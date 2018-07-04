{ Unité : Source TOF de la FICHE : CPVIDESECTIONS
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.50.001.004    10/06/05    JP     Création de l'unité : migration eAgl de ModTlEcr.Pas
--------------------------------------------------------------------------------------}
unit CPMODIFTABLELIBRE_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db,  dbtables, FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  Forms, SysUtils, ComCtrls, HCtrls, uTob, UTOF, Ent1, Graphics, Grids,
  {Pour la gestion des filtres}
  ParamSoc,		// GetParamSocSecur YMO
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UObjFiltres, Hent1, extctrls {YM04/10/2005 TTimer};

type

  TOF_CPMODIFTABLELIBRE = class (TOF)
    procedure OnLoad               ; override;
    procedure OnArgument(S :string); override;
    procedure OnClose              ; override;
    {Lancement du traitement de mise à jour des tables libres}
    procedure OnUpdate             ; override;
  private
    FormKeyDown : TKeyEvent;
    FAutoSearch     : TAutoSearch;
    FCritModified   : Boolean;        // Indicateur de modifiaction dans le PAGECONTROL
    FSearchTimer    : TTimer;
    FLoading        : Boolean;        // Indicateur pour en cours de chargement
    {Constitution de la requête de mise à jour des tables libres}
    function  RequeteUpdate(T : TOB) : string;
    {Mise à jour des tables libres dans la base}
    procedure InsererDansTable;
    {Branche les évènements sur les composants}
    procedure InitEvenement;
    {Affecte les pointeurs sur les composants}
    procedure InitComposant;
    {Gestion de la pérsentation de la grille}
    procedure InitTitreGrid(TypNat : Char);
    {Affichage des titres de la grille}
    procedure AfficheTitres(TypNat : Char);

    // YMO 04/10/2005 Ajout Autosearch + bouton Recherche standard (FQ16657)
    procedure InitAutoSearch;
    procedure CritereChange    (Sender: TObject);
    procedure SetCritModified  (Value: Boolean);
    procedure SearchTimerTimer (Sender: TObject);
    procedure ResetTimer       (Sender: TObject);
    property CritModified         : Boolean      read FCritModified        write SetCritModified;
    property AutoSearch           : TAutoSearch  read FAutoSearch          write FAutoSearch;

  protected
    FListe      : THGrid;
    Axe         : THValComboBox;
    Table       : THValComboBox;
    Exercice    : THValComboBox;
    Journal     : THValComboBox;
    QualifPiece : THValComboBox;
    NaturePiece : THValComboBox;
    Numeropiece : THNumEdit;
    Numeropiece_: THNumEdit;
    General     : THEdit;
    Auxiliaire  : THEdit;

    {Mise à jour des date en fonction de l'exercice}
    procedure EXERCICEChange (Sender : TObject);
    {Mise à jour du DataType des zones TableLibre en fonction de la nature du traitement}
    procedure TABLEChange    (Sender : TObject);
    {Mise à jour du DataType de la zone auxiliaire en fonction de l'axe}
    procedure AxeChange      (Sender : TObject);
    {Rafraîchissement de la grille}
    procedure BChercheClick  (Sender : TObject);
    {Gestion du boutons de désélection totale}
    procedure BDeTagClick    (Sender : TObject);
    {Gestion du boutons de sélection totale}
    procedure BTagClick      (Sender : TObject);
    {Sur la sélection / désélection d'une ligne}
    procedure OnFlipSelection(Sender : TObject);
    {Pour simuler le F9}
    procedure OnFormKeyDown  (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure AuxiElipsisClick(Sender : TObject);
  public
    ObjFiltre  : TObjFiltre;
    QuelTab    : string;
    TobTrait   : TOB;
    TotalSelec : Integer;
    Champ      : string;
    RunMaj     : Boolean;

    {On s'assure que la nouvelle valeur de table libre existe}
    function  ControleCpteOk : Boolean;
    {Fait un BChercheClick après le chargement éventuelle d'un filtre}
    procedure ApresChargementFiltre;
    {Initialisation des composants de la fiches en fonction de l'appel de la fiche}
    procedure InitLeCritere;
    {Affichage du nombre d'éléments sélectionnés}
    procedure CompteElemSelectionner;
    {Fabrication de la requête de sélection pour affichage dans la grille}
    procedure FabriqueLaRequete;
    {Affichage des champs voulus dans la grille}
    procedure RempliLaListe;
    {Gestion de la sélection / désélection des lignes de la grille}
    procedure TagDeTag(Select : Boolean);
  end;

procedure ModifSerieTableLibreEcr(QuelTab : string; FTypePiece : SetttTypePiece);

implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  HTB97, CpteUtil, HMsgBox, HStatus, LookUp, Math, Windows, UTofMulParamGen; {13/04/07 YMO F5 sur Auxiliaire }

var
  TypePiece : SetttTypePiece;

const
  GR_PROPORTION = 1.3;

{---------------------------------------------------------------------------------------}
procedure ModifSerieTableLibreEcr(QuelTab : string; FTypePiece : SetttTypePiece);
{---------------------------------------------------------------------------------------}
begin
  TypePiece := FTypePiece;
  AglLanceFiche('CP', 'CPMODIFTABLELIBRE', '', '', QuelTab + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
var
  Composant : TControlFiltre ;
begin
  inherited;
  QuelTab := ReadTokenSt(S);
  {Gestion des filtres, centralisée dans UObjFiltres}
  FSearchTimer := TTimer.Create(Ecran);
  Composant.Filtre   := TToolbarButton97(GetControl('BFILTRE'  ));
  Composant.Filtres  := THValComboBox   (GetControl('FFILTRES' ));
  Composant.PageCtrl := TPageControl    (GetControl('PCCONTROL'));
  ObjFiltre := TObjFiltre.Create(Composant, 'MODTLECR' + QuelTab[1]);
  ObjFiltre.ApresChangementFiltre := ApresChargementFiltre;

  {Affecte les pointeurs sur les composants de la fiche}
  InitComposant;
  {Affecte les pointeurs sur évènements des composants de la fiche}
  InitEvenement;
  {Préparation de la fiche en fonction de l'appel}
  InitLeCritere;
  {Tob de traitement et de chargement de la grille}
  TobTrait := TOB.Create('$$$', nil, -1);

  FSearchTimer.Enabled := False;
  FSearchTimer.OnTimer := SearchTimerTimer;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('EDAUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
  if Assigned(TobTrait)  then FreeAndNil(TobTrait);
  FreeAndNil(FSearchTimer);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TPageControl(GetControl('PCCONTROL')).ActivePageIndex := 0;
  RunMaj := False;

  Journal    .ItemIndex := 0;
  Table      .ItemIndex := 0;
  NaturePiece.ItemIndex := 0;

  {Gestion des dates}
  if VH^.CPExoRef.Code <> '' then begin
    Exercice.Value := VH^.CPExoRef.Code ;
    SetControlText('DATECOMPTABLE'  , DateToStr(VH^.CPExoRef.Deb));
    SetControlText('DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
  end
  else begin
    Exercice.Value := VH^.Entree.Code ;
    SetControlText('DATECOMPTABLE'  , DateToStr(V_PGI.DateEntree));
    SetControlText('DATECOMPTABLE_', DateToStr(V_PGI.DateEntree));
  end;

  {Chargement de l'éventuel filtre}
  ObjFiltre.Charger;

  {ATTENTION : A faire après le chargement des filtres : Gestion du qualifpiece et de l'aide}
  if TypePiece <> [] then begin
    QualifPiece.Enabled := False;
    if tpReel in TypePiece then begin
      Ecran.HelpContext := 7259500;
      QualifPiece.Value := 'N';
      SetControlText('CBQUALIFPIECE', 'N');
      QualifPiece.ItemIndex := QualifPiece.Values.IndexOf('N');
    end
    else if tpSim  in TypePiece then begin
      Ecran.HelpContext := 7280500;
      QualifPiece.Value := 'S';
    end
    else if tpSitu in TypePiece then begin
      Ecran.HelpContext := 7301500;
      QualifPiece.Value := 'U';
    end
    else if tpPrev in TypePiece then begin
      Ecran.HelpContext := 7316500;
      QualifPiece.Value := 'P';
    end
    else
      QualifPiece.Enabled := True;
   end
   else
     QualifPiece.Enabled := True;
   // YMO 04/10/2005 Ajout Autosearch + bouton Recherche standard (FQ16657)
   InitAutoSearch ;
end;

{Lancement du traitement sur le click du bouton Valider
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  i : Integer;
begin
  inherited;
  if not RunMaj then Exit ;

  if FListe.Cells[0, 1] = '' then begin
    HShowMessage('6;' + Ecran.Caption + ';Aucun compte ne correspond à votre critère.'#13#13 +
                                        'Aucun traitement à effectuer;W;O;O;O;', '', '');
    Exit;
  end;

  if TotalSelec = 0 then begin
    HShowMessage('7;' + Ecran.Caption + ';Aucun compte n''est sélectionné.'#13#13 +
                                        'Aucun traitement à effectuer;W;O;O;O;', '', '');
    Exit;
  end;

  if not ControleCpteOk then Exit;

  if GetControlText('EDNOUVAL') = GetControlText('EDANCVAL') then begin
    HShowMessage('9;' + Ecran.Caption + ';La nouvelle valeur est égale à l''ancienne valeur.'#13#13 +
                                        'Aucun traitement à effectuer.;W;O;O;O;', '', '');
    Exit;
  end;

  if HShowMessage('8;' + Ecran.Caption + ';Confirmez-vous le traitement de mise à jour des tables libres sélectionnées ?;Q;YN;N;N;',
                  '', '') <> mrYes then Exit;


  InitMove(FListe.RowCount - 1, '');
  for i := 1 to FListe.RowCount - 1 do begin
    MoveCur(False) ;
    if FListe.IsSelected(i) then begin
      TobTrait.Detail[i - 1].SetString(Trim(Champ), GetControlText('EDNOUVAL'));
      {On indique la mise à jour des enregistrements modifiés pour insérer dans table :
       Le traitement automatique ne fonctionne que pour les Tobs réelles : NumChamp < 1000}
      TobTrait.Detail[i - 1].SetBoolean('MODIFIE', True);
    end;
  end;

  {Insertion des modifications dans la table idoine}
  InsererDansTable;

  FiniMove;
  BChercheClick(GetControl('BCHERCHE')) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.ApresChargementFiltre;
{---------------------------------------------------------------------------------------}
begin
  {On lance la recherche après chargement du filtre}
  if GetControlText('FFILTRES' ) <> '' then BChercheClick(GetControl('BCHERCHE'));
end;

{Affecte les pointeurs sur les composants de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.InitComposant;
{---------------------------------------------------------------------------------------}
begin
  FListe      := THGrid(GetControl('FLISTE'));
  Axe         := THValComboBox(GetControl('CBAXE'));
  Table       := THValComboBox(GetControl('CBTABLELIBRE'));
  Exercice    := THValComboBox(GetControl('CBEXERCICE'));
  Journal     := THValComboBox(GetControl('CBJOURNAL'));
  QualifPiece := THValComboBox(GetControl('CBQUALIFPIECE'));
  NaturePiece := THValComboBox(GetControl('CBNATUREPIECE'));
  Numeropiece := THNumEdit(GetControl('EDNUMPIECE' ));
  Numeropiece_:= THNumEdit(GetControl('EDNUMPIECE_'));
  General     := THEdit(GetControl('EDGENERAL'));
  Auxiliaire  := THEdit(GetControl('EDAUXILIAIRE'));
end;

{affecte les pointeurs sur les évènements des composants de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.InitEvenement;
{---------------------------------------------------------------------------------------}
begin
  FormKeyDown := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnFormKeyDown;

  Axe     .OnChange      := AxeChange;
  Table   .OnChange      := TABLEChange;
  Exercice.OnChange      := EXERCICEChange;
  FListe.OnFlipSelection := OnFlipSelection;

  TToolbarButton97(GetControl('BCHERCHE')).OnClick := BChercheClick;
  TToolbarButton97(GetControl('BCHERCHE_')).OnClick := BChercheClick;
  TToolbarButton97(GetControl('BTAG'    )).OnClick := BTagClick;
  TToolbarButton97(GetControl('BDETAG'  )).OnClick := BDeTagClick;
end;

{gestion de l'affichage de la grille : titre, "centrage", taille des colonnes
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.InitTitreGrid(TypNat : Char);
{---------------------------------------------------------------------------------------}
begin
  case TypNat of
    'A', 'E' : begin
            FListe.ColWidths[0] := Round(40 * GR_PROPORTION);
            FListe.ColWidths[1] := Round(75 * GR_PROPORTION);
            FListe.ColWidths[2] := Round(50 * GR_PROPORTION);
            FListe.ColWidths[3] := Round(92 * GR_PROPORTION);
            FListe.ColWidths[4] := Round(92 * GR_PROPORTION);
            FListe.ColWidths[5] := Round(92 * GR_PROPORTION);
            FListe.ColWidths[6] := Round(92 * GR_PROPORTION);
          end;

    'U' : begin
            FListe.ColWidths[0] := Round(50 * GR_PROPORTION);
            FListe.ColWidths[1] := Round(76 * GR_PROPORTION);
            FListe.ColWidths[2] := Round(61 * GR_PROPORTION);
            FListe.ColWidths[3] := Round(61 * GR_PROPORTION);
            FListe.ColWidths[4] := Round(95 * GR_PROPORTION);
            FListe.ColWidths[5] := Round(95 * GR_PROPORTION);
            FListe.ColWidths[6] := Round(95 * GR_PROPORTION);
          end;
  end;
  {Affichage des titres de la grilles}
  AfficheTitres(TypNat);
end;

{Affichage des titres de la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.AfficheTitres(TypNat : Char);
{---------------------------------------------------------------------------------------}
begin
  FListe.Titres.Clear;
  Fliste.Titres.Add('Journal;C;');
  Fliste.Titres.Add('Date;C;');
  Fliste.Titres.Add('Numéro;C;');
  case TypNat of
    'A', 'E' : begin
                 Fliste.Titres.Add('Général;G;');
                 if TypNat = 'A' then Fliste.Titres.Add('Auxiliaire;G;')
                                 else Fliste.Titres.Add('Section;G;');
                 Fliste.Titres.Add('Débit;D;');
                 Fliste.Titres.Add('Crédit;D;');
               end;
    'U' :      begin
                 Fliste.Titres.Add('Nature;C;');
                 Fliste.Titres.Add('Général;G;');
                 Fliste.Titres.Add('Section;G;');
                 Fliste.Titres.Add('Etablissement;G;');
               end;
  end;
  FListe.UpdateTitres;
end;

{Initialisation des composants de la fiches en fonction de l'appel de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.InitLeCritere;
{---------------------------------------------------------------------------------------}
begin
  if QuelTab='' then Exit;
  {Chargement de la combobox TableLibre}
  ChargeComboTableLibre(QuelTab[1], TABLE.Values, TABLE.Items) ;

  case QuelTab[1] of
    'A' : begin
            SetControlProperty('CBJOURNAL'  , 'DATATYPE', 'TTJOURNAL');
            SetControlCaption('LBAUXILIAIRE', '&Section');
            SetControlVisible('EDAUXILIAIRE', True);
            SetControlCaption('LBNATURE'    , '&Axe');
            NaturePiece.Visible := False;
            Axe        .Visible := True;
            Axe.ItemIndex := 0;

            Ecran.Caption := 'Modification des tables libres des écritures analytiques';
            Ecran.HelpContext := 7370500;
          end;
    'E' : begin
            SetControlProperty('CBJOURNAL'  , 'DATATYPE', 'TTJOURNAL');
            SetControlCaption('LBAUXILIAIRE', '&Auxiliaire');
            SetControlVisible('EDAUXILIAIRE', True);
            SetControlCaption('LBNATURE'    , '&Nature');
            NaturePiece.Visible := True;
            Axe        .Visible := False;
            Ecran.Caption := 'Modification des tables libres des écritures comptables';
          end;
    'U' : begin
            SetControlProperty('CBJOURNAL'  , 'DATATYPE', 'TTBUDJAL');
            SetControlCaption('LBAUXILIAIRE', '&Axe');
            SetControlVisible('EDAUXILIAIRE', False);
            SetControlProperty('EDGENERAL'  , 'DATATYPE', 'TZBUDGEN');
            NaturePiece.DataType := 'TTNATECRBUD';
            NaturePiece.Visible := True;
            Axe        .Visible := True;
            Axe.Top := 10;
            Axe.ItemIndex := 0;
            Ecran.Caption := 'Modification des tables libres des écritures budgétaires';
          end;
  end;

  {gestion de la présentation de la grille}
  InitTitreGrid(QuelTab[1]);
  {Mise à jour du titre de l'écran}
  UpdateCaption(Ecran);
end;

{Mise à jour du DataType de la zone auxiliaire en fonction de l'axe
 AxeToDataType (dans Ent1) est la conversion de AxeToTz (ZoomTable) pour la propriété DataType
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.AxeChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (Axe.Value <> '') and (Length(Axe.Value) > 1) then begin
    if (QuelTab = 'A') then
      SetControlProperty('EDAUXILIAIRE', 'DATATYPE', AxeToDataType(Axe.Value));
  end;
end;

{Mise à jour du DataType des zones TableLibre en fonction de la nature du traitement
 NatureToDataType (dans Ent1) est la conversion de NatureToTz (ZoomTable) pour la propriété DataType
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.TABLEChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if Table.Value <> '' then begin
    SetControlProperty('EDANCVAL', 'DATATYPE', NatureToDataType(TABLE.Value));
    SetControlProperty('EDNOUVAL', 'DATATYPE', NatureToDataType(TABLE.Value));
  end;
end;

{Rafraîchissement de la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.BChercheClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  FLoading := true;
  FListe.VidePile(False);
  RunMaj := True;
  FabriqueLaRequete;
  RempliLaListe;
  CritModified := False;
  FLoading := false;
end;

{Gestion du boutons de désélection totale
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.BDeTagClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TagDeTag(False);
end;

{Gestion du boutons de sélection totale
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.BTagClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TagDeTag(True);
end;

{Gestion de la sélection / désélection des lignes de la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.TagDeTag(Select : Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 1 to FListe.RowCount - 1 do begin
    {Si on demande la sélection et que la ligne n'est pas sélectionnée ou au contraire
        on demande la désélection et la ligne est sélectionnée}
    if (Select and not FListe.IsSelected(n)) or (not Select and FListe.IsSelected(n)) then
      FListe.FlipSelection(n);
  end;

  FListe.Invalidate;
  {On désactive le bouton si rien n'est sélectionné}
  SetControlEnabled('BDETAG', Select);
  {On désactive le bouton si tout est sélectionné}
  SetControlEnabled('BTAG', not Select);
  {Affichage du nombres de lignes sélectionnées}
  CompteElemSelectionner;
end;

{Sur la sélection d'une ligne, activation / désactivation des boutons de sélection et
 calcul du nombre de lignes sélectionnées
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.OnFlipSelection(Sender: TObject);
{---------------------------------------------------------------------------------------}

begin
  CompteElemSelectionner;
  SetControlEnabled('BDETAG', FListe.nbSelected <> 0);
  SetControlEnabled('BTAG'  , (FListe.nbSelected <> FListe.RowCount - 1));
end;

{Mise à jour des date en fonction de l'exercice
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.EXERCICEChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ExoToDates(EXERCICE.Value, GetControl('DATECOMPTABLE'), GetControl('DATECOMPTABLE_'));
end;

{Fabrication de la requête de sélection pour affichage dans la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.FabriqueLaRequete;
{---------------------------------------------------------------------------------------}
var
  Sql : string;
  Ind : string;
begin
  if QuelTab = '' then Exit;
  if Length(Table.Value) < 3 then Exit;

  case QuelTab[1] of
    'A' : begin
            Champ := 'Y_TABLE' + Table.Value[3] + ' ';
            Ind   := ', Y_EXERCICE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ';
            Sql := 'SELECT Y_JOURNAL,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_GENERAL,Y_SECTION,Y_DEBIT,Y_CREDIT,' + Champ +
                   Ind + 'FROM ANALYTIQ WHERE '+
                   'Y_DATECOMPTABLE BETWEEN "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE'))) +
                   '" AND "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE_'))) + '" AND ' +
                   'Y_TABLE' + Table.Value[3] + ' = "' + GetControlText('EDANCVAL') + '" AND ' +
                   'Y_AXE = "' + Axe.Value + '" ';

            if Journal     .Value <> '' then Sql := Sql + 'AND Y_JOURNAL = "'    + Journal     .Value + '" ';
            if Exercice    .Value <> '' then Sql := Sql + 'AND Y_EXERCICE = "'   + Exercice    .Value + '" ';
            if General     .Text  <> '' then Sql := Sql + 'AND Y_GENERAL LIKE "' + General     .Text  + '%" ';
            if Auxiliaire  .Text  <> '' then Sql := Sql + 'AND Y_SECTION LIKE "' + Auxiliaire  .Text  + '%" ';
            if Numeropiece .Value  > 0  then Sql := Sql + 'AND Y_NUMEROPIECE >=' + Numeropiece .Text  + ' ';
            if Numeropiece .Value  > 0  then Sql := Sql + 'AND Y_NUMEROPIECE <=' + Numeropiece_.Text  + ' ';
            if QualifPiece .Value <> '' then Sql := Sql + 'AND Y_QUALIFPIECE = "'+ QualifPiece .Value + '" ';
          end;

    'E' : begin
            Champ := 'E_TABLE' + Table.Value[3] + ' ';
            Ind   := ', E_EXERCICE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE ';
            Sql := 'SELECT E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_DEBIT,E_CREDIT,' + Champ +
                   Ind + 'FROM ECRITURE WHERE '+
                   'E_DATECOMPTABLE BETWEEN "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE'))) +
                   '" AND "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE_'))) + '" AND ' +
                   'E_TABLE' + Table.Value[3] + ' = "' + GetControlText('EDANCVAL') + '" AND '+
                   'E_ECRANOUVEAU = "N" ';

            if NaturePiece .Value <> '' then Sql := Sql + 'AND E_NATUREPIECE = "'  + NaturePiece.Value + '" ';
            if Journal     .Value <> '' then Sql := Sql + 'AND E_JOURNAL = "'      + Journal.Value     + '" ';
            if Exercice    .Value <> '' then Sql := Sql + 'AND E_EXERCICE = "'     + Exercice.Value    + '" ';
            if General     .Text  <> '' then Sql := Sql + 'AND E_GENERAL LIKE "'   + General.Text      + '%" ';
            if Auxiliaire  .Text  <> '' then Sql := Sql + 'AND E_AUXILIAIRE LIKE "'+ Auxiliaire.Text   + '%" ';
            if Numeropiece .Value  > 0  then Sql := Sql + 'AND E_NUMEROPIECE >= '  + Numeropiece.Text  + ' ';
            if Numeropiece_.Value  > 0  then Sql := Sql + 'AND E_NUMEROPIECE <= '  + Numeropiece_.Text + ' ';

            if TypePiece <> [] then Sql := Sql + WhereSupp('E_', TypePiece);
          end;
    'U' : begin
            Champ := 'BE_TABLE' + Table.Value[3] + ' ';
            Ind   := ',BE_EXERCICE, BE_AXE, BE_QUALIFPIECE ';
            Sql := 'SELECT BE_BUDJAL,BE_DATECOMPTABLE,BE_NUMEROPIECE,BE_NATUREBUD,BE_BUDGENE,BE_BUDSECT,BE_ETABLISSEMENT,' + Champ +
                   Ind + 'FROM BUDECR WHERE '+
                   'BE_DATECOMPTABLE BETWEEN "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE'))) +
                   '" AND "' + USDateTime(StrToDateTime(GetControlText('DATECOMPTABLE_'))) + '" AND ' +
                   'BE_TABLE' + Table.Value[3] + ' = "' + GetControlText('EDANCVAL') + '" ';

            if NaturePiece .Value <> '' then Sql := Sql + 'AND BE_NATUREBUD = "'   + NaturePiece .Value + '" ';
            if Journal     .Value <> '' then Sql := Sql + 'AND BE_BUDJAL = "'      + Journal     .Value + '" ';
            if Exercice    .Value <> '' then Sql := Sql + 'AND BE_EXERCICE = "'    + Exercice    .Value + '" ';
            if General     .Text  <> '' then Sql := Sql + 'AND BE_BUDGENE LIKE "'  + General     .Text  + '%" ';
            if Numeropiece .Value  > 0  then Sql := Sql + 'AND BE_NUMEROPIECE >= ' + Numeropiece .Text  + ' ';
            if Numeropiece_.Value  > 0  then Sql := Sql + 'AND BE_NUMEROPIECE <= ' + Numeropiece_.Text  + ' ';
            if QualifPiece .Value <> '' then Sql := Sql + 'AND BE_QUALIFPIECE = "' + QualifPiece .Value + '" ';
          end;
  end;

  {Chargement dans la Tob}
  TobTrait.LoadDetailFromSQL(SQL);
  {Le setModified et le IsModified des champs des Tobs ne marchents que pour les champs réels (< 1000),
   on est obligé de le gérer à la main}
  if TobTrait.Detail.Count > 0 then
    TobTrait.Detail[0].AddChampSupValeur('MODIFIE', False, True);
end;

{Affichage des champs voulus dans la grille
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.RempliLaListe;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  for n := 0 to TobTrait.Detail.Count - 1 do begin
  // YMO 04/09/2005 Le compte était faux car cette ligne était placée en dehors
  // de la boucle, d'où plantage si pas d'enreg dans la sélection
    Fliste.RowCount := Max(n + 1, 2);
    FListe.Cells[0, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1000));
    FListe.Cells[1, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1001));
    FListe.Cells[2, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1002));
    FListe.Cells[3, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1003));
    FListe.Cells[4, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1004));
    FListe.Cells[5, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1005));
    FListe.Cells[6, n + 1] := TobTrait.Detail[n].GetString(TobTrait.Detail[n].GetNomChamp(1006));
  end;

  {On ne pourra lancer le traitement que si la tob n'est pas vide}
  RunMaj := TobTrait.Detail.Count > 0;

  {Affichage de titres de la grille}
  AfficheTitres(QuelTab[1]);
  FListe.Invalidate;
  {Par défaut on sélectionne toutes les lignes}
  if TobTrait.Detail.Count > 0 then
    TagDeTag(True);
end;

{Affichage du nombre d'éléments sélectionnés
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.CompteElemSelectionner;
{---------------------------------------------------------------------------------------}
begin
  TotalSelec := FListe.nbSelected;
  SetControlCaption('NB1', IntToStr(TotalSelec));
  if TotalSelec > 1 then SetControlCaption('Tex1', 'lignes sélectionnées')
                    else SetControlCaption('Tex1', 'ligne sélectionnée');
end;

{Mise à jour des tables libres dans la base
{---------------------------------------------------------------------------------------}
procedure TOF_CPMODIFTABLELIBRE.InsererDansTable;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  SQL : string;
begin
  for n := 0 to TobTrait.Detail.Count - 1 do begin
    {Si la table libre a été modifié ...}
    if TobTrait.Detail[n].GetBoolean('MODIFIE') then begin
      {... On prépare la requête d'UPDTE}
      SQL := RequeteUpdate(TobTrait.Detail[n]);
      {... et éventuelle mise à jour dans la table}
      if Trim(SQL) <> '' then
        ExecuteSQL(SQL);
    end;
  end;
end;

{Constitution de la requête de mise à jour des tables libres
{---------------------------------------------------------------------------------------}
function TOF_CPMODIFTABLELIBRE.RequeteUpdate(T : TOB) : string;
{---------------------------------------------------------------------------------------}
begin
  if QuelTab = '' then Exit;
  if Length(Table.Value) < 3 then Exit;

  case QuelTab[1] of
    'A' : Result := 'UPDATE ANALYTIQ SET '+ Champ + ' = "' + T.GetString(Champ)          + '" WHERE ' +
                    'Y_JOURNAL = "'       + T.GetString('Y_JOURNAL')                     + '" AND ' +
                    'Y_EXERCICE = "'      + T.GetString('Y_EXERCICE')                    + '" AND ' +
                    'Y_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('Y_DATECOMPTABLE')) + '" AND ' +
                    'Y_NUMEROPIECE = '    + T.GetString('Y_NUMEROPIECE')                 + ' AND ' +
                    'Y_NUMLIGNE = '       + T.GetString('Y_NUMLIGNE')                    + ' AND ' +
                    'Y_AXE = "'           + T.GetString('Y_AXE')                         + '" AND ' +
                    'Y_NUMVENTIL = '      + T.GetString('Y_NUMVENTIL')                   + ' AND ' +
                    'Y_QUALIFPIECE = "'   + T.GetString('Y_QUALIFPIECE')                 + '"';

    'E' : Result := 'UPDATE ECRITURE SET '+ Champ + ' = "' + T.GetString(Champ)          + '" WHERE ' +
                    'E_JOURNAL = "'       + T.GetString('E_JOURNAL')                     + '" AND ' +
                    'E_EXERCICE = "'      + T.GetString('E_EXERCICE')                    + '" AND ' +
                    'E_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('E_DATECOMPTABLE')) + '" AND ' +
                    'E_NUMEROPIECE = '    + T.GetString('E_NUMEROPIECE')                 + ' AND ' +
                    'E_NUMLIGNE = '       + T.GetString('E_NUMLIGNE')                    + ' AND ' +
                    'E_NUMECHE = '        + T.GetString('E_NUMECHE')                     + ' AND ' +
                    'E_QUALIFPIECE = "'   + T.GetString('E_QUALIFPIECE')                 + '"';

    'U' : Result := 'UPDATE BUDECR SET '  + Champ + ' = "' + T.GetString(Champ)          + '" WHERE ' +
                    'BE_BUDGENE = "'      + T.GetString('BE_BUDGENE')                    + '" AND ' +
                    'BE_BUDJAL = "'       + T.GetString('BE_BUDJAL')                     + '" AND ' +
                    'BE_EXERCICE = "'     + T.GetString('BE_EXERCICE')                   + '" AND ' +
                    'BE_DATECOMPTABLE = "'+ UsDateTime(T.GetDateTime('BE_DATECOMPTABLE'))+ '" AND ' +
                    'BE_BUDSECT = "'      + T.GetString('BE_BUDSECT')                    + '" AND ' +
                    'BE_AXE = "'          + T.GetString('BE_AXE')                        + '" AND ' +
                    'BE_NUMEROPIECE = '   + T.GetString('BE_NUMEROPIECE')                + ' AND ' +
                    'BE_QUALIFPIECE = "'  + T.GetString('BE_QUALIFPIECE')                + '"';
    else
      Result := '';
  end;
end;

{On s'assure que la nouvelle valeur de table libre existe
{---------------------------------------------------------------------------------------}
function TOF_CPMODIFTABLELIBRE.ControleCpteOk : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := GetControlText('EDNOUVAL') = '';
  if not Result then begin
    if not LookUpValueExist(GetControl('EDNOUVAL')) then begin
       TPageControl(GetControl('PCCONTROL')).ActivePageIndex := 0;
       SetFocusControl('EDNOUVAL');
       HShowMessage('5;' + Ecran.Caption + ';Le code renseigné n''existe pas;W;O;O;O;','','');
    end
    else
      Result := True;
  end;
end;

{---------------------------------------------------------------------------------------}

procedure TOF_CPMODIFTABLELIBRE.OnFormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F9 : BChercheClick(GetControl('BCHERCHE'));
    else
      FormKeyDown(Sender, Key, Shift);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout Autosearch + bouton Recherche standard (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.InitAutoSearch;
var
  i: integer;
begin

  if (V_PGI.AutoSearch) and (AutoSearch <> AsMouetteForce) then
    AutoSearch := asTimer;

  for i := 0 to Ecran.ComponentCount - 1 do
  begin
    if (Ecran.Components[i] is TControl) and
      (TControl(Ecran.Components[i]).Parent is TTabSheet) then
      if (Ecran.Components[i] is TControl) and
        TControl(Ecran.Components[i]).Visible and
        (TControl(Ecran.Components[i]).Enabled) then
      begin
        if (Ecran.Components[i] is THCritMaskEdit) and (not
          Assigned(THCritMaskEdit(Ecran.Components[i]).OnEnter)) then
          THCritMaskEdit(Ecran.Components[i]).OnEnter := V_PGI.EgaliseOnEnter;

        case AutoSearch of
          asChange:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).OnChange) then
                TEdit(Ecran.Components[i]).OnChange := SearchTimerTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).OnClick) then
                THValComboBox(Ecran.Components[i]).OnClick := SearchTimerTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).OnClick) then
                TCheckBox(Ecran.Components[i]).OnClick := SearchTimerTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
                THCritMaskEdit(Ecran.Components[i]).OnChange :=
                  SearchTimerTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
                THRadioGroup(Ecran.Components[i]).OnClick := SearchTimerTimer;
            end;
          asExit:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).onExit) then
                TEdit(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).onExit) then
                THValComboBox(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).onExit) then
                TCheckBox(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).onExit) then
                THCritMaskEdit(Ecran.Components[i]).onExit := SearchTimerTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).onExit) then
                THRadioGroup(Ecran.Components[i]).onExit := SearchTimerTimer;
            end;
          asTimer:
            begin
              if (Ecran.Components[i] is TEdit) and not
                assigned(TEdit(Ecran.Components[i]).OnChange) then
                TEdit(Ecran.Components[i]).OnChange := ResetTimer;
              if (Ecran.Components[i] is THValComboBox) and not
                assigned(THValComboBox(Ecran.Components[i]).OnClick) then
                THValComboBox(Ecran.Components[i]).OnClick := ResetTimer;
              if (Ecran.Components[i] is TCheckBox) and not
                assigned(TCheckBox(Ecran.Components[i]).OnClick) then
                TCheckBox(Ecran.Components[i]).OnClick := ResetTimer;
              if (Ecran.Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
                THCritMaskEdit(Ecran.Components[i]).OnChange := ResetTimer;
              if (Ecran.Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
                THRadioGroup(Ecran.Components[i]).OnClick := ResetTimer;
            end;
        else
          begin
            if (Ecran.Components[i] is TEdit) and not
              assigned(TEdit(Ecran.Components[i]).OnChange) then
              TEdit(Ecran.Components[i]).OnChange := CritereChange;
            if (Ecran.Components[i] is THValComboBox) and not
              assigned(THValComboBox(Ecran.Components[i]).OnClick) then
              THValComboBox(Ecran.Components[i]).OnClick := CritereChange;
            if (Ecran.Components[i] is TCheckBox) and not
              assigned(TCheckBox(Ecran.Components[i]).OnClick) then
              TCheckBox(Ecran.Components[i]).OnClick := CritereChange;
            if (Ecran.Components[i] is THCritMaskEdit) and not
              assigned(THCritMaskEdit(Ecran.Components[i]).OnChange) then
              THCritMaskEdit(Ecran.Components[i]).OnChange := CritereChange;
            if (Ecran.Components[i] is THRadioGroup) and not
              assigned(THRadioGroup(Ecran.Components[i]).OnClick) then
              THRadioGroup(Ecran.Components[i]).OnClick := CritereChange;
          end;

        end;
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Copie à partir de uTofViergeMul, pour importer le
Suite ........ : comportement du bouton de recherche (avec l'entonnoir) (FQ16657)
Mots clefs ... : YMO
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.CritereChange(Sender: TObject);
begin
  CritModified := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Copie à partir de uTofViergeMul, pour importer le
Suite ........ : comportement du bouton de recherche (avec l'entonnoir) (FQ16657)
Mots clefs ... : YMO
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.SetCritModified(Value: Boolean);
begin
  FCritModified     := Value;
  GetControl('BCHERCHE').Visible  := not Value;
  GetControl('BCHERCHE_').Visible := Value;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout fonction Autosearch  (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.SearchTimerTimer(Sender: TObject);
begin
  FSearchTimer.Enabled := False;
  CritModified := True;
  BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout fonction Autosearch  (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TOF_CPMODIFTABLELIBRE.ResetTimer(Sender: TObject);
begin
  if FLoading then Exit;

  CritModified := True;
  FSearchTimer.Enabled := False;
  FSearchTimer.Enabled := True;
end;

initialization
  RegisterClasses([TOF_CPMODIFTABLELIBRE]);

end.


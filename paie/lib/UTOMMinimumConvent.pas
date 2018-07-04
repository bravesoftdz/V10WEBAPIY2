{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 21/11/2001
Modifié le ... :   /  /
Description .. : TOM gestion des tables
Suite ........ : Attention : liens 1/N
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
unit UTOMMinimumConvent;
{ Cette Tom gère la table MinimumConvent, avec deux comportements : un de type
 tablette (type =val), un de type Intervale (type = INT). Bien que très différents
 sémantiquement parlant, on a tout laissé ensemble dans la même table, donc
 dans la même TOM : "On ne sait jamais".
 Pour ne pas avoir un code trop pollué, je positionnerai systématiquement un
 NomFiche : <SaisieValeur>, <SaisieIntervalle> ;
 Dans la saisie par Intervalle, j'ai rajouté un label invisible, destiné à
 positionner "NomFiche"

 Comme cela, je teste sur NomFiche.
 pour distinguer les fiches, les noms des zones sont préfixés par une lettre :
 Vxxxxxxxx : Il s'agit de la fiche liée à la table MINIMUMCONVENT  de saisie
             par valeur : façon tablette
 Ixxxxxxxx : Il s'agit de la fiche liée à la table MINIMUMCONVENT  de saisie
            par intervalle : façon variable
----------------------------------------------------------------------------------
 03/03/2000 Modif Ph : on ne gere plus la zone PREF_CODE mais uniquement IPMI_CODE
 donc pour simplifier la zone PREF_CODE existe tjrs mais est invisible
 Voir modif référencée par // Modif 03/03

 PT1-1 22/10/2001 SB V563 Fiche de bug n°321
                          bug ds la gestion des min. Conv. de la grille
 PT1-2 22/10/2001 SB V563 Fiche de bug n°322
                          Suppression des Enr. Associés sur OnDeleteRecord
 PT2   05/12/2002 PH V591 FQ 10209 test existance code lors de la creation
 PT3   05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer
 PT4   05/02/2003 PH V595 Les tables rendent comme valeur un élément national (New)
 PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
--------------------------------------------------------------------------
PH Le 18/02/2003 Refonte complète du source pour respecter la norme de dev
--------------------------------------------------------------------------
 PT6   16/05/2003 PH V_421 Traitement de la date de modif dans ma table MINCONVPAIE
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
 PT7   09/07/2003  SB V_421 FQ 10369 refonte edition des tables dossier pour intégration de la grille
 PT8   08/09/2003 PH V_421 Recuperation automatique de la nature
 PT9   17/09/2003 PH V_421 FQ 10080 Traitement datatype zone PMI_PROFIL pour affectation bonne tablette
 PT10  11/05/2004 PH V_50  FQ 10961 Message controle des bornes
 PT11  08/06/2004 PH V_50  FQ 11035 Controle unicité du code sur nature et code
 PT12  30/11/2004 PH V_60  FQ 11780 Filtre sur la convention collective
 PT12-1 16/03/2005 PH V_60  FQ 11780 Filtre actif en chargement de la fiche
 PT13  20/12/2004 PH V_60  Suppression double Validation si Annulation + Affectation combo grille en création
 PT14  20/12/2004 PH V_60  FQ 11331 Compatibilité DB2 Refonte de l'écriture INSERT remplacé par une TOB
 PT15  06/01/2005 PH V_60  ON force l'écriture dans le cas de la modification de la grille sans autre modif
 PT11-1 06/01/2005 PH V_60  FQ 11035 Controle unicité du code sur nature et code et type de table de travail
 PT11-2 16/03/2005 PH V_60  FQ 11035 Suppression controle Abandon FQ
 PT12   01/02/2005 SB V_60 FQ 11844 Creation coefficient : Affectation du code sur libellé
 PT13   10/05/2005 PH V_60 FQ 12104 Controle saisie du code nature en creation de table
 PT14   10/06/2005 SB V_60 CWAS : Chargement tablette compatibilite CWAS
 PT15   10/06/2005 PH V_60 FQ 12260 : On pré-renseigne le code CCN
 PT16   02/08/2005 PH V_60 FQ 12467 Ergonomie
 PT17   31/08/2005 PH V_60 Et on rechange d'avis Plus de filtre sur les coeff/qualif si table Toutes conventions
 PT18   27/04/2006 NA V_65 FQ 12651 Remplacement du combo par un Elipsis pour Elément National
 PT19   29/09/2006 PH V_70 FQ 12512 Activation accès grille de saisie  Point 1
 PT20   07/05/2007 PH V_72 Concept Paie
 PT21   11/07/2007 FC V_72 FQ 14386 concepts
}

interface
uses Windows, StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fiche, EdtREtat,
{$ELSE}
  eFiche, utileAgl,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97,  LookUp,  // PT18
  ParamSoc, HPanel, ParamDat, Dialogs, PgOutils, PgOutils2,
  Grids;//PT21

type
  TNomFiche = (SaisieValeur, SaisieIntervalle);

type

  TOM_Minimumconvent = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnClose; override;
    procedure OnAfterUpdateRecord; override;

    procedure PDetailCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure PDetailCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure Entrergrille(Sender: TObject);     // PT18
    procedure PDetailElipsisClick(Sender: TObject); // PT18
    procedure BTnInsClick(Sender: TObject);
    procedure BTnDelClick(Sender: TObject);
    procedure BTnAddClick(Sender: TObject);
    procedure Ferme(Sender: TObject);
    //procedure iPmi_codeExit(Sender: TObject);
    procedure ChangeGrid;
    procedure SaisieValeur_OnChange(F: TField);
    procedure SaisieIntervalle_OnChange(F: TField);
    procedure EnableChampsOngletIntervalle;
    procedure ControleZonesIntervalle;
    procedure ControleZonesValeur;
    procedure AlimGrilleIntervalle;
    procedure InitGrilleIntervalle;
    procedure AffecteComboGrille;
    procedure ClickImprimer(sender: Tobject); //PT7
  private
    PDETAIL: THGrid;
    TDETAIL: TOB;
    NomFiche: TNomFiche;
    LectureSeule, CEG, STD, DOS, TopUpdate, OnFerme: Boolean;
    newsaisie : boolean; // PT18
    NatureInt, LaNature, LaConv : string; // PT8 // PT15
    Natureresult : string ;  // PT18
    DerniereCreate: string;
  end;


implementation
uses P5Def;
{ TOM_Minimumconvent }

//------------------------------------------------------------------------------
// Nouvel enregistrement
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.OnNewRecord;
var
  Z: ThLabel;
  LabNature: THLabel;
begin
  inherited;
  // DEB PT15
  if LaConv = '' then SetField('PMI_CONVENTION', '000')
  else SetField('PMI_CONVENTION', LaConv);
  // FIN PT15
  SetField('PMI_PREDEFINI', 'DOS');
  Z := THLabel(GetControl('INTERVALLE'));

  if Z <> nil then NomFiche := SaisieIntervalle
  else NomFiche := SaisieValeur;

  if NomFiche = SaisieValeur then
  begin
    SetField('PMI_TYPENATURE', 'VAL');
    //   LabNature := THLabel(GetControl('Nature_Val'));
    if LabNature <> nil then SetField('PMI_NATURE', GetControlText('Nature_Val')); // ????
  end
  else SetField('PMI_TYPENATURE', 'INT');
end;

//------------------------------------------------------------------------------
// Chargement de l'enregistrement
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.OnLoadRecord;
var
  Nature: string;
  P: TPageControl;
  T: TTabSheet;
  Z, LabNature: THLabel;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
  //  PT8   08/09/2003 PH V_421 Recuperation automatique de la nature
  if (LaNature <> '') and (DS.State in [dsInsert]) then
  begin // Recupération de la nature en cas de Creation
    SetField('PMI_NATURE', LaNature);
    SetControlEnabled('PMI_NATURE', FALSE);
  end;


  Z := THLabel(GetControl('INTERVALLE'));
  TopUpdate := false;
  if Z <> nil then NomFiche := SaisieIntervalle
  else NomFiche := SaisieValeur;
  // Fiche saisie de valeurs
  if NomFiche = SaisieValeur then
  begin // begin 1
    P := TPageControl(GetControl('PAGES'));
    if P <> nil then
    begin //begin 2
      T := TTabSheet(GetControl('PGENERAL'));
      if T <> nil then
      begin // begin 3
        Nature := GetField('PMI_NATURE');
        if Nature = 'COE' then T.Caption := 'COEFFICIENTS' else
          if Nature = 'QUA' then T.Caption := 'QUALIFICATIONS' else
          if Nature = 'IND' then T.Caption := 'INDICES' else
          if Nature = 'NIV' then T.Caption := 'NIVEAUX';
      end; // begin 3
    end; // begin 2
    LabNature := THLabel(GetControl('Nature_Val'));
    if LabNature <> nil then LabNature.caption := Nature;
  end; // begin 1
  // Fin Fiche saisie de valeurs

  // Début Fiche saisie d'intervalles
  if NomFiche = SaisieIntervalle then
  begin // begin 1
    PDetail := THGrid(GetControl('DETAILMINI'));
    if PDetail = nil then exit;
    PDetail.CellEnterBeforeCombo := True;
    if GetField('PMI_NATURE') = 'DIV' then {DEB PT1-1 pour alim de la grille}
      SetControlProperty('PMI_NATUREINT', 'Datatype', 'PGTYPEINTERDIV')
    else
      //  PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
      if GetField('PMI_NATURE') = 'DIW' then SetControlProperty('PMI_NATUREINT', 'Datatype', 'PGDIVERSMINIMUM')
    else
      if ((GetField('PMI_NATURE') = 'AGE') or (GetField('PMI_NATURE') = 'ANC')) then
      SetControlProperty('PMI_NATUREINT', 'Datatype', 'PGTYPEINTERTPS'); {FIN PT1-1}
    AlimGrilleIntervalle;
    // Fin Fiche saisie d'intervalles
    EnableChampsOngletIntervalle;
  end; // begin1

  //Gestion du mode d'ouverture de la fiche
//  AccesPredefini('TOUS', CEG, STD, DOS);
  LectureSeule := FALSE;

  if (Getfield('PMI_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
    SetControlEnabled('DETAILMINI', True); //PT21
    SetControlEnabled('BDelete', CEG);
  end else
    if (Getfield('PMI_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('DETAILMINI', True); //PT21
    SetControlEnabled('BDelete', STD);
  end
  else
    if (Getfield('PMI_PREDEFINI') = 'DOS') then
    begin
      LectureSeule := (DOS = False);  //PT21
      PaieLectureSeule(TFFiche(Ecran), (DOS = False)); //PT21
      SetControlEnabled('DETAILMINI', True); //PT21
      SetControlEnabled('BDelete', DOS);
    end;
  SetControlEnabled('BINSERT_LINE', (LectureSeule = False));
  SetControlEnabled('BADD_LINE', (LectureSeule = False));
  SetControlEnabled('BDEL_LINE', (LectureSeule = False));
  SetControlEnabled('BInsert', True);
  SetControlEnabled('PMI_PREDEFINI', False);
  SetControlEnabled('PMI_CONVENTION', False);
  SetControlEnabled('PMI_CODE', False);
  if DS.State in [dsInsert] then
  begin
    LectureSeule := FALSE;
    PaieLectureSeule(TFFiche(Ecran), False);
    SetControlEnabled('PMI_PREDEFINI', True);
    SetControlEnabled('PMI_CONVENTION', True);
    SetControlEnabled('PMI_CODE', True);
    SetControlEnabled('BInsert', False);
    SetControlEnabled('BDelete', False);
    SetControlEnabled('BINSERT_LINE', True);
    SetControlEnabled('BADD_LINE', True);
    SetControlEnabled('BDEL_LINE', True);
  end;
end;


procedure TOM_Minimumconvent.OnClose;
var
  init: word;
begin
  if NomFiche <> SaisieIntervalle then exit;
  if LectureSeule then exit;
  if TopUpdate then exit;
  if not OnFerme then Exit; // PT13
  Init := HShowMessage('1;' + Ecran.caption + ';Voulez VOUS enregistrer les modifications ?;N;YN;Y;', '', '');
  if Init = mrYes then // je réinitialise la grille
    OnUpdateRecord;
end;

//  PT3   05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer revu la fonction

procedure TOM_Minimumconvent.AlimGrilleIntervalle;
var
  i, j: INTEGER;
  St: string;
  Nombre: string; // PT3 String au lieu d'integer
  Q: TQuery;
begin // begin 1
  PDetail.Col := 0;
  PDetail.Row := 1;
  for i := 1 to PDetail.rowcount do
  begin
    PDETAIL.Cells[0, i] := '';
    PDETAIL.Cells[1, i] := '';
  end;
  TDETAIL := TOB.Create('detail minimumconvent', nil, -1);
  st := 'SELECT * FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') + '" AND PCP_CODE="' +
    GetField('PMI_CODE') + '" AND PCP_CONVENTION="' + GetField('PMI_CONVENTION') + '" AND PCP_NODOSSIER="' +
    Getfield('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + getfield('PMI_PREDEFINI') + '"';
  Q := OpenSql(st, TRUE);
  TDETAIL.LoadDetailDB('MINCONVPAIE', '', '', Q, FALSE, FALSE);
  FERME(Q);
  if TDETAIL.Detail.Count <= 0 then
  begin // begin 2
    PDetail.RowCount := 2;
    end else // begin 2
  begin // begin 3
    PDetail.RowCount := TDETAIL.Detail.Count + 1;
    AffecteComboGrille;
    for i := 0 to TDETAIL.Detail.Count - 1 do
    begin // begin 4
      Nombre := TDETAIL.Detail[i].GetValue('PCP_NBRE');
      st := PDetail.ColFormats[0];
      j := Pos('|', St);
      if j = 0 then j := 60;
      st := Copy(PDetail.ColFormats[0], 4, j - 4);
      if Copy(PDETAIL.ColFormats[0], 1, 3) = 'CB=' then
        PDETAIL.CellValues[0, i + 1] := Nombre
        //rechdom(Copy(PDetail.ColFormats[0], 4, j - 4), Nombre, false) {PT1-1} // PT12-1
      else PDETAIL.Cells[0, i + 1] := Nombre;

      St := TDETAIL.Detail[i].GetValue('PCP_TAUX');
      if Copy(PDETAIL.ColFormats[1], 1, 3) = 'CB=' then
       // PDETAIL.Cells[1, i + 1] := rechdom(Copy(PDetail.ColFormats[1], 4, 60), St, false) // PT12-1
       PDETAIL.CellValues[1, i + 1] := st
      else PDETAIL.Cells[1, i + 1] := st;
    end; // begin 4
  end; // begin 3
  TDETAIL.Free;
end; // begin 1


procedure TOM_Minimumconvent.InitGrilleIntervalle;
var
  i: INTEGER;
begin // begin 1
  PDetail.Col := 0;
  PDetail.Row := 1;
  PDetail.RowCount := 6;
  for i := 1 to PDetail.rowcount do
  begin
    PDETAIL.Cells[0, i] := '';
    PDETAIL.Cells[1, i] := '';
    PDETAIL.CellValues[0, i] := ''; // PT18
    PDETAIL.CellValues[1, i] := ''; // PT18
  end;
end; // begin 1



procedure TOM_Minimumconvent.ChangeGrid;
var
  Nat: string;
  i: integer;
begin
  PDetail := THGrid(GetControl('DETAILMINI'));
  if PDetail = nil then exit;
  Nat := GetField('PMI_TAUXOUMONTANT');
  PDetail.ColAligns[1] := taRightJustify;
  if Nat = 'T' then PDetail.ColFormats[1] := '#,##0.00000' else
    if Nat = 'Q' then PDetail.ColFormats[1] := '#,##0' else
    if Nat = 'M' then PDetail.ColFormats[1] := '#,##0.00' else PDetail.ColFormats[1] := '';

  if PDetail.ColFormats[1] <> '' then
    for i := 1 to PDetail.RowCount - 1 do
      PDETAIL.Cells[1, i] := FormatFloat(PDetail.ColFormats[1], Valeur(PDETAIL.Cells[1, i]));

  PDetail.Col := 0;
end;

procedure TOM_Minimumconvent.OnArgument(stArgument: string);
var
  Btn, BtnIns, BtnDel, BtnAdd, Bferme: TToolBarButton97;
  st, Arg: string;
  LabNature: THLabel;
begin
  inherited;
  AccesPredefini('TOUS', CEG, STD, DOS); //PT21

  if NomFiche = SaisieValeur then
  begin
    // recherche du 2ème argument de stArgument pour récupérer la nature de l'élément.
    Arg := stArgument;
    St := Trim(ReadTokenPipe(Arg, ';'));
    st := Trim(ReadTokenPipe(Arg, ';'));
    //  PT8   08/09/2003 PH V_421 Recuperation automatique de la nature
    LaNature := Trim(ReadTokenPipe(Arg, ';'));
    LaConv := Trim(ReadTokenPipe(Arg, ';')); // PT15
    if LaNature = '' then LaNature := St;
    if LaNature <> '' then
    begin
      LabNature := THLabel(GetControl('Nature_Val'));
      if LabNature <> nil then LabNature.caption := LaNature;
    end;
    PaieConceptPlanPaie(Ecran); // PT20
  end
  else PaieConceptTabMinPaie(Ecran); // PT20



  {DEB PT7}
  if GetControl('INTERVALLE') <> nil then
  begin
    Btn := TToolBarButton97(GetControl('BImprimer'));
    if Btn <> nil then
    begin
      setControlvisible('BIMPRIMER', TRUE); // CWAS
      Btn.OnClick := ClickImprimer;
    end;
  end;
  {FIN PT7}
  NewSaisie := True; // PT18
  PDetail := THGrid(GetControl('DETAILMINI'));
  BtnIns := TToolBarButton97(GetControl('BINSERT_LINE'));
  BtnDel := TToolBarButton97(GetControl('BDEL_LINE'));
  BtnAdd := TToolBarButton97(GetControl('BADD_LINE'));
  if BtnIns <> nil then BtnIns.OnClick := BTnInsClick;
  if BtnDel <> nil then BtnDel.OnClick := BTnDelClick;
  if BtnAdd <> nil then BtnAdd.OnClick := BTnAddClick;
  if PDetail <> nil then
  begin
    PDetail.OnCellExit := PDetailCellExit;
    PDetail.Oncellenter := PdetailCellEnter;
    PDetail.Onenter := entrergrille;                  // PT18
    PDetail.OnElipsisClick := PDetailElipsisClick;    // PT18
    PDetail.OnDblClick := Entrergrille;    // PT18
    PDetail.Onclick := Entrergrille; // PT18
   end;
  BFerme := TToolBarButton97(GetControl('BFERME'));
  if BFerme <> nil then BFerme.OnMouseEnter := Ferme;
  //  PT9   17/09/2003 PH V_421 FQ 10080 Traitement datatype zone PMI_PROFIL pour affectation bonne tablette
  if GetControl('PMI_PROFIL') <> nil then SetControlProperty('PMI_PROFIL', 'DataType', 'PGPROFIL');
  SetControlEnabled('PMI_NATURE', FALSE);
end;

procedure TOM_Minimumconvent.OnChangeField(F: TField);
var
  Z: THLabel;
begin
  inherited;
  Z := THLabel(GetControl('INTERVALLE'));
  if Z <> nil then NomFiche := SaisieIntervalle
  else NomFiche := SaisieValeur;
  // Dispatch den 2 procedure OnChange
  if NomFiche = SaisieValeur then SaisieValeur_OnChange(F)
  else SaisieIntervalle_OnChange(F);

  //Visibilité des rubriques prédefini
  if (ds.state in [dsBrowse]) then
  begin
    if LectureSeule then
    begin
      PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', False);
    end;
    SetControlEnabled('BINSERT_LINE', (LectureSeule = False));
    SetControlEnabled('BADD_LINE', (LectureSeule = False));
    SetControlEnabled('BDEL_LINE', (LectureSeule = False));
    SetControlEnabled('PMI_NATUREINT', (LectureSeule = False));
    SetControlEnabled('PMI_NATURERESULT', (LectureSeule = False));
//    SetControlEnabled('DETAILMINI', (LectureSeule = False));
    SetControlEnabled('BDelete', (LectureSeule = False));

    SetControlEnabled('BInsert', True);
    SetControlEnabled('PMI_PREDEFINI', False);
    SetControlEnabled('PMI_CONVENTION', False);
    SetControlEnabled('PMI_CODE', False);
  end;
  //  SetControlEnabled ('PMI_NATURE', FALSE ) ;

end;

procedure TOM_Minimumconvent.Ferme(Sender: tobject);
begin
  TopUpdate := false;
end;

procedure TOM_Minimumconvent.SaisieValeur_OnChange(F: TField);
var
  Pred: string;
begin
  if (F.FieldName = 'PMI_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Pred := GetField('PMI_PREDEFINI');
    if Pred = '' then exit;
//PT21    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de rubrique prédéfinie Standard', 'Accès refusé');
//PT21      Pred := 'DOS';
//PT21      SetControlProperty('PMI_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PMI_PREDEFINI') then SetField('PMI_PREDEFINI', pred);
  end;
 { DEB PT12 }
  if (LaNature='COE') AND (F.FieldName = 'PMI_CODE') and (DS.State = dsinsert)  then
    Begin
    If GetField('PMI_LIBELLE') <> GetField('PMI_CODE') then
      SetField('PMI_LIBELLE', GetField('PMI_CODE'));
    End;
  { FIN PT12 }  
end;
//------------------------------------------------------------------------------
//   Traitemeent de l'événement OnChange sur la grille de gestion d'un intervalle
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.SaisieIntervalle_OnChange(F: TField);
var
  Nat, rubrique, mes, temprub, vide, pred: string;
{$IFNDEF EAGLCLIENT}
  MIPMI_NATUREINT, MIPMI_NATURERESULT, Nature: THDBVALCOMBOBOX;
{$ELSE}
  MIPMI_NATUREINT, MIPMI_NATURERESULT, Nature: THVALCOMBOBOX;
{$ENDIF}
  init: Word;
  //NatureResult, snature: string;     PT18
  snature: string;                //   PT18
  icode: integer;
  OkRub: boolean;
begin
  vide := '';
  PDetail := THGrid(GetControl('DETAILMINI'));

  if (F.FieldName = 'PMI_NATURE') then
  begin
{$IFNDEF EAGLCLIENT}
    Nature := THDBVALCOMBOBOX(GetControl('PMI_NATURE'));
{$ELSE}
    Nature := THVALCOMBOBOX(GetControl('PMI_NATURE'));
{$ENDIF}
    if Nature <> nil then
    begin // B1
      snature := GetField('PMI_NATURE');
      if ((sNature <> Nature.Value) and (sNature <> '')) then
      begin // B2
        Init := HShowMessage('1;Changement de nature;Cette modification entraînera la réinitialisation de la table. Voulez-vous continuer ?;Q;YN;N;N;', '', '');
        if Init = mrYes then // je réinitialise la grille
        begin // B3
          ExecuteSQL('DELETE FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') + '" AND PCP_CODE="' + GetField('PMI_CODE') + '" AND PCP_CONVENTION="' +
            GetField('PMI_CONVENTION') + '"'
            + ' AND PCP_NODOSSIER ="' + GetField('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + GetField('PMI_PREDEFINI') + '"');
          setField('PMI_NATUREINT', '');
          NewSaisie := True; // pt18
          InitGrilleIntervalle;
          EnableChampsOngletIntervalle;
          exit;
        end // B3
        else
          NATURE.Value := GetField('PMI_NATURE');
      end; //B2
      if Ecran <> nil then Ecran.Caption := 'Table ' + RechDom('PGNATURECOEFF', Nature.value, FALSE) + ' : ';
    end; //B1

    // Alimentation de la tablette d'intervalle en fct de la nature
    Nat := GetField('PMI_NATURE');
{$IFNDEF EAGLCLIENT}
    MIPMI_NATUREINT := THDBVALCOMBOBOX(GetControl('PMI_NATUREINT'));
{$ELSE}
    MIPMI_NATUREINT := THVALCOMBOBOX(GetControl('PMI_NATUREINT'));
{$ENDIF}
    if Nat = 'DIV' then MIPMI_NATUREINT.datatype := 'PGTYPEINTERDIV'
    else
      //  PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
      if Nat = 'DIW' then MIPMI_NATUREINT.datatype := 'PGDIVERSMINIMUM'
    else
      if ((Nat = 'AGE') or (Nat = 'ANC')) then MIPMI_NATUREINT.datatype := 'PGTYPEINTERTPS';
  end;

  if (F.FieldName = 'PMI_CODE') then
  begin
    Rubrique := Getfield('PMI_CODE');
    if (Rubrique = '') then exit;
    if ((isnumeric(Rubrique)) and (Rubrique <> '   ')) then
    begin
      iCode := strtoint(trim(Rubrique));
      TempRub := ColleZeroDevant(iCode, 3);
      if (DS.State = dsinsert) and (TempRub <> '') and (GetField('PMI_PREDEFINI') <> '') then
      begin
        OKRub := TestRubrique(GetField('PMI_PREDEFINI'), TempRub, 100);
        if (OkRub = False) or (Rubrique = '000') then
        begin
          Mes := MesTestRubrique('MIN', GetField('PMI_PREDEFINI'), 100);
          HShowMessage('2;Code Erroné: ' + TempRub + ' ;' + mes + ';W;O;O;;;', '', '');
          TempRub := '';
        end;
      end;
      if TempRub <> Rubrique then
      begin
        SetField('PMI_CODE', TempRub);
        SetFocusControl('PMI_CODE');
      end;
    end;
  end;

  if (F.FieldName = 'PMI_PREDEFINI') and (DS.State = dsinsert) then
  begin
    Rubrique := GetField('PMI_CODE');
    Pred := GetField('PMI_PREDEFINI');
    if Pred = '' then exit;
//PT21    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de code prédéfini CEGID', 'Accès refusé');
//PT21      Pred := 'DOS';
//PT21      SetControlProperty('PMI_PREDEFINI', 'Value', Pred);
    end;

    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de code prédéfini Standard', 'Accès refusé');
//PT21      Pred := 'DOS';
//PT21      SetControlProperty('PMI_PREDEFINI', 'Value', Pred);
    end;

    if (rubrique <> '') and (rubrique <> '   ') and (pred <> '') then
    begin
      OKRub := TestRubrique(pred, rubrique, 100);
      if (OkRub = False) or (Rubrique = '000') then
      begin
        Mes := MesTestRubrique('MIN', pred, 100);
        HShowMessage('2;Code Erroné: ' + Rubrique + ' ;' + mes + ';W;O;O;;;', '', '');
        SetField('PMI_CODE', vide);
        if Pred <> GetField('PMI_PREDEFINI') then SetField('PMI_PREDEFINI', pred);
        SetFocusControl('PMI_CODE');
        exit;
      end;
    end;

    if Pred <> GetField('PMI_PREDEFINI') then SetField('PMI_PREDEFINI', pred);
  end;


  // Modification des titres de la grid
  if (F.FieldName = 'PMI_NATUREINT') then
  begin
{$IFNDEF EAGLCLIENT}
    MIPMI_NATUREINT := THDBVALCOMBOBOX(GetControl('PMI_NATUREINT'));
{$ELSE}
    MIPMI_NATUREINT := THVALCOMBOBOX(GetControl('PMI_NATUREINT'));
{$ENDIF}
    if MIPMI_NATUREINT <> nil then
    begin // B1
      if NatureInt = '' then NatureInt := MIPMI_NATUREINT.Value; // PT12
      if ((NatureInt <> MIPMI_NATUREINT.Value) and
        (NatureInt <> '')) then
      begin // B2
        Init := HShowMessage('1;Changement d''intervalle;Cette modification entrainera le réinitialisation de la grille détail. Voulez vous continuer ?;Q;YN;N;N;', '', '');
        if Init = mrYes then // je réinitialise la grille
        begin // B3
          ExecuteSQL('DELETE FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') + '" AND PCP_CODE="' + GetField('PMI_CODE') + '" AND PCP_CONVENTION="' +
            GetField('PMI_CONVENTION')
            + '" AND PCP_NODOSSIER ="' + GetField('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + GetField('PMI_PREDEFINI') + '"');
          NewSaisie := True; // PT18
          InitGrilleIntervalle;
          EnableChampsOngletIntervalle;
          NatureInt := GetField('PMI_NATUREINT'); // PT12
         // exit;        PT18
        end // B3
        else
          MIPMI_NATUREINT.Value := GetField('PMI_NATUREINT');
      end //B2
      else AffecteComboGrille; // PT13 Pour forcer affectation en creation et 1ere saisie
    end; //B1

    Nat := GetField('PMI_NATURE');
    if PDetail <> nil then
      if ((Nat = 'AGE') or (Nat = 'ANC')) then
        PDetail.Cells[0, 0] := RechDom('PGTYPEINTERTPS', GetField('PMI_NATUREINT'), FALSE) else
        if Nat = 'DIV' then PDetail.cells[0, 0] := RechDom('PGTYPEINTERDIV', GetField('PMI_NATUREINT'), FALSE)
      else
        //  PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
        if Nat = 'DIW' then PDetail.cells[0, 0] := RechDom('PGDIVERSMINIMUM', GetField('PMI_NATUREINT'), FALSE);
  end;

  if (F.FieldName = 'PMI_NATURERESULT') then
  begin //B0
{$IFNDEF EAGLCLIENT}
    MIPMI_NATURERESULT := THDBVALCOMBOBOX(GetControl('PMI_NATURERESULT'));
{$ELSE}
    MIPMI_NATURERESULT := THVALCOMBOBOX(GetControl('PMI_NATURERESULT'));
{$ENDIF}
    if MIPMI_NATURERESULT <> nil then
    begin //B1
   //   NatureResult := GetField('PMI_NATURERESULT');            PT18
  //    if ((NatureResult <> MIPMI_NATURERESULT.Value) and       PT18
  //   (NatureResult <> '')) then                                PT18
  // PT18
    if NatureResult = '' then NatureResult := MIPMI_NATURERESULT.Value;
      if ((NatureResult <> MIPMI_NATURERESULT.Value) and
        (NatureResult <> '')) then
  // fin PT18
      begin //b2
        Init := HShowMessage('1;Changement de type de résultat;Cette modification entrainera le réinitialisation de la grille détail. Voulez vous continuer ?;Q;YN;N;N;', '', '');
        if Init = mrYes then // je réinitialise la grille
        begin // B3
          ExecuteSQL('DELETE FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') + '" AND PCP_CODE="' + GetField('PMI_CODE') + '" AND PCP_CONVENTION="' +
            GetField('PMI_CONVENTION') + '"'
            + ' AND PCP_NODOSSIER ="' + GetField('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + GetField('PMI_PREDEFINI') + '"');
          NewSaisie := True; // PT18 
          InitGrilleIntervalle;
          EnableChampsOngletIntervalle;
          NatureResult := GetField('PMI_NATURERESULT'); // PT18
        //  exit;     PT18
        end // b3
        else
          MIPMI_NATURERESULT.Value := GetField('PMI_NATURERESULT');
      end //B2
      else AffecteComboGrille; // PT13 Pour forcer affectation en creation et 1ere saisie
    end; //B1
    if PDetail <> nil then
      PDetail.Cells[1, 0] := RechDom('PGGESTIONMINI', GetField('PMI_NATURERESULT'), FALSE);
  end;
  if (DS.State in [dsInsert]) then  EnableChampsOngletIntervalle; // PT19
end;


(* 25/07/2006 Mise en commentaire, incompatible delphi7 et non utilisé
procedure TOM_Minimumconvent.Ipmi_codeExit(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  MdCode: THdbEdit;
{$ELSE}
  MdCode: THEdit;
{$ENDIF}
  i: integer;
  sCode : string;
  num: boolean;
begin
{$IFNDEF EAGLCLIENT}
  MdCode := THdbEdit(GetControl('PMI_CODE'));
  ;
{$ELSE}
  MdCode := THEdit(GetControl('PMI_CODE'));
  ;
{$ENDIF}
  if Mdcode <> nil then
  begin
    num := false;
    for i := 1 to Length(Mdcode.Text) do
      Begin
      if isnumeric(Mdcode.Text[i]) then
      begin
        num := true;
        break;
      end;
      End;
    if not num then
    begin
      HShowMessage('1;Attention;Le code doit être saisi et de type numérique;W;Y;Y;N;', '', '');
      SetFocusControl('PMI_CODE');
      exit;
    end;
    i := strtoint(trim(MdCode.Text));
    sCode := ColleZeroDevant(i, length(MdCode.Text));
    if sCode <> MdCode.Text then MdCode.Text := scode;
  end;
end;                    *)
// -----------------------------------------------------------------------------
//  On rend les champs de l'onglet saisissables ou non en fonction deu contexte
// -----------------------------------------------------------------------------

procedure TOM_Minimumconvent.EnableChampsOngletIntervalle;
var
  ch: string;
  s1, s2, s4, s5, s6, creation: boolean;
begin
  // Champ Nature
  Creation := (DS.State = dsinsert);
  ch := GetField('PMI_NATURE');
  SetControlEnabled('PMI_CONVENTION', creation);
  SetControlEnabled('PMI_NATUREINT', (Ch <> ''));
  SetControlEnabled('PMI_SENS', (Ch <> ''));
  SetControlEnabled('PMI_NATURERESULT', (Ch <> ''));
  // La grille de saisie ne devient accessible que lorsque l'ensemble des éléments
  // décrivant l'intervalle a été défini.
  s1 := (GetField('PMI_NATURE') <> '');
  s2 := (GetField('PMI_CODE') <> '');
  s4 := (GetField('PMI_NATUREINT') <> '');
  s5 := (GetField('PMI_SENS') <> '');
  s6 := (GetField('PMI_NATURERESULT') <> '');
  SetControlEnabled('DETAILMINI', (s1 and s2 and s4 and s5 and s6));
  AffecteComboGrille;

end;

procedure TOM_Minimumconvent.AffecteComboGrille;
var
{$IFNDEF EAGLCLIENT}
  NatureInt: THDbValComboBox;
{$ELSE}
  NatureInt: THValComboBox;
{$ENDIF}
  Natr, st: string;
begin

PDetail.colcombo := -1 ;    // PT18
if assigned(PDetail.ValCombo) then
  PDetail.ValCombo.Hide; //PT18

{$IFNDEF EAGLCLIENT}
  NatureInt := THDbValComboBox(getcontrol('PMI_NATUREINT'));
{$ELSE}
  NatureInt := THValComboBox(getcontrol('PMI_NATUREINT'));
{$ENDIF}
  if NatureInt <> nil then
  begin
    if NATUREINT.datatype = 'PGTYPEINTERDIV' then
    begin
      st := GetControlText('PMI_CONVENTION'); // PT12
      if Natureint.value = 'COE' then
        PDETAIL.colformats[0] := 'CB=PGCOEFFICIENT'
      else if Natureint.value = 'IND' then
        PDETAIL.colformats[0] := 'CB=PGINDICE'
      else if Natureint.value = 'QUA' then
        PDETAIL.colformats[0] := 'CB=PGQUALIFICATION'
      else if Natureint.value = 'NIV' then
        PDETAIL.colformats[0] := 'CB=PGNIVEAU';
      // DEB PT12
      if (st <> '') AND (st <> '000') then PDETAIL.colformats[0] := PDETAIL.colformats[0] + '| AND (PMI_CONVENTION="000" OR PMI_CONVENTION="' + st + '")';
//      else PDETAIL.colformats[0] := PDETAIL.colformats[0] + '| AND (PMI_CONVENTION="000")'; // PT17
      // FIN PT12
    end
    else
      if NATUREINT.datatype = 'PGDIVERSMINIMUM' then
    begin
         if Natureint.value = 'ELT' then             //  PT18
     //     PDETAIL.colformats[0] := 'CB=PGELEMENTNAT'  PT18
       PDETAIL.colformats[0] := '' // PT18
      else if Natureint.value = 'VAR' then
       PDETAIL.colformats[0] := 'CB=PGVARIABLE'
      else if Natureint.value = 'OR1' then
        PDETAIL.colformats[0] := 'CB=PGTRAVAILN1'
      else if Natureint.value = 'OR2' then
        PDETAIL.colformats[0] := 'CB=PGTRAVAILN2'
      else if Natureint.value = 'OR3' then
        PDETAIL.colformats[0] := 'CB=PGTRAVAILN3'
      else if Natureint.value = 'OR4' then
        PDETAIL.colformats[0] := 'CB=PGTRAVAILN4'
      else if Natureint.value = 'ETB' then
        PDETAIL.colformats[0] := 'CB=TTETABLISSEMENT'
      else if Natureint.value = 'STA' then
        PDETAIL.colformats[0] := 'CB=PGCODESTAT'
      else if Natureint.value = 'TC1' then
        PDETAIL.colformats[0] := 'CB=PGLIBREPCMB1'
      else if Natureint.value = 'TC2' then
        PDETAIL.colformats[0] := 'CB=PGLIBREPCMB2'
      else if Natureint.value = 'TC3' then
        PDETAIL.colformats[0] := 'CB=PGLIBREPCMB3'
      else if Natureint.value = 'TC4' then
        PDETAIL.colformats[0] := 'CB=PGLIBREPCMB4'
      else if (Copy(Natureint.value, 1, 2) = 'DT') then
      begin
        PDETAIL.ColTypes[0] := 'D';
        PDETAIL.colformats[0] := ShortDateFormat;
      end;
    end
    else
      if NATUREINT.datatype = 'PGTYPEINTERTPS' then PDETAIL.colformats[0] := '';
  end;

  if NatureInt <> nil then
  begin
    NatR := GetField('PMI_NATURERESULT');
    // PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
    if ((NatR = 'DIV') or (NatR = 'DIW') or (NatR = 'ANC') or (NatR = 'AGE')) then
      Pdetail.colformats[1] := 'CB=PGTABINT' + NatR
    else if NatR = 'VAR'
      then Pdetail.colformats[1] := 'CB=PGVARIABLE'
    else if NatR = 'ELT'
      //  PT4   05/02/2003 PH V595 Les tables rendent comme valeur un élément national (New)
   // then Pdetail.colformats[1] := 'CB=PGELEMENTNAT'  PT18
    then Pdetail.colformats[1] := ''    // PT18
      else Pdetail.colformats[1] := '';  
  end;
end;

procedure TOM_Minimumconvent.OnUpdateRecord;
var
  st: string;
begin
  inherited;
  LastError := 0;
  OnFerme := False;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PMI_CONVENTION')
  else
    if (DerniereCreate = GetField('PMI_CONVENTION')) then OnFerme := True; // le bug arrive on se casse !!!
  TopUpdate := true;
  if NomFiche = SaisieIntervalle then ControleZonesIntervalle;
  if NomFiche = SaisieValeur then ControleZonesValeur;
  if (DS.State = dsinsert) then
  begin
    if GetField('PMI_CONVENTION') = '' then SetField('PMI_CONVENTION', 000);
    if (GetField('PMI_PREDEFINI') <> 'DOS') then SetField('PMI_NODOSSIER', '000000')
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    else SetField('PMI_NODOSSIER', PgRendNoDossier());
  end;
  //  PT2   05/12/2002 PH V591 FQ 10209 test existance code lors de la creation
  if (Getfield('PMI_CODE') = '') or (Getfield('PMI_CODE') = '   ') then
  begin
    LastError := 1;
    PgiBox('Vous devez renseigner un code', Ecran.Caption);
  end;
  // DEB PT13
  if (Getfield('PMI_NATURE') = '') then
  begin
    LastError := 2;
    PgiBox('Vous devez renseigner une nature', Ecran.Caption);
  end;

  //DEB PT21
    if (GetField('PMI_PREDEFINI') = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de code prédéfini CEGID', 'Accès refusé');
      LastError := 3;
    end;
    if (GetField('PMI_PREDEFINI') = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez créer de code prédéfini Standard', 'Accès refusé');
      LastError := 3;
    end;
  //FIN PT21

  // FIN PT13
  // DEB PT11
  //  DEB PT11-1 Rajout de la nature de la table de travail
  // DEB PT11-2 Plus de controle
{    st := 'SELECT PMI_CODE FROM MINIMUMCONVENT WHERE PMI_CODE="' + GetField('PMI_CODE') +
      '" AND PMI_NATURE="' + GetField('PMI_NATURE') + '" AND ## PMI_PREDEFINI##';
//    '" AND PMI_NATUREINT <> "'+ GetField ('PMI_NATUREINT')+
    if ExisteSQL(st) and (DS.State in [dsInsert]) then
    begin
      LastError := 1;
      PgiError('Vous avez déjà une table existante avec les mêmes nature, code et table de travail', Ecran.Caption);
    end;}
  // FIN PT11 et  PT11-1 et PT11-2
    //Rechargement des tablettes
  if (LastError = 0) and (Getfield('PMI_CODE') <> '') and (Getfield('PMI_LIBELLE') <> '') then
      {$IFNDEF EAGLCLIENT}
      ChargementTablette(TFFiche(Ecran).TableName, '');
      {$ELSE}
      ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');   //PT14
      {$ENDIF}
end;

procedure TOM_Minimumconvent.ControleZonesValeur;
begin
  // Contrôle saisie Convention collective
  if (GetField('PMI_CONVENTION') = '') then
  begin
    LastError := 201;
    LastErrorMsg := 'Vous devez renseigner une convention collective';
    SetFocusControl('PMI_CONVENTION');
  end
  else
    // Contrôle présence Code
    if (GetField('PMI_CODE') = '') or (GetField('PMI_CODE') = '  ') then
  begin
    LastError := 202;
    LastErrorMsg := 'Vous devez renseigner un code';
    SetFocusControl('PMI_CODE');
  end
  else
    // Contrôle présence libellé
    if (GetField('PMI_LIBELLE') = '') then
  begin
    LastError := 203;
    LastErrorMsg := 'Vous devez renseigner un libellé';
    SetFocusControl('PMI_LIBELLE');
  end;
end;

procedure TOM_Minimumconvent.ControleZonesIntervalle;
var
  StNombre: string;
  i: INTEGER;
  St, St1: string;
  Predef: string;
  Q: TQuery;
  Tob_MinConvPaie, T1: TOB;
begin
  // Contrôle présence nature de l'intervalle
  if (GetField('PMI_NATUREINT') = '') then
  begin
    LastError := 101;
    LastErrorMsg := 'Vous devez renseigner une nature d''intervalle';
    SetFocusControl('PMI_NATUREINT');
  end
  else
    // Contrôle présence sens
    if (GetField('PMI_SENS') = '') then
  begin
    LastError := 102;
    LastErrorMsg := 'Vous devez renseigner un sens';
    SetFocusControl('PMI_SENS');
  end
  else
    // Contrôle présence nature du résultat
    if (GetField('PMI_NATURERESULT') = '') then
  begin
    LastError := 103;
    LastErrorMsg := 'Vous devez renseigner une nature de résultat';
    SetFocusControl('PMI_NATURERESULT');
  end
  else
    // Contrôle présence Libellé
    if (GetField('PMI_LIBELLE') = '') then
  begin
    LastError := 104;
    LastErrorMsg := 'Vous devez renseigner un libellé';
    SetFocusControl('PMI_LIBELLE');
  end
  else
    // Contrôle présence Code
    if (GetField('PMI_CODE') = '') then
  begin
    LastError := 104;
    LastErrorMsg := 'Vous devez renseigner un code';
    SetFocusControl('PMI_CODE');
  end;

  Predef := GetField('PMI_PREDEFINI');
  if (Predef <> 'CEG') and (Predef <> 'DOS') and (Predef <> 'STD') then
  begin
    LastError := 105;
    LastErrorMsg := 'Vous devez renseigner le champ prédéfini';
    SetFocusControl('PMI_PREDEFINI');
  end;

  st := 'SELECT PMI_PREDEFINI FROM MINIMUMCONVENT WHERE ' +
    ' PMI_NATURE="' + GetField('PMI_NATURE') + '" AND PMI_CODE="' + GetField('PMI_CODE') + '" AND PMI_CONVENTION="' + GetField('PMI_CONVENTION') + '"'
    + ' AND PMI_TYPENATURE ="INT" AND PMI_NATURE="' + GetField('PMI_NATURE') + '"';
  Q := Opensql(st, true);
  if DS.State in [dsInsert] then
    if not Q.eof then
    begin
      LastError := 106;
      LastErrorMsg := 'Cette table dossier existe déjà';
      SetFocusControl('PMI_CODE');
    end;
  Ferme(Q);

  TopUpdate := true;
  if (PDetail = nil) then exit;
  //  SetField('PMI_CODE',Thedit(getcontrol('IPMI_CODE')).text);   // beurk
  SetField('PMI_CODE', GetControlText('PMI_CODE')); // ?????
  st1 := GetField('PMI_NODOSSIER');
  if GetField('PMI_PREDEFINI') <> 'DOS' then st1 := '000000';
  ExecuteSQL('DELETE FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') +
    '" AND PCP_CODE="' + GetField('PMI_CODE') + '" AND PCP_CONVENTION="' + GetField('PMI_CONVENTION') +
    '" AND PCP_NODOSSIER ="' + GetField('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + GetField('PMI_PREDEFINI') + '"');
  if PDetail.RowCount - 1 >= 1 then
  begin
    Tob_MinConvPaie := TOB.Create('Ma_TOB_MinConvpaie', nil, -1); // PT14
    for i := 1 to PDetail.RowCount - 1 do
    begin {Deb PT1-1 pls modif. On récupére le cellvalues au lieu du cells, nouvel var StNombre}
      if ((PDetail.Cells[0, i]) = '') then Continue;
      // PT5   06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
      if (GetField('PMI_NATURE') <> 'DIV') and (GetField('PMI_NATURE') <> 'DIW') then
        StNombre := PDetail.Cells[0, i]
      else StNombre := PDetail.CellValues[0, i];
      St := PDetail.CellValues[1, i];
      //  PT3   05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer revu la fonction
      //  PT6   16/05/2003 PH V_421 Traitement de la date de modif dans ma table MINCONVPAIE
      // DEB PT14
      T1 := TOB.Create('MINCONVPAIE', Tob_MinConvPaie, -1);
      T1.PutValue('PCP_NODOSSIER', St1);
      T1.PutValue('PCP_PREDEFINI', GetField('PMI_PREDEFINI'));
      T1.PutValue('PCP_NATURE', GetField('PMI_NATURE'));
      T1.PutValue('PCP_CODE', GetField('PMI_CODE'));
      T1.PutValue('PCP_CONVENTION', GetField('PMI_CONVENTION'));
      T1.PutValue('PCP_NBRE', StNombre);
      T1.PutValue('PCP_TAUX', st);
      T1.PutValue('PCP_NOMBRE', 0);
      //      T1.PutValue('PCP_DATEMODIF', );
    end; {Fin PT1-1}
    Tob_MinConvPaie.InsertDb(nil, FALSE);
    // FIN PT14
  end;
end;

procedure TOM_Minimumconvent.PDetailCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
//  PDetail.Showcombo(Acol, Arow);    PT18

    // pt18
    PDetail.ElipsisButton := false;
 if (PDetail.col = 0) and (Getfield('PMI_NATUREINT') = 'ELT')
  then begin
  PDetail.ElipsisButton := TRUE;
  PDetail.ColFormats[0]:= '';
  if assigned(PDetail.ValCombo) then
     PDetail.ValCombo.Hide; //PT18
  end;
 if (PDetail.col = 1) and (Getfield('PMI_NATURERESULT') = 'ELT')
  then begin
  PDetail.ElipsisButton := TRUE;
  PDetail.ColFormats[1]:= '';
  if assigned(PDetail.ValCombo) then
     PDetail.ValCombo.Hide; //PT18
  end;
  // pt18
end;

procedure TOM_Minimumconvent.PDetailCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  i: Integer;
  st: string;
begin
  if LectureSeule = False then //PT21
  begin
    ForceUpdate ;  //  PT15
    if (ACol = 0) then
     begin
     if PDetail.colFormats[Acol] = '' then
     begin // Si on n'est pas lié à une combo, vérifier la numéricité et l'ascendance
        // des intervalles
       if (Natureint <> 'ELT') then     //PT18  si ce n'est pas ELT national
       begin                           // PT18
        if IsNumeric(THGrid(Sender).Cells[0, THGrid(Sender).Row]) then
        begin
         st := GetControlText('PMI_NATUREINT');
          if (st = 'MOI') or (St = 'ANN') or (St = 'JOU') then
            THGrid(Sender).Cells[0, THGrid(Sender).Row] := ColleZeroDevant(Round(Valeur(THGrid(Sender).Cells[0, THGrid(Sender).Row])), 3);
        end;
        if (THGrid(Sender).Row > 1) and
          (THGrid(Sender).Cells[0, THGrid(Sender).Row] <> '') and
        (Valeur(THGrid(Sender).Cells[0, THGrid(Sender).Row]) <= Valeur(THGrid(Sender).Cells[0, THGrid(Sender).Row - 1])) then
        begin
          LastError := 2;
          LastErrorMsg := 'Les intervalles doivent être saisis en ordre croissant';
          CheckError;
          cancel := TRUE;
        end;
        // PT10  11/05/2004 PH V_50  FQ 10961 Message controle des bornes
        if (ARow > 0) and (IsNumeric(THGrid(Sender).Cells[0, ARow]) and (ValeurI(THGrid(Sender).Cells[0, ARow]) > 999)) then
        begin
          LastError := 2;
          LastErrorMsg := 'La valeur des bornes doit être inférieure à 1000';
          CheckError;
          cancel := TRUE;
        end;

        if (ARow > 0) and (IsNumeric(THGrid(Sender).Cells[0, ARow]) = FALSE) then
        begin
          LastError := 2;
          // PT10  11/05/2004 PH V_50  FQ 10961 Message controle des bornes
          LastErrorMsg := 'Les bornes doivent être numériques et inférieures à 1000';
          CheckError;
          cancel := TRUE;
        end;
     end;
      end;
   end;
    // Tests de controle de dédoublonnage des bornes
    for i := 1 to THGrid(Sender).RowCount do
    begin
      if (THGrid(Sender).Cells[0, THGrid(Sender).Row] = THGrid(Sender).Cells[0, i]) and (i <> THGrid(Sender).Row) then
      begin
        if (THGrid(Sender).Cells[0, THGrid(Sender).Row]) <> '' then
        begin
          LastError := 2;
          LastErrorMsg := 'Vous ne pouvez pas avoir deux bornes identiques';
          CheckError;
          THGrid(Sender).Col := 0;
          cancel := TRUE;
        end;
      end;
    end;
  end;
end;

// PT18
procedure TOM_Minimumconvent.Entrergrille(Sender: TObject);
var Ligne , colonne : Integer;
    Cancel : Boolean;
begin
   If NewSaisie then
     begin
          Ligne := Pdetail.Row;
          colonne := Pdetail.col;
          Cancel := False;
          PdetailCellEnter(Pdetail,colonne,Ligne,Cancel);
          NewSaisie := False;
     end;
end;
//  PT18

// PT18
procedure TOM_Minimumconvent.PDetailElipsisclick(Sender: TObject);
begin
 LookupList(PDetail,'Elément national','ELTNATIONAUX','DISTINCT PEL_CODEELT','PEL_LIBELLE','','PEL_CODEELT',TRUE,-1);
end;
// PT18

//------------------------------------------------------------------------------
//                       Suppression D'UNE LIGNE DANS LE THGRID
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.BTnDelClick(Sender: TObject);
begin
  if PDetail.Row <> 0 then PDETAIL.DeleteRow(PDETAIL.Row);
end;
//------------------------------------------------------------------------------
//                       Insertion D'UNE LIGNE DANS LE THGRID
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.BTnInsClick(Sender: TObject);
begin
  PDETAIL.InsertRow(PDETAIL.Row);
end;
//------------------------------------------------------------------------------
//                       AJOUT D'UNE LIGNE DANS LE THGRID
//------------------------------------------------------------------------------

procedure TOM_Minimumconvent.BTnAddClick(Sender: TObject);
begin
  PDetail.RowCount := PDetail.RowCount + 5;
  PDetail.Col := 0;
end;

////                       Mes erreurs                                      ////
{
LastError;LastErrorMsg;Onglet
------------------------- Saisie par Valeur
201 ; LastErrorMsg:='Vous devez renseigner une convention collective' ;
202 ; LastErrorMsg:='Vous devez renseigner un code' ;
203 ; LastErrorMsg:='Vous devez renseigner un libellé' ;
204 ; LastErrorMsg:='le code doit être saisi et de type numérique';
------------------------- Saisie par Intervalle
101 ; LastErrorMsg:='Vous devez renseigner une nature d''intervalle' ;
102 ; LastErrorMsg:='Vous devez renseigner un sens' ;
103 ; LastErrorMsg:='Vous devez renseigner une nature de résultat' ;

}

procedure TOM_Minimumconvent.OnDeleteRecord;
begin
  inherited;
  {$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
  {$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), ''); //PT14
  {$ENDIF}
  {PT1-2}
  ExecuteSQL('DELETE FROM MINCONVPAIE WHERE PCP_NATURE="' + GetField('PMI_NATURE') +
    '" AND PCP_CODE="' + GetField('PMI_CODE') + '" AND PCP_CONVENTION="' + GetField('PMI_CONVENTION') +
    '" AND PCP_NODOSSIER ="' + GetField('PMI_NODOSSIER') + '" AND PCP_PREDEFINI ="' + GetField('PMI_PREDEFINI') + '"');
end;

procedure TOM_Minimumconvent.OnAfterUpdateRecord;
begin
  inherited;
  if OnFerme then Ecran.Close;
end;

{DEB PT7}

procedure TOM_Minimumconvent.ClickImprimer(sender: Tobject);
var
  Pages: TPageControl;
{$IFDEF EAGLCLIENT}
  StPages: string;
{$ENDIF}
begin
  Pages := TPageControl(GetControl('PAGES'));
  if Pages = nil then exit;
{$IFNDEF EAGLCLIENT}
  LanceEtat('E', 'PGA', 'PMN', True, False, False, Pages, '', '', False);
{$ELSE}
  StPages := AglGetCriteres(Pages, FALSE);
  LanceEtat('E', 'PGA', 'PMN', True, False, False, nil, '', '', False, 0, StPages);
{$ENDIF}
end;
{FIN PT7}

initialization
  registerclasses([TOM_MINIMUMCONVENT]);

end.


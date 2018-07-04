{***********UNITE*************************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... : 16/02/2007
Description .. : Source TOF de la FICHE : CPCONTROLCAISSE ()
Suite ........ : Reprise de l'ancien module CtrlCai.pas non compatible 
Suite ........ : CWAS
Mots clefs ... : TOF;CPCONTROLCAISSE
*****************************************************************}
Unit CPCONTROLCAISSE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,         //Agllancefiche
     {$IFNDEF DBXPRESS}
        dbtables,
     {$ELSE}
        uDbxDataSet,
     {$ENDIF}
     mul,
{$else}
     MainEAgl,      //AglLancefiche
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     AGLinit,         //thedata
     CritEdt,         // ClassCritEdt
     uTofCPGLGene,    //CPLanceFiche_CPGLGENE
     Graph,           //Visugraph
     Windows,         //Formkeydown
     Menus,           //popup
     UObjFiltres,     //TObjFiltre
     Hcompte,         //THCPTEDIT
     HTb97,           //TToolBarButton97
     HEnt1,           //STRFMONTANT
     Ent1,            //TfichierBase, TZoomTable, EXRF
     UTOF ;

Const
  //les colonnes de la grille
  ClFixe                : integer = 0;
  ClDate                : integer = 1;
  ClTotD                : integer = 2;
  CltotC                : integer = 3;
  ClSolde               : integer = 4;
  ClCumul               : integer = 5;

  StrErrorCpt           : Array [0..2] of String = (
                        'Aucun mouvement ne correspond aux critères que vous avez choisi',
                        'Vous devez renseigner un compte',
                        'Vous devez renseigner un compte existant'
                        );

Type
  TOF_CPCONTROLCAISSE = Class (TOF)
    Private
    ObjFiltre           : TObjFiltre;
    TobCaisse           : Tob;
    TitreDelaFiche      : String;
    fPopUp              : TPopUpMenu;
    procedure CoExoOnChange         (Sender : Tobject);
    function CreerLaRequete : String;
    procedure BRechercheOnClick     (Sender : Tobject);
    procedure RemplirGrille;
    procedure EffaceGrille;
    procedure LaFormOnKeyDown       (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure MajCaption            (LeTitre : String = '');
    procedure LeCompteOnChange      (Sender : Tobject);
    procedure LeCompteOnExit        (Sender : Tobject);
    procedure BGraphOnClick         (Sender : Tobject);
    procedure BImprimerOnClick      (Sender : Tobject);
    procedure BGrandLivreOnClick    (Sender: Tobject);
    procedure RedAgrandOnClick      (Sender : TObject);
    procedure BZoomGeneralOnClick   (Sender: Tobject);
    procedure OnPopupZoom           (Sender: Tobject);
    procedure OnChangeFFiltres      ;
    Public
    CoEtab              : THValComboBox;
    CoSens              : THValComboBox;
    CoExo               : THValComboBox;
    CoFiltre            : THValComboBox;
    LeCompte            : THEdit;
    LaDateDeb           : THEdit;
    LaDateFin           : THEdit;
    CbEcri              : THCheckbox;
    BRecherche          : TToolbarButton97;
    BValider            : TToolbarButton97;
    BFerme              : TToolbarButton97;
    BFiltre             : TToolbarButton97;
    BGraph              : TToolbarButton97;
    BImprimer           : TToolbarButton97;
    PopFiltre           : TPopupMenu;
    LaPage              : TPageControl;
    Fliste              : THGrid;

    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;


  end ;


procedure CPLanceFiche_ControlCaisse (Args : String);

Implementation

uses
    {$IFDEF MODENT1}
      CPProcGen,
      CPProcMetier,
      CPTypeCons,
    {$ENDIF MODENT1}
      Grids
    , TntStdCtrls
    , uLibExercice
    , HSysMenu
    , Vierge
    , uObjEtats
    , LookUp
    , uLibEcriture
    , CPGENERAUX_TOM;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /
Description .. : procédure le lancement de la fiche
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_ControlCaisse (Args : String);
begin
  AGLLanceFiche('CP', 'CPCONTROLCAISSE', '', '', Args);
end;


procedure TOF_CPCONTROLCAISSE.OnArgument (Arguments : String ) ;
var
  CompFiltre            : TControlFiltre;

begin
  Inherited ;
  TitreDelaFiche        := 'Contrôle de caisse';
  MajCaption;

  Ecran.HelpContext := 7631000;

  //Les controles
  CoEtab                := THValComboBox(GetControl('SETAB'));
  CoSens                := THValComboBox(GetControl('SSENS'));
  CoExo                 := THValComboBox(GetControl('SEXO'));
  CoFiltre              := THValComboBox(GetControl('FFILTRES'));
//  LeCompte              := THCptEdit(GetControl('SCPT'));
  LeCompte              := THEdit(GetControl('SCPT'));
  LaDateDeb             := THEdit(GetControl('SDATEC'));
  LaDateFin             := THEdit(GetControl('SDATEC_'));
  CbEcri                := THCheckbox(GetControl('SECRVALID'));
  PopFiltre             := TPopupMenu(GetControl('POPF'));
  BValider              := TToolbarButton97(GetControl('BVALIDER'));
  BFerme                := TToolbarButton97(GetControl('BFERME'));
  BRecherche            := TToolbarButton97(GetControl('BRECHERCHE'));
  BFiltre               := TToolbarButton97(GetControl('BFILTRE'));
  BGraph                := TToolbarButton97(GetControl('BGRAPH'));
  BImprimer             := TToolbarButton97(GetControl('BIMPRIMER'));
  LaPage                := TPageControl(GetControl('FE_PC1'));
  Fliste                := THGrid(GetControl('FLISTE'));
  fPopup                := TPopupMenu(GetCOntrol('POPZOOM'));
  fPopup.OnPopup        := OnPopupZoom;

  TToolbarButton97(GetControl('BAGRANDIR'       )).OnClick := RedAgrandOnClick;
  TToolbarButton97(GetControl('BREDUIRE'        )).OnClick := RedAgrandOnClick;

  //Valeurs par défaut
  //Si le Tag des différents critères n'est pas TAG=1 alors modification au niveau de la fiche AGL
  //la modif prend le dessus

  //appel tablette sens comptable
  if (Assigned(CoSens)) and (CoSens.Tag = 1) then
  begin
    CoSens.DataType     := 'TTRUBSIGNE';
    CoSens.Value        := 'POS';
  end;

  if assigned (CoExo) then CInitComboExercice(CoExo);

  //paramètres sur case à cocher des écritures valides
  if (Assigned(CbEcri)) and (CbEcri.Tag = 1) then
    CbEcri.State        := cbGrayed;

  //paramètres pour le champ compte général
  if (Assigned(LeCompte))  then
  begin
    LeCompte.DataType   := 'TZGENERAL';
    LeCompte.Plus := 'G_NATUREGENE = "CAI"';
    LeCompte.OnChange   := LeCompteOnChange;
    LeCompte.OnExit   := LeCompteOnExit;
  end;
  (*
  if Assigned(LeCompte) then
  begin
    LeCompte.MaxLength      := GetInfoCpta(fbGene).Lg;
    LeCompte.ElipsisButton  := True;
    LeCompte.OnElipsisClick := LeCompteElipsisClick;
  end;
*)

  //restriction sur les etablissements

  if Assigned(coEtab) then begin
    {Si l'on ne gère pas les établissement ...}
    if not VH^.EtablisCpta  then begin
      {... on affiche l'établissement par défaut}
      coEtab.Value := VH^.EtablisDefaut;
      {... on désactive la zone}
      coEtab.Enabled := False;
    end
    {On gère l'établisement, donc ...}
    else begin
      {... On commence par regarder les restrictions utilisateur}
      PositionneEtabUser(coEtab);
      {... s'il n'y a pas de restrictions, on reprend le paramSoc
       JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
      if coEtab.Value = '' then begin
        {... on affiche l'établissement par défaut
        coEtab.Value := VH^.EtablisDefaut;
        {... on active la zone
        coEtab.Enabled := True;
      end;}
    end;
  end;

  //Affectation pour les exercices
  if Assigned(CoExo) then
  begin
    if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
      CoExo.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
    else
     CoExo.Value      := CExerciceVersRelatif(VH^.Entree.Code) ;
    CoExo.OnChange      := CoExoOnChange;
    CExoRelatifToDates(CoExo.Value, LaDateDeb, LaDateFin);
  end;



  //paramètres pour le bouton de recherche
  if Assigned(BRecherche) then
  begin
    BRecherche.OnClick  := BRechercheOnClick;
  end;

  //paramètres pour le bouton de graphique
  if Assigned(BGraph) then
  begin
    BGraph.OnClick      := BGraphOnClick;
    Bgraph.Visible      := False;
  end;

  if Assigned(BImprimer) then
  begin
    BImprimer.OnClick      := BImprimerOnClick;
    BImprimer.Visible      := True;
  end;


  //les paramètres de la grille
  if Assigned(Fliste) then
  begin
    Fliste.RowCount     := 2;
    Fliste.ColCount     := 6;
    Fliste.FixedCols    := 1;
    Fliste.FixedRows    := 1;
    Fliste.DefaultRowHeight := 18;
    Fliste.MultiSelect  := False;
    Fliste.ScrollBars   := ssVertical;
    Fliste.SortEnabled  := False;
    Fliste.TitleBold    := True;
    Fliste.TitleCenter  := True;
    Fliste.Options      := Fliste.Options + [goRowSelect];
    Fliste.Cells[ClDate, 0]   := 'Date';
    Fliste.Cells[ClTotD, 0]   := 'Total débit';
    Fliste.Cells[CltotC, 0]   := 'Total crédit';
    Fliste.Cells[ClSolde, 0]  := 'Solde pour la date';
    Fliste.Cells[ClCumul, 0]  := 'Solde évolutif par date';
    Fliste.ColWidths[ClFixe]  := 10;
    Fliste.ColWidths[ClDate]  := 160;
    Fliste.ColWidths[ClTotD]  := 160;
    Fliste.ColWidths[CltotC]  := 160;
    Fliste.ColWidths[ClSolde] := 160;
    Fliste.ColWidths[ClCumul] := 160;

    Fliste.ColTypes[ClDate]   := 'D';
    Fliste.ColTypes[ClTotD]   := 'R';
    Fliste.ColFormats[ClTotD] := '##,##0.00';
    Fliste.ColAligns[ClTotD]  := taRightJustify;
    Fliste.ColTypes[ClTotC]   := 'R';
    Fliste.ColFormats[ClTotC] := '##,##0.00';
    Fliste.ColAligns[ClTotC]  := taRightJustify;
    Fliste.ColTypes[ClSolde]  := 'R';
    Fliste.ColFormats[ClSolde]:= '##,##0.00';
    Fliste.ColAligns[ClSolde] := taRightJustify;
    Fliste.ColTypes[ClCumul]  := 'R';
    Fliste.ColFormats[ClCumul]:= '##,##0.00';
    Fliste.ColAligns[ClCumul] := taRightJustify;

  end;

  //Gestion des filtres
  CompFiltre.PopupF     := PopFiltre;
  CompFiltre.Filtres    := CoFiltre;
  CompFiltre.Filtre     := BFiltre;
  CompFiltre.PageCtrl   := LaPage;
  ObjFiltre             := TObjFiltre.Create(CompFiltre, 'CPCONTROLCAISSE');
  ObjFiltre.Charger;
  objFiltre.ApresChangementFiltre := OnChangeFFiltres;

  TobCaisse             := Tob.Create('Le control de caisse', nil, -1);
  EffaceGrille;

  TForm(Ecran).OnKeyDown  := LaFormOnKeyDown;
  THSystemMenu(TFVierge(ECRAN).HMTrad).ResizeGridColumns(fListe);
end ;

procedure TOF_CPCONTROLCAISSE.OnClose ;
begin
  Inherited ;
  ObjFiltre.Free;
  TobCaisse.Free;

end ;

procedure TOF_CPCONTROLCAISSE.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /    
Description .. : Changement d'exercice
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLCAISSE.CoExoOnChange(Sender: Tobject);
begin
//  ExoToDates(CoExo.Value, LaDateDeb, LaDateFin);
  if CoExo.ItemIndex <= 0 then
  begin
    if VH^.Exov8.Code <> '' then
      LaDateDeb.Text := DateToStr(VH^.Exov8.Deb)
    else
      // Date de début du premier Exercice
      LaDateDeb.Text := DateToStr(iDate1900);
    LaDateFin.Text  := DateToStr(iDate2099-1);
  end
  else CExoRelatifToDates ( CoExo.Value, LaDateDeb, LaDateFin);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /
Description .. : Creation de la requete pour la recherche
Mots clefs ... : 
*****************************************************************}
function TOF_CPCONTROLCAISSE.CreerLaRequete: String;
var
  LaReq                 : String;

begin
  LaReq                 := 'SELECT E_DATECOMPTABLE AS DAT, SUM(E_DEBIT) AS DEB, SUM(E_CREDIT) AS CRE, ' +
                           'SUM(E_DEBIT-E_CREDIT) AS SOL FROM ECRITURE ' +
//                           'WHERE E_GENERAL="' + LeCompte.Text + '" ' +
                           'WHERE E_GENERAL="' + LeCompte.Text + '" ' +
                           ' AND E_QUALIFPIECE="N" '+
                           'AND E_DATECOMPTABLE >="' + USDATETIME(strtodate(LaDateDeb.Text)) + '" '+
                           'AND E_DATECOMPTABLE <="' + USDATETIME(strtodate(LaDateFin.Text)) + '" ';

  // Si exercice=tous, on ne prend pas les écritures d'à-nouveau
  if CoExo.Value = '' then LaReq := LaReq+' AND E_ECRANOUVEAU="N" ';

  case CbEcri.State of
    cbUnchecked :
      begin
        LaReq           := LaReq + 'AND E_VALIDE="-" ';
      end;
    cbChecked :
      begin
        LaReq           := LaReq + 'AND E_VALIDE="X" ';
      end;
  end;

  if CoExo.ItemIndex > 0 then
    LaReq               := LaReq + 'AND E_EXERCICE="' + CRelatifVersExercice ( CoExo.Value ) + '" ';

  if CoEtab.ItemIndex > 0 then
    LaReq               := LaReq + 'AND E_ETABLISSEMENT="' + CoEtab.Value + '" ';

  LaReq                 := LaReq + 'GROUP BY E_DATECOMPTABLE ORDER BY E_DATECOMPTABLE DESC';

  Result                := LaReq;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /
Description .. : Mise à jour des critères de recherche et lancement de la 
Suite ........ : requete
Suite ........ : affichage des données
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLCAISSE.BRechercheOnClick(Sender: Tobject);
var
  StrSql                : String;

begin
  //si le compte est vide
//  if LeCompte.Text = '' then
  if LeCompte.Text = '' then
  begin
    PGIError(StrErrorCpt[1], TForm(Ecran).Caption);
    Exit;
  end;

  //Creation de la requete
  StrSql                := CreerLaRequete;

  //Chargement des données
  if not TobCaisse.LoadDetailFromSQL(StrSql) then
  begin
    PGIError(StrErrorCpt[0], TForm(Ecran).Caption);
    EffaceGrille;
    BGraph.Visible      := False;
    Exit;
  end;

  RemplirGrille;
  BGraph.Visible        := True;

  THSystemMenu(TFVierge(ECRAN).HMTrad).ResizeGridColumns(fListe);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /    
Description .. : procedure pour remplir la grille à partir de la TobCaisse
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPCONTROLCAISSE.RemplirGrille;
var
  i                     : integer;
  MtDeb                 : Extended;
  MtCre                 : Extended;
  MtSol                 : Extended;
  MtCum                 : Extended;
  TobC                  : Tob;

begin
  //D'abord on vide l'ancienne grille
  EffaceGrille;

  //on affecte les données de la TobCaisse
  Fliste.RowCount       := TobCaisse.Detail.Count +1;

  for i := 0 to TobCaisse.Detail.Count -1 do
  begin
    TobC                := TobCaisse.Detail[i];
    MtDeb               := TobC.GetValue('DEB');
    MtCre               := TobC.GetValue('CRE');
    MtSol               := TobC.GetValue('SOL');
    Fliste.Cells[ClDate, i+1]   := Tobc.GetValue('DAT');
    Fliste.Cells[ClTotD, i+1]   := STRFMONTANT(MtDeb, 15, V_PGI.OkDecV, '', True);
    Fliste.Cells[ClTotC, i+1]   := STRFMONTANT(MtCre, 15, V_PGI.OkDecV, '', True);
    if CoSens.Value = 'POS' then
      Fliste.Cells[ClSolde, i+1]  := STRFMONTANT(MtSol, 15, V_PGI.OkDecV, '', True)
    else
      Fliste.Cells[ClSolde, i+1]  := STRFMONTANT(MtSol*(-1), 15, V_PGI.OkDecV, '', True);
  end;

  { Calcul du solde évolutif par date }
  mtCum := 0;
  for i := TobCaisse.Detail.Count - 1 downto 0 do
  begin
    TobC := TobCaisse.Detail[i];
    if TobC <> nil then
    begin
      MtSol := TobC.GetValue('SOL');
      MtCum := MtCum + MtSol;
      if MtCum < 0 then
        Fliste.Cells[ClCumul, i+1]  := PrintSolde(0, -MtCum, V_PGI.OkDecV, '', False)
      else
        Fliste.Cells[ClCumul, i+1]  := PrintSolde(MtCum, 0, V_PGI.OkDecV, '', False);
    end;
  end;
end;

procedure TOF_CPCONTROLCAISSE.EffaceGrille;
begin
  Fliste.RowCount           := 2;
  Fliste.Cells[ClDate, 1]   := '';
  Fliste.Cells[ClTotD, 1]   := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
  Fliste.Cells[ClTotC, 1]   := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
  Fliste.Cells[ClSolde, 1]  := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
  Fliste.Cells[ClCumul, 1]  := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/02/2007
Modifié le ... :   /  /    
Description .. : gestion des touches de la fenetre
Mots clefs ... :
*****************************************************************}
procedure TOF_CPCONTROLCAISSE.LaFormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of
    VK_F9 :
      begin
        BRechercheOnClick(nil);
      end;
  end;

end;

procedure TOF_CPCONTROLCAISSE.MajCaption(LeTitre: String);
var
  StrTitre              : String;

begin
  StrTitre              := TitreDelaFiche;
  if LeTitre <> '' then
    StrTitre            := StrTitre + ' - ' + LeTitre;

  TForm(Ecran).Caption  := StrTitre;
  UpdateCaption(TForm(Ecran));

end;

procedure TOF_CPCONTROLCAISSE.LeCompteOnChange(Sender: Tobject);
var
  StrCap                : String;

begin
  StrCap                := LeCompte.Text;
  MajCaption(StrCap);
end;

procedure TOF_CPCONTROLCAISSE.BGraphOnClick(Sender: Tobject);
var
  LesDates              : TStrings;
  i                     : integer;

begin
  LesDates              := TStringList.Create;
  LesDates.Add('');

  for i := 0 to TobCaisse.Detail.Count -1 do
    LesDates.Add(TobCaisse.Detail[i].GetValue('DAT'));

  VisuGraph('CTRLCAI', TForm(Ecran).Caption, Fliste, LesDates, 0, 1, Fliste.RowCount -1, True, [ClCumul], nil, nil);

end;

procedure TOF_CPCONTROLCAISSE.BImprimerOnClick(Sender: Tobject);
begin
  TObjEtats.GenereEtatGrille (fListe,ECRAN.Caption,False);
end;

procedure TOF_CPCONTROLCAISSE.LeCompteOnExit(Sender: Tobject);
var
  stCompte : string;
begin
  stCompte := BourreEtLess(LeCompte.Text,fbGene);
  if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+stCompte+'" AND G_NATUREGENE="CAI"') then
    LeCompte.Text := stCompte
  else LookupList(LeCompte,TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE','G_NATUREGENE="CAI" AND ' + CGenereSQLConfidentiel('G'),'',true, 1) ;
end;

procedure TOF_CPCONTROLCAISSE.BZoomGeneralOnClick(Sender: Tobject);
begin
  if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+LeCompte.text+'" AND G_NATUREGENE="CAI"')
  then FicheGene(Nil,'',LeCompte.Text,taConsult,0)
  else
  begin
    PGIInfo (TraduireMemoire('Veuillez choisir un compte de caisse correct.'));
    exit;
  end;
end;

procedure TOF_CPCONTROLCAISSE.BGrandLivreOnClick(Sender: Tobject);
Var
  ACritEdt              : ClassCritEdt;
  D1                    : TDateTime;
  D2                    : TDateTime;
  Etab                  : String;
begin
  if not ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+LeCompte.text+'" AND G_NATUREGENE="CAI"')
  then
  begin
    PGIInfo (TraduireMemoire('Veuillez choisir un compte de caisse correct.'));
    exit;
  end;
  ACritEdt              := ClassCritEdt.Create;
  try
    Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
    D1                        := StrToDate(LaDateDeb.text);
    D2                        := StrToDate(LaDateFin.text) ;
    ACritEdt.CritEdt.Date1    := D1;
    ACritEdt.CritEdt.DateDeb  := ACritEdt.CritEdt.Date1;
    ACritEdt.CritEdt.Date2    := D2;
    ACritEdt.CritEdt.DateDeb  := ACritEdt.CritEdt.Date2;
    ACritEdt.CritEdt.Cpt1     := LeCompte.Text;
    ACritEdt.CritEdt.Cpt2     := LeCompte.Text;
    ACritEdt.CritEdt.SCpt1    := '';
    ACritEdt.CritEdt.SCpt2    := '';
    Etab                      := EtabForce;
    if Etab <> '' then
      ACritEdt.CritEdt.Etab   := Etab;

    TheData                   := ACritEdt;
    CPLanceFiche_CPGLGENE();

  Finally
    ACritEdt.Free;
  end;
end;

procedure TOF_CPCONTROLCAISSE.RedAgrandOnClick(Sender: TObject);
begin
  if UpperCase((Sender as TToolbarButton97).Name) = 'BREDUIRE' then begin
    SetControlVisible('FE_PC1', True);
    SetControlVisible('BREDUIRE'   , False);
    SetControlVisible('BAGRANDIR'  , True);
  end
  else begin
    SetControlVisible('FE_PC1', False);
    SetControlVisible('BREDUIRE'   , True);
    SetControlVisible('BAGRANDIR'  , False);
  end;
end;

procedure TOF_CPCONTROLCAISSE.OnPopupZoom(Sender: Tobject);
begin
  // Bouton ZOOM
  TMenuItem(fPopUp.Items[0]).OnClick := BGrandLivreOnClick;
  TMenuItem(fPopUp.Items[1]).OnClick := BZoomGeneralOnClick;
end;

procedure TOF_CPCONTROLCAISSE.OnChangeFFiltres;
begin
  CExoRelatifToDates(CoExo.Value, LaDateDeb, LaDateFin, True);
  if CoExo.Value = '' then
  begin
    CoExo.ItemIndex := 0;
    CoExoOnChange (nil);
  end;
end;

Initialization
  registerclasses ( [ TOF_CPCONTROLCAISSE ] ) ;
end.




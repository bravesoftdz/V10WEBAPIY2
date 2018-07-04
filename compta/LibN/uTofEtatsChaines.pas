{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 08/11/2001
Modifié le ... : 12/10/2006
Description .. : Source TOF de la TABLE : ETATSCHAINES ()
Suite ........ : 24/06/2003 - CA - Ajout d'un nouveau type d'état : AM pour
Suite ........ : les amortissements
Mots clefs ... : TOF;ETATSCHAINES
*****************************************************************}
unit uTOFEtatsChaines;

interface

uses
  StdCtrls,
  Controls,
  Classes,
  forms,
  sysutils,
{$IFDEF EAGLCLIENT}
  MainEAgl,
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main,
{$ENDIF}
  Menus,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Htb97,
  uTob,
  Grids,
  Graphics,
  Ent1,
  Vierge,
  ShellAPI,
  Filtre,
  Dialogs,     // TOpenDialog, TPrintDialog
  Windows,     // VK_
  LookUp,      // LookUpList
  Messages,    // VK_TAB
  HSysMenu,    // THSystemMenu
  uDossierSelect,   // MD 27/04/04 : voir dans DP\Lib
  uTobDebug,        // TobDebug
  uLibWindows,      // IIF( , GetWindowsTempPath, FileExistsDelete
  uListByUser,      // TListByUser
  uTofLookUpTob,    // LookUpTob
  uLibExercice,     // CInitComboExercice(E_EXERCICE);
  uObjFiltres,      // Gestion des Filtres AGL
  uObjEtatsChaines, // lCEtatsChaines
{$IFDEF PAIEGRH}
  PgOutils2,        // Etats
{$ENDIF}
  uTXML;            // XMLDecodeSt

type

  TOF_ETATSCHAINES = class(TOF)

    ComboNature      : THValComboBox;
    ComboMode        : THValComboBox;
    ComboExercice    : THValComboBox;
    ComboType        : THMultiValComboBox;
    DateCDu          : THEdit;
    DateCAu          : THEdit;
    BValider         : TToolBarButton97;

    FCBCritere       : TCheckbox;
    FCPGBCritere     : TGroupBox;
    FPGGBCritere     : TGroupBox;

    PGDateDeb        : THEdit;
    PGDateFin        : THEdit;
// critères paie pour DUCS
    PGNatDeb         : THValComboBox;
    PGNatFin         : THValComboBox;
    PGEtabDeb        : THValComboBox;
    PGEtabFin        : THValComboBox;
//
    FCBAuFormatPdf   : TCheckBox;
    FGBPDF           : TGroupBox;
    FNomPdf          : THEdit;
    BNomPdf          : TToolBarButton97;
    FCBMultiPdf      : TCheckBox;

    FListe           : THGrid;
    FGrille          : THGrid;
    FPAGES           : TPageControl;

    procedure OnArgument(S: string);override;
    procedure OnLoad;               override;
    procedure OnClose;              override;
    procedure OnClickBValider       (Sender: TObject);
    procedure OnKeyDownEcran        (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClickBGauche        (Sender: TObject);
    procedure OnClickBDroit         (Sender: TObject);
    procedure OnClickBHaut          (Sender: TObject);
    procedure OnClickBBas           (Sender: TObject);
    procedure OnClickBEfface        (Sender: TObject);
    procedure OnClickFCBCritere     (Sender: TObject);
    procedure OnClickFCBAuFormatPdf (Sender: TObject);
    procedure OnChangeComboExercice (Sender: TObject);
    procedure OnDblClickFListe      (Sender: TObject);

    // Evénements des états sélectionnés
    procedure OnRowExitFGrille      (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure OnKeyPressFGrille     (Sender: TObject; var Key: Char);
    procedure OnCellEnterFGrille    (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnCellExitFGrille     (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure OnElipsisClickFGrille (Sender: TObject);

    // Evénements pour ouverture fichier .PDF
    procedure OnClickBtnNomPdf      (Sender: TObject);
    procedure OnElipsisClickFNomPdf (Sender: TObject);
    procedure OnExitFNomPdf         (Sender: TObject);

  private
    FObjFiltre       : TObjFiltre;     // Gestion des filtres par l'AGL
    FControlFiltre   : TControlFiltre;

    FTobYPAramEdition : Tob;         // Liste des états disponibles dans YPARAMEDITION
    FTobSelect        : Tob;         // TOB des états à imprimer

    //{IFDEF DP}
    FStFichierXML    : string;
    //{$ENDIF}

    // Variables de contexte
    FStType             : string ; // Type d'applications CP ou PG
    FFI_Table           : string ; // Valeur du champ Nom Filtre de la table Filtre (FI_TABLE)

    procedure InitFListe;
    procedure InitFGrille;
    procedure InitContexteEcran;
    {$IFDEF DP}
    procedure EnregistreEnXML;
    {$ENDIF}

    procedure MiseAJourFGrille;
    procedure LoadLookUpList(lOu: integer);
    procedure EnregistrementLigne(lRow: integer);
    procedure LoadTobFromGrille;

    procedure AfterShow;
    procedure AvantChangementFiltre;
    procedure ApresChangementFiltre;
    procedure SupprimeFiltre;
  end;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_EtatsChaines( vStArg : string ) : string;
////////////////////////////////////////////////////////////////////////////////

implementation

uses CbpPath; // TcbpPath.GetCegidUserTempPath

const
  cColFListeLibelle = 0;
  cColFListeID      = 1;

  cColID            = 0  ; // ID de l' état ( Clé unique dans le tableau )
  cColLibelleEtat   = 1  ; // Libelle de l'état
  cColNbExemplaire  = 2  ; // Nombre d'exemplaires pour les états
  cColFiltreUtilise = 3  ; // Filtre à utiliser dans les états
  cColCodeEtat      = 4  ; // Plus utilise depuis v 4.20.800.2 mais doit rester
  cColNatureEtat    = 5  ; // Plus utilise depuis v 4.20.800.2 mais doit rester
  cNbCol            = 6  ; // Nombre de colonne dans la grille de droite

  cCPTProduit       = 'COM';
  cPGTProduit       = 'SOC';
  cAMTProduit       = 'IMM';

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_EtatsChaines( vStArg : string ) : string;
begin
  Result := AGLLanceFiche('CP', 'ETATSCHAINES', '', '', vStArg);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/02/2003
Modifié le ... : 05/02/2003
Description .. : Description de S = "NOMDEFILTRE;MODE"
Suite ........ : NOMDEFILTRE = contient le nom du filtre à charger au démarrage
Suite ........ : MODE = CP pour la COMPTA ou PG pour la PaieGRH
Mots clefs ... : ETATSCHAINES
*****************************************************************}
procedure TOF_ETATSCHAINES.OnArgument(S: string);
var lStParam : string;
begin
  inherited;
  lStParam := S;
  // Contexte d'utilisation de l'écran, COMPTA ou PAIE ou AMORTISSEMENT
  FStType           := ReadTokenSt( lStParam ); // CP ou PG ou AM
  if FStType = 'CP' then FFI_Table := 'ETATSCHAINES';
  if FStType = 'PG' then FFI_Table := 'PGETATSCHAINES';
  if FStType = 'AM' then FFI_Table := 'AMETATSCHAINES';
  
  Ecran.Caption := TraduireMemoire('Etats chaînés');
  UpDateCaption(Ecran);
  Ecran.OnKeyDown := OnKeyDownEcran;

  FPages         := TPageControl(GetControl('PAGES', True));
  FListe         := THGrid(GetControl('FLISTE'));
  FGrille        := THGrid(GetControl('FGRILLE', True));

  FCBCritere     := TCheckBox(GetControl('FCBCRITERE'));
  FCPGBCritere   := TGroupBox(GetControl('FCPGBCRITERE', True));
  FPGGBCritere   := TGroupBox(GetControl('FPGGBCRITERE', True));

  // Critères d'impression PDF
  FGBPdf         := TGroupBox(GetControl('FGBPDF'));
  FCBAuFormatPdf := TCheckBox(GetControl('FCBAuFormatPdf'));
  FCBMultiPdf    := TCheckBox(GetControl('FCBMultiPdf'));
  BNomPdf        := TToolBarButton97(GetControl('BNOMPDF'));
  FNomPdf        := THEdit(GetControl('FNOMPDF'));

  // Critères Compta
  ComboNature    := THValComboBox(GetControl('ComboNature', True));
  ComboMode      := THValComboBox(GetControl('ComboMode', True));
  ComboExercice  := THValComboBox(GetControl('ComboExercice', True));
  CInitComboExercice( ComboExercice ); // Init de la combo Exercice en Relatif
  DateCDu        := THEdit(GetControl('DateCDu', True));
  DateCAu        := THEdit(GetControl('DateCAu', True));
  ComboType      := THMultiValComboBox(GetControl('COMBOTYPE', True));

  // Critères Paie
  PGDateDeb      := THEdit(GetControl('PGDATEDEB', True));
  PGDateFin      := THEdit(GetControl('PGDATEFIN', True));
  // critères spécifiques DUCS
  PGNatDeb       := THValComboBox(GetControl('PGNATDEB', True));
  PGNatFin       := THValComboBox(GetControl('PGNATFIN', True));
  PGEtabDeb      := THValComboBox(GetControl('PGETABDEB', True));
  PGEtabFin      := THValComboBox(GetControl('PGETABFIN', True));
  //

  FCBAuFormatPdf.OnClick := OnClickFCBAuFormatPdf;
  FNomPdf.OnElipsisClick := OnElipsisClickFNomPdf;
  FNomPdf.OnExit         := OnExitFNomPdf;

  FCBCritere.OnClick := OnClickFCBCritere;

  ComboExercice.OnChange := OnChangeComboExercice;
(*
{$IFDEF DP}
  // GCO - 29/01/2003 Activation des paramètres standards pour la PAIE
  FCBCritere.Enabled := (FstType = 'PG');
{$ELSE}
  FCBMultiPDF.Enabled := False;
  {$IFDEF COMPTA}
  // Initialisation de la ComboExercice avec l'exercice de référence ou sur l' exercice en cours
  ComboExercice.Value := CExerciceVersRelatif( IIF(VH^.CPExoRef.Code <> '', VH^.CPExoRef.Code, QUELEXODTBUD(VH^.Encours.Deb)));
  {$ENDIF}
{$ENDIF}*)
  
  TToolBarButton97(GetControl('BGAUCHE')).OnClick := OnClickBGauche;
  TToolBarButton97(GetControl('BDROIT')).OnClick  := OnClickBDroit;
  TToolBarButton97(GetControl('BHAUT')).OnClick   := OnClickBHaut;
  TToolBarButton97(GetControl('BBAS')).OnClick    := OnClickBBas;
  TToolBarButton97(GetControl('BEFFACE')).OnClick := OnClickBEfface;

  FGrille.OnRowExit := OnRowExitFGrille;
  FGrille.OnKeyPress := OnKeyPressFGrille;
  FGrille.OnCellEnter := OnCellEnterFGrille;
  FGrille.OnCellExit := OnCellExitFGrille;
  FGrille.OnElipsisClick := OnElipsisClickFGrille;

  BValider := TToolBarButton97(GetControl('BVALIDER', True));
  BValider.OnClick := OnClickBValider;
  TToolBarButton97(GetControl('BNOMPDF')).OnClick := OnClickBtnNomPdf;

  FTobYParamEdition := Tob.Create('', nil, -1);
  FTobSelect := Tob.Create('ETATSCHAINES', nil, -1);

  // Nouvelle Gestion des Filtres
  FControlFiltre.Filtre   := TToolBarButton97(GetControl('BFILTRE', True));
  FControlFiltre.Filtres  := THValComboBox(GetControl('FFILTRES'));
  FControlFiltre.PageCtrl := FPages;
  FObjFiltre := TObjFiltre.Create(FControlFiltre, FFI_TABLE, False);

  FObjFiltre.AvantChangementFiltre := AvantChangementFiltre;
  FObjFiltre.ApresChangementFiltre := ApresChangementFiltre;
  FObjFiltre.SupprimeFiltre        := SupprimeFiltre;
  TFVierge(Ecran).OnAfterFormShow := AfterShow;

  InitContexteEcran;
  InitFListe;        // Chargement de la liste des états disponibles
  InitFGrille;       // Initialisation de la grille des états sélectionnés
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... : 11/10/2006
Description .. : Chargement de l'affichage de l'écran
Suite ........ : Initialisation des critères si valeur passé en paramètre.
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnLoad;
begin
  inherited;
  // Chargement des Filtre spar L'AGL, on peut travailler par dessus
  FObjFiltre.Charger;

  FCBAuFormatPdf.OnClick( nil );
  FCBCritere.OnClick( nil );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/07/2004
Modifié le ... : 11/10/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.AfterShow;
begin
{$IFDEF CCSTD}
  FObjFIltre.ForceAccessibilite := faPublic;

{$ENDIF}
  if FListe.CanFocus then FListe.SetFocus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.AvantChangementFiltre;
begin
  FGrille.VidePile(False);
  FGrille.RowCount := 100;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.ApresChangementFiltre;
begin
  LoadTobFromGrille;
  MiseAJourFGrille;
{$IFDEF COMPTA}
  CExoRelatifToDates( ComboExercice.Value, DateCDu, DateCAu,True); // FQ 17312
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.SupprimeFiltre;
begin
  FTobSelect.ClearDetail;
  MiseAJourFGrille;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.InitContexteEcran;
{$IFDEF PAIEGRH}
var ExerPerEncours : string;
    DebPer         : string;
    FinPer         : string;
{$ENDIF}
begin
  // Comptabilité
  if FStType = 'CP' then
  begin
    FPGGBCritere.Visible := False;
  end;

  // Paie
  if FStType = 'PG' then
  begin
    FCPGBCritere.Visible := False;
  end;

  // Amortissement
  if FStType = 'AM' then
  begin
    FCBCritere.Visible := False;
    FCPGBCritere.Visible := False;
    FPGGBCritere.Visible := False;
  end;

  FStFichierXML := GetWindowsTempPath + FStTYpe + '_ETATSCHAINES_' + V_PGI.USER + '.XML';

{$IFDEF DP}
  // GCO - 29/01/2003 Activation des paramètres standards pour la PAIE
  FCBCritere.Enabled := (FstType = 'PG');
  FStFichierXML := GetWindowsTempPath + FStTYpe + '_ETATSCHAINES_' + V_PGI.USER + '.XML';
{$ELSE}
  FCBMultiPDF.Enabled := False;
{$ENDIF}

{$IFDEF COMPTA}
  ComboNature.ItemIndex := 0;
  ComboMode.ItemIndex   := 0;
  ComboType.Value       := TraduireMemoire('<<Tous>>');
  ComboType.SelectAll;

  // Initialisation de la ComboExercice avec l'exercice de référence ou sur l' exercice en cours
  ComboExercice.Value := CExerciceVersRelatif( IIF(VH^.CPExoRef.Code <> '', VH^.CPExoRef.Code, QUELEXODTBUD(VH^.Encours.Deb)));
{$ENDIF}

{$IFDEF PAIEGRH}
  if PgOutils2.RendPeriodeEnCours (ExerPerEncours, DebPer, FinPer) = True then
  begin
    PGDateDeb.Text := DebPer;
    PGDateFin.Text := FinPer;
  end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/10/2006
Modifié le ... : __/__/____
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF DP}
procedure TOF_ETATSCHAINES.EnregistreEnXML;
begin
  // Utilisation du Format PDF pour l'impression
  FTobSelect.AddChampSupValeur('AUFORMATPDF', False);
  FTobSelect.AddChampSupValeur('NOMPDF', '');
  FTobSelect.AddChampSupValeur('MULTIPDF', False);
  FTobSelect.AddChampSupValeur('CRITERE', False);
  FTobSelect.AddChampSupValeur('PGDATEDEB', '');
  FTobSelect.AddChampSupValeur('PGDATEFIN', '');
// critères paie spécifiques DUCS
  FTobSelect.AddChampSupValeur('PGNATDEB', '');
  FTobSelect.AddChampSupValeur('PGNATFIN', '');
  FTobSelect.AddChampSupValeur('PGETABDEB', '');
  FTobSelect.AddChampSupValeur('PGETABFIN', '');
//
  if FCBAuFormatPDF.Checked then
  begin
    FTobSelect.SetBoolean('AUFORMATPDF', True);
    FTobSelect.SetString('NOMPDF', FNomPdf.Text);
    if FCBMultiPdf.Checked then
      FTobSelect.SetBoolean('MULTIPDF', True);
  end;

  // Utilisation des critères standards
  if FCBCritere.Checked then
  begin
    FTobSelect.SetBoolean('CRITERE', True);
    //Paie
    if IsValidDate(PGDateDeb.Text) then
      FTobSelect.SetString('PGDATEDEB', PGDateDeb.Text);

    if IsValidDate(PGDateFin.Text) then
      FTobSelect.SetString('PGDATEFIN', PGDateFin.Text);
// critères paie spécifiques DUCS
      FTobSelect.SetString('PGNATFIN', PGDateFin.Text);
      FTobSelect.SetString('PGNATFIN', PGDateFin.Text);
      FTobSelect.SetString('PGETABDEB', PGEtabDeb.Text);
      FTobSelect.SetString('PGETABFIN', PGEtabFin.Text);
//
  end;

  // Export de la TOB dans le fichier XML
  FTobSelect.SaveToXmlFile(FStFichierXML, True);

end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. : Chargement de la liste des états disponibles
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.InitFListe;
var i : integer;
{$IFDEF YPARAMEDITION}
    lStSql : string;
{$ENDIF}
begin
  FListe.FixedRows := 1;
  
  FListe.OnDblClick := OnDblClickFListe;
  FListe.ColWidths[cColFListeLibelle] := FListe.Width;
  FListe.ColWidths[cColFListeID] := -1;
{$IFDEF YPARAMEDITION}
  if V_Pgi.SAV then
  begin
    FListe.ColWidths[cColFListeID]  := 100;
  end;
{$ENDIF}
  FListe.ColLengths[cColFListeID]     := -1;

  FListe.Cells[cColFListeLibelle,0]   := TraduireMemoire('Liste des états');
  FListe.ColAligns[cColFListeLibelle] := TaLeftJustify;
  FListe.RowCount := 2;

{$IFDEF YPARAMEDITION}
  lStSql := 'SELECT YED_TYPEETAT, YED_NATUREETAT, YED_CODEETAT, YED_LIBELLE, ' +
            'YED_CODEID, YED_TYPEFORME, YED_FORME, YED_NOMFILTRE ' +
            'FROM YPARAMEDITION WHERE ' +
            'YED_ETATCHAINE = "X"';

  if FStType = 'CP' then
    lStSql := lStSql + ' AND YED_TPRODUIT = "' + cCPTProduit + '"'
  else
  if FStType = 'PG' then
    lStSql := lStSql + ' AND YED_TPRODUIT = "' + cPGTProduit + '"'
  else
  if FStType = 'AM' then
    lStSql := lStSql + ' AND YED_TPRODUIT = "' + cAMtProduit + '"';

  lStSql := lStSql + ' ORDER BY YED_CODEID';

  FTobYParamEdition.LoadDetailFromSQL( lStSql );

  for i := 0 to FTobYParamEdition.Detail.Count-1 do
  begin
    FListe.Cells[cColFListeLibelle, FListe.RowCount-1] := FTobYParamEdition.Detail[i].Getstring('YED_LIBELLE');
    FListe.Cells[cColFListeID, FListe.RowCount-1]      := FTobYParamEdition.Detail[i].Getstring('YED_CODEID');
    FListe.RowCount := FListe.RowCount + 1 ;
  end;
{$ELSE}
  for i := 0 to cNbEtat do
  begin
    if FStType = cListeEtat[i].stType then
    begin
      FListe.Cells[cColFListeLibelle, FListe.RowCount-1] := cListeEtat[i].stLibelleEtat;
      FListe.Cells[cColFListeID, FListe.RowCount-1]      := cListeEtat[i].stID;
      FListe.RowCount := FListe.RowCount + 1 ;
    end;
  end;
{$ENDIF}

  FListe.RowCount := FListe.RowCount - 1 ;
  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FListe);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.InitFGrille;
begin
  FGrille.ColCount  := cNbCol;
  FGrille.Cells[cColID, 0]     := TraduireMemoire('ID Etat');
  FGrille.ColWidths[cColID]    := -1 ;
{$IFDEF YPARAMEDITION}
  if V_Pgi.Sav then FGrille.ColWidths[cColID]    := 50 ;
{$ENDIF}
  FGrille.ColLengths[cColID]   := FGrille.ColWidths[cColID];
  FGrille.ColEditables[cColID] := False;

  FGrille.Cells[cColLibelleEtat, 0] := TraduireMemoire('Libellé');
  FGrille.ColAligns[cColLibelleEtat] := TaLeftJustify;
  FGrille.ColWidths[cColLibelleEtat] := 175;
  FGrille.ColLengths[cColLibelleEtat] := -1;
  FGrille.ColEditables[cColLibelleEtat] := False;

  FGrille.Cells[cColNbExemplaire, 0] := TraduireMemoire('Nombre');
  FGrille.ColFormats[cColNbExemplaire] := 'F';
  FGrille.ColAligns[cColNbExemplaire] := TaLeftJustify;
  FGrille.ColWidths[cColNbExemplaire] := 45;

  FGrille.Cells[cColFiltreUtilise, 0] := TraduireMemoire('Filtre à utiliser');
  FGrille.ColAligns[cColFiltreUtilise] := TaLeftJustify;
  FGrille.ColWidths[cColFiltreUtilise] := 135;

  FGrille.Cells[cColCodeEtat, 0] := TraduireMemoire('Code Etat');
  FGrille.ColWidths[cColCodeEtat]  := -1 ;
  FGrille.ColLengths[cColCodeEtat] := -1 ;
{$IFDEF YPARAMEDITION}
  if V_Pgi.Sav then
  begin
    FGrille.ColWidths[cColCodeEtat]  := 50 ;
    FGrille.ColLengths[cColCodeEtat] := 50 ;
  end;  
{$ENDIF}
  FGrille.ColEditables[cColCodeEtat] := False;

  FGrille.Cells[cColNatureEtat, 0] := TraduireMemoire('Nature Etat');
  FGrille.ColWidths[cColNatureEtat]  := -1 ;
  FGrille.ColLengths[cColNatureEtat] := -1 ;

{$IFDEF YPARAMEDITION}
  if V_Pgi.Sav then
  begin
    FGrille.ColWidths[cColNatureEtat]  := 50 ;
    FGrille.ColLengths[cColNatureEtat] := 50 ;
  end;  
{$ENDIF}
  FGrille.ColEditables[cColNatureEtat] := False;

  FGrille.VidePile(False);

  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FGrille);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... :   /  /
Description .. : Changement dans la combo de l'exercice
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnChangeComboExercice(Sender: TObject);
begin
{$IFDEF COMPTA}
  CExoRelatifToDates( ComboExercice.Value, DateCDu, DateCAu);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. : Ajout d'un état dans la liste à imprimer
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBGauche(Sender: TObject);
var lTobTemp: Tob;
begin
  if FTobSelect.Detail.Count = 0 then Exit;

  //FGrille.Col := cColNombre;
  if FGrille.ValCombo <> nil then
    FGrille.ValCombo.Visible := False;

  lTobTemp := Tob(FGrille.Objects[0, FGrille.Row]);
  lTobTemp.Free;
  MiseAJourFGrille;
  FGrille.SetFocus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. : Suppression d'un état dans la liste à imprimer
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBDroit(Sender: TObject);
var lTobTemp: Tob;
    i : integer ;
{$IFDEf YPARAMEDITION}
    lTobEtatAjouter : Tob;
{$ENDIF}
begin
{$IFDEF YPARAMEDITION}
  lTobEtatAjouter := FTobYParamEdition.FindFirst(['YED_CODEID'],[FListe.Cells[cColFListeID, FListe.Row]], False);
  if lTobEtatAJouter <> nil then
  begin
    lTobTemp := Tob.Create('', FTobSelect, -1);
    lTobTemp.AddChampSupValeur('ID', lTobEtatAJouter.GetString('YED_CODEID'), False);
    lTobTemp.AddChampSupValeur('LIBELLEETAT', lTobEtatAJouter.Getstring('YED_LIBELLE'), False);
    lTobTemp.AddChampSupValeur('NBEXEMPLAIRE', 1, False);
    lTobTemp.AddChampSupValeur('FILTREUTILISE', '', False);
  end;

{$ELSE}
  for i:= 0 to cNbEtat do
  begin
    // Identification  de l"état à ajouter avec son ID
    if cListeEtat[i].stID = FListe.Cells[cColFListeID, FListe.Row] then
    begin
      lTobTemp := Tob.Create('', FTobSelect, -1);
      lTobTemp.AddChampSupValeur('ID', cListeEtat[i].stID, False);
      lTobTemp.AddChampSupValeur('LIBELLEETAT', cListeEtat[i].stLibelleEtat, False);
      lTobTemp.AddChampSupValeur('NBEXEMPLAIRE', 1, False);
      lTobTemp.AddChampSupValeur('FILTREUTILISE', '', False);
      Break;
    end;
  end;
{$ENDIF}

  MiseAJourFGrille;
  FGrille.Row := FGrille.RowCount - 1;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBHaut(Sender: TObject);
var lTobTemp : Tob;
    lIndex   : Integer;
begin
  if (FTobSelect.Detail.Count = 0) or (FGrille.Row <= 1) then Exit;
  EnregistrementLigne(FGrille.Row);
  lTobTemp := Tob(FGrille.Objects[0, FGrille.Row]);
  lIndex := lTobTemp.GetIndex;
  lTobTemp.ChangeParent(FTobSelect, lIndex - 1);
  MiseAJourFGrille;
  FGrille.Row := lIndex;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBBas(Sender: TObject);
var lTobTemp : Tob;
    lIndex   : Integer;
begin
  if (FTobSelect.Detail.Count = 0) or (FTobSelect.Detail.Count = FGrille.Row) then Exit;
  EnregistrementLigne(FGrille.Row);
  lTobTemp := Tob(FGrille.Objects[0, FGrille.Row]);
  lIndex := lTobTemp.GetIndex;
  lTobTemp.ChangeParent(FTobSelect, lIndex + 1);
  MiseAJourFGrille;
  FGrille.Row := lIndex+2;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/11/2001
Modifié le ... :   /  /
Description .. : Supprime tous les états de la TOB
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBEfface(Sender: TObject);
begin
  FTobSelect.ClearDetail;
  MiseAJourFGrille;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBValider(Sender: TObject);
{$IFDEF DP}
{$ELSE}
var lCEtatsChaines : TCEtatsChaines;
{$ENDIF}
begin
  inherited;

  //EnregistreEnXML;
  //Exit;

{$IFDEF CCSTD}
  Exit; // GCO - 25/04/2006 - FQ 17053
{$ENDIF}

  if FTobSelect.Detail.Count = 0 then
  begin
    PgiInfo('Vous devez sélectionner les états que vous voulez imprimer.', Ecran.Caption);
    Exit;
  end;

  if FCBAuFormatPdf.Checked then
  begin
    if Trim(FNomPdf.Text) = '' then
    begin
      PgiInfo('Vous devez renseigner le nom du fichier PDF à générer.', Ecran.Caption);
      Exit;
    end;
  end;

  // Pour enregistrer la liste sélectionné si la combo de choix du filtre est visible
  EnregistrementLigne( FGrille.Row );

  if (FCBAuFormatPdf.Checked) and (not FCBMultiPdf.Checked) then
  begin
    if FileExists( FNomPdf.Text ) then
    begin
      // Suppression du fichier PDF si il existe deja apres confimation du USER
      if PgiAsk('Le fichier pdf existe déjà, voulez vous l''écraser ?', Ecran.Caption) = MrYes then
        FileExistsDelete( FNomPdf.Text );
    end;
  end;

{$IFDEF DP}
  if not FileExistsDelete(FStFichierXML) then
  begin
    PgiError(FStFichierXML, 'Impossible de supprimer le fichier temporaire');
    Exit;
  end;

  EnregistreEnXML;

  TFVierge(Ecran).Retour := 'OUI; /' +  FFI_TABLE;
  Ecran.Close;
{$ELSE}

  // Traitement de la Compta : Création de l'objet des états chainés
  lCEtatsChaines := nil;
  try
    lCEtatsChaines := TCEtatsChaines.Create ;

    // Utilisation du Format PDF pour l'impression
    if FCBAuFormatPDF.Checked then
    begin
      lCEtatsChaines.CritEdtChaine.AuFormatPDF := FCBAuFormatPDF.Checked;
      lCEtatsChaines.CritEdtChaine.NomPDF      := FNomPdf.Text;
      lCEtatsChaines.CritEdtChaine.MultiPdf    := FCBMultiPdf.Checked;
    end;

    // Utilisation des critères standards
    if FCBCritere.Checked then
    begin
      // Compatibilité
      lCEtatsChaines.CritEdtChaine.UtiliseCritStd := True;
      lCEtatsChaines.CritEdtChaine.Exercice.Code  := ComboExercice.Value;
      if IsValidDate(DateCDu.Text) then
        lCEtatsChaines.CritEdtChaine.Exercice.Deb := StrToDate(DateCDu.Text);
      if IsValidDate(DateCAu.Text) then
        lCEtatsChaines.CritEdtChaine.Exercice.Fin := StrToDate(DateCAu.Text);
      lCEtatsChaines.CritEdtChaine.NatureCompte   := ComboNature.Value;
      lCEtatsChaines.CritEdtChaine.ModeSelection  := ComboMode.Value;
      lCEtatsChaines.CritEdtChaine.TypeEcriture   := ComboType.Value;

      //Paie
      if IsValidDate(GetControlText('PGDATEDEB')) then
        lCEtatsChaines.CritEdtChaine.PGDateDeb := StrToDate(GetControlText('PGDATEDEB'));
      if IsValidDate(GetControlText('PGDATEFIN')) then
        lCEtatsChaines.CritEdtChaine.PGDateFin := StrToDate(GetControlText('PGDATEFIN'));

      // Critères spécifiques DUCS
      lCEtatsChaines.CritEdtChaine.PGNatDeb  := GetControlText('PGNATDEB');
      lCEtatsChaines.CritEdtChaine.PGNatFin  := GetControlText('PGNATFIN');
      lCEtatsChaines.CritEdtChaine.PGEtabDeb := GetControlText('PGETABDEB');
      lCEtatsChaines.CritEdtChaine.PGEtabFin := GetControlText('PGETABFIN');
    end;

    lCEtatsChaines.Execute( FTobSelect ) ;

  finally
    lCEtatsChaines.Free;
  end;

  // Suppression du PDF générer suite à l' annulation
  if VH^.StopEdition then
  begin
    VH^.StopEdition := False;
    Pgiinfo('Traitement interrompu par l''utilisateur...', Ecran.Caption );
  end
  else
    Pgiinfo('Traitement terminé...', Ecran.Caption );

{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickFCBCritere(Sender: TObject);
begin
  if FstType = 'CP' then
  begin
    FCPGBCritere.Enabled := FCBCritere.Checked;
    ComboNature.Color    := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    ComboMode.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    ComboExercice.Color  := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    DateCDu.Color        := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    DateCAu.Color        := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    ComboType.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
  end
  else
  begin
    FPGGBCritere.Enabled := FCBCritere.Checked;
    PGDateDeb.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    PGDateFin.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    // critères paie spécifiques DUCS
    PGNatDeb.Color       := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    PGNatFin.Color       := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    PGEtabDeb.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    PGEtabFin.Color      := IIF(FCBCritere.Checked, clWindow, clBtnFace);
    // 
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 25/06/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickFCBAuFormatPdf(Sender: TObject);
begin
  FGBPDF.Enabled := FCBAuFormatPdf.Checked;
  FNomPdf.Color  := IIF(FCBAuFormatPdf.Checked, clWindow, clBtnFace);
  BNomPdf.Enabled:= FCBAuFormatPdf.Checked;

  if not FCBAuFormatPdf.Checked then
  begin
    FNomPdf.Text := '';
    FCBMultiPdf.Checked := False;
  end
  else
  begin
  {$IFNDEF EAGLCLIENT}
    if (Trim(FNomPDF.Text) = '') and (VH_Doss<>Nil) then
      FNomPDF.Text := VH_Doss.PathDos + '\' ;
  {$ENDIF}

    if FNomPDF.CanFocus then
    begin
      FNomPDF.SetFocus;
      FNomPDF.SelStart := Length(FNomPDF.Text); // Se place à la fin du mot
    end;
  end;

{$IFDEF DP}
  FCBMultiPdf.Enabled := FCBAuFormatPdf.Checked;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10: BValider.Click;

    VK_F5: if (FGrille.Col = cColFiltreUtilise) then LoadLookUpList(FGrille.Row);

    VK_F12: begin
              if FListe.Focused then FGrille.SetFocus
              else FListe.setFocus;
            end;

    VK_DELETE: if Shift = [ssCtrl] then OnClickBGauche(nil);

    VK_SPACE: if FListe.Focused then OnClickBDroit(nil);

    VK_RETURN: Key := VK_TAB;

    VK_UP: if Shift = [ssCtrl] then OnClickBHaut(nil);

    VK_DOWN: if Shift = [ssCtrl] then OnClickBBas(nil);

  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/11/2001
Modifié le ... : 11/10/2006
Description .. : Chargement de FTobSelect dans FGrille
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.MiseAJourFGrille;
var i : integer;
    lTobTemp : Tob;
begin
  FGrille.VidePile( False );

  if FTobSelect.Detail.Count = 0 then
  begin
    FGrille.RowCount := 2;
    FGrille.FixedRows := 1;
    FGrille.Cells[cColID, 1] := '';
    FGrille.Cells[cColLibelleEtat, 1] := '';
    FGrille.Cells[cColNbExemplaire, 1] := '';
    FGrille.Cells[cColFiltreUtilise, 1] := '';
    FGrille.Cells[cColNatureEtat, 1] := '';
    FGrille.Cells[cColCodeEtat, 1] := '';
  end
  else
  begin
    FGrille.RowCount := FTobSelect.Detail.Count + 1;
    for i := 0 to FTobSelect.Detail.Count - 1 do
    begin
      lTobTemp := FTobSelect.Detail[i];
      FGrille.Cells[cColID,            i + 1] := lTobTemp.GetValue('ID');
      FGrille.Cells[cColLibelleEtat,   i + 1] := lTobTemp.GetValue('LIBELLEETAT');
      FGrille.Cells[cColNbExemplaire,  i + 1] := lTobTemp.GetValue('NBEXEMPLAIRE');
      FGrille.Cells[cColFiltreUtilise, i + 1] := lTobTemp.GetValue('FILTREUTILISE');
      FGrille.Objects[0,               i + 1] := lTobTemp;
    end;
  end;

  FGrille.Row := 1;
  FGrille.Col := cColNbExemplaire;
  FGrille.ElipsisButton := False;

  THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(FGrille);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/12/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnDblClickFListe(Sender: TObject);
begin
  // Ajoute l'élément sélectionné dans la liste à imprimer
  OnClickBDroit(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnRowExitFGrille(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if FTobSelect.Detail.Count = 0 then Exit;
  EnregistrementLigne(Ou);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/06/2002
Modifié le ... : 26/06/2002
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnCellEnterFGrille(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if FTobSelect.Detail.Count = 0 then Exit;
{$IFDEF DP}
  FGrille.ElipsisButton := False;
{$ELSE}
  FGrille.ElipsisButton := (FGrille.Col = cColFiltreUtilise);
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Traitement lors de l' arrivée dans une cellule de la grille
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnCellExitFGrille(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if ACol = cColNbExemplaire then
  begin
    if Trim(FGrille.Cells[ACol, ARow]) = '' then
      FGrille.Cells[ACol, ARow] := '1';
  end
  else
    if ACol = cColFiltreUtilise then
      EnregistrementLigne(FGrille.Row);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnElipsisClickFGrille(Sender: TObject);
begin
  if FGrille.Col = cColFiltreUtilise then
    LoadLookUpList(FGrille.Row);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/10/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ETATSCHAINES.LoadLookUpList(lOu: integer);
var
  lTob             : Tob;
  lTobListeFiltre  : Tob;
  lTobFille        : Tob;
  lTobFiltreSelect : Tob;
  lSt              : string;
  lStName,lStUser  : string;
  lListeFiltre     : TListByUser;
  j                : integer;

{$IFDEF YPARAMEDITION}
  lTobTemp : Tob;
{$ENDIF}  

  function GetNomFiltre ( vID : string ) : string;
  var i : integer;
  begin
    Result := '';
  {$IFDEF YPARAMEDITION}
  {$ELSE}
    for i := 0 to cNbEtat do
    begin
      if vID = cListeEtat[i].stID then
        Result := cListeEtat[i].stNomFiltre;
    end;
  {$ENDIF}
  end;

begin
  lSt := FGrille.Cells[FGrille.Col, lOu];
  FGrille.Cells[FGrille.Col, lOu] := '';
  lTob := TOB(FGrille.Objects[0, lOu]);

  lListeFiltre := TListByUser.Create( toFiltre );
  if lListeFiltre <> nil then
  {$IFDEF YPARAMEDITION}
    lTobTemp := FTobYParamEdition.FindFirst(['YED_CODEID'],[lTob.GetString('ID')], False);
    if lTobTemp <> nil then
      lListeFiltre.LoadDB( lTobTemp.GetString('YED_NOMFILTRE') );
  {$ELSE}
    lListeFiltre.LoadDB( GetNomFiltre(lTob.GetString('ID')) );
  {$ENDIF}

  lTobListeFiltre := Tob.Create('', nil, -1);
  for j := 0 to lListeFiltre.GetListeObjects.count - 1 do
  begin
    lTobFille := Tob.Create( '', lTobListeFiltre, -1);
    lStName := lListeFiltre.GetListeObjects[j];
    lStUser := ReadTokenPipe(lStName,'|');  //le Nom du filtre reste dans lStName
    lTobFille.AddChampSupValeur('FILTRE', lStName, False);
    lTobFille.AddChampSupValeur('UTILISATEUR', lStUser, False);
  end;

  if lTobListeFiltre.Detail.Count <> 0 then
  begin
    lTobFiltreSelect := LookUpTob( FGrille, lTobListeFiltre, 'Liste des filtres', 'FILTRE', 'Nom du filtre' );
    if Assigned(lTobFiltreSelect) then
    begin
      lTob.PutValue('FILTREUTILISE', lTobFiltreSelect.GetValue('FILTRE'));
      FGrille.Cells[FGrille.Col, lOu] := lTobFiltreSelect.GetValue('FILTRE');
    end;
  end;

  if FGrille.Cells[FGrille.Col, lOu] = '' then
    FGrille.Cells[FGrille.Col, lOu] := lSt;

  FreeAndNil( lTobListeFiltre );
  FreeAndNil( lListeFiltre );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/12/2001
Modifié le ... : 12/12/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnKeyPressFGrille(Sender: TObject; var Key: Char);
begin
  if (FGrille.Col = cColNbExemplaire) then
  begin
    if not (Key in ['0'..'9', #8]) then
      Key := #0;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Enregistre les données de la ligne sélectionnée dans la TOB
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.EnregistrementLigne(lRow: integer);
var
  lTob: Tob;
begin
  lTob := TOB(FGrille.Objects[0, lRow]);
  if lTob <> nil then
  begin
    lTob.PutValue('NBEXEMPLAIRE', StrToInt(FGrille.Cells[cColNbExemplaire, lRow]));
    if (FGrille.ValCombo <> nil) and (FGrille.ValCombo.Text <> '') then
      lTob.PutValue('FILTREUTILISE', FGrille.ValCombo.Text)
    else
      lTob.PutValue('FILTREUTILISE', FGrille.Cells[cColFiltreUtilise, lRow]);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Parcours la grille et charge la tob avec les données
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.LoadTobFromGrille;
var
  i : integer;
  lTobTemp : Tob;
begin
  //if Trim(FGrille.Cells[cColLibelleEtat, 1]) = '' then Exit;

  FTobSelect.ClearDetail;
  for i := 1 to FGrille.RowCount - 1 do
  begin
    if Trim(FGrille.Cells[cColLibelleEtat, i]) <> '' then
    begin
      lTobTemp := Tob.Create('', FTobSelect, -1);
      lTobTemp.AddChampSupValeur('ID',           FGrille.Cells[cColID, i], False);
      lTobTemp.AddChampSupValeur('LIBELLEETAT',  FGrille.Cells[cColLibelleEtat, i], False);
      lTobTemp.AddChampSupValeur('NBEXEMPLAIRE', FGrille.Cells[cColNbExemplaire, i], False);

      if Trim(FGrille.Cells[cColFiltreUtilise, i]) = '' then
        lTobTemp.AddChampSupValeur('FILTREUTILISE', '', False)
      else
        lTobTemp.AddChampSupValeur('FILTREUTILISE', FGrille.Cells[cColFiltreUtilise, i], False);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Souad B.
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Ouverture du fichier .PDF
Mots clefs ... : ETATSCHAINES
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClickBtnNomPdf(sender: TObject);
var lFileN : String;
begin
  lFileN := GetControlText('FNOMPDF');
  if lFileN <> '' then
  begin
    if FileExists(lFileN) then
      ShellExecute( 0, PCHAR('open'), PChar(lFileN), nil, nil, SW_RESTORE)
    else
      PGIBox('Fichier inexistant.', TFVierge(Ecran).Caption);
  end
  else
    PGIBox('Fichier non renseigné.', TFVierge(Ecran).Caption);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/04/2003
Modifié le ... : 08/04/2003
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_ETATSCHAINES.OnElipsisClickFNomPdf(Sender: TObject);
var lOpenDialog : TOpenDialog;
    lStFileName : string;
begin
  lStFileName := FNomPdf.Text;
  lOpenDialog := TOpenDialog.Create( nil );
  lOpenDialog.Filter := 'Fichier pdf (*.pdf)|*.pdf';
  lOpenDialog.DefaultExt := 'pdf';
  lOpenDialog.FilterIndex := 1;
  lOpenDialog.InitialDir := FNomPdf.Text;
  lOpenDialog.Execute;
  FNomPdf.Text := IIF( lOpenDialog.FileName <> '', lOpenDialog.FileName, lStFileName);
  lOpenDialog.Free;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 16/03/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnExitFNomPdf(Sender: TObject);
begin
  if Pos('.PDF', UpperCase(FNomPdf.Text)) = 0 then
    FNomPdf.Text := FNomPdf.Text + '.pdf';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/11/2001
Modifié le ... : 12/10/2006
Description .. : Fermeture de la fenêtre
Mots clefs ... :
*****************************************************************}
procedure TOF_ETATSCHAINES.OnClose;
begin
  Application.ProcessMessages;

  FreeAndNil(FObjFiltre);
  FreeAndNil(FTobSelect);
  FreeAndNil(FTobYParamEdition);

  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_ETATSCHAINES]);
end.






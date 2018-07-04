{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/11/2002
Modifié le ... : 23/08/2005
Description .. : Source TOF de la FICHE : CPGLGENE ()
Suite ........ : GCO - 06/02/2004 -
Suite ........ : -> Correction du bug des ecritures de centralisation qui
Suite ........ :    n'apparaissaient pas lors de la sélection d'une devise
Suite ........ :    où d'un établissement
Suite ........ : GCO - 09/02/2004
Suite ........ : -> Correction du bug E_QUALIFPIECE pour la justification
Suite ........ : de
Suite ........ : solde ( on force le type d'écritures à N )
Suite ........ : GCO - 17/02/2004
Suite ........ : -> Passage des paramètres à l'état avec un CritEdt
Suite ........ : GCO - 07/07/2004
Suite ........ : -> FB 13405 ( Recalcul systématique des paramètres de
Suite ........ : l'onglet
Suite ........ :    DEV à chaque lancement d'édition )
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Suite ........ : GCO - 23/08/2005 - FQ 16404
Suite ........ : GCO - 05/09/2005 - FQ 16320
Mots clefs ... : TOF;CPGLGENE
*****************************************************************}
unit uTofCPGLGene;

interface

uses StdCtrls,
  Controls,
  Classes,
  Graphics,
{$IFDEF EAGLCLIENT}
  Maineagl,
  eQRS1,
{$ELSE}
  Fe_main,
  DB,
  {$IFDEF DBXPRESS}
  uDbxDataSet,
  {$ELSE}
  dbtables,
  {$ENDIF}
  QRS1,
{$ENDIF}
  Menus, // TPopUpMenu
  forms,
  Hqry,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  Htb97,
  Ent1,
  HMsgBox,
  TofMeth,
  UTOF,
  uTob,
  Windows,
  uLibWindows,
  uLibExercice, // CExerciceVersRelatif
  Spin,         // TSpinEdit
  SaisUtil,     // QuelExoDt
  ED_TOOLS,     // InitMoveProgressForm
  LicUtil,      // CryptageSt
  AGLInit,      // TheData
  ParamSoc,     // GetParamSocSecur
  uTobDebug,    // TobDebug
{$IFDEF EAGLCLIENT}
  uTileAGL,     // LanceEtatTob
{$ELSE}
  EdtREtat,     // LanceEtatTob
{$ENDIF}
  uPrintF1Book, // PDFBatch
  HPdfPrev,     // PreviewPDFFile
  HPdfViewer,   // PrintPDF
  Printers,     // Printer
  CritEdt,      // ClassCritEdt
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UtilEdt;      // MajEditionLegal

type
  TOF_CPGLGENE = class(TOF_METH)
  private
    EdLegale         : Boolean;
    E_GENERAL        : THEdit;
    E_GENERAL_       : THEdit;
    E_AUXILIAIRE     : THEdit;
    E_AUXILIAIRE_    : THEdit;
    GeneParAuxi      : THEdit;
    E_DATECOMPTABLE  : THEdit;
    E_DATECOMPTABLE_ : THEdit;
    E_NUMEROPIECE    : THEdit;
    E_NUMEROPIECE_   : THEdit;
    CPTEXCEPT        : THEdit;
    LIBEQUALIFPIECE  : THEdit;
    CUMULAU          : THEdit;
    E_REFINTERNE     : THEdit;
    E_EXERCICE       : THValComboBox;
    E_DEVISE         : THValComboBox;
    G_NATUREGENE     : THValComboBox;
    MODESELECT       : THValComboBox;
    TypeQuantite     : THValComboBox;

    E_QUALIFPIECE    : THMultiValComboBox;
    G_CYCLEREVISION  : THEdit;

    // Correspondances
    CORRESP          : THValComboBox;
    CORRESPDE        : THValComboBox;
    CORRESPA         : THValComboBox;

    // Tables Libres
    TABLELIBRE       : THValComboBox;
    LIBREDE          : THEdit;
    LIBREA           : THEdit;

    CBEnSituation    : TCheckBox;
    CBTous           : TCheckBox;
    ECRVALIDE        : TCheckBox;
    ECRLETTRE        : TCheckBox;
    ECRPOINTE        : TCheckBox;
    AFFANOUVEAU      : TCheckBox;
    AFFCENTRA        : TCheckBox;
    AFFCUMUL         : TCheckBox;

    EDateSituation   : THEdit;
    DateDebutExo     : THEdit;
    Exercice         : THEdit;

    Pages            : TPageControl;
    RGDevise         : THRadioGroup;
    RUPTURE          : THRadioGroup;
    RUPTURES         : TTabSheet;
    RUPTURETYPE      : THRadioGroup;
    RUPCORRESP       : TGroupBox;
    RUPLIBRES        : TGroupBox;
    RUPGROUPES       : TGroupBox;
    RUPCYCLEREVISION : TCheckBox;

    NIVORUPTURE      : TSpinEdit;
    AFFDEVISE        : THEdit;
    DEVPIVOT         : THEdit;
    XX_RUPTURE       : THEdit;
    AVECRUPTYPE      : THEdit;
    AVECRUPTURE      : THEdit;
    AVECNIVORUPTURE  : THEdit;
    OKENTETE         : THEdit;
    OKRESUME         : THEdit;
    ETATINCOMPLET    : THEdit;
    LIBANO           : THEdit;

    CBPreview        : TCheckBox;
    CPTCorrespExist  : TCheckBox;
    CPTLibreExist    : TCheckBox;

    procedure OnExitE_General         ( Sender : TObject );
    procedure OnExitE_Auxiliaire      ( Sender : TObject );
    procedure OnChangeG_NATUREGENE    ( Sender : TObject );
    procedure OnExitE_DateComptable   ( Sender : TObject );
    procedure OnChangeE_Exercice      ( Sender : TObject );
    procedure OnChangeE_Devise        ( Sender : TObject );
    procedure OnClickCBEnSituation    ( Sender : TObject );
    procedure OnClickRupture          ( Sender : TObject );
    procedure OnClickRuptureType      ( Sender : TObject );
    procedure FTimerTimer             ( Sender : TObject );
    procedure OnChangeCorresp         ( Sender : TObject );
    procedure OnChangeTableLibre      ( Sender : TObject );
    procedure OnEnterQualifPiece      ( Sender : TObject ); // SBO - 12/10/2004
    procedure OnClickRupCycleRevision ( Sender : TObject );
    procedure AffectePlusE_General;
    procedure AuxiElipsisClick        ( Sender : TObject ) ;

  public
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnNew; override;
    procedure ChargementCritEdt; override;
    procedure MySelectFiltre; override;
    procedure OnChangeFiltre ( Sender : TObject ) ; override;

  private
    FStArgument            : string;   // Argument passé en paramètre
    FStWhereEtat           : string;   // Condition Where envoyée à l'état
    FBoGeneParAuxi         : Boolean;  // Grand Livre Général par Auxi ?
    FBoGeneParQte          : Boolean;  // Grand Livre Général par Quantité ?
    FBoOkDecoupage         : Boolean;  // Découpage du grand livre
    FExoDate               : TExoDate; // Contexte de travail sur les dates

    procedure ControleAvantEdition;

    procedure GestionCentralisation;       // Point d'éntrée de la centralisation
    procedure SupprimeCentralisation;      // Suppression des écritures "Z"
    procedure EffectueCentralisation;      // Insertion des écritures "Z"
    procedure CBlocageGLGene;              // Blocage dans la Table COURRIER
    procedure CDeBlocageGLGene;            // Déblocage
    function  CEstBloqueGLGene : Boolean;  // Controle de la présence du Verrou

    procedure GenereCumulGL;

    function GenereConditionWhereEtat   : string;
    function RecupAutresCriteres        : string;
    function RecupCriteresCommuns       : string;
    function RecupCriteresModeSelection( vStWhereEtat : string ) : string;
    function RenvoiOrderBy              : string;

    procedure GestionCritereDEV;

{$IFDEF EAGLCLIENT}
{$ELSE}
    procedure GestionTailleGrandLivre( vStwhereCritere : string );
{$ENDIF}

  end;

function CPLanceFiche_CPGLGENE        ( vStParam: string = '') : string ;
function CPLanceFiche_CPGLGENEPARAUXI ( vStParam: string = '') : string ;
function CPLanceFiche_CPGLGENEPARQTE  ( vStParam: string = '') : string ;
function OnMyValideCritereEvent : Boolean;

implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  {$IFNDEF EDTQR}
    UTofMulParamGen,
  {$ENDIF}
    uLibCalcEdtCompta, // FTobCumulAu
    utilPGI,
    uMultiDossierUtil; // RequeteMultiDossier

const cNbEcrMax = 10000;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPGLGENE( vStParam : string = ''): string;
begin
  // Ne jamais toucher au vStParam
  Result := AGLLanceFiche('CP', 'CPGLGENE', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPGLGENEPARAUXI ( vStParam : string = '') : string ;
begin
  // Ne jamais toucher au vStParam
  Result := AGLLanceFiche('CP', 'CPGLGENEPARAUXI', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_CPGLGENEPARQTE  ( vStParam: string = '') : string ;
begin
  Result := AGLLanceFiche('CP', 'CPGLGENEPARQTE', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnArgument(S: string);
begin
  inherited;
  FStArgument := S; // Récupératrion des paramètres

  // GCO - 11/07/2006 - FQ 18363
  PAvances := TTabSheet(GetControl('AVANCES', true));
  Pavances.TabVisible := True;
  TFQRS1(Ecran).CritAvancesVisibled := True;

  if FStArgument = 'LEGALE' then EdLegale := True;

  if ReadTokenSt(FStArgument) = 'MULTIDOSSIER' then
  begin
    SetControlVisible( 'MULTIDOSSIER',   True ) ;
    SetControlVisible( 'TMULTIDOSSIER',  True ) ;
    SetControlVisible( 'BVMULTIDOSSIER', True ) ;
    //SetControlVisible( 'RUPTUREMAX',     True ) ;
  end;

  // Détermination du Type de Grand Livre Général
  FBoGeneParAuxi := IIF( Ecran.Name = 'CPGLGENEPARAUXI', True, False);
  FBoGeneParQte  := IIF( Ecran.Name = 'CPGLGENEPARQTE', True, False);

  // Branchement de HelpContext
  Ecran.HelpContext := IIF( FBoGeneParAuxi, 7415000, 7427000 );

  E_GENERAL   := THEdit(GetControl('EGENERAL', True));
  E_GENERAL_  := THEdit(GetControl('EGENERAL_', True));
  GeneParAuxi := THEdit(GetControl('GENEPARAUXI', True));

  E_DATECOMPTABLE  := THEdit(GetControl('EDATECOMPTABLE', True));
  E_DATECOMPTABLE_ := THEdit(GetControl('EDATECOMPTABLE_', True));
  E_NUMEROPIECE    := THEdit(GetControl('E_NUMEROPIECE', True));
  E_NUMEROPIECE_   := THEdit(GetControl('E_NUMEROPIECE_', True));
  E_REFINTERNE     := THEdit(GetControl('E_REFINTERNE', True));
  E_EXERCICE       := THValComboBox(GetControl('EEXERCICE', True));
  CInitComboExercice(E_EXERCICE);

  E_DEVISE         := THValComboBox(GetControl('E_DEVISE', True));
  E_QUALIFPIECE    := THMultiValComboBox(GetControl('EQUALIFPIECE', True));
  G_NATUREGENE     := THValComboBox(GetControl('G_NATUREGENE', True));

  MODESELECT       := THValComboBox(GetControl('MODESELECT', True));
  AFFDEVISE        := THEdit(GetControl('AFFDEVISE', True));
  DEVPIVOT         := THEdit(GetControl('DEVPIVOT', True));
  RGDEVISE         := THRadioGroup(GetControl('RGDEVISE', True));
  CPTEXCEPT        := THEdit(GetControl('CPTEXCEPT', True));

  ECRValide        := TCheckBox(GetControl('ECRVALIDE', True));
  ECRLETTRE        := TCheckBox(GetControl('ECRLETTRE', True));
  ECRPOINTE        := TCheckBox(GetControl('ECRPOINTE', True));
  AFFANOUVEAU      := TCheckBox(GetControl('AFFANOUVEAU', True));
  AFFCENTRA        := TCheckBox(GetControl('AFFCENTRA', True));
  AFFCUMUL         := TCheckBox(GetControl('AFFCUMUL', True));

  CBEnSituation    := TCheckBox(GetControl('CBENSITUATION', True));
  CBTous           := TCheckBox(GetControl('CBJUSTIFTOUS', True));
  EDateSituation   := THEdit(GetControl('EDATESITUATION', True));

  DateDebutExo     := THEdit(GetControl('DATEDEBUTEXO', True));
  Exercice         := THEdit(GetControl('EXERCICE', True));

  Pages            := TPageControl(GetControl('Pages', True));
  RUPTURE          := THRadioGroup(GetControl('RUPTURE', True));
  RUPTURES         := TTabSheet(GetControl('RUPTURES', True));
  RUPTURETYPE      := THRadioGroup(GetControl('RUPTURETYPE', True));
  RUPCORRESP       := TGroupBox(GetControl('RUPCORRESP', True));
  RUPLIBRES        := TGroupBox(GetControl('RUPLIBRES', True));
  RUPGROUPES       := TGroupBox(GetControl('RUPGROUPES', True));
  CORRESP          := THValComboBox(GetControl('CORRESP', True));
  CORRESPDE        := THValComboBox(GetControl('CORRESPDE', True));
  CORRESPA         := THValComboBox(GetControl('CORRESPA', True));

  CBPreview        := TCheckBox(GetControl('FAPERCU', True));
  // Case à cocher : Uniquement les comptes associés
  CPTCORRESPEXIST  := TCheckBox(GetControl('CPTCORRESPEXIST', True));
  CPTLIBREEXIST    := TCheckBox(GetControl('CPTLIBRESEXIST', True));

  // Zones cachées
  XX_RUPTURE       := THEdit(GetControl('XX_RUPTURE', True));
  AVECRUPTYPE      := THEdit(GetControl('AVECRUPTYPE', True));
  AVECRUPTURE      := THEdit(GetControl('AVECRUPTURE', True));
  AVECNIVORUPTURE  := THEdit(GetControl('AVECNIVORUPTURE', True));
  NIVORUPTURE      := TSpinEdit(GetControl('NIVORUPTURE', True));
  OKENTETE         := THEdit(GetControl('OKENTETE', True));
  OKRESUME         := THEdit(GetControl('OKRESUME', True));
  ETATINCOMPLET    := THEdit(GetControl('ETATINCOMPLET', True));
  LIBANO           := THEdit(GetControl('LIBANO', True));

  TABLELIBRE       := THValComboBox(Getcontrol('TABLELIBRE', True));
  LIBREDE          := THEdit(GetControl('LIBREDE', True));
  LIBREA	         := THEdit(GetControl('LIBREA', True));

  LIBEQUALIFPIECE  := THEdit(GetControl('LIBEQUALIFPIECE', True));
  CUMULAU          := THEdit(GetControl('CUMULAU', True));

  // ---------------------- Cycle de révision ----------------------------------
  // Pas de paramètre True pour le GetControl, le GL auxiliaire ne fait pas la révision
  G_CYCLEREVISION  := THEdit(GetControl('GCYCLEREVISION', False));
  RUPCYCLEREVISION := TCheckBox(GetControl('RUPCYCLEREVISION', False));

  // ----------------- Init des PROPERTY ---------------------------------------
  E_GENERAL.MaxLength  := VH^.CPta[fbGene].Lg;
  E_GENERAL_.MaxLength := VH^.CPta[fbGene].Lg;

  RUPTURES.TabVisible := False;
  RUPCORRESP.Visible  := False;
  RUPLIBRES.Visible   := False;
  RUPGROUPES.Visible  := True;

  // Type de plan comptable :
  if GetParamSocSecur('SO_CORSGE2', False, True) then
    Corresp.plus := 'AND (CO_CODE = "GE1" OR CO_CODE = "GE2")'
  else
    Corresp.plus := 'AND CO_CODE = "GE1"';

  // ----------------- Branchement des événements ------------------------------
  G_NATUREGENE.OnChange := OnChangeG_NATUREGENE;
  E_EXERCICE.OnChange   := OnChangeE_Exercice;

  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
    E_EXERCICE.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    E_EXERCICE.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  E_General.OnExit         := OnExitE_General;
  E_General_.OnExit        := OnExitE_General;
  CORRESP.OnChange         := OnChangeCorresp;
  TABLELIBRE.OnChange      := OnChangeTableLibre;
  E_DATECOMPTABLE.OnExit   := OnExitE_DateComptable;
  E_DATECOMPTABLE.OnExit   := OnExitE_DateComptable;
  E_DEVISE.OnChange        := OnChangeE_Devise;
  CBEnSituation.OnClick    := OnClickCBEnSituation;
  Rupture.OnClick          := OnClickRupture;
  RuptureType.OnClick      := OnClickRuptureType;
  FTimer.OnTimer           := FTimerTimer;
  E_QUALIFPIECE.OnEnter    := OnEnterQualifPiece ; // SBO 12/10/2004

  if (RupCycleRevision <> nil) and (G_CycleRevision <> nil) then // Vaut nil en Général par Auxiliaire
  begin
    RupCycleRevision.Enabled := VH^.Revision.Plan <> '';
    // GCO - 11/06/2007 - FQ 20630
    G_CycleRevision.Plus := ' AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"';
    RupCycleRevision.OnClick := OnClickRupCycleRevision;
  end;  

  // GCO - 27/10/2004 - FQ 14787
  {$IFDEF GIL}
    SetControlProperty('TABDEV', 'TABVISIBLE', True);
  {$ELSE}
    SetControlProperty('TABDEV', 'TABVISIBLE', (V_PGI.Sav) and (V_PGI.PassWord = CryptageSt(DayPass(Date))));
  {$ENDIF}

  if FBoGeneParAuxi then
  begin
    E_AUXILIAIRE            := THEdit(GetControl('EAUXILIAIRE', True));
    E_AUXILIAIRE_           := THEdit(GetControl('EAUXILIAIRE_', True));
    E_Auxiliaire.MaxLength  := VH^.Cpta[fbAux].Lg;
    E_Auxiliaire_.MaxLength := VH^.Cpta[fbAux].Lg;
    E_Auxiliaire.OnExit     := OnExitE_Auxiliaire;
    E_Auxiliaire_.OnExit    := OnExitE_Auxiliaire;
    CBEnSituation.Enabled   := False;
    EDateSituation.Enabled  := False;
    CBTous.Visible          := False;
    CBTous.Enabled          := False;
    CBTous.Checked          := False;

    if GetParamSocSecur('SO_CPMULTIERS', false) then
    begin
      E_Auxiliaire.OnElipsisClick:=AuxiElipsisClick;
      E_Auxiliaire_.OnElipsisClick:=AuxiElipsisClick;
    end;
  end;

  if FBoGeneParQte then
  begin
    TypeQuantite := THValComboBox(GetControl('TYPEQUANTITE', True));
    TypeQuantite.ItemIndex := 0;

    AFFCumul.Visible := False;
    AFFCumul.Checked := False;
  end;

  AffectePlusE_General;

  // GCO - 03/01/2007 - FQ 19375
  FBoOkDecoupage := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnUpdate;
begin
  inherited;           
  // Grand Livre
  TFQRS1(Ecran).WhereSQL := FStWhereEtat + RenvoiOrderBy;
  If EdLegale then MajEditionLegal('GLG',CRelatifVersExercice(E_EXERCICE.Value),E_DateComptable.Text,E_DateComptable_.Text);

  //06/12/2006 YMO Norme NF
  // FQ 19498 - CA - Pas de log dans le cas du grand-livre général par auxiliaire
  // BVE 30.08.07 : Nouvelle fonction de tracage
  if (FProvisoire.Text='') and (not FBoGeneParAuxi) then
{$IFNDEF CERTIFNF}
     CPEnregistreLog('GRAND-LIVRE '+E_DateComptable.Text+' - '+E_DateComptable_.Text);
{$ELSE}
     CPEnregistreJalEvent('CEG','Edition grand-livre','GRAND-LIVRE '+E_DateComptable.Text+' - '+E_DateComptable_.Text);
{$ENDIF}               
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.GenereCumulGL;
var lStSelect : string;
    lStWhere  : string;
    lStWhereCommun : string;
    lTobTemp : Tob;
    lStValEQualifPiece : string;
    lStLibEQualifPiece : string;
    lStWhereGene       : string;
    lStRgrp            : string;
    lQuery             : TQuery;
begin
  if FTobCumulAnoGL = nil then
    FTobCumulAnoGL := Tob.Create('', nil, -1)
  else
    FTobCumulAnoGL.ClearDetail;

  if FTobCumulAuGL = nil then
    FTobCumulAuGL := Tob.Create('', nil, -1)
  else
    FTobCumulAuGL.ClearDetail;

  if FTobBilanGestion = nil then
  begin
    FTobBilanGestion := Tob.Create('', nil, -1);
    lTobTemp := Tob.Create('', FTobBilanGestion, -1);
    lTobTemp.AddchampSupValeur('NATURE', 'BIL', False);
    lTobTemp := Tob.Create('', FTobBilanGestion, -1);
    lTobTemp.AddchampSupValeur('NATURE', 'CHA', False);
    lTobTemp := Tob.Create('', FTobBilanGestion, -1);
    lTobTemp.AddchampSupValeur('NATURE', 'PRO', False);
    lTobTemp := Tob.Create('', FTobBilanGestion, -1);
    lTobTemp.AddchampSupValeur('NATURE', 'EXT', False);
    lTobTemp.AddchampSupValeur('DEBIT', 0, True);
    lTobTemp.AddchampSupValeur('CREDIT', 0, True);
  end
  else
  begin
    FTobBilanGestion.PutValueAllFille('DEBIT', 0);
    FTobBilanGestion.PutValueAllFille('CREDIT', 0);
  end;

  lStSelect := 'SELECT E_GENERAL,' + IIF(FBoGeneParAuxi, 'E_AUXILIAIRE,', '');

  if AffDevise.Text = 'X' then
    lStSelect := lStSelect + ' SUM(E_DEBITDEV) DEBIT, SUM(E_CREDITDEV) CREDIT'
  else
    lStSelect := lStSelect + ' SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT';

  // GCO - 18/01/2006 - E_QUALIFPIECE
  TraductionTHMultiValComboBox( E_QualifPiece, lStValEQualifPiece, lStLibEQualifPiece, 'E_QUALIFPIECE', False );
  lStSelect := lStSelect + ' FROM ECRITURE WHERE ' + lStValEQualifPiece;

  // Where Commun aux 2 requêtes
  lStWhereCommun := '';
  if E_Exercice.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_EXERCICE = "' + CRelatifVersExercice(E_EXERCICE.Value) + '"';

  if ComboEtab.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_ETABLISSEMENT = "' + ComboEtab.Value + '"';

  if E_Devise.Value <> '' then
    lStWhereCommun := lStWhereCommun + ' AND E_DEVISE = "' + E_DEVISE.Value + '"';

  // FQ 18698 : Ajout test Généraux pour opti (SBO 24/08/2006)
  lStWhereGene   := ConvertitCaractereJokers(E_GENERAL, E_GENERAL_, 'E_GENERAL');
  if Trim( lStWhereGene ) <> '' then
    lStWhereCommun := lStWhereCommun + ' AND ' + lStWhereGene ;

  if FBoGeneParAuxi then
  begin
    lStWhereGene := ConvertitCaractereJokers(E_Auxiliaire, E_Auxiliaire_, 'E_AUXILIAIRE');
    if Trim( lStWhereGene ) <> '' then
      lStWhereCommun := lStWhereCommun + ' AND ' + lStWhereGene ;
  end ;
  // Fin FQ 18698

  lStWhereCommun := lStWhereCommun +
                    ' GROUP BY E_GENERAL' + IIF(FBoGeneParAuxi, ',E_AUXILIAIRE', '') +
                    ' ORDER BY E_GENERAL' + IIF(FBoGeneParAuxi, ',E_AUXILIAIRE', '');

  lStWhere := ' AND E_DATECOMPTABLE = "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '"';

  lStWhere := lStWhere + ' AND (E_ECRANOUVEAU = "OAN" OR E_ECRANOUVEAU = "H")';

  // GCO - 08/11/2007 - FQ 20377 Mode Multi-société
  if EstMultiSoc then
  begin
    try
      RequeteMultiDossier( Pages, lStRgrp );
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun, lStRgrp);
      FTobCumulAnoGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAnoGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun );

  // Calcul du Cumul AU
  lStWhere := ' AND E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '" AND ' +
               'E_DATECOMPTABLE < "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '" AND ' +
               'E_ECRANOUVEAU = "N"';

  // GCO - 08/11/2007 - FQ 20377
  if EstMultiSoc then
  begin
    try
      lQuery := OpenSelect( lStSelect + lStWhere + lStWhereCommun, lStRgrp);
      FTobCumulAuGL.LoadDetailDB( '', '', '', lQuery, False);
    finally
      Ferme(lQuery);
    end;
  end
  else
    FTobCumulAuGL.LoadDetailFromSql( lStSelect + lStWhere + lStWhereCommun );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.ControleAvantEdition;
var lStsql : string;
    lQuery : TQuery;
begin
  // Remplissage des bornes des comptes généraux
  if (Trim(E_General.Text) = '') or (Trim(E_General_.Text) = '') then
  begin
    // JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker
    if not TestJoker(E_General.Text) then
    begin
      try
        lStSql := 'SELECT MIN(G_GENERAL)MINI, MAX(G_GENERAL)MAXI FROM GENERAUX ' +
                  'WHERE ' + E_General.Plus;

        lQuery := OpenSql( lStSql , True);
        if not lQuery.Eof then
        begin
          if Trim(E_General.Text) = '' then
            E_General.Text := lQuery.FindField('MINI').AsString;

          if Trim(E_General_.Text) = '' then
            E_General_.Text := lQuery.FindField('MAXI').AsString;
        end;

      finally
        Ferme(lQuery);
      end;
    end;
  end;

  if FBoGeneParAuxi then
  begin
    if (Trim(E_Auxiliaire.Text) = '') or (Trim(E_Auxiliaire_.Text) = '') then
    begin
      // JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker
      if not TestJoker(E_Auxiliaire.Text) then begin
        try
          lQuery := OpenSql('SELECT MIN(T_AUXILIAIRE)MINI, MAX(T_AUXILIAIRE)MAXI FROM TIERS', True);
          if Trim(E_Auxiliaire.Text) = '' then
            E_Auxiliaire.Text := lQuery.FindField('MINI').AsString;

          if Trim(E_Auxiliaire_.Text) = '' then
            E_Auxiliaire_.Text := lQuery.FindField('MAXI').AsString;
        finally
          Ferme(lQuery);
        end;
      end;
    end;
  end;

  if E_Exercice.ItemIndex = 0 then
  begin
    if (VH^.ExoV8.Code <> '') then
    begin
      if StrToDate(E_DateComptable.Text) < VH^.ExoV8.Deb then
      begin
        PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(VH^.ExoV8.Deb) + '.' + #13 + #10 +
                'La date a été modifiée.', Ecran.Caption);
        E_DateComptable.Text := DateToStr(VH^.ExoV8.Deb);
      end;
    end
    else
    begin
      // GCO - 02/02/2003 FB 12796
      if StrToDate(E_DateComptable.Text) < VH^.Exercices[1].Deb then
      begin
        PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(VH^.Exercices[1].Deb) + '.' + #13 + #10 +
                'La date a été modifiée.', Ecran.Caption);
        E_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
      end;
    end;

    if CbEnSituation.Checked then
      FExoDate := CtxExercice.QuelExoDate(StrToDate(EDateSituation.Text))
    else
      FExoDate := CtxExercice.QuelExoDate(StrToDate(E_DateComptable.Text));
  end
  else
  begin
    FExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));
    if (StrToDate(E_DateComptable.Text) < FExoDate.Deb) then
    begin
      PgiInfo('Vous ne pouvez pas saisir une date inférieure au ' + DateToStr(FExoDate.Deb) + '.' + #13 + #10 +
              'La date a été modifiée.', Ecran.Caption);
      E_DateComptable.Text := DateToStr(FExoDate.Deb);
    end;
  end;

  // Correction du E_QUALIFPIECE
  // E_QUALIFPIECE est forcé à "N" en justificatif de solde
  if (E_QualifPiece.Text = '') or (CBEnSituation.Checked) then
    E_QualifPiece.Text := 'N;';
    
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/08/2003
Modifié le ... : 12/07/2006
Description .. : Prépare les composants cachés DEV
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.GestionCritereDEV;
var lStTemp : string;
begin
  // Init des CRI_ qui servent pour l'état avec le FExoDate
  GeneParAuxi.Text     := IIF( FBoGeneParAuxi, 'X', '-' );
  AvecRupture.Text     := 'SANS' ;
  AvecRupType.Text     := '' ;
  AvecNivoRupture.Text := '0' ;
  XX_Rupture.Text      := '' ;
  OkEntete.Text        := 'X';
  OkResume.Text        := 'X'; 

  LibAno.Text := TraduireMemoire('Total des A-Nouveaux');
  if E_Exercice.ItemIndex = 0 then
    Exercice.Text := '---'
  else
  begin
    Exercice.Text := FExoDate.Code;

    // Test de la présence d' A-Nouveaux
    if ExisteSQL('SELECT E_GENERAL FROM ECRITURE WHERE ' +
       'E_EXERCICE = "' + FExoDate.Code + '" AND ' +
       'E_DATECOMPTABLE = "' + UsDateTime(FExoDate.Deb) + '" AND ' +
       '(E_ECRANOUVEAU = "OAN" OR E_ECRANOUVEAU = "H")') then
    begin
      lStTemp := GetColonneSQL('EXERCICE', 'EX_ETATCPTA', 'EX_DATEFIN = "' + UsDateTime(FExoDate.Deb-1) + '"');
      if (lStTemp <> '') and (lStTemp <> 'CDE') then
        LibAno.Text := TraduireMemoire('Total des A-Nouveaux provisoires');
    end;    
  end;

  // Enregistrement de la date de début d'exercice séléctionné
  DateDebutExo.Text := DateToStr(FExoDate.Deb);

  // Date Cumul Au
  if CBEnSituation.Checked then
    CumulAu.Text := DateToStr(StrToDate(DateDebutExo.Text) - 1)
  else
    CumulAu.Text := DateToStr(StrToDate(E_DateComptable.Text) - 1);

  // Affichage en DEVISE ou en DEVISEPIVOT
  RGDevise.Items[0] := '&' + VH^.LibDevisePivot;
  if (RGDevise.Enabled) and (RGDevise.ItemIndex = 1) then
  begin // Affichage en devise
    AFFDEVISE.Text := 'X';
    DEVPIVOT.Text  := TraduireMemoire('Devise');
  end
  else
  begin
    AFFDEVISE.Text := '-';
    DEVPIVOT.Text  := VH^.LibDevisePivot;
  end;

  // Rupture Paramètrable
  if Rupture.Value <> 'SANS' then
  begin
    // GCO - 12/02/2007 - FQ 19352
    //if (RUPCYCLEREVISION <> nil) and (not RUPCYCLEREVISION.Checked) then
    //begin
      AvecRupType.Text := RuptureType.Value;
      AvecRupture.Text := Rupture.Value;
    //end;

    // Rupture sur Numéro de Compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.Value > 0) then
      AvecNivoRupture.Text := IntToStr(NivoRupture.Value);

    // Rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      XX_Rupture.Text := 'G_TABLE' + GetNumTableLibre;

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      if Corresp.value = 'GE1' then
        XX_Rupture.Text := 'G_CORRESP1'
      else
        XX_Rupture.Text := 'G_CORRESP2';
    end;
  end; // Fin Rupture Paramètrable

  (*
  // NE PAS IMPLEMENTER POUR LE MOMENT
  // GCO - 01/06/2006 - FQ 13051
  EtatIncomplet.Text := '-'; // Par défaut

  // Vérification des généraux
  if ExisteSQL('SELECT G_GENERAL FROM GENERAUX WHERE ' +
               'G_GENERAL >= "' + E_General.Text + '" AND ' +
               'G_GENERAL <= "' + E_General_.Text + '" AND ' +
               'G_CONFIDENTIEL > "' + V_PGI.Confidentiel + '"') then
    EtatIncomplet.Text := 'X';

  // Si les généraux ne sont pas confidentiels, mais que c'est un GLGENEPARAUXI
  if (EtatIncomplet.Text = '-') and (FBoGeneParAuxi) then
  begin //Vérification des auxiliaires
    if ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE ' +
                 'T_AUXILIAIRE >= "' + E_Auxiliaire.Text + '" AND ' +
                 'T_AUXILIAIRE <= "' + E_Auxiliaire_.Text + '" AND ' +
                 'T_CONFIDENTIEL > "' + V_PGI.Confidentiel + '"') then
      EtatInComplet.Text := 'X';
  end;
  *)
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 04/04/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnLoad;
begin
  inherited;

  // Remplissage des Zones GENERAL ET AUXILIAIRE et Vérification des dates
  ControleAvantEdition;

  // Gestion des critères DEV
  GestionCritereDEV;

  FDateFinEdition := StrToDate(E_DateComptable_.Text);

  // Création de la requête des centralisations et insertion dans la table ECRITURE
  GestionCentralisation;

  // Création de la condition Where de l'état
  FStWhereEtat := GenereConditionWhereEtat;

  FStWhereEtat := CMajRequeteExercice (E_EXERCICE.Value, FStWhereEtat);

  // Calcul des Cumuls ANO et Cumuls Au de la bande §RG6 de l'état
  GenereCumulGL;

{$IFDEF EAGLCLIENT}

{$ELSE}
  // GCO - 20/12/2005 - Correction PB BDE si trop d'ecritures à imprimer 900000
  // GCO - 03/01/2007 - FQ 19375
  if FBoOkDecoupage then
    GestionTailleGrandLivre( FStWhereEtat );
{$ENDIF}

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/02/2004
Modifié le ... : 18/02/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnNew;
begin
  inherited;
  SupprimeCentralisation;

  // ----------------- Init des composants si pas de Filtre DEFAUT -------------
  if FFiltres.Text = '' then
  begin
    G_NatureGene.ItemIndex    := 0;
    ModeSelect.Value          := 'FOU';
    E_QualifPiece.Text        := 'N;';
    {JP 04/10/05 : FQ 16149 : En cours !! En attendant de plus amples informations
                   sur le fonctionnement des établissements !!!}
    if (ComboEtab.Enabled) and (ComboEtab.ItemIndex < 0) then
      ComboEtab.ItemIndex := 0;

    E_Devise.ItemIndex := 0;
  end;

  OnChangeE_Devise( nil );
  OnClickCBEnSituation( nil );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... : 09/12/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnClose;
begin
  if Assigned(FTobCumulAnoGL)   then FreeAndNil(FTobCumulAnoGL);
  if Assigned(FTobCumulAuGL)    then FreeAndNil(FTobCumulAuGL);
  if Assigned(FTobBilanGestion) then FreeAndNil(FTobBilanGestion);
  SupprimeCentralisation;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.OnExitE_DateComptable(Sender: TObject);
var lExoDate : TExoDate;
begin
  // GCO - 26/10/2005 - FQ 14967
  if E_Exercice.ItemIndex <> 0 then
  begin
    lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));
    AFFANOUVEAU.Enabled := StrToDate(E_DateComptable.Text) = lExoDate.Deb;
  end
  else
  begin
    if VH^.ExoV8.Code <> '' then
      AFFANOUVEAU.Enabled := E_DateComptable.Text = DateToStr(VH^.ExoV8.Deb)
    else
      AFFANOUVEAU.Enabled := E_DateComptable.Text = DateToStr(VH^.Exercices[1].Deb);
  end;

  if not AFFANouveau.Enabled then
    AffANouveau.Checked := False;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPGLGENE.OnChangeE_Exercice(Sender: TObject);
begin
  if E_Exercice.ItemIndex = 0 then
  begin
    AFFANouveau.Checked := False;
    if VH^.ExoV8.Code <> '' then
      E_DateComptable.Text := DateToStr(VH^.ExoV8.Deb)
    else
      // Date de début du premier Exercice
      E_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
    E_DateComptable_.Text  := '30/12/2099';
  end
  else
  begin
    CExoRelatifToDates ( E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
    if StrToDate(E_DateComptable.Text) < VH^.ExoV8.Deb then
    begin
      AFFANouveau.Checked   := False;
      AFFANouveau.Enabled   := False;
      CBENSITUATION.Checked := False;
      CBENSITUATION.Enabled := False;
    end
    else
    begin
      if not FBoGeneParAuxi then
      begin
        AFFANouveau.Enabled   := True;
        CBENSITUATION.Enabled := True;
      end;  
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Changement de la Devise et
Mots clefs ... :
*****************************************************************}

procedure TOF_CPGLGENE.OnChangeE_Devise(Sender: TObject);
begin
  RGDevise.Enabled := (E_Devise.ItemIndex <> 0) and (E_Devise.Value <> V_PGI.DevisePivot);
  if E_Devise.ItemIndex = 0 then
    RGDevise.ItemIndex := 0;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/12/2002
Modifié le ... :   /  /
Description .. : Selection des écritures en situation
Suite ........ : et Activation la zone de saisie de la date
Mots clefs ... :
*****************************************************************}

procedure TOF_CPGLGENE.OnClickCBEnSituation(Sender: TObject);
begin
  EDateSituation.Enabled   := CBEnSituation.Checked;
  CBTous.Enabled           := CBEnSituation.Checked;
  EDateSituation.Color     := IIF(CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_DateComptable.Enabled  := not CBEnSituation.Checked;
  E_DateComptable_.Enabled := not CBEnSituation.Checked;
  E_DateComptable.Color    := IIF(not CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_DateComptable_.Color   := IIF(not CBEnSituation.Checked, ClWindow, ClBtnFace);
  E_QUALIFPIECE.Enabled    := not CBEnSituation.Checked;

  // GCO - 10/11/2004 - FQ 13498 et 14865 
  //AFFCumul.Checked := not CBEnSituation.Checked;
  //AFFCumul.Enabled := not CBEnSituation.Checked;

  if CBEnSituation.Checked then
  begin
    E_Exercice.ItemIndex := 0;
    E_Exercice.Enabled   := False;
    E_Exercice.OnChange(Self);
    AFFANouveau.Enabled  := False;
    AFFANouveau.Checked  := True;
    //AFFCentra.Enabled    := False;
    //AFFCentra.Checked    := False;
    E_NumeroPiece.Text   := '';
    E_NumeroPiece_.Text  := '';
    E_QualifPiece.Text   := 'N;'
  end
  else
  begin
    E_Exercice.Enabled   := True;
    AFFANouveau.Enabled  := True;
    AFFANouveau.Checked  := False;
    //AFFCentra.Enabled    := True;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPGLGENE.OnClickRupture(Sender: TObject);
begin
  Ruptures.TabVisible := (Rupture.ItemIndex <> 0);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPGLGENE.OnClickRuptureType(Sender: TObject);
begin
  RupGroupes.Visible := (RuptureType.ItemIndex = 0);
  RupLibres.Visible  := (RuptureType.ItemIndex = 1);
  RupCorresp.Visible := (RuptureType.ItemIndex = 2);

  if RuptureType.ItemIndex = 1 then
  begin
    TableLibre.ItemIndex := 0;
    OnChangeTableLibre(nil);
  end;

  if RuptureType.ItemIndex = 2 then
  begin
    if Corresp.Items.Count >= 1 then
      Corresp.ItemIndex := 1
    else
      Corresp.ItemIndex := 0;

    OnChangeCorresp(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.FTimerTimer(Sender: TObject);
begin
  if FCritEdtChaine <> nil then
  begin
    // GCO - 03/01/2007 - FQ 19375
    FBoOkDecoupage := False;

    with FCritEdtChaine do
    begin
      if CritEdtChaine.UtiliseCritStd then
      begin
        E_Exercice.Value      := CritEdtChaine.Exercice.Code;
        E_DateComptable.Text  := DateToStr(CritEdtChaine.Exercice.Deb);
        E_DateComptable_.Text := DateToStr(CritEdtChaine.Exercice.Fin);

        // GCO - 03/08/2007 - FQ 19377
        G_NatureGene.Value    := CritEdtChaine.NatureCompte;
        ModeSelect.Value      := CritEdtChaine.ModeSelection;
        E_QualifPiece.Value   := CritEdtChaine.TypeEcriture;
      end;
    end;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : Restreint la tablette des comptes en fonction de la nature
Suite ........ : de compte sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnChangeG_NATUREGENE(Sender: TObject);
begin
  AffectePlusE_General;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLGENE.RecupAutresCriteres : string;
begin
  // Borne Max de E_DATECOMPTABLE
  if not CbEnSituation.Checked then
    Result := ' AND E_DATECOMPTABLE <="' + USDateTime(StrToDate(E_DateComptable_.Text)) + '"' ;

  // E_Exercice, E_DATECOMPTABLE ne peuvent etre pris par le RecupWhereCriteres
  //car ne doit pas intervenir en date de Situation,
  if not CbEnsituation.Checked then
    Result := Result + ' AND E_DATECOMPTABLE >="' + USDateTime(StrToDate(E_DateComptable.Text)) + '"'
  else
    Result := Result + ' AND E_DATECOMPTABLE >="' + USDateTime(StrToDate(DateDebutExo.Text)) + '"';

  if (not CBEnSituation.Checked) then
  begin
    if E_Exercice.ItemIndex <> 0 then // Tous
      Result := Result + ' AND E_EXERCICE ="' + cRelatifVersExercice(E_Exercice.Value) + '"';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/09/2003
Modifié le ... :   /  /
Description .. : Récupère les critères communs à la condition WHERE du grand Livre
Mots clefs ... :
*****************************************************************}
function TOF_CPGLGENE.RecupCriteresCommuns: string;
var lStListeCptExcept    : string;
    lStCptExcept         : string;
    lStValEQualifPiece   : string;
    lStLibEQualifPiece   : string;
    lStTemp              : string;
begin
  Result := ' AND ' + ConvertitCaractereJokers(E_GENERAL, E_GENERAL_, 'E_GENERAL');

  // --> Compte Auxiliaire si Grand Livre Auxiliaire par Général
  if FBoGeneParAuxi then
    Result := Result + ' AND ' + ConvertitCaractereJokers(E_Auxiliaire, E_Auxiliaire_, 'E_AUXILIAIRE');

  // GCO - 11/07/2006 - FQ 18363  
  lStTemp := RecupWhereCritere(Pages);
  if lStTemp <> '' then
    lStTemp := Copy(lStTemp, 6, length(lStTemp)); // Dégage le mot WHERE

  Result := Result + ' AND ' + lStTemp;

  // Prise en compte du VH^.ExoV8 afin de cacher les ecritures antérieures
  // uniquement en date de situation
  if (CBEnSituation.Checked or AFFANOUVEAU.Checked) and (VH^.ExoV8.Code <> '') then
  begin
    Result := Result + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '" ';
  end;

  if CBEnSituation.Checked then
  begin
    if (CBTous.Enabled) and (not CBTous.Checked) then // Justif de solde des comptes lettrables uniquement
      Result := Result + ' AND G_LETTRABLE = "X"';
  end;

  // Sélection du type d'écritures et  Traduction du E_QualifPiece
  TraductionTHMultiValComboBox( E_QualifPiece, lStValEQualifPiece, lStLibEQualifPiece, 'E_QUALIFPIECE', False, '"Z"');

  LibEQualifPiece.Text := lStLibEQualifPiece;

  if Trim(lStValEqualifPiece) <> '' then
    Result := Result + ' AND ' + lStValEQualifPiece;

  // Ecritures validées
  if ECRValide.State <> CbGrayed then
  begin
    if ECRValide.Checked then
      Result := Result + ' AND E_VALIDE = "X" '
    else
      Result := Result + ' AND E_VALIDE = "-" ';
  end;

  // Ecritures Lettrés
  if ECRLettre.State <> CbGrayed then
  begin
    if ECRLettre.Checked then
      Result := Result + ' AND (E_ETATLETTRAGE = "TL" OR E_ETATLETTRAGE = "PL") '
    else
      Result := Result + ' AND (E_ETATLETTRAGE = "RI" OR E_ETATLETTRAGE = "AL" ) ';
  end;

  // Ecritures Pointées
  if ECRPointe.State <> CbGrayed then
  begin
    if ECRPointe.Checked then
      Result := Result + ' AND E_REFPOINTAGE <> "" '
    else
      Result := Result + ' AND E_REFPOINTAGE = "" ';
  end;

  // (Débit <> 0 ou Crédit <> 0) et (Ecritures non détruites)
  Result := Result + ' AND (E_DEBIT <> 0 OR E_CREDIT <> 0) AND E_CREERPAR <> "DET"' ;

  if FBoGeneParQte then
  begin
    if TypeQuantite.Value = 'QTE1' then
      Result := Result + ' AND E_QTE1 <> 0'
    else
      if TypeQuantite.Value = 'QTE2' then
        Result := Result + ' AND E_QTE2 <> 0';
  end;

  if (RupCycleRevision <> nil) and
     (RupCycleRevision.Checked) and
     (G_CycleRevision.Text <> '')then
  begin
    if TestJoker(G_CycleRevision.Text) then // GCO - 14/09/2007 - FQ 20913
      Result := Result + ' AND ' + ClauseAvecJoker(G_CycleRevision.Text, 'G_CYCLEREVISION')
    else
      Result := Result + ' AND G_CYCLEREVISION LIKE "' + G_CycleRevision.Text + '%"';
  end;

  // Comptes d'exceptions (les comptes peuvent être séparés par des ',' ou ';')
  if CptExcept.Text <> '' then
  begin
    lStListeCptExcept := FindEtReplace(CptExcept.Text, ',', ';', True);
    while (lStListeCptExcept <> '') do
    begin
      lStCptExcept := Trim(ReadTokenSt(lStListeCptExcept));
      if lStCptExcept <> '' then
        Result := Result + ' AND E_GENERAL NOT LIKE "' + lStCptExcept + '%"';
    end;
  end;

  // Gestion du V_Pgi.Confidentiel
  Result := Result + ' AND ' + CGenereSQLConfidentiel('G');
  if FBoGeneParAuxi then
    Result := Result + ' AND ' + CGenereSQLConfidentiel('T');

  // Rupture Paramètrable
  if Rupture.Value <> 'SANS' then
  begin
    // Rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      if CPTLibreExist.Checked then
      begin
        // GCO - 04/09/2006 - FQ 18432
        Result := Result + ' AND G_TABLE' + GetNumTableLibre + ' <> ''';

        if Trim(LibreDe.Text) <> '' then
          Result := Result + ' AND G_TABLE' + GetNumTableLibre + '>= "' + LibreDe.Text + '"';

        if Trim(LibreA.Text) <> '' then
          Result := Result + ' AND G_TABLE' + GetNumTableLibre + '<= "' + LibreA.Text + '"';
      end;
    end;

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
    begin
      if CPTCorrespExist.Checked then
      begin
        if Corresp.Value = 'GE1' then
        begin
          // GCO - 04/09/2006 - FQ 18432
          Result := Result + ' AND G_CORRESP1 <> ''';

          if Trim(CorrespDe.Text) <> '' then
            Result := Result + ' AND G_CORRESP1 >= "' + CorrespDe.Text + '"';

          if Trim(CorrespA.Text) <> '' then
            Result := Result + ' AND G_CORRESP1 <= "' + CorrespA.Text + '"';
        end
        else
        begin
          // GCO - 04/09/2006 - FQ 18432
          Result := Result + ' AND G_CORRESP2 <> ''';

          if Trim(CorrespDe.Text) <> '' then
            Result := Result + ' AND G_CORRESP2 >= "' + CorrespDe.Text + '"';

          if Trim(CorrespA.Text) <> '' then
            Result := Result + ' AND G_CORRESP2 <= "' + CorrespA.Text + '"';
        end;
      end;
    end;
  end; // Fin Rupture Paramètrable
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/08/2003
Modifié le ... :   /  /
Description .. : Changement de la rupture sur les plans de correspondance
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnChangeCorresp(Sender: TObject);
begin
  if RuptureType.value <> 'RUPCORRESP' then Exit;
  CorrespToCodes(Corresp,TComboBox(CorrespDe),TComboBox(CorrespA));
  CorrespDe.ItemIndex := 0 ;
  CorrespA.ItemIndex  := CorrespA.Items.Count - 1 ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/08/2003
Modifié le ... :   /  /    
Description .. : Changement de la rupture sur les Tables Libres
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.OnChangeTableLibre(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
  begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
  end
  else
  begin
    LibreDe.DataType := 'TZNATGENE' + GetNumTableLibre;
    LibreA.DataType  := 'TZNATGENE' + GetNumTableLibre;
    //LibreDe.Text     := '';
    //LibreA.Text      := '';
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2003
Modifié le ... : 12/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnExitE_General(Sender: TObject);
begin
  // GCO - 04/09/2007 - FQ 21324
  if (csDestroying in Ecran.ComponentState) then Exit;

  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + BourreEtLess(THEdit(Sender).Text, fbGene) + '"') then
  begin
    if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
      THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
  end
  else
    THEdit(Sender).ElipsisClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/09/2003
Modifié le ... : 12/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnExitE_Auxiliaire(Sender: TObject);
begin
  if (csDestroying in Ecran.ComponentState) then Exit;

  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
  begin
    if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
      THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
  end
  else
    THEdit(Sender).ElipsisClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.GestionCentralisation;
begin
  SupprimeCentralisation;
  if AFFCentra.Checked then
  begin
    if Transactions( EffectueCentralisation, 1 ) <> OeOk then
    begin
      PgiInfo('Erreur lors de la centralisation', Ecran.Caption);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/10/2003
Modifié le ... :   /  /
Description .. : Supprime les écritures de centralisation générées par les
Suite ........ : grand livres dont le E_Qualifpiece="Z"
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.SupprimeCentralisation;
begin
  try
    if CEstBloqueGLGene then
    begin
      BeginTrans;
      ExecuteSQL('DELETE FROM ECRITURE WHERE E_QUALIFPIECE="Z" AND E_UTILISATEUR = "' + V_PGI.User + '"');
      CDeBlocageGLGene; // Suppression du verrou dans la Table COURRIER
      CommitTrans;
    end;
  except
    on E: Exception do
    begin
      PgiError('Impossible de supprimer les enregistrements temporaires du Grand Livre', Ecran.Caption);
      RollBack;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/05/2004
Modifié le ... :   /  /
Description .. : Effectue la centralisation des écritures dans la TABLE ECRITURES
suite ........ : en créant des mouvements typés Z dans E_QUALIFPIECE
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.EffectueCentralisation;
var
  lQuery         : TQuery;
  lStSql         : string;
  lSt            : string;
  //lStTemp        : string;
  lStPeriode     : string;
  lDateFinDeMois : TDateTime;
  lNumLigne      : Integer;
  lNumeroPiece   : Integer;
begin
  try
    try
      lQuery := OpenSql('SELECT MAX(E_NUMEROPIECE) LENUM FROM ECRITURE WHERE E_QUALIFPIECE = "Z"', True);
      if lQuery.FindField('LENUM').AsInteger <> 0 then
        lNumeroPiece := lQuery.FindField('LENUM').AsInteger
      else
        lNumeroPiece := 1;
      Ferme(lQuery);

      // Requete des écritures centralisées
      lStSql := 'SELECT DISTINCT E_GENERAL, ' + IIF(FBoGeneParAuxi, 'E_AUXILIAIRE, ', '') +
                'E_EXERCICE, E_PERIODE, E_JOURNAL, E_DEVISE, E_ETABLISSEMENT, ' +
                'SUM(E_DEBIT) DEBIT, SUM(E_CREDIT) CREDIT, ' +
                'SUM(E_DEBITDEV) DEBITDEV, SUM(E_CREDITDEV) CREDITDEV FROM ECRITURE ' +
                'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
                'LEFT JOIN TIERS T ON T_AUXILIAIRE = E_AUXILIAIRE ' +
                'LEFT JOIN JOURNAL  ON J_JOURNAL = E_JOURNAL ' +
                'WHERE J_CENTRALISABLE = "X" AND G_CENTRALISABLE = "X"';

      if CBEnSituation.Checked then
      begin // On ne centralise pas les LETTRABLES car ils seront en justificatif de solde
        lStSql := lStSql + ' AND G_LETTRABLE = "-" ' +
        // GCO - 05/06/2007 - FQ 19482 - Pas Centralisation >= à la date de justif de solde 
                  'AND E_DATECOMPTABLE <="' + USDateTime(StrToDate(EDatesituation.Text)) + '"';
      end;

      lStSql := lStSql + ' '+ RecupAutresCriteres + RecupCriteresCommuns + ' GROUP BY E_GENERAL, ' +
                IIF(FBoGeneParAuxi, 'E_AUXILIAIRE, ', '') +
                'E_EXERCICE, E_PERIODE, E_JOURNAL, E_DEVISE, E_ETABLISSEMENT ' +
                'ORDER BY E_GENERAL, E_EXERCICE, E_PERIODE, E_JOURNAL';

      lQuery := OpenSql(lStSql, True);
      if not lQuery.Eof then
      begin
        // Création d'un verrou en Table COURRIER pour la présence de
        // Centralisation Mensuelle dans la Table ECRITURE
        CBlocageGLGene;
        InitMoveProgressForm(Ecran, Ecran.Caption, 'Traitement des centralisations en cours...', RecordsCount(lQuery), True, True);

        lNumLigne := 1;
        while not lQuery.Eof do
        begin
          if not MoveCurProgressForm('Compte ' + lQuery.FindField('E_GENERAL').AsString + ' - ' + 'Centralisation Mensuelle') then
            Break;

          lStPeriode := IntToStr(lQuery.FindField('E_PERIODE').AsInteger);
          lDateFinDeMois := StrToDate('01/' + Copy(lStPeriode, 5, 2) + '/' +  Copy(lStPeriode, 1, 4));
          lDateFinDeMois := FinDeMois(ldateFinDeMois);

          // gco - 05/06/2007 - FQ 19482
          if CBEnSituation.Checked and (lDateFinDeMois > StrToDate(EDatesituation.Text)) then
            lDateFinDeMois := StrToDate(EDateSituation.Text);

          if FBoGeneParAuxi then
            lSt := 'INSERT INTO ECRITURE (E_UTILISATEUR, E_GENERAL, E_AUXILIAIRE, '
          else
            lSt := 'INSERT INTO ECRITURE (E_UTILISATEUR, E_GENERAL, ';

          lSt := lSt + 'E_EXERCICE, E_JOURNAL, ' +
                 'E_DATECOMPTABLE, E_PERIODE, E_QUALIFPIECE, E_LIBELLE, E_DATEMODIF, ' +
                 'E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_ECRANOUVEAU,' +
                 'E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, E_CONFIDENTIEL, E_CREERPAR, ' +
                 'E_QTE1, E_QTE2, E_DATEECHEANCE, E_DEVISE, E_ETABLISSEMENT, E_ETATLETTRAGE) ' +
                 'VALUES (' +
                 '"' + string(V_PGI.User) + '",' +
                 '"' + lQuery.FindField('E_GENERAL').AsString + '",';

          if FBoGeneParAuxi then
            lSt := lSt + '"' + lQuery.FindField('E_AUXILIAIRE').AsString + '",';

          lSt := lSt + '"' + lQuery.FindField('E_EXERCICE').AsString + '",' +
                 '"' + lQuery.FindField('E_JOURNAL').AsString + '",' +
                 '"' + UsDateTime(lDateFinDeMois) + '",' +
                 '"' + lQuery.FindField('E_PERIODE').AsString + '",' +
                 '"Z", "Centralisation mensuelle",' +
                 '"' + UsDateTime(lDateFinDeMois) + '",' +
                 IntToStr( lNumeroPiece ) + ',' + IntToStr(lNumLigne) + ', 0, "N",' +
                 VariantToSql(lQuery.FindField('DEBIT').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('CREDIT').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('DEBITDEV').AsFloat) + ',' +
                 VariantToSql(lQuery.FindField('CREDITDEV').AsFloat) + ',' +
                 '"0", "SAI", ' + // E_CONFIDENTIEL A ZERO, E_CREERPAR = SAI
                 '"0", "0", "' + UsDateTime( iDate1900 ) + '",' +
                 '"' + lQuery.FindField('E_DEVISE').AsString + '",' +
                 '"' + lQuery.FindField('E_ETABLISSEMENT').AsString + '", "RI")';

          ExecuteSql(lSt);
          Inc(lNumLigne);
          lQuery.Next;
        end;
      end;
    except
      on E: Exception do PgiError(E.Message, Ecran.Caption);
    end;
  finally
    Ferme(lQuery);
    FiniMoveProgressForm;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLGENE.CEstBloqueGLGene: Boolean;
var lStSql  : string ;
    lQuery  : TQuery ;
begin
  Result := False;

  lStSql := 'SELECT MG_EXPEDITEUR FROM COURRIER WHERE ' +
            'MG_UTILISATEUR = "' + W_W + '" AND MG_TYPE=4000 AND ' +
            'MG_EXPEDITEUR = "' + String(V_PGI.User) + '"';

  lQuery := nil;
  try
    try
      lQuery := OpenSql(lStSql,True) ;
      Result := not lQuery.Eof ;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CEstBloqueGLGene');
    end;
  finally
    Ferme(lQuery);
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.CBlocageGLGene;
begin
  try
    if not CEstBloqueGLGene then
    begin
      ExecuteSQL('INSERT INTO COURRIER(MG_UTILISATEUR, MG_TYPE, MG_LIBELLE, '+
                 'MG_DATE, MG_EXPEDITEUR) VALUES("' + W_W + '",4000, ' +
                 '"Centralisation Mensuelle","'+USTime(Now)+'","' + V_PGI.User + '")');
    end;
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CBlocageGLGene');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/04/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.CDeBlocageGLGene;
begin
  try
    ExecuteSQL('DELETE FROM COURRIER WHERE MG_UTILISATEUR = "' + W_W + '" ' +
               'AND MG_TYPE=4000 ' +
               'AND MG_EXPEDITEUR = "' + String(V_PGI.User) + '"');
  except
    on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CDeBlocageGLGene');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.OnChangeFiltre(Sender: TObject);
begin
  inherited;
  CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_, True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 12/10/2004
Modifié le ... :   /  /
Description .. : Afin de rendre apparent le focus sur la
Suite ........ : THMultiValComboBax, car en readOnly, on perd le curseur
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnEnterQualifPiece(Sender: TObject);
begin
  CSelectionTextControl( Sender ) ;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/04/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.OnClickRupCycleRevision(Sender: TObject);
begin
  G_CycleRevision.Enabled := (RupCycleRevision.Checked);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.ChargementCritEdt;
begin
  inherited;

  if (TheData <> nil) and (TheData is (ClassCritEdt)) then
  begin
    E_General.Text  := ClassCritEdt(TheData).CritEdt.Cpt1;
    E_General_.Text := ClassCritEdt(TheData).CritEdt.Cpt2;

    if FBoGeneParAuxi then
    begin
      E_Auxiliaire.Text  := ClassCritEdt(TheData).CritEdt.SCpt1;
      E_Auxiliaire_.Text := ClassCritEdt(TheData).CritEdt.SCpt2;
    end;

    E_Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);

    // Date de Début de l'édition
    E_DateComptable.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);

    // Date de Fin de l'édition
    E_DateComptable_.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);

    // Numéro de Pièce
    //JP 26/08/04 : FQ 14062 : avec l'autocomplétion, si NumPiece1 = 0 et NumPiece2 = 9999999,
    //                         on se retrouve E_NUMEROPIECE à 99999999
    // GCO - 05/09/2005 - FQ 16320
    if ClassCritEdt(TheData).CritEdt.GL.NumPiece1 > 0 then
      E_NUMEROPIECE.Text  := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece1);

    if ClassCritEdt(TheData).CritEdt.GL.NumPiece2 > 0 then
      E_NUMEROPIECE_.Text := IntToStr(ClassCritEdt(TheData).CritEdt.GL.NumPiece2);

    // En situation
    CBEnSituation.Checked := ClassCritEdt(TheData).CritEdt.GL.EnDateSituation;
    if CBEnsituation.Checked then
    begin
      Ecran.Caption := TraduireMemoire('Justificatif de solde');
      UpdateCaption(Ecran);
      EDateSituation.Text := DateToStr( ClassCritEdt(TheData).CritEdt.Date2);
    end
    else
    begin
      // Détail A-Nouveaux
      AffANouveau.Checked := ClassCritEdt(TheData).CritEdt.GL.DetailAno;
      // GCO - 26/12/2005 - FQ 11773
      AffCentra.Checked   := ClassCritEdt(TheData).CritEdt.GL.AvecDetailCentralise;
    end;
    
    // Référence Interne
    E_REFINTERNE.Text := ClassCritEdt(TheData).CritEdt.ReferenceInterne;

    // E_Qualifpièce
    E_QUALIFPIECE.Text := FindEtReplace(ClassCritEdt(TheData).CritEdt.QualifPiece, '!', ';', True);

    // Etablissement
    ComboEtab.Value := ClassCritEdt(TheData).CritEdt.Etab;

    // Devise
    E_DEVISE.Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;

    // Ecritures validées ( '' = Grayed, OUI = Checked sinon not Checked
    if (ClassCritEdt(TheData).CritEdt.Valide = '') then
      ECRVALIDE.State := cbGrayed
    else
      ECRVALIDE.Checked := (ClassCritEdt(TheData).CritEdt.Valide = 'OUI');

    // Ecritures Lettrées ( 0 = Grayed, 1 = not Checked sinon 2 = Checked
    if (ClassCritEdt(TheData).CritEdt.GL.Lettrable = 0) then
      ECRLETTRE.State   := cbGrayed
    else
      ECRLETTRE.Checked := (ClassCritEdt(TheData).CritEdt.GL.Lettrable = 2);

    // Cycle de révision pour le GL Général uniquement
    if (ClassCritEdt(TheData).CritEdt.GL.CycleDeRevision <> '') and
       (RupCycleRevision <> nil) and
       (G_CycleRevision <> nil) then
    begin
      Rupture.ItemIndex := 1;
      RupCycleRevision.Checked := True;
      G_CycleRevision.Text := ClassCritEdt(TheData).CritEdt.GL.CycleDeRevision;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPGLGENE.MySelectFiltre;
var lExoDate : TExoDate;
begin
  inherited;
  lExoDate := CtxExercice.QuelExoDate(CRelatifVersExercice(E_Exercice.Value));

  if (CtxExercice.QuelExo(E_DateComptable.Text, False) <> lExoDate.Code) and
     (CtxExercice.QuelExo(E_DateComptable_.Text, False) <> lExoDate.Code) then
  begin
    CExoRelatifToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  end;

  // GCO - 30/07/2007 - FQ 21180
  if E_QUALIFPIECE.Value = '' then
  begin
    E_Qualifpiece.SelectAll;
    E_Qualifpiece.Text := TraduireMemoire('<<Tous>>');
  end;

  (*
  // SBO - 20/08/2007 - FQ 21250 ( violation d'accès en création de filtre )
  if (G_CycleRevision<>nil) and (G_CycleRevision.Value = '') then
  begin
    G_CycleRevision.SelectAll;
    G_CycleRevision.Text := TraduireMemoire('<<Tous>>');
  end;*)
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.AffectePlusE_General;
begin
  E_General.Text := '';
  E_General_.Text := '';
  E_General.Plus := '';

  if FBoGeneParAuxi then
  begin
    E_Auxiliaire.Text  := '';
    E_Auxiliaire_.Text := '';
    E_General.Plus  := ' G_COLLECTIF = "X"';
  end;

  if G_NatureGene.ItemIndex > 0 then
  begin
    if E_General.Plus = '' then
      E_General.Plus := ' G_NATUREGENE = "' + G_NatureGene.Value + '"'
    else
      E_General.Plus := E_General.Plus + ' AND G_NATUREGENE = "' + G_NatureGene.Value + '"';
  end;

  E_General_.Plus := E_General.Plus;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPGLGENE.AuxiElipsisClick(Sender: TObject);
begin
{$IFNDEF EDTQR}
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TOF_CPGLGENE.GestionTailleGrandLivre( vStwhereCritere : string );
var lStSelect   : string;          //
    lStGroupBy  : string;          //
    lStMini     : string;          //
    lStCompte   : string;          // Fourchette de comte que l'on ajoute au Where
    lStFileName : string;          // Nom du fichier temp du PDFBatch
    lBoOldNoPrintDialog : Boolean; //
    lBoOldQRPdf         : Boolean;
    lBoAbort            : Boolean;
    lNumPage, lNbEcr, i : integer; //
    lTotalPage : integer;
    lNbEcrMax           : integer; // Nombre d'écritures MAX avant de couper le GL
    lTobTrancheCompte   : Tob;     // Liste des fourchettes de Compte
    lTobDecoupage       : Tob;     // Nb d'écritures à imprimer par Compte
    lTobTemp            : Tob;     // Tob de Traitement
// GP 23152
    Okliste              : Boolean ;
begin
  TFQRS1(Ecran).OnValideCritereEvent := nil; // Conditionne le Critok du TQRS1
// GP 23152
  OkListe:=GetControlTExt('FLISTE')='X' ;

  lBoAbort := False;

  lTobTrancheCompte := Tob.Create('', nil, -1);

  lStSelect  := 'SELECT E_GENERAL';
  lStGroupBy := ' GROUP BY E_GENERAL';

  // GCO - 27/10/2006 - FQ 18923
  //if FBoGeneParAuxi then
  //begin
  //  lStSelect := lStSelect + ', E_AUXILIAIRE';
  //  lStGroupBy := lStGroupBy + ', E_AUXILIAIRE';
  //end;

  lStSelect := lStSelect + ', COUNT(*) TOTAL FROM ECRITURE ' +
               'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' +
               'LEFT JOIN TIERS ON T_AUXILIAIRE = E_AUXILIAIRE ' +
               'LEFT JOIN JOURNAL ON J_JOURNAL  = E_JOURNAL WHERE ';

  lTobDecoupage := Tob.Create('', nil, -1);
  lTobDecoupage.LoadDetailFromSQL( lStSelect + vStWhereCritere + lStGroupBy + ' ORDER BY E_GENERAL');

  // GCO - 26/12/2005 - Découpage possible si au moins deux comptes édités
  if lTobDecoupage.Detail.Count > 1 then
  begin
    lNbEcr := 0;
    If OkListe Then lNbEcrMax := GetSynRegKey('CPGLENREGMAX', cNbEcrMax+40000, True)
               Else lNbEcrMax := GetSynRegKey('CPGLENREGMAX', cNbEcrMax, True);

    lStMini := E_General.Text;
    for i := 0 to lTobDecoupage.Detail.Count-1 do
    begin
      lNbEcr := lNbEcr + lTobDecoupage.Detail[i].GetInteger('TOTAL');
      if lNbEcr >= lNbEcrMax then
      begin
        lNbEcr := 0;
        lTobTemp := Tob.Create('', lTobTrancheCompte, -1);
        lTobTemp.AddChampSupValeur('MINI', lStMini);
        lTobTemp.AddChampSupValeur('MAXI', lTobDecoupage.Detail[i].GetString('E_GENERAL'));
        if i <> lTobDecoupage.Detail.Count-1 then
          lStMini := lTobDecoupage.Detail[i+1].GetString('E_GENERAL')
        else
          lStMini := E_General_.Text; // Dernière ligne
      end;
    end;

    // Faut Faire la dernière tranche
    if lTobTrancheCompte.Detail.Count <> 0 then
    begin
      lTobTemp := Tob.Create('', lTobTrancheCompte, -1);
      lTobTemp.AddChampSupValeur('MINI', lStMini);
      lTobTemp.AddChampSupValeur('MAXI', E_General_.Text);
    end;

    // Edition découpée de l'état
    if lTobTrancheCompte.Detail.Count <> 0 then
    begin
      lBoOldNoPrintDialog := V_Pgi.NoPrintDialog;
      lBoOldQRPdf         := V_Pgi.QRPdf;

      V_Pgi.QRPdf := True;

      // CritOk devient FALSE, Plus de OnUpdate pour le QRS1
      TFQRS1(Ecran).OnValideCritereEvent := OnMyValideCritereEvent;

      lStFileName := TempFileName();

      If Not OkListe Then
        BEGIN
    {$IFDEF EAGLCLIENT}
      StartPdfBatch(lStFileName);
    {$ELSE}
      THPrintBatch.StartPdfBatch(lStFileName);
    {$ENDIF}
        END ;

      // Changement de la CLAUSE WHERE de l'état
      lTotalPage := 0;
      for i := 0 to lTobTrancheCompte.Detail.count -1 do
      begin
        if not lBoAbort then
        begin
          If Not OkListe Then OkEntete.Text := IIF(i = 0, 'X', '-')
                         Else OkEntete.Text := IIF(i = 0, 'X', 'X') ;
          OkResume.Text := IIF(i = lTobTrancheCompte.Detail.count -1, 'X', '-');
          lStCompte := ' AND E_GENERAL >= "' + lTobTrancheCompte.Detail[i].GetString('MINI') + '" AND ' +
                       'E_GENERAL <= "' + lTobTrancheCompte.Detail[i].GetString('MAXI') + '"';


        {$IFDEF EAGLCLIENT}
          lNumPage := LanceEtatTob('E', TFQRS1(Ecran).NatureEtat, TFQRS1(Ecran).CodeEtat, nil, CBPreview.Checked, False, False, Pages, FStWhereEtat + lStCompte + RenvoiOrderBy, Ecran.caption, False,0, '', lTotalPage);
        {$ELSE}
          lNumPage := LanceEtatTob('E', TFQRS1(Ecran).NatureEtat, TFQRS1(Ecran).CodeEtat, nil, CBPreview.Checked, OkListe, False, Pages, FStWhereEtat + lStCompte + RenvoiOrderBy, Ecran.caption, False,0,'',nil, lTotalPage);
        {$ENDIF}
          // GCO - 13/09/2006 - FQ 18667
          lTotalPage := lTotalPage + lNumPage;

          if lNumPage = -1 then
            lBoAbort := True;

          if not CBPreview.Checked then
            V_PGI.NoPrintDialog := True;

          If Not OkListe Then
            BEGIN
              THPrintBatch.AjoutPdf(V_PGI.QrPdfQueue, True);
            END ;
        end;
      end;

      If Not OkListe Then
        BEGIN
        {$IFDEF EAGLCLIENT}
          CancelPDFBatch;
        {$ELSE}
          THPrintBatch.StopPdfBatch();
        {$ENDIF}
          if CbPreview.Checked then
            PreviewPDFFile(Ecran.Caption, lStFileName, True)
          else
            PrintPDF(lStFileName, Printer.Printers[Printer.PrinterIndex],'');
         END ;


      // Remise en place du Contexte
      V_Pgi.NoPrintDialog := lBoOldNoPrintDialog;
      V_Pgi.QRPdf := lBoOldQRPdf;
    end;
  end;

  FreeAndNil(lTobDecoupage);
  FreeAndNil(lTobTrancheCompte);
end;
{$ENDIF}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function OnMyValideCritereEvent : Boolean;
begin
  Result := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/12/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPGLGENE.RenvoiOrderBy: String;
begin
  Result := ' ORDER BY ';

  if (RupCycleRevision <> nil) and (RupCycleRevision.Checked) then
    Result := Result + ' G_CYCLEREVISION,';

  if not ((Rupture.ItemIndex = 0) or
         ((Rupture.ItemIndex <> 0) and (RuptureType.ItemIndex = 0))) then
  begin
    Result := Result + XX_Rupture.Text + ', ';
  end;

  Result := Result + 'E_GENERAL,' + IIF( FBoGeneParAuxi, 'E_AUXILIAIRE,', '');

  // Problème SIC, on trie sur l'index E_EXERCICE, E_DATECOMPTABLE dès qu'un
  // exercice est sélectionné.
  // GCO - 25/05/2005 - FQ 15707 - Pas de tri sur exercice si un exo selectionné
  // et A-Nouveaux détaillés.
  if (E_Exercice.ItemIndex = 0) or
     ((E_Exercice.ItemIndex <> 0) and (AFFAnouveau.Checked)) then
    Result := Result + 'E_DATECOMPTABLE,E_EXERCICE,'
  else
    Result := Result + 'E_EXERCICE,E_DATECOMPTABLE,';

  Result := Result + 'E_PERIODE,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/12/2005
Modifié le ... :   /  /    
Description .. : La requête va se présenter sous cette forme WHERE
Suite ........ : ((1) or (2) or (3)) AND Criteres Communs
Mots clefs ... :
*****************************************************************}
function TOF_CPGLGENE.GenereConditionWhereEtat : string;
begin
  Result := '(';

  // Récupération du détail des A-Nouveaux pour les comptes Lettrables uniquement (1)
  if AFFANOUVEAU.Checked then
  begin
    Result := Result + '(';
    Result := Result + '(E_ECRANOUVEAU = "N" OR E_ECRANOUVEAU = "H") AND ' +
      'E_QUALIFPIECE = "N" AND ' + // Ajout GCO - 09/02/20004
      'G_LETTRABLE = "X" AND ' + 'E_DATECOMPTABLE <= "' + UsDateTime(FExoDate.Deb) + '" AND ' +
      '(E_ETATLETTRAGE <> "TL" OR (E_ETATLETTRAGE = "TL" AND E_DATEPAQUETMAX >= "' + UsDateTime(FExoDate.Deb) + '"';

    if CBEnSituation.Checked then
      Result := Result + ' AND E_DATEPAQUETMAX > "' + UsDateTime( StrToDate( EDateSituation.Text)) + '"))'
    else
      Result := Result + '))';

    Result := Result + ') OR ' ;
  end;
  // (1) Terminé ((Détail ANO) si coché)

  // Début de récupération des écritures A-Nouveaux PGI(2)
  // Toujours récupéré l'ANO pour forcer une rupture dans l'état
  Result := Result + '(';
  Result := Result + 'E_DATECOMPTABLE >= "' + UsDateTime(StrToDate(DateDebutExo.Text)) + '"'
                   + ' AND E_QUALIFPIECE = "N"';

  if E_Exercice.ItemIndex <> 0 then
  begin
    Result := Result + ' AND E_EXERCICE="' + CRelatifVersExercice(E_EXERCICE.Value) + '" AND ' +
                       'E_DATECOMPTABLE <= "' + UsDateTime(StrToDate(E_DateComptable.Text)) + '"'
  end;


  Result := Result + ' AND (E_ECRANOUVEAU="OAN" OR E_ECRANOUVEAU="H")';

  // Pas besoin des ANO des comptes Lettrables si détail est coché car (1) les récupèrent
  if AFFANouveau.Checked then
    Result := Result + ' AND G_LETTRABLE= "-"';

  Result := Result + ') OR ';
  // (2) Terminé ( OAN )

  // Début de Récupération des écritures concernées par les critères (3)
  Result := Result + '(E_ECRANOUVEAU = "N"';

  // Si centralisation des écritures, plus la peine de récupérer les ECR dont le
  // Compte et le journal sont centralisatibles, car celles ci sont créer une
  // seule écriture Typé E_QUALIFPIECE = "Z" ( voir EffectueCentralisation )
  if AffCentra.Checked then
  begin
    if CBEnSituation.Checked then
      Result  := Result  + ' AND ((G_CENTRALISABLE = "-" OR J_CENTRALISABLE = "-") OR G_LETTRABLE = "X")'
    else
      Result  := Result  + ' AND (G_CENTRALISABLE = "-" OR J_CENTRALISABLE = "-")';
  end;

  Result := Result  + ' ' + RecupAutresCriteres ;

  // GCO -- 11/07/2006 - FQ 18363
  //lStTemp := RecupWhereCritere(Pages);
  //if lStTemp <> '' then
  //  Result := Result  + ' AND ' + Copy(lStTemp, 6, length(lStTemp));
  
  // Ecritures en situation à une date donnée
  if CBEnSituation.Checked then
  begin
    if IsValidDate(EDateSituation.Text) then
    begin
      Result := Result + ' AND E_DATECOMPTABLE <= "' +
        UsDateTime(StrToDate(EDateSituation.Text)) + '" AND ' +
        '(E_ETATLETTRAGE <> "TL" OR ' +
        '(G_LETTRABLE = "-" OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX > "' +
        UsDateTime(StrToDate(EDateSituation.Text)) + '")))';
    end;
  end;
  // (3) Terminé

  Result := Result + ')';

  // Ajout des écritures de Centralisation dans la condition Where de la requête
  // On ne prend que les Ecritures "Z" du User, en effet plusieurs utilisateurs
  // peuvent faire un grand livre avec centralisation en même temps.
  if AFFCentra.Checked then
    Result := Result + ' OR (E_QUALIFPIECE="Z" AND E_UTILISATEUR = "' + string(V_PGI.User) + '")';

  // Fermeture dela parenthèse principale qui englobe tous les ( (1) OR (2) OR (3) )
  Result := Result + ')';

  // Ajout des critères communs à (1) or (2) (3)
  Result := Result + RecupCriteresCommuns;

  Result := Result + RecupCriteresModeSelection(Result);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 27/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPGLGENE.RecupCriteresModeSelection( vStWhereEtat : string ) : string;
begin
  Result := '';

  if ModeSelect.Value = 'FOU' then
    Exit; // Tous les Comptes

  // Mode de sélection
  if ModeSelect.Value = 'EXO' then
  begin // Comptes mouvementés sur l'exercice
    Result := Result + ' AND (E_GENERAL IN (SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE ' +
              vStWhereEtat + '))';
  end
  else
  if ModeSelect.Value = 'NSL' then
  begin // Comptes non soldés
    Result := Result + ' AND (E_GENERAL IN (SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE ' +
              vStWhereEtat + ' GROUP BY E_GENERAL HAVING SUM(E_DEBIT) <> SUM( E_CREDIT))) ';
  end
  else
  if ModeSelect.Value = 'PER' then
  begin // Comptes mouvementés sur la période
    Result := Result + ' AND (E_GENERAL IN (SELECT DISTINCT E_GENERAL FROM ECRITURE WHERE ' +
              vStWhereEtat + '))';
  end;
end;

////////////////////////////////////////////////////////////////////////////////

initialization
  registerclasses([TOF_CPGLGENE]);
end.



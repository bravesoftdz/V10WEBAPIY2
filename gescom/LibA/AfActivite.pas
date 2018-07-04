{***********UNITE*************************************************
Auteur  ...... : PL
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Ecran de Saisie d'activité
Mots clefs ... : SAISIE ACTIVITE
*****************************************************************}
unit AFActivite ;

interface


uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, HTB97, StdCtrls, HPanel, UIUtil, Hent1, Menus,
  HSysMenu, Mask, Buttons, M3FP, FactTOB, FactArticle,

{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main, HDB,
{$ENDIF}
  HStatus, hmsgbox, UTOF, UtilPGI,

  Hqry, UTOB, HFLabel, Ent1, EntGC, SaisUtil, LookUp, Math,
  AglInit, Paramsoc, clipbrd,                         
  facture, FactComm, FactUtil,
{$IFDEF BTP}
  CalcOleGenericBTP,
{$ENDIF}
  UtilArticle, AFUtilArticle, AffaireUtil,  ActiviteUtil, // ne pas dissocier la ligne
  ComCtrls, uafo_calendrier,uafo_ressource,afcalendrier,dicobtp,utilressource, GereTobInterne, utofGCPieceArtLie,
  HPop97, Spin, HRichEdt, HRichOLE, confidentaffaire, utofafjournalact, utofafactivite_suivi, FactAffaire,uEntCommun,
  TntStdCtrls, TntComCtrls, TntButtons, TntGrids, TntExtCtrls;


{==============================================================================================}
{======================================= Fonctions Générales ==================================}
{==============================================================================================}
Function AFSaisieActivite (CleAct : R_CLEACT ; Action : TActionFiche; bModale : boolean; bCritereFixe : boolean; AcbDateEntreeOK : boolean) : boolean ;
Function AFCreerActivite (tsaTypeSaisie : T_Sai; tacTypeAcces : T_Acc; sTypeActivite : string; sTiersFixe : string = ''; sRessFixe : string = '') : boolean ;
Function AFCreerActiviteModale (tsaTypeSaisie : T_Sai; tacTypeAcces : T_Acc; sTypeActivite, sRessource, sAffaire, sTiers : string; DateEntree : TDateTime; iFolio : integer = 1) : boolean ;

{==============================================================================================}
{==========================Classe calendrier surchargée pour l'activité =======================}
{==============================================================================================}
Type  TGridCalendrierActivite = Class (TGridCalendrier)
    Procedure ChargeSpecifGridCalendrier; override;
    Procedure ObjetCalendrierJour (LeJour : TDateTime; Col, Row, NumSemaine : integer); override;
    Procedure CalendrierDrawRect(Col,Row : integer; Rect : TRect; Ferie : boolean); override;
    Procedure CalendrierDblClick(Sender : TObject); override;
    Procedure CalendrierCellEnter(Sender : TObject; var ACol, ARow : Integer;  var Cancel : Boolean); override;
    procedure CalendrierSelectCell(Sender : TObject; Col, Row: Longint; var CanSelect : Boolean);  override;
    End;

{==============================================================================================}
{======================================= Classe de la Form ====================================}
{==============================================================================================}
type
  TFActivite = class(TForm)
    FindLigne: TFindDialog;
    Dock971: TDock97;
    HTitres: THMsgBox;
    BZoomArticle: THBitBtn;
    BZoomAssistant: THBitBtn;
    POPZ: TPopupMenu;
    GP_FACTUREHT: TCheckBox;
    POPY: TPopupMenu;
    CpltLigne: TMenuItem;
    BZoomAffaire: THBitBtn;
    HPanel1: THPanel;
    HPanel2: THPanel;
    GS: THGrid;
    TDescriptif: TToolWindow97;
    ACT_DESCRIPTIF: THRichEditOLE;
    Ajoutduneligne1: TMenuItem;
    Supprimeuneligne1: TMenuItem;
    Afficherlemmo1: TMenuItem;
    HLabelUniteBase: THLabel;
    panel3: THPanel;
    PEntete: THPanel;
    LabelMois: TLabel;
    LabelSemaine: TLabel;
    LabelAnnee: TLabel;
    LblIntervalle: THLabel;
    LblFolio: THLabel;
    LblTri: THLabel;
    GCalendrier: THGrid;
    ACT_ANNEE: TSpinEdit;
    ACT_MOIS: TSpinEdit;
    ACT_SEMAINE: TSpinEdit;
    ACT_FOLIO: TSpinEdit;
    HValComboBoxTri: THValComboBox;
    HMTrad: THSystemMenu;
    HActivite: THMsgBox;
    PEntete2: THPanel;
    LabelAssistant: TLabel;
    LIBELLEASSISTANT: THLabel;
    LblMois: THLabel;
    ACT_RESSOURCE: THCritMaskEdit;
    ACT_AFFAIRE: THCritMaskEdit;
    HNumEditTOTQTE: THNumEdit;
    HNumEditTOTPR: THNumEdit;
    HNumEditTOTPV: THNumEdit;
    LHNumEditTOTQTE: THLabel;
    LHNumEditTOTPR: THLabel;
    LHNumEditTOTPV: THLabel;
    N1: TMenuItem;
    N2: TMenuItem;
    Inverserslection1: TMenuItem;
    LabelAffaire: THLabel;
    ACT_AFFAIRE3: THCritMaskEdit;
    ACT_AFFAIRE2: THCritMaskEdit;
    ACT_AFFAIRE1: THCritMaskEdit;
    BRechAffaire: TToolbarButton97;
    LibelleAffaire: THLabel;
    ACT_TIERS: THCritMaskEdit;
    LIBELLETIERS: THLabel;
    bEffaceAffaireTiers: TToolbarButton97;
    ToolbarButton971: TToolbarButton97;
    Copier1: TMenuItem;
    Couper1: TMenuItem;
    Coller1: TMenuItem;
    BZoomCalendrierAssistant: THBitBtn;
    Slectionneruneligne1: TMenuItem;
    ACT_AFFAIRE0: THCritMaskEdit;
    ACT_AVENANT: THCritMaskEdit;
    BMoisMoins: TToolbarButton97;
    BMoisPlus: TToolbarButton97;
    HPBas: THPanel;
    PagePied: TPageControl;
    SheetPied: TTabSheet;
    PPied: THPanel;
    PTotaux1: THPanel;
    LblPrixUnit: THLabel;
    LblMontantTot: THLabel;
    LblPRCharge: THLabel;
    LblPV: THLabel;
    LblUnite: THLabel;
    LblQte: THLabel;
    LblDate: THLabel;
    LblActiviteReprise: THLabel;
    LibTypeHeure: THLabel;
    ACT_PUPRCHARGE: THNumEdit;
    ACT_TOTPRCHARGE: THNumEdit;
    ACT_PUVENTE: THNumEdit;
    ACT_TOTVENTE: THNumEdit;
    ACT_UNITE: THCritMaskEdit;
    ACT_QTE: THNumEdit;
    ACT_DATEACTIVITE: THCritMaskEdit;
    ACT_ACTIVITEREPRIS: THCritMaskEdit;
    ACT_TYPEHEURE: THCritMaskEdit;
    pInfo1: THPanel;
    LblAffaire: THLabel;
    LblArticle: THLabel;
    HLabelLibelle: THLabel;
    ACT_AFFAIREL: THLabel;
    ACT_ARTICLEL: THLabel;
    LblClient: THLabel;
    ACT_TIERSL: THLabel;
    ACT_RESSOURCEL: THLabel;
    ACT_AFFAIRE3_: THCritMaskEdit;
    ACT_AFFAIRE2_: THCritMaskEdit;
    ACT_AFFAIRE1_: THCritMaskEdit;
    ACT_AFFAIRE_: THCritMaskEdit;
    ACT_CODEARTICLE: THCritMaskEdit;
    ACT_LIBELLE: THCritMaskEdit;
    ACT_TIERS_: THCritMaskEdit;
    ACT_ARTICLE: THCritMaskEdit;
    ACT_AFFAIRE0_: THCritMaskEdit;
    ACT_RESSOURCE_: THCritMaskEdit;
    ACT_AVENANT_: THCritMaskEdit;
    SheetPiedCompl: TTabSheet;
    PPiedAvance: THPanel;
    MOIS_CLOT_CLIENT_: THCritMaskEdit;
    MOIS_CLOT_CLIENT: THCritMaskEdit;
    LabelMoisClotClient: TLabel;
    ACT_ACTIVITEEFFECT: TCheckBox;
    LblActiviteEffect: THLabel;
    BZoomPiece: THBitBtn;
    Bzoompieceach: THBitBtn;
    BZoomLigneAch: THBitBtn;
    Ptotaux3: THPanel;
    lblprixunit_: THLabel;
    lblmontanttot_: THLabel;
    lblprcharge_: THLabel;
    lblpv_: THLabel;
    lblunite_: THLabel;
    lblqte_: THLabel;
    lbldate_: THLabel;
    lblactivitereprise_: THLabel;
    lbltypeheure_: THLabel;
    act_puprcharge_: THNumEdit;
    act_totprcharge_: THNumEdit;
    act_puvente_: THNumEdit;
    act_totvente_: THNumEdit;
    act_unite_: THCritMaskEdit;
    act_qte_: THNumEdit;
    act_dateactivite_: THCritMaskEdit;
    act_activiterepris_: THCritMaskEdit;
    act_typeheure_: THCritMaskEdit;
    PTotaux2: THPanel;
    LblMoisTotaux: THLabel;
    lblTitreTotaux2: THLabel;
    LblTemps: THLabel;
    lblFacturable: THLabel;
    lblNonFacturable: THLabel;
    lblFrais: THLabel;
    CumulTemps: THLabel;
    CumulFac: THLabel;
    CumulNonFac: THLabel;
    CumulFrais: THLabel;
    PourcFac: THLabel;
    PourcNonFac: THLabel;
    BZoomTableauBord: THBitBtn;
    BZoomTiers: THBitBtn;
    HPanel3: THPanel;
    HPanel4: THPanel;
    HPanel6: THPanel;
    LibPiecevente: THLabel;
    NUMPIECEVENTE: THLabel;
    LibCreationFactAff: THLabel;
    BDATEFACTAFF: TToolbarButton97;
    LIBFACTURATION: THLabel;
    DATEFACTAFF: THCritMaskEdit;
    LibDateCreation: THLabel;
    LblDateModification: THLabel;
    ACT_DATECREATION: THCritMaskEdit;
    ACT_DATEMODIF: THCritMaskEdit;
    Datecreat: THLabel;
    DateModif: THLabel;
    LblCreateur: THLabel;
    LblUtilisateur: THLabel;
    Utilisateur: THLabel;
    Createur: THLabel;
    ACT_CREATEUR: THCritMaskEdit;
    ACT_UTILISATEUR: THCritMaskEdit;
    HPanel7: THPanel;
    LibPieceAchat: THLabel;
    NUMPIECEACHAT: THLabel;
    PrixAchat: THLabel;
    LIBACHAT: THLabel;
    HPanel5: THPanel;
    LblOrigine: THLabel;
    ACT_ACTORIGINE: THValComboBox;
    LblIntervalle2: THLabel;
    Toolbar973: TToolbar97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    BImprimer: TToolbarButton97;
    Toolbar974: TToolbar97;
    bNewligne: TToolbarButton97;
    bDelLigne: TToolbarButton97;
    BChercher: TToolbarButton97;
    bSupplement: TToolbarButton97;
    bDescriptif: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    bSelectAll: TToolbarButton97;
    bConsult: TToolbarButton97;
    b35heures: TToolbarButton97;
    bRafraichit: TToolbarButton97;
    bPlanning: TToolbarButton97;
    bInsererLgnAff: TToolbarButton97;
    USER: THCritMaskEdit;
    LIBELLEUSER: THLabel;

{=============================== Gestion de la Form =====================================}
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GereLeF2 (var Key: Word; TOBT : TOB; Vide, OkG, OKAss, OKDates : boolean; Shift: TShiftState; Atester : TShiftState);
    procedure GereLeF3F4 (var Key: Word; TOBT : TOB; Vide, OkG, OKAss, OKDates : boolean; Shift: TShiftState; Atester : TShiftState);
    procedure FormResize(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);

{=============================== Gestion du Grid ========================================}
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure GSElipsisClick(Sender: TObject);
    procedure GSTopLeftChanged(Sender: TObject);
    procedure GSExit(Sender: TObject);
    procedure GSMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GSMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GereLeCtrlF5(Sender: TObject);
    procedure GSGetEditMask(Sender: TObject; ACol, ARow: Integer; var Value: String);
    procedure Inverserslection1Click(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    function  ChangeDeCellule(OldCol, OldRow, NewCol, NewRow : integer):boolean;
    procedure VideCodesLigne ( ARow : integer ) ;
    procedure EffaceClientMission (ARow : integer);

{=============================== Gestion du Grid Calendrier ===================================}
    procedure GCalendrierMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GCalendrierExit(Sender: TObject);
    procedure ChargerCalendrierGrid;

{=============================== Gestion de l'Assistant/Ressource =============================}
    procedure ACT_RESSOURCEElipsisClick(Sender: TObject);
    procedure ACT_RESSOURCEExit(Sender: TObject);
    procedure ACT_RESSOURCEDblClick(Sender: TObject);
    procedure ACT_RESSOURCEChange(Sender: TObject);
    procedure ACT_RESSOURCE_Change(Sender: TObject);
    Procedure CodesAssToCodesLigne ( Control:TControl; AFOAss : TAFO_Ressource ; ARow : integer ) ;
    Function  IdentifierAssistant ( Control:TControl; Var ACol,ARow : integer ; Var Cancel : boolean ; Click : Boolean ) : boolean ;
    function BlocageCalendrierLigne(TOBLigne:TOB; AcbParle:boolean) : integer;
    function ControleSemaineGlissantes(AcsRessource, AcsFolio:string;  bRecupereLimites:boolean;
                                            var AviNumSemErreur1,AviNumSemErreur2:integer;
                                            var AvdNbHMaxSmnLim1:double; var AviNbSemLim1:integer; var AvbSemConsecLim1:boolean;
                                            var AvdNbHMaxSmnLim2:double; var AviNbSemLim2:integer; var AvbSemConsecLim2:boolean;
                                            var AvdNbHDepaceLim1,AvdNbHDepaceLim2:double
                                            ) : T_EnsBlocages35H;
    function ChargeTobSemaineAnalyse35H(AcsRessource, AcsFolio:string; AcdDateDebut, AcdDateFin : TDateTime) : TOB;

{=============================== Gestion de la Mission/Affaire ================================}
    procedure ACT_AFFAIRE0DblClick(Sender: TObject);
    procedure ACT_AFFAIRE1DblClick(Sender: TObject);
    procedure ACT_AFFAIRE2DblClick(Sender: TObject);
    procedure ACT_AFFAIRE3DblClick(Sender: TObject);
    procedure ACT_AFFAIRE0Change(Sender: TObject);
    procedure ACT_AFFAIRE1Change(Sender: TObject);
    procedure ACT_AFFAIRE2Change(Sender: TObject);
    procedure ACT_AFFAIRE2Exit(Sender: TObject);
    procedure ACT_AFFAIRE3Exit(Sender: TObject);
    procedure ACT_AFFAIRE0Exit(Sender: TObject);
    procedure ACT_AFFAIRE1Exit(Sender: TObject);
    procedure ACT_AFFAIREChange(Sender: TObject);
    procedure ACT_AFFAIRE_Change(Sender: TObject);
    procedure ACT_AFFAIREElipsisClick(Sender: TObject);
    procedure ACT_AFFAIREDblClick(Sender: TObject);
    procedure ACT_TIERS_Change(Sender: TObject);
    procedure ACT_TIERSExit(Sender: TObject);

{===================================== Actions, boutons =======================================}
    procedure bNewligneClick(Sender: TObject);
    procedure bDelLigneClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure BZoomAssistantClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomAffaireClick(Sender: TObject);
    procedure bEffaceAffaireTiersClick(Sender: TObject);
    procedure bSupplementClick(Sender: TObject);
    procedure bDescriptifClick(Sender: TObject);
    procedure BRechAffaireClick(Sender: TObject);
    procedure TDescriptifClose(Sender: TObject);
    procedure CpltLigneClick(Sender: TObject);
    procedure Afficherlemmo1Click(Sender: TObject);
    procedure Ajoutduneligne1Click(Sender: TObject);
    procedure Supprimeuneligne1Click(Sender: TObject);
    procedure SetConsultation(AcbConsult:boolean; ActTypBloc:T_TypeBlocageActivite);
    procedure ResizeBoutonsMois;

{===================================== Gestion de la date ==================================}
    procedure ACT_SEMAINEChange(Sender: TObject);
    procedure ACT_MOISChange(Sender: TObject);
    procedure ACT_ANNEEChange(Sender: TObject);

{===================================== Gestion des Totaux ==================================}
    procedure HNumEditTOTPRChange(Sender: TObject);

{=============================== Gestion des articles =================================}
    procedure ACT_CODEARTICLEExit(Sender: TObject);
    procedure ACT_CODEARTICLEChange(Sender: TObject);
    procedure ACT_CODEARTICLEElipsisClick(Sender: TObject);
    procedure ACT_CODEARTICLEEnter(Sender: TObject);
    procedure ACT_LIBELLEEnter(Sender: TObject);
    procedure ACT_LIBELLEExit(Sender: TObject);

{============================== Gestion des champs de saisie du pied ===============================}
    procedure ACT_UNITEExit(Sender: TObject);
    procedure ACT_QTEExit(Sender: TObject);
    procedure ACT_PUPRCHARGEExit(Sender: TObject);
    procedure ACT_TOTPRCHARGEExit(Sender: TObject);
    procedure ACT_ACTIVITEREPRISExit(Sender: TObject);
    procedure ACT_FOLIOChange(Sender: TObject);
    procedure ACT_PUVENTEExit(Sender: TObject);
    procedure ACT_TOTVENTEExit(Sender: TObject);
    procedure ACT_UNITEElipsisClick(Sender: TObject);
    procedure ACT_ACTIVITEREPRISEnter(Sender: TObject);
    procedure ACT_UNITEEnter(Sender: TObject);
    procedure ACT_QTEEnter(Sender: TObject);
    procedure ACT_PUPRCHARGEEnter(Sender: TObject);
    procedure ACT_PUVENTEEnter(Sender: TObject);
    procedure ACT_TOTPRCHARGEEnter(Sender: TObject);
    procedure ACT_TOTVENTEEnter(Sender: TObject);
    procedure ACT_TIERSElipsisClick(Sender: TObject);
    procedure ACT_TIERSDblClick(Sender: TObject);
    procedure ACT_TIERSChange(Sender: TObject);
    procedure pInfo1Enter(Sender: TObject);
    procedure Copier1Click(Sender: TObject);
    procedure Couper1Click(Sender: TObject);
    procedure Coller1Click(Sender: TObject);
    procedure POPYPopup(Sender: TObject);
    procedure PTotaux1Enter(Sender: TObject);
    procedure BZoomCalendrierAssistantClick(Sender: TObject);
    procedure Slectionneruneligne1Click(Sender: TObject);
    procedure BMoisMoinsClick(Sender: TObject);
    procedure BMoisPlusClick(Sender: TObject);
    procedure FormDeactivate(Sender: TObject);
    procedure GSMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure bSelectAllClick(Sender: TObject);
    procedure GSBeforeFlip(Sender: TObject; ARow: Integer;
      var Cancel: Boolean);
    procedure ACT_TYPEHEUREEnter(Sender: TObject);
    procedure ACT_TYPEHEUREExit(Sender: TObject);
    procedure BDATEFACTAFFClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure bConsultClick(Sender: TObject);
    procedure GSSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BZoomPieceAchClick(Sender: TObject);
    procedure BZoomLigneAchClick(Sender: TObject);
    procedure PTotaux2Enter(Sender: TObject);
    procedure BZoomTableauBordClick(Sender: TObject);
    procedure BZoomTiersClick(Sender: TObject);
    procedure ACT_DATECREATIONChange(Sender: TObject);
    procedure ACT_DATEMODIFChange(Sender: TObject);
    procedure ACT_CREATEURChange(Sender: TObject);
    procedure ACT_UTILISATEURChange(Sender: TObject);
    procedure ACT_ACTORIGINEExit(Sender: TObject);
    procedure b35heuresClick(Sender: TObject);
    procedure bRafraichitClick(Sender: TObject);
    procedure bPlanningClick(Sender: TObject);
    procedure bInsererLgnAffClick(Sender: TObject);
    procedure GereArtsLies (ARow: integer; Var ColRetour : integer);
    procedure TraiteLesLies (ARow: integer);
    procedure GSDblClick(Sender: TObject);

  private
    lColsLength:TStringList;
    FindFirst,FClosing : boolean ;
    iTableLigne : integer ;
    StCellCur,LesColonnes : String ;
    ColOld, LigOld : integer;
    ColsInter : Array of boolean ;
    DEV       : RDEVISE ;
    TOBActivite,
    // PL le 03/05/02 : suppression des lignes modifiées par un delete général
    //TOBActiviteOld,
    TOBAssistant, TOBArticles, TOBAffaire, TOBAffaires, TOBNbHeureJour, TOBNbHeureJourAffiche, TOBSommeSemaine, TOBValo,
    TOBLigneActivite, TOBEnteteLigneActivite, TOBListeConv, TOBListeConvAUC, TOBTiers, TOBTierss, TOBUnites, TOBValoZoom: TOB ;
    TOBCXV, TOBFormuleVar, TOBFormuleVarQte : TOB;
    AFOAssistants : TAFO_Ressources;
    GrCalendrier : TGridCalendrierActivite;
    GCanClose : boolean;
    gbSelectCellCalend:boolean;
    OrdreTri:string;
    OldIndexTri : integer;
    bDroitsAccesOK:boolean;
    gbSaisieManager:boolean;
    bAfficheLesValo:boolean;
    bAfficheFNF:boolean;
    gbInitOK:boolean;
    bSourisClickG:boolean;
    iRowDebClick, iNewRow, iOldRow:integer;
    iRowFinClick:integer;
    gbDateEntree:boolean;
    tEnsBlocages:T_EnsBlocagesActivite;
    bActModifiee:boolean;
    gsReqPourDelete:string;
    gdSommeFrais, gdNbHeureMois, gdNbHeureFac, gdNbHeureNFac:double;
    gCtrlF2, gCtrlF3, gCtrlF4:boolean;
    gbModifMois:boolean;
    gbModifFolio:boolean;
    gbModifAnnee:boolean;
    gbModifLibArticle:boolean;
    gbModifActRep:boolean;
    gbTri:boolean;
    giLigneAEnregistrer : integer;
    gbSaisieQtePermise : boolean;
{======================================= Initialisations ===================================}
    procedure InitEnteteDefaut ;
    procedure InitActiviteCreation ;
    procedure InitComposantVisible;
    procedure InitContextuelle(var AcdDateACharger:TDateTime) ;
    procedure ConfigureGridActivite ;
    procedure ClearPanel(AcPanel:THPanel);
//    procedure ClearPanelPTotaux2;
    procedure ReadOnlyPanel(AcPanel:THPanel; AcState:boolean);
    procedure AfficheChargeLesTOBLigne ;
    Procedure DimensionneSuivantContexte ;
    Procedure DimCalendrierSuivantEcran ;
    procedure GereLesRestrictionsDeCriteres;
    procedure InitFormatsColonnesGrid;

{=============================== Actions liées au Grid =====================================}
    function  ControlLigneValide (AciLig : integer; TOBL : TOB) : integer;
    procedure CompleteLaTOB(TOBL:TOB);
    procedure EtudieColsListe ;
    procedure FormateZoneSaisie (Control:TControl; ACol,ARow : Longint ) ;
    Procedure GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
    procedure MetColonnesNullesEnFinGrid(sGrid:THGrid);
    procedure DessineTotaux;
    procedure CollerLignes ;
    procedure CouperLignes ;
    procedure CopierLignes ;
    procedure CopierSelected (bAvecNumUnique, bAvecDeselection : boolean);
    procedure AfficheSuivantCriteres(AcdDateDebCal:TDateTime; AcbTotPied:boolean; AcbMajTotCal:boolean);
    procedure RafraichirLaSelection (DateAAfficher : TDateTime = 0);
    procedure FormatePartiesAffaire ;
    function  EntreDansLigneCourante(Col,Row:integer):boolean;
    function  OKSaisiePartieDroiteGridGS(AciLigne:integer; AcbChangeSiOK:boolean):boolean;
    Procedure ProtegePartieDroiteGS;
    Procedure DeprotegePartieDroiteGS(ARow:integer);
    procedure ConfigSuivantTypeValo;
    procedure ConfigSuivantFNF;
    procedure GotoEntete ;
    procedure GotoPied ;

{=============================== Gestion de l'Assistant/Ressource ==========================}
    Function  IdentifieAssistant(Control:TControl): boolean ;
    Function ChargeAssistant(Control:TControl) : boolean;
    Procedure IncidenceAssistant ;
    Procedure GereAssistantEnabled ;
    Procedure ZoomOuChoixRes ( ACol,ARow : integer );

{=============================== Gestion de la Mission/Affaire ================================}
    Procedure IncidenceAffaire ;
    procedure OnExitPartieAffaire(Sender: TObject);
    function  ChargeAffaireActivite(TestCle:Boolean; bAvecMessage:boolean):boolean ;
    Procedure GereAffaireEnabled ;
    function DernierePartieAffaireVisible : integer;

{=============================== Gestion Tiers ====================================}
    Function ChargeTiers(Control:TControl; AcsCodeTiers:string) : boolean;
    Function IdentifieTiers(Control:TControl) : boolean ;
    Procedure ZoomOuChoixTiers ( ACol,ARow : integer );

{=============================== Gestion Unites ====================================}
    procedure ChargeUnites (Control:TControl) ;
    Function IdentifieUnites(Control:TControl) : boolean ;

{===================================== Actions, boutons =======================================}
    procedure ClickInsert ( ARow : integer; boucle:boolean ) ;
    function  ClickDel (ARow : integer; boucle : boolean) : boolean;
    Procedure ZoomOuChoixArt ( ACol,ARow : integer ) ;
//    procedure ImprimerLaPiece ;
    procedure HValComboBoxTriChange(Sender: TObject);

{========================== Manipulation des TOB lignes activité ============================}
    procedure InsertTOBLigne ( ARow : integer ) ;
    Function  NewTOBLigne ( ARow : integer ) : TOB ;
    procedure CreerTOBLignes ( ARow : integer ) ;
    procedure DepileTOBLignes ( ARow,NewRow : integer ) ;
//    procedure VideLignesIntermediaires;
    procedure ChargerTOBDetailActivite(AcsCritere:string; AcsFolio:string);

{================= Manipulation des TOB heures par jour et par semaine =====================}
    function ChargerTOBNbHeureJour(AcsCritere:string; AcsFolio:string; AcdDateDebut:TDateTime; AcTypeAcc:T_Acc):TOB;
    procedure CompleteLesFraisTOBNbHeureJour;
    function CalculSommeFrais(AcsFolio:string):double;
    function ConversionTOBActivite( var TobActiviteSelect:TOB ; AcsChampRupture:string):Double;
    procedure MajTOBHeureCalendrier(AciLigneTraitee:integer);
    function ChargerTOBSommeSemaine( var TOBHeuresParJour:TOB ):TOB;
    function CalculSommeMois(TOBHeuresParJour:TOB; var AvdNbHeureFac, AvdNbHeureNFac : double ):double;
//    function CalculHeureMois(var AvdQteFac, AvdQteNFac:double):double;
    function MajNbHeureAnRessource (TOBL : TOB; bMoins : boolean) : double;
    procedure MajTotauxMensuels (TobLigneEnBase, TobLigneEnGrid : TOB);

{=============================== Manipulation des LIGNES ===================================}
    procedure InitialiseLigne ( ARow : integer ) ;
    Procedure InitAssistantDefaut ( TOBL : TOB; ARow : integer ; var AvbArtOK:boolean);
    Function  GetTOBLigne ( ARow : integer ) : TOB ;
//    Procedure NumeroteLignes ( GS : THGrid ; TOBActivite : TOB ) ;
    procedure AfficheLaTOBLigne ( ARow : integer ) ;
    procedure AfficheNouvellesLignesActivite;
    Function  GetChampLigne ( Champ : String ; ARow : integer ) : Variant ;
    function  BlocageLigne (TOBL : TOB; bParle : boolean = false; ARow : integer = 0) : boolean;
    function GSRafraichitLaLigne (LaLigne : integer) : TOB;

{=============================== Manipulation des articles =================================}
    Procedure TraiteArticle (Var ACol,ARow : integer ; Var Cancel : boolean ) ;
    Procedure CodesArtToCodesLigne ( Control:TControl; TOBArt : TOB ; ARow : integer; AcbMajTot : boolean; AcbArtChange : boolean=false);
    procedure CopieArtDansLaLigne (TOBArt, TOBLigneAct : TOB; bForcerPrix : boolean = false);
    Function  IdentifierArticle ( Control:TControl; Var ACol,ARow : integer ;  Var Cancel : boolean; Click : Boolean) : boolean;
    function  AppelMajTOBValo( TOBL : TOB; CodeArticle : string; var bPROK, bPVOK : boolean) : boolean;

{====================================== Calculs ============================================}
    procedure MajApresModifActivite;
    procedure MajTotauxActivite;
    Procedure MajTOTPR ( ARow : integer; TOBL : TOB; AbForceMaj:boolean ) ;
    function  TestMajTOTPR ( ARow : integer; TOBL : TOB ):boolean ;
    Procedure MajTOTPV ( ARow : integer; TOBL : TOB; AbForceMaj:boolean ) ;
    function  TestMajTOTPV ( ARow : integer; TOBL : TOB ):boolean ;
    procedure MajTotauxCalendrier(AcdDateDebut:TDateTime);
    procedure AfficheTotauxParMois;
    function ForcerRefreshCalendrier(AciNewYear, AciNewMonth, AciNewDay:integer):boolean;

{===================================== Gestion de la date ==================================}
    procedure MajSemaine;
    Procedure MajIntervalleSemaine;
    function ChargerDerniereDateSaisieCritere (AcbInit : boolean) : boolean;
    function AfficheDerniereDateSaisieCritere(AcdDateDefaut:TDateTime; AcbInit:boolean=false):TDateTime;
    function ControleDateDansIntervalle(AcdDate:TDateTime; AcbParle, AcbTous, AcbTestDatesMission:boolean; var AvbBloqueSaisie:boolean):integer;
    function ChercheJourAAfficherSemaine( AciOldSemaine, AciNewSemaine, AciOldMois : integer ):TDateTime;

{================================ Sauvegarde des champs ====================================}
    Procedure TraiteRessource ( ARow : integer; Var Cancel : boolean ) ;
    Procedure TraiteTypeHeure ( ARow : integer ) ;
    Procedure TraitePUCH ( ARow : integer ) ;
    Procedure TraiteTOTPRCH ( ARow : integer ) ;
    Procedure TraiteAffaire ( ACol,ARow : integer; Var Cancel : boolean ) ;
    procedure TraitePartieAffaire (ACol, ARow : integer; Var Cancel : boolean);
    procedure ViderPartiesInvisiblesCodeAffaire (TOBL : TOB; ARow : integer);
    Procedure TraiteTiers ( ARow : integer ) ;
    Procedure TraiteJour ( ARow : integer; Var Cancel : boolean ) ;
    Procedure TraiteQte  ( ARow : integer ) ;
    Procedure TraiteTotV( ARow : integer ) ;
    Procedure TraiteLib( ARow : integer ) ;
    Procedure TraiteActRep( ARow : integer ) ;
    Procedure TraiteTypA( ARow : integer ) ;
    Procedure TraitePV ( ARow : integer ) ;
    Procedure TraiteMntRemise (ARow : integer);
    function  PutJour( ARow : integer; var AciJour,AciMois,AciAnnee:integer ):integer ;
    Procedure TraiteMontantTVA (ARow : integer);

{================== Gestion globale écran (elipsis, accès menus context,  =================}
    Function  GereElipsis ( Control:TControl; LaCol : integer ) : boolean ;
    Procedure GereEnabled ( ARow : integer ) ;
    function GereSaisieEnabled( AcbAvecMess:boolean ):boolean ;
    Procedure AfficheDetailTOBLigne ( ARow : integer ) ;
    procedure RemplirLesTOBDesLignes;
    procedure RemplirLesTOBCourantes( ARow : integer ) ;
    Procedure GereDescriptif(ARow:Integer; Enter: Boolean);
    Procedure PutEcranPPiedAvance(TOBL : TOB);
    procedure TraiteDateFactAff (Ok : Boolean);
    function ProchaineColonneVide(AciColDepart:integer) : integer;
    procedure CriteresAvantAppel (var sTypeArticle, sDateDeb, sDateFin : string);

{======================================= VALIDATIONS ===================================}
//    procedure ValideLActivite ;
//    function  GereValider : TIOErr;
//    procedure DeleteLignesSuivantCriteres ;
    function  GereValiderLaLigne (NumLigne : integer) : TIOErr;
    procedure ValideLaLigneActivite;

{===================================== Gestion du calendrier ===============================}
    function  DateDebutCalendrier (AcdDateACharger : TDateTime) : TDateTime;
    function  CalendrierAAfficher (AcdDateACharger : TDateTime ) : TDateTime;
    function  ChargerCalendrierEnCours (AciJour, AciMois, AciAnnee : integer; AcbForcer : boolean ) : boolean;
    procedure MajApresChangeJourCalendrier;
{===================================== Gestion du FormuleVar ===============================}
    Procedure TraiteFormuleVar(ARow : integer; Action : String);

  public
     CleAct : R_CLEACT ;
     CleActOld : R_CLEACT;
     Action : TActionFiche ;
     iHauteurEcran, iLargeurEcran : integer;
     gbFicModale : boolean;
     gbCritereFixe : boolean;
     gsRessParDefaut : string;
  end;

implementation
uses   CbpMCD
		,BTPutil
  		,CbpEnumerator
      ;

// uses MenuOLG;

{$R *.DFM}

{==============================================================================================}
{======================================= Fonctions Générales ==================================}
{==============================================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 09/04/2003
Modifié le ... : 16/04/2003
Description .. : Point d'entrée dans la saisie d'activité en inside
                  entrées :
                  tsaTypeSaisie  :  (tsaRess,tsaClient)
                  tacTypeAcces   :  (tacTemps,tacFrais,tacFourn,tacGlobal)
                  sTypeActivite  : 'REA' = réalisé, 'BON' = appréciation
                  sTiersFixe     : code tiers en entrée
                  sRessFixe      : code ressource en entrée
                  sortie :
                  booleen        : false si les droits ne le permettent pas (saisie d'activité manager = toutes les ressources sont accessibles, sinon seule la ressource associée au user courant)


                  on rentre pour l'instant sur la dernière date saisie pour la ressource ou l'affaire stockée dans la registry
                  au moment de la fermeture précédente. Possibilité d'évolution comme la fonction d'appel modal
                  SaveSynRegKey ('DerniereRessource', ACT_RESSOURCE.Text, true);
                  SaveSynRegKey ('DerniereAffaire', ACT_AFFAIRE.Text, true);
                  SaveSynRegKey ('DernierTiers', ACT_TIERS.Text, true);

Mots clefs ... : ACTIVITE;SAISIE
*****************************************************************}
function AFCreerActivite (tsaTypeSaisie : T_Sai; tacTypeAcces : T_Acc; sTypeActivite : string;
                          sTiersFixe : string = ''; sRessFixe : string = '') : boolean;
var
  CleAct : R_CLEACT ;
  iYear, iMonth, iDay : word;
begin
  Result := False;

  if Not (SaisieActiviteManager) and (tsaTypeSaisie = tsaClient) then
    begin
      PGIBoxAf ('Les droits d''accès ne sont pas suffisants ', TitreHalley);
      Exit;
    end;

  Result := True ;
  // Pour l'instant,  on initialise avec la date du jour courant pour le cas où aucune saisie ne serait trouvée
  // pour l'assistant ou l'affaire choisie
  // On pourrait passer une date en paramètre...
  DecodeDate (date, iYear, iMonth, iDay);
  FillChar (CleAct, Sizeof (CleAct), #0);
  CleAct.Annee := iYear;
  CleAct.Mois := iMonth;
  CleAct.Semaine := NumSemaine (date);
  CleAct.Jour := iDay;
  CleAct.Folio := 1;
  CleAct.TypeSaisie := tsaTypeSaisie;
  CleAct.TypeAcces := tacTypeAcces;
  CleAct.TypeActivite := sTypeActivite;
  CleAct.Affaire := '';
  if Not (VH_GC.AFProposAct) then
     CleAct.Aff0 := 'A'
  else
     CleAct.Aff0 := '';
  CleAct.Aff1 := '';
  CleAct.Aff2 := '';
  CleAct.Aff3 := '';
  CleAct.Avenant := '';
  CleAct.Tiers := '';
  CleAct.Ressource := '';
  CleAct.TypeRess := '';
  CleAct.Fonction := '';

  if (tsaTypeSaisie = tsaClient) and (sTiersFixe <> '') then
    begin
      CleAct.Tiers := sTiersFixe;
      if (sRessFixe <> '') then
        CleAct.Ressource := sRessFixe;
    end
  else
  if (tsaTypeSaisie = tsaRess) and (sRessFixe <> '') then
    CleAct.Ressource := sRessFixe;


  if (sTiersFixe <> '') or (sRessFixe <> '') then
    AFSaisieActivite (CleAct, taCreat, false, true, false)
  else
    AFSaisieActivite (CleAct, taCreat, false, false, false);
end;


{***********A.G.L.***********************************************
Auteur  ...... : PL
Créé le ...... : 09/04/2003
Modifié le ... : 16/04/2003
Description .. : Point d'entrée sur la saisie d'activité en modale
                  entrées :
                  tsaTypeSaisie  :  (tsaRess,tsaClient)
                  tacTypeAcces   :  (tacTemps,tacFrais,tacFourn,tacGlobal)
                  sTypeActivite  : 'REA' = réalisé, 'BON' = appréciation
                  sRessource     : ressource souhaitée
                  sAffaire       : mission/affaire souhaitée
                  sTiers         : client souhaité
                  DateEntree     : date souhaitée (affichage du calendrier du mois contenant cette date)
                  iFolio         : numéro de folio, 1 par défaut
                  sortie :
                  booleen        : false si les droits ne le permettent pas (saisie d'activité manager = toutes les ressources sont accessibles, sinon seule la ressource associée au user courant)

Mots clefs ... : ACTIVITE; SAISIE
*****************************************************************}
function AFCreerActiviteModale (tsaTypeSaisie : T_Sai; tacTypeAcces : T_Acc;
                                sTypeActivite, sRessource, sAffaire, sTiers : string;
                                DateEntree : TDateTime; iFolio : integer = 1) : boolean;
  var
  CleAct : R_CLEACT;
  iYear, iMonth, iDay : word;
  bDateEntreeOK : boolean;
begin
  Result := True ;
  if (DateEntree <> 0) and (DateEntree > iDate1900) and (DateEntree < iDate2099) then
    begin
      bDateEntreeOK := true;
    end
  else
    begin
      DateEntree := date;
      bDateEntreeOK := false;
    end;

  DecodeDate (DateEntree, iYear, iMonth, iDay);
  FillChar (CleAct, Sizeof (CleAct), #0);
  CleAct.Annee := iYear;
  CleAct.Mois := iMonth;
  CleAct.Semaine := NumSemaine (DateEntree);
  CleAct.Jour := iDay;
  CleAct.Folio := iFolio;
  CleAct.TypeSaisie := tsaTypeSaisie;
  CleAct.TypeAcces := tacTypeAcces;
  CleAct.TypeActivite := sTypeActivite;
  CleAct.Affaire := sAffaire;
  CleAct.Tiers := sTiers;
  CleAct.Ressource := sRessource;
  CleAct.TypeRess := '';
  CleAct.Fonction := '';

  AFSaisieActivite (CleAct, taCreat, true, false, bDateEntreeOK);
end;

Function AFSaisieActivite (CleAct : R_CLEACT; Action : TActionFiche; bModale : boolean; bCritereFixe : boolean; AcbDateEntreeOK : boolean) : boolean;
var
  X : TFActivite;
  PP : THPanel;
  sTitre : string;
begin
  Result := true;
  SourisSablier;

  X := TFActivite.Create (Application); 
  X.gbDateEntree := AcbDateEntreeOK;
  X.CleAct := CleAct;
  X.CleActOld := CleAct;
  X.gbFicModale := bModale;
  X.gbCritereFixe := bCritereFixe;
  if (CleAct.TypeSaisie = tsaClient) then
    begin
      // Si on est en saisie par client/mission, si une ressource a été passée en paramètre, on la garde par défaut...
      X.gsRessParDefaut := '';
      if (CleAct.Tiers <> '') and (CleAct.Ressource <> '') then
        X.gsRessParDefaut := CleAct.Ressource;
    end;


  // Suivant l'endroit de l'appel, on change le titre du form
  if (CleAct.TypeSaisie = tsaRess) then
    sTitre :=  'Saisie d''activité par ressource'
  else
    sTitre := 'Saisie d''activité par client/affaire';

  // configuration du titre de l'écran suivant le type d'acces à la saisie, par temps, frais, fournitures, global
  case CleAct.TypeAcces of
    tacTemps : sTitre :=  sTitre + ' et par prestation';
    tacFrais : sTitre :=  sTitre + ' et par frais';
    tacFourn :
              if (ctxScot in V_PGI.PGIContexte) then
                 sTitre :=  sTitre + ' et par fourniture'
              else
                 sTitre :=  sTitre + ' et par article';
    tacGlobal : ;
  end;

  // Traduction du titre du form suivant le contexte Scot/Tempo
  X.caption := TraduitGA (sTitre);


  X.Action := Action;
  PP := FindInsidePanel;
  if PP = Nil then
    begin
      try
        X.ShowModal;
      finally
        X.Free;
      end;
      SourisNormale;
    end
  else
    begin
      InitInside(X,PP);
      X.Show;
    end;

END;

{==============================================================================================}
{=========================== Création Form et initialisations =================================}
{==============================================================================================}
procedure TFActivite.FormCreate (Sender : TObject);
begin
// Preinit du grid de saisie
GS.RowCount := NbRowsInit;
iTableLigne := PrefixeToNum ('ACT');

// Creation des TOB utilisees
TOBTiers := TOB.Create ('TIERS', Nil, -1);
TOBTierss := TOB.Create ('Les tiers', Nil, -1);
TOBUnites := TOB.Create ('MEA', Nil, -1);
TOBAssistant := TOB.Create ('RESSOURCE', Nil, -1);
TOBAffaire := TOB.Create ('AFFAIRE', Nil, -1);
TOBAffaires := TOB.Create ('Les affaires', Nil, -1);
TOBEnteteLigneActivite := TOB.Create ('Entete Ligne activite', Nil, -1);
TOBLigneActivite := TOB.Create ('ACTIVITE', TOBEnteteLigneActivite, -1);
TOBValo := Nil;
TOBValoZoom := Nil;
TOBActivite := TOB.Create ('Entete activite', Nil, -1);
// PL le 03/05/02 : suppression des lignes modifiées par un delete général
//TOBActiviteOld:=TOB.Create('Entete activite old',Nil,-1) ;
TOBArticles := TOB.Create ('Les articles', Nil, -1)  ;
AFOAssistants := TAFO_Ressources.Create;
TOBNbHeureJour := TOB.Create ('Nb heures par jour', Nil, -1);
TOBNbHeureJourAffiche := TOB.Create ('Nb heures par jour affichage', Nil, -1);
TOBSommeSemaine := TOB.Create ('Somme par semaine', Nil, -1);

// Initialise les numeros de colonnes du grid de saisie avant définition
InitLesColsAct;

FClosing := False;
end;

procedure TFActivite.FormClose(Sender: TObject; var Action: TCloseAction);
begin
// PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
if csDestroying in ComponentState then Exit ;
if (GS=nil) then exit;
//
// sauve le contexte d'affichage
//
if Not gbFicModale then
    begin
    SaveSynRegKey('ComplementActivite', HPBas.Visible, true);
    SaveSynRegKey('MemoActivite', TDescriptif.Visible, true);
    SaveSynRegKey('MemoX', TDescriptif.Top, true);
    SaveSynRegKey('MemoY', TDescriptif.Left, true);
    SaveSynRegKey('MemoW', TDescriptif.Width, true);
    SaveSynRegKey('MemoH', TDescriptif.Height, true);
    if (ACT_RESSOURCE.visible = true) then
       SaveSynRegKey('DerniereRessource', ACT_RESSOURCE.Text, true);
    if (ACT_AFFAIRE1.visible = true) then
       SaveSynRegKey('DerniereAffaire', ACT_AFFAIRE.Text, true);
    if (ACT_TIERS.visible = true) then
       SaveSynRegKey('DernierTiers', ACT_TIERS.Text, true);
    end;

GS.VidePile(True) ;

//
// Liberation de la mémoire
//
PurgePop(POPZ) ;
TOBTiers.Free; TOBTiers:=Nil;
TOBTierss.Free; TOBTierss:=Nil;
TOBUnites.Free; TOBUnites:=Nil;
TOBAssistant.Free ; TOBAssistant:=Nil ;
TOBAffaire.Free ; TOBAffaire:=Nil ;
TOBAffaires.Free ; TOBAffaires:=Nil ;
TOBEnteteLigneActivite.Free; TOBEnteteLigneActivite:=Nil;
// PL le 03/05/02 : suppression des lignes modifiées par un delete général
//TOBActiviteOld.Free ; TOBActiviteOld:=Nil ;
gsReqPourDelete:='';
///////////////////////////////
TOBFormuleVar.Free; TOBFormuleVar:=Nil ;
TOBFormuleVarQte.Free; TOBFormuleVarQte:=Nil ;
TOBActivite.Free ; TOBActivite:=Nil ;
TOBValo.Free ; TOBValo:=Nil ;
TOBValoZoom.Free ; TOBValoZoom:=Nil ;
TOBArticles.Free ; TOBArticles:=Nil ;
AFOAssistants.Free; AFOAssistants:=nil;
TOBNbHeureJour.Free ; TOBNbHeureJour:=Nil ;
TOBNbHeureJourAffiche.Free ; TOBNbHeureJourAffiche:=Nil ;
TOBSommeSemaine.Free ; TOBSommeSemaine:=Nil ;
gdSommeFrais:=0;gdNbHeureMois:=0;gdNbHeureFac:=0;gdNbHeureNFac:=0;

//TOBCXV.Free; // n'est plus libérée, puisque passée en paramètre par l'intermédiaire de la variable TheTOB
//TOBCXV := nil;
TOBListeConv.Free; TOBListeConv:=nil;
TOBListeConvAUC.Free; TOBListeConvAUC:=nil;
lColsLength.Free;lColsLength:=nil;

GCalendrier.SynEnabled:=false;
GCalendrier.Enabled := false;
GCalendrier.OnDrawCell:=nil;
GCalendrier.OnCellEnter:= nil;
GCalendrier.OnDblClick:= nil;
GCalendrier.OnSelectCell:= nil;
GrCalendrier.GrCal := nil;
GrCalendrier.Free;

if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

procedure TFActivite.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
Var
  Vide,OkG,OKAss, OKDates : boolean;
  index, Decalage : integer;
  wJour, wMois, wAnnee : word;
  TOBT : TOB;
  Atester : TShiftState;
begin
// PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
if csDestroying in ComponentState then Exit ;
if (GS=nil) then exit;
// PL le 11/09/2001 : pour résoudre les effets de la répétition auto des touches up et down
//if (key=VK_LEFT) or (key=VK_RIGHT) or (key=VK_UP) or (key=VK_DOWN) then
//    delay(80);
//GS.Refresh;
// fin PL 11/09/2001

OKAss := true;
OKDates := true;
if (GoRowSelect in GS.Options) then
    OKDates := false;

OkG := (Screen.ActiveControl = GS);
Vide := (Shift=[]);

TOBT := GetTOBLigne(GS.Row);

if (GetParamSoc ('SO_AFANCIENTOUCHE') = true) then
  begin
    Atester := [ssCtrl];

    if (Key = VK_F2) then // Recopie la cellule juste au dessus de la cellule courante dans le grid
                      begin
                        GereLeF2 (Key, TOBT, Vide, OkG, OKAss, OKDates, Shift, Atester);
                      end
    else
    if (Key = VK_RIGHT) then // place le curseur au bout de la saisie dans la cellule courante
                      begin
                        if (OkG) and (Shift = Atester) then
                          begin
                            GS.InplaceEditor.Undo;
                            GS.InplaceEditor.Deselect;
                            Key := 0;
                            Abort;
                          end;
                      end
    else
    if (Key = VK_F3)  or (Key = VK_F4) then // recopie la clé de la ligne juste au-dessus et ouvre la liste des frais ou des fournitures
                      begin
                        GereLeF3F4 (Key, TOBT, Vide, OkG, OKAss, OKDates, Shift, Atester);
                      end
    else
    if (Key = VK_F5) then  // ouvre la liste de recherche
                      begin
                        if OKAss and OKDates then
                          begin
                            if ((OkG) and (Vide)) then
                              BEGIN
                                Key := 0;
                                GSElipsisClick (Sender);
                              END
                            else
                              if ((OkG) and (Shift = Atester)) then
                                BEGIN
                                  Key := 0;
                                  GereLeCtrlF5 (Sender);
                                END;
                          end;
                      end
    else
    ;

  end
else
  begin
    Atester := [ssShift];
    if (Key = VK_F2) then // place le curseur au bout de la saisie dans la cellule courante
                      begin
                        if (OkG) then
                          begin
                            GS.InplaceEditor.Undo;
                            GS.InplaceEditor.Deselect;
                            Key := 0;
                            Abort;
                          end;
                      end
    else
    if (Key = VK_F3) or (Key = VK_F4) then // recopie la clé de la ligne juste au-dessus et ouvre la liste des frais ou des fournitures
                      begin
                        GereLeF3F4 (Key, TOBT, Vide, OkG, OKAss, OKDates, Shift, Atester);
                      end
    else
    if (Key = vk_F7) or (Key = vk_repro) then // Recopie la cellule juste au dessus de la cellule courante dans le grid
                      begin
                        GereLeF2 (Key, TOBT, Vide, OkG, OKAss, OKDates, Shift, Atester);
                      end
    else
    if (Key = VK_F5) or (Key = vk_recherche) then  // ouvre la liste de recherche
                      begin
                        if OKAss and OKDates then
                          begin
                            if ((OkG) and (Vide)) then
                              BEGIN
                                Key := 0;
                                GSElipsisClick (Sender);
                              END
                            else
                            if ((OkG) and (Shift = Atester)) then
                              BEGIN
                                Key := 0;
                                GereLeCtrlF5 (Sender);
                              END;
                          end;
                      end
    else
    ;
  end;


if (Key = VK_DOWN) then //  shift fleche bas, sélectionne la ligne et passe à la suivante
                  begin
                    if (OkG) and (Shift = [ssShift]) then
                      begin
                        GS.FlipSelection (GS.Row);
                      end;
                  end
else
if (Key = VK_UP) then //  shift fleche haut, sélectionne la ligne et passe à la précédente
                  begin
                    if (OkG) and (Shift = [ssShift]) then
                      begin
                        GS.FlipSelection (GS.Row);
                      end;
                  end
else
if (Key = 65) or (Key = vk_SelectAll) then    {Ctrl+A} // Tout/rien sélectionner
                  begin
                    if (OkG) and (Shift = [ssCtrl]) then
                      begin
                        bSelectAll.Down := Not bSelectAll.Down;
                        bSelectAllClick (nil);
                      end;
                  end
else
if (Key = 70) or (Key = vk_cherche) then    {Ctrl+F}  // Recherche dans le grid de saisie
                  begin
                    if (OkG) and (Shift = [ssCtrl]) then
                      begin
                        BChercherClick (nil);
                      end;
                  end
else
if (Key = VK_HOME) then   // Placer le focus dans la zone critères de sélection (en haut)
                  begin
                    if (Shift = [ssCtrl]) then
                      begin
                        Key := 0;
                        GotoEntete;
                      end;
                  end
else
if (Key = VK_END) then   // Placer le focus dans la zone complément (en bas)
                  begin
                    if Shift = [ssCtrl] then
                      begin
                        Key := 0;
                        GotoPied;
                      end;
                  end
else
if (Key = VK_TAB) or (Key = VK_RETURN) then // passage au champ suivant dans le grid
                  begin
                    if OkG and (TOBT <> nil) then
                      BlocageLigne (TOBT, true);

                    if ((OkG) and (Shift = [ssShift]) and (GS.Col = 1) and (GS.Row = 1)) then
                      begin
                        Key := 0;
                        Abort;
                      end
                    else
                      if (OkG) and (GS.Col = SG_Jour) and (GS.Cells[SG_Jour, GS.Row] = '') then
                        begin
                          wJour := 0;
                          DecodeDate (GrCalendrier.DateSelect (GrCalendrier.FDateDeb, GCalendrier.Col, GCalendrier.Row), wAnnee, wMois, wJour);
                          if (wJour <> 0) then
                            GS.Cells[SG_Jour,GS.Row] := inttostr (wJour);
                        end;

                    Key := VK_TAB;
                  end
else
if (Key = VK_F12) then   // Passage de la zone critère au grid et inversement
                  begin
                    if GS.InplaceEditor.Focused then
                      begin
                        GotoEntete;
                        Key := 0;
                      end
                    else
                      begin
                        GS.SetFocus;
                        GS.InplaceEditor.SetFocus;
                      end;
                  end
else
if (Key = 78) or (Key = VK_nouveau) then {Ctrl N}  // insérer une ligne dans le grid
                  begin
                    if OKAss and OKDates then
                      if ((OkG) and (Shift = [ssCtrl])) then
                        BEGIN
                          Key := 0 ;
                          ClickInsert (GS.Row, false);
                        END;
                  end
else
if (Key = VK_INSERT) then // insérer une ligne dans le grid
                  begin
                    if OKAss and OKDates then
                      if ((OkG) and (Vide)) then
                        begin
                          Key := 0 ;
                          ClickInsert (GS.Row, false);
                        end;
                  end
else
if (Key = VK_DELETE) then // supprimer une ligne dans le grid
                  begin
                    if OKAss and OKDates then
                      if ((OkG) and (Shift = [ssCtrl])) then
                        begin
                          GS.InplaceEditor.Deselect;
                          Key := 0;
                          ClickDel (GS.Row, false);
                        end;
                  end
else
if (Key = 67) then     {Ctrl+C} // Couper les lignes sélectionnées
                  begin
                    ClipBoard.Clear ;
                    if ((OkG) and (GS.NbSelected > 0) and (Shift = [ssCtrl])) then
                      BEGIN
                        Key := 0;
                        CopierLignes;
                      END;
                  end
else
if (Key = 69) then     {Ctrl+E} // Effacer les champs du code affaire dans le grid
                  begin

                    if (OkG) and (Shift=[ssCtrl]) then
                      EffaceClientMission (GS.Row);
                  end
else
if (Key = vk_imprime) or (Key = vk_print) then // imprimer
                  begin
                    if (OkG) and (Shift=[ssCtrl]) then
                      BImprimerClick (nil);
                  end
else
if (Key = 77) then  {Ctrl+M} // insérer les lignes d'affaire dans le grid de saisie
                  begin
                    if ((OkG) and (Shift = [ssCtrl])) then
                      BEGIN
                        Key := 0;
                        if (CleAct.TypeSaisie = tsaRess) then
                          bInsererLgnAffClick (nil);
                      END;
                  end
else
if (Key = VK_F9) or (Key = vk_applique) then   // Rafraichir la liste : PL le 04/06/03
                  begin
                    if (vide) then
                      BEGIN
                        bRafraichitClick (nil);
                        Key := 0;
                      END;
                  end
else
if (Key = 82) then {Ctrl+R} // Rafraichir la liste : PL le 04/06/03
                  begin
                    if ((OkG) and (Shift=[ssCtrl])) then
                      BEGIN
                        bRafraichitClick (nil);
                        Key := 0;
                      END;
                  end
else
if (Key = 86) then               {Ctrl+V} // Copier
                  begin
                    if OKAss and OKDates then
                      if ((OkG) and (TOBCXV.Detail.Count > 0) and (Shift = [ssCtrl])) then
                        BEGIN
                          Key := 0;
                          CollerLignes;
                        END;
                  end
else
if (Key = 88) then               {Ctrl+X} // Coller
                  begin
                    if OKAss and OKDates then
                      begin
                        GS.InplaceEditor.Deselect;
                        if ((OkG) and (GS.NbSelected > 0) and (Shift = [ssCtrl])) then
                          BEGIN
                            GS.Cells[SG_Curseur, GS.Row] := '';
                            Key := 0;
                            CouperLignes;
                            Abort;
                          END;
                      end;
                  end
else // Tri sur la colonne 1 à 6
if (Key = 49) or (Key = 50) or (Key = 51) or (Key = 52) or (Key = 53) or (Key = 54) then {Alt+1..6}
                  begin
                    if ((OkG) and (Shift=[ssAlt])) then
                      begin
                        index := -1;
                        Decalage:=0;

                        if (CleAct.TypeSaisie = tsaRess) then
                          begin
                            if (GS.ColLengths[SG_Aff2]=-1) then Decalage:=Decalage+1;
                            if (GS.ColLengths[SG_Aff3]=-1) then Decalage:=Decalage+1;
                          end;

                        // suivant la colonne désirée, on tri
                        if (Key-48 = SG_Jour) then index := HValComboBoxTri.Values.IndexOf('JOU') else
                        if (ctxScot  in V_PGI.PGIContexte) and (Key-48 = SG_Tiers) then
                              index := HValComboBoxTri.Values.IndexOf('CLI') else
                        if Not(ctxScot  in V_PGI.PGIContexte) and (Key-48+Decalage = SG_Tiers) then
                              index := HValComboBoxTri.Values.IndexOf('CLI') else
                        if ((Key-48 = SG_Aff0) and (GS.ColLengths[SG_Aff0]<>-1)) or
                           ((Key-48 = SG_Aff1) and (GS.ColLengths[SG_Aff1]<>-1)) or
                           ((Key-48 = SG_Aff2) and (GS.ColLengths[SG_Aff2]<>-1)) or
                           ((Key-48 = SG_Aff3) and (GS.ColLengths[SG_Aff3]<>-1)) then index := HValComboBoxTri.Values.IndexOf('MIS') else
                        if (Key-48+Decalage = SG_Article) then index := HValComboBoxTri.Values.IndexOf('PRE') else
                        if (CleAct.TypeSaisie <> tsaRess) and (Key-48 = SG_Ressource) then index := HValComboBoxTri.Values.IndexOf('RES') else
                        if (Key-48+Decalage = SG_TypA) then index := HValComboBoxTri.Values.IndexOf('TYP') ;

                        if (index<>-1) then
                          begin
                            HValComboBoxTri.ItemIndex := index;
                            HValComboBoxTriChange(HValComboBoxTri);
                            Abort;
                           end;
                        end;
                  end
else
;

end;


procedure TFActivite.GereLeF3F4 (var Key: Word; TOBT : TOB; Vide, OkG, OKAss, OKDates : boolean; Shift: TShiftState; Atester : TShiftState);
var
  Cancel : boolean;
  NewCol, Col, Row : integer;
  sCodeArt : string;
begin
  if OkG and (TOBT <> nil) then
    if BlocageLigne (TOBT, true) then
      OKAss := false;

  // F3 ou F4
  if (Shift = []) then
    begin
      if (Key = VK_F3) then
        begin
          if (ACT_MOIS.visible = true) then
            begin
              ACT_MOIS.Value := ACT_MOIS.Value - 1;
            end
          else
          if (ACT_SEMAINE.visible = true) then
            begin
              ACT_SEMAINE.Value := ACT_SEMAINE.Value - 1;
            end
          else
          if (ACT_FOLIO.visible = true) then
            begin
              ACT_FOLIO.Value := ACT_FOLIO.Value - 1;
            end;
        end
      else
        begin
          if (ACT_MOIS.visible = true) then
            begin
              ACT_MOIS.Value := ACT_MOIS.Value + 1;
            end
          else
          if (ACT_SEMAINE.visible = true) then
            begin
              ACT_SEMAINE.Value := ACT_SEMAINE.Value + 1;
            end
          else
          if (ACT_FOLIO.visible = true) then
            begin
              ACT_FOLIO.Value := ACT_FOLIO.Value + 1;
            end;
        end;
    end
  else
  // shift F3 ou shift F4
  if OKAss and OKDates then
    // Recopie de la ligne précédente jusqu'au frais
    if (Shift = Atester) and (CleAct.TypeAcces = tacGlobal) then
      begin
        if (Key = VK_F3) then
          begin
            sCodeArt := 'FRA';
            gCtrlF3 := true;
          end
        else
          begin
            sCodeArt := 'MAR';
            gCtrlF4 := true;
          end;

        try
        Key := 0 ;
        Cancel := false;
        if (GS.Row <= 1) then Abort;
        NewCol := SG_Jour;
        GS.Col := SG_Jour;
        StCellCur := GS.Cells[GS.Col, GS.Row] ;
        while (GS.Col < SG_TypA) and (Not Cancel) do
          begin
            GS.Cells[GS.Col, GS.Row] := GS.Cells[GS.Col, GS.Row-1];
            Col := GS.Col;
            Row := GS.Row;
            GSCellExit (GS, Col, Row, Cancel);
            if Not Cancel then
                begin
                //oldCol := GS.Col;
                GS.Col := NewCol;
                Col := GS.Col;
                Row := GS.Row;
                GSCellEnter (GS, Col, Row, Cancel);
                //if Cancel then GS.Col:=oldCol;
                end;

            NewCol := NewCol + 1;
          end;

        if Not Cancel then
          if ((GS.ColLengths[SG_TypA] <> -1) and (GS.ColWidths[SG_TypA] <> 0)) then
            begin
              // FRA ou MAR suivant le cas
              GS.Cells[SG_TypA, GS.Row] := sCodeArt;
              Col := SG_TypA;
              Row := GS.Row;
              GSCellExit(GS, Col, Row, Cancel);
              if Not Cancel then
                begin
                  //oldCol := GS.Col;
                  GS.Col := NewCol;
                  Col := GS.Col;
                  Row := GS.Row;
                  GSCellEnter(GS, Col, Row, Cancel);
                end;

              //NewCol := NewCol + 1;
            end;

        if Not Cancel then
          begin
            if ((GS.ColLengths[SG_Article] <> -1) and (GS.ColWidths[SG_Article] <> 0)) then
              begin
                GS.Col := SG_Article;
                GSElipsisClick (GS);
              end
            else
              GS.Col := SG_Jour;
          end;

        finally
          gCtrlF3 := false;
          gCtrlF4 := false;
        end;
      end;  // (Shift = [ssshift]) and (CleAct.TypeAcces = tacGlobal)
end;

procedure TFActivite.GereLeF2 (var Key: Word; TOBT : TOB; Vide, OkG, OKAss, OKDates : boolean; Shift: TShiftState; Atester : TShiftState);
var
  Cancel : boolean;
  NewCol, Col, oldCol, Row : integer;

begin
  if OkG and (TOBT <> nil) then
    if BlocageLigne (TOBT, true) then
      OKAss := false;

  if OKAss and OKDates then
    if Vide then
      BEGIN
        Key:=0 ;
        if (GS.Row <= 1) then Abort;
        GS.Cells[GS.Col, GS.Row] := GS.Cells[GS.Col, GS.Row-1];

        NewCol := GS.Col + 1;
        while (NewCol < GS.ColCount) and ((GS.ColLengths[NewCol] = -1) or (GS.ColWidths[NewCol] = 0)) do
          NewCol := NewCol + 1;

        if (NewCol < GS.ColCount) then
          begin
            Cancel := false;
            Col := GS.Col;
            Row := GS.Row;
            GSCellExit (GS, Col, Row, Cancel);
            if Not Cancel then
              begin
                oldCol := GS.Col;
                GS.Col := NewCol;
                Col := GS.Col;
                Row := GS.Row;
                GSCellEnter (GS, Col, Row, Cancel);
                if Cancel then GS.Col := oldCol;
              end;
          end;
      END
    else
    // Recopie de la ligne précédente complète
    if (Shift = Atester) then
      begin
        gCtrlF2 := true;
        try
          Key := 0;
          Cancel := false;
          if (GS.Row <= 1) then Abort;
          NewCol := SG_Jour;
          GS.Col := SG_Jour;
          StCellCur := GS.Cells [GS.Col, GS.Row];
          while (GS.Col < SG_Qte) and (Not Cancel) do
            begin
              // PL le 13/12/02 : si le jour est saisi, on ne l'écrase pas dans le Ctrl-F2
              if (GS.Col <> SG_Jour) or (GS.Cells [GS.Col, GS.Row] = '') then
                begin
                  GS.Cells[GS.Col, GS.Row] := GS.Cells[GS.Col, GS.Row-1];
                end;
              //////////// Fin PL le 13/12/02    demande FIDEA
              Col := GS.Col;
              Row := GS.Row;
              GSCellExit (GS, Col, Row, Cancel);
              if Not Cancel then
                begin
                  //oldCol := GS.Col;
                  GS.Col := NewCol;
                  Col := GS.Col;
                  Row := GS.Row;
                  GSCellEnter (GS, Col, Row, Cancel);
                  //if Cancel then GS.Col:=oldCol;
                end;

              NewCol := NewCol + 1;
            end;
          // act_activiterepris
          if Not Cancel then
            if ((GS.ColLengths[SG_Qte] <> -1) and (GS.ColWidths[SG_Qte] <> 0)) then
              GS.Col := SG_Qte else GS.Col := SG_Jour;

        finally
          gCtrlF2 := false;
        end;
      end;
end;


procedure TFActivite.FormResize (Sender: TObject);
begin
  // PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
  if csDestroying in ComponentState then Exit ;
  if (GS = nil) then exit;

  DessineTotaux;
  GCalendrier.Align := alNone;
  GCalendrier.Align := alRight;
  ResizeBoutonsMois;

  // place le libellé de l'assistant correctement par rapport à la zone de saisie
  if (CleAct.TypeSaisie = tsaRess) then
    LibelleAssistant.Left := ACT_RESSOURCE.Left + ACT_RESSOURCE.width + 5;
    
end;

procedure TFActivite.ResizeBoutonsMois;
begin
  //LblMois.Left := GCalendrier.Left + round(GCalendrier.width/2) - round((LblMois.width+BMoisMoins.width+BMoisPlus.width+10)/2);
  // Positionnement des boutons de changement de mois sur le calendrier
  //BMoisMoins.Left := LblMois.Left + LblMois.width + 5;
  //BMoisPlus.Left := BMoisMoins.Left + BMoisMoins.width + 5;
  BMoisPlus.Left := PEntete2.width - BMoisPlus.width - 10;
  BMoisMoins.Left := BMoisPlus.Left - BMoisMoins.width - 10;
  LblMois.Left := BMoisMoins.Left - LblMois.width - 10;
end;

procedure TFActivite.FormActivate (Sender : TObject);
begin
  // PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
  if csDestroying in ComponentState then Exit ;
  if (GS = nil) then exit;

  GereDescriptif (GS.Row, False);
end;

procedure TFActivite.FormDeactivate (Sender : TObject);
begin
  // PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
  if csDestroying in ComponentState then Exit ;
  if (GS = nil) then exit;

  GereDescriptif (GS.Row, true);
end;


procedure TFActivite.FormCloseQuery (Sender : TObject; var CanClose : Boolean);
begin
  GCanClose := true;
  // PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
  if csDestroying in ComponentState then Exit;
  if (GS = nil) then exit;

  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
  GSExit (GS);
  //if (GCanClose=true) then
     // Gere la validation des lignes d'activité saisies
  //   if (GereValider<>oeOk) then GCanClose:=false;

  CanClose := GCanClose;
end;

procedure TFActivite.FormShow (Sender : TObject);
  var
  dDateACharger : tdatetime;
  iAnnee, iMois, iJour : word;
begin

  aglEmpileFiche (self);
  // PL le 17/01/02 pour éviter le plantage aléatoire en entrée ou sortie ?????
  if csDestroying in ComponentState then Exit;
  if (GS = nil) then exit;

  gbInitOK := false;
  // configuration de la liste de saisie suivant le type de saisie par ressource ou affaire
  // en Scot ou Tempo
  if (CleAct.TypeSaisie = tsaRess) then
    begin
      if ctxScot in V_PGI.PGIContexte Then
        begin
          if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
            GS.ListeParam := 'AFSAISIEACTRESSAV'
          else
            GS.ListeParam := 'AFSAISIEACTRESSCT';
        end
      else
        begin
          if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
            GS.ListeParam := 'AFSAISIEACTRESTAV'
          else
            GS.ListeParam := 'AFSAISIEACTRESTMP';
        end;
    end
  else
    GS.ListeParam := 'AFSAISIEACTAFF';

  // Etude des colonnes de la liste de saisie pour accès directs : def des SG_ ...
  EtudieColsListe;

  // Initialisation des valeurs par defaut : affichage par defaut, valeurs à vide
  InitEnteteDefaut;

  // Configuration de l'écran suivant le type de saisie et le paramétrage
  InitActiviteCreation;

  // Configuration du grid de saisie d'activite
  ConfigureGridActivite;
  // Dimensionne et créé le calendrier suivant la résolution écran
  DimCalendrierSuivantEcran;

  // Initialisation de la presentation de l'ecran suivant le contexte (zones visibles ou pas)
  InitComposantVisible;

  // Reconditionne le grid pour eliminer les colonnes de largeur < 5
  MetColonnesNullesEnFinGrid (GS);
  EtudieColsListe;

  // Initialisation des formats des colonnes du Grid
  InitFormatsColonnesGrid;
  (*
  GS.ColTypes[SG_Aff2] := 'I';
  GS.ColFormats[SG_Aff2] := '####,##';
  GS.ColTypes[SG_totv] := 'I';
  GS.ColFormats[SG_totv] := '##0';
  *)
  // Initialisation des valeurs affichees suivant le contexte (saisie precedente, droits, config ecran...)
  InitContextuelle (dDateACharger);

  // Initialisation des tobs nécessaires à l'affichage des lignes d'activité lues en base
  RemplirLesTOBDesLignes;

  // Si la saisie n'est autorisée que pour l'assistant lié au user courant
  GereLesRestrictionsDeCriteres;

  // Si les droits d'accès sont vérouillés (attention : pour l'instant on est aussi forcement dans le cas
  // de la saisie limitée au user courant, on est donc forcement passé dans le test précédent)
  if (bDroitsAccesOK = false) then
     HActivite.Execute (12, Caption, '');

  gbInitOK := true;
  // PL le 28/01/02 pour éviter plantage aléatoire en arrivant sur l'écran
  // Attention : absolument après le gbInitOK := true; car on test gbInitOK dans la fonction
  DecodeDate (dDateACharger, iAnnee, iMois, iJour);
  ChargerCalendrierEnCours (iJour, iMois, iAnnee,  true);
  // Fin PL le 28/01/02

  // Après le chargement du calendrier
  // Dans le cas de la saisie globale par ressource, certains jours où seuls des frais sont saisis n'ont pas de TOB associée
  // dans le calendrier, ce qui fausse les calculs. Il faut les initialiser en création.
  if (CleAct.TypeAcces = tacGlobal) and (CleAct.TypeSaisie = tsaRess) then
    CompleteLesFraisTOBNbHeureJour;


  // retaille les boutons des Mois par rapport au calendrier
  ResizeBoutonsMois;
  // Init du libellé du mois dans le panel des totaux mensuels
  LblMoisTotaux.caption := LblMois.Caption;

end;

procedure TFActivite.InitFormatsColonnesGrid;
var
  mask : string;
begin
  // PL le 29/109/03
(*  mask :=  '#,##0.0' + StringOfChar('0', V_PGI.OkDecQ - 1);
  GS.ColTypes[SG_Qte] := 'R';
  GS.ColFormats[SG_Qte] := mask;*)
  if not GetParamSoc('SO_AFVARIABLES') then       // PL le 14/08/03 pour problème arrondi ONYX
    begin
      mask :=  '#,##0.0' + StringOfChar('0', V_PGI.OkDecQ - 1);
      GS.ColTypes[SG_Qte] := 'R';
      GS.ColFormats[SG_Qte] := mask;
    end
  else
    begin
      GS.ColTypes[SG_Qte] := 'R';
      GS.ColFormats[SG_Qte] := '#,##0.0000000';
    end;
end;

procedure TFActivite.GereLesRestrictionsDeCriteres;
begin
  if (gbSaisieManager = false) then
    begin
      if (CleAct.TypeSaisie = tsaRess) then
        begin
          ACT_RESSOURCE.ElipsisButton := false;
          ACT_RESSOURCE.OnElipsisClick := nil;
          ACT_RESSOURCE.OnDblClick := nil;
          ACT_RESSOURCE.OnChange := nil;
          ACT_RESSOURCE.OnExit := nil;
          ACT_RESSOURCE.ReadOnly := true;
          ACT_RESSOURCE.TabStop := false;
        end
      else
        begin
          // Ce cas ne devrait pas arriver car le test se trouve au niveau de l'appel dans le DispGA
          HActivite.Execute (13, Caption, '') ;
        end;
    end;


  if gbCritereFixe then
    begin
      if (CleAct.TypeSaisie = tsaRess) then
        begin
          ACT_RESSOURCE.ElipsisButton := false;
          ACT_RESSOURCE.OnElipsisClick := nil;
          ACT_RESSOURCE.OnDblClick := nil;
          ACT_RESSOURCE.OnChange := nil;
          ACT_RESSOURCE.OnExit := nil;
          ACT_RESSOURCE.ReadOnly := true;
          ACT_RESSOURCE.TabStop := false;
          // en contexte lanceur et critères fixes, on remplace l'affichage de la ressource par le user associé
          if (ctxlanceur in V_PGI.PGIContexte) and (CleAct.Ressource <> '') then
            begin
              USER.top := ACT_RESSOURCE.top;
              USER.Left := ACT_RESSOURCE.left;
              LIBELLEUSER.top := LIBELLEASSISTANT.Top;
              LIBELLEUSER.left := LIBELLEASSISTANT.left;
              ACT_RESSOURCE.visible := false;
              LIBELLEASSISTANT.visible := false;
              USER.visible := true;
              USER.ReadOnly := true;
              LIBELLEUSER.visible := true;
              USER.Text := V_PGI.UserLogin;
              LIBELLEUSER.Caption := V_PGI.UserName;
            end;
        end
    else
      if (CleAct.TypeSaisie = tsaClient) then
        begin
          ACT_TIERS.ElipsisButton := false;
          ACT_TIERS.OnElipsisClick := nil;
          ACT_TIERS.OnDblClick := nil;
          ACT_TIERS.OnChange := nil;
          ACT_TIERS.OnExit := nil;
          ACT_TIERS.ReadOnly := true;
          ACT_TIERS.TabStop := false;
        end;
    end;
end;

procedure TFActivite.EtudieColsListe;
  Var
  NomCol, LesCols : String;
  icol, ichamp, ii : integer;
  Mcd : IMCDServiceCOM;
  FieldList : IEnumerator ;

BEGIN
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  if (lColsLength <> nil) then
     begin
     lColsLength.Free;
     lColsLength := nil;
     end;
  lColsLength := TStringList.Create;

  SetLength (ColsInter, GS.ColCount);

  for ii := Low (ColsInter) to High (ColsInter) do
    ColsInter[ii] := False;

  LesCols := GS.Titres[0];
  LesColonnes := LesCols;
  icol := 0;
  Repeat
   NomCol := AnsiUppercase (Trim (ReadTokenSt (LesCols)));
   if NomCol <> '' then
      BEGIN
      ichamp := ChampToNum (NomCol);
      //ichamp:=QuelChampLigne(iTableLigne,NomCol) ;
      if ichamp>=0 then
         BEGIN
         if Pos('X',mcd.getField(NomCol).Control)>0 then ColsInter[icol]:=True ;
         if NomCol = 'ACT_TIERS'           then SG_Tiers:=icol else
         if NomCol = 'ACT_AFFAIRE'         then SG_Affaire:=icol else
         if NomCol = 'ACT_RESSOURCE'       then SG_Ressource:=icol else
         if NomCol = 'ACT_CODEARTICLE'     then SG_Article:=icol else
         if NomCol = 'ACT_UNITE'           then SG_Unite:=icol else
         if NomCol = 'ACT_QTE'             then SG_Qte:=icol else
         if NomCol = 'ACT_PUPRCHARGE'      then SG_PUCH:=icol else
         if NomCol = 'ACT_PUVENTE'         then SG_PV:=icol else
         if NomCol = 'ACT_TOTPRCHARGE'     then SG_TOTPRCH:=icol else
         if NomCol = 'ACT_LIBELLE'         then SG_Lib:=icol else
         if NomCol = 'ACT_TOTVENTE'        then SG_TotV:=icol else
         if NomCol = 'ACT_TYPEARTICLE'     then SG_TypA:=icol else
         if NomCol = 'ACT_ACTIVITEREPRIS'  then SG_ActRep:=icol else
         if NomCol = 'ACT_TYPEHEURE'       then SG_TypeHeure:=icol else
         if NomCol = 'ACT_MNTREMISE'       then SG_MntRemise  := icol else
         if NomCol = 'ACT_MONTANTTVA'      then SG_MontantTVA  := icol else
          ;
         END
      else
        BEGIN
         if NomCol='(CURSEUR)'           then SG_Curseur:=icol else
         if NomCol='(JOUR)'              then SG_Jour:=icol else
         if NomCol='(DESCRIPTIF)'        then SG_Desc:=icol else
         if NomCol='(PART0)'             then SG_Aff0:=icol else
         if NomCol='(PART1)'             then SG_Aff1:=icol else
         if NomCol='(PART2)'             then SG_Aff2:=icol else
         if NomCol='(PART3)'             then SG_Aff3:=icol else
         if NomCol='(PART4)'             then SG_Avenant:=icol else
         ;
        END;
      lColsLength.Add(inttostr(icol)+'='+inttostr(GS.ColLengths[icol]));
      END ;
   Inc(icol) ;
  Until ((LesCols='') or (NomCol='')) ;

END ;

procedure TFActivite.InitEnteteDefaut ;
  var
  ColName, StValeur : string;
begin
  // Init des variables globales liees à la forme
  bDroitsAccesOK := true;
  gbSaisieManager := true;
  OrdreTri := '';
  gbSelectCellCalend := false;
  //bTobNbHeuresJourMaj:=true;
  GCanClose := true;
  gbModifMois := false;
  gbModifFolio := false;
  gbModifAnnee := false;
  GbTri := false;

  gdSommeFrais := 0; gdNbHeureMois := 0; gdNbHeureFac := 0; gdNbHeureNFac := 0;
  ClipBoard.Clear;

  // PL le 07/02/03 : on remplace l'ancienne gestion de TheTOB par les fonctions de gestion génériques
  TOBCXV := MaTobInterne ('Mon Activite');

  GS.ColLengths[SG_Desc] := -1;
  GS.ColLengths[SG_Curseur] := -1;
  GS.ColLengths[SG_Affaire] := -1;
  GS.ColLengths[SG_Unite] := -1;

  // Initialisation des champs de saisie lies à la date
  ACT_ANNEE.Value := CleAct.Annee;
  ACT_MOIS.Value := CleAct.Mois;
  ACT_SEMAINE.Value := CleAct.Semaine;
  ACT_FOLIO.Value := CleAct.Folio;
  // Initialisation des champs de saisie lies à la ressource, l'affaire...
  ACT_RESSOURCE.Text := '';
  ACT_AFFAIRE.Text := '';
  ACT_TIERS.Text := '';
  MOIS_CLOT_CLIENT.Text := '';
  if Not (VH_GC.AFProposAct) then
     ACT_AFFAIRE0.Text := 'A'
  else
     ACT_AFFAIRE0.Text := '';
  ACT_AFFAIRE1.Text := '';
  ACT_AFFAIRE2.Text := '';
  ACT_AFFAIRE3.Text := '';
  ACT_AVENANT.Text := '';

  // PL le 29/09/03 : changement d'avis ONYX, on ne veut plus voir toutes les décimales
  // mais on veut toujours le calcul exact.
    // Format d'affichage et de saisie des quantités et prix
  if not GetParamSoc('SO_AFVARIABLES') then       // PL le 14/08/03 pour problème arrondi ONYX
    begin
      ACT_QTE.decimals := V_PGI.OkDecQ;
      ACT_QTE_.decimals := V_PGI.OkDecQ;
      ChangeMask(THNumEdit(ACT_QTE), V_PGI.OkDecQ, '') ;
      ChangeMask(THNumEdit(ACT_QTE_), V_PGI.OkDecQ, '') ;
    end;

  ACT_PUPRCHARGE.decimals := V_PGI.OkDecP;
  ACT_PUPRCHARGE_.decimals := V_PGI.OkDecP;
  ChangeMask(THNumEdit(ACT_PUPRCHARGE), V_PGI.OkDecP, '') ;
  ChangeMask(THNumEdit(ACT_PUPRCHARGE_), V_PGI.OkDecP, '') ;
  ACT_PUVENTE.decimals := V_PGI.OkDecP;
  ACT_PUVENTE_.decimals := V_PGI.OkDecP;
  ChangeMask(THNumEdit(ACT_PUVENTE), V_PGI.OkDecP, '') ;
  ChangeMask(THNumEdit(ACT_PUVENTE_), V_PGI.OkDecP, '') ;
  ACT_TOTPRCHARGE.decimals := V_PGI.OkDecV;
  ACT_TOTPRCHARGE_.decimals := V_PGI.OkDecV;
  ChangeMask(THNumEdit(ACT_TOTPRCHARGE), V_PGI.OkDecV, '') ;
  ChangeMask(THNumEdit(ACT_TOTPRCHARGE_), V_PGI.OkDecV, '') ;
  ACT_TOTVENTE.decimals := V_PGI.OkDecV;
  ACT_TOTVENTE_.decimals := V_PGI.OkDecV;
  ChangeMask(THNumEdit(ACT_TOTVENTE), V_PGI.OkDecV, '') ;
  ChangeMask(THNumEdit(ACT_TOTVENTE_), V_PGI.OkDecV, '') ;


  // Initialisation de la structure de gestion des devises
  FillChar (DEV, Sizeof (DEV), #0);
  DEV.code := V_PGI.DevisePivot;
  GetInfosDevise (DEV);

  // Initialisation de la TOB Activite Mère avec les valeurs passées en paramètre
  ColName := 'TYPESAISIE;TYPEACCES;ACT_TYPEACTIVITE;ACT_AFFAIRE;ACT_TIERS;ACT_RESSOURCE;ANNEE;MOIS;SEMAINE;JOUR;ACT_FOLIO';
  StValeur := Format('%d;%d;%s;%s;%s;%s;%d;%d;%d;%d',[
  Ord(CleAct.TypeSaisie), // Saisie par ressource ou par Client/Affaire
  Ord(CleAct.TypeAcces),  // Accès global, Temps, Frais, Fournitures
  CleAct.TypeActivite,    // Prévisionnel, Réalisé ou Simulation
  CleAct.Affaire,         // Code de l'affaire/Mission en cours
  CleAct.Tiers,           // Code du client en cours
  CleAct.Ressource,       // Code de la Ressource en cours
  CleAct.Annee,           // Annee courante
  CleAct.Mois,            // Mois courant
  CleAct.Semaine,         // Semaine courante
  CleAct.Jour,            // Jour courant
  CleAct.Folio           // Folio courant
  ]);
  TOBActivite.LoadFromSt (ColName, ';', StValeur, ';');

  // PL le 03/05/02 : suppression des lignes modifiées par un delete général
  //TOBActiviteOld.Dupliquer(TOBActivite,TRUE,TRUE,TRUE) ;
  gsReqPourDelete := '';
  ////////////////////////////
end;

procedure TFActivite.InitActiviteCreation ;
BEGIN
// Gestion de la semaine 53
// PL AGL545 : 08/10/02
ACT_SEMAINE.MinValue := 1;
ACT_SEMAINE.MaxValue := 53;

// configure la protection des colonnes PR et PV suivant le type de valorisation
ConfigSuivantTypeValo;

// configure la protection de la colonne F/NF suivant le concept
ConfigSuivantFNF;

// Init de la combo de tri
HValComboBoxTri.DataType := 'AFTRIACTIVITE';
HValComboBoxTri.Value := GetParamSocSecur('SO_AFTRIACTIVITE', 'JOU');
if (HValComboBoxTri.Text='') then
   HValComboBoxTri.Text:=HValComboBoxTri.Values[0];
OldIndexTri := HValComboBoxTri.ItemIndex;


// configuration de l'écran suivant le type d'acces à la saisie, par temps, frais, fournitures, global
case CleAct.TypeAcces of
   tacTemps :begin
            GS.ColLengths[SG_TypA]:=-1;
            GS.ColWidths[SG_TypA]:=0;
            GS.Cells[SG_Article, 0]:= 'Prestation';
            LblArticle.Caption:= 'Prestation';
            HTitres.Mess[0]:= 'Liste des prestations';
            HTitres.Mess[1]:= 'Prestations';
            BZoomArticle.Hint := 'Voir prestation';
            ACT_ACTIVITEEFFECT.visible:=true;
            LblActiviteEffect.visible:=true;
            end;
   tacFrais :begin
            GS.ColLengths[SG_TypA]:=-1;
            GS.ColWidths[SG_TypA]:=0;
            GS.Cells[SG_Article, 0]:= 'Frais';
            LblArticle.Caption:= 'Frais';
            LHNumEditTOTQTE.visible := False;
            HLabelUniteBase.visible := False;
            HTitres.Mess[0]:= 'Liste des frais';
            HTitres.Mess[1]:= 'Frais';
            BZoomArticle.Hint := 'Voir frais';
            HValComboBoxTri.Items[HValComboBoxTri.Values.IndexOf('PRE')]:='par frais';
            ACT_ACTIVITEEFFECT.visible:=false;
            LblActiviteEffect.visible:=false;
            b35heures.visible:=false;
            end;
   tacFourn :begin
            GS.ColLengths[SG_TypA]:=-1;
            GS.ColWidths[SG_TypA]:=0;
            LHNumEditTOTQTE.visible := False;
            HLabelUniteBase.visible := False;
            ACT_ACTIVITEEFFECT.visible:=false;
            LblActiviteEffect.visible:=false;
            GS.Cells[SG_Article, 0]:= 'Fourniture';
            LblArticle.Caption:= 'Fourniture';
            HTitres.Mess[0]:= 'Liste des fournitures';
            HTitres.Mess[1]:= 'Fournitures';
            BZoomArticle.Hint := 'Voir fourniture';
            HValComboBoxTri.Items[HValComboBoxTri.Values.IndexOf('PRE')]:='par fourniture';
            b35heures.visible:=false;
            end;
   tacGlobal:begin
            ACT_ACTIVITEEFFECT.visible:=true;
            LblActiviteEffect.visible:=true;

            if (ctxScot in V_PGI.PGIContexte) then
                begin
                GS.Cells[SG_Article, 0]:= 'Article';
                LblArticle.Caption:= 'Article';
                HTitres.Mess[0]:= 'Liste des articles';
                HTitres.Mess[1]:= 'Articles';
                BZoomArticle.Hint := 'Voir article';
                HValComboBoxTri.Items[HValComboBoxTri.Values.IndexOf('PRE')]:='par article';
                end
            else
                begin
                GS.Cells[SG_Article, 0]:= 'Code';
                LblArticle.Caption:= 'Code';
                HTitres.Mess[0]:= 'Liste des codes';
                HTitres.Mess[1]:= 'Codes';
                BZoomArticle.Hint := 'Voir code';
                HValComboBoxTri.Items[HValComboBoxTri.Values.IndexOf('PRE')]:='par code';
                end;
            end;
   end;

// Unite de conversion par defaut
TOBListeConv := ListeUnitesConversion(VH_GC.AFMESUREACTIVITE);
// Liste des unites 'autres' admises en saisie
TOBListeConvAUC := ListeUnitesConversion('KM');
END ;

procedure TFActivite.ConfigSuivantFNF;
begin
bAfficheFNF:=true;
// Confidentialité sur l'accès à la colonne F/NF
if (Not AffichageFNF)  then bAfficheFNF:=false;

if Not bAfficheFNF then
    begin
    GS.ColLengths[SG_ActRep]:=-1;
    GS.ColWidths[SG_ActRep]:=0;
    LblActiviteReprise.visible:=false;
    ACT_ACTIVITEREPRIS.visible:=false;
    end;
end;

procedure TFActivite.ConfigSuivantTypeValo;
begin
bAfficheLesValo:=true;
// Confidentialité sur l'accès aux prix en saisie de temps et globale
if ((Not AffichageValorisation) and ((CleAct.TypeAcces=tacTemps) or (CleAct.TypeAcces=tacGlobal))) then bAfficheLesValo:=false;

// Gestion de l'affichage de la valorisation
if (VH_GC.AFTYPEVALOACT = '001') and bAfficheLesValo then
// Mont PR + Mont PV
   begin
   GS.ColLengths[SG_PUCH]:=-1;
   GS.ColLengths[SG_PV]:=-1;
   GS.ColWidths[SG_PUCH]:=0;
   GS.ColWidths[SG_PV]:=0;
   end
else
if (VH_GC.AFTYPEVALOACT = '002') and bAfficheLesValo then
// Mont PR
   begin
   GS.ColLengths[SG_PUCH]:=-1;
   GS.ColLengths[SG_PV]:=-1;
   GS.ColLengths[SG_TOTV]:=-1;
   GS.ColWidths[SG_PUCH]:=0;
   GS.ColWidths[SG_PV]:=0;
   GS.ColWidths[SG_TOTV]:=0;
   LHNumEditTOTPV.visible := false;
   end
else
if (VH_GC.AFTYPEVALOACT = '003') and bAfficheLesValo then
// Mont PV
   begin
   GS.ColLengths[SG_PUCH]:=-1;
   GS.ColLengths[SG_PV]:=-1;
   GS.ColLengths[SG_TOTPRCH]:=-1;
   GS.ColWidths[SG_PUCH]:=0;
   GS.ColWidths[SG_PV]:=0;
   GS.ColWidths[SG_TOTPRCH]:=0;
   LHNumEditTOTPR.visible := false;
   end
else
if (VH_GC.AFTYPEVALOACT = '004') and bAfficheLesValo then
// PU + Mont PR
   begin
   GS.ColLengths[SG_TOTV]:=-1;
   GS.ColLengths[SG_PV]:=-1;
   GS.ColWidths[SG_TOTV]:=0;
   GS.ColWidths[SG_PV]:=0;
   LHNumEditTOTPV.visible := false;
   end
else
if (VH_GC.AFTYPEVALOACT = '005') and bAfficheLesValo then
// PU + Mont PV
   begin
   GS.ColLengths[SG_PUCH]:=-1;
   GS.ColLengths[SG_TOTPRCH]:=-1;
   GS.ColWidths[SG_PUCH]:=0;
   GS.ColWidths[SG_TOTPRCH]:=0;
   LHNumEditTOTPR.visible := false;
   end
else
if (VH_GC.AFTYPEVALOACT = '006') or Not bAfficheLesValo then
// Aucun
   begin
   GS.ColLengths[SG_TOTV]:=-1;
   GS.ColLengths[SG_PV]:=-1;
   GS.ColLengths[SG_PUCH]:=-1;
   GS.ColLengths[SG_TOTPRCH]:=-1;
   GS.ColWidths[SG_TOTV]:=0;
   GS.ColWidths[SG_PV]:=0;
   GS.ColWidths[SG_PUCH]:=0;
   GS.ColWidths[SG_TOTPRCH]:=0;
   LHNumEditTOTPR.visible := false;
   LHNumEditTOTPV.visible := false;
   end;

if Not bAfficheLesValo then
   begin
   LblMontantTot.visible:=false;
   LblPrixUnit.visible:=false;
   LblPRCharge.visible:=false;
   LblPV.visible:=false;
   ACT_PUPRCHARGE.visible:=false;
   ACT_PUVENTE.visible:=false;
   ACT_TOTPRCHARGE.visible:=false;
   ACT_TOTVENTE.visible:=false;
   LblMontantTot_.visible:=false;
   LblPrixUnit_.visible:=false;
   LblPRCharge_.visible:=false;
   LblPV_.visible:=false;
   ACT_PUPRCHARGE_.visible:=false;
   ACT_PUVENTE_.visible:=false;
   ACT_TOTPRCHARGE_.visible:=false;
   ACT_TOTVENTE_.visible:=false;
   end;
end;

procedure TFActivite.ConfigureGridActivite ;
begin
// Modif des options du grid de saisie de l'activite en fonction du mode (creat,modif …) => property rowSelect …
AffecteGrid(GS,Action) ;
// F° d'affichage du grid de saisie de l'activite
GS.GetCellCanvas:=GetCellCanvas ;
// Masque le bouton de passage en mode consult/modif quand on est en consult
if (Action=taConsult) then bConsult.visible:=false;
end;

procedure TFActivite.InitComposantVisible;
begin
Ptotaux1.visible:=true;
Ptotaux2.visible:=false;
Ptotaux3.visible:=false;
HPanel5.visible:=false;

// configuration de l'écran suivant le type de saisie par ressource ou affaire
if (CleAct.TypeSaisie = tsaRess) then
   begin
// gestion en 3 parties du code affaire
   GS.ColWidths[SG_Affaire]:=0 ;
   GS.ColLengths[SG_Affaire]:=-1;
//   if (GetParamSoc('SO_AFFGESTIONAVENANT')= false) then
//        begin
//        GS.ColWidths[SG_Avenant]:=0 ;
//        GS.ColLengths[SG_Avenant]:=-1;
//        end;
///////////////////
   GS.ColWidths[SG_Ressource]:=0;
   GS.ColLengths[SG_Ressource]:=-1;
   ACT_RESSOURCE.visible := true;
   LibelleAssistant.visible := true;
   ACT_TIERS.visible := false;
   MOIS_CLOT_CLIENT.visible := false;
   LabelMoisClotClient.visible := false;
   LibelleTiers.visible := false;
   ACT_AFFAIRE0.visible := false;
   ACT_AFFAIRE1.visible := false;
   ACT_AFFAIRE2.visible := false;
   ACT_AFFAIRE3.visible := false;
   ACT_AVENANT.visible := false;
   ACT_RESSOURCE_.visible := false;
   ACT_RESSOURCEL.visible := false;
//   ACT_AFFAIRE_.visible := true;
   ACT_AFFAIRE0_.visible := true;
   ACT_AFFAIRE1_.visible := true;
   ACT_AFFAIRE2_.visible := true;
   ACT_AFFAIRE3_.visible := true;
   ACT_AVENANT_.visible := true;
   LabelAffaire.visible := false;
   LabelAssistant.Caption := TraduitGA('Ressource');
   LibelleAssistant.Left := ACT_RESSOURCE.Left + ACT_RESSOURCE.width + 5;
   LibelleAffaire.visible := false;
   BRechAffaire.visible := false;
   LblAffaire.Caption := TraduitGA('Affaire');
   GS.Cells[SG_Affaire, 0]:= TraduitGA(GS.Cells[SG_Affaire, 0]);
//   if (GetParamSoc('SO_AFRECHTIERSPARNOM')=true) then
//        GS.Cells[SG_Tiers, 0]:= 'Raison sociale'
//   else
   GS.Cells[SG_Tiers, 0]:= TraduitGA(GS.Cells[SG_Tiers, 0]);
   bEffaceAffaireTiers.Visible := false;
   if (ctxScot in V_PGI.PGIContexte) then
      begin
        Ptotaux1.visible:=false;
        Ptotaux2.visible:=true;
        Ptotaux3.visible:=true;
      end
   else
      begin
// PL le 04/07/03 : on ne cache pas cette fonctionnalité pour l'instant : fait pour ONYX      
//        bInsererLgnAff.visible := true;
      end;
   end
else
   begin
   b35heures.visible:=false;
   GS.ColWidths[SG_Affaire]:=0 ;
   GS.ColLengths[SG_Affaire]:=-1;
   GS.ColWidths[SG_Tiers]:=0 ;
   GS.ColLengths[SG_Tiers]:=-1;
   ACT_RESSOURCE.visible := false;
   ACT_TIERS.visible := true;
   MOIS_CLOT_CLIENT.visible := true;
   LabelMoisClotClient.visible := true;
   LibelleAssistant.visible := false;
   LibelleTiers.visible := true;
   ACT_TIERS_.visible := false;
   MOIS_CLOT_CLIENT_.visible := false;
   ACT_TIERSL.visible := false;
   LblClient.visible := false;
   ACT_AFFAIRE0.visible := true;
   ACT_AFFAIRE1.visible := true;
   ACT_AFFAIRE2.visible := true;
   ACT_AFFAIRE3.visible := true;
   ACT_AVENANT.visible := true;
//   ACT_AFFAIRE_.visible := false;
   ACT_AFFAIRE0_.visible := false;
   ACT_AFFAIRE1_.visible := false;
   ACT_AFFAIRE2_.visible := false;
   ACT_AFFAIRE3_.visible := false;
   ACT_AVENANT_.visible := false;
   ACT_RESSOURCE_.visible := true;
   ACT_RESSOURCEL.visible := true;
   ACT_RESSOURCEL.top := ACT_AFFAIREL.top;
//   ACT_RESSOURCEL.left := ACT_AFFAIREL.left;
   ACT_TIERS.top := ACT_RESSOURCE.top;
   ACT_TIERS.left := ACT_RESSOURCE.left;
   LibelleTiers.Left := ACT_TIERS.Left + ACT_TIERS.width + 5;
   ACT_RESSOURCE_.top := ACT_AFFAIRE0_.top;
   ACT_RESSOURCE_.left := ACT_AFFAIRE0_.left;
   LblAffaire.Caption := TraduitGA('Ressource');
   GS.Cells[SG_Ressource, 0]:= TraduitGA(GS.Cells[SG_Ressource, 0]);
   // PL le 06/11/02 : GM ne veut plus Tiers, mais 'Client' comme en GI
   //   LabelAssistant.Caption := TraduitGA('Tiers');
   LabelAssistant.Caption := 'Client';
   ////////////////////////////////
   LibelleAffaire.visible := true;
   LabelAffaire.visible := true;
   BRechAffaire.visible := true;
   end;

// Gestion du type de saisie : par Mois, par Semaine, par Folio
if (VH_GC.AFTYPESAISIEACT = 'FOL') then
   begin
   ACT_SEMAINE.visible := false;
   LabelSemaine.visible := false;
   LblIntervalle.visible := false;
   LblIntervalle2.visible := false;
   ACT_FOLIO.Top := ACT_SEMAINE.Top;
   ACT_FOLIO.Left := ACT_SEMAINE.Left;
   LblFolio.Top := LabelSemaine.Top;
   LblFolio.Left := LabelSemaine.Left;
   LblTri.Top := LabelSemaine.Top;
   HValComboBoxTri.Top := ACT_FOLIO.Top;
   LblTri.Left := ACT_FOLIO.Left + ACT_FOLIO.width + 10;
   HValComboBoxTri.Left := LblTri.Left + LblTri.width + 5;
// PL le 20/01/03 : pas de sens de déplacer le libellé du mois, on le laisse au dessus du calendrier
//   LblMois.Top := LblIntervalle.Top;
//   LblMois.Left := LblIntervalle.Left;
   end
else
if (VH_GC.AFTYPESAISIEACT = 'MOI') then
   begin
   ACT_FOLIO.Visible := false;
   LblFolio.visible := false;
   ACT_SEMAINE.visible := false;
   LabelSemaine.visible := false;
   LblIntervalle.visible := false;
   LblIntervalle2.visible := false;
// PL le 20/01/03 : pas de sens de déplacer le libellé du mois, on le laisse au dessus du calendrier
//   LblMois.Top := LblIntervalle.Top;
//   LblMois.Left := LblIntervalle.Left;
   HValComboBoxTri.Top := ACT_SEMAINE.Top;
   HValComboBoxTri.Left := ACT_SEMAINE.Left;
   LblTri.Top := LabelSemaine.Top;
   LblTri.Left := LabelSemaine.Left;
   end
else
if (VH_GC.AFTYPESAISIEACT = 'SEM') then
  begin
  ACT_FOLIO.Visible := false;
  LblFolio.visible := false;
  LabelMois.visible := false;
  ACT_MOIS.visible := false;
  HValComboBoxTri.Top := ACT_SEMAINE.Top;
  HValComboBoxTri.Left := ACT_SEMAINE.Left;
  LblTri.Top := LabelSemaine.Top;
  LblTri.Left := LabelSemaine.Left;
  ACT_SEMAINE.Top := ACT_MOIS.Top;
  ACT_SEMAINE.Left := ACT_MOIS.Left;
  LabelSemaine.Top := LabelMois.Top;
  LabelSemaine.Left := LabelMois.Left;
  LblIntervalle.Top := LblIntervalle.Top - 6;
  LblIntervalle2.Top := LblIntervalle2.Top + 6;
  LblIntervalle2.Left := LblIntervalle.Left;
  end;

// Peut-on saisir sur des propositions ?
if Not (VH_GC.AFProposAct) then
   begin
   ACT_AFFAIRE0.visible := false;
   ACT_AFFAIRE0_.visible := false;
   ACT_AFFAIRE1.Left := ACT_AFFAIRE0.Left;
   ACT_AFFAIRE1_.Left := ACT_AFFAIRE0_.Left;
   end;
// Saisie possible du type d'heures
if (GetParamSoc('SO_AFTYPEHEUREACT')= False) or (CleAct.TypeAcces=tacFrais) or (CleAct.TypeAcces=tacFourn) then
   begin
   ACT_TYPEHEURE.Visible := False;
   LibTypeHeure.Visible := False; // PL le 15/01/02 pour V575
   act_typeheure_.Visible := False;
   lbltypeheure_.Visible := False; // PL le 16/12/02
   GS.ColWidths[SG_TypeHeure]:=0 ;
   GS.ColLengths[SG_TypeHeure]:=-1;
   end;

HLabelUniteBase.Caption := 'Unité d''activité en ' + RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
LblTemps.caption := 'Activité en ' + RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
LblFrais.caption := 'Montant des Frais en ' + V_PGI.DevisePivot;
HNumEditTOTPR.value := 0;
HNumEditTOTQTE.value := 0;
HNumEditTOTPV.value := 0;

if (CleAct.TypeSaisie = tsaRess) then
   // ATTENTION : les colonnes Aff0, Aff1, Aff2 et Aff3 n'existent pas en saisie par client/mission
   begin
   // affecte la bonne largeur aux colonnes paramétrables du code mission
   if Not (VH_GC.AFProposAct) then
      begin
      GS.ColWidths[SG_Aff0]:=0 ;
      GS.ColLengths[SG_Aff0]:=-1;
      end;

   if Not VH_GC.CleAffaire.Co2Visible then
      begin
      GS.ColWidths[SG_Aff2]:=0 ;
      GS.ColLengths[SG_Aff2]:=-1;
      end;

   if Not VH_GC.CleAffaire.Co3Visible then
      begin
      GS.ColWidths[SG_Aff3]:=0 ;
      GS.ColLengths[SG_Aff3]:=-1;
      end;
   end;

if (SG_MntRemise <> -1) and (GetParamSoc('SO_AFREMISEMONTANT') = false) then
  begin
    GS.ColWidths[SG_MntRemise] := 0;
    GS.ColLengths[SG_MntRemise] := -1;
  end;

if (SG_MontantTVA <> -1) and (GetParamSoc('SO_AFMONTANTTVA') = false) then
  begin
    GS.ColWidths[SG_MontantTVA] := 0;
    GS.ColLengths[SG_MontantTVA] := -1;
  end;

   // visible qu'en contexte Scot ou Tempo
if Not (ctxScot in V_PGI.PGIContexte) then
    begin
    MOIS_CLOT_CLIENT.visible := false;
    LabelMoisClotClient.visible := false;
    MOIS_CLOT_CLIENT_.visible := false;
    ACT_TIERS_.Left := MOIS_CLOT_CLIENT_.Left;
    ACT_TIERSL.Left := ACT_TIERS_.Left + ACT_TIERS_.Width + 6;
    //ACT_ACTORIGINE n'est visible qu'en contexte tempo
    HPanel5.visible:=true;
    Ptotaux3.visible:=false;
    end;

// PL le 27/08/02 : le libellé est saisissable sur 70 caractères
GS.ColLengths[SG_Lib] := 70;
end;

procedure TFActivite.MetColonnesNullesEnFinGrid(sGrid:THGrid);
var
i,j,NbColDep:integer;
w1,w2,l1,l2:integer;
tab1,tab2:boolean;
ColInter1,ColInter2:HTStrings;
T1,T2 : Char ;
A1,A2 : TAlignment ;
F1,F2 : String ;
sTitres:string;
TitresCols:TStringList;
begin
sTitres:=sGrid.Titres[0];
TitresCols := TStringList.Create;
TitresCols.Text := StringReplace(sTitres, ';', chr(VK_RETURN), [rfReplaceAll]);
try
with sGrid do
begin
i:=0;
NbColDep:=0;
while (i<=(ColCount-1-NbColDep)) do
   begin
   if (ColWidths[i]<5) then
      begin
      for j:=i+1 to ColCount-1 do
      begin
      ColInter1:= Cols[j-1];
      ColInter2:= Cols[j];
      T1:=ColTypes[j-1];
      T2:=ColTypes[j];
      A1:=ColAligns[j-1] ;
      A2:=ColAligns[j] ;
      F1:=ColFormats[j-1] ;
      F2:=ColFormats[j] ;
      l1:=ColLengths[j-1];
      l2:=ColLengths[j];
      w1:=ColWidths[j-1];
      w2:=ColWidths[j];
      tab1:=TabStops[j-1];
      tab2:=TabStops[j];

      Cols[j-1]:=ColInter2;
      Cols[j]:=ColInter1;
      ColTypes[j-1]:=T2;
      ColTypes[j]:=T1;
      ColAligns[j-1]:=A2 ;
      ColAligns[j]:=A1 ;
      ColFormats[j-1]:=F2 ;
      ColFormats[j]:=F1 ;
      ColLengths[j-1]:=l2;
      ColLengths[j]:=l1;
      ColWidths[j-1]:=w2;
      ColWidths[j]:=w1;
      TabStops[j-1]:=tab2;
      TabStops[j]:=tab1;

      TitresCols.exchange(j-1, j);
      end;
      NbColDep := NbColDep + 1;
      end
   else
      i:=i+1;

   end;
end;

GS.Titres[0]:= StringReplace(TitresCols.CommaText, ',', ';', [rfReplaceAll]);

finally
TitresCols.Free;
end;
end;

procedure TFActivite.InitContextuelle(var AcdDateACharger:TDateTime) ;
  var
  bRegKey : boolean;
  iIndiceRess, iPos : integer;
  sRegKey : string;
  bBloqueSaisie : boolean;
  AFORessources : TAFO_Ressources;
begin
  bRegKey := false;
  sRegKey:='';
  iPos:=1;
  // Remet le contexte d'affichage
  if Not gbFicModale then
    begin
      HPBas.Visible := GetSynRegKey('ComplementActivite', bRegKey, true);
      bSupplement.Down := HPBas.Visible;
      CpltLigne.Checked := HPBas.visible;
      TDescriptif.Visible := GetSynRegKey('MemoActivite', bRegKey, true);
      BDescriptif.Down:=TDescriptif.Visible;
      Afficherlemmo1.Checked :=TDescriptif.Visible;
      TDescriptif.Top := GetSynRegKey('MemoX', iPos, true);
      TDescriptif.Left := GetSynRegKey('MemoY', iPos, true);
      TDescriptif.Width := GetSynRegKey('MemoW', iPos, true);
      TDescriptif.Height := GetSynRegKey('MemoH', iPos, true);
    end;

  // Réajustement en cas de problème de dimension ou d'emplacement du descriptif
  if (TDescriptif.Top < 10) or (TDescriptif.Top > screen.Height - 50) then
    TDescriptif.Top := trunc(screen.Height / 2);
  if (TDescriptif.Left < 10) or (TDescriptif.Left > screen.Width - 50) then
    TDescriptif.Left := trunc(screen.Width / 2);
  if (TDescriptif.Width < 50) then
    TDescriptif.Width := 50;
  if (TDescriptif.Height < 50) then
    TDescriptif.Height := 50;

  // Confidentialité d'accès sur l'assistant
  gbSaisieManager := false;
  if SaisieActiviteManager then gbSaisieManager := true;

  // Remet le contexte des valeurs suivant la ressource précédente
  if (ACT_RESSOURCE.visible = true) then
    begin
      // on debranche les evenements de la gestion par affaire
      ACT_AFFAIRE.onchange := nil;
      ACT_AFFAIRE0.onchange := nil;
      ACT_AFFAIRE1.onchange := nil;
      ACT_AFFAIRE2.onchange := nil;
      ACT_AFFAIRE3.onchange := nil;
      // Affiche les zones de saisie de l'affaire par ligne sur le grid de saisie
      FormatePartiesAffaire ;

      // Si pas de Confidentialité d'accès sur l'assistant
      if gbSaisieManager then
        begin
          if gbFicModale or gbCritereFixe then
            // Si on est en critere fixe ou en saisie modale, on prend la ressource passée en paramètre
            ACT_RESSOURCE.Text := CleAct.Ressource
          else
            // sinon, On initialise la ressource courante avec la derniere ressource saisie
            ACT_RESSOURCE.Text := GetSynRegKey ('DerniereRessource', sRegKey, true);
        end
      else
        // si confidentialité, on va sur l'assistant associé au user courant
        begin
          ACT_RESSOURCE.Text := VH_GC.RessourceUser;
          if (ACT_RESSOURCE.Text = '') then
            begin
              // Recherche de l'assistant à partir du user
              AFORessources := TAFO_Ressources.Create;
              try
                ACT_RESSOURCE.Text := AFORessources.RessourceD1User (V_PGI.User);
                if (ACT_RESSOURCE.Text = '') then
                  bDroitsAccesOK := false;
              finally
                AFORessources.Free;
              end
            end;
        end;

      CleAct.Ressource := ACT_RESSOURCE.Text;
      // Affiche la derniere date de saisie pour la ressource courante
      AcdDateACharger := AfficheDerniereDateSaisieCritere (GrCalendrier.FDateDeb, true);

      // Teste l'existence de l'assistant et affiche les lignes correspondant aux criteres
      iIndiceRess := RemplirTOBAssistant (TOBAssistant, AFOAssistants, ACT_RESSOURCE.Text);
      if (iIndiceRess >= 0) then
        begin
          CleAct.TypeRess := TOBAssistant.GetValue ('ARS_TYPERESSOURCE');
          AfficheSuivantCriteres (DateDebutCalendrier (AcdDateACharger), true, true);
        end;

      // Test suivant la date courante si on peut entrer en saisie ou en consultation
      if (ControleDateDansIntervalle (AcdDateACharger, false, false, false, bBloqueSaisie) <> 0) then
        begin
          if bBloqueSaisie then
            SetConsultation (True, tbaDatesAct);
        end
      else
        SetConsultation (False, tbaDatesAct);


      ACT_RESSOURCE.SetFocus;
    end;

  // Remet le contexte des valeurs suivant l'affaire précédente
  if (ACT_AFFAIRE1.visible = true) then
    begin
      // On initialise Client et Affaire courante avec la derniere affaire saisie
      ACT_TIERS.onchange := nil;

      if gbFicModale or gbCritereFixe then
        ACT_TIERS.Text := CleAct.Tiers
      else
        ACT_TIERS.Text := GetSynRegKey('DernierTiers', sRegKey, true);
      ACT_TIERS.onchange := ACT_TIERSChange;

      ACT_AFFAIRE.onchange := nil;

      if Not gbFicModale then
        ACT_AFFAIRE.Text := GetSynRegKey('DerniereAffaire', sRegKey, true)
      else
        ACT_AFFAIRE.Text := CleAct.Affaire;
      if Not (VH_GC.AFProposAct) then
        if (copy (ACT_AFFAIRE.Text, 0, 1) <> 'A') then ACT_AFFAIRE.Text := '';
          ACT_AFFAIRE.onchange := ACT_AFFAIREChange;

      CleAct.Tiers := ACT_TIERS.Text;
      ChargeTiers(ACT_TIERS, CleAct.Tiers);
      CleAct.Affaire := ACT_AFFAIRE.Text;

      {$IFDEF BTP}
	  BTPCodeAffaireDecoupe(CleAct.Affaire,Cleact.Aff0,CleAct.Aff1,CleAct.Aff2,CleAct.Aff3,Cleact.Avenant,Action, false);
      {$ELSE}
      CodeAffaireDecoupe(CleAct.Affaire,Cleact.Aff0,CleAct.Aff1,CleAct.Aff2,CleAct.Aff3,Cleact.Avenant,Action, false);
      {$ENDIF}

      // Affiche la derniere date de saisie pour l'affaire courante
      AcdDateACharger := AfficheDerniereDateSaisieCritere(GrCalendrier.FDateDeb, true);
      AfficheSuivantCriteres(DateDebutCalendrier(AcdDateACharger), true, true);


      // Test suivant la date courante si on peut entrer en saisie ou en consultation
      if (ControleDateDansIntervalle (AcdDateACharger, false, false, true, bBloqueSaisie) <> 0) then
        begin
          if bBloqueSaisie then
            SetConsultation( True, tbaDatesAct);
        end
      else
        SetConsultation (False, tbaDatesAct);


      ChargeCleAffaire (ACT_AFFAIRE0, ACT_AFFAIRE1, ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, nil, Action, ACT_AFFAIRE.Text, False);
      ChargeAffaireActivite (true, false);

      if (gsRessParDefaut <> '') then
        // Si on a une ressource par défaut, on charge cette ressource
        begin
          iIndiceRess := AFOAssistants.AddRessource (gsRessParDefaut);

          if (iIndiceRess = -1) or (iIndiceRess = -2) then
            gsRessParDefaut := '';
        end;

      if (ACT_AFFAIRE0.canFocus) then
        ACT_AFFAIRE0.SetFocus
      else
        ACT_AFFAIRE1.SetFocus;
    end
  else
    begin
      ACT_AFFAIRE.Text := ACT_AFFAIRE_.Text;
    end;

  ChargeAssistant (ACT_RESSOURCE);

  // Dimensionne les composants suivant les paramètres, PGI...
  DimensionneSuivantContexte;

  // Branche les evenements de gestion des dates et du tri
  ACT_ANNEE.OnChange := ACT_ANNEEChange;
  ACT_MOIS.OnChange := ACT_MOISChange;
  ACT_SEMAINE.OnChange := ACT_SEMAINEChange;
  ACT_FOLIO.OnChange := ACT_FOLIOChange;
  HValComboBoxTri.OnChange := HValComboBoxTriChange;
end;

Procedure TFActivite.DimCalendrierSuivantEcran ;
Var DefRect: TDefRect;
begin
iHauteurEcran:=GetSystemMetrics(SM_CYSCREEN);
iLargeurEcran:=GetSystemMetrics(SM_CXSCREEN);
// Dimensionne le calendrier suivant la taille de l'écran
if (iLargeurEcran > 1000) then
   begin
   GCalendrier.width := 430;
   GCalendrier.DefaultColWidth := 52;
   end
else
   begin
   GCalendrier.width := 299;
   GCalendrier.DefaultColWidth := 40;
   end;

// Création et définition du grid calendrier
DefRect.NbLigneInRect := 1;
DefRect.NbColInRect := 2;
DefRect.IsBold := False;
DefRect.Cadrage := taLeftJustify;
GrCalendrier:=TGridCalendrierActivite.Create(GCalendrier, self, 1,6, True,
                True,true,True, EncodeDate(CleAct.Annee,CleAct.Mois,CleAct.Jour), LblMois,DefRect);
end;

Procedure TFActivite.DimensionneSuivantContexte ;
begin
// Dimensionne le Grid de saisie suivant la taille de l'écran, le nb de colonnes visibles pour utiliser tout
// l'espace dans le grid
HMTrad.ResizeGridColumns(GS);

// Taille des zones TOTAUX
DessineTotaux;

// Positionnement du champ libellé de l'affaire suivant la visibilité et la position des champs de saisie de l'affaire
if (ACT_AVENANT.visible) then
      LibelleAffaire.Left := ACT_AVENANT.Left + ACT_AVENANT.Width + 4
else
if (ACT_AFFAIRE3.visible) then
      LibelleAffaire.Left := ACT_AFFAIRE3.Left + ACT_AFFAIRE3.Width + 4
else
if (ACT_AFFAIRE2.visible) then
      LibelleAffaire.Left := ACT_AFFAIRE2.Left + ACT_AFFAIRE2.Width + 4
else
if (ACT_AFFAIRE1.visible) then
      LibelleAffaire.Left := ACT_AFFAIRE1.Left + ACT_AFFAIRE1.Width + 4;

// Positionnement du libellé du mois courant du calendrier suivant la position et la dimension de ce dernier
//LblMois.Left := GCalendrier.Left;
//LblMois.width := GCalendrier.width;

// Positionnement des boutons de changement de mois sur le calendrier
//ResizeBoutonsMois;
end;


Procedure TFActivite.GereEnabled (ARow : integer);
BEGIN
  BZoomArticle.Enabled := (GetChampLigne ('ACT_ARTICLE', ARow) <> '');
  GereAssistantEnabled;
  GereAffaireEnabled;
  BZoomTiers.Enabled := (GetChampLigne ('ACT_TIERS', ARow) <> '');

  BZoomPiece.Enabled := (GetChampLigne ('ACT_NUMPIECE', ARow) <> '');

  BZoomPieceach.Enabled := (GetChampLigne ('ACT_NUMPIECEACH', ARow) <> '');
  BZoomLigneAch.Enabled := (GetChampLigne ('ACT_NUMPIECEACH', ARow) <> '');
END ;

function TFActivite.GereSaisieEnabled (AcbAvecMess : boolean) : boolean;
begin
  Result := true;
  if (CleAct.TypeSaisie = tsaRess) then
    begin
      if ((ACT_RESSOURCE.Text = '') or (ACT_RESSOURCE.Text <> TOBAssistant.GetValue ('ARS_RESSOURCE')) and (ACT_RESSOURCE.CanFocus))
          and (CleAct.TypeAcces <> tacFourn)
        then
          begin
            if AcbAvecMess then
              HActivite.Execute (5, Caption, '');
            ACT_RESSOURCE.SetFocus;
            Result := false;
          end ;
    end
  else
    begin
      if ((ACT_AFFAIRE.Text = '') or (ACT_AFFAIRE.Text <> TOBAffaire.GetValue ('AFF_AFFAIRE')) and (ACT_AFFAIRE1.CanFocus)) then
        begin
          if AcbAvecMess then
            HActivite.Execute (9, Caption, '');
          ACT_AFFAIRE1.SetFocus;
          Result := false;
        end;
    end;
end;

procedure TFActivite.AfficheSuivantCriteres (AcdDateDebCal : TDateTime; AcbTotPied : boolean; AcbMajTotCal : boolean);
  var
  OldSyn : boolean;
  dDateAAfficher : TDateTime;
begin
  // Verouille tout affichage sur le calendrier
  OldSyn := GCalendrier.SynEnabled;
  GCalendrier.SynEnabled := false;
  GCalendrier.Enabled := false;
  try

    // Affiche les lignes correspondantes à la ressources affichée
    AfficheNouvellesLignesActivite;

    if AcbMajTotCal then
       begin
    //   bTobNbHeuresJourMaj:=false;
       // Reinitialise à nil tous les liens objets sur le calendrier
       GrCalendrier.VideObjetsCalendrier;

       // Décision de l'intervalle de date à afficher dans le calendrier
       dDateAAfficher := CalendrierAAfficher ( AcdDateDebCal );

       // Recalcul des totaux de temps affichés sur le calendrier
       MajTotauxCalendrier (dDateAAfficher);
       end;

    // Met à jour les totaux des lignes d'activité + pied (sommes en pied de grid)
    if AcbTotPied then
       MajApresModifActivite;

    // On se positionne sur le premier enregistrement, première colonne
    GS.Cells [SG_Curseur, GS.Row] := '';
    GS.Row := 1;
    GS.Col := SG_Jour;
    EntreDansLigneCourante (GS.Col, GS.Row);


  finally
    CleActOld := CleAct;
    // Reactive l'affichage du calendrier
    GS.TopRow := 1;
    GCalendrier.SynEnabled := OldSyn;
    GCalendrier.Enabled := true;

    // Mise à jour du libellé indiquant l'intervalle des dates
    // pour la semaine en cours
    MajIntervalleSemaine;

    // retaille les boutons des Mois par rapport au calendrier
    ResizeBoutonsMois;

    bActModifiee := false;
  end;
end;

procedure TFActivite.GotoEntete;
begin
  if (ACT_RESSOURCE.CanFocus) and (ACT_RESSOURCE.ReadOnly = false) then
    begin
      ACT_RESSOURCE.SetFocus;
      Exit;
    end
  else
    if (ACT_ANNEE.CanFocus) then
      ACT_ANNEE.SetFocus;

  if ACT_TIERS.CanFocus then
    begin
      ACT_TIERS.SetFocus;
      Exit;
    end;

  if USER.CanFocus then
    begin
      USER.SetFocus;
      Exit;
    end;
end;

procedure TFActivite.GotoPied;
begin
  if ACT_CODEARTICLE.CanFocus then
    begin
      ACT_CODEARTICLE.SetFocus;
      Exit;
    end;
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
// Actions sur les lignes
procedure TFActivite.CopierSelected (bAvecNumUnique, bAvecDeselection : boolean);
  Var
  i : integer;
  TOBL, TOBC : TOB;
begin
  TOBCXV.ClearDetail;

  try
  SourisSablier;

  for i := 1 to GS.RowCount - 1 do
    if GS.isSelected (i) then
      begin
        if Not EstRempliGCActivite (GS, i, false) then
          Continue;

        TOBL := GetTOBLigne (i);
        if TOBL = Nil then
          Continue;
        TOBC := TOB.Create ('', TOBCXV, -1);
        TOBC.Dupliquer (TOBL, True, True, False);
        if Not bAvecNumUnique then
          // on ne conserve pas le numéro unique
          TOBC.PutValue ('ACT_NUMLIGNEUNIQUE', 0);
      end;

  finally
    if bAvecDeselection then
      begin
        GS.ClearSelected;
        bSelectAll.Down := false;
      end;
    SourisNormale;
  end;
end ;

procedure TFActivite.CopierLignes;
begin
  if Action = taConsult then Exit;
  if GS.NbSelected <= 0 then Exit;
  CopierSelected (false, true);
  ClipBoard.Clear;
end ;

procedure TFActivite.CouperLignes;
  Var
  i : integer;
  TOBL : TOB;
begin
  if Action = taConsult then Exit;
  if GS.NbSelected <= 0 then Exit;

  // message d'alerte pour prévenir de la suppression en base
  if (PGIAskAf (HTitres.Mess[30], Caption) <> mrYes) then exit;

  CopierSelected (true, false);

  try
    SourisSablier;

    for i := GS.RowCount - 1 downto 1 do
      if GS.isSelected (i) then
        begin
          TOBL := GetTOBLigne (i);
          if (TOBL <> nil) then
            begin
              // on convertit pour être sûr de déduire la bonne quantité lors de la destruction de
              // la ligne
              TOBLigneActivite.Dupliquer (TOBL, TRUE, TRUE, TRUE);
              ConversionTobActivite (TOBEnteteLigneActivite, 'ACT_DATEACTIVITE');

              if not ClickDel (i, true) then break;
            end;
        end;

    if (GS.RowCount < NbRowsInit) then
      GS.RowCount := NbRowsInit;

    // Deselection des lignes à l'affichage
    GS.ClearSelected;
    bSelectAll.Down := false;

//    NumeroteLignes (GS, TOBActivite);

    EntreDansLigneCourante (GS.Col, GS.Row);
    GS.MontreEdit;
    GS.SynEnabled := True;
    GS.InplaceEditor.Deselect;

    MajTotauxActivite;
    ClipBoard.Clear;
  finally
    SourisNormale;
  end;
END ;

procedure TFActivite.CollerLignes;
  Var
  i : integer;
  Annee, Mois, Jour : word;
  TOBC, TOBL : TOB;
  Cancel : boolean;
  sTypeArt : string;
  DateCourante, PremierJour : TDateTime;
  iJourSemaine, YY : integer;
  RepCalend : integer;
BEGIN
  //if Not GS.Focused then exit;
  try
    SourisSablier;

    if Action = taConsult then Exit;
    if TOBCXV.Detail.Count <= 0 then Exit;

    // commencer par valider la ligne en cours de saisie
    GSExit (GS);

    // ensuite on peut insérer une ligne si tout est ok
    if (GCanClose = false) then exit;

    ClipBoard.Clear;
    Cancel := False;

    GSRowEnter (Nil, GS.Row + TOBCXV.Detail.Count, Cancel, False);

    for i := TOBCXV.Detail.Count - 1 downto 0 do
      BEGIN
        TOBC := TOBCXV.Detail [i];
        ClickInsert (GS.Row, true);
        TOBL := GetTOBLigne (GS.Row);
        TOBL.Dupliquer (TOBC, True, True);

        if (VH_GC.AFTYPESAISIEACT = 'SEM') then
          begin
            iJourSemaine := DayOfWeek (TOBL.GetValue ('ACT_DATEACTIVITE')) - 2;

            {$IFDEF AGLINF545}
            PremierJour := PremierJourSemaineTempo (CleAct.Semaine, CleAct.Annee);
            {$ELSE}
            // PL le 23/05/02 : réparation de la fonction  AGL550
            YY := CleAct.Annee;
            if (CleAct.Mois = 12) and (CleAct.Semaine = 1) then YY := YY + 1;
            if (CleAct.Mois = 1) and ((CleAct.Semaine = 52) or (CleAct.Semaine = 53)) then YY := YY - 1;
            PremierJour := PremierJourSemaine (CleAct.Semaine, YY);
            //////////////// AGL550
            {$ENDIF}

            if (iJourSemaine = -1) then iJourSemaine := 6;
            DateCourante := PremierJour + iJourSemaine;
          end
        else
          begin
            DecodeDate (TOBL.GetValue ('ACT_DATEACTIVITE'), Annee, Mois, Jour);
            DateCourante := EncodeDate (CleAct.Annee, CleAct.Mois, Jour);
          end;

        TOBL.PutValue ('ACT_DATEACTIVITE', DateCourante);
        TOBL.PutValue ('ACT_PERIODE', GetPeriode (DateCourante));
        TOBL.PutValue ('ACT_SEMAINE', NumSemaine (DateCourante));

        // PL le 13/12/02 : on est censé laisser le ACTIVITEREPRIS de la ligne d'origine FIDEA
        // PL le 14/10/03 : on affine le test : on ne modifie la nouvelle ligne à "facturable" que quand on était "facturé"
        if (TOBL.GetValue('ACT_ACTIVITEREPRIS') = 'FAC') then
          TOBL.PutValue('ACT_ACTIVITEREPRIS', 'F');
        ///////////////////////////
        TOBL.PutValue ('ACT_ACTORIGINE', 'SAI');
        TOBL.PutValue ('ACT_ETATVISA', 'ATT');
        TOBL.PutValue ('ACT_ETATVISAFAC', 'ATT');
        TOBL.PutValue ('ACT_VISEUR', V_PGI.User);
        TOBL.PutValue ('ACT_DATEVISA', NowH);
        TOBL.PutValue ('ACT_VISEURFAC', V_PGI.User);
        TOBL.PutValue ('ACT_DATEVISAFAC', NowH);

        TOBL.PutValue ('ACT_FOLIO', CleAct.Folio);
        // Init de la TOB copiée suivant le contexte par ressource ou par affaire
        if (CleAct.TypeSaisie = tsaRess) then
          begin
            TOBL.PutValue ('ACT_RESSOURCE', CleAct.Ressource);
            TOBL.PutValue ('ACT_TYPERESSOURCE', CleAct.TypeRess);
          end
        else
          begin
            TOBL.PutValue ('ACT_AFFAIRE', CleAct.Affaire);
            TOBL.PutValue ('ACT_AFFAIRE0', CleAct.Aff0);
            TOBL.PutValue ('ACT_AFFAIRE1', CleAct.Aff1);
            TOBL.PutValue ('ACT_AFFAIRE2', CleAct.Aff2);
            TOBL.PutValue ('ACT_AFFAIRE3', CleAct.Aff3);
            TOBL.PutValue ('ACT_AVENANT', CleAct.Avenant);
            TOBL.PutValue ('ACT_TIERS', CleAct.Tiers);
          end;

        if (CleAct.TypeAcces <> tacGlobal) then
          begin
            TOBL.PutValue ('ACT_TYPEACTIVITE', CleAct.TypeActivite);

            case CleAct.TypeAcces of
              tacTemps :  begin
                            sTypeArt := 'PRE';
                          end;
              tacFrais :  begin
                            sTypeArt := 'FRA';
                          end;
              tacFourn :  begin
                            sTypeArt := 'MAR';
                          end;
            end;

            if (TOBL.GetValue ('ACT_TYPEARTICLE') <> sTypeArt) then
              begin
                TOBL.PutValue ('ACT_TYPEARTICLE', sTypeArt);
                TOBL.PutValue ('ACT_ARTICLE', '');
                TOBL.PutValue ('ACT_CODEARTICLE', '');
                TOBL.PutValue ('ACT_FORMULEVAR', '');
                gbSaisieQtePermise := true;
                TOBL.PutValue ('ACT_UNITE', '');
                TOBL.PutValue ('ACT_UNITEFAC', '');
                TOBL.PutValue ('ACT_QTE', 0);
                TOBL.PutValue ('ACT_QTEFAC', 0);
                TOBL.PutValue ('ACT_QTEUNITEREF', 0);
                TOBL.PutValue ('ACT_LIBELLE', '');
                TOBL.PutValue ('ACT_PUPRCHARGE', 0);
                TOBL.PutValue ('ACT_PUPR',0);
                TOBL.PutValue ('ACT_PUPRCHINDIRECT', 0);
                TOBL.PutValue ('ACT_PUVENTE', 0);
                TOBL.PutValue ('ACT_PUVENTEDEV', 0);
                TOBL.PutValue ('ACT_TOTPRCHARGE', 0);
                TOBL.PutValue ('ACT_TOTPR', 0);
                TOBL.PutValue ('ACT_TOTPRCHINDI', 0);
                TOBL.PutValue ('ACT_TOTVENTE', 0);
                TOBL.PutValue ('ACT_TOTVENTEDEV', 0);
              end;
          end;

        MajTOBHeureCalendrier (GS.Row);
        AfficheLaTOBLigne (GS.Row);

        RepCalend := BlocageCalendrierLigne (TOBL, true);

        if (RepCalend < -100) then
          // Cas de blocage
          begin
            // la colonne de retour de modif est la quantité
            if (GS.ColLengths[SG_Qte] <> -1) and (gbSaisieQtePermise = true) then
              GS.Col := SG_Qte
            else
              GS.Col := SG_Article;

            // On réinitialise la quantité et les montants
            TOBL.PutValue ('ACT_QTE', 0);
            TOBL.PutValue ('ACT_QTEFAC', 0);
            TOBL.PutValue ('ACT_QTEUNITEREF', 0);
            TOBL.PutValue ('ACT_PUPRCHARGE', 0);
            TOBL.PutValue ('ACT_PUPR',0);
            TOBL.PutValue ('ACT_PUPRCHINDIRECT', 0);
            TOBL.PutValue ('ACT_PUVENTE', 0);
            TOBL.PutValue ('ACT_PUVENTEDEV', 0);
            TOBL.PutValue ('ACT_TOTPRCHARGE', 0);
            TOBL.PutValue ('ACT_TOTPR', 0);
            TOBL.PutValue ('ACT_TOTPRCHINDI', 0);
            TOBL.PutValue ('ACT_TOTVENTE', 0);
            TOBL.PutValue ('ACT_TOTVENTEDEV', 0);
            // On remet à jour le calendrier
            MajTOBHeureCalendrier (GS.Row);
            AfficheLaTOBLigne (GS.Row);

            break;
          end;

        MajNbHeureAnRessource (TOBL, false);

        GSRowExit (GS, GS.Row, Cancel, true);

        if (not Cancel and (TOBC.GetValue ('ACT_NUMLIGNEUNIQUE') <> 0)) then
          // On ne conserve plus la ligne après l'avoir collée dans le cas du couper
          TOBC.Free;

      END; // Boucle FOR

    EntreDansLigneCourante (GS.Col, GS.Row);

    MajTotauxActivite;

  finally
    GS.ClearSelected;
    bSelectAll.Down := false;
    GS.MontreEdit;
    GS.SynEnabled := True;
//    NumeroteLignes (GS, TOBActivite);
    SourisNormale;
  end;
END ;

function TFActivite.EntreDansLigneCourante (Col, Row : integer) : boolean;
  var
  ACol, ARow : integer;
  Chg, Cancel : boolean;
begin
  Chg := false;
  Cancel := false;
  ACol := Col;
  ARow := Row;

  try
    GSRowEnter (GS, ARow, Cancel, Chg);
    if (Not Cancel) then
      GSCellEnter (GS, ACol, ARow, Cancel);

  finally
    Result := Not Cancel;
  end;
end;

procedure TFActivite.CompleteLaTOB (TOBL : TOB);
  var
  Index : integer;
  ListeFonctionsRess : TStringList;
begin
  ListeFonctionsRess := nil;

  // On complete les donnees liees à la ressource
  Index := AFOAssistants.AddRessource (TOBL.GetValue ('ACT_RESSOURCE'));
  if (Index <> -1) and (Index <> -2) then
    begin
      try
        ListeFonctionsRess := TAFO_Ressource (AFOAssistants.GetRessource (Index)).FonctionDeLaRessource (TOBL.GetValue ('ACT_DATEACTIVITE'));
        if (ListeFonctionsRess <> nil) and (ListeFonctionsRess.Count <> 0) then
          TOBL.PutValue ('ACT_FONCTIONRES', ListeFonctionsRess [0]);
      finally
        ListeFonctionsRess.Free;
      end;
    end;

  // Mise à jour du dernier utilisateur ayant modifié la ligne et de l'heure
  TOBL.UpdateFieldsAgl;
//  TOBL.PutValue ('ACT_DATEMODIF', nowH);
//  TOBL.PutValue ('ACT_UTILISATEUR', V_PGI.User);
end;



// Renvoie 0 si au moins une des colonnes significative n'a pas été saisie
// Renvoie 1 si la ligne est valide
// Renvoie -1 si la ligne est invalide
function TFActivite.ControlLigneValide (AciLig : integer; TOBL : TOB) : integer;
  var
  StCellJour, mess : string;
  ColRetour, rep : integer;
  bBloqueSaisie : boolean;
  TOBArt : TOB;
  RepCalend : integer;
begin
  Result := 0;

  // L'une au-moins des colonnes est-elle saisie ?
  if EstRempliGCActivite (GS, AciLig) then
    begin
      Result := 1;
      // Si oui
      // Controle sur la saisie de la date
      StCellJour := GS.Cells [SG_Jour, AciLig];
      if (StCellJour = '') then
        begin
          PGIBoxAf (HTitres.Mess[8], Caption);
          GS.Col := SG_Jour;
          Result := -1;
          exit;
        end;

      // Controle de la cohérence du code affaire/ code tiers
      if (CleAct.TypeSaisie = tsaRess) then
        begin
          if (GS.ColLengths [SG_Affaire] = -1) then
            ACT_AFFAIRE.Text := TOBL.GetValue ('ACT_AFFAIRE') else ACT_AFFAIRE.Text := GS.Cells [SG_Affaire, AciLig];
          if (GS.ColLengths [SG_Aff0] = -1) then
            ACT_AFFAIRE0.Text := TOBL.GetValue ('ACT_AFFAIRE0') else ACT_AFFAIRE0.Text := GS.Cells [SG_Aff0, AciLig];
          if (GS.ColLengths [SG_Aff1] = -1) then
            ACT_AFFAIRE1.Text := TOBL.GetValue ('ACT_AFFAIRE1') else ACT_AFFAIRE1.Text := GS.Cells [SG_Aff1, AciLig];
          if (GS.ColLengths [SG_Aff2] = -1) then
            ACT_AFFAIRE2.Text := TOBL.GetValue ('ACT_AFFAIRE2') else ACT_AFFAIRE2.Text := GS.Cells [SG_Aff2, AciLig];
          if (GS.ColLengths [SG_Aff3] = -1) then
            ACT_AFFAIRE3.Text := TOBL.GetValue ('ACT_AFFAIRE3') else ACT_AFFAIRE3.Text := GS.Cells [SG_Aff3, AciLig];

          ACT_AVENANT.Text := TOBL.GetValue ('ACT_AVENANT');

          if (ACT_AFFAIRE1.Text = '') and (ACT_AFFAIRE2.Text = '') and (ACT_AFFAIRE3.Text = '') then
            begin
              ACT_AFFAIRE0.Text := '';
              ACT_AVENANT.Text := '';
            end;

          if (GS.ColLengths [SG_Tiers] = -1) or (GetParamSoc ('SO_AFRECHTIERSPARNOM') = true) then
            ACT_TIERS.Text := TOBL.GetValue ('ACT_TIERS')
          else
            ACT_TIERS.Text := GS.Cells [SG_Tiers, AciLig];


          if Not ((ACT_AFFAIRE1.Text = '') and (ACT_AFFAIRE2.Text = '') and (ACT_AFFAIRE3.Text = '')) then
            begin
              if Not ChargeAffaireActivite (true, false) then
                  begin
                  PGIBoxAf (HTitres.Mess [18], Caption);
                  if (GS.Col <> SG_Aff0) and (GS.Col <> SG_Aff1) and (GS.Col <> SG_Aff2) and (GS.Col <> SG_Aff3) then
                    GS.Col := SG_Tiers;
                  Result := -1;
                  exit;
                  end
              else
                  begin
                  TOBL.PutValue ('ACT_AFFAIRE', ACT_AFFAIRE.Text);

                  TOBL.PutValue ('ACT_AFFAIRE0',ACT_AFFAIRE0.Text) ;
                  TOBL.PutValue ('ACT_AFFAIRE1',ACT_AFFAIRE1.Text) ;
                  TOBL.PutValue ('ACT_AFFAIRE2',ACT_AFFAIRE2.Text) ;
                  TOBL.PutValue ('ACT_AFFAIRE3',ACT_AFFAIRE3.Text) ;
                  TOBL.PutValue ('ACT_AVENANT',ACT_AVENANT.Text) ;
                  TOBL.PutValue ('ACT_TIERS', ACT_TIERS.Text);


                  //AfficheLaTOBLigne(GS.Row) ;
                  AfficheLaTOBLigne (AciLig) ;
                  TOBL.PutEcran (Self, PPied) ;
                  PutEcranPPiedAvance (TOBL);
                  end;
            end
          else
            begin
              PGIBoxAf (HTitres.Mess[21], Caption);
              if (GS.Col <> SG_Aff0) and (GS.Col <> SG_Aff1) and (GS.Col <> SG_Aff2) and (GS.Col <> SG_Aff3) then
                GS.Col := SG_Aff1;
              Result := -1;
              exit;
            end;
        end;

      if (CleAct.TypeSaisie = tsaRess) then
        begin
          // En saisie par assistant, on verifie la date par rapport aux dates de la mission saisie et aux dates des paramètres
          // et blocage dus à la confidentialité
          rep := ControleDateDansIntervalle (TOBL.GetValue ('ACT_DATEACTIVITE'), false, true, true, bBloqueSaisie);
          if (rep <> 0) then
            // blocage
            begin
              mess := '';
              case rep of
                  2,3,4,5,6 : mess := 'La date choisie n''est pas dans l''intervalle des dates de saisie de l''activité ou de l''affaire.' + chr(10);
                  10 : mess := 'Blocage de confidentialité : l''affaire n''est pas accessible.' + chr(10);
                  11 : mess := 'Blocage sur l''état : l''affaire n''est pas accessible.' + chr(10);
                  12 : mess := 'Blocage sur la date : l''affaire n''est pas accessible.' + chr(10);
              end;

              // Si on veut annuler les modif de la ligne
              if (PGIAskAf (mess + HTitres.Mess[28], Caption) = mrYes) then
                if (TOBLigneActivite <> nil) then
                  begin
                    TOBL.Dupliquer (TOBLigneActivite, TRUE, TRUE, false);
                    TOBL.SetAllModifie (False);
                    AfficheLaTOBLigne (AciLig);
                  end;

              if (ctxScot in V_PGI.PGIContexte) then
                GS.Col := SG_Tiers
              else
                GS.Col := SG_Aff1;

              Result := -1;
              exit;
            end;
        end;

      if (CleAct.TypeSaisie = tsaClient) then
      // PL le 02/07/03 : en saisie par client/mission on test la validité de la ressource
        begin
          if BlocageLigne (TOBL, true, AciLig) then
            begin
              GS.Col := SG_Ressource;
              Result := -1;
              exit;
            end;
        end;

      // controle de la coherence generale de la ligne
      if (TOBL.GetValue ('ACT_AFFAIRE') = '') or (TOBL.GetValue ('ACT_TIERS') = '')
         or (TOBL.GetValue ('ACT_CODEARTICLE') = '') or (TOBL.GetValue ('ACT_TYPEARTICLE') = '')
         or ((TOBL.GetValue ('ACT_RESSOURCE') = '') and (CleAct.TypeAcces <> tacFourn))
         or ((ctxScot in V_PGI.PGIContexte) and (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') and (TOBL.GetValue ('ACT_QTE') = 0))
         or ((not (ctxScot in V_PGI.PGIContexte)) and (GS.Cells [SG_Qte, AciLig] = ''))
         then
        begin
          ColRetour := 1;
          if (TOBL.GetValue ('ACT_RESSOURCE') = '') then
            ColRetour := SG_Ressource;

          if (TOBL.GetValue ('ACT_TYPEARTICLE') = '') then
            ColRetour := SG_TypA
          else
          if ((CleAct.TypeAcces = tacGlobal) and (TOBL.GetValue ('ACT_TYPEARTICLE') = 'MAR')
                and (TOBL.GetValue ('ACT_RESSOURCE') = '')) then
            ColRetour := 0;

          if (TOBL.GetValue ('ACT_CODEARTICLE') = '') then
            ColRetour := SG_Article;
          if (TOBL.GetValue ('ACT_TIERS') = '') then
            ColRetour := SG_Tiers;
          if (TOBL.GetValue ('ACT_AFFAIRE') = '') then
            if (VH_GC.AFProposAct) then
              ColRetour := SG_Aff0
            else
              ColRetour := SG_Aff1;

          if ((ColRetour = 0) or (ColRetour = 1)) and (ctxScot in V_PGI.PGIContexte) then
            // Blocage d'une ligne quantité=0 pour une prestation seulement en scot
            if (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') and (TOBL.GetValue ('ACT_QTE') = 0) then
              begin
                ColRetour := SG_Qte;
              end;

          if ((ColRetour = 0) or (ColRetour = 1)) and not (ctxScot in V_PGI.PGIContexte) then
            // Blocage d'une ligne quantité vide seulement en GA
            if (GS.Cells [SG_Qte, AciLig] = '') then
              begin
                if (GS.ColLengths[SG_Qte] <> -1) and (gbSaisieQtePermise = true) then
                  ColRetour := SG_Qte
                else
                  ColRetour := SG_Article;
              end;

          if (ColRetour <> 0) then
            begin
              PGIBoxAf (HTitres.Mess[21], Caption);
              GS.Col := ColRetour;
              Result := -1;
              exit;
            end;
        end;

      // Controle de l'article pas fermé
      TOBArt := FindTOBArtSaisAff (TOBArticles, TOBL.GetValue ('ACT_ARTICLE'), true);
      if (TOBArt <> nil) then
        begin
          if (TOBArt.GetValue ('GA_FERME') = 'X') then
            begin
              PGIBoxAf (HTitres.Mess[29], Caption);
              GS.Col := SG_Article;
              Result := -1;
              exit;
            end;

          // Controles calendrier seulement si on saisie une prestation connue
          if (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') and (TOBL.GetValue ('ACT_ACTIVITEEFFECT') = 'X') and (TOBL.GetValue ('ACT_QTE') <> 0) then
            begin
              RepCalend := BlocageCalendrierLigne (TOBL, true);

              if (RepCalend < -100) then
                // Cas de blocage
                begin
                  if (GS.ColLengths[SG_Qte] <> -1) and (gbSaisieQtePermise = true) then
                    GS.Col := SG_Qte
                  else
                    GS.Col := SG_Article;

                  Result := -1;
                  exit;
                end;
            end;
        end;


      // Si tout est saisi et qu'il manque le facturable/NF, on met F par defaut
      if  (TOBL.GetValue ('ACT_ACTIVITEREPRIS') = '') then
        begin
          if (GS.Cells [SG_ActRep, AciLig] = 'N') then
            gbModifActRep := true;
          GS.Cells [SG_ActRep, AciLig] := 'F';
          TraiteActRep ( AciLig );
        end;

    end;

end;

// Actions sur les cellules
procedure TFActivite.FormateZoneSaisie (Control:TControl; ACol,ARow : Longint ) ;
  Var
  St, StC : String;
  sCodeType : string;
  bAffaire : boolean;
  Lng : integer;
BEGIN
  if (Control is THGrid) then
     St:=THGrid(Control).Cells[ACol,ARow]
  else
  if (Control is THCritMaskEdit) then
     St := THCritMaskEdit(Control).Text
  else
  if (Control is THNumEdit) then
     St := floattostr(THNumEdit(Control).Value);

  StC:=St;
  if (CleAct.TypeSaisie = tsaRess) then
     begin
     lng := 0;
     bAffaire:=true;
     if (ACol = SG_Aff1) then BEGIN sCodeType:=VH_GC.CleAffaire.Co1Type; lng := VH_GC.CleAffaire.Co1Lng; END
     else
     if (ACol = SG_Aff2) then BEGIN sCodeType:=VH_GC.CleAffaire.Co2Type; lng := VH_GC.CleAffaire.Co2Lng; END
     else
     if (ACol = SG_Aff3) then BEGIN sCodeType:=VH_GC.CleAffaire.Co3Type; lng := VH_GC.CleAffaire.Co3Lng; END
     else
        bAffaire:=false;
     if (St <> '') then
      BEGIN
        if bAffaire then
          if (sCodeType='CPT') then // compteur
            BEGIN
              if IsNumeric(St) then
               StC := Format('%*.*d', [lng, lng, StrToInt(St)])
             else
               StC:=AnsiUppercase(Trim(St)) ;
            //StC:=StrF00(Valeur(St),0)  // Voir fonction CompteurPartieAffaire...
            END
        else
           StC:=AnsiUppercase(Trim(St)) ;
      END;
     end;

  if (ACol in [SG_Ressource, SG_Affaire, SG_Tiers, SG_Article, SG_Unite, SG_TypA, SG_ActRep, SG_TypeHeure])
     then StC := AnsiUppercase (Trim (St)) else
   if (ACol in [SG_PUCH, SG_PV]) then StC := StrF00 (Valeur (St), V_PGI.OkDecP) else
   if (ACol in [SG_TOTPRCH, SG_TotV, SG_MntRemise, SG_MontantTVA]) then StC := StrF00 (Valeur (St), V_PGI.OkDecV) else
    if (ACol = SG_Qte)  then
        if not GetParamSoc('SO_AFVARIABLES') then       // PL le 14/08/03 pour problème arrondi ONYX
            StC := StrF00 (Valeur (St), V_PGI.OkDecQ) ;

  if (Control is THGrid) then
     THGrid(Control).Cells[ACol,ARow]:=StC
  else
  if (Control is THCritMaskEdit) then
     THCritMaskEdit(Control).Text:=StC
  else
  if (Control is THNumEdit) then
     if (StC<>'') then
     THNumEdit(Control).Value:=Valeur(StC);

END ;

Function TFActivite.GereElipsis (Control:TControl; LaCol : integer ) : boolean ;
var
//TOBL:TOB;
sTypeArt, sWhereTypArt:string;
BEGIN
//Cancel:=false;
sTypeArt := '';

Result:=False ;
// configuration de la selection suivant le type d'acces à la saisie, par temps, frais, fournitures, global
case CleAct.TypeAcces of
   tacTemps :begin
            sTypeArt := 'PRE';
            sWhereTypArt := 'CO_TYPE="TYA" AND CO_CODE="PRE"';
            end;
   tacFrais :begin
            sTypeArt := 'FRA';
            sWhereTypArt := 'CO_TYPE="TYA" AND CO_CODE="FRA"';
            end;
   tacFourn :begin
            sTypeArt := 'MAR';
            sWhereTypArt := 'CO_TYPE="TYA" AND CO_CODE="MAR"';
            end;
   tacGlobal:begin
//            TOBL:=GetTOBLigne(GS.Row) ;
//            sTypeArt := TOBL.GetValue('ACT_TYPEARTICLE'); PL le 11/07/03 : si on a changé à l'affichage, on doit prendre le contenu de la cellule et non pas le contenu de la TOB
            sTypeArt := GS.Cells[ SG_TypA, GS.Row];
            sWhereTypArt := 'CO_TYPE="TYA"';
            if (sTypeArt <> '') then
              // PL le 23/10/03 : gérait mal la sélection si on saisissait "P" et n'était pas adaptable si GA pour le type CTR
//               sWhereTypArt := sWhereTypArt + ' AND CO_CODE="'+sTypeArt+'"'  PlusTypeArticle (true)
               sWhereTypArt := sWhereTypArt + ' AND CO_CODE LIKE "'+sTypeArt+'%"' + PlusTypeArticle (true)
            else
//               sWhereTypArt := sWhereTypArt + ' AND (CO_CODE="PRE" OR CO_CODE="FRA" OR CO_CODE="MAR")';
               sWhereTypArt := sWhereTypArt + PlusTypeArticle (true);
            end;
end;

if (Control is THGrid) then
   begin
   if (LaCol<>SG_Article) and (LaCol<>SG_Ressource) and (LaCol<>SG_Unite) and (LaCol<>SG_TypeHeure) and (LaCol<>SG_ActRep)
   and (LaCol<>SG_TypA) and (LaCol<>SG_Aff0) and (LaCol<>SG_Aff1) and (LaCol<>SG_Aff2) and (LaCol<>SG_Aff3) and (LaCol<>SG_Tiers) then exit;
   end;

if (LaCol=SG_Article) then
   Result:=AFGetArticleRecherche(Control, sTypeArt, HTitres.Mess[1])
else
if (LaCol=SG_Ressource) then
   Result:=GetAssistantRecherche(Control, HTitres.Mess[11], true)
else
if (LaCol=SG_Unite) then
   Result:=GetUnite(Control, HTitres.Mess[13], CleAct.TypeAcces, sTypeArt)
else
if (LaCol=SG_TypeHeure) then
   Result:=GetTypeHeure(Control, HTitres.Mess[24])
else
if (LaCol=SG_ActRep) then
   // tablette AFTACTIVITEREPRISE
   Result:=LookupList(Control,HTitres.Mess[15],'COMMUN','CO_CODE','CO_LIBELLE','CO_TYPE="ARE" AND (CO_LIBRE="ACT" OR CO_LIBRE="ART")','CO_CODE',False,-1)
else
if (LaCol=SG_TypA) then
   // tablette GCTYPEARTICLE
   Result:=LookupList(Control,HTitres.Mess[17],'COMMUN','CO_CODE','CO_LIBELLE',sWhereTypArt,'CO_CODE',False,-1)
else
if (LaCol=SG_Aff0) then
   // tablette liée au champ ACT_AFFAIRE0
   Result:=LookUpPartiesAffaire( 0, GS)
else
if (LaCol=SG_Aff1) then
   // tablette liée au champ ACT_AFFAIRE1
   Result:=LookUpPartiesAffaire( 1, GS)
else
if (LaCol=SG_Aff2) then
   // tablette liée au champ ACT_AFFAIRE2
   Result:=LookUpPartiesAffaire( 2, GS)
else
if (LaCol=SG_Aff3) then
   // tablette liée au champ ACT_AFFAIRE3
   Result:=LookUpPartiesAffaire( 3, GS)
else
if (LaCol=SG_Tiers) then
   Result:=IdentifieTiers(Control)
;

END ;

procedure TFActivite.FindLigneFind(Sender: TObject);
var
OldRow, OldCol, NewCol, NewRow : integer;
begin
OldRow := GS.Row;
OldCol := GS.Col;
Rechercher(GS,FindLigne,FindFirst) ;
NewCol := GS.Col;
NewRow := GS.Row;

if Not ChangeDeCellule(OldCol, OldRow, NewCol, NewRow) then
   begin
   GS.Row := OldRow;
   GS.Col := OldCol;
   end;

end;

function TFActivite.ChangeDeCellule(OldCol, OldRow, NewCol, NewRow : integer):boolean;
var
Cancel, chg:boolean;
begin
Result := false;
Cancel := false; Chg := false;

if ((OldCol<>NewCol) or (OldRow<>NewRow)) then
   begin
   GSCellExit(GS, OldCol, OldRow, Cancel);
   if (Not Cancel) and (OldRow <> NewRow) then
      begin
      GSRowExit(GS, OldRow, Cancel, true);
      if Not Cancel then
         GSRowEnter(GS, NewRow, Cancel, Chg);
      end;

   if (Not Cancel) then
      GSCellEnter(GS, NewCol, NewRow, Cancel);

   if (Not Cancel) then
        Result:=true;
   end
else Result:=true;

end;

Procedure TFActivite.DeprotegePartieDroiteGS (ARow : integer);
  Var
  TOBL : TOB;
begin
  // Blocages sur le grid complet dus aux dates ou mode consultation
  if (GoRowSelect in GS.Options) then exit;

  // Blocages de la ligne dus à l'assistant
  TOBL:=GetTOBLigne(ARow) ;
  if (TOBL<>nil) then
      begin
      if (TOBL.GetValue('ACT_RESSOURCE')<>'') then
          begin
          if BlocageLigne (TOBL) then exit;
          end;
      end;


  //GS.ColLengths[SG_Unite]:= strtoint(lColsLength.Values[inttostr(SG_Unite)]);
  GS.ColLengths[SG_Qte]:= strtoint(lColsLength.Values[inttostr(SG_Qte)]);
  if (TOBL <> nil) then
  if (TOBL.GetValue('ACT_TYPEARTICLE')='PRE') then
     begin
     GS.ColLengths[SG_TypeHeure]:= strtoint(lColsLength.Values[inttostr(SG_TypeHeure)]);
     ACT_TYPEHEURE.ReadOnly:=false;
     end;
  GS.ColLengths[SG_ActRep]:= strtoint(lColsLength.Values[inttostr(SG_ActRep)]);
  GS.ColLengths[SG_PUCH]:= strtoint(lColsLength.Values[inttostr(SG_PUCH)]);
  GS.ColLengths[SG_PV]:= strtoint(lColsLength.Values[inttostr(SG_PV)]);
  GS.ColLengths[SG_TOTPRCH]:= strtoint(lColsLength.Values[inttostr(SG_TOTPRCH)]);
  GS.ColLengths[SG_TotV]:= strtoint(lColsLength.Values[inttostr(SG_TotV)]);
  if (SG_MntRemise <> -1) then
    GS.ColLengths[SG_MntRemise] := strtoint (lColsLength.Values[inttostr (SG_MntRemise)]);
  if (SG_MontantTVA <> -1) then
    GS.ColLengths[SG_MontantTVA] := strtoint (lColsLength.Values[inttostr (SG_MontantTVA)]);

      //ACT_UNITE.ReadOnly:=false;
  ACT_QTE.ReadOnly:=false;
  ACT_PUPRCHARGE.ReadOnly:=false;
  ACT_PUVENTE.ReadOnly:=false;
  ACT_TOTPRCHARGE.ReadOnly:=false;
  ACT_TOTVENTE.ReadOnly:=false;
  ACT_ACTIVITEREPRIS.ReadOnly:=false;
  ACT_ACTORIGINE.Enabled:=true;
  ACT_QTE_.ReadOnly:=false;
  ACT_PUPRCHARGE_.ReadOnly:=false;
  ACT_PUVENTE_.ReadOnly:=false;
  ACT_TOTPRCHARGE_.ReadOnly:=false;
  ACT_TOTVENTE_.ReadOnly:=false;
  ACT_ACTIVITEREPRIS_.ReadOnly:=false;

  // configure la protection des colonnes PR et PV suivant le type de valorisation
  ConfigSuivantTypeValo;
end;

Procedure TFActivite.ProtegePartieDroiteGS;
Var
TOBL:TOB;
begin
TOBL:=GetTOBLigne(GS.Row) ;
if (TOBL<>nil) then
    if BlocageLigne(TOBL) then exit;

//GS.ColLengths[SG_Unite]:= -1;
GS.ColLengths[SG_Qte]:= -1;
GS.ColLengths[SG_TypeHeure]:= -1;
GS.ColLengths[SG_ActRep]:= -1;
GS.ColLengths[SG_PUCH]:= -1;
GS.ColLengths[SG_PV]:= -1;
GS.ColLengths[SG_TOTPRCH]:= -1;
GS.ColLengths[SG_TotV]:= -1;
if (SG_MontantTVA <> -1) then
  GS.ColLengths[SG_MontantTVA]:= -1;
if (SG_MntRemise <> -1) then
  GS.ColLengths[SG_MntRemise]:= -1;
//ACT_UNITE.ReadOnly:=true;
ACT_QTE.ReadOnly:=true;
ACT_TYPEHEURE.ReadOnly:=true;
ACT_PUPRCHARGE.ReadOnly:=true;
ACT_PUVENTE.ReadOnly:=true;
ACT_TOTPRCHARGE.ReadOnly:=true;
ACT_TOTVENTE.ReadOnly:=true;
ACT_ACTIVITEREPRIS.ReadOnly:=true;
ACT_ACTORIGINE.Enabled:=false;
ACT_QTE_.ReadOnly:=true;
ACT_TYPEHEURE_.ReadOnly:=true;
ACT_PUPRCHARGE_.ReadOnly:=true;
ACT_PUVENTE_.ReadOnly:=true;
ACT_TOTPRCHARGE_.ReadOnly:=true;
ACT_TOTVENTE_.ReadOnly:=true;
ACT_ACTIVITEREPRIS_.ReadOnly:=true;
end;

function TFActivite.OKSaisiePartieDroiteGridGS(AciLigne:integer; AcbChangeSiOK:boolean):boolean;
var
affaire, ressource, tiers:string;
TOBL:TOB;
begin
Result:=false;
if (lColsLength=nil) then exit;
TOBL:=GetTOBLigne(AciLigne) ;
if (GS.ColLengths[SG_Affaire]=-1) and (TOBL<>nil) then
   affaire:=TOBL.GetValue('ACT_AFFAIRE') else affaire:=GS.Cells[SG_Affaire, AciLigne];
if ((GS.ColLengths[SG_Tiers]=-1) or (GetParamSoc('SO_AFRECHTIERSPARNOM')=true)) and (TOBL<>nil)  then
   Tiers:=TOBL.GetValue('ACT_TIERS') else Tiers:=GS.Cells[SG_Tiers, AciLigne];
if (GS.ColLengths[SG_Ressource]=-1) and (TOBL<>nil)  then
   Ressource:=TOBL.GetValue('ACT_RESSOURCE') else Ressource:=GS.Cells[SG_Ressource, AciLigne];

Result := (GS.Cells[SG_Jour,AciLigne]<>'')
            and (Tiers<>'')
            and (affaire<>'')
            // Soit la ressource n'est pas vide, soit on est en saisie de fourniture, soit on est en saisie globale et le type d'article est une fourniture
            and ((Ressource<>'') or (CleAct.TypeAcces=tacFourn) or ((CleAct.TypeAcces=tacGlobal) and (TOBL.GetValue('ACT_TYPEARTICLE')='MAR')))
            and (GS.Cells[SG_Article,AciLigne]<>'');


// Gestion des blocages
if (Result=true) then
    begin
    // Blocages sur le grid complet dus aux dates
    if (GoRowSelect in GS.Options) then Result:=false;
    // Blocages de la ligne dus à l'assistant
    if (TOBL<>nil) then
        if (Ressource<>'') then
        begin
        if BlocageLigne(TOBL) then Result:=false;
        end;
    end;

if AcbChangeSiOK then
   if Result then
      DeprotegePartieDroiteGS(AciLigne)
   else
      ProtegePartieDroiteGS;

end;

function TFActivite.ChargeTobSemaineAnalyse35H(AcsRessource, AcsFolio:string; AcdDateDebut, AcdDateFin : TDateTime) : TOB;
var
i:integer;
sReq:string;
QQ:TQuery;
begin
QQ := nil;
Result:=TOB.Create('Nb heures par semaine',Nil,-1) ;


Try
sReq := 'SELECT ACT_SEMAINE,SUM(ACT_QTE) AS ACT_QTE,ACT_UNITE,ACT_ACTIVITEEFFECT,ACT_ACTIVITEREPRIS FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                           +'" AND ACT_RESSOURCE="'+AcsRessource
                           +'" AND ACT_DATEACTIVITE>="'+UsDateTime(AcdDateDebut)
                           +'" AND ACT_DATEACTIVITE<"'+UsDateTime(AcdDateFin) + '"';


if (VH_GC.AFTYPESAISIEACT = 'FOL') and (AcsFolio <> '') then
   sReq := sReq + ' AND ACT_FOLIO="'+ AcsFolio + '"';

sReq := sReq + ' AND ACT_TYPEARTICLE="PRE"';

if (TOBListeConv<>nil) then
    if (TOBListeConv.Detail.Count<>0) then
      begin
      sReq := sReq + ' AND (ACT_UNITE="' + TOBListeConv.Detail[0].GetValue('GME_MESURE') + '"';

      for i:=1 to TOBListeConv.Detail.Count-1 do
         sReq := sReq + ' OR ACT_UNITE="' + TOBListeConv.Detail[i].GetValue('GME_MESURE')+ '"';

      sReq := sReq + ')' ;
      end;

sReq := sReq +' GROUP BY ACT_SEMAINE,ACT_UNITE,ACT_ACTIVITEEFFECT,ACT_ACTIVITEREPRIS ORDER BY ACT_SEMAINE' ;

QQ:=OpenSQL(sReq,true,-1,'',true);

if Not QQ.EOF then
    Result.LoadDetailDB('','','',QQ,false);

// Rajoute les champs sup pour stocker les lignes Fac et non Fac
// pas forcement utilisé ici, mais indispensable pour le bon fonctionnement de la fonction
if (Result.Detail.count<>0) then
    begin
    Result.Detail[0].AddChampSup('ACT_QTEFACTURE', true);
    Result.Detail[0].AddChampSup('ACT_QTENFACTURE', true);
    end;

// On remet toutes les occurrences de la TOB dans la même unité : en heure unite de conversion par defaut
ConversionTOBActivite( Result, 'ACT_SEMAINE' );

Finally
Ferme(QQ);
end;


end;

//
// bRecupereLimites à true : force la recuperation des limites dans les variables prévues à cet effet
// sinon on utilise celles passées en paramètres
function TFActivite.ControleSemaineGlissantes (AcsRessource, AcsFolio : string; bRecupereLimites : boolean;
                                            var AviNumSemErreur1, AviNumSemErreur2 : integer;
                                            var AvdNbHMaxSmnLim1 : double; var AviNbSemLim1 : integer; var AvbSemConsecLim1 : boolean;
                                            var AvdNbHMaxSmnLim2 : double; var AviNbSemLim2 : integer; var AvbSemConsecLim2 : boolean;
                                            var AvdNbHDepaceLim1, AvdNbHDepaceLim2 : double
                                            ) : T_EnsBlocages35H;
  var
  iNumSemainePrecedente, Index, i, j, iNbSemaineDepasse1, iNbSemaineDepasse2 : integer;
  dDerniereDateEffet, dDateDeRef : TDateTime;
  TobHS, TobCalRegle : TOB;
  dQteSurNSemaines1, dQteSurNSemaines2 : double;
begin
  Result := [];
  TobHS := nil;
  AvdNbHDepaceLim1 := 0; AvdNbHDepaceLim2 := 0;
  try
  //
  // Recuperation de la ressource courante
  //
  // La date de référence est la date de début d'analyse 35 Heures
  // elle sert à sélectionner le calendrier 35h avec la bonne date d'effet
  dDateDeRef:= GetParamSoc('SO_AFDATEDEB35');

  if bRecupereLimites then
      begin
      Index := AFOAssistants.AddRessource(AcsRessource, true);
      if (Index<>-1) and (Index<>-2) then
          begin
          AvdNbHMaxSmnLim1 := 0;
          AviNbSemLim1 := 0;
          AvbSemConsecLim1 := false;
          AvdNbHMaxSmnLim2 := 0;
          AviNbSemLim2 := 0;
          AvbSemConsecLim2 := false;
          TobCalRegle := TAFO_Ressource(AFOAssistants.GetRessource(Index)).TobCalendrierRegle;

          if (TobCalRegle<>nil) then
              begin
              // La Tob étant triée en ordre décroissant, on récupère la première date d'effet juste inférieure
              // à la date saisie et l'indice de la première TOB calendrier valable à la date d'effet
              dDerniereDateEffet:=TAFO_Ressource(AFOAssistants.GetRessource(Index)).DetDerniereDateEffetCalendrier(dDateDeRef, i);

              if (i<>-1) then
                  begin
                  // on boucle mais en principe on ne peut avoir qu'une seule règle par date d'effet
                  while (i<TobCalRegle.Detail.count) and (TobCalRegle.Detail[i].GetValue('ACG_DATEEFFET')=dDerniereDateEffet) do
                      begin
                      AvdNbHMaxSmnLim1 := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMNLIM1');
                      AviNbSemLim1 := TobCalRegle.Detail[i].GetValue('ACG_NBSEMLIM1');
                      AvbSemConsecLim1:=false;
                      if (TobCalRegle.Detail[i].GetValue('ACG_SEMCONSECLIM1')='X') then
                          AvbSemConsecLim1 := true;
                      AvdNbHMaxSmnLim2 := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMNLIM2');
                      AviNbSemLim2 := TobCalRegle.Detail[i].GetValue('ACG_NBSEMLIM2');
                      AvbSemConsecLim2:=false;
                      if (TobCalRegle.Detail[i].GetValue('ACG_SEMCONSECLIM2')='X') then
                          AvbSemConsecLim2 := true;
                      Inc(i);
                      end; // while
                  end; // if (i<TobCalRegle.Detail.count)
              end; // if (TobCalRegle<>nil)
          end; // if (Index<>-1) and (Index<>-2)
      end; // bRecupereLimites
  // inutile d'aller plus loin si l'on n'a pas les éléments de controles
  if ((AvdNbHMaxSmnLim1=0) or (AviNbSemLim1=0)) and ((AvdNbHMaxSmnLim2=0) or (AviNbSemLim2=0)) then exit;

  // Generation de la TOB par semaine sur la periode d'analyse des 35 heures
  TobHS := ChargeTobSemaineAnalyse35H(AcsRessource, AcsFolio, GetParamSoc('SO_AFDATEDEB35'), GetParamSoc('SO_AFDATEFIN35'));
  if (TobHS=nil) then exit;

  // Controle sur des semaines non consécutives
  AviNumSemErreur1:=0;
  AviNumSemErreur2:=0;
  if (Not AvbSemConsecLim1 and (AvdNbHMaxSmnLim1<>0) and (AviNbSemLim1<>0))
      or (Not AvbSemConsecLim2 and (AvdNbHMaxSmnLim2<>0) and (AviNbSemLim2<>0)) then
      begin
      iNbSemaineDepasse1:=0;
      iNbSemaineDepasse2:=0;
      for i:=0 to TOBHS.Detail.count-1 do
          begin
          if (TOBHS.Detail[i].GetValue('ACT_QTE')>AvdNbHMaxSmnLim1) then
              Inc(iNbSemaineDepasse1);
          if (TOBHS.Detail[i].GetValue('ACT_QTE')>AvdNbHMaxSmnLim2) then
              Inc(iNbSemaineDepasse2);
          end;

      if (Not AvbSemConsecLim1 and (AvdNbHMaxSmnLim1<>0) and (AviNbSemLim1<>0)) and (iNbSemaineDepasse1>AviNbSemLim1) then
          // la limite 1 est significative et la limite a été atteinte
          begin
          Result := Result + [tbhSemLim1];
          AviNumSemErreur1:=iNbSemaineDepasse1;
          end;
      if (Not AvbSemConsecLim2 and (AvdNbHMaxSmnLim2<>0) and (AviNbSemLim2<>0)) and (iNbSemaineDepasse2>AviNbSemLim2) then
          // la limite 2 est significative et la limite a été atteinte
          begin
          Result := Result + [tbhSemLim2];
          AviNumSemErreur2:=iNbSemaineDepasse2;
          end;
      end;


  // Controle sur des semaines consecutives
  if (AvbSemConsecLim1 and (AvdNbHMaxSmnLim1<>0) and (AviNbSemLim1<>0))
      or (AvbSemConsecLim2 and (AvdNbHMaxSmnLim2<>0) and (AviNbSemLim2<>0)) then
      begin
      for i := 0 to TOBHS.Detail.count - 1 do
          begin
          dQteSurNSemaines1 := 0;
          dQteSurNSemaines2 := 0;
          if (AvbSemConsecLim1 and (AvdNbHMaxSmnLim1 <> 0) and (AviNbSemLim1 <> 0)) and (AviNumSemErreur1 = 0) then
              begin
              iNumSemainePrecedente := 0;
              for j := 0 to AviNbSemLim1 - 1 do
                begin
                  if (i + j) > (TOBHS.Detail.count - 1) then
                    break;

                  // Si la semaine est bien consécutive à la précédente
                  if (iNumSemainePrecedente = 0) or (TOBHS.Detail[i+j].GetValue('ACT_SEMAINE') = (iNumSemainePrecedente + 1)) then
                    dQteSurNSemaines1 := dQteSurNSemaines1 + TOBHS.Detail[i+j].GetValue('ACT_QTE');

                  iNumSemainePrecedente := TOBHS.Detail[i+j].GetValue('ACT_SEMAINE');
                end;

              if ((dQteSurNSemaines1 / AviNbSemLim1) > AvdNbHMaxSmnLim1) then
                begin
                  AviNumSemErreur1 := TOBHS.Detail[i].GetValue('ACT_SEMAINE');
                  Result := Result + [tbhSemConsLim1];
                  AvdNbHDepaceLim1 := dQteSurNSemaines1 / AviNbSemLim1;
                end;
              end;

          if (AvbSemConsecLim2 and (AvdNbHMaxSmnLim2<>0) and (AviNbSemLim2<>0)) and (AviNumSemErreur2=0) then
              begin
              iNumSemainePrecedente := 0;
              for j := 0 to AviNbSemLim2 - 1 do
                begin
                  if (i + j) > (TOBHS.Detail.count - 1) then
                    break;

                  // Si la semaine est bien consécutive à la précédente
                  if (iNumSemainePrecedente = 0) or (TOBHS.Detail[i+j].GetValue('ACT_SEMAINE') = (iNumSemainePrecedente + 1)) then
                    dQteSurNSemaines2 := dQteSurNSemaines2 + TOBHS.Detail[i+j].GetValue('ACT_QTE');

                  iNumSemainePrecedente := TOBHS.Detail[i+j].GetValue('ACT_SEMAINE');
                end;

              if (dQteSurNSemaines2/AviNbSemLim2>AvdNbHMaxSmnLim2) then
                begin
                  AviNumSemErreur2 := TOBHS.Detail[i].GetValue('ACT_SEMAINE');
                  Result := Result + [tbhSemConsLim2];
                  AvdNbHDepaceLim2:=dQteSurNSemaines2/AviNbSemLim2;
                end;
              end;

          if  (Not AvbSemConsecLim1 or (AvdNbHMaxSmnLim1=0) or (AviNbSemLim1=0) or (AviNumSemErreur1<>0))
              and (Not AvbSemConsecLim2 or (AvdNbHMaxSmnLim2=0) or (AviNbSemLim2=0) or (AviNumSemErreur2<>0)) then
              // condition de sortie : l'une des limites n'est pas significative ou la limite a été atteinte
              break;

          end;
      end;

  finally
    TobHS.Free;
  end;
end;

// Il faudra changer la fonction au cas où les jours fériés ne seraient plus saisis dans la calendrier
// mais à recalculer dynamiquement
function TFActivite.BlocageCalendrierLigne(TOBLigne:TOB; AcbParle:boolean) : integer;
var
Index, i:integer;
TOBSemEncours, TobCalendrier, TobCalRegle, TOBQDuJour : TOB;
Col, Row : integer;
iJourSemaine:integer;
dQteSaisieJour,dSommeJour, dSommeSemaine, dSommeLigne, dSommeAn : double;
dDerniereDateEffet:TDateTime;
sMessage, sArgMess, sResCourante:string;
dDateSaisie:TDateTime;
bTestParAn:boolean;
begin
Result := 0;
dDerniereDateEffet := 0;
bTestParAn := false;
sMessage := ''; sArgMess := '';
Col := 0; Row := 0;
 
if bConsult.Down = true then exit;
if (CleAct.TypeSaisie <> tsaRess) or (TOBLigne = nil) then exit;

dQteSaisieJour := TOBLigne.GetValue('ACT_QTE');
sResCourante := TOBLigne.GetValue('ACT_RESSOURCE');
dDateSaisie := TOBLigne.GetValue('ACT_DATEACTIVITE');

GrCalendrier.CelluleDunJour (dDateSaisie, Col, Row);
TOBQDuJour := TOB(GCalendrier.Objects[Col,Row]);
TOBSemEncours :=TOBSommeSemaine.FindFirst (['SEMAINE'], [NumSemaine(dDateSaisie)], TRUE) ;
dSommeLigne := ConversionUnite (TOBLigne.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, dQteSaisieJour);
dSommeJour := dSommeLigne;

if (TOBQDuJour <> nil) then
    // Il y avait dejà une quantite saisie sur la ligne pour le même jour
    begin
    if (TOBLigneActivite <> nil) and (TOBQDuJour.GetValue('ACT_DATEACTIVITE') = TOBLigneActivite.GetValue('ACT_DATEACTIVITE')) then
        dSommeJour := TOBQDuJour.GetValue('ACT_QTE') + dSommeJour - ConversionUnite (TOBLigneActivite.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBLigneActivite.GetValue('ACT_QTE'))
    else
        dSommeJour := dSommeJour + TOBQDuJour.GetValue('ACT_QTE');
    end;


if (TOBSemEncours <> nil) then
    begin
    if (TOBLigneActivite <> nil) then
        dSommeSemaine := TOBSemEncours.GetValue('SOMME') + dSommeLigne - ConversionUnite (TOBLigneActivite.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBLigneActivite.GetValue('ACT_QTE'))
    else
        dSommeSemaine := TOBSemEncours.GetValue('SOMME') + dSommeLigne;
    end
else
    dSommeSemaine := dSommeJour;


// On complete les donnees liees à la ressource
Index := AFOAssistants.AddRessource (sResCourante, true);
if (Index <> -1) and (Index <> -2) then
    begin
    dSommeAn := TAFO_Ressource (AFOAssistants.GetRessource(Index)).gdTempsAnnuel;
    if (dDateSaisie >= GetParamSoc ('SO_AFDATEDEB35')) and (dDateSaisie <= GetParamSoc ('SO_AFDATEFIN35')) then
        begin
        bTestParAn := true;
        if (TOBLigneActivite <> nil) then
            dSommeAn := dSommeAn + dSommeLigne - ConversionUnite (TOBLigneActivite.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBLigneActivite.GetValue('ACT_QTE'))
        else
            dSommeAn := dSommeAn + dSommeLigne;
        end;

    TobCalendrier := TAFO_Ressource(AFOAssistants.GetRessource(Index)).TobCalendrier;
    TobCalRegle := TAFO_Ressource(AFOAssistants.GetRessource(Index)).TobCalendrierRegle;

    if (TobCalRegle <> nil) then
        begin
        // La Tob étant triée en ordre décroissant, on récupère la première date d'effet juste inférieure
        // à la date saisie et l'indice de la première TOB calendrier valable à la date d'effet
        dDerniereDateEffet := TAFO_Ressource(AFOAssistants.GetRessource(Index)).DetDerniereDateEffetCalendrier (dDateSaisie, i);

        if (i <> -1) then
            begin
            // on boucle mais en principe on ne peut avoir qu'une seule règle par date d'effet
            while (i < TobCalRegle.Detail.count) and (TobCalRegle.Detail[i].GetValue('ACG_DATEEFFET') = dDerniereDateEffet) do
                begin
                if (dDateSaisie >= TobCalRegle.Detail[i].GetValue('ACG_DATEEFFET')) then
                    // Ce qui est normalement forcement le cas
                    begin
                    if (GrCalendrier.JourFerie.TestJourFerie(dDateSaisie) and (TobCalRegle.Detail[i].GetValue('ACG_FERIEPERMIS')='-')) then
                        // Si c'est un jour férié et que ce n'est pas permis
                        begin
                        Result:=-111;
                        break;
                        end;
                    if ((DayOfWeek(dDateSaisie) = 1) and (TobCalRegle.Detail[i].GetValue('ACG_DIMPERMIS')='-')) then
                        // Si c'est un dimanche et que ce n'est pas permis
                        begin
                        Result:=-112;
                        break;
                        end;
                    if bTestParAn and (dSommeAn>TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANBLOC')) then
                        // Si le nombre d'heures saisies dans l'année est supérieur à la limite de blocage
                        begin
                        Result:=-115;
                        sArgMess := TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANBLOC');
                        break;
                        end;
                    if (dSommeJour>TobCalRegle.Detail[i].GetValue('ACG_NBHMAXJ')) then
                        // Si le nombre d'heures saisies dans la journée est supérieur à la limite
                        begin
                        Result:=-13;
                        sArgMess := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXJ');
                        if (TobCalRegle.Detail[i].GetValue('ACG_BLOCDEPHRJ')='X') then
                          begin
                          Result := Result-100;
                          break;
                          end;
                        end;
                    if (dSommeSemaine>TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMN')) then
                        // Si le nombre d'heures saisies dans la semaine est supérieur à la limite
                        begin
                        Result:=-14;
                        sArgMess := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMN');
                        if (TobCalRegle.Detail[i].GetValue('ACG_BLOCDEPHSMN')='X') then
                          begin
                          Result := Result-100;
                          break;
                          end;
                        end;
                    if bTestParAn and (dSommeAn>TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANALERT')) then
                        // Si le nombre d'heures saisies dans l'année est supérieur à la limite d'alerte
                        begin
                        Result:=-16;
                        sArgMess := TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANALERT');
                        break;
                        end;
                    end;

                Inc(i);
                end; // While
            end; // if (i<TobCalRegle.Detail.count)
        end; // if (TobCalRegle<>nil)

    if (TobCalendrier <> nil) and (Result > -100) then
      // Si aucun blocage n'a été détecté dans le calendrier règle
        begin
          for i := 0 to TobCalendrier.Detail.count - 1 do
          // sur toutes les lignes du calendrier lié à la ressource
            begin
              if (dDateSaisie = TobCalendrier.Detail[i].GetValue('ACA_DATE')) then
                // Si la date de saisie est la date courante du calendrier
                begin
                  if GrCalendrier.JourFerie.TestJourFerie (dDateSaisie) and (TobCalendrier.Detail[i].GetValue ('ACA_FERIETRAVAIL') = '-') then
                    // Si le jour férié n'est pas permis, on sort
                    begin
                      Result:=-101;
                      break;      
                    end;
                                            
                  if  (TobCalendrier.Detail[i].GetValue('ACA_JOUR') = 0) then
                    // Si on n'est pas sur un jour de semaine en particulier
                    if (TobCalendrier.Detail[i].GetValue('ACA_DUREE') < dSommeJour) then
                      // Si la durée de travail autorisée pour la journée est dépassée, on sort
                      begin
                        sArgMess := TobCalendrier.Detail[i].GetValue('ACA_DUREE');
                        Result := -2;
                        break;
                      end;
                end  // if (dDateSaisie = TobCalendrier.Detail[i].GetValue('ACA_DATE'))
              else
                // Si on est sur une condition liée au jour de la semaine : 'ACA_JOUR'<>0 et 'ACA_DATE'=iDate1900
                begin
                  iJourSemaine := DayOfWeek (dDateSaisie) - 1;
                  if (iJourSemaine = 0) then iJourSemaine := 7;
                  if ((TobCalendrier.Detail[i].GetValue('ACA_JOUR') = iJourSemaine)
                    and (TobCalendrier.Detail[i].GetValue('ACA_DATE') = iDate1900)) then
                      // Si la date saisie tombe le jour de la semaine courante du calendrier
//                    if (TobCalendrier.Detail[i].GetValue('ACA_DUREE')<>0)
//                        and (TobCalendrier.Detail[i].GetValue('ACA_DUREE')<dSommeJour) then
                      if (TobCalendrier.Detail[i].GetValue('ACA_DUREE') < dSommeJour) then
                        // Si la durée de travail autorisée pour la journée est dépassée, on sort
                        begin
                          sArgMess := TobCalendrier.Detail[i].GetValue('ACA_DUREE');
                          Result := -3;
                          break;
                        end;
                end;
            end; // For
        end; // if (TobCalendrier<>nil)
    end; // if (Index<>-1) and (Index<>-2) la ressource existe

if AcbParle and (Result<>0) then
    begin
    sMessage := 'Calendrier ' + RechDom ('AFTSTANDCALEN', TAFO_Ressource (AFOAssistants.GetRessource (Index)).tob_champs.GetValue('ARS_STANDCALEN'),false) + ' ';
    case Result of
        -1 : sMessage := sMessage + ': vous saisissez sur un jour férié non autorisé.';
        -2 : sMessage := sMessage + ': la durée autorisée ce jour est dépassée.' + chr(10) + 'Limite : '
                        + sArgMess + ' ' + RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
        -3 : sMessage := sMessage + ': la durée autorisée le ' + FormatDateTime ('dddd', dDateSaisie)
                        + ' est dépassée.' + chr(10) + 'Limite : ' + sArgMess + ' '
                        + RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
        -11,-111 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet) + ' '
                  + sMessage + ': Saisie sur un jour férié interdite.';
        -12,-112 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet) + ' '
                  + sMessage + ': Saisie sur un dimanche interdite.';
        -13,-113 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet) + ' '
                  + sMessage + ': la durée autorisée par jour est dépassée.'+ chr(10)+ 'Limite : '
                  + sArgMess + ' '+ RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
        -14,-114 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet) + ' '
                  + sMessage + ': la durée autorisée par semaine est dépassée.' + chr(10)+ 'Limite : '
                  + sArgMess + ' ' + RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
        -15,-115 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet) + ' '
                  + sMessage + ': la durée de blocage autorisée par an est dépassée.' + chr(10) + 'Limite : '
                  + sArgMess + ' '+ RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
        -16 : sMessage := 'Règle en date d''effet du ' + datetostr (dDerniereDateEffet)
                          + ' ' + sMessage + ': la durée d''alerte autorisée par an est dépassée.'
                          + chr(10) + 'Limite : ' + sArgMess + ' '
                          + RechDom ('GCQUALUNITTEMPS', VH_GC.AFMESUREACTIVITE, FALSE);
    end;

    PGIInfoAf (sMessage, Caption);
    end;
end;

{===========================================================================================}
{=============================== Evènements du Grid ========================================}
{===========================================================================================}
function TFActivite.GSRafraichitLaLigne (LaLigne : integer) : TOB;
  var
  TOBT, TOBLigneEnBase : TOB;
  LigAvant : integer;
begin
  Result := nil;
  LigAvant := GS.Row;
  TOBT := GetTOBLigne (GS.Row);
  if (TOBT = nil) then exit;
  if (TOBT.GetValue('ACT_TYPEACTIVITE') = '') then
    InitialiseLigneAct (TOBT, GS.Row, CleAct, Action, Dev.Code);

  if (TOBT.GetValue ('ACT_NUMLIGNEUNIQUE') = 0) then
    begin
      Result := TOBT;
      exit;
    end;

  TOBLigneEnBase := LaTobDeLaLigneUnique (TOBT.GetValue('ACT_TYPEACTIVITE'), TOBT.GetValue('ACT_AFFAIRE'),
                                  TOBT.GetValue('ACT_NUMLIGNEUNIQUE'));

  if (TOBLigneEnBase = nil) then
    // La ligne n'existe plus
    begin
    if (PGIAskAf ('La ligne courante a été supprimée par un autre utilisateur.' + chr(13) + 'Voulez-vous tout de même la conserver ?', Caption) = mrYes) then
      begin
      TOBT.SetAllModifie (True); // PL le 19/06/03 : on veut conserver la ligne donc la sauvegarder, on simule la modification
      Result := TOBT;
      end
    else
      begin
        RafraichirLaSelection;
        Result := nil;
        exit;
      end;
    end
  else
    begin
      if (TOBT.GetValue ('ACT_DATEMODIF') <> TOBLigneEnBase.GetValue ('ACT_DATEMODIF'))
      or ((GetParamSoc ('SO_AFVISAACTIVITE') = True) and (TOBT.GetValue ('ACT_DATEVISA') <> TOBLigneEnBase.GetValue ('ACT_DATEVISA')))
      or ((GetParamSoc ('SO_AFAPPPOINT') = True) and (TOBT.GetValue ('ACT_DATEVISAFAC') <> TOBLigneEnBase.GetValue ('ACT_DATEVISAFAC')))
        then
        // Si la date de modification n'a pas changé, inutile de la modifier
        // Sinon on alerte l'utilisateur qu'une modification est intervenue
        // et on bloque la modification de la ligne (il peut s'agir de la date de visa d'activité ou de facturation si elles sont gérées
        // manuellement)
        begin
          PGIInfoAf ('Attention, la ligne courante a été modifiée par un autre utilisateur.' + Chr(13)
                    + 'Votre sélection va être rafraîchie.' + Chr(13)
                    + 'Veuillez vous assurer que vous pouvez modifier cette sélection sans interaction avec d''autres utilisateurs.', Caption);

          RafraichirLaSelection;
        end;

      TOBLigneEnBase.free;

      if (LigAvant <> GS.Row) then // PL le 03/10/03 : bogue, en saisie à plusieurs si un user tiers modifie la ligne, on perdait la TOB
        Result := GetTOBLigne (GS.Row)
      else
        Result := TOBT;
    end;

end;


procedure TFActivite.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
dDateChargee:TDateTime;
PremierJourSem:TDateTime;
iAnnee,iMois,iJour:word;
Inc, iCol, iLigOrig, iLigne, YY : integer;
 bForcerCal:boolean;
TOBT:TOB;
begin
gbSaisieQtePermise := true;
bForcerCal:=true;
iAnnee := CleAct.Annee;
iMois := CleAct.Mois;
iJour := 1;
iLigOrig := LigOld;

TOBT := GSRafraichitLaLigne (GS.Row);
if (TOBT <> nil) then
  if GetParamSoc('SO_AFVARIABLES') and (TOBT.GetValue ('ACT_FORMULEVAR') <> '') then
    gbSaisieQtePermise := false;

///////////////////////////////////////////////////////////////////////////////
// Blocages de l'entree sur la ligne
//
//
//TOBT := GetTOBLigne(GS.Row);
if (TOBT<>nil) and Not(GoRowSelect in GS.Options) then
    // si pas de Blocage complet sur le grid du aux dates, on regarde s'il y a un blocage du à autre chose pour la ligne
    begin
    iLigne := GS.Row;
    if (iLigne < iLigOrig) and (iLigne>1) then Inc := -1 else Inc := 1;
    repeat
    TOBT := GetTOBLigne(iLigne);
    if BlocageLigne(TOBT) then
            begin
            iLigne:=iLigne+Inc;
            if (iLigne=0) then  begin Inc:=1; iLigne:=2; end;
            end;

    until (TOBT=nil) or Not (BlocageLigne(TOBT));
    GS.Row := iLigne;
    end ;
// Fin des blocages//////////////////////////////////////////////////


if GS.Row>=GS.RowCount-1 then GS.RowCount:=GS.RowCount+NbRowsPlus ;
if Action<>taConsult then CreerTOBLignes(GS.Row) ;

GS.Cells[SG_Curseur,GS.Row] := '4';

if (GS.Cells[SG_Jour,GS.Row]<>'') then
   begin
   iJour := strtoint(GS.Cells[SG_Jour,GS.Row]);

   if (VH_GC.AFTYPESAISIEACT='SEM') then
      begin
{      bDateValide:=true;
      try
      dDateChargee := EncodeDate(iAnnee, iMois, iJour) ;
      except
      bDateValide:=false;
      end;}
      TOBT := GetTOBLigne(GS.Row) ;
      dDateChargee := TOBT.GetValue('ACT_DATEACTIVITE') ;
      DecodeDate(dDateChargee, iAnnee, iMois, iJour);

      try

{$IFDEF AGLINF545}
      PremierJourSem:=PremierJourSemaineTempo(CleAct.Semaine, iAnnee);
{$ELSE}
// PL le 27/05/02 : Nouvelles fonctions AGL NumSemaine et PremierJourSemaine AGL550
      YY :=iAnnee;
      if (iMois=12) and (CleAct.Semaine=1) then YY := YY + 1;
      if (iMois=1) and ((CleAct.Semaine=52) or (CleAct.Semaine=53)) then YY := YY - 1;
      PremierJourSem:=PremierJourSemaine(CleAct.Semaine, YY);
//////////////////////// AGL550
{$ENDIF}

      if (dDateChargee<PremierJourSem) or (dDateChargee>PremierJourSem+6) then
         begin
         if (iJour>=15) then Inc:=-1 else Inc:=1;
         iMois := iMois+Inc;
         if (iMois=0) then begin iMois := 12; iAnnee:=iAnnee-1; end
         else if (iMois=13) then begin iMois := 1; iAnnee:=iAnnee+1; end;

         dDateChargee := EncodeDate(iAnnee,iMois,iJour);
         if (dDateChargee<PremierJourSem) or (dDateChargee>PremierJourSem+6) then
            begin
            raise ERangeError.Create('');
            end;
         end;
      except
        // PL : la gestion de l'exception dans ce cas de figure est sous controle mais pas très clean
        // c'est le moyen le plus rapide que j'ai trouvé pour traiter tous les cas...
        GS.Col := SG_Jour;
        GS.Cells[SG_Curseur,GS.Row] := '';
        PGIBoxAf(HTitres.Mess[9], Caption);
        Cancel := true;
        exit;
      end;
      end
   else
      begin
      if DaysPerMonth(iAnnee,iMois)<iJour then iJour:=DaysPerMonth(iAnnee,iMois);
//      dDateChargee := EncodeDate(iAnnee, iMois, iJour) ;
      end;
   bForcerCal:=ForcerRefreshCalendrier(iAnnee, iMois, iJour);
   end
else
   begin
   iJour := CleAct.Jour
   end;

// Positionne sur le jour courant dans le calendrier
if ChargerCalendrierEnCours( iJour, iMois, iAnnee,  bForcerCal) then
    // PL le 01/03/02 : pour éviter que le calendrier ne s'initialise à la date du jour
    // en gestion par semaine ce qui impliquait que la semaine était décalée
    if gbInitOK then
        MajApresChangeJourCalendrier;
    // Fin modif PL

if (Cancel = false) then
  begin
    gbModifLibArticle := false; // PL le 13/06/03 : on initialise le test de verif de modif du libelle
    // initialise l'acces aux colonnes suivant les colonnes saisies sur la premiere ligne
    iCol := GS.Col;
    OKSaisiePartieDroiteGridGS (GS.Row, true);
    while (iCol > SG_Jour) and (GS.ColLengths[iCol] = -1) do
      Dec (iCol);
    GS.Col := iCol;

    TOBValoZoom.Free; TOBValoZoom := nil;
    TOBValo.Free; TOBValo := nil;
    TOBFormuleVar.Free; TOBFormuleVar := Nil; // ONYX-AFORMULEVAR
    TOBFormuleVarQte.Free; TOBFormuleVarQte := Nil;

    TOBT := GetTOBLigne(GS.Row) ;
    if (GS.Row <> iLigOrig) then
    TOBLigneActivite.Dupliquer (TOBT, TRUE, TRUE, false);
    ConversionTobActivite (TOBEnteteLigneActivite, 'ACT_DATEACTIVITE');

    RemplirLesTOBCourantes (GS.Row);
    GereEnabled (GS.Row);
    GereDescriptif (GS.Row, True);

    // si la ligne a été créée à la volée (Ctrl M ou articles liés) on force la revalidation
    if TOBT.FieldExists ('VALIDATION') then
      if (TOBT.GetValue ('VALIDATION') = '-') then
        TOBT.SetAllModifie (true);
  end;

if (CleAct.TypeSaisie = tsaRess) then
  begin
    // Calcul de la somme par mois
    // A l'initialisation de l'écran de saisie et en saisie des frais et des fournitures
    // ou au changment de F/NF de la ligne
    // on initialise les sommes des heures par mois
//    if (gbInitOK = false) or (CleAct.TypeAcces = tacFrais) or (CleAct.TypeAcces = tacFourn)
//        or (gbModifMois = true) or (gbModifFolio = true) or (gbModifAnnee = true)
//        or (gbModifActRep = true) then
    gdNbHeureMois := CalculSommeMois (TOBNbHeureJourAffiche, gdNbHeureFac, gdNbHeureNFac);

    // Calcule somme des frais
    if (LblMoisTotaux.caption <> LblMois.Caption) then
      gdSommeFrais := CalculSommeFrais (ACT_FOLIO.Text);
    // affiche les totaux de temps le detail par mois
    AfficheTotauxParMois;
    LblMoisTotaux.caption := LblMois.Caption;
  end;
end;


procedure TFActivite.GSRowExit (Sender : TObject; Ou : Integer; var Cancel : Boolean; Chg : Boolean);
  var
  TOBL : TOB;
  bInitAff0, bJustValidate : boolean;
  RepControleLigne, ColRetour : integer;
begin
  ColRetour := -1;
  bJustValidate := false;
  if csDestroying in ComponentState then Exit;
  if Action = taConsult then Exit;

  DepileTOBLignes (Ou, GS.Row);
  TOBL := GetTOBLigne (Ou);
  if (TOBL = Nil) then
    begin
      GS.Cells [SG_Curseur, Ou] := '';
      Exit;
    end;

  if TOBL.IsOneModifie then
    begin
      bActModifiee := true;
      RepControleLigne := ControlLigneValide (Ou, TOBL);
      if (RepControleLigne = -1) then
        begin
          Cancel := true;
          // Vide la colonne curseur sur la nouvelle ligne
          GS.Cells [SG_Curseur,GS.Row] := '';

          // Au cas où la colonne vide sur laquelle on veut retourner le focus est protégée, on retourne sur la colonne Jour
          if (GS.ColLengths [GS.Col] = -1) then
            GS.Col := SG_Jour;

          // on force le repassage dans le cellexit si on n'a pas changé
          if (GS.Col = ColOld) then
            ColOld := SG_Jour;

          GS.Row := Ou;
          exit;
        end;

      // Ajustement des champs affaire : on vide le champ Aff0 dans le grid si aucun des autres champs de l'affaire n'a été rempli
      bInitAff0 := true;
      if (SG_Aff1 <> -1) then if (GS.Cells [SG_Aff1, Ou] <> '') then  bInitAff0 := false;
      if (SG_Aff2 <> -1) then if (GS.Cells [SG_Aff2, Ou] <> '') then  bInitAff0 := false;
      if (SG_Aff3 <> -1) then if (GS.Cells [SG_Aff3, Ou] <> '') then  bInitAff0 := false;
      if bInitAff0 and (SG_Aff0 <> -1) then
        GS.Cells[SG_Aff0, Ou] := '';

      // On génère les articles liés s'il y en a
      GereArtsLies (Ou, ColRetour); // PL le 16/10/03 : on ajoute la colonne de retour au cas où l'on veut forcer le retour sur la quantité

      // Ici on valide la ligne si la sortie est autorisée si la ligne est valide
      if (RepControleLigne = 1) then
        if (GereValiderLaLigne (Ou) <> oeOk) then
          begin
            Cancel := true;
            // Vide la colonne curseur sur la nouvelle ligne
            GS.Cells [SG_Curseur,GS.Row] := '';

            // Au cas où la colonne vide sur laquelle on veut retourner le focus est protégée, on retourne sur la colonne Jour
            if (GS.ColLengths [GS.Col] = -1) then GS.Col := SG_Jour;

            GS.Row := Ou;
            exit;
          end;

    end;

  // Vide la colonne curseur sur l'ancienne ligne courante
  GS.Cells[SG_Curseur, Ou] := '';

  if (Cancel = false) then
    begin
      // Pour les lignes créées à la volée (Ctrl M ou articles liés), on les marque comme validées
      if TOBL.FieldExists ('VALIDATION') then
        begin
          // PL le 16/10/03 : on veut ajouter les lignes ajoutées par Ctrl M ou articles liés dans le calendrier
          // seulement quand on les valide la première fois
          if (TOBL.GetValue ('VALIDATION') = '-') then
            bJustValidate := true;

          TOBL.PutValue ('VALIDATION', 'X');
        end;

      TOBL.SetAllModifie (false);

      if ((TOBLigneActivite = nil)
          or bJustValidate
          or (GS.Row > TOBActivite.detail.Count)
          or (TOBLigneActivite.getvalue ('ACT_NUMLIGNEUNIQUE') <> TOBActivite.detail[GS.Row - 1].getvalue ('ACT_NUMLIGNEUNIQUE'))
          or ((GS.Row < TOBActivite.detail.Count) and (TOBActivite.detail[GS.Row].getvalue ('ACT_NUMLIGNEUNIQUE') = 0))) then
        begin
          // Mise à jour des heures du calendrier
          MajTOBHeureCalendrier (Ou);

          MajNbHeureAnRessource (TOBL, false);
        end;

      // Deprotection des colonnes de la partie droite
      DeProtegePartieDroiteGS (GS.Row);

      // Si on a prévu une colonne de retour, on la positionne dans le grid
      if (ColRetour <> -1) then
        if (ColRetour = SG_Qte) and (GS.ColLengths[SG_Qte] <> -1) and (gbSaisieQtePermise = true) then
          GS.Col := SG_Qte;

    end;

end;

function TFActivite.MajNbHeureAnRessource (TOBL : TOB; bMoins : boolean) : double;
  var
  iIndiceRess : integer;
begin
  Result := 0;
  if (CleAct.TypeSaisie = tsaRess) then
    if (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') and (TOBL.GetValue ('ACT_ACTIVITEEFFECT') = 'X')
        and (TOBL.GetValue ('ACT_DATEACTIVITE') >= GetParamSoc ('SO_AFDATEDEB35'))
        and (TOBL.GetValue ('ACT_DATEACTIVITE') <= GetParamSoc ('SO_AFDATEFIN35')) then
      begin
        if (TOBL.GetValue ('ACT_UNITE') <> '') and (TOBL.GetValue ('ACT_QTE') <> 0) and (TOBL.GetValue ('ACT_RESSOURCE') <> '') then
          begin
            iIndiceRess := AFOAssistants.AddRessource (TOBL.GetValue ('ACT_RESSOURCE'));
            if (iIndiceRess <> -1) and (iIndiceRess <> -2) then
              begin
                Result := ConversionUnite (TOBL.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TOBL.GetValue ('ACT_QTE'));
                if bMoins then
                  Result := (-1) * Result;

                TAFO_Ressource (AFOAssistants.GetRessource (iIndiceRess)).gdTempsAnnuel :=
                            TAFO_Ressource (AFOAssistants.GetRessource (iIndiceRess)).gdTempsAnnuel +
                            Result;
              end;
          end;
      end;
end;

procedure TFActivite.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
bElipsisAffaire0,bElipsisAffaire1,bElipsisAffaire2,bElipsisAffaire3:boolean;
begin

if Action=taConsult then Exit ;
if Not Cancel then
   BEGIN
   StCellCur:=GS.Cells[GS.Col,GS.Row] ;
   bElipsisAffaire0 := (SG_Aff0<>-1) and ((GS.Col=SG_Aff0) and ACT_AFFAIRE0.ElipsisButton);
   bElipsisAffaire1 := (SG_Aff1<>-1) and ((GS.Col=SG_Aff1) and ACT_AFFAIRE1.ElipsisButton);
   bElipsisAffaire2 := (SG_Aff2<>-1) and ((GS.Col=SG_Aff2) and ACT_AFFAIRE2.ElipsisButton);
   bElipsisAffaire3 := (SG_Aff3<>-1) and ((GS.Col=SG_Aff3) and ACT_AFFAIRE3.ElipsisButton);

   GS.ElipsisButton:=(GS.Col=SG_Tiers) or (GS.Col=SG_Article) or (GS.Col=SG_Ressource) or (GS.Col=SG_Unite)  or (GS.Col=SG_ActRep)
                     or (GS.Col=SG_TypA) or bElipsisAffaire0 or bElipsisAffaire1 or bElipsisAffaire2 or bElipsisAffaire3 or (GS.Col=SG_TypeHeure);

   if GS.Col=SG_Article then GS.ElipsisHint:=HTitres.Mess[0] else
   if GS.Col=SG_Ressource then GS.ElipsisHint:=HTitres.Mess[10] else
   if GS.Col=SG_Unite then GS.ElipsisHint:=HTitres.Mess[12] else
   if GS.Col=SG_ActRep then GS.ElipsisHint:=HTitres.Mess[14] else
   if GS.Col=SG_TypA then GS.ElipsisHint:=HTitres.Mess[16] else
   if bElipsisAffaire0 then GS.ElipsisHint:=ACT_AFFAIRE0.Hint else
   if bElipsisAffaire1 then GS.ElipsisHint:=ACT_AFFAIRE1.Hint else
   if bElipsisAffaire2 then GS.ElipsisHint:=ACT_AFFAIRE2.Hint else
   if bElipsisAffaire3 then GS.ElipsisHint:=ACT_AFFAIRE3.Hint else
   if GS.Col=SG_Tiers then GS.ElipsisHint:=HTitres.Mess[22] else
   if GS.Col=SG_TypeHeure then GS.ElipsisHint:=HTitres.Mess[25] else
   ;

   // Affiche les données de la TOB dans le panel pied
   AfficheDetailTOBLigne(GS.Row) ;
   END ;

end;


procedure TFActivite.GSCellExit (Sender : TObject; var ACol, ARow : Integer; var Cancel : Boolean);
begin
  if csDestroying in ComponentState then Exit ;
  //if (ACol = ColOld) and (ARow = LigOld) and (GS.Cells[ACol,ARow]=StCellCur) then exit; // évite la réentrance, quand on change la cellule de destination, l'AGL passe deux fois ici...
  if (ACol = ColOld) and (ARow = LigOld) then exit; // évite la réentrance, quand on change la cellule de destination, l'AGL passe deux fois ici...

  if Action = taConsult then Exit;
  try
    GS.SynEnabled := false;
    if (GS.Cells [ACol, ARow] = StCellCur) then
      // Dans le cas ou l'on etait sur l'une des zones affaire, on traite quand meme car une des autres partie
      // a pu être changée et du coup modifier la cohérence du code complet...
      if  not  ((GS.Row = ARow)
        and (((ACol = SG_Aff0) or (ACol = SG_Aff1) or (ACol = SG_Aff2) or (ACol = SG_Aff3) or (ACol = SG_Avenant))
        or  (GetParamSoc('SO_AFVARIABLES') and (ACol = SG_Article) and (ACol < GS.Col)))) // PL le 08/09/03 : on veut que ça s'ouvre tout le temps pour ONYX
        // en onyx, si le code article est saisi et pas la quantite, on force le passage pour revalider les lignes générées en auto
        then
        Exit;

    FormateZoneSaisie (GS, ACol, ARow);
    if ACol = SG_Article    then TraiteArticle (ACol, ARow, Cancel) else
    if ACol = SG_Qte        then TraiteQte (ARow) else
    if ACol = SG_Jour       then TraiteJour (ARow, Cancel) else
    if ACol = SG_Tiers      then TraiteTiers (ARow) else
    if ACol = SG_Affaire    then TraiteAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Aff0       then TraitePartieAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Aff1       then TraitePartieAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Aff2       then TraitePartieAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Aff3       then TraitePartieAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Avenant    then TraitePartieAffaire (ACol, ARow, Cancel) else
//    if ACol = DerPartAff    then TraiteAffaire (ACol, ARow, Cancel) else
    if ACol = SG_Ressource  then TraiteRessource (ARow, Cancel) else
    if ACol = SG_PUCH       then TraitePUCH (ARow) else
    if ACol = SG_PV         then TraitePV (ARow) else
    if ACol = SG_TOTPRCH    then TraiteTOTPRCH (ARow) else
    if ACol = SG_TotV       then TraiteTotV (ARow) else
    if ACol = SG_Lib        then TraiteLib (ARow) else
    if ACol = SG_ActRep     then TraiteActRep (ARow) else
    if ACol = SG_TypA       then TraiteTypA (ARow) else
    //if ACol = SG_Unite      then TraiteUnite (ARow) else
    if ACol = SG_TypeHeure  then TraiteTypeHeure (ARow) else
    if ACol = SG_MntRemise  then TraiteMntRemise (ARow) else
    if ACol = SG_MontantTVA  then TraiteMontantTVA (ARow) else
    ;

  finally
    if not OKSaisiePartieDroiteGridGS (GS.Row, True) then
      begin
        if not ((ACol = SG_Jour) or (ACol = SG_Tiers) or (ACol = SG_Aff0) or (ACol = SG_Aff1) or (ACol = SG_Aff2) or (ACol = SG_Aff3)
            or (ACol = SG_Avenant)) and (GS.Row = ARow) and (GS.Cells[SG_Affaire, ARow] = '') then
//          GS.Col := DernierePartieAffaireVisible
//        else
          GS.Col := ProchaineColonneVide (GS.Col);
      end;

    if Not Cancel then
      begin
        ColOld := ACol;
        LigOld := ARow;
      end;

    GS.SynEnabled := true;
  end;
end;

procedure TFActivite.GSEnter (Sender : TObject);
begin
  gbSelectCellCalend := false;
  if Not GereSaisieEnabled ( true ) then exit;

  GS.ElipsisButton := true;

  EntreDansLigneCourante (GS.Col, GS.Row);
end;

procedure TFActivite.GSExit(Sender: TObject);
  var
  ACol, ARow : integer;
  Cancel : boolean;
begin
  if csDestroying in ComponentState then Exit;

  ACol := GS.Col;
  ARow := GS.Row;
  Cancel := false;
  GSCellExit (Sender, ACol, ARow, Cancel);
  if Not Cancel then
    GSRowExit (Sender, ARow, Cancel, false);

  if Cancel then
    begin
      GCanClose := false;
      GS.SetFocus;
      //Abort; PL le 05/12/01 : remplacé par un exit, car si on abort dans ce cas,
      // l'appli ne rend pas la main sur les menus et la barre d'outil...
      // Il faut dépiler la pile d'appel
      exit;
    end
  else
    GCanClose := true;

  ColOld := 0;
  LigOld := 0;
end;

procedure TFActivite.GSElipsisClick(Sender: TObject);
var
bElipsisAffaire0,bElipsisAffaire1,bElipsisAffaire2,bElipsisAffaire3:boolean;
Col, Row:integer;
Cancel:boolean;
begin
Col:=GS.Col;
Row:=GS.Row;
Cancel:=false;

bElipsisAffaire0 := (SG_Aff0<>-1) and ((GS.Col=SG_Aff0) and ACT_AFFAIRE0.ElipsisButton);
bElipsisAffaire1 := (SG_Aff1<>-1) and ((GS.Col=SG_Aff1) and ACT_AFFAIRE1.ElipsisButton);
bElipsisAffaire2 := (SG_Aff2<>-1) and ((GS.Col=SG_Aff2) and ACT_AFFAIRE2.ElipsisButton);
bElipsisAffaire3 := (SG_Aff3<>-1) and ((GS.Col=SG_Aff3) and ACT_AFFAIRE3.ElipsisButton);

if GS.Col=SG_Article then
   begin
   if (GS.Cells[SG_Article, GS.Row] <> '') then
    begin
      GSCellExit (GS, Col, Row, Cancel);
      if Cancel then exit;
//      if not Cancel then
//        GS.col := GS.col + 1;
    end;

   if (Col = GS.col) then
      // PL le 19/06/03 : si on est resté sur la même colonne c'est
      // qu'on n'a toujours pas choisi d'article, il faut donc relancer la recherche
      ZoomOuChoixArt (GS.Col, GS.Row)
   end else
if GS.Col=SG_Ressource then
   begin
   if (GS.Cells[SG_Ressource, GS.Row] <> '') then
      begin
        GSCellExit (GS, Col, Row, Cancel);
//        if not Cancel then
//          GS.col := GS.col + 1;
      end;

   if (Col = GS.col) then
      // PL le 19/06/03 : si on est resté sur la même colonne c'est
      // qu'on n'a toujours pas choisi de ressource, il faut donc relancer la recherche
      ZoomOuChoixRes (GS.Col, GS.Row)
   end else
if GS.Col=SG_Unite then IdentifieUnites(GS) else
if GS.Col=SG_ActRep then GereElipsis (GS, SG_ActRep) else
if GS.Col=SG_TypA then
    begin
    GSCellExit (GS, Col, Row, Cancel);
    if Cancel then exit;
    if (Col = GS.col) then
      // PL le 19/06/03 : si on est resté sur la même colonne c'est
      // qu'on n'a toujours pas choisi de type d'article, il faut donc relancer la recherche
      GereElipsis (GS, SG_TypA);
    end else
if GS.Col=SG_TypeHeure then
    begin
//    GSCellExit(GS, Col, Row, Cancel);
//    if Cancel then exit;
    GereElipsis (GS, SG_TypeHeure);
    end else
if bElipsisAffaire0 then GereElipsis (GS, SG_Aff0) else
if bElipsisAffaire1 then GereElipsis (GS, SG_Aff1) else
if bElipsisAffaire2 then GereElipsis (GS, SG_Aff2) else
if bElipsisAffaire3 then GereElipsis (GS, SG_Aff3) else
if GS.Col=SG_Tiers then
  begin
   if (GS.Cells[SG_Tiers, GS.Row] <> '') then
      begin
        GSCellExit (GS, Col, Row, Cancel);
//        if not Cancel then
//          GS.col := GS.col + 1;
      end;

   if (Col = GS.col) then
      // PL le 19/06/03 : si on est resté sur la même colonne c'est
      // qu'on n'a toujours pas choisi de tiers, il faut donc relancer la recherche
      ZoomOuChoixTiers (GS.Col, GS.Row);
//Result:=IdentifieTiers(Control)
//GereElipsis (GS, SG_Tiers)
  end else
  ;


ColOld := 0;
LigOld := 0;
end;

Procedure TFActivite.GetCellCanvas ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
  var
  TOBL, TOBArt, TOBT : TOB;
  Index : integer;
begin
  if (GS.ColLengths[ACol] = -1) and (ACol > SG_Desc)  then
    // Pour les colonnes que l'on a décalées à droite du grid, et qui sont non accessibles,
    // on change la couleur de la fonte de façon à ce que le contenu ne soit pas visible
    begin
      Canvas.Font.Color := canvas.brush.color;
    // GS.InplaceEditor.Brush.color := canvas.brush.color;
      exit;
    end;

  if (ARow = 0) then
  // si on est sur la ligne des titres
    begin
    // suivant la colonne courante, si le tri en cours correspond, on souligne le titre
      if ((ACol = SG_Jour) and  (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('JOU'))) or
        ((ACol = SG_Tiers) and (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('CLI'))) or
        (((ACol = SG_Affaire) or (ACol = SG_Aff1) or (ACol = SG_Aff0) or (ACol = SG_Aff2) or (ACol = SG_Aff3)) and (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('MIS'))) or
        ((ACol = SG_Article) and (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('PRE'))) or
        ((ACol = SG_Ressource) and (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('RES'))) or
        ((ACol = SG_TypA) and (HValComboBoxTri.ItemIndex = HValComboBoxTri.Values.IndexOf('TYP'))) then
          Canvas.Font.Style := Canvas.Font.Style + [fsUnderline];
    end
  else
  if (ACol = SG_Desc) then
    // Si on est sur le dessin de la colonne Descriptif
    begin
      Canvas.Font.Name := 'Wingdings';
      Canvas.Font.Size := 10;
      Canvas.Font.Color := clRed;
    end
  else
  if (ACol = SG_Curseur) then
    // Si on est sur le dessin de la colonne curseur
    begin
      Canvas.Font.Name := 'Marlett';
      Canvas.Font.Size := 14;
    end
  else
  if (ACol = SG_Aff2) then
    // Si on est sur le dessin de la colonne Aff2
    begin
      //   if (GS.Cells[ACol,ARow]<>'') then
      //    GS.Cells[ACol,ARow]:=Format('%-.4s-%.2s', [GS.Cells[ACol,ARow], copy(GS.Cells[ACol,ARow],5,2)]);
      //   GS.Cells[ACol,ARow]:=GS.Cells[ACol,ARow].EditText; GS.ColFormats[SG_Aff2]
    end
  else
  ;

  if not (gdFixed in AState) then
    begin
      TOBL := GetTOBLigne (ARow);
      if (TOBL <> nil) then
        begin
          // Italique sur la ligne visée
          if (GetParamSoc ('SO_AFVISAACTIVITE') = True) then
            begin
              if (TOBL.GetValue ('ACT_ETATVISA') = 'VIS') then
                canvas.Font.Style := canvas.Font.Style + [fsItalic]
              else
                canvas.Font.Style := canvas.Font.Style - [fsItalic];
            end
          else
            // Italique sur la ligne visée en facturation
          if (GetParamSoc ('SO_AFAPPPOINT') = True) then
            begin
              if (TOBL.GetValue ('ACT_ETATVISAFAC') = 'VIS') then
                canvas.Font.Style := canvas.Font.Style + [fsItalic]
              else
                canvas.Font.Style := canvas.Font.Style - [fsItalic];
            end
          else
          ;

          if (ACol = SG_Ressource) then
            if (TOBL.GetValue ('ACT_RESSOURCE') <> '') then
              begin
                Index := AFOAssistants.AddRessource (TOBL.GetValue ('ACT_RESSOURCE'));
                if (Index <> -1) and (Index <> -2) then
                  if (TAFO_Ressource (AFOAssistants.GetRessource (Index)).tob_Champs <> nil) then
                    if (TOB (TAFO_Ressource (AFOAssistants.GetRessource (Index)).tob_Champs).GetValue ('ARS_FERME')='X') then
                      Canvas.Font.Color := clRed;
              end;

          if (ACol = SG_Tiers) then
            if (TOBL.GetValue ('ACT_TIERS') <> '') then
              begin
                TOBT := TOBTierss.FindFirst (['T_TIERS', 'T_NATUREAUXI'], [TOBL.GetValue ('ACT_TIERS'), 'CLI'], false);
                if (TOBT <> nil) then
                  if (TOBT.GetValue ('T_FERME') = 'X') then
                    Canvas.Font.Color := clRed;
              end;

          if (ACol = SG_Article) then
            begin
              if (TOBL.GetValue ('ACT_ACTIVITEEFFECT') = '-') and (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
                begin
                  Canvas.Font.Color := clblue;
                end
              else
                begin
                  TOBArt := FindTOBArtSaisAff (TOBArticles, TOBL.GetValue ('ACT_ARTICLE'), true);
                  if (TOBArt <> nil) then
                    begin
                      if (TOBArt.GetValue ('GA_FERME') = 'X') then
                        Canvas.Font.Color := clRed;
                    end;
                end;
            end;
        end;
    end;

(* deux couleurs en fonction du jour courant
if not (gdFixed in AState) then
   begin
   canvas.brush.color := clWindow;
   GS.InplaceEditor.Brush.color := clWindow;

   if (HValComboBoxTri.Value='0') then
   // Si on est en mode tri par jour
      begin
      if (GS.Cells[SG_Jour,ARow]<>'') then
         if ((strtoint(GS.Cells[SG_Jour,ARow]) and $1)<>0) then
            begin
            canvas.brush.color := $00C9C9CB;
            end;

      if (gdFocused in AState) then
            GS.InplaceEditor.Brush.color := canvas.brush.color;
      end;
   end;
*)
end;

procedure TFActivite.GSBeforeFlip(Sender: TObject; ARow: Integer;
  var Cancel: Boolean);
begin
if bSelectAll.Down then
    if GS.isSelected(ARow) then
        bSelectAll.Down := false;

end;

procedure TFActivite.GSMouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
var
ACol, ARow:integer;
begin
GS.MouseToCell(X, Y, ACol, ARow);

if (ssCtrl in Shift) and (bSourisClickG=true) then
if (ARow<>iNewRow) then
if (iRowDebClick<>-1) then
    begin
    if (ARow=iOldRow) then GS.FlipSelection(iNewRow);

    GS.FlipSelection(ARow);
    iOldRow:=iNewRow;
    iNewRow:=ARow;
    end;
end;

procedure TFActivite.GSMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
ACol, ARow:integer;
begin
if (Button <> mbLeft) then exit;
bSourisClickG:=false;

GS.MouseToCell(X, Y, ACol, ARow);
if (ssCtrl in Shift) then
    begin
    GS.FlipSelection(ARow);
    iRowFinClick:=ARow;
    iNewRow:=-1;
    iOldRow:=-1;
    iRowDebClick:=-1;
    end;

DessineTotaux;
end;

procedure TFActivite.GSMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
ACol, ARow : integer;
index : integer;
begin
if (Button <> mbLeft) then exit;

GS.MouseToCell(X, Y, ACol, ARow);
if (ARow=0) then
// si on est sur la ligne des titres
   begin
   index := -1;
   // suivant la colonne sur laquelle on a cliqué, on tri
   if (ACol = SG_Jour) then index := HValComboBoxTri.Values.IndexOf('JOU') else
   if (ACol = SG_Tiers) then index := HValComboBoxTri.Values.IndexOf('CLI') else
   if (ACol = SG_Affaire) or (ACol = SG_Aff0) or (ACol = SG_Aff1) or (ACol = SG_Aff2) or (ACol = SG_Aff3) then index := HValComboBoxTri.Values.IndexOf('MIS') else
   if (ACol = SG_Article) then index := HValComboBoxTri.Values.IndexOf('PRE') else
   if (ACol = SG_Ressource) then index := HValComboBoxTri.Values.IndexOf('RES') else
   if (ACol = SG_TypA) then index := HValComboBoxTri.Values.IndexOf('TYP') ;

   if (index<>-1) then
      begin
      HValComboBoxTri.ItemIndex := index;
      HValComboBoxTriChange(HValComboBoxTri);
      end;
   end;

if (ARow<GS.FixedRows) then exit;

if (ACol<GS.FixedCols) then
    // si on a cliqué sur la colonne curseur
   begin
   if (ssCtrl in Shift) then
        begin
        iRowFinClick:=-1;
        iNewRow:=0;
        iOldRow:=-1;
        iRowDebClick:=ARow;
        bSourisClickG:=true;
        end;


   GS.SetFocus;
   ACol := SG_JOUR;
   end;

end;

procedure TFActivite.GSTopLeftChanged(Sender: TObject);
begin
DessineTotaux;
end;

procedure TFActivite.GereLeCtrlF5(Sender: TObject);
  var
  bElipsisAffaire0,bElipsisAffaire1,bElipsisAffaire2,bElipsisAffaire3,bElipsisAvenant,bTiers:boolean;
  Cancel,bChangeStatut,bProposition:boolean;
  TOBL, TOBValInter, TOBAffTempo : TOB;
  Col, Row, rep : integer;
  CodeAffaire,sTypeArt:string;
  Part0, Part1, Part2, Part3, Avenant : string;
  bBloqueSaisie : boolean;
begin
TOBValInter:=nil;
Part0 := ''; Part1 := ''; Part2 := ''; Part3 := ''; Avenant := '';
TOBL := GetTOBLigne(GS.Row) ;
if (TOBL=nil) then exit;
Cancel:=false;
bElipsisAffaire0 := (SG_Aff0<>-1) and (GS.Col=SG_Aff0) ;
bElipsisAffaire1 := (SG_Aff1<>-1) and (GS.Col=SG_Aff1) ;
bElipsisAffaire2 := (SG_Aff2<>-1) and (GS.Col=SG_Aff2) ;
bElipsisAffaire3 := (SG_Aff3<>-1) and (GS.Col=SG_Aff3) ;
bElipsisAvenant := (SG_Avenant<>-1) and (GS.Col=SG_Avenant) ;
bTiers := (SG_Tiers<>-1) and (GS.Col=SG_Tiers);
Col:=GS.Col;
Row:=GS.Row;

if bElipsisAffaire0 or bElipsisAffaire1 or bElipsisAffaire2 or bElipsisAffaire3 or bTiers or bElipsisAvenant then
   begin
   GSCellExit(Sender, Col, Row, Cancel);
   if Cancel then exit;

   Part0 := GS.Cells[SG_Aff0, Row];
   Part1 := GS.Cells[SG_Aff1, Row];
   Part2 := GS.Cells[SG_Aff2, Row];
   Part3 := GS.Cells[SG_Aff3, Row];
   if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
     Avenant := GS.Cells[SG_Avenant, Row];

   ACT_AFFAIRE.Text := TOBL.GetValue('ACT_AFFAIRE');
     if (ACT_AFFAIRE.Text = '') then
        ACT_AFFAIRE.Text := CodeAffaireRegroupe (Part0, Part1, Part2, Part3, Avenant, Action, false, false, false);

   ACT_AFFAIRE0.Text := Part0;
   ACT_AFFAIRE1.Text := Part1;
   ACT_AFFAIRE2.Text := Part2;
   ACT_AFFAIRE3.Text := Part3;
   ACT_TIERS.Text:=TOBL.GetValue('ACT_TIERS');
   if (VH_GC.AFProposAct) then bChangeStatut := true else bChangeStatut := False;
   if (ACT_AFFAIRE0.Text = 'P') then bProposition := True else bProposition := False;

   if GetAffaireEntete (ACT_AFFAIRE, ACT_AFFAIRE1, ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ACT_TIERS, bProposition, bChangeStatut, false, true, true,'SAT', true, true, true) then
      begin
      ChargeAffaireActivite(False, true);

      TOBAffTempo := TobAffaire;
      AFRemplirTOBAffaire (ACT_AFFAIRE.Text, TobAffaire, TobAffaires);
      // on verifie la date par rapport aux dates de la mission saisie et aux dates des paramètres
      // et blocage dus à la confidentialité
      if (GS.Cells[SG_Jour, Row] <> '') then
        begin
          rep := ControleDateDansIntervalle (TOBL.GetValue ('ACT_DATEACTIVITE'), true, false, true, bBloqueSaisie);
          if (rep <> 0) then
          // blocage
            begin
              TobAffaire := TOBAffTempo;
              exit;
            end;
        end;

//      GS.Cells[SG_Affaire,GS.Row] := ACT_AFFAIRE.Text;
      TOBL.PutValue('ACT_AFFAIRE', ACT_AFFAIRE.Text);

//      GS.Cells[SG_Aff1,GS.Row] := ACT_AFFAIRE1.Text;
//      GS.Cells[SG_Aff2,GS.Row] := ACT_AFFAIRE2.Text;
//      GS.Cells[SG_Aff3,GS.Row] := ACT_AFFAIRE3.Text;
//      GS.Cells[SG_Tiers,GS.Row] := ACT_TIERS.Text;
      TOBL.PutValue('ACT_AFFAIRE0',ACT_AFFAIRE0.Text) ;
      TOBL.PutValue('ACT_AFFAIRE1',ACT_AFFAIRE1.Text) ;
      TOBL.PutValue('ACT_AFFAIRE2',ACT_AFFAIRE2.Text) ;
      TOBL.PutValue('ACT_AFFAIRE3',ACT_AFFAIRE3.Text) ;
      TOBL.PutValue('ACT_AVENANT',ACT_AVENANT.Text) ;
      TOBL.PutValue('ACT_TIERS', ACT_TIERS.Text);

      AfficheLaTOBLigne(GS.Row) ;
      TOBL.PutEcran(Self, PPied) ;
      PutEcranPPiedAvance(TOBL);
      if (CleAct.TypeAcces = tacGlobal) then
        GS.Col := SG_TypA
      else
        GS.Col := SG_Article;
      end;
   end
else
if (GS.Col=SG_Article) then
   begin
   if (CleAct.TypeSaisie = tsaRess) then
      begin
//      CodeAffaire:=GS.Cells[SG_Affaire, GS.Row];
      CodeAffaire:=TOBL.GetValue('ACT_AFFAIRE');
      if ( CodeAffaire='' ) then
         begin
         PGIBoxAf(HTitres.Mess[20], Caption);
         exit;
         end;
      end
   else
      begin
      CodeAffaire:=CleAct.Affaire;
      end;

   sTypeArt:='';
   case CleAct.TypeAcces of
     tacTemps :  sTypeArt := 'PRE';
     tacFrais :  sTypeArt := 'FRA';
     tacFourn :  sTypeArt := 'MAR';
     tacGlobal : sTypeArt := GS.Cells[SG_TypA, GS.Row];

   end;

//   GSCellExit(Sender, Col, Row, Cancel);
//   if Cancel then exit;
   ZoomArticlesD1Affaire( CodeAffaire, GS.Cells[SG_Article, GS.Row], sTypeArt, ((Not AffichageValorisation) and ((CleAct.TypeAcces=tacTemps) or (CleAct.TypeAcces=tacGlobal))) ) ;

   if (TheTob <> nil) and (TheTOB.TOBNameExist('ARTICLE')) then
      begin
      //TOBValInter := TOB.Create('ARTICLE',Nil,-1) ;
      //TOBValInter.Dupliquer(TheTOB.Detail[0], true, true, false);
      TOBValInter := TheTOB.Detail[0];
      end;

   if (TOBValInter<>nil) then
      begin
      try
      if (TOBValoZoom<>nil) then
         begin
         TOBValoZoom.Free;
         TOBValoZoom:=nil;
         end;

      if (TOBValo<>nil) then
         begin
         TOBValo.Free;
         TOBValo:=nil;
         end;

      TOBValoZoom := TOB.Create('ARTICLE',Nil,-1) ;
      TOBValoZoom.Dupliquer(TOBValInter,true,true,false);
      GS.Cells[SG_Article, GS.Row]:= TOBValoZoom.GetValue('GA_CODEARTICLE');
      StCellCur:='';
      GSCellExit(Sender, Col, Row, Cancel);
      if Cancel then exit;
      GS.Col := SG_Lib;

      finally
      TOBValInter.Free;
      end;
      end;
   end
else
if (GS.Col = SG_Desc) then
  begin
    Afficherlemmo1Click(nil);
  end
else
if (GS.Col = SG_Qte) then
  begin
    if GetParamSoc ('SO_AFVARIABLES') then
      TraiteFormuleVar (Row, 'MODIF');
    GSCellExit (Sender, Col, Row, Cancel);
    if Cancel then exit;
  end
else
if (GS.Col = SG_Jour) then
  begin
  if GetParamSoc ('SO_AFVARIABLES') and (bConsult.Down = true) then
    begin
      TraiteFormuleVar (Row, 'CONSULT');
    end;
  end
;

end;

procedure TFActivite.GSGetEditMask(Sender: TObject; ACol, ARow: Integer; var Value: String);
begin
if (Acol=SG_Aff2) And (VH_GC.CleAffaire.Co2Type='SAI') And (VH_GC.AFFORMATEXER<>'AUC') then  // Formattage Exercice
   BEGIN
   Value := FormatZoneExercice(VH_GC.AFFORMATEXER);
   END;

end;

{==============================================================================================}
{=============================== Manipulation des TOB LIGNES ==================================}
{==============================================================================================}
function TFActivite.BlocageLigne (TOBL : TOB; bParle : boolean = false; ARow : integer = 0) : boolean;
  var
  Rep, Index, NumLigne : integer;
  MsgErreur : string;
  TOBAss : TOB;
  sRess : string;
  eOnSelectCell : TSelectCellEvent;

begin
  Result := false;

  if (TOBL = nil) then exit;

  NumLigne := GS.Row;
  if (ARow <> 0) then
    NumLigne := ARow;

  sRess := '';
  Rep := 0;

  try
//  if (TOBAss = nil) then
  if (CleAct.TypeSaisie = tsaClient) and (GS.Cells[SG_Ressource, NumLigne] <> '') then
    sRess := GS.Cells[SG_Ressource, NumLigne]
  else
  if (CleAct.TypeSaisie <> tsaClient) then
    sRess := TOBL.GetValue ('ACT_RESSOURCE');

  if (sRess <> '') then
    begin
      TOBAss := nil;
      Index := AFOAssistants.AddRessource (sRess);
      if (Index <> -1) and (Index <> -2) then
        TOBAss := TAFO_Ressource (AFOAssistants.GetRessource (Index)).tob_Champs;

      // Blocage sur l'assistant fermé
      if (TOBAss <> nil) then
      if (TOBAss.GetValue ('ARS_FERME') = 'X') then
        begin
          Rep := 1;
          if (CleAct.TypeSaisie = tsaClient) then
            begin
              eOnSelectCell := GS.OnSelectCell;
              GS.OnSelectCell := nil;
              GS.Col := SG_Ressource;
              GS.OnSelectCell := eOnSelectCell;
            end;
        end;
    end;

  if (TOBL <> nil) then
    begin
      // Blocage sur la ligne visée
      if (GetParamSoc ('SO_AFVISAACTIVITE') = True) then
        if (TOBL.GetValue ('ACT_ETATVISA') = 'VIS') then
          Rep := 2;

      // Blocage sur la ligne visée en facturation
      if (GetParamSoc ('SO_AFAPPPOINT') = True) then
        if (TOBL.GetValue ('ACT_ETATVISAFAC') = 'VIS') then
          Rep := 3;

      // Blocage sur la ligne FAC
      if (TOBL.GetValue ('ACT_ACTIVITEREPRIS') = 'FAC') then
        Rep := 4;
    end;

  if (Rep <> 0) and bParle then
    begin
      case Rep of
        1 : MsgErreur := 'L''assistant sélectionné est fermé.';
        2 : MsgErreur := 'La ligne est bloquée par le visa d''activité.';
        3 : MsgErreur := 'La ligne est bloquée par le visa de facturation.';
        4 : MsgErreur := 'La ligne est facturée.';
      end;

      PgiInfoAF (MsgErreur, Caption);
    end;

  finally
    if (Rep <> 0) then
      Result := true;
  end;
end;

procedure TFActivite.VideCodesLigne ( ARow : integer ) ;
Var TOBL : TOB ;
BEGIN

TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
TOBL.InitValeurs ;

END ;

Function TFActivite.GetTOBLigne ( ARow : integer ) : TOB ;
BEGIN
Result:=Nil ;
if (TOBActivite = nil) then exit;
if ((ARow<=0) or (ARow>TOBActivite.Detail.Count)) then Exit ;
Result:=TOBActivite.Detail[ARow-1] ;

END ;
(*
Procedure TFActivite.NumeroteLignes ( GS : THGrid ; TOBActivite : TOB ) ;
Var i, offset : integer ;
TOBL:TOB;
bModified:boolean;
BEGIN
offset:=0;
for i:=0 to TOBActivite.Detail.Count-1 do
    BEGIN
    TOBL :=GetTOBLigne(i+1);
    if (TOBL<>nil) then
      begin
      bModified:=TOBL.IsOneModifie;
//      if (TOBL.GetValue('ACT_TYPEARTICLE')='FRA') then offset:=1000;
//      if (TOBL.GetValue('ACT_TYPEARTICLE')='MAR') then offset:=2000;
      TOBL.PutValue('ACT_NUMLIGNE',offset+i+1) ;
      if (bModified=false) then TOBL.SetAllModifie(false);
      end;
    END ;
END ;
*)
procedure TFActivite.InitialiseLigne ( ARow : integer );
  var
  TOBL : TOB ;
begin
  TOBL := GetTOBLigne (ARow);
  InitialiseLigneAct ( TOBL, ARow , CleAct, Action, Dev.Code);
end;

Procedure TFActivite.InitAssistantDefaut (TOBL : TOB; ARow : integer; var AvbArtOK : boolean);
var
  Cancel : boolean;
  Index : integer;
  TOBArt : TOB;
  bArticleOK : boolean;
begin
  AvbArtOK := false;
  bArticleOK := false;
  TOBArt := nil;
  if (TOBL.GetValue ('ACT_RESSOURCE') = '') then exit;

  Index := AFOAssistants.AddRessource (CleAct.Ressource);
  if (Index <> -1) and (Index <> -2) then
    begin
      if (Trim (TAFO_Ressource(AFOAssistants.GetRessource (Index)).tob_Champs.GetValue('ARS_ARTICLE')) = '') then exit;
      //      TOBL.PutValue('ACT_ARTICLE', TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_Champs.GetValue('ARS_ARTICLE'));

        // Associe l'article par defaut à l'objet ressource courant si necessaire
      if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article = nil)
         or (TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_Champs.GetValue('ARS_ARTICLE') <> TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_ARTICLE')) then
        TAFO_Ressource(AFOAssistants.GetRessource(Index)).AssocieArticle ('');

      if (TOBL.GetValue('ACT_ARTICLE') <> '') then
        TOBArt := TOBArticles.FindFirst (['GA_ARTICLE'], [TOBL.GetValue('ACT_ARTICLE')], False);

      if (TOBArt = nil) and (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article <> nil) then
        begin
           // On controle que l'article trouve soit bien associable à la ligne d'activité saisie :
           // - l'article doit etre en rapport avec le type de saisie : une prestation si on saisie par prestation, etc..
           // - Si on est en saisie globale, on vérifie que le type saisi est bien celui de l'article
           case CleAct.TypeAcces of
              tacTemps :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE') = 'PRE') then
                    bArticleOK := true;
              tacFrais :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE') = 'FRA') then
                    bArticleOK := true;
              tacFourn :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE') = 'MAR') then
                    bArticleOK := true;
              tacGlobal :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE') = TOBL.GetValue('ACT_TYPEARTICLE')) then
                    bArticleOK := true;
           end;

           // On ajoute l'article dans la TOB générale
           if bArticleOK then
              begin
              TOBArt := CreerTOBArt (TOBArticles);
              TOBArt.Dupliquer (TAFO_Ressource (AFOAssistants.GetRessource (Index)).Article, true, true, false);
              end;
           end;

        if (TOBArt <> nil) then
           begin
           // On controle que l'article trouve soit bien associable à la ligne d'activité saisie :
           // - l'article doit etre en rapport avec le type de saisie : une prestation si on saisie par prestation, etc..
           // - Si on est en saisie globale, on vérifie que le type saisi est bien celui de l'article
           case CleAct.TypeAcces of
              tacTemps :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE')='PRE') then
                    bArticleOK := true;
              tacFrais :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE')='FRA') then
                    bArticleOK := true;
              tacFourn :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE')='MAR') then
                    bArticleOK := true;
              tacGlobal :
                 if (TAFO_Ressource(AFOAssistants.GetRessource(Index)).Article.GetValue('GA_TYPEARTICLE')=TOBL.GetValue('ACT_TYPEARTICLE')) then
                    bArticleOK := true;
           end;

           if bArticleOK then
              begin
              Cancel := false;
              GS.Cells[SG_Article, ARow] := TOBArt.GetValue ('GA_CODEARTICLE');
              StCellCur := '';
              // PL le 29/01/02 : problème de revalorisation alors que la mission n'etait pas renseignée
              // - On doit avoir saisi une mission pour pouvoir lancer la valorisation
              if (TOBL.GetValue ('ACT_AFFAIRE') <> '') and (TOBL.GetValue ('ACT_AFFAIRE1') <> '') then
                  GSCellExit (nil, SG_Article, ARow, Cancel)
              else
                  if (IdentifierArticle ( GS, SG_Article, ARow, Cancel, False)) then
                      if (TOBArt <> nil) then
                          ACT_ARTICLEL.Caption := TOBArt.GetValue('GA_LIBELLE')
                      else
                          ACT_ARTICLEL.Caption := '';
              // Fin PL le 29/01/02
              if Not Cancel then AvbArtOK := true;
              end;
           end;
        end;
end;

procedure TFActivite.AfficheLaTOBLigne ( ARow : integer ) ;
Var TOBL : TOB ;
    i    : integer ;
    iJour, iMois, iAnnee : word;
    dActivite:TDateTime;
    sPart0,sPart1,sPart2,sPart3,sAvenant:string;
//    T:TOB;
BEGIN
//T:=nil;
TOBL:=GetTOBLigne(ARow) ;
if TOBL = Nil then exit;
TOBL.PutLigneGrid(GS,ARow,False,False,LesColonnes) ;
for i:=1 to GS.ColCount-1 do FormateZoneSaisie(GS, i,ARow) ;

if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc('SO_AFRECHTIERSPARNOM')=true) then
    begin
    if RemplirTOBTiersP ( TOBTiers, TOBTierss, TOBL.GetValue('ACT_TIERS') ) then
        GS.Cells[SG_Tiers,ARow] := TOBTiers.GetValue('T_LIBELLE')
    else
        GS.Cells[SG_Tiers,ARow] := '';
    end;

if IsValidDate(TOBL.GetValue('ACT_DATEACTIVITE')) then
   begin
   dActivite := TOBL.GetValue('ACT_DATEACTIVITE');
   DecodeDate(dActivite, iAnnee, iMois, iJour);
   GS.Cells[SG_Jour,ARow] := inttostr(iJour);
   end;

if (VartoStr(TOBL.GetValue('ACT_DESCRIPTIF'))<>'') and (VartoStr(TOBL.GetValue('ACT_DESCRIPTIF'))<>#$00) then
   begin
   GS.Cells[SG_Desc,ARow] := '&';
   end;


if (CleAct.TypeSaisie = tsaRess) and not gCtrlF2 and not gCtrlF3 and not gCtrlF4 then
   begin
   if (TOBL.GetValue('ACT_AFFAIRE')<>'') then
      begin
	  {$IFDEF BTP}
      BTPCodeAffaireDecoupe(TOBL.GetValue('ACT_AFFAIRE'),sPart0, sPart1, sPart2, sPart3,sAvenant,Action, false);
      {$ELSE}
      CodeAffaireDecoupe(TOBL.GetValue('ACT_AFFAIRE'),sPart0, sPart1, sPart2, sPart3,sAvenant,Action, false);
      {$ENDIF}
      GS.Cells[SG_Aff0,ARow] := sPart0;
      GS.Cells[SG_Aff1,ARow] := sPart1;
      GS.Cells[SG_Aff2,ARow] := sPart2;
      GS.Cells[SG_Aff3,ARow] := sPart3;
      if (GetParamSoc('SO_AFFGESTIONAVENANT')= true) then
        GS.Cells[SG_Avenant,ARow] := sAvenant;
      end
   else
      begin
      GS.Cells[SG_Aff0,ARow] := '';
      GS.Cells[SG_Aff1,ARow] := '';
      GS.Cells[SG_Aff2,ARow] := '';
      GS.Cells[SG_Aff3,ARow] := '';
      if (GetParamSoc('SO_AFFGESTIONAVENANT')= true) then
        GS.Cells[SG_Avenant,ARow] := '';
      end;
   end;


if (GS.Row = ARow) then
   begin
   GS.Cells[SG_Curseur,ARow] := '4';
   end;

END ;

procedure TFActivite.AfficheNouvellesLignesActivite;
var
i:integer;
sCritere:string;
begin
// Vide la TOB Activité
GS.ClearSelected ;
bSelectAll.Down := false;
TobActivite.ClearDetail;

// Vide le Grid
for i:=1 to GS.RowCount do
   GS.Rows[i].Clear;

// Charge la TOB Activité
if (CleAct.TypeSaisie = tsaRess) then
   sCritere := ACT_Ressource.Text
else
   sCritere := ACT_Affaire.Text;

ChargerTOBDetailActivite(sCritere, ACT_FOLIO.Text);

// Copie dans la TOBold
// PL le 03/05/02 : suppression des lignes modifiées par un delete général
//TOBActiviteOld.ClearDetail;
//TOBActiviteOld.Dupliquer(TOBActivite,TRUE,TRUE,TRUE) ;

// Affiche les lignes de la TOB Activité dans le Grid et rempli les tob nécessaires
AfficheChargeLesTOBLigne ;

//EntreDansLigneCourante(GS.Col,GS.Row);
end;


Function TFActivite.GetChampLigne ( Champ : String ; ARow : integer ) : Variant ;
Var TOBL : TOB ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
Result:=TOBL.GetValue(Champ) ;
END ;

procedure TFActivite.InsertTOBLigne (ARow : integer);
  Var
  TOBL : TOB;
BEGIN
  TOBL := NewTOBLigne (ARow);
  InitialiseLigneAct (TOBL, ARow, CleAct, Action, Dev.Code);
END ;

Function TFActivite.NewTOBLigne (ARow : integer) : TOB;
  Var
  TOBL : TOB;
BEGIN
  TOBL := TOB.Create ('ACTIVITE', TOBActivite, ARow - 1);
  TOBL.AddChampSup('ARTSLIES', false);

  Result := TOBL;
END ;

procedure TFActivite.DepileTOBLignes (ARow, NewRow : integer);
  Var i, Limite, MaxL : integer ;
BEGIN
//
// Suppression des lignes vides de la fin
//
//if NewRow>ARow then Exit ;
  Limite := TOBActivite.Detail.Count + 1;
  MaxL := Max(Limite, ARow);
  
  for i := MaxL downto NewRow + 1 do
    if Not EstRempliGCActivite (GS, i) then
      Limite := i
    else
      Break;

  for i := TOBActivite.Detail.Count - 1 downto Limite - 1 do
    BEGIN
      TOBActivite.Detail[i].Free;
    END ;
end;
(*
procedure TFActivite.VideLignesIntermediaires;
var
  i : integer;
  NbTob : integer;
begin
  //
  // Suppression des lignes vides intermédiaires
  //
  try
    NbTob := TOBActivite.Detail.Count;
    if (NbTob = 0) then exit;
    i := 1;
    while i <= NbTob do
      if Not EstRempliGCActivite (GS, i) then
        begin
          GS.DeleteRow (i) ;
          TOBActivite.Detail[i - 1].Free;
          Dec (NbTob);
        end
      else
        begin
          Inc (i);
        end;

  finally
    if GS.RowCount < NbRowsInit then GS.RowCount := GS.RowCount + NbRowsPlus;
//  NumeroteLignes (GS, TOBActivite);
  end;

END ;
*)

procedure TFActivite.CreerTOBLignes ( ARow : integer );
  Var
  NumL : integer ;
begin

  while TOBActivite.Detail.Count < ARow do
    begin
    NewTOBLigne (0);
    NumL := TOBActivite.Detail.Count;
    InitialiseLigne (NumL);
    TOBActivite.Detail[NumL-1].SetAllModifie (false);
    end;

end;


Procedure TFActivite.AfficheDetailTOBLigne ( ARow : integer ) ;
Var TOBLig,TOBArt : TOB ;
Article:string;
Assistant:string;
BEGIN
Article:='';
ACT_ARTICLEL.Caption := '';
ACT_RESSOURCEL.Caption := '';

if ARow<=TOBActivite.Detail.Count then
   BEGIN
   TOBLig:=TOBActivite.Detail[ARow-1] ;
   Article:=TOBLig.GetValue('ACT_ARTICLE') ;
   Assistant:=TOBLig.GetValue('ACT_RESSOURCE') ;

   // Libelle de l'article
   TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[Article],False) ;
   if (TOBArt=nil) then
      begin
      TOBArt:=CreerTOBArt(TOBArticles);
      TrouverArticleSQL_GI(false, TOBLig.GetValue('ACT_CODEARTICLE'),TOBArt, TOBLig.GetValue('ACT_TYPEARTICLE'));
      end;
   if (TOBArt<>nil) then
      ACT_ARTICLEL.Caption := TOBArt.GetValue('GA_LIBELLE');

   // Libelle de l'assistant
   if RemplirTOBAssistant( TOBAssistant, AFOAssistants, Assistant )<0 then
      ACT_RESSOURCEL.Caption := ''
   else
      ACT_RESSOURCEL.Caption := TOBAssistant.GetValue('ARS_LIBELLE');

   // Libelle du tiers
   if (RemplirTOBTiersP(TOBTiers,TOBTierss,TOBLig.GetValue('ACT_TIERS'))=true) then
      begin
      LibelleTiers.Caption:=TOBTiers.GetValue('T_LIBELLE') ;
      ACT_TIERSL.Caption :=TOBTiers.GetValue('T_LIBELLE') ;
      MOIS_CLOT_CLIENT.Text := TOBTiers.GetValue('T_MOISCLOTURE');
      MOIS_CLOT_CLIENT_.Text := TOBTiers.GetValue('T_MOISCLOTURE');
      end
   else
      begin
      LibelleTiers.Caption:='' ;
      ACT_TIERSL.Caption :='' ;
      MOIS_CLOT_CLIENT.Text := '';
      MOIS_CLOT_CLIENT_.Text := '';
      end;

   if (ACT_AFFAIRE1.visible=false) then
      begin
      ACT_AFFAIRE.Text := TOBLig.GetValue('ACT_AFFAIRE');
      ACT_AFFAIRE0.Text := TOBLig.GetValue('ACT_AFFAIRE0');
      ACT_AFFAIRE1.Text := TOBLig.GetValue('ACT_AFFAIRE1');
      ACT_AFFAIRE2.Text := TOBLig.GetValue('ACT_AFFAIRE2');
      ACT_AFFAIRE3.Text := TOBLig.GetValue('ACT_AFFAIRE3');
      ACT_AVENANT.Text := TOBLig.GetValue('ACT_AVENANT');
      ACT_TIERS.Text := TOBLig.GetValue('ACT_TIERS');
      end;

   // Libelle de l'affaire
   if (AFRemplirTOBAffaire(TOBLig.GetValue('ACT_AFFAIRE'),TobAffaire,TobAffaires)=true) then
      begin
      LibelleAffaire.Caption:=TobAffaire.GetValue('AFF_LIBELLE') ;
      ACT_AFFAIREL.Caption :=TobAffaire.GetValue('AFF_LIBELLE') ;
      TraiteDateFactAff(True); // traitement de la date de création du factaff
      end
   else
      begin
      LibelleAffaire.Caption:='' ;
      ACT_AFFAIREL.Caption :='' ;
      TraiteDateFactAff(False); // traitement de la date de création du factaff
      end;

//   ChargeAffaireActivite( true, false );

   TOBLig.PutEcran(Self, PPied) ;
   PutEcranPPiedAvance(TOBLig);
   END ;

END ;

// Affichage ou pas du traitement de la date de création du factaff
procedure TFActivite.TraiteDateFactAff (Ok : Boolean);
BEGIN
if TobAffaire = Nil then Ok := False else
if ok then if (TobAffaire.GetValue('AFF_GENERAUTO')<>'ACT') then Ok := False;
LibCreationFactAff.Visible := Ok; DATEFACTAFF.Visible := Ok; BDATEFACTAFF.Visible := Ok;
END;
(*
procedure TFActivite.RemplirLesTOBDesLignes;
Var i, Moni : integer ;
   LaListe : TStringList;
begin
try
LaListe := TStringList.Create;
for i := 1 to TOBActivite.Detail.Count do
    begin
      RemplirLesTOBCourantes (i);
      Moni := ProchainIndiceAffaires (TOBActivite.Detail[i-1].GetValue('ACT_AFFAIRE'), LaListe);
    end;
finally
LaListe.free;
end;
end;
*)

procedure TFActivite.RemplirLesTOBCourantes (ARow : integer);
  Var
  TOBL : TOB;
begin
  TOBL := TOBActivite.Detail[ARow - 1];

  RemplirTOBAssistant (TOBAssistant, AFOAssistants, TOBL.GetValue('ACT_RESSOURCE'));
  RemplirTOBTiersP (TOBTiers, TOBTierss, TOBL.GetValue('ACT_TIERS'));
  AFRemplirTOBAffaire (TOBL.GetValue('ACT_AFFAIRE'), TobAffaire, TobAffaires);
end;

procedure TFActivite.RemplirLesTOBDesLignes;
  Var
  i : integer;
begin
  for i := 1 to TOBActivite.Detail.Count do
      begin
        RemplirLesTOBCourantes (i);
      end;
end;


procedure TFActivite.AfficheChargeLesTOBLigne;
Var i : integer ;
    TOBL,TOBArt : TOB ;
    RefUnique : String ;
BEGIN
GS.RowCount:=TOBActivite.Detail.Count+NbRowsPlus ;

for i:=0 to TOBActivite.Detail.Count-1 do
    BEGIN
    TOBL:=TOBActivite.Detail[i] ;

    // Récupération article
    RefUnique:=TOBL.GetValue('ACT_ARTICLE') ;
    if RefUnique<>'' then
       BEGIN
       // Lecture ou création TOB article
       TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefUnique],False) ;
       if TOBArt=Nil then
            BEGIN
            TOBArt:=CreerTOBArt(TOBArticles) ;
            TOBArt.SelectDB('"'+RefUnique+'"',Nil) ;
            END ;
       END ;

    // Récupération Ressource
    AFOAssistants.AddRessource(TOBL.GetValue('ACT_RESSOURCE'));

    // Affichage
    AfficheLaTOBLigne(i+1) ;
    END ;
END ;


{==============================================================================================}
{=============================== Manipulation des TOB GLOBALES ================================}
{==============================================================================================}
// Chargement de la TOB des lignes activités en cours
procedure TFActivite.ChargerTOBDetailActivite(AcsCritere:string; AcsFolio:string);
var QQ : TQuery;
    dDateDebut:TDateTime;
    dDateFin:TDateTime;
    YY,iAnneeFin:word;
    iMoisFin:word;
    sReq, sTri, sTypeArt, sWhere : string;
begin
sWhere:=''; QQ:=nil;
If TobActivite=Nil Then exit;

Try
if (VH_GC.AFTYPESAISIEACT = 'MOI') or (VH_GC.AFTYPESAISIEACT = 'FOL') then
   // Sélection sur le mois
   begin
   dDateDebut := EncodeDate(CleAct.Annee,CleAct.Mois,1);
   iMoisFin := CleAct.Mois;
   iAnneeFin := CleAct.Annee;
   if iMoisFin = 12 then
      begin
      iAnneeFin := iAnneeFin + 1;
      iMoisFin := 1;
      end
   else
      iMoisFin := iMoisFin + 1;
   dDateFin := EncodeDate(iAnneeFin,iMoisFin,1);
   end
else
   // Sélection sur la semaine
   begin

{$IFDEF AGLINF545}
   dDateDebut := PremierJourSemaineTempo(CleAct.Semaine, CleAct.Annee);
{$ELSE}
// PL le 27/05/02 : Nouvelles fonctions AGL NumSemaine et PremierJourSemaine AGL550
    YY:=CleAct.Annee;
    if (CleAct.Mois=12) and (CleAct.Semaine=1) then YY := YY + 1;
    if (CleAct.Mois=1) and ((CleAct.Semaine=52) or (CleAct.Semaine=53)) then YY := YY - 1;
    dDateDebut := PremierJourSemaine(CleAct.Semaine,  YY);
///////////////////////// AGL550
{$ENDIF}

   dDateFin := dDateDebut + GrCalendrier.NbJourSemaine;
   end;

// SELECT * indispensable pour l'instant car on fait un SetAllModifie en validation avant le insertorupdate
// si on ne faisait pas de select *, les champs non sélectionnés seraient écrasés
sReq := 'SELECT * FROM ACTIVITE ';
if (CleAct.TypeSaisie = tsaRess) then
   begin
   sWhere := 'WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                           +'" AND ACT_RESSOURCE="'+AcsCritere
                           +'" AND ACT_DATEACTIVITE>="'+UsDateTime(dDateDebut)
                           +'" AND ACT_DATEACTIVITE<"'+UsDateTime(dDateFin);
   end
Else
   begin
   sWhere := 'WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                           +'" AND ACT_AFFAIRE="'+AcsCritere
                           +'" AND ACT_DATEACTIVITE>="'+UsDateTime(dDateDebut)
                           +'" AND ACT_DATEACTIVITE<"'+UsDateTime(dDateFin);
   end;

if (VH_GC.AFTYPESAISIEACT = 'FOL') and (AcsFolio <> '') then
   sWhere := sWhere + '" AND ACT_FOLIO="'+ AcsFolio;

// configuration de la selection suivant le type d'acces à la saisie, par temps, frais, fournitures, global
case CleAct.TypeAcces of
   tacTemps :begin
            sTypeArt := 'PRE';
            end;
   tacFrais :begin
            sTypeArt := 'FRA';
            end;
   tacFourn :begin
            sTypeArt := 'MAR';
            end;
   tacGlobal:begin
            sTypeArt := '';
            end;
end;
if (sTypeArt<>'') then
   sWhere := sWhere + '" AND ACT_TYPEARTICLE="' + sTypeArt;

sWhere := sWhere + '"';
sReq := sReq + sWhere;

sTri:='';
if (HValComboBoxTri.Text<>'') then
   begin
   sReq := sReq + ' ORDER BY ';
   if (HValComboBoxTri.Value='JOU') then sTri := 'ACT_DATEACTIVITE' else   // Jour
   if (HValComboBoxTri.Value='CLI') then sTri := 'ACT_TIERS' else          // Client
   if (HValComboBoxTri.Value='MIS') then sTri := 'ACT_AFFAIRE' else        // Mission
   if (HValComboBoxTri.Value='RES') then sTri := 'ACT_RESSOURCE' else      // Ressource
   if (HValComboBoxTri.Value='PRE') then sTri := 'ACT_CODEARTICLE' else    // Prestation
   if (HValComboBoxTri.Value='TYP') then sTri := 'ACT_TYPEARTICLE' else    // type de Prestation
   if (HValComboBoxTri.Value='SAI') then sTri := 'ACT_NUMLIGNE' ;       // ordre de saisie
   sReq := sReq +  sTri ;

   if (sTri <> '') and (OrdreTri <> '') then
      sReq := sReq + ' '+ OrdreTri ;
   end;

  gsReqPourDelete := sWhere;

  QQ := OpenSQL (sReq, true,-1,'',true);

  if Not QQ.EOF then
    begin
      TobActivite.LoadDetailDB ('ACTIVITE', '', '', QQ, False);

      TobActivite.Detail[0].AddChampSupValeur ('ARTSLIES', '', true); // Ce champs boolean se doit d'être initialisé à vide
    end;
  Finally
    Ferme(QQ);
  end;

end;

// Chargement de la TOB du nombre d'heures d'activité par jour
function TFActivite.ChargerTOBNbHeureJour(AcsCritere:string; AcsFolio:string; AcdDateDebut:TDateTime; AcTypeAcc:T_Acc):TOB;
var QQ : TQuery;
    dDateDebut : TDateTime;
    dDateFin : TDateTime;
    sReq, sUnite, sTypeArt: string;
    i : integer;
begin
QQ:=nil;
Result:=TOB.Create('Nb heures par jour',Nil,-1) ;


Try
// somme les quantites sur toutes les cellules affichées du calendrier
dDateDebut := AcdDateDebut;
dDateFin := dDateDebut + (GrCalendrier.NbJourSemaine * GrCalendrier.NbRow);
// configuration de la selection suivant le type d'acces à la saisie, par temps, frais, fournitures, global
case AcTypeAcc of
   tacTemps :begin
            sTypeArt := 'PRE';
            sUnite := VH_GC.AFMESUREACTIVITE;
            end;
   tacFrais :begin
            sTypeArt := 'FRA';
            sUnite := '';
            end;
   tacFourn :begin
            sTypeArt := 'MAR';
            sUnite := '';
            end;
   tacGlobal:begin
            sTypeArt := '';
            sUnite := VH_GC.AFMESUREACTIVITE;
//            sUnite := ''; // PL le 02/07/03 : ne créait pas de TOB pour comptabiliser les frais mensuel en saisie globale
            end;
end;

if (CleAct.TypeSaisie = tsaRess) then
   begin
   // PL le 06/03/02 : INDEX 2
   sReq := 'SELECT ACT_DATEACTIVITE,ACT_UNITE,SUM(ACT_QTE) AS ACT_QTE,ACT_ACTIVITEEFFECT,ACT_ACTIVITEREPRIS FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                           +'" AND ACT_RESSOURCE="'+AcsCritere
                           +'" AND ACT_DATEACTIVITE>="'+UsDateTime(dDateDebut)
                           +'" AND ACT_DATEACTIVITE<"'+UsDateTime(dDateFin) + '"'

   end
Else
   begin
   // PL le 06/03/02 : INDEX 4
   sReq := 'SELECT ACT_DATEACTIVITE,ACT_UNITE,SUM(ACT_QTE) AS ACT_QTE,ACT_ACTIVITEEFFECT,ACT_ACTIVITEREPRIS FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                           +'" AND ACT_AFFAIRE="'+AcsCritere
                           +'" AND ACT_DATEACTIVITE>="'+UsDateTime(dDateDebut)
                           +'" AND ACT_DATEACTIVITE<"'+UsDateTime(dDateFin) + '"'
   end;

if (VH_GC.AFTYPESAISIEACT = 'FOL') and (AcsFolio <> '') then
   sReq := sReq + ' AND ACT_FOLIO="'+ AcsFolio + '"';

if (sTypeArt<>'') then
  begin
    sReq := sReq + ' AND ACT_TYPEARTICLE="' + sTypeArt + '"';
    if (sTypeArt='PRE') then  (* ???? *)
  end;

if (sUnite = VH_GC.AFMESUREACTIVITE) then
   begin
   if (TOBListeConv<>nil) then
   if (TOBListeConv.Detail.Count<>0) then
      begin
      sReq := sReq + ' AND (ACT_UNITE="' + TOBListeConv.Detail[0].GetValue('GME_MESURE') + '"';

      for i:=1 to TOBListeConv.Detail.Count-1 do
         sReq := sReq + ' OR ACT_UNITE="' + TOBListeConv.Detail[i].GetValue('GME_MESURE')+ '"';

      sReq := sReq + ')' ;
      end;
   end;

sReq := sReq + ' GROUP BY ACT_DATEACTIVITE,ACT_UNITE,ACT_ACTIVITEREPRIS,ACT_ACTIVITEEFFECT ORDER BY ACT_DATEACTIVITE' ;

QQ:=OpenSQL(sReq,true,-1,'',true);

if Not QQ.EOF then
    Result.LoadDetailDB('Activites par jour','','',QQ,false);


if (sUnite <> VH_GC.AFMESUREACTIVITE) then exit;

// Rajoute les champs sup pour stocker les lignes Fac et non Fac
if (Result.Detail.count<>0) then
    begin
    Result.Detail[0].AddChampSup('ACT_QTEFACTURE', true);
    Result.Detail[0].AddChampSup('ACT_QTENFACTURE', true);
    end;

// On remet toutes les occurrences de la TOB dans la même unité : en heure unite de conversion par defaut
ConversionTOBActivite( Result, 'ACT_DATEACTIVITE' );

Finally
Ferme(QQ);
end;

end;


procedure TFActivite.CompleteLesFraisTOBNbHeureJour;
  var
  QQ : TQuery;
  TOBDesFrais, TOBFrais, TOBQ : TOB;
  Col, Row, i, iMoisFin, iAnneeFin : integer;
  ReqFrais : string;
  dDateDebut, dDateFin : TDateTime;
begin
  dDateDebut := EncodeDate (CleAct.Annee, CleAct.Mois, 1);
  iMoisFin := CleAct.Mois;
  iAnneeFin := CleAct.Annee;
  if iMoisFin = 12 then
    begin
    iAnneeFin := iAnneeFin + 1;
    iMoisFin := 1;
    end
  else
    iMoisFin := iMoisFin + 1;
  dDateFin := EncodeDate (iAnneeFin, iMoisFin, 1);

  ReqFrais := 'SELECT ACT_DATEACTIVITE FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + CleAct.TypeActivite
                         +'" AND ACT_RESSOURCE="' + CleAct.Ressource
                         +'" AND ACT_DATEACTIVITE>="' + UsDateTime (dDateDebut)
                         +'" AND ACT_DATEACTIVITE<"' + UsDateTime (dDateFin)
                         +'" AND ACT_TYPEARTICLE="FRA"';

  if (VH_GC.AFTYPESAISIEACT = 'FOL') and (CleAct.Folio <> 0) then
    ReqFrais := ReqFrais + ' AND ACT_FOLIO="'+ inttostr(CleAct.Folio) + '"';

  ReqFrais := ReqFrais + ' GROUP BY ACT_DATEACTIVITE ORDER BY ACT_DATEACTIVITE';


  TOBDesFrais := TOB.Create ('Frais par jour', Nil, -1);
  try
  QQ := OpenSQL (ReqFrais, true,-1,'',true);

  if Not QQ.EOF then
    begin
      TOBDesFrais.LoadDetailDB ('Frais par jour', '', '', QQ, false);

      for i := 0 to TOBDesFrais.detail.count - 1 do
        begin
          TOBFrais := TOBDesFrais.detail[i];
          GrCalendrier.CelluleDunJour(TOBFrais.GetValue('ACT_DATEACTIVITE'), Col, Row);
          TOBQ := TOB(GCalendrier.Objects[Col,Row]);
          if (TOBQ = nil) then
            // aucune TOB n'était encore liée
            begin
              TOBQ := TOB.Create ('ACTIVITE', TOBNbHeureJour, -1);
              TOBQ.PutValue('ACT_DATEACTIVITE', TOBFrais.GetValue('ACT_DATEACTIVITE'));
              TOBQ.PutValue('ACT_PERIODE', GetPeriode(TOBFrais.GetValue('ACT_DATEACTIVITE')));
              TOBQ.PutValue('ACT_SEMAINE', NumSemaine(TOBFrais.GetValue('ACT_DATEACTIVITE')));
              TOBQ.PutValue('ACT_QTE', 0);
              TOBQ.AddChampSup('ACT_QTEFACTURE', false);
              TOBQ.AddChampSup('ACT_QTENFACTURE', false);
              TOBQ.PutValue('ACT_QTENFACTURE', 0);
              TOBQ.PutValue('ACT_QTEFACTURE', 0);
              GCalendrier.Objects[Col,Row] := TOBQ;
            end;
        end;
    end;

  finally
    TOBDesFrais.Free;
    Ferme(QQ);
  end;
end;

// Calcul de la somme des Frais en PV pour la periode courante
function TFActivite.CalculSommeFrais (AcsFolio : string) : double;
var
  QQ : TQuery;
  dDateDebut, dDateFin : TDateTime;
  sCritere, sReq : string;
  FYear, FMois : word;
begin
  QQ := nil;
  Result := 0;

  // Charge la TOB Activité
  if (CleAct.TypeSaisie = tsaRess) then
     sCritere := ACT_Ressource.Text
  else
     sCritere := ACT_Affaire.Text;


  Try
  dDateDebut := EncodeDate (CleAct.Annee, CleAct.Mois, 1);
  if (CleAct.Mois = 12) then
      begin
        FMois := 1;
        FYear := CleAct.Annee + 1;
      end
  else
      begin
        FMois := CleAct.Mois + 1;
        FYear := CleAct.Annee;
      end;

  dDateFin := EncodeDate (FYear, Fmois, 1);


  if (CleAct.TypeSaisie = tsaRess) then
     begin
     sReq := 'SELECT SUM(ACT_TOTVENTE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + CleAct.TypeActivite
                             + '" AND ACT_RESSOURCE="' + sCritere
                             + '" AND ACT_DATEACTIVITE>="' + UsDateTime (dDateDebut)
                             + '" AND ACT_DATEACTIVITE<"' + UsDateTime (dDateFin) + '"'

     end
  Else
     begin
     sReq := 'SELECT SUM(ACT_TOTVENTE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + CleAct.TypeActivite
                             + '" AND ACT_AFFAIRE="' + sCritere
                             + '" AND ACT_DATEACTIVITE>="' + UsDateTime (dDateDebut)
                             + '" AND ACT_DATEACTIVITE<"' + UsDateTime (dDateFin) + '"'
     end;

  if (VH_GC.AFTYPESAISIEACT = 'FOL') and (AcsFolio <> '') then
     sReq := sReq + ' AND ACT_FOLIO="'+ AcsFolio + '"';

  sReq := sReq + ' AND ACT_TYPEARTICLE="FRA"';


  QQ := OpenSQL (sReq, true,-1,'',true);

  if Not QQ.EOF then
    begin
      Result := QQ.Fields[0].Asfloat;
    end;


  Finally
    Ferme (QQ);
  end;

end;


// Conversion d'une TOB et filles dans l'unite courante
function TFActivite.ConversionTobActivite( var TobActiviteSelect:TOB ; AcsChampRupture:string):Double;
var
TOBD, TOBInter:TOB;
i:integer;
dQteCourante, dQuotHeure, dQuotConv, dQteFac, dQteNFac, dQteConv:double;
sUniteCourante:string;
begin
Result := 0;

if (TobActiviteSelect.Detail.Count = 0) then exit;
if (TOBListeConv=nil) or (TOBListeConv.Detail.Count = 0) then exit;

i:=0;
TOBD := TOBListeConv.FindFirst(['GME_MESURE'],[VH_GC.AFMESUREACTIVITE],TRUE) ;
if (TOBD=nil) then exit;
dQuotHeure := TOBD.GetValue('GME_QUOTITE');
if (dQuotHeure=0) then exit;
while (i<TobActiviteSelect.Detail.Count) do
   begin
   sUniteCourante := TobActiviteSelect.Detail[i].GetValue('ACT_UNITE');
   TOBD := TOBListeConv.FindFirst(['GME_MESURE'],[sUniteCourante],TRUE) ;
   if (TOBD <> nil) then
      begin
      dQuotConv := TOBD.GetValue('GME_QUOTITE');
      dQteCourante := 0; dQteNFac:=0; dQteFac := 0;
      if (TobActiviteSelect.Detail[i].GetValue('ACT_ACTIVITEEFFECT')='X') then
            begin
            dQteConv:=(TobActiviteSelect.Detail[i].GetValue('ACT_QTE')*dQuotConv)/dQuotHeure;
            dQteCourante:= dQteConv;
            if TobActiviteSelect.Detail[i].FieldExists('ACT_QTEFACTURE') then
                begin
                if (TobActiviteSelect.Detail[i].GetValue('ACT_ACTIVITEREPRIS')='N') then
                    dQteNFac := dQteNFac + dQteConv;
                if (TobActiviteSelect.Detail[i].GetValue('ACT_ACTIVITEREPRIS')<>'N') then
                    dQteFac := dQteFac + dQteConv;
                end;
            end;
// PL le 15/01/02 pour V575
      while ((i+1) < TobActiviteSelect.Detail.Count) and
      (TobActiviteSelect.Detail[i+1].GetValue(AcsChampRupture)=TobActiviteSelect.Detail[i].GetValue(AcsChampRupture) and
      (TobActiviteSelect.Detail[i+1].GetValue('ACT_ACTIVITEEFFECT')='X'))
// Fin PL le 15/01/02 pour V575
      do
         begin
         TOBInter := TobActiviteSelect.Detail[i];
         TOBD := TOBListeConv.FindFirst(['GME_MESURE'],[TobActiviteSelect.Detail[i+1].GetValue('ACT_UNITE')],TRUE) ;
         if (TOBD <> nil) then
            begin
            dQuotConv := TOBD.GetValue('GME_QUOTITE');
            dQteConv := (TobActiviteSelect.Detail[i+1].GetValue('ACT_QTE')*dQuotConv)/dQuotHeure;
            dQteCourante := dQteCourante + dQteConv;
            if TobActiviteSelect.Detail[i+1].FieldExists('ACT_QTEFACTURE') then
                begin
                if (TobActiviteSelect.Detail[i+1].GetValue('ACT_ACTIVITEREPRIS')='N') then
                    dQteNFac := dQteNFac + dQteConv;
                if (TobActiviteSelect.Detail[i+1].GetValue('ACT_ACTIVITEREPRIS')<>'N') then
                    dQteFac := dQteFac + dQteConv;
                end;
            end;
//         TobActiviteSelect.Detail.Delete(i);
         TOBInter.Free;
         end;

      Result := Result + dQteCourante;
      TobActiviteSelect.Detail[i].PutValue('ACT_QTE',dQteCourante);
      TobActiviteSelect.Detail[i].PutValue('ACT_UNITE',VH_GC.AFMESUREACTIVITE);
      if TobActiviteSelect.Detail[i].FieldExists('ACT_QTEFACTURE') then
            begin
            TobActiviteSelect.Detail[i].PutValue('ACT_QTEFACTURE',dQteFac);
            TobActiviteSelect.Detail[i].PutValue('ACT_QTENFACTURE',dQteNFac);
            end;
      end;
   Inc(i);
   end;
end;

// Chargement de la TOB des sommes des heures d'activité par semaine
function TFActivite.ChargerTOBSommeSemaine(var TOBHeuresParJour:TOB ):TOB;
var
i:integer;
OldNumSemaine:integer;
dSommeSemaine:double;
TOBT:TOB;
begin
Result := nil;
if (TOBHeuresParJour.Detail.Count=0) then exit;

Result:=TOB.Create('Somme par semaine',Nil,-1) ;

i:=0;
while (i<TOBHeuresParJour.Detail.Count) do
   begin
   OldNumSemaine := NumSemaine(TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_DATEACTIVITE'));
   dSommeSemaine:=0;

   while (i<TOBHeuresParJour.Detail.Count)
         and (NumSemaine(TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_DATEACTIVITE'))=OldNumSemaine) do
      begin
      dSommeSemaine := dSommeSemaine + TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_QTE');
      Inc(i);
      end;

   TOBT:= TOB.Create('Somme une semaine',Result,-1) ;
   TOBT.LoadFromSt('SEMAINE;SOMME;',';',inttostr(OldNumSemaine)+':'+ floattostr(dSommeSemaine),':') ;
   end;

end;

// Calcul de la somme des heures d'activité pour le mois courant
function TFActivite.CalculSommeMois (TOBHeuresParJour : TOB; var AvdNbHeureFac, AvdNbHeureNFac : double ):double;
  Var
  YY, MM, DD : Word;
  i : integer;
  ddate : TDateTime;
BEGIN
  Result := 0;
  AvdNbHeureFac := 0;
  AvdNbHeureNFac := 0;
  if (TOBHeuresParJour.Detail.Count = 0) then exit;

  try
  i := 0;
  while (i < TOBHeuresParJour.Detail.Count) do
     begin
     ddate := TOB (TOBHeuresParJour.Detail[i]).GetValue ('ACT_DATEACTIVITE');
     DecodeDate (ddate, YY, MM, DD);

     if (MM = CleAct.Mois) and (TOB (TOBHeuresParJour.Detail[i]).GetValue ('ACT_ACTIVITEEFFECT') <> '-') then
          begin
          Result := Result + TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_QTE');

          if (TOB(TOBHeuresParJour.Detail[i]).FieldExists('ACT_QTEFACTURE')) then
              AvdNbHeureFac := AvdNbHeureFac + TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_QTEFACTURE');

          if (TOB(TOBHeuresParJour.Detail[i]).FieldExists('ACT_QTENFACTURE')) then
              AvdNbHeureNFac := AvdNbHeureNFac + TOB(TOBHeuresParJour.Detail[i]).GetValue('ACT_QTENFACTURE');

          end;
     Inc(i);
     end;

  finally
    gbModifActRep := false; // PL le 16/06/03 : on initialise le test de verif de modif du F/NF
  end;
end;


// fonction de remise à jour des heures saisies par jour dans le calendrier,
// dans la TOB associée à la cellule du calendrier
// au jour le jour (saisie sur le grid d'une ligne d'activité => modification en direct dans le calendrier)
procedure TFActivite.MajTOBHeureCalendrier(AciLigneTraitee:integer);
var
TOBT, TOBLigneEncours, TOBSommeEncours:TOB;
TOBEntete:TOB;
TOBU, TOBQuantite:TOB;
dDateSaisie:TDateTime;
Col,Row:integer;
dQteOld, dQteLig :double;
bPremierCas:boolean;
OldSyn,OldEnabled :boolean;
iSemaine:integer;
begin
OldSyn := GCalendrier.SynEnabled;
OldEnabled := GCalendrier.Enabled;
GCalendrier.CacheEdit ; GCalendrier.SynEnabled:=False ; GCalendrier.Enabled:=false;

TOBT := GetTOBLigne(AciLigneTraitee);
if (TOBT=nil) then exit;
try
TOBEntete := TOB.Create('Entete Ligne activite',Nil,-1) ;
TOBLigneEncours := TOB.Create('ACTIVITE',TOBEntete,-1) ;
TOBLigneEncours.Dupliquer(TOBT,TRUE,TRUE,FALSE) ;
bPremierCas := false;

// Si la ligne etait vide, il faut verifier si une TOB etait deja liée a la date saisie
// Si non, on cree la TOB, on l'initialise, et on la lie au calendrier
// Si oui, on modifie la TOB deja liee au calendrier
if (TOBLigneActivite=nil) then
   bPremierCas := true
else
   if (TOBLigneActivite.GetValue('ACT_QTE')=0) or Not IsValidDate(TOBLigneActivite.GetValue('ACT_DATEACTIVITE')) then
      bPremierCas := true;

TOBU := TOBListeConv.FindFirst(['GME_MESURE'],[TOBLigneActivite.GetValue('ACT_UNITE')],TRUE) ;

if (bPremierCas=true) then
   begin
   if (TOBLigneEncours.GetValue('ACT_QTE')=0) or Not IsValidDate(TOBLigneEncours.GetValue('ACT_DATEACTIVITE')) then exit;

   dDateSaisie := TOBLigneEncours.GetValue('ACT_DATEACTIVITE');

   dQteLig := ConversionTobActivite( TOBEntete, 'ACT_DATEACTIVITE' );
   GrCalendrier.CelluleDunJour(dDateSaisie, Col, Row);
   iSemaine:=NumSemaine(dDateSaisie);
   TOBSommeEncours:=TOBSommeSemaine.FindFirst(['SEMAINE'],[iSemaine],TRUE) ;

   TOBQuantite := TOB(GCalendrier.Objects[Col,Row]);
   if (TOBQuantite = nil) then
      // aucune TOB n'était encore liée
      begin
      TOBQuantite := TOB.Create('ACTIVITE',TOBNbHeureJour,-1) ;
      TOBQuantite.PutValue('ACT_DATEACTIVITE', dDateSaisie);
      TOBQuantite.PutValue('ACT_PERIODE', GetPeriode(dDateSaisie));
      TOBQuantite.PutValue('ACT_SEMAINE', NumSemaine(dDateSaisie));
      TOBQuantite.PutValue('ACT_QTE', dQteLig);
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      TOBQuantite.AddChampSup('ACT_QTEFACTURE', false);
      TOBQuantite.AddChampSup('ACT_QTENFACTURE', false);
      TOBQuantite.PutValue('ACT_QTENFACTURE', 0);
      TOBQuantite.PutValue('ACT_QTEFACTURE', 0);
      if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
        TOBQuantite.PutValue('ACT_QTENFACTURE', dQteLig)
      else
        TOBQuantite.PutValue('ACT_QTEFACTURE', dQteLig);

      GCalendrier.Objects[Col,Row]:= TOBQuantite;
      if (TOBSommeEncours=nil) then
         begin
         TOBSommeEncours:= TOB.Create('Somme une semaine',TOBSommeSemaine,-1) ;
         TOBSommeEncours.LoadFromSt('SEMAINE;SOMME;',';',inttostr(iSemaine)+':0',':') ;
         end;

      TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);
      end
   else
      // une TOB était deja liee, on la modifie
      begin
      dQteOld := TOBQuantite.GetValue('ACT_QTE');
      TOBQuantite.PutValue('ACT_QTE', dQteLig + dQteOld );
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      if TOBQuantite.FieldExists('ACT_QTEFACTURE') then
        begin
        if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          TOBQuantite.PutValue('ACT_QTENFACTURE', TOBQuantite.GetValue('ACT_QTENFACTURE') + dQteLig)
        else
          TOBQuantite.PutValue('ACT_QTEFACTURE', TOBQuantite.GetValue('ACT_QTEFACTURE') + dQteLig);
        end;
      if (TOBSommeEncours<>nil) then
         TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);
      end;

    // On met à jour le montant de frais si c'est ce qui a été saisi
   if (TOBLigneEncours.GetValue('ACT_TYPEARTICLE') = 'FRA') then
     gdSommeFrais := gdSommeFrais + TOBLigneEncours.GetValue('ACT_TOTVENTE');

   end
else
// on arrive ici avec au moins une TOB d'entree sur la ligne correcte (date OK, Qte<>0) sinon cas precedent
if (Not IsValidDate(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'))) or
   (TOBLigneActivite.GetValue('ACT_DATEACTIVITE')<>TOBLigneEncours.GetValue('ACT_DATEACTIVITE')) then
   begin
   // Si la date a ete modifiee, il faut soustraire l'ancienne quantite à l'ancienne date, et ajouter la nouvelle quantite
   // à la nouvelle date

   // Ancienne date
   iSemaine:=NumSemaine(TOBLigneActivite.GetValue('ACT_DATEACTIVITE'));
   TOBSommeEncours:=TOBSommeSemaine.FindFirst(['SEMAINE'],[iSemaine],TRUE) ;
   GrCalendrier.CelluleDunJour(TOBLigneActivite.GetValue('ACT_DATEACTIVITE'), Col, Row);
   TOBQuantite := TOB(GCalendrier.Objects[Col,Row]);
   // On met à jour le montant de frais si c'est ce qui a été saisi
   if (TOBLigneActivite.GetValue('ACT_TYPEARTICLE') = 'FRA') then
     gdSommeFrais := gdSommeFrais - TOBLigneActivite.GetValue('ACT_TOTVENTE');
   if ((TOBQuantite <> nil) and (TOBU<>nil)) then
      begin
      dQteLig := TOBLigneActivite.GetValue('ACT_QTE');
      dQteOld := TOBQuantite.GetValue('ACT_QTE');
      TOBQuantite.PutValue ('ACT_QTE', dQteOld - dQteLig);
      if TOBQuantite.FieldExists('ACT_QTEFACTURE') then
        begin
        if (TOBLigneActivite.GetValue ('ACT_ACTIVITEREPRIS') = 'N') then
          TOBQuantite.PutValue ('ACT_QTENFACTURE', TOBQuantite.GetValue('ACT_QTENFACTURE')- dQteLig)
        else
          TOBQuantite.PutValue ('ACT_QTEFACTURE', TOBQuantite.GetValue('ACT_QTEFACTURE') - dQteLig);
        end;
      if (TOBSommeEncours=nil) then
         begin
         TOBSommeEncours:= TOB.Create('Somme une semaine',TOBSommeSemaine,-1) ;
         TOBSommeEncours.LoadFromSt('SEMAINE;SOMME;',';',inttostr(iSemaine)+':0',':') ;
         end;

      TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') - dQteLig);
      end;

   // Nouvelle date
   if (Not IsValidDate(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'))) then exit;
   dQteLig := ConversionTobActivite( TOBEntete, 'ACT_DATEACTIVITE' );
   iSemaine:=NumSemaine(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'));
   TOBSommeEncours:=TOBSommeSemaine.FindFirst(['SEMAINE'],[iSemaine],TRUE) ;
   GrCalendrier.CelluleDunJour(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'), Col, Row);
   TOBQuantite := TOB(GCalendrier.Objects[Col,Row]);
   if (TOBQuantite = nil) then
      // aucune TOB n'était encore liée
      begin
      TOBQuantite := TOB.Create('ACTIVITE',TOBNbHeureJour,-1) ;
      TOBQuantite.PutValue('ACT_DATEACTIVITE', TOBLigneEncours.GetValue('ACT_DATEACTIVITE'));
      TOBQuantite.PutValue('ACT_PERIODE', GetPeriode(TOBLigneEncours.GetValue('ACT_DATEACTIVITE')));
      TOBQuantite.PutValue('ACT_SEMAINE', NumSemaine(TOBLigneEncours.GetValue('ACT_DATEACTIVITE')));
      TOBQuantite.PutValue('ACT_QTE', dQteLig);
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      TOBQuantite.AddChampSup('ACT_QTEFACTURE', false);
      TOBQuantite.AddChampSup('ACT_QTENFACTURE', false);
      TOBQuantite.PutValue('ACT_QTENFACTURE', 0);
      TOBQuantite.PutValue('ACT_QTEFACTURE', 0);
      if TOBQuantite.FieldExists('ACT_QTEFACTURE') then
        begin
        if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          TOBQuantite.PutValue('ACT_QTENFACTURE', dQteLig)
        else
          TOBQuantite.PutValue('ACT_QTEFACTURE', dQteLig);
        end;
      GCalendrier.Objects[Col,Row]:= TOBQuantite;
      if (TOBSommeEncours=nil) then
         begin
         TOBSommeEncours:= TOB.Create('Somme une semaine',TOBSommeSemaine,-1) ;
         TOBSommeEncours.LoadFromSt('SEMAINE;SOMME;',';',inttostr(iSemaine)+':0',':') ;
         end;

      TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);
      end
   else
      // une TOB était deja liee, on la modifie
      begin
      dQteOld := TOBQuantite.GetValue('ACT_QTE');
      TOBQuantite.PutValue('ACT_QTE', dQteLig + dQteOld );
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      if TOBQuantite.FieldExists('ACT_QTEFACTURE') then
        begin
        if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          TOBQuantite.PutValue('ACT_QTENFACTURE', TOBQuantite.GetValue('ACT_QTENFACTURE') + dQteLig)
        else
          TOBQuantite.PutValue('ACT_QTEFACTURE', TOBQuantite.GetValue('ACT_QTEFACTURE') + dQteLig);
        end;
      if (TOBSommeEncours<>nil) then
         TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);
      end;
   // On met à jour le montant de frais si c'est ce qui a été saisi
   if (TOBLigneEncours.GetValue('ACT_TYPEARTICLE') = 'FRA') then
      gdSommeFrais := gdSommeFrais + TOBLigneEncours.GetValue('ACT_TOTVENTE');
   end
else
// on arrive ici avec une TOB d'entree sur la ligne correcte (date OK, Qte<>0),
// une date de TOB de saisie correcte et non modifiée par rapport à la date de la TOB d'entree
   begin
   // Si la date n'a pas changée, il faut soustraire l'ancienne quantite à la nouvelle et ajouter cette difference
   //  à la date traitee
   dQteLig := ConversionTobActivite( TOBEntete, 'ACT_DATEACTIVITE' );
   dQteOld := TOBLigneActivite.GetValue('ACT_QTE');

   iSemaine:=NumSemaine(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'));
   TOBSommeEncours:=TOBSommeSemaine.FindFirst(['SEMAINE'],[iSemaine],TRUE) ;
   GrCalendrier.CelluleDunJour(TOBLigneEncours.GetValue('ACT_DATEACTIVITE'), Col, Row);
   TOBQuantite := TOB(GCalendrier.Objects[Col,Row]);
   if (TOBQuantite = nil) then
      // aucune TOB n'était encore liée
      begin
      TOBQuantite := TOB.Create('ACTIVITE',TOBNbHeureJour,-1) ;
      TOBQuantite.PutValue('ACT_DATEACTIVITE', TOBLigneEncours.GetValue('ACT_DATEACTIVITE'));
      TOBQuantite.PutValue('ACT_PERIODE', GetPeriode(TOBLigneEncours.GetValue('ACT_DATEACTIVITE')));
      TOBQuantite.PutValue('ACT_SEMAINE', NumSemaine(TOBLigneEncours.GetValue('ACT_DATEACTIVITE')));
      TOBQuantite.PutValue('ACT_QTE', dQteLig);
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      TOBQuantite.AddChampSup('ACT_QTEFACTURE', false);
      TOBQuantite.AddChampSup('ACT_QTENFACTURE', false);
      TOBQuantite.PutValue('ACT_QTENFACTURE', 0);
      TOBQuantite.PutValue('ACT_QTEFACTURE', 0);
      if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
        TOBQuantite.PutValue('ACT_QTENFACTURE', dQteLig)
      else
        TOBQuantite.PutValue('ACT_QTEFACTURE', dQteLig);
      GCalendrier.Objects[Col,Row]:= TOBQuantite;
      if (TOBSommeEncours=nil) then
         begin
         TOBSommeEncours:= TOB.Create('Somme une semaine',TOBSommeSemaine,-1) ;
         TOBSommeEncours.LoadFromSt('SEMAINE;SOMME;',';',inttostr(iSemaine)+':0',':') ;
         end;

      TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);
      // On met à jour le montant de frais si c'est ce qui a été saisi
      if (TOBLigneEncours.GetValue('ACT_TYPEARTICLE') = 'FRA') then
        gdSommeFrais := gdSommeFrais + TOBLigneEncours.GetValue('ACT_TOTVENTE');
      end
   else
      // une TOB était deja liee, on la modifie
      begin
      // Si l'unite de l'ancienne ligne n'etait pas bonne, on n'avait jamais comptabilise la quantite old
      // la quantite à ajouter est donc la quantite indiquee new
      if (TOBU<>nil) then
         // sinon, la quantite à ajouter est la difference entre la nouvelle et l'ancienne quantite
         dQteLig := dQteLig - dQteOld;

      dQteOld := TOBQuantite.GetValue('ACT_QTE');
      TOBQuantite.PutValue('ACT_QTE', dQteLig + dQteOld );
      TOBQuantite.PutValue('ACT_UNITE', TOBLigneEncours.GetValue('ACT_UNITE'));
      TOBQuantite.PutValue('ACT_ACTIVITEREPRIS', TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS'));
      if TOBQuantite.FieldExists('ACT_QTEFACTURE') then
        begin
        if (TOBLigneEncours.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          TOBQuantite.PutValue('ACT_QTENFACTURE', TOBQuantite.GetValue('ACT_QTENFACTURE') + dQteLig)
        else
          TOBQuantite.PutValue('ACT_QTEFACTURE', TOBQuantite.GetValue('ACT_QTEFACTURE') + dQteLig);
        end;
      if (TOBSommeEncours<>nil) then
         TOBSommeEncours.PutValue('SOMME', TOBSommeEncours.GetValue('SOMME') + dQteLig);


      // On met à jour le montant de frais si c'est ce qui a été saisi
      if (TOBLigneEncours.GetValue('ACT_TYPEARTICLE') = 'FRA') then
        begin
        if (TOBLigneActivite.GetValue('ACT_TYPEARTICLE') = 'FRA') then
          gdSommeFrais := gdSommeFrais + TOBLigneEncours.GetValue('ACT_TOTVENTE') - TOBLigneActivite.GetValue('ACT_TOTVENTE')
        else
          gdSommeFrais := gdSommeFrais + TOBLigneEncours.GetValue('ACT_TOTVENTE');
        end;
      end;

   end;

TOBLigneActivite.Dupliquer(TOBLigneEncours,TRUE,TRUE,TRUE) ;
ConversionTobActivite( TOBEnteteLigneActivite, 'ACT_DATEACTIVITE' );

GCalendrier.Repaint;
finally
TOBEntete.Free;
GCalendrier.MontreEdit ; GCalendrier.SynEnabled:=OldSyn ; GCalendrier.Enabled:=OldEnabled;
end;

// Calcul des heures par mois si on affiche le nombre d'heure dans le calendrier
//if (CleAct.TypeAcces = tacTemps) or (CleAct.TypeAcces = tacGlobal) then
//  gdNbHeureMois := CalculHeureMois(gdNbHeureFac, gdNbHeureNFac);
end;

// fonction de calcul des heures saisies par mois
(*
function TFActivite.CalculHeureMois(var AvdQteFac, AvdQteNFac:double):double;
var
iCol, iRow : integer;
TobActCourante:TOB;
Annee, Mois, Jour : word;
begin
Result:=0; AvdQteFac:=0; AvdQteNFac:=0;

for iCol:=0 to GCalendrier.ColCount-1 do
    for iRow:=0 to GCalendrier.RowCount-1 do
        begin
        TobActCourante:=TOB(GCalendrier.Objects[iCol,iRow]);
        if (TobActCourante<>nil) then
            if (TobActCourante.NomTable='ACTIVITE') or (TobActCourante.NomTable='Activites par jour') then
                begin
                DecodeDate(TobActCourante.GetValue('ACT_DATEACTIVITE'), Annee, Mois, Jour);
                if (Mois=CleAct.Mois) then
                    begin
                    Result := Result + TobActCourante.GetValue('ACT_QTE');
//                    if (TobActCourante.GetValue('ACT_ACTIVITEREPRIS')='N') then
                    AvdQteNFac := AvdQteNFac + TobActCourante.GetValue('ACT_QTENFACTURE');
//                    if (TobActCourante.GetValue('ACT_ACTIVITEREPRIS')<>'N') then
                    AvdQteFac := AvdQteFac + TobActCourante.GetValue('ACT_QTEFACTURE');
                    end;
                end;
        end;

end;
*)
{===========================================================================================}
{=============================== Gestion de l'Assistant/Ressource ==========================}
{===========================================================================================}
procedure TFActivite.ACT_RESSOURCEElipsisClick(Sender: TObject);
var
sAssistant:string;
begin
sAssistant := ACT_RESSOURCE.Text;
// Identifie l'assistant saisi
IdentifieAssistant(ACT_RESSOURCE) ;
if (sAssistant=ACT_RESSOURCE.Text) then exit;

// Gere la validation des lignes d'activité saisies
//if (GereValider<>oeOk) then Abort;
  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

RafraichirLaSelection;
end;

procedure TFActivite.ACT_RESSOURCEDblClick(Sender: TObject);
var
sAssistant:string;
begin
sAssistant := ACT_RESSOURCE.Text;
// Identifie l'assistant saisi
IdentifieAssistant(ACT_RESSOURCE) ;
if (sAssistant=ACT_RESSOURCE.Text) then exit;

// Gere la validation des lignes d'activité saisies
//if (GereValider<>oeOk) then Abort;

  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

RafraichirLaSelection;
end;

Procedure TFActivite.GereAssistantEnabled;
var
  TOBL : TOB;
  Enabled : boolean;
BEGIN
  Enabled := false;
  if (CleAct.TypeSaisie = tsaRess) then
    begin
      Enabled := ((TOBAssistant.GetValue('ARS_RESSOURCE') <> '') and (ACT_RESSOURCE.Text = TOBAssistant.GetValue('ARS_RESSOURCE')));
      if (ctxlanceur in V_PGI.PGIContexte) and (CleAct.Ressource <> '') then
        Enabled := false;
    end
  else
    begin
      TOBL := GetTOBLigne (GS.Row);
      if (TOBL <> nil) then
        Enabled := ((TOBAssistant.GetValue('ARS_RESSOURCE') <> '') and (TOBL.GetValue('ACT_RESSOURCE') = TOBAssistant.GetValue('ARS_RESSOURCE')));
    end;

  BZoomAssistant.Enabled := Enabled;
  BZoomCalendrierAssistant.Enabled := Enabled and (TOBAssistant.GetValue('ARS_STANDCALEN') <> '');
  if (b35heures.visible = true) then
    b35heures.Enabled := BZoomCalendrierAssistant.Enabled;
END ;

Function TFActivite.IdentifieAssistant (Control : TControl) : boolean;
BEGIN
  Result := GetAssistantRecherche (Control, HTitres.Mess[11], false);
  if Result then
    ChargeAssistant (Control);
  GereAssistantEnabled;
END ;

function TFActivite.ChargeAssistant (Control : TControl) : boolean;
var
  CodeRessource : string;
  iIndiceRess : integer;
BEGIN
  Result := true;
  if (ACT_RESSOURCE.visible = true) then
    // Gestion de la selection par la ressource
    begin
      iIndiceRess := RemplirTOBAssistant(TOBAssistant, AFOAssistants, ACT_RESSOURCE.Text );

      if iIndiceRess<0 then
        begin
          ACT_RESSOURCE.Text:=TOBAssistant.GetValue('ARS_RESSOURCE');
          Result:=false;
        end
      else
        TAFO_Ressource(AFOAssistants.GetRessource(iIndiceRess)).ChargeTempsAnnuel35H;

      CodeRessource:= ACT_RESSOURCE.Text;
      if (TOBAssistant.GetValue('ARS_FERME')='X') then
        begin
          SetConsultation(True, tbaAssistant);
          ACT_RESSOURCE.Font.Color := clRed;
        end
      else
        begin
          SetConsultation(False, tbaAssistant);
          ACT_RESSOURCE.Font.Color := clBlack;
        end;
    end
  else
    // Selection de la ressource sur les lignes d'activite
    begin
      if (Control is THGrid) then
        begin
          if (RemplirTOBAssistant (TOBAssistant, AFOAssistants, THGrid(Control).Cells[SG_Ressource, THGrid(Control).Row]) < 0) then
            THGrid(Control).Cells[SG_Ressource, THGrid(Control).Row] := TOBAssistant.GetValue('ARS_RESSOURCE');
          CodeRessource := THGrid(Control).Cells[SG_Ressource, THGrid(Control).Row];
          ACT_RESSOURCE_.Text := CodeRessource;
        end
      else
        begin
          if RemplirTOBAssistant (TOBAssistant, AFOAssistants, ACT_RESSOURCE_.Text)<0 then
            begin
              ACT_RESSOURCE_.Text := TOBAssistant.GetValue('ARS_RESSOURCE');
              Result := false;
            end;
          CodeRessource := ACT_RESSOURCE_.Text;
        end;
      ACT_RESSOURCEL.caption := TOBAssistant.GetValue('ARS_LIBELLE') ;
    end;

  LibelleAssistant.Caption := TOBAssistant.GetValue('ARS_LIBELLE') ;
  CleAct.TypeRess := TOBAssistant.GetValue('ARS_TYPERESSOURCE');
  if (Action = taCreat) then IncidenceAssistant ;

  CleAct.Ressource := CodeRessource;
END ;

Procedure TFActivite.IncidenceAssistant ;
BEGIN
if Action<>taCreat then Exit ;

GereAssistantEnabled;

TOBActivite.GetEcran(self, PEntete) ;
TOBActivite.GetEcran(self, PEntete2) ;
END ;

Function TFActivite.IdentifierAssistant ( Control:TControl; Var ACol,ARow : integer ; Var Cancel : boolean ; Click : Boolean ) : boolean ;
Var RefSais, CodeAssistant : String ;
    Index:integer;
BEGIN
if (Control is THGrid) then
   RefSais := THGrid(Control).Cells[SG_Ressource,ARow]
else
   RefSais := THCritMaskEdit(Control).Text;

Result:=False ; CodeAssistant:='' ;
// On essai de l'ajouter dans la TOB
Index := AFOAssistants.AddRessource(RefSais);

if ((Index<>-1) and (Index<>-2)) then
   // Si la ressource existe
   CodeAssistant:=TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_Champs.GetValue('ARS_RESSOURCE')
else
   begin
   if (Control is THGrid) then begin THGrid(Control).Col:=SG_Ressource ; THGrid(Control).Row:=ARow ; end;

   if GereElipsis(Control, SG_Ressource) then
      BEGIN
      // Assistant choisi dans la liste affichée
      if (Control is THGrid) then CodeAssistant:=THGrid(Control).Cells[SG_Ressource,ARow] else CodeAssistant:= THCritMaskEdit(Control).Text;

      // On essai de l'ajouter dans la TOB
      Index := AFOAssistants.AddRessource(CodeAssistant);
      if ((Index<>-1) and (Index<>-2)) then
         CodeAssistant:=TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_Champs.GetValue('ARS_RESSOURCE');

      if (Control is THGrid) then THGrid(Control).Cells[SG_Ressource,ARow]:=CodeAssistant else THCritMaskEdit(Control).Text := CodeAssistant;
      END ;
   end;

if ((Index=-1) or (Index=-2))  then
   BEGIN
   Cancel:=True ;
   if (Control is THGrid) then begin THGrid(Control).Col:=ACol ; THGrid(Control).Row:=ARow ; end;
   Exit ;
   END ;

CodesAssToCodesLigne(Control, TAFO_Ressource(AFOAssistants.GetRessource(Index)), ARow) ;
Result:=True ;
END ;

Procedure TFActivite.CodesAssToCodesLigne ( Control:TControl; AFOAss : TAFO_Ressource ; ARow : integer ) ;
Var TOBL : TOB ;
    RefSais : String ;
BEGIN

TOBL:=GetTOBLigne(ARow) ;
if (Control is THGrid) then
   RefSais:=THGrid(Control).Cells[SG_Article,ARow]
else
   RefSais:=THCritMaskEdit(Control).Text;

if (AFOAss=nil) then
   begin
   TOBL.PutValue('ACT_RESSOURCE','') ;
   TOBL.PutValue('ACT_TYPERESSOURCE','') ;
   end
else
   begin
   TOBL.PutValue('ACT_RESSOURCE', AFOAss.tob_Champs.GetValue('ARS_RESSOURCE')) ;
   TOBL.PutValue('ACT_TYPERESSOURCE', AFOAss.tob_Champs.GetValue('ARS_TYPERESSOURCE')) ;
   end;
END ;

procedure TFActivite.ACT_RESSOURCE_Change(Sender: TObject);
begin
if (ACT_RESSOURCE_.Text='') then ACT_RESSOURCEL.caption:='';
end;

procedure TFActivite.ACT_RESSOURCEExit(Sender: TObject);
begin

if csDestroying in ComponentState then Exit ;
// PL le 28/02/02 pour ne plus que le libellé s'efface quand on resaisit de même code
if (ACT_RESSOURCE.Text=CleAct.Ressource) and (ACT_RESSOURCE.Text<>'') then
    begin
    LibelleAssistant.Caption:=TOBAssistant.GetValue('ARS_LIBELLE') ;
    exit;
    end;
// Fin PL le 28/02/02

GereAssistantEnabled;

if ((CleAct.TypeAcces=tacFourn) and (ACT_RESSOURCE.Text='')) then
   // en saisie de fournitures, la ressource n'est pas obligatoire
   begin
   TOBAssistant.InitValeurs(false);
   end
else
    begin
    if Not ChargeAssistant(ACT_RESSOURCE) then
        begin
        ACT_RESSOURCE.SetFocus;
        Abort;
        end;
//    else
//        TAFO_Ressource(AFOAssistants.GetRessource(iIndiceRess)).ChargeTempsAnnuel35H;

    end;
(*    begin
    iIndiceRess:=RemplirTOBAssistant(TOBAssistant,AFOAssistants, ACT_RESSOURCE.Text);
    if (iIndiceRess<0) or (ACT_RESSOURCE.Text='') then
        begin
        if Not IdentifieAssistant(ACT_RESSOURCE) then
            begin
            ACT_RESSOURCE.SetFocus;
            Abort;
            end;
        end
    else
        TAFO_Ressource(AFOAssistants.GetRessource(iIndiceRess)).ChargeTempsAnnuel35H;

    end;*)

// Gere la validation des lignes d'activité saisies
//if (GereValider<>oeOk) then Abort;
  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

LibelleAssistant.Caption:=TOBAssistant.GetValue('ARS_LIBELLE') ;
if Action=taCreat then IncidenceAssistant ;
CleAct.Ressource := ACT_RESSOURCE.Text;
CleAct.TypeRess:=TOBAssistant.GetValue('ARS_TYPERESSOURCE');// mcd 22/04/02 pas ok si change de type entre 2 assistants
GereAssistantEnabled ;

RafraichirLaSelection;
end;


procedure TFActivite.ACT_RESSOURCEChange(Sender: TObject);
begin
if (ACT_RESSOURCE.visible=true) then
   begin
   gbSelectCellCalend:=false;
   ClipBoard.Clear;
   end;

if (ACT_RESSOURCE.Text='') then LIBELLEASSISTANT.Caption:='';
end;

{===========================================================================================}
{======================== Gestion du Tiers/Affaire (client/Mission) ========================}
{===========================================================================================}
procedure TFActivite.ACT_AFFAIREElipsisClick(Sender: TObject);
begin
// Gere la validation des lignes d'activité saisies
//if (GereValider<>oeOk) then Abort;
  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

RafraichirLaSelection;
end;

procedure TFActivite.ACT_AFFAIREDblClick(Sender: TObject);
begin
  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

RafraichirLaSelection;
end;

Procedure TFActivite.GereAffaireEnabled ;
  var
  TOBL : TOB;
  Enabled : boolean;
  sConf : string;
BEGIN
  Enabled := false;
  if (CleAct.TypeSaisie = tsaClient) then
    begin
      Enabled := ((TOBAffaire.GetValue ('AFF_AFFAIRE') <> '') and (ACT_AFFAIRE.Text = TOBAffaire.GetValue ('AFF_AFFAIRE'))) ;
    end
  else
    begin
      TOBL := GetTOBLigne (GS.Row) ;
      if (TOBL <> nil) then
        Enabled := ((TOBAffaire.GetValue ('AFF_AFFAIRE') <> '') and (TOBL.GetValue ('ACT_AFFAIRE') = TOBAffaire.GetValue ('AFF_AFFAIRE')));
    end;

  // PL le 17/01/02
  // L'accès au menu affaire du bouton zoom n'est possible que si on ne gère pas la confidentialité
  sConf := GetParamSoc ('SO_AFTYPECONF');
  if (sConf <> 'A00') then Enabled := False;
  // L'accès au menu affaire du bouton zoom n'est possible que si le user courant n'est pas un invité
  //Enabled := Not VH_GC.UserInvite and Enabled;

  BZoomAffaire.Enabled := Enabled;
  BZoomTableauBord.Enabled := Enabled;
END ;

function TFActivite.ChargeAffaireActivite(TestCle:Boolean; bAvecMessage:boolean):boolean;
Var NbAff : integer;
    bAffaireVisible, bProposition :boolean;
BEGIN
bAffaireVisible:=ACT_AFFAIRE1.visible;
Result:=false;
bProposition := VH_GC.AFProposAct;
//try
NbAff:=ChargeAffaire(ACT_AFFAIRE,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT,ACT_TIERS,ACT_AFFAIRE0,Testcle,Action,TobAffaire,Tobaffaires,bAvecMessage,False,bProposition) ;
if (NbAff =1) then Result := True;
//except
//end;

If (Result=true) Then
    Begin
    LibelleAffaire.Caption:=TOBAffaire.GetValue('AFF_LIBELLE') ;
    ACT_AFFAIREL.Caption:=TOBAffaire.GetValue('AFF_LIBELLE');
    if Action=taCreat then IncidenceAffaire;
    CleAct.Affaire := ACT_AFFAIRE.Text;
    CleAct.Aff0 := ACT_AFFAIRE0.Text;
    CleAct.Aff1 := ACT_AFFAIRE1.Text;
    CleAct.Aff2 := ACT_AFFAIRE2.Text;
    CleAct.Aff3 := ACT_AFFAIRE3.Text;
    CleAct.Avenant := ACT_AVENANT.Text;
    CleAct.Tiers := ACT_TIERS.Text;
    End
Else
    Begin
    LibelleTiers.Caption:='';
    LibelleAffaire.Caption:='';
    ACT_AFFAIREL.Caption:='';
    CleAct.Affaire := '';
    if Not(VH_GC.AFProposAct) then
      CleAct.Aff0       := 'A'
    else
      CleAct.Aff0       := '' ;
    CleAct.Aff1 := '';
    CleAct.Aff2 := '';
    CleAct.Aff3 := '';
    CleAct.Avenant := '';
    CleAct.Tiers := '';
    End;

// Suite au chargement de l'affaire, les propriétés visibles peuvent être ré-affectées
// on les remet donc comme précédemment
if (bAffaireVisible=false) then
   begin
   ACT_AFFAIRE0.visible:=false;
   ACT_AFFAIRE1.visible:=false;
   ACT_AFFAIRE2.visible:=false;
   ACT_AFFAIRE3.visible:=false;
   ACT_AVENANT.visible:=false;
   end
else
   ACT_AFFAIREL.visible:=false;


END ;

Procedure TFActivite.IncidenceAffaire ;
Var
CodeTiers,LibTiers : String ;
BEGIN

if Action<>taCreat then Exit ;

if ((ACT_TIERS.Text='') Or (ACT_TIERS.Text<>TOBTiers.GetValue('T_TIERS'))) then
    Begin
    CodeTiers:=TOBAffaire.GetValue('AFF_TIERS');
    LibTiers:=RechDom('GCTIERSSAISIE',CodeTiers,False) ;
    if ((LibTiers='') or (LibTiers='Error')) then
        BEGIN
        ACT_AFFAIRE.Text:='' ; LibelleAffaire.Caption:='';
        ChargeAffaireActivite( true, true );
//        ChargeCleAffaire;
//        HActivite.Execute(4,Caption,'');
        END else
        BEGIN
        ACT_TIERS.Text:=CodeTiers ; ChargeTiers(ACT_TIERS, CodeTiers) ;
        END ;
    End;

CleAct.Tiers:=ACT_TIERS.Text;
//TOBActivite.PutValue('ACT_AFFAIRE_',CodeAffaire) ;
TOBActivite.GetEcran(self, PEntete) ;
TOBActivite.GetEcran(self, PEntete2) ;

END ;

procedure TFActivite.ACT_AFFAIRE0DblClick(Sender: TObject);
begin
BRechAffaireClick(Sender);
end;

procedure TFActivite.ACT_AFFAIRE1DblClick(Sender: TObject);
begin
BRechAffaireClick(Sender);
end;

procedure TFActivite.ACT_AFFAIRE2DblClick(Sender: TObject);
begin
BRechAffaireClick(Sender);
end;

procedure TFActivite.ACT_AFFAIRE3DblClick(Sender: TObject);
begin
BRechAffaireClick(Sender);
end;

procedure TFActivite.ACT_AFFAIRE0Change(Sender: TObject);
begin
//GoSuiteCodeAffaire(ACT_AFFAIRE0,ACT_AFFAIRE1,0);
end;

procedure TFActivite.ACT_AFFAIRE1Change(Sender: TObject);
begin
GoSuiteCodeAffaire(ACT_AFFAIRE1,ACT_AFFAIRE2,1);
end;

procedure TFActivite.ACT_AFFAIRE2Change(Sender: TObject);
begin
GoSuiteCodeAffaire(ACT_AFFAIRE2,ACT_AFFAIRE3,2);
end;

procedure TFActivite.OnExitPartieAffaire(Sender: TObject);
var
iPartErreur:integer;
bProposition : Boolean;
begin
if csDestroying in ComponentState then Exit ;
bProposition := VH_GC.AFProposAct;
ACT_AFFAIRE.Text:= DechargeCleAffaire(ACT_AFFAIRE0,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT,ACT_TIERS.Text,Action,False,True,bProposition,iPartErreur);
ChargeAffaireActivite(True, true);
GereAffaireEnabled;
end;

procedure TFActivite.ACT_AFFAIRE2Exit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
OnExitPartieAffaire(Sender);
end;

procedure TFActivite.ACT_AFFAIRE3Exit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
OnExitPartieAffaire(Sender);
end;

procedure TFActivite.ACT_AFFAIRE0Exit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
OnExitPartieAffaire(Sender);
end;

procedure TFActivite.ACT_AFFAIRE1Exit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
OnExitPartieAffaire(Sender);
end;

procedure TFActivite.ACT_AFFAIREChange(Sender: TObject);
begin
  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

if (ACT_AFFAIRE.visible=true) then
   begin
   gbSelectCellCalend:=false;
   ClipBoard.Clear;
   end;

CleAct.Affaire := ACT_AFFAIRE.Text;

RafraichirLaSelection;
end;

procedure TFActivite.ACT_AFFAIRE_Change(Sender: TObject);
begin
if (ACT_AFFAIRE_.Text='') then ACT_AFFAIREL.caption:='';
//ChargeCleAffaire(ACT_AFFAIRE0_,ACT_AFFAIRE1_,ACT_AFFAIRE2_,ACT_AFFAIRE3_,ACT_AVENANT,Action,ACT_AFFAIRE_.Text,False);

end;

procedure TFActivite.FormatePartiesAffaire ;
var
bVisiblePart0,bVisiblePart1,bVisiblePart2,bVisiblePart3,bVisiblePart4:boolean;
eOnChange0,eOnChange1,eOnChange2,eOnChange3:TNotifyEvent;
iWidth, icanvas:integer;
begin
bVisiblePart0:=ACT_AFFAIRE0.visible;
bVisiblePart1:=ACT_AFFAIRE1.visible;
bVisiblePart2:=ACT_AFFAIRE2.visible;
bVisiblePart3:=ACT_AFFAIRE3.visible;
bVisiblePart4:=ACT_AVENANT.visible;
eOnChange0:=ACT_AFFAIRE0.OnChange;
eOnChange1:=ACT_AFFAIRE1.OnChange;
eOnChange2:=ACT_AFFAIRE2.OnChange;
eOnChange3:=ACT_AFFAIRE3.OnChange;
try

ChargeCleAffaire(ACT_AFFAIRE0,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT,nil,Action,ACT_AFFAIRE.Text,False);

iWidth := ACT_AFFAIRE0.Width;
ACT_AFFAIRE0_.Width := iWidth;
GS.ColWidths[SG_Aff0]:=iWidth;
if Not(VH_GC.AFProposAct) then
   begin
   ACT_AFFAIRE0_.Visible := false;
   GS.ColWidths[SG_Aff0]:=0 ;
   GS.ColLengths[SG_Aff0]:=-1;
   end
else
   ACT_AFFAIRE0_.Visible := true;

iWidth := ACT_AFFAIRE1.Width;
ACT_AFFAIRE1_.Width := ACT_AFFAIRE1.Width;
ACT_AFFAIRE1_.Visible := ACT_AFFAIRE1.Visible;
icanvas:=GS.Canvas.TextWidth(VH_GC.CleAffaire.Co1Lib);
if (iWidth < (icanvas+30)) then iWidth := icanvas+30;
GS.ColWidths[SG_Aff1]:=iWidth;
if (ACT_AFFAIRE1.visible=false) then
   begin
   GS.ColWidths[SG_Aff1]:=0 ;
   GS.ColLengths[SG_Aff1]:=-1;
   end;

iWidth := ACT_AFFAIRE2.Width;
ACT_AFFAIRE2_.Width := ACT_AFFAIRE2.Width;
ACT_AFFAIRE2_.Visible := ACT_AFFAIRE2.Visible;
if VH_GC.CleAffaire.Co2Visible then
   begin
   icanvas:=GS.Canvas.TextWidth(VH_GC.CleAffaire.Co2Lib);
   if (iWidth < (icanvas+30)) then iWidth := icanvas+30;
   GS.ColWidths[SG_Aff2]:=iWidth;
   end
else
   begin
   GS.ColWidths[SG_Aff2]:=0 ;
   GS.ColLengths[SG_Aff2]:=-1;
   end;

iWidth := ACT_AFFAIRE3.Width;
ACT_AFFAIRE3_.Width := ACT_AFFAIRE3.Width;
ACT_AFFAIRE3_.Visible := ACT_AFFAIRE3.Visible;
if VH_GC.CleAffaire.Co3Visible then
   begin
   icanvas:=GS.Canvas.TextWidth(VH_GC.CleAffaire.Co3Lib);
   if (iWidth < (icanvas+30)) then iWidth := icanvas+30;
   GS.ColWidths[SG_Aff3]:=iWidth;
   end
else
   begin
   GS.ColWidths[SG_Aff3]:=0 ;
   GS.ColLengths[SG_Aff3]:=-1;
   end;

iWidth := ACT_AVENANT.Width;
ACT_AVENANT_.Width := iWidth;
ACT_AVENANT_.Visible := ACT_AVENANT.Visible;
if (ACT_AVENANT.visible=false) then
   begin
   if (GetParamSoc('SO_AFFGESTIONAVENANT')= true) then
        begin
        GS.ColWidths[SG_Avenant]:=0 ;
        GS.ColLengths[SG_Avenant]:=-1;
        end;
   end;

GS.Cells[SG_Aff0, 0]:= 'St';
GS.Cells[SG_Aff1, 0]:= VH_GC.CleAffaire.Co1Lib;
GS.Cells[SG_Aff2, 0]:= VH_GC.CleAffaire.Co2Lib;
GS.Cells[SG_Aff3, 0]:= VH_GC.CleAffaire.Co3Lib;
if (GetParamSoc('SO_AFFGESTIONAVENANT')= true) then
    GS.Cells[SG_Avenant, 0]:= 'Av';

finally
ACT_AFFAIRE0.visible:=bVisiblePart0;
ACT_AFFAIRE1.visible:=bVisiblePart1;
ACT_AFFAIRE2.visible:=bVisiblePart2;
ACT_AFFAIRE3.visible:=bVisiblePart3;
ACT_AVENANT.visible:=bVisiblePart4;
ACT_AFFAIRE0.OnChange:=eOnChange0;
ACT_AFFAIRE1.OnChange:=eOnChange1;
ACT_AFFAIRE2.OnChange:=eOnChange2;
ACT_AFFAIRE3.OnChange:=eOnChange3;
end;

if (ACT_AFFAIRE0_.visible=false) then
   ACT_AFFAIRE1_.Left:=ACT_AFFAIRE0_.Left
else
   ACT_AFFAIRE1_.Left:=ACT_AFFAIRE0_.Left + ACT_AFFAIRE0_.Width + 1;
if (ACT_AFFAIRE1_.visible=false) then
   ACT_AFFAIRE2_.Left:=ACT_AFFAIRE1_.Left
else
    ACT_AFFAIRE2_.Left:=ACT_AFFAIRE1_.Left + ACT_AFFAIRE1_.Width + 1;
if (ACT_AFFAIRE2_.visible=false) then
   ACT_AFFAIRE3_.Left:=ACT_AFFAIRE2_.Left
else
   ACT_AFFAIRE3_.Left:=ACT_AFFAIRE2_.Left + ACT_AFFAIRE2_.Width + 1;
if (ACT_AFFAIRE3_.visible=false) then
   ACT_AVENANT_.Left:=ACT_AFFAIRE3_.Left
else
   ACT_AVENANT_.Left:=ACT_AFFAIRE3_.Left + ACT_AFFAIRE3_.Width + 1;


// Positionnement du champ libellé de l'affaire suivant la visibilité et la position des champs de saisie de l'affaire
if (ACT_AVENANT_.visible) then
   ACT_AFFAIREL.Left := ACT_AVENANT_.Left + ACT_AVENANT_.Width + 4
else
if (ACT_AFFAIRE3_.visible) then
   ACT_AFFAIREL.Left := ACT_AFFAIRE3_.Left + ACT_AFFAIRE3_.Width + 4
else
if (ACT_AFFAIRE2_.visible) then
   ACT_AFFAIREL.Left := ACT_AFFAIRE2_.Left + ACT_AFFAIRE2_.Width + 4
else
if (ACT_AFFAIRE1_.visible) then
   ACT_AFFAIREL.Left := ACT_AFFAIRE1_.Left + ACT_AFFAIRE1_.Width + 4;

end;

procedure TFActivite.ACT_TIERSChange(Sender: TObject);
begin
if (ACT_TIERS.visible=true) then
   begin
   gbSelectCellCalend:=false;
   ClipBoard.Clear;
   end;

if (ACT_TIERS.Text='') then LIBELLETIERS.Caption:='';
end;

procedure TFActivite.ACT_TIERSExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
if (ACT_TIERS.Text=CleAct.Tiers) then exit;

//if Not RemplirTOBTiersP(TOBTiers,TOBTierss,ACT_TIERS.Text) or (ACT_TIERS.Text='') then
//   if Not IdentifieTiers(ACT_TIERS) then
//      begin
//      ACT_TIERS.SetFocus;
//      Abort;
//      end;

if Not ChargeTiers(ACT_TIERS, ACT_TIERS.Text) then
    begin
    ACT_TIERS.SetFocus;
    Abort;
    end;

  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

// On vide les champs affaire : si on a changé le tiers, il faut reselectionner une affaire
ACT_AFFAIRE.Text:='';
ACT_AFFAIRE0.Text:='';
ACT_AFFAIRE1.Text:='';
ACT_AFFAIRE2.Text:='';
ACT_AFFAIRE3.Text:='';
ACT_AVENANT.Text:='';


LibelleTiers.Caption:=TOBTiers.GetValue('T_LIBELLE') ;
ACT_TIERSL.Caption :=TOBTiers.GetValue('T_LIBELLE') ;
MOIS_CLOT_CLIENT_.Text := TOBTiers.GetValue('T_MOISCLOTURE');
MOIS_CLOT_CLIENT.Text := TOBTiers.GetValue('T_MOISCLOTURE');
CleAct.Tiers:=ACT_TIERS.Text;

RafraichirLaSelection;
end;


{===========================================================================================}
{=============================== Gestion de l'unité ========================================}
{===========================================================================================}
Function TFActivite.IdentifieUnites(Control:TControl) : boolean ;
var
TOBL:TOB;
sTypeArt:string;
BEGIN
TOBL:=GetTOBLigne(GS.Row) ;
sTypeArt := TOBL.GetValue('ACT_TYPEARTICLE');

Result:=GetUnite(Control, HTitres.Mess[13], CleAct.TypeAcces, sTypeArt);
if Result then ChargeUnites(Control) ;
//GereUnitesEnabled ;
END ;

procedure TFActivite.ChargeUnites (Control:TControl) ;
Var
    OldU : String ;
    CodeUnite:string;
    bOK:boolean;
    TOBA:TOB;
BEGIN
TOBA:=nil;
bOK:=true;
if (Control is THGrid) then
   with THGrid(Control) do
      CodeUnite := Cells[Col, Row]
else
   CodeUnite :=THCritMaskEdit(Control).Text;

OldU:=TOBUnites.GetValue('GME_MESURE') ;
if Not RemplirTOBUnites( TOBListeConv, CodeUnite, TOBUnites) then
   begin
   // Si on ne trouve pas l'unite dans la liste principale, on cherche dans la liste
   // des unites de type 'AUC'
   if (TOBListeConvAUC<>nil) then
      TOBA := TOBListeConvAUC.FindFirst(['GME_MESURE'],[CodeUnite],TRUE) ;
   if (TOBA<>nil) then
      begin
      CodeUnite:=TOBA.GetValue('GME_MESURE') ;
      end
   else
      begin
      HActivite.Execute(11,Caption,'') ;
      bOK:=false;
      CodeUnite:='' ;
      end;
   end
else
   CodeUnite:=TOBUnites.GetValue('GME_MESURE') ;

if (Control is THGrid) then
   with THGrid(Control) do
      Cells[Col, Row] := CodeUnite
else
   begin
   THCritMaskEdit(Control).Text := CodeUnite;
   if (bOK=False) and (THCritMaskEdit(Control).canFocus) then
      THCritMaskEdit(Control).SetFocus;
   end;

if (Action=taCreat) then
   BEGIN
   if (CleAct.TypeSaisie = tsaRess) then
        TOBActivite.GetEcran(self, PTotaux3)
   else
        TOBActivite.GetEcran(self, PTotaux1) ;
   END ;
END ;


{===========================================================================================}
{=============================== Gestion du Tiers/client ===================================}
{===========================================================================================}
function TFActivite.ChargeTiers (Control : TControl; AcsCodeTiers : string) : boolean;
  var
  CodeTiers : string;
BEGIN
  Result := true;
  if (Control is THGrid) then
    begin
      if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc ('SO_AFRECHTIERSPARNOM') = true) then
        CodeTiers := AcsCodeTiers
      else
        with THGrid (Control) do
          CodeTiers := Cells [col, Row];
    end
  else
    CodeTiers := THCritMaskEdit (Control).Text;

  (*if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc('SO_AFRECHTIERSPARNOM')=true) and (Control is THGrid) then
      begin
      if Not RemplirTOBTiersLibP(TOBTiers,TOBTierss, CodeTiers) then
         begin
         HActivite.Execute(10,Caption,'') ;
         CodeTiers:=TOBTiers.GetValue('T_TIERS') ;
         Result:=false;
         end;
      end
  else
      begin*)
  //   if Not IdentifieTiers(ACT_TIERS) then
  if Not RemplirTOBTiersP (TOBTiers, TOBTierss, CodeTiers) then
    begin
      HActivite.Execute (10,Caption, '');
      CodeTiers := TOBTiers.GetValue ('T_TIERS');
      Result := false;
    end;
    //    end;

  LIBELLETIERS.Caption := TOBTiers.GetValue ('T_LIBELLE');
  ACT_TIERSL.Caption := TOBTiers.GetValue ('T_LIBELLE');
  ACT_TIERS_.text := TOBTiers.GetValue ('T_TIERS');
  MOIS_CLOT_CLIENT.Text := TOBTiers.GetValue ('T_MOISCLOTURE');

  if (Control is THGrid) then
    begin
      if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc ('SO_AFRECHTIERSPARNOM') = true) then
        with THGrid (Control) do
          Cells [Col, Row] := TOBTiers.GetValue ('T_LIBELLE')
      else
        with THGrid (Control) do
          Cells [Col, Row] := CodeTiers
    end
  else
    begin
      THCritMaskEdit (Control).Text := CodeTiers;
      THCritMaskEdit (Control).Font.Color := clBlack;
      if (TOBTiers <> nil) then
        if (CodeTiers = TOBTiers.GetValue ('T_TIERS')) then
          if (TOBTiers.GetValue ('T_FERME') = 'X') then
            begin
              THCritMaskEdit (Control).Font.Color := clRed;
              PGIInfoAF ('Attention, vous avez sélectionné un client fermé.', Caption);
            end;

      if (Result = False) and (THCritMaskEdit (Control).canFocus) then
        THCritMaskEdit (Control).SetFocus;
    end;

  if (Action = taCreat) then
    BEGIN
      TOBActivite.GetEcran (self, PEntete);
      TOBActivite.GetEcran (self, PEntete2);
    END;
END;

procedure TFActivite.ACT_TIERS_Change (Sender : TObject);
begin
  if (ACT_TIERS_.Text = '') then
    begin
      ACT_TIERSL.caption := '';
      MOIS_CLOT_CLIENT_.Text := '';
    end;

end;

Function TFActivite.IdentifieTiers ( Control:TControl  ) : boolean ;
  var
  bRechParCode : boolean;
  sCodeTiers : string;
BEGIN
  bRechParCode := true;
  sCodeTiers := '';
  if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc ('SO_AFRECHTIERSPARNOM') = true) then
    bRechParCode := false;

  Result := GetTiersRechercheAF (Control, ACT_TIERS.Hint, bRechParCode, sCodeTiers);

  if Result then
    BEGIN
      ChargeTiers(Control, sCodeTiers);
      ChargeAffaireActivite(True, true);
    END;

END;

procedure TFActivite.ACT_TIERSElipsisClick(Sender: TObject);
begin
IdentifieTiers(ACT_TIERS) ;
end;

procedure TFActivite.ACT_TIERSDblClick(Sender: TObject);
begin
IdentifieTiers(ACT_TIERS) ;
end;

{==============================================================================================}
{=============================== Manipulation des articles ====================================}
{==============================================================================================}

Function TFActivite.IdentifierArticle ( Control:TControl; Var ACol,ARow : integer ; Var Cancel : boolean ; Click : Boolean ) : boolean ;
Var RefSais, RefUnique,CodeArticle, TypeArticle : String ;
    TOBArt  : TOB ;
    QQ : TQuery ;
    RechArt : T_RechArt ;
BEGIN
Result:=False ;
TOBArt:=nil;
if (Control is THGrid) then
   RefSais := THGrid(Control).Cells[SG_Article,ARow]
else
   RefSais := THCritMaskEdit(Control).Text;

try
if (Click=false) then
   if (RefSais='') then exit;

TypeArticle:=GetChampLigne('ACT_TYPEARTICLE',ARow);

RefUnique:='' ; CodeArticle:='' ;
TOBArt:=FindTOBArtSaisAff(TOBArticles, RefSais, false) ;
if TOBArt<>Nil then
   BEGIN
   RechArt:=traOk ;
   END else
   BEGIN
   TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.InitValeurs(FALSE) ;
   RechArt:=TrouverArticleSQL_GI(false, RefSais, TOBArt, TypeArticle) ;
   END ;
Case RechArt of
       traOk : BEGIN
               CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
               RefUnique:=TOBArt.GetValue('GA_ARTICLE') ;
               END ;
    traAucun : BEGIN
               if (Control is THGrid) then begin THGrid(Control).Col:=SG_Article ; THGrid(Control).Row:=ARow ; end;
//               else THCritMaskEdit(Control).SetFocus;

               if GereElipsis(Control, SG_Article) then
                  BEGIN
                  if (Control is THGrid) then RefUnique:=THGrid(Control).Cells[SG_Article,ARow]
                                          else RefUnique:= THCritMaskEdit(Control).Text;

                  TOBArt:=FindTOBArtSaisAff(TOBArticles, RefUnique, true) ;
                  if TOBArt<>Nil then
                     BEGIN
                     CodeArticle:=TOBArt.GetValue('GA_CODEARTICLE') ;
                     END else
                     BEGIN
                     // SELECT * justifié par le fait qu'on ne ramène qu'un article et que beaucoup de champ sont nécessaires
                     // dans l'exploitation en aval, dont une partie pas encore développée ce qui nous obligerait à revenir ici pour
                     // rajouter des champs si on listait un par un tous ceux qui sont vraiment nécessaires
                     QQ:=OpenSQL('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+RefUnique+'" ',True,-1,'',true) ;
                     if Not QQ.EOF then
                        BEGIN
                        TOBArt:=CreerTOBArt(TOBArticles) ; TOBArt.SelectDB('',QQ) ;
                        CodeArticle:=QQ.FindField('GA_CODEARTICLE').AsString ;
                        END;
                     Ferme(QQ) ;
                     END ;
                  if (Control is THGrid) then  THGrid(Control).Cells[SG_Article,ARow]:=CodeArticle else THCritMaskEdit(Control).Text := CodeArticle;
                  END ;
               END ;
   END ;

if ((CodeArticle='') or (RefUnique='')) then
   BEGIN
   Cancel:=True ;
   if (Control is THGrid) then begin THGrid(Control).Col:=ACol ; THGrid(Control).Row:=ARow ; end;
//   else THCritMaskEdit(Control).SetFocus;

   TOBArt.Free ;
   TOBArt:=nil;
   Exit ;
   END ;

Result:=True ;
finally
CodesArtToCodesLigne(Control,TOBArt,ARow, true) ;
end;
END ;


Procedure TFActivite.CodesArtToCodesLigne ( Control:TControl; TOBArt : TOB ; ARow : integer; AcbMajTot:boolean; AcbArtChange : boolean = false);
  Var
  TOBL : TOB;
BEGIN

  TOBL := GetTOBLigne (ARow);
  if (TOBL = nil) then exit;

  CopieArtDansLaLigne (TOBArt, TOBL);

  // Met à jour l'article lié
  if AcbArtChange then
    begin
      TOBL.PutValue ('ARTSLIES', '');

      TraiteLesLies (ARow);
    end;

  // Mise à jour des totaux
  if AcbMajTot then
    MajTotauxActivite;

  AfficheLaTOBLigne (ARow);
  TOBL.PutEcran (Self, PPied);
  PutEcranPPiedAvance (TOBL);
END;

procedure TFActivite.CopieArtDansLaLigne (TOBArt, TOBLigneAct : TOB; bForcerPrix : boolean = false);
begin

  if (TOBArt = nil) then
    begin
      TOBLigneAct.PutValue ('ACT_ARTICLE', '');
      TOBLigneAct.PutValue ('ACT_CODEARTICLE', '');
      TOBLigneAct.PutValue ('ACT_FORMULEVAR', '');
      gbSaisieQtePermise := true;
      TOBLigneAct.PutValue ('ACT_UNITE', '');
      TOBLigneAct.PutValue ('ACT_UNITEFAC', '');
      // PL le 13/06/03 : on n'efface plus le libelle quand il a été modifié
      if (gbModifLibArticle = false) and (TOBLigneAct.GetValue('ACT_NUMLIGNEUNIQUE') = 0) then
        TOBLigneAct.PutValue('ACT_LIBELLE', '');
      TOBLigneAct.PutValue ('ACT_ACTIVITEREPRIS', '');
      TOBLigneAct.PutValue ('ACT_QTE', 0);
      TOBLigneAct.PutValue ('ACT_QTEFAC', 0);
      TOBLigneAct.PutValue ('ACT_QTEUNITEREF', 0);
      TOBLigneAct.PutValue ('ACT_PUPRCHARGE', 0);
      TOBLigneAct.PutValue ('ACT_PUPR', 0);
      TOBLigneAct.PutValue ('ACT_PUPRCHINDIRECT', 0);
      TOBLigneAct.PutValue ('ACT_PUVENTE', 0);
      TOBLigneAct.PutValue ('ACT_PUVENTEDEV', 0);
      TOBLigneAct.PutValue ('ACT_TOTPRCHARGE', 0);
      TOBLigneAct.PutValue ('ACT_TOTPR', 0);
      TOBLigneAct.PutValue ('ACT_TOTPRCHINDI', 0);
      TOBLigneAct.PutValue ('ACT_TOTVENTE', 0);
      TOBLigneAct.PutValue ('ACT_TOTVENTEDEV', 0);
      TOBLigneAct.PutValue ('ACT_ACTIVITEEFFECT', 'X');
    end
  else
    begin
   // PL le 13/06/03 : Le libelle de l'article n'est modifié en contexte scot que s'il était vide
   // avant ou si on ne l'a pas changé à la main
   // en contexte Affaire, il est mis à jour dans ce cas et aussi quand le code a changé
   //
   if ((ctxScot in V_PGI.PGIContexte) and (TOBLigneAct.GetValue('ACT_NUMLIGNEUNIQUE')=0) and ((TOBLigneAct.GetValue('ACT_LIBELLE')='') or (gbModifLibArticle = false)) )
      or (not (ctxScot in V_PGI.PGIContexte) and (TOBLigneAct.GetValue('ACT_CODEARTICLE')<>TOBArt.GetValue('GA_CODEARTICLE')) or (TOBLigneAct.GetValue('ACT_LIBELLE')=''))
      or (bForcerPrix) then
        TOBLigneAct.PutValue('ACT_LIBELLE', TOBArt.GetValue('GA_LIBELLE')) ;
   //////////// Fin PL le 13/06/03
      TOBLigneAct.PutValue ('ACT_ARTICLE', TOBArt.GetValue ('GA_ARTICLE'));
      TOBLigneAct.PutValue ('ACT_CODEARTICLE', TOBArt.GetValue ('GA_CODEARTICLE'));
      TOBLigneAct.PutValue ('ACT_TYPEARTICLE', TOBArt.GetValue ('GA_TYPEARTICLE'));
      TOBLigneAct.PutValue ('ACT_FORMULEVAR', TOBArt.GetValue ('GA_FORMULEVAR'));
      gbSaisieQtePermise := true;
      if GetParamSoc('SO_AFVARIABLES') and (TOBArt.GetValue ('GA_FORMULEVAR') <> '') then
        gbSaisieQtePermise := false;
      if TOBArt.FieldExists('NUMORDRE') then  //ONYX-AFORMULEVAR
        TOBLigneAct.PutValue ('ACT_NUMORDRE',TOBArt.GetValue ('NUMORDRE'));
      if TOBArt.FieldExists('NUMPIECEORIG') then  // PL le 06/08/03 : pour CB ajout NumPieceOrig
        TOBLigneAct.PutValue ('ACT_NUMPIECEORIG',TOBArt.GetValue ('NUMPIECEORIG'));
      if TOBArt.FieldExists('MNTREMISE') then  //ONYX  // PL le 11/07/03
        TOBLigneAct.PutValue ('ACT_MNTREMISE',TOBArt.GetValue ('MNTREMISE'));

      if (TOBLigneAct.GetValue ('ACT_ACTIVITEREPRIS') <> TOBArt.GetValue ('GA_ACTIVITEREPRISE')) then
        begin
          // PL le 20/06/03 : en scot on n'initialise avec l'activite reprise de l'article uniquement si c'était vide
          // en GA, on l'initialise dans tous les cas
          if not (ctxScot in V_PGI.PGIContexte) or ((ctxScot in V_PGI.PGIContexte) and (TOBLigneAct.GetValue ('ACT_ACTIVITEREPRIS') = '')) then
            begin
              TOBLigneAct.PutValue ('ACT_ACTIVITEREPRIS', TOBArt.GetValue ('GA_ACTIVITEREPRISE')) ;
              gbModifActRep := true;
            end;
        end;

      if (TOBArt.GetValue ('GA_ACTIVITEEFFECT') = '') and (TOBArt.GetValue ('GA_TYPEARTICLE') = 'PRE') then
        TOBLigneAct.PutValue ('ACT_ACTIVITEEFFECT', 'X')
      else
        TOBLigneAct.PutValue ('ACT_ACTIVITEEFFECT', TOBArt.GetValue ('GA_ACTIVITEEFFECT'));

      if ((TOBValo <> nil) and (TOBValo.GetValue ('GA_CODEARTICLE') = TOBArt.GetValue ('GA_CODEARTICLE'))) then
        begin
          TOBLigneAct.PutValue ('ACT_PUPRCHARGE', TOBValo.GetValue ('GA_PMRP')) ;
          TOBLigneAct.PutValue ('ACT_PUPR', TOBLigneAct.GetValue ('ACT_PUPRCHARGE')) ;
          TOBLigneAct.PutValue ('ACT_PUPRCHINDIRECT', TOBLigneAct.GetValue ('ACT_PUPRCHARGE')) ;
          TOBLigneAct.PutValue ('ACT_PUVENTE', TOBValo.GetValue ('GA_PVHT')) ;
          TOBLigneAct.PutValue ('ACT_PUVENTEDEV', TOBValo.GetValue ('GA_PVHT')) ;

          // PL le 20/06/03 : on peut initialiser le PV par le PR dans certains cas...
          // Paramsoc + si le PV était vide
          if (GetParamSoc('SO_AFPRDANSPV') and (TOBLigneAct.GetValue('ACT_PUVENTE') = 0)) then
            begin
              TOBLigneAct.PutValue('ACT_PUVENTE', TOBLigneAct.GetValue('ACT_PUPRCHARGE'));
              TOBLigneAct.PutValue('ACT_PUVENTEDEV', TOBLigneAct.GetValue('ACT_PUPRCHARGE'));
            end;

          // if (TOBL.GetValue('ACT_UNITE')='') or (TOBL.GetValue('ACT_LIBELLE')<>TOBArt.GetValue('GA_LIBELLE')) then
          // Dans le cas de l'initialisation par la TOBValo, on ne modifie l'unite que si elle est vide ou
          // si le libellé de l'article a changé (=l'article vient de la ligne de pièce)
          // begin
          TOBLigneAct.PutValue('ACT_UNITE', TOBArt.GetValue('GA_QUALIFUNITEACT')) ;
          TOBLigneAct.PutValue('ACT_UNITEFAC', TOBArt.GetValue('GA_QUALIFUNITEACT')) ;
          //  end;
        end
     else
        begin
          if bForcerPrix then
            begin
              TOBLigneAct.PutValue ('ACT_PUPRCHARGE', TOBArt.GetValue ('GA_PMRP')) ;
              TOBLigneAct.PutValue ('ACT_PUPR', TOBLigneAct.GetValue ('ACT_PUPRCHARGE')) ;
              TOBLigneAct.PutValue ('ACT_PUPRCHINDIRECT', TOBLigneAct.GetValue ('ACT_PUPRCHARGE')) ;
              TOBLigneAct.PutValue ('ACT_PUVENTE', TOBArt.GetValue ('GA_PVHT')) ;
              TOBLigneAct.PutValue ('ACT_PUVENTEDEV', TOBArt.GetValue ('GA_PVHT')) ;
            end;
          TOBLigneAct.PutValue('ACT_UNITE', TOBArt.GetValue('GA_QUALIFUNITEACT')) ;
          TOBLigneAct.PutValue('ACT_UNITEFAC', TOBArt.GetValue('GA_QUALIFUNITEACT')) ;
        end;
     end;
end;

procedure TFActivite.ACT_CODEARTICLEChange(Sender: TObject);
begin
if (ACT_CODEARTICLE.Text='') then ACT_ARTICLEL.caption:='';
end;

procedure TFActivite.ACT_CODEARTICLEElipsisClick(Sender: TObject);
var
TOBA:TOB;
ACol, ARow : integer;
Cancel:boolean;
begin
if ACT_CODEARTICLE.readonly then exit;
ACol := SG_Article;
ARow := GS.Row;
Cancel:=false;

if IdentifierArticle(ACT_CODEARTICLE ,ACol,ARow,Cancel,True) then
   begin
   TOBA := FindTOBArtSaisAff(TOBArticles, ACT_ARTICLE.Text, true );
   if (TOBA<>nil) then
      ACT_ARTICLEL.Caption := TOBA.GetValue('GA_LIBELLE');
   end;

end;

procedure TFActivite.ACT_CODEARTICLEExit(Sender: TObject);
var
TOBA, TOBL :TOB;
ACol, ARow : integer;
Cancel:boolean;
begin
if csDestroying in ComponentState then Exit ;
FormateZoneSaisie (ACT_CODEARTICLE, SG_Article, GS.Row );
ACol := SG_Article;
ARow := GS.Row;
Cancel:=false;
TOBL:=GetTOBLigne(GS.Row) ;
if TOBL=Nil then exit;
if (ACT_CODEARTICLE.Text=TOBL.GetValue('ACT_CODEARTICLE')) then exit;


if IdentifierArticle(ACT_CODEARTICLE ,ACol,ARow,Cancel,False) then
   begin
   TOBA := FindTOBArtSaisAff(TOBArticles, ACT_CODEARTICLE.Text, false );
   if (TOBA<>nil) then
      ACT_ARTICLEL.Caption := TOBA.GetValue('GA_LIBELLE');
   end;

// Mise à jour des totaux
MajTotauxActivite;

MajTOBHeureCalendrier(GS.Row);
bActModifiee:=true;
end;

procedure TFActivite.ACT_CODEARTICLEEnter(Sender: TObject);
var
TOBT:TOB;
begin
if Not GereSaisieEnabled( true ) then abort;

// on memorise la TOB activite de la ligne avant toute modif
TOBT := GetTOBLigne(GS.Row) ;
TOBLigneActivite.Dupliquer(TOBT,TRUE,TRUE,TRUE) ;
ConversionTobActivite( TOBEnteteLigneActivite, 'ACT_DATEACTIVITE' );
end;

procedure TFActivite.ACT_LIBELLEEnter(Sender: TObject);
begin
if Not GereSaisieEnabled( true ) then abort;
end;

{==========================================================================================}
{=============================== Sauvegarde des champs ====================================}
{==========================================================================================}
function TFActivite.PutJour (ARow : integer; var AciJour, AciMois, AciAnnee : integer ) : integer;
  var
  TOBL : TOB ;
  dDateActivite : TDateTime;
  PremierJourSem : TDateTime;
  YY, Inc : integer;
  bDateValide : boolean;
begin
  Result := 0;
  dDateActivite := idate1900;
  {$IFDEF AGLINF545}
  PremierJourSem := PremierJourSemaineTempo (CleAct.Semaine, CleAct.Annee);
  {$ELSE}
  // PL le 23/05/02 : réparation de la fonction AGL550
  YY := CleAct.Annee;
  if (CleAct.Mois = 12) and (CleAct.Semaine = 1) then YY := YY + 1;
  if (CleAct.Mois = 1) and ((CleAct.Semaine = 52) or (CleAct.Semaine = 53)) then YY := YY - 1;
  PremierJourSem := PremierJourSemaine (CleAct.Semaine, YY);
  //////////////////////AGL550
  {$ENDIF}

  TOBL := GetTOBLigne (ARow);
  if TOBL = Nil then Exit ;
  if (GS.Cells [SG_Jour, ARow] = '') then exit;

  try
  Result := 1;

  if (VH_GC.AFTYPESAISIEACT = 'SEM') then
    begin
      try
        bDateValide := true;
        dDateActivite := EncodeDate (CleAct.Annee, CleAct.Mois, AciJour);
      except
        // Si la date générée n'est pas valide ou hors intervalle de validité, on doit gérer la date autrement
        // ce qui est l'objet du bloc suivant
        bDateValide:=false;
      end;

      if (bDateValide = false) or (dDateActivite < PremierJourSem) or (dDateActivite > PremierJourSem + 6) then
        // Controle sur la validité de la date ou sur la date dans la semaine courante
        begin
          if (AciJour >= 15) then Inc := -1 else Inc := 1;
          AciMois := CleAct.Mois + Inc;
          AciAnnee := CleAct.Annee;
          if (AciMois = 0) then
            begin
              AciMois := 12;
              AciAnnee := AciAnnee - 1;
            end
          else
            if (AciMois = 13) then
              begin
                AciMois := 1;
                AciAnnee := AciAnnee + 1;
              end;

          dDateActivite := EncodeDate (AciAnnee, AciMois, AciJour);
          if (dDateActivite < PremierJourSem) or (dDateActivite > PremierJourSem + 6) then
             begin
              Result := 9;
              exit;
             end;
        end;
    end
  else
    begin
      dDateActivite := EncodeDate (CleAct.Annee, CleAct.Mois, AciJour);
    end;
  (*
  if (dDateActivite<VH_GC.AFDateDebutAct) or (dDateActivite>VH_GC.AFDateFinAct) then
     // Controle sur l'intervalle de dates du paramétrage
     begin
     Result := 19;
     exit;
     end;
  *)
  TOBL.PutValue ('ACT_DATEACTIVITE', dDateActivite) ;
  TOBL.PutValue ('ACT_PERIODE', GetPeriode (dDateActivite));
  TOBL.PutValue ('ACT_SEMAINE', NumSemaine (dDateActivite));

  except
    // gestion de l'exception sous controle
    Result := 8;
  end;
end;

Procedure TFActivite.TraiteJour (ARow : integer; Var Cancel : boolean);
  var
  iJour, iMois, iAnnee : integer;
  iErr : integer;
  TOBL : TOB ;
  bOK : boolean;
  bBloqueSaisie : boolean;
begin

  iErr := 1;
  if (GS.Cells [SG_Jour, ARow] = '') then exit;

  TOBL := GetTOBLigne (ARow) ;
  if TOBL = Nil then Exit ;

  try
    iErr := 8;
    iJour := strtoint (GS.Cells [SG_Jour, ARow]);
    iMois := CleAct.Mois;
    iAnnee := CleAct.Annee;
    iErr := PutJour (ARow, iJour, iMois, iAnnee);

    case iErr of
       0      : Exit;
       1      :;  // OK
       8,9,19 :   // le format de La date n'est pas correcte ou
                  // La date saisie n'est pas dans la semaine courante ou
                  // La date saisie n'est pas dans l'intervalle defini au parametrage
                raise ERangeError.Create ('');
    end;

    if (CleAct.TypeSaisie = tsaClient) then
      if (ControleDateDansIntervalle (EncodeDate (iAnnee, iMois, iJour), true, false, true, bBloqueSaisie) <> 0) then
        begin
          GS.Col := SG_Jour;
          Cancel := true;
        end;

    // Positionne sur le jour courant dans le calendrier
    if ChargerCalendrierEnCours (iJour, iMois, iAnnee , ForcerRefreshCalendrier (iAnnee, iMois, iJour)) then
      MajApresChangeJourCalendrier;

    // En saisie par Client/mission et à partir du lanceur, s'il existe un assistant par défaut, on le remplit ici
    if (ctxlanceur in V_PGI.PGIContexte) and (CleAct.TypeSaisie = tsaClient) and (gsRessParDefaut <> '') then
      begin
        GS.Cells[SG_Ressource, ARow] := gsRessParDefaut;
        StCellCur := '';
        GSCellExit (nil, SG_Ressource, ARow, Cancel);
      end;


    // Si on a une ressource courante et qu'aucun article n'est encore saisi, on affiche l'article associé
    if (TOBL.GetValue ('ACT_ARTICLE') = '') and (TOBL.GetValue ('ACT_RESSOURCE') <> '') then
      InitAssistantDefaut (TOBL, ARow, bOK);

  except
    // Gestion d'une exception controlée et non répercutée
    GS.Col := SG_Jour;
    PGIBoxAf (HTitres.Mess [iErr], Caption);
    Cancel := true;
  end;

end;


Procedure TFActivite.TraiteAffaire (ACol, ARow : integer; Var Cancel : boolean);
  Var
  TOBL : TOB;
  Part0, Part1, Part2, Part3, Avenant, Tiers : string;
  CodeAffaire : string;
  NewCol, NewRow, rep : integer;
  bAffaireChange, bAffaire0Change, bAffaire1Change, bAffaire2Change, bAffaire3Change, bAvenantChange, bOK : boolean;
  iPartErreur : integer;
  Cancel2, bMissionComplete, bMissionChange, bBloqueSaisie : boolean;
//  ddate : TDateTime;
  TOBAffTempo : TOB;
BEGIN
  Cancel2 := false;
  bMissionComplete := false;
  bMissionChange := false;
  Part0 := ''; Part1 := ''; Part2 := ''; Part3 := ''; Avenant := '';
  NewCol := GS.Col; NewRow := GS.Row;
  TOBL := GetTOBLigne (ARow);
  if TOBL = Nil then
    Exit;
  Cancel := false;
  bOK := true;
  iPartErreur := 0;
  TOBAffTempo := nil;
  //ddate := TOBL.Getvalue('act_datemodif'); // pour suivre la date en DEBUG

  Part0 := GS.Cells[SG_Aff0, ARow];
  Part1 := GS.Cells[SG_Aff1, ARow];
  Part2 := GS.Cells[SG_Aff2, ARow];
  Part3 := GS.Cells[SG_Aff3, ARow];
  if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
      Avenant := GS.Cells[SG_Avenant, ARow];

  CodeAffaire := CodeAffaireRegroupe (Part0, Part1, Part2, Part3, Avenant, Action, false, false, false);

  // PL le 16/05/02 : doit forcer la valorisation dans le cas où la mission a changé
  if (CodeAffaire <> TOBL.GetValue('ACT_AFFAIRE')) then bMissionChange := true;
  // La mission est elle complete ?
  if (Part0 <> '') and (Part1 <> '')
      and ((VH_GC.CleAffaire.NbPartie < 2) or (Part2 <> ''))
      and ((VH_GC.CleAffaire.NbPartie < 3) or (Part3 <> ''))
      and ((GetParamSoc ('SO_AFFGESTIONAVENANT') = false) or (Avenant <> '')) then
    bMissionComplete := true;
  ////////////////////////////

    // On est en recherche du code client/mission mais pas par un Ctrl-F2 ou autre
    if ((CleAct.TypeSaisie = tsaRess)
      and (gCtrlF2 = false) and (gCtrlF3 = false) and (gCtrlF4 = false)
      and ((Part1='') or ((VH_GC.CleAffaire.NbPartie >= 2) and (Part2='')) or ((VH_GC.CleAffaire.NbPartie >= 3) and (Part3='')) ))
      and (GS.Col > Acol) then
        begin
          ACT_AFFAIRE0.text := Part0;
          ACT_AFFAIRE1.text := Part1;
          ACT_AFFAIRE2.text := Part2;
          ACT_AFFAIRE3.text := Part3;
          if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
            ACT_AVENANT.Text := Avenant;

          if GetAffaireEnteteSt (ACT_AFFAIRE0, ACT_AFFAIRE1, ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ACT_TIERS,
                          CodeAffaire, (Part0 = 'P'), VH_GC.AFProposAct, false, true,true,'SAT', true, false, true) then
              begin
                TOBAffTempo := TobAffaire;
                AFRemplirTOBAffaire (CodeAffaire, TobAffaire, TobAffaires);
              end
          else
              bOK := False;

        end;

    if (bOK = true) then
      begin
        // on verifie la date par rapport aux dates de la mission saisie et aux dates des paramètres
        // et blocage dus à la confidentialité
        rep := ControleDateDansIntervalle (TOBL.GetValue ('ACT_DATEACTIVITE'), true, false, true, bBloqueSaisie);
        if (rep <> 0) then
          // blocage
          begin
            iPartErreur := 1;
            bOK := False;
          end;

        if (TOBAffTempo <> nil) then
          begin
            TobAffaire := TOBAffTempo;
            if (bOK = true) then
              begin
                //ddate := TOBL.Getvalue('act_datemodif');  // pour suivre la date en DEBUG
                ACT_AFFAIRE.Text := CodeAffaire;
				{$IFDEF BTP}
                BTPCodeAffaireDecoupe (ACT_AFFAIRE.Text, Part0, Part1, Part2, Part3, Avenant, Action, False);
				{$ELSE}
                CodeAffaireDecoupe (ACT_AFFAIRE.Text, Part0, Part1, Part2, Part3, Avenant, Action, False);
                {$ENDIF}
                ACT_AFFAIRE0.text := Part0;
                ACT_AFFAIRE1.text := Part1;
                ACT_AFFAIRE2.text := Part2;
                ACT_AFFAIRE3.text := Part3;
                ACT_AVENANT.Text := Avenant;
                ACT_TIERS_.Text := TobAffaire.GetValue('AFF_TIERS');
              end;
        end;
      end;

  // Si chaque partie est bonne, on doit controler la cohérence du code affaire issu de la concatenation
  // Si on sort de la zone de saisie de l'affaire
  if (bOK = true) then
     begin
     ACT_AFFAIRE0.Text := Part0;
     ACT_AFFAIRE1.Text := Part1;
     ACT_AFFAIRE2.Text := Part2;
     ACT_AFFAIRE3.Text := Part3;
     ACT_AVENANT.Text := Avenant;
     ACT_AFFAIRE.Text := CodeAffaire;
     ACT_TIERS.Text := ACT_TIERS_.Text;

     ACT_AFFAIRE.Text:= DechargeCleAffaire(ACT_AFFAIRE0,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT,ACT_TIERS.Text,Action,False,True,VH_GC.AFProposAct,iPartErreur);
     if (Not ChargeAffaireActivite(True, true)) then
  //   if (Not ChargeAffaireActivite(True, true)) and  (iPartErreur<>0) then // PL le 03/07/03 : on ne gère plus ici le controle sur les parties du code affaire
     // Si le code affaire n'est pas coherent et qu'il s'agit d'une erreur sur une partie du code
        bOK:=false
     else
        begin
        Part0:= ACT_AFFAIRE0.Text;
        Part1:= ACT_AFFAIRE1.Text;
        Part2:= ACT_AFFAIRE2.Text;
        Part3:= ACT_AFFAIRE3.Text;
        Avenant:= ACT_AVENANT.Text;
        CodeAffaire:= ACT_AFFAIRE.Text;
        Tiers := ACT_TIERS.Text;
        end;

      GereAffaireEnabled;
     end;

  //ddate := TOBL.Getvalue('act_datemodif');  // pour suivre la date en DEBUG

  if bOK = true then
     begin
     with TOBL do
        begin
        if ((Part1='') and (Part2='') and (Part3='') and (ARow<>NewRow)) then
          begin Part0:=''; Avenant:=''; end;
        bAffaireChange:=GetValue('ACT_AFFAIRE')<>CodeAffaire;
        bAffaire0Change:=(GetValue('ACT_AFFAIRE0')<>Part0);
        bAffaire1Change:=GetValue('ACT_AFFAIRE1')<>Part1;
        bAffaire2Change:=GetValue('ACT_AFFAIRE2')<>Part2;
        bAffaire3Change:=GetValue('ACT_AFFAIRE3')<>Part3;
        bAvenantChange:=GetValue('ACT_AVENANT')<>Avenant;
        if (bAffaireChange or bAffaire0Change or bAffaire1Change or bAffaire2Change or bAffaire3Change or bAvenantChange) then
           begin
           PutValue('ACT_AFFAIRE',CodeAffaire);
           GS.Cells[SG_Affaire, ARow]:= CodeAffaire;
           PutValue('ACT_AFFAIRE0',Part0);
           GS.Cells[SG_Aff0, ARow]:= Part0;
           PutValue('ACT_AFFAIRE1',Part1);
           GS.Cells[SG_Aff1, ARow]:= Part1;
           PutValue('ACT_AFFAIRE2',Part2);
           if (SG_Aff2 <> -1) then
            GS.Cells[SG_Aff2, ARow]:= Part2;
           PutValue('ACT_AFFAIRE3',Part3);
           if (SG_Aff3 <> -1) then
            GS.Cells[SG_Aff3, ARow]:= Part3;
           PutValue('ACT_AVENANT',Avenant);
           if (SG_Avenant <> -1) then
            GS.Cells[SG_Avenant, ARow] := Avenant;
           end;

        if (GetValue('ACT_TIERS')<>Tiers) then
          if (SG_Tiers <> -1) then
              begin
              PutValue('ACT_TIERS', Tiers);

              if (GetParamSoc('SO_AFRECHTIERSPARNOM') = false) then
                begin
                  GS.Cells[SG_Tiers, ARow] := Tiers;
                end
              else
                begin
                  if (TOBTiers <> nil) then
                    GS.Cells[SG_Tiers, ARow] := TOBTiers.GetValue ('T_LIBELLE');
                end;
              end;
        end;

     GS.Col := NewCol ; GS.Row := NewRow ;
     end
  else
     begin
  //   GS.Cells[SG_Affaire,ARow] := TOBL.GetValue('ACT_AFFAIRE');
     GS.Col := ACol ; GS.Row := ARow ;
     if (iPartErreur <> 0) then
        begin
        Cancel := true;
        if (GS.Col<>SG_Aff0+iPartErreur) then
           begin
           Cancel := false;    // PL le 03/07/03 : Si on choisit la colonne de retour nous même, inutile d'annuler le déplacement
           GS.Col := SG_Aff0 + iPartErreur;
           end;
        end;
     end;

  // PL le 15/05/03 : changement de clé de la table ACTIVITE : on doit remettre le numero de ligne unique à 0 quand l'affaire a changée
  if (CleAct.TypeSaisie = tsaRess) and bMissionChange then
    begin
      TOBL.PutValue('ACT_NUMLIGNEUNIQUE', 0);
      // On vient d'une ligne deja saisie, on ne peut plus modifier le libelle
      gbModifLibArticle := true;
    end;

  //ddate := TOBL.Getvalue('act_datemodif'); // pour suivre la date en DEBUG

  // PL le 16/05/02 : doit forcer la valorisation dans le cas où la mission a changé
  // En saisie par ressource, Si on valorise par client/Mission, que la mission est complete
  // et qu'on a deja valorise, on efface l'article pour revaloriser
  // PL le 13/06/03 : on n'efface plus l'article que si on valorise par affaire (comité technique)
  if (CleAct.TypeSaisie = tsaRess) and bMissionComplete and bMissionChange and (TOBValo<>nil) and
  //    ((VH_GC.AFValoActPR<>'ART') or (VH_GC.AFValoActPV<>'ART')) then
       ((VH_GC.AFValoActPR='AFF') or (VH_GC.AFValoActPV='AFF')) then
          begin
          TOBL.PutValue('ACT_ARTICLE','');
          TOBL.PutValue('ACT_CODEARTICLE','');
          TOBL.PutValue('ACT_FORMULEVAR','');
          gbSaisieQtePermise := true;
          GS.Cells[SG_Article,ARow]:='';
          TraiteArticle(SG_Article,ARow, Cancel2);
          end;
  ////////////////////////

END ;

procedure TFActivite.EffaceClientMission (ARow : integer);
  var
  TOBL : TOB;
begin
  TOBL := GetTOBLigne (ARow);
  TOBL.PutValue('ACT_AFFAIRE0', '');
  TOBL.PutValue('ACT_AFFAIRE1', '');
  TOBL.PutValue('ACT_AFFAIRE2', '');
  TOBL.PutValue('ACT_AFFAIRE3', '');
  TOBL.PutValue('ACT_AVENANT', '');
  TOBL.PutValue('ACT_AFFAIRE', '');
  TOBL.PutValue('ACT_TIERS', '');
  if (SG_Aff0 <> -1) then
    GS.Cells[SG_Aff0, ARow] := '';
  if (SG_Aff1 <> -1) then
    GS.Cells[SG_Aff1, ARow] := '';
  if (SG_Aff2 <> -1) then
    GS.Cells[SG_Aff2, ARow] := '';
  if (SG_Aff3 <> -1) then
    GS.Cells[SG_Aff3, ARow] := '';
  if (SG_Avenant <> -1) then
    GS.Cells[SG_Avenant, ARow] := '';
  if (SG_Tiers <> -1) then
    GS.Cells[SG_Tiers, ARow] := '';
  if (SG_Affaire <> -1) then
    GS.Cells[SG_Affaire, ARow] := '';
  ACT_AFFAIRE0.Text := '';
  ACT_AFFAIRE0_.Text := '';
  ACT_AFFAIRE1.Text := '';
  ACT_AFFAIRE1_.Text := '';
  ACT_AFFAIRE2.Text := '';
  ACT_AFFAIRE2_.Text := '';
  ACT_AFFAIRE3.Text := '';
  ACT_AFFAIRE3_.Text := '';
  ACT_AVENANT.Text := '';
  ACT_AVENANT_.Text := '';
  ACT_AFFAIRE.Text := '';
  ACT_AFFAIRE_.Text := '';
  ACT_TIERS.Text := '';
  ACT_TIERS_.Text := '';
  if (ctxScot in V_PGI.PGIContexte) then
    GS.Col := SG_Tiers
  else
    GS.Col := SG_Aff1;
end;

function TFActivite.DernierePartieAffaireVisible : integer;
begin
  // PL le 24/09/03 : en contexte affaire on ne veut pas attendre la sortie du tiers pour trouver le code affaire
  (*if Not (ctxScot in V_PGI.PGIContexte) then
    // En contexte Affaire la dernière partie visible est forcément le client
    begin
      Result := SG_Tiers;
      exit;
    end
  else*)
    Result := SG_Aff1;

  case VH_GC.CleAffaire.NbPartie of
    1 : ;
    2 : if (VH_GC.CleAffaire.Co2Visible) then
          Result := SG_Aff2;
    3 :
      begin
      if (VH_GC.CleAffaire.Co3Visible) then
          Result := SG_Aff3
      else
      if (VH_GC.CleAffaire.Co2Visible) then
          Result := SG_Aff2;
      end;
  end;

  if (GetParamSoc ('SO_AFFGESTIONAVENANT') = true) then
    Result := SG_Avenant;

end;

procedure TFActivite.ViderPartiesInvisiblesCodeAffaire (TOBL : TOB; ARow : integer);
begin
  if (TOBL = nil) then Exit;

  case VH_GC.CleAffaire.NbPartie of
    1 : ; // Pas de gestion des parties non définies
    2 : // Si deux parties mais la deuxième non visible, on l'efface
        begin
          if not (VH_GC.CleAffaire.Co2Visible) then
            begin
              TOBL.PutValue ('ACT_AFFAIRE2', '');
              ACT_AFFAIRE2.text := '';
              ACT_AFFAIRE2_.text := '';
              TOBL.PutValue ('ACT_AFFAIRE', '');
              ACT_AFFAIRE.text := '';
              ACT_AFFAIRE_.text := '';
              GS.Cells[SG_Affaire, ARow] := '';
              if (SG_Aff2 <> -1) then
                GS.Cells[SG_Aff2, ARow] := '';
            end;
        end;
    3 : // Si trois parties on efface les parties non visibles
        begin
          if not (VH_GC.CleAffaire.Co2Visible) then
            begin
              TOBL.PutValue ('ACT_AFFAIRE2', '');
              ACT_AFFAIRE2.text := '';
              ACT_AFFAIRE2_.text := '';
              TOBL.PutValue ('ACT_AFFAIRE', '');
              ACT_AFFAIRE.text := '';
              ACT_AFFAIRE_.text := '';
              GS.Cells[SG_Affaire, ARow] := '';
              if (SG_Aff2 <> -1) then
                GS.Cells[SG_Aff2, ARow] := '';
            end;

          if not (VH_GC.CleAffaire.Co3Visible) then
            begin
              TOBL.PutValue ('ACT_AFFAIRE3', '');
              ACT_AFFAIRE3.text := '';
              ACT_AFFAIRE3_.text := '';
              TOBL.PutValue ('ACT_AFFAIRE', '');
              ACT_AFFAIRE.text := '';
              ACT_AFFAIRE_.text := '';
              GS.Cells[SG_Affaire, ARow] := '';
              if (SG_Aff3 <> -1) then
                GS.Cells[SG_Aff3, ARow] := '';
            end
        end;
  end;

end;

procedure TFActivite.TraitePartieAffaire (ACol, ARow : integer; Var Cancel : boolean);
  var
  TOBL : TOB;
  DerPartAff : integer;
  NomPartieCodeAffaire : string;
  THCME, THCME_ : THCritMaskEdit;
begin
  Cancel := false;
  TOBL := GetTOBLigne (ARow);
  if (TOBL = nil) then Exit;
  DerPartAff := DernierePartieAffaireVisible;

  // On determine le nom de la partie du code affaire sur laquelle on est
  if (ACol = SG_Aff0) then
    begin
      NomPartieCodeAffaire := 'ACT_AFFAIRE0';
      THCME := ACT_AFFAIRE0;
      THCME_ := ACT_AFFAIRE0_;
    end
  else
  if (ACol = SG_Aff1) then
    begin
      NomPartieCodeAffaire := 'ACT_AFFAIRE1';
      THCME := ACT_AFFAIRE1;
      THCME_ := ACT_AFFAIRE1_;
    end
  else
  if (ACol = SG_Aff2) then
    begin
      NomPartieCodeAffaire := 'ACT_AFFAIRE2';
      THCME := ACT_AFFAIRE2;
      THCME_ := ACT_AFFAIRE2_;
    end
  else
  if (ACol = SG_Aff3) then
    begin
      NomPartieCodeAffaire := 'ACT_AFFAIRE3';
      THCME := ACT_AFFAIRE3;
      THCME_ := ACT_AFFAIRE3_;
    end
  else
  if (ACol = SG_Avenant) then
    begin
      NomPartieCodeAffaire := 'ACT_AVENANT';
      THCME := ACT_AVENANT;
      THCME_ := ACT_AVENANT_;
    end
  else
  exit;

  // Si on a changé la zone, on vide les parties non visibles
  if (GS.Cells [ACol, ARow] <> TOBL.GetValue (NomPartieCodeAffaire)) then
    ViderPartiesInvisiblesCodeAffaire (TOBL, ARow);

  if (GS.Cells [ACol, ARow] = '') then
    begin
      TOBL.PutValue (NomPartieCodeAffaire, '');
      THCME.Text := '';
      THCME_.Text := '';
      TOBL.PutValue ('ACT_AFFAIRE', '');
      GS.Cells[SG_Affaire, ARow] := '';
    end
  else
    begin
      TOBL.PutValue (NomPartieCodeAffaire, GS.Cells [ACol, ARow]);
      THCME.Text := GS.Cells [ACol, ARow];
      THCME_.Text := GS.Cells [ACol, ARow];
    end;

    if (Acol = DerPartAff) and (GS.Col >= Acol) then
    // Si on est sur la dernière partie visible du code affaire
    // on traite le code en global et on sort
    begin
      TraiteAffaire (ACol, ARow, Cancel);
    end;

end;

procedure TFActivite.TraiteTiers ( ARow : integer ) ;
Var TOBL : TOB ;
NewCol,NewRow:integer;
bOK, RecodeAff, bChargeLaTob : boolean;
CodeAff, Part0,Part1,Part2,Part3,Part4 : string;
Cancel,bChangeStatut,bProposition,bTiersChange,bTiersPasVide:boolean;
BEGIN
Cancel:=False ;
bTiersChange:=false; bTiersPasVide:=false;
NewCol:=GS.Col ; NewRow:=GS.Row ;
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
bOK:=true;  RecodeAff := False; bChargeLaTob:=false;


// PL le 16/05/02 : doit forcer la valorisation dans le cas où le tiers a changé
if (TOBL.GetValue('ACT_TIERS')<>'') then bTiersPasVide:=true;
if (GetParamSoc('SO_AFRECHTIERSPARNOM')=false) then
    begin
    if (GS.Cells[SG_Tiers,ARow]<>TOBL.GetValue('ACT_TIERS')) then bTiersChange:=true;
    end
else
    begin
    if (TOBTiers<>nil) then
        if (GS.Cells[SG_Tiers,ARow]<>AnsiUpperCase(TOBTiers.GetValue('T_LIBELLE'))) then bTiersChange:=true;
    end;
////////////////////////////


if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc('SO_AFRECHTIERSPARNOM')=true) then
    begin
    if Not RemplirTOBTiersLibP(TOBTiers,TOBTierss, GS.Cells[SG_Tiers,ARow]) then
       begin
       GS.Col:=SG_Tiers;
       GS.Row:=ARow;
       bOK:=GereElipsis (GS, SG_Tiers);
       end;
    end
else
    begin
    if Not RemplirTOBTiersP(TOBTiers,TOBTierss, GS.Cells[SG_Tiers,ARow]) then
       begin
       GS.Col:=SG_Tiers;
       GS.Row:=ARow;
       bOK:=GereElipsis (GS, SG_Tiers);
       end;
    end;


if (bOK = true) and (bTiersChange = true) then
   begin
   // On enregistre le tiers
   TOBL.PutValue('ACT_TIERS',TOBTiers.GetValue('T_TIERS')) ;
   ACT_TIERS.Text := TOBL.GetValue('ACT_TIERS');
   MOIS_CLOT_CLIENT.Text := TOBTiers.GetValue('T_MOISCLOTURE');
   MOIS_CLOT_CLIENT_.Text := TOBTiers.GetValue('T_MOISCLOTURE');
   GS.Col:=NewCol ; GS.Row:=NewRow ;

   // on reverifie le couple Tiers/affaire
   if (VH_GC.CleAffaire.Co2Visible=false) then
        begin
        TOBL.PutValue('ACT_AFFAIRE2','') ;
        ACT_AFFAIRE2.Text:='';
        GS.Cells[SG_Aff2, ARow]:='';
        RecodeAff := True;
        end;
   if (VH_GC.CleAffaire.Co3Visible=false) then
        begin
        TOBL.PutValue('ACT_AFFAIRE3','') ;
        ACT_AFFAIRE3.Text:='';
        GS.Cells[SG_Aff3, ARow]:='';
        RecodeAff := true;
        end;
   if (DernierePartieAffaireVisible = SG_Tiers) then
    RecodeAff := true;

   if RecodeAff then
        begin
        Part0 := TOBL.GetValue('ACT_AFFAIRE0'); Part1 := TOBL.GetValue('ACT_AFFAIRE1');
        Part2 := TOBL.GetValue('ACT_AFFAIRE2'); Part3 := TOBL.GetValue('ACT_AFFAIRE3');
        Part4 := TOBL.GetValue('ACT_AVENANT');
        CodeAff:=CodeAffaireRegroupe(Part0, Part1 ,Part2 ,Part3,Part4, Action, false,false,false);
        TOBL.PutValue('ACT_AFFAIRE',CodeAff) ;
        ACT_AFFAIRE.Text:=CodeAff;
        GS.Cells[SG_Affaire, ARow]:=CodeAff;
        end;

        // On est en recherche auto du code client/mission mais pas par un Ctrl-F2
   if ((CleAct.TypeSaisie = tsaRess) and (GetParamSoc('SO_AFRECHCLIMISSAUTO')=true)
        and (gCtrlF2=false) and (gCtrlF3=false) and (gCtrlF4=false)
        and ((TOBL.GetValue('ACT_AFFAIRE1')='') or (TOBL.GetValue('ACT_AFFAIRE2')=''))) then
        begin
        if (VH_GC.AFProposAct) then bChangeStatut := true else bChangeStatut := False;
        if (ACT_AFFAIRE0.Text = 'P') then bProposition := True else bProposition := False;
        if GetAffaireEntete(ACT_AFFAIRE,ACT_AFFAIRE1,ACT_AFFAIRE2,ACT_AFFAIRE3,ACT_AVENANT,ACT_TIERS,bProposition,bChangeStatut,false,true, True,'SAT', true, true, true) then
            begin
			{$IFDEF BTP}
            BTPCodeAffaireDecoupe(ACT_AFFAIRE.Text, Part0,Part1, Part2, Part3,Part4, Action, False);
            {$ELSE}
            CodeAffaireDecoupe(ACT_AFFAIRE.Text, Part0,Part1, Part2, Part3,Part4, Action, False);
            {$ENDIF}
            ACT_AFFAIRE0.text:=Part0; ACT_AFFAIRE1.text := Part1; ACT_AFFAIRE2.text := Part2; ACT_AFFAIRE3.text := Part3; ACT_AVENANT.Text := Part4;
            bChargeLaTob:=true;
            if (CleAct.TypeAcces = tacGlobal) then
              GS.Col := SG_TypA
            else
              GS.Col := SG_Article;
            end;
        end;

   if bChargeLaTob or (TOBL.GetValue('ACT_AFFAIRE1')<>'') or (TOBL.GetValue('ACT_AFFAIRE2')<>'') or (TOBL.GetValue('ACT_AFFAIRE3')<>'') then
      if ChargeAffaireActivite(true, false) then
            bChargeLaTob:=true;


    if (bChargeLaTob=true) then
         begin
         TOBL.PutValue('ACT_AFFAIRE', ACT_AFFAIRE.Text);

         TOBL.PutValue('ACT_AFFAIRE0',ACT_AFFAIRE0.Text) ;
         TOBL.PutValue('ACT_AFFAIRE1',ACT_AFFAIRE1.Text) ;
         TOBL.PutValue('ACT_AFFAIRE2',ACT_AFFAIRE2.Text) ;
         TOBL.PutValue('ACT_AFFAIRE3',ACT_AFFAIRE3.Text) ;
         TOBL.PutValue('ACT_AVENANT',ACT_AVENANT.Text) ;
         TOBL.PutValue('ACT_TIERS', ACT_TIERS.Text);
         // mcd 03/04/02 ligne pas chargée GS.Col := ProchaineColonneVide;
         end;
   end;

// PL le 16/05/02 : doit forcer la valorisation dans le cas où le tiers a changé
// En saisie par ressource, Si on valorise par client/Mission ou par ressource (=pas par article) et que le tiers a changé,
// on efface l'article pour revaloriser
// PL le 13/06/03 : on n'efface plus l'article que si on valorise par affaire (comité technique)
if (CleAct.TypeSaisie = tsaRess) and bTiersPasVide and bTiersChange and
//    ((VH_GC.AFValoActPR<>'') and (VH_GC.AFValoActPV<>'ART')) then
     ((VH_GC.AFValoActPR='AFF') or (VH_GC.AFValoActPV='AFF')) then
        begin
        TOBL.PutValue('ACT_CODEARTICLE','');
        TOBL.PutValue('ACT_ARTICLE','');
        GS.Cells[SG_Article,ARow]:='';
        TraiteArticle(SG_Article,ARow, Cancel);
        end;
////////////////////////


AfficheLaTOBLigne(ARow) ;
//if (bChargeLaTob=true) then
//         begin        // mcd 03/04/02 pour e^tre sur colonne vide
//         GS.Col := ProchaineColonneVide;
//         end;

TOBL.PutEcran(Self, PPied) ;
PutEcranPPiedAvance (TOBL);
END ;

procedure TFActivite.TraiteQte (ARow : integer);
  var
  TOBL : TOB ;
  bForceMajPR : boolean;
  bForceMajPV : boolean;
  bPROK, bPVOK : boolean;
begin
  bPROK := true;
  bPVOK := true;
  TOBL := GetTOBLigne (ARow);
  if (TOBL = nil) then Exit;

  // Pour Onyx
  if GetParamSoc('SO_AFVARIABLES') and not gbSaisieQtePermise then
    begin
      GS.Cells [SG_Qte, ARow] := TOBL.GetValue ('ACT_QTE');
      FormateZoneSaisie (GS, SG_Qte, ARow);
      exit;
    end;

  bForceMajPR := TestMajTOTPR ( ARow, TOBL);
  bForceMajPV := TestMajTOTPV ( ARow, TOBL);

  if (GS.Cells [SG_Qte,ARow] = '') then
    begin
      TOBL.PutValue ('ACT_QTE', '0');
      TOBL.PutValue ('ACT_QTEFAC', '0');
      TOBL.PutValue ('ACT_QTEUNITEREF', '0');
    end
  else
    begin
      TOBL.PutValue ('ACT_QTE', Valeur(GS.Cells [SG_Qte, ARow]));
      TOBL.PutValue ('ACT_QTEFAC', Valeur(GS.Cells [SG_Qte, ARow]));
      if (TOBL.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
        TOBL.PutValue ('ACT_QTEUNITEREF', ConversionUnite (TOBL.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, Valeur (GS.Cells [SG_Qte, ARow])))
      else
        TOBL.PutValue ('ACT_QTEUNITEREF', '0');

        // Valorisation
      if (TOBValo = nil) then
        AppelMajTOBValo (TOBL, TOBL.GetValue ('ACT_CODEARTICLE'), bPROK, bPVOK);

      (*if (TOBValo = nil) then
        TOBValo := MajTOBValo (TOBL.GetValue ('ACT_DATEACTIVITE'), CleAct.TypeAcces, TOBL.GetValue ('ACT_AFFAIRE'),
                                TOBL.GetValue ('ACT_RESSOURCE'), TOBL.GetValue ('ACT_CODEARTICLE'),
                                TOBL.GetValue ('ACT_TYPEHEURE'), TOBArticles, TOBAffaires, nil, AFOAssistants,
                                true, VH_GC.AFValoActPR, VH_GC.AFValoActPV, bPROK, bPVOK);*)

      // on utilise le zoom dans le cas de la revalorisation par affaire, donc les booleens bPROK, bPVOK sont forcement
      // à true car on a du choisir un article même s'il y en avait plusieurs sur l'affaire en cours

      if (TOBValo <> nil) then
        begin
          CodesArtToCodesLigne (GS, TOBValo, ARow, false, true);

          ACT_ARTICLEL.Caption := TOBValo.GetValue ('GA_LIBELLE');
          bForceMajPR := true;
          bForceMajPV := true;
        end;
    end;

  MajTOTPR (ARow, TOBL, bForceMajPR);
  MajTOTPV (ARow, TOBL, bForceMajPV);

  // Mise à jour des totaux
  MajTotauxActivite;

end;

Procedure TFActivite.TraitePUCH ( ARow : integer ) ;
Var TOBL : TOB ;
bForceMaj:boolean;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
bForceMaj:=TestMajTOTPR ( ARow, TOBL);

if (GS.Cells[SG_PUCH,ARow]='') then
   TOBL.PutValue('ACT_PUPRCHARGE',0)
else
   TOBL.PutValue('ACT_PUPRCHARGE',GS.Cells[SG_PUCH,ARow]) ;


TOBL.PutValue('ACT_PUPR',TOBL.GetValue('ACT_PUPRCHARGE')) ;
TOBL.PutValue('ACT_PUPRCHINDIRECT',TOBL.GetValue('ACT_PUPRCHARGE')) ;

// PL le 20/06/03 : on peut initialiser le PV par le PR dans certains cas...
if (GetParamSoc('SO_AFPRDANSPV') and (TOBL.GetValue('ACT_PUVENTE') = 0)) then
  begin
    TOBL.PutValue('ACT_PUVENTE', TOBL.GetValue('ACT_PUPRCHARGE'));
    TOBL.PutValue('ACT_PUVENTEDEV', TOBL.GetValue('ACT_PUPRCHARGE'));
  end;

MajTOTPR ( ARow, TOBL, bForceMaj ) ;

// Mise à jour des totaux
MajTotauxActivite;
bActModifiee:=true;
END ;

Procedure TFActivite.TraitePV ( ARow : integer ) ;
Var TOBL : TOB ;
bForceMaj:boolean;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
bForceMaj:=TestMajTOTPV ( ARow, TOBL);

if (GS.Cells[SG_PV,ARow]='') then
   begin
   TOBL.PutValue('ACT_PUVENTE',0);
   TOBL.PutValue('ACT_PUVENTEDEV',0);
   end
else
   begin
   TOBL.PutValue('ACT_PUVENTE',GS.Cells[SG_PV,ARow]) ;
   TOBL.PutValue('ACT_PUVENTEDEV',TOBL.GetValue('ACT_PUVENTE')) ;
   end;

MajTOTPV ( ARow, TOBL, bForceMaj ) ;

// Mise à jour des totaux
MajTotauxActivite;
bActModifiee:=true;
END ;

function TFActivite.TestMajTOTPV ( ARow : integer; TOBL : TOB ):boolean ;
var
  fQte : double;
  cPUV : double;
  cTOTPVCalc : double;
  cTOTPVSauv : double;
begin
  Result := False;
  if TOBL = Nil then Exit;
  fQte := Valeur (TOBL.GetValue ('ACT_QTE'));
  cPUV := Valeur (TOBL.GetValue ('ACT_PUVENTE'));
  cTOTPVSauv := Valeur (TOBL.GetValue ('ACT_TOTVENTE'));
  cTOTPVCalc := fQte * cPUV;

  if (cTOTPVSauv <> cTOTPVCalc) then Result := True;
end;

Procedure TFActivite.MajTOTPV ( ARow : integer; TOBL : TOB; AbForceMaj:boolean ) ;
var
fQte:double;
cPUV:double;
cTOTPV:double;
begin
if TOBL=Nil then Exit ;
fQte := Valeur (TOBL.GetValue('ACT_QTE'));
cPUV := Valeur (TOBL.GetValue('ACT_PUVENTE'));
cTOTPV := fQte * cPUV;

if Not AbForceMaj then
   if (GS.Cells[SG_TotV,ARow]<>'') then
      if (GS.Cells[SG_TotV,ARow]<>'0') then exit;

TOBL.PutValue('ACT_TOTVENTE',cTOTPV) ;
TOBL.PutValue('ACT_TOTVENTEDEV',TOBL.GetValue('ACT_TOTVENTE')) ;
GS.Cells[SG_TotV,ARow] := floattostr(cTOTPV);
FormateZoneSaisie(GS,SG_TotV,ARow) ;
TOBL.PutEcran(Self, PPied) ;
PutEcranPPiedAvance(TOBL);
end;

function TFActivite.TestMajTOTPR ( ARow : integer; TOBL : TOB ):boolean ;
var
fQte:double;
cPUCH:double;
cTOTPRCalc:double;
cTOTPRSauv:double;
begin
Result := False;
// Test sur le changement de résultat du calcul
if TOBL=Nil then Exit ;
fQte := Valeur (TOBL.GetValue('ACT_QTE'));
cPUCH := Valeur (TOBL.GetValue('ACT_PUPRCHARGE'));
cTOTPRSauv := Valeur (TOBL.GetValue('ACT_TOTPRCHARGE'));
cTOTPRCalc := fQte * cPUCH;

if (cTOTPRSauv <> cTOTPRCalc) then Result := True;

end;

Procedure TFActivite.MajTOTPR ( ARow : integer; TOBL : TOB; AbForceMaj:boolean ) ;
var
fQte:double;
cPUCH:double;
cTOTPR:double;
begin
if TOBL=Nil then Exit ;
fQte := Valeur (TOBL.GetValue('ACT_QTE'));
cPUCH := Valeur (TOBL.GetValue('ACT_PUPRCHARGE'));
cTOTPR := fQte * cPUCH;

if Not AbForceMaj then
   if (GS.Cells[SG_TOTPRCH,ARow]<>'') then
      if (GS.Cells[SG_TOTPRCH,ARow]<>'0') then exit;

TOBL.PutValue('ACT_TOTPRCHARGE',cTOTPR) ;
TOBL.PutValue('ACT_TOTPR',TOBL.GetValue('ACT_TOTPRCHARGE')) ;
TOBL.PutValue('ACT_TOTPRCHINDI',TOBL.GetValue('ACT_TOTPRCHARGE')) ;
GS.Cells[SG_TOTPRCH,ARow] := floattostr(cTOTPR);
FormateZoneSaisie(GS,SG_TOTPRCH,ARow) ;
TOBL.PutEcran(Self, PPied) ;
PutEcranPPiedAvance(TOBL);
end;

Procedure TFActivite.TraiteTOTPRCH ( ARow : integer ) ;
Var TOBL : TOB ;
fQte:double;
cPUR:double;
cTOTPR:double;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;

// Si le contenu du champ a changé suite à une saisie manuelle, on doit recalculer le prix unitaire
if ((GS.Cells[SG_TOTPRCH,ARow]='') and (TOBL.GetValue('ACT_TOTPRCHARGE')<>0)) or
   (TOBL.GetValue('ACT_TOTPRCHARGE')<>GS.Cells[SG_TOTPRCH,ARow]) then
   begin
   fQte := Valeur (TOBL.GetValue('ACT_QTE'));
   cTOTPR := Valeur(GS.Cells[SG_TOTPRCH,ARow]);
   if (fQte=0) then
      begin
      fQte:=1;
      TOBL.PutValue('ACT_QTE', 1);
      TOBL.PutValue('ACT_QTEFAC', 1);
      GS.Cells[SG_Qte,ARow]:='1';
      FormateZoneSaisie(GS,SG_Qte,ARow) ;
      end;
   cPUR := cTOTPR / fQte;
   TOBL.PutValue('ACT_PUPRCHARGE',cPUR);
   TOBL.PutValue('ACT_PUPR',TOBL.GetValue('ACT_PUPRCHARGE')) ;
   TOBL.PutValue('ACT_PUPRCHINDIRECT',TOBL.GetValue('ACT_PUPRCHARGE')) ;
   GS.Cells[SG_PUCH,ARow] := floattostr(cPUR);
   FormateZoneSaisie(GS,SG_PUCH,ARow) ;
   end;

if (GS.Cells[SG_TOTPRCH,ARow]='') then
   TOBL.PutValue('ACT_TOTPRCHARGE',0)
else
   TOBL.PutValue('ACT_TOTPRCHARGE',GS.Cells[SG_TOTPRCH,ARow]) ;

TOBL.PutValue('ACT_TOTPR',TOBL.GetValue('ACT_TOTPRCHARGE')) ;
TOBL.PutValue('ACT_TOTPRCHINDI',TOBL.GetValue('ACT_TOTPRCHARGE')) ;

// Mise à jour des totaux
MajTotauxActivite;
bActModifiee:=true;
END ;

Procedure TFActivite.TraiteTotV( ARow : integer ) ;
Var TOBL : TOB ;
fQte:double;
cPUV:double;
cTOTPV:double;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;

// Si le contenu du champ a changé suite à une saisie manuelle, on doit recalculer le prix unitaire
if ((GS.Cells[SG_TotV,ARow]='') and (TOBL.GetValue('ACT_TOTVENTE')<>0)) or
   (TOBL.GetValue('ACT_TOTVENTE')<>GS.Cells[SG_TotV,ARow]) then
   begin
   fQte := Valeur (TOBL.GetValue('ACT_QTE'));
   cTOTPV := Valeur(GS.Cells[SG_TotV,ARow]);
   if (fQte=0) then
      begin
      fQte:=1;
      TOBL.PutValue('ACT_QTE', 1);
      TOBL.PutValue('ACT_QTEFAC', 1);
      GS.Cells[SG_Qte,ARow]:='1';
      FormateZoneSaisie(GS,SG_Qte,ARow) ;
      end;

   cPUV := cTOTPV / fQte;
   TOBL.PutValue('ACT_PUVENTE',cPUV);
   TOBL.PutValue('ACT_PUVENTEDEV',TOBL.GetValue('ACT_PUVENTE')) ;
   GS.Cells[SG_PV,ARow] := floattostr(cPUV);
   FormateZoneSaisie(GS,SG_PV,ARow) ;
   end;

if (GS.Cells[SG_TotV,ARow]='') then
   begin
   TOBL.PutValue('ACT_TOTVENTE',0);
   TOBL.PutValue('ACT_TOTVENTEDEV',0) ;
   end
else
   begin
   TOBL.PutValue('ACT_TOTVENTE',GS.Cells[SG_TotV,ARow]) ;
   TOBL.PutValue('ACT_TOTVENTEDEV',TOBL.GetValue('ACT_TOTVENTE')) ;
   end;

// Mise à jour des totaux
MajTotauxActivite;
bActModifiee:=true;
END ;

Procedure TFActivite.TraiteMntRemise (ARow : integer);
  Var
  TOBL : TOB;
BEGIN
  TOBL := GetTOBLigne(ARow); if TOBL = Nil then Exit;

  if (GS.Cells[SG_MntRemise, ARow] = '') then
    begin
      TOBL.PutValue ('ACT_MNTREMISE', 0);
    end
  else
    begin
      TOBL.PutValue ('ACT_MNTREMISE', GS.Cells[SG_MntRemise, ARow]);
    end;

  bActModifiee := true;
END ;

Procedure TFActivite.TraiteMontantTVA (ARow : integer);
  Var
  TOBL : TOB;
BEGIN
  TOBL := GetTOBLigne(ARow); if TOBL = Nil then Exit;

  if (GS.Cells[SG_MontantTVA, ARow] = '') then
    begin
      TOBL.PutValue ('ACT_MONTANTTVA', 0);
    end
  else
    begin
      TOBL.PutValue ('ACT_MONTANTTVA', GS.Cells[SG_MontantTVA, ARow]);
    end;

  bActModifiee := true;
END ;

Procedure TFActivite.TraiteLib( ARow : integer ) ;
Var TOBL : TOB ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
TOBL.PutValue('ACT_LIBELLE',GS.Cells[SG_Lib,ARow]) ;
gbModifLibArticle := true; // PL le 13/06/03 : le libelle article a ete modifie
END ;

Procedure TFActivite.TraiteActRep( ARow : integer ) ;
Var TOBL : TOB ;
NewCol,NewRow : integer ;
bOK:boolean;
LibAct:string;
BEGIN
NewCol:=GS.Col ; NewRow:=GS.Row ;
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;

bOK := true;
LibAct:=RechDom('AFTACTIVITEREPRISE',GS.Cells[SG_ActRep,ARow],FALSE);
if ((LibAct='Error') or (LibAct='')) and (GS.Cells[SG_ActRep,ARow] <>'') then
   begin
   GS.Col:=SG_ActRep ;
   GS.Row:=ARow ;
   bOK:=GereElipsis (GS, SG_ActRep);
   end;

if bOK=true then
   begin
   TOBL.PutValue('ACT_ACTIVITEREPRIS',GS.Cells[SG_ActRep,ARow]) ;
   GS.Col:=NewCol ; GS.Row:=NewRow ;
   end
else
   GS.Cells[SG_ActRep,ARow] := TOBL.GetValue('ACT_ACTIVITEREPRIS');

gbModifActRep := true;
   
TOBL.PutEcran(Self, PPied) ;
PutEcranPPiedAvance(TOBL);
bActModifiee:=true;
END ;

Procedure TFActivite.TraiteTypA( ARow : integer ) ;
Var TOBL : TOB ;
bModif:boolean;
NewCol,NewRow : integer ;
bOK,bArtOK:boolean;
LibTArt:string;
BEGIN
NewCol:=GS.Col ; NewRow:=GS.Row ;
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;
bModif:=false;
if (TOBL.GetValue('ACT_TYPEARTICLE')<>GS.Cells[SG_TypA,ARow]) then bModif:=true;

bOK := true;
bArtOK := false;
if (bModif=true) then
   begin
   LibTArt:=RechDom('GCTYPEARTICLE',GS.Cells[SG_TypA,ARow],FALSE);
   if ((LibTArt='Error') or (LibTArt='')) and (GS.Cells[SG_TypA,ARow] <>'') then
      begin
      GS.Col:=SG_TypA ;
      GS.Row:=ARow ;
      bOK:=GereElipsis (GS, SG_TypA);
      end;

   if bOK=true then
      begin
      TOBL.PutValue('ACT_TYPEARTICLE',GS.Cells[SG_TypA,ARow]) ;
      if (TOBL.GetValue('ACT_TYPEARTICLE') = 'PRE') then
            TOBL.PutValue('ACT_ACTIVITEEFFECT', 'X')
      else
            TOBL.PutValue('ACT_ACTIVITEEFFECT', '-');

      GS.Col:=NewCol ; GS.Row:=NewRow ;

      TOBL.PutValue('ACT_ARTICLE','') ;
      TOBL.PutValue('ACT_CODEARTICLE','') ;
      TOBL.PutValue('ACT_FORMULEVAR','') ;
      gbSaisieQtePermise := true;
      TOBL.PutValue('ACT_UNITE','') ;
      TOBL.PutValue('ACT_UNITEFAC','') ;
      // PL le 13/06/03 : on n'efface plus le libelle quand il a été modifié
      if (gbModifLibArticle = false) and (TOBL.GetValue('ACT_NUMLIGNEUNIQUE')=0) then
        TOBL.PutValue('ACT_LIBELLE','') ;
      TOBL.PutValue('ACT_ACTIVITEREPRIS','') ;

      if (TOBL.GetValue('ACT_RESSOURCE')<>'') then
         InitAssistantDefaut(TOBL, ARow, bArtOK);
      end
    else
      GS.Cells[SG_TypA,ARow] := TOBL.GetValue('ACT_TYPEARTICLE');

   // Mise à jour des totaux
   MajTotauxActivite;

   AfficheLaTOBLigne(ARow) ;
   TOBL.PutEcran(Self, PPied) ;
   PutEcranPPiedAvance(TOBL);
   end;
END ;


Procedure TFActivite.TraiteTypeHeure ( ARow : integer ) ;
  Var
  TOBL : TOB ;
  bOK : boolean;
  NewCol, NewRow : integer ;
  TxPR, TXPV : double;
  bForceMajPR : boolean;
  bForceMajPV : boolean;
  bPROK, bPVOK : boolean;
BEGIN
  bPROK := true; bPVOK := true;
  bForceMajPR := false; bForceMajPV := false;
  NewCol := GS.Col ; NewRow := GS.Row;
  TOBL := GetTOBLigne(ARow);
  // Cas de sortie de la fonction
  if (TOBL = Nil) then
    Exit;
  if GS.Cells[SG_TypeHeure,ARow] = TOBL.GetValue('ACT_TYPEHEURE') then
    Exit;

  bOK := true;
  if (GS.Cells[SG_TypeHeure, ARow] <> '') then
  if Not TauxHSupFromCode (GS.Cells[SG_TypeHeure, ARow], TxPR, TXPV) then
    begin
      GS.Col := SG_TypeHeure;
      GS.Row := ARow;
      bOK := GereElipsis (GS, SG_TypeHeure);
    end;

  if (bOK = true) then
    begin
      GS.Col := NewCol;
      GS.Row := NewRow;
      TOBL.PutValue ('ACT_TYPEHEURE', GS.Cells[SG_TYPEHEURE, ARow]);
      if (TOBValo <> nil) then
        begin
          TOBValo.Free;
          TOBValo := nil;
        end;

      if (TOBL.GetValue('ACT_QTE')<>0) then
        begin
          // Valorisation
          AppelMajTOBValo (TOBL, TOBL.GetValue ('ACT_CODEARTICLE'), bPROK, bPVOK);

          (* TOBValo:=MajTOBValo(TOBL.GetValue('ACT_DATEACTIVITE'), CleAct.TypeAcces, TOBL.GetValue('ACT_AFFAIRE'),
                                TOBL.GetValue('ACT_RESSOURCE'), TOBL.GetValue('ACT_CODEARTICLE'),
                                TOBL.GetValue('ACT_TYPEHEURE'), TOBArticles, TOBAffaires, nil, AFOAssistants,
                                true, VH_GC.AFValoActPR, VH_GC.AFValoActPV, bPROK, bPVOK);*)
          // on utilise le zoom dans le cas de la revalorisation par affaire, donc les booleens bPROK, bPVOK sont forcement
          // à true car on a du choisir un article même s'il y en avait plusieurs sur l'affaire en cours

          if (TOBValo<>nil) then
            begin
              CodesArtToCodesLigne(GS,TOBValo,ARow, false, true);

              ACT_ARTICLEL.Caption := TOBValo.GetValue('GA_LIBELLE');
              bForceMajPR:=true;
              bForceMajPV:=true;
            end;

          MajTOTPR ( ARow, TOBL, bForceMajPR  ) ;
          MajTOTPV ( ARow, TOBL, bForceMajPV  ) ;

          // Mise à jour des totaux
          MajTotauxActivite;
        end;
    end;

  AfficheLaTOBLigne(ARow) ;
  TOBL.PutEcran(Self, PPied) ;
  PutEcranPPiedAvance(TOBL);
  bActModifiee:=true;
END ;

Procedure TFActivite.TraiteRessource ( ARow : integer; Var Cancel : boolean );
Var TOBL : TOB ;
bOK, bArtOK:boolean;
NewCol,NewRow : integer ;
bRessChange,bRessPasVide :boolean;
BEGIN
Cancel:=False ;
bRessChange:=false; bRessPasVide:=false;
NewCol:=GS.Col ; NewRow:=GS.Row ;
TOBL:=GetTOBLigne(ARow) ; if TOBL=Nil then Exit ;

bOK := true;
if RemplirTOBAssistant(TOBAssistant,AFOAssistants, GS.Cells[SG_Ressource,ARow])<0 then
   begin
   GS.Col:=SG_Ressource ;
   GS.Row:=ARow ;
   bOK:=GereElipsis (GS, SG_Ressource);
   end;

// PL le 14/05/02 : doit forcer la valorisation dans le cas où la ressource a change
if (TOBL.GetValue('ACT_RESSOURCE')<>'') then bRessPasVide:=true;
if (GS.Cells[SG_Ressource,ARow]<>TOBL.GetValue('ACT_RESSOURCE')) then bRessChange:=true;
////////////////////////////

if bOK=true then
   begin
   if (TOBAssistant.GetValue('ARS_FERME')='X') then
        begin
        //PGIBoxAf(HTitres.Mess[26], Caption);
        PgiInfoAF ('L''assistant sélectionné est fermé.', Caption);
        GS.Col:=SG_Ressource ;
        GS.Row:=ARow ;
        GS.Cells[SG_Ressource,ARow]:='';
        TOBL.PutValue('ACT_RESSOURCE', '') ;
        exit;
        end
   else
        begin
        TOBL.PutValue('ACT_RESSOURCE',GS.Cells[SG_Ressource,ARow]) ;
        TOBL.PutValue('ACT_TYPERESSOURCE',TOBAssistant.GetValue('ARS_TYPERESSOURCE')) ;
        // PL le 29/01/02
        CleAct.Ressource := GS.Cells[SG_Ressource,ARow];
        // Fin PL le 29/01/02
        GS.Col:=NewCol ; GS.Row:=NewRow ;
        end;
   end
else
  // PL le 11/07/03
  begin
    ColOld := SG_Jour; // En cas d'erreur il faut forcer le repassage dans le Cellexit pour controler la nouvelle saisie
    Cancel := true;
    exit;
  end;

// PL le 14/05/02 : doit forcer la valorisation dans le cas où la ressource a changé
// En saisie par client/Mission, Si on valorise par assistant et que l'assistant a changé, on efface l'article pour revaloriser
if (CleAct.TypeSaisie = tsaClient) and bRessPasVide and bRessChange and
    ((VH_GC.AFValoActPR='RES') or (VH_GC.AFValoActPV='RES')) then
        begin
        TOBL.PutValue('ACT_ARTICLE','');
        TOBL.PutValue('ACT_CODEARTICLE','');
        GS.Cells[SG_Article,ARow]:='';
        TOBL.PutValue('ACT_FORMULEVAR','');
        gbSaisieQtePermise := true;
        TraiteArticle(SG_Article,ARow, Cancel);
        end;
////////////////////////

bArtOK:=false;
// Si on a une ressource courante et qu'aucun article n'est encore saisi, on affiche l'article associé
if (TOBL.GetValue('ACT_ARTICLE')='') and (TOBL.GetValue('ACT_RESSOURCE')<>'') then
   InitAssistantDefaut(TOBL, ARow, bArtOK);

// Si la mise à jour de l'article est OK, ce n'est pas la peine de repasser dans les fonctions d'affichage
// c'est deja fait
if Not bArtOK then
   begin
   AfficheLaTOBLigne(ARow) ;
   TOBL.PutEcran(Self, PPied) ;
   PutEcranPPiedAvance(TOBL);
   end;
END ;

Procedure TFActivite.TraiteArticle (Var ACol,ARow : integer ; Var Cancel : boolean ) ;
  Var
  OkArt, bMemoQtePermise, bArtChange : Boolean;
  NewCol, NewRow : integer;
  TOBA, TOBL : TOB;
  bPROK, bPVOK : boolean;
BEGIN
  bPROK := true;
  bPVOK := true;
  TOBL := GetTOBLigne (ARow);
  if TOBL = Nil then Exit;

  NewCol := GS.Col;
  NewRow := GS.Row;
  if (TOBValo <> nil) then
    if (TOBValo.GetValue ('GA_CODEARTICLE') <> GS.Cells [SG_Article, ARow]) then
      // Si le code a changé, la tobvalo n'est plus valide
      begin
        TOBValo.Free;
        TOBValo := nil;
      end;


  // Valorisation
  bArtChange := AppelMajTOBValo (TOBL, GS.Cells [SG_Article, ARow], bPROK, bPVOK);

(*
    // PL le 17/06/03 : ONYX
    // Le code article n'a pas change et la formule est deja prévue pourtant la tob zoom est vide
    // = on a deja choisi un article : il faut reinitialiser une tob avec les elements qu'on
    // a deja
  bArtChange := (TOBL.getvalue('ACT_CODEARTICLE') <> GS.Cells [SG_Article, ARow]);
  if (TOBValoZoom = nil)
      and ((TOBL.getvalue('ACT_FORMULEVAR') <> '') or (TOBL.getvalue('ACT_NUMORDRE') <> 0))
      and not bArtChange then
    begin
      TOBValoZoom := TOB.Create ('ARTICLE', Nil, -1);
      TOBValoZoom.AddChampSup('NUMORDRE', false);
      TOBValoZoom.AddChampSup('NUMPIECEORIG', false); // PL le 06/08/03 : pour CB ajout NumPieceOrig
      TOBValoZoom.AddChampSup('MNTREMISE', false);  // PL le 11/07/03

      TOBValoZoom.PutValue('GA_CODEARTICLE', TOBL.getvalue('ACT_CODEARTICLE'));
      TOBValoZoom.PutValue('GA_LIBELLE', TOBL.getvalue('ACT_LIBELLE'));
      TOBValoZoom.PutValue('GA_PMRP', TOBL.getvalue('ACT_PUPRCHARGE'));   // 0
      TOBValoZoom.PutValue('GA_PVHT', TOBL.getvalue('ACT_PUVENTE'));       // 0
      TOBValoZoom.PutValue('NUMORDRE', TOBL.getvalue('ACT_NUMORDRE'));
      TOBValoZoom.PutValue('NUMPIECEORIG', TOBL.getvalue('ACT_NUMPIECEORIG')); // PL le 06/08/03 : pour CB ajout NumPieceOrig

      TOBValoZoom.PutValue('MNTREMISE', TOBL.getvalue('ACT_MNTREMISE')); // PL le 11/07/03
      TOBValoZoom.PutValue('GA_FORMULEVAR', TOBL.getvalue('ACT_FORMULEVAR'));
    end;

  if (TOBValo = nil) and (GS.Cells [SG_Article, ARow] <> '') then
    TOBValo := MajTOBValo (TOBL.GetValue ('ACT_DATEACTIVITE'), CleAct.TypeAcces, TOBL.GetValue ('ACT_AFFAIRE'),
                            TOBL.GetValue ('ACT_RESSOURCE'), GS.Cells [SG_Article, ARow], TOBL.GetValue ('ACT_TYPEHEURE'),
                            TOBArticles, TOBAffaires, TOBValoZoom, AFOAssistants, true, VH_GC.AFValoActPR,
                            VH_GC.AFValoActPV, bPROK, bPVOK);

  // on utilise le zoom dans le cas de la revalorisation par affaire, donc les booleens bPROK, bPVOK sont forcement
  // à true car on a du choisir un article même s'il y en avait plusieurs sur l'affaire en cours

  if (TOBValoZoom <> nil) then
    begin
      TOBValoZoom.Free;
      TOBValoZoom := nil;
    end;
*)
  if (TOBValo <> nil) then
    begin
      GS.Col := NewCol;
      GS.Row := NewRow;

      CodesArtToCodesLigne (GS, TOBValo, ARow, true, bArtChange);

      // Mise à jour des totaux par ligne
      MajTOTPR (ARow, TOBL, true);
      MajTOTPV (ARow, TOBL, true);
      // Mise à jour des totaux
      MajTotauxActivite;

      ACT_ARTICLEL.Caption := TOBValo.GetValue ('GA_LIBELLE');

      if (GetParamSoc('SO_AFVARIABLES')) and (not gbSaisieQtePermise) then
        begin
          TraiteFormuleVar (ARow, 'MODIF');
          bMemoQtePermise := gbSaisieQtePermise;
          gbSaisieQtePermise := true;
          TraiteQte (ARow);
          gbSaisieQtePermise := bMemoQtePermise;
        end;

      exit;
    end;


  // Cas de la saisie normale sans qu'on ait pu valoriser
  OkArt := IdentifierArticle (GS, ACol, ARow, Cancel, False);
  if ((OkArt) and (Not Cancel)) then
    begin
      GS.Col := NewCol;
      GS.Row := NewRow;

      TOBA := FindTOBArtSaisAff (TOBArticles, GS.Cells [ACol, ARow], false);

      if (TOBA <> nil) then
        ACT_ARTICLEL.Caption := TOBA.GetValue ('GA_LIBELLE')
      else
        ACT_ARTICLEL.Caption := '';

    end;

END;

function TFActivite.AppelMajTOBValo (TOBL : TOB; CodeArticle : string; var bPROK, bPVOK : boolean) : boolean;
begin

    // PL le 17/06/03 : ONYX
    // Le code article n'a pas change et la formule est deja prévue pourtant la tob zoom est vide
    // = on a deja choisi un article : il faut reinitialiser une tob avec les elements qu'on
    // a deja
  Result := (TOBL.getvalue('ACT_CODEARTICLE') <> CodeArticle);
  if (TOBValoZoom = nil)
      and ((TOBL.getvalue('ACT_FORMULEVAR') <> '') or (TOBL.getvalue('ACT_NUMORDRE') <> 0))
      and not Result then
    begin
      TOBValoZoom := TOB.Create ('ARTICLE', Nil, -1);
      TOBValoZoom.AddChampSup('NUMORDRE', false);
      TOBValoZoom.AddChampSup('NUMPIECEORIG', false); // PL le 06/08/03 : pour CB ajout NumPieceOrig
      TOBValoZoom.AddChampSup('MNTREMISE', false);  // PL le 11/07/03

      TOBValoZoom.PutValue('GA_CODEARTICLE', TOBL.getvalue('ACT_CODEARTICLE'));
      TOBValoZoom.PutValue('GA_LIBELLE', TOBL.getvalue('ACT_LIBELLE'));
      TOBValoZoom.PutValue('GA_PMRP', TOBL.getvalue('ACT_PUPRCHARGE'));   // 0
      TOBValoZoom.PutValue('GA_PVHT', TOBL.getvalue('ACT_PUVENTE'));       // 0
      TOBValoZoom.PutValue('NUMORDRE', TOBL.getvalue('ACT_NUMORDRE'));
      TOBValoZoom.PutValue('NUMPIECEORIG', TOBL.getvalue('ACT_NUMPIECEORIG')); // PL le 06/08/03 : pour CB ajout NumPieceOrig

      TOBValoZoom.PutValue('MNTREMISE', TOBL.getvalue('ACT_MNTREMISE')); // PL le 11/07/03
      TOBValoZoom.PutValue('GA_FORMULEVAR', TOBL.getvalue('ACT_FORMULEVAR'));
    end;

  if (TOBValo = nil) and (CodeArticle <> '') then
    TOBValo := MajTOBValo (TOBL.GetValue ('ACT_DATEACTIVITE'), CleAct.TypeAcces, TOBL.GetValue ('ACT_AFFAIRE'),
                            TOBL.GetValue ('ACT_RESSOURCE'), CodeArticle, TOBL.GetValue ('ACT_TYPEHEURE'),
                            TOBArticles, TOBAffaires, TOBValoZoom, AFOAssistants, true, VH_GC.AFValoActPR,
                            VH_GC.AFValoActPV, bPROK, bPVOK);

  // on utilise le zoom dans le cas de la revalorisation par affaire ou dans la cas où l'on veut récupérer des données déjà saisies,
  // donc les booleens bPROK, bPVOK sont forcement
  // à true car on a du choisir un article même s'il y en avait plusieurs sur l'affaire en cours

  if (TOBValoZoom <> nil) then
    begin
      TOBValoZoom.Free;
      TOBValoZoom := nil;
    end;
end;

Procedure TFActivite.TraiteFormuleVar(ARow : integer; Action : String);  //ONYX-AFORMULEVAR
Var TOBL : TOB ;
    ACol : integer;
    Cancel, bMemoQtePermise : Boolean;
    Qte : double ;
    FormuleVar,CodeAffaire : string;
begin
  if not (GetParamSoc('SO_AFVARIABLES')) then exit;
  TOBL := GetTOBLigne (ARow);
  if (TOBL = Nil) then Exit;
          CodeAffaire := TOBL.GetValue ('ACT_AFFAIRE');
  FormuleVar  := TOBL.GetValue ('ACT_FORMULEVAR');

  if (FormuleVar <> '') and  (CodeAffaire <> '') then
  begin
    if (TOBFormuleVar = nil) or (TOBFormuleVarQte = nil) then
    begin
      LoadLesAFFormuleVar (FormuleVar,TOBL,TOBFormuleVar,TOBFormuleVarQte);
    end;

    if (Action ='SUPPRIME') then
    begin
      if (TobFormuleVarQte <> nil) and (TobFormuleVarQte.GetValue('AVV_NUMLIGNEUNIQUE') = TOBL.GetValue('ACT_NUMLIGNEUNIQUE')) then
        TobFormuleVarQte.DeleteDB;
      TobFormuleVarQte.Free;
      TobFormuleVarQte:=nil;
      if (TOBFormuleVar <> nil) then
        TOBFormuleVar.cleardetail;
      Exit;
    end;

    Qte := EvaluationAFFormuleVar ('ACT',FormuleVar,Action,TOBFormuleVar,TOBFormuleVarQte,TOBL);
    if (Qte >= 0) and (SG_QF > 0) then
    begin
//      GS.Cells[SG_QTE, ARow] := StrF00 (Qte, V_PGI.OkDecQ);
      GS.Cells[SG_QTE, ARow] := floattostr(Qte); // PL le 14/08/03 pour problème arrondi ONYX
      if Action = 'MODIF' then
        begin
//          ACol := GS.Col;
          Cancel := False;
//          GS.Col := ACol + 1;
          GS.Col := SG_QTE;
          ACol := SG_QTE;
          bMemoQtePermise := gbSaisieQtePermise;
          gbSaisieQtePermise := true;
          try
            GSCellExit (Nil, ACol, ARow, Cancel);
          finally
            gbSaisieQtePermise := bMemoQtePermise;
          end;
        end
      else
      if Action <> 'CONSULT' then
        begin
          bMemoQtePermise := gbSaisieQtePermise;
          gbSaisieQtePermise := true;
          try
            TraiteQte (ARow);
          finally
            gbSaisieQtePermise := bMemoQtePermise;
          end;
        end;
    end ;
  end;
end;

{=======================================================================================}
{======================================= VALIDATIONS ===================================}
{=======================================================================================}
(*function TFActivite.GereValider : TIOErr;
  var
  reponse : TModalResult;
begin
  Result := oeOk;
  //if TOBActivite.Detail.Count=0 then exit;
  // Test s'il y a eu modification
  if TOBActivite <> nil then

  //if Not TOBActivite.IsOneModifie then exit
  if Not bActModifiee then exit
  else
     // demande confirmation de l'enregistrement des modifs
     begin
       Reponse := HActivite.Execute (8, Caption, '');
       case Reponse of
            mrYes : ;
            mrNo  : exit;
            else
              begin
                Result := oeUnknown;
                exit;
              end;
       end;
     end;

  Result := Transactions (ValideLActivite, 5);
  Case Result of
        oeOk :      BEGIN
                      bActModifiee := false;
                      CreerTOBLignes (GS.Row);
                    END;
        oeUnknown : BEGIN
                      PGIBoxAf (HTitres.Mess[5], Caption);
                      Exit;
                    END;
        oeSaisie :  BEGIN
                      PGIBoxAf (HTitres.Mess[6], Caption);
                      Exit;
                    END;
  end;

end;
*)
(*
procedure TFActivite.ValideLActivite ;
  var
  i : integer;
BEGIN
  if Action = taConsult then Exit;
  try
    SourisSablier;
    InitMove (3, '');

    TOBActivite.GetEcran (Self, PEntete);
    TOBActivite.GetEcran (Self, PEntete2);

    DepileTOBLignes (GS.Row, 1);
    VideLignesIntermediaires;

    // Suppression des lignes de TOBActiviteOld dans la table
    // PL le 03/05/02 : suppression des lignes modifiées par un delete général
    //if Not TOBActiviteOld.DeleteDB then V_PGI.IoError:=oeSaisie ;
    DeleteLignesSuivantCriteres;
    /////////////////////////////////
    MoveCur (False);

    if V_PGI.IoError = oeOk then
      BEGIN
        // PL le 23/10/02 : ajout du try pour trapper les exceptions sur la boucle de maj de la tob activite
        // et les remonter à l'AGL
        try
          for i := 0 to TOBActivite.Detail.count - 1 do
            begin
              // Gestion automatique du visa sur l'activité
              if (GetParamSoc('SO_AFVISAACTIVITE') = false) then
                begin
                  TOBActivite.Detail[i].PutValue ('ACT_ETATVISA', 'VIS');
                  TOBActivite.Detail[i].PutValue ('ACT_VISEUR', V_PGI.User);
                  TOBActivite.Detail[i].PutValue ('ACT_DATEVISA', NowH);
                end
              else
                begin
                  if TOBActivite.Detail[i].GetValue ('ACT_ETATVISA') = '' then
                    TOBActivite.Detail[i].PutValue ('ACT_ETATVISA', 'ATT');
                end;

              // Gestion automatique du visa de facturation sur l'activité
              if (GetParamSoc('SO_AFAPPPOINT') = false) then
                begin
                  TOBActivite.Detail[i].PutValue ('ACT_ETATVISAFAC', 'VIS');
                  TOBActivite.Detail[i].PutValue ('ACT_VISEURFAC', V_PGI.User);
                  TOBActivite.Detail[i].PutValue ('ACT_DATEVISAFAC', NowH);
                end
              else
                begin
                  if TOBActivite.Detail[i].GetValue ('ACT_ETATVISAFAC') = '' then
                    TOBActivite.Detail[i].PutValue ('ACT_ETATVISAFAC', 'ATT');
                end;
            end;


          TOBActivite.SetAllModifie (True) ;
        except
          // PL le 23/10/02 : ajout du try pour trapper les exceptions sur la boucle de maj de la tob activite
          // et les remonter à l'AGL
          V_PGI.IoError := oeSaisie;
        end;

        MoveCur (False);

        // Insertion des lignes de TOBActivite dans la table
        if V_PGI.IoError = oeOk then
          begin
            if Not TOBActivite.InsertDB (nil) then
              V_PGI.IoError := oeUnknown
            else
              begin
                TOBActivite.SetAllModifie (False);
              end;
          end;

        MoveCur (False);
      END;

  finally
    FiniMove;
    SourisNormale;
  end;
END;
*)
function TFActivite.GereValiderLaLigne (NumLigne : integer) : TIOErr;
var
  OldTransErrorMessage : string;
begin
  Result := oeOk;

  // Test s'il y a eu modification
  if Not bActModifiee then exit;

  try
    // Aucun message ne doit s'afficher en jaune dans les erreurs SQL
    OldTransErrorMessage := V_PGI.TransacErrorMessage;
    V_PGI.TransacErrorMessage := '';
    giLigneAEnregistrer := NumLigne;
    Result := Transactions (ValideLaLigneActivite, 5);

    Case Result of
      oeOk :      begin
                    bActModifiee := false;
                    //CreerTOBLignes (GS.Row);
                  end;
      oeUnknown : begin
                    PGIBoxAf (HTitres.Mess[5], Caption);
                    Exit;
                  end ;
      oeSaisie :  begin
                    PGIInfoAf ('Attention, la ligne courante a été modifiée par un autre utilisateur.' + Chr(13)
                      + 'Votre sélection va être rafraîchie et vous perdrez vos modifications sur cette ligne.' + Chr(13)
                      + 'Veuillez vous assurer que vous pouvez modifier cette sélection sans interaction avec d''autres utilisateurs.', Caption);

                    RafraichirLaSelection;
                    Exit;
                  end ;
    end ;

  finally
    // remise en état de l'ancien message SQL error
    V_PGI.TransacErrorMessage := OldTransErrorMessage;
  end;
end;

procedure TFActivite.ValideLaLigneActivite;
  var
  NumLigneUnique : integer;
  TobLigneEnBase, TobLigneEnGrid : TOB;
  bNouvelle : boolean;
  dDateModifEnBase, dDateModifEnGrid, dDateVisaEnBase, dDateVisaEnGrid, dDateVisaFacEnBase, dDateVisaFacEnGrid : TDateTime;
BEGIN
  TobLigneEnBase := nil;
  if Action = taConsult then Exit;

  try
    SourisSablier;
    InitMove (3, '');

    TobLigneEnGrid := GetTOBLigne (giLigneAEnregistrer);
    if (TobLigneEnGrid = nil) then exit;

  //
  // Vérification de la stabilité de la ligne entre le moment où on a commencé à la modifier
  // et le moment où on la valide
  //
    // On ramène la tob de la ligne telle qu'elle est en base pour vérifier que personne ne l'a modifiée entre temps
    // prévoir un verrou ici en cas de volonté future de gestion de blocage de la ligne modifiée
    // PL le 04/06/03
    if (TOBLigneActivite <> nil) then
      TobLigneEnBase := LaTobDeLaLigneUnique (TOBLigneActivite.GetValue ('ACT_TYPEACTIVITE'),
                                              TOBLigneActivite.GetValue ('ACT_AFFAIRE'),
                                              TOBLigneActivite.GetValue ('ACT_NUMLIGNEUNIQUE'));

    if (TobLigneEnBase <> nil) then
      begin
        dDateModifEnBase := TOBLigneEnBase.GetValue ('ACT_DATEMODIF');
        dDateModifEnGrid := TobLigneEnGrid.GetValue ('ACT_DATEMODIF');
        dDateVisaEnBase := TOBLigneEnBase.GetValue ('ACT_DATEVISA');
        dDateVisaEnGrid := TobLigneEnGrid.GetValue ('ACT_DATEVISA');
        dDateVisaFacEnGrid := TobLigneEnGrid.GetValue ('ACT_DATEVISAFAC');
        dDateVisaFacEnBase := TOBLigneEnBase.GetValue ('ACT_DATEVISAFAC');

                // Si cette ligne existe deja, on vérifie qu'elle n'a pas été modifiée entre temps
        if (dDateModifEnGrid <> dDateModifEnBase)
            or ((GetParamSoc ('SO_AFVISAACTIVITE') = True) and (dDateVisaEnGrid <> dDateVisaEnBase))
            or ((GetParamSoc ('SO_AFAPPPOINT') = True) and (dDateVisaFacEnGrid <> dDateVisaFacEnBase))
          then
          // Si la date de modification a changé, on alerte l'utilisateur qu'une modification est intervenue
          // et on rafraichit la saisie en perdant les modifications de la ligne en cours
          begin
            V_PGI.IoError := oeSaisie;
            exit;
          end;
      end;
   MoveCur (False);

  //
  // on complète la ligne suivant le contexte
  //
  // Gestion automatique du visa sur l'activité
  if (GetParamSoc ('SO_AFVISAACTIVITE') = false) then
    begin
      TobLigneEnGrid.PutValue ('ACT_ETATVISA', 'VIS');
      TobLigneEnGrid.PutValue ('ACT_VISEUR', V_PGI.User);
      TobLigneEnGrid.PutValue ('ACT_DATEVISA', NowH);
    end
  else
    begin
      if TobLigneEnGrid.GetValue ('ACT_ETATVISA') = '' then
        TobLigneEnGrid.PutValue ('ACT_ETATVISA', 'ATT');
    end;

  // Gestion automatique du visa de facturation sur l'activité
  if (GetParamSoc ('SO_AFAPPPOINT') = false) then
    begin
    TobLigneEnGrid.PutValue ('ACT_ETATVISAFAC', 'VIS');
    TobLigneEnGrid.PutValue ('ACT_VISEURFAC', V_PGI.User);
    TobLigneEnGrid.PutValue ('ACT_DATEVISAFAC', NowH);
    end
  else
    begin
      if TobLigneEnGrid.GetValue ('ACT_ETATVISAFAC') = '' then
        TobLigneEnGrid.PutValue ('ACT_ETATVISAFAC', 'ATT');
    end;

  bNouvelle := false;
  // PL le 04/06/03
  if (TobLigneEnBase <> nil) then
    begin
      if ((TobLigneEnGrid.GetValue ('ACT_TYPEACTIVITE') <> TobLigneEnBase.GetValue ('ACT_TYPEACTIVITE'))
          or (TobLigneEnGrid.GetValue ('ACT_AFFAIRE') <> TobLigneEnBase.GetValue ('ACT_AFFAIRE'))) then
        // Si un element de la cle a changé, il faut supprimer l'ancienne ligne avant toutes choses
        begin
          bNouvelle := true;
          if not TobLigneEnBase.DeleteDB then
            begin
              V_PGI.IoError := oeSaisie;
              exit;
            end
          else
            TobLigneEnGrid.ChargeCle1; // PL le 10/07/03 : on recharge absolument la clé !!!
        end
      else
        // Il se peut qu'on ait changé le code mission puis remis l'ancien
        // dans ce cas la clé n'a pas changée, il suffit de remettre l'ancien numéro unique
      if (TobLigneEnGrid.GetValue ('ACT_NUMLIGNEUNIQUE') = 0) and (TobLigneEnBase.GetValue ('ACT_NUMLIGNEUNIQUE') <> 0) then
          TobLigneEnGrid.PutValue ('ACT_NUMLIGNEUNIQUE', TobLigneEnBase.GetValue ('ACT_NUMLIGNEUNIQUE'));
    end;

  if (TobLigneEnBase = nil) or (bNouvelle = true) then
    // Si la ligne n'existait pas en base ou si la cle a changé
    begin
      bNouvelle := true;

      //if (TobLigneEnGrid.GetValue ('ACT_NUMLIGNEUNIQUE') = 0) then
        // si aucun numéro n'est prévu, il faut trouver le numéro unique qui lui correspond
        //begin
      NumLigneUnique := ProchainPlusNumLigneUniqueActivite (TobLigneEnGrid.GetValue('ACT_TYPEACTIVITE'), TobLigneEnGrid.GetValue('ACT_AFFAIRE'));
      TobLigneEnGrid.PutValue ('ACT_NUMLIGNEUNIQUE', NumLigneUnique);
        //end;
    end;
  /////////////// Fin PL le 04/06/03

  CompleteLaTOB (TobLigneEnGrid);  // Met a jour la dateModif et le user entre autre

  if V_PGI.IoError = oeOk then
    MajTotauxMensuels (TobLigneEnBase, TobLigneEnGrid);


  MoveCur (False);

  if (V_PGI.IoError = oeOk) then   //ONYX-AFORMULEVAR
    // la fonction est conditionnée par if GetParamSoc('SO_AFVARIABLES')
    ValideLesAFFormuleVar (TobLigneEnGrid, TobLigneEnBase, TOBFormuleVarQte, bNouvelle);

  //
  // On valide la ligne en base
  //
  // Insertion des lignes de TOBActivite dans la table
  if V_PGI.IoError = oeOk then
    begin
      if bNouvelle then
        begin
          if not TobLigneEnGrid.InsertDB (nil) then
            V_PGI.IoError := oeSaisie
        end
      else
        begin
          if Not TobLigneEnGrid.UpdateDB then
            V_PGI.IoError := oeUnknown;
        end;


      if V_PGI.IoError = oeOk then
        begin
          TobLigneEnGrid.SetAllModifie (False);
          TOBActivite.SetAllModifie (False);
        end;
    end;

  MoveCur (False) ;


  finally
    TobLigneEnBase.free;
    FiniMove;
    SourisNormale;
  end;
END ;


procedure TFActivite.MajTotauxMensuels (TobLigneEnBase, TobLigneEnGrid : TOB);
  var
  StValeur, sdate, sQte, sUnite, sActEff, sActRep, sdFac, sdNFac : string;
  dNFact, dFact : double;
  TobNewHeures, TobOldHeures : TOB;
//  ddate : TDateTime;
begin
  TobNewHeures := nil;
  TobOldHeures := nil;

  // On met à jour la tob de suivi des montants mensuels (total, F, NF, ...)
  // on cherche si on a deja une tob pour cette date
  if (TobLigneEnBase <> nil) then
    begin
      // ddate := TobLigneEnBase.GetValue('ACT_DATEACTIVITE');   // Pour DEBUG

      TobOldHeures := TOBNbHeureJourAffiche.FindFirst (['ACT_DATEACTIVITE'],
                                                       [TobLigneEnBase.GetValue('ACT_DATEACTIVITE')],
                                                       false);
    end;

  if (TobLigneEnGrid <> nil) then
    begin
      // ddate := TobLigneEnGrid.GetValue('ACT_DATEACTIVITE'); // pour DEBUG

      TobNewHeures := TOBNbHeureJourAffiche.FindFirst (['ACT_DATEACTIVITE'],
                                                       [TobLigneEnGrid.GetValue('ACT_DATEACTIVITE')],
                                                       false);
    end;

  // on commence par enlever des comptes la ligne en base si elle existait
  if (TobOldHeures <> nil) then
    if (TobLigneEnBase.GetValue('ACT_ACTIVITEEFFECT') = 'X') then
      begin
        TobOldHeures.PutValue ('ACT_QTE', TobOldHeures.GetValue ('ACT_QTE') - ConversionUnite (TobLigneEnBase.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnBase.GetValue('ACT_QTE')));
        if (TobLigneEnBase.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          begin
           TobOldHeures.PutValue ('ACT_QTENFACTURE', TobOldHeures.GetValue ('ACT_QTENFACTURE') - ConversionUnite (TobLigneEnBase.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnBase.GetValue('ACT_QTE')));
          end
        else
          TobOldHeures.PutValue ('ACT_QTEFACTURE', TobOldHeures.GetValue ('ACT_QTEFACTURE') - ConversionUnite (TobLigneEnBase.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnBase.GetValue('ACT_QTE')));
      end;

  // on ajoute la nouvelle ligne s'il y en a une et s'il s'agit de travail effectif
  if (TobLigneEnGrid.GetValue('ACT_ACTIVITEEFFECT') = 'X') then
    begin
    if (TobNewHeures <> nil) then
      begin
        TobNewHeures.PutValue ('ACT_QTE', TobNewHeures.GetValue ('ACT_QTE') + ConversionUnite (TobLigneEnGrid.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnGrid.GetValue('ACT_QTE')));
        if (TobLigneEnGrid.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          TobNewHeures.PutValue ('ACT_QTENFACTURE', TobNewHeures.GetValue ('ACT_QTENFACTURE') + ConversionUnite (TobLigneEnGrid.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnGrid.GetValue('ACT_QTE')))
        else
          TobNewHeures.PutValue ('ACT_QTEFACTURE', TobNewHeures.GetValue ('ACT_QTEFACTURE') + ConversionUnite (TobLigneEnGrid.GetValue('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, TobLigneEnGrid.GetValue('ACT_QTE')));
      end
    else
      // Il n'y a pas encore de ligne de compte pour cette date, on la créé
      begin
        dNFact := 0;
        dFact := 0;
        if (TobLigneEnGrid.GetValue('ACT_ACTIVITEREPRIS') = 'N') then
          dNFact := TobLigneEnGrid.GetValue('ACT_QTE')
        else
          dFact := TobLigneEnGrid.GetValue('ACT_QTE');

        TobNewHeures := TOB.Create ('Activites par jour', TOBNbHeureJourAffiche, -1);
        sdate := floattostr(TobLigneEnGrid.GetValue('ACT_DATEACTIVITE'));
        sQte := vartostr(TobLigneEnGrid.GetValue('ACT_QTE'));
        sUnite := TobLigneEnGrid.GetValue('ACT_UNITE');
        sActEff := TobLigneEnGrid.GetValue('ACT_ACTIVITEEFFECT');
        sActRep := TobLigneEnGrid.GetValue('ACT_ACTIVITEREPRIS');
        sdFac := floattostr(dFact);
        sdNFac := floattostr(dNFact);
        StValeur := sdate + ';'
                    + sUnite + ';'
                    + sQte + ';'
                    + sActEff + ';'
                    + sActRep + ';'
                    + sdFac + ';'
                    + sdNFac ;

        TobNewHeures.LoadFromSt ('ACT_DATEACTIVITE;ACT_UNITE;ACT_QTE;ACT_ACTIVITEEFFECT;ACT_ACTIVITEREPRIS;ACT_QTEFACTURE;ACT_QTENFACTURE', ';', StValeur, ';');
      end;
    end;
end;

(*
procedure TFActivite.DeleteLignesSuivantCriteres ;
  var
  sReqDel : string;
begin
  if gsReqPourDelete = '' then exit;

  sReqDel := 'DELETE FROM ACTIVITE ';
  sReqDel := sReqDel + gsReqPourDelete;

  try
  ExecuteSQL (sReqDel);

  except
    V_PGI.IoError := oeSaisie;
  end;

end;
*)
{==============================================================================================}
{===================================== Actions, boutons =======================================}
{==============================================================================================}
function TFActivite.ClickDel (ARow : integer; boucle : boolean) : boolean;
  var
  TobLigneEnGrid, TobLigneEnBase : TOB;
  iLigneDuHaut : integer;
  bAvecAlerte : boolean;
BEGIN
  TobLigneEnBase := nil;
  bAvecAlerte := false;
  Result := true;
  if Action = taConsult then Exit;

  if ((ARow < 1) or (ARow > TOBActivite.Detail.Count)) then Exit;

  TobLigneEnGrid := GetTOBLigne (ARow);
  if (TobLigneEnGrid = nil) then exit;

  if not boucle then
    begin
      bAvecAlerte := true;
      // message d'alerte pour prévenir de la suppression en base
      if (PGIAskAf (HTitres.Mess[31], Caption) <> mrYes) then exit;
    end;

  try
    //
    // Vérification de la stabilité de la ligne entre le moment où on a commencé à la modifier et le moment où on la valide
    //
    // On ramène la tob de la ligne telle qu'elle est en base pour vérifier que personne ne l'a modifiée entre temps
    // prévoit un verrou ici en cas de volonté future de gestion de blocage de la ligne modifiée
    TobLigneEnBase := LaTobDeLaLigneUnique (TobLigneEnGrid.GetValue ('ACT_TYPEACTIVITE'), TobLigneEnGrid.GetValue ('ACT_AFFAIRE'),
                                            TobLigneEnGrid.GetValue ('ACT_NUMLIGNEUNIQUE'));

    if (TobLigneEnBase <> nil) then
      begin
        // Si cette ligne existe deja, on vérifie qu'elle n'a pas été modifiée entre temps
        if (TobLigneEnGrid.GetValue ('ACT_DATEMODIF') <> TOBLigneEnBase.GetValue ('ACT_DATEMODIF'))
            or ((GetParamSoc ('SO_AFVISAACTIVITE') = True) and (TobLigneEnGrid.GetValue ('ACT_DATEVISA') <> TOBLigneEnBase.GetValue ('ACT_DATEVISA')))
            or ((GetParamSoc ('SO_AFAPPPOINT') = True) and (TobLigneEnGrid.GetValue ('ACT_DATEVISAFAC') <> TOBLigneEnBase.GetValue ('ACT_DATEVISAFAC')))
            then
          // Si la date de modification a changé, on alerte l'utilisateur qu'une modification est intervenue
          // et on rafraichit la saisie en perdant les modifications de la ligne en cours
          begin
            Result := false;
            PGIInfoAf ('Attention, la ligne courante a été modifiée par un autre utilisateur.' + Chr(13)
              + 'Votre sélection va être rafraîchie sans tenir compte de la suppression demandée.' + Chr(13)
              + 'Veuillez vous assurer que vous pouvez modifier cette sélection sans interaction avec d''autres utilisateurs.', Caption);

            RafraichirLaSelection;
            Exit;
          end;
      end;

    if (TobLigneEnGrid.GetValue ('ACT_RESSOURCE') <> '') then
      begin
        AFOAssistants.AddRessource (TobLigneEnGrid.GetValue ('ACT_RESSOURCE'));
      end;

    // On verifie qu'il n'y ait pas de blocage sur la ligne ou sur l'assistant
    // PL le 19/09/03 : seulement si la ligne existait en base auparavant
    if (TobLigneEnBase <> nil) then
      if BlocageLigne (TobLigneEnGrid, bAvecAlerte) then
        begin
          Result := false;
          exit;
        end;

    //
    // Si aucun blocage, on supprime la ligne en base
    //
    if (TobLigneEnBase <> nil) then
      if not TobLigneEnBase.DeleteDB then
        V_PGI.IoError := oeSaisie;

    if GetParamSoc('SO_AFVARIABLES') and (TobLigneEnBase <> nil) and (V_PGI.IoError = oeOk) then
      TraiteFormuleVar (ARow, 'SUPPRIME'); //ONYX-AFORMULEVAR

    if V_PGI.IoError = oeOk then
      begin
        // On met à jour les TOB et le Grid
        // On met à jour le nombre d'heures annuelle de la ressource courante
        MajNbHeureAnRessource (TobLigneEnGrid, true);
        TobLigneEnGrid.PutValue ('ACT_QTE', 0);
        // On met à jour le nombre d'heures
        MajTOBHeureCalendrier (ARow);
        MajTotauxMensuels (TobLigneEnBase, TobLigneEnGrid);

        if (CleAct.TypeAcces in [tacFrais, tacGlobal]) then
          gdSommeFrais := CalculSommeFrais (ACT_FOLIO.Text);


        GS.CacheEdit;
        GS.SynEnabled := False;
        GS.Cells [SG_Curseur, ARow] := '';
        iLigneDuHaut := GS.TopRow; // correction Bogue AGL
        GS.DeleteRow (ARow);
        GS.TopRow := iLigneDuHaut; // correction Bogue AGL
        if ((TOBActivite.Detail.Count > 1) and (ARow <> TOBActivite.Detail.Count)) then
          TOBActivite.Detail [ARow-1].Free
        else
          TOBActivite.Detail [ARow-1].InitValeurs;
      end;

  finally
    TobLigneEnBase.free;
  end;

  if (boucle = false) then
    begin
      TobLigneEnGrid := GetTOBLigne (ARow);
      if Not EstRempliGCActivite (GS, ARow) then
        InitialiseLigneAct (TobLigneEnGrid, ARow, CleAct, Action, Dev.Code);

      TOBLigneActivite.Dupliquer (TobLigneEnGrid, TRUE, TRUE, TRUE);
      ConversionTobActivite (TOBEnteteLigneActivite, 'ACT_DATEACTIVITE');

      if GS.RowCount < NbRowsInit then GS.RowCount := GS.RowCount + NbRowsPlus;
      GS.MontreEdit;
      GS.SynEnabled := True;
      //NumeroteLignes (GS, TOBActivite);

      GS.Col := SG_Jour;
      EntreDansLigneCourante (GS.Col, GS.Row);

      MajTotauxActivite;
    end;
END;

procedure TFActivite.bDelLigneClick (Sender : TObject);
begin
  ClickDel (GS.Row, false);
end;

procedure TFActivite.ClickInsert (ARow : integer; boucle : boolean );
BEGIN
  if Action = taConsult then Exit;
  if ((ARow < 1) or (ARow > TOBActivite.Detail.Count)) then Exit ;

  if (boucle = false) then
    begin
      // commencer par valider la ligne en cours de saisie
      GSExit (GS);
      // ensuite on peut insérer une ligne si tout est ok
      if (GCanClose = false) then exit;
    end;

  GS.CacheEdit;
  GS.SynEnabled := False;
  GS.Cells [SG_Curseur, ARow] := '';
  InsertTOBLigne (ARow);
  GS.InsertRow (ARow);
  TOBLigneActivite.Free;
  TOBLigneActivite := TOB.Create ('ACTIVITE', TOBEnteteLigneActivite, -1);

  // PL le 11/09/03 : articles liés : Ce champs boolean se doit d'être initialisé à vide
  TOBLigneActivite.AddChampSupValeur ('ARTSLIES', '', false); // Ce champs boolean se doit d'être initialisé à vide

  GS.Row := ARow;

  if (boucle = false) then
    begin
      GS.Cells [SG_Curseur, ARow] := '4';
      GS.MontreEdit;
      GS.SynEnabled := True;
//      NumeroteLignes (GS, TOBActivite);
      AfficheDetailTOBLigne (ARow);
    end;
END ;

procedure TFActivite.bNewligneClick (Sender : TObject);
begin
  ClickInsert (GS.Row, false);
end;

procedure TFActivite.bInsererLgnAffClick (Sender : TObject);
  Var
  TOBL, TOBC, TOBX, TOBArtCourant, TobLesArticles : TOB;
  ii : integer;
BEGIN
  if Action = taConsult then Exit;

  TOBX := nil;
  GS.ClearSelected;
  bSelectAll.Down := false;

  TOBL := GetTOBLigne (GS.Row);
  if TOBL = Nil then Exit;

  // On vérifie que le minimum d'info est saisi
  if (TOBL.GetValue ('ACT_AFFAIRE') = '') or (TOBL.GetValue ('ACT_DATEACTIVITE') = 0)
      or (TOBL.GetValue ('ACT_DATEACTIVITE') = idate1900) or (TOBL.GetValue ('ACT_DATEACTIVITE') = idate2099) then
    begin
      PGIInfoAF ('Le jour et l''affaire doivent au minimum être saisis pour utiliser cette fonctionnalité.', caption);
      exit;
    end;

  // On interdit de prendre une ligne déjà validée comme mère
  if (TOBL.GetValue ('ACT_NUMLIGNEUNIQUE') <> 0) then
    begin
      PGIInfoAF ('Cette fonctionnalité ne peut être lancée depuis une ligne validée.', caption);
      exit;
    end;

  // Recherche des articles de l'affaire
  TobLesArticles := TOBArticlesD1AffaireOnyx (TOBL.GetValue ('ACT_AFFAIRE'), '', '');

  try
  SourisSablier;
  //GS.InplaceEditor.Deselect;

  TOBX := TOB.Create ('Les activites crees', nil, -1);

  // Génération des lignes d'activité à partir des lignes d'affaire
  if (TobLesArticles <> nil) and (TobLesArticles.Detail.Count <> 0) then
    begin
      for ii := 0 to TobLesArticles.Detail.Count - 1 do
        begin
          TOBArtCourant := TobLesArticles.Detail[ii];

          // on ne reprend que les type d'article correspondant au type de saisie
          if (CleAct.TypeAcces <> tacGlobal) then
            if ((CleAct.TypeAcces = tacTemps) and (TOBArtCourant.GetValue('GA_TYPEARTICLE') <> 'PRE'))
                or ((CleAct.TypeAcces = tacFrais) and (TOBArtCourant.GetValue('GA_TYPEARTICLE') <> 'FRA'))
                or ((CleAct.TypeAcces = tacFourn) and (TOBArtCourant.GetValue('GA_TYPEARTICLE') <> 'MAR'))
              then continue;

          TOBC := TOB.Create ('', TOBX, -1);
          TOBC.Dupliquer (TOBL, True, True, False);

          // on copie l'article récupéré dans la ligne d'activité en forcant la mise à jour des prix
          CopieArtDansLaLigne (TOBArtCourant, TOBC, true);

          // on ne conserve pas le numéro unique : forcément à 0  puisque la ligne n'est pas validée
          //TOBC.PutValue ('ACT_NUMLIGNEUNIQUE', 0);
        end;
    end
  else
    begin
      PGIInfoAF ('Aucune ligne n''a été trouvée pour l''affaire en cours', caption);
      exit;
    end;

  // copie des lignes générées dans l'activité
//      for ii := 0 to TOBX.Detail.Count - 1 do
// PL le 14/10/03 : on insert dans le même ordre que les lignes ont été saisies
      for ii := TOBX.Detail.Count - 1 downto 0 do
        begin
          TOBC := TOBX.Detail[ii];
          if (ii <> TOBX.Detail.Count - 1) then
            ClickInsert (GS.Row, true);
          TOBL := GetTOBLigne (GS.Row);
          TOBL.Dupliquer (TOBC, True, True);
          // Création du champ d'indication de ligne validée ou non, car créée à la volée
          TOBL.AddChampSupValeur ('VALIDATION', '-', false);
          TOBL.SetAllModifie (true);

          // Afficher la ligne dans le grid
          AfficheLaTOBLigne (GS.Row);
        end;

//      if (GS.ColLengths[SG_Qte] <> -1) and (gbSaisieQtePermise = true) then
//        GS.Col := SG_Qte else GS.Col := SG_Jour;

  finally
    GS.MontreEdit;
    GS.SynEnabled := True;
    GS.InplaceEditor.Deselect;
    SourisNormale;
    TOBX.Free;
  end;
end;

procedure TFActivite.BMenuZoomMouseEnter (Sender : TObject);
begin
  PopZoom97 (BMenuZoom, POPZ);
end;

procedure TFActivite.BZoomTiersClick (Sender : TObject);
  var
  Tiers : string;
begin

  if Not BZoomTiers.Enabled then Exit;

  Tiers := GetChampLigne ('ACT_TIERS', GS.Row);
  if (Tiers <> '') then ZoomClient (Tiers);

end;

procedure TFActivite.BZoomAssistantClick (Sender : TObject);
  var
  Ressource : string;
begin

  if Not BZoomAssistant.Enabled then Exit;
  Ressource := GetChampLigne ('ACT_RESSOURCE', GS.Row);
  if (Ressource <> '') then ZoomRessource (Ressource);
end;

procedure TFActivite.BZoomCalendrierAssistantClick (Sender : TObject);
  var
  Index : integer;
begin
  if (TOBAssistant = nil) then exit;

  if (TOBAssistant.GetValue ('ARS_CALENSPECIF') = 'X') then
    BEGIN
      AglLanceFiche ('AFF', 'HORAIRESTD', '', '',
                    'TYPE:RES;CODE:' + TOBAssistant.GetValue ('ARS_RESSOURCE') + ';LIBELLE:' + TOBAssistant.GetValue ('ARS_LIBELLE') +
                    ';STANDARD:' + TOBAssistant.GetValue ('ARS_STANDCALEN'));
    END
  else
    BEGIN
      AglLanceFiche ('AFF', 'HORAIRESTD', '', '', 'TYPE:STD;STANDARD:' + TOBAssistant.GetValue ('ARS_STANDCALEN') + ';');
    END;

  // On complete remet à jour le calendrier de la ressource
  Index := AFOAssistants.AddRessource (TOBAssistant.GetValue ('ARS_RESSOURCE'));
  if (Index <> -1) and (Index <> -2) then
    begin
      TAFO_Ressource (AFOAssistants.GetRessource (Index)).ChargeCalendriers (true);
    end;
end;

procedure TFActivite.BValiderClick (Sender : TObject);
begin
  (*
  GCanClose:=true;
  // Gere la sortie du Grid pour validation
  GSExit(GS);
  if (GCanClose=true) then
    GereValider;
  *)
end;

procedure TFActivite.bRafraichitClick (Sender : TObject);
begin
  try
    gbModifMois := true;
    RafraichirLaSelection( GrCalendrier.DateSelect (GrCalendrier.FDateDeb, GCalendrier.Col, GCalendrier.Row) );
  finally
    gbModifMois := false;
  end;
end;

procedure TFActivite.RafraichirLaSelection (DateAAfficher : TDateTime = 0);
  var
  bBloqueSaisie : boolean;
begin
  if (DateAAfficher = 0) then
    // Affiche la derniere date de saisie pour l'affaire courante dans le cas où on n'a pas de date en entrée
    DateAAfficher := AfficheDerniereDateSaisieCritere (GrCalendrier.FDateDeb);

  AfficheSuivantCriteres (DateDebutCalendrier (DateAAfficher), true, true);


  // Test suivant la date courante si on peut entrer en saisie ou en consultation
  if (ControleDateDansIntervalle (DateAAfficher, false, false, true, bBloqueSaisie) <> 0) then
    begin
      if bBloqueSaisie then
        SetConsultation (True, tbaDatesAct);
    end
  else
    SetConsultation (False, tbaDatesAct);

  ChargerCalendrierEnCours( CleAct.Jour, CleAct.Mois, CleAct.Annee,  true);

  // Dans le cas de la saisie globale par ressource, certains jours où seuls des frais sont saisis n'ont pas de TOB associée
  // dans le calendrier, ce qui fausse les calculs. Il faut les initialiser en création.
  if (CleAct.TypeAcces = tacGlobal) and (CleAct.TypeSaisie = tsaRess) then
    CompleteLesFraisTOBNbHeureJour;

end;

procedure TFActivite.BImprimerClick (Sender : TObject);
  var
  TypeArticle, DateDeb, DateFin : string;
begin

  CriteresAvantAppel (TypeArticle, DateDeb, DateFin);

  if (CleAct.TypeSaisie = tsaRess) then
    AFLanceFiche_JournalActivite1Ress (ACT_RESSOURCE.Text, TypeArticle, DateDeb, DateFin)
  else
    AFLanceFiche_JournalClient1Tiers (ACT_TIERS.Text, ACT_AFFAIRE.Text, TypeArticle, DateDeb, DateFin);
end;

procedure TFActivite.CriteresAvantAppel (var sTypeArticle, sDateDeb, sDateFin : string);
  var
  YY, MM, DD, Semaine : word;
begin
  sTypeArticle := '';

  case CleAct.TypeAcces of
    tacTemps : sTypeArticle :=  'PRE';
    tacFrais : sTypeArticle :=  'FRA';
    tacFourn : sTypeArticle := 'MAR';
    tacGlobal : ;
  end;

  if (VH_GC.AFTYPESAISIEACT = 'SEM') then
    begin
      Semaine := strtoint (ACT_SEMAINE.Text);
      YY := strtoint (ACT_ANNEE.Text);
      if (strtoint (ACT_MOIS.Text) = 12) and (Semaine = 1) then YY := YY + 1;
      if (strtoint (ACT_MOIS.Text) = 1) and ((Semaine = 52) or (Semaine = 53)) then YY := YY - 1;
      sDateDeb := datetostr (PremierJourSemaine (Semaine, YY));

      sDateFin := datetostr(PremierJourSemaine (Semaine, YY) + 6);
    end
  else
    begin
      YY := strtoint (ACT_ANNEE.Text);
      MM := strtoint (ACT_MOIS.Text);
      DD := 1;
      sDateDeb := DateToStr (EncodeDate (YY, MM, DD));
      if (MM = 12) then
        begin
          YY := YY + 1;
          MM := 1;
        end
      else
        Inc (MM);

      sDateFin := DateToStr (EncodeDate (YY, MM, 1) - 1);
    end;

end;


procedure TFActivite.BChercherClick (Sender : TObject);
begin
  if TobActivite.Detail.Count <= 2 then Exit;
  GS.SetFocus;
  FindFirst := True;
  FindLigne.Execute;
end;

Procedure TFActivite.ZoomOuChoixRes ( ACol,ARow : integer ) ;
Var Ressource : String ;
BEGIN
if ACol<>SG_Ressource then Exit ;

//Ressource:=GetChampLigne('ACT_RESSOURCE',ARow);
Ressource := GS.Cells[ACol,ARow];
if Ressource<>'' then
   BEGIN
   ZoomRessource(Ressource) ;
   END
else
   BEGIN
   IdentifieAssistant(GS);
   END ;

END ;

Procedure TFActivite.ZoomOuChoixTiers ( ACol,ARow : integer ) ;
Var Tiers : String ;
bRechParCode : boolean;
BEGIN
if ACol<>SG_Tiers then Exit;

Tiers := GS.Cells[ACol, ARow];
if Tiers <> '' then
   BEGIN
   ZoomClient (Tiers);
   END
else
   BEGIN
//   IdentifieTiers (GS);
    bRechParCode := true;
    if (CleAct.TypeSaisie = tsaRess) and (GetParamSoc ('SO_AFRECHTIERSPARNOM') = true) then
      bRechParCode := false;

    if GetTiersRechercheAF (GS, ACT_TIERS.Hint, bRechParCode, Tiers) then
      ;
   END ;

END ;

Procedure TFActivite.ZoomOuChoixArt ( ACol,ARow : integer ) ;
Var RefUnique, TypeArticle : String ;
    Cancel, OkArt : boolean ;
    TOBA, TOBL:TOB;
BEGIN
if ACol<>SG_Article then Exit ;

//RefUnique:=GetChampLigne('ACT_ARTICLE',ARow);
RefUnique:=GS.Cells[ACol,ARow];
TypeArticle:=GetChampLigne('ACT_TYPEARTICLE',ARow);
if RefUnique<>'' then
   BEGIN
   ZoomArticle(RefUnique, TypeArticle) ;
   END
else
   BEGIN
   Cancel:=False ; OkArt:=IdentifierArticle(GS, ACol,ARow,Cancel,True) ;
    if ((OkArt) and (Not Cancel)) then
       BEGIN
       TOBValo.Free; TOBValo:=nil;
       TOBL:=GetTOBLigne(GS.Row);
       if (TOBL=nil) then exit;

       StCellCur:=GS.Cells[ACol,ARow] ;
       TOBA := FindTOBArtSaisAff(TOBArticles, TOBL.GetValue('ACT_ARTICLE'), true );
       if (TOBA<>nil) then
         ACT_ARTICLEL.Caption := TOBA.GetValue('GA_LIBELLE')
       else
         ACT_ARTICLEL.Caption := '';
       END ;
   END ;

END ;

procedure TFActivite.bSelectAllClick(Sender: TObject);
var
i:integer;
begin

for i:=1 to GS.RowCount-1 do
    if (bSelectAll.Down) then
        begin
        if Not GS.isSelected(i) then
            if EstRempliGCActivite(GS, i, false) then
                GS.FlipSelection(i);
        end
    else
        begin
        if GS.isSelected(i) then
            GS.FlipSelection(i);
        end;
end;

procedure TFActivite.bSupplementClick(Sender: TObject);
begin
//
HPBas.Visible:=bSupplement.Down;

CpltLigne.Checked := HPBas.visible;
end;

procedure TFActivite.bDescriptifClick(Sender: TObject);
Begin
TDescriptif.Visible:=BDescriptif.Down;

Afficherlemmo1.Checked :=TDescriptif.Visible;
end;

procedure TFActivite.Supprimeuneligne1Click(Sender: TObject);
begin
ClickDel(GS.Row, false) ;
end;

procedure TFActivite.Ajoutduneligne1Click(Sender: TObject);
begin
ClickInsert(GS.Row, false) ;
end;

procedure TFActivite.Slectionneruneligne1Click(Sender: TObject);
begin
GS.FlipSelection(GS.Row) ;
end;

procedure TFActivite.Afficherlemmo1Click(Sender: TObject);
begin
if (bDescriptif.Down=true) then
   bDescriptif.Down:=false
else
   bDescriptif.Down:=true;

TDescriptif.Visible:=BDescriptif.Down;
Afficherlemmo1.Checked := bDescriptif.Down;
end;

procedure TFActivite.CpltLigneClick(Sender: TObject);
Begin
if (bSupplement.Down=true) then
   bSupplement.Down:=false
else
   bSupplement.Down:=true;

HPBas.visible := bSupplement.Down;
TMenuItem(Sender).Checked := bSupplement.Down;
end;

procedure TFActivite.HValComboBoxTriChange(Sender: TObject);
var
bTriOK:boolean;
begin
bTriOK:=true;
if (TOBActivite.Detail.Count=0) then exit;
GbTri := true;
try
// Cas des tris non traités
if (CleAct.TypeSaisie = tsaRess) then
    begin
    if (HValComboBoxTri.Values[HValComboBoxTri.ItemIndex]='RES') then
        bTriOK:=false
    else
    if (HValComboBoxTri.Values[HValComboBoxTri.ItemIndex]='TYP') then
        if (CleAct.TypeAcces <> tacGlobal) then
            bTriOK:=false;

    end
else
    begin
    if (HValComboBoxTri.Values[HValComboBoxTri.ItemIndex]='CLI') OR (HValComboBoxTri.Values[HValComboBoxTri.ItemIndex]='MIS') then
        bTriOK:=false
    else
    if (HValComboBoxTri.Values[HValComboBoxTri.ItemIndex]='TYP') then
        if (CleAct.TypeAcces <> tacGlobal) then
            bTriOK:=false;

    end;

if not bTriOK then
    begin
    PGIBoxAf(HTitres.Mess[27], Caption);
    HValComboBoxTri.ItemIndex := OldIndexTri;
    exit;
    end;

gbSelectCellCalend:=false;
GS.Row:=1;
if (OldIndexTri = HValComboBoxTri.ItemIndex) then
   begin
   if (OrdreTri = '') then OrdreTri := 'DESC' else OrdreTri := '';
   end
else
   OrdreTri := '';

  // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
GSExit (GS);
if (GCanClose = false) then Abort;

OldIndexTri := HValComboBoxTri.ItemIndex;

// Affiche les lignes correspondant à la ressources affichée
AfficheSuivantCriteres(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GCalendrier.Col,GCalendrier.Row), False, False);

finally
  GbTri := false;
end
end;



procedure TFActivite.BRechAffaireClick(Sender: TObject);
  var
  bProposition, bChangeStatut : boolean;
  sAffaire : string;
begin
  ACT_AFFAIRE.onchange := nil;
  try
    if VH_GC.AFProposAct then bChangeStatut := true else bChangeStatut := False;
    if (ACT_AFFAIRE0.Text = 'P') then bProposition := True else bProposition := False;
    if GetAffaireEntete (ACT_AFFAIRE, ACT_AFFAIRE1, ACT_AFFAIRE2, ACT_AFFAIRE3, ACT_AVENANT, ACT_TIERS,
              bProposition, bChangeStatut, false, Not gbCritereFixe,True, 'SAT', true, true, true) then
        ChargeTiers (ACT_TIERS, ACT_TIERS.Text);

  finally
    sAffaire := ACT_AFFAIRE.Text;
    // on force le onchange
    ACT_AFFAIRE.Text := '£';
    CleActOld.Affaire := '£';
    ACT_AFFAIRE.onchange := ACT_AFFAIREChange;
  end;

  CleAct.Tiers := ACT_TIERS.Text;
  ACT_AFFAIRE.Text := sAffaire;
  CleAct.Affaire := sAffaire;
  CleAct.Tiers := ACT_TIERS.Text;
  {$IFDEF BTP}
  BTPCodeAffaireDecoupe (CleAct.Affaire, CleAct.Aff0, CleAct.Aff1, CleAct.Aff2, CleAct.Aff3, CleAct.Avenant, Action, false);
  {$ELSE}
  CodeAffaireDecoupe (CleAct.Affaire, CleAct.Aff0, CleAct.Aff1, CleAct.Aff2, CleAct.Aff3, CleAct.Avenant, Action, false);
  {$ENDIF}

  ACT_AFFAIRE0.onchange := nil;
  try
    ACT_AFFAIRE0.Text := CleAct.Aff0;
  finally
    ACT_AFFAIRE0.onchange := ACT_AFFAIRE0Change;
  end;

  ACT_AFFAIRE1.onchange := nil;
  try
    ACT_AFFAIRE1.Text := CleAct.Aff1;
  finally
    ACT_AFFAIRE1.onchange := ACT_AFFAIRE1Change;
  end;

  ACT_AFFAIRE2.onchange := nil;
  try
    ACT_AFFAIRE2.Text := CleAct.Aff2;
  finally
    ACT_AFFAIRE2.onchange := ACT_AFFAIRE2Change;
  end;

  ACT_AFFAIRE3.Text := CleAct.Aff3;
  ACT_AVENANT.Text := CleAct.Avenant;
end;

procedure TFActivite.BZoomArticleClick(Sender: TObject);
Var RefUnique, sTypeArticle :string;
begin
if Not BZoomArticle.Enabled then Exit ;
RefUnique:= GetChampLigne('ACT_ARTICLE',GS.Row) ;
sTypeArticle:= GetChampLigne('ACT_TYPEARTICLE',GS.Row) ;
if (RefUnique <>'') then ZoomArticle(RefUnique, sTypeArticle) ;
end;

procedure TFActivite.BZoomAffaireClick(Sender: TObject);
begin
ZoomAffaire(ACT_AFFAIRE.Text);
end;

procedure TFActivite.BZoomTableauBordClick(Sender: TObject);
begin
//OuvreFiche('AFF','AFTBVIEWER','','','TYPETB:AFF;AFFAIRE:'+GetChamp('AFF_AFFAIRE')+';TIERS:'+GetChamp('AFF_TIERS'));
AglLanceFiche ('AFF','AFTBVIEWER','','','TYPETB:AFF;AFFAIRE:'+ACT_AFFAIRE.Text+';TIERS:'+ACT_TIERS.Text);
end;

procedure TFActivite.bEffaceAffaireTiersClick(Sender: TObject);
begin
ACT_TIERS.OnChange := nil;
ACT_AFFAIRE0.OnChange := nil;
ACT_AFFAIRE1.OnChange := nil;
ACT_AFFAIRE2.OnChange := nil;
try
// PL le 25/04/03 : en critères fixes, on ne peux pas effacer le tiers passé en paramètre d'entrée
if Not gbCritereFixe then ACT_TIERS.Text:='';

ACT_AFFAIRE0.Text:='';
ACT_AFFAIRE1.Text:='';
ACT_AFFAIRE2.Text:='';
ACT_AFFAIRE3.Text:='';
ACT_AVENANT.Text:='';
finally
ACT_TIERS.OnChange := ACT_TIERSChange;
ACT_AFFAIRE0.OnChange := ACT_AFFAIRE0Change;
ACT_AFFAIRE1.OnChange := ACT_AFFAIRE1Change;
ACT_AFFAIRE2.OnChange := ACT_AFFAIRE2Change;
end;

ACT_AFFAIRE.Text:='';
GereAffaireEnabled;
LibelleAffaire.Caption:='';

// PL le 25/04/03 : en critères fixes, on ne peux pas effacer le tiers passé en paramètre d'entrée
if Not gbCritereFixe then
  begin
    LibelleTiers.Caption:='';
    MOIS_CLOT_CLIENT.Text := '';
  end;

CleAct.Affaire:='';
if Not(VH_GC.AFProposAct) then
   CleAct.Aff0       := 'A'
else
   CleAct.Aff0       := '' ;
CleAct.Aff1:='';
CleAct.Aff2:='';
CleAct.Aff3:='';
CleAct.Avenant:='';
CleActOld.Affaire:='';
if Not(VH_GC.AFProposAct) then
   CleActOld.Aff0       := 'A'
else
   CleActOld.Aff0       := '' ;
CleActOld.Aff1:='';
CleActOld.Aff2:='';
CleActOld.Aff3:='';
CleActOld.Avenant:='';
if (ACT_AFFAIRE0.CanFocus) then ACT_AFFAIRE0.SetFocus
else if (ACT_AFFAIRE1.CanFocus) then ACT_AFFAIRE1.SetFocus;
end;

procedure TFActivite.Inverserslection1Click(Sender: TObject);
  var
  i : integer;
begin
  for i := 1 to GS.RowCount - 1 do
    GS.FlipSelection (i);
end;

procedure TFActivite.Copier1Click (Sender : TObject);
  var
  OkG : boolean;
begin
  OkG := (Screen.ActiveControl = GS);

  ClipBoard.Clear;

  if ((OkG) and (GS.NbSelected > 0)) then
    CopierLignes;
end;

procedure TFActivite.Couper1Click (Sender : TObject);
  var
  OkG : boolean;
begin
  OkG := (Screen.ActiveControl = GS);

  GS.InplaceEditor.Deselect;
  if ((OkG) and (GS.NbSelected > 0)) then
    BEGIN
      GS.Cells [SG_Curseur, GS.Row] := '';
      CouperLignes;
    END ;
end;

procedure TFActivite.Coller1Click (Sender : TObject);
  var
  OkG : boolean;
begin
  OkG := (Screen.ActiveControl = GS);
  if ((OkG) and (TOBCXV.Detail.Count > 0)) then
    CollerLignes;
end;

procedure TFActivite.POPYPopup(Sender: TObject);
var
OK:boolean;
begin
if (GS.NbSelected <> 0) then
  begin
    Copier1.Enabled:=true;
    Copier1.Tag:=0;
    Couper1.Enabled:=true;
    Couper1.Tag:=0;
  end
else
  begin
    Copier1.Enabled := false;
    Copier1.Tag := -1;
    Couper1.Enabled := false;
    Couper1.Tag := -1;
  end;

if (TOBCXV=nil) or (TOBCXV.Detail.Count=0) then
    begin
    Coller1.Enabled:=false;
    Coller1.Tag:=-1;
    end
else
    begin
    OK:=true;
    if (TOBAssistant<>nil) then
    if (TOBAssistant.GetValue('ARS_FERME')='X') then OK := false;

    if OK then
        begin
        Coller1.Enabled:=true;
        Coller1.Tag:=0;
        end;
    end;
end;

{===========================================================================================}
{===================================== Gestion de la date ==================================}
{===========================================================================================}
procedure TFActivite.ACT_SEMAINEChange(Sender: TObject);
  var
  iOldAnnee, iOldMois, iSemaine : integer;
  iNewAnnee, iNewMois, iNewJour : word;
  dJourAAfficher : TDateTime;
  eOnChange : TNotifyEvent;
  bBloqueSaisie : boolean;
begin
  try
    // Saisie de la semaine vide
    if (ACT_ANNEE.Text = '') or (ACT_SEMAINE.Text = '') then exit;
    // Pas de saisie de la semaine
    if (ACT_SEMAINE.Text = inttostr (CleAct.Semaine)) then exit;

    // Format de la semaine correct
    iOldMois := strtoint (ACT_MOIS.Text);
    iOldAnnee := strtoint (ACT_ANNEE.Text);
    iSemaine := strtoint (ACT_SEMAINE.Text);

    if (iSemaine < 1) then
    // PL AGL545 : le 08/10/02
    {$IFDEF AGLINF545}
       if V_PGI.Semaine53 then begin ACT_SEMAINE.Text := '2'; iSemaine := 2; end
       else
    {$ENDIF}
        begin
        ACT_SEMAINE.Text := '1';
        iSemaine := 1;
        end;

    if (iSemaine > 53) then
    // PL AGL545 : le 08/10/02
    {$IFDEF AGLINF545}
       if V_PGI.Semaine53 then begin ACT_SEMAINE.Text := '53'; iSemaine := 53; end
       else begin ACT_SEMAINE.Text := '52'; iSemaine := 52; end;
    {$ELSE}
        begin
        // la dernière semaine de l'année courante
        iSemaine := NumSemaine (EncodeDate (iOldAnnee, 12, 31));
        ACT_SEMAINE.Text := inttostr (iSemaine);
        end;
    {$ENDIF}

      // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
    GSExit (GS);
    if (GCanClose = false) then
    // Gere la validation des lignes d'activité saisies
    //if (GereValider<>oeOk) then
       begin
       ACT_SEMAINE.Text := inttostr (CleAct.Semaine);
       ChargerCalendrierEnCours (CleAct.Jour, iOldMois, iOldAnnee, false);
       Abort;
       end;

    // On cherche le jour à afficher suivant les contraintes
    dJourAAfficher := ChercheJourAAfficherSemaine (CleAct.Semaine, strtoint (ACT_SEMAINE.Text), iOldMois);

    if ACT_SEMAINE.Focused or gcalendrier.focused then
       if (ControleDateDansIntervalle (dJourAAfficher, true, false, (CleAct.TypeSaisie = tsaClient), bBloqueSaisie) <> 0) then
            begin
            if bBloqueSaisie then
                SetConsultation (True, tbaDatesAct);
            end
       else
            SetConsultation (False, tbaDatesAct);

    //If Not ControleDateDansIntervalle(dJourAAfficher) then
      // begin
       // PL le 19/07/2000 on ne bloque plus si la date n'est pas dans l'intervalle de la mission
       //ACT_SEMAINE.Text := inttostr(CleAct.Semaine);
       //ChargerCalendrierEnCours( CleAct.Jour, iOldMois, iOldAnnee, false);
       //Abort;
      // end;

    DecodeDate (dJourAAfficher, iNewAnnee, iNewMois, iNewJour);

    // On modifie le mois si le changement de semaine a occasionné un changement de mois
    if (iNewMois <> iOldMois) then
       begin
       CleAct.Mois := strtoint (ACT_MOIS.Text);
       eOnChange := ACT_MOISChange;
       ACT_MOIS.OnChange := nil;
       ACT_MOIS.Text := inttostr (iNewMois);
       CleAct.Mois := iNewMois;
       ACT_MOIS.OnChange := eOnChange;
       // Si on change le mois, il faudra forcer le rafraichissement des données mensuelles sur les compléments
       gbModifMois := true;
       end;

    CleAct.Semaine := strtoint (ACT_SEMAINE.Text);
    CleAct.Jour := iNewJour;

    AfficheSuivantCriteres (GrCalendrier.PremierJourAffiche (dJourAAfficher), true, true);

    // Dans le cas de la saisie globale par ressource, certains jours où seuls des frais sont saisis n'ont pas de TOB associée
    // dans le calendrier, ce qui fausse les calculs. Il faut les initialiser en création.
    if (CleAct.TypeAcces = tacGlobal) and (CleAct.TypeSaisie = tsaRess) then
      CompleteLesFraisTOBNbHeureJour;

    // Mise à jour du libellé indiquant l'intervalle des jours pour la semaine en cours
    //MajIntervalleSemaine;

  finally
    // Si le mois a changé, on reinitialise à false en sortant
    gbModifMois := false;
  end;
end;

function TFActivite.ChercheJourAAfficherSemaine( AciOldSemaine, AciNewSemaine, AciOldMois : integer ):TDateTime;
var
PremierJour:TDateTime;
iJourSemaineCalend, Inc, Annee:integer;
FAnnee,FMois,FJour:word;
begin
// cherche dans le calendrier le même jour de semaine que le jour courant
// sauf en changement d'année : premier jour dans l'annee de la semaine selectionnee
iJourSemaineCalend := DayOfWeek(GrCalendrier.DateSelect(GrCalendrier.FDateDeb,GCalendrier.Col,GCalendrier.Row))-2;

{$IFDEF AGLINF545}
PremierJour := PremierJourSemaineTempo(AciNewSemaine, CleAct.Annee);
{$ELSE}
// PL le 27/05/02  AGL550
Annee := CleAct.Annee;
if (AciOldMois=12) and (AciNewSemaine=1) then Annee := Annee + 1;
if (AciOldMois=1) and ((AciNewSemaine=52) or (AciNewSemaine=53)) then Annee := Annee - 1;
PremierJour := PremierJourSemaine(AciNewSemaine, Annee);
/////////////////////// AGL550
{$ENDIF}

if (iJourSemaineCalend=-1) then iJourSemaineCalend:=6;
PremierJour := PremierJour + iJourSemaineCalend;

DecodeDate(PremierJour,FAnnee,FMois,FJour);

Inc :=0;
if (FAnnee > CleAct.Annee) then Inc := -1
else if (FAnnee < CleAct.Annee) then Inc := 1;

while (FAnnee <> CleAct.Annee) do
   begin
   PremierJour := PremierJour + Inc;
   DecodeDate(PremierJour,FAnnee,FMois,FJour);
   end;

Result:=PremierJour;
end;

procedure TFActivite.ACT_FOLIOChange(Sender: TObject);
begin
  try
    gbModifFolio := true;
    gbSelectCellCalend := false;
    // Saisie de la semaine vide
    if (ACT_FOLIO.Text = '') then exit;
    // Pas de saisie de la semaine
    if (ACT_FOLIO.Text = inttostr (CleAct.Folio)) then exit;

    // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
    GSExit (GS);

    if (GCanClose = false) then
    // Gere la validation des lignes d'activité saisies
    //if (GereValider<>oeOk) then
      begin
        ACT_FOLIO.Text := inttostr (CleAct.Folio);
        Abort;
      end;

    CleAct.Folio := strtoint (ACT_FOLIO.Text);

    AfficheSuivantCriteres (GrCalendrier.FDateDeb, true, true);

    // Dans le cas de la saisie globale par ressource, certains jours où seuls des frais sont saisis n'ont pas de TOB associée
    // dans le calendrier, ce qui fausse les calculs. Il faut les initialiser en création.
    if (CleAct.TypeAcces = tacGlobal) and (CleAct.TypeSaisie = tsaRess) then
      CompleteLesFraisTOBNbHeureJour;

  finally
    gbModifFolio := false;
  end;

end;

procedure TFActivite.ACT_MOISChange(Sender: TObject);
  var
  iMois : integer;
  iAnnee, Rep : integer;
  eOnChange : TNotifyEvent;
  bBloqueSaisie : boolean;
  dDernierJourMoisPrecedent : TDateTime;
begin
  try
    gbModifMois := true;
    gbSelectCellCalend := false;
    // Saisie du mois vide
    if (ACT_MOIS.Text = '') then exit;
    // Pas de saisie du mois
    if (ACT_MOIS.Text = inttostr (CleAct.Mois)) then exit;
    // Format du mois correct
    iMois := strtoint (ACT_MOIS.Text);
    if (iMois < 0) then
       begin
       ACT_MOIS.Text := '1';
       exit;
       end;
    if (iMois > 13) then
       begin
       ACT_MOIS.Text := '12';
       exit;
       end;
    if (iMois = 13) then
       begin
       eOnChange := ACT_ANNEEChange;
       ACT_ANNEE.OnChange := nil;
       iAnnee := strtoint (ACT_ANNEE.Text);
       ACT_ANNEE.Text := inttostr (iAnnee + 1);
       ACT_ANNEE.OnChange := eOnChange;
       ACT_MOIS.Text := '1';
       exit;
       end;
    if (iMois = 0) then
       begin
       eOnChange := ACT_ANNEEChange;
       ACT_ANNEE.OnChange := nil;
       iAnnee := strtoint (ACT_ANNEE.Text);
       ACT_ANNEE.Text := inttostr (iAnnee - 1);
       ACT_ANNEE.OnChange := eOnChange;
       ACT_MOIS.Text := '12';
       exit;
       end;

    // On affiche le premier jour du mois
    // PL le 22/10/02 : bogue de non passage au mois suivant en saisie par semaine trouvé par MCF
    CleAct.Jour:=1;
    if (VH_GC.AFTYPESAISIEACT = 'SEM') then
        begin
        // cas du passage au mois suivant
        if ((strtoint(ACT_MOIS.Text) > CleAct.Mois) or ((CleAct.Mois = 12) and (strtoint (ACT_MOIS.Text) = 1))) then
          begin
          dDernierJourMoisPrecedent := EncodeDate (strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), 1) - 1;
          // En saisie par semaine, si on veut forcer le changement de mois suivant, il faut aller chercher le premier jour
          // de ce mois qui n'est pas à cheval sur la dernière semaine du mois précédent. Autrement, on va se repositionner
          // sur le premier jour contenant de l'activité de cette semaine et ce sera peut être dans le mois d'origine, et pas
          // dans le suivant => bouclage et on reste toujours sur le même mois.
          while (NumSemaine (EncodeDate (strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), CleAct.Jour)) = NumSemaine (dDernierJourMoisPrecedent)) do
            Inc (CleAct.Jour);
          end
        else
        if ((strtoint (ACT_MOIS.Text) < CleAct.Mois) or ((CleAct.Mois = 1) and (strtoint (ACT_MOIS.Text) = 12))) then
          begin
          dDernierJourMoisPrecedent := EncodeDate (strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), 1) - 1;
          // En saisie par semaine, si on veut forcer le passage au mois précédent, il faut aller chercher le premier jour
          // de ce mois qui n'est pas à cheval sur la dernière semaine du mois précédent. Autrement, on va se repositionner
          // sur le premier jour contenant de l'activité de cette semaine et ce sera peut être dans le mois d'origine, et pas
          // dans le précédent => saut d'un mois précédent supplémentaire.
          while (NumSemaine (EncodeDate (strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), CleAct.Jour)) = NumSemaine (dDernierJourMoisPrecedent)) do
            Inc (CleAct.Jour);
          end;
        end;
    /////////////////////////////


    // controle la validite de la date à afficher
    if (Not ACT_ANNEE.Focused and Not ACT_SEMAINE.Focused and Not GS.Focused and Not GCalendrier.Focused) then
    //   ControleDateDansIntervalle(EncodeDate(strtoint(ACT_ANNEE.Text), strtoint(ACT_MOIS.Text), CleAct.Jour), true, false, (CleAct.TypeSaisie = tsaClient));
        begin
        Rep := ControleDateDansIntervalle (EncodeDate (strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), CleAct.Jour), true, false, (CleAct.TypeSaisie = tsaClient), bBloqueSaisie);
        if (Rep <> 0) then
            begin
            if bBloqueSaisie then
                SetConsultation (True, tbaDatesAct);
            end
        else
            SetConsultation (False, tbaDatesAct);
        end;

      // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
    GSExit (GS);
    if (GCanClose = false) then
    // Gere la validation des lignes d'activité saisies
    //if (GereValider<>oeOk) then
       begin
       eOnChange := ACT_ANNEEChange;
       ACT_ANNEE.OnChange := nil;
       ACT_ANNEE.Text := inttostr (CleAct.Annee);
       ACT_ANNEE.OnChange := eOnChange;
       ACT_MOIS.Text := inttostr (CleAct.Mois);
       Abort;
       end;

    CleAct.Annee := strtoint (ACT_ANNEE.Text);
    CleAct.Mois := strtoint (ACT_MOIS.Text);

    MajSemaine;

  finally
    gbModifMois := false;
  end;
end;

procedure TFActivite.ACT_ANNEEChange(Sender: TObject);
  var
  iAnnee : integer;
  FJour, Fmois, FYear : Word ;
  bBloqueSaisie : boolean;
begin
  try
    gbModifAnnee := true;
    gbSelectCellCalend := false;
    // Saisie de l'année vide
    if (ACT_ANNEE.Text = '') then exit;
    // Pas de saisie de l'année
    if (ACT_ANNEE.Text = inttostr (CleAct.Annee)) then exit;

    // Format de l'année correct
    iAnnee := strtoint (ACT_ANNEE.Text);
    if (iAnnee < 1900) or (iAnnee > 2100) then
       begin
       DecodeDate (date, FYear, FMois, FJour);
       ACT_ANNEE.Text := inttostr (FYear);
       end;

    // controle la validite de la date à afficher
    if ACT_ANNEE.Focused then
    //   ControleDateDansIntervalle(EncodeDate(strtoint(ACT_ANNEE.Text), strtoint(ACT_MOIS.Text), CleAct.Jour), true, false, (CleAct.TypeSaisie = tsaClient));
       if (ControleDateDansIntervalle (EncodeDate(strtoint (ACT_ANNEE.Text), strtoint (ACT_MOIS.Text), CleAct.Jour), true, false, (CleAct.TypeSaisie = tsaClient), bBloqueSaisie) <> 0) then
            begin
            if bBloqueSaisie then
                SetConsultation (True, tbaDatesAct);
            end
       else
            SetConsultation (False, tbaDatesAct);

      // Gere la sortie du Grid de saisie d'activite (qui valide la ligne en cours par la même occasion)
    GSExit (GS);

    if (GCanClose = false) then
    // Gere la validation des lignes d'activité saisies
    //if (GereValider<>oeOk)  then
       begin
       ACT_ANNEE.Text := inttostr (CleAct.Annee);
       Abort;
       end;

    CleAct.Annee := strtoint (ACT_ANNEE.Text);

    MajSemaine;

  finally
    gbModifAnnee := false;
  end;
end;

procedure TFActivite.MajSemaine;
var
iAnnee:integer;
iMois:integer;
iSemaine:integer;
eOnChange:TNotifyEvent;
begin
iAnnee := strtoint(ACT_ANNEE.Text);
iMois := strtoint(ACT_MOIS.Text);

eOnChange := ACT_SEMAINEChange;
ACT_SEMAINE.OnChange := nil;
// PL le 22/10/02 : bogue de non passage au mois suivant en saisie par semaine trouvé par MCF
//iSemaine := NumSemaine(EncodeDate(iAnnee,iMois,1));
iSemaine := NumSemaine(EncodeDate(iAnnee,iMois,Cleact.Jour));
//////////////////////////////

// PL le 23/05/02 : suite reparation de la fonction agl (V5.4.0i) NumSemaine
{$IFDEF AGLINF545}
if (iSemaine > 52) then
    begin
// PL AGL550 : le 08/10/02
    iSemaine := 53;
    if (Not V_PGI.Semaine53) then iSemaine := 1;
    end;
{$ENDIF}
//////////////////////////////
CleAct.Semaine := iSemaine;
ACT_SEMAINE.Text := inttostr(iSemaine);
ACT_SEMAINE.OnChange := eOnChange;

AfficheSuivantCriteres(GrCalendrier.PremierJourAffiche(EncodeDate(iAnnee, iMois, CleAct.Jour)), true, true);

// Dans le cas de la saisie globale par ressource, certains jours où seuls des frais sont saisis n'ont pas de TOB associée
// dans le calendrier, ce qui fausse les calculs. Il faut les initialiser en création.
if (CleAct.TypeAcces = tacGlobal) and (CleAct.TypeSaisie = tsaRess) then
  CompleteLesFraisTOBNbHeureJour;

end;


Procedure TFActivite.MajIntervalleSemaine;
var
dDateDebut, dDateFin : TDateTime;
YY,Semaine:integer;
begin
if (VH_GC.AFTYPESAISIEACT <> 'SEM') then exit;

if ((ACT_SEMAINE.Text='') or (ACT_ANNEE.Text='')) then LblIntervalle.Caption := ''
else
   begin
{$IFDEF AGLINF545}
   dDateDebut := PremierJourSemaineTempo(strtoint(ACT_SEMAINE.Text), strtoint(ACT_ANNEE.Text));
{$ELSE}
// PL le 23/05/02 : réparation de la fonction AGL550
   Semaine:=strtoint(ACT_SEMAINE.Text); YY:=strtoint(ACT_ANNEE.Text);
   if (strtoint(ACT_MOIS.Text)=12) and (Semaine=1) then YY := YY + 1;
   if (strtoint(ACT_MOIS.Text)=1) and ((Semaine=52) or (Semaine=53)) then YY := YY - 1;
   dDateDebut := PremierJourSemaine(Semaine,YY);
/////////////////////////// AGL550
{$ENDIF}

// PL le 27/05/02 : Nouvelles fonctions AGL NumSemaine et PremierJourSemaine
   dDateFin := dDateDebut;
//   if V_PGI.Semaine53 and (strtoint(ACT_SEMAINE.Text)=53) then
//        dDateFin := EncodeDate(strtoint(ACT_ANNEE.Text)+1, 1, 1)-1
//   else
        while DayOfWeek(dDateFin)<>1 do dDateFin := dDateFin + 1 ;
/////////////////////////

   LblIntervalle.Caption := 'du ' + datetostr(dDateDebut);
   LblIntervalle2.Caption := 'au ' + datetostr(dDateFin);
   end;
end;


// Recherche de la dernière date d'activité saisie pour une ressource ou une affaire
function TFActivite.ChargerDerniereDateSaisieCritere  (AcbInit : boolean) : boolean;
  var
  QQ : TQuery;
  sReq, sTypeArt : string;
  dDateMax : TDateTime;
  iJour, iMois, iAnnee : word;
begin
  QQ := nil;
  Result := false;
  Try

  if (CleAct.TypeSaisie = tsaRess) then
    begin
      (*   sReq := 'SELECT ACT_DATEACTIVITE FROM ACTIVITE WHERE ACT_TYPEACTIVITE="'+CleAct.TypeActivite
                             +'" AND ACT_RESSOURCE="'+CleAct.Ressource;*)
      sReq := 'SELECT MAX(ACT_DATEACTIVITE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + CleAct.TypeActivite
              + '" AND ACT_RESSOURCE="' + CleAct.Ressource;
    end
  else
    begin
      // PL le 06/03/02 : INDEX 4
      sReq := 'SELECT MAX(ACT_DATEACTIVITE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="' + CleAct.TypeActivite;
      if (CleAct.Tiers <> '') then sReq := sReq + '" AND ACT_TIERS="' + CleAct.Tiers;
      if (CleAct.Affaire <> '') then sReq := sReq + '" AND ACT_AFFAIRE="' + CleAct.Affaire;
    end;

  // PL le 10/10/03 : un client saisissait des folios 6 sans avoir saisi les 1...5 avant,
  // en initialisation, le max(date) where folio=1 n'est plus bon dans ce cas
  // un future modification doit être faite engestion par folio
  // select Max(ACT_DATEACTIVITE) as DATEMAX, ACT_FOLIO from activite where act_ressource="5" group by act_folio order by DATEMAX desc, act_folio desc
  // ramène en premier la date la plus grande et le folio le plus grand pour cette date. Il faut ramener le folio en paramètres pour positionner la sélection sur celui-ci
  if Not AcbInit then
    if (VH_GC.AFTYPESAISIEACT = 'FOL')  then
      sReq := sReq + '" AND ACT_FOLIO="' + inttostr (CleAct.Folio);

  // configuration de la selection suivant le type d'acces à la saisie, par temps, frais, fournitures, global
  case CleAct.TypeAcces of
     tacTemps : begin
                  sTypeArt := 'PRE';
                end;
     tacFrais : begin
                  sTypeArt := 'FRA';
                end;
     tacFourn : begin
                  sTypeArt := 'MAR';
                end;
     tacGlobal: begin
                  sTypeArt := '';
                end;
  end;

  if (sTypeArt <> '') then
     sReq := sReq + '" AND ACT_TYPEARTICLE="' + sTypeArt;

  sReq := sReq + '"';
  QQ := OpenSQL (sReq, True,-1,'',true);

  if Not QQ.EOF then
    begin
      dDateMax := QQ.Fields[0].AsDateTime;
      (*   QQ.Last;
      dDateMax := QQ.FieldByName('ACT_DATEACTIVITE').AsDateTime;*)
      if IsValidDate (datetostr (dDateMax)) and (dDateMax <> 0) and (dDateMax <> iDate1900) and (dDateMax <> iDate2099) then
        begin
          DecodeDate (dDateMax, iAnnee, iMois, iJour);
          CleAct.Annee := iAnnee;
          CleAct.Mois := iMois;
          CleAct.Jour := iJour;
          CleAct.Semaine := NumSemaine (dDateMax);
          Result := true;
        end;
     end;

  Finally
     Ferme (QQ);
  end;

end;

function TFActivite.AfficheDerniereDateSaisieCritere (AcdDateDefaut : TDateTime; AcbInit : boolean = false) : TDateTime;
  var
  eAnneeOnChange : TNotifyEvent;
  eMoisOnChange : TNotifyEvent;
  eSemaineOnChange : TNotifyEvent;
begin
  // débranche les evenements de gestion des dates
  eAnneeOnChange := ACT_ANNEE.OnChange;
  eMoisOnChange := ACT_MOIS.OnChange;
  eSemaineOnChange := ACT_SEMAINE.OnChange;
  ACT_ANNEE.OnChange := nil;
  ACT_MOIS.OnChange := nil;
  ACT_SEMAINE.OnChange := nil;

  try
    // Si on a passé une date en entrée et qu'on est en création de fiche
    if Not AcbInit or Not gbDateEntree then
      // Charge la derniere date de saisie pour le critere courant (ressource ou client/affaire)
      ChargerDerniereDateSaisieCritere (AcbInit);

    ACT_ANNEE.Value := CleAct.Annee;
    ACT_MOIS.Value := CleAct.Mois;
    ACT_SEMAINE.Value := CleAct.Semaine;
    Result := EncodeDate (CleAct.Annee, CleAct.Mois, CleAct.Jour);

  finally
    // Branche les evenements de gestion des dates
    ACT_ANNEE.OnChange := eAnneeOnChange;
    ACT_MOIS.OnChange := eMoisOnChange;
    ACT_SEMAINE.OnChange := eSemaineOnChange;
  end;
end;

function TFActivite.ControleDateDansIntervalle(AcdDate:TDateTime; AcbParle, AcbTous, AcbTestDatesMission :boolean;
                                                var AvbBloqueSaisie:boolean):integer;
var
dDateDebItvAutorise, dDateFinItvAutorise, dDateDebItvAff, dDateFinItvAff, dDateDebItvAct, dDateFinItvAct : TDateTime;
dDebItvCourant, dFinItvCourant : TDateTime;
bAlerteParle:boolean;
Annee, Mois, Jour, Semaine:word;
RepBlocage : T_TypeBlocAff;
begin
bAlerteParle:=true;
// En mode consultation, aucun message n'est déclenché
if bConsult.Down=true then
    begin
    AcbParle:=false;
    bAlerteParle:=false;
    end;

dDebItvCourant:=idate1900;
dFinItvCourant:=idate2099;
dDateDebItvAct:=idate1900;
dDateFinItvAct:=idate2099;
dDateDebItvAff:=idate1900;
dDateFinItvAff:=idate2099;

AvbBloqueSaisie:=false;
Result := -1;
if Not IsValidDate(DatetoStr(AcdDate)) then exit;

//
// Gestion des blocages sur la mission
//
if AcbTestDatesMission then
   begin
   // Blocage sur les dates ou l'état de la mission
   RepBlocage := BlocageAffaire( 'SAT', TOBAffaire.GetValue('AFF_AFFAIRE'), V_PGI.groupe, AcdDate, AcbParle, bAlerteParle, AcbTous, dDateDebItvAff, dDateFinItvAff, TOBAffaires);

   if (RepBlocage=tbaConfident) then Result:=10;
   if (RepBlocage=tbaEtat) then Result:=11;
   if (RepBlocage=tbaDates) then Result:=12;
   end;


//
// Gestion du blocage sur l'intervalle de dates de la saisie d'activité
//
if AcbTous or (Result=-1) then
    begin
    Result:=TestBlocageDates(AcdDate, Result);

    // Message d'avertissement pour les dates d'activite
    if (Result>1) and (Result<10) then
         begin
         IntervalleDatesActivite(dDateDebItvAct, dDateFinItvAct);
         if AcbParle then
            PGIInfoAf(HTitres.Mess[19], Caption);
         end;
    end;


//
// Analyse du résultat pour bloquage ou non du grid en saisie
//
if (Result=10) then
// si le blocage est sur la confidentialité de l'affaire on bloque
    begin
    AvbBloqueSaisie:=true;
    end
else
if (Result=11) then
// si le blocage est sur l'etat de l'affaire et qu'on est en saisie par client/mission, on bloque
    begin
    if (CleAct.TypeSaisie = tsaClient) then AvbBloqueSaisie:=true else AvbBloqueSaisie:=false;
    end
else
if (Result=12) then
// si le blocage est sur les dates de l'affaire et qu'on est en saisie par ressource, on ne bloque pas
    begin
    if (CleAct.TypeSaisie = tsaRess) then AvbBloqueSaisie:=false else AvbBloqueSaisie:=true;
    end
else
if (Result>=1) and (Result<10) then
    // Si le blocage concerne les dates sauf les cas précédents
    begin
    // on prend le plus petit intervalle
    if (dDateDebItvAct>dDateDebItvAff) then dDateDebItvAutorise:=dDateDebItvAct else dDateDebItvAutorise:=dDateDebItvAff;
    if (dDateFinItvAct<dDateDebItvAff) then dDateFinItvAutorise:=dDateFinItvAct else dDateFinItvAutorise:=dDateFinItvAff;
    // Calcul de l'intervalle courant suivant la date d'entree et le type de critères (par mois ou par semaine)
    if (VH_GC.AFTYPESAISIEACT = 'MOI') or (VH_GC.AFTYPESAISIEACT = 'FOL') then
        begin
        dDebItvCourant := DebutDeMois(AcdDate);
        dFinItvCourant := FinDeMois(AcdDate);
        end
    else
    if (VH_GC.AFTYPESAISIEACT = 'SEM') then
        begin
        DecodeDate(AcdDate, Annee, Mois, Jour);
{$IFDEF AGLINF545}
        dDebItvCourant:=PremierJourSemaineTempo(NumSemaine(AcdDate), Annee);
{$ELSE}
// PL le 23/05/02 : réparation de la fonction AGL550
        Semaine:=NumSemaine(AcdDate);
        if (Mois=12) and (Semaine=1) then Annee := Annee + 1;
        if (Mois=1) and ((Semaine=52) or (Semaine=53)) then Annee := Annee - 1;
        dDebItvCourant:=PremierJourSemaine(Semaine, Annee);
////////////////////////// AGL550
{$ENDIF}

        dFinItvCourant := dDebItvCourant+6;
        end;

    // Test sur les dates pour savoir si on doit bloquer la saisie ou non
    // on bloque si l'intervalle des dates de la saisie est en dehors de l'intervalle autorisé
    // on ne bloque pas s'il est à cheval ou dans l'intervalle des dates autorisées
    if ((dDateDebItvAutorise>dFinItvCourant) or (dDateFinItvAutorise<dDebItvCourant)) then
       AvbBloqueSaisie:=true;

    end;



if (Result=-1) then
    Result := 0;
end;


function TFActivite.DateDebutCalendrier(AcdDateACharger:TDateTime):TDateTime;
begin
result := GrCalendrier.PremierJourAffiche(AcdDateACharger);
{
DecodeDate(AcdDateACharger, Year1, Month1,Day1 );
dDateInter:=GrCalendrier.PremierJourAffiche(AcdDateACharger);

DecodeDate(dDateInter, Year2, Month2,Day2 );

if (Month1=Month2) then
   result := dDateInter
else
   begin
   result := GrCalendrier.PremierJourAffiche(dDateInter);
   end;
}
end;

function  TFActivite.CalendrierAAfficher( AcdDateACharger:TDateTime ):TDateTime;
var
dDebCalFirstLine,dDateFirstLine:TDateTime;
iAnnee, iMois, iJour:Word;
TOBL:TOB;
begin
Result := AcdDateACharger;
TOBL := GetTOBLigne(1);
if (TOBL=nil) or (TOBL.GetValue('ACT_DATEACTIVITE')=0) then exit;

dDateFirstLine := TOBL.GetValue('ACT_DATEACTIVITE');
DecodeDate(dDateFirstLine, iAnnee, iMois, iJour);

dDebCalFirstLine := GrCalendrier.PremierJourAffiche(EncodeDate(iAnnee, iMois, 1));
Result := dDebCalFirstLine;

end;


{============================================================================================}
{==========================Fonctions liees au calendrier de l'activité ======================}
{============================================================================================}
// Fonctions de l'objet TGridCalendrierActivite
Procedure TGridCalendrierActivite.ChargeSpecifGridCalendrier;
begin

end;

Procedure TGridCalendrierActivite.ObjetCalendrierJour (LeJour:TDateTime; Col,Row,NumSemaine:integer);
var
TOBJ:TOB;
begin
if (TFActivite(Ecran)=nil) then exit;
TOBJ:=nil;

if (LeJour<>0) then
   // Pour la gestion des sommes par jour
    begin
    if (TFActivite(Ecran).TOBNbHeureJour<>nil) then
        TOBJ:=TFActivite(Ecran).TOBNbHeureJour.FindFirst(['ACT_DATEACTIVITE'],[LeJour],TRUE) ;
    end
else
   // Pour la gestion de la somme par semaine
   if Not (TFActivite(Ecran).CleAct.TypeAcces in [tacFrais,tacFourn]) then
      // Seulement dans le cas de la saisie par prestation ou en global
      begin
      // Dans ce cas, c'est le numéro de semaine qui est passé dans la variable NumSemaine
      if (TFActivite(Ecran).TOBSommeSemaine<>nil) then
            TOBJ:=TFActivite(Ecran).TOBSommeSemaine.FindFirst(['SEMAINE'],[NumSemaine],TRUE) ;
      end;

if (TFActivite(Ecran).GCalendrier<>nil) then
    TFActivite(Ecran).GCalendrier.Objects[Col,Row]:= TOBJ;
end;

Procedure TGridCalendrierActivite.CalendrierDrawRect(Col,Row :integer;Rect:TRect;Ferie:boolean);
var
TOBJ:TOB;
sNbHeures:string;
R : TRect;
begin
if (TFActivite(Ecran)=nil) then exit;
try
try
//if (TFActivite(Ecran).bTobNbHeuresJourMaj=false) then
//   exit;

TOBJ:=nil;
R:= Rect;
sNbHeures := '';
if (TFActivite(Ecran).GCalendrier<>nil) then
    TOBJ:= TOB(TFActivite(Ecran).GCalendrier.Objects[Col,Row]);
if (TOBJ<>nil) then
   begin
//   try
   if (Col<>TFActivite(Ecran).GCalendrier.ColCount-1) then
      // Colonne jour
      sNbHeures := TOBJ.GetValue('ACT_QTE')
   else
      // colonne semaine
      sNbHeures := TOBJ.GetValue('SOMME');
//   except
//   end;

   if (TFActivite(Ecran).CleAct.TypeAcces in [tacFrais, tacFourn]) then
      begin
//      if (sNbHeures <> '0') then
         sNbHeures := 'X'
//      else
//         sNbHeures := '';
      end
   else
      if (sNbHeures = '0') then
         sNbHeures := '';
   end;

if (TFActivite(Ecran).GCalendrier<>nil) then
if (Col<>TFActivite(Ecran).GCalendrier.ColCount-1) then
   begin
   // Ecriture Colonne 3 du Rect
   TFActivite(Ecran).GCalendrier.Canvas.Font.Color:= clGreen;
   if (sNbHeures<>'') then
      TFActivite(Ecran).GCalendrier.Canvas.Brush.Color:= clInfoBk;
   R.Left:=R.Right-((Rect.Right - Rect.Left) div FDefRect.NbColInRect);
   end
else
   begin
   if (sNbHeures<>'') then
      sNbHeures:=StrF00(Valeur(sNbHeures),V_PGI.OkDecQ);

   if (TFActivite(Ecran).CleAct.TypeAcces in [tacFrais, tacFourn]) then
      TFActivite(Ecran).GCalendrier.Canvas.Brush.Color:= clLightGray;
   end;

finally
if (TFActivite(Ecran).GCalendrier<>nil) then
TFActivite(Ecran).GCalendrier.canvas.TextRect(R,
            R.Left + (R.Right - R.Left - TFActivite(Ecran).GCalendrier.canvas.TextWidth(sNbHeures)) div 2,
            R.Top + (R.Bottom - R.Top - TFActivite(Ecran).GCalendrier.canvas.TextHeight(sNbHeures)) div 2,
            sNbHeures);
end;
except
// si je laisse trapper ici, on boucle sur les erreurs à l'affichage (en principe il ne doit pas y en avoir !!!)
end;
end;

Procedure TGridCalendrierActivite.CalendrierDblClick(Sender: TObject);
var  IYear,IMonth,IDay : Word;
dDateCell:TDateTime;
begin

with TFActivite(Ecran) do
begin
if Not GereSaisieEnabled( true ) then exit;

dDateCell := DateSelect(FDateDeb,GCalendrier.Col,GCalendrier.Row);
DecodeDate(dDateCell ,IYear,IMonth,IDay);
ClickInsert(GS.Row, false) ;
GS.SetFocus;
GS.Cells[SG_Jour,GS.Row]:= inttostr(IDay);
GetTOBLigne(GS.Row).PutValue('ACT_DATEACTIVITE',dDateCell) ;
GetTOBLigne(GS.Row).PutValue('ACT_PERIODE', GetPeriode(dDateCell));
GetTOBLigne(GS.Row).PutValue('ACT_SEMAINE', NumSemaine(dDateCell));
end;

end;

Procedure TGridCalendrierActivite.CalendrierCellEnter(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
begin
if Not TFActivite(Ecran).gbInitOK then exit;
try
TFActivite(Ecran).MajApresChangeJourCalendrier;
finally
TFActivite(Ecran).gbSelectCellCalend:=false;
end;

end;

procedure TGridCalendrierActivite.CalendrierSelectCell(Sender: TObject; Col, Row: Longint; var CanSelect: Boolean);
Var DateDebutSemaine, DateFinSemaine, DateCell : TDateTime;
bBloqueSaisie:boolean;
YY,Semaine:integer;
Begin
if Not TFActivite(Ecran).gbInitOK then exit;

Inherited;
DateCell := DateSelect(FDateDeb,Col,Row);

If ((CanSelect = False) And (VH_GC.AFTYPESAISIEACT = 'SEM')) Then
    if TFActivite(Ecran).GCalendrier.Focused Then CanSelect := True
    else
      begin

{$IFDEF AGLINF545}
      DateDebutSemaine := PremierJourSemaineTempo(StrToInt(TFActivite(Ecran).ACT_SEMAINE.Text),StrToInt(TFActivite(Ecran).ACT_ANNEE.Text));
{$ELSE}
// PL le 23/05/02 : réparation de la fonction AGL550
      Semaine:=StrToInt(TFActivite(Ecran).ACT_SEMAINE.Text);
      YY:=StrToInt(TFActivite(Ecran).ACT_ANNEE.Text);
      if (StrToInt(TFActivite(Ecran).ACT_MOIS.Text)=12) and (Semaine=1) then YY := YY + 1;
      if (StrToInt(TFActivite(Ecran).ACT_MOIS.Text)=1) and ((Semaine=52) or (Semaine=53)) then YY := YY - 1;
      DateDebutSemaine := PremierJourSemaine(Semaine,YY);
//////////////////////////// AGL550
{$ENDIF}

      DateFinSemaine := DateDebutSemaine+ NbJourSemaine ;
      If ((DateCell >= DateDebutSemaine ) Or (DateCell <  DateFinSemaine)) Then CanSelect := True;
      End;

// Controle de la validite de la date
If CanSelect then
   if TFActivite(Ecran).GCalendrier.Focused Then
      TFActivite(Ecran).ControleDateDansIntervalle(DateCell, true, false, (TFActivite(Ecran).CleAct.TypeSaisie = tsaClient), bBloqueSaisie);

End;

// Fonctions de l'objet TFActivite
procedure TFActivite.MajApresChangeJourCalendrier;
  var
  IYear, IMonth, IDay : Word;
  IYearOld, IMonthOld : Word;
  eOnChange : TNotifyEvent;
  Semaine : string;
begin
  //GCalendrier.Enabled := false;
  try
    IYearOld := CleAct.Annee;
    IMonthOld := CleAct.Mois;
    // PL le 27/05/02 : Nouvelles fonctions AGL NumSemaine et PremierJourSemaine
    ///////////////////////////

    DecodeDate (GrCalendrier.DateSelect (GrCalendrier.FDateDeb, GCalendrier.Col, GCalendrier.Row), IYear, IMonth, IDay);
    // Annee
    if (IYear <> IYearOld) then
      begin
        eOnChange := ACT_ANNEEChange;
        ACT_ANNEE.OnChange := nil;
        ACT_ANNEE.Text := inttostr (IYear);
        CleAct.Annee := IYear;
        ACT_ANNEE.OnChange := eOnChange;
      end;

    // Mois
    if (VH_GC.AFTYPESAISIEACT <> 'SEM') then
      begin
        ACT_MOIS.Text := inttostr (IMonth);
        CleAct.Mois := IMonth;
      end
    else
    if (IMonth <> IMonthOld) then
      begin
        eOnChange := ACT_MOISChange;
        ACT_MOIS.OnChange := nil;
        ACT_MOIS.Text := inttostr (IMonth);
        CleAct.Mois := IMonth;
        ACT_MOIS.OnChange := eOnChange;
        // Si on change le mois, il faudra forcer le rafraichissement des données mensuelles sur les compléments
        gbModifMois := true;
      end;
    // Jour
    CleAct.Jour := IDay;


    // Semaine
    Semaine := inttostr (NumSemaine (EncodeDate (IYear, IMonth, IDay)));
    if (VH_GC.AFTYPESAISIEACT = 'SEM') then
      begin
        ACT_SEMAINE.Text := Semaine;
        Semaine := ACT_SEMAINE.Text;
      end;

    CleAct.Semaine := strtoint (Semaine);

  finally
    //GCalendrier.Enabled := true;
  end;
end;

procedure TFActivite.ChargerCalendrierGrid;
  var
  iJour, iMois, iAnnee : integer;
  Inc : integer;
  sJour : string;
  iSemaineCourante, iSemaineCible : integer;
  bDateValide : boolean;
begin
  iJour := 1;
  iSemaineCible := 0;
  iAnnee := CleAct.Annee;
  iMois := CleAct.Mois;
  sJour := GS.Cells [SG_Jour, GS.Row];
  if (sJour <> '') then
    begin
      iJour := strtoint (sJour);
      if (VH_GC.AFTYPESAISIEACT = 'SEM') then
        // En saisie par semaine, on s'assure que la date cible est bien dans la même semaine que la
        // date courante du calendrier
        begin
          //   try
          iSemaineCourante := strtoint (ACT_SEMAINE.Text);
          try
            bDateValide := true;
            iSemaineCible := NumSemaine (EncodeDate (iAnnee, iMois, iJour));
          except
            // Toujours la problèmatique de gestion de date non valide (30/02/XXXX par exemple)
            // gérée par le bloque qui suit
            bDateValide := false;
          end;

          if (iSemaineCourante <> iSemaineCible) or (bDateValide = false)then
            begin
              if (iJour >= 15) then Inc := -1 else Inc := 1;
              iMois := iMois + Inc;
              if (iMois = 0) then
                begin
                  iMois := 12;
                  iAnnee:=iAnnee-1;
                end
              else
              if (iMois=13) then
                begin
                  iMois := 1;
                  iAnnee:=iAnnee+1;
                end;

              iSemaineCible := NumSemaine (EncodeDate (iAnnee, iMois, iJour));
              if ((iSemaineCourante <> 53) or (iSemaineCible <> 1)) and ((iSemaineCourante <> 1) or (iSemaineCible <> 53)) then
              if (iSemaineCourante <> iSemaineCible) then exit;
            end;
          //   except
          //   exit;
          //   end;
        end;
    end;

  if ChargerCalendrierEnCours (iJour, iMois, iAnnee, true) then
    MajApresChangeJourCalendrier;

end;

function TFActivite.ChargerCalendrierEnCours (AciJour, AciMois, AciAnnee : integer; AcbForcer : boolean) : boolean;
  var
  DateCible : TDateTime;
begin
  Result := false;
  // PL le 28/01/02 pour éviter plantage aléatoire en arrivant sur l'écran
  if not gbInitOK then exit;
  // Fin PL le 28/01/02
  Result := true;
  try
    DateCible := EncodeDate (AciAnnee, AciMois, AciJour);

    GrCalendrier.ChangerDate (DateCible, AcbForcer);
  except
    // on trappe le cas de la date non valide (voir excepts précédents) toujours de la même
    // façon. En l'occurrence, on a juste besoin de savoir que la fonction s'est mal passé
    Result := false;
  end;
end;

function TFActivite.ForcerRefreshCalendrier(AciNewYear, AciNewMonth, AciNewDay:integer):boolean;
var
iNewWeek:integer;
bChangeCritere:boolean;
begin
Result := True;
bChangeCritere:=false;
if (CleAct.TypeSaisie = tsaRess) then
   begin
   if (CleActOld.Ressource <> CleAct.Ressource) then bChangeCritere:=true;
   end
else
   begin
   if ((CleActOld.Affaire <> CleAct.Affaire) or (CleActOld.Tiers <> CleAct.Tiers)) then
      bChangeCritere:=true;
   end;

if ((AciNewYear <> CleAct.Annee) or (AciNewMonth <> CleAct.Mois)
   or (CleActOld.Annee <> CleAct.Annee) or (CleActOld.Mois <> CleAct.Mois)
   or (CleActOld.Folio <> CleAct.Folio) or bChangeCritere) then exit;

if (VH_GC.AFTYPESAISIEACT = 'SEM') then
   begin
   iNewWeek := NumSemaine(EncodeDate(AciNewYear, AciNewMonth, AciNewDay ));
   if (iNewWeek <> CleAct.Semaine) then exit;
   if (CleActOld.Semaine <> CleAct.Semaine) then exit;
   end;

Result := False;
end;

procedure TFActivite.MajTotauxCalendrier(AcdDateDebut:TDateTime);
var
  TOBMaj, TOBSemMAJ : TOB;
  sCritere : string;
begin
  TOBSemMAJ := nil;
  // Charge la TOB Activité
  if (CleAct.TypeSaisie = tsaRess) then
    sCritere := ACT_Ressource.Text
  else
    sCritere := ACT_Affaire.Text;

  // Vide la TOB NBHeureJour et la recharge
  if (gbSelectCellCalend = false) then
    begin
      TOBMaj := ChargerTOBNbHeureJour (sCritere, ACT_FOLIO.Text, AcdDateDebut, CleAct.TypeAcces);
      if (TOBMaj <> nil) then
        begin
          try
            // Calcul de la somme par semaine
            TOBSemMAJ := ChargerTOBSommeSemaine (TOBMaj);

            // Maj de la TOB nombre d'heures par jour
            TOBNbHeureJour.ClearDetail;
            TOBNbHeureJour.Dupliquer (TOBMaj, TRUE, TRUE, TRUE);

            // Calcul de la somme par mois
            TOBNbHeureJourAffiche.ClearDetail;
            if (CleAct.TypeAcces in [tacTemps, tacGlobal]) then
              TOBNbHeureJourAffiche.Dupliquer (TOBMaj, TRUE, TRUE, TRUE)
            else
              begin
              if (TOBNbHeureJourAffiche <> nil) then TOBNbHeureJourAffiche.Free;
              TOBNbHeureJourAffiche := ChargerTOBNbHeureJour(sCritere, ACT_FOLIO.Text, AcdDateDebut, tacTemps);
              end;

            // Maj de la TOB nombre d'heures par semaine
            TOBSommeSemaine.ClearDetail;
            if (TOBSemMAJ <> nil) then
              begin
                TOBSommeSemaine.Dupliquer (TOBSemMAJ, TRUE, TRUE, TRUE);
              end;

          finally
            TOBMaj.Free;
            TOBSemMAJ.Free;
          end;
        end;
    end;
end;

procedure TFActivite.AfficheTotauxParMois;
var
dPourcFac, dPourcNFac : double;
begin

// Temps
dPourcFac:=0;
dPourcNFac:=0;
if (gdNbHeureMois<>0) then
    dPourcFac:=(gdNbHeureFac/gdNbHeureMois)*100;
if (gdNbHeureMois<>0) then
    dPourcNFac:=(gdNbHeureNFac/gdNbHeureMois)*100;
CumulTemps.caption := StrF00(gdNbHeureMois,V_PGI.OkDecQ) ;
if (CumulTemps.caption='') then CumulTemps.caption:='0';
CumulFac.caption := StrF00(gdNbHeureFac,V_PGI.OkDecQ) ;
if (CumulFac.caption='') then CumulFac.caption:='0';
CumulNonFac.caption := StrF00(gdNbHeureNFac,V_PGI.OkDecQ) ;
if (CumulNonFac.caption='') then CumulNonFac.caption:='0';
PourcFac.caption := Format('(%3.1f%%)', [dPourcFac]);
PourcNonFac.caption := Format('(%3.1f%%)', [dPourcNFac]);
// Frais
CumulFrais.caption := StrF00(gdSommeFrais,V_PGI.OkDecV);
if (CumulFrais.caption='') then CumulFrais.caption:='0';
end;

procedure TFActivite.GCalendrierExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;
gbSelectCellCalend:=false;
end;

procedure TFActivite.GCalendrierMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
gbSelectCellCalend:=true;
end;

{============================================================================================}
{========================== Dessin des totaux en bas de Grid ================================}
{============================================================================================}
procedure TFActivite.DessineTotaux;
var
Rect:TRect;
begin
Rect := GS.cellRect( SG_TOTPRCH, 0);
LHNumEditTOTPR.Left := Rect.Left;
LHNumEditTOTPR.width := Rect.Right - Rect.Left;

Rect := GS.cellRect( SG_Qte, 0);
LHNumEditTOTQTE.Left := Rect.Left;
LHNumEditTOTQTE.width := Rect.Right - Rect.Left;
HLabelUniteBase.Left := LHNumEditTOTQTE.Left-HLabelUniteBase.width - 20;

Rect := GS.cellRect( SG_TotV, 0);
LHNumEditTOTPV.Left := Rect.Left;
LHNumEditTOTPV.width := Rect.Right - Rect.Left;

//LblMois.Left := GCalendrier.Left;
end;

procedure TFActivite.MajTotauxActivite;
var
cTotMontants, cTotMoins:double;
cTotMontantsPV, cTotMoinsPV:double;
cTotQte:double;
TOBC:TOB;
begin
try
TOBC:=TOB.Create('Entete activite',Nil,-1) ;
cTotMontants:=0;
cTotMontantsPV:=0;
cTotQte:=0;
// Mise à jour des totaux
// PL le 15/01/02 pour V575
if (TOBActivite.Detail.Count <> 0) then
if (CleAct.TypeSaisie = tsaRess) then
   begin
   if (CleAct.TypeAcces in [tacTemps]) then
        begin
        cTotMontants:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_RESSOURCE','ACT_ACTIVITEEFFECT'],[ACT_RESSOURCE.Text,'X'],TRUE) ;
        cTotMontantsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_RESSOURCE','ACT_ACTIVITEEFFECT'],[ACT_RESSOURCE.Text,'X'],TRUE) ;
        end
   else
        begin
        cTotMontants:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_RESSOURCE'],[ACT_RESSOURCE.Text],TRUE) ;
        cTotMontantsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_RESSOURCE'],[ACT_RESSOURCE.Text],TRUE) ;
        cTotMoins:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_RESSOURCE','ACT_TYPEARTICLE','ACT_ACTIVITEEFFECT'],[ACT_RESSOURCE.Text,'PRE','-'],TRUE) ;
        cTotMoinsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_RESSOURCE','ACT_TYPEARTICLE','ACT_ACTIVITEEFFECT'],[ACT_RESSOURCE.Text,'PRE','-'],TRUE) ;
        cTotMontants := cTotMontants - cTotMoins;
        cTotMontantsPV := cTotMontantsPV - cTotMoinsPV;
        end;
//   cTotQte:=TOBActivite.Somme('ACT_QTE',['ACT_RESSOURCE'],[ACT_RESSOURCE.Text],TRUE) ;
   end
else
   begin
   if (CleAct.TypeAcces in [tacTemps]) then
        begin
        cTotMontants:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_AFFAIRE','ACT_ACTIVITEEFFECT'],[ACT_AFFAIRE.Text,'X'],TRUE) ;
        cTotMontantsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_AFFAIRE','ACT_ACTIVITEEFFECT'],[ACT_AFFAIRE.Text,'X'],TRUE) ;
        end
   else
        begin
        cTotMontants:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_AFFAIRE'],[ACT_AFFAIRE.Text],TRUE) ;
        cTotMontantsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_AFFAIRE'],[ACT_AFFAIRE.Text],TRUE) ;
        cTotMoins:=TOBActivite.Somme('ACT_TOTPRCHARGE',['ACT_AFFAIRE','ACT_TYPEARTICLE','ACT_ACTIVITEEFFECT'],[ACT_AFFAIRE.Text,'PRE','-'],TRUE) ;
        cTotMoinsPV:=TOBActivite.Somme('ACT_TOTVENTE',['ACT_AFFAIRE','ACT_TYPEARTICLE','ACT_ACTIVITEEFFECT'],[ACT_AFFAIRE.Text,'PRE','-'],TRUE) ;
        cTotMontants := cTotMontants - cTotMoins;
        cTotMontantsPV := cTotMontantsPV - cTotMoinsPV;
        end;
//   cTotQte:=TOBActivite.Somme('ACT_QTE',['ACT_AFFAIRE'],[ACT_AFFAIRE.Text],TRUE) ;
   end;
// Fin PL le 15/01/02 pour V575

// On fait la somme de toutes les quantités de la TOB Activite dans la même unité : unite de conversion par defaut
if (LHNumEditTOTQTE.visible = true) then
   begin
   TOBC.Dupliquer(TOBActivite,TRUE,TRUE,false) ;
   cTotQte := ConversionTOBActivite( TOBC, 'ACT_DATEACTIVITE' );
   end;

HNumEditTOTPR.Value := cTotMontants;
HNumEditTOTPV.Value := cTotMontantsPV;
HNumEditTOTQTE.Value := cTotQte;


finally
TOBC.Free;
end;
end;

procedure TFActivite.MajApresModifActivite;
var
TOBL:TOB;
begin
// Mise à jour des totaux
MajTotauxActivite;

// Mise à jour des champs du pieds de page
TOBL:=nil;
if (TOBActivite <> nil) then
   begin
   TOBL := GetTOBLigne(GS.Row);
   if TOBL <> nil then AfficheDetailTOBLigne( GS.Row ) ;
   end;

if (TOBL = nil) then
   begin
   ClearPanel(PPied);
   ClearPanel(PPiedAvance);
//   ClearPanelPTotaux2;
   if (bDescriptif.Down = true) then
      ACT_DESCRIPTIF.Clear;
   end
else
   // Maj de la zone complement à l'écran
   GereDescriptif(GS.Row,True);

end;

{=================================================================================================}
{================================= Gestion du pied ===============================================}
{========================== Range dans le grid les modifs du pied ================================}
{=================================================================================================}
procedure TFActivite.ACT_UNITEExit (Sender : TObject);
  Var
  TOBL, TOBA : TOB;
  bOK : boolean;
BEGIN
  if csDestroying in ComponentState then Exit ;
  if (ACT_UNITE.ReadOnly = true) then exit;
  TOBA := nil;
  FormateZoneSaisie (ACT_UNITE, SG_Unite, GS.Row);

  TOBL := GetTOBLigne(GS.Row);
  if TOBL = Nil then Exit;

  bOK := true;
  if (ACT_UNITE.Text <> '') then
  if Not RemplirTOBUnites ( TOBListeConv, ACT_UNITE.Text, TOBUnites) then
    // Si on ne trouve pas l'unite dans la liste principale, on cherche dans la liste
    // des unites de type 'AUC'
    begin
      if (TOBListeConvAUC <> nil) then
        TOBA := TOBListeConvAUC.FindFirst (['GME_MESURE'], [ACT_UNITE.Text], TRUE);
      if (TOBA <> nil) then
        bOK := true
      else
        bOK := GereElipsis (ACT_UNITE, SG_Unite);
    end;

  if bOK = true then
    begin
      TOBL.PutValue ('ACT_UNITE', ACT_UNITE.Text);
      TOBL.PutValue ('ACT_UNITEFAC', ACT_UNITE.Text);
    end
  else
    begin
      ACT_UNITE.Text := TOBL.GetValue ('ACT_UNITE');
      if (ACT_UNITE.Text <> '') then
        ACT_UNITE.SetFocus
      else
        begin
          TOBL.PutValue ('ACT_UNITE', '');
          TOBL.PutValue ('ACT_UNITEFAC', '');
        end;
    end;

  // Mise à jour des totaux
  MajTotauxActivite;

  MajTOBHeureCalendrier (GS.Row);

  AfficheLaTOBLigne (GS.Row);
  TOBL.PutEcran (Self, PPied);
  PutEcranPPiedAvance (TOBL);
end;

procedure TFActivite.ACT_UNITEElipsisClick (Sender : TObject);
begin
  if (ACT_UNITE.ReadOnly = true) then exit;
  IdentifieUnites (ACT_UNITE);
end;

procedure TFActivite.ACT_QTEExit (Sender : TObject);
  Var
  TOBL : TOB ;
begin
  if csDestroying in ComponentState then Exit;
  TOBL := GetTOBLigne (GS.Row);
  if TOBL = Nil then Exit;
  if (THNumEdit (Sender).ReadOnly = true) then exit;
  GS.Cells[SG_QTE, GS.Row] := THNumEdit (Sender).Text;
  TraiteQte (GS.Row);

  MajTOBHeureCalendrier (GS.Row);
  AfficheLaTOBLigne (GS.Row);
  TOBL.PutEcran (Self, PPied);
  PutEcranPPiedAvance (TOBL);

end;

procedure TFActivite.ACT_PUPRCHARGEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if (THNumEdit(Sender).ReadOnly=true) then exit;
  GS.Cells[SG_PUCH,GS.Row]:=THNumEdit(Sender).Text;
  TraitePUCH ( GS.Row ) ;
  AfficheLaTOBLigne ( GS.Row ) ;
end;

procedure TFActivite.ACT_PUVENTEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if (THNumEdit(Sender).ReadOnly=true) then exit;
  GS.Cells[SG_PV,GS.Row]:=THNumEdit(Sender).Text;
  TraitePV ( GS.Row ) ;
  AfficheLaTOBLigne ( GS.Row ) ;
end;

procedure TFActivite.ACT_TOTPRCHARGEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;
  if (THNumEdit(Sender).ReadOnly=true) then exit;
  GS.Cells[SG_TOTPRCH,GS.Row]:=THNumEdit(Sender).Text;
  TraiteTOTPRCH ( GS.Row ) ;
  AfficheLaTOBLigne ( GS.Row ) ;

end;

procedure TFActivite.ACT_TOTVENTEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;

  if (THNumEdit(Sender).ReadOnly=true) then exit;
  GS.Cells[SG_TotV,GS.Row]:=THNumEdit(Sender).Text;
  TraiteTotV ( GS.Row ) ;
  AfficheLaTOBLigne ( GS.Row ) ;

end;

procedure TFActivite.ACT_ACTIVITEREPRISExit(Sender: TObject);
  Var TOBL : TOB ;
begin
  if csDestroying in ComponentState then Exit ;
  if (THCritMaskEdit(Sender).ReadOnly=true) then exit;
  TOBL:=GetTOBLigne(GS.Row) ; if TOBL=Nil then Exit ;
  FormateZoneSaisie (THCritMaskEdit(Sender), SG_ActRep, GS.Row );

  if Not LookupValueExist(THCritMaskEdit(Sender)) and (THCritMaskEdit(Sender).Text <>'') then
     THCritMaskEdit(Sender).SetFocus;

  TOBL.PutValue('ACT_ACTIVITEREPRIS',THCritMaskEdit(Sender).Text) ;
  gbModifActRep := true;
  AfficheLaTOBLigne ( GS.Row ) ;
end;

procedure TFActivite.ACT_ACTORIGINEExit(Sender: TObject);
  Var TOBL : TOB ;
begin
  if csDestroying in ComponentState then Exit ;
  if (THValComboBox(Sender).Enabled=false) then exit;
  TOBL:=GetTOBLigne(GS.Row) ; if TOBL=Nil then Exit ;

  TOBL.PutValue('ACT_ACTORIGINE',THValComboBox(Sender).Value) ;
end;


procedure TFActivite.ACT_TYPEHEUREEnter(Sender: TObject);
begin
  if (THCritMaskEdit(Sender).ReadOnly=true) then THCritMaskEdit(Sender).DataType:=''
                                   else THCritMaskEdit(Sender).DataType:='AFTTYPEHEURE';
  if Not GereSaisieEnabled( true ) then abort;
end;

procedure TFActivite.ACT_TYPEHEUREExit(Sender: TObject);
  Var TOBL : TOB;
begin
  if csDestroying in ComponentState then Exit ;
  TOBL:=GetTOBLigne(GS.Row) ; if TOBL=Nil then Exit ;
  if (THCritMaskEdit(Sender).ReadOnly=true) then exit;
  GS.Cells[SG_TYPEHEURE,GS.Row]:=THCritMaskEdit(Sender).Text;
  if Not LookupValueExist(THCritMaskEdit(Sender)) and (THCritMaskEdit(Sender).Text <>'') then
     BEGIN THCritMaskEdit(Sender).SetFocus; Exit; END;

  TraiteTypeHeure ( GS.Row ) ;
  AfficheLaTOBLigne ( GS.Row ) ;
  TOBL.PutEcran(Self, PPied) ;
  PutEcranPPiedAvance(TOBL);
end;


procedure TFActivite.ACT_LIBELLEExit(Sender: TObject);
  Var TOBL : TOB ;
begin
  if csDestroying in ComponentState then Exit ;
  TOBL:=GetTOBLigne(GS.Row) ; if TOBL=Nil then Exit ;

  if (TOBL.GetValue('ACT_LIBELLE')=ACT_LIBELLE.Text) then exit;

  TOBL.PutValue('ACT_LIBELLE', ACT_LIBELLE.Text);
  gbModifLibArticle := true; // PL le 13/06/03 : le libelle article a été modifié


  AfficheLaTOBLigne ( GS.Row ) ;
  bActModifiee:=true;
end;

procedure TFActivite.ClearPanel(AcPanel:THPanel);
  var
  i:integer;
begin
  for i:=0 to AcPanel.ControlCount-1 do
     begin
     if (AcPanel.Controls[i] is THNumEdit) then
        THNumEdit(AcPanel.Controls[i]).Value := 0;

     if (AcPanel.Controls[i] is THCritMaskEdit) then
        THCritMaskEdit(AcPanel.Controls[i]).Text := '';

     if (AcPanel.Controls[i] is THPanel) then
        ClearPanel(THPanel(AcPanel.Controls[i]));
     end;
end;
(*
procedure TFActivite.ClearPanelPTotaux2;
begin
  CumulTemps.Caption:='0';
  CumulFac.Caption:='0';
  CumulNonFac.Caption:='0';
  CumulFrais.Caption:='0';
  LblMoisTotaux.Caption:='';
  PourcFac.Caption:='(0%)';
  PourcNonFac.Caption:='(0%)';
end;
*)
procedure TFActivite.ReadOnlyPanel(AcPanel:THPanel; AcState:boolean);
  var
  i:integer;
begin
  for i:=0 to AcPanel.ControlCount-1 do
     begin
     if (AcPanel.Controls[i] is THNumEdit) OR (AcPanel.Controls[i] is THCritMaskEdit) then
          begin
          if (AcState = true) then
              begin
              if (TEdit(AcPanel.Controls[i]).ReadOnly = true) then
                  TEdit(AcPanel.Controls[i]).Tag := -1
              else
                  TEdit(AcPanel.Controls[i]).ReadOnly := true;
              end
          else
              begin
              if (TEdit(AcPanel.Controls[i]).Tag <> -1) then
                  TEdit(AcPanel.Controls[i]).ReadOnly := false;
              end;
          end
     else
     if (AcPanel.Controls[i] is THValComboBox) then
          begin
          if (AcState = true) then
              begin
              if (THValComboBox(AcPanel.Controls[i]).Enabled = false) then
                  THValComboBox(AcPanel.Controls[i]).Tag := -1
              else
                  THValComboBox(AcPanel.Controls[i]).Enabled := false;
              end
          else
              begin
              if (THValComboBox(AcPanel.Controls[i]).Tag <> -1) then
                  THValComboBox(AcPanel.Controls[i]).Enabled := true;
              end;
          end;

     if (AcPanel.Controls[i] is THPanel) then
        ReadOnlyPanel(THPanel(AcPanel.Controls[i]), AcState);
     end;
end;

procedure TFActivite.ACT_ACTIVITEREPRISEnter(Sender: TObject);
begin
  if (THCritMaskEdit(Sender).ReadOnly=true) then
     THCritMaskEdit(Sender).DataType:=''
  else
     THCritMaskEdit(Sender).DataType:='AFTACTIVITEREPRISE';

  if Not GereSaisieEnabled( true ) then abort;
end;

procedure TFActivite.ACT_UNITEEnter(Sender: TObject);
  var
  TOBT:TOB;
begin
  if Not GereSaisieEnabled( true ) then abort;

  // on memorise la TOB activite de la ligne avant toute modif
  TOBT := GetTOBLigne(GS.Row) ;
  TOBLigneActivite.Dupliquer(TOBT,TRUE,TRUE,TRUE) ;
  ConversionTobActivite( TOBEnteteLigneActivite, 'ACT_DATEACTIVITE' );

end;

procedure TFActivite.ACT_QTEEnter (Sender : TObject);
  var
  TOBT : TOB;
begin
  if Not GereSaisieEnabled (true) then abort;

  // on memorise la TOB activite de la ligne avant toute modif
  TOBT := GetTOBLigne (GS.Row);
  TOBLigneActivite.Dupliquer (TOBT, TRUE, TRUE, TRUE);
  ConversionTobActivite (TOBEnteteLigneActivite, 'ACT_DATEACTIVITE');
end;

procedure TFActivite.ACT_PUPRCHARGEEnter (Sender : TObject);
begin
  if Not GereSaisieEnabled (true) then abort;
end;

procedure TFActivite.ACT_PUVENTEEnter (Sender : TObject);
begin
  if Not GereSaisieEnabled (true) then abort;
end;

procedure TFActivite.ACT_TOTPRCHARGEEnter (Sender : TObject);
begin
  if Not GereSaisieEnabled (true) then abort;
end;

procedure TFActivite.ACT_TOTVENTEEnter (Sender : TObject);
begin
  if Not GereSaisieEnabled (true) then abort;
end;

procedure TFActivite.pInfo1Enter (Sender : TObject);
begin
  OKSaisiePartieDroiteGridGS (GS.Row, true);
  if Not GereSaisieEnabled (true) then abort;
end;

procedure TFActivite.HNumEditTOTPRChange (Sender : TObject);
begin
  AfficheTotauxDansLibelle (Sender, Self );
end;


{=================================================================================================}
{================================ Gestion du descriptif (Memo) ===================================}
{=================================================================================================}
Procedure TFActivite.GereDescriptif (ARow : Integer; Enter : Boolean);
  Var
  TOBL : Tob;
  Tmp : String;
Begin

  TOBL := GetTobLigne (ARow);
  if TOBL = Nil then Exit ;

  If Enter Then
    Begin
      StringtoRich (ACT_DESCRIPTIF, TOBL.GetValue ('ACT_DESCRIPTIF'));
      TDescriptif.Caption := Format ('Descriptif complémentaire de la ligne N° %d', [ARow]);

      if BlocageLigne (TOBL) then
        ACT_DESCRIPTIF.Enabled := false
      else
        ACT_DESCRIPTIF.Enabled := true;
    End
  Else
    begin
      Tmp := ACT_DESCRIPTIF.Text;
      if (Length (Tmp) <> 0) and (Tmp <> #$D#$A) then
        begin
          TOBL.PutValue ('ACT_DESCRIPTIF', ExRichToString (ACT_DESCRIPTIF));
          GS.Cells [SG_Desc,ARow] := '&';
        end
      else
        begin
          if (TOBL.GetValue ('ACT_DESCRIPTIF') <> '') then
            TOBL.PutValue ('ACT_DESCRIPTIF', '');
          GS.Cells [SG_Desc, ARow] := '';
        end;
    end;
End;

procedure TFActivite.TDescriptifClose(Sender: TObject);
begin
  GereDescriptif (GS.Row, False);

  bDescriptif.Down := False;
  Afficherlemmo1.Checked := bDescriptif.Down;
end;

Procedure TFActivite.PutEcranPPiedAvance (TOBL : TOB);
  var
  CleDocVente : R_CleDoc ;
  NumPiece : integer;
  PA : Double;
begin
  TOBL.PutEcran (Self, PPiedAvance);
  // Traitement manuel des num de pièces
  if TOBL.GetValue ('ACT_NUMPIECE') <> '' then
    begin
      DecodeRefPiece (TOBL.GetValue ('ACT_NUMPIECE'), CleDocVente);
      NumPieceVente.Caption := Format ('%s n° %d du %s', [CleDocVente.NaturePiece, CleDocVente.NumeroPiece, DateToStr (CleDocVente.DatePiece)]);
    end
  else
    NumPieceVente.Caption := '';
  if TOBL.GetValue ('ACT_NUMERO') <> 0 then
    begin
      NumPiece := TOBL.GetValue ('ACT_NUMERO');
      NumPieceAchat.Caption := Format ('%s n° %d ',[TOBL.GetValue('ACT_NATUREPIECEG'), NumPiece]);
      PA := TOBL.GetValue ('ACT_PUPR');
      PrixAchat.Caption := Format ('Prix d''achat unitaire : %n ', [PA]);
    end
  else
    begin
      NumPieceAchat.Caption := '';
      PrixAchat.Caption := '';
    end;
end;

procedure TFActivite.PTotaux1Enter (Sender: TObject);
begin
  OKSaisiePartieDroiteGridGS (GS.Row, true);
end;

procedure TFActivite.PTotaux2Enter (Sender: TObject);
begin
  OKSaisiePartieDroiteGridGS (GS.Row, true);
end;


procedure TFActivite.BMoisMoinsClick(Sender: TObject);
begin
  if (ACT_MOIS.CanFocus) then ACT_MOIS.SetFocus;
  ACT_MOIS.Value := ACT_MOIS.Value - 1;
end;

procedure TFActivite.BMoisPlusClick(Sender: TObject);
begin
  if (ACT_MOIS.CanFocus) then ACT_MOIS.SetFocus;
  ACT_MOIS.Value := ACT_MOIS.Value + 1;
end;



procedure TFActivite.BDATEFACTAFFClick(Sender: TObject);
  Var DateFact : TDateTime;
begin
  if (TobAffaire = Nil) or (TOBLigneActivite = Nil) Then Exit;
  if TobAffaire.GetValue ('AFF_GENERAUTO') <> 'ACT' then Exit;
  if TOBLigneActivite.GetValue ('ACT_AFFAIRE') = '' then Exit;
  if (TOBLigneActivite.GetValue ('ACT_AFFAIRE') <> TOBAffaire.GetValue ('AFF_AFFAIRE')) then exit;

  DateFact := StrToDate(DateFactAff.Text);
  if DateFact = iDate1900 then
    BEGIN
      PGIBoxAf ('Date d''échéance de facturation incorrecte', 'Génération d''échéance de facturation');
      Exit;
    END;
  if not (CreerFactAff (TOBAffaire, Datefact, 'NOR')) then
     PGIBoxAf ('la création d''échéance a échouée', 'Génération d''échéance de facturation')
  else
     PGIBoxAf ('Echéance générée avec succès', 'Génération d''échéance de facturation');

end;


Function AGLAFCreerActivite (Parms : array of variant ; nb : integer) : Variant;
BEGIN
  Result := AFCreerActivite (T_Sai (Integer (Parms [0])), T_Acc(Integer (Parms [1])), Parms [2]);
END;

procedure TFActivite.BAbandonClick(Sender: TObject);
begin
  Close;
  if FClosing and IsInside (Self) then THPanel (parent).CloseInside;
end;

procedure TFActivite.SetConsultation (AcbConsult : boolean; ActTypBloc : T_TypeBlocageActivite);
begin
  if AcbConsult then
    tEnsBlocages := tEnsBlocages + [ActTypBloc]
  else
    tEnsBlocages := tEnsBlocages - [ActTypBloc];



  if (tEnsBlocages <> []) then
    begin
      ReadOnlyPanel (PPied, True);
      ReadOnlyPanel (PPiedAvance, True);
      AffecteGrid (GS, taConsult) ;
      bNewligne.OnClick := nil;
      bDelLigne.OnClick := nil;
      GCalendrier.OnDblClick := nil;
      Supprimeuneligne1.Enabled := false;
      Ajoutduneligne1.Enabled := false;
      Couper1.Enabled := false;
      Coller1.Enabled := false;
    end
  else
    begin
      ReadOnlyPanel (PPied, false);
      ReadOnlyPanel (PPiedAvance, false);
      AffecteGrid (GS, taCreat) ;
      bNewligne.OnClick := bNewligneClick;
      bDelLigne.OnClick := bDelLigneClick;
      GCalendrier.OnDblClick := GrCalendrier.CalendrierDblClick;
      Supprimeuneligne1.Enabled := true;
      Ajoutduneligne1.Enabled := true;
      Couper1.Enabled := true;
      if (Coller1.Tag <> -1) then Coller1.Enabled := true;
    end;
end ;

procedure TFActivite.bConsultClick(Sender: TObject);
begin

// Modif des options du grid de saisie de l'activite en fonction du mode (creat,modif …) => property rowSelect …
if bConsult.Down then
    begin
    bConsult.Down:=false;
    GSExit(GS);
    if (GCanClose=true) then
        begin
        SetConsultation(true, tbaModeConsult);
        bConsult.Down:=true;
        end;
    end
else
    begin
    if (Action<>taConsult) then
        begin
        SetConsultation(false, tbaModeConsult);
        GS.Cells[SG_Curseur, GS.Row] := '';
        GS.Row := 1;
        GS.Col := SG_Jour;
        EntreDansLigneCourante(GS.Col,GS.Row);
        end;
    end;
end;

procedure TFActivite.b35heuresClick(Sender: TObject);
var
i,Index,iNumSemErreur1,iNumSemErreur2:integer;
TobCalRegle:TOB;
dDateDeRef, dDerniereDateEffet:TDateTime;
dSommeAn:double;
Reponse:integer;
sMessage,sArgBlocAn,sArgAlertAn:string;
tEnsBloc35h:T_EnsBlocages35H;
dNbHMaxSmnLim1,dNbHMaxSmnLim2,dNbHDepaceLim1,dNbHDepaceLim2:double;
iNbSemLim1,iNbSemLim2:integer;
bSemConsecLim1,bSemConsecLim2:boolean;
begin
dDerniereDateEffet:=idate1900;
tEnsBloc35h:=[];
Reponse:=0; iNumSemErreur1:=0; iNumSemErreur2:=0;
dNbHMaxSmnLim1:=0;dNbHMaxSmnLim2:=0;dNbHDepaceLim1:=0;dNbHDepaceLim2:=0;dSommeAn:=0;
iNbSemLim1:=0;iNbSemLim2:=0;
bSemConsecLim1:=false;bSemConsecLim2:=false;
sMessage:='';sArgBlocAn:='';sArgAlertAn:='';
if (CleAct.TypeSaisie <> tsaRess) then exit;
//
// commencer par valider la ligne en cours de saisie
//
GSExit(GS);
//
// ensuite on valide la saisie
//
if (GCanClose=false) then exit;

//
// Teste sur le nombre d'heures dans l'année
//
// La date de référence est la date de début d'analyse 35 Heures
// elle sert à sélectionner le calendrier 35h avec la bonne date d'effet
dDateDeRef:= GetParamSoc('SO_AFDATEDEB35');

Index := AFOAssistants.AddRessource(CleAct.Ressource, true);
if (Index<>-1) and (Index<>-2) then
    begin
    dSommeAn := TAFO_Ressource(AFOAssistants.GetRessource(Index)).gdTempsAnnuel;
    TobCalRegle := TAFO_Ressource(AFOAssistants.GetRessource(Index)).TobCalendrierRegle;

    if (TobCalRegle<>nil) then
        begin
        // La Tob étant triée en ordre décroissant, on récupère la première date d'effet juste inférieure
        // à la date saisie et l'indice de la première TOB calendrier valable à la date d'effet
        dDerniereDateEffet:=TAFO_Ressource(AFOAssistants.GetRessource(Index)).DetDerniereDateEffetCalendrier(dDateDeRef, i);

        if (i<>-1) then
            begin
            // on boucle mais en principe on ne peut avoir qu'une seule règle par date d'effet
            while (i<TobCalRegle.Detail.count) and (TobCalRegle.Detail[i].GetValue('ACG_DATEEFFET')=dDerniereDateEffet) do
                begin
                sArgBlocAn := TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANBLOC');
                sArgAlertAn := TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANALERT');
                dNbHMaxSmnLim1 := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMNLIM1');
                iNbSemLim1 := TobCalRegle.Detail[i].GetValue('ACG_NBSEMLIM1');
                bSemConsecLim1 := false;
                if (TobCalRegle.Detail[i].GetValue('ACG_SEMCONSECLIM1')='X') then
                    bSemConsecLim1 := true;
                dNbHMaxSmnLim2 := TobCalRegle.Detail[i].GetValue('ACG_NBHMAXSMNLIM2');
                iNbSemLim2 := TobCalRegle.Detail[i].GetValue('ACG_NBSEMLIM2');
                bSemConsecLim2 := false;
                if (TobCalRegle.Detail[i].GetValue('ACG_SEMCONSECLIM2')='X') then
                    bSemConsecLim2 := true;
                if (dSommeAn>TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANBLOC')) then
                        // Si le nombre d'heures saisies dans l'année est supérieur à la limite de blocage
                        begin
                        tEnsBloc35h := tEnsBloc35h + [tbhAnBloc];
                        Reponse:=-1;
                        end;
                if (dSommeAn>TobCalRegle.Detail[i].GetValue('ACG_NBMAXHANALERT')) then
                        // Si le nombre d'heures saisies dans l'année est supérieur à la limite d'alerte
                        begin
                        tEnsBloc35h := tEnsBloc35h + [tbhAnAlert];
                        Reponse:=-1;
                        end;
                Inc(i);
                end; // while
            end; // if (i<TobCalRegle.Detail.count)
        end; // if (TobCalRegle<>nil)
    end; // if (Index<>-1) and (Index<>-2)

//
// Test sur les semaines glissantes
//
if ((dNbHMaxSmnLim1<>0) and (iNbSemLim1<>0)) or ((dNbHMaxSmnLim2<>0) and (iNbSemLim2<>0)) then
    tEnsBloc35h := tEnsBloc35h + ControleSemaineGlissantes( CleAct.Ressource, inttostr(CleAct.Folio), false,
                                                            iNumSemErreur1, iNumSemErreur2,
                                                            dNbHMaxSmnLim1, iNbSemLim1, bSemConsecLim1,
                                                            dNbHMaxSmnLim2, iNbSemLim2, bSemConsecLim2,
                                                            dNbHDepaceLim1,dNbHDepaceLim2 ) ;


//
// Gestion des messages d'erreur
//
//if (dDateSaisie>=GetParamSoc('SO_AFDATEDEB35')) and (dDateSaisie<=GetParamSoc('SO_AFDATEFIN35')) then
sMessage:='Intervalle d''analyse du '+datetostr(GetParamSoc('SO_AFDATEDEB35'))+ ' au '+datetostr(GetParamSoc('SO_AFDATEFIN35'))+chr(10);
sMessage:=sMessage + 'Règle en date d''effet du '+datetostr(dDerniereDateEffet)+' sur le calendrier ' + RechDom('AFTSTANDCALEN', TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_champs.GetValue('ARS_STANDCALEN'),false) + chr(10) + chr(10);
sMessage:=sMessage + 'Saisie annuelle de la ressource '+TAFO_Ressource(AFOAssistants.GetRessource(Index)).tob_champs.GetValue('ARS_LIBELLE') + ' : ' + floattostr(dSommeAn)+ ' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE) + Chr(10);

if (sArgAlertAn<>'') then
    begin
    sMessage:=sMessage+'La durée d''alerte autorisée par an : '+ sArgAlertAn + ' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
    if (tbhAnAlert in tEnsBloc35h) then
        sMessage:=sMessage+ ' est dépassée.'
    else
        sMessage:=sMessage+ ' n''est pas atteinte.';
    sMessage:=sMessage+chr(10);
    end;
if (sArgBlocAn<>'') then
    begin
    sMessage:=sMessage+'La durée de blocage autorisée par an : '+ sArgBlocAn + ' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
    if (tbhAnBloc in tEnsBloc35h) then
        sMessage:=sMessage+ ' est dépassée.'
    else
        sMessage:=sMessage+ ' n''est pas atteinte.';
    sMessage:=sMessage+chr(10);
    end;

sMessage:=sMessage+chr(10);
sMessage:=sMessage + 'Controle de la durée hebdomadaire limite 1 : '+chr(10);
if (tbhSemLim1 in tEnsBloc35h) then
    begin
    sMessage:=sMessage + 'la limite de '+ inttostr(iNbSemLim1) +' semaines à plus de '+ floattostr(dNbHMaxSmnLim1)+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE)+ ' est dépassée';
    sMessage:=sMessage + ' : '+inttostr(iNumSemErreur1)+' semaines dépassent la limite';
    end
else
if (tbhSemConsLim1 in tEnsBloc35h) then
    begin
    sMessage:=sMessage + 'la limite de '+ inttostr(iNbSemLim1) +' semaines consécutives à plus de '+ floattostr(dNbHMaxSmnLim1)+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE)+ ' en moyenne est dépassée';
    sMessage:=sMessage + ' : à partir de la semaine '+inttostr(iNumSemErreur1)+' la moyenne sur les '+ inttostr(iNbSemLim1) +' semaines suivantes est de '+ floattostr(arrondi(dNbHDepaceLim1,V_PGI.OkDecQ))+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
    end
else
    sMessage:=sMessage + 'les limites ne sont pas atteintes.';

sMessage:=sMessage+chr(10);
sMessage:=sMessage + 'Controle de la durée hebdomadaire limite 2 : '+chr(10);
if (tbhSemLim2 in tEnsBloc35h) then
    begin
    sMessage:=sMessage + 'la limite de '+ inttostr(iNbSemLim2) +' semaines à plus de '+ floattostr(dNbHMaxSmnLim2)+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE)+ ' est dépassée';
    sMessage:=sMessage + ' : '+inttostr(iNumSemErreur2)+' semaines dépassent la limite';
    end
else
if (tbhSemConsLim2 in tEnsBloc35h) then
    begin
    sMessage:=sMessage + 'la limite de '+ inttostr(iNbSemLim2) +' semaines consécutives à plus de '+ floattostr(dNbHMaxSmnLim2)+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE)+ ' en moyenne est dépassée';
    sMessage:=sMessage + ' : à partir de la semaine '+inttostr(iNumSemErreur2)+' la moyenne sur les '+ inttostr(iNbSemLim2) +' semaines suivantes est de '+ floattostr(arrondi(dNbHDepaceLim2,V_PGI.OkDecQ))+' '+ RechDom('GCQUALUNITTEMPS',VH_GC.AFMESUREACTIVITE,FALSE);
    end
else
    sMessage:=sMessage + 'les limites ne sont pas atteintes.';

if (Reponse<>0) then
    begin

    end;

PGIInfoAf(sMessage, Caption);
end;

procedure TFActivite.GSSelectCell (Sender : TObject; ACol, ARow : Integer; var CanSelect : Boolean);
  var
  TOBT : TOB;
begin
  CanSelect := true;
  TOBT := GetTOBLigne (ARow);
  if (TOBT <> nil) and Not (GoRowSelect in GS.Options) then
    // si pas de Blocage complet sur le grid du aux dates,
    // on regarde s'il y a un blocage du à autre chose pour la ligne
    begin
      if BlocageLigne (TOBT) then
        begin
          CanSelect := false;
        end;
    end;
end;

procedure TFActivite.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFActivite.BZoomLigneAchClick(Sender: TObject);
var znum : string;
    CleDoc : R_CleDoc ;
    DD : Tdatetime;
    StD ,arg : String;
BEGIN
  if not BZoomLigneAch.Enabled  then exit;

  znum  := GetChampLigne('ACT_NUMPIECEACH',GS.Row);
  DecodeRefPiece(znum,Cledoc) ;

  DD:=CleDoc.DatePiece ; StD:=FormatDateTime('dd/mm/yyyy',DD) ;
  znum :=CleDoc.NaturePiece+';'+StD+';'+ CleDoc.Souche +';'
  +IntToStr(CleDoc.NumeroPiece)+';'+IntToStr(CleDoc.Indice)+';;';


  Arg := 'CONSULTATION;NUMERO:'+IntToStr(CleDoc.NumeroPiece)+';NATURE:'+CleDoc.NaturePiece+';';
  AGLLanceFiche('BTP','BTLIGNEACH_MUL','GPP_VENTEACHAT=ACH','',Arg) ;
END ;

procedure TFActivite.BZoomPieceClick(Sender: TObject);
var znum : string;
    CleDoc : R_CleDoc ;
    DD : Tdatetime;
    StD : String;
BEGIN
	if not BZoomPiece.Enabled  then exit;
{
if (bConsult.Down=true) then
    AppelPiece([ GetChampLigne('ACT_NUMPIECE',GS.Row), 'ACTION=CONSULTATION'],2)
else
    AppelPiece([ GetChampLigne('ACT_NUMPIECE',GS.Row), 'ACTION=MODIFICATION'],2);
}
// ajout gm 14/3/2002
    znum  := GetChampLigne('ACT_NUMPIECE',GS.Row);
    DecodeRefPiece(znum,Cledoc) ;

    DD:=CleDoc.DatePiece ; StD:=FormatDateTime('dd/mm/yyyy',DD) ;
    znum :=CleDoc.NaturePiece+';'+StD+';'+ CleDoc.Souche +';'
        +IntToStr(CleDoc.NumeroPiece)+';'+IntToStr(CleDoc.Indice)+';;';

   AppelPiece([ znum , 'ACTION=CONSULTATION'],2);

END ;
procedure TFActivite.BZoomPieceAchClick(Sender: TObject);
var znum : string;
    CleDoc : R_CleDoc ;
    DD : Tdatetime;
    StD : String;
BEGIN
	if not BZoomPieceAch.Enabled  then exit;
{
if (bConsult.Down=true) then
    AppelPiece([ GetChampLigne('ACT_NUMPIECE',GS.Row), 'ACTION=CONSULTATION'],2)
else
    AppelPiece([ GetChampLigne('ACT_NUMPIECE',GS.Row), 'ACTION=MODIFICATION'],2);
}
// ajout gm 14/3/2002
    znum  := GetChampLigne('ACT_NUMPIECEACH',GS.Row);
    DecodeRefPiece(znum,Cledoc) ;

    DD:=CleDoc.DatePiece ; StD:=FormatDateTime('dd/mm/yyyy',DD) ;
    znum :=CleDoc.NaturePiece+';'+StD+';'+ CleDoc.Souche +';'
        +IntToStr(CleDoc.NumeroPiece)+';'+IntToStr(CleDoc.Indice)+';;';

   AppelPiece([ znum , 'ACTION=CONSULTATION'],2);

END ;


function TFActivite.ProchaineColonneVide (AciColDepart : integer) : integer;
  var
  i : integer;
begin
  Result := SG_Jour;
  i := AciColDepart;
  while ((GS.Cells[i, GS.Row] <> '') or (GS.ColLengths[i] = -1)
        or (GS.ColWidths[i] = 0)) and (i < GS.ColCount - 1) do
    Inc (i);

  if (i < GS.ColCount - 1) then
    Result := i
  else
  if (i = GS.ColCount - 1) then
    Result := AciColDepart;
end;





procedure TFActivite.ACT_DATECREATIONChange(Sender: TObject);
begin
DateCreat.caption:=ACT_DATECREATION.Text;
end;

procedure TFActivite.ACT_DATEMODIFChange(Sender: TObject);
begin
DateModif.Caption:=ACT_DATEMODIF.Text;
end;

procedure TFActivite.ACT_CREATEURChange(Sender: TObject);
begin
Createur.Caption:=ACT_CREATEUR.Text;
end;

procedure TFActivite.ACT_UTILISATEURChange(Sender: TObject);
begin
Utilisateur.Caption:=ACT_UTILISATEUR.Text;
end;





procedure TFActivite.bPlanningClick(Sender: TObject);
  var
  arg : string;
  TypeArticle, DateDeb, DateFin : string;
begin

  if (CleAct.TypeSaisie = tsaRess) then
    arg := 'RESSOURCE=' + ACT_RESSOURCE.Text
  else
    arg := 'AFFAIRE=' + ACT_AFFAIRE.Text;


  CriteresAvantAppel (TypeArticle, DateDeb, DateFin);

  if (TypeArticle <> '') then
    arg := arg + ';TYPEARTICLE=' + TypeArticle;

  arg := arg + ';FOLIO=' + inttostr (CleAct.Folio);
  arg := arg + ';DATEACTIVITE=' + DateDeb + DateFin;
  arg := arg + ';IMMEDIAT';

  AFLanceFiche_Suivi_Activite (arg);

end;

procedure TFActivite.GereArtsLies (ARow : integer; Var ColRetour : integer);
var RefUnique, CodeArticle, LesRefsLies, StLoc, QteRef : string;
  TOBL, TOBX, TOBC, TOBArt : TOB;
  //Qte : Double;
  Cancel : boolean;
  ii, NbLies : integer;
begin
  // Conditions d'execution
  if (Action = taConsult) then Exit;
  if (CleAct.TypeAcces <> tacGlobal) then Exit;
  if not VH_GC.GCArticlesLies then Exit;

  GS.ClearSelected;
  bSelectAll.Down := false;

  //
  TOBL := GetTOBLigne (ARow);
  if TOBL = nil then Exit;
  if TOBL.GetValue ('ARTSLIES') <> 'X' then Exit;
  TOBL.PutValue ('ARTSLIES', '-');
  //Qte := TOBL.GetValue ('ACT_QTE');
  CodeArticle := Trim (Copy (TOBL.GetValue ('ACT_ARTICLE'), 1, 18));
  LesRefsLies := SelectArticlesLies (CodeArticle);
  if LesRefsLies = '' then Exit;
  StLoc := LesRefsLies;
  NbLies := 0;


  TOBX := TOB.Create ('Les activites crees', nil, -1);
  try
  SourisSablier;

  {Création des lignes liées}
  repeat
    RefUnique := ReadTokenSt (LesRefsLies);
    QteRef := ReadTokenSt (LesRefsLies);
    if (RefUnique <> '') then
      begin
        TOBArt := TOBArticles.FindFirst (['GA_ARTICLE'], [RefUnique], False);
        if (TOBArt = nil) then
          begin
            TOBArt := CreerTOBArt (TOBArticles);
            TrouverArticleSQL_GI (true, RefUnique, TOBArt, '');
          end;

        if (TOBArt <> nil) then
          begin
            TOBC := TOB.Create ('', TOBX, -1);
            TOBC.Dupliquer (TOBL, True, True, False);

            // on copie l'article récupéré dans la ligne d'activité en forcant la mise à jour des prix
            CopieArtDansLaLigne (TOBArt, TOBC, true);

            // On réinitialise les quantités à 0 pour forcer l'utilisateur à revalider les lignes
            // et ne pas afficher n'importe quoi
            TOBC.PutValue ('ACT_QTE', 0);
            TOBC.PutValue ('ACT_QTEFAC', 0);
            TOBC.PutValue ('ACT_TOTPRCHARGE', 0);
            TOBC.PutValue ('ACT_TOTPR', 0);
            TOBC.PutValue ('ACT_TOTPRCHINDI', 0);
            TOBC.PutValue ('ACT_TOTVENTE', 0);
            TOBC.PutValue ('ACT_TOTVENTEDEV', 0);

          end;

        Inc (NbLies);
      end;

  {copie des lignes générées dans l'activité}
  until ((RefUnique = '') or (LesRefsLies = ''));


  if (NbLies = 0) then exit;

  GSRowEnter (Nil, GS.Row + 1, Cancel, False);
  if (Cancel = true) then exit;

  for ii := 0 to TOBX.Detail.Count - 1 do
    begin
      TOBC := TOBX.Detail[ii];
      ClickInsert (GS.Row, true);
      TOBL := GetTOBLigne (GS.Row);
      if (TOBL <> nil) then
        begin
          TOBL.Dupliquer (TOBC, True, True);
//          TOBL.SetAllModifie (true);
          // Création du champ d'indication de ligne validée ou non, car créée à la volée
          TOBL.AddChampSupValeur ('VALIDATION', '-', false);

          // Afficher la ligne dans le grid
          AfficheLaTOBLigne (GS.Row);
        end;
    end;

  ColRetour := SG_Qte;

  finally
    GS.MontreEdit;
    GS.SynEnabled := True;
    GS.InplaceEditor.Deselect;
    SourisNormale;
    TOBX.Free;
  end;

end;

procedure TFActivite.TraiteLesLies (ARow : integer);
var CodeArticle: string;
  TOBL: TOB;
  OkL: boolean;
begin
  if (Action = taConsult) then Exit;
  if (CleAct.TypeAcces <> tacGlobal) then Exit;
  if not VH_GC.GCArticlesLies then Exit;

  TOBL := GetTOBLigne (ARow);
  if TOBL = nil then Exit;

  if TOBL.GetValue('ARTSLIES') = '-' then Exit;
  CodeArticle := TOBL.GetValue('ACT_CODEARTICLE');
  if CodeArticle = '' then Exit;
  OkL := ExisteSQL('SELECT GAL_ARTICLE FROM ARTICLELIE WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="' + CodeArticle + '"');
  if OkL then TOBL.PutValue('ARTSLIES', 'X') else TOBL.PutValue('ARTSLIES', '-');
end;

procedure TFActivite.GSDblClick (Sender : TObject);
begin
  GSElipsisClick (Sender);
end;

Initialization
RegisterAglFunc('AFCreerActivite',False,3,AGLAFCreerActivite) ;

end.



unit AssistParamSocBtp;

interface

uses Windows,
     Messages,
     SysUtils,
     Classes,
     Graphics,
     Controls,
     Forms,
     Dialogs,
  	 assist,
     HSysMenu,
     hmsgbox,
     StdCtrls,
     ComCtrls,
     ExtCtrls,
     HPanel,
     HTB97,
{$IFDEF EAGLCLIENT}
     eFiche,
     MaineAGL,
     utileAGL,
     eFichlist,
     Spin,
     etablette,
{$ELSE}
     tablette,
     db,
     FE_Main,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     DbCtrls,
     HDB,
     Fiche,
     FichList,
     AglIsoflex,
     dbgrids,
{$ENDIF}
     Hent1,
     Ent1,
     AglInit,
     splash,
     UtilSoc,
     ParamSoc,
     UTOB,
     IniFiles,
     FichComm,
     UTomUtilisat,
     Usergrp_tom,
     UtilPgi,
     UtilLine,
     UtilActionComSx,
     OuvreExo,
     ULibExercice,
     FileCtrl,
     Hctrls,
     Menus,
     Mask,
     Spin,
     HRotLab, HRichEdt, HRichOLE, TntComCtrls, TntStdCtrls, TntExtCtrls;

type
  TFassistParamSoc = class(TFAssist)
    InfoG					: TTabSheet;
    GroupBox2			: TGroupBox;
    MenuOption		: TTabSheet;
    POPCPTA				: TPopupMenu;
    InfoGen				: TMenuItem;
    Fourchettesdecomptes1: TMenuItem;
    Comptespciaux1: TMenuItem;
    GestionTVA1		: TMenuItem;
    Monnaiedetenue1: TMenuItem;
    Divers1				: TMenuItem;
    Paramtragedelettrage1: TMenuItem;
    Modedactionsurlettrage1: TMenuItem;
    ODGetInfosTRA	: TOpenDialog;
    TFINISASSISTANT: TCheckBox;
    PTitre				: TPanel;
    ParamSoc1			: TTabSheet;
    LTitre1				: THLabel;
    Line1					: TGroupBox;
    SO_SOCIETE		: TEdit;
    GrpAdresse		: TGroupBox;
    SO_ADRESSE1		: TEdit;
    SO_ADRESSE2		: TEdit;
    SO_ADRESSE3		: TEdit;
    SO_CODEPOSTAL	: TEdit;
    SO_VILLE			: TEdit;
    SO_DIVTERRIT	: TEdit;
    PREFERBTP			: TToolbarButton97;
    ParamEchangeCpta: TToolbarButton97;
    BPrefcompta		: TToolbarButton97;
    SO_LIBELLE		: TEdit;
    SO_NATUREJURIDIQUE: THValComboBox;
    SO_PAYS				: THValComboBox;
    GroupBox4			: TGroupBox;
    GrpCoordonnees: TGroupBox;
    SO_TELEPHONE	: TEdit;
    SO_FAX				: TEdit;
    SO_TELEX			: TEdit;
    SO_MAIL				: TEdit;
    LSO_CONTACT		: TLabel;
    SO_CONTACT		: TEdit;
    ParamSoc2			: TTabSheet;
    HLabel5				: THLabel;
    GroupBox7			: TGroupBox;
    LTXTJURIDIQUE	: TLabel;
    SO_TXTJURIDIQUE: TEdit;
    SO_NIF				: TEdit;
    SO_RVA				: TEdit;
    LRVA					: TLabel;
    SO_APE				: TEdit;
    LAPE					: TLabel;
    LSIRET				: TLabel;
    SO_SIRET			: TEdit;
    LRCS					: TLabel;
    SO_RC					: TEdit;
    LCAPITAL			: TLabel;
    SO_CAPITAL		: TEdit;
    LLOGO					: TLabel;
    SO_GCLOGO_ETATS: THCritMaskEdit;
    BRetMenu			: TToolbarButton97;
    BCoordonnees	: TToolbarButton97;
    LTitre6				: THLabel;
    EchangeComptable: TTabSheet;
    HLabel1				: THLabel;
    GroupBox8			: TGroupBox;
    LCPRDREPERTOIRE: TLabel;
    SO_CPRDREPERTOIRE: THCritMaskEdit;
    Label14				: TLabel;
    SO_FREQUENCESX: THValComboBox;
    GrpResCabinet	: TGroupBox;
    Label15				: TLabel;
    SO_CPRDNOMCLIENT: TEdit;
    Label16				: TLabel;
    SO_CPRDEMAILCLIENT: TEdit;
    BEMAIL				: TToolbarButton97;
    LCPRDSIRET		: TLabel;
    SO_CPRDSIRET	: TEdit;
    LCPRDDATERECEPTION: TLabel;
    SO_CPRDDATERECEPTION: TEdit;
    SO_CPMODESYNCHRO	: TCheckBox;
    LCPPOINTAGESX			: TLabel;
    SO_CPPOINTAGESX		: THValComboBox;
    SauveRestaure			: TTabSheet;
    HLabel2						: THLabel;
    GroupBox6					: TGroupBox;
    LREPSAUVE					: TLabel;
    SO_FICHIERSAUVE		: THCritMaskEdit;
    SO_SAUVEAUTO			: TCheckBox;
    SO_ASKBEFORESAVE	: TCheckBox;
    SO_ZIPFILE: TCheckBox;
    LHOWLASTSAVE	: TLabel;
    LLASTSAVE			: TLabel;
    GroupBox9			: TGroupBox;
    SO_BTMETREDOC	: TCheckBox;
    Label9				: TLabel;
    SO_BTREPMETR	: THCritMaskEdit;
    CODEARTICLE		: TTabSheet;
    HLabel7				: THLabel;
    GroupBox10		: TGroupBox;
    GBSUFFIXE			: TGroupBox;
    GBPREFIXE					: TGroupBox;
    SO_GCPREFIXEART		: TEdit;
    SO_GCPREFIXENOM		: TEdit;
    SO_GCPREFIXEPRE		: TEdit;
    SO_BTPREFIXEPOS: TEdit;
    SO_GCTYPSUFART		: TCheckBox;
    PREFCOMPTA1				: TTabSheet;
    Line11						: TGroupBox;
    GBVENTILE					:	TGroupBox;
    SO_GCVENTCPTAART	: TCheckBox;
    SO_GCVENTCPTAAFF	: TCheckBox;
    SO_GCVENTCPTATIERS: TCheckBox;
    GBJOURNAUX		: TGroupBox;
    GbPREFAUX			: TGroupBox;
    GBLGCPT				: TGroupBox;
    HLabel8				: THLabel;
    POPTEDIT			: TPanel;
    SO_BOURREGEN	: TEdit;
    SO_BOURREAUX	: TEdit;
    PREFCOMPTA2		: TTabSheet;
    HLabel9				: THLabel;
    GroupBox13		: TGroupBox;
    GRPCOLLECTIF	: TGroupBox;
    SO_DEFCOLCLI	: THCritMaskEdit;
    SO_DEFCOLFOU	: THCritMaskEdit;
    SO_DEFCOLDDIV	: THCritMaskEdit;
    SO_DEFCOLCDIV	: THCritMaskEdit;
    PChoixLiaison	: TPanel;
    CPTAEX				: TRadioButton;
    CPTALINE			: TRadioButton;
    CPTAWINNER		: TRadioButton;
    CPTAQUADRA		: TRadioButton;
    CPTASUITE			: TRadioButton;
    NOCPTA				: TRadioButton;
    IMPORTPARAM		: TCheckBox;
    PTransfert		: TPanel;
    FileTrF				: THCritMaskEdit;
    HLabel3				: THLabel;
    BLanceRecupParam	: TToolbarButton97;
    AnimGetParam			: TAnimate;
    BoptionEtats			: TToolbarButton97;
    GRPCPTACHATVENTE	: TGroupBox;
    Label30						: TLabel;
    Label31						: TLabel;
    SO_GCCPTEESCACH		: THCritMaskEdit;
    SO_GCCPTEREMACH		: THCritMaskEdit;
    SO_GCCPTEESCVTE		: THCritMaskEdit;
    SO_GCCPTEREMVTE		: THCritMaskEdit;
    SO_GCCPTEHTACH		: THCritMaskEdit;
    SO_GCCPTEPORTACH	: THCritMaskEdit;
    GRPCPTECART				: TGroupBox;
    SO_GCECARTCREDIT	: THCritMaskEdit;
    SO_GCECARTDEBIT		: THCritMaskEdit;
    SO_GCCPTEHTVTE		: THCritMaskEdit;
    SO_GCCPTEPORTVTE	: THCritMaskEdit;
    SO_GCCPTERGVTE		: THCritMaskEdit;
    RotateLabel1			: TRotateLabel;
    RotateLabel2			: TRotateLabel;
    GPP_JOURNALCPTAA	: THValComboBox;
    GPP_JOURNALCPTAV	: THValComboBox;
    GroupBox16		: TGroupBox;
    Label37				: TLabel;
    Label39				: TLabel;
    Label40				: TLabel;
    POPBTP				: TPopupMenu;
    InfoClient		: TMenuItem;
    InfoAffaires	: TMenuItem;
    FamCptaArt		: TMenuItem;
    N1						: TMenuItem;
    VentilCpta1		: TMenuItem;
    VentilTVA			: TMenuItem;
    N2						: TMenuItem;
    LibFamArt			: TMenuItem;
    MESURE				: TMenuItem;
    ModedePaiement1		: TMenuItem;
    ConditionRegl			: TMenuItem;
    BExercices				: TToolbarButton97;
    ExoComptable			: TTabSheet;
    HLabel4						: THLabel;
    GroupBox17				: TGroupBox;
    GRPEXOENCOURS			: TGroupBox;
    HLabel6						: THLabel;
    HLabel10					: THLabel;
    EX_DATEFIN				: TMaskEdit;
    EX_DATEDEBUT			: TMaskEdit;
    EX_LIBELLE				: TEdit;
    EX_EXERCICE				: TEdit;
    EX_ABREGE					: TEdit;
    LTitre10					: THLabel;
    EX_ETATCPTA				: TEdit;
    EX_ETATBUDGET			: TEdit;
    BCreatNewExo			: TToolbarButton97;
    Benregistre				: TToolbarButton97;
    PRESENTDOC1				: TTabSheet;
    //
    GB_SOUSDETAIL			: TGroupBox;
    GB_PERSO					: TGroupBox;
    GroupBox18				: TGroupBox;
    Ligne10						: TGroupBox;
    GB_Lignes					: TGroupBox;
    GB_DESCRIPTIF			: TGroupBox;
    GBPARAMDOC				: TGroupBox;
    GB_OPTEDIT				: TGroupBox;
    //
    RB_SDESC					: TRadioButton;
    RB_IDESC					: TRadioButton;
    RB_TDESC					: TRadioButton;
    RB_SSD1						: TRadioButton;
    RB_SSD2						: TRadioButton;
    RB_SSD3						: TRadioButton;
    //
    CB_CODE						: TCheckBox;
    CB_LIBELLE				: TCheckBox;
    CB_QTE						: TCheckBox;
    CB_UNITE					: TCheckBox;
    CB_PU							: TCheckBox;
    CB_MONTANT				: TCheckBox;
    CB_CODE1					: TCheckBox;
    CB_LIBELLE1				: TCheckBox;
    CB_QTE1						: TCheckBox;
    CB_UNITE1					: TCheckBox;
    CB_PU1						: TCheckBox;
    CB_MONTANT1				: TCheckBox;
    C_DESCREMPLACE		: TCheckBox;
    C_SAUTTXTDEB			: TCheckBox;
    C_SAUTTXTFIN			: TCheckBox;
    C_IMPTOTPAR				: TCheckBox;
    C_IMPTOTSSP				: TCheckBox;
    C_IMPRECPAR				: TCheckBox;
    C_IMPMETRE				: TCheckBox;
    C_IMPRECSIT				: TCheckBox;
    C_TYPBLOCNOTE			: TCheckBox;
    C_IMPTABTOTSIT		: TCheckBox;
    C_IMPRVIATOB			: TCheckBox;
    C_IMPCOLONNES			: TCheckBox;
    C_IMPBASPAGE			: TCheckBox;
    //
    HautDePage				: TToolbarButton97;
    BasDePage					: TToolbarButton97;
    BTYPEDOC					: TToolbarButton97;
    //
    BPD_IMPCOLONNES		: TEdit;
    BPD_IMPDESCRIPTIF	: TEdit;
    BPD_IMPRECSIT			: TEdit;
    BPD_IMPRECPAR			: TEdit;
    BPD_SAUTAPRTXTDEB	: TEdit;
    BPD_SAUTAVTTXTFIN	: TEdit;
    BPD_IMPTOTPAR			: TEdit;
    BPD_IMPTOTSSP			: TEdit;
    BPD_NATUREPIECE		: TEdit;
    BPD_SOUCHE				: TEdit;
    BPD_DESCREMPLACE	: TEdit;
    BPD_IMPRVIATOB		: TEdit;
    BPD_TYPBLOCNOTE		: TEdit;
    BPD_IMPTABTOTSIT	:	TEdit;
    BPD_IMPMETRE			: TEdit;
    BPD_IMPBASPAGE		: TEdit;
    //
    BPD_TYPEPRES			: TSpinEdit;
    BPD_TYPESSD				: TSpinEdit;
    BPD_NUMPIECE			: TSpinEdit;
    GPP_NBEXEMPLAIRE	: TSpinEdit;
    //
    LGCPREFIXEART			: TLabel;
    LGCPREFIXENOM			: TLabel;
    LGCPREFIXEPRE			: TLabel;
    LGCPREFIXEPOS			: TLabel;
    Label1						: TLabel;
    Label2						: TLabel;
    Label3						: TLabel;
    Label4						: TLabel;
    Label5						: TLabel;
    Label6						: TLabel;
    Label7						: TLabel;
    Label8						: TLabel;
    Label10						: TLabel;
    Label11						: TLabel;
    Label12						: TLabel;
    Label13						: TLabel;
    Label17						: TLabel;
    Label18						: TLabel;
    Label19						: TLabel;
    Label20						: TLabel;
    Label21						: TLabel;
    Label22						: TLabel;
    Label23						: TLabel;
    Label26						: TLabel;
    Label27						: TLabel;
    Label28						: TLabel;
    Label29						: TLabel;
    Label32						: TLabel;
    Label33						: TLabel;
    Label34						: TLabel;
    Label35						: TLabel;
    Label36						: TLabel;
    Label38						: TLabel;
    Label41						: TLabel;
    Label42						: TLabel;
    Label43						: TLabel;
    Label44						: TLabel;
    Label45						: TLabel;
    Label46						: TLabel;
    Label47						: TLabel;
    Label48						: TLabel;
    Label49						: TLabel;
    Label50						: TLabel;
    Label51						: TLabel;
    Label52						: TLabel;
    Label53						: TLabel;
    Label54						: TLabel;
    Label55						: TLabel;
    //
    SO_BTECARTPMA			: THCritMaskEdit;
    SO_CPCODEEUROS3		: THCritMaskEdit;
    HautBas						: THRichEditOLE;
    //
    GPP_IMPIMMEDIATE	: TCheckBox;
    GPP_APERCUAVIMP		: TCheckBox;
    GPP_IMPMODELE			: THValComboBox;
    GPP_IMPETAT				: THValComboBox;
    //
    _ETATSDISPO				: THMultiValComboBox;
    SO_EXPJRNX				: THMultiValComboBox;
    SO_MBOCPTAEXCLPA1	: THMultiValComboBox;
    SO_MBOCPTAEXCLPA2	: THMultiValComboBox;
    MemoLog						: TRichEdit;
    Label56: TLabel;
    SO_BTMOINTERNE: THMultiValComboBox;
    SO_BTREPLPOURCENT: THCritMaskEdit;
    Label57: TLabel;
    ParamRessources: TTabSheet;
    HLabel11: THLabel;
    GroupBox3: TGroupBox;
    SO_AFPRESTATIONRES: THCritMaskEdit;
    Label59: TLabel;
    GroupBox5: TGroupBox;
    SO_AFRESCALCULPR: TCheckBox;
    Label58: TLabel;
    SO_AFFRAISGEN1: THCritMaskEdit;
    LblPrestationRes: TLabel;
    Label24: TLabel;
    SO_GCPREFIXEAUXI: TEdit;
    Label25: TLabel;
    SO_GCPREFIXEAUXIFOU: TEdit;
    SO_BTCOMPTAREGL: TCheckBox;
    SO_BTOPTRGCOLLTIE: TCheckBox;
    SO_BTLIENCPTAS1: THValComboBox;
    Label60: TLabel;
    Label61: TLabel;
    SO_BTMODEPAIEASS: THValComboBox;
    SO_BTGAMME: TEdit;
    SO_HOWLASTSAVE: THSpinEdit;
    SO_DECQTE: THSpinEdit;
    SO_DECPRIX: THSpinEdit;
    SO_LGCPTEGEN: THSpinEdit;
    SO_LGCPTEAUX: THSpinEdit;
    CPTDEVIS: THSpinEdit;
    CPTFACTURE: THSpinEdit;
    CPTAVOIR: THSpinEdit;
    GBCODEART00: TGroupBox;
    SO_GCNUMARTAUTO: TCheckBox;
    LGCLGNUMART: TLabel;
    SO_GCLGNUMART: THSpinEdit;
    LGCCOMPTEURART: TLabel;
    SO_GCCOMPTEURART: TEdit;
    SO_ECARTFACTURATION: THCritMaskEdit;
    TSO_ECARTFACTURATION: TLabel;
    TSO_PRESTASSOC: TLabel;
    SO_BTPRESTPRIXPOS: THCritMaskEdit;
    GBFACTUREAVANC: TGroupBox;
    Label63: TLabel;
    Label65: TLabel;
    SO_BTETATSIT: THValComboBox;
    Label62: TLabel;
    SO_BTETATSITDIR: THValComboBox;
    Label64: TLabel;
    SO_BTETATSIR: THValComboBox;
    Label66: TLabel;
    Label67: TLabel;
    SO_BTETATSIRDIR: THValComboBox;
    SO_METRESEXCEL: TCheckBox;
    FormesJuridiques: TMenuItem;
    LSO_BTCHEMINCOMSX: TLabel;
    SO_BTCHEMINCOMSX: THCritMaskEdit;
	  // --------
    //
    procedure bFinClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure FileTrFElipsisClick(Sender: TObject);
    procedure BLanceRecupParamClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    //
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    //
    procedure bCoordonneesClick(Sender: TObject);
    procedure InfoGenClick(Sender: TObject);
    procedure MnVentilCptaClick(Sender: TObject);
    procedure VentilCpta1Click(Sender: TObject);
    procedure TFINISASSISTANTClick(Sender: TObject);
    procedure InfoclientClick(Sender: TObject);
    procedure LibFamArtClick(Sender: TObject);
    procedure FamCptaArtClick(Sender: TObject);
    procedure InfoAffairesClick(Sender: TObject);
    procedure ParamEchangeCptaClick(Sender: TObject);
    procedure CPTASUITEClick(Sender: TObject);
    procedure NOCPTAClick(Sender: TObject);
    procedure CPTAEXClick(Sender: TObject);
    procedure CPTALINEClick(Sender: TObject);
    procedure CPTAWINNERClick(Sender: TObject);
    procedure CPTAQUADRAClick(Sender: TObject);
    procedure VentilTVAClick(Sender: TObject);
    procedure ModePaiementClick(Sender: TObject);
    procedure ConditionReglClick(Sender: TObject);
    procedure FormesjuridiquesClick(Sender: TObject);
    procedure BPrefcomptaClick(Sender: TObject);

    procedure MESUREClick(Sender: TObject);
    procedure ChargeEcranbyTob;
    procedure SO_GCNUMARTAUTOClick(Sender: TObject);
    procedure BRetMenuClick(Sender: TObject);
    procedure BExercicesClick(Sender: TObject);
    procedure BCreatNewExoClick(Sender: TObject);
    procedure BEMAILClick(Sender: TObject);
    procedure BoptionEtatsClick(Sender: TObject);
    procedure HautDePageClick(Sender: TObject);
    procedure BasDePageClick(Sender: TObject);
    procedure CB_CODEClick(Sender: TObject);
    procedure CB_LIBELLEClick(Sender: TObject);
    procedure CB_QTEClick(Sender: TObject);
    procedure CB_UNITEClick(Sender: TObject);
    procedure CB_PUClick(Sender: TObject);
    procedure CB_MONTANTClick(Sender: TObject);
    procedure RB_IDESCClick(Sender: TObject);
    procedure RB_SDESCClick(Sender: TObject);
    procedure RB_TDESCClick(Sender: TObject);
    procedure C_DESCREMPLACEClick(Sender: TObject);
    procedure RB_SSD1Click(Sender: TObject);
    procedure RB_SSD2Click(Sender: TObject);
    procedure RB_SSD3Click(Sender: TObject);
    procedure CB_UNITE1Click(Sender: TObject);
    procedure CB_QTE1Click(Sender: TObject);
    procedure CB_PU1Click(Sender: TObject);
    procedure CB_MONTANT1Click(Sender: TObject);
    procedure CB_LIBELLE1Click(Sender: TObject);
    procedure CB_CODE1Click(Sender: TObject);
    procedure C_TYPBLOCNOTEClick(Sender: TObject);
    procedure C_SAUTTXTFINClick(Sender: TObject);
    procedure C_SAUTTXTDEBClick(Sender: TObject);
    procedure C_IMPTOTSSPClick(Sender: TObject);
    procedure C_IMPTOTPARClick(Sender: TObject);
    procedure C_IMPTABTOTSITClick(Sender: TObject);
    procedure C_IMPRVIATOBClick(Sender: TObject);
    procedure C_IMPRECSITClick(Sender: TObject);
    procedure C_IMPRECPARClick(Sender: TObject);
    procedure C_IMPMETREClick(Sender: TObject);
    procedure C_IMPCOLONNESClick(Sender: TObject);
    procedure C_IMPBASPAGEClick(Sender: TObject);
    procedure BTYPEDOCClick(Sender: TObject);
    procedure SO_AFPRESTATIONRESExit(Sender: TObject);
    procedure SO_METRESEXCELClick(Sender: TObject);
    procedure SO_CAPITALExit(Sender: TObject);
    
  private
    { Dclarations prives }
    Init 						: boolean;
    IMPORT 					: boolean;
    //
    TOBParamSoc			: TOB;
    TOBCptPiece			: TOB;
    TOBExercice			: TOB;
    TobParPiece			: TOB;
		TobParamImp			: TOB;
    TobLiensOLE			: TOB;
    TOBSauv 				: TOB;
    //
    TheEchangeComSx : TEchangeComSx;
    //
    Code_Choix			: String;
    //
    TabIndex 				: Integer;
    TypePres				: Integer;
    Typessd					: Integer;
    //
    NaturePiece			: String;
    Souche					: String;
    NumeroPiece			: String;
    Origine_Txt			: String;
    NumEcr 					: integer;
    //
    MemoHaut				: THRichEditOLE;
    Memobas					: THRichEditOLE;
    //
    procedure PositionneSuivantCpta(Etat: boolean);
    procedure PositioneChecked;
    procedure PositionneTypeVentil;
    procedure EcritInfoSoc;
    procedure EcritInfoEts;
    procedure ChargeParams;
    procedure NettoieAvantImport;
    procedure ValideParamSocAvantImport;
    procedure SetTiersFactureHT;
    procedure ReInitCompta;
    procedure ReinitVentilTva;
    procedure ReInitCollectif;
    procedure ReInitBilan;
    procedure ReInitProBilan;
    procedure ReInitChaBilan;
    procedure ReInitExtBilan;
    procedure ReInitSpeciaux;
    procedure ReInitLGCompte;
    procedure ReinitParamGescom;

    //Gestion page suivante et prcdentes en fonction des option choisies au menu
    procedure AffichageChoixCompta;
    procedure AffichageTabRestore;
    procedure AlimenteBlocNotes;
    procedure AlimenteOptEdition;
    procedure ChargeParImpDoc(UneTob: Tob);
    procedure ChargeTobParamImp;
    procedure ChargementLogErreurImport;
    procedure CreationLienOLETOB(NumRang: Integer; Memo: THRichEditOLE);
    procedure EcritParamImp;
    procedure EnregistreLiensOle;
    procedure EnregistrementParametres;
    procedure EnregistreOptEdition(UneTob : Tob);
    procedure FormatAvantImport(FileName: String);
    procedure InitParImpDoc;
    procedure NextPageSoc;
    procedure NextPageCPTA;
    procedure NextPageEDI;
    procedure PositionneCptPieceEcran;
    procedure PositionneCompteurs;
    procedure PositionneExoEcran;
    procedure PositionneParamImpEcran;
    procedure PositionneParamPiece;
    procedure PositionneParPieceEcran;
    procedure PrevPageCPTA;
    procedure PrevPageEDI;
    //
    function VerifInfoVentilCompta: boolean;
    function VerifRepSauvegarde: Boolean;
    function VerifCompteur: boolean;
    function ExistDocument(Nature: string; Numero: integer): boolean;
    function LectureArticle(CodePrestation : String; TypeArticle : String) : String;
    //
  public
    { Dclarations publiques }
  end;

// DEFINITION DES TYPES DE PRESENTATION DES DETAILS D'OUVRAGES
const DOU_AUCUN		: integer = 0;
  		DOU_CODE		: integer = 1;
	  	DOU_LIBELLE	: integer = 2;
  		DOU_QTE			: integer = 4;
  		DOU_UNITE		: integer = 8;
			DOU_PU			: integer = 16;
			DOU_MONTANT	: integer = 32;
 			DOU_TOUS		: integer = 63;

function AppelParamSociete (Init : Boolean=false) : boolean;

implementation
uses UTOF_VideInside,UiUtil,BTTVA_TOF,BTPUTil;

{$R *.DFM}

function AppelParamSociete (Init : Boolean=false) : boolean;
	var XX : TFassistParamSoc;
begin
	XX := TFassistParamSoc.Create (application);
  XX.init := Init;

  XX.ShowModal;
  result := XX.IMPORT;
  XX.free;
end;

procedure TFassistParamSoc.RB_IDESCClick(Sender: TObject);
begin
  inherited;

	if RB_IDESC.checked then
     begin
     BPD_IMPDESCRIPTIF.text := 'I';
     C_DESCREMPLACE.visible := True;
     end;

end;

procedure TFassistParamSoc.RB_SDESCClick(Sender: TObject);
begin
  inherited;

	if RB_SDESC.checked then
     begin
     BPD_IMPDESCRIPTIF.text := 'S';
     C_DESCREMPLACE.visible := False;
     BPD_DESCREMPLACE.text  := '-';
     end;

end;

procedure TFassistParamSoc.RB_TDESCClick(Sender: TObject);
begin
  inherited;

  if RB_TDESC.checked then
     begin
     BPD_IMPDESCRIPTIF.text := 'T';
     C_DESCREMPLACE.visible := True;
     end;
end;

procedure TFassistParamSoc.C_DESCREMPLACEClick(Sender: TObject);
begin
  inherited;

	if C_DESCREMPLACE.Checked then
		 BPD_DESCREMPLACE.text := 'X'
	else
		 BPD_DESCREMPLACE.text := '-';

end;

procedure TFassistParamSoc.bFinClick(Sender: TObject);
Var UneTob : Tob;
begin
  inherited;
  // Retour  la page menu principal
  RestorePage;
  P.ActivePage := FirstPage;
  Pchange(nil);

  Code_Choix := '';

  //Controle des n de compteur
  if not verifCompteur then
	   begin
	   RestorePage;
     P.ActivePageIndex := Tabindex;
     Pchange(nil);
	   end;

  BRetMenu.Enabled := false;

  //Mise  jour des options ditions par pice
  UneTob := TobParPiece.FindFirst (['GPP_NATUREPIECEG'],[NaturePiece],false);
  EnregistreOptEdition(UneTob);

  //appel de la procedure d'enregistrement
  EnregistrementParametres;

  //Rechargement des paramtres aprs modification
  ChargeSocieteHalley;
  //
	TobParPiece.Free;
  TOBCptPiece.Free;
  TOBExercice.Free;
  TobParamImp.Free;
  TobLiensOLE.Free;
  //
	close;

end;


procedure TFassistParamSoc.EnregistrementParametres;
Var Indice	: integer;
    UneTob	: TOB;
begin

	V_PGI.ioerror := OeOK;

  //Ecriture des diffrents onglets lis  la table paramtres !!!
  lAide.caption := 'Mise  jour de la table des paramètres';
  for indice := 0 to  P.PageCount-1 do
      Begin
			lAide.caption := 'Traitement de l''onglet ' +  P.Pages[Indice].Caption;
		  DescentStockeUnNiveau(P.Pages[Indice], TobParamSoc);
      end;

  //UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTCOMPTAREGL'],true);
  //if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','-');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIVCHANTIER'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','-');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTAXE3'],true);
  if UneTOB <> nil then UNeTOB.putValue('SOC_DATA','-');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTAXE2'],true);
  if UneTOB <> nil then UNeTOB.putValue('SOC_DATA','-');

  if not TOBParamSoc.UpdateDB (true) then V_PGI.ioerror := OeUnknown;

  //Ecriture des infos socit dans la table Socit
  lAide.caption := 'Ecriture des infos société dans la table';
  EcritInfoSoc;

  //Ecriture des infos Socit dans la table Etablissement
  lAide.caption := 'Mise  jour de la table Etablissements';
  EcritInfoEts;

  //Ecriture des compteurs N de Pice dans table...
  lAide.caption := 'Mise à jour de la table des Compteurs Automatiques';
  PositionneCompteurs;

	//Ecriture des compteurs N de Pice dans table...
  lAide.caption := 'Mise à jour de la table des Paramètres de Pièces';
  PositionneParamPiece;

  //Ecriture des Paramtres d'impression
  lAide.caption := 'Mise à jour de la table des Paramètres d''impression';
  EcritParamImp;

  //Mise  jour de la table des liens OLE pour haut et bas de page
  lAide.caption := 'Mise à jour des texte de haut et bas de page';
  TOBLiensOLE.SetAllModifie (true);
  if not TOBLiensOLE.InsertOrUpdateDB(true) then V_PGI.ioerror := OeUnknown;

  if V_PGI.Ioerror <> OeOk then
     PGiBOx ('Erreur durant l''écriture des paramètres',Caption)
  else
     if not TFINISASSISTANT.checked then
        Begin
        UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTPARAMSOCOK'],true);
	      if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','X');
        UneTob.InsertOrUpdateDB(true)
        End;
  //
	lAide.caption := '';
end;

procedure TFassistParamSoc.bAnnulerClick(Sender: TObject);
begin
  inherited;
//
end;


procedure TFassistParamSoc.FormShow(Sender: TObject);
begin
  inherited;

  NOCPTA.Checked := true;
	AnimGetParam.visible := false;
  bfin.Visible	:= false;
  BEnregistre.Visible := True;
  PositionneSuivantCpta (false);
  //
  ParamEchangeCpta.enabled := true;
  PositioneChecked;
  //
  NaturePiece := 'DBT';
  Souche 			:= 'DBT';
  NumeroPiece := '0';
  NumEcr 			:= 0;
  //
  MemoHaut := THRichEditOLE.Create(self);
  MemoHaut.name := 'MEMOHAUT';
  MemoHaut.Parent := Self;
  Memohaut.Visible := false;

  MemoBas	 := THRichEditOLE.Create(Self);
  MemoBas.name := 'MEMOBAS';
  memoBas.parent := self;
  MemoBas.visible := false;
  //
  TOBExercice := TOB.Create ('LES EXERCICES',nil,-1);
  TOBCptPiece := TOB.Create ('LES SOUCHES',nil,-1);
  TobParPiece := TOB.Create ('LES PARPIECES',nil,-1);
  TobParamImp	:= TOB.Create ('LES IMPRESSIONS',nil,-1);
	TobLiensOLE := TOB.Create ('LES LIENS OLE',nil,-1);
  //
  ChargeEcranbyTob;

	lAide.caption := '';

end;

procedure TFassistParamSoc.FileTrFElipsisClick(Sender: TObject);
var uneTOb : TOB;
begin
  inherited;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPRDREPERTOIRE'],true);
  if UneTOB <> nil then
  	ODGetInfosTRA.InitialDir := UneTOB.GetValue('SOC_DATA')
  else
  	ODGetInfosTRA.InitialDir := 'C:\';

  if ODGetInfosTRA.Execute then
  	 begin
     if ODGetInfosTRA.FileName <> '' then
        begin
    	  FileTrf.Text := ODGetInfosTRA.FileName;
        BLanceRecupParam.enabled := true;
        end;
     end;

end;

procedure TFassistParamSoc.BLanceRecupParamClick(Sender: TObject);
var MsgErreur : String;
begin
  inherited;

  //Contrle existence fichier de rcup...
  if fileTrf.text = '' then
     Begin
     PgiInfo('La zone Fichier transfert est vide !', 'Erreur Tranfert');
     ImportParam.Checked := False;
     AffichageChoixCompta;
     Exit;
     end;

  if not FileExists(FileTrf.Text) then
     Begin
     PgiInfo('La Fichier transfert n''existe pas !', 'Erreur Tranfert');
		 FileTrFElipsisClick(Self);
     Exit;
     end;

  //Vrification si ComsX.exe existe dans le rpertoire d'execution...
  if not fileExists(ExtractFileDir (Application.ExeName)+'\COMSX.exe') then
     Begin
     PgiInfo('La Fichier ComsX.exe n''existe pas dans le rpertoire ' + ExtractFileDir (Application.ExeName) + ' !', 'Erreur installation programme');
     ImportParam.Checked := False;
		 ImportParam.Visible := False;
     AffichageChoixCompta;
     Exit;
     end;

	AnimGetParam.active := true;
	AnimGetParam.visible := true;

  //Formatage du fichier Texte pour intgration sans problme...
  LAide.caption := 'Formatage du fichier TRA avant importation';
  FormatAvantImport(FileTrf.Text);

  //
  LAide.caption := '';
  ValideParamSocAvantImport;
  //
  ReInitCompta;
  //

  //suppression des Exercices, Etablissements, Ecritures, Gnraux, Tiers, Analytique
  MsgErreur := 'Attention : l''intgration des paramtres comptables va entrainer la suppression des lments suivants :';
  MsgErreur := MsgErreur + Chr(10) + '   - Exercices';
  MsgErreur := MsgErreur + Chr(10) + '   - Etablissements';
  MsgErreur := MsgErreur + Chr(10) + '   - Ecritures';
  MsgErreur := MsgErreur + Chr(10) + '   - Comptes Gnraux';
  MsgErreur := MsgErreur + Chr(10) + '   - Comptes Tiers';
  MsgErreur := MsgErreur + Chr(10) + '   - Analytique';
  MsgErreur := MsgErreur + Chr(10) + 'Dsirez-vous continuer la rcupration des paramtres comptables?';

  if PgiAsk(MsgErreur, 'Rcupration Paramtres Comptables') = MrNo then
     Begin
     RestorePage;
     TabIndex := PrefCompta1.PageIndex;
     P.ActivePageIndex := TabIndex;
     Pchange(nil);
     exit;
     end;
  //
  NettoieAvantImport; // obligatoire pour l'importation
  //
  TheEchangeComSx.Init;
  TheEchangeComSx.ExecuteAction (TmaSxImpInit,FileTrf.Text);
  //
  IMPORT := true;
  //
	AnimGetParam.active := false;
	AnimGetParam.visible := false;

  //Affichage du log de traitement... (FihcierOK ou FichierERR)
  lAide.Caption := 'Lecture du fichier log après importation';
  ChargementLogErreurImport;

  // rechargement des paramètres mise à jour par COmsx
  lAide.Caption := 'Rechargement des paramètres';
  ChargeParams;

  //Chargement des différents écran en fonction des Tobs
  LAide.Caption := 'Réaffichage des écrans paramètres';
  ChargeEcranbyTob;

  //appel de la procedure d'enregistrement
  LAide.caption := 'Enregistrement des nouveaux paramètres';
  EnregistrementParametres;

  //Rechargement des paramètres après modification
  LAide.caption := 'Rechargement des paramètres Société';
  ChargeSocieteHalley;

  //Mise  jour en fonction des paramètres imports...
  LAide.caption := '';
  SetTiersFactureHT;

end;

Procedure TFAssistParamSoc.FormatAvantImport(FileName : String);
Var NewFile : TextFile;
    LigFile : String;
    I				: Integer;
Begin

	//Active la vèrification des Erreurs
  {$I+}

   //Ouverture et chargement du fichier Text
   AssignFile(NewFile, FileName);

   //initialisation du fichier (positionnement sur la première ligne)
   Reset(NewFile);

   //On enlve le retour chariot automatique
   MemoLog.WordWrap := False;
   MemoLog.Clear;

   //Lecture fichier TRA et modification
   while not Eof(NewFile) do
      Begin
	    Readln(NewFile, LigFile);                          { Lecture de la premire ligne du fichier }
      if copy(LigFile,4, 2) =  'S5' then StringReplace(LigFile,Copy(LigFile,4,2),'S1',[rfReplaceAll]);
      if copy(LigFile,1, 3) = '***' then MemoLog.Lines.Add(LigFile);
      end;

   //Suppression du fichier pour rcriture
   LigFile := '';
   Rewrite(NewFile);

   For i := 0 To MemoLog.Lines.Count do
       Begin
       LigFile := MemoLog.Lines[I];
       WriteLn(NewFile, LigFile); //Ecriture des lignes...
       end;

   //Fermeture du fichier
   CloseFile(NewFile);

   //Initialisation du blob
   MemoLog.WordWrap := True;
   MemoLog.Clear;

   //Dsactive la vrification des Erreurs
   {$I-}

end;

Procedure TFAssistParamSoc.ChargementLogErreurImport;
Var FichierOK 	: String;
    FichierERR	: String;
    NomFichier	: String;
    RepFichier 	: String;
    Ok_Affiche	: Boolean;
Begin

    MemoLog.Clear;
    MemoLog.PlainText := True;

    RepFichier := ExtractFiledir(FileTrf.Text) + '\';

    //chargement Fichier Erreur
    FichierERR := ChangeFileExt(FileTrf.Text, '.ERR');
    FichierERR := 'ListeCom' + ExtractFileName(FichierERR);
    FichierERR := RepFichier + FichierERR;

    //Chargement fichier Russite
    FichierOK	 := ChangeFileExt(FileTrf.Text, '.OK');
    FichierOK  := 'ListeCom' + ExtractFileName(FichierOK);
    FichierOK  := RepFichier + FichierOK;

    Ok_Affiche := False;

		//Chargement rapport d'erreur dans bloc-note...
    if FileExists(FichierERR) then
       Begin
       Nomfichier := FichierERR;
       if PGIASk('Le traitement ne s''est pas déroulé correctement.' + CHR(10) + 'Voulez-vous visualiser le fichier de rapport ?', 'Import Fichier TRA') = Mryes then
          Ok_Affiche := True;
       end;

    if FileExists(FichierOK) then
       Begin
       Nomfichier := FichierOK;
       if PGIASk('Le traitement s''est effectué correctement.' + CHR(10) + 'Voulez-vous visualiser le fichier log ?', 'Import Fichier TRA') = Mryes then
          Ok_Affiche := True;
       end;

    if Ok_Affiche then
       Begin
	     PTransfert.Visible := False;
       MemoLog.Visible := True;
       MemoLog.Lines.LoadFromFile(NomFichier);
       end;

    DeleteFile(NomFichier);

end;

procedure TFassistParamSoc.bSuivantClick(Sender: TObject);
Var NomTab : String;
begin
  inherited;

  if NumEcr > 0 then
     Begin
		 RestorePage;
     P.ActivePageindex := P.pagecount - 1;
     Pchange(nil);
     end;

  NomTab := uppercase(P.Pages[P.ActivePageindex-1].Name);
  TabIndex := P.ActivePageindex-1;

  //Branchement  partir de MenuOption
 	if CODE_CHOIX = 'SOC' then
     begin
     NextPageSoc;
     exit;
     end
  Else if CODE_CHOIX = 'CPTA' then
  	 Begin
     NextPageCPTA;
     Exit;
     end
  Else if CODE_CHOIX = 'EDI' then
     Begin
     NextPageEDI;
     Exit;
     end;

  //Gestion, du bouton suivant dans un cas normal
  if NomTab = uppercase('PARAMSOC1') then
  Else if NomTab = uppercase('PARAMSOC2') then
     Begin
     //Controle des n de compteur
     if not verifCompteur then
		 	  begin
	      RestorePage;
   	    P.ActivePageIndex := Tabindex;
        Pchange(nil);
	      end;
     exit;
     end
  Else if NomTab = uppercase('CODEARTICLE') then
  Else if NomTab = uppercase('SauveRestaure') then
     Begin
     if Not VerifRepSauvegarde then
				Begin
 	      RestorePage;
   	    P.ActivePageIndex := Tabindex;
        Pchange(nil);
        end;
     end
  Else if NomTab = Uppercase('EXOCOMPTABLE') then
  Else If NomTab = UpperCase('InfoG') then
     Begin
     //Si pas de compta on considr qu'il n'y a plus d'onglet
     if NOCPTA.Checked then
       	Begin
        BFin.Visible := True;
				BEnregistre.Visible := False;
        BSuivant.enabled := False;
        End;
     //Si importation fichier TRA est coch on change l'cran de PREFCOMPTA1
		 if IMPORTPARAM.Checked then
	       if not PTransfert.visible then	AffichageTabRestore;
     end
  Else if NomTab = uppercase('PREFCOMPTA1') then
  Else if NomTab = UpperCase('EchangeComptable') then
     Begin
     if not VerifInfoVentilCompta then
        Begin
        RestorePage;
	      TabIndex := PrefCompta2.PageIndex;
        P.ActivePageIndex := TabIndex;
        Pchange(nil);
        end;
     end;

  if P.ActivePageIndex = P.PageCount-1 then
  	 begin
     if uppercase(P.Pages[P.ActivePageindex].Name) <> 'PRESENTDOC1' then
        Begin
	      bfin.Visible	:= True;
	  		BEnregistre.Visible := False;
		    bSuivant.Enabled := false;
     		end
     else
	      NextPageEDI;
	      If NumEcr = 3 then bSuivant.Enabled := false;
     end
  else
     begin
     BRetMenu.Enabled := True;
     bfin.Visible	:= false;
	   BEnregistre.Visible := True;
     NumEcr := 0;
     end;

end;

//Gestion du bouton suivant si click sur bouton gestion entreprise !!!
Procedure TFassistParamSoc.NextPageSoc;
Var NomTab : String;
Begin

	NomTab := uppercase(P.ActivePage.Name);

	if NomTab = uppercase('SauveRestaure') then
   	 begin
     bfin.Visible := True;
	   BEnregistre.Visible := False;
		 bSuivant.Enabled := false;
     end
  else
     begin
     BRetMenu.Enabled := True;
     bfin.Visible	:= false;
	   BEnregistre.Visible := True;
     end;

end;

Procedure TFassistParamSoc.NextPageCPTA;
Var NomTab : String;
Begin

	NomTab := upperCase(P.ActivePage.Name);

  if NomTab = uppercase('PREFCOMPTA1') then
     Begin
	   if IMPORTPARAM.Checked then
        if not PTransfert.Visible then
           AffichageTabRestore;
     end
  else if NomTab = upperCase('EchangeComptable') then
   	 begin
     bfin.Visible	:= True;
		 BEnregistre.Visible := False;
  	 bSuivant.Enabled := false;
     end
  else
     begin
     BRetMenu.Enabled := True;
     bfin.Visible	:= false;
	   BEnregistre.Visible := True;
     end;

end;

Procedure TFassistParamSoc.NextPageEDI;
Var NomTab : String;
    UneTob : Tob;
Begin

	NomTab := uppercase(P.Pages[P.PageCount-1].Name);
  TabIndex := P.PageCount-1;

	ChargeTobParamImp;

  //enregistrement des texte de haut et pied de page...
	EnregistreLiensOle;

  UneTob := TobParPiece.FindFirst (['GPP_NATUREPIECEG'],[NaturePiece],false);
  EnregistreOptEdition(UneTob);

  NumEcr := NumEcr + 1;

  if NumEcr = 1 then
     NaturePiece := 'DBT'
  else if NumEcr = 2 then
     NaturePiece := 'FBT'
  else if NumEcr = 3 then
     NaturePiece := 'ABT'
  else
     Begin
     NumEcr := 3;
	   NaturePiece := 'ABT';
     end;

	Souche := NaturePiece;

  PositionneParamImpEcran;

	RestorePage;
	P.ActivePageIndex := TabIndex;
	Pchange(nil);

	if NumEcr < 3 then
     Begin
     bSuivant.Enabled := True;
	   bfin.Visible := False;
	   BEnregistre.Visible := True;
     End
  else
     begin
     bSuivant.Enabled := false;
     bfin.Visible := True;
	   BEnregistre.Visible := False;
	   end;

end;

Procedure TFassistParamSoc.ChargeTobParamImp;
var	UneTOB : TOB;
Begin

  //Chargement des informations lies  la tob PARAMIMP
	UneTOB := TOBParamImp.findFirst(['BPD_NATUREPIECE'],[NaturePiece],true);
  if UneTOB <> nil then
     Begin
     UneTOB.PutValue('BPD_NATUREPIECE', NaturePiece);
     UneTOB.PutValue('BPD_NUMPIECE', NumeroPiece);
     if TypePres < 0 then TypePres := TypePres * -1;
     UneTOB.PutValue('BPD_TYPEPRES', Typepres);
     UneTOB.PutValue('BPD_IMPTOTPAR', BPD_IMPTOTPAR.Text);
     UneTOB.PutValue('BPD_IMPTOTSSP', BPD_IMPTOTSSP.Text);
     //if Typessd < 0 then Typessd := Typessd * -1;
     UneTOB.PutValue('BPD_TYPESSD',  Typessd);
     UneTOB.PutValue('BPD_IMPCOLONNES', BPD_IMPCOLONNES.Text);
		 UneTOB.PutValue('BPD_IMPBASPAGE', BPD_IMPBASPAGE.Text);
     UneTOB.PutValue('BPD_SOUCHE', BPD_SOUCHE.Text);
     UneTOB.PutValue('BPD_TYPBLOCNOTE', BPD_TYPBLOCNOTE.Text);
     UneTOB.PutValue('BPD_SAUTAPRTXTDEB', BPD_SAUTAPRTXTDEB.Text);
     UneTOB.PutValue('BPD_SAUTAVTTXTFIN', BPD_SAUTAVTTXTFIN.Text);
     UneTOB.PutValue('BPD_IMPRECPAR', BPD_IMPRECPAR.Text);
     UneTOB.PutValue('BPD_IMPRECSIT', BPD_IMPRECSIT.Text);
     UneTOB.PutValue('BPD_IMPDESCRIPTIF', BPD_IMPDESCRIPTIF.Text);
     UneTOB.PutValue('BPD_IMPMETRE', BPD_IMPMETRE.Text);
     UneTOB.PutValue('BPD_IMPTABTOTSIT', BPD_IMPTABTOTSIT.Text);
     UneTOB.PutValue('BPD_DESCREMPLACE', BPD_DESCREMPLACE.Text);
     UneTOB.PutValue('BPD_IMPRVIATOB', BPD_IMPRVIATOB.Text);
     UneTob.SetAllModifie(True);
     end;

end;

Procedure TFassistParamSoc.PrevPageEDI;
Var NomTab  : String;
    UneTob	: Tob;
Begin

  NomTab := uppercase(P.Pages[P.PageCount-1].Name);
  TabIndex := P.PageCount-1;

  ChargeTobParamImp;

  //enregistrement des texte de haut et pied de page...
	EnregistreLiensOle;

  UneTob := TobParPiece.FindFirst (['GPP_NATUREPIECEG'],[NaturePiece],false);
  EnregistreOptEdition(UneTob);

  NumEcr := NumEcr - 1;

  if NumEcr = 1 then
     NaturePiece := 'DBT'
  else if NumEcr = 2 then
     NaturePiece := 'FBT'
  else if NumEcr = 3 then
     NaturePiece := 'ABT'
  else
     Begin
	   NumEcr := 1;
	   NaturePiece := 'DBT';
     end;

  Souche := NaturePiece;

  PositionneParamImpEcran;

	RestorePage;
	P.ActivePageIndex := TabIndex;
	Pchange(nil);

  if NumEcr = 1 then
     if CODE_CHOIX = 'EDI' then bPrecedent.Enabled := False;

  bfin.Visible := False;
  BEnregistre.Visible := True;
  BSuivant.Enabled := True;

end;

Procedure TfAssistParamSoc.AffichageTabRestore;
Begin

  //si le blocnote est affich on change d'onglet !!!!!
  if Memolog.visible then
  	 begin
     //en fin de traitement au passe  l'cran suivant de paramtrage !!!!
     Memolog.visible := false;
     TabIndex := Prefcompta1.PageIndex;
     P.ActivePageIndex := TabIndex;
     Pchange(nil);
     exit;
     end;

  PChoixLiaison.Visible := False;

  ltitre6.Caption := 'Récupration Fichier Export Comptable';

  PTransfert.Visible := true;
  MemoLog.Visible := False;

  RestorePage;
  TabIndex := InfoG.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

end;

Procedure TFAssistParamSoc.AffichageChoixCompta;
Begin

  PTransfert.Visible := False;
  MemoLog.Visible := False;

  ltitre6.Caption := 'Option de Liaison Comptable';

  PChoixLiaison.Visible := True;

  RestorePage;

  TabIndex := InfoG.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  if not fileExists(ExtractFileDir (Application.ExeName)+'\COMSX.exe') then
     Begin
     ImportParam.Checked := False;
		 ImportParam.Visible := False;
     end;       

  BPrecedent.Enabled := false;

end;

procedure TFassistParamSoc.bPrecedentClick(Sender: TObject);
Var NomTab : String;
begin
  inherited;

  if CODE_CHOIX = 'CPTA' then
  	 Begin
     PrevPageCPTA;
     Exit;
     end
  else if CODE_CHOIX = 'EDI' then
  	 Begin
     PrevPageEDI;
     Exit;
     end;

  if NumEcr > 1 then
     PrevPageEDI
  else
	   NumEcr := 0;

	NomTab := UpperCase(P.ActivePage.Name);

  if NomTab = Uppercase('Paramsoc1') then
     if CODE_CHOIX = 'SOC' then
        BPrecedent.Enabled := false
	Else if NomTab = Uppercase('EXOCOMPTABLE') then
     if getParamSoc('SO_BTLIENCPTAS1')= 'AUC' then
        Begin
        TabIndex := SauveRestaure.PageIndex;
        P.ActivePageIndex := TabIndex;
        Pchange(nil)
        end;

  if NomTab = UpperCase('MenuOption') then
     Begin
     BRetMenu.Enabled := False;
     BPrecedent.Enabled := False;
     End;

  if P.ActivePageIndex < P.PageCount-1 then
     Begin
     bfin.Visible	:= false;
	   BEnregistre.Visible := True;
		 end;

end;


Procedure TFassistParamSoc.PrevPageCPTA;
Var NomTab : String;
Begin

	NomTab := UpperCase(P.ActivePage.Name);

  if NomTab = UpperCase('ExoComptable') then
     if IMPORTPARAM.Checked then
        AffichageChoixCompta;

  if NomTab = UpperCase('InfoG') then
     if not PTransfert.Visible then
        Begin
        PChoixLiaison.visible	:= True;
        BPrecedent.Enabled		:= false;
        end;

End;

procedure TFassistParamSoc.ChargeParams;
var QQ : TQuery;
begin

  TOBParamSoc.ClearDetail;

  QQ := OpenSql ('SELECT * FROM PARAMSOC',true);
  TOBParamSoc.LoadDetailDB ('PARAMSOC','','',QQ,false);

  ferme (QQ);

  //AddParamVentil;

end;

procedure TFassistParamSoc.FormCreate(Sender: TObject);
var SPlash : TFsplashScreen;
begin
  inherited;
  IMPORT := false;
  TOBParamSoc := TOB.Create ('LES PARAMSOC',nil,-1);
  TOBSauv := TOB.Create ('LA SAUVEGARDE',nil,-1);
  splash := TFsplashScreen.Create (application);
  splash.Label1.Caption := 'Initialisation de l''assistant de création en cours...';
  splash.Show;
  splash.Refresh;
  ChargeParams;
  splash.free;
  TheEchangeComSx := TEchangeComSx.Create;
end;

procedure TFassistParamSoc.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
	TOBParamSoc.free;
  TOBSauv.free;
  TheEchangeComSx.Free;
  TheTOB:=Nil;
end;

procedure TFassistParamSoc.bCoordonneesClick(Sender: TObject);
begin
  inherited;

	TheTOB := TOBparamSoc;

  TabIndex := 0;


  CODE_CHOIX := 'SOC';

  //AglLanceFiche ('BTP','BTPARAMSOCS','','',''); //==> envoie sur ttab de 1  3

  RestorePage;

  TabIndex := Paramsoc1.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  BRetMenu.Enabled := True;
  BPrecedent.Enabled := false;

  TheTOB := nil;
  
end;


procedure TFassistParamSoc.BExercicesClick(Sender: TObject);
begin
  inherited;

	TheTOB := TOBparamSoc;

  TabIndex := 0;

  CODE_CHOIX := 'EXO';

  RestorePage;

  TabIndex := ExoComptable.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  BRetMenu.Enabled := True;
  BSuivant.Enabled := False;
  BPrecedent.Enabled := False;

  bfin.Visible	:= True;
  BEnregistre.Visible := False;


  TheTOB := nil;

end;

procedure TFassistParamSoc.BPrefcomptaClick(Sender: TObject);
begin
  inherited;

	TheTOB := TOBparamSoc;

  CODE_CHOIX := 'CPTA';

  PChoixLiaison.Visible := True;
  PTransfert.Visible := False;

  RestorePage;

  TabIndex := Infog.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  
  if not fileExists(ExtractFileDir (Application.ExeName)+'\COMSX.exe') then
     Begin
     ImportParam.Checked := False;
		 ImportParam.Visible := False;
     end;

  BPrecedent.enabled := False;
  BRetMenu.Enabled := True;

  if NOCPTA.Checked then
     Begin
     BSuivant.enabled := False;
     BPrecedent.enabled := False;
     bfin.Visible	:= True;
  	 BEnregistre.Visible := False;
     IMPORTPARAM.Visible := False;
     end;

end;

procedure TFassistParamSoc.ParamEchangeCptaClick(Sender: TObject);
begin
  inherited;

  CODE_CHOIX := 'ECH';

  RestorePage;
  TabIndex := EchangeComptable.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  BRetMenu.Enabled := True;
  BSuivant.Enabled := False;
  BPrecedent.Enabled := False;
  bfin.Visible	:= True;
  BEnregistre.Visible := False;

end;

procedure TFassistParamSoc.InfoGenClick(Sender: TObject);
begin
  inherited;
	TheTOB := TOBparamSoc;
  AglLanceFiche ('BTP','BTPARSOCCPTAGS1','','','');
  TheTOB := nil;
end;

procedure TFassistParamSoc.MnVentilCptaClick(Sender: TObject);
begin
  inherited;
	TheTOB := TOBparamSoc;
  AglLanceFiche ('BTP','BTMETHODVENTILS1','','','');
  TheTOB := nil;
  PositionneTypeVentil;
end;

procedure TFassistParamSoc.VentilCpta1Click(Sender: TObject);
begin
  inherited;
  AglLanceFiche ('GC','GCCODECPTA','','','');
end;

procedure TFassistParamSoc.TFINISASSISTANTClick(Sender: TObject);
begin
  inherited;
	if TFINISASSISTANT.checked then
     SetParamSoc ('SO_BTPARAMSOCOK','X')
  else
     SetParamSoc ('SO_BTPARAMSOCOK','-');
end;

procedure TFassistParamSoc.InfoclientClick(Sender: TObject);
begin
  inherited;
	ParamTable('GCCOMPTATIERS',taCreat,0,nil) ;
  AvertirTable ('GCCOMPTATIERS');
end;

procedure TFassistParamSoc.LibFamArtClick(Sender: TObject);
begin
  inherited;
  ParamTable('GCLIBFAMILLE',taModif,0,nil);
  AvertirTable ('GCLIBFAMILLE');
end;

procedure TFassistParamSoc.FamCptaArtClick(Sender: TObject);
begin
  inherited;
  ParamTable('GCCOMPTAARTICLE',taCreat,0,nil) ;
  AvertirTable ('GCCOMPTAARTICLE');
end;

procedure TFassistParamSoc.InfoAffairesClick(Sender: TObject);
begin
  inherited;
  ParamTable('AFCOMPTAAFFAIRE',taCreat,0,nil) ;
  AvertirTable ('AFCOMPTAAFFAIRE');
end;

procedure TFassistParamSoc.PositionneSuivantCpta (Etat : boolean);
begin
  Infoclient.Enabled := Etat;
  FamCptaArt.enabled := Etat;
  InfoAffaires.Enabled := Etat;
  //MnVentilCpta.Enabled := Etat;
  VentilCpta1.enabled := Etat;
  //VentilCptaSuite.enabled := Etat;
  if etat then PositionneTypeVentil;
end;

procedure TFassistParamSoc.CPTASUITEClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','');
  SO_BTLIENCPTAS1.Value := '';
  PositionneSuivantCpta (true);
  ParamEchangeCpta.enabled := false;
end;

procedure TFassistParamSoc.NOCPTAClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;

  if Uppercase(P.ActivePage.Name ) = Uppercase('InfoG') then
     Begin
     BSuivant.enabled := false;
     BPrecedent.enabled := false;
     IMPORTPARAM.Visible := false;
     bfin.Visible	:= True;
	   BEnregistre.Visible := False;
     BRetMenu.Enabled := true;
     end;

  if getParamSoc('SO_BTLIENCPTAS1')= 'AUC' then BExercices.Enabled := False;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','X');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','AUC');
  SO_BTLIENCPTAS1.Value := 'AUC';
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAART'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTATIERS'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAAFF'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');

  PositionneSuivantCpta (false);
  ParamEchangeCpta.enabled := false;

end;

procedure TFassistParamSoc.CPTAEXClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S5');
  SO_BTLIENCPTAS1.Value := 'S5';
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S5');
  PositionneSuivantCpta (true);
  ParamEchangeCpta.enabled := true;
end;

procedure TFassistParamSoc.CPTALINEClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;

  BSuivant.enabled := True;
  BPrecedent.enabled := false;
  IMPORTPARAM.Visible := True;

  BRetMenu.Enabled := True;

  bfin.Visible	:= false;
  BEnregistre.Visible := True;

  BExercices.Enabled := True;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S1');
  SO_BTLIENCPTAS1.Value := 'S1';
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S1');
  PositionneSuivantCpta (true);
  ParamEchangeCpta.enabled := true;
end;

procedure TFassistParamSoc.CPTAQUADRAClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;

  BSuivant.enabled := True;
  BPrecedent.enabled := false;
  IMPORTPARAM.Visible := True;

  BRetMenu.Enabled := True;

  bfin.Visible	:= false;
  BEnregistre.Visible := True;

  BExercices.Enabled := True;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCPREFIXEAUXI'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','0');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCPREFIXEAUXIFOU'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','9');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','QUA');
  SO_BTLIENCPTAS1.Value := 'QUA';

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S1');

  PositionneSuivantCpta (true);

  ParamEchangeCpta.enabled := true;

end;

procedure TFassistParamSoc.CPTAWINNERClick(Sender: TObject);
var UneTOB : TOB;
begin
  inherited;

  BSuivant.enabled := True;
  BPrecedent.enabled := false;
  IMPORTPARAM.Visible := True;

  BRetMenu.Enabled := True;

  bfin.Visible	:= false;
  BEnregistre.Visible := True;

  BExercices.Enabled := True;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','-');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','WIN');
  SO_BTLIENCPTAS1.Value := 'WIN';
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCPREFIXEAUXI'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','0');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCPREFIXEAUXIFOU'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','9');
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if UneTOB <> nil then UneTOB.putValue('SOC_DATA','S1');

  PositionneSuivantCpta (true);

  ParamEchangeCpta.enabled := true;

end;

procedure TFassistParamSoc.PositioneChecked;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);

  if UneTOB <> nil then
     begin
     if UneTOB.GetValue('SOC_DATA')='AUC' Then
        begin
        NOCPTAClick  (Self);
        NOCPTA.Checked := true;
        end
     else if UneTOB.GetValue('SOC_DATA')='S1' Then
        begin
    	  CPTALINEClick  (Self);
        CPTALINE.Checked := true;
        end
     else if UneTOB.GetValue('SOC_DATA')='WIN' Then
        begin
    	  CPTAWINNERClick  (Self);
        CPTAWINNER.Checked := true;
        end
     else if UneTOB.GetValue('SOC_DATA')='QUA' Then
        begin
    	  CPTAQUADRAClick  (Self);
        CPTAQUADRA.Checked := true;
        end
     else if (UneTOB.GetValue('SOC_DATA')='S5') then
        begin
    	  CPTAEXClick (Self);
        CPTAEX.Checked := true;
        end
     else if UneTOB.GetValue('SOC_DATA')='' then
        begin
    	  CPTASUITEClick  (Self);
        CPTASUITE.Checked := true;
        end
     else
        begin
    	  NOCPTAClick  (Self);
        NOCPTA.Checked := true;
       end;
     end
  else
     begin
  	 NOCPTAClick  (Self);
     NOCPTA.Checked := true;
     end;

end;

procedure TFassistParamSoc.PositionneTypeVentil;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAART'],true);
  if UneTOB <> nil then
  	 begin
  	 FamCptaArt.enabled := (UneTOB.GetValue('SOC_DATA')='X');
  	 end;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTATIERS'],true);
  if UneTOB <> nil then
  	 begin
  	 Infoclient.enabled := (UneTOB.GetValue('SOC_DATA')='X');
	   end;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCVENTCPTAAFF'],true);
  if UneTOB <> nil then
  	 begin
  	 InfoAffaires.enabled := (UneTOB.GetValue('SOC_DATA')='X');
  	 end;

end;

procedure TFassistParamSoc.EcritInfoSoc;
var TOBSOC,TOBINI : TOB;
		UneTOB : TOB;
    QQ : TQuery;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_SOCIETE'],true);

  if UneTOB = nil then exit;

  TOBSoc := TOB.Create ('SOCIETE',nil,-1);
  TOBINI := TOB.Create ('SOCIETE',nil,-1);
  QQ := OpenSql ('SELECT * FROM SOCIETE',True);
	TOBINI.selectDb ('',QQ);
	ferme (QQ);
  //
  TOBSOC.Dupliquer(TOBINI,false,true);
  //
  if UneTOb <> nil then TOBSOC.putValue('SO_SOCIETE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_LIBELLE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_LIBELLE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE1'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_ADRESSE1',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE2'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_ADRESSE2',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE3'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_ADRESSE3',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DIVTERRIT'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_DIVTERRIT',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CODEPOSTAL'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_CODEPOSTAL',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_VILLE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_VILLE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PAYS'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_PAYS',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_TELEPHONE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_TELEPHONE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FAX'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_FAX',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_TELEX'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_TELEX',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_MAIL'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_MAIL',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_RVA'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_RVA',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CONTACT'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_CONTACT',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_NIF'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_NIF',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_APE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_APE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_SIRET'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_SIRET',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CAPITAL'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_CAPITAL',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_RC'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_RC',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_TXTJURIDIQUE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_TXTJURIDIQUE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DECQTE'],true);
  if UneTOb <> nil then TOBSOC.putValue('SO_DECQTE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DECPRIX'],true);

  if UneTOb <> nil then TOBSOC.putValue('SO_DECPRIX',UneTOB.GetValue('SOC_DATA'));
  if not TOBINI.DeleteDB then V_PGI.Ioerror := OeUnknown;
  if not TOBSoc.InsertDB(nil) then V_PGI.IoError := oeUnknown;

  TOBSOC.free;
  TOBINI.free;
end;

procedure TFassistParamSoc.VentilTVAClick(Sender: TObject);
begin
  inherited;
  ParamTVATPF(true);
end;

procedure TFassistParamSoc.ModePaiementClick(Sender: TObject);
begin
  inherited;
	FicheModePaie_AGL('');
end;

procedure TFassistParamSoc.ConditionReglClick(Sender: TObject);
begin
  inherited;
  FicheRegle_AGL ('',False,taModif);
end;


procedure TFassistParamSoc.FormesJuridiquesClick(Sender: TObject);
begin
  inherited;
  ParamTable('ttFormeJuridique',taCreat,0,Nil) ;
  AvertirTable ('ttFormeJuridique');
end;

procedure TFassistParamSoc.EcritInfoEts;
var UneTOB 	: TOB;
    TOBEtab : Tob;
    LETAB 	: TOB;
    QQ 			: TQuery;
begin

  TOBEtab := TOB.Create ('LES ETABLISSEMENTS',nil,-1);
  QQ := OpenSql ('SELECT * FROM ETABLISS',True);
  TOBEtab.LoadDetailDB ('ETABLISS','','',QQ,false,true);
  ferme (QQ);

  if TOBEtab.detail.count = 0 then exit;

  LETAB := TOBETAB.detail[0];

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_SOCIETE'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_SOCIETE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_LIBELLE'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_LIBELLE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE1'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_ADRESSE1',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE2'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_ADRESSE2',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_ADRESSE3'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_ADRESSE3',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DIVTERRIT'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_DIVTERRIT',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CODEPOSTAL'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_CODEPOSTAL',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_VILLE'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_VILLE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PAYS'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_PAYS',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_TELEPHONE'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_TELEPHONE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FAX'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_FAX',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_TELEX'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_TELEX',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_MAIL'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_EMAIL',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_APE'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_APE',UneTOB.GetValue('SOC_DATA'));

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_SIRET'],true);
  if UneTOb <> nil then LETAB.PutValue('ET_SIRET',UneTOB.GetValue('SOC_DATA'));

  if not Letab.UpdateDB (true) then V_PGI.ioerror := OeUnknown;

  TOBEtab.free;
  
end;

procedure TFassistParamSoc.NettoieAvantImport;
begin
  ExecuteSQl ('DELETE FROM EXERCICE');
  ExecuteSQl ('DELETE FROM ETABLISS');
  ExecuteSQl ('DELETE FROM ECRITURE');
  ExecuteSQl ('DELETE FROM GENERAUX');
  ExecuteSQl ('DELETE FROM TIERS');
  ExecuteSQl ('DELETE FROM ANALYTIQ');
end;

procedure TFassistParamSoc.ValideParamSocAvantImport;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCDESACTIVECOMPTA'],true);
  if UneTOB <> nil then UneTOB.InsertOrUpdateDB ;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BTLIENCPTAS1'],true);
  if UneTOB <> nil then UneTOB.InsertOrUpdateDB ;

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CPLIENGAMME'],true);
  if UneTOB <> nil then UneTOB.InsertOrUpdateDB ;

end;

procedure TFassistParamSoc.SetTiersFactureHT;
begin

  //Mise  jour automatique des tiers client en Facture HT au lieu de TTC
  ExecuteSQL('UPDATE TIERS SET T_FACTUREHT="X" WHERE T_NATUREAUXI="CLI"');

  //Lecture des tiers pour contrle longueur de compte !!!
  //Opensql('SELECT T_COLLECTIF, T_AUXILLIAIRE FROM TIERS')

end;

procedure TFassistParamSoc.ReInitCompta;
begin

  // Sur la tob
  ReinitParamGescom;
  ReInitCollectif;
  ReInitBilan;
  ReInitProBilan;
  ReInitChaBilan;
  ReInitExtBilan;
  ReInitSpeciaux;
  ReInitLGCompte;

  // Reinitialisation des Ventilation de TVA
  ReinitVentilTva;

end;

procedure TFassistParamSoc.ReinitVentilTva;
var TOBVT,UneTOB : TOB;
    QQ :TQuery;
    indice : integer;
begin

  TOBVT := TOB.Create ('LES VENTILATIONS COMPTABLES',nil,-1);
  QQ := OpenSql ('SELECT * FROM TXCPTTVA',True);
  if not QQ.eof then TOBVT.LoadDetailDB ('TXCPTTVA','','',QQ,false);

  for Indice := 0 to TOBVT.detail.count -1 do
  	begin
    UneTOB := TOBVT.detail[Indice];
    UneTOB.putValue('TV_TAUXACH',0);
    UneTOB.putValue('TV_TAUXVTE',0);
    UneTOB.putValue('TV_CPTEACH','');
    UneTOB.putValue('TV_CPTEVTE','');
    UneTOB.putValue('TV_ENCAISACH','');
    UneTOB.putValue('TV_ENCAISVTE','');
    UneTOB.putValue('TV_FORMULETAXE','');
    UneTOB.putValue('TV_CPTVTERG','');
    UneTOB.putValue('TV_CPTACHRG','');
    UneTOB.putValue('TV_CPTACHFARFAE','');
    UneTOB.putValue('TV_CPTVTEFARFAE','');
  	end;

  TOBVT.UpdateDB;

end;

procedure TFassistParamSoc.ReInitCollectif;
var UneTOB : TOB;
begin
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLCLI'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLSAL'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLDIV'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLFOU'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLCDIV'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DEFCOLDDIV'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
end;

procedure TFassistParamSoc.ReInitBilan;
var UneTOB : TOB;
begin
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILDEB1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILDEB2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILDEB3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILDEB4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILDEB5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILFIN1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILFIN2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILFIN3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILFIN4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_BILFIN5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
end;

procedure TFassistParamSoc.ReInitProBilan;
var UneTOB : TOB;
begin
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PRODEB1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PRODEB2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PRODEB3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PRODEB4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PRODEB5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PROFIN1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PROFIN2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PROFIN3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PROFIN4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_PROFIN5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
end;

procedure TFassistParamSoc.ReInitChaBilan;
var UneTOB : TOB;
begin
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHADEB1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHADEB2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHADEB3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHADEB4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHADEB5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHAFIN1'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHAFIN2'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHAFIN3'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHAFIN4'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CHAFIN5'],true);
  if UneTOB <> nil then
  BEGIN
    UneTOB.PutValue('SOC_DATA','');
    UneTOB.UpDateDB;
  END;
end;

procedure TFassistParamSoc.ReInitExtBilan;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_EXTDEB1'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_EXTDEB2'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_EXTFIN1'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_EXTFIN2'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  TOBParamSoc.UpDateDB(True);

end;

procedure TFassistParamSoc.ReInitSpeciaux;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GENATTEND'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_CLIATTEND'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FOUATTEND'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_SALATTEND'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_DIVATTEND'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_OUVREBIL'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FERMEBIL'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_RESULTAT'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_OUVREPERTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FERMEPERTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_OUVREBEN'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_FERMEBEN'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  TOBParamSoc.UpDateDB(True);

end;

procedure TFassistParamSoc.ReInitLGCompte;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_LGCPTEGEN'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA',0);

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_LGCPTEAUX'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA',0);

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_LGMAXBUDGET'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA',0);

  TOBParamSoc.UpDateDB(True);

end;

procedure TFassistParamSoc.ReinitParamGescom;
var UneTOB : TOB;
begin

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEESCACH'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEREMACH'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEHTACH'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEPORTVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEPORTACH'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEESCVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEREMVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEHTVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTEPORTVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCCPTERGVTE'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCECARTCREDIT'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  UneTOB := TOBParamSoc.findFirst(['SOC_NOM'],['SO_GCECARTDEBIT'],true);
  if UneTOB <> nil then UneTOB.PutValue('SOC_DATA','');

  TOBParamsoc.UpDateDB(True);

end;

procedure TFassistParamSoc.MESUREClick(Sender: TObject);
begin
  inherited;
  AGLLanceFiche('GC','GCUNITEMESURE','','','');
  AvertirTable ('GCQUALUNITTOUS');
end;

//Chargement de l'cran avec les valeur de la TOB Paramtre societe
Procedure TFAssistParamSoc.ChargeEcranbyTob();
Var QQ 		 : Tquery;
    Indice : integer;
Begin

   //chargement des diffrents onglets lis  la table paramtres !!!
	 for indice := 0 to  P.PageCount-1 do
      Begin
		  DescentChargeUnNiveau(P.Pages[Indice], TobParamSoc);
      end;

   //Chargement onglet Exercices comptable
   if not PrepareOuvreExoSansCompta then
      begin
      BExercices.Enabled := false;
      ExoComptable.Visible := false;
      end;

	 LblPrestationRes.caption := LectureArticle(SO_AFPRESTATIONRES.Text, 'PRE');

   QQ := OpenSQL ('SELECT * FROM EXERCICE WHERE EX_ETATCPTA="OUV" ORDER BY EX_DATEDEBUT', True);
   TOBExercice.LoadDetailDB ('EXERCICE','','',QQ,false);
   ferme (QQ);

   if tobexercice.detail.count = 0 then TobExercice.Create('EXERCICE',nil, -1);

   PositionneExoEcran;

   //Chargement table souche...
   QQ := OpenSql ('SELECT * FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE IN ("ABT","FBT","DBT")',True);
   TOBCptPiece.LoadDetailDB ('SOUCHE','','',QQ,false);
   ferme (QQ);
   //
   PositionneCptPieceEcran;

   //Chargement Table ParPice...
   //QQ := OpenSql ('Select * FROM PARPIECE Where GPP_VENTEACHAT IN ("VEN", "ACH")',True);
   QQ := OpenSql ('Select * FROM PARPIECE Where GPP_NATUREPIECEG IN ("DBT", "FBT", "FF", "ABT")',True);
   TOBParPiece.LoadDetailDB ('PARPIECE','','',QQ,false);
   ferme (QQ);
   //
   PositionneParPieceEcran;

   //Chargement de la table des paramtres d'impression
   QQ	:= OpenSQL ('SELECT * FROM BTPARDOC WHERE BPD_NATUREPIECE IN ("DBT", "FBT", "ABT") AND BPD_NUMPIECE=0', True);
   TOBParamImp.LoadDetailDB ('BTPARDOC','','',QQ,false);
   Ferme (QQ);

   //Chargement des liens OLE
   QQ:=OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="GP" AND LO_IDENTIFIANT IN ("DBT:DBT:0:0", "FBT:FBT:0:0", "ABT:ABT:0:0")', True);
   TOBLIENSOLE.LoadDetailDB('LIENSOLE','','',QQ,False) ;
   ferme (QQ);

   //Chargement des zones necessaire  la gestion des impressions
   PositionneParamImpEcran;

   //Gestion coche ne plus afficher
   IF GetParamSocSecur('SO_BTPARAMSOCOK', False) THEN
      Begin
      TFINISASSISTANT.Checked := True;
      TFINISASSISTANT.Visible := False;
      end
   else
      Begin
      TFINISASSISTANT.Checked := False;
      TFINISASSISTANT.Visible := True;
      end;

End;

//Chargement des valeur du paramtrage d'impression des documents dans les zones cran
Procedure TFassistParamSoc.PositionneParamImpEcran;
Var UneTOB : TOB;
Begin

  //chargement des zones cran !!!
  if NaturePiece = 'DBT' then
     LTitre10.Caption := 'Préférences Document Type Devis'
  Else if NaturePiece = 'FBT' then
     LTitre10.Caption := 'Préférences Document Type Facture'
  Else if NaturePiece = 'ABT' then
     LTitre10.Caption := 'Préférences Document Type Avoir';

	UneTOB := TOBParamImp.findFirst(['BPD_NATUREPIECE'],[NaturePiece],true);
  if UneTob = nil then
     InitParImpDoc
	else
	   ChargeParImpDoc(UneTob);

  //Rinitialisation des champsmmo !!!
  MemoBas.Clear;
  MemoHaut.Clear;

  POPTEDIT.Visible := False;

  HautBas.Visible := False;
  Origine_Txt := '';

  //chargement des zones cran avec les tob adaptes
  LAide.Caption := 'chargement des options d''édition';
  AlimenteOptEdition;

  //chargement des zones cran avec les tob adaptes
  LAide.Caption := 'Chargement des haut et bas de page dans zones écran';
  AlimenteBlocNotes;

end;

Procedure TFassistParamSoc.InitParImpDoc;
begin

    Typepres := DOU_TOUS;
    BPD_IMPTOTPAR.text := 'X';
    BPD_IMPTOTSSP.text :=  'X';
    Typessd  := 0;
    BPD_IMPCOLONNES.Text := '-';
    BPD_IMPBASPAGE.Text := 'X';
    BPD_IMPMETRE.Text := 'S';
    BPD_TYPBLOCNOTE.Text := 'D';
    BPD_SAUTAPRTXTDEB.Text:= '-';
    BPD_SAUTAVTTXTFIN.text:= '-';
    BPD_IMPRECPAR.text := '-';
    BPD_IMPRECSIT.text := '-';
    BPD_IMPTABTOTSIT.text := '-';
    BPD_IMPDESCRIPTIF.text:= 'I';
    BPD_DESCREMPLACE.text := '-';
    BPD_IMPRVIATOB.text := '-';

end;

Procedure TFassistParamSoc.ChargeParImpDoc(UneTob : Tob);
Var TypeEs : integer;

Begin

   BPD_NATUREPIECE.Text := UneTob.GetString('BPD_NATUREPIECE');
   BPD_SOUCHE.Text := UneTob.GetString('BPD_SOUCHE');
   BPD_NUMPIECE.Text := UneTob.GetString('BPD_NUMPIECE');
   //
   BPD_DESCREMPLACE.Text :=  UneTob.GetString('BPD_DESCREMPLACE');
   BPD_IMPRVIATOB.Text := UneTob.GetString('BPD_IMPRVIATOB');
   BPD_TYPBLOCNOTE.Text := UneTob.GetString('BPD_TYPBLOCNOTE');
   BPD_IMPTABTOTSIT.Text := UneTob.GetString('BPD_IMPTABTOTSIT');
   BPD_IMPMETRE.Text := UneTob.GetString('BPD_IMPMETRE');
   //
   BPD_IMPCOLONNES.Text := UneTob.GetString('BPD_IMPCOLONNES');
	 BPD_IMPBASPAGE.Text := UneTob.GetString('BPD_IMPBASPAGE');
   BPD_IMPDESCRIPTIF.Text := UneTob.GetString('BPD_IMPDESCRIPTIF');
   BPD_IMPRECSIT.text := UneTob.GetString('BPD_IMPRECSIT');
   BPD_IMPRECPAR.text := UneTob.GetString('BPD_IMPRECPAR');
   BPD_SAUTAPRTXTDEB.text := UneTob.GetString('BPD_SAUTAPRTXTDEB');
   BPD_SAUTAVTTXTFIN.text := UneTob.GetString('BPD_SAUTAVTTXTFIN');
   BPD_IMPTOTPAR.text := UneTob.GetString('BPD_IMPTOTPAR');
   BPD_IMPTOTSSP.text := UneTob.GetString('BPD_IMPTOTSSP');

   CB_CODE.checked := False;
   CB_LIBELLE.Checked := False;
   CB_QTE.Checked := False;
   CB_UNITE.Checked := False;
   CB_PU.Checked := False;
   CB_MONTANT.Checked := False;
   //
   CB_CODE1.checked := False;
   CB_LIBELLE1.Checked := False;
   CB_QTE1.Checked := False;
   CB_UNITE1.Checked := False;
   CB_PU1.Checked := False;
   CB_MONTANT1.Checked := False;

   TypePres := StrToInt(UneTob.GetString('BPD_TYPEPRES'));
   TypeEs   := TypePres;
   TypePres := DOU_AUCUN;

   if (TypeEs and DOU_CODE) = DOU_CODE then CB_CODE.checked :=True;
   if (TypeEs and DOU_LIBELLE) = DOU_LIBELLE then CB_LIBELLE.Checked := True;
   if (TypeEs and DOU_QTE) = DOU_QTE then CB_QTE.Checked := True;
   if (TypeEs and DOU_UNITE) = DOU_UNITE then CB_UNITE.Checked := True;
   if (TypeEs and DOU_PU) = DOU_PU then CB_PU.Checked := True;
   if (TypeEs and DOU_MONTANT) = DOU_MONTANT then CB_MONTANT.Checked := True;

   Typessd 	:= StrToInt(UneTob.GetString('BPD_TYPESSD'));
	 TypeEs 	:= Typessd;

	 if Typessd < DOU_AUCUN then
      RB_SSD1.Checked := True
	 else if Typessd = DOU_AUCUN then
      RB_SSD2.Checked := True
	 else
      begin
      RB_SSD3.checked := True;
      GB_Perso.Visible := true;
      Typessd := DOU_AUCUN;
      if (TypeEs and DOU_CODE) = DOU_CODE then CB_CODE1.checked := True;
      if (TypeEs and DOU_LIBELLE) = DOU_LIBELLE then CB_LIBELLE1.Checked := True;
      if (TypeEs and DOU_QTE) = DOU_QTE then CB_QTE1.Checked := True;
      if (TypeEs and DOU_UNITE) = DOU_UNITE then CB_UNITE1.Checked := True;
      if (TypeEs and DOU_PU) = DOU_PU then CB_PU1.checked := True;
      if (TypeEs and DOU_MONTANT) = DOU_MONTANT then CB_MONTANT1.Checked := True;
      end;

	 C_IMPRECPAR.Checked := BPD_IMPRECPAR.text = 'X';
	 C_IMPRECSIT.Visible := False;
   C_IMPTABTOTSIT.Visible := False;
   C_IMPTOTPAR.checked := BPD_IMPTOTPAR.text = 'X';
   C_IMPTOTSSP.checked := BPD_IMPTOTSSP.text = 'X';
   C_IMPCOLONNES.checked := BPD_IMPCOLONNES.text = 'X';
	 C_IMPBASPAGE.checked := BPD_IMPBASPAGE.text = 'X';
   C_IMPMETRE.checked := BPD_IMPMETRE.text = 'T';
   C_TYPBLOCNOTE.checked := BPD_TYPBLOCNOTE.text = 'L';
   C_SAUTTXTDEB.checked := BPD_SAUTAPRTXTDEB.text = 'X';
   C_SAUTTXTFIN.checked := BPD_SAUTAVTTXTFIN.text = 'X';

	 if BPD_IMPDESCRIPTIF.text = 'S' then
  		begin
  		RB_SDESC.Checked := True;
  		C_DESCREMPLACE.Visible := False;
  		end
	 else if BPD_IMPDESCRIPTIF.text = 'T' then
		  RB_TDESC.checked := True
	 else
  		RB_IDESC.Checked := True;

	 C_DESCREMPLACE.checked := BPD_DESCREMPLACE.text = 'X';

   C_IMPRVIATOB.checked := (BPD_IMPRVIATOB.text='X');
//   C_IMPRVIATOB.Visible := false;

end;

//Chargement des valeur du compteur dans les zones ecran
procedure TFassistParamSoc.PositionneParPieceEcran;
var UneTOB : TOB;
begin

  UneTOB := TOBParPiece.findFirst(['GPP_NATUREPIECEG'],['FBT'],true);
  if UneTob <> nil then GPP_JOURNALCPTAV.Value := UneTOB.GetValue('GPP_JOURNALCPTA');

  UneTOB := TOBParPiece.findFirst(['GPP_NATUREPIECEG'],['FF'],true);
  if UneTob <> nil then GPP_JOURNALCPTAA.Value := UneTOB.GetValue('GPP_JOURNALCPTA');

end;

procedure TFassistParamSoc.PositionneCptPieceEcran;
var Indice : integer;
    UneTOB : TOB;
begin

  for Indice := 0 To TOBCptPiece.detail.count -1 do
		  begin
    	UneTOB := TOBCptPiece.detail[Indice];
	    if UneTOB.GetValue('SH_SOUCHE')='ABT' then
  	     CPTAVOIR.Value := UneTOB.GetValue('SH_NUMDEPART')
	    else if UneTOB.GetValue('SH_SOUCHE')='FBT' then
   	     CPTFACTURE.Value := UneTOB.GetValue('SH_NUMDEPART')
	    else
  	     CPTDEVIS.Value := UneTOB.GetValue('SH_NUMDEPART');
	    end;

end;

//Chargement des zone ecran dans les valeur du compteur
procedure TFassistParamSoc.PositionneCompteurs;
var Indice : integer;
    UneTOB : TOB;
begin

  for Indice := 0 To TOBCptPiece.detail.count -1 do
  		begin
    	UneTOB := TOBCptPiece.detail[Indice];
    	if UneTOB.GetValue('SH_SOUCHE')='ABT' then
      	 UneTOB.PutValue('SH_NUMDEPART',CPTAVOIR.Value)
      else if UneTOB.GetValue('SH_SOUCHE')='FBT' then
      	 UneTOB.PutValue('SH_NUMDEPART',CPTFACTURE.Value)
    	else
	       UneTOB.PutValue('SH_NUMDEPART',CPTDEVIS.Value);
  		end;

  TOBCptPiece.SetAllModifie(true);
  if not TOBCptPiece.UpdateDB (true) then V_PGI.ioerror := OeUnknown;

end;

procedure TFassistParamSoc.PositionneParamPiece;
var Indice : integer;
    UneTOB : TOB;
begin

  for Indice := 0 To TOBParPiece.detail.count -1 do
  		begin
    	UneTOB := TOBParPiece.detail[Indice];
      if UneTOB.GetValue('GPP_NATUREPIECEG')='FF' then
				 UneTOB.PutValue('GPP_JOURNALCPTA', GPP_JOURNALCPTAA.Value)
  		Else if UneTOB.GetValue('GPP_NATUREPIECEG')='DBT' then
         Begin
         UneTOB.PutValue('GPP_REFINTEXT','INT');
	       UneTOB.PutValue('GPP_TYPEPRESENT',62) ;
         end
			else if UneTOB.GetValue('GPP_NATUREPIECEG')='FBT' then
         begin
         UneTOB.PutValue('GPP_JOURNALCPTA', GPP_JOURNALCPTAV.Value);
         UneTOB.PutValue('GPP_TYPEECRCPTA','NOR');
         UneTOB.PutValue('GPP_REFINTEXT','INT');
         UneTOB.PutValue('GPP_TYPEPRESENT',62);
         end
      else if UneTOB.GetValue('GPP_NATUREPIECEG')='ABT' then
       	 begin
         UneTOB.PutValue('GPP_TYPEECRCPTA','NOR');
         UneTOB.PutValue('GPP_REFINTEXT','INT');
         UneTOB.PutValue('GPP_TYPEPRESENT',62);
       	 end;
  	  end;



  TobParPiece.SetAllModifie(true);
  if not TOBParPiece.UpdateDB (true) then V_PGI.ioerror := OeUnknown;

end;

procedure TFassistParamSoc.EcritParamImp;
begin

	ChargeTobParamImp;

  TOBPAramImp.SetAllModifie(true);
	if not TOBParamImp.InsertOrUpdateDB (true) then V_PGI.ioerror := OeUnknown;

end;

procedure TFassistParamSoc.PositionneExoEcran;
begin

  if TobExercice.Detail.Count = 0 Then
     begin
     EX_EXERCICE.Text  		:= '001';
     EX_DATEDEBUT.Text 		:= DateToStr(Idate1900);
     EX_DATEFIN.Text	 		:= DateToStr(Idate1900);
     EX_ETATCPTA.Text 		:= 'Non Ouvert';
     EX_ETATBUDGET.Text		:= 'Ouvert';
     BCreatNewExo.Visible := True;
     end
  else
     Begin
     EX_EXERCICE.Text 	:= TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_EXERCICE');
     EX_ABREGE.Text 		:= TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_ABREGE');
     EX_LIBELLE.Text 		:= TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_LIBELLE');
     EX_DATEDEBUT.Text 	:= DateToStr(TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_DATEDEBUT'));
     EX_DATEFIN.Text 		:= DateToStr(TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_DATEFIN'));
     EX_ETATCPTA.Text 	:= TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_ETATCPTA');
     EX_ETATBUDGET.Text := TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_ETATBUDGET');
     if TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_ETATCPTA') = 'OUV' then
        EX_ETATCPTA.Text := 'Ouvert'
     Else
        EX_ETATCPTA.Text := 'Non Ouvert';
     if TobExercice.Detail[TobExercice.Detail.Count -1].GetValue('EX_ETATBUDGET') = 'OUV' then
        EX_ETATBUDGET.Text := 'Ouvert'
     Else
        EX_ETATBUDGET.Text := 'Non Ouvert';
     BCreatNewExo.Visible := False;
     end;

end;

function TFassistParamSoc.ExistDocument(Nature: string; Numero: integer): boolean;
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT GP_NUMERO FROM PIECE WHERE GP_NATUREPIECEG="'+ Nature +'" AND GP_NUMERO='+ IntToStr(Numero),True);
  result := not QQ.eof;
  ferme (QQ);
end;

function TFassistParamSoc.verifCompteur: boolean;
begin

  result := false;

  if ExistDocument('DBT',CPTDEVIS.Value) then
	   begin
     PgiBox('Le Numéro de devis existe déjà');
     exit;
	   end;

  if ExistDocument('FBT',CPTFACTURE.Value) then
	   begin
     PgiBox('Le Numéro de facture existe déjà');
     exit;
 		 end;

  if ExistDocument('ABT',CPTAVOIR.Value) then
  	 begin
     PgiBox('Le Numéro d''avoir existe déjà');
     exit;
	   end;

  result := true;

end;

procedure TFassistParamSoc.SO_GCNUMARTAUTOClick(Sender: TObject);
begin
  inherited;

  GBPREFIXE.Enabled := SO_GCNUMARTAUTO.Checked;
  GBSUFFIXE.Enabled := SO_GCNUMARTAUTO.Checked;

end;

procedure TFassistParamSoc.BRetMenuClick(Sender: TObject);
begin
  inherited;

  RestorePage;
  P.ActivePage := FirstPage;
  Pchange(nil);

  Code_Choix := '';

  NumEcr := 0;

  BRetMenu.Enabled := false;

end;

procedure TFassistParamSoc.BCreatNewExoClick(Sender: TObject);
begin
  inherited;

  IF not CControleDureeExercice(StrToDate(EX_DATEDEBUT.Text), StrToDate(EX_DATEFIN.text)) then
     begin
     PGIInfo('La durée d''exercice ne doit pas excèder 24 mois.', 'Ouverture d''exercice');
     EX_DATEFIN.SetFocus;
     exit;
     end;

  if PGIASk('Confirmez-vous la cération de cet exercice ?', 'Cération Exercie') = MrNo then exit;

  if EX_LIBELLE.TEXT = '' then
     EX_LIBELLE.Text :=LibelleExerciceDefaut(StrToDate(EX_DATEDEBUT.Text), StrToDate(EX_DATEFIN.Text), False);

  if EX_ABREGE.TEXT = '' then
     EX_ABREGE.Text :=LibelleExerciceDefaut(StrToDate(EX_DATEDEBUT.Text), StrToDate(EX_DATEFIN.Text), True);

  //Maj de la Tob pour mise  jour de la base
  TobExercice.PutValue('EX_EXERCICE', EX_EXERCICE.Text);
  TobExercice.PutValue('EX_LIBELLE', EX_LIBELLE.Text);
  TobExercice.PutValue('EX_ABREGE', EX_ABREGE.Text);
	TobExercice.PutValue('EX_DATEDEBUT', StrToDate(EX_DATEDEBUT.Text));
  TobExercice.PutValue('EX_DATEFIN', StrToDate(EX_DATEFIN.Text));
  TobExercice.PutValue('EX_ETATCPTA', 'OUV');
	TobExercice.PutValue('EX_ETATBUDGET', 'OUV');
	TobExercice.PutValue('EX_NATEXO', '');
	TobExercice.PutValue('EX_ETATADV', 'NON');
	TobExercice.PutValue('EX_ETATAPPRO', 'NON');
	TobExercice.PutValue('EX_ETATPROD', 'NON');
	TobExercice.PutValue('EX_SOCIETE', V_PGI.CodeSociete);
	TobExercice.PutValue('EX_VALIDEE', '------------------------');
	TobExercice.PutValue('EX_DATECUM', idate1900);
	TobExercice.PutValue('EX_DATECUMRUB', idate1900);
	TobExercice.PutValue('EX_DATECUMBUD', idate1900);
	TobExercice.PutValue('EX_DATECUMBUDGET', idate1900);
 	TobExercice.PutValue('EX_BUDJAL', '');

  TobExercice.InsertOrUpdateDB(true);

  //---- On averti le serveur de la mise  jour des paramSoc et de la table exercice ----// SBO : 15/07/2004
  AvertirCacheServer( 'EXERCICE' ) ;

  VH^.Suivant.Code := EX_EXERCICE.Text;
  VH^.Suivant.Deb := StrToDate(EX_DATEDEBUT.text);
  VH^.Suivant.Fin := StrToDate(EX_DATEFIN.Text);

  BCreatNewExo.Visible := False;

  if Code_Choix = 'EXO' then
     BRetMenuClick(self)
  else
     Begin
		 TabIndex := InfoG.PageIndex;
		 P.ActivePageIndex := TabIndex;
		 Pchange(nil);
     end;

end;

Function TFAssistParamSoc.VerifRepSauvegarde : Boolean;
begin

  result := false;

  if (SO_FICHIERSAUVE.text = '' ) or (not DirectoryExists(SO_FICHIERSAUVE.text)) then
 	   Begin
   	 HShowMessage('3;Société;Vous devez renseigner un répertoire valide pour stocker vos sauvegardes;W;O;O;O;','','') ;
     Exit ;
     end;

	if GetParamSoc('SO_BTLIENCPTAS1')= 'AUC' then
     Begin
     AffichageChoixCompta;
     bSuivant.Enabled := True;
     end;

  result := True;

end;

function TFassistParamSoc.VerifInfoVentilCompta: boolean;
begin

	result := false;


  if Not GenExiste('SO_GCCPTEESCACH',Self,True,'le Compte escompte Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEREMACH',Self,True,'le Compte Remise Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEHTACH',Self,True,'le Compte H.T. Achat n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEPORTACH',Self,True,'le Compte Port Achat n''existe pas') then Exit ;

  if GPP_JOURNALCPTAA.Text = '' then
  	 BEGIN
  	 PGiError('Vous devez renseigner un journal de ventilation des Achats',Self.caption);
     exit;
	   END;

  if Not GenExiste('SO_GCCPTEESCVTE',self,True,'le Compte escompte Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEREMVTE',Self,True,'le Compte Remise Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEHTVTE',Self,True,'le Compte H.T. Vente n''existe pas') then Exit ;
  if Not GenExiste('SO_GCCPTEPORTVTE',Self,True,'le Compte Port Vente n''existe pas') then Exit ;

  if SO_GCECARTDEBIT.text <> '' then
  	 if Not GenExiste('SO_GCECARTDEBIT',self,True,'le Compte d''écart débiteur n''existe pas') then Exit ;

  if SO_GCECARTCREDIT.Text <> '' then
  	 if Not GenExiste('SO_GCECARTCREDIT',Self,True,'le Compte d''écart créditeur n''existe pas') then Exit ;

  if GPP_JOURNALCPTAV.text ='' then
  	 BEGIN
  	 PGiError('Vous devez renseigner un journal de ventilation des ventes','Ventilation Comptable');
     exit;
  	 END;

  result := true;

end;

procedure TFassistParamSoc.BEMAILClick(Sender: TObject);
Var ParametreMail : ParamMail;
begin
  inherited;

  ParametreMail.Email := SO_CPRDEMAILCLIENT.Text;
  ParametreMail.MessMail := 'Test d''envoi de Mail à l''adresse suivante : ' + ParametreMail.Email;
  ParametreMail.Sujet := 'Envoi Mail Cabinet comptable';
  ParametreMail.Fichier := '';
  ParametreMail.CopieA := '';

  EnvoieMail(ParametreMail);

end;

procedure TFassistParamSoc.BoptionEtatsClick(Sender: TObject);
begin
  inherited;

  NumEcr 			:= 1;

  CODE_CHOIX := 'EDI';

  NaturePiece := 'DBT';

  RestorePage;
  TabIndex := PRESENTDOC1.PageIndex;
  P.ActivePageIndex := TabIndex;
  Pchange(nil);

  BRetMenu.Enabled := True;
  BSuivant.Enabled := True;
  BPrecedent.Enabled := False;
  bfin.Visible	:= False;
  BEnregistre.Visible := True;

end;

procedure TFassistParamSoc.HautDePageClick(Sender: TObject);
Var TypeDoc : string;
begin

  inherited;

  if NaturePiece = 'DBT' then TypeDoc := 'Type Devis'
  Else if NaturePiece = 'FBT' then TypeDoc := 'Type Facture'
  Else if NaturePiece = 'ABT' then TypeDoc := 'Type Avoir';

  //Si le bloc-note n'est pas affich on charge hautbas avec le bon doc.
  if not HautBas.Visible then
     Begin
     if Origine_Txt = 'EDIT' then POPTEDIT.Visible := False;
     Ltitre10.Caption := 'Haut de Page ' + TypeDoc;
     Origine_Txt		 	:= 'HAUT';
     HautBas.Clear;
     HautBas.Hint 	 	:= 'Saisie du texte de haut de page';
     HautBas.Visible 	:= True;
     HautBas.text := MemoHaut.text;
     HautBas.SetFocus;
     exit;
     end;

  //Si bloc-note appel par bouton Bas on naffiche haut  la place de bas
  if Origine_Txt = 'BAS' then
     Begin
     Ltitre10.Caption := 'Haut de Page ' + TypeDoc;
     Origine_Txt 			:= 'HAUT';
     HautBas.Hint 	 	:= 'Saisie du texte de haut de page';
     MemoBas.Text := HautBas.Text;
     HautBas.Text := MemoHaut.Text;
     HautBas.SetFocus;
     end
  else
     Begin
     MemoHaut.text := HautBas.text;
     HautBas.Visible := False;
     Origine_Txt := '';
     if NaturePiece = 'DBT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc
     Else if NaturePiece = 'FBT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc
     Else if NaturePiece = 'ABT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc;
     end;

end;

procedure TFassistParamSoc.BasDePageClick(Sender: TObject);
Var TypeDoc : string;
begin
  inherited;

  if NaturePiece = 'DBT' then TypeDoc := 'Type Devis'
  Else if NaturePiece = 'FBT' then TypeDoc := 'Type Facture'
  Else if NaturePiece = 'ABT' then TypeDoc := 'Type Avoir';

  if Not HautBas.Visible then
     Begin
     if Origine_Txt = 'EDIT' then POPTEDIT.Visible := False;
     Ltitre10.Caption := 'Bas de Page ' + TypeDoc;
     Origine_Txt		 	:= 'BAS';
     HautBas.Clear;
     HautBas.Hint 	 	:= 'Saisie du texte de Bas de page';
     HautBas.Visible 	:= True;
     HautBas.text := MemoBas.text;
     HautBas.SetFocus;
     exit;
     end;

	if Origine_Txt = 'HAUT' then
     Begin
     Ltitre10.Caption := 'Bas de Page ' + TypeDoc;
     Origine_Txt		 	:= 'BAS';
     HautBas.Hint 	 	:= 'Saisie du texte de Bas de page';
     MemoHaut.Text := HautBas.Text;
     HautBas.Text := MemoBas.Text;
     HautBas.SetFocus;
     end
  else
     Begin
     MemoBas.text := HautBas.text;
     HautBas.Visible := False;
     Origine_Txt := '';
     if NaturePiece = 'DBT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc
     Else if NaturePiece = 'FBT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc
     Else if NaturePiece = 'ABT' then
        LTitre10.Caption := 'Préférences Document ' + TypeDoc;
     end;

end;


procedure TFassistParamSoc.AlimenteBlocNotes;
var UneTob 	: TOB;
    Piece		: String;
begin

   Piece := NaturePiece + ':' + Souche + ':0:0';

   //Chargement memo haut de page
 	 UneTob := TOBLIENSOLE.FindFirst (['LO_IDENTIFIANT','LO_RANGBLOB'],[Piece,'1'],false);
   if UneTob <> nil then StringToRich(MemoHaut, UneTob.GetValue('LO_OBJET'));

   //chargement Memo bas de page
   UneTob := TOBLIENSOLE.FindFirst (['LO_IDENTIFIANT','LO_RANGBLOB'],[Piece,'2'],false);
   if UneTob <> nil then StringToRich(MemoBas, UneTob.GetValue('LO_OBJET'));

end;

procedure TFassistParamSoc.AlimenteOptEdition;
var UneTob 	: TOB;
Begin

   //Chargement memo haut de page
 	 UneTob := TOBParPiece.FindFirst (['GPP_NATUREPIECEG'],[NaturePiece],false);
   if UneTob <> nil then
      Begin
      GPP_NBEXEMPLAIRE.Text := UneTob.Getvalue('GPP_NBEXEMPLAIRE');
      GPP_IMPIMMEDIATE.Checked := UneTob.Getvalue('GPP_IMPIMMEDIATE')='X';
      GPP_APERCUAVIMP.Checked := UneTob.Getvalue('GPP_APERCUAVIMP')='X';
      GPP_IMPMODELE.Value  := UneTob.Getvalue('GPP_IMPMODELE');
      GPP_IMPETAT.Value  := UneTob.Getvalue('GPP_IMPETAT');
      end;


End;

procedure TFassistParamSoc.EnregistreLiensOle();
var UneTob : TOB;
    Piece  : string;
begin

	piece := NaturePiece + ':' + Souche + ':0:0';

  //Suppression des enregistrements dj existant...
	//if (TOBLiensOLE.detail.count > 0) then
  //	 begin
  // 	 TOBLiensOLE.DeleteDB;
  // 	 TOBLiensOLe.ClearDetail;
  // 	 end;

 	UneTob := TOBLIENSOLE.FindFirst (['LO_IDENTIFIANT','LO_RANGBLOB'],[Piece,'1'],false);
  if UneTob <> nil then
     Begin
	   if (length (MemoHaut.Text) <> 0) and (MemoHaut.text <> #$D#$A) and (MemoHaut.text <> '') then
  	 		begin
        UneTob.PutValue('LO_TABLEBLOB', 'GP');
        UneTob.PutValue('LO_LIBELLE', 'Haut de Page Document');
        UneTob.PutValue('LO_QUALIFIANTBLOB', 'MEM');
        UneTob.PutValue('LO_IDENTIFIANT', Piece);
        UneTob.PutValue('LO_RANGBLOB', 1);
        UneTob.PutValue('LO_OBJET', ExRichToString(MemoHaut));
        end;
     end
  Else
     CreationLienOLETob(1, MemoHaut);

  UneTob := TOBLIENSOLE.FindFirst (['LO_IDENTIFIANT','LO_RANGBLOB'],[Piece,'2'],false);
  if UneTob <> nil then
     Begin
		 if (length (MemoBas.Text) <> 0) and (MemoBas.text <> #$D#$A) and (MemoBas.text <> '') then
        begin
        UneTob.PutValue('LO_TABLEBLOB', 'GP');
        UneTob.PutValue('LO_LIBELLE', 'Bas de Page Document');
        UneTob.PutValue('LO_QUALIFIANTBLOB', 'MEM');
        UneTob.PutValue('LO_IDENTIFIANT', Piece);
        UneTob.PutValue('LO_RANGBLOB', 2);
        UneTob.PutValue('LO_OBJET', ExRichToString(MemoBas));
        end;
     end
  else
     CreationLienOLETob(2, Memobas);

end;

procedure TFassistParamSoc.EnregistreOptEdition(UneTob : Tob);
begin

  if UneTob <> nil then
     Begin
     UneTob.Putvalue('GPP_NBEXEMPLAIRE',GPP_NBEXEMPLAIRE.Text);
     if GPP_IMPIMMEDIATE.Checked then
        UneTob.Putvalue('GPP_IMPIMMEDIATE','X')
     else
        UneTob.Putvalue('GPP_IMPIMMEDIATE','-');
     If GPP_APERCUAVIMP.Checked Then
        UneTob.Putvalue('GPP_APERCUAVIMP','X')
     Else
	      UneTob.Putvalue('GPP_APERCUAVIMP','-');
     UneTob.Putvalue('GPP_IMPMODELE', GPP_IMPMODELE.Value);
     UneTob.Putvalue('GPP_IMPETAT', GPP_IMPETAT.Value);
		 end;

end;

Procedure TFassistParamSoc.CreationLienOLETOB(NumRang : Integer; Memo : THRichEditOLE);
Var UneTOB : TOB;
    Piece  : String;
Begin

	piece := NaturePiece + ':' + Souche + ':0:0';

	UneTOB := Tob.Create('LIENSOLE', TOBLiensOLE, -1);

  UneTob.PutValue('LO_TABLEBLOB', 'GP');
  UneTob.PutValue('LO_LIBELLE', 'Memo Document ' + NaturePiece);
  UneTob.PutValue('LO_QUALIFIANTBLOB', 'MEM');
  UneTob.PutValue('LO_IDENTIFIANT', Piece);
  UneTob.PutValue('LO_RANGBLOB', NumRang);
  UneTob.PutValue('LO_OBJET', ExRichToString(Memo));

end;

procedure TFassistParamSoc.CB_CODEClick(Sender: TObject);
begin
  inherited;

	if CB_CODE.checked then
     Typepres := Typepres + 1
	else
     Typepres := Typepres - 1;

end;

procedure TFassistParamSoc.CB_LIBELLEClick(Sender: TObject);
begin
  inherited;

	if CB_LIBELLE.checked then
     Typepres := Typepres + 2
	else
     Typepres := Typepres - 2;

end;

procedure TFassistParamSoc.CB_QTEClick(Sender: TObject);
begin
  inherited;

	if CB_QTE.checked then
     Typepres := Typepres + 4
	else
     Typepres := Typepres - 4;

end;

procedure TFassistParamSoc.CB_UNITEClick(Sender: TObject);
begin
  inherited;

	if CB_UNITE.checked then
     Typepres := Typepres + 8
	else
     Typepres := Typepres - 8;

end;

procedure TFassistParamSoc.CB_PUClick(Sender: TObject);
begin
  inherited;

	if CB_PU.checked then
     Typepres := Typepres + 16
	else
     Typepres := Typepres - 16;

end;

procedure TFassistParamSoc.CB_MONTANTClick(Sender: TObject);
begin
  inherited;

	if CB_MONTANT.checked then
     Typepres := Typepres + 32
	else
     Typepres := Typepres - 32;

end;

procedure TFassistParamSoc.RB_SSD1Click(Sender: TObject);
begin
  inherited;

  if RB_SSD1.checked then
		 begin
		 GB_PERSO.visible  := False;
		 Typessd := -1;
		 end

end;

procedure TFassistParamSoc.RB_SSD2Click(Sender: TObject);
begin
  inherited;

  if RB_SSD2.checked then
		 begin
		 GB_PERSO.visible  := False;
		 Typessd := 0;
		 end
end;

procedure TFassistParamSoc.RB_SSD3Click(Sender: TObject);
begin
  inherited;

  if RB_SSD3.checked then
		 begin
		 GB_PERSO.visible  := True;
		 Typessd := 0;
		 end

end;

procedure TFassistParamSoc.CB_UNITE1Click(Sender: TObject);
begin
  inherited;

	if CB_UNITE1.checked then
     Typessd := Typessd + 8
	else
     Typessd := Typessd - 8;

end;

procedure TFassistParamSoc.CB_QTE1Click(Sender: TObject);
begin
  inherited;

 	if CB_QTE1.checked then
     Typessd := Typessd + 4
	else
     Typessd := Typessd - 4;

end;

procedure TFassistParamSoc.CB_PU1Click(Sender: TObject);
begin
  inherited;

	if CB_PU1.checked then
     Typessd := Typessd + 16
	else
     Typessd := Typessd - 16;

end;

procedure TFassistParamSoc.CB_MONTANT1Click(Sender: TObject);
begin
  inherited;

	if CB_MONTANT1.checked then
     Typessd := Typessd + 32
	else
     Typessd := Typessd - 32;

end;

procedure TFassistParamSoc.CB_LIBELLE1Click(Sender: TObject);
begin
  inherited;

	if CB_LIBELLE1.checked then
     Typessd := Typessd + 2
	else
     Typessd:= Typessd - 2;

end;

procedure TFassistParamSoc.CB_CODE1Click(Sender: TObject);
begin
  inherited;

	if CB_CODE1.checked then
     Typessd := Typessd + 1
	else
     Typessd := Typessd - 1;

end;

procedure TFassistParamSoc.C_TYPBLOCNOTEClick(Sender: TObject);
begin
  inherited;

	if C_TYPBLOCNOTE.Checked then
		 BPD_TYPBLOCNOTE.text := 'L'
	else
		 BPD_TYPBLOCNOTE.text := 'D';
end;

procedure TFassistParamSoc.C_SAUTTXTFINClick(Sender: TObject);
begin
  inherited;

	if C_SAUTTXTFIN.checked then
		 BPD_SAUTAVTTXTFIN.text := 'X'
	else
		 BPD_SAUTAVTTXTFIN.text := '-';
     
end;

procedure TFassistParamSoc.C_SAUTTXTDEBClick(Sender: TObject);
begin
  inherited;

	if C_SAUTTXTDEB.Checked then
	   BPD_SAUTAPRTXTDEB.text := 'X'
	else
		 BPD_SAUTAPRTXTDEB.text := '-';

end;

procedure TFassistParamSoc.C_IMPTOTSSPClick(Sender: TObject);
begin
  inherited;

	if C_IMPTOTSSP.checked then
		 BPD_IMPTOTSSP.text := 'X'
	else
		 BPD_IMPTOTSSP.text := '-';

end;

procedure TFassistParamSoc.C_IMPTOTPARClick(Sender: TObject);
begin
  inherited;

	if C_IMPTOTPAR.Checked then
		 BPD_IMPTOTPAR.text := 'X'
	else
		 BPD_IMPTOTPAR.text := '-';

end;

procedure TFassistParamSoc.C_IMPTABTOTSITClick(Sender: TObject);
begin
  inherited;

  if C_IMPTABTOTSIT.checked then
     BPD_IMPTABTOTSIT.text := 'X'
  else
     BPD_IMPTABTOTSIT.text := '-';

end;

procedure TFassistParamSoc.C_IMPRVIATOBClick(Sender: TObject);
begin
  inherited;

  if C_IMPRVIATOB.Checked then
     BPD_IMPRVIATOB.text := 'X'
  else
     BPD_IMPRVIATOB.text := '-';

end;

procedure TFassistParamSoc.C_IMPRECSITClick(Sender: TObject);
begin
  inherited;

	if C_IMPRECSIT.Checked then
		 BPD_IMPRECSIT.text := 'X'
	else
		 BPD_IMPRECSIT.text := '-';
end;

procedure TFassistParamSoc.C_IMPRECPARClick(Sender: TObject);
begin
  inherited;

  if C_IMPRECPAR.Checked then
     BPD_IMPRECPAR.text := 'X'
  else
     BPD_IMPRECPAR.text := '-';

end;

procedure TFassistParamSoc.C_IMPMETREClick(Sender: TObject);
begin
  inherited;

  if C_IMPMETRE.Checked then
     BPD_IMPMETRE.text := 'T'
  else
     BPD_IMPMETRE.text := 'S';
     
end;

procedure TFassistParamSoc.C_IMPCOLONNESClick(Sender: TObject);
begin
  inherited;

	if C_IMPCOLONNES.Checked then
		 begin
     BPD_IMPCOLONNES.text  := 'X';
     BPD_TYPBLOCNOTE.text  := 'D';
     C_TYPBLOCNOTE.visible := False;
   	 end
  else
     begin
     BPD_IMPCOLONNES.text  := '-';
     C_TYPBLOCNOTE.visible := True;
     end

end;

procedure TFassistParamSoc.BTYPEDOCClick(Sender: TObject);
Var TypeDoc : String;
begin
  inherited;

  if NaturePiece = 'DBT' then
  begin
  	TypeDoc := 'Devis';
  end Else if NaturePiece = 'FBT' then
  begin
  	TypeDoc := 'Factures';
    GBFACTUREAVANC.visible := true;
  end Else if NaturePiece = 'ABT' then
  begin
  	TypeDoc := 'Avoir';
  end;

  LTitre10.Caption := 'Préférences Document ' + TypeDoc;

  if not POPTEDIT.Visible then
     Begin
     if Origine_Txt <> 'EDIT' then HautBas.Visible := false;
	   GB_OPTEDIT.Caption := 'Edition des ' + TypeDoc;
     Origine_Txt		 	:= 'EDIT';
     POPTEDIT.Visible := True;
     GPP_NBEXEMPLAIRE.SetFocus;
     exit;
     end;

  if Origine_Txt <> 'EDIT' then
     Begin
     IF Origine_Txt = 'BAS' then
        BasDePageClick(Self)
     else if Origine_Txt = 'HAUT' then
	      HautDePageClick(Self);
     Exit;
     end
  else
     Begin
     POPTEDIT.Visible := False;
     Origine_Txt := '';
     end;


end;

procedure TFassistParamSoc.C_IMPBASPAGEClick(Sender: TObject);
begin

	if C_IMPBASPAGE.Checked then
     BPD_IMPBASPAGE.text  := 'X'
  else
     BPD_IMPBASPAGE.text  := '-';

end;

procedure TFassistParamSoc.SO_AFPRESTATIONRESExit(Sender: TObject);
begin
  inherited;

  LblPrestationRes.caption := LectureArticle(SO_AFPRESTATIONRES.Text, 'PRE');

end;

Function TFassistParamSoc.LectureArticle(CodePrestation : String; TypeArticle : String) : String;
Var QQ 	: TQuery;
begin

	Result := '';

  QQ := OpenSql('Select GA_LIBELLE FROM ARTICLE WHERE GA_CODEARTICLE="' + CodePrestation + '" AND GA_TYPEARTICLE = "PRE"',true);

  if Not QQ.Eof then Result := QQ.FindField('GA_LIBELLE').AsString ;

  ferme (QQ);

end;

procedure TFassistParamSoc.SO_METRESEXCELClick(Sender: TObject);
begin
  inherited;
  SO_BTREPMETR.enabled := (SO_METRESEXCEL.checked);
end;

procedure TFassistParamSoc.SO_CAPITALExit(Sender: TObject);
begin
  inherited;
  SO_CAPITAL.text := StringReplace (SO_CAPITAL.text,'.',',',[rfReplaceAll]);
end;

end.



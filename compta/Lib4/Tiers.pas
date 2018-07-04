{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/10/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : 07/10/2003 : FQ 12266 Modif DFM pour prise en compte
Suite ........ : des champs manquants
Mots clefs ... :
*****************************************************************}
unit Tiers;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, {$IFNDEF DBXPRESS}dbtables, HSysMenu, Menus, hmsgbox, DBCtrls,
  StdCtrls, ComCtrls, HRichEdt, HRichOLE, Hctrls, Hcompte, HRegCpte, Mask,
  ExtCtrls, Buttons, TntComCtrls, TntStdCtrls{$ELSE}uDbxDataSet{$ENDIF}, Grids, DBGrids, StdCtrls, Buttons, ExtCtrls, DBCtrls, Mask,
  hmsgbox, ComCtrls, ent1, Hcompte, HDB, hctrls, Spin, HEnt1, Menus, Filtre,UtilPGI,
{$IFNDEF PGIIMMO}
  CritEdt,
{$ENDIF}
  UTob,
  HRichEdt, HSysMenu, HRichOLE, Region, MajTable, Hqry, HRegCpte, ParamSoc,
  LicUtil,fe_main, HStatus;

Procedure FicheTiers(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Procedure FicheTiersMZS(Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : String);

type
  TFCpteTiersCpt = class(TForm)
    STiers: TDataSource;
    DBNav : TDBNavigator;
    HPB: TPanel;
    BFerme      : TBitBtn;
    BFirst      : TBitBtn;
    BPrev       : TBitBtn;
    BNext       : TBitBtn;
    BLast       : TBitBtn;
    BImprimer   : TBitBtn;
    BValider    : TBitBtn;
    BAnnuler    : TBitBtn;
    BAide       : TBitBtn;
    BInsert     : TBitBtn;
    BContact    : TBitBtn;
    BRIB        : TBitBtn;
    FAutoSave   : TCheckBox;
    MsgBox      : THMsgBox;
    PInformation: TTabSheet;
    Pages       : TPageControl;
    PComptable  : TTabSheet;
    PCommercial : TTabSheet;
    PCoordonnees: TTabSheet;
    TReglement  : TTabSheet;
    TT_AUXILIAIRE       : THLabel;
    TT_CORRESP1         : THLabel;
    TT_CORRESP2         : THLabel;
    T_AUXILIAIRE        : TDBEdit;
    TT_TARIFTIERS       : THLabel;
    T_TARIFTIERS        : THDBValComboBox;
    TT_REMISE           : THLabel;
    T_REMISE            : TDBEdit;
    T_FACTUREHT         : TDBCheckBox;
    TT_ASSURANCE: TGroupBox;
    TT_TOTALDEBIT       : THLabel;
    TT_EAN              : THLabel;
    T_EAN               : TDBEdit;
    T_LIBELLE           : TDBEdit;
    TT_LIBELLE          : THLabel;
    TT_ADRESSE          : THLabel;
    T_ADRESSE1          : TDBEdit;
    T_ADRESSE2          : TDBEdit;
    T_ADRESSE3          : TDBEdit;
    T_CODEPOSTAL        : TDBEdit;
    TT_CODEPOSTAL       : THLabel;
    TT_VILLE            : THLabel;
    T_VILLE             : TDBEdit;
    T_PAYS              : THDBValComboBox;
    TT_PAYS             : THLabel;
    TT_DIVTERRIT        : THLabel;
    TT_TELEPHONE        : THLabel;
    T_TELEPHONE         : TDBEdit;
    TT_FAX              : THLabel;
    T_FAX               : TDBEdit;
    T_TELEX             : TDBEdit;
    TT_TELEX            : THLabel;
    TT_TRESORERIE       : TGroupBox;
    TT_RELANCEREGLEMENT : THLabel;
    TT_RELANCETRAITE    : THLabel;
    TT_LETTREPAIEMENT   : THLabel;
    TT_RESIDENTETRANGER : THLabel;
    TT_NATUREECONOMIQUE : THLabel;
    T_RELANCEREGLEMENT  : THDBValComboBox;
    T_RELANCETRAITE     : THDBValComboBox;
    T_LETTREPAIEMENT    : THDBValComboBox;
    T_RESIDENTETRANGER  : THDBValComboBox;
    T_NATUREECONOMIQUE  : THDBValComboBox;
    T_MODEREGLE         : THDBValComboBox;
    T_JOURPAIEMENT1     : TDBEdit;
    T_JOURPAIEMENT2     : TDBEdit;
    TT_MODEREGLE        : THLabel;
    TT_JOURPAIEMENT1: THLabel;
    TT_NIF              : THLabel;
    T_NIF               : TDBEdit;
    TT_SIRET            : THLabel;
    TT_APE              : THLabel;
    TT_COMMENTAIRE      : THLabel;
    TT_LANGUE           : THLabel;
    T_LANGUE            : THDBValComboBox;
    T_COMMENTAIRE       : TDBEdit;
    TT_RELEVEFACTURE    : TGroupBox;
    TT_FREQRELEVE       : THLabel;
    TT_JOURRELEVE       : THLabel;
    T_JOURRELEVE        : TDBEdit;
    TT_PAYEUR           : THLabel;
    T_PAYEUR            : THDBCpteEdit;
    TT_FACTURE          : THLabel;
    T_FACTURE           : THDBCpteEdit;
    TT_JOURPAIEMENT2: THLabel;
    TT_ADRESSE2         : THLabel;
    TT_ADRESSE3         : THLabel;
    HGBDates            : TGroupBox;
    TT_DATEOUVERTURE    : THLabel;
    TT_DATEFERMETURE    : THLabel;
    TT_DATEMODIFICATION : THLabel;
    TT_DATECREATION     : THLabel;
    T_DATEMODIF: TDBEdit;
    T_DATEOUVERTURE     : TDBEdit;
    T_DATEFERMETURE     : TDBEdit;
    T_DATECREATION      : TDBEdit;
    T_FERME             : TDBCheckBox;
    TT_COLLECTIF        : THLabel;
    T_COLLECTIF         : THDBCpteEdit;
    T_RELEVEFACTURE     : TDBCheckBox;
    T_ABREGE            : TDBEdit;
    TT_ABREGE           : THLabel;
    T_RVA               : TDBEdit;
    TT_RVA              : THLabel;
    T_JURIDIQUE         : THDBValComboBox;
    T_MOTIFVIREMENT     : THDBValComboBox;
    TT_MOTIFVIREMENT    : THLabel;
    T_FREQRELEVE        : THDBValComboBox;
    T_DATEDERNRELEVE    : TDBText;
    HLabel9             : THLabel;
    TT_NATUREAUX        : THLabel;
    T_NATUREAUXI        : THDBValComboBox;
    TT_COUTHORAIRE      : THLabel;
    T_COUTHORAIRE       : TDBEdit;
    NatAux              : TEdit;
    LNat: TLabel;
    PopZ                : TPopupMenu;
    BMenuZoom           : TBitBtn;
    BGLAux              : TBitBtn;
    BBalAux             : TBitBtn;
    BBalAge             : TBitBtn;
    BBalVen             : TBitBtn;
    BCumP               : TBitBtn;
    BEchean             : TBitBtn;
    BJustSold           : TBitBtn;
    BCumul              : TBitBtn;
    BZecrimvt           : TBitBtn;
    BJustif             : TBitBtn;
    T_CONFIDENTIEL      : TDBCheckBox;
    SolCpta: THNumEdit;
    Bevel4: TBevel;
    Bevel5: TBevel;
    T_APE: TDBEdit;
    T_SIRET: TDBEdit;
    TT_OptionImpression: TGroupBox;
    T_SOLDEPROGRESSIF: TDBCheckBox;
    T_TOTAUXMENSUELS: TDBCheckBox;
    T_SAUTPAGE: TDBCheckBox;
    HGBDERNMOUV: TGroupBox;
    TG_DEBITDERNMVT: THLabel;
    TG_CREDITDERNMVT: THLabel;
    TG_NUMDERNMVT: THLabel;
    HLabel10: THLabel;
    T_DATEDERNMVT: TDBEdit;
    T_DEBITDERNMVT: TDBEdit;
    T_CREDITDERNMVT: TDBEdit;
    T_NUMDERNMVT: TDBEdit;
    T_LIGNEDERNMVT: TDBEdit;
    TT_TVAENCAISSEMENT: THLabel;
    T_TVAENCAISSEMENT: THDBValComboBox;
    HLRegimeTVA: THLabel;
    T_REGIMETVA: THDBValComboBox;
    T_SOUMISTPF: TDBCheckBox;
    TT_DEVISE: THLabel;
    T_DEVISE: THDBValComboBox;
    T_MULTIDEVISE: TDBCheckBox;
    TT_ESCOMPTE: THLabel;
    T_ESCOMPTE: TDBEdit;
    TT_ESCOMPTEPOURCENT: THLabel;
    GroupBox1: TGroupBox;
    HLabel2: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    T_TOTDEBP: TDBEdit;
    T_TOTCREP: TDBEdit;
    T_TOTDEBE: TDBEdit;
    T_TOTCREE: TDBEdit;
    T_TOTDEBS: TDBEdit;
    T_TOTCRES: TDBEdit;
    TSOLCREP: THNumEdit;
    TSOLCREE: THNumEdit;
    TSOLCRES: THNumEdit;
    TA_BLOCNOTE: TGroupBox;
    T_BLOCNOTE: THDBRichEditOLE;
    T_DERNLETTRAGE: TDBEdit;
    TT_DERNLETTRAGE: THLabel;
    T_LETTRABLE: TDBCheckBox;
    Bevel6: TBevel;
    Bevel1: TBevel;
    Bevel2: TBevel;
    Bevel8: TBevel;
    Bevel10: TBevel;
    HMTrad: THSystemMenu;
    HLabel11: THLabel;
    BBalAuxGen: TBitBtn;
    T_DIVTERRIT: TDBEdit;
    T_CORRESP1: THDBCpteEdit;
    T_CORRESP2: THDBCpteEdit;
    ZL: TTabSheet;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE0: THDBCpteEdit;
    T_TABLE1: THDBCpteEdit;
    T_TABLE2: THDBCpteEdit;
    T_TABLE3: THDBCpteEdit;
    T_TABLE4: THDBCpteEdit;
    T_TABLE5: THDBCpteEdit;
    T_TABLE6: THDBCpteEdit;
    T_TABLE7: THDBCpteEdit;
    T_TABLE8: THDBCpteEdit;
    T_TABLE9: THDBCpteEdit;
    QTiers: TQuery;
    QTiersT_AUXILIAIRE: TStringField;
    QTiersT_NATUREAUXI: TStringField;
    QTiersT_LIBELLE: TStringField;
    QTiersT_EAN: TStringField;
    QTiersT_COLLECTIF: TStringField;
    QTiersT_ABREGE: TStringField;
    QTiersT_COMMENTAIRE: TStringField;
    QTiersT_ADRESSE1: TStringField;
    QTiersT_ADRESSE2: TStringField;
    QTiersT_ADRESSE3: TStringField;
    QTiersT_CODEPOSTAL: TStringField;
    QTiersT_VILLE: TStringField;
    QTiersT_DIVTERRIT: TStringField;
    QTiersT_PAYS: TStringField;
    QTiersT_LANGUE: TStringField;
    QTiersT_DEVISE: TStringField;
    QTiersT_MULTIDEVISE: TStringField;
    QTiersT_NIF: TStringField;
    QTiersT_SIRET: TStringField;
    QTiersT_APE: TStringField;
    QTiersT_TELEPHONE: TStringField;
    QTiersT_FAX: TStringField;
    QTiersT_TELEX: TStringField;
    QTiersT_FACTURE: TStringField;
    QTiersT_PAYEUR: TStringField;
    QTiersT_SECTEUR: TStringField;
    QTiersT_REMISE: TFloatField;
    QTiersT_ESCOMPTE: TFloatField;
    QTiersT_QUALIFESCOMPTE: TStringField;
    QTiersT_MODEREGLE: TStringField;
    QTiersT_JOURPAIEMENT1: TIntegerField;
    QTiersT_JOURPAIEMENT2: TIntegerField;
    QTiersT_FACTUREHT: TStringField;
    QTiersT_REGIMETVA: TStringField;
    QTiersT_SOUMISTPF: TStringField;
    QTiersT_LETTRABLE: TStringField;
    QTiersT_CORRESP1: TStringField;
    QTiersT_CORRESP2: TStringField;
    QTiersT_TOTALDEBIT: TFloatField;
    QTiersT_TOTALCREDIT: TFloatField;
    QTiersT_BLOCNOTE: TMemoField;
    QTiersT_DATECREATION: TDateTimeField;
    QTiersT_DATEMODIF: TDateTimeField;
    QTiersT_DATEOUVERTURE: TDateTimeField;
    QTiersT_DATEFERMETURE: TDateTimeField;
    QTiersT_FERME: TStringField;
    QTiersT_DATEDERNMVT: TDateTimeField;
    QTiersT_DEBITDERNMVT: TFloatField;
    QTiersT_CREDITDERNMVT: TFloatField;
    QTiersT_NUMDERNMVT: TIntegerField;
    QTiersT_LIGNEDERNMVT: TIntegerField;
    QTiersT_CONFIDENTIEL: TStringField;
    QTiersT_SOLDEPROGRESSIF: TStringField;
    QTiersT_SAUTPAGE: TStringField;
    QTiersT_TOTAUXMENSUELS: TStringField;
    QTiersT_COUTHORAIRE: TFloatField;
    QTiersT_SOCIETE: TStringField;
    QTiersT_RELANCEREGLEMENT: TStringField;
    QTiersT_RELANCETRAITE: TStringField;
    QTiersT_RESIDENTETRANGER: TStringField;
    QTiersT_NATUREECONOMIQUE: TStringField;
    QTiersT_MOTIFVIREMENT: TStringField;
    QTiersT_LETTREPAIEMENT: TStringField;
    QTiersT_RELEVEFACTURE: TStringField;
    QTiersT_FREQRELEVE: TStringField;
    QTiersT_JOURRELEVE: TIntegerField;
    QTiersT_RVA: TStringField;
    QTiersT_DERNLETTRAGE: TStringField;
    QTiersT_JURIDIQUE: TStringField;
    QTiersT_TVAENCAISSEMENT: TStringField;
    QTiersT_TOTDEBP: TFloatField;
    QTiersT_TOTCREP: TFloatField;
    QTiersT_TOTDEBE: TFloatField;
    QTiersT_TOTCREE: TFloatField;
    QTiersT_TOTDEBS: TFloatField;
    QTiersT_TOTCRES: TFloatField;
    QTiersT_TOTDEBANO: TFloatField;
    QTiersT_TOTCREANO: TFloatField;
    QTiersT_TOTDEBANON1: TFloatField;
    QTiersT_TOTCREANON1: TFloatField;
    QTiersT_DATEDERNRELEVE: TDateTimeField;
    QTiersT_SCORERELANCE: TIntegerField;
    QTiersT_CREERPAR: TStringField;
    QTiersT_EXPORTE: TStringField;
    QTiersT_SCORECLIENT: TIntegerField;
    QTiersT_PAYEURECLATEMENT: TStringField;
    QTiersT_TABLE0: TStringField;
    QTiersT_TABLE1: TStringField;
    QTiersT_TABLE2: TStringField;
    QTiersT_TABLE3: TStringField;
    QTiersT_TABLE4: TStringField;
    QTiersT_TABLE5: TStringField;
    QTiersT_TABLE6: TStringField;
    QTiersT_TABLE7: TStringField;
    QTiersT_TABLE8: TStringField;
    QTiersT_TABLE9: TStringField;
    TT_CONSO: THLabel;
    T_CONSO: TDBEdit;
    QTiersT_CONSO: TStringField;
    QTiersT_UTILISATEUR: TStringField;
    QTiersT_TARIFTIERS: TStringField;
    QTiersT_PROFIL: TStringField;
    QTiersT_CREDITDEMANDE: TFloatField;
    QTiersT_CREDITACCORDE: TFloatField;
    QTiersT_DOSSIERCREDIT: TStringField;
    QTiersT_DATECREDITDEB: TDateTimeField;
    QTiersT_DATECREDITFIN: TDateTimeField;
    QTiersT_CREDITPLAFOND: TFloatField;
    QTiersT_DATEPLAFONDDEB: TDateTimeField;
    QTiersT_DATEPLAFONDFIN: TDateTimeField;
    QTiersT_NIVEAURISQUE: TStringField;
    T_CREDITDEMANDE: TDBEdit;
    T_CREDITACCORDE: TDBEdit;
    TT_CREDITDEMANDE: THLabel;
    TT_CREDITACCORDE: THLabel;
    H_0: THLabel;
    H_1: THLabel;
    H_2: THLabel;
    H_3: THLabel;
    H_4: THLabel;
    H_5: THLabel;
    H_6: THLabel;
    H_7: THLabel;
    H_8: THLabel;
    H_9: THLabel;
    QTiersT_CODEIMPORT: TStringField;
    QTiersT_TRANSPORTEUR: TStringField;
    QTiersT_COEFCOMMA: TFloatField;
    QTiersT_FRANCO: TFloatField;
    QTiersT_DATEDERNPIECE: TDateTimeField;
    QTiersT_NUMDERNPIECE: TIntegerField;
    QTiersT_TOTDERNPIECE: TFloatField;
    QTiersT_TIERS: TStringField;
    QTiersT_NATIONALITE: TStringField;
    QTiersT_PRENOM: TStringField;
    QTiersT_JOURNAISSANCE: TIntegerField;
    QTiersT_MOISNAISSANCE: TIntegerField;
    QTiersT_ANNEENAISSANCE: TIntegerField;
    QTiersT_COMPTATIERS: TStringField;
    T_ISPAYEUR: TDBCheckBox;
    T_AVOIRRBT: TDBCheckBox;
    QTiersT_AVOIRRBT: TStringField;
    QTiersT_ISPAYEUR: TStringField;
    QTiersT_DEBRAYEPAYEUR: TStringField;
    T_DEBRAYEPAYEUR: TDBCheckBox;
    QTiersT_APPORTEUR: TStringField;
    QTiersT_ZONECOM: TStringField;
    QTiersT_MOISCLOTURE: TIntegerField;
    QTiersT_EURODEFAUT: TStringField;
    QTiersT_PARTICULIER: TStringField;
    QTiersT_SEXE: TStringField;
    QTiersT_PASSWINTERNET: TStringField;
    TT_CREDITPLAFOND: THLabel;
    T_CREDITPLAFOND: TDBEdit;
    QTiersT_ETATRISQUE: TStringField;
    QTiersT_SOCIETEGROUPE: TStringField;
    QTiersT_TELEPHONE2: TStringField;
    QTiersT_REPRESENTANT: TStringField;
    QTiersT_EMAIL: TStringField;
    QTiersT_FORMEJURIDIQUE: TStringField;
    QTiersT_PRESCRIPTEUR: TStringField;
    QTiersT_PUBLIPOSTAGE: TStringField;
    QTiersT_ORIGINETIERS: TStringField;
    QTiersT_PHONETIQUE: TStringField;
    QTiersT_DATEPROCLI: TDateTimeField;
    QTiersT_DOMAINE: TStringField;
    QTiersT_DATEINTEGR: TDateTimeField;
    QTiersT_NIVEAUIMPORTANCE: TStringField;
    QTiersT_CLETELEPHONE: TStringField;
    QTiersT_ENSEIGNE: TStringField;
    QTiersT_EMAILING: TStringField;
    QTiersT_REGION: TStringField;
    procedure BLastClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure STiersDataChange(Sender: TObject; Field: TField);
    procedure FormShow(Sender: TObject);
    procedure BContactClick(Sender: TObject);
    procedure BJustifClick(Sender: TObject);
    procedure BRIBClick(Sender: TObject);
    procedure BCumulClick(Sender: TObject);
    procedure T_AUXILIAIREExit(Sender: TObject);
    procedure QTiersBeforeDelete(DataSet: TDataSet);
    procedure BZecrimvtClick(Sender: TObject);
    procedure T_NATUREAUXIChange(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BMenuZoomClick(Sender: TObject);
    procedure BGLAuxClick(Sender: TObject);
    procedure BBalAuxClick(Sender: TObject);
    procedure BBalAgeClick(Sender: TObject);
    procedure BBalVenClick(Sender: TObject);
    procedure BCumPClick(Sender: TObject);
    procedure BEcheanClick(Sender: TObject);
    procedure BJustSoldClick(Sender: TObject);
    procedure T_COLLECTIFExit(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure T_CODEPOSTALDblClick(Sender: TObject);
    procedure T_VILLEDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure T_FACTUREExit(Sender: TObject);
    procedure T_LETTRABLEClick(Sender: TObject);
    procedure T_DEVISEChange(Sender: TObject);
    procedure T_MULTIDEVISEClick(Sender: TObject);
    procedure BBalAuxGenClick(Sender: TObject);
    procedure T_DIVTERRITDblClick(Sender: TObject);
    procedure T_PAYSChange(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure T_ISPAYEURClick(Sender: TObject);
    procedure T_AUXILIAIREChange(Sender: TObject);
    procedure PagesChanging(Sender: TObject; var AllowChange: Boolean);
  private    { Déclarations privées }
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Procedure NewEnreg ;
    Procedure ChargeEnreg ;
    Function  IsMouvemente : Boolean ;
    Function  CodeValide : Boolean ;
    Function  CodeCollectifValide : Boolean ;
    Procedure InitModif ;
    Procedure OnEstEnSaisieCompta ;
    Procedure LettrageChgAutorise ;
    Procedure NatureChgAutorise ;
    Function  LettrageChanger : Boolean ;
    Procedure DelRibContact ;
    Procedure AlerteMisEnSommeil(Sender : TObject) ;
    Function  VerifCoherenceTL : Boolean ;

    Function  Supprime(St : String) : Boolean ;

    Procedure InitModifEnSerie(StModif : string) ;
    Procedure AffecteLe(Champ, Valeur : string) ;
    Function OkConfidentiel : Boolean ;
    Procedure EnabledGB(GB : TGroupBox) ;
    Procedure AutoriseReleveFacture ;
    Procedure PourSalarie ;
    Procedure PositionnePayeur ;
    Procedure FormateLesMontants ;
    procedure Ergo35 ;
    procedure PayeurEna ;
    procedure CreateEnregGescomGRC;

    // GCO - 18/04/2002
    procedure ActivationControl;
    // Fin GCO

  public    { Déclarations publiques }
    Q : TQuery ;
    Lequel, LesModif,LeCompte  : String ;
    AvecMvt                  : Boolean ;
    Mode                     : TActionFiche ;
    LaPage                   : Integer ;
    LgCode                   : Byte ;
    MNat,ML,Dev,MultiDev     : String ;
    EnSaisieCpta         : Boolean ;
  end;


implementation

uses ventil, codepost, FichComm,
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
     QRBaAuGe,QrBalVen,QRCumulG,QrBalAge,
{$ENDIF}
     CumMens, Zecrimvt, Justisol,
     QRTier, QRGLAUX,  QrBalAux, QRJuSold, QREchean,
{$ENDIF}
{$ENDIF}
     Choix, General,
     SAISUTIL,
{$IFNDEF PGIIMMO}
     UTILEDT,
     uLibWindows,
{$ENDIF}
     UtilSais,
     CpteSav ;

{$R *.DFM}


Procedure FicheTiers(Q : TQuery; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var FCpteTiersCPT: TFCpteTiersCPT ;
BEGIN
Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not ExJaiLeDroitConcept(TConcept(ccAuxCreat),True) then Exit ;
              taModif,taModifEnSerie : if Not ExJaiLeDroitConcept(TConcept(ccAuxModif),True) then Exit ;
   END ;
FCpteTiersCPT:=TFCpteTiersCPT.Create(Application) ;
try
  FCpteTiersCPT.Q:=Q ;
  FCpteTiersCPT.Lequel:=Compte ;
  FCpteTiersCPT.Mode:=Comment ;
  FCpteTiersCPT.LaPage:=QuellePage ;
  FCpteTiersCPT.ShowModal ;
  Finally
  FCpteTiersCPT.Free ;
  End ;
Screen.Cursor:=SyncrDefault ;
END ;


Procedure FicheTiersMZS(Compte : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : String);
var FCpteTiersCPT : TFCpteTiersCPT ;
BEGIN
Case Comment of
   taCreat,taCreatEnSerie,taCreatOne : if Not ExJaiLeDroitConcept(TConcept(ccAuxCreat),True) then Exit ;
              taModif,taModifEnSerie : if Not ExJaiLeDroitConcept(TConcept(ccAuxModif),True) then Exit ;
   END ;
FCpteTiersCPT:=TFCpteTiersCPT.Create(Application) ;
try
  FCpteTiersCPT.Lequel:=Compte ;
  FCpteTiersCPT.Mode:=Comment ;
  FCpteTiersCPT.LaPage:=QuellePage ;
  FCpteTiersCPT.LesModif:=LesModif ;
  FCpteTiersCPT.ShowModal ;
  Finally
  FCpteTiersCPT.Free ;
  End ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFCpteTiersCpt.IsMouvemente : Boolean  ;
BEGIN
if ((Mode=taConsult) or (QTiers.State=dsInsert)) then BEGIN Result:=False ; Exit ; END ;
Result:=ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE="'+T_AUXILIAIRE.Text+'" '+
                  'AND EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_AUXILIAIRE="'+T_AUXILIAIRE.Text+'" )') ;
END ;

Procedure TFCpteTiersCpt.InitModif ;
BEGIN

MNat:=QTiersT_NATUREAUXI.AsString ; ML:=QTiersT_LETTRABLE.AsString ;
Dev:=QTiersT_DEVISE.AsString ; MultiDev:=QTiersT_MULTIDEVISE.AsString ;
if MultiDev='X' then T_DEVISE.Enabled:=False
                else T_DEVISE.Enabled:=True ;
END ;

Procedure TFCpteTiersCpt.OnEstEnSaisieCompta ;
BEGIN
if Not EnSaisieCpta then Exit ;
T_NATUREAUXI.Enabled:=False ; T_LETTRABLE.Enabled:=False ; BInsert.Enabled:=False ;
END ;

Function TFCpteTiersCpt.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
result:=FALSE  ;
if QTiers.Modified then
   BEGIN
   if (Mode in [taCreat..taCreatOne]) And
      (Bourreless(QTiersT_AUXILIAIRE.AsString,fbAux)='') then Rep:=mrNo
   else if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if not Bouge(nbPost) then exit ;
  mrNo  : BEGIN DelRibContact ; if not Bouge(nbCancel) then exit ; END ;
  mrCancel : Exit ;
  end ;
result:=TRUE  ;
end ;

Function TFCpteTiersCpt.CodeCollectifValide : Boolean ;
Var Q : TQuery ;
    NatAux,NatGen : String3;
BEGIN
Result:=True ;
if QTiers.state in [dsInsert,dsEdit] then
   BEGIN
   if(QTiersT_COLLECTIF.AsString='')then
      BEGIN
      Pages.ActivePage:=PComptable ; T_COLLECTIF.SetFocus ;
      MsgBox.execute(16,'','') ; Result:=False ; Exit ;
      END ;
   if Not Presence('GENERAUX','G_GENERAL',QTiersT_COLLECTIF.AsString) then
      BEGIN
      Pages.ActivePage:=PComptable ; T_COLLECTIF.SetFocus ;
      MsgBox.execute(17,'','') ; Result:=False ; Exit ;
      END else
      BEGIN
      NatAux:=T_NATUREAUXI.Value ;
      Q:=OpenSql('Select G_NATUREGENE From GENERAUX Where G_GENERAL="'+QTiersT_COLLECTIF.AsString+'"',True) ;
      NatGen:=Q.Fields[0].AsString ;
      Ferme(Q) ;
      if not NATUREGENAUXOK(NatGen, NatAux) then  // FQ 12769
        BEGIN
        Pages.ActivePage:=PComptable ; T_COLLECTIF.SetFocus ;
        MsgBox.execute(18,'','') ; Result:=False ; Exit ;
        END ;
      END ;
   END ;
END ;

Function TFCpteTiersCpt.CodeValide : Boolean ;
Var StAux : String ;
    CarInter : boolean ;
BEGIN
Result:=True ; CarInter:=False ;
if QTiers.state in [dsinsert] then
   BEGIN
   StAux:=QTiersT_AUXILIAIRE.AsString ;
   if StAux='' then
      BEGIN
      Pages.ActivePage:=PComptable ; T_AUXILIAIRE.SetFocus ;
      MsgBox.Execute(2,'','') ; Result:=False ; Exit ;
      END ;
   if ExisteCarInter(StAux) then CarInter:=true ;
   if CarInter then
      BEGIN
      Pages.ActivePage:=PComptable ; T_AUXILIAIRE.SetFocus ;
      MsgBox.Execute(41,'','') ; Result:=False ; Exit ;
      END ;
   if Presence('TIERS','T_AUXILIAIRE',StAux) then
      BEGIN
      Pages.ActivePage:=PComptable ; T_AUXILIAIRE.SetFocus ;
      MsgBox.Execute(4,'','') ; Result:=False ; Exit ;
      END ;
   if Presence('TIERS','T_TIERS',BourreLess(stAux,fbAux)) then
      BEGIN
      Pages.ActivePage:=PComptable ; T_AUXILIAIRE.SetFocus ;
      MsgBox.Execute(42,'','') ; Result:=False ; Exit ;
      END ;

   END ;
END ;

Function TFCpteTiersCpt.LettrageChanger : Boolean ;
Var AvecMvtTemp,AeteModifier : Boolean ;
BEGIN
Result:=False ; AvecMvtTemp:=False ;
if ((ML='-') And (Not(T_LETTRABLE.Checked)) Or
   ((ML='X') And (T_LETTRABLE.Checked))) then AeteModifier:=False
                                         else AeteModifier:=True ;
if (AeteModifier) And (Not AvecMvt) then AvecMvtTemp:=IsMouvemente ;
if AvecMvtTemp then
   BEGIN
   if (ML='X') And (Not(T_LETTRABLE.Checked)) then
       BEGIN
       MsgBox.Execute(36,'','') ; AvecMvt:=True ; QTiersT_LETTRABLE.AsString:=ML ; Exit ;
       END ;
   AvecMvt:=True ;
   END ;
if Not AvecMvt then Exit ;
if QTiers.State in [dsBrowse,dsInsert] then Exit ;
Result:=((ML='-') And (T_LETTRABLE.Checked)) ;
END ;

Function TFCpteTiersCpt.EnregOK : boolean ;
Var ModPaie : String ;
BEGIN
Result:=False  ;
if QTiers.state in [dsinsert,dsedit]=False then Exit ;
if Not CodeValide then Exit ;
if Not CodeCollectifValide then Exit ;
if QTiers.state in [dsinsert,dsedit] then
   BEGIN
   if QTiersT_LIBELLE.asString='' then
      BEGIN
      Pages.ActivePage:=PCoordonnees ; T_LIBELLE.SetFocus ;
      MsgBox.execute(3,'','') ; Exit ;
      END ;
   if QTiersT_PAYEUR.AsString=QTiersT_AUXILIAIRE.AsString then
      BEGIN
      Pages.ActivePage:=TReglement ; T_PAYEUR.SetFocus ;
      MsgBox.Execute(40,'','') ; Exit ;
      END ;
   if QTiersT_MODEREGLE.AsString='' then
      BEGIN
      Pages.ActivePage:=TReglement ; T_MODEREGLE.SetFocus ;
      MsgBox.execute(31,'','') ; Exit ;
      END ;
   if (QTiersT_REGIMETVA.AsString='') And (T_NATUREAUXI.Value<>'SAL') then
      BEGIN
      Pages.ActivePage:=PComptable ; T_REGIMETVA.SetFocus ;
      MsgBox.execute(32,'','') ; Exit ;
      END ;
   if Not VerifCoherenceTL then Exit ;
   END ;
Case VerifCorrespondance(2,QTiersT_CORRESP1.AsString,QTiersT_CORRESP2.AsString) of
    0 : ;
    1 : BEGIN Msgbox.Execute(37,'','') ; Exit ; END ;
    2 : BEGIN Msgbox.Execute(38,'','') ; Exit ; END ;
    END ;
if QTiers.state in [dsEdit] then
   BEGIN
   if LettrageChanger then
      BEGIN
      if ((Not _BlocCarFiche) and (MsgBox.Execute(23,'','')=mrYes)) then
         BEGIN
         MsgBox.Execute(24,'','') ; ModPaie:=Choisir(MsgBox.Mess[25],'MODEPAIE','MP_LIBELLE','MP_MODEPAIE','','') ;
         if ModPaie<>'' then MajLettrageEcriture(ModPaie,T_AUXILIAIRE.Text,fbAux) else
            BEGIN
            MsgBox.Execute(26,'','') ; QTiersT_LETTRABLE.AsString:=ML ;
            END ;
         END else QTiersT_LETTRABLE.AsString:=ML ;
      END ;
   END ;
DateModification(QTiers,'T') ;
if QTiersT_Tiers.AsString='' then QTiersT_Tiers.AsString:=BourreLess(QTiersT_Auxiliaire.AsString,fbAux) ;
if(T_NATUREAUXI.Value='CLI') Or (T_NATUREAUXI.Value='AUD') then else
  BEGIN
  QTiersT_RELANCEREGLEMENT.AsString:='' ; QTiersT_RELANCETRAITE.AsString:='' ;
  END ;
if T_NATUREAUXI.Value<>'CLI' then
   BEGIN
   QTiersT_RELEVEFACTURE.AsString:='-' ; QTiersT_JOURRELEVE.AsInteger:=0 ;
   QTiersT_FREQRELEVE.AsString:='' ;
   END ;
if QTiers.state in [dsinsert] then QTiersT_CODEIMPORT.AsString:=QTiersT_AUXILIAIRE.AsString ;
Result:=TRUE  ;
END ;

Procedure TFCpteTiersCpt.ChargeEnreg ;
BEGIN
if LeCompte=QTiersT_Auxiliaire.AsString then
   BEGIN
   if Mode=taConsult then FicheReadOnly(Self) ;
   Exit ;
   END ;
LeCompte:=QTiersT_Auxiliaire.AsString ;
NatAux.Text:=T_NATUREAUXI.Text ;
InitCaption(Self,T_AUXILIAIRE.text,T_LIBELLE.text) ; PourSalarie ; PositionnePayeur ; 
AfficheLeSolde(TSOLCREP,QTiersT_TOTDEBP.AsFloat,QTiersT_TOTCREP.AsFloat) ;
AfficheLeSolde(TSOLCREE,QTiersT_TOTDEBE.AsFloat,QTiersT_TOTCREE.AsFloat) ;
AfficheLeSolde(TSOLCRES,QTiersT_TOTDEBS.AsFloat,QTiersT_TOTCRES.AsFloat) ;
AfficheLeSolde(SolCpta,QTiersT_TOTALDEBIT.AsFloat,QTiersT_TOTALCREDIT.AsFloat) ;
if Mode<>taConsult then T_COLLECTIF.ExisteH ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
T_AUXILIAIRE.Enabled:=False ;
AvecMvt:=IsMouvemente ; InitModif ; OnEstEnSaisieCompta ;
LettrageChgAutorise ; NatureChgAutorise ; AutoriseReleveFacture ;
if Not T_RELANCETRAITE.Focused then
   if T_RELANCETRAITE.Value='' then T_RELANCETRAITE.ItemIndex:=T_RELANCETRAITE.Items.Indexof(MsgBox.Mess[33]) ;
if Not T_RELANCEREGLEMENT.Focused then
   if T_RELANCEREGLEMENT.Value='' then T_RELANCEREGLEMENT.ItemIndex:=T_RELANCEREGLEMENT.Items.Indexof(MsgBox.Mess[33]) ;
PayeurEna ; 
END ;

Procedure TFCpteTiersCpt.NewEnreg ;
Var i : integer ;
    LL  : THLabel ;
BEGIN
InitNew(QTiers) ;
AvecMvt:=False ;
T_LETTRABLE.Enabled:=Not AvecMvt ;
T_NATUREAUXI.Enabled:=True ;
T_NATUREAUXI.Value:='DIV' ;
NatAux.Text:=T_NATUREAUXI.Text ;
QTiersT_SOLDEPROGRESSIF.AsString:='X' ;
QTiersT_LETTRABLE.AsString:='X' ;
QTiersT_CONFIDENTIEL.AsInteger:=0 ;
QTiersT_DEVISE.AsString:=V_PGI.DevisePivot ;
{ Ajout ME le 19/04/2001 }
if (ctxPCL in V_PGI.PGIContexte) then
begin
  QTiersT_REGIMETVA.Value := GetParamSoc ('SO_REGIMEDEFAUT');
  QTiersT_MODEREGLE.Value := GetParamSoc ('SO_GCMODEREGLEDEFAUT');
  QTiersT_TVAENCAISSEMENT.Value := GetParamSoc ('SO_CODETVADEFAUT');
end;
{ Fin Ajout ME le 19/04/2001 }
QTiersT_FACTUREHT.AsString:='X' ;
Pages.ActivePage:=PComptable ;
T_AUXILIAIRE.Enabled:=TRUE ; T_AUXILIAIRE.Setfocus ;
LgCode:=VH^.Cpta[fbAux].Lg ;
for i:=0 to 9 do
    BEGIN
    LL:=THLabel(FindComponent('H_'+IntToStr(i))) ;
    if LL<>Nil then LL.Caption:='' ;
    END ;
END ;

Function TFCpteTiersCpt.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : if Not OnSauve  then Exit ;
   nbPost           : if Not EnregOK  then Exit ;
   nbDelete         : if Not Supprime(T_AUXILIAIRE.Text) then Exit ;
   end ;
if(QTiers.state=dsInsert) And (T_DEVISE.Value='') And (Not T_MULTIDEVISE.Checked)then T_DEVISE.Value:=V_PGI.DevisePivot ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(Msgbox.Mess[27]) ;
if (Button = nbPost) then CreateEnregGescomGRC;
Result:=True ;
if Button=NbInsert then NewEnreg ;
END ;

procedure TFCpteTiersCpt.BLastClick(Sender: TObject);
begin Bouge(nbLast) ; end;

procedure TFCpteTiersCpt.BNextClick(Sender: TObject);
begin Bouge(nbNext) ; end;

procedure TFCpteTiersCpt.BPrevClick(Sender: TObject);
begin Bouge(nbPrior) ; end;

procedure TFCpteTiersCpt.BFirstClick(Sender: TObject);
begin Bouge(nbFirst) ; end;

procedure TFCpteTiersCpt.BValiderClick(Sender: TObject);
begin
if Bouge(nbPost) then
   BEGIN
   if Mode=taCreatEnSerie then begin Bouge(nbInsert) ; exit ; end ;
   if ((Mode=taCreatOne) or (Mode=taModifEnSerie) or (V_PGI.MonoFiche)) then Close ;
   END ;
end;

procedure TFCpteTiersCpt.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

procedure TFCpteTiersCpt.BInsertClick(Sender: TObject);
begin if ExJaiLeDroitConcept(TConcept(ccAuxCreat),True) then Bouge(nbInsert) ; end;

procedure TFCpteTiersCpt.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFCpteTiersCpt.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin BFerme.SetFocus ; CanClose:=OnSauve ; end;

procedure TFCpteTiersCpt.STiersDataChange(Sender: TObject; Field: TField);
var UpEnable, DnEnable: Boolean;
begin
Binsert.Enabled:=(Not (QTiers.State in [dsInsert,dsEdit])) ;
BMenuZoom.Enabled:=(Not (QTiers.State in [dsInsert])) ;
BImprimer.Enabled:=(Not (QTiers.State in [dsInsert])) ;

// GCO - 17/04/2002
If (ctxPCL in V_PGI.PGIContexte) Then
  BEGIN
  T_NatureAuxi.Enabled := Not IsMouvemente;
  if T_NatureAuxi.Enabled then
    T_NatureAuxi.Color := ClWindow
  else
    T_NatureAuxi.Color :=   ClBtnFace;
  END ;
// FIn - GCO

if EnSaisieCpta then BInsert.Enabled:=False ;
if Field=Nil then
   BEGIN
   UpEnable := Enabled and not QTiers.BOF;
   DnEnable := Enabled and not QTiers.EOF;
   BFirst.Enabled := UpEnable; BPrev.Enabled := UpEnable;
   BNext.Enabled  := DnEnable; BLast.Enabled := DnEnable;
   ChargeEnreg ;
   END else
   BEGIN
   if ((Field.FieldName='T_LIBELLE') and (T_ABREGE.Field.AsString='')) then
      T_ABREGE.Field.AsString:=Copy(Field.AsString,1,17) ;
   END ;
end;

procedure TFCpteTiersCpt.Ergo35 ;
Var i : integer ;
    CC : THCpteEdit ;
    LL : THLabel ;
BEGIN
if ZL.TabVisible then
   BEGIN
   for i:=0 to 9 do
       BEGIN
       CC:=THCpteEdit(FindComponent('T_TABLE'+IntTostr(i))) ;
       if CC<>Nil then
          BEGIN
          if CC.Text<>'' then CC.ExisteH else
             BEGIN
             LL:=THLabel(FindComponent('H_'+IntToStr(i))) ;
             LL.Caption:='' ;
             END ;
          END ;
       END ;
   END ;
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   T_CORRESP2.Visible:=False ; TT_CORRESP2.Visible:=False ;
   T_LANGUE.Visible:=False ; TT_LANGUE.Visible:=False ;
   if EstSerie(S3) then
      BEGIN
      T_CORRESP1.Visible:=False ; TT_CORRESP1.Visible:=False ;
      T_CONSO.Visible:=False ; TT_CONSO.Visible:=False ;
      T_ISPAYEUR.Visible:=False ; T_AVOIRRBT.Visible:=False ;
      T_DEBRAYEPAYEUR.Visible:=False ; T_PAYEUR.Visible:=False ; TT_PAYEUR.Visible:=False ;
      T_SOUMISTPF.Visible:=False ;
      bBalven.Enabled:=False ; bBalAuxGen.Enabled:=False ;
// Fiche 10130 GP le 27/08/2002      PCommercial.TabVisible:=False ;
      T_CONFIDENTIEL.Visible:=False ;
//      TT_RELEVEFACTURE.Visible:=False ;
      T_JOURPAIEMENT1.Visible:=False ; TT_JOURPAIEMENT1.Visible:=False ;
      T_JOURPAIEMENT2.Visible:=False ; TT_JOURPAIEMENT2.Visible:=False ;
      END ;
   END ;
END ;

procedure TFCpteTiersCpt.FormShow(Sender: TObject);
Var OkPourChangeSQL : Boolean ;
begin
MakeZoomOLE(Handle) ; OkPourChangeSQL:=FALSE ;
T_RELANCETRAITE.Items.Add(MsgBox.Mess[33]) ; T_RELANCETRAITE.Values.Add('') ;
T_RELANCEREGLEMENT.Items.Add(MsgBox.Mess[33]) ; T_RELANCEREGLEMENT.Values.Add('') ;
RecupWhereSQL(Q,QTiers) ;
if ((Q=NIL) and (Lequel<>'')) then QTiers.SQL.Add('Where T_AUXILIAIRE="'+Lequel+'"') else
 if ((Q=Nil) and (Lequel='') and ((Mode=taCreatEnSerie) Or (Mode=taCreatEnSerie))) then
   BEGIN
   QTiers.SQL.Add('Where T_AUXILIAIRE="zzz"') ; OkPourChangeSQL:=TRUE ;
   END ;
if (Q=NIL) then QTiers.SQL.Add('Order by T_AUXILIAIRE') ;//lhe le 19/02/97 pour trier [tacreat..tacreatenserie] avec navigation
// JLD 08/08/2001 Pourquoi ne pas le faire systématiquement ??? Cela plante sous Oracle en création
if ((Q=NIL) and (Lequel<>'')) then ChangeSQL(QTiers) else If OkPourChangeSQL Then ChangeSQL(QTiers) ;
//ChangeSQL(QTiers) ;
ChangeSizeMemo(QTiersT_BLOCNOTE) ;
T_NATUREAUXI.DataType:='ttNatTiersCPTA' ;
Pages.ActivePage:=Pages.Pages[LaPage] ; QTiers.Open ; FormateLesMontants ;
if Mode<>taConsult then LgCode:=VH^.Cpta[fbAux].Lg ;
if ((not V_PGI.MonoFiche) and (Lequel<>'') And ((Mode in [taCreat..taCreatOne])=False)) then
  BEGIN
  if Not QTiers.Locate('T_AUXILIAIRE',Lequel,[]) then //vt avant = loCaseInsensitive
     BEGIN MessageAlerte(Msgbox.Mess[27]) ; PostMessage(Handle,WM_CLOSE,0,0) ; Exit ; END ;
  END ;
if Not OkConfidentiel then Exit ;
Case Mode Of
     taConsult           : FicheReadOnly(Self) ;
     taCreat..taCreatOne : BEGIN Bouge(nbInsert) ; T_AUXILIAIRE.Text:=Lequel ; BAnnuler.enabled:=False ; END ;
     taModif             : EnSaisieCpta:=SaisieLancer ;
     taModifEnSerie      : InitModifEnSerie(LesModif);
   End ;
BFirst.Visible := not V_PGI.MonoFiche ;
BPrev.Visible := not V_PGI.MonoFiche ;
BNext.Visible := not V_PGI.MonoFiche ;
BLast.Visible := not V_PGI.MonoFiche ;
Ergo35 ;
end;

Procedure TFCpteTiersCpt.PositionnePayeur ;
Var NatTiers : String ;
BEGIN
T_PAYEUR.ZoomTable:=tzTPayeur ; NatTiers:=T_NATUREAUXI.Value ;
if NatTiers='' then Exit ;
if NatTiers='CLI' then T_PAYEUR.ZoomTable:=tzTPayeurCLI else
   if NatTiers='FOU' then T_PAYEUR.ZoomTable:=tzTPayeurFOU else
      if NatTiers='SAL' then T_PAYEUR.ZoomTable:=tzTPayeur else
         if NatTiers='AUD' then T_PAYEUR.ZoomTable:=tzTPayeurCLI else
            if NatTiers='AUC' then T_PAYEUR.ZoomTable:=tzTPayeurFOU else
               if NatTiers='DIV' then T_PAYEUR.ZoomTable:=tzTPayeur ;
END ;

procedure TFCpteTiersCpt.T_NATUREAUXIChange(Sender: TObject);
Var NatTiers : String3 ;
begin
NatAux.Text:=T_NATUREAUXI.Text ;
if AvecMvt then Exit ;
NatTiers:=T_NATUREAUXI.Value ;
if NatTiers='CLI' then QTiersT_COLLECTIF.AsString:=VH^.DefautCli else
   if NatTiers='FOU' then QTiersT_COLLECTIF.AsString:=VH^.DefautFou else
      if NatTiers='SAL' then QTiersT_COLLECTIF.AsString:=VH^.DefautSal else
         if NatTiers='AUD' then QTiersT_COLLECTIF.AsString:=VH^.DefautDDivers else
            if NatTiers='AUC' then QTiersT_COLLECTIF.AsString:=VH^.DefautCDivers else
               if NatTiers='DIV' then QTiersT_COLLECTIF.AsString:=VH^.DefautDivers ;
AutoriseReleveFacture ; PourSalarie ; PositionnePayeur ;
end;

Procedure TFCpteTiersCpt.NatureChgAutorise ;
BEGIN
T_NATUREAUXI.Enabled:=True ;
//if Mode<>taModif then Exit ;
if EnSaisieCpta  then Exit ;
if Not AvecMvt   then Exit ;
T_NATUREAUXI.Enabled:=False ;
END ;

Procedure TFCpteTiersCpt.LettrageChgAutorise ;
BEGIN
T_LETTRABLE.Enabled:=True ;
//if Mode<>taModif then Exit ;
if EnSaisieCpta  then Exit ;
if Not AvecMvt   then Exit ;
T_LETTRABLE.Enabled:=(Not T_LETTRABLE.Checked) ;
END ;

procedure TFCpteTiersCpt.BRIBClick(Sender: TObject);
Var Ok : Boolean ;
    LeMode : TActionFiche ;
begin
if QTiers.State=dsInsert then
   if Not CodeValide then Exit ;
LeMode:=Mode ;
Ok:=ExisteSQL('Select R_AUXILIAIRE from RIB where R_AUXILIAIRE="'+T_AUXILIAIRE.Text+'"') ;
if((LeMode=taConsult)And(Not Ok))then BEGIN Msgbox.execute(19,'','') ; Exit ; END ;
if((LeMode<>taConsult)And(Not Ok)) then if Msgbox.execute(20,'','')<>mrYes then Exit else LeMode:=taCreatOne ;
if Ok And (LeMode in [taCreat..taCreatOne]) then LeMode:=taModif ;
FicheRIB_AGL(T_AUXILIAIRE.Text,LeMode,False,'',TRUE) ;
end;

procedure TFCpteTiersCpt.BContactClick(Sender: TObject);
Var Ok : Boolean ;
    LeMode : TActionFiche ;
begin
if QTiers.State=dsInsert then
   if Not CodeValide then Exit ;
LeMode:=Mode ;
Ok:=ExisteSQL('Select C_AUXILIAIRE from CONTACT where C_AUXILIAIRE="'+T_AUXILIAIRE.Text+'"') ;
if((LeMode=taConsult)And(Not Ok))then
  BEGIN Msgbox.Execute(21,'','') ; Exit ; END ;
if((LeMode<>taConsult)And(Not Ok)) then if Msgbox.Execute(22,'','')<>mrYes then Exit else LeMode:=taCreat ;
if Ok And (LeMode in [taCreat..taCreatOne]) then LeMode:=taModif ;
FicheContact_AGL(T_AUXILIAIRE.Text,LeMode) ;
end;

procedure TFCpteTiersCpt.BJustifClick(Sender: TObject);
begin (* rony 14/04/97 JustifSolde(T_AUXILIAIRE.text,fbAux); *) end;

procedure TFCpteTiersCpt.BCumulClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
CumulCpteMensuel(fbAux,QTiersT_AUXILIAIRE.AsString,QTiersT_LIBELLE.AsString,VH^.Entree) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.BZecrimvtClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
ZoomEcritureMvt(T_AUXILIAIRE.Text, fbAux,'MULMMVTS' ) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.T_AUXILIAIREExit(Sender: TObject);
begin
if(QTiers.State=dsInsert)And(T_AUXILIAIRE.Field.AsString<>'')then
   BEGIN
   QTiersT_AUXILIAIRE.AsString:=Trim(QTiersT_AUXILIAIRE.AsString) ;
   if Length(QTiersT_AUXILIAIRE.AsString)>LgCode then
//      QTiersT_AUXILIAIRE.AsString:=Copy(QTiersT_AUXILIAIRE.AsString,1,LgCode) ;
   BEGIN
      QTiersT_AUXILIAIRE.AsString:=Copy(QTiersT_AUXILIAIRE.AsString,1,LgCode) ;
      QTiersT_AUXILIAIRE.AsString:=Trim(QTiersT_AUXILIAIRE.AsString) ;
   END;   
   if Length(QTiersT_AUXILIAIRE.AsString)<LgCode then
      QTiersT_AUXILIAIRE.AsString:=BourreLaDonc(QTiersT_AUXILIAIRE.AsString,fbAux) ;
   END ;
{ Ajout ME le 19/04/2001 }
if (QTiersT_COLLECTIF.AsString <> '') and (QTiersT_AUXILIAIRE.asstring <> '') and (QTiersT_LIBELLE.asstring = '') and (ctxPCL in V_PGI.PGIContexte) then
begin
QTiersT_LIBELLE.asstring := AGLLanceFiche('CP','LIBELLECPTE','','','') ;
if QTiersT_MODEREGLE.AsString = '' then QTiersT_MODEREGLE.AsString := T_MODEREGLE.values[0];
end;
end;

procedure TFCpteTiersCpt.T_FACTUREExit(Sender: TObject);
begin
if Mode<>taConsult then AlerteMisEnSommeil(Sender) ;
end;

Procedure TFCpteTiersCpt.AlerteMisEnSommeil(Sender : TObject) ;
Var QLoc : TQuery ;
BEGIN
if(QTiers.FindField(THDBCpteEdit(Sender).Name).AsString='') Or
  (QTiers.FindField(THDBCpteEdit(Sender).Name).IsNull) then Exit ;
if Length(THDBCpteEdit(Sender).Text)<LgCode then
   THDBCpteEdit(Sender).Text:=BourreLaDonc(THDBCpteEdit(Sender).Text,fbAux) ;
if THDBCpteEdit(Sender).ExisteH=0 then
   BEGIN MsgBox.Execute(35,'','') ; THDBCpteEdit(Sender).SetFocus ; Exit ; END ;
QLoc:=OpenSql('Select T_FERME From TIERS Where T_AUXILIAIRE="'+THDBCpteEdit(Sender).Text+'"',True) ;
if QLoc.Fields[0].AsString='X' then MsgBox.Execute(34,'','') ;
Ferme(QLoc) ;
END ;

procedure TFCpteTiersCpt.QTiersBeforeDelete(DataSet: TDataSet);
begin
if BourreLess(QTiersT_AUXILIAIRE.AsString,fbAux)='' then Exit ;
ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+QTiersT_AUXILIAIRE.AsString+'"') ;
ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+QTiersT_AUXILIAIRE.AsString+'"') ;
//ExecuteSQL('Delete from TARIF Where F_AUXILIAIRE="'+QTiersT_AUXILIAIRE.AsString+'"') ;
//Table tarif non encore créée le 20/02/97
end;

Function TFCpteTiersCpt.Supprime(St : String) : Boolean ;
BEGIN
Result:=False ;
if MsgBox.Execute(1,'','')<>mrYes then Exit ;
Result:=True ;
END ;


//**************************
Procedure TFCpteTiersCpt.InitModifEnSerie(StModif : string) ;
var St,Champ, Valeur : string;
    i             : integer;
    B             : TBitBtn;
BEGIN
if QTiers.State=dsBrowse then QTiers.Edit ;
While StModif<>'' do
   BEGIN
   St:=ReadTokenSt(StModif);
   i:=Pos('=',St); if i>0 then Champ:=Trim(Copy(St,1,i-1));
   i:=Pos('"',St); if i>0 then St:=Trim(Copy(St,i+1,Length(St)));
   i:=Pos('"',St); if i>0 then Valeur:=Trim(Copy(St,1,i-1));
   AffecteLe(Champ,Valeur);
   END;
For i:=0 to HPB.ControlCount-1 do
   if HPB.Controls[i] is TBitBtn then
      BEGIN
      B:=TBitBtn(HPB.Controls[i]);
      if ((B.Name<>'BValider') and (B.Name<>'BFerme') and (B.Name<>'BAide')) then B.Enabled:=false;
      END;
END;

Procedure TFCpteTiersCpt.AffecteLe(Champ, Valeur : string) ;
var C : TControl;
BEGIN
C:=TControl(FindComponent(Champ)) ;
if(C is TDBCheckBox)Or(C is THDBValComboBox)Or(C is TDBEdit)Or(C is THDBCpteEdit)then
   BEGIN
   QTiers.FindField(Champ).AsString:=Valeur ; TEdit(C).Font.Color:=clRed;
   END else if C is THDBSpinEdit then
   BEGIN
   QTiers.FindField(Champ).AsInteger:=StrToInt(Valeur) ; THDBSpinEdit(C).Font.Color:=clRed;
   END ;
END;

procedure TFCpteTiersCpt.BImprimerClick(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
PlanAuxi(T_AUXILIAIRE.Text,True) ;
{$ENDIF}
{$ENDIF}
end;

Function TFCpteTiersCpt.OkConfidentiel : Boolean ;
BEGIN
Result:=False ;
if V_PGI.Confidentiel='0' then
   BEGIN
   if T_Confidentiel.Checked then
     BEGIN
     MessageAlerte(MsgBox.Mess[30]) ;
     PostMessage(Handle,WM_CLOSE,0,0) ; Exit ;
     END ;
   END ;
T_CONFIDENTIEL.Visible:=(V_PGI.Confidentiel='1') ;
T_CONFIDENTIEL.Enabled:=V_PGI.Superviseur ;
Result:=True ;
END ;

procedure TFCpteTiersCpt.BMenuZoomClick(Sender: TObject);
BEGIN PopZoom(BMenuZoom,POPZ) ; END ;

{---------- Procédures de zoom sur les états à partir de la fiche compte -------------------}

procedure TFCpteTiersCpt.BGLAuxClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1,D2 : TdateTime ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
if VH^.Entree.Code=VH^.Suivant.Code then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neGL ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
GLAuxiliaireZoom(Crit) ;
{$ENDIF}
{$ENDIF}
end ;

procedure TFCpteTiersCpt.BBalAuxClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1 ,D2 : TdateTime ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
if VH^.Entree.Code=VH^.Suivant.Code then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
BalanceAuxiZoom(Crit) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.BBalAgeClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1 : TdateTime ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ;
if VH^.Entree.Code=VH^.Suivant.Code then D1:=VH^.Suivant.Deb ;
Crit.Date1:=D1 ;
Crit.NatureEtat:=neGlV ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
BalanceAgeZoom(Crit) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.BBalVenClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1 : TdateTime ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ;
if VH^.Entree.Code=VH^.Suivant.Code then D1:=VH^.Suivant.Deb ;
Crit.Date1:=D1 ;
Crit.NatureEtat:=neGlV ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
BalVentileZoom(Crit) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.BCumPClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit  : TCritEdt ;
    D1,D2 : TdateTime ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neCum ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.Text ;
Crit.Cpt2:=Crit.Cpt1 ;
Crit.NatureCpt:=T_NATUREAUXI.Value ;
CumulPeriodiqueZoom(Crit,fbAux) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.BEcheanClick(Sender: TObject);
(*Var Crit  : TCritEdt ;
    D1,D2 : TdateTime ; *)
begin
(* rony 14/04/97
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=V_PGI.Encours.Deb ; D2:=V_PGI.Encours.Fin ;
if V_PGI.Entree.Code=V_PGI.Suivant.Code then BEGIN D1:=V_PGI.Suivant.Deb ; D2:=V_PGI.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.NatureEtat:=neEch ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
EcheancierZoom(Crit) ;
*)
end;

procedure TFCpteTiersCpt.BJustSoldClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit  : TCritEdt ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
Fillchar(Crit,SizeOf(Crit),#0) ; Crit.Date1:=V_PGI.DateEntree ; Crit.NatureEtat:=neJu ;
Crit.JU.OnTiers:=TRUE ;
InitCritEdt(Crit) ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
JustSoldeZoom(Crit) ;
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.T_COLLECTIFExit(Sender: TObject);
begin
if Mode<>taConsult then T_COLLECTIF.ExisteH ;
{ Ajout ME le 19/04/2001 }
if (QTiersT_LIBELLE.asstring = '') and (ctxPCL in V_PGI.PGIContexte) then
QTiersT_LIBELLE.asstring := AGLLanceFiche('CP','LIBELLECPTE','','','') ;

end;

Procedure TFCpteTiersCpt.DelRibContact ;
BEGIN
if BourreLess(QTiersT_AUXILIAIRE.AsString,fbAux)='' then Exit ;
if Mode in [taCreat..taCreatOne] then
   BEGIN
   ExecuteSQL('Delete from CONTACT Where C_AUXILIAIRE="'+QTiersT_AUXILIAIRE.AsString+'"') ;
   ExecuteSQL('Delete from RIB Where R_AUXILIAIRE="'+QTiersT_AUXILIAIRE.AsString+'"') ;
   END ;
END ;

procedure TFCpteTiersCpt.FormCreate(Sender: TObject);
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
{$ENDIF}
{$ENDIF}
Q:=NIL ;
LeCompte:='azeaze' ;
end;

procedure TFCpteTiersCpt.T_CODEPOSTALDblClick(Sender: TObject);
begin VerifCodePostal(QTiers,T_CODEPOSTAL,T_VILLE,TRUE) ; end;

procedure TFCpteTiersCpt.T_VILLEDblClick(Sender: TObject);
begin VerifCodePostal(QTiers,T_CODEPOSTAL,T_VILLE,FALSE) ; end;

procedure TFCpteTiersCpt.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
if Vide then
   BEGIN
   Case Key of
        VK_F3  : BPrevClick(Nil) ;
        VK_F4  : BNextClick(Nil) ;
        VK_F10 : BValiderClick(Nil) ;
        VK_RETURN :BEGIN
                   if ActiveControl is TCustomMemo then Exit ;
                   FindNextControl(ActiveControl,True,True,False).SetFocus ;
                   END ;
     END ;
   END ;
end;

Procedure TFCpteTiersCpt.EnabledGB(GB : TGroupBox) ;
Var l : Integer ;
BEGIN
For l:=0 To GB.ControlCount-1 Do GB.Controls[l].Enabled:=GB.Enabled ;
END ;



Procedure TFCpteTiersCpt.AutoriseReleveFacture ;
BEGIN
if(T_NATUREAUXI.Value='CLI') Or (T_NATUREAUXI.Value='AUD') then
  BEGIN
  if T_LETTRABLE.Checked then
     BEGIN
     if Not T_RELANCEREGLEMENT.Enabled then
        BEGIN
        T_RELANCEREGLEMENT.Enabled:=True ; T_RELANCETRAITE.Enabled:=True ;
        TT_RELEVEFACTURE.Enabled:=True ;
        END ;
     END else
     BEGIN
     T_RELANCEREGLEMENT.Enabled:=False ; T_RELANCETRAITE.Enabled:=False ;
     TT_RELEVEFACTURE.Enabled:=False ;
     T_RELEVEFACTURE.Checked:=False ; T_JOURRELEVE.Text:='0' ;
     T_FREQRELEVE.ItemIndex:=-1 ;
     END ;
  {Assurance crédit}
  TT_ASSURANCE.Enabled:=True ;
  END else
  BEGIN
  T_RELANCEREGLEMENT.Enabled:=False ; T_RELANCETRAITE.Enabled:=False ;
  TT_RELEVEFACTURE.Enabled:=False ;
  {Assurance crédit}
  TT_ASSURANCE.Enabled:=False ;
  END ;
EnabledGB(TT_ASSURANCE) ; SolCpta.Enabled:=False ;
EnabledGB(TT_RELEVEFACTURE) ;
TT_RELANCEREGLEMENT.Enabled:=T_RELANCEREGLEMENT.Enabled ; TT_RELANCETRAITE.Enabled:=T_RELANCETRAITE.Enabled ;
END ;

procedure TFCpteTiersCpt.T_LETTRABLEClick(Sender: TObject);
begin AutoriseReleveFacture ; end;

Procedure TFCpteTiersCpt.PourSalarie ;
BEGIN
if T_NATUREAUXI.Value='SAL' then
   BEGIN
   if Pages.ActivePage=PCommercial then Pages.ActivePage:=PComptable ;
   PCommercial.TabVisible:=False ;
   T_COUTHORAIRE.Visible:=True ; TT_COUTHORAIRE.Visible:=True ;
   T_PAYEUR.Enabled:=False ; TT_PAYEUR.Enabled:=False ;
   T_DEBRAYEPAYEUR.Enabled:=False ;
   BContact.Enabled:=False ; T_SOUMISTPF.Visible:=False ; HLRegimeTVA.Visible:=False ; TT_TRESORERIE.Enabled:=False ;
   T_REGIMETVA.Visible:=False ; TT_TVAENCAISSEMENT.Visible:=False ; T_TVAENCAISSEMENT.Visible:=False ;
   TT_NIF.Visible:=False ; T_NIF.Visible:=False ; TT_SIRET.Visible:=False ; T_SIRET.Visible:=False ;
   TT_APE.Visible:=False ; T_APE.Visible:=False ;
   END else
   BEGIN
   PCommercial.TabVisible:=True ;
   T_COUTHORAIRE.Visible:=False ; TT_COUTHORAIRE.Visible:=False ;
   T_PAYEUR.Enabled:=True ; TT_PAYEUR.Enabled:=True ;
   T_DEBRAYEPAYEUR.Enabled:=True ;
   BContact.Enabled:=True ; If Not EstSerie(S3) Then T_SOUMISTPF.Visible:=True ; HLRegimeTVA.Visible:=True ;
   T_REGIMETVA.Visible:=True ; TT_TVAENCAISSEMENT.Visible:=True ; T_TVAENCAISSEMENT.Visible:=True ;
   TT_NIF.Visible:=True ; T_NIF.Visible:=True ; TT_SIRET.Visible:=True ; T_SIRET.Visible:=True ;
   TT_APE.Visible:=True ; T_APE.Visible:=True ; TT_TRESORERIE.Enabled:=True ;
   END ;
END ;

procedure TFCpteTiersCpt.T_DEVISEChange(Sender: TObject);
begin
if T_MULTIDEVISE.Checked then Exit ;
if Dev<>T_DEVISE.Value then
   if AvecMvt then
      if Dev<>'' then T_DEVISE.Value:=Dev ;
end;

procedure TFCpteTiersCpt.T_MULTIDEVISEClick(Sender: TObject);
begin
if QTiers.State=dsBrowse then Exit ;
if (MultiDev='X') And (Not T_MULTIDEVISE.Checked) then
   BEGIN
   if AvecMvt then
      BEGIN
      QTiersT_MULTIDEVISE.AsString:=MultiDev ; QTiersT_DEVISE.AsString:='' ;
      END else
      BEGIN
      if Dev='' then QTiersT_DEVISE.AsString:=V_PGI.DevisePivot
                else QTiersT_DEVISE.AsString:=Dev ;
      END ;
   END else
   if T_MULTIDEVISE.Checked then QTiersT_DEVISE.AsString:='' else
      if(MultiDev='-') And (Not T_MULTIDEVISE.Checked) then
         BEGIN
         if Dev='' then QTiersT_DEVISE.AsString:=V_PGI.DevisePivot
                   else QTiersT_DEVISE.AsString:=Dev ;
         END ;
if Not T_MULTIDEVISE.Checked then T_DEVISE.Enabled:=True
                             else T_DEVISE.Enabled:=False ;
end;

procedure TFCpteTiersCpt.BBalAuxGenClick(Sender: TObject);
{$IFNDEF PGIIMMO}
Var Crit : TCritEdt ;
    D1,D2 : TdateTime ;
    Etab : String ;
{$ENDIF}
begin
{$IFNDEF PGIIMMO}
{$IFNDEF IMP}
{$IFNDEF CCS3}
Fillchar(Crit,SizeOf(Crit),#0) ;
D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
Crit.Date1:=D1 ; Crit.Date2:=D2 ;
Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
Crit.NatureEtat:=neBal ;
Crit.QualifPiece:='NOR' ;
Crit.Bal.TypCpt:=0 ;
InitCritEdt(Crit) ;
Crit.Bal.TypCpt:=1 ;
Crit.Cpt1:=T_AUXILIAIRE.text ;
Crit.Cpt2:=Crit.Cpt1 ;
Etab:=EtabForce ; if Etab<>'' then Crit.Etab:=Etab ;
BalAuGeZoom(Crit) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

procedure TFCpteTiersCpt.T_DIVTERRITDblClick(Sender: TObject);
begin
PaysRegion(T_PAYS,T_DIVTERRIT,True) ;
end;

procedure TFCpteTiersCpt.T_PAYSChange(Sender: TObject);
begin
PaysRegion(T_PAYS,T_DIVTERRIT,False) ;
end;

Procedure TFCpteTiersCpt.FormateLesMontants ;
BEGIN
QTiersT_TOTALDEBIT.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTALCREDIT.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_DEBITDERNMVT.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_CREDITDERNMVT.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTDEBP.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTCREP.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTDEBE.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTCREE.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTDEBS.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTCRES.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTDEBANO.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTCREANO.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTDEBANON1.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_TOTCREANON1.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_CREDITPLAFOND.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_CREDITDEMANDE.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
QTiersT_CREDITACCORDE.DisPlayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
ChangeMask(TSOLCREP,V_PGI.OkdecV,'') ;
ChangeMask(TSOLCREE,V_PGI.OkdecV,'') ;
ChangeMask(TSOLCRES,V_PGI.OkdecV,'') ;
END ;

procedure TFCpteTiersCpt.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFCpteTiersCpt.VerifCoherenceTL : Boolean ;
Var i : Integer ;
    C : TComponent ;
    Alerte : Boolean ;
BEGIN
Result:=True ;
if Not ZL.TabVisible then BEGIN Result:=True ; Exit ; END ;
Alerte:=False ;
for i:=0 to 9 do
  BEGIN
  C:=FindComponent('T_TABLE'+IntToStr(i)) ;
  if Not(THDBCpteEdit(C).Enabled) then Continue ;
  if THDBCpteEdit(C).Text='' then Continue ;
  if THDBCpteEdit(C).ExisteH<=0 then Alerte:=True ;
  END ;
if Alerte then if MsgBox.Execute(39,'','')<>mrYes then Result:=False ;
END ;

procedure TFCpteTiersCpt.HMTradBeforeTraduc(Sender: TObject);
begin
LibellesTableLibre(ZL,'TT_TABLE','T_TABLE','T') ;
end;

procedure TFCpteTiersCpt.PayeurEna ;
BEGIN
if T_ISPAYEUR.Checked then
   BEGIN
   T_PAYEUR.Enabled:=False ; TT_PAYEUR.Enabled:=False ;
   T_PAYEUR.Text:='' ; T_DEBRAYEPAYEUR.Enabled:=False ;
   T_AVOIRRBT.Enabled:=False ; T_AVOIRRBT.Checked:=False ;
   END else
   BEGIN
   T_PAYEUR.Enabled:=True ; TT_PAYEUR.Enabled:=True ;
   T_AVOIRRBT.Enabled:=True ; T_DEBRAYEPAYEUR.Enabled:=True ;
   END ;
END ;

procedure TFCpteTiersCpt.T_ISPAYEURClick(Sender: TObject);
begin
PayeurEna ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/04/2002
Modifié le ... : 18/04/2002
Description .. : Active ou Désactive les contrôles de la fiche lors de
Suite ........ : la création d'un compte auxiliaire afin de forcer
Suite ........ : la saisie de  T_AUXILIAIRE et T_NATUREAUXI
Mots clefs ... :
*****************************************************************}
procedure TFCpteTiersCpt.ActivationControl;
var lStAuxiliaire : string;
begin
{$IFNDEF PGIIMMO}
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  lStAuxiliaire := Trim(T_Auxiliaire.Text);

  T_Collectif.Enabled    := ( lStAuxiliaire <> '' );
  T_Collectif.Color      := IIF( T_Collectif.Enabled, ClWindow, ClBtnFace );

  T_Lettrable.Enabled    := ( lStAuxiliaire <> '' );
  T_MultiDevise.Enabled  := ( lStAuxiliaire <> '' );
  T_Confidentiel.Enabled := (T_Confidentiel.Visible) and ( lStAuxiliaire <> '' );

  T_Devise.Enabled       := ( lStAuxiliaire <> '' );
  T_Devise.Color         := IIF( T_Devise.Enabled, ClWindow, ClBtnFace );

  T_SoumisTpf.Enabled    := ( lStAuxiliaire <> '' );

  T_RegimeTva.Enabled    := ( lStAuxiliaire <> '' );
  T_RegimeTva.Color      := IIF( T_RegimeTva.Enabled, ClWindow, ClBtnFace );

  T_TvaEncaissement.Enabled := ( lStAuxiliaire <> '' );
  T_TvaEncaissement.Color   := IIF( T_TvaEncaissement.Enabled, ClWindow, ClBtnFace );

  T_Corresp1.Enabled := ( lStAuxiliaire <> '' );
  T_Corresp1.Color   := IIF( T_Corresp1.Enabled, ClWindow, ClBtnFace );

  T_Corresp2.Enabled := ( lStAuxiliaire <> '' );
  T_Corresp2.Color   := IIF( T_Corresp2.Enabled, ClWindow, ClBtnFace );

{$ENDIF}
end;

procedure TFCpteTiersCpt.PagesChanging(Sender: TObject; var AllowChange: Boolean);
begin
  if not ( CtxPcl in V_Pgi.PGIContexte ) then Exit;

  AllowChange := (Trim(T_Auxiliaire.Text) <> '');

  if not AllowChange then
  begin
    if Trim(T_Auxiliaire.Text) = '' then
      T_Auxiliaire.SetFocus
    else
      T_NatureAuxi.SetFocus;
  end;
end;

procedure TFCpteTiersCpt.T_AUXILIAIREChange(Sender: TObject);
begin
  ActivationControl;
end;

// Création d'un enregistrement vierge dans la table PROSPECTS et TIERSCOMPL si on a
// respectivement le contexte ctxGRC et ctxGescom
procedure TFCpteTiersCpt.CreateEnregGescomGRC;
var
  T : Tob;
begin
  // Tables PROSPECTS
  if (ctxGRC in V_PGI.PGIContexte) AND (T_NATUREAUXI.Value = 'CLI') then begin
    T := Tob.Create('PROSPECTS', nil, -1);
    T.InitValeurs;
    T.PutValue('RPR_AUXILIAIRE', T_AUXILIAIRE.Text);
    T.InsertDB(nil);
    T.Free;
  end;

  // Table TIERSCOMPL
  if (ctxGescom in V_PGI.PGIContexte) then begin
    T := Tob.Create('TIERSCOMPL', nil, -1);
    T.InitValeurs;
    T.PutValue('YTC_AUXILIAIRE', T_AUXILIAIRE.Text);
    T.PutValue('YTC_TIERS', BourreLess(T_AUXILIAIRE.Text, fbAux));
    T.InsertDB(nil);
    T.Free;
  end;
end;

end.

unit SOCSISCO;

interface

uses SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
     Forms, Dialogs, DBCtrls, StdCtrls, Buttons, ExtCtrls, Spin, HDB, Mask,
     Hctrls, DBTables,
     Grids, DBGrids, TabNotBk, CodePost, Hcompte,
     ComCtrls, hmsgbox, Ent1, HEnt1, HSysMenu, Region, HTB97, MajTable,
     Hqry, HRegCpte, HPanel, UiUtil, Saisutil, SaisComm, ADODB, udbxDataset,
  DB ;

Procedure FicheSocieteSISCO ( Comment : TActionFiche ; QuellePage : integer ; Verif : Boolean ; StFichier : String ) ;

Type TabCpte=Array[1..10]of String ;

type
  TFSocSISCO = class(TForm)
    TSociete: THTable;
    SSociete : TDataSource;
    Pages           : TPageControl;
    PCoordonnees    : TTabSheet;
    PComptables     : TTabSheet;
    PFourchettes    : TTabSheet;
    PComptesSpeciaux: TTabSheet;
    PDates          : TTabSheet;
    PLet            : TTabSheet;
    TSO_SOCIETE         : THLabel;
    SO_SOCIETE          : TDBEdit;
    TSO_LIBELLE         : THLabel;
    SO_LIBELLE          : TDBEdit;
    TSO_ADRESSE1        : THLabel;
    SO_ADRESSE1         : TDBEdit;
    SO_ADRESSE2         : TDBEdit;
    SO_ADRESSE3         : TDBEdit;
    TSO_CODEPOSTAL      : THLabel;
    SO_CODEPOSTAL       : TDBEdit;
    SO_VILLE            : TDBEdit;
    TSO_VILLE           : THLabel;
    SO_DIVTERRIT        : TDBEdit;
    TSO_DIVTERRIT       : THLabel;
    TSO_TELEPHONE       : THLabel;
    SO_TELEPHONE        : TDBEdit;
    HLFax               : THLabel;
    SO_FAX              : TDBEdit;
    SO_TELEX            : TDBEdit;
    TSO_TELEX           : THLabel;
    TSO_CONTACT         : THLabel;
    SO_CONTACT          : TDBEdit;
    SO_PAYS             : THDBValComboBox;
    TSO_PAYS            : THLabel;
    TSO_NIF             : THLabel;
    TSO_SIRET           : THLabel;
    TSO_APE             : THLabel;
    TSO_CAPITAL         : THLabel;
    TSO_RC              : THLabel;
    SO_RC               : TDBEdit;
    SO_CAPITAL          : TDBEdit;
    SO_APE              : TDBEdit;
    SO_SIRET            : TDBEdit;
    SO_NIF              : TDBEdit;
    SO_TXTJURIDIQUE     : TDBEdit;
    TSO_TXTJURIDIQUE    : THLabel;
    TSO_RVA             : THLabel;
    SO_RVA              : TDBEdit;
    SO_MAIL             : TDBEdit;
    TSO_MAIL            : THLabel;
    HGroupBox1          : TGroupBox;
    TSO_NBJECRAVANT     : THLabel;
    TSO_NBJECRAPRES     : THLabel;
    SO_NBJECRAVANT      : THDBSpinEdit;
    SO_NBJECRAPRES      : THDBSpinEdit;
    GbBil      : TGroupBox;
    TSO_BILDEB1: THLabel;
    TSO_BILDEB2: THLabel;
    TSO_BILDEB3: THLabel;
    TSO_BILDEB4: THLabel;
    TSO_BILDEB5: THLabel;
    TSO_BILFIN1: THLabel;
    TSO_BILFIN2: THLabel;
    TSO_BILFIN3: THLabel;
    TSO_BILFIN4: THLabel;
    TSO_BILFIN5: THLabel;
    SO_BILDEB1 : THDBCpteEdit;
    SO_BILDEB2 : THDBCpteEdit;
    SO_BILDEB3 : THDBCpteEdit;
    SO_BILDEB4 : THDBCpteEdit;
    SO_BILDEB5 : THDBCpteEdit;
    SO_BILFIN1 : THDBCpteEdit;
    SO_BILFIN2 : THDBCpteEdit;
    SO_BILFIN3 : THDBCpteEdit;
    SO_BILFIN4 : THDBCpteEdit;
    SO_BILFIN5 : THDBCpteEdit;
    GbProd     : TGroupBox;
    TSO_PRODEB1: THLabel;
    TSO_PRODEB2: THLabel;
    TSO_PRODEB3: THLabel;
    TSO_PRODEB4: THLabel;
    TSO_PRODEB5: THLabel;
    TSO_PROFIN1: THLabel;
    TSO_PROFIN2: THLabel;
    TSO_PROFIN3: THLabel;
    TSO_PROFIN4: THLabel;
    TSO_PROFIN5: THLabel;
    SO_PRODEB1 : THDBCpteEdit;
    SO_PRODEB2 : THDBCpteEdit;
    SO_PRODEB3 : THDBCpteEdit;
    SO_PRODEB4 : THDBCpteEdit;
    SO_PRODEB5 : THDBCpteEdit;
    SO_PROFIN1 : THDBCpteEdit;
    SO_PROFIN2 : THDBCpteEdit;
    SO_PROFIN3 : THDBCpteEdit;
    SO_PROFIN4 : THDBCpteEdit;
    SO_PROFIN5 : THDBCpteEdit;
    GbExt      : TGroupBox;
    TSO_EXTDEB1: THLabel;
    TSO_EXTDEB2: THLabel;
    TSO_EXTFIN1: THLabel;
    TSO_EXTFIN2: THLabel;
    SO_EXTDEB1 : THDBCpteEdit;
    SO_EXTDEB2 : THDBCpteEdit;
    SO_EXTFIN1 : THDBCpteEdit;
    SO_EXTFIN2 : THDBCpteEdit;
    GbChar     : TGroupBox;
    TSO_CHADEB1: THLabel;
    TSO_CHADEB2: THLabel;
    TSO_CHADEB3: THLabel;
    TSO_CHADEB4: THLabel;
    TSO_CHADEB5: THLabel;
    TSO_CHAFIN1: THLabel;
    TSO_CHAFIN2: THLabel;
    TSO_CHAFIN3: THLabel;
    TSO_CHAFIN4: THLabel;
    TSO_CHAFIN5: THLabel;
    SO_CHADEB1 : THDBCpteEdit;
    SO_CHADEB2 : THDBCpteEdit;
    SO_CHADEB3 : THDBCpteEdit;
    SO_CHADEB4 : THDBCpteEdit;
    SO_CHADEB5 : THDBCpteEdit;
    SO_CHAFIN1 : THDBCpteEdit;
    SO_CHAFIN2 : THDBCpteEdit;
    SO_CHAFIN3 : THDBCpteEdit;
    SO_CHAFIN4 : THDBCpteEdit;
    SO_CHAFIN5 : THDBCpteEdit;
    HGroupBox7   : TGroupBox;
    TSO_GENATTEND: THLabel;
    TSO_CLIATTEND: THLabel;
    TSO_FOUATTEND: THLabel;
    SO_GENATTEND : THDBCpteEdit;
    HGroupBox8   : TGroupBox;
    TSO_FERMEBIL : THLabel;
    TSO_OUVREBIL : THLabel;
    TSO_RESULTAT : THLabel;
    TSO_FERMEPERTE: THLabel;
    TSO_OUVREPERTE: THLabel;
    TSO_FERMEBEN  : THLabel;
    TSO_OUVREBEN  : THLabel;
    TSO_JALFERME  : THLabel;
    TSO_JALOUVRE  : THLabel;
    TTSO_GENATTEND       : THLabel;
    TTSO_CLIATTEND       : THLabel;
    TTSO_FOUATTEND       : THLabel;
    MsgBox               : THMsgBox;
    GroupBox4       : TGroupBox;
    TSO_LGCPTEGEN   : THLabel;
    SO_LGCPTEGEN    : THDBSpinEdit;
    SO_LGCPTEAUX    : THDBSpinEdit;
    TSO_LGCPTEAUX   : THLabel;
    TSO_LGMAXBUDGET : THLabel;
    TSO_LGMINBUDGET : THLabel;
    SO_LGMINBUDGET  : THDBSpinEdit;
    SO_LGMAXBUDGET  : THDBSpinEdit;
    TSO_BOURRETIERS : THLabel;
    SO_BOURREAUX    : TDBEdit;
    SO_BOURREGEN    : TDBEdit;
    TSO_BOURREGEN   : THLabel;
    GroupBox3         : TGroupBox;
    SO_CLIATTEND      : THDBCpteEdit;
    SO_FOUATTEND      : THDBCpteEdit;
    SO_RESULTAT       : THDBCpteEdit;
    SO_FERMEPERTE     : THDBCpteEdit;
    SO_OUVREPERTE     : THDBCpteEdit;
    SO_FERMEBEN       : THDBCpteEdit;
    SO_OUVREBEN       : THDBCpteEdit;
    SO_JALFERME       : THDBCpteEdit;
    SO_JALOUVRE       : THDBCpteEdit;
    SO_OUVREBIL       : THDBCpteEdit;
    SO_FERMEBIL       : THDBCpteEdit;
    TTSO_FERMEBIL     : THLabel;
    TTSO_OUVREBIL     : THLabel;
    TTSO_RESULTAT     : THLabel;
    TTSO_FERMEPERTE   : THLabel;
    TTSO_OUVREPERTE   : THLabel;
    TTSO_FERMEBEN     : THLabel;
    TTSO_OUVREBEN     : THLabel;
    TTSO_JALFERME     : THLabel;
    TTSO_JALOUVRE     : THLabel;
    SO_CORSGE1: TDBCheckBox;
    SO_CORSAU1: TDBCheckBox;
    SO_CORSA11: TDBCheckBox;
    SO_CORSA21: TDBCheckBox;
    SO_CORSA31: TDBCheckBox;
    SO_CORSA41: TDBCheckBox;
    SO_CORSA51: TDBCheckBox;
    SO_CORSBU1: TDBCheckBox;
    GLC          : TGroupBox;
    SO_LETDATEFC : THDBValComboBox;
    SO_LETDATERC : THDBValComboBox;
    SO_LETREFFC  : THDBValComboBox;
    SO_LETREFRC  : THDBValComboBox;
    SO_LETEGALC  : TDBCheckBox;
    SO_LETTOLERC : THDBSpinEdit;
    TSO_LETREFFC: THLabel;
    TSO_LETREFRC: THLabel;
    TSO_LETDATEFC: THLabel;
    TSO_LETDATERC: THLabel;
    TSO_LETTOLERC: THLabel;
    HLabel17     : THLabel;
    GroupBox6    : TGroupBox;
    TSO_LETREFFF : THLabel;
    TSO_LETREFRF : THLabel;
    TSO_LETDATEFF: THLabel;
    TSO_LETDATERF: THLabel;
    TSO_LETTOLERF: THLabel;
    HLabel23     : THLabel;
    SO_LETDATEFF : THDBValComboBox;
    SO_LETDATERF : THDBValComboBox;
    SO_LETREFFF  : THDBValComboBox;
    SO_LETREFRF  : THDBValComboBox;
    SO_LETEGALF  : TDBCheckBox;
    SO_LETTOLERF : THDBSpinEdit;
    Bevel1       : TBevel;
    Bevel2       : TBevel;
    GroupBox1    : TGroupBox;
    TSO_DECQTE   : THLabel;
    TSO_DECPRIX  : THLabel;
    SO_DECQTE    : THDBSpinEdit;
    SO_DECPRIX   : THDBSpinEdit;
    BVentil: TToolbarButton97;
    TSO_FURTIVITE       : THLabel;
    SO_FURTIVITE        : TDBEdit;
    SO_SUIVILOG         : TDBCheckBox;
    SO_NATUREJURIDIQUE  : THDBValComboBox;
    TSO_NATUREJURIDIQUE : THLabel;
    HLabel1             : THLabel;
    GroupBox8           : TGroupBox;
    SO_NBJECHAVANT      : THDBSpinEdit;
    SO_NBJECHAPRES      : THDBSpinEdit;
    HLabel2             : THLabel;
    HLabel3             : THLabel;
    GbCpteDef: TGroupBox;
    TDefautCli: THLabel;
    TDefautFou: THLabel;
    TDefautSal: THLabel;
    TDefautDDivers: THLabel;
    TDefautCDivers: THLabel;
    TDefautDivers: THLabel;
    GbBudget: TGroupBox;
    GDivReg: TGroupBox;
    HMTrad: THSystemMenu;
    CbJfer: TGroupBox;
    Cbj3: TCheckBox;
    Cbj2: TCheckBox;
    Cbj4: TCheckBox;
    Cbj5: TCheckBox;
    Cbj6: TCheckBox;
    Cbj7: TCheckBox;
    Cbj1: TCheckBox;
    Bevel5: TBevel;
    Bevel6: TBevel;
    GroupBox5: TGroupBox;
    SO_MONTANTNEGATIF: TDBCheckBox;
    SO_ETABLISDEFAUT: THDBValComboBox;
    TSO_ETABLISDEFAUT: THLabel;
    SO_ETABLISCPTA: TDBCheckBox;
    SO_CORSGE2: TDBCheckBox;
    SO_CORSAU2: TDBCheckBox;
    SO_CORSA12: TDBCheckBox;
    SO_CORSA22: TDBCheckBox;
    SO_CORSA32: TDBCheckBox;
    SO_CORSA42: TDBCheckBox;
    SO_CORSA52: TDBCheckBox;
    SO_CORSBU2: TDBCheckBox;
    HLabel8: THLabel;
    HPlan2: THLabel;
    GroupBox2: TGroupBox;
    TSO_DATECLOTUREPRO: THLabel;
    SO_DATECLOTUREPER: TDBText;
    TSO_DATECLOTUREPER: THLabel;
    SO_DATECLOTUREPRO: TDBText;
    TSO_DATEDERNENTREE: THLabel;
    SO_DATEDERNENTREE: TDBText;
    SO_NUMPLANREF: THDBSpinEdit;
    TSO_NUMPLANREF: THLabel;
    Defaut1: THDBCpteEdit;
    Defaut2: THDBCpteEdit;
    Defaut3: THDBCpteEdit;
    Defaut4: THDBCpteEdit;
    Defaut5: THDBCpteEdit;
    Defaut6: THDBCpteEdit;
    SO_DUPSECTBUD: TDBCheckBox;
    TSO_JALCTRLBUD: THLabel;
    SO_JALCTRLBUD: THDBValComboBox;
    TSO_BOURRELIB: THLabel;
    SO_BOURRELIB: TDBEdit;
    SO_OUITVAENC: TDBCheckBox;
    TSO_LIBRETEXTE0: THLabel;
    SO_LIBRETEXTE0: TDBEdit;
    HLabel4: THLabel;
    SO_TAUXCOUTFITIERS: TDBEdit;
    HLabel5: THLabel;
    TSO_TVAENCAISSEMENT: THLabel;
    SO_TVAENCAISSEMENT: THDBValComboBox;
    SO_COLLCLIENC: TDBEdit;
    TSO_COLLCLIENC: THLabel;
    TSO_COLLFOUENC: THLabel;
    SO_COLLFOUENC: TDBEdit;
    SO_VERIFRIB: TDBCheckBox;
    TSEuro: TTabSheet;
    GBMonnaie: TGroupBox;
    TSO_DEVISEPRINC: THLabel;
    TSO_DECVALEUR: THLabel;
    SO_DEVISEPRINC: THDBValComboBox;
    SO_DECVALEUR: THDBSpinEdit;
    GBConv: TGroupBox;
    GBCarac: TGroupBox;
    SO_TENUEEURO: TDBCheckBox;
    TSO_DATEDEBUTEURO: THLabel;
    SO_DATEDEBUTEURO: TDBEdit;
    TSO_TAUXEURO: THLabel;
    SO_TAUXEURO: TDBEdit;
    TSO_ECCEURODEBIT: THLabel;
    SO_ECCEURODEBIT: THDBCpteEdit;
    HSO_ECCEURODEBIT: THLabel;
    TSO_ECCEUROCREDIT: THLabel;
    SO_ECCEUROCREDIT: THDBCpteEdit;
    HSO_ECCEUROCREDIT: THLabel;
    GBEquil: TGroupBox;
    TSO_REGLEEQUILSAIS: THLabel;
    SO_REGLEEQUILSAIS: THDBValComboBox;
    TSO_JALECARTEURO: THLabel;
    SO_JALECARTEURO: THDBValComboBox;
    DBNav: TDBNavigator;
    FAutoSave: TCheckBox;
    HPB: TToolWindow97;
    Dock: TDock97;
    BAnnuler: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    TSO_DATEBASCULE: THLabel;
    SO_DATEBASCULE: TDBEdit;
    BRecupSISCO: TToolbarButton97;
    TSocieteSO_SOCIETE: TStringField;
    TSocieteSO_LIBELLE: TStringField;
    TSocieteSO_ADRESSE1: TStringField;
    TSocieteSO_ADRESSE2: TStringField;
    TSocieteSO_ADRESSE3: TStringField;
    TSocieteSO_CODEPOSTAL: TStringField;
    TSocieteSO_VILLE: TStringField;
    TSocieteSO_DIVTERRIT: TStringField;
    TSocieteSO_PAYS: TStringField;
    TSocieteSO_TELEPHONE: TStringField;
    TSocieteSO_FAX: TStringField;
    TSocieteSO_TELEX: TStringField;
    TSocieteSO_NIF: TStringField;
    TSocieteSO_APE: TStringField;
    TSocieteSO_SIRET: TStringField;
    TSocieteSO_CONTACT: TStringField;
    TSocieteSO_LGCPTEGEN: TIntegerField;
    TSocieteSO_LGCPTEAUX: TIntegerField;
    TSocieteSO_BOURREGEN: TStringField;
    TSocieteSO_BOURREAUX: TStringField;
    TSocieteSO_VERIFRIB: TStringField;
    TSocieteSO_MONTANTNEGATIF: TStringField;
    TSocieteSO_SOLDEGEN: TStringField;
    TSocieteSO_NBJECRAVANT: TIntegerField;
    TSocieteSO_NBJECRAPRES: TIntegerField;
    TSocieteSO_NBJECHAVANT: TIntegerField;
    TSocieteSO_NBJECHAPRES: TIntegerField;
    TSocieteSO_TVAENCAISSEMENT: TStringField;
    TSocieteSO_BILDEB1: TStringField;
    TSocieteSO_BILDEB2: TStringField;
    TSocieteSO_BILDEB3: TStringField;
    TSocieteSO_BILDEB4: TStringField;
    TSocieteSO_BILDEB5: TStringField;
    TSocieteSO_BILFIN1: TStringField;
    TSocieteSO_BILFIN2: TStringField;
    TSocieteSO_BILFIN3: TStringField;
    TSocieteSO_BILFIN4: TStringField;
    TSocieteSO_BILFIN5: TStringField;
    TSocieteSO_CHADEB1: TStringField;
    TSocieteSO_CHADEB2: TStringField;
    TSocieteSO_CHADEB3: TStringField;
    TSocieteSO_CHADEB4: TStringField;
    TSocieteSO_CHADEB5: TStringField;
    TSocieteSO_CHAFIN1: TStringField;
    TSocieteSO_CHAFIN2: TStringField;
    TSocieteSO_CHAFIN3: TStringField;
    TSocieteSO_CHAFIN4: TStringField;
    TSocieteSO_CHAFIN5: TStringField;
    TSocieteSO_PRODEB1: TStringField;
    TSocieteSO_PRODEB2: TStringField;
    TSocieteSO_PRODEB3: TStringField;
    TSocieteSO_PRODEB4: TStringField;
    TSocieteSO_PRODEB5: TStringField;
    TSocieteSO_PROFIN1: TStringField;
    TSocieteSO_PROFIN2: TStringField;
    TSocieteSO_PROFIN3: TStringField;
    TSocieteSO_PROFIN4: TStringField;
    TSocieteSO_PROFIN5: TStringField;
    TSocieteSO_EXTDEB1: TStringField;
    TSocieteSO_EXTDEB2: TStringField;
    TSocieteSO_EXTFIN1: TStringField;
    TSocieteSO_EXTFIN2: TStringField;
    TSocieteSO_GENATTEND: TStringField;
    TSocieteSO_CLIATTEND: TStringField;
    TSocieteSO_FOUATTEND: TStringField;
    TSocieteSO_FERMEBIL: TStringField;
    TSocieteSO_OUVREBIL: TStringField;
    TSocieteSO_RESULTAT: TStringField;
    TSocieteSO_FERMEPERTE: TStringField;
    TSocieteSO_OUVREPERTE: TStringField;
    TSocieteSO_FERMEBEN: TStringField;
    TSocieteSO_OUVREBEN: TStringField;
    TSocieteSO_JALFERME: TStringField;
    TSocieteSO_JALOUVRE: TStringField;
    TSocieteSO_DATECLOTUREPER: TDateTimeField;
    TSocieteSO_DATECLOTUREPRO: TDateTimeField;
    TSocieteSO_DATEDERNENTREE: TDateTimeField;
    TSocieteSO_NATUREJURIDIQUE: TStringField;
    TSocieteSO_CAPITAL: TFloatField;
    TSocieteSO_RC: TStringField;
    TSocieteSO_RVA: TStringField;
    TSocieteSO_MAIL: TStringField;
    TSocieteSO_TXTJURIDIQUE: TStringField;
    TSocieteSO_ETABLISCPTA: TStringField;
    TSocieteSO_ETABLISDEFAUT: TStringField;
    TSocieteSO_LGMAXBUDGET: TIntegerField;
    TSocieteSO_LGMINBUDGET: TIntegerField;
    TSocieteSO_CORSGE1: TStringField;
    TSocieteSO_CORSGE2: TStringField;
    TSocieteSO_CORSAU1: TStringField;
    TSocieteSO_CORSAU2: TStringField;
    TSocieteSO_CORSA11: TStringField;
    TSocieteSO_CORSA12: TStringField;
    TSocieteSO_CORSA21: TStringField;
    TSocieteSO_CORSA22: TStringField;
    TSocieteSO_CORSA31: TStringField;
    TSocieteSO_CORSA32: TStringField;
    TSocieteSO_CORSA41: TStringField;
    TSocieteSO_CORSA42: TStringField;
    TSocieteSO_CORSA51: TStringField;
    TSocieteSO_CORSA52: TStringField;
    TSocieteSO_CORSBU1: TStringField;
    TSocieteSO_CORSBU2: TStringField;
    TSocieteSO_DEVISEPRINC: TStringField;
    TSocieteSO_FURTIVITE: TIntegerField;
    TSocieteSO_RECUPCPTA: TStringField;
    TSocieteSO_DECQTE: TIntegerField;
    TSocieteSO_DECVALEUR: TIntegerField;
    TSocieteSO_DECPRIX: TIntegerField;
    TSocieteSO_SUIVILOG: TStringField;
    TSocieteSO_LETREFFC: TStringField;
    TSocieteSO_LETREFFF: TStringField;
    TSocieteSO_LETREFRC: TStringField;
    TSocieteSO_LETREFRF: TStringField;
    TSocieteSO_LETDATEFC: TStringField;
    TSocieteSO_LETDATEFF: TStringField;
    TSocieteSO_LETDATERC: TStringField;
    TSocieteSO_LETDATERF: TStringField;
    TSocieteSO_LETEGALC: TStringField;
    TSocieteSO_LETEGALF: TStringField;
    TSocieteSO_LETTOLERC: TIntegerField;
    TSocieteSO_LETTOLERF: TIntegerField;
    TSocieteSO_EXOV8: TStringField;
    TSocieteSO_DATEREVISION: TDateTimeField;
    TSocieteSO_OUITVAENC: TStringField;
    TSocieteSO_TAUXCOUTFITIERS: TFloatField;
    TSocieteSO_NUMPLANREF: TIntegerField;
    TSocieteSO_ETAPES: TStringField;
    TSocieteSO_JOURFERMETURE: TStringField;
    TSocieteSO_TAUXEURO: TFloatField;
    TSocieteSO_VERSIONBASE: TIntegerField;
    TSocieteSO_DEFCOLCLI: TStringField;
    TSocieteSO_DEFCOLFOU: TStringField;
    TSocieteSO_DEFCOLSAL: TStringField;
    TSocieteSO_DEFCOLDDIV: TStringField;
    TSocieteSO_DEFCOLCDIV: TStringField;
    TSocieteSO_DEFCOLDIV: TStringField;
    TSocieteSO_COLLCLIENC: TStringField;
    TSocieteSO_COLLFOUENC: TStringField;
    TSocieteSO_ETAPEIMPORT: TStringField;
    TSocieteSO_LIBRETEXTE0: TStringField;
    TSocieteSO_DUPSECTBUD: TStringField;
    TSocieteSO_PARAMTABLE: TStringField;
    TSocieteSO_JALCTRLBUD: TStringField;
    TSocieteSO_BOURRELIB: TStringField;
    TSocieteSO_TENUEEURO: TStringField;
    TSocieteSO_DATEDEBUTEURO: TDateTimeField;
    TSocieteSO_ECCEURODEBIT: TStringField;
    TSocieteSO_ECCEUROCREDIT: TStringField;
    TSocieteSO_JALECARTEURO: TStringField;
    TSocieteSO_REGLEEQUILSAIS: TStringField;
    TSocieteSO_DATEBASCULE: TDateTimeField;
    procedure FormShow(Sender: TObject);
    procedure SO_CODEPOSTALExit(Sender: TObject);
    procedure SO_VILLEExit(Sender: TObject);
    procedure SO_CODEPOSTALDblClick(Sender: TObject);
    procedure SO_VILLEDblClick(Sender: TObject);
    procedure SO_ETABLISCPTAClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure TSocieteAfterPost(DataSet: TDataSet);
    procedure SO_BILDEB1Exit(Sender: TObject);
    procedure SO_DEVISEPRINCChange(Sender: TObject);
    procedure Cbj2MouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SO_DIVTERRITDblClick(Sender: TObject);
    procedure SO_PAYSChange(Sender: TObject);
    procedure Defaut1Exit(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SO_DATEDEBUTEUROExit(Sender: TObject);
    procedure BRecupSISCOClick(Sender: TObject);
  private
    Mode   : TActionFiche ;
    LaPage : Integer ;
    Verification : Boolean ;
    LgGen,LgAux,MaxBud,MinBud,DecDev : Integer ;
    PariteEuro : double ;
    CodeDev : String3 ;
    MemoFcpteBil,MemoFcptePro,MemoFcpteCha,MemoFcpteExt : TabCpte ;
    FcpteBil,FcptePro,FcpteCha,FcpteExt : TabCpte ;
    StFichier : String ;
    Function  OnSauve : boolean ;
    Function  EnregOK : boolean ;
    Function  Bouge(Button: TNavigateBtn) : boolean ;
    Procedure ChargeEnreg ;
    Procedure ActiveLaPage(Sender : TObject ; QuellePage : Integer) ;
    Function  LongueurCpte(Sender : TObject) : Boolean ;
    Function  LongueurCpteOk : Boolean ;
    Procedure MemoChamps ;
    Procedure RempliTabloCpte ;
    Function  ChangementTabloCpte : Boolean ;
    Function  FourchetteVide(Sender : TObject) : Boolean ;
    Function  FourchetteOk : Boolean ;
    Function  GeneAttenteOk : Boolean ;
    Function  ExisteCptesAttente : boolean ;
    Procedure GetCompteFourchette(Sender : TObject ; Var Cpte : TabCpte) ;
    Function  ChevauchementCpteFourchetteOk : Boolean ;
    Function  CompareTabloCpteOk(Tab1,Tab2 : TabCpte) : Boolean ;
    Function  ManqueUnCpte : Boolean ;
    Function  NatureCpteOk : Boolean ;
    Function  VerifiNatureCpteOk(Sender : TObject ; Tab : TabCpte) : Boolean ;
    Procedure BourreLeCpte(Sender : TObject) ;
    Procedure RecupDecDev ;
    Function  EtabDefautOk : Boolean ;
    Procedure BlocageDevise ;
    Function  TrouveUnEnreg(Q : TQuery ; Sql : String) : Boolean ;
    Function  DevisePrincEnMvt : Boolean ;
    Procedure EcritJourFermeture ;
    Procedure LitJourFermeture ;
    Function  DefautCpteOk : Boolean ;
    Function  FaitLeWhere(St : String) : String ;
    Function  ExisteLeCompte(LeWhere : String) : Boolean ;
    Function  VerifDateEntreeEuroOk : Boolean ;
    function  VerifPariteEuro : boolean ;
    function  VerifCptesEuro : boolean ;
    Function  CarBourreOk : Boolean ;
    procedure CacheFonctions35 ;
  public  { Public-déclarations }
  end;


implementation

{$R *.DFM}

uses PrintDBG, CPAxe_TOM, QRE, HStatus, ImpFicU ;

Procedure FicheSocieteSISCO ( Comment : TActionFiche ; QuellePage : integer ; Verif : Boolean ; StFichier : String ) ;
var FSociete: TFSocSISCO;
    PP       : THPanel ;
begin
if Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage'],True,'nrCloture') then Exit ;
SourisSablier ;
FSociete:=TFSocSISCO.Create(Application) ;
FSociete.Mode:=Comment ;
FSociete.LaPage:=QuellePage ;
FSociete.Verification:=Verif ;
FSociete.StFichier:=StFichier ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
    FSociete.ShowModal ;
   finally
    FSociete.Free ;
    Bloqueur('nrCloture',False) ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FSociete,PP) ;
   FSociete.Show ;
   END ;
end ;

procedure TFSocSISCO.BImprimerClick(Sender: TObject);
begin
//PrintDBGrid(Nil,Nil,Caption,'PRT_SOCIETE') ;
if ((EstSerie(S3)) or (EstSerie(S5))) Then PrintPageDeGarde(Pages,TRUE,TRUE,FALSE,Caption,101)
                                          Else PrintPageDeGarde(Pages,TRUE,TRUE,FALSE,Caption,1) ;
end;

procedure TFSocSISCO.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin CanClose:=OnSauve ; end;

procedure TFSocSISCO.BValiderClick(Sender: TObject);
begin Bouge(nbPost) ; end;

procedure TFSocSISCO.BAnnulerClick(Sender: TObject);
begin Bouge(nbCancel) ; end;

Function TFSocSISCO.OnSauve : boolean ;
Var Rep : Integer ;
BEGIN
Result:=FALSE  ; NextControl(Self) ;
if TSociete.Modified then
   BEGIN
   if FAutoSave.Checked then Rep:=mrYes else Rep:=MsgBox.Execute(0,'','') ;
   END else Rep:=321 ;
Case rep of
  mrYes : if Not Bouge(nbPost)   then Exit ;
  mrNo  : if Not Bouge(nbCancel) then Exit ;
  mrCancel : Exit ;
  end ;
Result:=TRUE  ;
end ;

Function TFSocSISCO.EnregOK : boolean ;
BEGIN
Result:=FALSE  ; NextControl(Self) ;
if TSociete.state in [dsEdit]=False then Exit ;
if TSociete.state in [dsEdit] then
   BEGIN
   if TSocieteSO_LIBELLE.AsString='' then BEGIN ActiveLaPage(SO_LIBELLE,0) ; MsgBox.Execute(1,'','') ; Exit ; END ;
   if TSocieteSO_ADRESSE1.AsString='' then BEGIN ActiveLaPage(SO_ADRESSE1,0) ; MsgBox.Execute(3,'','') ; Exit ; END ;
   RempliTabloCpte ;
   if Not LongueurCpteOk then Exit ;
   if Not GeneAttenteOk  then Exit ;
//   if not ExisteCptesAttente then Exit ;
   if ChangementTabloCpte then
      BEGIN
      if Not FourchetteOk then Exit ;
      if ManqueUnCpte then Exit ;
      if Not ChevauchementCpteFourchetteOk then Exit ;
      if Not NatureCpteOk then Exit ;
      END ;
   if Not DefautCpteOk then Exit ;
   if Not VerifDateEntreeEuroOk then Exit ;
   if not VerifPariteEuro then Exit ;
   if not VerifCptesEuro then Exit ;
   if Not CarBourreOk then Exit ;
   END ;
// GP 10/01/97 if (SO_ETABLISCPTA.Checked=False) then TSocieteSO_ETABLISDEFAUT.AsString:='' ;
If Not EtabDefautOk Then Exit ;
EcritJourFermeture ;
Result:=TRUE  ;
END ;

Procedure TFSocSISCO.RecupDecDev ;
var Q : TQuery ;
BEGIN
Q:=OpenSQL('SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="'+SO_DEVISEPRINC.Value+'"',True) ;
If Not Q.EOF then SO_DECVALEUR.Value:=Q.Fields[0].AsInteger ;
Ferme(Q) ;
END ;

procedure TFSocSISCO.BFermeClick(Sender: TObject);
begin Close ; end;

Procedure TFSocSISCO.ChargeEnreg ;
BEGIN
InitCaption(Self,SO_SOCIETE.text,SO_LIBELLE.text) ;
SO_SOCIETE.Enabled:=False ;
TSocieteSO_CAPITAL.DisplayFormat:=StrfMask(V_PGI.OkdecV,'',True) ;
SO_GENATTEND.ExisteH  ; SO_CLIATTEND.ExisteH  ; SO_FOUATTEND.ExisteH ;
SO_FERMEBIL.ExisteH   ; SO_OUVREBIL.ExisteH   ; SO_RESULTAT.ExisteH ;
SO_OUVREPERTE.ExisteH ; SO_FERMEPERTE.ExisteH ; SO_FERMEBEN.ExisteH ;
SO_OUVREBEN.ExisteH   ; SO_JALFERME.ExisteH   ; SO_JALOUVRE.ExisteH   ;
SO_ECCEURODEBIT.ExisteH ; SO_ECCEUROCREDIT.ExisteH ;
SO_DEVISEPRINCChange(Nil) ; LitJourFermeture ;
if Trim(SO_BOURRELIB.Text)='' then SO_BOURRELIB.Text:='.' ;
end;

Function TFSocSISCO.Bouge(Button: TNavigateBtn) : boolean ;
BEGIN
result:=FALSE  ;
Case Button of
   nblast,nbprior,nbnext,
   nbfirst,nbinsert : Exit ;
   nbPost           : if Not EnregOK then Exit ;
   nbDelete         : Exit ;
   end ;
if not TransacNav(DBNav.BtnClick,Button,10) then MessageAlerte(MsgBox.Mess[2]) ;
if Button=nbPost then BEGIN ChargeSocieteHalley ; MemoChamps ; END ;
result:=TRUE ;
END ;

Procedure TFSocSISCO.MemoChamps ;
BEGIN
LgGen:=TSocieteSO_LGCPTEGEN.AsInteger ; LgAux:=TSocieteSO_LGCPTEAUX.AsInteger ;
MaxBud:=TSocieteSO_LGMAXBUDGET.AsInteger ; MinBud:=TSocieteSO_LGMINBUDGET.AsInteger ;
CodeDev:=TSocieteSO_DEVISEPRINC.AsString ;
DecDev:=TSocieteSO_DECVALEUR.AsInteger ;
PariteEuro:=TSocieteSO_TAUXEURO.AsFloat ;
FillChar(MemoFcpteBil,SizeOf(MemoFcpteBil),#0) ; FillChar(MemoFcptePro,SizeOf(MemoFcptePro),#0) ;
FillChar(MemoFcpteCha,SizeOf(MemoFcpteCha),#0) ; FillChar(MemoFcpteExt,SizeOf(MemoFcpteExt),#0) ;
GetCompteFourchette(GbBil,MemoFcpteBil)  ; GetCompteFourchette(GbProd,MemoFcptePro) ;
GetCompteFourchette(GbChar,MemoFcpteCha) ; GetCompteFourchette(GbExt,MemoFcpteExt)  ;
END ;

procedure TFSocSISCO.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TSociete.Close ; TSociete.Open ;
Pages.ActivePage:=Pages.Pages[LaPage] ;
MemoChamps ;
if Mode=taConsult then BEGIN FicheReadOnly(Self) ; Exit ; END ;
BlocageDevise ; ChargeEnreg ;
SO_TAUXEURO.Enabled:=(V_PGI.DateEntree<V_PGI.DateDebutEuro) ;
TSO_TAUXEURO.Enabled:=SO_TAUXEURO.Enabled ;
SourisNormale ;
CacheFonctions35 ;
SO_TENUEEURO.Enabled:=False ;
if SO_TENUEEURO.Checked then SO_TAUXEURO.Enabled:=False ; 
if SO_JALCTRLBUD.Vide then SO_JALCTRLBUD.Items[0]:='' ;
if VH^.TenueEuro then SO_DATEDEBUTEURO.Enabled:=False ;
If StFichier<>'' Then BRecupSISCO.Visible:=TRUE ;
end;

Procedure TFSocSISCO.ActiveLaPage(Sender : TObject ; QuellePage : Integer) ;
BEGIN Pages.ActivePage:=Pages.Pages[QuellePage] ; TWinControl(Sender).SetFocus ; END ;

procedure TFSocSISCO.SO_CODEPOSTALExit(Sender: TObject);
begin VerifCodePostal(TSociete,SO_CODEPOSTAL,SO_VILLE,TRUE) ; end;

procedure TFSocSISCO.SO_VILLEExit(Sender: TObject);
begin VerifCodePostal(TSociete,SO_CODEPOSTAL,SO_VILLE,TRUE) ; (*FALSE*) end;

procedure TFSocSISCO.SO_CODEPOSTALDblClick(Sender: TObject);
begin VerifCodePostal(TSociete,SO_CODEPOSTAL,SO_VILLE,True) ; end;

procedure TFSocSISCO.SO_VILLEDblClick(Sender: TObject);
begin VerifCodePostal(TSociete,SO_CODEPOSTAL,SO_VILLE,False) ; end;

procedure TFSocSISCO.SO_ETABLISCPTAClick(Sender: TObject);
begin
// GP 10/01/97
//TSO_ETABLISDEFAUT.Visible:=SO_ETABLISCPTA.Checked ;
//SO_ETABLISDEFAUT.Visible:=SO_ETABLISCPTA.Checked ;
TSO_ETABLISDEFAUT.Visible:=TRUE ;
SO_ETABLISDEFAUT.Visible:=TRUE ;
end;

procedure TFSocSISCO.BVentilClick(Sender: TObject);
begin FicheAxeAna ('') ; end;

Procedure TFSocSISCO.RempliTabloCpte ;
BEGIN
FillChar(FcpteBil,SizeOf(FcpteBil),#0) ; FillChar(FcptePro,SizeOf(FcptePro),#0) ;
FillChar(FcpteCha,SizeOf(FcpteCha),#0) ; FillChar(FcpteExt,SizeOf(FcpteExt),#0) ;
GetCompteFourchette(GbBil,FcpteBil)  ; GetCompteFourchette(GbProd,FcptePro) ;
GetCompteFourchette(GbChar,FcpteCha) ; GetCompteFourchette(GbExt,FcpteExt)  ;
END ;

Function TFSocSISCO.ChangementTabloCpte : Boolean ;
Var i : Integer ;
BEGIN
Result:=False ; i:=1 ;
While i<=10 do
   if(MemoFcpteBil[i]<>FcpteBil[i])Or(MemoFcptePro[i]<>FcptePro[i])Or
     (MemoFcpteCha[i]<>FcpteCha[i])Or(MemoFcpteExt[i]<>FcpteExt[i])then
      BEGIN Result:=True ; Exit ; END else Inc(i) ;
END ;

Function TFSocSISCO.LongueurCpte(Sender : TObject) : Boolean ;
Var Sql : String ;
    Q : TQuery ;
BEGIN
Result:=True ;
if(THDBSpinEdit(Sender).Name='SO_LGCPTEGEN')then Sql:='Select G_GENERAL From GENERAUX' else
  if(THDBSpinEdit(Sender).Name='SO_LGCPTEAUX')then Sql:='Select T_AUXILIAIRE From TIERS' else
     if(THDBSpinEdit(Sender).Name='SO_LGMAXBUDGET')Or
       (THDBSpinEdit(Sender).Name='SO_LGMINBUDGET') then Sql:='Select BG_BUDGENE From BUDGENE' ;
Q:=OpenSql(Sql,True) ;
if Not Q.Eof then
   BEGIN
   ActiveLaPage(Sender,1) ; MsgBox.Execute(4,'','') ; Result:=False ;
   END ;
Ferme(Q) ;
END ;

Function TFSocSISCO.LongueurCpteOk : Boolean ;
BEGIN
Result:=False ;
if TSocieteSO_LGCPTEGEN.AsInteger<>LgGen then if Not LongueurCpte(SO_LGCPTEGEN)   then Exit ;
if TSocieteSO_LGCPTEAUX.AsInteger<>LgAux then if Not LongueurCpte(SO_LGCPTEAUX)   then Exit ;
if TSocieteSO_LGMAXBUDGET.AsInteger<>MaxBud then if Not LongueurCpte(SO_LGMAXBUDGET) then Exit ;
if TSocieteSO_LGMINBUDGET.AsInteger<>MinBud then if Not LongueurCpte(SO_LGMINBUDGET) then Exit ;
Result:=True ;
END ;

Function TFSocSISCO.GeneAttenteOk : Boolean ;
Var Q : TQuery ;
    PasOk : Boolean ;
    Ind : Byte ;
BEGIN
Result:=False ; PasOk:=FALSE ; Ind:=0 ;
if TSocieteSO_GENATTEND.AsString='' then BEGIN Result:=True ; Exit ; END ;
Q:=OpenSql('Select G_GENERAL,G_VENTILABLE,G_COLLECTIF,G_LETTRABLE,G_POINTABLE From GENERAUX '+
           'Where G_GENERAL="'+TSocieteSO_GENATTEND.AsString+'"',True) ;
if Q.Fields[0].AsString='' then BEGIN Result:=True ; Ferme(Q) ; Exit ; END else
   BEGIN
   PasOk:=True ;
   if(Q.Fields[1].AsString='X')And(PasOk) then BEGIN PasOk:=False ; Ind:=5 ; END ;
   if(Q.Fields[2].AsString='X')And(PasOk) then BEGIN PasOk:=False ; Ind:=6 ; END ;
   if(Q.Fields[3].AsString='X')And(PasOk) then BEGIN PasOk:=False ; Ind:=7 ; END ;
   if(Q.Fields[4].AsString='X')And(PasOk) then BEGIN PasOk:=False ; Ind:=8 ; END ;
   END ;
Ferme(Q) ;
if Not PasOk then BEGIN ActiveLaPage(SO_GENATTEND,3) ; MsgBox.Execute(Ind,'','') ; Exit ; END ;
Result:=True ;
END ;

function TFSocSISCO.ExisteCptesAttente : boolean ;
BEGIN
Result:=False ;
if not Verification then begin Result := True ; exit ; end ;
if (SO_GENATTEND.ExisteH<=0) then BEGIN MsgBox.Execute(13,'','') ; Exit ; END ;
if (SO_CLIATTEND.ExisteH<=0) then BEGIN MsgBox.Execute(14,'','') ; Exit ; END ;
if (SO_FOUATTEND.ExisteH<=0) then BEGIN MsgBox.Execute(15,'','') ; Exit ; END ;
Result:=True ;
END ;

Function TFSocSISCO.FourchetteVide(Sender : TObject) : Boolean ;
Var i : Integer ;
    Pref,Indice : String ;
BEGIN
Result:=True ;
if TGroupBox(Sender).Name='GbBil' then Pref:='BIL' else
   if TGroupBox(Sender).Name='GbProd' then Pref:='PRO' else
      if TGroupBox(Sender).Name='GbChar' then Pref:='CHA' else
          if TGroupBox(Sender).Name='GbExt' then Pref:='EXT' ;
for i:=0 to TGroupBox(Sender).ControlCount-1 do
    BEGIN
    if TGroupBox(Sender).Controls[i] is THDBCpteEdit then
        BEGIN
        if(THDBCpteEdit(TGroupBox(Sender).Controls[i]).Text)<>'' then
           BEGIN
           Indice:=Copy(THDBCpteEdit(TGroupBox(Sender).Controls[i]).Name,
                        Length(THDBCpteEdit(TGroupBox(Sender).Controls[i]).Name),1) ;
           if(THDBCpteEdit(FindComponent('SO_'+Pref+'DEB'+Indice+'')).Text<>'')And
             (THDBCpteEdit(FindComponent('SO_'+Pref+'FIN'+Indice+'')).Text<>'')then
             BEGIN Result:=False ; Exit ; END ;
           END ;
        END ;
    END ;
Case Pref[1] of
     'B' : ActiveLaPage(SO_BILDEB1,2) ;
     'C' : ActiveLaPage(SO_CHADEB1,2) ;
     'E' : ActiveLaPage(SO_EXTDEB1,2) ;
     'P' : ActiveLaPage(SO_PRODEB1,2) ;
    End ;
MsgBox.Execute(9,'','') ;
END ;

Function TFSocSISCO.FourchetteOk : Boolean ;
BEGIN
Result:=False ;
if FourchetteVide(GbBil) then Exit ;
if FourchetteVide(GbProd)then Exit ;
if FourchetteVide(GbChar)then Exit ;
Result:=True ;
END ;

procedure TFSocSISCO.TSocieteAfterPost(DataSet: TDataSet);
begin
if ((TSocieteSO_DECVALEUR.AsInteger<>DecDev) or (TSocieteSO_TAUXEURO.AsFloat<>PariteEuro)) then
   ExecuteSql('Update DEVISE SET D_DECIMALE='+IntToStr(TSocieteSO_DECVALEUR.AsInteger)+', '+
              'D_PARITEEURO='+StrfPoint(TSocieteSO_TAUXEURO.AsFloat)+' '+
              'Where D_DEVISE="'+TSocieteSO_DEVISEPRINC.AsString+'"') ;
end;

Procedure TFSocSISCO.GetCompteFourchette(Sender : TObject ; Var Cpte : TabCpte) ;
Var i,Max,Indice : Integer ;
    Pref : String ;
BEGIN
Max:=0 ;
for i:=0 to TGroupBox(Sender).ControlCount-1 do
   if TGroupBox(Sender).Controls[i] is THDBCpteEdit then Inc(Max) ;

if TGroupBox(Sender).Name='GbBil' then Pref:='BIL' else
   if TGroupBox(Sender).Name='GbProd' then Pref:='PRO' else
      if TGroupBox(Sender).Name='GbChar' then Pref:='CHA' else
          if TGroupBox(Sender).Name='GbExt' then Pref:='EXT' ;
Indice:=1 ; i:=1 ;
While i<Max do
    BEGIN
    Cpte[i]:=THDBCpteEdit(FindComponent('SO_'+Pref+'DEB'+IntToStr(Indice)+'')).Text ;
    Cpte[i+1]:=THDBCpteEdit(FindComponent('SO_'+Pref+'FIN'+IntToStr(Indice)+'')).Text ;
    Inc(Indice) ; i:=i+2 ;
    END ;
END ;

Function TFSocSISCO.ChevauchementCpteFourchetteOk : Boolean ;
Var Ok : Boolean ;
BEGIN
Ok:=True ;
if Not CompareTabloCpteOk(FcpteBil,FcptePro) then BEGIN ActiveLaPage(SO_PRODEB1,2) ; Ok:=False ; END ;
if Not CompareTabloCpteOk(FcpteBil,FcpteCha) then BEGIN ActiveLaPage(SO_CHADEB1,2) ; Ok:=False ; END ;
if Not CompareTabloCpteOk(FcpteBil,FcpteExt) then BEGIN ActiveLaPage(SO_EXTDEB1,2) ; Ok:=False ; END ;
if Not CompareTabloCpteOk(FcptePro,FcpteCha) then BEGIN ActiveLaPage(SO_CHADEB1,2) ; Ok:=False ; END ;
if Not CompareTabloCpteOk(FcptePro,FcpteExt) then BEGIN ActiveLaPage(SO_EXTDEB1,2) ; Ok:=False ; END ;
if Not CompareTabloCpteOk(FcpteCha,FcpteExt) then BEGIN ActiveLaPage(SO_EXTDEB1,2) ; Ok:=False ; END ;
Result:=Ok ;
if Not Ok then MsgBox.Execute(11,'','') ;
END ;

Function TFSocSISCO.CompareTabloCpteOk(Tab1,Tab2 : TabCpte) : Boolean ;
Var i,j : Integer ;
BEGIN
Result:=True ; i:=1 ;
While i<10 do
 BEGIN
 if(Tab1[i]<>'')And(Tab1[i+1]<>'') then
   BEGIN
   j:=1 ;
   While j<10 do
    BEGIN
    if(Tab2[j]<>'')And(Tab2[j+1]<>'') then
       BEGIN
       if Tab2[j]=Tab1[i]     then BEGIN Result:=False ; Exit ; END ;
       if Tab2[j+1]=Tab1[i+1] then BEGIN Result:=False ; Exit ; END ;
       if Tab2[j]<Tab1[i] then
         if Tab2[j+1]>Tab1[i]then BEGIN Result:=False ; Exit ; END ;
       if Tab2[j]>Tab1[i] then
          if Tab2[j]<Tab1[i+1] then BEGIN Result:=False ; Exit ; END ;
       END ;
    j:=j+2 ;
    END ;
   END ;
 i:=i+2 ;
 END ;
END ;

Function TFSocSISCO.ManqueUnCpte : Boolean ;
Var i,Indice,LeGb : Integer ;
    ManqueCpte : Boolean ;
    LeFocus : TObject ;
BEGIN
Result:=False ; i:=1 ; ManqueCpte:=False ; Indice:=0 ; LeGb:=0 ; LeFocus:=Nil ;
While(i<10)And(Not ManqueCpte) do
  BEGIN
  if(((FcpteBil[i]<>'')And(FcpteBil[i+1]=''))Or((FcpteBil[i]='')And(FcpteBil[i+1]<>''))) then
    BEGIN ManqueCpte:=True ; LeGb:=1 ; Indice:=i ; Break ; END ;
  if(((FcptePro[i]<>'')And(FcptePro[i+1]=''))Or((FcptePro[i]='')And(FcptePro[i+1]<>''))) then
    BEGIN ManqueCpte:=True ; LeGb:=2 ; Indice:=i ; Break ;  END ;
  if(((FcpteCha[i]<>'')And(FcpteCha[i+1]=''))Or((FcpteCha[i]='')And(FcpteCha[i+1]<>''))) then
    BEGIN ManqueCpte:=True ; LeGb:=3 ; Indice:=i ; Break ;  END ;
  if(((FcpteExt[i]<>'')And(FcpteExt[i+1]=''))Or((FcpteExt[i]='')And(FcpteExt[i+1]<>''))) then
    BEGIN ManqueCpte:=True ; LeGb:=4 ; Indice:=i ; Break ;  END ;
  i:=i+2 ;
  END ;
if ManqueCpte then
  BEGIN
  Case Indice of
       1:BEGIN
          Case LeGb of
               1:LeFocus:=SO_BILDEB1 ;
               2:LeFocus:=SO_PRODEB1 ;
               3:LeFocus:=SO_CHADEB1 ;
               4:LeFocus:=SO_EXTDEB1 ;
              End ;
         END ;
       3:BEGIN
          Case LeGb of
               1:LeFocus:=SO_BILDEB2 ;
               2:LeFocus:=SO_PRODEB2 ;
               3:LeFocus:=SO_CHADEB2 ;
               4:LeFocus:=SO_EXTDEB2 ;
              End ;
         END ;
       5:BEGIN
          Case LeGb of
               1:LeFocus:=SO_BILDEB3 ;
               2:LeFocus:=SO_PRODEB3 ;
               3:LeFocus:=SO_CHADEB3 ;
              End ;
         END ;
       7:BEGIN
          Case LeGb of
               1:LeFocus:=SO_BILDEB4 ;
               2:LeFocus:=SO_PRODEB4 ;
               3:LeFocus:=SO_CHADEB4 ;
              End ;
         END ;
       9:BEGIN
          Case LeGb of
               1:LeFocus:=SO_BILDEB5 ;
               2:LeFocus:=SO_PRODEB5 ;
               3:LeFocus:=SO_CHADEB5 ;
              End ;
         END ;
   End ;
  Result:=True ; ActiveLaPage(LeFocus,2) ; MsgBox.Execute(10,'','') ;
  END ;
END ;

Function TFSocSISCO.NatureCpteOk : Boolean ;
BEGIN
Result:=False ;
if Not VerifiNatureCpteOk(GbBil,FcpteBil) then Exit ;
if Not VerifiNatureCpteOk(GbProd,FcptePro) then Exit ;
if Not VerifiNatureCpteOk(GbChar,FcpteCha) then Exit ;
if Not VerifiNatureCpteOk(GbExt,FcpteExt) then Exit ;
Result:=True ;
END ;

Function TFSocSISCO.VerifiNatureCpteOk(Sender : TObject ; Tab : TabCpte) : Boolean ;
Var Q : TQuery ;
    i : Integer ;
    Pref : Char ;
    PasOk : Boolean ;
    LeFocus : TObject ;
BEGIN
Result:=True ; PasOk:=False ; i:=1 ;
if TGroupBox(Sender).Name='GbBil' then BEGIN Pref:='B' ; LeFocus:=SO_BILDEB1 ; END else
   if TGroupBox(Sender).Name='GbProd' then BEGIN Pref:='P' ; LeFocus:=SO_PRODEB1 ; END else
      if TGroupBox(Sender).Name='GbChar' then BEGIN Pref:='C' ; LeFocus:=SO_CHADEB1 ; END else
          if TGroupBox(Sender).Name='GbExt' then BEGIN Pref:='E' ; LeFocus:=SO_EXTDEB1 ; END Else Exit ;
While i<10 do
 BEGIN
 if Tab[i]<>'' then
    BEGIN
    Q:=OpenSql('Select G_NATUREGENE,G_GENERAL from GENERAUX Where G_GENERAL>="'+Tab[i]+'" And '+
               'G_GENERAL<="'+Tab[i+1]+'" Order by G_GENERAL',True) ;
    While (Not Q.Eof)And(Not PasOk) do
      BEGIN
      Case Pref of
           'B': if(Q.Fields[0].AsString='CHA')Or(Q.Fields[0].AsString='PRO')Or
                  (Q.Fields[0].AsString='EXT') then PasOk:=True ;
           'P': if(Q.Fields[0].AsString<>'PRO')then PasOk:=True ;
           'C': if(Q.Fields[0].AsString<>'CHA')then PasOk:=True ;
           'E': if(Q.Fields[0].AsString<>'EXT')then PasOk:=True ;
          End ;
      Q.Next ;
      END ;
    Ferme(Q) ;
    if PasOk then
       BEGIN
       ActiveLaPage(LeFocus,2) ; MsgBox.Execute(12,'','') ; Result:=False ; Exit ;
       END ;
    END ;
 i:=i+2 ;
 END ;
END ;

Procedure TFSocSISCO.BourreLeCpte(Sender : TObject );
Var Lefb : TFichierBase ;
BEGIN
if THCpteEdit(Sender).Text='' then Exit ;
LeFb:=CaseFic(THCpteEdit(Sender).ZoomTable) ;
if Length(THCpteEdit(Sender).Text)<VH^.Cpta[Lefb].Lg then
   THCpteEdit(Sender).Text:=BourreLaDonc(THCpteEdit(Sender).Text,Lefb) ;
END ;

procedure TFSocSISCO.SO_BILDEB1Exit(Sender: TObject);
begin BourreLeCpte(Sender) ; end;

Function TFSocSISCO.DevisePrincEnMvt : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select D_DEVISE From DEVISE Where D_DEVISE="'+SO_DEVISEPRINC.Value+'" And '+
              '(Exists (Select E_DEVISE From ECRITURE Where E_DEVISE="'+SO_DEVISEPRINC.Value+'"))',True) ;
Result:=QLoc.Eof ; Ferme(QLoc) ;
END ;

procedure TFSocSISCO.SO_DEVISEPRINCChange(Sender: TObject);
begin RecupDecDev ; SO_DECVALEUR.Enabled:=DevisePrincEnMvt ; end;

Function TFSocSISCO.EtabDefautOk : Boolean ;
BEGIN
Result:=TRUE ;
If ((Verification) and (SO_ETABLISDEFAUT.Value='')) Then BEGIN Result:=FALSE ; MsgBox.Execute(17,'','') ; END ;
END ;

Procedure TFSocSISCO.BlocageDevise ;
Var QLoc : TQuery ;
    Sql : String ;
BEGIN
QLoc:=TQuery.Create(Application) ;
QLoc.DataBaseName:='SOC' ;
Sql:='Select J_JOURNAL From JOURNAL Where Exists(Select E_JOURNAL From ECRITURE)' ;
if TrouveUnEnreg(Qloc,Sql) then
   BEGIN SO_DEVISEPRINC.Enabled:=False ; QLoc.Close ; QLoc.Free ; Exit ; END  ;
Sql:='Select J_JOURNAL From JOURNAL Where Exists(Select Y_JOURNAL From ANALYTIQ)' ;
if TrouveUnEnreg(Qloc,Sql) then
   BEGIN SO_DEVISEPRINC.Enabled:=False ; QLoc.Close ; QLoc.Free ; Exit ; END  ;
Sql:='Select Distinct CU_TYPE From CUMULS Where CU_TYPE<>"CON"' ;
if TrouveUnEnreg(Qloc,Sql) then
   BEGIN SO_DEVISEPRINC.Enabled:=False ; QLoc.Close ; QLoc.Free ; Exit ; END  ;
Sql:='Select J_JOURNAL From JOURNAL Where Exists(Select BE_BUDJAL From BUDECR)' ;
if TrouveUnEnreg(Qloc,Sql) then
   BEGIN SO_DEVISEPRINC.Enabled:=False ; QLoc.Close ; QLoc.Free ; Exit ; END  ;
Sql:='Select EG_TYPE From ECRGUI' ;
if TrouveUnEnreg(Qloc,Sql) then
   BEGIN SO_DEVISEPRINC.Enabled:=False ; QLoc.Close ; QLoc.Free ; Exit ; END  ;
QLoc.Close ; QLoc.Free ;
END ;

Function TFSocSISCO.TrouveUnEnreg(Q : TQuery ;Sql : String) : Boolean ;
BEGIN
Result:=True ; Q.Close ; Q.Sql.Clear ;
Q.Sql.Add(Sql) ; ChangeSql(Q) ; Q.Open ;
if Q.Eof then Result:=False ; Q.Close ;
END ;

procedure TFSocSISCO.Cbj2MouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if TSociete.State in [dsEdit,dsInsert] then Exit ;
TSociete.Edit ;
TSocieteSO_SOCIETE.AsString:=TSocieteSO_SOCIETE.AsString ;
end;

Procedure TFSocSISCO.LitJourFermeture ;
Var i : Byte ;
    C : TComponent ;
    St : String[7] ;
BEGIN
St:=TSocieteSO_JOURFERMETURE.AsString ;
if Length(St)<>7 then Exit ;
for i:=1 to 7 do
   BEGIN
   C:=FindComponent('Cbj'+IntToStr(i)) ;
   TCheckBox(C).Checked:=(St[i]='X') ;
   END ;
END ;

Procedure TFSocSISCO.EcritJourFermeture ;
Var i : Byte ;
    C : TComponent ;
    St : String[7] ;
BEGIN
St:='' ; 
for i:=1 to 7 do
   BEGIN
   C:=FindComponent('Cbj'+IntToStr(i)) ;
   if TCheckBox(C).Checked then St:=St+'X' else St:=St+'-' ;
   END ;
if TSociete.State<>dsEdit then TSociete.Edit ;
TSocieteSO_JOURFERMETURE.AsString:=St ;
END ;

procedure TFSocSISCO.SO_DIVTERRITDblClick(Sender: TObject);
begin
PaysRegion(SO_PAYS,SO_DIVTERRIT,True) ;
end;

procedure TFSocSISCO.SO_PAYSChange(Sender: TObject);
begin
PaysRegion(SO_PAYS,SO_DIVTERRIT,False) ;
end;

procedure TFSocSISCO.Defaut1Exit(Sender: TObject);
begin BourreLeCpte(Sender) ; end;

Function TFSocSISCO.DefautCpteOk : Boolean ;
Var Q : TQuery ;
    i : Byte ;
    Sql,LeWhere : String ;
    St : String ;
BEGIN
Result:=True ; Q:=Tquery.Create(Self) ; Q.DataBaseName:='SOC' ;
for i:=1 to 6 do
   BEGIN
   St:=THCpteEdit(FindComponent('Defaut'+IntToStr(i)+'')).Text ;
   if St='' then Continue ;
   Case i of
        1: LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COC"' ;
        2: LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COF"' ;
        3: LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COS"' ;
        4: LeWhere:='And G_COLLECTIF="X" And (G_NATUREGENE="COC" Or G_NATUREGENE="COD")' ;
        5: LeWhere:='And G_COLLECTIF="X" And (G_NATUREGENE="COF" Or G_NATUREGENE="COD")' ;
        6: LeWhere:='And G_COLLECTIF="X" And G_NATUREGENE="COD"' ;
      End ;
   Sql:='Select G_GENERAL From GENERAUX Where G_GENERAL="'+St+'" '+LeWhere+'' ;
   Q.Close ; Q.Sql.Clear ; Q.Sql.Add(Sql) ; ChangeSql(Q) ; Q.Open ;
   St:=Q.Fields[0].AsString ;
   if St='' then
      BEGIN
      Pages.ActivePage:=PDates ; THCpteEdit(FindComponent('Defaut'+IntToStr(i)+'')).SetFocus ;
      MsgBox.Execute(16,'','') ;  Result:=False ; Break ;
      END ;
   END ;
Ferme(Q) ;
END ;

procedure TFSocSISCO.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

Function TFSocSISCO.FaitLeWhere(St : String) : String ;
Var S1,S2 : String ;
BEGIN
if Pos(':',St)>0 then
   BEGIN
   S1:=Copy(St,1,Pos(':',St)-1) ; S2:=Copy(St,Pos(':',St)+1,Length(St)) ;
   S2:=BourreLaDonc(S2,fbGene) ;
   Result:='Where G_GENERAL >="'+S1+'%" And G_GENERAL <="'+S2+'"' ;
   END else
   BEGIN
   if Length(St)<VH^.Cpta[fbGene].Lg then Result:='Where G_GENERAL Like"'+St+'%"'
                                    else Result:='Where G_GENERAL ="'+St+'"' ;
   END ;
END ;

Function TFSocSISCO.ExisteLeCompte(LeWhere : String) : Boolean ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select G_GENERAL From GENERAUX '+LeWhere,True) ;
Result:=Not QLoc.Eof ; Ferme(QLoc) ;
END ;

Function TFSocSISCO.VerifDateEntreeEuroOk : Boolean ;
BEGIN
Result:=True ;
if Not Verification then Exit ;
if TSocieteSO_DATEDEBUTEURO.AsDateTime<Encodedate(1999,01,01) then BEGIN MsgBox.Execute(20,'','') ; Result:=False ; END ;
if TSocieteSO_DATEDEBUTEURO.AsDateTime>Encodedate(1999,12,31) then
   BEGIN
   if ((SO_DEVISEPRINC.Value<>'') and (V_PGI.DevisePivot<>'') and (Not So_DEVISEPRINC.Enabled) and (Not EstMonnaieIn(SO_DEVISEPRINC.Value)))
      then else BEGIN MsgBox.Execute(20,'','') ; Result:=False ; END ;
   END ;
END ;

function TFSocSISCO.VerifCptesEuro : boolean ;
Var CptD,CptC : String ;
    Okok : Boolean ;
    QQ   : TQuery ;
BEGIN
Result:=True ; Okok:=True ;
if Not Verification then Exit ;
CptD:=TSocieteSO_ECCEURODebit.AsString ;
CptC:=TSocieteSO_ECCEUROCredit.AsString ;
if ((CptD='') and (CptC='')) then Exit ;
if ((CptD<>'') and (SO_ECCEuroDebit.ExisteH<=0)) then
   BEGIN
   Result:=False ; MsgBox.Execute(23,'',''); Exit ;
   END ;
if ((CptC<>'') and (SO_ECCEuroCredit.ExisteH<=0)) then
   BEGIN
   Result:=False ; MsgBox.Execute(23,'',''); Exit ;
   END ;
if CptD<>'' then
   BEGIN
   QQ:=OpenSQL('Select * from GENERAUX Where G_GENERAL="'+CptD+'" AND G_VENTILABLE="X"',True) ;
   if Not QQ.EOF then Okok:=False ;
   Ferme(QQ) ;
   if Not Okok then BEGIN Result:=False ; MsgBox.Execute(24,'',''); Exit ; END ;
   END ;
if CptC<>'' then
   BEGIN
   QQ:=OpenSQL('Select * from GENERAUX Where G_GENERAL="'+CptC+'" AND G_VENTILABLE="X"',True) ;
   if Not QQ.EOF then Okok:=False ;
   Ferme(QQ) ;
   if Not Okok then BEGIN Result:=False ; MsgBox.Execute(24,'',''); Exit ; END ;
   END ;
Result:=Okok ;   
END ;

function TFSocSISCO.VerifPariteEuro : boolean ;
Var StParite : String ;
    Parite   : Integer ;
BEGIN
Result:=True ;
if Not Verification then Exit ;
StParite:=Trim(SO_TAUXEURO.Text) ;
if Pos(V_PGI.SepDecimal,StParite)<>0 then Delete(StParite,Pos(V_PGI.SepDecimal,StParite),1) ;
If StParite<>'' Then Parite:=StrToInt(StParite) Else Parite:=0 ;
if Parite<=0 then BEGIN MsgBox.Execute(22,'','') ; Result:=False ; Exit ; END ;
StParite:=IntToStr(Parite) ;
if Length(StParite)>6 then BEGIN MsgBox.Execute(22,'','') ; Result:=False ; END ;
END ;

Function TFSocSISCO.CarBourreOk : Boolean ;
Var Okok : boolean ;
    C : Char ;
    CC : Set of Char ;
BEGIN
Result:=False ; Okok:=True ;
CC:=[' ','*','%','?','#',':','_',',','|','"','''',';'] ;
if TSocieteSO_BOURREGEN.AsString='' then Okok:=False else
   BEGIN
   C:=TSocieteSO_BOURREGEN.AsString[1] ; if C in CC then Okok:=False ;
   END ;
if Not Okok then
   BEGIN
   MsgBox.Execute(21,'','') ; Pages.ActivePage:=PComptables ; SO_BOURREGEN.SetFocus ; Exit ;
   END ;
if TSocieteSO_BOURREAUX.AsString='' then Okok:=False else
   BEGIN
   C:=TSocieteSO_BOURREAUX.AsString[1] ; if C in CC then Okok:=False ;
   END ;
if Not Okok then
   BEGIN
   MsgBox.Execute(21,'','') ; Pages.ActivePage:=PComptables ; SO_BOURREAUX.SetFocus ; Exit ;
   END ;
Result:=True ;
END ;

procedure TFSocSISCO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   Bloqueur('nrCloture',False) ;
   Action:=caFree ;
   END ;
end;

procedure TFSocSISCO.CacheFonctions35 ;
BEGIN
if ((EstSerie(S3)) or (EstSerie(S5))) then
   BEGIN
   // Plan ref 2
   SO_CORSGE2.Visible:=False ; SO_CORSAU2.Visible:=False ;
   SO_CORSA12.Visible:=False ; SO_CORSA22.Visible:=False ;
   SO_CORSA32.Visible:=False ; SO_CORSA42.Visible:=False ;
   SO_CORSA52.Visible:=False ; SO_CORSBU2.Visible:=False ;
   // Axes > 2
// JLD5Axes
   if EstSerie(S3) then BEGIN SO_CORSA31.Visible:=False ; SO_CORSA41.Visible:=False ; SO_CORSA51.Visible:=False ; END ;
   HPlan2.Visible:=False ; SO_CORSBU1.Top:=SO_CORSA31.Top ;
   SO_BOURRELIB.Visible:=False ; TSO_BOURRELIB.Visible:=False ;
   END ;
if EstSerie(S3) then
   BEGIN
   // Axes > 1
   SO_CORSA21.Visible:=False ;
   SO_CORSBU1.Top:=SO_CORSA21.Top ;
   END ;
END ;

procedure TFSocSISCO.SO_DATEDEBUTEUROExit(Sender: TObject);
begin
SO_TAUXEURO.Enabled:=(V_PGI.DateEntree<TSocieteSO_DATEDEBUTEURO.AsDateTime) ;
TSO_TAUXEURO.Enabled:=SO_TAUXEURO.Enabled ;
end;

procedure TFSocSISCO.BRecupSISCOClick(Sender: TObject);
Var Fichier : TextFile ;
    St : String ;
    Ok00,Ok01,Ok02,Ok03,Ok50 : Boolean ;
begin
If Not EstFichierSISCO(StFichier,TRUE) Then Exit ;
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
InitMove(1000,'') ;
Ok00:=FALSE ; Ok01:=FALSE ; Ok02:=FALSE ; Ok03:=FALSE ; Ok50:=FALSE ;
if TSociete.State<>dsEdit then TSociete.Edit ;
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If (Copy(St,1,2)='00') Then
    BEGIN
    Ok00:=TRUE ;
    TSocieteSO_LGCPTEGEN.AsInteger:=StrToInt(Copy(St,21,2)) ;
    TSocieteSO_LGCPTEAUX.AsInteger:=StrToInt(Copy(St,21,2)) ;
    END ;
  If (Copy(St,1,2)='01') Then
    BEGIN
    Ok01:=TRUE ;
    TSocieteSO_LIBELLE.ASString:=Copy(St,3,30) ;
    END ;
  If (Copy(St,1,2)='02') Then
    BEGIN
    Ok02:=TRUE ;
    TSocieteSO_ADRESSE1.ASString:=Trim(Copy(St,3,32)) ;
    TSocieteSO_ADRESSE2.ASString:=Trim(Copy(St,35,32)) ;
    TSocieteSO_ADRESSE3.ASString:=Trim(Copy(St,67,32)) ;
    END ;
  If (Copy(St,1,2)='03') Then
    BEGIN
    Ok03:=TRUE ;
    TSocieteSO_SIRET.ASString:=Trim(Copy(St,3,14)) ;
    TSocieteSO_APE.ASString:=Trim(Copy(St,17,4)) ;
    TSocieteSO_TELEPHONE.ASString:=Trim(Copy(St,21,12)) ;
    If Copy(St,54,1)='N' Then Ok50:=TRUE ;
    END ;
  If (Copy(St,1,2)='50') Then
    BEGIN
    Ok50:=TRUE ;
//    SLgAna.Text:=Copy(St,44,2) ;
    END ;
  If Ok00 And Ok01 And Ok02 And Ok03 And Ok50 Then Break ;
  END ;
System.Close(Fichier) ;
FiniMove ;
end;

end.

unit Uscript;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(*                        I N T E R F A C E                           *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

interface

uses Classes, Windows, SysUtils, DB,
  uDbxDataSet, Variants, ADODB, HQRY,
  Forms, StdCtrls,
  Dialogs, FileCtrl, inifiles, menus,
  HEnt1, HMsgBox, Hctrls,
{$IFNDEF EAGLCLIENT}
  UTOB,
{$ENDIF}
{$IFDEF CISXPGI}
  uYFILESTD, ULibWindows,
{$ENDIF}
  cbpPath
 ,LicUtil;


Var
  CurrentDossier,CurrentDonnee : String ;
  
const
  ftTypes: array[0..2] of TFieldType = (ftString, ftFloat, ftString);

type
  TTypeInfo = (tiNull, tiScript, tiCaractere, tiShort, tiChaine, tiChamp,
    tiOption, tiCalcul, tiConcat, tiConstante, tiReference,
    tiTableCorr, tiInteger, tiBool);
  TTypeSortie = (tsSVACCESS, tsSVASCII);
  TTraitCpt = set of (tcRAZCPTGEN, tcRAZCPTAUX, tcTRTCPTAUX, tcTRTCTRAuto,
    tcRAZECRNUL, tcTRTCREcptAuto, tcTRTLumEcr);
  TViderTable = (vtJAMAIS, vtTOUJOURS, vtDEMANDER);
  TTransformMode = (tmAUCUN, tmMAJUSCULE, tmMINUSCULE);
  TPhaseMultiSeq = (pmDEBUT, pmMILIEU, pmFIN);
  TMultiSeqMod = (msLIGNEALIGNE, msCRITERE);

  TTypeCar = (tcOEM, tcANSI);
  TOperationChampNull = (ocnDefaut, ocnRemplacer, ocnLaisser, ocnEffacer);
  TOpTblCorr = (otcToujours, otcSurConf, otcJamais, otcDefaut);

{$IFNDEF CISXPGI}
Type LaVariablesCISX = RECORD
     Ini, Domaine, Mode,
     Nature, Complement,
     Directory, Script,
     NomFichier, Option,
     Famille, Appelant,
     RepSTD, LISTECOMMANDE,
     Monotraitement, compress : string;
     ATob                     : TOB;
     ListeFichier             : string;
    end;
Var VHCX : ^LaVariablesCISX ;
FUNCTION  CHARGECIX (FichierIni : string=''): Boolean ;
procedure LibereCX;
function GetInfoVHCX : LaVariablesCISX;

{$ELSE}
{$IFNDEF EAGLCLIENT}
TZVarCisx = class
  private
    Ini,
    RepSTD, LISTECOMMANDE    : string;
  public
    Domaine                  : string;
    NomFichier               : string;
    Directory                : string;
    Appelant                 : string;
    Famille                  : string;
    Mode                     : string;
    Script                   : string;
    Nature                   : string;
    Complement               : string;
    Option                   : string;
    ATob                     : TOB;
    Monotraitement, compress : string;
    ListeFichier             : string;
    constructor Create ; virtual ;
    procedure CHARGECIX(FichierIni : string='');
    destructor  Destroy ; override ;

end;
function GetInfoVHCX : TZVarCisx;
{$ENDIF}
{$ENDIF}

function ParLigneCommande : boolean;


Var CodeSeria : String ;

const
  TTypeChamp = [tiNull, tiCalcul, tiConcat, tiConstante, tiReference,
    tiTableCorr];

  CarSepField : array[0..4] of char = (';', ',', #9, '|', ':');
  AccessBase = 'PGZIMPACCESS.mdb';
  AccesBaseDonnee = 'PGZIMPACCESSD.mdb';

type
  TSqlMemo = class(TObject)
    NomMemo: string;
    NomCle: TStringList;
    constructor Create;
    destructor Destroy; override;
  end;

  TModeOuverture = (moNormal, moUser, moTest);
  TParOrig = (poGlobal, poDossier, poUtil, poClient);

  TFourchette = class(TObject)
  public
    deb, fin: string;
    bFourchette: boolean;

    destructor Destroy; override;
    function evalue(str: string): boolean;
  end;

  TMyWriter = class;
  TChampInfo = set of (ciUn, ciDeux, ciStocke, ciRetenu);

  PField1 = ^TField1;
  TField1 = record
    A: Pointer;
    B: Pointer;
    Name: array[0..64] of char;
    len: Smallint;
  end;

  TField1Array = array[0..500] of PField1;

  TOutputList = record
    DW1: Longint;
    DW2: Longint;
    DW3: Longint;
    data: ^TField1Array;
    Lim: Integer;
    used: Integer;
  end;

  TNextRecordEvent = function(Sender: TObject; AEtape: Integer; AB: PChar;
    bSelected: Boolean; var OutputList: TOutputList): integer of object;
      stdcall;

  TNextEvent = function(ANumEnr: Longint; AEtape: Integer): integer of object;
    stdcall;

  TInfoEvent = function(ANumEnr: Longint; AEtape: Integer)
    : integer of object; stdcall;

  TConstante = class(TPersistent)
  private
    FValeur: string;
    FPas: SmallInt;
    FLongueur: SmallInt;
  public
    procedure Assign(AValue: TConstante); reintroduce; overload ; virtual ;

  published
    property Valeur: string read FValeur write FValeur;
    property Pas: Smallint read FPas write FPas default 0;
    property Longueur: Smallint read FLongueur write FLongueur default 0;
  end;

  TReference = class(TPersistent)
  protected
    FCondition: string;
    FPos: Smallint;
    FLon: Smallint;
    FPost: Bool;
  public
    procedure Assign(AValue: TReference); reintroduce; overload ; virtual ;
  published
    property Condition: string read FCondition write FCondition;
    property Pos: Smallint read FPos write FPos;
    property Lon: Smallint read FLon write FLon;
    property Post: Bool read FPost write FPost default False;
  end;

  TCalcul = class(TPersistent)
  private
    FFormule: string;
  public
    procedure Assign(AValue: TCalcul); reintroduce; overload ; virtual ;
  published
    property Formule: string read FFormule write FFormule;
  end;

  PTableCorrRec = ^TTableCorrRec;
  TTableCorrRec = record
    FValeur: integer; (* numero de champ, negatif si externe *)
    FCount: Integer;
    FAssociee: Boolean; (* associe a un champ *)
    FEntree: TStringList; { table interne }
    FSortie: TStringList; { table interne }
    FEntreeExt: TStringList; { table externe }
    FSortieExt: TStringList; { table externe }
  end;

  TWriteEvent = procedure of object;
  TMyObjet = class(TPersistent)
    FProperties: TList;
  end;

  (* ------------------------------------------------------------------ *)
  (*                            T S c r ip t C o r r e s p              *)
  (* ------------------------------------------------------------------ *)

  TScriptCorresp = class(TComponent)
  protected
    FName  : string;
    Fichier : string;
    FEntree: TStringList; { table interne }
    FSortie: TStringList; { table interne }
  public
    procedure SaveTo(AStreamTable: TStream);
    destructor Destroy; override;
    constructor Create(AOwner: TComponent); override;
  published
    property Name: string read FName write FName;
    property FFichier: string read Fichier write Fichier;
    property LFEntree: TStringList read FEntree write FEntree;
    property LFSortie: TStringList read FSortie write FSortie;
  end;


  (* ------------------------------------------------------------------ *)
  (*                            T C H A M P                             *)
  (* ------------------------------------------------------------------ *)

  TChampList = class;
  TChampClass = class of TChamp;
  TChamp = class(TCollectionItem)
  protected
    FName  : string[64];
    FDeb   : Smallint;
    FLon   : Smallint; (* longueur paramétrée                    *)
    FTyp   : Smallint; (* 0:Alpha 1:Numeric 2:Date 3:Heure       *)
    FSiz   : Smallint; (*  longueur du champ dans TRA            *)
    FSel   : Smallint; (* 0:Tout 1:positif 2:negatif 4:NonRetenu *)
    (* 8:Non stocke  16:non supprimable       *)
    FTypeInfo : TTypeInfo;

    FCompl    : Boolean; (*                                        *)
    FComplLgn : Smallint;
    FComplCar : Char;
    FAlignLeft: Boolean; (*                                        *)

    FConstante: TConstante;
    FReference: TReference;
    FTableCorr: PTableCorrRec;

    FCalcul   : TCalcul;
    FConcat   : TChampList; (* liste des elements a concatener *)

    FTableExist  : Boolean;
    FTableExterne: Boolean;
    FNomTableExt : string;

    FCache       : Boolean;
    FTransform   : TTransformMode;
    FOpChampNull : TOperationChampNull;
    FOpTblCorr   : TOpTblCorr;
    FFiche       : Smallint;
    FbInterne    : Boolean; // permet de dire si le champ est temporaire
    FCommentTableCorr: string;
    FFormatDate      : smallint;
    FFormatDateSortie: smallint;
    FRang            : smallint; //ordre dans la liste des références
    FConditionChamp  : string;
    FbCondition      : Boolean;
    FbUnenreg        : Boolean;
    FListProfile     : TStringlist; // ajout me liste dprofile des correspondances
    FNbDecimal       : smallint;    // ajout me pour le nombre decimal
    FbArrondi        : Boolean;     //Ajout me pour arrondir un champ
    FNomFichExt      : string;
    FLienInter       : string;
    FLOrder          : string;  // ordre de tri
    FLFamilleCorr    : string;   //Famille de correspondance
    FLngRef          : Boolean;
    procedure SetName(AValue: string);
    function GetName : string;
    function GetCache: Boolean;
    procedure SetCache(AValue: Boolean);

    function IsStoredConstante: Boolean;
    function IsStoredReference: Boolean;
    function IsStoredTableCorr: Boolean;
    function IsStoredCalcul   : Boolean;
    function IsStoredConcat   : Boolean;
    function IsStoredComplLgn : Boolean;
    function IsStoredComplCar : Boolean;
  protected
    procedure SetTypeInfo(AValue: TTypeInfo);
    procedure SetTableExist(AValue: Boolean);
  public
    constructor Create(ACollection: TCollection); override;
    destructor Destroy; override;

    procedure Assign(AChamp: TChamp); reintroduce; overload ; virtual ;

    procedure IniTableCorr(LisCorresp: TStringList=nil; LisCorrespsortie : TStringList=nil);

    property TableExist: Boolean read FTableExist write FTableExist;
    property Transform: TTransformMode read FTransform write FTransform;
    property TableCorr: PTableCorrRec read FTableCorr write FTableCorr;
  published (* csLoading *)
    property Name: string read GetName write SetName;
    property Deb: Smallint read FDeb write FDeb default 0;
    property Lon: Smallint read FLon write FLon default 0;
    property Typ: Smallint read FTyp write FTyp default 0;
    property Siz: Smallint read FSiz write FSiz default 0;
    property Sel: Smallint read FSel write FSel default 0;
    property TypeInfo: TTypeInfo read FTypeInfo write SetTypeInfo;
    property Compl: Boolean read FCompl write FCompl;
    property FormatDate: SmallInt read FFormatDate write FFormatDate default 1;
    property FormatDateSortie: SmallInt read FFormatDateSortie write FFormatDateSortie default 1;
    property Rang: SmallInt read FRang write FRang default 0;
    property ComplLgn: Smallint read FComplLgn write FComplLgn stored
      IsStoredComplLgn;
    property ComplCar: Char read FComplCar write FComplCar stored
      IsStoredComplCar;
    property AlignLeft: Boolean read FAlignLeft write FAlignLeft;
    property TranformMode: TTransformMode read FTransform write FTransform
      default tmAUCUN;
    property TableExit: Boolean read FTableExist write SetTableExist default
      false;
    property NomTableExt: string read FNomTableExt write FNomTableExt;

    property NomFichExt: string read FNomFichExt write FNomFichExt;
    property TableExterne: Boolean read FTableExterne write FTableExterne default
      false;

    property CommentTableCorr: string read FCommentTableCorr write
      FCommentTableCorr;

    property Constante: TConstante read FConstante write FConstante stored
      IsStoredConstante;
    property Reference: TReference read FReference write FReference stored
      IsStoredReference;
    property Calcul: TCalcul read FCalcul write FCalcul stored IsStoredCalcul;
    property Concat: TChampList read FConcat write FConcat stored
      IsStoredConcat;
    property OpChampNull: TOperationChampNull read FOpChampNull write
      FOpChampNull default ocnDefaut;
    property OpTblCorr: TOpTblCorr read FOpTblCorr write FOpTblCorr default
      otcDefaut;
    //        property OpLibelleNull : TOperationLibelleNull read FOpLibelleNull write FOpLibelleNull Default ocnDefaut;
    property fiche: Smallint read FFiche write FFiche;
    property ConditionChamp: string read FConditionChamp write FConditionChamp;
    property bCondition: boolean read FbCondition write FbCondition;
    property bUnenreg: boolean read FbUnenreg write FbUnenreg;
    property Cache: Boolean read GetCache write SetCache;
    property ListProfile: TStringList read FListProfile write
      FListProfile;
    property NbDecimal: smallint read FNbDecimal write FNbDecimal default 2;
    property FArrondi: Boolean read FbArrondi Write FbArrondi default  FALSE;    // ajout me
    property LienInter: string read FLienInter write FLienInter;
    property LOrder: string read FLOrder write FLOrder;
    property FFamilleCorr: string read FLFamilleCorr write FLFamilleCorr;
    property LngRef: Boolean read FLngRef write FLngRef;
    property bInterne: boolean read FbInterne write FbInterne;   // pour les champs temporaires
  end;

  (* ------------------------------------------------------------------ *)
  (*                            T O P T I O N S                         *)
  (* ------------------------------------------------------------------ *)

  TOptions = class(TPersistent)
  private
    FDecimal: Char;
    FMillier: Char;
    FSepField: Char;
    FFileName: string;
    FCondition: string;
    FASCIIFileName: string;
    FTabSize: Word;
    FNbIgnoreLignes: Longint;
    FTypeCar: TTypeCar;
    FViderTable: TViderTable;
    FMultiSequence: Boolean;
    FPhaseMultiSeq: TPhaseMultiSeq;
    FMultiSeqMod: TMultiSeqMod;
    FMultiSeqDeb: Integer;
    FMultiSeqLon: Integer;
    FOpChampNull: TOperationChampNull;
    FOpTblCorr: TOpTblCorr;
    FModifEnrExistant: Boolean;
    FAliasODBC: string;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Assign(AOptions: TOptions); reintroduce; overload ; virtual ;
    property AliasODBC: string read FAliasODBC write FAliasODBC;
  published
    property Decimal: Char read FDecimal write FDecimal default '.';
    property Millier: Char read FMillier write FMillier default #0;
    property SepField: Char read FSepField write FSepField default ',';
    property Condition: string read FCondition write FCondition;
    property FileName: string read FFileName write FFileName;
    property ASCIIFileName: string read FASCIIFileName write FASCIIFileName;
    property TabSize: Word read FTabSize write FTabSize default 0;
    property ViderTable: TViderTable read FViderTable write FViderTable;
    property TypeCar: TTypeCar read FTypeCar write FTypeCar;
    property NbIgnoreLignes: Longint read FNbIgnoreLignes write FNbIgnoreLignes
      default 0;
    property MultiSequence: Boolean read FMultiSequence write FMultiSequence;
    property PhaseMultiSeq: TPhaseMultiSeq read FPhaseMultiSeq write
      FPhaseMultiSeq stored FMultiSequence;
    property MultiSeqMod: TMultiSeqMod read FMultiSeqMod write FMultiSeqMod
      stored FMultiSequence;
    property MultiSeqDeb: Integer read FMultiSeqDeb write FMultiSeqDeb stored
      FMultiSequence;
    property MultiSeqLon: Integer read FMultiSeqLon write FMultiSeqLon stored
      FMultiSequence;
    property OpChampNull: TOperationChampNull read FOpChampNull write
      FOpChampNull default ocnRemplacer;
    property OpTblCorr: TOpTblCorr read FopTblCorr write FopTblCorr default
      otcToujours;
    property ModifEnrExistant: Boolean read FModifEnrExistant write
      FModifEnrExistant default false;
  end; (* val *)

  (* ------------------------------------------------------------------ *)
  (*                    T O P T I O N C O M P T A                       *)
  (* ------------------------------------------------------------------ *)

  TInfoCpt = class(TCollectionItem)
  private
    FAux: string;
    FColl: string;
    FVari: string;
    FRaci: string;
    FSel: Smallint;
    //        FLonRacine : smallint;
 // protected
  public
    ListFourchette: TList;
    function EvalueRacine(compte: string): integer;
    destructor destroy; override;
    constructor create(Collection: Tcollection); override;
    procedure Assign(AValue: TPersistent); reintroduce; overload ; virtual ;
  published
    property Aux: string read FAux write FAux;
    property Coll: string read FColl write FColl;
    property Raci: string read FRaci write FRaci;
    property Vari: string read FVari write FVari;
    property Sel: Smallint read FSel write FSel;
    //        property LonRacine : smallint read  FLonRacine write FLonRacine;
  end;

  TCorrespList = class(TCollection)
  private
    function Get(itemIndex: Integer): TInfoCpt;
    procedure Put(index: integer; AValue: TInfoCpt);
  protected

  public
    property Items[Index: Integer]: TInfoCpt read Get write Put; default;
  end;

  TAxeAna = class(TCollectionItem)
  public
    FNomAxe: string;
    Fcondition: string;
    Fsel: boolean;
    FNumAxe: integer;
    procedure assign(Value: TPersistent); override;
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
  published
    property NomAxe: string read FNomAxe write FNomAxe;
    property NumAxe: integer read FNumAxe write FNumAxe;
    property condition: string read Fcondition write Fcondition;
    property sel: boolean read FSel write FSel default False;
  end;

  TListAxe = class(TCollection)
  public
    constructor Create;
    procedure Assign(AList: TListAxe);  reintroduce; overload ; virtual ;

    function Add: TAxeAna;
    procedure Delete(Index: Integer);
    function IndexOf(AValue: string): integer;
    procedure Move(curIndex, NewIndex: Integer);
    procedure OrganizeByNumAxe;
  end;

  TOptionCompta = class(TPersistent)
  private
    FOptions: TTraitCpt; (* 0:Aucune, 1:TraitAuxil 2:reset auxil *)
    FNature: Smallint; (* 0: Rattachement 1:Substitution 2:Regroupement *)
    FRattSubsLon: Smallint;
    FModeEcriture: Smallint;
      (* 0: Ecriture soldee 1:Numero Ecriture 2:ligne a ligne *)
    FLonRacine: Smallint; (* Taille de la racine 1..4 *)

    FcompteCTR: string;
    FLibelleCTR: string;

    FCodeJalANouveau: string; (* code du journal des a-nouveau *)
    FCritLigVent: string; (* critere de selection des lignes de ventilations *)
    FCritligAmortTech: string;

    FCorrespondance: TCorrespList; (* liste des correspondances *)
    FListAxe: TListAxe; //array [1..6] of TaxeAna;
    FGestionLumiere: SmallInt;
      // -1 : InfoLum avant; 0 : Pas d'infoLum; 1: InfoLum Après
    function IsStoredCorrespondance: Boolean;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Assign(AOptionCompta: TOptionCompta); reintroduce; overload ; virtual ;
    procedure DecoupeCritere(S: string; var Liste: Tlist);
  published
    property Options: TTraitCpt read FOptions write FOptions;
    property Nature: Smallint read FNature write FNature;
    property RattSubsLon: Smallint read FRattSubsLon write FRattSubsLon;
    property ModeEcriture: Smallint read FModeEcriture write FModeEcriture;
    property LonRacine: Smallint read FLonRacine write FLonRacine;
    property CodeJalANouveau: string read FCodeJalANouveau write
      FCodeJalANouveau;
    property CritLigVent: string read FCritLigVent write FCritLigVent;
    property CritLigAmortTech: string read FCritLigAmortTech write
      FCritLigAmortTech;
    property Correspondance: TCorrespList read FCorrespondance write
      FCorrespondance
      stored IsStoredCorrespondance;
    property Compte: string read FCompteCTR write FCompteCTR;
    property Libelle: string read FlibelleCTR write FLibelleCTR;
    property ListAxe: TListAxe read FListAxe write FListAxe;
    property GestionLumiere: SmallInt read FGestionLumiere write
      FGestionLumiere;
  end;

  (* ------------------------------------------------------------------ *)
  (*             T P R E T R T   e t   T P R E T R T L I S T            *)
  (* ------------------------------------------------------------------ *)

  TPreTrt = class(TCollectionItem)
  private
    FNom: string;
    FParam: TStringList;
  public
    constructor Create(ACollection: TCollection); override;
    destructor Destroy; override;

    procedure Assign(AValue: TPersistent); reintroduce; overload ; virtual ;
  published
    property Nom: string read FNom write FNom;
    property Param: TStringList read FParam write FParam;
  end;

  TPreTrtList = class(TCollection)
  private
    function Get(itemIndex: Integer): TPreTrt;
    procedure Put(index: integer; AValue: TPreTrt);
  protected

  public
    destructor Destroy; override;

    procedure Delete(Index: Integer);
    procedure Move(curIndex, NewIndex: Integer);
    procedure Assign(AValue: TPreTrtList);reintroduce; overload ; virtual ;

    property Items[Index: Integer]: TPreTrt read Get write Put; default;
  end;

  (* ------------------------------------------------------------------ *)
  (*                       T C H A M P L I S T                          *)
  (* ------------------------------------------------------------------ *)
  TChampList = class(TCollection)
  private
    function Get(itemIndex: Integer): TChamp;
    procedure Put(index: integer; AValue: TChamp);
  protected

  public
    destructor Destroy; override;

    procedure Assign(AList: TChampList); reintroduce; overload ; virtual ;

    procedure AddObject(AName: string; AChamp: TChamp);
    procedure Delete(Index: Integer);
    function IndexOf(AValue: string; deb: integer=0; ParFamille: Boolean=FALSE): integer;
    procedure Move(curIndex, NewIndex: Integer);

    property Items[Index: Integer]: TChamp read Get write Put; default;
  end;

  (* ------------------------------------------------------------------ *)
  (*                        T V A R I A B L E D E F                     *)
  (*                    T V A R I A B L E D E F L I S T                 *)
  (* ------------------------------------------------------------------ *)

  TVariableDefClass = class of TVariableDef;
  TVariableDef = class(TCollectionItem)
  private
    FTypeVar: Word;
    FName: string;
    FLibelle: string;
    FText: string;
    FDemandable: Boolean;
    FItems: TStringList;

  protected
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;

    procedure Assign(AValue: TPersistent); override;

  published
    property TypeVar: Word read FTypeVar write FTypeVar;
    property Demandable: Boolean read FDemandable write FDemandable default
      False;
    property Name: string read FName write FName;
    property Libelle: string read FLibelle write FLibelle;
    property Text: string read FText write FText;
    property Items: TStringList read FItems;
  end;

  TVariableDefList = class(TCollection)
  private
    function Get(itemIndex: Integer): TVariableDef;
  protected

  public
    constructor Create;

    procedure Assign(AList: TVariableDefList); reintroduce; overload ; virtual ;

    function AddObject(AName: string): TVariableDef;
    procedure Delete(Index: Integer);
    function IndexOf(AValue: string): integer;
    procedure Move(curIndex, NewIndex: Integer);

    property Items[Index: Integer]: TVariableDef read Get; default;
  end;

  (* ------------------------------------------------------------------ *)
  (*                                                                    *)
  (* ------------------------------------------------------------------ *)

  TFieldRec = class
  private
    FBuffer: PChar;
    FField: TField; // reference sur le champ associe

    function GetValue: string;
    procedure SetValue(AValue: string); virtual;
  public
    FName: string; // nom du champ
    FTableName: string; // nom de la table
    FFieldName: string; // nom du champ de table
    FFieldIndexName: string; // nom du champ d'index pour memo
    FFieldIndexValue: string; // nom de la valeur d'index pour memo
    FFieldMemoName: string; // nom de la zone dans le memo
    FOffset: Smallint; // offset dans le buffer d'entree
    FLon: Smallint; // longueur dans le buffer d'entree
    FTyp: Smallint; // type de champ
    FChamp: TChamp; // nom de la table
    FFiche: integer; // Fiche de Memo pour les info Trs.
    FbInterne: Boolean;
    Fsel: Smallint;
    FNbdecimal : Smallint; // ajout me
    FbArrondi  : Boolean; // ajout me
    function AsCurrency: Currency;

    procedure SetBuffer(ABuffer: PChar);

    property Field: TField write FField;
    property Value: string read GetValue write SetValue;
    property bInterne: boolean read FbInterne write FbInterne;
    property Fiche: integer read FFiche write FFiche;
    property Sel: Smallint read FSel write Fsel;
  end;

  TFieldList = class(TList)
    function Get(Index: Integer): TFieldRec;
  public
    function IndexOf(AValue: string): Integer;
    property Values[Index: Integer]: TFieldRec read Get; default;
  end;

  PTableRec = ^TTableRec;
  TTableRec = record
    FTable: TADOTABLE; // reference sur la table associee
    FFields: TFieldList; // liste des champs associes
    FFieldUsed: Integer;
    FAnnul: boolean;
  end;

  (* ------------------------------------------------------------------ *)
  (*                    T C O H E R E N C E                             *)
  (* ------------------------------------------------------------------ *)
  TSVParser = class;
  TCondCoherence = class(TCollectionItem)
  public //Item contenant le nom et la condition de la coherence
    FNom: string;
    FsCondition: string;
    FsMessage: string;
    FbStop: boolean;
    constructor create(NomCoherence: string; condition: string; collection:
      TCollection); reintroduce; overload ; virtual ;
    procedure Assign(Value: TPersistent); reintroduce; overload ; virtual ;
    function Evalue(SV: TSVParser; sText: PChar): boolean;
  published
    property Nom: string read FNom write FNom;
    property sCondition: string read FsCondition write FsCondition;
    property sMessage: string read FsMessage write FsMessage;
    property bStop: boolean read FbStop write FbStop;
  end;

  TCoherence = class(TCollection)
  public
    constructor create;
    destructor Destroy; override;

    procedure Assign(AList: TCoherence); reintroduce; overload ; virtual ;

    function Add: TCondCoherence;
    procedure Delete(Index: Integer);
    function IndexOf(AValue: string): integer;
    //		    property Items[Index: Integer] : TCondCoherence read Get write Put; default;
  end;


  (* ------------------------------------------------------------------ *)
  (*                          T S C R I P T                             *)
  (* ------------------------------------------------------------------ *)
// Version TComponent de 7 diffère avec la version 5 on passe par un record en paramètre pour la
// impfic7.dll
  TScriptRecord = Record
    FSignature: Longint;
    FOptions  : TOptions;
    FOptionCpt: TOptionCompta;
    FChamps   : TChampList;
    FTableCorr: TList;
    FName     : string[12];
    FFileType : Smallint; (* 0:base 1:ASCII *)
    FASCIIMODE: Smallint; (* 0:Fixe 1:Delimite *)
    FDestTable: Smallint; (* 0:Dossier 1:Utilisateur *)

    FScriptSuivant: string;

    FParName      : string; (* CLE PARAMETRE  *)
    FVariable     : TStringList;
    FVariableList : TVariableDefList;
    FParOrig      : TParOrig;
  end;

  TScript = class(TComponent)
  private
    FSignature: Longint;
    FOptions: TOptions;
    FOptionCpt: TOptionCompta;
    FChamps: TChampList;
    FTableCorr: TList;
    FName: string[12];
    FFileType: Smallint; (* 0:base 1:ASCII *)
    FASCIIMODE: Smallint; (* 0:Fixe 1:Delimite *)
    FDestTable: Smallint; (* 0:Dossier 1:Utilisateur *)

    FScriptSuivant: string;

    FParName: string; (* CLE PARAMETRE  *)
    FVariable: TStringList;
    FVariableList: TVariableDefList;
    FParOrig: TParOrig;

    FFields: TFieldList;
    FTables: TStringList;
    FPreTrt: TPreTrtList;

    FExecuteModeTest: Boolean; { mode normal par defaut }

    FOnNextRecord: TNextRecordEvent;
    FOnOuvrirTables: TNotifyEvent;
    FOnFermerTables: TNotifyEvent;
    FVersion: integer;
    FFileDest: Smallint; (* 0:Paradox 1:ASCII *)
    FCoherence: TCoherence;
    FModeparam:  Smallint; (* 0 mode complet 1 mode simplifié*)
    FShellexec   : string;
    function Get(Index: Integer): TChamp;
    function GetName: string;
    procedure SetName(AValue: string); reintroduce; overload ; virtual ;

    procedure CompareTableCorrRec(Sender: TObject; e1, e2: Word; var Action:
      Integer);


    function IsStoredPreTrt: Boolean;

  protected
    function GetAxe: boolean;
    function GetParName: string;
    function GetFileName: string;
    procedure SetFileName(AValue: string);

  public
    FOrigine: TParOrig;
    FComment: string;
    FEcrAna: Boolean;
    FEcrEch: Boolean;
    FTrsFiche: Boolean;
    FNouveau: Boolean;
    bTablecorr, bTableCorrExt, bAbort: boolean;

    NiveauAcces: Integer;
    DejaExecuter: Boolean;
    Destination: TParOrig;
    DeuxiemeExec: Boolean;
    Listtbl: TStringList;

    DllHandle: THandle;
    Axe: integer; // il n'y a pas de property pour Axe, on ne le sauvegarde pas
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure AddChamp(AChamp: TChamp);
    procedure InitTableAndFields; virtual;
    function IndexOfChamp(AValue: string): integer;
    function ChampByName(S: string): TChamp;
    function TableByName(S: string): PTableRec;
    function CloneScript: TScript;

    procedure InvalideVentilation;
    procedure SelectVentilation;

    procedure InvalideEch;
    procedure SelectEch;

    function EcrVent: boolean;
    function EcrEch: boolean;

    function ChampLocalEcr(Champ: TChamp): boolean;
    procedure SaveTo(AStream: TStream; AStreamTable: TStream);
    function Executer(AOnNextRecord: TNextRecordEvent): Boolean; virtual;
    procedure AssignSr (var SR :TScriptRecord);

    function Scruter(FChamp : TChamp; FScript : TScript; AonNextRecord: TNextEvent): string;

    function NewTableCorr: PTableCorrRec;
    procedure DelTableCorr(val: integer);
    function LaCallBack(FileName: PChar; bSelected: Boolean;
      var OutputList: TOutputList): integer; stdcall;

    procedure Assign(AScript: TScript; Racine :string='';  Profile : string=''; ListeFamille : HTStringList=nil); reintroduce; overload ; virtual ;
    procedure PrepareTableChampsASCII;
    procedure ConstruitVariable(bReplace: Boolean);
    procedure InterpreteSyntax;
    procedure OuvrirTables(ATable: TADOTable; N: Integer; EcrOrBal: Integer);
    procedure CloneTable(from, dest: TTable);
    function TraiteCoherence: boolean;
    function MajTableCorr(aChamp: TChamp; AonNextRecord: TNextEvent):
      TStringList; stdcall;
    function Callback(ANumEnr: Longint; AEtape: Integer): integer; stdcall;
    procedure CreerTables(var ATable: TTableRec);
    procedure BdeToOdbc(DataSet: TDataset; dsDest: TDataset);
    procedure SwapTableCorrRec(Sender: TObject; e1, e2: Word);
    procedure CompareTableList(Sender: TObject; e1, e2: Word; var Action:
      Integer);

    { liste des proprietes }

    property Champs[Index: Integer]: TChamp read Get;
    property Fields: TFieldList read FFields;
    property Tables: TStringList read FTables;
    property ExecuteModeTest: Boolean read FExecuteModeTest write
      FExecuteModeTest;
    property FileName: string read GetFileName write SetFileName;
    property ParOrig: TParOrig read FParOrig write FParOrig;

    property OnNextRecord: TNextRecordEvent write FOnNextRecord;
    property OnOuvrirTables: TNotifyEvent read FOnOuvrirTables write
      FOnOuvrirTables;
    property OnFermerTables: TNotifyEvent read FOnFermerTables write
      FOnFermerTables;
  published
    property Name: string read GetName write SetName;
    property ScriptSuivant: string read FScriptSuivant write FScriptSuivant;
    property FileType: Smallint read FFileType write FFileType;
    property ParName: string read FParName write FParName;
    property ASCIIMode: Smallint read FASCIIMode write FASCIIMode;
    property Version: integer read FVersion write FVersion;
    property Options: TOptions read FOptions write FOptions;
    property OptionsComptable: TOptionCompta read FOptionCpt write FOptionCpt;
    property PreTrt: TPreTrtList read FPreTrt write FPreTrt stored
      IsStoredPreTrt;
    property Variable: TStringList read FVariable write FVariable;
    property Champ: TChampList read FChamps write FChamps;
    property DestTable: Smallint read FDestTable write FDestTable;
      (* 0:Dossier 1:Utilisateur *)
    property Variables: TVariableDefList read FVariableList write FVariableList;
      (* 0:Dossier 1:Utilisateur *)
    property FileDest: Smallint read FFileDest write FFileDest default 1;
    property Coherence: TCoherence read FCoherence write FCoherence;
    property Modeparam: Smallint read FModeparam write FModeparam;
    property Shellexec  : String read FShellexec write FShellexec;

  end;

  TScriptclass = class of TScript;


  TMyWriter = class(TWriter)
  public // Objet TMyWriter
    procedure WriteType(ti: TTypeInfo);
  end;

  TCallBack = function(FileName: PChar): Integer; stdcall;

  TSVImport = class(TObject)
  public
    class function Create: TSVImport;
    procedure Free; virtual; stdcall; abstract;
    procedure SetInfo(Script: TScript); virtual; stdcall; abstract;

    function ImportFichier(uUserData: Longint; FileName: PChar; pInfo: Pointer;
      FCallBack: TCallBack): integer; virtual; stdcall; abstract;
    function TraiteCriter(crOff: integer; crLen: integer): PChar; virtual;
      stdcall; abstract;

    property Info: TScript write SetInfo;
  end;

  TScruterList = class(TObject)
  public
    procedure Free; virtual; stdcall; abstract;
    function Scrute(FileName: PChar; uUserData: Longint; pInfo: Pointer;
      AChamp: TChamp; FCallBack: TInfoEvent; dummy:integer=0): Integer; virtual; stdcall;
        abstract;
  end;

  TSVParser = class
  public
    procedure Free; virtual; stdcall; abstract;
    procedure SetMode(mode: integer); virtual; stdcall; abstract;
    function IsSelect(Ligne: PChar): Integer; virtual; stdcall; abstract;
    function IsTextSelect(Ligne: PChar): Integer; virtual; stdcall; abstract;
    procedure Analyse; virtual; stdcall; abstract;
    procedure SetText(p: PCHAR); virtual; stdcall; abstract;
    procedure SetScript(pInfo: Pointer); virtual; stdcall; abstract;
    procedure SetRefScript(s: TScript); virtual; stdcall; abstract;
    function GetResult: PChar; virtual; stdcall; abstract;
    procedure ResetFile; virtual; stdcall; abstract;
    procedure SetLine(ALine: integer); virtual; stdcall; abstract;
      // pour que la dll retrouve la ligne qu'il faut analyser

    property Script : Pointer write SetScript;
    property Text: PChar write SetText;
    property Mode: Integer write SetMode;
    property iLine: Integer write SetLine;
  end;

  TTVerifTblCorr = class(TObject)
  private // objet gérant la multiplicité des tables de correspondances
    ListeTbl: TStringList;
  public
    constructor create;
    destructor destroy; override;
  end;
  type
  PCreateScruteList = function: TScruterList; stdcall;

function LoadScriptFromStream(AStream: TStream; AStreamTable: TStream): TScript;
function LitScript(AStream: TStream; AStreamTable: TStream): TScript;
function InterpreteVar(SInput: string; StrList: TStringList): string;
function InterpreteVarSyntax(SInput: string; StrList: TStringList): string;
function InterpreteVarSyntaxSt(SInput: string): string;
function Interprete(SInput: string): string;
procedure AjouteVariableAlias(SL: TStringList);
function RecomposeChemin(SInput: string): string;
function retire(c: string; str: string): string;
procedure ImportAddField(Q: TDataset; Name: string; fType: TFieldType; flen:
  word; required: boolean; calc: boolean; Fieldvisible: boolean);
procedure ExecuteShell (Executable : string);
function LoadScriptCorresp(AStream: TStream): TScriptCorresp;
function CreerLigPop ( Name : string ; Owner : TComponent) : TMenuItem ;

var
  Impfic2: boolean = false;
  bTRACE: boolean = false;
  ImportSql: boolean = false;
  iTRACE: integer = 0;
  ImpficDll: string;
  Alias_Table_Ascii: string = '';
  GestionAxe: Boolean;
  ImportModeReseau: Boolean;

  CreerSVImport: function: TSVImport; // fonction de creation du TSVImport
  CreateScruteList: function: TScruterlist;
  CreerSVParser: function: TSVParser;

function IsEmpty(S: string): Boolean;

FUNCTION ADroite (st : string; l : Integer ; C : Char) : string ;
FUNCTION AGuauche (st : string; l : Integer ; C : Char) : string ;
function Estalpha (Chaine : string) : Boolean ;
Function ConvertDate(St : String) : TDateTime ;
Procedure FaitSISCOPer(Var FF : TextFile ; Datefin, Datedeb : TDateTime) ;
function ChargeFile (Command : string; indice:integer) : Boolean;
Procedure DeleteFichierRepertoire (Dossier : string; Attributs: integer);
function FusionFichierRepertoire (Dossier: string; Attributs: integer) : string;
{$IFNDEF EAGLCLIENT}
function InitScriptAuto(table : string) : integer;
{$IFDEF CISXPGI}
function RendTobparametre  (Domaine : string; var TOBParam : TOB; crit2 : string='COMPTA'; crit3: string='PARAM'; Crit4 : string=''): integer;
{$ENDIF}
{$ENDIF}
function ReconnaissanceScript (Requete : string) : string;
function AccessType(fd : TFieldType; Size : integer):string;


(* ------------------------------------------------------------------ *)
(*                                                                    *)
(*                   I M P L E M E N T A T I O N                      *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

implementation

uses BDE, Registry, Qsort,
{$IFDEF CISXPGI}
 ULibCpContexte,
{$ENDIF}
UDMIMP;

const
  cScriptVersion = 103;
  Signature = $C0DEABAD;

var
  CompteurReference: integer = 0;
  CompteurTableCorr: integer = 0;

  CurScript: TScript = nil;

class function TSVImport.Create: TSVImport;
begin
           @CreerSVParser := GetProcAddress(DMImport.HdlImpfic, 'CreerSVParser');
           if not Assigned(CreerSVParser) then
                raise Exception.Create('la fonction "CreerSVParser" n''a pas été trouvée.');

           @CreateScruteList := GetProcAddress(DMImport.HdlImpfic, 'CreerScruterList');
           if not Assigned(@CreateScruteList) then
                raise Exception.Create('fonction "CreerScruterList" non trouvée.');

  Result := CreerSVImport;
end;

constructor TSqlMemo.Create;
begin
  inherited Create;
  NomCle := TStringList.Create;
end;

destructor TSqlMemo.Destroy;
begin
  NomCle.Free;
  inherited Destroy;
end;

// ajoute les fields au nouveau Query
var
  Cpt: Integer = 1;

procedure ImportAddField(Q: TDataset; Name: string; fType: TFieldType; flen:
  word; required: boolean; calc: boolean; Fieldvisible: boolean);
var
  F: Tfield;
begin
  inc(Cpt);
  case fType of
    ftString: F := TStringField.Create(Q);
    ftSmallint: F := TSmallIntField.Create(Q);
    ftMemo: F := TMemoField.Create(Q);
    ftcurrency: F := TCurrencyField.Create(Q);
    ftBlob: F := TBlobField.Create(Q);
    ftBoolean: F := TBooleanField.Create(Q);
    ftInteger: F := TIntegerField.Create(Q);
    ftTime: F := TDateTimeField.Create(Q);
    ftFloat: F := TFloatField.Create(Q);
    ftDate: F := TDateField.Create(Q);
    ftGraphic: F := TGraphicField.Create(Q);
  else
    F := TField.Create(Q);
  end;
  F.SetFieldType(fType);
  try
    F.Name := Trim(Name);
  except
    F.Name := 'C' + IntToStr(Cpt);
  end;
  F.FieldName := Name;
  if fType <> ftsmallint then
  try
    F.Size := fLen - 1;
  except F.Size := 0;
  end;
  F.Calculated := Calc;
  F.Visible := Fieldvisible;
  F.DataSet := TDataSet(Q);
end;

(* ------------------------------------------------------------------ *)
(* InterpreteVar                                                      *)
(* -------------                                                      *)
(*   remplace les variables de la forme $(NOM) par leur valeur        *)
(*   trouve dans la liste de chaine StrList                           *)
(*                                                                    *)
(*   exemple :  $(NOM).$(EXT)                                         *)
(*      si StrList contient                                           *)
(*           NOM=fichier                                              *)
(*           EXT=asc                                                  *)
(*      la variable resultat contiendra                               *)
(*           fichier.asc                                              *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

function InterpreteVar(SInput: string; StrList: TStringList): string;
var
  P: Pchar;
  S2, Mot: string;
  FMode: Integer;
begin
  P := PChar(SInput);
  S2 := '';
  Fmode := 0;
  while P^ <> #0 do
  begin
    if Fmode = 1 then
    begin
      if P^ = ')' then
      begin
        S2 := S2 + StrList.Values[Mot];
        Fmode := 0;
      end
      else
        Mot := Mot + P^;
    end
    else if (P[0] = '$') and (P[1] = '(') then
    begin
      Fmode := 1;
      Mot := '';
      Inc(P);
    end
    else
      S2 := S2 + P^;
    Inc(P);
  end;
  Result := S2;
end;

function InterpreteVarSyntax(SInput: string; StrList: TStringList): string;
var
  P: Pchar;
  S2, Mot: string;
  FMode: Integer;
begin
  P := PChar(SInput);
  S2 := '';
  Fmode := 0;
  while P^ <> #0 do
  begin
    if Fmode = 1 then
    begin
      if P^ = ')' then
      begin
        S2 := S2 + '''' + StrList.Values[Mot] + '''';
        Fmode := 0;
      end
      else
        Mot := Mot + P^;
    end
    else if (P[0] = '$') and (P[1] = '(') then
    begin
      Fmode := 1;
      Mot := '';
      Inc(P);
    end
    else
      S2 := S2 + P^;
    Inc(P);
  end;
  Result := S2;
end;

function InterpreteVarSyntaxSt(SInput: string): string;
var
  P      : Pchar;
  S2     : string;
begin
  P := PChar(SInput);
  S2 := '';
  while P^ <> #0 do
  begin
    if (P[0] = #13) or (P[0] = #10)then begin Inc(P); continue;  end;
      S2 := S2 + P^;
    Inc(P);
  end;
  Result := S2;
end;


function Interprete(sInput: string): string; // renvoie le bon chemin tronqué
begin // vérification du chemin rentré par l'utilisateur
  result := sInput;
  if sInput = '' then
    exit;
end;

procedure AjouteVariableAlias(SL: TStringList);
begin // fonction intégrant les diff. variables obligatoires
end;

function RecomposeChemin(sInput: string): string;
  // renvoie le bon chemin tronqué
var
  Pass,Fich         : string;
   cheminInt        : string;
begin // vérification du chemin rentré par l'utilisateur
  result := sInput;
  if pos('($STD)\', sInput) <> 0 then
  begin
    {$IFDEF CISXPGI}
        Pass := CryptageSt(DayPass(Date));
    {$ELSE}
        Pass := DayPass(Date);
    {$ENDIF}

       Fich := Copy (sInput,pos('($STD)\', sInput)+7, length(sInput));

       if V_PGI.PassWord=Pass then
       begin
           CheminInt := ExtractFileDir (Application.ExeName)+'\parametre\'+Fich ;
           Result := CheminInt;
       end
       else
       begin
         if not FileExists(CurrentDossier+'\'+Fich) then
          Copyfile (PChar(ExtractFileDir (Application.ExeName)+'\parametre\'+Fich), Pchar(CurrentDossier+'\'+Fich),TRUE);
         CheminInt := CurrentDossier+'\'+Fich ;
         Result := CheminInt;
       end;
  end;
  if sInput = '' then
    exit;
end;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)
(*
function TTVerifTblCorr.VerifTBLCorr(script: Tscript; AOnNextRecord:
  TNextEvent): Boolean;
begin
end;

procedure TTVerifTblCorr.ShowCorr(curchamp: Tchamp; FScript: TScript; CorrFound:
  HTStringList);
begin
end;
*)
constructor TTVerifTblCorr.create;
begin
  ListeTbl := TStringList.create;
end;

destructor TTVerifTblCorr.destroy;
begin
  ListeTbl.Free;
end;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

procedure TMyWriter.WriteType(ti: TTypeInfo);
begin
  Write(ti, sizeof(ti));
end;

{ }

function TFieldList.Get(Index: Integer): TFieldRec;
begin
  Result := Items[Index];
end;

function TFieldList.IndexOf(AValue: string): Integer;
var
  N: Integer;
begin
  Result := -1;
  for N := 0 to Count - 1 do
    if AnsiCompareText(TFieldRec(Items[N]).FName, AValue) = 0 then
    begin
      Result := N;
      exit;
    end;
end;

procedure TFieldRec.SetBuffer(ABuffer: PChar);
begin
  FBuffer := ABuffer;
end;

function TFieldRec.GetValue: string;
begin
  //  if ((Sel and 4) <> 4) then begin result := '';	exit; end;
  try
    SetString(Result, FBuffer + FOffset, FLon);
    Result := TrimRight(Result);
  except
    Result := '';
  end;
end;

function TFieldRec.AsCurrency: Currency;
var
  S: string;
begin
  SetString(S, FBuffer + FOffset, FLon);
  try
    Result := StrToFloat(S);
  except Result := 0;
  end;
end;

procedure TFieldRec.SetValue(AValue: string);
begin
  try
    if FField <> nil then
       FField.AsString := AValue;
  except
    if FField is TNumericField then
      FField.Value := 0;
  end;
  //	except FField.AsString := '0'; end;
end;

{ --- }

procedure TConstante.Assign(AValue: TConstante);
begin
  FValeur := AValue.FValeur;
  FPas := AValue.FPas;
  FLongueur := AValue.FLongueur;
end;

procedure TReference.Assign(AValue: TReference);
begin
  FCondition := AValue.FCondition;
  FPos := AValue.FPos;
  FLon := AValue.FLon;
  FPost := AValue.FPost;
end;

procedure TCalcul.Assign(AValue: TCalcul);
begin
  FFormule := AValue.FFormule;
end;

(* ----------------------------------------------------------------------- *)
(*                                                                         *)
(* ----------------------------------------------------------------------- *)

constructor TOptions.Create;
begin
  FDecimal := '.';
  FMillier := #0; // pas de separateur de millier
  FSepField := ','; // separateur de champ
  FASCIIFileName := '';
  FOpChampNull := ocnRemplacer;

end;

destructor TOptions.Destroy;
begin
end;

procedure TOptions.Assign(AOptions: TOptions);
begin
  FDecimal := AOptions.FDecimal;
  FMillier := AOptions.FMillier;
  FSepField := AOptions.FSepField;
  FFileName := AOptions.FFileName;
  FCondition := AOptions.FCondition;
  FASCIIFileName := AOptions.FASCIIFileName;
  FTabSize := AOptions.FTabSize;
  FNbIgnoreLignes := AOptions.FNbIgnoreLignes;
  FTypeCar := AOptions.FTypeCar;
  FViderTable := AOptions.FViderTable;
  FMultiSequence := AOptions.FMultiSequence;
  FPhaseMultiSeq := AOptions.FPhaseMultiSeq;
  FMultiSeqMod := AOptions.FMultiSeqMod;
  FMultiSeqDeb := AOptions.FMultiSeqDeb;
  FMultiSeqLon := AOptions.FMultiSeqLon;
  FopChampNull := AOptions.FOpChampNull;
  FOpTblCorr := AOptions.FOpTblCorr;
  ModifEnrExistant := AOptions.ModifEnrExistant;
  AliasODBC := AOptions.AliasODBC;
end;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* ------------------------------------------------------------------ *)

procedure TInfoCpt.Assign(AValue: TPersistent);
var
  i: integer;
begin
  if AValue is TInfoCpt then
  begin
    FAux := TInfoCpt(AValue).Aux;
    FColl := TInfoCpt(AValue).Coll;
    FRaci := TInfoCpt(AValue).Raci;
    FVari := TInfoCpt(AValue).Vari;
    FSel := TInfoCpt(AValue).Sel;
    //        FLonRacine := TInfoCpt(AValue).LonRacine;
    if Tinfocpt(avalue).listfourchette <> nil then
    begin
      ListFourchette.Clear;
      for I := 0 to Tinfocpt(avalue).listfourchette.count - 1 do
        Listfourchette.Add(Tinfocpt(avalue).listfourchette.items[I]);
    end;
  end
  else
    inherited Assign(AValue);
end;

(* ----------------------------------------------------------------------- *)
(*                                                                         *)
(* ----------------------------------------------------------------------- *)

function TCorrespList.Get(itemIndex: integer): TInfoCpt;
begin
  Result := TInfoCpt(inherited Items[itemIndex]);
end;

procedure TCorrespList.Put(index: integer; AValue: TInfoCpt);
begin
  Items[index].Assign(AValue);
end;

(* ----------------------------------------------------------------------- *)
(*                                                                         *)
(* ----------------------------------------------------------------------- *)

constructor TOptionCompta.Create;
begin
  FCorrespondance := TCorrespList.Create(TInfoCpt);
  FOptions := [tcTRTCREcptAuto, tcTRTLumEcr];
  FlistAxe := TListAxe.Create;
end;

destructor TOptionCompta.Destroy;
begin
  FCorrespondance.Free;
  FListAxe.Free;
  inherited Destroy;
end;

function TOptionCompta.IsStoredCorrespondance: Boolean;
begin
  Result := tcTRTCPTAUX in FOptions;
end;

procedure TOptionCompta.DecoupeCritere(S: string; var Liste: TList);
var
  P, pch, pch2: PChar;
  S2: string;
  F: TFourchette;
  deb, fin: string;
begin // découpe et analyse la chaine syntaxiquement et sémantiquement
  if Liste = nil then
    Liste := TList.create;
  Liste.clear;
  setstring(S2, Pchar(S), length(S)); // ne pas optimiser
  P := PChar(S2);

  repeat
    pch := StrScan(P, ',');
    if pch <> nil then
    begin
      pch^ := #0;
      inc(pch);
    end;
    (* éclatement de l'élément *)
    pch2 := StrScan(P, '-');
    if pch2 <> nil then
    begin
      pch2^ := #0;
      inc(pch2);
      deb := Trim(P);
      fin := Trim(pch2);
    end
    else
    begin
      deb := Trim(P);
      fin := deb;
    end;

    if (deb <> '') or (fin <> '') then
    begin
      F := TFourchette.Create;
      F.deb := deb;
      F.fin := fin;
      F.bFourchette := not (deb = fin);
      Liste.Add(F);
    end;
    P := pch;
  until pch = nil;
end;

procedure TOptionCompta.Assign(AOptionCompta: TOptionCompta);
var
  ic: TInfoCpt;
  N, M: Integer;
begin
  FOptions := AOptionCompta.FOptions;
  FModeEcriture := AOptionCompta.FModeEcriture;
  FNature := AOptionCompta.FNature;
  FRattSubsLon := AOptionCompta.FRattSubsLon;
  FCodeJalANouveau := AOptionCompta.FCodeJalANouveau;
  FCritLigVent := AOptionCompta.FCritLigVent;
  FCritligAmortTech := AOptionCompta.FCritligAmortTech;
  FcompteCTR := AOptionCompta.FcompteCTR;
  FLibelleCTR := AOptionCompta.FLibelleCTR;
  FLonRacine := AOptioncompta.FLonracine;
  FGestionLumiere := AOptioncompta.GestionLumiere;
  Flistaxe.assign(Aoptioncompta.ListAxe);
  FCorrespondance.Clear;

  for N := 0 to AOptionCompta.FCorrespondance.Count - 1 do
  begin
    ic := FCorrespondance.Add as TInfoCpt;
    ic.Assign(AOptionCompta.FCorrespondance.Items[N]);
    { -- Ajoute les elements dans la liste par ordre croissant sur FAux }
    M := 0;
    while M < FCorrespondance.Count do
    begin
      if TInfoCpt(FCorrespondance.Items[M]).FAux > ic.FAux then
        break;
      Inc(M);
    end;
    if M <> FCorrespondance.Count then
      ic.Index := M;
  end;
end;

(* ------------------------------------------------------------------ *)
(* TChampList                                                         *)
(* ------------------------------------------------------------------ *)

destructor TChampList.Destroy;
begin
  inherited Destroy;
end;

procedure TChampList.Assign(AList: TChampList);
var
  N: Integer;
  AChamp: TChamp;
begin
  for N := 0 to AList.Count - 1 do
  begin
    AChamp := TChamp.Create(self);
    AChamp.Assign(AList[N]);
  end;
end;

function TChampList.Get(itemIndex: integer): TChamp;
begin
  Result := TChamp(inherited Items[itemIndex]);
end;

procedure TChampList.Put(index: integer; AValue: TChamp);
begin
  Items[index].Assign(AValue);
end;

function TChampList.IndexOf(AValue: string; deb: integer=0; ParFamille: Boolean=FALSE): integer;
var
  N, i : integer;
  NN   : string;
  RChamp: string;
begin
  Result := -1;
  for N := 0 to Count - 1 do
  begin
       if ParFamille then
          RChamp := UpperCase(TChamp(Items[N]).FFamilleCorr)
       else
           RChamp :=  TChamp(Items[N]).FName;
       if AnsiCompareText(RChamp, AValue) = 0 then
       begin
               Result := N;
               exit;
       end;
  end;
  if Result = -1 then
  begin
        for N := deb to Count - 1 do
        begin
          i := pos('_', TChamp(Items[N]).FName);
          if i <> 0 then NN := copy(TChamp(Items[N]).FName, 1, i-1);
          if AnsiCompareText(NN, AValue) = 0  then
          begin
               result := N;
               exit;
          end;
        end;
  end;
end;

procedure TChampList.Move(curIndex, NewIndex: Integer);
begin
  Items[curIndex].Index := NewIndex;
end;

procedure TChampList.Delete(Index: Integer);
begin { New Dispose}
  TChamp(Items[Index]).Free;
end;

procedure TChampList.AddObject(AName: string; AChamp: TChamp);
begin
  Add.Assign(AChamp);
end;

(* ------------------------------------------------------------------ *)
(* TChamp                                                             *)
(* ------------------------------------------------------------------ *)

function TChamp.IsStoredConstante: Boolean;
begin
  Result := TypeInfo = tiConstante;
end;

function TChamp.IsStoredReference: Boolean;
begin
  Result := TypeInfo = tiReference;
end;

function TChamp.IsStoredTableCorr: Boolean;
begin
  Result := TypeInfo = tiTableCorr;
end;

function TChamp.IsStoredCalcul: Boolean;
begin
  Result := TypeInfo = tiCalcul;
end;

function TChamp.IsStoredConcat: Boolean;
begin
  Result := TypeInfo = tiConcat;
end;

function TChamp.IsStoredComplLgn: Boolean;
begin
  Result := FCompl;
end;

function TChamp.IsStoredComplCar: Boolean;
begin
  Result := FCompl;
end;

function TChamp.GetCache: Boolean;
begin
  Result := FCache;
end;

procedure TChamp.SetCache(AValue: Boolean);
begin
  FCache := AValue;
end;

procedure TChamp.SetName(AValue: string);
begin
  FName := AValue;
end;

function TChamp.GetName: string;
begin
  Result := FName;
end;

procedure TChamp.IniTableCorr(LisCorresp: TStringList=nil; LisCorrespsortie : TStringList=nil);
var
  Ft      : TextFile;
  TempStr : string;
  N       : integer;
  function GetEntree(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := 1 to length(str) do
    begin
      if str[i] = ';' then
        break;
      Result := Result + str[i];
    end;
  end;

  function GetSortie(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := pos(';', str) to (length(str) - 1) do
      Result := Result + str[i + 1];
  end;
begin
  if TableExterne then
  begin
    if FTableCorr^.FEntreeExt = nil then
      FTableCorr^.FEntreeExt := TStringList.create;
    if FTableCorr^.FSortieExt = nil then
      FTableCorr^.FSortieExt := TStringList.Create;
    FTableCorr^.FentreeExt.Clear;
    FTableCorr^.FSortieExt.Clear;
      AssignFile(Ft, RecomposeChemin(InterpreteVar(NomTableExt, LisCorresp)));
      Reset(Ft);
      while not EOF(Ft) do
      begin
        Readln(Ft, TempStr);
        FTableCorr^.FEntreeExt.add(GetEntree(TempStr));
        FTableCorr^.FSortieExt.add(GetSortie(TempStr));
      end;
        closeFile(Ft);
  end
  else
  begin // Table interne
          if LisCorresp = nil then exit;
          if FTableCorr = nil then   SetTableExist(TRUE);
          FTableExist := TRUE;

          FTableCorr^.Fentree.Clear;
          FTableCorr^.FSortie.Clear;
          for  N :=0 to LisCorresp.count-1 do
          begin
               FTableCorr^.FEntree.add(LisCorresp.strings[N]);
               FTableCorr^.FSortie.add(LisCorrespsortie.strings[N]);
          end;
          FTableCorr^.FCount := FTableCorr^.FEntree.Count;
  end;
end;


constructor TChamp.Create(ACollection: TCollection);
begin
  FFormatDate := 1;
  FFormatDateSortie := 1;
  Nbdecimal   := 2;
  FbArrondi   := FALSE;
  inherited Create(ACollection);
      // ajout me
  FListProfile := TStringList.create;
end;

destructor TChamp.Destroy;
begin
  FCalcul.Free;
  FConcat.Free;
  FConstante.Free;
  FReference.Free;
 //ajout me
  FListProfile.free;

  if assigned(FTableCorr) then
  begin
    FtableCorr.FEntree.free;
    FtableCorr.FSortie.free;
    FtableCorr.FEntreeExt.free;
    FtableCorr.FSortieExt.free;
    FtableCorr.FEntree := nil;
    FtableCorr.FSortie := nil;
    FtableCorr.FEntreeExt := nil;
    FtableCorr.FSortieExt := nil;
    Dispose(FtableCorr);
  end;
  inherited Destroy;
end;

procedure TChamp.Assign(AChamp: TChamp);
var
  N: Integer;
  Ft: TextFile;
  TempStr: string;

  function GetEntree(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := 1 to length(str) do
    begin
      if str[i] = ';' then
        break;
      Result := Result + str[i];
    end;
  end;

  function GetSortie(str: string): string;
  var
    i: integer;
  begin
    Result := '';
    for i := pos(';', str) to (length(str) - 1) do
      Result := Result + str[i + 1];
  end;
begin
  FName := AChamp.FName;
  FDeb := AChamp.FDeb;
  FLon := AChamp.FLon;
  FRang := AChamp.FRang;
  FTyp := AChamp.FTyp; (* 0:Alpha 1:Numeric 2:Date 3:Heure       *)
  FSiz := AChamp.FSiz; (*                                        *)
  FSel := AChamp.FSel; (* 0:Tout 1:positif 2:negatif 4:NonRetenu *)
  (* 8:Non stocke  16:non supprimable       *)
  FTypeInfo := AChamp.FTypeInfo;

  FCompl       := AChamp.FCompl;
  FComplLgn    := AChamp.FComplLgn;
  FComplCar    := AChamp.FComplCar;
  FAlignLeft   := AChamp.FAlignLeft;
  FFormatDate  := AChamp.FFormatDate;
  FLngRef      := AChamp.FLngRef;
  FbInterne    := Achamp.bInterne;
  FFormatDateSortie := AChamp.FFormatDateSortie;
  case FTypeInfo of
    tiConstante:
      begin
        FConstante := TConstante.Create;
        FConstante.Assign(AChamp.FConstante);
      end;
    tiReference:
      begin
        FReference := TReference.Create;
        FReference.Assign(AChamp.FReference);
      end;
    tiCalcul:
      begin
        FCalcul := TCalcul.Create;
        FCalcul.assign(AChamp.FCalcul);
      end;
    tiConcat:
      begin
        if not Assigned(FConcat) then
          FConcat := TChampList.Create(TChamp);
        FConcat.Assign(AChamp.FConcat); (* liste des elements a concatener *)
      end;
  end;

  if Assigned(AChamp.FTableCorr) then
  begin (* TList *)
    New(FTableCorr);
    with FTableCorr^ do
    begin
      FEntree := TStringList.Create;
      FSortie := TStringList.Create;
      FEntreeExt := TStringList.Create;
      FSortieExt := TStringList.Create;
      FValeur := AChamp.FTableCorr^.FValeur;
      FCount := AChamp.FTableCorr^.FCount;
      FAssociee := AChamp.FTableCorr^.FAssociee;
      if AChamp.TableExterne then
      begin
           TableExterne :=  AChamp.TableExterne;
           NomTableExt  := RecomposeChemin(AChamp.NomTableExt);
           FTableExist := TRUE;
           AChamp.FTableCorr^.FEntreeExt.Clear;
           AChamp.FTableCorr^.FSortieExt.Clear;
           AChamp.FTableCorr^.FEntree.Clear;
           AChamp.FTableCorr^.FSortie.Clear;
           if not FileExists(NomTableExt) then
                    PGIInfo ('Le fichier de correspondance '+ NomTableExt+ ' sur le champ'+ FName + ' n''existe pas','')
           else
           begin
               AssignFile(Ft, NomTableExt);
               {$I-} Reset (Ft) ; {$I+}
               while not EOF(Ft) do
                begin
                  Readln(Ft, TempStr);
                  AChamp.FTableCorr^.FEntreeExt.add(GetEntree(TempStr));
                  AChamp.FTableCorr^.FSortieExt.add(GetSortie(TempStr));
                  AChamp.FTableCorr^.FEntree.add(GetEntree(TempStr));
                  AChamp.FTableCorr^.FSortie.add(GetSortie(TempStr));
                end;
                closeFile(Ft);
            end;
            TableCorr^.FEntree.Assign(AChamp.FTableCorr^.FEntree);
            TableCorr^.FSortie.Assign(AChamp.FTableCorr^.FSortie);
            TableCorr^.FCount := AChamp.FTableCorr^.FEntree.Count;
      end;

      if Assigned(AChamp.FTableCorr^.FEntree) then
        FEntree.Assign(AChamp.FTableCorr^.FEntree);
      if Assigned(AChamp.FTableCorr^.FSortie) then
        FSortie.Assign(AChamp.FTableCorr^.FSortie);
      if Assigned(AChamp.FTableCorr^.FEntreeExt) then
        FEntreeExt.Assign(AChamp.FTableCorr^.FEntreeExt);
      if Assigned(AChamp.FTableCorr^.FSortieExt) then
         FSortieExt.Assign(AChamp.FTableCorr^.FSortieExt);
    end;
  end;

  FTableExist          := AChamp.FTableExist; // TableExterne ? pour Table de Corr.
  FTableExterne        := AChamp.FTableExterne; // TableExterne ? pour Table de Corr.
  FNomTableExt         := AChamp.FNomTableExt; // nom du fichier pour la table externe
  FNomFichExt          := AChamp.FNomFichExt; // nom du fichier pour la table externe
  FLFamilleCorr         := AChamp.FLFamilleCorr;  // ajout me

  FCache               := AChamp.FCache;
  FTransform           := AChamp.FTransform;
  FOpChampNull         := AChamp.FOpChampNull; // opération sur champ nul
  FOpTblCorr           := AChamp.FOpTblCorr; // opération sur tables de corr.
  Ffiche               := AChamp.FFiche;
  FCommentTableCorr    := Achamp.CommentTableCorr;
  conditionChamp       := Achamp.conditionChamp;
  bCondition           := Achamp.bcondition;
  for N := 0 to Achamp.ListProfile.count-1 do
      FListProfile.add(Achamp.ListProfile.strings[N]);
  FNbDecimal            := Achamp.FNbDecimal;
  FbArrondi             := Achamp.FbArrondi;
  LienInter             := Achamp.FLienInter;
  LOrder                := Achamp.FLOrder;
end;

procedure TChamp.SetTableExist(AValue: Boolean);
begin
  if AValue and not Assigned(FTableCorr) then
    FTableCorr := CurScript.NewTableCorr;
  FTableExist := AValue;
end;

procedure TChamp.SetTypeInfo(AValue: TTypeInfo);
begin
  if (AValue = tiConcat) and not Assigned(FConcat) then
    FConcat := TChampList.Create(TChamp);
  if (AValue = tiConstante) and not Assigned(FConstante) then
    FConstante := TConstante.Create;
  if (AValue = tiCalcul) and not Assigned(FCalcul) then
    FCalcul := TCalcul.Create;
  if (AValue = tiReference) and not Assigned(FReference) then
    FReference := TReference.Create;
  FTypeInfo := AValue;
end;

(* ------------------------------------------------------------------ *)
(* TVariableDef                                                       *)
(* ------------------------------------------------------------------ *)

constructor TVariableDef.Create(Collection: TCollection);
begin
  FItems := TStringList.Create;
  inherited Create(Collection);
end;

destructor TVariableDef.Destroy;
begin
  FItems.Free;
  inherited Destroy;
end;

procedure TVariableDef.Assign(AValue: TPersistent);
begin
  FTypeVar := TVariableDef(AValue).FTypeVar;
  FName := TVariableDef(AValue).FName;
  FLibelle := TVariableDef(AValue).FLibelle;
  FText := TVariableDef(AValue).FText;
  FDemandable := TVariableDef(AValue).FDemandable;
  FItems.Assign(TVariableDef(AValue).FItems);
end;

constructor TVariableDefList.Create;
begin
  inherited Create(TVariableDef);
end;

function TVariableDefList.Get(itemIndex: Integer): TVariableDef;
begin
  Result := TVariableDef(inherited Items[itemIndex]);
end;

procedure TVariableDefList.Assign(AList: TVariableDefList);
var
  I: Integer;
begin
  Clear;
  for I := 0 to AList.Count - 1 do
    Add.Assign(AList[I]);
end;

function TVariableDefList.AddObject(AName: string): TVariableDef;
begin
  Result := TVariableDef(Add);
end;

procedure TVariableDefList.Delete(Index: Integer);
begin
  TVariableDef(Items[Index]).free;
end;

function TVariableDefList.IndexOf(AValue: string): integer;
var
  N: integer;
begin
  for N := 0 to Count - 1 do
    if AnsiCompareText(TVariableDef(Items[N]).Name, AValue) = 0 then
    begin
      Result := N;
      exit;
    end;
  Result := -1;
end;

procedure TVariableDefList.Move(curIndex, NewIndex: Integer);
begin
  ;
end;

(* ------------------------------------------------------------------ *)
(* TChampList                                                         *)
(* ------------------------------------------------------------------ *)

constructor TPreTrt.Create(ACollection: TCollection);
begin
  inherited Create(ACollection);
  FParam := TStringList.Create;
end;

destructor TPreTrt.Destroy;
begin
  FParam.Free;
  inherited Destroy;
end;

procedure TPreTrt.Assign(AValue: TPersistent);
begin
  Nom := TPreTrt(AValue).Nom;
  Param.Assign(TPreTrt(AValue).Param);
end;

destructor TPreTrtList.Destroy;
begin
  inherited Destroy;
end;

function TPreTrtList.Get(itemIndex: integer): TPreTrt;
begin
  Result := TPreTrt(inherited Items[itemIndex]);
end;

procedure TPreTrtList.Put(index: integer; AValue: TPreTrt);
begin
  Items[index].Assign(AValue);
end;

procedure TPreTrtList.Assign(AValue: TPreTrtList);
var
  N: integer;
  aPreTrt: TPreTrt;
begin
  Clear;
  for N := 0 to AValue.Count - 1 do
  begin
    aPreTrt := Add as TPreTrt;
    aPreTrt.Assign(AValue[N]);
  end;
end;

procedure TPreTrtList.Move(curIndex, NewIndex: Integer);
begin
  Items[curIndex].Index := NewIndex;
end;

procedure TPreTrtList.Delete(Index: Integer);
begin { New Dispose}
  TPreTrt(Items[Index]).Free;
end;

(* ------------------------------------------------------------------ *)
(*                      T S C R I P T                                 *)
(* ------------------------------------------------------------------ *)

constructor TScript.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  CurScript := Self;
  FSignature := Signature;
  FName := '';

  FVariableList := TVariableDefList.Create;
  FVariable := TStringList.Create;
  FVariable.Sorted := True;
  FChamps := TChampList.Create(TChamp);
  FTableCorr := TList.Create;
  FOptions := TOptions.Create;
  FOptionCpt := TOptionCompta.Create;
  FFields := TFieldList.Create;
  FTables := TStringList.Create;
  FPreTrt := TPreTrtList.Create(TPreTrt);
  Coherence := TCoherence.Create;

  NiveauAcces := 1;
  FileDest := 1; //Destination AsciiTemp pour les AsciiTemporaires
end;


destructor TScript.Destroy;
var
  N : Integer;
begin
  FPreTrt.Free;
  FOptionCpt.Free;
  for N := FTables.Count - 1 downto 0 do
  begin
//    PTableRec(FTables.Objects[N])^.FTable.free;
    PTableRec(FTables.Objects[N])^.FTable.close;
    PTableRec(FTables.Objects[N])^.FFields.Free;
    Dispose(PTableRec(FTables.Objects[N]));
  end;
  FTables.Free;
  for N := FFields.Count - 1 downto 0 do
    TFieldRec(FFields[N]).Free;
  FFields.Free;
  FOptions.Free;
  FTableCorr.Free;
  FChamps.Free;
  FVariable.Free;
  FVariableList.Free;
  Coherence.Free;
  if assigned(ListTbl) then
    for N := 0 to listTbl.count - 1 do
      if Assigned(TStringList(ListTbl.objects[N])) then
        TStringList(ListTbl.objects[N]).Free;
  ListTbl.free;

end;

(* ------------------------------------------------------------------ *)
(*                                                                    *)
(* PROCEDURE xxxxxxxxxxxx                                             *)
(* ------------------------------------------------------------------ *)

function TScript.IsStoredPreTrt: Boolean;
begin
  Result := FPreTrt.Count <> 0;
end;

const
  NBChampLocalEcr = 64;

  ChampLocalEcrArray: array[0..NBChampLocalEcr] of string =
  ('NUMECRIT', 'NUMECRIT1', 'NUMLIGNE', 'Compte', //4
    'DateComptable', 'Debit', 'Credit', 'Solde', //4
    'Libelle', 'Libelle_du_compte', 'Complement', 'CodeJournal', 'TypeJournal',
      //5 =13
    'Folio', 'PieceComptable', 'Codejustificatif', 'PieceJustificative',
      'CodeType', //5 =18
    'LigneType', 'DateLettrage', 'CodeLettrage', 'Datepointage', 'CodePointage',
      //5 =23
    'CodeDev', 'TauxTVA', 'CodeSociete', 'CodeEtablissement', 'QuantiteDebit',
      //5 =28
    'QuantiteCredit', 'QuantiteSolde', 'QUantiteCode', 'DevDebit', 'DevCredit',
      //5 =33
    'DevSolde', 'TauxConversion', 'DateTauxConversion', 'CodeDev2', //4 =37
    'DevDebit3', 'DevCredit3', 'DevSolde3', 'TauxConversion3',
      'DateTauxConversion3', 'CodeDev3', //6 =43
    'Critere.Critere1', 'Critere.Critere2', 'Critere.Critere3',
      'Critere.Critere4', 'Critere.Critere5', 'Critere.Critere6',
      'Critere.Critere7', 'Critere.Critere8',
    'Remise', 'Releve', 'InformationAlpha1.Numero', 'InformationAlpha1.Libelle',
      'InformationAlpha1.Montant', 'InformationAlpha1.Date', //4 =57
    'CodeEtabOrigine', 'CodeSocieteOrigine', 'CreationEcriture.Date',
      'CreationEcriture.Operateur', 'ModificationEcriture.Date',
      'ModificationEcriture.Operateur', //4 =63
    'ValidationEcriture.Date', 'ValidationEcriture.Operateur'); //2 champs =65

  ChampVent: array[0..13] of string =
  ('CompteAnal', 'VDateComptable',
    'Section1', 'Section2', 'Section3', 'Section4', 'Section5', 'Section6',
    'VPourcentage', 'VDebit', 'VCredit', 'VSolde',
    'VQSolde', 'VLibelle');
  ChampEch: array[0..7] of string =
  ('Date_Echeance', 'Mode_Reglement', 'Code_Echeance', 'Pourc_Echeance',
    'Montant_Echeance', 'Bon_A_Payer', 'MontantEcheance2', 'Visa');

  ChampFiche: array[0..7] of string = // pour les infos tiers
  ('Compte', 'Texte', 'Adresse', 'CodePostal', 'Ville', 'CodePays',
    'Pays', 'FormeJuridique');

function TScript.EcrVent: boolean;
var
  N, M: Integer;
begin
  Result := false;
  for N := 0 to FChamps.Count - 1 do
  begin
    if ParName <> 'TIERS' then // pour les ecritures
    begin
      for M := 0 to 13 do
        if (AnsiCompareText(FChamps[N].Name, ChampVent[M]) = 0) and
          ((FChamps[N].FSel and 4) <> 4) then
        begin
          Result := True;
          exit;
        end;
    end
    else // pour les info tiers
    begin
      if (StrPos(PChar(FFields[N].FFieldName), 'TEXTE') <> nil)
        and ((FChamps[N].FSel and 4) <> 4) then
      begin
        Result := True;
        exit;
      end;

      {       for M:=0 to 7 do

   if (not(AnsiCompareText(FChamps[N].Name, ChampFiche[M]) = 0)) and
     ((FChamps[N].FSel and 4) <> 4) then
               begin
    Result := True;
    exit;
   end;            }
    end;
  end;
end;

function TScript.EcrEch: boolean;
var
  N, M: Integer;
begin
  Result := false;
  for N := 0 to FChamps.Count - 1 do
  begin
    for M := 0 to 7 do
    begin
      if (AnsiCompareText(FChamps[N].Name, ChampEch[M]) = 0) and
        ((FChamps[N].FSel and 4) <> 4) then
      begin
        Result := True;
        exit;
      end;
    end;
  end;
end;

procedure TScript.InvalideVentilation;
var
  N, M: Integer;
begin
  for N := 0 to FChamps.Count - 1 do
  begin
    if ParName <> 'TIERS' then
    begin
      for M := 0 to 13 do
        if AnsiCompareText(FChamps[N].Name, ChampVent[M]) = 0 then
        begin
          FChamps[N].FSel := 4;
        end;
    end
    else // info Trs
    begin
      if (StrPos(PChar(FFields[N].FFieldName), 'TEXTE') <> nil) then
      begin
        FChamps[N].FSel := 4;
      end;
    end;
  end;
end;

procedure TScript.SelectVentilation;
var
  N, M: Integer;
  Sel: Boolean;
begin
  for N := 0 to FChamps.Count - 1 do
  begin
    Sel := False;
    if ParName <> 'TIERS' then
    begin // ecriture normal
      for M := 0 to 13 do
        if AnsiCompareText(FChamps[N].Name, ChampVent[M]) = 0 then
        begin
          Sel := True;
          break;
        end;
    end
    else // info TIERS
    begin
      if (StrPos(Pchar(FFields[N].FFieldName), 'TEXTE') <> nil) then
      begin
        Sel := True;
        break;
      end;
    end;
    if not Sel then
      if not ChampLocalEcr(Champ[N]) then
        Champ[N].FSel := 4;
  end;
end;

function TScript.ChampLocalEcr(Champ: TChamp): boolean;
var
  j: integer;
begin
  Result := true;
  for J := 0 to 7 do
  begin
    if AnsiCompareText(Champ.Name, ChampEch[J]) = 0 then
      Result := false;
    if not result then
      break;
  end;

  for J := 0 to 13 do
  begin
    if AnsiCompareText(Champ.Name, ChampVent[J]) = 0 then
      Result := false;
    if not result then
      break;
  end;

  for J := 0 to NBCHAMPLocalEcr do
  begin
    if AnsiCompareText(Champ.Name, ChampLocalEcrArray[J]) = 0 then
      Result := false;
    if not Result then
      break;
  end;
end;

procedure TScript.InvalideEch;
var
  N, M: Integer;
begin
  for N := 0 to FChamps.Count - 1 do
  begin
    for M := 0 to 7 do
      if AnsiCompareText(FChamps[N].Name, ChampEch[M]) = 0 then
      begin
        FChamps[N].FSel := 4;
      end;
  end;
end;

procedure TScript.SelectEch;
var
  N, M: Integer;
  Sel: Boolean;
begin
  for N := 0 to FChamps.Count - 1 do
  begin
    Sel := False;
    for M := 0 to 7 do
      if AnsiCompareText(FChamps[N].Name, ChampEch[M]) = 0 then
      begin
        Sel := True;
        break;
      end;

    if not Sel then
      FChamps[N].FSel := 4;
  end;
end;

function TScript.GetName: string;
begin
  Result := FName;
end;

procedure TScript.SetName(AValue: string);
begin
  FName := AValue;
end;

function TScript.IndexOfChamp(AValue: string): integer;
begin
  Result := FChamps.IndexOf(AValue);
end;

function TScript.ChampByName(S: string): TChamp;
var
  index: integer;
begin
  index := FChamps.IndexOf(S);
  if index >= 0 then
    Result := FChamps[index]
  else
    Result := nil;
end;

function TScript.TableByName(S: string): PTableRec;
var
  index: integer;
begin
  index := FTables.IndexOf(S);
  if index >= 0 then
    Result := PTableRec(FTables.Objects[index])
  else
    Result := nil;
end;

procedure TScript.AddChamp(AChamp: TChamp);
begin
  FChamps.AddObject(AChamp.Name, AChamp);
end;

function TScript.Get(index: integer): TChamp;
begin
  Result := TChamp(FChamps[index]);
end;

procedure TScript.CompareTableCorrRec(Sender: TObject; e1, e2: Word; var Action:
  Integer);
begin
  with Sender as TList do
    Action := PTableCorrRec(items[e1 - 1])^.FValeur - PTableCorrRec(items[e2 -
      1])^.FValeur;
end;

procedure TScript.CompareTableList(Sender: TObject; e1, e2: Word; var Action:
  Integer);
begin
  with FChamps do
  begin
    Action := AnsiCompareText(Items[e1 - 1].Name, Items[e1 - 1].Name);
  end;
end;

procedure TScript.SwapTableCorrRec(Sender: TObject; e1, e2: Word);
var
  p: pointer;
begin
  with Sender as TList do
  begin
    p := items[e1 - 1];
    items[e1 - 1] := items[e2 - 1];
    items[e2 - 1] := p;
  end;
end;

procedure TScript.SaveTo(AStream: TStream; AStreamTable: TStream);
var
  ST1, ST2: TStream;
  WR: TMyWriter;
  N: Integer;
  NbTableCorrAssociee: Integer;
  ATri: TQSort;

  procedure SauveTableCorr(ATable: TTableCorrRec);
  var
    M: Integer;
  begin
    WR.WriteInteger(ATable.FCount);
    with ATable do
      for M := 0 to FCount - 1 do
      begin
        WR.WriteStr(FEntree.Strings[M]);
        WR.WriteStr(FSortie.Strings[M]);
      end;
  end;

  procedure ParcoursChampsInterprete(slChamp: TChampList);
  var
    N: Integer;
    CH: TChamp;
  begin
    for N := 0 to slChamp.Count - 1 do
    begin
      CH := slChamp[N];
      if CH.TypeInfo = tiConcat then
        ParcoursChampsInterprete(CH.Concat);
      if CH.TableExist and CH.TableExterne then
        CH.NomTableExt := Interprete(CH.NomTableExt);
    end;
  end;

  procedure ParcoursChampsTable(slChamp: TChampList);
  var
    N: Integer;
    CH: TChamp;
  begin
    for N := 0 to slChamp.Count - 1 do
    begin
      CH := slChamp[N];
      if CH.TypeInfo = tiConcat then
        ParcoursChampsTable(CH.Concat);
      if CH.TableExist and not CH.TableExterne then
        SauveTableCorr(CH.TableCorr^);
    end;
  end;
begin (* SaveTo *)
  // Interpretation des noms de table de correspondance
  if AStreamTable <> nil then ParcoursChampsInterprete(Champ);
  // - - - - - - - - - - - - - - - - - - - - - - - - - -
  Version := cScriptVersion;
  ST1 := TMemoryStream.Create;
  WR := TMyWriter.Create(ST1, 4096);
  try
    WR.WriteRootComponent(Self);
  except
    WR.Free;
    ST1.Free;
    ShowMessage(Exception(ExceptObject).message);
    Exit;
  end;
  WR.Free;

  ST1.Seek(0, 0);
  ST2 := AStream;
  try
    ObjectBinaryToText(ST1, ST2);
  finally
    //      ShowMessage(Exception(ExceptObject).message);
    ST1.Free;
  end;

  (* Sauvegarde des tables de correspondances internes *)
  (* tri de la liste des tables *)
  if FTableCorr.Count > 1 then
  begin
    ATri := TQSort.Create(nil);
    ATri.Compare := CompareTableCorrRec;
    ATri.Swap := SwapTableCorrRec;
    ATri.DoQSort(FTableCorr, FTableCorr.Count);
    ATri.Free;
  end;

  WR := TMyWriter.Create(AStreamTable, 4096);
  NbTableCorrAssociee := 0;
  for N := 0 to FTableCorr.Count - 1 do
    if PTableCorrRec(FTableCorr[N])^.FAssociee then
      Inc(NbTableCorrAssociee);

  WR.WriteInteger(NbTableCorrAssociee);
  ParcoursChampsTable(champ);
  for N:=0 to FTableCorr.Count-1 do
   begin
    if not PTableCorrRec(FTableCorr[N])^.FAssociee then continue;
    SauveTableCorr(PTableCorrRec(FTableCorr[N])^);
   end;
  WR.Free;
end;

type
  TMonReader = class(TReader)
  protected
    function Error(const Message: string): Boolean; override;
  end;

var
  ErrorReader: TStringList; // repertorie toutes les erreurs du TMonReader

function TMonReader.Error(const Message: string): Boolean;
begin
  if not Assigned(ErrorReader) then
    ErrorReader := TStringList.Create;
  ErrorReader.Add(Message);
  Result := True;
end;

function LoadScriptFromStream(AStream: TStream; AStreamTable: TStream): TScript;
var
  N, CountTableCorr: integer;
  AScript: TScript;
  RD: TMonReader;
  ST1, ST2: TStream;

  procedure ChargeTableCorr(var ATable: TTableCorrRec);
  var
    M, CountCorr: Integer;
  begin
        CountCorr := RD.ReadInteger;
        with ATable do
          for M := 0 to CountCorr - 1 do
          begin
            FEntree.Add(RD.ReadStr);
            FSortie.Add(RD.ReadStr);
          end;
        ATable.FCount := CountCorr;
  end;
begin (* LoadScriptFromStream *)
  ST1 := AStream;
  ST2 := TMemoryStream.Create;
  try (* transform l'objet texte en binaire *)
    ObjectTextToBinary(ST1, ST2);
  except
    ST2.Free;
    raise;
  end;
  ST2.Seek(0, 0); (* creer l'objet *)
  RD := TMonReader.Create(ST2, 4096);
  try
    AScript := RD.ReadRootComponent(nil) as TScript;
  except RD.Free;
    ST2.Free;
    raise;
  end;
  RD.Free;
  ST2.Free;

  Result := AScript;

  { lecture de(s) table(s) de correspondances }
  if AStreamTable = nil then
    exit;
 try
  RD := TMonReader.Create(AStreamTable, 4096);
  CountTableCorr := RD.ReadInteger;
  with AScript do
      for N := 1 to CountTableCorr do
      begin
        ChargeTableCorr(PTableCorrRec(FTableCorr[N - 1])^);
        PTableCorrRec(FTableCorr[N - 1])^.FAssociee := True;
      end;
  RD.Free;
  except
  end;
  AScript.InitTableAndFields;
end;

function LitScript(AStream: TStream; AStreamTable: TStream): TScript;
var
  N, CountTableCorr: integer;
  AScript: TScript;
  RD: TMonReader;
  ST2: TStream;

  procedure ChargeTableCorr(var ATable: TTableCorrRec);
  var
    M, CountCorr: Integer;
  begin
    CountCorr := RD.ReadInteger;
    with ATable do
      for M := 0 to CountCorr - 1 do
      begin
        FEntree.Add(RD.ReadStr);
        FSortie.Add(RD.ReadStr);
      end;
    ATable.FCount := CountCorr;
  end;
begin (* LoadScriptFromStream *)
  //ST1 := AStream;
  ST2 := TMemoryStream.Create;
  //	try (* transform l'objet texte en binaire *)
  //		ObjectTextToBinary(ST1, ST2);
  //	except
  //		ST2.Free;
  //		raise;
  //	end;
  ST2.Seek(0, 0); (* creer l'objet *)
  RD := TMonReader.Create(ST2, 4096);
  try
    AScript := RD.ReadRootComponent(nil) as TScript;
  except RD.Free;
    ST2.Free;
    raise;
  end;
  RD.Free;
  ST2.Free;

  Result := AScript;

  { lecture de(s) table(s) de correspondances }
  RD := TMonReader.Create(AStreamTable, 4096);
  CountTableCorr := RD.ReadInteger;
  for N := 1 to CountTableCorr do
  begin
      ChargeTableCorr(PTableCorrRec(Ascript.FTableCorr[N - 1])^);
      PTableCorrRec(aScript.FTableCorr[N - 1])^.FAssociee := True;
  end;
  RD.Free;
  AScript.InitTableAndFields;
end;

function retire(c: string; str: string): string;
var
  n : integer;
begin
  Result := '';
  for N := 1 to length(str) do
  begin
  if str[N] <> c then
//    if (strcomp(PChar(str[N]), PChar(c)) <> 0) then
      Result := Result + str[N];
  end;
end;

procedure TScript.InitTableAndFields;
var
  N: Integer;
  AChamp: TChamp;
  AFieldRec: TFieldRec;
begin
  for N := 0 to Champ.Count - 1 do
  begin
    AChamp := Champ[N];
    AFieldRec := TFieldRec.Create;
    with AFieldRec do
    begin
      FTableName := Options.ASCIIFileName;
      FName := AChamp.FName;
      FFieldName := AChamp.FName;
      FTyp := AChamp.FTyp;
      FLon := AChamp.FSiz;
      FNbdecimal := AChamp.FNbdecimal; // ajout me
      FbArrondi  := Achamp.FbArrondi;
      FChamp := AChamp;
      if (FTyp = 0) and (FLon = 0) then
        FLon := AChamp.FLon;
    end;
    FFields.Add(AFieldRec);
  end;
end;

function TScript.NewTableCorr: PTableCorrRec;
begin // Ajoute une nouvelle tablecorr à la liste
  New(Result);
  FillChar(Result^, sizeof(Result^), 0);
  with Result^ do
  begin
    FValeur := CompteurTableCorr;
    FEntree := TStringList.Create;
    FSortie := TStringList.Create;
    FEntreeExt := TStringList.Create;
    FSortieExt := TStringList.Create;

    Inc(CompteurTableCorr);
    FTableCorr.Add(Result);
  end;
end;

procedure TScript.DelTableCorr(val: integer);
var
  I: integer;
begin
  for I := 0 to FTableCorr.count - 1 do
  begin
    if PTablecorrRec(FTableCorr[I]).FValeur = val then
    begin
      PTablecorrRec(FTableCorr[I]).FEntree.Free;
      PTablecorrRec(FTableCorr[I]).FEntree := nil;
      PTablecorrRec(FTableCorr[I]).FSortie.free;
      PTablecorrRec(FTableCorr[I]).FSortie := nil;
      PTablecorrRec(FTableCorr[I]).FEntreeExt.Free;
      PTablecorrRec(FTableCorr[I]).FEntreeExt := nil;
      PTablecorrRec(FTableCorr[I]).FSortieExt.Free;
      PTablecorrRec(FTableCorr[I]).FSortieExt := nil;
      FTableCorr.Delete(I);
      Dec(CompteurTableCorr);
    end;
  end;
end;

procedure TScript.Assign(AScript: TScript; Racine :string=''; Profile : string=''; ListeFamille: HTStringList=nil);
var
  N       : Integer;
  AChamp  : TChamp;
  procedure findvariable (Namechamp : string; index : integer);
  var
  id        : integer;
  St,SQL    : string;
  NTable    : string;
  Q1        : TQuery;
  begin
       For id:=0 To AScript.Variables.count-1 do
       begin
                    if (not AScript.Variables.Items[id].demandable) and ((pos (Racine, AScript.Variables.Items[id].Name) <> 0) or (Namechamp = AScript.Variables.Items[id].Name)) then
                    begin
                            if (Namechamp = AScript.Variables.Items[id].Name) then
                            begin
                              St :=  AScript.Champ[index].Name;
                              St := ReadTokenPipe (St, '_');
                              NTable := AScript.Name+St;
                              if (Variables.Items[id].FTypeVar = 1) and (Variables.Items[id].FLibelle = 'SOMME')then
                                   SQL := 'SELECT SUM(' + AScript.Variables.Items[id].Name+') from '+NTable
                              else
                              SQL := 'SELECT ' + AScript.Variables.Items[id].Name+' from '+NTable;
                              Q1 := OpenSQLADO (SQL, DMImport.DBGlobalD.ConnectionString);
                              if not Q1.EOF then
                              begin
                                   if (Variables.Items[id].FTypeVar = 1) and (Variables.Items[id].FLibelle = 'SOMME') then // numérique
                                      Variables.Items[id].Text :=  FloatToStr(Arrondi(Q1.fields[0].asfloat,AScript.Champ[Index].FNbDecimal))
                                   else
                                      Variables.Items[id].Text := Q1.findfield(AScript.Variables.Items[id].Name).asstring;
                                   AScript.FVariableList.Items[id].Text :=Variables.Items[id].Text;
                              end;
                              Ferme(Q1);
                            end;
                    end;
       end;
  end;
  procedure ChargeCorresp (Profile,famille : string; ch : Tchamp);
  var
  TCorr : TADOTable;
  AStreamTable          : TmemoryStream;
  ATblCorField          : TField;
  ScriptCorr            : TScriptCorresp;
  begin
                 try
                      ScriptCorr := nil;
                      TCorr := TADOTable.create(Application);
                      TCorr.Connection := DMImport.DBGLOBAL;
                      TCorr.TableName := DMImport.GzImpCorresp.TableName;
                      with TCorr do
                      begin
                           Open;
                           if Locate('Profile;Champcode',VarArrayOf([Profile,famille]), [loCaseInsensitive]) then
                           begin
                                        ATblCorField := FieldByName('TBLCOR');
                                        AStreamTable := TmemoryStream.create;
                                        TBlobField(ATblCorField).SaveToStream (AStreamTable);
                                        AStreamTable.Seek (0,0);
                                        ScriptCorr := LoadScriptCorresp(AStreamTable);
                                        AStreamTable.free;
                                        With ScriptCorr do ch.IniTableCorr(LFEntree, LFSortie);
                           end;
                           close;
                      end;
                      TCorr.free;
                      ScriptCorr.destroy;
                 except;
                 end;
  end;
begin
  FSignature := AScript.FSignature;
  FName := AScript.FName;
  FVersion := Ascript.FVersion;
  FFileType := AScript.FFileType; (* 0:Paradox 1:ASCII *)
  FASCIIMODE := AScript.FASCIIMODE; (* 0:Fixe 1:Delimite *)
  FDestTable := AScript.FDestTable; (* 0:Dossier 1:Utilisateur *)
  FOrigine := AScript.FOrigine;
  destination := AScript.destination;
  NiveauAcces := Ascript.NiveauAcces; (* 1:Modifiable  0:Non Modifiable *)
  FileDest := Ascript.filedest;
  FModeparam := AScript.FModeparam; (* 0:complet 1: simplifié*)

  FScriptSuivant := AScript.FScriptSuivant;

  FVariable.Assign(AScript.FVariable);
  FVariableList.Assign(AScript.FVariableList);
  FPreTrt.Assign(Ascript.preTrt);
  FParOrig := AScript.FParOrig;
  FexecuteModeTest := AScript.ExecuteModeTest;

  FParName := AScript.FParName+Racine;
  FOptions.Assign(AScript.FOptions);
  FOptionCpt.Assign(AScript.FOptionCpt);
  bTableCorr := Ascript.btablecorr;
  options.condition := Ascript.options.condition;
  FShellexec := Ascript.Shellexec;
  for N := 0 to AScript.FChamps.Count - 1 do
  begin
      if AScript.Variables.count <> 0 then
      findvariable (AScript.Champ[N].Name, N);

      // pour Cas Exercice et Exercice1 test avec '_'
       if (Copy (AScript.Champ[N].Name, 0, Length (Racine)+1) = Racine+'_')
       or (Racine = '') then
       begin
            AChamp := TChamp.Create(FChamps);
            if (AChamp.TableExterne) then
            begin
                 bTableCorr := AChamp.TableExterne;
            end;
            AChamp.Assign(AScript.FChamps[N]);
            AChamp.bUnenreg := AScript.FChamps[N].bUnenreg;
            if AScript.FChamps[N].conditionchamp <> '' then
            begin
               if (AScript.FChamps[N].bcondition) or (options.condition = '') then
                  options.condition := AScript.FChamps[N].conditionChamp
               else
                  options.condition := options.condition + ' ET '  + AScript.FChamps[N].conditionChamp;
            end;
            if (ListeFamille <> nil) and (AScript.FChamps[N].FLFamilleCorr<> '') and (ListeFamille.count <> 0) then
            begin
                 if ListeFamille.IndexOf(AScript.FChamps[N].FLFamilleCorr) <> -1 then
                 begin
                  ChargeCorresp(Profile,AScript.FChamps[N].FLFamilleCorr, AChamp);
                 end;
            end;
       end;
  end;
  Coherence.Assign(Ascript.coherence);
  InitTableAndFields;
end;

function TScript.GetParName: string;
begin
  Result := InterpreteVar(FParName, FVariable);
end;

function TScript.GetFileName: string;
begin
  Result := InterpreteVar(FOptions.FFileName, FVariable);
end;

procedure TScript.SetFileName(AValue: string);
begin
  FOptions.FFileName := AValue;
end;

(* ------------------------------------------------------------------ *)
(* ************************ EXECUTER ******************************** *)
(* ------------------------------------------------------------------ *)

// ----------------------- LaCallBack -------------------------------
//***********************************************************************

function TScript.LaCallBack(FileName: PChar; bSelected: Boolean;
  var OutputList: TOutputList): integer; stdcall;
begin
  Application.ProcessMessages;
  Result := 1; (* CONTINUE *)

  try
    if Assigned(FOnNextRecord) then
      Result := FOnNextRecord(self, 1, FileName, bSelected, OutputList);
  except on e: Exception do
    begin
      //	  	     ShowMessage(e.message);
      Result := -1; (* ABANDONNE *)
    end;
  end;
end;

procedure TScript.AssignSr (var SR :TScriptRecord);
begin
  // Initialisation de record pour la Dll
  SR.FSignature := FSignature;
  SR.FOptions   := FOptions;
  SR.FOptionCpt := FOptionCpt;
  SR.FChamps    := FChamps;
  SR.FTableCorr := FTableCorr;
  SR.FName      := FName;
  SR.FFileType  := FFileType; (* 0:base 1:ASCII *)
  SR.FASCIIMODE := FASCIIMODE; (* 0:Fixe 1:Delimite *)
  SR.FDestTable := FDestTable; (* 0:Dossier 1:Utilisateur *)
  SR.FScriptSuivant:= FScriptSuivant;
  SR.FParName   := FParName; (* CLE PARAMETRE  *)
  SR.FVariable  := FVariable;
  SR.FVariableList:= FVariableList;
  SR.FParOrig   := FParOrig;
end;

//------------------ Fonction Executer pour les BALANCES -------------------
//**************************************************************************

function TScript.Executer(AOnNextRecord: TNextRecordEvent): Boolean;
type
  TImportFichier = function(uUserData: Longint; FileName: PChar;
    pInfo: Pointer;
    FCallBack: TCallBack): Integer; stdcall;
var
  ImportFichier: TImportFichier;
  N, index     : Integer;
  aFieldRec: TFieldRec;
  aTableRec: PTableRec;
  AFileName, S: string;
  POL: ^TOutputList;
  PrevDateSeparator, PrevDecimalSeparator: Char;
  SVImport: TSVImport;
  SR :TScriptRecord;

  procedure ParcoursChamps(slChamp: TChampList);
  var
    N: Integer;
    CH: TChamp;
  begin
    for N := 0 to slChamp.Count - 1 do
    begin
      CH := slChamp[N];
      if CH.TypeInfo = tiConcat then
        ParcoursChamps(CH.Concat);
      if CH.TableExist and CH.TableExterne then
        CH.NomTableExt := RecomposeChemin(CH.NomTableExt);
    end;
  end;

begin
  N := 0; POL := nil; SVImport := nil; aFieldRec := nil; ImportFichier := nil;
  ParcoursChamps(Champ);
  OnNextRecord  := AOnNextRecord;

  // Initialisation de record pour la Dll
  SR.FSignature := FSignature;
  SR.FOptions   := FOptions;
  SR.FOptionCpt := FOptionCpt;
  SR.FChamps    := FChamps;
  SR.FTableCorr := FTableCorr;
  SR.FName      := FName;
  SR.FFileType  := FFileType; (* 0:base 1:ASCII *)
  SR.FASCIIMODE := FASCIIMODE; (* 0:Fixe 1:Delimite *)
  SR.FDestTable := FDestTable; (* 0:Dossier 1:Utilisateur *)
  SR.FScriptSuivant:= FScriptSuivant;
  SR.FParName   := FParName; (* CLE PARAMETRE  *)
  SR.FVariable  := FVariable;
  SR.FVariableList:= FVariableList;
  SR.FParOrig   := FParOrig;


  (* PREPARATION DE LA LISTE DES TABLES ET DES CHAMPS *)
  case TTypeSortie(FFileType) of
    tsSVACCESS: ;
    tsSVASCII : PrepareTableChampsASCII;
  end;
  if not impfic2 then
  begin
    @ImportFichier := GetProcAddress(DMImport.HdlImpfic, 'ImportFichier');
    if DMImport.HdlImpfic = 0 then
      raise
        Exception.Create('la fonction "ImportFichier" n''a pas été trouvée');
    DllHandle := DMImport.HdlImpfic;
  end
  else
  begin
    SVImport := TSvImport.create;
    SVImport.info := self;
  end;

  DllHandle := DMImport.HdlImpfic;

  PrevDateSeparator := DateSeparator;
  PrevDecimalSeparator := DecimalSeparator;

  try
    try
      if Assigned(FOnNextRecord) then
        FOnNextRecord(self, 0, nil, False, POL^);
      { ouverture des tables }
      try
        for N := 0 to FTables.Count - 1 do
          CreerTables(PTableRec(FTables.Objects[N])^);
        N :=  FTables.Count;
      except on e: exception do
        begin
          S := Format('Probleme sur l''ouverture de la table %s',
            [PTableRec(FTables.Objects[N])^.FTable.TableName]);
          MessageBox(0, PChar(S), '', MB_OK);
          ShowMessage(e.message);
          for N := FTables.Count - 1 downto 0 do
            PTableRec(FTables.Objects[N]).FTable.Close;
          Result := False;
          exit;
        end;
      end;
      { fin d'ouverture des tables }
      DecimalSeparator := '.'; // fixe la valeur du point decimal
      DateSeparator := '/'; // fixe la valeur du point decimal
      with FFields do
        for N := 0 to Count - 1 do
        begin { TField }
        // si champ temporaire
          if (Champ[N].binterne) then continue;
          aFieldRec := TFieldRec(Fields[N]);
          index := FTables.IndexOf(aFieldRec.FTableName);
          aTableRec := PTableRec(FTables.Objects[index]);
          aFieldRec.FField :=
            aTableRec.FTable.FieldByName(aFieldRec.FFieldName);
          AfieldRec.sel := Champ[N].sel;
        end;
    except on e: Exception do
      begin
        ShowMessage(Format('%s'#10#10'%s', [e.message, aFieldRec.FTableName]));
        for N := FTables.Count - 1 downto 0 do
          PTableRec(FTables.Objects[N]).FTable.Close;
        raise;
      end; // end Except
    end; // end begin

    for N := 0 to Champ.Count - 1 do
    begin
      if Champs[N].FTypeInfo = tiConcat then
        Champs[N].FLon := 0;
    end;
    try
      AFileName := FileName;
      // ShowMessage(Format('Executer Count=%d',[CallbackDlg.CountLu]));
      // voir table de correspondance pour la révision impfic.dll FentreeExt Tsringlist en Htstringlist
      if not impfic2 then
        ImportFichier(Longint(Self), PChar(AFileName), @SR, @TScript.LaCallBack)
      else
        SVImport.ImportFichier(Longint(Self), PChar(AFileName), Self,
          @TScript.LaCallBack);
    finally { Fermeture des tables }
      if Assigned(FOnNextRecord) then
        FOnNextRecord(self, 2, nil, False, POL^);
      for N := FTables.Count - 1 downto 0 do
        PTableRec(FTables.Objects[N]).FTable.Close;
    end;
  finally
    DateSeparator := PrevDateSeparator;
    DecimalSeparator := PrevDecimalSeparator;
    if impfic2 then
      SVImport.free;
    Freelibrary(DMImport.HdlImpfic);
    DMImport.HdlImpfic := loadlibrary(PChar('IMPFIC7.dll'));
  end;
  Result := True;
end;

// -------------------- PrepareTableChampsASCII -------------------------------
//*****************************************************************************

procedure TScript.PrepareTableChampsASCII;
var
  N, Offset: Integer;
  AFieldRec: TFieldRec;
  ATablerec: PTableRec;
begin
  (* suppression des anciennes listes *)
  for N := FTables.Count - 1 downto 0 do
  begin
    ATablerec := PTableRec(FTables.Objects[N]);
    ATableRec^.FFields.Free;
    ATableRec^.FTable.free;
    Dispose(ATableRec);
  end;
  FTables.Clear;
  for N := FFields.Count - 1 downto 0 do
    TFieldRec(FFields[N]).Free;
  FFields.Clear;

  New(ATablerec);
  ATablerec.FTable := DMImport.ADOTableD;
  ATablerec.FTable.TableName := ParName;
  if FileDest = 0 then
    ATablerec.FTable.TableName := ExtractFileName(ParName);
    ATablerec.FTable.Connection := DMImport.DBGLOBALD;

  ATablerec.FFields := TFieldList.Create;
  FTables.AddObject(ParName, TObject(ATablerec));

  (* parcours du script pour creer la liste des champs *)
  Offset := 0;
  for N := 0 to FChamps.Count - 1 do
  begin
    AFieldRec := TFieldRec.Create;
    with AFieldRec do
    begin
      FName := Champs[N].FName;
      FTableName := ParName;
      FFieldName := Champs[N].FName;
      FOffset := Offset;
      // modif me FLon := Champs[N].FLon;
      FTyp := Champs[N].FTyp;
      FLon := Champs[N].FSiz;
      FNbdecimal := Champs[N].FNbdecimal; // ajout me
      FbArrondi  := Champs[N].FbArrondi;
      Inc(Offset, FLon);
    end;
    FFields.Add(AFieldRec);
    ATablerec.FFields.Add(AFieldRec);
  end;
end;

procedure TScript.ConstruitVariable(bReplace: Boolean);
var
  I: integer;
begin
  Variable.Sorted := false;
  for I := 0 to Variables.count - 1 do
  begin
    if (Variable.indexOfName(Variables[I].Name) = -1) then
    begin // la variable n'est pas encore ajoutée
      Variable.add(Variables[I].Name + '=' + Variables[I].text);
      if (itrace > 2) then
        ShowMessage('variable ajoutée : ' + Variables[I].Name + '=' +
          Variables[I].text);
    end
    else if bReplace then // la variable existe deja
      Variable.Values[Variables[I].Name] := Variables[I].text;
  end;
  Variable.Sorted := true;
end;

procedure TScript.InterpreteSyntax;
  procedure ParcoursChampsInterpreteVar(slChamp: TChampList);
  var
    N: Integer;
    CH: TChamp;
  begin
    for N := 0 to slChamp.Count - 1 do
    begin
      CH := slChamp[N];
      if CH.TypeInfo = tiConcat then
        ParcoursChampsInterpreteVar(CH.Concat);
      if (CH.Calcul <> nil) and (CH.Calcul.FFormule <> '') then
        CH.Calcul.FFormule := InterpreteVarSyntax(CH.Calcul.FFormule, Variable);
      if (CH.FReference <> nil) and (CH.FReference.FCondition <> '') then
        // AKE 28.07.99
        CH.FReference.FCondition :=
          InterpreteVarSyntax(CH.FReference.FCondition, Variable);
    end;
  end;
begin
  Options.condition := InterpreteVarSyntax(Options.condition, Variable);
  ParcoursChampsInterpreteVar(self.FChamps);
end;

procedure  TScript.OuvrirTables(ATable: TADOTable; N: Integer; EcrOrBal: Integer);
begin

end;

// ******************** CreerTables *********************************
// ******************************************************************
function AccessType(fd : TFieldType; Size : integer):string;
begin
 case fd of
  ftString: Result:='TEXT('+IntToStr(Size)+')';
  ftWideString : Result:='TEXT('+IntToStr(Size)+')';
  ftSmallint: Result:='SMALLINT';
  ftInteger: Result:='INTEGER';
  ftWord: Result:='WORD';
  ftBoolean: Result:='YESNO';
  ftFloat : Result:='FLOAT';
  ftCurrency: Result := 'CURRENCY';
  ftDate, ftTime, ftDateTime: Result := 'DATETIME';
  ftAutoInc: Result := 'COUNTER';
  ftBlob, ftGraphic: Result := 'LONGBINARY';
  ftMemo, ftFmtMemo: Result := 'MEMO';
 else
  Result:='MEMO';
 end;
end;

procedure TScript.CreerTables(var ATable: TTableRec);
var
S : string;
  procedure DefFields(Fields: TFieldList);
  var
    N          : Integer;
  begin
    for N := 0 to Fields.Count - 1 do
    begin
      if (Champ[N].binterne) then continue;
      if ATable.FTable.FieldDefs.indexof (Champ[N].Fname) > 0 then continue;
      with Fields[N] do
      try
        case FTyp of
        // string
          0,2: ATable.FTable.FieldDefs.Add(FNAME, ftTypes[FTyp], FLon, False);

          1: ATable.FTable.FieldDefs.Add(FNAME, ftTypes[FTyp], 0, False); //float
          253: ATable.FTable.FieldDefs.Add(FNAME, ftTypes[FTyp], 0, False);
        end;
        if (Ftyp <> 0) and (Ftyp <> 2) then FLon := 0;
         s:=s + ' ' + FNAME;
         s:=s + ' ' + AccessType(ftTypes[FTyp], Flon);
         s:=s + ',';

      except
        if (iTrace = 2) then
          ShowMessage('Champ ' + FName + ' : ' +
            Exception(EXceptObject).Message);
      end;
    end;
    s[Length(s)]:=')';
  end;

begin { CreerTables }
  with ATable do
  begin { TTable TDatabase }
    FTable.Connection := DMImport.DBGLOBALD;
    FTable.TableName := ParName;
    { definition des zones }
    FTable.FieldDefs.Clear;
    FTable.IndexDefs.Clear;
    DefFields(FFields);
    try
     DMImport.CmdDonnee.CommandText:='DROP TABLE ' + ParName;
     DMImport.CmdDonnee.Execute;
    except end;

    try;
     DMImport.CmdDonnee.CommandText:='CREATE TABLE ' + ParName + ' (' + S;
     DMImport.CmdDonnee.Execute;
    except on e: Exception do
        Application.ShowException(e);
    end;
    try
      FTable.Open;
    except on e: Exception do
        Application.ShowException(e);
    end;
  end;
  if (FileDest = 2) then // sortie ODBC
  begin
    if ((pos('\', options.asciiFilename) <> 0) or (pos('/',
      options.asciiFilename) <> 0)) and (FileDest = 0) then
      options.asciiFilename := extractFileName(options.asciiFilename);
  end;
end;

// ********************** clonage de table ***********************************
// ***************************************************************************

procedure TScript.CloneTable(from, dest: TTable);

begin
{$IFNDEF DBXPRESS}
  try
    dest.open;
    dest.close;
  except // problème sur l'ouverture donc la table n'existe pas
    dest.FieldDefs.Assign(from.FieldDefs);

    from.IndexDefs.Update;
    dest.IndexDefs.Assign(from.IndexDefs);

    try
      dest.CreateTable;
    except on e: EDBEngineError do
      begin
        case e.errors[0].errorCode of
          10245: // DBIERR_FILELOCKED
            ShowMessage(format(('La table %s est vérouillée par une autre application'#13#10'le traitement ne peut continuer. '), [Dest.TableName]));

          10024: // DBIERR_NOSUCHTABLE
            ShowMessage(format(('La table %s n''existe pas'#13#10'le traitement ne peut continuer.'), [Dest.TableName]));

          10243: // DBIERR_FILEISOPEN
            ShowMessage(format(('La table %s est ouverte par une autre application'#13#10'le traitement ne peut continuer.'), [Dest.TableName]));

        else
          ShowMessage(e.Message);

        end;
        // ShowMessage(e.Message);
        raise;
      end;
    end; (* Except *)
    {        except
                  ShowMessage(Exception(ExceptObject).message);
            end;   }
  end; (* Except *)
{$ENDIF}
end;

// ------------------------- SCRUTER  -----------------------------------                                                           *)
// **********************************************************************

function TScript.Scruter(FChamp : TChamp; FScript : TScript; AonNextRecord: TNextEvent): string;
var
	TempCond : string;
        N : integer;
	CreateScruteList : PCreateScruteList;
	szFileName : array[0..260] of char;
  SL : TScruterList;
  CurrentPath : String;
  SR :TScriptRecord;

begin
//    if Assigned(SL) then SL.Free;
//	SL := nil;	(* -- ne pas supprimer *)
	if pos('$', FScript.Options.FileName) <> 0 then
	begin
	   ShowMessage('Le nom du fichier est incorrect ('+FScript.Options.FileName+').');
	   exit;
	end;
  // Initialisation de record pour la Dll
  SR.FSignature := FSignature;
  SR.FOptions   := FOptions;
  SR.FOptionCpt := FOptionCpt;
  SR.FChamps    := FChamps;
  SR.FTableCorr := FTableCorr;
  SR.FName      := FName;
  SR.FFileType  := FFileType; (* 0:base 1:ASCII *)
  SR.FASCIIMODE := FASCIIMODE; (* 0:Fixe 1:Delimite *)
  SR.FDestTable := FDestTable; (* 0:Dossier 1:Utilisateur *)
  SR.FScriptSuivant:= FScriptSuivant;
  SR.FParName   := FParName; (* CLE PARAMETRE  *)
  SR.FVariable  := FVariable;
  SR.FVariableList:= FVariableList;
  SR.FParOrig   := FParOrig;


	try
		@CreateScruteList := GetProcAddress(DMImport.HdlImpfic, 'CreerScruterList');
		if not Assigned(@CreateScruteList) then
			raise Exception.Create('fonction "CreerScruterList" non trouvée');

		if FChamp.TypeInfo = tiReference then
		begin  // remplacement de la condition pour les references
		  TempCond := FScript.Options.Condition;
		  FScript.Options.condition := Fchamp.Reference.Condition;
		end;
		SL := CreateScruteList;
		try
                  CurrentPath := CurrentDossier;
                  GetDir(0,CurrentPath); CurrentPath := CurrentPath + '\' + FChamp.Name + '.tmp';
                  for  N:=1  to 260 do
                    szFileName[N-1] := CurrentPath[N];
                  SL.Scrute(szFileName, Integer(Self), @SR, FChamp, AOnNextRecord);
		except
		  if (itrace > 0) then ShowMessage(Exception(ExceptObject).message);
		end;

		 if Fchamp.TypeInfo = tiReference then
			FScript.options.condition := TempCond;

	finally
//		if Assigned(SL) then
//			SL.Free;
	end;
    Result := szFileName;
end;

{$ifdef AUCASOU}
function TScript.Scruter(AObject: TObject; AChamp: TChamp;
  AOnNextRecord: TInfoEvent): Boolean;
type
  TChercheTableCorr = function(uUserData: Longint; pInfo: Pointer;
    AChamp: TChamp; FCallBack: TCallBack): Integer; stdcall;
var
  hImpFic: THandle;
  ChercheTableCorr: TChercheTableCorr;
begin
  (* CHARGEMENT DE LA LIBRAIRIE C++ *)
  Result := False;
  if DMImport.HdlImpfic <> 0 then
  begin
    @ChercheTableCorr := GetProcAddress(DMImport.HdlImpfic, 'ChercheTableCorr');
    try
      if Assigned(ChercheTableCorr) then
      begin
        ChercheTableCorr(Longint(AObject), Self, AChamp,
          Pointer(@AOnNextRecord));
        Result := True;
      end;
    finally
      FreeLibrary(hImpFic);
    end;
  end;
end;
{$endif}

// ---------------------- duplication du script ------------------------------
// ---------------------------------------------------------------------------

function TScript.CloneScript: TScript;
begin
  Result := TScriptClass(self.ClassType).Create(nil);
end;



function TScript.Callback(ANumEnr: Longint; AEtape: Integer): integer; stdcall;
begin
  Application.ProcessMessages;
  //	NumEnr := ANumEnr;
  if bAbort then
    Result := -1
  else
    Result := 1;
end;

function TScript.GetAxe: boolean;
begin
  result := (Variable.indexofName('AXE') <> -1) and GestionAxe;
end;

function TScript.MajTableCorr(aChamp: TChamp; AonNextRecord: TNextEvent):
  TStringList; stdcall;
var // vérifie si les tables sont toujours les mêmes
  S, TempCond: string;
  StrG: TStringList;
  f: TextFile;
  SL: TScruterList;
  //	CreateScruteList : PCreateScruteList;
  T: integer;
  himpfic: THandle;
  c: Tchamp;
  szFileName: array[0..260] of char;
  I: integer;

  function diff(str1, str2: TStringList): boolean;
  var
    i: integer;
  begin
    result := false;
    for i := 0 to str1.count - 1 do
      if str2.indexof(str1.strings[i]) = -1 then
        Result := true;
          // il manque des items dans srt2, il faut refaire les table de corr.
  end;
begin
  c := achamp;
  if not assigned(listTbl) then
    ListTbl := TStringList.create; // gestion multiplicité des tbl
  I := ListTbl.Indexof(c.NomTableExt);
  if I = -1 then
  begin //la table n'a pas été parcouru
    hImpFic := 0;
    if impfic2 then
      SL := CreateScruteList
    else
    begin
      SL := nil; (* -- ne pas supprimer *)
      @CreateScruteList := GetProcAddress(DMImport.HdlImpfic, 'CreerScruterList');
      if not Assigned(@CreateScruteList) then
        raise Exception.Create('fonction "CreerScruterList" non trouvée');
      SL := CreateScruteList;
    end;

    StrG := TStringList.Create;

    Result := Strg;
    if c.TypeInfo = tiReference then
    begin // remplacement de la condition pour les references
      TempCond := Options.Condition;
      Options.condition := c.Reference.Condition;
    end;
    strcopy(szFileName, PChar(FileName));
    try
      T := SL.Scrute(szFileName, Integer(Self), self, c, AonNextRecord);
      //        except if itrace > 0 then ShowMessage('TBLCORR (ScruteList) : '+Exception(ExceptObject).Message); raise; end;
      if T = -1 then
        exit; //erreur
      if c.TypeInfo = tiReference then
        Options.condition := TempCond;

      AssignFile(f, szFileName);
      Reset(f);
      while not eof(f) do
      begin
        readln(f, s);
        S := trim(S);
        if StrG.IndexOf(S) < 0 then
          StrG.add(S);
      end;
      CloseFile(f);

      // préparation de la table de corresp. du champ en cours
      if c.TableExterne then
        c.IniTableCorr(Variable); //récup. de la table depuis le fichier externe

      // comparer les 2 tables de corresp.
      result := Strg;
      if C.TableExterne then
      begin // chargement des tables de corr.
        C.TableCorr^.FEntree.Clear;
        C.TableCorr^.FSortie.Clear;
        C.TableCorr^.FEntree.Assign(C.TableCorr^.FEntreeExt);
        C.TableCorr^.FSortie.Assign(C.TableCorr^.FSortieExt);
        C.TableCorr^.FCount := C.TableCorr^.FEntree.Count;
      end;
      if (himpfic <> 0) or Impfic2 then
        if diff(StrG, C.TableCorr^.FEntree) then
          result := Strg
        else
          Strg.Clear; // cette stringList est libéré par le Tscript.destroy
    except
      if itrace > 2 then
        ShowMessage('MajTblCorr : ' + Exception(ExceptObject).message);
    end;
    SL.Free;
    if not impfic2 then
      FreeLibrary(Himpfic);
    ListTbl.addObject(c.NomTableExt, Strg);
  end
  else
    result := TStringList(listtbl.objects[I]);
      // cette stringList est libéré par le Tscript.destroy
end;

{$IFDEF VAR}

function TScript.SetVariableList: TStringList;
var
  List: TStringList;
  V: TVariableDef;
  N: integer;
  function Convert(I: integer): string;
  var
    str: string;
  begin
    str := 'Edit';
    case I of
      0, 1: str := 'EDIT';
      2: str := '[' + v.FText + ']';
      3: str := 'CHECK';
      4: str := 'DLGBOX';
    end;
    result := str;
  end;
begin
  List := TStringList.create;
  for N := 0 to FVariableList.Count - 1 do
  begin
    V := FVariableList.Items[N];
    List.add(v.FName + '=' + v.FLibelle + ',' + v.FText + ',' +
      Convert(v.FTypeVar));
  end;
  Result := List;
end;
{$ENDIF VAR}

function TScript.TraiteCoherence: boolean;
var
  I: integer;
  CC: Tcondcoherence;
  F: TextFile;
  sLigne: string;
  hdll: integer;
  aSVparser: TSVParser;
begin
  Result := (Coherence.count = 0); hdll := 0;
  try
    if not impfic2 then
    begin
      Hdll := loadlibrary(PChar(IMPFICDLL));
      if Hdll = 0 then
        raise Exception.Create('la librarie n''a pas été trouvée');
      @CreerSVParser := GetProcAddress(hdll, 'CreerSVParser');
      if not Assigned(CreerSVParser) then
        raise
          Exception.Create('la fonction "CreerSVParser" n''a pas été trouvée');
    end;

    aSVParser := CreerSVParser;
  //  aSVParser.setScript(self);
    if impfic2 then
      Hdll := 0;
  except
    if not impfic2 then
      freelibrary(hdll);
    if (itrace > 0) then
      showMessage(Exception(ExceptObject).message);
    Result := true; //Pour ne pas bloquer l'importation
    raise
      Exception.Create('Le traitement de la cohérence ne s''est pas déroulé correctement.(' + Exception(ExceptObject).message + ')');
  end;

  // récupérer la première Ligne du fichier
  assignFile(F, options.fileName);
  reset(F);
  ReadLn(F, sLigne);
  close(F);

  for I := 0 to Coherence.count - 1 do
  begin
    CC := TCondCoherence(Coherence.items[I]);
    Result := CC.Evalue(aSVparser, PChar(sLigne));
    if ((not result) and (CC.sMessage <> '')) then
      ShowMessage(CC.sMessage);
    if (not Result) and (CC.bstop) then
      break;
  end;
  ASVParser.Free;
  if not impfic2 then
    freelibrary(hdll);
end;

procedure TScript.BdeToOdbc(DataSet: TDataset; dsDest: TDataset);
begin
end;

destructor TFourchette.Destroy;
begin
  inherited Destroy;
end;

function TFourchette.evalue(str: string): boolean;
begin // évalue la chaine fournit par rapport à la règle définie
  if deb = fin then
    Result := copy(str, 1, length(deb)) = deb
  else // Chaine entre 'Deb' et 'Fin'
    Result := (copy(str, 1, length(deb)) >= PChar(deb)) and (copy(str, 1,
      length(deb)) <= PChar(fin));
end;

function TInfocpt.EvalueRacine(compte: string): integer;
var
  i: integer;
begin
  Result := -1;
  for I := 0 to ListFourchette.count - 1 do
  begin
    if TFourchette(Listfourchette.items[I]).evalue(compte) then
    begin
      Result := I;
      exit;
    end;
  end;
end;

constructor TInfoCpt.create(Collection: Tcollection);
begin
  inherited create(collection);
  listfourchette := Tlist.create;
end;

destructor TInfoCpt.destroy;
begin
  if assigned(ListFourchette) then
    Listfourchette.Free;
  inherited;
end;
//---------------
// TListAxe
// --------------

constructor TAxeAna.Create(Collection: TCollection);
begin
  inherited Create(Collection);
end;

destructor TAxeAna.Destroy;
begin
  inherited Destroy;
end;

procedure TAxeAna.Assign(Value: TPersistent);
begin
  NomAxe := TAxeAna(Value).Nomaxe;
  NumAxe := TAxeAna(Value).Numaxe;
  Condition := TAxeAna(Value).Condition;
  sel := TAxeAna(Value).Sel;
end;

constructor TListAxe.Create;
begin
  inherited Create(TAxeAna);
end;

procedure TListAxe.Assign(AList: TListAxe);
var
  I: Integer;
begin
  Clear;
  for I := 0 to AList.Count - 1 do
    Add.Assign(AList.items[I]);
end;

function TListAxe.Add: TAxeAna;
begin
  Result := TAxeAna(inherited Add);
end;

procedure TListAxe.Delete(Index: Integer);
begin
  TAxeAna(Items[Index]).free;
end;

function TListAxe.IndexOf(AValue: string): integer;
var
  N: integer;
begin
  for N := 0 to Count - 1 do
    if AnsiCompareText(TAxeAna(Items[N]).NomAxe, AValue) = 0 then
    begin
      Result := N;
      exit;
    end;
  Result := -1;
end;

procedure TListaxe.OrganizeByNumAxe;
var
  last: integer;
begin
  last := count - 1;
  while (last > 0) do
  begin
    while (TaxeAna(Items[Last]).Numaxe < TaxeAna(Items[Last - 1]).Numaxe) do
      move(Last, last - 1);
    last := last - 1;
  end;
end;

procedure TListAxe.Move(curIndex, NewIndex: Integer);
var
  TempItem: TAxeana;
begin
  TempItem := TAxeana.create(nil);
  TempItem.assign(TAxeAna(Items[NewIndex]));
  TAxeana(Items[NewIndex]).assign(TAxeAna(Items[curIndex]));
  TAxeAna(Items[curIndex]).assign(TempItem);
  TempItem.free;
end;

// -----------------------------------------------------------------------------
//                       T C O H E R E N C E
// -----------------------------------------------------------------------------

constructor TCondCoherence.create(NomCoherence: string; condition: string;
  Collection: TCollection);
begin
  sCondition := condition;
  Nom := NomCoherence;
  inherited create(Collection);
end;

procedure TCondCoherence.Assign(Value: TPersistent);
begin
  Nom := TCondCoherence(Value).Nom;
  sCondition := TCondCoherence(Value).sCondition;
  sMessage := TCondCoherence(Value).sMessage;
  bStop := TCondCoherence(Value).bStop;
end;

function TCondCoherence.Evalue(SV: TSVParser; sText: PChar): boolean;
begin // Il faut que le setscript soit deja fait
  if SV = nil then
  begin
    Result := FALSE; exit;
  end;
  SV.SetText(PChar(sCondition));
  if (itrace > 0) then
    ShowMessage('Coherence 1 : ligne[' + sText + ']');
  Result := (SV.IsTextSelect(sText) = 1);
end;

constructor TCoherence.create;
begin // creation de la liste TCoherence
  inherited create(TCondCoherence);
end;

destructor TCoherence.Destroy;
begin
  inherited destroy;
end;

procedure TCoherence.delete;
begin
  TCondCoherence(Items[Index]).free;
end;

function TCoherence.Add: TCondCoherence;
begin
  Result := TCondCoherence(inherited Add);
end;

procedure TCoherence.Assign(AList: TCoherence);
var
  I: Integer;
begin
  Clear;
  for I := 0 to AList.Count - 1 do
    Add.Assign(AList.items[I]);
end;

function TCoherence.IndexOf(AValue: string): integer;
var
  N: integer;
begin
  for N := 0 to Count - 1 do
    if AnsiCompareText(TCondCoherence(Items[N]).Nom, AValue) = 0 then
    begin
      Result := N;
      exit;
    end;
  Result := -1;
end;

function IsEmpty(S: string): Boolean;
var
  P: PChar;
begin
  P := PChar(S);
  while P^ = ' ' do
    Inc(P);
  IsEmpty := P^ = #0;
end;

FUNCTION ADroite (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=C+St1 ;
Result:=st1 ;
END ;

FUNCTION AGuauche (st : string; l : Integer ; C : Char) : string ;
var St1 : String ;
BEGIN
st1:=Trim(st) ;
if ((l>0) and (l<Length(St1))) then St1:=Copy(St1,1,l) ;
While Length(St1)<l Do St1:=St1+C ;
Result:=st1 ;
END ;

function Estalpha (Chaine : string) : Boolean ;
var
  j: integer;
begin
    Result := TRUE;
         for j:=1 To  length(Chaine) do
         begin
             if Chaine[j]<> ' ' then
             if Chaine[j] in Alpha then exit;
             if (Chaine[j] in ['|','!'..'/',':'..'`']) and (Chaine[j] <> '.') then exit;
         end;
    Result := FALSE;
end;

Function ConvertDate(St : String) : TDateTime ;
Var D,M,Y : Word ;
BEGIN
Result:=iDate1900 ;
if Trim(St)='' then Exit ;
if length(St) < 6 then exit;
Y:=StrToInt(Copy(St,5,2)) ; M:=StrToInt(Copy(St,3,2)) ; D:=StrToInt(Copy(St,1,2)) ;
If Y<90 Then Y:=2000+Y Else Y:=1900+y ;
Result:=EncodeDate(Y,M,D) ;
END ;

Procedure FaitSISCOPer(Var FF : TextFile ; Datefin, Datedeb : TDateTime) ;
Var St : String ;
    Y,M,D,Y1,M1,D1 : Word ;
    i : Integer ;
BEGIN
Decodedate(Datedeb,Y,M,D) ; Decodedate(Datefin,Y1,M1,D1) ; i:=12*(Y1-Y)+M1-M+1 ;
St:='M'+FormatFloat('00',i) ;
WriteLn(FF,St) ;
END ;

procedure ExecuteShell (Executable : string);
var
Rep    : string;
Cmd,St : string;
Dir    : string;
begin
       Cmd := Executable;
       St := ReadTokenPipe (Cmd,' ');
       Dir := RecomposeChemin(St);
       Rep := ExtractFileDir(Dir);
       if DirectoryExists(Rep) then ChDir(Rep);
       Cmd := Dir+' '+Cmd;
// exemple C:\tmp\00002\AppelDos.bat Fich1.txtfich2.txt ici.txt
       if (pos('.EXE', Uppercase(Cmd)) = 0) then
       begin
           if not ChargeFile (Cmd, 1) and (pos('.BAT', Uppercase(Cmd)) <> 0) then
              Winexec (PCHAR(Cmd) ,SW_SHOWNORMAL);
       end
       else
              Winexec (PCHAR(Cmd) ,SW_SHOWNORMAL);
end;

procedure TScriptCorresp.SaveTo(AStreamTable: TStream);
var
  ST1, ST2: TStream;
  WR: TWriter;
begin (* SaveTo *)
  ST1 := TMemoryStream.Create;
  WR := TWriter.Create(ST1, 4096);
  try
    WR.WriteRootComponent(Self);
  except
    WR.Free;
    ST1.Free;
    ShowMessage(Exception(ExceptObject).message);
    Exit;
  end;
  WR.Free;

  ST1.Seek(0, 0);
  ST2 := AStreamTable;
  try
    ObjectBinaryToText(ST1, ST2);
  finally
    ST1.Free;
  end;
end;

destructor TScriptCorresp.Destroy;
begin
    LFEntree.free;
    LFSortie.free;
    LFEntree := nil;
    LFSortie := nil;
  inherited Destroy;
end;

constructor TScriptCorresp.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  LFEntree := TStringList.Create;
  LFSortie := TStringList.Create;
end;


function LoadScriptCorresp(AStream: TStream): TScriptCorresp;
var
  AScriptc: TScriptCorresp;
  RD: TMonReader;
  ST1, ST2: TStream;
begin
  ST1 := AStream;
  ST2 := TMemoryStream.Create;
  try
    ObjectTextToBinary(ST1, ST2);
  except
    ST2.Free;
    raise;
  end;
  ST2.Seek(0, soFromBeginning);
  RD := TMonReader.Create(ST2, 4096);
  try
    AScriptc := RD.ReadRootComponent(nil) as TScriptCorresp;
  except
    RD.Free;
    ST2.Free;
    raise;
  end;
  RD.Free;
  ST2.Free;
  Result := AScriptc;
end;

function CreerLigPop ( Name : string ; Owner : TComponent) : TMenuItem ;
Var
    T   : TMenuItem ;
BEGIN
T:=TMenuItem.Create(nil) ; T.Name:=Name ;
T.Caption:=Name ;
Result:=T ;
END ;

{$IFNDEF EAGLCLIENT}
{$IFNDEF CISXPGI}
FUNCTION  CHARGECIX (FichierIni : string=''): Boolean ;
var
Commande, tmp       : string;
sTempo              : string;
FicIni              :TIniFile;
CurrentDir          : string;
ListeFile,Domaine   : string;
TFV                 : TOB;
begin
  Result := TRUE; FicIni := nil;
     if (FichierIni <> '') or (pos ('/INI=', ParamStr(1)) <> 0)then  // si lancement par la ligne de commande
     begin
          if FichierIni = '' then
          begin
             Commande := ParamStr(1);
             tmp := ReadTokenPipe (Commande, ';');
             CurrentDir := ReadTokenPipe (Commande, ';');
             if CurrentDir = '' then CurrentDir := ParamStr(1);
          end
          else
             currentDir := FichierIni;

          if (Copy(currentDir, 1, 5) = '/INI=') then
          begin
               sTempo := Copy(currentDir, 6, length(CurrentDir));
               if not FileExists(sTempo) then
               begin
                    PGIInfo ('Le fichier '+ sTempo+ ' n''existe pas','');
                    FicIni.free;
                    Result := FALSE;
                    exit;
               end;
               FicIni        := TIniFile.Create(sTempo);
               New(VHCX) ; FillChar(VHCX^,Sizeof(VHCX^),#0) ;
               VHCX^.Ini        := sTempo;
               VHCX^.Domaine    := FicIni.ReadString ('PRODUIT', 'Domaine', '');
               VHCX^.Mode       := FicIni.ReadString ('PRODUIT', 'Mode', '');
               VHCX^.Nature     := FicIni.ReadString ('PRODUIT', 'NATURE', '');
               VHCX^.Complement := FicIni.ReadString ('PRODUIT', 'COMPLEMENT', '');
               VHCX^.Option     := FicIni.ReadString ('PRODUIT', 'OPTION', '');
               VHCX^.Famille    := FicIni.ReadString ('PRODUIT', 'FAMILLE', '');
               VHCX^.Appelant   := FicIni.ReadString ('PRODUIT', 'APPELANT', '');
               //VHCX^.RepSTD     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDossier     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDonnee   := FicIni.ReadString ('PRODUIT', 'REPDOS', '');
               if not DirectoryExists(CurrentDonnee) then CreateDir(CurrentDonnee);

               VHCX^.LISTECOMMANDE  := FicIni.ReadString ('PRODUIT', 'LISTECOMMANDE', '');
               VHCX^.ATob := nil;
               if VHCX^.LISTECOMMANDE <> '' then
               begin
                     ListeFile := VHCX^.LISTECOMMANDE;
                     Commande  := ListeFile;
                     Domaine   := VHCX^.Domaine;
                     While Commande <> '' do
                     begin
                          Commande    := ReadTokenPipe(ListeFile, ';');
                          if Commande <> '' then
                          begin
                               if VHCX^.ATob = nil then VHCX^.ATob := TOB.Create('COMMANDE', nil,-1);
                               TFV := TOB.Create(Commande, VHCX^.ATob, -1);
                               TFV.AddChampSupValeur('REPERTOIRE', FicIni.ReadString (Commande, 'REPERTOIRE', ''));
                               TFV.AddChampSupValeur('SCRIPT', FicIni.ReadString (Commande, 'SCRIPT', ''));
                               TFV.AddChampSupValeur('NOMFICHIER', FicIni.ReadString (Commande, 'NOMFICHIER', ''));
                               TFV.AddChampSupValeur('LISTEFICHIER', FicIni.ReadString (Commande, 'LISTEFICHIER', ''));
                               TFV.AddChampSupValeur('MONOTRAITEMENT', FicIni.ReadString (Commande, 'MONOTRAITEMENT', ''));
                               TFV.AddChampSupValeur('COMPRESS', FicIni.ReadString (Commande, 'COMPRESS', ''));
                               Domaine := ReadTokenPipe(VHCX^.Domaine, ';');
                               TFV.AddChampSupValeur('DOMAINE', Domaine);
                          end;
                     end;
               end
               else
               begin
                              VHCX^.Directory  := FicIni.ReadString ('COMMANDE', 'REPERTOIRE', '');
                              VHCX^.Script     := FicIni.ReadString ('COMMANDE', 'SCRIPT', '');
                              VHCX^.NomFichier := FicIni.ReadString ('COMMANDE', 'NOMFICHIER', '');
                              VHCX^.ListeFichier:= FicIni.ReadString ('COMMANDE', 'LISTEFICHIER', '');
                              VHCX^.Monotraitement := FicIni.ReadString ('COMMANDE', 'MONOTRAITEMENT', '');
                              VHCX^.Compress       := FicIni.ReadString ('COMMANDE', 'COMPRESS', '');
               end;
               FicIni.free;
          end;
     end
     else
     begin
                    New(VHCX) ; FillChar(VHCX^,Sizeof(VHCX^),#0) ;
     end;
end;

procedure LibereCX;
begin
     if assigned(VHCX) then
     begin
          if VHCX^.ATOB <> nil then
          begin VHCX^.ATOB.free; VHCX^.ATOB := nil end;
          Dispose(VHCX) ;
     end;
end;
{$ENDIF}
{$ENDIF}
                                    // Dossier = Repertoire+ extension
Procedure DeleteFichierRepertoire (Dossier : string; Attributs: integer);
var
          Resultat       : Integer;
          SearchRec      : TSearchRec;
begin
          Resultat := FindFirst (Dossier, Attributs, SearchRec);
          while Resultat = 0 do
          begin
            Application.ProcessMessages;
            if ((SearchRec.Attr and faDirectory) <= 0) then
              DeleteFile (SearchRec.Name);
             Resultat:= FindNext (SearchRec);
          end;
end;

Function FusionFichierRepertoire (Dossier: string; Attributs: integer) : string;
var
  FichierTrouve: string;
  Resultat       : Integer;
  SearchRec      : TSearchRec;
  St             : string;
  Extent,tmp     : string;
  Chemin         : string;
  F,FS           : TextFile;
begin
  tmp := ReadTokenPipe(Dossier, ' ');
  Chemin := ExtractFileDir(tmp);
  Extent := Copy(Dossier, pos('.', Dossier), length(Dossier));
  if Extent = '' then Extent := '.Fusion';
  DeleteFichierRepertoire (Chemin+'\*'+Extent, Attributs);


  Resultat := FindFirst (tmp, Attributs, SearchRec);
  while Resultat = 0 do
  begin
    Application.ProcessMessages;
    if ((SearchRec.Attr and faDirectory) <= 0) then // On a trouvé un Fichier (et non un dossier)
    begin
      FichierTrouve := SearchRec.Name;
      AssignFile(FS,  Chemin+'\'+Copy(FichierTrouve, 1,5)+Extent) ;
      if FileExists(Chemin+'\'+Copy(FichierTrouve, 1,5)+Extent)  then
         Append (FS)
      else
         rewrite(FS) ;
      AssignFile(F, FichierTrouve);
      {$I-} Reset (F) ; {$I+}
      While Not EOF(F) do
      begin
              Readln(F, St);
              writeln(FS, St);
      end;
      CloseFile(F);
      CloseFile(FS);
    end;
     Resultat:= FindNext (SearchRec);
  end;
  Result := '*'+ Extent;
end;

function ChargeFile (Command : string; indice:integer) : Boolean;
var
  F,FS              : TextFile;
  S,Chemin          : string;
  Tmp,ff            : string;
  P, pch            : PChar;
  Fichier           : string;
  Lines             : TStringList;

begin
Result := FALSE;
if (pos('.BAT', Uppercase(Command)) <> 0) then
begin
      Tmp := ReadTokenPipe(Command, ' ');
      AssignFile(F, Tmp);
      {$I-} Reset (F) ; {$I+}
      if EOF(F) then Result := FALSE;
      if Not EOF(F) then
      begin
              Readln(F, S);
              Tmp := ReadTokenPipe(S, ' ');
      end;
      CloseFile(F);
      if pos('SCRUTE', Uppercase(Tmp)) <> 0  then
      begin
           if S = '' then exit;
           FusionFichierRepertoire (S, FaDirectory + faHidden + faSysFile);
           Result := TRUE;
           exit;
      end;
      if pos('RCOPY', Uppercase(Tmp)) = 0  then exit;
      if S = '' then exit;
      Chemin := ExtractFileDir(S);
      setstring(S, Pchar(S), length(S));
end
else
begin
      Tmp := ReadTokenPipe(Command, ' ');
      if pos('RCOPY', Uppercase(Tmp)) = 0  then exit;
      if Command = '' then exit;
      Chemin := ExtractFileDir(Command);
      setstring(S, Pchar(Command), length(Command));
end;
if Chemin = '' then exit;
Lines := TStringList.Create;
P := PChar(S);
repeat
    pch := StrScan(P, '+');
    if pch <> nil then
    begin
      pch^ := #0;
      inc(pch);
      Fichier := P;
      Tmp := Chemin+'\'+ExtractFileName(Fichier);
      Lines.Add(Tmp);
    end
    else
    begin
         pch := StrScan(P, ' ');
         if pch <> nil then
         begin
            pch^ := #0;
            inc(pch);
            Fichier := P;
            Tmp := Chemin+'\'+ExtractFileName(Fichier);
            Lines.Add(Tmp);
         end
         else
            Fichier := P;
    end;
    P := pch;
until pch = nil;

AssignFile(FS,  Chemin+'\'+Fichier) ;
Rewrite(FS) ;

for indice := 0 to Pred(Lines.Count) do
begin
//RenameFile(Fichier, Fichier+'_S');
      if not FileExists(Lines.Strings[indice])  then continue;
      AssignFile(F, Lines.Strings[indice]);
      {$I-} Reset (F) ; {$I+}
      While Not EOF(F) do
      begin
              Readln(F, S);
              ff := Format('%2.2d',[indice+1]);
              writeln(FS, ff+S);
      end;
      CloseFile(F);
end;
CloseFile(FS);
Result := TRUE;
end;

{$IFNDEF CISXPGI}
{$IFNDEF EAGLCLIENT}
function InitScriptAuto(table : string) : integer;
var
STable        : TTable;
S             : TmemoryStream;
AStreamTable  : TmemoryStream;
Ascript       : TScript;
isc           : integer;
SQL,St        : string;
TF1,TOBVar    : TOB;
QP            : TQuery;
begin
STable := nil; S := nil; St := ''; AStreamTable := nil; TOBVar:= nil;
Result := -1; //Script inexistant
try
  STable := TTable.Create(Application);
  STable.DatabaseName :=  DMImport.DBGlobal.Name;
  STable.Tablename := DMImport.GzImpReq.TableName;
  STable.Open;
  AStreamTable := TmemoryStream.create;
  if Table <> '' then
  begin
       if not STable.FindKey([table]) then exit;
       s := TmemoryStream.create;
       TBlobField(STable.FieldByName('PARAMETRES')).SaveToStream (s);
       s.Seek (0,0);
       TBlobField(STable.FieldByName('TBLCOR')).SaveToStream (AStreamTable);
       AStreamTable.Seek (0,0);

       AScript := LoadScriptFromStream(s, AStreamTable);
       if AScript.Variables <> nil then
       begin
            Result := 1; // Script et variables trouvés
            TOBVar := TOB.Create('Enregistrement', nil, -1);
            For isc:=0 To AScript.Variables.count-1 do
            begin
                       St := AScript.Variables.Items[isc].Libelle;
                       St := ReadTokenpipe (St,'=');
                       TF1 := TOB.Create ('',TOBVar,-1);
                       TF1.AddChampSupValeur('Name',AScript.Variables.Items[isc].Name);
                       TF1.AddChampSupValeur('Libelle',St);
                       TF1.AddChampSupValeur('Text', AScript.Variables.Items[isc].Text);
                       TF1.AddChampSupValeur('Demandable', AScript.Variables.Items[isc].demandable);
            end;
            if VHCX^.Directory = '' then VHCX^.Directory := CurrentDonnee;
            if VHCX^.NomFichier <> '' then
               TOBVar.SaveToFile(VHCX^.Directory+'\'+VHCX^.NomFichier,True,True,True)
            else
               TOBVar.SaveToFile(VHCX^.Directory+'\'+table+'Variables.TXT',True,True,True);
       end
       else
             Result := 0;
  end
  else
  begin
      St := '';
      if VHCX^.Domaine <> '' then SQL := ' Domaine="' + VHCX^.Domaine + '" ';
      if VHCX^.Complement <> ''  then SQL := SQL + ' and CLEPAR = "' + VHCX^.Complement + '" ';
      if VHCX^.Nature <> ''  then SQL := SQL + ' and Table0 = "' + VHCX^.Nature + '" ';
      QP := TQuery.Create(nil);
      QP.DatabaseName := DMImport.DBGlobal.Name;
      if SQL <> '' then  SQL := 'AND ' + SQL;
      St := ' SELECT Table1,CLE,COMMENT,CLEPAR,DATEDEMODIF,Table0,Domaine from '+ DMImport.GzImpReq.TableName + ' Where (CLEPAR<>"SQL" or CLEPAR="") '+ SQL;
      QP.Close;
      QP.SQL.Clear;
      QP.SQL.Add(St);
      QP.Open;
      TOBVar := TOB.Create('', nil, -1);
      TOBVar.LoadDetailDB(DMImport.GzImpReq.TableName, '', '', QP, TRUE, FALSE);
      TOBVar.SaveToFile(VHCX^.Directory+'\'+VHCX^.NomFichier,True,True,True);
      Result := 0;
  end;
  Finally
    STable.Close;
    STable.Free;
    s.free;
    AStreamTable.free;
    TOBVar.free;
  end;
end;
{$ENDIF}

{$ELSE}

{$IFNDEF EAGLCLIENT}

function InitScriptAuto(table : string) : integer;
var
S             : TmemoryStream;
AStreamTable  : TmemoryStream;
Ascript       : TScript;
isc           : integer;
St            : string;
TF1,TOBVar    : TOB;
QP            : TQuery;
begin
S := nil; St := ''; AStreamTable := nil; TOBVar:= nil;
Result := -1; //Script inexistant
try
  QP := OpenSQl ('SELECT * FROM CPGZIMPREQ Where CIS_CLE="'+table+'"', TRUE);

  AStreamTable := TmemoryStream.create;
  if not QP.EOF then
  begin
       Result := 0; // Pas de variables
       s := TmemoryStream.create;
       TBlobField(QP.FindField('CIS_PARAMETRES')).SaveToStream (s);
       s.Seek (0,0);

       AScript := LoadScriptFromStream(s, AStreamTable);
       if AScript.Variables <> nil then
       begin
            Result := 1; // Script et variables trouvés
            TOBVar := TOB.Create('Enregistrement', nil, -1);
            For isc:=0 To AScript.Variables.count-1 do
            begin
                       St := AScript.Variables.Items[isc].Libelle;
                       St := ReadTokenpipe (St,'=');
                       TF1 := TOB.Create ('',TOBVar,-1);
                       TF1.AddChampSupValeur('Name',AScript.Variables.Items[isc].Name);
                       TF1.AddChampSupValeur('Libelle',St);
                       TF1.AddChampSupValeur('Text', AScript.Variables.Items[isc].Text);
                       TF1.AddChampSupValeur('Demandable', AScript.Variables.Items[isc].demandable);
            end;
(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

            if FileExists(V_PGI.DosPath + '\Echanges\Echanges_LSR.INI') then
               SysUtils.DeleteFile (V_PGI.DosPath + '\Echanges\Echanges_LSR.INI');
            TOBVar.SaveToFile(V_PGI.DosPath + '\Echanges\Echanges_LSR.INI',True,True,True);
*)
            if FileExists(TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI') then
               SysUtils.DeleteFile (TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI');
            TOBVar.SaveToFile(TcbpPath.GetCegidUserTempPath + 'Echanges\Echanges_LSR.INI',True,True,True);

       end;
  end
  Finally
    Ferme (QP);
    s.free;
    AStreamTable.free;
    TOBVar.free;
  end;
end;

function RendTobparametre  (Domaine : string; var TOBParam : TOB; crit2 : string='COMPTA'; crit3: string='PARAM'; Crit4 : string=''): integer;
var
Filename                   : string;
CodeRetour                 : integer;
entete                     : boolean;
encode                     : string;

begin
        Filename := GetWindowsTempPath +'PGI\STD\CISX\';
        if Domaine <> '' then Filename := Filename + Domaine+'\';
        if crit2 <> '' then Filename := Filename + crit2 +'\';
        if crit3 <> '' then Filename := Filename + crit3 +'\';
        if crit4 <> '' then Filename := Filename +crit4 +'\';
        Filename := Filename +  V_PGI.LanguePrinc+'\CEG\'+
        Domaine+'Compta.Cix';
        if FileExists (Filename) then
                      DeleteFile(Filename);

(* CA - 05/07/2007 - Remplacement DosPath car n'existe pas en EAGL

        if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
        Filename :=  V_PGI.DosPath +'\'+Domaine+'Compta.Cix'
        else Filename :=  V_PGI.DosPath +Domaine+'Compta.Cix';
*)
    Filename :=  TcbpPath.GetCegidUserTempPath+Domaine+'Compta.Cix';

        CodeRetour :=  AGL_YFILESTD_EXTRACT (Filename, 'CISX', ExtractFileName(Filename), Domaine, Crit2, Crit3, crit4, '', FALSE, V_PGI.LanguePrinc, 'CEG');
        if CodeRetour = -1 then
        begin
              if TOBParam <> nil then TOBParam.free;
              TOBParam := TOB.create ('', nil, -1);
              TOBParam.LoadFromXMLFile(Filename, entete, encode);
        end;
        Result := CodeRetour;
end;
{$ENDIF}
{$ENDIF}

function ParLigneCommande : boolean;
var i : integer;
    St : string;
begin
  Result := False;
  for i:=1 to ParamCount do
  begin
    St := ParamStr(i);
    if (pos('/INI', St) <> 0) then
    begin
      Result := True;
      break;
    end;
  end;
end;

{$IFNDEF EAGLCLIENT}
{$IFDEF CISXPGI}
procedure TZVarCisx.CHARGECIX(FichierIni : string='');
var
Commande, tmp       : string;
sTempo              : string;
FicIni              :TIniFile;
CurrentDir          : string;
ListeFile           : string;
TFV                 : TOB;
begin
     FicIni := nil;
     if (FichierIni <> '') or (pos ('/INI=', ParamStr(1)) <> 0)then  // si lancement par la ligne de commande
     begin
          if FichierIni = '' then
          begin
             Commande := ParamStr(1);
             tmp := ReadTokenPipe (Commande, ';');
             CurrentDir := ReadTokenPipe (Commande, ';');
             if CurrentDir = '' then CurrentDir := ParamStr(1);
          end
          else
             currentDir := FichierIni;

          if (Copy(currentDir, 1, 5) = '/INI=') then
          begin
               sTempo := Copy(currentDir, 6, length(CurrentDir));
               if not FileExists(sTempo) then
               begin
                    PGIInfo ('Le fichier '+ sTempo+ ' n''existe pas','');
                    FicIni.free;
                    exit;
               end;
               FicIni        := TIniFile.Create(sTempo);
               Ini        := sTempo;
               Domaine    := FicIni.ReadString ('PRODUIT', 'Domaine', '');
               Mode       := FicIni.ReadString ('PRODUIT', 'Mode', '');
               Nature     := FicIni.ReadString ('PRODUIT', 'NATURE', '');
               Complement := FicIni.ReadString ('PRODUIT', 'COMPLEMENT', '');
               Option     := FicIni.ReadString ('PRODUIT', 'OPTION', '');
               Famille    := FicIni.ReadString ('PRODUIT', 'FAMILLE', '');
               Appelant   := FicIni.ReadString ('PRODUIT', 'APPELANT', '');
               CurrentDossier     := FicIni.ReadString ('PRODUIT', 'REPSTD', '');
               CurrentDonnee   := FicIni.ReadString ('PRODUIT', 'REPDOS', '');
               if not DirectoryExists(CurrentDonnee) then CreateDir(CurrentDonnee);

               LISTECOMMANDE  := FicIni.ReadString ('PRODUIT', 'LISTECOMMANDE', '');
               if LISTECOMMANDE <> '' then
               begin
                     ListeFile := LISTECOMMANDE;
                     Commande  := ListeFile;
                     While Commande <> '' do
                     begin
                          Commande    := ReadTokenPipe(ListeFile, ';');
                          if Commande <> '' then
                          begin
                               if ATob = nil then ATob := TOB.Create('COMMANDE', nil,-1);
                               TFV := TOB.Create(Commande, ATob, -1);
                               TFV.AddChampSupValeur('REPERTOIRE', FicIni.ReadString (Commande, 'REPERTOIRE', ''));
                               TFV.AddChampSupValeur('SCRIPT', FicIni.ReadString (Commande, 'SCRIPT', ''));
                               TFV.AddChampSupValeur('NOMFICHIER', FicIni.ReadString (Commande, 'NOMFICHIER', ''));
                               TFV.AddChampSupValeur('LISTEFICHIER', FicIni.ReadString (Commande, 'LISTEFICHIER', ''));
                               Domaine := ReadTokenPipe(Domaine, ';');
                               TFV.AddChampSupValeur('DOMAINE', Domaine);
                          end;
                     end;
               end
               else
               begin
                              Directory  := FicIni.ReadString ('COMMANDE', 'REPERTOIRE', '');
                              Script     := FicIni.ReadString ('COMMANDE', 'SCRIPT', '');
                              NomFichier := FicIni.ReadString ('COMMANDE', 'NOMFICHIER', '');
                              ListeFichier:= FicIni.ReadString ('COMMANDE', 'LISTEFICHIER', '');
                              Monotraitement := FicIni.ReadString ('COMMANDE', 'MONOTRAITEMENT', '');
                              Compress       := FicIni.ReadString ('COMMANDE', 'COMPRESS', '');
               end;
               FicIni.free;
          end;
     end
     else
                            Domaine    := 'X';
end;

constructor TZVarCisx.Create ;
begin
     ATob := TOB.Create('COMMANDE', nil,-1);
     inherited ;
end;

destructor TZVarCisx.Destroy;
begin
  if atob <> nil then
  FreeAndNil(ATob);
  inherited ;
end;

{$ENDIF}


{$IFNDEF CISXPGI}
function GetInfoVHCX : LaVariablesCISX;
{$ELSE}
function GetInfoVHCX : TZVarCisx;
{$ENDIF}
begin
{$IFNDEF CISXPGI}
 if VHCX = nil then exit;
 Result := VHCX^;
{$ELSE}
 Result := TCPContexte.GetCurrent.VarCisx;
{$ENDIF}
end ;

{$ENDIF EAGLCLIENT}

// Recherche du script des reconnaissances
function ReconnaissanceScript (Requete : string) : string;
var
Q       : TQuery;
Index   : integer;
begin

      Q := TQuery.Create(nil);
      Q.ConnectionString := DMImport.DBGlobalD.ConnectionString;
      Q.SQL.Add (requete);
      Q.Open;
      if Q.FindField('Origine_Editeur').AsString <> '' then
         Result := Q.FindField('Origine_Processus').AsString;
      Q.Close;
      Q.Free;
end;


initialization
  RegisterClasses([TScript, TChamp, TChampList, TConstante, TCalcul,
    TReference, TScriptCorresp]);
{$IFNDEF CISXPGI}
{$IFNDEF EAGLCLIENT}
  CHARGECIX;
{$ENDIF}
{$ENDIF}
finalization
{$IFNDEF CISXPGI}
{$IFNDEF EAGLCLIENT}
  LibereCX;
{$ENDIF}
{$ENDIF}

end.




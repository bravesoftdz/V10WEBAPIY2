{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 18/10/2002
Modifié le ... : 18/10/2002
Description .. : fonctions et procédures non métier.
Mots clefs ... :
*****************************************************************}
unit wCommuns;

interface

uses
  Windows,
  Classes,
  SysUtils,
  HCtrls,
  Hent1,
  uTob,
  Dialogs,
  Forms,
  Controls,
  comctrls,
  stdctrls,
  extCtrls,
  ed_tools,
  Variants,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    Db,
  {$ENDIF}
  UTom,
  {$IFDEF GESCOM}
  uTomComm,
  {$ENDIF GESCOM}
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      {$IFNDEF EAGLCLIENT}
        MUL,
        Fiche,
        FichList,
        EdtEtat,
        EdtREtat,
        FE_Main,
        MenuOlg,
      {$ELSE  !EAGLCLIENT}
        eMul,
        eFiche,
        eFichList,
        utilEagl,
        mainEagl,
        MenuOLx,
      {$ENDIF !EAGLCLIENT}
      SaisieList,
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  HMsgBox,
  Hdb,
  M3FP,
  Graphics,
  Menus,
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      {$IFNDEF EAGLCLIENT}
        M3CODE,
      {$ENDIF !EAGLCLIENT}
      M3VM,
    {$ENDIF !ERADIO}
  {$ENDIF EAGLSERVER}
  wRapport,
  HPanel,
  hStatus,
  CBPTrace
  , CommonTools
  ;

type
  //GP_20071217_TS_GP14539 >>>
  MyArrayField = Array of String;
	MyArrayValue = Array of Variant;

  ExceptionW = class(Exception)
  public
    constructor Create(const Msg: string);
  end;

  { Class de structuration d'une liste paramétrable }
  //GP_20071217_TS_GP14539 <<<
  TWListeString = Set of (wlsObligatoire, wlsVisible);
  TWListe = class;
  TWFieldListe = class
  private
    FParent        : TWListe;
    FName, FCaption: String;
    FObligatoire,
    FVisible,
    FCumul,
    FLibelleComplet: Boolean;
    FAlignment: TAlignment;
    //GP_20071217_TS_GP14539
    FFormat, FSql: String;
    FWidth: Integer;
    FImage: Boolean;
    FEmptyIfNull: Boolean;
    FTypeField: Char;
    //GP_20071217_TS_GP14539 >>>
    FSqlFields,
    FSqlValues: MyArrayField;
    FSqlWhere: String;
    
    { Getter }
    function GetIndex: Integer;
    function GetTableName: String;
    procedure SetSql(const Value: String);
    function GetTokenPosFromTheEnd(const S, TokenToFind: String): Integer;
    procedure TrimChar(var St: String; const C1, C2: Char);
    function DecodeWhere(var Sql: String): Boolean;
    function FindInFormule(FieldName: HString): Variant;
    //GP_20071217_TS_GP14539 <<<
  public
    constructor Create(AListe: TWListe; const AName, ACaption: String; const AWidth: Integer; Params: String);

    //GP_20071217_TS_GP14539
    function isSql: Boolean;

    property Index: Integer read GetIndex;
    property Name: String read FName;
    property Caption: String read FCaption;
    property Alignment: TAlignment read FAlignment write FAlignment;
    property Format: String read FFormat write FFormat;
    property Width: Integer read FWidth write FWidth;
    property Image: Boolean read FImage write FImage;
    property Obligatoire: Boolean read FObligatoire;
    //GP_20071217_TS_GP14539
    property Visible: Boolean read FVisible write FVisible;
    property EmptyIfNull: Boolean read FEmptyIfNull;
    property Cumul: Boolean read FCumul;
    property LibelleComplet: Boolean read FLibelleComplet;
    property TableName: String read GetTableName;
    property TypeField: Char read FTypeField;
    //GP_20071217_TS_GP14539 >>>
    property Sql: String read FSql write SetSql;
    property SqlFields: MyArrayField read FSqlFields write FSqlFields;
    property SqlValues: MyArrayField read FSqlValues write FSqlValues;
    property SqlWhere: String read FSqlWhere;
    //GP_20071217_TS_GP14539 <<<
  end;

  TWListe = class
  private
    FDBListe: String;
    FChamps: TList;
    FTables: Array of String;
    FJoin: String;
    FOrderBy: String;
    FJoinByFrom: Boolean;
    //GP_20071217_TS_GP14539
    FIndexAlias: Integer;
//GP_20080409_TS_GP14539
    FGereSqlFields: Boolean;
    function GetChamps(Index: Integer): TWFieldListe;
    function GetChampsByName(ColName: String): TWFieldListe;
    function GetCount: Integer;
    function GetCountVisible: Integer;
    function GetIsChamps(Index: Integer): Boolean;
    function GetTable(Index: Integer): String;
  public
//GP_20080409_TS_GP14539 >>>
    constructor Create(const DBListe: String; Ecran: TObject; const OnlyFields: Boolean = False;
                       const GereSqlFields: Boolean = False); reintroduce;
//GP_20080409_TS_GP14539 <<<
    destructor Destroy; override;
    procedure Clear;
    function GetFields(const Format: TWListeString): String;
    function GetCaptions(const Format: TWListeString): String;
    function TablesCount: Integer;
    function SQL: String;

    procedure AddField(const ColName, ColTitle: String;
                       const ColWidth: Integer; const ColType: Char;
                       const ColFormat: String; const ColAlign: TAlignment;
                       const ColImage: Boolean = False; const ColCumul: Boolean = False;
                       const ColLibelleComplet: Boolean = False);
    property DBListe: String read FDBListe;
    property ChampsCount: Integer read GetCount;
    property ChampsVisibleCount: Integer read GetCountVisible;
    //GP_20071217_TS_GP14539
    property Champs[Index: Integer]: TWFieldListe read GetChamps; default;
    property IsChamps[Index: Integer]: Boolean read GetIsChamps;
    property ChampsByName[ColName: String]: TWFieldListe read GetChampsByName;
    property Tables[Index: Integer]: String read GetTable;
    property Join: String read FJoin;
    property OrderBy: String read FOrderBy;
    property JoinByFrom: Boolean read FJoinByFrom;
  end;

{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    TMemoFindDialog = class(TFindDialog)
    private
      FMemo: TMemo;
      FStringToFind: String;
      FEcran: TForm;
      FOnFound: TNotifyEvent;
    protected
      procedure Find; override;
    public
      procedure FindNext;

      property Memo: TMemo read FMemo write FMemo;
      property Ecran: TForm read FEcran write FEcran;
      property StringToFind: String read FStringToFind write FStringToFind;
      property OnFound: TNotifyEvent read FOnFound write FOnFound;
    end;
  {$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{$IFNDEF EAGLSERVER}
type
  TPrinterOptions = record
    OldPrinterIndex,
    OldNbDocCopies  : Integer;
    OldSilentMode,
    OldNoPrintDialog: Boolean;
  end;
{$ENDIF !EAGLSERVER}

type
	MyNewLine = (wNormal, wException, wExceptionLig, wInsertion, wDuplication);
	TFuncWRP = function(Q: TQuery): String;

  tArrondiQuoi = (tTPMTaux, tTPMPrix, tTPMMont);
	tArrondiPrec = (tAPCentaineDeMillier, tAPDizaineDeMillier, tAPMillier, tAPCentaine, tAPDizaine, tAPUnite, tAPDizieme, tAPCentieme, tAPMillieme, tAPDixMillieme);
  tArrondiMeth = (tAMInferieure, tAMSuperieure, tAMPlusProche);

	{ Manipulation dans DETABLES, DECHAMPS }
	function  wGetLibChamp(Const FieldName:String): String;
	function  wExistFieldInDechamps(Const FieldName:String): Boolean;
	function  wExistTableInDetables(Const TableName:String): Boolean;
	function  wGetNbrFieldInKey(Const TableName: String): Integer;
	function  wGetPrefixe(Const FieldName: String): String;
	function  wGetSuffixe(Const FieldName: String; const WithUnderScore: Boolean = False): String;
	procedure wGetTableFieldsName(Const TableName: String; TSCodeChamp: HTStrings; TSLibChamp: HTStrings = nil);
  function  wGetTypeField(Const FieldName: String): String;
  function  wGetSimpleTypeField(Const FieldName: String): Char; overload;
  function  wGetSimpleTypeField(const FieldName: String; const iTable, iField: Integer): Char; overload;
  function  wGetInitValue(Const FieldName: String): Variant; overload;
  function  wGetInitValue(Const TypeField: Char): Variant; overload;

  { Make }
  function wMakeRange(Const Prefixe, PrefixeValeurChamp:String; Const WithFieldName:Boolean; NombreDeChamp:Integer; Const DS: TDataSet ): String; overload;
  function wMakeRange(Const Prefixe, PrefixeValeurChamp:String; Const WithFieldName:Boolean; NombreDeChamp:Integer; Const FieldValue: MyArrayValue): String; overload;
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      { ********************************* TS ********************************** }
      {$IFNDEF EAGLCLIENT}
        function wMakeRestrictedWhereFromList(Const Prefixe: String; Const Liste: ThDbGrid; Const Q: tQuery; Const IsAValidRecord: TFuncWRP; Const WithRapport: Boolean = false; Const Rapport: TWRapport = nil): string;
      {$ELSE}
        function wMakeRestrictedWhereFromList(Const Prefixe: String; Const Liste: ThGrid;   Const Q: tQuery; Const IsAValidRecord: TFuncWRP; Const WithRapport: Boolean = false; Const Rapport: TWRapport = nil): string;
      {$ENDIF EAGLCLIENT}
      { *********************************************************************** }
      {$IFNDEF EAGLCLIENT}
        {
          Pour info : J'ai rajouté une clé en paramètre facultatif
          On en a besoin pour pouvoir utiliser la multisélection ds les tarifs
          car YTARIFS n'a pas de clé naturelle fixe
        }
        function wMakeWhereFromList(Const Prefixe: string; Const Liste: ThDbGrid; Const Q: tQuery; Const Cle: string=''): string;
      {$ELSE}
        function wMakeWhereFromList(Const Prefixe: string; Const Liste: ThGrid; Const Q: tQuery; Const Cle: string=''): string;
      {$ENDIF EAGLCLIENT}
    {$ENDIF !ERADIO}
  {$ENDIF EAGLSERVER}
  function wMakeWhereFromFSL(Const Prefixe: string; Const Liste: THGrid; Const DS: tDataSet; Const Cle: string=''): string;
	function wMakeFieldArray (Const TableName: String ) : MyArrayField; overload;
  function wMakeFieldArray (Const TableName: String; Const MaxChamp:Integer) : MyArrayField; overload;
  function wMakeFieldArrayValue (Const TableName: String; Const DS:TDataSet ) : MyArrayField;
  function wMakeFieldString(Const TableName, Separ: String): String;
  function wMakeWhereSQL(Const sClef1, sValuesClef1: String): String;
  function wFieldValueToString(Const FieldName :string; Const DS:TDataSet):String; overload;
  function wFieldValueToString(Const FieldName:String; Const FieldValue:Variant):String; overload;
	function wFieldValueToStringWithCote(Const FieldName: String; Const FieldValue: Variant): String;
	function wFabriqueWhere(Const FieldArray: array of String; Const FieldValue: Array of Variant) : String; overload;
  function wFabriqueWhere(Const Prefixe: string; Const DS: tDataSet) : String; overload;
  function wFabriqueRange(Const TableName: String; Const FieldsName: String; Const DS: TDataSet ): String;
  function wGetValueClef1(Const TableName: String; Const F: TForm; Lequel: string=''): String ; overload;
  function wGetValueClef1(Const T: Tob): String; overload;
  function wGetValueClef1(Const TableName: String; Const TobData: Tob): String; overload;
  function WMakeRangeFromFplus(Const Fplus: string ): string;
  function wGetFieldsClef1 (Const TableName: String): String ;
  function wConvertStringToField (const TableName, FieldName, Value: String): Variant ;
	function wExtractParametres (const S: String ) : String;
  function wMakeInFromMultiValCombo(Const Value: string): string;
  function GetOrderByFromDBListe(const ListName: String): String;
  function GetListeChampsFromDBListe(const ListName: String): String;
  function Etoile(Const TableName: Hstring): HString;

  { Divers }
  function wCtrlCaractere(Const sMessage, sAControler, sAutorise: string) : string;
  function wTTOperator2SQLOperator(const TTOperator: String): String;
  function wGetSQLCondition(const FieldName, TTOperator: String; Value: Variant): String;
  function wEvalCondition(Const Valeur1, Operateur, Valeur2: string; FieldName: String): Boolean;
  function wEvalDelphiString(sScript: String; Const Params: Array of Variant; var Error: Boolean): Variant; overload;
  function wEvalDelphiString(sScript: string; Const ParamsName: String; Const Params: Array of Variant; var Error: Boolean): Variant;overload;
  function wDezipString(Const St: String): String;
  function wZipString(Const St: String): String;
  procedure wCheckNullFromTob(Const T: Tob);
  procedure wCreateChampSupFromTable(Const T: Tob; const TableName: String);
  function wCopyItemFromListToList(const Item: String; FromList, ToList: TStrings; const ByItemName: Boolean;
                                   const ExcludeDoublon: Boolean = True;
                                   const DeleteItemOrigine: Boolean = False;
                                   const InsertAtPos: Integer = -1): Integer;
  function wExtractSectionFromMemo(const Section: String; var Memo: String; const ModifyMemo: Boolean): String;
  function isVarInt(const ValueV: Variant): Boolean;
  function GetFieldSize(FieldName : string) : integer;

  { Outil Tob }
  function Find1erTobOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): Tob;
  function Find1erRangOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): integer;
  function FindRangOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): integer;
  function FindOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): Tob;
  procedure TobRealToVirtual(TobData: Tob; var TobVirtual: Tob);

  { Sql }
  function wGetSQLFieldValue(Const FieldName, TableName, Where: string; Const OrderBy : string=''): variant;
  function wGetSQLFieldsValues(Const FieldsName: Array of string; Const TableName, Where: string; Const OrderBy : string=''): MyArrayValue;
  function wDeleteSql(const TableName: String; Const FieldsName: array of String; Const FieldsValue: array of variant): Integer;
  function InsertIntoTable(Const TableName, Select, From: string): integer;

  { COMMUN, CHOIXCOD, CHOIXEXT }
  function wExistCO(const sType, Code: String): Boolean;
  function wExistCC(const sType, Code: String): Boolean;
  function wGetTabletteFromCommun(Const sType: String; Const TheTob: TOB; Const OrderByAbrege: Boolean = False): Boolean;
	function wGetTabletteFromChoixCod(Const sType: String; Const TheTob: TOB): Boolean;
  function wGetTabletteFromChoixExt(Const sType: String; Const TheTob: TOB): Boolean;

  { Tablette }
  function wIsDataTypeArticle(Const TheDataType: String): Boolean;
  function wIsDataTypeCodeArticle(Const TheDataType: String): Boolean;
  function wIsDataTypeTiers(Const TheDataType: String): Boolean;
  function wIsFieldFromKey1(Const FieldName: String): Boolean;

  { Tests }
  function iif(Const Expression, TruePart, FalsePart: Boolean): Boolean; overload;
  function iif(Const Expression: Boolean; Const TruePart, FalsePart: Integer): Integer; Overload;
  function iif(Const Expression: Boolean; Const TruePart, FalsePart: Double): Double; overload;
  function iif(Const Expression: Boolean; Const TruePart, FalsePart: String): String; overload;
  function iif(Const Expression: Boolean; Const TruePart, FalsePart: char): char; overload;
  function iifV(Const Expression: Boolean; Const TruePart, FalsePart: Variant): Variant;

  function wBetween(Const LaValeur, PetiteValeur, GrandeValeur: Variant; Const EgalitePetiteValeur:boolean=true; Const EgaliteGrandeValeur: Boolean=true): Boolean;

  { String }
  function wRight(Const Chaine: string; Const Long: integer): string;
  function wLeft(Const Chaine: string; Const Long: integer): string;
  function wPadRight(Const Chaine:string; Const Long:integer; Const Remplissage:string = ' ') : string;
  function wPadLeft(Const Chaine:string; Const Long:integer; Const Remplissage:string = ' ') : string;
	function wExtractGuillemet(Const Chaine: String): String;
  function wStringRepeat(Const Chaine: string; Nb: integer): string;
  function wStringToArray(Chaine: string; Const Delimiteur: String): MyArrayField;
  //GP_20071217_TS_GP14539 >>>
  function wArrayToString(MyArray: MyArrayField; Const Delimiteur: String): String; overload;
  function wArrayToString(MyArray: MyArrayValue; Const Delimiteur: String): String; overload;
  //GP_20071217_TS_GP14539 <<<
  procedure wStringToTString(Const Chaine: string; Result: TStrings; const Sep: String = ';');
  function wTStringToString(TString: TStrings; const Sep: String = ';'): String;
  function AglUpperCase(Parms : array of variant; nb : integer): Variant;
  function AglLowerCase(Parms : array of variant; nb : integer): Variant;
  function wStringToAlphaNumString(const St: String; const CharsToIgnore: String = ''): String;
  function wFileSizeToStr(FileSize: Integer; const NbDecimal: Integer): String;
  function wIntToC2(i: Integer): String;

  { Date }
  function wDateTimeToStr(Const LaDate: tDateTime; Const cFormat:string = 'AAAA-S-T-MM-JJ'; Const lFormat: boolean = True): string;
  function AddAnnee(Const LaDate: tDateTime; Const iNbAn: integer; Const sJour: string = 'J'): tDateTime;
  function AddSemestre(Const LaDate: tDateTime; Const iNbSemestre: integer; Const sJour: string = 'J'): tDateTime;
  function AddTrimestre(Const LaDate: tDateTime; Const iNbTrimestre: integer; Const sJour: string = 'J'): tDateTime;
  function AddMois(Const LaDate: tDateTime; Const iNbMois: integer; Const sJour: string = 'J'): tDateTime;
  function AddSemaine(Const LaDate: tDateTime; Const NbSemaine: integer; Const Jour: char = 'J'): tDateTime;
  function AddJour(Const LaDate: tDateTime; Const iNbJour: integer; Const sJour: string = 'J'): tDateTime;
  function wYear(Const LaDate: tDateTime): word;
  function wMonth(Const LaDate: tDateTime): word;
  function wDay(Const LaDate: tDateTime): word;
  function wGetNumDay(Const LaDate: tDateTime): Integer;
	function wGetMonth(Const LeMois: Integer; Const WithMajuscule: Boolean = False): String;
	function wGetLastWeek(Const LaDate: tDateTime): Word;
  function wGetTrimestreFromMois(Const LeMois: Integer): Integer;
  function wGetTrimestre(Const LaDate: tDateTime): Integer;
  function wGet1erMoisFromTrimestre(Const LeTrimestre: Integer): Integer;
  function wGetSemestreFromMois(Const LeMois: Integer): Integer;
  function wGet1erMoisFromSemestre(Const LeSemestre: Integer): Integer;
  function wGetSemestre(Const LaDate: tDateTime): Integer;
  function wHour(Const Heure: tTime): word;
  function wMin(Const Heure: tTime): word;
  function wSec(Const Heure: tTime): word;
  function wMsec(Const Heure: tTime): word;
  function IsDateVide(const sDate: String): Boolean;

  { Double }
  function MyDoubleToStr(MyDouble: Double; Const Remove: Boolean = false): string;

  { Conversion de temps }
	function HHCC(Const Temps: double): Double;
	function HHMM(const Temps: double): Double;

	{ Division }
	function wDivise(const Dividande, Diviseur: double; Const WithMessage: boolean = False) : double;

  { Delete }
  function wDeleteTable(Const TableName, Where: string; Const WithProgressForm: Boolean = False; const Rapport: TWRapport = nil): Boolean;
  function wDeleteTable2(Const TableName, Sql: string; Const WithProgressForm: Boolean = False; const Rapport: TWRapport = nil): Boolean;

  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      { Liste }
      {$IFNDEF EAGLCLIENT}
        function wMultiSelected(Const L : THDBGrid): boolean; overload;
        function wMultiSelected(Const L : THGrid): boolean; overload;
      {$ELSE}
        function wMultiSelected(Const L : THGrid): boolean;
      {$ENDIF}
      procedure wChangeDBListe(Const Mul: TFMul; const DBListe: String);
      procedure wHideListColumns(Const Ecran: TForm; Const ColumnsName: Array of string);
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}

  { Tob }
  function wSelectTobFromSQL(Const Sql: String; Const t: Tob; Const WithMemo: Boolean = True): Boolean;
  function wLoadTobFromListe(const DBListe: String; T: Tob; const OnlyExistingFields: Boolean = True): Boolean;
  procedure wMakeTobFromDS(Const DS: TDataSet; Const TheTob: TOB; Const OnRecords: Boolean = False; Const ToutEnAddChampSup: Boolean = False);
  procedure wSetDSFromTob(Const TheTob: TOB; Const DS: TDataSet);
	procedure wCopyTobBySuffixe(Const TobSource, TobDest: Tob; Const PrefixeSource, PrefixeDest: String; Const WithKey: Boolean = True; Const AddToNomChamp: String = '');
  function wTobCountDetails(T: Tob): Integer;
  function VarAsTob(T: Variant): Tob;
  function GetArgumentByTob(TobData: Tob; WriteKey: Boolean = True; WriteBlocNote: Boolean = False; SansPrefixe: Boolean = False): String;
  function wUpdateSingleTob(T: Tob): Boolean;
  function wInsertSingleTob(T: Tob): Boolean;
  function wDeleteSingleTob(T: Tob): Boolean;
  function wInsertOrUpdateSingleTob(T: Tob): Boolean;
  function wTobGetAbsoluteIndex(T: Tob): Integer;
  function TobExchange(TheTob: Tob; Const index1, index2: integer): boolean;
  procedure LoadTobFromArgument(const TobData: Tob; Argument: String; const Separateur: String = ';');
  //GP_20071217_TS_GP14539
  function TobToArrayValue(const TobData: Tob; const FieldNames: MyArrayField): MyArrayValue;

  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
    { Propriétés }
    procedure wCallProperties(Const Prefixe, Identifiant, Clef: String; Const PointeurFiche: String = '');
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}

  { Affichage, couleurs }
  function HSLtoRGB(H, S, L: Double): TColor;
  procedure RGBtoHSL(RGB: TColor; var H, S, L : Double);
	function wAssombrirCouleur(CouleurDOrigine: TColor; Pourcent: Integer): tColor;
	function wEclaircirCouleur(CouleurDOrigine: TColor; Pourcent: Integer): tColor;
  function wInvertColor(Couleur: TColor): tColor;
  procedure DrawTitlePanel(const P: THPanel);

  { Token }
  function wGetTokenPos(St: String; const SubSt: String; const ExactSubSt: Boolean = True): Integer;
  function wGetTokenAt(St: String; const Index: Integer): String;
  function wCountToken(St: String; const Separateur: String = ';'): Integer;
  function wMakeJoinOnCle1(const Table1, Table2: String): String;
  function wMakeSelectPresqueEtoile(const TableName: String; const FieldsToExclude: MyArrayField = nil; const ExcludeBlobField: Boolean = True): String;
  function wStringToAction(Const Action: String): TActionFiche;
  function GetMaxLenToken(St: String): String;

  procedure uAfficheGBUNIC (Const GbUnic: tGroupBox; Const Affiche:boolean); // Affichage du GroupBox GBUNIC
  function GetTokenFieldName(const Token: String): String;
  
  { Droits }
  function JAiLeDroitGestion(Const Prefixe: string): boolean;
  function JAiLeDroitConsult(Const Prefixe: string): boolean;

  { Debug }
  procedure PutToDebugLog(Const Fonction: string; Const Debut: boolean; Const Notes: string='');
  procedure wShowMeTheTob(Const T: Tob);
  {$IFNDEF EAGLCLIENT}
    Procedure www(T: Tob); overload;
    Procedure www(DS: TDataSet); overload;
  {$ELSE  !EAGLCLIENT}
    Procedure www(T: Tob); overload;
  {$ENDIF !EAGLCLIENT}
  Procedure www(const Sql: String); overload;

  { Listes }
  function wGetColsVisibilityFromListe(const ListeName: String; const OnlyFields: Boolean = False): String;
  function GetSelectColsFromListe(const ListeName: String; Const Prefixes: String=''): String;

  { Imprimante }
  procedure BeginPrint(const NbCopies: Integer = 1; const SilentMode: Boolean = False);
  procedure EndPrint;
  Procedure SelectIndexPrinter(const PrinterName: hString);

  { ProgressForm }
  procedure wInitProgressForm(Const Form: tWinControl; Const Caption, Title: string; Const MaxValue: integer; Const WithCancelButton, DisabledButtons: boolean);
  function wMoveProgressForm(Const Title: string = ''): boolean;
  procedure wFiniProgressForm;

  function EgaliteDouble(Const n1, n2: double): boolean;
  function wArrondir(Const nValeur: double; const ArrondiPrec: tArrondiPrec; Const ArrondiMeth: tArrondiMeth): double;
  function ArrondiPrecToNbDec(Const ArrondiPrec: tArrondiPrec): integer;
  function ArrondiPrecToPoids(Const ArrondiPrec: tArrondiPrec): double;
  function NbDecToArrondiPrec(Const iNbDec: integer): tArrondiPrec;
  function PoidsToArrondiPrec(Const nPoids: double): tArrondiPrec;
  function MethodeToArrondiMeth(Const sMethode: string; lMessageErreur: boolean = True): tArrondiMeth;
  function QuoiToArrondiQuoi(const sQuoi: string; lMessageErreur: boolean = True): tArrondiQuoi;

  { Renvoie la string d'une occurence dans un argument }
  function GetTokenByNum(Chaine: string; Num: integer): string;

  { Transforme une liste de valeurs issue d'une MultiComboBox en liste conforme à un IN pour WHERE }
  function GetListeINFromListeCombo(sListeDeValeurCombo: string): string;
  function TStringsToSqlIn(TS: TStrings): String;
  function TobToSqlIn(aTob: Tob; FieldName: String): String;

  procedure wLoadPopUpFromCommun(Const PopUp: TPopupMenu; Const CoType: String; Const AOnClick: TNotifyEvent; Const Plus: String = '');

  { Mot de passe du jour }
  function IsDayPass: boolean;

  { Contrôles }
  procedure GetCaptionByFocusControl(Container: TWinControl; T: Tob);

  { Pour compilation en process serveur > Voir à débloquer au niveau de l'AGL }
  {$IFDEF EAGLSERVER}
    function wActionToString(const Action: TActionFiche): String;
  {$ENDIF EAGLSERVER}
  function wTrue: char;
  function wFalse: char;
  Type tTemps = record
                  Fonction: string;
                  Temps   : tDateTime;
                end;
//GP_20080204_TS_GP14791 >>>
  { Messages communs }
  function TransactionErrorMessage: String;

  { Utilisation des Plugins }
  {$IF Defined(EAGLCLIENT) and not Defined(EAGLSERVER)}
    type
      twPluginParam = record
                        varName : String;
                        varType : TCSType;
                        varValue: Variant;
                      end;
      twPluginParams = Array of twPluginParam;

    { - Générique }
    function wExecutePluginVerb(const ClassName, FunctionName, VerbParam: String; FunctionParams: twPluginParams;
                                const ControlDllFonctionName: String = ''): Tob;
  {$IFEND EAGLCLIENT and not(EAGLSERVER)}
  { - Verbes communs }
  function wExecuteSqlOnClonedDB(const Sql: String): Integer;
//GP_20080204_TS_GP14791 <<<

  function GetWindowsCaptionFromArgument(Argument, DefaultValue : string) : string;
  function GetFieldType(FieldName : string) : tTypeField;
const
  SELECT_FROM_LISTE = 'SELECT * FROM LISTE WHERE LI_LISTE="%s" AND LI_UTILISATEUR="%s" AND LI_LANGUE="%s" ORDER BY LI_UTILISATEUR DESC';

{$IFNDEF EAGLSERVER}
var
  RelativeTime            : Array of tTemps;
  ProgressFormImbrication : Integer;
  CurrentForm             : THForm;
  PrinterOptions          : TPrinterOptions;
{$ENDIF EAGLSERVER}

implementation

{ Attention, pas de uses métiers! }
uses
  Math,
  LicUtil,
  {$IFDEF GPAO}
    {$IFNDEF EAGLSERVER}
      {$IFNDEF ERADIO}
        uTobDebug,
      {$ENDIF !ERADIO}
    {$ENDIF EAGLSERVER}
  {$ENDIF GPAO}
  {$IFDEF EAGLSERVER}
    eSession,
  {$ELSE  EAGLSERVER}
    HDrawXP,
  {$ENDIF EAGLSERVER}
//GP_20080204_TS_GP14791 >>>
  {$IFDEF EAGLCLIENT}
    uHttp,
  {$ENDIF EAGLCLIENT}
//GP_20080204_TS_GP14791 <<<
  HZStream,
  Printers,
  Messages,
  UtilPGI,
  {$IFDEF GCGC}
//  yTarifsCommun,
  {$ENDIF GCGC}
  //GP_20071217_TS_GP14539
  Formule,
  CbpMCD,
  CbpEnumerator,
  DateUtils
  ;
{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 10/09/2001
Description .. : Cette fonction permet d'extraire une zone délimitée par [...].
Suite ........ : C'est la méthode utilisée pour transmettre des paramètres
Suite ........ : via AglLanceFiche/OuvreFiche.
*****************************************************************}
function wExtractParametres ( const S: String ) : String;
var
  i1,i2: integer;
begin
  Result := '';
  i1 := pos ( '[', S );
  if i1 > 0 then
  begin
  	i2 := pos ( ']', S );
   	if i2 > 0 then
    	Result := Copy ( S, I1+1, I2-I1-1 );
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/03/2005
Modifié le ... :   /  /
Description .. : Renvoi un IN sql fabriqué à partir d'une valeur de
Suite ........ : MultiValComboBox
Mots clefs ... :
*****************************************************************}
function wMakeInFromMultiValCombo(Const Value: string): string;
begin
  Result := '';

  if Value = '' then exit;
  if Pos('<<', Value) > 0 then exit;

  Result := '("' + Value + '")';

  Result := StringReplace(Result, '";', '"'  , [rfReplaceAll]);
  Result := StringReplace(Result, ';"', '"'  , [rfReplaceAll]);
  Result := StringReplace(Result, ';' , '","', [rfReplaceAll]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 22/06/2007
Modifié le ... :   /  /    
Description .. : Permet de récupérer l'ORDER BY d'une DBListe
Mots clefs ... : 
*****************************************************************}
function GetOrderByFromDBListe(const ListName: String): String;
var
  NA1, NA4: String;
  NA2, NA3, sColCaption: HString;
  NA5, NA6: Boolean;
  sTables, sJoin, sOrderBy,
  sColName, sColWidth, sColParams: String;
begin
  if ChargeHListeUser(ListName, V_Pgi.User, sTables, sJoin, sOrderBy, sColName, sColCaption,
                      sColWidth, sColParams, NA1, NA2, NA3, NA4, NA5, NA6) then
    Result := sOrderBy
  else
    Result := ''
end;

{***********A.G.L.***********************************************
Auteur  ...... : Garnier Marie-Nöelle
Créé le ...... : 11/12/2007
Modifié le ... :   /  /
Description .. : Permet de récupérer la liste des champs d'une DBListe
Mots clefs ... :
*****************************************************************}
function GetListeChampsFromDBListe(const ListName: String): String;
var
  wListe: TWListe;
  i : integer;
begin
  Result := '';
  wListe := TWListe.Create(ListName, nil, true);
  try
    if (wListe.ChampsCount > 0) then
      begin
      for i := 0 to Pred(wListe.ChampsCount) do
        begin
        if  wListe.Champs[i].Visible then
          Result := Result + iif(Result <> '', ';', '') + wListe.Champs[i].Name;
        end
      end;
  finally
    wListe.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 29/06/2007
Modifié le ... :   /  /    
Description .. : Retourne la liste des champs de TableName pour faire un 
Suite ........ : Select *
Mots clefs ... : 
*****************************************************************}
function Etoile(Const TableName: Hstring): HString;
begin
  Result := wMakeSelectPresqueEtoile(TableName);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 05/10/2001
Description .. : Récupère le préfixe dans un nom de champ
*****************************************************************}
function wGetPrefixe(Const FieldName: String): String;
var
  i: Integer;
begin
  i := Pos ( '_', FieldName );
  if i > 0 then
  	Result := Copy(FieldName, 1, i - 1)
  else
  	Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 05/10/2001
Description .. : Récupère le suffixe d'un nom de champ
Description .. : (Avec ou sans UnderScore)
*****************************************************************}
function wGetSuffixe(Const FieldName: String; Const WithUnderScore: Boolean = False): String;
var
	i:Integer;
begin
	i := Pos('_', FieldName);
	if (i <> 0) then
  begin
    if WithUnderScore then 
      Result := Copy(FieldName, i, Length(FieldName) - i + 1)
    else
      Result := Copy(FieldName, i + 1, Length(FieldName) - i)
  end
  else
    Result := FieldName;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier Persouyre
Créé le ...... : 20/09/2001
Description .. : Fabrique un range du type 'VALEUR1,VALEUR2' à utiliser
Description .. : pour 'ranger' les fiches simples
Modifié le ... : 03/04/2002 par Thierry
Description .. : Gestion nouvel ordre de tri de DECHAMPS
Mots clefs ... :
*****************************************************************}
function wFabriqueRange(Const TableName: String; Const FieldsName: String; Const DS: TDataSet): String;
var
  s: String;
  Champ: String ;
  Value: String ;
  iTable, iChamp: Integer;
  TypeField: String ;
  Ok: Boolean;

  Mcd : IMCDServiceCOM;
  McdTable  : ITableCOM ;
  Field     : IFieldCOM ;

begin
  if not ((DS.Eof) and (DS.Bof)) then
  begin
    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();
    //
    { Valeur par défaut : veux dire que erreur }
    Result := '' ;
    McdTable := MCD.getTable(TableName);
    if Assigned(McdTable) then
    begin
      { Ecrit la chaine de résultat }
      s := FieldsName; Ok := True;
      while (s <> '') and Ok do
      begin
        Champ := ReadTokenSt(s); Champ := Trim(Champ);
        if (Champ <> '') then
        begin
          Field := Mcd.GetField(Champ);
//          iChamp := ChampToNum(Champ);
          if Assigned(Field) then
          begin
            Value := '' ;
            TypeField := Field.tipe ;
            if ( TypeField = 'INTEGER' ) OR ( TypeField = 'SMALLINT' ) then
              Value := IntToStr ( DS.FindField ( Champ ).AsInteger )
            else if ( TypeField = 'DOUBLE' ) OR ( TypeField = 'RATE' ) OR ( TypeField = 'EXTENDED' ) then
              Value := FloatToStr ( DS.FindField ( Champ ).AsFloat )
            else if ( TypeField = 'DATE' ) then
              Value := DateTimeToStr ( DS.FindField ( Champ ).AsDateTime )
            else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then
            else Value := DS.FindField ( Champ ).AsString ;
            if Result = '' then Result := Value else Result := Result + ';' + Value ;
          end;
        end
        else
          Ok := False;
      end;
      if wRight(Result, 1) = ';' then
        Result := Result + ';';
      if not Ok then Result := '';
    end;
  end
  else
    Result := ''
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 05/10/2001
Modifié le ... :   /  /
Description .. : Fabrique un range du type 'WGT_CHAMP1=VALEUR1,WGT_CHAMP2=VALEUR2'
Suite ........ : à utiliser par exemple dans 'OuvreFiche' pour initialiser 'Range' et 'Lequel'.
Mots clefs ... :
Exemple ...... :

  WMakeRange('WGT', 'WGT', True, -1, DS)
     => WGT_NATURETRAVAIL=[DS.FieldValue('WGT_NATURETRAVAIL')],
     => WGT_CODEARTICLE=[DS.FieldValue('WGT_CODEARTICLE')],
     => WGT_MAJEUR=[DS.FieldValue('WGT_MAJEUR')]    Permet de faire un Range

  WMakeRange('WGT', 'WGT', False, -1, DS)
  	  => [DS.FieldValue('WGT_NATURETRAVAIL')],
     => [DS.FieldValue('WGT_CODEARTICLE')],
     => [DS.FieldValue('WGT_MAJEUR')]                Permet de faire un Lequel

  WMakeRange('WGL', 'WGT', True, 3, DS)
     Récupération du nom des 3 premiers champs de la clé1 de la table WGAMMELIG
     et récupération dans le DS de la valeur des champs WGT_
     => WGL_NATURETRAVAIL=[DS.FieldValue('WGT_NATURETRAVAIL')],
     => WGL_CODEARTCILE=[DS.FieldValue('WGT_CODEARTICLE')],
     => WGL_MAJEUR=[DS.FieldValue('WGT_MAJEUR')]

*****************************************************************}
function WMakeRange(Const Prefixe, PrefixeValeurChamp:String; Const WithFieldName:Boolean; NombreDeChamp:Integer; Const DS: TDataSet ): String; Overload;
var
  TableFieldName, TableFieldValue:String;
  ArrayFieldName, ArrayFieldValue:MyArrayField;
  i:Integer;
  Separ:String;
begin
  SetLength(ArrayFieldName, 0);
  SetLength(ArrayFieldValue, 0);
  if not DS.Eof then
  begin
    { Récupération du nom de la table }
    TableFieldName := PrefixeToTable ( Prefixe );
    TableFieldValue := PrefixeToTable ( PrefixeValeurChamp );
    if NombredeChamp < 1 then NombreDeChamp := WGetNbrFieldInKey(TableFieldValue);
    { Forme le tableau des noms des champs }
    ArrayFieldName := WMakeFieldArray( TableFieldName, NombreDeChamp);
    { Forme le tableau des valeurs des champs }
    ArrayFieldValue := WMakeFieldArrayValue( TableFieldValue, DS);
    { Chaine de résultat }
    Separ := ';';
    Result := '';
    for i := 0 to NombreDeChamp-1 do
    begin
      if Result = '' then
        Result := iif(WithFieldName, ArrayFieldName[i] + '=', '') + ArrayFieldValue[i]
      else
        Result := Result + Separ + iif(WithFieldName, ArrayFieldName[i] + '=', '') + ArrayFieldValue[i];
    end;
  end
  else
    Result := ''
end;

function WMakeRange(Const Prefixe, PrefixeValeurChamp:String; Const WithFieldName:Boolean; NombreDeChamp:Integer; Const FieldValue: MyArrayValue): String; overload;
{ Fabrique un range en utilisant les valeurs des champs dans un tableau }
var
  TableFieldName, TableFieldValue:String;
  ArrayFieldName:MyArrayField;
  i:Integer;
  Separ:String;
begin
  { Récupération du nom de la table }
  TableFieldName := PrefixeToTable ( Prefixe );
  TableFieldValue := PrefixeToTable ( PrefixeValeurChamp );
  if NombredeChamp < 1 then NombreDeChamp := WGetNbrFieldInKey(TableFieldValue);
  { Forme le tableau des noms des champs }
	ArrayFieldName := WMakeFieldArray( TableFieldName, NombreDeChamp);
  { Chaine de résultat }
  Separ := ';'; Result := '';
  for i := 0 to NombreDeChamp-1 do
  begin
    if Result = '' then
      Result := iif(WithFieldName, ArrayFieldName[i] + '=', '') + WFieldValueToString(ArrayFieldName[i], FieldValue[i])
    else
      Result := Result + Separ + iif(WithFieldName, ArrayFieldName[i] + '=', '') + WFieldValueToString(ArrayFieldName[i], FieldValue[i]);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/04/2002
Description .. : A partir d'un tableau contenant des noms de champ et d'un
Suite ........ : tableau contenat des valeurs, forme un Where réutilisable
Suite ........ : dans une requete SQL
Modifié le ... : 11/04/2002 par Thierry
Modifié le ... : Gestion nouvel ordre de tri de DECHAMPS
Mots clefs ... : WHERE;SQL
*****************************************************************}
function wFabriqueWhere(Const FieldArray: array of String; Const FieldValue: Array of Variant) : String;
var
  I			        : Integer;
  Value		      : String;
  Prefixe 	    : String;
  Table		      : String;
  iTable, iChamp: Integer;
  TypeField     : String;
  Ok            : Boolean;
  Mcd : IMCDServiceCOM;
  Field     : IFieldCOM ;
begin
	Result := '';
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();

  { Forme le where en collant les nom de champ et leur valeur }
  i := Low(FieldArray) - 1; Ok := True;
  while (i < High(FieldArray)) and Ok do
  begin
    Inc(i);
    Field := MCD.getField(FieldArray[i]);
    if Assigned(Field) then
    begin
      TypeField := Field.tipe;
      if ( TypeField = 'INTEGER' ) OR ( TypeField = 'SMALLINT' ) then
      begin
        if isVarInt(FieldValue[i]) then
          Value := IntToStr ( FieldValue[i] )
        else
          Value := FieldValue[i];
      end
      else
        if ( TypeField = 'DOUBLE' ) OR ( TypeField = 'RATE' ) OR ( TypeField = 'EXTENDED' ) then
        begin
          if VarType ( FieldValue[i] ) <> VarNull then
            Value := FloatToStr ( FieldValue[i] )
          else
            Value := '';
        end
        else if ( TypeField = 'DATE' ) then
          Value := '"' + DateTimeToStr ( FieldValue[i] ) + '"'
        else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then
        else Value := '"' + FieldValue[i] + '"';

        if Result = '' then
          Result := FieldArray[i] + '=' + Value
        else
          Result := Result + ' AND ' + FieldArray[i] + '=' + Value;
    end
    else
      Ok := False;
  end;
  if not Ok then Result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/05/2002
Modifié le ... :   /  /
Description .. : Fabrique un where à partir du Prefixe et du DataSet
Mots clefs ... :
*****************************************************************}
function wFabriqueWhere (Const Prefixe: string; Const DS: tDataSet) : String; overload;
var
	aField : MyArrayField;
  aValue : MyArrayValue;

  function MakeVariant(a: MyArrayField): MyArrayValue;
  var
  	i : integer;
  begin
  	SetLength(Result, High(a) + 1);
    for i := 0 to High(a) do
     	Result[i] := a[i];
  end;
begin
	aField := wMakeFieldArray(PrefixeToTable(Prefixe));
  aValue := MakeVariant(wMakeFieldArrayValue(PrefixeToTable(Prefixe), DS));
	Result := wFabriqueWhere(aField, aValue);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 11/09/2001
Modifié le ... :   /  /
Description .. : Cette fonction supprime les guillements en première et
Suite ........ : dernière position de la string passée en paramètre.
Mots clefs ... :
*****************************************************************}
function WExtractGuillemet(Const Chaine: String ): String;
begin
	Result := Chaine;
  if Length(Result) > 0 then
  begin
    if Result[1] = '"' then
      Delete(Result, 1, 1);
    if Result[Length(Result)] = '"' then
      Delete(Result, Length(Result), 1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 05/10/2001
Modifié le ... :   /  /
Description .. : A partir d'un nom de table, fabrique un tableau de string
Suite ........ : contenant les champs de la cle1 de la table;
Suite ........ : MaxChamp permet de limiter le tableau au n premier champs
Suite ........ : de le clé.
Suite ........ : Si Maxchamp = -1 on prend tous les champs de la clé
Mots clefs ... :
*****************************************************************}
function WMakeFieldArray (Const TableName: String; Const MaxChamp:Integer) : MyArrayField; overload;
var
  iTable: Integer;
  Indice: Integer;
  S,St: String;
  Mcd : IMCDServiceCOM;
  McdTable  : ITableCOM ;
  Keys : IStringEnumerator;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  McdTable := MCD.getTable(TableName);
  { Récupération des champs de la clé principale }
  Keys := McdTable.Keys;
  Keys.reset;
  Keys.movenext;
  S := keys.Current;
  //
  St := trim(ReadTokenPipe( S, ','));
  Indice := 0;
  while St <> '' do
  begin
    Inc(Indice);
    SetLength(Result, Indice);
    Result[Indice-1] := St;
    if (MaxChamp <> -1) and (Indice = MaxChamp) then BREAK;
    St := trim(ReadTokenPipe (S, ','));
  end;
end;

function WMakeFieldArray ( const TableName: String) : MyArrayField; overload;
begin
  Result := WMakeFieldArray( TableName, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Permet d'évaluer complètemnt une expression sans passer
Suite ........ : par les IF, ELSE IF...
Mots clefs ... :
*****************************************************************}
function iif(Const Expression, TruePart, FalsePart: Boolean): Boolean; overload;
begin
  Result := Tools.iif(Expression, TruePart, FalsePart);
end;

function iif(Const Expression: Boolean; Const TruePart, FalsePart: Integer): Integer; overload;
begin
  Result := Tools.iif(Expression, TruePart, FalsePart);
end;

function iif(Const Expression: Boolean; Const TruePart, FalsePart: Double): Double; overload;
begin
  Result := Tools.iif(Expression, TruePart, FalsePart);
end;

function iif(Const Expression: Boolean; Const TruePart, FalsePart: String): String; overload;
begin
  Result := Tools.iif(Expression, TruePart, FalsePart);
end;

function iif(Const Expression: Boolean; Const TruePart, FalsePart: Char): Char; overload;
begin
  Result := Tools.iif(Expression, TruePart, FalsePart);
end;

function iifV(Const Expression: Boolean; Const TruePart, FalsePart: Variant): Variant;
begin
	if Expression then
		Result := TruePart
	else
		Result := FalsePart;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : A partir d'un N° de table, d'un N° de champ dans
Suite ........ : V_PGI.DECHAMPS, retourne la valeur du champ contenu
Suite ........ : dans le Dataset et la convertie en string
Mots clefs ... : DATASET,DECHAMPS
*****************************************************************}
function wFieldValueToString(Const FieldName :string; Const DS:TDataSet):String; overload;
var
	TypeField:String;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  TypeField := (Mcd.GetField(FieldName) as IFieldCom).Tipe;

  if ( TypeField = 'INTEGER' ) OR ( TypeField = 'SMALLINT' ) then
    Result := IntToStr ( DS.FindField ( FieldName ).AsInteger )
  else if ( TypeField = 'DOUBLE' ) OR ( TypeField = 'RATE' ) OR ( TypeField = 'EXTENDED' ) then
    Result := FloatToStr ( DS.FindField (FieldName).AsFloat )
  else if ( TypeField = 'DATE' ) then
    Result := '"' + DateTimeToStr ( DS.FindField ( FieldName ).AsDateTime ) + '"'
  else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then
  else
    Result := DS.FindField ( FieldName ).AsString;
end;

function WFieldValueToString(Const FieldName:String; Const FieldValue:Variant):String; overload;
var
	TypeField:String;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  TypeField := (Mcd.GetField(FieldName) as IFieldCom).Tipe;
  if ( TypeField = 'INTEGER' ) OR ( TypeField = 'SMALLINT' ) then
    Result := IntToStr ( FieldValue )
  else if ( TypeField = 'DOUBLE' ) OR ( TypeField = 'RATE' ) OR ( TypeField = 'EXTENDED' ) then
    Result := FloatToStr ( FieldValue )
  else if ( TypeField = 'DATE' ) then
    Result := '"' + DateTimeToStr ( FieldValue ) + '"'
  else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then
  else
    Result := FieldValue;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 27/03/2002
Modifié le ... :   /  /
Description .. : A partir d'un nom de champ et d'une valeur retourne la
Suite ........ : valeur en string encadrée par des " si besoin.
Suite ........ : Utile pour les requetes de mise à jour
Mots clefs ... :
*****************************************************************}
function WFieldValueToStringWithCote(Const FieldName: String; Const FieldValue: Variant): String;
var
	TypeField:String;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  TypeField := (Mcd.GetField(FieldName) as IFieldCom).Tipe;
  if ( TypeField = 'INTEGER' ) OR ( TypeField = 'SMALLINT' ) then
  begin
    if isVarInt(FieldValue) then Result := IntToStr ( FieldValue );
  end
  else
    if ( TypeField = 'DOUBLE' ) OR ( TypeField = 'RATE' ) OR ( TypeField = 'EXTENDED' ) then
    begin
      if VarType ( FieldValue ) <> VarNull then Result := StrFpoint ( FieldValue );
    end
    else if ( TypeField = 'DATE' ) then Result := '"' + UsDateTime(( FieldValue )) + '"'
    else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then Result := ''
    else Result := '"' + FieldValue + '"';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : A partir d'un nom de table, retourne un tableau contenant
Suite ........ : la valeur des différents champs composant la clé. Ces
Suite ........ : valeurs sont recherchées dans le Dataset passé en
Suite ........ : paramètres
Mots clefs ... : CLE,DATASET
*****************************************************************}
function WMakeFieldArrayValue (Const TableName: String; Const DS:TDataSet ) : MyArrayField;
var
  iTable,iChamp: Integer;
  Indice: Integer;
  S,St: String;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  //
  if not DS.Eof then
  begin
    S := Mcd.TableToCle1(TableName);
    St := ReadTokenPipe ( S, ',' );
    Indice := 0;
    while St <> '' do
    begin
      Inc ( Indice );
      SetLength ( Result, Indice );
			Result[Indice-1] := WFieldValueToString (st,DS);
      St := ReadTokenPipe ( S, ',' );
    end;
  end
  else
    SetLength(Result, 0) 
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 04/10/2001
Modifié le ... :   /  /
Description .. : Pour une table, fabrique une chaine contenant les champs
Suite ........ : de la clé primaire
Suite ........ : Exemple : ('WGAMMETET', ';') retourne la chaine :
Suite ........ : WGT_NATURETRAVAIL;WGT_CODEARTICLE;WGT_MAJEUR
Mots clefs ... :
*****************************************************************}
Function WMakeFieldString ( const TableName, Separ: String  ) : String;
var
  i : Integer;
  SavePos : MyArrayField;
begin
  Result := '';
  SavePos := WMakeFieldArray ( TableName );
  for i := Low(SavePos) to High(SavePos) do
  begin
    if i = 0 then
      Result := SavePos[i]
    else
      Result := Result + Separ + SavePos[i];
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/06/2003
Modifié le ... :   /  /
Description .. : Permet de composer le where SQL avec les données de la
Suite ........ : clef.
Suite ........ : Paramètres :
Suite ........ : - sClef1 : 'CHAMP1;...;CHAMPN'
Suite ........ : - sValuesClef1 : 'Valeur_CHAMP1;...;ValeurCHAMPN'
Suite ........ : S'utilise avec les fonction :
Suite ........ : - wGetValueClef1(GetTableName, Ecran)
Suite ........ : - wMakeFieldString(GetTableName, ';');
Mots clefs ... :
*****************************************************************}
function wMakeWhereSQL(Const sClef1, sValuesClef1: String): String;
var
  sID, sClef, sField: String;

  function GetStr(cType: Char): String;
  begin
    Result := ReadTokenSt(sID);
    if Result = '' then
    begin
      case cType of
        'B'     : Result := wFalse;
        'N', 'I': Result := '0';
        'D'     : Result := USDateTime(iDate1900);
      end
    end
    else if cType = 'D' then
      Result := UsDateTime(StrToDateTime(Result));
  end;

begin
  sID := sValuesClef1;
  sClef := sClef1;
  Result := '';

  while sClef <> '' do
  begin
    sField := ReadTokenSt(sClef);
    case wGetSimpleTypeField(sField) of
      'C', 'B', 'D': Result := Result + iif(Result = '', '', ' AND ') + sField + '="' + GetStr(wGetSimpleTypeField(sField)) + '"';
      'N', 'I'     : Result := Result + iif(Result = '', '', ' AND ') + sField + '='  + GetStr(wGetSimpleTypeField(sField));
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/11/2001
Modifié le ... :   /  /
Description .. : Renvoie la valeur d'un champ avec un sql
Mots clefs ... :
*****************************************************************}
function wGetSQLFieldValue(Const FieldName, TableName, Where: string; Const OrderBy : string=''): variant;
var
	Sql: string;
  T  : Tob;
begin
	Result := null;
	Sql := 'SELECT ' + FieldName + ' AS VAL'
       + ' FROM ' + TableName
       + iif(Where   <> '', ' WHERE '    + Where  , '')
       + iif(OrderBy <> '', ' ORDER BY ' + OrderBy, '')
       ;

  T := Tob.Create ('', nil, -1);
  try
    T.LoadDetailDBFromSQL ('', Sql);
    if T.Detail.Count > 0 then
    begin
      if (Pos('MIN(', UpperCase(FieldName)) > 0) or (Pos('MAX(', UpperCase(FieldName)) > 0) or (Pos('SUM(', UpperCase(FieldName)) > 0) then
      begin
        if T.Detail[0].GetString('VAL') = '' then
          Result := 0.0
        else
          Result := T.Detail[0].GetDouble('VAL');
      end
      else
      begin
        Result := T.Detail[0].GetValue('VAL');
        if VarIsNull(Result) then
          Result := wGetInitValue(FieldName)
      end;
    end
    else
    begin
      Result := wGetInitValue(FieldName)
    end;
  finally
    T.free;
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/11/2001
Modifié le ... :   /  /
Description .. : Retourne des valeurs en fonction d'un where
Mots clefs ... :
*****************************************************************}
function wGetSQLFieldsValues(Const FieldsName: Array of string; Const TableName, Where: string; Const OrderBy : string=''): MyArrayValue;
var
  i					    : integer;
	Sql, TypeField: String;
	Q					    : TQuery;
begin
  Result := nil;

  Sql := 'SELECT ';
	for i := Low(FieldsName) to High(FieldsName) do
  begin
   	Sql:= Sql + iif(i = 0, FieldsName[i], ', ' + FieldsName[i]);
	end;

  Sql := Sql + ' FROM ' + TableName;
	Sql := Sql + iif(Where   <> '', ' WHERE '    + Where  , '');
  Sql := Sql + iif(OrderBy <> '', ' ORDER BY ' + OrderBy, '');

  Q := OpenSQL(Sql, true, 1);
  try
    if not Q.Eof then
    begin
      SetLength(Result, High(FieldsName) + 1);
      for i := Low(Fieldsname) to High(FieldsName) do
      begin
      	Result[i] := null;
        Result[i] := Q.Fields[i].Value;

				if VarIsNull(Result[i]) then
        begin
			   	TypeField := wGetTypeField(FieldsName[i]);
					if (TypeField = 'INTEGER') or (TypeField = 'SMALLINT') then Result[i] := 0
			   	else if (TypeField = 'DOUBLE') or (TypeField = 'RATE') or (TypeField = 'EXTENDED') then Result[i] := 0.0
			   	else if (TypeField = 'DATE') then Result[i] := iDate1900
					else Result[i] := ''
			  end;
      end;
    end;
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 19/09/2001
Modifié le ... :   /  /
Description .. : Retourne le nombre de champs composant une clé
Mots clefs ... :
*****************************************************************}
function WGetNbrFieldInKey (Const TableName: String): Integer;
var
  iTable: Integer;
  S,St: String;
  Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  S := Mcd.TableToCle1(TableName);
  St := ReadTokenPipe ( S, ',' );
  Result := 0;
  while St <> '' do
  begin
    Inc ( Result );
    St := ReadTokenPipe ( S, ',' );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 22/09/2001
Modifié le ... :   /  /
Description .. : Cette fonction supprime les lignes
Suite ........ : Les différents paramètres sont :
Suite ........ : La table
Suite ........ : La liste des champs concernés
Suite ........ : Les valeurs des champs  concernés
Suite ........ : Result := -6 indique que les paramètres sont incorrectes.
Suite ........ : Result := >= 0 indique le nombre de ligne supprimée
Suite ........ : Result := < -1 Erreur
Mots clefs ... : SQL;DELETE
*****************************************************************}
function WDeleteSql ( const TableName: String ; Const FieldsName: array of String ; Const FieldsValue: array of variant ) : Integer ;
var
	ValeurDuWhere: String ;                  // Retour de la fonction FFabriqueUpdate
begin
   { Valeur de retour par défaut }
  Result := -1 ;
  try
    { Génération de la close SET }
  	ValeurDuWhere := WFabriqueWhere ( FieldsName, FieldsValue );

    { Est-ce que la close est remplie ? }
    if ValeurDuWhere <> '' then
      Result := ExecuteSQL( 'DELETE FROM ' + TableName + ' WHERE ' + ValeurDuWhere ) ;
  except
    on E: Exception do
    begin
      Result := -6 ;
    end ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 02/03/2005
Modifié le ... :   /  /
Description .. : Permet de réaliser un InsertInto à partir d'un select.
Suite ........ : Contrôle le nombre de champs.
Suite ........ : Met les champ du select dans l'ordre
Mots clefs ... :
*****************************************************************}
function InsertIntoTable(Const TableName, Select, From: string): integer;
var
  iTable, iField, De, A, P                 : integer;
  FieldsName, LastErrorMsg, MonSelect, Sql,NomChamps,TypeChamps : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
  II : integer;
begin
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();
  //
  LastErrorMsg := '';
  Table := Mcd.GetTable(tableName);

  if Assigned(Table) then
  begin
    FieldList := Table.Fields;
    FieldList.Reset();
    II := 1;
    While FieldList.MoveNext do
    begin
      NomChamps := (FieldList.Current as IFieldCOM).name;
      TypeChamps := (FieldList.Current as IFieldCOM).Tipe;
      if Pos('AS ' + NomChamps , Select) > 0 then
        FieldsName := FieldsName + iif(II > 1, ',', '') + NomChamps;
      if (Pos('AS ' + NomChamps, Select) <= 0) and
         (TypeChamps <> 'BLOB') then
      begin
        LastErrorMsg := 'Le champ n''existe pas dans le select (AS ' + NomChamps + ')';
        Break;
      end
      else
      if Pos('AS ' + NomChamps, Select) <= 0 then continue
      else
      begin
        { De }
        De := Pos('AS ' + NomChamps, Select);
        P  := 0;
        While (De > 1) and ((Select[De] <> ',') or (p > 0)) do
        begin
          Case Select[De] of
            ')': Inc(p);
            '(': Dec(p);
          end;
          Dec(De);
        end;
        if Select[De] = ',' then Inc(De);

        { A }
        A := Pos('AS ' + NomChamps, Select);
        While (A <= Length(Select)) and (Select[A] <> ',') do
          Inc(A);

        MonSelect := MonSelect + iif(iField > 1, ',', '') + Copy(Select, De, A-De);
      end;
      inc(II);
    end;

    { Constitution du sql }
    if LastErrorMsg = '' then
    begin
      Sql := 'INSERT INTO ' + TableName
           + ' (' + Fieldsname + ')'
           + ' SELECT ' + MonSelect
           + ' FROM ' + From
           ;
    end;
  end
  else
  begin
    LastErrorMsg := 'La table: ' + TableName + ' n''existe pas';
  end;

  { Exécution du Sql }
  if LastErrorMsg = '' then
  begin
    Result := ExecuteSql(Sql)
  end
  else
  begin
    Result := -1;
    if V_Pgi.Sav then
      PgiError(LastErrorMsg, 'wCommuns.InsertIntoTable')
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 05/10/2001
Description .. : Retrouve dans le dataset, la valeur des champs de la clé 1
Suite ........ : d'une table.
Modifié le ... : 03/04/2002 par Thierry
Description .. : Suite ordre changement de tri dans V_PGI.DECHAMPS
Mots clefs ... :
*****************************************************************}
function wGetValueClef1(const TableName: String; Const F: TForm; Lequel: string=''): String ; overload;
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
var
	iTable, iChamp: Integer;
	Champ, TypeField: String;
	Value: String;
	DS: TDataSet;
	Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;

  function GetLequel: string;
  begin
    if Lequel = '' then
    begin
      if F is TFFICHE then
        Result := TFFICHE(F).UniqueName
      else
 		    Result := WGetFieldsClef1(TableName);
    end
    else
      Result := StringReplace(Lequel, ';', ',', [rfReplaceAll]);
  end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
begin
	Result := '' ;
  MCD := TMCD.GetMcd;
	if not mcd.loaded then mcd.WaitLoaded();

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
	{ Récupération du DataSet de la fiche }
  DS := nil;
  if Assigned(F) then
  begin
    {$IFNDEF EAGLCLIENT}
    if F is TfMul then DS := TDataSet(TFMul(F).Q)
    {$ELSE}
    if F is TfMul then
    begin
      TFMul(F).Q.TQ.seek(TFMul(F).FListe.Row - 1);
      DS := TDataSet(TFMul(F).Q.TQ)
    end
   {$ENDIF}
    else if F is TFFICHE then
      DS := TDataSet(TFFICHE(F).QFiche)
    else if F is TFFicheListe then
      DS := TDataSet(TFFicheListe(F).Ta)
    else if F is TFSaisieList then
      DS := TFSaisieList(F).LaTOF.GetDataSet
    else DS := nil;
  end;

  if Assigned(DS) then
  begin
    Table := Mcd.GetTable(TableName);

    if Assigned(table) then
    begin
      { Récupération des champs de la clef1 de la table }
      Lequel := GetLequel;

      { Ecrit la chaine de résultat }
      while (Lequel <> '') do
      begin
        Champ := Trim(ReadTokenPipe(Lequel, ','));
        TypeField := (Mcd.GetField(Champ) as IFieldCOM).tipe ;
        if ( TypeField = 'INTEGER' ) or ( TypeField = 'LONGINT' ) then
          Value := IntToStr ( DS.FindField ( Champ ).AsInteger )
        else if ( TypeField = 'RATE' ) OR ( TypeFIeld = 'DOUBLE' ) OR ( TypeField = 'EXTENDED' ) then
          Value := FloatToStr ( DS.FindField ( Champ ).AsFloat )
        else if ( TypeFIeld = 'DATE' ) then
          Value := DateTimeToStr( DS.FindField ( Champ ).AsDateTime )
        else if ( TypeField = 'BLOB' ) OR ( TypeField = 'DATA' ) then
          Value := ''
        else { Les VarChar, Char, Boolean }
          Value := DS.FindField ( Champ ).AsString ;
        Result := Result + Value + ';';
      end;
      if Result <> '' then
        Delete(Result, Length(Result), 1)
    end;
  end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end ;

function wGetValueClef1(Const T: Tob): String;
var
  Clef1: String;
begin
  Result := '';
  if Assigned(T) and (T.NomTable <> '') then
  begin
    Clef1 := wMakeFieldString(T.NomTable, ';');
    while Clef1 <> '' do
      Result := Result + T.GetString(ReadTokenSt(Clef1)) + ';';

    if Result <> '' then
      Delete(Result, Length(Result), 1)
  end;
end;

function wGetValueClef1(Const TableName: String; Const TobData: Tob): String; overload;
var
  FieldName, Clef1: String;
begin
  Clef1 := wMakeFieldString(TableName, ';');
  Result := '';
  if Assigned(TobData) then
  begin
    while (Clef1 <> '') do
    begin
      FieldName := ReadTokenSt(Clef1);
      if TobData.FieldExists(FieldName) then
        Result := Result + TobData.GetString(FieldName);
      Result := Result + ';'
    end;
    if Result <> '' then
      Delete(Result, Length(Result), 1)
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 22/09/2001
Modifié le ... :   /  /
Description .. : Cette fonction permet de récupérer le clef1 de la table
Suite ........ : passée en paramètre.
Mots clefs ... : SQL;CLEF;KEY
*****************************************************************}
function WGetFieldsClef1 ( Const TableName: String  ) : String ;
var
	iTable: Integer; // Indice de la table dans V_PGI.DETABLES
	Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  Result := Mcd.TableToCle1(TableName);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 22/09/2001
Description .. : Cette fonction permet de convertir des strings en valeur du type de champ
Suite ........ : Seuls les champs DATA et BLOB ne sont pas traités.
Suite ........ : Attention, la table doit être présente dans le dictionnaire.
Modifié le ... : 03/04/2002 Par Thierry
Description .. : Suite changement ordre de tri de V_PGI.DECHAMP en valeur du type de champ
Mots clefs ... : FIELD;SQL;CONVERTIR
*****************************************************************}
//GP_20071116_MM_GP14563 Déb
function wConvertStringToField(Const TableName, FieldName, Value: String) : Variant ;
var
  FieldType: String;
	Mcd : IMCDServiceCOM;
	Field     : IFieldCOM ;
begin
  { Valeur de retour par défaut }
  Result := Null;
  //
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldType := (Mcd.GetField(FieldName) as IFieldCOM).tipe ;
  if ( FieldType = 'INTEGER' ) OR ( FieldType = 'SMALLINT' ) then
    Result := ValeurI ( Value )
  else if ( FieldType = 'DATE' ) then
    Result := Value
  else if ( FieldType = 'DOUBLE' ) OR ( FieldType = 'RATE' ) OR ( FieldType = 'EXTENDED' ) then
    Result := Valeur( Value )
  else
  begin
    Result := Value;
    {$IFNDEF EAGLCLIENT}
      if (Result = '') and isOracle() then
        Result := Null
    {$ENDIF !EAGLCLIENT}
  end
end;
//GP_20071116_MM_GP14563 Fin

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 05/10/2001
Modifié le ... :   /  /
Description .. : Copie les champs d'un dataset dans une TOB
Mots clefs ... :
*****************************************************************}
procedure wMakeTobFromDS(Const DS: TDataSet; Const TheTob: TOB; Const OnRecords: Boolean = False; Const ToutEnAddChampSup: Boolean = False);
var
  T: Tob;

  procedure LoadTobFromDS(UneTob: Tob);
  var
    i:Integer;
  begin
    for i:=0 to DS.FieldCount-1 do
    begin
      if UneTob.FieldExists(DS.Fields[i].FieldName) or ToutEnAddChampSup then
      begin
        if ToutEnAddChampSup then
        begin
          if DS.Fields[i].AsVariant <> null then
            UneTob.AddChampSupValeur(DS.Fields[i].FieldName, DS.Fields[i].AsVariant)
          else
            UneTob.AddChampSupValeur(DS.Fields[i].FieldName, wGetInitValue(DS.Fields[i].FieldName))
        end
        else
        begin
          if DS.Fields[i].AsVariant <> null then
            UneTob.PutValue(DS.Fields[i].FieldName, DS.Fields[i].AsVariant)
          else
            UneTob.PutValue(DS.Fields[i].FieldName, wGetInitValue(DS.Fields[i].FieldName))
        end;
      end;
    end;
  end;

begin
  if Assigned(DS) and Assigned(TheTob) then
  begin
    if not OnRecords then
      LoadTobFromDS(TheTob)
    else
    begin
      DS.First;
      while not DS.Eof do
      begin
        T := Tob.Create(TheTob.NomTable, TheTob, -1);
        LoadTobFromDS(T);
        DS.Next;
      end;
      DS.First;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/10/2001
Modifié le ... :   /  /
Description .. : Copie les valeurs des fields d'un dataset dans une TOB,
Description .. : Si ceux-ci ont été modifiés dans la TOB
Description .. : A utiliser conjointement avec la procédure ci-dessus
Mots clefs ... : TOB;DATASET
*****************************************************************}
Procedure WSetDSFromTob(Const TheTob:TOB; Const DS:TDataSet);
var
  i:Integer;
begin
  if (Assigned(DS)) and (Assigned(TheTob)) and (TheTob.IsOneModifie) then
  begin
    if DS.State = dsBrowse then DS.Edit; {PMJEAGL}
    for i:=0 to DS.FieldCount-1 do
    begin
      if TheTob.FieldExists(DS.Fields[i].FieldName) then
      begin
        DS.Fields[i].Value := TheTob.GetValue(DS.Fields[i].FieldName);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/10/2001
Modifié le ... :   /  /
Description .. : Répète x fois la même chaine
Mots clefs ... :
*****************************************************************}
function wStringRepeat(Const Chaine: string; Nb: integer): string;
begin
	Result := '';
  While Nb > 0 do
  begin
   	Result := Result + Chaine;
    Dec(Nb);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/05/2003
Modifié le ... :   /  /
Description .. : Converti une chaîne de caractères en un tableau de
Suite ........ : chaînes :
Suite ........ : - s: Chaîne à convertir
Suite ........ : - Delimiteur: Délimiteur de tokens
Mots clefs ... :
*****************************************************************}
function wStringToArray(Chaine: string; Const Delimiteur: String): MyArrayField;
begin
  SetLength(Result, 0);
  repeat
    SetLength(Result, Length(Result) + 1);
    Result[Length(Result) - 1] := Trim(ReadTokenPipe(Chaine, Delimiteur));
  until Chaine = '';
  if (Length(Result) = 1) and (Result[Length(Result) - 1] = '') then
    SetLength(Result, 0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2006
Modifié le ... :   /  /    
Description .. : Converti un tableau en chaîne de caractères tokenisée
Mots clefs ... : 
*****************************************************************}
//GP_20071217_TS_GP14539 >>>
function wArrayToString(MyArray: MyArrayField; Const Delimiteur: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := Low(MyArray) to High(MyArray) do
    Result := Result + iif(Result <> '', Delimiteur, '') + MyArray[i];
end;
//GP_20071217_TS_GP14539 <<<
function wArrayToString(MyArray: MyArrayValue; Const Delimiteur: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := Low(MyArray) to High(MyArray) do
    Result := Result + iif(Result <> '', Delimiteur, '') + VarToStr(MyArray[i]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 04/02/2005
Modifié le ... :   /  /    
Description .. : Convertit une chaîne quelconque en une chaîne 
Suite ........ : composée de caractères amlpha numériques
Mots clefs ... : 
*****************************************************************}
function wStringToAlphaNumString(const St: String; const CharsToIgnore: String = ''): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(St) do
    if ((St[i] >= 'a') and (St[i] <= 'z'))
    or ((St[i] >= 'A') and (St[i] <= 'Z'))
    or ((St[i] >= '0') and (St[i] <= '9'))
    or (Pos(St[i], CharsToIgnore) > 0) then
      Result := Result + St[i]
end;

function wFileSizeToStr(FileSize: Integer; const NbDecimal: Integer): String;
var
  FileSize2, Cpt, Res: Integer;
  ResDecimal: Real;
  sUnite: String;
begin
  ResDecimal := FileSize;
  Cpt := 0;
  repeat
    FileSize2 := FileSize div 1024;
    FileSize := FileSize mod 1024;
    if FileSize > (1024 div 2) then
      Inc(FileSize2);
    if FileSize2 > 0 then
      Inc(Cpt);
  until FileSize div 1024 = 0;

  Res := iif(FileSize2 > 0, FileSize2, FileSize);

  case Cpt of
    0: sUnite := ' octet' + iif(Res > 1, 's', '');
    1: begin
         sUnite := ' ko';
         ResDecimal := ResDecimal / 1024
       end;
    2: begin
         sUnite := ' mo';
         ResDecimal := ResDecimal / Sqr(1024)
       end;
  else
    sUnite := ' ??o';
  end;
  Result := FloatToStr(Round(ResDecimal * IntPower(10, NbDecimal)) / IntPower(10, NbDecimal)) + sUnite;
end;  

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 06/12/2001
Modifié le ... :   /  /
Description .. : Sous chaine de droite
Mots clefs ... :
*****************************************************************}
Function wRight(Const Chaine: string; const Long: integer): string;
begin
	Result := Copy(Chaine, Length(Chaine) - Long + 1, Long);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 06/12/2001
Modifié le ... :   /  /
Description .. : Sous chaine de gauche
Mots clefs ... :
*****************************************************************}
Function wLeft(Const Chaine: string; Const Long: integer): string;
begin
	Result := Copy(Chaine, 1, Long);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 19/12/2001
Modifié le ... :   /  /
Description .. : Retourne le libelle d'un champ dans DECHAMPS
Description .. : à partir de son nom
Mots clefs ... :
*****************************************************************}
Function wGetLibChamp(Const FieldName:String):String;
var
	Mcd : IMCDServiceCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  result := (Mcd.GetField(FieldName) as IFieldCOM).Libelle;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 26/03/2002
Modifié le ... :   /  /
Description .. : Test l'existence d'un nom de champ dans V_PGI.Dechamp
Mots clefs ... : DECHAMP;
*****************************************************************}
function wExistFieldInDechamps(Const FieldName:String): Boolean;
var
  iTable: Integer;
begin
  iTable := TableToNum( PrefixeToTable( ExtractPrefixe( FieldName ) ) );
  {$IFDEF EAGLCLIENT}
    if High(V_Pgi.DEChamps[iTable]) <= 0 then
      ChargeDeChamps(iTable, ExtractPrefixe(FieldName));
  {$ENDIF EAGLCLIENT}
  if iTable <> -1 then
    Result := ChampToNum(FieldName) <> -1
  else
    Result := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/09/2002
Modifié le ... :   /  /
Description .. : Test l'existence d'un nom de table dans V_PGI.Dttable
Mots clefs ... : DTTABLE;
*****************************************************************}
function wExistTableInDetables(Const TableName:String): Boolean;
begin
	Result := TableToNum(TableName) <> 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 27/12/2001
Modifié le ... :   /  /
Description .. : Récupère dans une tstrings la liste des noms de champs
Suite ........ : d'une table
Mots clefs ... : DECHAMPS;DETABLE;
*****************************************************************}
procedure wGetTableFieldsName(Const TableName: String; TSCodeChamp: HTStrings; TSLibChamp: HTStrings = nil);
var
	iTable,i : Integer;
	Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
  Field     : IFieldCOM ;
begin
  if Assigned(TSCodeChamp) then
  begin
		MCD := TMCD.GetMcd;
		if not mcd.loaded then mcd.WaitLoaded();
    //
    Table := Mcd.GetTable(TableName);
    if assigned (table) then
    begin
      FieldList := Table.Fields;
      FieldList.Reset();
      While FieldList.MoveNext do
      begin
      	Field := FieldList.Current as IFieldCOM ;
        if assigned(field) then
        begin
          TSCodeChamp.Add(Field.Name);
          if Assigned(TSLibChamp) then
            TSLibChamp.Add(Field.Libelle)
        end;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin & Marc Morretton
Créé le ...... : 02/01/2002
Modifié le ... :   /  /
Description .. : Récupère le contenu d'une tablette de Communs dans une TOB
Mots clefs ... :
*****************************************************************}
function wGetTabletteFromCommun(Const sType: String; Const TheTob: TOB; Const OrderByAbrege: Boolean = False): Boolean;
var
	Sql: string;
begin
	Sql := 'SELECT *'
       + ' FROM COMMUN'
       + ' WHERE CO_TYPE="' + sType + '"'
       + iif(OrderByAbrege, ' ORDER BY CO_ABREGE', '')
       ;
	Result := (Assigned(TheTob)) and TheTob.LoadDetailDBFromSql('COMMUN', Sql)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 03/03/2005
Modifié le ... :   /  /    
Description .. : Test une valeur dans une tablette
Mots clefs ... : 
*****************************************************************}
function wExistCO(const sType, Code: String): Boolean;
begin
  Result := ExisteSQL('SELECT 1'
                    + ' FROM COMMUN'
                    + ' WHERE CO_TYPE="' + sType + '"'
                    + ' AND CO_CODE="' + Code + '"')
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gérard JUGDE
Créé le ...... : 13/03/2006
Modifié le ... : 13/03/2006
Description .. : Test une valeur dans une tablette CHOIXCOD
Mots clefs ... : 
*****************************************************************}
function wExistCC(const sType, Code: String): Boolean;
begin
  Result := ExisteSQL('SELECT 1'
                    + ' FROM CHOIXCOD'
                    + ' WHERE CC_TYPE="' + sType + '"'
                    + ' AND CC_CODE="' + Code + '"')
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/09/2002
Modifié le ... :   /  /
Description .. : Récupère le contenu d'une tablette de choixcod dans une
Suite ........ : tob
Mots clefs ... :
*****************************************************************}
function wGetTabletteFromChoixCod(Const sType: String; Const TheTob: TOB): Boolean;
var
  sql: String;
begin
	Result := False;
  if Assigned(TheTob) then
  begin
   	sql := 'SELECT *'
         + ' FROM CHOIXCOD'
         + ' WHERE CC_TYPE="' + sType + '"';
    Result := TheTob.LoadDetailDBFromSql('CHOIXCOD', sql);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 11/09/2002
Description .. : Récupère le contenu d'une tablette de choixcod dans une
Suite ........ : tob
*****************************************************************}
function wGetTabletteFromChoixExt(Const sType: String; Const TheTob: TOB): Boolean;
var
  Sql: String;
begin
	Result := False;
  if Assigned(TheTob) then
  begin
   	sql := 'SELECT *'
         + ' FROM CHOIXEXT'
         + ' WHERE YX_TYPE="' + sType + '"';
    Result := TheTob.LoadDetailDBFromSql('CHOIXEXT', Sql);
  end;
end;

{-----------------------------------------------------------------------------------------------------------------------}
{- Protection de la division par 0 avec ou Sans Message  ---------------------------------------------------------------}
{-----------------------------------------------------------------------------------------------------------------------}
function wDivise(Const Dividande, Diviseur: double; Const WithMessage: boolean = False) : double;
begin
	if Diviseur = 0 then
  begin
  	Result := 0;
    if WithMessage then
      PGIError(TraduireMemoire('Erreur : Division par 0'), '');
  end
  else
   	Result := Dividande / Diviseur;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 07/02/2002
Modifié le ... : 07/02/2002
Description .. : Controle dans une chaine que  tous ses caractères sont
Suite ........ : bien autorisés.
Mots clefs ... :
*****************************************************************}
Function wCtrlCaractere(Const sMessage, sAControler, sAutorise : string) : string;
var
  sCaractereATester : string;
  iPos : integer;
begin
  Result := '';
  for iPos:=1 to length(sAControler) do
  begin
    sCaractereATester := copy(sAcontroler,iPos,1);
    if (pos(sCaractereATester,sAutorise)=0) then
    begin
      Result:=Result+iif(sCaractereATester=' ','espace',sCaractereATester)+' ';
    end;
  end;
  Result:=iif(Result<>'',sMessage+' '+sAControler+' contient des caractères interdits : '+Result,'');
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFNDEF EAGLCLIENT}
function wMakeWhereFromList(Const Prefixe: string; Const Liste: ThDbGrid; Const Q: tQuery; Const Cle: string=''): string;
{$ELSE}
function wMakeWhereFromList(Const Prefixe: string; Const Liste: ThGrid; Const Q: tQuery; Const Cle: string=''): string;
{$ENDIF}
var
	i: integer;
  {$IFDEF EAGLCLIENT}
    wBookMark: integer;
  {$ELSE}
    wBookMark: tBookMark;
  {$ENDIF}

	procedure AddToWhere(var Where: string);
  var
    sClef1: String;

    function GetFieldsValues(sClef: String): String;
    begin
      Result := '';
      if sClef <> '' then
      	Result := Q.findField(ReadTokenSt(sClef)).AsString;
      while sClef <> '' do
        Result := Result + ';' + Q.findField(ReadTokenSt(sClef)).AsString;
    end;

  begin
    if Cle <> '' then
      sClef1 := Cle
    else
      sClef1 := wMakeFieldString(PrefixeToTable(Prefixe), ';');
    Where := Where + iif(Where <> '', ' OR ', '') + '(' + wMakeWhereSQL(sClef1, GetFieldsValues(sClef1)) + ')';
  end;

begin
  Result := '';
	if (Liste.NbSelected > 0) or (Liste.AllSelected) then
  begin
    Q.DisableControls;
    try
      {$IFDEF EAGLCLIENT}
      wBookMark := Liste.Row; {PMJEAGL}
      if Liste.AllSelected then
      begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          AddToWhere(Result);
          Q.next;
        end;
      end
      else
      begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          if Liste.IsSelected( i ) then
            AddToWhere(Result);
          Q.next;
        end;
      end;
      Liste.Row := wBookMark;
      if Liste.Row>0 then
        Q.Seek( Liste.Row - 1 )
      else
        Q.First;
      {$ELSE}
      wBookMark := Liste.DataSource.DataSet.GetBookmark;
      if Liste.AllSelected then
      begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          AddToWhere(Result);
          Q.next;
        end;
      end
      else
      begin
        for i := 0 to Liste.NbSelected - 1 do
        begin
          Liste.GotoLeBOOKMARK(i);
          AddToWhere(Result);
        end;
      end;
      Q.GotoBookmark(wBookMark);
      {$ENDIF}
    finally
      Q.EnableControls;
    end;
  end
  else
  begin
    {$IFDEF EAGLCLIENT}
    { Resyncriniser le Q }
    Q.seek(Liste.Row-1);
    {$ENDIF}
	  AddToWhere(Result);
  end;

  Result := '(' + Result + ')';
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 21/10/2004
Modifié le ... :   /  /
Description .. : Construit un WHERE avec les lignes sélectionnées par la
Suite ........ : sélection multiple.
Mots clefs ... :
*****************************************************************}
function wMakeWhereFromFSL(Const Prefixe: string; Const Liste: THGrid; Const DS: tDataSet; Const Cle: string=''): string;
Var
  i: integer;

	procedure AddToWhere(var Where: string);
    function GetFieldsValues(sClef: String): String;
    begin
      Result := '';
      while sClef <> '' do
        Result := Result + iif(Result = '', '', ';') + DS.FindField(ReadTokenSt(sClef)).AsString;
    end;
  begin
    if Cle <> '' then
    	Where := Where + iif(Where <> '', ' OR ', '') + '(' + wMakeWhereSQL(Cle, GetFieldsValues(Cle)) + ')'
    else
	    Where := Where + iif(Where <> '', ' OR ', '') + Prefixe + '_IDENTIFIANT=' + DS.FindField(Prefixe + '_IDENTIFIANT').AsString;
  end;

begin
//GP_20080606_DKZ_GP15168 Déb
  Result := '';
  if (Liste.NbSelected > 0) or (Liste.AllSelected) then
  begin
    i := Liste.FixedRows;
    DS.First;
    try
      while not DS.EOF do
      begin
        if (Liste.IsSelected(i)) then
          AddToWhere(Result);

        Inc(i);
        DS.next;
      end;
    finally
      DS.First; // pour ne pas positionner le dataset en fin de fichier (FQ GP15168).
    end
  end
  else
  begin
    {$IFDEF EAGLCLIENT}
      Liste.GotoLeBookMark(0);
      DS.Seek(Liste.Row-Liste.FixedRows);
    {$ENDIF EAGLCLIENT}

    AddToWhere(Result);
  end;
//GP_20080606_DKZ_GP15168 Fin
end;

(*
procedure AGLwLanceEtat (parms: array of variant; nb: integer);
var
	NatureEtat, Etat : String;
begin
  NatureEtat := parms[0];
  Etat := parms[1];
  LanceEtat('E', NatureEtat, Etat, True, False, False, Nil, '', '', False);
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 08/03/2002
Modifié le ... :   /  /
Description .. : Delete d'une tob
Mots clefs ... :
*****************************************************************}
function wDeleteTob(TableName: string; TheTob: Tob; WithProgressForm: Boolean; Rapport: TWRapport = nil): Boolean;
var
	iTob 	: integer;
  {$IFDEF GESCOM}
    TheTom: tTomComm;
  {$ELSE GESCOM}
    TheTom: Tom;
  {$ENDIF GESCOM}  
begin
  {$IFDEF GESCOM}
    TheTom := tTomComm(CreateTOM(TableName , nil, false, true));
  {$ELSE GESCOM}
    TheTom := CreateTOM(TableName , nil, false, true);
  {$ENDIF GESCOM}
  {$IFDEF EAGLSERVER}
    WithProgressForm := False;
  {$ENDIF EAGLSERVER}
  if withProgressForm then
    wInitProgressForm(nil, TraduireMemoire('Suppression en cours'), '', TheTob.Detail.Count, False, True);
  try
    Result := TheTob.Detail.Count > 0;
    for iTob := 0 to Pred(TheTob.Detail.Count) do
    begin
      if withProgressForm then
        wMoveProgressForm;
      { OnDelete }
      TheTob.Detail[iTob].AddChampSupValeur('IKC', 'S');
      if Assigned(Rapport) then
        TheTob.Detail[iTob].AddChampSupValeur(WRPTobFieldName, LongInt(Rapport));
      try
        Result := TheTom.DeleteTOB(TheTob.Detail[iTob]) and TheTob.Detail[iTob].DeleteDB;

        {$IFDEF GESCOM}
          if Result then
            TheTom.AfterDeleteTob(TheTob.Detail[iTob]);
        {$ENDIF GESCOM}
      finally
        if Assigned(Rapport) then
          TheTob.Detail[iTob].DelChampSup(WRPTobFieldName, false);
      end;
    end;
  finally
    if withProgressForm then
      wFiniProgressForm;
    TheTom.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 08/03/2002
Modifié le ... :   /  /
Description .. : Delete d'une table
Mots clefs ... :
*****************************************************************}
function wDeleteTable(Const TableName, Where: string; Const WithProgressForm: Boolean = False; Const Rapport: TWRapport = nil): Boolean;
var
	Sql	: string;
	TheTob: Tob;
begin
  Sql := 'SELECT *'
       + ' FROM ' + TableName
       + ' WHERE ' + Where
       ;

  if Where <> '' then
  begin
    TheTob := Tob.Create('TOB', nil, -1);
    try
      Result := TheTob.LoadDetailDBFromSql(TableName, Sql) and wDeleteTob(TableName, TheTob, {$IFNDEF ERADIO}withProgressForm{$ELSE}False{$ENDIF !ERADIO}, Rapport);
    finally
      TheTob.Free;
    end;
  end
  else
  begin
    if V_PGI.SAV then PGIINFO('Attention ! Where vide ! : ' + #13 + #10 + Sql, 'SAV');
    Result := False;
  end;
end;

function wDeleteTable2(Const TableName, Sql: string; Const WithProgressForm: Boolean = False; const Rapport: TWRapport = nil): Boolean;
var
	TheTob: Tob;
begin
  TheTob := Tob.Create('TOB', nil, -1);
  try
    Result := TheTob.LoadDetailDBFromSql(TableName, Sql) and wDeleteTob(TableName, TheTob, {$IFNDEF ERADIO}withProgressForm{$ELSE}False{$ENDIF !ERADIO}, Rapport);
  finally
    TheTob.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : Fixe la taille d'un champ caractère
Suite ........ :
Suite ........ : en entrée
Suite ........ :    * chaine de caractère
Suite ........ :    * taille fixe
Suite ........ :
Suite ........ : en sortie
Suite ........ :    * chaine à longueur fixe
Mots clefs ... :
*****************************************************************}
function wPadRight(Const Chaine:string; Const Long:integer; Const Remplissage:string = ' ') : string;
begin
  Result := wLeft(Chaine + wStringRepeat(Remplissage, Long), Long);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : KOZA Denis
Créé le ...... : 20/06/2002
Modifié le ... : 20/06/2002
Description .. : Fixe la taille d'une chaîne
Suite ........ : en complétant à gauche avec un caractère
Suite ........ :
Suite ........ : Entrée : Chaîne, Longueur fixe, Caractère
Suite ........ :
Suite ........ : Sortie : Chaîne de longueur fixe
Mots clefs ... :
*****************************************************************}
function wPadLeft(Const Chaine:string; Const Long:integer; Const Remplissage:string = ' ') : string;
begin
  if Length(Chaine) < Long then
    Result := wStringRepeat(Remplissage, Long - Length(Chaine)) + Chaine
  else
    Result := Copy(Chaine, 1, Long );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 18/03/2002
Modifié le ... :   /  /
Description .. : Replicate un caractère un certain nombre de  fois
Suite ........ :
Suite ........ : en entrée :
Suite ........ :    * caractère à répliquer
Suite ........ :    * nombre de replication
Suite ........ :
Suite ........ : en sortie
Suite ........ :    * chaine de caractère
Mots clefs ... :
*****************************************************************}
function wReplicate(s:string; iLong:integer) : string;
var
	i: integer;
begin
  Result:='';
  for i:=1 to iLong do
    Result := Result + s;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 18/03/2002
Modifié le ... : 18/03/2002
Description .. : Transforme une date + une heure en une chaine de
Suite ........ : caractère de format précisé dans l'appel :
Suite ........ :
Suite ........ : Date du 14/07/2002 à 11 h 55 minutes et 23 secondes
Suite ........ :
Suite ........ : Format :
Suite ........ :
Suite ........ : AAAA                           -> 2002
Suite ........ : AAAA-S                         -> 2002-2
Suite ........ : AAAA-S-T                       -> 2002-2-3
Suite ........ : AAAA-S-T-MM                    -> 2002-2-3-07
Suite ........ : AAAA-S-T-MM-JJ                 -> 2002-2-3-07-14
Suite ........ : AAAA-S-T-MM-JJ-HH              -> 2002-2-3-07-14-11
Suite ........ : AAAA-S-T-MM-JJ-HHMM            -> 2002-2-3-07-14-1155
Suite ........ : AAAA-S-T-MM-JJ-HHMMSS          -> 2002-2-3-07-14-115523
Suite ........ :
Suite ........ : Par défaut le paramètre est AAAAMMJJ
Suite ........ :
Mots clefs ... :
*****************************************************************}
function wDateTimeToStr(Const LaDate: tDateTime; Const cFormat:string = 'AAAA-S-T-MM-JJ'; Const lFormat: boolean = True): string;
var
  dDate : tDateTime;
begin
  if lFormat then
  begin
  	Result := DateTimeToStr(LaDate);
	  Result := copy(Result,7,4)+'-'+FloatToStr(wArrondir(wMonth(LaDate)/6, tAPUnite, tAMSuperieure))+'-'+FloatToStr(wArrondir(wMonth(LaDate)/3, tAPUnite, tAMSuperieure))+'-'+copy(Result,4,2)+'-'+copy(Result,1,2)+'-'+copy(Result,12,2)+copy(Result,15,2)+copy(result,18,2);
    Result := wLeft(Result,length(cFormat));
  end
  else
  begin
    if      (cFormat='AAAA') then
      dDate := AddAnnee(LaDate, 0, 'D')
    else if (cFormat='AAAA-S') then
      dDate := AddSemestre(LaDate, 0, 'D')
    else if (cFormat='AAAA-S-T') then
      dDate := AddTrimestre(LaDate, 0, 'D')
    else if (cFormat='AAAA-S-T-MM') then
      dDate := AddMois(LaDate, 0, 'D')
    else if (cFormat='AAAA-S-T-MM-JJ') then
      dDate := AddJour(LaDate, 0, 'D')
    else
      dDate := LaDate;

    Result := DateTimeToStr(dDate);
	  Result := copy(Result,7,4)+'-'+FloatToStr(wArrondir(wMonth(dDate)/6, tAPUnite, tAMSuperieure))+'-'+FloatToStr(wArrondir(wMonth(dDate)/3, tAPUnite, tAMSuperieure))+'-'+copy(Result,4,2)+'-'+copy(Result,1,2)+'-'+copy(Result,12,2)+copy(Result,15,2)+copy(result,18,2);
  end;
end;

{-------------------------------------------------------------------------------
  Additionne un nombre d'années à la date demandée
  Renvoie la date en initialisant le jour en fonction du paramètre demandé.
   * le 1er jour de l'année calculée
   * le dernier jour de l'année calculée
   * le même jour de l'année calculée
--------------------------------------------------------------------------------}
function AddAnnee(Const LaDate: tDateTime; Const iNbAn: integer; Const sJour: string = 'J'): tDateTime;
begin {AddAnnee}
  if      (sJour='J') then
    Result := PlusDate(LaDate, iNbAn, 'A') //StrToDateTime(wPadLeft(IntToStr(integer(wDay(LaDate))),2,'0')+'/'+wPadLeft(IntToStr(integer(wMonth(LaDate))),2,'0')+'/'+IntToStr(integer(wYear(LaDate)+iNbAn)))
  else if (sJour='D') then
    Result := StrToDateTime('01/01/'+IntToStr(wYear(LaDate) + iNbAn)+' 00:00:01')
  else if (sJour='F') then
    Result := StrToDateTime('31/12/'+IntToStr(wYear(LaDate) + iNbAn)+' 23:59:59')
  else
    Result := LaDate;
end; {AddAnnee}

{-------------------------------------------------------------------------------
  Additionne un nombre de semestres à la date demandée
  Renvoie la date en initialisant le jour en fonction du paramètre demandé.
   * le 1er jour du semestre calculé
   * le dernier jour du semestre calculé
   * le même jour du semestre calculé
--------------------------------------------------------------------------------}
function AddSemestre(Const LaDate: tDateTime; Const iNbSemestre: integer; Const sJour: string = 'J'): tDateTime;
var
  dDate : tDateTime;
begin {AddSemestre}
  if      (sJour='J') then
  begin
    Result := PlusDate(LaDate, iNbSemestre*6, 'M');
  end
  else if (sJour='D') then
  begin
    dDate  := PlusDate(LaDate, iNbSemestre*6, 'M');
    Result := StrToDateTime('01/'+wPadLeft(IntToStr(wGet1erMoisFromSemestre(wGetSemestreFromMois(wMonth(dDate)))),2,'0')+'/'+IntTostr(Integer(wYear(dDate)))+' 00:00:01');
  end
  else if (sJour='F') then
  begin
    dDate  := PlusDate(LaDate, (iNbSemestre+1)*6, 'M');
    Result := StrToDateTime('01/'+wPadLeft(IntToStr(wGet1erMoisFromSemestre(wGetSemestreFromMois(wMonth(dDate)))),2,'0')+'/'+IntTostr(Integer(wYear(dDate)))+' 23:59:59')-1;
  end
  else
    Result := LaDate;
end; {AddSemestre}

{-------------------------------------------------------------------------------
  Additionne un nombre de trimestres à la date demandée
  Renvoie la date en initialisant le jour en fonction du paramètre demandé.
   * le 1er jour du trimestre calculé
   * le dernier jour du trimestre calculé
   * le même jour du trimestre calculé
--------------------------------------------------------------------------------}
function AddTrimestre(Const LaDate: tDateTime; Const iNbTrimestre: integer; Const sJour: string = 'J'): tDateTime;
var
  dDate : tDateTime;
begin {AddTrimestre}
  if      (sJour='J') then
  begin
    Result := PlusDate(LaDate, iNbTrimestre*3, 'M');
  end
  else if (sJour='D') then
  begin
    dDate  := PlusDate(LaDate, iNbTrimestre*3, 'M');
    Result := StrToDateTime('01/'+wPadLeft(IntToStr(wGet1erMoisFromTrimestre(wGetTrimestreFromMois(wMonth(dDate)))),2,'0')+'/'+IntTostr(Integer(wYear(dDate)))+' 00:00:01');
  end
  else if (sJour='F') then
  begin
    dDate  := PlusDate(LaDate, (iNbTrimestre+1)*3, 'M');
    Result := StrToDateTime('01/'+wPadLeft(IntToStr(wGet1erMoisFromTrimestre(wGetTrimestreFromMois(wMonth(dDate)))),2,'0')+'/'+IntTostr(Integer(wYear(dDate)))+' 23:59:59')-1;
  end
  else
    Result := LaDate;
end; {AddTrimestre}

{-------------------------------------------------------------------------------
  Additionne un nombre de mois à la date demandée
  Renvoie la date en initialisant le jour en fonction du paramètre demandé.
   * le 1er jour du mois calculé
   * le dernier jour du mois calculé
   * le même jour du mois calculé
--------------------------------------------------------------------------------}
function AddMois(Const LaDate: tDateTime; Const iNbMois: integer; Const sJour: string = 'J'): tDateTime;
var
  dDate : tDateTime;
  sDate : string;
begin {AddMois}
  if      (sJour='J') then
  begin
    Result := PlusDate(LaDate, iNbMois, 'M');
  end
  else if (sJour='D') then
  begin
    dDate  := PlusDate(LaDate, iNbMois, 'M');
    sDate  := DateToStr(dDate);
    Result := StrToDateTime('01'+ copy(sDate,3,length(sDate)-2)+' 00:00:01');
  end
  else if (sJour='F') then
  begin
    dDate  := PlusDate(LaDate, iNbMois+1, 'M');
    sDate  := DateToStr(dDate);
    Result := StrToDateTime('01'+ copy(sDate,3,length(sDate)-2)+' 23:59:59')-1;
  end
  else
    Result := LaDate;
end; {AddMois}

{***********A.G.L.***********************************************
Auteur  ...... : Ka Joua VANG
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Ajoute ou retire le nombre de semaine
Suite ........ : Jour=D: se positionne en début de la nouvelle semaine
Suite ........ : Jour=F: se positionne en fin de semaine
Mots clefs ... : 
*****************************************************************}
function AddSemaine(Const LaDate: tDateTime; Const NbSemaine: integer; Const Jour: char = 'J'): tDateTime;
begin
  Result := LaDate + 7*NbSemaine;

  Case Jour of
    'D': while DayOfTheWeek(Result) > 1 do Result := Result-1;
    'F': while DayOfTheWeek(Result) < 7 do Result := Result+1;
  end;
end;

{-------------------------------------------------------------------------------
  Additionne un nombre de jour à la date demandée
  Renvoie la date en initialisant le jour en fonction du paramètre demandé.
   * la 1èere heure du jour calculé
   * la dernière heure du jour calculé
   * à k'heure du jour calculé
--------------------------------------------------------------------------------}
function AddJour(Const LaDate: tDateTime; Const iNbJour: integer; Const sJour: string = 'J'): tDateTime;
begin {AddJour}
  if      (sJour='J') then
  begin
    Result := PlusDate(LaDate, iNbJour, 'J');
  end
  else if (sJour='D') then
  begin
    Result := StrToDateTime(DateToStr(PlusDate(LaDate, iNbJour, 'J'))+' 00:00:01');
  end
  else if (sJour='F') then
  begin
    Result := StrToDateTime(DateToStr(PlusDate(LaDate, iNbJour, 'J'))+' 23:59:59');
  end
  else
    Result := LaDate;
end; {AddJour}

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 22/03/2002
Modifié le ... : 22/03/2002
Description .. : Dit si une Grid est multi-sélectionnée
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT} //PMJEAGL
function wMultiSelected(Const L : THDBGrid): boolean; overload;
{$ELSE  !EAGLCLIENT}
function wMultiSelected(Const L : THGrid): boolean;
{$ENDIF !EAGLCLIENT}
begin
	Result := (L.NbSelected > 0) or (L.AllSelected)
end;

{$IFNDEF EAGLCLIENT} 
function wMultiSelected(Const L : THGrid): boolean; overload;
begin
	Result := (L.NbSelected > 0) or (L.AllSelected)
end;
{$ENDIF !EAGLCLIENT}
{$ENDIF !EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/02/2003
Modifié le ... :   /  /
Description .. : Effectue un TOb.SelectDB à l'aide d'un query.
Suite ........ : Celà évite d'oublier le Try...Finally...End.
Mots clefs ... :
*****************************************************************}
function wSelectTobFromSQL(Const Sql: String; Const t: Tob; Const WithMemo: Boolean = True): Boolean;
var
  Q: tQuery;
begin
  Result := False;
	Q := OpenSql(Sql, True, 1);
  try
	  if not Q.Eof then
    begin
      Result := t.SelectDB('', Q, not WithMemo);
    end;
  finally
  	Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 18/10/2005
Modifié le ... :   /  /
Description .. : Charge une tob selon les champs d'une liste
Suite ........ : (ne prend que les champs de "droite", dans le 
Suite ........ : paramétrage de la liste)
Mots clefs ... :
*****************************************************************}
function wLoadTobFromListe(const DBListe: String; T: Tob; const OnlyExistingFields: Boolean = True): Boolean;
var
  wListe: TWListe;
begin
  wListe := TWListe.Create(DBListe, nil, OnlyExistingFields);
  try
    if (wListe.TablesCount > 0) and (wListe.ChampsCount > 0) then
      Result := T.LoadDetailDBFromSQL(TableToPrefixe(wListe.Tables[0]), wListe.SQL, False)
    else
      Result := False
  finally
    wListe.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : change le typefield pour tests divers zet variés
Description .. : char, varchar, combo, longchar => C
Description .. : integer, smallint => I
Description .. : double, extended, rate => N
Description .. : date => D
Description .. : boolean => B
Mots clefs ... :
*****************************************************************}
function GetSimpleTypeField(const TypeField: String): Char;
begin
  if      (TypeField = 'INTEGER') or (TypeField = 'SMALLINT') then Result := 'I'
  else if (TypeField = 'DOUBLE') or (TypeField = 'RATE') or (TypeField = 'EXTENDED') then Result := 'N'
  else if (TypeField = 'DATE') then Result := 'D'
  else if (TypeField = 'BLOB') or (TypeField = 'DATA') then Result := 'M'
  else if (TypeField = 'BOOLEAN') then Result := 'B'
  else Result := 'C';
end;

function wGetSimpleTypeField(const FieldName: String): Char;
begin
  Result := GetSimpleTypeField(wGetTypeField(FieldName))
end;

function wGetSimpleTypeField(const FieldName: String; const iTable, iField: Integer): Char;
var Mcd : IMCDServiceCOM;
		TypeField : string;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  TypeField := (Mcd.GetField(FieldName) as IFieldCOM).tipe ;

  Result := GetSimpleTypeField(TypeField);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : Permet de ramener une valeur par défaut en fonction d'un 
Suite ........ : type de champ
Mots clefs ... : 
*****************************************************************}
function wGetInitValue(Const TypeField: Char): Variant;
begin
  case TypeField of
    'C', 'M': Result := '';
    'D'     : Result := VarAsType(iDate1900, varDate);
    'N'     : Result := VarAsType(0, varDouble);
    'I'     : Result := VarAsType(0, varInteger);
    'B'     : Result := wFalse;
  else
    Result := null
  end
end;

function wGetInitValue(const FieldName: String): Variant;
begin
  Result := wGetInitValue(wGetSimpleTypeField(FieldName))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/06/2002
Modifié le ... :   /  /
Description .. : Renvoie le type d'un champ
Mots clefs ... :
*****************************************************************}
function wGetTypeField(Const FieldName: String): String;
var
  iTable, iChamp: Integer;
	Mcd : IMCDServiceCOM;
begin
  Result := '';
  if (Pos('MIN(', UpperCase(FieldName)) > 0)
  or (Pos('MAX(', UpperCase(FieldName)) > 0)
  or (Pos('SUM(', UpperCase(FieldName)) > 0) then
    Result := wGetTypeField(Trim(Copy(FieldName, 5, Length(FieldName) - 5)))
  else
  begin
    MCD := TMCD.GetMcd;
    if not mcd.loaded then mcd.WaitLoaded();
    result := (Mcd.GetField(FieldName) as IFieldCOM).tipe ;
  end;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 04/07/2002
Modifié le ... :   /  /
Description .. : Se charge d'appeler les propriétés d'un enregistrement
Suite ........ : PtrEcran = String(@Ecran) pour récupérer l'écran dans la WPROPERTIES_TOF
Mots clefs ... :
*****************************************************************}
procedure wCallProperties(Const Prefixe, Identifiant, Clef: String; Const PointeurFiche: String = '');
var
	Argument: string;
begin
	Argument := 'PREFIXE=' + Prefixe
            + ';IDENTIFIANT=' + Identifiant
            + ';CLEF=' + Clef
            + ';PTRFICHE=' + PointeurFiche
            ;
	AglLanceFiche('W', 'WPROPERTIES', '', '', Argument);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

function HSLtoRGB(H, S, L: Double): TColor;
var
  M1, M2: Double;

  function HueToColourValue(Hue: Double): Byte;
  var
    V: Double;
  begin
    if Hue < 0 then
      Hue := Hue + 1
    else if Hue > 1 then
      Hue := Hue - 1;

    if 6 * Hue < 1 then
      V := M1 + (M2 - M1) * Hue * 6
    else if 2 * Hue < 1 then
      V := M2
    else if 3 * Hue < 2 then
      V := M1 + (M2 - M1) * (2/3 - Hue) * 6
    else
      V := M1;
    Result := Min(Round(255 * V), 255)
  end;

var
  R, G, B: Byte;
begin
  L := Min(1, L);

  if S = 0 then
  begin
    R := Round(255 * L);
    G := R;
    B := R
  end
  else
  begin
    if L <= 0.5 then
      M2 := L * (1 + S)
    else
      M2 := L + S - L * S;
    M1 := 2 * L - M2;
    R := HueToColourValue (H + 1/3);
    G := HueToColourValue (H);
    B := HueToColourValue (H - 1/3)
  end;

  Result := RGB (R, G, B)
end;

procedure RGBtoHSL(RGB: TColor; var H, S, L : Double);
var
  R, G, B, D, Cmax, Cmin: Double;
begin
  RGB := ColorToRGB(RGB);
  R := GetRValue(RGB) / 255;
  G := GetGValue(RGB) / 255;
  B := GetBValue(RGB) / 255;
  Cmax := Max(R, Max(G, B));
  Cmin := Min(R, Min(G, B));

  // Luminosité
  L := (Cmax + Cmin) / 2;

  if Cmax = Cmin then  // Niveaux de gris
  begin
    H := 0; // Non définis
    S := 0
  end
  else
  begin
    D := Cmax - Cmin;

    // Saturation
    if L = 0 then
      S := 0
    else if L < 0.5 then
      S := D / (Cmax + Cmin)
    else
      S := D / (2 - Cmax - Cmin);

    // Teinte
    if R = Cmax then
      H := (G - B) / D
    else if G = Cmax then
      H  := 2 + (B - R) /D
    else
      H := 4 + (R - G) / D;

    H := H / 6;
    if H < 0 then
      H := H + 1
  end
end;

{ Si intens>0 alors effet plus lumineux
  Si intens<0 alors effet plus sombre }
function HighlightColor(const Color: TColor; Intens: ShortInt): TColor;
var
  H, S, L: Double;
  Factor: Shortint;
begin
  if Intens = 0 then
    Result := Color
  else
  begin
    Intens := Min(100, Intens);
    Intens := Max(-100, Intens);
    Factor := Round(Intens / Abs(Intens));
    Intens := Abs(Intens);
    RGBToHSL(Color, H, S, L);
    L := Min(L, 1) + Factor * (Min(L, 1) * (Intens / 100));
    Result := HSLToRGB(H, S, L)
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : Permet d'assombrir une couleur
Description .. : Ainsi, on est lié à la couleur définie par l'utilisateur dans windows
Mots clefs ... :
*****************************************************************}
function wAssombrirCouleur(CouleurDOrigine: TColor; Pourcent: Integer): tColor;
begin
  Result := HighlightColor(CouleurDOrigine, -Pourcent)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 19/07/2002
Modifié le ... :   /  /
Description .. : Permet d'éclaircir une couleur
Description .. : Ainsi, on est lié à la couleur définie par l'utilisateur dans windows
Mots clefs ... :
*****************************************************************}
function wEclaircirCouleur(CouleurDOrigine: TColor; Pourcent: Integer): tColor;
begin
  Result := HighlightColor(CouleurDOrigine, Pourcent)
end;


function wInvertColor(Couleur: TColor): tColor;
var
  R, G, B   : Integer;
  CouleurRGB: LongInt;
begin
  CouleurRGB := ColorToRGB(Couleur);
  R := GetRValue(CouleurRGB);
  G := GetGValue(CouleurRGB);
  B := GetBValue(CouleurRGB);
  R := ABS(R - 255);
  G := ABS(G - 255);
  B := ABS(B - 255);
  Result := RGB(R, G, B);
end;

procedure DrawTitlePanel(const P: THPanel);
begin
	{$IFNDEF EAGLSERVER}
    if Assigned(P) then
    begin
      P.ParentBackGround := False;
      if V_PGI.Draw2003 and V_Pgi.DrawXP then
      begin
        P.BackGroundEffect := bdDown;
        P.ColorEnd := FTitreGEndColor;
        P.ColorStart := FTitreGStartColor;
      end
      else
        P.Color := clBtnShadow
    end
	{$ENDIF !EAGLSERVER}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 26/07/2002
Modifié le ... :   /  /
Description .. : Permet de savoir si une valeur est comprise entre X & Y
Mots clefs ... :
*****************************************************************}
function wBetween(Const LaValeur, PetiteValeur, GrandeValeur: Variant; Const EgalitePetiteValeur: boolean=true; Const EgaliteGrandeValeur: Boolean=true): Boolean;
begin
	if EgalitePetiteValeur and not EgaliteGrandeValeur then           Result := (LaValeur >= PetiteValeur) and (LaValeur < GrandeValeur)
  else if not EgalitePetiteValeur and EgaliteGrandeValeur then      Result := (LaValeur > PetiteValeur) and (LaValeur <= GrandeValeur)
  else if not EgalitePetiteValeur and not EgaliteGrandeValeur then  Result := (LaValeur > PetiteValeur) and (LaValeur < GrandeValeur)
  else if EgalitePetiteValeur and EgaliteGrandeValeur then   	      Result := (LaValeur >= PetiteValeur) and (LaValeur <= GrandeValeur)
  else                                                              Result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 07/08/2002
Modifié le ... : 17/12/2002  Thierry
Description .. : Convertit un temps hh,ss en hh,cc (Heure Seconde en Heure centième)
Ndp : nTemps est convertit au format xxxxxx,mmsspp
*****************************************************************}
function HHCC(Const Temps: Double): Double;
var
  nDec: Extended;
	sDec: String;
	Err, hh, mm, ss, pp: Integer;
begin
  Result := 0;
  Err := 0;
  if (Temps < 999999) then
  begin
    { Extrait la partie entière de nTemps }
    hh := Trunc(Temps);
    { Formate et Extrait la partie décimale de nTemps }
    nDec := Frac(Temps);
    sDec := Format('%8.6f', [nDec]);
    { Extrait les valeurs dans la chaine de la partie décimale formatée }
    mm := ValeurI(Copy(sDec, 3, 2));
    ss := ValeurI(Copy(sDec, 5, 2));
    pp := ValeurI(Copy(sDec, 7, 2));
    if (Temps < 0) or (mm >= 60) or (ss >= 60) then
      Err := 2
    else
      Result := Arrondi( hh + ( mm / 60 ) + ( ss / 3600 ) + ( pp / 360000 ) , 6 );
  end
  else
    Err := 1;

  if Err <> 0 then
  begin
    PgiError(Format(TraduireMemoire('Conversion de temps impossible (wCommuns.HhCc). %s n''est pas un temps correct !'), [FloatToStr(Temps)]));
    Result := 0;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc Morretton
Créé le ...... : 07/08/2002
Modifié le ... : 07/08/2002
Description .. : Convertit un temps hh,cc en hh,ss (Heure centième en Heure seconde)
Mots clefs ... :
*****************************************************************}
function HHMM(Const Temps: Double): Double;
var
  hh, mm, ss: integer;
begin
  { Les heures }
  hh := Trunc(Temps);

  { Les secondes totales}
  ss := round((Temps - hh)*3600);

  { Les minutes }
  mm := Max(0, Min(59, trunc(wDivise(ss, 60))));

  { Les secondes }
  ss := Max(0, Min(59, round(ss - mm*60)));

  { Résultat }
  Result := hh + mm/100 + ss/10000;
end;

//celui qui le fait nativement ne marche pas ...
function wStringToAction(Const Action: String): TActionFiche;
begin
	if Action = 'CONSULTATION' then Result := taConsult
  else if Action = 'CREATION' then Result := taCreat
  else Result := taModif;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/10/2002
Description .. : Retourne Vrai si le DataType est lié à un article
Suite ........ : (utilisé par le SetElipsisClick des articles de la wTom)
*****************************************************************}
function wIsDataTypeArticle(Const TheDataType: String): Boolean;
begin
  //GP_20071214_MM_GP14669
	Result := (TheDataType = 'GCARTICLE') or (TheDataType = 'WARTICLE') or (TheDataType='WCODEARTICLEWNT') or (TheDataType='WCODEARTICLEWGT');
// Ndp : Voir pour faire kkchoze de + générique, en lisant les suffixes de clé étrangére des tablettes par exemple...
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 18/10/2002
Description .. : Retourne Vrai si le DataType est lié à un CodeArticle
Suite ........ : (utilisé par le SetElipsisClick des articles de la wTom)
*****************************************************************}
function wIsDataTypeCodeArticle(Const TheDataType: String): Boolean;
begin
  Result := (TheDataType = 'GCARTICLEGENERIQUE') or (TheDataType = 'WCODEARTICLE');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 01/06/2006
Description .. : Retourne Vrai si le DataType est lié à un Tiers
Suite ........ : (utilisé par le SetElipsisClick des tiers de la wTom/wTof)
*****************************************************************}
function wIsDataTypeTiers(Const TheDataType: String): Boolean;
begin
  Result := (TheDataType = 'GCTIERS') or (TheDataType = 'GCTIERSCLI')
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Eric CHEVALLIER
Créé le ...... : 13/02/2003
Modifié le ... :   /  /
Description .. : Copie, d'une tob dans une autre, les champs ayant le même
Suite ........ : suffixe
Mots clefs ... :
*****************************************************************}
procedure wCopyTobBySuffixe(Const TobSource, TobDest: Tob; Const PrefixeSource, PrefixeDest: String; Const WithKey: Boolean = True; Const AddToNomChamp: String = '');
var
  n: Integer;
	chDest, chSrc, Clef: String;
  iTable, iChDest: Integer;

  procedure FillTobByTypeField(const TDest: Tob; const FieldName: String; const Value: Variant; const TypeField: Char);
  begin
    case TypeField of
      'D': TDest.SetDateTime(FieldName, VarAsType(Value, varDate));
      'N': TDest.SetDouble  (FieldName, VarAsType(Value, varDouble));
      'I': TDest.SetInteger (FieldName, VarAsType(Value, varInteger));
    else
      TDest.SetString(FieldName, VarToStr(Value)) //on traite aussi les booléens en STRING
    end
  end;

begin
  if Assigned(TobSource) and Assigned(TobDest) then
  begin
    if not WithKey then
      Clef := WMakeFieldString(PrefixeToTable(PrefixeDest), ';') + ';'
    else
      Clef := '';

    iTable := TobDest.NumTable;
    if (iTable = -1) and (PrefixeDest <> '') then
      iTable := TableToNum(PrefixeToTable(PrefixeDest));

    for n := 1 to TobSource.NbChamps do
    begin
      chSrc := tobSource.GetNomChamp(n);
      ChDest := wGetSuffixe(chSrc);

      if (PrefixeDest <> '') and (copy(chDest, 1, 1) <> '_') then
        ChDest := PrefixeDest + '_' + ChDest + AddToNomChamp;

      if ((PrefixeSource = '') or (wGetPrefixe(chSrc) = PrefixeSource)) and (WithKey or (not WithKey and (Pos(iif(ChDest = '', '@', ChDest) + ';', Clef) = 0))) Then
      begin
        if not TobDest.FieldExists(chDest) then
          TobDest.AddChampSupValeur(chDest, TobSource.GetValeur(n))
        else 
        begin
          iChDest := ChampToNumDicho(iTable, chDest);
          if iChDest >= 0 then
            FillTobByTypeField(TobDest, chDest, TobSource.GetValeur(n), wGetSimpleTypeField(chDest, iTable, iChDest))
          else
            TobDest.PutValue(chDest, VarAsType(TobSource.GetValeur(n), VarType(TobSource.GetValeur(n))))
        end
      end;
    end;

    //Champ supp.
    for n := 1000 to (1000 + TobSource.ChampsSup.Count - 1) do
    begin
      ChDest := wGetSuffixe(tobSource.GetNomChamp(n));

      if (PrefixeDest <> '') and (copy(chDest, 1, 1) <> '_') then
        ChDest := PrefixeDest + '_' + ChDest + AddToNomChamp;

      if ((PrefixeSource = '') or (wGetPrefixe(tobSource.GetNomChamp(n)) = PrefixeSource)) and (WithKey or (not WithKey and (Pos(iif(ChDest = '', '@', ChDest) + ';', Clef) = 0))) Then
      begin
        if TobDest.FieldExists(ChDest) then
        begin
          iChDest := ChampToNumDicho(iTable, chDest);
          if iChDest >= 0 then
            FillTobByTypeField(TobDest, chDest, TobSource.GetValeur(n), wGetSimpleTypeField(chDest, iTable, iChDest))
          else
            TobDest.PutValue(ChDest, VarAsType(TobSource.GetValeur(n), VarType(TobSource.GetValeur(n))))
        end
        else if AddToNomChamp <> '' then
          TobDest.AddChampSupValeur(ChDest, TobSource.GetValeur(n));
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 15/02/2005
Modifié le ... :   /  /    
Description .. : Ramène le nombre de tob filles (récursivement) de la tob 
Suite ........ : "T"
Mots clefs ... : TOB;COUNT
*****************************************************************}
function wTobCountDetails(T: Tob): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Pred(T.Detail.Count) do
    Result := Result + 1 + wTobCountDetails(T.Detail[i])
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 10/03/2006
Modifié le ... :   /  /    
Description .. : Converti un variant en Tob
Mots clefs ... : 
*****************************************************************}
function VarAsTob(T: Variant): Tob;
begin
  Result := Tob(LongInt(T))
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/12/2005
Modifié le ... :   /  /    
Description .. : Ramène l'index absolut de la tob passée en paramètres.
Mots clefs ... : 
*****************************************************************}
function wTobGetAbsoluteIndex(T: Tob): Integer;

  function CountSibilingDetails(TP: Tob; const Index: Integer): Integer;
  var
    i: Integer;
  begin
    Result := 0;
    if Assigned(TP) then
      for i := 0 to Pred(Index) do
        Result := Result + wTobCountDetails(TP.Detail[i])
  end;

begin
  if Assigned(T) then
    Result := (T.GetIndex + 1) + CountSibilingDetails(T.Parent, T.GetIndex) + wTobGetAbsoluteIndex(T.Parent)
  else
    Result := 0
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 03/08/2005
Modifié le ... :   /  /    
Description .. : Echange les deux filles en fonction des index donnés
Description .. : Renvoie Vrai si l'échange a été fait
Mots clefs ... :
*****************************************************************}
function TobExchange(TheTob: Tob; Const index1, index2: integer): boolean;
var
  tob1, tob2: Tob;
begin
  Result := false;
  if not Assigned(TheTob) then exit;
  if not Assigned(TheTob.Detail[index1]) then exit;
  if not Assigned(TheTob.Detail[index2]) then exit;

  tob1 := TheTob.Detail[index1];
  tob2 := TheTob.Detail[index2];
  tob1.ChangeParent(TheTob, index2);
  tob2.ChangeParent(TheTob, index1);
  Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Remplissage d'une Tob depuis un argument type chaîne :
Suite ........ : MYFIELD1=MYVALUE1;MYFIELD2=MYVALUE2;...
Mots clefs ... : 
*****************************************************************}
procedure LoadTobFromArgument(const TobData: Tob; Argument: String; const Separateur: String = ';');
var
  Token, FieldName: String;
begin
  while Argument <> '' do
  begin
    Token := Trim(ReadTokenPipe(Argument, Separateur));
    FieldName := GetTokenFieldName(Token);
    TobData.AddChampSupValeur(FieldName, GetArgumentValue(Token, FieldName, False, Separateur));
  end;
end;

//GP_20071217_TS_GP14539 >>>
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 12/12/2007
Modifié le ... :   /  /    
Description .. : Rempli un tableau de variants depuis les valeur d'une Tob 
Suite ........ : pour les champs donnés
Mots clefs ... : 
*****************************************************************}
function TobToArrayValue(const TobData: Tob; const FieldNames: MyArrayField): MyArrayValue;
var
  i: Integer;
begin
  SetLength(Result, Length(FieldNames));
  for i := Low(FieldNames) to High(FieldNames) do
    Result[i] := TobData.GetValue(FieldNames[i])
end;
//GP_20071217_TS_GP14539 <<<

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Composition d'un argument type chaîne : 
Suite ........ : MYFIELD1=MYVALUE1;MYFIELD2=MYVALUE2;...
Mots clefs ... : 
*****************************************************************}
function GetArgumentByTob(TobData: Tob; WriteKey: Boolean = True; WriteBlocNote: Boolean = False; SansPrefixe: Boolean = False): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to TobData.NbChamps do
    if  (WriteKey      or not wIsFieldFromKey1(TobData.GetNomChamp(i)))
    and (WriteBlocNote or (wGetSimpleTypeField(TobData.GetNomChamp(i)) <> 'M'))  then
      Result := Result + copy(TobData.GetNomChamp(i), pos('_', TobData.GetNomChamp(i)), 18) + '=' + VarToStr(TobData.GetValeur(i)) + ';';

  for i := 1000 to 1000+Pred(TobData.ChampsSup.Count) do
    Result := Result + copy(TobData.GetNomChamp(i), pos('_', TobData.GetNomChamp(i))+1, 18) + '=' + VarToStr(TobData.GetValeur(i)) + ';'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Update d'un seule Tob dans une hiérarchie de Tob.
Suite ........ : Un "UpdateDb" étant récursif sur les filles, cette fonction
Suite ........ : permet d'effectuer l'update d'une Tob indépendemment de sa
Suite ........ : structure hiérarchique
Mots clefs ... :
*****************************************************************}
function wUpdateSingleTob(T: Tob): Boolean;
var
  TobTable: Tob;
begin
  TobTable := Tob.Create(T.NomTable, nil, -1);
  try
    TobTable.Dupliquer(T, False, True, True);
    TobTable.ChargeCle1();
    Result := TobTable.UpdateDB()
  finally
    TobTable.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Insert d'un seule Tob dans une hiérarchie de Tob.
Suite ........ : Un "InsertDb" étant récursif sur les filles, cette fonction
Suite ........ : permet d'effectuer l'Insert d'une Tob indépendemment de sa
Suite ........ : structure hiérarchique
Mots clefs ... :
*****************************************************************}
function wInsertSingleTob(T: Tob): Boolean;
var
  TobTable: Tob;
begin
  TobTable := Tob.Create(T.NomTable, nil, -1);
  try
    TobTable.Dupliquer(T, False, True, True);
    TobTable.ChargeCle1();
    Result := TobTable.InsertDB(nil)
  finally
    TobTable.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Delete d'un seule Tob dans une hiérarchie de Tob.
Suite ........ : Un "DeleteDb" étant récursif sur les filles, cette fonction
Suite ........ : permet d'effectuer le Delete d'une Tob indépendemment de sa
Suite ........ : structure hiérarchique
Mots clefs ... :
*****************************************************************}
function wDeleteSingleTob(T: Tob): Boolean;
var
  TobTable: Tob;
begin
  TobTable := Tob.Create(T.NomTable, nil, -1);
  try
    TobTable.Dupliquer(T, False, True, True);
    TobTable.ChargeCle1();
    Result := TobTable.DeleteDB();
    if not Result then
      Result := not TobTable.ExistDB()
  finally
    TobTable.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Variante de wInsertSingleTob & wUpdateSingleTob
Mots clefs ... :
*****************************************************************}
function wInsertOrUpdateSingleTob(T: Tob): Boolean;
var
  TobTable: Tob;
begin
  TobTable := Tob.Create(T.NomTable, nil, -1);
  try
    TobTable.Dupliquer(T, False, True, True);
    TobTable.ChargeCle1();
    Result := TobTable.InsertOrUpdateDB()
  finally
    TobTable.Free
  end
end;

{**********************************************************************************************************************}
{*                                        TS : FONCTIONS SUR LES DATES                                                *}
{**********************************************************************************************************************}

{ ramène l'année d'une date }
function wYear(Const LaDate: tDateTime): word;
var
  y, m, d: word;
begin
  Decodedate(LaDate, y, m, d);
  Result := y;
end;

{ ramène le mois d'une date }
function wMonth(Const LaDate: tDateTime): word;
var
  y, m, d: word;
begin
  Decodedate(LaDate, y, m, d);
  Result := m;
end;

{ ramène le jour d'une date }
function wDay(Const LaDate: tDateTime): word;
var
  y, m, d: word;
begin
  Decodedate(LaDate, y, m, d);
  Result := d;
end;
{ ramène le quantième d'une date }
function wGetNumDay(Const LaDate: tDateTime): Integer;
begin
  Result := trunc(LaDate - EncodeDate(wYear(LaDate), 1, 1)) + 1;
end;

{ ramène l'heure }
function wHour(Const Heure: tTime): word;
var
  hour, min, sec, msec: word;
begin
  Decodetime(Heure, hour, min, sec, msec);
  Result := hour;
end;

{ ramène les minutes }
function wMin(Const Heure: tTime): word;
var
  hour, min, sec, msec: word;
begin
  Decodetime(Heure, hour, min, sec, msec);
  Result := min;
end;

{ ramène les secondes }
function wSec(Const Heure: tTime): word;
var
  hour, min, sec, msec: word;
begin
  Decodetime(Heure, hour, min, sec, msec);
  Result := sec;
end;

{ ramène les Millisecondes }
function wMSec(Const Heure: tTime): word;
var
  hour, min, sec, msec: word;
begin
  Decodetime(Heure, hour, min, sec, msec);
  Result := msec;
end;

function IsDateVide(const sDate: String): Boolean;
begin
  Result := False;
  try
    StrToDate(sDate)
  except
    on E: EConvertError do
      Result := True
  end
end;

{ ramène le nom du mois avec ou sans majuscule }
function wGetMonth(Const LeMois: Integer; Const WithMajuscule: Boolean = False): String;
begin
  case LeMois of
    1 : Result := TraduireMemoire('janvier');
    2 : Result := TraduireMemoire('février');
    3 : Result := TraduireMemoire('mars');
    4 : Result := TraduireMemoire('avril');
    5 : Result := TraduireMemoire('mai');
    6 : Result := TraduireMemoire('juin');
    7 : Result := TraduireMemoire('juillet');
    8 : Result := TraduireMemoire('août');
    9 : Result := TraduireMemoire('septembre');
    10: Result := TraduireMemoire('octobre');
    11: Result := TraduireMemoire('novembre');
    12: Result := TraduireMemoire('décembre');
    else Result := '';
  end;
  if WithMajuscule and (Result <> '') then
  	Result[1] := UpCase(Result[1]);
end;

{ récupère le n° de la dernière semaine de l'année de la date ... hem j'me comprends }
function wGetLastWeek(Const LaDate: tDateTime): Word;
var
	j, m, a: Word;
begin
	DecodeDate(LaDate, a, m, j);
  j := 31;
  m := 12;
  Result := NumSemaine(EncodeDate(a, m, j));
  while Result = 1 do
  begin
  	Dec(j);
		Result := NumSemaine(EncodeDate(a, m, j));
  end;
end;

{ Ramène le numéro de trimestre }
function wGetTrimestreFromMois(Const LeMois: Integer): Integer;
begin
  case LeMois of
    1..3  : Result := 1;
    4..6  : Result := 2;
    7..9  : Result := 3;
    10..12: Result := 4;
    else    Result := 0;
  end;
end;

{ Ramène le 1er mois du trimestre }
function wGet1erMoisFromTrimestre(Const LeTrimestre: Integer): Integer;
begin
  case LeTrimestre of
    1 : Result :=  1;
    2 : Result :=  4;
    3 : Result :=  7;
    4 : Result := 10;
  else  Result :=  0;
  end;
end;

{ Ramène le numéro de trimestre }
function wGetTrimestre(Const LaDate: tDateTime): Integer;
begin
  Result := wGetTrimestreFromMois(wMonth(LaDate));
end;

{ Ramène le numéro de semestre }
function wGetSemestreFromMois(Const LeMois: Integer): Integer;
begin
  case LeMois of
    1..6  : Result := 1;
    7..12 : Result := 2;
    else    Result := 0;
  end;
end;

{ Ramène le 1er mois d'un semestre }
function wGet1erMoisFromSemestre(Const LeSemestre: Integer): Integer;
begin
  case LeSemestre of
    1 : Result := 1;
    2 : Result := 7;
  else  Result := 0;
  end;
end;
{ Ramène le numéro de semestre }
function wGetSemestre(Const LaDate: tDateTime): Integer;
begin
  Result := wGetSemestreFromMois(wMonth(LaDate));
end;
{**********************************************************************************************************************}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Florence DURANTET
Créé le ...... : 25/10/2002
Modifié le ... :   /  /
Description .. : Permet d'afficher le GroupBox GBUNIC (Quantité/Durée unique
Suite ........ : ou Ressource unique) selon le mode de gestion de la
Suite ........ : gamme
Mots clefs ... : GBUNIC
*****************************************************************}
procedure uAfficheGBUNIC (Const GbUnic: tGroupBox; Const Affiche:boolean); // Affichage du GroupBox GBUNIC
begin
  GbUnic.Visible := Affiche;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 17/02/2003
Modifié le ... :   /  /
Description .. : Permet de restreindre la composition d'un where lors d'une
Suite ........ : multi-sélection d'un mul.
Suite ........ : Gère également un rapport personnalisé grâce à la function
Suite ........ : "IsAValidRecord" de type TFuncWRP passée en paramètre
Suite ........ : **********************************************************************
Suite ........ : prototype de la fonction :
Suite ........ : function MyFunc(Q_ETR: TQuery): String;
Suite ........ : Cette fonction doit renvoyer une chaîne de la forme :
Suite ........ : '[ERROR];MESSAGE=Tout c''est pas bien passé'
Suite ........ : '[ERROR];' => permet de repérer l'erreur rencontrée et ainsi de
Suite ........ : filtrer la requête.
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{$IFNDEF EAGLCLIENT}
function wMakeRestrictedWhereFromList(Const Prefixe: String; Const Liste: ThDbGrid; Const Q: tQuery; Const IsAValidRecord: TFuncWRP; Const WithRapport: Boolean = false; Const Rapport: TWRapport = nil): string;
{$ELSE  !EAGLCLIENT}
function wMakeRestrictedWhereFromList(Const Prefixe: String; Const Liste: ThGrid;   Const Q: tQuery; Const IsAValidRecord: TFuncWRP; Const WithRapport: Boolean = false; Const Rapport: TWRapport = nil): string;
{$ENDIF !EAGLCLIENT}
var
	i: integer;
  {$IFDEF EAGLCLIENT}
  wBookMark: integer; {PMJEAGL}
  {$ELSE}
  wBookMark: tBookMark;
  {$ENDIF}
  sMsg: String;

	procedure AddToWhere(var Where: string);
  var
    sClef1, sFieldsValues: String;

    function GetFieldsValues(sClef: String): String;
    begin
      Result := '';
      while sClef <> '' do
        Result := Result + iif(Result = '', '', ';') + Q.findField(ReadTokenSt(sClef)).AsString;
    end;

  begin
    sClef1 := wMakeFieldString(PrefixeToTable(Prefixe), ';');
    if not Q.Eof then
    begin
      sFieldsValues := GetFieldsValues(sClef1);
      sMsg := IsAValidRecord(Q);
      if Pos('[ERROR]', sMsg) = 0 then
      begin
        Where := Where + iif(Where <> '', ' OR ', '') + '(' + wMakeWhereSQL(sClef1, sFieldsValues) + ')';//+ Prefixe + '_IDENTIFIANT' + SuffixeCompl + '=' + Q.FindField(Prefixe + '_IDENTIFIANT' + SuffixeCompl).AsString;
        { Rapport }
        if WithRapport and (Assigned(Rapport)) and (GetArgumentValue(sMsg, 'MESSAGE', False) <> '') then
          Rapport.Add(TWRP_Done, sFieldsValues, GetArgumentValue(sMsg, 'MESSAGE', False));
      end
      else
      begin
        { Rapport }
        if WithRapport and (Assigned(Rapport)) then
          Rapport.Add(TWRP_Error, sFieldsValues, GetArgumentValue(sMsg, 'MESSAGE', False));
      end;
    end;
  end;

begin
  Result := '';
	if (Liste.NbSelected > 0) or (Liste.AllSelected) then
  begin
   	Q.DisableControls;
    try
      {$IFDEF EAGLCLIENT}
      wBookMark := Liste.Row; {PMJEAGL}
      if Liste.AllSelected then
      Begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          AddToWhere(Result);
          Q.next;
        end;
      end
      else
      Begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          if Liste.IsSelected( i ) then
          begin
            AddToWhere(Result);
          end;
          Q.next;
        end;
      end;
      Liste.Row := wBookMark; {PMJEAGL}
      if Liste.Row>0 then
        Q.Seek( Liste.Row - 1 )
      else
        Q.First;
      {$ELSE}
      wBookMark := Liste.DataSource.DataSet.GetBookmark;
      if Liste.AllSelected then
      Begin
        Q.First;
        for i := 1 to Q.recordCount do
        begin
          AddToWhere(Result);
          Q.next;
        end;
      end
      else
      Begin
        for i := 0 to Liste.NbSelected - 1 do
        Begin
          Liste.GotoLeBOOKMARK(i);
          AddToWhere(Result);
        end;
      end;
      Q.GotoBookmark(wBookMark);
      {$ENDIF}
    finally
      Q.EnableControls;
    end;
  end
  else
  begin
	  AddToWhere(Result);
  end;

  { Rapport }
  if Result = '' then
    Result := '1 = 2';
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. : Dit si on a le droit de modifier la fiche correspondante
Mots clefs ... :
*****************************************************************}
function JAiLeDroitGestion(Const Prefixe: string): boolean;
var iTag: integer;
begin
  if (CtxAffaire in V_PGI.PgiContexte) then
  begin
    if      Prefixe = 'GA'      then Result := JAiLeDroitTag( 71511)
    else if Prefixe = 'TAR'     then Result := JAiLeDroitTag( 74346)
    else if Prefixe = 'FN1'     then Result := JAiLeDroitTag( 74341)
    else if Prefixe = 'FN2'     then Result := JAiLeDroitTag( 74342)
    else if Prefixe = 'FN3'     then Result := JAiLeDroitTag( 74343)
    else if Prefixe = 'YTP'     then Result := JAiLeDroitTag(215102)
    else if Prefixe = 'ARS'     then Result := JAiLeDroitTag( 71104)
    else if Prefixe = 'T'       then Result := JAiLeDroitTag( 71101)
    else if Prefixe = 'TRC'     then Result := JAiLeDroitTag( 74315)
    else if Prefixe = 'TRF'     then Result := JAiLeDroitTag( 74318)
    else if Prefixe = 'GCA'     then Result := JAiLeDroitTag(138602)
    {$IFDEF GCGC}
    else if copy(Prefixe,1,3)='YTS' then  //GA_AB_200801-GA14539
    begin
      itag := 0;
      {Calcul du Tag en fonction du module}
      {
      if      (copy(Prefixe,5,3)=sTarifClient) then
        iTag := 210400
      else
      if (copy(Prefixe,5,3)=sTarifFournisseur) then
        iTag := 211400
      ;
      }
      {Calcul du Tag en fonction : Saisie ou Consultation}
      if      (copy(Prefixe,9,3)='SAI') then
        iTag := iTag+  10
      else if (copy(Prefixe,5,3)='CON') then
        iTag := iTag+  20
      ;
      {Calcul du Tag en fonction de l'orientation demandé}
      if      (copy(Prefixe,13,3)='TIE') then
        iTag := iTag+   1
      else if (copy(Prefixe,13,3)='ART') then
        iTag := iTag+   2
      ;
      Result := JAiLeDroitTag(iTag);
    end
    {$ENDIF GCGC}    
    else
    begin
      Result := false;
      if V_Pgi.sav then PgiError('Appel de la fonction wCommuns.JAiLeDroitGestion avec le préfixe: ' + Prefixe + ' non implémentée')
    end;
  end
  else
  begin
    if      Prefixe = 'WNT' then Result := JAiLeDroitTag(126120)
    else if Prefixe = 'WGT' then Result := JAiLeDroitTag(126130)
//GP_20071015_MM_GP14362
    else if Prefixe = 'WPE' then Result := JAiLeDroitTag(126420)
    else if Prefixe = 'WPC' then Result := JAiLeDroitTag(125110)
    else if Prefixe = 'WOT' then Result := JAiLeDroitTag(120110)
    else if Prefixe = 'WOL' then Result := JAiLeDroitTag(120120)
    else if Prefixe = 'WOP' then Result := JAiLeDroitTag(120130)
    else if Prefixe = 'WOB' then Result := JAiLeDroitTag(120140)
    else if Prefixe = 'WLB' then Result := JAiLeDroitTag(128120)
    else if Prefixe = 'WRB' then Result := JAiLeDroitTag(128130)
    else if Prefixe = 'QWB' then Result := JAiLeDroitTag(352110)
    else if Prefixe = 'QIT' then Result := JAiLeDroitTag(121110)
    else if Prefixe = 'QCI' then Result := JAiLeDroitTag(121120)
    else if Prefixe = 'GA'  then Result := JAiLeDroitTag(325011)
    else if Prefixe = 'GCA' then Result := JAiLeDroitTag(31702)
    else if Prefixe = 'TAR' then Result := JAiLeDroitTag( 65302)
    else if Prefixe = 'FN1' then Result := JAiLeDroitTag( 65321)
    else if Prefixe = 'FN2' then Result := JAiLeDroitTag( 65322)
    else if Prefixe = 'FN3' then Result := JAiLeDroitTag( 65323)
    else if Prefixe = 'GM'  then Result := JAiLeDroitTag(211901)
    else if Prefixe = 'GCQ' then Result := JAiLeDroitTag(211911)
    else if Prefixe = 'FVS' then Result := JAiLeDroitTag(212301)
    else if Prefixe = 'YTP' then Result := JAiLeDroitTag(215102)
    else if Prefixe = 'ARS' then Result := JAiLeDroitTag(126314)
    else if Prefixe = 'T'   then Result := JAiLeDroitTag( 30501)
    else if Prefixe = 'TRC' then Result := JAiLeDroitTag( 65108)
    else if Prefixe = 'TRF' then Result := JAiLeDroitTag( 65109)
    else if Prefixe = 'WWF' then Result := JAiLeDroitTag(215205)
    else if Prefixe = 'RQN' then Result := JAiLeDroitTag( 124101)
    else if Prefixe = 'RQD' then Result := JAiLeDroitTag( 124102)
    else if Prefixe = 'RQP' then Result := JAiLeDroitTag( 124103)
    else if Prefixe = 'RQT' then Result := JAiLeDroitTag( 124742)
    else if Prefixe = 'RAC' then Result := JAiLeDroitTag( 92104)
    else if Prefixe = 'SCC' then Result := JAiLeDroitTag( 65106)
    else if Prefixe = 'GOR' then Result := JAiLeDroitTag( 65107)
    else if Prefixe = 'REN' then Result := JAiLeDroitTag( 65110)
    {$IFDEF GCGC}
    else if copy(Prefixe,1,3)='YTS' then
    begin
      itag := 0;
      {Calcul du Tag en fonction du module}
      {
      if      (copy(Prefixe,5,3)=sTarifClient) then
        iTag := 210400
      else
      if (copy(Prefixe,5,3)=sPrixVenteMarge) then
        iTag := 210500
      else if (copy(Prefixe,5,3)=sTarifFournisseur) then
        iTag := 211400
      else if (copy(Prefixe,5,3)=sTarifSousTraitantAchat) then
        iTag := 211500
      else if (copy(Prefixe,5,3)=sTarifSousTraitantPhase) then
        iTag := 211700
      else
      if (copy(Prefixe,5,3)=sFraisAnnexes) then
        iTag := 211800
      else if (copy(Prefixe,5,3)=sCoutsIndirects) then
        iTag := 320000
      ;
      }
      {Calcul du Tag en fonction : Saisie ou Consultation}
      if      (copy(Prefixe,9,3)='SAI') then
        iTag := iTag+  10
      else if (copy(Prefixe,5,3)='CON') then
        iTag := iTag+  20
      ;
      {Calcul du Tag en fonction de l'orientation demandé}
      if      (copy(Prefixe,13,3)='TIE') then
        iTag := iTag+   1
      else if (copy(Prefixe,13,3)='ART') then
        iTag := iTag+   2
      ;

      Result := JAiLeDroitTag(iTag);
    end
    {$ENDIF GCGC}
    else
    begin
      Result := false;
      if V_Pgi.sav then PgiError('Appel de la fonction wCommuns.JAiLeDroitGestion avec le préfixe: ' + Prefixe + ' non implémentée')
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 25/02/2003
Modifié le ... :   /  /
Description .. : Dit si on a le droit de consulter la fiche correspondante
Mots clefs ... :
*****************************************************************}
function JAiLeDroitConsult(Const Prefixe: string): boolean;
begin
  if (CtxAffaire in V_PGI.PgiContexte) then
  begin
    if      Prefixe = 'GA'      then Result := JAiLeDroitTag( 71511)
    else if Prefixe = 'TAR'     then Result := JAiLeDroitTag( 74346)
    else if Prefixe = 'FN1'     then Result := JAiLeDroitTag( 74341)
    else if Prefixe = 'FN2'     then Result := JAiLeDroitTag( 74342)
    else if Prefixe = 'FN3'     then Result := JAiLeDroitTag( 74343)
    else if Prefixe = 'YTP'     then Result := JAiLeDroitTag(215102)
    else if Prefixe = 'ARS'     then Result := JAiLeDroitTag( 71104)
    else if Prefixe = 'T'       then Result := JAiLeDroitTag( 71101)
    else if Prefixe = 'TRC'     then Result := JAiLeDroitTag( 74315)
    else if Prefixe = 'TRF'     then Result := JAiLeDroitTag( 74318)
    else
    begin
      Result := false;
      if V_Pgi.sav then PgiError('Appel de la fonction wCommuns.JAiLeDroitConsult avec le préfixe: ' + Prefixe + ' non implémentée')
    end;
  end
  else
  begin
    if      Prefixe = 'WNT' then Result := JAiLeDroitTag(126220)
    else if Prefixe = 'WGT' then Result := JAiLeDroitTag(126230)
    else if Prefixe = 'WPC'     then Result := JAiLeDroitTag(125210)
    else if Prefixe = 'WOT' then Result := JAiLeDroitTag(120210)
    else if Prefixe = 'WOL' then Result := JAiLeDroitTag(120220)
    else if Prefixe = 'WOP' then Result := JAiLeDroitTag(120230)
    else if Prefixe = 'WOB' then Result := JAiLeDroitTag(120240)
    else if Prefixe = 'WPH' then Result := JAiLeDroitTag(120022)
    else if Prefixe = 'QIT' then Result := JAiLeDroitTag(121210)
    else if Prefixe = 'QWB' then Result := JAiLeDroitTag(352210)
    else if Prefixe = 'QCI' then Result := JAiLeDroitTag(121220)
    else if Prefixe = 'GA'  then Result := JAiLeDroitTag( 32501)
    else if Prefixe = 'GCA' then Result := JAiLeDroitTag( 31702)
    else if Prefixe = 'TAR' then Result := JAiLeDroitTag( 65302)
    else if Prefixe = 'FN1' then Result := JAiLeDroitTag( 65321)
    else if Prefixe = 'FN2' then Result := JAiLeDroitTag( 65322)
    else if Prefixe = 'FN3' then Result := JAiLeDroitTag( 65323)
    else if Prefixe = 'GM'  then Result := JAiLeDroitTag(211901)
    else if Prefixe = 'GCQ' then Result := JAiLeDroitTag(211911)
    else if Prefixe = 'FVS' then Result := JAiLeDroitTag(212301)
    else if Prefixe = 'YTP' then Result := JAiLeDroitTag(215102)
    else if Prefixe = 'ARS' then Result := JAiLeDroitTag(126240)
    else if Prefixe = 'T'   then Result := JAiLeDroitTag( 30501)
    else if Prefixe = 'TRC' then Result := JAiLeDroitTag( 65108)
    else if Prefixe = 'TRF' then Result := JAiLeDroitTag( 65109)
    else if Prefixe = 'SCC' then Result := JAiLeDroitTag( 65106)
    else if Prefixe = 'GOR' then Result := JAiLeDroitTag( 65107)
    else if Prefixe = 'REN' then Result := JAiLeDroitTag( 65110)
    else
    begin
      Result := false;
      if V_Pgi.sav then PgiError('Appel de la fonction wCommuns.JAiLeDroitConsult avec le préfixe: ' + Prefixe + ' non implémentée')
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 30/05/2005
Modifié le ... :   /  /
Description .. : Permet de traduire une donnée de la tablette
Suite ........ : WINITCHAMPOPERCOND, en opérateur SQL
Suite ........ : Ex. : "EGA" -> "=" ; "DIF" -> "<>" ; "COM" -> "LIKE"
Mots clefs ... : 
*****************************************************************}
function wTTOperator2SQLOperator(const TTOperator: String): String;
begin
  if      TTOperator = 'EGA' then Result := '='
  else if TTOperator = 'DIF' then Result := '<>'
  else if TTOperator = 'INF' then Result := '<'
  else if TTOperator = 'SUP' then Result := '>'
  else if TTOperator = 'SEG' then Result := '>='
  else if TTOperator = 'NCO' then Result := 'NOT LIKE'
  else if (TTOperator = 'IEG') or (TTOperator = 'PER') then Result := '<='
  else if (TTOperator = 'COM') or (TTOperator = 'CON') then Result := 'LIKE'
  else Result := ''
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 30/05/2005
Modifié le ... :   /  /    
Description .. : Permet de contruire une condition SQL :
Suite ........ : NomDeChamps Opérateur ValeurCondition
Mots clefs ... : 
*****************************************************************}
function wGetSQLCondition(const FieldName, TTOperator: String; Value: Variant): String;
var
  TypeField: Char;
  WithQuote: Boolean;
  sValue   : String;
begin
  TypeField := wGetSimpleTypeField(FieldName);
  WithQuote := TypeField in ['C', 'D', 'B'];
  case TypeField of
    'D': sValue := UsDateTime(Trunc(Value));
    'I': sValue := IntToStr(Value);
    'N': sValue := StrFPoint(Value);
    'B': sValue := BoolToStr_(Value);
  else
    sValue := VarToStr(Value)
  end;

  Result := FieldName + ' '
          + wTTOperator2SQLOperator(TTOperator) + ' '
          + iif(WithQuote, '"', '') + iif(Pos(TTOperator + ';', 'CON;') > 0, '%', '')
          + sValue
          + iif(Pos(TTOperator + ';', 'CON;NCO;COM;') > 0, '%', '') + iif(WithQuote, '"', '')
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 18/03/2003
Modifié le ... : 18/03/2003
Description .. : Compare 2 valeur en string en entrée dans leur type
Suite ........ : d'origine.
Suite ........ : Valeur1 & 2 : String
Suite ........ : Operateur : Valeur de la tablette WINITOPERCOND
Mots clefs ... :
*****************************************************************}
function wEvalCondition(Const Valeur1, Operateur, Valeur2: string; FieldName: String): Boolean;
var
  Val1, Val2: Variant;
  TypeField: Char;
begin
  TypeField := wGetSimpleTypeField(FieldName);
  case TypeField of
    'N', 'I':
          begin
            Val1 := Valeur(Valeur1);
            Val2 := Valeur(Valeur2);
          end;
    'D':  begin
            Val1 := StrToDate(FormatDateTime('dd/mm/yyyy', StrToDateTime(Valeur1)));
            Val2 := StrToDate(FormatDateTime('dd/mm/yyyy', StrToDateTime(Valeur2)));
          end
    else  begin
            Val1 := Valeur1;
            Val2 := Valeur2;
          end;
  end;

  if Operateur = 'COM' then
    Result := Pos(String(Val2), String(Val1)) = 1
  else if Operateur = 'CON' then
    Result := Pos(String(Val2), String(Val1)) <> 0
  else if Operateur = 'NCO' then
    Result := Pos(String(Val2), String(Val1)) = 0
  else if Operateur = 'EGA' then
  begin
    if TypeField = 'N' then
      Result := EgaliteDouble(Val1, Val2)
    else
      Result := Val1 = Val2
  end
  else if (Operateur = 'DIF') or (Operateur = 'MOD') then
  begin
    if TypeField = 'N' then
      Result := not EgaliteDouble(Val1, Val2)
    else
      Result := Val1 <> Val2
  end
  else if Operateur = 'INF' then
    Result := Val1 < Val2
  else if Operateur = 'IEG' then
    Result := Val1 <= Val2
  else if Operateur = 'SUP' then
    Result := Val1 > Val2
  else if Operateur = 'SEG' then
    Result := Val1 >= Val2
  else
    Result := False
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Permet d'évaluer grâce au script, une expression sous forme
Suite ........ : de STRING écrite en delphi :
Suite ........ : Exemples :
Suite ........ : 1) ((True and True) and False) => un booléen à FALSE
Suite ........ : 2) (5 + 2) * 3 => 21
Suite ........ : ...
Mots clefs ... :
*****************************************************************}
function wEvalDelphiString(sScript: String; Const Params: Array of Variant; var Error: Boolean): Variant; overload;
begin
  result:=wEvalDelphiString(sSCript,'',Params,Error);
end;

function wEvalDelphiString(sScript: string; Const ParamsName: String; Const Params: Array of Variant; var Error: Boolean): Variant;overload;
var
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    FScriptVM: TM3VM;
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
  sParams  : String;
  sNone    : String;
  i        : Integer;
begin
  Error := False;

  sParams := ParamsName;

  if sParams='' then
  begin
    for i := 0 to Length(Params) - 1 do
      sParams := sParams + 'Param_' + IntToStr(i) + iif((i < Length(Params) - 1) and (Length(Params) > 1), ', ', '');

    sNone := Params[0];
    if Pos('NONE', Trim(UpperCase(sNone))) <> 0 then
      sParams := '';
  end;

  sScript := 'function GetEvalExpr(' + sParams + ')'
           + #13'begin'
           + '  Result := ' + sScript
           + #13'end;'
           ;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  FScriptVM := TM3VM.Create;
  try
    FScriptVM.debugMode := M3_DEBUG_RAISE;
    FScriptVM.OpenScript(nil, sScript);
    if FScriptVM.getLastOpenError <> '' then
    begin
      PGIError(FScriptVM.getLastOpenError, 'Erreur de compilation');
      Error := True;
    end;
    Try
      Result := FScriptVM.callFunc(nil, 'GetEvalExpr', Params);
      Error := False;
    except
      on E: Exception do
      begin
        PGIError('Ligne : ' + IntToStr(FScriptVM.getLineFromPC(FScriptVM.RegPC)) + #13#10 + 'Erreur : ' + E.Message, 'Erreur d''exécution' );
        Error := True;
      end;
    end;
  finally
    FScriptVM.CloseScript(nil);
    FScriptVM.Free;
  end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

procedure PutToDebugLog(Const Fonction: string; Const Debut: boolean; Const Notes: string='');

{$IFNDEF EAGLSERVER}
  function GetIndexFonction: integer;
  var
    i: integer;
  begin
    Result := -1;
    for i := Low(RelativeTime) to High(RelativeTime) do
    begin
      if RelativeTime[i].Fonction = Fonction then
      begin
        Result := i;
        Break;
      end;
    end;

    if Result = -1 then
    begin
      SetLength(RelativeTime, Length(RelativeTime)+1);
      RelativeTime[High(RelativeTime)].Fonction := Fonction;
      Result := High(RelativeTime);
    end;
  end;

  procedure RazTime;
  begin
    RelativeTime[GetIndexFonction].Temps := Time;
  end;

  function GetTime: tDateTime;
  begin
    if Debut then
    begin
      RazTime;
      Result := 0;
    end
    else
      Result := Time - RelativeTime[GetIndexFonction].Temps;
  end;

  function GetIoError: string;
  begin
    Case V_Pgi.IoError of
      oeLettrage : Result := 'oeLettrage';
      oeOk       : Result := 'oeOk';
      oePiece    : Result := 'oePiece';
      oePointage : Result := 'oePointage';
      oeReseau   : Result := 'oeReseau';
      oeSaisie   : Result := 'oeSaisie';
      oeStock    : Result := 'oeStock';
      oeTiers    : Result := 'oeTiers';
      oeUnknown  : Result := 'oeUnknown';
    end;
  end;
{$ENDIF !EAGLSERVER}

begin
{$IFNDEF EAGLSERVER}
  if not V_PGI.Debug then exit;

  Trace.TraceInformation('INFO', '===========================================');
  Trace.TraceInformation('INFO', ' Fonction: ' + Fonction);
  if Notes <> '' then
    Trace.TraceInformation('INFO', ' Notes   : ' + Notes);
  Trace.TraceInformation('INFO', ' Temps   : ' + FormatDateTime('hh:nn:ss:zzz', GetTime));
  Trace.TraceInformation('INFO', ' IoError : ' + GetIoError);
  Trace.TraceInformation('INFO', '===========================================');
{$ENDIF !EAGLSERVER}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 15/01/2004
Modifié le ... :   /  /    
Description .. : Count le nombre de tokens dans la chaîne
Mots clefs ... : 
*****************************************************************}
function wCountToken(St: String; const Separateur: String = ';'): Integer;
begin
  Result := 0;
  while St <> '' do
  begin
    ReadTokenPipe(St, Separateur);
    Inc(Result)
  end
end;

function wGetTokenPos(St: String; const SubSt: String; const ExactSubSt: Boolean = True): Integer;
var
  i: Integer;
  St2: String;
begin
  Result := -1;
  i := 1;
  while (St <> '') and (Result = -1) do
  begin
    St2 := UpperCase(ReadTokenSt(St));
    if (ExactSubSt     and (UpperCase(SubSt) = St2))
    or (not ExactSubSt and (Pos(UpperCase(SubSt), St2) = 1)) then
      Result := i;
    Inc(i)
  end
end;

function wGetTokenAt(St: String; const Index: Integer): String;
begin
  if St = '' then
    Result := ''
  else if Index = 1 then
    Result := ReadTokenSt(St)
  else
  begin
    ReadTokenSt(St);
    Result := wGetTokenAt(St, Index - 1)
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/09/2004
Modifié le ... :   /  /    
Description .. : Renvoie le code SQL d'une jointure Clé primaire <=> Clé
Suite ........ : étrangère
Mots clefs ... : 
*****************************************************************}
function wMakeJoinOnCle1(const Table1, Table2: String): String;
var
  Key1, Key2, sField: String;
begin
  Key1 := StringReplace(wGetFieldsClef1(Table1), ',', ';', [rfIgnoreCase, rfReplaceAll]);
  Key2 := StringReplace(wGetFieldsClef1(Table2), ',', ';', [rfIgnoreCase, rfReplaceAll]);

  Result := '';

  while Key1 <> '' do
  begin
    sField := Trim(ReadTokenSt(Key1));
    if Pos('_' + ExtractSuffixe(sField), Key2) <> 0 then
      Result := Result + iif(Result <> '', ' AND ', '') + sField + '=' + TableToPrefixe(Table2) + '_' + ExtractSuffixe(sField)
  end;

  if Result <> '' then
    Result := '(' + Result + ')'
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 08/09/2005
Modifié le ... :   /  /
Description .. : Renvoi le token de longueur max (sert à extraire
Suite ........ : le message d'un V_Pgi.LastHShowMessage
*****************************************************************}
function GetMaxLenToken(St: String): String;
var
  sRes: String;
begin
  Result := '';
  sRes := '';
  while St <> '' do
  begin
    sRes := ReadTokenSt(St);
    if Length(sRes) > Length(Result) then
      Result := sRes
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 02/09/2004
Modifié le ... :   /  /
Description .. : Renvoie la syntaxe SQL d'un select * mais avec une liste
Suite ........ : complète des champs moins ceux présents dans le
Suite ........ : paramètre "FieldsToExclude"
Mots clefs ... :
*****************************************************************}
function wMakeSelectPresqueEtoile(const TableName: String; const FieldsToExclude: MyArrayField = nil; const ExcludeBlobField: Boolean = True): String;
var Mcd : IMCDServiceCOM;
		Table     : ITableCOM ;
    FieldList : IEnumerator ;
    Field     : IFieldCOM ;
  	iTable, i: Integer;

  function IsFieldToExclude(const FieldName: String): Boolean;
  var
    iTab: Integer;
  begin
    Result := Length(FieldsToExclude) > 0;
    if Result then
    begin
      iTab := Low(FieldsToExclude);
      Result := False;
      while not Result and (iTab <= High(FieldsToExclude)) do
      begin
        Result := FieldName = FieldsToExclude[iTab];
        Inc(iTab)
      end
    end
  end;

begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  Result := '';
  Table := Mcd.GetTable(TableName);
  if assigned(table) then
  begin
    FieldList := Table.Fields;
    FieldList.Reset();
    While FieldList.MoveNext do
    begin
      Field := FieldList.Current as IFieldCOM ;
	    if not IsFieldToExclude(Field.name) and (not ExcludeBlobField or (Pos(Field.tipe, 'BLOB;DATA;') = 0)) then
	      Result := Result + iif(Result <> '', ', ', '') + Field.name;
    end;
	end
end;

function AglUpperCase(Parms : array of variant; nb : integer): Variant;
begin
  Result := UpperCase(string(Parms[1]));
end;

function AglLowerCase(Parms : array of variant; nb : integer): Variant;
begin
  Result := LowerCase(string(Parms[1]));
end;

procedure wInitProgressForm(Const Form: tWinControl; Const Caption, Title: string; Const MaxValue: integer; Const WithCancelButton, DisabledButtons: boolean);
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
	  if not V_Pgi.SilentMode then
	  begin
	    if ProgressFormImbrication = 0 then
      begin
        if DisabledButtons then
        begin
          if Assigned(Form) then
            CurrentForm := THForm(Form)
          else
            CurrentForm := THForm(Screen.ActiveForm);
          if not Assigned(CurrentForm) or (Assigned(CurrentForm.Parent) and CurrentForm.Parent.InheritsFrom(TPanel)) then
            CurrentForm :=THForm(Application.MainForm);
          CurrentForm.Enabled := False;
        end
        else
          CurrentForm := nil;

        try
  	      InitMoveProgressForm(CurrentForm, Caption, Title, MaxValue, WithCancelButton, DisabledButtons);
        except
          if DisabledButtons then
            CurrentForm.Enabled := True
        end
      end;

	    Inc(ProgressFormImbrication)
	  end
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

function wMoveProgressForm(Const Title: string = ''): boolean;
begin
  Result := True;
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
	  if not V_Pgi.SilentMode then
    begin
	    if ProgressFormImbrication = 1 then
	      Result := MoveCurProgressForm(Title)
      else
        TextProgressForm(Title);
    end
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

procedure wFiniProgressForm;
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
	  if not V_Pgi.SilentMode then
	  begin
	    ProgressFormImbrication := Max(0, ProgressFormImbrication - 1);

	    if ProgressFormImbrication = 0 then
      begin
	      FiniMoveProgressForm;
        if Assigned(CurrentForm) then
          CurrentForm.Enabled := True
      end;
	  end
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopulos
Créé le ...... : 20/08/2003
Modifié le ... :   /  /
Description .. : Construit un Range (séparé par des points virgule) à partir
Suite ........ : d'un Flus.
Suite ........ : Ex : Fplus renvoyé par une tablette dans le paramètre 'Range' du DISPATCHTT .
Suite ........ :   => 'ADR_TYPEADRESSE="TIE" AND ADR_REFCODE="C4M" AND ADR_LIVR="X"'
Mots clefs ... :
*****************************************************************}
function WMakeRangeFromFplus(Const Fplus: string): string;
var
  i:Integer;
  Arg1, Arg2:String;
begin
  { Chaine de résultat }
	Result := '';

  i:= pos('AND', Fplus);
  if i > 0 then
  begin
    // 1er argument
    Arg1 := Trim(Copy(Fplus, 1, i-1));
    // Suppression des guillemets dans la chaîne de caractère de l'argument
    Result:= Copy(Arg1, 1, Pos('=', Arg1)) + wExtractGuillemet(Copy(Arg1, Pos('=', Arg1) + 1, length(Arg1)));
    // Reste de l'argument
    Arg2 := Trim(Copy(Fplus, i+4, length(Fplus)));
    i:= pos('AND', Arg2);

    While i > 0 do
    begin
      Arg1 := Trim(Copy(Arg2, 1, i-1));
      Result:= Result + ';' + Copy(Arg1, 1, Pos('=', Arg1)) + wExtractGuillemet(Copy(Arg1, Pos('=', Arg1) + 1, length(Arg1)));
      Arg2 := Trim(Copy(Arg2, i+4, length(Fplus)));
      i:= pos('AND', Arg2);
    end;
    Result:= Result + ';' + Copy(Arg2, 1, Pos('=', Arg2)) + wExtractGuillemet(Copy(Arg2, Pos('=', Arg2) + 1, length(Arg2)));
  end
  else
    Result:= Fplus;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/09/2003
Modifié le ... : 29/09/2003
Description .. : Conversions d'unité de temps entre elles.
Suite ........ : Particularité: Gère aussi le temps en Heures centièmes
Suite ........ :
Mots clefs ... :
*****************************************************************}
function wConversionTemps(Temps: double; Const UniteDe: string; Const QuotiteDe: Double; Const UniteA: string; Const QuotiteA: Double): Double;
begin
  if Temps <> 0 then
  begin
    if UniteDe = 'H' then Temps := HHCC(Temps);
    Result := wDivise(Temps * QuotiteDe, QuotiteA);
    if UniteA = 'H' then Result := HHMM(Result);
  end
  else
    Result := 0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 10/02/2005
Modifié le ... :   /  /    
Description .. : Teste l'égalité de deux Double
Suite ........ : BUG Delphi corrigé en 6.0
Mots clefs ... : 
*****************************************************************}
function EgaliteDouble(Const n1, n2: double): boolean;
begin
  Result := ((n1 = 0) and (n2 = 0)) or (abs(n1-n2) < Min(abs(n1),abs(n2)) * 1E-12);
end; 

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Arrondir une quantité, un prix ou tout autre chose en
Suite ........ : focntion d'une précision d'arrondi et d'une méthode d'arrondi
Mots clefs ... :
*****************************************************************}
function wArrondir(Const nValeur: double; const ArrondiPrec: tArrondiPrec; Const ArrondiMeth: tArrondiMeth): double;
var
  nPoids: Double;
begin {wArrondir}
  nPoids := ArrondiPrecToPoids(ArrondiPrec);

  Result := nValeur / nPoids;
  case ArrondiMeth of
    tAMInferieure: Result := Trunc(Result);
    tAMPlusProche: Result := Arrondi(Result, 0);
    tAMSuperieure: Result := Arrondi(Result+0.4989, 0);
  end;
  Result := Result * nPoids;
end; {wArrondir}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Transforme le nombre de décimal d'arrondi sauvegardé dans la base
Suite ........ : en un type d'arrondi géré dans les écran et la fonction
Suite ........ : d'arrondi
Mots clefs ... :
*****************************************************************}
function NbDecToArrondiPrec(Const iNbDec: integer): tArrondiPrec;
begin {NbDecToArrondiPrec}
  Result := tApDixMillieme;
  Case iNbDec of
    -5: Result := tAPCentaineDeMillier;
    -4: Result := tAPDizaineDeMillier;
    -3: Result := tAPMillier;
    -2: Result := tAPCentaine;
    -1: Result := tAPDizaine;
     0: Result := tApUnite;
    +1: Result := tApDizieme;
    +2: Result := tApCentieme;
    +3: Result := tApMillieme;
    +4: Result := tApDixMillieme;
  else if V_Pgi.Sav then
    PgiError('Le nombre de décimal d''arrondi n''a pas de correspondance connue, l''arrondi se fera au 10 millième');
  end;
end; {NbDecToArrondiPrec}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Transforme un poids d'arrondi sauvegardé dans la base
Suite ........ : en un type d'arrondi géré dans les écran et la fonction
Suite ........ : d'arrondi
Mots clefs ... :
*****************************************************************}
function PoidsToArrondiPrec(Const nPoids: double): tArrondiPrec;
begin {PoidsToArrondiPrec}
  Result := tApDixMillieme;
  if      EgaliteDouble(nPoids,100000     ) then Result := tAPCentaineDeMillier
  else if EgaliteDouble(nPoids, 10000     ) then Result := tAPDizaineDeMillier
  else if EgaliteDouble(nPoids,  1000     ) then Result := tAPMillier
  else if EgaliteDouble(nPoids,   100     ) then Result := tAPCentaine
  else if EgaliteDouble(nPoids,    10     ) then Result := tAPDizaine
  else if EgaliteDouble(nPoids,     1     ) then Result := tApUnite
  else if EgaliteDouble(nPoids,     0.1   ) then Result := tApDizieme
  else if EgaliteDouble(nPoids,     0.01  ) then Result := tApCentieme
  else if EgaliteDouble(nPoids,     0.001 ) then Result := tApMillieme
  else if EgaliteDouble(nPoids,     0.0001) then Result := tApDixMillieme
  else if V_Pgi.Sav then
  begin
    PgiError('Le poids d''arrondi n''a pas de correspondance connue, l''arrondi se fera au 10 millième');
  end;
end; {PoidsToArrondiPrec}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 09/02/2005
Modifié le ... : 10/02/2005
Description .. : Transforme  une précision d'arrondi
Suite ........ : en un nombre de décimal d'arrondi
Mots clefs ... :
*****************************************************************}
function ArrondiPrecToNbDec(Const ArrondiPrec: tArrondiPrec): integer;
begin {ArrondiPrecToNbDec}
  Result := 0;
  Case ArrondiPrec of
    tAPCentaineDeMillier : Result := -5;
    tAPDizaineDeMillier  : Result := -4;
    tAPMillier           : Result := -3;
    tAPCentaine          : Result := -2;
    tAPDizaine           : Result := -1;
    tApUnite             : Result :=  0;
    tApDizieme           : Result := +1;
    tApCentieme          : Result := +2;
    tApMillieme          : Result := +3;
    tApDixMillieme       : Result := +4;
  else if V_Pgi.Sav then
    PgiError('La précision d''arrondi n''est pas référencé dans le programme');
  end;
end; {ArrondiPrecToNbDec}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. : Transforme  une précision d'arrondi
Suite ........ : en un poids d'arrondi
Mots clefs ... :
*****************************************************************}
function ArrondiPrecToPoids(Const ArrondiPrec: tArrondiPrec): double;
begin {ArrondiPrecToPoids}
  Result := 0.0;
  Case ArrondiPrec of
    tAPCentaineDeMillier : Result := 100000     ;
    tAPDizaineDeMillier  : Result :=  10000     ;
    tAPMillier           : Result :=   1000     ;
    tAPCentaine          : Result :=    100     ;
    tAPDizaine           : Result :=     10     ;
    tApUnite             : Result :=      1     ;
    tApDizieme           : Result :=      0.1   ;
    tApCentieme          : Result :=      0.01  ;
    tApMillieme          : Result :=      0.001 ;
    tApDixMillieme       : Result :=      0.0001;
  else if V_Pgi.Sav then
    PgiError('La précision d''arrondi n''est pas référencé dans le programme');
  end;
end; {ArrondiPrecToPoids}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 09/02/2005
Modifié le ... :   /  /
Description .. : Transforme une méthode d'arrondi en type d'arrondi
Mots clefs ... :
*****************************************************************}
function MethodeToArrondiMeth(Const sMethode: string; lMessageErreur: boolean = True): tArrondiMeth;
begin {MethodeToArrondiMeth}
  if      (sMethode='I') then Result := tAMInferieure
  else if (sMethode='P') then Result := tAMPlusProche
  else if (sMethode='S') then Result := tAMSuperieure
  else
  begin
    if lMessageErreur then
      PgiError('La méthode d''arrondi n''a pas de correspondance connue, l''arrondi se fera au plus proche');
    Result := tAMPlusProche;
  end;
end; {MethodeToArrondiMeth}

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 25/08/2006
Modifié le ... :   /  /
Description .. : Transforme le quoi en type  de quoi
Mots clefs ... :
*****************************************************************}
function QuoiToArrondiQuoi(const sQuoi: string; lMessageErreur: boolean = True): tArrondiQuoi;
begin {QuoiToArrondiQuoi}
  if      (sQuoi='T') then Result := tTPMTaux
  else if (sQuoi='P') then Result := tTPMPrix
  else if (sQuoi='M') then Result := tTPMMont
  else
  begin
    if lMessageErreur then
      PgiError('La type de donnée à arrondir n''a pas de correspondance connue, l''arrondi se fera dans une logique de montant');
    Result := tTPMMont;
  end;
end; {QuoiToArrondiQuoi}

function wIsFieldFromKey1(Const FieldName: String): Boolean;
var
	Mcd : IMCDServiceCOM;
  Table : ITableCOM;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Table := Mcd.GetTable(PrefixeToTable(ExtractPrefixe(Trim(UpperCase(FieldName)))));
  //
  Result := ( Assigned(Table)) and (Pos(Trim(UpperCase(FieldName)) + ',', Trim(UpperCase(Mcd.TableToCle1(Table.Name))) + ',') <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ..... : Thierry Petetin
Créé le ..... : 24/02/2004
Description . : Appel de TobDebug de l'AGL : ModeSAV + Mot de passe du jour
*****************************************************************}
procedure wShowMeTheTob(Const T: Tob);
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    { Vérifie si on peut montrer la Tob : Mode SAV + Mot de Passe du jour }
    {$IFDEF GPAO}
      if Assigned(T) and V_PGI.SAV and (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
        TobDebug(T);
    {$ENDIF GPAO}
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 24/02/2004
Description .. : Pour les nostalgiques ....  :o)
*****************************************************************}
Procedure www(T: Tob);
begin
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    {$IFDEF GPAO}
      TobDebug(T);
    {$ENDIF GPAO}
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
end;

{$IFNDEF EAGLCLIENT}
Procedure www(DS: TDataSet);
var
  T: Tob;
begin
  if Assigned(DS) then
  begin
    T := Tob.Create('_DataSet-Debug_', nil, -1);
    try
      wMakeTobFromDS(DS, T, False, True);
      wMakeTobFromDS(DS, T, True , True);
      www(T)
    finally
      T.Free
    end;
  end
end;
{$ENDIF !EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 21/12/2006
Modifié le ... :   /  /    
Description .. : Permet de visualiser le résultat d'une requête dans une Tob. 
Suite ........ : Sert beaucoup dans le cadre d'une transaction où le 
Suite ........ : processus en cours lock une table. Du coup il n'est pas 
Suite ........ : possible de contrôler en temps réel les instructions SQL 
Suite ........ : effectuées.
Mots clefs ... : 
*****************************************************************}
Procedure www(const Sql: String); 
var
  T: Tob;
begin
  if Sql = '' then
    Exit;
  T := Tob.Create('_Sql-Debug_', nil, -1);
  try
    if T.LoadDetailDBFromSQL('', Sql) then
      www(T)
  finally
    T.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 12/12/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Dans un argument, renvoie la string d'un rang
Mots clefs ... :
*****************************************************************}
function GetTokenByNum(Chaine: string; Num: integer): string;
begin
  While Num > 0 do
  begin
    Result := ReadTokenSt(Chaine);
    Dec(Num);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 13/04/2004
Modifié le ... : 13/04/2004
Description .. :
Suite ........ : Transforme une liste de valeurs issue d'une MultiComboBox
Suite ........ : en liste conforme à un IN pour WHERE
Mots clefs ... :
*****************************************************************}
function GetListeINFromListeCombo(sListeDeValeurCombo: string): string;
begin
  Result := '';
  while (sListeDeValeurCombo<>'') do
  begin
    Result := Result+'"'+ReadTokenSt(sListeDeValeurCombo)+'",'
  end;
  Result := copy(Result, 1 , length(Result)-1 );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/02/2005
Modifié le ... :   /  /    
Description .. : Dit si on est rentré avec le mot de passe du jour
Mots clefs ... :
*****************************************************************}
function IsDayPass: boolean;
begin
  Result := DelphiRunning or (V_Pgi.PassWord = CryptageSt(DayPass(V_Pgi.DateEntree)))
end;

{ Permet de décompresser une chaîne : algo ZIP }
function wDezipString(Const St: String): String;
begin
  try
    Result := DecompressAsciiString(St)
  except
    Result := St
  end
end;

{ Permet de compresser une chaîne : algo ZIP }
function wZipString(Const St: String): String;
begin
  try
    Result := CompressAsciiString(St)
  except
    Result := St
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : Balaie les champs d'une tob :
Suite ........ : Si "Null" trouvé, on initialise la valeur.
Suite ........ : Permet en ORACLE de ne pas avoir de surprise avec un Tob.PutEcran par ex.
Mots clefs ... : 
*****************************************************************}
procedure wCheckNullFromTob(Const T: Tob);
var
  i: Integer;
begin
  for i := 1 to T.NbChamps do
    if T.GetValeur(i) = Null then
      T.PutValeur(i, wGetInitValue(T.GetNomChamp(i)));

  for i := 1000 to (1000 + Pred(T.ChampsSup.Count)) do
    if T.GetValeur(i) = Null then
      T.PutValeur(i, wGetInitValue(T.GetNomChamp(i)))
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{ permet de changer la DBListe associée au mul dans le OnArgument de la TOF }
procedure wChangeDBListe(Const Mul: TFMul; const DBListe: String);
begin
  if Assigned(Mul.Q) then
    Mul.SetDBListe(DBListe);
  UpdateCaption(Mul);
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 24/08/2004
Modifié le ... :   /  /
Description .. : crée autant de champsup dans la tob "T" que de champs
Suite ........ : dans la table "TableName"
Mots clefs ... :
*****************************************************************}
procedure wCreateChampSupFromTable(Const T: Tob; const TableName: String);
var
	Table     : ITableCOM ;
  NomChamps,TypeChamps : string;
	Mcd : IMCDServiceCOM;
	FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  if Assigned(T) then
  begin
    Table := Mcd.GetTable(TableName);
    if Assigned(table) then
    begin
      FieldList := Table.Fields;
      FieldList.Reset();
      While FieldList.MoveNext do
      begin
				NomChamps := (FieldList.Current as IFieldCOM).name;
				TypeChamps := (FieldList.Current as IFieldCOM).tipe;
	      if (NomChamps<>'') then
	        T.AddChampSupValeur(NomChamps, wGetInitValue(TypeChamps));
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 01/09/2005
Modifié le ... :   /  /    
Description .. : Copie un Item d'une TStringList vers une autre TStringList.
Suite ........ : Si ExcludeDoublon est à "Vrai", la copy interdit les doublons
Suite ........ : Si DeleteItemOrigine est à "Vrai", l'item sera supprimé de la
Suite ........ : première liste.
Suite ........ : ATTENTION : Tout est basé sur TStringList.Name... pas Item
Suite ........ : Retour = 0 : l'item a été copié
Suite ........ : Retour < 0 : l'item d'origine n'existe pas
Suite ........ : Retour > 0 : l'item de destination existe déjà
Mots clefs ... : 
*****************************************************************}
function wCopyItemFromListToList(const Item: String; FromList, ToList: TStrings; const ByItemName: Boolean;
                                 const ExcludeDoublon: Boolean = True;
                                 const DeleteItemOrigine: Boolean = False;
                                 const InsertAtPos: Integer = -1): Integer;
var
  FromIndexOf, ToIndexOf: Integer;
begin
  if ByItemName then
    Result := FromList.IndexOfName(Item)
  else
    Result := FromList.IndexOf(Item);
    
  FromIndexOf := Result;
  if Result >= 0 then
  begin
    if ByItemName then
      ToIndexOf := ToList.IndexOfName(Item)
    else
      ToIndexOf := ToList.IndexOf(Item);

    if ToIndexOf >= 0 then
      Result := 1
    else
      Result := 0;

    if (Result = 0) or not ExcludeDoublon then
    begin
      if InsertAtPos = -1 then
        ToList.Add(FromList[FromIndexOf])
      else
        ToList.Insert(InsertAtPos, FromList[FromIndexOf]);
      if DeleteItemOrigine then
        FromList.Delete(FromIndexOf)
    end
  end
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/09/2005
Modifié le ... :   /  /    
Description .. : Permet d'extraire une section d'une chaîne de caractères, 
Suite ........ : façon .ini :
Suite ........ : [SECTION1] bla bla [SECTION2] bla2 bla2
Suite ........ : Section : Nom de la section à extraire sans les '[]'
Suite ........ : Memo : Chaîne source
Suite ........ : ModifyMemo : Permet de mettre à jour le param. "Memo" 
Suite ........ : (Memo - Section extraite)
Mots clefs ... : 
*****************************************************************}
function wExtractSectionFromMemo(const Section: String; var Memo: String; const ModifyMemo: Boolean): String;
var
  TS, TSSection: TStringList;
  i: Integer;

  function EoSection(const i: Integer): Boolean;
  begin
    Result := (i >= TS.Count)
           or ((Pos('[', TS[i]) > 0) and (Pos('[', TS[i]) < Pos(']', TS[i])))
  end;

  function GetISection: Integer;
  var
    i: Integer;
  begin
    Result := -1;
    i := 0;
    while (Result = -1) and (i < TS.Count) do
    begin
      if Pos('[' + Section + ']', TS[i]) > 0 then
        Result := i;
      Inc(i)
    end
  end;

begin
  TS := TStringList.Create();
  TSSection := TStringList.Create();
  try
    TS.SetText(PChar(Memo));
    Result := '';
    i := GetISection;
    if i >= 0 then
    begin
      TS[i] := Trim(StringReplace(TS[i], '[' + Section + ']', '', [rfIgnoreCase]));
      if TS[i] = '' then
        TS.Delete(i);
      while not EoSection(i) do
      begin
        TSSection.Add(TS[i]);
        TS.Delete(i)
      end;
      Result := Trim(TSSection.Text);
      if ModifyMemo then
        Memo := TS.Text
    end;
  finally
    TSSection.Free;
    TS.Free
  end
end;

function isVarInt(const ValueV: Variant): Boolean;
begin
  Result := VarType(ValueV) in [varShortInt, varWord, varByte, varLongWord, varInt64, varInteger]
end;

function GetFieldType(FieldName : string) : tTypeField;
begin
//*******************************
  Result := Tools.GetFieldType(FieldName{$IF defined(APPSRV)}, '', '', nil{$IFEND APPSRV});
//*******************************
end;

function GetFieldSize(FieldName : string) : integer;
var
  FieldType : string;
  Extract   : string;
begin
  if GetFieldType(FieldName) = ttfText then
  begin
    FieldType := ChampToType(FieldName);
    Extract   := Copy(FieldType, pos('(', FieldType) + 1, Length(FieldType));
    Result    := StrToInt(Copy(Extract, 1, Length(Extract) -1));
  end else
    Result := -1;
end;

{ ******************************** Gestion des listes * DEBUT ************************* }

function wGetColsVisibilityFromListe(const ListeName: String; const OnlyFields: Boolean = False): String;
var
  L: TWListe;
  i: Integer;
begin
  Result := '';
  L := TWListe.Create(ListeName, nil, OnlyFields);
  try
    for i := 0 to Pred(L.ChampsCount) do
      Result := Result + iif(Result <> '', ';', '') + L.Champs[i].Name + '=' + BoolToStr_(L.Champs[i].Visible)
  finally
    L.Free
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 18/06/2008
Modifié le ... :   /  /
Description .. : Renvoie la liste des champs d'une liste faisant partie des
Suite ........ : préfixes demandé afin de les mettre dans un SELECT
Suite ........ : > Ne renvoie que des champs
Suite ........ : > ne renvoie pas les invisibles
Mots clefs ... :
*****************************************************************}
function GetSelectColsFromListe(const ListeName: String; Const Prefixes: String = ''): String;
var
  L: TWListe;
  i: Integer;
begin
  Result := '';
  L := TWListe.Create(ListeName, nil, true);
  try
    for i := 0 to Pred(L.ChampsCount) do
      if ((Prefixes = '') or (Pos(wGetPrefixe(L.Champs[i].Name), Prefixes) > 0))
         and (L.Champs[i].Visible) then
        Result := Result + iif(Result <> '', ',', '') + L.Champs[i].Name
  finally
    L.Free
  end
end;

{ TWFieldListe ************ DEBUT ************** }

constructor TWFieldListe.Create(AListe: TWListe; const AName, ACaption: String; const AWidth: Integer; Params: String);
const
  NbParams = 2;
begin
  FParent := AListe;
  FName := AName;
  if Params <> '' then
    FWidth := Round(AWidth * 6.3)
  else
    FWidth := AWidth;
  FCaption := ACaption;
  if Params <> '' then
  begin
    FLibelleComplet := Pos('$', Params) > 0;
    FObligatoire    := Pos('O', Params) > 0;
    case Params[1] of
      'C': FAlignment := taCenter;
      'D': FAlignment := taRightJustify;
    else
      FAlignment := taLeftJustify
    end;
    FFormat := '#';
    if Params[2] = '.' then
      FFormat := FFormat + ' ';
    FFormat := FFormat + '##0';
    if ValeurI(Params[3]) > 0 then
      FFormat := FFormat + '.' + wStringRepeat('0', ValeurI(Params[3]));

    Params := Copy(Params, Length(Params) - NbParams, 3);
    FVisible := not StrToBool_(Params[1]);
    FEmptyIfNull := StrToBool_(Params[2]);
  end;

  //GP_20071217_TS_GP14539 >>>
  SetLength(FSqlValues, 0);
  SetLength(FSqlFields, 0);

  if (ExtractPrefixe(AName) <> '') and (ChampToNum(AName) > 0) then
    FTypeField := wGetSimpleTypeField(AName)
  else
  begin
    { Type par défaut numérique pour les formules }
    FTypeField := 'N';
    { Traitement des champs formules  }
//GP_20080409_TS_GP14539
    if FParent.FGereSqlFields and (Pos('SELECT ', UpperCase(AName)) > 0) then
      Sql := UpperCase(Trim(AName));
  end;
  //GP_20071217_TS_GP14539 <<<

  if (FTypeField = 'N') or (FTypeField = 'I') then
  begin
    FImage := False;
    if Params <> '' then
      FCumul := StrToBool_(Params[3]);
  end
  else
  begin
    if Params <> '' then
      FImage := StrToBool_(Params[3]);
    FCumul := False;
  end
end;

function TWFieldListe.GetIndex: Integer;
begin
  Result := FParent.FChamps.IndexOf(Self)
end;

function TWFieldListe.GetTableName: String;
begin
  Result := PrefixeToTable(ExtractPrefixe(Name))
end;

//GP_20071217_TS_GP14539 >>>
{ Vire les C1 & C2 respectivement de début et fin }
procedure TWFieldListe.TrimChar(var St: String; const C1, C2: Char);
begin
  if (St <> '') and (St[Length(St)] = C2) and (St[1] = C1) then
  begin
    Delete(St, 1, 1);
    Delete(St, Length(St), 1);
    St := Trim(St);
    TrimChar(St, C1, C2)
  end;
end;

{ Trouve la position du dernier jeton trouvé }
function TWFieldListe.GetTokenPosFromTheEnd(const S, TokenToFind: String): Integer;
var
  Token: String;
  CharOkToken: Char;
begin
  Result := Length(S);
  Token := '';
  { on recherche l'alias depuis la fin }
  while (Token <> TokenToFind) and (Result > 0) do
  begin
    if Length(TokenToFind) - Length(Token) <= 0 then
      CharOkToken := #0
    else
      CharOkToken := TokenToFind[Length(TokenToFind) - Length(Token)];
    if CharOkToken = #0 then
      CharOkToken := TokenToFind[Length(TokenToFind)];
    if S[Result] = CharOkToken then
      Token := S[Result] + Token
    else
      Token := '';
    Dec(Result);
  end;
  if Result > 0 then
    Inc(Result);
end;

{ Appelé depuis Formule.ChampFormule afin de lister les variables }
function TWFieldListe.FindInFormule(FieldName: HString): Variant;
var
  iPosValue,
  iPosField, i: Integer;
  Where: String;
begin
  SetLength(FSqlValues, Length(FSqlValues) + 1);
  FSqlValues[Pred(Length(FSqlValues))] := FieldName;

  { Liste également les champs comparés }
  Where := FSqlWhere;
  iPosValue := Pos('[' + FieldName + ']', Where);
  if iPosValue > 0 then
  begin
    iPosField := GetTokenPosFromTheEnd(wLeft(Where, iPosValue), '=');
    if iPosField > 0 then
    begin
      i := iPosField - 1;
      { ignore les ' ' }
      while (Length(Where) >= i) and (Where[i] = ' ') do
        Delete(Where, i, 1);
      { récupère le nom de champ }
      while (Length(Where) >= i) and (Where[i] in ['A'..'Z', '0'..'9', '_']) do
        Dec(i);
      if i <> iPosField - 1 then
      begin
        SetLength(FSqlFields, Length(FSqlFields) + 1);
        FSqlFields[Pred(Length(FSqlFields))] := Copy(Where, i, iPosField - 1);
      end;
    end;
  end
end;

{ récupère le WHERE de la formule SQL. Ces formules étant "simples", on considère
  que le dernier WHERE de la requête est celui concernant la requête "principale" }
function TWFieldListe.DecodeWhere(var Sql: String): Boolean;
var
  iPosStartWhere,
  iPosEndWhere  : Integer;
const
  _WHERE_ = ' WHERE ';
begin
  { les parenthèses "superflues" sont ici déjà exclues }
  { 1. on retrouve le WHERE }
  iPosStartWhere := GetTokenPosFromTheEnd(Sql, _WHERE_);
  Result := iPosStartWhere > 0;
  if Result then
  begin
    iPosStartWhere := iPosStartWhere + Length(_WHERE_);
    { 2. on l'isole -> FSqlWhere }
    iPosEndWhere := GetTokenPosFromTheEnd(Sql, ' ORDER BY ');
    if iPosEndWhere < iPosStartWhere then
      iPosEndWhere := GetTokenPosFromTheEnd(Sql, ' GROUP BY ');
    if iPosEndWhere < iPosStartWhere then
      iPosEndWhere := Length(Sql);

    FSqlWhere := Trim(Copy(Sql, iPosStartWhere, iPosEndWhere - iPosStartWhere + 1));
    Delete(Sql, iPosStartWhere, iPosEndWhere - iPosStartWhere + 1);
    Insert('%s', Sql, iPosStartWhere + 1);
    TrimChar(FSqlWhere, '(', ')');
    { 3. on extrait les critères }
    Result := FSqlWhere <> '';
    if Result then
      ChampFormule(FSqlWhere, FindInFormule, nil, 1)
  end;
end;

{ cas d'une formule SQL, renseigne les propriétés :
  . SQL avec la chaîne SQL simplifiée
  . NAME avec un alias automatique ou trouvé dans la chaîne
  Ne se veut pas être un parser de chaîne SQL, mais analyse un SQL simple }
procedure TWFieldListe.SetSql(const Value: String);
var
  i, iPosAlias,
  iPosFrom: Integer;
  Alias: String;
const
  _AS_ = ' AS ';
  _FROM_ = ' FROM ';
begin
  FSql := Value;

  { Extraction des éventuelles parenthèses début/fin }
  TrimChar(FSql, '(', ')');

  { on écarte les formules type @ }
  if Pos('@', FSql) in [1, 2] then
    FSql := '';

  { Extraction/Initialisation de l'alias }
  if FSql <> '' then
  begin
    iPosFrom := GetTokenPosFromTheEnd(FSql, _FROM_);
    { ' AS ' trouvé }
    iPosAlias := GetTokenPosFromTheEnd(FSql, _AS_);
    if iPosAlias > 0 then
    begin
      { Suppression du ' AS ' dans la chaîne sql }
      if iPosAlias > iPosFrom then
        Delete(FSql, iPosAlias, Length(_AS_))
      else
        Inc(iPosAlias, Length(_AS_));
      { Affecte l'alias de la position trouvée jusqu'à la fin de la chaîne }
      Alias := Trim(Copy(FSql, iPosAlias, Length(FSql) - iPosAlias + 1));
      if Alias <> '' then
      begin
        { Alias entre quotes }
        if Alias[1] = '"' then
        begin
          if iPosAlias > iPosFrom then
            Delete(FSql, iPosAlias, 1)
          else
            Inc(iPosAlias, 1);
          Delete(Alias, 1, 1);
          i := 1;
          { recherche de la " suivante }
          while (i <= Length(Alias)) and (Alias[i] <> '"') do
            Inc(i);
          { erreur de " }
          if i > Length(Alias) then
            Alias := ''
          else { Extraction de l'alias }
          begin
            if iPosAlias > iPosFrom then
              Delete(FSql, iPosAlias + i - 1, 1);
            Alias := wLeft(Alias, i - 1);
          end;
          { suppression de l'alias dans la chaîne Sql si après le FROM (i.e. en fin de SELECT) }
          if iPosAlias > iPosFrom then
            Delete(FSql, iPosAlias, Length(Alias));
        end
        else { Extraction de l'alias sur les caractères autorisés }
        begin
          i := 1;
          while (i <= Length(Alias)) and (Alias[i] in ['A'..'Z', '0'..'9', '_']) do
            Inc(i);
          Dec(i);
          Alias := Copy(Alias, 1, i);
          { suppression de l'alias dans la chaîne Sql }
          if iPosAlias > iPosFrom then
            Delete(FSql, iPosAlias, Length(Alias));
        end;
      end
    end;

    { Alias automatique }
    if Alias = '' then
    begin
      Inc(FParent.FIndexAlias);
      Alias := 'SQLFIELD' + wPadLeft(IntToStr(FParent.FIndexAlias), 2, '0');
    end
    else { L'alias renseigné correspond peut-être à un champ. On récupère son type }
      FTypeField := wGetSimpleTypeField(Alias);

    FSql := Trim(FSql);

    { Extraction des éventuelles parenthèses début/fin restantes }
    TrimChar(FSql, '(', ')');
    FName := Alias;

    { décode un where simple (CHAMP1=CHAMP2, Etc) }
    if DecodeWhere(FSql) then
    begin
      if Length(FSqlFields) > 0 then
      begin
        iPosFrom := GetTokenPosFromTheEnd(FSql, _FROM_);
        Insert(', ' + wArrayToString(FSqlFields, ', '), FSql, iPosFrom);
      end
    end
    else
    begin
      { erreur : le champ n'est pas considéré comme SQL et sera invisible }
      SetLength(FSqlValues, 0);
      SetLength(FSqlFields, 0);
      FName := Value;
      Sql := '';
      FVisible := False
    end;
  end;
end;

{ s'agit-il d'une formule type SQL ? }
function TWFieldListe.isSql: Boolean;
begin
  Result := FSql <> ''
end;
//GP_20071217_TS_GP14539 <<<

{ TWFieldListe ************* FIN *************** }

{ TWListe ***************** DEBUT ************** }

{ Permet de "décoder" une liste de paramétrage }
//GP_20080409_TS_GP14539 >>>
constructor TWListe.Create(const DBListe: String; Ecran: TObject; const OnlyFields: Boolean = False;
                           const GereSqlFields: Boolean = False);
//GP_20080409_TS_GP14539 <<<
var
  sColParams,
  sColName,
  sColWidth,
  sTables,
  sP, sN, sC  : String;
  sColCaption : HString;   
  LastIndex, iW: Integer;
  T            : Tob;

  procedure DeleteField;
  begin
    Champs[LastIndex].Free;
    FChamps.Delete(LastIndex);
  end;

  procedure SetTables;
  begin
    sTables := StringReplace(sTables, ',', ';', [rfIgnoreCase, rfReplaceAll]);
    while sTables <> '' do
    begin
      SetLength(FTables, Length(FTables) + 1);
      FTables[Pred(Length(FTables))] := Trim(ReadTokenSt(sTables))
    end
  end;

var
  NA1,  NA4: String;
  NA2, NA3 : hstring;
  NA5, NA6: Boolean;
begin
//GP_20080409_TS_GP14539
  FGereSqlFields := GereSqlFields;
  FDBListe := DBListe;
  FChamps := TList.Create();
  SetLength(FTables, 0);
  T := Tob.Create('LISTE', nil, -1);
  try
    if ChargeHListeUser(FDBListe, V_Pgi.User, sTables, FJoin, FOrderBy, sColName, sColCaption,
                                     sColWidth, sColParams, NA1, NA2, NA3, NA4, NA5, NA6) then
    begin
      FJoinByFrom := Pos(',', sTables) > 0;
      SetTables;
      while sColName <> '' do
      begin
        sN := ReadTokenSt(sColName);
        sC := ReadTokenSt(sColCaption);
        iW := ValeurI(ReadTokenSt(sColWidth));
        sP := ReadTokenSt(sColParams);

        if not OnlyFields or (ChampToNum(sN) > 0) then
        begin
          LastIndex := FChamps.Add(TWFieldListe.Create(Self, sN, sC, iW, sP));

          {$IFNDEF EAGLSERVER}
          {$IFNDEF ERADIO}
            if Assigned(Ecran) then
            begin
              { épuration selon contexte }
              {$IFNDEF EAGLCLIENT}
                if not Champs[LastIndex].FVisible then
                  DeleteField;
              {$ELSE !EAGLCLIENT}
                if (Ecran is TFSaisieList) and not Champs[LastIndex].FVisible then
                  DeleteField
                else if (Ecran.InheritsFrom(TForm) and not (Ecran is TFSaisieList)) and Champs[LastIndex].FLibelleComplet and Champs[LastIndex].FObligatoire then
                begin
                  { Ajout d'un champ car Code + Libellé présents dans la grille }
                  LastIndex := FChamps.Add(TWFieldListe.Create(Self, sN, sC, iW, sP));
                  Champs[LastIndex].FName := Champs[LastIndex].FName + '_CODE';
                end
              {$ENDIF !EAGLCLIENT}
            end
          {$ENDIF !ERADIO}
          {$ENDIF EAGLSERVER}
        end;
      end;
    end;
  finally
    T.Free;
  end
end;

procedure TWListe.Clear;
begin
  while FChamps.Count > 0 do
  begin
    TWFieldListe(FChamps[Pred(FChamps.Count)]).Free;
    FChamps.Delete(Pred(FChamps.Count))
  end;
  SetLength(FTables, 0)
end;

destructor TWListe.Destroy;
begin
  Clear;
  FChamps.Free;

  inherited;
end;

function TWListe.GetChamps(Index: Integer): TWFieldListe;
begin
  if wBetween(Index, 0, Pred(FChamps.Count)) then
    Result := TWFieldListe(FChamps[Index])
  else
    Result := nil
end;

function TWListe.GetChampsByName(ColName: String): TWFieldListe;
var
  Index: Integer;
  Found: Boolean;
begin
  Found := False;
  Index := 0;
  while not Found and (Index < FChamps.Count) do
  begin
    Found := UpperCase(Trim(TWFieldListe(FChamps[Index]).Name)) = UpperCase(Trim(ColName));
    if not Found then
      Inc(Index)
  end;
  if Found then
    Result := TWFieldListe(FChamps[Index])
  else
    Result := nil
end;

function TWListe.GetCount: Integer;
begin
  Result := FChamps.Count
end;

function TWListe.GetCountVisible: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Pred(FChamps.Count) do
    if Champs[i].Visible then
      Inc(Result)
end;

function TWListe.GetIsChamps(Index: Integer): Boolean;
begin
  Result := ChampToNum(Champs[Index].Name) > 0
end;

function TWListe.GetFields(const Format: TWListeString): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to Pred(FChamps.Count) do
  begin
    if  ((Champs[i].Visible     and (wlsVisible     in Format)) or not (wlsVisible     in Format))
    and ((Champs[i].Obligatoire and (wlsObligatoire in Format)) or not (wlsObligatoire in Format)) then
      Result := Result + iif(Result <> '', ';', '') + Champs[i].Name
  end
end;

function TWListe.GetCaptions(const Format: TWListeString): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to Pred(FChamps.Count) do
  begin
    if  ((Champs[i].Visible     and (wlsVisible     in Format)) or not (wlsVisible     in Format))
    and ((Champs[i].Obligatoire and (wlsObligatoire in Format)) or not (wlsObligatoire in Format)) then
      Result := Result + iif(Result <> '', ';', '') + Champs[i].Caption
  end
end;

function TWListe.TablesCount: Integer;
begin
  Result := Length(FTables)
end;

function TWListe.GetTable(Index: Integer): String;
begin
  if wBetween(Index, Low(FTables), High(FTables)) then
    Result := FTables[Index]
  else
    Result := ''
end;

function TWListe.SQL: String;

  function GetSelect: String;
  var
    i: Integer;
  begin
    Result := '';
    for i := 0 to Pred(ChampsCount) do
      Result := Result + iif(Result <> '', ', ', '') + Champs[i].Name
  end;

  function GetFrom: String;
  var
    i: Integer;
  begin
    Result := '';
    if JoinByFrom then
      for i := 0 to Min(1, Pred(TablesCount)) do
        Result := Result + iif(Result <> '', ', ', '') + Tables[i]
    else
      Result := Tables[0]
  end;

begin
  Result := 'SELECT ' + GetSelect
          + ' FROM '  + GetFrom
          ;
  if Join <> '' then
    Result := Result + iif(not JoinByFrom, ' LEFT JOIN ' + Join, ' WHERE ' + Join);
  if OrderBy <> '' then
    Result := Result + ' ORDER BY ' + OrderBy
end;

procedure TWListe.AddField(const ColName, ColTitle: String;
                           const ColWidth: Integer; const ColType: Char;
                           const ColFormat: String; const ColAlign: TAlignment;
                           const ColImage: Boolean = False; const ColCumul: Boolean = False;
                           const ColLibelleComplet: Boolean = False);
var
  FieldListe: TWFieldListe;
begin
  FieldListe := TWFieldListe.Create(Self, ColName, ColTitle, ColWidth, '');
  with FieldListe do
  begin
    Format := ColFormat;
    Alignment := ColAlign;
    FVisible := True;
    Image := ColImage and (TypeField = 'C');
    FObligatoire := True;
    FEmptyIfNull := (TypeField = 'D') or (TypeField = 'I') or (TypeField = 'N');
    FCumul := ColCumul and ((TypeField = 'I') or (TypeField = 'N'));
    FLibelleComplet := ColLibelleComplet and (TypeField = 'C');
  end;
  FChamps.Add(FieldListe);
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin & Thibaut Sublet
Créé le ...... : 17/03/2005
Description .. : Rendre invisible une ou plusieurs colonnes de la liste d'un
Suite ........ : mul (Pour les mul non issus d'une wTOF)
Suite ........ : A placer dans le OnUpdate de la Tof
*****************************************************************}
procedure wHideListColumns(Const Ecran: TForm; Const ColumnsName: Array of string);
var
  iChamp: integer;
{$IFDEF EAGLCLIENT}
  wListe: TWListe;
  wField: TWFieldListe;
{$ENDIF EAGLCLIENT}
begin
{$IFDEF EAGLCLIENT}
  wListe := TWListe.Create(TFMul(Ecran).DBListe, Ecran);
  try
    for iChamp := Low(ColumnsName) to High(ColumnsName) do
    begin
      wField := wListe.ChampsByName[ColumnsName[iChamp]];
      if Assigned(wField) and (wField.Index > -1) then
      begin
        TFMul(Ecran).SetVisibleColumn (wField.Name, false);
      end;
    end;
  finally
    wListe.Free;
  end;
{$ELSE EAGLCLIENT}
  for iChamp := Low(ColumnsName) to High(ColumnsName) do
    TFMul(Ecran).SetVisibleColumn(ColumnsName[iChamp], False);
{$ENDIF EAGLCLIENT}
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{ ******************************** Gestion des listes *  FIN  ************************* }

procedure wLoadPopUpFromCommun(Const PopUp: TPopupMenu; Const CoType: String; Const AOnClick: TNotifyEvent; Const Plus: String = '');
var
  i: Integer;
  Sql: String;
  TobCO: Tob;
  T: Array of TMenuItem;
begin
  TobCO := Tob.Create('CO', nil, -1);
  try
    { Charge la tablette }
    Sql := 'SELECT CO_CODE,CO_LIBELLE '
         + 'FROM COMMUN '
         + 'WHERE CO_TYPE="' + CoType + '"'
         ;
    if Plus <> '' then
      Sql := Sql + ' ' + Plus;
    TobCO.LoadDetailDBFromSQL('COMMUN', Sql);
    { Crée les items }
    for i := 0 to Pred(TobCO.Detail.Count) do
    begin
      SetLength(T, Length(T) + 1);
      T[Pred(Length(T))] := TMenuItem.Create(PopUp);
      with T[Pred(Length(T))] do
      begin
        Name := TobCO.Detail[i].GetString('CO_CODE');
        Caption := TobCO.Detail[i].GetString('CO_LIBELLE');
        OnClick := AOnClick;
      end;
    end;
    PopUp.Items.Add(T);
  finally
    TobCO.Free;
  end;
end;


function wTrue: char;
begin
  Result := BoolToStr_(True)[1];
end;

function wFalse: char;
begin
  Result := BoolToStr_(False)[1];
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 30/03/2005
Modifié le ... :   /  /    
Description .. : Renvoie le rang de la tob en fonction de la clef et de la 
Suite ........ : valeur.
Suite ........ : > Utilise la dichotomie
Suite ........ : > La tob doit être triée selon la clef
Suite ........ : > Retourne nil si non trouvé
Mots clefs ... :
*****************************************************************}
function MiddleIndex(TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant; MinIndex, MaxIndex: integer): integer;
var
  Trouve   : Boolean;
  i, Middle: integer;
begin
  if MinIndex <= MaxIndex then
  begin
    Middle := (MinIndex + MaxIndex) div 2;
    Trouve := True;
    for i := Low(FieldsName) to High(FieldsName) do
    begin
      if FieldsValue[i] < TheTob.Detail[Middle].GetValue(FieldsName[i]) then
      begin
        MaxIndex := Middle - 1;
        Trouve := false;
        Break;
      end
      else if FieldsValue[i] > TheTob.Detail[Middle].GetValue(FieldsName[i]) then
      begin
        MinIndex := Middle + 1;
        Trouve := false;
        Break;
      end;
    end;

    if Trouve then
      Result := Middle
    else
      Result := MiddleIndex(TheTob, FieldsName, FieldsValue, MinIndex, MaxIndex);
  end
  else
    Result := -1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 22/03/2006
Modifié le ... : 22/03/2006
Description .. : Renvoie le rang de la tob de la clef recherchée sur la
Suite ........ : méthode de recherche par dichotomie.
Suite ........ : Attention au clef homonymiques où cette fonction renvoie 
Suite ........ : un des rang et non pas le premier.
Suite ........ : 
Suite ........ : Voir Find1erRangOnTobSorted
Suite ........ : 
Mots clefs ... : 
*****************************************************************}
function FindRangOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): integer;
begin
  Result := MiddleIndex(TheTob, FieldsName, FieldsValue, 0, TheTob.Detail.Count-1)
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 22/03/2006
Modifié le ... : 22/03/2006
Description .. : Cette fonction renvoie la 1ere Tob 
Suite ........ : sur la clef demandée dans le cas où la clef est 
Suite ........ : homonymique.
Mots clefs ... : 
*****************************************************************}
function Find1erTobOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): Tob;
var
  iCpt, iRang, iNbChamps    : integer;
  lMemeClef                 : boolean;
begin
  Result := nil;
  iRang := FindRangOnTobSorted(TheTob, FieldsName, FieldsValue);
  {Recherche de la première ligne de la tob sur la clef}
  if (iRang<>-1) then
  begin
    for iCpt:=iRang downto 0 do
    begin
      lMemeClef:=True;
    	for iNbChamps := Low(FieldsName) to High(FieldsName) do
      begin
        lMemeClef := lMemeClef and (TheTob.Detail[iCpt].GetValue(FieldsName[iNbChamps])=FieldsValue[iNbChamps]);
	    end;
      if lMemeClef then
        Result := TheTob.Detail[iCpt]
      else
        break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 22/03/2006
Modifié le ... : 22/03/2006
Description .. : Cette fonction renvoie le rang de la première ligne de la Tob 
Suite ........ : sur la clef demandée dans le cas où la clef est 
Suite ........ : homonymique.
Mots clefs ... : 
*****************************************************************}
function Find1erRangOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): integer;
var
  iCpt, iRang, iNbChamps    : integer;
  lMemeClef                 : boolean;
begin
  iRang := FindRangOnTobSorted(TheTob, FieldsName, FieldsValue);
  {Recherche de la première ligne de la tob sur la clef}
  if (iRang<>-1) then
  begin
    for iCpt:=iRang downto 0 do
    begin
      lMemeClef:=True;
    	for iNbChamps := Low(FieldsName) to High(FieldsName) do
      begin
        lMemeClef := lMemeClef and (TheTob.Detail[iCpt].GetValue(FieldsName[iNbChamps])=FieldsValue[iNbChamps]);
	    end;
      if not lMemeClef then
        break;
    end;
    iRang := iCpt+1;
  end;
  Result := iRang;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marc MORRETTON
Créé le ...... : 22/03/2006
Modifié le ... : 22/03/2006
Description .. : Renvoie la tob de la clef recherchée sur la méthode de 
Suite ........ : recherche par dichotomie
Mots clefs ... : 
*****************************************************************}
function FindOnTobSorted(Const TheTob: Tob; Const FieldsName: array of string; FieldsValue: array of Variant): Tob;
var
  Rang: integer;
begin
  Rang := MiddleIndex(TheTob, FieldsName, FieldsValue, 0, TheTob.Detail.Count-1);
  if Rang = -1 then
    Result := nil
  else
    Result := TheTob.Detail[Rang];
end;

procedure TobRealToVirtual(TobData: Tob; var TobVirtual: Tob);
var
  iChamp    : integer;
begin
  if TobData.NumTable > 0 then
  begin
    { Recopie des champs réels }
    for iChamp := 1 to TobData.NbChamps do
      TobVirtual.AddChampSupValeur(TobData.GetNomChamp(iChamp), TobData.GetValue(TobData.GetNomChamp(iChamp)));
    { Recopie des champs supp. }
    for iChamp := 1000 to Pred(1000 + TobData.ChampsSup.Count) do
      TobVirtual.AddChampSupValeur(TobData.GetNomChamp(iChamp), TobData.GetValeur(iChamp))
  end
  else
    TobVirtual.Dupliquer(TobData, false, true);
end;

procedure GetCaptionByFocusControl(Container: TWinControl; T: Tob);
var
  i, NbrFind, j: Integer;

  function GetTobIndex(ControlName: String): Integer;
  var
    i: Integer;
  begin
    i := -1; Result := -1;
    while (i < T.Detail.Count - 1) and (Result = -1) do
    begin
      Inc(i);
      if (T.Detail[i].FieldExists('CONTROLNAME')) and (UpperCase(T.Detail[i].GetString('CONTROLNAME')) = UpperCase(ControlName)) then
        Result := i;
    end;
  end;

begin
  { Init. des résultats }
  for i := 0 to T.Detail.Count - 1 do
  begin
    if T.Detail[i].FieldExists('CONTROLNAME') then
      T.Detail[i].SetString('CONTROLNAME', UpperCase(T.Detail[i].GetString('CONTROLNAME')));
  end;
  { Recherche des contrôles dans le container }
  i := -1; NbrFind := 0;
  while (i < Container.ControlCount - 1) and (NbrFind < T.Detail.Count) do
  begin
    Inc(i);
    if (Container.Controls[i] is TWinControl) and (TWinControl(Container.Controls[i]).ControlCount > 1)
                                               or (Container.Controls[i] is ThPanel) then
    begin
  		if V_PGI.Debug then
      	Trace.TraceInformation('INFO',TWinControl(Container.Controls[i]).Name + ' ->');
      GetCaptionByFocusControl(TWinControl(Container.Controls[i]), T);
    end
    else if (Container.Controls[i] is ThLabel) then
    begin
		  if V_PGI.Debug then
    	  Trace.TraceInformation('INFO',TWinControl(Container.Controls[i]).Name + ' ThLabel');
      if Assigned(ThLabel(Container.Controls[i]).FocusControl) then
      begin
        j := GetTobIndex(ThLabel(Container.Controls[i]).FocusControl.Name);
        if j <> -1 then
        begin
				  if V_PGI.Debug then
        	  Trace.TraceInformation('INFO',TWinControl(Container.Controls[i]).Name + ' Trouvé');
          T.Detail[j].AddChampSupValeur('CAPTION', StringReplace(ThLabel(Container.Controls[i]).Caption, '&', '', [rfreplaceall]));
          Inc(NbrFind);
        end;
      end;
    end;
  end;
end;

procedure wStringToTString(Const Chaine: string; Result: TStrings; const Sep: String = ';');
var
  s1, s2: String;
begin
  if Assigned(Result) then
  begin
    s1 := Chaine;
    while (s1 <> '') do
    begin
      s2 := READTOKENPipe(s1, Sep);
      if s2 <> '' then
        Result.Add(s2);
    end;
  end;
end;

function wTStringToString(TString: TStrings; const Sep: String = ';'): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to Pred(TString.Count) do
    Result := Result + iif(Result <> '', Sep, '') + TString[i]
end;

function wIntToC2(i: Integer): String;
begin
  if i < 10 then
    Result := '0';
  Result := Result + IntToStr(i);
end;

{ TMemoFindDialog }

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure TMemoFindDialog.Find;
var
  PrevFoundAt,
  FoundAt : LongInt;
  i,
  CountCar,
  StartPos: Integer;
  Opts    : TSearchTypes;
  St      : String;

  function IsValidChar(const c: Char): Boolean;
  begin
    Result := not ((c >= '0') and (c <= '9'))
              and not ((c >= 'a') and (c <= 'z'))
              and not ((c >= 'A') and (c <= 'Z'))
  end;

begin
  FStringToFind := FindText;
  if Assigned(Memo) and Memo.Visible then
  begin
    with Memo do
    begin
      StartPos := SelStart + iif(frDown in Options, SelLength, 0);

      Opts := [];
      if frWholeWord in Options then
        Opts := Opts + [stWholeWord];
      if frMatchCase in Options then
        Opts := Opts + [stMatchCase];

      St := Text;
      FoundAt := 0;
      repeat
        PrevFoundAt := FoundAt;
        if FoundAt <> 0 then
        begin
          St := StringReplace(St, FindText, StringOfChar('@', Length(FindText)), [rfIgnoreCase]);
          if ((frDown in Options) and (StartPos < FoundAt)) then
            StartPos := FoundAt;
        end;
        FoundAt := iif(not (stMatchCase in Opts), Pos(UpperCase(FindText), UpperCase(St)), Pos(FindText, St));
        if  (FoundAt <> 0)
        and (stWholeWord in Opts)
        and (((FoundAt > 1) and not IsValidChar(St[FoundAt - 1])) or ((FoundAt + Length(FindText) - 1 < Length(St)) and not IsValidChar(St[FoundAt + Length(FindText)]))) then
          FoundAt := -1
      until (((frDown in Options) and (FoundAt > StartPos))
      or    (not (frDown in Options) and (FoundAt >= StartPos))) or (FoundAt = 0);

      if ((frDown in Options) and (FoundAt > 0))
      or (not (frDown in Options) and (PrevFoundAt > 0))then
      begin
        SelStart := iif(frDown in Options, FoundAt - 1, PrevFoundAt - 1);
        SelLength := Length(FindText);
        CountCar := 0;
        i := 0;
        while (CountCar < SelStart) and (i < Lines.Count) do
        begin
          CountCar := CountCar + Length(Lines[i]);
          Inc(i)
        end;
        if CountCar >= SelStart then
        begin
          Dec(i);
          Perform(EM_LINEINDEX, i, 0);
          Perform(EM_LINEFROMCHAR, CountCar - SelStart + SelLength, 0);
          Perform(EM_SCROLLCARET, 0, 0);
          if Assigned(FOnFound) then
            FOnFound(Memo)
        end
      end
      else
        PGIInfo(TraduireMemoire('Impossible de trouver') + ' : "' + FindText + '".', TraduireMemoire('Recherche'));
    end
  end
  else
    CloseDialog;

  inherited;
end;

procedure TMemoFindDialog.FindNext;
begin
  Find()
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{ Utils }
procedure BeginPrint(const NbCopies: Integer = 1; const SilentMode: Boolean = False);
begin
{$IFNDEF EAGLSERVER}
  { Sauvegarde des paramèters Pgi par défaut }
  with PrinterOptions do
  begin
    OldPrinterIndex  := Printer.PrinterIndex;
    OldNbDocCopies   := V_Pgi.NbDocCopies;
    OldSilentMode    := V_Pgi.SilentMode;
    OldNoPrintDialog := V_Pgi.NoPrintDialog;
  end;

  V_Pgi.NbDocCopies := NbCopies;
  if SilentMode then
  begin
    V_Pgi.NoPrintDialog := True;
    V_Pgi.SilentMode := True;
    V_Pgi.FirstHShowMessage := '';
    V_Pgi.LastHShowMessage := '';
  end
{$ENDIF !EAGLSERVER}
end;

procedure EndPrint;
begin
{$IFNDEF EAGLSERVER}
  { Restitution des paramèters Pgi par défaut }
  with PrinterOptions do
  begin
    V_Pgi.NoPrintDialog := OldNoPrintDialog;
    V_Pgi.NbDocCopies := OldNbDocCopies;
    V_Pgi.SilentMode := OldSilentMode;
    Printer.PrinterIndex := OldPrinterIndex;
  end
{$ENDIF !EAGLSERVER}
end;

Procedure SelectIndexPrinter(const PrinterName: hString);
begin
  if PrinterName <> '' then
    Printer.PrinterIndex := Printer.Printers.IndexOf(PrinterName);
end;

function TStringsToSqlIn(TS: TStrings): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to TS.Count - 1 do
  begin
    if Result <> '' then
      Result := Result + ',';
    Result := Result + '"' + Trim(TS[i]) + '"';
  end;
  if Result <> '' then
    Result := '(' + Result + ')';

end;

function TobToSqlIn(aTob: Tob; FieldName: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 0 to aTob.Detail.Count - 1 do
  begin
    if Result <> '' then
      Result := Result + ',';
    Result := Result + '"' + aTob.Detail[i].GetString(FieldName) + '"';
  end;
  if Result <> '' then
    Result := '(' + Result + ')';
end;

{-------------------------------------------------------------------------------
  Contrsuit une string à partir d'un double en respectant le signe
  afin d'éviter +-10 pour un nombre négatif
--------------------------------------------------------------------------------}
function MyDoubleToStr(MyDouble: Double; const Remove: Boolean = false): string;
begin
  if Remove then
    MyDouble := -MyDouble;
  Result := iif(MyDouble >= 0, '+', '') + StrFPoint(MyDouble)
end;

{$IFDEF EAGLSERVER}
function wActionToString(const Action: TActionFiche): String;
begin
  case Action of
    taCreat,
    taCreatEnSerie,
    taCreatOne      : Result := 'CREATION';
    taModif,
    taModifEnSerie  : Result := 'MODIFICATION';
    taDuplique      : result := 'DUPLICATION';
  else
    Result := 'CONSULTATION'
  end
end;
{$ENDIF EAGLSERVER}

function GetTokenFieldName(const Token: String): String;
begin
  Result := UpperCase(Trim(wLeft(Token, Pos('=', Token) - 1)))
end;

//GP_20080204_TS_GP14791 >>>
{ Utilisation des Plugins
  - Compilés et utilisables uniquement via un client Web Access ! }

{$IF Defined(EAGLCLIENT) and not Defined(EAGLSERVER)}
{ - Verbes génériques }
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 14/01/2008
Modifié le ... :   /  /    
Description .. : Exécution d'un verbe de Plugin Serveur.
Suite ........ : Fonction générique d'encapsulation des "requêtes" serveur 
Suite ........ : (non au sens SQL).
Suite ........ : ClassName : Nom de la classe au sens plugin (préfixe de la fonction)
Suite ........ : FunctionName : Nom de la fonction serveur à exécuter
Suite ........ : Retour : TobResponse.
Mots clefs ... : PLUGIN;
*****************************************************************}
function wExecutePluginVerb(const ClassName, FunctionName, VerbParam: String; FunctionParams: twPluginParams;
                            const ControlDllFonctionName: String = ''): Tob;
var
  TobRequest: Tob;
  i: Integer;

  function CSTTypeToVarType(const varType: TCSType): Word;
  begin
    case varType of
      CSTInteger: Result := varInteger;
      CSTDouble : Result := varDouble;
      CSTDate   : Result := varDate;
      CSTBoolean: Result := varBoolean;
      CSTString : Result := varString;
    else
      Result := varVariant
    end
  end;

const
  DllLoaded : Boolean = False;
begin
  TobRequest := Tob.Create('_Request_', nil, -1);
  try
    Result := nil;

    if not DllLoaded and (ControlDllFonctionName <> '') then
    begin
      Result := wExecutePluginVerb(ClassName, ControlDllFonctionName, '', nil);
      DllLoaded := Assigned(Result);
      if Assigned(Result) then
        FreeAndNil(Result);
    end;

    if DllLoaded or (ControlDllFonctionName = '') then
    begin
      if Length(FunctionParams) > 0 then
        for i := Low(FunctionParams) to High(FunctionParams) do
          with FunctionParams[i] do
            TobRequest.AddChampSupValeur(varName, VarAsType(varValue, CSTTypeToVarType(varType)), varType);

      Result := AppServer.Request(ClassName + '.' + FunctionName, VerbParam, TobRequest, '', '');
    end
  finally
    TobRequest.Free
  end
end;
{$IFEND EAGLCLIENT and not(EAGLSERVER)}

{ - Verbes communs }
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 14/01/2008
Modifié le ... :   /  /
Description .. : Exécution d'une requête SQL sur un DataBase clonée afin
Suite ........ : d'éviter les Locks en transaction.
Suite ........ : A utiliser à bon escient !
Mots clefs ... : SQL;TRANSACTION
*****************************************************************}
function wExecuteSqlOnClonedDB(const Sql: String): Integer;
{$IF not Defined(EAGLCLIENT) or Defined(EAGLSERVER)}
var
  ClonedDB: TDataBase;
{$ELSE  !EAGLCLIENT or EAGLSERVER}
  //TobResponse: Tob;
  //Params: twPluginParams;
{$IFEND !EAGLCLIENT or EAGLSERVER}
begin
{$IF not Defined(EAGLCLIENT) or Defined(EAGLSERVER)}
  if V_PGI.NbTransaction >= 1 then
  begin
    { En 2 Tiers ou EAGLSERVER, Clonage de la DB ... }
    ClonedDB := CloneConnection(DBSOC);
    try
      if Assigned(ClonedDB) then
      begin
        { ... et exécution de la requête sur cette DB }
        Result := ExecuteSQLDB(Sql, ClonedDB);
      end
      else
      begin
        if V_Pgi.Sav then
          PgiError('ClonedDB not assigned !', 'Erreur système (non bloquante)');
        Result := ExecuteSQL(Sql)
      end;
    finally
      if Assigned(ClonedDB) then
        ClonedDB.Free
    end;
  end
  else
    Result := ExecuteSql(Sql)
{$ELSE  !EAGLCLIENT or EAGLSERVER}
  Result := ExecuteSql(Sql)
{$IFEND !EAGLCLIENT or EAGLSERVER}
(* { En client Web Access, exécution de l'équivalent 2 Tiers via un plugin }
SetLength(Params, 1);
with Params[0] do
begin
  varName := 'SQL';
  varType := CSTString;
  varValue := Sql;
end;
TobResponse := wExecutePluginVerb('GPCommon', 'ExecuteSqlOnClonedDB', '', Params, 'PluginLoaded');
//'SQL=' + Sql, 'PluginLoaded');
try
  if Assigned(TobResponse) then
  begin
    if TobResponse.GetString('Error') <> '' then
    begin
      Result := 0;
      PgiError(TobResponse.GetString('Error'))
    end
    else
      Result := TobResponse.GetInteger('Result');
  end
  else
  begin
    if V_Pgi.Sav then
      PgiError(Format(TraduireMemoire('Plugin %s non chargé !'), ['GPPluginCommon.dll']));
    { Si le plugin n'existe pas ou n'est pas chargé, on ne doit pas bloquer le fonctionnel }
    Result := ExecuteSQL(Sql)
  end
finally
  if Assigned(TobResponse) then
    TobResponse.Free;
end; *)
end;

{ Récupère un message d'erreur commun à tous pour signaler une interruption
  de traitement liée à une exception dze transaction }
function TransactionErrorMessage: String;
begin
  if V_Pgi.Sav and (V_Pgi.NbTransaction > 0) and (V_Pgi.TransacErrorMessage <> '') then
    Result := V_Pgi.TransacErrorMessage
  else
    Result := TraduireMemoire('Traitement interrompu en cours de transaction.') + #13#10
            + TraduireMemoire('Veuillez relancer le traitement.')
end;
//GP_20080204_TS_GP14791 <<<

{ WException }
constructor ExceptionW.Create(const Msg: string);
begin
  V_Pgi.TransacErrorMessage := Msg;

  inherited Create(Msg)
end;

function GetWindowsCaptionFromArgument(Argument, DefaultValue : string) : string;
begin
  if Argument <> '' then
    Result := GetArgumentString(Argument, 'WINDOWCAPTION', False);
  if Result = '' then
    Result := DefaultValue;
end;

{$IFNDEF EAGLSERVER}
Initialization
  ProgressFormImbrication := 0;
{$ENDIF !EAGLSERVER}
end.

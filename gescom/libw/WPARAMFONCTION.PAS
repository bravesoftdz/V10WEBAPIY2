{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 08/10/2002
Modifié le ... :   /  /
Description .. : Paramétrage des enchaînements d'actions
Mots clefs ... :
*****************************************************************}
unit wParamFonction;

interface

uses
	Windows,
  Classes,
  SysUtils,
  Hctrls,
  Hent1,
  uTob,
  Dialogs,
  Menus,
  Messages,
  Hmsgbox,
  wCommuns,
  {$IFNDEF EAGLSERVER}
    HTB97,
    Controls,
  {$ENDIF !EAGLSERVER}
  EntPgi
  ;

type
  tCleWPF = record
            { Général }
            User          : string;
            Contexte      : string;
            Prefixe       : string;

            { Commercial }
            Secteur           : string;
            SocieteGroupe     : string;
            Tiers             : string;
            TarifTiers        : string;
            TableLibreTiers   : Array[1..10] of string;

            { Pièce }
            VenteAchat    : string;
            NaturePieceG  : String;

            { Lieu }
            Etablissement : string;
            Depot         : string;

            { Production }
            NatureTravail : string;
            TypeOrdre     : string;

            { Scm }
            TypComp       : string;
            FamilleMat    : string;

            { Process }
            CodIti        : string;
            Circuit       : string;
            Pole          : string;
            Site          : string;
            Grp           : string;
            PhasePred     : string;
            Phase         : string;
            PhaseNext     : string;
            OpeIti        : string;

            { Mouvements }
            StkFlux       : string;
            QualifMvt     : string;

            { Article }
            Article       : string;
            CodeArticle   : string; //Sert uniquement pour gestion WPARAMFONCTION_TOM (Contrôle de saisie => pour éviter les doublons)
//GP_20080604_MM_GP12663 Déb
            FamilleNiv    : Array[1..3] of string;
            FamilleValo   : string;
            Collection    : string;
//GP_20080604_MM_GP12663 Fin
//GP_20080930_TS_GP15021
            MethProd      : String;
            LibreArt      : Array[1..10] of string;
            CodeUL        : string;

            { Tarifs }
            FormuleMt     : string;

            { Quotas fournisseurs }
            CodeMarche    : string;

            { Situation géographique }
            Pays          : string;
            Region        : string;
            CodePostal    : string;
            SecteurGeo    : string;

            { Prix de revient : consolidation }
            TypePdr       : string;
            NomChamp      : string;
//GP_20080527_MM_GP12663 Déb

            { Coût indirect }
            CodePort      : string;

            { Ressources }
            Ressource     : string;
            TypeRessource : string;
            TarifRessource: string;
//GP_20080527_MM_GP12663 Fin
           end;
  PCleWPF = ^TCleWPF;

  tDataWPF = record
               MyBoolean : Array[1..15] of boolean;
               MyInteger : Array[1..10] of integer;
               MyDouble  : Array[1..10] of double;
               MyDateTime: Array[1..10] of tDateTime;
               MyCombo   : Array[1..10] of String;
               MyString  : Array[1..12] of String;
               BlocNote  : String;
               CleWPF    : TCleWPF;
             end;
  PDataWPF = ^TDataWPF;

  { Gestion mémoire de wParamFonction }
  TWPF = class(TMemoryTob)
  private
    FCodeFonction,
    FSelect,
    FOrder, FGroup: String;
    FDataFields,
    FCritFields,
    FOrderFields,
    FGroupFields: MyArrayField;
    FActive: Boolean;

    function BuildSelect: String;
    function BuildSql: String;
    function BuildOrder: String;
    function BuildGroup: String;
    function FindInWPF(const CleWPF: PCleWPF; const TestInitValue: Boolean): Tob;
    function HasFieldIn(const FieldName: String; const Fields: MyArrayField): Boolean;
    procedure SetActive(const Value: Boolean);
  public
    procedure Load(Const AvertirCache: Boolean = False); override;
    function HasFieldInOrder(const FieldName: String; out SortType: String): Boolean;
    function HasFieldInWhere(const FieldName: String): Boolean;
    function HasFieldInSelect(const FieldName: String): Boolean;
    function FindWPF(const CleWPF: PCleWPF; var DataWPF: TDataWPF; const TestInitValue: Boolean; TobResultWPF: Tob): Boolean; overload;
    function FindWPF(const CleWPF: PCleWPF; const TestInitValue: Boolean = False): Boolean; overload;

    constructor Create(const CodeFonction: String; Select, Where: String;
                       Order: String = ''; Group: String = ''); reintroduce;
    destructor Destroy; override;

    property CodeFonction: String read FCodeFonction;
    property OrderBy: String read FOrder write FOrder;
    property Active: Boolean read FActive write SetActive;
  end;

  TParamFonction = class
  private
    FLastCodeFonction: String;
    FLastWPF: TWPF;
    FList: TList;
    function GetAWPF(CodeFonction: String): TWPF;
    procedure RemoveWPF(const Index: Integer); overload;
  public
    procedure Clear;
    function Count: Integer;
    procedure Initialisation;
    function Exists(const CodeFonction: String): Boolean;
    function AddWPF(const CodeFonction, Select, Where: String; const Order: String = ''; const Group: String = ''): TWPF;
    procedure RemoveWPF(const CodeFonction: String); overload;

    constructor Create;
    Destructor Destroy; override;

    property WPF[CodeFonction: String]: TWPF read GetAWPF; default;
  end;

{$IFNDEF EAGLSERVER}
  TWPFButton = class(TToolBarButton97)
  private
    FControler: TWinControl;
    FCodeFonction: String;
    FArgument: String;
    procedure SetControler(const Value: TWinControl);
    procedure WPFButtonClick(Sender: TObject);
  published
  public
    constructor Create(const AControler: TWinControl; const ACodeFonction, AHint: String;
                       const AArgument: String = ''); reintroduce;
    destructor Destroy; override;

    class function Add(const Controler: TWinControl;
                       const CodeFonction, Hint: String;
                       const Argument: String = ''): TWPFButton;

    property Controler: TWinControl read FControler write SetControler;
    property CodeFonction: String read FCodeFonction;
    property Argument: String read FArgument write FArgument;
  end;
{$ENDIF !EAGLSERVER}

function WhereWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; Const TestInitValue: boolean=false): String;
function ExistWPF(Const CodeFonction: string): Boolean; overload;
function ExistWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; Const TestInitValue: boolean=false): Boolean; overload;
function GetWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; var DataWPF: tDataWPF; const WithBlocNote: Boolean = False; const TestInitValue: Boolean = True): Boolean;
function GetTobWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; var DataWPF: tDataWPF; TobResultWPF: Tob; const WithBlocNote: Boolean = False; const TestInitValue: Boolean = True): Boolean;
function OrderByWPF(Const CodeFonction: string = ''): String;
function wIniWPF(Const CodeFonction: string; Const CleWPF: TCleWPF; Const ChampsInitWPF: MyArrayField = nil; Const ValeursChampsInitWPF: MyArrayValue = nil): Boolean;
function GetCleWPF(Const Argument: string; const Separator: String = ';'): tCleWPF;
//Result type String car aujourd'hui il n'y a que des champs texte à tester
function GetValueCleWPF(const CleWPF: PCleWPF; const FieldName: String): String;
procedure AssignDataWPFFromTobWPF(const CodeFonction: String; var DataWPF: TDataWPF; const TobWPF: Tob);

implementation

uses
  EntGC,
  Math,
  UtilPGI,
  {$IFNDEF EAGLSERVER}
    {$IFDEF EAGLCLIENT}
      MainEAgl,
    {$ELSE  EAGLCLIENT}
      FE_Main,
    {$ENDIF EAGLCLIENT}
    StdCtrls,
  {$ENDIF !EAGLSERVER}
  uTom
  ;

Const
  TableName = 'WPARAMFONCTION';

function OrderByWPF(Const CodeFonction: string = ''): String;
var
  WPF: TWPF;
  SortType: String;

  { <-- Gestion mémoire de WPF --> }
  function isInOrder(const FieldName: String): Boolean;
  begin
    Result := not Assigned(WPF) or not WPF.Active or WPF.HasFieldInOrder(FieldName, SortType)
  end;

begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];
  if Assigned(WPF) and (WPF.OrderBy <> '') then
    Result := WPF.OrderBy
  else
  begin
    SortType := 'DESC';
    Result := iif(isInOrder('WPF_ARTICLE')       , ', ISNULL(WPF_ARTICLE, " ") '       + SortType, '')
            + iif(isInOrder('WPF_CODEARTICLE')   , ', ISNULL(WPF_CODEARTICLE, " ") '   + SortType, '')
//GP_20080604_MM_GP12663 Déb
            + iif(isInOrder('WPF_FAMILLEVALO')   , ', ISNULL(WPF_FAMILLEVALO, " ") '   + SortType, '')
//GP_20080930_TS_GP15021
            + iif(isInOrder('WPF_METHPROD')      , ', ISNULL(WPF_METHPROD, " ") '      + SortType, '')
            + iif(isInOrder('WPF_FAMILLENIV3')   , ', ISNULL(WPF_FAMILLENIV3, " ") '   + SortType, '')
            + iif(isInOrder('WPF_FAMILLENIV2')   , ', ISNULL(WPF_FAMILLENIV2, " ") '   + SortType, '')
            + iif(isInOrder('WPF_FAMILLENIV1')   , ', ISNULL(WPF_FAMILLENIV1, " ") '   + SortType, '')
            + iif(isInOrder('WPF_COLLECTION')    , ', ISNULL(WPF_COLLECTION, " ") '    + SortType, '')
//GP_20080604_MM_GP12663 Fin
            + iif(isInOrder('WPF_LIBREARTA')     , ', ISNULL(WPF_LIBREARTA, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART9')     , ', ISNULL(WPF_LIBREART9, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART8')     , ', ISNULL(WPF_LIBREART8, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART7')     , ', ISNULL(WPF_LIBREART7, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART6')     , ', ISNULL(WPF_LIBREART6, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART5')     , ', ISNULL(WPF_LIBREART5, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART4')     , ', ISNULL(WPF_LIBREART4, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART3')     , ', ISNULL(WPF_LIBREART3, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART2')     , ', ISNULL(WPF_LIBREART2, " ") '     + SortType, '')
            + iif(isInOrder('WPF_LIBREART1')     , ', ISNULL(WPF_LIBREART1, " ") '     + SortType, '')
            + iif(isInOrder('WPF_TIERS')         , ', ISNULL(WPF_TIERS, " ") '         + SortType, '')
            + iif(isInOrder('WPF_TARIFTIERS')    , ', ISNULL(WPF_TARIFTIERS, " ") '    + SortType, '')
            + iif(isInOrder('WPF_QUALIFMVT')     , ', ISNULL(WPF_QUALIFMVT, " ") '     + SortType, '')
            + iif(isInOrder('WPF_STKFLUX')       , ', ISNULL(WPF_STKFLUX, " ") '       + SortType, '')
            + iif(isInOrder('WPF_PHASENEXT')     , ', ISNULL(WPF_PHASENEXT, " ") '     + SortType, '')
            + iif(isInOrder('WPF_PHASE')         , ', ISNULL(WPF_PHASE, " ") '         + SortType, '')
            + iif(isInOrder('WPF_PHASEPRED')     , ', ISNULL(WPF_PHASEPRED, " ") '     + SortType, '')
            + iif(isInOrder('WPF_OPEITI')        , ', ISNULL(WPF_OPEITI, " ") '        + SortType, '')
            + iif(isInOrder('WPF_GRP')           , ', ISNULL(WPF_GRP, " ") '           + SortType, '')
            + iif(isInOrder('WPF_SITE')          , ', ISNULL(WPF_SITE, " ") '          + SortType, '')
            + iif(isInOrder('WPF_CIRCUIT')       , ', ISNULL(WPF_CIRCUIT, " ") '       + SortType, '')
            + iif(isInOrder('WPF_CODITI')        , ', ISNULL(WPF_CODITI, " ") '        + SortType, '')
            + iif(isInOrder('WPF_TYPEORDRE')     , ', ISNULL(WPF_TYPEORDRE, " ") '     + SortType, '')
            + iif(isInOrder('WPF_NATURETRAVAIL') , ', ISNULL(WPF_NATURETRAVAIL, " ") ' + SortType, '')
            + iif(isInOrder('WPF_FAMILLEMAT')    , ', ISNULL(WPF_FAMILLEMAT, " ") '    + SortType, '')
            + iif(isInOrder('WPF_TYPCOMP')       , ', ISNULL(WPF_TYPCOMP, " ") '       + SortType, '')
            + iif(isInOrder('WPF_DEPOT')         , ', ISNULL(WPF_DEPOT, " ") '         + SortType, '')
            + iif(isInOrder('WPF_NATUREPIECEG')  , ', ISNULL(WPF_NATUREPIECEG, " ") '  + SortType, '')
            + iif(isInOrder('WPF_VENTEACHAT')    , ', ISNULL(WPF_VENTEACHAT, " ") '    + SortType, '')
            + iif(isInOrder('WPF_ETABLISSEMENT') , ', ISNULL(WPF_ETABLISSEMENT, " ") ' + SortType, '')
            + iif(isInOrder('WPF_SOCIETEGROUPE') , ', ISNULL(WPF_SOCIETEGROUPE, " ") ' + SortType, '')
            + iif(isInOrder('WPF_SECTEUR')       , ', ISNULL(WPF_SECTEUR, " ") '       + SortType, '')
            + iif(isInOrder('WPF_PREFIXE')       , ', ISNULL(WPF_PREFIXE, " ") '       + SortType, '')
            + iif(isInOrder('WPF_CONTEXTE')      , ', ISNULL(WPF_CONTEXTE, " ") '      + SortType, '')
            + iif(isInOrder('WPF_USER')          , ', ISNULL(WPF_USER, " ") '          + SortType, '')
            + iif(isInOrder('WPF_FORMULEMT')     , ', ISNULL(WPF_FORMULEMT, " ") '     + SortType, '')
            + iif(isInOrder('WPF_SECTEURGEO')    , ', ISNULL(WPF_SECTEURGEO, " ") '    + SortType, '')
            + iif(isInOrder('WPF_CODEPOSTAL')    , ', ISNULL(WPF_CODEPOSTAL, " ") '    + SortType, '')
            + iif(isInOrder('WPF_REGION')        , ', ISNULL(WPF_REGION, " ") '        + SortType, '')
            + iif(isInOrder('WPF_PAYS')          , ', ISNULL(WPF_PAYS, " ") '          + SortType, '')
            + iif(isInOrder('WPF_TYPEPDR')       , ', ISNULL(WPF_TYPEPDR, " ") '       + SortType, '')
            + iif(isInOrder('WPF_NOMCHAMP')      , ', ISNULL(WPF_NOMCHAMP, " ") '      + SortType, '')
//GP_20080527_MM_GP12663 Déb
            + iif(isInOrder('WPF_CODEPORT')      , ', ISNULL(WPF_CODEPORT, " ") '      + SortType, '')
            + iif(isInOrder('WPF_RESSOURCE')     , ', ISNULL(WPF_RESSOURCE, " ")'      + SortType, '')
            + iif(isInOrder('WPF_TYPERESSOURCE') , ', ISNULL(WPF_TYPERESSOURCE, " ")'  + SortType, '')
            + iif(isInOrder('WPF_TARIFRESSOURCE'), ', ISNULL(WPF_TARIFRESSOURCE, " ")' + SortType, '')
//GP_20080527_MM_GP12663 Fin
            ;
    { Retire le ', ' du début }
    if Result <> '' then
      System.Delete(Result, 1, 2);
    { <-- Gestion mémoire de WPF --> }
    if Assigned(WPF) then
      WPF.OrderBy := Result
  end
end;

function WhereWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; Const TestInitValue: boolean=false): String;
var
  i: integer;
  WPF: TWPF;

  function WhereFieldName(Const FieldName, Value: string): string;
  begin
    if not Assigned(WPF) or not WPF.Active or WPF.HasFieldInWhere(FieldName) then
    begin
      if Value <> '' then
        Result := ' AND (' + FieldName + '="' + Value + '"' + iif(TestInitValue, ' OR ISNULL(' + FieldName + ',"")=""', '') + ')'
      else
        Result := ' AND ISNULL(' + FieldName + ',"")=""'
    end
  end;

begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];

  with CleWPF do
  begin
    Result := 'WPF_CODEFONCTION="' + CodeFonction + '"';

    { Général }
    Result := Result
            + WhereFieldName('WPF_USER'    , User)
            + WhereFieldName('WPF_CONTEXTE', Contexte)
            + WhereFieldName('WPF_PREFIXE' , Prefixe)
            ;

    { Commercial }
    Result := Result
            + WhereFieldName('WPF_SECTEUR'      , Secteur)
            + WhereFieldName('WPF_SOCIETEGROUPE', SocieteGroupe)
            + WhereFieldName('WPF_TIERS'        , Tiers)
            + WhereFieldName('WPF_TARIFTIERS'   , TarifTiers)
            ;
    for i := 1 to 10 do
      Result := Result
              + WhereFieldName('WPF_TABLELIBRETIERS' + IntToStr(i), TableLibreTiers[i])
              ;

    { Pièces }
    Result := Result
            + WhereFieldName('WPF_VENTEACHAT'  , VenteAchat)
            + WhereFieldName('WPF_NATUREPIECEG', NaturePieceG)
            ;

    { Lieux }
    Result := Result
            + WhereFieldName('WPF_ETABLISSEMENT', Etablissement)
            + WhereFieldName('WPF_DEPOT'        , Depot)
            ;

    { Production }
    Result := Result
            + WhereFieldName('WPF_NATURETRAVAIL', NatureTravail)
            + WhereFieldName('WPF_TYPEORDRE'    , TypeOrdre)
            ;

    { Scm }
    Result := Result
            + WhereFieldName('WPF_FAMILLEMAT', FamilleMat)
            + WhereFieldName('WPF_TYPCOMP'   , TypComp)
            ;

    { Process }
    Result := Result
            + WhereFieldName('WPF_CODITI'   , CodIti)
            + iif(CodeFonction<>'CBNCIRCDEP',WhereFieldName('WPF_CIRCUIT'  , Circuit),'')
            + WhereFieldName('WPF_POLE'     , Pole)
            + WhereFieldName('WPF_SITE'     , Site)
            + WhereFieldName('WPF_GRP'      , Grp)
            + WhereFieldName('WPF_PHASEPRED', PhasePred)
            + WhereFieldName('WPF_PHASE'    , Phase)
            + WhereFieldName('WPF_PHASENEXT', PhaseNext)
            + WhereFieldName('WPF_OPEITI'   , OpeIti)
            ;

    { Mouvement }
    Result := Result
            + WhereFieldName('WPF_STKFLUX'  , StkFlux)
            + WhereFieldName('WPF_QUALIFMVT', QualifMvt)
            ;

    { Article }
//GP_20080604_MM_GP12663 Déb
    Result := Result
            + WhereFieldName('WPF_ARTICLE'     , Article)
            + WhereFieldName('WPF_CODEARTICLE' , iif(CodeArticle <> '', CodeArticle, Trim(wLeft(Article, 18))))
            + WhereFieldName('WPF_COLLECTION'  , Collection)
//GP_20080930_TS_GP15021
            + WhereFieldName('WPF_METHPROD'    , MethProd)
            + WhereFieldName('WPF_FAMILLEVALO' , FamilleValo)
            ;
    for i := 1 to 3 do
      Result := Result
              + WhereFieldName('WPF_FAMILLENIV' + IntToStr(i), FamilleNiv[i])
              ;
//GP_20080604_MM_GP12663 Fin
    for i := 1 to 10 do
      Result := Result
              + WhereFieldName('WPF_LIBREART' + IntToHex(i, 0), LibreArt[i])
              ;
    if CodeFonction = 'RCOLISAGE' then
    begin
      Result := Result
              + WhereFieldName('WPF_VARCHAR01', CodeUL)
    end;

    { Tarifs }
    Result := Result
            + WhereFieldName('WPF_FORMULEMT', FormuleMt)
            ;

    {Quotas fournisseurs }
    if CodeFonction = 'CBNQUOTAFO' then
    begin
      Result := Result
              + WhereFieldName('WPF_VARCHAR01', CodeMarche)
    end;

    { Situation géographique }
    Result := Result
            + WhereFieldName('WPF_PAYS', Pays)
            + WhereFieldName('WPF_REGION', Region)
            + WhereFieldName('WPF_CODEPOSTAL', CodePostal)
            + WhereFieldName('WPF_SECTEURGEO', SecteurGeo)
            ;

    { Prix de revient : consolidation }
    Result := Result
            + WhereFieldName('WPF_TYPEPDR' , TypePdr)
            + WhereFieldName('WPF_NOMCHAMP', NomChamp)
            ;

//GP_20080527_MM_GP12663 Déb
    { Coût indirect }
    Result := Result
            + WhereFieldName('WPF_CODEPORT', CodePort)
            ;

    { Ressource }
    Result := Result
            + WhereFieldName('WPF_RESSOURCE'     , Ressource)
            + WhereFieldName('WPF_TYPERESSOURCE' , TypeRessource)
            + WhereFieldName('WPF_TARIFRESSOURCE', TarifRessource)
            ;
//GP_20080527_MM_GP12663 Fin
  end;
end;

function ExistWPF(Const CodeFonction: string): Boolean;
var
  Sql: String;
  WPF: TWPF;
begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];
  if Assigned(WPF) and WPF.Active then
  begin
    WPF.Load();
    Result := WPF.Detail.Count > 0
  end
  else
  begin
    Sql := 'SELECT 1'
         + ' FROM ' + TableName
         + ' WHERE WPF_CODEFONCTION="' + CodeFonction + '"'
         ;
    Result := ExisteSQL(Sql)
  end
end;

function ExistWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; Const TestInitValue: boolean=false): Boolean;
var
  Sql: String;
  WPF: TWPF;
begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];
  if Assigned(WPF) and WPF.Active then
    Result := WPF.FindWPF(@CleWPF, TestInitValue)
  else
  begin
    Sql := 'SELECT 1'
         + ' FROM ' + TableName
         + ' WHERE ' + WhereWPF(CodeFonction, CleWPF, TestInitValue)
         ;
    Result := ExisteSQL(Sql);
  end
end;

procedure AssignDataWPFFromTobWPF(const CodeFonction: String; var DataWPF: TDataWPF; const TobWPF: Tob);
var
  WPF: TWPF;
  i  : Integer;
  sNumField: String;

  function isDataField(const FieldName: String): Boolean;
  begin
    Result := not Assigned(WPF) or not WPF.Active or WPF.HasFieldInSelect(FieldName)
  end;

  function isKeyField(const FieldName: String): Boolean;
  begin
    Result := not Assigned(WPF) or not WPF.Active or WPF.HasFieldInWhere(FieldName)
  end;

begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];

  with DataWPF do
  begin
    { Recopie des données trouvées }
    for i := 1 to 15 do
    begin
      sNumField := wRight('0' + IntToStr(i), 2);
      if isDataField('WPF_BOOLEAN'     + sNumField) then MyBoolean[i]  := TobWPF.GetBoolean ('WPF_BOOLEAN' + sNumField);
      if i <= 12 then
      begin
        if i <= 10 then
        begin
          if isDataField('WPF_INTEGER' + sNumField) then MyInteger[i]  := TobWPF.GetInteger ('WPF_INTEGER' + sNumField);
          if isDataField('WPF_DOUBLE'  + sNumField) then MyDouble[i]   := TobWPF.GetDouble  ('WPF_DOUBLE'  + sNumField);
          if isDataField('WPF_DATE'    + sNumField) then MyDateTime[i] := TobWPF.GetDateTime('WPF_DATE'    + sNumField);
          if isDataField('WPF_COMBO'   + sNumField) then MyCombo[i]    := TobWPF.GetString  ('WPF_COMBO'   + sNumField);
        end;
        if isDataField('WPF_VARCHAR'   + sNumField) then MyString[i]   := TobWPF.GetString  ('WPF_VARCHAR' + sNumField);
      end
    end;
    if isDataField('WPF_BLOCNOTE') then BlocNote := TobWPF.GetString('WPF_BLOCNOTE');

    { Recopie la clé trouvée }
    with TobWPF do
    begin
      { Général }
      if isKeyField('WPF_USER')     then CleWPF.User     := GetString('WPF_USER');
      if isKeyField('WPF_CONTEXTE') then CleWPF.Contexte := GetString('WPF_CONTEXTE');
      if isKeyField('WPF_PREFIXE')  then CleWPF.Prefixe  := GetString('WPF_PREFIXE');

      { Commercial }
      if isKeyField('WPF_SECTEUR')       then CleWPF.Secteur       := GetString('WPF_SECTEUR');
      if isKeyField('WPF_SOCIETEGROUPE') then CleWPF.SocieteGroupe := GetString('WPF_SOCIETEGROUPE');
      if isKeyField('WPF_TIERS')         then CleWPF.Tiers         := GetString('WPF_TIERS');
      if isKeyField('WPF_TARIFTIERS')    then CleWPF.TarifTiers    := GetString('WPF_TARIFTIERS');

      for i := 1 to 10 do
        if isKeyField('WPF_TABLELIBRETIERS' + IntToStr(i)) then CleWPF.TableLibreTiers[i] := GetString('WPF_TABLELIBRETIERS' + IntToStr(i));

      { Pièces }
      if isKeyField('WPF_VENTEACHAT')   then CleWPF.VenteAchat    := GetString('WPF_VENTEACHAT');
      if isKeyField('WPF_NATUREPIECEG') then CleWPF.NaturePieceG  := GetString('WPF_NATUREPIECEG');

      { Lieux }
      if isKeyField('WPF_ETABLISSEMENT') then CleWPF.Etablissement := GetString('WPF_ETABLISSEMENT');
      if isKeyField('WPF_DEPOT')         then CleWPF.Depot         := GetString('WPF_DEPOT');

      { Production }
      if isKeyField('WPF_NATURETRAVAIL') then CleWPF.NatureTravail := GetString('WPF_NATURETRAVAIL');
      if isKeyField('WPF_TYPEORDRE')     then CleWPF.TypeOrdre     := GetString('WPF_TYPEORDRE');

      { Scm }
      if isKeyField('WPF_TYPCOMP')    then CleWPF.TypComp    := GetString('WPF_TYPCOMP');
      if isKeyField('WPF_FAMILLEMAT') then CleWPF.FamilleMat := GetString('WPF_FAMILLEMAT');

      { Process }
      if isKeyField('WPF_CODITI')    then CleWPF.CodIti    := GetString('WPF_CODITI');
      if isKeyField('WPF_CIRCUIT')   then CleWPF.Circuit   := GetString('WPF_CIRCUIT');
      if isKeyField('WPF_POLE')      then CleWPF.Pole      := GetString('WPF_POLE');
      if isKeyField('WPF_SITE')      then CleWPF.Site      := GetString('WPF_SITE');
      if isKeyField('WPF_GRP')       then CleWPF.Grp       := GetString('WPF_GRP');
      if isKeyField('WPF_PHASEPRED') then CleWPF.PhasePred := GetString('WPF_PHASEPRED');
      if isKeyField('WPF_PHASE')     then CleWPF.Phase     := GetString('WPF_PHASE');
      if isKeyField('WPF_PHASENEXT') then CleWPF.PhaseNext := GetString('WPF_PHASENEXT');
      if isKeyField('WPF_OPEITI')    then CleWPF.OpeIti    := GetString('WPF_OPEITI');

      { Mouvements }
      if isKeyField('WPF_STKFLUX')   then CleWPF.StkFlux   := GetString('WPF_STKFLUX');
      if isKeyField('WPF_QUALIFMVT') then CleWPF.QualifMvt := GetString('WPF_QUALIFMVT');

      { Article }
//GP_20080604_MM_GP12663 Déb
      if isKeyField('WPF_ARTICLE') then CleWPF.Article    := GetString('WPF_ARTICLE');
      for i := 1 to 3 do
        if isKeyField('WPF_FAMILLENIV' + IntToStr(i))
          then CleWPF.FamilleNiv[i] := GetString('WPF_FAMILLENIV' + IntToStr(i));
      if isKeyField('WPF_FAMILLEVALO') then CleWPF.FamilleValo := GetString('WPF_FAMILLEVALO');
      if isKeyField('WPF_COLLECTION') then CleWPF.Collection := GetString('WPF_COLLECTION');
//GP_20080604_MM_GP12663 Fin
//GP_20080930_TS_GP15021
      if isKeyField('WPF_METHPROD') then CleWPF.MethProd := GetString('WPF_METHPROD');
      for i := 1 to 10 do
        if isKeyField('WPF_LIBREART' + IntToHex(i, 0)) then CleWPF.LibreArt[i] := GetString('WPF_LIBREART' + IntToHex(i, 0));

     if CodeFonction = 'RCOLISAGE' then
        if isKeyField('WPF_VARCHAR01') then CleWPF.CodeUL := GetString('WPF_VARCHAR01');

      { Formule de calcul de montant pour la recherche tarifaire }
      if isKeyField('WPF_FORMULEMT') then CleWPF.FormuleMt := GetString('WPF_FORMULEMT');

      { Quotas fournisseurs }
      if CodeFonction = 'CBNQUOTAFO' then
        if isKeyField('WPF_VARCHAR01') then CleWPF.CodeMarche := GetString('WPF_VARCHAR01');

      { Situation géographique }
      if isKeyField('WPF_PAYS') then CleWPF.Pays             := GetString('WPF_PAYS');
      if isKeyField('WPF_REGION') then CleWPF.Region         := GetString('WPF_REGION');
      if isKeyField('WPF_CODEPOSTAL') then CleWPF.CodePostal := GetString('WPF_CODEPOSTAL');
      if isKeyField('WPF_SECTEURGEO') then CleWPF.SecteurGeo := GetString('WPF_SECTEURGEO');

      { Prix de revient : consolidation }
      if isKeyField('WPF_TYPEPDR')  then CleWPF.TypePdr  := GetString('WPF_TYPEPDR');
      if isKeyField('WPF_NOMCHAMP') then CleWPF.NomChamp := GetString('WPF_NOMCHAMP');

//GP_20080527_MM_GP12663 Déb
      { Coût indirect  }
      if isKeyField('WPF_CODEPORT') then CleWPF.CodePort := GetString('WPF_CODEPORT');

      { Ressource }
      if isKeyField('WPF_RESSOURCE')      then CleWPF.Ressource      := GetString('WPF_RESSOURCE');
      if isKeyField('WPF_TYPERESSOURCE')  then CleWPF.TypeRessource  := GetString('WPF_TYPERESSOURCE');
      if isKeyField('WPF_TARIFRESSOURCE') then CleWPF.TarifRessource := GetString('WPF_TARIFRESSOURCE');
//GP_20080527_MM_GP12663 Fin
    end
  end
end;

function GetTobWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; var DataWPF: tDataWPF; TobResultWPF: Tob; const WithBlocNote: Boolean = False; const TestInitValue: Boolean = True): Boolean;
var
  Sql   : string;
  TobWPF: Tob;
  WPF   : TWPF;
begin
  { <-- Gestion mémoire de WPF --> }
  WPF := VH_GC.ParamFonction[CodeFonction];
  if Assigned(WPF) and WPF.Active then
    Result := WPF.FindWPF(@CleWPF, DataWPF, TestInitValue, TobResultWPF)
  else
  begin
    Result := false;
    Sql := '';
    if IsOracle then Sql := 'SELECT * FROM (';
    Sql := Sql
         + 'SELECT ' + iif(isOracle, '', '##TOP 1##')
         + ' ' + wMakeSelectPresqueEtoile(TableName, nil, not WithBlocNote)
         + ' FROM ' + TableName
         + ' WHERE ' + WhereWPF(CodeFonction, CleWPF, TestInitValue)
         + ' ORDER BY ' + OrderByWPF(CodeFonction)
         ;
    if IsOracle then Sql := Sql + ') WHERE ROWNUM <= 1';

    TobWPF := Tob.Create('WPF', nil, -1);
    try
      if TobWPF.LoadDetailDBFromSQL(TableName, Sql) then
      begin
        Result := true;
        AssignDataWPFFromTobWPF(CodeFonction, DataWPF, TobWPF.Detail[0]);

        if Assigned(TobResultWPF) then
          TobResultWPF.Dupliquer(TobWPF.Detail[0], False, True, False)
      end;
    finally
      TobWPF.Free;
    end;
  end
end;

function GetWPF(Const CodeFonction: string; Const CleWPF: tCleWPF; var DataWPF: tDataWPF; const WithBlocNote: Boolean = False; const TestInitValue: Boolean = True): Boolean;
begin
  Result := GetTobWPF(CodeFonction, CleWPF, DataWPF, nil, WithBlocNote, TestInitValue)
end;

function wIniWPF(Const CodeFonction: string; Const CleWPF: TCleWPF; Const ChampsInitWPF: MyArrayField = nil; Const ValeursChampsInitWPF: MyArrayValue = nil): Boolean;
var
  TobWPF: Tob;
  TomWPF: Tom;
  i     : Integer;
begin
  Result := false;
  if not ExistWPF(CodeFonction, CleWPF) then
  begin
    TobWPF := Tob.Create(TableName, nil, -1);
    TomWPF := CreateTOM(TableName, nil, false, true);
    try
      { NewRecord }
      TomWPF.InitTOB(TobWPF);

      { Key }
      TobWPF.SetString('WPF_CODEFONCTION', CodeFonction);

      { Autres champs pour Init. }
      for i := Low(ChampsInitWPF) to High(ChampsInitWPF) do
        TobWPF.AddChampSupValeur(ChampsInitWPF[i], ValeursChampsInitWPF[i], False);

      { UpDateRecord }
      TobWPF.AddChampSupValeur('IKC', 'C', false);
      Result := TomWPF.VerifTOB(TobWPF) and TobWPF.InsertDB(nil);
      if Result then
        TomWPF.AfterVerifTOB(TobWPF)
    finally
      TobWPF.free;
      TomWPF.free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/02/2006
Modifié le ... :   /  /    
Description .. : Renvoie une cleWPF en fonction d'un argument
Mots clefs ... : 
*****************************************************************}
function GetCleWPF(Const Argument: string; const Separator: String = ';'): tCleWPF;
var
  i: integer;

  function GetArgStr(const MyArg: String): String;
  begin
    Result := GetArgumentString(Argument, MyArg, True, Separator)
  end;

begin
  { Général }
  Result.User     := GetArgStr('WPF_USER');
  Result.Contexte := GetArgStr('WPF_CONTEXTE');
  Result.Prefixe  := GetArgStr('WPF_PREFIXE');

  { Commercial }
  Result.Secteur       := GetArgStr('WPF_SECTEUR');
  Result.SocieteGroupe := GetArgStr('WPF_SOCIETEGROUPE');
  Result.Tiers         := GetArgStr('WPF_TIERS');
  Result.TarifTiers    := GetArgStr('WPF_TARIFTIERS');
  for i := 1 to 10 do
    Result.TableLibreTiers[i] := GetArgStr('WPF_TABLELIBRETIERS' + IntToStr(i));

  { Pièce }
  Result.VenteAchat   := GetArgStr('WPF_VENTEACHAT');
  Result.NaturePieceG := GetArgStr('WPF_NATUREPIECEG');

  { Lieu }
  Result.Etablissement := GetArgStr('WPF_ETABLISSEMENT');
  Result.Depot         := GetArgStr('WPF_DEPOT');

  { Production }
  Result.NatureTravail := GetArgStr('WPF_NATURETRAVAIL');
  Result.TypeOrdre     := GetArgStr('WPF_TYPEORDRE');

  { Scm }
  Result.TypComp    := GetArgStr('WPF_TYPCOMP');
  Result.FamilleMat := GetArgStr('WPF_FAMILLEMAT');

  { Process }
  Result.CodIti    := GetArgStr('WPF_CODITI');
  Result.Circuit   := GetArgStr('WPF_CIRCUIT');
  Result.Pole      := GetArgStr('WPF_POLE');
  Result.Site      := GetArgStr('WPF_SITE');
  Result.Grp       := GetArgStr('WPF_GRP');
  Result.PhasePred := GetArgStr('WPF_PHASEPRED');
  Result.Phase     := GetArgStr('WPF_PHASE');
  Result.PhaseNext := GetArgStr('WPF_PHASENEXT');
  Result.OpeIti    := GetArgStr('WPF_OPEITI');

  { Mouvements }
  Result.StkFlux   := GetArgStr('WPF_STKFLUX');
  Result.QualifMvt := GetArgStr('WPF_QUALIFMVT');

  { Article }
//GP_20080604_MM_GP12663 Déb
  Result.Article     := GetArgStr('WPF_ARTICLE');
  for i := 1 to 3 do
    Result.FamilleNiv[i] := GetArgStr('WPF_FAMILLENIV' + IntToStr(i));
  Result.FamilleValo := GetArgStr('WPF_FAMILLEVALO');
  Result.Collection  := GetArgStr('WPF_COLLECTION');
//GP_20080604_MM_GP12663 Fin
//GP_20080930_TS_GP15021
  Result.MethProd := GetArgStr('WPF_METHPROD');
  for i := 1 to 10 do
    Result.LibreArt[i] := GetArgStr('WPF_LIBREART' + IntToHex(i, 0));
  Result.CodeUL := GetArgStr('WPF_VARCHAR01');
  { Tarifs }
  Result.FormuleMt := GetArgStr('WPF_FORMULEMT');

  { Quotas fournisseurs }
  Result.CodeMarche := GetArgStr('WPF_CODEMARCHE');

  { Situation géographique }
  Result.Pays       := GetArgStr('WPF_PAYS');
  Result.Region     := GetArgStr('WPF_REGION');
  Result.CodePostal := GetArgStr('WPF_CODEPOSTAL');
  Result.SecteurGeo := GetArgStr('WPF_SECTEURGEO');

  { Prix de revient : consolidation }
  Result.TypePdr    := GetArgStr('WPF_TYPEPDR');
  Result.NomChamp   := GetArgStr('WPF_NOMCHAMP');

//GP_20080527_MM_GP12663 Déb
  { Coût indirect }
  Result.CodePort := GetArgStr('WPF_CODEPORT');

  { Ressource }
  Result.Ressource      := GetArgStr('WPF_RESSOURCE');
  Result.TypeRessource  := GetArgStr('WPF_TYPERESSOURCE');
  Result.TarifRessource := GetArgStr('WPF_TARIFRESSOURCE');
//GP_20080527_MM_GP12663 Fin
end;

//Result type String car aujourd'hui il n'y a que des champs texte à tester
function GetValueCleWPF(const CleWPF: PCleWPF; const FieldName: String): String;

  function ExtractNumFromField(const Hexa: Boolean = False): Integer;
  var
    StNum: String;
  begin
    StNum := wRight(FieldName, 2);
    if not (StNum[1] in ['0'..'9']) then
      System.Delete(StNum, 1, 1);
    if Hexa then
      Result := StrToInt('$' + StNum)
    else
      Result := ValeurI(StNum)
  end;

begin
  with CleWPF^ do
  begin
    { Général }
    if      FieldName = 'WPF_USER'     then Result := User
    else if FieldName = 'WPF_CONTEXTE' then Result := Contexte
    else if FieldName = 'WPF_PREFIXE'  then Result := Prefixe
    { Commercial }
    else if FieldName = 'WPF_SECTEUR'       then Result := Secteur
    else if FieldName = 'WPF_SOCIETEGROUPE' then Result := SocieteGroupe
    else if FieldName = 'WPF_TIERS'         then Result := Tiers
    else if FieldName = 'WPF_TARIFTIERS'    then Result := TarifTiers
    else if Pos('WPF_TABLELIBRETIERS', FieldName) = 1 then Result := TableLibreTiers[ExtractNumFromField()]
    { Pièce }
    else if FieldName = 'WPF_VENTEACHAT'  then Result := VenteAchat
    else if FieldName = 'WPF_NATUREPIECEG'then Result := NaturePieceG
    { Lieu }
    else if FieldName = 'WPF_ETABLISSEMENT' then Result := Etablissement
    else if FieldName = 'WPF_DEPOT'         then Result := Depot
    { Production }
    else if FieldName = 'WPF_NATURETRAVAIL' then Result := NatureTravail
    else if FieldName = 'WPF_TYPEORDRE'     then Result := TypeOrdre
    { Scm }
    else if FieldName = 'WPF_TYPCOMP'    then Result := TypComp
    else if FieldName = 'WPF_FAMILLEMAT' then Result := FamilleMat
    { Process }
    else if FieldName = 'WPF_CODITI'    then Result := CodIti
    else if FieldName = 'WPF_CIRCUIT'   then Result := Circuit
    else if FieldName = 'WPF_POLE'      then Result := Pole
    else if FieldName = 'WPF_SITE'      then Result := Site
    else if FieldName = 'WPF_GRP'       then Result := Grp
    else if FieldName = 'WPF_PHASEPRED' then Result := PhasePred
    else if FieldName = 'WPF_PHASE'     then Result := Phase
    else if FieldName = 'WPF_PHASENEXT' then Result := PhaseNext
    else if FieldName = 'WPF_OPEITI'    then Result := OpeIti
    { Mouvements }
    else if FieldName = 'WPF_STKFLUX'   then Result := StkFlux
    else if FieldName = 'WPF_QUALIFMVT' then Result := QualifMvt
    { Article }
    else if FieldName = 'WPF_ARTICLE'            then Result := Article
    else if FieldName = 'WPF_CODEARTICLE'        then Result := iif(CodeArticle <> '', CodeArticle, Trim(wLeft(Article, 18)))
//GP_20080604_MM_GP12663 Déb
    else if Pos('WPF_FAMILLENIV', FieldName) = 1 then Result := FamilleNiv[ExtractNumFromField()]
    else if FieldName = 'WPF_FAMILLEVALO'        then Result := FamilleValo
    else if FieldName = 'WPF_COLLECTION'         then Result := Collection
//GP_20080604_MM_GP12663 Fin
//GP_20080930_TS_GP15021
    else if FieldName = 'WPF_METHPROD'           then Result := MethProd
    else if Pos('WPF_LIBREART', FieldName) = 1   then Result := LibreArt[ExtractNumFromField(True)]
    else if FieldName = 'WPF_VARCHAR01' then Result := CodeUL
    { Tarifs }
    else if FieldName = 'WPF_FORMULEMT' then Result := FormuleMt
    { Quotas fournisseurs }
    else if FieldName = 'WPF_CODEMARCHE' then Result := CodeMarche
    { Situation géographique }
    else if FieldName = 'WPF_PAYS'       then Result := Pays
    else if FieldName = 'WPF_REGION'     then Result := Region
    else if FieldName = 'WPF_CODEPOSTAL' then Result := CodePostal
    else if FieldName = 'WPF_SECTEURGEO' then Result := SecteurGeo
    { Prix de revient : consolidation }
    else if FieldName = 'WPF_TYPEPDR'  then Result := TypePdr
    else if FieldName = 'WPF_NOMCHAMP' then Result := NomChamp
//GP_20080527_MM_GP12663 Déb
    { Coût indirect }
    else if FieldName = 'WPF_CODEPORT' then Result := CodePort
    { Ressource }
    else if FieldName = 'WPF_RESSOURCE'      then Result := Ressource
    else if FieldName = 'WPF_TYPERESSOURCE'  then Result := TypeRessource
    else if FieldName = 'WPF_TARIFRESSOURCE' then Result := TarifRessource
//GP_20080527_MM_GP12663 Fin
    else
      Result := ''
  end
end;

{ TWParamFonction }

function TParamFonction.AddWPF(const CodeFonction, Select, Where, Order, Group: String): TWPF;
begin
  if not Exists(CodeFonction) then
  begin
    Result := TWPF.Create(CodeFonction, Select, Where, Order, Group);
    FList.Add(Result)
  end
  else
    Result := nil
end;

procedure TParamFonction.Clear;
begin
  while Count > 0 do
    RemoveWPF(Pred(Count))
end;

function TParamFonction.Count: Integer;
begin
  Result := FList.Count
end;

constructor TParamFonction.Create;
begin
  FList := TList.Create();
end;

destructor TParamFonction.Destroy;
begin
  Clear();
  FList.Free;
  
  inherited;
end;

function TParamFonction.Exists(const CodeFonction: String): Boolean;
begin
  Result := WPF[CodeFonction] <> nil
end;

function TParamFonction.GetAWPF(CodeFonction: String): TWPF;
var
  i: Integer;
begin
  CodeFonction := UpperCase(Trim(CodeFonction));
  if CodeFonction = '' then
    Result := nil
  else
  begin
    if CodeFonction = FLastCodeFonction then
      Result := FLastWPF
    else
    begin
      Result := nil;
      i := 0;
      while not Assigned(Result) and (i < Count) do
      begin
        if CodeFonction = TWPF(FList[i]).CodeFonction then
          Result := TWPF(FList[i]);
        Inc(i)
      end;
      FLastCodeFonction := CodeFonction;
      FLastWPF := Result
    end
  end
end;

{ Phase d'init : Registre des CodeFonction + Paramètres à gérer dans cette classe }
procedure TParamFonction.Initialisation;
var
  c: Char;
begin
  { Seule la première initialisation compte }
  if Count = 0 then
  begin
    AddWPF('AFFCONTREM'   , 'WPF_BOOLEAN01;WPF_COMBO01'
                          , 'WPF_ARTICLE;WPF_TIERS;WPF_CONTEXTE'
                          , 'WPF_ARTICLE;WPF_TIERS;+WPF_CONTEXTE');
    AddWPF('AUTOWOG'      , 'WPF_INTEGER09..10;WPF_INTEGER01;WPF_INTEGER05;WPF_VARCHAR01..08;WPF_BOOLEAN01'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_CODITI;WPF_CIRCUIT;WPF_POLE;WPF_SITE;'
                          + 'WPF_GRP;WPF_PHASEPRED;WPF_PHASE;WPF_PHASENEXT');
    AddWPF('AUTOWOL'      , 'WPF_BOOLEAN01..13'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_CODITI;WPF_CIRCUIT');
    AddWPF('AUTOWOP'      , 'WPF_BOOLEAN01..13'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_CODITI;WPF_CIRCUIT;WPF_POLE;WPF_SITE;'
                          + 'WPF_GRP;WPF_PHASEPRED;WPF_PHASE;WPF_PHASENEXT');
    AddWPF('SAISIEREBUT'  , 'WPF_COMBO01'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_CODITI;WPF_CIRCUIT;WPF_POLE;WPF_SITE;'
                          + 'WPF_GRP;WPF_PHASE');
    AddWPF('CHXQARTI'     , 'WPF_VARCHAR01'
                          , 'WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_LIBREART1..A');
    AddWPF('CHXQTIERS'    , 'WPF_VARCHAR01'
                          , 'WPF_ARTICLE;WPF_TIERS;WPF_FAMILLENIV1..3;WPF_LIBREART1..A');
    AddWPF('CONVU'        , 'WPF_DOUBLE01..02;WPF_BOOLEAN01..06'
                          , 'WPF_CODEARTICLE;WPF_TIERS');
    AddWPF('DELDISPO'     , 'WPF_INTEGER01'
                          , 'WPF_DEPOT;WPF_TIERS;WPF_SECTEUR;WPF_SOCIETEGROUPE;WPF_QUALIFMVT;WPF_STKFLUX;WPF_ARTICLE;WPF_FAMILLENIV1..3');
    AddWPF('DUREEVIE'     , 'WPF_INTEGER01'
                          , 'WPF_ARTICLE;WPF_FAMILLENIV1..3');
    AddWPF('IMPQTEWOP'    , 'WPF_BOOLEAN01'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_CODITI;WPF_CIRCUIT;WPF_POLE;WPF_SITE;'
                          + 'WPF_GRP;WPF_PHASEPRED;WPF_PHASE;WPF_PHASENEXT');
    AddWPF('MARQARTI'     , 'WPF_VARCHAR01'
                          , 'WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_LIBREART1..A');
    AddWPF('MARQTIERS'    , 'WPF_VARCHAR01'
                          , 'WPF_ARTICLE;WPF_TIERS;WPF_FAMILLENIV1..3;WPF_LIBREART1..A');
    AddWPF('PALMARES'     , ''
                          , 'WPF_VENTEACHAT;WPF_NATUREPIECEG;WPF_TIERS;WPF_TARIFTIERS;WPF_PAYS;'
                          + 'WPF_REGION;WPF_CODEPOSTAL;WPF_SECTEURGEO');
    AddWPF('PREPAEXPEXT'  , 'WPF_BOOLEAN01;WPF_BLOCNOTE'
                          , 'WPF_USER;WPF_CONTEXTE');
//GP_20080507_TS_GP15065
//GP_20080507_TS_GP13386
    AddWPF('PREPAEXPGEN'  , 'WPF_BOOLEAN01..10;WPF_BLOCNOTE'
                          , 'WPF_USER;WPF_CONTEXTE;WPF_NATUREPIECEG');
    AddWPF('PREPAEXPRGP'  , 'WPF_BOOLEAN01;WPF_BLOCNOTE'
                          , 'WPF_USER;WPF_CONTEXTE');
    AddWPF('REPRISELOT'   , 'WPF_COMBO01'
                          , 'WPF_NATURETRAVAIL;WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION');
    AddWPF('RGESTION'     , 'WPF_VARCHAR01'
                          , 'WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION');
    AddWPF('BAC'          , 'WPF_DOUBLE01..02'
                          , 'WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A');
    AddWPF('CONSOTOTREC'  , 'WPF_BOOLEAN01'
                          , 'WPF_NATURETRAVAIL;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION');
    AddWPF('DELGARDE'     , 'WPF_INTEGER01'
                          , 'WPF_ARTICLE;WPF_FAMILLENIV1..3');
//GP_20080616_MM_GP15052
    AddWPF('DELAIMOYFAB'  , 'WPF_INTEGER01;WPF_DOUBLE01;WPF_VARCHAR01;WPF_VARCHAR02'
                          , 'WPF_NATURETRAVAIL;WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION;WPF_CODITI;WPF_CIRCUIT;WPF_PHASE;WPF_OPEITI;');
    AddWPF('PDRREGROUPE'  , 'WPF_VARCHAR01..11'
                          , 'WPF_TYPEPDR;WPF_NOMCHAMP');
    for c := '1' to '3' do
      AddWPF('GSTCLI' + c , 'WPF_DOUBLE01..02'
                          , 'WPF_TIERS;WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION;WPF_LIBREART1..A');
//GP_20080930_TS_GP15021
    AddWPF('RESTRICSO'    , 'WPF_BOOLEAN01'
                          , 'WPF_DEPOT;WPF_NATURETRAVAIL;WPF_TYPEORDRE;WPF_METHPROD;WPF_FAMILLENIV1..3;'
                          + 'WPF_COLLECTION;WPF_LIBREART1..A;WPF_ARTICLE');
    AddWPF('SURCONSOREC'  , 'WPF_BOOLEAN01'
                          , 'WPF_NATURETRAVAIL;WPF_DEPOT;WPF_ARTICLE;WPF_FAMILLENIV1..3;WPF_COLLECTION');
//GP_20080527_MM_GP12663 Déb
    AddWPF('CI'           , 'WPF_CODEPORT'
                          , 'WPF_CODEPORT;WPF_DEPOT;WPF_QUALIFMVT;WPF_NATUREPIECEG;WPF_STKFLUX;'
                           +'WPF_TYPEPDR;WPF_NATURETRAVAIL;WPF_TRAITEMENT;WPF_PHASE;WPF_CIRCUIT;WPF_RESSOURCE;WPF_TYPERESSOURCE;WPF_TARIFRESSOURCE;'
                           +'WPF_ARTICLE;WPF_FAMILLEVALO;WPF_FAMILLENIV1..3;WPF_COLLECTION;WPF_LIBREART1..A;WPF_TIERS;WPF_TARIFTIERS'
                          );
//GP_20080527_MM_GP12663 Fin
  end
end;

procedure TParamFonction.RemoveWPF(const Index: Integer);
var
  AWPF: TWPF;
begin
  AWPF := TWPF(FList[Index]);
  if Assigned(AWPF) then
  begin
    if AWPF.CodeFonction = FLastCodeFonction then
      FLastWPF := nil;
    AWPF.ClearDetail();
    AWPF.Free;
    FList.Delete(Index)
  end
end;

procedure TParamFonction.RemoveWPF(const CodeFonction: String);
var
  AWPF: TWPF;
begin
  AWPF := WPF[CodeFonction];
  if Assigned(AWPF) then
    RemoveWPF(FList.IndexOf(AWPF))
end;

{ TWPF }

function TWPF.BuildGroup: String;
var
  i: Integer;
begin
  if FGroup <> '' then
    Result := FGroup
  else
  begin
    Result := '';
    for i := Low(FGroupFields) to High(FGroupFields) do
      Result := Result + iif(Result <> '', ', ', '') + FGroupFields[i];
    FGroup := Result
  end
end;

function TWPF.BuildOrder: String;
begin
  if FOrder <> '' then
    Result := FOrder
  else
  begin
    Result := OrderByWPF(FCodeFonction);
    FOrder := Result
  end
end;

function TWPF.BuildSelect: String;
var
  i: Integer;

  function inDataField(const FieldName: String): Boolean;
  var
    i: Integer;
  begin
    Result := False;
    i := Low(FDataFields);
    while not Result and (i <= High(FDataFields)) do
    begin
      Result := FieldName = FDataFields[i];
      Inc(i)
    end
  end;

begin
  if FSelect <> '' then
    Result := FSelect
  else
  begin
    Result := '';
    { Champs du select... }
    for i := Low(FDataFields) to High(FDataFields) do
      Result := Result + iif(Result <> '', ', ', '') + FDataFields[i];
    { ... + Champs du where pour effectuer la recherche en mémoire }
    for i := Low(FCritFields) to High(FCritFields) do
      if not inDataField(FCritFields[i]) then
        Result := Result + iif(Result <> '', ', ', '') + FCritFields[i];
    FSelect := Result
  end
end;

function TWPF.BuildSql: String;
begin
  Result := 'SELECT ' + BuildSelect()
          + ' FROM '  + TableName
          + ' WHERE WPF_CODEFONCTION="' + FCodeFonction + '"'
          + iif(Length(FGroupFields) > 0, ' GROUP BY ' + BuildGroup(), '')
          + iif(Length(FOrderFields) > 0, ' ORDER BY ' + BuildOrder(), '')
end;

constructor TWPF.Create(const CodeFonction: String; Select, Where: String;
                        Order: String = ''; Group: String = '');

  procedure StoreFields(var Fields: MyArrayField; var StFields: String; const CheckSigne: Boolean);
  var
    FieldName: String;

    procedure AddInFields(const FieldName: String);
    var
      sSortType: String;
    begin
      if  (Pos('WPF_CODEARTICLE', FieldName) > 0) and (Length(FieldName) <= Length('WPF_CODEARTICLE') + 1)
      and ((CheckSigne and HasFieldInOrder('WPF_CODEARTICLE', sSortType)) or (not CheckSigne and HasFieldIn('WPF_CODEARTICLE', Fields))) then
        Exit;

      SetLength(Fields, Length(Fields) + 1);
      Fields[High(Fields)] := FieldName;
      if CheckSigne and (Pos(wLeft(FieldName, 1)[1], '-+') = 0) then
        Fields[High(Fields)] := '-' + FieldName;

      if (Pos('WPF_ARTICLE', FieldName) > 0) and (Length(FieldName) <= Length('WPF_ARTICLE') + 1)
      and ((CheckSigne and not HasFieldInOrder('WPF_CODEARTICLE', sSortType)) or (not CheckSigne and not HasFieldIn('WPF_CODEARTICLE', Fields))) then
        AddInFields(StringReplace(FieldName, 'WPF_ARTICLE', 'WPF_CODEARTICLE', [rfIgnoreCase]))
    end;
  var
    sCptStart,
    sCptEnd         : String;
    iPosDots, CptLen,
    i, iStart, iStop,
    MaxCptLen       : Integer;
    isHexa          : Boolean;
  begin
    if StFields <> '' then
    begin
      FieldName := ReadTokenSt(StFields);
      if FieldName <> '' then
      begin
        iPosDots := Pos('..', FieldName);
        { Cas de plages de champs }
        if iPosDots > 0 then
        begin
          { Compteur de fin + taille de compteur maxi }
          sCptEnd   := wRight(FieldName, Length(FieldName) - iPosDots - 1);
          MaxCptLen := Length(sCptEnd);
          System.Delete(FieldName, iPosDots, MaxCptLen + 2);
          { Détection Hexa }
          isHexa := (MaxCptLen = 1) and not (sCptEnd[1] in ['0'..'9']);
          if isHexa then
            iStop := StrToInt('$' + sCptEnd)
          else
            iStop := ValeurI(sCptEnd);
          { Compteur de début + taille de compteur mini }
          sCptStart := wRight(FieldName, MaxCptLen);
          if (Length(sCptStart) > 1) and not (sCptStart[1] in ['0'..'9']) then { on considère que les plages sont au max sur 2 digits }
            System.Delete(sCptStart, 1, 1);
          CptLen := Length(sCptStart);
          System.Delete(FieldName, Length(FieldName) - CptLen + 1, CptLen);
          { Détection Hexa }
          isHexa := isHexa or (CptLen = 1) and not (sCptStart[1] in ['0'..'9']);
          if isHexa then
            iStart := StrToInt('$' + sCptStart)
          else
            iStart := ValeurI(sCptStart);
          for i := iStart to iStop do
            AddInFields(FieldName + wPadLeft(iif(isHexa, IntToHex(i, 1), IntToStr(i)), MaxCptLen, '0'))
        end
        else
          AddInFields(FieldName)
      end;
      StoreFields(Fields, StFields, CheckSigne)
    end;
  end;

begin
  inherited Create('WPF_' + CodeFonction);
  { le StackName de la TMemoryTob est alors 'WPF_' + CodeFonction pour garantir son unicité }

  FCodeFonction := CodeFonction;
  if Order = '' then
    Order := Where;
  StoreFields(FDataFields , Select, False);
  StoreFields(FCritFields , Where , False);
  StoreFields(FGroupFields, Group , False);
  StoreFields(FOrderFields, Order , True );
  FActive := True;
end;

destructor TWPF.Destroy;
begin
  SetLength(FDataFields , 0);
  SetLength(FCritFields , 0);
  SetLength(FOrderFields, 0);
  SetLength(FGroupFields, 0);

  inherited;
end;

function TWPF.FindInWPF(const CleWPF: PCleWPF; const TestInitValue: Boolean): Tob;
var
  i, j: Integer;
  OkFields: Boolean;
begin
  if V_Pgi.Sav then
    PutToDebugLog('TWPF.FindInWPF', True, CodeFonction
                 + #13#10 + iif(not isLoaded, Format('Premier chargement !!!'#13#10'-> %d enregistrements', [Detail.Count]), ''));

  Load();
  Result := nil;
  i := 0;
  while not Assigned(Result) and (i < Detail.Count) do
  begin
    j := Low(FCritFields);
    OkFields := True;
    while OkFields and (j <= High(FCritFields)) do
    begin
      //Tob.GetString car aujourd'hui il n'y a que des champs texte à tester
      OkFields := (Detail[i].GetString(FCritFields[j]) = GetValueCleWPF(CleWPF, FCritFields[j]))
               or (TestInitValue and (Detail[i].GetString(FCritFields[j]) = '')); //'' car aujourd'hui il n'y a que des champs texte à tester
      Inc(j)
    end;
    if OkFields then
      Result := Self.Detail[i];
    Inc(i)
  end;

  if V_Pgi.Sav then
    PutToDebugLog('TWPF.FindInWPF', False, CodeFonction);
end;

{ Recherche d'un enregistrement WPF }
function TWPF.FindWPF(const CleWPF: PCleWPF; var DataWPF: TDataWPF; const TestInitValue: Boolean; TobResultWPF: Tob): Boolean;
var
  T: Tob;
begin
  T := FindInWPF(CleWPF, TestInitValue);
  Result := Assigned(T);
  if Result then
  begin
    { Renseigne les données }
    AssignDataWPFFromTobWPF(FCodeFonction, DataWPF, T);
    if Assigned(TobResultWPF) then
      TobResultWPF.Dupliquer(T, False, True, False)
  end
end;

{ Recherche pour contrôler l'existence }
function TWPF.FindWPF(const CleWPF: PCleWPF; const TestInitValue: Boolean): Boolean;
begin
  Result := FindInWPF(CleWPF, TestInitValue) <> nil
end;

function TWPF.HasFieldIn(const FieldName: String; const Fields: MyArrayField): Boolean;
var
  i: Integer;
begin
  Result := False;
  i := Low(Fields);
  while not Result and (i <= High(Fields)) do
  begin
    Result := FieldName = wRight(Fields[i], Min(Length(FieldName), Length(Fields[i])));
    Inc(i)
  end
end;

function TWPF.HasFieldInWhere(const FieldName: String): Boolean;
begin
  Result := HasFieldIn(FieldName, FCritFields)
end;

function TWPF.HasFieldInSelect(const FieldName: String): Boolean;
begin
  Result := HasFieldIn(FieldName, FDataFields)
end;

function TWPF.HasFieldInOrder(const FieldName: String; out SortType: String): Boolean;
var
  i: Integer;
begin
  Result := False;
  i := Low(FOrderFields);
  SortType := 'DESC';
  while not Result and (i <= High(FOrderFields)) do
  begin
    Result := FieldName = wRight(FOrderFields[i], Length(FOrderFields[i]) - 1);
    if Result and (FOrderFields[i][1] = '+') then
      SortType := 'ASC';
    Inc(i)
  end
end;

procedure TWPF.Load(const AvertirCache: Boolean);
begin
  if SqlQuery = '' then
    SqlQuery := BuildSql();

  inherited Load(AvertirCache);
end;

procedure TWPF.SetActive(const Value: Boolean);
begin
  FActive := Value;
  AvertirCache()
end;

{$IFNDEF EAGLSERVER}

{ TWPFButton }
{ Class dérivée du TToolBarButton97 pour créer par programme, des boutons liés
  à la fois à des paramètres société et des "wParamFunc" (paramètres métiers)
  L'implémentation dans les PARAMSOC se fait au niveau d'UilSoc.ChargePageSoc
  de la manière suivante :
  > if ExisteNam('SO_WEXTFORMS', FF) then
      WPFAddWPFButton(TWinControl(GetFromNam('SO_MYPARAMSOC', FF)), 'MYCODEFONCTION', TraduireMemoire('Mon hint de bouton'));}

constructor TWPFButton.Create(const AControler: TWinControl;
  const ACodeFonction, AHint: String; const AArgument: String = '');
begin
  Assert(Assigned(AControler));

  inherited Create(AControler.Owner);

  Controler := AControler;
  FCodeFonction := UpperCase(ACodeFonction);
  FArgument := AArgument;
  OnClick := WPFButtonClick;
  Name := 'BtWPF_' + ACodeFonction;
  Hint := AHint;
  ShowHint := Hint <> '';
  XPEnabled := V_Pgi.DrawXP;
  Mode2003 := V_Pgi.Draw2003;
  GlobalIndexImage := 'Z0008_S16G1'; //<- Equerre rouge
  ChangeForme(AControler)            //<- pour dessin correct du bouton
end;

destructor TWPFButton.Destroy;
begin
  inherited Destroy();
end;

procedure TWPFButton.SetControler(const Value: TWinControl);
var
  w, h: Integer;
begin
  FControler := Value;
  Parent := FControler.Parent;
  if FControler.InheritsFrom(TCustomEdit) then
    h := FControler.Height
  else
    h := 22;
  w := h + 1;
  SetBounds(FControler.Left + FControler.Width + 5, FControler.Top + ((FControler.Height - h) div 2), w, h);
end;

procedure TWPFButton.WPFButtonClick(Sender: TObject);
begin
  AGLLanceFiche('W' ,'WPF_' + FCodeFonction + '_M'  , '', '',
                'ACTION=MODIFICATION;CODEFONCTION=' + FCodeFonction + iif(FArgument <> '', ';' + FArgument, ''))
end;

class function TWPFButton.Add(const Controler: TWinControl;
  const CodeFonction, Hint: String; const Argument: String = ''): TWPFButton;
begin
  Result := TWPFButton.Create(Controler, CodeFonction, Hint, Argument)
end;

{$ENDIF !EAGLSERVER}

end.

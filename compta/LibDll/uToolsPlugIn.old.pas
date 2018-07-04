unit uToolsPlugin;

{$IFNDEF EAGLSERVER}abghgahgh{$ENDIF}

interface

uses classes,
  HEnt1,
  HCtrls,
  DBTables,
  uTobView,
  UTOB;

function LoadFormeXML(RequestTob: TOB; Nature, Forme: string): boolean;
function LoadXMLData(RequestTob: TOB; cSQL, cXMLDATA: string): boolean;
function SaveXMLData(TobsXml, TobXml: TOB; cXMLDATA: string; Compressed: boolean = false): boolean;
procedure ChargeData(TobCtrl, DataTob: Tob; ADossier: string = '');

function GetListeMultiDossierDB(Groupement, DBName: string): string;

type
  HString = string;
  TVignettePlugin = class(TObject)
  private
    LaVignette: Tob;
    FormeToSend: Tob;
    PropsToSend: Tob;
    DatasToSend: Tob;
    DatasRequest: Tob;
    CritRequest: Tob;
    CritToSend: TOB;
    DatasLinked: Tob;
    FPlugInName: string;
    function GetControlCount: integer;
    function GetCritInstance(ControlName: string): Tob;
    procedure SetSqlMutliDossier(var SQL: string);
  protected
  public
    constructor Create(TobIn, TobOut: Tob); virtual;
    destructor Destroy; override;
    function GetDataInstance(ControlName: string): Tob;
    function GetPropInstance(PropName, ControlName: string): Tob;
    function GetControlValue(ControlName: string): variant; overload;
    function GetControlValue(ControlIndex: integer): variant; overload;
    function GetGridValue(GridName, ColumnName: string; ARow: integer): variant;
    function GetControlName(ControlIndex: integer): string;
    procedure SetControlVisible(ControlName: string; Value: Boolean);
    procedure SetControlEnabled(ControlName: string; Value: Boolean);
    procedure SetControlValue(ControlName: string; Value: Variant);
    procedure SetGridColumns(ControlName: string; Datas: Tob); overload;
    procedure SetGridColumns(ControlName: string; Datas: Tob; DBListe : string; UserName: string); overload;
    procedure PutGridDetail(ControlName: string; Datas: Tob; DBListe : string ='' ; UserName: string = '');
    procedure PutEcran(Datas: Tob);
    procedure SetDesignedData(ControlType, ControlName: string; ControlValue: variant);
    procedure SetComboDetail(ControlName: string; InfoTob: TOB);
    function GetControlProperty(ControlName, PropertyName: string): variant;
    function WhereSQL: HString;
    function GetCritereValue(AControlName: string): hString;
    procedure SetCritereValue(AControlName: string; AValue: Variant);
    function GetDateTime(NomChp: string; Critere: Boolean = True): TDateTime;
    function GetDouble(NomChp: string; Critere: Boolean = True): Double;
    function GetInteger(NomChp: string; Critere: Boolean = True): Integer;
    function GetString(NomChp: string; Critere: Boolean = True): string;
    function GetTobProperty(Grille, Chp: string): TOB;
    function GetDossier: string;
    function OpenSelect(Requete: string; ADossier: string = ''): TQuery;
    procedure SetFormatCol(Grille, Chp, Format: string);
    procedure SetTitreCol(Grille, Chp, Lib: string);
    procedure SetVisibleCol(Grille, Chp: string; VisibleOk: Boolean);
    procedure SetWidthCol(Grille, Chp: string; Largeur: Integer);
    function GetSQLListe(AUserName, AListName: string): string;
    function GetSQLCount(ATableName, AWhereSt: string): integer;
    function GetSQLFiche(ATableName: string; var ARowCount: integer): string;
    function GetGridDataSelected(GridName: string; var ABookmark: integer): TOB;
    function GetLinkedValue(ControlName: string): tob;
    procedure ConvertFieldValue(laTOB: TOB);
    function GetTobViewerColumn(AControlName, AColName: string): TOB;
    procedure PutTobViewerDetail(ControlName: string; Datas: Tob; AColNames: string = '');
    procedure SetTobViewerCalcOpField(AControlName, AColName: string; ACalcOp: TCalcOp);
    procedure SetTobViewerColumns(ControlName: string; Datas: Tob; AColNames: string);
    procedure SetTobViewerGroupField(AControlName, AColName: string; AFixedPos: TFixedPos);
    procedure SetTobViewerRegField(AControlName, AColName: string; ARegIndex: integer; ARegSortOrder: TSortOrder = soNone);
    procedure SetTobViewerSortField(AControlName, AColName: string; ASortIndex: integer; ASortOrder: TSortOrder = soNone);
    procedure SetTobViewerWidthField(AControlName, AColName: string; AWidth: integer);
    procedure SetTobViewerCaptionField(AControlName, AColName, ACaption: string);
  private
    function GetListeSoc: string;
    function GetRegroupement: string;
  public
    function IsCritereVisible(AControlName: string): Boolean;
    function IsForceDataBaseName(var databasename: string; PlugInName: string): boolean;

    property Dossier: string read getDossier; {Dossier sur lequel il faut lancer les requêtes}
    property ControlCount: integer read GetControlCount;
    property ListeSoc: string read GetListeSoc;
    property Regroupement: string read GetRegroupement;
  end;

implementation

uses
  windows,
  eSession,
  sysutils,
  uMultiDossierUtil,
  comctrls,
  hDTLinks,
  DB,
  MajTable,
  FileCtrl,
  IniFiles,
  Forms,
  hQry;

function WildChar(St: string): string;
var
  St1: string;
  i: integer;
begin
  St1 := '';
  for i := 1 to Length(St) do
  begin
    case St[i] of
      '*': St1 := St1 + '%';
      '?': St1 := St1 + '_';
    else
      St1 := St1 + St[i];
    end;
  end;
  WildChar := St1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joseph MOREL
Créé le ...... : 21/09/2005
Modifié le ... :   /  /
Description .. : Charge en TOB le flux XML d'une fiche sauvegardée dans
Suite ........ : FORMESXML.
Mots clefs ... :
*****************************************************************}

function LoadFormeXML(RequestTob: TOB; Nature, Forme: string): boolean;
var
  TobFormes, TobFormesXml, TobFormeXml: TOB;
  TobScript: TOB; //PGR 25/10/2005
//  StringStream: TStringStream;
  MemoryStream : TMemoryStream;
  Entete: Boolean;
  Buffer, Encoding: string;
begin
  Result := False;

  // voir pour stocker l'ancetre dans SFORMESXML !!
  TobFormes := Tob.Create('_FORMES_', nil, -1);
  TobFormes.LoadDetailFromSQL('SELECT DFM_ANCETRE FROM FORMES WHERE (DFM_TYPEFORME = "' + Nature + '") AND (DFM_FORME = "' + Forme + '")');
  //ddwriteln('LoadFormeXML:Requete='+'SELECT DFM_ANCETRE FROM FORMES WHERE (DFM_TYPEFORME = "' + Nature + '") AND (DFM_FORME = "' + Forme + '")');
  //ddwriteln('LoadFormeXML:TobFormes.Detail.count'+IntToStr(TobFormes.Detail.count));
  if TobFormes.Detail.count > 0 then
    RequestTob.AddChampSupValeur('ANCETRE', TobFormes.Detail[0]['DFM_ANCETRE'].Valeur);
  TobFormes.Free;

  TobFormesXml := Tob.Create('_SFORMESXML_', nil, -1);
  TobFormesXml.LoadDetailFromSQL('SELECT SFX_XMLDATA FROM SFORMESXML WHERE (SFX_TYPEFORME = "' + Nature + '") AND (SFX_FORME = "' + Forme + '")');

  //ddwriteln('LoadFormeXML:Requete='+'SELECT SFX_XMLDATA FROM SFORMESXML WHERE (SFX_TYPEFORME = "' + Nature + '") AND (SFX_FORME = "' + Forme + '")');
  //ddwriteln('LoadFormeXML:TobFormesXml.Detail.Count'+IntToStr(TobFormesXml.Detail.Count));

  if TobFormesXml.Detail.Count > 0 then
  begin
//    StringStream := TStringStream.Create(TobFormesXml.Detail[0]['SFX_XMLDATA'].Valeur);

// à mettre du coté client apres avoir avoir vu avec P Gronchi
    Buffer := TobFormesXml.Detail[0]['SFX_XMLDATA'].Valeur;
    MemoryStream := TMemoryStream.Create;
    MemoryStream.Clear;
    DecompressTextObjectStream(MemoryStream);
    MemoryStream.writeBuffer(Pointer(Buffer)^,Length(Buffer));
    DeCompressTextObjectStream(MemoryStream);
    MemoryStream.Seek(0, 0);

    TobFormeXml := Tob.Create('', nil, -1);
    TobFormeXml.LoadFromXMLStream(MemoryStream, Entete, Encoding);
//    StringStream.Free;
    MemoryStream.Free;

    //PGR 25/10/2005
    TobScript := Tob.Create('T_SCRIPT', nil, -1);
    TobScript.LoadDetailFromSQL('SELECT SRC_DATA FROM SCRIPTS WHERE SRC_TYPESCRIPT="' + Nature + '" AND SRC_SCRIPT="' + Forme + '"');

    if TobScript.Detail.Count > 0 then
    begin
      TobScript.Detail[0].ReaffecteNomTable('T_SCRIPT', True);
      //      TobScript.Detail[0].ChangeParent(TobFormeXml, -1);
      TobScript.Detail[0].ChangeParent(TobFormeXml, -1, True);
    end;

    TobScript.Free;

    //    TobFormeXml.ChangeParent(RequestTob, -1);
    TobFormeXml.ChangeParent(RequestTob, -1, True);
    Result := True;
  end;

  TobFormesXml.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joseph MOREL
Créé le ...... : 21/09/2005
Modifié le ... :   /  /
Description .. : Charge en TOB le flux XML sauvegardé dans un champ
Suite ........ : XML. cSQL doit contenir la requête sur ce champ
Mots clefs ... :
*****************************************************************}

function LoadXMLData(RequestTob: TOB; cSQL, cXMLDATA: string): boolean;
var
  TobsXml, TobXml: TOB;
  StringStream: TStringStream;
  Entete: Boolean;
  Encoding: string;
  i: integer;
begin
  Result := False;
  TobsXml := Tob.Create('la tob xml', nil, -1);
  TobsXml.LoadDetailFromSQL(cSQL);
  try
    if TobsXml.Detail.Count > 0 then
    begin
      for i:=0 to TobsXml.Detail.Count-1 do
      begin
        StringStream := TStringStream.Create(TobsXml.Detail[i][cXMLDATA].Valeur);
        if StringStream.Size > 0 then
        begin
          TobXml := Tob.Create('', nil, -1);
          try
            TobXml.LoadFromXMLStream(StringStream, Entete, Encoding);
            Result := True;
          finally
            StringStream.Free;
            TobXml.ChangeParent(RequestTob, -1, True);
          end;
        end
        else
          StringStream.Free;
      end;
    end;
  finally
    TobsXml.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joseph MOREL
Créé le ...... : 21/09/2005
Modifié le ... :   /  /
Description .. : Sauvegarde la TobXml dans un mémo de base de données
Mots clefs ... :
*****************************************************************}

function SaveXMLData(TobsXml, TobXml: TOB; cXMLDATA: string; Compressed: boolean = false): boolean;
var
  Buffer: string;
  MemoryStream: TMemoryStream;
begin
  Result := false;

  MemoryStream := TMemoryStream.Create;

  if TobXml <> nil then
  begin
    try
      TobXml.SaveToXmlStream(MemoryStream, LookupCurrentSession.XMLEntete, LookupCurrentSession.XMLEncoding);

      if Compressed then
        CompressTextObjectStream(MemoryStream);

      if MemoryStream.Size > 0 then
      begin
        MemoryStream.Position := 0;
        SetLength(Buffer, MemoryStream.Size);
        MemoryStream.ReadBuffer(Pointer(Buffer)^, MemoryStream.Size);
      end;

      MemoryStream.Free;
    except
      MemoryStream.Free;
      Exit;
    end;
  end;

  BeginTrans;
  try
    if TobXml <> nil then
      TobsXml.PutValue(cXMLDATA, Buffer);
    TobsXml.InsertOrUpdateDb(false);
    CommitTrans;
  except
    RollBack;
    Exit;
  end;

  Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Joseph MOREL
Créé le ...... : 21/09/2005
Modifié le ... :   /  /
Description .. : Génère une DataTob avec les données implicites d'une
Suite ........ : TobForme.
Mots clefs ... :
*****************************************************************}

procedure ChargeData(TobCtrl, DataTob: Tob; ADossier: string = '');
var
  iDetail: Integer;
  NewData, NewTob, ComboTob: Tob;
  stSQL: hstring;
  item, value: variant;
begin
  if TobCtrl.NomTable = 'T_FORM' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end
  else if TobCtrl.NomTable = 'T_VALCOMBOBOX' then
  begin
    if TobCtrl.getString('DATATYPE') <> '' then
    begin
      stSql := GetTabletteSql(TobCtrl.getString('DATATYPE'), TobCtrl.getString('PLUS'));
      if ADossier <> '' then
        stSQL := sqlToMultiBases(stSQL, ADossier);
      ComboTob := TOB.Create('ComboTob', nil, -1);
      ComboTob.LoadDetailDBFromSQL('VirtualDataType', stSql);
      if ComboTob.Detail.Count > 0 then
      begin
        NewData := Tob.Create('T_DATAVALCOMBOBOX', DataTob, -1);
        NewData.AddChampSupValeur('NAME', TobCtrl['NAME'].Valeur, CSTString);
        value := '';
        NewData.AddChampSupValeur('VALUE', value, CSTString);
        for iDetail := 0 to ComboTob.Detail.Count - 1 do
        begin
          NewTob := Tob.Create('T_ITEMVALUE', NewData, -1);
          item := ComboTob.Detail[iDetail][1].Valeur;
          NewTob.AddChampSupValeur('ITEM', item, CSTString);
          value := ComboTob.Detail[iDetail][0].Valeur;
          NewTob.AddChampSupValeur('VALUE', value, CSTString);
        end;
      end;
      FreeAndNil(ComboTob);
    end;
  end
  else if TobCtrl.NomTable = 'T_EDIT' then
  begin
  end
  else if TobCtrl.NomTable = 'T_LABEL' then
  begin
  end
  else if TobCtrl.NomTable = 'T_PAGECONTROL' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end
  else if TobCtrl.NomTable = 'T_TABSHEET' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end
  else if TobCtrl.NomTable = 'T_PANEL' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end
  else if TobCtrl.NomTable = 'T_SCROLLBOX' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end
  else if TobCtrl.NomTable = 'T_CHECKBOX' then
  begin
  end
  else if TobCtrl.NomTable = 'T_RADIOBUTTON' then
  begin
  end
  else if TobCtrl.NomTable = 'T_MEMO' then
  begin
  end
    {
    else if Ctrl is THImage then
    begin
    end
    }
  else if TobCtrl.NomTable = 'T_VIGNETTECONTEXTE' then
  begin
  end
  else if TobCtrl.NomTable = 'T_GRID' then
  begin
  end
  else if TobCtrl.NomTable = 'T_TOOLBAR' then
  begin
    for iDetail := 0 to TobCtrl.Detail.Count - 1 do
      ChargeData(TobCtrl.Detail[iDetail], DataTob, ADossier);
  end;
end;

{ TVignettePlugin }

constructor TVignettePlugin.Create(TobIn, TobOut: Tob);
var
  i: integer;
begin
  inherited create;
  LaVignette := Tob.Create('T_VIGNETTE', TobOut, -1); //renvoyé au client
  FormeToSend := Tob.Create('T_FORM', LaVignette, -1); //renvoyé au client
  PropsToSend := Tob.Create('T_CONTROLPROPS', LaVignette, -1); //renvoyé au client
  DatasToSend := Tob.Create('T_DATACONTEXTE', LaVignette, -1); //renvoyé au client
  CritToSend := Tob.Create('T_DATACRITERES', LaVignette, -1); //renvoyé au client
  DatasRequest := nil;
  CritRequest := nil;
  DatasLinked := nil;
  if TobIn.FieldExists('PLUGINNAME') then
    FPlugInName := TobIn.GetValue('PLUGINNAME');
  for i := 0 to TobIn.detail.count - 1 do
  begin
    if TobIn.Detail[i].NomTable = 'T_DATACONTEXTE' then
    begin
      DatasRequest := TobIn.Detail[i];
      //    break;
    end
    else if TobIn.Detail[i].NomTable = 'T_DATACRITERES' then
    begin
      CritRequest := TobIn.Detail[i];
    end
    else if TobIn.Detail[i].NomTable = 'T_LINKDATAS' then
    begin
      DatasLinked := TobIn.Detail[i];
    end
    else
      ;
  end;
end;

destructor TVignettePlugin.Destroy;
begin
  inherited;
end;

function TVignettePlugin.GetGridValue(GridName, ColumnName: string; ARow: integer): variant;
var
  TG, T: TOB;
  i: integer;
begin
  TG := DatasRequest.FindFirst(['NAME'], [GridName], True);
  if TG <> nil then
  begin
    T := nil;
    for i := 0 to TG.Detail.count - 1 do
    begin
      if TG.Detail[i].NomTable = 'T_GRIDVALUES' then
        T := TG.Detail[i];
    end;
    TG := T;
    if Assigned(TG) and (ARow < TG.Detail.Count) then
    begin
      T := TG.Detail[ARow + 1]; // correspondance num ligne / indice tob
      if T.FieldExists(ColumnName) then
        result := T.GetValue(ColumnName)
      else
        result := 'error: Colonne inconnue !';
    end;
  end
  else
    result := 'error: Objet non trouvé !';
end;

function TVignettePlugin.GetGridDataSelected(GridName: string; var ABookmark: integer): TOB;
var
  TG, TB: TOB;
  i: integer;
begin
  TG := DatasRequest.FindFirst(['NAME'], [GridName], True);
  result := nil;
  if TG <> nil then
  begin
    for i := 0 to TG.Detail.count - 1 do
    begin
      if TG.Detail[i].NomTable = 'T_GRIDVALUES' then
      begin
        TB := TG.Detail[i];
        Inc(ABookmark);
        if Assigned(TB) and (ABookmark < TB.Detail.Count) then
        begin
          result := TB.Detail[ABookmark];
        end;
      end;
    end;
  end;
end;

(*
function TVignettePlugin.GetGridBookmark(GridName: string; var AllSelected: boolean; ABookmark: integer): variant;
var
  TG,T: TOB;
  i: integer;
begin
  TG := DatasRequest.FindFirst(['NAME'], [GridName], True);
  if TG <> nil then
  begin
    T := nil;
    for i:=0 to TG.Detail.count-1 do
    begin
      if TG.Detail[i].NomTable='T_BOOKMARKS' then
        T := TG.Detail[i];
    end;
    TG := T;
    if Assigned(TG) and (ABookmark < TG.Detail.Count) then
    begin
      if TG.NomTable='T_GRIDBOOKMARKS' then
        AllSelected := TG.GetValue('ALLSELECTED');
      T := TG.Detail[ABookmark];
      if T.FieldExists('ROW') then result := T.GetValue('ROW')
      else result := -1;
    end;
  end else result := -1;
end;
*)

function TVignettePlugin.GetLinkedValue(ControlName: string): tob;
begin
  result := nil;
  if not assigned(DatasLinked) then
    exit;
  result := DatasLinked.FindFirst(['NAME'], [ControlName], True);
end;

function TVignettePlugin.GetControlValue(ControlName: string): variant;
var
  T: TOB;
  i: integer;
  V: variant;
begin
  T := DatasRequest.FindFirst(['NAME'], [ControlName], True);
  if T <> nil then
  begin
    V := T.GetValue('VALUE');
    i := T.GetNumChamp('VALUE');
    if i >= 1000 then
    begin
      //TCSType = (CSTUknow, CSTInteger, CSTDouble, CSTDate, CSTBoolean, CSTString, CSTBase64Binary, CSTHexBinary);
      case TCS(T.ChampsSup[i - 1000]).CSType of
        CSTinteger: Result := VarAsType(V, varInteger);
        CSTDouble: Result := VarAsType(V, varDouble);
        CSTDate: Result := VarAsType(V, varDate); // todo les autres type
      else
        Result := V;
      end;
    end
    else
      Result := V;
  end;
end;

function TVignettePlugin.GetDataInstance(ControlName: string): Tob;
var
  T: TOB;
begin
  Result := DatasToSend.FindFirst(['NAME'], [ControlName], True); // si on trouve dans DatasToSend, ok
  if Result = nil then
  begin //sinon on cherche dans DatasRequest
    T := DatasRequest.FindFirst(['NAME'], [ControlName], True);
    if T <> nil then //on réplique les données dans DatasToSend
    begin
      Result := Tob.Create(T.NomTable, DatasToSend, -1);
      Result.Assign(T);
    end;
  end;
end;

function TVignettePlugin.GetPropInstance(PropName, ControlName: string): Tob;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to PropsToSend.Detail.Count - 1 do
    if PropsToSend.Detail[i].NomTable = 'T_CONTROLPROP' then
      if (PropsToSend.Detail[i].GetValue('PROPNAME') = PropName) and (PropsToSend.Detail[i].GetValue('CONTROLNAME') = ControlName) then
      begin
        result := PropsToSend.Detail[i];
        break;
      end;
end;

procedure TVignettePlugin.SetControlEnabled(ControlName: string; Value: Boolean);
var
  Tp: Tob;
begin
  Tp := GetPropInstance('ENABLED', ControlName);
  if Tp = nil then
  begin
    Tp := Tob.Create('T_CONTROLPROP', PropsToSend, -1);
    Tp.AddChampSupValeur('PROPNAME', 'ENABLED', CSTString);
    Tp.AddChampSupValeur('CONTROLNAME', ControlName, CSTString);
  end;
  Tp.AddChampSupValeur('VALUE', Value, CSTBoolean);
end;

procedure TVignettePlugin.SetControlValue(ControlName: string; Value: Variant);
var
  T: TOB;
begin
  T := GetDataInstance(ControlName);
  if T <> nil then
    T.PutValue('VALUE', Value);
end;

procedure TVignettePlugin.SetControlVisible(ControlName: string; Value: Boolean);
var
  Tp: Tob;
begin
  Tp := GetPropInstance('VISIBLE', ControlName);
  if Tp = nil then
  begin
    Tp := Tob.Create('T_CONTROLPROP', PropsToSend, -1);
    Tp.AddChampSupValeur('PROPNAME', 'VISIBLE', CSTString);
    Tp.AddChampSupValeur('CONTROLNAME', ControlName, CSTString);
  end;
  Tp.AddChampSupValeur('VALUE', Value, CSTBoolean);
end;

function CalcType(Typ: string): string;
begin
  if (Typ = 'COMBO') then
    Result := ''
  else if (Pos('CHAR', Typ) > 0) then
    Result := ''
  else if Typ = 'BOOLEAN' then
    Result := 'B'
  else if Typ = 'DATE' then
    Result := 'D'
  else if Typ = 'INTEGER' then
    Result := 'I'
  else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
    Result := 'R'
  else if Typ = 'BLOB' then
    Result := 'M'
  else
    Result := ''
end;

procedure TVignettePlugin.SetGridColumns(ControlName: string; Datas: Tob);
var
  i: integer;
  T, TNew, TCol: TOB;
  Name, Lib, Typ: string;

  function CalcWidth: integer;
  var
    il: integer;
  begin
    if (Typ = 'COMBO') then
      Result := 16
    else if (Pos('CHAR', Typ) > 0) then
    begin
      il := ValeurI(Typ);
      if il <= 20 then
        Result := (il * 3) + 2
      else if il <= 50 then
        Result := (il * 2)
      else
        Result := 100;
    end
    else if Typ = 'BOOLEAN' then
      Result := 10
    else if Typ = 'DATE' then
      Result := 24
    else if Typ = 'INTEGER' then
      Result := 24
    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      Result := 30
    else if Typ = 'BLOB' then
      Result := 8
    else
      Result := 16;
  end;

  function CalcAlign: string;
  begin
    if (Typ = 'COMBO') then
      Result := 'G.0  ---'
    else if (Pos('CHAR', Typ) > 0) then
      Result := 'G.0  ---'
    else if Typ = 'BOOLEAN' then
      Result := 'C.0  ---'
    else if Typ = 'DATE' then
      Result := 'G.0  ---'
    else if Typ = 'INTEGER' then
      Result := 'D.0  ---'
    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      Result := 'D.2  ---'
    else if Typ = 'BLOB' then
      Result := 'G.0  ---'
    else
      Result := 'G.0  ---'
  end;

  function ChampToCSType(ch: TCS): string;
  begin
    case ch.CSType of
      CSTInteger: result := 'INTEGER';
      CSTDouble: result := 'DOUBLE';
      CSTDate: result := 'DATETIME';
      CSTBoolean: result := 'BOOLEAN';
      CSTString: result := 'STRING';
      CSTBase64Binary: result := 'DATA';
      //      CSTHexBinary:  Typ := '';
      CSTBooleanStr: result := 'BOOLEAN';
    else
      result := ChampToType(ch.Nom);
    end;
  end;

begin
  if Datas.Detail.count = 0 then
    exit;

  T := GetDataInstance(ControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATAGRIDCOLUMNS' then
      begin
        T.detail[0].Free;
        break;
      end;
    end;
    TNew := Tob.Create('T_DATAGRIDCOLUMNS', T, -1);
    T := Datas.Detail[0];
    for i := 1 to T.nbChamps do
    begin
      TCol := Tob.Create('T_DATAGRIDCOLUMN', TNew, -1);
      Name := T.GetNomChamp(i);
      Typ := ChampToType(Name);
      TCol.AddChampSupValeur('CHAMP', name, CSTString);
      TCol.AddChampSupValeur('LIBELLE', ChampToLibelle(Name), CSTString);
      TCol.AddChampSupValeur('WIDTH', CalcWidth, CSTInteger);
      TCol.AddChampSupValeur('ALIGN', CalcAlign, CSTString);
      TCol.AddChampSupValeur('VISIBLE', True, CSTBoolean);
    end;
    for i := 0 to T.ChampsSup.Count - 1 do
    begin
      TCol := Tob.Create('T_DATAGRIDCOLUMN', TNew, -1);
      Name := TCS(T.ChampsSup[i]).Nom;
      Typ := ChampToCSType(TCS(T.ChampsSup[i]));
      Lib := ChampToLibelle(Name);
      if Lib = '??' then
        Lib := Name;
      TCol.AddChampSupValeur('CHAMP', name, CSTString);
      TCol.AddChampSupValeur('LIBELLE', Lib, CSTString);
      TCol.AddChampSupValeur('WIDTH', CalcWidth, CSTInteger);
      TCol.AddChampSupValeur('ALIGN', CalcAlign, CSTString);
      TCol.AddChampSupValeur('TYPE', CalcType(Typ), CSTString);
      TCol.AddChampSupValeur('VISIBLE', True, CSTBoolean);
    end;
  end;
end;

procedure TVignettePlugin.SetGridColumns(ControlName: string; Datas: Tob; DBListe : string; UserName: string);
var
    OkTri, OkNumCol: boolean;
    FSources, FChamps, FLiaison, FFPerso, FParams, FLargeurs, FJustifs, FTitres, FLibelle, FTris, FNumCols: string;
    StF, Typ, Align, FF, F, StN, stL, stJ, Stt : string;
    T, TCol, TNew: TOB;
    i: integer;
    Dec: Integer;
    Sep, Obli, Lib, Visu, Nulle, Cumul: Boolean;
begin
  ChargeHListeUser (DBListe, UserName, FSources, FLiaison, FTris, FChamps, FTitres, FLargeurs, FJustifs, FParams, FLibelle, FNumCols, FFPerso, OkTri, OkNumCol);
  StF := FChamps;
  StL := FLargeurs;
  StJ := FJustifs;
  StT := FTitres;
  StN := FNumCols;

  if Datas.Detail.count = 0 then
    exit;

  T := GetDataInstance(ControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATAGRIDCOLUMNS' then
      begin
        T.detail[0].Free;
        break;
      end;
    end;
    TNew := Tob.Create('T_DATAGRIDCOLUMNS', T, -1);
    while StF <> '' do
    begin
      TCol := Tob.Create('T_DATAGRIDCOLUMN', TNew, -1);
      F := Trim (ReadTokenSt (StF) ) ;
      Typ := ChampToType(F);
      Align := ReadtokenSt (StJ);
      TransAlign (Align , FF, Dec, Sep, Obli, Lib, Visu, Nulle, Cumul) ;
      TCol.AddChampSupValeur('CHAMP', F, CSTString);
      TCol.AddChampSupValeur('LIBELLE', ReadTokenSt (StT), CSTString);
      TCol.AddChampSupValeur('WIDTH', ReadtokenSt (StL), CSTInteger);
      TCol.AddChampSupValeur('ALIGN', ALign, CSTString);
      TCol.AddChampSupValeur('VISIBLE', Visu, CSTBoolean);
      TCol.AddChampSupValeur('TYPE', CalcType(Typ), CSTString);
    end;
  end;
end;

procedure TVignettePlugin.PutGridDetail(ControlName: string; Datas: Tob; DBListe : string ='' ; UserName: string = '');
var
  i: integer;
  T, TNew: TOB;
begin
  if (DBListe<>'') and (UserName<>'') then
    SetGridColumns(ControlName, Datas, DBListe, UserName)
  else
    SetGridColumns(ControlName, Datas); //génère les colonnes
  for i := 0 to Datas.Detail.count - 1 do
    Datas.detail[i].ReaffecteNomTable('T_DATAGRIDLIGNE', true); //pour le message XML
  T := GetDataInstance(ControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATAGRIDLIGNES' then
      begin
        T.detail[0].Free;
        break;
      end;
    end;
    TNew := Tob.Create('T_DATAGRIDLIGNES', T, -1);
    while Datas.detail.count > 0 do
      Datas.Detail[0].ChangeParent(TNew, -1, True);
  end;
end;

procedure TVignettePlugin.PutEcran(Datas: Tob);
var
  i: integer;
begin
  for i := 1 to Datas.nbChamps do
    SetControlValue(Datas.GetNomChamp(i), Datas.GetValeur(i));
  for i := 0 to Datas.ChampsSup.Count - 1 do
    SetControlValue(Datas.GetNomChamp(i + 1000), Datas.GetValeur(i + 1000));
end;

function TVignettePlugin.GetControlCount: integer;
begin
  if assigned(DatasRequest) then
    result := DatasRequest.Detail.Count
  else
    result := 0;
end;

function TVignettePlugin.GetControlName(ControlIndex: integer): string;
begin
  result := '';
  if ControlIndex >= DatasRequest.Detail.Count then
    exit;
  result := DatasRequest.Detail[ControlIndex].GetValue('NAME');
end;

function TVignettePlugin.GetControlValue(ControlIndex: integer): variant;
var
  Name: string;
begin
  result := '';
  if ControlIndex >= DatasRequest.Detail.Count then
    exit;
  Name := GetControlName(ControlIndex);
  if Name <> '' then
    result := GetControlValue(Name);
end;

procedure TVignettePlugin.SetDesignedData(ControlType, ControlName: string; ControlValue: variant);
var
  T: TOB;
begin
  T := GetDataInstance(ControlName);
  if T = nil then // normalement, c'est un control de l'ancetre
  begin
    T := Tob.Create(ControlType, DatasToSend, -1);
    T.AddChampSupValeur('NAME', ControlName);
    T.AddChampSupValeur('VALUE', ControlValue);
  end;
end;

procedure TVignettePlugin.SetComboDetail(ControlName: string; InfoTob: TOB);
var
  i: integer;
  s, v: string;
  NewData, NewTob: TOB;
begin
  try
    NewData := GetDataInstance(ControlName);
    //    T := InfoTob.Detail[0];
    NewData.ClearDetail;
    for i := 0 to InfoTob.Detail.Count - 1 do
    begin
      s := InfoTob.Detail[i].GetValue('ITEM');
      v := InfoTob.Detail[i].GetValue('VALUE');
      if i = 0 then
        NewData.AddChampSupValeur('VALUE', v, CSTString);
      NewTob := Tob.Create('T_ITEMVALUE', NewData, -1);
      NewTob.AddChampSupValeur('ITEM', s, CSTString);
      NewTob.AddChampSupValeur('VALUE', v, CSTString);
    end;
  finally
  end;
end;

function TVignettePlugin.GetControlProperty(ControlName, PropertyName: string): variant;
var
  T: TOB;
  i: integer;
  V: variant;
begin
  T := DatasRequest.FindFirst(['NAME'], [ControlName], True);
  if T <> nil then
  begin
    if T.FieldExists(PropertyName) then
    begin
      V := T.GetValue(PropertyName);
      i := T.GetNumChamp(PropertyName);
      if i >= 1000 then
      begin
        //TCSType = (CSTUknow, CSTInteger, CSTDouble, CSTDate, CSTBoolean, CSTString, CSTBase64Binary, CSTHexBinary);
        case TCS(T.ChampsSup[i - 1000]).CSType of
          CSTinteger: Result := VarAsType(V, varInteger);
          CSTDouble: Result := VarAsType(V, varDouble);
          CSTDate: Result := VarAsType(V, varDate); // todo les autres type
        else
          Result := V;
        end;
      end
      else
        Result := V;
    end
    else
    begin
      result := varUnknown;
    end;
  end;
end;

function StrToOperateur(Operateur: string): TOperateur;
begin
  if Operateur = 'COMMENCE' then
    Result := Commence
  else if Operateur = 'CONTIENT' then
    Result := Contient
  else if Operateur = 'SUPERIEUR' then
    Result := Superieur
  else if Operateur = 'INFERIEUR' then
    Result := Inferieur
  else if Operateur = 'DIFFERENT' then
    Result := Different
  else
    result := egal;
end;

function StrToOpeType(OpeType: string): TOpeType;
begin
  if OpeType = 'DATE' then
    Result := otDate
  else if OpeType = 'HEURE' then
    Result := otHeure
  else if OpeType = 'REEL' then
    Result := otReel
  else
    Result := otString;
end;


function OpeTypeToStr(OpeType: TOpeType): string;
begin
  case OpeType of
    otDate: result := 'DATE';
    otHeure: result := 'HEURE';
    otString: result := 'STRING';
    otReel: result := 'REEL';
  end;
end;

function OperateurToStr(Operateur: TOperateur): string;
begin
  case operateur of
    Commence: result := 'COMMENCE';
    Egal: result := 'EGAL';
    Contient: result := 'CONTIENT';
    Superieur: result := 'SUPERIEUR';
    Inferieur: result := 'INFERIEUR';
    Different: result := 'DIFFERENT';
  end;
end;

function TVignettePlugin.WhereSQL: HString;
var
  i: integer;
  T: TOB;
  stSQL, Name, Value, EditMask: HString;
  Operateur: TOperateur;
  OpeType: TOpeType;
  Visible: boolean;
  Nam, Groupby, Having, Select, From, Join, RuptBy, SortBy: HString;
  OrWhere: HString;

  procedure Control2Criteres(var St: hstring);
  var
    w, ww: string;
    stText: string;
    Chaine: hstring;
    FType: string;
    s: string;

    function USHEUREVal(Val: HString): string;
    var
      D: TDateTime;
    begin
      D := strtotime(Val);
      if D < 0 then
        D := 0;
      UsHeureVal := UsTime(D);
    end;

    function USDATEVal(Val: Hstring): string;
    var
      D: TDateTime;
    begin
      D := strtoDate(val);
      if d < 0 then
        D := 0;
      UsDateVal := UsDateTime(D);
    end;

    function USDateSuivante(Val: string): string;
    var
      D: TDateTime;
    begin
      D := strtoDate(val);
      if d < 0 then
        D := 0;
      D := D + 1;
      result := UsDateTime(D);
    end;

  begin
    s := Nam;
//ddwriteln('ControleToCriteres,Name='+Name);
//ddwriteln('ControleToCriteres,T.NomTable='+T.NomTable);
    if Pos('.', s) > 0 then
      system.Delete(s, 1, Pos('.', s));
    if (Pos('_', s) <= 6) and (Pos('_', s) > 1) then
    begin
      (*
          if Copy(Nam, 1, 3) = 'Z_C' then
          begin
            nb := ValeurI(Copy(Nam, 4, 1));
            if P.Controls[Va[nb]] is THEdit then //pour unicode
              StAv[Nb] := TransComp(THValComboBox(C).Value, THValComboBox(P.Controls[O[nb]]).Value, THEdit(P.Controls[Va[nb]]).Text)
            else
              StAv[Nb] := TransComp(THValComboBox(C).Value, THValComboBox(P.Controls[O[nb]]).Value, TEdit(P.Controls[Va[nb]]).Text);
          end
          else
      *)
      //A FAIRE      if C is THMultiValComboBox then St := St + THMultiValComboBox(C).GetSQLValue else
      if ((Copy(Name, 1, 10) = 'XX_GROUPBY')) then
      else if ((Copy(Name, 1, 9) = 'XX_HAVING')) then
      else if ((Copy(Name, 1, 9) = 'XX_SELECT')) then
      else if ((Copy(Name, 1, 7) = 'XX_FROM')) then
      else if ((Copy(Name, 1, 7) = 'XX_JOIN')) then
      else if ((Copy(Name, 1, 10) = 'XX_ORDERBY')) then
      else if ((Copy(Name, 1, 11) = 'XX_VARIABLE')) then
      else if ((Copy(Name, 1, 10) = 'XX_RUPTURE')) then
      else if ((Copy(Name, 1, 8) = 'XX_WHERE')) then
      begin
        if Value <> '' then
        begin
          stText := Trim(Value);
          if Copy(Value, 1, 3) = 'OR ' then
            OrWhere := stText
          else if Copy(Value, 1, 6) = 'SELECT' then
            St := St + stText
          else
            St := St + ' AND (' + stText + ')';
        end;
      end
      else if (T.NomTable = 'T_DATAEDIT') and (Value <> '') then
      begin
        if ((OpeType = otDate)
          and (Value <> StDateVide)
          and (StrToDate(Value) <> iDate1900)
          and (StrToDate(Value) <> iDate2099)) then
        begin
          case Operateur of
            Commence: St := St + ' AND ' + Nam + ' LIKE "' + USDateVal(Value) + '%"';
            Egal:
              begin
                if ((EditMask = ChangeDateMask(EditMask, DateMask)) or (EditMask = ChangeDateMask(EditMask, DateMask99))) then
                  St := St + ' AND ' + Nam + ' >= "' + USDateVal(Value) + '"' + ' AND ' + Nam + ' < "' + USDATESUIVANTE(Value) + '"'
                else
                  St := St + ' AND ' + Nam + ' = "' + USDateVal(Value) + '"';
              end;
            Contient: St := St + ' AND ' + Nam + ' LIKE "%' + USDateVal(Value) + '%"';
            Superieur: St := St + ' AND ' + Nam + ' >= "' + USDateVal(Value) + '"';
            Inferieur:
              begin
                if ((EditMask = ChangeDateMask(EditMask, DateMask)) or (EditMask = ChangeDateMask(EditMask, DateMask99))) then
                  St := St + ' AND ' + Nam + ' < "' + USDATESUIVANTE(Value) + '"'
                else
                  St := St + ' AND ' + Nam + ' <= "' + USDateVal(Value) + '"';
              end;
            Different: St := St + ' AND ' + Nam + ' <> "' + USDateVal(Value) + '"';
          end;
        end
        else if ((OpeType = otHeure) and (Value <> '00:00:00') and (Value <> '99:99:99')) then
        begin
          case Operateur of
            Commence: St := St + ' AND ' + Nam + ' LIKE "' + USHeureVal(Value) + '%"';
            Egal: St := St + ' AND ' + Nam + ' = "' + USDateVal(Value) + '"';
            Contient: St := St + ' AND ' + Nam + ' LIKE "%' + USHeureVal(Value) + '%"';
            Superieur: St := St + ' AND ' + Nam + ' >= "' + USHeureVal(Value) + '"';
            Inferieur: St := St + ' AND ' + Nam + ' <= "' + USHeureVal(Value) + '"';
            Different: St := St + ' AND ' + Nam + ' <> "' + USHeureVal(Value) + '"';
          end;
        end
        else if OpeType = otString then
        begin
          case Operateur of
            //      {$IFNDEF AGL580}
            //          Commence: St := St + ' AND ' + Nam + ' LIKE "' + CheckdblQuote(THCritMaskEdit(C).Text) + '%"';
            //      {$ELSE}
            Commence:
              begin
                (*
                                if DisableFourchette and
                                  ((pos('*',Value)>0) or
                                   (pos('%',Value)>0) or
                                   (pos('?',Value)>0) or
                                   (pos('_',Value)>0) or
                                   (pos('[',Value)>0) or
                                   (pos(']',Value)>0)) then
                                begin
                                  Chaine := FindEtReplace(Value,'?','_',True);
                                  Chaine := FindEtReplace(Chaine,'*','%',True);
                                  if pos('%',Chaine)>0 then
                                    St := St + ' AND ' + Nam + ' LIKE "' + CheckdblQuote(Chaine) + '"'
                                  else
                                    St := St + ' AND ' + Nam + ' LIKE "' + CheckdblQuote(Chaine) + '%"';
                                end
                                else
                *)                               
                St := St + ' AND ' + Nam + ' LIKE "' + CheckdblQuote(Value) + '%"';
              end;
            //      {$ENDIF}
            Egal: St := St + ' AND ' + Nam + ' = "' + CheckdblQuote(WildChar(Value)) + '"';
            Contient: St := St + ' AND ' + Nam + ' LIKE "%' + CheckdblQuote(Value) + '%"';
            Superieur: St := St + ' AND ' + Nam + ' >= "' + CheckdblQuote(Value) + '"';
            Inferieur: St := St + ' AND ' + Nam + ' <= "' + CheckdblQuote(Value) + '"';
            Different: St := St + ' AND ' + Nam + ' <> "' + CheckdblQuote(Value) + '"';
          end;
        end
        else if OpeType = otReel then
        begin
          case Operateur of
            Superieur: St := St + ' AND ' + Nam + ' >= ' + StrFPoint(Valeur(Value));
            Inferieur: St := St + ' AND ' + Nam + ' <= ' + StrFPoint(Valeur(Value));
            Different: St := St + ' AND ' + Nam + ' <> ' + StrFPoint(Valeur(Value)); {PMJ545} // un oubli ??
          else
            St := St + ' AND ' + Nam + ' = ' + StrFPoint(Valeur(Value));
          end;
        end;
(*
        else if ((C is TEdit) and (TEdit(C).Text <> '')) then
        begin
          Ope := 0;
          if Copy(Nam, Length(Nam) - 3, 4) = '_SUP' then
          begin
            Nam := Copy(Nam, 1, Length(Nam) - 4);
            Ope := 1;
          end;
          if Copy(Nam, Length(Nam) - 3, 4) = '_INF' then
          begin
            Nam := Copy(Nam, 1, Length(Nam) - 4);
            Ope := 2;
          end;
          if Pos('*', Value) > 0 then
            Ope := 0;
          if Pos('?', Value) > 0 then
            Ope := 0;
          case Ope of
            2: St := St + ' AND ' + Nam + ' <="' + CheckdblQuote(Value) + '"';
            1: St := St + ' AND ' + Nam + ' >="' + CheckdblQuote(Value) + '"';
          else
            St := St + ' AND ' + Nam + ' LIKE "' + CheckdblQuote(WildChar(Value)) + '%"';
          end;
        end;
*)
      end
      else if T.NomTable = 'T_DATANUMEDIT' then
      begin
        if Value <> '' then
          St := St + ' AND ' + Nam + ' = ' + StrFPoint(Valeur(Value));
      end
      else if T.NomTable = 'T_DATAVALCOMBOBOX' then
      begin
        if Value <> '' then
        begin
          Ftype := ChampToType(Name);
          if (V_PGI.Driver = dbDB2) and ((Ftype = 'INTEGER') or (Ftype = 'SMALLINT') or (Ftype = 'DOUBLE') or (Ftype = 'RATE') or (Ftype = 'EXTENDED')) then
            Chaine := ''
          else
            chaine := '"';
          case Operateur of
            Commence: St := St + ' AND ' + Nam + ' LIKE "' + Value + '%"';
            Egal:
              begin
                if Pos(';', Value) > 0 then
                begin
                  w := '';
                  ww := Value;
                  while ww <> '' do
                  begin
                    w := w + Name + '=' + Chaine + ReadTokenSt(ww) + Chaine + ' ';
                    if ww <> '' then
                      w := w + ' or ';
                  end;
                  if W <> '' then
                    St := St + ' AND (' + w + ')';
                end
                else
                begin
                  if Pos('"', Value) > 0 then
                    St := St + ' AND (' + Nam + ' = ' + Chaine + WildChar(Value) + Chaine + ')'
                  else
                    St := St + ' AND ' + Nam + ' = ' + Chaine + WildChar(Value) + Chaine;
                end;
              end;
            Contient: St := St + ' AND ' + Nam + ' LIKE "%' + Value + '%"';
            Superieur: St := St + ' AND ' + Nam + ' >= ' + Chaine + Value + Chaine;
            Inferieur: St := St + ' AND ' + Nam + ' <= ' + Chaine + Value + Chaine;
            Different: St := St + ' AND ' + Nam + ' <> ' + Chaine + Value + Chaine;
          end;
        end
      end
      else if T.NomTable = 'T_DATACHECKBOX' then
      begin
        if ((Length(Value) = 2) and (Value = '-1')) or
          (char(Value[1]) in ['X', '1', 'O']) then
          St := St + ' AND ' + Nam + '="X"'
        else
          St := St + ' AND ' + Nam + '<>"X"';
      end;
    end;
  end;

begin
  if assigned(CritRequest) then
  else
    exit;
  try
if V_PGI.SAV then
if DirectoryExists('c:\xml') then
  CritRequest.Savetofile('c:\xml\CritRequest.txt',false,true,true);
    for i := 0 to CritRequest.Detail.Count - 1 do
    begin
      T := CritRequest.Detail[i];
      Name := uppercase(T['NAME'].Valeur);
      Visible := T['VISIBLE'].Valeur;
      if not visible then
        continue;
      if T.NomTable = 'T_DATAVALCOMBOBOX' then
      begin
        if Name = 'DATABASES' then
          continue; //reservé nom du dossier
        Operateur := StrToOperateur(T.Detail[0]['OPERATEUR'].Valeur);
      end
      else if T.NomTable = 'T_DATACHECKBOX' then
      begin
        if Name = 'CBAUTODATA' then
           continue;//reservé
        if Name = 'CBAUTOSAVE' then
           continue;//reservé
        ;
      end
      else if T.NomTable = 'T_DATAEDIT' then
      begin
if V_PGI.SAV then
if DirectoryExists('c:\xml') then
  T.Savetofile('c:\xml\t_dataedit.txt',false,true,true);
ddwriteln('T.Detail.Count=' + inttostr(T.Detail.Count));
        Operateur := StrToOperateur((T.Detail[0]['OPERATEUR'].Valeur));
        OpeType := StrToOpeType(T.Detail[0]['OPETYPE'].Valeur);
ddwriteln('Operateur=' + OPerateurToStr(Operateur));
ddwriteln('OpeType=' + IntToStr(Ord(OpeType)) + '=' + OpeTypeToStr(OpeType));
      end
      else if T.NomTable = 'T_DATANUMEDIT' then
      else
        continue;

      Value := T['VALUE'].Valeur;
      Nam := Name;
ddwriteln('WhereSQL=' + Name + ', Value=' + Value);
      if (Nam <> '') and (Nam[Length(Nam)] = '_') then
        System.Delete(Nam, Length(Nam), 1);
      Nam := FindEtReplace(nam, '__', '.', FALSE); //CR modif pour gérer les alias du 4/10/02

      if ((Copy(Name, 1, 10) = 'XX_GROUPBY')) then
      begin
        Groupby := Value;
      end;
      if ((Copy(Name, 1, 9) = 'XX_HAVING')) then
      begin
        Having := Value;
      end;
      if ((Copy(Name, 1, 9) = 'XX_SELECT')) then
      begin
        Select := Value;
      end;
      if ((Copy(Name, 1, 7) = 'XX_FROM')) then
      begin
        From := Value;
      end;
      if ((Copy(Name, 1, 7) = 'XX_JOIN')) then
      begin
        Join := Value;
      end;
      if ((Copy(Name, 1, 10) = 'XX_ORDERBY')) then
      begin
        Sortby := Value;
      end;
      if (Copy(Name, 1, 10) = 'XX_RUPTURE') then
      begin
        RuptBy := RuptBy + ';' + Copy(Name, 11, 20) + '=' + Value;
      end;
ddwriteln('Control2Criteres:stSQL=' + stSQL);
      Control2Criteres(stSQL);

    end;
  finally
    stSQL := Trim(stSQL);
    if Copy(stSQL, 1, 3) = 'AND' then
      Delete(stSQL, 1, 3);
    ddWriteLN('==== SQL=' + stSQL);
    result := Trim(stSQL);
  end;
end;

function TVignettePlugin.GetCritereValue(AControlName: string): String;
var
//  i: integer;
  Name, Nom: string;
  T: TOB;
begin
  result := '';
  Nom := uppercase(AControlName);
//EPZ  for i := 0 to CritRequest.Detail.Count - 1 do
//EPZ  begin
//EPZ    T := CritRequest.Detail[i];
    T := GetCritInstance(AControlName);
    if T <> nil then
	  begin
      Name := uppercase(T['NAME'].Valeur);
      if Name = Nom then
        result := T['VALUE'].Valeur;
  end;
//EPZ  end;
end;

procedure TVignettePlugin.SetCritereValue(AControlName: string; AValue: Variant);
var
  T: TOB;
  s: hstring;
begin
  T := GetCritInstance(AControlName);
  if T <> nil then
  begin
    case VarType(AValue) of
      varDouble: s := Format('%f', [AValue]);
      varInteger: s := Format('%d', [AValue]);
      varDate: s := DateTimeToStr(AValue);
    else
      s := AValue;
    end;
    T.PutValue('VALUE', s);
  end;
end;

function TVignettePlugin.GetCritInstance(ControlName: string): Tob;
var
  T: TOB;
begin
  Result := CritToSend.FindFirst(['NAME'], [ControlName], True); // si on trouve dans DatasToSend, ok
  if Result = nil then
  begin //sinon on cherche dans DatasRequest
    T := CritRequest.FindFirst(['NAME'], [ControlName], True);
    if T <> nil then //on réplique les données dans DatasToSend
    begin
      Result := Tob.Create(T.NomTable, CritToSend, -1);
      Result.Assign(T);
    end;
  end;
end;

function TVignettePlugIn.GetDateTime(NomChp: string; Critere: Boolean = True): TDateTime;
var
  V: Variant;
begin
  Result := iDate1900;
  if Critere then
    V := GetCritereValue(NomChp)
  else
    V := GetControlValue(NomChp);
  if (VarType(V) > 1) then
  try
    Result := VarAsType(V, varDate);
  except
  end;
end;

function TVignettePlugIn.GetDouble(NomChp: string; Critere: Boolean = True): Double;
var
  V: Variant;
begin
  Result := 0;
  if Critere then
    V := GetCritereValue(NomChp)
  else
    V := GetControlValue(NomChp);
  if (VarType(V) > 1) then
  try
    Result := VarAsType(V, varDouble);
  except
  end;
end;

function TVignettePlugIn.GetInteger(NomChp: string; Critere: Boolean = True): Integer;
var
  V: Variant;
begin
  Result := 0;
  if Critere then
    V := GetCritereValue(NomChp)
  else
    V := GetControlValue(NomChp);
  if (VarType(V) > 1) then
  try
    Result := VarAsType(V, varInteger);
  except
  end;
end;

function TVignettePlugIn.GetString(NomChp: string; Critere: Boolean = True): string;
var
  V: Variant;
begin
  Result := '';
  if Critere then
    V := GetCritereValue(NomChp)
  else
    V := GetControlValue(NomChp);
  if (VarType(V) > 1) then
  try
    Result := VarAsType(V, varString);
  except
  end;
end;

//Retourne le dossier
function TVignettePlugIn.GetDossier: string;
var
  stDossier, StDriver,
    StServer, StPath,
    StUser, StPassWord,
    StODBC, StOptions,
    StGroup, StShare: string;
begin
  Result := '';
  if IsCritereVisible('DATABASES') then
  begin
    stDossier := GetString('DATABASES');
    ChargeDBParams(stDossier, StDriver, StServer, StPath, Result, StUser, StPassWord, StODBC, StOptions, StGroup, StShare)
  end;
end;

(*
function TVignettePlugIn.GetDossier: string;
begin
  Result := GetString('DATABASES');
end;
*)

//Ouvre la requête en fonction du dossier de la vignette.
//Si Dossier = '', la requête s'exécute sur la base de connection
function TVignettePlugIn.OpenSelect(Requete: string; ADossier: string = ''): TQuery;
var
  s: string;
begin
  if ADossier = '' then
    s := ListeSoc
  else
    s := ADossier;
  if s = '' then
    s := Dossier;
  Result := OpenSQL(Requete, True, -1, '', False, s);
end;
(*
begin
  if ADossier='' then Result := OpenSQL(Requete, True, -1, '', False, Dossier)
  else Result := OpenSQL(Requete, True, -1, '', False, ADossier);
end;
*)

//Retourne la Tob contenant la définition d'une colonne de la grille
function TVignettePlugIn.GetTobProperty(Grille, Chp: string): TOB;

var
  T: TOB;
begin
  Result := nil;
  T := GetDataInstance(Grille);
  if T <> nil then
    Result := T.FindFirst(['CHAMP'], [Chp], True);
end;

//Definit le titre d'une colonne de la grille
procedure TVignettePlugIn.SetTitreCol(Grille, Chp, Lib: string);
var
  T: TOB;
begin
  T := GetTobProperty(Grille, Chp);
  if T <> nil then
    T.SetString('LIBELLE', Lib);
end;

//Definit le format des données d'une colonne de la grille
procedure TVignettePlugIn.SetFormatCol(Grille, Chp, Format: string);

var
  T: TOB;
begin
  T := GetTobProperty(Grille, Chp);
  if T <> nil then
    T.SetString('ALIGN', Format);
end;

//Definit la largeur d'une colonne de la grille
procedure TVignettePlugIn.SetWidthCol(Grille, Chp: string; Largeur: Integer);
var
  T: TOB;
begin
  T := GetTobProperty(Grille, Chp);
  if T <> nil then
    T.SetString('WIDTH', IntToStr(Largeur));
end;

procedure TVignettePlugIn.SetVisibleCol(Grille, Chp: string; VisibleOk: Boolean);
begin
  SetWidthCol(Grille, Chp, -1);
end;

(*
procedure TVignettePlugIn.SetVisibleCol(Grille, Chp: string; VisibleOk: Boolean);
var
  T: TOB;
begin
  T := GetTobProperty(Grille, Chp);
  if T <> nil then
    T.SetString('VISIBLE', BoolToStr(VisibleOk));
end;
*)

// récupère la liste des dossiers pour un groupement donné
function GetListeMultiDossierDB(Groupement, DBName: string): string;
var
  L: TStringList;
  stsql, st, base, dossier: string;
  Q: TQuery;
begin
  result := '';
  if DBName = LookupCurrentSession.DBName then
  begin
    result := GetListeMultiDossier(Groupement);
    exit;
  end;
  stSQL := 'SELECT YMD_CODE,YMD_DETAILS FROM YMULTIDOSSIER WHERE YMD_CODE="' + Groupement + '"';
  Q := OpenSQL(stSQL, true, -1, '', false, DBName);
  if not Q.EOF then
  begin
    L := TStringList.Create;
    L.Text := Q.Fields[1].AsString;
    if L.Count = 2 then
    begin
      result := '';
      st := L[0]; //Liste des dossiers
      while st <> '' do
      begin
        base := ReadTokenSt(st);
        dossier := ReadTokenPipe(base, '|');
        result := result + base + ';';
      end;
    end;
  end;
  Ferme(Q);
  FreeAndNil(L);
end;

function TVignettePlugIn.GetSQLListe(AUserName, AListName: string): string;
var
  Q: THQuery;
  stMulti, stDbName, stSQL: string;
  StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare: string;
begin
  Q := THQuery.Create(nil);
  Q.DatabaseName := LookupCurrentSession.DBName;
  Q.CacheBlobs := FALSE;
  Q.UpdateMode := upWhereChanged;
  Q.SQL.Clear;
  Q.MajComboAvance := FALSE;
  Q.Criteres := WhereSQL;
  Q.UserListe := AUserName;
  Q.Liste := AListName;
  Q.ChargeChampSQL;
  Q.FabricSQL;
  stSQL := Q.SQL.Text; //@@ la requete est déjà traduite
  Ferme(Q);
  stDBName := GetDossier;//GetCritereValue('DATABASES');
  if stDBName <> '' then
    ChargeDBParams(stDBName, StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare)
  else
    stBase := LookupCurrentSession.DbName;

  stMulti := GetCritereValue('MULTIDOSSIERPORTAIL');
  if stMulti <> '' then
    stBase := GetListeMultiDossierDB(stMulti, stDBName);

  result := '@@' + sqlToMultiBases(stsql, stBase);
end;

function TVignettePlugIn.GetSQLCount(ATableName, AWhereSt: string): integer;
var
  st, stDBName: string;
  Q: TQuery;
  StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare: string;
begin
  result := 0;
  st := 'SELECT COUNT(*) FROM ' + ATableName;
  st := InsertSQLWhere(st, AWhereSt);

  stDBName := GetDossier;//GetCritereValue('DATABASES');
  if stDBName <> '' then
    ChargeDBParams(stDBName, StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare)
  else
    stBase := LookupCurrentSession.DbName;

  Q := OpenSelect(st, stBase);
  if assigned(Q) then
    result := Q.Fields[0].AsInteger - 1;
  Ferme(Q);
end;

function TVignettePlugIn.GetSQLFiche(ATableName: string; var ARowCount: integer): string;
var
  stSQL, stField, stWhereSQL, stSelect, stMulti, stDBName, stOrderBy: string;
  StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare: string;
  i: integer;
begin
  stWhereSQL := WhereSQL;
  // indique que l'on veut le dernier il faut fair un select count
  ARowCount := GetSQLCount(ATableName, stWhereSQL);
  for i := 0 to ControlCount - 1 do
  begin
    stField := GetControlName(i);
    if ChampToNum(stField) > -1 then
    begin
      if stSelect <> '' then
        stSelect := stSelect + ',';
      stSelect := stSelect + stField;
    end;
  end;
  if stSelect = '' then
    stSelect := '*';
  stSQL := 'SELECT ' + stSelect + ' FROM ' + ATableName;
  stSQL := InsertSQLWhere(stSQL, stWhereSQL);
  stOrderBy := TableToCle1(ATableName);
  if stOrderBy <> '' then
    stSQL := stSQL + ' ORDER BY ' + stOrderBy;

  stDBName := GetDossier;//GetCritereValue('DATABASES');
  if stDBName <> '' then
    ChargeDBParams(stDBName, StDriver, StServer, StPath, Stbase, StUser, StPassWord, StODBC, StOptions, StGroup, StShare)
  else
    stBase := LookupCurrentSession.DbName;

  stMulti := GetCritereValue('MULTIDOSSIERPORTAIL');
  if stMulti <> '' then
    stBase := GetListeMultiDossierDB(stMulti, stDBName);

  result := '@@' + sqlToMultiBases(stsql, stBase);
end;

procedure TVignettePlugIn.SetSqlMutliDossier(var SQL: string);
var
  s: string;
begin
  s := GetListeSoc;
  if s <> '' then
    SQL := SqlToMultiBases(SQL, s);
end;

function TVignettePlugin.GetListeSoc: string;
begin
  Result := Regroupement;
  if Result <> '' then
    Result := GetListeMultiDossierDB(Result, Dossier);
end;

function TVignettePlugin.GetRegroupement: string;
begin
  Result := '';
  if IsCritereVisible('MULTIDOSSIERPORTAIL') then
    Result := GetString('MULTIDOSSIERPORTAIL');
end;

function TVignettePlugin.IsCritereVisible(AControlName: string): Boolean;
var
  n: Integer;
  T: TOB;
  Name: string;
  Nom: string;
begin
  Result := False;
  if not Assigned(CritRequest) then
    Exit;
  Nom := UpperCase(AControlName);
  for n := 0 to CritRequest.Detail.Count - 1 do
  begin
    T := CritRequest.Detail[n];
    Name := UpperCase(T['NAME'].Valeur);
    if Name = Nom then
    begin
      Result := T['VISIBLE'].Valeur <> 0;
      Break;
    end;
  end;
end;

procedure TVignettePlugin.ConvertFieldValue(laTOB: TOB);
var
  i, j: integer;
  Nam, Value: string;
begin
  for i := 0 to LaTob.detail.Count - 1 do
  begin
    (*
      for j:=0 to DataTob.detail[i].ChampsSup.Count-1 do
      begin
        Nam := TCS(DataTob.detail[i].ChampsSup[j]).Nom;
    ddwriteln('stSQL=Nam:'+Nam);
        if ChampToType(Nam)='BOOLEAN' then
        begin
          Value := DataTob.detail[i].GetValue(Nam);
    ddwriteln('stSQL=Value:'+Value);
          if Value='-' then DataTob.detail[i].PutValue(Nam,'-1')
          else DataTob.detail[i].PutValue(Nam,'0');
        end;
      end;
    *)
    for j := 0 to LaTob.detail[i].NbChamps - 1 do
    begin
      Nam := LaTob.detail[i].GetNomChamp(j);
      if ChampToType(Nam) = 'BOOLEAN' then
      begin
        Value := LaTob.detail[i].GetValue(Nam);
        if Value = '-' then
          LaTob.detail[i].PutValue(Nam, '-1')
        else
          LaTob.detail[i].PutValue(Nam, '0');
      end;
    end;
  end;
end;

procedure TVignettePlugin.SetTobViewerColumns(ControlName: string; Datas: Tob; AColNames: string);
var
  i: integer;
  T, TNew, TCol: TOB;
  Name, Lib, Typ: string;
  Visu: boolean;

  function CalcWidth: integer;
  var
    il: integer;
  begin
    if (Typ = 'COMBO') then
      Result := 16
    else if (Pos('CHAR', Typ) > 0) then
    begin
      il := ValeurI(Typ);
      if il <= 20 then
        Result := (il * 3) + 2
      else if il <= 50 then
        Result := (il * 2)
      else
        Result := 100;
    end
    else if Typ = 'BOOLEAN' then
      Result := 10
    else if Typ = 'DATE' then
      Result := 24
    else if Typ = 'INTEGER' then
      Result := 24
    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      Result := 30
    else if Typ = 'BLOB' then
      Result := 8
    else
      Result := 16;
  end;

  function CalcAlign: string;
  begin
    if (Typ = 'COMBO') then
      Result := 'taLeftJustify'
    else if (Pos('CHAR', Typ) > 0) then
      Result := 'taLeftJustify'
    else if Typ = 'BOOLEAN' then
      Result := 'taCenterJustify'
    else if Typ = 'DATE' then
      Result := 'taCenterJustify'
    else if Typ = 'INTEGER' then
      Result := 'taRightJustify'
    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      Result := 'taRightJustify'
    else if Typ = 'BLOB' then
      Result := 'taLeftJustify'
    else
      Result := 'taLeftJustify'
  end;

  function ChampToCSType(ch: TCS): string;
  begin
    case ch.CSType of
      CSTInteger: result := 'dtInteger';
      CSTDouble: result := 'dtDouble';
      CSTDate: result := 'dtDateTime';
      CSTBoolean: result := 'dtBoolean';
      CSTString: result := 'dtString';
      CSTBase64Binary: result := 'dtString';
      //      CSTHexBinary:  Typ := '';
      CSTBooleanStr: result := 'dtBoolean';
    else
      result := 'dtString';
    end;
  end;

begin
  if Datas.Detail.count = 0 then
    exit;

  T := GetDataInstance(ControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATATOBVIEWERCOLUMNS' then
      begin
        T.detail[0].Free;
        break;
      end;
    end;
    TNew := Tob.Create('T_DATATOBVIEWERCOLUMNS', T, -1);
    T := Datas.Detail[0];
    for i := 0 to T.ChampsSup.Count - 1 do
    begin
      TCol := Tob.Create('T_DATATOBVIEWERCOLUMN', TNew, -1);
      Name := uppercase(TCS(T.ChampsSup[i]).Nom);
      Typ := ChampToCSType(TCS(T.ChampsSup[i]));
      Lib := ChampToLibelle(Name);
      if Lib = '??' then
        Lib := Name;
      Visu := (Trim(AColNames)='') or (Pos(Name,AColNames)>0);
      TCol.AddChampSupValeur('CHAMP', name, CSTString);
      TCol.AddChampSupValeur('LIBELLE', Lib, CSTString);
      TCol.AddChampSupValeur('WIDTH', CalcWidth, CSTInteger);
      TCol.AddChampSupValeur('ALIGN', CalcAlign, CSTString);
      TCol.AddChampSupValeur('TYPE', CalcType(Typ), CSTString);
      TCol.AddChampSupValeur('VISIBLE', BoolToStr(Visu), CSTBoolean);
      TCol.AddChampSupValeur('FIXEDPOS','fp2Center', CSTString);
      TCol.AddChampSupValeur('CALCOP', 'opNone', CSTString);
      TCol.AddChampSupValeur('REGINDEX',-1, CSTInteger);
      TCol.AddChampSupValeur('REGSORTORDER','soNone', CSTString);
      TCol.AddChampSupValeur('SORTORDER','soNone', CSTString);
      TCol.AddChampSupValeur('SORTINDEX',-1, CSTInteger);
    end;
  end;
end;

procedure TVignettePlugin.PutTobViewerDetail(ControlName: string; Datas: Tob; AColNames: string = '');
var
  i: integer;
  T, TNew: TOB;
begin
  SetTobViewerColumns(ControlName, Datas, uppercase(AColNames)); //génère les colonnes
  for i := 0 to Datas.Detail.count - 1 do
    Datas.detail[i].ReaffecteNomTable('T_DATATOBVIEWERLIGNE', true); //pour le message XML
  T := GetDataInstance(ControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATATOBVIEWERLIGNES' then
      begin
        T.detail[0].Free;
        break;
      end;
    end;
    TNew := Tob.Create('T_DATATOBVIEWERLIGNES', T, -1);
    while Datas.detail.count > 0 do
      Datas.Detail[0].ChangeParent(TNew, -1, True);
  end;
end;

//C_COl, F_SRT, A_DIV, R_HDR, R_REG, P_IMP
function TVignettePlugin.GetTobViewerColumn(AControlName, AColName: string): TOB;
var
  i,j: integer;
  T, TT: TOB;
begin
  result := nil;
  T := GetDataInstance(AControlName);
  if T <> nil then
  begin
    for i := 0 to T.Detail.count - 1 do
    begin
      if T.detail[0].NomTable = 'T_DATATOBVIEWERCOLUMNS' then
      begin
        TT := T.Detail[0];
        for j := 0 to TT.Detail.count - 1 do
        begin
          if TT.detail[0].NomTable = 'T_DATATOBVIEWERCOLUMNS' then
            result := TT.detail[0].FindFirst(['NAME'],[AColName],false);
        end;
      end;
    end;
  end;
end;

procedure TVignettePlugin.SetTobViewerCalcOpField (AControlName,AColName: string; ACalcOp: TCalcOp);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
    TCol.PutValue('CALCOP','opNone');
end;

procedure TVignettePlugin.SetTobViewerCaptionField (AControlName,AColName: string; ACaption: string);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
    TCol.PutValue('LIBELLE',ACaption);
end;

procedure TVignettePlugin.SetTobViewerWidthField (AControlName,AColName: string; AWidth: integer);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
    TCol.PutValue('WIDTH',AWidth);
end;

function EncodeSo (Value: TSortOrder) : string;
const SoStr: array [TSortOrder] of string = ('soAscending', 'soDescending', 'soNone') ;
begin
  Result := SoStr [Value] ;
end;

procedure TVignettePlugin.SetTobViewerRegField (AControlName,AColName: string; ARegIndex: integer; ARegSortOrder: TSortOrder = soNone);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
  begin
    TCol.PutValue('REGINDEX',ARegIndex);
    TCol.PutValue('REGSORTORDER',EncodeSo(ARegSortOrder));
  end;
end;

procedure TVignettePlugin.SetTobViewerSortField (AControlName,AColName: string; ASortIndex: integer; ASortOrder: TSortOrder = soNone);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
  begin
    TCol.PutValue('SORTINDEX',ASortIndex);
    TCol.PutValue('SORTORDER',EncodeSo(ASortOrder));
  end;
end;

function EncodeFp (Value: TFixedPos) : string;
const FpStr: array [TFixedPos] of string = ('fp1Left', 'fp3Right', 'fp2Center') ;
begin
  Result := FpStr [Value] ;
end;

procedure TVignettePlugin.SetTobViewerGroupField (AControlName,AColName: string; AFixedPos: TFixedPos);
var
  TCol: TOB;
begin
  TCol := GetTobViewerColumn(AControlName,AColName);
  if TCol <> nil then
    TCol.PutValue('FIXEDPOS',EncodeFp(AFixedPos));
end;

function TVignettePlugin.IsForceDataBaseName(var databasename: string; PlugInName: string): boolean;
var
  ini: TIniFile;
begin
databasename := '';
if PlugInName<>'' then
begin
  ini:= TIniFile.Create(ExtractFilePath(Application.ExeName)+'portail.ini');
  databasename := ini.ReadString('DATABASES',PlugInName,'');
  ddwriteln('~~~~~~~~~~PlugInName='+PlugInName);
  ddwriteln('~~~~~~~~~~databasename='+databasename);
  result := databasename<>'';
  FreeAndNil(ini);
end;
end;

end.



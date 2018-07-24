unit uWSDataService;

interface

uses
  Classes
  , UConnectWSConst
  ;

type
  TReadWSDataService = class (TObject)
  private
    class function GetODataOperator(Operator : string) : string;
  public
    class function GetData(DSType : T_WSDataService; ServerName, FolderName : string; var TslResult, TslViewFields : TStringList; TslFilter : TStringList=nil; KnownUrl : string=''; TslFieldsType : TStringList=nil) : string;
  end;

implementation

uses
  CommonTools
  , WinHttp_TLB
  , uLkJSON
  , SysUtils
  , Variants
  {$IF not defined(APPSRV)}
  , UConnectWSCEGID
  {$IFEND !APPSRV}
  ;

class function TReadWSDataService.GetODataOperator(Operator : string) : string;
begin
  case Tools.CaseFromString(Operator, ['=', '<>', '>', '>=', '<', '<=', 'AND', 'OR', 'NOT']) of
    {=}   0 : Result := 'eq';
    {<>}  1 : Result := 'ne';
    {>}   2 : Result := 'gt';
    {>=}  3 : Result := 'ge';
    {<}   4 : Result := 'lt';
    {>=}  5 : Result := 'le';
    {AND} 6 : Result := 'and';
    {OR}  7 : Result := 'or';
    {NOT} 8 : Result := 'not';
  else
    Result := '';
  end;
end;

{ DSType : WebApi
  TslResult : TstringList renvoyée
  TslFilter : TStringList des filtres sous la forme :
  . 1ère valeur : opérateur logique (; s'il n'y en a pas, sinon (ex) AND;)
  . 2ème valeur : nom du champ
  . 3ème valeur : opérateur (à traduire pour oData)
  . 4ème valeur :
    . valeur
      ou
    . ( début de groupe
    . ) fin de groupe

  EXEMPLE :
  Avoir la liste des clients et fournisseurs créés ou modifiés après le 01/04/2018 13:53:48
    TSlFilter.Add(';;;(');                                         : début groupe
    TSlFilter.Add(';T_NATUREAUXI;=;CLI');                          : condition
    TSlFilter.Add('OR;T_NATUREAUXI;=;FOU');                        : condition
    TSlFilter.Add(';;;)');                                         : fin groupe
    TSlFilter.Add('AND;T_DATEMODIF;>=;'2018-04-01T13:53:48.000Z'); : condition
  Le filtre ci dessus correspond :
    pour oData : $filter=(Nature eq 'CLI' or Nature eq 'FOU') and DateModif ge 02018-04-01T13:53:48.000Z
    pour SQL   : WHERE (T_NATUREAUXI='CLI OR T_NATUREAUXI='FOU') AND T_DATEMODIF >= '202018-04-01T13:53:48.000Z'

  ATTENTION, si passage de date en paramètre, celle-ci doit être au format "string"
}
class function TReadWSDataService.GetData(  DSType            : T_WSDataService // Type
                                          ; ServerName        : string          // Serveur
                                          ; FolderName        : string          // Dossier
                                          ; var TslResult     : TStringList     // TstringList de retour
                                          ; var TslViewFields : TStringList     // Liste des champs
                                          ; TslFilter         : TStringList=nil // Filtres
                                          ; KnownUrl          : string=''       // Url si déjà connue
                                          ; TslFieldsType     : TStringList=nil // Lsite des champs avec types pour éviter des qry
                                         ) : string;
var
  http     : IWinHttpRequest;
  Url      : string;
  Response : string;
  Values   : string;
  Value    : string;
  Fields   : string;
  Alias    : string;
  Field    : string;
  NewUrl   : string;
  JSon     : TlkJSONBase;
  Items    : TlkJSONBase;
  Item     : TlkJSONBase;
  Cpt      : Integer;
  Cpt1     : Integer;
//  TslFieldsType : TStringList;

  function GetFieldName(Value : string) : string;
  begin
    Result := Copy(Value, 1, pos('=', Value) -1);
  end;

  function GetAliasName(Value : string) : string;
  begin
    Result := Copy(Value, pos('=', Value) + 1, Length(Value));
  end;

  function GetFieldsFromDSType : string;
  var
    iPos : integer;
  begin
    iPos := TslViewFields.IndexOfName(TGetFromDSType.dstWSName(DSType));
    if iPos > -1 then
      Result := TslViewFields.ValueFromIndex[iPos];
  end;

  function GetFilter : string;
  var
    CptF            : integer;
    CptT            : integer;
    Filter          : string;
    FilterField     : string;
    FilterOperator  : string;
    FilterValue     : string;
    FilterLogicalOp : string;
    Values          : string;
    Field           : string;
    Alias           : string;
    FieldCache      : string;
    Value           : Variant;
    StartEndGroup   : boolean;
    IsFieldDate     : boolean;

    function GetPrefix(FieldName : string) : string;
    begin
      Result := copy(FieldName, pos('_', FieldName), length(FieldName));
    end;

  begin
    if (Assigned(TslFilter)) and (TslFilter[0] <> WSCDS_EmptyValue) then
    begin
      for CptF := 0 to pred(TslFilter.Count) do
      begin
        Filter          := TslFilter.Strings[CptF];
        FilterLogicalOp := GetODataOperator(Tools.ReadTokenSt_(Filter, ';'));
        FilterField     := Tools.ReadTokenSt_(Filter, ';');
        FilterOperator  := GetODataOperator(Tools.ReadTokenSt_(Filter, ';'));
        FilterValue     := Tools.ReadTokenSt_(Filter, ';');
        Values          := GetFieldsFromDSType;
        while Values <> '' do
        begin
          Value := Tools.ReadTokenSt_(Values, ';');
          Field := GetFieldName(Value);
          Alias := GetAliasName(Value);
          if Field = FilterField then
            Break;
        end;
        StartEndGroup := ((FilterValue = '(') or (FilterValue = ')'));
        if not StartEndGroup then
        begin
          IsFieldDate := False;
          if assigned(TslFieldsType) then
          begin
            for CptT := 0 to pred(TslFieldsType.count) do
            begin
              FieldCache := copy(TslFieldsType[CptT], 1, pos(ToolsTobToTsl_Separator, TslFieldsType[CptT])-1);
              if dsType = wsdsAnalyticalSection then // Exception pour analytique (préfixe V10 = S, préfixe Y2 = CSP)
                IsFieldDate := ((pos('^DATE^', TslFieldsType[CptT]) > 0) and (GetPrefix(FilterField) = GetPrefix(FieldCache)))
              else
                IsFieldDate := ((FieldCache = FilterField) and (pos('^DATE^', TslFieldsType[CptT]) > 0));
              if IsFieldDate then
                break;
            end;
          end else
            IsFieldDate := (Tools.GetFieldType(FilterField{$IF defined(APPSRV)}, ServerName, FolderName{$IFEND !APPSRV}) = ttfDate);
          if IsFieldDate then
            FilterValue := FormatDateTime('yyyy-mm-dd', Int(StrToDateTime(FilterValue))) + 'T' +  FormatDateTime('hh:nn:ss.zzz', StrToDateTime(FilterValue)) + 'Z'
          else
            FilterValue := '''' + FilterValue + '''';
          Result := Result
                  + Tools.iif(FilterLogicalOp <> '', ' ' + FilterLogicalOp, '')
                  + ' ' + Alias
                  + ' ' + FilterOperator
                  + ' ' + FilterValue
                    ;                                  
        end else
          Result := Result + FilterValue;
      end;                                              
      Result := '?$filter=' + Result;
    end else                                                                           
      Result := '';
  end;

begin
  if (Assigned(TslResult)) and (DSType <> wsdsNone) then
  begin
    if KnownUrl = '' then
      Url := 'http://'
           + ServerName
           + '/CegidDataService/odata'
           + '/' + FolderName
           + '/' + TGetFromDSType.dstWSName(DSType)
           + GetFilter
    else
      Url := KnownUrl;
    http := CoWinHttpRequest.Create;
    try
      http.SetAutoLogonPolicy(0);
      http.Open('GET', Url, False);
      http.Send(EmptyParam);
      try
        if http.Status = 200 then
        begin
          Response := '[' + http.ResponseText + ']';
          JSon := TlkJSON.ParseText(Response);
          for Cpt := 0 to pred(JSon.Count) do
          begin
            Items := JSon.Child[Cpt].Field['value'];
            for Cpt1 := 0 to Pred(Items.Count) do
            begin
              Item := Items.Child[Cpt1];
              TslResult.Add(WSCDS_IndiceField + IntToStr(Cpt1) + '#=' + IntToStr(Cpt1));
              Values := GetFieldsFromDSType;
              while Values <> '' do
              begin
                Fields := Tools.ReadTokenSt_(Values, ';');
                if Fields <> '' then
                begin
                  Field  := GetFieldName(Fields);
                  Alias  := GetAliasName(Fields);
                  if Tools.GetFieldType(Field{$IF defined(APPSRV)}, ServerName, FolderName{$IFEND APPSRV}) = ttfBoolean then
                    Value := Tools.iif(Item.Field[Alias].Value, 'X', '-')
                  else
                    Value := VarToStr(Item.Field[Alias].Value);
                  TslResult.Add(Field + '=' + Value);
                end;
              end;
            end;
          end;
          if pos(WSCDS_NextUrlValue, Response) > 0 then
          begin
            NewUrl := Copy(Response, Pos(WSCDS_NextUrlValue, Response) + Length(WSCDS_NextUrlValue), Length(Response)-2);
            NewUrl := Copy(NewUrl, 1, Pos('"', NewUrl) -1);
            GetData(DSType, ServerName, FolderName, TslResult, TslViewFields, TslFilter, NewUrl);
          end;
          Result := WSCDS_GetDataOk;
        end else
          Result :=  Format('Erreur %s - %s (%s)', [IntToStr(http.Status), http.StatusText, Url]); //http.ResponseText]);;
      except
        Result := Format('Erreur interne (exception) lors de l''appel de %s', [Url]); //http.ResponseText]);;
      end;
    finally
      http := nil;
    end;
  end else
    Result := 'TslResult ou DSType non assigné.';
end;

end.

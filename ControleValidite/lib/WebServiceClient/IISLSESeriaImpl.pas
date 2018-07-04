{ Fichier d'implémentation invocable pour TIISLSESeria qui implémente IIISLSESeria }

unit IISLSESeriaImpl;

interface

uses InvokeRegistry, Types, XSBuiltIns, IISLSESeriaIntf,DB, ADODB,
     AdoInt, OleDB,forms;

type

  { TIISLSESeria }
  TIISLSESeria = class(TInvokableClass, IIISLSESeria)
  public
    function GetSeriaTiers(const CodeTiers : String): TSeriaDBArray; stdcall;
  end;

implementation

function TIISLSESeria.GetSeriaTiers( const CodeTiers: String): TSeriaDBArray;
var LSeriaArray : TSeriaDBArray;
		QQ : TADOQuery;
begin
  SetLength  (LSeriaArray, 10);
  QQ := TADOQuery.Create(application);
  TRY
    QQ.SQL.Add('SELECT ');
    QQ.Open;
    
    Result := LseriaArray;
  finally
    QQ.free;
    QQ := nil;
  end;
end;

initialization
  { Les classes invocables doivent être recensées }
  InvRegistry.RegisterInvokableClass(TIISLSESeria);
  RemClassRegistry.RegisterXSClass(TSeriaDB);
  RemClassRegistry.RegisterXSInfo(TypeInfo(TSeriaDBArray));
end.

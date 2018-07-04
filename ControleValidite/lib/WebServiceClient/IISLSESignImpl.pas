{ Fichier d'implémentation invocable pour TIISLSESign qui implémente IIISLSESign }

unit IISLSESignImpl;

interface

uses InvokeRegistry, Types, XSBuiltIns, IISLSESignIntf;

type

  { TIISLSESign }
  TIISLSESign = class(TInvokableClass, IIISLSESign)
  public
    function echoEnum(const Value: TEnumTest): TEnumTest; stdcall;
    function echoDoubleArray(const Value: TDoubleArray): TDoubleArray; stdcall;
    function echoMyEmployee(const Value: TMyEmployee): TMyEmployee; stdcall;
    function echoDouble(const Value: Double): Double; stdcall;
  end;

implementation

function TIISLSESign.echoEnum(const Value: TEnumTest): TEnumTest; stdcall;
begin
  { A FAIRE : Implémenter la méthode echoEnum }
  Result := Value;
end;

function TIISLSESign.echoDoubleArray(const Value: TDoubleArray): TDoubleArray; stdcall;
begin
  { A FAIRE : Implémenter la méthode echoDoubleArray }
  Result := Value;
end;

function TIISLSESign.echoMyEmployee(const Value: TMyEmployee): TMyEmployee; stdcall;
begin
  { A FAIRE : Implémenter la méthode echoMyEmployee }
  Result := TMyEmployee.Create;
end;

function TIISLSESign.echoDouble(const Value: Double): Double; stdcall;
begin
  { A FAIRE : Implémenter la méthode echoDouble }
  Result := Value;
end;

initialization
  { Les classes invocables doivent être recensées }
  InvRegistry.RegisterInvokableClass(TIISLSESign);

end.
 
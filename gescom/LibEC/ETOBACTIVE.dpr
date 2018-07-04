library ETOBACTIVE;



uses
  ComServ,
  ETOBACTIVE_TLB in 'ETOBACTIVE_TLB.pas',
  uTOBActive in 'uTOBActive.pas' {TOBActive: CoClass},
  UtilTarif in '..\Lib\UtilTarif.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.

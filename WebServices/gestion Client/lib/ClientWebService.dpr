program ClientWebService;

uses
  Forms,
  Uficheprinc in 'Uficheprinc.pas' {FPrinc},
  HManifest in 'HManifest.pas',
  WSDL in 'WSDL.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFPrinc, FPrinc);
  Application.Run;
end.

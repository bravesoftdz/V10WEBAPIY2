program WSLSEDesign;

uses
  Forms,
  HManifest in 'HManifest.pas',
  UPrincipale in 'UPrincipale.pas' {Fprincipale},
  UgestionINI in '..\..\Commun\Lib\UgestionINI.pas',
  UNewCollab in 'UNewCollab.pas' {fNewID},
  USQLServer in 'USQLServer.pas',
  ftest in 'ftest.pas' {Form1},
  Udefinitions in '..\..\Commun\Lib\Udefinitions.pas',
  UCryptage in '..\..\Commun\Lib\UCryptage.pas',
  Uapplications in '..\..\..\lanceur LSE\lib\Uapplications.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFprincipale, Fprincipale);
  Application.Run;
end.

program LSEClientMaj;

uses
  Forms,
  UClientMajP in 'UClientMajP.pas' {FlanceurMajLse},
  UControleUAC in 'UControleUAC.pas',
  UEncryptage in 'UEncryptage.pas',
  Uregistry in 'Uregistry.pas',
  Uconstantes in 'Uconstantes.pas',
  USystemInfo in 'USystemInfo.pas',
  RunElevatedSupport in 'RunElevatedSupport.pas',
  UfronctionBase in '..\..\gescom\LibCreateRepo\UfronctionBase.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas',
  ZPatience in 'ZPatience.pas' {FPatience};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Mise à jour des applications LSE-CEGID';
  Application.CreateForm(TFlanceurMajLse, FlanceurMajLse);
  Application.Run;
end.

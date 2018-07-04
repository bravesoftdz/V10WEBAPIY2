program LseMajInfo;

uses
  Forms,
  UMajInfoBR in 'UMajInfoBR.pas' {FmajInfoBR},
  UControleUAC in 'UControleUAC.pas',
  UEncryptage in 'UEncryptage.pas',
  UbrowseRepert in 'UbrowseRepert.pas',
  UFolders in 'UFolders.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Mise à jour des informations pour update';
  Application.CreateForm(TFmajInfoBR, FmajInfoBR);
  Application.Run;
end.

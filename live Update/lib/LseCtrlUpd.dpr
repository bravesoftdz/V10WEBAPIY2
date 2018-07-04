program LseCtrlUpd;

uses
  Forms,
  UEncryptage in 'UEncryptage.pas',
  UControleUAC in 'UControleUAC.pas',
  UverifServeur in 'UverifServeur.pas',
  UFolders in 'UFolders.pas',
  UXmlUpdInfo in 'UXmlUpdInfo.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.ShowMainForm := false;
  Application.Run;
  SetInfoMaj;
end.

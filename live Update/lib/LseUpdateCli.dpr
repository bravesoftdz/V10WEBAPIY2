program LseUpdateCli;

uses
  Forms,
  UmajCli in 'UmajCli.pas' {FmajCli},
  UXmlUpdInfo in 'UXmlUpdInfo.pas',
  Uregistry in 'Uregistry.pas',
  UfileCtrl in 'UfileCtrl.pas',
  UControleUAC in 'UControleUAC.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'mise à jour du poste client';
  Application.Run;
  lanceMajCli;
end.

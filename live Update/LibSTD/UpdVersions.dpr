program UpdVersions;

uses
  Forms,
  UMainUpdate in 'UMainUpdate.pas' {fMainUpdate},
  UfronctionBase in '..\..\gescom\LibCreateRepo\UfronctionBase.pas',
  Uregistry in '..\lib\Uregistry.pas',
  UControleUAC in '..\lib\UControleUAC.pas',
  Uconstantes in '..\lib\Uconstantes.pas',
  UEncryptage in '..\lib\UEncryptage.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas',
  UCommonFuncs in '..\lib\UCommonFuncs.pas',
  USystemInfo in '..\lib\USystemInfo.pas',
  UlanceurUpdate in '..\lib\UlanceurUpdate.pas' {Form3},
  ZPatience in '..\lib\ZPatience.pas' {FPatience},
  CoolTrayIcon in '..\lib\CoolTrayIcon.pas',
  SimpleTimer in '..\lib\SimpleTimer.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm3, Form3);
  Application.Run;
end.

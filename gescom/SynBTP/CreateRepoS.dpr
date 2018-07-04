program CreateRepoS;

{$APPTYPE CONSOLE}

uses
  Windows,
  SysUtils,
  Registry,
  SHellApi,
  UEncryptage in '..\..\live Update\lib\UEncryptage.pas',
  UfronctionBase in '..\LibCreateRepo\UfronctionBase.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas',
  Uconstantes in '..\..\live Update\lib\Uconstantes.pas',
  UControleUAC in '..\..\live Update\lib\UControleUAC.pas',
  Uregistry in '..\..\live Update\lib\Uregistry.pas',
  ZPatience in '..\..\live Update\lib\ZPatience.pas' {FPatience},
  UCommonFuncs in '..\..\live Update\lib\UCommonFuncs.pas',
  USystemInfo in '..\..\live Update\lib\USystemInfo.pas';

begin
  { TODO -oUser -cConsole Main : placez le code ici }
  ConstitueRepository ;
end.

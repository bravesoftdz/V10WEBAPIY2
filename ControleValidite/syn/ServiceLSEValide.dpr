program ServiceLSEValide;

uses
  SvcMgr,
  UServiceControleValide in '..\lib\Service\UServiceControleValide.pas' {LSEAutorise: TService},
  UConnectionData in '..\lib\Service\UConnectionData.pas',
  IdHash in '..\..\..\..\Borland\Delphi7\Source\Indy\IdHash.pas',
  Ulog in '..\lib\Commun\Ulog.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TLSEAutorise, LSEAutorise);
  Application.Run;
end.

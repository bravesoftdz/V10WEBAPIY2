program ServiceLSEValide;

uses
  SvcMgr,
  UServiceControleValide in '..\lib\UServiceControleValide.pas' {LSEAutorise: TService};

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TLSEAutorise, LSEAutorise);
  Application.Run;
end.

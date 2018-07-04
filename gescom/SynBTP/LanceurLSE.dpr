program LanceurLSE;

uses
  Forms,
  HManifest in '..\..\lanceur LSE\lib\HManifest.pas',
  FichePrinc in '..\..\lanceur LSE\lib\FichePrinc.pas' {FFichePrinc},
  USplashLanceur in '..\..\lanceur LSE\lib\USplashLanceur.pas' {Fsplash},
  Uapplications in '..\..\lanceur LSE\lib\Uapplications.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Lanceur L.S.E';
  Application.CreateForm(TFFichePrinc, FFichePrinc);
  Application.Run;
end.

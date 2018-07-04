program activate;

uses
  Forms,
  fcalcul in '..\lib\fcalcul.pas' {calcul},
  CLKBtpLib_TLB in '..\lib\CLKBtpLib_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Calcul de clef d''activation';
  Application.CreateForm(Tcalcul, calcul);
  Application.Run;
end.

program LSEProtect;

uses
  Forms,
  HManifest in '..\..\lanceur LSE\lib\HManifest.pas',
  Cryptage in '..\..\gescom\LibBTP\Cryptage.pas',
  UAcceuil in '..\Lib\UAcceuil.pas' {Fprincipale},
  UCalcProtect in '..\Lib\UCalcProtect.pas' {FCalcProtec},
  UdroitUtilisation in '..\..\gescom\LibBTP\UdroitUtilisation.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Protection L.S.E';
  Application.CreateForm(TFprincipale, Fprincipale);
  Application.Run;
end.

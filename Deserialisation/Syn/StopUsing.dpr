program StopUsing;

uses
  Forms,
  UUtilSeria in '..\Lib\UUtilSeria.pas',
  UtilEnvoieMail in '..\Lib\UtilEnvoieMail.pas',
  UADeseria in '..\Lib\UADeseria.pas' {TFDeseria},
  WinUAC in '..\Lib\WinUAC.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Arret d''utilisation des produits BUSINESS';
  Application.CreateForm(TTFDeseria, TFDeseria);
  Application.Run;
end.

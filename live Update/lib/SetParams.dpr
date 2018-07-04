program SetParams;

{$R 'UAC\ManifestAdmin.res' 'UAC\ManifestAdmin.rc'}

uses
  Forms,
  USetParams in 'USetParams.pas' {Form1},
  UEncryptage in 'UEncryptage.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'L.S.E Update Params';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

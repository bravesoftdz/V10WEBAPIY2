program ExportParams;

{$R 'UAC\ManifestAdmin.res' 'UAC\ManifestAdmin.rc'}

uses
  Forms,
  UExportParams in 'UExportParams.pas' {Form2},
  Uregistry in 'Uregistry.pas',
  Uconstantes in 'Uconstantes.pas',
  Ulog in '..\..\commun\Lib\Ulog.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'L.S.E Export Update Params';
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.

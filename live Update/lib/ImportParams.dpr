program ImportParams;

{$R 'UAC\ManifestAdmin.res' 'UAC\ManifestAdmin.rc'}

uses
  Forms,
  UimportParams in 'UimportParams.pas' {FimportParams},
  Uconstantes in 'Uconstantes.pas',
  Uregistry in 'Uregistry.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFimportParams, FimportParams);
  Application.Run;
end.

program ExportDSL;

uses
  Forms,
  UExportCsv in '..\..\ExportCSV\UExportCsv.pas' {Form1},
  Ulog in '..\..\commun\Lib\Ulog.pas',
  IdFTP in '..\..\ExportCSV\IdFTP.pas',
  IdAntiFreeze in '..\..\Indy10\Core\IdAntiFreeze.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Export DSL';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

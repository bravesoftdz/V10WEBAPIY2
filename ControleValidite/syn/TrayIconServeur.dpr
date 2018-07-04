program TrayIconServeur;

uses
  Forms,
  UParamControleValide in '..\lib\UParamControleValide.pas' {Form1},
  UConnectionData in '..\lib\UConnectionData.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

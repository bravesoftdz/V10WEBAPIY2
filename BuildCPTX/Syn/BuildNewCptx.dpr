program BuildNewCptx;

uses
  Forms,
  Umain in '..\lib\Umain.pas' {Form1},
  galPatience in '..\..\Cegidpgi\Lib\galPatience.pas' {FPatience},
  UControleUAC in '..\..\live Update\lib\UControleUAC.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'Outil de génération des Build CPTX';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.

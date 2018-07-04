unit UimportParams;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Buttons, StdCtrls,FileCtrl,Uregistry,UConstantes;


type
  TFimportParams = class(TForm)
    Label1: TLabel;
    Repository: TEdit;
    BFindEmplacement: TSpeedButton;
    Panel1: TPanel;
    BtCache: TSpeedButton;
    Bvalide: TSpeedButton;
    procedure BFindEmplacementClick(Sender: TObject);
    procedure BtCacheClick(Sender: TObject);
    procedure BvalideClick(Sender: TObject);
  private
    { Déclarations privées }
    procedure ImportParams;
    function readToken(var TheChaine: string): string;
  public
    { Déclarations publiques }
  end;

var
  FimportParams: TFimportParams;

implementation

{$R *.dfm}

procedure TFimportParams.BFindEmplacementClick(Sender: TObject);
var TheRepert : string;
begin
  TheRepert := Repository.text;
  if SelectDirectory ('Sélectionner le répertoire de stockage',Repository.text,Therepert) then Repository.text := TheRepert;
end;

procedure TFimportParams.BtCacheClick(Sender: TObject);
begin
  close;
end;

procedure TFimportParams.BvalideClick(Sender: TObject);
begin
  if Repository.Text = '' then
  begin
    Application.MessageBox('Merci de renseigner le répertoire de stockage','Export des paramètres',MB_OKCANCEL or MB_ICONEXCLAMATION);
    exit;
  end;
  ImportParams;
  close;
end;

function TFimportParams.readToken (var TheChaine : string) : string;
var II : integer;
begin
  II :=  POs('|',TheChaine);
  if II> 0 then
  begin
    result:= copy(TheChaine,1,II-1);
    TheChaine := copy( TheChaine,II+1,Length(TheCHaine));
  end else
  begin
    result := TheChaine;
    TheChaine := '';
  end;
end;

procedure TFimportParams.ImportParams;
var DIrectory,NbMins,OutilsVisible : string;
    InfoFile : string;
    f : TextFile;
begin

  AssignFile(f, IncludeTrailingBackslash(Repository.text)+NAMEPARAM);
  reset(f);
  readln ( f, InfoFile);
  CloseFile(f);

  DIrecTory := readToken(infofile);
  if length(infofile)>0 then NbMins := readToken(infofile);
  if length(infofile)>0 then OutilsVisible := readToken(infofile);;
  SetInfoStocke ('OutilsVisible',OutilsVisible);

  SetInfoStocke ('ServerRepository',Directory);
  SetInfoStocke ('Timer',NbMins);
end;

end.

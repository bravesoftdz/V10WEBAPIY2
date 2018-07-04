unit USetParams;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, ExtCtrls, StdCtrls, ComCtrls,FileCtrl,UEncryptage;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    BtCache: TSpeedButton;
    Bvalide: TSpeedButton;
    Label1: TLabel;
    Repository: TEdit;
    BFindEmplacement: TSpeedButton;
    Label2: TLabel;
    Label3: TLabel;
    TB: TTrackBar;
    NBmins: TLabel;
    USER: TEdit;
    Label4: TLabel;
    PASSW: TEdit;
    Label5: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure BFindEmplacementClick(Sender: TObject);
    procedure BtCacheClick(Sender: TObject);
    procedure BvalideClick(Sender: TObject);
    procedure TBChange(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  Form1: TForm1;

implementation
uses Uregistry;

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
var ThePass : string;
    RR : string;
begin
	Repository.text:= GetInfoStocke ('ServerRepository');
  RR := getInfoStocke('Timer');
  if RR <> '' then TB.Position := StrToInt(RR) else TB.position := 60; 
end;

procedure TForm1.BFindEmplacementClick(Sender: TObject);
var TheRepert : string;
begin
  TheRepert := Repository.text;
  if SelectDirectory ('Séléctionner un dépot',Repository.text,Therepert) then Repository.text := TheRepert;
end;

procedure TForm1.BtCacheClick(Sender: TObject);
begin
  close;
end;

procedure TForm1.BvalideClick(Sender: TObject);
begin
  if (Repository.text = '') or (TB.Position=00) then
  begin
    Application.MessageBox('Merci de renseigner les informations demandées','Mise à jour des versions',MB_OKCANCEL or MB_ICONEXCLAMATION);
    exit;
  end;
  SetInfoStocke('ServerRepository',Repository.text);
  SetInfoStocke('Timer',InttoStr(TB.Position));
  close;
end;

procedure TForm1.TBChange(Sender: TObject);
begin
  NBmins.Caption := InttoStr(TB.Position);  
end;

end.

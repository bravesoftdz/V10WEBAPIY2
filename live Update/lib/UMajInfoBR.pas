unit UMajInfoBR;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ImgList, XPMan, Menus, CoolTrayIcon, Mask, StdCtrls,
  Spin, ComCtrls, Buttons,UControleUAC, UEncryptage, UbrowseRepert,Uregistry;

type
  TFmajInfoBR = class(TForm)
    Panel1: TPanel;
    BtCache: TSpeedButton;
    Bvalide: TSpeedButton;
    TB1: TTabControl;
    PC1: TPageControl;
    TabParam: TTabSheet;
    Label2: TLabel;
    VerifAuto: TCheckBox;
    Intervalle: TSpinEdit;
    TabSheet1: TTabSheet;
    Label3: TLabel;
    GBmethode: TGroupBox;
    Label1: TLabel;
    BFindEmplacement: TSpeedButton;
    ServerUpd: TRadioButton;
    LocalUpd: TRadioButton;
    RepertStockage: TEdit;
    User: TEdit;
    FindRepert: TOpenDialog;
    XPManifest1: TXPManifest;
    Label5: TLabel;
    UpdServer: TEdit;
    FindServer: TSpeedButton;
    OpenServer: TOpenDialog;
    procedure BFindEmplacementClick(Sender: TObject);
    procedure BvalideClick(Sender: TObject);
    procedure BtCacheClick(Sender: TObject);
    procedure LocalUpdClick(Sender: TObject);
    procedure ServerUpdClick(Sender: TObject);
    procedure VerifAutoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FindServerClick(Sender: TObject);
  private
    procedure EnregInfo;
    procedure GetInfos;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FmajInfoBR: TFmajInfoBR;

implementation

{$R *.dfm}

procedure TFmajInfoBR.BFindEmplacementClick(Sender: TObject);
var LeRepert : string;
begin
	lerepert := getrepert (self);
  if lerepert <> '' then
  begin
    RepertStockage.Text := LeRepert;
  end;
end;

procedure TFmajInfoBR.BvalideClick(Sender: TObject);
begin
	EnregInfo;
  close;
end;

procedure TFmajInfoBR.EnregInfo;
begin
	SetInfoStocke ('ServerUpd',UpdServer.Text);
	SetInfoStocke ('LocalFolder',RepertStockage.Text);
  if VerifAuto.Checked then SetInfoStocke('AutoScan','1') else SetInfoStocke('AutoScan','0');
  SetInfoStocke('Timer',intToStr(Intervalle.Value));
  SetInfoStocke ('user',User.text);
// 	SetInfoStocke ('Password',crypte(Password.Text));
  if ServerUpd.Checked then SetInfoStocke('UpdMethod','0') else SetInfoStocke('UpdMethod','1');
end;

procedure TFmajInfoBR.BtCacheClick(Sender: TObject);
begin
  close;
end;

procedure TFmajInfoBR.GetInfos;
var fTimer : string;
begin
	UpdServer.Text := GetInfoStocke ('ServerUpd');
	RepertStockage.Text := GetInfoStocke ('LocalFolder');
  VerifAuto.Checked := (getInfoStocke('AutoScan')='1');
  fTimer := getInfoStocke('Timer');
  if fTimer <> '' then Intervalle.Value := StrToInt(getInfoStocke('Timer'))
  							  else Intervalle.Value := 5;
  User.Text := GetInfoStocke ('user');
//  Password.Text := decrypte(GetInfoStocke ('Password'));
  if getInfoStocke('UpdMethod')='0' then ServerUpd.Checked := true else LocalUpd.Checked := true;
end;

procedure TFmajInfoBR.VerifAutoClick(Sender: TObject);
begin
	Intervalle.Enabled := VerifAuto.Checked;
end;

procedure TFmajInfoBR.ServerUpdClick(Sender: TObject);
begin
  RepertStockage.enabled := false;
end;

procedure TFmajInfoBR.LocalUpdClick(Sender: TObject);
begin
  RepertStockage.enabled := true;
end;

procedure TFmajInfoBR.FormCreate(Sender: TObject);
begin
  GetInfos;
end;

procedure TFmajInfoBR.FindServerClick(Sender: TObject);
begin
//
end;

end.

unit ZEch;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, hmsgbox, StdCtrls, Buttons, Mask, Hctrls,
  HEnt1,
  ParamDat; { Mac3 }

function SaisieZEch(var Mode : string; var Ech : TDateTime; Action : TActionFiche ; DateEcr : tDateTime) : boolean ;

type
  TZFEch = class(TForm)
    FMODE: THValComboBox;
    HLabel1: THLabel;
    HLabel2: THLabel;
    FEch: TMaskEdit;
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    Baide: THBitBtn;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FEchKeyPress(Sender: TObject; var Key: Char);
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    DateEcr : tDateTime ;
  public
    bModeRO : Boolean ;
  end;

implementation

{$R *.DFM}

Uses SaisUtil ;

function SaisieZEch(var Mode : string; var Ech : TDateTime; Action : TActionFiche ; DateEcr : tDateTime) : boolean ;
var ZFEch : TZFEch ;
begin
ZFEch:=TZFEch.Create(Application) ;
  try
  ZFEch.FMode.Value:=Mode ;
  ZFEch.FEch.Text:=DateToStr(Ech) ;
  ZFEch.bModeRO:=(Action=taConsult) ;
  ZFEch.DateEcr:=DateEcr ;
  Result:=(ZFEch.ShowModal=mrOK) ;
  if Result then begin Mode:=ZFEch.FMode.Value ; Ech:=StrToDate(ZFEch.FEch.Text) ; end ;
  finally
  ZFEch.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TZFEch.BValiderClick(Sender: TObject);
var Ech: TDateTime ;
begin
try
  StrToDate(FEch.Text) ;
except
  HM.Execute(2,'','') ; Exit ;
end ;
If DateEcr<>0 Then
  BEGIN
  Ech:=StrToDate(FEch.Text) ;
  if Not NbJoursOK(DateEcr,Ech) then BEGIN HM.Execute(2,'','') ; Exit ; END ;
  END ;
(*
if Not IsValidDate(FDateEche.Text) then FDateEche.Text:=DateToStr(EcheInit) ;
if Not IsValidDate(FDateValeur.Text) then FDateValeur.Text:=DateToStr(ValInit) ;
*)
if FMode.Value='' then begin HM.Execute(0,'','') ; Exit ; end ;
ModalResult:=mrOk ;
end;

procedure TZFEch.FormShow(Sender: TObject);
begin
if bModeRO then begin FMODE.Enabled:=FALSE ; FEch.Enabled:=FALSE ; end ;
end;

procedure TZFEch.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
try
  StrToDate(FEch.Text) ;
except
  HM.Execute(2,'','') ; CanClose:=FALSE ;
end ;
if FMode.Value='' then begin HM.Execute(0,'','') ; CanClose:=FALSE ; end ;
end;

procedure TZFEch.FEchKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end ;

procedure TZFEch.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Vide : Boolean ;
begin
Vide:=(Shift=[]) ;
case Key of
     VK_F10 : if (Vide) then begin Key:=0 ; BValiderClick(nil) ;  end ;
//  VK_ESCAPE : if (Vide) then begin Key:=0 ; BAbandonClick(nil) ; end ;
  end ;
end;

end.

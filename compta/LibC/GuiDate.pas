{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/06/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit GuiDate ;

interface

uses
  SysUtils,
  Windows,
  Classes,
  Controls,
  HEnt1,
  Forms,
  StdCtrls,
  Mask,
  Hctrls,
  Buttons,
  hmsgbox,
  HSysMenu;

Const MethodeGuideDate1 : Boolean = TRUE ;
Procedure SaisirDateGuide ( Var DD : TDateTime ) ;

type
  TFGuiDate = class(TForm)
    HLabel2: THLabel;
    FDate: TMaskEdit;
    BValider: THBitBtn;
    BAnnuler: THBitBtn;
    Baide: THBitBtn;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BaideClick(Sender: TObject);
  private
  public
    DD : tDateTime ;
  end;

implementation

{$R *.DFM}

Uses SaisUtil ;

Procedure SaisirDateGuide ( Var DD : TDateTime ) ;
var FGuiDate : TFGuiDate ;
    OkOk : Boolean ;
BEGIN
FGuiDate:=TFGuiDate.Create(Application) ;
try
  FGuiDate.DD:=DD ;
  OkOk:=(FGuiDate.ShowModal=mrOK) ;
  if OkOk then
     BEGIN
     DD:=StrToDate(FGuiDate.FDate.Text) ;
     END ;
  finally
  FGuiDate.Free ;
  end;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFGuiDate.FormShow(Sender: TObject);
begin
FDate.Text:=DateToStr(DD) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
BorderIcons:=[] ; BorderStyle:=bsNone ; Caption:='' ;
(*
BValider.Visible:=False ; BAnnuler.Visible:=False ; BAide.Visible:=False ;
BValider.Enabled:=False ; BAnnuler.Enabled:=False ; BAide.Enabled:=False ;
*)
end;

procedure TFGuiDate.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if Shift<>[] then Exit ;
if Key=VK_F10 then BEGIN Key:=0 ; BValiderClick(Nil) ; END else
 if Key=VK_ESCAPE then BEGIN Key:=0 ; ModalResult:=mrCancel ; END ;
end;

procedure TFGuiDate.BValiderClick(Sender: TObject);
Var Err : Integer ;
begin
Err:=ControleDate(FDate.Text) ;
if Err>0 then BEGIN HM.Execute(Err,caption,'') ; FDate.SetFocus ; Exit ; END ;
ModalResult:=mrOk ;
end;

procedure TFGuiDate.BAnnulerClick(Sender: TObject);
begin
HM.Execute(0,'','') ;
end;

procedure TFGuiDate.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if  (ModalResult<>mrOk) then BEGIN HM.Execute(0,'','') ; CanClose:=False ; END ;
end;

procedure TFGuiDate.BaideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.

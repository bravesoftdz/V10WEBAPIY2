unit MsgQrTVA;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DB, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  hmsgbox, StdCtrls, Hctrls, Buttons, ExtCtrls,
  ENT1, HEnt1, HSysMenu, Mask ;

Function MsgTva(Lequel : Byte ; Deductible : Boolean ; D1,D2 : TDateTime ; RegimeTva,CodeTva : String) : Boolean ;

type
  TFMsgTVA = class(TForm)
    HPB: TPanel;
    BValider: THBitBtn;
    BFerme: THBitBtn;
    GBFermeDEF: TGroupBox;
    HLabel1: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HMess: THMsgBox;
    HLabel6: THLabel;
    HMTrad: THSystemMenu;
    GroupBox1: TGroupBox;
    HLabel14: THLabel;
    Label1: TLabel;
    Fdate1: TMaskEdit;
    FDate2: TMaskEdit;
    HLabel2: THLabel;
    TypeTVA: TRadioGroup;
    HLabel7: THLabel;
    FCodeTva: TEdit;
    FRegimeTva: TEdit;
    procedure BValiderClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    Lequel : Byte ;
    OkOk,Deductible : Boolean ;
    D1,D2 : TDateTime ;
    RegimeTva,CodeTva : String ;
  public
  end;

implementation

{$R *.DFM}


Function MsgTva(Lequel : Byte ; Deductible : Boolean ; D1,D2 : TDateTime ; RegimeTva,CodeTva : String) : Boolean ;
var FMsgTVA: TFMsgTVA;
    OkMess : Boolean ;
begin
Result:=FALSE ;
FMsgTVA:=TFMsgTVA.Create(Application) ; OkMess:=FALSE ;
 try
   FMsgTVA.Lequel:=Lequel ;
   FMsgTVA.Deductible:=Deductible ;
   FMsgTVA.D1:=D1 ;
   FMsgTVA.D2:=D2 ;
   FMsgTVA.RegimeTva:=RegimeTva ;
   FMsgTVA.CodeTva:=CodeTva ;
   FMsgTVA.OkOk:=OkMess ;
   FMsgTVA.ShowModal ;
 Finally
   OkMess:=FMsgTVA.OkOk ;
   FMsgTVA.free ;
 End ;
If OkMess Then Result:=TRUE ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFMsgTVA.BValiderClick(Sender: TObject);
var i : Integer ;
begin
i:=HMess.Execute(Lequel,'','') ;
If i<>mrYes Then Exit ; Screen.Cursor:=SynCrDefault ;
OkOk:=TRUE ;
end;

procedure TFMsgTVA.FormShow(Sender: TObject);
begin
HLabel1.Caption:=HMess.Mess[Lequel+3] ;
FDate1.Text:=DateToStr(D1) ; FDate2.Text:=DateToStr(D2) ;
If RegimeTva='' Then RegimeTva:=Traduirememoire('<<Tous>>') ;
FCodeTva.Text:=RechDom('TTTVA',CodeTva,FALSE) ; FRegimeTva.Text:=RegimeTva ;
If Deductible Then TypeTva.ItemIndex:=1 Else TypeTva.ItemIndex:=0 ;
end;

end.

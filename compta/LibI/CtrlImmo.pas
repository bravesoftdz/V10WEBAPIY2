unit CtrlImmo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ImEnt, HEnt1, Hctrls, hmsgbox, HSysMenu, HPanel, UiUtil,
  HTB97 ;

type
  TFCtrlImmo = class(TForm)
    Panel1: TPanel;
    FControleMvt: TGroupBox;
    FAmort: TCheckBox;
    FOp: TCheckBox;
    FcontroleCpt: TGroupBox;
    FImmos: TCheckBox;
    FCptes: TCheckBox;
    MsgRien: THMsgBox;
    MsgBar: THMsgBox;
    HMTrad: THSystemMenu;
    FEche: TCheckBox;
    Mes: THMsgBox;
    TTravail: TLabel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    procedure MajCaption(M : byte) ;
    procedure LanceControl ;
    function Selection : boolean ;
  public
    { Déclarations publiques }
  end;

procedure ControleFichiersImmo ;

implementation

uses VerImmo, UtilPgi ;

{$R *.DFM}

procedure ControleFichiersImmo ;
var CFic : TFCtrlImmo ;
    PP : THPanel ;
begin
if not _BlocageMonoPoste(False) then Exit ;
CFic:=TFCtrlImmo.Create(Application) ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     CFic.ShowModal ;
   finally
     CFic.Free ;
     _DeblocageMonoPoste(False) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(CFic,PP) ;
  CFic.Show ;
  end ;
end ;

procedure TFCtrlImmo.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TTravail.Caption:='' ;
FImmos.Checked:=True ; FCptes.Checked:=True ;
FAmort.Checked:=True ; FOp.Checked:=True ; FEche.Checked:=True ;
//if V^.SAV then BEGIN FNumPL.Visible:=True ; FNumPL.Checked:=True ; END
//          Else BEGIN FNumPL.Visible:=False ; FNumPL.Checked:=False ; END
end;

procedure TFCtrlImmo.BValiderClick(Sender: TObject);
begin
  if not Selection then begin MsgRien.Execute(0,'','') ; Exit ; end ;
  if MsgRien.Execute(2,'','')<>mryes then exit ;
  TTravail.Caption:=MsgBar.Mess[0] ;
  LanceControl ;
  TTravail.Caption:='' ;
end;

procedure TFCtrlImmo.LanceControl ;
begin
  EnableControls(Self,False) ; TTravail.Enabled:=True ;
  if FImmos.Checked then ExecuteVerifImmos(tvImmos, FALSE) ;
  if FCptes.Checked then ExecuteVerifImmos(tvCptes, FALSE) ;
  if FAmort.Checked then ExecuteVerifImmos(tvAmort, FALSE) ;
  if FOp.Checked    then ExecuteVerifImmos(tvOp,    FALSE) ;
  if FEche.Checked  then ExecuteVerifImmos(tvEch,   FALSE) ;
  EnableControls(Self,True) ;
  MsgRien.Execute(1,'','') ;
end ;

function TFCtrlImmo.Selection : Boolean ;
begin
Result:=(FImmos.Checked or FCptes.Checked or FAmort.Checked or FOp.Checked
         or FEche.Checked) ;
end ;

procedure TFCtrlImmo.MajCaption( M : byte ) ;
begin
TTravail.Caption:=MsgBar.Mess[0]+' '+MsgBar.Mess[M] ;
end ;

procedure TFCtrlImmo.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCtrlImmo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   begin
   _DeblocageMonoPoste(False) ;
   Action:=caFree ;
   end ;
end;

end.

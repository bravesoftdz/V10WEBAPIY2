unit CtrlFic;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, Ent1, HEnt1, Hctrls, hmsgbox, HSysMenu, HPanel, UiUtil,
  HTB97 ;

procedure ControlFic ;


type
  TFCtrlFic = class(TForm)
    Panel1: TPanel;
    FControleMvt: TGroupBox;
    FComptable: TCheckBox;
    FLettrage: TCheckBox;
    FcontroleCpt: TGroupBox;
    FGen: TCheckBox;
    FAux: TCheckBox;
    FSec: TCheckBox;
    FJal: TCheckBox;
    MsgRien: THMsgBox;
    MsgBar: THMsgBox;
    HMTrad: THSystemMenu;
    FNumPL: TCheckBox;
    Mes: THMsgBox;
    TTravail: TLabel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
  private
    Cpt, Cpta, Piece : Boolean ;
    Procedure MajCaption( M : byte ) ;
    Procedure LanceControl ;
    Procedure AvecQui ;
    Function Selection : Boolean ;
  public
  end;

implementation

{$R *.DFM}

uses VerCpta, VerCpte, VerPiece, UtilPgi ;

procedure ControlFic ;
var CFic : TFCtrlFic ;
    PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(False,'',TRUE) then Exit ;
CFic:=TFCtrlFic.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
    CFic.ShowModal ;
   finally
    CFic.Free ;
    _DeblocageMonoPoste(False,'',TRUE) ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  END else
  BEGIN
  InitInside(CFic,PP) ;
  CFic.Show ;
  END ;
END ;

procedure TFCtrlFic.FormCreate(Sender: TObject);
begin
{$IFDEF CCS3}
  FSec.Visible := not EstComptaSansAna;
{$ENDIF}
end;

procedure TFCtrlFic.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TTravail.Caption:='';
FGen.Checked:=True;
FAux.Checked:=True;
FSec.Checked:=(True and FSec.Visible);
FJal.Checked:=True;
FComptable.Checked:=True;
FLettrage.Checked:=True;
FNumPL.Checked:=True ;
FNumPL.Visible := V_PGI.SAV;
FNumPL.Checked := V_PGI.SAV;
end;

procedure TFCtrlFic.BValiderClick(Sender: TObject);
begin
if Not Selection then begin MsgRien.Execute(0,'','') ; Exit ; end ;
if MsgRien.Execute(2,'','')<>mryes then exit ;
TTravail.Caption:=MsgBar.Mess[0] ;
AvecQui ;
LanceControl ;
TTravail.Caption:='' ;
end;

Procedure TFCtrlFic.AvecQui ;
BEGIN
Cpt:=(FGen.Checked or FAux.Checked or FSec.Checked or FJal.Checked) ;
Cpta:=(FComptable.Checked or FLettrage.Checked) ;
Piece:=FNumPL.Checked ;
END ;

Procedure TFCtrlFic.LanceControl ;
Var i : Byte ; NbNoPossible : integer ;
    AuMoinsUn : Boolean ;
BEGIN
EnableControls(Self,False) ; TTravail.Enabled:=True ;
(*if Cpt then
   BEGIN
   if FGen.Checked then begin MajCaption(1) ; VerCompte(1) ; end ;
   if FAux.Checked then begin MajCaption(2) ; VerCompte(2) ; end ;
   if FSec.Checked then begin MajCaption(3) ; VerCompte(3) ; end ;
   if FJal.Checked then begin MajCaption(4) ; VerCompte(4) ;end ;
   END ;
*)
AuMoinsUn:=(FGen.Checked or FAux.Checked or FSec.Checked or FJal.Checked) ;
if Cpt then
   BEGIN
   if AuMoinsUn and V_PGI.SAV then VerifCompte Else
      BEGIN
      if FGen.Checked then
         BEGIN
         NbNoPossible:=0 ; MajCaption(1) ;
         VerCompteMAJ(1,NbNoPossible, True) ;
         if NbNoPossible<>0 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',Mes.Mess[0]) ;
         END ;
      if FAux.Checked then
         BEGIN
         NbNoPossible:=0 ; MajCaption(2) ;
         VerCompteMAJ(2,NbNoPossible, True) ;
         if NbNoPossible<>0 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',Mes.Mess[1]) ;
         END ;
      if FSec.Checked then
         BEGIN
         NbNoPossible:=0 ; MajCaption(3) ;
         VerCompteMAJ(3,NbNoPossible, True) ;
         if NbNoPossible<>0 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',Mes.Mess[2]) ;
         END ;
      if FJal.Checked then
         BEGIN
         NbNoPossible:=0 ; MajCaption(4) ;
         VerCompteMAJ(4,NbNoPossible, True) ;
         if NbNoPossible<>0 then MsgRien.Execute(3,IntToStr(NbNoPossible)+' ',Mes.Mess[3]) ;
         END ;
      END ;
   END ;
if Cpta then
   BEGIN
   I:=100 ;
   if (FComptable.Checked and FLettrage.Checked) then
      BEGIN
      i:=0 ; MajCaption(7) ;
      END Else
   if  FComptable.Checked then
       BEGIN
       i:=1 ; MajCaption(5) ;
       END Else
   if FLettrage.Checked then
      BEGIN
      i:=2 ; MajCaption(6) ;
      END ;
   If i<>100 Then ControleMvt(i) ;
   END ;
if Piece then begin MajCaption(8) ; VerifPiece ; end ;
EnableControls(Self,True) ;
MsgRien.Execute(1,'','') ;
END ;

Function TFCtrlFic.Selection : Boolean ;
BEGIN
Result:=(FGen.Checked or FAux.Checked or FSec.Checked or FJal.Checked
         or FComptable.Checked or FLettrage.Checked or FNumPL.Checked) ;
END ;

Procedure TFCtrlFic.MajCaption( M : byte ) ;
BEGIN
TTravail.Caption:=MsgBar.Mess[0]+' '+MsgBar.Mess[M] ;
END ;

procedure TFCtrlFic.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFCtrlFic.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(False) ;
   Action:=caFree ;
   END ;
end;

procedure TFCtrlFic.BFermeClick(Sender: TObject);
begin
  //SG6 09.03.05
  Close;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

end.

unit AnnulRS ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1, HEnt1,
  HCtrls, hmsgbox, HSysMenu, StdCtrls, ExtCtrls, HTB97, HPanel, UiUtil, TImpFic ;

Procedure AnnuleRecupSISCO ;

type
  TFAnnulRecup = class(TForm)
    Panel1: THPanel;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    CAnnul: TRadioGroup;
    HMTrad: THSystemMenu;
    HMess: THMsgBox;
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure AnnuleRecupSISCO ;
var FFR: TFAnnulRecup;
    PP       : THPanel ;
begin
FFR:=TFAnnulRecup.Create(Application) ;
PP:=FindInsidePanel ;
if PP=NIL then
   BEGIN
   try
    FFR.ShowModal ;
   finally
    FFR.Free ;
   end ;
   END else
   BEGIN
   InitInside(FFR,PP) ;
   FFR.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFAnnulRecup.BValiderClick(Sender: TObject);
begin
If HMess.Execute(0,Caption,'')<>mrYes then exit ;
EnableControls(Self,False) ;
Try
 BeginTrans ;
 If (CAnnul.ItemIndex=0) Or  (CAnnul.ItemIndex=2) Then
   BEGIN
   ExecuteSQL('DELETE FROM GENERAUX') ;
   ExecuteSQL('DELETE FROM TIERS') ;
   ExecuteSQL('DELETE FROM SECTION') ;
   ExecuteSQL('DELETE FROM JOURNAL') ;
   ExecuteSQL('DELETE FROM RIB') ;
   ExecuteSQL('DELETE FROM VENTIL') ;
   END ;
 If (CAnnul.ItemIndex=1) Or  (CAnnul.ItemIndex=2) Then
   BEGIN
   ExecuteSQL('DELETE FROM ECRITURE') ;
   ExecuteSQL('DELETE FROM ANALYTIQ') ;
   // CA - 30/11/2001
   ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART=1 WHERE SH_TYPE="CPT"') ;
   END ;
 CommitTrans ;
 HMess.Execute(1,Caption,'') ;
Except
 HMess.Execute(2,Caption,'') ;
 Rollback ;
End ;
EnableControls(Self,TRUE) ;
end;

procedure TFAnnulRecup.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFAnnulRecup.FormShow(Sender: TObject);
begin
If RecupSISCO Then Caption:=HMess.Mess[3] Else
  If RecupSERVANT Then Caption:=HMess.Mess[4] ;
UpdateCaption(Self) ;
end;

end.

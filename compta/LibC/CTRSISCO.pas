unit CTRSISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, HFLabel, HPanel, HSysMenu, Ent1, HEnt1, HCtrls, {$IFNDEF DBXPRESS}dbtables,
  hmsgbox{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  hmsgbox ;

Function ControleAvantRecupSISCO : Boolean ;

type
  TFCtrlRecup = class(TForm)
    Panel1: THPanel;
    L1: TLabel;
    L2: TLabel;
    L3: TLabel;
    L4: TLabel;
    L5: TLabel;
    L6: TLabel;
    L7: TLabel;
    FlashCtrl: TFlashingLabel;
    EC1: TFlashingLabel;
    EC2: TFlashingLabel;
    EC3: TFlashingLabel;
    EC4: TFlashingLabel;
    EC5: TFlashingLabel;
    EC6: TFlashingLabel;
    EC7: TFlashingLabel;
    Ok1: TImage;
    OK2: TImage;
    OK3: TImage;
    OK4: TImage;
    OK5: TImage;
    OK6: TImage;
    OK7: TImage;
    Stop1: TImage;
    Stop2: TImage;
    Stop3: TImage;
    Stop4: TImage;
    Stop5: TImage;
    Stop6: TImage;
    Stop7: TImage;
    HMTrad: THSystemMenu;
    LeTimer: TTimer;
    HMess: THMsgBox;
    procedure LeTimerTimer(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    NbSec : Integer ;
    LanceControle,OkFerme : Boolean ;
    ControleOk: Boolean ;
    procedure Affiche(i : Integer ; OkOk : Boolean) ;
    Function Controles : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function ControleAvantRecupSISCO : Boolean ;
var FFR: TFCtrlRecup;
begin
Result:=TRUE ;
FFR:=TFCtrlRecup.Create(Application) ;
FFR.ControleOk:=TRUE ;
try
 FFR.ShowModal ;
finally
 Result:=FFR.ControleOk ;
 FFR.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFCtrlRecup.Affiche(i : Integer ; OkOk : Boolean) ;
Var EC : tControl ;
BEGIN
EC:=TControl(FindComponent('EC'+IntToStr(i))) ;
If EC<>NIL Then EC.Visible:=FALSE ;
If i<7 Then
  BEGIN
  EC:=TControl(FindComponent('EC'+IntToStr(i+1))) ;
  If EC<>NIL Then EC.Visible:=TRUE ;
  END ;
If OkOk Then EC:=TControl(FindComponent('OK'+IntToStr(i)))
        Else EC:=TControl(FindComponent('Stop'+IntToStr(i))) ;
If EC<>NIL Then Ec.Visible:=TRUE ;
Application.ProcessMessages ;
Delay(1000) ;
END ;

Function CtrlFourchettes : Boolean ;
Var Q : TQuery ;
    OkCha,OkPro,OkBan : Boolean ;
    i : Integer ;
BEGIN
OkCha:=FALSE ; OkPro:=FALSE ; OkBan:=FALSE ;
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="SIS"',TRUE) ;
While Not Q.Eof Do
  BEGIN
  If Copy(Q.FindField('CC_CODE').AsString,1,2)='BA' Then OkBan:=TRUE ;
  Q.Next ;
  END ;
Ferme(Q) ;
For i:=1 To 5 Do If (VH^.FCha[i].Deb<>'') And (VH^.FCha[i].Fin<>'') Then OkCha:=TRUE ;
For i:=1 To 5 Do If (VH^.FPro[i].Deb<>'') And (VH^.FPro[i].Fin<>'') Then OkPro:=TRUE ;
Result:=OkCha And OkPro And OkBan ;
END ;

Function CtrlParam : Boolean ;
BEGIN
Result:=(VH^.Cpta[fbGene].Lg<>0) And (VH^.Cpta[fbGene].CB<>'') And
        (VH^.Cpta[fbAux].Lg<>0) And (VH^.Cpta[fbAux].CB<>'') ;                           
END ;

Function CtrlEtab : Boolean ;
Var Q : TQuery ;
    OkEtab : Boolean ;
BEGIN
OkEtab:=TRUE ;
Q:=OpenSQL('SELECT * FROM ETABLISS',TRUE) ;
If Not Q.Eof Then
  BEGIN
  If (VH^.EtablisDefaut='') Then OkEtab:=FALSE ;
  END Else OkEtab:=FALSE ;
Ferme(Q) ;
Result:=OkEtab ;
END ;

Function CtrlDevise : Boolean ;
Var Q : TQuery ;
    OkFongible : Boolean ;
    Err : Integer ;
BEGIN
Result:=TRUE ; Err:=0 ; OkFongible:=FALSE ;
Q:=OpenSQL('SELECT * FROM DEVISE',TRUE) ;
While Not Q.Eof Do
  BEGIN
  If (Q.FindField('D_FONGIBLE').AsString='X') Then
    BEGIN
    OkFongible:=TRUE ;
    If (Q.FindField('D_MONNAIEIN').AsString<>'X') Then Err:=1 ;
    END ;
  If (Q.FindField('D_MONNAIEIN').AsString='X') And (Arrondi(Q.FindField('D_PARITEEURO').AsFloat,5)=0) Then Err:=2 ;
  Q.Next ;
  END ;
Ferme(Q) ;
//If Not ParamEuroOk Then Err:=3 ;
If Not OkFongible Then Err:=4 ; 
If Err<>0 Then Result:=FALSE ;
END ;

Function CtrlModePaie : Boolean ;
Var Q : TQuery ;
    OkMP : Boolean ;
BEGIN
OkMP:=TRUE ;
Q:=OpenSQL('SELECT * FROM MODEPAIE',TRUE) ;
If Not Q.Eof Then
  BEGIN
  END Else OkMP:=FALSE ;
Ferme(Q) ;
Result:=OkMP ;
END ;

Function CtrlModeRegle : Boolean ;
Var Q : TQuery ;
    OkMR : Boolean ;
BEGIN
OkMR:=TRUE ;
Q:=OpenSQL('SELECT * FROM MODEREGL',TRUE) ;
If Not Q.Eof Then
  BEGIN
  END Else OkMR:=FALSE ;
Ferme(Q) ;
Result:=OkMR ;
END ;

Function CtrlRegime : Boolean ;
Var Q : TQuery ;
    OkRegime : Boolean ;
BEGIN
OkRegime:=TRUE ;
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="RTV"',TRUE) ;
If Not Q.Eof Then
  BEGIN
  END Else OkRegime:=FALSE ;
Ferme(Q) ;
Result:=OkRegime ;
END ;

Function TFCtrlRecup.Controles : Boolean ;
Var OkOk : Boolean ;
BEGIN
InitMove(7,'') ; EC1.Visible:=TRUE ; Application.ProcessMessages ; Result:=TRUE ;
OkOk:=CtrlFourchettes ; MoveCur(FALSE) ; Affiche(1,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlParam       ; MoveCur(FALSE) ; Affiche(2,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlEtab        ; MoveCur(FALSE) ; Affiche(3,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlDevise      ; MoveCur(FALSE) ; Affiche(4,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlModePaie    ; MoveCur(FALSE) ; Affiche(5,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlModeRegle   ; MoveCur(FALSE) ; Affiche(6,OkOk) ; If Not OkOk Then Result:=FALSE ;
OkOk:=CtrlRegime      ; MoveCur(FALSE) ; Affiche(7,OkOk) ; If Not OkOk Then Result:=FALSE ;
FiniMove ;
END ;

procedure TFCtrlRecup.LeTimerTimer(Sender: TObject);
Var OkOk : Boolean ;
begin
If Not LanceControle Then
  BEGIN
  LanceControle:=TRUE ; OkOk:=Controles ; FlashCtrl.Flashing:=FALSE ;
  If Not OkOk Then BEGIN FlashCtrl.Caption:=HMess.Mess[0] ; ControleOk:=FALSE ; END
              Else BEGIN FlashCtrl.Caption:=HMess.Mess[1] ; ControleOk:=TRUE ; END ;
  Application.ProcessMessages  ;
  OkFerme:=TRUE ;
  END Else If OkFerme Then
  BEGIN
  Inc(NbSec) ; If NbSec>3 Then Close ;
  END ;
end;

procedure TFCtrlRecup.FormShow(Sender: TObject);
begin
LanceControle:=FALSE ; NbSec:=0 ; OkFerme:=FALSE ;
LeTimer.Enabled:=TRUE ;
end;

end.

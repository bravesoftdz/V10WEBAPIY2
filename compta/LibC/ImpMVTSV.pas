unit ImpMVTSV;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls, {$IFNDEF DBXPRESS}dbtables,
  HTB97, StdCtrls, HPanel{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  HTB97, StdCtrls, ImpFicU, SISCO, TImpFic, HPanel, UiUtil, LicUtil, Paramsoc ;

Function ImportMvtSERVANT : Boolean ;

type
  TFRecupMvtSERVANT = class(TForm)
    Panel1: THPanel;
    HMTrad: THSystemMenu;
    HMess: THMsgBox;
    HLabel2: THLabel;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    Sauve: TSaveDialog;
    Dock: TDock97;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    CBTypeRecupMvtSISCO: TRadioGroup;
    CBTriRecupSISCO: TRadioGroup;
    CBShuntTransfert: TCheckBox;
    procedure RechFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBShuntTransfertClick(Sender: TObject);
    procedure CBTypeRecupMvtSISCOClick(Sender: TObject);
  private
    { Déclarations privées }
    ShunteTransfert : Boolean ;
    Function OkRecupCpt : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

uses FmtChoix, AssistRS, MulJalSI, moteurIP, AssImp ;

{$R *.DFM}

Const LgLs=11 ;

Function ImportMvtSERVANT : Boolean ;
var FFR: TFRecupMvtSERVANT;
    PP       : THPanel ;
begin
Result:=TRUE ;
FFR:=TFRecupMvtSERVANT.Create(Application) ;
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

procedure TFRecupMvtSERVANT.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFRecupMvtSERVANT.FormShow(Sender: TObject);
begin
ShunteTransfert:=FALSE ;
If ParamCount=1 Then ShunteTransfert:=(ParamStr(1)='ROBERT LE COCHON') Or (CryptageSt(ParamStr(1))=CryptageSt(DayPass(Date))) ;
If ShunteTransfert Then CBShuntTransfert.Visible:=TRUE Else CBShuntTransfert.Checked:=FALSE ;
end;

Function TFRecupMvtSERVANT.OkRecupCpt : Boolean ;
Var Q,Q1 : TQuery ;
    i : Integer ;
    Sect,Cpt : String ;
BEGIN
Result:=FALSE ; i:=0 ;
If HMess.Execute(3,'','')<>mrYes Then Exit ;
if not FileExists(FileName.Text) then BEGIN HMess.Execute(2,'','') ; Exit ; END ;
{$IFDEF SPEC302}
Q:=Opensql('SELECT SO_GENATTEND FROM SOCIETE',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Q1:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+Q.Fields[0].AsString+'" ',TRUE) ;
  If Q1.Eof Then i:=5 ;
  Ferme(Q1) ;
  END ;
Ferme(Q) ;
{$ELSE}
Q1:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+GetParamSoc('SO_GENATTEND')+'"',TRUE) ;
If Q1.Eof Then i:=5 ;
Ferme(Q1) ;
{$ENDIF}
Q:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_VENTILABLE="X"',TRUE) ;
If Not Q.Eof Then
  BEGIN
  Sect:='' ;
  Q1:=OpenSQL('SELECT X_SECTIONATTENTE FROM AXE WHERE X_AXE="A1"',TRUE) ;
  If Not Q1.Eof Then Sect:=Q1.Fields[0].AsString ;
  Ferme(Q1) ;
  If Sect<>'' Then
    BEGIN
    Q1:=OpenSQL('Select S_SECTION FROM SECTION WHERE S_SECTION="'+Sect+'" ',TRUE) ;
    If Q1.Eof Then i:=6 ;
    Ferme(Q1) ;
    END ;
  END ;
Ferme(Q) ;
If Not ParamEuroOk Then i:=7 ;
Q:=Opensql('SELECT J_COMPTEURNORMAL FROM JOURNAL WHERE J_COMPTEURNORMAL=""',TRUE) ;
If Not Q.Eof Then i:=8 ;
Ferme(Q) ;
If i<>0 Then BEGIN HMess.Execute(i,'','') ; Exit ; END ;
MajNomFicSISCO(FileName.Text) ;
Result:=TRUE ;
END ;

procedure TFRecupMvtSERVANT.BValiderClick(Sender: TObject);
Var Env : TEnvImpAuto ;
    Pb : Boolean ;
    Err : Integer ;
    PasDeTransfert : Boolean ;
    NomFic : String ;
begin
If Not OkRecupCpt Then Exit ;
PasDeTransfert:=ShunteTransfert And CBShuntTransfert.Visible And CBShuntTransfert.Checked ;
NomFic:=FileName.Text ;
//Fillchar(TRS,SizeOf(TRS),#0) ;
//PrepareTRS(TRS,FileName.Text) ;
Env.Format:='CGE' ; Env.Lequel:='FEC' ; Env.NomFic:=NomFic ; Env.CodeFormat:='' ;
Case CBTypeRecupMvtSISCO.ItemIndex Of
  0 : MoteurImportAuto(Env,NIL) ;
  1 : LanceImportRecupSISCO('FEC',NomFic) ;
  END ;
end;

procedure TFRecupMvtSERVANT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFRecupMvtSERVANT.CBShuntTransfertClick(Sender: TObject);
begin
//CBTriRecupSISCO.Enabled:=Not CBShuntTransfert.Checked ;
CBTypeRecupMvtSISCO.Enabled:=Not CBShuntTransfert.Checked ;
end;

procedure TFRecupMvtSERVANT.CBTypeRecupMvtSISCOClick(Sender: TObject);
begin
If CBTypeRecupMvtSISCO.ItemIndex=0 Then
  BEGIN
  CBShuntTransfert.Enabled:=FALSE ;
  CBShuntTransfert.Checked:=FALSE ;
  END Else CBShuntTransfert.Enabled:=TRUE ;
end;

end.

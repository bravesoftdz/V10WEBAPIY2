unit ImpMVTSI;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls,  HStatus,
  HTB97, StdCtrls, ImpFicU, SISCO, TImpFic, HPanel, UiUtil, LicUtil, Paramsoc, udbxDataset ;

Function ImportMvtSISCO : Boolean ;

type
  TFRecupMvtSISCO = class(TForm)
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
    CBLett: TCheckBox;
    CBBigVol: TCheckBox;
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
    procedure GereLettSISCO ;
  public
    { Déclarations publiques }
  end;

implementation

uses FmtChoix, AssistRS, MulJalSI, moteurIP, AssImp, LETTREF ;

{$R *.DFM}

Const LgLs=11 ;

Function ImportMvtSISCO : Boolean ;
var FFR: TFRecupMvtSISCO;
    PP       : THPanel ;
begin
Result:=TRUE ;
FFR:=TFRecupMvtSISCO.Create(Application) ;
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

procedure TFRecupMvtSISCO.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFRecupMvtSISCO.GereLettSISCO ;
BEGIN
CBLett.Visible:=FALSE ; CBLett.Checked:=FALSE ;
If VH^.RecupLTL Or VH^.RecupSISCOPGI Then
  BEGIN
  CBLett.Visible:=TRUE ;
  CBBigVol.Visible:=TRUE ;
  END ;
END ;

procedure TFRecupMvtSISCO.FormShow(Sender: TObject);
Var Q : Tquery ;
begin
ShunteTransfert:=FALSE ;
If ParamCount=1 Then ShunteTransfert:=(ParamStr(1)='ROBERT LE COCHON') Or (CryptageSt(ParamStr(1))=CryptageSt(DayPass(Date))) ;
If ShunteTransfert Then CBShuntTransfert.Visible:=TRUE ;
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="SIS" AND CC_CODE="NOM"',TRUE) ;
If Not Q.Eof Then FileName.Text:=Q.FindField('CC_LIBRE').ASString ;
Ferme(Q) ;
SourisNormale ;
GereLettSISCO ;
end;

Function TFRecupMvtSISCO.OkRecupCpt : Boolean ;
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

Procedure PrepareTRS(var TRS : TRecupSISCO ; StFichier : String) ;
BEGIN
TRS.NomFicOrigine:=StFichier ;
TRS.NomFicCpt:=NewNomFicEtDir(StFichier,'Cpt','SISCO') ;
TRS.NomFicMvt:=NewNomFicEtDir(StFichier,'Mvt','SISCO') ;
TRS.NomFicCptCGE:=NewNomFic(TRS.NomFicCpt,'CGE') ;
TRS.NomFicMvtCGE:=NewNomFic(TRS.NomFicMvt,'CGE') ;
TRS.RuptEc:=OnFolio ;
END ;

Procedure AlimListe(St : String ; Var TRS : TRecupSISCO) ;
Var StJ : String ;
    St1 : String ;
    StOr : String ;
// Liste triée : Jour(2) OrigineSaisie(1) N°(8)+
BEGIN
If St[1]='E' Then
  BEGIN
  Inc(TRS.LigSISCO) ;
  StJ:=Copy(St,2,2) ; StOr:=Trim(Copy(St,139,1)) ; If StOr='' Then StOr:='F' ;
  St:=StJ+StOr+FormatFloat('00000000',TRS.Ind)+St+'&[+'+FormatFloat('00',Trs.LigSISCO) ;
  END Else
  BEGIN
  St1:=TRS.LM[TRS.LM.Count-1] ;
  TRS.LM.Delete(TRS.LM.Count-1) ;
  St:=St1+'###'+St ;
  END ;
TRS.LM.Add(St) ; Inc(TRS.Ind) ;
END ;

Procedure GereRupture(Var NEwFichier : TextFile ; St : String ; Var TRS : TRecupSISCO) ;
Var IsRuptureFolio,IsRuptureJal,IsRupturePeriode,IsRuptureAutre,IsRupture,OkRupt : Boolean ;
    StLM,St1,St2 : String ;
    ll,ll1,Lg : Integer ;
    i,j : Integer ;
    RuptATraiter : Boolean ;
BEGIN
RuptATraiter:=(TRS.RuptEc<>OnQued) And (TRS.LM.Count>0) ;
If RuptATraiter Then
  BEGIN
  IsRuptureFolio:=(St[1]='F') ; IsRupturePeriode:=(St[1]='M') ;
  IsRuptureJal:=(St[1]='J') ; IsRuptureAutre:=(St[1]<>'E') And (St[1]<>'v') And (St[1]<>'F') And (St[1]<>'J') ;
  IsRupture:=IsRuptureFolio Or IsRupturePeriode Or IsRuptureJal Or IsRuptureAutre ;
  OkRupt:=(IsRupture And (TRS.RuptEc=OnFolio)) Or (IsRupture And (TRS.RuptEc=OnJal) And (Not IsRuptureFolio)) ;
  If OkRupt Then
    BEGIN
    TRS.LM.Sorted:=TRUE ;
    For i:=0 To TRS.LM.Count-1 Do
      BEGIN
      St1:='' ; St2:='' ;
      St1:=Copy(TRS.LM[i],LgLs+1,Length(TRS.LM[i])-LgLs) ;
      StLM:=St1 ;
      ll:=Pos('###',TRS.LM[i]) ;
      If ll>0 Then
        BEGIN
        St1:=Copy(TRS.LM[i],LgLs+1,ll-1-LgLs) ;
        If St1<>'' Then Writeln(NewFichier,St1) ;
        Lg:=Length(St1) ;
        While Pos('###',StLM)>0 Do
          BEGIN
          ll:=Pos('###',StLM) ; For j:=ll To ll+2 Do StLM[j]:='/' ;
          ll1:=Pos('###',StLM) ;
          If ll1>0 Then St2:=Copy(StLM,ll+3,ll1-ll-3)
                   Else St2:=Copy(StLM,ll+3,Length(StLM)-Lg-3) ;
          If St2<>'' Then Writeln(NewFichier,St2) ;
          Lg:=Lg+Length(St2) ;
          END ;
        END Else If St1<>'' Then Writeln(NewFichier,St1) ;
      END ;
    TRS.LM.Clear ;
    TRS.LM.Sorted:=FALSE ;
    TRS.Ind:=0 ; TRS.LIGSISCO:=0 ;
    END ;
  END ;
If (St[1]<>'E') And (St[1]<>'v') Then Writeln(NewFichier,St) ;
END ;

Procedure PrepareFichierPhase1(Var TRS : TRecupSISCO) ;
Var Fichier,NewFichier : TextFile ;
    St,StNewFichier : String ;
// M J F E v
BEGIN
If (TRS.RuptEc=OnQued) Then Exit ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,TRS.NomFicMvt) ; Reset(Fichier) ;
StNewFichier:=FileTemp('.SII') ;
StNewFichier:=NewNomFic(TRS.NomFicMvt,'Tst') ;
TRS.LM:=TStringList.Create ;
TRS.LM.Sorted:=FALSE ;
TRS.LM.Duplicates:=DupIgnore ;
TRS.Ind:=0 ; TRS.LigSISCO:=0 ;
AssignFile(NewFichier,StNewFichier) ; Rewrite(NewFichier) ;
While (Not EOF(Fichier))  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If Trim(St)='' Then
    BEGIN
    St:='E' ;
    END Else
    BEGIN
    If TRS.RuptEc=OnQued Then Writeln(NewFichier,St) Else
      BEGIN
      If ((St[1]='E') Or (St[1]='v')) Then AlimListe(St,TRS) Else GereRupture(NewFichier,St,TRS) ;
      END ;
    END ;
  END ;
FiniMove ; Flush(NewFichier) ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
{$i-}
AssignFile(Fichier,TRS.NomFicMvt) ; Erase(Fichier) ;
{$i+}
renamefile(StNewFichier,TRS.NomFicMvt) ;
FichierOnDisk(TRS.NomFicMvt,TRUE) ;
FiniMove ;
TRS.LM.Clear ; TRS.LM.Free ;
END ;

Function NatureEnr(St : String) : Integer ;
Var Code1,Code2 : String ;
(*
Cpt 03 05 06 07 08 09 10 C 30 50 P L D I f s R S
Mvt 00 01 02 03 04 11 M J F E 12 p t r A B H v w G
*)

BEGIN
Result:=2 ;
Code1:=Copy(St,1,1) ; Code2:=Copy(St,1,2) ;
If (Code1='C') Or (Code1='P') Or (Code1='L') Or (Code1='D') Or (Code1='I') Or
   (Code1='f') Or (Code1='s') Or (Code1='R') Or (Code1='S') Or
   (Code2='05') Or (Code2='06') Or (Code2='07') Or (Code2='08') Or (Code2='09') Or
   (Code2='10') Or (Code2='30') Or (Code2='50') Then Result:=0 ;
If (Code1='M') Or (Code1='J') Or (Code1='F') Or (Code1='E') Or (Code1='p') Or
   (Code1='t') Or (Code1='r') Or (Code1='A') Or (Code1='B') Or
   (Code1='H') Or (Code1='v') Or (Code1='w') Or (Code1='G') Or
   (Code2='00') Or (Code2='01') Or (Code2='02') Or (Code2='04') Or
   (Code2='11') Or (Code2='12') Then Result:=1 ;

END ;

Procedure InitFichierPhase1(Var TRS : TRecupSISCO) ;
Var Fichier,NewFichier1,NewFichier2 : TextFile ;
    St : String ;
    Pb : Boolean ;
    What : Integer ;
BEGIN
Pb:=FALSE ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,TRS.NomFicOrigine) ; Reset(Fichier) ;
AssignFile(NewFichier1,TRS.NomFicCpt) ; Rewrite(NewFichier1) ;
AssignFile(NewFichier2,TRS.NomFicMvt) ; Rewrite(NewFichier2) ;
TRS.NbLCpt:=0 ; TRS.NbLMvt:=0 ;
While (Not EOF(Fichier)) And (Not Pb)  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  What:=NatureEnr(St) ;
  Case What Of
    0 : BEGIN Writeln(NewFichier1,St) ; Inc(TRS.NbLCpt) ; END ; // Cpt
    1 : BEGIN Writeln(NewFichier2,St) ; Inc(TRS.NbLMvt) ; END ; // Mvt
    2 : BEGIN Writeln(NewFichier1,St) ; Writeln(NewFichier2,St) ; Inc(TRS.NbLCpt) ; Inc(TRS.NbLMvt) ; END ;// Les 2
    END ;
  END ;
FiniMove ; Flush(NewFichier1) ; Flush(NewFichier2) ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
FichierOnDisk(TRS.NomFicCpt,TRUE) ; FichierOnDisk(TRS.NomFicMvt,TRUE) ;
FiniMove ;
END ;


procedure TFRecupMvtSISCO.BValiderClick(Sender: TObject);
Var TRS : TRecupSISCO ;
    Env : TEnvImpAuto ;
    Pb : Boolean ;
    Err,k : Integer ;
    PasDeTransfert : Boolean ;
begin
//If CBLett.Checked Then BEGIN LettrageParCode(TRUE) ; Exit ; END ;
If Not OkRecupCpt Then Exit ;
PasDeTransfert:=ShunteTransfert And CBShuntTransfert.Visible And CBShuntTransfert.Checked ;
Fillchar(TRS,SizeOf(TRS),#0) ;
PrepareTRS(TRS,FileName.Text) ; pb:=FALSE ;
If Not PasDeTransfert Then
  BEGIN
  Case CBTriRecupSISCO.ItemIndex Of
    0 : TRS.RuptEc:=OnQued ; 1 : TRS.RuptEc:=OnJal ; 2 : TRS.RuptEc:=OnFolio ;
    END ;
  InitFichierPhase1(TRS) ;
  PrepareFichierPhase1(TRS) ;
  Err:=TransfertSisco(TRS.NomFicMvt,FALSE,k,'',NIL,CbBigVol.Checked) ;
  Pb:=FALSE ;
  Case Err Of
    1,2 : BEGIN HMess.Execute(4,'','') ; Pb:=TRUE ; END ;
    END ;
  END ;
If Not Pb Then
  BEGIN
  Env.Format:='CGE' ; Env.Lequel:='FEC' ; Env.NomFic:=TRS.NomFicMvtCGE ; Env.CodeFormat:='' ;
  Case CBTypeRecupMvtSISCO.ItemIndex Of
    0 : MoteurImportAuto(Env,NIL) ;
    1 : LanceImportRecupSISCO('FEC',TRS.NomFicOrigine,CbLett.Checked,CbBigVol.Checked) ;
    END ;
  END ;
end;

procedure TFRecupMvtSISCO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFRecupMvtSISCO.CBShuntTransfertClick(Sender: TObject);
begin
CBTriRecupSISCO.Enabled:=Not CBShuntTransfert.Checked ;
CBTypeRecupMvtSISCO.Enabled:=Not CBShuntTransfert.Checked ;
end;

procedure TFRecupMvtSISCO.CBTypeRecupMvtSISCOClick(Sender: TObject);
begin
If (Not VH^.RecupLTL) And (Not VH^.RecupSISCOPGI) Then Exit ;
CbLett.Visible:=CBTypeRecupMvtSISCO.ItemIndex=1 ;
If Not CbLett.Visible Then CbLett.Checked:=FALSE ;
end;

end.

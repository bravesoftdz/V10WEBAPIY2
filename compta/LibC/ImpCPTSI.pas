unit ImpCPTSI;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls, {$IFNDEF DBXPRESS}dbtables,
  Menus, StdCtrls, Buttons, HTB97, HPanel{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  HTB97, StdCtrls, ImpFicU, SISCO, TImpFic, Buttons, Menus, HPanel, UiUtil ;

Function ImportCompteSISCO : Boolean ;
Procedure MAJCptBqe ;

type
  TFRecupCptSISCO = class(TForm)
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
    BMenuZoom: TToolbarButton97;
    POPZ: TPopupMenu;
    VisuInfo: TBitBtn;
    VisuJalFic: TBitBtn;
    VisuJalS7: TBitBtn;
    procedure RechFileClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BMenuZoomMouseEnter(Sender: TObject);
    procedure VisuInfoClick(Sender: TObject);
    procedure VisuJalFicClick(Sender: TObject);
    procedure VisuJalS7Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    Function OkRecupCpt : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

uses FmtChoix, AssistRS, MulJalSI, moteurIP, FouSISCO, JalSISCO, InfFicSI ;

{$R *.DFM}

Function ImportCompteSISCO : Boolean ;
var FFR: TFRecupCptSISCO;
    PP       : THPanel ;
begin
Result:=TRUE ;
FFR:=TFRecupCptSISCO.Create(Application) ;
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

procedure TFRecupCptSISCO.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFRecupCptSISCO.FormShow(Sender: TObject);
Var Q : Tquery ;
begin
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="SIS" AND CC_CODE="NOM"',TRUE) ;
If Not Q.Eof Then FileName.Text:=Q.FindField('CC_LIBRE').ASString ;
Ferme(Q) ;
SourisNormale ;
end;

Function TFRecupCptSISCO.OkRecupCpt : Boolean ;
BEGIN
Result:=FALSE ;
If HMess.Execute(3,'','')<>mrYes Then Exit ;
if not FileExists(FileName.Text) then BEGIN HMess.Execute(2,'','') ; Exit ; END ;
MajNomFicSISCO(FileName.Text) ;
Result:=TRUE ;
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

Procedure PrepareTRS(var TRS : TRecupSISCO ; StFichier : String) ;
BEGIN
ChargeFourchetteCompte('SIS',TRS.InfoSISCO.FourchetteImport) ;
TRS.NomFicOrigine:=StFichier ;
TRS.NomFicCpt:=NewNomFicEtDir(StFichier,'Cpt','SISCO') ;
TRS.NomFicMvt:=NewNomFicEtDir(StFichier,'Mvt','SISCO') ;
TRS.NomFicCptCGE:=NewNomFic(TRS.NomFicCpt,'CGE') ;
TRS.NomFicMvtCGE:=NewNomFic(TRS.NomFicMvt,'CGE') ;
chargeCptSISCO(TRS.InfoSISCO.FourchetteSISCO) ;
TRS.InfoSISCO.ShuntAnaCharge:=Not OkRecupAnaSISCO('ANC') ;
TRS.InfoSISCO.ShuntAnaProduit:=Not OkRecupAnaSISCO('ANP') ;
ChargeCharRemplace(TRS.InfoSISCO) ;
ChargeDiversDefaut(TRS.InfoSISCO) ;
END ;

Procedure PrepareFichierPhase1(Var TRS : TRecupSISCO) ;
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

Procedure RecupCpt(Var TRS : TRecupSISCO) ;
Var FichierCpt,FichierCptCGE : TextFile ;
    St,St1,St2 : String ;
    Env : TEnvImpAuto ;
BEGIN
AssignFile(FichierCpt,TRS.NomFicCpt) ; Reset(FichierCpt) ;
AssignFile(FichierCptCGE,TRS.NomFicCptCGE) ; Rewrite(FichierCptCGE) ;
Writeln(FichierCptCGE,TRS.NomFicCptCGE) ;
InitMove(TRS.NbLCpt,'') ;
While Not Eof(FichierCpt) Do
  BEGIN
  MoveCur(FALSE) ; ReadLn(FichierCpt,St) ;
  St1:='' ;
  If Copy(St,1,2)='03' Then RecupSociete(3,St,TRS.InfoSISCO) Else
   If St[1]='C' Then RecupCptSISCO(St,St1,St2,TRS.InfoSISCO) Else
    If St[1]='S' Then RecupSectionSISCO(St,St1,TRS.InfoSISCO) ;
  If St1<>'' Then Writeln(FichierCptCGE,St1) ;
  END ;
FiniMove ;
Close(FichierCpt) ; Close(FichierCptCGE) ;
Env.Format:='CGE' ; Env.Lequel:='FEC' ; Env.NomFic:=TRS.NomFicCptCGE ; Env.CodeFormat:='' ;
BeginTrans ;
MoteurImportAuto(Env,NIL) ;
CommitTrans ;
END ;

Procedure RecupJal(Var TRS : TRecupSISCO) ;
Var Fichier : TextFile ;
    St,Cpt : String ;
    Jal : String ;
    Q : TQuery ;
    Nat : String ;
    NatJalSISCO : String ;
BEGIN
Q:=PrepareSQL('Select * FROM JOURNAL WHERE J_JOURNAL=:CPT',FALSE) ;
AssignFile(Fichier,TRS.NomFicCpt) ; Reset(Fichier) ;
InitMove(TRS.NbLCpt,'') ;
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If Copy(St,1,2)='05' Then
    BEGIN
    Jal:=Trim(AnsiUpperCase(Copy(St,3,2))) ;
    Jal:=CptRemplace(Jal,TRS.InfoSISCO.CharRJ) ;
    Q.Close ; Q.Params[0].ASString:=Jal ; Q.Open ;
    If Q.Eof Then
      BEGIN
      Q.Insert ;
      InitNew(Q) ;
      Q.FindField('J_JOURNAL').AsString:=Jal ;
      Q.FindField('J_LIBELLE').AsString:=Trim(Copy(St,5,20)) ;
      Q.FindField('J_ABREGE').AsString:=Trim(Copy(St,5,17)) ;
      Q.FindField('J_NATUREJAL').AsString:='OD' ;
      Q.FindField('J_CREERPAR').AsString:='RCS' ;
      // J_MODESAISIE 
      NatJalSISCO:=Copy(St,36,1) ;
      If NatJalSISCO='T' Then
        BEGIN
        Cpt:=AnsiUpperCase(Trim(Copy(St,25,10))) ;
        Cpt:=BourreEtLess(Cpt,fbGene) ;
        Nat:=NatureCptImport(Cpt,TRS.InfoSISCO.FourchetteImport) ;
        If (Nat='BQE') Or (Nat='CAI') Then
          BEGIN
          Q.FindField('J_NATUREJAL').AsString:=Nat ;
          Q.FindField('J_CONTREPARTIE').AsString:=Cpt ;
          Q.FindField('J_TYPECONTREPARTIE').AsString:='MAN' ;
          END ;
        END Else If NatJalSISCO='X' Then Q.FindField('J_NATUREJAL').AsString:='EXT' Else
              If (NatJalSISCO='B') Or (NatJalSISCO='P') Then
              BEGIN Q.FindField('J_NATUREJAL').AsString:='ODA' ; Q.FindField('J_AXE').AsString:='A1' ; END ;
      Q.FindField('J_MODESAISIE').AsString:='-' ; 
      Q.Post ;
      END ;
    END ;
  END ;
FiniMove ;
Ferme(Q) ;
Close(Fichier) ;
Q:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CREERPAR="RCS"',TRUE) ;
If Not Q.Eof Then BEGIN Ferme(Q) ; MulticritereJournalSISCO(taModif) ; VerifJalSISCO(TRUE) ; END Else Ferme(Q) ;
END ;

Procedure MAJCptBqe ;
Var Q,Q1 : TQuery ;
    Gen : String ;
BEGIN
Q:=OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE="BQE"',TRUE) ;
InitMove(20,'') ;
While Not Q.Eof Do
  BEGIN
  MoveCur(FALSE) ;
  Gen:=Q.FindField('G_GENERAL').AsString ;
  Q1:=OpenSQL('Select * from BANQUECP where BQ_GENERAL="'+Gen+'" ',FALSE) ;
  If Q1.Eof Then
    BEGIN
    Q1.Insert ;
    InitNew(Q1) ;
    Q1.FindField('BQ_CODE').AsString:=Gen ;
    Q1.FindField('BQ_GENERAL').AsString:=Gen ;
    Q1.FindField('BQ_LIBELLE').AsString:='RIB d''attente' ;
    Q1.FindField('BQ_DEVISE').AsString:=V_PGI.DevisePivot ;
    Q1.Post ;
    END ;
  Ferme(Q1) ;
  Q.Next ;
  END ;
Ferme(Q) ;
FiniMove ;
END ;

procedure TFRecupCptSISCO.BValiderClick(Sender: TObject);
Var TRS : TRecupSISCO ;
begin
If Not OkRecupCpt Then Exit ;
EnableControls(Self,False) ;
Fillchar(TRS,SizeOf(TRS),#0) ;
PrepareTRS(TRS,FileName.Text) ;
PrepareFichierPhase1(TRS) ;
RecupCpt(TRS) ;
RecupJal(TRS) ;
MajCptBqe ;
EnableControls(Self,TRUE) ;
end;

procedure TFRecupCptSISCO.BMenuZoomMouseEnter(Sender: TObject);
begin
PopZoom97(BMenuZoom,POPZ) ;
end;

procedure TFRecupCptSISCO.VisuInfoClick(Sender: TObject);
begin
If Trim(FileName.Text)<>'' Then InfoFicSISCO(FileName.Text) ;
end;

procedure TFRecupCptSISCO.VisuJalFicClick(Sender: TObject);
begin
If Trim(FileName.Text)<>'' Then VisuJalSISCO(FileName.text) ;
end;

procedure TFRecupCptSISCO.VisuJalS7Click(Sender: TObject);
begin
MulticritereJournalSISCO(taModif) ; VerifJalSISCO(TRUE) ;
end;

procedure TFRecupCptSISCO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

end.

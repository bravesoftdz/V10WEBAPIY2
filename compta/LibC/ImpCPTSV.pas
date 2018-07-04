unit ImpCPTSV;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls, {$IFNDEF DBXPRESS}dbtables,
  HTB97, StdCtrls, HPanel{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  HTB97, StdCtrls, ImpFicU, SISCO, TImpFic, Buttons, Menus, HPanel, UiUtil ;

Function ImportCompteSERVANT(Qui : Integer) : Boolean ;
Procedure MAJCptBqe ;

type
  TFRecupCptSERVANT = class(TForm)
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
    procedure RechFileClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    Qui : Integer ;
    Function OkRecupCpt : Boolean ;
  public
    { Déclarations publiques }
  end;

implementation

uses FmtChoix, AssistRS, MulJalSI, ImpUtil ;

{$R *.DFM}

Function ImportCompteSERVANT(Qui : Integer) : Boolean ;
var FFR: TFRecupCptSERVANT;
    PP       : THPanel ;
begin
Result:=TRUE ;
FFR:=TFRecupCptSERVANT.Create(Application) ;
FFR.Qui:=Qui ;
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

procedure TFRecupCptSERVANT.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

Function TFRecupCptSERVANT.OkRecupCpt : Boolean ;
BEGIN
Result:=FALSE ;
If HMess.Execute(3,Caption,'')<>mrYes Then Exit ;
if not FileExists(FileName.Text) then BEGIN HMess.Execute(2,Caption,'') ; Exit ; END ;
Result:=TRUE ;
END ;

Procedure ImporteLesJournaux(NomFic : String) ;
Var Fichier : TextFile ;
    St,Cpt : String ;
    Jal : String ;
    Q : TQuery ;
    Nat : String ;
    NatJalSISCO : String ;
    NbImport : Integer ;
BEGIN
NbImport:=0 ;
AssignFile(Fichier,NomFic) ;
{$I-} Reset (Fichier) ; {$I+}
ReadLn(Fichier,St) ;
While not EOF(Fichier) do BEGIN NbImport:=NbImport+1 ; Readln(Fichier,St) ; END ;
CloseFile(Fichier) ;
Q:=PrepareSQL('Select * FROM JOURNAL WHERE J_JOURNAL=:CPT',FALSE) ;
AssignFile(Fichier,NomFic) ; Reset(Fichier) ;
InitMove(NbImport,'') ;
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  Jal:=Trim(AnsiUpperCase(Copy(St,4,3))) ;
  Q.Close ; Q.Params[0].ASString:=Jal ; Q.Open ;
  If Q.Eof Then
    BEGIN
    Q.Insert ;
    InitNew(Q) ;
    Q.FindField('J_JOURNAL').AsString:=Jal ;
    Q.FindField('J_LIBELLE').AsString:=Trim(Copy(St,7,35)) ;
    Q.FindField('J_ABREGE').AsString:=Trim(Copy(St,7,17)) ;
    Q.FindField('J_NATUREJAL').AsString:='OD' ;
    Q.FindField('J_CREERPAR').AsString:='RCS' ;
    Q.Post ;
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

Procedure AlimInfoImp(Var InfoImp :TInfoImport) ;
BEGIN
FillChar(InfoImp,SizeOf(InfoImp),#0) ; CreateListeImp(InfoImp) ;
InfoImp.Lequel:='FEC' ; InfoImp.Format:='CGE' ;
ChargeScenarioImport(InfoImp,FALSE) ;
END ;

Function ImporteUnCompte(St : String ; Var InfoImp : TInfoImport ; QFiche : TQFiche ; Var OkGen : Boolean) : Boolean ;
Var OkOk : Boolean ;
    ResImp : tResultImportCpte ;
    Val : String ;
BEGIN
Result:=FALSE ;
If EstUneLigneCptOk(St,'CGE') Then OkGen:=TRUE ;
OkOk:=EstUneLigneCptOk(St,'CGE') Or EstUneLigneCptOk(St,'CAU') Or EstUneLigneCptOk(St,'CAE') Or
      EstUneLigneCptOk(St,'SAN') Or EstUneLigneCptOk(St,'SSA') Or EstUneLigneCptOk(St,'TL') ;
(*
Case Lequel[1] Of
  'G' : OkOk:=EstUneLigneCptOk(St,'CGE') ;
  'X' : OkOk:=EstUneLigneCptOk(St,'CAU') Or EstUneLigneCptOk(St,'CAE');
  'A' : OkOk:=EstUneLigneCptOk(St,'SAN') Or EstUneLigneCptOk(St,'SSA');
  Else OkOk:=FALSE ;
  END ;
*)
If OkOk Then
  BEGIN
  Result:=TRUE ; Val:=Trim(Copy(St,7,17)) ;
  ResImp:=TraiteImportCompte(St,InfoImp,QFiche) ;
  (*
  Case ResImp Of
    resCreer    : BEGIN
                  Inc(NbImport) ;
                  MajCompteRendu(0,False,Val,Val,17) ;
                  END ;
    resModifier : BEGIN
                  Inc(NbModifies) ;
                  MajCompteRendu(1,False,Val,Val,17) ;
                  END ;
    END ;
  *)
  END ;
END ;

procedure ImporteLesComptes(Qui : Integer ; NomFic : String ; Var OkGen : Boolean) ;
Var TRS : TRecupSISCO ;
    InfoImp : PtTInfoImport ;
    QFiche : TQFiche ;
    i : Integer ;
    NbImport : Integer ;
    Fichier : TextFile ;
    St : String ;
begin
New(InfoImp) ; AlimInfoImp(InfoImp^) ; OkGen:=FALSE ;
For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
NbImport:=0 ;
AssignFile(Fichier,NomFic) ;
{$I-} Reset (Fichier) ; {$I+}
ReadLn(Fichier,St) ;
While not EOF(Fichier) do BEGIN NbImport:=NbImport+1 ; Readln(Fichier,St) ; END ;
CloseFile(Fichier) ;
InitMove(NbImport,'') ; MoveCur(False) ;
AssignFile(Fichier,NomFic) ;
{$I-} Reset (Fichier) ; {$I+}
Readln(Fichier,St) ;
SourisSablier ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(False) ; Readln(Fichier,St) ; St:=Ascii2Ansi(St) ;
  ImporteUnCompte(St,InfoImp^,QFiche,OkGen) ;
  END ;
CloseFile(Fichier) ;
VideListeInfoImp(InfoImp^,TRUE) ; Dispose(InfoImp) ;
For i:=0 To 3 Do Ferme(QFiche[i]) ;
If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
FiniMove ;
SourisNormale ;
end;


procedure TFRecupCptSERVANT.BValiderClick(Sender: TObject);
Var OkGen : Boolean ;
begin
If Not OkRecupCpt Then Exit ;
EnableControls(Self,False) ;
Try
 BeginTrans ;
 Case Qui Of
   0 : ImporteLesComptes(Qui,FileName.Text,OkGen) ;
   1 : ImporteLesJournaux(FileName.Text) ;
   End ;
 If (Qui=0) And OkGen Then MajCptBqe ;
 CommitTrans ;
Except
 RollBack ;
End ;
EnableControls(Self,TRUE) ;
end;

procedure TFRecupCptSERVANT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFRecupCptSERVANT.FormShow(Sender: TObject);
begin
Caption:=HMess.Mess[Qui+4] ; UpdateCaption(Self) ;
end;

end.

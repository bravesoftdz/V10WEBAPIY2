unit Transfic;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu, HTB97, StdCtrls, Hctrls, Buttons, ExtCtrls, FmtChoix, Ent1, Hent1, ImpFic,
  hmsgbox ;

Procedure TransformeFichier ;

type
  TFFic = class(TForm)
    Outils: TPanel;
    GroupBox1: TGroupBox;
    HLabel2: THLabel;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    Sauve: TSaveDialog;
    HSystemMenu1: THSystemMenu;
    Label1: TLabel;
    C1: TEdit;
    Label2: TLabel;
    C2: TEdit;
    Patience: TLabel;
    Label3: TLabel;
    Ref: TCheckBox;
    Montant: TCheckBox;
    NumP: TCheckBox;
    Lib: TCheckBox;
    MP: TCheckBox;
    TE: TCheckBox;
    BValider: TBitBtn;
    BFerme: TBitBtn;
    BAide: TBitBtn;
    TNBLC: TLabel;
    NbLC: TLabel;
    BVoir: TButton;
    Sens: TCheckBox;
    TC: TCheckBox;
    Label4: TLabel;
    ModifAll: TCheckBox;
    HMess: THMsgBox;
    procedure RechFileClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BVoirClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    Function OkMaj(i : Integer) : Boolean ;
  public
    { Déclarations publiques }
  end;

var
  FFic: TFFic;

implementation

{$R *.DFM}

Procedure TransformeFichier ;
Var X : TFFic ;
BEGIN
X:=TFFic.Create(Application) ;
 Try
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
END ;

Function TFFic.OkMaj(i : Integer) : Boolean ;
BEGIN
Result:=FALSE ;
If Ref.Checked And (i>=39) And (i<52) Then Result:=TRUE ;
If Lib.Checked And (i>=52) And (i<77) Then Result:=TRUE ;
If MP.Checked And (i=77) Then Result:=TRUE ;
If Montant.Checked And (i>=85) And (i<105) Then Result:=TRUE ;
If NumP.Checked And (i>=106) And (i<113) Then Result:=TRUE ;
If TE.Checked And (i=105) Then Result:=TRUE ;
If TC.Checked And (i=25) Then Result:=TRUE ;
If Sens.Checked And (i=84) Then Result:=TRUE ;
If ModifAll.Checked Then Result:=TRUE ;
END ;

procedure TFFic.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFFic.BValiderClick(Sender: TObject);
Var Fichier,NewFichier : TextFile ;
    St : String ;
    NewFileName : String ;
    Ca1,Ca2 : Char ;
    OkCorr : Boolean ;
    i,NbLCorr : Integer ;
begin
if not FileExists(FileName.Text) then BEGIN ShowMessage('Fichier introuvable') ; Exit ; END ;
If HMess.Execute(0,'','')<>mrYes Then Exit ;
AssignFile(Fichier,FileName.Text) ;
{$I-} Reset(Fichier) ; {$I+} if IoResult<>0 then Exit ;
Patience.Visible:=TRUE ; NBLC.Caption:='0' ;
Application.ProcessMessages ;
If C1.Text='' Then C1.Text:=' ' ; If C2.Text='' Then C2.Text:=' ' ;
Ca1:=C1.Text[1] ; Ca2:=C2.Text[1] ;
NewFileName:=FileTemp('.PNM') ;
AssignFile(NewFichier,NewFileName) ; Rewrite(NewFichier) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier,St) ; NbLCorr:=0 ;
While Not EOF(Fichier) do
   BEGIN
   Readln(Fichier,St) ; OkCorr:=FALSE ;
   For i:=1 To length(St) Do If (St[i]=Ca1) And (OkMaj(i)) Then BEGIN St[i]:=Ca2 ; OkCorr:=TRUE ; END ;
   WriteLn(NewFichier,St) ; If OkCorr Then Inc(NbLCorr) ;
   END ;
CloseFile(Fichier) ; CloseFile(NewFichier) ;
AssignFile(Fichier,FileName.Text) ; Erase(Fichier) ;
renamefile(NewFileName,FileName.Text) ;
NBLC.Caption:=IntToStr(NbLCorr) ;
Patience.Visible:=FALSE ; Application.ProcessMessages ;
ShowMessage('Traitement effectué') ;
Screen.Cursor:=SynCrDefault ;
end;

procedure TFFic.BVoirClick(Sender: TObject);
var StErr : String ;
    fmtFic : Integer ;
begin
StErr :='' ;
FmtFic:=0 ;
if FileExists(FileName.text) then VisuLignesErreurs(FileName.text,StErr,FmtFic,TRUE)
                             else ShowMessage('Fichier introuvable') ;
end;

procedure TFFic.FormShow(Sender: TObject);
begin
FileName.SetFocus ;
end;

end.

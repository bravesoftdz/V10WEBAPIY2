unit USISCO ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, Hctrls, ExtCtrls, ComCtrls, HStatus, HEnt1,HPanel, UiUtil ;

Procedure FaitFicSISCO ;

type
  TVisufichier = class(TForm)
    Panel1: TPanel;
    HLabel2: THLabel;
    FileName: TEdit;
    RechFile: TToolbarButton97;
    BVoir: TButton;
    Sauve: TSaveDialog;
    Status: THStatusBar;
    procedure FileNameChange(Sender: TObject);
    procedure RechFileClick(Sender: TObject);
    procedure BVoirClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

uses ImpFic, SISCO ;

{$R *.DFM}

procedure FaitFicSISCO ;
var X: TVisufichier;
    PP : THPanel ;
begin
X:=TVisufichier.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     X.ShowModal ;
    finally
     X.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end;


procedure TVisufichier.FileNameChange(Sender: TObject);
begin
BVoir.Enabled:=FileName.Text<>'' ;
end;

procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
var j,i : integer ;
BEGIN
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then BEGIN j:=i ; Break ; END ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
END ;

procedure TVisufichier.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TVisufichier.BVoirClick(Sender: TObject);
Var k : Integer ;
begin
if FileExists(FileName.text) then TransfertSISCO(FileName.text,FALSE,k)
                             else ShowMessage('Fichier introuvable') ;
end;

end.

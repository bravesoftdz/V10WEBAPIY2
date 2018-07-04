unit JalSISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, PrintDBG, Ent1, HSysMenu,
  hmsgbox,HStatus,HEnt1,{$IFNDEF DBXPRESS}dbtables, HTB97{$ELSE}uDbxDataSet{$ENDIF}, ColMemo, ComCtrls, HTB97, VisuEnr, HCompte,RappType,ImpFicU,TImpFic,
{$IFNDEF PROGEXT}
  RapSuppr,
{$ENDIF}

  Menus ;

Procedure VisuJalSISCO (StFichier : String) ;
Function  VisuJalSISCO2 (StFichier : String) : String ;


type
  TFJalSISCO = class(TForm)
    Outils: TPanel;
    HMTrad: THSystemMenu;
    G: THGrid;
    FindMvt: TFindDialog;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Uutils97: TToolbar97;
    BImprimer: TToolbarButton97;
    BRecherche: TToolbarButton97;
    ToolbarSep971: TToolbarSep97;
    Valide97: TToolbar97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
  private
    StJal : String ;
    Function ChargeGrid : Integer ;
  public
    StFichier : String ;
    FindFirst : Boolean ;
  end;

implementation

{$R *.DFM}


Procedure VisuJalSISCO (StFichier : String) ;
Var X : TFJalSISCO ;
BEGIN
X:=TFJalSISCO.Create(Application) ;
 Try
  X.StFichier:=StFichier ;
  X.ShowModal ;
 Finally
  X.Free ;
 End ;
END ;

Function VisuJalSISCO2 (StFichier : String) : String ;
Var X : TFJalSISCO ;
BEGIN
Result:='' ;
X:=TFJalSISCO.Create(Application) ;
 Try
  X.StFichier:=StFichier ;
  X.BValider.Visible:=TRUE ;
  X.G.MultiSelect:=TRUE ;
  X.ShowModal ;
 Finally
  If X.ModalResult=mrOk Then Result:=X.StJal ;
  X.Free ;
 End ;
END ;

Function TFJalSISCO.ChargeGrid : Integer ;
Var FF   : TextFile ;
    St : String ;
    OkOk : Boolean ;
BEGIN
Result:=0 ;
AssignFile(FF,StFichier) ;
{$I-} Reset(FF) ; {$I+} if IoResult<>0 then Exit ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
OkOk:=FALSE ;
While Not EOF(FF) do
   BEGIN
   MoveCur(FALSE) ;
   Readln(FF,St) ;
   If (Copy(St,1,2)='05')  Then
     BEGIN
     OkOk:=TRUE ;
     G.Cells[0,G.RowCount-1]:=Trim(Copy(St,3,2)) ;
     G.Cells[1,G.RowCount-1]:=Trim(Copy(St,5,20)) ;
     G.RowCount:=G.RowCount+1 ;
     END Else If (Copy(St,1,2)<>'05') And (Copy(St,1,2)<>'06') And OkOk Then Break ;
   END ;
G.RowCount:=G.rowCount-1 ;
CloseFile(FF) ;
FiniMove ;
END ;


procedure TFJalSISCO.FormShow(Sender: TObject);
Var i : Integer ;
begin
Caption:=Caption+' '+StFichier ;
ChargeGrid ;
end;

procedure TFJalSISCO.BImprimerClick(Sender: TObject);
begin
PrintDBGrid(G,Nil,Caption,'') ;
end;

procedure TFJalSISCO.BRechercheClick(Sender: TObject);
begin
FindFirst:=True ; FindMvt.Execute ;
end;

procedure TFJalSISCO.FindMvtFind(Sender: TObject);
begin
Rechercher(G,FindMvt,FindFirst) ;
end;

procedure TFJalSISCO.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
StJal:='' ;
For i:=0 To G.RowCount-1 Do If G.IsSelected(i) And (Length(StJal)<190) Then
  BEGIN
  StJal:=StJal+G.Cells[0,i]+';' ;
  END ;
end;

end.

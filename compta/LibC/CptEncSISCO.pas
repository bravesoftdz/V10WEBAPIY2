unit CptEncSISCO;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, StdCtrls, Buttons, ExtCtrls, PrintDBG, Ent1, HSysMenu,
  hmsgbox,HStatus,HEnt1,{$IFNDEF DBXPRESS}dbtables, HTB97{$ELSE}uDbxDataSet{$ENDIF}, ColMemo, ComCtrls, HTB97, VisuEnr, HCompte,RappType,ImpFicU,TImpFic,
{$IFNDEF PROGEXT}
  RapSuppr,
{$ENDIF}

  Menus ;

Procedure VisuCptHTEncSISCO (StFichier : String) ;


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
    procedure FormShow(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindMvtFind(Sender: TObject);
  private
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

Function TFJalSISCO.ChargeGrid : Integer ;
Var FF   : TextFile ;
    St,Cpt : String ;
    OkOk : Boolean ;
    LL : tStringList ;
BEGIN
Result:=0 ;
AssignFile(FF,StFichier) ;
{$I-} Reset(FF) ; {$I+} if IoResult<>0 then Exit ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
OkOk:=FALSE ;
LL:=tStringList.Create ;
LL.Sorted:=TRUE ; LL.Duplicates:=dupIgnore ;
While Not EOF(FF) do
   BEGIN
   MoveCur(FALSE) ;
   Readln(FF,St) ;
   If (Copy(St,1,1)='t')  Then
     BEGIN
     Cpt:=Copy(St,2,10) ;
     If LL.IndexOf(Cpt)<0 Then
       BEGIN
       OkOk:=TRUE ;
       G.Cells[0,G.RowCount-1]:=Trim(Copy(St,3,2)) ;
       G.Cells[1,G.RowCount-1]:=Trim(Copy(St,5,20)) ;
       G.RowCount:=G.RowCount+1 ;
       END ;
     END Else If (Copy(St,1,2)<>'05') And (Copy(St,1,2)<>'06') And OkOk Then Break ;
   END ;
G.RowCount:=G.rowCount-1 ;
CloseFile(FF) ;
FiniMove ;
LL.Clear ;
LL.Free ;
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

end.

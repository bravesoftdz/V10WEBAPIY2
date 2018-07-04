unit AfficheParam;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, Utob, StdCtrls, ComCtrls;

type
  TAffParam = class(TForm)
    HGrid1: THGrid;
    GCA_REFERENCE: TEdit;
    GCA_LIBELLE: TEdit;
    TreeView1: TTreeView;
    procedure FormActivate(Sender: TObject);
    procedure TreeView1Change(Sender: TObject; Node: TTreeNode);
    procedure HGrid1CellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}


procedure TAffParam.FormActivate(Sender: TObject);
var T1 : TOB;
begin
    T1 := Tob.Create('Liste des arrondis', NIL, -1);
    Hgrid1.VidePile(true) ;
    T1.LoadDetailDB ('ARRONDI', '', '',NIL,FALSE,TRUE);
    T1.PutGridDetail (HGrid1, TRUE, TRUE, 'GAR_CODEARRONDI;GAR_LIBELLE');
    T1.PutTreeView (TreeView1, NIL, '"Arrondi"; GAR_LIBELLE',-1);
end;

procedure TAffParam.TreeView1Change(Sender: TObject; Node: TTreeNode);
var T3 : TOB;
begin
    if  TreeView1.selected = Nil then exit;
    T3 := TOB (TreeView1.selected.data);
    T3.PutEcran(self);
end;

procedure TAffParam.HGrid1CellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var T2 : TOB;
begin
   T2 := TOB(Hgrid1.objects[0,Hgrid1.row]);
   T2.PutEcran(self);
end;

end.

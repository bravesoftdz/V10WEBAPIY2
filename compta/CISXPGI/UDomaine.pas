unit UDomaine;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichGrid, HSysMenu, hmsgbox, Db, DBTables, Hqry, HTB97, StdCtrls,
  ExtCtrls, DBCtrls, Grids, DBGrids, HDB, HPanel, UIUtil,Hctrls;

type
  TFFicheDomaine = class(TFFicheGrid)
    procedure FormShow(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FFicheDomaine: TFFicheDomaine;

procedure ParamDomained(PP : THPanel);

implementation

uses UDMIMP;

{$R *.DFM}

procedure ParamDomained(PP : THPanel);
var XX : TFFicheDomaine ;
BEGIN
XX:=TFFicheDomaine.Create(Application);
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free ;
    End ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
end;

procedure TFFicheDomaine.FormShow(Sender: TObject);
begin
//********
end;

procedure TFFicheDomaine.BDeleteClick(Sender: TObject);
var
  tbcharger: TTable;
begin
  tbcharger := TTable.create(Application);
  tbcharger.DatabaseName := DMImport.DBGlobal.DatabaseName;
  tbcharger.Tablename := DMImport.GzImpDomaine.TableName;
  tbCharger.open;
  try
    if tbCharger.findKey([FListe.SelectedField.Text]) then
    begin
      tbCharger.Delete;
    end;
  except
    PGIBox(Exception(ExceptObject).message, '');
    exit;
  end;
  tbcharger.free;

end;


end.

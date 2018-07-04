{***********UNITE*************************************************
Auteur  ...... : Mathieu GRANDIS
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : Sélection des scripts de reconnaissance
Mots clefs ... : 
*****************************************************************}
unit URecoSelecScripts;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, HSysMenu, HTB97,HmsgBox, Db, DBTables, Hqry, ADODB,
  udbxDataset;

type
  TFSelecScript = class(TFVierge)
    LblSelecScripts: TLabel;
    ListScriptsDispo: TListBox;
    ListScriptsSelec: TListBox;
    BMonter : TToolbarButton97;
    BAjouter: TToolbarButton97;
    BRetirer: TToolbarButton97;
    BBaisser: TToolbarButton97;
    SQ      : TDataSource;
    Q       : THQuery;
    Label1: TLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure BAjouterClick(Sender: TObject);
    procedure BBaisserClick(Sender: TObject);
    procedure BRetirerClick(Sender: TObject);
    procedure BMonterClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  FSelecScript: TFSelecScript;

implementation

uses UDMIMP;

{$R *.DFM}

procedure TFSelecScript.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  Action := caFree;
end;

procedure TFSelecScript.FormCreate(Sender: TObject);
var
  SelectQL : string;
begin
  inherited;

  // Récupérons les scripts de reconnaissance
  Q.DatabaseName := DMImport.DBGlobal.Name;
  SelectQL := 'SELECT CLE FROM '+ DMImport.GzImpReq.TableName + ' WHERE Domaine=' + '"R"';
  Q.Close;
  Q.SQL.Clear;
  Q.SQL.Add(SelectQL);
  Q.Open;

  // Remplissons la ListBox
  while not Q.Eof do
  begin
    ListScriptsDispo.Items.Add( Q.FindField('CLE').AsString );
    Q.Next;
  end;

  Q.Close;
  Q.Free;
end;


procedure TFSelecScript.BAjouterClick(Sender: TObject);
var
  N : integer;
begin
  inherited;
    for N:=ListScriptsDispo.Items.Count-1 downto 0 do
    begin
      if ListScriptsDispo.Selected[N] then
      begin
        ListScriptsSelec.Items.AddObject(ListScriptsDispo.Items[N], ListScriptsDispo.Items.Objects[N]);
        ListScriptsDispo.Items.Delete(N);
      end;
    end;
end;


procedure TFSelecScript.BBaisserClick(Sender: TObject);
var
  N : integer;
begin
  inherited;
  N := ListScriptsSelec.ItemIndex;
  if (N <> ListScriptsSelec.items.count-1) and (ListScriptsSelec.items.count <> 0) then
  begin
     ListScriptsSelec.Items.Exchange( N,N+1);
  end;
end;


procedure TFSelecScript.BRetirerClick(Sender: TObject);
var
  N : integer;
begin
  inherited;
  for N:=ListScriptsSelec.Items.Count-1 downto 0 do
    begin
      if ListScriptsSelec.Selected[N] then
      begin
        ListScriptsDispo.Items.AddObject(ListScriptsSelec.Items[N], ListScriptsSelec.Items.Objects[N]);
        ListScriptsSelec.Items.Delete(N);
      end;
    end;
end;

procedure TFSelecScript.BMonterClick(Sender: TObject);
var
  N : integer;
begin
  inherited;
  N := ListScriptsSelec.ItemIndex;
  if ( N <> 0 ) then
  begin
    ListScriptsSelec.Items.Exchange( N , N-1 );
  end;
end;

end.

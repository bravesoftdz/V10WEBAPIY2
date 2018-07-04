unit USelectCorresp;

interface

uses
  Windows, Messages, Classes, SysUtils, Graphics, Controls, StdCtrls, Forms,
  Dialogs, DBCtrls, DB, DBGrids,Grids, ExtCtrls,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB,{$ENDIF}
  HEnt1
  ;

type
  TFSelCorresp = class(TForm)
    QTablename: TStringField;
    QProfile: TStringField;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    Panel2: TPanel;
    Q: TQuery;
    ScriptSelect: TEdit;
    procedure FormShow(Sender: TObject);
    procedure DBGrid1DblClick(Sender: TObject);
  private
    { déclarations privées }
    TableName : string;
  public
    { déclarations publiques }
  end;

var
  FSelCorresp: TFSelCorresp;
Function  SelectCorresp(Latable : string) : String;

implementation

uses UDMIMP;

{$R *.DFM}

Function  SelectCorresp(Latable : string) : String;
BEGIN
  FSelCorresp :=TFSelCorresp.Create(Application) ;
  Try
       FSelCorresp.TableName := Latable;
       FSelCorresp.ShowModal ;
  Finally
       Result := FSelCorresp.ScriptSelect.Text;
       FSelCorresp.Free ;
  End ;
END ;


procedure TFSelCorresp.FormShow(Sender: TObject);
var
SelectQL,WWhere    : string;
begin
    //Q.DatabaseName := DMImport.DBGlobal.Name;
    if TableName = TraduireMemoire('Liste de correspondance') then
    WWhere := ' Where Ident = 1 and TableName="'+ TableName+'"'
    else
    if TableName <> '' then WWhere := ' Where TableName="'+ TableName+'"'
    else
    WWhere := '';
    SelectQL := ' SELECT Profile,TableName,Codedepart from '+DMImport.GzImpCorresp.TableName+ WWhere;
    if  not DMImport.DBGlobal.Connected then
     DMImport.DBGlobal.Connected := TRUE ;
    Q := OpenSQLADO (SelectQL, DMImport.DBGlobal.ConnectionString);
    DataSource1.DataSet := Q;
{$IFNDEF EAGLCLIENT}
    CentreDBGrid(DBGrid1);
{$ENDIF}
end;

procedure TFSelCorresp.DBGrid1DblClick(Sender: TObject);
begin
	ScriptSelect.Text := DBGrid1.Fields[0].Text;
        Q.DatabaseName := DMImport.DBGlobal.Name;
	ModalResult := mrOK;
end;

end.
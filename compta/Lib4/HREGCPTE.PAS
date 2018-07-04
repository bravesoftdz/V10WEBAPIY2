unit HRegCpte ;

interface

uses
  SysUtils, WinTypes, WinProcs, Messages, Classes, Graphics, Controls,
  Forms, Dialogs, StdCtrls, DB,
  {$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF !EAGLCLIENT}
  Buttons, ExtCtrls, Hctrls, Grids,
  DBGrids, hmsgbox,HEnt1,HSysMenu,Ent1,HCompte,dbCtrls ;

PROCEDURE Register;

type
  THDBCpteEdit = class(THCpteEdit)
  private  { Déclarations private }
    FdataLink : TFieldDataLink;
    Procedure DataChange(sender:Tobject);
    function getDataField: String;
    Function GetDataSource : TDataSource;
    Procedure SetDataField(const value:String);
    Procedure SetDataSource(value : TDataSource);
    procedure CMEnter(var Message: TCMGotFocus); message CM_ENTER;
    Procedure CMExit(var message:TCMExit); message CM_EXIT;
    Procedure UpdateData(Sender:Tobject);
  protected  { Déclarations protected }
    procedure Change; override;
  public    { Déclarations public }
    Constructor Create(AOwner:Tcomponent); override;
    Destructor Destroy; override;
  published  { Déclarations published }
    Property DataField : string read GetDataField write SetDataField;
    property DataSource: TDataSource read GetDataSource write SetDataSource;
  end;

implementation

PROCEDURE Register;
BEGIN
  RegisterComponents('Halley', [THCpteEdit,THDBCpteEdit]);
END ;

{===============================================================================}
{============================= CPTEEDIT ====================================}
{===============================================================================}
Constructor THDBCpteEdit.Create(AOwner : Tcomponent);
begin
inherited create(AOwner);
FdataLink:= TfieldDataLink.Create;
Fdatalink.OnDataChange := DataChange;
FdataLink.OnUpdateData := UpdateData;
end;

Destructor THDBCpteEdit.Destroy;
begin
Fdatalink.Free;
Inherited destroy;
end;

Procedure THDBCpteEdit.DataChange(sender:tobject);
begin
If FdataLink.Field = nil then Text:='' Else Text:=Fdatalink.Field.AsString;
end;

Function THDBCpteEdit.GetDataField : String;
begin
result := FdataLink.FieldName;
end;

Function THDBCpteEdit.GetDataSource : TDataSource;
begin
 Result := FdataLink.DataSource;
end;

Procedure THDBCpteEdit.SetDataField(const value : string);
begin
 FdataLink.FieldName:=Value;
end;

procedure THDBCpteEdit.SetDataSource(value : TDataSource);
begin
 FdataLink.DataSource:=Value;
end;

Procedure THDBCpteEdit.UpdateData(Sender: TObject);
begin
if Fdatalink.edit then FdataLink.Field.AsString:=Text     ;
(*                  else Text:=Fdatalink.Field.AsString;*)
end;

Procedure THDBCpteEdit.Change;
begin
if FDataLink<>Nil then FdataLink.Modified;
 Inherited Change;
end;

procedure THDBCpteEdit.CMEnter(var Message: TCMGotFocus);
begin
  inherited;
  UpdateData(Self);
end;

Procedure THDBCpteEdit.CMEXIT(Var message:TCMEXIT);
begin
 try
  FdataLink.UpdateRecord;
 except
  setfocus;
  raise;
 end;
 inherited;
end;


end.


{
Free Delphi Database Programming Course

CHAPTER 11: From Paradox to Access with ADO and Delphi
http://delphi.about.com/library/weekly/aa062601a.htm
Focusing on the TADOCommand components and using the SQL DDL language to help porting your BDE/Paradox data to ADO/Access.

Zarko Gajic, BSCS
About.com Guide to Delphi Programming
http://delphi.about.com
email: delphi.guide@about.com
free newsletter: http://delphi.about.com/library/blnewsletter.htm
forum: http://forums.about.com/ab-delphi/start/
}

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  DBTables, ADODB, Db, StdCtrls;

type
  TForm1 = class(TForm)
    cboBDETblNames: TComboBox;
    Memo1: TMemo;
    Button1: TButton;
    Button2: TButton;
    ADOConnection1: TADOConnection;
    ADOTable: TADOTable;
    ADOCommand: TADOCommand;
    BDETable: TTable;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private
    { Private declarations }
    Construction : Boolean;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}

function AccessType(fd:TFieldDef):string;
begin
 case fd.DataType of
  ftString: Result:='TEXT('+IntToStr(fd.Size)+')';
  ftSmallint: Result:='SMALLINT';
  ftInteger: Result:='INTEGER';
  ftWord: Result:='WORD';
  ftBoolean: Result:='YESNO';
  ftFloat : Result:='FLOAT';
  ftCurrency: Result := 'CURRENCY';
  ftDate, ftTime, ftDateTime: Result := 'DATETIME';
  ftAutoInc: Result := 'COUNTER';
  ftBlob, ftGraphic: Result := 'LONGBINARY';
  ftMemo, ftFmtMemo: Result := 'MEMO';
 else
  Result:='MEMO';
 end;
end;



procedure TForm1.FormCreate(Sender: TObject);
begin

 Session.GetTableNames('DBDEMOS', '*.db',False, False, cboBDETblNames.Items);
 Construction := FALSE;
end;

procedure TForm1.Button1Click(Sender: TObject);
var i:integer;
    s:string;
begin
 BDETable.TableName:='C:\devpgi\CISX\'+cboBDETblNames.Text;
 BDETable.FieldDefs.Update;

 s:='CREATE TABLE ' + cboBDETblNames.Text + ' (';
 with BDETable.FieldDefs do begin
  for i:=0 to Count-1 do begin
   s:=s + ' ' + Items[i].Name;
   s:=s + ' ' + AccessType(Items[i]);
   s:=s + ',';
  end; //for
  s[Length(s)]:=')';
 end;//with

 Memo1.Clear;
 Memo1.lines.Add (s);
 Construction := TRUE;

end;

procedure TForm1.Button2Click(Sender: TObject);
var i:integer;
    tblName:string;
begin
 tblName:=cboBDETblNames.Text;

//refresh
 if not Construction then Button1Click(Sender);

//drop & create table
// ADOCommand.CommandText:='DROP TABLE ' + tblName;
// ADOCommand.Execute;

 ADOCommand.CommandText:=Memo1.Text;
 ADOCommand.Execute;

 ADOTable.TableName:=tblName;

//copy data
 BDETable.Open;
 ADOTable.Open;
 try
  while not BDETable.Eof do begin
   ADOTable.Insert;
   for i:=0 to BDETable.Fields.Count-1 do begin
    ADOTable.FieldByName
   (BDETable.FieldDefs[i].Name).Value :=
      BDETable.Fields[i].Value;
   end;//for
   ADOTable.Post;
   BDETable.Next
  end;//while
 finally
  BDETable.Close;
  ADOTable.Close;
 end;//try

end;

end.

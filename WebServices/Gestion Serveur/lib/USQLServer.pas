unit USQLServer;

interface
uses DB, ADODB, Variants, ActiveX,
     ComObj, AdoInt, OleDB,classes,SysUtils,Dialogs;

procedure ListAvailableSQLServers(Names : TStringList);
procedure ListDatabasesOnServer(Server : string; Databases : TStringList);
procedure ListAvailableUsers (SERVER,DATABASE: string;UserList,UserValue : TStringList);
function getUserFromSql (SERVER,DATABASE,RESSOURCE : string) : string;

implementation

procedure ListAvailableSQLServers(Names : TStringList);
var
  RSCon: ADORecordsetConstruction;
  Rowset: IRowset;
  SourcesRowset: ISourcesRowset;
  SourcesRecordset: _Recordset;
  SourcesName, SourcesType: TField;

    function PtCreateADOObject
             (const ClassID: TGUID): IUnknown;
    var
      Status: HResult;
      FPUControlWord: Word;
    begin
      asm
        FNSTCW FPUControlWord
      end;
      Status := CoCreateInstance(
                  CLASS_Recordset,
                  nil,
                  CLSCTX_INPROC_SERVER or
                  CLSCTX_LOCAL_SERVER,
                  IUnknown,
                  Result);
      asm
        FNCLEX
        FLDCW FPUControlWord
      end;
      OleCheck(Status);
    end;
begin
  SourcesRecordset := 
      PtCreateADOObject(CLASS_Recordset) 
      as _Recordset;
  RSCon := 
      SourcesRecordset 
      as ADORecordsetConstruction;
  SourcesRowset := 
      CreateComObject(ProgIDToClassID('SQLOLEDB Enumerator')) 
      as ISourcesRowset;
  OleCheck(SourcesRowset.GetSourcesRowset(
       nil, 
       IRowset, 0, 
       nil, 
       IUnknown(Rowset)));
  RSCon.Rowset := RowSet;
  with TADODataSet.Create(nil) do
  try
    Recordset := SourcesRecordset;
    SourcesName := FieldByName('SOURCES_NAME');
    SourcesType := FieldByName('SOURCES_TYPE');
    Names.BeginUpdate;
    try
      while not EOF do
      begin
        if 
           (SourcesType.AsInteger = DBSOURCETYPE_DATASOURCE) 
           and (SourcesName.AsString <> '') then
          Names.Add(SourcesName.AsString);
        Next;
      end;
    finally
      Names.EndUpdate;
    end;
  finally
    Free;
  end;
end;

procedure ListDatabasesOnServer(Server : string; Databases : TStringList);
var
	QQ : TADOQuery;
  Conn : string;
begin
  Databases.Clear;
  Conn := 'Provider=SQLOLEDB.1;Password=ADMIN;User ID=ADMIN;Data Source='+Server+';Auto Translate=True;Use Encryption for Data=False;';
  QQ := TADOQuery.create (nil);
  with QQ do
  try
    ConnectionString := Conn;

    SQL.clear;
    SQL.Add('SELECT name FROM master..sysdatabases where dbid > 4');
    try
      Open;
      Databases.BeginUpdate;
      while not Eof do
      begin
        Databases.Add( QQ.Fields [0].AsString);
        Next;
      end;
    	Databases.EndUpdate;
    finally
    	Close;
    end;
  except
      on e:exception do MessageDlg(e.Message,mtError, [mbOK],0);
  end;
  QQ.Free;
end;


procedure ListAvailableUsers (SERVER,DATABASE: string;UserList,UserValue : TStringList);
var  Conn : string;
		 QQ : TADOQuery;
begin
	UserList.Clear;
  UserValue.Clear;
  QQ :=  TAdoQuery.Create(nil);
  try
    //simple ConnectionString without the DB name
    Conn := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+Database+';Data Source='+Server+';';
    QQ.ConnectionString := Conn;
    QQ.SQL.clear;
    QQ.SQL.Add('SELECT ARS_RESSOURCE,ARS_LIBELLE FROM RESSOURCE WHERE ARS_TYPERESSOURCE=''SAL'' AND ARS_UTILASSOCIE<>''''');
    try
  		QQ.Open;
      while not QQ.Eof do
      begin
        UserList.Add (QQ.Fields [1].AsString);
        UserValue.add (QQ.Fields [0].AsString);
        QQ.Next;
      end;
      QQ.Close;
    except
      on e:exception do
        MessageDlg(e.Message,mtError, [mbOK],0);
    end;
  finally
    QQ.Free;
  end;
end;

function getUserFromSql (SERVER,DATABASE,RESSOURCE : string) : string;
var  Conn : string;
		 QQ : TADOQuery;
     Sql : string;
begin
  SQl := 'SELECT US_ABREGE FROM UTILISAT WHERE US_UTILISATEUR=(SELECT ARS_UTILASSOCIE FROM RESSOURCE WHERE ARS_RESSOURCE='''+Ressource+''')';
  QQ :=  TAdoQuery.Create(nil);
  try
    //simple ConnectionString without the DB name
    Conn := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+Database+';Data Source='+Server+';';
    QQ.ConnectionString := Conn;
    QQ.SQL.clear;

    QQ.SQL.Add(SQL);
    try
  		QQ.Open;
      QQ.First;
      Result := QQ.Fields [0].AsString;
      QQ.Close;
    except
      on e:exception do
        MessageDlg(e.Message,mtError, [mbOK],0);
    end;
  finally
    QQ.Free;
  end;
end;

end.

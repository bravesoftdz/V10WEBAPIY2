unit BTPGetVersions;

interface

uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  CBPDatabases, forms,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Hdb,
  majtable,
{$ELSE}
  uWaini,
{$ENDIF}
  uhttp,
  utob;

const    HKEY_LOCAL_MACHINE    = $80000002;

{$IFNDEF EAGLCLIENT}
function Okversion (VersionRef : string) : boolean;
function BTPGetVersionRef  : string;
function BTPGetVersion : string;
{$ELSE}

{$IFDEF EAGLSERVER}
procedure BTPGetVersionRef (TOBResult : TOB) ;
{$ENDIF}

{$ENDIF}


implementation

{$IFNDEF EAGLCLIENT}
function Okversion (VersionRef : string) : boolean;
var QQ : TQuery;
		VersionLoc : string;
begin
	result := false;
  QQ := OPenSql ('SELECT BTV_VERSIONBASEB FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT="VERSION"',true,1,'',false);
  TRY
    if not QQ.eof then
    begin
      VersionLoc := QQ.findField('BTV_VERSIONBASEB').AsString;
      if VersionLoc >= VersionRef then result := true;
    end;
  FINALLY
  	ferme (QQ);
  END;
end;

function BTPGetVersion : string;
VAR QQ : TQuery;
begin
  Result := '';
  QQ := OPenSql ('SELECT BTV_VERSIONBASEB FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT="VERSION"',true,1,'',false);
  TRY
    if not QQ.eof then
    begin
      result := QQ.findField('BTV_VERSIONBASEB').AsString;
    end;
  FINALLY
  	ferme (QQ);
  END;

end;

function BTPGetVersionRef : string;
var DbRef : TDataBase;
    DestDriver: TDBDriver;
    DestODBC: Boolean;
    RefDriver : TDBDriver;
    RefODBC : Boolean;
    QQ : TQuery;
begin
	result := '';
  DestDriver := V_PGI.Driver; DestODBC := V_PGI.ODBC; V_PGI.StopCourrier := TRUE;
  //
  DBRef := TDatabase.Create(Application);
  DBRef.Name := 'DBREF';
  // Overture de la base REF
  ConnecteDB('Reference', DBRef, 'DBREF');
  RefDriver := V_PGI.Driver; RefODBC := V_PGI.ODBC;
  QQ := OpenSqlDb (DBREF,'SELECT BTV_VERSIONBASEB FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT="VERSION"',true,-1,'',true);
  if not QQ.eof then result := QQ.FindField('BTV_VERSIONBASEB').AsString ;
  ferme (QQ);
  DBRef.Connected := False;
  DBRef.Free;
  V_PGI.Driver := DestDriver; V_PGI.ODBC := DestODBC; V_PGI.StopCourrier := FALSE;
end;


{$ELSE !EAGLCLIENT}
{$IFDEF EAGLSERVER}
procedure BTPGetVersionRef (TOBResult : TOB);
var DBRef : TDatabase;
		QQ :TQuery;
    MaBase : String;
    RepInstCWS : string;
begin
  RepInstCWS:=GetFromRegistry(HKEY_LOCAL_MACHINE,'Software\CEGID\Cegid Business','DIRCWAS',RepInstCWS,TRUE) ;
  if RepInstCWS = '' then
  begin
  	Raise Exception.Create ('Erreur d''inscription du répertoire d''installation du serveur CWAS') ;
    exit;
  end;
  //
	MaBase := IncludeTrailingBackslash(repInstCWS)+'SocRefBTP.mdb';
  //
  DBREF := CBPDatabases.TCBPDataBases.CreateNamedDataBase ('Reference');
	DBREF.LoginPrompt := False ;
	DBREF.DriverName := 'MSACCESS' ;
  DBREF.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+ MaBase +
  ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
  'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;'+
  'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";'+
  'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;'+
  'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False';
  TRY
  //
    If Not CBPDatabases.TCBPDatabases.ConnectDatabase (DBREF) then
    begin
      Raise Exception.Create ('Erreur de connexion à la base…') ;
    end;
    QQ := OpenSQLDb (DBREF,'SELECT BTV_VERSIONBASEB FROM BTMAJSTRUCTURES WHERE BTV_TYPEELT="VERSION"',true);
    TRY
      if not QQ.eof then
      begin
        TOBResult.AddChampSupValeur('VERSIONREF',QQ.findField('BTV_VERSIONBASEB').AsString);
      end;
    FINALLY
      ferme (QQ);
    END;
  FINALLY
  	CBPDatabases.TCBPDataBases.FreeDataBase(DBREF);
    DBREF := nil;
  end;
end;
{$ENDIF EAGLSERVER}
{$ENDIF !EAGLCLIENT}

end.

unit UAccesDatabase;

interface
uses DB, ADODB,
     AdoInt, OleDB,classes;

function FindTiers (Server,Database,CodeClient : PChar) : Boolean;
function SetConnectionString (Server,Database : PChar) : string;

implementation

function SetConnectionString (Server,Database : PChar) : string;
begin
  result := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+Database+';Data Source='+Server+';';
end;


function FindTiers (Server,Database ,CodeClient : PChar) : Boolean;
var
	QQ : TADOQuery;
  Conn : string;
begin
  Result := false;
  Conn := SetConnectionString (Server,Database);
  QQ := TADOQuery.create (nil);
  with QQ do
  try
    ConnectionString := Conn;

    SQL.clear;
    SQL.Add('SELECT T_TIERS FROM TIERS WHERE T_TIERS="'+CodeClient+'" AND T_NATUREAUXI="CLI"');
    try
      Open;
      Result := not QQ.Eof;
    finally
    	Close;
    end;
  finally
  	QQ.Free;
  end;
end;


end.

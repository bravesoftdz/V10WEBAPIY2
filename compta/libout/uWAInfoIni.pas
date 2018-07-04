unit uWAInfoIni;

interface
uses 
  sysutils,
  hent1,
  classes,
  uCptWA
{$IFNDEF EAGLCLIENT}

{$IFDEF ODBCDAC}
  ,
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac
{$ELSE}
,
  uDbxDataSet
{$ENDIF}

{$ENDIF EAGLCLIENT}
  ;

type
  cWAInfoIni = class (cCptWA)
  private
    function getNomDeBase: string;
    function getDriver: string;
    function getODBC: string;
    function getBase: string;
    function getUser: string;
    function getPass: string;
    function getOptions: string;
    function getserver: string;
    function getTeamLink: string;
    function getShare: string;
    function getPath: string;
    //function getGamme: string;
    function getGroupe: string;
    function getNotConnectable: boolean ;
    function getDBDriver: TDBDriver;

    procedure setNomDeBase (valeur: string) ;
    procedure setDriver (valeur: string) ;
    procedure setODBC (valeur: string) ;
    procedure setBase (valeur: string) ;
    procedure setUser (valeur: string) ;
    procedure setPass (valeur: string) ;
    procedure setOptions (valeur: string) ;
    procedure setserver (valeur: string) ;
    procedure setTeamLink (valeur: string) ;
    procedure setShare (valeur: string) ;
    //procedure setGamme (valeur: string) ;
    procedure setGroupe (valeur: string) ;
    procedure setPath (valeur: string) ;
    procedure setNotConnectable(valeur: boolean) ;
  public
    function clone () : cWAInfoIni;
    constructor create () ; override;
    destructor destroy; override;
    class procedure afficheTrace (WAInfoIni: cWAInfoIni) ;
  published
    property NomDeBase: string read getNomDeBase write setNomDeBase;
    property Driver: string read getDriver write setDriver;
    property ODBC: string read getODBC write setODBC;
    property Base: string read getBase write setBase;
    property User: string read getUser write SetUser;
    property Pass: string read getPass write setPass;
    property Options: string read getOptions write setOptions;
    property server: string read getserver write setserver;
    property TeamLink: string read getTeamLink write setTeamLink;
    property Share: string read getShare write setShare;
    property Path: string read getPath write setPath;
    //property Gamme: string read getGamme write setGamme;
    property Groupe: string read getGroupe write setGroupe;
    property NotConnectable: boolean read getNotConnectable write setNotConnectable;
    property DBDriver: TDBDriver read getDBDriver;
  end;

implementation

uses inifiles,
  majtable,
  forms
{$IFDEF EAGLCLIENT}
{$ELSE}
  ,
  uWAiniFile
{$IFNDEF BASEEXT}
  ,
  uWAiniBase
{$ENDIF BASEEXT}

{$ENDIF EAGLCLIENT}
{$IFDEF EAGLSERVER}
  ,
  eSession
{$ELSE}
{$ENDIF EAGLSERVER}
  ;

(*var
  INSTANCE_OWAINI: coWAINI = nil;
  delINSTANCE_OWAINI: boolean = false ;*)

const
  CST_NOMDEBASE       = 'NOMDEBASE' ;
  CST_DRIVER          = 'DRIVER' ;
  CST_ODBC            = 'ODBC' ;
  CST_BASE            = 'BASE' ;
  CST_USER            = 'USER' ;
  CST_PASS            = 'PASS' ;
  CST_OPTIONS         = 'OPTIONS' ;
  CST_SERVER          = 'SERVER' ;
  CST_TEAMLINK        = 'TEAMLINK' ;
  CST_SHARE           = 'SHARE' ;
  CST_PATH            = 'PATH' ;
  CST_GROUPE          = 'GROUPE' ;
  CST_NOTCONNECTABLE  = 'NOTCONNECTABLE' ;

constructor cWAInfoIni.create () ;
begin
  inherited;
  ReaffecteNomTable (ClassName, true) ;
end;

function cWAInfoIni.clone () : cWAInfoIni;
begin
  result := cWAInfoIni.create () ;
  if result <> nil then result.Dupliquer (self, true, true) ;
end;

destructor cWAInfoIni.destroy;
begin
  inherited; //
end;

class procedure cWAInfoIni.afficheTrace (WAInfoIni: cWAInfoIni) ;
begin
  try
    if WAInfoIni <> nil then
    begin
      with WAInfoIni do
      begin
        TraceExecution ('WAInfoIni.NomDeBase  : ' + NomDeBase) ;
        TraceExecution ('WAInfoIni.Driver     : ' + Driver) ;
        TraceExecution ('WAInfoIni.ODBC       : ' + ODBC) ;
        TraceExecution ('WAInfoIni.Base       : ' + Base) ;
        TraceExecution ('WAInfoIni.User       : ' + User) ;
        TraceExecution ('WAInfoIni.Pass       : ' + Pass) ;
        TraceExecution ('WAInfoIni.Options    : ' + Options) ;
        TraceExecution ('WAInfoIni.server     : ' + server) ;
        TraceExecution ('WAInfoIni.TeamLink   : ' + TeamLink) ;
        TraceExecution ('WAInfoIni.Share      : ' + Share) ;
        TraceExecution ('WAInfoIni.Path       : ' + Path) ;
      end;
    end
    else
    begin
      TraceExecution ('WAInfoIni.WAInfoIni  : nil') ;
    end;
  except
  end;
end;

function cWAInfoIni.getNomDeBase: string;
begin
  if (not self.FieldExists (CST_NOMDEBASE) ) then self.AddChampSupValeur (CST_NOMDEBASE, '') ;
  result := self.getString (CST_NOMDEBASE) ;
end;

function cWAInfoIni.getDriver: string;
begin
  if (not self.FieldExists (CST_DRIVER) ) then self.AddChampSupValeur (CST_DRIVER, '') ;
  result := self.getString (CST_DRIVER) ;
end;

function cWAInfoIni.getODBC: string;
begin
  if (not self.FieldExists (CST_ODBC) ) then self.AddChampSupValeur (CST_ODBC, '') ;
  result := self.getString (CST_ODBC) ;
end;

function cWAInfoIni.getBase: string;
begin
  if (not self.FieldExists (CST_BASE) ) then self.AddChampSupValeur (CST_BASE, '') ;
  result := self.getString (CST_BASE) ;
end;

function cWAInfoIni.getUser: string;
begin
  if (not self.FieldExists (CST_USER) ) then self.AddChampSupValeur (CST_USER, '') ;
  result := self.getString (CST_USER) ;
end;

function cWAInfoIni.getPass: string;
begin
  if (not self.FieldExists (CST_PASS) ) then self.AddChampSupValeur (CST_PASS, '') ;
  result := self.getString (CST_PASS) ;
end;

function cWAInfoIni.getOptions: string;
begin
  if (not self.FieldExists (CST_OPTIONS) ) then self.AddChampSupValeur (CST_OPTIONS, '') ;
  result := self.getString (CST_OPTIONS) ;
end;

function cWAInfoIni.getserver: string;
begin
  if (not self.FieldExists (CST_SERVER) ) then self.AddChampSupValeur (CST_SERVER, '') ;
  result := self.getString (CST_SERVER) ;
end;

function cWAInfoIni.getTeamLink: string;
begin
  if (not self.FieldExists (CST_TEAMLINK) ) then self.AddChampSupValeur (CST_TEAMLINK, '') ;
  result := self.getString (CST_TEAMLINK) ;
end;

function cWAInfoIni.getShare: string;
begin
  if (not self.FieldExists (CST_SHARE) ) then self.AddChampSupValeur (CST_SHARE, '') ;
  result := self.getString (CST_SHARE) ;
end;

function cWAInfoIni.getPath: string;
begin
  if (not self.FieldExists (CST_PATH) ) then self.AddChampSupValeur (CST_PATH, '') ;
  result := self.getString (CST_PATH) ;
end;

function cWAInfoIni.getGroupe: string;
begin
  if (not self.FieldExists (CST_GROUPE) ) then self.AddChampSupValeur (CST_GROUPE, '') ;
  result := self.getString (CST_GROUPE) ;
end;

function cWAInfoIni.getNotConnectable: boolean;
begin
  if (not self.FieldExists (CST_NOTCONNECTABLE) ) then self.AddChampSupValeur (CST_NOTCONNECTABLE, '') ;
  result := self.GetBoolean (CST_NOTCONNECTABLE) ;
end;

(*function cWAInfoIni.getGamme: string;
begin
  if (not self.FieldExists ('GAMME') ) then self.AddChampSupValeur ('GAMME', '') ;
  result := self.getString ('GAMME') ;
end;*)

function cWAInfoIni.getDBDriver: TDBDriver;
begin
  result := StToDriver (Driver, TRUE) ;
  //result:=xxDBDriver ;
end;

procedure cWAInfoIni.setNomDeBase (valeur: string) ;
begin
  if (not self.FieldExists (CST_NOMDEBASE) ) then self.AddChampSup (CST_NOMDEBASE, false) ;
  self.PutValue (CST_NOMDEBASE, valeur) ;
end;

procedure cWAInfoIni.setDriver (valeur: string) ;
begin
  if (not self.FieldExists (CST_DRIVER) ) then self.AddChampSup (CST_DRIVER, false) ;
  self.PutValue (CST_DRIVER, valeur) ;
end;

procedure cWAInfoIni.setODBC (valeur: string) ;
begin
  if (not self.FieldExists (CST_ODBC) ) then self.AddChampSup (CST_ODBC, false) ;
  self.PutValue (CST_ODBC, valeur) ;
end;

procedure cWAInfoIni.setBase (valeur: string) ;
begin
  if (not self.FieldExists (CST_BASE) ) then self.AddChampSup (CST_BASE, false) ;
  self.PutValue (CST_BASE, valeur) ;
end;

procedure cWAInfoIni.setUser (valeur: string) ;
begin
  if (not self.FieldExists (CST_USER) ) then self.AddChampSup (CST_USER, false) ;
  self.PutValue (CST_USER, valeur) ;
end;

procedure cWAInfoIni.setPass (valeur: string) ;
begin
  if (not self.FieldExists (CST_PASS) ) then self.AddChampSup (CST_PASS, false) ;
  self.PutValue (CST_PASS, valeur) ;
end;

procedure cWAInfoIni.setOptions (valeur: string) ;
begin
  if (not self.FieldExists (CST_OPTIONS) ) then self.AddChampSup (CST_OPTIONS, false) ;
  self.PutValue (CST_OPTIONS, valeur) ;
end;

procedure cWAInfoIni.setserver (valeur: string) ;
begin
  if (not self.FieldExists (CST_SERVER) ) then self.AddChampSup (CST_SERVER, false) ;
  self.PutValue (CST_SERVER, valeur) ;
end;

procedure cWAInfoIni.setTeamLink (valeur: string) ;
begin
  if (not self.FieldExists (CST_TEAMLINK) ) then self.AddChampSup (CST_TEAMLINK, false) ;
  self.PutValue (CST_TEAMLINK, valeur) ;
end;

procedure cWAInfoIni.setShare (valeur: string) ;
begin
  if (not self.FieldExists (CST_SHARE) ) then self.AddChampSup (CST_SHARE, false) ;
  self.PutValue (CST_SHARE, valeur) ;
end;

procedure cWAInfoIni.setPath (valeur: string) ;
begin
  if (not self.FieldExists (CST_PATH) ) then self.AddChampSup (CST_PATH, false) ;
  self.PutValue (CST_PATH, valeur) ;
end;

procedure cWAInfoIni.setGroupe (valeur: string) ;
begin
  if (not self.FieldExists (CST_GROUPE) ) then self.AddChampSup (CST_GROUPE, false) ;
  self.PutValue (CST_GROUPE, valeur) ;
end;

procedure cWAInfoIni.setNotConnectable (valeur: boolean) ;
begin
  if (not self.FieldExists (CST_NOTCONNECTABLE) ) then self.AddChampSup (CST_NOTCONNECTABLE, false) ;
  self.PutValue (CST_NOTCONNECTABLE, valeur) ;
end;

(*procedure cWAInfoIni.setGamme (valeur: string) ;
begin
  if (not self.FieldExists ('GAMME') ) then self.AddChampSup ('GAMME', false) ;
  self.PutValue ('GAMME', valeur) ;
end;    *)

end.

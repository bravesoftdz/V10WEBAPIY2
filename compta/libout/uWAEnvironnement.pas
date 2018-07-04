unit uWAEnvironnement;

interface
uses hctrls,
  sysutils,
  hent1,
  classes,
  uCptWA, utob
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
  cWAEnvironnement = class (cCptWA)
  private
    function getIDEnv: string;
    function getLibelle: string;
    function getVersion: string;
    function getBlocNote: string;
    function getSocref: string;
    function getPrefixeBase: string;

    procedure setIDEnv (valeur: string) ;
    procedure setLibelle (valeur: string) ;
    procedure setVersion (valeur: string) ;
    procedure setBlocNote (valeur: string) ;
    procedure setSocref (valeur: string) ;
    procedure setPrefixeBase (valeur: string) ;
  public
    function clone () : cWAEnvironnement;
    constructor create () ; override;
    destructor destroy; override;
    class procedure afficheTrace (WAEnvironnement: cWAEnvironnement) ;
    function setValuesFromTob(values: Tob): boolean ;
  published
    property IDEnv: string read getIDEnv write setIDEnv;
    property Libelle: string read getLibelle write setLibelle;
    property Version: string read getVersion write setVersion;
    property BlocNote: string read getBlocNote write setBlocNote;
    property Socref: string read getSocref write setSocref;
    property PrefixeBase: string read getPrefixeBase write setPrefixeBase;
  //08/08/2007
  private
    function getServeurCwas : string;
    procedure setServeurCWAS(valeur: string) ;
  published
    property SerVeurCwas : string read getServeurCwas write setServeurCwas ;

  //08/08/2007
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

const
  CST_IDENV       = 'IDEnv' ;
  CST_LIBELLE     = 'Libelle' ;
  CST_VERSION     = 'Version' ;
  CST_BLOCNOTE    = 'BlocNote' ;
  CST_SOCREF      = 'Socref' ;
  CST_PREFIXEBASE = 'PrefixeBase' ;
  CST_SERVEURCWAS = 'ServeurCwas' ; //08/08/2007

(*var
  INSTANCE_OWAINI: coWAINI = nil;
  delINSTANCE_OWAINI: boolean = false ;*)

constructor cWAEnvironnement.create () ;
begin
  inherited;
  ReaffecteNomTable (ClassName, true) ;
end;

function cWAEnvironnement.clone () : cWAEnvironnement;
begin
  result := cWAEnvironnement.create () ;
  if result <> nil then result.Dupliquer (self, true, true) ;
end;

destructor cWAEnvironnement.destroy;
begin
  inherited; //
end;

class procedure cWAEnvironnement.afficheTrace (WAEnvironnement: cWAEnvironnement) ;
begin
  try
    if WAEnvironnement <> nil then
    begin
      with WAEnvironnement do
      begin
        TraceExecution ('WAEnvironnement.IDEnv      : ' + IDEnv) ;
        TraceExecution ('WAEnvironnement.Libelle    : ' + Libelle) ;
        TraceExecution ('WAEnvironnement.Version    : ' + Version) ;
        TraceExecution ('WAEnvironnement.BlocNote   : ' + BlocNote) ;
        //08/08/2007
        TraceExecution ('WAEnvironnement.SocRef     : ' + Socref) ;
        TraceExecution ('WAEnvironnement.PrefixeBase: ' + PrefixeBase) ;
        TraceExecution ('WAEnvironnement.ServerCWAS : ' + ServeurCWAS) ;
        //08/08/2007
      end;
    end
    else
    begin
      TraceExecution ('WAEnvironnement.WAEnvironnement  : nil') ;
    end;
  except
  end;
end;

function cWAEnvironnement.setValuesFromTob(values: Tob): boolean ;
begin
result:=false ;
try
  if (values<>nil) and (values.detail.count>0) then
    begin
    setIDEnv(values.detail[0].getValue('AEN_IDENV')) ;
    setLibelle(values.detail[0].getValue('AEN_LIBELLE')) ;
    setVersion(values.detail[0].getValue('AEN_VERSION')) ;
    setBlocNote(values.detail[0].getValue('AEN_BLOCNOTE')) ;
    setSocref(values.detail[0].getValue('AEN_NOMSOCREF')) ;
    setPrefixeBase (values.detail[0].getValue('AEN_PREFIXEBASE')) ;
    setServeurCWAS(values.detail[0].getValue('AEN_SERVEURCWAS')) ; //08/08/2007
    result:=true ;
    end ;
  except
    on E: Exception do
      begin
      ErrorMessage := 'cWAEnvironnement.setValuesFromTob: ' + E.Message;
      end;
    end ;
end ;

function cWAEnvironnement.getIDEnv: string;
begin
  if (not self.FieldExists (CST_IDENV) ) then self.AddChampSupValeur (CST_IDENV, '') ;
  result := self.getString (CST_IDENV) ;
end;

function cWAEnvironnement.getLibelle: string;
begin
  if (not self.FieldExists (CST_LIBELLE) ) then self.AddChampSupValeur (CST_LIBELLE, '') ;
  result := self.getString (CST_LIBELLE) ;
end;

function cWAEnvironnement.getVersion: string;
begin
  if (not self.FieldExists (CST_VERSION) ) then self.AddChampSupValeur (CST_VERSION, '') ;
  result := self.getString (CST_VERSION) ;
end;

function cWAEnvironnement.getBlocNote: string;
begin
  if (not self.FieldExists (CST_BLOCNOTE) ) then self.AddChampSupValeur (CST_BLOCNOTE, '') ;
  result := self.getString (CST_BLOCNOTE) ;
end;

function cWAEnvironnement.getSocref(): string ;
begin
  if (not self.FieldExists (CST_SOCREF) ) then self.AddChampSupValeur (CST_SOCREF, '') ;
  result := self.getString (CST_SOCREF) ;
end ;

function cWAEnvironnement.getPrefixeBase(): string ;
begin
  if (not self.FieldExists (CST_PREFIXEBASE) ) then self.AddChampSupValeur (CST_PREFIXEBASE, '') ;
  result := self.getString (CST_PREFIXEBASE) ;
end ;

procedure cWAEnvironnement.setIDEnv (valeur: string) ;
begin
  if (not self.FieldExists (CST_IDENV) ) then self.AddChampSup (CST_IDENV, false) ;
  self.PutValue (CST_IDENV, valeur) ;
end;

procedure cWAEnvironnement.setLibelle (valeur: string) ;
begin
  if (not self.FieldExists (CST_LIBELLE) ) then self.AddChampSup (CST_LIBELLE, false) ;
  self.PutValue (CST_LIBELLE, valeur) ;
end;

procedure cWAEnvironnement.setVersion (valeur: string) ;
begin
  if (not self.FieldExists (CST_VERSION) ) then self.AddChampSup (CST_VERSION, false) ;
  self.PutValue (CST_VERSION, valeur) ;
end;

procedure cWAEnvironnement.setBlocNote (valeur: string) ;
begin
  if (not self.FieldExists (CST_BLOCNOTE) ) then self.AddChampSup (CST_BLOCNOTE, false) ;
  self.PutValue (CST_BLOCNOTE, valeur) ;
end;

procedure cWAEnvironnement.setSocref(valeur: string)  ;
begin
  if (not self.FieldExists (CST_SOCREF) ) then self.AddChampSup (CST_SOCREF, false) ;
  self.PutValue (CST_SOCREF, valeur) ;
end ;

procedure cWAEnvironnement.setPrefixeBase(valeur: string)  ;
begin
  if (not self.FieldExists (CST_PREFIXEBASE) ) then self.AddChampSup (CST_PREFIXEBASE, false) ;
  self.PutValue (CST_PREFIXEBASE, valeur) ;
end ;

//08/08/2007
function cWAEnvironnement.getServeurCwas : string;
begin
  if (not self.FieldExists (CST_SERVEURCWAS) ) then self.AddChampSupValeur (CST_SERVEURCWAS, '') ;
  result := self.getString (CST_SERVEURCWAS) ;
end ;

procedure cWAEnvironnement.setServeurCWAS(valeur: string) ;
begin
  if (not self.FieldExists (CST_SERVEURCWAS) ) then self.AddChampSup (CST_SERVEURCWAS, false) ;
  self.PutValue (CST_SERVEURCWAS, valeur) ;
end ;
//08/08/2007
end.

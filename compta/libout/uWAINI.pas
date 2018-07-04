unit uWAINI;

interface


uses hctrls,
  sysutils,
  hent1,
  classes,
  uCptWA,
  uWAInfoIni,
  uWAEnvironnement,
  utob
{$IFDEF EAGLSERVER}
  , eisapi
{$ELSE}
  , hmsgbox
{$ENDIF}
{$IFNDEF EAGLCLIENT}
{$IFDEF ODBCDAC}
  ,
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac
{$ELSE}
,
  uDbxDataSet,ActiveX
{$ENDIF}
{$ENDIF EAGLCLIENT}
  , cbpdatabases
  ;

type
  TDossierVerrou = (dvAucun, dvinconnu, dvMaj, dvEnl, dvSos, dvPar, dvMar);

  cWAINI = class (cCptWA)
  private
    function getMode: string; //type de ini (INI, BAS, ...)
    function getNomIni: string;
    //function getGamme: string;
    function getIDHeberge: string;
    function getIDClient: string;
    function getIDDossier: string;
    function getNomHeberge: string;
    function getNomClient: string;
    function getServer: string;
    function getDriver: string;
    function getForceBase: Boolean;

    procedure setMode (valeur: string) ;
    procedure setNomIni (valeur: string) ;
    //procedure setGamme (valeur: string) ;
    procedure setIDHeberge (valeur: string) ;
    procedure setIDClient (valeur: string) ;
    procedure setIDDossier (valeur: string) ;
    procedure setNomHeberge (valeur: string) ;
    procedure setNomClient (valeur: string) ;
    procedure setServer (valeur: string) ;
    procedure setDriver (valeur: string) ;
    procedure setForceBase (valeur: boolean) ;

  public
    destructor destroy; override;
    function clone () : cWAINI;
    constructor create () ; overload; override;
    class function HeritedInstance (AvecForce: boolean = false) : cWAINI;
    class procedure afficheTrace (wcWAINI: cWAINI) ;

    function GetNumSocRef () : integer;
    function GetNumSoc (WAInfoIni: cWAInfoIni) : integer;
    function ModifUser (NomBase: string; User, Password: string) : boolean;
    function AddUser (NomBase: string; User, Password: string) : boolean;
    function TestConnection (WAInfoIni: cWAInfoIni) : boolean;
    class function GetUserNumberFromSeria (GroupName: string) : integer;
    function DossierVerrouToString(Etat: TDossierVerrou): string;
    function StringToDossierVerrou(EtatString: string): TDossierVerrou;
    {$IFNDEF EAGLCLIENT}
    class procedure Deconnectedb (wDB: TDataBase) ;
    class function connectedb (WAInfoIni: cWAInfoIni; var wDB: TDataBase; stNom: string) : boolean; overload;
    function connectedb (NomBase: string; var wDB: TDataBase; stNom: string) : boolean; overload;
    class function Action (UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB) : boolean;
    {$ENDIF EAGLCLIENT}
    {$IFDEF EAGLSERVER}
    {$IFNDEF BASEEXT}
    class function isBaseChapeau (AvecForce: boolean = false) : boolean;
    {$ENDIF BASEEXT}
    {$ENDIF EAGLSERVER}

    //méthodes à surcharger dans les classes descendantes
    function Chargedossier (VireRef, VireDP, VireModel: boolean) : HTStringList; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function Ajoutedossier (WAInfoIni: cWAInfoIni) : boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function Supprimedossier (WAInfoIni: cWAInfoIni) : boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function Modifiedossier (WAInfoIni: cWAInfoIni) : boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function Existedossier (WAInfoIni: cWAInfoIni) : boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    procedure ChargeDBParams (WAInfoIni: cWAInfoIni) ; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function getListOfHeberge(): HTStringList; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function getListOfClient(): HTStringList; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function getListOfDossier(): HTStringList; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function UserPassValidate(user, pass: string): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function IsAdmin(login: string): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetHeberge: TOB; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetClient: TOB; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetDossier: TOB; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function setHeberge(laTob: TOB): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function setClient(laTob: TOB): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function setDossier(laTob: TOB; BaseFrom: string; AvecRecupHeberge: boolean): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function DelHeberge: boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function DelClient: boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function DelDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique: boolean): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetNextID(leQuel: integer): string; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function SetDossierVerrou(BaseNom: string; Etat: TDossierVerrou): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetDossierVerrou(BaseNom: string): TDossierVerrou; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function getProduit: string; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetMaxUserNumber: integer; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetSocRef () : string; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
    function IsFicBase(): boolean; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}

  protected
    procedure QuickSort (FDossier: HTStrings; deb, fin: Integer) ;
  published
    property Mode: string read getMode write setMode;
    property NomIni: string read getNomIni write setNomIni;
    //property Gamme: string read getGamme write setGamme;
    property IDHeberge: string read getIDHeberge write setIDHeberge;
    property IDClient: string read getIDClient write setIDClient;
    property IDDossier: string read getIDDossier write setIDDossier;
    property NomHeberge: string read getNomHeberge write setNomHeberge;
    property NomClient: string read getNomClient write setNomClient;
    property DRIVER: string read getDRIVER write setDRIVER;
    property SERVER: string read getSERVER write setSERVER;
    property ForceBase: boolean read getForceBase write setForceBase;
  public
    function GetListofEnvironnement() : HTStringList; {$IFNDEF EAGLCLIENT} virtual ; {$ENDIF}
  {$IFNDEF EAGLCLIENT} //YCP 15/05/07 
  private
    xxDB: TDatabase;
    function getDB: TDatabase;
    procedure setDB(valeur: TDatabase);
  public
    class function connectedb (WAInfoIni: cWAInfoIni; var wDB: TDataBase) : boolean; overload;
    function connectedb (NomBase: string; var wDB: TDataBase) : boolean; overload;
  published
    property LaDB: TDataBase read getDB write setDB ;
  {$ENDIF EAGLCLIENT}
  end;


const
  BASECHAPEAUINI: string = 'ChapeauPGI.ini';
implementation

uses windows,
  inifiles,
  majtable,
  forms,
  uWAiniFile
{$IFNDEF BASEEXT}
  ,
  uWAINIBase
{$ENDIF BASEEXT}
{$IFNDEF EAGLCLIENT}
  ,
  licutil
{$ENDIF EAGLCLIENT}

{$IFDEF EAGLSERVER}
  ,
  eSession
{$ENDIF EAGLSERVER}
  ;

//var //YCP 15/05/07 
//  nbConnect: integer = 0;
//  cpte: integer = 0;
  //fIsBaseChapeau: string = '';

{$IFDEF EAGLSERVER}
//  wDBSession: tSession; //YCP 15/05/07 
{$ENDIF EAGLSERVER}
  (*var
    INSTANCE_cWAINI: cWAINI = nil;
    delINSTANCE_cWAINI: boolean = false ;*)

const
  CST_IDHEBERGE    = 'IDHEBERGE' ;
  CST_IDCLIENT     = 'IDCLIENT' ;
  CST_IDDOSSIER    = 'IDDOSSIER' ;
  //CST_GAMME        = 'GAMME' ;
  CST_MODE         = 'MODE' ;
  CST_NOMINI       = 'NOMINI' ;
  CST_NOMHEBERGE   = 'NOMHEBERGE';
  CST_NOMCLIENT    = 'NOMCLIENT';
  CST_SERVER       = 'SERVER' ;
  CST_DRIVER       = 'DRIVER' ;
  CST_FORCEBASE    = 'FORCEBASE' ;

constructor cWAINI.create () ;
begin
  inherited;
  ReaffecteNomTable (className, true) ;
{$IFDEF EAGLCLIENT}
  nomIni := HalSocIni;
{$ELSE}
  nomIni := GetHalSocIni;
{$ENDIF EAGLCLIENT}
end;

{$IFNDEF EAGLCLIENT} //YCP 15/05/07 
function cWAINI.getDB: TDatabase;
begin
  if (xxDB = nil) then
    begin
    //inc(cpte);
    xxDB:=TCBPDatabases.createNamedDataBase('LOCALMASTER') ;
    end;

  result := xxDB;
end;

procedure cWAINI.setDB(valeur: TDatabase);
begin
  TCBPDatabases.freeDataBase(xxDB) ;
  xxDB := valeur;
end;
{$ENDIF EAGLCLIENT}

function getWindowsDir () : string;
var
  Buffer: array [0..1023] of Char;
begin
  GetWindowsDirectory (Buffer, sizeof (Buffer) ) ;
  result := Buffer;
end;

{$IFDEF EAGLSERVER}
{$IFNDEF BASEEXT}
class function cWAINI.isBaseChapeau (AvecForce: boolean = false) : boolean;
begin
  if TISession.PluginExists ('dbss')
    and (FileExists (getWindowsDir + '\' + uWAini.BASECHAPEAUINI)
         or FileExists (ExtractFilePath (Application.exeName) + '\' + BASECHAPEAUINI)
         or FileExists (ExtractFilePath (Application.exeName) + '..\' + BASECHAPEAUINI) //YCP Process serveur dans /Scripts, avec le .ini dans c:/CWS         
         )
    and (AvecForce or (uppercase (ExtractFileName (GetHalSocini) ) <> uppercase (BASECHAPEAUINI) )) then //YCP 03/01/05
    result:=true  //YCP 15/05/07 fIsBaseChapeau := 'OUI'
  else
    result:=false ; //YCP 15/05/07 fIsBaseChapeau := 'NON';
  //YCP 15/05/07 result := fIsBaseChapeau = 'OUI';
end;
{$ENDIF BASEEXT}
{$ENDIF EAGLSERVER}

class function cWAINI.HeritedInstance (AvecForce: boolean = false) : cWAINI;
begin
{$IFDEF EAGLSERVER}
{$IFNDEF BASEEXT}
  if isBaseChapeau (AvecForce) then
    result := cWAINIBase.create
  else
{$ENDIF BASEEXT}
{$ENDIF EAGLSERVER}
    result := cWAINIFile.create;
  TraceExecution ('Creation d''objet : ' + result.ClassName) ;
end;

function cWAINI.clone () : cWAINI;
begin
  TraceExecution ('Clone ' + si (forceBase, 'avec', 'sans') + ' forcçage') ;
  result := HeritedInstance (forceBase) ;
  if result <> nil then result.Dupliquer (self, true, true) ;
end;

destructor cWAINI.destroy;
begin
{$IFNDEF EAGLCLIENT} //YCP 15/05/07 
  if assigned(xxDB) then
    setDB(nil);
{$ENDIF EAGLCLIENT}
  inherited; //
end;

class procedure cWAINI.afficheTrace (wcWAINI: cWAINI) ;
begin
  with wcWAINI do
  begin
    TraceExecution (classname + '.Mode       : ' + Mode) ;
    TraceExecution (classname + '.NomIni     : ' + NomIni) ;
    //TraceExecution (classname + '.Gamme      : ' + Gamme) ;
    TraceExecution (classname + '.IDHeberge     : ' + IDHeberge) ;
    TraceExecution (classname + '.IDClient     : ' + IDClient) ;
    TraceExecution (classname + '.NomHeberge  : ' + NomHeberge) ;
    TraceExecution (classname + '.NomClient  : ' + NomClient) ;
  end;
end;

function cWAINI.getMode: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists (CST_MODE) ) then self.AddChampSupValeur (CST_MODE, '') ;
  result := self.getString (CST_MODE) ;
end;

function cWAINI.getNomIni: string;
var
  xxNomIni: string;
  filPath: String ;
begin
  if (not self.FieldExists (CST_NOMINI) ) then self.AddChampSupValeur (CST_NOMINI, '') ;
  xxNomIni := self.getString (CST_NOMINI) ;
  filPath:=lowerCase(ExtractFilePath (application.ExeName)) ;
  if (FileExists (filPath + '\' + xxNomIni) ) then
    xxNomIni := filPath + '\' + xxNomIni
  else if (pos('scripts', filPath)>0) and FileExists (filPath + '\..\' + xxNomIni) then
    xxNomIni := filPath + '\..\' + xxNomIni;

  result := xxNomIni;
end;

(*function cWAINI.getGamme: string;
begin
  if (not self.FieldExists (CST_GAMME) ) then self.AddChampSupValeur (CST_GAMME, '') ;
  result := self.getString (CST_GAMME) ;
end;*)

function cWAINI.getIDHeberge: string;
begin
  if self.FieldExists('ENTITE') and not self.FieldExists(CST_IDHEBERGE) then
    self.AddChampSupValeur (CST_IDHEBERGE, self.getString('ENTITE')) ;
  if (not self.FieldExists (CST_IDHEBERGE) ) then self.AddChampSupValeur (CST_IDHEBERGE, '') ;
  result := self.getString (CST_IDHEBERGE) ;
end;

function cWAINI.getIDClient: string;
begin
  if self.FieldExists('GROUPE') and not self.FieldExists(CST_IDCLIENT) then
    self.AddChampSupValeur (CST_IDCLIENT, self.getString('GROUPE')) ;
  if (not self.FieldExists (CST_IDCLIENT) ) then self.AddChampSupValeur (CST_IDCLIENT, '') ;
  result := self.getString (CST_IDCLIENT) ;
end;

function cWAINI.getIDDossier: string;
begin
  if self.FieldExists('DOSSIER') and not self.FieldExists(CST_IDDOSSIER) then
    self.AddChampSupValeur (CST_IDDOSSIER, self.getString('DOSSIER')) ;
  if (not self.FieldExists (CST_IDDOSSIER) ) then self.AddChampSupValeur (CST_IDDOSSIER, '') ;
  result := self.getString (CST_IDDOSSIER) ;
end;

function cWAINI.getNomHeberge: string;
begin
  if (not self.FieldExists (CST_NOMHEBERGE) ) then self.AddChampSupValeur (CST_NOMHEBERGE, '') ;
  result := self.getString (CST_NOMHEBERGE) ;
end;

function cWAINI.getNomClient: string;
begin
  if (not self.FieldExists (CST_NOMCLIENT) ) then self.AddChampSupValeur (CST_NOMCLIENT, '') ;
  result := self.getString (CST_NOMCLIENT) ;
end;

function cWAINI.getSERVER: string;
begin
  if (not self.FieldExists (CST_SERVER) ) then self.AddChampSupValeur (CST_SERVER, '') ;
  result := self.getString (CST_SERVER) ;
end;

function cWAINI.getDRIVER: string;
begin
  if (not self.FieldExists (CST_DRIVER) ) then self.AddChampSupValeur (CST_DRIVER, '') ;
  result := self.getString (CST_DRIVER) ;
end;

function cWAINI.getFORCEBASE: boolean;
begin
  if (not self.FieldExists (CST_FORCEBASE) ) then self.AddChampSupValeur (CST_FORCEBASE, '-') ;
  result := (self.getString (CST_FORCEBASE) = 'X') ;
end;

procedure cWAINI.setMode (valeur: string) ;
begin
  if (not self.FieldExists (CST_MODE) ) then self.AddChampSup (CST_MODE, false) ;
  self.PutValue (CST_MODE, valeur) ;
end;

procedure cWAINI.setNomIni (valeur: string) ;
begin
  if (not self.FieldExists (CST_NOMINI) ) then self.AddChampSup (CST_NOMINI, false) ;
  self.PutValue (CST_NOMINI, valeur) ;
end;

(*procedure cWAINI.setGamme (valeur: string) ;
begin
  if (not self.FieldExists (CST_GAMME) ) then self.AddChampSup (CST_GAMME, false) ;
  self.SetString (CST_GAMME, valeur) ;
end;*)

procedure cWAINI.setIDHeberge (valeur: string) ;
begin
  if (pos ('@', valeur) > 0) then
  begin
    NomHeberge := copy (valeur, pos ('@', valeur) + 1, length (valeur) ) ;
    NomClient := copy (valeur, 1, pos ('@', valeur) - 1) ;
  end
  else
  begin
    if (not self.FieldExists (CST_IDHEBERGE) ) then self.AddChampSup (CST_IDHEBERGE, false) ;
    self.PutValue (CST_IDHEBERGE, valeur) ;
  end;
end;

procedure cWAINI.setIDClient (valeur: string) ;
begin
  if (pos ('@', valeur) > 0) then
  begin
    NomHeberge := copy (valeur, pos ('@', valeur) + 1, length (valeur) ) ;
    NomClient := copy (valeur, 1, pos ('@', valeur) - 1) ;
    ;
  end
  else
  begin
    if (not self.FieldExists (CST_IDCLIENT) ) then self.AddChampSup (CST_IDCLIENT, false) ;
    self.PutValue (CST_IDCLIENT, valeur) ;
  end;
end;

procedure cWAINI.setIDDossier (valeur: string) ;
begin
  if (not self.FieldExists (CST_IDDOSSIER) ) then self.AddChampSup (CST_IDDOSSIER, false) ;
  self.PutValue (CST_IDDOSSIER, valeur) ;
end;

procedure cWAINI.setNomHeberge (valeur: string) ;
begin
  if (not self.FieldExists (CST_NOMHEBERGE) ) then self.AddChampSup (CST_NOMHEBERGE, false) ;
  self.PutValue (CST_NOMHEBERGE, valeur) ;
end;

procedure cWAINI.setNomClient (valeur: string) ;
begin
  if (not self.FieldExists (CST_NOMCLIENT) ) then self.AddChampSup (CST_NOMCLIENT, false) ;
  self.PutValue (CST_NOMCLIENT, valeur) ;
end;

procedure cWAINI.setSERVER (valeur: string) ;
begin
  if (not self.FieldExists (CST_SERVER) ) then self.AddChampSup (CST_SERVER, false) ;
  self.PutValue (CST_SERVER, valeur) ;
end;

procedure cWAINI.setDRIVER (valeur: string) ;
begin
  if (not self.FieldExists (CST_DRIVER) ) then self.AddChampSup (CST_DRIVER, false) ;
  self.PutValue (CST_DRIVER, valeur) ;
end;

procedure cWAINI.setFORCEBASE (valeur: boolean) ;
begin
  if (not self.FieldExists (CST_FORCEBASE) ) then self.AddChampSup (CST_FORCEBASE, false) ;
  self.PutValue (CST_FORCEBASE, si (valeur, 'X', '-') ) ;
end;

function cWAINI.DossierVerrouToString(Etat: TDossierVerrou): string;
begin
  case Etat of
    dvAucun: result := 'AUC';
    dvMaj: result := 'MAJ';
    dvEnl: result := 'ENL';
    dvSos: result := 'SOS';
    dvPar: result := 'PAR';
    dvMar: result := 'MAR';
  else
    result := '';
  end;
end;

function cWAINI.StringToDossierVerrou(EtatString: string): TDossierVerrou;
begin
  if EtatString = 'AUC' then
    result := dvAucun
  else if EtatString = 'MAJ' then
    result := dvMaj
  else if EtatString = 'ENL' then
    result := dvEnl
  else if EtatString = 'SOS' then
    result := dvSos
  else if EtatString = 'PAR' then
    result := dvPar
  else if EtatString = 'MAR' then
    result := dvMar
  else
    result := dvInconnu;
end;

procedure cWAINI.QuickSort (FDossier: HTStrings; deb, fin: Integer) ;
var
  I, J, P: Integer;
begin
  repeat
    I := deb;
    J := fin;
    P := integer (FDossier.Objects [(deb + fin) shr 1] ) ;
    repeat
      while (integer (FDossier.Objects [I] ) < P) do
        Inc (I) ;
      while (integer (FDossier.Objects [J] ) > P) do
        Dec (J) ;
      if I <= J then
      begin
        FDossier.Exchange (I, J) ;
        Inc (I) ;
        Dec (J) ;
      end;
    until I > J;
    if deb < J then QuickSort (FDossier, deb, J) ;
    deb := I;
  until I >= fin;
end;

function cWAINI.GetNumSocRef () : integer;
var
  ResultTob: Tob;
  st: string;
{$IFNDEF EAGLCLIENT}
  WAInfoIni: cWAInfoIni;
{$ENDIF}
begin
  result := 0;
  if ModeEagl then
  begin
    ResultTOB := Request (dbss_dll + '.cWAINI', 'GetNumSocRef', self, '', '') ;
    // $$$ JP 18/01/08
    try
      if GetRetour (ResultTob, st) then
        if IsNumeric (st) then result := StrToInt (st) ;
    // $$$ JP 18/01/08
    finally
      ResultTob.free;
    end;
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
    WAInfoIni := cWAInfoIni.create () ;

    // $$$ JP 18/01/08
    try
      WAInfoIni.NomDeBase := GetSocRef () ; //YCP attention compatible uniquement .ini
      result := GetNumSoc (WAInfoIni) ;
    // $$$ JP 18/01/08
    finally
      WAInfoIni.free;
    end;
{$ENDIF}
  end;
end;

function incCode(code: string; pos: integer=-1): string ;
const lc : array [0..35] of char = ('0','1','2','3','4','5','6','7','8','9','A',
'B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U',
'V','W','X','Y','Z') ;
var i : integer ;
begin
if (pos=-1) then pos:=Length(code) ;
if pos>Length(code) then
  begin
  result:=''
  end else
  begin
  i:=low(lc) ;
  while i<high(lc) do
    begin
    if (lc[i]=code[pos]) and (i<high(lc)-1) then
      begin
      code[pos]:=lc[i+1] ;
      break ;
      end;
    inc(i) ;
    end ;
  if (i=high(lc)) then
    begin
    code[pos]:=lc[0] ;
    result:=incCode(code,pos-1) ;
    end else
    begin
    result:=code ;
    end ;
  end ;
end ;

function cWAINI.AddUser (NomBase: string; User, Password: string) : boolean;
var
  tmpTob, ResultTob: tob;
{$IFNDEF EAGLCLIENT}
  WAInfoIni: cWAInfoIni;
  q: TQuery;
  sql: string;
  NumUser, st: string;
  i: integer;
  okok: boolean ;
{$ENDIF EAGLCLIENT}
begin
  result := false;
  user:=UpperCase(user) ;
  Password:=UpperCase(Password) ;
  if ModeEagl then
  begin
    tmpTOB := tob.create ('OPTIONS', nil, -1) ;
    tmpTOB.AddChampSupValeur ('NomBase', NomBase) ;
    tmpTOB.AddChampSupValeur ('User', User) ;
    tmpTOB.AddChampSupValeur ('Password', Password) ;
    AjouteUneTob (self, tmpTob) ;
    ResultTOB := Request (dbss_dll + '.cWAINI', 'AddUser', self, '', '') ;
    result := GetRetour (ResultTob) ;
    ResultTob.free;
    self.ClearDetail;
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
    WAInfoIni := cWAInfoIni.create () ;
    // $$$ JP 18/01/08
    try
     WAInfoIni.NomDeBase := NomBase;
     if Existedossier (WAInfoIni) then
     begin
      q := nil;
      try
        ChargeDBParams (WAInfoIni) ;
        getDB() ;
        if connecteDB (WAInfoIni, xxDB, 'WAINITMPSOC' + inttostr(GetCurrentThreadId) ) then
        begin
          i := pos ('&@#', user) ; //astuce pour eviter de redefinir l'interface (SUPER CRADE) !!!
          if (i > 0) then
          begin
            user := copy (user, i + 3, 255) ;
            sql := 'UPDATE PARAMSOC SET SOC_DATA="' + Password + '" WHERE SOC_NOM="' + user + '"';
            traceexecution (Sql) ;
            ExecuteSQLDB (sql, LaDB) ;
            result := true;
          end
          else
          begin
            sql := 'SELECT US_UTILISATEUR FROM UTILISAT WHERE US_ABREGE="' + user + '"';
            traceexecution (Sql) ;
            q := DBOpenSQL (LaDB, Sql, WAInfoIni.DBDriver) ;
            result := Q.Eof;
            ferme (Q) ;
            if not result then
            begin
              ErrorMessage := 'L''utilisateur ' + user + ' existe déjà ';
            end else
            begin
              sql := 'SELECT US_UTILISATEUR FROM UTILISAT order by US_UTILISATEUR';
              traceexecution (Sql) ;
              q := DBOpenSQL (LaDB, Sql, WAInfoIni.DBDriver) ;
              numUser:='000' ;  okok:=false ;
              while not okok and (numUser<>'')  do
                begin
                okok:=true ; q.first ;
                while not q.eof do
                  begin
                  st:=Q.fields [0].ASString ;
                  if st=numUser then
                    begin
                    okok:=false ;
                    break ;
                    end;
                  q.next ;
                  end ;

                if not okok then
                  begin
                  numUser:=incCode(numUser) ;
                  end
                else if (numUser<>'') then
                  begin
                  sql := 'INSERT INTO UTILISAT '
                     + ' (US_UTILISATEUR,US_LIBELLE,US_ABREGE,US_CONTROLEUR,US_PASSWORD,US_PRESENT,US_SUPERVISEUR,US_GROUPE)'
                     + 'values ("' + numuser + '","' + user + '","' + user + '","-","' + CryptageSt ( password ) + '","-","X","ADM")';
                  traceexecution (Sql) ;
                  ExecuteSQLDB (sql, LaDB) ;
                  result := true;
                  numUser:=incCode(numUser) ;
                  end ;
                end ;
              end ;
            ferme (Q) ;
          end ;
        end;
        LaDB:=nil ;
      except
        on E: Exception do
        begin
          LaDB:=nil ;
          ErrorMessage := E.Message;
        end;
      end;
     end;
    // $$$ JP 18/01/08
    finally
      FreeAndNil (WAInfoIni);
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINI.ModifUser (NomBase: string; User, Password: string) : boolean;
var
  tmpTob, ResultTob: tob;
{$IFNDEF EAGLCLIENT}
  WAInfoIni: cWAInfoIni;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  result := false;
  user:=UpperCase(user) ;
  Password:=UpperCase(Password) ;
  if ModeEagl then
  begin
    tmpTOB := tob.create ('OPTIONS', nil, -1) ;
    tmpTOB.AddChampSupValeur ('NomBase', NomBase) ;
    tmpTOB.AddChampSupValeur ('User', User) ;
    tmpTOB.AddChampSupValeur ('Password', Password) ;
    AjouteUneTob (self, tmpTob) ;
    ResultTOB := Request (dbss_dll + '.cWAINI', 'ModifUser', self, '', '') ;
    result := GetRetour (ResultTob) ;
    ResultTob.free;
    self.ClearDetail;
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
   WAInfoIni := cWAInfoIni.create () ;
   // $$$ JP 18/01/08
   try
    WAInfoIni.NomDeBase := NomBase;
    if Existedossier (WAInfoIni) then
    begin
      try
        ChargeDBParams (WAInfoIni) ;
        getDB() ;
        if connecteDB (WAInfoIni, xxDB, 'WAINITMPSOC' + inttostr(GetCurrentThreadId) ) then
        begin
          sql := 'UPDATE UTILISAT SET US_PASSWORD="'+CryptageSt (password)+'" WHERE US_ABREGE="' + user + '"' ;
          traceexecution (Sql) ;
          result:=ExecuteSQLDB (sql, LaDB)>0 ;
          if not result then
              ErrorMessage := 'L''utilisateur ' + user + ' n''existe pas';
        end;
        LaDB:=nil ;
      except
        on E: Exception do
        begin
          LaDB:=nil ;
          ErrorMessage := E.Message;
        end;
      end;
    end;
   // $$$ JP 18/01/08
   finally
    FreeAndNil (WAInfoIni);
   end;

{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINI.GetNumSoc (WAInfoIni: cWAInfoIni) : integer;
var
  ResultTob: tob;
  st: string;
{$IFNDEF EAGLCLIENT}
  q: TQuery;
  sql: string;
{$ENDIF EAGLCLIENT}
begin
  result := 0;
  if ModeEagl then
  begin
    AjouteUneTob (self, WAInfoIni) ;
    ResultTOB := Request (dbss_dll + '.cWAINI', 'GetNumSoc', self, '', '') ;
    RecupereUneTob (self, WAInfoIni.NomTable) ;
    if GetRetour (ResultTob, st) then
      if IsNumeric (st) then result := StrToInt (st) ;
    ResultTob.free;
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
    if Existedossier (WAInfoIni) then
    begin
      q := nil;
      try
        ChargeDBParams (WAInfoIni) ;
        getDB() ;
        if connecteDB (WAInfoIni, xxDB, 'WAINITMPSOC' + inttostr(GetCurrentThreadId) ) then
        begin
          sql := 'select SO_VERSIONBASE from societe';
          traceexecution (Sql) ;
          q := DBOpenSQL (LaDB, Sql, WAInfoIni.DBDriver) ;
          if not Q.eof then result := q.fields [0] .AsInteger;
          q.free;
        end;
        LaDB:=nil ;
      except
        on E: Exception do
        begin
          LaDB:=nil ;
          if Assigned (q) then q.free;
          ErrorMessage := E.Message;
        end;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINI.TestConnection (WAInfoIni: cWAInfoIni) : boolean;
var
  ResultTob: tob;
begin
  result := false;
  if ModeEagl then
  begin
    AjouteUneTob (self, WAInfoIni) ;
    ResultTOB := Request (dbss_dll + '.cWAINI', 'TestConnection', self, '', '') ;
    RecupereUneTob (self, WAInfoIni.NomTable) ;
    result := GetRetour (ResultTob) ;
    ResultTob.free;
  end
  else
  begin
{$IFNDEF EAGLCLIENT}
    try
      getDB();
      result := connecteDB (WAInfoIni.NomDeBase, xxDB, 'WAINITEST' + inttostr(GetCurrentThreadId) ) ;
      LaDB:=nil ;
    except
      on E: Exception do
      begin
        LaDB:=nil ;
        ErrorMessage := E.Message;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

function cWAINI.GetSocRef () : string;
var
  ResultTob: Tob;
  st: string;
begin
 if ModeEagl then
 begin
  ResultTOB := Request (dbss_dll + '.cWAINI', 'GetSocRef', self, '', '') ;

  // $$$ JP 18/01/08
  try
    if GetRetour (ResultTob, st) then
      result := st;
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
 end
 else
 begin
    result:='Reference' ;
 end ;
end;

function cWAINI.GetEnvironnement(WAEnvironnement : cWAEnvironnement): boolean;
var
  ResultTob, tmptob: Tob;
  //st: string;
begin
if ModeEagl then
  begin
  AjouteUneTob (self, WAEnvironnement) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'GetEnvironnement', self, '', '') ;
  //WAEnvironnement:=cWAEnvironnement(RecupereUneTob (self, WAEnvironnement.NomTable)) ;
  //result := GetRetour (ResultTob, st) ;
  // $$$ JP 18/01/08
  try
    RecupereUneTob (self, WAEnvironnement.NomTable) ;
    result:=GetRetour (ResultTob) ;
    if result then
    begin
      tmpTob := RecupereUneTob (ResultTob, cWAEnvironnement.ClassName) ;
      if (tmpTob <> nil) then
      begin
        WAEnvironnement.Dupliquer (tmpTob.detail [0] , true, true) ;
      end;
    end;
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
  end
  else
  begin
    result:=true ;
  end ;
end ;


procedure cWAINI.ChargeDBParams (WAInfoIni: cWAInfoIni) ;
var
  tmpTob, ResultTob: tob;
begin
 if ModeEagl then
 begin
  AjouteUneTob (self, WAInfoIni) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'chargeDBParam', self, '', '') ;
  // $$$ JP 18/01/08
  try
    RecupereUneTob (self, WAInfoIni.NomTable) ;
    if GetRetour (ResultTob) then
    begin
      tmpTob := RecupereUneTob (ResultTob, cWAInfoIni.ClassName) ;
      if (tmpTob <> nil) then
      begin
        WAInfoIni.Dupliquer (tmpTob.detail [0] , true, true) ;
      end;
    end;
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
 end
 else
 begin
   //do nothing
 end ;
end;

function cWAINI.getListOfHeberge(): HTStringList;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  result := HTStringList.Create();
  ResultTOB := Request(dbss_dll + '.cWAINI', 'getListOfHeberge', self, '', '');
  GetRetour(ResultTob, result);
  ResultTob.free;
  end else
  begin
  result:=HTStringList.create() ;
  end ;
end ;

function cWAINI.getListOfClient(): HTStringList;
var
  ResultTob: Tob;
begin
 if ModeEagl then
 begin
  result := HTStringList.Create();
  ResultTOB := Request(dbss_dll + '.cWAINI', 'getListOfClient', self, '', '');
  // $$$ JP 18/01/08
  try
   GetRetour(ResultTob, result);
  // $$$ JP 18/01/08
  finally
   ResultTob.free;
  end;
  end
 else
 begin
   result:=HTStringList.create() ;
 end ;
end ;

function cWAINI.GetListOfEnvironnement() : HTStringList;
var
  ResultTob: Tob;
begin
 if ModeEagl then
 begin
  result := HTStringList.Create();
  ResultTOB := Request(dbss_dll + '.cWAINI', 'GetListOfEnvironnement', self, '', '');
  // $$$ JP 18/01/08
  try
   GetRetour(ResultTob, result);
  // $$$ JP 18/01/08
  finally
   ResultTob.free;
  end;
 end
 else
 begin
  result:=HTStringList.create() ;
 end ;
end ;

function cWAINI.getListOfDossier(): HTStringList;
var
  ResultTob: Tob;
begin
 if ModeEagl then
 begin
  result := HTStringList.Create();
  ResultTOB := Request(dbss_dll + '.cWAINI', 'getListOfDossier', self, '', '');
  // $$$ JP 18/01/08
  try
    GetRetour(ResultTob, result);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
 end
 else
 begin
   result:=HTStringList.create() ;
 end ;
end ;

function cWAINI.UserPassValidate(user, pass: string): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('user', user);
  tmpTOB.AddChampSupValeur('pass', pass);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'UserPassValidate', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.IsFicBase(): boolean;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  ResultTob := Request(dbss_dll + '.cWAINI', 'IsFicBase', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.IsAdmin(login: string): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('user', login);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'IsAdmin', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail ;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.GetHeberge: TOB;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  result := TOB.Create('', nil, -1);
  ResultTob := Request(dbss_dll + '.cWAINI', 'getHeberge', self, '', '');
  GetRetour(ResultTob, result);
  ResultTob.free;
  end else
  begin
  result:=TOB.create('',nil,-1) ;
  end ;
end ;

function cWAINI.GetClient: TOB;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  result := TOB.Create('', nil, -1);
  ResultTob := Request(dbss_dll + '.cWAINI', 'GetClient', self, '', '');
  GetRetour(ResultTob, result);
  ResultTob.free;
  end else
  begin
  result:=TOB.Create('',nil,-1) ;
  end ;
end ;

function cWAINI.GetDossier: TOB;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  result := TOB.Create('', nil, -1);
  ResultTob := Request(dbss_dll + '.cWAINI', 'GetDossier', self, '', '');
  GetRetour(ResultTob, result);
  ResultTob.free;
  end else
  begin
  result:=TOB.create('',nil,-1) ;
  end ;
end ;

function cWAINI.setHeberge(laTob: TOB): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create(laTob.NomTable, nil, -1);
  tmpTob.Dupliquer(laTob, true, true);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'setHeberge', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.setClient(laTob: TOB): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create(laTob.NomTable, nil, -1);
  tmpTob.Dupliquer(laTob, true, true);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'setClient', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.setDossier(laTob: TOB; BaseFrom: string; AvecRecupHeberge: boolean): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create(laTob.NomTable, nil, -1);
  tmpTob.Dupliquer(laTob, true, true);
  AjouteUneTob(self, tmpTob);

  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('BaseFrom', BaseFrom);
  tmpTOB.AddChampSupValeur('AvecRecupHeberge', AvecRecupHeberge);
  AjouteUneTob(self, tmpTob);

  ResultTOB := Request(dbss_dll + '.cWAINI', 'setDossier', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.DelHeberge: boolean;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  ResultTob := Request(dbss_dll + '.cWAINI', 'delHeberge', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.DelClient: boolean;
var
  ResultTob: Tob;
begin
if ModeEagl then
  begin
  ResultTob := Request(dbss_dll + '.cWAINI', 'delClient', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.DelDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique: boolean): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('AvecSupprPhysique', AvecSupprPhysique);
  tmpTOB.AddChampSupValeur('ForceSupprSiEchecPhysique', ForceSupprSiEchecPhysique);
  AjouteUneTob(self, tmpTob);

  ResultTob := Request(dbss_dll + '.cWAINI', 'delDossier', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail ;

  end else
  begin
  result:=false ;
  end ;
end ;

function cWAINI.GetNextID(leQuel: integer): string;
var
  tmpTob, ResultTob: Tob;
  st: string;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('lequel', lequel);
  AjouteUneTob(self, tmpTob);

  ResultTob := Request(dbss_dll + '.cWAINI', 'GetNextID', self, '', '');
  if GetRetour(ResultTob, st) then
    result := st;
  ResultTob.free;
  self.ClearDetail ;

  end else
  begin
  result:='' ;
  end ;
end ;

function cWAINI.SetDossierVerrou(BaseNom: string; Etat: TDossierVerrou): boolean;
var
  tmpTob, ResultTob: Tob;
begin
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('BaseNom', BaseNom);
  tmpTOB.AddChampSupValeur('Etat', DossierVerrouToString(Etat));
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'SetDossierVerrou', self, '', '');
  result := GetRetour(ResultTob);
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=false ;
  end ;

end ;

function cWAINI.GetDossierVerrou(BaseNom: string): TDossierVerrou;
var
  tmpTob, ResultTob: Tob;
  st: string;
begin
result := dvinconnu;
if ModeEagl then
  begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('BaseNom', BaseNom);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWAINI', 'GetDossierVerrou', self, '', '');
  if GetRetour(ResultTob, st) then
    result := StringToDossierVerrou(st);
  ResultTob.free;
  self.ClearDetail;
  end ;
end ;



function cWAINI.getProduit: string;
var
  ResultTob: tob;
  st: string;
begin
  result := '';
if ModeEagl then
  begin
  ResultTOB := Request (dbss_dll + '.cWAINI', 'getProduit', self, '', '') ;
  if GetRetour (ResultTob, st) then result := st;
  ResultTob.free;
  end ;
end;

function cWAINI.GetMaxUserNumber: integer;
var
  ResultTob: tob;
  st: string;
begin
  result := -1;
if ModeEagl then
  begin
  ResultTOB := Request (dbss_dll + '.cWAINI', 'GetMaxUserNumber', self, '', '') ;
  if GetRetour (ResultTob, st) and isNumeric (st) then result := StrToInt (st) ;
  ResultTob.free;
  end ;
end;

function cWAINI.Ajoutedossier (WAInfoIni: cWAInfoIni) : boolean;
var
  ResultTob: tob;
begin
if ModeEagl then
  begin
  AjouteUneTob (self, WAInfoIni) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'Ajoutedossier', self, '', '') ;
  RecupereUneTob (self, WAInfoIni.NomTable) ;
  result := GetRetour (ResultTob) ;
  ResultTob.free;
  self.ClearDetail ;
  end else
  begin
  result:=false ;
  end ;
end;

function cWAINI.Supprimedossier (WAInfoIni: cWAInfoIni) : boolean;
var
  ResultTob: tob;
begin
if ModeEagl then
  begin
  AjouteUneTob (self, WAInfoIni) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'Supprimedossier', self, '', '') ;
  RecupereUneTob (self, WAInfoIni.NomTable) ;
  result := GetRetour (ResultTob) ;
  ResultTob.free;
  self.ClearDetail ;
  end else
  begin
  result:=false ;
  end ;
end;


function cWAINI.Modifiedossier (WAInfoIni: cWAInfoIni) : boolean;
var
  ResultTob: tob;
begin
if ModeEagl then
  begin
  AjouteUneTob (self, WAInfoIni) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'Modifiedossier', self, '', '') ;
  RecupereUneTob (self, WAInfoIni.NomTable) ;
  result := GetRetour (ResultTob) ;
  ResultTob.free;
  self.ClearDetail ;
  end else
  begin
  result:=false ;
  end ;
end;

function cWAINI.Existedossier (WAInfoIni: cWAInfoIni) : boolean;
var
  ResultTob: tob;
begin
if ModeEagl then
  begin
  AjouteUneTob (self, WAInfoIni) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'Existedossier', self, '', '') ;
  RecupereUneTob (self, WAInfoIni.NomTable) ;
  result := GetRetour (ResultTob) ;
  ResultTob.free;
  self.ClearDetail ;
  end else
  begin
  result:=false ;
  end ;
end;

function cWAINI.Chargedossier (VireRef, VireDP, VireModel: boolean) : HTStringList;
var
  TmpTOB, ResultTob: Tob;
begin
if ModeEAGL then
  begin
  result := HTStringList.Create;
  tmpTOB := tob.create ('OPTIONS', nil, -1) ;
  tmpTOB.AddChampSupValeur ('Viref', VireRef) ;
  tmpTOB.AddChampSupValeur ('VireDP', VireDP) ;
  tmpTOB.AddChampSupValeur ('VireModel', VireModel) ;
  AjouteUneTob (self, tmpTob) ;
  ResultTOB := Request (dbss_dll + '.cWAINI', 'Chargedossier', self, '', '') ;
  GetRetour (ResultTob, result) ;
  ResultTob.free;
  self.ClearDetail;
  end else
  begin
  result:=HTStringList.create() ;
  end ;
end;

{$IFNDEF EAGLCLIENT}
class procedure cWAINI.Deconnectedb (wDB: TDataBase) ;
begin
  //redirigé vers TCbpDataBases
  if assigned(wDB) and (wDB.Connected) then
    begin
    TCBPDatabases.DisConnectDatabase(wDB) ;
    //nbConnect := nbConnect - 1;
    TraceExecution ('=======  cWAINI.Deconnectedb ' + ' (' + wDB.DatabaseName + '/'+ wDB.Session.SessionName + ')' ) ;
    end;
end;
//YCP 15/05/07 début

class function cWAINI.connectedb (WAInfoIni: cWAInfoIni; var wDB: TDataBase; stNom: string) : boolean;
begin
result:=false ;
if wDB<>nil then
  begin
  try
    if (trim(stNom)='') and (trim(wDB.Databasename)='') then
      stNom := 'TEMPO' + IntToStr(GetCurrentThreadId)+inttostr(GetTickCount)
      else stNom := '' ;
    AssignDBParamsStrings (stNom, WAInfoIni.Driver, WAInfoIni.server, WAInfoIni.Path,
      WAInfoIni.base, WAInfoIni.user, WAInfoIni.pass, WAInfoIni.odbc,
      WAInfoIni.options, wDB) ;
    //YCP wDB.Connected := False;
    wDB.EnablePooling:=false ;    //YCP Pour ne pas mettre la base dans le pool de connexion... pb de suppression
    //if (trim(stNom)<>'') or (trim(wDB.DatabaseName)='') then
       //wDB.DatabaseName := stNom+ IntToStr(GetCurrentThreadId)+inttostr(GetTickCount) ; //24/05/2007
    result := TCBPDatabases.ConnectDatabase(wDB) ;
    //nbConnect := nbConnect + 1;
    TraceExecution ('======= cWAINI.Connectedb ' + ' (' + wDB.DatabaseName + '/'+ wDB.Session.SessionName +')' ) ;
  except
    on e: Exception do
        {$IFDEF EAGLSERVER}
        LogeAGL('Exception lors du connectedb : ' + e.message, TRUE, FALSE);
        {$ELSE}
        if V_PGI.SAV then PGIBox('Exception lors du connectedb : ' + e.message);
        {$ENDIF EAGLSERVER}
    end;
  end ;
end;

class function cWAINI.connectedb (WAInfoIni: cWAInfoIni; var wDB: TDataBase) : boolean;
begin
result:=connectedb(WAInfoIni,wDB,wDB.dataBaseName) ;
end ;

function cWAINI.connectedb (NomBase: string; var wDB: TDataBase; stNom: string) : boolean;
var
  WAInfoIni: cWAInfoIni;
begin
  result := false;
  try
    traceexecution ('Connection à :' + NomBase) ;
    WAInfoIni := cWAInfoIni.create;
    try
      WAInfoIni.NomDeBase := NomBase;
      chargeDBParams (WAInfoIni) ;
      result:=connectedb (WAInfoIni, wDB, stNom) ;
    finally
      WAInfoIni.free;
      end ;
    //wDB.Connected := true;
    //result := wDB.Connected;
  except
    on E: Exception do ErrorMessage := E.Message;
  end;
end;

function cWAINI.connectedb (NomBase: string; var wDB: TDataBase) : boolean;
begin
result:=connectedb(NomBase,wDB,wDB.DatabaseName);
end ;
//YCP 15/05/07 
{$ENDIF EAGLCLIENT}

class function cWAINI.GetUserNumberFromSeria (GroupName: string) : integer;
{$IFDEF EAGLCLIENT}
var
  WAini: cWAINI;
begin
  WAini := cWAINI.create;
  try
    WAIni.IDClient := GroupName;
    result := WAIni.GetMaxUserNumber;
    WAIni.free;
  except
    on E: Exception do
    begin
      result := -1;
      if assigned (WAini) then WAini.Free;

    end;
  end;
{$ELSE}
begin
  Result := -1;
{$ENDIF EAGLCLIENT}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Fonction dbss sur le server
Suite ........ :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
class function cWAINI.Action (UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB) : boolean;
var
  T1,
  tmpTob, laTob   : tob;
  Viref, VireDP,
  VireModel        : boolean;
  tt               : HTStringList;
  WAIni            : cWAINI;
  WAInfoIni        : cWAInfoIni;
  WAEnvironnement  : cWAEnvironnement;
  NomBase, User,
  Password, wSt,
  BaseNom, BaseFrom,
  Etat       : string;
  ForceSupprSiEchecPhysique,
  AvecSupprPhysique,
  AvecRecupHeberge : boolean;
  lequel           : integer;
begin
  result := true;
  Viref := false;
  VireDP := false;
  VireModel := false;
  WAInfoIni := nil;
  WAEnvironnement := nil ;
  T1 := nil;
  tt := nil;
  wSt := '';
  lequel := -1;
  AvecSupprPhysique := false;
  ForceSupprSiEchecPhysique := false;
  AvecRecupHeberge := false;
  WAIni := nil;
  BaseNom := '' ; BaseFrom := '' ; User := '' ; Password := '' ; wSt:='' ; nomBase:=''; //YCP 15/05/07
  Etat := '' ;
  TraceExecution ('dbssPGI.cWAINI.' + UneAction + ' :') ;
  try
    UneAction := UpperCase (UneAction) ;
    tmpTob := RecupereUneTob (RequestTOB, 'OPTIONS') ;
    if (tmpTob <> nil) then
    begin
      if tmpTob.detail [0] .FieldExists ('Viref')                     then Viref := tmpTob.detail [0] .GetValue ('Viref') ;
      if tmpTob.detail [0] .FieldExists ('VireDP')                    then VireDP := tmpTob.detail [0] .GetValue ('VireDP') ;
      if tmpTob.detail [0] .FieldExists ('VireModel')                 then VireModel := tmpTob.detail [0] .GetValue ('VireModel') ;
      if tmpTob.detail [0] .FieldExists ('NomBase')                   then NomBase := tmpTob.detail [0] .GetValue ('NomBase') ;
      if tmpTob.detail [0] .FieldExists ('User')                      then User := tmpTob.detail [0] .GetValue ('User') ;
      if tmpTob.detail [0] .FieldExists ('Password')                  then Password := tmpTob.detail [0] .GetValue ('Password') ;
      if tmpTob.detail [0] .FieldExists ('BaseNom')                   then BaseNom := tmpTob.detail[0].GetValue('BaseNom');
      if tmpTob.detail [0] .FieldExists ('BaseFrom')                  then BaseFrom := tmpTob.detail[0].GetValue('BaseFrom');
      if tmpTob.detail [0] .FieldExists ('AvecRecupHeberge')          then AvecRecupHeberge := tmpTob.detail[0].GetValue('AvecRecupHeberge');
      if tmpTob.detail [0] .FieldExists ('AvecSupprPhysique')         then AvecSupprPhysique := tmpTob.detail[0].GetValue('AvecSupprPhysique');
      if tmpTob.detail [0] .FieldExists ('ForceSupprSiEchecPhysique') then ForceSupprSiEchecPhysique := tmpTob.detail[0].GetValue('ForceSupprSiEchecPhysique');
      if tmpTob.detail [0] .FieldExists ('lequel')                    then lequel := tmpTob.detail[0].GetValue('lequel');
      if tmpTob.detail [0] .FieldExists ('Etat')                      then Etat := tmpTob.detail[0].GetValue('Etat');
      tmpTob.free;
    end;
    tmpTob := RecupereUneTob (RequestTOB, cWAInfoIni.ClassName) ;
    if (tmpTob <> nil) then
    begin
      WAInfoIni := cWAInfoIni (tmpTob.detail [0] ) .clone;
      tmpTob.free;
    end;

    tmpTob := RecupereUneTob (RequestTOB, cWAEnvironnement.ClassName) ;
    if (tmpTob <> nil) then
    begin
      WAEnvironnement := cWAEnvironnement (tmpTob.detail [0] ) .clone;
      tmpTob.free;
    end;


    laTob := RecupereUneTob(RequestTOB, 'XXHEBERGE');
    if laTob = nil then
      laTob := RecupereUneTob(RequestTOB, 'XXCLIENT');
    if laTob = nil then
      laTob := RecupereUneTob(RequestTOB, 'XXSOCIETE');


    WAIni := cWAINI (RequestTOB) .clone;

    with WAIni do
    begin
      afficheTrace (cWAINI (RequestTOB) ) ;
      TraceExecution (classname + '.Viref                     :' + si (viref, 'true', 'false') ) ;
      TraceExecution (classname + '.VireDP                    :' + si (vireDP, 'true', 'false') ) ;
      TraceExecution (classname + '.VireModel                 :' + si (VireModel, 'true', 'false') ) ;
      TraceExecution (classname + '.User                      :' + User);
      TraceExecution (classname + '.Pass                      :' + Password);
      TraceExecution (classname + '.BaseFrom                  :' + BaseFrom);
      TraceExecution (classname + '.AvecRecupHeberge          :' + si(AvecRecupHeberge, 'true', 'false'));
      TraceExecution (classname + '.AvecSupprPhysique         :' + si(AvecSupprPhysique, 'true', 'false'));
      TraceExecution (classname + '.ForceSupprSiEchecPhysique :' + si(ForceSupprSiEchecPhysique, 'true', 'false'));
      TraceExecution (classname + '.lequel:' + intToStr(lequel));

      cWAInfoIni.afficheTrace (WAInfoIni) ;
      try
        if UneAction = 'CHARGEDOSSIER' then
          tt := Chargedossier (Viref, VireDP, VireModel)
        else if UneAction = 'GETSOCREF' then
          wSt := GetSocRef ()
        else if UneAction = 'GETNUMSOCREF' then
          wSt := IntToStr (GetNumSocRef () )
        else if UneAction = 'GETMAXUSERNUMBER' then
          wSt := IntToStr (GetMaxUserNumber () )
        else if UneAction = 'GETPRODUIT' then
          wSt := getProduit ()
        else if UneAction = 'GETNUMSOC' then
          wSt := IntToStr (GetNumSoc (WAInfoIni) )
        else if UneAction = 'TESTCONNECTION' then
          result := TestConnection (WAInfoIni)
        else if UneAction = 'CHARGEDBPARAM' then
          ChargeDBParams (WAInfoIni)
        else if UneAction = 'AJOUTEDOSSIER' then
          result := Ajoutedossier (WAInfoIni)
        else if UneAction = 'SUPPRIMEDOSSIER' then
          result := Supprimedossier (WAInfoIni)
        else if UneAction = 'MODIFIEDOSSIER' then
          result := Modifiedossier (WAInfoIni)
        else if UneAction = 'EXISTEDOSSIER' then
          result := Existedossier (WAInfoIni)
        else if UneAction = 'ADDUSER' then
          result := AddUser (NomBase, User, Password)
        else if UneAction = 'MODIFUSER' then
          result := ModifUser (NomBase, User, Password)
        else if (UneAction = 'GETLISTOFHEBERGE') then
          tt := getListOfHeberge
        else if (UneAction = 'GETLISTOFCLIENT') then
          tt := getListOfClient()
        else if (UneAction = 'GETLISTOFDOSSIER') then
          tt := getListOfDossier()
        else if (UneAction = 'USERPASSVALIDATE') then
          result := UserPassValidate(user, password)
        else if (UneAction = 'ISADMIN') then
          result := isadmin(user)
        else if (UneAction = 'GETHEBERGE') then
          T1 := GetHeberge()
        else if (UneAction = 'GETCLIENT') then
          T1 := getClient()
        else if (UneAction = 'GETDOSSIER') then
          T1 := GetDossier()
        else if (UneAction = 'SETHEBERGE') then
          result := setHeberge(laTob.detail[0])
        else if (UneAction = 'SETCLIENT') then
          result := setClient(laTob.detail[0])
        else if (UneAction = 'SETDOSSIER') then
          result := setDossier(laTob.detail[0], BaseFrom, AvecRecupHeberge)
        else if (UneAction = 'DELHEBERGE') then
          result := delHeberge()
        else if (UneAction = 'DELCLIENT') then
          result := delClient()
        else if (UneAction = 'DELDOSSIER') then
          result := delDossier(AvecSupprPhysique, ForceSupprSiEchecPhysique)
        else if (UneAction = 'GETNEXTID') then
          wst := getnextID(lequel)
        else if (UneAction = 'GETDOSSIERVERROU') then
          wst := DossierVerrouToString(GetDossierVerrou(BaseNom))
        else if (UneAction = 'SETDOSSIERVERROU') then
          result := SetDossierVerrou(BaseNom, StringToDossierVerrou(Etat))
        else if (UneAction = 'GETENVIRONNEMENT') then
          result := getEnvironnement(WAEnvironnement)
        else if (uneAction = 'GETLISTOFENVIRONNEMENT') then
          tt := getListOfEnvironnement()
        else if (uneAction = 'ISFICBASE') then
          result := IsFicBase()
        else
        begin
          result := false;
          ErrorMessage := 'Action inconnue : ' + UneAction;
        end;
        if ResponseTOB = nil then
          ResponseTOB := RequestTOB
        else
          ResponseTOB.Dupliquer (RequestTOB, true, true) ;
        if WAInfoIni <> nil then AjouteUneTob (ResponseTOB, WAInfoIni) ;
        if WAEnvironnement <> nil then AjouteUneTob (ResponseTOB, WAEnvironnement) ;
        setRetour (ResponseTOB, result, T1, tt, wSt) ;
      except
        on E: Exception do
          ErrorMessage := E.Message;
      end;
    end;
  finally
    if (WAIni <> nil) then WAIni.free;
    if T1 <> nil then T1.free;
    if tt <> nil then tt.free;
  end;
end;
{$ENDIF EAGLCLIENT}

end.

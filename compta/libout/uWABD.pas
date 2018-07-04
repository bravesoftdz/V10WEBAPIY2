{***********UNITE*************************************************
Auteur  ...... : Y. PELUD
Créé le ...... : 11/10/2004
Modifié le ... :   /  /    
Description .. : XP 11-10-2004 : CTRL-D
Mots clefs ... : 
*****************************************************************}
unit uWABD;
{$IFDEF BASEEXT}azertyuiop{$ENDIF BASEEXT}

interface
uses
  classes, uHTTP, utob, hctrls, stdctrls, uCptWA, uWAIni, hent1
{$IFNDEF EAGLCLIENT}
, uDbxDataSet
{$ENDIF EAGLCLIENT}
  ;

type
  cWABD = class(cCptWA)
  private
    //xxioWAIni: coWAIni ;
    //function getioWAIni: coWAIni ;
    function getDriver: string;
    function getBase: string;
    function getBaseModel: string;
    function getfichier: string;
    function getBackUpFile: string;
    //procedure setioWAIni(valeur: coWAIni) ;
    procedure setDriver(valeur: string);
    procedure setBase(valeur: string);
    procedure setBaseModel(valeur: string);
    procedure setfichier(valeur: string);
    procedure setBackUpFile(valeur: string);
  public
    constructor create(); overload; override;
    class function create(wdriver: string): cWABD; reintroduce; overload;
    destructor destroy; override;
    function clone(): cWABD;
{$IFDEF EAGLCLIENT}
    function CreerBase(): boolean; overload;
    function CreerBaseFrom(nomBase: string): boolean; overload;
    function SupprimerBase(): boolean;
    function ArchiverBase(): boolean;
    function RestorerBase(): boolean;
    function AttacherBase(): boolean;
    function DetacherBase(): boolean;
    function ExisteBase(): boolean;
    function TestConnexion(): boolean;
{$ELSE}
    function CreerBase(): boolean; overload; virtual; abstract;
    function CreerBaseFrom(nomBase: string): boolean; overload; virtual; abstract;
    function SupprimerBase(): boolean; virtual; abstract;
    function ArchiverBase(): boolean; virtual; abstract;
    function RestorerBase(): boolean; virtual; abstract;
    function AttacherBase(): boolean; virtual; abstract;
    function DetacherBase(): boolean; virtual; abstract;
    function ExisteBase(): boolean; virtual; abstract;
    function TestConnexion(): boolean; virtual; abstract;
{$ENDIF EAGLCLIENT}
{$IFDEF EAGLSERVER}
    class function Action(UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB): boolean;
{$ENDIF EAGLSERVER}
    class function isMSSQL(Driver: String): boolean ; overload ;
    class function isMSSQL(Driver: TDBDriver): boolean ; overload ;
  protected
  published
    property Driver: string read getDriver write setDriver;
    property Base: string read getBase write setBase;
    property BaseModel: string read getBaseModel write setBaseModel;
    property fichier: string read getfichier write setfichier;
    property BackUpFile: string read getBackUpFile write setBackUpFile;
  end;


implementation
uses
  uWABDSQLSERVER, uWABDMSACCESS
{$IFNDEF EAGLCLIENT}
  , uWAINIBase, uWAINIFile, sysutils, forms
{$ENDIF EAGLCLIENT}
  ;
constructor cWABD.create();
begin
  inherited create;
  ReaffecteNomTable(className, true);
  //ioWAIni:=nil ;
end;

destructor cWABD.destroy;
begin
  //ioWAIni:=nil ;
  inherited;
end;

function cWABD.clone(): cWABD;
begin
{$IFDEF EAGLCLIENT}
  result := cWABD.create;
{$ELSE}
  result := cWABD.create(driver);
{$ENDIF EAGLCLIENT}
  if result <> nil then result.Dupliquer(self, true, true);
end;

(*
**
** MSACCESS
** INTRBASE
** ODBC_MSSQL
** ORACLE
** ORACLE8
** MSSQL
** ODBC_SYBASE
** ODBC_DB2
** INFORMIX
** STANDARD
** ODBC_SQLANYWHERE
** SQLBASE
** ODBC_POL
** ODBC_MSACCESS
** ODBC_MYSQL
** ODBC_INTRBASE
** ODBC_PROGRESS
** ODBC_ORACLE8
** DB2
** SYBASE
** ODBC_ASIQ
**
*)

class function cWABD.create(wdriver: string): cWABD;
begin
  if (wdriver = 'ODBC_MSSQL') then result := cWABDSQLSERVER.Create
  //else if (wdriver = 'MSACCESS') then result := cWABDMSACCESS.Create
  else result := cWABDSQLSERVER.Create;
end;


{$IFDEF EAGLCLIENT}
function cWABD.CreerBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'CreerBase', self, '', '');

  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.CreerBaseFrom(nomBase: string): boolean;
var
  tmpTOB, ResultTob: Tob;
begin
  tmpTOB := tob.create('OPTIONS', nil, -1);
  tmpTOB.AddChampSupValeur('nomBase', nomBase);
  AjouteUneTob(self, tmpTob);
  ResultTOB := Request(dbss_dll + '.cWABD', 'CreerBaseFrom', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.SupprimerBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'SupprimerBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.ArchiverBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'ArchiverBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.RestorerBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'RestorerBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.AttacherBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'AttacherBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.DetacherBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'DetacherBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;

function cWABD.ExisteBase(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := Request(dbss_dll + '.cWABD', 'ExisteBase', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;


function cWABD.TestConnexion(): boolean;
var
  ResultTob: Tob;
begin
  ResultTOB := request(dbss_dll + '.cWABD', 'TestConnexion', self, '', '');
  // $$$ JP 18/01/08
  try
    result := GetRetour(ResultTob);
  // $$$ JP 18/01/08
  finally
    ResultTob.free;
  end;
end;
{$ENDIF EAGLCLIENT}

(*function cWABD.getioWAIni: coWAIni ;
var
 i: integer ;
begin
if (xxioWAIni=nil) then
  begin
  for i:=0 to self.Detail.Count-1 do
    begin
    {$IFDEF EAGLCLIENT}
    xxioWAIni:=coWAINI(self.Detail[i]) ;
    {$ELSE}
    if (self.Detail[i].NomTable=coWAINIBase.ClassName) then xxioWAIni:=coWAINIBase(self.Detail[i]) else
    if (self.Detail[i].NomTable=coWAINIFile.ClassName) then xxioWAIni:=coWAINIFile(self.Detail[i]) ;
    {$ENDIF EAGLCLIENT}
    if (xxioWAIni<>nil) then break ;
    end ;
  if (xxioWAIni=nil) then
    begin
    {$IFDEF EAGLCLIENT}
    xxioWAIni:=coWAINI.create ;
    {$ELSE}
    xxioWAIni:=coWAINI.create('') ;
    {$ENDIF EAGLCLIENT}
    xxioWAIni.ChangeParent(self,0) ;
    end ;
  end ;
result:=xxioWAIni ;
end ;

procedure cWABD.setioWAIni(valeur: coWAIni) ;
begin
if (xxioWAIni<>nil) then
  begin
  xxioWAIni.ChangeParent(nil,-1) ;
  xxioWAIni.free ;
  end ;
if (valeur<>nil) then
  begin
  valeur.ReaffecteNomTable(valeur.ClassName,true) ;
  valeur.ChangeParent(self,0) ;
  end ;
xxioWAIni:=valeur ;
end ;*)

function cWABD.getDriver: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('Driver')) then self.AddChampSupValeur('Driver', '');
  result := self.getString('Driver');
end;

procedure cWABD.setDriver(valeur: string);
begin
  if (not self.FieldExists('Driver')) then self.AddChampSup('Driver', false);
  self.PutValue('Driver', valeur);
end;

function cWABD.getBase: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('Base')) then self.AddChampSupValeur('Base', '');
  result := self.getString('Base');
end;

procedure cWABD.setBase(valeur: string);
begin
  if (not self.FieldExists('Base')) then self.AddChampSup('Base', false);
  self.PutValue('Base', valeur);
end;

function cWABD.getBaseModel: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('BaseModel')) then self.AddChampSupValeur('BaseModel', '');
  result := self.getString('BaseModel');
end;

procedure cWABD.setBaseModel(valeur: string);
begin
  if (not self.FieldExists('BaseModel')) then self.AddChampSup('BaseModel', false);
  self.PutValue('BaseModel', valeur);
end;

function cWABD.getfichier: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('fichier')) then self.AddChampSupValeur('fichier', '');
  result := self.getString('fichier');
end;

procedure cWABD.setfichier(valeur: string);
begin
  if (not self.FieldExists('fichier')) then self.AddChampSup('fichier', false);
  self.PutValue('fichier', valeur);
end;

function cWABD.getBackupFile: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('BackupFile')) then self.AddChampSupValeur('BackupFile', '');
  result := self.getString('BackupFile');
end;

procedure cWABD.setBackupFile(valeur: string);
begin
  if (not self.FieldExists('BackupFile')) then self.AddChampSup('BackupFile', false);
  self.PutValue('BackupFile', valeur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Fonction dbss sur le server
Suite ........ :
Mots clefs ... :
*****************************************************************}
{$IFDEF EAGLSERVER}
class function cWABD.Action(UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB): boolean;
var
  tmpTOB: tob;
  ioWABD: cWABD;
  NomBase: string;
begin
  result := true;
  ioWABD := nil;
  TraceExecution('dbssPGI.cWABD.' + UneAction + ' :');
  try
    ioWABD := cWABD(RequestTOB).clone;
    ;
    UneAction := UpperCase(UneAction);
    tmpTob := RecupereUneTob(RequestTOB, 'OPTIONS');
    if (tmpTob <> nil) then
    begin
      if tmpTob.detail[0].FieldExists('NomBase') then NomBase := tmpTob.detail[0].GetValue('NomBase');
      tmpTob.free;
    end;

    with ioWABD do
    begin
      TraceExecution('*=>BaseModel  : ' + ioWABD.BaseModel);
      TraceExecution('*=>fichier    : ' + ioWABD.fichier);
      TraceExecution('*=>BackUpFile : ' + ioWABD.BackUpFile);
      TraceExecution('*=>NomBase    : ' + NomBase);

      try
        if UneAction = 'CREERBASE' then result := CreerBase()
        else if UneAction = 'CREERBASEFROM' then result := CreerBaseFrom(nomBase)
        else if UneAction = 'SUPPRIMERBASE' then result := SupprimerBase()
        else if UneAction = 'ARCHIVERBASE' then result := ArchiverBase()
        else if UneAction = 'RESTORERBASE' then result := RestorerBase()
        else if UneAction = 'ATTACHERBASE' then result := AttacherBase()
        else if UneAction = 'DETACHERBASE' then result := DetacherBase()
        else if UneAction = 'EXISTEBASE' then result := ExisteBase()
        else if UneAction = 'TESTCONNEXION' then result := TestConnexion()
        else
        begin
          result := false;
          ErrorMessage := 'Action inconnue : ' + UneAction;
        end;
      except
        on E: Exception do
          ErrorMessage := E.Message;
      end;
      if ResponseTOB = nil then ResponseTOB := TOB.Create('', nil, -1);
      ResponseTOB.Dupliquer(ioWABD, true, true);
      setRetour(ResponseTOB, result, nil, nil, '');
    end;
  finally
    if ioWABD <> nil then ioWABD.free;
  end;
end;
{$ENDIF EAGLSERVER}



class function cWABD.isMSSQL(Driver: String): boolean ;
begin
result:=isMSSQL(StToDriver(Driver,false)) ;
end ;

class function cWABD.isMSSQL(Driver: TDBDriver): boolean ;
begin
result:= (Driver in [dbMSSQL, dbMSSQL2005, dbMSSQL2008]) ;
end ;

end.


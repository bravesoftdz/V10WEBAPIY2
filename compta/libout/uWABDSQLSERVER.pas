{***********UNITE*************************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 17/12/2002
Modifié le ... :   /  /
Description .. : Contient les fonctions qui permettent de gérer un server
Suite ........ : SQL (local). Ainsi que les déclarations de ces fonctions
Suite ........ : executablent depuis les scripts
Suite ........ : XP 18-10-2004 : CTRL + D
Mots clefs ... : S1;DATABASE;
*****************************************************************}
unit uWABDSQLSERVER;
{$IFDEF BASEEXT}azertyuiop{$ENDIF BASEEXT}


interface
uses
  utob, Classes, HenT1, MajTable, hctrls, uHttp, stdctrls
    {$IFNDEF EAGLCLIENT}
    ,db ,uDbxDataSet
    {$ENDIF EAGLCLIENT}
  , uwaBD, uWAINI, uWAInfoIni;

type
  cWABDSQLSERVER = class(cWABD)
  private
{$IFNDEF EAGLCLIENT}
    xxForceBackupDirectory: string ;
    xxDB: TDatabase;
    xxDumpFile: boolean ;
    xxDumpName: string ;
    function getDB: TDatabase;
    procedure setDB(valeur: TDatabase);
    //function GetConnecte: boolean ;
    //procedure SetConnecte(Value: boolean) ;
    function ExecuteDBSpec(wSQL: string): boolean;
    function LoadDetailFromSQLSpec(wSql: string): TOB;
    property DB: TDatabase read getDB ; //write setDB;
{$ENDIF EAGLCLIENT}
    class function GetLienODBC(): string;
{$IFNDEF EAGLCLIENT}
    //property Connecte: boolean read GetConnecte write SetConnecte ;
{$ENDIF}
    function getMasterServer: string;
    function getMasterUser: string;
    function getMasterPass: string;
    function getmdfFile: string;
    function getldfFile: string;
    function getFileGrowth: string;
    function getJobName: string;
    function getTypeFreq: string;
    function getDateDebut: string;
    function getHeureDebut: string;
    function getInterVal: string;
    function getjob_id: string;
    procedure setMasterServer(valeur: string);
    procedure setMasterUser(valeur: string);
    procedure setMasterPass(valeur: string);
    procedure setmdfFile(valeur: string);
    procedure setldfFile(valeur: string);
    procedure setFileGrowth(valeur: string);
    procedure setJobName(valeur: string);
    procedure setTypeFreq(valeur: string);
    procedure setDateDebut(valeur: string);
    procedure setHeureDebut(valeur: string);
    procedure setInterVal(valeur: string);
    procedure setjob_id(valeur: string);

  public
    constructor Create; override;
    destructor Destroy; override;
    function ReparerDatabase(): boolean;
    function AjouterLogin(UserName, Password: string): boolean;
    function ModifierLoginCEGIDSA: boolean;
    function ListerDatabase(All: boolean): TOB;
    function ListerUser(): TOB;
    function ListerJob(): TOB;
    function ListerTables(): TOB;
    function AjouterJob(): boolean;
    function SupprimerJob(): boolean;
    function GetInfosJob(): TOB;
    function new_backup(nomBase: string): TOB  ;
    function NEW_RESTORE(nomBase,filename: string): TOB  ;


{$IFNDEF BASEEXT}
    class procedure GetInfoConnection(Servername, user, password: THEdit); overload;
    class procedure GetInfoConnection(var Servername, user, password: string); overload;
    class procedure SetInfoConnection(Servername, user, password: string);
{$ENDIF BASEEXT}
{$IFNDEF EAGLCLIENT}
    function GetPhysicalName(lequel: integer): string;
    function CreerBase(): boolean; overload; override;
    function CreerBaseFrom(nomBase: string): boolean; overload; override;
    function SupprimerBase(): boolean; override;
    function ArchiverBase(): boolean; override;
    function RestorerBase(): boolean; override;
    function AttacherBase(): boolean; override;
    function DetacherBase(): boolean; override;
    function ExisteBase(): boolean; override;
    function TestConnexion(): boolean; override;

{$ENDIF EAGLCLIENT}
{$IFNDEF EAGLCLIENT}
    class function Action(UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB): boolean;
{$ENDIF EAGLCLIENT}
  published
    property MasterUser: string read getMasterUser write setMasterUser;
    property MasterPass: string read getMasterPass write setMasterPass;
    property MasterServer: string read getMasterServer write setMasterServer;
    property mdfFile: string read getmdfFile write setmdfFile;
    property ldfFile: string read getldfFile write setldfFile;
    property FileGrowth: string read getFileGrowth write setFileGrowth;
    property JobName: string read getJobName write setJobName;
    property TypeFreq: string read getTypeFreq write setTypeFreq;
    property DateDebut: string read getDateDebut write setDateDebut;
    property HeureDebut: string read getHeureDebut write setHeureDebut;
    property InterVal: string read getInterVal write setInterVal;
    property job_id: string read getjob_id write setjob_id;

{$IFNDEF EAGLCLIENT}
   property ForceBackupDirectory: string read xxForceBackupDirectory write xxForceBackupDirectory;
{$ENDIF EAGLCLIENT}
  public
    function AutocloseBase(bOption : Boolean): boolean;
{$IFNDEF EAGLCLIENT}

  public
    function isSQL2005(): boolean ;
    function isSQL2008(): boolean ;
    function getVersion(): string ;

{$ENDIF}
  end;

implementation

uses
  Windows, SysUtils, Forms, controls, (* FileCtrl, *)
  hmsgbox, M3fp, registry, inifiles, licutil, cbptrace
{$IFNDEF BASEEXT}
  , ParamSoc
{$ENDIF BASEEXT}
  , cbpDatabases
{$IFDEF EAGLSERVER}
  , esession
{$ENDIF EAGLSERVER}
  ;

//var cpte: integer = 0; //YCP 15/05/07 

{$IFNDEF EAGLCLIENT}
function cWABDSQLSERVER.CreerBase(): boolean;
var Directory,St : String ;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
    result := ExecuteDBSpec('use master;');

  if result then
  begin
    if ldfFile = '' then ldfFile := ChangeFileExt(mdfFile, '.ldf');
    if FileGrowth = '' then FileGrowth := '10%';
    (* YCP 17/12/2004 Le serveur SQL n'est pas accessible
    if not DirectoryExists(ExtractFilePath(mdfFile)) then ForceDirectories(ExtractFilePath(mdfFile));
    if not DirectoryExists(ExtractFilePath(mdfFile)) then
    begin
      result := false;
      ErrorMessage := 'La création du répertoire de destination a échoué ' + ExtractFilePath(mdfFile);
    end else*)
    //24/05/2007 Assurer le répertoitre
    DirectoRy:=ExtractFilePath(mdfFile) ;
    st:='EXEC xp_cmdshell ''mkdir '+Directory+' '' ' ;
    result:=ExecuteDBSpec(st) ;
    if result then
    //24/05/2007 Assurer le répertoitre fin
    begin
      result := ExecuteDBSpec('CREATE DATABASE ' + Base
        + ' ON ( NAME = ''' + Base + '_data'', FILENAME = ''' + mdfFile + ''', '
        + '      MAXSIZE = UNLIMITED, FILEGROWTH = ' + FileGrowth + ' )'
        + ' LOG ON ( NAME = ''' + Base + '_log'', FILENAME = ''' + ldfFile + ''', '
        + '      MAXSIZE = UNLIMITED, FILEGROWTH = ' + FileGrowth + ' );');

      if result then result := ReparerDatabase();
      //if result then
        //result := ExecuteDBSpec('use master;');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''auto update statistics'', ''true'';');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''auto create statistics'', ''true'';');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''autoclose'', ''true'';');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''autoshrink'', ''true'';');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''select into/bulkcopy'', ''true'';');
      if result then result := ExecuteDBSpec('exec sp_dboption ' + Base + ', ''trunc. log on chkpt.'', ''true'';');

      if result then ActionMessage := 'La base a été créée avec succès'
      else ErrorMessage := 'La création de la base a échoué'; // XP 24.09.2007 FQ 14296
    end;
  end;
end;

function cWABDSQLSERVER.getVersion(): string ;
var
  t1: tob;
begin
result:='' ;
T1 := LoadDetailFromSQLSpec('SELECT CAST(@@VERSION as VarChar(30)) AS CAST_VERS');
try
if (T1 <> nil) and  (T1.detail.count>0) then
  result:=T1.detail[0].GetValue('CAST_VERS') ;
finally
  T1.free ;
  end ;
end ;

function cWABDSQLSERVER.isSQL2005(): boolean ;
var
  res: string ;
  p: integer ;
begin
  res:=getVersion() ;
  p := pos(' 2005 ', res);
  if p = 0 then
    p := pos(' 2008 ', res);
  result := p > 0 ;
end ;

// md 20.03.2009 suite FQ 16051
function cWABDSQLSERVER.isSQL2008(): boolean ;
var
  res: string ;
  p: integer ;
begin
  res:=getVersion() ;
  p := pos(' 2008 ', res);
  result := p > 0 ;
end ;


function cWABDSQLSERVER.CreerBaseFrom(nomBase: string): boolean;
var
  Directory,St : String ;
begin
  //YCP 06/01/05 result := false;
  if nomBase = '' then
  begin
    result := CreerBase
  end else
  begin
    //YCP 06/09/2007 if result then
    begin
    if ldfFile = '' then ldfFile := ChangeFileExt(mdfFile, '.ldf');
    (* YCP 17/12/2004 le serveur SQL n'est pas accessible
    if not DirectoryExists(ExtractFilePath(mdfFile)) then ForceDirectories(ExtractFilePath(mdfFile));
    if not DirectoryExists(ExtractFilePath(mdfFile)) then
    begin
      ErrorMessage := 'La création du répertoire de destination a échoué'
    end else*)
      DirectoRy:=ExtractFilePath(mdfFile) ;
      st:='EXEC xp_cmdshell ''mkdir '+Directory+' '' ' ;
      result:=ExecuteDBSpec(st) ;
    if result then
      begin
      if not isSQL2008 then
        result:=ExecuteDBSpec('BACKUP LOG '+nomBase+' WITH NO_LOG;') ; //YCP 06/01/05
     
      if result then
        begin
        xxDumpFile:=true ;
        xxDumpName:=FormatDateTime('"DATA_"yymmdd"_"hhnnsszzz""',now()) ;
        BackUpFile:=FormatDateTime('"DATA_"yymmdd"_"hhnnsszzz".BAK"',now()) ;
        result:=ExecuteDBSpec('sp_addumpdevice ''disk'', '''+xxDumpName+''', '''+BackUpFile+'''') ;
        try
          if result then
            begin
            result:=ExecuteDBSpec('BACKUP DATABASE '+nomBase+' TO '+xxDumpName+' WITH  INIT;') ;
            if result then result:=RestorerBase() ; //YCP 06/01/05
            end ;
        finally
          ExecuteDBSpec('sp_dropdevice '''+xxDumpName+''', DELFILE') ;
          xxDumpFile:=false ;
          end ;
        end ;
      end ;

    if result then ActionMessage := 'La base a été créée avec succès'
    else ErrorMessage := 'La création de la base a échoué'; // XP 19.09.2007 FQ 14296
    end ;
  end ;
end ;
    (*YCP 17/12/2004

      WABDSQLSERVER := cS1WABDSQLSERVER.Create;
      try



        (*YCP 17/12/2004
        WABDSQLSERVER.Dupliquer(self, true, true);
        WABDSQLSERVER.Base := nomBase;
        WABDSQLSERVER.mdfFile := WABDSQLSERVER.GetPhysicalName(1);
        WABDSQLSERVER.ldfFile := WABDSQLSERVER.GetPhysicalName(2);

        result := ExecuteDBSpec('CREATE DATABASE ' + Base
           + ' ON ( NAME = ''' + Base + '_data'', FILENAME = ''' + mdfFile + ''', '
           + '      MAXSIZE = UNLIMITED, FILEGROWTH = 10% )'
           + ' LOG ON ( NAME = ''' + Base + '_log'', FILENAME = ''' + ldfFile + ''', '
           + '      MAXSIZE = UNLIMITED, FILEGROWTH = 10% );');

        if WABDSQLSERVER.DetacherBase then
        begin
          result := CopyFile(PChar(WABDSQLSERVER.MdfFile), PChar(mdfFile), true);
          if WABDSQLSERVER.ldfFile <> '' then result := CopyFile(PChar(WABDSQLSERVER.ldfFile), PChar(ldfFile), true)
          else ldfFile := '';
          WABDSQLSERVER.AttacherBase;
        end;
        if result then result := AttacherBase;
        WABDSQLSERVER.free;  
      except
        on E: Exception do
        begin
          ErrorMessage := E.Message;
          WABDSQLSERVER.free;
        end;
      end;
    end;
  end;
end;    *)

function cWABDSQLSERVER.SupprimerBase(): boolean;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
  begin
    result := ExecuteDBSpec('DROP DATABASE ' + Base + ';');
    //if result then SupprimeFichierIni(BaseName,IniName,nil) ;
    if result then ActionMessage := 'La base a été supprimée avec succès'
    else ErrorMessage := 'La suppression de la base a échoué';
  end;
end;

function cWABDSQLSERVER.ArchiverBase(): boolean;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
  begin
    result := not FileExists(BackupFile);
    if not result then
      result := not parle or (PgiAsk('Le fichier de sauvegarde existe déjà. Voulez-vous l''écraser ?', 'Sauvegarde') = mrYes);
    result := result and ExecuteDBSpec
      ('BACKUP DATABASE ' + Base + ' TO DISK=''' + BackupFile + ''' WITH  INIT, NOSKIP, NOFORMAT');
    if result then ActionMessage := 'La base a été sauvegardée avec succès'
    else ErrorMessage := 'La sauvegarde de la base à échoué';
  end;
end;

function cWABDSQLSERVER.RestorerBase(): boolean;
var
  DirectoRy, mdf, ldf, st,table: string ;
  okTable: boolean ;
  T1: Tob;
  i: integer ;
begin
  //connecte:=true ;
  okTable:=false  ; //YCP 06/01/05
  result := DB.connected;
  if result then
  begin
    result := ExecuteDBSpec('use master;');
    if result then
      begin
      DirectoRy:=ExtractFilePath(mdfFile) ;
      (*st:='DECLARE @result int;'
         +'EXEC @result = xp_cmdshell ''dir '''''+Directory+'*.*'''' '' ;'
         +'IF (@result = 0)'
         +'   PRINT ''Success'' '
         +'ELSE '*)
      st:='EXEC xp_cmdshell ''mkdir '+Directory+' '' ' ;
      result:=ExecuteDBSpec(st) ;
      end ;

    table := FormatDateTime('"#DBFiles_"yymmdd"_"hhnnsszzz', now());
      if result then
      begin
      if ( not isSQL2005 ) then
        begin
        st := 'CREATE TABLE ' + table + '  ('
          + ' LogicalName nvarchar( 128 ) , PhysicalName nvarchar( 512 ) , Type char( 1 ) ,'
          +'FileGroupName nvarchar( 128 ) , [Size] varchar( 30 ) , [MaxSize] varchar( 30 ) ) ;';
        end else
        begin
        st := 'CREATE TABLE ' + table + '  ('
          + ' LogicalName nvarchar(128),  PhysicalName nvarchar(260),  Type char(1), FileGroupName nvarchar(128),'
          + ' Size numeric(20,0),  [MaxSize] numeric(20,0),  FileID bigint,  CreateLSN numeric(25,0),'
          + ' DropLSN numeric(25,0) NULL, UniqueID uniqueidentifier, ReadOnlyLSN numeric(25,0) NULL,'
          + ' ReadWriteLSN numeric(25,0) NULL, BackupSizeInBytes bigint, SourceBlockSize int, FileGroupID int,'
          + ' LogGroupGUID uniqueidentifier NULL, DifferentialBaseLSN numeric(25,0) NULL,'
          + ' DifferentialBaseGUID uniqueidentifier, IsReadOnly bit, IsPresent bit';
         if isSQL2008 then
           st := st + ', TDEThumbprint varbinary(32) );' // FQ 16051
         else
           st := st + ');' ;
        end ;
        result := ExecuteDBSpec(st);
        okTable := true;
      end;

    if result then
    begin
      st := 'DECLARE @CmdStr varchar( 8000 ) ;';
      if xxDumpFile then
        st := st + 'SET @CmdStr = ''RESTORE FILELISTONLY FROM ' + xxDumpName + ' '' ;'
      else
        st := st + 'SET @CmdStr = ''RESTORE FILELISTONLY FROM DISK=''''' + BackUpFile + ''''' '' ;';
      st := st + 'INSERT INTO ' + table (*+ ('
        + '  LogicalName , PhysicalName , Type , FileGroupName ,'
        + '  [Size] , [MaxSize]'
        + '  ) EXEC( @CmdStr );';*)
        + ' EXEC( @CmdStr );';

       result:=ExecuteDBSpec(st) ;

      T1 := LoadDetailFromSQLSpec('select CAST(LogicalName as VarChar(128)) [filename],CAST(Type as VarChar(1)) [type] from '+table);
      if (T1 <> nil) then
      begin
        for i := 0 to T1.detail.count-1 do
        begin
          if (T1.detail[i].getValue('TYPE')='D') then mdf := T1.detail[i].getValue('FILENAME');
          if (T1.detail[i].getValue('TYPE')='L') then ldf := T1.detail[i].getValue('FILENAME');
        end;
        T1.free ;
      end;
    end ;

    if result or okTable then result:=ExecuteDBSpec('DROP TABLE '+table) ;

    if xxDumpFile then st:='RESTORE DATABASE '+Base+' FROM '+xxDumpName
                  else st:='RESTORE DATABASE '+Base+' FROM DISK='''+BackUpFile+'''' ;
    st:=st+' WITH REPLACE,'
      +' MOVE '''+mdf+''' TO '''+mdfFile+''','
      +' MOVE '''+ldf+''' TO '''+ldfFile+'''' ;
    if result then result:=ExecuteDBSpec(st) ;

    //result := ExecuteDBSpec('RESTORE DATABASE ' + Base + ' FROM DISK=''' + BackUpFile + ''' WITH REPLACE;');
    if result then ActionMessage := 'La base a été restaurée avec succès'
    else ErrorMessage := 'La restauration de la base a échoué';
  end;
end;

function cWABDSQLSERVER.AttacherBase(): boolean;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
    result := ExecuteDBSpec('use master;');

  if result then
  begin
    if ldfFile <> '' then
      result := ExecuteDBSpec('EXEC sp_attach_db @dbname = ''' + Base + ''','
        + '  @filename1 = ''' + mdfFile + ''','
        + '  @filename2 = ''' + ldfFile + ''';')
    else
      result := ExecuteDBSpec('EXEC sp_attach_single_file_db @dbname = ''' + Base + ''','
        + '  @physname = ''' + mdfFile + ''';');
    //if result then CreeFichierINI(Base,Group,IniName,nil) ;

    if result then ActionMessage := 'La base a été attachée avec succès'
    else ErrorMessage := 'L''attachement de la base a échoué';
  end;
end;

function cWABDSQLSERVER.DetacherBase(): boolean;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
    result := ExecuteDBSpec('use master;');

  if result then
  begin
    TraceExecution('dans detacher connect ok !');
    result := ExecuteDBSpec('EXEC sp_detach_db ''' + Base + ''', ''true'';');
    //if result then SupprimeFichierIni(Base,IniName,nil) ;
    if result then ActionMessage := 'La base a été détachée avec succès'
    else ErrorMessage := 'Le détachement de la base a échoué';
  end;
end;

function cWABDSQLSERVER.ExisteBase(): boolean;
var
  T1: TOB;
begin
  //connecte:=true ;
  result := DB.connected;
  if result then
  begin
    T1 := LoadDetailFromSQLSpec('select name from master.dbo.sysdatabases where name=''' + Base + '''');
    if (T1 = nil) or (T1.detail.count = 0) then
      Result := False
    else Result := True;
    if (T1 <> nil) then T1.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... : 04/12/2002
Description .. : Tester la connexion au server
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.TestConnexion: boolean;
var
{$IFNDEF EAGLCLIENT}
  oldParle: boolean;
{$ENDIF EAGLCLIENT}
begin
  oldParle := Parle;
  Parle := false;
  result := DB.connected;
  Parle := oldParle;
  if result then ActionMessage := 'Test de connexion réussi'
  else ErrorMessage := 'Test de connexion échoué';
end;
{$ENDIF EAGLCLIENT}

function cWABDSQLSERVER.AutocloseBase(bOption : Boolean): boolean;
var
  ResultTob,tmpTOB: Tob; //YCP 15/05/07 
begin
  if ModeEagl then
  begin
    tmpTOB := tob.create('OPTIONS', nil, -1);
    tmpTOB.AddChampSupValeur('autoclose', bOption);
    AjouteUneTob(self, tmpTob);
    (*with tob.Create('OPTIONS', self, -1) do begin AddChampSupValeur('autoclose', bOption); end;*)
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'AutocloseBase', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
    self.ClearDetail ;
  end else
  begin
    {$IFNDEF EAGLCLIENT}
    result:=DB.connected ;
    if result then
      result := ExecuteDBSpec('use master;');
    if result then
      begin
    if bOption then
      result:=ExecuteDBSpec('EXEC sp_dboption '''+Base+''', autoclose, true')
      else result:=ExecuteDBSpec('EXEC sp_dboption '''+Base+''', autoclose, false');
      end ;
    {$ELSE} // XP WC 04-07-2005
    result := false;
    {$ENDIF EAGLCLEINT}
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Cronstructeur de l'objet "outil"
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
constructor cWABDSQLSERVER.Create;
begin
  inherited create();
{$IFNDEF EAGLCLIENT}
  setDB(nil);
  xxForceBackupDirectory:='' ;
{$ENDIF EAGLCLIENT}
  ReaffecteNomTable(className, true);
  driver:=DriverToSt(V_PGI.Driver,true) ;
  FileGrowth := '10%';
end;


{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Destructeur de l'objet "outil"
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
destructor cWABDSQLSERVER.Destroy;
begin
{$IFNDEF EAGLCLIENT}
  setDB(nil) ;
{$ENDIF}
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Récupère le nom du lien ODBC de la machine local
Mots clefs ... :
*****************************************************************}
class function cWABDSQLSERVER.GetLienODBC(): string;
var
  TR: TRegistry;
  ts: TStringList;
  ii: integer;
  Buffer: array[0..1023] of Char;
  wst: string;
begin
  result := '';
  TR := TRegistry.Create;
  TS := TStringList.Create;
  try
    with TR, TS do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', false) then GetValueNames(ts);
      CloseKey;

      for ii := 0 to ts.Count - 1 do
      begin
        if OpenKey('SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', false)
          and ValueExists(Strings[ii]) and (ReadString(Strings[ii]) = 'SQL Server') then
        begin
          CloseKey;
          if OpenKey('SOFTWARE\ODBC\ODBC.INI\' + Strings[ii], false)
            and ValueExists('Server') and (ReadString('Server') = '(local)')
            then
          begin
            result := Strings[ii];
            closeKey;
            break;
          end;
        end;
        closeKey;
      end;

      if result = '' then
      begin
        result := 'PGI' + NomHalley + '_LOCAL';
        if OpenKey('SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources', true) then
        begin
          WriteString(result, 'SQL Server');
          CloseKey;
        end;

        if OpenKey('SOFTWARE\ODBC\ODBC.INI\' + result, true) then
        begin
          WriteString('Server', '(local)');
          GetSystemDirectory(Buffer, 1023);
          System.SetString(wst, Buffer, StrLen(Buffer));
          WriteString('Driver', wst + '\sqlsrv32.dll');
          WriteString('QuotedId', 'No');
          WriteString('Language', 'us_english');
          WriteString('Lastuser', 'sa');
          CloseKey;
        end;
      end;
    end;
  finally
    TR.free;
    TS.free;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... : 04/12/2002
Description .. : Réparer une base SQLSERVER
Suite ........ :   - ajout de l'utilisateur ADMIN/ADMIN
Suite ........ : Param5: Nom de la base
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ReparerDatabase(): boolean;
var
  ResultTob: Tob;
{$IFNDEF EAGLCLIENT}
  T1: TOB;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'REPARE', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
  end else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    if ( not isSQL2005 ) then //YCP 25/10/2007 pb de création de l'utilisateur ADMIN
      begin
      //connecte:=true ;
      if DB.connected then result := true;
      if result then
      begin
        result := ExecuteDBSpec('use ' + Base + ';');
        if result then
        begin
          AjouterLogin('ADMIN', 'ADMIN');
          T1 := LoadDetailFromSQLSpec('select count(*) AS NB from  sysusers left join master.dbo.syslogins '
            + ' on sysusers.sid=master.dbo.syslogins.sid where master.dbo.syslogins.name=''ADMIN''');
          if (T1 = nil) or (T1.detail.count = 0) or (T1.detail[0].GetValue('NB') = 0) then
            result := ExecuteDBSpec('EXEC sp_adduser ''ADMIN'', ''ADMIN'', ''db_owner'' ');
          if (T1 <> nil) then T1.free;
        end;
      end ;
    end else
    begin
    result:=true ;
    end ;
{$ENDIF EAGLCLIENT}
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : ajouter un login
Suite ........ : Param1: Nom du login
Suite ........ : Param2: Pass du login
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.AjouterLogin(UserName, Password: string): boolean;
var
  tmpTOB, ResultTob: Tob;
{$IFNDEF EAGLCLIENT}
  T1: Tob;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    tmpTOB := tob.create('OPTIONS', nil, -1);
    tmpTOB.AddChampSupValeur('UserName', UserName);
    tmpTOB.AddChampSupValeur('Password', Password);
    AjouteUneTob(self, tmpTob);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'ADDLOGIN', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
    self.cleardetail ; //YCP 15/05/07 
  end else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected then result := true;
    if result then
      result := ExecuteDBSpec('use master;');
      if result then
      begin
        T1 := LoadDetailFromSQLSpec('select name from syslogins where name = ''' + UserName + '''');
        if (T1 = nil) or (T1.detail.count = 0) then
        begin
          result := ExecuteDBSpec('EXEC sp_addlogin @loginame = ''' + UserName + ''', @passwd = ''' + Password + '''');
          result := result and ExecuteDBSpec('EXEC sp_addsrvrolemember ''' + UserName + ''', ''sysadmin''');
        end;
        if (T1 <> nil) then T1.free;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;


{$WARN SYMBOL_PLATFORM OFF}
function cWABDSQLSERVER.NEW_BACKUP(nomBase: string): TOB  ;
var
{$IFNDEF EAGLCLIENT}
  xxDumpName: string ;
  Backupfile: string ;
  repertoire : string ;
  OkOK: boolean ;
  q: Tob ;
{$ENDIF EAGLCLIENT}
  tmpTOB, ResultTob: Tob;
begin
  if ModeEagl then
  begin
    tmpTOB := tob.create('OPTIONS', nil, -1);
    tmpTOB.AddChampSupValeur('nomBase', nomBase);
    AjouteUneTob(self, tmpTob);
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'NEW_BACKUP', self, '', '');
    GetRetour(ResultTob,result);
    ResultTob.free;
  end else
  begin
  result := nil;
  {$IFNDEF EAGLCLIENT}
  {$IFDEF EAGLSERVER}
  if nombase='' then
    begin
    nombase:=LookupCurrentSession.DBName ;
    end ;
  {$ENDIF EAGLSERVER}
  OkOK:=ExecuteDBSpec('BACKUP LOG '+nomBase+' WITH NO_LOG;') ; //YCP 06/01/05
  if OkOK then
    begin
    xxDumpName:=FormatDateTime('"DATA_"yymmdd"_"hhnnsszzz""',now()) ;
    if xxForceBackupDirectory<>'' then
      begin
      repertoire:=xxForceBackupDirectory
      end 
    else
      begin
      repertoire:='C:\TEMP' ;
      q:=LoadDetailFromSQLSpec('SELECT SOC_DATA FROM PGIADMIN..PARAMSOC WHERE SOC_NOM=''SO_CHEMINSAVE''') ;
      if (q <> nil) and (q.detail.count > 0) and (q.Detail[0].FieldExists('SOC_DATA')) then
        repertoire:=q.Detail[0].getValue('SOC_DATA') ;
      q.free ;
      end ;
    {$WARN SYMBOL_PLATFORM OFF}
    BackUpFile:=IncludeTrailingBackslash(repertoire)+ChangeFileExt(xxDumpName,'.BAK') ;
    {$WARN SYMBOL_PLATFORM ON}
    OkOK:=ExecuteDBSpec('sp_addumpdevice ''disk'', '''+xxDumpName+''', '''+BackUpFile+'''') ;
    try
    if OkOK then
      begin
      OkOK:=ExecuteDBSpec('BACKUP DATABASE '+nomBase+' TO '+xxDumpName+' WITH  INIT;');
      if OkOK then
        begin
        result:=LoadDetailFromSQLSpec('SELECT * FROM OPENROWSET(BULK N'''+BackUpFile+''', SINGLE_BLOB) as test') ;
        end ;
      end ;
    finally
      ExecuteDBSpec('sp_dropdevice '''+xxDumpName+''', DELFILE') ;
      end ;
    end ;
  {$ENDIF EAGLCLIENT}
  end ;
end ;
{$WARN SYMBOL_PLATFORM ON}


{$WARN SYMBOL_PLATFORM OFF}
function cWABDSQLSERVER.NEW_RESTORE(nomBase,filename: string): TOB  ;
var
{$IFNDEF EAGLCLIENT}
  xxDumpName: string ;
  Backupfile: string ;
  repertoire : string ;
  OkOK: boolean ;
{$ENDIF EAGLCLIENT}
  tmpTOB, ResultTob: Tob;
begin
  if ModeEagl then
  begin
    tmpTOB := tob.create('OPTIONS', nil, -1);
    tmpTOB.AddChampSupValeur('nomBase', nomBase);

    AjouteUneTob(self, tmpTob);
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'NEW_BACKUP', self, '', '');
    GetRetour(ResultTob,result);
    ResultTob.free;
  end else
  begin
  result := nil;
  {$IFNDEF EAGLCLIENT}
  OkOK:=ExecuteDBSpec('BACKUP LOG '+nomBase+' WITH NO_LOG;') ; //YCP 06/01/05
  if OkOK then
    begin
    xxDumpName:=FormatDateTime('"DATA_"yymmdd"_"hhnnsszzz""',now()) ;
    repertoire:='c:\temp' ;
    {$WARN SYMBOL_PLATFORM OFF}
    BackUpFile:=IncludeTrailingBackslash(repertoire)+ChangeFileExt(xxDumpName,'.BAK') ;
    {$WARN SYMBOL_PLATFORM ON}
    OkOK:=ExecuteDBSpec('sp_addumpdevice ''disk'', '''+xxDumpName+''', '''+BackUpFile+'''') ;
    try
    if OkOK then
      begin
      OkOK:=ExecuteDBSpec('BACKUP DATABASE '+nomBase+' TO '+xxDumpName+' WITH  INIT;');
      if OkOK then
        begin
        result:=LoadDetailFromSQLSpec('SELECT * FROM OPENROWSET(BULK N'''+BackUpFile+''', SINGLE_BLOB) as test') ;
        end ;
      end ;
    finally
      ExecuteDBSpec('sp_dropdevice '''+xxDumpName+''', DELFILE') ;
      end ;
    end ;
  {$ENDIF EAGLCLIENT}
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : modifier le login de sa
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ModifierLoginCEGIDSA: boolean;
var
  ResultTob: Tob;
begin
  if ModeEagl then
  begin
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'CEGIDSA', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
  end else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected then result := true;
    if result then
      result := ExecuteDBSpec('use master;');

    if result then
    begin
      result := ExecuteDBSpec('exec sp_password NULL,''cegid'',''sa''');
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Recupérer une TOB des bases du server
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ListerDatabase(all: boolean): TOB;
var
  ResultTOB,tmpTOB: TOB; //YCP 15/05/07 
{$IFNDEF EAGLCLIENT}
  st: string;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    Result := Tob.Create('', nil, -1);
    tmpTOB := tob.create('OPTIONS', nil, -1);
    tmpTOB.AddChampSupValeur('all', all);
    AjouteUneTob(self, tmpTob);
    (*with tob.Create('OPTIONS', self, -1) do begin AddChampSupValeur('all', all); end; *)
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'LISTDATABASE', self, '', '');
    GetRetour(ResultTob, result);
    ResultTob.free;
    self.ClearDetail ;
  end
  else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected and ExecuteDBSpec('use master;') then
    begin
      st := 'SELECT CAST(Name as VarChar(255)) [name],dbid, CAST(filename as VarChar(255)) [filename]'
        + ' from sysdatabases';
      if not all then
        st := st + ' where name not in (''master'',''model'',''msdb'',''tempdb'')';
      result := LoadDetailFromSQLSpec(st);
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Recupérer la liste des utilistareurs dans une TOB
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ListerUser(): TOB;
var
  ResultTOB: TOB;
begin
  if ModeEagl then
  begin
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'LISTUSER', self, '', '');
    GetRetour(ResultTob, result);
    ResultTob.free;
  end else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected and ExecuteDBSpec('use master') then
    begin
      result := LoadDetailFromSQLSpec('select spid,CAST(loginame as VarChar(128)) [loginame],CAST(hostname as VarChar(128)) [hostname],'
        + 'CAST(program_name as VarChar(128)) [program_name] ,  CAST(status as VarChar(128)) [status] '
        + ' from sysprocesses where hostname<>'''' order by spid');
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le . .. :   /  /
Description .. : Recupérer la liste des utilistareurs dans une TOB
Suite ........ :  connectés
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ListerTables(): TOB;
var
  ResultTOB: TOB;
begin
  if ModeEagl then
  begin
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'LISTTABLE', self, '', '');
    GetRetour(ResultTob, result);
    ResultTob.free;
  end else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected and ExecuteDBSpec('use ' + Base + ';') then
    begin
      result := LoadDetailFromSQLSpec('SELECT CAST(Name as VarChar(255)) [name] from sysobjects where xtype=''U'' order by name');
    end;
{$ENDIF EAGLCLIENT}
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le . .. :   /  /
Description .. : Recupérer la liste des utilistareurs dans une TOB
Suite ........ :  connectés
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ListerJob(): TOB;
var
  ResultTOB: TOB;
{$IFNDEF EAGLCLIENT}
  ii: integer;
  st: string;
  T1: Tob;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'LISTJOB', self, '', '');
    GetRetour(ResultTob, result);
    ResultTob.free;
  end else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected and ExecuteDBSpec('use msdb') then
    begin
      result := LoadDetailFromSQLSpec('select CAST(name as VarChar(128)) [name], CAST(job_id as VarChar(128)) [job_id]'
        + ' from sysjobs');
      if result <> nil then
      begin
        for ii := 0 to result.Detail.Count - 1 do with result.detail[ii] do
          begin
            T1 := LoadDetailFromSQLSpec('exec sp_help_job @job_name=''' + GetValue('Name') + '''');
            if (T1 <> nil) and (T1.detail.count > 0) then
            begin
              st := T1.detail[0].GetValue('enabled');
              AddChampSupValeur('enabled', si(st = '1', 'X', '-'));
              st := string(T1.detail[0].GetValue('last_run_time')) ;
              while length(st)<6 do st:='0'+st ;
              st := string(T1.detail[0].GetValue('last_run_date')) + st;
              if st = '0000000' then
                AddChampSupValeur('last_run_date', '')
              else
                AddChampSupValeur('last_run_date', copy(st, 7, 2) + '/' + copy(st, 5, 2) + '/' + copy(st, 1, 4) + ' ' + copy(st, 9, 2) + ':' + copy(st, 11, 2) + ':' + copy(st, 13, 2));

              st := string(T1.detail[0].GetValue('next_run_time')) ;
              while length(st)<6 do st:='0'+st ;
              st := string(T1.detail[0].GetValue('next_run_date')) + st;
              if st = '0000000' then
                AddChampSupValeur('next_run_date', '')
              else
                AddChampSupValeur('next_run_date', copy(st, 7, 2) + '/' + copy(st, 5, 2) + '/' + copy(st, 1, 4) + ' ' + copy(st, 9, 2) + ':' + copy(st, 11, 2) + ':' + copy(st, 13, 2));
            end;
            T1.free;
          end;
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;

end;



{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 09/12/2002
Modifié le ... :   /  /
Description .. : Ajoute un job de sauvegarde au server
Suite ........ : Param1: Nom de la base
Suite ........ : Param2: Fichier de sauvegarde
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.AjouterJob(): boolean;
var
  ResultTob: Tob;
begin
  if ModeEagl then
  begin
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'ADDJOB', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
  end else
  begin
    result := false;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected then result := true;
    if result then
    begin
      result := ExecuteDBSpec('USE msdb;');
      if result then
      begin
        Supprimerjob();
        if result then
          result := ExecuteDBSpec('exec sp_add_job @job_name=''' + JobName + ''';');

        if result then
          result := ExecuteDBSpec('EXEC sp_add_jobstep @job_name=''' + JobName + ''','
            + '@step_name=''Sauvegarde de la base ' + Base + ''','
            + '@subsystem=''TSQL'','
            + '@command  =''BACKUP DATABASE [' + Base + '] TO DISK=''''' + BackupFile + ''''' WITH  INIT , NOSKIP ,  NOFORMAT'','
            + '@retry_attempts = 5, @retry_interval = 5;');

        if result then
          result := ExecuteDBSpec('EXEC sp_add_jobschedule @job_name=''' + jobname + ''''
            + ',@name=''Planification de la base ' + Base + ''',@freq_type=' + TypeFreq
            + si(TypeFreq <> '1', ',@freq_interval=' + InterVal, '')
            + si((TypeFreq = '8') or (TypeFreq = '16'), ',@freq_recurrence_factor=1', '')
            + si(DateDebut <> '', ',@active_start_date=' + DateDebut, '')
            + si(HeureDebut <> '', ',@active_start_time=' + HeureDebut, ''));

        if result then
          result := ExecuteDBSpec('EXEC sp_add_jobserver @job_name=''' + jobname + ''';');
      end;
    end;
{$ENDIF EAGLCLIENT}
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : récupère les infos d'une sauvegarde
Mots clefs ... :
*****************************************************************}
function cWABDSQLSERVER.GetInfosJob(): TOB;
var
  ResultTOB: TOB;
{$IFNDEF EAGLCLIENT}
  st: string;
  T1: Tob;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    Result := Tob.Create('', nil, -1);
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'INFOSJOB', self, '', '');
    GetRetour(ResultTob, result);
    ResultTob.free;
  end else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    //connecte:=true ;
    if DB.connected and ExecuteDBSpec('use msdb') then
    begin
      st:='select CAST(j.name as VarChar(30)) [name], CAST(t.command as VarChar(254)) [command], ' ;
      st:= st + ' freq_type,freq_interval,active_start_date,active_start_time';
      st:= st + ' from msdb..sysjobs j';

      if (isSQL2005()) then
        begin
        st:= st + ' join msdb..sysJobSchedules s on j.job_id = s.job_id';
        st:= st + ' join msdb..sysschedules k on k.schedule_id = s.schedule_id';
        st:= st + ' join msdb..sysJobSteps t on j.job_id = t.job_id' ;
        end else
        begin
        st:= st + ' join msdb..sysJobSchedules s on	j.job_id = s.job_id';
        st:= st + ' join msdb..sysJobSteps t on	j.job_id = t.job_id ';
        end ;
      st:= st + ' where j.job_id='''+job_id+'''' ;
      T1 := LoadDetailFromSQLSpec(st);
      if (T1 <> nil) and (T1.detail.count > 0) then result := T1.Detail[0] else result := nil;
      if result <> nil then with result do
        begin
          st := GetValue('command');
          AddChampSupValeur('basename', copy(copy(st, 1, pos(']', st) - 1), pos('[', st) + 1, 250));

          st := string(GetValue('Active_Start_Date'));
          PutValue('Active_Start_Date', copy(st, 7, 2) + '/' + copy(st, 5, 2) + '/' + copy(st, 1, 4));

          st := string(GetValue('Active_Start_Time'));
          //YCP 31/08/05
          while length(st)<6 do st:='0'+st ;
          PutValue('Active_Start_Time', copy(st, 1, 2) + ':' + copy(st, 3, 2))  ;

          st := GetValue('command');
          st := copy(st, pos('''', st) + 1, length(st));
          st := copy(st, 1, pos('''', st) - 1);
          AddChampSupValeur('BackupFile', st);
        end;
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{$IFNDEF EAGLCLIENT}
function cWABDSQLSERVER.GetPhysicalName(lequel: integer): string;
var
  i: integer;
  T1: Tob;
begin
  //connecte:=true ;
  if DB.connected and ExecuteDBSpec('use msdb') then
  begin
    T1 := LoadDetailFromSQLSpec('select CAST(filename as VarChar(250)) [filename] from ' + base + '.dbo.sysfiles');
    if (T1 <> nil) then
    begin
      for i := 0 to T1.detail.count do
      begin
        if (lequel = 1) and (pos('MDF', UpperCase(T1.detail[i].getValue('FILENAME'))) > 0) then result := T1.detail[i].getValue('FILENAME');
        if (lequel = 2) and (pos('LDF', UpperCase(T1.detail[i].getValue('FILENAME'))) > 0) then result := T1.detail[i].getValue('FILENAME');
        if result <> '' then break;
      end;
    end;
  end;
end;
{$ENDIF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 09/12/2002
Modifié le ... :   /  /
Description .. : supprime un job de sauvegarde au server
Suite ........ : Param1: Nom du job
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.SupprimerJob(): boolean;
var
  ResultTob: Tob;
begin
  if ModeEagl then
  begin
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'DELETEJOB', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
  end else
  begin
{$IFDEF EAGLCLIENT}
    result := false; //cas impossible
{$ELSE}
    //connecte:=true ;
    result := DB.connected;
    if result then
    begin
      result := ExecuteDBSpec('USE msdb;');
      if result then
        result := ExecuteDBSpec('if exists (select name from sysjobs where job_id = ''' + Job_id + ''' )' + #13
          + '    EXEC sp_delete_job @job_id=''' + Job_id + ''';');
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{$IFNDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Recupérer une TOB de la requette
Suite ........ : Param1: requette
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.LoadDetailFromSQLSpec(wSql: string): TOB;
{$IFDEF EAGLCLIENT}
{$ELSE}
var
  TQ: TQuery;
  jj: integer;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    result := Request(dbss_dll + '.cWABDSQLSERVER', 'LoadDetailFromSQLSpec', self, '', '');
  end else
  begin
    result := nil;
{$IFNDEF EAGLCLIENT}
    TQ := nil;
    if (DB.Connected) then
    begin
      result := Tob.Create('LoadDetailFromSQLSpec', nil, -1);
      try
        TQ := TQuery.Create(Application);
        with TQ do
        begin
          SQLConnection := DB;
          SessionName := DB.SessionName;
          SQL.Clear;
          SQL.Add(wSQL);
          Trace.TraceInformation('UWA',' SQL : '+wSQL);
          RequestLive := FALSE;
          Open;
          while not TQ.Eof do
          begin
            with Tob.Create('sysdatabases', result, -1) do
              for jj := 0 to TQ.fields.count - 1 do
              begin
                AddChampSup(TQ.fields[jj].fieldname, false);
                PutValue(TQ.fields[jj].fieldname, TQ.fields[jj].Value);
              end;
            next;
          end;
          close;
          Free;
        end;
      except
        on E: Exception do
        begin
          if TQ <> nil then TQ.free;
          if result <> nil then
          begin
            result.free;
            result := nil;
          end;
          ErrorMessage := E.Message;
        end;
      end;
    end else
    begin
      ErrorMessage := 'Erreur de connection au serveur';
    end;
{$ENDIF EAGLCLIENT}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 09/12/2002
Modifié le ... :   /  /
Description .. : Execute une requette SQL
Suite ........ : Param1: Requette SQL
Mots clefs ... : S1;;DATABASE;
*****************************************************************}
function cWABDSQLSERVER.ExecuteDBSpec(wSQL: string): boolean;
var
  ResultTob: Tob;
{$IFDEF EAGLCLIENT}
{$ELSE}
  TQ: TQuery;
{$ENDIF EAGLCLIENT}
begin
  if ModeEagl then
  begin
    ResultTOB := Request(dbss_dll + '.cWABDSQLSERVER', 'EXECUTESQL', self, '', '');
    result := GetRetour(ResultTob);
    ResultTob.free;
  end else
  begin
{$IFDEF EAGLCLIENT}
    result := false; //cas impossible
{$ELSE}
    TQ := nil;
    result := false;
    if (DB.Connected) then
    begin
      try
        TQ := TQuery.Create(Application);
        with TQ do
        begin
          SQLConnection := DB;
          SessionName := DB.SessionName;
          SQL.Clear;
          SQL.Add(wSQL);
          RequestLive := FALSE;
          ExecSql;
          close;
          Free;
        end;
        result := true;
      except
        on E: Exception do
        begin
          result := false;
          if (TQ <> nil) then
          begin
            TQ.Close;
            TQ.free;
          end;
          ErrorMessage := E.Message;
        end;
      end;
    end else
    begin
      ErrorMessage := 'Erreur de connection au serveur';
    end;
{$ENDIF EAGLCLIENT}
  end;
end;
{$ENDIF EAGLCLIENT}


{$IFNDEF BASEEXT}
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : Récupère les infos de connections
Mots clefs ... :
*****************************************************************}
class procedure cWABDSQLSERVER.GetInfoConnection(Servername, user, password: THEdit);
var
  s, u, p: string;
begin
  GetInfoConnection(s, u, p);
  SERVERNAME.Text := s;
  USER.Text := u;
  PASSWORD.Text := p;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : récupère les infos de connections à la base
Mots clefs ... :
*****************************************************************}
class procedure cWABDSQLSERVER.GetInfoConnection(var Servername, user, password: string);
var
  CheminCle, st: string;
begin
  CheminCle := 'Software\' + Apalatys + '\' + NomHalley + '\';

  if V_PGI.okOuvert then
  begin
    SErverName := GetParamSocSecur('SO_MSDESERVER', '');
    if SErverName = '' then ServerName := GetLienODBC();
    USER := GetParamSocSecur('SO_MSDEUSER', '');
    st := GetParamSocSecur('SO_MSDEPASSWORD', '');
    if st <> '' then PASSWORD := DeCryptageSt(st);
  end else
  begin
    SErverName := GetFromRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'SERVERNAME', '');
    if SErverName = '' then ServerName := GetLienODBC();
    USER := GetFromRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'USER', '');
    st := GetFromRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'PASSWORD', '');
    if st <> '' then PASSWORD := DeCryptageSt(st);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 02/03/2004
Modifié le ... :   /  /
Description .. : renseigne les infos de connections à la base
Mots clefs ... :
*****************************************************************}
class procedure cWABDSQLSERVER.setInfoConnection(Servername, user, password: string);
var
  CheminCle: string;
begin
  CheminCle := 'Software\' + Apalatys + '\' + NomHalley + '\';
  if V_PGI.okOuvert then
  begin
    SetParamSoc('SO_MSDESERVER', SErverName);
    SetParamSoc('SO_MSDEUSER', USER);
    SetParamSoc('SO_MSDEPASSWORD', CryptageSt(PASSWORD));
  end else
  begin
    SaveInRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'SERVERNAME', SErverName);
    SaveInRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'USER', USER);
    SaveInRegistry(HKEY_LOCAL_MACHINE, CheminCle, 'PASSWORD', CryptageSt(PASSWORD));
  end;
end;
{$ENDIF BASEEXT}

{$IFNDEF EAGLCLIENT}
function cWABDSQLSERVER.getDB: TDatabase;
var
  WAInfoIni: cWAInfoIni;
begin
  if (xxDB = nil) then
  begin
    //inc(cpte);
    WAInfoIni := cWAInfoIni.create;
    try
      WAInfoIni.Base := 'master';
      WAInfoIni.server := MasterServer;
      WAInfoIni.User := MasterUser;
      WAInfoIni.Pass := MasterPass;

      WAInfoIni.Driver := DriverToSt(dbMSSQL,true) ; //'ODBC_MSSQL'; mko 19/01/2010 fq 18209
      WAInfoIni.ODBC := 'Microsoft OLEDB Driver';
      xxDB:=TCBPDatabases.CreateNamedDataBase('LOCALMASTER') ;
      cWAINI.connectedb(WAInfoIni, xxDB) ; //, 'LOCALMASTER') ; //24/05/2007+ IntToStr(GetCurrentThreadId));
    finally
      WAInfoIni.free;
    end;
  end;
  result := xxDB;
end;

procedure cWABDSQLSERVER.setDB(valeur: TDatabase);
begin
if assigned(xxDB) then //YCP 15/05/07
  TCBPDatabases.FreeDataBase(xxDB) ;
xxDB := valeur;
end;


(*function cS1WABDSQLSERVER.GetConnecte: boolean ;
begin
result:=DB.Connected ;
end ;

procedure cS1WABDSQLSERVER.SetConnecte(Value: boolean) ;
BEGIN
try
  DB.Connected:=Value ;
except
  on E: Exception do
    ErrorMessage:=E.Message ;
  end ;
end ;*)
{$ENDIF EAGLCLIENT}

function cWABDSQLSERVER.getmdfFile: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('mdfFile')) then self.AddChampSupValeur('mdfFile', '');
  result := self.getString('mdfFile');
end;

procedure cWABDSQLSERVER.setmdfFile(valeur: string);
begin
  if (not self.FieldExists('mdfFile')) then self.AddChampSup('mdfFile', false);
  self.PutValue('mdfFile', valeur);
end;

function cWABDSQLSERVER.getMasterPass: string; //type de ini (INI, BAS, ...)
begin
{$IFDEF EAGLCLIENT}
  if (not self.FieldExists('MasterPass')) then self.AddChampSupValeur('MasterPass', '');
{$ELSE}
  if (not self.FieldExists('MasterPass')) then self.AddChampSupValeur('MasterPass', DBSOC.GetParamConnection('PASSWORD'));
{$ENDIF EAGLCLIENT}
  result := self.getString('MasterPass');
end;

procedure cWABDSQLSERVER.setMasterPass(valeur: string);
begin
  if (not self.FieldExists('MasterPass')) then self.AddChampSup('MasterPass', false);
  self.PutValue('MasterPass', valeur);
end;

function cWABDSQLSERVER.getMasterUser: string; //type de ini (INI, BAS, ...)
begin
{$IFDEF EAGLCLIENT}
  if (not self.FieldExists('MasterUser')) then self.AddChampSupValeur('MasterUser', '');
{$ELSE}
  if (not self.FieldExists('MasterUser')) then self.AddChampSupValeur('MasterUser', DBSOC.GetParamConnection('User ID'));
{$ENDIF EAGLCLIENT}
  result := self.getString('MasterUser');
end;

procedure cWABDSQLSERVER.setMasterUser(valeur: string);
begin
  if (not self.FieldExists('MasterUser')) then self.AddChampSup('MasterUser', false);
  self.PutValue('MasterUser', valeur);
end;

function cWABDSQLSERVER.getMasterserver: string; //type de ini (INI, BAS, ...)
begin
{$IFDEF EAGLCLIENT}
  if (not self.FieldExists('Masterserver')) then self.AddChampSupValeur('Masterserver', '');
{$ELSE}
  if (not self.FieldExists('Masterserver')) then self.AddChampSupValeur('Masterserver', DBSOC.GetParamConnection('Data Source'));
{$ENDIF EAGLCLIENT}
  result := self.getString('Masterserver');
end;

procedure cWABDSQLSERVER.setMasterserver(valeur: string);
begin
  if (not self.FieldExists('Masterserver')) then self.AddChampSup('Masterserver', false);
  self.PutValue('Masterserver', valeur);
end;

function cWABDSQLSERVER.getjob_id: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('job_id')) then self.AddChampSupValeur('job_id', '');
  result := self.getString('job_id');
end;

procedure cWABDSQLSERVER.setjob_id(valeur: string);
begin
  if (not self.FieldExists('job_id')) then self.AddChampSup('job_id', false);
  self.PutValue('job_id', valeur);
end;


function cWABDSQLSERVER.getInterVal: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('InterVal')) then self.AddChampSupValeur('InterVal', '');
  result := self.getString('InterVal');
end;

procedure cWABDSQLSERVER.setInterVal(valeur: string);
begin
  if (not self.FieldExists('InterVal')) then self.AddChampSup('InterVal', false);
  self.PutValue('InterVal', valeur);
end;

function cWABDSQLSERVER.getJobName: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('JobName')) then self.AddChampSupValeur('JobName', '');
  result := self.getString('JobName');
end;

procedure cWABDSQLSERVER.setJobName(valeur: string);
begin
  if (not self.FieldExists('JobName')) then self.AddChampSup('JobName', false);
  self.PutValue('JobName', valeur);
end;

function cWABDSQLSERVER.getTypeFreq: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('TypeFreq')) then self.AddChampSupValeur('TypeFreq', '');
  result := self.getString('TypeFreq');
end;

procedure cWABDSQLSERVER.setTypeFreq(valeur: string);
begin
  if (not self.FieldExists('TypeFreq')) then self.AddChampSup('TypeFreq', false);
  self.PutValue('TypeFreq', valeur);
end;

function cWABDSQLSERVER.getDateDebut: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('DateDebut')) then self.AddChampSupValeur('DateDebut', '');
  result := self.getString('DateDebut');
end;

procedure cWABDSQLSERVER.setDateDebut(valeur: string);
begin
  if (not self.FieldExists('DateDebut')) then self.AddChampSup('DateDebut', false);
  self.PutValue('DateDebut', valeur);
end;

function cWABDSQLSERVER.getHeureDebut: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('HeureDebut')) then self.AddChampSupValeur('HeureDebut', '');
  result := self.getString('HeureDebut');
end;

procedure cWABDSQLSERVER.setHeureDebut(valeur: string);
begin
  if (not self.FieldExists('HeureDebut')) then self.AddChampSup('HeureDebut', false);
  self.PutValue('HeureDebut', valeur);
end;

function cWABDSQLSERVER.getldfFile: string; //type de ini (INI, BAS, ...)
begin
  if (not self.FieldExists('ldfFile')) then self.AddChampSupValeur('ldfFile', '');
  result := self.getString('ldfFile');
end;

function cWABDSQLSERVER.getFileGrowth: string; // mode accroissement : 10% = en pourcentage, 10 = en Mo
begin
  if (not self.FieldExists('FileGrowth')) then self.AddChampSupValeur('FileGrowth', '');
  result := self.getString('FileGrowth');
end;

procedure cWABDSQLSERVER.setldfFile(valeur: string);
begin
  if (not self.FieldExists('ldfFile')) then self.AddChampSup('ldfFile', false);
  self.PutValue('ldfFile', valeur);
end;

procedure cWABDSQLSERVER.setFileGrowth(valeur: string);
begin
  if (not self.FieldExists('FileGrowth')) then self.AddChampSup('FileGrowth', false);
  self.PutValue('FileGrowth', valeur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Fonction cdgi de gestion de base
Suite ........ :
Suite ........ : Attention :
Suite ........ :   format JOBDATE (yyymmdd)
Suite ........ :   format JOBTIME (hhmm00)
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
class function cWABDSQLSERVER.Action(UneAction: string; RequestTOB: TOB; var ResponseTOB: TOB): boolean;
var
  (*  HeureDebut,DateDebut,Interval,TypeFreq,BaseOri,mdfFile,ldfFile,
    ServerName, ,BaseName,bakFile,JobName,wSql, job_id,IniName,Group: string ;*)
  UserName, Password,nomBase: string;
  tmpTOB, T1: TOB;
  autoclose, all: boolean;
  ioWABDSQLSERVER: cWABDSQLSERVER;
begin
  result := true;
  all := true;   autoclose:=false ;
  ioWABDSQLSERVER := nil;
  T1 := nil;
  TraceExecution('dbssPGI.cWABDSQLSERVER.' + UneAction + ' :');
  try
    ioWABDSQLSERVER := cWABDSQLSERVER(cWABD(RequestTOB).clone);

    UneAction := UpperCase(UneAction);
    tmpTob := RecupereUneTob(RequestTOB, 'OPTIONS');
    if (tmpTob <> nil) and (tmpTob.detail.count > 0) then
    begin
      if tmpTob.detail[0].FieldExists('UserName') then UserName := tmpTob.detail[0].GetValue('UserName');
      if tmpTob.detail[0].FieldExists('Password') then Password := tmpTob.detail[0].GetValue('Password');
      if tmpTob.detail[0].FieldExists('all') then all := tmpTob.detail[0].GetValue('all');
      if tmpTob.detail[0].FieldExists('autoclose') then autoclose := tmpTob.detail[0].GetValue('autoclose');
      if tmpTob.detail[0].FieldExists('nomBase') then nomBase := tmpTob.detail[0].GetValue('nomBase');
      tmpTob.free;
    end;

    with ioWABDSQLSERVER do
    begin
      TraceExecution('*=>fichier:' + fichier);
      TraceExecution('*=>BackUpFile:' + BackUpFile);

      try
        if UneAction = 'ADDLOGIN' then result := AjouterLogin(UserName, Password)
        else if UneAction = 'CEGIDSA' then result := ModifierLoginCEGIDSA()
        else if UneAction = 'REPARE' then result := ReparerDatabase()
        else if UneAction = 'EXISTDB' then result := ExisteBase()
        else if UneAction = 'ADDJOB' then result := AjouterJob()
        else if UneAction = 'INFOSJOB' then T1 := GetInfosJob()
        else if UneAction = 'DELETEJOB' then result := SupprimerJob()
        else if UneAction = 'LISTDATABASE' then T1 := ListerDatabase(All)
        else if UneAction = 'LISTUSER' then T1 := ListerUser()
        else if UneAction = 'LISTJOB' then T1 := ListerJob()
        else if UneAction = 'LISTTABLE' then T1 := ListerTables()
        else if UneAction = 'AUTOCLOSEBASE' then result := autoclosebase(autoclose)
        else if UneAction = 'NEW_BACKUP' then T1 := NEW_BACKUP(nombase)
        ;
      except
        on E: Exception do
          ErrorMessage := E.Message;
      end;
      setRetour(ResponseTOB, result, T1, nil, '');
    end;
  finally
    if T1 <> nil then T1.free;
    if ioWABDSQLSERVER <> nil then ioWABDSQLSERVER.free;
  end;
end;
{$ENDIF EAGLCLIENT}


end.



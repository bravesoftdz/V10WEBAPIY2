{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/11/2001
Modifié le ... : 02/11/2001
Description .. : Source TOF de la FICHE : MFOGENEREDBFORM ()
Mots clefs ... : TOF;MFOGENEREDBFORM
*****************************************************************}
unit MFOGENEREDBFORM_TOF;

interface

uses
  Classes, forms, sysutils, FileCtrl, Controls,
  {$IFNDEF EAGLCLIENT}
  dbtables, MajTable,
  {$ENDIF}
  HEnt1, HCtrls, HMsgBox, UTOF;

type
  TOF_MFOGENEREDBFORM = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    SocForm: THValComboBox; // choix de la base de formation
    procedure MessTraitement(Indice: Integer; Fini: Boolean);
    procedure SauvegardeDB;
    function RestaureDB(var Base: string): Boolean;
    procedure MajDBFormation(Base: string);
  end;

implementation

///////////////////////////////////////////////////////////////////////////////////////
//  FOConnecteDB : connexion à une base autre que la base de travail (cf. kb_assis.pas)
///////////////////////////////////////////////////////////////////////////////////////

function FOConnecteDB(DBName, BaseName: string): TDatabase;
var OldDriver: TDBDriver;
  Ok: Boolean;
begin
  OldDriver := V_PGI.Driver;
  Result := nil;
  try
    Application.ProcessMessages;
    Ok := ConnecteDB(DBName, Result, BaseName);
    Application.ProcessMessages;
  finally
    V_PGI.Driver := OldDriver;
  end;
  if not Ok then PgiBox('Impossible de se connecter à la société ' + DBName, BaseName);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOExtractDir : extrait le répertoire de stockage des bases de données
///////////////////////////////////////////////////////////////////////////////////////

function FOExtractDir(FileName, BaseName: string): string;
begin
  Result := ExtractFilePath(FileName);
  if Copy(Result, Length(Result), 1) = '\' then Delete(Result, Length(Result), 1);
  while Copy(Result, Length(Result), 1) <> '\' do Delete(Result, Length(Result), 1);
  if Result <> '' then
  begin
    Result := Result + 'D' + BaseName;
    //if not FileExists(Result) then CreateDir(Result) ;
    if not DirectoryExists(Result) then ForceDirectories(Result);
    Result := IncludeTrailingBackslash(Result);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  FOCheminSauve : indique le répertoire de stockage de la sauvegarde
///////////////////////////////////////////////////////////////////////////////////////

function FOCheminSauve: string;
begin
  Result := V_PGI.CheminSauve;
  if Result = '' then Result := 'C:\PGI00\SAUVE';
  //if not FileExists(Result) then CreateDir(Result) ;
  if not DirectoryExists(Result) then ForceDirectories(Result);
  Result := IncludeTrailingBackslash(Result);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MessTraitement : affichage d'un message d'état du traitement
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.MessTraitement(Indice: Integer; Fini: Boolean);
var Libelle: string;
begin
  if Fini then Libelle := 'Ok' else Libelle := 'En cours...';
  SetControlText('ETAT' + IntToStr(Indice), TraduireMemoire(Libelle));
  Application.ProcessMessages;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  SauvegardeDB : Sauvegarde de la base courante
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.SauvegardeDB;
var DBName, BKUPName, Sql: string;
begin
  DBName := GetDBPathName(True);
  BKUPName := 'N''' + FOCheminSauve + DBName + '.bak''';
  Sql := 'BACKUP DATABASE [' + DBName + ']'
    + ' TO  DISK = ' + BKUPName
    + ' WITH  INIT,  NOUNLOAD,  NAME = N''Backup ' + DBName + ''',  NOSKIP,  STATS = 10,  NOFORMAT';
  MessTraitement(1, False);
  ExecuteSQLDB(Sql, DBSoc, True);
  MessTraitement(1, True);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  RestaureDB : Sauvegarde de la base Formation
///////////////////////////////////////////////////////////////////////////////////////

function TOF_MFOGENEREDBFORM.RestaureDB(var Base: string): Boolean;
var QQ, QF: TQuery;
  NbMdf, NbLdf: Integer;
  DBName, BKUPName, FormDBName, FName, UName, PName, Sql: string;
  sDriver, sServer, sPath, sUser, sPassWord, sODBC, sOptions, sGroup: string;
begin
  Result := False;
  FormDBName := SocForm.Text;
  DBName := GetDBPathName(True);
  BKUPName := 'N''' + FOCheminSauve + DBName + '.bak''';
  // Fabrication de l'ordre de restauration de la sauvegarde dans la base Formation
  // Liste des fichiers de la base de données
  Sql := 'select cast (name as char(128)) as NAME, groupid as GROUPID from sysfiles';
  QQ := OpenSQL(Sql, True);
  if not QQ.EOF then
  begin
    // Paramètres de la base Formation
    ChargeDBParams(FormDBName, sDriver, sServer, sPath, Base, sUser, sPassWord, sODBC, sOptions, sGroup);
    // Recherche du répertoire de la base Formation
    PName := '';
    Sql := 'select cast (filename as char(255)) as FILENAME from master..sysdatabases'
      + ' where name = "' + Base + '"';
    QF := OpenSQL(Sql, True);
    if not QF.EOF then
    begin
      FName := Trim(QF.FindField('FILENAME').AsString);
      PName := IncludeTrailingBackslash(ExtractFilePath(FName));
    end else
    begin
      Ferme(QF);
      // Recherche du répertoire de la base courante
      Sql := 'select cast (filename as char(255)) as FILENAME from master..sysdatabases'
        + ' where name = "' + DBName + '"';
      QF := OpenSQL(Sql, True);
      if not QF.EOF then
      begin
        FName := Trim(QF.FindField('FILENAME').AsString);
        PName := FOExtractDir(FName, FormDBName);
      end;
    end;
    Ferme(QF);
    if PName = '' then
    begin
      PGIBox('La base de Formation n''est pas correctement définie.', Ecran.Caption);
      Ferme(QQ);
      Exit;
    end;
    NbMdf := 0;
    NbLdf := 0;
    Sql := 'RESTORE DATABASE [' + Base + '] FROM  DISK = ' + BKUPName
      + ' WITH  FILE = 1,  NOUNLOAD,  STATS = 10,  RECOVERY,  REPLACE';
    while not QQ.EOF do
    begin
      UName := Trim(QQ.FindField('NAME').AsString);
      FName := PName + Base;
      if QQ.FindField('GROUPID').AsInteger = 0 then
      begin
        // fichier de LOG
        if NbLdf > 0 then FName := FName + IntToStr(NbLdf);
        FName := FName + '.LDF';
        Inc(NbLdf);
      end else
      begin
        // fichier de données
        if NbMdf > 0 then FName := FName + IntToStr(NbMdf);
        FName := FName + '.MDF';
        Inc(NbMdf);
      end;
      Sql := Sql + ', MOVE N''' + UName + ''' TO N''' + FName + '''';
      QQ.Next;
    end;
  end else Sql := '';
  Ferme(QQ);
  if Sql <> '' then
  begin
    // Restauration de la sauvegarde dans la base Formation
    MessTraitement(2, False);
    ExecuteSQLDB(Sql, DBSoc, True);
    MessTraitement(2, True);
    Result := True;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  MajDBFormation : mise à jour de la base Formation
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.MajDBFormation(Base: string);
var DBForm: TDatabase;
  Sql, FormDBName: string;
begin
  // Ouverture de la base Formation
  FormDBName := SocForm.Text;
  DBForm := FOConnecteDB(FormDBName, Base);
  if DBForm <> nil then
  begin
    MessTraitement(3, False);
    // mise à jour des paramètres société de la base Formation
    Sql := 'update PARAMSOC set SOC_DATA="' + TraduireMemoire('BASE DE FORMATION') + '"'
      + ' where SOC_NOM="SO_LIBELLE"';
    ExecuteSQLDB(Sql, DBForm, False);
    Sql := 'update PARAMSOC set SOC_DATA="" where SOC_NOM="SO_GCFOCAISREFTOX"';
    ExecuteSQLDB(Sql, DBForm, False);
    Sql := 'update PARAMSOC set SOC_DATA="X" where SOC_NOM="SO_GCFOBASEFORMATION"';
    ExecuteSQLDB(Sql, DBForm, False);
    // libération des utilisateurs
    Sql := 'update UTILISAT set US_PRESENT="-" where US_PRESENT="X"';
    ExecuteSQLDB(Sql, DBForm, False);
    // modification du paramétrage des caisses dans la base Formation
    Sql := 'update PARCAISSE set GPK_INUSE="-", GPK_AFFINVERSE="-", GPK_TPETYPE="",'
      + ' GPK_TPEPORT="", GPK_TPEBAUDS="", GPK_TIROIRTYPE="", GPK_TIROIRPIN5="-",'
      + ' GPK_AFFMESG="*** ' + TraduireMemoire('BASE DE FORMATION') + '***"';
    ExecuteSQLDB(Sql, DBForm, False);
    // Remise à blanc des tables de la Tox dans la base Formation
    ExecuteSQLDB('delete STOXCHRONO', DBForm, False);
    ExecuteSQLDB('delete STOXEVENTS', DBForm, False);
    ExecuteSQLDB('delete STOXGSITES', DBForm, False);
    ExecuteSQLDB('delete STOXINFOCOMP', DBForm, False);
    ExecuteSQLDB('delete STOXPARMS', DBForm, False);
    ExecuteSQLDB('delete STOXQUERYS', DBForm, False);
    ExecuteSQLDB('delete STOXSITES', DBForm, False);
    ExecuteSQLDB('delete STOXVARS', DBForm, False);
    // Fermeture de la base Formation
    DBForm.Connected := False;
    DBForm.Free;
    MessTraitement(3, True);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnNew :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnNew;
begin
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnDelete :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnDelete;
begin
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnUpdate : constitution d'une base formation
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnUpdate;
var FormDBName, sBase: string;
begin
  inherited;
  FormDBName := SocForm.Text;
  if PGIAsk('Confirmez-vous ce traitement sur la société ' + FormDBName + ' ?', Ecran.Caption) = mrNo then Exit;
  SetControlVisible('ENCOURS', True);
  SetControlEnabled('SOCFORM', False);
  SaveSynRegKey('SOCIETEFORMATION', FormDBName, True);
  // Sauvegarde de la base en cours
  SauvegardeDB;
  // Restauration de la sauvegarde dans la base Formation et mise à jour de la base Formation
  if RestaureDB(sBase) then MajDBFormation(sBase);
  SetControlVisible('ENCOURS', False);
  SetControlEnabled('BVALIDER', False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnLoad :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnLoad;
var Ind: Integer;
  Stg: string;
begin
  inherited;
  // Chargement de la liste des sociétès
  SocForm := THValComboBox(GetControl('SOCFORM'));
  ChargeDossier(SocForm.Items, True);
  Ind := SocForm.Items.IndexOf(V_PGI.CurrentAlias);
  if Ind > -1 then SocForm.Items.Delete(Ind);
  Stg := '';
  Stg := GetSynRegKey('SOCIETEFORMATION', Stg, True);
  Ind := SocForm.Items.IndexOf(Stg);
  if Ind > -1 then SocForm.ItemIndex := Ind else SocForm.ItemIndex := 0;
  for Ind := 1 to 3 do SetControlText('ETAT' + IntToStr(Ind), '');
  SetControlVisible('ENCOURS', False);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnArgument :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnArgument(S: string);
begin
  inherited;
  // Uniquement pour les bases MSDE
  if V_PGI.Driver <> dbMSSQL then
  begin
    PGIBox('Le type de base de données ne permet pas l''utilisation de cette fonction.', Ecran.Caption);
    Close;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnClose :
///////////////////////////////////////////////////////////////////////////////////////

procedure TOF_MFOGENEREDBFORM.OnClose;
begin
  inherited;
end;

initialization
  registerclasses([TOF_MFOGENEREDBFORM]);
end.

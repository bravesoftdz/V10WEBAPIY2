unit galSystem;

interface

uses
  Windows, Messages, SysUtils, FileCtrl, Forms, Hent1, HCtrls, HMsgBox,
  UTob,
{$IFDEF VER150}
  Variants,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  Db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
//  uWABDSqlServer,
  Classes, StdCtrls, UTOZ, Dialogs;

type
  // TObject au lieu de Record car pb pour passer tableau dyn. en param de ChargeTabEspaces
  TRecEspace = class (TObject)
    Requis: Int64;  // taille des dossiers sélectionnés sur 1 disque
    Libre : Int64;  // espace libre sur 1 disque
    Total : Int64;  // espace total du disque
  end;

  TRegProc = function : HResult; Stdcall;

procedure TailleRepertoire (const APath: String; var ASize: Int64; lemasque: String='*.*');
function  SomeFilesIn (ladir: String): Boolean;
function  WaitWithMessageLoopUntil (hEvent : thandle; timeout: Integer; var Stop: Boolean) : boolean;
function  FileExecAndWaitUntil (sCommandLine  : string; timeout: Integer; var Stop: Boolean) : boolean;
function  IsValidHeureMinute (const hr: String): Boolean;
procedure OuvreLog (var chemlog: String; var FLog: TextFile; bIncrementer: Boolean=False);
procedure FermeLog (var FLog: TextFile);
function  GetVersionMSSQL: String;
function  PremierFichier (ladir, debutfichier, ext: String): String;
function  CopyDirectorySauf (laSource, laDestination, leMasque, saufext: String;
                             bWithSubDir, bFailIfExist: Boolean; saufterm: String=''; RemoveRO: Boolean=False): Boolean;
function  LectureSeule ( DirPath : String ) : Boolean ;
function  CopyDirectory2 (laSource, laDestination, leMasque: String; bWithSubDir,
                          bFailIfExist: Boolean; RemoveRO: Boolean): Boolean;
function  DeleteFileRO (lefichier: String; RemoveRO: Boolean=False): Boolean;
                                                                                                // ajout me 23/11/2005
function  RemoveInDir2 (laDir: String; bWithRoot,bWithSubDir: Boolean; RemoveRO: Boolean=False; Ext:string='\*.*'): Boolean;
procedure RetireRO (lefichier: String);
function  SeTerminePar (lefichier, lafin: String): Boolean;
function  GetSynRegKey_CB  (nomcle: String; user: Boolean): Boolean;
procedure SaveSynRegKey_CB (nomcle: String; vbool: Boolean; user: Boolean);
function  FicheEstChargee (NomFiche: String): Boolean;
function  RepertoireAuDessus(rep: String): String;
function  IsDossierUtilise (nodossier: string): Boolean;

function  CreeBaseSql              (NomBase, FichierMdf, FichierLdf, Accroissement: String): Boolean;
function  AttacherBaseSqlAutoclose (NomBase, FichierMdf, FichierLdf: String): Boolean;
function  SupprimeBaseSql          (NomBase: String): Boolean;
function  DBExists                 (NomBase:string):boolean;

{$IFDEF EAGLCLIENT}
{$ELSE}
function  DBTablePhysExiste (nomtbl: String; DB: TDatabase): Boolean;
{$IFNDEF EAGLSERVER}
function  ConnectDBAccess ( sAliasName_p, sDatabaseName_p, sFileName_p : string;
                          sUser_p : string = ''; sPassword_p : string = '';
                          bLoginPrompt_p : boolean = false ) : TDatabase;
function  DeconnectDBAccess (DBBase_p : TDatabase) : boolean;
{$ENDIF EAGLSERVER}

function  OpenSQLDBAvecTry (DB: TDatabase; SQL : String ; RO : Boolean; bNoChangeSQL: Boolean=False) : TQuery ;
function  ExecScriptDB(DB: TDatabase; Script: TStringList; ShowErr: Boolean = True): Boolean;
function  ExecCmdShell (Quoi: String): Boolean;
// $$$ JP 17/01/07: apparement obsolète function  ExisteSQLDB(DB: TDatabase; SQL : String) : boolean;

{$IFNDEF EAGLSERVER}
procedure ShrinkDatabaseMsSql(nombase: String; DB: TDatabase = Nil);
function  GetNbPartages: Integer;
procedure ChargeTabEspaces (lAide: THLabel; lsResultat: TListBox; gdEspace: THGrid;
                            LstDoss, LstDisq, RepLocaux: TStringList; var TabEspace: TList;
                            var BadEspace: Boolean; traitmt: String);
                            // impossible de passer un tableau dynamique en param
                            // donc pas de var TabEspace: Array of TRecEspace
{$ENDIF EAGLSERVER}
function  GetCollationBase (strBase: String): String;

{$ENDIF EAGLCLIENT}

{$IFDEF EAGL}
function  IsConnexionSSO(): Boolean;
{$ENDIF EAGL}

// Gestion Fichiers Zip classique pour les sauvegardes dossiers
function  VireExtension(Fichier: String): String;
function  InitialiseZip(CheminZip: String): TOZ;
function  FinaliseZip(var zip: TOZ): Boolean;
function  AjouteAuZip(var zip: TOZ; chemin: String): Boolean;
function  ZippeDirectory(var zip: TOZ; laSource, leMasque: String; bWithSubDir, bFailIfExist: Boolean; Exclure : string=''): Boolean;

procedure ChercheUnFichierDans(root, nomfichier: String; var resultat: String; niveau: Integer; var nivtrouve: Integer);
procedure ChercheUnRepertDans (root, nomrepert: String; var resultat: String; niveau: Integer; var nivtrouve: Integer);
procedure EnregistrerDll (sNomFichier : String);
function  IsProcessLoaded (const sModuleNames : String) : Boolean;
function  GetFileVersion (strFile:string):string; // $$$ JP 11/10/06

function  GetFullRegKey (const RegValueName : String ) : String ;
function  AffecteGroupe (nodoss:string; strGroupeConf:string) : string;
function  IsNoDossierOk(stNoDossier: string): boolean;
function  JaiBloque(Niveau: string): boolean;
function ISTablePhysExiste (NomTbl : string) : Boolean;



///////////// IMPLEMENTATION ///////////
implementation

uses galFileTools, Psapi;

{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 01/06/2001
Modifié le ... :   /  /
Description .. : Calcule l'espace occupé par un répertoire
Suite ........ : et ses sous-répertoires, en octets.
Mots clefs ... : SIZE;DIR;SPACE;ESPACE;REPERTOIRE
*****************************************************************}
procedure TailleRepertoire(const APath: String; var ASize: Int64; lemasque: String='*.*');
var SearchRec: TSearchRec;
begin
  if not DirectoryExists(APath) then exit;
  FindFirst(APath+'\'+lemasque, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
      begin
      if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          TailleRepertoire(APath+'\'+SearchRec.Name, ASize);
        end
      else
        begin
        ASize := ASize + SearchRec.Size;
        end;
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
end;


function SomeFilesIn(ladir: String): Boolean;
var SearchRec: TSearchRec;
begin
  Result := False;
  if not DirectoryExists(ladir) then exit;
  FindFirst(ladir+'\*.*', faAnyFile, SearchRec);
  repeat
    If SearchRec.Name<>'' then
      BEGIN
      if ((SearchRec.Attr and SysUtils.faDirectory)=0)
       or ((SearchRec.Name<>'.') and (SearchRec.Name<>'..')) then
        begin Result := True; break; end;
      END;
  until FindNext(SearchRec)<>0 ;
  FindClose(SearchRec);
end;


function WaitWithMessageLoopUntil(hEvent : thandle; timeout: Integer; var Stop: Boolean) : boolean;
// Copie de WaitWithMessageLoop pour rajout timeout
var
  msg : TMSG;
  dwRet : DWORD;
  depart : TDateTime;
begin
  result := false;
  depart := Now();
  while true do
    begin
    dwRet := MsgWaitForMultipleObjects( 1,    // One event to wait for
                     hEvent,        // The array of events
                     FALSE,          // Wait for 1 event
                     INFINITE,       // Timeout value
                     QS_ALLINPUT);   // Any message wakes up
    if dwRet = WAIT_OBJECT_0 then
      begin
      result := true; // The event was signaled, return
      break;
      end
    else if dwRet = WAIT_OBJECT_0 + 1 then
      begin
          // There is a window message available. Dispatch it.
      while PeekMessage(msg,0,0,0,PM_REMOVE) do
        begin
        Application.ProcessMessages; // pour appui sur stop
        if Stop then exit; // SORTIE
        TranslateMessage(msg);
        DispatchMessage(msg);
        // TIMEOUT en MINUTES
        if (Now-Depart)*24*60>timeout then exit; // SORTIE
        end;
      end
    else
      begin
      result := false;
      break;
      end;
    Application.ProcessMessages; // pour appui sur stop
    if Stop then exit; // SORTIE
    end;
end;


function FileExecAndWaitUntil(sCommandLine  : string; timeout: Integer; var Stop: Boolean) : boolean;
// Copie de FileExecAndWait pour rajout d'un timeout (+ stop)
var
  tsi           : TStartupInfo;
  tpi           : TProcessInformation;
begin
  result := false;
  FillChar(tsi, SizeOf(TStartupInfo), 0);
  tsi.cb := SizeOf(TStartupInfo);
  if CreateProcess( nil, //PChar(sExe),
                    PChar(sCommandLine),
                    nil, nil, False, 0, nil, nil, tsi, tpi) then
  begin
    result := WaitWithMessageLoopUntil(tpi.hProcess, timeout, Stop);
    CloseHandle(tpi.hProcess);
    CloseHandle(tpi.hThread);
  end;

end;


function IsValidHeureMinute(const hr: String): Boolean;
var tmp: String;
begin
  Result := False;
  tmp := Trim(hr);
  // date vide
  if (tmp='') or (tmp=':') or (tmp='.') then exit;
  tmp := TimeToStr(StrToTime(hr));
  if tmp='' then exit;
  Result := True;
end;


procedure OuvreLog(var chemlog: String; var FLog: TextFile; bIncrementer: Boolean=False);
var i: Integer;
    racine, extension, newlog : String;
begin
  If FileExists(chemlog) then
    begin
    // On remplace le fichier
    if Not bIncrementer then
      DeleteFile(chemlog)
    else
    // On génère (et retourne) un nom de fichier incrémenté
      begin
      racine := VireExtension(chemlog);
      // contient le . de l'extension !
      extension := ExtractFileExt(chemlog);
      i := 0;
      newlog := chemlog;
      While FileExists(newlog) do
        begin
        i := i + 1;
        newlog := racine + '_' + FormatDateTime('yyyymmdd', Date()) + '_' + IntToStr(i) + extension;
        end;
      chemlog := newlog;
      end;
    end;

  For i:=1 to 1000 do Application.ProcessMessages;
  try
    AssignFile(FLog, chemlog);
    // FileMode := 2; inutile, c'est déjà le cas
    Rewrite(FLog)
    // Rq : pour compléter un fichier existant : Append(FLog);
  except
    on E:Exception do
      begin
      PGIInfo(E.Message+#13+#10+'Impossible de créer le journal '+chemlog,TitreHalley) ;
      try CloseFile(FLog) Except end;
      end;
  end;
end;


procedure FermeLog(var FLog: TextFile);
begin
  try
    CloseFile(FLog);
  except
  end;
end;


function GetVersionMSSQL: String;
// retourne '', 'SQL7', 'SQL2000'
// MD  18/05/06 : maintenant on retourne '', 'SQL 7', 'SQL2000', 'SQL2005'
// RP0 06/02/08 : maintenant on retourne '', 'SQL 7', 'SQL2000', 'SQL2005', 'SQL2008'
var
    Q, Q2 : TQuery;
begin
  Result := '';
  if (V_PGI.Driver<>dbMSSQL) and (V_PGI.Driver<>dbMSSQL2005)  then exit;
  try
    Q := OpenSQL('select cast(@@Version as varchar(30))', True);
    if Not Q.Eof then
    begin
      { SQL 7 :
        Microsoft SQL Server  7.00 - 7
        SQL 2000 :
        Microsoft SQL Server  2000 - 8
        SQL 2005 :
        Microsoft SQL Server 2005 - 9.
        SQL 2008 :
        Microsoft SQL Server code name "Katmai" (CTP) - 10.
        Pour l'instant @@Version renvoie un nom de code, donc dans le doute, on fait un serverproperty("ProductVersion"),
        qui renvoie le numero de version seul, qui lui sera toujours de la forme 10.x.xxxx.x, même plus tard
      }
      if Pos('7.00', Q.Fields[0].AsString)>0 then
        Result := 'SQL 7' // espace important pour les comparaisons de chaines 'SQL 7' < 'SQL2000'
      else if Pos('2000', Q.Fields[0].AsString)>0 then
        Result := 'SQL2000'
      else if Pos('2005', Q.Fields[0].AsString)>0 then
        Result := 'SQL2005'
      else
        try
          Q2 := OpenSQL('select convert(varchar(3), serverproperty("ProductVersion"))', True);
          if Not Q2.Eof then
              if Pos('10.', Q2.Fields[0].AsString)>0 then Result := 'SQL2008';
        finally
          Ferme(Q2);
        end;
    end;
  finally
    Ferme(Q);
  end;
{
  #### Rq : code ci-dessous ne marche pas en sql7 car serverproperty non implémentée
  SQL := 'select convert(varchar(128), serverproperty("ProductVersion"))';
}
end;


function PremierFichier(ladir, debutfichier, ext: String): String;
// retourne le chemin du 1er fichier trouvé
// avec le nom [débutfichier], l'extension [ext], dans le répertoire [racine]
var
  SearchRec: TSearchRec;
begin
  Result := '';
  if not DirectoryExists(ladir) then exit;

  FindFirst(ladir+'\'+debutfichier+'*.'+ext, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') and (SearchRec.Attr and faDirectory = 0) then
      begin
      Result := laDir+'\'+SearchRec.Name;
      break; // Fin de boucle
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
end;


function CopyDirectorySauf(laSource, laDestination, leMasque, saufext: String;
  bWithSubDir, bFailIfExist: Boolean; saufterm: String=''; RemoveRO: Boolean=False): Boolean;
// saufext : fournir les éventuelles extensions des fichiers à ne pas copier, séparées par des ;
//           (ne pas mettre le . devant chaque extension)
// saufterm : ne copie pas un fichier si il se termine par ces chaines
//           (extension comprise, séparées par des ;)
var SearchRec: TSearchRec;
    strDe, strVers, ext, sauflesext, badext, sauflesterm, badterm: string;
    bad : Boolean;
begin
  Result := False;
  if not DirectoryExists(laSource) then exit;
  if not DirectoryExists(laDestination) then CreateDir(laDestination);
  FindFirst(laSource+'\'+lemasque, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
      begin
      strDe := laSource+'\'+SearchRec.Name;
      strVers := laDestination+'\'+SearchRec.Name;
      if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if (bWithSubDir) and (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          begin
          // CopyDirectory2(strDe,strVers,leMasque,bWithSubDir,bFailIfExist, RemoveRO);
          // => pas bon ! ne tenait pas compte des options saufext, saufterm...
          // => il faut faire un appel récursif :
          CopyDirectorySauf(strDe, strVers, leMasque, saufext, bWithSubDir, bFailIfExist, saufterm, RemoveRO);
          if RemoveRO then RetireRO(strVers);
          end;
        end
      else
        begin
        ext := Copy(strDe, Length(strDe)-2, 3);
        sauflesext := UpperCase(saufext);
        bad := False;
        // élimine fichiers ayant telle extension...
        While sauflesext<>'' do
          begin
          badext :=ReadTokenSt(sauflesext);
          if (badext<>'') and (UpperCase(ext)=badext) then
            begin bad := True; break; end;
          end;
        if bad then Continue;
        // élimine fichiers terminés par ...
        sauflesterm := UpperCase(saufterm);
        While sauflesterm<>'' do
          begin
          badterm := ReadTokenSt(sauflesterm);
          if (badterm<>'') and SeTerminePar(strDe, badterm) then
            begin bad := True; break; end;
          end;
        if bad then Continue;
        // copie d'un fichier
        if CopyFile(PChar(strDe),PChar(strVers), bFailIfExist) and RemoveRO then
          RetireRO(strVers);
        end;
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
  Result := True;
end;


function LectureSeule( DirPath : String ) : Boolean ;
// #### Attente dispo PathIsReadOnly de E.P. dans AGL > 530e
var
  f              :     file;
  UniqueFName    :     String ;
  oldDir         : array [0..MAX_PATH] of Char;
begin
  Result         := True ;

  DirPath      := trim( DirPath );
  if DirPath = '' then DirPath := GetCurrentDir ;
  if length( DirPath ) = 1 then DirPath := DirPath + ':\' ;
  if DirPath[ length( DirPath ) ] <> '\' then DirPath := DirPath + '\' ;

  randomize;
  repeat
    UniqueFName := IntToHex( random( $FFFF ) , 8 ) + '.tmp' ;
  until not FileExists( DirPath + UniqueFName );

  // tester si le lecteur existe
  GetCurrentDirectory(MAX_PATH,oldDir);
  {$I-}
  ChDir(DirPath);
  if IOResult <> 0 then exit;
  ChDir(OldDir);
  {$I+}

  // tester si on peut ecrire
  Assign( f , DirPath + UniqueFName );
  try
    Rewrite( f );
    if IOResult = 0 then begin
      Close( f );
      DeleteFile( DirPath + UniqueFName ) ;
      Result := False ;
    end;
  except
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 28/09/2002
Modifié le ... : 
Description .. : Depuis, il existe CopyDirectory dans l'AGL (Hent1)
Mots clefs ... : 
*****************************************************************}
function CopyDirectory2(laSource, laDestination, leMasque: String; bWithSubDir,
  bFailIfExist: Boolean; RemoveRO: Boolean): Boolean;
var SearchRec: TSearchRec;
    strDe, strVers : string;
begin
  Result := False;
  if not DirectoryExists(laSource) then exit;
  if not DirectoryExists(laDestination) then CreateDir(laDestination);
  FindFirst(laSource+'\'+lemasque, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
      begin
      strDe := laSource+'\'+SearchRec.Name;
      strVers := laDestination+'\'+SearchRec.Name;
      if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if (bWithSubDir) and (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          CopyDirectory2(strDe,strVers,leMasque,bWithSubDir,bFailIfExist, RemoveRO);
        end
      else
        begin
        If CopyFile(PChar(strDe),PChar(strVers), bFailIfExist) and RemoveRO then
          RetireRO(strVers);
        end;
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
  Result := True;
end;


function DeleteFileRO(lefichier: String; RemoveRO: Boolean=False): Boolean;
// Suppression possible même si fichier en lecture seule
// True si suppr ok, False sinon
var Attrs: Integer;
begin
  Result := True;
  If Not FileExists(lefichier) then exit;
  If RemoveRO then
    begin
    Attrs := FileGetAttr(lefichier);
    // SysUtils. permet de compiler même si on uses l'unité DB.pas
    if (Attrs and SysUtils.faReadOnly) <> 0 then
      FileSetAttr(lefichier, Attrs - SysUtils.faReadOnly);
    end;
  Result := DeleteFile(lefichier);
end;


function RemoveInDir2(laDir: String; bWithRoot,bWithSubDir: Boolean; RemoveRO: Boolean=False; Ext:string='\*.*'): Boolean;
// supprime tous les fichiers d'un répertoire
// avec possibilité de supprimer ceux en lecture seule
var SearchRec: TSearchRec;
    strDir : string;
begin
  Result := False;
  if not DirectoryExists(laDir) then exit;
  // ajout me 23/11/2005 Ext
  FindFirst(laDir+Ext, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
      begin
      strDir := laDir+'\'+SearchRec.Name;
      if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if bWithSubDir and (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          begin
          if RemoveRO then RetireRO(strDir);
          RemoveInDir2(strDir,True,bWithSubDir, RemoveRO);
          end;
        end
      else
        DeleteFileRO(PChar(strDir), RemoveRO);
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
  Result := True;
  // supprime le répertoire lui-même
  if bWithRoot then
    begin
    if RemoveRO then RetireRO(ladir);
    Result := RemoveDir(laDir);
    end;
end;


procedure RetireRO(lefichier: String);
// aucune vérif d'existence du répertoire ou du fichier
var Attrs: Word;
begin
  Attrs := FileGetAttr(lefichier);
  // SysUtils. permet de compiler même si on uses l'unité DB.pas
  if (Attrs and SysUtils.faReadOnly) <> 0 then
    FileSetAttr(lefichier, Attrs - SysUtils.faReadOnly);
end;


function SeTerminePar(lefichier, lafin: String): Boolean;
// indique si le fichier se termine précisément par lafin indiquée
// (extension comprise)
var lg: Integer;
begin
  Result := False;
  lg := Length(lefichier);
  if lg >= Length(lafin) then
    if UpperCase(Copy(lefichier, lg-Length(lafin)+1, Length(lafin)))=UpperCase(lafin) then
      Result := True;
end;


// aucune vérif d'existence du répertoire ou du fichier
procedure RemoveFileAttr(lefichier: String ; Attrib : integer);
var Attrs: Word;
begin
  Attrs := FileGetAttr(lefichier);
  if (Attrs and Attrib) > 0 then
    FileSetAttr(lefichier, Attrs - Attrib);
end;


function GetSynRegKey_CB(nomcle: String; user: Boolean): Boolean;
// récupère une valeur booléene (checkbox) dans la registry
// 1 = True, autres = False
begin
  Result := False;
  if GetSynRegKey(nomcle, '', user)='1' then Result := True;
end;


procedure SaveSynRegKey_CB(nomcle: String; vbool: Boolean; user: Boolean);
// sauve une valeur booléene dans la registry
// 1 = True, 0 = False
begin
  if vbool then SaveSynRegKey(nomcle, '1', user)
  else SaveSynRegKey(nomcle, '0', user);
end;


function FicheEstChargee(NomFiche: String): Boolean;
// teste si une fiche est chargée (même si en mode inside dans FMenuG !)
var i: Integer;
begin
  Result := False;
  for i:=0 to Screen.FormCount-1 do
    begin
    if UpperCase(Screen.Forms[i].Name)=UpperCase(NomFiche) then
      begin
      Result := True;
      Break;
      end;
    end;
end;


function RepertoireAuDessus(rep: String): String;
var tmp: String;
    i : Integer;
begin
  Result := '';
  tmp := rep;
  i := pos('\', tmp);
  While i>0 do
    begin
    if Result<>'' then Result := Result+'\';
    Result := Result + Copy(tmp, 1, i-1);
    tmp := Copy(tmp, i+1, Length(tmp)-i);
    i := pos('\', tmp);
    end;
end;


function IsDossierUtilise (nodossier: string): Boolean;
//Vérifie si une base est en cours d'utilisation (au sens SQL).
begin
  Result := False;
  if (V_PGI.Driver<>dbMSSQL) and (V_PGI.Driver<>dbMSSQL2005) then exit;
  Result := ExisteSQL('SELECT 1 FROM MASTER.DBO.SYSPROCESSES WHERE DBID=(SELECT DBID FROM MASTER.DBO.SYSDATABASES WHERE NAME="DB'+nodossier+'")');
end;


function CreeBaseSql (NomBase, FichierMdf, FichierLdf, Accroissement: String): Boolean;
begin
  (*
     with cWABDSQLSERVER.Create do
     begin
          Parle      := V_PGI.SAV; // évite le message "Base créée avec succès"
          Base       := NomBase;
          mdfFile    := FichierMdf;
          ldfFile    := FichierLdf;
          FileGrowth := Accroissement;
          Result     := CreerBase;
          Free;
     end;
  *)
end;


function AttacherBaseSqlAutoclose (NomBase, FichierMdf, FichierLdf: String): Boolean;
begin
  (*
     with cWABDSQLSERVER.Create do
     begin
          Parle   := V_PGI.SAV; // évite le message "Base attachée avec succès"
          Base    := NomBase;
          mdfFile := FichierMdf;
          ldfFile := FichierLdf;
          Result  := AttacherBase;
          // Tentative sans gérer le log si non trouvé
          // Permet aussi d'attacher une base avec _log.ldf
          if Result = FALSE then
          begin
            ldfFile := '';
            Result := AttacherBase;
          end;
          AutocloseBase(True);
          Free;
     end;
  *)
end;


function  SupprimeBaseSql (NomBase: String): Boolean;
begin
  (*
     with cWABDSQLSERVER.Create do
     begin
          Parle   := V_PGI.SAV; // évite le message "Base attachée avec succès"
          Base    := NomBase;
          Result  := SupprimerBase;
          Free;
     end;
  *)
end;


function DBExists (NomBase:string):boolean;
begin
     // Pour dépanner, avant AGL 7.0.0.9 :
     // Result := ExisteSQL('select 1 from master.dbo.sysdatabases where name="'+strBase+'"');
     // Exit;
     (*
     with cWABDSQLSERVER.Create do
     begin
          Parle        := V_PGI.SAV;
          Base         := NomBase;
          Result       := ExisteBase;
          Free;
     end;
     *)
end;

{$IFDEF EAGLCLIENT}
  // #### A FAIRE EAGL
{$ELSE}
function ISTablePhysExiste (NomTbl : string) : Boolean;
var Q: TQuery;
begin
  Result := False;
  Q := OpenSQL('select object_id('''+nomtbl+''') as resultat', True,-1,'',True,'',true);
  if Q<>Nil then
  begin
    if Not Q.Eof then
      Result := Not VarIsNull(Q.FindField('resultat').AsVariant);
    Ferme(Q);
  end;
end;

function DBTablePhysExiste(nomtbl: String; DB: TDatabase): Boolean;
// teste existence table physique directement dans la base
var Q: TQuery;
begin
  Result := False;
  if DB=Nil then exit;
  if Not DB.Connected then exit;
  Q := OpenSQLDBAvecTry(DB, 'select object_id('''+nomtbl+''') as resultat', True, True);
  if Q<>Nil then
    begin
    if Not Q.Eof then
      Result := Not VarIsNull(Q.FindField('resultat').AsVariant);
    Ferme(Q);
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 16/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
function ConnectDBAccess( sAliasName_p, sDatabaseName_p, sFileName_p : string;
                          sUser_p : string = ''; sPassword_p : string = '';
                          bLoginPrompt_p : boolean = false ) : TDatabase;
var
   DBBase_l : TDatabase;
begin
   DBBase_l := nil;
//   result := nil;
   try
      DBBase_l := TDatabase.Create(Appli);
      {$IFNDEF DBXPRESS}
      DBBase_l.AliasName := sAliasName_p;
      {$ENDIF}
      DBBase_l.DatabaseName := sDatabaseName_p;
      DBBase_l.DriverName := 'MSACCESS';
      DBBase_l.Params.Clear;
      DBBase_l.Params.Add('DATABASE NAME=' + sFileName_p);
      DBBase_l.Params.Add('USER=' + sUser_p);
      DBBase_l.Params.Add('PASSWORD=' + sPassword_p);
      DBBase_l.LoginPrompt := bLoginPrompt_p;
      DBBase_l.Connected := true;
   finally
      result := DBBase_l;
   end;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 16/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLSERVER}
function DeconnectDBAccess(DBBase_p : TDatabase) : boolean;
begin
//   result := false;
   try
      DBBase_p.Close;
      DBBase_p.Connected := false;
   finally
      result := true;
   end;
end;
{$ENDIF}

function OpenSQLDBAvecTry(DB: TDatabase; SQL : String ; RO : Boolean; bNoChangeSQL: Boolean=False) : TQuery ;
// copié de Opensql (HCtrls) le 22/02/02
// rajout try except pour autoriser test de tables inexistantes
// Attention : non testé avec des req du type nombase.dbo.nomtable
//  (voir si il manque des guillemets autour : "nombase.dbo.nomtable")
// #### Attention : obsolète avec ADO ou alors revoir le mode de connexion !
var Q : TQuery;
begin
  try
    // copié de Opensql (HCtrls) le 22/02/02
    Q:=TQuery.Create(Appli) ;
    Q.DatabaseName:=DB.DataBaseName ;
    Q.SessionName:=DB.SessionName ;
    Q.CacheBlobs:=FALSE ; //Simon 12/8/99
    Q.UpdateMode:=upWhereChanged ;
    Q.SQL.Clear ;
    Q.SQL.Add(SQL) ;
    Q.RequestLive:=Not RO ;
    //deconne avec ACCESS ! if UniDirection and RO then Q.UniDirectional:=TRUE ;
    If Not bNoChangeSQL then ChangeSQL(Q) ;
    Q.Open ;
    FetchSQLODBC(Q) ; //Simon 7/12/99
    Result:=Q ;
  except
    Result := Nil;
  end;
end;

function ExecScriptDB(DB: TDatabase; Script: TStringList; ShowErr: Boolean = True): Boolean;
// Exécution du script SQL sur le DB ou sur DBSOC en cours, mais sans aucune traduction du SQL
// (#### inspiré de ExecuteSQLDB, donc à maintenir...)
// ATTENTION aux scripts contenant des USE NOMBASE : cela change la base en cours !!
var
  Q: TQuery;
begin
  Result := False;
  Q := nil;
  try
    if (DB<>Nil) and (Not DB.Connected) then exit;
    Q := TQuery.Create(Appli);
    if DB=Nil then
      begin
{$IFNDEF DBXPRESS}
      Q.DatabaseName := DBSOC.DataBaseName;
{$ELSE}
      Q.SQLConnection := DBSOC;
{$ENDIF}
      Q.SessionName := DBSOC.SessionName;
      end
    else
      begin
{$IFNDEF DBXPRESS}
      Q.DatabaseName := DB.DataBaseName;
{$ELSE}
      Q.SQLConnection := DB;
{$ENDIF}
      Q.SessionName := DB.SessionName;
      end;

    Q.SQL.Clear;
    //PMJ 16/08/2005, avec ADO et UNICODE on est obligé de passer par les requêtes paramétrées
    Q.RequestLive := FALSE;
    // éxécution du script
    Q.SQL.Text := Script.Text;
    try
      Q.ExecSQL;
      Result := True;
    except
      on e: exception do
        if ShowErr then PgiError(E.Message);
    end;
  finally
    if Q <> nil then
    begin
      Q.Close;
      Q.Free;
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 20/06/2007
Modifié le ... :   /  /
Description .. : Exécution d'une commande shell sur serveur :
Suite ........ : elle peut contenir des caractères interprétés
Suite ........ : par le driver ( "  :  * )
Suite ........ : donc on passe par une mise à en paramètre
Mots clefs ... :
*****************************************************************}
function  ExecCmdShell (Quoi: String): Boolean;
var Q : TQuery;
begin
  Result := False;
  if (V_PGI.Driver<>dbMSSQL) and (V_PGI.Driver<>dbMSSQL2005) then exit;
  try
    Q := PrepareSQL('@@master.dbo.xp_cmdshell :QUOI',True) ;
    Q.ParamByName('QUOI').AsString:= Quoi;
    Q.ExecSQL;
    Result := True;
  except
    on E:Exception do PGIInfo(E.Message);
  end;
  if Q<>Nil then Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 17/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
// $$$ JP 17/01/07: apparement plus du tout utilisé
{function ExisteSQLDB(DB: TDatabase; SQL : String) : boolean;
var
   QRYRequete_l : TQuery;
begin
   try
      QRYRequete_l := TQuery.Create(Appli) ;
      QRYRequete_l.DatabaseName := DB.DataBaseName;
      QRYRequete_l.SessionName := DB.SessionName;
      QRYRequete_l.CacheBlobs := FALSE ;
      QRYRequete_l.UpdateMode := upWhereChanged;
      QRYRequete_l.SQL.Clear;
      QRYRequete_l.SQL.Add(SQL);
      QRYRequete_l.RequestLive := false;
      ChangeSQL(QRYRequete_l);
      QRYRequete_l.Open;
      FetchSQLODBC(QRYRequete_l);
      Ferme(QRYRequete_l);
      Result := true;
   except
      Ferme(QRYRequete_l);
      Result := false;
   end;
end;}


{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 10/02/2005
Modifié le ... :   /  /
Description .. : Récupération de l'espace inoccupé dans les fichiers
Suite ........ : de la base de données Sql Server (mdf/ldf)
Mots clefs ... :
*****************************************************************}
procedure  ShrinkDatabaseMsSql(NomBase: String; DB: TDatabase = Nil);
var
    Script: TStringList;
    BaseEnCours: String;
begin
    // Mémorise la base en cours sur la connexion
    If DB=Nil then
      BaseEnCours := V_PGI.DbName // ou DBSOC.Params.Values['DataBase'] (BDE : ['DATABASE NAME'])
    else
      BaseEnCours := DB.Params.Values['DataBase'];
    Script := TStringList.Create;
    Script.Add('use '+NomBase);
    Script.Add('checkpoint');
    Script.Add('backup log '+NomBase+' with no_log');
    Script.Add('dbcc shrinkdatabase('+NomBase+')');
    Script.Add('dbcc shrinkfile(2)'); // au lieu de (MODELE_PCL_S5_LOG),
    // car si base sql7 réparée, le nom du journal serait ##nodossier_log !
    // Restitue la base en cours
    Script.Add('use '+BaseEnCours);
    ExecScriptDB(DB, Script, False);
    Script.Free;
end;


function GetNbPartages: Integer;
var i: Integer;
begin
  // Result := V_PGI_Env.Partages.Count;
  Result := 0;
  for i := 1 to 8 do
    begin
    if V_PGI.Partages[i]='' then break;
    Result := i;
    end;
end;

procedure ChargeTabEspaces(lAide: THLabel; lsResultat: TListBox; gdEspace: THGrid;
  LstDoss, LstDisq, RepLocaux: TStringList; var TabEspace: TList; var BadEspace: Boolean;
  traitmt: String);
// calcul + affichage espace total, libre, et requis sur les partages disques disponibles
//  calculé par rapport à un accroissement (en dur) de 20% des fichiers mdf/ldf
//  pour les dossiers fournis dans la liste
// => convient pour la mise à jour par lot et la réintégration suite à transport
// => calcul espace requis dépend du traitmt :
//  'Mise à jour par lots' => ajoute 20% aux MDF+LDF des dossiers du serveur
//  'Réintégration' => prend l'espace total des rep dossiers locaux à réintégrer
var nodisq, i : Integer;
    nodoss, rep : String;
    Filetool : TFileTools;
    espace : Int64;
    esp : TRecEspace;
    NbPartages : Integer;
begin
  if TabEspace=Nil then exit;

  // CALCULS PREALABLES
  lAide.Caption :='Calcul de l''espace disponibles sur le(s) disque(s)...';
  NbPartages := GetNbPartages;
  for i:=0 to NbPartages-1 do
    begin
    rep := V_PGI.Partages[i+1]; // V_PGI_Env.Partages[i];
    esp := TRecEspace.create;
    esp.Requis := 0;
    esp.Libre := 0;
    esp.Total := 0;
    FileTool := TFileTools.Create;
    FileTool.LoadFileProp(rep);
    esp.Libre := (FileTool.FreeOnDisk Div 1024) div 1024;
    esp.Total := (FileTool.SizeOfDisk div 1024) div 1024;
    FileTool.Free;
    TabEspace.Add(esp);
    end;

  lAide.Caption :='Calcul de l''espace occupé par les dossiers sélectionnés...';
  // exploite liste des dossiers fournie
  for i:=0 to LstDoss.Count-1 do
    BEGIN
    nodoss := LstDoss[i];
    nodisq := StrToInt(LstDisq[i]);
    // nodisq démarre à 1
    if nodisq>0 then
      begin
      esp := TRecEspace(TabEspace[nodisq-1]);
      espace := 0;
      if traitmt='Mise à jour par lots' then
        begin
        // TailleRepertoire(V_PGI_Env.Partages[nodisq-1]+'\D'+nodoss, espace, '*.MDF');
        TailleRepertoire(V_PGI.Partages[nodisq]+'\D'+nodoss, espace, '*.MDF');
        esp.Requis := esp.Requis + round(0.20 * espace);
        espace := 0;
        // TailleRepertoire(V_PGI_Env.Partages[nodisq-1]+'\D'+nodoss, espace, '*.LDF');
        TailleRepertoire(V_PGI.Partages[nodisq]+'\D'+nodoss, espace, '*.LDF');
        esp.Requis := esp.Requis + round(0.20 * espace);
        end
      else if (traitmt='Réintégration') then
        begin
        if nodisq>RepLocaux.Count then
          PGIInfo('Le disque de données n° '+IntToStr(nodisq)+' n''est pas défini en local.', 'Calcul de l''espace disque nécessaire')
        else
          begin
          TailleRepertoire(RepLocaux[nodisq-1]+'\D'+nodoss, espace);
          esp.Requis := esp.Requis + espace;
          end;
        end;
      end
    else
      PGIInfo('Disque '+IntToStr(nodisq)+ ' du dossier '+nodoss
       +#13+#10+'inconnu dans la liste des partages.', TitreHalley);
    END;
  // Conversion Octets => Mo
  for i:=0 to NbPartages-1 do
    TRecEspace(TabEspace[i]).Requis := (TRecEspace(TabEspace[i]).Requis div 1024) div 1024;

  // Remplissage grille
  gdEspace.RowCount := NbPartages+1;
  lsResultat.Clear;
  for i:=1 to 3 do gdEspace.ColAligns[i] := taRightJustify;
  for i:=0 to NbPartages-1 do
    begin
    esp := TRecEspace(TabEspace[i]);
    // gdEspace.Cells[0, i+1] := V_PGI_Env.Partages[i];
    gdEspace.Cells[0, i+1] := V_PGI.Partages[i+1];
    // Taille
    gdEspace.Cells[1, i+1] := IntToStr(esp.Total)+' Mo';
    // Requis
    gdEspace.Cells[2, i+1] := IntToStr(esp.Requis)+' Mo';
    // Esp libre
    gdEspace.Cells[3, i+1] := IntToStr(esp.Libre)+' Mo';
    // Affichage
    if (esp.Requis > esp.Libre) then
      begin
      // lsResultat.Items.Add('Attention : l''espace libre sur '+V_PGI_Env.Partages[i]+' est insuffisant');
      lsResultat.Items.Add('Attention : l''espace libre sur '+V_PGI.Partages[i+1]+' ');
      lsResultat.Items.Add('est insuffisant pour la '+traitmt+' !');
      lsResultat.Items.Add('Veuillez d''abord déplacer certains dossiers vers un autre partage...');
      BadEspace := True;
      end;
    end;
  // message si ok
  if Not BadEspace then
    begin
    lsResultat.Items.Add('L''espace libre sur le(s) partage(s) est théoriquement suffisant ');
    lsResultat.Items.Add('pour la '+traitmt+'.');
    lsResultat.Items.Add('En cas de doute (limite proche), il est recommandé de faire d''abord');
    lsResultat.Items.Add('un déplacement de dossiers pour ne pas risquer la saturation...');
    end;
  lAide.Caption :='';
end;

{$ENDIF EAGLSERVER}
{$ENDIF EAGLCLIENT}


// $$$ JP 19/05/06 FQ 10616: en eagl, fonction CheckCollationBase dans le plugin: cegidpgiutil.pas
function GetCollationBase(strBase: String): String;
// sur sql2000, récupère la collation
var
    Q      : TQuery;
begin
  Result := '';
  if GetVersionMsSql<'SQL2000' then exit;

  // test existence
  if Not DBExists(strBase) then
    Result  := 'Base absente ou détachée !'
  else
    try
      // Nécessaire de mettre une clause from pour ouvrir la base qui est autoclose
      // sinon databasepropertyex renvoie Null
      Q := OpenSQL('SELECT CONVERT (VARCHAR (100), DATABASEPROPERTYEX ("' + strBase + '","collation")), fileid from ' +strBase+ '.dbo.sysfiles', True);
      If Not Q.Eof then Result := Q.Fields[0].AsString ;
      Ferme(Q);
    except
        on E : exception do
              {$IFDEF EAGLSERVER}
              ddWriteln('######## Erreur GetCollationBase : '+E.Message);
              {$ELSE}
              if V_PGI.SAV then PGIInfo(E.Message, TitreHalley) ;
              {$ENDIF}
    end;
end;

{$IFDEF EAGL}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : SL
Créé le ...... : 11/04/2008
Modifié le ... :   /  /    
Description .. : Indique si la connexion Web Access s'est faite via une 
Suite ........ : authentification Single Sign-On
Suite ........ : ###ATTENTION : L'authentification forte va écrire dans V_PGI.DomainName
Suite ........ : =>A prendre en compte
Mots clefs ... :
*****************************************************************}
function  IsConnexionSSO(): Boolean;
var
  DomainName:       String;
begin
  Result := False;
  DomainName := V_PGI.DomainName;
  if DomainName <> '' then
    Result := True;
end;
{$ENDIF EAGL}

function VireExtension(Fichier: String): String;
// retourne le chemin complet du fichier, sans l'extension
begin
  Result := Fichier;
  if ExtractFileExt(Fichier)<>'' then
    // ExtractFileExt contient le . !
    Result := Copy(Fichier, 1, Length(Fichier)-Length(ExtractFileExt(Fichier)));
end;

function  InitialiseZip(CheminZip: String): TOZ;
var zip : TOZ;
begin
  zip := TOZ.Create;
  zip.OpenZipFile(CheminZip, moCreate);
  zip.ActiveSwitchDisqueDefault(False, False, True, True);
  zip.ZipFileWithPath := True;
  zip.NiveauDeCompression := 9;
  Result := zip;
end;

function  FinaliseZip(var zip: TOZ): Boolean;
begin
  Result := False;
  if zip=Nil then exit;
  // Compression effective
  Result := zip.CloseSession;
  if Not Result then zip.CancelSession;
  zip.Free;
  zip := Nil;
end;

// AJOUT ME 30/10/2005 déplacement de utilSauvDossier
function AjouteAuZip(var zip: TOZ; chemin: String): Boolean;
begin
  Result := False;
{ MD 11/09/06 - on ne ferme plus la session à chaque fichier ajouté au zip
  // si on n'a pas déjà ouvert la session...
  if zip.TypeSession=osNone then zip.OpenSession(osAdd);
  zip.NiveauDeCompression := 9;
  zip.ActiveSwitchDisqueDefault(False, False, True, True);
  zip.ZipFileWithPath := True;
  if zip.ProcessFile(chemin) then // ajoute le fichier à la liste à traiter
    zip.CloseSession // compression effective
  else
    begin
    zip.CancelSession;
    exit;
    end;
  Result := True; }
  if zip.TypeSession=osNone then exit;
  if zip.ProcessFile(chemin) then Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 01/01/2003
Modifié le ... : 05/06/2007
Description .. : Exclure : pour préciser des extensions à exclure,
Suite ........ :           présenter sous la forme .XXX;.YYY;.ZZZ
Mots clefs ... :
*****************************************************************}
function ZippeDirectory(var zip: TOZ; laSource, leMasque: String;
  bWithSubDir, bFailIfExist: Boolean; Exclure : string=''): Boolean;
var SearchRec: TSearchRec;
    strDe, extension : string;
begin
  Result := False;
  if zip=Nil then exit;
  if not DirectoryExists(laSource) then exit;
  FindFirst(laSource+'\'+lemasque, faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
    begin
      strDe := laSource+'\'+SearchRec.Name;
        if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if (bWithSubDir) and (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          begin
          Result := ZippeDirectory(zip, strDe,leMasque,bWithSubDir,bFailIfExist, Exclure);
          if Result=False then begin sysutils.FindClose(SearchRec); exit; end;
          end;
        end
        else
        begin
          extension := UpperCase(ExtractFileExt(strDe));
          if pos(extension, Exclure) <> 0 then
            Result := True
          else
            Result := AjouteAuZip(zip, strDe);
          if Result=False then begin sysutils.FindClose(SearchRec); exit; end;
        end;
    end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
  Result := True;
end;


procedure ChercheUnFichierDans(root, nomfichier: String; var resultat: String;
  niveau: Integer; var nivtrouve: Integer);
// trouve le chemin du nomfichier dans la root fournie
// en prenant le plus haut trouvé dans l'arborescence (évite de se récupérer
// les sauvegardes mano faites en sous-répertoire, genre \sauve\nomfichier ...)
var chemin: String;
    SearchRec: TSearchRec;
begin
  if not DirectoryExists(root) then exit;
  FindFirst(root+'\*.*', faAnyFile, SearchRec);
  repeat
    if (SearchRec.Name<>'') then
      begin
      chemin := root+'\'+SearchRec.Name;
      // On est sur un répertoire : on cherche dedans en précisant qu'on
      // passe au niveau hiérarchique en dessous (niveau+1)
      if (SearchRec.Attr and SysUtils.faDirectory) <> 0 then
        begin
        if (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
          ChercheUnFichierDans(chemin, nomfichier, resultat, niveau+1, nivtrouve);
        end
      // On est sur un fichier, on le valide comme resultat,
      // uniquement si il est plus haut que le précédent trouvé
      else
        begin
        if UpperCase(ExtractFileName(chemin))=UpperCase(nomfichier) then
          begin
          // niveau plus "haut" (0=la racine de l'arbo)
          // ou pas encore de niveau trouve (-1)
          if (niveau<nivtrouve) or (nivtrouve=-1) then
            begin
            nivtrouve := niveau;
            resultat := chemin;
            sysutils.FindClose(SearchRec);
            exit;
            end;
          end;
        end;
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
end;


procedure ChercheUnRepertDans(root, nomrepert: String; var resultat: String;
  niveau: Integer; var nivtrouve: Integer);
// trouve le chemin du nomrepert dans la root fournie
// en prenant le plus haut trouvé dans l'arborescence
var chemin: String;
    SearchRec: TSearchRec;
begin
  if not DirectoryExists(root) then exit;
  FindFirst(root+'\*.*', faDirectory, SearchRec);
  repeat
    if (SearchRec.Name<>'') and (SearchRec.Name<>'.') and (SearchRec.Name<>'..') then
      begin
      chemin := root+'\'+SearchRec.Name;
      // Si ce n'est pas le répertoire cherché, on cherche dedans en précisant qu'on
      // passe au niveau hiérarchique en dessous (niveau+1)
      // (ExtractFileName renvoit bien le nom de la directory finale du chemin)
      if UpperCase(ExtractFileName(chemin))<>UpperCase(nomrepert) then
        begin
        ChercheUnRepertDans(chemin, nomrepert, resultat, niveau+1, nivtrouve);
        end
      // On est sur le bon répert, on le valide comme resultat,
      // uniquement si il est plus haut que le précédent trouvé
      else
        begin
        // niveau plus "haut" (0=la racine de l'arbo)
        // ou pas encore de niveau trouve (-1)
        if (niveau<nivtrouve) or (nivtrouve=-1) then
          begin
          nivtrouve := niveau;
          resultat := chemin;
          sysutils.FindClose(SearchRec);
          exit;
          end;
        end;
      end;
  until FindNext(SearchRec)<>0 ;
  sysutils.FindClose(SearchRec);
end;


procedure EnregistrerDll (sNomFichier : String);
var LibHandle : THandle;
    RegProc   : TRegProc;
begin
 LibHandle:=LoadLibrary (PChar (sNomFichier));
 if (LibHandle<>0) then
  begin
   try
    @RegProc := GetProcAddress(LibHandle, 'DllRegisterServer');
    if (V_PGI.Debug) then
     begin
      if (@RegProc=nil) then PGIInfo('Erreur sur DllRegisterServer', sNomFichier) else
      if (RegProc <> 0) then PGIInfo('Erreur d''enregistrement', sNomFichier);
     end;

   finally
    FreeLibrary(LibHandle);
   end;
  end
 else
  if V_PGI.Debug then PGIInfo('Erreur sur LoadLibrary', sNomFichier)
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 21/06/2006
Modifié le ... :   /  /    
Description .. : Teste si l'un des modules demandés est chargé en mémoire.
Suite ........ : Les modules doivent être précisés sous la forme 
Suite ........ : nom1.exe;nom2.exe;nom3.exe ...
Mots clefs ... :
*****************************************************************}
function  IsProcessLoaded (const sModuleNames : String) : Boolean;
var
  PIDArray   : array[0..1023] of DWORD;
  cb         : DWORD;
  i, j,
  ProcCount  : Integer;
  hMod       : HMODULE;
  hProcess   : THandle;
  LstModules : TStringList;
  ModuleName : array[0..300] of Char;
  St         : String;
begin
  Result := False;

  LstModules := TStringList.Create;
  St := sModuleNames;
  While St<>'' do
    LstModules.Add(UpperCase(ReadTokenSt(St)));
  if LstModules.Count=0 then begin LstModules.Free; exit; end;

  // Si on gèrait une listview
  // Liste.Clear;

  // Recupération de la liste de tous les processus sous NT4
  EnumProcesses(@PIDArray, SizeOf(PIDArray), cb);
  ProcCount := cb div SizeOf(DWORD);
  for i := 0 to ProcCount - 1 do
  begin
    If PIDArray[i] = 0 Then ModuleName := '[Idle]'
    Else
    begin
      hProcess := OpenProcess(PROCESS_QUERY_INFORMATION or PROCESS_VM_READ, False, PIDArray[i]);
      if (hProcess <> 0) then
      begin
        EnumProcessModules(hProcess, @hMod, SizeOf(hMod), cb);
        FillChar(ModuleName, 300, 0);
        ModuleName := 'System';
        GetModuleBaseName(hProcess, hMod, ModuleName, SizeOf(ModuleName));
        CloseHandle(hProcess);
      end
      Else ModuleName := 'Erreur';
    End;
    St := PChar(@ModuleName);
    St := UpperCase(St);

    // Module trouvé
    for j:=0 to LstModules.Count-1 do
      begin
      if St=LstModules[j] then
        begin
        Result := True;
        LstModules.Free;
        exit;
        end;
      end;

    // Si on gèrait une listview
  { With Liste.Items.Add Do
    Begin
      // Affichage de chaque process
      Caption:=IntToHex(PIDArray[i],8);
      SubItems.Add(ModuleName);

      // Data va conserver le ProcessID afin de pouvoir
      // en demander la fermeture par la suite
      Data:=Pointer(PIDArray[i]);
    End; }
  end;
end;

// $$$ JP 11/10/06
function GetFileVersion (strFile:string):string;
var
   dwFileVerSize  :DWORD;
   Handle         :DWORD;
   pVerHandle     :Pointer;
   pVerLangHandle :Pointer;
   dwVerLangLen   :DWORD;
   pVerBuffer     :Pointer;
   dwVerBufferLen :DWORD;
   strInfoVer     :string;
begin
     // Par défaut, pas de version connue du fichier
     Result := '';

     // Lecture information de version de fichier de la dll principale de Ews
     dwFileVerSize := GetFileVersionInfoSize (pchar (strFile), Handle);
     if dwFileVerSize > 0 then
     begin
          pVerHandle := AllocMem (dwFileVerSize);
          try
             if GetFileVersionInfo (pchar (strFile), Handle, dwFileVerSize, pVerHandle) = TRUE then
             begin
                  if VerQueryValue (pVerHandle, '\VarFileInfo\Translation', pVerLangHandle, dwVerLangLen) then
                  begin
                       strInfoVer := Format ('\StringFileInfo\%0.4x%0.4x\FileVersion'#0, [LoWord (LongInt (pVerLangHandle^)), HiWord (LongInt (pVerLangHandle^))]);
                       if VerQueryValue (pVerHandle, @strInfoVer [1], pVerBuffer, dwVerBufferLen) then
                          Result := pchar (pVerBuffer);
                  end;
             end;
          finally
                 FreeMem (pVerHandle);
          end;
     end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MB
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : Fournit une racine de clé user dans la registry,
Suite ........ : avec un préfixage par nom de société et nom de serveur 
Suite ........ : eagl.
Mots clefs ... : 
*****************************************************************}
function GetFullRegKey (const RegValueName : String ) : String ;
begin
    result :=  RegValueName+'_'+V_PGI.FCurrentAlias ;
    {$IFDEF EAGLCLIENT}
       Result := Result+'_'+V_PGI.eAGLHost ;
    {$ENDIF}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MB
Créé le ...... : 11/04/2007
Modifié le ... :   /  /    
Description .. : Affecte des groupes de travail à un dossier,
Suite ........ : via la nouvelle gestion des groupes de travail
Suite ........ : (tables :
Suite ........ : GRPDONNEES / LIENDOSGRP / LIENDONNEES)
Mots clefs ... : 
*****************************************************************}
function AffecteGroupe (nodoss:string; strGroupeConf:string):string;
var UneTob,UneTobTempo,UneTobEnreg,Tmp : Tob;
    LeGroupe : String ;
    MyId, MyParent : Integer ;
    procedure AddToTobParent(Id : Integer ; Mark : boolean) ;
    begin
       UneTobEnreg:=TOB.Create ('LIENDOSGRP', UneTob, -1);
       UneTobEnreg.SetString('LDO_NOM','GROUPECONF');
       UneTobEnreg.SetInteger('LDO_GRPID',Id);
       UneTobEnreg.SetBoolean('LDO_MARK',Mark);
       UneTobEnreg.SetString('LDO_NODOSSIER',NoDoss);
    end;
begin

 Result := '';
 if nodoss = '' then exit;

 //--- Effacement des anciennes données
 ExecuteSQL ('DELETE FROM LIENDOSGRP WHERE LDO_NODOSSIER="'+NoDoss+'"');

 //--- Construction de la tob
 UneTob:=TOB.Create('', Nil, -1);
 if (UneTob<>nil) then
  begin
   UneTobTempo := TOB.Create('GRPDONNEES',nil,0);
   UneTobTempo.LoadDetailFromSQL('SELECT * FROM GRPDONNEES WHERE GRP_NOM="GROUPECONF"');
   while (StrGroupeConf<>'') do
   begin
       LeGroupe := ReadToKenPipe (StrGroupeConf,';') ;
       Tmp := UneTobTempo.FindFirst(['GRP_CODE'],[LeGroupe],false) ;
       if Tmp <> nil then
       begin
           MyId := Tmp.GetInteger('GRP_ID') ;
           MyParent := Tmp.GetInteger('GRP_IDPERE');
           AddToTobParent(MyId,True) ;
           while MyParent > -1 do
           begin
              Tmp := UneTobTempo.FindFirst(['GRP_ID'],[myparent],false) ;
              if Tmp <> nil then
              begin
                 MyId := Tmp.GetInteger('GRP_ID') ;
                 MyParent := Tmp.GetInteger('GRP_IDPERE');
                 AddToTobParent(MyId,false) ;
              end
              else
                 MyParent := -1 ;
           end;
       end;
   end;
   UneTobTempo.free ;
   UneTob.SetAllModifie(True);
   UneTob.InsertOrUpdateDB (FALSE);
   UneTob.Free;
   Result := '  -> Dossier '+NoDoss+' affecté au(x) groupe(s) ';
  end
 else
  Result := '  -> Impossible d''affecter le dossier '+NoDoss+' à un groupe.';
end;


function IsNoDossierOk(stNoDossier: string): boolean;
// Vérifie la validité du No de dossier (True si OK)
// Attention, le no de dossier contient maintenant 6 c (même si varchar 8)
// Interdit tout caractère autre que des chiffres et des lettres...
var i: integer;
begin
  result := false;
  if stNoDossier = '' then exit;
  result := True;
  for i := 1 to Length(stNoDossier) do
  begin
    if not (stNoDossier[i] in ['0'..'9']) and not (stNoDossier[i] in ['A'..'Z']) then
    begin
      result := false;
      break;
    end;
  end;
end;


function JaiBloque(Niveau: string): boolean;
var Q: TQuery;
begin
  Niveau := Uppercase(Trim(Niveau));
  if ((not V_PGI.VersionReseau) or (Niveau = '') or (Niveau = 'NRAUCUN')) then begin Result := False; Exit; end;
  Q := OpenSQL('SELECT MG_EXPEDITEUR FROM COURRIER WHERE MG_UTILISATEUR="' + W_W + '" AND MG_COMBO="' + Niveau + '" AND MG_TYPE=1000 AND MG_EXPEDITEUR="' + V_PGI.User + '"', TRUE, 1);
  JaiBloque := (not Q.EOF);
  Ferme(Q);
end;

end.


unit UfronctionBase;

interface
uses
  Windows,
  SysUtils,
  Registry,
  SHellApi,
  Uconstantes,
  UEncryptage,
  UControleUAC,
  Uregistry,
  Ulog,
  ComCtrls,
  Forms,
  ZPatience
  ;

procedure PrepareRepository;
function ConstitueRepository : Boolean;
procedure  GetServerFiles (RepertStockage : string);
procedure BuildRepositoryDetail (TheRepository,TheSubDirectory,FileOut : string;  WithEcho : Boolean=true; PB : TprogressBar = nil);
procedure RegenConfFiles ;
function UpdateUpdProducts : boolean;
function CreateUniqueGUIDFileName(sPath, sPrefix, sExtension : string) : string;

implementation
uses UCommonFuncs;

procedure TraiteFile(TheRepository,TheSubDirectory,FileTrait,FileOut : string);
var NameConfig : string;
    DataStr : string;
begin
  DataStr := FileTrait+SEPARATOR+MD5(IncludeTrailingBackslash(TheSubDirectory)+FileTrait);
  TRY
    ecritLog (IncludeTrailingBackslash(TheRepository),DataStr,FileOut);
  EXCEPT
    OSError ('Ecriture du fichier : '+FileTrait);
  END;
end;


procedure BuildRepositoryDetail (TheRepository,TheSubDirectory,FileOut : string;  WithEcho : Boolean=true; PB : TprogressBar = nil);

  procedure MoveDefile(PB : TprogressBar);
  begin
    if PB.Position < PB.Max then
       PB.Position := PB.Position + 1
    else
       PB.Position := PB.Min ;

   PB.Invalidate;
   Application.ProcessMessages;
  end;
var Info   : TSearchRec;
    II : Integer;
begin
  II := 0;
  If FindFirst(IncludeTrailingBackslash(TheSubDirectory)+'*.*',faAnyFile,Info)=0 Then
  Begin
    if WithEcho then Writeln('Creation des MD5 pour '+ TheSubDirectory);
    Repeat
      If (info.Name<>'.')And(info.Name<>'..') then
      begin
        If ((Info.Attr and faDirectory)<>faDirectory) then
        begin
          if PB <> nil then MoveDefile(PB);
          TraiteFile(TheRepository,TheSubDirectory,Info.Name,FileOut);
          inc(II);
        end;
      end;
    { Il faut ensuite rechercher l'entrée suivante }
    Until FindNext(Info)<>0;
    if WithEcho then Writeln('MD5 pour '+ TheSubDirectory+' --> OK');
  end;
end;

procedure PrepareRepository;
var Treg : TRegistry;
    TheShare,TheInstallDirCegid,TheInstallDirLSE,TheRepository,TheRepoExeCegid,TheRepoExeLSe,TheRepoData,TheProgramData,cmd : string;
begin
  TheShare := '';
  TheRepository := '';
  TheRepoData := '';
  TheInstallDirCegid := '';
  TheInstallDirLSE := '';
  Treg := TRegistry.Create;
  TRY
    with Treg do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
      TheShare := ReadString('DIRCOPY');
      TheInstallDirCegid := ReadString('INSTALLDIR');
      TheInstallDirLSE := ReadString('INSTDIRLSE');
      CloseKey;
      if not OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',false) then Exit;
      TheProgramData := ReadString('Common AppData');
    end;
  finally
    Treg.Free;
  end;

  if TheShare = '' then Exit;
  if TheProgramData = '' then Exit;
  //
  if not DirectoryExists (TheShare) then
  begin
    if not CreateDir (TheShare) then Exit;
  end;
  //
  TheRepository := IncludeTrailingBackslash (TheShare)+'REPOSITORY';
  if not DirectoryExists (TheRepository) then
  begin
    if not CreateDir (TheRepository) then Exit;
  end;
  TheRepoExeCEGID := IncludeTrailingBackslash(TheRepository)+'CEGID';
  if not DirectoryExists (TheRepoExeCEGID) then
  begin
    if not CreateDir (TheRepoExeCEGID) then Exit;
  end;
  TheRepoExeLSE := IncludeTrailingBackslash(TheRepository)+'LSE';
  if not DirectoryExists (TheRepoExeLSE) then
  begin
    if not CreateDir (TheRepoExeLSE) then Exit;
  end;
  TheRepoData := IncludeTrailingBackslash(TheRepository)+'STD';
  if not DirectoryExists (TheRepoData) then
  begin
    if not CreateDir (TheRepoData) then Exit;
  end;
  TheRepoData := IncludeTrailingBackslash(TheRepository)+'DATA';
  if not DirectoryExists (TheRepoData) then
  begin
    if not CreateDir (TheRepoData) then Exit;
  end;
end;

function ConstitueRepository : Boolean;
var Treg : TRegistry;
    TheShare,TheInstallDirCegid,TheInstallDirLSE,TheRepository,TheRepoExeCegid,TheRepoExeLSe,TheRepoData,TheMode,TheProgramData,cmd,TheRepoSTD,TheInstallDirSTD,TheRepoUPD,TheInstallDirUPD : string;
begin
  Result := false;
  TheShare := '';
  TheRepository := '';
  TheRepoData := '';
  TheMode := '';
  TheInstallDirCegid := '';
  TheInstallDirLSE := '';
  TheInstallDirSTD := '';
  TheRepoSTD := '';
  TheRepoUPD := '';
  Treg := TRegistry.Create;
  TRY
    with Treg do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
      TheShare := ReadString('DIRCOPY');
      TheInstallDirCegid := ReadString('INSTALLDIR');
      TheInstallDirLSE := ReadString('INSTDIRLSE');
      TheInstallDirSTD := 'C:\PGI00\STD';
      CloseKey;
      if not OpenKey ('SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion',False) then Exit;
      TheInstallDirUPD := IncludeTrailingBackslash(readString('ProgramFilesDir'))+'CEGID\LSE Live Update\APP';
      CloseKey;
      if not OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',false) then Exit;
      TheProgramData := ReadString('Common AppData');
    end;
  finally
    Treg.Free;
  end;

  if TheShare = '' then Exit;
  if TheProgramData = '' then Exit;
  //
  if not DirectoryExists (TheShare) then
  begin
    if not CreateDir (TheShare) then Exit;
  end;
  //
  TheRepository := IncludeTrailingBackslash (TheShare)+'REPOSITORY';
  if not DirectoryExists (TheRepository) then
  begin
    if not CreateDir (TheRepository) then Exit;
  end;
  TheRepoExeCEGID := IncludeTrailingBackslash(TheRepository)+'CEGID';
  if not DirectoryExists (TheRepoExeCEGID) then
  begin
    if not CreateDir (TheRepoExeCEGID) then Exit;
  end;
  TheRepoExeLSE := IncludeTrailingBackslash(TheRepository)+'LSE';
  if not DirectoryExists (TheRepoExeLSE) then
  begin
    if not CreateDir (TheRepoExeLSE) then Exit;
  end;
  TheRepoData := IncludeTrailingBackslash(TheRepository)+'DATA';
  if not DirectoryExists (TheRepoData) then
  begin
    if not CreateDir (TheRepoData) then Exit;
  end;
  TheRepoSTD := IncludeTrailingBackslash(TheRepository)+'STD';
  if not DirectoryExists (TheRepoSTD) then
  begin
    if not CreateDir (TheRepoSTD) then Exit;
  end;
  TheRepoUPD := IncludeTrailingBackslash(TheRepository)+'UPDATE';
  if not DirectoryExists (TheRepoUPD) then
  begin
    if not CreateDir (TheRepoUPD) then Exit;
  end;

  if fileExists(IncludeTrailingBackslash(TheRepository)+CONFCEGIDFILE) then  DeleteFile(IncludeTrailingBackslash(TheRepository)+CONFCEGIDFILE);
  if fileExists(IncludeTrailingBackslash(TheRepository)+CONFLSEFILE) then DeleteFile(IncludeTrailingBackslash(TheRepository)+CONFLSEFILE);
  if fileExists(IncludeTrailingBackslash(TheRepository)+PARAMUPDATE) then DeleteFile(IncludeTrailingBackslash(TheRepository)+PARAMUPDATE);
  if fileExists(IncludeTrailingBackslash(TheRepository)+CONFSTDFILE) then DeleteFile(IncludeTrailingBackslash(TheRepository)+CONFSTDFILE);
  if fileExists(IncludeTrailingBackslash(TheRepository)+CONFUPDFILE) then DeleteFile(IncludeTrailingBackslash(TheRepository)+CONFUPDFILE);

  if TheInstallDirCegid <> '' then
  begin
    Writeln('preparation du dépot CEGID');
    ShellExecAndWait('c:\windows\system32\xcopy', '"'+IncludeTrailingBackslash(TheInstallDirCegid)+'APP'+'"'+' '+TheRepoExeCEGID+' /C/Y',SW_HIDE);
  end;
  if TheInstallDirLSE <> '' then
  begin
    Writeln('preparation du dépot LSE');
    ShellExecAndWait('c:\windows\system32\xcopy', '"'+IncludeTrailingBackslash(TheInstallDirLSE)+'APP'+'"'+' '+TheRepoExeLSE+' /C/Y',SW_HIDE);
  end;
  if TheInstallDirSTD <> '' then
  begin
    Writeln('preparation du dépot STD');
    ShellExecAndWait('c:\windows\system32\xcopy', '"'+TheInstallDirSTD+'"'+' '+TheRepoSTD+' /C/Y',SW_HIDE);
  end;
  if TheInstallDirUPD <> '' then
  begin
    Writeln('preparation du dépot UPDATE');
    ShellExecAndWait('c:\windows\system32\xcopy', '"'+TheInstallDirUPD+'"'+' '+TheRepoUPD+' /C/Y',SW_HIDE);
  end;
  if FileExists(IncludeTrailingBackslash(TheProgramData)+'Cegid\SocRefBtp.mdb') then
  begin
    Writeln('preparation du dépot DATA');
    ShellExecAndWait('c:\windows\system32\xcopy', '"'+IncludeTrailingBackslash(TheProgramData)+'Cegid\SocRefBtp.mdb" '+TheRepoData+' /C/Y',SW_HIDE);
  end;
  //
  BuildRepositoryDetail (TheRepository,TheRepoExeCegid,CONFCEGIDFILE);
  BuildRepositoryDetail (TheRepository,TheRepoExeLSe,CONFLSEFILE);
  BuildRepositoryDetail (TheRepository,TheRepoSTD,CONFSTDFILE);
  BuildRepositoryDetail (TheRepository,TheRepoData,PARAMUPDATE);
  BuildRepositoryDetail (TheRepository,TheRepoUPD,CONFUPDFILE);
  Sleep(3000);
  //
  Result := True;
end;

procedure RegenConfFiles ;
var Treg : Tregistry;
    RootKey : string;
    TheShare : string;
    TheProgramData : string;
    TheRepository,TheRepoExeCEGID,TheRepoExeLSE,TheRepoData,TheRepoSTD,TheRepoUPD : string;
    XX : TFPatience;
begin
  Treg := TRegistry.Create;
  TRY
    with Treg do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
      TheShare := ReadString('DIRCOPY');
      CloseKey;
      if not OpenKey ('SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion',False) then Exit;
      TheRepoUPD := IncludeTrailingBackslash(readString('ProgramFilesDir'))+'CEGID\LSE Live Update\APP';
      CloseKey;
      if not OpenKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',false) then Exit;
      TheProgramData := ReadString('Common AppData');
    end;
  finally
    Treg.Free;
  end;

  TheRepository := IncludeTrailingBackslash (TheShare)+'REPOSITORY';
  TheRepoExeCEGID := IncludeTrailingBackslash(TheRepository)+'CEGID';
  TheRepoExeLSE := IncludeTrailingBackslash(TheRepository)+'LSE';
  TheRepoSTD := IncludeTrailingBackslash(TheRepository)+'STD';
  TheRepoData := IncludeTrailingBackslash(TheRepository)+'DATA';
  TheRepoUpd := IncludeTrailingBackslash(TheRepository)+'UPDATE';

  XX := FenetrePatience ('-');
  XX.StartK2000;

  if not FileExists (IncludeTrailingBackslash(TheRepository)+CONFCEGIDFILE) then
  begin
    XX.SetTitre('CALCUL DES CRC CEGID');
    BuildRepositoryDetail (TheRepository,TheRepoExeCEGID,CONFCEGIDFILE,false,XX.ProgressBar1);
  end;

  if not FileExists (IncludeTrailingBackslash(TheRepository)+CONFLSEFILE) then
  begin
    XX.SetTitre('CALCUL DES CRC LSE');
    BuildRepositoryDetail (TheRepository,TheRepoExeLSE,CONFLSEFILE,false,XX.ProgressBar1);
  end;

  if not FileExists (IncludeTrailingBackslash(TheRepository)+CONFSTDFILE) then
  begin
    XX.SetTitre('CALCUL DES CRC STD');
    BuildRepositoryDetail (TheRepository,TheRepoSTD,CONFSTDFILE,false,XX.ProgressBar1);
  end;

  if not FileExists (IncludeTrailingBackslash(TheRepository)+CONFUPDFILE) then
  begin
    XX.SetTitre('CALCUL DES CRC UPDATE');
    BuildRepositoryDetail (TheRepository,TheRepoUpd,CONFUPDFILE,false,XX.ProgressBar1);
  end;

  if not FileExists (IncludeTrailingBackslash(TheRepository)+PARAMUPDATE) then
  begin
    XX.SetTitre('CALCUL DES CRC COMMUN');
    BuildRepositoryDetail (TheRepository,TheRepoData,PARAMUPDATE,false,XX.ProgressBar1);
  end;
  XX.StopK2000;
  XX.Free;
end;

procedure  GetServerFiles (RepertStockage : string);
var fileLTemp,fileCTemp,fileDTemp,ServerFile,FileSTemp,FileUtemp : string;
begin
  // -- RECUP DES FICHIERS SERVEURS EN LOCAL POUR DETERMINATION SI BESOIN MAJ
  fileLTemp := IncludeTrailingBackslash (GetTempDir)+CONFLSEFILE;
  if FileExists(fileLTemp) then DeleteFile(fileLTemp);
  fileCTemp := IncludeTrailingBackslash (GetTempDir)+CONFCEGIDFILE;
  if FileExists(fileCTemp) then DeleteFile(fileCTemp);
  fileDTemp := IncludeTrailingBackslash (GetTempDir)+PARAMUPDATE;
  if FileExists(fileDTemp) then DeleteFile(fileDTemp);
  FileSTemp := IncludeTrailingBackslash (GetTempDir)+CONFSTDFILE;
  if FileExists(FileSTemp) then DeleteFile(FileSTemp);
  FileUTemp := IncludeTrailingBackslash (GetTempDir)+CONFUPDFILE;
  if FileExists(FileUTemp) then DeleteFile(FileUTemp);

  ServerFile := IncludeTrailingBackslash (RepertStockage)+CONFLSEFILE;
  if FileExists(ServerFile) then
  begin
    CopyFile(PAnsiChar(ServerFile),PansiChar(fileLTemp),false);
  end else Exit;
  //----
  ServerFile := IncludeTrailingBackslash (RepertStockage)+CONFCEGIDFILE;
  if FileExists(ServerFile) then
  begin
    CopyFile(PAnsiChar(ServerFile),PansiChar(fileCTemp),false);
  end  else Exit;
  //----
  ServerFile := IncludeTrailingBackslash (RepertStockage)+CONFSTDFILE;
  if FileExists(ServerFile) then
  begin
    CopyFile(PAnsiChar(ServerFile),PansiChar(FileSTemp),false);
  end  else Exit;
  // ---
  ServerFile := IncludeTrailingBackslash (RepertStockage)+PARAMUPDATE;
  if FileExists(ServerFile) then
  begin
    CopyFile(PAnsiChar(ServerFile),PansiChar(fileDTemp),false);
  end  else Exit;

  ServerFile := IncludeTrailingBackslash (RepertStockage)+CONFUPDFILE;
  if FileExists(ServerFile) then
  begin
    CopyFile(PAnsiChar(ServerFile),PansiChar(FileUTemp),false);
  end  else Exit;
end;


function TraitDetailUpdate (RepertStockage,TheRepository,TheInstallDirUPD: string; TTL,TTS: TProductList) : boolean;
var II : Integer;
    From,TODest,ToTemp,TheName,cmd  : string;
    ServerFile,TempFile,DestFile : string;
    LT,DT : TElement;
    isOneTreated : Boolean;
begin
  //
  isOneTreated := False;
  result := false;
  From := IncludeTrailingBackslash(IncludeTrailingBackslash(RepertStockage)+'UPDATE');
  ToTemp := IncludeTrailingBackslash(IncludeTrailingBackslash(TheRepository)+'UPDATE');
  TODest := IncludeTrailingBackslash(TheInstallDirUPD)+'APP';

  for II := 0 to TTL.Count -1 do
  begin
    LT := TTL.items[II];
    TheName := LT.Element;
    DT := TTS.Find(TheName);
    if DT <> nil then
    begin
      if DT.Signature <> LT.Signature then
      begin
        result := True;
        if not isOneTreated then isOneTreated := True;
        //
        ServerFile := IncludeTrailingBackslash(From)+TheName;
        TempFile := IncludeTrailingBackslash(ToTemp)+TheName;
        DestFile := IncludeTrailingBackslash(ToDest)+TheName;
        ShellExecAndWait('c:\windows\system32\xcopy', '"'+ServerFile+'"  "'+TempFile+'" /C/Y',SW_HIDE);
        ShellExecAndWait('c:\windows\system32\xcopy', '"'+Tempfile+'"  "'+DestFile+'" /C/Y',SW_HIDE);
        //
      end;
    end;
  end;
  Result := isOneTreated;
end;

function TraiteUpdProducts (RepertStockage,TheRepository,TheInstallDirUPD : string) : boolean;
var TTL,TTS : TProductList;
    LLOC,LTemp : string;
    ATraiter : boolean;
begin

  result := false;
  TTL := TProductList.Create;
  TTS := TProductList.Create;
  TRY
    LLOc := IncludeTrailingBackslash (TheRepository)+CONFUPDFILE ;
    if FileExists(LLOC) then
    begin
      ChargeListeProducts (LLOc,TTL);
      LTemp := IncludeTrailingBackslash (GetTempDir)+CONFUPDFILE;
      if FileExists (LTemp) then
      begin
        ChargeListeProducts (LTemp,TTS);
        ATraiter := TraitDetailUpdate (RepertStockage,TheRepository,TheInstallDirUPD,TTL,TTS);
        if ATraiter then DeleteFile(LLOC);
      end;
    end;
  FINALLY
    Result := ATraiter;
    TTL.free;
    TTS.free;
  END;

end;

function UpdateUpdProducts : boolean;
var TReg : TRegistry;
    TheShare,TheInstallDirUPD,TheRepository,RepertStockage : string;
begin
  Result := false;
  RepertStockage:= GetInfoStocke ('ServerRepository');
  Treg := TRegistry.Create;
  TRY
    with Treg do
    begin
      RootKey := HKEY_LOCAL_MACHINE;
      if not OpenKey ('SOFTWARE\Wow6432Node\Cegid\Cegid Business',False) then Exit;
      TheShare := ReadString('DIRCOPY');
      CloseKey;
      if not OpenKey ('SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion',False) then Exit;
      TheInstallDirUPD := IncludeTrailingBackslash(readString('ProgramFilesDir'))+'CEGID\LSE Live Update\APP';
    end;
  finally
    Treg.Free;
  end;
  if TheShare = '' then exit;
  TheRepository := IncludeTrailingBackslash (TheShare)+'REPOSITORY';
  Result := TraiteUpdProducts (RepertStockage,TheRepository,TheInstallDirUPD);
end;

function CreateUniqueGUIDFileName(sPath, sPrefix, sExtension : string) : string;
var sFileName : string;
    Guid : TGUID;
begin
  Result := '';
  repeat
   SFileName := '';
   CreateGUID(Guid);
   SFileName := sPath + sPrefix + GUIDtoString(GUID);
   Result := ChangeFileExt(sFileName, sExtension)
  until not FileExists(Result);
end;


end.

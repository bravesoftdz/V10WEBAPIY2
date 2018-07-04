unit UControleUAC;

interface
uses shellapi,Windows,SysUtils,ShlObj,Messages,TLHelp32;

const
  LOGON_WITH_PROFILE=$00000001;

function CreateProcessWithLogon(lpUsername        :PWideChar;
                                lpDomain          :PWideChar;
                                lpPassword        :PWideChar;
                                dwLogonFlags      :DWORD;
                                lpApplicationName :PWideChar;
                                lpCommandLine     :PWideChar;
                                dwCreationFlags   :DWORD;
                                lpEnvironment     :Pointer;
                                lpCurrentDirectory:PWideChar;
                                var lpStartupInfo :TStartupInfo;
                                var lpProcessInfo :TProcessInformation):BOOL;stdcall;external 'advapi32.dll' name 'CreateProcessWithLogonW';
function CreateEnvironmentBlock(var lpEnvironment:Pointer;hToken:THandle;bInherit:BOOL):BOOL;stdcall;external 'userenv.dll';
function DestroyEnvironmentBlock(pEnvironment:Pointer):BOOL;stdcall;external 'userenv.dll';

function RunProcess(Command : String;Parameters:array of string;WorkingDirectory:string=''; Wait: Boolean=true) : integer;
function RunProcessAs(Command:string;Parameters:array of string;Username,Password:string;Domain:string='';WorkingDirectory:string='';Wait:Boolean=False):Cardinal;
function Getuser (UserComposed : string) : string;
function GetDomain (UserComposed : string) : string;
procedure KillProcess(ClassName: PChar; Titre: PChar);
function IsProcessAlive(Titre: PChar) : boolean;
procedure RunAsAdmin(hWnd : HWND; aFile : String; aParameters : String);
function ShellExecAndWait(ExeFile: string; Parameters: string = ''; ShowWindow: Word = SW_SHOWNORMAL): Boolean;
function ExeRunning(NomApplication : string; StopProcess:Boolean):Boolean;
function RunAs(User, Password, Command: String): Integer;
procedure OSError(s:string);

implementation


function ExeRunning(NomApplication : string; StopProcess:Boolean):Boolean;
var
  ProcListExec : TProcessentry32;
  PrhListExec : Thandle;
  Continu : Boolean;
  isStarted : Boolean;
  HandleProcessCourant : Cardinal;
  PathProcessCourant : string;
  ProcessCourant :String;
begin
  ProcListExec.dwSize:=sizeof(ProcListExec);
  Continu := True;
  isStarted := False;

  Try
    // Récupére la liste des process en cours d'éxécution au moment de l'appel
    PrhListExec:=CreateToolhelp32Snapshot(TH32CS_SNAPALL,0);
    if (PrhListExec <> INVALID_HANDLE_VALUE) then
    begin
      //On se place sur le premier process
      Process32First(PrhListExec,ProcListExec);

      // Tant que le process recherché n'est pas trouvé et qu'il reste
      // des process dans la liste, on parcourt et analyse la liste
      while Continu do
      begin
        ProcessCourant := Uppercase(ExtractFileName(ProcListExec.szExeFile));
        ProcessCourant := ChangeFileExt(ProcessCourant,'');
        isStarted := (ProcessCourant = Uppercase(NomApplication));
        if isStarted then
        begin
          HandleProcessCourant := ProcListExec.th32ProcessID;
          PathProcessCourant := ExtractFilepath(ProcListExec.szExeFile);
          Continu := False;
        end
        else // Recherche le process suivant dans la liste
        Continu := Process32Next(PrhListExec,ProcListExec);
      end;

      if StopProcess then
      begin
        if isStarted then
        begin  // Termine le process en indiquant le code de sortie zéro
          TerminateProcess(OpenProcess(PROCESS_TERMINATE,False,HandleProcessCourant),0);
          Sleep(500); // Laisse le temps au process en cours de suppression de s'arrêter
        end;
      end;
    end;
  Finally
    CloseHandle(PrhListExec); // Libère les ressources
    Result := isStarted;
  end;

end;

function ShellExecAndWait(ExeFile: string; Parameters: string = ''; ShowWindow: Word = SW_SHOWNORMAL): Boolean;

  procedure WaitFor(processHandle: THandle);
  var
    AMessage : TMsg;
    Result   : DWORD;
  begin
    repeat
      Result := MsgWaitForMultipleObjects(1,
                                          processHandle,
                                          False,
                                          INFINITE,
                                          QS_PAINT or
                                          QS_SENDMESSAGE);
      if Result = WAIT_FAILED then
        Exit;
      if Result = ( WAIT_OBJECT_0 + 1 ) then
      begin
        while PeekMessage(AMessage, 0, WM_PAINT, WM_PAINT, PM_REMOVE) do
          DispatchMessage(AMessage);
      end;
    until result = WAIT_OBJECT_0;
  end;

var
  ExecuteCommand: array[0 .. 512] of Char;
  PathToExeFile : string;
  StartUpInfo   : TStartupInfo;
  ProcessInfo   : TProcessInformation;
begin
  StrPCopy(ExecuteCommand , ExeFile + ' ' + Parameters);
  FillChar(StartUpInfo, SizeOf(StartUpInfo), #0);
  StartUpInfo.cb  := SizeOf(StartUpInfo);
  StartUpInfo.dwFlags := STARTF_USESHOWWINDOW;
  StartUpInfo.wShowWindow := ShowWindow;//SW_SHOWNORMAL; //SW_MINIMIZE;
  PathToExeFile := ExtractFileDir(ExeFile);
  if PathToExeFile = '' then
    PathToExeFile := ExtractFilePath (ExeFile);
  if CreateProcess(nil,
                   ExecuteCommand,
                   nil,
                   nil,
                   False,
                   CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS,
                   nil,
                   PChar(PathToExeFile),
                   StartUpInfo,
                   ProcessInfo) then
  begin
    WaitFor(ProcessInfo.hProcess);
    CloseHandle(ProcessInfo.hProcess);
    CloseHandle(ProcessInfo.hThread);
    Result := true;
  end
  else
    Result := false;
end;

function IsProcessAlive(Titre: PChar) : boolean;
  var Appli : HWND;
begin
  Appli := FindWindow(nil, PChar(Titre));
  result := (Appli <> 0);
end;

procedure KillProcess(ClassName: PChar; Titre: PChar);
var
  ProcessHandle : THandle;
  ProcessID: Integer;
  Appli : HWND;
begin
  Appli := FindWindow(PChar(ClassName), PChar(Titre));
  if Appli <> 0 then
  begin
    GetWindowThreadProcessID(Appli, @ProcessID);
    ProcessHandle := OpenProcess(PROCESS_TERMINATE, FALSE, ProcessId);
    TerminateProcess(ProcessHandle,4);
  end;
end;

function FormatParam(s:string):string;
var
  a:Integer;
  t:Boolean;
begin
  Result:=s;
  t:=False;
  for a:=Length(s) downto 1 do
    if s[a] in [' ',#32] then begin
      t:=True;
      Break;
    end;
  if t and not (s[1] in ['''','"']) then
    Result:='"'+s+'"';
end;

procedure Error(s:string);
begin
  raise Exception.Create(s);
end;

procedure OSError(s:string);
begin
 raise Exception.Create(s+#13#10+SysErrorMessage(GetLastError));
end;


function RunProcess(Command : String;Parameters:array of string;WorkingDirectory:string=''; Wait: Boolean=true) : integer;
var
  StartInfo : TStartupInfo;
  ProcInfo : TProcessInformation;
  CreateOK : Boolean;
  wCommandLine,wCurrentDirectory : string;
  a:Integer;
begin
    { fill with known state }
	result := -1;
  FillChar(StartInfo,SizeOf(TStartupInfo),#0);
  FillChar(ProcInfo,SizeOf(TProcessInformation),#0);
  StartInfo.cb := SizeOf(TStartupInfo);

  if WorkingDirectory='' then
    wCurrentDirectory:=GetCurrentDir
  else
    wCurrentDirectory:=WorkingDirectory;
  wCurrentDirectory := IncludeTrailingBackslash(wCurrentDirectory);
  wCommandLine:='';
  // if not SetCurrentDir (wCurrentDirectory) then RaiseLastOsError;
  for a:=Low(Parameters) to High(Parameters) do
    wCommandLine:=wCommandLine+' "'+Parameters[a]+'" ';
  CreateOK := CreateProcess(PAnsiChar(Command), PChar(wCommandLine), nil, nil,False,
              CREATE_NEW_PROCESS_GROUP+NORMAL_PRIORITY_CLASS,
              nil, PAnsiChar(wCurrentDirectory), StartInfo, ProcInfo);
    { check to see if successful }
  if CreateOK then
    begin
        //may or may not be needed. Usually wait for child processes
      if Wait then
        WaitForSingleObject(ProcInfo.hProcess, INFINITE);
    end
  else
    begin
      RaiseLastOSError;
     end;
  CloseHandle(ProcInfo.hProcess);
  CloseHandle(ProcInfo.hThread);
end;

function RunProcessAs(Command:string;Parameters:array of string;Username,Password:string;Domain:string='';WorkingDirectory:string='';Wait:Boolean=False):Cardinal;
(*
  Execute the Command with the given Parameters, Username, Domain, Password and Working Directory. Parameters containing white spaces are
  automatically embraced into quotes before being sent to avoid having them splitted by the system. If either Domain or Working Directory
  are empty the current one will be used instead.

  If Wait is specified the function will wait till the command is completely executed and will return the exit code of the process,
  otherwise zero.

  Suitable Delphi exceptions will be thrown in case of API failure.
*)
var
  a:Integer;
  n:Cardinal;
  h:THandle;
  p:Pointer;
  PI:TProcessInformation;
  SI:TStartupInfo;
  t:array[0..MAX_PATH] of WideChar;
	wUser,wDomain,wPassword,wCommandLine,wCurrentDirectory:WideString;
begin
	result := 0;
  ZeroMemory(@PI,SizeOf(PI));
  ZeroMemory(@SI,SizeOf(SI));
  SI.cb:=SizeOf(SI);
  if not LogonUser(PChar(Username),nil,PChar(Password),LOGON32_LOGON_INTERACTIVE,LOGON32_PROVIDER_DEFAULT,h) then
  begin
  	OSError ('Connexion sur le compte utilisateur impossible');
  	exit;
  end;
  try
    if not CreateEnvironmentBlock(p,h,True) then
    begin
			OSError('Environnement utilisateur inaccessible');
      exit;
    end;
    try
      wUser:=Username;
      wPassword:=Password;
      wCommandLine:=Command;
      for a:=Low(Parameters) to High(Parameters) do
        wCommandLine:=wCommandLine+' '+FormatParam(Parameters[a]);
      if Domain='' then
      begin
        n:=SizeOf(t);
        if not GetComputerNameW(t,n) then
          exit;
        wDomain:=t;
      end else
        wDomain:=Domain;
      if WorkingDirectory='' then
        wCurrentDirectory:=GetCurrentDir
      else
        wCurrentDirectory:=WorkingDirectory;
      if not CreateProcessWithLogon(PWideChar(wUser),PWideChar(wDomain),PWideChar(wPassword),LOGON_WITH_PROFILE,nil,PWideChar(wCommandLine),CREATE_UNICODE_ENVIRONMENT,p,PWideChar(wCurrentDirectory),SI,PI) then
      begin
      	OSError('Impossible de créer le processus');
        exit;
      end;
      try
        if Wait then
        begin
          WaitForSingleObject(PI.hProcess,INFINITE);
          if not GetExitCodeProcess(PI.hProcess,Result) then
            exit;
        end else
         Result:=0;
      finally
        CloseHandle(PI.hProcess);
        CloseHandle(PI.hThread);
      end;
    finally
      DestroyEnvironmentBlock(p);
    end;
  finally
  	CloseHandle(h);
  end;
end;

function Getuser (UserComposed : string) : string;
var iPos : integer;
begin
  result := userComposed;
  ipos := pos ('/',userComposed);
  if iPos = 0 then ipos := pos ('\',userComposed);
  if Ipos = 0 then exit;
  result := copy(UserComposed,ipos+1,255);
end;

function GetDomain (UserComposed : string) : string;
var iPos : integer;
begin
	result := '';
  iPos := pos ('/',userComposed);
  if iPos = 0 then ipos := pos ('\',userComposed);
  if Ipos = 0 then exit;
  result := copy(UserComposed,1,ipos-1);
end;

procedure RunAsAdmin(hWnd : HWND; aFile : String; aParameters : String);
Var
 Sei : TShellExecuteInfoA;
begin
 Fillchar(sei,SizeOf(sei),0);
 sei.cbSize := SizeOf(sei);
 sei.Wnd    := hWnd;
 sei.fMask  := SEE_MASK_FLAG_DDEWAIT or SEE_MASK_FLAG_NO_UI;
 sei.lpfile := PChar(aFile);
 sei.lpVerb := 'runas';
 sei.lpParameters := PChar(aParameters);
 sei.nShow := SW_SHOWNORMAL;
 if not ShellExecuteEx(@sei) then
  RaiseLastOSError;
end;



function RunAs(User, Password, Command: String): Integer;
var  dwSize:        DWORD;
     hToken:        THandle;
     lpvEnv:        Pointer;
     pi:            TProcessInformation;
     si:            TStartupInfo;
     szPath:        Array [0..MAX_PATH] of WideChar;
begin

  ZeroMemory(@szPath, SizeOf(szPath));
  ZeroMemory(@pi, SizeOf(pi));
  ZeroMemory(@si, SizeOf(si));
  si.cb:=SizeOf(TStartupInfo);

  if LogonUser(PChar(User), nil, PChar(Password), LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken) then
  begin
     if CreateEnvironmentBlock(lpvEnv, hToken, True) then
     begin
        dwSize:=SizeOf(szPath) div SizeOf(WCHAR);
        if (GetCurrentDirectoryW(dwSize, @szPath) > 0) then
        begin
           if (CreateProcessWithLogon(PWideChar(WideString(User)), nil, PWideChar(WideString(Password)),
               LOGON_WITH_PROFILE, nil, PWideChar(WideString(Command)), CREATE_UNICODE_ENVIRONMENT,
               lpvEnv, szPath, si, pi)) then
           begin
              result:=ERROR_SUCCESS;
              CloseHandle(pi.hProcess);
              CloseHandle(pi.hThread);
           end
           else
              result:=GetLastError;
        end
        else
           result:=GetLastError;
        DestroyEnvironmentBlock(lpvEnv);
     end
     else
        result:=GetLastError;
     CloseHandle(hToken);
  end
  else
     result:=GetLastError;

end;


end.

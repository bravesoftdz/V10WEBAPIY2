unit ConstServices;

interface

type
  T_WSLogValues       = Record
                          LogLevel          : integer;
                          LogMoMaxSize      : double;
                          LogMaxQty         : integer;
                          LogDebug          : integer;
                          LogDebugMoMaxSize : double;
                          DebugEvents       : integer;
                          OneLogPerDay      : boolean;
                       end;

  T_SvcTypeLog  = (ssbylNone, ssbylLog, ssbylWindows, ssbylAll);

  TServicesLog = class (TObject)
    class function GetFilePath(ServiceName, Extension: string): string;
    class function CreateLog(LogValues : T_WSLogValues; ServiceName : string): string;
    class procedure WriteLog(TypeDebug: T_SvcTypeLog; Text, ServiceName: string; LogValues : T_WSLogValues; LineLevel: integer; WithoutDateTime: Boolean=true; AddFileName : string='');
  end;


const
  ServiceName_BTPY2     = 'SvcSynBTPY2';
  ServiceName_BTPVerdon = 'SvcSynBTPVerdon';
  WSCDS_ErrorMsg        = '##### ERREUR';
  WSCDS_DebugMsg        = '***** DEBUG : ';


implementation

uses
  SvcMgr
  , SysUtils
  , Windows
  , CommonTools
  , UConnectWSConst
  ;

{ TServicesLgo }

class function TServicesLog.GetFilePath(ServiceName, Extension: string): string;
begin
  Result := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), ServiceName, Extension]);
end;

class function TServicesLog.CreateLog(LogValues : T_WSLogValues; ServiceName : string): string;
var
  SizeFile    : Extended;
  SearchFile  : TSearchRec;
  MaxSize     : double;
begin
  if LogValues.LogLevel > 0 then
  begin
    Result := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), ServiceName, 'log']);
    if Logvalues.OneLogPerDay then
      Result := Format('%s_%s.log', [Copy(Result, 1, pos('.log', Result) -1), Tools.CastDateTimeForQry(Now)]);
    if LogValues.DebugEvents > 0 then TServicesLog.WriteLog(ssbylLog, Format('%s - TSvcSyncBTPY2Execute.LogsManagement : LogFilePath = %s', [WSCDS_DebugMsg, Result]), ServiceName, LogValues, 0);
    if not LogValues.OneLogPerDay then
    begin
      MaxSize := LogValues.LogMoMaxSize;
      { Si dépasse la taille max, supprime puis créé un nouveau }
      if (MaxSize > 0) then
      begin
        if FindFirst(Result, faAnyFile, SearchFile) = 0 then
        try
          begin
            SizeFile := Tools.GetFileSize(Result, tssMo);
            if SizeFile > MaxSize then
              SysUtils.DeleteFile(Result);
          end;
        finally
          SysUtils.FindClose(SearchFile);
        end;
      end;
    end else
      Result := Format('%s_%s.log', [Copy(Result, 1, pos('.log', Result) -1), Tools.CastDateTimeForQry(Now)]);
  end else
    Result := '';
end;

class procedure TServicesLog.WriteLog(TypeDebug: T_SvcTypeLog; Text, ServiceName: string; LogValues: T_WSLogValues; LineLevel: integer; WithoutDateTime: Boolean=true; AddFileName : string='');
var
  LogText     : string;
  LogFilePath : string;
  WindowsLog  : TEventLogger;
  LogFile     : TextFile;

  procedure WriteWindowsLog;
  begin
    LogFilePath := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), ServiceName, 'exe']);
    WindowsLog := TEventLogger.Create(ExtractFileName(LogFilePath));
    try
      WindowsLog.LogMessage(Text, EVENTLOG_INFORMATION_TYPE);
    finally
      WindowsLog.Free;
    end;
  end;

  procedure WriteUserLog;
  begin
    if LogValues.LogLevel > 0 then
    begin
      LogFilePath := Format('%s%s.%s', [ExtractFilePath(ParamStr(0)), ServiceName, 'log']);
      if Logvalues.OneLogPerDay then
        LogFilePath := Format('%s_%s.log', [Copy(LogFilePath, 1, pos('.log', LogFilePath) -1), Tools.CastDateTimeForQry(Now)]);
      if AddFileName <> '' then
        LogFilePath := Format('%s_%s.log', [Copy(LogFilePath, 1, pos('.log', LogFilePath) -1), AddFileName]);
      AssignFile(LogFile, LogFilePath);
      try
        if FileExists(LogFilePath) then
          Append(LogFile)
        else
          Rewrite(LogFile);
        if Text <> '' then
        begin
          if WithoutDateTime then
            LogText := Format('%s : %s%s', [DateTimeToStr(Now), StringOfChar(' ', LineLevel), Text])
          else
            LogText := Format('%s : %s%s', [Copy(Text, 1, pos('=', Text) - 1), StringOfChar(' ', LineLevel), Copy(Text, Pos('=', Text) + 1, length(Text))]);
        end else
          LogText := '';
        Writeln(LogFile, LogText);
      finally
        CloseFile(LogFile);
      end;
    end else
    begin
      { Si pas de log métier, on écrit dans le log windows uniquement pour le débug }
      if (LogValues.DebugEvents > 0) and (pos(WSCDS_DebugMsg, Text) > 0 ) then
        WriteWindowsLog;
    end;
  end;

begin
  case TypeDebug of
    ssbylLog     : WriteUserLog;
    ssbylWindows : WriteWindowsLog;
    ssbylAll     : begin
                     if LogValues.LogLevel > 0 then
                       WriteUserLog;
                     WriteWindowsLog;
                   end;
  end;
end;

end.

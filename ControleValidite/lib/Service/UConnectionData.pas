unit UConnectionData;

interface
uses Messages,forms,registry,Windows,ShlObj,SysUtils,IdHashCRC;

function CalculCRC (P1: TDateTime;P2 : Integer) : Double;

procedure CreateInfosregistry;
//
procedure CreateRegDWord ( Clef : string);
procedure CreateRegInteger ( Clef : string);
procedure CreateRegString ( Clef : string);
procedure CreateRegDate ( Clef : string);
//
function GetInfosRegDWord ( Clef : string) : double;
function GetInfosRegInteger ( Clef : string) : integer ;
function GetInfosRegString ( Clef : string) : string;
function GetInfosRegDate ( Clef : string) : TdateTime;

function ACCESConnectionString(Database,Passw : string) : string;

implementation
uses DateUtils, ComObj;

function ACCESConnectionString(Database,Passw : string) : string;
begin
	result := 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+DataBase+
  ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
  'Jet OLEDB:Database Password="'+Passw+'";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;'+
  'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";'+
  'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;'+
  'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False';
end;

function GetInfosRegString ( Clef : string) : string;
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      result := TT.ReadString (Clef);
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

function GetInfosRegDWord ( Clef : string) : double;
var TT : TRegistry;
begin
  Result := 0;
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      result := TT.ReadFloat(Clef);
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

function GetInfosRegDate ( Clef : string) : TDateTime;
var TT : TRegistry;
begin
  Result := StrToDateTime('01/01/1900 00:00:00');
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      result := TT.ReadDateTime(Clef);
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

function GetInfosRegInteger ( Clef : string) : integer ;
var TT : TRegistry;
begin
  Result := 0;
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      result := TT.ReadInteger (Clef);
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

procedure CreatelesKeys ;
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    TT.OpenKey('\SOFTWARE\LSE',true);
    TT.CloseKey;

  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    TT.CloseKey;

  FINALLY
    TT.Free;
  END;
end;



procedure CreateInfosregistry;
begin
  CreatelesKeys;
  CreateRegString('CodeClient');
  CreateRegString('LSEAdress');
  CreateRegDate('LastDate');
  CreateRegInteger('Counter');
  CreateRegDWord ('CRC');
end;

procedure CreateRegString ( Clef : string);
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      if not (TT.KeyExists(Clef)) then
      begin
        TT.WriteString(Clef,'');
      end;
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

procedure CreateRegDate ( Clef : string);
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      if not (TT.KeyExists(Clef)) then
      begin
        TT.WriteDateTime (Clef,StrToDate('01/01/1900'));
      end;
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

procedure CreateRegInteger ( Clef : string);
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      if not (TT.KeyExists(Clef)) then
      begin
        TT.WriteInteger (Clef,0);
      end;
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

procedure CreateRegDWord ( Clef : string);
var TT : TRegistry;
begin
	TT := TRegistry.Create;
  TRY
  	TT.RootKey := HKEY_LOCAL_MACHINE;
    if TT.OpenKey('\SOFTWARE\LSE\SvcAutorise',true) then
    begin
      if not (TT.KeyExists(Clef)) then
      begin
        TT.WriteFloat (Clef,0);
      end;
    end;
    TT.CloseKey;
  FINALLY
    TT.Free;
  END;
end;

function GetAppData : string;
var RecPath : PChar;
begin
  RecPath := StrAlloc(MAX_PATH);
  try
    FillChar(RecPath^, MAX_PATH, 0);
    if SHGetSpecialFolderPath(0, RecPath, CSIDL_APPDATA, false) then
    	result := RecPath
    else result := '';
  finally
  	StrDispose(RecPath);
  end;
end;

function CalculCRC (P1: TDateTime;P2 : Integer) : Double;
var fLocal : string;
begin
  flocal := DateTimeToStr(IncDay(P1,P2));
  with TIdHashCRC32.Create do begin
    Result := HashValue(fLocal);
    Free;
  end;
end;

end.

unit Uregistry;

interface
uses Registry,Windows;

function IsVersionDispo : string;
function GetInfoStocke ( Clef : string) : string;
procedure SetInfoStocke (Clef,Valeur : string);
procedure SetversionDispo (dispo : boolean);
function GetCegidInstallFolder : string;
function GetLSEInstallFolder : string;
function GetLocalSpecialFolder(folder:string) :string;
function GetCommonSpecialFolder(folder:string) :string;
function GetTempDir: string;

implementation


function GetTempDir: string;
begin;
  SetLength(Result, 255);
  SetLength(Result, GetTempPath(255, (PChar(Result))));
end;

function IsVersionDispo : string;
var Reg : TRegistry;
		res : string;
begin
	Reg := TRegistry.Create;
  result := '';
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.KeyExists('SOFTWARE\LSE\AutoUpdate') then
    begin
      if Reg.OpenKey('SOFTWARE\LSE\AutoUpdate', False) then
      begin
        if Reg.ValueExists('ServerUpdated') then
          res := Reg.ReadString('ServerUpdated');
        if res = '1' then result := 'S';
      end;
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

function GetInfoStocke ( Clef : string) : string;
var Reg : TRegistry;
begin
	Reg := TRegistry.Create;
  result := '';
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.KeyExists('SOFTWARE\LSE\AutoUpdate') then
    begin
      if Reg.OpenKey('SOFTWARE\LSE\AutoUpdate', False) then
      begin
        if Reg.ValueExists(Clef) then
        begin
          result := Reg.ReadString(Clef);
        end;
      end;
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

procedure SetInfoStocke (Clef,Valeur : string);
var Reg : TRegistry;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if not Reg.KeyExists('SOFTWARE\LSE\AutoUpdate') then
    begin
    	reg.CreateKey('SOFTWARE\LSE\AutoUpdate');
    end;
    if Reg.OpenKey('SOFTWARE\LSE\AutoUpdate', False) then
    begin
    	Reg.WriteString(Clef,Valeur);
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

procedure SetversionDispo (dispo : boolean);
var Reg : TRegistry;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\LSE\AutoUpdate', False) then
    begin
    	if dispo then reg.WriteFloat('ServerUpdated',1)
      				 else reg.WriteFloat('ServerUpdated',0);
    end;
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
end;

function GetCegidInstallFolder : string;
var Reg : TRegistry;
		res : string;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\Cegid\Cegid business\', False)
    then res := Reg.ReadString('INSTALLDIR')
    else res := '';
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
  result := res;
end;

function GetLSEInstallFolder : string;
var Reg : TRegistry;
		res : string;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\Cegid\Cegid business\', False)
    then res := Reg.ReadString('INSTDIRLSE')
    else res := '';
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
  result := res;
end;

function GetLocalSpecialFolder(folder:string) :string;
var Reg : TRegistry;
		res : string;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', False)
    then res := Reg.ReadString(folder)
    else res := '';
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
  result := res;
end;

function GetCommonSpecialFolder(folder:string) :string;
var Reg : TRegistry;
		res : string;
begin
	Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if Reg.OpenKey('Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders', False)
    then res := Reg.ReadString(folder)
    else res := '';
  finally
    Reg.CloseKey;
    Reg.Free;
  end;
  result := res;
end;

end.

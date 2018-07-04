unit UUtilSeria;

interface
uses Windows,Messages,Registry,SysUtils,Nb30 ;

type

TMACAddress = packed array[0..5] of Byte;
TAStat = record
            Adapt: TAdapterStatus;
            NameBuff: array[0..30] of TNameBuffer;
				 end;


function SeriaServerpresent : Boolean;
function ComputerName: string;
function GetAdrSeria : string;
function DeleteSeria : Boolean;

implementation

function GetMacAddress(AdapterNum: Integer): TMACAddress;
var  Ncb: TNCB;
  uRetCode: Char;
  j: Integer;
  Adapter: TAStat;
begin
  FillChar(NCB, SizeOf(NCB), 0);
  with NCB do
  begin
    ncb_command := Char(NCBRESET);
    ncb_lana_num := Char(AdapterNum);
  end;
  uRetCode := NetBios(@Ncb);
  if uRetCode <> #0 then raise Exception.CreateFmt('NetBIOS error code = %f', [Ord(uRetCode)]);
  FillChar(NCB, SizeOf(NCB), 0);
  with NCB do
  begin
    ncb_command := Char(NCBASTAT);
    ncb_lana_num := Char(AdapterNum);
    StrCopy(ncb_callname, '* ');
    ncb_buffer := @Adapter;
    ncb_length := SizeOf(Adapter);
  end;
  uRetCode := NetBios(@Ncb);
  if uRetCode <> #0 then
  raise Exception.CreateFmt('NetBIOS error code = %f', [Ord(uRetCode)]);
  for j := 0 to 5 do Result[j] := Ord(Adapter.Adapt.Adapter_address[j]);
end;

function LanAdapterCount: Integer;
var
  Ncb: TNCB;
  uRetCode: Char;
  lEnum: TLanaEnum;
begin
  FillChar(NCB, SizeOf(NCB), 0);
  with NCB do
  begin
    ncb_command := Char(NCBENUM);
    ncb_buffer := @lEnum;
    ncb_length := SizeOf(lEnum);
  end;
  uRetCode := Netbios(@Ncb);
  if uRetCode <> #0 then
  raise Exception.CreateFmt('NetBIOS error code = %f', [Ord(uRetCode)] );
  Result := Ord(lenum.length);
end;


function GetAdressMac : string;
var MAC: TMACAddress;
			i: Integer;
			s: string;
begin
  MAC := GetMacAddress(LanAdapterCount());
  s := '';
  for i := 0 to 5 do
  s := s + IntToHex(MAC[i], 2) + ' ';
  Result := s;
end;

function SeriaServerpresent : Boolean;
var r : TRegistry;
begin
  Result := False;
  r := TRegistry.Create ;
  r.RootKey := HKEY_LOCAL_MACHINE ;
  if r.OpenKeyReadOnly('SOFTWARE\CEGID_RM\CegidSeria') then
  begin
    Result :=  r.KeyExists('Version5');
    r.CloseKey ;
  end;
  r.Free ;
end;


function GetAdrSeria : string;
var r : TRegistry;
begin
  Result := '';
  r := TRegistry.Create ;
  r.RootKey := HKEY_LOCAL_MACHINE ;
  if r.OpenKeyReadOnly('SOFTWARE\CEGID_RM\CegidSeria\Version5') then
  begin
    Result :=  r.ReadString('RefKey');
    if Result = '' then
    begin
      Result := GetAdressMac;
    end;
    r.CloseKey ;
  end;
  r.Free ;
end;

function DeleteSeria : Boolean;
var r : TRegistry;
begin
	result := false;
  r := TRegistry.Create ;
  r.RootKey := HKEY_LOCAL_MACHINE ;
  if r.OpenKey('SOFTWARE\CEGID_RM\CegidSeria\Version5',false) then
  begin
    if r.ValueExists('Client') then
    begin
    	Result :=  r.DeleteValue  ('Client');
    end else if r.ValueExists('Server') then
    begin
    	Result :=  r.DeleteValue ('Server');
    end;
    r.CloseKey ;
  end;
  r.Free ;
end;

function ComputerName: string;
var
  lpBuffer: array[0..MAX_COMPUTERNAME_LENGTH] of char;
  nSize: dword;
begin
  nSize:= Length(lpBuffer);
  if GetComputerName(lpBuffer, nSize) then
    result:= lpBuffer
  else
    result:= '';
end;

end.

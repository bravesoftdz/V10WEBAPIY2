unit Uconstantes;
interface

uses windows,commctrl;

const
  {$EXTERNALSYM BCM_FIRST}
  BCM_FIRST               = $1600;      // Button control messages
  CONFLSEFILE = 'LSEUPDATE.conf';
  CONFCEGIDFILE = 'CEGIDUPDATE.conf';
  PARAMUPDATE = 'DATAUPDATE.conf';
  CONFCOMMONFILE = 'COMMONUPDATE.conf';
  CONFSTDFILE = 'STDUPDATE.conf';
  CONFUPDFILE = 'UPDUPDATE.conf';
  NAMEPARAM = 'PARAMUPDATE.LSE';
  SEPARATOR = '|';
  SECURITY_NT_AUTHORITY: TSIDIdentifierAuthority = (Value: (0, 0, 0, 0, 0, 5));
  SECURITY_BUILTIN_DOMAIN_RID = $00000020;
  DOMAIN_ALIAS_RID_ADMINS = $00000220;

  BCM_SETSHIELD = BCM_FIRST + $000C;

function Button_SetElevationRequiredState(wnd: HWND; fRequired: bool): LRESULT;

implementation

function Button_SetElevationRequiredState(wnd: HWND; fRequired: bool): LRESULT;
begin
  Result := SendMessage(wnd, BCM_SETSHIELD, 0, LPARAM(fRequired));
end;

end.

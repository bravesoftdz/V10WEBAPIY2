unit Udefinitions;

interface
uses classes,Registry,Windows,SysUtils;
type
    TActionFiche = (taCreat,TaModif);

function getCheminIni : string;

implementation

function getCheminIni : string;
var iReg : TregInifile;
begin
  Ireg := TRegIniFile.Create;
  IReg.RootKey := HKEY_LOCAL_MACHINE;
  result := IncludeTrailingPathDelimiter(IReg.ReadString('SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders','Common AppData','C:\ProgramData\Cegid\LSE Business\WebServices'));
  FreeAndNil(iReg);
  result := result+'CEGID\LSE Business\WebServices';
end;

end.

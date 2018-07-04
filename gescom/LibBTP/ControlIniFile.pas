unit ControlIniFile;

interface
uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
	windows,
  uhttp,paramsoc,
  inifiles,
  CBPPath,forms,
  utob,uEntCommun,Classes;

const INIFILE : string = 'CEGIDPGI.INI';
			PRESOCREF = 'socref.mdb';
      NEWSOCREF = 'SocRefbtp.mdb';
      DRIVER2008 = 'ODBC_MSSQL2008';
      DRIVER2005 = 'ODBC_MSSQL2005';

function FindPGIIniFile : string;
procedure ControlPGIINI;

implementation

function WindowsDirectory: String;
const
  dwLength: DWORD = 255;
var
  WindowsDir: PChar;
begin
  GetMem(WindowsDir, dwLength);
  GetWindowsDirectory(WindowsDir, dwLength);
  Result := String(WindowsDir);
  FreeMem(WindowsDir, dwLength);
end;

function FindPGIIniFile : string;
  var Localrepert : string;
  		Iexist : integer;
begin
	Iexist := -1;
  // recherche sur repertoire courant
  TRY
  	Result := '';
    Localrepert := ExtractFilePath (Application.ExeName); // repertoire de l'application
    Iexist := FileOpen(IncludeTrailingBackslash(Localrepert)+INIFILE,fmOpenRead );
    if Iexist <0 then
    begin
      Localrepert := TCbpPath.GetCegidUserRoamingAppData; // repertoire Utilisateur
      Iexist := FileOpen(IncludeTrailingBackslash(Localrepert)+INIFILE,fmOpenRead );
      if Iexist <0 then
      begin
        Localrepert := TCBPPath.GetCegidData; // Data de All Users
        Iexist := FileOpen(IncludeTrailingBackslash(Localrepert)+INIFILE,fmOpenRead);
        if Iexist < 0 then
        begin
           LocalRepert := WindowsDirectory;         // Emplacement de windows
           Iexist := FileOpen(IncludeTrailingBackslash(Localrepert)+INIFILE,fmOpenRead);
           if Iexist >= 0 then
           begin
              Result := Localrepert;
           end;
        end else
        begin
          Result := Localrepert;
        end;
      end else
      begin
        Result := Localrepert;
      end;
    end else
    begin
      Result := Localrepert;
    end;
  FINALLY
   if Iexist >= 0 then FileClose(Iexist);
  END;
end;

procedure  ControleDriverVers (lesSection : TStringList; IniPgiFile : Tinifile);
var Indice : integer;
    LaSection,LeDriver : string;
begin
	For Indice:= 0 to lesSection.Count -1 do
  begin
    laSection := lesSection.Strings [Indice];
    LeDriver := IniPgiFile.ReadString(LaSection,'Driver','');
		if UpperCase(leDriver) = DRIVER2008 then
    begin
      IniPgiFile.DeleteKey(LaSection, 'Driver');
      IniPgiFile.WriteString(LaSection,'Driver',DRIVER2005);
    end;
  end;
end;

procedure ControlPGIINI;
var	lesSections : Tstringlist;
{$IFNDEF EAGLCLIENT}
  var RepertIniFile,PreDbffile,NewDbffile : string;
  		preDbfName : string;
      IniPGIFile : Tinifile;
{$ENDIF}
begin
	INIFILE := HalSocIni;
{$IFNDEF EAGLCLIENT}
	lesSections := TStringList.Create;
	RepertIniFile :=  FindPGIIniFile;
	//
  if RepertIniFile <> '' then
  begin
    IniPgiFile := TiniFile.create (IncludeTrailingBackslash (RepertInifile)+INIFILE);
    (*
    PreDbffile := IniPgiFile.ReadString('Reference','Database','');
    PreDbfName := ExtractFileName(preDbffile);
    if (LowerCase (PreDbfName) = PRESOCREF) then
    begin
    	NewDbffile := IncludeTrailingBackslash(ExtractFilePath(PreDbffile))+NEWSOCREF;
      if not FileExists (NewDbffile) then  RenameFile(PreDbffile,NewDbffile);
      IniPgiFile.DeleteKey('Reference', 'Database');
      IniPgiFile.WriteString('Reference','Database',NewDbffile);
    end;
    *)
    IniPGIFile.ReadSections(lesSections);
//    ControleDriverVers (lesSections,IniPgiFile);
    IniPGIFile.Free;
  end;
  lesSections.free;
{$ENDIF}
end;

end.

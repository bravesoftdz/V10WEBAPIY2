unit UFolders;

interface
uses windows,SHFolder;

const

// - Ajout Vista et Seven -----------------
	CSIDL_COMMON_APPDATA          = $0023;
  CSIDL_WINDOWS                 = $0024;
  CSIDL_SYSTEM                  = $0025;
  CSIDL_PROGRAM_FILES           = $0026;
  CSIDL_MYPICTURES              = $0027;
  CSIDL_PROFILE                 = $0028;
  CSIDL_SYSTEMX86               = $0029;
  CSIDL_PROGRAM_FILESX86        = $002A;     // application 32 bits sur OS version 64 bits
  CSIDL_PROGRAM_FILES_COMMON    = $002B;
  CSIDL_PROGRAM_FILES_COMMONX86 = $002C;     // application 32 bits sur OS version 64 bits
  CSIDL_COMMON_TEMPLATES        = $002D;
  CSIDL_COMMON_DOCUMENTS        = $002E;
  CSIDL_COMMON_ADMINTOOLS       = $002F;
  CSIDL_ADMINTOOLS              = $0030;
  CSIDL_FLAG_CREATE             = $8000;
  CSIDL_FLAG_DONT_VERIFY        = $4000;
  CSIDL_FLAG_MASK               = $FF00;
// ---------------------------------------

function GetSpecialFolders(CSIDL : integer) : string;

implementation

function GetSpecialFolders(CSIDL : integer) : string;
const SHGFP_TYPE_CURRENT = 0;
var path: array [0..MAX_PATH] of char;
begin
  if SUCCEEDED(SHGetFolderPath(0,CSIDL,0,SHGFP_TYPE_CURRENT,@path[0])) then
  	Result := path
  else
  	Result := '';
end;

end.
 
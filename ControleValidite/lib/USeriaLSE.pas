unit USeriaLSE;

interface
uses Windows,CBPPath,Forms;

function LoadLibrarySeriaLSE : THandle;
function GetDirLsePlugins : string;
procedure FreeLibrarySeriaLSE(Handle : THandle);

implementation

uses SysUtils;

function GetDirLsePlugins : string;
begin
  result := IncludeTrailingBackslash( ExtractFilePath(Application.ExeName))+ 'Plugins';
end;

function LoadLibrarySeriaLSE : THandle;
var Dll : string;
begin
  Dll := IncludeTrailingBackslash(GetDirLsePlugins)+'SeriaServeurLse.dll';
  //Chargement de la DLL. On récupère son handle
  result := LoadLibrary(PAnsiChar(Dll));
end;

procedure FreeLibrarySeriaLSE(Handle : THandle) ;
begin
	FreeLibrary(Handle)
end;

end.

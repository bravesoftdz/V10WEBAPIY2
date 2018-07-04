unit InfoVersion;

interface

uses
  Windows, HCtrls,SysUtils, Forms;


type
  TVersionResources = (vrCompanyName, vrFileDescription, vrFileVersion,
                         vrInternalName, vrLegalCopyright, vrOriginalFilename,
                         vrProductName, vrProductVersion, vrComments, vrFlags);

  TInfoVersion = class (TObject) 
  private
      FLangCharset: string;
      function GetStringFileInfo(VersionResourceKey:string; Buffer: Pchar; size: integer): string;
      Procedure GetInfo;
  public
    { Public declarations }
    FileVersion :string ;
    Comments:string;
    VersionSocRef:string;
    VMajeure,VMineure,SousVersion,Build:string;
    Erreur:string;
    constructor Create;
  end;

    {The order of this array must be the same as the VersionResources
    enum type as that is used for the index lookup}
{    VersionLookup: array[TVersionResources, 0..1] of string = (
                    ('CompanyName', 'Company Name:'),
                    ('FileDescription', 'File Description:'),
                    ('FileVersion', 'File Version:'),
                    ('InternalName', 'Internal Name:'),
                    ('LegalCopyright', 'Legal Copyright:'),
                    ('OriginalFilename', 'Original Filename:'),
                    ('ProductName', 'Product Name:'),
                    ('ProductVersion', 'Product Version:'),
                    ('Comments', 'Comments:'),
                    ('Flags', 'Flags:'));
}

Function IVGetVersion:string;
Function IVGetBuild:string;
Function IVGetSocRef:Integer;

implementation
Function IVGetVersion:string;
var TVersion : TInfoVersion;
begin
TVersion:=TInfoVersion.create;
with TVersion do
      result:=Vmajeure+'.'+VMineure+SousVersion ;
TVersion.free;
end;

Function IVGetBuild:string;
var TVersion : TInfoVersion;
begin
TVersion:=TInfoVersion.create;
with TVersion do
      result:=Build;
TVersion.free;
end;

Function IVGetSocRef:integer;
var TVersion : TInfoVersion;
begin
TVersion:=TInfoVersion.create;
result:=0;
with TVersion do
      if copy(VersionSocRef,1,1)<>'<' then result:=strToInt(VersionSocRef) ;
TVersion.free;
end;


constructor TInfoVersion.Create;
begin
    inherited Create;
    //WordWrap := false;
    //Autosize := true;
    //ShowInfoPrefix := true;
    FLangCharset :='-1';   {-1 = auto detect}
    //VersionResource := vrFileVersion;
    Erreur:='';
    GetInfo;
end;

function TInfoVersion.GetStringFileInfo(VersionResourceKey:string;Buffer: Pchar; size: integer): string;
var vallen, Translen: DWORD;
    VersionPointer, TransBuffer: pchar;
    Temp: integer;
    CalcLangCharSet: string;
begin
    if FLangCharSet = '-1' then
    begin
        VerQueryValue(buffer, '\VarFileInfo\Translation',
                        pointer(TransBuffer), TransLen);
        if TransLen >= 4 then
        begin
            StrLCopy(@temp, TransBuffer, 2);
            CalcLangCharSet:=IntToHex(temp, 4);
            StrLCopy(@temp, TransBuffer+2, 2);
            CalcLangCharSet := CalcLangCharSet+IntToHex(temp, 4);
            FLangCharSet := CalcLangCharSet;
        end
        else
        begin
            Result := '< Error retrieving translation info >';
            exit;
        end;
    end;

    if VerQueryValue(buffer, pchar('\StringFileInfo\'+FLangCharSet+'\'+
                     VersionResourceKey),
                     pointer(VersionPointer), vallen) then
    begin
        if (Vallen > 1) then
        begin
            SetLength(Result, vallen);
            StrLCopy(Pchar(Result), VersionPointer, vallen);
        end
        else Result := '< No Version Info >';
    end
    else result := '< Error retrieving version info >';
end;

procedure TInfoVersion.GetInfo;
var dump, size: DWORD;
    buffer: pchar;
    stVersion:string;
begin
size := GetFileVersionInfoSize(pchar(Application.Exename), dump);
if  size = 0 then
    begin
    Erreur := '< No Data Available >';
    end
else
    begin
    buffer := StrAlloc(size+1);
    try
       if not GetFileVersionInfo(Pchar(Application.Exename), 0,size, buffer) then
          Erreur := '< Error retrieving version info >'
       else
           begin
           FileVersion := GetStringFileInfo('FileVersion',buffer, size);
           Comments := GetStringFileInfo('Comments',buffer, size);
           VersionSocRef := GetStringFileInfo('VersionSocRef',buffer, size);

           stVersion:=FileVersion;
           VMajeure:=readtokenPipe(stVersion,'.');
           VMineure:=readtokenPipe(stVersion,'.');
           SousVersion:=readtokenPipe(stVersion,'.');
           Build:=readtokenPipe(stVersion,'.');
           end;
    finally
           StrDispose(Buffer);
    end;
    end;
end;

end.
 
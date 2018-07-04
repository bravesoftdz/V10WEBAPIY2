unit cdBurnStream;

interface

uses ActiveX, ComObj, Classes, Math, SysUtils, Windows;

Type
  TStorage = Class ;

  IStorage_EX = interface
  ['{E81D59A5-CC55-41A2-B51E-01CD7A28CEB1}']
    procedure CreateFileStream(filename, streamName: string;openMode:Word);
    function CreateStorageDirectory(storageName: string): IStorage;
    function GetSubStorageCount:integer;
    function GetFileCount:integer;
    function GetFileName(i:integer):string;
    function GetFileStorageName(i:integer):string;
    function GetStreamByIndex(i:integer):IStream;
    function GetStorageByIndex(i:integer):IStorage;
  end;


  TStreamAdapterEX  = class(TStreamAdapter)
  private
    FStreamName:String ;
  public
    constructor Create(Stream: TStream;StreamName:String);reintroduce;
    function Stat(out statstg: TStatStg;
      grfStatFlag: Longint): HResult; override; stdcall;
    function GetStreamName:string ;
  end ;


  TStorage = class(TInterfacedObject, IStorage,IStorage_EX)
  private

    FSubStorageIndex: TStringList;
    FSubStorages:IInterfaceList;

    FStreamNameIdx:TStringList;
    FFileNameIdx:TStringList;

    fStatstg: TStatStg;

    procedure CreateFileStream(filename, streamName: string;openMode:word);
    function GetFileCount: integer;
    function GetFileName(i: integer): string;
    function GetFileStorageName(i: integer): string;
    function GetStorageByIndex(i: integer): IStorage;
    function GetStreamByIndex(i: integer): IStream;
    function GetSubStorageCount: integer;
    function CreateStorageDirectory(storageName: string): IStorage;

  public
    constructor Create(name: string);
    destructor Destroy; override;

    function CreateStream(pwcsName: POleStr; grfMode: Longint; reserved1: Longint;
      reserved2: Longint; out stm: IStream): HResult; stdcall;
    function OpenStream(pwcsName: POleStr; reserved1: Pointer; grfMode: Longint;
      reserved2: Longint; out stm: IStream): HResult; stdcall;
    function CreateStorage(pwcsName: POleStr; grfMode: Longint;
      dwStgFmt: Longint; reserved2: Longint; out stg: IStorage): HResult;
      stdcall;
    function OpenStorage(pwcsName: POleStr; const stgPriority: IStorage;
      grfMode: Longint; snbExclude: TSNB; reserved: Longint;
      out stg: IStorage): HResult; stdcall;
    function CopyTo(ciidExclude: Longint; rgiidExclude: PIID;
      snbExclude: TSNB; const stgDest: IStorage): HResult; stdcall;
    function MoveElementTo(pwcsName: POleStr; const stgDest: IStorage;
      pwcsNewName: POleStr; grfFlags: Longint): HResult; stdcall;
    function Commit(grfCommitFlags: Longint): HResult; stdcall;
    function Revert: HResult; stdcall;
    function EnumElements(reserved1: Longint; reserved2: Pointer; reserved3: Longint;
      out enm: IEnumStatStg): HResult; stdcall;
    function DestroyElement(pwcsName: POleStr): HResult; stdcall;
    function RenameElement(pwcsOldName: POleStr;
      pwcsNewName: POleStr): HResult; stdcall;
    function SetElementTimes(pwcsName: POleStr; const ctime: TFileTime;
      const atime: TFileTime; const mtime: TFileTime): HResult;
      stdcall;
    function SetClass(const clsid: TCLSID): HResult; stdcall;
    function SetStateBits(grfStateBits: Longint; grfMask: Longint): HResult;
      stdcall;
    function Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
      stdcall;


  end;

  TEnumStorageElements = class(TInterfacedObject, IEnumStatStg)
  private
    fStorage: IStorage;
    fStreamOffset, fStorageOffset: Integer;
  public
    constructor Create(storage: IStorage);

    function NextOne(out stat: TStatStg): boolean;
    function Next(celt: Longint; out elt; pceltFetched: PLongint): HResult; stdcall;
    function Skip(celt: Longint): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out enm: IEnumStatStg): HResult; stdcall;
  end;

implementation




{ TStreamAdapterEX }

constructor TStreamAdapterEX.Create(Stream: TStream; StreamName: String);
begin
  inherited create(Stream,soOwned);
  FStreamName := StreamName;

end;

function TStreamAdapterEX.GetStreamName: string;
begin
  result := FStreamName ;
end;

function TStreamAdapterEX.Stat(out statstg: TStatStg;
  grfStatFlag: Integer): HResult;
var
  fileInfo: TByHandleFileInformation;
begin
  if (Assigned(@statstg)) then
  begin

    if ((grfStatFlag = STATFLAG_DEFAULT) or (grfStatFlag = STATFLAG_NONAME)) then
    begin
      try
        FillChar(statstg, sizeof(TStatStg), 0);
        statstg.dwType := STGTY_STREAM;
        statstg.cbSize := Stream.Size;

        if (GetFileInformationByHandle(TFileStream(Stream).Handle, fileInfo)) then
        begin
          statstg.mtime := fileInfo.ftLastWriteTime;
          statstg.ctime := fileInfo.ftCreationTime;
          statstg.atime := fileInfo.ftLastAccessTime;
        end;

        if (grfStatFlag <> STATFLAG_NONAME) then
        begin
          statstg.pwcsName := CoTaskMemAlloc(sizeof(WideChar) * (length(fStreamName) + 1));
          { TODO -oansonh : Check return value of CoTaskMemAlloc }
          StringToWideChar(fStreamName, statstg.pwcsName, length(fStreamName) + 1);
        end;

        Result := S_OK;
      except
        Result := E_UNEXPECTED;
      end;
    end
    else
      Result := STG_E_INVALIDFLAG;
  end
  else
    Result := STG_E_INVALIDPOINTER;
end;



 {TStorage}

function TStorage.GetStreamByIndex(i: integer): IStream;
var openMode:word;
begin
  openMode := fmOpenRead or fmShareDenyWrite; //  Word(FFileNameIdx.Objects[i]) ;
  result := TStreamAdapterEX.Create(TFileStream.Create(FFileNameIdx[i],openMode),FStreamNameIdx[i]);
end;



procedure TStorage.CreateFileStream(filename, streamName: string;openMode:word);
begin
  FStreamNameIdx.AddObject(streamName,TObject(FFileNameIdx.AddObject(FileName,TObject(OpenMode) )));
end;

function TStorage.GetFileCount: integer;
begin
  result := FFileNameIdx.Count;
end;

function TStorage.GetFileName(i: integer): string;
begin
  result := FFileNameIdx[i];
end;

function TStorage.GetFileStorageName(i: integer): string;
begin
  result := FSubStorageIndex[i];
end;

function TStorage.GetStorageByIndex(i: integer): IStorage;
begin
  result := FSubStorages.Items[i] as IStorage;
end;

function TStorage.GetSubStorageCount: integer;
begin
  result := FSubStorageIndex.Count;
end;

constructor TStorage.Create(name: string);
begin
  inherited Create;
  FStreamNameIdx:=TStringList.Create;
  FFileNameIdx:=TStringList.Create;
  FSubStorageIndex := TStringList.Create;
  FSubStorages := TInterfaceList.Create;

  FillChar(fStatstg, sizeof(TStatStg), 0);

  fStatstg.pwcsName := CoTaskMemAlloc(sizeof(WideChar) * (length(name) + 1));
  StringToWideChar(name, fStatstg.pwcsName, length(name) + 1);

  fStatstg.dwType := STGTY_STORAGE;
end;

destructor TStorage.Destroy;
begin
  FStreamNameIdx.Free;
  FFileNameIdx.Free;
  FSubStorageIndex.Free;
  FSubStorages := nil;
  CoTaskMemFree(fStatstg.pwcsName);

  inherited Destroy;
end;

function TStorage.CreateStream(pwcsName: POleStr; grfMode: Longint; reserved1: Longint;
  reserved2: Longint; out stm: IStream): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.OpenStream(pwcsName: POleStr; reserved1: Pointer; grfMode: Longint;
  reserved2: Longint; out stm: IStream): HResult;
var
  streamName: string;
  streamIndex: Integer;
  fileIndex:integer;
  OpenMode:word ;
begin
  stm := nil;

  if (Assigned(pwcsName)) then
  begin
    streamName := WideCharToString(pwcsName);
    streamIndex := FStreamNameIdx.IndexOf(streamName);
    fileIndex   :=  Integer(FStreamNameIdx.Objects[StreamIndex]);
    if (streamIndex <> -1) then
    begin
      OpenMode := fmOpenRead or fmShareDenyWrite; // Word(FFileNameIdx.Objects[fileIndex]) ;
      stm := TStreamAdapter.Create(TFileStream.Create(FFileNameIdx[fileIndex],OpenMode ),soOwned);
      Result := S_OK;
    end
    else
      Result := STG_E_FILENOTFOUND;
  end
  else
    Result := STG_E_INVALIDPOINTER;
end;

function TStorage.CreateStorage(pwcsName: POleStr; grfMode: Longint;
  dwStgFmt: Longint; reserved2: Longint; out stg: IStorage): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.CreateStorageDirectory(storageName: string): IStorage;
begin
  result := TStorage.Create(storageName);
  FSubStorageIndex.AddObject(storageName,TObject( FSubStorages.Add(result) ));
end;

function TStorage.OpenStorage(pwcsName: POleStr; const stgPriority: IStorage;
  grfMode: Longint; snbExclude: TSNB; reserved: Longint; out stg: IStorage): HResult;
var
  storageName: string;
  storageIndex: Integer;
begin
  stg := nil;

  if (Assigned(pwcsName)) then
  begin
    storageName := WideCharToString(pwcsName);
    storageIndex := FSubStorageIndex.IndexOf(storageName);
    if (storageIndex <> -1) then
    begin
      stg := FSubStorages.Items[Integer(FSubStorageIndex.Objects[StorageIndex])] as IStorage;
      Result := S_OK;
    end
    else
      Result := STG_E_FILENOTFOUND;
  end
  else
    Result := STG_E_INVALIDPOINTER;
end;

function TStorage.CopyTo(ciidExclude: Longint; rgiidExclude: PIID;
      snbExclude: TSNB; const stgDest: IStorage): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.MoveElementTo(pwcsName: POleStr; const stgDest: IStorage;
  pwcsNewName: POleStr; grfFlags: Longint): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.Commit(grfCommitFlags: Longint): HResult;
begin
  Result := S_OK;
end;

function TStorage.Revert: HResult;
begin
  Result := S_OK;
end;

function TStorage.EnumElements(reserved1: Longint; reserved2: Pointer; reserved3: Longint;
  out enm: IEnumStatStg): HResult;
begin
  enm := TEnumStorageElements.Create(self);

  Result := S_OK;
end;

function TStorage.DestroyElement(pwcsName: POleStr): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.RenameElement(pwcsOldName: POleStr; pwcsNewName: POleStr): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.SetElementTimes(pwcsName: POleStr; const ctime: TFileTime;
  const atime: TFileTime; const mtime: TFileTime): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.SetClass(const clsid: TCLSID): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.SetStateBits(grfStateBits: Longint; grfMask: Longint): HResult;
begin
  Result := E_NOTIMPL;
end;

function TStorage.Stat(out statstg: TStatStg; grfStatFlag: Longint): HResult;
begin
  if (Assigned(@statstg)) then
  begin
    if (grfStatFlag <> STATFLAG_NONAME) then
    begin
      statstg.pwcsName := CoTaskMemAlloc((lstrlenW(fStatstg.pwcsName) + 1) * sizeof(WideChar));
      lstrcpyW(statstg.pwcsName, fStatstg.pwcsName);
    end
    else
      statstg.pwcsName := nil;

    statstg.dwType := fStatstg.dwType;
    statstg.cbSize := fStatstg.cbSize;
    statstg.mtime := fStatstg.mtime;
    statstg.ctime := fStatstg.ctime;
    statstg.atime := fStatstg.atime;
    statstg.grfMode := fStatstg.grfMode;
    statstg.grfLocksSupported := fStatstg.grfLocksSupported;
    statstg.clsid := fStatstg.clsid;
    statstg.grfStateBits := fStatstg.grfStateBits;
    statstg.reserved := fStatstg.reserved;

    Result := S_OK;
  end
  else
    Result := STG_E_INVALIDPOINTER;
end;

//////////////////////////////////////////////////////////////////////////////
////  TEnumStorageElements implementation
/////////////////////////////////////////////////////////////////////////////


{TEnumStorageElements}

constructor TEnumStorageElements.Create(storage: IStorage);
begin
  fStorage := storage;
  fStreamOffset := 0;
  fStorageOffset := 0;
end;

function TEnumStorageElements.NextOne(out stat: TStatStg): boolean;
var iResult:HResult ;
    AStream:IStream ;
begin
  if (fStreamOffset < (fStorage as IStorage_EX).GetFileCount) then
  begin
    AStream :=  (fstorage as IStorage_EX).GetStreamByIndex(fStreamOffset);
    iResult :=  AStream.Stat(stat, STATFLAG_DEFAULT) ;
    if SUCCEEDED(iResult) then
    begin
      Result := true;
      Inc(fStreamOffset);
      AStream :=Nil ;
    end
    else
      Result := false;
  end
  else
    if (fStorageOffset < (fStorage as IStorage_EX).GetSubStorageCount) then
    begin
      if (SUCCEEDED((fStorage as IStorage_EX).GetStorageByIndex(fStorageOffset).Stat(stat, STATFLAG_DEFAULT))) then
      begin
        Result := true;
        Inc(fStorageOffset);
      end
      else
        Result := false;
    end
    else
      Result := false;
end;

function TEnumStorageElements.Next(celt: Longint; out elt; pceltFetched: PLongint): HResult;
var
  returned: Longint;
  index: integer;
  okay: boolean;
begin
  if (celt > 1) then
  begin
    Result := S_FALSE;
    Exit;
  end;

  returned := 0;

  for index := fStreamOffset to (fStorage as IStorage_EX).GetfileCount - 1 do
  begin
    if (returned >= celt) then
      break;

    okay := NextOne(TStatStg(elt));

    if (okay) then
    begin
      Inc(returned);
 //     Inc(PByteArray(elt), sizeof(TStatStg));
    end
    else
      break;
  end;

  for index := fStorageOffset to (fStorage as IStorage_EX).GetSubStorageCount - 1 do
  begin
    if (returned >= celt) then
      break;

    okay := NextOne(TStatStg(elt));

    if (okay) then
    begin
      Inc(returned);
 //     Inc(PByteArray(elt), sizeof(TStatStg));
    end
    else
      break;
  end;

  if (Assigned(pceltFetched)) then
    pceltFetched^ := returned;

  if (returned = celt) then
    Result := S_OK
  else
    Result := S_FALSE;
end;

function TEnumStorageElements.Skip(celt: Longint): HResult;
begin
  Result := E_NOTIMPL;
end;

function TEnumStorageElements.Reset: HResult;
begin
  fStorageOffset := 0;
  fStreamOffset := 0;

  Result := S_OK;
end;

function TEnumStorageElements.Clone(out enm: IEnumStatStg): HResult;
begin
  Result := E_NOTIMPL;
end;



{$IFDEF DEBUG}

{ HACK: Apparently ntdll.dll has a 'debug.break' located in a method that this component makes use of.
  Conseqeuently, when debugged the program will always stop, even when no breakpoint is aparent (encounters an int 3).
  The code below was taken from the net - can't guarantee whether it's the right fix, but it appears to work.  Only
  enabled for debug. }
procedure PatchINT3;
var
  NOP: Byte;
  BytesWritten: DWORD;
  NtDll: THandle;
  P: Pointer;
begin
  if (Win32Platform <> VER_PLATFORM_WIN32_NT) then
    Exit;

  NtDll := GetModuleHandle('NTDLL.DLL');

  if (NtDll = 0) then
    Exit;

  P := GetProcAddress(NtDll, 'DbgBreakPoint');
  if (P = nil) then
    Exit;

  try
    if (char(P^) <> #$CC) then
      Exit;

    NOP := $90;
    if (WriteProcessMemory(GetCurrentProcess, P, @NOP, 1, BytesWritten)) and
     (BytesWritten = 1) then
      FlushInstructionCache(GetCurrentProcess, P, 1);
    except on EAccessViolation do ;
    else raise;
  end;
end;





initialization
  PatchINT3;

{$ENDIF}

end.

unit cdBurnCom;

interface

uses Windows, ActiveX;

const
  OLE32 = 'ole32.dll';
  STGFMT_FILE = 3;
  STGFMT_ANY = 4;
  CLSID_MSDiscMasterObj: TGUID =  '{520CCA63-51A5-11D3-9144-00104BA11C5E}';
  IID_IStorage : TGUID =          '{0000000B-0000-0000-C000-000000000046}';
  IID_IDiscMaster: TGUID =        '{520CCA62-51A5-11D3-9144-00104BA11C5E}';
  IID_IDiscRecorder: TGUID =      '{85AC9776-CA88-4cf2-894E-09598C078A41}';
  IID_IJolietDiscMaster: TGUID =  '{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}';
  IID_IRedBookDiscMaster: TGUID = '{E3BC42CD-4E5C-11D3-9144-00104BA11C5E}';

function StgOpenStorageEx(const pwcsName: POleStr; grfMode: LongInt; stgfmt: DWORD; grfAttrs: DWORD; pStgOptions: Pointer;
  reserved2: Pointer; riid: PGUID; out stgOpen : IStorage ) : HResult; stdcall; external OLE32;
function StgCreateStorageEx(const pwcsName: POleStr; grfMode: LongInt; stgfmt: DWORD; grfAttrs: DWORD; pStgOptions: Pointer;
  reserved2: Pointer; riid: PGUID; out stgOpen: IStorage): HResult; stdcall; external OLE32;

type
  // Progress events

  TQueryCancel = procedure(bCancel: boolean) of object;
  TNotifyPnPActivity = procedure of object;
  TNotifyCDProgress = procedure(nCompletedSteps, nTotalSteps: Integer) of object;
  TNotifyEstimatedTime = procedure(nEstimatedSeconds: Integer) of object;
  TNotifyCompletionStatus = procedure(status: HResult) of object;

  // End Progress events

  ULONG = LongWord;
  PByte = ^Byte;
  PUINT = ^UINT;

  // Reviewed 8/31/2002 - changed WordBool (2 bytes), to boolean (1 byte).  Removed long typedef, use Integer instead.
  IDiscMasterProgressEvents = interface(IUnknown)
    ['{EC9E51C1-4E5D-11D3-9144-00104BA11C5E}']
    function QueryCancel(out pbCancel: boolean): HResult; stdcall;
    function NotifyPnPActivity: HResult; stdcall;
    function NotifyAddProgress(nCompletedSteps, nTotalSteps: Integer): HResult; stdcall;
    function NotifyBlockProgress(nCompleted, nTotal: Integer): HResult; stdcall;
    function NotifyTrackProgress(nCurrentTrack, nTotalTracks: Integer): HResult; stdcall;
    function NotifyPreparingBurn(nEstimatedSeconds: Integer): HResult; stdcall;
    function NotifyClosingDisc(nEstimatedSeconds: Integer): HResult; stdcall;
    function NotifyBurnComplete(status: HResult): HResult; stdcall;
    function NotifyEraseComplete(status: HResult): HResult; stdcall;
  end;

  // Reviewed 8/31/2002
  IEnumDiscMasterFormats = interface(IUnknown)
    ['{DDF445E1-54BA-11d3-9144-00104BA11C5E}']
    function Next(cFormats: ULONG; out lpiidFormatID; out pcFetched: ULONG): HResult; stdcall;
    function Skip(cFormats: ULONG): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppEnum: IEnumDiscMasterFormats): HResult; stdcall;
  end;

  // Reviewed 8/31/2002 - changed WordBool (2 bytes), to boolean (1 byte).  Removed long typedef and use Integer instead.
  IDiscRecorder = interface(IUnknown)
    ['{85AC9776-CA88-4cf2-894E-09598C078A41}']
    function Init(pbyUniqueID: pbyte; nulIDSize, nulDriveNumber: ULONG): HResult; stdcall;
    function GetRecorderGUID(var pbyUniqueID: byte; ulBufferSize: ULONG; out pulReturnSizeRequired: ULONG): HResult; stdcall;
    function GetRecorderType(out fTypeCode: Integer): HResult; stdcall;
    function GetDisplayNames(var pbstrVendorID: TBSTR; var pbstrProductID: TBSTR; var pbstrRevision: TBSTR): HResult; stdcall;
    function GetBasePnPID(out pbstrBasePnPID: WideString): HResult; stdcall;
    function GetPath(out pbstrPath: TBSTR): HResult; stdcall;
    function GetRecorderProperties(out ppPropStg: IPropertyStorage): HResult; stdcall;
    function SetRecorderProperties(pPropStg: IPropertyStorage): HResult; stdcall;
    function GetRecorderState(out pulDevStateFlags: ULONG): HResult; stdcall;
    function OpenExclusive: HResult; stdcall;
    function QueryMediaType(out fMediaType: Integer; out fMediaFlags: Integer): HResult; stdcall;
    function QueryMediaInfo(out pbSessions: byte; out pbLastTrack: byte; out ulStartAddress: ULONG; out ulNextWritable: ULONG; out ulFreeBlocks: ULONG): HResult; stdcall;
    function Eject: HResult; stdcall;
    function Erase(bFullErase: boolean): HResult; stdcall;
    function Close: HResult; stdcall;
  end;

  // Reviewed 8/31/2002
  IEnumDiscRecorders = interface(IUnknown)
    ['{9B1921E1-54AC-11d3-9144-00104BA11C5E}']
    function Next(cRecorders: ULONG; out ppRecorder; out pcFetched: ULONG): HResult; stdcall;
    function Skip(cRecorders: ULONG): HResult; stdcall;
    function Reset: HResult; stdcall;
    function Clone(out ppEnum: IEnumDiscRecorders): HResult; stdcall;
  end;

  // Reviewed 8/31/2002 - changed WordBool (2 bytes), to boolean (1 byte)
  IDiscMaster = interface(IUnknown)
    ['{520CCA62-51A5-11D3-9144-00104BA11C5E}']
    function Open: HResult; stdcall;
    function EnumDiscMasterFormats(out ppEnum: IEnumDiscMasterFormats): HResult; stdcall;
    function GetActiveDiscMasterFormat(out lpiid: TGUID): HResult; stdcall;
    function SetActiveDiscMasterFormat(var riid: TGUID; out ppUnk: IUnknown): HResult; stdcall;
    function EnumDiscRecorders(out ppEnum: IEnumDiscRecorders ): HResult; stdcall;
    function GetActiveDiscRecorder(out ppRecorder: IDiscRecorder):HResult; stdcall;
    function SetActiveDiscRecorder(pRecorder: IDiscRecorder): HResult; stdcall;
    function ClearFormatContent: HResult; stdcall;
    function ProgressAdvise(pEvents: IDiscMasterProgressEvents; out pvCookie: PUINT):HResult; stdcall;
    function ProgressUnadvise(vCookie: PUINT): HResult; stdcall;
    function RecordDisc(bSimulate, bEjectAfterBurn: boolean): HResult; stdcall;
    function Close: HResult; stdcall;
  end;

  // Reviewed 8/31/2002
  IJolietDiscMaster = interface(IUnknown)
    ['{E3BC42CE-4E5C-11D3-9144-00104BA11C5E}']
    function GetTotalDataBlocks(out pnBlocks: Integer): HResult; stdcall;
    function GetUsedDataBlocks(out pnBlocks: Integer): HResult; stdcall;
    function GetDataBlockSize(out pnBlockBytes: Integer): HResult; stdcall;
    function AddData(pStorage: IStorage; lFileOverwrite: Integer): HResult; stdcall;
    function GetJolietProperties(var ppPropStg: IPropertyStorage): HResult; stdcall;
    function SetJolietProperties(pPropStg: IPropertyStorage): HResult; stdcall;
  end;

   // Reviewed 8/31/2002 - Removed long typedef use Integer instead
  IRedbookDiscMaster = interface(IUnknown)
    ['{E3BC42CD-4E5C-11D3-9144-00104BA11C5E}']
    function GetTotalAudioTracks(out pnTracks: Integer): HResult; stdcall;
    function GetTotalAudioBlocks(out pnBlocks: Integer): HResult; stdcall;
    function GetUsedAudioBlocks(out pnBlocks: Integer): HResult; stdcall;
    function GetAvailableAudioTrackBlocks(out pnBlocks: Integer): HResult; stdcall;
    function GetAudioBlockSize(out pnBlockBytes: Integer): HResult; stdcall;
    function CreateAudioTrack(nBlocks: Integer): HResult; stdcall;
    function AddAudioTrackBlocks(var pby: byte; cb: Integer): HResult; stdcall;
    function CloseAudioTrack: HResult; stdcall;
  end;

//////////////////////////////////////////////////////////////////////////////////////////////////
//                                    End COM Stuff
//////////////////////////////////////////////////////////////////////////////////////////////////

implementation


end.

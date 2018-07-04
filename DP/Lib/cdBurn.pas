// LAST UPDATE 7:02a.m. 3/9/2003

// ©2003 Bayden Systems.  All rights reserved.
{//DEFINE _NOT_REGISTERED}

{ TODO -oansonh -cnext version : Set access, modified, and created times on directories }
{ TODO -oAnsonh -cFit and Finish : Factor property reads so that there isn't so much duplicated code }
unit cdBurn;

interface

uses
  Graphics, Controls, Forms, Windows, Messages, SysUtils, Classes,
  COMObj, shlObj, activex, Dialogs, cdBurnCom, AxCtrls, StdCtrls, ExtCtrls
  , cdBurnStream;


const
  WM_RECORDERCHANGE = WM_USER + 1;
  WM_ADDPROGRESS = WM_USER + 2;
  WM_BLOCKPROGRESS = WM_USER + 3;
  WM_TRACKPROGRESS = WM_USER + 4;
  WM_PREPARINGBURN = WM_USER + 5;
  WM_CLOSINGDISC = WM_USER + 6;
  WM_BURNCOMPLETE = WM_USER + 7;
  WM_ERASECOMPLETE = WM_USER + 8;
  WM_MEDIABLANK    = WM_USER+9;

  // { 08/27/2002 Anson - changed these constants to their correct values}
  IMAPI_S_PROPERTIESIGNORED     = 262656;
  IMAPI_S_BUFFER_TO_SMALL       = 262657;
  IMAPI_E_NOTOPENED             = -2147220981;
  IMAPI_E_NOTINITIALIZED        = -2147220980;
  IMAPI_E_USERABORT             = -2147220979;
  IMAPI_E_GENERIC               = -2147220978;
  IMAPI_E_MEDIUM_NOTPRESENT     = -2147220977;
  IMAPI_E_MEDIUM_INVALIDTYPE    = -2147220976;
  IMAPI_E_DEVICE_NOPROPERTIES   = -2147220975;
  IMAPI_E_DEVICE_NOTACCESSIBLE  = -2147220974;
  IMAPI_E_DEVICE_NOTPRESENT     = -2147220973;
  IMAPI_E_DEVICE_INVALIDTYPE    = -2147220972;
  IMAPI_E_INITIALIZE_WRITE      = -2147220971;
  IMAPI_E_INITIALIZE_ENDWRITE   = -2147220970;
  IMAPI_E_FILESYSTEM            = -2147220969;
  IMAPI_E_FILEACCESS            = -2147220968;
  IMAPI_E_DISCINFO              = -2147220967;
  IMAPI_E_TRACKNOTOPEN          = -2147220966;
  IMAPI_E_TRACKOPEN             = -2147220965;
  IMAPI_E_DISCFULL              = -2147220964;
  IMAPI_E_BADJOLIETNAME         = -2147220963;
  IMAPI_E_INVALIDIMAGE          = -2147220962;
  IMAPI_E_NOACTIVEFORMAT        = -2147220961;
  IMAPI_E_NOACTIVERECORDER      = -2147220960;
  IMAPI_E_WRONGFORMAT           = -2147220959;
  IMAPI_E_ALREADYOPEN           = -2147220958;
  IMAPI_E_WRONGDISC             = -2147220957;
  IMAPI_E_FILEEXISTS            = -2147220956;
  IMAPI_E_STASHINUSE            = -2147220955;
  IMAPI_E_DEVICE_STILL_IN_USE   = -2147220954;
  IMAPI_E_LOSS_OF_STREAMING     = -2147220953;
  IMAPI_E_COMPRESSEDSTASH       = -2147220952;
  IMAPI_E_ENCRYPTEDSTASH        = -2147220951;
  IMAPI_E_NOTENOUGHDISKFORSTASH = -2147220950;
  IMAPI_E_REMOVABLESTASH        = -2147220949;
  IMAPI_E_CANNOT_WRITE_TO_MEDIA = -2147220948;
  IMAPI_E_TRACK_NOT_BIG_ENOUGH  = -2147220947;
  IMAPI_E_BOOTIMAGE_AND_NONBLANK_DISC = -2147220946;

  SEVERITY_SUCCESS : HResult = 0;
  SEVERITY_ERROR : HResult = 3 shl 30;
  FACILITY_ITF : HResult = 4 shl 16;


{
  Limitations:
   1) Doesn't allow burning to more then one CD drive at a time
X (fixed with implementation of IStorage and IStream) 2) Doesn't allow burning paths which have any component over 32 characters in length (including filename)
}

Type
  // Allows for up to 5 properties at once
(*  TPropSpecArray = array[0..4] of TPropSpec;
  PPropSpecArray = ^TPropSpecArray;

  TPropVariantArray = array[0..4] of TPropVariant;
  PPropVariantArray = ^TPropVariantArray;
*)
  TPropSpecArray = array of TPropSpec;
  PPropSpecArray = ^TPropSpecArray;

  TPropVariantArray = array of TPropVariant;
  PPropVariantArray = ^TPropVariantArray;


  TXPBurn = class;

  {This window is used *only* to synchronize progress messages sent from the
    burner component.  It is created on the same thread as the initializing application
    (the GUI thread).  When the burner component recevies progress from the XP burn
    COM object, it posts a message to this window.  When the message queue of this window
    processes the message, it calls the event on TXPBurn, which, if hooked, calls the
    associated method on whatever application is using the component.

    Roger Harris Nov 2003. TMessageQueue has now been modified to inherit from TObject and contains its
    own window handle
   }

{  TMessageQueue = class
  private
    FOwner: TXPBurn;
    FWindowHandle: HWND;

    procedure DoRecorderChange(var Amsg: TMessage);
    procedure DoAddProgress(var Amsg: TMessage);
    procedure DoBlockProgress(var Amsg: TMessage);
    procedure DoTrackProgress(var Amsg: TMessage);
    procedure DoPreparingBurn(var Amsg: TMessage);
    procedure DoClosingDisc(var Amsg: TMessage);
    procedure DoBurnComplete(var Amsg: TMessage);

    procedure DoEraseComplete(var Amsg: TMessage);
    procedure WndProc(var AMsg: TMessage);

  public
    constructor Create(AOwner:TXPBurn);
    destructor destroy;override ;
    property Burner:TXpBurn read FOwner write FOwner ;
    property Handle:HWND read FWindowHandle;
  end;}


//  TProgressEvents = class;

  TSupportedRecordTypes = (sfNone, sfData, sfMusic, sfBoth);
  TRecordType = (afMusic, afData);
  TEraseKind = (ekQuick, ekFull);
  TRecorderType = (rtCDR, rtCDRW);
  TMedia = RECORD
    iResult:HResult;
    isPresent:boolean ;
    isBlank: boolean;
    isRW: boolean;
    isWritable: boolean;
    isUsable: boolean;
    SessionCount:byte;
    LastTrack:byte;
    StartAddress:Integer;
    NextWriteable:Integer ;
    FreeBlocks:Integer ;
    FreeSpace:Integer ;
    TotalSpace:Integer;
  end;

  TRecorders = class ;

  TRecorder = class(TCollectionItem)
  private
    FVendor: String;
    FRevision: String;
    FProductID: string;
    //FDisplayName: string;
    FRecorderType: TRecorderType;
    FIDiscRecorder: IDiscRecorder;
    FDrivePath: string;
    FMaxWriteSpeed: integer;
    FDosDevice: string;

    function GetRecordTypeAsString: string;
    function GetXpBurner: TXPBurn;
    function GetWriteSpeed: integer;
    function GetMediaInfo: TMedia;
    function GetDriveID: integer;

  protected
    function GetDisplayName: string;override;

  public
    constructor Create(Owner: TRecorders; Vendor,ProductID,Revision: string;RecorderType:TRecorderType);reintroduce;
    procedure Eject ;
    procedure Erase(eraseType: TEraseKind);
//    procedure EraseThread(eraseType: TEraseKind);
    property DriveID:integer  read GetDriveID ;
    property DosDevice:string read FDosDevice ;
    property MediaInfo:TMedia read GetMediaInfo ;
    property XPBurner:TXPBurn read GetXpBurner;
    property Vendor:String read FVendor;
    property ProductID:string read FProductID ;
    property Revision:String read FRevision ;
    property RecorderType:TRecorderType read FRecorderType ;
    property DisplayName:string read GetDisplayName ;
    property RecordTypeAsString:string read GetRecordTypeAsString;
    property DrivePath:string read FDrivePath ;
    property WriteSpeed:integer read GetWriteSpeed ;
    property MaxWriteSpeed:integer read FMaxWriteSpeed ;
    property DiscRecorder:IDiscRecorder read FIDiscRecorder;
  end;

  TDataRecorder = class(TRecorder)
  private
    function GetVolumeName: string;
    procedure SetVolumeName(volume: string);
    function GetDiscSize: Integer;
    function GetFreeDiscSpace: Integer;
    function GetDiscMaster:IJolietDiscMaster ;
   protected

  public
    property VolumeName: string read GetVolumeName write SetVolumeName;
    property DiscSpace: Integer read GetDiscSize;
    property FreeDiscSpace: Integer read GetFreeDiscSpace;
    property DataDiscWriter:IJolietDiscMaster read GetDiscMaster;

  end;


  TRecorderClass = class of TRecorder ;

  TRecorders = class (TCollection)
  private
    FDiscMaster:IUnKnown ;
    FXPBurner: TXPBurn;

  public

    function AddRecorder: TRecorder;
    function AddDiscRecorder(Recorder:IDiscRecorder):TRecorder;
    property XPBurner:TXPBurn read FXPBurner;
  end;

  TDataRecorders = class(TRecorders)
  private
    function getRecorderByPath(DrivePath: String): TDataRecorder;
  protected
    function GetJolietDiscMaster:IJolietDiscMaster ;
    function GetRecorder(Index: integer): TDataRecorder;

  public
    property DataDiscWriter:IJolietDiscMaster read GetJolietDiscMaster;
    property Items[Index:integer]:TDataRecorder read GetRecorder ;
    property ItemsByPath[DrivePath:String]:TDataRecorder read getRecorderByPath;
  end;

{  TMusicRecorders = class(TRecorders)
  private
    function GetRedBookDiscMaster: IRedBookDiscMaster;
  public
    property RedBookDiscMaster:IRedBookDiscMaster read GetRedBookDiscMaster;
  end;}


  TActiveDataRecorder = class
  private
    FSimulate: boolean;
    FEjectAfterBurn: Boolean;
    FRecorder: TDataRecorder;
    FDiscMaster:IDiscMaster;
    FVolumeName: String;
    procedure SetEjectAfterBurn(const Value: Boolean);
    procedure SetRecorder(const Value: TDataRecorder);
    procedure SetSimulate(const Value: boolean);
    function GetFreeSpace: integer;
    function GetTotalSpace: integer;
    function GetMediaInfo: TMedia;
    procedure SetVolumeName(const Value: String);


  public
    constructor Create(DiscMaster:IDiscMaster);
    function    AddData(Stg:IStorage;FileOverWrite:integer):HResult;
    function    SetRecorderAsActive:HResult ;
    function    RecordDisc:HResult ;
    function    OpenExclusive:HResult;
    function    Close:HResult ;
    procedure   QuickErase ;
    procedure   FullErase ;
    procedure   Erase(eraseKind:TEraseKind);
//    procedure   EraseThread(eraseKind:TEraseKind);
    procedure   Eject ;
    procedure   ApplyVolumeName ;
    property    VolumeName:String read FVolumeName write SetVolumeName;
    property    Recorder:TDataRecorder  read FRecorder write SetRecorder;
    property    Simulate:boolean read FSimulate write SetSimulate;
    property    EjectAfterBurn:Boolean  read FEjectAfterBurn write SetEjectAfterBurn;
    property    FreeSpace:integer read GetFreeSpace ;
    property    TotalSpace:integer read GetTotalSpace ;
    property    MediaInfo:TMedia read GetMediaInfo ;


  end;

  TXPBurn = class(TComponent)
  private
    // Class which responds to progress events
//    fProgressEvents: TProgressEvents;
    // Queue which synchronizes progress events with main thread
//    fMessageQueue: TMessageQueue;

    // Event backing stores
    fOnRecorderChange: TNotifyPnPActivity;
    fOnAddProgress: TNotifyCDProgress;
    fOnBlockProgress: TNotifyCDProgress;
    fOnTrackProgress: TNotifyCDProgress;
    fOnPreparingBurn: TNotifyEstimatedTime;
    fOnClosingDisc: TNotifyEstimatedTime;
    fOnBurnComplete: TNotifyCompletionStatus;
    fOnEraseComplete: TNotifyCompletionStatus;
    // End event backing stores

    // Property backing stores
    fLastError: string;
    fCancel: boolean;
    fIsBurning: boolean;
    fActiveFormat: TRecordType;
    fSupportedFormats: TSupportedRecordTypes;
    //fBurnerDrive: Integer;
    fFiles: TStringList;

    fDiscMaster: IDiscMaster;
    fActiveDataRecorder: TActiveDataRecorder;
    //fEjectAfterBurn: boolean;
    //fSimulate: boolean;

    //fProgressCookie: PUINT;
    fFileListOffset: Integer;
    FDataRecorders: TDataRecorders;
//    FMusicRecorders:TMusicRecorders ;
    FIsErasing: boolean;

    // Methods

    function CheckIMAPIError(hr: HResult):boolean;
    function  CreateIStorage:IStorage;
    procedure EnumerateDiscRecorders;
    procedure StreamHelper(storage: IStorage; path, filename: string);
    procedure WriteStream(storage: IStorage; streamName: string);
    procedure SetCancel(const Value: boolean);

    function GetDosDevice(CdDriveID:integer):string ;

    // End methods
  public
    // Constructor
    constructor Create(AOwner: TComponent); override;
    // Destructor
    destructor Destroy; override;
    procedure DiscMasterClose ;
    // Properties
    property DiscMaster:IDiscMaster read FDiscMaster ;
    property DataRecorders:TDataRecorders read FDataRecorders;
    procedure GetRecorderDrives(DriveList: TStringList);
    procedure GetFilesToBurn(FilesToBurnList:TStringList);
    property ActiveDataRecorder:TActiveDataRecorder read FActiveDataRecorder ;
    property Cancel: boolean read fCancel write SetCancel; // Used to cancel the burn if IsBurning is true
    property IsBurning: boolean read fIsBurning; // Checks to see if the drive is currently being burned
    property IsErasing:boolean read FIsErasing;
    property LastError: string read fLastError;

    procedure AddFile(filename, nameOnCD: string); // Adds a file to the file table to be burned to cd with a default open mode of read only + RH: shared
    procedure ClearFiles;
    procedure RemoveFile(filename, nameOnCD: string); // Removes a file from the file table
    procedure RecordDisc(simulate, ejectAfterBurn: boolean);

    // Not impelemented
    { TODO -oAnsonh : Add implementations }
    procedure AddFolder(folderName, folderNameOnCD: string);
    // End not implemented

    // End methods

  published
    // Events
    property OnRecorderChange: TNotifyPnPActivity read fOnRecorderChange write fOnRecorderChange default nil;
    property OnAddProgress: TNotifyCDProgress read fOnAddProgress write fOnAddProgress default nil;
    property OnBlockProgress: TNotifyCDProgress read fOnBlockProgress write fOnBlockProgress default nil;
    property OnTrackProgress: TNotifyCDProgress read fOnTrackProgress write fOnTrackProgress default nil;
    property OnPreparingBurn: TNotifyEstimatedTime read fOnPreparingBurn write fOnPreparingBurn default nil;
    property OnClosingDisc: TNotifyEstimatedTime read fOnClosingDisc write fOnClosingDisc default nil;
    property OnBurnComplete: TNotifyCompletionStatus read fOnBurnComplete write fOnBurnComplete default nil;
    property OnEraseComplete: TNotifyCompletionStatus read fOnEraseComplete write fOnEraseComplete default nil;
    // End Events
  end;

{  TProgressEvents = class(TInterfacedObject, IDiscMasterProgressEvents)
  private
    fOwner: TXPBurn;

    function QueryCancel(out pbCancel: boolean): HResult; stdcall;
    function NotifyPnPActivity: HResult; stdcall;
    function NotifyAddProgress(nCompletedSteps, nTotalSteps: Integer): HResult; stdcall;
    function NotifyBlockProgress(nCompleted, nTotal: Integer): HResult; stdcall;
    function NotifyTrackProgress(nCurrentTrack, nTotalTracks: Integer): HResult; stdcall;
    function NotifyPreparingBurn(nEstimatedSeconds: Integer): HResult; stdcall;
    function NotifyClosingDisc(nEstimatedSeconds: Integer): HResult; stdcall;
    function NotifyBurnComplete(status: HResult): HResult; stdcall;
    function NotifyEraseComplete(status: HResult): HResult; stdcall;

    constructor Create(owner: TXPBurn);
  public
  end;}

  EBurnException = class(Exception);

function CheckError(iResult: HResult):string;
procedure Register;

implementation

const

  MEDIA_NOTPRESENT = $0 ;
  MEDIA_BLANK = $1;
  MEDIA_RW = $2;
  MEDIA_WRITABLE = $4;
  MEDIA_UNUSABLE = $8;

//Type
{  TBurnThread = class(TThread)
  private
    fOwner: TXPBurn;
  protected
    procedure Execute; override;
  public
    constructor Create(Owner: TXPBurn);
  end;}


{  TEraseThread = class(TThread)
  private
    fOwner: TXPBurn;
    FEraseKind:TEraseKind ;
  protected
    procedure Execute; override;
  public
    constructor Create(Owner: TXPBurn;eraseKind:TEraseKind);
  end;}





function GetStgPropString(PS:IPropertyStorage;PropName:String):String;
var
  PSA: TPropSpecArray;
  PVA: TPropVariantArray;
  return:Cardinal ;
begin
  result := '';
  SetLength(PSA,1);
  SetLength(PVA,1);

  PSA[0].ulKind := 0 ;
  try
    PSA[0].lpwstr := CoTaskMemAlloc(sizeof(WideChar) * (length(PropName) + 1));
    StringToWideChar(propName, PSA[0].lpwstr, length(propName) + 1);
    return := Ps.ReadMultiple(1,@PSA[0],@PVA[0])  ;
    if return = 0 then
      result := OleStrToString(PVA[0].bStrVal) ;
   finally
     CoTaskMemFree(PSA[0].lpwstr);
     SetLength(PSA,0);
     SetLength(PVA,0);
   end;
end;

function GetStgPropInteger(PS:IPropertyStorage;PropName:PWideChar):Integer ;
var
  PSA: TPropSpecArray;
  PVA: TPropVariantArray;
  return:Cardinal ;
begin
  result := 0;
  try
  SetLength(PSA,1);
  SetLength(PVA,1);
  PSA[0].ulKind := 0 ;
  PSA[0].lpwstr := PropName ;
  return := Ps.ReadMultiple(1,@PSA[0],@PVA[0])  ;
  if return = 0 then
    result := PVA[0].ulVal ;
  finally
  SetLength(PSA,0);
  SetLength(PVA,0);
  end;

end;







procedure Register;
begin
  RegisterComponents('System', [TXPBurn]);
end;

function IntToBinary(i: Int64): String;
var
  x: integer;
Begin
  for x:=31 downto 0 do
  Begin
    if (i AND (1 shl x)) > 0 then
      Result := Result + '1'
    else
      Result := Result + '0';
  End;
End;

function CheckError(iResult: HResult):string;
begin
  iResult := (iResult);

  case iResult of
    0                             : Result := 'Success!';
    IMAPI_S_PROPERTIESIGNORED     : Result := 'IMAPI_S_PROPERTIESIGNORED';
    IMAPI_S_BUFFER_TO_SMALL       : Result := 'IMAPI_S_BUFFER_TO_SMALL';
    IMAPI_E_NOTOPENED             : Result := 'IMAPI_E_NOTOPENED ';
    IMAPI_E_NOTINITIALIZED        : Result := 'IMAPI_E_NOTINITIALIZED ';
    IMAPI_E_USERABORT             : Result := 'IMAPI_E_USERABORT ';
    IMAPI_E_GENERIC               : Result := 'IMAPI_E_GENERIC ';
    IMAPI_E_MEDIUM_NOTPRESENT     : Result := 'IMAPI_E_MEDIUM_NOTPRESENT ';
    IMAPI_E_MEDIUM_INVALIDTYPE    : Result := 'IMAPI_E_MEDIUM_INVALIDTYPE ';
    IMAPI_E_DEVICE_NOPROPERTIES   : Result := 'IMAPI_E_DEVICE_NOPROPERTIES ';
    IMAPI_E_DEVICE_NOTACCESSIBLE  : Result := 'IMAPI_E_DEVICE_NOTACCESSIBLE ';
    IMAPI_E_DEVICE_NOTPRESENT     : Result := 'IMAPI_E_DEVICE_NOTPRESENT ';
    IMAPI_E_DEVICE_INVALIDTYPE    : Result := 'IMAPI_E_DEVICE_INVALIDTYPE ';
    IMAPI_E_INITIALIZE_WRITE      : Result := 'IMAPI_E_INITIALIZE_WRITE ';
    IMAPI_E_INITIALIZE_ENDWRITE   : Result := 'IMAPI_E_INITIALIZE_ENDWRITE ';
    IMAPI_E_FILESYSTEM            : Result := 'IMAPI_E_FILESYSTEM ';
    IMAPI_E_FILEACCESS            : Result := 'IMAPI_E_FILEACCESS ';
    IMAPI_E_DISCINFO              : Result := 'IMAPI_E_DISCINFO ';
    IMAPI_E_TRACKNOTOPEN          : Result := 'IMAPI_E_TRACKNOTOPEN ';
    IMAPI_E_TRACKOPEN             : Result := 'IMAPI_E_TRACKOPEN ';
    IMAPI_E_DISCFULL              : Result := 'IMAPI_E_DISCFULL ';
    IMAPI_E_BADJOLIETNAME         : Result := 'IMAPI_E_BADJOLIETNAME ';
    IMAPI_E_INVALIDIMAGE          : Result := 'IMAPI_E_INVALIDIMAGE ';
    IMAPI_E_NOACTIVEFORMAT        : Result := 'IMAPI_E_NOACTIVEFORMAT ';
    IMAPI_E_NOActiveRecorder      : Result := 'IMAPI_E_NOActiveRecorder ';
    IMAPI_E_WRONGFORMAT           : Result := 'IMAPI_E_WRONGFORMAT ';
    IMAPI_E_ALREADYOPEN           : Result := 'IMAPI_E_ALREADYOPEN ';
    IMAPI_E_WRONGDISC             : Result := 'IMAPI_E_WRONGDISC ';
    IMAPI_E_FILEEXISTS            : Result := 'IMAPI_E_FILEEXISTS ';
    IMAPI_E_STASHINUSE            : Result := 'IMAPI_E_STASHINUSE ';
    IMAPI_E_DEVICE_STILL_IN_USE   : Result := 'IMAPI_E_DEVICE_STILL_IN_USE ';
    IMAPI_E_LOSS_OF_STREAMING     : Result := 'IMAPI_E_LOSS_OF_STREAMING ';
    IMAPI_E_COMPRESSEDSTASH       : Result := 'IMAPI_E_COMPRESSEDSTASH ';
    IMAPI_E_ENCRYPTEDSTASH        : Result := 'IMAPI_E_ENCRYPTEDSTASH ';
    IMAPI_E_NOTENOUGHDISKFORSTASH : Result := 'IMAPI_E_NOTENOUGHDISKFORSTASH ';
    IMAPI_E_REMOVABLESTASH        : Result := 'IMAPI_E_REMOVABLESTASH ';
    IMAPI_E_CANNOT_WRITE_TO_MEDIA : Result := 'IMAPI_E_CANNOT_WRITE_TO_MEDIA ';
    IMAPI_E_TRACK_NOT_BIG_ENOUGH  : Result := 'IMAPI_E_TRACK_NOT_BIG_ENOUGH ';
    IMAPI_E_BOOTIMAGE_AND_NONBLANK_DISC : Result := 'IMAPI_E_BOOTIMAGE_AND_NONBLANK_DISC ';
  else
    Result:='Unknown: ' + IntToBinary(iResult) + '['+IntToStr(iResult)+']';
   end;
end;


//////////////////////////////////////////////////////////////////////////////////////////////////
//                                    BEGIN TMessageQueue impelmentation
//////////////////////////////////////////////////////////////////////////////////////////////////

{constructor TMessageQueue.Create(AOwner:TXPBurn);
begin
  inherited create;
  FOwner := AOwner;
  FWindowHandle := Classes.AllocateHWnd(WndProc);
end;

destructor TMessageQueue.destroy;
begin
  Classes.DeallocateHWnd(FWindowHandle);
  inherited;
end;

procedure TMessageQueue.WndProc(var AMsg: TMessage);
begin

  with AMsg do
  begin
  if Msg > WM_USER  then
    try
      case Msg of
        WM_RECORDERCHANGE:    DoRecorderChange(Amsg);
        WM_ADDPROGRESS:       DoAddProgress(Amsg);
        WM_BLOCKPROGRESS:     DoBlockProgress(Amsg);
        WM_TRACKPROGRESS:     DoTrackProgress(Amsg);
        WM_PREPARINGBURN:     DoPreparingBurn(Amsg);
        WM_CLOSINGDISC:       DoClosingDisc(Amsg);
        WM_BURNCOMPLETE:      DoBurnComplete(Amsg);
        WM_ERASECOMPLETE:     DoEraseComplete(Amsg);
      end;

    except
      Application.HandleException(Self);
    end
  else
     Result := DefWindowProc(FWindowHandle, Msg, wParam, lParam);
  end
end;


procedure TMessageQueue.DoRecorderChange(var Amsg: TMessage);
begin
  if assigned(FOwner.OnRecorderChange) then
     FOwner.OnRecorderChange;
end;

procedure TMessageQueue.DoAddProgress(var Amsg: TMessage);
begin
  if assigned(FOwner.OnAddProgress) then
    FOwner.OnAddProgress(Amsg.WParam, Amsg.LParam);
end;

procedure TMessageQueue.DoBlockProgress(var Amsg: TMessage);
begin
  if assigned(FOwner.OnBlockProgress) then
    FOwner.OnBlockProgress(Amsg.WParam, Amsg.LParam);
end;

procedure TMessageQueue.DoTrackProgress(var Amsg: TMessage);
begin
  if assigned(FOwner.OnTrackProgress) then
    FOwner.OnTrackProgress(Amsg.WParam, Amsg.LParam);
end;

procedure TMessageQueue.DoPreparingBurn(var Amsg: TMessage);
begin
  if assigned(FOwner.OnPreparingBurn) then
    FOwner.OnPreparingBurn(Amsg.WParam);
end;

procedure TMessageQueue.DoClosingDisc(var Amsg: TMessage);
begin
  if assigned(FOwner.OnClosingDisc) then
    FOwner.OnClosingDisc(Amsg.WParam);
end;

procedure TMessageQueue.DoBurnComplete(var Amsg: TMessage);
begin
  if assigned(FOwner.OnBurnComplete) then
    FOwner.OnBurnComplete(Amsg.WParam);
end;

procedure TMessageQueue.DoEraseComplete(var Amsg: TMessage);
begin
  if assigned(FOwner.OnEraseComplete) then
    FOwner.OnEraseComplete(Amsg.WParam);
end;}



{constructor TProgressEvents.Create(owner: TXPBurn);
begin
  inherited Create;
  fOwner := owner;
end;

function TProgressEvents.QueryCancel(out pbCancel: boolean): HResult;
begin
  pbCancel := fOwner.fCancel;
  Result := S_OK;
end;

function TProgressEvents.NotifyPnPActivity: HResult;
begin
  if Assigned(fOwner.fOnRecorderChange) then
  begin
    fOwner.EnumerateDiscRecorders;
    PostMessage(fOwner.fMessageQueue.Handle, WM_RECORDERCHANGE, 0, 0);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyAddProgress(nCompletedSteps, nTotalSteps: Integer): HResult;
begin
  if Assigned(fOwner.fOnAddProgress) then
  begin
    PostMessage(fOwner.fMessageQueue.Handle, WM_ADDPROGRESS, nCompletedSteps, nTotalSteps);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyBlockProgress(nCompleted, nTotal: Integer): HResult;
begin
  if Assigned(fOwner.fOnBlockProgress) then
  begin
    PostMessage(fOwner.fMessageQueue.Handle, WM_BLOCKPROGRESS, nCompleted, nTotal);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyTrackProgress(nCurrentTrack, nTotalTracks: Integer): HResult;
begin
  if Assigned(fOwner.fOnTrackProgress) then
  begin
    PostMessage(fOwner.fMessageQueue.Handle, WM_TRACKPROGRESS, nCurrentTrack, nTotalTracks);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyPreparingBurn(nEstimatedSeconds: Integer): HResult;
begin
  if Assigned(fOwner.fOnPreparingBurn) then
  begin
    PostMessage(fOwner.fMessageQueue.Handle, WM_PREPARINGBURN, nEstimatedSeconds, 0);
    Result := S_OK;
  end
  else
    Result :=  E_NOTIMPL;
end;

function TProgressEvents.NotifyClosingDisc(nEstimatedSeconds: Integer): HResult;
begin
  if Assigned(fOwner.fOnClosingDisc) then
  begin
    PostMessage(fOwner.fMessageQueue.Handle, WM_CLOSINGDISC, nEstimatedSeconds, 0);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyBurnComplete(status: HResult): HResult;
begin
  if Assigned(fOwner.fOnBurnComplete) then
  begin
    fOwner.fIsBurning := false;
    fOwner.fLastError := CheckError(status);
    PostMessage(fOwner.fMessageQueue.Handle, WM_BURNCOMPLETE, status, 0);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;

function TProgressEvents.NotifyEraseComplete(status: HResult): HResult;
begin
  if Assigned(fOwner.fOnEraseComplete) then
  begin
    fOwner.fLastError := CheckError(status);
    PostMessage(fOwner.fMessageQueue.Handle, WM_ERASECOMPLETE, status, 0);
    Result := S_OK;
  end
  else
    Result := E_NOTIMPL;
end;
}


{constructor TBurnThread.Create(Owner: TXPBurn);
begin
  freeOnTerminate := true;
  fOwner := Owner;
  inherited Create(false);
end;


procedure TBurnThread.Execute;
var AStorage:IStorage ;
var iError:HResult ;
begin
  CoInitializeEx(nil, COINIT_MULTITHREADED);

  try

    // dont care if there is an error here. Just want to ensure recorder is closed.
    fOwner.ActiveDataRecorder.Close ;

    iError := fOwner.ActiveDataRecorder.SetRecorderAsActive ;
    if iError = S_OK then
      begin
        AStorage := fOwner.CreateIStorage;
        FOwner.ActiveDataRecorder.ApplyVolumeName ;
        iError := fOwner.ActiveDataRecorder.AddData(AStorage, 1);

        if iError = S_OK then
          fOwner.ActiveDataRecorder.RecordDisc;
          // RecordDisc calls NotifyBurnComplete
      end;

  finally
    fOwner.fIsBurning := false;
    CoUninitialize;
    if iError <> S_OK then
      fOwner.fProgressEvents.NotifyBurnComplete(iError);

  end;
end;
}

                 {*************************}
                 {*         TXPBurn       *}
                 {*************************}
constructor TXPBurn.Create(AOwner: TComponent);
var
   pUnknown           :IUnknown;
   pEnumDiscFormats   :IEnumDiscMasterFormats;
   lpguidFormatID     :TGUID;
   pcFetched          :ULONG;
begin
     inherited Create (AOwner);

     if (csDesigning in ComponentState) then
        exit;

     fCancel    := false;
     fIsBurning := false;
     fLastError := '';

     //  fMessageQueue := TMessageQueue.Create(self);
     //  fProgressEvents := TProgressEvents.Create(self);
     // Liste des fichiers à graver
     fFiles                   := TStringList.Create;
//     FDataRecorders           := TDataRecorders.Create (TRecorder);

//  FMusicRecorders := TMusicRecorders.Create(TRecorder) ;
//  FMusicRecorders.FXPBurner := self ;

  try
      pUnknown                 := CreateComObject (CLSID_MSDiscMasterObj);
      fDiscMaster              := pUnknown as IDiscMaster;
      FDataRecorders           := TDataRecorders.Create (TRecorder);
      FDataRecorders.FXPBurner := Self;
      FActiveDataRecorder      := TActiveDataRecorder.Create (FDiscMaster);
  except
        on e:Exception do
        begin
             // Raise an exception here, because this component cannot continue in these circumstances. Containing application should wrap create in try catch
             fLastError := 'IMAPI non disponible: ' + e.Message;
             raise EBurnException.Create (fLastError);
        end;
  end;

  CheckIMAPIError (fDiscMaster.Open);
//  CheckIMAPIError (fDiscMaster.ProgressAdvise(fProgressEvents, fProgressCookie));
  CheckIMAPIError (fDiscMaster.EnumDiscMasterFormats(pEnumDiscFormats));

  while((SUCCEEDED(pEnumDiscFormats.Next(1, lpguidFormatID, pcFetched))) and (pcFetched = 1)) do
  begin
       if (IsEqualGuid(lpguidFormatID, IID_IJolietDiscMaster)) then
          fSupportedFormats := TSupportedRecordTypes(Integer(fSupportedFormats) or 1);
       if (IsEqualGuid(lpguidFormatID, IID_IRedBookDiscMaster)) then
          fSupportedFormats := TSupportedRecordTypes(Integer(fSupportedFormats) or 2);
  end;

  if (fSupportedFormats = sfNone) then
  begin
       DiscMasterClose;
       raise EBurnException.Create ('This API does not support the formats understood by this component');
  end;
  fActiveFormat := afData;
  EnumerateDiscRecorders;

  if FDataRecorders.Count > 0 then
     FActiveDataRecorder.Recorder := FDataRecorders.Items[0];
end;

destructor TXPBurn.Destroy;
begin
     if (not (csDesigning in ComponentState)) then
     begin
          //    fMessageQueue.Free;
          FreeAndNil (FActiveDataRecorder);
          //FreeAndNil (FDataRecorders); // $$$ JP 27/03/06 - fuite mémoire, mais tant pis: plantage dans certains cas, hélas donc on libère pas!
          if (Assigned(FDiscMaster)) then
             DiscMasterClose;
          FreeAndNil (fFiles);
     end;

     inherited Destroy;
end;

procedure TXPBurn.DiscMasterClose;
begin
//     fDiscMaster.ProgressUnadvise (fProgressCookie);
     fDiscMaster.Close;
end;

function TXPBurn.CheckIMAPIError(hr: HResult):boolean;
begin
     Result := TRUE;
     if (not Succeeded(hr)) then
     begin
          //$$$ JP 06/12/05 - warning delphi -> Result := FALSE;
          fLastError := CheckError(hr);
          raise EBurnException.Create(fLastError);
     end;
end;

// Returns the paths to the recorder drives for the current active format (list returned should be deleted by the caller)
// Precondition: Requires that fMusicRecorderDrives and fDataRecorderDives have been created
// Postcondition: Returns a new StringList instance, this must be deleted by the caller
procedure TXPBurn.GetRecorderDrives(DriveList: TStringList);
var
  count: Integer;
  recorderDrives: TRecorders;
begin
  recorderDrives := nil;

  case fActiveFormat of
//    afMusic: recorderDrives := fMusicRecorders;
    afData: recorderDrives :=  fDataRecorders;
  end;
  if RecorderDrives = Nil then exit;

  DriveList.BeginUpdate;

  for count := 0 to recorderDrives.Count - 1 do
    DriveList.Add(recorderDrives.Items[count].DisplayName);
  DriveList.EndUpdate;
end;

// Returns the list of files (not including folders) currently set to be burned to the disc (list returned should be deleted by the caller)
// Precondition: fFiles has been created
// Postcondition: Returns a new StringList instance, this must be deleted by the caller (values contains the path on the CD to burn the file to)
procedure TXPBurn.GetFilesToBurn(FilesToBurnList: TStringList);
var
  count: Integer;
begin

  FilesToBurnList.BeginUpdate;
  for count := 0 to fFiles.Count - 1 do
  begin
    FilesToBurnList.Add(fFiles.Values[fFiles.Strings[count]]);
    FilesToBurnList.Values[fFiles.Names[count]] := fFiles.Names[count];
  end;
  FilesToBurnList.EndUpdate;
end;

// This function will 'add files' onto a queue.  The second parameter indicates whether the file will be named differently on disc.
// Notice that the nameOnCD parameter may include drive information, or may not, but either way, it must include a filename
// Precondition: File must exist on disc
// Postcondition: File table is updated with information (note that this information isn't fully resolved until burn is called)
procedure TXPBurn.AddFile(filename, nameOnCD: string);
var
  cdName: string;
begin
  if ((fileName = '') or (nameOnCD = '') or (ExtractFilePath(nameOnCD) = '')) then
    raise EBurnException.Create('Neither argument to AddFile may be empty strings');

  if (FileExists(filename)) then
  begin
    cdName := ExtractRelativePath(ExtractFileDrive(nameOnCD), nameOnCD);
    fFiles.Add(cdName + '=' + filename );
  end
  else
    raise EBurnException.Create('File ' + filename + ' must exist to be burned to a CD');
end;

// This function removes files from the list to be added to the CD.  The filename passed in should be identical to the filename
// that was passed to the AddFile method.  This method will fail if filename is nil, or if the file isn't located in the list.
// Precondition: Filename must not be nil, and must exist in the fFiles list
// Postcondition: File is removed from the file table
procedure TXPBurn.RemoveFile(filename, nameOnCD: string);
var
  index: Integer;
  findString: string;
begin
  if (filename <> '') then
  begin
    findString := filename + '=' + nameOnCD;
    index := fFiles.IndexOf(findString);
    if (index <> -1) then
    begin
      fFiles.Delete(index);
    end;
  end;
end;

{ TODO -oansonh : Add implementation }
procedure TXPBurn.AddFolder (FolderName, FolderNameOnCD: string);
var
   sr    :TSearchRec;
begin
     if FindFirst (FolderName + '\*.*', faAnyFile, sr) = 0 then
     begin
          repeat
                if (sr.Attr and faDirectory) = sr.Attr then
                begin
                     if (sr.Name <> '.') and (sr.Name <> '..') then
                        AddFolder (FolderName + '\' + sr.Name, FolderNameOnCD + '\' + sr.Name)
                end
                else
                    AddFile (FolderName + '\' + sr.Name, FolderNameOnCD + '\' + sr.Name);
          until FindNext (sr) <> 0;
          FindClose (sr);
     end;
end;



procedure TXPBurn.RecordDisc (simulate, ejectAfterBurn:boolean);
var
   AStorage    :IStorage;
   iError      :HResult;
begin
     fIsBurning := true;
     iError:=S_OK;

     try
        ActiveDataRecorder.EjectAfterBurn := ejectAfterBurn;
        ActiveDataRecorder.Simulate       := simulate;
        ActiveDataRecorder.Close;
        iError := ActiveDataRecorder.SetRecorderAsActive;
        if iError = S_OK then
        begin
             AStorage := CreateIStorage;
             ActiveDataRecorder.ApplyVolumeName;
             iError := ActiveDataRecorder.AddData (AStorage, 1);
             if iError = S_OK then
                ActiveDataRecorder.RecordDisc;
        end;
     finally
            fIsBurning := false;
            if iError <> S_OK then
               raise EBurnException.Create (CheckError (iError));
     end;
end;

procedure TXPBurn.WriteStream(storage: IStorage; streamName: string);
var
  fileToWrite: string;
begin
  fileToWrite := fFiles.Values[fFiles.Names[fFileListOffset]];

  if (FileExists(fileToWrite)) then
  begin
    streamName := ExtractFileName(streamName);
    (storage as IStorage_EX).CreateFileStream(fileToWrite, streamName,word(FFiles.Objects[FFileListOffSet]));
  end;
end;

procedure TXPBurn.StreamHelper(storage: IStorage; path, filename: string);
var
  nestedFilename, subStorageName: string;
  index: Integer;
  newStorage: IStorage;
begin
  index := Pos ('\', filename);
  if (index <> 0) then
  begin
    //  Create sub-storage
    nestedFilename := Copy(filename, index + 1, Length(filename) - index);
    subStorageName := Copy(filename, 1, index - 1);
    // Create the storage here (subStorageName contains the storage name)

    newStorage := (Storage as IStorage_EX).CreateStorageDirectory(subStorageName);

    StreamHelper(newStorage, path + subStorageName + '\', nestedFilename);
  end
  else
  begin
    // Write Stream - name is filename
    WriteStream((storage as IStorage), filename);
    Inc(fFileListOffset);
  end;
  if (fFileListOffset < fFiles.Count) then
  begin
    if (path <> '') then
    begin
      nestedFilename := fFiles.Names[fFileListOffset];
      index := Pos(path, nestedFilename);
      if (index <> 0) then
      begin
        StreamHelper(storage, path, Copy(nestedFilename, index + Length(path), Length(nestedFilename) - (index + Length(path)) + 1));
      end;
    end;
  end;
end;

function  TXPBurn.CreateIStorage:IStorage;
Var RStorage:IStorage ;
begin
  RStorage := TStorage.Create('RootStorage');

  // First sort fFiles (prepare to create directory strucutre in compound file)
  fFiles.Sort;
  fFileListOffset := 0;

  try
    while(fFileListOffset < fFiles.Count) do
      StreamHelper (RStorage, '', fFiles.Names[fFileListOffset]);
  finally
    Result := RStorage ;
  end;
end;

procedure TXPBurn.EnumerateDiscRecorders;
var
//  i:integer ;
  pUnknown: IUnknown;
  lpguidFormatID: TGUID;
  ppRecorder: IDiscRecorder;
  pEnumDiscRecorders: IEnumDiscRecorders;
  pcFetched: ULONG;
  DR:TRecorder ;
begin
  FDataRecorders.Clear ;
  if ((fSupportedFormats = sfBoth) or (fSupportedFormats = sfData)) then
  begin
    lpguidFormatID := IID_IJolietDiscMaster;
    CheckIMAPIError(fDiscMaster.SetActiveDiscMasterFormat(lpguidFormatID, pUnknown));
    FDataRecorders.FDiscMaster := pUnknown as IJolietDiscMaster ;
    CheckIMAPIError(fDiscMaster.EnumDiscRecorders(pEnumDiscRecorders));

    while((SUCCEEDED(pEnumDiscRecorders.Next(1, ppRecorder, pcFetched))) and (pcFetched = 1))do
    begin
      DR :=  FDataRecorders.AddDiscRecorder(ppRecorder)   ;
      DR.FDosDevice := GetDosDevice(Dr.DriveID);

    end;
  end;


  if (fDataRecorders.Count = 0) then
    raise EBurnException.Create('No XP compatible CDR (which supports afData) on this system');
end;


procedure TXPBurn.SetCancel(const Value: boolean);
begin
  if Value then
    begin
      if (FIsBurning and Value) then
        fCancel := Value
      else
        fCancel := false;
    end;
end;


procedure TXPBurn.ClearFiles;
begin
   fFiles.Clear ;
end;

function TXPBurn.GetDosDevice(CDDriveID:integer):string;
var
  Buffer : array[0..500] of char;
  TmpPC  : PChar;
  CdCtr:integer ;
begin
  CdCtr := 0 ;
  result := '';
  GetLogicalDriveStrings(SizeOf(Buffer),Buffer);
  TmpPC := Buffer;
   while TmpPC[0] <> #0 do begin
    if Windows.GetDriveType(TmpPC) = DRIVE_CDROM then
        begin
         if CdCtr = CdDriveId then
           begin
             result := TmpPC;
             exit;
           end;
           CdCtr := CdCtr+1;
        end ;
    TmpPC := StrEnd(TmpPC)+1;
   end;
end;

{ TRecorder }

constructor TRecorder.Create(Owner: TRecorders; Vendor,
  ProductID, Revision: string; RecorderType: TRecorderType);
begin
  inherited Create(Owner);
  FVendor := Vendor;
  FProductID := ProductID;
  FRevision  := Revision;
  FRecorderType := RecorderType;
end;

procedure TRecorder.Eject;
begin
   XpBurner.CheckIMAPIError(DiscRecorder.OpenExclusive);
  try
    XPBurner.CheckIMAPIError(DiscRecorder.Eject);
  finally
    XPBurner.CheckIMAPIError(DiscRecorder.Close);
  end;

end;


function TRecorder.GetMediaInfo: TMedia;
var
  mediaType, mediaFlags: Integer;
  FreeBlocks:ULONG ;
  Sessions:byte;
  LastTrack:byte ;
  StartAddress:ULONG;
  NextWriteable:ULONG;
//  iResult:HResult;
begin
  Result.iResult := DiscRecorder.OpenExclusive ;

  try
  if Succeeded(Result.iResult) then
  begin

    Result.isPresent := false ;
    Result.isBlank := false;
    Result.isRW := false;
    Result.isWritable := false;
    Result.isUsable := false;

    Result.iResult := DiscRecorder.QueryMediaType(mediaType, mediaFlags);
    if succeeded(Result.iResult) then
    begin

      Result.isPresent := not (mediaFlags = MEDIA_NOTPRESENT) and  not(mediaType = MEDIA_NOTPRESENT) ;
      Result.isBlank := ((mediaFlags and MEDIA_BLANK) = MEDIA_BLANK);
      Result.isRW := ((mediaFlags and MEDIA_RW) = MEDIA_RW);
      Result.isWritable := ((mediaFlags and MEDIA_WRITABLE) = MEDIA_WRITABLE);
      Result.isUsable := not ((mediaFlags and MEDIA_UNUSABLE) = MEDIA_UNUSABLE);
    end ;
    result.SessionCount := 0;
    result.LastTrack := 0;
    result.StartAddress := 0;
    result.NextWriteable := 0;
    result.FreeBlocks := 0;
    result.FreeSpace :=   0;
    result.TotalSpace := 0;

    if (succeeded(result.iResult) and result.isPresent)  then
    begin
      Result.iResult := DiscRecorder.QueryMediaInfo(sessions,lasttrack,startaddress,nextwriteable,FreeBlocks);
      if succeeded(Result.iResult) then
      begin
        result.SessionCount := Sessions;
        result.LastTrack := LastTrack;
        result.StartAddress := StartAddress;
        result.NextWriteable := NextWriteable;
        result.FreeBlocks := FreeBlocks;
        result.FreeSpace :=   FreeBlocks*2048 ;
        result.TotalSpace := (FreeBlocks+NextWriteable)*2048;
      end ;
    end;
  end;
  finally
    XPBurner.CheckIMAPIError(DiscRecorder.Close);
  end;


end;



{procedure TRecorder.EraseThread(eraseType: TEraseKind);
var
  bFullErase: Boolean;
  nError:HResult ;
begin

  if ((RecorderType = rtCDRW) and (GetMediaInfo.isRW)) then
  begin
    if GetMediaInfo.isBlank then
     begin
       nError := XpBurner.fProgressEvents.NotifyEraseComplete(WM_MEDIABLANK);
       exit ;
     end;

    try
     TEraseThread.Create(XpBurner,EraseType);
    finally
      XpBurner.fIsErasing := false ;
    end;
  end
  else
    raise EBurnException.Create('Recorder type and Media type must be CDRW to erase disc');
end;
}

procedure TRecorder.Erase(eraseType: TEraseKind);
var
  bFullErase: Boolean;
  nError:HResult ;
begin

  if ((RecorderType = rtCDRW) and (GetMediaInfo.isRW)) then
  begin
{    if GetMediaInfo.isBlank then
     begin
       nError := XpBurner.fProgressEvents.NotifyEraseComplete(WM_MEDIABLANK);
       exit ;
     end;}

    try
      nError := DiscRecorder.OpenExclusive ;
      if nError = 0 then
      begin
        bFullErase := eraseType= ekFull ;
        XpBurner.fIsErasing := true ;

        DiscRecorder.erase(bFullErase);
      end;
      // RH: it seems notify erase complete is not called by IDiscRecorder ;

    finally
      //$$$ JP 06/12/05 - warning delphi -> nError := DiscRecorder.Close;
      XPBurner.FIsErasing := false ;
//      XPBurner.fProgressEvents.NotifyEraseComplete(nError);
    end;
  end
  else
    raise EBurnException.Create('Recorder type and Media type must be CDRW to erase disc');
end;





function TRecorder.GetDisplayName: string;
begin
  result := FVendor+' '+FProductID ;
end;


function TRecorder.GetRecordTypeAsString: string;
begin
  result := '';
  case FRecorderType of
    rtCDR:result := 'CD-R';
    rtCDRW:result := 'CD-RW';
  end;
end;

function TRecorder.GetXpBurner: TXPBurn;
begin
   result := (Collection As TRecorders).XPBurner ;
end;



function TRecorder.GetWriteSpeed: integer;
Var
  ppPropStg: IPropertyStorage;
begin
  //$$$ JP 06/12/05 - warning delphi -> Result := 0;
  XPBurner.CheckIMAPIError(DiscRecorder.GetRecorderProperties(ppPropStg));
  result := GetStgPropInteger (ppPropStg, 'WriteSpeed');
end;


function TRecorder.GetDriveID: integer;
var last:char ;
begin
  result:=0;
  last := FdrivePath[length(FdrivePath)];
  if last in ['0','1','2','3'] then
     result := strtoint(last) ;
end;

{ TRecorders }

function TRecorders.AddDiscRecorder( Recorder: IDiscRecorder): TRecorder;
var
  vendor, productID, revision:TBSTR ;

  pbStrPath: TBSTR;
  typeCode: Integer;
  ppPropStg: IPropertyStorage;
begin
  result := AddRecorder ;
  result.FIDiscRecorder := Recorder ;
  try
    // must be assigned Nil prior to call
    vendor := nil;
    productID := nil;
    revision := nil;
    pbStrPath := Nil ;
    XPBurner.CheckIMAPIError(Recorder.GetDisplayNames(vendor, productID, revision));

    // Copy data into pascal strings
    result.fVendor := OleStrToString(vendor);
    result.fProductID := OleStrToString(productID);
    result.fRevision := OleStrToString(revision);


    XPBurner.CheckIMAPIError(Recorder.GetPath(pbstrPath));
    result.FDrivePath := OleStrToString(pbstrPath);


    XPBurner.CheckIMAPIError(Recorder.GetRecorderType(typeCode));

    case typeCode of
      $1: result.fRecorderType := rtCDR;
      $2: result.fRecorderType := rtCDRW;
    end;

    // get write speed
    XPBurner.CheckIMAPIError(Recorder.GetRecorderProperties(ppPropStg));
    result.FMaxWriteSpeed := GetStgPropInteger(ppPropStg,'MaxWriteSpeed') ;
  finally
    SysFreeString(vendor);
    SysFreeString(productID);
    SysFreeString(revision);
    SysFreeString(pbstrPath);
  end;

end;

function TRecorders.AddRecorder: TRecorder;
begin
  result := TRecorder(inherited add) ;
end;

{ TDataRecorders }

function TDataRecorders.GetJolietDiscMaster: IJolietDiscMaster;
begin
  result := FDiscMaster as IJolietDiscMaster ;
end;

function TDataRecorders.GetRecorder(
  Index: integer): TDataRecorder;
begin
  Result := TDataRecorder(inherited Items[Index]);
end;

function TDataRecorders.getRecorderByPath(
  DrivePath: String): TDataRecorder;
var i:integer ;
begin
  result := nil;
  for i := 0 to count-1 do
    if items[i].DrivePath = DrivePath then
      begin
        result := items[i] ;
        exit;
      end;
end;

{ TActiveDataRecorder }

function  TActiveDataRecorder.AddData(Stg: IStorage; FileOverWrite: integer):HResult;
begin
  result:= Recorder.DataDiscWriter.AddData(Stg,FileOverWrite) ;
end;

procedure TActiveDataRecorder.ApplyVolumeName;
begin
  if VolumeName <> '' then
    Recorder.VolumeName := VolumeName ;
end;

function TActiveDataRecorder.Close: HResult;
begin
  result := Recorder.DiscRecorder.Close ;
end;

constructor TActiveDataRecorder.Create(DiscMaster: IDiscMaster);
begin
  inherited Create;
  FSimulate := false ;
  FEjectAfterBurn := false;
  FDiscMaster := DiscMaster ;
end;

procedure TActiveDataRecorder.Eject;
begin
  Recorder.Eject ;
end;

procedure TActiveDataRecorder.Erase(eraseKind: TEraseKind);
begin
  recorder.Erase(eraseKind);
end;

{procedure TActiveDataRecorder.EraseThread(eraseKind: TEraseKind);
begin
  recorder.EraseThread(eraseKind);
end;}

procedure TActiveDataRecorder.FullErase;
begin
//  EraseThread(ekfull);
end;

function TActiveDataRecorder.GetFreeSpace: integer;
begin
  result := recorder.FreeDiscSpace ;
end;

function TActiveDataRecorder.GetMediaInfo: TMedia;
begin
  result := recorder.MediaInfo ;
end;

function TActiveDataRecorder.GetTotalSpace: integer;
begin
  result := Recorder.DiscSpace ;
end;

function TActiveDataRecorder.OpenExclusive: HResult;
begin
   result := Recorder.DiscRecorder.OpenExclusive ;
end;

procedure TActiveDataRecorder.QuickErase;
begin
//  EraseThread(ekQuick);
end;

function TActiveDataRecorder.RecordDisc:HResult;
begin
//  Recorder.VolumeName := 'TestOne' ;
  result := FDiscMaster.RecordDisc(FSimulate,FEjectAfterBurn) ;
end;

procedure TActiveDataRecorder.SetEjectAfterBurn(const Value: Boolean);
begin
  FEjectAfterBurn := Value;
end;

procedure TActiveDataRecorder.SetRecorder(const Value: TDataRecorder);
begin
  FRecorder := Value;
end;

function TActiveDataRecorder.SetRecorderAsActive:HResult;
begin
  result := FDiscMaster.SetActiveDiscRecorder(Recorder.DiscRecorder) ;
end;

procedure TActiveDataRecorder.SetSimulate(const Value: boolean);
begin
  FSimulate := Value;
end;

procedure TActiveDataRecorder.SetVolumeName(const Value: String);
begin
     FVolumeName := Value;
end;


{function TMusicRecorders.GetRedBookDiscMaster: IRedBookDiscMaster;
begin
  result := FDiscMaster as IRedBookDiscMaster ;
end;}


function TDataRecorder.GetVolumeName: string;
var MaximumComponentLength : dWord;
    FileSystemFlags : dWord;
    VolumeName : String;
begin
 SetLength(VolumeName, 64);


 GetVolumeInformation(PChar(DosDevice),
                      PChar(VolumeName),
                      Length(VolumeName),
                      nil,
                      MaximumComponentLength,
                      FileSystemFlags,
                      nil,
                      0);
 Result := Copy(VolumeName, 0, StrLen(PChar(VolumeName)));
end;




procedure TDataRecorder.SetVolumeName(volume: string);

var

  ppPropStg: IPropertyStorage;
  rgPropID: PPropSpec;
  rgPropVar: PPropVariant;
  propertyID: string;

begin
    propertyID := 'VolumeName';
    DataDiscWriter.GetJolietProperties(ppPropStg);
    rgPropID := GetMemory(sizeof(tagPROPSPEC) * 1);
    rgPropVar := GetMemory(sizeof(tagPROPVARIANT) * 1);
    rgPropID^.ulKind := 0;
    rgPropID^.lpwstr := CoTaskMemAlloc(sizeof(WideChar) * (length(propertyID) + 1));
    StringToWideChar(propertyID, rgPropID^.lpwstr, length(propertyID) + 1);
    try
      if (SUCCEEDED(ppPropStg.ReadMultiple(1, rgPropID, rgPropVar))) then
      begin
        rgPropVar^.vt := VT_BSTR;
        rgPropVar^.bstrVal := StringToOleStr(volume);
        if (SUCCEEDED(ppPropStg.WriteMultiple(1, rgPropID, rgPropVar, 2))) then

          XPBurner.CheckIMAPIError(DataDiscWriter.SetJolietProperties(ppPropStg))

        else

        begin

          raise EBurnException.Create('Write multiple failed');

        end;

      end

      else

      begin


        raise EBurnException.Create('Read multiple failed');

      end;

  finally

      CoTaskMemFree(rgPropID^.lpwstr);
      SysFreeString(rgPropVar^.bstrVal);
      FreeMem(rgPropID);
      FreeMem(rgPropVar);

    end;

 end ;





(*

procedure TDataRecorder.SetVolumeName(volume:String) ;
var
  PS: IPropertyStorage;
  PSA: TPropSpecArray;
  PVA: TPropVariantArray;
begin
  try

  if not succeeded(DataDiscWriter.GetJolietProperties(ps)) then
        raise EBurnException.Create('Get Joliet Properties failed');


  SetLength(PSA,1 );
  SetLength(PVA, 1);
  PSA[0].ulKind := PRSPEC_LPWSTR ; // 0
  PSA[0].lpwstr := 'VolumeName' ;

  if not succeeded(Ps.ReadMultiple(1,@PSA[0],@PVA[0])) then
    raise EBurnException.Create('Read multiple failed');

  PVA[0].vt := VT_LPWSTR; // or VT_BSTR ;
  PVA[0].pwszVal  := StringToOleStr(Volume) ;
  if not succeeded( PS.WriteMultiple(1, @PSA[0], @PVA[0], PID_FIRST_USABLE)) then
    raise EBurnException.Create('Write multiple failed');

  if not succeeded(DataDiscWriter.SetJolietProperties(ps)) then
    raise EBurnException.Create('Set Joliet Properties failed');




  finally
    SysFreeString(PVA[0].pwszval);
    SetLength(PSA, 0);
    SetLength(PVA, 0);
  end;

end;
*)
function TDataRecorder.GetFreeDiscSpace: Integer;
begin
  Result := MediaInfo.FreeSpace;
end;



function TDataRecorder.GetDiscSize: Integer;
begin
  Result := MediaInfo.TotalSpace ;
end;


function TDataRecorder.GetDiscMaster: IJolietDiscMaster;
begin
   result := (Collection as TDataRecorders).FDiscMaster as IJolietDiscMaster ;
end;


{constructor TEraseThread.Create(Owner: TXPBurn;eraseKind:TEraseKind);
begin
  freeOnTerminate := true;
  fOwner := Owner;
  fEraseKind := eraseKind ;
  inherited Create(false);
end;

procedure TEraseThread.Execute;
var nError:HResult;
    bFullErase:boolean ;
begin
  CoInitializeEx(nil, COINIT_MULTITHREADED);
    try
      nError := FOwner.ActiveDataRecorder.OpenExclusive ;
      if nError = 0 then
      begin
        bFullErase := fEraseKind = ekFull ;
        fOwner.fIsErasing := true ;
        fOwner.ActiveDataRecorder.Recorder.DiscRecorder.erase(bFullErase);
      end;
      // RH: it seems notify erase complete is not called by IDiscRecorder ;

    finally
      nError := FOwner.ActiveDataRecorder.Close;
      fOwner.fIsErasing := false ;
//      fOwner.fProgressEvents.NotifyEraseComplete(nError);
      CoUninitialize;
    end;
end;}


initialization
              CoInitializeEx (nil, COINIT_MULTITHREADED);

finalization
              CoUninitialize;

end.


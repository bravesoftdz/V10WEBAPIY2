{ *****************************************************************************
  WARNING: This component file was generated using the EventSinkImp utility.
           The contents of this file will be overwritten everytime EventSinkImp
           is asked to regenerate this sink component.

  NOTE:    When using this component at the same time with the XXX_TLB.pas in
           your Delphi projects, make sure you always put the XXX_TLB unit name
           AFTER this component unit name in the USES clause of the interface
           section of your unit; otherwise you may get interface conflict
           errors from the Delphi compiler.

           EventSinkImp is written by Binh Ly (bly@techvanguards.com)
  *****************************************************************************
  //Sink Classes//
  TEdisys_IULM_CoreIDetectionEvents2
}

{$IFDEF VER100}
{$DEFINE D3}
{$ENDIF}

//SinkUnitName//
unit Edisys_IULM_CoreEvents;

interface

uses
  Windows, ActiveX, Classes, ComObj, OleCtrls
  , Edisys_IULM_Core_TLB
  ;

type
  { backward compatibility types }
  {$IFDEF D3}
  OLE_COLOR = TOleColor;
  {$ENDIF}

  TEdisys_IULM_CoreEventsBaseSink = class (TComponent, IUnknown, IDispatch)
  protected
    { IUnknown }
    function QueryInterface(const IID: TGUID; out Obj): HResult; {$IFNDEF D3} override; {$ENDIF} stdcall;
    function _AddRef: Integer; stdcall;
    function _Release: Integer; stdcall;

    { IDispatch }
    function GetIDsOfNames(const IID: TGUID; Names: Pointer;
      NameCount, LocaleID: Integer; DispIDs: Pointer): HResult; virtual; stdcall;
    function GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult; virtual; stdcall;
    function GetTypeInfoCount(out Count: Integer): HResult; virtual; stdcall;
    function Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; stdcall;
  protected
    FCookie: integer;
    FCP: IConnectionPoint;
    FSinkIID: TGUID;
    FSource: IUnknown;
    function DoInvoke (DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; virtual; abstract;
  public
    destructor Destroy; override;
    procedure Connect (const ASource: IUnknown);
    procedure Disconnect;
    property SinkIID: TGUID read FSinkIID write FSinkIID;
    property Source: IUnknown read FSource;
  end;

  TIDetectionEvents2EngineReadyEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2AccountDataRequiredEvent = procedure (Sender: TObject; AccountType: AccountType; var accountKey: WideString; const AccountData: IAccountData) of object;
  TIDetectionEvents2AccountStateChangedEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2UserRequestNewDealsEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2UserRequestSelectedDealsEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2UserRequestUpdatedDealsEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2UserRequestDealsWithAvailableDIEEvent = procedure (Sender: TObject) of object;
  TIDetectionEvents2UserValidationRequiredEvent = procedure (Sender: TObject; var password: WideString; var isValidated: WordBool) of object;

  //SinkComponent//
  TEdisys_IULM_CoreIDetectionEvents2 = class (TEdisys_IULM_CoreEventsBaseSink )
  protected
    function DoInvoke (DispID: Integer; const IID: TGUID; LocaleID: Integer;
      Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
      VarResult, ExcepInfo, ArgErr: Pointer): HResult; override;

    //ISinkInterfaceMethods//
  public
    { system methods }
    constructor Create (AOwner: TComponent); override;
  protected
    //SinkInterface//
    procedure DoEngineReady; safecall;
    procedure DoAccountDataRequired(AccountType: AccountType; var accountKey: WideString; const AccountData: IAccountData); safecall;
    procedure DoAccountStateChanged; safecall;
    procedure DoUserRequestNewDeals; safecall;
    procedure DoUserRequestSelectedDeals; safecall;
    procedure DoUserRequestUpdatedDeals; safecall;
    procedure DoUserRequestDealsWithAvailableDIE; safecall;
    procedure DoUserValidationRequired(var password: WideString; var isValidated: WordBool); safecall;
  protected
    //SinkEventsProtected//
    FEngineReady: TIDetectionEvents2EngineReadyEvent;
    FAccountDataRequired: TIDetectionEvents2AccountDataRequiredEvent;
    FAccountStateChanged: TIDetectionEvents2AccountStateChangedEvent;
    FUserRequestNewDeals: TIDetectionEvents2UserRequestNewDealsEvent;
    FUserRequestSelectedDeals: TIDetectionEvents2UserRequestSelectedDealsEvent;
    FUserRequestUpdatedDeals: TIDetectionEvents2UserRequestUpdatedDealsEvent;
    FUserRequestDealsWithAvailableDIE: TIDetectionEvents2UserRequestDealsWithAvailableDIEEvent;
    FUserValidationRequired: TIDetectionEvents2UserValidationRequiredEvent;
  published
    //SinkEventsPublished//
    property EngineReady: TIDetectionEvents2EngineReadyEvent read FEngineReady write FEngineReady;
    property AccountDataRequired: TIDetectionEvents2AccountDataRequiredEvent read FAccountDataRequired write FAccountDataRequired;
    property AccountStateChanged: TIDetectionEvents2AccountStateChangedEvent read FAccountStateChanged write FAccountStateChanged;
    property UserRequestNewDeals: TIDetectionEvents2UserRequestNewDealsEvent read FUserRequestNewDeals write FUserRequestNewDeals;
    property UserRequestSelectedDeals: TIDetectionEvents2UserRequestSelectedDealsEvent read FUserRequestSelectedDeals write FUserRequestSelectedDeals;
    property UserRequestUpdatedDeals: TIDetectionEvents2UserRequestUpdatedDealsEvent read FUserRequestUpdatedDeals write FUserRequestUpdatedDeals;
    property UserRequestDealsWithAvailableDIE: TIDetectionEvents2UserRequestDealsWithAvailableDIEEvent read FUserRequestDealsWithAvailableDIE write FUserRequestDealsWithAvailableDIE;
    property UserValidationRequired: TIDetectionEvents2UserValidationRequiredEvent read FUserValidationRequired write FUserValidationRequired;
  end;

implementation

uses
  SysUtils;

{ globals }

procedure BuildPositionalDispIds (pDispIds: PDispIdList; const dps: TDispParams);
var
  i: integer;
begin
  Assert (pDispIds <> nil);
  
  { by default, directly arrange in reverse order }
  for i := 0 to dps.cArgs - 1 do
    pDispIds^ [i] := dps.cArgs - 1 - i;

  { check for named args }
  if (dps.cNamedArgs <= 0) then Exit;

  { parse named args }
  for i := 0 to dps.cNamedArgs - 1 do
    pDispIds^ [dps.rgdispidNamedArgs^ [i]] := i;
end;


{ TEdisys_IULM_CoreEventsBaseSink }

function TEdisys_IULM_CoreEventsBaseSink.GetIDsOfNames(const IID: TGUID; Names: Pointer;
  NameCount, LocaleID: Integer; DispIDs: Pointer): HResult;
begin
  Result := E_NOTIMPL;
end;

function TEdisys_IULM_CoreEventsBaseSink.GetTypeInfo(Index, LocaleID: Integer; out TypeInfo): HResult;
begin
  Result := E_NOTIMPL;
  pointer (TypeInfo) := nil;
end;

function TEdisys_IULM_CoreEventsBaseSink.GetTypeInfoCount(out Count: Integer): HResult;
begin
  Result := E_NOTIMPL;
  Count := 0;
end;

function TEdisys_IULM_CoreEventsBaseSink.Invoke(DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var Params; VarResult, ExcepInfo, ArgErr: Pointer): HResult;
var
  dps: TDispParams absolute Params;
  bHasParams: boolean;
  pDispIds: PDispIdList;
  iDispIdsSize: integer;
begin
  { validity checks }
  if (Flags AND DISPATCH_METHOD = 0) then
    raise Exception.Create (
      Format ('%s only supports sinking of method calls!', [ClassName]
    ));

  { build pDispIds array. this maybe a bit of overhead but it allows us to
    sink named-argument calls such as Excel's AppEvents, etc!
  }
  pDispIds := nil;
  iDispIdsSize := 0;
  bHasParams := (dps.cArgs > 0);
  if (bHasParams) then
  begin
    iDispIdsSize := dps.cArgs * SizeOf (TDispId);
    GetMem (pDispIds, iDispIdsSize);
  end;  { if }

  try
    { rearrange dispids properly }
    if (bHasParams) then BuildPositionalDispIds (pDispIds, dps);
    Result := DoInvoke (DispId, IID, LocaleID, Flags, dps, pDispIds, VarResult, ExcepInfo, ArgErr);
  finally
    { free pDispIds array }
    if (bHasParams) then FreeMem (pDispIds, iDispIdsSize);
  end;  { finally }
end;

function TEdisys_IULM_CoreEventsBaseSink.QueryInterface(const IID: TGUID; out Obj): HResult;
begin
  if (GetInterface (IID, Obj)) then
  begin
    Result := S_OK;
    Exit;
  end
  else
    if (IsEqualIID (IID, FSinkIID)) then
      if (GetInterface (IDispatch, Obj)) then
      begin
        Result := S_OK;
        Exit;
      end;
  Result := E_NOINTERFACE;
  pointer (Obj) := nil;
end;

function TEdisys_IULM_CoreEventsBaseSink._AddRef: Integer;
begin
  Result := 2;
end;

function TEdisys_IULM_CoreEventsBaseSink._Release: Integer;
begin
  Result := 1;
end;

destructor TEdisys_IULM_CoreEventsBaseSink.Destroy;
begin
  Disconnect;
  inherited;
end;

procedure TEdisys_IULM_CoreEventsBaseSink.Connect (const ASource: IUnknown);
var
  pcpc: IConnectionPointContainer;
begin
  Assert (ASource <> nil);
  Disconnect;
  try
    OleCheck (ASource.QueryInterface (IConnectionPointContainer, pcpc));
    OleCheck (pcpc.FindConnectionPoint (FSinkIID, FCP));
    OleCheck (FCP.Advise (Self, FCookie));
    FSource := ASource;
  except
    raise Exception.Create (Format ('Unable to connect %s.'#13'%s',
      [Name, Exception (ExceptObject).Message]
    ));
  end;  { finally }
end;

procedure TEdisys_IULM_CoreEventsBaseSink.Disconnect;
begin
  if (FSource = nil) then Exit;
  try
    OleCheck (FCP.Unadvise (FCookie));
    FCP := nil;
    FSource := nil;
  except
    pointer (FCP) := nil;
    pointer (FSource) := nil;
  end;  { except }
end;


//SinkImplStart//

function TEdisys_IULM_CoreIDetectionEvents2.DoInvoke (DispID: Integer; const IID: TGUID; LocaleID: Integer;
  Flags: Word; var dps: TDispParams; pDispIds: PDispIdList;
  VarResult, ExcepInfo, ArgErr: Pointer): HResult;
type
  POleVariant = ^OleVariant;
begin
  Result := DISP_E_MEMBERNOTFOUND;

  //SinkInvoke//
    case DispId of
      128 :
      begin
        DoEngineReady ();
        Result := S_OK;
      end;
      129 :
      begin
        DoAccountDataRequired (AccountType (dps.rgvarg^ [pDispIds^ [0]].lval), dps.rgvarg^ [pDispIds^ [1]].pbstrval^, IUnknown (dps.rgvarg^ [pDispIds^ [2]].unkval) as IAccountData);
        Result := S_OK;
      end;
      130 :
      begin
        DoAccountStateChanged ();
        Result := S_OK;
      end;
      131 :
      begin
        DoUserRequestNewDeals ();
        Result := S_OK;
      end;
      132 :
      begin
        DoUserRequestSelectedDeals ();
        Result := S_OK;
      end;
      133 :
      begin
        DoUserRequestUpdatedDeals ();
        Result := S_OK;
      end;
      134 :
      begin
        DoUserRequestDealsWithAvailableDIE ();
        Result := S_OK;
      end;
      135 :
      begin
        DoUserValidationRequired (dps.rgvarg^ [pDispIds^ [0]].pbstrval^, dps.rgvarg^ [pDispIds^ [1]].pbool^);
        Result := S_OK;
      end;
    end;  { case }
  //SinkInvokeEnd//
end;

constructor TEdisys_IULM_CoreIDetectionEvents2.Create (AOwner: TComponent);
begin
  inherited Create (AOwner);
  //SinkInit//
  FSinkIID := IDetectionEvents2;
end;

//SinkImplementation//
procedure TEdisys_IULM_CoreIDetectionEvents2.DoEngineReady;
begin
  if not Assigned (EngineReady) then System.Exit;
  EngineReady (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoAccountDataRequired(AccountType: AccountType; var accountKey: WideString; const AccountData: IAccountData);
begin
  if not Assigned (AccountDataRequired) then System.Exit;
  AccountDataRequired (Self, AccountType, accountKey, AccountData);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoAccountStateChanged;
begin
  if not Assigned (AccountStateChanged) then System.Exit;
  AccountStateChanged (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoUserRequestNewDeals;
begin
  if not Assigned (UserRequestNewDeals) then System.Exit;
  UserRequestNewDeals (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoUserRequestSelectedDeals;
begin
  if not Assigned (UserRequestSelectedDeals) then System.Exit;
  UserRequestSelectedDeals (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoUserRequestUpdatedDeals;
begin
  if not Assigned (UserRequestUpdatedDeals) then System.Exit;
  UserRequestUpdatedDeals (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoUserRequestDealsWithAvailableDIE;
begin
  if not Assigned (UserRequestDealsWithAvailableDIE) then System.Exit;
  UserRequestDealsWithAvailableDIE (Self);
end;

procedure TEdisys_IULM_CoreIDetectionEvents2.DoUserValidationRequired(var password: WideString; var isValidated: WordBool);
begin
  if not Assigned (UserValidationRequired) then System.Exit;
  UserValidationRequired (Self, password, isValidated);
end;


end.

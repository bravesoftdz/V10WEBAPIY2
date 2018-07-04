unit MC_Comm;

interface

{uses
  Windows, SysUtils, Classes, Forms, Dialogs, HEnt1, MC_Timer, mc_Erreur  ;}

Uses Windows, Classes, MC_Timer ;


type
  // COM Port Baud Rates
  TComPortBaudRate = ( br110, br300, br600, br1200, br2400, br4800,
                       br9600, br14400, br19200, br38400, br56000,
                       br57600, br115200, br128000, br256000 );
  // COM Port Numbers
  TComPortNumber = ( pnNone, pnFile, pnCOM1, pnCOM2, pnCOM3, pnCOM4, pnLPT1, pnLPT2 );
  // COM Port Data bits
  TComPortDataBits = ( db5BITS, db6BITS, db7BITS, db8BITS );
  // COM Port Stop bits
  TComPortStopBits = ( sb1BITS, sb1HALFBITS, sb2BITS );
  // COM Port Parity
  TComPortParity = ( ptNONE, ptODD, ptEVEN, ptMARK, ptSPACE );
  // COM Port Hardware Handshaking
  TComPortHwHandshaking = ( hhNONE, hhRTSCTS, hhDTRDSR );
  // COM Port Software Handshaing
  TComPortSwHandshaking = ( shNONE, shXONXOFF );

  TComPortReceiveDataEvent = procedure( Sender: TObject; DataPtr: pointer; DataSize: integer ) of object;


  TRecPOrt = Record
             Speed    : TComPortBaudRate ;
             Parite   : TComPortParity ;
             DataBits : TComPortDataBits ;
             StopBits : TComPortStopBits ;
             Handle   : THandle ;
             Conteur  : Integer ;
             End ;

  TArrayPorts = Array[pnCOM1..pnLPT2] of TRecPort ;

Var ArrayPorts : TArrayPorts ;
Type

  TMC_CommportDriver = class(TComponent)
  private

    FHandle             : THANDLE; // COM Port Device Handle
    FComPort            : TComPortNumber; // COM Port to use (1..4)
    FBaudRate           : TComPortBaudRate; // COM Port speed (brXXXX)
    FDataBits           : TComPortDataBits; // Data bits size (5..8)
    FStopBits           : TComPortStopBits; // How many stop bits to use (1,1.5,2)
    FParity             : TComPortParity; // Type of parity to use (none,odd,even,mark,space)
    FHwHandshaking      : TComPortHwHandshaking; // Type of hw handshaking to use
    FSwHandshaking      : TComPortSwHandshaking; // Type of sw handshaking to use
    FInBufSize          : word; // Size of the input buffer
    FOutBufSize         : word; // Size of the output buffer
    FXOnChar            : Char ;
    FXOffChar           : Char ;
    FErrorChar          : Char ;
    FReceiveData        : TComPortReceiveDataEvent; // Event to raise on data reception

    FNomFichierSortie   : String ; //XMG 30/03/04

    FTempInBuffer              : pointer;
    FPosInBuffer               : DWord ;
    FPrtTimer                  : TMC_Timer ;

    FLastError                 : DWord ;

    FAttendre                  : Boolean ;

    procedure SetComPort( Value: TComPortNumber );
    procedure SetBaudRate( Value: TComPortBaudRate );
    procedure SetDataBits( Value: TComPortDataBits );
    procedure SetStopBits( Value: TComPortStopBits );
    procedure SetParity( Value: TComPortParity );
    procedure SetHwHandshaking( Value: TComPortHwHandshaking );
    procedure SetSwHandshaking( Value: TComPortSwHandshaking );
    procedure SetInBufSize( Value: word );
    procedure SetOutBufSize( Value: word );
    procedure SetXOnChar ( Value: Char );
    procedure SetXOffChar ( Value: Char );
    procedure SetErrorChar ( Value: Char );

    Function  ApplyCOMSettings : Boolean ;

    procedure SetInterval( Value: Cardinal );
    Function GetIsPrinter : Boolean ;

    Procedure IsTime ( Sender : tObject ) ;
    Function GetInterval : Cardinal ;
    Function GetPartage : Boolean ;
    procedure SetAttendre ( Value: Boolean );
  public
    constructor Create( AOwner: TComponent ); override;
    destructor Destroy; override;

    function Connect: boolean;
    Function Disconnect : Boolean ;
    function Connected: boolean;

    function SendData( DataPtr: pointer; DataSize: LongWord ): boolean;
    function SendString( s: string ): boolean;
    Property IsPrinter : Boolean Read GetIsPrinter ;
    Property LastError : LongWord read FLastError ;
    Property Partage   : Boolean Read GetPartage ;
    Property AttendreResponse : Boolean Read FAttendre write setAttendre ;
  published
    // Com port device Handle used
    property Handle : THANDLE read FHandle ;
    // Which COM Port to use
    property ComPort: TComPortNumber read FComPort write SetComPort default pnNone ;
    // COM Port speed (bauds)
    property Speed: TComPortBaudRate read FBaudRate write SetBaudRate default br9600;
    // Data bits to used (5..8, for the 8250 the use of 5 data bits with 2 stop bits is an invalid combination,
    // as is 6, 7, or 8 data bits with 1.5 stop bits)
    property DataBits: TComPortDataBits read FDataBits write SetDataBits default db8BITS;
    // Stop bits to use (1, 1.5, 2)
    property StopBits: TComPortStopBits read FStopBits write SetStopBits default sb1BITS;
    // Parity Type to use (none,odd,even,mark,space)
    property Parity: TComPortParity read FParity write SetParity default ptNONE;
    // Hardware Handshaking Type to use:
    //  cdNONE          no handshaking
    //  cdCTSRTS        both cdCTS and cdRTS apply (** this is the more common method**)
    property HwHandshaking: TComPortHwHandshaking read FHwHandshaking write SetHwHandshaking default hhNONE;
    // Software Handshaking Type to use:
    //  cdNONE          no handshaking
    //  cdXONXOFF       XON/XOFF handshaking
    property SwHandshaking: TComPortSwHandshaking read FSwHandshaking write SetSwHandshaking default shNONE;
    // Input Buffer size
    property InBufSize: word read FInBufSize write SetInBufSize default 2048;
    // Output Buffer size
    property OutBufSize: word read FOutBufSize write SetOutBufSize default 2048;
    // ms of delay between COM port pollings
    property PollingDelay: Cardinal read GetInterval write SetInterval default 100;
    // character for XOn
    property XOnChar : Char read FXOnChar write SetXOnChar default #17 ;
    // character for XOff
    property XOffChar : Char read FXOffChar write SetXOffChar default #19 ;
    // remplace parity errors for character
    property ErrorChar : Char read FErrorChar write SetErrorChar default #0 ;
    // Event to raise when there is data available (input buffer has data)
    property OnReceiveData: TComPortReceiveDataEvent read FReceiveData write FReceiveData;

    Property NomFichierSortie : String read FNomFichierSortie write FNomFichierSortie ; //XMG 30/03/04
  end;


implementation

uses Dialogs, Hent1, MC_Erreur, Sysutils, Forms ;

{$R *.DCR}

Function PortInUse(PortCom : TComPortNumber) : Boolean ;
Begin
if PortCom=pnFile then Result:=FALSE
   else Result:=ArrayPorts[PortCom].Handle<>INVALID_HANDLE_VALUE ;
End ;

Function CompatiblePort(PortCom : TComPortNumber ; Speed : TComPortBaudRate ; Parite : TComPortParity ; DataBits : TComPortDataBits ; StopBits : TComPortStopBits) : Boolean ;
Begin
Result:=FALSE ;
if ArrayPorts[PortCom].Speed<>Speed      then exit ;
if ArrayPorts[PortCom].Parite>Parite     then exit ;
if ArrayPorts[PortCom].DataBits>DataBits then exit ;
if ArrayPorts[PortCom].StopBits>StopBits then exit ;
if (PortCom in [pnLPT1,pnLPT2]) and (ArrayPorts[PortCom].Conteur>0) then exit ;
Result:=TRUE ;
End ;


Const MaxPourcentBuffer = 0.8 ;  //80% du buffer

constructor TMC_CommportDriver.Create( AOwner: TComponent );
begin
  inherited Create( AOwner );
  // Initialize to default values
  FHandle             := INVALID_HANDLE_VALUE ;       // Not connected
  FComPort            := pnNone ; // Not one port
  FBaudRate           := br9600;  // 9600 bauds
  FDataBits           := db8BITS; // 8 data bits
  FStopBits           := sb1BITS; // 1 stop bit
  FParity             := ptNONE;  // no parity
  FHwHandshaking      := hhNONE;  // no hardware handshaking
  FSwHandshaking      := shNONE;  // no software handshaking
  FInBufSize          := 2048;    // input buffer of 512 bytes
  FOutBufSize         := 2048;    // output buffer of 512 bytes
  FReceiveData        := nil;     // no data handler
  GetMem( FTempInBuffer, FInBufSize ); // Temporary buffer for received data ;
  FPosInBuffer        :=0 ;       // current position in temporary input buffer

  XOnChar             :=#17 ;
  XOffChar            :=#19 ;
  ErrorChar           :=#0 ;

  FAttendre:=TRUE ;
  FPrtTimer                     :=TMC_Timer.Create(Self) ;
  FPrtTimer.Enabled             :=FALSE ;
  FPrtTimer.delay               :=20 ; // 1/1000 de seconde
  FPrtTimer.OnTimer             :=IsTime ;
end;

destructor TMC_CommportDriver.Destroy;
begin
  // Be sure to release the COM device
  Disconnect;
  // Free the temporary buffer
  FreeMem( FTempInBuffer, InBufSize );
  // Destroy the timer's window

  FPrtTimer.Free ;
  inherited Destroy;
end;

procedure TMC_CommportDriver.SetComPort( Value: TComPortNumber );
begin
  if Value=ComPort then exit ;
  // Be sure we are not using any COM port
  if Connected then
    exit;
  // Change COM port
  FComPort := Value;
end;

procedure TMC_CommportDriver.SetBaudRate( Value: TComPortBaudRate );
begin
  if Value=speed then exit ;
  // Set new COM speed
  FBaudRate := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetDataBits( Value: TComPortDataBits );
begin
  if Value=DataBits then exit ;
  // Set new data bits
  FDataBits := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetStopBits( Value: TComPortStopBits );
begin
  if Value=StopBits then exit ;
  // Set new stop bits
  FStopBits := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetParity( Value: TComPortParity );
begin
  if Value=Parity then exit ;
  // Set new parity
  FParity := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetHwHandshaking( Value: TComPortHwHandshaking );
begin
  if Value=HwHandshaking then exit ;
  // Set new hardware handshaking
  FHwHandshaking := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetSwHandshaking( Value: TComPortSwHandshaking );
begin
  if Value=SwHandshaking then exit ;
  // Set new software handshaking
  FSwHandshaking := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetInBufSize( Value: word );
begin
  if Value=InBufSize then exit ;
  // Free the temporary input buffer
  FreeMem( FTempInBuffer, InBufSize );
  // Set new input buffer size
  FInBufSize := Value;
  // Allocate the temporary input buffer
  GetMem( FTempInBuffer, InBufSize );
  FPosInBuffer:=0 ;
  Fillchar(FTempInBuffer^,InBufsize,#0) ;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetOutBufSize( Value: word );
begin
  if Value=OutBufSize then exit ;
  // Set new output buffer size
  FOutBufSize := Value;
  // Apply changes
  if Connected then
    ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetXOnChar ( Value: Char );
begin
if Value=XOnChar then exit ;
FXOnChar:= Value;
if Connected then ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetXOffChar ( Value: Char );
begin
if Value=XOffChar then exit ;
FXOffChar:= Value;
if Connected then ApplyCOMSettings;
end;

procedure TMC_CommportDriver.SetErrorChar ( Value: Char );
begin
if Value=ErrorChar then exit ;
FErrorChar:= Value;
if Connected then ApplyCOMSettings;
end;

Function TMC_CommportDriver.GetInterval : Cardinal ;
Begin
Result:=FPrtTimer.delay ;
End ;

Function TMC_CommportDriver.GetPartage : Boolean ;
Begin
result:=Portinuse(Comport) and (ArrayPorts[Comport].Conteur>1) ;
End ;

procedure TMC_CommportDriver.SetAttendre ( Value: Boolean );
Begin
if Value=Attendreresponse then exit ;
FAttendre:=Value ;
if connected then FPrttimer.Enabled:=Value ;
End ;

procedure TMC_CommportDriver.SetInterval ( Value: Cardinal );
begin
  if Value = FPrtTimer.delay then exit ;
  FPrtTimer.delay:= Value;
end;

Function TMC_CommportDriver.GetIsPrinter : Boolean ;
Begin
Result:=(ComPort in [pnLPT1..pnLPT2]) ;
End ;

const
  Win32BaudRates: array[br110..br256000] of DWORD =
    ( CBR_110, CBR_300, CBR_600, CBR_1200, CBR_2400, CBR_4800, CBR_9600,
      CBR_14400, CBR_19200, CBR_38400, CBR_56000, CBR_57600, CBR_115200,
      CBR_128000, CBR_256000 );
  DevicePortNames : Array[pnCOM1..pnLPT2] of String =
    ('COM1', 'COM2', 'COM3', 'COM4', 'LPT1', 'LPT2' ) ;

const
  dcb_Binary              = $00000001;
  dcb_ParityCheck         = $00000002;
  dcb_OutxCtsFlow         = $00000004;
  dcb_OutxDsrFlow         = $00000008;
  dcb_DtrControlMask      = $00000030;
    dcb_DtrControlDisable   = $00000000;
    dcb_DtrControlEnable    = $00000010;
    dcb_DtrControlHandshake = $00000020;
  dcb_DsrSensivity        = $00000040;
  dcb_TXContinueOnXoff    = $00000080;
  dcb_OutX                = $00000100;
  dcb_InX                 = $00000200;
  dcb_ErrorChar           = $00000400;
  dcb_NullStrip           = $00000800;
  dcb_RtsControlMask      = $00003000;
    dcb_RtsControlDisable   = $00000000;
    dcb_RtsControlEnable    = $00001000;
    dcb_RtsControlHandshake = $00002000;
    dcb_RtsControlToggle    = $00003000;
  dcb_AbortOnError        = $00004000;
  dcb_Reserveds           = $FFFF8000;

// Apply COM settings.
Function TMC_CommportDriver.ApplyCOMSettings : Boolean ;
var dcb: TDCB;
begin
Result:=FALSE ;
  // Do nothing if not connected
  if not Connected then exit;
  if Not IsPrinter then  //XMG 23/06/03
     Begin
     // Retrieve old dcb
     if Not GetCommState(Handle,Dcb) then Begin Disconnect ; Exit ; End ;
     // Setup dcb (Device Control Block) fields
     dcb.DCBLength := sizeof(dcb); // dcb structure size
     dcb.BaudRate := Win32BaudRates[ Speed]; // baud rate to use
     dcb.Flags := dcb_Binary or  // Set fBinary: Win32 does not support non binary mode transfers
                                // (also disable EOF check)
                  dcb_ParityCheck ; //Enables parityCheck
     if Errorchar<>#0 then
        Begin
        dcb.flags:=dcb.Flags or dcb_ErrorChar ; //if Errorchar<>#0 then allow replace on error
        dcb.ErrorChar:=ErrorChar ;
        End ;

     case HwHandshaking of // Type of hw handshaking to use
       hhNONE:; // No hardware handshaking
       hhRTSCTS: // RTS/CTS (request-to-send/clear-to-send) hardware handshaking
         dcb.Flags := dcb.Flags or dcb_OutxCtsFlow or dcb_RtsControlHandshake;
       hhDTRDSR: //DTR / DSR (Data terminal ready/Data set reasy) harrdware handshaking
         dcb.Flags := dcb.Flags or dcb_OutxDsrFlow or dcb_DsrSensivity or dcb_DtrControlHandshake ;
     end;
     case SwHandshaking of // Type of sw handshaking to use
       shNONE:; // No software handshaking
       shXONXOFF: // XON/XOFF handshaking
         dcb.Flags := dcb.Flags or dcb_OutX or dcb_InX;
     end;
     dcb.XONLim := InBufSize div 4; // Specifies the minimum number of bytes allowed
                                    // in the input buffer before the XON character is sent
     dcb.XOFFLim := 1; // Specifies the maximum number of bytes allowed in the input buffer
                       // before the XOFF character is sent. The maximum number of bytes
                       // allowed is calculated by subtracting this value from the size,
                       // in bytes, of the input buffer
     dcb.ByteSize := 5 + ord(DataBits); // how many data bits to use
     dcb.Parity := ord(Parity); // type of parity to use
     dcb.StopBits := ord(Stopbits); // how many stop bits to use
     dcb.XONChar := XOnChar ;  // XON ASCII char
     dcb.XOFFChar := XOffChar ;  // XOFF ASCII char
     dcb.ErrorChar:=Errorchar ;
     if Not SetCommState( Handle, dcb ) then Begin Disconnect ; Exit ; End ;
     // Setup buffers size
     if Not SetupComm( Handle, InBufSize, OutBufSize ) then Begin Disconnect ; Exit ; End ;
     End ;
Result:=TRUE ;
end;

function TMC_CommportDriver.Connect: boolean;
var tms: TCOMMTIMEOUTS;
    NomFile : String ;
    Exist : Integer ;
    SaveDlg : TSaveDialog ;
begin
Result:=FALSE ;
if Comport=pnNone then exit ;
// Do nothing if already connected
Result := Connected;
if Result then exit;
if Comport=pnFile then
   Begin
   //XMG 30/03/04 début
   if trim(NomFichierSortie)<>'' then Nomfile:=NomFichierSortie else
      Begin
      SaveDlg:=TSaveDialog.Create(Self) ;
      try
         with SaveDlg do
           Begin
           FileName:='*.txt' ;
           Filter:=TraduireMemoire(MC_MsgErrDefaut(1038)) ;
           FilterIndex:=0 ;
           Title:=traduireMemoire(MC_MsgErrDefaut(1039)) ;
           Options:=[ofOverwritePrompt,ofHideReadOnly,ofShowHelp,ofPathMustExist] ;
           if Execute then NomFile:=FileName ;
           End ;
        Finally
         SaveDlg.Free ;
        End ;
      End ;
   //XMG 30/03/04 fin
   If trim(NomFile)='' then exit ;
   Exist:=CREATE_ALWAYS ;
   End Else
   Begin
   NomFile:=DevicePortNames[ComPort] ;
   Exist:=OPEN_EXISTING ;
   End ;
if not PortinUse(Comport) then
   Begin // Open the port
   FHandle := CreateFile(       PChar(NomFile),
                                GENERIC_READ or GENERIC_WRITE,
                                0,// Not shared
                                nil, // No security attributes
                                Exist,
                                FILE_ATTRIBUTE_NORMAL,
                                //FILE_FLAG_WRITE_THROUGH or FILE_FLAG_NO_BUFFERING or FILE_FLAG_SEQUENTIAL_SCAN,
                                0 // No template
                              ) ;
   Result := Connected;
   if (not Result) or (ComPort=pnFile) then exit;
   // Apply settings

   if not ApplyCOMSettings then Begin Result:=FALSE ; exit ; ENd ; //XMG 23/06/03
   if not IsPrinter then //XMG 23/06/03
     Begin
     // Setup timeouts: we disable timeouts because we are polling the com port!
     tms.ReadIntervalTimeout := high(LongWord) ; //1 ; // Specifies the maximum time, in milliseconds,  //XMG 23/07/01
                                   // allowed to elapse between the arrival of two
                                   // characters on the communications line
     tms.ReadTotalTimeoutMultiplier := high(LongWord) ; //0 ; // Specifies the multiplier, in milliseconds, //XMG 23/07/01
                                          // used to calculate the total time-out period
                                          // for read operations.
     tms.ReadTotalTimeoutConstant := 1 ; // Specifies the constant, in milliseconds,  //XMG 23/07/01
                                        // used to calculate the total time-out period
                                        // for read operations.
     tms.WriteTotalTimeoutMultiplier := ord(IsPrinter) ; // Specifies the multiplier, in milliseconds,
                                                         // used to calculate the total time-out period
                                                         // for write operations.
     tms.WriteTotalTimeoutConstant :=  1000 ; //TimeOutConst ; // Specifies the constant, in milliseconds,
                                                      // used to calculate the total time-out period
                                                      // for write operations.
     if not SetCommTimeOuts( Handle, tms ) then
        Begin
        Result:=FALSE ;
        Disconnect ;
        Exit ;
        End ;
     End ;
   ArrayPorts[ComPort].Handle:=Handle ;
   ArrayPorts[ComPort].Speed:=Speed ;
   ArrayPorts[ComPort].Parite:=Parity ;
   ArrayPorts[ComPort].DataBits:=DataBits ;
   ArrayPorts[ComPort].StopBits:=StopBits ;
   End else
   if CompatiblePort(ComPort,Speed,parity,Databits,stopbits) then FHandle:=ArrayPorts[Comport].Handle
      else FLastError:=5 ;
if connected then
   Begin
   if ComPort in [pnCOM1..pnLPT2] then inc(ArrayPorts[ComPort].Conteur) ;
   // Start the timer (used for polling)
   FPrtTimer.Enabled:=AttendreResponse ;
   Result:=TRUE ;
   End ;
end;

Function TMC_CommportDriver.Disconnect : Boolean ;
begin
Result:=FALSE ;
if Connected then
  begin
  // Stop the timer (used for polling)
  FPrtTimer.Enabled:=FALSE ;
  if ComPort in [pnCOM1..pnLPT2] then dec(ArrayPorts[Comport].Conteur) ;
  if ((ComPort in [pnCOM1..pnLPT2])=FALSE) or (ArrayPorts[Comport].Conteur<=0) then
     Begin
     Result:=CloseHandle( Handle );
     if ComPort in [pnCOM1..pnLPT2] then
        Begin
        FillChar(ArrayPorts[ComPort],sizeof(ArrayPorts[ComPort]),#0) ;
        ArrayPorts[ComPort].Handle:=INVALID_HANDLE_VALUE ;
        End ;
     End else result:=TRUE ;
  FHandle := INVALID_HANDLE_VALUE ;
  end;
end;

function TMC_CommportDriver.Connected: boolean;
begin
Result := Handle <> INVALID_HANDLE_VALUE ;
end;

function TMC_CommportDriver.SendData( DataPtr: pointer; DataSize: LongWord ): boolean;
var nsent : DWORD;
Begin
FLastError:=0 ;
if (DataSize<=0) or (Not Connected) then
   Begin
   result:=TRUE ;
   Exit ;
   End ;
try
   Result:=WriteFile(Handle,DataPtr^,DataSize,nSent,nil );
   if Not Result then FLastError:=GetLastError
      else Result :=(Result) and (nsent=DataSize);
 except
   raise ;
 End ;
end;

function TMC_CommportDriver.SendString( s: string ): boolean;
begin
Result := SendData( pchar(s), length(s) );
end;

procedure TMC_CommportDriver.IsTime ( Sender : tObject ) ;
var nRead      : dword;
    CourantPos : Pointer ;
    i          : DWord ;
begin
if Connected then
  begin
  nRead := 0;
  if (FPosInBuffer>(InBufSize*MaxPourcentbuffer)) and (Not Assigned(OnReceiveData)) then //Si le MaxPourcentbuffer% du buffer est plein et il n'y a pas d'event liée, alors on le supprime
     Begin
     FPosInBuffer:=0 ;
     Fillchar(FTempInBuffer^,InBufsize,#0) ;
     End ;
  CourantPos:=FTempInBuffer ;
  i:=FPosInBuffer ; while i>0 do Begin inc(PChar(CourantPos)) ; Dec(i) ; End ;
  if ReadFile( Handle, CourantPos^, InBufSize-FPosInBuffer , nRead, Nil) then
    if (nRead <> 0) then
       Begin
       FPosInBuffer:=FPosInBuffer+nRead ;
       if Assigned(OnReceiveData) then
          Begin
          OnReceiveData( Self, FTempInBuffer, FPosInBuffer);
          FPosInBuffer:=0 ;
          End ;
       End ;
  end;
Application.ProcessMessages ;
end;



var i : TComPortNumber ;
initialization

Fillchar(ArrayPorts,sizeof(ArrayPorts),#0) ;

for i:=low(ArrayPorts) to High(ArrayPorts) do ArrayPorts[i].Handle:= INVALID_HANDLE_VALUE ;
end.

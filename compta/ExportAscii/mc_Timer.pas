unit mc_Timer;
interface

uses Classes, Windows, Sysutils ;

type
 TMC_Timer=class(TComponent)
  private
    FDelay       : Cardinal;
    FResolution  : Cardinal;
    OnTimerEvent : TNotifyEvent;
    TimerHandle  : THandle ;
    FPAused      : Boolean ;
  protected
    procedure InitTimer;
    Function  GetEnabled : Boolean ;
    procedure SetEnabled (NewOn: Boolean);
    procedure SetPaused(AValue : Boolean);
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure Resume;
    procedure Pause;
    property  Paused   : Boolean read FPaused write setPaused ;
    property  Enabled  : Boolean read GetEnabled write SetEnabled default False;
  published
    property Delay: Cardinal read FDelay write FDelay default 100;
    property Resolution: Cardinal read FResolution write FResolution default 10;
    property OnTimer: TNotifyEvent read OnTimerEvent write OnTimerEvent;
  end;

 ELP_Timer=class(Exception);



implementation

uses mmSystem ; 

/////////////////////////////////////////////////////////////////////////////////////////
constructor TMC_Timer.Create(AOwner: TComponent);
begin
inherited Create(AOwner);
Enabled:=False;
FDelay:=100;
Resolution:=10;
TimerHandle:=0 ;
FPaused:=TRUE ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
destructor TMC_Timer.Destroy;
begin
Enabled:=False;
inherited Destroy;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure LP_TimerCallBack(NA1,NA2,TimerUser,NA3,NA4: Integer); stdcall;
var Timer: TMC_Timer;
begin
Timer:=TMC_Timer(TimerUser);
if (not Timer.Paused) and (Assigned(Timer.OnTimer)) then Timer.OnTimer(Timer) ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TMC_Timer.InitTimer;
begin
TimerHandle:=timesetevent(FDelay,Resolution,@LP_TimerCallBack,Integer(Self),TIME_PERIODIC) ;
if TimerHandle=0 then
   begin
   Enabled:=False;
   raise ELP_Timer.Create('Erreur lors la création de LP_timer.');
   end;
end;
/////////////////////////////////////////////////////////////////////////////////////////
Function TMC_Timer.GetEnabled : Boolean ;
Begin
Result:=TimerHandle<>0 ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TMC_Timer.SetEnabled (NewOn: Boolean);
begin
if (NewOn=Enabled) or (csDesigning in ComponentState) then exit ;
if NewOn then
   begin
   InitTimer ;
   Resume ;
   end;
if not NewOn then
   begin
   TimeKillEvent(TimerHandle) ;
   TimerHandle:=0 ;
   end;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TMC_Timer.SetPaused(AValue : Boolean);
Begin
if (Enabled) and (AValue<>Paused) then FPaused:=AValue ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TMC_Timer.Pause;
begin
if (Enabled) and (not Paused) then FPaused:=TRUE ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
procedure TMC_Timer.Resume;
begin
if (Enabled) and (Paused) then FPaused:=FALSE ;
end;
/////////////////////////////////////////////////////////////////////////////////////////
end.

unit tThreadTiers;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadTiers = class(TThread)
  public
    TiersValues : T_TiersValues;
    LogValues   : T_WSLogValues;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;
  private
    procedure SetName;
  protected
    procedure Execute; override;
  end;

implementation

uses
  CommonTools
  , SysUtils
  ;

{ Important : les m�thodes et propri�t�s des objets de la VCL peuvent uniquement �tre
  utilis�s dans une m�thode appel�e en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  o� UpdateCaption serait de la forme

    procedure ThreadTiers.UpdateCaption;
    begin
      Form1.Caption := 'Mis � jour dans un thread';
    end; }

{$IFDEF MSWINDOWS}
type
  TThreadNameInfo = record
    FType: LongWord;     // doit �tre 0x1000
    FName: PChar;        // pointeur sur le nom (dans l'espace d'adresse de l'utilisateur)
    FThreadID: LongWord; // ID de thread (-1=thread de l'appelant)
    FFlags: LongWord;    // r�serv� pour une future utilisation, doit �tre z�ro
  end;
{$ENDIF}

{ ThreadTiers }

procedure ThreadTiers.SetName;
{$IFDEF MSWINDOWS}
var
  ThreadNameInfo: TThreadNameInfo;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  ThreadNameInfo.FType := $1000;
  ThreadNameInfo.FName := 'ThreadNameTiers';
  ThreadNameInfo.FThreadID := $FFFFFFFF;
  ThreadNameInfo.FFlags := 0;

  try
    RaiseException( $406D1388, 0, sizeof(ThreadNameInfo) div sizeof(LongWord), @ThreadNameInfo );
  except
  end;
{$ENDIF}
end;

constructor ThreadTiers.Create(CreateSuspended : boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
  TServicesLog.WriteLog(ssbylWindows, TUtilBTPVerdon.GetMsg(tnTiers, True), ServiceName_BTPVerdon, LogValues, 0);
end;

destructor ThreadTiers.Destroy;
begin
  inherited;
  TServicesLog.WriteLog(ssbylWindows, TUtilBTPVerdon.GetMsg(tnTiers, False), ServiceName_BTPVerdon, LogValues, 0);
end;

procedure ThreadTiers.Execute;
begin
  SetName;
  Sleep(10000);
  { Placer le code du thread ici }
end;

end.

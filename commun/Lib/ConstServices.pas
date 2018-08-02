unit ConstServices;

interface

type
  T_WSLogValues       = Record
                          LogLevel          : integer;
                          LogMoMaxSize      : double;
                          LogMaxQty         : integer;
                          LogDebug          : integer;
                          LogDebugMoMaxSize : double;
                          DebugEvents       : integer;
                          OneLogPerDay      : boolean;
                       end;

  T_SvcTypeLog  = (ssbylNone, ssbylLog, ssbylWindows);

const
  ServiceName_BTPY2     = 'SvcSynBTPY2';
  ServiceName_BTPVerdon = 'SvcSynBTPVerdon';


implementation

end.

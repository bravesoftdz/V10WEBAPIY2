unit BullOLE;

interface

uses Dialogs, Classes, StdCtrls, ComCtrls, Forms, SysUtils
{$IFDEF SERIE1}
;
{$ELSE}
,Ent1;
{$endIF}

//Procedure ImZoomEdtEtat(Quoi : String) ;
function BullCalcOLEEtat(sf,sp : string) : variant ;

implementation

uses PGEdtEtat;

(*Procedure ImZoomEdtEtat(Quoi : String) ;
begin
{$IFDEF AMORTISSEMENT}
{$IFDEF SERIE1}
{$ELSE}
{$IFDEF EAGLCLIENT}
{$ELSE}
ZoomEdtEtat(Quoi);
{$ENDIF}
{$ENDIF}
{$ENDIF}
END ;*)

function BullCalcOLEEtat(sf,sp : string) : variant ;
begin
{$IFDEF SERIE1}
{$ELSE}
result := CalcOLEEtatPG(sf,sp) ;
{$ENDIF}
END ;

end.


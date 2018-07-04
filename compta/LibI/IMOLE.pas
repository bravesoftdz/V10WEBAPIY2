{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/12/2005
Modifié le ... : 21/12/2005 MVg pour tests Teamlink
               : 28/02/2006 MVG - Report des corrections de C. Ayel
Description .. :
Mots clefs ... :
*****************************************************************}
unit ImOLE;

interface

uses Dialogs, Classes, StdCtrls, ComCtrls, Forms, SysUtils
{$IFDEF SERIE1}
;
{$ELSE}
,Ent1;
{$endIF}

procedure ImZoomEdtEtatImmo(Quoi : String) ;
Procedure ImZoomEdtEtat(Quoi : String) ;
function ImCalcOLEEtat(sf,sp : string) : variant ;

implementation

{$IFDEF SERIE1}
{$ELSE}
  uses
  {$IFDEF CCSTD}
  {$ELSE}
    CPZOOM,
  {$ENDIF}
  {$IFDEF SERIE1}
  {$ELSE}
    uLibCalcEdtCompta;
  {$ENDIF}
{$ENDIF}

procedure ImZoomEdtEtatImmo(Quoi : String) ;
begin
{$IFDEF AMORTISSEMENT}
{$IFDEF SERIE1}
{$ELSE}
{$IFNDEF IMP}
{$IFNDEF CCMP}
{$IFDEF EAGLCLIENT}
{$ELSE}
ZoomEdtEtatImmo(Quoi) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

Procedure ImZoomEdtEtat(Quoi : String) ;
begin
{$IFDEF AMORTISSEMENT}
{$IFDEF SERIE1}
{$ELSE}
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFDEF CCSTD}
{$ELSE}
ZoomEdtEtat(Quoi);
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}
END ;

function ImCalcOLEEtat(sf,sp : string) : variant ;
begin
{$IFDEF SERIE1}
{$ELSE}
result := CalcOLEEtat(sf,sp) ;
{$ENDIF}
END ;

end.


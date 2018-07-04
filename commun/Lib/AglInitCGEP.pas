unit AglInitCGEP;

interface
uses M3FP,
{$IFDEF EAGLCLIENT}
     MenuOLX,Maineagl
{$ELSE}
     MenuOLG
{$ENDIF} ;


implementation
Procedure AGLMenuRemoveItem(Parms: Array of variant; Nb: Integer );
begin
FMenuG.RemoveItem(Parms[0]);
end;

Procedure AGLMenuRemoveGroup(Parms: Array of variant; Nb: Integer );
begin
FMenuG.RemoveGroup(Parms[0]);
end;

Initialization
RegisterAglProc( 'MenuRemoveItem', FALSE , 1, AGLMenuRemoveItem);
RegisterAglProc( 'MenuRemoveGroup', FALSE , 1, AGLMenuRemoveGroup);
end.





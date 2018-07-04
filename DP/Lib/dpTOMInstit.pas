Unit dpTOMInstit ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTob,
{$IFDEF EAGLCLIENT}
     eFiche,
     eFichList,
{$ELSE}
     Fiche,
     FichList,
     dbTables,
{$ENDIF}
     UTOM;

Type
  TOM_DPINSTIT = Class (TOM)
    procedure OnArgument ( S: String )   ; override ;
    end ;

Implementation


procedure TOM_DPINSTIT.OnArgument ( S: String ) ;
begin
  Inherited ;
  // 20/06/01
  SetControlVisible('BINSERT', False);
end ;

Initialization
  registerclasses ( [ TOM_DPINSTIT ] ) ;
end.


Unit dpTOMEco ;
// TOM de la table DPECO

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOM, UTob ;

Type
  TOM_DPECO = Class (TOM)
    procedure OnLoadRecord ; override ;
    end ;

////////////// IMPLEMENTATION //////////////
Implementation

uses DpJurOutils;


procedure TOM_DPECO.OnLoadRecord ;
begin
  Inherited ;
  if Ecran.Name='DPCOMECO' then ModeEdition(DS);
end ;

Initialization
  registerclasses ( [ TOM_DPECO ] ) ;
end.


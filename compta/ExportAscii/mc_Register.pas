unit mc_Register;

interface

procedure Register;

implementation

uses Classes, MC_Timer, LP_Base ; 

{$IFNDEF EXPORTASCII}
const MC_NomOnglet = 'Mat. Caisse' ;
{$ELSE}
const MC_NomOnglet = 'Export ASCII' ;
{$ENDIF EXPORTASCII}

procedure Register;
begin
  RegisterComponents(MC_NomOnglet,[TMC_Timer,TLPBase]) ;
end;

end.

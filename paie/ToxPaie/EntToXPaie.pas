unit EntToxPaie;

interface

uses SysUtils ;

Procedure ChargeParamsPGI ;   // Dans ChargeMagHalley (Encore dans Ent1 )

var ProcChargeV_PGI : TProcedure;

implementation

Uses  Hent1,
//      Ent1,
      EntPaie;

Procedure ChargeParamsPGI ;
Begin
  ChargeParamsPaie ;
End;

Procedure InitVariablesPGI ;
Begin
//  InitLaVariableHalley ;  // Compta pour tout le monde !
  InitLaVariablePaie ;
End;

Procedure LibereParamsPGI ;
Begin
  VH_Paie.Free ;
  VH_Paie:=Nil ;
End;

Procedure InitialisationAGL ;
Begin
  If assigned(ProcChargeV_PGI) then TProcedure(ProcChargeV_PGI);
  InitVariablesPGI ;
End;

Initialization
  InitProcAGL := InitialisationAGL ;
Finalization
  LibereParamsPGI ;
end.

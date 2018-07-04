unit PlanningUtil;

interface

Uses HEnt1, HCtrls, UTOB, Ent1,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     DB, DBTables,FE_Main,
{$ENDIF}
     SysUtils, Dialogs, UtilPGI, AGLInit, EntGC, HMsgBox, paramsoc;

function ChargerTOBPlanning():TOB;
procedure AppelPlanningMDisp (Num : integer) ;

implementation

//uses planning;

procedure AppelPlanningMDisp (Num : integer) ;
Var CodePlanning:string;
    TobModelePlanning:TOB;
    IndiceTob:integer;
begin
try
TobModelePlanning:=ChargerTOBPlanning;

if (TobModelePlanning<>nil) then
if (TobModelePlanning.Detail.count >= 1) then
  begin
    IndiceTob := Num - 153201;
    CodePlanning:=TobModelePlanning.Detail[IndiceTob].GetValue('HPP_PARAMPLANNING');
    //AGLLanceFiche('AFF','PLANNING','','','CODE='+ CodePlanning);
    //ExecPlanning;
  end;

finally
TobModelePlanning.Free;
end;
end;


function ChargerTOBPlanning():TOB;
var
TOBPlanning:TOB;
begin
Result:=nil;
TOBPlanning:=TOB.Create ('les modeles', Nil, -1);
TOBPlanning.LoadDetailDB ('HRPARAMPLANNING', '', '',Nil, True) ;
Result:= TOBPlanning;
end;

end.

unit CtrCptaEdServer;

interface

uses classes, UTOB, Ent1, Paramsoc, CalCulContr;

procedure InitApplication;
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  sysutils, DBTables,
  HEnt1, HCtrls,
  ULibCpContexte,
  eSession;

function CgiCtrCptaDispatch(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function CgiCtrCptaEdInit(Action, Param: string; RequestTOB: TOB): TOB;  forward;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
 ddwriteln('dispatch CGiCtrCpta................................');
 Action := UpperCase(Action);
 result := CgiCtrCptaDispatch(Action, Param, RequestTOB);
end;

function CgiCtrCptaDispatch(Action, Param: string; RequestTOB: TOB): TOB;
begin
    Result := CgiCtrCptaEdInit(Action, Param, RequestTOB)
end;

procedure InitApplication;
begin
  ddWriteLN('CgiCtrCpta loading...');
end;


function CgiCtrCptaEdInit(Action, Param : string ; RequestTOB : TOB) : TOB;
var
 ResponseTOB : TOB ;
begin
 ResponseTOB := TOB.Create('', nil, -1) ;
 if not ResponseTOB.FieldExists ('ERROR') then
    ResponseTOB.AddChampSupValeur ('ERROR', '');

 // A OPTIMISER : on devrait recharger systématiquement les paramsocs utilisés
 ChargeTOBSOC;

(* if V_GedFiles = nil then
 begin
    if V_PGI.RunFromLanceur then
      InitializeGedFiles(V_PGI.DefaultSectionDbName, nil)
    else
      InitializeGedFiles(V_PGI.DbName, nil);
 end;
*)
 if Action = 'CALCTREPARTIE' then
  begin
    ddwriteln('dispatch CALCTREPARTIE................................');
    LanceTraitementContreparties(RequestTOB);
  end;
  Result := ResponseTOB;

end;



end.


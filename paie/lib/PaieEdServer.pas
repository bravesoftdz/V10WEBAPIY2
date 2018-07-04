unit PaieEdServer;

interface

uses classes, UTOB,
  BullEdCalc,
  uLibPGContexte
;

procedure InitApplication;
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  sysutils, DBTables,
  HEnt1, HCtrls
  ,BullEnt
  ;


function EditionPaieDispatch(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function EditionPaieEdInit(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function EditionPaieEdClose(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function CalcOleEditionPaie(sf, sp: string): variant; forward;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
  ddWriteLN('PGIEditionPaie : DISPATCH');
  Action := UpperCase(Action);
  result := EditionPaieDispatch(Action, Param, RequestTOB);
end;

function EditionPaieDispatch(Action, Param: string; RequestTOB: TOB): TOB;
begin
  Result := nil;
  if Action = 'PAIEEDOPEN' then
    Result := EditionPaieEdInit(Action, Param, RequestTOB)
  else if Action = 'PAIEEDCLOSE' then
    Result := EditionPaieEdClose(Action, Param, RequestTOB)
end;

procedure InitApplication;
begin
  ddWriteLN('PGIEditionPaieAPI loading...');
end;

function EditionPaieEdInit(Action, Param: string; RequestTOB: TOB): TOB;
var
  ResponseTOB: TOB;
begin
  RegisterClassContext(TEditionPaieContext) ;
  TEditionPaieContext.GetCurrent.MySession.UserCalcEdt := CalcOleEditionPaie;

  ddWriteLN('PGIEditionPaieAPI doing ' + Action + ', ' + Param + ', appel#' +
    inttostr(TEditionPaieContext.GetCurrent.Count));
  ResponseTOB := TOB.Create('le test', nil, -1);
  ResponseTOB.AddChampSupValeur('RESULT', True);
  Result := ResponseTOB;

  if Assigned(Result) then
    ddWriteLN('PGIEditionPaieAPI Result : ' + IntToStr(Result.ChampsSup.Count)
      + ' champs, ' +
      IntToStr(Result.Detail.Count) + ' filles')
  else
    ddWriteLN('PGIEditionPaieAPI Result = nil')
end;

function EditionPaieEdClose(Action, Param: string; RequestTOB: TOB): TOB;
var
  ResponseTOB: TOB;
begin
  ResponseTOB := TOB.Create('le test', nil, -1);
  ResponseTOB.AddChampSupValeur('RESULT', True);
  Result := ResponseTOB;

  TEditionPaieContext.Release;
end;

function CalcOleEditionPaie(sf, sp: string): variant;
begin
  ddWriteLN('UserCalcEdt = ' + sf + '(' + sp + ')');
  Result := CalcOLEEtatBull(sf, sp);
end;

end.


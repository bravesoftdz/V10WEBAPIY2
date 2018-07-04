unit AmEdServer;

interface

uses classes, UTOB;

procedure InitApplication;
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  sysutils, DBTables,
  HEnt1, HCtrls,
  eSession,
  { AMORTISSEMENT}
  Ent1,
  imPlanInfo,
  ImEdCalc
  ;

function AmortissementDispatch(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function AmortissementEdInit(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function AmortissementEdClose(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function CalcOleAmortissement(sf, sp: string): variant; forward;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
  Action := UpperCase(Action);
  result := AmortissementDispatch(Action, Param, RequestTOB);
end;

function AmortissementDispatch(Action, Param: string; RequestTOB: TOB): TOB;
begin
  Result := nil;
  if Action = 'AMEDOPEN' then
    Result := AmortissementEdInit(Action, Param, RequestTOB)
  else if Action = 'AMEDCLOSE' then
    Result := AmortissementEdClose(Action, Param, RequestTOB)
end;

procedure InitApplication;
begin
  ddWriteLN('PGIAmortissementAPI loading...');
end;

type
  TAmortissementContext = class
  public { Déclarations publiques}
  protected { Déclarations protégées}
    Count: integer;
    PlanInfo: TPlanInfo;
    Exercices: array[1..20] of TExoDate;
  private { Déclarations privées}
    constructor Create;
    destructor Destroy; override;
    procedure ChargeInfoExo;
  end;

function AmortissementEdInit(Action, Param: string; RequestTOB: TOB): TOB;
var
  ResponseTOB: TOB;
  i: integer;

  pSession: TISession;
  pContext: TAmortissementContext;


begin
  pSession := LookupCurrentSession;
  i := pSession.UserObjects.IndexOf('Amortissement');
  ddWriteLN('PGIAmortissementAPI index#' + inttostr(i));
  if (i >= 0) then
  begin
    pContext := TAmortissementContext(pSession.UserObjects.Objects[i]);
  end
  else
  begin
    pContext := TAmortissementContext.Create;
    pSession.UserObjects.AddObject('Amortissement', pContext);
  end;
  Inc(pContext.Count);

  pSession.UserCalcEdt := CalcOleAmortissement;

  ddWriteLN('PGIAmortissementAPI doing ' + Action + ', ' + Param + ', appel#' +
    inttostr(pContext.Count));
  ResponseTOB := TOB.Create('le test', nil, -1);
  ResponseTOB.AddChampSupValeur('RESULT', True);
  Result := ResponseTOB;

  if Assigned(Result) then
    ddWriteLN('PGIAmortissementAPI Result : ' + IntToStr(Result.ChampsSup.Count)
      + ' champs, ' +
      IntToStr(Result.Detail.Count) + ' filles')
  else
    ddWriteLN('PGIAmortissementAPI Result = nil')
end;

function AmortissementEdClose(Action, Param: string; RequestTOB: TOB): TOB;
var
  i: integer;
  pSession: TISession;
  ResponseTOB: TOB;
begin
  pSession := LookupCurrentSession;

  i := pSession.UserObjects.IndexOf('Amortissement');
  ddWriteLN('PGIAmortissementAPI index#' + inttostr(i));
  if (i >= 0) then
  begin
    TAmortissementContext(pSession.UserObjects.Objects[i]).PlanInfo.Free;
    TAmortissementContext(pSession.UserObjects.Objects[i]).PlanInfo := nil;
    pSession.UserObjects.Delete(i);
  end;
  ResponseTOB := TOB.Create('le test', nil, -1);
  ResponseTOB.AddChampSupValeur('RESULT', True);
  Result := ResponseTOB;
end;

function CalcOleAmortissement(sf, sp: string): variant;
var
  pSession: TISession;
  pContext: TAmortissementContext;
  i: integer;
begin
  pSession := LookupCurrentSession;
  i := pSession.UserObjects.IndexOf('Amortissement');
  if i < 0 then
    exit;

  pContext := TAmortissementContext(pSession.UserObjects.Objects[i]);

  ddWriteLN('UserCalcEdt = ' + sf + '(' + sp + ')');
  Result := CalcOLEEtatImmo(pContext.PlanInfo, sf, sp);
end;

{ TAmortissementContext }

procedure TAmortissementContext.ChargeInfoExo;
var
  ll: integer;
  i, j, k: Word;
  Q: TQuery;
begin
  FillChar(Exercices, SizeOF(Exercices), #0);
  { Chargement de la liste des exercices }
  Q := OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT', TRUE, -1,
    'EXERCICE');
  if not Q.Eof then
  begin
    Q.First;
    ll := 1;
    while ((not Q.Eof) and (ll <= 20)) do
    begin
      Exercices[ll].Code := Q.FindField('EX_EXERCICE').AsString;
      Exercices[ll].Deb := Q.FindField('EX_DATEDEBUT').AsDateTime;
      Exercices[ll].Fin := Q.FindField('EX_DATEFIN').AsDateTime;
      Exercices[ll].DateButoir := Q.FindField('EX_DATECUM').AsDateTime;
      Exercices[ll].DateButoirRub := Q.FindField('EX_DATECUMRUB').AsDateTime;
      Exercices[ll].DateButoirBud := Q.FindField('EX_DATECUMBUD').AsDateTime;
      Exercices[ll].DateButoirBudgete :=
        Q.FindField('EX_DATECUMBUDGET').AsDateTime;
      NombrePerExo(Exercices[ll], i, j, k);
      Exercices[ll].NombrePeriode := k;
      Exercices[ll].EtatCpta := Q.FindField('EX_ETATCPTA').AsString;
      Inc(ll);
      Q.Next;
    end;
    Exercices[ll].Code := '';
    Exercices[ll].Deb := iDate1900;
    Exercices[ll].Fin := iDate1900;
  end
  else
  begin
    Exercices[1].Code := '';
    Exercices[1].Deb := iDate1900;
    Exercices[1].Fin := iDate1900;
  end;
  Ferme(Q);
end;

constructor TAmortissementContext.Create;
begin
  PlanInfo := TPlanInfo.Create('');
  ChargeInfoExo;
  PlanInfo.SetExercicesServer(Exercices);
end;

destructor TAmortissementContext.Destroy;
begin
  PlanInfo.Free;
  inherited;
end;

end.


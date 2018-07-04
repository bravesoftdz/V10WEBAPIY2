unit AmEdServer;

interface

uses classes, UTOB,
  { AMORTISSEMENT}
  Ent1,
  // eSession,      // BTY 03/06
  //imPlanInfo,     // BTY 03/06
  ImEdCalc,
  uLibCpContexte
;

// BTY 03/06 Class remontée dans ImEnt
//type
//  TAmortissementContext = class (TCPContexte)
//  private { Déclarations privées}
//    FPlanInfo: TPlanInfo;
//    procedure ChargeInfoExo;
//    // BTY 02/06 Rattacher les paramètres de TVA au contexte d'édition
//    procedure ChargeInfoTVA;
//    // BTY 03/06 Rattacher les coefficients dégressifs au contexte d'édition
//    procedure ChargeCoeffDegressif;
//  protected { Déclarations protégées}
//
//    // BTY 02/06 FQ 17547   Montants faux en Web Access
//    // => Démarrer à l'indice 0 pour pouvoir passer le tableau en paramètre
//    Exercices: array[0..20] of TExoDate; //array[1..20]
//    // BTY 02/06
//    TobTva: TOB;
//    // BTY 03/06 Pour les éditions si besoin dégressif
//    CoeffDegressif : TOB;
//    constructor Create( MySession : TISession ) ; override;
//    destructor Destroy; override;
//  public { Déclarations publiques}
//    property PlanInfo : TPlanInfo read FPlanInfo;
//  end;

procedure InitApplication;
function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;

implementation

uses
  sysutils, DBTables,
  HEnt1, HCtrls
  ,ImEnt     // BTY 03/06 Coeff dégressifs
  ;


function AmortissementDispatch(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function AmortissementEdInit(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function AmortissementEdClose(Action, Param: string; RequestTOB: TOB): TOB;  forward;
function CalcOleAmortissement(sf, sp: string): variant; forward;

function Dispatch(Action, Param: string; RequestTOB: TOB): TOB; stdcall;
begin
  ddWriteLN('PGIAMORTISSEMENT : DISPATCH');
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

function AmortissementEdInit(Action, Param: string; RequestTOB: TOB): TOB;
var
  ResponseTOB: TOB;

  (*
  pSession: TISession;
  pContext: TAmortissementContext;
   *)
begin
(*
  pSession := LookupCurrentSession;
  i := pSession.UserObjects.IndexOf('Amortissement');
  ddWriteLN('PGIAmortissementAPI index#' + inttostr(i));
  *)
(*
  if (i >= 0) then
  begin
    pContext := TAmortissementContext(pSession.UserObjects.Objects[i]);
  end
  else
  begin
    pContext := TAmortissementContext.Create;
    pSession.UserObjects.AddObject('Amortissement', pContext);
  end;
*)
  RegisterClassContext(TAmortissementContext) ;
  TAmortissementContext.GetCurrent.MySession.UserCalcEdt := CalcOleAmortissement;

//  pSession.UserCalcEdt := CalcOleAmortissement;

  ddWriteLN('PGIAmortissementAPI doing ' + Action + ', ' + Param + ', appel#' +
    inttostr(TAmortissementContext.GetCurrent.Count));
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
  ResponseTOB: TOB;
(*
  pSession: TISession;
  *)
begin
(*
  pSession := LookupCurrentSession;

  i := pSession.UserObjects.IndexOf('Amortissement');
  ddWriteLN('PGIAmortissementAPI index#' + inttostr(i));
  if (i >= 0) then
  begin
    TAmortissementContext(pSession.UserObjects.Objects[i]).PlanInfo.Free;
    TAmortissementContext(pSession.UserObjects.Objects[i]).PlanInfo := nil;
    pSession.UserObjects.Delete(i);
  end;
  *)
  ResponseTOB := TOB.Create('le test', nil, -1);
  ResponseTOB.AddChampSupValeur('RESULT', True);
  Result := ResponseTOB;

  TAmortissementContext.Release;
end;

function CalcOleAmortissement(sf, sp: string): variant;
begin
  ddWriteLN('UserCalcEdt = ' + sf + '(' + sp + ')');
  Result := CalcOLEEtatImmo(TAmortissementContext(TAmortissementContext.GetCurrent).PlanInfo, sf, sp);
end;

{ TAmortissementContext }

{procedure TAmortissementContext.ChargeInfoExo;
var
  ll: integer;
  i, j, k: Word;
  Q: TQuery;
begin
  FillChar(Exercices, SizeOF(Exercices), #0);
  // Chargement de la liste des exercices
  Q := OpenSQL('SELECT * FROM EXERCICE ORDER BY EX_DATEDEBUT', TRUE, -1,
    'EXERCICE');
  if not Q.Eof then
  begin
    Q.First;
    ll := 0; // 1;  BTY 02/06 FQ 17547
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
    // BTY 02/06 FQ 17547
    //Exercices[1].Code := '';
    //Exercices[1].Deb := iDate1900;
    //Exercices[1].Fin := iDate1900;
    Exercices[0].Code := '';
    Exercices[0].Deb := iDate1900;
    Exercices[0].Fin := iDate1900;
  end;
  Ferme(Q);
end;

// BTY 02/06 Rattacher les paramètres de TVA au contexte d'édition
procedure TAmortissementContext.ChargeInfoTVA;
begin
 TobTva := TOB.Create('',Nil,-1);
 TobTva.LoadDetailDB('TXCPTTVA','','',Nil,True)
end;

// BTY 03/06 Rattacher les coefficients dégressifs au contexte d'édition
procedure TAmortissementContext.ChargeCoeffDegressif;
begin
 if CoeffDegressif= nil then
    begin
    CoeffDegressif := TOB.Create('',Nil,-1);
    AmChargeCoefficientDegressif ( CoeffDegressif );
    end;
end;

constructor TAmortissementContext.Create ( MySession : TISession );
begin
  inherited  Create(MySession);
  FPlanInfo := TPlanInfo.Create('');
  ChargeInfoExo;
  FPlanInfo.SetExercicesServer(Exercices);
  // BTY 02/06 Charger la TVA dans le contexte d'édition
  ChargeInfoTVA;
  // BTY 02/06 Recopier la TVA dans PlanInfo qui est visible des formules de ImEdCalc
  FPlanInfo.SetTauxTvaServer(TobTva);
  // BTY 03/06 Charger les coefficients dégressifs
  ChargeCoeffDegressif;
  // BTY 03/06 Recopier la TOB dans PlanInfo qui est visible des formules de ImEdCalc
  FPlanInfo.SetCoeffDegressifServer(CoeffDegressif);
end;

destructor TAmortissementContext.Destroy;
begin
  FPlanInfo.Free;
  // BTY 02/06
  TobTVA.Free;
  // BTY 03/06
  CoeffDegressif.Free;
  inherited;
end; }

end.


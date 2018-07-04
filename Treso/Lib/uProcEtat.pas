unit uProcEtat;

interface

uses
  Classes, uTob
  {$IFDEF EAGLSERVER}
  , uLibCpContexte, eSession
  {$ELSE}
  , uEdtComp  
  {$ENDIF}
  ;

{$IFDEF EAGLSERVER}
type
  TEtatContext = class (TCPContexte)

  protected { Déclarations protégées}
    constructor Create( MySession : TISession ) ; override;
  public { Déclarations publiques}
    TobParamSoc : TOB;
    destructor Destroy; override;
  end;
{$ENDIF EAGLSERVER}

procedure GereDllEtat(ChargeOk : Boolean);


{05/10/06 : Fonctions pour les états}
function TrCalcOLEEtat(sf, sp : string): Variant;
function TrGetParamSocEtat(NoDossier, Chp : string) : string;
function CalcArobaseEtat (sf, sp : string) : Variant;

implementation

uses
  {$IFDEF EAGLCLIENT}
  uLanceProcess, uHTTP, HMsgBox,
  {$ENDIF EAGLCLIENT}
  sysutils, HEnt1, HCtrls, ParamSoc, Commun;

function ReadParam(var St : string): string; forward;

var
  NumLigne  : Integer = 0;

{---------------------------------------------------------------------------------------}
procedure GereDllEtat(ChargeOk : Boolean);
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLCLIENT}
var
  T, TResponse : TOB;
{$ENDIF EAGLCLIENT}
begin
  {$IFDEF EAGLCLIENT}
  T := TOB.Create ('',nil,-1);
  try
    if ChargeOk then
      TResponse := AppServer.Request('PGITresorerieAPI.GETPARAMSOCETAT', '', T, '', '')
    else
      TResponse := AppServer.Request('PGITresorerieAPI.FreeContext', '', T, '', '');
    if TResponse = nil then PGIBox('PGITresorerieAPI : serveur d''éditions introuvable');
    TResponse.Free;
    Delay(100); // on laisse le temps au serveur de se fermer
  finally
    T.Free;
  end;
  {$ELSE}

   if ChargeOk then //ProcCalcEdt := TrCalcOLEEtat
               else (*ProcCalcEdt := CalcAroBaseEtat*);
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
function CalcArobaseEtat (sf, sp : string) : Variant ;
{---------------------------------------------------------------------------------------}
begin
  Result := '' ;

  if sf = 'DEBUTPAGE' then
    NumLigne:=0
  else if sf='LIGNEBLANCHE' then begin
   Inc(NumLigne) ;
   if Not Odd(NumLigne) then Result:='OK' ;
  end
  else if sf='LIGNEGRISE' then begin
    if Odd(NumLigne) then Result:='OK' ;
  end;
end;

{---------------------------------------------------------------------------------------}
function ReadParam(var St : string): string;
{---------------------------------------------------------------------------------------}
var
  i: Integer;
begin
  i := Pos(',', St);
  if i <= 0 then
    i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;

{05/10/06 : Fonctions pour les états}
{---------------------------------------------------------------------------------------}
function TrCalcOLEEtat(sf, sp : string): Variant;
{---------------------------------------------------------------------------------------}
var
  s1 : string;
  s2 : string;
begin
  if sf = 'GETPARAMSOCETAT' then begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    Result := TrGetParamSocEtat(s1, s2);
  end
  else if sf = 'AAAAAAA' then begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    Result := '';
  end;
end;

{---------------------------------------------------------------------------------------}
function TrGetParamSocEtat(NoDossier, Chp : string) : string;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLSERVER}
  T       : TOB;
  {$ENDIF EAGLSERVER}
  NomBase : string;
begin
  Result := '';
  {$IFDEF EAGLSERVER}
  T := TEtatContext(TEtatContext.GetCurrent).TobParamSoc.FindFirst(['NODOSSIER', 'CHAMP'], [NoDossier, Chp], True);
  if Assigned(T) then
    Result := T.GetString('VALEUR')
  else begin
    T := Tob.Create('YYYY', TEtatContext(TEtatContext.GetCurrent).TobParamSoc, -1);
    T.AddChampSupValeur('NODOSSIER', NoDossier);
    T.AddChampSupValeur('CHAMP', Chp);
    NomBase := GetInfosFromDossier('DOS_NODOSSIER', NoDossier, 'DOS_NOMBASE');
    Result := GetParamsocDossierSecur(Chp, '', NomBase);
    T.AddChampSupValeur('VALEUR', Result);
  end;
  {$ELSE}
  NomBase := GetInfosFromDossier('DOS_NODOSSIER', NoDossier, 'DOS_NOMBASE');
  Result := GetParamsocDossierSecur(Chp, '', NomBase);
  {$ENDIF EAGLSERVER}
end;

{$IFDEF EAGLSERVER}
{---------------------------------------------------------------------------------------}
constructor TEtatContext.Create(MySession : TISession);
{---------------------------------------------------------------------------------------}
begin
  inherited  Create(MySession);
  TobParamSoc := Tob.Create('PARA', nil, -1);
end;

{---------------------------------------------------------------------------------------}
destructor TEtatContext.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobParamSoc) then FreeAndNil(TobParamSoc);
  inherited;
end;
{$ENDIF EAGLSERVER}

end.




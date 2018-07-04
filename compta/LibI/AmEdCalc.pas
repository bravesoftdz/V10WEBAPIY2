unit AmEdCalc;

interface

uses { VCL }
  SysUtils,
  { AGL }
  Hctrls, HEnt1,
  { Comptabilité }
  Ent1,
  { Amortissement }
  ImPlanInfo;

{$IFDEF EAGLSERVER}
function AmCalcOLEEtat(Exercices: array of TExoDate; PlanInfo: TPlanInfo; sf,
  sp: string): variant;
{$ELSE}
function AmCalcOLEEtat(sf, sp: string): variant;
{$ENDIF}

implementation

procedure UpdatePlanInfo(Exercices: array of TExoDate; PlanInfo: TPlanInfo;
  CodeImmo: string; DateRef: TDateTime);
begin
  ddwriteln('UpdatePlanInfo : ' + CodeImmo);
  if (PlanInfo = nil) then
  begin
    PlanInfo := TPlanInfo.Create('');
    PlanInfo.SetExercicesServer(Exercices);
    ddwriteln('PlanInfo = nil');
  end;
  if (PlanInfo.Plan.CodeImmo <> CodeImmo) or (PlanInfo.DateRef <> DateRef) then
  begin
    ddwriteln('PlanInfo = Avant Charge');
    PlanInfo.ChargeImmo(CodeImmo);
    ddwriteln('PlanInfo = Avant Calcul');
    PlanInfo.Calcul(DateRef, true);
    ddwriteln('PlanInfo = Après Calcul');
  end;
end;
(*
function GetMontantAchat(PlanInfo : TPlanInfo; CodeImmo:string; DateCalcul: TDateTime;EstImmoCedee: string): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo, DateCalcul);
  Result := PlanInfo.GetValeurAchat(DateCalcul,EstImmoCedee='CES');
end;

function GetCessionEco(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  Result:=PlanInfo.CessionEco ;
end;

function GetCessionFisc(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result:=PlanInfo.CessionFisc ;
end;
*)

function GetDotationEco(Exercices: array of TExodate; PlanInfo: TPlanInfo;
  CodeImmo: string; DateCalcul: TDateTime; EstImmoCedee: string): double;
begin
  UpdatePlanInfo(Exercices, PlanInfo, CodeImmo, DateCalcul);
  result := PlanInfo.DotationEco;
  //  if EstImmoCedee='CES' then result:=GetCessionEco(PlanInfo, CodeImmo,DateCalcul)+PlanInfo.ExcepEco ;
end;
(*
function GetDotationFisc(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime;EstImmoCedee: string): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result:=PlanInfo.DotationFisc;
  if EstImmoCedee='CES' then result:=GetCessionFisc(PlanInfo, CodeImmo,DateCalcul) ;
end;

function GetCumulAntEco(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime;EstImmoCedee: string): double ;
begin
  UpdatePlanInfo( PlanInfo, CodeImmo,DateCalcul);
  Result:=PlanInfo.CumulAntEco ;
  if EstImmoCedee='CES' then Result := PlanInfo.GetCumulAntEco( DateCalcul);
end;

function GetCumulAntFisc(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime;EstImmoCedee: string): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result:=PlanInfo.CumulAntFisc ;
  if EstImmoCedee='CES' then Result := PlanInfo.GetCumulAntFisc( DateCalcul);
end;

function GetCession(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime): string ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  Result:=PlanInfo.GetTypeSortie ( DateCalcul );
end;

function GetBaseEco ( PlanInfo : TPlanInfo; CodeImmo : string; DateCalcul : TDateTime ) : double;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result := PlanInfo.BaseEco;
end;

function GetExceptionnel(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result:=PlanInfo.ExcepEco ;
end;

function GetNonDeductible(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  result:=PlanInfo.NonDeductible;
end;

function GetVncEco(PlanInfo : TPlanInfo; CodeImmo: string ; DateCalcul: TDateTime;EstImmoCedee: string): double ;
begin
  UpdatePlanInfo(PlanInfo, CodeImmo,DateCalcul);
  Result:=PlanInfo.VNCEco ;
  if Result<0 then result:=0 ;
end;
  *)
{$IFDEF EAGLSERVER}

function AmCalcOLEEtat(Exercices: array of TExoDate; PlanInfo: TPlanInfo; sf,
  sp: string): variant;
var
{$ELSE}

function AmCalcOLEEtat(sf, sp: string): variant;
var
  PlanInfo: TPlanInfo;
{$ENDIF}
  s1, s2, s3: string;
begin
{$IFDEF EAGLSERVER}
{$ELSE}
  PlanInfo := VHImmo^.PlanInfo;
{$ENDIF}
  if sf = 'GETDOTATIONECO' then
  begin
    s1 := ReadTokenSt(sp);
    s2 := ReadTokenSt(sp);
    s3 := ReadTokenSt(sp);
    Result := GetDotationEco(Exercices, PlanInfo, s1, StrToFloat(S2), s3);
  end
  else
    Result := '';
  (*
    if sf = 'GETMONTANTACHAT' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      Result := GetMontantAchat (PlanInfo, s1, StrToFloat(s2), s3 );
    end else if sf='GETCESSION' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      result:=GetCession(PlanInfo, s1,StrToFloat(S2)) ;
    end else if sf='GETBASEECO' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      result:=GetBaseEco(PlanInfo, s1,StrToFloat(S2)) ;
    end else if sf='GETEXCEPTIONNEL' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      result:=GetExceptionnel(PlanInfo, s1,StrToFloat(S2)) ;
    end else if sf='GETNONDEDUCTIBLE' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      result:=GetNonDeductible(PlanInfo, s1,StrToFloat(S2)) ;
    end else if sf='GETDOTATIONECO' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      result:=GetDotationEco(PlanInfo, s1,StrToFloat(S2),s3) ;
    end else if sf='GETDOTATIONFISC' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      result:=GetDotationFisc(PlanInfo, s1,StrToFloat(S2),s3) ;
    end else if sf='GETCUMULANTECO' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      result:=GetCumulAntEco(PlanInfo, s1,StrToFloat(S2),s3) ;
    end else if sf='GETCUMULANTFISC' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      result:=GetCumulAntFisc(PlanInfo, s1,StrToFloat(S2),s3) ;
    end else if sf='GETVNCECO' then
    begin
      s1:=ReadTokenSt(sp) ;
      s2:=ReadTokenSt(sp) ;
      s3:=ReadTokenSt(sp) ;
      result:=GetVncEco(PlanInfo, s1,StrToFloat(S2),s3) ;
    end;
    *)
end;

end.


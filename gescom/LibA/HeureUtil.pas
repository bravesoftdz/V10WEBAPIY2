unit HeureUtil;
 // mcd 28/03/03 mise fct HeureBase100to60 dans CalcOleGenricAff pour @ ds les états
interface

Uses SysUtils, HMsgBox, HEnt1, Windows,DateUtils,

{$IFDEF BTP}
CalcOleGenericBTP;
{$ENDIF}

Function TimeToFloat      ( THeure : TDateTime ; b100:Boolean =False): Double;
Function FloatToTime      ( Heure : Double  ; b100:Boolean =False) : TDateTime;
Function HeureBase60To100 ( Heure : Double) : Double;
Function StrTimeIsNull    ( StrTime : string ) : Boolean;
Function StrTimeToFloat   ( StrTime : string ; b100:Boolean =False) : Double;
Function FloatToStrTime   ( Heure : Double; Format : String; b100:Boolean =False ) : String;

Function CalculEcartHeure ( HeureDeb, HeureFin : Double) : Double;
Function AjouteHeure      ( Heure1 , Heure2 : Double) : Double;
function TestHeure        ( heure : double) : Boolean;
function  CurrentDate: TDateTime;

Const formatHM  = 'hh:nn';

implementation
uses PlanUtil;

function TimeToFloat (THeure:TDateTime ; b100: Boolean =False) : Double;
Var Hour, Min, Sec, MSec : Word;
    //Op  : integer;
BEGIN
DecodeTime(THeure,Hour,Min,Sec,MSec);
Result := Hour+(Min/100);
if b100 then Result := HeureBase60To100(Result);
END;

function FloatToTime (Heure:Double; b100:Boolean =False) : TDateTime;
Var Hour, Min : word;
    tmp : double;
    //Op  : integer;
BEGIN
if b100 then Heure := BTPHeureBase100To60(Heure);
tmp := (Frac(Heure)*100);
if (tmp = 0) then Min:=0 else Min:=Trunc(Arrondi(tmp,0));
Hour := Trunc(Heure);
Result:= EncodeTime(Hour,Min,0,0);
END;


Function HeureBase60To100 ( Heure : Double) : Double;
var tmp : double;
BEGIN
if Heure = 0 then BEGIN Result := 0; Exit; END;
tmp := (Frac(Heure)/0.6);
if (tmp = 0) then Result := Trunc(Heure) else Result :=Trunc(Heure) + Arrondi(tmp,4);
END;

Function FloatToStrTime( Heure : Double ; Format : String; b100:Boolean =False ) : String ;
Var THeure : TDateTime;
BEGIN
if Heure > 24.0 then Heure := 0;
THeure := FloatToTime(Heure, b100); Result:= FormatDateTime(Format,THeure);
end;

Function StrTimeToFloat ( StrTime : string ; b100 : Boolean =False) : Double ;
Const Num = ['0'..'9'] ;
Var Hour , Min, Position : Integer;
    St : String;
    tmp, Op : double;

BEGIN
Result := 0.0;
if StrTime = '' then Exit;
if b100 then Op := 60.0 else Op := 100.0;
Position := Pos(':',StrTime);
if Position = 0 then Position := Pos('.',StrTime);
if Position = 0 then Position := Pos(',',StrTime);
if (Position <> 0 ) then
    BEGIN
    St:=Copy(StrTime,1,Position-1);
    if (St<> '')  then
        BEGIN
 //       if (Not (St[1] in Num)) then St[1] := '0';
 //       if (Not (St[2] in Num)) then St[2] := ' ';
        st := trim(st); Hour:=StrToInt(St);
        END else Hour :=0;
    St:=Copy(StrTime,Position+1,Length(StrTime)-Position);
    if (St<> '')  then
        BEGIN
        st := trim(st);
        Min :=StrToInt(St);
        END else Min:=0;
    tmp := Min / Op;
    if (Hour + tmp= 0)  then Result := 0
                        else Result := Arrondi ( Hour + tmp , 4 );
    END;
END;

Function StrTimeIsNull ( StrTime : string ) : Boolean ;
Var Hour,Min:Integer;
    St:String;
BEGIN
St:=Trim(Copy(StrTime,1,2)); if (St<> '') and IsNumeric(St) then Hour:=StrToInt(St) else Hour :=0;
St:=Trim(Copy(StrTime,4,2)); if (St<> '') and IsNumeric(St) then Min :=StrToInt(St) else Min:=0;
if (Hour=0) and (Min=0) then result :=True
                        else Result:=False;
End;


function TestHeure (Heure : double) : Boolean;
var d : double;
BEGIN
result := True;
d := Int (Heure);
if (d < 0) or (d >= 24) then result := False;
if (Result) and ( frac(Heure)>0.6 ) then result := False;
if (Result = False) then PGIBox ('Horaire invalide',TitreHalley);
END;

Function CalculEcartHeure ( HeureDeb , HeureFin : Double) : Double;
Var DTDeb, DTFin : TDateTime;
BEGIN
DTDeb := FloatToTime(HeureDeb); DTFin := FloatToTime(HeureFin);
if (HeureFin < HeureDeb) then
    BEGIN
    // On considère qu'il s'agit d'un changement de jour
    Result := TimeToFloat ( 1 - DTDeb + DTFin);
    END
else
    BEGIN
    Result := TimeToFloat ( DTFin - DTDeb);
    END;
if (Result < 0) Then result :=0;
END;

Function AjouteHeure ( Heure1 , Heure2 : Double) : Double;
//Var DTDeb, DTFin : TDateTime;
BEGIN
if Heure1 = 0 then BEGIN Result := Heure2; Exit; END;
if Heure2 = 0 then BEGIN Result := Heure1; Exit; END;

Result := TimeToFloat( FloatToTime (Heure1) + FloatToTime (Heure2));
END;

function CurrentDate: TDateTime;
var
  SystemTime: TSystemTime;
begin
  GetLocalTime(SystemTime);
  Result := SystemTimeToDateTime(SystemTime);
end;


end.

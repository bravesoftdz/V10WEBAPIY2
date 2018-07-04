unit DMUtil;

interface

uses SysUtils;

function extraitChp(const S: string; var Chp: string; var S2: string; Car: char): boolean;
function genenreg(N: integer; Enreg: array of string; Car: char): string;

implementation

// donne le 1er champ et le reste du string

function extraitChp(const S: string; var Chp: string; var S2: string; Car: char): boolean;
var K: integer;
begin
  if S = ''
    then Result := false
  else
  begin
    K := ansipos(Car, S);
    if K = 0
      then
    begin
      Chp := S;
      S2 := '';
    end
    else
    begin
      Chp := copy(S, 0, K - 1);
      S2 := copy(S, K + 1, length(S)); // le reste
    end;
    Result := true;
  end;
end;

// concatène dans un string les champ de données après les champ clés

function genenreg(N: integer; Enreg: array of string; Car: char): string;
var S: string;
begin
  if N > high(Enreg)
    then Result := ''
  else
  begin
    S := genEnreg(N + 1, Enreg, #1);
    {if S=''
      then Result:= Enreg[N]
      else } Result := Enreg[N] + Car + S;
  end;
end;

end.

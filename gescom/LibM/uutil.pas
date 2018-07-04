unit UUtil;

interface

uses SysUtils, Windows, Classes, Graphics, Controls;

type

  PMaListeInt = ^MaListeInt;
  MaListeInt = record
    donnee: longint;
    suivant: PMaListeInt;
  end;

  PListePotCal = ^TListePotCal;
  TListePotCal = record
    Matricule: string;
    Datej: string;
    Valeur: string;
  end;

  PListeParam = ^TListeParam;
  TListeParam = record
    TypeParam: string;
    Niveau: integer;
    Numero: double;
    Libelle: string;
    Variable: integer;
    Cle: integer;
    TypeVal: string;
    ListeVal: TStringList;
    Valeur: string;
  end;

function pluspetit(A, B: integer): integer;
function plusgrand(A, B: integer): integer;
function pluspetitr(A, B: double): double;
function plusgrandr(A, B: double): double;
function liste(const L: array of string): TStringList;

procedure addListeInt(var Liste: PMaListeInt; Val: longint);
procedure destroyListeInt(var Liste: PMaListeInt);

function rgb125(Code: integer): TColor;
function invrgb125(Couleur: TColor): integer;
function indrgb125(Couleur: TColor): integer;
function decale(X: integer): integer;

function retireblanc(const Texte: string): string;

{$IFDEF TRACE}
function DonneTemps(): TdateTime;
function DonneTemps2(): string;
function DonneTemps3(): TdateTime;
{$ENDIF}

function TransformeCodeSqlSansGuill(codesql: string): string;
function TransformeCodeSqlAvecGuill(codesql: string): string;

function MetZero(const S: string; Lg: integer): string;

var
  //    listeCal: Tlist;
  tps: TDateTime;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Céline
Créé le ...... : 16/07/2001
Modifié le ... :   /  /
Description .. : Ajoute des zéros en début de la chaine donnée pour arriver
Suite ........ : à la longueur Lg donnée
Mots clefs ... : AJOUTE ZERO
*****************************************************************}

function MetZero(const S: string; Lg: integer): string;
var x, i: integer;
  res: string;
begin
  res := S;
  x := Lg - length(S);
  for i := 0 to x - 1 do
  begin
    res := '0' + res;
  end;
  result := res;
end;

//enlève dans la chaine représentant une requête
//les "
//pour pouvoir enregistrer les filtres

function TransformeCodeSqlSansGuill(codesql: string): string;
var i: integer;
begin
  for i := 1 to Length(codesql) do
  begin
    if codesql[i] = '"'
      then codesql[i] := 'µ';
  end;
  result := codesql;
end;

//remet les " dans la chaine représentant une requête

function TransformeCodeSqlAvecGuill(codesql: string): string;
var i: integer;
begin
  for i := 1 to Length(codesql) do
  begin
    if codesql[i] = 'µ'
      then codesql[i] := '"';
  end;
  result := codesql;
end;

{$IFDEF TRACE}

function DonneTemps(): TdateTime;
var x: TDateTime;
begin
  x := time;
  result := x - tps;
  tps := x;
end;

function DonneTemps2(): string;
var x, y: TDateTime;
  h, m, s, ms: word;
begin
  x := time;
  y := x - tps;
  Decodetime(y, h, m, s, ms);
  result := inttostr(h) + ' : ' + inttostr(m) + ' : ' + inttostr(s) + ' : ' + inttostr(ms);
  tps := x;
end;

function DonneTemps3(): tdatetime;
var x: TDateTime;
begin
  x := time;
  result := x - tps;
  tps := x;
end;

{$ENDIF}

function liste(const L: array of string): TstringList;
var
  P: TStringList;
  I: integer;
begin
  P := TStringList.create;
  for I := 0 to high(L) do
    P.add(L[I]);
  Result := P;
end;

procedure addListeInt(var Liste: PMaListeInt; Val: longint);
var
  PLPred, PLc, PL: PMalisteInt;
begin
  new(PL);
  PL^.donnee := Val;
  PL^.suivant := nil;
  PLc := Liste;
  PLPRed := nil;
  if PLc = nil then Liste := PL
  else
  begin
    while PLc <> nil do
    begin
      PLPred := PLC;
      PLc := PLPred^.suivant;
    end;
    PLPred^.suivant := PL;
  end;
end;

procedure destroyListeInt(var Liste: PMaListeInt);
var
  PL, PLd: PMalisteInt;
  P: MaListeInt;
begin
  PL := Liste;
  while PL <> nil do
  begin
    P := PL^;
    PLd := PL;
    PL := P.suivant;
    dispose(PLd);
  end;
  Liste := nil;
end;

function pluspetit(A, B: integer): integer;
begin
  if A > B then Result := B else Result := A;
end;

function plusgrand(A, B: integer): integer;
begin
  if A > B then Result := A else Result := B;
end;

function invrgb125(Couleur: TColor): integer;
var
  L, P, M: integer;
  Color: TColor;
begin
  Result := 0;
  for P := 0 to 4 do
  begin
    for M := 0 to 24 do
    begin
      L := P * 25 + decale(M);
      Color := rgb125(L);
      if Couleur = Color then
        Result := L
    end;
  end;
end;

function indrgb125(Couleur: TColor): integer;
var
  L, P, M: integer;
  I: integer;
  Color: TColor;
begin
  I := 0;
  Result := 0;
  for P := 0 to 4 do
  begin
    for M := 0 to 24 do
    begin
      L := P * 25 + decale(M);
      Color := rgb125(L);
      if Couleur = Color then
        Result := I
      else
        I := I + 1;
    end;
  end;
end;

function decale(X: integer): integer;
begin
  case X of
    0: Result := 0;
    1: Result := 1;
    2: Result := 2;
    3: Result := 7;
    4: Result := 8;
    5: Result := 3;
    6: Result := 4;
    7: Result := 9;
    8: Result := 13;
    9: Result := 14;
    10: Result := 19;
    11: Result := 24;
    12: Result := 23;
    13: Result := 22;
    14: Result := 21;
    15: Result := 20;
    16: Result := 18;
    17: Result := 17;
    18: Result := 16;
    19: Result := 15;
    20: Result := 12;
    21: Result := 11;
    22: Result := 10;
    23: Result := 6;
    24: Result := 5;
  else
    Result := 0;
  end;
end;

function rgb125(Code: integer): TColor;
var
  I, J, K, bred, bgreen, bblue: integer;
begin
  K := Code mod 5;
  J := (Code div 5) mod 5;
  I := (Code div 25) mod 5;
  bRed := 64 * I;
  if bRed = 256 then bRed := 255;
  bGreen := 64 * J;
  if bGreen = 256 then bGreen := 255;
  bBlue := 64 * K;
  if bBlue = 256 then bBlue := 255;
  Result := TColor(RGB(bRed, bGreen, bBlue));
end;

function retireblanc(const Texte: string): string;
begin
  Result := TrimRight(Texte);
end;

function pluspetitr(A, B: double): double;
begin
  if A > B then Result := B else Result := A;
end;

function plusgrandr(A, B: double): double;
begin
  if A > B then Result := A else Result := B;
end;

end.

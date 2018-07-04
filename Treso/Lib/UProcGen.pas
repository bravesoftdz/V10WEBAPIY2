{Unité de manipulations génériques de chaines (cf : *CHAINES*),
                                      dates (cf : *DATES*),
                                      conversion de format  (cf : *TRANSTYPAGE*),
                                      Chiffres (cf : *CHIFFRES*)
                                      divers (cf : *DIVERS*),
                                      Montants ...}
unit UProcGen;

interface

uses SysUtils, HEnt1, HTB97, Classes, Controls, HCtrls;

                 { *TRANSTYPAGE* }
{Transforme un variant en entier}
function VarToInt(MonVar : Variant) : Integer;
{Transforme un variant en double}
function VarToDle(MonVar : Variant) : Double;
{Transforme une string en char si elle est de longueur 1, sinon renvoie #0}
function StrToChr(MaString : string): Char;
{Transorme une Chaine représentant un Double en string SQL}
function VarStrToPoint(MaString : string): string;

                 { *CHAINES* }
{Complète une chaine sur la gauche avec le car sur la longueur Long}
function PadL(Chaine : string; Car : Char; Long : Byte) : string;
{Complète une chaine sur la droite avec le car sur la longueur Long}
function PadR(Chaine : string; Car : Char; Long : Byte) : string;
{Renvoie les "Combien" derniers caractères d'une chaine}
function StrRight(Chaine : string; Combien : Integer) : string;
{Renvoie les "Combien" premiers caractères d'une chaine}
function StrLeft(Chaine : string; Combien : Integer) : string;
{Renvoie la position de la première occurence de SubStr en partant de la droite}
function RPos(SubStr, Chaine : string) : Integer;
{RPosAbs la position en partant de 0 à gauche contre 0 en partant de droite avec RPos}
function RPosAbs(SubStr, Chaine : string) : Integer;
{Transforme une chaine 'AA;BB;CC' en '"AA", "BB", "CC"' : notamment pour
 constituer une clause IN à partir d'un THMultiValComboBox}
function GetClauseIn(Chaine : string) : string;
{Cherche la valeur supérieure de Code, en précisant sa longeur}
function IncCode(var Code : string; Lg : Byte = 3) : Boolean;
{Supprime ADetruire de la chaine}
function DelChaine(ADetruire, AChaine : string) : string;
{Retourne le nombre d'occurences de SsCh dans Chaine}
function NbOccurences(SsCh, Chaine : string) : Integer;
{Correspond au FindEtReplace, à la nuance près que NewValue peut contenir OldValue sans faire une Boucle infinie}
function TrTrouveEtRemplace(const aChaine, OldValue, NewValue : HString; Tous : Boolean) : HString;

                 { *DATES* }
{Calcule la date de Pâques pour une année donnée}
function CalculPaques(Year : Integer) : TDateTime;
{Retourne le dernier jour du trimestre}
function FinTrimestre(Dt : TDate) : string;
{Retourne le premier jour du trimestre}
function DebutTrimestre(Dt : TDate) : string;
{Retourne le dernier jour du trimestre}
function FinDeTrimestre(Dt : TDate) : TDateTime;
{Retourne le premier jour du trimestre}
function DebutDeTrimestre(Dt : TDate) : TDateTime;
{Retourne le dernier jour du Semestre}
function FinSemestre(Dt : TDate) : string;
{Retourne le premier jour du Semestre}
function DebutSemestre(Dt : TDate) : string;
{Retourne le mois et l'année en lettre : dt est la date de départ et nb est le nombre de
 mois à rajouter à l'année}
function  RetourneMois(dt : string; nb : Integer) : string;
{Retourne le mois et l'année à partir d'une date}
function  DateToMois  (dt : string   ) : string; overload;
function  DateToMois  (dt : TDateTime) : string; overload;
{Retourne le Trimestre et l'année à partir d'une date}
function  DateToTrimestre(dt : TDateTime) : string;
{Retourne le premier (ou dernier si FinOk) jour du mois en fonction d'un mois et d'une année en lettre}
function  MoisToDate  (ch : string; FinOk : Boolean = False) : TDateTime;
{Calcul le nombre de mois entre deux dates}
function  GetNbMoisEntre(Date1, Date2 : TDateTime) : Integer;
function  RetourneAnneeW(Dt : TDateTime) : Word;
function  RetourneMoisW (Dt : TDateTime) : Word;
function  RetourneJourW (Dt : TDateTime) : Word;

                 { *CHIFFRES* }
{Arrondi un montant selon la condition}
function ArrondirTreso(Arrondi : string; Mnt: Double): Double;
{Regarde si "ATester" appartient à [Inf, Sup] }
function Between(ATester, Inf, Sup : Double) : Boolean; overload;
function Between(ATester, Inf, Sup : string) : Boolean; overload;
{21/10/04 : Fait un FormatFloat en fonction du nombre de décimales}
function FormateMontant(Mnt : Double; Dec : Integer) : string;
{03/01/05 : s'assure que Chiffre est différent de zéro ; Defaut est la valeur à renvoyer si c'était le cas}
function DoubleNotNul(Chiffre : Double; Defaut : Double = 1) : Double;

                 { *DIVERS* }
{Chargement dynamique d'un glyph}
procedure ChargeGlyph(bb : TToolbarButton97; Img : string = '');
{Libération des objets contenus dans une StringList}
procedure LibereListe(var L : TStringList; Detruire : Boolean);
{Libération de la mémoire allouée à des variables dynamiques contenues dans une TLIST}
procedure DisposeListe(var L : TList; Detruire : Boolean);
{Constitution d'un nom de fichier}
function GetNomFichier(Rep, Nom, Ext : string) : string;
{Copie d'un fichier : Fonction est le traitement appelant pour la génération
 du nom du fichier}
function CopierFichier(FicSource, RepDest, Fonction, Ext : string) : Boolean;
{Execute un RollBack en cachant le message en jaune sur le HPanel}
procedure RollBackDiscret;
{15/03/05 : Calcule le montant HT à partir du montant TTC et du Taux de TVA}
function GetMntHTFromTTC(MntTTC, TxTVA : Double; NbDec : Integer = -1) : Double;
{15/03/05 : Calcule le montant de la TVA à partir du montant TTC et du Taux de TVA}
function GetMntTVAFromTTC(MntTTC, TxTVA : Double; NbDec : Integer = -1) : Double;


implementation

uses
  HMsgBox, FileCtrl, Windows;

{Transforme un variant en entier
{---------------------------------------------------------------------------------------}
function VarToInt(MonVar : Variant) : Integer;
{---------------------------------------------------------------------------------------}
begin
  if TVarData(MonVar).VType = varNull then Result := 0
                                      else Result := TVarData(MonVar).VInteger;
end;

{Transforme un variant en entier
{---------------------------------------------------------------------------------------}
function VarToDle(MonVar : Variant) : Double;
{---------------------------------------------------------------------------------------}
begin
  if TVarData(MonVar).VType = varNull then Result := 0
                                      else Result := TVarData(MonVar).VDouble;
end;

{Transforme une string en char si elle est de longueur 1, sinon renvoie #0}
{---------------------------------------------------------------------------------------}
function StrToChr(MaString : string): Char;
{---------------------------------------------------------------------------------------}
begin
  Result := #0;
  if Length(MaString) = 1 then
    Result := MaString[1];
end;

{Transorme une Chaine représentant un Double en string SQL
{---------------------------------------------------------------------------------------}
function VarStrToPoint(MaString : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := StrFPoint(Valeur(MaString));
end;

{Complète une chaine sur la gauche avec le car sur la longueur Long
{---------------------------------------------------------------------------------------}
function PadL(Chaine : string; Car : Char; Long : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  if Car = '' then Car := ' ';
  Result := Chaine;
  while Length(Result) < Long do Result := Car + Result;
end;

{Complète une chaine sur la droite avec le car sur la longueur Long
{---------------------------------------------------------------------------------------}
function PadR(Chaine : string; Car : Char; Long : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  if Car = '' then Car := ' ';
  Result := Chaine;
  while Length(Result) < Long do Result := Result + Car;
end;

{Renvoie les n derniers caractères d'une chaine
{---------------------------------------------------------------------------------------}
function StrRight(Chaine : string; Combien : Integer) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  Result := Copy(chaine, Length(chaine) - Combien + 1, Combien);
end;

{Renvoie les n premiers caractères d'une chaine
{---------------------------------------------------------------------------------------}
function StrLeft(Chaine : string; Combien : Integer) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  Result := Copy(chaine, 1, Combien);
end;

{Renvoie la position de la première occurence de SubStr en partant de la droite
{---------------------------------------------------------------------------------------}
function RPos(SubStr, Chaine : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  Result := Length(Chaine);
  ch := Chaine;
  while Pos(SubStr, ch) > 0 do begin
    Result := Result - Pos(SubStr, ch);
    ReadTokenPipe(ch, SubStr);
  end;
end;

{RPosAbs la position en partant de 0 à gauche contre 0 en partant de droite avec RPos
{---------------------------------------------------------------------------------------}
function RPosAbs(SubStr, Chaine : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  Result := 0;
  ch := Chaine;
  while Pos(SubStr, ch) > 0 do begin
    Result := Result + Pos(SubStr, ch);
    ReadTokenPipe(ch, SubStr);
  end;
end;

{Transforme une chaine 'AA;BB;CC' en '"AA", "BB", "CC"' : notamment pour
 constituer une clause IN à partir d'un THMultiValComboBox
{---------------------------------------------------------------------------------------}
function GetClauseIn(Chaine : string) : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := '';
  {26/05/05 : on rajoute un ';' en fin de chaine s'il n'est pas présent, et si bien
              sûr la chaine n'est pas vide}
  if (chaine <> '') and (Chaine[Length(Chaine)] <> ';') then
    Chaine := Chaine + ';';

  while Pos(';', Chaine) > 0 do begin
    s := ReadTokenSt(Chaine);
    Result := Result + ' "' + s + '",';
  end;
  System.Delete(Result, Length(Result), 1);
end;

{Cherche la valeur supérieure de Code, en précisant sa longeur
{---------------------------------------------------------------------------------------}
function IncCode(var Code : string; Lg : Byte = 3) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Byte;
  c : Char;
  p : Char;
  s : string;
begin
  Result := True;
  {On met le code sur sa longueur standard}
  if Length(Code) < Lg then Padr(Code, '0', Lg - Length(Code));
  {On regarde si le code est de Type numérique}
  n := 1;
  while n <= Lg do begin
    if not (Code[n] in ['0'..'9']) then Break;
    Inc(n);
  end;

  {Si le code est de type numérique}
  if n = Lg + 1 then
    {On l'incrémente de 1}
    Code := Padl(IntToStr(StrToInt(Code) + 1), '0', Lg)
  else begin
    for n := Lg downto 1 do begin
      c := Code[n];
      p := Code[n];
      {On met le code en majuscule}
      if c in ['a'..'z'] then begin
        UpCase(c);
        UpCase(p);
      end;
      if c in ['0'..'9', 'A'..'Z'] then
        Inc(c)
      else begin
        {On ne génère pas de code avec des caractères spéciaux}
        Result := False;
        HShowMessage('0;Génération d''un code; Le code dépasse l''étendue alphanumérique.'#13 +
                     'Traitement intérrompu.;E;O;O;O;', '', '');
        Exit;
      end;

      if (p in ['0'..'9', 'A'..'Z']) and not (c in ['0'..'9', 'A'..'Z']) and (n = 1) then begin
        {Si on arrive au maximum des caractères alphanumérique, on avertit et on sort}
        Result := False;
        HShowMessage('0;Génération d''un code; Le code dépasse l''étendue alphanumérique.'#13 +
                     'Traitement intérrompu.;E;O;O;O;', '', '');
        Exit;
      end
      else if (p in ['A'..'Z']) and not (c in ['A'..'Z']) then c := 'A'
      else if (p in ['0'..'9']) and not (c in ['0'..'9']) then c := '0'
      else begin
        s := c + s;
        Code := Copy(Code, 1, n - 1) + s;
        Break;
      end;
      s := c + s;
    end;
  end;
end;

{Supprime ADetruire de la chaine
{---------------------------------------------------------------------------------------}
function DelChaine(ADetruire, AChaine : string) : string;
{---------------------------------------------------------------------------------------}
var
  n, p : Integer;
begin
  Result := AChaine;
  n := Pos(ADetruire, AChaine);
  if n > 0 then begin
    p := Length(ADetruire);
    System.Delete(Result, n, p);
  end;
end;

{Retourne le nombre d'occurences de SsCh dans Chaine}
{---------------------------------------------------------------------------------------}
function NbOccurences(SsCh, Chaine : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  S : string;
begin
  Result := 0;
  S := Chaine;
  while Pos(SsCh, S) <> 0 do begin
    Inc(Result);
    S := Copy(S, Pos(SsCh, S) + Length(SsCh), Length(S));
  end
end;

{Correspond au FindEtReplace, à la nuance près que NewValue peut contenir OldValue sans faire une Boucle infinie}
{---------------------------------------------------------------------------------------}
function TrTrouveEtRemplace(const aChaine, OldValue, NewValue : HString; Tous : Boolean) : HString;
{---------------------------------------------------------------------------------------}
var
  s : HString;
  c : HString;
begin
  if Pos(OldValue, aChaine) = 0 then begin
    Result := aChaine;
    Exit;
  end;

  Result := '';
  c := aChaine;
  s := ReadTokenPipe(c, OldValue);
  while s <> '' do begin
    Result := Result + s + NewValue;

    if not Tous or (Pos(OldValue, c) = 0) then begin
      Result := Result + c;
      Break;
    end;

    s := ReadTokenPipe(c, OldValue);
  end;
end;

{Calcule la date de Pâques pour une année donnée}
{---------------------------------------------------------------------------------------}
function CalculPaques(Year: Integer) : TDateTime;
{---------------------------------------------------------------------------------------}
var
  c, n, k,
  i, j, l,
  m, d   : Integer;
begin
  c := Year div 100;
  n := Year mod 19;
  k := (c - 17) div 25;
  i := (c - c div 4 - (c - k) div 3 + 19 * n + 15) mod 30;
  i := i - (i div 28) * (1 - (i div 28) * (29 div (i + 1)) * ((21 - n) div 11));
  j := (Year + Year div 4 + i + 2 - c + c div 4) mod 7;
  l := i - j;
  m := 3 + (l + 40) div 44;
  d := l + 28 - 31 * (m div 4);
  Result := EncodeDate(Year, m, d);
end;

{Retourne le premier jour du trimestre}
{---------------------------------------------------------------------------------------}
function DebutTrimestre(Dt : Tdate) : string ;
{---------------------------------------------------------------------------------------}
begin
  Result := DateToStr(DebutDeTrimestre(Dt));
end;

{Retourne le dernier jour du trimestre}
{---------------------------------------------------------------------------------------}
function FinTrimestre(Dt : Tdate) : string ;
{---------------------------------------------------------------------------------------}
begin
  Result := DateToStr(FinDeTrimestre(Dt));
end;

{Retourne le premier jour du trimestre}
{---------------------------------------------------------------------------------------}
function DebutDeTrimestre(Dt : Tdate) : TDateTime;
{---------------------------------------------------------------------------------------}
var
  AA, MM, JJ : Word ;
begin
  Result := Date;
  DecodeDate(Dt, AA, MM, JJ);
  case MM of
    1..3   : Result := DebutDeMois(EncodeDate(AA, 1, 1));
    4..6   : Result := DebutDeMois(EncodeDate(AA, 4, 1));
    7..9   : Result := DebutDeMois(EncodeDate(AA, 7, 1));
    10..12 : Result := DebutDeMois(EncodeDate(AA,10, 1));
  end;
end;

{Retourne le dernier jour du trimestre}
{---------------------------------------------------------------------------------------}
function FinDeTrimestre(Dt : Tdate) : TDateTime;
{---------------------------------------------------------------------------------------}
var
  AA, MM, JJ : Word ;
begin
  Result := Date;
  DecodeDate(Dt, AA, MM, JJ);
  case MM of
    1..3   : Result := FinDeMois(EncodeDate(AA, 3, JJ));
    4..6   : Result := FinDeMois(EncodeDate(AA, 6, JJ));
    7..9   : Result := FinDeMois(EncodeDate(AA, 9, JJ));
    10..12 : Result := FinDeMois(EncodeDate(AA,12, JJ));
  end ;
end;

{Retourne le dernier jour du Semestre}
{---------------------------------------------------------------------------------------}
function FinSemestre(Dt : TDate) : string ;
{---------------------------------------------------------------------------------------}
var
  AA, MM, JJ : Word ;
begin
  DecodeDate(Dt, AA, MM, JJ);
  case MM of
    1..6  : Result := DateToStr(FinDeMois(EncodeDate(AA, 6, JJ)));
    7..12 : Result := DateToStr(FinDeMois(EncodeDate(AA,12, JJ)));
  end;
end;

{Retourne le premier jour du Semestre}
{---------------------------------------------------------------------------------------}
function DebutSemestre(Dt : TDate) : string ;
{---------------------------------------------------------------------------------------}
var
  AA, MM, JJ : Word ;
begin
  DecodeDate(Dt, AA, MM, JJ);
  case MM of
    1..6  : Result := DateToStr(DebutDeMois(EncodeDate(AA, 1, JJ)));
    7..12 : Result := DateToStr(DebutDeMois(EncodeDate(AA, 7, JJ)));
  end;
end;

{Retourne le mois et l'année à partir d'une date
{---------------------------------------------------------------------------------------}
function DateToMois(dt : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := DateToMois(StrToDate(dt));
end;

{Retourne le mois et l'année à partir d'une date
{---------------------------------------------------------------------------------------}
function DateToMois(dt : TDateTime) : string;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(dt, a, m, j);
  case m of
    1  : Result := TraduireMemoire('Janvier ')   + IntToStr(a);
    2  : Result := TraduireMemoire('Fevrier ')   + IntToStr(a);
    3  : Result := TraduireMemoire('Mars ')      + IntToStr(a);
    4  : Result := TraduireMemoire('Avril ')     + IntToStr(a);
    5  : Result := TraduireMemoire('Mai ')       + IntToStr(a);
    6  : Result := TraduireMemoire('Juin ')      + IntToStr(a);
    7  : Result := TraduireMemoire('Juillet ')   + IntToStr(a);
    8  : Result := TraduireMemoire('Aout ')      + IntToStr(a);
    9  : Result := TraduireMemoire('Septembre ') + IntToStr(a);
    10 : Result := TraduireMemoire('Octobre ')   + IntToStr(a);
    11 : Result := TraduireMemoire('Novembre ')  + IntToStr(a);
    12 : Result := TraduireMemoire('Decembre ')  + IntToStr(a);
  end;

  Result := AnsiUpperCase(Result);
end;

{Retourne le Trimestre et l'année à partir d'une date}
{---------------------------------------------------------------------------------------}
function DateToTrimestre(dt : TDateTime) : string;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(dt, a, m, j);
  case m of
    1..3   : Result := TraduireMemoire('1er Trimestre ' ) + IntToStr(a);
    4..6   : Result := TraduireMemoire('2ème Trimestre ') + IntToStr(a);
    7..9   : Result := TraduireMemoire('3ème Trimestre ') + IntToStr(a);
    10..12 : Result := TraduireMemoire('4ème Trimestre ') + IntToStr(a);
  end;
end;

{---------------------------------------------------------------------------------------}
function  RetourneAnneeW(Dt : TDateTime) : Word;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(dt, a, m, j);
  Result := a;
end;

{---------------------------------------------------------------------------------------}
function  RetourneMoisW(Dt : TDateTime) : Word;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(dt, a, m, j);
  Result := m;
end;

{---------------------------------------------------------------------------------------}
function  RetourneJourW(Dt : TDateTime) : Word;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(dt, a, m, j);
  Result := j;
end;

{Retourne le premier (ou dernier si FinOk) jour du mois en fonction d'un mois et d'une année en lettre
{---------------------------------------------------------------------------------------}
function  MoisToDate(ch : string; FinOk : Boolean = False) : TDateTime;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  j := 1;
  m := 0;
  try
    a := StrToInt(StrRight(ch, 4));
         if Pos(TraduireMemoire('janvier ')   , AnsiLowerCase(ch)) > 0 then m := 1
    else if Pos(TraduireMemoire('fevrier ')   , AnsiLowerCase(ch)) > 0 then m := 2
    else if Pos(TraduireMemoire('mars ')      , AnsiLowerCase(ch)) > 0 then m := 3
    else if Pos(TraduireMemoire('avril ')     , AnsiLowerCase(ch)) > 0 then m := 4
    else if Pos(TraduireMemoire('mai ')       , AnsiLowerCase(ch)) > 0 then m := 5
    else if Pos(TraduireMemoire('juin ')      , AnsiLowerCase(ch)) > 0 then m := 6
    else if Pos(TraduireMemoire('juillet ')   , AnsiLowerCase(ch)) > 0 then m := 7
    else if Pos(TraduireMemoire('aout ')      , AnsiLowerCase(ch)) > 0 then m := 8
    else if Pos(TraduireMemoire('septembre ') , AnsiLowerCase(ch)) > 0 then m := 9
    else if Pos(TraduireMemoire('octobre ')   , AnsiLowerCase(ch)) > 0 then m := 10
    else if Pos(TraduireMemoire('novembre ')  , AnsiLowerCase(ch)) > 0 then m := 11
    else if Pos(TraduireMemoire('decembre ')  , AnsiLowerCase(ch)) > 0 then m := 12;

    if m = 0 then Result := Date
             else Result := EncodeDate(a, m, j);

    if FinOk then Result := FinDeMois(Result);{28/12/2004}
  except
    Result := Date;
  end;
end;

{Retourne le mois et l'année en lettre : dt est la date de départ et nb est le nombre de
 mois à rajouter à l'année
{---------------------------------------------------------------------------------------}
function RetourneMois(dt : string; nb : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  Mois    : Word;
  NbAnnee : Word;
begin
  DecodeDate(StrToDate(dt), a, m, j);
  Mois := (m + nb) mod 12;

  if (m + nb) > 12 then begin
    NbAnnee := (m + nb) div 12;
    if Mois = 0 then Inc(a, NbAnnee - 1)
                else Inc(a, NbAnnee);
  end;

  case Mois of
    0  : Result := 'Decembre '  + IntToStr(a);
    1  : Result := 'Janvier '   + IntToStr(a);
    2  : Result := 'Fevrier '   + IntToStr(a);
    3  : Result := 'Mars '      + IntToStr(a);
    4  : Result := 'Avril '     + IntToStr(a);
    5  : Result := 'Mai '       + IntToStr(a);
    6  : Result := 'Juin '      + IntToStr(a);
    7  : Result := 'Juillet '   + IntToStr(a);
    8  : Result := 'Aout '      + IntToStr(a);
    9  : Result := 'Septembre ' + IntToStr(a);
    10 : Result := 'Octobre '   + IntToStr(a);
    11 : Result := 'Novembre '  + IntToStr(a);
  end;

  Result := AnsiUpperCase(Result);
end;

{Calcul le nombre de mois entre deux dates}
{---------------------------------------------------------------------------------------}
function GetNbMoisEntre(Date1, Date2 : TDateTime) : Integer;
{---------------------------------------------------------------------------------------}
var
  a, e, b, c, d : Word;
begin
  DecodeDate(Date1, a, c, b);
  DecodeDate(Date2, e, d, b); {01/12/05 : Date2 c'est mieux que Date1 !!!}
  if c > d then Result := Abs(12 - c + d)
  else if (c = d) and (a + 1 = e) then Result := 12
                                  else Result := Abs(c - d);
  if a > e + 1 then Result := Result + 12 * Abs(a - e - 1);
end;

{Arrondi un montant selon la condition
{---------------------------------------------------------------------------------------}
function ArrondirTreso(Arrondi : string; Mnt: Double): Double;
{---------------------------------------------------------------------------------------}
var
  N : Double;
begin
  if Arrondi = '0' then begin {Pas d'arrondi}
    Result := Mnt;
    Exit;
  end;

  if Arrondi = '01' then {Au dixième près}
    N := 0.1
  else if Arrondi = '1' then {A l'unité près}
    N := 1
  else if Arrondi = '10' then {A 10 près}
    N := 10
  else if Arrondi = '100' then {A 100 près}
    N := 100
  else if Arrondi = 'K' then {A 1000 près}
    N := 1000
  else {'M' Au million}
    N := 1000000;
  Mnt := Mnt / N; {Séparer du Round sinon erreur de calcul !}
  Result := Round(Mnt) * N;
end;

{Regarde si "ATester" appartient à [Inf, Sup] }
{---------------------------------------------------------------------------------------}
function Between(ATester, Inf, Sup : Double) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (ATester >= Inf) and (ATester <= Sup);
end;

{---------------------------------------------------------------------------------------}
function Between(ATester, Inf, Sup : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (ATester >= Inf) and (ATester <= Sup);
end;

{21/10/04 : Fait un FormatFloat en fonction du nombre de décimales}
{---------------------------------------------------------------------------------------}
function FormateMontant(Mnt : Double; Dec : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  D, Z : string;
  n    : Integer;
begin
//  D := '#,';
  Z := '0.';
  for n := 1 to Dec do begin
//    D := D + '#';
    Z := Z + '0';
  end;
  Result := FormatFloat(D + Z, Mnt);
end;

{03/01/05 : s'assure que Chiffre est différent de zéro ;
           Defaut est la valeur à renvoyer si c'était le cas
{---------------------------------------------------------------------------------------}
function DoubleNotNul(Chiffre : Double; Defaut : Double = 1) : Double;
{---------------------------------------------------------------------------------------}
begin
  if Between(Defaut , - 0.00000000001, 0.00000000001) then Defaut := 1;
  if Between(Chiffre, - 0.00000000001, 0.00000000001) then Result := Defaut
                                                      else Result := Chiffre;
end;

{Chargement dynamique d'un glyph}
{---------------------------------------------------------------------------------------}
procedure ChargeGlyph(bb : TToolbarButton97; Img : string = '');
{---------------------------------------------------------------------------------------}
const
  {Rem : dans Alt + F12 de delphi, les Glyph.Data ne sont pas au bon format;
    if faut enlever ce qui est devant 424D42 sur la première ligne
    42010000424D4201000000000000760000002800000011000000110000000100 ->
            424D4201000000000000760000002800000011000000110000000100
    }
  Data =
      '424D4201000000000000760000002800000011000000110000000100' +
      '040000000000CC000000CE0E0000D80E00001000000000000000000000000000' +
      'BF0000BF000000BFBF00BF000000BF00BF00BFBF0000C0C0C000808080000000' +
      'FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00777777777777' +
      '77777000000070000000000000077000000070FFF0FFFFFF0F077000000070F0' +
      'F0F0FF0F0F077000000070000000000000077000000077709999999077777000' +
      '0000777090000907777770000000777090709007777770000000777090099700' +
      '77777000000077709099070007777000000077709990770BB077700000007770' +
      '9907770BB07770000000777090777770BB0770000000777007777770B0077000' +
      '00007770777777770BB070000000777777777777000770000000777777777777' +
      '777770000000';
var
  s   : TMemoryStream  ;
begin
  if Img = '' then Img := Data;

  s := TMemoryStream .Create;
  try
    S.SetSize(Length(Img) div 2);
    HexToBin(PAnsiChar(Img),s.Memory,s.Size);
    s.Seek(0,0);
    bb.Glyph.LoadFromStream(s);
  finally
    s.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure LibereListe(var L : TStringList; Detruire : Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  o : TObject;
begin
  if Assigned(L)  then begin
    try
      for n := 0 to L.Count - 1 do
        if Assigned(L.Objects[n]) then begin
          o := L.Objects[n];
          FreeAndNil(o);
        end;
    finally
      if Detruire then FreeAndNil(L)
                  else L.Clear;
    end;
  end;
end;

{Libération de la mémoire allouée à des variables dynamiques contenues dans une TLIST
{---------------------------------------------------------------------------------------}
procedure DisposeListe(var L : TList; Detruire : Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if Assigned(L)  then begin
    for n := 0 to L.Count - 1 do
      if Assigned(L[n]) then begin
        Dispose(L[n]);
        L[n] := nil;
      end;

    if Detruire then FreeAndNil(L)
                else L.Clear;
  end;
end;

{Constitution d'un nom de fichier
{---------------------------------------------------------------------------------------}
function GetNomFichier(Rep, Nom, Ext : string) : string;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  n : Integer;
begin
  if (Ext <> '') and (Ext[1] <> '.') then Ext := '.' + Ext;

  n := Length(Rep);
  if Rep[n] = '\' then System.Delete(Rep, n, 1);
  n := 1;

  Result := Nom;
  DecodeDate(Date, a, m, j);
  Result := Result + '_' + IntTostr(a) + '_' + PadL(IntTostr(m), '0', 2) + '_' + PadL(IntTostr(j), '0', 2) + '_';
  while FileExists(Rep + '\' + Result + IntToStr(n) + Ext) do Inc(n);
  Result := Result + IntToStr(n) + Ext;
end;

{Copie d'un fichier : Fonction est le traitement appelant pour la génération du nom du fichier
{---------------------------------------------------------------------------------------}
function CopierFichier(FicSource, RepDest, Fonction, Ext : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  S : string;
  P, Q : PChar;
begin
  if DirectoryExists(RepDest) then begin
    try
      {Generation du nom du fichier de destination}
      S := GetNomFichier(RepDest, Fonction, Ext);
      P := PChar(FicSource);
      S := RepDest + '\' + S;
      Q := PChar(S);
      {Copie du fichier source}
      Result := CopyFile(P, Q, True);
    except
      on E : Exception do begin
        PGIError(TraduireMemoire('Impossible de faire la copie du fichier.') + #13 + E.Message);
        Result := False;
      end;
    end
  end
  else
    Result := False;
end;

{Execute un RollBack en cachant le message en jaune sur le HPanel
{---------------------------------------------------------------------------------------}
procedure RollBackDiscret;
{---------------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
var
  e : THEdit;
{$ENDIF EAGLSERVER}
begin
  {$IFDEF EAGLSERVER}
  RollBack;
  {$ELSE}
  e := THEdit(Valerte);
  THEdit(Valerte) := nil;
  RollBack;
  THEdit(Valerte) := e;
  {$ENDIF EAGLSERVER}
end;

{15/03/05 : Calcule le HT à partir du montant TTC et du Taux de TVA
{---------------------------------------------------------------------------------------}
function GetMntHTFromTTC(MntTTC, TxTVA : Double; NbDec : Integer = -1) : Double;
{---------------------------------------------------------------------------------------}
begin
  if NbDec = - 1 then NbDec := V_PGI.OkDecV;
  Result := Arrondi(MntTTC / (1 + Abs(TxTVA)), NbDec);
end;

{15/03/05 : Calcule le montant de la TVA à partir du montant TTC et du Taux de TVA
{---------------------------------------------------------------------------------------}
function GetMntTVAFromTTC(MntTTC, TxTVA : Double; NbDec : Integer = -1) : Double;
{---------------------------------------------------------------------------------------}
begin
  if NbDec = - 1 then NbDec := V_PGI.OkDecV;
  Result := Arrondi(MntTTC, NbDec) - GetMntHTFromTTC(MntTTC, TxTVA);
end;

end.

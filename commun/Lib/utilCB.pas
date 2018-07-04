{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Gestion des cartes bancaires
Mots clefs ... : CB
*****************************************************************}

unit UtilCB;

interface

uses
  SysUtils,
  {$IFNDEF EAGLCLIENT}
  db,
  {$ENDIF}
  UTob, LicUtil, HEnt1, HMsgBox, HCtrls;

const
  TailleNoCarteCBDef = 16; // nombre de caractère d'un numéro de carte bancaire

type
  TCardType = (ctInvalid, ctVISA, ctDinerClub, ctMasterCard, ctAmEx, ctDiscover, ctCarteBleue, ctUnknown);

function VerifNoCarteCB(NCB: string): boolean;
function GoodCBNumber(NCB: string; CT: TCardType): boolean;
function GoodCBNumberCode(NCB, CodeCT: string): boolean;
function CardType(NCB: string): TCardType;
function CodeCardType(NCB: string): string;
function DeCrypteNoCarteCB(NumCarte: string; ModeUS: boolean; Complet: boolean = True; ForceDecrypte: boolean = False): string;
function CrypteNoCarteCB(NumCarte: string): string;
function FormuleMULDeCrypteCB(Params: string; TT: TDataset): string;
function FormuleOLEDeCrypteCB(Params: string): string;
function VerifDateExpireCB(dateExpire: string): boolean;

implementation

uses
  Ent1,UFonctionsCBP;

function Numbaz(S: string): string;
var i: integer;
begin
  for i := Length(S) downto 1 do if not (S[i] in ['0'..'9']) then Delete(S, i, 1);
  Result := S;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Contrôle un numéro de carte bancaire.
Mots clefs ... : CB
*****************************************************************}

function VerifNoCarteCB(NCB: string): boolean;
var
  Sum: Word;
  Fois2: boolean;
  i, N: integer;
begin
  Result := False;
  NCB := Numbaz(NCB);
  if Length(NCB) < 4 then Exit;
  Sum := 0;
  Fois2 := True;
  for i := Length(NCB) - 1 downto 1 do
  begin
    N := Ord(NCB[i]) - Ord('0');
    if Fois2 then N := N * 2;
    Fois2 := not Fois2;
    Sum := Sum + (N mod 10) + (N div 10);
  end;
  Sum := ((((Sum - 1) div 10) + 1) * 10) - Sum;
  Result := (Sum = (Ord(NCB[Length(NCB)]) - Ord('0')));
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Determine le type de carte en fonction d'un numéro de carte
Suite ........ : bancaire.
Mots clefs ... : CB
*****************************************************************}

function CardType(NCB: string): TCardType;
var Deb: integer;
begin
  NCB := Numbaz(NCB);
  Deb := StrToIntDef(Copy(NCB, 1, 6), -1);
  if (Deb = -1) or (not VerifNoCarteCB(NCB)) then Result := ctInvalid else
    if ((Deb >= 510000) and (Deb <= 559999)) and (Length(NCB) = 16) then Result := ctMasterCard else
    if ((Deb >= 400000) and (Deb <= 499999)) and ((Length(NCB) = 13) or (Length(NCB) = 16)) then Result := ctVISA else
    if (((Deb >= 300000) and (Deb <= 305999)) or ((Deb >= 360000) and (Deb <= 369999)) or ((Deb >= 380000) and (Deb <= 389999))) and (Length(NCB) = 15) then
    Result := ctAmEx else
    if (((Deb >= 340000) and (Deb <= 349999)) or ((Deb >= 370000) and (Deb <= 379999))) and (Length(NCB) = 14) then Result := ctDinerClub else
    if ((Deb >= 601100) and (Deb <= 601199)) and (Length(NCB) = 16) then Result := ctDiscover
  else Result := ctUnknown;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : En fonction d'un numéro de carte bancaire, détermine le
Suite ........ : type de carte selon la tablette YYCARTEINTERNET
Mots clefs ... : CB
*****************************************************************}

function CodeCardType(NCB: string): string;
var CT: TCardType;
begin
  CT := CardType(NCB);
  case CT of
    ctInvalid: Result := '';
    ctVISA: Result := 'VIS';
    ctDinerClub: Result := 'DIN';
    ctMasterCard: Result := 'MAS';
    ctAmEx: Result := 'AME';
    ctDiscover: Result := 'DIS';
    ctCarteBleue: Result := 'CB';
  else Result := 'CB';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Vérifie la cohérence entre un numéro de carte et un type de
Suite ........ : carte.
Mots clefs ... : CB
*****************************************************************}

function GoodCBNumber(NCB: string; CT: TCardType): boolean;
begin
  Result := (CardType(NCB) = CT);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Vérifie la cohérence entre un numéro de carte et un type de
Suite ........ : carte selon la tablette YYCARTEINTERNET.
Mots clefs ... : CB
*****************************************************************}

function GoodCBNumberCode(NCB, CodeCT: string): boolean;
begin
  Result := (CodeCardType(NCB) = CodeCT);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Indique si un numéro d'une carte bancaire est crypté
Mots clefs ... : CB
*****************************************************************}

function NoCarteCBIsCrypte(NumCarte: string): boolean;
begin
  Result := (Length(NumCarte) > TailleNoCarteCBDef);
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Décrypte en partie ou complétement un numéro de carte
Suite ........ : bancaire.
Suite ........ :
Suite ........ : Le décryptage partiel donne au format US
Suite ........ : ************1234 sinon ******123456789*
Mots clefs ... : CB
*****************************************************************}

function DeCrypteNoCarteCB(NumCarte: string; ModeUS: boolean; Complet: boolean = True; ForceDecrypte: boolean = False): string;
var
  Stg: string;
  Ind, NbCar: integer;
  Ch: char;
begin
  if NoCarteCBIsCrypte(NumCarte) then Stg := DeCryptageSt(NumCarte) else Stg := NumCarte;
  if (Complet) and ((ForceDecrypte) or (ExJaiLeDroitConcept(TConcept(gcNumCBComplet), False))) then
  begin
    Result := Stg;
  end else
  begin
    if ModeUS then
    begin
      Ind := 13;
      NbCar := 4;
    end else
    begin
      Ind := 7;
      NbCar := 9;
    end;
    Ch := '*';
    if Ind > 1 then Result := StringOfChar(Ch, (Ind -1));
    Result := Result + Copy(Stg, Ind, NbCar);
    NbCar := Length(Stg) - Length(Result);
    if NbCar > 0 then Result := Result + StringOfChar(Ch, NbCar);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Crypte le numéro d'une carte bancaire
Mots clefs ... : CB
*****************************************************************}

function CrypteNoCarteCB(NumCarte: string): string;
begin
  if NoCarteCBIsCrypte(NumCarte) then
    Result := NumCarte
  else
    Result := CryptageSt(NumCarte);
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Décryptage d'un numéro de carte bancaire depuis un MUL
Suite ........ : 
Suite ........ : @DECRYPTENOCARTECB(GPE_CBINTERNET;
Suite ........ : MP_AFFICHNUMCBUS)
Mots clefs ... : CB;MUL
*****************************************************************}
function FormuleMULDeCrypteCB(Params: string; TT: TDataset): string;
var
  s1, s2, s3: string;
  sCode: string;
  Fld: TField;
  ModeUS: boolean;
begin
  s1 := ReadTokenPipe(Params, ';'); // Nom du champ n° carte
  s2 := ReadTokenPipe(Params, ';'); // Nom du champ mode d'affichage (US ou français)
  s3 := ReadTokenPipe(Params, ';'); // Mode d'affichage (US ou français) (si pas de s2)
  Fld := TT.FindField(s1);
  if Fld <> nil then
  begin
    sCode := Fld.AsString;
    if sCode <> '' then
    begin
      Fld := TT.FindField(s2);
      ModeUS :=(((Fld <> nil) and (Fld.AsString = 'X')) or (s3 = 'X'));
      Result := DeCrypteNoCarteCB(sCode, ModeUS);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Décryptage d'un numéro de carte bancaire depuis un état
Suite ........ :
Suite ........ : @DECRYPTENOCARTECB([GPE_CBINTERNET];
Suite ........ : [MP_AFFICHNUMCBUS)
Mots clefs ... : CB;MUL
*****************************************************************}
function FormuleOLEDeCrypteCB(Params: string): string;
var
  s1, s2, s3: string;
begin
  s1 := ReadTokenPipe(Params, ';'); // Nom du champ n° carte
  s2 := ReadTokenPipe(Params, ';'); // Nom du champ mode d'affichage (US ou français)
  s3 := ReadTokenPipe(Params, ';'); // ModeComplet)

  Result := DeCrypteNoCarteCB(s1, (s2 = 'X'), (s3 = 'X'));
end;

function VerifDateExpireCB(dateExpire: string): boolean;
var Mois, Annee: string;
  i, iMois, iAnnee: integer;
  Year, Month, Day: Word;
begin
  Result := False;
  Mois := copy(DateExpire, 1, 2);
  Annee := copy(DateExpire, 4, 4);
  for i := 1 to Length(Mois) do
    if (not (Mois[i] in ['0'..'9'])) then exit;
  for i := 1 to Length(Annee) do
    if (not (Annee[i] in ['0'..'9'])) then exit;
  DecodeDate(Now, Year, Month, Day);
  iMois := strToInt(Mois);
  if (iMois < 1) or (iMois > 12) then exit;
  iAnnee := strToInt(Annee);
  if (iAnnee < 2000) or (iAnnee > 2099) then exit;
  if iAnnee < Year then exit;
  if (iannee = year) and (iMois < Month) then exit;
  result := True;
end;

end.

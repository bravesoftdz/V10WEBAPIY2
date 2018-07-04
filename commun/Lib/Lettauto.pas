{***********UNITE*************************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 04/05/2000
Modifié le ... : 04/05/2000
Description .. : Rapprochement automatique récursif entre un montant à
Suite ........ : rapprocher et une liste d'éléments
Suite ........ : Cherche une solution exacte (sinon voir l'unité
Suite ........ : Approxim.pas)
Mots clefs ... : RAPPROCHEMENT;COMBINATOIRE;
*****************************************************************}
unit LettAuto;

interface

uses Windows, Messages, SysUtils, Classes, HEnt1;

type REC_AUTO = record
    Nival, NbD, Decim, Temps: integer;
    Unique: boolean;
  end;
const MaxDroite = 1000;
const MaxNivProf = 10;
const MaxTempo = 30;
const MaxSolRech = 10;
type T_D = array[0..MaxDroite - 1] of Double;
  T_I = array[0..MaxDroite - 1] of Smallint;
type TSolLet = (slPasSolution, slSolutionUnique, slSolutionMultiple, slPaquet);
type T_LAS = array[1..100] of TSolLet;
  P_LAS = class
    S: T_LAS;
  end;
  P_I = class
    I: T_I;
  end;


function LettrageAuto(ALettrer: double; var ListeDroite: T_D; var ListeSortie: T_I;
  Infos: REC_AUTO): integer;

implementation

var Encours: T_I;
  TotalNeg, TotalPos: T_D;
  NbAppels, Debray, Arrete: Comp;
  TT1: TDateTime;
var Vale: double;
  Niveau, NbDec, Tempo: integer;
  NbDroite: Longint;
  MtantRappro {,DeltaRappro  }: Double;
const NbDebray: Longint = 3125;

type X_TRI = array[0..MaxDroite - 1] of
  record
    PosI: integer;
    X: Double;
  end;

var Montant: X_TRI;
  SolUnique: boolean;

{============================================================================}

procedure POSITIONNEDEBRAYE;
var NbS: Longint;
  HH, MM, SS, CC: Word;
begin
  DecodeTime(Time - TT1, HH, MM, SS, CC);
  NbS := 3600 * HH + 60 * MM + SS;
  if NbS >= Tempo then begin Arrete := 1; Exit; end;
  if NbS > 0 then
  begin
    Arrete := Round(NbDebray / NbS * 120);
    Debray := Round(NbDebray / NbS * Tempo);
    Arrete := Debray;
  end else if NbDebray < 2000000000 then
  begin
    NbDebray := 2 * NbDebray;
  end else Arrete := 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Louis DECOSSE
Créé le ...... : 04/05/2000
Modifié le ... : 04/05/2000
Description .. : Algo récursif fondamental pour le rappro combinatoire
Mots clefs ... : RAPPROCHEMENT;COMBINATOIRE;RECURSIVITE;
*****************************************************************}

function ALGOCOMBI(Reste: double; Pos, L: integer): integer;
var ii, Position, Res: integer;
begin
  AlgoCombi := 0;
{gestion du "débrayage"}
  NbAppels := NbAppels + 1;
  if NbAppels = NbDebray then PositionneDebraye;
  if ((Arrete > 0) and (NbAppels > Arrete)) then begin AlgoCombi := -2; Exit; end;
  if ((Debray > 0) and (NbAppels > Debray)) then begin AlgoCombi := -2; Exit; end;
{ Algorithme récursif }
  Position := Pos;
  for ii := Position to NbDroite - 1 do
  begin
    Vale := Reste + Montant[ii].X; if EnCours[ii] > 0 then Continue;
    if Arrondi(Vale - MtantRappro, NbDec) = 0 then begin EnCours[ii] := 1; AlgoCombi := 1; Exit; end;
    if (Vale + TotalNeg[ii] > MtantRappro) or (Vale + TotalPos[ii] < MtantRappro) then Continue;
    if l >= Niveau then Continue;
    EnCours[ii] := 1; Res := AlgoCombi(Vale, ii + 1, l + 1);
    if Res <> 0 then begin AlgoCombi := Res; Exit; end;
    EnCours[ii] := 0;
  end;
end;

{============================================================================}

function REMPLITNEGPOS: integer;
var i, Niv: Longint;
  X, STN, STP: Double;
begin
  RemplitNegPos := 0; STN := 0; STP := 0; {X:=0 ;} Niv := 0;
  for i := NbDroite - 1 downto 0 do
  begin
    if Arrondi(Montant[i].X - MtantRappro, NbDec) = 0 then begin RemplitNegPos := 1000 + Montant[i].PosI; Exit; end;
    TotalNeg[i] := STN; TotalPos[i] := STP;
    if Montant[i].X > 0 then STP := STP + Montant[i].X else STN := STN + Montant[i].X;
  end;
  if SolUnique then Exit;
{Peut on espérer une solution ?}
  if ((STP > 0) and (STP < MtantRappro)) then Exit;
  if ((STN < 0) and (STN > MtantRappro)) then Exit;
{Somme Totale = Montant}
  if not SolUnique then
  begin
    if Arrondi(STP + STN - MTantRappro, NbDec) = 0 then begin RemplitNegPos := 2; Exit; end;
    if Arrondi(STP - MTantRappro, NbDec) = 0 then begin RemplitNegPos := 3; Exit; end;
    if Arrondi(STN - MTantRappro, NbDec) = 0 then begin RemplitNegPos := 4; Exit; end;
  end;
{si que des positifs}
  if ((STP > 0) and (STN = 0)) then
  begin
    if MtantRappro <= 0 then Exit; if Montant[NbDroite - 1].X > MtantRappro then Exit;
    i := NbDroite - 1; X := 0;
    while ((X <= MtantRappro) and (i >= 0)) do begin X := X + Montant[i].X; Dec(i); Inc(Niv); end;
    Dec(Niv, 1); if ((Niv >= 0) and (Niv < Niveau)) then Niveau := Niv;
  end;
{si que des négatifs}
  if ((STP = 0) and (STN < 0)) then
  begin
    if MtantRappro > 0 then Exit; if Montant[0].X < MtantRappro then Exit;
    i := 0; X := 0;
    while ((X >= MtantRappro) and (i <= NbDroite - 1)) do begin X := X + Montant[i].X; Inc(i); Inc(Niv); end;
    Dec(Niv, 1); if ((Niv >= 0) and (Niv < Niveau)) then Niveau := Niv;
  end;
  RemplitNegPos := 1;
end;

{============================================================================}

procedure DEMARRETIME;
begin
  TT1 := Time; NbAppels := 0; Debray := 0; Arrete := 0;
end;

{============================================================================}
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : - LG - 27/04/2006 - FB 15201 - on respecte l'ordre
Suite ........ : chrnologique pour deux montants identiques
Mots clefs ... :
*****************************************************************}

procedure TriMontants(var LD: T_D);
var i, j: Longint;
  X: real;
  Pos: Longint;
  MM: X_TRI;
begin
  FillChar(Montant, Sizeof(Montant), #0); FillChar(MM, Sizeof(MM), #0);
  for i := 0 to NbDroite - 1 do begin Montant[i].PosI := i; Montant[i].X := LD[i]; end;
  for i := NbDroite - 1 downto 1 do
  begin
    for j := 0 to i - 1 do
    begin
      if (Arrondi(Montant[j].X - Montant[i].X, NbDec) > 0) or
        ((Arrondi(Montant[j].X - Montant[i].X, NbDec) = 0) and (Montant[j].PosI < Montant[i].PosI)) then
      begin
        X := Montant[j].X; Pos := Montant[j].PosI;
        Montant[j].X := Montant[i].X; Montant[j].PosI := Montant[i].PosI;
        Montant[i].X := X; Montant[i].PosI := Pos;
      end;
    end;
  end;
  for i := 0 to NbDroite - 1 do MM[i] := Montant[NbDroite - 1 - i];
  Montant := MM;
end;


{============================================================================}

function LettrageAuto(ALettrer: double; var ListeDroite: T_D; var ListeSortie: T_I;
  Infos: REC_AUTO): integer;
var Res, i: integer;
begin
  Result := 0; NbDroite := Infos.NbD; Niveau := Infos.Nival; MtantRappro := ALettrer;
  NbDec := Infos.Decim; Tempo := Infos.Temps; SolUnique := Infos.Unique;
  FillChar(TotalNeg, Sizeof(TotalNeg), #0); FillChar(TotalPos, Sizeof(TotalPos), #0); FillChar(Encours, Sizeof(Encours), #0);
  TriMontants(ListeDroite);
  Res := RemplitNegPos;
  case Res of
    0: Exit; {pas de solution possible}
    1: begin {possible solution combinatoire}
        DemarreTime; Res := AlgoCombi(0, 0, 0);
        if Res > 0 then
        begin
          for i := 0 to NbDroite - 1 do if Encours[i] > 0 then ListeSortie[Montant[i].PosI] := 1;
          Result := 1; Exit;
        end else if Res = -2 then
        begin
          Result := 2; Exit;
        end;
      end;
    2: begin {solution triviale (Somme liste droite = montant à lettrer)}
        for i := 0 to NbDroite - 1 do ListeSortie[i] := 1; Result := 1; Exit;
      end;
    3: begin {solution triviale (Somme liste positif = montant à lettrer)}
        for i := 0 to NbDroite - 1 do if Montant[i].X > 0 then ListeSortie[Montant[i].PosI] := 1; Result := 1; Exit;
      end;
    4: begin {solution triviale (Somme liste neg = montant à lettrer)}
        for i := 0 to NbDroite - 1 do if Montant[i].X < 0 then ListeSortie[Montant[i].PosI] := 1; Result := 1; Exit;
      end;
  else begin {un élément de liste droite couvre totalement le à lettrer}
      if Res >= 1000 then begin ListeSortie[Res - 1000] := 1; Result := 1; Exit; end;
    end;
  end;
end;

{============================================================================}
{============================================================================}
end.

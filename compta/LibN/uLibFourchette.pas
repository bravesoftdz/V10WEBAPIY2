{***********UNITE*************************************************
Auteur  ...... : Christophe AYEL
Créé le ...... : 30/10/2007
Modifié le ... :   /  /
Description .. : Unité de gestion des fourchettes de comptes.
Mots clefs ... :
*****************************************************************}
unit uLibFourchette;
{
  TEnsembleFourchettes : Objet de gestion des unions, instersections, complément
    d'un ensemble de fourchettes

  TFourchetteItem :
    Objet fouchette comprenant juste un CompteDebut, et un CompteFin
}

interface

uses classes, Contnrs;

type

  TCompteNumero = string;

  //description voir la méthode WhereIs
  TWhere = (wkUnknown, wkIn, wkSelfIn, wkOnDown, wkOnUp, wkOutDown, wkOutUp, wkCrossDown, wkCrossUp);

  TFourchetteItem = class(TObject)
  	protected
    	FCompteDebut : TCompteNumero;
      FCompteFin : TCompteNumero;
    public
      // procedure Assign(AValue : TCustomFourchette); virtual;
    	property CompteDebut : TCompteNumero read FCompteDebut;
      property CompteFin : TCompteNumero read FCompteFin;
      function WhereIs(aItem : TFourchetteItem) : TWhere;
    end;

  TEnsembleFourchettes = class(TObjectList)
  private
    FCompteLongueur: integer; // FParametreDossier : IParametreDossier;
    FFourchetteSeparateur: string;
    FCompteSeparateur: string;
    FCompteMin,
    FCompteMax : TCompteNumero;
    procedure SetFourchettes(Index : integer; Valeur : TFourchetteItem);
    function GetFourchettes(Index : integer) : TFourchetteItem;
    procedure SetFourchettesString(const Value: string);
    function GetFourchettesString: string;
    procedure SetCompteSeparateur(const Value: string);
    procedure SetFourchetteSeparateur(const Value: string);
    procedure SetCompteMax(const Value: TCompteNumero);
    procedure SetCompteMin(const Value: TCompteNumero);
    function GetFourchettesStringComplement: string;
  protected
    function  Indice(Compte : TCompteNumero) : integer;
    procedure AddFourchettesString(const Value: string; ViderBefore : boolean);
  public
    constructor Create(const iCompteLongueur: integer); reintroduce;
    //ajouter une fouchette
    function Ajouter(CpteDeb, CpteFin : TCompteNumero) : TFourchetteItem;
    procedure AjouterFourchettesString(aFourchettesString : string);
    procedure Complement(Resultat : TEnsembleFourchettes);
    function Get_A_NotIn_B(const A, B : string) : string;
    function GetFourchettesStringEnsembleComplement(const aReference, aSousEnsemble : string): string;
    procedure GetFourchettesStrings(const Liste: TStrings);
    //renvoi l'intersection c.a.d si 10000000-15000000 x 14000000-16000000 -> 14000000-15000000
    procedure Intersections(Resultat : TEnsembleFourchettes); overload;
    //renvoi les fourchettes en cause c.a.d si 10000000-15000000 x 14000000-16000000
    //  -> 10000000-15000000 et 14000000-16000000
    procedure Intersections(Liste : TStringList); overload;
    procedure Union(Resultat : TEnsembleFourchettes);
    function CompteRechercher(const sCompte: string): TFourchetteItem; overload;
    function Rechercher(CpteDeb, CpteFin : TCompteNumero) : TFourchetteItem; overload;
    function Rechercher(const aFourchette : string) : TFourchetteItem; overload;
    procedure Trier;
    //une foruchette par ligne
    procedure SetFourchettesStrings(const Liste: TStrings);

    property Fourchettes[Index : integer] : TFourchetteItem read GetFourchettes
             write SetFourchettes; default;
    property FourchettesString : string read GetFourchettesString write SetFourchettesString;
    property FourchettesStringComplement : string read GetFourchettesStringComplement;
    property FourchetteSeparateur : string read FFourchetteSeparateur write SetFourchetteSeparateur;
    property CompteSeparateur : string read FCompteSeparateur write SetCompteSeparateur;
    property CompteMax : TCompteNumero read FCompteMax write SetCompteMax;
    property CompteMin : TCompteNumero read FCompteMin write SetCompteMin;
  end;

implementation

uses sysUtils;

const
//  CST_COMPTEMIN = '00000000';
//  CST_COMPTEMAX = '99999999';
  DefaultSeparatorCpte = '-';
  DefaultSeparatorFour = ';';

function CompteNumeroCompleter(const sCompte: string;
  const cComplete: Char;
  const iCompteLongueur: integer): string;
begin
  Result := StringReplace(Trim(sCompte), ' ', cComplete, [rfReplaceAll]);

  // compléter avec le caractère de complément
  if Length(Result) < iCompteLongueur then
    while length(Result) < iCompteLongueur do
      Result :=  Result + cComplete

  // tronquer les numéros en trop
  else if length(Result) > iCompteLongueur then
    Result := Copy(Result, 1, iCompteLongueur);
end;

function SepareString(const S : string; var Deb, Fin : string; separator : string) : boolean;
var
  P : Pchar;
  index : integer;
  {scinde une chaine en deux chaines, le découpage se trouve à la position
   du separator, celui-lui est retiré du résultat}
begin
  P := PChar(S);
  index := Pos(separator, P);
  if (index > 0) then
  begin
    Deb := Copy(P, 1, index-1);
    Inc(P, index);
    Fin := P;
    Result := True;
  end
  else
    Result := False;
end;

function GetNombre(const s : string; var Res : string) : integer;
var
  P : PChar;
  {retourne le début de la chaine commençant par des chiffres}
begin
  P := PChar(S);
  while (P^ in ['0'..'9']) do
    Inc(P);
  Result := P - PChar(S);
  Res := Copy(S, 1, Result);
end;

function FourchetteValid(const S : string; separator : string) : boolean;
var
  index : integer;
  Res : string;
  pStr : PChar;
begin
  Result := False;
  pStr := PChar(S);
  index := GetNombre(pStr, Res);
  if (index > 0) then
  begin
    Inc(pStr, index);
    if pStr^ = separator then
    begin
      Inc(pStr);
      Inc(pStr, GetNombre(pStr, Res));
      //si je suis en fin de chaine c'est ok
      Result := (pStr^ = #0) and (Res <> '');
    end
  end;
end;

function GetFourchette(const S : string; var Res : string; separator : string) : integer;
begin
  Result := Pos(separator, S);
  if Result > 0 then
    Res := Copy(S, 1, Result-1)
  else
  begin
    Res := S;
    Result := Length(S);
  end;
end;

function FourchettesValid(const S : string; Liste : TStrings; SeparatorFour,
                          SeparatorCpte : string) : boolean;
var
  index : integer;
  Res, Deb, Fin : string;
  pStr : PChar;
  DupList : TStringList;
begin
  pStr := PChar(S);
  index := GetFourchette(pStr, Res, SeparatorFour);
  if index <= 0 then
    Result := False
  else
    Result := True;
  DupList := TStringList.Create;
  try
    while index > 0 do
    begin
      Inc(pStr, index);
      if Res <> '' then
      begin
        Result := Result and FourchetteValid(Res, SeparatorCpte);
        if Result then
        begin
          SepareString(Res, Deb, Fin, SeparatorCpte);
          DupList.Add(Deb);
          DupList.Add(Fin);
        end;
      end;
      index := GetFourchette(pStr, Res, SeparatorFour);
    end;
    {si toutes les fouchettes sont valides, alors je recopie les fourchettes}
    if (Result) and (Liste <> nil) then
      Liste.Assign(DupList);
  finally
    DupList.Free;
  end;
end;

function ComptePlusUn(aCompte : TCompteNumero; const iCompteLongueur: integer) : TCompteNumero;
begin
  Result := CompteNumeroCompleter(IntToStr(StrToInt(aCompte) + 1), '0', iCompteLongueur)
end;

function CompteMoinsUn(aCompte : TCompteNumero; const iCompteLongueur: integer) : TCompteNumero;
begin
  Result := IntToStr(StrToInt(aCompte) - 1);
  while Length(Result) < iCompteLongueur do
    Result := '0' + Result;
end;

{ TFourchetteItem }

function TFourchetteItem.WhereIs(aItem: TFourchetteItem): TWhere;
{
    wkUnknown : il est nulle part ! -> aItem est nil
    wkIn : aItem est complètement dans Self
    wkSelfIn : Self est complètement dans aItem
    wkOnDown :  CompteDebut = aItem.CompteFin au minimum
    wkOnUp :  CompteFin = aItem.CompteDebut au minimum
    wkOutDown : aItem complement < à Self
    wkOutUp : aItem complement > à Self
    wkCrossDown : CompteDebut < aItem.CompteFin <= CompteFin
    wkCrossUp :   CompteDebut <= aItem.CompteDebut < CompteFin
}
var
  cmp, cmp1 : integer;
begin
  if aItem = nil then
  begin
    Result := wkUnknown;
    Exit;
  end;
  cmp := AnsiCompareStr(aItem.FCompteFin, FCompteDebut);
  if cmp = 0 then //un compte de chevauchement
    Result := wkOnDown
  else if cmp < 0 then //inférieure strictement
    Result := wkOutDown
  else
  begin
    //ici aItem.CompteFin > CompteDebut
    cmp := AnsiCompareStr(aItem.FCompteDebut, FCompteFin);
    if cmp = 0 then
      Result := wkOnUp
    else if cmp > 0 then
      Result := wkOutUp
    else
    begin
      //ici aItem.CompteDebut < CompteFin
      cmp := AnsiCompareStr(aItem.FCompteDebut, FCompteDebut);
      cmp1 := AnsiCompareStr(aItem.FCompteFin, FCompteFin);
      if (cmp >= 0) and (cmp1 <= 0) then
        Result := wkIn
      else if (cmp < 0) and (cmp1 <= 0) then
        Result := wkCrossDown
      else if (cmp < 0) and (cmp1 > 0) then
        Result := wkSelfIn
      else
        Result := wkCrossUp;
    end;
  end;
end;

{ TEnsembleFourchettes }

procedure TEnsembleFourchettes.AddFourchettesString(const Value: string;
  ViderBefore: boolean);
var
  i : integer;
  Liste : TStringList;
begin
  Liste := TStringList.Create;
  try
    if ViderBefore then
      Clear;
    if FourchettesValid(Value, Liste, FFourchetteSeparateur, FCompteSeparateur) then
    begin
      i := 0;
      while i < Liste.Count-1 do
      begin
        Ajouter(CompteNumeroCompleter(Liste[i], '0', FCompteLongueur),
                CompteNumeroCompleter(Liste[i+1], '9', FCompteLongueur));
        Inc(i, 2);
      end;
    end;
  finally
    Liste.Free;
  end;
end;

function TEnsembleFourchettes.Ajouter(CpteDeb,
  CpteFin: TCompteNumero): TFourchetteItem;
{ !!!!!!!! les comptes doivent être formater !!!!!!!!!}
begin
  // Result := TFourchetteItem(DoAjouter);
  Result := TFourchetteItem.Create;
  Add(Result);

  with Result do
  begin
    //je m'assure que CompteDebut est toujours inférieure à CompteFin
    if AnsiCompareStr(CpteDeb, CpteFin) <= 0 then
    begin
      FCompteDebut := CpteDeb;
      FCompteFin := CpteFin;
    end
    else
    begin
      FCompteDebut := CpteFin;
      FCompteFin := CpteDeb;
    end;
  end;
  //c'est important que la liste soit triée, pour les fonctions d'intersections,
  //d'union, de complément
  Trier;
end;

procedure TEnsembleFourchettes.AjouterFourchettesString(
  aFourchettesString: string);
begin
  AddFourchettesString(aFourchettesString, False);
end;

procedure TEnsembleFourchettes.Complement(Resultat: TEnsembleFourchettes);
{ Resultat -> toutes les fouchettes n'appartenant pas à Self}
var
  TmpRes : TEnsembleFourchettes;
  Cpte, Cpte1 : TCompteNumero;
  i : integer;
begin
  if Count = 0 then
    Resultat.Ajouter(FCompteMin, FCompteMax)
  else
  begin
    TmpRes := TEnsembleFourchettes.Create(FCompteLongueur);
    try
      // élimination des d'intersections
      Union(TmpRes);
      with TmpRes do
      begin
        CompteMin := Self.CompteMin;
        CompteMax := Self.CompteMax;
        // commencer par les extrémités
        with Fourchettes[0] do
          if AnsiCompareStr(CompteMin, CompteDebut) < 0 then
            Resultat.Ajouter(CompteMin, CompteMoinsUn(CompteDebut, FCompteLongueur));
        with Fourchettes[Count-1] do
          if AnsiCompareStr(CompteMax, CompteFin) > 0 then
            Resultat.Ajouter(ComptePlusUn(CompteFin, FCompteLongueur), CompteMax);
        //je regarde entre chaque fouchette
        if Count > 1 then
          for i := 0 to Count-2 do
          begin
            Cpte := ComptePlusUn(Fourchettes[i].CompteFin, FCompteLongueur);
            Cpte1 := CompteMoinsUn(Fourchettes[i+1].CompteDebut, FCompteLongueur);
            if CompareStr(Cpte, Cpte1) <= 0 then
              Resultat.Ajouter(Cpte, Cpte1);
          end;
      end;
    finally
      TmpRes.Free;
    end;
  end;
end;

constructor TEnsembleFourchettes.Create(const iCompteLongueur: integer {aParametreDossier : IParametreDossier});
begin
  inherited Create(True);
  FCompteLongueur := iCompteLongueur;
  // FClassFourchette := TFourchetteItem;
  FFourchetteSeparateur := DefaultSeparatorFour;
  FCompteSeparateur := DefaultSeparatorCpte;
  FCompteMin := CompteNumeroCompleter('', '0', FCompteLongueur);// CST_COMPTEMIN;
  FCompteMax := CompteNumeroCompleter('', '9', FCompteLongueur);// CST_COMPTEMAX;
end;

function TEnsembleFourchettes.GetFourchettesString: string;
var
  i : integer;
begin
  Result := '';
  for i:=0 to Count-1 do
  begin
    if i > 0 then
      Result := Result + FFourchetteSeparateur;
    Result := Result + Fourchettes[i].CompteDebut + FCompteSeparateur + Fourchettes[i].CompteFin;
  end;
end;

function TEnsembleFourchettes.GetFourchettesStringComplement: string;
var
  tmp : TEnsembleFourchettes;
begin
  tmp := TEnsembleFourchettes.Create(FCompteLongueur);
  try
    Complement(tmp);
    Result := tmp.FourchettesString;
  finally
    tmp.Free;
  end;
end;

function TEnsembleFourchettes.Get_A_NotIn_B(const A, B: string): string;
var
  inter, tmp, tmp1, tmp2 : TEnsembleFourchettes;
  one : string;
  i : integer;
begin
  if A = '' then
    Result := ''
  else if B = '' then
    Result := A
  else
  begin
    Result := '';
    tmp := TEnsembleFourchettes.Create(FCompteLongueur);
    try
      tmp1 := TEnsembleFourchettes.Create(FCompteLongueur);
      tmp2 := TEnsembleFourchettes.Create(FCompteLongueur);
      inter := TEnsembleFourchettes.Create(FCompteLongueur);
      try
        tmp1.FourchettesString := B;
        tmp1.AddFourchettesString(A, False);
        tmp1.Intersections(inter);
        tmp1.FourchettesString := A;
        for i := 0 to tmp1.Count-1 do
        begin
          one := tmp1.Fourchettes[i].CompteDebut + tmp1.CompteSeparateur + tmp1.Fourchettes[i].CompteFin;
          tmp.FourchettesString := one;
          tmp.AddFourchettesString(inter.FourchettesString, False);
          tmp2.Clear;
          tmp.Intersections(tmp2);
          if tmp2.Count = 0 then
          begin //pas d'intersections donc cette fouchette ne fait pas partie de B
            if Result <> '' then
              Result := Result + FFourchetteSeparateur;
            Result := Result + one;
          end
          else
          begin
            tmp2.CompteMax := tmp1.Fourchettes[i].CompteFin;
            tmp2.CompteMin := tmp1.Fourchettes[i].CompteDebut;
            one := tmp2.FourchettesStringComplement;
            if one <> '' then
            begin
              if Result <> '' then
                Result := Result + FFourchetteSeparateur;
              Result := Result + one;
            end;
          end;
        end;
      finally
        inter.Free;
        tmp2.Free;
        tmp1.Free;
      end;
    finally
      tmp.Free;
    end;
  end;
end;

function TEnsembleFourchettes.GetFourchettesStringEnsembleComplement(
  const aReference, aSousEnsemble : string): string;
{
  fonction qui renvoie le complement de l'ensemble aSousEnsemble par rapport
  à l'ensemble aReference
  Attention le sous ensemble doit être plus petit ou égal à la référence
  - renvoi les fourchettes de l'ensemble aReference qui ne sont pas dans aSousEnsemble
}
begin
  Result := Get_A_NotIn_B(aReference, aSousEnsemble);
end;

procedure TEnsembleFourchettes.GetFourchettesStrings(const Liste: TStrings);
var
  i : integer;
begin
  Liste.Clear;
  for i := 0 to Count-1 do
    Liste.Add(Fourchettes[i].CompteDebut + CompteSeparateur + Fourchettes[i].CompteFin);
end;

procedure TEnsembleFourchettes.Intersections(
  Resultat: TEnsembleFourchettes);
{ Renvoi dans resultat un ensemble des fouchettes resultant de l'intersection}
var
  i, index : integer;
  Item, Item1 : TFourchetteItem;
  TmpRes : TEnsembleFourchettes;
begin
  if Count < 1 then
    Exit;

  Resultat.Clear;
  TmpRes := TEnsembleFourchettes.Create(FCompteLongueur);
  try
    i := 0;
    while i < Count-1 do
    begin
      Item := Fourchettes[i] as TFourchetteItem;
      index := i + 1;
      while index < Count do
      begin
        Item1 := Fourchettes[index] as TFourchetteItem;
        case Item.WhereIs(Item1) of
          wkIn :
            TmpRes.Ajouter(Item1.CompteDebut, Item1.CompteFin);
          wkSelfIn :
            TmpRes.Ajouter(Item.CompteDebut, Item.CompteFin);
          wkOnDown, wkCrossDown  :
            TmpRes.Ajouter(Item.CompteDebut, Item1.CompteFin);
          wkOutDown, wkOutUp :
              Break;
          wkOnUp, wkCrossUp :
            TmpRes.Ajouter(Item1.CompteDebut, Item.CompteFin);
        end;
        Inc(Index);
      end; //while index < Count do
      Inc(i);
    end; //while i < Count-1 do
    TmpRes.Union(Resultat);
  finally
    TmpRes.Free;
  end;
end;

procedure TEnsembleFourchettes.Intersections(Liste: TStringList);
{ Renvoi dans resultat un ensemble des fouchettes resultant de l'intersection}
var
  i, index : integer;
  Item, Item1 : TFourchetteItem;
begin
  if Count < 1 then
    Exit;
  Liste.Clear;
  i := 0;
  while i < Count-1 do
  begin
    Item := Fourchettes[i] as TFourchetteItem;
    index := i + 1;
    while index < Count do
    begin
      Item1 := Fourchettes[index] as TFourchetteItem;
      case Item.WhereIs(Item1) of
        wkIn, wkSelfIn, wkOnDown, wkCrossDown, wkOnUp, wkCrossUp :
        begin
          Liste.Add(Item.CompteDebut + FCompteSeparateur + Item.CompteFin);
          Liste.Add(Item1.CompteDebut + FCompteSeparateur + Item1.CompteFin);
        end;
        wkOutDown, wkOutUp :
            Break;
      end;
      Inc(Index);
    end; //while index < Count do
    Inc(i);
  end; //while i < Count-1 do
end;

function TEnsembleFourchettes.CompteRechercher(const sCompte: string): TFourchetteItem;
{ Recherche dicotomique de la fourchette contenant un compte }
var
  iFourchette : integer;
begin
  iFourchette := Indice(sCompte);
  if iFourchette < 0 then
    Result := nil
  else
    Result := GetFourchettes(iFourchette);
end;

function TEnsembleFourchettes.Rechercher(CpteDeb,
  CpteFin: TCompteNumero): TFourchetteItem;
var
  i : integer;
begin
  Result := nil;
  for i := 0 to Count-1 do
  begin
    if (Fourchettes[i].CompteDebut = CpteDeb) and
       (Fourchettes[i].CompteFin = CpteFin) then
    begin
      Result := (Fourchettes[i] as TFourchetteItem);
      exit;
    end;
  end;
end;

function TEnsembleFourchettes.Rechercher(
  const aFourchette: string): TFourchetteItem;
var
  CpteDe, CpteA : string;
begin
  if FourchetteValid(aFourchette, CompteSeparateur) and
     SepareString(aFourchette, CpteDe, CpteA, CompteSeparateur) then
    Result := Rechercher(CpteDe, CpteA)
  else
    Result := nil;
end;

procedure TEnsembleFourchettes.SetCompteMax(const Value: TCompteNumero);
var
  NewValue : TCompteNumero;
begin
  if Value <> FCompteMax then
  begin
    NewValue := CompteNumeroCompleter(Value, '9', FCompteLongueur);
    if (AnsiCompareStr(FCompteMin, NewValue) <= 0) and
    (FourchetteValid(FCompteMin + FCompteSeparateur + NewValue, FCompteSeparateur) ) then
    begin
      FCompteMax := NewValue;
    end;
  end;
end;

procedure TEnsembleFourchettes.SetCompteMin(const Value: TCompteNumero);
var
  NewValue : TCompteNumero;
begin
  if Value <> FCompteMin then
  begin
    NewValue := CompteNumeroCompleter(Value, '0', FCompteLongueur);
    if (AnsiCompareStr(NewValue, FCompteMax) <= 0) and
    (FourchetteValid(NewValue + FCompteSeparateur + FCompteMax, FCompteSeparateur) ) then
    begin
      FCompteMin := NewValue;
    end;
  end;
end;

procedure TEnsembleFourchettes.SetCompteSeparateur(const Value: string);
begin
  FCompteSeparateur := Value;
end;

procedure TEnsembleFourchettes.SetFourchetteSeparateur(
  const Value: string);
begin
  FFourchetteSeparateur := Value;
end;

procedure TEnsembleFourchettes.SetFourchettesString(const Value: string);
begin
  AddFourchettesString(Value, True);
end;

procedure TEnsembleFourchettes.SetFourchettesStrings(const Liste: TStrings);
var
  i : integer;
  CpteDe, CpteA : string;
begin
  Clear;
  for i := 0 to Liste.Count-1 do
  begin
    if FourchetteValid(Liste[i], CompteSeparateur) and
       SepareString(Liste[i], CpteDe, CpteA, CompteSeparateur) then
    begin
      Ajouter(CompteNumeroCompleter(CpteDe, '0', FCompteLongueur),
              CompteNumeroCompleter(CpteA, '9', FCompteLongueur));
    end;
  end;
end;

procedure TEnsembleFourchettes.Union(Resultat: TEnsembleFourchettes);
{ Renvoi dans resultat l'ensemble des fouchettes sans intersection}
var
  i : integer;
  Item, Item1 : TFourchetteItem;
begin
  if Count = 0 then
    Exit;

  Resultat.Clear;
  Item := Fourchettes[0] as TFourchetteItem;
  Item := Resultat.Ajouter(Item.CompteDebut, Item.CompteFin);
  for i := 1 to Count-1 do
  begin
    Item1 := Fourchettes[i] as TFourchetteItem;
    case Item.WhereIs(Item1) of
      wkIn : ;//je ne fais rien
      //ne devrait pas arriver car la liste est triée par ordre croissant
      wkSelfIn, wkOnDown, wkCrossDown, wkOutDown : ;
      wkOutUp :
        begin
          Item := Resultat.Ajouter(Item1.CompteDebut, Item1.CompteFin);
        end;
      wkOnUp, wkCrossUp :
        Item.FCompteFin := Item1.CompteFin;
    end;
  end;
end;

function FourchetteComparer(Item1, Item2: Pointer): Integer;
{ Fonction de comparaison pour l'appel de Sort sur la liste }

  function InternalComparer(Item1,
    Item2: TFourchetteItem): integer;
  { Fonction de compaaison des fourchettes }
  begin
    Result := AnsiCompareStr(Item1.FCompteDebut, Item2.FCompteDebut) ;
    if Result = 0 then
      Result := AnsiCompareStr(Item1.FCompteFin, Item2.FCompteFin) ;
  end;

begin
  Result := InternalComparer(TFourchetteItem(Item1), TFourchetteItem(Item2));
end;

procedure TEnsembleFourchettes.Trier;
begin
  Sort(FourchetteComparer);
end;

function TEnsembleFourchettes.GetFourchettes(
  Index: integer): TFourchetteItem;
begin
  Result := Items[Index] as TFourchetteItem
end;

procedure TEnsembleFourchettes.SetFourchettes(Index: integer;
  Valeur: TFourchetteItem);
begin
  Items[Index] := Valeur;
end;

function TEnsembleFourchettes.Indice(Compte: TCompteNumero): integer;
{ Recherche dicotomique de l'indice de la fourchette contenant un compte }
var
  cFourchette : TFourchetteItem;
  iMin,
  iMax : integer;
  bTrouve : boolean;
begin
  if Count = 0 then
    begin
      Result := -1;
      Exit;
    end;

  iMin := 0;
  iMax := Count;
  Result := iMin;
  bTrouve := False;

  cFourchette := GetFourchettes(iMin);
  if (cFourchette.FCompteDebut <= Compte) and
     (GetFourchettes(iMax - 1).FCompteFin >= Compte) then
    // le compte peut être compris dans une des fourchettes triées
    repeat
      if Compte > cFourchette.FCompteFin then
        // recherche d'une fourchette plus grande
        begin
          iMin := Result;
          Result := (Result + iMax + 1) shr 1;
          cFourchette := GetFourchettes(Result);
        end;

      if Compte < cFourchette.FCompteDebut then
        // recherche d'une fourchette plus petite
        begin
          iMax := Result;
          Result := (Result + iMin) shr 1;
          cFourchette := GetFourchettes(Result);
        end;

      bTrouve := (Compte >= cFourchette.FCompteDebut) and
                 (Compte <= cFourchette.FCompteFin);
    until bTrouve or (iMin + 1 = iMax);  // si iMin + 1 = iMax, on tourne en rond

  if bTrouve then
    begin
      // recherche de la plus petite fourchette comprenant le compte
      while (Result > 0) and
            (Compte >= GetFourchettes(Result - 1).FCompteDebut) and
            (Compte <= GetFourchettes(Result - 1).FCompteFin) do
        dec(Result);
    end
  else
    Result := -1;
end;

end.

unit IntegGen;

interface

uses SysUtils, classes, Ent1;

const FMT_GENERAL = '%3.3s%3.3s%17.17s%35.35s%3.3s%1s%1s%1s%1s%1s%1s%1s';
const FMT_AUXILIAIRE = '%3.3s%3.3s%17.17s%35.35s%3.3s%1s%17.17s%303.303s%3.3s%3.3s';
const FMT_ECRITURE = '%3.3s%8.8s%2.2s%17.17s%1s%17.17s%35.35s%35.35s%3.3s%8.8s%1s%20.20s%1s%8.8s';
  { Entête Domix }
const
  EX_LGCONTRAINTE =   24;
  EX_OFFSETVALEUR =   27;
  EX_TYPE =           'TYPE DE BASE **********:';
  EX_TITRE =          'TITRE DE LA BASE*******:';
  EX_CODEACTION =     'CODE ACTION PARTIC*****:';
  EX_CODETABLE =      'CODE TABLE DE CONVERS**:';
  EX_DECOUPE =        'DECOUPE VERTICALE******:';
  EX_NBDECIMAL=       'Nbr decimal            :';
  EX_TYPERECUP=       'TYPE RECUPERATION *****:';
  EX_NBCONTRAINTE =   'CONTRAINTE NOMBRE******:';
  EX_TYPELIGNE =      'Type ligne             :';
  EX_NOMATRICE =      'Numero matrice         :';
  EX_CONTRAINTE =     'REPERE contrainte      :';
  EX_DEBUTCOMPTE =    'COMPTE NOMBRE******(C)*:';
  EX_DEBUTMOIS =      'MOIS NOMBRE********(M)*:';
  EX_DEBUTJOURNAL =   'JOURNAL NOMBRE*****(*)*:';
  EX_DEBUTECRITURE =  'ECRITURE NOMBRE****(E)*:';
  EX_CPOSNC =         'Pos. debut compte  (NC):';
  EX_CTAILLENC =      'Taille compte      (NC):';
  EX_CPOSLC =         'Pos. debut libelle (LC):';
  EX_CTAILLELC =      'Taille libelle     (LC):';
  EX_CPOSDB =         'Pos. debut debit   (DB):';
  EX_CTAILLEDB =      'Taille debit       (DB):';
  EX_CPOSCR =         'Pos. debut credit  (CR):';
  EX_CTAILLECR =      'Taille credit      (CR):';
  EX_MPOSDJ =         'Pos. debut jour    (DJ):';
  EX_MPOSDM =         'Pos. debut mois    (DM):';
  EX_MPOSDA =         'Pos. debut annee   (DA):';
  EX_MPOSMR =         'Pos. deb. mois rela(MR):';
  EX_MPOSJX =         'Pos. debut journal (JX):';
  EX_MTAILLEJX =      'Taille journal     (JX):';
  EX_JPOSJ =          'Pos. debut journal (J ):';
  EX_JTAILLEJ =       'Taille journal     (J ):';
  EX_JPOSL =          'Pos. debut libelle (L ):';
  EX_JTAILLEL =       'Taille libelle     (L ):';
  EX_EPOSDJ =         'Pos. debut jour    (DJ):';
  EX_EPOSDM =         'Pos. debut mois    (DM):';
  EX_EPOSDA =         'Pos. debut annee   (DA):';
  EX_EPOSNC =         'Pos. debut compte  (NC):';
  EX_ETAILLENC =      'Taille compte      (NC):';
  EX_EPOSLC =         'Pos. debut libelle (LC):';
  EX_ETAILLELC =      'Taille libelle     (LC):';
  EX_EPOSS =          'Pos. debut stat    (S ):';
  EX_ETAILLES =       'Taille statistique (S ):';
  EX_EPOSP =          'Pos. debut piece   (P ):';
  EX_ETAILLEP =       'Taille piece       (P ):';
  EX_EPOSDB =         'Pos. debut debit   (DB):';
  EX_ETAILLEDB =      'Taille debit       (DB):';
  EX_EPOSCR =         'Pos. debut credit  (CR):';
  EX_ETAILLECR =      'Taille credit      (CR):';
  EX_EPOSA =          'Pos. debut lettrage(A ):';
  EX_ETAILLEA =       'Taille lettrage    (A ):';
  EX_EPOSD1 =         'Pos. debut debit-1 (D1):';
  EX_ETAILLED1 =      'Taille debit-1     (D1):';
  EX_EPOSC1 =         'Pos. debut credit-1(C1):';
  EX_ETAILLEC1 =      'Taille credit-1    (C1):';
  EX_EPOSQ =          'Pos. debut quantite(Q ):';
  EX_ETAILLEQ =       'Taille quantite    (Q ):';
  EX_EPOSRG =         'Pos. debut reglemen(Rg):';
  EX_ETAILLERG =      'Taille reglement   (Rg):';
  EX_EPOSEJ =         'Pos. debut jour ech(EJ):';
  EX_EPOSEM =         'Pos. debut mois ech(EM):';
  EX_EPOSEA =         'Pos debut annee ech(EA):';
  EX_EPOSJ =          'Pos. debut journal (J ):';
  EX_ETAILLEJ =       'Taille journal     (J ):';
  EX_EPOSLJ =         'Pos deb lib compte (LJ):';
  EX_ETAILLELJ =      'Taille  lib compte (LJ):';
  { Fin Entête Domix }

type
  TIntegTypeLigne = (ilCompte, ilEcriture, ilJournal, ilMois, ilSautCompte, ilRejet, ilNeutre);
  TIntegModeFonc = (mlAutonome, mlCompta);

  TIntegInfoCompte = record
    General : string;
    Auxiliaire : string;
    Nature : string;
  end;

  TIntegCollectif = record
    General : string;
    AuxiInf : string;
    AuxiSup : string;
    Nature : string;
  end;
  PIntegCollectif = ^TIntegCollectif;

  TIntegTabConv = record
    Depart : string;
    Arrivee : string;
  end;
  PIntegTabConv = ^TIntegTabConv;

  TIntegCtxEcriture = record
    Compte : string;
    Journal : string;
    Jour : integer;
    Mois : integer;
    Annee : integer;
    MoisRel : integer;
  end;

  TIntegContexte = record
    DebEx : TDateTime;
    FinEx : TDateTime;
    EncEx : TDateTime;
    Base : string;
    FileName : string;
    LConvCpte : TList;
    LConvJal : TList;
    LCollectif : TList;
    Ecriture : TIntegCtxEcriture;
  end;

  TIntegZoneDef = record
    Pos : integer;
    Lg : integer;
    Val : variant;
  end;

function IntegExtractEntete (St : string) : string;
function IntegExtractValeur (St : string) : integer;
function IntegExtractChaine (St : string) : string;
function IntegBourreCompte (stCompte : string) : string;
procedure IntegGetCompteEquiv (Ctx : TIntegContexte; var stCompte, stAuxi : string);
procedure IntegGetCompteConv (LTabConv : TList ; var stCompte : string);
function IntegBuildNewString(var St : string; stMask, stNewMask : string) : boolean;
procedure IntegAffecteAuxiGene (L : TList; var stGene, stAuxi : string);
function IntegIntrancheAuxi (AColl : PIntegCollectif; stCompte : string) : boolean;
function IntegGetNatureCompte (L : TList; stCompte : string) : string;
procedure IntegUpdateZoneDef (var ZD : TIntegZoneDef; Position , Longueur : integer);

var IntegModeFonc : TIntegModeFonc;

implementation

function IntegExtractEntete (St : string) : string;
begin
  Result := Copy (St, 1, EX_LGCONTRAINTE);
end;

function IntegExtractValeur (St : string) : integer;
begin
  Result := StrToInt(Trim(Copy(St,EX_OFFSETVALEUR,8)));
end;

function IntegExtractChaine (St : string) : string;
begin
  Result := Trim(Copy(St,EX_OFFSETVALEUR,(Length(St) - EX_OFFSETVALEUR + 1)));
end;

procedure IntegGetCompteEquiv (Ctx : TIntegContexte; var stCompte, stAuxi : string);
begin
  IntegGetCompteConv (Ctx.LConvCpte,stCompte);
  stCompte := IntegBourreCompte(stCompte);
  IntegAffecteAuxiGene (Ctx.LCollectif,stCompte, stAuxi);
end;

function IntegBourreCompte (stCompte : string) : string;
var i, ll, LG : integer;
  Bourre : Char;
begin
  if IntegModeFonc = mlAutonome then
  begin
    Lg := 10; Bourre := '0'; // on force bourrage à 0 sur lg=10
    Result:=stCompte ; ll:=Length(Result) ;
    if ll<Lg then for i:=ll+1 to Lg do Result:=Result+Bourre ;
  end else Result := BourreLaDonc(stCompte, fbGene);
end;

procedure IntegGetCompteConv (LTabConv : TList ; var stCompte : string);
var i : integer;
  AConv : PIntegTabConv;
  stDepart : string;
begin
  for i:=0 to LTabConv.Count - 1 do
  begin
    AConv := LTabConv[i];
    stDepart := AConv^.Depart;
    if IntegBuildNewString (stCompte,AConv^.Depart, AConv^.Arrivee) then break;
  end;
end;

function IntegBuildNewString(var St : string; stMask, stNewMask : string) : boolean;
var i : integer;
  bMaskOK : boolean;
  stTmpx, stNew : string;
begin
  bMaskOK := True; stTmpx := ''; stNew := '';
  for i:=0 to Length(stMask) do
  begin
    if stMask[i]='x' then stTmpx := stTmpx+St[i]
    else if St[i] <> stMask[i] then
    begin
      bMaskOK := False;
      break;
    end;
  end;
  if bMaskOK then
  begin
    for i:=0 to Length(stNewMask) do
    begin
      if stNewMask[i]='x' then
      begin
        stNew := stNew+stTmpx[1];
        Delete(stTmpx,1,1);
      end else stNew := stNew+stNewMask[i];
    end;
    St := stNew;
  end;
  Result := bMaskOK;
end;

procedure IntegAffecteAuxiGene (L : TList; var stGene, stAuxi : string);
var stCompte : string;
  i : integer;
  AColl : PIntegCollectif;
begin
  stCompte := stGene;
  for i:=0 to L.Count - 1 do
  begin
    AColl := L[i];
    if IntegIntrancheAuxi (L[i],stCompte) then
    begin
      stAuxi := stCompte;
      stGene := AColl^.General;
      break;
    end;
  end;
end;

function IntegIntrancheAuxi (AColl : PIntegCollectif; stCompte : string) : boolean;
begin
  Result := (stCompte >= AColl^.AuxiInf) and (stCompte <= AColl^.AuxiSup);
end;

function IntegGetNatureCompte (L : TList; stCompte : string) : string;
var i : integer;
  AColl : PIntegCollectif;
  var stNature : string;
begin
  stNature := 'DIV';
  for i:=0 to L.Count - 1 do
  begin
    AColl := L[i];
    if (AColl^.General = stCompte) then
    begin
      stNature := AColl^.Nature;
      break;
    end else if IntegIntrancheAuxi (AColl,stCompte) then
    begin
      if AColl^.Nature='COC' then stNature := 'CLI' // Client
      else if AColl^.Nature='COF' then stNature := 'FOU' // Fournisseur
      else stNature := 'DIV';
      break;
    end;
  end;
  Result := stNature;
end;

procedure IntegUpdateZoneDef (var ZD : TIntegZoneDef; Position , Longueur : integer);
begin
  ZD.Pos := Position;
  ZD.Lg := Longueur;
end;

end.

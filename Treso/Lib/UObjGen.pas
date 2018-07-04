unit UObjGen;

interface

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
  {$ENDIF}
  sysutils, Ent1, hctrls , Constantes, Hent1,
  UTOB, ParamSoc, Classes, Math, Commun, UProcGen;

type
  {Objet pour le calcul des dates des valeurs}
  TDateValeur = class (TObject)
  private
    tCodeCIB  : TOB;
    tModePaie : TOB;
    lCorresp  : TStringList;
    lDateVal  : TStringList;
    procedure GetTobCib;
    procedure GetCorrespondance;
  public
    constructor Create;
    destructor  Destroy; override;
    procedure RecupConditionValeur;
    procedure GetDateValeur(General, ModePaie, Sens : string; var dtValeur : TDateTime; var CIB : string);
    procedure AjouterALaListe(BQ, Gen, Paie, Sens : string);

    property CodeCIB  : TOB         read tCodeCIB  write tCodeCIB;
    property ModePaie : TOB         read tModePaie write tModePaie;
    property Corresp  : TStringList read lCorresp  write lCorresp;
  end;

  {Objet pour l'affectation / récupération des code rubriques pour les
   écritures importées de la comptabilité}
  TObjRubrique = class (TObject)
  private
    tRubrique : TOB;
    tGeneraux : TOB;
    tCibCour  : TOB;
    lCorresp  : TStringList;

    function  GetTobCibCour : TOB; {09/08/06 : Geter de la property CibCour}
    procedure GetTobRubrique;
    {08/03/05 : Déplacer en fonction globale, car utilisée aussi dans ULibBonAPayer
    function  CompareGeneral(AComparer, Gen, Excpt : string) : Boolean;}
  public
    constructor Create;
    destructor  Destroy; override;

    procedure SetListeCorrespondance;
    procedure AjouterALaListe(Gen : string);
    procedure GetCorrespondance  (var Cpte, Cod, Lib : string; Sens, Paie : string);
    {09/08/06 : Mise en place d'une procédure spéciale pour les comptes courants}
    procedure GetCorrespondanceMS(var Cpte, Cod, Lib : string; Sens, Paie : string);

    property Rubrique : TOB         read tRubrique write tRubrique;
    property Generaux : TOB         read tGeneraux write tGeneraux;
    property CibCour  : TOB         read GetTobCibCour;
    property Corresp  : TStringList read lCorresp  write lCorresp;
  end;

  {Objet de gestion de la TVA pour les écritures à intégrer dans la comptabilité}
  TObjTVA = class (TObject)
  private
    tTVA      : TOB;
    tCompte   : TOB;
    tGeneraux : TOB;
    tFluxTres : TOB;
    lCorresp  : TStringList;

    function  IsAchat(Cpte : string) : Boolean;

    procedure GetTobFluxTreso;
    procedure GetTobGeneraux;
    procedure GetTobTVA;
    procedure GetTobCpteBancaire;
    procedure SetListeCorrespondance;
  public
    constructor Create;
    destructor  Destroy; override;

    procedure GetCorrespondance(var Obj : TObjDetailTVA; Cpte, Flx : string);
    property Corresp  : TStringList read lCorresp  write lCorresp;
  end;

  {Objet stockant le cours des taux d'intérêts}
  TObjTaux = class (TObject)
  private
    lTaux    : TStringList;
    FDateDeb : TDateTime;
    FDateFin : TDateTime;

    procedure ChargerListeTaux(MaxOk : Boolean);
  public
    constructor Create(DateDeb, DateFin : TDateTime; MaxOk : Boolean = False);
    destructor  Destroy; override;

    procedure   ChangerDates(DateDeb, DateFin : TDateTime; MaxOk : Boolean = False; ClearOk : Boolean = False);
    function    GetTaux(Taux, EnDate : string) : Double;

    property DateDeb : TDateTime read FDateDeb write FDateDeb;
    property DateFin : TDateTime read FDateFin write FDateFin;
  end;

  {Objet de gestion des devises pour les écritures de trésorerie à générer}
  TObjDevise = class (TObject)
  private
    tBanque   : TOB;
    lCorresp  : TStringList;
    lListeDev : TStringList;
    FDateValeur : TDateTime;

    procedure GetTobBanque;
    procedure SetListeCorrespondance;
    procedure GetTauxDevise(var Taux : Double; Dev : string; Dt : TDateTime = 1);
  public
    constructor Create(DtValeur : TDateTime; AvecComptes : Boolean = True);
    destructor  Destroy; override;

    procedure ConvertitMnt(var MntE, MntD : Double; DevOrigine, Compte : string; Dt : TDateTime = 1);
    {06/12/04 : Renvoie le montant dans la devise du compte à partir d'un montant en devise pivot}
    function  GetMntDevFromEur(MntE : Double; Cpte : string; Dt : TDateTime = 1) : Double;
    function  GetDeviseCpt    (Compte : string) : string;
    {06/12/04 : Renvoie le montant à mettre dans les Champs _COTATION}
    function  Get_Cotation    (Cpte : string; Dt : TDateTime) : Double;
    function  GetParite       (DevOrig, DevDest : string) : Double;
    {23/05/05 : Retourne le nombre de décimales en fonction de la devise}
    function  GetNbDecimalesFromDev(Dev : string) : Integer;
    {23/05/05 : Retourne le nombre de décimales en fonction du compte}
    function  GetNbDecimalesFromCpt(Cpt : string) : Integer;

    property Corresp  : TStringList read lCorresp    write lCorresp;
    property ListeDev : TStringList read lListeDev   write lListeDev;
    property DateValeur : TDateTime read FDateValeur write FDateValeur;
  end;

  {Objet de gestion des ParamSoc des banques prévisionnelles}
  TObjBanquePrevi = class
  private
    ParamAbst  : Boolean;
    ParamValid : TValideNatCpte;
    ParamLibel : TLibelleParam;
    ParamCptes : TComptesParam;
    LBanqPrevi : TStringList;
    LBnqFiltre : TStringList;

    procedure ChargerParam(FromDisk : Boolean);
    procedure MajTableaux;
    procedure MajListeFiltre(Valeur : string);
  public
    NoDossier : string;
    constructor Create;
    destructor  Destroy; override;
    procedure Recharger (FromDisk : Boolean);
    function IndToNature(Ind : Byte  ) : string;
    function NatToIndice(Nat : string) : Byte;
    function GetLibelle (Ind : Byte  ) : string;
    function GetCompte  (Ind : Byte  ) : string;
    {Récupération des infos du Rib}
    function GetEtabBq  (Cpt : string) : string;
    function GetGuichet (Cpt : string) : string;
    function GetNumCpt  (Cpt : string) : string;
    function GetCleRib  (Cpt : string) : string;
    function GetIban    (Cpt : string) : string;
    function GetInfosBq (Cpt : string) : TObjInfosBque;

    function IsValide   (Ind : Byte)                : Boolean;
    function IsAllOk    (var Msg : string)          : Boolean;
    function IsFiltre   (BqPrevi, NatGene : string) : Boolean;

    property BqFiltre : string write MajListeFiltre;
  end;

  {12/10/04 : Objet permettant de récupérer un cib en fonction de son mode paiement et inversement}
  TObjCIBModePaie = class
  private
    tCib : TOB;
    tBqe : TOB;
    procedure ChargerTOB;
    function  GetBanque(General : string) : string;
  public
    constructor Create;
    destructor  Destroy; override;
    function GetCIB     (General, ModePaie, Sens : string) : string;
    function GetModePaie(General, CodeCib , Sens : string) : string;
  end;

  {Cet objet stocke les comptes issus de BANQUECP et gère leur correspondance :
   1/ BQ_CODE <-> G_GENERAL
   2/ NODOSSIER <-> CLIENSSOC ....}
  TobjComptesTreso = class
  private
    TobBqe : TOB;
    FDosSource : string;
    FDosDest   : string;
    FEstMulti  : Boolean;
    CptPoubelle: string;

    procedure ChargeLesComptes;
  public
    constructor Create(DosSource, DosDest : string);
    destructor  Destroy; override;
    function    GetCouFromGene(General, NoDossier : string) : string;
    function    GetCouFromDoss(DosSource, DosDest : string) : string;
    function    GetGeneFromCode(BqCode : string) : string;
    procedure   SetLeRib(var T : TOB);
  end;

  {Objet pour la saisie automatique lors du pointage à partir des TRGUIDEs}
  TobjSaisieAuto = class

  private
    FTobReleve : TOB;
    FTobGuide  : TOB;
    FCompte    : string;
    FDevise    : string;
    FRefPtage  : string;
    FDatePtage : TDateTime;

    {$IFDEF TRESO}
    FCurGuide  : TOB;
    function RechercheGuide(LigReleve : TOB) : Boolean;
    {$ENDIF TRESO}
  public
    TobEcriture : TOB;
    constructor Create(TobReleve : TOB; Cpte, Ref, Devise : string; DatePtg : TDateTime);
    destructor  Destroy; override;

    procedure GenereEcritures;
  end;

  {Objet générant automatiquement un code alpha numérique unique sur 3 caractères en fonction *
   de la Table passée en paramètre dans le Create}
  TObjCodeCombo = class
  private
    FTobCode  : TOB;
    FNomTable : string;
    FNomCode  : string;
    FLastCode : string;
  public
    class procedure GetCodeComboUnique(var CodeCombo : string);
    constructor     Create(NomTable, ChampCode  : string);
    destructor      Destroy; override;
    function        GetNewCode : string;
    
    property LastCode : string read FLastCode write FLastCode;
  end;



  TGrpObjet = record
    oBqPrev   : TObjBanquePrevi;
    oCibMp    : TObjCIBModePaie;
    oDevise   : TObjDevise;
    oTaux     : TObjTaux;
    oTVA      : TObjTVA;
    oRub      : TObjRubrique;
    oValeur   : TDateValeur;
    iCode     : Integer;
    sChaine   : string;
    sGene     : string;
    dDate     : TDatetime;
    NomBase   : string;
    NoDossier : string;
    Societe   : string;
    ListSolde : TStringList;
    TobRejet  : TOB;
    TobDate   : TOB; {20/11/06}
  end;

{JP 08/03/06 : AComparer : En général un compte général justement
               Gen : contient des racine de compte généraux
               Excpt : contient des racine de compte généraux à exclure par rapport à Gen
               Renvoie True si AComparer appartient à Gen et pas à Excpt :
               Utiliser en Tréso pour les rubriques, dans les bons à payer sur les comptes
               de charge à traiter ...}
function CompareGeneral(AComparer, Gen, Excpt : string) : Boolean;


implementation

uses
  {$IFNDEF EAGLSERVER}
  AglInit,
  {$ENDIF EAGLSERVER}

  {$IFDEF TRESO}
  UProcSolde,
  UProcEcriture,
  {$ENDIF TRESO}

  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}

  HMsgBox,
  UlibWindows;


{JP 08/03/06 : AComparer : En général un compte général justement
               Gen : contient des racine de compte généraux
               Excpt : contient des racine de compte généraux à exclure par rapport à Gen
               Renvoie True si AComparer appartient à Gen et pas à Excpt :
               Utiliser en Tréso pour les rubriques, dans les bons à payer sur les comptes
               de charge à traiter ...}
{---------------------------------------------------------------------------------------}
function CompareGeneral(AComparer, Gen, Excpt : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Part : string;
  Lgth : Byte;
  s, c, a  : string;
  APrendre : Boolean;
begin
  Result := False;
  {04/05/04 : Deux possibilités : ou Gen est une racine de compte ...}
  if Pos(':', Gen) = 0 then begin
    {On récupère le début du compte général pour le comparer avec la "racine" de comptes
    paramétrée dans la rubrique}
    Lgth := Length(Gen);
    Part := Copy(AComparer, 1, Lgth);
    APrendre := Part = Gen;
  end
  {... ou bien Gen est une fourchette de compte}
  else begin
    a    := AComparer;
    s    := Gen;
    c    := ReadTokenPipe(s, ':');
    Part := ReadTokenPipe(s, ':');
    {Recherche de la chaine la plus longue}
    Lgth := Max(Length(AComparer), Max(Length(c), Length(Part)));
    {On complète les trois chaines pour qu'elles soient de la même taille}
    c    := PadR(c   , '0', Lgth);
    Part := PadR(Part, '9', Lgth);
    a    := PadR(a   , '1', Lgth);
    {Si le compte général appartient à la fourchette ...}
    APrendre := (c <= a) and (a <= Part);
  end;

  {Le compte général "fait bien partie" de la rubrique. On regarde maintenant s'il ne
   fait pas partie des comptes à exclure}
  if APrendre then begin
    {Par défaut on considère que ce compte est rattaché  à la rubrique}
    Result := True;
    {23/11/04 : FQ 10178 : gestion de la virgule qui sert de séparateurb entre les différents
                comptes ou fourchettes de comptes à exclure cf. TOM_RUBRIQUE.AnalyseCompteOk}
    s := ReadTokenPipe(Excpt, ',');
    {23/11/04 : on boucle tant qu'il y a des compte et que Acomparer n'est pas à exclure}
    while (s <> '') and Result do begin
      {Si on trouve ":", c'est que l'on a affaire à une fourchette de compte}
      if (Pos(':', s) > 0) then begin
        a    := AComparer;
        c    := ReadTokenPipe(s, ':');
        Part := ReadTokenPipe(s, ':');
        {Recherche de la chaine la plus longue}
        Lgth := Max(Length(AComparer), Max(Length(c), Length(Part)));
        {On complète les trois chaines pour qu'elles soient de la même taille}
        c    := PadR(c   , '0', Lgth);
        Part := PadR(Part, '9', Lgth);
        a    := PadR(a   , '1', Lgth);
        {Si le compte général appartient à la fourchette ce compte n'est donc pas rattaché à la rubrique}
        Result := not ((c <= a) and (a <= Part));
      end

      {Sinon, s'il y a un compte d'exclusion, on teste le compte par rapport à la "racine" d'exclusion}
      else if Trim(s) <> '' then begin
        Lgth := Length(s);
        Part := Copy(AComparer, 1, Lgth);
        {Si le compte général appartient à la "racine" d'exclusion}
        Result := not (Part = s);
      end;

      {On récupère la prochaine exclusion de comptes}
      s := ReadTokenPipe(Excpt, ',');
    end;{while s <> ''}
  end;{if APrendre}
end;
            {***************************************************************}
            {**********************    TDateValeurs   **********************}
            {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TDateValeur.Create;
{---------------------------------------------------------------------------------------}
begin
  tCodeCIB  := TOB.Create('$CIB', nil, -1);
  tModePaie := TOB.Create('$CHQ', nil, -1);
  lCorresp  := TStringList.Create;
  lCorresp.Duplicates := dupIgnore;
  lCorresp.Sorted     := True;
  lDateVal  := TStringList.Create;
  lDateVal.Duplicates := dupIgnore;
  lDateVal.Sorted     := True;
end;

{---------------------------------------------------------------------------------------}
destructor TDateValeur.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tCodeCIB)  then FreeAndNil(tCodeCIB);
  if Assigned(tModePaie) then FreeAndNil(tModePaie);
  LibereListe(lCorresp, True);
  LibereListe(lDateVal, True);

  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TDateValeur.RecupConditionValeur;
{---------------------------------------------------------------------------------------}
begin
  GetTobCib;
  GetCorrespondance;
end;

{On calcule et stocke la date de valeur à partir du mode de paiement et du sens d'une opération
 dtValeur est en entrée de procedure la date comptable : si on ne trouve pas les informations de
          conversion, dtValeur ne sera pas modifiée et par donc la date de valeur sera la date
          comptable.
{---------------------------------------------------------------------------------------}
procedure TDateValeur.GetDateValeur(General, ModePaie, Sens : string; var dtValeur : TDateTime; var CIB : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  k : Integer;
  O : TObjCibTanque;
  v : TObjDtValeur;
  s : string;
begin
  {Le nom de chaque ligne de la liste est composé de "General; ModePaie"}
  n := lCorresp.IndexOf(General + ';' + ModePaie + ';' + Sens);
  {Si le mode de paiement n'est pas paramétré, la date de valeur sera la date comptable
   et le code CIB de TRECRITURE sera vide}
  if n < 0 then Exit;

  O := TObjCibTanque(lCorresp.Objects[n]);
  CIB := O.CodeCIB;

  {Pour éviter un trop grand nombre de requêtes, on stocke les dates de valeur ...}
  s := DateToStr(dtValeur);
  {On regarde si la date de valeur pour ce compte, ce CIB et cette date comptable a déjà été calculée ...}
  k := lDateVal.IndexOf(General + ';' + O.CodeCIB + ';' + s);

  if k < 0 then begin
    {... Si ce n'est pas le cas, on la calcule ...}
    dtValeur := CalcDateValeur(O.CodeCIB, General, dtValeur);
    v := TObjDtValeur.Create;
    v.DateVal := dtValeur;
    {... et on la stocke}
    lDateVal.AddObject(General + ';' + O.CodeCIB + ';' + s, v);
  end
  else
    {... sinon, on se contente de récupérer la date déjà calculée}
    dtValeur := TObjDtValeur(lDateVal.Objects[k]).DateVal;
end;

{Chargement de tous les codes CIB paramétrés
{---------------------------------------------------------------------------------------}
procedure TDateValeur.GetTobCib;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TCI_CODECIB, TCI_BANQUE, TCI_MODEPAIE, TCI_SENS FROM CIB ' +
               'WHERE (TCI_MODEPAIE <> "" AND TCI_MODEPAIE IS NOT NULL) AND ' +
               '      (TCI_SENS <> "" AND TCI_SENS IS NOT NULL)', True);
  tCodeCIB.LoadDetailDB('CIB', '', '', Q, False);
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
procedure TDateValeur.AjouterALaListe(BQ, Gen, Paie, Sens : string);
{---------------------------------------------------------------------------------------}
var
  T   : TOB;
  Obj : TObjCibTanque;
begin
  T := tCodeCIB.FindFirst(['TCI_MODEPAIE', 'TCI_BANQUE', 'TCI_SENS'], [Paie, Bq, Sens], False);

  {S'il n'y a pas de code CIB de paramétrés pour le sens, on regarde "MIX" }
  if not Assigned(T) and (Sens <> 'MIX') then
    T := tCodeCIB.FindFirst(['TCI_MODEPAIE', 'TCI_BANQUE', 'TCI_SENS'], [Paie, Bq, 'MIX'], False);

  {S'il n'y a pas de code CIB de paramétrés pour la banque, on regarde dans la "banque DEFAUT" }
  if not Assigned(T) then
    T := tCodeCIB.FindFirst(['TCI_MODEPAIE', 'TCI_BANQUE', 'TCI_SENS'], [Paie, CODECIBREF, Sens], False);

  {S'il n'y a pas de code CIB de paramétrés pour le sens, on regarde "MIX" }
  if not Assigned(T) and (Sens <> 'MIX') then
    T := tCodeCIB.FindFirst(['TCI_MODEPAIE', 'TCI_BANQUE', 'TCI_SENS'], [Paie, CODECIBREF, 'MIX'], False);

  if Assigned(T) then begin
    Obj := TObjCibTanque.Create;
    Obj.CodeCIB := T.GetString('TCI_CODECIB');
    Obj.CodeBQE := Bq;
    lCorresp.AddObject(Gen + ';' + Paie + ';' + Sens, Obj);
  end;
end;

{Constitue la StringList de correspondance entre un compte général et son code CIB
{---------------------------------------------------------------------------------------}
procedure TDateValeur.GetCorrespondance;
{---------------------------------------------------------------------------------------}
var
  n    : Integer;
  B,
  S,
  M, C : string;
begin
  {On vide les listes}
  LibereListe(lCorresp, False);
  LibereListe(lCorresp, False);

  if not Assigned(tModePaie) then FreeAndNil(tModePaie);

  {27/11/2003 : A la demande d'Olivier Guillaud, ajout de la gestion du sens}
  {On balaie la liste avec les divers modes de paiement nécessaires}
  tCodeCIB.Detail.Sort('TCI_BANQUE;TCI_MODEPAIE;TCI_SENS');
  for n := 0 to tModePaie.Detail.Count - 1 do begin
    {31/05/06 : Ajout d'un préfixe aux alias pour éviter une erreur en multi-sociétés}
    S := tModePaie.Detail[n].GetString('AAA_SENS');
    M := tModePaie.Detail[n].GetString('AAA_MODEPAIE');
    B := tModePaie.Detail[n].GetString('AAA_BANQUE');
    C := tModePaie.Detail[n].GetString('AAA_GENERAL');
    AjouterALaListe(B, C, M, S);
  end;
end;

          {***************************************************************}
          {**********************    TObjRubrique   **********************}
          {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TObjRubrique.Create;
{---------------------------------------------------------------------------------------}
begin
  tGeneraux := TOB.Create('$GEN', nil, -1);
  tRubrique := TOB.Create('$RUB', nil, -1);

  {Récupération des rubriques de Tréso}
  GetTobRubrique;
  {09/08/06 : Cette Tob ne sera créée que lors du premier appel à GetCorrespondanceMS}
  tCibCour := nil;

  lCorresp  := TStringList.Create;
  lCorresp.Duplicates := dupIgnore;
  lCorresp.Sorted     := True;
end;

{---------------------------------------------------------------------------------------}
destructor TObjRubrique.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tGeneraux) then FreeAndNil(tGeneraux);
  if Assigned(tRubrique) then FreeAndNil(tRubrique);
  if Assigned(tCibCour ) then FreeAndNil(tCibCour);
  LibereListe(lCorresp, True);

  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjRubrique.GetTobRubrique;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  {04/05/04 : ajout du signe de la rubrique dans la requête
   27/09/04 : FQ 10119 : ajout de RB_EXCLUSION2 dans lequel sont stockés les éventuels modes de paiement}
  Q := OpenSQL('SELECT RB_RUBRIQUE, RB_LIBELLE, RB_COMPTE1, RB_EXCLUSION1, RB_EXCLUSION2, RB_SIGNERUB ' +
               'FROM RUBRIQUE WHERE RB_NATRUB = "TRE" AND RB_CLASSERUB = "TRE"', True);
  tRubrique.LoadDetailDB('RUBRIQUE', '', '', Q, False);
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function TObjRubrique.GetTobCibCour : TOB;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(tCibCour) then begin
    tCibCour := TOB.Create('****', nil, -1);
    tCibCour.LoadDetailFromSQL('SELECT TCI_CODECIB, TCI_MODEPAIE FROM CIB WHERE TCI_BANQUE = "' + CODECOURANTS + '"');
  end;
  Result := tCibCour;
end;

{Génère la liste de correspondance Comptes - Sens/Rubriques pour les comptes généraux demandés
 JP 04/05/04 : et en fonction de leur sens à partir de la 01.50.001.001 FQ 10062
 (Ils sont affectés par l'appelant dans tGeneraux)
FQ 10119 : gestion des modes de paiement qui sont stockés dans le champ RB_EXCLUSION2
{---------------------------------------------------------------------------------------}
procedure TObjRubrique.SetListeCorrespondance;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  Gen : string;
begin
  {On vide les listes}
  LibereListe(lCorresp, False);

  if not Assigned(tGeneraux) then Exit;

  for n := 0 to tGeneraux.Detail.Count - 1 do begin
    Gen := tGeneraux.Detail[n].GetString('GENERAL');
    AjouterALaListe(Gen);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjRubrique.AjouterALaListe(Gen : string);
{---------------------------------------------------------------------------------------}
var
  Obj : TObjLibRub;
  p   : Integer;
  Sens, MP,
  pGen, pGenDet,
  pEx , pExDet : string;
  ModePaiement : string;
begin
  for p := 0 to tRubrique.Detail.Count - 1 do begin
    pGen := tRubrique.Detail[p].GetString('RB_COMPTE1');
    pEx  := tRubrique.Detail[p].GetString('RB_EXCLUSION1');
    {JP 04/05/04 : Le sens qui pouvait être saisi sur la fiche des rubriques n'était pas utilisé : maintenant un
                   même compte pourra être rattaché à deux rubriques en fonction de son sens (positif ou négatif)}
    Sens := tRubrique.Detail[p].GetString('RB_SIGNERUB');

    pGenDet := ReadTokenSt(pGen);
    pExDet  := ReadTokenSt(pEx );
    while pGenDet <> '' do begin
      {On regarde si le compte général fait partie de la rubrique}
      if CompareGeneral(Gen, pGenDet, pExDet) then begin
        ModePaiement := tRubrique.Detail[p].GetString('RB_EXCLUSION2');
        {27/09/04 : Ajout du mode de paiement, on crée autant d'objets qu'il y a de mode paiement}
        MP := ReadTokenSt(ModePaiement);
        repeat
          {13/09/06 : Ne devrait pas arriver, mais comme j'ai des fuites mémoires ...}
          if lCorresp.IndexOf(Gen + '|' + Sens + '|' + MP) = -1 then begin
            Obj := TObjLibRub.Create;
            Obj.CodRub := tRubrique.Detail[p].GetString('RB_RUBRIQUE');
            Obj.LibRub := tRubrique.Detail[p].GetString('RB_LIBELLE');
            lCorresp.AddObject(Gen + '|' + Sens + '|' + MP, Obj);
          end;
          MP := ReadTokenSt(ModePaiement);
        until Trim(MP) = '';

        Break;
      end;
      pGenDet := ReadTokenSt(pGen);
      pExDet  := ReadTokenSt(pEx );
    end;
  end;
end;

(* 08/03/05 : Déplacer dans ULibWindows, car utilisée aussi dans ULibBonAPayer
{On teste si un compte général (AComparer) est rattaché à une rubrique, c'est à dire s'il
 a la même racine que Gen et s'il ne fait pas partie des comptes à exclure de la "racine"
 JP 04/05/04 : FQ 10052 : maintenant la partie à comparer pourra être aussi une fourchette
               de rubriques
{---------------------------------------------------------------------------------------}
function TObjRubrique.CompareGeneral(AComparer, Gen, Excpt : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Part : string;
  Lgth : Byte;
  s, c, a  : string;
  APrendre : Boolean;
begin
  Result := False;
  {04/05/04 : Deux possibilités : ou Gen est une racine de compte ...}
  if Pos(':', Gen) = 0 then begin
    {On récupère le début du compte général pour le comparer avec la "racine" de comptes
    paramétrée dans la rubrique}
    Lgth := Length(Gen);
    Part := Copy(AComparer, 1, Lgth);
    APrendre := Part = Gen;
  end
  {... ou bien Gen est une fourchette de compte}
  else begin
    a    := AComparer;
    s    := Gen;
    c    := ReadTokenPipe(s, ':');
    Part := ReadTokenPipe(s, ':');
    {Recherche de la chaine la plus longue}
    Lgth := Max(Length(AComparer), Max(Length(c), Length(Part)));
    {On complète les trois chaines pour qu'elles soient de la même taille}
    c    := PadR(c   , '0', Lgth);
    Part := PadR(Part, '9', Lgth);
    a    := PadR(a   , '1', Lgth);
    {Si le compte général appartient à la fourchette ...}
    APrendre := (c <= a) and (a <= Part);
  end;

  {Le compte général "fait bien partie" de la rubrique. On regarde maintenant s'il ne
   fait pas partie des comptes à exclure}
  if APrendre then begin
    {Par défaut on considère que ce compte est rattaché  à la rubrique}
    Result := True;
    {23/11/04 : FQ 10178 : gestion de la virgule qui sert de séparateurb entre les différents
                comptes ou fourchettes de comptes à exclure cf. TOM_RUBRIQUE.AnalyseCompteOk}
    s := ReadTokenPipe(Excpt, ',');
    {23/11/04 : on boucle tant qu'il y a des compte et que Acomparer n'est pas à exclure}
    while (s <> '') and Result do begin
      {Si on trouve ":", c'est que l'on a affaire à une fourchette de compte}
      if (Pos(':', s) > 0) then begin
        a    := AComparer;
        c    := ReadTokenPipe(s, ':');
        Part := ReadTokenPipe(s, ':');
        {Recherche de la chaine la plus longue}
        Lgth := Max(Length(AComparer), Max(Length(c), Length(Part)));
        {On complète les trois chaines pour qu'elles soient de la même taille}
        c    := PadR(c   , '0', Lgth);
        Part := PadR(Part, '9', Lgth);
        a    := PadR(a   , '1', Lgth);
        {Si le compte général appartient à la fourchette ce compte n'est donc pas rattaché à la rubrique}
        Result := not ((c <= a) and (a <= Part));
      end

      {Sinon, s'il y a un compte d'exclusion, on teste le compte par rapport à la "racine" d'exclusion}
      else if Trim(s) <> '' then begin
        Lgth := Length(s);
        Part := Copy(AComparer, 1, Lgth);
        {Si le compte général appartient à la "racine" d'exclusion}
        Result := not (Part = s);
      end;

      {On récupère la prochaine exclusion de comptes}
      s := ReadTokenPipe(Excpt, ',');
    end;{while s <> ''}
  end;{if APrendre}
end;
*)

{Renvoie le code et le libellé de la rubrique rattachée à un compte général
{---------------------------------------------------------------------------------------}
procedure TObjRubrique.GetCorrespondance(var Cpte, Cod, Lib : string; Sens, Paie : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  n := lCorresp.IndexOf(Cpte + '|' + Sens + '|' + Paie);
  if n > -1 then begin
    Cod := TObjLibRub(lCorresp.Objects[n]).CodRub;
    Lib := TObjLibRub(lCorresp.Objects[n]).LibRub;
  end
  else begin
    n := lCorresp.IndexOf(Cpte + '|' + Sens + '|');
    if n > -1 then begin
      Cod := TObjLibRub(lCorresp.Objects[n]).CodRub;
      Lib := TObjLibRub(lCorresp.Objects[n]).LibRub;
    end
    else begin
      Cod := '';
      Lib := '';
    end;
  end;
end;

{09/08/06 : Mise en place d'une procédure spéciale pour les comptes courants
            On va boucler sur tous les modes de paiement du CIB pour trouver s'il
            existe une rubriqe correspondant au compte et au sens de l'écriture
{---------------------------------------------------------------------------------------}
procedure TObjRubrique.GetCorrespondanceMS(var Cpte, Cod, Lib : string; Sens, Paie : string);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  {On boucle sur tous les modes de paiement correspondant au Cib passé en paramètres ...}
  T := CibCour.FindFirst(['TCI_CODECIB'], [Paie], True);
  while Assigned(T) do begin
    {... et pour chaque mode de paiement, appel à la fonction standard comme fait par ailleur dans AccImportTreso}
    GetCorrespondance(Cpte, Cod, Lib, Sens, T.GetString('TCI_MODEPAIE'));
    if Cod = '' then begin
      AjouterALaListe(Cpte);
      GetCorrespondance(Cpte, Cod, Lib, Sens, T.GetString('TCI_MODEPAIE'));
      if Cod <> '' then Break;
    end
    else
      Break;
    T := CibCour.FindNext(['TCI_CODECIB'], [Paie], True);
  end;
end;

          {***************************************************************}
          {**********************       TObjTVA     **********************}
          {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TObjTVA.Create;
{---------------------------------------------------------------------------------------}
begin
  tGeneraux := TOB.Create('$GEN', nil, -1);
  tCompte   := TOB.Create('$BAN', nil, -1);
  tFluxTres := TOB.Create('$FLU', nil, -1);
  tTVA      := TOB.Create('$TVA', nil, -1);
  {Liste permettant de récupérer le taux de TVA et le compte général de tva à partir d'un compte
   bancaire et d'un code flux cf GetCorrespondance}
  lCorresp  := TStringList.Create;
  lCorresp.Duplicates := dupIgnore;
  lCorresp.Sorted     := True;

  {Récupération des Flux de Tréso qui demande une gestion de la TVA}
  GetTobFluxTreso;
  {Récupération des comptes concernés par les flux récupérés ci-dessus}
  GetTobGeneraux;
  {Récupération des taux et des comptes de tva}
  GetTobTVA;
  {Récupération des comptes bancaires pour connaître leur regime TVA (rattaché à l'agence)}
  GetTobCpteBancaire;
  {Constitution de la liste de correspondance CodeFlux-CompteGeneBancaire / TauxTva-CompteGenTva}
  SetListeCorrespondance;
end;

{---------------------------------------------------------------------------------------}
destructor TObjTVA.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tCompte)   then FreeAndNil(tCompte);
  if Assigned(tGeneraux) then FreeAndNil(tGeneraux);
  if Assigned(tFluxTres) then FreeAndNil(tFluxTres);
  if Assigned(tTVA     ) then FreeAndNil(tTVA);
  LibereListe(lCorresp, True);
  inherited Destroy;
end;

{Constitution de la tob des flux de trésorerie soumis à TVA
{---------------------------------------------------------------------------------------}
procedure TObjTVA.GetTobFluxTreso;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TFT_FLUX, TFT_GENERAL, TFT_TVAENCAIS, TFT_CODETVA FROM FLUXTRESO ' +
               'WHERE TFT_ASSUJETTITVA = "X"', True);
  tFluxTres.LoadDetailDB('FLUXTRESO', '', '', Q, False);
  Ferme(Q);
end;

{Constitution de la tob des comptes génraux affectés à des flux de trésorerie qui gère la TVA
{---------------------------------------------------------------------------------------}
procedure TObjTVA.GetTobGeneraux;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT G_GENERAL, G_NATUREGENE FROM GENERAUX WHERE G_NATUREGENE IN ("CHA", "PRO")', True);
  tGeneraux.LoadDetailDB('GENERAUX', '', '', Q, False);
  Ferme(Q);
end;

{Constitution de la tob TVA
{---------------------------------------------------------------------------------------}
procedure TObjTVA.GetTobTVA;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT * FROM TXCPTTVA WHERE TV_REGIME IN (SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="RTV")', True);
  tTVA.LoadDetailDB('GENERAUX', '', '', Q, False);
  Ferme(Q);
end;

{Constitution de la tob des agences bancaires et de leur regime TVA
{---------------------------------------------------------------------------------------}
procedure TObjTVA.GetTobCpteBancaire;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT BQ_CODE GENERAL, TRA_REGIMETVA REGIMETVA FROM BANQUECP ' + 
               'LEFT JOIN AGENCE ON TRA_AGENCE = BQ_AGENCE', True);
  tCompte.LoadDetailDB('BQE', '', '', Q, False);
  Ferme(Q);
end;

{Génère la liste de correspondance Comptes bancaires/Taux et comptes TVA
{---------------------------------------------------------------------------------------}
procedure TObjTVA.SetListeCorrespondance;
{---------------------------------------------------------------------------------------}
var
  Obj  : TObjDetailTVA;
  n, p : Integer;
  T    : TOB;
  Ach,
  Enc  : Boolean;
  Reg,
  CBq,
  CTv,
  Tau,
  Gen,
  Flx  : string;
begin
  {On vide les listes}
  LibereListe(lCorresp, False);

  {On boucle sur la liste des flux soumis à TVA}
  for n := 0 to tFluxTres.Detail.Count - 1 do begin
    Flx := tFluxTres.Detail[n].GetString('TFT_FLUX');
    CTv := tFluxTres.Detail[n].GetString('TFT_CODETVA');
    Enc := tFluxTres.Detail[n].GetString('TFT_TVAENCAIS') = 'X';
    Ach := IsAchat(tFluxTres.Detail[n].GetString('TFT_GENERAL'));
    {On récupère le nom des champs à utiliser dans la table TXCPTTVA}
    if Ach then begin
      Tau := 'TV_TAUXACH';
      if Enc then Gen := 'TV_ENCAISACH'
             else Gen := 'TV_CPTEACH';
    end
    else begin
      Tau := 'TV_TAUXVTE';
      if Enc then Gen := 'TV_ENCAISVTE'
             else Gen := 'TV_CPTEVTE';
    end;

    {On boucle sur les comptes bancaires pour récupérer leur régime TVA}
    for p := 0 to tCompte.Detail.Count - 1 do begin
      Reg := tCompte.Detail[p].GetString('REGIMETVA');
      CBq := tCompte.Detail[p].GetString('GENERAL');
      if Trim(Reg) = '' then Continue;
      T := tTVA.FindFirst(['TV_TVAOUTPF','TV_CODETAUX','TV_REGIME'], ['TX1', CTv, Reg], True);
      if not Assigned(T) then Continue;

      {On récupère le taux de TVA et le compte général sur lequel imputé la TVA}
      Obj := TObjDetailTVA.Create;
      Obj.Taux := T.GetDouble(Tau)/100;
      Obj.Cpte := T.GetString(Gen);
      {11/01/06 : FQ 10323 / 10324 : Mémorisation du code et du Régime de TVA}
      Obj.Code := CTv;
      Obj.Regm := Reg;

      lCorresp.AddObject(Flx + ';' + CBq, Obj);
    end;
  end;

  {On vide les listes pour libérer de la mémoire}
  tCompte  .ClearDetail;
  tTVA     .ClearDetail;
  tFluxTres.ClearDetail;
  tGeneraux.ClearDetail;
end;

{Renvoie True si le compte général passé en paramètre est de nature 'CHA'
{---------------------------------------------------------------------------------------}
function TObjTVA.IsAchat(Cpte : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := False;
  T := tGeneraux.FindFirst(['G_GENERAL'], [Cpte], True);
  if Assigned(T) then Result := T.GetString('G_NATUREGENE') = 'CHA';
end;

{Renvoie le taux et le compte de TVA pour un flux et un compte bancaire donnés
{---------------------------------------------------------------------------------------}
procedure TObjTVA.GetCorrespondance(var Obj : TObjDetailTVA; Cpte, Flx : string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  n := lCorresp.IndexOf(Flx + ';' + Cpte);
  if n > -1 then Obj := TObjDetailTVA(lCorresp.Objects[n])
            else Obj := nil;
end;

          {***************************************************************}
          {*********************       TObjTaux     **********************}
          {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TObjTaux.Create(DateDeb, DateFin : TDateTime; MaxOk : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  FDateDeb := DateDeb;
  FDateFin := DateFin;
  lTaux    := TStringList.Create;
  lTaux.Duplicates := dupIgnore;
  lTaux.Sorted := True;
  ChargerListeTaux(MaxOk);
end;

{---------------------------------------------------------------------------------------}
destructor TObjTaux.Destroy;
{---------------------------------------------------------------------------------------}
begin
  LibereListe(lTaux , True);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjTaux.ChangerDates(DateDeb, DateFin : TDateTime; MaxOk : Boolean = False; ClearOk : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  if (FDateDeb = DateDeb) and (FDateFin = DateFin) then Exit;
  FDateDeb := DateDeb;
  FDateFin := DateFin;
  if ClearOk then LibereListe(lTaux, False);
  ChargerListeTaux(MaxOk);
end;

{---------------------------------------------------------------------------------------}
function TObjTaux.GetTaux(Taux, EnDate : string) : Double;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := 1;{ou 0, à voir !!!}
  n := lTaux.IndexOf(Taux + '|' + EnDate);
  if n > -1 then Result := TObjNombre(lTaux.Objects[n]).Nombre;
end;

{---------------------------------------------------------------------------------------}
procedure TObjTaux.ChargerListeTaux(MaxOk : Boolean);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  S : string;
  O : TObjNombre;
  n : Integer;
begin
  {03/05/04 : FQ 10056 : si l'on veut récupérer un taux même s'il n'y en a pas dans la période de référence,
              on prend le dernier taux trouvé}
  if MaxOk then
    S := 'SELECT TTA_CODE, TTA_COTATION FROM COTATIONTAUX C1 WHERE TTA_DATE IN (SELECT MAX(TTA_DATE) ' +
         'FROM COTATIONTAUX C2 WHERE C2.TTA_CODE = C1.TTA_CODE AND C2.TTA_DATE <="' + USDATETIME(FDateFin) + '")'
  else
    S := 'SELECT TTA_CODE, TTA_DATE, TTA_COTATION FROM COTATIONTAUX WHERE TTA_DATE >= "' + USDATETIME(FDateDeb) +
         '" AND TTA_DATE <= "' + USDATETIME(FDateFin) + '" ORDER BY TTA_CODE, TTA_DATE DESC';

  Q := OpenSQL(S, True);
  try
    S := DateToStr(FDateFin);
    while not Q.EOF do begin
      if not MaxOk then S := Q.FindField('TTA_DATE').AsString;
      n := lTaux.IndexOf(Q.FindField('TTA_CODE').AsString + '|' + S);
      if n = -1 then begin
        O := TObjNombre.Create;
        O.Nombre := Q.FindField('TTA_COTATION').AsFloat;
        lTaux.AddObject(Q.FindField('TTA_CODE').AsString + '|' + S, O);
      end;
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;
end;


          {***************************************************************}
          {********************       TObjDevise     *********************}
          {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TObjDevise.Create(DtValeur : TDateTime; AvecComptes : Boolean = True);
{---------------------------------------------------------------------------------------}
begin
  FDateValeur := DtValeur;
  tBanque     := TOB.Create('£££', nil, -1);
  lCorresp    := TStringList.Create;
  lListeDev   := TStringList.Create;

  {Si l'on utilise en liaison avec les Comptes bancaires}
  if AvecComptes then begin
    {Chargement des comptes et de leur devises}
    GetTobBanque;
    {Constitution de la liste de correspondance}
    SetListeCorrespondance;
  end;
end;

{---------------------------------------------------------------------------------------}
destructor TObjDevise.Destroy;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(tBanque);
  LibereListe(lCorresp , True);
  LibereListe(lListeDev, True);
  inherited Destroy;
end;

{Chargement des comptes et de leur devise
{---------------------------------------------------------------------------------------}
procedure TObjDevise.GetTobBanque;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT BQ_DEVISE, BQ_CODE FROM BANQUECP', True);
  tBanque.LoadDetailDB('', '', '', Q, False);
  Ferme(Q);
end;

{Pour chaque compte, on récupère la cotation de la devise
{---------------------------------------------------------------------------------------}
procedure TObjDevise.SetListeCorrespondance;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  Obj : TObjDetailDevise;
begin
  {On vide la liste}
  LibereListe(lCorresp, False);

  for n := 0 to tBanque.Detail.Count - 1 do begin
    Obj := TObjDetailDevise.Create;
    Obj.Devise    := tBanque.Detail[n].GetString('BQ_DEVISE');
    Obj.DtCotat   := FDateValeur;
    {23/05/05 : Gestion des décimales, par défaut celle de la devise pivot}
    Obj.NbDecimal := V_PGI.OkDecV;
    {$IFNDEF EAGLSERVER}
    TheData := Obj;
    {$ENDIF EAGLSERVER}
    Obj.Cotation  := RetPariteEuro(Obj.Devise, FDateValeur);
    lCorresp.AddObject(tBanque.Detail[n].GetString('BQ_CODE'), Obj);
  end;
  {$IFNDEF EAGLSERVER}
  TheData := nil;
  {$ENDIF EAGLSERVER}
end;

{06/12/04 : Renvoie le montant dans la devise du compte à partir d'un montant en devise pivot
{---------------------------------------------------------------------------------------}
function TObjDevise.GetMntDevFromEur(MntE : Double; Cpte : string; Dt : TDateTime = 1) : Double;
{---------------------------------------------------------------------------------------}
var
  Taux : Double;
  n    : Integer;
begin
  Result := MntE;
  Taux := 1;
  n    := lCorresp.IndexOf(Cpte);
  if n = -1 then Exit;

  {Si on ne s'occupe pas de la date de conversion}
  if (Dt = 1) and (TObjDetailDevise(lCorresp.Objects[n]).Cotation <> 0) then
    Result := MntE / TObjDetailDevise(lCorresp.Objects[n]).Cotation

  else
  {On veut un taux à une date précise}
  if Dt <> 1 then begin
    {On récupère la cotation}
    GetTauxDevise(Taux, TObjDetailDevise(lCorresp.Objects[n]).Devise, Dt);
    if Taux <> 0 then
      Result := MntE / Taux;
  end;
  {24/05/05 : Gestion du format du résultat}
  Result := Arrondi(Result, GetNbDecimalesFromCpt(Cpte));
end;

{Convertit un montant d'une devise donnée en euro et dans la devise du compte donné
{---------------------------------------------------------------------------------------}
procedure TObjDevise.ConvertitMnt(var MntE, MntD : Double; DevOrigine, Compte : string; Dt : TDateTime = 1);
{---------------------------------------------------------------------------------------}
var
  Taux : Double;
  n    : Integer;
begin
  {Remarque : Dans la trésorerie, on part du principe MontantEur * Cotation = MontantDev.
              Cependant, historiquement, RetPariteEuro renvoie le taux et non la cotation :
              en fait de cotation (1€ = x.xx Dev), ici il faut entendre Taux (1 Dev = x.xx €)}
  Taux := 1;
  MntD := MntE;
  n    := lCorresp.IndexOf(Compte);

  if n = -1 then Exit;

  {Le montant est en Euro}
  if DevOrigine = V_PGI.DevisePivot then begin
    {06/12/04 : Appel de la fonction GetPariteFromEur}
    MntD := GetMntDevFromEur(MntE, Compte, Dt);
  end

  {Le montant est dans la devise du compte}
  else if DevOrigine = TObjDetailDevise(lCorresp.Objects[n]).Devise then begin
    MntD := MntE;
    {Si on ne se préoccupe pas de la date}
    if Dt = 1 then
      MntE := MntE * TObjDetailDevise(lCorresp.Objects[n]).Cotation
    else begin
      {On récupère la cotation}
      GetTauxDevise(Taux, DevOrigine, Dt);
      MntE := MntE * Taux;
    end
  end

  {Le montant est dans une autre devise}
  else begin
    {On récupère le taux de la devise source par rapport à l'euro}
    GetTauxDevise(Taux, DevOrigine, Dt);
    {On convertit en Euro}
    MntE := MntE * Taux;
    {On récupère le taux de la devise du compte par rapport à l'euro}
    GetTauxDevise(Taux, TObjDetailDevise(lCorresp.Objects[n]).Devise, Dt);
    {Puis dans la devise du compte}
    MntD := MntE / Taux;
  end;
  {20/05/05 : Format des montants renvoyés}
  MntE := Arrondi(MntE, V_PGI.OkDecV);
  MntD := Arrondi(MntD, GetNbDecimalesFromCpt(Compte));
end;

{---------------------------------------------------------------------------------------}
procedure TObjDevise.GetTauxDevise(var Taux : Double; Dev : string; Dt : TDateTime = 1);
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  DOk : Boolean;
  Obj : TObjDetailDevise;
begin
  {Remarque : Dans la trésorerie, on part du principe MontantEur * Cotation = MontantDev.
              Cependant, historiquement, RetPariteEuro renvoie le taux et non la cotation :
              en fait de cotation (1Dev = x.xx €), ici il faut entendre Taux (1 € = x.xx Dev)}
  Taux := 1;
  {On regarde si la devise appartient à la liste des devises des comptes bancaires}
  for n := 0 to lCorresp.Count - 1 do begin
    {La date de change est ok si on gère la dernière trouvé (1) ou si les dates correspondent}
    DOk := (Dt = 1) or (TObjDetailDevise(lCorresp.Objects[n]).DtCotat = Dt);
    if Dok and(TObjDetailDevise(lCorresp.Objects[n]).Devise = Dev) then begin
      Taux := TObjDetailDevise(lCorresp.Objects[n]).Cotation;
      Exit;
    end;
  end;

  {Sinon, on regarde si on a déjà traité la devise ...}
  n := lListeDev.IndexOf(Dev + ';' + DateToStr(Dt));
  if n > -1 then
    {JP 02/03/04 : Cela ressemble à un gros Bug, mais bizzarement,
                  l'erreur (Index Hors limite) ne s'était jamais produite :
                  Taux := TObjDetailDevise(lCorresp.Objects[n]).Cotation}
    Taux := TObjDetailDevise(lListeDev.Objects[n]).Cotation

  {... sinon on récupère le taux et on la stocke}
  else begin
    {Si on ne se préoccupe pas de la date}
    if Dt = 1 then Taux := RetPariteEuro(Dev, FDateValeur)
              {... sinon on récupère la parité la plus proche}
              else Taux := RetPariteEuro(Dev, Dt);

    if Taux = 0 then Taux := 1;
    Obj := TObjDetailDevise.Create;
    Obj.Cotation := Taux;
    if Dt = 1 then Obj.DtCotat := FDateValeur
              else Obj.DtCotat := Dt;
    lListeDev.AddObject(Dev + ';' + DateToStr(Dt), Obj);
  end;
end;

{Renvoie la devise du compte
{---------------------------------------------------------------------------------------}
function TObjDevise.GetDeviseCpt(Compte : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  n := lCorresp.IndexOf(Compte);
  if n = -1 then Result := V_PGI.DevisePivot
            else Result := TObjDetailDevise(lCorresp.Objects[n]).Devise;
end;

{23/05/05 : Retourne le nombre de décimales en fonction de la devise
{---------------------------------------------------------------------------------------}
function TObjDevise.GetNbDecimalesFromCpt(Cpt : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {Par défaut, on récupère le nombre de décimales de la devise pivot}
  Result := V_PGI.OkDecV;
  n := lCorresp.IndexOf(Cpt);
  if n > -1 then Result := TObjDetailDevise(lCorresp.Objects[n]).NbDecimal;
end;

{23/05/05 : Retourne le nombre de décimales en fonction de la devise
{---------------------------------------------------------------------------------------}
function TObjDevise.GetNbDecimalesFromDev(Dev : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  {Par défaut, on récupère le nombre de décimales de la devise pivot}
  Result := V_PGI.OkDecV;
  for n := 0 to lCorresp.Count - 1 do
    if TObjDetailDevise(lCorresp.Objects[n]).Devise = Dev then
      Result := TObjDetailDevise(lCorresp.Objects[n]).NbDecimal;
end;

{Renvoie la parité entre DevOrig et DevDest
{---------------------------------------------------------------------------------------}
function TObjDevise.GetParite(DevOrig, DevDest : string) : Double;
{---------------------------------------------------------------------------------------}
var
  Tx1, Tx2 : Double;
begin
  {Remarque : Dans la trésorerie, on part du principe MontantEur * Cotation = MontantDev.
              Cependant, historiquement, RetPariteEuro renvoie le taux et non la cotation :
              en fait de cotation (1€ = x.xx Dev), ici il faut entendre Taux (1 Dev = x.xx €)}

  {On récupère le taux de la devise source par rapport à l'euro}
  GetTauxDevise(Tx1, DevOrig);
  {On récupère le taux de la devise destination par rapport à l'euro}
  GetTauxDevise(Tx2, DevDest);
  if Tx2 = 0 then Tx2 := 1;
  Result := Arrondi(Tx1 / Tx2, NBDECIMALTAUX);
end;

{06/12/04 : Renvoie le montant à mettre dans les Champs _COTATION}
{---------------------------------------------------------------------------------------}
function TObjDevise.Get_Cotation(Cpte : string; Dt : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
begin
  Result := 1;
  GetTauxDevise(Result, GetDeviseCpt(Cpte), Dt);
  if Result = 0 then Result := 1
                else Result := Arrondi(1 / Result, NBDECIMALTAUX);
end;

          {***************************************************************}
          {********************   TObjBanquePrevi    *********************}
          {***************************************************************}

{---------------------------------------------------------------------------------------}
constructor TObjBanquePrevi.Create;
{---------------------------------------------------------------------------------------}
begin
  LBanqPrevi := TStringList.Create;
  LBnqFiltre := TStringList.Create;
  Recharger(False);
end;

{---------------------------------------------------------------------------------------}
destructor TObjBanquePrevi.Destroy;
{---------------------------------------------------------------------------------------}
begin
  LibereListe(LBanqPrevi, True);
  if Assigned(LBnqFiltre) then FreeAndNil(LBnqFiltre);
  inherited Destroy;
end;

{26/07/05 : FQ 10280 : Ajout de FromDisk pour Relancer la requête sur les PramSoc
{---------------------------------------------------------------------------------------}
procedure TObjBanquePrevi.Recharger(FromDisk : Boolean);
{---------------------------------------------------------------------------------------}
begin
  ChargerParam(FromDisk);
  MajTableaux;
end;

{26/07/05 : FQ 10280 : Ajout de FromDisk pour Relancer la requête sur les PramSoc
{---------------------------------------------------------------------------------------}
procedure TObjBanquePrevi.ChargerParam(FromDisk : Boolean);
{---------------------------------------------------------------------------------------}
var
  n : Byte;
begin
  {On récupère les comptes banquaires}
  ParamCptes[nc_Client] := GetParamSocSecur('SO_COLCLIENT'     , '', FromDisk);
  ParamCptes[nc_Fourni] := GetParamSocSecur('SO_COLFOURNISSEUR', '', FromDisk);
  ParamCptes[nc_Divers] := GetParamSocSecur('SO_COLDIVERS'     , '', FromDisk);
  ParamCptes[nc_Salari] := GetParamSocSecur('SO_COLSALARIE'    , '', FromDisk);
  ParamCptes[nc_Debite] := GetParamSocSecur('SO_TIERSDEB'      , '', FromDisk);
  ParamCptes[nc_Credit] := GetParamSocSecur('SO_TIERSCRED'     , '', FromDisk);

  {On regarde ceux qui ont été paramétrés.
   On renseigne la variable précisant que le paramétrage est absent}
  ParamAbst := True;
  for n := 0 to max_Nature do begin
    ParamValid[n] := ParamCptes[n] <> '';
    if not ParamValid[n] then ParamLibel[n] := TraduireMemoire('Veuillez définir la banque prévis.');
    ParamAbst := ParamAbst and not ParamValid[n];
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjBanquePrevi.MajTableaux;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  D : TOB;
  w : string;
  n : Integer;
  O : TObjInfosBque;
begin
  {Si aucun paramétrage n'est en place}
  if ParamAbst then Exit;
  {On nettoie la liste contenant les informations banquaires}
  LibereListe(LBanqPrevi, False);

  {Constitution de la clause where}
  for n := 0 to max_Nature do
    if ParamValid[n] then w := w + '"' + ParamCptes[n] + '",';
  System.Delete(w, Length(w), 1);

  T := TOB.Create('****', nil, -1);
  try
    {JP 19/07/05 : on ne récupérait pas BQ_BANQUE}
    T.LoadDetailFromSQL('SELECT BQ_BANQUE, BQ_CODE, BQ_LIBELLE, BQ_ETABBQ, PQ_ETABBQ, BQ_GENERAL, BQ_NODOSSIER, ' +
                        'BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN, G_LIBELLE FROM BANQUECP ' +
                       // 'LEFT JOIN BANQUES ON PQ_BANQUE  = BQ_BANQUE WHERE BQ_CODE IN (' + w + ')');
                        'LEFT JOIN GENERAUX ON G_GENERAL = BQ_GENERAL ' +
                        'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE WHERE BQ_GENERAL IN (' + w + ')');
    {On mémorise toute les infos dans d'une StringList}
    for n := 0 to T.Detail.Count - 1 do begin
      D := T.Detail[n];
      O := TObjInfosBque.Create;
      {JP 19/07/05 : on ne récupérait pas BQ_BANQUE, ce qui posait quelque problème dans AccImporttreso.TravaillerDonnees}
      O.Banque  := D.GetString('BQ_BANQUE');
      O.General := D.GetString('BQ_GENERAL');
      O.Code    := D.GetString('BQ_CODE');
      O.Dossier := D.GetString('BQ_NODOSSIER');
      O.CodeBq  := D.GetString('BQ_ETABBQ');
      O.CodeGch := D.GetString('BQ_GUICHET');
      O.CodeCpt := D.GetString('BQ_NUMEROCOMPTE');
      O.Cle     := D.GetString('BQ_CLERIB');
      O.Iban    := D.GetString('BQ_CODEIBAN');
      if O.CodeBq = '' then O.CodeBq := D.GetString('PQ_ETABBQ');
      LBanqPrevi.AddObject(O.Code, O);
    end;

    {Récupération des libellés des comptes bancaires}
    for n := 0 to max_Nature do begin
      D := T.FindFirst(['BQ_GENERAL'], [ParamCptes[n]], True);
      if Assigned(D) then begin
        ParamLibel[n] := D.GetString('G_LIBELLE');
      end
      {S'il y a un paramétrage de défini, mais que le compte est introuvable}
      else if ParamValid[n] then begin
        ParamValid[n] := False;
        ParamLibel[n] := TraduireMemoire('Le paramétrage est incorrect');
      end;
    end;

    {On s'assure qu'il existe au moins un paramétrage de correct}
    ParamAbst := True;
    for n := 0 to max_Nature do
      ParamAbst := ParamAbst and not ParamValid[n];
  finally
    if Assigned(T) then FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.IndToNature(Ind : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  case Ind of
    nc_Client : Result := 'COC';
    nc_Fourni : Result := 'COF';
    nc_Divers : Result := 'COD';
    nc_Salari : Result := 'COS';
    nc_Debite : Result := 'TID';
    nc_Credit : Result := 'TIC';
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.NatToIndice(Nat : string) : Byte;
{---------------------------------------------------------------------------------------}
begin
  Result := nc_Erreur;
  if (Nat = 'BQE') or
     (Nat = 'CAI')    then Result := nc_BqCais
  else if Nat = 'DIV' then Result := nc_CouTit
  else if Nat = 'COC' then Result := nc_Client
  else if Nat = 'COF' then Result := nc_Fourni
  else if Nat = 'COD' then Result := nc_Divers
  else if Nat = 'COS' then Result := nc_Salari
  else if Nat = 'TID' then Result := nc_Debite
  else if Nat = 'TIC' then Result := nc_Credit;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetLibelle(Ind : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := ParamLibel[Ind];
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetCompte(Ind : Byte) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := ParamCptes[Ind];
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetEtabBq(Cpt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjInfosBque;
begin
  Result := '';
  Obj := GetInfosBq(Cpt);
  if Assigned(Obj) then
    Result := Obj.CodeBq;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetGuichet(Cpt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjInfosBque;
begin
  Result := '';
  Obj := GetInfosBq(Cpt);
  if Assigned(Obj) then
    Result := Obj.CodeGch;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetNumCpt(Cpt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjInfosBque;
begin
  Result := '';
  Obj := GetInfosBq(Cpt);
  if Assigned(Obj) then
    Result := Obj.CodeCpt;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetCleRib(Cpt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjInfosBque;
begin
  Result := '';
  Obj := GetInfosBq(Cpt);
  if Assigned(Obj) then
    Result := Obj.Cle;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetIban(Cpt : string) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjInfosBque;
begin
  Result := '';
  Obj := GetInfosBq(Cpt);
  if Assigned(Obj) then
    Result := Obj.Iban;
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.IsValide(Ind : Byte) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (Ind = nc_BqCais) or (Ind = nc_CouTit) or ParamValid[Ind];
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.IsFiltre(BqPrevi, NatGene : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  n : Byte;
begin
  Result := False;

  {Il n'y a pas de filtre}
  if LBnqFiltre.Count = 0 then Exit;

  {S'il s'agit d'une écriture de banque, les écritures sont filtrées dans la requête}
  if (NatGene = 'BQE') or (NatGene = 'CAI') then Exit;

  {Si la banque prévisionnelle est renseignée, on regarde si elle doit être traitée}
  if BqPrevi <> '' then begin
    Result := LBnqFiltre.IndexOf(NoDossier + '|' + BqPrevi) > -1;
    Exit;
  end;

  {Sinon, on regarde si la banque paramétrée est dans le filtre}
  n := NatToIndice(NatGene);
  {Il ne s'agit pas d'un Tic/Tid ou d'un collectif, on considère qu'il faut traiter l'écriture}
  if n > max_nature then Exit;
  Result := LBnqFiltre.IndexOf(NoDossier + '|' + ParamCptes[n]) > -1;
end;

{LBnqFiltre Contient la liste des comptes à ne pas traiter
{---------------------------------------------------------------------------------------}
procedure TObjBanquePrevi.MajListeFiltre(Valeur : string);
{---------------------------------------------------------------------------------------}
begin
  LBnqFiltre.Clear;
  if (Valeur = '') or (Pos('<<', Valeur) > 0) then Exit;
  while Valeur <> '' do LBnqFiltre.Add(ReadTokenSt(Valeur));
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.IsAllOk(var Msg : string) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
  n  : Byte;
begin
  Msg    := '';
  Result := not ParamAbst;
  {On ne lance pas le traitement pour l'échéancier si aucune banque prévisionnelle n'est paramétrée}
  if ParamAbst then begin
    Msg := TraduireMemoire('Les banques prévisionnelles ne sont pas paramétrées ou le sont mal.'#13 +
                           'Veuillez corriger votre paramétrage.');
    Exit;
  end;

  {On regarde si toutes les banques sont correctement paramétrées}
  Ok := True;
  for n := 0 to max_nature do
    Ok := Ok and ParamValid[n];
  if not Ok then
    Msg := TraduireMemoire('Le paramétrage de certaines banques prévisionnelles n''est pas correct.'#13 +
                           'Voulez-vous poursuivre et ignorer les écritures concernées ?');
end;

{---------------------------------------------------------------------------------------}
function TObjBanquePrevi.GetInfosBq(Cpt : string) : TObjInfosBque;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  Q : TQuery;
begin
  Result := nil;
  n := LBanqPrevi.IndexOf(Cpt);
  {La banque prévisionnelle n'a pas encore été stockée dans la liste}
  if n = -1 then begin
    Q := OpenSQL('SELECT BQ_BANQUE, PQ_ETABBQ, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, ' +
                 'BQ_GENERAL, BQ_NODOSSIER, BQ_CODEIBAN FROM BANQUECP ' +
                 'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE WHERE BQ_CODE = "' + Cpt + '"', True);
    if not Q.EOF then begin
      Result := TObjInfosBque.Create;
      Result.Code    := Cpt;
      Result.Dossier := Q.FindField('BQ_NODOSSIER').AsString;
      Result.General := Q.FindField('BQ_GENERAL').AsString;
      Result.Banque  := Q.FindField('BQ_BANQUE').AsString;
      Result.CodeBq  := Q.FindField('BQ_ETABBQ').AsString;
      Result.CodeGch := Q.FindField('BQ_GUICHET').AsString;
      Result.CodeCpt := Q.FindField('BQ_NUMEROCOMPTE').AsString;
      Result.Cle     := Q.FindField('BQ_CLERIB').AsString;
      Result.Iban    := Q.FindField('BQ_CODEIBAN').AsString;
      if Result.CodeBq = '' then Result.CodeBq := Q.FindField('PQ_ETABBQ').AsString;
    end;
    {Si le compte n'a pas été trouvé, on mémorise nil}
    LBanqPrevi.AddObject(Cpt, Result);
  end
  else begin
    Result := TObjInfosBque(LBanqPrevi.Objects[n]);
  end;
end;



          {***************************************************************}
          {********************   TObjCIBModePaie    *********************}
          {***************************************************************}


{---------------------------------------------------------------------------------------}
constructor TObjCIBModePaie.Create;
{---------------------------------------------------------------------------------------}
begin
  tBqe := TOB.Create('£££', nil, -1);
  tCib := TOB.Create('$$$', nil, -1);
  ChargerTOB;
end;

{---------------------------------------------------------------------------------------}
destructor TObjCIBModePaie.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(tCib) then FreeAndNil(tCib);
  if Assigned(tBqe) then FreeAndNil(tBqe);
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjCIBModePaie.ChargerTOB;
{---------------------------------------------------------------------------------------}
var
  Wh : string;
begin
  {$IFNDEF TRESO}
  Wh := ' WHERE BQ_NODOSSIER = "' + V_PGI.NoDossier + '"';
  {$ENDIF TRESO}
  tBqe.LoadDetailFromSQL('SELECT BQ_GENERAL, BQ_CODE, BQ_BANQUE FROM BANQUECP' + Wh);
  tCib.LoadDetailFromSQL('SELECT TCI_MODEPAIE, TCI_SENS, TCI_BANQUE, TCI_CODECIB FROM CIB');
end;

{---------------------------------------------------------------------------------------}
function TObjCIBModePaie.GetBanque(General : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  {$IFDEF TRESO}
  T := tBqe.FindFirst(['BQ_CODE'], [General], False);
  {$ELSE}
  T := tBqe.FindFirst(['BQ_GENERAL'], [General], False);
  {$ENDIF TRESO}
  if Assigned(T) then Result := T.GetString('BQ_BANQUE');
end;

{Permet de retourner le CIB pour une écriture de trésorerie à partir
 de son mode de paiement , se son sens et de son compte général
{---------------------------------------------------------------------------------------}
function TObjCIBModePaie.GetCIB(General, ModePaie, Sens : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_MODEPAIE', 'TCI_SENS'], [GetBanque(General), ModePaie, Sens], False);
  if Assigned(T) then
    Result := T.GetString('TCI_CODECIB')
  else begin
    {Si on n'a pas trouver de cib pour le sens indiqué, on regarde s'il en existe un qui soit mixte}
    T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_MODEPAIE', 'TCI_SENS'], [GetBanque(General), ModePaie, 'MIX'], False);
    if Assigned(T) then
      Result := T.GetString('TCI_CODECIB')
    else begin
      {Si on n'a rien trouvé, on cherche au niveau de la référence}
      T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_MODEPAIE', 'TCI_SENS'], [CODECIBREF, ModePaie, Sens], False);
      if Assigned(T) then
        Result := T.GetString('TCI_CODECIB')
      else begin
        T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_MODEPAIE', 'TCI_SENS'], [CODECIBREF, ModePaie, 'MIX'], False);
        if Assigned(T) then
          Result := T.GetString('TCI_CODECIB');
      end;
    end;
  end;
end;

{Permet de retrouver le mode de paiement d'une écriture de trésorerie à partir
 de son CIB, se son sens et de son compte général
{---------------------------------------------------------------------------------------}
function TObjCIBModePaie.GetModePaie(General, CodeCib , Sens : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_CODECIB', 'TCI_SENS'], [GetBanque(General), CodeCib, Sens], False);
  if Assigned(T) then
    Result := T.GetString('TCI_MODEPAIE')
  else begin
    {Si on n'a pas trouver de mode de paiement pour le sens indiqué, on regarde s'il en existe un qui soit mixte}
    T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_CODECIB', 'TCI_SENS'], [GetBanque(General), CodeCib, 'MIX'], False);
    if Assigned(T) then
      Result := T.GetString('TCI_MODEPAIE')
    else begin
      {Si on n'a rien trouvé, on cherche au niveau de la référence}
      T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_CODECIB', 'TCI_SENS'], [CODECIBREF, CodeCib, Sens], False);
      if Assigned(T) then
        Result := T.GetString('TCI_MODEPAIE')
      else begin
        T := tCIB.FindFirst(['TCI_BANQUE', 'TCI_CODECIB', 'TCI_SENS'], [CODECIBREF, CodeCib, 'MIX'], False);
        if Assigned(T) then
          Result := T.GetString('TCI_MODEPAIE');
      end;
    end;
  end;
end;


          {***************************************************************}
          {********************   TobjComptesTreso   *********************}
          {***************************************************************}


{---------------------------------------------------------------------------------------}
constructor TobjComptesTreso.Create(DosSource, DosDest: string);
{---------------------------------------------------------------------------------------}
begin
  TobBqe := TOB.Create('****', nil, -1);
  FDosSource := DosSource;
  FDosDest   := DosDest;
  FEstMulti  := IsTresoMultiSoc;
  //(FDosSource <> '') and (FDosDest <> '') and (FDosSource <> CODEDOSSIERDEF) and (FDosDest <> CODEDOSSIERDEF);
  ChargeLesComptes;
end;

{---------------------------------------------------------------------------------------}
destructor TobjComptesTreso.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobBqe) then FreeAndNil(TobBqe);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TobjComptesTreso.ChargeLesComptes;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  B : string;
  S : string;
  D : string;
  T : TOB;
begin
  TobBqe.ClearDetail;

  if FEstMulti then begin
    {Si les deux dossier sont identiques, on récupère le "compte poubelle" des virements de compte à compte}
    if FDosSource = FDosDest then begin
      Q := OpenSQL('SELECT TFT_GENERAL FROM FLUXTRESO WHERE TFT_FLUX = "' + CODETRANSACDEP + '"', True);
      if not Q.EOF then CptPoubelle := Q.FindField('TFT_GENERAL').AsString;
      Ferme(Q);
      {Chargement des comptes des dossiers source et destination}
      TobBqe.LoadDetailFromSQL('SELECT BQ_CODE, BQ_GENERAL, BQ_NODOSSIER DOSSOURCE, "" DOSDEST, BQ_ETABBQ, ' +
                               'BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN FROM BANQUECP ' +
                               'WHERE BQ_NODOSSIER = "' + FDosDest + '"');
    end
    else begin
      {Chargement des comptes des dossiers source et destination}
      TobBqe.LoadDetailFromSQL('SELECT BQ_CODE, BQ_GENERAL, BQ_NODOSSIER DOSSOURCE, "" DOSDEST, BQ_ETABBQ, ' +
                               'BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN FROM BANQUECP ' +
                               'WHERE BQ_NODOSSIER IN ("' + FDosSource + '", "' + FDosDest + '")');

      {Récupération des noms de base source et destination}
      S := GetInfosFromDossier('DOS_NODOSSIER', FDosSource, 'DOS_NOMBASE');
      D := GetInfosFromDossier('DOS_NODOSSIER', FDosDest  , 'DOS_NOMBASE');
      B := S + ';' + D + ';';
      {25/06/07 : FQ 10493 : préfixage de * pour Oracle}
      Q := OpenSQL('SELECT SYSDOSSIER, CLIENSSOC.* FROM CLIENSSOC WHERE CLS_DOSSIER IN ("' + S + '", "' + D + '")', True, -1, '', False, B);
      try
        while not Q.EOF do begin
          {Recherche du compte de liaison entre la base source et la base destination }
          if (Q.FindField('CLS_DOSSIER').AsString = D) and
             (Q.FindField('SYSDOSSIER' ).AsString = S) then begin
            T := TobBqe.FindFirst(['BQ_GENERAL', 'DOSSOURCE'], [Q.FindField('CLS_GENERAL').AsString, FDosSource], True);
            if Assigned(T) then T.SetString('DOSDEST', FDosDest);
          end
          {Recherche du compte de liaison entre la base destination et la base source }
          else if (Q.FindField('CLS_DOSSIER').AsString = S) or
                  (Q.FindField('SYSDOSSIER' ).AsString = D) then begin
            T := TobBqe.FindFirst(['BQ_GENERAL', 'DOSSOURCE'], [Q.FindField('CLS_GENERAL').AsString, FDosDest], True);
            if Assigned(T) then T.SetString('DOSDEST', FDosSource);
          end;
          Q.Next;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end
  else begin
    {On récupère le "compte poubelle" des virements de compte à compte}
    Q := OpenSQL('SELECT TFT_GENERAL FROM FLUXTRESO WHERE TFT_FLUX = "' + CODETRANSACDEP + '"', True);
    if not Q.EOF then CptPoubelle := Q.FindField('TFT_GENERAL').AsString;
    Ferme(Q);

    {Chargement de la Tob Pour la converion de BQ_CODE <-> BQ_GENERAL}
    TobBqe.LoadDetailFromSQL('SELECT BQ_CODE, BQ_GENERAL, BQ_NODOSSIER DOSSOURCE, "" DOSDEST, BQ_ETABBQ, ' +
                             'BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN  FROM BANQUECP ' +
                             'WHERE ' + BQCLAUSEWHERE);
  end;
end;

{Récupère le compte courant de liaison entre deux dossiers
{---------------------------------------------------------------------------------------}
function TobjComptesTreso.GetCouFromDoss(DosSource, DosDest : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  {Si on n'est pas en Multi dossiers ou que l'on fait un virement interne à un dossier ...}
  if DosSource = DosDest then
    {... on retourne le compte de virement de compte à compte}
    Result := CptPoubelle
  else begin
    {... sinon, on récupère le compte de liaison entre les deux dossiers}
    T := TobBqe.FindFirst(['DOSSOURCE', 'DOSDEST'], [DosSource, DosDest], True);
    if Assigned(T) then Result := T.GetString('BQ_CODE');
  end;
end;

{---------------------------------------------------------------------------------------}
function TobjComptesTreso.GetCouFromGene(General, NoDossier : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  T := TobBqe.FindFirst(['DOSSOURCE', 'BQ_GENERAL'], [NoDossier, General], True);
  if Assigned(T) then Result := T.GetString('BQ_CODE');
end;

{Retourne le compte général à partir de BQ_CODE
{---------------------------------------------------------------------------------------}
function TobjComptesTreso.GetGeneFromCode(BqCode : string) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := '';
  if BqCode = CptPoubelle then
    Result := BqCode
  else begin
    T := TobBqe.FindFirst(['BQ_CODE'], [BqCode], True);
    if Assigned(T) then Result := T.GetString('BQ_GENERAL');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TobjComptesTreso.SetLeRib(var T : TOB);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  if Assigned(T) then begin
    F := TobBqe.FindFirst(['BQ_CODE'], [T.GetString('TE_GENERAL')], True);
    if Assigned(F) then begin
      T.SetString('TE_CODEBANQUE' , F.GetString('BQ_ETABBQ'      ));
      T.SetString('TE_CODEGUICHET', F.GetString('BQ_GUICHET'     ));
      T.SetString('TE_NUMCOMPTE'  , F.GetString('BQ_NUMEROCOMPTE'));
      T.SetString('TE_CLERIB'     , F.GetString('BQ_CLERIB'      ));
      T.SetString('TE_IBAN'       , F.GetString('BQ_CODEIBAN'    ));
     end;
  end;
end;


          {***************************************************************}
          {*********************   TobjSaisieAuto   **********************}
          {***************************************************************}



{---------------------------------------------------------------------------------------}
constructor TobjSaisieAuto.Create(TobReleve : TOB; Cpte, Ref, Devise : string; DatePtg : TDateTime);
{---------------------------------------------------------------------------------------}
var
  Wh : string;
begin
  FCompte     := Cpte;
  FDevise     := Devise;
  FRefPtage   := Ref; 
  FTobReleve  := TobReleve;
  FDatePtage := DatePtg;
  FTobGuide   := TOB.Create('_TRGUIDE', nil, -1);
  TobEcriture := TOB.Create('', nil, -1);
  if Cpte = '' then Wh := 'WHERE TGU_GENERAL = "" OR TGU_GENERAL IS NULL'
               else Wh := 'WHERE TGU_GENERAL = "" OR TGU_GENERAL IS NULL OR TGU_GENERAL = "' + Cpte + '"';
  FTobGuide.LoadDetailFromSQL('SELECT * FROM TRGUIDE ' + Wh);
end;

{---------------------------------------------------------------------------------------}
destructor TobjSaisieAuto.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobGuide  ) then FreeAndNil(FTobGuide);
  if Assigned(TobEcriture) then FreeAndNil(TobEcriture);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TobjSaisieAuto.GenereEcritures;
{---------------------------------------------------------------------------------------}
{$IFDEF TRESO}
var
  n : Integer;
  F : TOB;
  T : TOB;

    {----------------------------------------------------}
    function _GetChp(LibOk : Boolean) : string;
    {----------------------------------------------------}
    begin
      if LibOk then begin
             if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_Libelle   then Result := F.GetString('CEL_LIBELLE')
        else if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_Libelle1  then Result := F.GetString('CEL_LIBELLE1')
        else if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_Libelle2  then Result := F.GetString('CEL_LIBELLE2')
        else if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_Libelle3  then Result := F.GetString('CEL_LIBELLE3')
        else if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_Reference then Result := F.GetString('CEL_REFORIGINE')
        else if FCurGuide.GetString('TGU_CHAMPLIBELLE') = tgu_NumeroPie then Result := F.GetString('CEL_REFPIECE')
                                                                        else Result := F.GetString('CEL_LIBELLE');
      end
      else begin
             if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_Libelle   then Result := F.GetString('CEL_LIBELLE')
        else if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_Libelle1  then Result := F.GetString('CEL_LIBELLE1')
        else if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_Libelle2  then Result := F.GetString('CEL_LIBELLE2')
        else if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_Libelle3  then Result := F.GetString('CEL_LIBELLE3')
        else if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_Reference then Result := F.GetString('CEL_REFORIGINE')
        else if FCurGuide.GetString('TGU_CHAMPREFERENCE') = tgu_NumeroPie then Result := F.GetString('CEL_REFPIECE')
                                                                          else Result := F.GetString('CEL_REFORIGINE');
      end;
    end;

var
  DtC : TDateTime;
  Msg : Boolean;
  Num : Integer;
{$ENDIF TRESO}
begin
{$IFDEF TRESO}
  DtC := Date;
  Msg := False;
  Num := 0;

  for n := 0 to FTobReleve.Detail.Count - 1 do begin
    F := FTobReleve.Detail[n];
    if RechercheGuide(F) then begin
      {Génération de l'écriture}
      T := CreerTrEcriture(FCompte, FCurGuide.GetString('TGU_CODEFLUX'),
                           _GetChp(True), na_Realise,
                           Abs(F.GetDouble('CEL_CREDITEURO') + F.GetDouble('CEL_CREDITDEV') - F.GetDouble('CEL_DEBITEURO') - F.GetDouble('CEL_DEBITDEV')),
                           F.GetDateTime('CEL_DATEOPERATION'), F.GetDateTime('CEL_DATEVALEUR'), F.GetString('CEL_CODEAFB'),_GetChp(False), FDevise);
      {Si l'écriture a été générée ...}
      if Assigned(T) then begin
        Inc(Num);
        {Pointage de l'écriture et du mouvement bancaire}
        T.SetDateTime('TE_DATERAPPRO', FDatePtage);
        T.SetString('TE_CODERAPPRO', IntToStr(Num));
        T.SetString('TE_REFPOINTAGE', FRefPtage);
        F.SetDateTime('CEL_DATEPOINTAGE', FDatePtage);
        F.SetString('CEL_CODEPOINTAGE', IntToStr(Num));
        F.SetString('CEL_REFPOINTAGE', FRefPtage);

        {... Changement de parent}
        T.ChangeParent(TobEcriture, -1, True);

        {... On mémorise la plus ancienne date comptable pour le recalcul des soldes}
        if DtC > F.GetDateTime('CEL_DATEOPERATION') then DtC := F.GetDateTime('CEL_DATEOPERATION');

        if not Msg then begin
          {...Avertissement sur les dates de réinitialisation}
          if TestDateEtMillesime(F.GetDateTime('CEL_DATEVALEUR'), Date) or
             TestDateEtMillesime(F.GetDateTime('CEL_DATEOPERATION'), Date) then begin
             PgiBox(TraduireMemoire('Certaines écritures ont des dates antérieures au début d''exercice.') + #13 +
                    TraduireMemoire('Cela va avoir un impact sur les soldes d''initialisation.'), TraduireMemoire('Création d''écritures'));
             Msg := True;
          end;
        end;
      end;
    end;
  end;

  if TobEcriture.Detail.Count > 0 then begin
    {Insertion de la TOB ...}
    TobEcriture.InsertDb(nil);
    {... recalcul des soldes}
    RecalculSolde(F.GetString('CEL_GENERAL'), DateToStr(DtC), 0, True);
  end;
{$ENDIF TRESO}
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
function TobjSaisieAuto.RechercheGuide(LigReleve : TOB): Boolean;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  F   : TOB;
  Chp : string;
  Val : string;
begin
  Result := False;
  FCurGuide := nil;
  for n := 0 to FTobGuide.Detail.Count - 1 do begin
    F := FTobGuide.Detail[n];
    Chp := F.GetString('TGU_CHAMPGUIDE');
    Val := F.GetString('TGU_VALEURCHAMP');
    Result := ((Chp = tgu_CIB     ) and (Val = LigReleve.GetString('CEL_CODEAFB')) or
              ((Chp = tgu_Libelle ) and CompareAvecJoker(Val, LigReleve.GetString('CEL_LIBELLE'))) or
              ((Chp = tgu_Libelle1) and CompareAvecJoker(Val, LigReleve.GetString('CEL_LIBELLE1'))) or
              ((Chp = tgu_Libelle2) and CompareAvecJoker(Val, LigReleve.GetString('CEL_LIBELLE2'))) or
              ((Chp = tgu_Libelle3) and CompareAvecJoker(Val, LigReleve.GetString('CEL_LIBELLE3'))));

    if Result then begin
      FCurGuide := F;
      Break;
    end;
  end;
end;
{$ENDIF TRESO}

{ TObjCodeCombo }

{---------------------------------------------------------------------------------------}
class procedure TObjCodeCombo.GetCodeComboUnique(var CodeCombo: string);
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------}
    function _GetNextChar(Ordre : Byte) : Char;
    {----------------------------------------------------------------}
    begin
      if Ordre < 48 {0} then
        Result := '0'
      else if Between(Ordre, 48, 56) {0..8} then {Poinctuation :, ; ...}
        Result := Chr(Ordre + 1)
      else if Ordre = 57 {9} then
        Result := Chr(65) {A}
      else if Between(Ordre, 58, 64) then {Poinctuation :, ; ...}
        Result := Chr(65) {A}
      else if Between(Ordre, 65, 89 {Y}) then {Poinctuation :, ; ...}
        Result := Chr(Ordre + 1)
      else
        Result := '$';
    end;

var
  Code : string;
  Tmp  : string;
  C1   : Char;
  Lg   : Integer;
begin
  Tmp := '';
  Lg := Length(CodeCombo);
  if Lg = 0 then
    CodeCombo := '000'
  else if Lg = 1 then
    CodeCombo := CodeCombo + '00'
  else if Lg = 2 then
    CodeCombo := CodeCombo + '0'
  else begin
    if Lg > 3 then Tmp := Copy(CodeCombo, 4, lg - 3);
    Code := Copy(CodeCombo, 1, 3);

    C1 := _GetNextChar(Ord(Code[3]));
    if C1 <> '$' then
      CodeCombo := Copy(CodeCombo, 1, 2) + C1 + Tmp
    else begin
      C1 := _GetNextChar(Ord(Code[2]));
      if C1 <> '$' then
        CodeCombo := Code[1] + C1 + Code[3] + Tmp
      else begin
        C1 := _GetNextChar(Ord(Code[1]));
        if C1 <> '$' then
          CodeCombo := C1 + Copy(CodeCombo, 2, 2) + Tmp
        else
          CodeCombo := '$$$';
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
constructor TObjCodeCombo.Create(NomTable, ChampCode : string);
{---------------------------------------------------------------------------------------}
begin
  FNomTable := NomTable;
  FNomCode  := ChampCode;
  FTobCode := TOB.Create('_CODE', nil, -1);
  FTobCode.LoadDetailFromSQL('SELECT ' + FNomCode + ' FROM ' + FNomTable)
end;

{---------------------------------------------------------------------------------------}
destructor TObjCodeCombo.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobCode) then FreeAndNil(FTobCode);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjCodeCombo.GetNewCode : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  Result := FLastCode;
  repeat
    GetCodeComboUnique(Result);
    T := FTobCode.FindFirst([FNomCode], [Result], True);
  until T = nil;

  if Result <> '$$$' then FLastCode := Result
                     else Result    := '';
end;

end.


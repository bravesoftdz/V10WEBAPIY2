{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 29/11/2002
Modifié le ... : 29/11/2002
Description .. : Fonctions pour le traitement d'affectation des commandes
Mots clefs ... : AFFCDE
*****************************************************************}
unit UtilAffCde;

interface

uses
  sysutils, Math, Windows, Stdctrls, Controls,
  {$IFNDEF EAGLCLIENT}
  DBTables,
  {$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOB, HStatus, AGLInit;

type
  TTypeEtapeAffCde = (afcSelection, afcReservation, afcAffectation, afcPreparation, afcFin);

function TraitementAffectationCde(CodeAff: string; Etape: TTypeEtapeAffCde; NbRow: integer; DataMem, ArtFerme, CliFerme: boolean; NatPiece: string): integer;
function TraitementPreparationCde(CodeAff: string; Eclate, DeGroupeRemise, AvecComment, AvecCompteRendu, ContinuOnError, MajToutStock: boolean;
  TraitFinTravaux: integer ; CompteRendu : TMemo): integer;

function InsertAFFCDELIGNE(CodeAff, Where: string): integer;
function InsertGeneriqueAFFCDELIGNE(CodeAff: string): integer;
function InsertPieceAFFCDELIGNE(CodeAff: string): integer;
function InsertAFFCDEDISPO(CodeAff: string): integer;
function UpdateAFFCDEDISPO(CodeAff: string): integer;
function DeleteAFFCDEDISPO(CodeAff: string): integer;
function DeleteAFFCDEPIECE(CodeAff: string): integer;
function DeleteAFFCDELIGNE(CodeAff, TypeDim: string): integer;

function VerifStatutAFFCDEENTETE(CodeAff: string; Etape: TTypeEtapeAffCde): boolean;
procedure MajStatutAFFCDEENTETE(CodeAff, CodeStatut: string; Termine: boolean);
function DonneCodeStatutAFFCDE(Etape: TTypeEtapeAffCde): string;
function GetQteStockAFFCDE(TOBStock: TOB; Etape: TTypeEtapeAffCde): double;
procedure MajStockAFFCDE(TOBStock: TOB; QteLivrer: double; Etape: TTypeEtapeAffCde; AncienneQte: double = 0; ForceMaj: boolean = False);

implementation

uses
  EntGC, utilPGI, FactGrp, FactComm, FactUtil, FactPiece, FactTOB, 
  UDataMemAGL, UEnreg;

type
  TTypeJointure = (joinPIECE, joinARTICLE, joinTIERS, joinTIERSCOMPL);
  TTypeJointures = set of TTypeJointure;

type
  T_AFFECTATIONCDE = class
  private
    Etape: TTypeEtapeAffCde;
    CodeAff: string;
    Depot: string;
    NaturePiece: string; // nature des pièces
    MaxNbRow: integer; // nombre de lignes à charger en TOB
    UseDataMem: boolean; // utilisation des DataMem
    VerifArtFerme: boolean; // vérification si les articles sont fermés
    VerifCliFerme: boolean; // vérification si les clients sont fermés
    AvecInitMove: boolean;
    TOBAffectation: TOB; // paramètres de l'affectation
    QQBesoin: TQuery; // liste des lignes de commandes à traiter
    IndBesoin: integer; // Indice dans l'ensemble de données du premier enregistrement chargé en TOB.
    TOBBesoin: TOB; // liste des lignes de commandes en cours de traitement
    TOBLgGen: TOB; // liste des lignes génériques
    TOBPiece: TOB; // liste des commandes
    TOBTiersPiece: TOB; // liste des exceptions par tiers et natures de pièce
    TOBDispo: TOB; // liste des stocks disponibles
    TOBGen: TOB; // dernière ligne générique traitée  !! avec les DataMem ce sont des TOBs orphelines,
    TOBCde: TOB; // dernière commande traitée         !! sinon elles pointent sur la TOB fille courante
    TOBStk: TOB; // dernier stock traité              !! de la TOB mère associée (TOBLgGen, TOBPiece ou TOBDispo)
    procedure ChargeTobAffect;
    procedure ChargeBesoin;
    function ChargeBesoinEnTob: boolean;
    procedure ChargeTobDispo;
    procedure ChargeTobPiece;
    procedure ChargeTobLgGen;
    procedure MakeOrderBy(var sOrder: string; var TypeJoin: TTypeJointures);
    function ArticleAutorise(TOBLigne: TOB; NaturePiece: string): boolean;
    function TiersAutorise(TOBLigne: TOB; NaturePiece: string): boolean;
    procedure MajLigneGenerique(TOBLigne: TOB; QteLivrer: double);
    function FindLigneGen(TOBLigne: TOB): TOB;
    function FindPiece(TOBLigne: TOB): TOB;
    procedure MajPiece(TOBLigne: TOB; OldQteLiv, OldMtLiv: double);
    function FindStock(RefArt: string; TOBLigne: TOB): TOB;
    procedure MajStock(TOBStock: TOB; QteLivrer: double);
    procedure AlloueLesTOB;
    procedure ChargeLesTOB;
    procedure MajDesTOB;
    procedure MajTOBDispo;
    procedure MajTOBPiece;
    procedure MajTOBLgGen;
    procedure VideLesTOB;
    procedure LibereLesTOB;
    function CalculQteLivrer(QteDispo, QteBesoin: double; TOBLigne: TOB): double;
    function CalculQteReservee(TOBLigne, TOBStock: TOB): boolean;
    procedure CalculMontantLigne(TOBLigne: TOB; OldQteLiv: double);
    procedure MajDataMem;
    function TraitementUnLot: integer;
  public
    constructor Create(CodeAffCde: string; CodeEtape: TTypeEtapeAffCde; NbRow: integer; DataMem, ArtFerme, CliFerme: boolean; NatPiece: string); reintroduce; overload; virtual;
    function Traitement: integer;
    destructor Destroy; override;
  end;

type
  T_RESELECTIONCDE = class
  private
    CodeAff: string;
    Depot: string;
    AvecInitMove: boolean;
    MajToutLeStock: boolean;
    TOBLesPieces: TOB; // liste des pièces à traiter
    CleDoc: R_CleDoc; // référence de la pièce en cours de traitement
    TOBPiece: TOB; // pièce et lignes en cours
    TOBAff: TOB; // pièce et lignes d'affectation de la pièce en cours
    TOBAff_O: TOB; // anciennes lignes d'affectation de la pièce
    TOBStock: TOB; // liste des stocks d'affectation
    TOBDispo: TOB; // liste des stocks disponibles
    procedure ReSelectCommande;
    procedure CreationEntete;
    procedure CreationLigne(NoLigne: integer);
    procedure ChargeTOBDispo;
    procedure DecrementStock(NoLigne: integer);
    procedure MajStock(TOBLigne: TOB; QRes, OldQRes, QAff, OldQAff: double);
    procedure AlloueLesTOB;
    procedure MajDesTOB;
    procedure LibereLesTOB;
  public
    constructor Create(CodeAffCde, CodeDepot: string; TOBPieces: TOB; MajToutStock: boolean); reintroduce; overload; virtual;
    procedure Traitement;
    destructor Destroy; override;
  end;

const
  NBPRIORAFFCDE = 3; // nombre d'ordres de priorité

  IDM_AFFCDEPIECE = 1; // DataMem des pièces
  GEP_CODEAFF = 1;
  GEP_NATUREPIECEG = 2;
  GEP_SOUCHE = 3;
  GEP_NUMERO = 4;
  GEP_INDICEG = 5;
  GEP_TOTALHT = 6;
  GEP_TOTALQTE = 7;
  GEP_TOTHTRESERVEE = 8;
  GEP_TOTALQTERES = 9;
  GEP_POUQTERES = 10;
  GEP_POUMTRES = 11;
  GEP_MODIF = 12;
  IDM_AFFCDEDISPO = 2; // DataMem des stocks
  GED_CODEAFF = 1;
  GED_DEPOT = 2;
  GED_ARTICLE = 3;
  GED_QTEPREVI = 4;
  GED_QTERESERVEE = 5;
  GED_DISPORESERVEE = 6;
  GED_MODIF = 7;
  IDM_AFFCDELIGNE = 3; // DataMem des lignes génériques
  GEL_CODEAFF = 1;
  GEL_NATUREPIECEG = 2;
  GEL_SOUCHE = 3;
  GEL_NUMERO = 4;
  GEL_INDICEG = 5;
  GEL_NUMLIGNE = 6;
  GEL_QTEALIVRER = 7;
  GEL_MONTANTHT = 8;
  GEL_QTERESERVEE = 9;
  GEL_MTHTRESERVEE = 10;
  GEL_MODIF = 11;

  RC_CODEAFF = 1; // Clefs des DataMem
  RC_NATUREPIECEG = 2;
  RC_SOUCHE = 3;
  RC_NUMERO = 4;
  RC_INDICEG = 5;
  RC_NUMLIGNE = 6;
  RC_DEPOT = 7;
  RC_ARTICLE = 8;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Calcul du montant de la ligne d'affectation
Mots clefs ... :
*****************************************************************}
procedure CalculMontantLigneAff ( TOBLigne : TOB ; CodeEtape : TTypeEtapeAffCde ) ;
var
  Qte, Montant, PxUnit: double ;
  NomChamp: string ;
begin
  if TOBLigne = nil then Exit ;
  // calcul du prix unitaire
  Qte := TOBLigne.GetValue ('GEL_QTEALIVRER') ;
  if Qte = 0 then Exit ;
  Montant := TOBLigne.GetValue ('GEL_MONTANTHT') ;
  PxUnit := Arrondi (Montant / Qte, 6) ;
  // recherche de la quantité et du champ à alimenter
  if CodeEtape = afcReservation then
  begin
    Qte := TOBLigne.GetValue ('GEL_QTERESERVEE') ;
    NomChamp := 'GEL_MTHTRESERVEE' ;
  end else
  begin
    Qte := TOBLigne.GetValue ('GEL_QTEAFFECTEE') ;
    NomChamp := 'GEL_MTHTAFFECTEE' ;
  end ;
  Montant := Arrondi (Qte * PxUnit, 6) ;
  TOBLigne.PutValue (NomChamp, Montant) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Calcul des totaux de la pièce d'affectation
Mots clefs ... : 
*****************************************************************}
procedure CalculMontantPieceAff(TOBPiece, TOBLigne: TOB; CodeEtape: TTypeEtapeAffCde; OldMtLiv, OldQteLiv: double; CumulQteLivrer: boolean);
var
  Mont, Qte, Total: double;
  NomChamp: string;
begin
  if (TOBPiece = nil) or (TOBLigne = nil) then Exit;
  // calcul du à livrer
  if CumulQteLivrer then
  begin
    Mont := TOBLigne.GetValue('GEL_MONTANTHT');
    Total := TOBPiece.GetValue('GEP_TOTALHT');
    Total := Arrondi(Total + Mont, 6);
    TOBPiece.PutValue('GEP_TOTALHT', Total);
    Qte := TOBLigne.GetValue('GEL_QTEALIVRER');
    Total := TOBPiece.GetValue('GEP_TOTALQTE');
    Total := Arrondi(Total + Qte, 6);
    TOBPiece.PutValue('GEP_TOTALQTE', Total);
  end;
  // calcul du total HT
  if CodeEtape = afcReservation then
  begin
    Mont := TOBLigne.GetValue('GEL_MTHTRESERVEE');
    NomChamp := 'GEP_TOTHTRESERVEE';
  end else
  begin
    Mont := TOBLigne.GetValue('GEL_MTHTAFFECTEE');
    NomChamp := 'GEP_TOTHTAFFECTEE';
  end;
  Total := TOBPiece.GetValue(NomChamp);
  Total := Arrondi(Total - OldMtLiv + Mont, 6);
  TOBPiece.PutValue(NomChamp, Total);
  // calcul du pourcentage en montant
  Qte := TOBPiece.GetValue('GEP_TOTALHT');
  if Qte = 0 then
    Mont := 0
  else
    Mont := Arrondi((Total * 100) / Qte, 6);
  if CodeEtape = afcReservation then
    NomChamp := 'GEP_POUMTRES'
  else
    NomChamp := 'GEP_POUMTAFF';
  TOBPiece.PutValue(NomChamp, Mont);
  // calcul de la quantité totale
  if CodeEtape = afcReservation then
  begin
    Qte := TOBLigne.GetValue('GEL_QTERESERVEE');
    NomChamp := 'GEP_TOTALQTERES';
  end else
  begin
    Qte := TOBLigne.GetValue('GEL_QTEAFFECTEE');
    NomChamp := 'GEP_TOTALQTEAFF';
  end;
  Total := TOBPiece.GetValue(NomChamp);
  Total := Arrondi(Total - OldQteLiv + Qte, 6);
  TOBPiece.PutValue(NomChamp, Total);
  // calcul du pourcentage en quantité
  Qte := TOBPiece.GetValue('GEP_TOTALQTE');
  if Qte = 0 then
    Mont := 0
  else
    Mont := Arrondi((Total * 100) / Qte, 6);
  if CodeEtape = afcReservation then
    NomChamp := 'GEP_POUQTERES'
  else
    NomChamp := 'GEP_POUQTEAFF';
  TOBPiece.PutValue(NomChamp, Mont);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Recherche de la ligne générique si le numéro de ligne
Suite ........ : générique n'est pas connu
Mots clefs ... :
*****************************************************************}

function TrouveLigneGeneriqueAff(TOBPiece, TOBLigne: TOB): TOB;
var
  Ind, NumL, NumG, NumLig: integer;
  NomChamp: string;
  LstChamp: array[1..5] of string;
  LstValeur: array[1..5] of variant;
  TOBL, TOBG: TOB;
begin
  Result := nil;
  if (TOBPiece = nil) or (TOBLigne = nil) then Exit;
  // recherche des lignes génériques de l'article
  for Ind := Low(LstChamp) to High(LstChamp) do
  begin
    case Ind of
      1: NomChamp := 'GEL_NATUREPIECEG';
      2: NomChamp := 'GEL_SOUCHE';
      3: NomChamp := 'GEL_NUMERO';
      4: NomChamp := 'GEL_INDICEG';
      5: NomChamp := 'GEL_CODESDIM';
    end;
    LstChamp[Ind] := NomChamp;
    if Ind = 5 then NomChamp := 'GEL_CODEARTICLE';
    LstValeur[Ind] := TOBLigne.GetValue(NomChamp);
  end;
  NumLig := TOBLigne.GetValue('GEL_NUMLIGNE');
  TOBG := TOBPiece.FindFirst(LstChamp, LstValeur, False);
  if TOBG <> nil then
  begin
    NumG := TOBG.GetValue('GEL_NUMLIGNE');
    TOBL := TOBG;
    // recherche de la ligne générique avec le n° de ligne inférieur le plus proche
    while TOBL <> nil do
    begin
      TOBL := TOBPiece.FindNext(LstChamp, LstValeur, False);
      if TOBL <> nil then
      begin
        NumL := TOBL.GetValue('GEL_NUMLIGNE');
        if (NumL < NumLig) and (NumL > NumG) then
        begin
          TOBG := TOBL;
          NumG := NumL;
        end;
      end;
    end;
  end;
  Result := TOBG;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Recherche de la ligne générique si le numéro de ligne
Suite ........ : générique est connu
Mots clefs ... :
*****************************************************************}

function TrouveNumLigneGeneriqueAff(TOBPiece, TOBLigne, TOBGen: TOB): TOB;
var
  Ind: integer;
  NomChamp: string;
  LstChamp: array[1..5] of string;
  LstValeur: array[1..5] of variant;
  Trouve: boolean;
begin
  Result := nil;
  if (TOBPiece = nil) or (TOBLigne = nil) then Exit;
  if TOBGen = nil then
    Trouve := False
  else
    Trouve := True;
  // critères de recherche de la ligne générique
  for Ind := Low(LstChamp) to High(LstChamp) do
  begin
    case Ind of
      1: NomChamp := 'GEL_NATUREPIECEG';
      2: NomChamp := 'GEL_SOUCHE';
      3: NomChamp := 'GEL_NUMERO';
      4: NomChamp := 'GEL_INDICEG';
      //5: NomChamp := 'GEL_NUMLIGNE';
      5: NomChamp := 'GEL_NUMORDRE';
    end;
    LstChamp[Ind] := NomChamp;
    //if Ind = 5 then NomChamp := 'GEL_NUMLIGNEGEN';
    if Ind = 5 then NomChamp := 'GEL_NUMORDREGEN';
    LstValeur[Ind] := TOBLigne.GetValue(NomChamp);
    // vérifie si la ligne générique précédente est celle recherchée
    if (Trouve) and (TOBGen <> nil) then
    begin
      if TOBGen.GetValue(NomChamp) <> LstValeur[Ind] then
        Trouve := False;
    end;
  end;
  // recherche de la ligne générique dans la liste des lignes
  if not Trouve then
    TOBGen := TOBPiece.FindFirst(LstChamp, LstValeur, False);
  Result := TOBGen;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Mise à jour de la ligne générique
Mots clefs ... :
*****************************************************************}

function MajLigneGeneriqueAff(TOBPiece, TOBLigne, TOBGen: TOB; CodeEtape: TTypeEtapeAffCde; QteLivrer: double; CumulQteLivrer: boolean): TOB;
var
  Qte, OldQte: double;
  NomChamp: string;
  TOBG: TOB;
begin
  Result := nil;
  if TOBLigne = nil then Exit;
  if (TOBPiece = nil) and (TOBGen = nil) then Exit;
  if (QteLivrer <= 0) and (not CumulQteLivrer) then Exit;
  if TOBLigne.GetValue('GEL_TYPEDIM') <> 'DIM' then Exit;
  // recherche des lignes génériques de l'article
  if TOBPiece = nil then
  begin
    TOBG := TOBGen;
  end else
  begin
    //if TOBLigne.GetValue('GEL_NUMLIGNEGEN') > 0 then
    if TOBLigne.GetValue('GEL_NUMORDREGEN') > 0 then
      TOBG := TrouveNumLigneGeneriqueAff(TOBPiece, TOBLigne, TOBGen)
    else
      TOBG := TrouveLigneGeneriqueAff(TOBPiece, TOBLigne);
  end;
  // incrémentation de la quantité réservée de la ligne générique
  if TOBG <> nil then
  begin
    //if TOBLigne.GetValue('GEL_NUMLIGNEGEN') = 0 then
    if TOBLigne.GetValue('GEL_NUMORDREGEN') = 0 then
    begin
      TOBLigne.PutValue('GEL_NUMLIGNEGEN', TOBG.GetValue('GEL_NUMLIGNE'));
      TOBLigne.PutValue('GEL_NUMORDREGEN', TOBG.GetValue('GEL_NUMORDRE'));
    end ;
    if CumulQteLivrer then
    begin
      OldQte := TOBG.GetValue('GEL_QTESTOCK');
      Qte := TOBLigne.GetValue('GEL_QTESTOCK');
      TOBG.PutValue('GEL_QTESTOCK', (OldQte + Qte));
      OldQte := TOBG.GetValue('GEL_QTEALIVRER');
      Qte := TOBLigne.GetValue('GEL_QTEALIVRER');
      TOBG.PutValue('GEL_QTEALIVRER', (OldQte + Qte));
      OldQte := TOBG.GetValue('GEL_MONTANTHT');
      Qte := TOBLigne.GetValue('GEL_MONTANTHT');
      TOBG.PutValue('GEL_MONTANTHT', (OldQte + Qte));
    end;
    if CodeEtape = afcReservation then
      NomChamp := 'GEL_QTERESERVEE'
    else
      NomChamp := 'GEL_QTEAFFECTEE';
    OldQte := TOBG.GetValue(NomChamp);
    Qte := OldQte + QteLivrer;
    TOBG.PutValue(NomChamp, Qte);
    // Calcul du montant réservée
    CalculMontantLigneAff(TOBG, CodeEtape);
    Result := TOBG;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 19/12/2002
Modifié le ... : 19/12/2002
Description .. : Execute un ordre SQL sur la base de données en cours en
Suite ........ : l'intégrant dans une transaction notamment pour Oracle.
Mots clefs ... :
*****************************************************************}

function ExecuteSQLAFFCDE(sSql: string): integer;
begin
  Result := 0;
  BeginTrans;
  try
    Result := ExecuteSQL(sSql);
    CommitTrans;
  except
    RollBack;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Recopie des lignes de commandes dans la table spécifique
Suite ........ : des affectations
Mots clefs ... :
*****************************************************************}

function InsertAFFCDELIGNE(CodeAff, Where: string): integer;
var
  sSql, Stg: string;
  JoinArt, JoinTiersCompl: boolean;
begin
  // Ajout du nom de la table LIGNE devant les champs GL_ de la clause WHERE
  Stg := UpperCase(Where);
  Stg := StringReplace(Stg, 'GL_', 'LIGNE.GL_', [rfReplaceAll, rfIgnoreCase]);
  JoinArt := (Pos('GA_', Stg) > 0);
  JoinTiersCompl := (Pos('YTC_', Stg) > 0);
  // Recopie des lignes de commandes
  sSql := 'INSERT INTO AFFCDELIGNE '
    + '(GEL_CODEAFF,GEL_NATUREPIECEG '
    + ' ,GEL_DATEPIECE,GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG, '
    + ' GEL_NUMLIGNE,GEL_NUMORDRE,GEL_TIERS,GEL_TIERSLIVRE, '
    + ' GEL_ARTICLE,GEL_CODEARTICLE,GEL_CODESDIM, '
    + ' GEL_TYPEDIM,GEL_NUMLIGNEGEN,GEL_NUMORDREGEN, '
    + ' GEL_TYPEARTICLE,GEL_TENUESTOCK, '
    + ' GEL_QTESTOCK,GEL_QTEALIVRER, '
    + ' GEL_QTERESERVEE,GEL_QTEAFFECTEE,GEL_MONTANTHT, '
    + ' GEL_MTHTRESERVEE, GEL_MTHTAFFECTEE, '
    + ' GEL_DATELIVRAISON,GEL_PRIORITE,GEL_MODIFIABLE,GEL_CREERPAR) '
    + 'SELECT "' + CodeAff + '" AS GEL_CODEAFF,LIGNE.GL_NATUREPIECEG, '
    + ' MAX(LIGNE.GL_DATEPIECE),LIGNE.GL_SOUCHE,LIGNE.GL_NUMERO,LIGNE.GL_INDICEG, '
    + ' LIGNE.GL_NUMLIGNE,MAX(LIGNE.GL_NUMORDRE),MAX(LIGNE.GL_TIERS),MAX(LIGNE.GL_TIERSLIVRE), '
    + ' MAX(LIGNE.GL_ARTICLE),MAX(LIGNE.GL_CODEARTICLE),MAX(LIGNE.GL_CODESDIM), '
    + ' MAX(LIGNE.GL_TYPEDIM), ISNULL(MAX(LGEN.GL_NUMLIGNE),0), ISNULL(MAX(LGEN.GL_NUMORDRE),0), '
    + ' MAX(LIGNE.GL_TYPEARTICLE),MAX(LIGNE.GL_TENUESTOCK), '
    + ' MAX(LIGNE.GL_QTESTOCK),MAX(LIGNE.GL_QTERESTE), '
    + ' 0,0,MAX((LIGNE.GL_MONTANTHT*LIGNE.GL_QTERESTE)/IIF( LIGNE.GL_QTEFACT=0, 1, LIGNE.GL_QTEFACT)), '
    + ' 0,0, '
    + ' MAX(LIGNE.GL_DATELIVRAISON),MAX(GP_PRIORITE),"X","GEN" '
    + 'FROM LIGNE '
    + 'LEFT JOIN LIGNE LGEN'
    + ' ON LGEN.GL_NATUREPIECEG=LIGNE.GL_NATUREPIECEG '
    + ' AND LGEN.GL_SOUCHE=LIGNE.GL_SOUCHE AND LGEN.GL_NUMERO=LIGNE.GL_NUMERO '
    + ' AND LGEN.GL_INDICEG=LIGNE.GL_INDICEG '
    + ' AND LGEN.GL_CODESDIM=LIGNE.GL_CODEARTICLE AND LGEN.GL_TYPELIGNE="COM" '
    + ' AND LGEN.GL_TYPEDIM="GEN" AND LGEN.GL_NUMLIGNE<LIGNE.GL_NUMLIGNE '
    + 'LEFT JOIN PIECE ON PIECE.GP_NATUREPIECEG=LIGNE.GL_NATUREPIECEG '
    + ' AND PIECE.GP_SOUCHE=LIGNE.GL_SOUCHE AND PIECE.GP_NUMERO=LIGNE.GL_NUMERO '
    + ' AND PIECE.GP_INDICEG=LIGNE.GL_INDICEG ';
  if JoinArt then
    sSql := sSql + 'LEFT JOIN ARTICLE ON GA_ARTICLE=LIGNE.GL_ARTICLE ';
  if JoinTiersCompl then
    sSql := sSql + 'LEFT JOIN TIERSCOMPL ON YTC_TIERS=LIGNE.GL_TIERS ';
  sSql := sSql + 'WHERE ' + Stg
    + 'AND NOT EXISTS (SELECT GEL_CODEAFF FROM AFFCDELIGNE '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_NATUREPIECEG=LIGNE.GL_NATUREPIECEG '
    + ' AND GEL_SOUCHE=LIGNE.GL_SOUCHE AND GEL_NUMERO=LIGNE.GL_NUMERO '
    + ' AND GEL_INDICEG=LIGNE.GL_INDICEG AND GEL_NUMLIGNE=LIGNE.GL_NUMLIGNE) '
    + 'GROUP BY LIGNE.GL_NATUREPIECEG,LIGNE.GL_SOUCHE,LIGNE.GL_NUMERO, '
    + ' LIGNE.GL_INDICEG,LIGNE.GL_NUMLIGNE';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 03/12/2002
Description .. : Création des lignes génériques pour la gestion des articles
Suite ........ : dimensionnés
Mots clefs ... :
*****************************************************************}

function InsertGeneriqueAFFCDELIGNE(CodeAff: string): integer;
var
  sSql: string;
begin
  // création de la ligne générique
  sSql := 'INSERT INTO AFFCDELIGNE '
    + '(GEL_CODEAFF,GEL_NATUREPIECEG,GEL_DATEPIECE, '
    + ' GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG,GEL_NUMLIGNE,GEL_NUMORDRE, '
    + ' GEL_TIERS,GEL_TIERSLIVRE, '
    + ' GEL_ARTICLE,GEL_CODEARTICLE, '
    + ' GEL_CODESDIM,GEL_TYPEDIM, '
    + ' GEL_NUMLIGNEGEN,GEL_NUMORDREGEN,GEL_TYPEARTICLE,GEL_TENUESTOCK, '
    + ' GEL_QTESTOCK, '
    + ' GEL_QTEALIVRER,GEL_QTERESERVEE, '
    + ' GEL_QTEAFFECTEE,GEL_MONTANTHT, '
    + ' GEL_MTHTRESERVEE,GEL_MTHTAFFECTEE, '
    + ' GEL_DATELIVRAISON,GEL_PRIORITE, '
    + ' GEL_MODIFIABLE,GEL_CREERPAR) '
    + 'SELECT GEL_CODEAFF,GEL_NATUREPIECEG,MAX(GEL_DATEPIECE) AS GEL_DATEPIECE, '
    + ' GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG,GEL_NUMLIGNEGEN AS GEL_NUMLIGNE, '
    + ' GEL_NUMORDREGEN AS GEL_NUMORDRE, '
    + ' MAX(GEL_TIERS) AS GEL_TIERS,MAX(GEL_TIERSLIVRE) AS GEL_TIERSLIVRE, '
    + ' "" AS GEL_ARTICLE,MAX(GEL_CODEARTICLE) AS GEL_CODEARTICLE, '
    + ' MAX(GEL_CODEARTICLE) AS GEL_CODESDIM,"GEN" AS GEL_TYPEDIM, '
    + ' 0 AS GEL_NUMLIGNEGEN, 0 AS GEL_NUMORDREGEN, "" AS GEL_TYPEARTICLE,"-" AS GEL_TENUESTOCK, '
    + ' SUM(GEL_QTESTOCK) AS GEL_QTESTOCK, '
    + ' SUM(GEL_QTEALIVRER) AS GEL_QTEALIVRER,SUM(GEL_QTERESERVEE) AS GEL_QTERESERVEE, '
    + ' SUM(GEL_QTEAFFECTEE) AS GEL_QTEAFFECTEE,SUM(GEL_MONTANTHT) AS GEL_MONTANTHT, '
    + ' SUM(GEL_MTHTRESERVEE) AS GEL_MTHTRESERVEE,SUM(GEL_MTHTAFFECTEE) AS GEL_MTHTAFFECTEE, '
    + ' MAX(GEL_DATELIVRAISON) AS GEL_DATELIVRAISON,MAX(GEL_PRIORITE) AS GEL_PRIORITE, '
    + ' "X" AS GEL_MODIFIABLE,"GEN" AS GEL_CREERPAR '
    + 'FROM AFFCDELIGNE '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_TYPEDIM="DIM" '
    + 'GROUP BY GEL_CODEAFF,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO, '
    + ' GEL_INDICEG,GEL_NUMLIGNEGEN,GEL_NUMORDREGEN';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 24/12/2002
Description .. : Recopie des commandes dans la table spécifique des
Suite ........ : affectations
Mots clefs ... :
*****************************************************************}

function InsertPieceAFFCDELIGNE(CodeAff: string): integer;
var
  sSql: string;
begin
  // création de la pièce
  sSql := 'INSERT INTO AFFCDEPIECE '
    + '(GEP_CODEAFF,GEP_NATUREPIECEG,GEP_DATEPIECE, '
    + ' GEP_SOUCHE,GEP_NUMERO,GEP_INDICEG, '
    + ' GEP_TOTALHT,GEP_TOTALQTE, '
    + ' GEP_TOTHTRESERVEE, '
    + ' GEP_TOTALQTERES, '
    + ' GEP_POUQTERES, '
    + ' GEP_POUMTRES, '
    + ' GEP_TOTHTAFFECTEE, '
    + ' GEP_TOTALQTEAFF, '
    + ' GEP_POUQTEAFF, '
    + ' GEP_POUMTAFF) '
    + 'SELECT GEL_CODEAFF,GEL_NATUREPIECEG,MAX(GEL_DATEPIECE) AS GEP_DATEPIECE, '
    + ' GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG, '
    + ' SUM(GEL_MONTANTHT) AS GEP_TOTALHT,SUM(GEL_QTEALIVRER) AS GEP_TOTALQTE, '
    + ' SUM(GEL_MTHTRESERVEE) AS GEP_TOTHTRESERVEE, '
    + ' SUM(GEL_QTERESERVEE) AS GEP_TOTALQTERES, '
    + ' (SUM(GEL_QTERESERVEE) * 100) / IIF(SUM(GEL_QTEALIVRER)=0, 1, SUM(GEL_QTEALIVRER)) AS GEP_POUQTERES, '
    + ' (SUM(GEL_MTHTRESERVEE) * 100) / IIF(SUM(GEL_MONTANTHT)=0, 1, SUM(GEL_MONTANTHT)) AS GEP_POUMTRES, '
    + ' SUM(GEL_MTHTAFFECTEE) AS GEP_TOTHTAFFECTEE, '
    + ' SUM(GEL_QTEAFFECTEE) AS GEP_TOTALQTEAFF, '
    + ' (SUM(GEL_QTEAFFECTEE) * 100) / IIF(SUM(GEL_QTEALIVRER)=0, 1, SUM(GEL_QTEALIVRER)) AS GEP_POUQTEAFF, '
    + ' (SUM(GEL_MTHTAFFECTEE) * 100) / IIF(SUM(GEL_MONTANTHT)=0, 1, SUM(GEL_MONTANTHT)) AS GEP_POUMTAFF '
    + 'FROM AFFCDELIGNE '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_TYPEDIM<>"DIM" '
    + 'GROUP BY GEL_CODEAFF,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Recopie des quantités disponibles du stock dans la table
Suite ........ : spécifique en tenant compte des quantités déjà réservées.
Mots clefs ... :
*****************************************************************}

function InsertAFFCDEDISPO(CodeAff: string): integer;
var
  sSql: string;
begin
  sSql := 'INSERT INTO AFFCDEDISPO '
    + '(GED_CODEAFF, GED_DEPOT, GED_ARTICLE, GED_PHYSIQUE, GED_QTEPREVI, '
    + ' GED_QTERESERVEE, GED_QTEAFFECTEE, '
    + ' GED_DISPORESERVEE, '
    + ' GED_DISPOAFFECTEE) '
    + 'SELECT DISTINCT GEL_CODEAFF, GEA_DEPOT, GEL_ARTICLE, '
    + ' GQ_PHYSIQUE, GQ_PHYSIQUE - GQ_PREPACLI AS GED_QTEPREVI, '
    + ' 0 AS GED_QTERESERVE, 0 AS GED_QTEAFFECTEE, '
    + ' GQ_PHYSIQUE - GQ_PREPACLI AS GED_QTERESERVEE, '
    + ' GQ_PHYSIQUE - GQ_PREPACLI AS GED_DISPOAFFECTEE '
    + 'FROM AFFCDELIGNE, AFFCDEENTETE, DISPO '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" '
    + ' AND AFFCDEENTETE.GEA_CODEAFF=AFFCDELIGNE.GEL_CODEAFF '
    + ' AND DISPO.GQ_ARTICLE=AFFCDELIGNE.GEL_ARTICLE '
    + ' AND DISPO.GQ_DEPOT=AFFCDEENTETE.GEA_DEPOT '
    + ' AND DISPO.GQ_CLOTURE="-" '
    + ' AND DISPO.GQ_DATECLOTURE="01/01/1900" ';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Mise à jour des quantités disponibles du stock dans la table
Suite ........ : spécifique en tenant compte des quantités déjà réservées.
Mots clefs ... :
*****************************************************************}

function UpdateAFFCDEDISPO(CodeAff: string): integer;
var
  sSql: string;
  QQ: TQuery;
  TOBDispo: TOB;
  Ind, MaxNbRow: integer;
  Ok: boolean;
begin
  // suppression du stock existant
  DeleteAFFCDEDISPO(CodeAff);
  // recopie du stock réel
  sSql := 'INSERT INTO AFFCDEDISPO '
    + '(GED_CODEAFF, GED_DEPOT, GED_ARTICLE, GED_PHYSIQUE, GED_QTEPREVI, '
    + ' GED_QTERESERVEE, GED_QTEAFFECTEE, '
    + ' GED_DISPORESERVEE, '
    + ' GED_DISPOAFFECTEE) '
    + 'SELECT DISTINCT GEL_CODEAFF, GEA_DEPOT, GEL_ARTICLE, '
    + ' GQ_PHYSIQUE, GQ_PHYSIQUE - GQ_PREPACLI AS GED_QTEPREVI, '
    + ' 0 AS GED_QTERESERVE, 0 AS GED_QTEAFFECTEE, '
    + ' GQ_PHYSIQUE - GQ_PREPACLI AS GED_QTERESERVEE, '
    + ' GQ_PHYSIQUE - GQ_PREPACLI AS GED_DISPOAFFECTEE '
    + 'FROM AFFCDELIGNE, AFFCDEENTETE, DISPO '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" '
    + ' AND AFFCDEENTETE.GEA_CODEAFF=AFFCDELIGNE.GEL_CODEAFF '
    + ' AND DISPO.GQ_ARTICLE=AFFCDELIGNE.GEL_ARTICLE '
    + ' AND DISPO.GQ_DEPOT=AFFCDEENTETE.GEA_DEPOT '
    + ' AND DISPO.GQ_CLOTURE="-" '
    + ' AND DISPO.GQ_DATECLOTURE="01/01/1900" ';
  Result := ExecuteSQLAFFCDE(sSql);
  // cumul des quantités réservées et affectées
  sSql := 'SELECT MAX(GEL_CODEAFF) AS GED_CODEAFF, MAX(GEA_DEPOT) AS GED_DEPOT, '
    + 'GEL_ARTICLE AS GED_ARTICLE, MAX(GED_PHYSIQUE) AS GED_PHYSIQUE, '
    + ' MAX(GED_QTEPREVI) AS GED_QTEPREVI, '
    + ' SUM(GEL_QTERESERVEE) AS GED_QTERESERVEE, '
    + ' SUM(GEL_QTEAFFECTEE) AS GED_QTEAFFECTEE, '
    + ' (MAX(GED_QTEPREVI)-SUM(GEL_QTERESERVEE)) AS GED_DISPORESERVEE, '
    + ' (MAX(GED_QTEPREVI)-SUM(GEL_QTEAFFECTEE)) AS GED_DISPOAFFECTEE '
    + 'FROM AFFCDELIGNE, AFFCDEENTETE, AFFCDEDISPO '
    + 'WHERE GEL_CODEAFF="' + CodeAff + '" '
    + ' AND AFFCDEENTETE.GEA_CODEAFF=AFFCDELIGNE.GEL_CODEAFF '
    + ' AND AFFCDEDISPO.GED_CODEAFF=AFFCDELIGNE.GEL_CODEAFF '
    + ' AND AFFCDEDISPO.GED_DEPOT=AFFCDEENTETE.GEA_DEPOT '
    + ' AND AFFCDEDISPO.GED_ARTICLE=AFFCDELIGNE.GEL_ARTICLE '
    + 'GROUP BY GEL_ARTICLE';
  QQ := OpenSQl(sSql, True);
  Ok := (not QQ.Eof);
  TOBDispo := TOB.Create('', nil, -1);
  Ind := 0;
  MaxNbRow := 1000;
  while Ok do
  begin
    Ok := TOBDispo.LoadDetailDB('AFFCDEDISPO', '', '', QQ, False, False, MaxNbRow, Ind);
    if Ok then
    begin
      Ind := 1; // pour ne pas traiter 2 fois le dernier enregistrement
      TOBDispo.SetAllModifie(True);
      TOBDispo.UpdateDB;
    end;
  end;
  TOBDispo.Free;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Suppression des enregistrements dans la table spécifique
Suite ........ : du stock
Mots clefs ... :
*****************************************************************}

function DeleteAFFCDEDISPO(CodeAff: string): integer;
var
  sSql: string;
begin
  sSql := 'DELETE AFFCDEDISPO WHERE GED_CODEAFF="' + CodeAff + '"';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Suppression des enregistrements dans la table spécifique
Suite ........ : des commandes
Mots clefs ... :
*****************************************************************}

function DeleteAFFCDEPIECE(CodeAff: string): integer;
var
  sSql: string;
begin
  sSql := 'DELETE AFFCDEPIECE WHERE GEP_CODEAFF="' + CodeAff + '"';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Suppression des enregistrements dans la table spécifique
Suite ........ : des lignes de commandes
Mots clefs ... :
*****************************************************************}

function DeleteAFFCDELIGNE(CodeAff, TypeDim: string): integer;
var
  sSql: string;
begin
  sSql := 'DELETE AFFCDELIGNE WHERE GEL_CODEAFF="' + CodeAff + '"';
  if TypeDim <> '' then sSql := sSql + ' AND GEL_TYPEDIM="' + TypeDim + '"';
  Result := ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 10/12/2002
Modifié le ... : 10/12/2002
Description .. : Vérifie le statut de l'affectation de commandes avant de
Suite ........ : lancer une étape.
Mots clefs ... :
*****************************************************************}

function VerifStatutAFFCDEENTETE(CodeAff: string; Etape: TTypeEtapeAffCde): boolean;
var
  sSql, CodeStatut: string;
  Termine: boolean;
  QQ: TQuery;
begin
  Result := False;
  sSql := 'SELECT GEA_STATUTAFF,GEA_TERMINE FROM AFFCDEENTETE'
    + ' WHERE GEA_CODEAFF="' + CodeAff + '"';
  QQ := OpenSQl(sSql, True);
  if not QQ.Eof then
  begin
    Termine := (QQ.FindField('GEA_TERMINE').AsString = 'X');
    CodeStatut := QQ.FindField('GEA_STATUTAFF').AsString;
    if CodeStatut = '' then Termine := True; // jamais exécuté
    case Etape of
      afcSelection:
        begin
          // Sélection des lignes de commandes
          Result := ((Termine) and ((CodeStatut = '') or (CodeStatut = '010') or
            (CodeStatut = '050')));
        end;
      afcReservation:
        begin
          // Réservation des commandes
          Result := ((Termine) and (CodeStatut = '010'));
        end;
      afcAffectation:
        begin
          // Affectation des commandes
          Result := (((Termine) and (CodeStatut = '020')) or
            ((Termine) and (CodeStatut = '030')) or
            ((Termine) and (CodeStatut = '040')));
        end;
      afcPreparation:
        begin
          // Préparation des commandes
          Result := (((Termine) and (CodeStatut = '030')) or
            ((Termine) and (CodeStatut = '040')));
        end;
      afcFin:
        begin
          // Fin de l'affectation des commandes
          Result := True;
        end;
    end;
  end;
  Ferme(QQ);
  if not Result then
    PgiBox('Lancement impossible : l''étape précédente n''a pas été effectuée.');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 28/11/2002
Modifié le ... : 28/11/2002
Description .. : Mise à jour du statut de l'affectation de commandes
Mots clefs ... :
*****************************************************************}

procedure MajStatutAFFCDEENTETE(CodeAff, CodeStatut: string; Termine: boolean);
var
  sSql, Stg: string;
begin
  if Termine then Stg := 'X' else Stg := '-';
  sSql := 'UPDATE AFFCDEENTETE SET'
    + ' GEA_STATUTAFF="' + CodeStatut + '",'
    + ' GEA_DATEEXECUTION="' + USTime(NowH) + '",'
    + ' GEA_TERMINE="' + Stg + '"'
    + ' WHERE GEA_CODEAFF="' + CodeAff + '"';
  ExecuteSQLAFFCDE(sSql);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 02/12/2002
Description .. : Détermine le code statut en fonction de l'étape en cours
Suite ........ : d'exécution
Mots clefs ... :
*****************************************************************}

function DonneCodeStatutAFFCDE(Etape: TTypeEtapeAffCde): string;
begin
  case Etape of
    afcSelection:
      Result := '010';
    afcReservation:
      Result := '020';
    afcAffectation:
      Result := '030';
    afcPreparation:
      Result := '040';
    afcFin:
      Result := '050';
  else
    Result := '';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Donne la quantité disponible en stock
Mots clefs ... :
*****************************************************************}

function GetQteStockAFFCDE(TOBStock: TOB; Etape: TTypeEtapeAffCde): double;
var
  NomChamp: string;
begin
  Result := 0;
  if TOBStock = nil then Exit;
  if Etape = afcReservation then
    NomChamp := 'GED_DISPORESERVEE'
  else
    if Etape = afcAffectation then
    NomChamp := 'GED_DISPOAFFECTEE'
  else
    NomChamp := 'GED_QTEPREVI';
  Result := TOBStock.GetValue(NomChamp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour des quantités du stock
Mots clefs ... :
*****************************************************************}

procedure MajStockAFFCDE(TOBStock: TOB; QteLivrer: double; Etape: TTypeEtapeAffCde; AncienneQte: double = 0; ForceMaj: boolean = False);
var
  QteStock, QteDisponible: double;
  NomChamp: string;
begin
  if TOBStock = nil then Exit;
  if (QteLivrer <= 0) and (not ForceMaj) then Exit;
  // incrémentation de la quantité réservée
  if Etape = afcReservation then
    NomChamp := 'GED_QTERESERVEE'
  else
    NomChamp := 'GED_QTEAFFECTEE';
  QteStock := TOBStock.GetValue(NomChamp);
  QteStock := (QteStock - AncienneQte) + QteLivrer;
  TOBStock.PutValue(NomChamp, QteStock);
  // décrementation de la quantité disponible
  QteDisponible := TOBStock.GetValue('GED_QTEPREVI');
  QteDisponible := QteDisponible - QteStock ;
  if Etape = afcReservation then
    NomChamp := 'GED_DISPORESERVEE'
  else
    NomChamp := 'GED_DISPOAFFECTEE';
  TOBStock.PutValue(NomChamp, QteDisponible);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Chargement en TOB des paramètres de l'affectation.
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeTobAffect;
begin
  TOBAffectation.SelectDB('"' + CodeAff + '"', nil);
  Depot := TOBAffectation.GetValue('GEA_DEPOT');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Recherche des ordres de priorité donc des ordres de tri
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MakeOrderBy(var sOrder: string; var TypeJoin: TTypeJointures);
var
  Ind: integer;
  NomChamp, Prfx: string;
begin
  //if Etape in [afcReservation, afcAffectation] then sOrder := 'GEL_MODIFIABLE';
  if Etape = afcAffectation then
  begin
    sOrder := 'GEL_MODIFIABLE,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG,GEL_NUMLIGNE';
  end else
  begin
    // on traite d'abord les lignes avec une quantité réservée ou affectée
    if Etape = afcReservation then
      sOrder := 'GEL_MODIFIABLE'
    else
      sOrder := '';
    if TOBAffectation <> nil then
    begin
      for Ind := 1 to NBPRIORAFFCDE do
      begin
        NomChamp := TOBAffectation.GetValue('GEA_PRIOAFFCDE' + IntToStr(Ind));
        if NomChamp <> '' then NomChamp := RechDom('GCAFFCDEPRIORITE', NomChamp, True);
        if NomChamp <> '' then
        begin
          if Copy(NomChamp, 1, 15) = 'TABLELIBRETIERS' then
            NomChamp := 'YTC_' + NomChamp;
          if sOrder <> '' then sOrder := sOrder + ',';
          sOrder := sOrder + NomChamp;
          if TOBAffectation.GetValue('GEA_PRIODESC' + IntToStr(Ind)) = 'X' then
            sOrder := sOrder + ' DESC';
          // jointure
          Prfx := ExtractPrefixe(NomChamp);
          if Prfx = 'YTC' then TypeJoin := TypeJoin + [joinTIERSCOMPL] else
          if Prfx = 'GP' then TypeJoin := TypeJoin + [joinPIECE] else
          if Prfx = 'GA' then TypeJoin := TypeJoin + [joinARTICLE] else
          if Prfx = 'T' then TypeJoin := TypeJoin + [joinTIERS];
        end;
      end;
    end;
  end;
  if sOrder <> '' then sOrder := ' ORDER BY ' + sOrder;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Chargement des besoins c'est-à-dire des lignes de
Suite ........ : commandes en TOB.
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeBesoin;
var
  sSql, sOrder: string;
  TypeJoin: TTypeJointures;
begin
  TypeJoin := [];
  if VerifCliFerme then TypeJoin := TypeJoin + [joinTIERS];
  if VerifArtFerme then TypeJoin := TypeJoin + [joinARTICLE];
  // Ordre de tri
  MakeOrderBy(sOrder, TypeJoin);
  // Sélection des lignes avec si besoin jointure sur le complément tiers
  sSql := 'SELECT GEL_CODEAFF,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG,'
    + ' GEL_NUMLIGNE,GEL_TIERS,GEL_ARTICLE,GEL_CODEARTICLE,GEL_CODESDIM,'
    + ' GEL_TYPEDIM,GEL_NUMLIGNEGEN,GEL_TENUESTOCK,GEL_QTEALIVRER,GEL_QTERESERVEE,'
    + ' GEL_QTEAFFECTEE,GEL_MONTANTHT,GEL_MTHTRESERVEE,GEL_MTHTAFFECTEE,'
    + ' GEL_MODIFIABLE';
  if VerifCliFerme then sSql := sSql + ',T_FERME';
  if VerifArtFerme then sSql := sSql + ',GA_FERME,GA_INTERDITVENTE';
  sSql := sSql + ' FROM AFFCDELIGNE';
  // Jointures
  if joinTIERS in TypeJoin then
    sSql := sSql + ' LEFT JOIN TIERS ON GEL_TIERS=T_TIERS';
  if joinARTICLE in TypeJoin then
    sSql := sSql + ' LEFT JOIN ARTICLE ON GEL_ARTICLE=GA_ARTICLE';
  if joinTIERSCOMPL in TypeJoin then
    sSql := sSql + ' LEFT JOIN TIERSCOMPL ON GEL_TIERS=YTC_TIERS';
  if joinPIECE in TypeJoin then
    sSql := sSql + ' LEFT JOIN PIECE ON GEL_NATUREPIECEG=GP_NATUREPIECEG '
      + ' AND GEL_SOUCHE=GP_SOUCHE AND GEL_NUMERO=GP_NUMERO '
      + ' AND GEL_INDICEG=GP_INDICEG';
  // Clause where
  sSql := sSql + ' WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_TYPEDIM<>"GEN" '
    + sOrder;
  // chargement de la TOB
  QQBesoin := OpenSQl(sSql, True);
  IndBesoin := 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 11/02/2003
Description .. : Chargement des lignes de commandes à traiter en TOB.
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.ChargeBesoinEnTob: boolean;
begin
  Result := False;
  if (QQBesoin <> nil) and (not QQBesoin.Eof) then
  begin
    Result := TOBBesoin.LoadDetailDB('AFFCDELIGNE', '', '', QQBesoin, False, False, MaxNbRow, IndBesoin);
    IndBesoin := 1; // pour ne pas traiter 2 fois le dernier enregistrement
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Chargement des quantités disponibles en TOB.
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeTobDispo;
var
  sSql: string;
  QQ: TQuery;
begin
  sSql := 'SELECT GED_CODEAFF,GED_DEPOT,GED_ARTICLE,GED_QTEPREVI';
  if Etape = afcReservation then
    sSql := sSql + ',GED_QTERESERVEE,GED_DISPORESERVEE'
  else
    sSql := sSql + ',GED_QTEAFFECTEE,GED_DISPOAFFECTEE';
  if UseDataMem then sSql := sSql + ',"-" AS MODIF';
  sSql := sSql + ' FROM AFFCDEDISPO WHERE GED_CODEAFF="' + CodeAff + '"'
    + ' AND GED_DEPOT="' + Depot + '"'
    + ' ORDER BY GED_CODEAFF,GED_DEPOT,GED_ARTICLE';

  if UseDataMem then
  begin
    dm_sqlInit(IDM_AFFCDEDISPO, 0, sSql, ['A', 'A', 'A'], True,
      [RC_CODEAFF, RC_DEPOT, RC_ARTICLE]);
  end else
  begin
    QQ := OpenSQl(sSql, True);
    if not QQ.EOF then TOBDispo.LoadDetailDB('AFFCDEDISPO', '', '', QQ, False);
    Ferme(QQ);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 22/11/2002
Description .. : Chargement des pièces en TOB.
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeTobPiece;
var
  sSql: string;
  QQ: TQuery;
begin
  sSql := 'SELECT GEP_CODEAFF,GEP_NATUREPIECEG,GEP_SOUCHE,GEP_NUMERO,GEP_INDICEG,'
    + ' GEP_TOTALHT,GEP_TOTALQTE';
  if Etape = afcReservation then
    sSql := sSql + ',GEP_TOTHTRESERVEE,GEP_TOTALQTERES,GEP_POUQTERES,GEP_POUMTRES'
  else
    sSql := sSql + ',GEP_TOTHTAFFECTEE,GEP_TOTALQTEAFF,GEP_POUQTEAFF,GEP_POUMTAFF';
  if UseDataMem then sSql := sSql + ',"-" AS MODIF';
  sSql := sSql + ' FROM AFFCDEPIECE WHERE GEP_CODEAFF="' + CodeAff + '"'
    + ' ORDER BY GEP_CODEAFF,GEP_NATUREPIECEG,GEP_SOUCHE,GEP_NUMERO,GEP_INDICEG';

  if UseDataMem then
  begin
    dm_sqlInit(IDM_AFFCDEPIECE, 0, sSql, ['A', 'A', 'A', 'N', 'N'], True,
      [RC_CODEAFF, RC_NATUREPIECEG, RC_SOUCHE, RC_NUMERO, RC_INDICEG]);
  end else
  begin
    QQ := OpenSQl(sSql, True);
    if not QQ.EOF then TOBPiece.LoadDetailDB('AFFCDEPIECE', '', '', QQ, False);
    Ferme(QQ);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 11/02/2003
Modifié le ... : 11/02/2003
Description .. : Chargement des lignes génériques en TOB.
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeTobLgGen;
var
  sSql: string;
  QQ: TQuery;
begin
  sSql := 'SELECT GEL_CODEAFF,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO,GEL_INDICEG,'
    + ' GEL_NUMLIGNE,GEL_QTEALIVRER,GEL_MONTANTHT';
  if Etape = afcReservation then
    sSql := sSql + ',GEL_QTERESERVEE,GEL_MTHTRESERVEE'
  else
    sSql := sSql + ',GEL_QTEAFFECTEE,GEL_MTHTAFFECTEE';
  if UseDataMem then sSql := sSql + ',"-" AS MODIF';
  sSql := sSql + ' FROM AFFCDELIGNE WHERE GEL_CODEAFF="' + CodeAff + '"'
    + ' AND GEL_TYPEDIM="GEN"'
    + ' ORDER BY GEL_CODEAFF,GEL_NATUREPIECEG,GEL_SOUCHE,GEL_NUMERO,'
    + ' GEL_INDICEG,GEL_NUMLIGNE';

  if UseDataMem then
  begin
    dm_sqlInit(IDM_AFFCDELIGNE, 0, sSql, ['A', 'A', 'A', 'N', 'N', 'N'], True,
      [RC_CODEAFF, RC_NATUREPIECEG, RC_SOUCHE, RC_NUMERO, RC_INDICEG, RC_NUMLIGNE]);
  end else
  begin
    QQ := OpenSQl(sSql, True);
    if not QQ.EOF then TOBLgGen.LoadDetailDB('AFFCDELIGNE', '', '', QQ, False);
    Ferme(QQ);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Vérification si l'article est utilisable pour la nature de piéce
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.ArticleAutorise(TOBLigne: TOB; NaturePiece: string): boolean;
var
  TOBNat, TOBG: TOB;
  RefUnique: string;
  isFerme: boolean;
begin
  Result := True;
  if not VerifArtFerme then Exit;
  if TOBLigne = nil then Exit;
  RefUnique := TOBLigne.GetValue('GEL_ARTICLE');
  if RefUnique = '' then Exit;
  // article générique
  if TOBLigne.GetValue('GEL_TYPEDIM') = 'GEN' then
  begin
    Result := False;
    Exit;
  end;
  // article fermé
  isFerme := (TOBLigne.GetValue('GA_FERME') = 'X');
  if isFerme then
  begin
    TOBNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'], [NaturePiece], False);
    if TOBNat <> nil then
    begin
      TOBG := TOBNat.FindFirst(['GAP_NATUREPIECEG', 'GAP_ARTICLE'], [NaturePiece, RefUnique], False);
      if TOBG = nil then
        Result := False
      else Result := (TOBG.GetValue('GAP_SUSPENSION') = 'X');
    end;
  end;
  // article interdit à la vente
  if Result then
  begin
    TOBNat := VH_GC.TOBParPiece.FindFirst(['GPP_NATUREPIECEG'], [NaturePiece], False);
    if TOBNat <> nil then
    begin
      if TOBNat.GetValue('GPP_VENTEACHAT') = 'VEN' then
        Result := (TOBLigne.GetValue('GA_INTERDITVENTE') = 'X');
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Vérification si le tiers est utilisable pour la nature de piéce
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.TiersAutorise(TOBLigne: TOB; NaturePiece: string): boolean;
var
  TOBT: TOB;
  CodeTiers: string;
  isFerme: boolean;
begin
  Result := True;
  if not VerifCliFerme then Exit;
  if TOBLigne = nil then Exit;
  CodeTiers := TOBLigne.GetValue('GEL_TIERS');
  if CodeTiers = '' then Exit;
  isFerme := (TOBLigne.GetValue('T_FERME') = 'X');
  if not isFerme then Exit;
  TOBT := TOBTiersPiece.FindFirst(['GTP_TIERS', 'GTP_NATUREPIECEG'], [CodeTiers, NaturePiece], False);
  if TOBT = nil then
  begin
    TOBT := TOB.Create('TIERSPIECE', TOBTiersPiece, -1);
    if not TOBT.SelectDB('"' + CodeTiers + '";"' + NaturePiece + '"', nil) then
    begin
      TOBT.InitValeurs;
      TOBT.PutValue('GTP_TIERS', CodeTiers);
      TOBT.PutValue('GTP_NATUREPIECEG', NaturePiece);
      TOBT.PutValue('GTP_SUSPENSION', '-');
    end;
  end;
  if TOBT = nil then
    Result := False
  else Result := (TOBT.GetValue('GTP_SUSPENSION') = 'X');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour de la ligne générique
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajLigneGenerique(TOBLigne: TOB; QteLivrer: double);
var
  TOBG: TOB;
begin
  if Etape = afcReservation then Exit;
  TOBG := FindLigneGen(TOBLigne);
  if TOBG <> nil then
  begin
    TOBGen := TOBG;
    MajLigneGeneriqueAff(TOBLgGen, TOBLigne, TOBGen, Etape, QteLivrer, False);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Recherche de la ligne générique
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.FindLigneGen(TOBLigne: TOB): TOB;
var
  NoErr: integer;
  Enr: TEnreg;
  Trouve: boolean;
begin
  Result := nil;
  if TOBLigne.GetValue('GEL_TYPEDIM') <> 'DIM' then Exit;
  // recherche de la ligne générique de l'article
  if UseDataMem then
  begin
    if TOBGen <> nil then
    begin
      // si la ligne générique courante est celle recherchée
      Trouve := ((TOBLigne.GetValue('GEL_NATUREPIECEG') = TOBGen.GetValue('GEL_NATUREPIECEG')) and
        (TOBLigne.GetValue('GEL_SOUCHE') = TOBGen.GetValue('GEL_SOUCHE')) and
        (TOBLigne.GetValue('GEL_NUMERO') = TOBGen.GetValue('GEL_NUMERO')) and
        (TOBLigne.GetValue('GEL_INDICEG') = TOBGen.GetValue('GEL_INDICEG')) and
        (TOBLigne.GetValue('GEL_NUMLIGNEGEN') = TOBGen.GetValue('GEL_NUMLIGNE')));
      if Trouve then
      begin
        Result := TOBGen;
        Exit;
      end;
      FreeAndNil(TOBGen);
    end;
    NoErr := dm_trouveEnr(IDM_AFFCDELIGNE, [CodeAff,
      TOBLigne.GetValue('GEL_NATUREPIECEG'),
        TOBLigne.GetValue('GEL_SOUCHE'),
        IntToStr(TOBLigne.GetValue('GEL_NUMERO')),
        IntToStr(TOBLigne.GetValue('GEL_INDICEG')),
        IntToStr(TOBLigne.GetValue('GEL_NUMLIGNEGEN'))], Enr);
    if NoErr = 0 then
    begin
      Result := TOB.Create('AFFCDELIGNE', nil, -1);
      Result.InitValeurs;
      Result.PutValue('GEL_CODEAFF', Enr.ch(GEL_CODEAFF));
      Result.PutValue('GEL_NATUREPIECEG', Enr.ch(GEL_NATUREPIECEG));
      Result.PutValue('GEL_SOUCHE', Enr.ch(GEL_SOUCHE));
      Result.PutValue('GEL_NUMERO', StrToInt(Enr.ch(GEL_NUMERO)));
      Result.PutValue('GEL_INDICEG', StrToInt(Enr.ch(GEL_INDICEG)));
      Result.PutValue('GEL_NUMLIGNE', StrToInt(Enr.ch(GEL_NUMLIGNE)));
      Result.PutValue('GEL_QTEALIVRER', Valeur(Enr.ch(GEL_QTEALIVRER)));
      Result.PutValue('GEL_MONTANTHT', Valeur(Enr.ch(GEL_MONTANTHT)));
      if Etape = afcReservation then
      begin
        Result.PutValue('GEL_QTERESERVEE', Valeur(Enr.ch(GEL_QTERESERVEE)));
        Result.PutValue('GEL_MTHTRESERVEE', Valeur(Enr.ch(GEL_MTHTRESERVEE)));
      end else
      begin
        Result.PutValue('GEL_QTEAFFECTEE', Valeur(Enr.ch(GEL_QTERESERVEE)));
        Result.PutValue('GEL_MTHTAFFECTEE', Valeur(Enr.ch(GEL_MTHTRESERVEE)));
      end;
    end;
  end else
  begin
    if TOBLigne.GetValue('GEL_NUMLIGNEGEN') > 0 then
      Result := TrouveNumLigneGeneriqueAff(TOBLgGen, TOBLigne, TOBGen)
    else
      Result := TrouveLigneGeneriqueAff(TOBLgGen, TOBLigne);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Recherche de la pièce
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.FindPiece(TOBLigne: TOB): TOB;
var
  Ind, NoErr: integer;
  NomChamp: string;
  LstChamp: array[1..4] of string;
  LstValeur: array[1..4] of variant;
  Trouve: boolean;
  Enr: TEnreg;
begin
  Result := nil;
  if TOBLigne = nil then Exit;
  if TOBCde = nil then
    Trouve := False
  else
    Trouve := True;
  // critères de recherche de la pièce
  for Ind := Low(LstChamp) to High(LstChamp) do
  begin
    case Ind of
      1: NomChamp := 'GEP_NATUREPIECEG';
      2: NomChamp := 'GEP_SOUCHE';
      3: NomChamp := 'GEP_NUMERO';
      4: NomChamp := 'GEP_INDICEG';
    end;
    LstChamp[Ind] := NomChamp;
    LstValeur[Ind] := TOBLigne.GetValue('GEL_' + ExtractSuffixe(NomChamp));
    // vérifie si la pièce précédente est celle recherchée
    if (Trouve) and (TOBCde <> nil) then
    begin
      if TOBCde.GetValue(NomChamp) <> LstValeur[Ind] then
        Trouve := False;
    end;
  end;
  if not Trouve then
  begin
    if UseDataMem then
    begin
      NoErr := dm_trouveEnr(IDM_AFFCDEPIECE, [CodeAff,
        TOBLigne.GetValue('GEL_NATUREPIECEG'),
          TOBLigne.GetValue('GEL_SOUCHE'),
          IntToStr(TOBLigne.GetValue('GEL_NUMERO')),
          IntToStr(TOBLigne.GetValue('GEL_INDICEG'))], Enr);
      if NoErr = 0 then
      begin
        if TOBCde <> nil then FreeAndNil(TOBCde);
        TOBCde := TOB.Create('AFFCDEPIECE', nil, -1);
        TOBCde.InitValeurs;
        TOBCde.PutValue('GEP_CODEAFF', Enr.ch(GEP_CODEAFF));
        TOBCde.PutValue('GEP_NATUREPIECEG', Enr.ch(GEP_NATUREPIECEG));
        TOBCde.PutValue('GEP_SOUCHE', Enr.ch(GEP_SOUCHE));
        TOBCde.PutValue('GEP_NUMERO', StrToInt(Enr.ch(GEP_NUMERO)));
        TOBCde.PutValue('GEP_INDICEG', StrToInt(Enr.ch(GEP_INDICEG)));
        TOBCde.PutValue('GEP_TOTALHT', Valeur(Enr.ch(GEP_TOTALHT)));
        TOBCde.PutValue('GEP_TOTALQTE', Valeur(Enr.ch(GEP_TOTALQTE)));
        if Etape = afcReservation then
        begin
          TOBCde.PutValue('GEP_TOTHTRESERVEE', Valeur(Enr.ch(GEP_TOTHTRESERVEE)));
          TOBCde.PutValue('GEP_TOTALQTERES', Valeur(Enr.ch(GEP_TOTALQTERES)));
          TOBCde.PutValue('GEP_POUQTERES', Valeur(Enr.ch(GEP_POUQTERES)));
          TOBCde.PutValue('GEP_POUMTRES', Valeur(Enr.ch(GEP_POUMTRES)));
        end else
        begin
          TOBCde.PutValue('GEP_TOTHTAFFECTEE', Valeur(Enr.ch(GEP_TOTHTRESERVEE)));
          TOBCde.PutValue('GEP_TOTALQTEAFF', Valeur(Enr.ch(GEP_TOTALQTERES)));
          TOBCde.PutValue('GEP_POUQTEAFF', Valeur(Enr.ch(GEP_POUQTERES)));
          TOBCde.PutValue('GEP_POUMTAFF', Valeur(Enr.ch(GEP_POUMTRES)));
        end;
      end;
    end else
    begin
      TOBCde := TOBPiece.FindFirst(LstChamp, LstValeur, False);
    end;
  end;
  Result := TOBCde;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour de la pièce
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajPiece(TOBLigne: TOB; OldQteLiv, OldMtLiv: double);
var
  TOBP: TOB;
begin
  if Etape = afcReservation then Exit;
  if TOBLigne = nil then Exit;
  if TOBLigne.GetValue('GEL_TYPEDIM') = 'GEN' then Exit;
  // recherche de la pièce
  TOBP := FindPiece(TOBLigne);
  if TOBP <> nil then
    CalculMontantPieceAff(TOBP, TOBLigne, Etape, OldMtLiv, OldQteLiv, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Recherche du stock de l'article
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.FindStock(RefArt: string; TOBLigne: TOB): TOB;
var
  NoErr: integer;
  Enr: TEnreg;
begin
  Result := nil;
  if RefArt = '' then Exit;
  if TOBLigne.GetValue('GEL_TENUESTOCK') = '-' then Exit;
  if (TOBStk = nil) or (RefArt <> TOBStk.GetValue('GED_ARTICLE')) then
  begin
    if UseDataMem then
    begin
      NoErr := dm_trouveEnr(IDM_AFFCDEDISPO, [CodeAff, Depot, RefArt], Enr);
      if NoErr = 0 then
      begin
        if TOBStk <> nil then FreeAndNil(TOBStk);
        TOBStk := TOB.Create('AFFCDEDISPO', nil, -1);
        TOBStk.InitValeurs;
        TOBStk.PutValue('GED_CODEAFF', Enr.ch(GED_CODEAFF));
        TOBStk.PutValue('GED_DEPOT', Enr.ch(GED_DEPOT));
        TOBStk.PutValue('GED_ARTICLE', Enr.ch(GED_ARTICLE));
        TOBStk.PutValue('GED_QTEPREVI', Valeur(Enr.ch(GED_QTEPREVI)));
        if Etape = afcReservation then
        begin
          TOBStk.PutValue('GED_QTERESERVEE', Valeur(Enr.ch(GED_QTERESERVEE)));
          TOBStk.PutValue('GED_DISPORESERVEE', Valeur(Enr.ch(GED_DISPORESERVEE)));
        end else
        begin
          TOBStk.PutValue('GED_QTEAFFECTEE', Valeur(Enr.ch(GED_QTERESERVEE)));
          TOBStk.PutValue('GED_DISPOAFFECTEE', Valeur(Enr.ch(GED_DISPORESERVEE)));
        end;
      end;
    end else
    begin
      TOBStk := TOBDispo.FindFirst(['GED_ARTICLE'], [RefArt], False);
    end;
  end;
  Result := TOBStk;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour des quantités du stock
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajStock(TOBStock: TOB; QteLivrer: double);
begin
  MajStockAFFCDE(TOBStock, QteLivrer, Etape);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 03/12/2002
Description .. : Calcul de la quantité réservée en fonction des règles
Suite ........ : d'affectation
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.CalculQteLivrer(QteDispo, QteBesoin: double; TOBLigne: TOB): double;
var
  Pourc, QteMini: double;
begin
  if QteDispo < QteBesoin then
  begin
    // quantité disponible insuffisante
    if TOBAffectation.GetValue('GEA_GERERELIQUAT') = 'X' then
    begin
      // les reliquats autorisés
      Pourc := TOBAffectation.GetValue('GEA_POURCENTAGE');
      if Pourc > 0 then
      begin
        // vérification du % mini du besoin à conserver
        QteMini := QteBesoin * Pourc / 100;
        if QteDispo >= QteMini then
          Result := QteDispo
        else
          Result := 0;
      end else
      begin
        Result := QteDispo;
      end;
    end else
    begin
      // les reliquats interdits
      Result := 0;
    end;
  end else
  begin
    // quantité disponible suffisante
    Result := QteBesoin;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Calcul de la quantité réservée
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.CalculQteReservee(TOBLigne, TOBStock: TOB): boolean;
var
  QteDispo, QteBesoin: double;
  QteLiv, QteLivInit: double;
  NomChamp: string;
begin
  Result := False;
  if TOBLigne = nil then Exit;
  // recherche du besoin et du champ à alimenter appelé quantité réservée
  if Etape = afcReservation then
  begin
    QteBesoin := TOBLigne.GetValue('GEL_QTEALIVRER');
    NomChamp := 'GEL_QTERESERVEE';
  end else
  begin
    QteBesoin := TOBLigne.GetValue('GEL_QTERESERVEE');
    NomChamp := 'GEL_QTEAFFECTEE';
  end;
  if QteBesoin <= 0 then Exit;
  QteLivInit := TOBLigne.GetValue(NomChamp);
  QteLiv := QteLivInit;
  // calcul de la quantité réservée
  if TOBLigne.GetValue('GEL_TENUESTOCK') = '-' then
  begin
    // cas d'un article non tenu en stock
    if (TOBLigne.GetValue('GEL_MODIFIABLE') = 'X') and (QteBesoin > QteLiv) then
      QteLiv := QteBesoin;
  end else
  begin
    if TOBStock = nil then Exit;
    // recherche de la quantité disponible
    QteDispo := GetQteStockAFFCDE(TOBStock, Etape);
    if QteDispo < 0 then QteDispo := 0;
    if QteLiv > 0 then
    begin
      // cas d'une ligne avec une quantité réservée
      if QteLiv > QteDispo then
      begin
        // plus assez en stock
        QteLiv := CalculQteLivrer(QteDispo, QteLiv, TOBLigne);
      end else
        if (TOBLigne.GetValue('GEL_MODIFIABLE') = 'X') and (QteBesoin > QteLiv) then
      begin
        // on complète la quantité réservée à concurrence de la quantité à livrer
        QteLiv := CalculQteLivrer(QteDispo, QteBesoin, TOBLigne);
      end;
    end else
    begin
      // cas d'une ligne sans quantité réservée
      QteLiv := CalculQteLivrer(QteDispo, QteBesoin, TOBLigne);
    end;
    // décrementation du stock
    MajStock(TOBStock, QteLiv);
  end;
  // mise à jour de la quantité réservée
  if QteLivInit <> QteLiv then
  begin
    TOBLigne.PutValue(NomChamp, QteLiv);
    Result := True;
    // mise à jour du montant réservé
    CalculMontantLigne(TOBLigne, QteLivInit);
    // mise à jour de la ligne générique
    MajLigneGenerique(TOBLigne, QteLiv);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Calcul du montant réservée
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.CalculMontantLigne(TOBLigne: TOB; OldQteLiv: double);
var
  OldMont: double;
begin
  if TOBLigne = nil then Exit;
  if Etape = afcReservation then
    OldMont := TOBLigne.GetValue('GEL_MTHTRESERVEE')
  else
    OldMont := TOBLigne.GetValue('GEL_MTHTAFFECTEE');
  // Calcul du montant réservée
  CalculMontantLigneAff(TOBLigne, Etape);
  // mise à jour de la pièce
  MajPiece(TOBLigne, OldQteLiv, OldMont);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Allocation des TOB
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.AlloueLesTOB;
begin
  TOBTiersPiece := TOB.Create('Les TiersPiece', nil, -1);
  TOBAffectation := TOB.Create('AFFCDEENTETE', nil, -1);
  TOBBesoin := TOB.Create('Les besoins', nil, -1);
  if UseDataMem then
  begin
    TOBPiece := nil;
    TOBDispo := nil;
    TOBLgGen := nil;
  end else
  begin
    TOBDispo := TOB.Create('Les disponibles', nil, -1);
    TOBPiece := TOB.Create('Les pièces', nil, -1);
    TOBLgGen := TOB.Create('Les lignes génériques', nil, -1);
  end;
  TOBGen := nil;
  TOBCde := nil;
  TOBStk := nil;
  QQBesoin := nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Chargement des TOB
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.ChargeLesTOB;
begin
  ChargeTobAffect;
  ChargeTobDispo;
  ChargeTobPiece;
  ChargeTobLgGen;
  ChargeBesoin;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour des TOB
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajDesTOB;
begin
  if V_PGI.IoError = oeOk then TOBBesoin.UpdateDB(AvecInitMove);
  if V_PGI.IoError = oeOk then MajTOBDispo;
  if V_PGI.IoError = oeOk then MajTOBPiece;
  if V_PGI.IoError = oeOk then MajTOBLgGen;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour de la TOB du stock
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajTOBDispo;
var
  NoErr: integer;
  Enr: TEnreg;
  Pseq: TDMPseq;
  sSql: string;
begin
  if UseDataMem then
  begin
    NoErr := dm_pseqAll(IDM_AFFCDEDISPO, Pseq);
    if AvecInitMove then InitMove(dm_Nb_Enregistrement(IDM_AFFCDEDISPO), '');
    while (NoErr = 0) and (dm_pseqLect(Pseq, Enr) = 0) do
    begin
      if AvecInitMove then MoveCur(False);
      if Enr.Ch(GED_MODIF) = 'X' then
      begin
        sSql := 'UPDATE AFFCDEDISPO SET';
        if Etape = afcReservation then
        begin
          sSql := sSql + ' GED_QTERESERVEE=' + Enr.Ch(GED_QTERESERVEE)
            + ',GED_DISPORESERVEE=' + Enr.Ch(GED_DISPORESERVEE);
        end else
        begin
          sSql := sSql + ' GED_QTEAFFECTEE=' + Enr.Ch(GED_QTERESERVEE)
            + ',GED_DISPOAFFECTEE=' + Enr.Ch(GED_DISPORESERVEE);
        end;
        sSql := sSql + ' WHERE GED_CODEAFF="' + Enr.Ch(GED_CODEAFF) + '"'
          + ' AND GED_DEPOT="' + Enr.Ch(GED_DEPOT) + '"'
          + ' AND GED_ARTICLE="' + Enr.Ch(GED_ARTICLE) + '"';
        ExecuteSQL(sSql);
        dm_modifchcour(IDM_AFFCDEDISPO, GED_MODIF, '-');
      end;
    end;
    if AvecInitMove then FiniMove;
  end else
  begin
    TOBDispo.UpdateDB(AvecInitMove);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour de la TOB des pîèces
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajTOBPiece;
var
  NoErr: integer;
  Enr: TEnreg;
  Pseq: TDMPseq;
  sSql: string;
begin
  if UseDataMem then
  begin
    NoErr := dm_pseqAll(IDM_AFFCDEPIECE, Pseq);
    if AvecInitMove then InitMove(dm_Nb_Enregistrement(IDM_AFFCDEPIECE), '');
    while (NoErr = 0) and (dm_pseqLect(Pseq, Enr) = 0) do
    begin
      if AvecInitMove then MoveCur(False);
      if Enr.Ch(GEP_MODIF) = 'X' then
      begin
        sSql := 'UPDATE AFFCDEPIECE SET';
        if Etape = afcReservation then
        begin
          sSql := sSql + ' GEP_TOTHTRESERVEE=' + Enr.Ch(GEP_TOTHTRESERVEE)
            + ',GEP_TOTALQTERES=' + Enr.Ch(GEP_TOTALQTERES)
            + ',GEP_POUQTERES=' + Enr.Ch(GEP_POUQTERES)
            + ',GEP_POUMTRES=' + Enr.Ch(GEP_POUMTRES);
        end else
        begin
          sSql := sSql + ' GEP_TOTHTAFFECTEE=' + Enr.Ch(GEP_TOTHTRESERVEE)
            + ',GEP_TOTALQTEAFF=' + Enr.Ch(GEP_TOTALQTERES)
            + ',GEP_POUQTEAFF=' + Enr.Ch(GEP_POUQTERES)
            + ',GEP_POUMTAFF=' + Enr.Ch(GEP_POUMTRES);
        end;
        sSql := sSql + ' WHERE GEP_CODEAFF="' + Enr.Ch(GEP_CODEAFF) + '"'
          + ' AND GEP_NATUREPIECEG="' + Enr.Ch(GEP_NATUREPIECEG) + '"'
          + ' AND GEP_SOUCHE="' + Enr.Ch(GEP_SOUCHE) + '"'
          + ' AND GEP_NUMERO="' + Enr.Ch(GEP_NUMERO) + '"'
          + ' AND GEP_INDICEG="' + Enr.Ch(GEP_INDICEG) + '"';
        ExecuteSQL(sSql);
        dm_modifchcour(IDM_AFFCDEPIECE, GEP_MODIF, '-');
      end;
    end;
    if AvecInitMove then FiniMove;
  end else
  begin
    TOBPiece.UpdateDB(AvecInitMove);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour de la TOB des lignes génériques
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajTOBLgGen;
var
  NoErr: integer;
  Enr: TEnreg;
  Pseq: TDMPseq;
  sSql: string;
begin
  if UseDataMem then
  begin
    NoErr := dm_pseqAll(IDM_AFFCDELIGNE, Pseq);
    if AvecInitMove then InitMove(dm_Nb_Enregistrement(IDM_AFFCDELIGNE), '');
    while (NoErr = 0) and (dm_pseqLect(Pseq, Enr) = 0) do
    begin
      if AvecInitMove then MoveCur(False);
      if Enr.Ch(GEL_MODIF) = 'X' then
      begin
        sSql := 'UPDATE AFFCDELIGNE SET';
        if Etape = afcReservation then
        begin
          sSql := sSql + ' GEL_QTERESERVEE=' + Enr.Ch(GEL_QTERESERVEE)
            + ',GEL_MTHTRESERVEE=' + Enr.Ch(GEL_MTHTRESERVEE);
        end else
        begin
          sSql := sSql + ' GEL_QTEAFFECTEE=' + Enr.Ch(GEL_QTERESERVEE)
            + ',GEL_MTHTAFFECTEE=' + Enr.Ch(GEL_MTHTRESERVEE);
        end;
        sSql := sSql + ' WHERE GEL_CODEAFF="' + Enr.Ch(GEL_CODEAFF) + '"'
          + ' AND GEL_NATUREPIECEG="' + Enr.Ch(GEL_NATUREPIECEG) + '"'
          + ' AND GEL_SOUCHE="' + Enr.Ch(GEL_SOUCHE) + '"'
          + ' AND GEL_NUMERO="' + Enr.Ch(GEL_NUMERO) + '"'
          + ' AND GEL_INDICEG="' + Enr.Ch(GEL_INDICEG) + '"'
          + ' AND GEL_NUMLIGNE="' + Enr.Ch(GEL_NUMLIGNE) + '"';
        ExecuteSQL(sSql);
        dm_modifchcour(IDM_AFFCDELIGNE, GEL_MODIF, '-');
      end;
    end;
    if AvecInitMove then FiniMove;
  end else
  begin
    TOBLgGen.UpdateDB(AvecInitMove);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Vide les TOB
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.VideLesTOB;
begin
  if TOBBesoin <> nil then TOBBesoin.ClearDetail;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Libération des TOB
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.LibereLesTOB;
begin
  if QQBesoin <> nil then
  begin
    Ferme(QQBesoin);
    QQBesoin := nil;
  end;
  if TOBTiersPiece <> nil then FreeAndNil(TOBTiersPiece);
  if TOBAffectation <> nil then FreeAndNil(TOBAffectation);
  if TOBPiece <> nil then FreeAndNil(TOBPiece);
  if TOBBesoin <> nil then FreeAndNil(TOBBesoin);
  if TOBDispo <> nil then FreeAndNil(TOBDispo);
  if TOBLgGen <> nil then FreeAndNil(TOBLgGen);
  if UseDataMem then
  begin
    dm_supprime(IDM_AFFCDEDISPO);
    dm_supprime(IDM_AFFCDEPIECE);
    dm_supprime(IDM_AFFCDELIGNE);
    if TOBStk <> nil then FreeAndNil(TOBStk);
    if TOBCde <> nil then FreeAndNil(TOBCde);
    if TOBGen <> nil then FreeAndNil(TOBGen);
  end else
  begin
    TOBGen := nil;
    TOBCde := nil;
    TOBStk := nil;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 12/02/2003
Modifié le ... : 12/02/2003
Description .. : Mise à jour des Datamem à partir des TOBs
Mots clefs ... :
*****************************************************************}

procedure T_AFFECTATIONCDE.MajDataMem;
begin
  if not UseDataMem then Exit;
  // mise à jour du datamem stock
  if TOBStk <> nil then
  begin
    if Etape = afcReservation then
    begin
      dm_modifchcour(IDM_AFFCDEDISPO, GED_QTERESERVEE, StrfPoint(TOBStk.GetValue('GED_QTERESERVEE')));
      dm_modifchcour(IDM_AFFCDEDISPO, GED_DISPORESERVEE, StrfPoint(TOBStk.GetValue('GED_DISPORESERVEE')));
    end else
    begin
      dm_modifchcour(IDM_AFFCDEDISPO, GED_QTERESERVEE, StrfPoint(TOBStk.GetValue('GED_QTEAFFECTEE')));
      dm_modifchcour(IDM_AFFCDEDISPO, GED_DISPORESERVEE, StrfPoint(TOBStk.GetValue('GED_DISPOAFFECTEE')));
    end;
    dm_modifchcour(IDM_AFFCDEDISPO, GED_MODIF, 'X');
  end;
  // mise à jour du datamem ligne générique
  if TOBGen <> nil then
  begin
    if Etape = afcReservation then
    begin
      dm_modifchcour(IDM_AFFCDELIGNE, GEL_QTERESERVEE, StrfPoint(TOBGen.GetValue('GEL_QTERESERVEE')));
      dm_modifchcour(IDM_AFFCDELIGNE, GEL_MTHTRESERVEE, StrfPoint(TOBGen.GetValue('GEL_MTHTRESERVEE')));
    end else
    begin
      dm_modifchcour(IDM_AFFCDELIGNE, GEL_QTERESERVEE, StrfPoint(TOBGen.GetValue('GEL_QTEAFFECTEE')));
      dm_modifchcour(IDM_AFFCDELIGNE, GEL_MTHTRESERVEE, StrfPoint(TOBGen.GetValue('GEL_MTHTAFFECTEE')));
    end;
    dm_modifchcour(IDM_AFFCDELIGNE, GEL_MODIF, 'X');
  end;
  // mise à jour du datamem pièce
  if TOBCde <> nil then
  begin
    if Etape = afcReservation then
    begin
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_TOTHTRESERVEE, StrfPoint(TOBCde.GetValue('GEP_TOTHTRESERVEE')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_TOTALQTERES, StrfPoint(TOBCde.GetValue('GEP_TOTALQTERES')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_POUQTERES, StrfPoint(TOBCde.GetValue('GEP_POUQTERES')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_POUMTRES, StrfPoint(TOBCde.GetValue('GEP_POUMTRES')));
    end else
    begin
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_TOTHTRESERVEE, StrfPoint(TOBCde.GetValue('GEP_TOTHTAFFECTEE')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_TOTALQTERES, StrfPoint(TOBCde.GetValue('GEP_TOTALQTEAFF')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_POUQTERES, StrfPoint(TOBCde.GetValue('GEP_POUQTEAFF')));
      dm_modifchcour(IDM_AFFCDEPIECE, GEP_POUMTRES, StrfPoint(TOBCde.GetValue('GEP_POUMTAFF')));
    end;
    dm_modifchcour(IDM_AFFCDEPIECE, GEP_MODIF, 'X');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 11/02/2003
Description .. : Réservation ou affectation des quantités de l'ensemble des
Suite ........ : lignes de commande chargées selon les ordres de priorité
Suite ........ : définis.
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.TraitementUnLot: integer;
var
  TOBBe, TOBDi: TOB;
  Ind: integer;
  RefArt: string;
begin
  Result := 0;
  if AvecInitMove then InitMove(TOBBesoin.Detail.Count, '');
  for Ind := 0 to TOBBesoin.Detail.Count - 1 do
  begin
    if AvecInitMove then MoveCur(False);
    TOBBe := TOBBesoin.Detail[Ind];
    if (TiersAutorise(TOBBe, NaturePiece)) and
      (ArticleAutorise(TOBBe, NaturePiece)) then
    begin
      RefArt := TOBBe.GetValue('GEL_ARTICLE');
      if RefArt <> '' then
      begin
        // recherche de la quantité en stock
        TOBDi := FindStock(RefArt, TOBBe);
        // calcul de la quantité à réserver
        if CalculQteReservee(TOBBe, TOBDi) then
        begin
          Inc(Result);
          MajDataMem;
        end;
      end;
    end;
  end;
  if AvecInitMove then FiniMove;
  if Result > 0 then Transactions(MajDesTOB, 1);
  VideLesTOB;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Réservation ou affectation des quantités des lignes de
Suite ........ : commande selon les ordres de priorité définis.
Mots clefs ... :
*****************************************************************}

function T_AFFECTATIONCDE.Traitement: integer;
var
  Nb: integer;
begin
  Result := 0;
  ChargeLesTOB;
  while ChargeBesoinEnTob do
  begin
    Nb := TraitementUnLot;
    Inc(Result, Nb);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Création du composant
Mots clefs ... :
*****************************************************************}

constructor T_AFFECTATIONCDE.Create(CodeAffCde: string; CodeEtape: TTypeEtapeAffCde; NbRow: integer; DataMem, ArtFerme, CliFerme: boolean; NatPiece: string);
begin
  inherited Create;
  CodeAff := CodeAffCde;
  Etape := CodeEtape;
  AvecInitMove := True;
  if NbRow > 0 then MaxNbRow := NbRow else MaxNbRow := 1000;
  IndBesoin := 0;
  UseDataMem := DataMem;
  VerifArtFerme := ArtFerme;
  VerifCliFerme := CliFerme;
  NaturePiece := NatPiece;
  AlloueLesTOB;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Destruction du composant
Mots clefs ... :
*****************************************************************}

destructor T_AFFECTATIONCDE.Destroy;
begin
  LibereLesTOB;
  inherited Destroy;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Affectation des lignes de commandes
Mots clefs ... :
*****************************************************************}

function TraitementAffectationCde(CodeAff: string; Etape: TTypeEtapeAffCde; NbRow: integer; DataMem, ArtFerme, CliFerme: boolean; NatPiece: string): integer;
var
  AffCde: T_AFFECTATIONCDE;
begin
  AffCde := T_AFFECTATIONCDE.Create(CodeAff, Etape, NbRow, DataMem, ArtFerme, CliFerme, NatPiece);
  try
    Result := AffCde.Traitement;
  finally
    AffCde.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.ReSelectCommande;
var
  sSql: string;
  Ind: integer;
  QQ: TQuery;
  ReliquatOk: boolean;
begin
  if TOBPiece = nil then Exit;
  CleDoc := TOB2CleDoc(TOBPiece);
  // vérifie si une commande reliquat existe
  sSql := 'SELECT GP_INDICEG FROM PIECE WHERE GP_VIVANTE="X"'
    + ' AND GP_NATUREPIECEG="' + CleDoc.NaturePiece + '"'
    + ' AND GP_SOUCHE="' + CleDoc.Souche + '"'
    + ' AND GP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
    + ' AND GP_INDICEG=' + IntToStr(CleDoc.Indice);
  ReliquatOk := ExisteSQL(sSql) ;
  // changement d'indice de la pièce d'affectation
  TOBPiece.PutValue('GP_INDICEG', CleDoc.Indice);
  // chargement des lignes
  sSql := 'SELECT GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,'
    + 'GL_NUMLIGNE,GL_NUMORDRE,GL_PIECEPRECEDENTE,GL_TIERS,GL_TIERSLIVRE,'
    + 'GL_ARTICLE,GL_TYPEDIM,GL_CODESDIM,GL_CODEARTICLE,GL_TYPEARTICLE,'
    + 'GL_TENUESTOCK,GL_DATELIVRAISON,GL_TYPELIGNE,GL_QTERESTE,GL_QTESTOCK,GL_MONTANTHT'
    + ' FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False)
    + ' ORDER BY GL_NUMLIGNE';
  QQ := OpenSQl(sSql, True);
  if not QQ.EOF then TOBPiece.LoadDetailDB('LIGNE', '', '', QQ, False);
  Ferme(QQ);
  if TOBPiece.Detail.Count <= 0 then Exit;
  // création des TOBs
  AlloueLesTOB;
  // chargement des quantités réservées
  sSql := 'SELECT GEL_NUMLIGNE,GEL_NUMORDRE,GEL_ARTICLE,GEL_TENUESTOCK,GEL_QTERESERVEE,'
        + ' GEL_QTEAFFECTEE FROM AFFCDELIGNE'
        + ' WHERE GEL_CODEAFF="' + CodeAff + '"'
        + ' AND GEL_NATUREPIECEG="' + CleDoc.NaturePiece + '"'
        + ' AND GEL_SOUCHE="' + CleDoc.Souche + '"'
        + ' AND GEL_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
        + ' AND GEL_INDICEG=' + IntToStr(CleDoc.Indice)
        + ' ORDER BY GEL_NUMLIGNE';
  QQ := OpenSQl(sSql, True);
  if not QQ.EOF then TOBAff_O.LoadDetailDB('AFFCDELIGNE', '', '', QQ, False);
  Ferme(QQ);
  // mise à jour du stock
  ChargeTOBDispo;
  for Ind := 0 to TOBAff_O.Detail.Count - 1 do DecrementStock(Ind);
  // création de la pièce d'affectation
  if ReliquatOk then
  begin
    CreationEntete;
    // constitution des lignes d'affectation
    for Ind := 0 to TOBPiece.Detail.Count - 1 do CreationLigne(Ind);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.CreationEntete;
begin
  if TobPiece = nil then Exit;
  TOBAff.PutValue('GEP_CODEAFF', CodeAff);
  TOBAff.PutValue('GEP_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TOBAff.PutValue('GEP_DATEPIECE', TOBPiece.GetValue('GP_DATEPIECE'));
  TOBAff.PutValue('GEP_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
  TOBAff.PutValue('GEP_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TOBAff.PutValue('GEP_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.CreationLigne(NoLigne: integer);
var
  TOBA, TOBL, TOBQ: TOB;
  QStk, QRes, QAff: double;
begin
  if (TOBPiece = nil) or (TOBAff = nil) then Exit;
  if (NoLigne < 0) or (NoLigne >= TOBPiece.Detail.Count) then Exit;
  TOBL := TOBPiece.Detail[NoLigne];
  QStk := TOBL.GetValue('GL_QTERESTE');
  if QStk <= 0 then Exit ;
  // recherche de l'ancienne ligne d'affectation
  TOBQ := TOBAff_O.FindFirst(['GEL_NUMORDRE'], [TOBL.GetValue('GL_NUMORDRE')], False);
  if TOBQ = nil then Exit;
  // création de la nouvelle ligne d'affectation
  TOBA := TOB.Create('AFFCDELIGNE', TOBAff, -1);
  TOBA.InitValeurs;
  TOBA.PutValue('GEL_CODEAFF', CodeAff);
  TOBA.PutValue('GEL_NATUREPIECEG', TOBL.GetValue('GL_NATUREPIECEG'));
  TOBA.PutValue('GEL_DATEPIECE', TOBL.GetValue('GL_DATEPIECE'));
  TOBA.PutValue('GEL_SOUCHE', TOBL.GetValue('GL_SOUCHE'));
  TOBA.PutValue('GEL_NUMERO', TOBL.GetValue('GL_NUMERO'));
  TOBA.PutValue('GEL_INDICEG', TOBL.GetValue('GL_INDICEG'));
  TOBA.PutValue('GEL_NUMLIGNE', TOBL.GetValue('GL_NUMLIGNE'));
  TOBA.PutValue('GEL_NUMORDRE', TOBL.GetValue('GL_NUMORDRE'));
  TOBA.PutValue('GEL_TIERS', TOBL.GetValue('GL_TIERS'));
  TOBA.PutValue('GEL_TIERSLIVRE', TOBL.GetValue('GL_TIERSLIVRE'));
  TOBA.PutValue('GEL_ARTICLE', TOBL.GetValue('GL_ARTICLE'));
  if TOBL.GetValue('GL_TYPEDIM') = 'GEN' then
    TOBA.PutValue('GEL_CODEARTICLE', TOBL.GetValue('GL_CODESDIM'))
  else
    TOBA.PutValue('GEL_CODEARTICLE', TOBL.GetValue('GL_CODEARTICLE'));
  TOBA.PutValue('GEL_CODESDIM', TOBL.GetValue('GL_CODESDIM'));
  TOBA.PutValue('GEL_TYPEDIM', TOBL.GetValue('GL_TYPEDIM'));
  TOBA.PutValue('GEL_TYPEARTICLE', TOBL.GetValue('GL_TYPEARTICLE'));
  TOBA.PutValue('GEL_TENUESTOCK', TOBL.GetValue('GL_TENUESTOCK'));
  TOBA.PutValue('GEL_DATELIVRAISON', TOBL.GetValue('GL_DATELIVRAISON'));
  TOBA.PutValue('GEL_MODIFIABLE', 'X');
  TOBA.PutValue('GEL_CREERPAR', 'GEN');
  TOBA.PutValue('GEL_PRIORITE', TOBPiece.GetValue('GP_PRIORITE'));
  if TOBA.GetValue('GEL_TYPEDIM') <> 'GEN' then
  begin
    QRes := TOBQ.GetValue('GEL_QTERESERVEE');
    QAff := TOBQ.GetValue('GEL_QTEAFFECTEE');
    // décrémentation de la quantité déjà affectée pour obtenir la nouvelle quantité réservée
    QRes := Arrondi((QRes - QAff), 6);
    if QRes > QStk then QRes := QStk;
    QAff := 0;
    TOBA.PutValue('GEL_QTEALIVRER', QStk);
    TOBA.PutValue('GEL_QTERESERVEE', QRes);
    TOBA.PutValue('GEL_QTEAFFECTEE', QAff);
    TOBA.PutValue('GEL_QTESTOCK', TOBL.GetValue('GL_QTESTOCK'));
    TOBA.PutValue('GEL_MONTANTHT', TOBL.GetValue('GL_MONTANTHT'));
    CalculMontantLigneAff(TOBA, afcReservation);
    CalculMontantPieceAff(TOBAff, TOBA, afcReservation, 0, 0, True);
    // cas des lignes dimensionnées
    if TOBA.GetValue('GEL_TYPEDIM') = 'DIM' then
      MajLigneGeneriqueAff(TOBAff, TOBA, nil, afcReservation, QRes, True);
    // mise à jour du stock
    MajStock(TOBA, QRes, 0, QAff, 0);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.ChargeTOBDispo;
var
  TOBL: TOB;
  TabArt: array of string;
  sSql, RefArt: string;
  QQ: TQuery;
  Ind, Nb, jj: integer;
  Trouve: boolean;
begin
  if MajToutLeStock then Exit;
  // mise en forme de la requête
  Nb := 0;
  sSql := 'SELECT GQ_ARTICLE,GQ_PHYSIQUE,GQ_PREPACLI,GED_QTERESERVEE,GED_QTEAFFECTEE,'
    + 'GED_DISPORESERVEE,GED_DISPOAFFECTEE,GED_QTEPREVI FROM DISPO'
    + ' LEFT JOIN AFFCDEDISPO ON GED_CODEAFF="' + CodeAff + '"'
    + ' AND GED_DEPOT=GQ_DEPOT AND GED_ARTICLE=GQ_ARTICLE'
    + ' WHERE GQ_DEPOT="' + Depot + '" AND GQ_CLOTURE="-"'
    + ' AND GQ_DATECLOTURE="01/01/1900" AND GQ_ARTICLE IN (';
  for Ind := 0 to TOBAff_O.Detail.Count - 1 do
  begin
    TOBL := TOBAff_O.Detail[Ind];
    if TOBL.GetValue('GEL_TENUESTOCK') <> 'X' then Continue;
    RefArt := TOBL.GetValue('GEL_ARTICLE');
    if RefArt = '' then Continue;
    Trouve := False;
    for jj := Low(TabArt) to High(TabArt) do
      if TabArt[jj] = RefArt then
      begin
        Trouve := True;
        Break;
      end;
    if not Trouve then
    begin
      if Nb > 0 then sSql := sSql + ',';
      sSql := sSql + '"' + RefArt + '"';
      SetLength(TabArt, (Nb + 1));
      TabArt[Nb] := RefArt;
      Inc(Nb);
    end;
  end;
  sSql := sSql + ')';
  if Nb <= 0 then Exit;
  // lecture du stock
  QQ := OpenSQl(sSql, True);
  if not QQ.EOF then TOBDispo.LoadDetailDB('DISPO', '', '', QQ, False);
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.DecrementStock(NoLigne: integer);
var
  TOBA: TOB;
  QRes_O, QAff_O: double;
begin
  if MajToutLeStock then Exit;
  if (TOBAff_O = nil) then Exit;
  if (NoLigne < 0) or (NoLigne >= TOBAff_O.Detail.Count) then Exit;
  TOBA := TOBAff_O.Detail[NoLigne];
  // mise à jour du stock
  QRes_O := TOBA.GetValue('GEL_QTERESERVEE');
  QAff_O := TOBA.GetValue('GEL_QTEAFFECTEE');
  MajStock(TOBA, 0, QRes_O, 0, QAff_O);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.MajStock(TOBLigne: TOB; QRes, OldQRes, QAff, OldQAff: double);
var
  TOBS, TOBD: TOB;
  RefArt: string;
  Qte: double;
  //**************************************************************************
  function GetQteStk(NomChamp: string): double;
  var
    Valeur: variant;
  begin
    Valeur := TOBD.GetValue(NomChamp);
    if VarIsNull(Valeur) then
      Result := 0
    else
      Result := VarAsType(Valeur, VarDouble);
  end;
  //**************************************************************************
begin
  if MajToutLeStock then Exit;
  if (TOBLigne = nil) or (TOBStock = nil) then Exit;
  if Depot = '' then Exit;
  // recherche du stock
  RefArt := TOBLigne.GetValue('GEL_ARTICLE');
  if RefArt = '' then Exit;
  if TOBLigne.GetValue('GEL_TENUESTOCK') <> 'X' then Exit;
  TOBS := TOBStock.FindFirst(['GED_ARTICLE'], [RefArt], False);
  if TOBS = nil then
  begin
    TOBD := TOBDispo.FindFirst(['GQ_ARTICLE'], [RefArt], False);
    if TOBD <> nil then
    begin
      TOBS := TOB.Create('AFFCDEDISPO', TOBStock, -1);
      TOBS.InitValeurs;
      TOBS.PutValue('GED_CODEAFF', CodeAff);
      TOBS.PutValue('GED_DEPOT', Depot);
      TOBS.PutValue('GED_ARTICLE', RefArt);
      Qte := GetQteStk('GQ_PHYSIQUE');
      TOBS.PutValue('GED_PHYSIQUE', Qte);
      Qte := GetQteStk('GQ_PHYSIQUE') - GetQteStk('GQ_PREPACLI');
      TOBS.PutValue('GED_QTEPREVI', Qte);
      Qte := GetQteStk('GED_QTERESERVEE');
      TOBS.PutValue('GED_QTERESERVEE', Qte);
      Qte := GetQteStk('GED_QTEAFFECTEE');
      TOBS.PutValue('GED_QTEAFFECTEE', Qte);
      Qte := GetQteStk('GED_DISPORESERVEE') - GetQteStk('GED_QTEPREVI')
        + GetQteStk('GQ_PHYSIQUE') - GetQteStk('GQ_PREPACLI');
      TOBS.PutValue('GED_DISPORESERVEE', Qte);
      Qte := GetQteStk('GED_DISPOAFFECTEE') - GetQteStk('GED_QTEPREVI')
        + GetQteStk('GQ_PHYSIQUE') - GetQteStk('GQ_PREPACLI');
      TOBS.PutValue('GED_DISPOAFFECTEE', Qte);
      TOBS.SetAllModifie(True);
    end;
  end;
  // mise à jour du stock
  if TOBS <> nil then
  begin
    MajStockAFFCDE(TOBS, QRes, afcReservation, OldQRes, True);
    MajStockAFFCDE(TOBS, QAff, afcAffectation, OldQAff, True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Allocation des TOB
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.AlloueLesTOB;
begin
  if not MajToutLeStock then
  begin
    TOBDispo := TOB.Create('', nil, -1);
    TOBStock := TOB.Create('', nil, -1);
  end;
  TOBAff_O := TOB.Create('', nil, -1);
  TOBAff := TOB.Create('AFFCDEPIECE', nil, -1);
  TOBAff.InitValeurs;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Mise à jour des TOB
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.MajDesTOB;
var
  sSql: string;
begin
  // suppression de la commande et des lignes
  sSql := 'DELETE AFFCDEPIECE WHERE GEP_CODEAFF="' + CodeAff + '"'
    + ' AND GEP_NATUREPIECEG="' + CleDoc.NaturePiece + '"'
    + ' AND GEP_SOUCHE="' + CleDoc.Souche + '"'
    + ' AND GEP_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
    + ' AND GEP_INDICEG=' + IntToStr(CleDoc.Indice);
  ExecuteSQL(sSql);
  sSql := 'DELETE AFFCDELIGNE WHERE GEL_CODEAFF="' + CodeAff + '"'
    + ' AND GEL_NATUREPIECEG="' + CleDoc.NaturePiece + '"'
    + ' AND GEL_SOUCHE="' + CleDoc.Souche + '"'
    + ' AND GEL_NUMERO=' + IntToStr(CleDoc.NumeroPiece)
    + ' AND GEL_INDICEG=' + IntToStr(CleDoc.Indice);
  ExecuteSQL(sSql);
  // insertion de la commande reliquat et mise à jour du stock
  if (V_PGI.IoError = oeOk) and (TOBAff <> nil) and (TOBAff.Detail.Count > 0) then
    TOBAff.InsertDB(nil, AvecInitMove);
  if (V_PGI.IoError = oeOk) and (TOBStock <> nil) then
    TOBStock.InsertOrUpdateDB(AvecInitMove);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Libération des TOB
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.LibereLesTOB;
begin
  if TOBPiece <> nil then FreeAndNil(TOBPiece);
  if TOBAff_O <> nil then FreeAndNil(TOBAff_O);
  if TOBAff <> nil then FreeAndNil(TOBAff);
  if TOBStock <> nil then FreeAndNil(TOBStock);
  if TOBDispo <> nil then FreeAndNil(TOBDispo);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Réservation ou affectation des quantités des lignes de
Suite ........ : commande selon les ordres de priorité définis.
Mots clefs ... :
*****************************************************************}

procedure T_RESELECTIONCDE.Traitement;
var
  Ind: integer;
begin
  if TOBLesPieces = nil then Exit;
  // Traitement de chaque commande préparée
  for Ind := 0 to TOBLesPieces.Detail.Count - 1 do
  begin
    TOBPiece := TOB.Create('', nil, -1);
    TOBPiece.Dupliquer(TOBLesPieces.Detail[Ind], True, True, True);
    ReSelectCommande;
    // Mise à jour et libération des TOBs
    Transactions(MajDesTOB, 1);
    // libération des TOBs
    LibereLesTOB;
  end;
  // recalcul du stock de tous les articles
  if MajToutLeStock then UpdateAFFCDEDISPO(CodeAff);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Création du composant
Mots clefs ... :
*****************************************************************}

constructor T_RESELECTIONCDE.Create(CodeAffCde, CodeDepot: string; TOBPieces: TOB; MajToutStock: boolean);
begin
  inherited Create;
  CodeAff := CodeAffCde;
  Depot := CodeDepot;
  TOBLesPieces := TOBPieces;
  AvecInitMove := True;
  MajToutLeStock := MajToutStock;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 22/11/2002
Modifié le ... : 26/11/2002
Description .. : Destruction du composant
Mots clefs ... :
*****************************************************************}

destructor T_RESELECTIONCDE.Destroy;
begin
  LibereLesTOB;
  inherited Destroy;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TraitementReSelectAffCde(CodeAff, Depot: string; TOBPieces: TOB; MajToutStock: boolean);
var
  SelCde: T_RESELECTIONCDE;
begin
  SelCde := T_RESELECTIONCDE.Create(CodeAff, Depot, TOBPieces, MajToutStock);
  try
    SelCde.Traitement;
  finally
    SelCde.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Affichage du détail des pièces traitées
Mots clefs ... :
*****************************************************************}

procedure AfficheDetailPiece(CodeAff: string; CompteRendu: TMemo);
var
  Ind              : integer ;
  TOBPieces, TOBL  : TOB ;
  Stg, sSql, Texte : string ;
  CleDoc           : R_CleDoc;
  QQ               : TQuery;
begin
  if CompteRendu = nil then Exit ;
  CompteRendu.Text := '' ;
  Texte := TraduireMemoire('Préparation des commandes')
         + ' - ' + RechDom('GCAFFCDEENTETE', CodeAff, False)
         + ' - ' + DateTimeToStr(Now) ;
  CompteRendu.lines.Append(Texte) ;
  CompteRendu.lines.Append('') ;
  CompteRendu.Visible := True ;
  CompteRendu.Align := alBottom;
  // Sélection des commandes
  TOBPieces := TOB.Create('Les pieces', nil, -1);
  sSql := 'SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DEVENIRPIECE FROM PIECE '
        + 'WHERE EXISTS (SELECT * FROM AFFCDELIGNE '
        + 'WHERE GEL_CODEAFF="' + CodeAff + '"'
        + ' AND GEL_NATUREPIECEG=GP_NATUREPIECEG AND GEL_SOUCHE=GP_SOUCHE '
        + ' AND GEL_NUMERO=GP_NUMERO AND GEL_INDICEG=GP_INDICEG)';
  QQ := OpenSQl(sSql, True);
  if not QQ.Eof then TOBPieces.LoadDetailDB('PIECE', '', '', QQ, False);
  Ferme(QQ);
  // Affichage des pièces
  for Ind := 0 to TOBPieces.Detail.Count -1 do
  begin
    TOBL := TOBPieces.Detail[Ind] ;
    Texte := RechDom('GCNATUREPIECEG', TOBL.GetValue('GP_NATUREPIECEG'), False)
           + ' n° ' + IntToStr(TOBL.GetValue('GP_NUMERO')) ;
    Stg := TOBL.GetValue('GP_DEVENIRPIECE');
    if Stg <> '' then
    begin
      DecodeRefPiece(Stg, CleDoc) ;
      Texte := Texte + ' => ' + RechDom('GCNATUREPIECEG', CleDoc.NaturePiece, False)
             + ' n° ' + IntToStr(CleDoc.NumeroPiece) ;
    end ;
    CompteRendu.lines.Append (Texte) ;
  end ;
  TOBPieces.Free ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Affichage du compte rendu
Mots clefs ... :
*****************************************************************}

procedure AfficheCompteRendu(TOBListePiece: TOB; NaturePrepaCde: string);
var
  Stg: string;
  PremNum, LastNum, Ind: integer;
begin
  if TOBListePiece.Detail.Count > 0 then
  begin
    PremNum := TOBListePiece.Detail[0].GetValue('NUMERO');
    Ind := TOBListePiece.Detail.Count - 1;
    LastNum := TOBListePiece.Detail[Ind].GetValue('NUMERO');
  end else
  begin
    PremNum := 0;
    LastNum := 0;
  end;
  Stg := TraduireMemoire('Le traitement s''est correctement effectué.')
    + '#13#10' + IntToStr(TOBListePiece.Detail.Count) + ' '
    + TraduireMemoire('pièce(s) générée(s) de nature ')
    + RechDom('GCNATUREPIECEG', NaturePrepaCde, False)
    + TraduireMemoire(' du N° ') + IntToStr(PremNum)
    + TraduireMemoire(' au N° ') + IntToStr(LastNum);
  PGIInfo(Stg, 'Compte-rendu de génération');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Fabrique l'ordre de sélection des lignes de commandes
Mots clefs ... :
*****************************************************************}

function MakeSelect ( CodeAff : string ; TOBAffCde : TOB ) : string;
var
  PouQte, PouMt : integer ;
begin
  if TOBAffCde <> nil then
  begin
    PouQte := TOBAffCde.GetValue('GEA_POUMINIQTECDE');
    PouMt := TOBAffCde.GetValue('GEA_POUMINIMTCDE');
  end else
  begin
    PouQte := 0;
    PouMt := 0;
  end;
  if (PouQte > 0) or (PouMt > 0) then
  begin
    Result := 'SELECT * FROM PIECE WHERE EXISTS (SELECT * FROM AFFCDEPIECE '
            + 'LEFT JOIN AFFCDELIGNE ON GEL_CODEAFF=GEP_CODEAFF '
            + ' AND GEL_NATUREPIECEG=GEP_NATUREPIECEG AND GEL_SOUCHE=GEP_SOUCHE '
            + ' AND GEL_NUMERO=GEP_NUMERO AND GEL_INDICEG=GEP_INDICEG '
            + 'WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_QTEAFFECTEE > 0 '
            + ' AND GEL_NATUREPIECEG=GP_NATUREPIECEG AND GEL_SOUCHE=GP_SOUCHE '
            + ' AND GEL_NUMERO=GP_NUMERO AND GEL_INDICEG=GP_INDICEG';
    if PouQte > 0 then
      Result := Result + ' AND GEP_POUQTEAFF > '+ IntToStr(PouQte);
    if PouMt > 0 then
      Result := Result + ' AND GEP_POUMTAFF > '+ IntToStr(PouMt);
    Result := Result + ')';
  end else
  begin
    Result := 'SELECT * FROM PIECE WHERE EXISTS (SELECT * FROM AFFCDELIGNE '
            + 'WHERE GEL_CODEAFF="' + CodeAff + '" AND GEL_QTEAFFECTEE > 0 '
            + ' AND GEL_NATUREPIECEG=GP_NATUREPIECEG AND GEL_SOUCHE=GP_SOUCHE '
            + ' AND GEL_NUMERO=GP_NUMERO AND GEL_INDICEG=GP_INDICEG)';
  end ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 06/12/2002
Modifié le ... : 06/12/2002
Description .. : Préparation des lignes de commandes
Mots clefs ... :
*****************************************************************}

function TraitementPreparationCde(CodeAff: string; Eclate, DeGroupeRemise, AvecComment, AvecCompteRendu, ContinuOnError, MajToutStock: boolean;
  TraitFinTravaux: integer ; CompteRendu : TMemo ): integer;
var
  TOBPieces, TOBListePiece, TOBAffCde: TOB;
  sSql, Depot, NaturePrepaCde: string;
  QQ: TQuery;
begin
  Result := 0;
  TOBListePiece := nil;
  // Recherche du dépôt associé au code affectation
  TOBAffCde := TOB.Create('AFFCDEENTETE', nil, -1) ;
  TOBAffCde.SelectDB('"'+ CodeAff +'"', nil);
  Depot := TOBAffCde.GetValue('GEA_DEPOT');
  NaturePrepaCde := TOBAffCde.GetValue('GEA_NATPIECEGRP');
  // Sélection des commandes avec des lignes affectées
  TOBPieces := TOB.Create('Les pieces', nil, -1);
  sSql := MakeSelect(CodeAff, TOBAffCde);
  QQ := OpenSQl(sSql, True);
  if not QQ.Eof then TOBPieces.LoadDetailDB('PIECE', '', '', QQ, False);
  Ferme(QQ);
  // Génération des pièces
  try
    if TOBPieces.Detail.Count > 0 then
    begin
      AddLesSupLigne(TOBPieces.Detail[0], True);
      TheTob := nil;
      RegroupeLesPieces(TOBPieces, NaturePrepaCde, Eclate, DeGroupeRemise, AvecComment,
        0, 0, False, True, ContinuOnError, False, False, CodeAff);
      TOBListePiece := TheTob;
      TheTob := nil;
      // Affichage du détail des pièces traitées
      if AvecCompteRendu then AfficheDetailPiece(CodeAff, CompteRendu);
      // Traitement final
      case TraitFinTravaux of
        2: // fin de l'affectation
          AvecCompteRendu := False;
        1: // sélection des reliquats
          TraitementReSelectAffCde(CodeAff, Depot, TOBPieces, MajToutStock);
      end;
      // Compte rendu
      if TOBListePiece <> nil then
      begin
        Result := TOBListePiece.Detail.Count;
        if AvecCompteRendu then AfficheCompteRendu(TOBListePiece, NaturePrepaCde);
        FreeAndNil(TOBListePiece);
      end;
    end;
  finally
    TOBPieces.Free;
  end;
end;

end.

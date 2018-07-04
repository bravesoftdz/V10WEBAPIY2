{***********UNITE*************************************************
Auteur  ...... : Julien DITTMAR
Créé le ...... : 19/02/2001
Modifié le ... : 18/08/2003
Description .. : Unité de calcul d'une nouvelle affectation de proposition
Mots clefs ... : TRANSFERT;PROPOSITION;AFFECTATION;
*****************************************************************}
unit PropoAffectTrf;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFNDEF EAGLCLIENT}
  DBTables, HDB,
  {$ENDIF}
  Hctrls, uTob, StdCtrls, HEnt1, Grids, uiutil, hpanel,
  HMsgBox, Spin, HStatus, UtilArticle, SaisUtil;
  //, FOWaitFor;

function TraitementAffectationTrf(FF: TForm; TobAffectationTrf, TobMethode, TobListeEtab: TOB; TypePropoTrf: string): TOB;
procedure ChargeTob_DispoDepot(FF: TForm; TobAffectationTrf, TobMethode: TOB; CodDepot: string);
function ChargeDepotAffectation(TobArticle, TobDuplicSelArt, TobListeEtab: TOB; CodDepot: string): boolean;
function ChargeStockLigneEnTOB(ListeArticles : TStringList;
         StPeriode, NomChampSUM, ListeNaturePiece, ListeEtab : string;
         BVivante : Boolean) : TOB;
procedure ChargeDepotQteARepartir(FF: TForm; TobAffectationTrf: TOB);
function ChargeArticle_PropositionTable(TobArticle, TobListeEtab: TOB): boolean;
procedure TraitementMethodesAffectation(TobAffectationTrf, TobMethode: TOB);
function TraitementUneMethode(TobAffectationTrf: TOB; CodeMethode: string): boolean;
procedure AffectationEtablissement(TobAffectationTrf, TobEtabMeth: TOB; TypeMethode, DebPeriodeVte, FinPeriodeVte: string; QteSaisie: double);
procedure Charge_TobStockEtab(TobStockEtab: TOB; ListeArticles : TStringList);
procedure Init_TobEtabAff(TobEtabAff: TOB; CodeArt, IdentArt, StatutArt, CodeEtablis: string;
  NGTL_CODEPTRF, NGTL_ARTICLE, NGTL_CODEARTICLE, NGTL_STATUTART, NGTL_SURPLUS,
  NGTL_DEPOTDEST, NGTL_STOCKINITIAL, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_QTEMIN,
  NGTL_QTEMAX, NGTL_QTEVTE, NGTL_MANQUANT, NGTL_PROPOSITION, NGTL_STOCKLU,
  NGTL_STOCKALF, NGTL_GENERER, NGTL_GENECF : integer);
procedure ChargeStock_TobEtabAff(TobDepotAff, TobEtabAff, TobStockEtab: TOB;
  IdentArt, CodeEtablis: string;
  NGTL_STOCKINITIAL, NGTL_SURPLUS, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_STOCKLU,
  NGTL_RESERVECLI, NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_PROPOSITION,
  NGQ_PHYSIQUE, NGQ_STOCKMIN, NGQ_STOCKMAX, NGQ_RESERVECLI, NGQ_RESERVEFOU,
  NGQ_PREPACLI, N_ARTDEPOT : integer );
procedure ChargeVente_TobEtabAff(TobEtabAff, TobVenteEtab: TOB; IdentArt, CodeEtablis, DebPeriodeVte, FinPeriodeVte: string; NVente_ARTETAB : integer);
procedure ChargeTransit_TobEtabAff(TobEtabAff, TobTransitEtab: TOB; IdentArt, CodeEtablis : string; NTransit_ARTETAB : integer);
procedure ChargeQteMax_TobEtabAff(TobDepotAff, TobEtabAff: TOB; TypeMethode: string;
  QteSaisie, CoefRepart: Double; NGTL_STOCKMAXI, NGTL_QTEMAX, NGTL_STOCKINITIAL,
  NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_RESERVECLI, NGTL_PROPOSITION,
  NGTL_SURPLUS, NGTL_MANQUANT : integer);
procedure ChargeProposition_TobEtabAff(TobDepotAff, TobMethEtabAff: TOB;
  TypeMethode: string; NGTL_PROPOSITION, NGTL_MANQUANT, NGTL_SURPLUS : integer);
procedure Init_TobCalculAff(TobCalculAff: TOB; CodeEtablis: string);
function ChargeBesoin_TobCalculAff(TobEtabAff, TobCalculAff: TOB;
  TypeMethode: string; QteSaisie, CoefRepart: Double;
  NGTL_STOCKINITIAL, NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_RESERVECLI,
  NGTL_STOCKALF, NGTL_PROPOSITION, NGTL_STOCKMINI, NGTL_QTEVTE, NGTL_QTEMIN,
  NGTL_QTEMAX : integer): Double;
procedure ChargeAffectation_TobCalculAff(TobMethEtabAff: TOB; TypeMethode: string; QteARepartir, TotBesoin: Double);
procedure Alimentation_TobEcrPTrf(TobAffectationTrf, TobEcrPTrf: TOB);
function ArrondirQte(MethodeArr: string; Qte: double): double;
procedure AlimenteTobMethode(CodeMethode, TypeMethode, ArrondiMeth, UtilCoef: string; QteFixe, QteMin, QteMax: double; DebPerVte, FinPerVte: string;
  TobMethode: TOB);
procedure AjouteLesLignesGEN(TOBLignesDIM: TOB);
function CreerListeArticle(Const TOBArts : TOB; Const AvecCodeArticle : Boolean=False) : TStringList;
function ConstruireInArticle(Const ListeArticles : TStringList; Var Indice : Integer) : string;
function ChargeStockTransit(ListeArticles : TStringList; BSiege : Boolean): TOB;
function ChargeStockVente(ListeArticles : TStringList; DebPeriodeVte, FinPeriodeVte : String; BDateIntegr : Boolean): TOB;

implementation

var RechStockEtab: boolean; // Est-ce que le stock des établissements à approvisionner a été chargé en TOB.
  AffectMaxi: boolean; // Est-ce qu'une méthode, contrôlant la quantité maxi à approvisionner , a été affectée.
  UtilStockEtab: boolean; // Prise en compte du stock de la boutique
  UtilCoefEtab: boolean; // Application du coefficient par boutique
  RepartRatio: boolean; // Répartion par ratio
  VerifStkMaxi: boolean; // Vérification du stock maxi
  ArretRupture: boolean; // Ne pas exécuter les méthodes après rupture de stock au dépôt
  PropEncours: boolean; // Proposition en cours de saisie
  ReaffectArticle: boolean; // Réaffectation des articles déjà traités
  ReaffectTsBout: boolean; // Réaffectation sur toutes boutiques ou uniquement celles à approvisionner
  LimiteEnreg: boolean; // Permet de supprimer les enreg qui n'ont pas d'affectation
  CodeDepot: string; // Code de l'établissement émetteur
  CodePropTrf: string; // Code de la proposition de transfert
  TypePropTrf: string; // Type de proposition de transfert : 'M'anuel ou 'A'utomatique
  MethodeArrondi: string; // Méthode d'arrondi : Inférieur, au plus Près, Supérieur
  ListeEtabMeth: string; // Liste des établissements d'une méthode d'affectation
  Tob_DispoDepot: TOB; // Tob sur laquelle est chargée le Dispo du dépôt
  Tob_PropoTable: TOB; // Tob sur laquelle est chargée la proposition en cours
  SiegeGereStockNet : boolean; // Gestion du stock net pour le dépôt
  SiegeGereStockTransit : boolean; // Gestion du stock transit pour le dépôt
  BTQGereStockNet : boolean; // Gestion du stock net pour les boutiques
  BTQGereStockTransit : boolean; // Gestion du stock transit pour les boutiques
  NG_DEPOTDEST, NG_BESOIN, NG_RATIO, NG_AFFECTATION, NG_QTESTK, NG_QTEMAX : integer;

const
  NbArtParRequete : integer = 1000;

procedure FindFirstValeur (var TobFound : Tob; var TobMere : Tob; NomChamp: array of string; ValeurChamp: array of Variant; MultiNiveau : boolean;
                           NumeroChamp : array of integer; TobSorted : boolean) ;
var max, jj, kk, kmin, kmax : integer ;
    okok : boolean;
  Procedure FindDicho (stSearch : String);
  var Min, Max, k, u : integer;
  Begin
    Min := 0;
    Max := TobMere.Detail.Count-1;
    while True do
    begin
      if Min > Max then
      begin
        TobFound := nil;
        Break;
      end
      else
      begin
        k := Round( (Min + Max) / 2 );
        u := AnsiCompareStr( stSearch, TobMere.Detail[k].GetValeur(NumeroChamp[0]));
        if u = 0 then
          begin
          TobFound := TobMere.Detail[k];
          Break;
          end
        else if u < 0 then
          Max := k - 1
        else
          Min := k + 1
      end;
    end;
  end;
begin
  TobFound := nil;
  if TobMere.Detail.Count < 1 then exit;
  if High(NomChamp) <> High(ValeurChamp) then exit;
  if High(NomChamp) <> High(NumeroChamp) then exit;
  // Si plus d'un champ le tri par dichotomie n'est pas géré !
  if Low(NomChamp) <> High(NomChamp) then TobSorted := False;
  // MultiNiveau non géré. On force appel à FindFirst.
  // Normalement, aucun appel avec MultiNiveau à True dans la recherche de tarif !
  If MultiNiveau then NumeroChamp[0] := -1;

  // si numchamp pas trouvé, comme avant...
  if (NumeroChamp[0] = -1) then
    TobFound := TobMere.FindFirst(NomChamp, ValeurChamp, MultiNiveau)
  // sinon on boucle sur le même
  else
  begin
    if TobMere <> nil then
    begin
      if TobSorted then
      begin
        // Recherche dans une Tob triée
        FindDicho (ValeurChamp[0]) ;
      end
      else
      begin
        // Recherche dans une Tob NON triée
        max := TobMere.detail.count-1;
        kmin := Low(NumeroChamp);
        kmax := High(NumeroChamp);
        okok := False; // Recherche un seul élément, pas de next donc pas de gestion de TobMere.Detail[jj].FoundByFind
        for jj := 0 to max do
        begin
          for kk := kmin to kmax do
          begin
            okok := TobMere.Detail[jj].GetValeur(NumeroChamp[kk]) = ValeurChamp[kk];
            if not okok then Break;
          end;
          if okok then // Trouvé
          begin
            TobFound := TobMere.Detail[jj];
            Break;
          end;
        end;
      end;
    end;
  end;
end;

function TraitementAffectationTrf(FF: TForm; TobAffectationTrf, TobMethode, TobListeEtab: TOB; TypePropoTrf: string): TOB;
var i, j: integer;
  NbArt_Tot: integer;
  ReaffectItem: integer;
  TobArticle: TOB;
  TobDuplicSelArt: TOB;
  TobEcrPTrf: TOB;
  QQ: TQuery;
begin
  // ----------------------------------------------------------------------------
  // TobAffectationTrf est la Tob dans laquelle sera stockée toute la proposition
  // d'affectation de transferts, lors de la phase de génération de celle-ci.
  // Cette Tob contient sur son premier niveau, la liste des articles sélectionnés
  // depuis le MUL Article : TobArticle.
  // ----------------------------------------------------------------------------
  result := nil;
  TypePropTrf := TypePropoTrf;

  CodePropTrf := THEdit(FF.FindComponent('CODPROPAFF')).Text;

  ReaffectItem := THRadioGroup(FF.FindComponent('REAFFECT')).ItemIndex;
  ReaffectTsBout := False;
  if ReaffectItem = 0 then ReaffectArticle := False
  else
  begin
    ReaffectArticle := True;
    if ReaffectItem = 2 then ReaffectTsBout := True;
  end;

  ArretRupture := False;
  if TypePropTrf = 'A' then ArretRupture := TCheckBox(FF.FindComponent('ARRETRUPTURE')).Checked;

  // Lecture de la Table Entête proposition de transfert : PROPTRANSFENT
  CodeDepot := '';
  QQ := OpenSQL('Select GTE_DEPOTEMET, GTE_ENCOURSPTRF, GTE_STKTRANSIT, ' +
  'GTE_STKNET, GTE_STKTRANSITBTQ, GTE_STKNETBTQ from PROPTRANSFENT ' +
  'Where GTE_CODEPTRF="' + CodePropTrf + '"', TRUE);
  if QQ.EOF then
  begin
    Ferme(QQ);
    exit;
  end;
  CodeDepot := QQ.Findfield('GTE_DEPOTEMET').AsString;
  PropEncours := Boolean(QQ.Findfield('GTE_ENCOURSPTRF').AsString = 'X');
  SiegeGereStockTransit := (QQ.Findfield('GTE_STKTRANSIT').AsString = 'X');
  BTQGereStockTransit := (QQ.Findfield('GTE_STKTRANSITBTQ').AsString = 'X');
  SiegeGereStockNet := (QQ.Findfield('GTE_STKNET').AsString = 'X');
  BTQGereStockNet := (QQ.Findfield('GTE_STKNETBTQ').AsString = 'X');

  Ferme(QQ);

  //-----------------------------------------------------------------------------
  // Duplication de TobAffectationTrf, si des articles dimensionnés ont été sélectionnés,
  // afin de pouvoir contrôler par la suite qu'il n'y ait pas redondance entre la
  // sélection d'articles génériques et la sélection d'articles dimensionnés.
  //-----------------------------------------------------------------------------
  TobDuplicSelArt := nil;
  if TobAffectationTrf.FindFirst(['GA_STATUTART'], ['DIM'], False) <> nil then
  begin
    TobDuplicSelArt := TOB.Create('Duplic-SelArt', nil, -1);
    TobDuplicSelArt.Dupliquer(TobAffectationTrf, True, True, True);
  end;

  // -----------------------------------------------------------------------------
  // Fabrication d'une requête SQL, permettant de charger en TOB, pour l'ensemble
  // des articles sélectionnés :
  // - le stock du dépôt emetteur
  // - la proposition en cours déjà mémorisée dans la base.
  // -----------------------------------------------------------------------------
  Tob_DispoDepot := TOB.CREATE('Dispo-Depot', nil, -1);
  Tob_PropoTable := TOB.CREATE('Propo-Table', nil, -1);
  ChargeTob_DispoDepot(FF, TobAffectationTrf, TobMethode, CodeDepot);

  //------------------------------------------------------------------------------
  // Sur chaque article sélectionné dans TobAffectationTrf, création d'une Tob fille
  // listant tous les articles dimensionnés qui possèdent dans l'établissement "Dépôt"
  // de la marchandise à dispatcher : TobDepotAff.
  // Si un article sélectionné ne possède aucune fille (pas de stock à dispatcher),
  // celui-ci est supprimé du premier niveau : TobArticle.
  // -----------------------------------------------------------------------------
  j := 0;
  NbArt_Tot := TobAffectationTrf.Detail.Count;
  InitMove(NbArt_Tot, 'Recherche du stock du dépôt émetteur ...');
  for i := 0 to NbArt_Tot - 1 do
  begin
    TobArticle := TobAffectationTrf.Detail[j];
    if ChargeDepotAffectation(TobArticle, TobDuplicSelArt, TobListeEtab, CodeDepot) = True
      then inc(j)
    else TobArticle.Free;
    MoveCur(False);
  end;
  FiniMove;

  Tob_DispoDepot.Free;
  Tob_DispoDepot := nil;
  Tob_PropoTable.Free;
  Tob_PropoTable := nil;
  if TobDuplicSelArt <> nil then TobDuplicSelArt.Free;

  // -------------------------------------------------------------------------
  // Calcul de la quantité à répartir pour chaque article dimensionné du dépôt,
  // tenant compte du stock disponible et de la quantité à laisser au dépôt.
  // -------------------------------------------------------------------------
  ChargeDepotQteARepartir(FF, TobAffectationTrf);

  // ---------------------------------------------------------------
  // Traitement des méthodes d'affectation les unes après les autres
  // ---------------------------------------------------------------
  TraitementMethodesAffectation(TobAffectationTrf, TobMethode);

  // ---------------------------------------------------------------------------
  // Alimentation de la Tob finale au format de la table stockant la proposition
  // de transfert. Ajout du cumul sur les articles génériques.
  // ---------------------------------------------------------------------------
  TobEcrPTrf := TOB.Create('Prop_Trf', nil, -1);
  Alimentation_TobEcrPTrf(TobAffectationTrf, TobEcrPTrf);
  Result := TobEcrPTrf;
end;

function LimiteAuxVentes(TobMethode: TOB): string;
var StDatePiece, StDateIntegr, StSQL : string;
    NbFille: integer;
begin
  StDatePiece := '';
  StDateIntegr := '';
  for NbFille := 0 to TobMethode.Detail.count - 1 do
  begin
    With TobMethode.Detail[NbFille] do
    begin
      if GetValue('TYPE') = 'PVT' then
      begin
        if StDatePiece <> '' then StDatePiece := StDatePiece + ' or ';
        StDatePiece := StDatePiece + '(GL_DATEPIECE>="' + GetValue('DATEDEBUS')
        + '" and GL_DATEPIECE<="' + GetValue('DATEFINUS') + '")';
      end
      else if GetValue('TYPE') = 'PVV' then
      begin
        if StDateIntegr <> '' then StDateIntegr := StDateIntegr + ' or ';
        StDateIntegr := StDateIntegr + '(GP_DATEINTEGR>="' + GetValue('DATEDEBUS')
        + '" and GP_DATEINTEGR<="' + GetValue('DATEFINUS') + '")';
      end;
    end;
  end;
  
  if StDatePiece <> '' then
    StSQL := ' AND GA_ARTICLE IN ' +
    '(Select GL_ARTICLE From LIGNE Where GL_NATUREPIECEG IN ("FFO","FAC") and ' +
    '(' + StDatePiece + '))';

  if StDateIntegr <> '' then
    StSQL := StSQL + ' AND GA_ARTICLE IN ' +
    '(Select GL_ARTICLE From PIECE LEFT JOIN LIGNE ' +
    'ON GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE ' +
    'AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG ' +
    'AND (' + StDateIntegr + ') ' +
    'Where GL_NATUREPIECEG IN ("FFO","FAC"))';

  Result := StSQL;
end;

// -----------------------------------------------------------------------------
// Fabrication d'une requête SQL, permettant de charger en TOB, pour l'ensemble
// des articles sélectionnés :
// - le stock du dépôt emetteur
// - la proposition en cours déjà mémorisée dans la base.
// -----------------------------------------------------------------------------

procedure ChargeTob_DispoDepot(FF: TForm; TobAffectationTrf, TobMethode: TOB; CodDepot: string);
var y : integer;
  StSelect, StWhere, StSQL, StLimiteArt, StWhereArticles : string;
  ListeArticles : TStringList;
  QDispo, QPropo: TQuery;
begin
  ListeArticles := CreerListeArticle(TobAffectationTrf, True);

  StSelect := 'Select GA_ARTICLE,GA_CODEARTICLE,GA_STATUTART,' +
  'GQ_PHYSIQUE,GQ_RESERVEFOU,GQ_RESERVECLI,GQ_STOCKMIN,GQ_STOCKMAX,GQ_PREPACLI ';
  StWhere := '';

  if TypePropTrf = 'M' then LimiteEnreg := False
  else LimiteEnreg := TCheckBox(FF.FindComponent('GTE_LIMITEENREG')).Checked;

  if ListeArticles.Count > 0 then
  begin
    if TypePropTrf = 'M' then StLimiteArt := 'AUC'
    else StLimiteArt := THRadioGroup(FF.FindComponent('LIMITEART')).Value;
    if StLimiteArt = 'AUC' then
      // Prise en compte de tous les articles mouvementés ou non dans le dépôt
      StSelect := StSelect + 'From ARTICLE Left join DISPO '+
      'ON GA_ARTICLE=GQ_ARTICLE and GQ_DEPOT="' + CodDepot + '" and GQ_CLOTURE="-" '
    else if StLimiteArt = 'STK' then
    begin
      // Prise en compte uniquement des articles mouvementés
      StSelect := StSelect + 'From ARTICLE Left join DISPO on GA_ARTICLE=GQ_ARTICLE ';
      StWhere := ' and GQ_DEPOT="' + CodDepot + '" and GQ_CLOTURE="-"';
    end
    else if StLimiteArt = 'STD' then
    begin
      // Prise en compte uniquement des articles ayant du stock au dépôt
      StSelect := StSelect + 'From ARTICLE Left join DISPO on GA_ARTICLE=GQ_ARTICLE ';
      StWhere := ' and GQ_DEPOT="' + CodDepot + '" and GQ_CLOTURE="-" and GQ_PHYSIQUE>0';
    end
    else if StLimiteArt = 'VTE' then
    begin
      // Prise en compte uniquement des articles ayant du stock au dépôt
      StSelect := StSelect + 'From ARTICLE Left join DISPO ' +
      'ON GA_ARTICLE=GQ_ARTICLE and GQ_DEPOT="' + CodDepot + '" and GQ_CLOTURE="-" ';
      StWhere := LimiteAuxVentes(TobMethode);
    end;

    y := 0;
    While y < ListeArticles.Count do
    begin
      StWhereArticles := ConstruireInArticle(ListeArticles, y);
      if StWhereArticles <> '' then
      begin
        StSQL := StSelect + 'Where GA_STATUTART<>"GEN" AND GA_CODEARTICLE IN (' + StWhereArticles + ')' + StWhere;
        QDispo := OpenSQL(StSQL, True);
        if not QDispo.EOF then Tob_DispoDepot.LoadDetailDB('', '', '', QDispo, True);
        Ferme(QDispo);

        if PropEncours = True then
        begin
          StSQL := 'Select * from PROPTRANSFLIG Where GTL_CODEPTRF="' + CodePropTrf + '" and ' +
            'GTL_CODEARTICLE IN (' + StWhereArticles + ') and GTL_STATUTART<>"GEN"';
          QPropo := OpenSQL(StSQL, True);
          if not QPropo.EOF then Tob_PropoTable.LoadDetailDB('PROPTRANSFLIG', '', '', QPropo, True);
          Ferme(QPropo);
        end;
      end;
    end;
  end;

  ListeArticles.Free;
end;

//------------------------------------------------------------------------------
// Sur chaque article sélectionné dans TobAffectationTrf, création d'une Tob fille
// listant tous les articles dimensionnés qui possèdent dans l'établissement
// émetteur "Dépôt" de la marchandise à dispatcher.
// -----------------------------------------------------------------------------

function ChargeDepotAffectation(TobArticle, TobDuplicSelArt, TobListeEtab: TOB; CodDepot: string): boolean;
var CodeArt: string;
  IdentArt: string;
  StatutArt: string;
  ReaffectOk: boolean;
  TobDepotAff: TOB;
  TobDispo: TOB;
begin
  result := True;
  CodeArt := TobArticle.GetValue('GA_CODEARTICLE');
  IdentArt := TobArticle.GetValue('GA_ARTICLE');
  StatutArt := TobArticle.GetValue('GA_STATUTART');

  if StatutArt = 'DIM' then
  begin
    // Recherche si l'article générique a été sélectionné.
    // Si Oui, la sélection sur l'article dimensionné est annulée, puisque tous les
    // articles dimensionnés de ce générique seront traités par défaut.
    if TobDuplicSelArt.FindFirst(['GA_CODEARTICLE', 'GA_STATUTART'], [CodeArt, 'GEN'], False) <> nil then
    begin
      result := False;
      exit;
    end;
  end;

  // Si une proposition d'affectation a déjà été validée (En cours), chargement dans
  // TobAffectationTrf de la proposition initiale de l'article si celle-ci existe.
  ReaffectOk := False;
  if PropEncours = True then
  begin
    ReaffectOk := ChargeArticle_PropositionTable(TobArticle, TobListeEtab);
  end;

  // Recherche du DISPO de l'article pour l'établissement émetteur
  if StatutArt = 'GEN' then
    TobDispo := Tob_DispoDepot.FindFirst(['GA_CODEARTICLE'], [CodeArt], False)
  else TobDispo := Tob_DispoDepot.FindFirst(['GA_ARTICLE'], [IdentArt], False);

  if TobDispo = nil then
  begin
    // Pas de ligne DISPO pour cet article sélectionné
    if ReaffectOk = True then result := True else result := False;
    exit;
  end;

  while (TobDispo <> nil) do
  begin
    IdentArt := TobDispo.GetValue('GA_ARTICLE');
    if ReaffectOk = False then TobDepotAff := nil
    else TobDepotAff := TobArticle.FindFirst(['GTL_ARTICLE'], [IdentArt], False);
    if TobDepotAff = nil then
    begin
      TobDepotAff := TOB.Create('PROPTRANSFLIG', TobArticle, -1);
      With TobDepotAff do
      begin
        PutValue('GTL_CODEPTRF', CodePropTrf);
        PutValue('GTL_ARTICLE', IdentArt);
        PutValue('GTL_CODEARTICLE', CodeArt);
        PutValue('GTL_DEPOTDEST', '...');
        PutValue('GTL_QTEMIN', 0.0);
        PutValue('GTL_QTEMAX', 0.0);
        PutValue('GTL_QTEVTE', 0.0);
        PutValue('GTL_MANQUANT', 0.0);
        PutValue('GTL_SURPLUS', 0.0);
        PutValue('GTL_PROPOSITION', 0.0);
        PutValue('GTL_STOCKLU', 'X');
        PutValue('GTL_STOCKALF', 0.0);
        PutValue('GTL_RESERVECLI', 0.0);
        PutValue('GTL_RESERVEFOU', 0.0);
        PutValue('GTL_PREPACLI', 0.0);
        PutValue('GTL_GENERER', '-');
        PutValue('GTL_GENECF', '-');
      end;
    end;

    With TobDepotAff do
    begin
      PutValue('GTL_STATUTART', TobDispo.GetValue('GA_STATUTART'));
      PutValue('GTL_STOCKINITIAL', TobDispo.GetValue('GQ_PHYSIQUE'));
      PutValue('GTL_STOCKMINI', TobDispo.GetValue('GQ_STOCKMIN'));
      PutValue('GTL_STOCKMAXI', TobDispo.GetValue('GQ_STOCKMAX'));
      PutValue('GTL_RESERVECLI', TobDispo.GetValue('GQ_RESERVECLI'));
      PutValue('GTL_RESERVEFOU', TobDispo.GetValue('GQ_RESERVEFOU'));
      PutValue('GTL_PREPACLI', TobDispo.GetValue('GQ_PREPACLI'));
    end;

    if StatutArt = 'GEN' then
      TobDispo := Tob_DispoDepot.FindNext(['GA_CODEARTICLE'], [CodeArt], False)
    else TobDispo := nil;
  end;
end;

// --------------------------------------------------------------------------
// Chargement de la proposition initiale de l'article dans TobAffectationTrf.
// --------------------------------------------------------------------------

function ChargeArticle_PropositionTable(TobArticle, TobListeEtab: TOB): boolean;
var CodeArt: string;
  CodeEtablis: string;
  ReaffOk: boolean;
  i, j: integer;
  TotManquant, TotProposition: Double;
  TobDepotAff: TOB;
  TobEtabAff: TOB;
  TobPropArticle: TOB;
  TobProp: TOB;
begin
  result := False;
  CodeArt := TobArticle.GetValue('GA_CODEARTICLE');

  TobProp := Tob_PropoTable.FindFirst(['GTL_CODEARTICLE'], [CodeArt], False);
  // Pas de proposition en cours pour cet article
  if TobProp = nil then Exit;

  // On isole la proposition de l'article
  TobPropArticle := TOB.Create('', nil, -1);
  while (TobProp <> nil) do
  begin
    TobProp.ChangeParent(TobPropArticle, -1);
    TobProp := Tob_PropoTable.FindFirst(['GTL_CODEARTICLE'], [CodeArt], False);
  end;

  // Reprise de la proposition d'affectation de tous articles dimensionnés sur le
  // dépôt émetteur.
  TobProp := TobPropArticle.FindFirst(['GTL_DEPOTDEST'], ['...'], False);
  while TobProp <> nil do
  begin
    TobProp.ChangeParent(TobArticle, -1);
    TobProp := TobPropArticle.FindFirst(['GTL_DEPOTDEST'], ['...'], False);
  end;

  // Pour chaque article dimensionné, reprise de l'affectation sur les différents
  // établissements destinataires
  for i := 0 to TobArticle.Detail.Count - 1 do
  begin
    TobDepotAff := TobArticle.Detail[i];
    TobProp := TobPropArticle.FindFirst(['GTL_ARTICLE'], [TobDepotAff.GetValue('GTL_ARTICLE')], False);
    while TobProp <> nil do
    begin
      TobProp.ChangeParent(TobDepotAff, -1);
      TobProp := TobPropArticle.FindFirst(['GTL_ARTICLE'], [TobDepotAff.GetValue('GTL_ARTICLE')], False);
    end;
  end;

  // Remise à zéro de la proposition si demande de réaffectation
  if ReaffectArticle = True then
  begin
    for i := 0 to TobArticle.Detail.Count - 1 do
    begin
      TobDepotAff := TobArticle.Detail[i];
      TotManquant := 0.0;
      TotProposition := 0.0;
      for j := 0 to TobDepotAff.Detail.Count - 1 do
      begin
        TobEtabAff := TobDepotAff.Detail[j];
        ReaffOk := True;
        if ReaffectTsBout = False then
        begin
          CodeEtablis := TobEtabAff.GetValue('GTL_DEPOTDEST');
          if TobListeEtab.FindFirst(['ETABLISSEMENT'], [CodeEtablis], False) = nil then
            ReaffOk := False;
        end;
        if ReaffOk = True then
        begin
          TobEtabAff.PutValue('GTL_QTEMIN', 0.0);
          TobEtabAff.PutValue('GTL_QTEMAX', 0.0);
          TobEtabAff.PutValue('GTL_QTEVTE', 0.0);
          TobEtabAff.PutValue('GTL_DEBPERVTE', StrToDate('01/01/1900'));
          TobEtabAff.PutValue('GTL_FINPERVTE', StrToDate('01/01/1900'));
          TobEtabAff.PutValue('GTL_MANQUANT', 0.0);
          TobEtabAff.PutValue('GTL_PROPOSITION', 0.0);
          TobEtabAff.PutValue('GTL_STOCKALF', 0.0);
          TobEtabAff.PutValue('GTL_RESERVECLI', 0.0);
          TobEtabAff.PutValue('GTL_RESERVEFOU', 0.0);
          TobEtabAff.PutValue('GTL_PREPACLI', 0.0);
          TobEtabAff.PutValue('GTL_STOCKLU', '-');
        end else
        begin
          TotManquant := ToTManquant + TobEtabAff.GetValue('GTL_MANQUANT');
          TotProposition := TotProposition + TobEtabAff.GetValue('GTL_PROPOSITION');
        end;
      end;
      TobDepotAff.PutValue('GTL_SURPLUS', 0.0);
      TobDepotAff.PutValue('GTL_MANQUANT', TotManquant);
      if TotProposition = 0.0 then
        TobDepotAff.PutValue('GTL_PROPOSITION', 0.0)
      else TobDepotAff.PutValue('GTL_PROPOSITION', -TotProposition);
    end;
  end;
  TobPropArticle.Free;
  result := True;
end;

function ConstruireInArticle(Const ListeArticles : TStringList; Var Indice : Integer) : string;
var StIn : string;
    NbArt, i : integer;
begin
  StIn := '';
  NbArt := 0;
  For i := Indice to ListeArticles.Count - 1 do
  begin
    if StIn <> '' then StIn := StIn + ',';
    StIn := StIn + '"' + ListeArticles.Strings[i] + '"';
    Inc(NbArt);
    if NbArt >= NbArtParRequete then Break;
  end;
  if (i = ListeArticles.Count) then Indice := i else Indice := i + 1;
  Result := StIn;
end;

function CreerListeArticle(Const TOBArts : TOB; Const AvecCodeArticle : Boolean=False) : TStringList;
var NbArt, NbArtDim, i : integer;
    ListeArticles : TStringList;
    StArticle : string;
    TOBArtGen : TOB;
    Num_CODEARTICLE, Num_ARTICLE : integer;
begin
  ListeArticles := TStringList.Create;
  ListeArticles.Sorted := True;

  if AvecCodeArticle then
  begin
    Num_CODEARTICLE := TOBArts.Detail[0].GetNumChamp('GA_CODEARTICLE'); //add
    for NbArt := 0 to TOBArts.Detail.Count - 1 do
    begin
      //StArticle := TOBArts.Detail[NbArt].GetValue('GA_CODEARTICLE');
      StArticle := TOBArts.Detail[NbArt].GetValeur(Num_CODEARTICLE); //add
      if Not ListeArticles.Find(StArticle, i) then ListeArticles.Add(StArticle);
    end;
  end else
  begin
    Num_ARTICLE := TOBArts.Detail[0].Detail[0].GetNumChamp('GTL_ARTICLE'); //add
    for NbArt := 0 to TOBArts.Detail.Count - 1 do
    begin
      TOBArtGen := TOBArts.Detail[NbArt];
      for NbArtDim := 0 to TOBArtGen.Detail.Count - 1 do
      begin
        //StArticle := TOBArtGen.Detail[NbArtDim].GetValue('GTL_ARTICLE');
        StArticle := TOBArtGen.Detail[NbArtDim].GetValeur(Num_ARTICLE); //add
        if Not ListeArticles.Find(StArticle, i) then ListeArticles.Add(StArticle);
      end;
    end;
  end;

  Result := ListeArticles;
end;

function ChargeStockLigneEnTOB(ListeArticles : TStringList;
         StPeriode, NomChampSUM, ListeNaturePiece, ListeEtab : string;
         BVivante : Boolean) : TOB;
var
  y : integer;
  StInArticle, StSQL, StVivante, StJointurePiece : string;
  QLigne : TQuery;
  TOBSUM : TOB;
begin
  TOBSUM := TOB.Create('Tob_vituelle', Nil, -1);
  if BVivante then StVivante := ' and GL_VIVANTE = "X"' else StVivante := '';
  if StPeriode = '' then StJointurePiece := ''
  else
    StJointurePiece := 'LEFT JOIN PIECE ON GP_NATUREPIECEG=GL_NATUREPIECEG ' +
    'AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO AND GP_INDICEG=GL_INDICEG ';
  if ListeArticles.Count > 0 then
  begin
    y := 0;
    while y < ListeArticles.Count do
    begin
      StInArticle := ConstruireInArticle(ListeArticles, y);
      if StInArticle <> '' then
      begin
        StSQL := 'Select GL_ARTICLE,GL_ETABLISSEMENT,' +
        'GL_ARTICLE||GL_ETABLISSEMENT AS ARTETAB, SUM(GL_QTEFACT) as ' +
        NomChampSUM + ' From LIGNE ' + StJointurePiece +
        'where GL_ARTICLE IN (' + StInArticle + ') ' +
        'and GL_NATUREPIECEG IN (' + ListeNaturePiece + ') ' +
        StPeriode + ' and GL_ETABLISSEMENT IN (' + ListeEtab + ')' +
        StVivante + ' Group By GL_ARTICLE,GL_ETABLISSEMENT ' +
        'Order By GL_ARTICLE,GL_ETABLISSEMENT';
        QLigne := OpenSQL(StSQL, True);
        if not QLigne.EOF then TOBSUM.LoadDetailDB('', '', '', QLigne, True);
        Ferme(QLigne);
      end;
    end;
  end;

  Result := TOBSUM;
end;

function ChargeStockTransit(ListeArticles : TStringList; BSiege : Boolean): TOB;
var ListeEtab : String;
begin
  if BSiege then ListeEtab := '"' + CodeDepot + '"'
  else ListeEtab := ListeEtabMeth;

  Result := ChargeStockLigneEnTOB(ListeArticles, '', 'TRANSIT', '"ALF","TRV"', ListeEtab, True);
end;

function ChargeStockVente(ListeArticles : TStringList; DebPeriodeVte, FinPeriodeVte : String; BDateIntegr : Boolean): TOB;
var LimiteDate : String;
begin
  if BDateIntegr then
    LimiteDate := ' and GP_DATEINTEGR >= "' + UsDateTime(StrToDate(DebPeriodeVte)) +
    '" and GP_DATEINTEGR <= "' + UsDateTime(StrToDate(FinPeriodeVte)) + '"'
  else
    LimiteDate := ' and GP_DATEPIECE >= "' + UsDateTime(StrToDate(DebPeriodeVte)) +
    '" and GP_DATEPIECE <= "' + UsDateTime(StrToDate(FinPeriodeVte)) + '"';

  Result := ChargeStockLigneEnTOB(ListeArticles, LimiteDate, 'VENDU', '"FFO","FAC"', ListeEtabMeth, False);
end;


// ---------------------------------------------------------------------------
// Chargement en TOB de la quantité à répartir pour chaque article dimensionné
// du dépôt, tenant compte du stock disponible et de la quantité à laisser au dépôt.
// ---------------------------------------------------------------------------

procedure ChargeDepotQteARepartir(FF: TForm; TobAffectationTrf: TOB);
var TobArticle, TobTransit, TOBArtTransit : TOB;
  i, j, NbQteFixe, NbPctStock: integer;
  QteStock, Qte_ARepartir: Double;
  DepotVide, DepotQteFixe, DepotPctStock : boolean;
  Q : TQuery;
  ListeArticles : TStringList;
begin
  if TypePropTrf = 'A' then
  begin
    DepotVide := TCheckBox(FF.FindComponent('VIDDEPOT')).Checked;
    DepotQteFixe := TCheckBox(FF.FindComponent('QTEFIXE')).Checked;
    DepotPctStock := TCheckBox(FF.FindComponent('TXSTOCK')).Checked;
    NbQteFixe := TSpinEdit(FF.FindComponent('NBQTEFIXE')).Value;
    NbPctStock := 100 - (TSpinEdit(FF.FindComponent('NBTXSTOCK')).Value);
  end else
  begin
    DepotVide := True;
    DepotQteFixe := False;
    DepotPctStock := False;
    NbQteFixe := 0;
    NbPctStock := 0;
    Q := OpenSQL('Select GTE_VIDDEPOT, GTE_QTEFIXE, GTE_NBQTEFIXE, GTE_TXSTOCK,' +
    'GTE_NBTXSTOCK From PROPTRANSFENT Where GTE_CODEPTRF="' + CodePropTrf + '"', TRUE);
    if not Q.EOF then
    begin
      if not ((Q.Findfield('GTE_VIDDEPOT').AsString) = 'X') then DepotVide := False;
      if (Q.Findfield('GTE_QTEFIXE').AsString) = 'X' then
      begin
        DepotQteFixe := True;
        NbQteFixe := Q.Findfield('GTE_NBQTEFIXE').AsInteger;
      end;
      if (Q.Findfield('GTE_TXSTOCK').AsString) = 'X' then
      begin
        DepotPctStock := True;
        NbPctStock := 100 - (Q.Findfield('GTE_NBTXSTOCK').AsInteger);
      end;
    end;
    Ferme(Q);
  end;

  TobTransit := Nil;
  ListeArticles := Nil;
  if SiegeGereStockTransit then
  begin
    ListeArticles := CreerListeArticle(TobAffectationTrf);
    TobTransit := ChargeStockTransit(ListeArticles, True);
  end;

  for i := 0 to TobAffectationTrf.Detail.Count - 1 do
  begin
    TobArticle := TobAffectationTrf.Detail[i];
    for j := 0 to TobArticle.Detail.Count - 1 do
    begin
      With TobArticle.Detail[j] do //Cette TOB contient le stock du siège d'un article
      begin
        PutValue('GTL_STOCKALF', 0);
        if SiegeGereStockTransit then
        begin
          //TOBArtTransit := TobTransit.FindFirst(['GL_ARTICLE'], [GetValue('GTL_ARTICLE')], False);
          FindFirstValeur(TOBArtTransit, TobTransit, ['GL_ARTICLE'], [GetValue('GTL_ARTICLE')], False, [TobTransit.GetNumChamp('GL_ARTICLE')], True); //add
          if TOBArtTransit <> Nil then PutValue('GTL_STOCKALF', TOBArtTransit.GetValue('TRANSIT'));
        end;

        if SiegeGereStockNet then QteStock := GetValue('GTL_STOCKINITIAL') +
           GetValue('GTL_RESERVEFOU') - GetValue('GTL_PREPACLI') - GetValue('GTL_RESERVECLI')
        else QteStock := GetValue('GTL_STOCKINITIAL');
        Qte_ARepartir := QteStock - GetValue('GTL_STOCKMINI'); // Valeur par défaut

        if DepotVide then Qte_ARepartir := QteStock
        else if DepotQteFixe then Qte_ARepartir := QteStock - NbQteFixe
        else if DepotPctStock then Qte_ARepartir := Arrondi((QteStock * NbPctStock) / 100.0, 0);
        Qte_ARepartir := Qte_ARepartir + GetValue('GTL_PROPOSITION') - GetValue('GTL_STOCKALF');
        if Qte_ARepartir < 0 then Qte_ARepartir := 0.0;
        PutValue('GTL_SURPLUS', Qte_ARepartir);
      end;
    end;
  end;

  if SiegeGereStockTransit then
  begin
    TobTransit.Free;
    if ListeArticles <> Nil then ListeArticles.Free;
  end;
end;

// -------------------------------------------------------------------
// Traitement des différentes méthodes d'affectation de la proposition
// -------------------------------------------------------------------

procedure TraitementMethodesAffectation(TobAffectationTrf, TobMethode: TOB);
var CodeMethode: string;
  i: integer;
begin
  UtilStockEtab := False;
  RechStockEtab := False;
  AffectMaxi := False;
  for i := 0 to TobMethode.Detail.Count - 1 do
  begin
    CodeMethode := TobMethode.Detail[i].GetValue('CODE');
    if TraitementUneMethode(TobAffectationTrf, CodeMethode) then
      RechStockEtab := True;
  end;
end;

// --------------------------------------------------------
// Traitement d'une méthode d'affectation de la proposition
// --------------------------------------------------------

function TraitementUneMethode(TobAffectationTrf: TOB; CodeMethode: string): boolean;
var QQ: TQuery;
  i: integer;
  TypeMethode: string;
  DebPeriodeVte, FinPeriodeVte: string;
  QteSaisie: double;
  TobEtabMeth: TOB; // Tob stockant les établissements d'une méthode
begin
  Result := True;

  // Chargement des établissements de la méthode en TOB
  TobEtabMeth := TOB.Create('', nil, -1);
  QQ := OpenSQL('Select * from PROPMETETAB where GTQ_CODEMETPAFF="' + CodeMethode + '" ' +
    'Order by GTQ_LIGMETPAFF', True);
  if QQ.EOF then
  begin
    Ferme(QQ);
    Result := False;
    exit;
  end;
  TobEtabMeth.LoadDetailDB('Virtuel_PROPMETETAB', '', '', QQ, False, False);
  Ferme(QQ);
  // Enregistrement dans une chaine de caractères de la liste des établissements de la méthode
  ListeEtabMeth := '';
  for i := 0 to TobEtabMeth.Detail.Count - 1 do
  begin
    if i > 0 then ListeEtabMeth := ListeEtabMeth + ',';
    ListeEtabMeth := ListeEtabMeth + '"' + TobEtabMeth.Detail[i].GetValue('GTQ_ETABLISSEMENT') + '"';
  end;

  // Lecture du paramétrage de la méthode d'affectation
  QQ := OpenSQL('Select * from PROPMETHODE where GTM_CODEMETPAFF="' + CodeMethode + '"', TRUE);
  if QQ.EOF then
  begin
    Ferme(QQ);
    Result := False;
    exit;
  end;

  // Type de la méthode d'affectation
  TypeMethode := QQ.Findfield('GTM_TYPEMETPAFF').AsString;

  // Quantité saisi par dimension pour l'affectation (Qté fixe, qté mini ou qté maxi)
  QteSaisie := 0.0;
  if TypeMethode = 'QTE' then QteSaisie := QQ.Findfield('GTM_QTEDIM').AsFloat
  else if TypeMethode = 'SMS' then QteSaisie := QQ.Findfield('GTM_QTEMINDIM').AsFloat
  else if TypeMethode = 'SXS' then QteSaisie := QQ.Findfield('GTM_QTEMAXDIM').AsFloat;

  // Période de vente
  DebPeriodeVte := QQ.Findfield('GTM_DEBPERVTE').AsString;
  FinPeriodeVte := QQ.Findfield('GTM_FINPERVTE').AsString;

  // Vérification du stock Maxi.
  VerifStkMaxi := Boolean(QQ.Findfield('GTM_VERIFSTKMAX').AsString = 'X');

  // Prise en compte du stock de l'établissement
  if (TypeMethode = 'PVT') or (TypeMethode = 'PVV') then
    UtilStockEtab := Boolean(QQ.Findfield('GTM_UTILSTKETAB').AsString = 'X')
  else if ((TypeMethode = 'PDS') or (TypeMethode = 'QTE')) and
    ((AffectMaxi = False) or (VerifStkMaxi = False)) then
    UtilStockEtab := False
  else UtilStockEtab := True;

  // Application du coef. par boutique
  UtilCoefEtab := Boolean(QQ.Findfield('GTM_UTILCOEFETAB').AsString = 'X');
  if TypeMethode = 'YSM' then UtilCoefEtab := False;

  // Répartition hiérarchique ou par ratio
  RepartRatio := Boolean(QQ.Findfield('GTM_REPARATIO').AsString = 'X');
  if TypeMethode = 'PDS' then RepartRatio := True;

  // Méthode d'arrondi
  MethodeArrondi := QQ.Findfield('GTM_ARRONDI').AsString;

  // Alimentation d'une TOB fille par établissement à approvisionner.
  AffectationEtablissement(TobAffectationTrf, TobEtabMeth, TypeMethode, DebPeriodeVte, FinPeriodeVte, QteSaisie);

  Ferme(QQ);
  TobEtabMeth.Free;
end;

//-----------------------------------------------------------------------
// Pour chaque article dimensionné du Dépôt, alimentation d'une TOB fille
// mémorisant la proposition de transfert qui sera calculée pour chaque
// établissement à approvisionner.
// ----------------------------------------------------------------------

procedure AffectationEtablissement(TobAffectationTrf, TobEtabMeth: TOB; TypeMethode, DebPeriodeVte, FinPeriodeVte: string; QteSaisie: double);
var
  TobArticle, TobDepotAff, TobEtabAff, TobStockEtab, TobVenteEtab: TOB;
  TobMethEtabAff, TobCalculAff, TobTransitEtab : TOB;
  i, j, k : integer;
  CoefRepart, QteARepartir: double;
  QteBesoin, TotBesoin: double;
  IdentArt, CodeArt, StatutArt, CodeEtablis : string;
  ListeArticles : TStringList;
  NGQ_PHYSIQUE, NGQ_STOCKMIN, NGQ_STOCKMAX, NGQ_RESERVECLI, NGQ_RESERVEFOU,
  NGQ_PREPACLI, N_ARTDEPOT : integer;
  NGTL_CODEPTRF, NGTL_ARTICLE, NGTL_CODEARTICLE, NGTL_STATUTART, NGTL_SURPLUS,
  NGTL_DEPOTDEST, NGTL_STOCKINITIAL, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_QTEMIN,
  NGTL_QTEMAX, NGTL_QTEVTE, NGTL_MANQUANT, NGTL_PROPOSITION, NGTL_STOCKLU,
  NGTL_STOCKALF, NGTL_GENERER, NGTL_GENECF, NGTL_RESERVECLI, NGTL_RESERVEFOU,
  NGTL_PREPACLI : integer;
  NGTQ_ETABLISSEMENT, NGTQ_COEFREPAR  : integer;
  NVente_ARTETAB, NTransit_ARTETAB : integer;
begin
  NG_DEPOTDEST := -1;
  NG_BESOIN := -1;
  NG_RATIO := -1;
  NG_AFFECTATION := -1;
  NG_QTESTK := -1;
  NG_QTEMAX := -1;
  ListeArticles := CreerListeArticle(TobAffectationTrf);
  // Création d'une Tob stockant le stock de tous les articles dimensionnés
  // sur tous les établissements à approvisionner.
  TobStockEtab := TOB.Create('Stock-Etablissement', nil, -1);
  Charge_TobStockEtab(TobStockEtab, ListeArticles);
  if (TobStockEtab <> Nil) and (TobStockEtab.Detail.Count > 0) then
  begin
    With TobStockEtab.Detail[0] do
    begin
      NGQ_PHYSIQUE := GetNumChamp('GQ_PHYSIQUE');
      NGQ_STOCKMIN := GetNumChamp('GQ_STOCKMIN');
      NGQ_STOCKMAX := GetNumChamp('GQ_STOCKMAX');
      NGQ_RESERVECLI := GetNumChamp('GQ_RESERVECLI');
      NGQ_RESERVEFOU := GetNumChamp('GQ_RESERVEFOU');
      NGQ_PREPACLI := GetNumChamp('GQ_PREPACLI');
      N_ARTDEPOT := GetNumChamp('ARTDEPOT');
    end;
  end
  else
  begin
    NGQ_PHYSIQUE := -1;
    NGQ_STOCKMIN := -1;
    NGQ_STOCKMAX := -1;
    NGQ_RESERVECLI := -1;
    NGQ_RESERVEFOU := -1;
    NGQ_PREPACLI := -1;
    N_ARTDEPOT := -1;
  end;
  // Création d'une Tob stockant le cumul des ventes de tous les articles dimensionnés
  // pour tous les établissements à approvisionner.
  TobVenteEtab := Nil;
  if (TypeMethode = 'PVT') or (TypeMethode = 'PVV') then
  begin
    TobVenteEtab := ChargeStockVente(ListeArticles, DebPeriodeVte, FinPeriodeVte, TypeMethode = 'PVV');
    if (TobVenteEtab <> Nil) and (TobVenteEtab.Detail.Count > 0) then
      NVente_ARTETAB := TobVenteEtab.detail[0].GetNumChamp('ARTETAB')
    else NVente_ARTETAB := -1;
  end;
  // Création d'une Tob stockant le cumul des annonces de livraison de tous les articles dimensionnés
  // d'un article pour tous les établissements à approvisionner.
  TobTransitEtab := Nil;
  if BTQGereStockTransit then
  begin
    TobTransitEtab := ChargeStockTransit(ListeArticles, False);
    if (TobTransitEtab <> Nil) and (TobTransitEtab.Detail.Count > 0) then
      NTransit_ARTETAB := TobTransitEtab.detail[0].GetNumChamp('ARTETAB')
    else NTransit_ARTETAB := -1;
  end;
  InitMove(TobAffectationTrf.Detail.Count, 'Traitement d''une méthode d''affectation en cours ...');
  if (TobTransitEtab <> Nil) and (TobAffectationTrf.Detail.Count > 0)
   and (TobAffectationTrf.Detail[0].detail.Count > 0) then
  begin
    With TobAffectationTrf.Detail[0].Detail[0] do
    begin
      NGTL_CODEPTRF := GetNumChamp('GTL_CODEPTRF');
      NGTL_ARTICLE := GetNumChamp('GTL_ARTICLE');
      NGTL_CODEARTICLE := GetNumChamp('GTL_CODEARTICLE');
      NGTL_STATUTART := GetNumChamp('GTL_STATUTART');
      NGTL_SURPLUS := GetNumChamp('GTL_SURPLUS');
      NGTL_DEPOTDEST := GetNumChamp('GTL_DEPOTDEST');
      NGTL_STOCKINITIAL := GetNumChamp('GTL_STOCKINITIAL');
      NGTL_STOCKMINI := GetNumChamp('GTL_STOCKMINI');
      NGTL_STOCKMAXI := GetNumChamp('GTL_STOCKMAXI');
      NGTL_QTEMIN := GetNumChamp('GTL_QTEMIN');
      NGTL_QTEMAX := GetNumChamp('GTL_QTEMAX');
      NGTL_QTEVTE := GetNumChamp('GTL_QTEVTE');
      NGTL_MANQUANT := GetNumChamp('GTL_MANQUANT');
      NGTL_PROPOSITION := GetNumChamp('GTL_PROPOSITION');
      NGTL_STOCKLU := GetNumChamp('GTL_STOCKLU');
      NGTL_STOCKALF := GetNumChamp('GTL_STOCKALF');
      NGTL_GENERER := GetNumChamp('GTL_GENERER');
      NGTL_GENECF := GetNumChamp('GTL_GENECF');
      NGTL_RESERVECLI := GetNumChamp('GTL_RESERVECLI');
      NGTL_RESERVEFOU := GetNumChamp('GTL_RESERVEFOU');
      NGTL_PREPACLI := GetNumChamp('GTL_PREPACLI');
    end;
  end
  else
  begin
    NGTL_CODEPTRF := -1;
    NGTL_ARTICLE := -1;
    NGTL_CODEARTICLE := -1;
    NGTL_STATUTART := -1;
    NGTL_SURPLUS := -1;
    NGTL_DEPOTDEST := -1;
    NGTL_STOCKINITIAL := -1;
    NGTL_STOCKMINI := -1;
    NGTL_STOCKMAXI := -1;
    NGTL_QTEMIN := -1;
    NGTL_QTEMAX := -1;
    NGTL_QTEVTE := -1;
    NGTL_MANQUANT := -1;
    NGTL_PROPOSITION := -1;
    NGTL_STOCKLU := -1;
    NGTL_STOCKALF := -1;
    NGTL_GENERER := -1;
    NGTL_GENECF := -1;
    NGTL_RESERVECLI := -1;
    NGTL_RESERVEFOU := -1;
    NGTL_PREPACLI := -1;
  end;
  if (TobEtabMeth <> Nil) and (TobEtabMeth.Detail.Count > 0) then
  begin
    With TobEtabMeth.Detail[0] do
    begin
      NGTQ_ETABLISSEMENT := GetNumChamp('GTQ_ETABLISSEMENT');
      NGTQ_COEFREPAR := GetNumChamp('GTQ_COEFREPAR');
    end;
  end
  else
  begin
    NGTQ_ETABLISSEMENT := -1;
    NGTQ_COEFREPAR := -1;
  end;

  for i := 0 to TobAffectationTrf.Detail.Count - 1 do
  begin
    TobArticle := TobAffectationTrf.Detail[i];
    for j := 0 to TobArticle.Detail.Count - 1 do
    begin
      TobDepotAff := TobArticle.Detail[j];
      CodeArt := TobDepotAff.GetValeur(NGTL_CODEARTICLE);
      IdentArt := TobDepotAff.GetValeur(NGTL_ARTICLE);
      StatutArt := TobDepotAff.GetValeur(NGTL_STATUTART);
      // Abandon de l'article, si plus de marchandise à répartir
      QteARepartir := TobDepotAff.GetValeur(NGTL_SURPLUS);
      if (QteARepartir <= 0) and (ArretRupture = True)  and
        (TypeMethode <> 'SXP') and (TypeMethode <> 'SXS') then Continue;
      // Création d'une TOB qui stockera temporairement l'affectation d'une méthode
      // par établissement. Cette TOB servira au stockage du besoin et au calcul de
      // la répartition par ratio.
      TobMethEtabAff := TOB.Create('Affect-Meth-Etablis', nil, -1);
      // Traitement de l'affectation pour chaque établissement de la méthode
      TotBesoin := 0.0;
      for k := 0 to TobEtabMeth.Detail.Count - 1 do
      begin
        CodeEtablis := TobEtabMeth.Detail[k].GetValeur(NGTQ_ETABLISSEMENT);
        if (UtilCoefEtab = False) and (TypeMethode <> 'PDS') then CoefRepart := 1
        else CoefRepart := TobEtabMeth.Detail[k].GetValeur(NGTQ_COEFREPAR);
        // Recherche s'il existe une Tob fille concernant l'établissement à approvisionner
        TobEtabAff := nil;
        if TobDepotAff.Detail.Count > 0 then
          TobEtabAff := TobDepotAff.FindFirst(['GTL_DEPOTDEST'], [CodeEtablis], False);
        if TobEtabAff = nil then
        begin
          // Création de la Tob fille Etablissement pour l'affectation globale de la proposition.
          TobEtabAff := TOB.Create('PROPTRANSFLIG', TobDepotAff, -1);
          Init_TobEtabAff(TobEtabAff, CodeArt, IdentArt, StatutArt, CodeEtablis,
            NGTL_CODEPTRF, NGTL_ARTICLE, NGTL_CODEARTICLE, NGTL_STATUTART, NGTL_SURPLUS,
            NGTL_DEPOTDEST, NGTL_STOCKINITIAL, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_QTEMIN,
            NGTL_QTEMAX, NGTL_QTEVTE, NGTL_MANQUANT, NGTL_PROPOSITION, NGTL_STOCKLU,
            NGTL_STOCKALF, NGTL_GENERER, NGTL_GENECF);
        end;
        if (TobEtabAff.GetValeur(NGTL_STOCKLU) = '-') then
        begin
          // Chargement dans cette Tob fille, du stock de l'article de l'établissement à approvisionner.
          ChargeStock_TobEtabAff(TobDepotAff, TobEtabAff, TobStockEtab, IdentArt, CodeEtablis,
            NGTL_STOCKINITIAL, NGTL_SURPLUS, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_STOCKLU,
            NGTL_RESERVECLI, NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_PROPOSITION, NGQ_PHYSIQUE,
            NGQ_STOCKMIN, NGQ_STOCKMAX, NGQ_RESERVECLI, NGQ_RESERVEFOU, NGQ_PREPACLI, N_ARTDEPOT );
          if BTQGereStockTransit then
            ChargeTransit_TobEtabAff(TobEtabAff, TobTransitEtab, IdentArt, CodeEtablis, NTransit_ARTETAB);
        end;
        if (TypeMethode = 'PVT') or (TypeMethode = 'PVV') then
          // Recherche des ventes de l'article sur une période, pour l'établissement à approvisionner
          ChargeVente_TobEtabAff(TobEtabAff, TobVenteEtab, IdentArt, CodeEtablis, DebPeriodeVte, FinPeriodeVte, NVente_ARTETAB);
        if (TypeMethode = 'SXP') or (TypeMethode = 'SXS') then
        begin
          // Méthode permettant de contrôler la quantité maximum à envoyer dans chaque établissement.
          AffectMaxi := True;
          // Chargement en Tob de la Quantité Maxi et mise à jour de la proposition
          // si le nouveau stock dépasse cette valeur.
          ChargeQteMax_TobEtabAff(TobDepotAff, TobEtabAff, TypeMethode, QteSaisie, CoefRepart,
            NGTL_STOCKMAXI, NGTL_QTEMAX, NGTL_STOCKINITIAL, NGTL_RESERVEFOU,
            NGTL_PREPACLI, NGTL_RESERVECLI, NGTL_PROPOSITION, NGTL_SURPLUS, NGTL_MANQUANT);
        end else
        begin
          if TypeMethode <> 'YSM' then
          begin
            // Création d'une TOB fille pour le calcul d'affectation d'une méthode par établissement.
            TobCalculAff := TOB.Create('Affect-Meth-Etab', TobMethEtabAff, -1);
            Init_TobCalculAff(TobCalculAff, CodeEtablis);
            // Mémorisation du besoin de l'établissement suivant la méthode
            QteBesoin := ChargeBesoin_TobCalculAff(TobEtabAff, TobCalculAff,
              TypeMethode, QteSaisie, CoefRepart, NGTL_STOCKINITIAL, NGTL_RESERVEFOU,
              NGTL_PREPACLI, NGTL_RESERVECLI, NGTL_STOCKALF, NGTL_PROPOSITION,
              NGTL_STOCKMINI, NGTL_QTEVTE, NGTL_QTEMIN, NGTL_QTEMAX);
            TotBesoin := TotBesoin + QteBesoin;
          end;
        end;
      end; // fin boucle sur chaque établissement de la méthode

      if (TobMethEtabAff.Detail.Count > 0) and ((TotBesoin > 0) or (TypeMethode = 'PDS')) then
      begin
        // Calcul de l'affectation en fonction du besoin de chaque établissement et
        // de la quantité totale à répartir.
        QteARepartir := TobDepotAff.GetValeur(NGTL_SURPLUS);
        ChargeAffectation_TobCalculAff(TobMethEtabAff, TypeMethode, QteARepartir, TotBesoin);
        // Mise à jour de la TOB mémorisant la proposition d'affectation globale de
        // chaque établissement et mise à jour sur le Dépôt de la quantité totale affectée.
        ChargeProposition_TobEtabAff(TobDepotAff, TobMethEtabAff, TypeMethode,
          NGTL_PROPOSITION, NGTL_MANQUANT, NGTL_SURPLUS);
      end;

      TobMethEtabAff.Free;
    end;
    MoveCur(False);
  end;
  FiniMove;
  if (TypeMethode = 'PVT') OR (TypeMethode = 'PVV') then TobVenteEtab.Free;
  if BTQGereStockTransit then TobTransitEtab.Free;
  TobStockEtab.Free;
  ListeArticles.Free;
end;

// ----------------------------------------------------------------------------
// Chargement en TOB depuis la base de données, du stock de tous les articles
// dimensionnés d'un article sur l'ensemble des établissements à approvisionner.
// ----------------------------------------------------------------------------

procedure Charge_TobStockEtab(TobStockEtab: TOB; ListeArticles : TStringList);
var
  SQL, LesArticles : string;
  Q: TQUERY;
  Compt : integer;
begin
  Compt := 0;
  While Compt < ListeArticles.Count do
  begin
    LesArticles := ConstruireInArticle(ListeArticles, Compt);
    if LesArticles <> '' then
    begin
      SQL := 'Select GQ_ARTICLE,GQ_DEPOT,GQ_ARTICLE||GQ_DEPOT AS ARTDEPOT,' +
      'GQ_PHYSIQUE,GQ_STOCKMIN,GQ_STOCKMAX,GQ_RESERVEFOU,GQ_RESERVECLI,' +
      'GQ_PREPACLI From DISPO Where GQ_ARTICLE IN (' + LesArticles +
      ') and GQ_DEPOT IN (' + ListeEtabMeth + ') and GQ_CLOTURE="-" ' +
      'ORDER BY GQ_ARTICLE,GQ_DEPOT';
      Q := OpenSQL(SQL, True);
      if not Q.EOF then TobStockEtab.LoadDetailDB('', '', '', Q, True, False);
      Ferme(Q);
    end;
  end;
end;

// -----------------------------------------------
// Initialisation des champs de la Tob: TobEtabAff
// -----------------------------------------------

procedure Init_TobEtabAff(TobEtabAff: TOB; CodeArt, IdentArt, StatutArt, CodeEtablis: string;
  NGTL_CODEPTRF, NGTL_ARTICLE, NGTL_CODEARTICLE, NGTL_STATUTART, NGTL_SURPLUS,
  NGTL_DEPOTDEST, NGTL_STOCKINITIAL, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_QTEMIN,
  NGTL_QTEMAX, NGTL_QTEVTE, NGTL_MANQUANT, NGTL_PROPOSITION, NGTL_STOCKLU,
  NGTL_STOCKALF, NGTL_GENERER, NGTL_GENECF : integer);
begin
  with TobEtabAff do
  begin
    PutValeur(NGTL_CODEPTRF, CodePropTrf);
    PutValeur(NGTL_ARTICLE, IdentArt);
    PutValeur(NGTL_CODEARTICLE, CodeArt);
    PutValeur(NGTL_STATUTART, StatutArt);
    PutValeur(NGTL_DEPOTDEST, CodeEtablis);
    PutValeur(NGTL_STOCKINITIAL, 0.0);
    PutValeur(NGTL_STOCKMINI, 0.0);
    PutValeur(NGTL_STOCKMAXI, 0.0);
    PutValeur(NGTL_QTEMIN, 0.0);
    PutValeur(NGTL_QTEMAX, 0.0);
    PutValeur(NGTL_QTEVTE, 0.0);
    PutValeur(NGTL_MANQUANT, 0.0);
    PutValeur(NGTL_SURPLUS, 0.0);
    PutValeur(NGTL_PROPOSITION, 0.0);
    PutValeur(NGTL_STOCKLU, '-');
    PutValeur(NGTL_STOCKALF, 0.0);
    PutValeur(NGTL_GENERER, '-');
    PutValeur(NGTL_GENECF, '-');
  end;
end;

// -------------------------------------------------------------------------------------
// Chargement dans la TOB d'affectation globale par établissement, du stock de l'article.
// -------------------------------------------------------------------------------------

procedure ChargeStock_TobEtabAff(TobDepotAff, TobEtabAff, TobStockEtab: TOB;
  IdentArt, CodeEtablis: string;
  NGTL_STOCKINITIAL, NGTL_SURPLUS, NGTL_STOCKMINI, NGTL_STOCKMAXI, NGTL_STOCKLU,
  NGTL_RESERVECLI, NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_PROPOSITION,
  NGQ_PHYSIQUE, NGQ_STOCKMIN, NGQ_STOCKMAX, NGQ_RESERVECLI, NGQ_RESERVEFOU,
  NGQ_PREPACLI, N_ARTDEPOT : integer );
var TobS: TOB;
    StkPhysique : Double;
begin
  if CodeEtablis = CodeDepot then
  begin
    TobEtabAff.PutValeur(NGTL_STOCKINITIAL, TobDepotAff.GetValeur(NGTL_STOCKINITIAL)
    - TobDepotAff.GetValeur(NGTL_SURPLUS) + TobDepotAff.GetValeur(NGTL_PROPOSITION));
    TobEtabAff.PutValeur(NGTL_STOCKMINI, TobDepotAff.GetValeur(NGTL_STOCKMINI));
    TobEtabAff.PutValeur(NGTL_STOCKMAXI, TobDepotAff.GetValeur(NGTL_STOCKMAXI));
    TobEtabAff.PutValeur(NGTL_RESERVECLI, TobDepotAff.GetValeur(NGTL_RESERVECLI));
    TobEtabAff.PutValeur(NGTL_RESERVEFOU, TobDepotAff.GetValeur(NGTL_RESERVEFOU));
    TobEtabAff.PutValeur(NGTL_PREPACLI, TobDepotAff.GetValeur(NGTL_PREPACLI));
  end
  else
  begin
    // A ne plus faire ICI
    //Charge_TobStockEtab(TobStockEtab, IdentArt);
    //TobS := TobStockEtab.FindFirst(['GQ_ARTICLE', 'GQ_DEPOT'], [IdentArt, CodeEtablis], False);
    FindFirstValeur(TobS, TobStockEtab, ['ARTDEPOT'], [IdentArt + CodeEtablis], False, [N_ARTDEPOT], True);
    if TobS <> nil then
    begin
      StkPhysique := TobS.GetValeur(NGQ_PHYSIQUE);
      if (StkPhysique < 0.0) then TobEtabAff.PutValeur(NGTL_STOCKINITIAL, 0.0)
      else TobEtabAff.PutValeur(NGTL_STOCKINITIAL, StkPhysique);
      TobEtabAff.PutValeur(NGTL_STOCKMINI, TobS.GetValeur(NGQ_STOCKMIN));
      TobEtabAff.PutValeur(NGTL_STOCKMAXI, TobS.GetValeur(NGQ_STOCKMAX));
      TobEtabAff.PutValeur(NGTL_RESERVECLI, TobS.GetValeur(NGQ_RESERVECLI));
      TobEtabAff.PutValeur(NGTL_RESERVEFOU, TobS.GetValeur(NGQ_RESERVEFOU));
      TobEtabAff.PutValeur(NGTL_PREPACLI, TobS.GetValeur(NGQ_PREPACLI));

      TobS.Free;
    end;
  end;
  TobEtabAff.PutValeur(NGTL_STOCKLU, 'X');
end;

// ------------------------------------------------------------------------
// Chargement dans la TOB d'affectation globale par établissement, du cumul
// des ventes pour un article dimensionné.
// ------------------------------------------------------------------------

procedure ChargeVente_TobEtabAff(TobEtabAff, TobVenteEtab: TOB; IdentArt, CodeEtablis, DebPeriodeVte, FinPeriodeVte: string; NVente_ARTETAB : integer);
var TobV: TOB;
begin
  //TobV := TobVenteEtab.FindFirst(['GL_ARTICLE', 'GL_ETABLISSEMENT'], [IdentArt, CodeEtablis], False);
  FindFirstValeur(TobV, TobVenteEtab, ['ARTETAB'], [IdentArt + CodeEtablis], False, [NVente_ARTETAB], True); //add
  if TobV <> nil then
  begin
    TobEtabAff.PutValue('GTL_QTEVTE', TobV.GetValue('VENDU'));
    TobEtabAff.PutValue('GTL_DEBPERVTE', StrToDate(DebPeriodeVte));
    TobEtabAff.PutValue('GTL_FINPERVTE', StrToDate(FinPeriodeVte));

    TobV.Free;
  end;
end;

// ------------------------------------------------------------------------
// Chargement dans la TOB d'affectation globale par établissement, du cumul
// des Annonces de livraison pour un article dimensionné.
// ------------------------------------------------------------------------

procedure ChargeTransit_TobEtabAff(TobEtabAff, TobTransitEtab: TOB; IdentArt, CodeEtablis : string; NTransit_ARTETAB : integer);
var TobV: TOB;
begin
  //TobV := TobTransitEtab.FindFirst(['GL_ARTICLE', 'GL_ETABLISSEMENT'], [IdentArt, CodeEtablis], False);
  FindFirstValeur(TobV, TobTransitEtab, ['ARTETAB'], [IdentArt + CodeEtablis], False, [NTransit_ARTETAB], True); //add
  if TobV <> nil then
  begin
    TobEtabAff.PutValue('GTL_STOCKALF', TobV.GetValue('TRANSIT'));
    TobV.Free;
  end;
end;

// ----------------------------------------------------------------------
// Chargement en Tob de la Quantité Maxi et mise à jour de la proposition
// si le nouveau stock dépasse cette valeur.
// ----------------------------------------------------------------------

procedure ChargeQteMax_TobEtabAff(TobDepotAff, TobEtabAff: TOB; TypeMethode: string;
  QteSaisie, CoefRepart: Double; NGTL_STOCKMAXI, NGTL_QTEMAX, NGTL_STOCKINITIAL,
  NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_RESERVECLI, NGTL_PROPOSITION,
  NGTL_SURPLUS, NGTL_MANQUANT : integer);
var QteStock, QteMaxi, QtePropo : double;
  QteManquant, QteNew, QteDiff : double;
begin
  // Quantité maxi à ne pas dépasser
  if TypeMethode = 'SXS' then QteMaxi := ArrondirQte(MethodeArrondi, (QteSaisie * CoefRepart))
  else QteMaxi := ArrondirQte(MethodeArrondi, (TobEtabAff.GetValeur(NGTL_STOCKMAXI) * CoefRepart));
  TobEtabAff.PutValeur(NGTL_QTEMAX, QteMaxi);
  // Contrôle que la proposition ne dépasse pas cette quantité Maxi
  With TobEtabAff do
  begin
    if BTQGereStockNet then QteStock := GetValeur(NGTL_STOCKINITIAL) +
      GetValeur(NGTL_RESERVEFOU) - GetValeur(NGTL_PREPACLI) - GetValeur(NGTL_RESERVECLI)
    else QteStock := GetValeur(NGTL_STOCKINITIAL);
    QtePropo := GetValeur(NGTL_PROPOSITION);
  end;
  if (QtePropo > 0) and (QteStock + QtePropo > QteMaxi) then
  begin
    QteNew := QteMaxi - QteStock;
    if QteNew < 0 then QteNew := 0.0;
    QteDiff := QtePropo - QteNew;
    // Mise à jour de la proposition sur l'établissement destinataire
    TobEtabAff.PutValeur(NGTL_PROPOSITION, QteNew);
    // Mise à jour de la proposition sur l'établissement émetteur (Dépôt)
    TobDepotAff.PutValeur(NGTL_PROPOSITION, TobDepotAff.GetValeur(NGTL_PROPOSITION) + QteDiff);
    TobDepotAff.PutValeur(NGTL_SURPLUS, TobDepotAff.GetValeur(NGTL_SURPLUS) + QteDiff);
    QtePropo := QteNew;
  end;
  // Contrôle que le manquant ne dépasse pas cette quantité Maxi
  QteManquant := TobEtabAff.GetValeur(NGTL_MANQUANT);
  if (QteManquant > 0) and (QteStock + QtePropo + QteManquant > QteMaxi) then
  begin
    QteNew := QteMaxi - QteStock - QtePropo;
    if QteNew < 0 then QteNew := 0.0;
    QteDiff := QteManquant - QteNew;
    // Mise à jour du manquant sur l'établissement destinataire
    TobEtabAff.PutValeur(NGTL_MANQUANT, QteNew);
    // Mise à jour du manquant sur l'établissement émetteur (Dépôt)
    TobDepotAff.PutValeur(NGTL_MANQUANT, TobDepotAff.GetValeur(NGTL_MANQUANT) - QteDiff);
  end;
end;

// -------------------------------------------------------------------------------
// Mise à jour de la TOB mémorisant la proposition d'affectation globale de
// chaque établissement et mise à jour sur le Dépôt de la quantité totale affectée.
// -------------------------------------------------------------------------------

procedure ChargeProposition_TobEtabAff(TobDepotAff, TobMethEtabAff: TOB;
  TypeMethode: string; NGTL_PROPOSITION, NGTL_MANQUANT, NGTL_SURPLUS : integer);
var i: integer;
  Affectation, TotAffectation, TotSurplus, Manquant_Meth, Manquant_Propo: double;
  Manquant_Prec, Manquant_New, TotManquant: double;
  CodeEtablis: string;
  TobCalculAff, TobEtabAff: TOB;
begin
  TotAffectation := 0.0;
  TotSurplus := 0.0;
  TotManquant := 0.0;
  for i := 0 to TobMethEtabAff.Detail.Count - 1 do
  begin
    TobCalculAff := TobMethEtabAff.Detail[i];
    Affectation := TobCalculAff.GetValeur(NG_AFFECTATION);
    if TypeMethode = 'PDS' then Manquant_Meth := 0.0
    else Manquant_Meth := TobCalculAff.GetValeur(NG_BESOIN) - Affectation;

    CodeEtablis := TobCalculAff.GetValeur(NG_DEPOTDEST);
    TobEtabAff := TobDepotAff.FindFirst(['GTL_DEPOTDEST'], [CodeEtablis], False);
    if TobEtabAff <> nil then
    begin
      if Affectation <> 0 then
      begin
        TobEtabAff.PutValeur(NGTL_PROPOSITION, TobEtabAff.GetValeur(NGTL_PROPOSITION) + Affectation);
        if CodeEtablis <> CodeDepot then TotAffectation := TotAffectation + Affectation;
        TotSurplus := TotSurplus + Affectation;
      end;
      Manquant_Prec := TobEtabAff.GetValeur(NGTL_MANQUANT);
      Manquant_Propo := Manquant_Prec - Affectation;
      if Manquant_Propo > Manquant_Meth
        then Manquant_New := Manquant_Propo
      else Manquant_New := Manquant_Meth;
      TobEtabAff.PutValeur(NGTL_MANQUANT, Manquant_New);
      TotManquant := TotManquant + Manquant_New - Manquant_Prec;
    end;
  end;
  TobDepotAff.PutValeur(NGTL_PROPOSITION, TobDepotAff.GetValeur(NGTL_PROPOSITION) - TotAffectation);
  TobDepotAff.PutValeur(NGTL_SURPLUS, TobDepotAff.GetValeur(NGTL_SURPLUS) - TotSurplus);
  TobDepotAff.PutValeur(NGTL_MANQUANT, TobDepotAff.GetValeur(NGTL_MANQUANT) + TotManquant);
end;

// -------------------------------------------------
// Initialisation des champs de la Tob: TobCalculAff
// -------------------------------------------------

procedure Init_TobCalculAff(TobCalculAff: TOB; CodeEtablis: string);
begin
  With TobCalculAff do
  begin
    AddChampSupValeur('DEPOTDEST', CodeEtablis);
    AddChampSupValeur('BESOIN', 0.0);
    AddChampSupValeur('RATIO', 0.0);
    AddChampSupValeur('AFFECTATION', 0.0);
    AddChampSupValeur('QTESTK', 0.0);
    AddChampSupValeur('QTEMAX', 0.0);
    if (NG_DEPOTDEST = -1) then
    begin
      NG_DEPOTDEST := GetNumChamp('DEPOTDEST');
      NG_BESOIN := GetNumChamp('BESOIN');
      NG_RATIO := GetNumChamp('RATIO');
      NG_AFFECTATION := GetNumChamp('AFFECTATION');
      NG_QTESTK := GetNumChamp('QTESTK');
      NG_QTEMAX := GetNumChamp('QTEMAX');
    end;
  end;
end;

// ------------------------------------------------------------------------------------
// Calcul du besoin de l'établissement suivant la méthode et stockage dans TobCalculAff
// ------------------------------------------------------------------------------------
function ChargeBesoin_TobCalculAff(TobEtabAff, TobCalculAff: TOB;
  TypeMethode: string; QteSaisie, CoefRepart: Double;
  NGTL_STOCKINITIAL, NGTL_RESERVEFOU, NGTL_PREPACLI, NGTL_RESERVECLI,
  NGTL_STOCKALF, NGTL_PROPOSITION, NGTL_STOCKMINI, NGTL_QTEVTE, NGTL_QTEMIN,
  NGTL_QTEMAX : integer): Double;
var QteStock, QteMini, QteMaxi, QteVendu, QtePropo, QteBesoin, QteTransit: double;
begin
  With TobEtabAff do
  begin
    if BTQGereStockNet then QteStock := GetValeur(NGTL_STOCKINITIAL) +
      GetValeur(NGTL_RESERVEFOU) - GetValeur(NGTL_PREPACLI) - GetValeur(NGTL_RESERVECLI)
    else QteStock := GetValeur(NGTL_STOCKINITIAL);
    QteTransit := TobEtabAff.GetValeur(NGTL_STOCKALF);
    QtePropo := TobEtabAff.GetValeur(NGTL_PROPOSITION) + QteTransit;
  end;
  QteMini := 0.0;
  QteBesoin := 0.0;
  if TypeMethode = 'QTE' then QteBesoin := ArrondirQte(MethodeArrondi, QteSaisie * CoefRepart)
  else if TypeMethode = 'SMS' then
  begin
    QteMini := ArrondirQte(MethodeArrondi, QteSaisie * CoefRepart);
    QteBesoin := QteMini - QteStock - QtePropo;
  end
  else if TypeMethode = 'SMP' then
  begin
    QteMini := ArrondirQte(MethodeArrondi, (TobEtabAff.GetValeur(NGTL_STOCKMINI) * CoefRepart));
    QteBesoin := QteMini - QteStock - QtePropo;
  end
  else if (TypeMethode = 'PVT') or (TypeMethode = 'PVV') then
  begin
    QteVendu := ArrondirQte(MethodeArrondi, (TobEtabAff.GetValeur(NGTL_QTEVTE) * CoefRepart));
    QteBesoin := QteVendu - QtePropo;
    if UtilStockEtab then QteBesoin := QteBesoin - QteStock;
  end
  else if TypeMethode = 'PDS' then
  begin
    QteBesoin := 0.0; // Calculé plus tard
    TobCalculAff.PutValeur(NG_RATIO, CoefRepart);
  end;

  // Chargement en Tob de la Quantité Mini suivant le type de méthode appliqué.
  if QteMini <> 0 then TobEtabAff.PutValeur(NGTL_QTEMIN, QteMini);

  // "Ecrétage" du besoin, si celui-ci dépasse le seuil Maxi.
  // Mémorisation du Stock et de la quantité Maxi en Tob, pour pouvoir affecter par
  // la suite la quantité résiduelle générée par les arrondis de calcul de répartition,
  // en tenant compte là aussi du seuil Maxi.
  if (AffectMaxi = True) and (VerifStkMaxi = True) then
  begin
    QteMaxi := TobEtabAff.GetValeur(NGTL_QTEMAX);
    TobCalculAff.PutValeur(NG_QTESTK, QteStock + QtePropo);
    TobCalculAff.PutValeur(NG_QTEMAX, QteMaxi);
    if (QteBesoin > 0) and (QteStock + QtePropo + QteBesoin > QteMaxi) then
      QteBesoin := QteMaxi - QteStock - QtePropo;
  end;
  if QteBesoin < 0 then QteBesoin := 0.0;
  TobCalculAff.PutValeur(NG_BESOIN, QteBesoin);
  Result := QteBesoin;
end;

// ------------------------------------------------------------------------
// Calcul de l'affectation en fonction du besoin de chaque établissement et
// de la quantité totale à répartir.
// ------------------------------------------------------------------------

procedure ChargeAffectation_TobCalculAff(TobMethEtabAff: TOB; TypeMethode: string; QteARepartir, TotBesoin: Double);
var i: integer;
  MajAff: boolean;
  Ratio, Affectation, Reste, Besoin, QteStock, QteMaxi: double;
  CodeEtablis: string;
  TobCalculAff: TOB;
begin
  Reste := QteARepartir;
  for i := 0 to TobMethEtabAff.Detail.Count - 1 do
  begin
    TobCalculAff := TobMethEtabAff.Detail[i];

    if TypeMethode = 'PDS' then
    begin
      // Répartion en fonction du poids de l'établissement
      Ratio := TobCalculAff.GetValeur(NG_RATIO);
      Affectation := ArrondirQte(MethodeArrondi, (QteARepartir * Ratio) / 100.0);
      Reste := Reste - Affectation;
      TobCalculAff.PutValeur(NG_AFFECTATION, Affectation);
      TobCalculAff.PutValeur(NG_BESOIN, Affectation);
      TotBesoin := QteARepartir + 1;
    end
    else
    begin
      if RepartRatio = True then
      begin
        Ratio := Arrondi((TobCalculAff.GetValeur(NG_BESOIN) / TotBesoin) * 100.0, 2);
        TobCalculAff.PutValeur(NG_RATIO, Ratio);
        if QteARepartir < TotBesoin then
          Affectation := ArrondirQte(MethodeArrondi, (QteARepartir * Ratio) / 100.0)
        else Affectation := TobCalculAff.GetValeur(NG_BESOIN);
      end else
      begin
        Affectation := TobCalculAff.GetValeur(NG_BESOIN);
      end;
      if (RepartRatio = False) and (Affectation > Reste) then Affectation := Reste;
      Reste := Reste - Affectation;
      TobCalculAff.PutValeur(NG_AFFECTATION, Affectation);
    end;
  end;

  // Test si une quantité résiduelle a été générée suite aux écarts d'arrondis.
  if (RepartRatio = True) and (QteARepartir < TotBesoin) and ((Reste < -0.5) or (Reste > 0.5)) then
  begin
    if Reste < -0.5 then
    begin
      // Si des pièces ont été transférées en trop, parcours des établissements
      // en partant du bas de la liste avec retrait d'une pièce par établissement
      // jusqu'à régulation.
      // Si le dépôt figure parmi les établissements à approvisionner, le retrait sur
      // celui-ci peut être total, afin de favoriser les autres établissements.
      for i := TobMethEtabAff.Detail.Count - 1 downto 0 do
      begin
        TobCalculAff := TobMethEtabAff.Detail[i];
        Affectation := TobCalculAff.GetValeur(NG_AFFECTATION);
        CodeEtablis := TobCalculAff.GetValeur(NG_DEPOTDEST);
        while Affectation > 0.5 do
        begin
          Affectation := Affectation - 1;
          Reste := Reste + 1;
          TobCalculAff.PutValeur(NG_AFFECTATION, Affectation);
          if Reste > -0.5 then break;
          if CodeEtablis <> CodeDepot then break;
        end;
        if Reste > -0.5 then break;
      end;
    end;

    if Reste > 0.5 then
    begin
      // Si des pièces ont été transférées en moins, parcours des établissements
      // en partant du haut de la liste avec ajout d'une pièce jusqu'à régulation.
      for i := 0 to TobMethEtabAff.Detail.Count - 1 do
      begin
        TobCalculAff := TobMethEtabAff.Detail[i];
        Besoin := TobCalculAff.GetValeur(NG_BESOIN);
        Affectation := TobCalculAff.GetValeur(NG_AFFECTATION);
        MajAff := False;
        if (Besoin > Affectation) or ((TypeMethode = 'PDS')
          and (TobCalculAff.GetValeur(NG_RATIO) > 0)) then
        begin
          Affectation := Affectation + 1;
          MajAff := True;
        end;

        // Test si l'affectation ne dépasse pas le seuil Maxi.
        if (MajAff = True) and (AffectMaxi = True) and (VerifStkMaxi = True) then
        begin
          QteStock := TobCalculAff.GetValeur(NG_QTESTK);
          QteMaxi := TobCalculAff.GetValeur(NG_QTEMAX);
          if (QteStock + Affectation > QteMaxi) then MajAff := False;
        end;

        if MajAff = True then
        begin
          Reste := Reste - 1;
          TobCalculAff.PutValeur(NG_AFFECTATION, Affectation);
          if Reste < 0.5 then break;
        end;
      end;
    end;
  end;

  // Dans le cas d'une répartition en fonction du poids de l'établissement,
  // l'"Ecrétage" de l'affectation, si celui-ci dépasse le seuil Maxi, n'a pas
  // encore été fait.
  if (TypeMethode = 'PDS') and (AffectMaxi = True) and (VerifStkMaxi = True) then
  begin
    for i := 0 to TobMethEtabAff.Detail.Count - 1 do
    begin
      TobCalculAff := TobMethEtabAff.Detail[i];
      Affectation := TobCalculAff.GetValeur(NG_AFFECTATION);
      QteStock := TobCalculAff.GetValeur(NG_QTESTK);
      QteMaxi := TobCalculAff.GetValeur(NG_QTEMAX);
      if (Affectation > 0) and (QteStock + Affectation > QteMaxi) then
        Affectation := QteMaxi - QteStock;
      if Affectation < 0 then Affectation := 0.0;
      TobCalculAff.PutValeur(NG_AFFECTATION, Affectation);
    end;
  end;
end;

// ---------------------------------------------------------------------------
// Alimentation de la Tob finale au format de la table stockant la proposition
// de transfert. Ajout du cumul sur les articles génériqies.
// ---------------------------------------------------------------------------

procedure Alimentation_TobEcrPTrf(TobAffectationTrf, TobEcrPTrf: TOB);
var TobArticle: TOB;
  TobDepotAff: TOB;
  TobEtabAff: TOB;
  TobPTrf: TOB;
  i, j, k: integer;
  abandon: boolean;
  CodeEtablis: string;
begin
  InitMove(TobAffectationTrf.Detail.Count, 'Mise en forme des données en cours...');
  for i := 0 to TobAffectationTrf.Detail.Count - 1 do
  begin
    TobArticle := TobAffectationTrf.Detail[i];

    for j := 0 to TobArticle.Detail.Count - 1 do
    begin
      // Affectation du dépôt émetteur pour un article
      TobDepotAff := TobArticle.Detail[j];

      Abandon := False;
      if (ArretRupture = False) and (TypePropTrf = 'A') then
      begin
        // Si pas de mouvement de stock enregistré sur l'article, celui-ci est
        // abandonné.
        if (TobDepotAff.GetValue('GTL_STOCKINITIAL') <= 0) and
          (TobDepotAff.GetValue('GTL_MANQUANT') = 0.0) and
          (TobDepotAff.GetValue('GTL_SURPLUS') = 0.0) and
          (TobDepotAff.GetValue('GTL_PROPOSITION') = 0.0) then abandon := True;
      end;

      if LimiteEnreg then
      begin
        // Si pas de proposition, suppresion de l'enregistrement
        if (TobDepotAff.GetValue('GTL_PROPOSITION') = 0.0)
          and (TobDepotAff.GetValue('GTL_MANQUANT') = 0.0) then abandon := True;
      end;

      if Abandon = False then
      begin
        // Duplication de l'affectation dans la nouvelle Tob
        TobPTrf := TOB.Create('PROPTRANSFLIG', TobEcrPTrf, -1);
        TobPTrf.Dupliquer(TobDepotAff, False, True, True);

        for k := 0 to TobDepotAff.Detail.Count - 1 do
        begin
          // Affectation de l'établissement destinataire pour un article dimensionné
          TobEtabAff := TobDepotAff.Detail[k];
          CodeEtablis := TobEtabAff.GetValue('GTL_DEPOTDEST');
          if CodeEtablis <> CodeDepot then
          begin
            if not ((LimiteEnreg) and (TobEtabAff.GetValue('GTL_PROPOSITION') = 0.0)
              and (TobEtabAff.GetValue('GTL_MANQUANT') = 0.0)) then
            begin //Création des lignes > 0 dans le cas des ventes
              // Duplication de l'affectation dans la nouvelle Tob
              TobPTrf := TOB.Create('PROPTRANSFLIG', TobEcrPTrf, -1);
              TobPTrf.Dupliquer(TobEtabAff, False, True, True);
            end;
          end;
        end;
      end;
      TobDepotAff.ClearDetail;
    end;
    TobArticle.ClearDetail;
    MoveCur(False);
  end;
  TobAffectationTrf.ClearDetail;
  FiniMove;

  AjouteLesLignesGEN(TobEcrPTrf); // JD New 18_06_2002
end;

procedure AjouteLesLignesGEN(TOBLignesDIM: TOB);
var i: Integer;
  StCodeArticle, StDepotArticle, IdentArt: string;
  TOBLigneGEN: TOB;
  StockInitial, StockMini, StockMaxi, QteMin, QteMax, QteVte,
  Manquant, Surplus, Proposition, StockTransit, StockNet,
  ReserveCli, ReserveFou, PrepaCli : Double;
begin
  TOBLignesDIM.Detail.Sort('GTL_CODEPTRF;GTL_DEPOTDEST;GTL_ARTICLE');
  InitMove(TOBLignesDIM.Detail.Count, 'Mis à jour des articles génériques...');
  i := 0;
  while i < TOBLignesDIM.Detail.Count do
  begin
    if TOBLignesDIM.Detail[i].GetValue('GTL_STATUTART') <> 'DIM' then
    begin
      i := i + 1;
      Continue;
      MoveCur(False);
    end;
    TOBLigneGEN := TOB.Create('PROPTRANSFLIG', TOBLignesDIM, i);
    i := i + 1;
    StCodeArticle := TOBLignesDIM.Detail[i].GetValue('GTL_CODEARTICLE');
    StDepotArticle := TOBLignesDIM.Detail[i].GetValue('GTL_DEPOTDEST');
    IdentArt := CodeArticleUnique2(StCodeArticle, '');
    TOBLigneGEN.Dupliquer(TOBLignesDIM.Detail[i], False, True, True);
    with TOBLigneGEN do
    begin
      PutValue('GTL_ARTICLE', IdentArt);
      PutValue('GTL_STATUTART', 'GEN');
      //Fait la somme des lignes dimensionnées de cette article générique
      StockInitial := 0;
      StockMini := 0;
      StockMaxi := 0;
      QteMin := 0;
      QteMax := 0;
      QteVte := 0;
      Manquant := 0;
      Surplus := 0;
      Proposition := 0;
      StockTransit := 0;
      ReserveCli := 0;
      ReserveFou := 0;
      PrepaCli := 0;
      StockNet := 0;
      while (i < TOBLignesDIM.Detail.Count) and
        (StCodeArticle = TOBLignesDIM.Detail[i].GetValue('GTL_CODEARTICLE')) and
        (StDepotArticle = TOBLignesDIM.Detail[i].GetValue('GTL_DEPOTDEST')) do
      begin
        StockInitial := StockInitial + TOBLignesDIM.Detail[i].GetValue('GTL_STOCKINITIAL');
        StockMini := StockMini + TOBLignesDIM.Detail[i].GetValue('GTL_STOCKMINI');
        StockMaxi := StockMaxi + TOBLignesDIM.Detail[i].GetValue('GTL_STOCKMAXI');
        QteMin := QteMin + TOBLignesDIM.Detail[i].GetValue('GTL_QTEMIN');
        QteMax := QteMax + TOBLignesDIM.Detail[i].GetValue('GTL_QTEMAX');
        QteVte := QteVte + TOBLignesDIM.Detail[i].GetValue('GTL_QTEVTE');
        Manquant := Manquant + TOBLignesDIM.Detail[i].GetValue('GTL_MANQUANT');
        Surplus := Surplus + TOBLignesDIM.Detail[i].GetValue('GTL_SURPLUS');
        Proposition := Proposition + TOBLignesDIM.Detail[i].GetValue('GTL_PROPOSITION');
        StockTransit := StockTransit + TOBLignesDIM.Detail[i].GetValue('GTL_STOCKALF');
        ReserveCli := ReserveCli + TOBLignesDIM.Detail[i].GetValue('GTL_RESERVECLI');
        ReserveFou := ReserveFou + TOBLignesDIM.Detail[i].GetValue('GTL_RESERVEFOU');
        PrepaCli := PrepaCli + TOBLignesDIM.Detail[i].GetValue('GTL_PREPACLI');
        StockNet := (StockInitial - PrepaCli - ReserveCli + ReserveFou);
        i := i + 1;
        MoveCur(False);
      end;
      PutValue('GTL_STOCKINITIAL', StockInitial);
      PutValue('GTL_STOCKMINI', StockMini);
      PutValue('GTL_STOCKMAXI', StockMaxi);
      PutValue('GTL_QTEMIN', QteMin);
      PutValue('GTL_QTEMAX', QteMax);
      PutValue('GTL_QTEVTE', QteVte);
      PutValue('GTL_MANQUANT', Manquant);
      PutValue('GTL_SURPLUS', Surplus);
      PutValue('GTL_PROPOSITION', Proposition);
      PutValue('GTL_STOCKALF', StockTransit);
      PutValue('GTL_RESERVECLI', ReserveCli);
      PutValue('GTL_RESERVEFOU', ReserveFou);
      PutValue('GTL_PREPACLI', PrepaCli);
      AddChampSupValeur('GTL_STOCKNET', StockNet);
    end;
  end;
  FiniMove;
end;

// ----------------------------------------------------------------------------
// Retourne une quantité entière arrondie en fonction de la méthode d'arrondi :
//   Inférieur, au plus Près, Supérieur.
// ----------------------------------------------------------------------------

function ArrondirQte(MethodeArr: string; Qte: double): double;
begin
  Result := Qte;
  if MethodeArr = '' then MethodeArr := 'P';
  if MethodeArr = 'I' then Result := Trunc(Qte)
  else if MethodeArr = 'P' then Result := Arrondi(Qte, 0)
  else if MethodeArr = 'S' then Result := Arrondi(Qte + 0.4989, 0);
end;

procedure AlimenteTobMethode(CodeMethode, TypeMethode, ArrondiMeth, UtilCoef: string; QteFixe, QteMin, QteMax: double; DebPerVte, FinPerVte: string;
  TobMethode: TOB);
var TobM: TOB;
begin
  TobM := TOB.Create('LMethodes', TobMethode, -1);
  TobM.AddChampSup('ORDRE', False);
  TobM.AddChampSup('CODE', False);
  TobM.AddChampSup('LIBELLE', False);
  TobM.AddChampSup('TYPE', False);
  TobM.AddChampSup('UTILCOEF', False);
  TobM.AddChampSup('ARRONDI', False);
  TobM.InitValeurs;
  TobM.PutValue('ORDRE', TobMethode.Detail.count);
  TobM.PutValue('CODE', CodeMethode);
  if (TypeMethode = 'QTE') then
    TobM.PutValue('LIBELLE', 'Quantité fixe de ' + strF00(QteFixe, 0) + ' pièce(s) par Dimension')
  else if (TypeMethode = 'SMS') then
    TobM.PutValue('LIBELLE', 'Quantité minimum de ' + strF00(QteMin, 0) + ' pièce(s) par Dimension')
  else if (TypeMethode = 'SXS') then
    TobM.PutValue('LIBELLE', 'Quantité maximum de ' + strF00(QteMax, 0) + ' pièce(s) par Dimension')
  else if (TypeMethode = 'SMP') then
    TobM.PutValue('LIBELLE', 'Quantité en stock minimum paramètrée')
  else if (TypeMethode = 'SXP') then
    TobM.PutValue('LIBELLE', 'Quantité en stock maximum paramètrée')
  else if (TypeMethode = 'PVT') then
    TobM.PutValue('LIBELLE', 'Ventes du ' + DebPerVte + ' au ' + FinPerVte)
  else if (TypeMethode = 'PVV') then
    TobM.PutValue('LIBELLE', 'Ventes intégrées du ' + DebPerVte + ' au ' + FinPerVte)
  else if (TypeMethode = 'PDS') then
    TobM.PutValue('LIBELLE', 'Application des poids de répartition des boutiques')
  else if (TypeMethode = 'YSM') then
    TobM.PutValue('LIBELLE', 'Saisie manuelle');
  TobM.PutValue('TYPE', TypeMethode);
  TobM.PutValue('UTILCOEF', UtilCoef);
  TobM.PutValue('ARRONDI', ArrondiMeth);

  TobM.AddChampSupValeur('DATEDEBUS', USDATETIME(StrToDate(DebPerVte)));
  TobM.AddChampSupValeur('DATEFINUS', USDATETIME(StrToDate(FinPerVte)));
end;

end.

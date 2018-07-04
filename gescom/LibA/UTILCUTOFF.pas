unit UTILCUTOFF;

interface

uses
{$IFDEF VER150}
Variants,
{$ENDIF}
  Utob,
  Hctrls,
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}

  hmsgbox
  ;

type
  T_ModeEclat = (tmeRessArt, tmeArticle, tmeRessource, tmeRessTypA, tmeGlobal, tmeSans, tmeTypeArticle);
type
  T_ModeCalcul = (tmcPrest, tmcFrais, tmcFourn, tmcGlobal, tmcArticle);

type
  TFoncFormule = class

  public
    dDateF: TDateTime;
    dDateD: TDateTime;
    bAvecCutoff: boolean;
    bGenerAzero: boolean; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
    sEtablissement, sAffaire, sTiers, sArticle, sFamilleTaxe, sRegimeTaxe, sTypeCumul: string;
    iNumEclat: integer;
    TOBProd: TOB;
    TOBBM: TOB;
    TOBAff: TOB;
    TOBPiecePrepa: TOB;
    TOBActivitePrepa: TOB;
    TOBArticlePrepa: TOB;
    tModeCalc: T_ModeCalcul;

    function DonneesFormuleCutOff (AcsChampFormule: Hstring): variant;
    function GetWherePiece (stcle: string): string;
    function GetlignePrepa (TobL: TOB): Boolean;
  end;

  TTraiteCutOff = class
  public
    typeCumul: string;
    dDateD: TDateTime;
    dDateF: TDateTime;
    bAvecCutoff: boolean;
    bGenerAzero: boolean;
    bGlobal: boolean;
    sAffaire, sTiers: string;
    TobdetAff: TOB;
    iNbLignes: integer;
    bEnCompta: boolean; //AB-200602-
    procedure TraitCutOff;
    procedure TraitCutOffAchat;
    function GAChargeAffaire: boolean;
    function MajCutOff: integer; //AB-200701-Ajouté en local dans la classe
    function MajCutOffProrata: integer;
    function MajCutOffLigneArticle: integer;
    function MajCutOffTermeEchu: integer;
  end;

  TAFComptaCutOff = class
  public
    fDateCutOff: TdateTime;
  private
    fTobCutOff, fTobErreur: TOB;
    TypeCumul: string;
    procedure chargeCutoff (fTobCumAffaire: Tob);
    function GenerLesCutoffEnCompta: integer;
    procedure TraiteLignesErreur;
  end;

procedure Gener_CutOff (ZDateDeb, ZDateFin: TdateTime; bAvecCutoff: boolean = true);
procedure Gener_CutOffAchat (Global: boolean; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean = true; bGenerAzero: boolean = false);
function BoucleCutOff (Global: boolean; ZDateDebut, ZDatefin: TdateTime; TOb_Aff: Tob; bAvecCutoff: boolean = true; bGenerAzero: boolean = false): boolean;
function BoucleCutOffAchat (Global: boolean; ZDateDebut, ZDatefin: TdateTime; Tob_AFF: TOb; bAvecCutoff: boolean = true; bGenerAzero: boolean = false): boolean;
procedure ChargeTobActiviteCut (TobProd, TobBM: TOB; Aff: string; ZDateDeb, ZDateFin: TdateTime);
function EcrireDansAfCumulCut (TobProd: TOB; TobAff: Tob; TF: TFoncFormule; Mtt, MtPre, MtFra, MtFour: double): integer;
procedure StockEnrgtCutOff (TypeArt, Art, typeRess, ress: string; TF: TFoncFormule;
  Totvente, totpr, coeff, mtt: double; TobAfcumul: TOB; Tobaff: Tob);
function DetermineModeEclatCutOff: T_ModeEclat;
function DetermineModeEclatCutOffAch: T_ModeEclat;
function DetermineModeEclatFact: T_ModeEclat;
procedure EclatementCutoff (AcdSommeSolde, AcdSommeFAE, AcdSommeAAE, AcdSommePCA: double; TobReg: TOB; AcsAffaire: string; AcsTypeArticle: string);
procedure AFGenerComptaCutOff (fDateCutOff: TdateTime; fTypeCumul: string; fTobCumAffaire: Tob = nil);

//function MajCutOff (TobAff: tob; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean; bGenerAzero: boolean = false): integer;
//function MajCutOffProrata (typeCumul: string; TobAff: Tob; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean): integer;
//function MajCutOffLigneArticle (typeCumul: string; TobAff: Tob; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean): integer;
//function MajCutOffTermeEchu (typeCumul: string; TobAff: Tob; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean): integer;
const
  FC_ERR_AUX = -20;
  FC_ERR_RES = -21;
  FC_ERR_AFF = -22;
  FC_ERR_TIE = -23;
  FC_ERR_ART = -24;
  FC_ERR_ETA = -25;
const
  FC_ERR_GEN = -1;
  FC_ERR_TRAITE1 = -4;
  FC_ERR_TRAITE2 = -5;

implementation
uses EntGC,
  ParamSoc,
  Hent1,
  ent1,
  SysUtils,
  HStatus,
  lookup,
  CPProcGen,
  utilArticle,
  AffaireUtil,
  formule,
  DicoBTP,
  FactCpta,
  AFPrepfact;

const
  TexteMessage: array [1..25] of string = (
    {1}'Erreur générale à la génération des cut-off en comptabilité. Toutes les lignes n''ont pas été générées.',
    {2}'Erreur de paramétrage. Le journal n''est pas renseigné dans les paramètres.',
    {3}'Erreur de paramétrage. Le code journal saisi dans les paramètres est inconnu.',
    {4}'Erreur générale n°1 au moment de la gestion des lignes en erreur.',
    {5}'Erreur générale n°2 au moment de la gestion des lignes en erreur.',
    {6}'Génération des cut-off de vente',
    {7}'Génération des cut-off d''achat',
    {8}'Génération des cut-off en comptabilité',
    {9}'%d lignes de cut-off de vente générées au %s.',
    {10}'%d lignes de cut-off d''achat générées au %s.',
    {11}'%d lignes de cut-off générées en comptabilité et %d lignes présentent des erreurs.',
    {12}'Vous n''avez pas de cut-off à générer en comptabilité',
    {13}'La génération des cut-off de l''affaire %s n''a pu être effectuée.',
    {14}'Les cut-off de l''affaire %s sont en cours de modification par un autre utilisateur.',
    {15}'La génération des cut-off a été effectuée avec succès.',
    {16}'Certaines générations ont été arrêtées par des blocages affaires sur la génération des cut-off.',
    {17}'Un problème est survenu lors de la génération des cut-off.',
    {18}'Vous n''avez pas de cut-off à générer pour cette date %s.',
    {19}'Aucune ligne de cut-off n''a été générée en comptabilité.',
    {20}'Les cut-off de %d affaires ne sont pas générés car des cut-off sont passés en comptabilité au %s.',
    {21}'Les cut-off ne sont pas générés car des cut-off sont passés en comptabilité au %s.',
    {22}'Les cut-off de l''affaire %s ne sont pas complètement générés #10#13car des cut-off sont passés en comptabilité au %s.',
    {23}'Les cut-off de l''affaire %s ne sont pas générés #10#13car des cut-off sont passés en comptabilité au %s.',
    {24}'%d affaires ont des blocages affaires et ne sont pas générées.',
    {25}'%d affaires ont des cut-off passés en comptabilité %s et ne sont pas générées.'
    );

procedure MajCutOffTraite (TOBInter : TOB; bLignesOK : boolean);
begin
	// Rien pour BTP
end;

function  PasserCutOffEnCompta (TobEnCompta : TOB) : integer;
begin
	// Rien pour BTP
end;

procedure TTraiteCutOff.TraitCutOff;
var
  sFormule: string;
begin
  // si mise à jour par affaire, il faut détruire les enrgt concernant l'affaire
//  If not(bGlobal) then
   //AB-200602- Ne pas écraser les cut-off générés en compta
  bEnCompta := Existesql ('SELECT ACU_DATE FROM AFCUMUL WHERE ACU_TYPEAC="CVE" AND ACU_DATE="'
    + UsDAteTime (dDateF) + '" AND ACU_AFFAIRE="' + sAffaire + '" AND ACU_ERRCPTA = 1');
  if bEnCompta and (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') <> 'ADT') then
    Exit;
  // Calcul du Cut Off
  // Le calcul du Facturé prorata M6 ou M8 se fait sur les lignes de pièce
  sFormule := GetParamSocSecur ('SO_AFFORMULCUTOFF', '');
  if (pos ('M6', sFormule) > 0) and (pos ('M0', sFormule) = 0) and (pos ('M1', sFormule) = 0) then
    iNbLignes := MajCutOffProrata
  else
  if (pos ('M8', sFormule) > 0) and (pos ('M0', sFormule) = 0) and (pos ('M1', sFormule) = 0) then //GA_200809_AB-FQ15171 Facturé Proratisé avec 30 jours pour les mois entiers    
    iNbLignes := MajCutOffProrata
  else
    if (pos ('M7', sFormule) > 0) then //AB-200610- Facturé à terme échu
    iNbLignes := MajCutOffTermeEchu
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then //AB-200512- Ligne par article sans prorata
    iNbLignes := MajCutOffLigneArticle
  else
    iNbLignes := MajCutOff;
end;

procedure TTraiteCutOff.TraitCutOffAchat;
var
  TOBAfCumul, TobAfCumulDet, tobligne_CA: TOB;
  cutOff: double;
  TF: TFoncFormule;
{$IFDEF VER150}
  sFormuleCO: hString;
{$ELSE}
  sFormuleCO,
{$ENDIF}
  TypeArt, St: string;
  vResultat: variant;
  i_ind: integer;
  QQ: TQuery;
begin
  if (pos ('M6', GetParamSocSecur ('SO_AFFORMULECUTOFFACH', '[M2]-[M6]')) = 0)
  or (pos ('M8', GetParamSocSecur ('SO_AFFORMULECUTOFFACH', '[M2]-[M6]')) = 0) then //GA_200809_AB-FQ15171 Facturé Proratisé avec 30 jours pour les mois entiers
    Exit;
  //AB-200602-Ne pas modifier les cut-off générés en compta
  bEnCompta := false;
  sFormuleCO := '';
  CutOff := 0;
  TOBAfCumul := TOB.Create ('les cut-off', nil, -1);
  tobligne_CA := Tob.Create ('CA proratise', nil, -1);
  TF := TFoncFormule.Create;
  try // Analyse de la formule
    sFormuleCO := GetParamSocSecur ('SO_AFFORMULECUTOFFACH', '[M2]-[M6]');
    sFormuleCO := '{"#.###,0"' + sFormuleCO + '}';
    if TestFormule (sFormuleCO) then
    begin
      TF.sAffaire := sAffaire;
      TF.sTiers := sTiers;
      TF.sEtablissement := '';
      TF.TOBAff := nil;
      TF.dDateD := dDateD;
      TF.dDateF := dDateF;
      TF.bAvecCutoff := bAvecCutoff;
      TF.TOBProd := nil;
      TF.TOBBM := nil;
      TF.tModeCalc := tmcGlobal;
      TF.sTypeCumul := 'CAC';
      //AB-200605-FQ12748
      St := 'SELECT ACU_DATE FROM AFCUMUL WHERE ACU_TYPEAC="CAC" AND ACU_DATE="' + UsDAteTime (dDateF) + '" AND ACU_ERRCPTA=1';
      if sAffaire <> '' then
        St := St + ' AND ACU_AFFAIRE="' + sAffaire + '"';
      bEnCompta := Existesql (St);
      if bEnCompta then
        Exit;
      if bGenerAZero then
      begin
        TF.iNumEclat := 1;
        TF.bGenerAzero := true; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
        if GAChargeAffaire then
          StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobDetAff)
        else
          StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, nil);
      end
      else
      begin
        st := 'SELECT DISTINCT GL_ETABLISSEMENT,GL_AFFAIRE,GL_TIERS,GL_CODEARTICLE,GL_TYPEARTICLE,GL_REGIMETAXE,GL_FAMILLETAXE1 FROM LIGNE WHERE ' + TF.GetWherePiece ('') + ' AND ' +
          'GL_AFFAIRE not in (SELECT ACU_AFFAIRE FROM AFCUMUL WHERE ACU_TYPEAC="CAC" AND ACU_DATE="' + UsDAteTime (dDateF) + '" AND ACU_ERRCPTA=1) ' +
          'ORDER BY GL_AFFAIRE ';
        QQ := OpenSql (st, True);
        if not QQ.EOF then
          tobligne_CA.LoadDetailDB ('les lignes', '', '', QQ, True);
        Ferme (QQ);
        for i_ind := 0 to tobligne_CA.detail.count - 1 do
        begin
          TF.sEtablissement := tobligne_CA.detail [i_ind] .GetString ('GL_ETABLISSEMENT');
          sAffaire := tobligne_CA.detail [i_ind] .GetString ('GL_AFFAIRE');
          TF.sAffaire := sAffaire;
          TF.sTiers := tobligne_CA.detail [i_ind] .GetString ('GL_TIERS');
          TF.sArticle := tobligne_CA.detail [i_ind] .GetString ('GL_CODEARTICLE');
          TF.sFamilleTaxe := tobligne_CA.detail [i_ind] .GetString ('GL_FAMILLETAXE1');
          TF.sRegimeTaxe := tobligne_CA.detail [i_ind] .GetString ('GL_REGIMETAXE');
          TypeArt := tobligne_CA.detail [i_ind] .GetString ('GL_TYPEARTICLE');
          vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
          if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then //AB-20070-FQ14338
            CutOff := 0
          else
            CutOff := vResultat;
          TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
          if TobAfCumuLDet <> nil then
            while TobAfCumuLDet <> nil do
            begin
              TF.iNumEclat := TobAfCumulDet.GetInteger ('ACU_NUMECLAT') + 1;
              TobAfCumulDet := TobAfcumul.Findnext (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
            end
          else
            TF.iNumEclat := 1;
          if GAChargeAffaire then
            StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobDetAff)
          else
            StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, nil);
        end;
      end;
      iNbLignes := TobAfCumul.detail.count;
      if bGlobal then
      begin
        TF.sAffaire := '';
        TF.sTiers := '';
        TF.sArticle := '';
        TF.sEtablissement := '';
        TF.sFamilleTaxe := '';
        TF.sRegimeTaxe := '';
        St := 'DELETE from AFCUMUL where ACU_TYPEAC="CAC" AND ACU_DATE="' + UsDAteTime (dDateF) + '" AND ACU_ERRCPTA <> 1 AND ACU_AFFAIRE in ' +
          '(SELECT DISTINCT GL_AFFAIRE FROM LIGNE WHERE ' + TF.GetWherePiece ('') + ')';
      end
      else
      begin
        St := 'DELETE from AFCUMUL where ACU_TYPEAC="CAC" AND ACU_DATE="' + UsDAteTime (dDateF) + '" AND ACU_ERRCPTA <> 1';
        if sAffaire <> '' then
          St := St + ' AND ACU_AFFAIRE="' + sAffaire + '"';
        if sTiers <> '' then
          St := St + ' AND ACU_TIERS="' + sTiers + '"';
      end;
      ExecuteSql (St);
      TobAfCumul.InsertOrUpdatedb (False);
    end;
  finally
    //    if TobDetAff<> nil then TobDetAff.free; AB-20040924
    TF.Free;
    TOBAfCumul.free;
    tobligne_CA.free;
  end;
end;

function TTraiteCutOff.GAChargeAffaire: boolean;
var
  Q: TQuery;
begin
  Result := False;
  if trim (sAffaire) = '' then
    Exit;
  Q := OpenSQL ('SELECT AFF_TIERS,AFF_AFFAIRE,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AFFAIRE0,AFF_AVENANT FROM AFFAIRE WHERE AFF_AFFAIRE="' + sAffaire + '"', True);
  if not Q.EOF then
  begin
    if TobDetAff = nil then
      TobDetAff := TOB.Create ('AFFAIRE', nil, -1);
    TobDetAff.SelectDB ('', Q);
    Result := true;
  end;
  Ferme (Q);
end;

function TFoncFormule.DonneesFormuleCutOff (AcsChampFormule: Hstring): variant;
var
  sWhere, sWhereTypArt, sRefPiece: string;
  QQ: TQuery;
  dPrevu, dPlanifie, dRealise, dEcart, dCoeff, dPeriode1, dperiode2: double;
  i_ind1, i_ind2, i_numero: integer;
  tobPieces_CA, tobPiece, TobL: TOB;
  dtDEBFAC, dtFINFAC: TDateTime;
  {Retourne le nombre de jours de la période avec 30 jours pour les mois entiers}
  function NombreJoursMois30 (dtDebut,dtFin:TDateTime):double;
  var dtjour : TDateTime;
  begin
    dtjour := dtdebut;
    dtjour := FinDeMois(dtjour)+1;
    if dtdebut = DebutDeMois(dtdebut) then
      result := 30
    else
      result := dtJour - dtdebut ;
    dtJour := PlusMois (dtJour,1);
    while dtJour-1 <= dtFin do
    begin
      result := result + 30;
      dtJour := PlusMois (dtJour,1);
    end;
    if dtFin < FinDeMois(dtFin) then
    result := result + dtFin - DebutDeMois(dtFin) + 1;
  end;
begin
  // ATTENTION : CHAQUE AJOUT DE CHAMP DE L'AFFAIRE TRAITE DANS LA FORMULE IMPLIQUE DE MODIFIER LA FONCTION
  // DE SELECTION GLOBALE DES AFFAIRES DANS LA FONCTION Gener_CutOff, ET EGALEMENT LA VUE AFAFFTIERS, POUR
  // QUE LES DEUX SELECTIONS RAMENENT LES CHAMPS DE L'AFFAIRE NECESSAIRES A LA FORMULE

  Result := 1;
  if (AcsChampFormule = 'M0') then
    // Réalisé vente
  begin
    if (TOBProd <> nil) then
    begin
      if (tModeCalc = tmcGlobal) then
        Result := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE'] , [sAffaire] , False)
      else
        if (tModeCalc = tmcPrest) then
        Result := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'PRE'] , False)
      else
        if (tModeCalc = tmcFrais) then
        Result := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'FRA'] , False)
      else
        if (tModeCalc = tmcFourn) then
        Result := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'MAR'] , False)
      else
        if (tModeCalc = tmcArticle) then //AB-200512- Ligne par article sans prorata
        Result := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_CODEARTICLE'] , [sAffaire, sArticle] , False);
    end;
  end
  else
    if (AcsChampFormule = 'M1') then
    // Réalisé revient
  begin
    if (TOBProd <> nil) then
    begin
      if (tModeCalc = tmcGlobal) then
        Result := TOBProd.Somme ('ACT_TOTPR', ['ACT_AFFAIRE'] , [sAffaire] , False)
      else
        if (tModeCalc = tmcPrest) then
        Result := TOBProd.Somme ('ACT_TOTPR', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'PRE'] , False)
      else
        if (tModeCalc = tmcFrais) then
        Result := TOBProd.Somme ('ACT_TOTPR', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'FRA'] , False)
      else
        if (tModeCalc = tmcFourn) then
        Result := TOBProd.Somme ('ACT_TOTPR', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'MAR'] , False)
      else
        if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
        Result := TOBProd.Somme ('ACT_TOTPR', ['ACT_AFFAIRE', 'ACT_CODEARTICLE'] , [sAffaire, sArticle] , False);
    end;
  end
  else
    if (AcsChampFormule = 'M2') then // récupération mtt facturé
  begin
    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(GL_TOTALHT) FROM LIGNE WHERE ' + GetWherePiece (''), True);
      ////////////////////////// fin modif PL le 25/10/02
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
  {M6 	Facturé Proratisé}
    if (AcsChampFormule = 'M6') then // Calcul mtt Facturé Proratisé
  begin
    QQ := nil;
    tobPieces_CA := Tob.Create ('CA proratise', nil, -1);
    //AB-200701-FQ 13738
    sWhere := 'SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DATEDEBUTFAC,GP_DATEFINFAC,GL_TOTALHT ' +
              ' FROM PIECE,LIGNE WHERE ' + GetWherePiece ('') +
              ' AND GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO ' +
              ' ORDER BY GP_NATUREPIECEG,GP_NUMERO';
    QQ := OpenSql (sWhere, true);
    if not QQ.EOF then
      tobPieces_CA.LoadDetailDB ('_PIECE_', '', '', QQ, True);
    Ferme (QQ);
    result := 0;
    i_numero := 0;
    sRefPiece := '';
    dCoeff := 1;
    for i_ind1 := 0 to tobPieces_CA.detail.count - 1 do
    begin
      tobPiece := tobPieces_CA.detail [i_ind1] ;
      if i_numero <> tobPiece.Getinteger('GP_NUMERO') then
      begin
        i_numero := tobPiece.Getinteger('GP_NUMERO');
        dtDEBFAC := tobPiece.GetDateTime ('GP_DATEDEBUTFAC');
        dtFINFAC := tobPiece.GetDateTime ('GP_DATEFINFAC');
        { CA * (date cut-off - début facturation   / fin facturation - début facturation) }
        dPeriode1 := dDateF - dtDEBFAC + 1;
        dPeriode2 := dtFINFAC - dtDEBFAC + 1;
        dCoeff := 1;
        if (dtFINFAC <= dDateF) or (dtDEBFAC = idate1900) or (dtFINFAC = idate1900) or (dtFINFAC = idate2099) then
          dCoeff := 1
        else
          if (dtDEBFAC > dDateF) and (dtFINFAC > dDateF) then
          dCoeff := 0
        else
          if (dPeriode1 > 0) and (dPeriode2 > 0) and (dPeriode2 >= dPeriode1) and (dtFINFAC > dDateF) then
          dCoeff := dPeriode1 / dPeriode2;
      end;
      result := result + (tobPiece.GetDouble ('GL_TOTALHT') * dCoeff);
    end;
    tobPieces_CA.free;
  end
  else
  {M8 	Facturé Proratisé en mois de 30 jours}
    if (AcsChampFormule = 'M8') then //GA_200809_AB-FQ15171 Facturé Proratisé avec 30 jours pour les mois entiers
  begin
    QQ := nil;
    tobPieces_CA := Tob.Create ('CA proratise', nil, -1);
    sWhere := 'SELECT GP_NATUREPIECEG,GP_NUMERO,GP_DATEDEBUTFAC,GP_DATEFINFAC,GL_TOTALHT ' +
              ' FROM PIECE,LIGNE WHERE ' + GetWherePiece ('') +
              ' AND GP_NATUREPIECEG=GL_NATUREPIECEG AND GP_SOUCHE=GL_SOUCHE AND GP_NUMERO=GL_NUMERO ' +
              ' ORDER BY GP_NATUREPIECEG,GP_NUMERO';
    QQ := OpenSql (sWhere, true);
    if not QQ.EOF then
      tobPieces_CA.LoadDetailDB ('_PIECE_', '', '', QQ, True);
    Ferme (QQ);
    result := 0;
    i_numero := 0;
    sRefPiece := '';
    dCoeff := 1;
    for i_ind1 := 0 to tobPieces_CA.detail.count - 1 do
    begin
      tobPiece := tobPieces_CA.detail [i_ind1] ;
      if i_numero <> tobPiece.Getinteger('GP_NUMERO') then
      begin
        i_numero := tobPiece.Getinteger('GP_NUMERO');
        dtDEBFAC := tobPiece.GetDateTime ('GP_DATEDEBUTFAC');
        dtFINFAC := tobPiece.GetDateTime ('GP_DATEFINFAC');
        { CA * (date cut-off - début facturation   / fin facturation - début facturation) }
        dCoeff := 1;
        if (dtFINFAC <= dDateF) or (dtDEBFAC = idate1900) or (dtFINFAC = idate1900) or (dtFINFAC = idate2099) then
          dCoeff := 1
        else
          if (dtDEBFAC > dDateF) and (dtFINFAC > dDateF) then
          dCoeff := 0
        else
        begin
          { NombreJoursMois30 = début jour mois non entier + nombre de mois entiers * 30 + fin jour mois non entier }
          dPeriode1 :=  NombreJoursMois30 (dtDEBFAC,dDateF);
          dPeriode2 :=  NombreJoursMois30 (dtDEBFAC,dtFINFAC);
          if (dPeriode1 > 0) and (dPeriode2 > 0) and (dPeriode2 >= dPeriode1) and (dtFINFAC > dDateF) then
          dCoeff := dPeriode1 / dPeriode2;

        end;
      end;
      result := result + (tobPiece.GetDouble ('GL_TOTALHT') * dCoeff);
    end;
    tobPieces_CA.free;
  end
  else
    { Formule M7 : facturé à terme échu }
    if (AcsChampFormule = 'M7') then //AB-200610- Cutoff avec prépa facture
  begin
    Result := 0;
    for i_ind2 := 0 to TOBPiecePrepa.detail.count - 1 do
    begin
      TobL := TOBPiecePrepa.detail [i_ind2] ;
      if GetlignePrepa (TobL) then
      begin
        { Période de facturation calculée avec le mode de génération des échéances }
        dtDEBFAC := TobL.GetDateTime ('DATEDEBUTFAC');
        dtFINFAC := TobL.GetDateTime ('DATEFINFAC');
        { CA * (date cut-off - début facturation   / fin facturation - début facturation) }
        dPeriode1 := dDateF - dtDEBFAC + 1;
        dPeriode2 := dtFINFAC - dtDEBFAC + 1;
        dCoeff := 1;
        if (dtFINFAC <= dDateF) or (dtDEBFAC = idate1900) or (dtFINFAC = idate1900) or (dtFINFAC = idate2099) then
          dCoeff := 1
        else
          if (dtDEBFAC > dDateF) and (dtFINFAC > dDateF) then
          dCoeff := 0
        else
          if (dPeriode1 > 0) and (dPeriode2 > 0) and (dPeriode2 >= dPeriode1) and (dtFINFAC > dDateF) then
          dCoeff := dPeriode1 / dPeriode2;
        { Lignes de pièce générées en préparation sont proratisées sur datefac et finfac}
        result := result + (TobL.GetDouble ('GL_TOTALHT') * dCoeff);
      end;
    end;

    { Lignes d'activité reprise ne sont pas proratisées }
    for i_ind2 := 0 to TOBActivitePrepa.detail.count - 1 do
    begin
      TobL := TOBActivitePrepa.detail [i_ind2] ;
      if GetlignePrepa (TobL) then
        result := result + TobL.GetDouble ('GL_TOTALHTDEV');
    end
  end
  else
    if (AcsChampFormule = 'M3') then
  begin
    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ABU_MTPVBUD) FROM AFBUDGET WHERE ABU_TYPEAFBUDGET="PVT" AND ABU_AFFAIRE="'
        + sAffaire + '"', True);

      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'M4') then
    // TOTAH HT Global Affaire
  begin
    Result := TOBAff.GetDouble ('AFF_TOTALHTGLO');
  end
  else
    if (AcsChampFormule = 'M5') then
    // BoniMali vente
  begin
    if (TOBBM <> nil) then
    begin
      if (tModeCalc = tmcGlobal) then
        Result := TOBBM.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE'] , [sAffaire] , False)
      else
        if (tModeCalc = tmcPrest) then
        Result := TOBBM.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'PRE'] , False)
      else
        if (tModeCalc = tmcFrais) then
        Result := TOBBM.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'FRA'] , False)
      else
        if (tModeCalc = tmcFourn) then
        Result := TOBBM.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'MAR'] , False);
    end;
  end
  else
    if (AcsChampFormule = 'Q0') then
    // Somme des temps des prestations (réalisé)
  begin
    if (TOBProd <> nil) then
      Result := TOBProd.Somme ('ACT_QTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'PRE'] , False);
  end
  else
    if (AcsChampFormule = 'Q1') then // gm 05/11/02
  begin
    // Plan de charge somme qté planifiée enrgt PC
    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(APL_QTEPLANIFIEE) FROM AFPLANNING WHERE APL_TYPEPLA="PC" AND APL_AFFAIRE="'
        + sAffaire + '" AND APL_DATEDEBPLA >="' + UsDateTime (dDateF) + '" AND APL_TYPEARTICLE="PRE"'
        + ' AND APL_RESSOURCE<>"" ', True);
      if not QQ.EOF then
        result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'Q2') then
    // PC :Quantité rest à faire = (Prévu) - (Planifié depuis la date de cutoff) - (Réalisé depuis la date de cutoff) + (écart saisi)
  begin
    dPrevu := 0;
    dPlanifie := 0; {dRealise:=0 conseil}
    ;
    dEcart := 0;
    QQ := nil; // Prevu
    try
      QQ := OpenSql ('SELECT SUM(ATA_QTEINITIALE), SUM(ATA_ECARTQTEINIT) FROM TACHE WHERE ATA_AFFAIRE="'
        + sAffaire + '" AND ATA_TYPEARTICLE="PRE"', True);
      if not QQ.EOF then
      begin
        dPrevu := QQ.Fields [0] .AsFloat;
        dEcart := QQ.Fields [1] .AsFloat;
      end;
    finally
      Ferme (QQ);
    end;
    QQ := nil;
    try //gm 05/11/02 PLANIFIEE
      QQ := OpenSql ('SELECT SUM(APL_QTEPLANIFIEE) FROM AFPLANNING WHERE  APL_TYPEPLA="PC" AND APL_AFFAIRE="'
        + sAffaire + '" AND APL_DATEDEBPLA >="' + UsDateTime (dDateF) + '" AND APL_TYPEARTICLE="PRE"'
        + ' AND APL_RESSOURCE<>""', True);
      if not QQ.EOF then
        dPlanifie := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
    dRealise := 0; // Realise
    if (TOBProd <> nil) then
      dRealise := TOBProd.Somme ('ACT_QTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [sAffaire, 'PRE'] , False);
    Result := dPrevu - dPlanifie - dRealise + dEcart;
  end
  else
    if (AcsChampFormule = 'Q6') then
  begin
    // Planning jour    somme qté planifiée
    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(APL_QTEPLANIFIEE) FROM AFPLANNING WHERE APL_TYPEPLA="PLA" AND APL_AFFAIRE="'
        + sAffaire + '" AND APL_DATEDEBPLA >="' + UsDateTime (dDateF) + '" AND APL_TYPEARTICLE="PRE"'
        + ' AND APL_RESSOURCE<>"" ', True);
      if not QQ.EOF then
        result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'Q7') then
    // Planning jour :Quantité à plannifier = (Prévu) - (Planifié depuis la date de cutoff)
  begin
    dPrevu := 0;
    dPlanifie := 0;
    QQ := nil; // Prevu
    try
      QQ := OpenSql ('SELECT SUM(ATA_QTEINITPLA) FROM TACHE WHERE ATA_AFFAIRE="'
        + sAffaire + '" AND ATA_TYPEARTICLE="PRE"', True);
      if not QQ.EOF then
        dPrevu := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
    QQ := nil;
    try // PLANIFIEE
      QQ := OpenSql ('SELECT SUM(APL_QTEPLANIFIEE) FROM AFPLANNING WHERE  APL_TYPEPLA="PLA" AND APL_AFFAIRE="'
        + sAffaire + '" AND APL_DATEDEBPLA >="' + UsDateTime (dDateF) + '" AND APL_TYPEARTICLE="PRE"'
        + ' AND APL_RESSOURCE<>""', True);
      if not QQ.EOF then
        dPlanifie := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
    Result := dPrevu - dPlanifie;
  end
  else
    if (AcsChampFormule = 'VL1') then
    // Montant libre 1 Affaire
  begin
    if (TOBAff <> nil) then
      Result := TOBAff.GetDouble ('AFF_VALLIBRE1');
  end
  else
    if (AcsChampFormule = 'VL2') then
    // Montant libre 2 Affaire
  begin
    if (TOBAff <> nil) then
      Result := TOBAff.GetDouble ('AFF_VALLIBRE2');
  end
  else
    if (AcsChampFormule = 'VL3') then
    // Montant libre 3 Affaire
  begin
    if (TOBAff <> nil) then
      Result := TOBAff.GetDouble ('AFF_VALLIBRE3');
  end
  else
    if (AcsChampFormule = 'FAE') then
    // FAE dernier Cut off
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';
    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_FAE) FROM AFCUMUL WHERE ACU_TYPEAC="CVE" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;

    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'AAE') then
    // AAE dernier Cut off
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';

    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_AAE) FROM AFCUMUL WHERE ACU_TYPEAC="CVE" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'PCA') then
    // PCA dernier Cut off
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';

    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_PCA) FROM AFCUMUL WHERE ACU_TYPEAC="CVE" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'FAR') then
    // FAR dernier Cut off achat
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';

    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_FAE) FROM AFCUMUL WHERE ACU_TYPEAC="CAC" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'AAR') then
    // AAR dernier Cut off achat
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';

    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_AAE) FROM AFCUMUL WHERE ACU_TYPEAC="CAC" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    if (AcsChampFormule = 'CCA') then
    // CCA dernier Cut off achat
  begin
    if not bAvecCutoff then
    begin
      Result := 0;
      exit;
    end;
    if (tModeCalc = tmcGlobal) then
      sWhereTypArt := ''
    else
      if (tModeCalc = tmcPrest) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="PRE"'
    else
      if (tModeCalc = tmcFrais) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="FRA"'
    else
      if (tModeCalc = tmcFourn) then
      sWhereTypArt := ' AND ACU_TYPEARTICLE="MAR"'
    else
      if (tModeCalc = tmcArticle) then //AB-200512- Eclatement par article sans répartition
      sWhereTypArt := ' AND ACU_CODEARTICLE="' + sarticle + '"';

    QQ := nil;
    try
      QQ := OpenSql ('SELECT SUM(ACU_PCA) FROM AFCUMUL WHERE ACU_TYPEAC="CAC" AND ACU_AFFAIRE="'
        + sAffaire + '" AND ACU_DATE="' + UsDateTime (dDateD) + '"' + sWhereTypArt, True);
      if not QQ.EOF then
        Result := QQ.Fields [0] .AsFloat;
    finally
      Ferme (QQ);
    end;
  end
  else
    ;

end;

{ AB-200610- Cutoff avec prépa facture }
{ Retour à vrai si la Tob Ligne correspond au mode d'éclatement }

function TFoncFormule.GetlignePrepa (TobL: TOB): Boolean;
begin
  Result := false;
  if TobL.GetString ('GL_TYPELIGNE') <> 'ART' then
    Exit;
  if (tModeCalc = tmcPrest) and (TobL.GetString ('GL_TYPEARTICLE') <> 'PRE') then
    Exit
  else
    if (tModeCalc = tmcFrais) and (TobL.GetString ('GL_TYPEARTICLE') <> 'FRA') then
    Exit
  else
    if (tModeCalc = tmcFourn) and (TobL.GetString ('GL_TYPEARTICLE') <> 'MAR') then
    Exit;
  if (trim (sArticle) <> '') and (TobL.GetString ('GL_CODEARTICLE') <> sArticle) then
    Exit;
  if (trim (sFamilleTaxe) <> '') and (TobL.GetString ('GL_FAMILLETAXE1') <> sFamilleTaxe) then
    Exit;
  if (trim (sRegimeTaxe) <> '') and (TobL.GetString ('GL_REGIMETAXE') <> sRegimeTaxe) then
    Exit;
  result := true;
end;

function TFoncFormule.GetWherePiece (stcle: string): string;
var
  sWhere: string;
begin
  if trim (stcle) = '' then
  begin
    if sTypeCumul = 'CAC' then
    begin
      sWhere := '(GL_NATUREPIECEG="FF" OR GL_NATUREPIECEG="AF")' +
        ' AND GL_DATEPIECE >"' + UsDateTime (dDateD) + '" AND GL_DATEPIECE <="' + UsDateTime (dDateF) + '"';
      if sEtablissement <> '' then
        sWhere := sWhere + ' AND GL_ETABLISSEMENT="' + sEtablissement + '"';
      if sTiers <> '' then
        sWhere := sWhere + ' AND GL_TIERS="' + sTiers + '"';
      if sAffaire <> '' then
        sWhere := sWhere + ' AND GL_AFFAIRE="' + sAffaire + '"';
    end
    else
      sWhere := '(GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE" OR GL_NATUREPIECEG="AVC")' +
      ' AND GL_AFFAIRE="'+ sAffaire + '"' +
      ' AND GL_DATEPIECE >"' + UsDateTime (dDateD) + '" AND GL_DATEPIECE <="' + UsDateTime (dDateF) + '"';
  end
  else
    sWhere := stcle;
  sWhere := sWhere + ' AND GL_TYPELIGNE="ART"';
  if (tModeCalc = tmcGlobal) then
      AjoutTypeArticle ('GL', sWhere)
  else
    if (tModeCalc = tmcPrest) then
    sWhere := sWhere + ' AND GL_TYPEARTICLE="PRE" '
  else
    if (tModeCalc = tmcFrais) then
    sWhere := sWhere + ' AND GL_TYPEARTICLE="FRA" '
  else
    if (tModeCalc = tmcFourn) then
    sWhere := sWhere + ' AND GL_TYPEARTICLE="MAR" ';
  if trim (sArticle) <> '' then
    sWhere := sWhere + ' AND GL_CODEARTICLE="' + sArticle + '" ';
  if trim (sFamilleTaxe) <> '' then
    sWhere := sWhere + ' AND GL_FAMILLETAXE1="' + sFamilleTaxe + '" ';
  if trim (sRegimeTaxe) <> '' then
    sWhere := sWhere + ' AND GL_REGIMETAXE="' + sRegimeTaxe + '" ';
  result := sWhere;
end;

procedure EclatementCutoff (AcdSommeSolde, AcdSommeFAE, AcdSommeAAE, AcdSommePCA: double; TobReg: TOB; AcsAffaire: string;
  AcsTypeArticle: string);
var
  i: integer;
  dTaux, dSommeFAENew, dSommeAAENew, dSommePCANew: double;
  TOBDet: TOB;
begin

  for i := 0 to TobReg.Detail.count - 1 do
  begin
    TOBDet := TobReg.Detail [i] ;
    // Si on eclate par type article et que le type en cours n'est pas celui de la ligne,
    // on passe à la ligne suivante
    if (AcsTypeArticle <> '') and (AcsTypeArticle <> TOBDet.GetString ('ACU_TYPEARTICLE')) then
      continue;

    dTaux := 0;
    if (AcdSommeSolde <> 0) then
      dTaux := TOBDet.GetDouble ('ACU_CUTOFFORIG') / AcdSommeSolde
    else
      if (DetermineModeEclatCutOff = tmeGlobal) and (TobReg.Detail.count = 1) then
      dTaux := 1; //AB-200602- enregistrement par affaire si éclatement global

    TOBDet.PutValue ('ACU_FAE', Arrondi (AcdSommeFAE * dTaux, V_PGI.OkDecV));
    TOBDet.PutValue ('ACU_FAEDEV', Arrondi (AcdSommeFAE * dTaux, V_PGI.OkDecV));
    TOBDet.PutValue ('ACU_AAE', Arrondi (AcdSommeAAE * dTaux, V_PGI.OkDecV));
    TOBDet.PutValue ('ACU_AAEDEV', Arrondi (AcdSommeAAE * dTaux, V_PGI.OkDecV));
    TOBDet.PutValue ('ACU_PCA', Arrondi (AcdSommePCA * dTaux, V_PGI.OkDecV));
    TOBDet.PutValue ('ACU_PCADEV', Arrondi (AcdSommePCA * dTaux, V_PGI.OkDecV));
  end;

  // test de coherence des sommes et reparation sur la premiere echeance
  if (AcsTypeArticle <> '') then
  begin
    dSommeFAENew := TobReg.Somme ('ACU_FAE', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [AcsAffaire, AcsTypeArticle] , true);
    dSommeAAENew := TobReg.Somme ('ACU_AAE', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [AcsAffaire, AcsTypeArticle] , true);
    dSommePCANew := TobReg.Somme ('ACU_PCA', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [AcsAffaire, AcsTypeArticle] , true);
  end
  else
  begin
    dSommeFAENew := TobReg.Somme ('ACU_FAE', ['ACU_AFFAIRE'] , [AcsAffaire] , true);
    dSommeAAENew := TobReg.Somme ('ACU_AAE', ['ACU_AFFAIRE'] , [AcsAffaire] , true);
    dSommePCANew := TobReg.Somme ('ACU_PCA', ['ACU_AFFAIRE'] , [AcsAffaire] , true);
  end;

  if (AcdSommeFAE <> dSommeFAENew) then
  begin
    TobReg.Detail [0] .PutValue ('ACU_FAEDEV', TobReg.Detail [0] .GetDouble ('ACU_FAEDEV') - (dSommeFAENew - AcdSommeFAE));
    TobReg.Detail [0] .PutValue ('ACU_FAE', TobReg.Detail [0] .GetDouble ('ACU_FAEDEV'));
  end;
  if (AcdSommeAAE <> dSommeAAENew) then
  begin
    TobReg.Detail [0] .PutValue ('ACU_AAEDEV', TobReg.Detail [0] .GetDouble ('ACU_AAEDEV') - (dSommeAAENew - AcdSommeAAE));
    TobReg.Detail [0] .PutValue ('ACU_AAE', TobReg.Detail [0] .GetDouble ('ACU_AAEDEV'));
  end;
  if (AcdSommePCA <> dSommePCANew) then
  begin
    TobReg.Detail [0] .PutValue ('ACU_PCADEV', TobReg.Detail [0] .GetDouble ('ACU_PCADEV') - (dSommePCANew - AcdSommePCA));
    TobReg.Detail [0] .PutValue ('ACU_PCA', TobReg.Detail [0] .GetDouble ('ACU_PCADEV'));
  end;

end;

function DetermineModeEclatCutOff: T_ModeEclat;
begin
  Result := tmeSans;
  if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'AAR') then
    Result := tmeRessArt
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ART') then
    Result := tmeArticle
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ASS') then
    Result := tmeRessource
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ATY') then
    Result := tmeRessTypA
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'GLO') then
    Result := tmeGlobal
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'TYP') then
    Result := tmeTypeArticle
  else
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then
    Result := tmeArticle; //AB-200512- Ligne par article sans prorata

end;

{ Prend en  compte ce mode d'éclatement en modification détaillée des cut-off des achats }

function DetermineModeEclatCutOffAch: T_ModeEclat;
var
  StModeEclat: string;
begin
  Result := tmeSans;
  StModeEclat := GetParamSocSecur ('SO_AFCUTMODEECLATACH', '');
  if StModeEclat = '' then
    StModeEclat := GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO');
  if StModeEclat = 'AAR' then
    Result := tmeRessArt
  else
    if StModeEclat = 'ART' then
    Result := tmeArticle
  else
    if StModeEclat = 'ASS' then
    Result := tmeRessource
  else
    if StModeEclat = 'ATY' then
    Result := tmeRessTypA
  else
    if StModeEclat = 'GLO' then
    Result := tmeGlobal
  else
    if StModeEclat = 'TYP' then
    Result := tmeTypeArticle
  else
    if StModeEclat = 'ADT' then
    Result := tmeArticle;
end;

function DetermineModeEclatFact: T_ModeEclat;
begin
  Result := tmeSans;
  if (GetParamSocSecur ('SO_AFFACTPARRES', 'SAN') = 'COD') then
    Result := tmeRessArt
  else
    if (GetParamSocSecur ('SO_AFFACTPARRES', 'SAN') = 'GLO') then
    Result := tmeRessource
  else
    if (GetParamSocSecur ('SO_AFFACTPARRES', 'SAN') = 'TYP') then
    Result := tmeRessTypA
  else
end;

procedure Gener_CutOffAchat (Global: boolean; ZDateDebut, ZDateFin: TdateTime; bAvecCutoff: boolean = true; bGenerAzero: boolean = false);
var
  tob_aff :tob;
  QQ :Tquery;
begin
  Tob_AFF := Tob.Create ('Liste Affaire', nil, -1);
  try
    QQ := OpenSql ('SELECT AFF_AFFAIRE,AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_ETATAFFAIRE,'
      + 'AFF_AVENANT,AFF_ETABLISSEMENT,AFF_TIERS,AFF_VALLIBRE1,AFF_VALLIBRE2,AFF_VALLIBRE3,AFF_TOTALHTGLO,'
      + 'AFF_TOTALHTGLODEV FROM AFFAIRE WHERE AFF_MODELE<>"X" AND AFF_ADMINISTRATIF<>"X" AND AFF_STATUTAFFAIRE="AFF"'
      + ' AND AFF_TYPEAFFAIRE="NOR" AND AFF_ETATAFFAIRE="ENC"', True);  //AB-200708- affaire encours

    if not QQ.EOF then
      Tob_Aff.loadDetailDb ('Detail', '', '', QQ, False);
    //AB-200708-génération totale
    BoucleCutOffAchat (false, ZDateDebut, ZDateFin, Tob_Aff, bAvecCutOff, GetParamSocSecur ('SO_AFALIMCUTOFFACH', False));//AB-200708- FQ14469

  finally
    Tob_Aff.Free;
  end;
end;

procedure Gener_CutOff (ZDateDeb, ZDateFin: TdateTime; bAvecCutoff: boolean = true);
var
  Tob_Aff: TOB;
  QQ: TQuery;
  repGener, bArrete, bGenerAzero: boolean;
begin
  Tob_Aff := nil;
  QQ := nil;
  // on détruit les enrgt afcumul à meme date
  //ExecuteSql ('DELETE from AFCUMUL where ACU_TYPEAC="CVE" AND ACU_DATE="' + UsDAteTime(Zdate)+'"');

  // on charge toutes les affaires pour traitement
  // A FAIRE : voir si il faut faire une sélection sur ces affaires ??? il y en a qui ne sont plus valide ....
  // affaire non fermées : GM?
  // ==> non car on peut générer un cut-off sur des affaires qui viennent d'être fermées
  // exclure les affaires "administrative"
  try
    Tob_AFF := Tob.Create ('Liste Affaire', nil, -1);
    try
      // ICI ON DOIT AJOUTER LES CHAMPS NECESSAIRES A LA FORMULE TRAITEE PLUS LOIN
      QQ := OpenSql ('SELECT AFF_AFFAIRE,AFF_AFFAIRE0,AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_ETATAFFAIRE,'
        + 'AFF_AVENANT,AFF_ETABLISSEMENT,AFF_TIERS,AFF_VALLIBRE1,AFF_VALLIBRE2,AFF_VALLIBRE3,AFF_TOTALHTGLO,'
        + 'AFF_TOTALHTGLODEV, T_REGIMETVA FROM AFFAIRE ' +
        'INNER JOIN TIERS ON (AFF_TIERS = T_TIERS) ' + // BDU - 07/11/07 - FQ : 13138. Ajout de la jointure sur TIERS pour avoir le régime TVA
        'WHERE AFF_MODELE<>"X" AND AFF_ADMINISTRATIF<>"X" AND AFF_STATUTAFFAIRE="AFF"'
        + ' AND AFF_TYPEAFFAIRE="NOR" AND AFF_ETATAFFAIRE="ENC"', True);  //AB-200708- affaire encours

      if not QQ.EOF then
        Tob_Aff.loadDetailDb ('Detail', '', '', QQ, False);

    finally
      Ferme (QQ);
    end;
    bGenerAzero := GetParamSocSecur ('SO_AFALIMCUTOFF', False);//AB-200708- FQ14469
    repGener := BoucleCutOff (True, ZDateDeb, ZDateFin, Tob_Aff, bAvecCutoff,bGenerAzero);

  finally
    Tob_Aff.Free;
  end;

  // Arrêté de période en liaison
  if (repGener = true) and (ctxScot in V_PGI.PGIContexte) then
    // on est en cutoff général
    // Dans Scot uniquement pour l'instant
  begin
    bArrete := true;
    if (GetParamSocSecur ('SO_AFCLIENT', 0) = cInClientAmyot) and not (GetParamSocSecur ('SO_AFALIMCUTOFF', False) = false) then
      // Pour Amyot Seulement si on alimente automatiquement les montants : PL le 29/10/03 suite demande Conan ce jour
      bArrete := false;

    // PL le 04/11/02 : on ouvre la saisie d'activité à partir du lendemain de l'arrêté de période
    // SetParamSoc('SO_AFDATEDEBUTACT', ZDate);
    if bArrete then
      SetParamSoc ('SO_AFDATEDEBUTACT', ZDateFin + 1);
    ///////////////////////
  end;

end;

// fct qui boucle sur toutes les Affaires passée dans la TOB
// et qui pour chacune des affaire fait la mise à jour du cut off.
// Deux points d'entrée différents :
//      - Boucle sur toutes les affaires valides
//      - Boucle sur une sélection d'affaires

function BoucleCutOff (Global: boolean; ZDateDebut, ZDatefin: TdateTime; Tob_AFF: TOb; bAvecCutoff: boolean = true; bGenerAzero: boolean = false): boolean;
var
  TobDetAff: Tob;
  wi, iNbLignes, iEnCompta, iBlocage: integer;
  Rep: TIOErr;
  TCO: TTraiteCutOff;
  bProbleme, bBlocageAff: boolean;
  RepBlocage: T_TypeBlocAff;
  dDateDebItv, dDateFinItv: TDateTime;
begin
  Result := false;
  bProbleme := false;
  bBlocageAff := false;
  iNbLignes := 0;
  iEnCompta := 0;
  iBlocage := 0;
  try
    InitMove (TOb_Aff.detail.count, '');
    // PL A FAIRE : voir si on peut afficher un message pour indiquer quelle est l'affaire en cours de traitement
    for wi := 0 to TOb_Aff.detail.count - 1 do
    begin
      TCO := nil;
      tobDetAff := Tob_Aff.detail [wi] ;
      try
        TCO := TTraiteCutOff.Create;
        TCO.typecumul := 'CVE';
        TCO.dDateD := ZDateDebut;
        TCO.dDateF := ZDatefin;
        TCO.bAvecCutoff := bAvecCutoff;
        TCO.bGenerAZero := bGenerAZero; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
        TCO.bGlobal := Global;
        TCO.sAffaire := TobDetAff.GetString ('AFF_AFFAIRE');
        TCO.TobdetAff := tobDetAff;
        MoveCur (False);
        RepBlocage := BlocageAffaire ('CUT', TobDetAff.GetString ('AFF_AFFAIRE'), V_PGI.groupe,
          ZDatefin, false, false, false, dDateDebItv, dDateFinItv, nil);

        Rep := oeOk;
        if (RepBlocage = tbaAucun) then
          Rep := Transactions (TCO.TraitCutOff, 1)
        else
        begin
          bBlocageAff := true;
          inc (iBlocage);
        end;

        if not Global then
        begin
          case Rep of
            oeOk:
              begin
                iNbLignes := iNbLignes + TCO.iNbLignes;
                //AB-200602-Ne pas modifier les cut-off générés en compta
                if TCO.bEnCompta then
                begin
                  inc (iEnCompta);
                  if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT' then
                    PgiBoxAF (TexteMessage [22] , TexteMessage [6] , [TCO.sAffaire, dateToStr (ZDatefin)] )
                    {22'Les cut-off de l''affaire %s ne sont pas complètement générés #10#13car des cut-off sont passés en comptabilité au %s.',}
                  else
                    PgiBoxAF (TexteMessage [23] , TexteMessage [6] , [TCO.sAffaire, dateToStr (ZDatefin)] )
                    {23'Les cut-off de l''affaire %s ne sont pas générés #10#13car des cut-off sont passés en comptabilité au %s.'}
                end;
              end;
            oeUnknown:
              begin
                PGIBox (format (traduitGA (TexteMessage [13] ), [TCO.sAffaire] ), TraduitGA (TexteMessage [6] ));
                {13 'La génération des cut-off de l''affaire %s n''a pu être effectuée.',}
                bProbleme := true;
                Continue;
              end;
            oeSaisie:
              begin
                PGIBox (format (traduitGA (TexteMessage [14] ), [TCO.sAffaire] ), TraduitGA (TexteMessage [6] ));
                {14 'Les cut-off de l''affaire %s sont en cours de modification par un autre utilisateur.',}
                bProbleme := true;
                Continue;
              end;
          end;
        end
        else
        begin
          if (Rep = oeOk) then
            iNbLignes := iNbLignes + TCO.iNbLignes
          else
            if (Rep = oeUnknown) or (Rep = oeSaisie) then
          begin
            bProbleme := true;
            continue;
          end;
          if TCO.bEnCompta then
            inc (iEnCompta);
        end;

      finally
        TCO.Free;
      end;
    end; // fin For

  finally
    FiniMove;
  end;

  if (bProbleme = true) then
    PGIInfoAF (TexteMessage [17] , TexteMessage [6] )
      {17 'Un problème est survenu lors de la génération des cut-off.',}
  else
    PGIBox (format (traduitGA (TexteMessage [9] ), [iNbLignes, dateToStr (ZDatefin)] ), TraduitGA (TexteMessage [6] ));
  {9 '%d lignes de cut-off de ventes générées au %s.',}

  if bBlocageAff then
    PgiBoxAF (TexteMessage [24] , TexteMessage [6] , [iBlocage] );
  {24'%d affaires ont des blocages affaires et ne sont pas générées.'}

  if iEnCompta > 0 then
    PgiBoxAF (TexteMessage [25] , TexteMessage [6] , [iEnCompta, dateToStr (ZDatefin)] );
  {25'%d affaires ont des cut-off passés en comptabilité %s et ne sont pas générées.'}

// mise à jour de la date de dernier cut off dans les paramètres
  if not bProbleme then
  begin
    SetParamsoc ('SO_DATECUTOFF', ZDateFin);
    Result := true;
  end;
end;

//AB-200512- Ligne par article sans prorata

function TTraiteCutOff.MajCutOffLigneArticle: integer;
var
  TOBAfCumul, TobAfCumulDet, Tobligne_CA, TobProd, TobBM, TobArticles, TobA, TOBAfCumulCpta: TOB;
  cutOff: double;
  TF: TFoncFormule;
{$IFDEF VER150}
  sFormuleCO: hString;
{$ELSE}
  sFormuleCO,
{$ENDIF}
  TypeArt: string;
  vResultat: variant;
  i_ind: integer;
  QQ: TQuery;
begin
  sFormuleCO := '';
  Result := 0;
  TOBAfCumul := TOB.Create ('les cut-off', nil, -1);
  Tobligne_CA := Tob.Create ('CA proratise', nil, -1);
  TobProd := Tob.Create ('la production', nil, -1);
  TobArticles := Tob.Create ('les articles', nil, -1);
  TobBM := Tob.Create ('les boniMalis', nil, -1);
  TOBAfCumulCpta := Tob.Create ('En Compta', nil, -1);
  TF := TFoncFormule.Create;
  try // Analyse de la formule
    sFormuleCO := GetParamSocSecur ('SO_AFFORMULCUTOFF', '');
    sFormuleCO := '{"#.###,0"' + sFormuleCO + '}';
    ChargeTobActiviteCut (TobProd, TobBM, TobdetAff.GetString ('AFF_AFFAIRE'), dDateD, dDateF);

    if TestFormule (sFormuleCO) then
    begin
      if typeCumul = 'CVE' then
      begin
        TF.sAffaire := TobdetAff.GetString ('AFF_AFFAIRE');
        TF.sTiers := TobdetAff.GetString ('AFF_TIERS');
        TF.TOBAff := TobdetAff;
        TF.TOBProd := TobProd;
      end
      else
      begin
        Exit;
        TF.sAffaire := '';
        TF.TOBAff := nil;
        TF.TOBProd := nil;
      end;
      TF.dDateD := dDateD;
      TF.dDateF := dDateF;
      TF.bAvecCutoff := bAvecCutoff;
      TF.bGenerAzero := bGenerAzero; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
      TF.TOBBM := nil;
      TF.tModeCalc := tmcArticle;
      TF.sTypeCumul := typeCumul;
      //AB-200602-Ne pas écraser les cut-off générés en compta
      QQ := OpenSql ('SELECT * FROM AFCUMUL WHERE ACU_TYPEAC="CVE" AND ACU_AFFAIRE="' + TF.sAffaire + '" AND ACU_DATE="' + UsDateTime (TF.dDateF) + '" AND ACU_ERRCPTA <> 0', true);
      if not QQ.EOF then
      begin
        TOBAfCumulCpta.LoadDetailDB ('en compta', '', '', QQ, True);
      end;
      Ferme (QQ);

      QQ := OpenSql ('SELECT DISTINCT GL_CODEARTICLE,GL_TYPEARTICLE FROM LIGNE WHERE ' + TF.GetWherePiece (''), True);
      if not QQ.EOF then
        tobligne_CA.LoadDetailDB ('les lignes', '', '', QQ, True);
      Ferme (QQ);

      for i_ind := 0 to tobligne_CA.detail.count - 1 do
      begin
        TobA := Tob.Create ('un article', TobArticles, -1);
        TobA.AddChampSupValeur ('CODEARTICLE', tobligne_CA.detail [i_ind] .GetString ('GL_CODEARTICLE'), False);
        TobA.AddChampSupValeur ('TYPEARTICLE', tobligne_CA.detail [i_ind] .GetString ('GL_TYPEARTICLE'), False);
      end;
      for i_ind := 0 to tobProd.detail.count - 1 do
      begin
        if TobArticles.FindFirst (['CODEARTICLE'] , [tobProd.detail [i_ind] .GetString ('ACT_CODEARTICLE')] , False) <> nil then
          Continue;
        TobA := Tob.Create ('un article', TobArticles, -1);
        TobA.AddChampSupValeur ('CODEARTICLE', tobProd.detail [i_ind] .GetString ('ACT_CODEARTICLE'), False);
        TobA.AddChampSupValeur ('TYPEARTICLE', tobProd.detail [i_ind] .GetString ('ACT_TYPEARTICLE'), False);
      end;
      for i_ind := 0 to TobArticles.detail.count - 1 do
      begin
        TF.sArticle := TobArticles.detail [i_ind] .GetString ('CODEARTICLE');
        TypeArt := TobArticles.detail [i_ind] .GetString ('TYPEARTICLE');
        vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
        if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then //AB-20070-FQ14338
          CutOff := 0
        else
          CutOff := vResultat;
        TobAfCumulDet := TOBAfCumulCpta.Findfirst (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
        if TobAfCumuLDet <> nil then
          Continue; //AB-200602-Ne pas écraser les cut-off générés en compta
        TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
        if TobAfCumuLDet <> nil then
        begin
          while TobAfCumuLDet <> nil do
          begin
            TF.iNumEclat := TobAfCumulDet.GetInteger ('ACU_NUMECLAT') + 1;
            TobAfCumulDet := TobAfcumul.Findnext (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
          end
        end
        else
          TF.iNumEclat := 1;
        if (typeCumul = 'CVE') and
        ((pos ('M2', sFormuleCO) > 0) and (pos ('M6', sFormuleCO) > 0) and (pos ('M6', sFormuleCO) > pos ('M2', sFormuleCO))) or
        ((pos ('M0', sFormuleCO) > 0) and (pos ('M2', sFormuleCO) > 0) and (pos ('M0', sFormuleCO) > pos ('M2', sFormuleCO))) or
        ((pos ('M2', sFormuleCO) > 0) and (pos ('M8', sFormuleCO) > 0) and (pos ('M8', sFormuleCO) > pos ('M2', sFormuleCO))) //GA_200809_AB-FQ15171 Facturé Proratisé avec 30 jours pour les mois entiers
        then //GA_200802_AB-GA14934
          CutOff := CutOff * (-1);
        StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobdetAff);
      end;

      Result := TobAfCumul.detail.count;
      //AB-200602- Ne pas écraser les cut-off générés en compta
      ExecuteSql ('DELETE from AFCUMUL where ACU_TYPEAC="CVE" AND ACU_DATE="'
        + UsDAteTime (dDateF) + '" AND ACU_AFFAIRE="' + TobdetAff.GetString ('AFF_AFFAIRE') + '" AND ACU_ERRCPTA <> 1');
      TobAfCumul.InsertOrUpdatedb (False);
      // mise à jour date dernier cutoff dans affaire
      ExecuteSql ('UPDATE AFFAIRE SET AFF_DATECUTOFF="' + UsDateTime (dDateF)
        + '" WHERE AFF_AFFAIRE="'
        + TobdetAff.GetString ('AFF_AFFAIRE')
        + '" AND AFF_DATECUTOFF<"' + UsDateTime (dDateF) + '"');
    end;
  finally
    TF.Free;
    TOBAfCumul.free;
    tobligne_CA.free;
    TobBM.Free;
    TobProd.free;
    TobArticles.free;
    TOBAfCumulCpta.free;
  end;
end;

function TTraiteCutOff.MajCutOff: integer;
var
  TobProd, TobBM: TOB;
  cutOff, coPres, coFra, coFour: double;
  TF: TFoncFormule;
{$IFDEF VER150}
  sFormuleCO: hString;
{$ELSE}
  sFormuleCO: string;
{$ENDIF}
  vResultat, vResPre, vResFra, vResFour: variant;
begin
  sFormuleCO := '';
  Result := 0;
  // fct qui charge activite en fct paramétrage de génération
  TobProd := Tob.Create ('la production', nil, -1);
  TobBM := Tob.Create ('les boniMalis', nil, -1);
  TF := TFoncFormule.Create;
  try
    ChargeTobActiviteCut (TobProd, TobBM, TobdetAff.GetString ('AFF_AFFAIRE'), dDateD, dDateF);
    ///////////////////////////////////////////////// PL 14/08/02
    // Analyse de la formule
    sFormuleCO := GetParamSocSecur ('SO_AFFORMULCUTOFF', '');
    sFormuleCO := '{"#.###,0"' + sFormuleCO + '}';
    if TestFormule (sFormuleCO) then
    begin
      TF.sAffaire := TobdetAff.GetString ('AFF_AFFAIRE');
      TF.TOBAff := TobdetAff;
      TF.dDateD := dDateD;
      TF.dDateF := dDateF;
      TF.bAvecCutoff := bAvecCutoff;
      TF.bGenerAzero := bGenerAzero; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
      TF.TOBProd := TOBProd;
      TF.TOBBM := TobBM;
      TF.tModeCalc := tmcGlobal;
      TF.iNumEclat := 1;
      TF.sTypeCumul := 'CVE';
      vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
      if (GetParamSocSecur ('SO_AFCLIENT', 0) = cInClientAmyot) then
      begin
        TF.tModeCalc := tmcPrest;
        vResPre := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
        TF.tModeCalc := tmcFrais;
        vResFra := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
        TF.tModeCalc := tmcFourn;
        vResFour := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
      end;

      // PL le 10/12/03 : cas d'erreur en sortie de la fonction GFormule renvoie de '' si le montant est très petit 1,13687E-13
      //          if (VarType(vResultat) = varString) and (vResultat = 'Erreur') then
      if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then //AB-20070-FQ14338
        CutOff := 0
      else
      begin
        CutOff := vResultat;
        if (pos ('M0', sFormuleCO) > 0) and (pos ('M2', sFormuleCO) > 0) and (pos ('M0', sFormuleCO) > pos ('M2', sFormuleCO)) then
          CutOff := CutOff * (-1); //AB-200701-FQ12155
      end;

      //          if (VarType(vResPre) = varString) and (vResPre = 'Erreur')  then
      if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResPre = 'Erreur') or (vResPre = '')) then
        coPres := 0
      else
        coPres := vResPre;

      //          if (VarType(vResFra) = varString) and (vResFra = 'Erreur')then
      if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResFra = 'Erreur') or (vResFra = '')) then
        coFra := 0
      else
        coFra := vResFra;

      //          if (VarType(vResFour) = varString) and (vResFour = 'Erreur')  then
      if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResFour = 'Erreur') or (vResFour = '')) then
        coFour := 0
      else
        coFour := vResFour;
      ///////////////////////////////////////////////// PL 14/08/02

  // PL le 06/05/03 : on veut toutes les affaires, même si le résultat de la formule est =0
  //    If (CutOff <> 0) then
  //      begin

      Result := EcrireDansAfCumulCut (TobProd, TobdetAff, TF, CutOff, coPres, coFra, coFour);

      // mise à jour date dernier cutoff dans affaire
      ExecuteSql ('UPDATE AFFAIRE SET AFF_DATECUTOFF="' + UsDateTime (dDateF)
        + '" WHERE AFF_AFFAIRE="'
        + TobdetAff.GetString ('AFF_AFFAIRE')
        + '" AND AFF_DATECUTOFF<"' + UsDateTime (dDateF) + '"');
    end;
  finally
    TobBM.Free;
    TobProd.free;
    TF.free;
  end;
end;

procedure ChargeTobActiviteCut (TobProd, TobBM: TOB; aff: string; ZDateDeb, ZDateFin: TdateTime);
var
  Req, Req1, zwhere, wheredate, zorder: string;
  QQ: Tquery;
  TobDet: TOB;
  dDebut, dfin: TdateTime;
  CumTotPre, CumTotFra, CumTotMar, CumTot, Coeff: Double;
  wi: integer;
begin
  CumTotPre := 0;
  CumTotFra := 0;
  CumTotMar := 0;
  // fct qui charge l'activité sur la période voulue et en fct de
  // paramètres de la base, afin d'avoir la base pour cut off
      // sélect zones en fct option d'éclatement
  // A FAIRE : tester tous les # cas d'éclatement
  Req1 := 'SELECT SUM(ACT_TOTVENTE) AS ACT_TOTVENTE,SUM(ACT_TOTPR) AS ACT_TOTPR,SUM(ACT_QTE) AS ACT_QTE,ACT_AFFAIRE';

  if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ART') or (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then
    req1 := req1 + ',ACT_CODEARTICLE,ACT_TYPEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ASS' then
    req1 := req1 + ',ACT_TYPERESSOURCE,ACT_RESSOURCE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'AAR' then
    req1 := req1 + ',ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_CODEARTICLE,ACT_TYPEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ATY' then
    req1 := req1 + ',ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_TYPEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'TYP' then
    req1 := req1 + ',ACT_TYPEARTICLE'
      ;
  //else  req1 := req1 + ',ACT_TYPEARTICLE';   // gm 05/11/02  pl?
  req1 := req1 + ' FROM ACTIVITE ';

  dDebut := ZDateDeb;
  dfin := ZDateFin;
  zwhere := ' WHERE  act_typeactivite="REA" AND ACT_AFFAIRE="' + Aff + '"';
  zwhere := zwhere + ' AND ACT_ACTIVITEREPRIS <> "N"'; //AB-200512- Eclatement par article
  AjoutTypeArticle ('ACT', zwhere);
  wheredate := ' AND ACT_DATEACTIVITE > "' + usdatetime (dDebut) + '"'; // PL le 22/04/03 : dorenavant, on part du dernier cut off
  wheredate := wheredate + ' AND ACT_DATEACTIVITE <= "' + usdatetime (dfin) + '"'; // PL le 22/04/03
  wheredate := wheredate + '  AND ACT_ETATVISA="VIS"';

  // group by  en fct option d'éclatement
  zorder := '  group by act_affaire';
  //AB-200512- Ligne par article sans prorata
  if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ART') or (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then
    zorder := zorder + ',ACT_TYPEARTICLE,ACT_CODEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ASS' then
    zorder := zorder + ',ACT_TYPERESSOURCE,ACT_RESSOURCE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'AAR' then
    zorder := zorder + ',ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_TYPEARTICLE,ACT_CODEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ATY' then
    zorder := zorder + ',ACT_TYPERESSOURCE,ACT_RESSOURCE,ACT_TYPEARTICLE'
  else
    if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'TYP' then
    zorder := zorder + ',ACT_TYPEARTICLE'
      ;
  //else zorder := zorder + ',ACT_TYPEARTICLE';  // gm 05/11/02  pl?

  Req := Req1 + zwhere + wheredate + zorder;

  // Le Realise
  QQ := nil;
  try
    QQ := OpenSQL (Req, true);
    if not QQ.EOF then
      TobProd.LoadDetailDB ('_ACTIVITE_', '', '', QQ, True);

  finally
    Ferme (QQ);
  end;

  // Les boni/Malis
  Req := StringReplace (Req, '"REA"', '"BON"', [rfReplaceAll] );
  QQ := nil;
  try
    QQ := OpenSQL (Req, true);
    if not QQ.EOF then
      TobBM.LoadDetailDB ('_ACTIVITE_', '', '', QQ, True);
  finally
    Ferme (QQ);
  end;

  if TobProd.detail.count = 0 then
    exit;

  // calcul du coeff pour chaque ligne en fct paramétrage
  CumTot := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE'] , [aff] , False);
  // PL le 19/05/03 : spécif Amyot : éclatement par coeff
  if (GetParamSocSecur ('SO_AFCLIENT', 0) = cInClientAmyot) then
  begin
    CumTotPre := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [aff, 'PRE'] , False);
    CumTotFra := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [aff, 'FRA'] , False);
    CumTotMar := TOBProd.Somme ('ACT_TOTVENTE', ['ACT_AFFAIRE', 'ACT_TYPEARTICLE'] , [aff, 'MAR'] , False);
  end;
  // fin PL le 19/05/03

  for wi := 0 to TObProd.detail.count - 1 do
  begin
    TobDet := TobProd.detail [wi] ;
    TobDet.AddChampSup ('COEFF', False); // ajout champ coeff sur enrgt
    // gm 05/11/02 , Si pas d'éclatement le coef est inutile le mettre à 1
    if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'GLO') then
    begin
      Coeff := 1;
    end
    else
    begin
      Coeff := 0;
      // PL le 19/05/03 : spécif Amyot : éclatement par coeff
      if (GetParamSocSecur ('SO_AFCLIENT', 0) = cInClientAmyot) then
      begin
        if (TObDet.GetString ('ACT_TYPEARTICLE') = 'PRE') and (CumTotPre <> 0) then
          coeff := TObDet.GetDouble ('ACT_TOTVENTE') / CumTotPre
        else
          if (TObDet.GetString ('ACT_TYPEARTICLE') = 'FRA') and (CumTotFra <> 0) then
          coeff := TObDet.GetDouble ('ACT_TOTVENTE') / CumTotFra
        else
          if (TObDet.GetString ('ACT_TYPEARTICLE') = 'MAR') and (CumTotMar <> 0) then
          coeff := TObDet.GetDouble ('ACT_TOTVENTE') / CumTotMar;
      end
      else
        if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT' then //AB-200512- Ligne par article sans prorata
        Coeff := 1
      else
      begin
        // PL le 22/04/03 : on ne doit pas arrondir le coefficient : problèmes d'arrondis, on vérifie l'arrondi de la somme sur les montants et non pas sur le coefficient
        if CumTot <> 0 then
          coeff := TObDet.GetDouble ('ACT_TOTVENTE') / CumTot;
        // fin PL le 22/04/03
      end;
      ////////////// : fin PL le 19/05/03
    end;

    TobDet.PutValue ('COEFF', Coeff);
  end;
end;

function EcrireDansAfCumulCut (TobProd: TOB; TobAff: Tob; TF: TFoncFormule; Mtt, MtPre, MtFra, MtFour: double): integer;
var
  TObAfCumul, TobProddet, TobAfcumulDet, TOBAffaires: TOB;
  wi: integer;
  Ress, typeArt, Art, TypeRess: string;
  CutOff, Coeff, totvente, Totpr, tot1: double;
begin
  TOBAfCumul := TOB.Create ('AFCUMUL', nil, -1);
  TOBAffaires := TOB.Create ('Les affaires', nil, -1);
  // fct qui pour le mtt passé,  alimente la table
  // afcumul en fct du paramétrage et du coeff calculé dans la prod
  try
    if TobProd.detail.count > 0 then
    begin
      for wi := 0 to TobProd.detail.count - 1 do
      begin
        TobProdDet := TobProd.detail [wi] ;
        Ress := '';
        Art := '';
        TypeArt := '';
        TypeRess := '';
        // on charge les valeurs par défaut de afcumul en fct du paramétrage
        // (les zones n'existent pas toutes dans la tob)
        if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ART') or (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then //AB-200512- Ligne par article sans prorata
        begin
          art := TobProdDet.GetString ('ACT_CODEARTICLE');
          typeart := TobProdDet.GetString ('ACT_TYPEARTICLE');
        end
        else
          if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'TYP' then
        begin
          typeart := TobProdDet.GetString ('ACT_TYPEARTICLE');
        end
        else
          if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'AAR' then
        begin
          typeart := TobProdDet.GetString ('ACT_TYPEARTICLE');
          art := TobProdDet.GetString ('ACT_CODEARTICLE');
          Ress := TobProdDet.GetString ('ACT_RESSOURCE');
          TypeRess := TobProdDet.GetString ('ACT_TYPERESSOURCE');
        end
        else
          if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ASS' then
        begin
          Ress := TobProdDet.GetString ('ACT_RESSOURCE');
          TypeRess := TobProdDet.GetString ('ACT_TYPERESSOURCE');
          typeart := 'PRE'; // sinon l'alimentation de aftableaubord ne marche pas. ne traite que les types prestation valides ..
        end
        else
          if GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ATY' then
        begin
          typeart := TobProdDet.GetString ('ACT_TYPEARTICLE');
          Ress := TobProdDet.GetString ('ACT_RESSOURCE');
          TypeRess := TobProdDet.GetString ('ACT_TYPERESSOURCE');
        end
        else
          typeart := 'PRE'; // sinon l'alimentation de aftableaubord ne marche pas. ne traite que les types prestation valides ..

        CutOff := Mtt;
        if (GetParamSocSecur ('SO_AFCLIENT', 0) = cInClientAmyot) then
        begin
          if (typeart = 'PRE') then
            CutOff := MtPre
          else
            if (typeart = 'FRA') then
            CutOff := MtFra
          else
            if (typeart = 'MAR') then
            CutOff := MtFour;
        end;

        Coeff := TobProdDet.GetDouble ('COEFF');
        Totvente := TobProdDet.GetDouble ('ACT_TOTVENTE');
        TotPr := TobProdDet.GetDouble ('ACT_TOTPR');
        StockEnrgtCutOff (TypeArt, Art, typeRess, ress, TF, Totvente, totpr, coeff, CutOff, TobAfcumul, TobAff);
      end;
    end
    else
    begin
      // cas ou on a un mtt de cutoff , sans activité correspondante, il faut quand même écrire enrgt
      // voir si prévoir des codes par défaut dans les paramètres, actuellement tout mis sur des codes à blanc
      Ress := '';
      Art := '';
      TypeArt := 'PRE';
      TypeRess := '';
      if (GetParamSocSecur ('SO_AFCLIENT', 0) <> cInClientAmyot) then
        StockEnrgtCutOff (TypeArt, Art, typeRess, ress, TF, 0, 0, 1, Mtt, TobAfCumul, TobAff)
      else
      begin
        StockEnrgtCutOff ('PRE', Art, typeRess, ress, TF, 0, 0, 1, MtPre, TobAfCumul, TobAff);
        StockEnrgtCutOff ('FRA', Art, typeRess, ress, TF, 0, 0, 1, MtFra, TobAfCumul, TobAff);
        StockEnrgtCutOff ('MAR', Art, typeRess, ress, TF, 0, 0, 1, MtFour, TobAfCumul, TobAff);
      end;
    end;

    if mtt <> 0 then
    begin
      if (GetParamSocSecur ('SO_AFCLIENT', 0) <> cInClientAmyot) then
      begin
        tot1 := TOBAfCumul.Somme ('ACU_CUTOFFORIG', ['ACU_AFFAIRE'] , [TobAff.GetString ('AFF_AFFAIRE')] , False);
        if tot1 <> Arrondi (mtt, V_PGI.OkDecV) then
        begin
          TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE'] , [TobAff.GetString ('AFF_AFFAIRE')] , True);
          totvente := mtt - tot1;
          if TobAfCumuLDet <> nil then
          begin
            TobAfCumulDet.putvalue ('ACU_CUTOFFORIG', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIG') + totvente);
            TobAfCumulDet.putvalue ('ACU_CUTOFFORIGDEV', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIGDEV') + totvente);
          end;
        end;
      end
      else
        // PL le 19/05/03 : spécif Amyot : éclatement par coeff donc arrondis par coeff
      begin
        // Les prestations
        tot1 := TOBAfCumul.Somme ('ACU_CUTOFFORIG', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'PRE'] , False);
        TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'PRE'] , True);
        if (TobAfCumulDet <> nil) then
        begin
          //              Coeff := valeur (TobAfCumulDet.GetDouble ('COEFF'));
          //              if (tot1 <> Arrondi (MtPre * Coeff, V_PGI.OkDecV)) then
          if (tot1 <> Arrondi (MtPre, V_PGI.OkDecV)) then
          begin
            //                  totvente := (MtPre * Coeff) - tot1;
            totvente := MtPre - tot1;
            if TobAfCumuLDet <> nil then
            begin
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIG', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIG') + totvente);
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIGDEV', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIGDEV') + totvente);
            end;
          end;
        end;
        // Les frais
        tot1 := TOBAfCumul.Somme ('ACU_CUTOFFORIG', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'FRA'] , False);
        TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'FRA'] , True);
        if (TobAfCumulDet <> nil) then
        begin
          //              Coeff := valeur (TobAfCumulDet.GetDouble ('COEFF'));
          //              if (tot1 <> Arrondi (MtFra * Coeff, V_PGI.OkDecV)) then
          if (tot1 <> Arrondi (MtFra, V_PGI.OkDecV)) then
          begin
            //                  totvente := (MtFra * Coeff) - tot1;
            totvente := MtFra - tot1;
            if TobAfCumuLDet <> nil then
            begin
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIG', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIG') + totvente);
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIGDEV', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIGDEV') + totvente);
            end;
          end;
        end;
        // Les fournitures
        tot1 := TOBAfCumul.Somme ('ACU_CUTOFFORIG', ['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'MAR'] , False);
        TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TYPEARTICLE'] , [TobAff.GetString ('AFF_AFFAIRE'), 'MAR'] , True);
        if (TobAfCumulDet <> nil) then
        begin
          //              Coeff := valeur (TobAfCumulDet.GetDouble ('COEFF'));
          //              if (tot1 <> Arrondi (MtFour * Coeff, V_PGI.OkDecV)) then
          if (tot1 <> Arrondi (MtFour, V_PGI.OkDecV)) then
          begin
            //                  totvente := (MtFour * Coeff) - tot1;
            totvente := MtFour - tot1;
            if TobAfCumuLDet <> nil then
            begin
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIG', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIG') + totvente);
              TobAfCumulDet.putvalue ('ACU_CUTOFFORIGDEV', TobAfCumulDet.GetDouble ('ACU_CUTOFFORIGDEV') + totvente);
            end;
          end;
        end;
      end;
    end;

    if not TF.bGenerAzero then //AB-200708-FQ14469
      // Si on alimente automatiquement les cut off et que l'affaire en cours n'est pas terminée
    begin
      // on regarde si mtt éclaté = mtt  origine. sinon mis écart sur 1er
      if mtt > 0 then
        tot1 := TOBAfCumul.Somme ('ACU_FAE', ['ACU_AFFAIRE'] , [TobAff.GetString ('AFF_AFFAIRE')] , False)
      else
      begin
        // PL le 24/10/02 : modif de dernière minute suite remarques AMYOT
//          Tot1:=TOBAfCumul.Somme('ACU_AAE',['ACU_AFFAIRE'],[TobAff.GetString('AFF_AFFAIRE')],False);
        tot1 := TOBAfCumul.Somme ('ACU_PCA', ['ACU_AFFAIRE'] , [TobAff.GetString ('AFF_AFFAIRE')] , False);
        // fin modif PL le 24/10/02
        tot1 := tot1 * (-1);
      end;

      if Tot1 <> Arrondi (mtt, V_PGI.OkDecV) then
      begin
        TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE'] , [TobAff.GetString ('AFF_AFFAIRE')] , True);
        totvente := mtt - tot1;
        if TobAfCumuLDet <> nil then
        begin
          if mtt > 0 then
          begin
            TobAfCumulDet.putvalue ('ACU_FAE', TobAfCumulDet.GetDouble ('ACU_FAE') + totvente);
          end
          else
          begin
            // PL le 24/10/02 : modif de dernière minute suite remarques AMYOT
            TobAfCumulDet.putvalue ('ACU_PCA', TobAfCumulDet.GetDouble ('ACU_PCA') + (totvente * -1));
            // fin modif PL le 24/10/02
          end;
        end;
      end;
    end;
    Result := TobAfCumul.detail.count;
    ExecuteSql ('DELETE from AFCUMUL where ACU_TYPEAC="CVE" AND ACU_DATE="'
      + UsDAteTime (TF.dDateF) + '" AND ACU_AFFAIRE="' + TobAff.GetString ('AFF_AFFAIRE') + '"'); //AB-200602-
    TobAfCumul.InsertOrUpdatedb (False);

  finally
    TobAfCumul.free;
    TOBAffaires.free;
  end;
end;

procedure StockEnrgtCutOff (TypeArt, Art, typeRess, ress: string; TF: TFoncFormule;
  Totvente, totpr, coeff, mtt: double; TobAfcumul: TOB; Tobaff: Tob);
var
  TobAfCumulDet: TOB;
begin
  //Si non cocher 'Génération des cut-off à zéro' et cocher 'Ignorer les valeurs nulles des calculs générés'
  if (TF.sTypeCumul = 'CVE') and not TF.bGenerAzero and not GetParamSocSecur ('SO_AFALIMCUTOFF', False)
    and GetParamSocSecur ('SO_AFCUTOFFNULS', False) and (mtt = 0) then
    Exit; //Ne pas générer des lignes de cut-off à zéro
  if (TF.sTypeCumul = 'CAC') and not TF.bGenerAzero and not GetParamSocSecur ('SO_AFALIMCUTOFFACH', False)
    and GetParamSocSecur ('SO_AFCUTOFFNULSACH', False) and (mtt = 0) then
    Exit; //Ne pas générer des lignes de cut-off à zéro
  TobAfCumulDet := Tob.create ('AFCUMUL', TobAfCumul, -1);
  TobAfCumulDet.InitValeurs;
  TobAfCumulDet.AddChampSup ('COEFF', False); // ajout champ coeff sur enrgt
  TobAfCumulDet.putvalue ('ACU_TYPEAC', TF.sTypeCumul);
  TobAfCumulDet.putvalue ('ACU_NUMECLAT', TF.iNumEclat);
  TobAfCumulDet.putvalue ('ACU_CODEARTICLE', Art);
  if trim (art) <> '' then
    TobAfCumulDet.putvalue ('ACU_ARTICLE', CodeArticleUnique2 (Art, ''));
  TobAfCumulDet.putvalue ('ACU_TYPEARTICLE', Typeart);
  TobAfCumulDet.putvalue ('ACU_RESSOURCE', Ress);
  TobAfCumulDet.putvalue ('ACU_ETABLISSEMENT', TF.sEtablissement);
  TobAfCumulDet.putvalue ('ACU_TIERS', TF.sTiers);
  if TobAff <> nil then
  begin
    TobAfCumulDet.putvalue ('ACU_AFFAIRE', TobAff.GetString ('AFF_AFFAIRE'));
    TobAfCumulDet.putvalue ('ACU_AFFAIRE1', TobAff.GetString ('AFF_AFFAIRE1'));
    TobAfCumulDet.putvalue ('ACU_AFFAIRE2', TobAff.GetString ('AFF_AFFAIRE2'));
    TobAfCumulDet.putvalue ('ACU_AFFAIRE3', TobAff.GetString ('AFF_AFFAIRE3'));
    TobAfCumulDet.putvalue ('ACU_AFFAIRE0', TobAff.GetString ('AFF_AFFAIRE0'));
    TobAfCumulDet.putvalue ('ACU_AVENANT', TobAff.GetString ('AFF_AVENANT'));
    if TF.sTiers = '' then
      TobAfCumulDet.putvalue ('ACU_TIERS', TobAff.GetString ('AFF_TIERS'));
    if TF.sEtablissement = '' then
      TobAfCumulDet.putvalue ('ACU_ETABLISSEMENT', TobAff.GetString ('AFF_ETABLISSEMENT'));
    // BDU - 07/11/07 - FQ : 13138. Charge le régime TVA
    if TF.sRegimeTaxe = '' then
    begin
      TobAfCumulDet.PutValue('ACU_REGIMETAXE', TobAff.GetString('T_REGIMETVA'));
      TF.sRegimeTaxe  := TobAff.GetString('T_REGIMETVA');
    end;
  end;
  TobAfCumulDet.putvalue ('ACU_DATE', TF.dDateF);
  TobAfCumulDet.putvalue ('ACU_SEMAINE', NumSemaine (TF.dDateF));
  TobAfCumulDet.putvalue ('ACU_PERIODE', GetPeriode (TF.dDateF));
  TobAfCumulDet.putvalue ('ACU_TYPERESSOURCE', TypeRess);
  // cumul zone pour enrgt à mettre à jour
  TobAfCumulDet.putvalue ('ACU_PVPROD', TotVente);
  TobAfCumulDet.putvalue ('ACU_PVPRODDEV', TotVente);
  TobAfCumulDet.putvalue ('ACU_PRPROD', TotPr);
  TobAfCumulDet.putvalue ('ACU_PRPRODDEV', TotPr);
  // Alimenter ACU_CUTOFFORIG par mtt * coeff
  TobAfCumulDet.putvalue ('COEFF', Coeff);
  TobAfCumulDet.putvalue ('ACU_CUTOFFORIG', Arrondi (Coeff * mtt, V_PGI.OkDecV));
  TobAfCumulDet.putvalue ('ACU_CUTOFFORIGDEV', Arrondi (Coeff * mtt, V_PGI.OkDecV));
  TobAfCumulDet.putvalue ('ACU_DATEMODIF', V_PGI.DateEntree);

  // if (GetParamSocSecur ('SO_AFALIMCUTOFF', False) = true) then
  if not TF.bGenerAzero then //AB-200701-FQ13401-Générer les cut offs des ventes à 0
  begin
    if TF.sTypeCumul = 'CAC' then //AB-20040929 On inverse pour les cut-off d'achats
    begin
      if Arrondi (Coeff * mtt, V_PGI.OkDecV) > 0 then
      begin
        TobAfCumulDet.putvalue ('ACU_PCA', Arrondi (Coeff * mtt, V_PGI.OkDecV));
        TobAfCumulDet.putvalue ('ACU_PCADEV', Arrondi (Coeff * mtt, V_PGI.OkDecV));
      end
      else
      begin
        TobAfCumulDet.putvalue ('ACU_FAE', Arrondi (Coeff * mtt, V_PGI.OkDecV) * -1);
        TobAfCumulDet.putvalue ('ACU_FAEDEV', Arrondi (Coeff * mtt, V_PGI.OkDecV) * -1);
      end;
    end
    else
    begin
      // Si on alimente automatiquement les cut off de ventes
      if Arrondi (Coeff * mtt, V_PGI.OkDecV) > 0 then
      begin
        TobAfCumulDet.putvalue ('ACU_FAE', Arrondi (Coeff * mtt, V_PGI.OkDecV));
        TobAfCumulDet.putvalue ('ACU_FAEDEV', Arrondi (Coeff * mtt, V_PGI.OkDecV));
      end
      else
      begin
        // PL le 24/10/02 : modif de dernière minute suite remarques AMYOT
        TobAfCumulDet.putvalue ('ACU_PCA', Arrondi (Coeff * mtt, V_PGI.OkDecV) * -1);
        TobAfCumulDet.putvalue ('ACU_PCADEV', Arrondi (Coeff * mtt, V_PGI.OkDecV) * -1);
        // Fin modif PL du 24/10/02
      end;
    end;
    TobAfCumulDet.putvalue ('ACU_FAMILLETAXE', TF.sFamilleTaxe);
    TobAfCumulDet.putvalue ('ACU_REGIMETAXE', TF.sRegimeTaxe);
  end;
end;

function BoucleCutOffAchat (Global: boolean; ZDateDebut, ZDatefin: TdateTime; Tob_AFF: TOb; bAvecCutoff: boolean = true; bGenerAzero: boolean = false): boolean;
var
  TobDetAff: Tob;
  wi, iEnCompta: integer;
  Rep: TIOErr;
  TCO: TTraiteCutOff;
  bProbleme, bBlocageAff: boolean;
  iNbLignes: Integer;
begin
  Result := false;
  bProbleme := false;
  bBlocageAff := false;
  iNbLignes := 0;
  iEnCompta := 0;
  try
    InitMove (TOb_Aff.detail.count, '');
    for wi := 0 to TOb_Aff.detail.count - 1 do
    begin
      TCO := nil;
      tobDetAff := Tob_Aff.detail [wi] ;
      try
        TCO := TTraiteCutOff.Create;
        TCO.dDateD := ZDateDebut;
        TCO.dDateF := ZDatefin;
        TCO.bAvecCutoff := bAvecCutoff;
        TCO.bGenerAZero := bGenerAZero;
        TCO.bGlobal := Global;
        TCO.sAffaire := TobDetAff.GetString ('AFF_AFFAIRE');
        TCO.sTiers := '';
        TCO.TobdetAff := tobDetAff;
        MoveCur (False);
        Rep := Transactions (TCO.TraitCutOffAchat, 1);
        if not Global then
        begin
          case Rep of
            oeOk:
              begin
                iNbLignes := iNbLignes + TCO.iNbLignes;
                if TCO.bEnCompta then //AB-200602-Ne pas modifier les cut-off générés en compta
                  inc (iEnCompta);
              end;
            oeUnknown:
              begin
                PGIBox (format (traduitGA (TexteMessage [13] ), [TCO.sAffaire] ), TraduitGA (TexteMessage [7] ));
                {13 'La génération des cut-off de l''affaire %s n''a pu être effectuée.',}
                bProbleme := true;
                Continue;
              end;
            oeSaisie:
              begin
                PGIBox (format (traduitGA (TexteMessage [14] ), [TCO.sAffaire] ), TraduitGA (TexteMessage [7] ));
                {14 'Les cut-off de l''affaire %s sont en cours de modification par un autre utilisateur.',}
                bProbleme := true;
                Continue;
              end;
          end;
        end
        else
        begin
          if (Rep = oeOk) then
          begin
            iNbLignes := iNbLignes + TCO.iNbLignes;
            if TCO.bEnCompta then
              inc (iEnCompta);
          end
          else
            if (Rep = oeUnknown) or (Rep = oeSaisie) then
          begin
            bProbleme := true;
            continue;
          end;
        end;

      finally
        TCO.Free;
      end;
    end; // fin For

  finally
    FiniMove;
  end;

  if (bProbleme = true) then
    PGIInfoAF (TexteMessage [17] , TexteMessage [7] )
      {17 'Un problème est survenu lors de la génération des cut-off.',}
  else
    if iEnCompta > 0 then //AB-200602-Ne pas modifier les cut-off générés en compta
    PgiBoxAF (TexteMessage [20] , TexteMessage [7] , [iEnCompta, dateToStr (ZDatefin)] )
    {20 'Les cut-off de %d affaires ne sont pas générés car des cut-off sont passés en comptabilité au %s.'}
  else
    if (bBlocageAff = false) then
    PGIBox (format (traduitGA (TexteMessage [10] ), [iNbLignes, dateToStr (ZDatefin)] ), TraduitGA (TexteMessage [7] ))
    {10 '%d lignes de cut-off d''achats générées au %s.',}
  else
    PgiInfoAf (TexteMessage [16] , TexteMessage [7] );
  {16 'Certaines générations ont été arrêtées par des blocages affaires sur la génération des cut-off.',}
// mise à jour de la date de dernier cut off dans les paramètres
  if not bProbleme then
  begin
    SetParamsoc ('SO_DATECUTOFFACH', ZDateFin);
    Result := true;
  end;
end;

{ M6 Facturé Proratisé}
{ M8 Facturé Proratisé en mois de 30 jours}
function TTraiteCutOff.MajCutOffProrata: integer;
var
  TOBAfCumul, TobAfCumulDet, tobligne_CA: TOB;
  cutOff: double;
  TF: TFoncFormule;
{$IFDEF VER150}
  sFormuleCO: hString;
{$ELSE}
  sFormuleCO,
{$ENDIF}
  TypeArt: string;
  vResultat: variant;
  i_ind: integer;
  QQ: TQuery;
begin
  sFormuleCO := '';
  Result := 0;
  TOBAfCumul := TOB.Create ('les cut-off', nil, -1);
  tobligne_CA := Tob.Create ('CA proratise', nil, -1);
  TF := TFoncFormule.Create;
  try // Analyse de la formule
    sFormuleCO := GetParamSocSecur ('SO_AFFORMULCUTOFF', '');
    sFormuleCO := '{"#.###,0"' + sFormuleCO + '}';
    if TestFormule (sFormuleCO) then
    begin
      if typeCumul = 'CVE' then
      begin
        TF.sAffaire := TobdetAff.GetString ('AFF_AFFAIRE');
        TF.sTiers := TobdetAff.GetString ('AFF_TIERS');
        TF.TOBAff := TobdetAff;
        TF.TOBProd := nil;
      end
      else
      begin
        TF.sAffaire := '';
        TF.TOBAff := nil;
        TF.TOBProd := nil;
      end;
      TF.dDateD := dDateD;
      TF.dDateF := dDateF;
      TF.bAvecCutoff := bAvecCutoff;
      TF.bGenerAzero := bGenerAzero; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
      TF.TOBBM := nil;
      TF.tModeCalc := tmcGlobal;
      TF.sTypeCumul := typeCumul;
      if (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ART') or (GetParamSocSecur ('SO_AFCUTMODEECLAT', 'GLO') = 'ADT') then
      begin
        QQ := OpenSql ('SELECT DISTINCT GL_ETABLISSEMENT,GL_CODEARTICLE,GL_TYPEARTICLE,GL_REGIMETAXE,GL_FAMILLETAXE1 FROM LIGNE WHERE ' + TF.GetWherePiece (''), True);
        if not QQ.EOF then
          tobligne_CA.LoadDetailDB ('les lignes', '', '', QQ, True);
        Ferme (QQ);
        for i_ind := 0 to tobligne_CA.detail.count - 1 do
        begin
          TF.sEtablissement := tobligne_CA.detail [i_ind] .GetString ('GL_ETABLISSEMENT');
          TF.sArticle := tobligne_CA.detail [i_ind] .GetString ('GL_CODEARTICLE');
          TF.sFamilleTaxe := tobligne_CA.detail [i_ind] .GetString ('GL_FAMILLETAXE1');
          TF.sRegimeTaxe := tobligne_CA.detail [i_ind] .GetString ('GL_REGIMETAXE');
          TypeArt := tobligne_CA.detail [i_ind] .GetString ('GL_TYPEARTICLE');
          vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
          if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then //AB-20070-FQ14338
            CutOff := 0
          else
            CutOff := vResultat;
          TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
          if TobAfCumuLDet <> nil then
            while TobAfCumuLDet <> nil do
            begin
              TF.iNumEclat := TobAfCumulDet.GetInteger ('ACU_NUMECLAT') + 1;
              TobAfCumulDet := TobAfcumul.Findnext (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
            end
          else
            TF.iNumEclat := 1;
          if (typeCumul = 'CVE') and (pos ('M2', sFormuleCO) > 0)
          and ((pos ('M6', sFormuleCO) > pos ('M2', sFormuleCO)) or (pos ('M8', sFormuleCO) > pos ('M2', sFormuleCO))) then
            CutOff := CutOff * (-1);
          StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobdetAff);
        end;
      end
      else
      begin
        vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
        if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then
          CutOff := 0
        else
        begin
          CutOff := vResultat;
          if (typeCumul = 'CVE') and (pos ('M2', sFormuleCO) > 0)
          and ((pos ('M6', sFormuleCO) > pos ('M2', sFormuleCO)) or  (pos ('M8', sFormuleCO) > pos ('M2', sFormuleCO)))then
            CutOff := CutOff * (-1);
        end;
        StockEnrgtCutOff ('', '', '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobdetAff);
      end;
      Result := TobAfCumul.detail.count;
      //AB-200602-
      ExecuteSql ('DELETE from AFCUMUL where ACU_TYPEAC="CVE" AND ACU_DATE="'
        + UsDAteTime (dDateF) + '" AND ACU_AFFAIRE="' + TobdetAff.GetString ('AFF_AFFAIRE') + '"');
      TobAfCumul.InsertOrUpdatedb (False);
      // mise à jour date dernier cutoff dans affaire
      ExecuteSql ('UPDATE AFFAIRE SET AFF_DATECUTOFF="' + UsDateTime (dDateF)
        + '" WHERE AFF_AFFAIRE="'
        + TobdetAff.GetString ('AFF_AFFAIRE')
        + '" AND AFF_DATECUTOFF<"' + UsDateTime (dDateF) + '"');
    end;
  finally
    TF.Free;
    TOBAfCumul.free;
    tobligne_CA.free;
  end;
end;

{ M7 Facturé à terme échu}
{ Lancement d'une préparation de facture sans enregistrer en base sur les échéances non-facturées}
{ Le résultat sera le montant des échéances à facturer et le montant de l'activité à facturer saisies dans la période }

function TTraiteCutOff.MajCutOffTermeEchu: integer;
var
  TOBAfCumul, TobAfCumuLDet: TOB;
  cutOff: double;
  TF: TFoncFormule;
{$IFDEF VER150}
  sFormuleCo: hString;
{$ELSE}
  sFormuleCo: string;
{$ENDIF}
  TypeArt, stSql: string;
  vResultat: variant;
  i_ind2: integer;
  TOBEcheances, TOBArticles: tob;
  TOBPiecePrepa, TOBActivitePrepa, TobL: tob;
begin
  sFormuleCO := '';
  Result := 0;
  cutOff := 0;
  TOBAfCumul := TOB.Create ('les cut-off', nil, -1);
  TOBPiecePrepa := TOB.Create ('les Pieces', nil, -1);
  TOBEcheances := Tob.Create ('les Echéances', nil, -1);
  TOBArticles := TOB.Create ('les ARTICLES', nil, -1);
  TOBActivitePrepa := TOB.Create ('les activitées', nil, -1);
  stSql := 'SELECT AFF_AFFAIRE, AFF_AFFAIRE1, AFF_AFFAIRE2, AFF_AFFAIRE3, AFF_AVENANT, AFF_LIBELLE, T_TIERS,'
    + 'AFF_REGROUPEFACT, AFF_PRINCIPALE, AFF_PROFILGENER, AFA_AFFAIRE, AFA_NUMECHE, AFF_AFFAIREREF,'
    + 'AFF_ISAFFAIREREF, AFA_TYPECHE, AFA_DATEECHE, AFA_GENERAUTO, AFA_MONTANTECHE, AFF_REGSURCAF,'
    + 'AFA_NUMECHEBIS, AFF_NATUREPIECEG  '
    + 'FROM AFACTTIERSAFFAIRE '
    + 'WHERE AFA_DATEECHE >= "' + USDateTime (dDateD) + '" AND AFA_DATEECHE < "' + USDateTime (dDateF) + '" AND AFA_ECHEFACT<>"X" '
    + 'AND AFA_LIQUIDATIVE<>"X" AND AFF_MODELE<>"X" AND AFA_TYPECHE = "NOR" AND AFF_REGSURCAF<>"X" '
    + 'AND AFF_AFFAIRE = "' + TobdetAff.GetString ('AFF_AFFAIRE') + '" ORDER BY AFA_NUMECHE DESC';

  TOBEcheances.LoadDetailDBFromSQL ('Echéances', stSql);
  { Lancement d'une préparation de facture sans enregistrer en base sur les échéances non-facturées }
  AFPrepFact_SansMajBase (TOBEcheances, TOBArticles, TOBPiecePrepa, TOBActivitePrepa, datetostr (dDateD), datetostr (dDateF));

  TF := TFoncFormule.Create;
  try // Analyse de la formule
    sFormuleCO := GetParamSocSecur ('SO_AFFORMULCUTOFF', '');
    sFormuleCO := '{"#.###,0"' + sFormuleCO + '}';
    if TestFormule (sFormuleCO) then
    begin
      TF.sAffaire := TobdetAff.GetString ('AFF_AFFAIRE');
      TF.sTiers := TobdetAff.GetString ('AFF_TIERS');
      TF.TOBAff := TobdetAff;
      TF.TOBProd := nil;
      TF.TOBPiecePrepa := TOBPiecePrepa;
      TF.TOBActivitePrepa := TOBActivitePrepa;
      TF.TOBArticlePrepa := TOBArticles;
      TF.dDateD := dDateD;
      TF.dDateF := dDateF;
      TF.bAvecCutoff := bAvecCutoff;
      TF.bGenerAzero := bGenerAzero; //AB-200701-FQ13401-Générer les cut offs des ventes à 0
      TF.TOBBM := nil;
      TF.sTypeCumul := typeCumul;
      if DetermineModeEclatCutOff = tmeArticle then
      begin
        TF.tModeCalc := tmcArticle;
        for i_ind2 := 0 to TOBArticles.Detail.Count - 1 do
        begin
          TobL := TOBPiecePrepa.FindFirst (['GL_ARTICLE'] , [TOBArticles.detail [i_ind2] .GetString ('GA_ARTICLE')] , False);
          if (TobL = nil) and (TOBActivitePrepa.detail.count > 0) then
            TobL := TOBActivitePrepa.FindFirst (['GL_ARTICLE'] , [TOBArticles.detail [i_ind2] .GetString ('GA_ARTICLE')] , False);
          if TobL <> nil then
          begin
            TF.sEtablissement := TobL.GetString ('GL_ETABLISSEMENT');
            TF.sArticle := TobL.GetString ('GL_CODEARTICLE');
            TF.sFamilleTaxe := TobL.GetString ('GL_FAMILLETAXE1');
            TF.sRegimeTaxe := TobL.GetString ('GL_REGIMETAXE');
            TypeArt := TobL.GetString ('GL_TYPEARTICLE');
            vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
            if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then //AB-20070-FQ14338
              CutOff := 0
            else
              CutOff := vResultat;
            TobAfCumulDet := TobAfcumul.Findfirst (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
            if TobAfCumuLDet <> nil then
              while TobAfCumuLDet <> nil do
              begin
                TF.iNumEclat := TobAfCumulDet.GetInteger ('ACU_NUMECLAT') + 1;
                TobAfCumulDet := TobAfcumul.Findnext (['ACU_AFFAIRE', 'ACU_TIERS', 'ACU_CODEARTICLE'] , [TF.sAffaire, TF.sTiers, TF.sArticle] , True);
              end
            else
              TF.iNumEclat := 1;
            StockEnrgtCutOff (TypeArt, TF.sArticle, '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobdetAff);
          end;
        end
      end
      else
      begin
        TF.tModeCalc := tmcGlobal;
        vResultat := GFormule (sFormuleCO, TF.DonneesFormuleCutOff, nil, 1);
        if ((VarType (vResultat) = varString) or (VarType (vResultat) = varOleStr)) and ((vResultat = 'Erreur') or (vResultat = '')) then
          CutOff := CutOff + 0
        else
          CutOff := CutOff + vResultat;
      end;
      if DetermineModeEclatCutOff <> tmeArticle then
        StockEnrgtCutOff ('', '', '', '', TF, 0, 0, 1, CutOff, TobAfcumul, TobdetAff);
      Result := TobAfCumul.detail.count;
      //AB-200602-
      ExecuteSql ('DELETE from AFCUMUL where ACU_TYPEAC="CVE" AND ACU_DATE="'
        + UsDAteTime (dDateF) + '" AND ACU_AFFAIRE="' + TobdetAff.GetString ('AFF_AFFAIRE') + '"');
      TobAfCumul.InsertOrUpdatedb (False);
      // mise à jour date dernier cutoff dans affaire
      ExecuteSql ('UPDATE AFFAIRE SET AFF_DATECUTOFF="' + UsDateTime (dDateF)
        + '" WHERE AFF_AFFAIRE="'
        + TobdetAff.GetString ('AFF_AFFAIRE')
        + '" AND AFF_DATECUTOFF<"' + UsDateTime (dDateF) + '"');
    end;
  finally
    TF.Free;
    TOBAfCumul.free;
    TOBPiecePrepa.free;
    TOBEcheances.free;
    TOBArticles.free;
    TOBActivitePrepa.free;
  end;
end;

{*****************************************************************************
Auteur  ...... : AB
Créé le ...... : 06/2004
Modifié le ... :   /  /
Description .. : Génération des CutOff en Comptablité
Suite ........ : fDateCutOff : Date du cutoff
Suite ........ : fTypeCumul=CVE pour les ventes
Suite ........ : fTypeCumul=CAC pour les achats
Suite ........ : Appel de fonction d'écriture comptable de FactCpta.pas
Suite ........ : NbLignesOK := PasserCutOffEnCompta (TobEnCompta);
*******************************************************************************}

procedure AFGenerComptaCutOff (fDateCutOff: TdateTime; fTypeCumul: string; fTobCumAffaire: Tob = nil);
var
  TCPTA: TAFComptaCutOff;
begin
  TCPTA := TAFComptaCutOff.Create;
  TCPTA.fDateCutOff := fDateCutOff;
  TCPTA.TypeCumul := fTypeCumul;
  try
    TCPTA.chargeCutoff (fTobCumAffaire);
    if (TCPTA.fTobCutOff <> nil) and (TCPTA.fTobCutOff.detail.count > 0) then
      TCPTA.GenerLesCutoffEnCompta;
  finally
    TCPTA.Free;
  end
end;

{*** Chargement pour la génération Totale et fTobCumAffaire pour la génération avec sélection ************}

procedure TAFComptaCutOff.chargeCutoff (fTobCumAffaire: Tob);
var
  Q: TQuery;
  st, stWhere: string;
begin
  stWhere := '';
  if fTobCumAffaire = nil then
  begin
    fTobCutOff := TOB.Create ('cut-off', nil, -1);
    st := 'SELECT ACU_TYPEAC,ACU_DATE,ACU_TIERS,ACU_AFFAIRE,ACU_NUMECLAT,ACU_RESSOURCE,ACU_TYPEARTICLE,ACU_CODEARTICLE,' +
      ' ACU_ETABLISSEMENT,ACU_ARTICLE,ACU_REGIMETAXE,ACU_FAMILLETAXE,ACU_FAE,ACU_AAE,ACU_PCA,' +
      ' AFF_ETABLISSEMENT,AFF_LIBELLE,T_AUXILIAIRE,T_LIBELLE' +
      ' FROM AFCUMUL' +
      ' LEFT JOIN AFFAIRE ON ACU_AFFAIRE=AFF_AFFAIRE' +
      ' LEFT JOIN TIERS ON ACU_TIERS=T_TIERS' +
      ' WHERE ACU_DATE="' + UsDateTime (fDateCutOff) + '" AND ACU_TYPEAC="' + TypeCumul + '"' +
      ' AND ACU_ERRCPTA <> 1 AND (ACU_FAE<>0 OR ACU_AAE<>0 OR ACU_PCA<>0) ' + stWhere +
      ' ORDER BY ACU_ETABLISSEMENT,ACU_AFFAIRE,T_AUXILIAIRE,ACU_CODEARTICLE';
    Q := OpenSQL (st, True);
    fTobCutOff.LoadDetailDB ('Lignes CutOff', '', '', Q, True);
    Ferme (Q);
  end
  else
    fTobCutOff := fTobCumAffaire;
  if (fTobCutOff.detail.count = 0) then
    PGIInfoAF (TexteMessage [12] , TexteMessage [8] );
  {12'Vous n''avez pas de cut-off à générer en comptabilité',}{8'Génération des cut-off en comptabilité',}
end;


function TAFComptaCutOff.GenerLesCutoffEnCompta: integer;
var
  ii: integer;
  RepGen, sDateGen: string;
  repGener, CodeErreurAppel, NbLignesOK: integer;
  TobEnCompta, TOBMCpta, TOBLCpta, TOBMere, TobEnErreur, TOBL, TOBTVA: TOB;
  Etablissement, RegimeTaxe, FamilleTaxe, LibEcriture, RefExterne, RefInterne: string;
  MtHT, MtTTC, MtTVA, Taux: currency;
  MtHTFAE, MtHTFAR, MtHTAAE, MtHTAAR, MtHTPCA, MtHTCCA: currency;
  io1: TIoErr;
begin
  TobEnCompta := nil;
  TobEnErreur := nil;
  repGener := 0;
  io1 := oeOk;
  sDateGen := '';
  RepGen := '';
  try
    try
      TobEnCompta := TOB.Create ('CutOffDansCompta', nil, -1);
      TobEnErreur := TOB.Create ('CutOffEnErreur', nil, -1);
      for ii := 0 to fTobCutOff.Detail.Count - 1 do
      begin
        CodeErreurAppel := 0;
        MtHT := 0;
        MtTTC := 0;
        MtTVA := 0;
        MtHTFAE := 0;
        MtHTAAE := 0;
        MtHTPCA := 0;
        MtHTFAR := 0;
        MtHTAAR := 0;
        MtHTCCA := 0;
        RefExterne := '';
        TOBMere := TobEnCompta; // Si tout va bien on rangera la tob en cours dans les frais à passer
        TobL := fTobCutOff.Detail [ii] ;
        //        if Not VarIsNull(TobL.GetString('ACU_ETABLISSEMENT')) then
        // PL le 04/01/05 : compatibilité avec AGL 580
        if (TobL.GetString ('ACU_ETABLISSEMENT') <> '') then
          Etablissement := trim (TobL.GetString ('ACU_ETABLISSEMENT'));
        if (Etablissement = '') and (TobL.GetString ('ACU_AFFAIRE') <> '') then
          Etablissement := TobL.GetString ('AFF_ETABLISSEMENT');
        if Etablissement = '' then
          Etablissement := VH^.EtablisDefaut;
        // Tests qu'on a bien tous les éléments avant l'appel
        if (TobL.GetString ('T_AUXILIAIRE') = '') then
          CodeErreurAppel := FC_ERR_AUX;
        if (TobL.GetString ('ACU_TIERS') = '') then
          CodeErreurAppel := FC_ERR_TIE;
        if (Etablissement = '') then
          CodeErreurAppel := FC_ERR_ETA;
        if (CodeErreurAppel <> 0) then
          TOBMere := TobEnErreur;

        if (TobL.GetDouble ('ACU_FAE') <> 0.0) then //AB-200610- FQ 13247
        begin
          if TypeCumul = 'CVE' then
            MtHTFAE := TobL.GetDouble ('ACU_FAE')
          else
            MtHTFAR := TobL.GetDouble ('ACU_FAE');
        end;
        if (TobL.GetDouble ('ACU_AAE') <> 0.0) then //AB-200610- FQ 13247
        begin
          if TypeCumul = 'CVE' then
            MtHTAAE := TobL.GetDouble ('ACU_AAE')
          else
            MtHTAAR := TobL.GetDouble ('ACU_AAE');
        end;
        if (TobL.GetDouble ('ACU_PCA') <> 0.0) then //AB-200610- FQ 13247
        begin
          if TypeCumul = 'CVE' then
            MtHTPCA := TobL.GetDouble ('ACU_PCA')
          else
            MtHTCCA := TobL.GetDouble ('ACU_PCA');
        end;
        //AB-200701-FQ13137-Génération compta sur chaque montant cut-off renseigné (FAE AAE...)
        TOBMCpta := TOB.Create ('CutOffDansCompta', nil, -1);
        TOBMCpta.AddChampSupValeur ('TYPEGEN', 'CUTOFF');
        if TypeCumul = 'CVE' then
          TOBMCpta.AddChampSupValeur ('VENTEACHAT', 'VEN')
        else
          TOBMCpta.AddChampSupValeur ('VENTEACHAT', 'ACH');

        if (TobL.GetString ('ACU_AFFAIRE') <> '') and (TobL.GetString ('AFF_LIBELLE') <> '') then
          LibEcriture := TobL.GetString ('AFF_LIBELLE')
        else
          LibEcriture := TobL.GetString ('T_LIBELLE');
        RegimeTaxe := trim (TobL.GetString ('ACU_REGIMETAXE'));
        FamilleTaxe := trim (TobL.GetString ('ACU_FAMILLETAXE'));

        TOBMCpta.AddChampSupValeur ('DATEECRITURE', TobL.GetDateTime ('ACU_DATE')); // Date de l'écriture
        TOBMCpta.AddChampSupValeur ('AUXILIAIRE', TobL.GetString ('T_AUXILIAIRE')); // compte auxiliaire
        TOBMCpta.AddChampSupValeur ('ETABLISSEMENT', Etablissement); // Etablissement

        TOBMCpta.AddChampSupValeur ('ACT_ARTICLE', TobL.GetString ('ACU_ARTICLE')); // Code Article
        TOBMCpta.AddChampSupValeur ('ACT_AFFAIRE', TobL.GetString ('ACU_AFFAIRE')); // Code Affaire
        TOBMCpta.AddChampSupValeur ('ACT_TIERS', TobL.GetString ('ACU_TIERS')); // Code Client

        TOBMCpta.AddChampSupValeur ('CODETVA', FamilleTaxe); // Famille TVA
        TOBMCpta.AddChampSupValeur ('REGIMETAXE', RegimeTaxe); // Régime Taxe
        TOBMCpta.AddChampSupValeur ('LIBELLEECRITURE', LibEcriture); // Libellé écriture
        TOBMCpta.AddChampSupValeur ('DATEREFEXTERNE', ''); // Date de référence externe
        // clé pour toper: ACU_TYPEAC,ACU_DATE,ACU_TIERS,ACU_AFFAIRE,ACU_NUMECLAT,ACU_RESSOURCE,ACU_TYPEARTICLE,ACU_CODEARTICLE
        TOBMCpta.AddChampSupValeur ('ACU_TYPEAC', TobL.GetString ('ACU_TYPEAC'));
        TOBMCpta.AddChampSupValeur ('ACU_DATE', TobL.GetDateTime ('ACU_DATE'));
        TOBMCpta.AddChampSupValeur ('ACU_TIERS', TobL.GetString ('ACU_TIERS'));
        TOBMCpta.AddChampSupValeur ('ACU_AFFAIRE', TobL.GetString ('ACU_AFFAIRE'));
        TOBMCpta.AddChampSupValeur ('ACU_NUMECLAT', TobL.GetInteger ('ACU_NUMECLAT'));
        TOBMCpta.AddChampSupValeur ('ACU_RESSOURCE', TobL.GetString ('ACU_RESSOURCE'));
        TOBMCpta.AddChampSupValeur ('ACU_TYPEARTICLE', TobL.GetString ('ACU_TYPEARTICLE'));
        TOBMCpta.AddChampSupValeur ('ACU_CODEARTICLE', TobL.GetString ('ACU_CODEARTICLE'));
        TOBMCpta.AddChampSupValeur ('ACU_ERRCPTA', 0); //Code erreur en retour de passation comptable
        //AB-200701-FQ13137-
        //On genère 2 lignes compta pour chaque montant cut-off renseigné (FAE AAE...)
        while (MtHTFAE <> 0) or (MtHTFAR <> 0) or (MtHTAAE <> 0) or (MtHTAAR <> 0) or (MtHTPCA <> 0) or (MtHTCCA <> 0) do
        begin
          TOBLCpta := TOB.Create ('CutOffDansCompta', TOBMere, -1);
          TOBLCpta.Dupliquer (TOBMCpta, false, true);
          RefExterne := '';
          //On crée une ligne compta MtHTFAE MtHTAAE MtHTPCA MtHTFAR  MtHTAAR  MtHTCCA
          if (MtHTFAE <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'FAE');
            MtHT := MtHTFAE;
            MtHTFAE := 0.0;
          end
          else
            if (MtHTAAE <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'AAE');
            MtHT := MtHTAAE;
            MtHTAAE := 0.0;
          end
          else
            if (MtHTPCA <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'PCA');
            MtHT := MtHTPCA;
            MtHTPCA := 0.0;
          end
          else
            if (MtHTFAR <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'FAR');
            MtHT := MtHTFAR;
            MtHTFAR := 0.0;
          end
          else
            if (MtHTAAR <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'AAR');
            MtHT := MtHTAAR;
            MtHTAAR := 0.0;
          end
          else
            if (MtHTCCA <> 0.0) then
          begin
            TOBLCpta.AddChampSupValeur ('TYPE', 'CCA');
            MtHT := MtHTCCA;
            MtHTCCA := 0.0;
          end;
          if (TOBLCpta.GetString ('TYPE') <> 'PCA') and (TOBLCpta.GetString ('TYPE') <> 'CCA')
            and (RegimeTaxe <> '') and (FamilleTaxe <> '') then
          begin
            TOBTVA := VH^.LaTOBTVA.FindFirst (['TV_TVAOUTPF', 'TV_REGIME', 'TV_CODETAUX'] , ['TX1', RegimeTaxe, FamilleTaxe] , False);
            if (TOBTVA <> nil) then
            begin
              if TypeCumul = 'CVE' then
                Taux := TOBTVA.GetDouble ('TV_TAUXVTE') / 100.0
              else
                Taux := TOBTVA.GetDouble ('TV_TAUXACH') / 100.0;
              if taux <> 0.0 then
              begin
                MtTTC := Arrondi (MtHT * (1.0 + Taux), V_PGI.OkDecV);
                MtTVA := Arrondi (MtHT * Taux, V_PGI.OkDecV);
                RefExterne := 'TVA sur ' + TOBLCpta.GetString ('TYPE') + format (' à %2.1f', [taux * 100] ) + '%';
              end;
            end;
          end;
          RefInterne := TOBLCpta.GetString ('TYPE') + ' du ' + datetostr (TobL.GetDateTime ('ACU_DATE'));
          TOBLCpta.AddChampSupValeur ('REFINTERNE', RefInterne); // Référence interne
          TOBLCpta.AddChampSupValeur ('REFEXTERNE', RefExterne); // Référence externe
          TOBLCpta.AddChampSupValeur ('ACT_MONTANTHT', MtHT); // Montant HT
          TOBLCpta.AddChampSupValeur ('ACT_MONTANTTTC', MtTTC); // Montant TTC
          TOBLCpta.AddChampSupValeur ('ACT_MONTANTTVA', MtTVA); // Montant TVA
        end; //fin while sur les montants
        TOBMCpta.free;
      end;

      { Appel de la fonction Passation des écritures comptables des Cut-Off dans FacCpta }
      if (TOBEnCompta.Detail.Count > 0) then
      begin
        NbLignesOK := PasserCutOffEnCompta (TobEnCompta);
        repGener := NbLignesOK;

        if (RepGener >= 0) then
        begin
          if (RepGener = fTobCutOff.Detail.Count) then
            PGIInfoAF (TexteMessage [15] , TexteMessage [8] )
              {15 'La génération des cut-off a été effectuée avec succès.',}
          else
            if (RepGener < fTobCutOff.Detail.Count) and (RepGener <> 0) then
            PGIBox (format (traduitGA (TexteMessage [11] ), [RepGener, fTobCutOff.Detail.Count - RepGener] ), TraduitGA (TexteMessage [8] ))
              {11'%d lignes de cut-off générées en comptabilité et %d lignes présentent des erreurs.',}
          else
            if (RepGener = 0) then
            PGIInfoAF (TexteMessage [19] , TexteMessage [8] )
              {19 'Aucune ligne de cut-off n''a été générée en comptabilité.'}
        end
        else
          if (RepGener < 0) then
          PGIInfoAf (TexteMessage [Abs (RepGener)] , TexteMessage [8] );
      end;
    except
      repGener := FC_ERR_GEN;
    end;

  finally
    // En cas d'erreur générale dans les tests avant la fonction de génération en compta, on saute cette étape
    if (repGener >= 0) then
      // On traite les lignes en erreur dans les tests avant appel
      if (TOBEnErreur <> nil) and (TOBEnErreur.detail.count > 0) then
      begin
        fTobErreur := TOBEnErreur;
        io1 := Transactions (TraiteLignesErreur, 2);
      end;

    // Maj des lignes cut-off avec le ACU_ERRCPTA en retour de la Passation des écritures comptables
    if (repGener >= 0) then
      if (TobEnCompta <> nil) and (TobEnCompta.detail.count > 0) then
      begin
        fTobErreur := TobEnCompta;
        Transactions (TraiteLignesErreur, 2);
      end;

    if io1 <> oeOk then
      repGener := FC_ERR_TRAITE1;
    TobEnCompta.Free;
    TOBEnErreur.Free;
  end;
  Result := repGener;
end;

procedure TAFComptaCutOff.TraiteLignesErreur;
begin
  MajCutOffTraite (fTobErreur, false);
end;

end.

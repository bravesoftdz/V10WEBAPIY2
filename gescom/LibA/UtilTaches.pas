unit UtilTaches;

interface
uses
  {$IFDEF VER150} Variants, {$ENDIF}
  Hctrls,
  SysUtils,
  UtilArticle,
  AFPlanningCst,
  paramsoc,
  SaisUtil,
  AffaireUtil,
  HEnt1,
  UTOB,
  //UtilPlanning,
  ENTGC,
  ActiviteUtil,
  UAFO_Ressource,
  UAFO_REGLES,
  HRichOle,
  StdCtrls,
  utilRessource,
  Dicobtp,
  HMsgBox,
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF DBXPRESS}dbtables, {BDE}
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}
  FactComm,uEntCommun, vierge, Forms
  ;

// Chargement / Validation Tob des taches ...
procedure GetArticle(pStArticle: string; var pArticle: RecordArticle);
procedure LoadQteRAF(var pTobTaches: Tob; pStPrefixe: string);
function ChargeLesTaches(var pTobTaches: TOB; Affaire, pStTache: string; bCalcRaf: boolean): Boolean;
procedure LoadAdresseTache(pBoCreateTache : Boolean; pTobTache: Tob);
function ValideLesTaches(pTob, pTobTache, pTobRes: TOB; pStAffaire: string; pTabRes: TTabRes): Boolean;
procedure InitNewTobTache(TobTache: TOB; pCodeArticle: string; pCurAffaire: AffaireCourante);
procedure InitNewRessource(vTobRes: Tob; pCurAffaire: AffaireCourante);
// Organise tache / tob filles tacheressource
procedure OrganiseTacheRessource(TobTaches, TobTachesRes: TOB);

// numérotation des tâches
function GetNumTache(pStAffaire: string): integer;
function LoadAdresse(pBoCreateTache : Boolean; pStNiveauRecherche, pStAffaire, pStTiers, pStTypeAdr1, pStNumeroAdr1, pStTypeAdr2, pStNumeroAdr2: string; var pStTypeAdr3, pStNumeroAdr3 : string; pTob, pTobAffaire, pTobTiers: Tob): string;
function ControleCodeArticle(var pStCodeArticle, pStTypeArticle, pStArticle, pStFacturable, pStUnite, PstLibelle : string; var PR, PV: double; var pStStatutPla : string): boolean;
function ExisteUtilisateur(pStCodeRes: string): string;
function ExisteRessources(pStSqlWhere: string): boolean;
function ExisteTache(pStSqlWhere: string): Boolean;
function ExisteTacheRessource(pStAffaire, pStTache, pStRessource: string): Boolean;
function RessourceLib(pStCodeRes: string): string;
function RessourceNomPrenom(pStCodeRes: string): string;
function IsTachePrestation(pTypeArt: string): boolean;

// gestion des arguments
procedure DecodeArgument(pStTmp: string; var pStChamp, pStValeur: string);

// valorisation du planning
procedure AffectationTachePourRessource(pTob: Tob; pBoPRUnitaire, pBoPVUnitaire, pBoValoPR, pBoValoPV: Boolean);

procedure AffectationTobPR(pDevise: RDevise; pTobSource, pTobDest: Tob;
  pStPrefixe, pStPR: string; pBoTableTache: Boolean);

procedure AffectationTobPV(pDevise: RDevise; pTobSource, pTobDest: Tob;
  pStPrefixe, pStPV: string; pBoTableTache: Boolean);

function CalculMntDev(pDevise: RDevise; pRdMntPivot: Double): Double;
function GetModeValPR: Integer;
function GetModeValPV: Integer;
//procedure ValorisationReal(pTob: Tob);
procedure ValorisationPlanning(pTob: Tob; pStPrefixe: string; pAFOAssistants: TAFO_Ressources; pTOBAffaires, pTobArticles: Tob; PR, PV: boolean);

function CalculPlanifie(pStAffaire, pStNumeroTache, pStChamp, pStRessource, Typepla: string; pBoTout: Boolean): Double;
function CalculRealise(pStAffaire, pStFonction, pStChamp, pStRessource, pStArticle, Unite: string; pBoPDC: Boolean): Double;
//function CalculToutRealise(pStAffaire, pStFonction, pStChamp, pStRessource, pStArticle: string): Double;
function CalculRAFQte(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
function CalculRAFMnt(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
function CalculRAFMntPV(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
function calculPRArticle(pStArticle: string): Double;
procedure TOBCopieModeleTache(pCurAffaire: Affairecourante; FromTOB, ToTOB: TOB);
{$IFNDEF EAGLSERVER}
procedure SaisieIntervenantsTache(pStAffaire, pStNumeroTache, pStIdentligne, pStLibTache, pStAction: string);
{$ENDIF EAGLSERVER}
function FamilleTacheLie : String;
function FindFamilleTacheLie(pStCodeArticle : String) : String;
//
Procedure GenereAppel(fTOBDet : TOB);
Procedure GenererListesAppel(pListeJours : TListDate; FTOBDet : TOB; NumLigne : Integer; Duree: Double);
//procedure DefiniLibelle(fTOBDet : Tob);
Function RechRessPrinc(Affaire, NoTache : String) : String;

implementation

uses UtilPGI,
  FactAffaire,
  AppelsUtil,
  UdateUtils,
  DateUtils,
  FactAdresse,
  //UtomAffTiers,
  FactUtil,
  CalcOleGenericAff;
  //UAFO_ValorisationPrix;

const

  TexteMessage: array[1..3] of string = (
    {1}'Erreur lors de l''enregistrement de la tâche.',
    {2}'Tâche',
    {3}'Erreur de valorisation.'
       );


procedure DefiniLibelle(TOBD : TOB; Memo : THRichEditOLE);
var lng : integer;
    Indice : integer;
begin

  if Memo=nil then exit;
  IF not TOBD.FieldExists ('AFF_LIBELLE') then TOBD.AddChampSupValeur('AFF_LIBELLE','');
  // trouve la premiere ligne non vide
  Indice := 0;
  repeat
    lng := Pos (#$D#$A, Memo.lines [indice] ) ;
    if lng > 1 then break else inc(indice);
  until indice > Memo.lines.count;

  if Indice > Memo.lines.count then exit;

  if (lng > 35) then
  begin
    TOBD.PutValue ('AFF_LIBELLE', Copy (Memo.lines [Indice] , 1, 35) ) ;
  end
  else
  begin
    TOBD.PutValue ('AFF_LIBELLE', Copy (Memo.lines [indice] , 1, lng - 1) ) ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 07/11/2002
Modifié le ... :
Description .. : Calcul des quantités Restant a faire
                 Formule pour PDC
                 08/11/2005 suppression du test si on gere le RAF
                 testé en amont dans les fonctions chargeLeTaches
Mots clefs ... :
*****************************************************************}

procedure LoadQteRAF(var pTobTaches: Tob; pStPrefixe: string);
var
  i: Integer;
  vRdQte: Double;

begin

  pTobTaches.AddChampSup('QTEAPLANIFIERCALC', true);
  pTobTaches.AddChampSup('RAPPTPRCAL', true);
  pTobTaches.AddChampSup('RAPPTPVCAL', true);
  for i := 0 to pTobTaches.detail.count - 1 do
  begin
    vRdQte := CalculRAFQte(pTobTaches.detail[i], pStPrefixe, True);
    pTobTaches.detail[i].AddChampSupValeur('QTEAPLANIFIERCALC', vRdQte);
    pTobTaches.detail[i].AddChampSupValeur('RAPPTPRCAL', CalculRAFMnt(pTobTaches.detail[i], pStPrefixe, True));
    pTobTaches.detail[i].AddChampSupValeur('RAPPTPVCAL', CalculRAFMntPv(pTobTaches.detail[i], pStPrefixe, True));

    if (GetParamSocSecur('SO_AFCLIENT',0) = cInClientAlgoe) and (vRdQte = 0) and (pStPrefixe = 'ATA')
      and IsTachePrestation(pTobTaches.detail[i].GetString('ATA_TYPEARTICLE')) then  //AB-200605 - FQ 12918
    begin
      pTobTaches.detail[i].PutValue('RAPPTPRCAL', 0);
      pTobTaches.detail[i].PutValue('RAPPTPVCAL', 0);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 28/02/2002
Modifié le ... : 28/02/2002
Description .. : chargement des taches
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : TACHES
*****************************************************************}

function ChargeLesTaches(var pTobTaches: TOB; Affaire, pStTache: string; bCalcRaf: boolean): Boolean;
var
  i: integer;
  vQr: TQuery;
  vStSQL: string;
  TobTachesRes: TOB;
  vStQuelleAdresse: string;

  procedure LoadAdresseTaches(pTobTaches: Tob);
  var
    i           : Integer;
    vTob        : Tob;
    vTobTiers   : Tob;
    vTobAffaire : Tob;
    vStRien1    : String;
    vStRien2    : String;
    
  begin
    vTob := Tob.create('MONADRESSE', nil, -1);
    vTobAffaire := Tob.create('ADRESSEAFFAIRE', nil, -1);
    vTobTiers := Tob.create('ADRESSETIERS', nil, -1);
    vTobAffaire.AddChampSupValeur('ADR_LIBELLE', '', True);
    vTobAffaire.AddChampSupValeur('ADR_LIBELLE2', '', True);
    vTobAffaire.AddChampSupValeur('ADR_ADRESSE1', '', True);
    vTobAffaire.AddChampSupValeur('ADR_ADRESSE2', '', True);
    vTobAffaire.AddChampSupValeur('ADR_ADRESSE3', '', True);
    vTobAffaire.AddChampSupValeur('ADR_CODEPOSTAL', '', True);
    vTobAffaire.AddChampSupValeur('ADR_VILLE', '', True);
    vTobAffaire.AddChampSupValeur('ADR_TELEPHONE', '', True);
    vTobAffaire.AddChampSupValeur('QUELLEADRESSE', '', True); // CLIENT, TACHE, INTERVENTION
    vTobTiers.AddChampSupValeur('T_LIBELLE', '', True);
    vTobTiers.AddChampSupValeur('T_PRENOM', '', True);
    vTobTiers.AddChampSupValeur('T_ADRESSE1', '', True);
    vTobTiers.AddChampSupValeur('T_ADRESSE2', '', True);
    vTobTiers.AddChampSupValeur('T_ADRESSE2', '', True);
    vTobTiers.AddChampSupValeur('T_CODEPOSTAL', '', True);
    vTobTiers.AddChampSupValeur('T_VILLE', '', True);
    vTobTiers.AddChampSupValeur('T_TELEPHONE', '', True);
    vTobTiers.AddChampSupValeur('QUELLEADRESSE', '', True); // CLIENT, TACHE, INTERVENTION

    try
      for i := 0 to pTobTaches.detail.count - 1 do
      begin
        if i = 0 then
        begin
          pTobTaches.detail[i].AddChampSupValeur('ADR_LIBELLE', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_LIBELLE2', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_ADRESSE1', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_ADRESSE2', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_ADRESSE3', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_CODEPOSTAL', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_VILLE', '', True);
          pTobTaches.detail[i].AddChampSupValeur('ADR_TELEPHONE', '', True);
          pTobTaches.detail[i].AddChampSupValeur('QUELLEADRESSE', '', True); // CLIENT, TACHE, INTERVENTION
        end;

        vStQuelleAdresse := LoadAdresse(false, 'TACHE', pTobTaches.detail[i].GetString('ATA_AFFAIRE'),
        pTobTaches.detail[i].GetString('ATA_TIERS'),
        pTobTaches.detail[i].GetString('ATA_TYPEADRESSE'),
        pTobTaches.detail[i].GetString('ATA_NUMEROADRESSE'),
        '', '', vStRien1, vStRien2, vTob, vTobAffaire, vTobTiers);

        pTobTaches.detail[i].PutValue('QUELLEADRESSE', vStQuelleAdresse);
        pTobTaches.detail[i].PutValue('ADR_LIBELLE', vTob.GetString('CLIENTLIBELLE'));
        pTobTaches.detail[i].PutValue('ADR_LIBELLE2', vTob.GetString('CLIENTPRENOM'));
        pTobTaches.detail[i].PutValue('ADR_ADRESSE1', vTob.GetString('CLIENTADRESSE1'));
        pTobTaches.detail[i].PutValue('ADR_ADRESSE2', vTob.GetString('CLIENTADRESSE2'));
        pTobTaches.detail[i].PutValue('ADR_ADRESSE3', vTob.GetString('CLIENTADRESSE3'));
        pTobTaches.detail[i].PutValue('ADR_CODEPOSTAL', vTob.GetString('CLIENTCODEPOSTAL'));
        pTobTaches.detail[i].PutValue('ADR_VILLE', vTob.GetString('CLIENTVILLE'));
        pTobTaches.detail[i].PutValue('ADR_TELEPHONE', vTob.GetString('CLIENTTELEPHONE'));
      end;
    finally
      FreeAndNil(vTob);
      FreeAndNil(vTobAffaire);
      FreeAndNil(vTobTiers);
    end;
  end;

begin

  Result := False;

  if (pTobTaches <> nil) then
  begin

    // on charge tous les champs de la tache, fait expres
    vStSQL := 'SELECT * FROM TACHE WHERE ATA_AFFAIRE="' + Affaire + '"';
    if pStTache <> '' then
      vStSQL := vStSQL + ' AND ATA_NUMEROTACHE = ' + pStTache;
    vQR := OpenSQL(vStSQL, True);

    try
      if (not vQR.EOF) then
      begin
        Result := True;
        pTobTaches.LoadDetailDB('TACHE', '', '', vQR, True);

        pTobTaches.detail[0].AddChampSupValeur('QTEAPLANIFIERCALC', 0, True);
        pTobTaches.detail[0].AddChampSupValeur('RAPPTPRCAL', 0, True);
        pTobTaches.detail[0].AddChampSupValeur('RAPPTPVCAL', 0, True);

{Modif FV if VH_GC.GCIfDefCEGID then
        begin
          for i := 0 to pTobTaches.detail.count - 1 do
          begin
            if ArticleForfait(pTobTaches.detail[i].GetString('ATA_AFFAIRE'), pTobTaches.detail[i].GetString('ATA_NUMEROTACHE')) then
              pTobTaches.detail[i].AddChampSupValeur('ARTICLE_FORFAIT', 'X')
            else
              pTobTaches.detail[i].AddChampSupValeur('ARTICLE_FORFAIT', '-');
          end;
        end;}
      end;

      // C.B 08/11/2005
      // Chargement de l'adresse de toutes les taches
      LoadAdresseTaches(pTobTaches);

    finally
      Ferme(vQR);
    end;

    TobTachesRes := Tob.Create('liste tachesRes', nil, -1);

    try
      vStSQL := 'SELECT ATR_AFFAIRE, ATR_NUMEROTACHE, ATR_RESSOURCE,';
      vStSQL := vStSQL + 'ATR_TIERS, ATR_FONCTION, ATR_AFFAIRE0,';
      vStSQL := vStSQL + 'ATR_AFFAIRE1, ATR_AFFAIRE2, ATR_AFFAIRE3,';
      vStSQL := vStSQL + 'ATR_AVENANT, ATR_COMPETENCE1, ATR_COMPETENCE2,';
      vStSQL := vStSQL + 'ATR_COMPETENCE3, ATR_QTEINITIALE,ATR_QTEINITPLA,';
      vStSQL := vStSQL + 'ATR_ECARTQTEINIT, ATR_POURCENTAGE,';
      vStSQL := vStSQL + 'ATR_DEVISE, ATR_PUPR, ATR_PUPRDEV, ATR_PUVENTEHT,';
      vStSQL := vStSQL + 'ATR_PUVENTEDEVHT, ATR_INITPTPR, ATR_INITPTPRDEV,';
      vStSQL := vStSQL + 'ATR_INITPTVENTEHT, ATR_INITPTVTDEVHT,';
      vStSQL := vStSQL + 'ATR_ECARTPR, ATR_ECARTPRDEV, ATR_ECARTPVHT,';
      vStSQL := vStSQL + 'ATR_ECARTPVHTDEV, ';
      vStSQL := vStSQL + ' ATR_STATUTRES, ';
      vStSQL := vStSQL + ' ARS_LIBELLE ';
      vStSQL := vStSQL + ' FROM TACHERESSOURCE, RESSOURCE ';
      vStSQL := vStSQL + ' WHERE ATR_AFFAIRE="' + Affaire + '"';
      if pStTache <> '' then
        vStSQL := vStSQL + ' AND ATR_NUMEROTACHE = ' + pStTache;
      vStSQL := vStSQL + ' AND ATR_RESSOURCE = ARS_RESSOURCE ';
      vStSQL := vStSQL + ' AND ATR_STATUTRES <> "INA"';
      vQr := nil;
      try
        vQR := OpenSQL(vStSQL, True);
        if not (vQR.EOF) then
          TobTachesRes.LoadDetailDB('TACHERESSOURCE', '', '', vQR, True);
      finally
        Ferme(vQR);
      end;

      // Positionnement des tâches/ressources sous les tâches
      OrganiseTacheRessource(pTobTaches, TobTachesRes);
      
      // calcul des qtes et montants RAF pour le plan de charge
      if bCalcRaf then
        LoadQteRAF(pTobTaches, 'ATA');
      // lien ligne affaire
      for i := 0 to pTobTaches.detail.count - 1 do
        if pTobTaches.Detail[i].getvalue('ATA_IDENTLIGNE') <> '' then
          pTobTaches.Detail[i].AddChampSupValeur('IDENTLIGNE', 'X')
        else
          pTobTaches.Detail[i].AddChampSupValeur('IDENTLIGNE', '-');

    finally
      TobTachesRes.free;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 08/11/2005
Modifié le ... :   /  /
Description .. : chargement de l'adresse de la tache selectionnée
Mots clefs ... :
*****************************************************************}
procedure LoadAdresseTache(pBoCreateTache : Boolean; pTobTache: Tob);
var
  vTob              : Tob;
  vStQuelleAdresse  : string;
  vStTypeAdresse    : String;
  vStNumeroAdresse  : String;
  vStQuelContact    : String;
begin

  vTob := Tob.create('MONADRESSE', nil, -1);

  try
    vStQuelleAdresse := LoadAdresse(pBoCreateTache, 'TACHE', pTobTache.GetString('ATA_AFFAIRE'),
                                    pTobTache.GetString('ATA_TIERS'), pTobTache.GetString('ATA_TYPEADRESSE'),
                                    pTobTache.GetString('ATA_NUMEROADRESSE'),'', '', vStTypeAdresse, vStNumeroAdresse, vTob, nil, nil);

    if vStQuelleAdresse = 'AFFAIRE' then
       begin
       vStQuelleAdresse := 'TACHE';
       pTobTache.PutValue('ATA_TYPEADRESSE', vStTypeAdresse);
       pTobTache.PutValue('ATA_NUMEROADRESSE', vStNumeroAdresse);
       end;

    pTobTache.PutValue('QUELLEADRESSE', vStQuelleAdresse);
    pTobTache.PutValue('ADR_LIBELLE', vTob.GetString('CLIENTLIBELLE'));
    pTobTache.PutValue('ADR_LIBELLE2', vTob.GetString('CLIENTPRENOM'));
    pTobTache.PutValue('ADR_ADRESSE1', vTob.GetString('CLIENTADRESSE1'));
    pTobTache.PutValue('ADR_ADRESSE2', vTob.GetString('CLIENTADRESSE2'));
    pTobTache.PutValue('ADR_ADRESSE3', vTob.GetString('CLIENTADRESSE3'));
    pTobTache.PutValue('ADR_CODEPOSTAL', vTob.GetString('CLIENTCODEPOSTAL'));
    pTobTache.PutValue('ADR_VILLE', vTob.GetString('CLIENTVILLE'));
    pTobTache.PutValue('ADR_TELEPHONE', vTob.GetString('CLIENTTELEPHONE'));
    pTobTache.PutValue('ATA_FAITPARQUI', vTob.GetString('CLIENTCONTACT'));
  finally
    FreeAndNil(vTob);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : pBoPRUnitaire : on ne calcule pas les prix unitaires
                 par ressource car cela n'a pas de sens
Ptob : tob tache + tache ressource
Affecte champs Prix unitaire, (PUPr et PUPV), champs mtt sur qté unitaire (init*)
et champs mtt sur qte reste à faire (RAPP*)
Mots clefs ... :
*****************************************************************}

procedure AffectationTachePourRessource(pTob: Tob; pBoPRUnitaire, pBoPVUnitaire, pBoValoPR, pBoValoPV: Boolean);
var
  vRdINITPTPR: Double;
  vRdINITPTVENTEHT: Double;
  vRdINITPTPRDEV: Double;
  vRdINITPTVTDEVHT: Double;
  vRdRAPPTPR: Double;
  vRdRAPPTVENTEHT: Double;
  vRdRAPPTPRDEV: Double;
  vRdRAPPTVTDEVHT: Double;
  i: Integer;

begin

  if pBoValoPR then
  begin
    if pTob.detail.count = 0 then
    begin
      pTob.putvalue('ATA_PUPR', 0);
      pTob.putvalue('ATA_PUPRDEV', 0);
      pTob.putvalue('ATA_INITPTPR', 0);
      pTob.putvalue('ATA_INITPTPRDEV', 0);
      pTob.putvalue('ATA_ECARTPR', 0);
      pTob.putvalue('ATA_ECARTPRDEV', 0);
    end
    else
    begin
      if pBoPRUnitaire then
      begin
        pTob.putvalue('ATA_PUPR', pTob.detail[0].GetValue('ATR_PUPR'));
        pTob.putvalue('ATA_PUPRDEV', pTob.detail[0].GetValue('ATR_PUPRDEV'));
      end
      else
      begin
        pTob.putvalue('ATA_PUPR', 0);
        pTob.putvalue('ATA_PUPRDEV', 0);
      end;

      vRdINITPTPR := 0;
      vRdINITPTPRDEV := 0;
      vRdRAPPTPR := 0;
      vRdRAPPTPRDEV := 0;

      for i := 0 to pTob.detail.count - 1 do
      begin
        vRdINITPTPR := vRdINITPTPR + valeur(pTob.detail[i].GetValue('ATR_INITPTPR'));
        vRdINITPTPRDEV := vRdINITPTPRDEV + valeur(pTob.detail[i].GetValue('ATR_INITPTPRDEV'));
        vRdRAPPTPR := vRdRAPPTPR + valeur(pTob.detail[i].GetValue('ATR_ECARTPR'));
        vRdRAPPTPRDEV := vRdRAPPTPRDEV + valeur(pTob.detail[i].GetValue('ATR_ECARTPRDEV'));
      end;
      pTob.putvalue('ATA_INITPTPR', vRdINITPTPR);
      pTob.putvalue('ATA_INITPTPRDEV', vRdINITPTPRDEV);
      pTob.putvalue('ATA_ECARTPR', vRdRAPPTPR);
      pTob.putvalue('ATA_ECARTPRDEV', vRdRAPPTPRDEV);
    end;
  end;

  if pBoValoPV then
  begin
    if pTob.detail.count = 0 then
    begin
      pTob.putvalue('ATA_PUVENTEHT', 0);
      pTob.putvalue('ATA_PUVENTEDEVHT', 0);
      pTob.putvalue('ATA_INITPTVENTEHT', 0);
      pTob.putvalue('ATA_INITPTVTDEVHT', 0);
      pTob.putvalue('ATA_ECARTPVHT', 0);
      pTob.putvalue('ATA_ECARTPVHTDEV', 0);
    end
    else
    begin
      if pBoPVUnitaire then
      begin
        pTob.putvalue('ATA_PUVENTEHT', pTob.detail[0].GetValue('ATR_PUVENTEHT'));
        pTob.putvalue('ATA_PUVENTEDEVHT', pTob.detail[0].GetValue('ATR_PUVENTEDEVHT'));
      end
      else
      begin
        pTob.putvalue('ATA_PUVENTEHT', 0);
        pTob.putvalue('ATA_PUVENTEDEVHT', 0);
      end;

      vRdINITPTVENTEHT := 0;
      vRdINITPTVTDEVHT := 0;
      vRdRAPPTVENTEHT := 0;
      vRdRAPPTVTDEVHT := 0;

      for i := 0 to pTob.detail.count - 1 do
      begin
        vRdINITPTVENTEHT := vRdINITPTVENTEHT + valeur(pTob.detail[i].GetValue('ATR_INITPTVENTEHT'));
        vRdINITPTVTDEVHT := vRdINITPTVTDEVHT + valeur(pTob.detail[i].GetValue('ATR_INITPTVTDEVHT'));
        vRdRAPPTVENTEHT := vRdRAPPTVENTEHT + valeur(pTob.detail[i].GetValue('ATR_ECARTPVHT'));
        vRdRAPPTVTDEVHT := vRdRAPPTVTDEVHT + valeur(pTob.detail[i].GetValue('ATR_ECARTPVHTDEV'));
      end;

      pTob.putvalue('ATA_INITPTVENTEHT', vRdINITPTVENTEHT);
      pTob.putvalue('ATA_INITPTVTDEVHT', vRdINITPTVTDEVHT);
      pTob.putvalue('ATA_ECARTPVHT', vRdRAPPTVENTEHT);
      pTob.putvalue('ATA_ECARTPVHTDEV', vRdRAPPTVTDEVHT);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Fct qui permet de valoriser les mtt réaliser de afplanning
Suite ........ : appeler lors de la génération du planning ou lors de la saisie
Suite ........ : d'un item du planning.
Suite ........ : Prend le PU en PR et PV qui a été calculer avant pour la
Suite ........ : valorisation de la qté planifiée, et le multiplue par la qté
Suite ........ : réalisée
Mots clefs ... :
*****************************************************************}
//AB-200612- Fonction integrée dans ValorisationPlanning
{procedure ValorisationReal(pTob: Tob);
var vDevise: RDevise;
begin
  vDevise.code := pTob.GetString('APL_DEVISE');  //AB-200602-
  GetInfosDevise(vDevise);
  pTob.putvalue('APL_REALPTPR', Arrondi(pTob.GetDouble('APL_PUPR') * pTob.GetDouble('APL_QTEREALISE'), V_PGI.OkDecV));
  pTob.putvalue('APL_REALPTPRDEV', Arrondi(pTob.GetDouble('APL_PUPRDEV') * pTob.GetDouble('APL_QTEREALISE'), vDevise.Decimale));
  pTob.putvalue('APL_REALPTVENTEHT', Arrondi(pTob.GetDouble('APL_PUVENTEHT') * pTob.GetDouble('APL_QTEREALISE'), V_PGI.OkDecV));
  pTob.putvalue('APL_REALPTVTDEVHT', Arrondi(pTob.GetDouble('APL_PUVENTEDEVHT') * pTob.GetDouble('APL_QTEREALISE'), vDevise.Decimale));
end; }

procedure AffectationTobPR(pDevise: RDevise; pTobSource, pTobDest: Tob;
  pStPrefixe, pStPR: string; pBoTableTache: Boolean);
var dPUPRHT,dPUPRDEVHT :double;
begin
  //AB-200601- Ajuste le nombre de décimal Px unitaire avec V_PGI.OkDecP et montant avec euro dev
  dPUPRHT := arrondi(pTobSource.GetDouble (pStPR), V_PGI.OkDecP);
  dPUPRDEVHT := arrondi(CalculMntDev(pDevise, dPUPRHT), V_PGI.OkDecP);

  pTobDest.putvalue(pStPrefixe + '_PUPR', dPUPRHT);
  pTobDest.putvalue(pStPrefixe + '_PUPRDEV', dPUPRDEVHT);
  pTobDest.putvalue(pStPrefixe + '_DEVISE', pDevise.Code);

  if pBoTableTache then
  begin
    pTobDest.putvalue(pStPrefixe + '_INITPTPR', Arrondi(dPUPRHT * pTobDest.GetDouble(pStPrefixe + '_QTEINITIALE'),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_INITPTPRDEV', Arrondi(dPUPRDEVHT * pTobDest.GetDouble(pStPrefixe + '_QTEINITIALE'),pDevise.Decimale));

    // calculer le raf
    pTobDest.putvalue(pStPrefixe + '_ECARTPR', Arrondi(dPUPRHT * pTobDest.GetDouble(pStPrefixe + '_ECARTQTEINIT'),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_ECARTPRDEV', Arrondi(dPUPRDEVHT * pTobDest.GetDouble(pStPrefixe + '_ECARTQTEINIT'),pDevise.Decimale));
  end
  else
  begin
    pTobDest.putvalue(pStPrefixe + '_INITPTPR', Arrondi(dPUPRHT * pTobDest.GetDouble(pStPrefixe + '_QTEPLANIFIEE'),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_INITPTPRDEV', Arrondi(dPUPRDEVHT * pTobDest.GetDouble(pStPrefixe + '_QTEPLANIFIEE'),pDevise.Decimale));
  end;
end;

procedure AffectationTobPV(pDevise: RDevise; pTobSource, pTobDest: Tob;
  pStPrefixe, pStPV: string; pBoTableTache: Boolean);
var dPUVENTEHT,dPUVENTEDEVHT:double;
begin
  // pTobDest.putvalue(pStPrefixe + '_PUVENTEHT', valeur(pTobSource.GetValue(pStPV)));
  //AB-200601- Ajuste le nombre de décimal Px unitaire avec V_PGI.OkDecP et montant avec euro dev
  dPUVENTEHT := arrondi(pTobSource.GetDouble (pStPV),V_PGI.OkDecP);
  dPUVENTEDEVHT := arrondi(CalculMntDev(pDevise, dPUVENTEHT),V_PGI.OkDecP);
  pTobDest.putvalue(pStPrefixe + '_PUVENTEHT', dPUVENTEHT);
  pTobDest.putvalue(pStPrefixe + '_PUVENTEDEVHT', dPUVENTEDEVHT);

  if pBoTableTache then
  begin
    pTobDest.putvalue(pStPrefixe + '_ECARTPVHT', Arrondi(dPUVENTEHT * valeur(pTobDest.GetValue(pStPrefixe + '_ECARTQTEINIT')),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_ECARTPVHTDEV', Arrondi(dPUVENTEDEVHT * valeur(pTobDest.GetValue(pStPrefixe + '_ECARTQTEINIT')),pDevise.Decimale));

    pTobDest.putvalue(pStPrefixe + '_INITPTVENTEHT', Arrondi(dPUVENTEHT * valeur(pTobDest.GetValue(pStPrefixe + '_QTEINITIALE')),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_INITPTVTDEVHT', Arrondi(dPUVENTEDEVHT * valeur(pTobDest.GetValue(pStPrefixe + '_QTEINITIALE')),pDevise.Decimale));
  end
  else
  begin
    pTobDest.putvalue(pStPrefixe + '_INITPTVENTEHT', Arrondi(dPUVENTEHT * valeur(pTobDest.GetValue(pStPrefixe + '_QTEPLANIFIEE')),V_PGI.OkDecV));
    pTobDest.putvalue(pStPrefixe + '_INITPTVTDEVHT', Arrondi(dPUVENTEDEVHT * valeur(pTobDest.GetValue(pStPrefixe + '_QTEPLANIFIEE')),pDevise.Decimale));
end;
end;

function CalculMntDev(pDevise: RDevise; pRdMntPivot: Double): Double;
begin
  // C.B 18/06/2003 Suppression contrevaleur
  ConvertPivotToDev(pDevise, pRdMntPivot, result);
end;

function GetModeValPR: Integer;
begin
  if GetparamSocSecur('SO_AFVALOPR','') = 'ART' then
    result := cInValArticle
  else
    if GetparamSocSecur('SO_AFVALOPR','') = 'FON' then
    result := cInValFonction
  else
    if GetparamSocSecur('SO_AFVALOPR','') = 'RES' then
    result := cInValRessource
  else
    if GetparamSocSecur('SO_AFVALOPR','') = 'TAC' then
    result := cInValTache
  else
    result := cInValRessource;
end;

function GetModeValPV: Integer;
begin
  if GetparamSocSecur('SO_AFVALOPV','') = 'ART' then
    result := cInValArticle
  else
    if GetparamSocSecur('SO_AFVALOPV','') = 'FON' then
    result := cInValFonction
  else
    if GetparamSocSecur('SO_AFVALOPV','') = 'RES' then
    result := cInValRessource
  else
    if GetparamSocSecur('SO_AFVALOPV','') = 'TAC' then
    result := cInValTache
  else
    if GetparamSocSecur('SO_AFVALOPV','') = 'TAR' then
{Modif FV result := cInValTarif  //AB-200611- Valorisation par tarifs clients}
  else
    result := cInValArticle;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : valorisation des taches, des tachesressource
  Ptob : tob à revalorise,
 Pstprefixe indiquant le type de table TAC pour TAChe et tache ressource, APL pour planning ..
  prend en compte pour le PR et le PV le mode de valorisation des paramsoc
 Mots clefs ... :
*****************************************************************}

procedure ValorisationPlanning(pTob: Tob; pStPrefixe: string; pAFOAssistants: TAFO_Ressources;
  pTOBAffaires, pTobArticles: Tob; PR, PV: Boolean);
var
  vInModeValPR: Integer;
  vInModeValPV: Integer;
  vSt: string;
  vQr: TQuery;
  vTobValo: Tob;
  vDevise: RDevise;
  i: Integer;
  vBoValoPV: Boolean;
  vBoPRUnitaire: Boolean;
  vBoPVUnitaire: Boolean;
  vBoTache: Boolean;
  DateLigne: Tdatetime;
  //AFO_Valorisation: TAFValorisationPrix;

  procedure ValorisationReal(pTob: Tob); //AB-200612- On valorise les montants réalisés de l'item du planning
  begin
    pTob.putvalue('APL_REALPTPR', Arrondi(pTob.GetDouble('APL_PUPR') * pTob.GetDouble('APL_QTEREALISE'), V_PGI.OkDecV));
    pTob.putvalue('APL_REALPTPRDEV', Arrondi(pTob.GetDouble('APL_PUPRDEV') * pTob.GetDouble('APL_QTEREALISE'), vDevise.Decimale));
    pTob.putvalue('APL_REALPTVENTEHT', Arrondi(pTob.GetDouble('APL_PUVENTEHT') * pTob.GetDouble('APL_QTEREALISE'), V_PGI.OkDecV));
    pTob.putvalue('APL_REALPTVTDEVHT', Arrondi(pTob.GetDouble('APL_PUVENTEDEVHT') * pTob.GetDouble('APL_QTEREALISE'), vDevise.Decimale));
  end;
begin
  try
    // valorisation
    //vBoTache permet de savoir si on est sur les taches ou sur le planning
    //pour valoriser differemment si la methode de valorisation est ressource
    vBoTache := (pStPrefixe <> 'APL');
    //mcd maintenenant que l'on n'a plus la  QTe et MMT, il ne faut surtout pas recalculé les prix si qte =0
    if vbotache and (Ptob.getValue('ATA_QTEINITIALE') = 0) then
      exit;
    if not vbotache and (ptob.Getvalue('APL_QTEPLANIFIEE') = 0) then
      exit;
    // renseigner les prix unitaires de la planning
    vBoPRUnitaire := true;
    vBoPVUnitaire := true;
    vBoValoPV := false;
    // devise
    if pStPrefixe = 'ATA' then
      vDevise.code := pTob.GetValue('ATA_DEVISE');
    if (vDevise.code = '') or (vDevise.code = #0) then
      vDevise.Code := V_PGI.DevisePivot;
    // si ligne de planning, il faut valoriser les ligne sur la date réelle de la ligne (util en ressource car histo de prix)
    if pStPrefixe = 'APL' then
      DateLigne := pTob.Getvalue('APL_DATEDEBPLA')
    else
      DateLigne := now;
    GetInfosDevise(vDevise);
    vInModeValPR := GetModeValPR;
    vInModeValPV := GetModeValPV;

{Modif FV AFO_Valorisation := TAFValorisationPrix.create(pAFOAssistants,pTobArticles,pTOBAffaires,nil);
    try
    with AFO_Valorisation do
      begin
      gStDevise := vDevise.Code;
      gDate := DateLigne;
      if not PR then vInModeValPr := 99; // cas ou on ne veut pas la valo par PR
      //Appel AFO_Valorisation.MajTOBValo uniquement si :
      //valorisation Prix revient du planning est RES ou valorisation Prix du planning est RES ou TAR
      if vInModeValPR in [cInValRessource] then gStValoActPR := GetparamSocSecur('SO_AFVALOPR','');
      if vInModeValPV in [cInValRessource,cInValTarif] then gStValoActPV := GetparamSocSecur('SO_AFVALOPV','');
      case vInModeValPR of
           cInValArticle: // revalo article pour tache
           begin
           // Article on utilise le PMRP (prix moren pondéré car s'est le seul saisi)
           vSt := 'SELECT GA_PVHT, GA_PMRP FROM ARTICLE WHERE GA_ARTICLE = "' + PTob.Getvalue(pStPrefixe + '_ARTICLE') + '"';
           vQr := nil;
           vTobValo := Tob.Create('tob valo', nil, -1);
           try
              vQR := OpenSQL(vSt, True);
              if (not vQR.EOF) then
                begin
                vTobValo.LoadDetailDB('tob valo', '', '', vQR, True);
                AffectationTobPR(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'GA_PMRP', vBoTache);
                if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                  for i := 0 to pTob.detail.count - 1 do
                    AffectationTobPR(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'GA_PMRP', vBoTache);
                if (vInModeValPV = cInValArticle) and (PV) then //mcd 10/09/04 ajput PV ... si false, il ne faut pas faire
                begin
                  AffectationTobPV(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'GA_PVHT', vBoTache);
                  if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                    for i := 0 to pTob.detail.count - 1 do
                      AffectationTobPV(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                  vBoValoPV := true;
                end;
              end;
            finally
              Ferme(vQR);
              FreeAndNil(vTobValo);
            end;
           end; // fin cas revalo articl pour planning

        cInValFonction:
          begin
            if PTob.GetString (pStPrefixe + '_FONCTION') <> '' then
            begin
              vSt := 'SELECT AFO_TAUXREVIENTUN, AFO_PVHT FROM FONCTION WHERE AFO_FONCTION = "' + PTob.GetString(pStPrefixe + '_FONCTION') + '"';
            vQr := nil;
            vTobValo := Tob.Create('tob valo', nil, -1);
            try
              vQR := OpenSQL(vSt, True);
              if (not vQR.EOF) then
              begin
                vTobValo.LoadDetailDB('tob valo', '', '', vQR, True);
                AffectationTobPR(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'AFO_TAUXREVIENTUN', vBoTache);
                if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                  for i := 0 to pTob.detail.count - 1 do
                    AffectationTobPR(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'AFO_TAUXREVIENTUN', vBoTache);
                if (vInModeValPV = cInValFonction) and (PV) then //mcd 10/09/04 ajput PV ... si false, il ne faut pas faire
                begin
                  AffectationTobPV(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'AFO_PVHT', vBoTache);
                  if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                    for i := 0 to pTob.detail.count - 1 do
                      AffectationTobPV(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'AFO_PVHT', vBoTache);
                  vBoValoPV := true;
                end;
              end;
            finally
              Ferme(vQR);
              FreeAndNil(vTobValo);
            end;
          end; // fin revalo par fct pour planning
          end; // fin revalo par fct pour planning
        cInValRessource:
          begin
            if vBoTache then
            begin
              for i := 0 to pTob.detail.count - 1 do
              begin
                gStAffaire := PTob.Getvalue('ATA_AFFAIRE');
                gStTiers := PTob.Getvalue('ATA_TIERS');
                gStCodeArticle := PTob.Getvalue('ATA_CODEARTICLE');
                gStRessource := PTob.detail[i].Getvalue('ATR_RESSOURCE');
                if MajTOBValo then
                begin
                  vTOBValo := gTobArticleValorise; //tob article avec GA_PMRP et GA_PVHT
                  if gbPROK and (vTobValo <> nil) then
                  begin
                    affectationTobPR(vDevise, vTobValo, PTob.detail[i], 'ATR', 'GA_PMRP', vBoTache);
                    vBoPRUnitaire := false;
                  end;
                  if PV and gbPVOK and (vTobValo <> nil) and  (vInModeValPV in [cInValRessource,cInValTarif]) then //mcd 10/09/04 ajput PV ... si false, il ne faut pas faire
                  begin
                    AffectationTobPV(vDevise, vTobValo, pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                    vBoValoPV := true;
                    if vInModeValPV in [cInValRessource] then
                      vBoPVUnitaire := False;
                  end;
                end;
              end; // fin for  sur détail tob
              // prix de la tache = somme des prix des ressources Mcd modif Bvalo appel fct : on est dans le cas de valo par PR
              AffectationTachePourRessource(pTob, vBoPRUnitaire, vBoPVUnitaire, true, False);
              vTobValo := LibereTobValo;
            end // fin vbotache
            else // planning -> ressource est connue
            begin // ligne plan de charge ou planning APL_
              gStAffaire := PTob.GetString('APL_AFFAIRE');
              gStTiers := PTob.GetString('APL_TIERS');
              gStRessource := PTob.GetString('APL_RESSOURCE');
              gStCodeArticle := PTob.GetString('APL_CODEARTICLE');
              gStUniteQte := PTob.GetString('APL_UNITETEMPS');
              gdbQte := PTob.GetDouble('APL_QTEPLANIFIEE');
              if MajTOBValo then
              begin
                vTOBValo := gTobArticleValorise;  //tob article avec GA_PMRP et GA_PVHT
                if gbPROK and (vTobValo <> nil) then
                  AffectationTobPR(vDevise, vTobValo, pTob, pStPrefixe, 'GA_PMRP', vBoTache);
                if PV and gbPVOK and (vTobValo <> nil) and (vInModeValPV in [cInValRessource,cInValTarif]) then //mcd 10/09/04 ajput PV ... si false, il ne faut pas faire
                begin
                  AffectationTobPV(vDevise, vTobValo, pTob, pStPrefixe, 'GA_PVHT', vBoTache);
                  vBoValoPV := true;
                end;
              end;
              vTobValo := LibereTobValo; //on libere AFO_Valorisation.gTobArticleValorise
            end; // fin sur planning
          end; // fin valo par ressource
        cInValTache:
          begin // Tache
            if not vBoTache then //AB-200510-
            begin
              vSt := 'SELECT ATA_PUVENTEHT, ATA_PUPR FROM TACHE WHERE ATA_AFFAIRE= "' + PTob.GetString(pStPrefixe + '_AFFAIRE') + '" AND ATA_NUMEROTACHE= ' + PTob.GetString(pStPrefixe + '_NUMEROTACHE');
              vTobValo := Tob.Create('tob valo', nil, -1);
              vQR := OpenSQL(vSt, True);
              if (not vQR.EOF) then
              begin
                vTobValo.SelectDB('', vQR);
                AffectationTobPR(vDevise, vTobValo, pTob, pStPrefixe, 'ATA_PUPR', vBoTache);
                if vInModeValPV = cInValTache then
                  AffectationTobPV(vDevise, vTobValo, pTob, pStPrefixe, 'ATA_PUVENTEHT', vBoTache)
                else
                  vBoPVUnitaire := False;
              end;
              Ferme(vQR);
              FreeAndNil(vTobValo);
            end
            else
            begin
              AffectationTobPR(vDevise, pTob, pTob, pStPrefixe, pStPrefixe + '_PUPR', vBoTache);
              for i := 0 to pTob.detail.count - 1 do
                AffectationTobPR(vDevise, pTob, pTob.detail[i], 'ATR', pStPrefixe + '_PUPR', vBoTache);
              if vInModeValPV = cInValTache then
              begin
                AffectationTobPV(vDevise, pTob, pTob, pStPrefixe, pStPrefixe + '_PUVENTEHT', vBoTache);
                for i := 0 to pTob.detail.count - 1 do
                  AffectationTobPV(vDevise, pTob, pTob.detail[i], 'ATR', pStPrefixe + '_PUVENTEHT', vBoTache);
                vBoValoPV := true;
              end
              else
                vBoPVUnitaire := False;
            end
          end; // fin valo par tache
      end; // fin case
      // si la valorisation PV n'est pas encore faite
      if PV and not vBoValoPV then
        case vInModeValPV of
          cInValArticle: //Valorisation PV par article
            begin
              vSt := 'SELECT GA_PVHT, GA_PMRP FROM ARTICLE WHERE GA_ARTICLE = "' + PTob.GetString(pStPrefixe + '_ARTICLE') + '"';
              vQr := nil;
              vTobValo := Tob.Create('tob valo', nil, -1);
              try
                vQR := OpenSQL(vSt, True);
                if (not vQR.EOF) then
                begin
                  vTobValo.LoadDetailDB('tob valo', '', '', vQR, True);
                  AffectationTobPV(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'GA_PVHT', vBoTache);
                  if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                    for i := 0 to pTob.detail.count - 1 do
                      AffectationTobPV(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                end;
              finally
                Ferme(vQR);
                FreeAndNil(vTobValo);
              end;
            end; // fin valo article
          cInValFonction:  //Valorisation PV par fonction
            begin
              if PTob.GetString (pStPrefixe + '_FONCTION') <> '' then
              begin
                vSt := 'SELECT AFO_TAUXREVIENTUN, AFO_PVHT FROM FONCTION WHERE AFO_FONCTION = "' + PTob.GetString(pStPrefixe + '_FONCTION') + '"';
              vQr := nil;
              vTobValo := Tob.Create('tob valo', nil, -1);
              try
                vQR := OpenSQL(vSt, True);
                if (not vQR.EOF) then
                begin
                  vTobValo.LoadDetailDB('tob valo', '', '', vQR, True);
                  AffectationTobPV(vDevise, vTobValo.detail[0], pTob, pStPrefixe, 'AFO_PVHT', vBoTache);
                  if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                    for i := 0 to pTob.detail.count - 1 do
                      AffectationTobPV(vDevise, vTobValo.detail[0], pTob.detail[i], 'ATR', 'AFO_PVHT', vBoTache);
                end;
              finally
                Ferme(vQR);
                FreeAndNil(vTobValo);
              end;
            end; // fin valo fct
            end; // fin valo fct

          cInValRessource: //Valorisation PV par ressource
            begin
              if vBoTache then
              begin // on est en valo de la tache par ressource
                gStAffaire := PTob.Getvalue('ATA_AFFAIRE');
                gStTiers := PTob.Getvalue('ATA_TIERS');
                gStCodeArticle := PTob.Getvalue('ATA_CODEARTICLE');
                for i := 0 to pTob.detail.count - 1 do
                begin
                  gStRessource := PTob.detail[i].Getvalue('ATR_RESSOURCE');
                  if MajTOBValo then
                  begin
                    vTOBValo := gTobArticleValorise; //tob article avec GA_PVHT valorisé
                    if gbPVOK and (vTobValo <> nil) then
                      AffectationTobPV(vDevise, vTobValo, pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                    vBoPVUnitaire := False;
                  end;
                end;
                // prix de la tache = somme des prix des ressources Mcd modif zone valo.. on estd ans le cas de valo pv uniquement
                affectationTachePourRessource(pTob, vBoPRUnitaire, vBoPVUnitaire, False, true);
                vTobValo := LibereTobValo; //on libere AFO_Valorisation.gTobArticleValorise
              end // fin valo tache par ressourc
              else
              begin // planning -> ressource est connue
                gStAffaire := PTob.Getvalue(pStPrefixe + '_AFFAIRE');
                gStTiers := PTob.Getvalue(pStPrefixe + '_TIERS');
                gStRessource := PTob.Getvalue(pStPrefixe + '_RESSOURCE');
                gStCodeArticle := PTob.Getvalue(pStPrefixe + '_CODEARTICLE');
                if MajTOBValo then
                begin
                  vTOBValo := gTobArticleValorise; //tob article avec GA_PVHT valorisé
                  if gbPVOK and (vTobValo <> nil) then
                  begin
                    AffectationTobPV(vDevise, vTobValo, pTob, pStPrefixe, 'GA_PVHT', vBoTache);
                    if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                      for i := 0 to pTob.detail.count - 1 do
                        AffectationTobPV(vDevise, vTobValo, pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                  end;
                end;
                vTobValo := LibereTobValo; //on libere AFO_Valorisation.gTobArticleValorise
              end;
            end; // fin valo par ressource

          cInValTache:  //Valorisation PV par tache
            begin
              if not vBoTache then //AB-200510-
              begin
                vSt := 'SELECT ATA_PUVENTEHT, ATA_PUPR FROM TACHE WHERE ATA_AFFAIRE= "' + PTob.GetString(pStPrefixe + '_AFFAIRE') + '" AND ATA_NUMEROTACHE= ' + PTob.GetString(pStPrefixe + '_NUMEROTACHE');
                vTobValo := Tob.Create('tob valo', nil, -1);
                vQR := OpenSQL(vSt, True);
                if (not vQR.EOF) then
                begin
                  vTobValo.SelectDB('', vQR);
                  AffectationTobPV(vDevise, vTobValo, pTob, pStPrefixe, 'ATA_PUVENTEHT', vBoTache);
                end;
                Ferme(vQR);
                FreeAndNil(vTobValo);
              end
              else
              begin
                AffectationTobPV(vDevise, pTob, pTob, pStPrefixe, pStPrefixe + '_PUVENTEHT', vBoTache);
                for i := 0 to pTob.detail.count - 1 do
                  AffectationTobPV(vDevise, pTob, pTob.detail[i], 'ATR', pStPrefixe + '_PUVENTEHT', vBoTache);
              end;
            end;
          cInValTarif:  //AB-200611- Valorisation PV par tarifs clients
              begin
                gStAffaire := PTob.GetString(pStPrefixe + '_AFFAIRE');
                gStTiers := PTob.GetString(pStPrefixe + '_TIERS');
                gStRessource := PTob.GetString(pStPrefixe + '_RESSOURCE');
                gStCodeArticle := PTob.GetString(pStPrefixe + '_CODEARTICLE');
                if not vBoTache then
                begin
                  gStUniteQte := PTob.GetString(pStPrefixe + 'APL_UNITETEMPS');
                  gdbQte := PTob.GetDouble(pStPrefixe + 'APL_QTEPLANIFIEE');
                end;
                if MajTOBValo then
                begin
                  vTOBValo := gTobArticleValorise; //tob article avec GA_PVHT valorisé
                  if gbPVOK and (vTobValo <> nil) then
                  begin
                    AffectationTobPV(vDevise, vTobValo, pTob, pStPrefixe, 'GA_PVHT', vBoTache);
                    if vBoTache then //mcd ... à ne pas faire surmaj afplanning ..
                      for i := 0 to pTob.detail.count - 1 do
                        AffectationTobPV(vDevise, vTobValo, pTob.detail[i], 'ATR', 'GA_PVHT', vBoTache);
                  end;
                end;
                vTobValo := LibereTobValo; //on libere AFO_Valorisation.gTobArticleValorise
              end;
        end; // fin du case
    end;
    finally
      AFO_Valorisation.Destroy;
    end;}

    //AB-200612- On valorise les montants réalisés de l'item du planning
    if not vBoTache then
      ValorisationReal(pTob);
  except
    on E: Exception do
    begin
      MessageAlerte(texteMessage[3] + #10#13 + E.Message);
      V_PGI.IOError := oeSaisie;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : C.B
Créé le ...... : 28/02/2002
Modifié le ... : 28/02/2002
Description .. : Validation des tâches
Suite ........ : Numérotation automatique des lignes
Suite ........ : Suppression des ressources (TACHERESSOURCE)
Suite ........ : et des lignes de planning (AFPLANNING)
Mots clefs ... : TACHES
*****************************************************************}

function ValideLesTaches(pTob, pTobTache, pTobRes: TOB; pStAffaire: string; pTabRes: TTabRes): Boolean;
var
  i, vInNum: integer;
  vTOBAffaires: Tob;
  vTobArticles: Tob;
  vAFOAssistants: TAFO_Ressources;

  procedure SupprimerTache(pStAffaire, pStNumTache: string);
  var
    vSt: string;

  begin

    vSt := 'Delete From TACHERESSOURCE where ATR_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND ATR_NUMEROTACHE = ' + pStNumTache;

    ExecuteSql(vSt);

    vSt := 'Delete From TACHE where ATA_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND ATA_NUMEROTACHE = ' + pStNumTache;

    ExecuteSql(vSt);
  end;

  procedure SupprimerPlanning(pStAffaire, pStNumTache: string);
  var
    vSt: string;

  begin
    vSt := 'Delete From AFPLANNING where APL_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND APL_NUMEROTACHE = ' + pStNumTache;

    ExecuteSql(vSt);
  end;

  procedure SuppRessource(pArRessource: array of RecordRessource; pStUniteTemps: string);
  var
    vSt: string;
    i: Integer;

  begin

    for i := 0 to length(pArRessource) - 1 do
    begin
      vSt := 'Delete From TACHERESSOURCE where ATR_AFFAIRE = "' + pArRessource[i].rdAffaire.StAffaire + '"';
      vSt := vSt + ' AND ATR_NUMEROTACHE = ' + pArRessource[i].StNumTache;
      vSt := vSt + ' AND ATR_RESSOURCE = "' + pArRessource[i].StOldRes + '"';

      ExecuteSql(vSt);

      vSt := 'Delete From AFPLANNING where APL_AFFAIRE = "' + pArRessource[i].rdAffaire.StAffaire + '"';
      vSt := vSt + ' AND APL_NUMEROTACHE = ' + pArRessource[i].StNumTache;
      vSt := vSt + ' AND APL_RESSOURCE = "' + pArRessource[i].StOldRes + '"';

      ExecuteSql(vSt);

      // C.B 04/05/05
      // mise à jour de la quantité planifiée dans la table tâche maintenant que cette quantité est stockée
//Modif FV UpdatePlanifieTache(pArRessource[i].rdAffaire.StAffaire, pArRessource[i].StNumTache, pStUniteTemps);
    end;
  end;

begin

  Result := True;

  // suppression des taches
  if assigned(pTobTache) then
  begin
    SupprimerTache(pTobTache.GetValue('ATA_AFFAIRE'), pTobTache.GetValue('ATA_NUMEROTACHE'));
    SupprimerPlanning(pTobTache.GetValue('ATA_AFFAIRE'), pTobTache.GetValue('ATA_NUMEROTACHE'));
  end
  else
  begin
    // suppression des ressources
    if assigned(pTobRes) and (length(pTabRes) > 0) then
    begin
      SuppRessource(pTabRes, pTob.GetValue('ATA_UNITETEMPS'));
      pTobRes.DeleteDB;
    end;

    vInNum := GetNumTache(pStAffaire);
    if pTob.GetValue('ATA_NUMEROTACHE') = 0 then
      pTob.PutValue('ATA_NUMEROTACHE', vInNum);

    // maj du num de tâche sur les tâches/ressources si creation
    for i := 0 to pTob.Detail.count - 1 do
    begin
      if pTob.Detail[i].GetValue('ATR_NUMEROTACHE') = 0 then
        pTob.Detail[i].PutValue('ATR_NUMEROTACHE', pTob.GetValue('ATA_NUMEROTACHE'));

      // ajout de la fonction de la tache
      if pTob.Detail[i].GetValue('ATR_FONCTION') <> pTob.GetValue('ATA_FONCTION') then
        pTob.Detail[i].PutValue('ATR_FONCTION', pTob.GetValue('ATA_FONCTION'));
    end;

    vAFOAssistants := TAFO_Ressources.Create;
    vTOBAffaires := TOB.Create('#AFFAIRE', nil, -1);
    vTOBArticles := TOB.Create('#ARTICLE', nil, -1);

    if GetparamSocSecur('SO_BTGESTIONRAF', false) then
    begin // mcd  inutile de faire des recalculs de InitPt* et Ecratµ si pas de gestion du raf
      try
        // valorisation d'une tache et de ses ressources
        if pTob.Modifie or pTob.IsOneModifie then
          ValorisationPlanning(pTob, 'ATA', vAFOAssistants, vTOBAffaires, vTobArticles, true, true);
      finally
        vAFOAssistants.Free;
        vTobArticles.Free;
        vTOBAffaires.Free;
      end;
    end;

    try
      Result := pTob.InsertOrUpdateDB(false);
    except
      on e: exception do
      begin
        result := false;
        PGIBoxAF(TexteMessage[1] + #10#13 + e.Message, texteMessage[2]);
      end;
    end;

  end;
end;

{**********************************************************************************
Auteur  ...... :
Créé le ...... : 13/04/2004
Modifié le ... :   /  /
Description .. : Initialise une nouvelle tâche
Mots clefs ... : nouvelle TobTache , code article par défault, champ de l'affaire
***********************************************************************************}

procedure InitNewTobTache(TobTache: TOB; pCodeArticle: string; pCurAffaire: AffaireCourante);
var
  vStTypeArticle, vStArticle : String;
  vStCodeArticle, vStFacturable : String;
  vStUnite, vStlibelle : String;
  vStStatutPla : String;
  PR, PV : double;
  vStDate : String;

begin

  // C.B 22/06/2006
  // création des champs sup adresse
  TobTache.AddChampSupValeur('ADR_LIBELLE', '', True);
  TobTache.AddChampSupValeur('ADR_LIBELLE2', '', True);
  TobTache.AddChampSupValeur('ADR_ADRESSE1', '', True);
  TobTache.AddChampSupValeur('ADR_ADRESSE2', '', True);
  TobTache.AddChampSupValeur('ADR_ADRESSE3', '', True);
  TobTache.AddChampSupValeur('ADR_CODEPOSTAL', '', True);
  TobTache.AddChampSupValeur('ADR_VILLE', '', True);
  TobTache.AddChampSupValeur('ADR_TELEPHONE', '', True);
  TobTache.AddChampSupValeur('QUELLEADRESSE', '', True); // CLIENT, TACHE, INTERVENTION

  TobTache.InitValeurs;

  // C.B 30/03/2006
  // paramétrage des dates de début de taches
  vStDate := GetParamsocSecur('SO_BTTACHEDATEDEB','');
  if vStDate = 'ACC' then
     TobTache.PutValue('ATA_DATEDEBPERIOD', pCurAffaire.DtDateSigne)
  else if vStDate = 'AFF' then
     TobTache.PutValue('ATA_DATEDEBPERIOD', pCurAffaire.DtDebutAff)
  else if vstDate = 'FAC' then
     TobTache.PutValue('ATA_DATEDEBPERIOD', pCurAffaire.DtDateDebGener);

  // C.B 22/08/2006
  // correction de la date ata_datefinperiod
  vStDate := GetParamsocSecur('SO_BTTACHEDATEFIN','');
  if vStDate = 'AFF' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtFinAff)
  else if vStDate = 'FAC' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtDateFinGener)
  else if vStDate = 'LIM' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtDateLimite)
  else if vStDate = 'CLO' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtDateCloture)
  else if vStDate = 'GAR' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtDateGarantie)
  else if vStDate = 'RES' then
     TobTache.PutValue('ATA_DATEFINPERIOD', pCurAffaire.DtResil);

  TobTache.PutValue('ATA_TIERS', pCurAffaire.StTiers);
  TobTache.PutValue('ATA_AFFAIRE', pCurAffaire.StAffaire);
  TobTache.PutValue('ATA_AFFAIRE0', pCurAffaire.StAff0);
  TobTache.PutValue('ATA_AFFAIRE1', pCurAffaire.StAff1);
  TobTache.PutValue('ATA_AFFAIRE2', pCurAffaire.StAff2);
  TobTache.PutValue('ATA_AFFAIRE3', pCurAffaire.StAff3);
  TobTache.PutValue('ATA_AVENANT', pCurAffaire.StAvenant);
  TobTache.PutValue('ATA_DEVISE', pCurAffaire.StDevise);
  TobTache.PutValue('ATA_PERIODICITE', 'M');

  TobTache.PutValue('ATA_MOISSEMAINE', '1');
  TobTache.PutValue('ATA_MODEGENE', '1');
  TobTache.PutValue('ATA_QTEINTERVENT', '1');
  TobTache.PutValue('ATA_ANNEENB', '1'); //mcd 05/07/2004
  TobTache.PutValue('ATA_FAITPARQUI', 'CAB'); //mcd 05/07/2004

  // ajout des parametres de decalage
  TobTache.PutValue('ATA_NBJOURSDECAL', GetparamSocSecur('SO_BTJOURSDECAL',0));
  TobTache.PutValue('ATA_METHODEDECAL', GetparamSocSecur('SO_BTJOURSPLANIF',''));

  // article
  TobTache.PutValue('ATA_CODEARTICLE', pCodeArticle);
  vStCodeArticle := pCodeArticle;
  if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStunite, vStLibelle, PR, PV, vStStatutPla) then
     begin
     TobTache.PutValue('ATA_TYPEARTICLE', vStTypeArticle);
     if IsTachePrestation (vStTypeArticle) then  //AB-200605 - FQ 12918
        TobTache.PutValue('ATA_UNITETEMPS', vStUnite)
     else
        TobTache.PutValue('ATA_UNITETEMPS', 'J'); // pour frais et marchandise on met à J
     if TobTache.GetValue('ATA_LIBELLETACHE1') = '' then TobTache.PutValue('ATA_LIBELLETACHE1', vStLibelle);
     //C.B 19/09/2006
     if TobTache.GetValue('ATA_STATUTPLA') = '' then TobTache.PutValue('ATA_STATUTPLA', vStStatutPla);
     TobTache.PutValue('ATA_ARTICLE', vStArticle);
     if GetparamSocSecur('SO_AFVALOPR','') = 'ART' then TobTache.PutValue('ATA_PUPR', PR);
     if GetparamSocSecur('SO_AFVALOPV','') = 'ART' then TobTache.PutValue('ATA_PUVENTEHT', PV);
     TobTache.PutValue('ATA_ACTIVITEREPRIS', vStFacturable); // AB-030325-Facturable
     end
  else
     begin
     TobTache.PutValue('ATA_CODEARTICLE', '');
     TobTache.PutValue('ATA_TYPEARTICLE', GetparamSocSecur('SO_BTTYPEARTDEF','PRE'));
     TobTache.PutValue('ATA_ARTICLE', '');
     TobTache.PutValue('ATA_UNITETEMPS', 'J');
     end;

  TobTache.PutValue('ATA_FAMILLETACHE', GetparamSocSecur('SO_BTFAMILLEDEF',''));
  TobTache.PutValue('ATA_DATEANNUELLE', Idate1900);
  TobTache.PutValue('ATA_QTEINITPLA', 0);
  TobTache.PutValue('ATA_QTEINITIALE', 0);

  if GetparamSocSecur('SO_AFGENELIMIT', false) then
    TobTache.putvalue('ATA_RAPPLANNING', 'X')
  else
    TobTache.putvalue('ATA_RAPPLANNING', '-');

  TobTache.PutValue('RAPPTPRCAL', 0);
  TobTache.PutValue('RAPPTPVCAL', 0);
  TobTache.PutValue('QTEAPLANIFIERCALC', 0);

end;

procedure InitNewRessource(vTobRes: Tob; pCurAffaire: AffaireCourante);
begin
  //vTobRes.AddChampSup('ARS_LIBELLE', True);
  vTobRes.putValue('ATR_TIERS', pCurAffaire.StTiers);
  vTobRes.putValue('ATR_AFFAIRE', pCurAffaire.StAffaire);
  vTobRes.putValue('ATR_AFFAIRE0', pCurAffaire.StAff0);
  vTobRes.putValue('ATR_AFFAIRE1', pCurAffaire.StAff1);
  vTobRes.putValue('ATR_AFFAIRE2', pCurAffaire.StAff2);
  vTobRes.putValue('ATR_AFFAIRE3', pCurAffaire.StAff3);
  vTobRes.putValue('ATR_AVENANT', pCurAffaire.StAvenant);
  vTobRes.putValue('ATR_DEVISE', pCurAffaire.StDevise);
  // ajout de la fonction de la tache juste avant l'enregistrement : ATR_FONCTION
  vTobRes.putValue('ATR_ECARTQTEINIT', 0);
  vTobRes.putValue('ATR_POURCENTAGE', 0);
  // les prix sont alimentés lors de la valorisation
  vTobRes.putValue('ATR_STATUTRES', 'ACT');
end;

//*********************** Numérotation des tâches ******************************

function GetNumTache(pStAffaire: string): integer;
var
  Q: Tquery;
begin
  Result := 1;
  if pStAffaire <> '' then
  begin
    Q := nil;
    try
      Q := OpenSQL('SELECT MAX(ATA_NUMEROTACHE) FROM TACHE WHERE ATA_AFFAIRE = "' + pStAffaire + '"', True);
      if not Q.EOF then
        Result := Q.Fields[0].AsInteger + 1;
    finally
      Ferme(Q);
    end
  end;
end;

// **********************  Organisation Tâche / Tob Filles tâchesRessources ****

procedure OrganiseTacheRessource(TobTaches, TobTachesRes: TOB);
var
  i, j: integer;
  TobDetRes: TOB;
  TobDetTache: TOB;

begin
  for i := TobTachesRes.Detail.count - 1 downto 0 do
  begin
    TobDetRes := TobTachesRes.Detail[i];
    for j := 0 to TobTaches.Detail.Count - 1 do
    begin
      TobDetTache := TobTaches.Detail[j];
      if TobDetRes.GetValue('ATR_NUMEROTACHE') = TobDetTache.GetValue('ATA_NUMEROTACHE') then
        TobDetRes.ChangeParent(TobDetTache, -1);
    end;
  end;
end;

function ControleCodeArticle(var pStCodeArticle, pStTypeArticle, pStArticle,
pStFacturable, pStUnite, pStLibelle : string; var PR, PV: double; var pStStatutPla : String): boolean;
var
  vSt           : String;
  vStDim        : string;
  vStCodeArt    : string;
  vStTypeArt    : string;
  vQr           : TQuery;
  vBoStatutPla  : Boolean;

begin

  // contrôle sur code Article
  vBoStatutPla := True;
  vStDim := '';
  if pStCodeArticle <> '' then
  begin
    vStCodeArt := CodeArticleUnique(pStCodeArticle, vStDim, vStDim, vStDim, vStDim, vStDim);
    vStTypeArt := '';

    //C.B 19/09/2006
    vSt := 'SELECT GA_TYPEARTICLE,GA_ACTIVITEREPRISE, GA_PMRP,GA_PVHT,';
    vSt := vSt + 'GA_QUALIFUNITEACT,GA_QUALIFUNITEVTE,GA_LIBELLE ';
    vSt := vSt + 'FROM ARTICLE, ARTICLECOMPL WHERE GA_ARTICLE="' + vStCodeArt + '"';
    vSt := vSt + 'AND GA_ARTICLE = GA2_ARTICLE';

    vQr := OpenSQL(vSt, true);
    if vQr.EOF then
    begin
      ferme(vQr);
      vSt := 'SELECT GA_TYPEARTICLE,GA_ACTIVITEREPRISE, GA_PMRP,GA_PVHT,';
      vSt := vSt + 'GA_QUALIFUNITEACT,GA_QUALIFUNITEVTE,GA_LIBELLE ';
      vSt := vSt + 'FROM ARTICLE WHERE GA_ARTICLE="' + vStCodeArt + '"';
      vQr := OpenSQL(vSt, true);
      vBoStatutPla := False;
    end;

    if not vQr.EOF then
    begin
      vStTypeArt := vQr.findField('GA_TYPEARTICLE').asString;
      pStLibelle := vQr.findField('GA_LIBELLE').asString;
      pStFacturable := vQr.findField('GA_ACTIVITEREPRISE').asString;
      pStUnite := vQr.findField('GA_QUALIFUNITEACT').asString;
      if vBoStatutPla then
        pStStatutPla := vQr.findField('GA2_STATUTPLA').asString;

      PR := vQr.findField('GA_PMRP').asfloat;
      PV := vQr.findField('GA_PVHT').asfloat;
      pStTypeArticle := vStTypeArt;
      result := true;
    end
    else
      Result := false;

    ferme(vQr);
    pStArticle := CodeArticleUnique(vStCodeArt, '', '', '', '', '');
  end
  else
  begin
    pStArticle := '';
    result := false;
  end;
end;

procedure GetArticle(pStArticle: string; var pArticle: RecordArticle);
var
  vQr: TQuery;

begin
  vQr := OpenSQL('SELECT GA_TYPEARTICLE, GA_CODEARTICLE, GA_ACTIVITEREPRISE FROM ARTICLE WHERE GA_ARTICLE="' + pStArticle + '"', true);
  if not vQr.EOF then
  begin
    pArticle.StArticle := pStArticle;
    pArticle.StCodeArticle := vQr.findField('GA_CODEARTICLE').asString;
    pArticle.StTypeArticle := vQr.findField('GA_TYPEARTICLE').asString;
  end;
  ferme(vQr);
end;

function ExisteUtilisateur(pStCodeRes: string): string;
var
  vQr: TQuery;
begin
  Result := '';
  if pStCodeRes = '' then
    Exit;
  vQR := OpenSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE = "' + pStCodeRes + '"', true);
  try
    if not vQR.Eof then
      Result := vQR.FindField('ARS_RESSOURCE').AsString
    else
      Result := ''
  finally
    ferme(vQR);
  end;
end;

function ExisteTache(pStSqlWhere: string): Boolean;
var
  vSt: string;
begin
  {
     vSt := 'SELECT ATA_AFFAIRE FROM TACHE, AFFAIRE ';
     vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE AND ' + pStSqlWhere;
     }
  vSt := 'SELECT ATA_AFFAIRE FROM TACHE, AFFAIRE,PIECE ';
  vSt := vSt + ' WHERE ATA_AFFAIRE = AFF_AFFAIRE ';
  vSt := vSt + ' AND AFF_AFFAIRE = GP_AFFAIRE';
  vSt := vSt + ' AND GP_NATUREPIECEG = AFF_NATUREPIECEG  AND '; //GM à verifier
  vSt := vSt + pStSqlWhere;
  result := existeSql(vSt);
end;
 
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 20/12/2002
Modifié le ... :
Description .. : test d'existance de ressources pour cette affaire

Mots clefs ... :
*****************************************************************}
function ExisteRessources(pStSqlWhere: string): boolean;
var
  vStAffaire : string;
  vStSqlWhere : String;
  vSt : string;
  i : Integer;
  vBoModif : Boolean;
  vStSqlRef : String;

begin

  if Pos('AFFAIRE =', pStSqlWhere) <> 0 then
  begin
    vBoModif := False;
		vStSqlRef := pStSqlWhere;

    //C.B 12/06/2006
    //suppression des références à la table gp
		if pos('AND GP_PLANIFIABLE="X"', vStSqlRef) > 0 then
    begin
    	vStSqlWhere := copy(vStSqlRef, 0, pos('AND GP_PLANIFIABLE="X"', vStSqlRef) - 1);
    	vStSqlWhere := vStSqlWhere + copy(vStSqlRef, pos('AND GP_PLANIFIABLE="X"', vStSqlRef) + length('AND GP_PLANIFIABLE="X"'), length(vStSqlRef));
      vBoModif := True;
    end
		else if pos('AND GP_PLANIFIABLE<>"X"', vStSqlRef) > 0 then
    begin
    	vStSqlWhere := copy(vStSqlRef, 0, pos('AND GP_PLANIFIABLE<>"X"', vStSqlRef) - 1);
    	vStSqlWhere := vStSqlWhere + copy(vStSqlRef, pos('AND GP_PLANIFIABLE<>"X"', vStSqlRef) + length('AND GP_PLANIFIABLE<>"X"'), length(vStSqlRef));
      vBoModif := True;
    end;

    if vBoModif then
      vStSqlRef := vStSqlWhere;

		if pos('AND GP_VIVANTE="X"', vStSqlRef) > 0 then
    begin
    	vStSqlWhere := copy(vStSqlRef, 0, pos('AND GP_VIVANTE="X"', vStSqlRef) - 1);
    	vStSqlWhere := vStSqlWhere + copy(vStSqlRef, pos('AND GP_VIVANTE="X"', vStSqlRef) + length('AND GP_VIVANTE="X"'), length(vStSqlRef));
      vBoModif := True;
    end
		else if pos('AND GP_VIVANTE<>"X"', vStSqlRef) > 0 then
    begin
    	vStSqlWhere := copy(vStSqlRef, 0, pos('AND GP_VIVANTE<>"X"', vStSqlRef) - 1);
    	vStSqlWhere := vStSqlWhere + copy(vStSqlRef, pos('AND GP_VIVANTE<>"X"', vStSqlRef) + length('AND GP_VIVANTE<>"X"'), length(vStSqlRef));
      vBoModif := True;
    end;

		if not vBoModif then
			vStSqlWhere := pStSqlWhere;

    i := Pos('AFF_AFFAIRE =', vStSqlWhere);
    if i <> 0 then
    begin
      vSt := 'SELECT ATR_AFFAIRE FROM TACHERESSOURCE, AFFAIRE, TIERS ';
      vSt := vSt + 'WHERE ATR_AFFAIRE = AFF_AFFAIRE ';
      vSt := vSt + 'AND AFF_TIERS = T_TIERS ';
      vSt := vSt + ' AND ' + vStSqlWhere;
      result := existeSql(vSt);
    end
    else
    begin
      i := Pos('APL_AFFAIRE =', pStSqlWhere);
      if i <> 0 then
      begin
        vStAffaire := copy(pStSqlWhere, i + 13, length(pStSqlWhere));
        vSt := 'SELECT ATR_AFFAIRE FROM TACHERESSOURCE ';
        // C.B 30/11/2006
        // cette ligne est inutile dans un exist
        // vSt := vSt + ' LEFT OUTER JOIN AFPLANNING ON ATR_AFFAIRE = APL_AFFAIRE ';
        vSt := vSt + ' WHERE ATR_AFFAIRE = ' + vStAffaire;
        result := existeSql(vSt);
      end   
      else
        result := true;
    end;
  end
  else
    result := true;
end;

function RessourceLib(pStCodeRes: string): string;
var
  vQr: TQuery;
begin
  if pStCodeRes = '' then
    Exit;
  vQr := nil;
  try
    vQR := OpenSQL('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE = "' + pStCodeRes + '"', true);
    Result := vQR.FindField('ARS_LIBELLE').AsString;
  finally
    ferme(vQR);
  end;
end;

function RessourceNomPrenom(pStCodeRes: string): string;
var
  vQr: TQuery;
begin
  if pStCodeRes = '' then
    Exit;
  vQr := nil;
  try
    vQR := OpenSQL('SELECT ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_RESSOURCE = "' + pStCodeRes + '"', true);
    Result := vQR.FindField('ARS_LIBELLE').AsString;
    if trim(vQR.FindField('ARS_LIBELLE2').AsString) <> '' then
      Result := Result + '  ' + vQR.FindField('ARS_LIBELLE2').AsString;
  finally
    ferme(vQR);
  end;
end;

procedure DecodeArgument(pStTmp: string; var pStChamp, pStValeur: string);
var
  vInPos: Integer;
begin
  vInPos := pos(':', pStTmp);
  if vInPos = 0 then
    vInPos := pos('=', pStTmp);
  if vInPos <> 0 then
  begin
    pStChamp := copy(pStTmp, 1, vInPos - 1);
    pStValeur := Copy(pStTmp, vInPos + 1, length(pStTmp) - vInPos);
  end
  else
    pStChamp := pStTmp;
end;

function ExisteTacheRessource(pStAffaire, pStTache, pStRessource: string): Boolean;
var
  vQR: TQuery;
  vSt: string;
begin
  vSt := 'SELECT ATR_RESSOURCE FROM TACHERESSOURCE ';
  vSt := vSt + ' WHERE ATR_RESSOURCE =  "' + pStRessource + '"';
  vSt := vSt + ' AND ATR_AFFAIRE =  "' + pStAffaire + '"';
  vSt := vSt + ' AND ATR_NUMEROTACHE =  ' + pStTache;
  vQr := nil;
  try
    vQr := OpenSQL(vSt, True);
    result := not vQr.Eof;
  finally
    ferme(vQr);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 18/06/2002
Modifié le ... : 13/02/2004
Description .. : Calcul de la quantité Affecté pour une tache a partir
Suite ........ : Recoit l'affaire, le n° tache, le champ à lire dans afplanning
Suite ........ : (qte, mtt)
Suite ........ : la ressource, le type de planning (PC ou PLA)
Suite ........ : et si Pbotout = true : on limite la date début > à la date du
Suite ........ : mois en cours (si PC) et on prend que les enrgt dont
Suite ........ : ressource <> blanc
Suite ........ : (mais tout !!)
Suite ........ : Fi Pbotout=False, on ne prend que les enrgt de la ressource
Suite ........ : passée
Suite ........ : sinon, pas de restriction de date, mais sélection sur la
Suite ........ : ressource si ressource passée
Mots clefs ... :
*****************************************************************}

function CalculPlanifie(pStAffaire, pStNumeroTache, pStChamp, pStRessource, TypePla: string; pBoTout: Boolean): Double;
var
  S: string;
  vDtDepart: TDateTime;
  vQr: TQuery;
  vTob: Tob;

begin
  if not (GetparamSocSecur('SO_BTRAFPLANNING', false)) and (typepla = 'PC') then
  begin
    result := 0; // mcd mis dans cette fct, plutôt que de le faire à chaque fois
    exit;
  end;

  // chargement des mois pour la ligne fonction (pas de ressource affectée) de cette tache
  S := 'SELECT sum(' + pStChamp + ') AS QTE FROM AFPLANNING ';
  S := S + ' where APL_TYPEPLA="' + TypePla + '" AND APL_AFFAIRE = "' + pStAffaire + '"';
  if pStNumeroTache <> '' then
    S := S + ' and APL_NUMEROTACHE = ' + pStNumeroTache;
  // prise en compte de la date du mois que sur PC ... à voir si modif traitement PC
  if (not pBoTout) and (typepla = 'PC') then
  begin
    vDtDepart := DebutDeMois(now);
    S := S + ' and APL_DATEDEBPLA >= "' + UsDateTime(vDtDepart) + '"';
  end;

  // on ne calcule que le planning des ressources
(* mcd il faut otut prendr, avec et sans ressource (la ressource n'est pas obligatorie si pe pas de valo par tache
  et doit être OK maintenant que l'on n'écrit plus enrgt tache avec ress à blanc et le détail
    if pBoTout then
    S := S + ' and APL_RESSOURCE <> ""'
  else  *)

  if (not pBoTout) and (pStRessource <> '') then
    S := S + ' and APL_RESSOURCE = "' + pStRessource + '"';

  vQr := nil;
  vTob := Tob.Create('qteaffecte', nil, -1);
  try
    vQr := OpenSql(S, True);
    if not vQr.Eof then
    begin
      vTob.LoadDetailDB('qteaffecte', '', '', vQr, False, True);
      if vTob.Detail[0].GetValue('QTE') = NULL then
        result := 0
      else
        result := valeur(vTob.Detail[0].GetValue('QTE'));
    end
    else
      Result := -1;

  finally
    if vQr <> nil then
      Ferme(vQr);
    vTob.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : chargement du realise saisi dans l'activite
                 C.B 04/12/2003
                 Ajout de la règle suivante :
                 - si Plan de Charge, on ne prend pas en compte le réalisé du mois courant
                 - sinon, on prend tout le réalisé
pStChamp : nom du champ à lire dans activité
MCd 26/02/04 :passe unité pour convertir dans unit de la tache !!!
Attention : article = champ Article (et non pas codearticle)
Mots clefs ... :
*****************************************************************}
function CalculRealise(pStAffaire, pStFonction, pStChamp, pStRessource,
pStArticle, Unite: string; pBoPDC: Boolean): Double;
var
  S: string;
  vQrRealise: TQuery;
  vTobRealise: Tob;
  vStFamille: string;
  vDate: TDateTime;
begin
  vStFamille := '';
  result := 0;  
  if GetparamSocSecur('SO_BTLIENARTTACHE','') = 'FAM' then
  begin
    // on va recherche e code famille 1 associé à l'artilce (Haden ??)
    S := 'SELECT GA_FAMILLENIV1 FROM ARTICLE WHERE GA_ARTICLE = "' + pStArticle + '"';
    vQrRealise := nil;
    try
      vQrRealise := OpenSql(S, True);
      if not vQrRealise.Eof then
        vStFamille := vQrRealise.FindField('GA_FAMILLENIV1').asString
      else
      begin
        Result := 0;
        exit;
      end;
    finally
      if vQrRealise <> nil then
        Ferme(vQrRealise);
    end;
  end;
  vTobRealise := TOB.create('QteRealise', nil, -1);
  // chargement du realise de l'activite
  S := 'SELECT SUM(' + pStChamp + ') AS QTE, ACT_UNITE ';
  S := S + ' FROM ACTIVITE, ARTICLE ';
  S := S + ' WHERE ACT_AFFAIRE = "' + pStAffaire + '"';
  if GetparamSocSecur('SO_AFCLIENT',0) = cInClientAlgoe then
    S := S + ' AND GA_FAMILLENIV1 <> "PRO"';
  S := S + ' AND GA_ARTICLE = ACT_ARTICLE';
  if GetparamSocSecur('SO_BTLIENARTTACHE','') = 'FAM' then
    S := S + ' AND GA_FAMILLENIV1 = "' + vStFamille + '"'
  else
    S := S + ' AND GA_ARTICLE = "' + pStArticle + '"';
  S := S + ' AND ACT_ETATVISA="VIS" ';
  if GetparamSocSecur('SO_AFCLIENT',0) = cInClientAlgoe then //mcd 18/08/04 ajout du test pour spécif Algoe sur cette valeur
    S := S + ' AND ACT_ACTIVITEREPRIS <> "A"';
  if pStRessource <> '' then
    S := S + ' AND ACT_RESSOURCE = "' + pStRessource + '"';
  (* mcd ??? on passe le code article, donc on ne doit rien faire sur le type !!!
    // article montant(frais ...)
  if pStChamp = 'ACT_TOTPR' then S := S + ' AND ACT_TYPEARTICLE <> "PRE"';
    // article prestation
  if pStChamp = 'ACT_QTE'   then S := S + ' AND ACT_TYPEARTICLE = "PRE"';
  *)

  // Avec la gestion du reste, on prend le réalisé antérieur au mois en cours .
  if pBoPDC and GetparamSocSecur('SO_BTGESTIONRAF', false) and GetparamSocSecur('SO_BTRAFPLANNING', false) then
  begin
    vDate := DebutdeMois(now);
    S := S + ' AND ACT_DATEACTIVITE < "' + usDateTime(vDate) + '"';
  end;
  S := S + ' GROUP BY ACT_UNITE ';
  vQrRealise := nil;
  try
    vQrRealise := OpenSql(S, True);
    if not vQrRealise.Eof then
    begin
      vTobRealise.LoadDetailDB('QteRealise', '', '', vQrRealise, False, True);
      if vTobRealise.Detail[0].GetValue('QTE') = NULL then
        result := 0
      else
        // Convertir les unites dans l'unité passé en paramètre
        if pStChamp = 'ACT_QTE' then
//Modif FV result := AFConversionUnite(vTobRealise.Detail[0].GetValue('ACT_UNITE'), Unite, valeur(vTobRealise.Detail[0].GetValue('QTE')))
      else
        result := valeur(vTobRealise.Detail[0].GetValue('QTE'));
    end
    else
      Result := 0;
  finally
    if vQrRealise <> nil then
      Ferme(vQrRealise);
    vTobRealise.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/05/2002
Modifié le ... :   /  /
Description .. : chargement de tout le realise
Mots clefs ... :
*****************************************************************}
//C.B 26/01/2007 fonction manifestetment plus utilisée depuis longtemps
{function CalculToutRealise(pStAffaire, pStFonction, pStChamp, pStRessource, pStArticle: string): Double;
var
  S: string;
  vQrRealise: TQuery;
  vTobRealise: Tob;

begin

  vTobRealise := TOB.create('QteRealise', nil, -1);

  // chargement du realise de l'activite
  S := 'SELECT SUM(' + pStChamp + ') AS QTE, ACT_UNITE ';
  S := S + ' FROM ACTIVITE, ARTICLE ';
  S := S + ' WHERE ACT_AFFAIRE = "' + pStAffaire + '"';

  if GetparamSocSecur('SO_AFCLIENT', 0) = cInClientAlgoe then
    S := S + ' AND GA_FAMILLENIV1 <> "PRO"';

  S := S + ' AND GA_ARTICLE = ACT_ARTICLE';
  S := S + ' AND GA_ARTICLE = "' + pStArticle + '"';
  S := S + ' AND ACT_ETATVISA = "VIS" ';
  //  if pStFonction <> '' then S := S + ' AND ACT_FONCTIONRES = "' + pStFonction + '"';

  S := S + ' AND ACT_ACTIVITEREPRIS <> "A"';

  if pStRessource <> '' then
    S := S + ' AND ACT_RESSOURCE = "' + pStRessource + '"';

  if pStChamp = 'ACT_TOTPR' then
    S := S + ' AND ACT_TYPEARTICLE <> "PRE" AND ACT_TYPEARTICLE <> "CTR"';  //AB-200605 - FQ 12918
  if pStChamp = 'ACT_QTE' then
    S := S + ' AND (ACT_TYPEARTICLE = "PRE" OR ACT_TYPEARTICLE = "CTR")';

  S := S + ' GROUP BY ACT_UNITE ';

  vQrRealise := OpenSql(S, True);

  try

    if not vQrRealise.Eof then
    begin
      vTobRealise.LoadDetailDB('QteRealise', '', '', vQrRealise, False, True);
      if vTobRealise.Detail[0].GetValue('QTE') = NULL then
        result := 0
      else
        // Convertir les unites de l'activite en jour
        result := AFConversionUnite(vTobRealise.Detail[0].GetValue('ACT_UNITE'), 'J', valeur(vTobRealise.Detail[0].GetValue('QTE')));
    end
    else
      Result := 0;

  finally
    if vQrRealise <> nil then
      Ferme(vQrRealise);
    vTobRealise.Free;
  end;
end;}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : calcul du RAF en quantite avec les options suivantes :
                  tenir compte ou non du planifie
                  tenir compte ou non des temps (realise)
Mots clefs ... :
*****************************************************************}

function CalculRAFQte(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
var
  vRdRealise: double;
  vRdRes: double;
  vRdPlanifie: double;
  vStArticle: string;
  vStRessource, Unite: string;

begin

  // ressource et article
  if pStPrefixe = 'ATR' then
  begin
    vStArticle := pTob.Parent.GetValue('ATA_ARTICLE');
    Unite := pTob.Parent.GetValue('ATA_UNITETEMPS');
    vStRessource := pTob.GetValue(pStPrefixe + '_RESSOURCE');
  end
  else
  begin
    vStArticle := pTob.GetValue(pStPrefixe + '_ARTICLE');
    Unite := pTob.GetValue(pStPrefixe + '_UNITETEMPS');
    vStRessource := '';
  end;

  // planifie Mcd en PC ....
  vRdPlanifie := CalculPlanifie(pTob.getValue(pStPrefixe + '_AFFAIRE'), pTob.getValue(pStPrefixe + '_NUMEROTACHE'),
    'APL_QTEPLANIFIEE', vStRessource, 'PC', False);

  // realise
  vRdRealise := CalculRealise(pTob.getValue(pStPrefixe + '_AFFAIRE'), pTob.getValue(pStPrefixe + '_FONCTION'),
    'ACT_QTE', vStRessource, vStArticle, Unite, pBoPDC);

  vRdRes := valeur(pTob.getValue(pStPrefixe + '_QTEINITIALE')) + valeur(pTob.GetValue(pStPrefixe + '_ECARTQTEINIT'));
  vRdRes := vRdRes - vRdRealise - vRdPlanifie;
  if (GetparamSocSecur('SO_AFCLIENT', 0) = cInClientAlgoe) and (vRdRes < 0) then
    vRdRes := 0;
  // C.B 17/09/02 ajout du formatage du RAP
  // C.B 26/10/2006 remplaement de floattostr par valeur
  result := valeur(formatFloat('0.##', vRdRes));

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 17/07/2002
Modifié le ... :   /  /
Description .. : calcul du RAF en montant utilisé dans le plan de charge
Mots clefs ... :
*****************************************************************}

function CalculRAFMnt(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
var
  vRdRealise: double;
  vRdPlanifie: double;
  vStArticle: string;

begin

  if pStPrefixe = 'ATR' then
    vStArticle := pTob.Parent.GetValue('ATA_ARTICLE')
  else
    vStArticle := pTob.GetValue(pStPrefixe + '_ARTICLE');

  // realise
  vRdRealise := CalculRealise(pTob.getValue(pStPrefixe + '_AFFAIRE'), '', 'ACT_TOTPR', '', vStArticle, '', pBoPDC);

  //planifie  MCd utilisé depuis ChargeLesATche, LoadQte :==> uniquement pour le PC
  vRdPlanifie := CalculPlanifie(pTob.getValue(pStPrefixe + '_AFFAIRE'), pTob.getValue(pStPrefixe + '_NUMEROTACHE'), 'APL_INITPTPR', '', 'PC', true);

  Result := valeur(pTob.getValue(pStPrefixe + '_INITPTPR')) - vRdPlanifie
    - vRdRealise + pTob.GetValue(pStPrefixe + '_ECARTPR');

end;

function CalculRAFMntPV(pTob: Tob; pStPrefixe: string; pBoPDC: Boolean): Double;
var
  vRdRealise: double;
  vRdPlanifie: double;
  vStArticle: string;
begin
  // calcul du mtt RAF du PC, mais fait en PV
  if pStPrefixe = 'ATR' then
    vStArticle := pTob.Parent.GetValue('ATA_ARTICLE')
  else
    vStArticle := pTob.GetValue(pStPrefixe + '_ARTICLE');
  // realise
  vRdRealise := CalculRealise(pTob.getValue(pStPrefixe + '_AFFAIRE'), '', 'ACT_TOTVENTE', '', vStArticle, '', pBoPDC);
  //planifie  MCd utilisé depuis ChargeLesATche, LoadQte :==> uniquement pour le PC
  vRdPlanifie := CalculPlanifie(pTob.getValue(pStPrefixe + '_AFFAIRE'), pTob.getValue(pStPrefixe + '_NUMEROTACHE'), 'APL_INITPTVENTEHT', '', 'PC', true);

  Result := valeur(pTob.getValue(pStPrefixe + '_INITPTVENTEHT')) - vRdPlanifie
    - vRdRealise + pTob.GetValue(pStPrefixe + '_ECARTPVHT');
end;

function calculPRArticle(pStArticle: string): Double;
var
  vSt: string;
  vQr: TQuery;
begin

  // Article on utilise le PMRP (prix moren pondéré car s'est le seul saisi)
  vSt := 'SELECT GA_PMRP FROM ARTICLE WHERE GA_ARTICLE = "' + pStArticle + '"';

  vQr := nil;
  try
    vQR := OpenSQL(vSt, True);
    if (not vQR.EOF) then
      result := vQR.findField('GA_PMRP').AsFloat
    else
      result := -1;
  finally
    ferme(vQr);
  end;
end;

procedure TOBCopieModeleTache(pCurAffaire: AffaireCourante; FromTOB, ToTOB: TOB);
var
  i_pos, i_ind1: integer;
  FieldNameTo, FieldNameFrom, St: string;
  PrefixeTo, PrefixeFrom: string;
  vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle: string;
  vStStatutPla : String;
  PR, PV: double;
begin

  PrefixeFrom := TableToPrefixe(FromTOB.NomTable);
  PrefixeTo := TableToPrefixe(ToTOB.NomTable);

  if PrefixeFrom = '' then PrefixeFrom := 'GL'; // on n'a pas la table totale des lignes, donc pas de préfixe

  if PrefixeFrom = 'GL' then
    begin
    ToTOB.PutValue('ATA_ACTIVITEREPRIS', FromTOB.GetValue('GL_FACTURABLE')); //mcd 05/10/2004 11493
    ToTOB.PutValue('ATA_LIBELLETACHE1', FromTOB.GetValue('GL_LIBELLE'));
    ToTOB.PutValue('ATA_CODEARTICLE', FromTOB.GetValue('GL_CODEARTICLE'));
    ToTOB.PutValue('ATA_TYPEARTICLE', FromTOB.GetValue('GL_TYPEARTICLE'));
    ToTOB.PutValue('ATA_UNITETEMPS', FromTOB.GetValue('GL_QUALIFQTEVTE'));
    ToTOB.PutValue('ATA_ACTIVITEREPRIS', FromTOB.GetValue('GL_FACTURABLE')); //AB-200510-
    if IsTachePrestation(ToTob.getValue('ATA_TYPEARTICLE')) then //AB-200605 - FQ 12918
      // quantité initiale du plan de charge
      ToTOB.PutValue('ATA_QTEINITIALE', FromTOB.GetValue('GL_QTEFACT'))
    else
      begin
      if GetparamSocSecur('SO_BTGESTIONRAF', false) then
        begin
        //mcd 27/09/04 ajout test gestion RAF.. sinon ces zones ne sont pas à renseignée...
        ToTOB.PutValue('ATA_INITPTVENTEHT', FromTOB.GetValue('GL_TOTALHT'));
        ToTOB.PutValue('ATA_INITPTVTDEVHT', FromTOB.GetValue('GL_TOTALHTDEV'));

        //mcd 24/09/04 11493 ajout PR    .. je considère que DEV= mttt normal.. pour le PR,
        // pas de notion de devise... de plus devise pas gérer dans planning .. si pb à voir ..
        ToTOB.PutValue('ATA_INITPTPR', FromTOB.GetValue('GL_PMRP') * FromTOB.GetValue('GL_QTEFACT'));
        ToTOB.PutValue('ATA_INITPTPRDEV', FromTOB.GetValue('GL_PMRP') * FromTOB.GetValue('GL_QTEFACT'));
        end;
      end;
    if (GetModeValPR = cInValTache) then //AB-200510-
      begin
      ToTOB.PutValue('ATA_PUPR', FromTOB.GetValue('GL_PMRP'));
      ToTOB.PutValue('ATA_PUPRDEV', FromTOB.GetValue('GL_PMRP'));
      end;
    if (GetModeValPV = cInValTache) then //AB-200510-
      begin
      ToTOB.PutValue('ATA_PUVENTEHT', FromTOB.GetValue('GL_PUHT'));
      ToTOB.PutValue('ATA_PUVENTEDEVHT', FromTOB.GetValue('GL_PUHTDEV'));
      end;
    ToTOB.PutValue('ATA_QTEINITPLA', FromTOB.GetValue('GL_QTEFACT'));
    ToTOB.PutValue('ATA_IDENTLIGNE', EncodeRefPiece(FromTOB));
    ToTOB.AddChampSupValeur('IDENTLIGNE', 'X');
    end
  else
    for i_ind1 := 1 to FromTOB.NbChamps do
      begin
      FieldNameFrom := FromTOB.GetNomChamp(i_ind1);
      St := FieldNameFrom;
      i_pos := Pos('_', St);
      System.Delete(St, 1, i_pos - 1);
      FieldNameTo := PrefixeTo + St;
      if ToTOB.FieldExists(FieldNameTo) then
         ToTOB.PutValue(FieldNameTo, FromTOB.GetValue(FieldNameFrom))
      end;

  ToTOB.PutValue('ATA_TIERS', pCurAffaire.StTiers);
  ToTOB.PutValue('ATA_AFFAIRE', pCurAffaire.StAffaire);
  ToTOB.PutValue('ATA_AFFAIRE0', pCurAffaire.StAff0);
  ToTOB.PutValue('ATA_AFFAIRE1', pCurAffaire.StAff1);
  ToTOB.PutValue('ATA_AFFAIRE2', pCurAffaire.StAff2);
  ToTOB.PutValue('ATA_AFFAIRE3', pCurAffaire.StAff3);
  ToTOB.PutValue('ATA_AVENANT', pCurAffaire.StAvenant);
  ToTOB.PutValue('ATA_DEVISE', pCurAffaire.StDevise);

  if (ToTOB.getvalue('ATA_DATEDEBPERIOD') < pCurAffaire.DtDebutAff) or
     (ToTOB.getvalue('ATA_DATEDEBPERIOD') > pCurAffaire.DtFinAff) then
      ToTOB.putvalue('ATA_DATEDEBPERIOD', pCurAffaire.DtDebutAff);

  if (ToTOB.getvalue('ATA_DATEFINPERIOD') > pCurAffaire.DtFinAff) or
     (ToTOB.getvalue('ATA_DATEFINPERIOD') < pCurAffaire.DtDebutAff) then
      ToTOB.putvalue('ATA_DATEFINPERIOD', pCurAffaire.DtFinAff);

  vStCodeArticle := ToTOB.GetValue('ATA_CODEARTICLE');
  if controleCodeArticle(vStCodeArticle, vStTypeArticle, vStArticle, vStFacturable, vStUnite, vStLibelle, PR, PV, vStStatutPla) then
     begin
     ToTOB.PutValue('ATA_TYPEARTICLE', vStTypeArticle);
     ToTOB.PutValue('ATA_ARTICLE', vStArticle);
     if PrefixeFrom <> 'GL' then //AB-200510-
        ToTOB.PutValue('ATA_ACTIVITEREPRIS', vStFacturable);
     if ToTOB.GetValue('ATA_LIBELLETACHE1') = '' then
        ToTOB.PutValue('ATA_LIBELLETACHE1', vStLibelle);
     //C.B 19/09/2006
     if ToTOB.GetValue('ATA_STATUTPLA') = '' then
        ToTOB.PutValue('ATA_STATUTPLA', vStStatutPla);

     if GetparamSocSecur('SO_AFVALOPR','') = 'ART' then
        ToTOB.PutValue('ATA_PUPR', PR);
     if GetparamSocSecur('SO_AFVALOPV','') = 'ART' then
        ToTOB.PutValue('ATA_PUVENTEHT', PV);
     if IsTachePrestation (vStTypeArticle) then //AB-200605 - FQ 12918
        ToTOB.PutValue('ATA_UNITETEMPS', vStUnite)
     else
      begin
      ToTOB.PutValue('ATA_UNITETEMPS', 'J'); // pour frais et marchandise on met à J
      if PrefixeFrom = 'GL' then
         begin
         ToTOB.PutValue('ATA_QTEINITIALE', 0);
         ToTOB.PutValue('ATA_QTEINITPLA', 0);
         end;
      end;
  end
end;

{***********A.G.L.****************************************************************
Auteur  ...... : A Buchet
Créé le ...... : 21-03-2005
Modifié le ... :
Description .. : Saisie des Intervenants d'une tâche
                 Appel depuis la tache ou du planning
Mots clefs ... : Affaire - Numéro de Tache - Identligne - Lib Tache - Action
                 AFT_TYPEORIG : TAC - clé :affaire + Numéro de tache
                 AFT_TYPEORIG : LIG - clé :affaire + Numéro d'ordre ligne pièce
***********************************************************************************}

{$IFNDEF EAGLSERVER}
procedure SaisieIntervenantsTache(pStAffaire, pStNumeroTache, pStIdentligne, pStLibTache, pStAction: string);
var
  vStRange, vStArgument, vStSql: string;
  vStNumOrig, vStTypeOrig, vStTypeInterv, vStIdentligne: string;
  CleDoc: R_CleDoc;
  Q: TQuery;
  vBAffaireEcole: boolean;
begin
  vStNumOrig := pStNumeroTache;
  vStTypeOrig := 'TAC';
  vStTypeInterv := 'TYPEINTERV=CON;';
  vBAffaireEcole := false;
  if (GetparamSocSecur('SO_BTMODEPLANNING','') = 'ART') then
  begin
    vStIdentligne := pStIdentligne;
    Q := OpenSQL('SELECT ATA_IDENTLIGNE,AFF_ISECOLE FROM TACHE,AFFAIRE WHERE AFF_AFFAIRE="' + pStAffaire + '" AND ATA_AFFAIRE="' + pStAffaire + '" AND ATA_NUMEROTACHE=' + pStNumeroTache, true);
    if not Q.EOF then
    begin
      vStIdentligne := Q.Fields[0].AsString;
      vBAffaireEcole := (Q.Fields[1].AsString = 'X');
    end;
    ferme(Q);
    if trim(vStIdentligne) <> '' then
    begin
      vStTypeOrig := 'LIG';
      DecodeRefPiece(vStIdentligne, CleDoc);
      vStNumOrig := IntToStr(CleDoc.NumOrdre);
      if vBAffaireEcole then
        vStTypeInterv := 'TYPEINTERV=RES;';
    end;
  end;

  vStRange := pStAffaire + ';' + vStTypeOrig + ';' + vStNumOrig + ';';
  vStArgument := vStTypeOrig + ';AFFAIRE=' + pStAffaire + ';NUMERO=' + vStNumOrig + ';';
  if trim(pStLibTache) <> '' then
    vStArgument := vStArgument + 'LIBELLE=' + pStLibTache + ';';
  vStArgument := vStArgument + vStTypeInterv + pStAction + ';';

//Modif FV AFLanceFiche_AFIntervenant(vStRange, '', vStArgument);

  if (pStNumeroTache <> '') then
  begin
    vStSql := 'UPDATE TACHE SET ATA_NBPARTINSCRIT = ';
    vStSql := vStSql + ' (SELECT COUNT(AFT_RANG) FROM AFFTIERS WHERE AFT_AFFAIRE="' + pStAffaire + '" AND AFT_TYPEORIG="' + vStTypeOrig + '" AND AFT_NUMORIG=' + vStNumOrig + ')';
    vStSql := vStSql + ' WHERE  ATA_AFFAIRE = "' + pStAffaire + '"';
    vStSql := vStSql + ' AND ATA_NUMEROTACHE = ' + pStNumeroTache;
    ExecuteSql(vStSql);
  end;
end;
{$ENDIF EAGLSERVER}


function FamilleTacheLie : String;
var
  vSt 	: String;
  vTob 	: Tob;

begin
  result := '';
 	vSt := 'SELECT YLT_ORIGINETAB FROM YLIENTABLETTE WHERE YLT_DESTINATIONTAB = "AFFAMILLETACHE"';
  vTob := TOB.Create('',Nil,-1);
  Try
	 	vTob.LoadDetailFromSQL(vSt);
    if vTob.detail.count = 1 then
    begin
    	if vTob.Detail[0].GetString('YLT_ORIGINETAB') = 'GCFAMILLENIV1' then
      	result := 'GA_FAMILLENIV1'
    	else if vTob.Detail[0].GetString('YLT_ORIGINETAB') = 'GCFAMILLENIV2' then
        result := 'GA_FAMILLENIV2'
    	else if vTob.Detail[0].GetString('YLT_ORIGINETAB') = 'GCFAMILLENIV3' then
        result := 'GA_FAMILLENIV3'
      else if vTob.Detail[0].GetString('YLT_ORIGINETAB') = 'GCCOMPTAARTICLE' then
				result := 'GA_COMPTAARTICLE';
    end;
  Finally
    FreeAndNil(vTob);
  end;
end;

// C.B 03/04/2006
// si tablettes identiques entre article et tache, on initialise la
// tablette de la tâche en création
function FindFamilleTacheLie(pStCodeArticle : String) : String;
var
  vSt 		: String;
  vTob 		: Tob;
  vStFam	: String;
   
begin
                             
	vStFam := familleTacheLie;
  if vStFam <> '' then
  begin
    vSt := 'SELECT ' + vStFam + ' FROM ARTICLE WHERE GA_CODEARTICLE = "' + pStCodeArticle + '"';
    vTob := TOB.Create('',Nil,-1);
    Try
      vTob.LoadDetailFromSQL(vSt);
      if vTob.detail.count = 1 then
        result := vTob.Detail[0].GetString(vStFam);
    Finally
      FreeAndNil(vTob);
    end;
	end;
end;

//AB-200605 - FQ 12918 - les articles de type prestation et contrat sont gérés comme des prestations
function IsTachePrestation(pTypeArt: string): boolean;
begin
  result := (pTypeArt = 'PRE') or (pTypeArt ='CTR');
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 04/11/2005
Modifié le ... :
Description .. : Chargement de l'adresse de la tache ou de l'intervention
                 principe : on recherche l'adresse au niveau de la tache
                 si elle est définie
                 si non, on recherche l'adresse de l'affaire
                 si non, on recherche l'adresse du client
                 pStNiveauRecherche : 2 niveaux de recherche
                 intervention ou tache

                 26/10/2006 en création de tache
                            on recherche l'adresse de l'affaire seulement en création de tâche
                            donc Boolean CreateTache 
Mots clefs ... :
*****************************************************************}
function LoadAdresse(pBoCreateTache : Boolean; pStNiveauRecherche, pStAffaire, pStTiers, pStTypeAdr1, pStNumeroAdr1, pStTypeAdr2, pStNumeroAdr2: string; var pStTypeAdr3, pStNumeroAdr3 : string; pTob, pTobAffaire, pTobTiers: Tob): string;
var                                                           
  vTobAdresses: Tob;
  vTobTiers   : Tob;
  vTobContact : Tob;
  vSt         : string;
begin
  result := '';

  pTob.AddChampSup('CLIENTLIBELLE', false);
  pTob.AddChampSup('CLIENTPRENOM', false);
  pTob.AddChampSup('CLIENTADRESSE1', false);
  pTob.AddChampSup('CLIENTADRESSE2', false);
  pTob.AddChampSup('CLIENTADRESSE3', false);
  pTob.AddChampSup('CLIENTCODEPOSTAL', false);
  pTob.AddChampSup('CLIENTVILLE', false);
  PTob.AddChampSup('CLIENTCONTACT', False);
  pTob.AddChampSup('CLIENTTELEPHONE', false);

  if (pStNiveauRecherche = 'INT') and (pStNumeroAdr2 <> '') and (pStTypeAdr2 <> '') then
  begin
    vTobAdresses := TOB.Create('#ADRESSES', nil, -1);
    try
      vSt := 'SELECT ADR_LIBELLE, ADR_LIBELLE2, ADR_ADRESSE1, ADR_ADRESSE2, ';
      vSt := vSt + ' ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE, ADR_TELEPHONE FROM ADRESSES';
      vSt := vSt + ' WHERE ADR_TYPEADRESSE = "INT"';
      vSt := vSt + ' AND ADR_TYPEADRESSE = "' + pStTypeAdr2 + '"';
      vSt := vSt + ' AND ADR_NUMEROADRESSE = ' + pStNumeroAdr2;
      vTobAdresses.LoadDetailFromSQL(vSt);
      if vTobAdresses.detail.count > 0 then
         begin
         pTob.PutValue('CLIENTLIBELLE', vTobAdresses.detail[0].GetString('ADR_LIBELLE'));
         pTob.PutValue('CLIENTPRENOM', vTobAdresses.detail[0].GetString('ADR_LIBELLE2'));
         pTob.PutValue('CLIENTADRESSE1', vTobAdresses.detail[0].GetString('ADR_ADRESSE1'));
         pTob.PutValue('CLIENTADRESSE2', vTobAdresses.detail[0].GetString('ADR_ADRESSE2'));
         pTob.PutValue('CLIENTADRESSE3', vTobAdresses.detail[0].GetString('ADR_ADRESSE3'));
         pTob.PutValue('CLIENTCODEPOSTAL', vTobAdresses.detail[0].GetString('ADR_CODEPOSTAL'));
         pTob.PutValue('CLIENTVILLE', vTobAdresses.detail[0].GetString('ADR_VILLE'));
         pTob.PutValue('CLIENTTELEPHONE', vTobAdresses.detail[0].GetString('ADR_TELEPHONE'));
         result := 'INTERVENTION';
         end;
    finally
      vTobAdresses.Free;
    end;
  end;

  // chargement de l'adresse d'intervention définie au niveau de la tâche
  if (result = '') and (pStNumeroAdr1 <> '') and (pStTypeAdr1 <> '') then
  begin
    vTobAdresses := TOB.Create('#ADRESSES', nil, -1);
    try
      vSt := 'SELECT ADR_LIBELLE, ADR_LIBELLE2, ADR_ADRESSE1, ADR_ADRESSE2, ';
      vSt := vSt + ' ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE, ADR_TELEPHONE FROM ADRESSES';
      vSt := vSt + ' WHERE ADR_TYPEADRESSE = "INT"';
      vSt := vSt + ' AND ADR_TYPEADRESSE = "' + pStTypeAdr1 + '"';
      vSt := vSt + ' AND ADR_NUMEROADRESSE = ' + pStNumeroAdr1;
      vTobAdresses.LoadDetailFromSQL(vSt);

      if vTobAdresses.detail.count > 0 then
      begin
        pTob.PutValue('CLIENTLIBELLE', vTobAdresses.detail[0].GetString('ADR_LIBELLE'));
        pTob.PutValue('CLIENTPRENOM', vTobAdresses.detail[0].GetString('ADR_LIBELLE2'));
        pTob.PutValue('CLIENTADRESSE1', vTobAdresses.detail[0].GetString('ADR_ADRESSE1'));
        pTob.PutValue('CLIENTADRESSE2', vTobAdresses.detail[0].GetString('ADR_ADRESSE2'));
        pTob.PutValue('CLIENTADRESSE3', vTobAdresses.detail[0].GetString('ADR_ADRESSE3'));
        pTob.PutValue('CLIENTCODEPOSTAL', vTobAdresses.detail[0].GetString('ADR_CODEPOSTAL'));
        pTob.PutValue('CLIENTVILLE', vTobAdresses.detail[0].GetString('ADR_VILLE'));
        pTob.PutValue('CLIENTTELEPHONE', vTobAdresses.detail[0].GetString('ADR_TELEPHONE'));
        result := 'TACHE';
      end;
    finally
      vTobAdresses.Free;
    end;
  end

  // chargement de la premiere adresse d'intervention trouvée
  else if (result = '') then
  begin
    if pBoCreateTache then
    begin
      vTobAdresses := TOB.Create('#ADRESSES', nil, -1);
      try
        if assigned(pTobAffaire) and (pTobAffaire.GetString('ADR_LIBELLE') <> '') then
        begin
          pTob.PutValue('CLIENTLIBELLE', pTobAffaire.GetString('ADR_LIBELLE'));
          pTob.PutValue('CLIENTPRENOM', pTobAffaire.GetString('ADR_LIBELLE2'));
          pTob.PutValue('CLIENTADRESSE1', pTobAffaire.GetString('ADR_ADRESSE1'));
          pTob.PutValue('CLIENTADRESSE2', pTobAffaire.GetString('ADR_ADRESSE2'));
          pTob.PutValue('CLIENTADRESSE3', pTobAffaire.GetString('ADR_ADRESSE3'));
          pTob.PutValue('CLIENTCODEPOSTAL', pTobAffaire.GetString('ADR_CODEPOSTAL'));
          pTob.PutValue('CLIENTVILLE', pTobAffaire.GetString('ADR_VILLE'));
          pTob.PutValue('CLIENTTELEPHONE', pTobAffaire.GetString('ADR_TELEPHONE'));

          pStTypeAdr3 := pTobAffaire.GetString('ADR_TYPEADRESSE');
          pStNumeroAdr3 := pTobAffaire.GetString('ADR_NUMEROADRESSE');

          result := 'AFFAIRE';
        end
        else
        begin
          vSt := 'SELECT ADR_TYPEADRESSE, ADR_NUMEROADRESSE, ADR_LIBELLE, ADR_LIBELLE2, ADR_ADRESSE1, ADR_ADRESSE2, ';
          vSt := vSt + ' ADR_ADRESSE3, ADR_CODEPOSTAL, ADR_VILLE, ADR_TELEPHONE FROM ADRESSES';
          vSt := vSt + ' WHERE ADR_TYPEADRESSE = "INT"';
          vSt := vSt + ' AND ADR_REFCODE ="' + pStAffaire + '"';
          vTobAdresses.LoadDetailFromSQL(vSt);
          if vTobAdresses.detail.count > 0 then
          begin
            pTob.PutValue('CLIENTLIBELLE', vTobAdresses.detail[0].GetString('ADR_LIBELLE'));
            pTob.PutValue('CLIENTPRENOM', vTobAdresses.detail[0].GetString('ADR_LIBELLE2'));
            pTob.PutValue('CLIENTADRESSE1', vTobAdresses.detail[0].GetString('ADR_ADRESSE1'));
            pTob.PutValue('CLIENTADRESSE2', vTobAdresses.detail[0].GetString('ADR_ADRESSE2'));
            pTob.PutValue('CLIENTADRESSE3', vTobAdresses.detail[0].GetString('ADR_ADRESSE3'));
            pTob.PutValue('CLIENTCODEPOSTAL', vTobAdresses.detail[0].GetString('ADR_CODEPOSTAL'));
            pTob.PutValue('CLIENTVILLE', vTobAdresses.detail[0].GetString('ADR_VILLE'));
            pTob.PutValue('CLIENTTELEPHONE', vTobAdresses.detail[0].GetString('ADR_TELEPHONE'));

            pStTypeAdr3 := vTobAdresses.detail[0].GetString('ADR_TYPEADRESSE');
            pStNumeroAdr3 := vTobAdresses.detail[0].GetString('ADR_NUMEROADRESSE');

            if Assigned(pTobAffaire) and (pTobAffaire.GetString('ADR_LIBELLE') = '') then
              pTobAffaire.Dupliquer(vTobAdresses.detail[0], False, True);
            result := 'AFFAIRE';
          end;
        end;
      finally
        vTobAdresses.Free;
      end;
    end;
  end;

  if (result = '') then
  begin
    vTobTiers := TOB.Create('#TIERS', nil, -1);
    try
      if assigned(pTobTiers) and (pTobTiers.GetString('T_LIBELLE') <> '') then
      begin
        pTob.PutValue('CLIENTLIBELLE', pTobTiers.GetString('T_LIBELLE'));
        pTob.PutValue('CLIENTPRENOM', pTobTiers.GetString('T_PRENOM'));
        pTob.PutValue('CLIENTADRESSE1', pTobTiers.GetString('T_ADRESSE1'));
        pTob.PutValue('CLIENTADRESSE2', pTobTiers.GetString('T_ADRESSE2'));
        pTob.PutValue('CLIENTADRESSE3', pTobTiers.GetString('T_ADRESSE3'));
        pTob.PutValue('CLIENTCODEPOSTAL', pTobTiers.GetString('T_CODEPOSTAL'));
        pTob.PutValue('CLIENTVILLE', pTobTiers.GetString('T_VILLE'));
        pTob.PutValue('CLIENTTELEPHONE', pTobTiers.GetString('T_TELEPHONE'));
      end
      else
      begin
        vSt := 'SELECT T_LIBELLE, T_PRENOM, T_ADRESSE1, T_ADRESSE2, T_ADRESSE3, ';
        vSt := vSt + 'T_CODEPOSTAL, T_VILLE, T_TELEPHONE FROM TIERS ';
        vSt := vSt + 'WHERE T_TIERS = "' + pStTiers + '"';
        vTobTiers.LoadDetailFromSQL(vSt);
        if vTobTiers.detail.count > 0 then
        begin
          pTob.AddChampSupValeur('CLIENTLIBELLE',     vTobTiers.detail[0].GetString('T_LIBELLE'));
          pTob.AddChampSupValeur('CLIENTPRENOM',      vTobTiers.detail[0].GetString('T_PRENOM'));
          pTob.AddChampSupValeur('CLIENTADRESSE1',    vTobTiers.detail[0].GetString('T_ADRESSE1'));
          pTob.AddChampSupValeur('CLIENTADRESSE2',    vTobTiers.detail[0].GetString('T_ADRESSE2'));
          pTob.AddChampSupValeur('CLIENTADRESSE3',    vTobTiers.detail[0].GetString('T_ADRESSE3'));
          pTob.AddChampSupValeur('CLIENTCODEPOSTAL',  vTobTiers.detail[0].GetString('T_CODEPOSTAL'));
          pTob.AddChampSupValeur('CLIENTVILLE',       vTobTiers.detail[0].GetString('T_VILLE'));
          pTob.AddChampSupValeur('CLIENTTELEPHONE',   vTobTiers.detail[0].GetString('T_TELEPHONE'));
          if assigned(pTobTiers) and (pTobTiers.GetString('T_LIBELLE') = '') then
            pTobTiers.Dupliquer(vTobTiers.detail[0], False, True);
        end;
      end;
      result := 'CLIENT';
    finally
      vTobTiers.Free;
    end;
  end;

  pTob.PutValue('CLIENTCONTACT', '');

  //Recherche du contact principal du tiers
  vTobContact := TOB.Create('#CONTACT', nil, -1);
  vSt := 'SELECT C_NUMEROCONTACT FROM CONTACT WHERE C_TYPECONTACT = "T" ';
  vSt := vSt + 'AND C_TIERS = "' + pStTiers + '" AND C_PRINCIPAL="X"';
  vTobContact.LoadDetailFromSQL(vSt);
  if vTobContact.detail.count > 0 then
     pTob.AddChampSupValeur('CLIENTCONTACT', vTobContact.detail[0].GetString('C_NUMEROCONTACT'))
  else
     Begin
     vSt := 'SELECT C_NUMEROCONTACT FROM CONTACT WHERE C_TYPECONTACT = "T" ';
     vSt := vSt + 'AND C_TIERS = "' + pStTiers + '"';
     vTobContact.LoadDetailFromSQL(vSt);
     if vTobContact.detail.count > 0 then
        pTob.AddChampSupValeur('CLIENTCONTACT', vTobContact.detail[0].GetString('C_NUMEROCONTACT'));
     end;

  vTobContact.Free;

end;

//Génération de l'appel en fonction du contrat et de la tache
Procedure GenereAppel(fTOBDet : TOB);
var AnneeDeb  : String;
    AnneeFin  : String;
    Affaire   : String;
    StTiers   : String;
    StLibelle : String;
    Frequence : String;
    NoTache   : String;
    TypeAction: String;
    vSt       : String;
    ListInt   : TListDate;
    NbJrDecal : Integer;
    NbAppel   : Integer;
    i         : Integer;
    Duree     : Double;
    vAFReglesTache      : TAFRegles;
const
    TexteMessage: array[1..7] of string = (
    {1}'Affaire : %s.' + chr(10) + ' Tiers : %s. ' + Chr(10) + 'Aucune intervention n''a été générée pour la tâche %s. ',
    {2}'Si vous avez généré en utilisant ''Ajouter au planning existant'', vous planifiez à partir de la dernière date de génération précisée dans l''onglet ''Règles de la fiche Tâche''. ',
    {3}'Attention à la date de fin de génération paramètrée qui est actuellement au %s.',
    {4}'La génération de la tache %s c''est déroulée normalement. ' + Chr(10) + 'Nombre d''appel généré : %s ',
    {5}'La tâche à dupliquer a déjà été générée.' + chr(10) + 'Voulez-vous continuer en réinitialisant la tâche dupliquée ?',
    {6}'La date de fin de la tâche à générer est égale à 2099.' + chr(10) + 'Seule les 5 premières années seront générées !',
    {7}'La date de Début de la tâche à générer est égale à 1900.' + chr(10) + 'Cette date sera remplacé par la date de début de facturation !'
    );

begin

  if fTOBDet.GetString('ATA_FAITPARQUI')= '' then
  begin
    PGIBoxAF('Vous devez renseigner un contact', 'Génération Interventions Préventives');
    exit;
  end;

  //if fTobDet.GetString('ATA_TERMINE') = 'X' then
  //begin
  //  PGIBoxAF('Cette tâche a déjà été générée', 'Génération Interventions Préventives');
  //  exit;
  //end;

  TypeAction := fTobDet.GetString('ATA_BTETAT');
  if TypeAction = '' then
  Begin
    PGIBoxAF('Le type d''action est obligatoire pour la génération d''une tâche', 'Génération Interventions Préventives');
    exit;
  end;

  //Contrôle de la date de debut et de la date de fin
  AnneeDeb := FormatDateTime('yyyy',fTobDet.GetDateTime('ATA_DATEDEBPERIOD'));
  AnneeFin := FormatDateTime('yyyy',fTobDet.GetDateTime('ATA_DATEFINPERIOD'));

  if fTobDet.GetDateTime('ATA_DATEDEBPERIOD') = idate1900 then
  Begin
    PGIBoxAF(TexteMessage[7], 'Génération Interventions Préventives');
    AnneeDeb := FormatDateTime('yyyy',StrToDate(DateToStr(now)));
  end;

  if fTobDet.GetDateTime('ATA_DATEFINPERIOD') = idate2099 then
  Begin
    PGIBoxAF(TexteMessage[6], 'Génération Interventions Préventives');
    AnneeFin := FormatDateTime('yyyy', StrToDate(dateToStr(PlusDate(now,5,'A'))));
  end;

  if fTobDet.GetDateTime('ATA_DATEFINPERIOD') > PlusDate(now,5,'A') then
  Begin
    PGIBoxAF(TexteMessage[6], 'Génération Interventions Préventives');
    AnneeFin := FormatDateTime('yyyy', StrToDate(dateToStr(PlusDate(now,5,'A'))));
  end;

  //Définition de l'outil de calcul des règles d'échéance
  vAFReglesTache := TAFRegles.create;

  //génération en fonction des règles...
  NoTache   := FtobDet.GetString('ATA_NUMEROTACHE');
  Affaire   := FtobDet.GetString('ATA_AFFAIRE');
  StTiers   := FtobDet.GetString('ATA_TIERS');
  Frequence := fTobDet.GetString('ATA_PERIODICITE');
  NbJrDecal := fTobDet.GetInteger('ATA_NBJOURSDECAL');
  Duree     := fTobDet.GetValue('ATA_QTEINTERVENT');
  StLibelle := fTobDet.GetString('ATA_LIBELLETACHE1');
  vAFReglesTache.LoadDBReglesTaches(Affaire, NoTache, 1);
  ListInt := vAFReglesTache.GenereListeJours(Duree, NbJrDecal, frequence, -1);
  if ListInt = nil then exit;

  for i := ListInt.count - 1 downto 0 do
  begin
    if ListInt.Items[i] <= StrToDate(FtobDet.GetString('ATA_LASTDATEGENE')) then ListInt.delete(i);
  end;

  nbAppel := 0;

  if ListInt.count = 0 then
     begin
     //5 'Affaire : %s. Tiers : %s. Aucune intervention n''a été générée pour la tâche %s. Si vous avez généré en utilisant "Ajouter au planning existant", vous planifiez à partir de la dernière date de génération précisée dans l''onglet "Règles de la fiche Tâche".'
     vSt := format(TraduitGa(TexteMessage[1]), [affaire, StTiers, StLibelle]);
     vSt := vST + Chr(10) + TexteMessage[3] + DateToStr(FtobDet.GetValue('ATA_DATEANNUELLE'));
     if vAFReglesTache.DtDateFin < now then
        vSt := vSt + Chr(10) + format(TraduitGa(TexteMessage[2]), [dateToStr(vAFReglesTache.DtDateFin)]);
     PGIBoxAF(vst, 'Génération Interventions Préventives');
     end
  else
  Begin
    for i := 0 to ListInt.count - 1 do
    begin
     //controle si la date de génération est supérieure à la date du jour...
     if ListInt.Items[i] > now then
        Begin
        GenererListesAppel(ListInt, FTobDet, i, Duree);
        NbAppel := NbAppel+1;
        end;
    end;
    vst :=  'La génération de la tache ' + NoTache + ' s''est déroulée normalement. ' + Chr(10);
    vst := vst + 'Nombre d''appel(s) généré(s) : ' + inttostr(NbAppel);
    PGIBoxAF(vst,'Génération Interventions Préventives');
  End;

  vAFReglesTache.Free;

  ListInt.Free;

end;

Procedure GenererListesAppel(pListeJours : TListDate; FTOBDet : TOB; NumLigne : Integer; Duree: Double);
Var CodeAppel   : String;
    fCurAffaire : AffaireCourante; // tache en cours
    P0          : String;
    P1          : String;
    P2          : String;
    P3          : String;
    Avn         : String;
    //
    Req         : String;
    Ressource   : String;
    Contact     : String;
    BlocNote 	  : THRichEditOLE;
    NoAdresse   : Integer;
    DureeEnJour : Double;
    DateDeFin   : TDateTime;
    DateDeb     : TDateTime;
    TobAppel    : Tob;
    TobAdrInt   : Tob;
    TobAdrIntB   : Tob;
    TobRes      : Tob;
    QQ          : TQuery;
    Ok_Saisie   : Boolean;
    iPartErreur : Integer;
    IDuree      : integer;
begin

  // Si en création calcul du N° d'appel
  CodeAppel := CalculCodeAppel('W', P1, P2, P3, Avn, FTOBDet.GetString('ATA_TIERS'));

  if CodeAppel = '' then
  Begin
   PGIBox('Le code appel n''a pu être généré. Veuillez vérifier votre paramétrage.', 'Génération Interventions Préventives'); //Le code appel n'a pu être généré veuillez vérifier votre paramétrage
   exit;
  End;

  //Création de l'appel
  TobAppel := Tob.Create('AFFAIRE',Nil, -1);

  //découpage du code Appel
  CodeAppelDecoupe(CodeAppel, P0, P1, P2, P3, Avn);

  //Mise à jour de la table Affaire avec les zones de l'écran
  TobAppel.PutValue('AFF_AFFAIRE', CodeAppel);
  TobAppel.PutValue('AFF_AFFAIREREF', CodeAppel);

  TobAppel.PutValue('AFF_AFFAIRE0', 'W');
  TobAppel.PutValue('AFF_AFFAIRE1', P1);
  TobAppel.PutValue('AFF_AFFAIRE2', P2);
  TobAppel.PutValue('AFF_AFFAIRE3', P3);
  TobAppel.PutValue('AFF_AVENANT', Avn);
  TobAppel.PutValue('AFF_STATUTAFFAIRE', 'APP');
  TobAppel.PutValue('AFF_GENERAUTO', 'DIR');

  TobAppel.PutValue('AFF_DEVISE', FTOBDet.GetString('ATA_DEVISE'));

  TobAppel.PutValue('AFF_TIERS', FTOBDet.GetString('ATA_TIERS'));
  DateDeb := pListeJours.Items[NumLigne];
  TobAppel.PutValue('AFF_DATEDEBUT', DateDeb);

  if Duree < 1 then Duree := 1;
  IDuree := StrToInt(FloatToStr(Duree));
  if FTOBDet.GetString('ATA_UNITETEMPS') = 'H' Then
  begin
    DateDeFin := IncHour(DateDeb,IDuree);
  end else if FTOBDet.GetString('ATA_UNITETEMPS') = 'J' Then
  begin
    DateDeFin := IncDay(DateDeb,IDuree);
  end else if FTOBDet.GetString('ATA_UNITETEMPS') = 'MIN' Then
  begin
    DateDeFin := IncMinute(DateDeb,IDuree);
  end;

  TobAppel.PutValue('AFF_DATEFIN', DateDeFin);

  TobAppel.PutValue('AFF_DATELIMITE', iDate2099);
  TobAppel.PutValue('AFF_DATEGARANTIE', iDate2099);
  TobAppel.PutValue('AFF_DATEREPONSE', iDate2099);
  TobAppel.PutValue('AFF_ETATAFFAIRE', 'ECO');
  TobAppel.PutValue('AFF_AFFAIREINIT', FTOBDet.GetString('ATA_AFFAIRE'));
  TobAppel.PutValue('AFF_CHANTIER', '');
  TobAppel.PutValue('AFF_NUMEROTACHE', FTOBDet.GetString('ATA_NUMEROTACHE'));

  //Contrôle du contact
  Contact := FTOBDet.GetString('ATA_FAITPARQUI');
  if Contact <> '' then TobAppel.PutValue('AFF_NUMEROCONTACT', Contact);

  TobAppel.PutValue('AFF_CREERPAR', 'TAC');

  //Descriptif de l'appel obligatoire
  if isBlobVide(Application.mainForm,fTOBDet,'ATA_DESCRIPTIF' ) then
  Begin
    TobAppel.PutValue('AFF_DESCRIPTIF', FTOBDet.GetString('ATA_LIBELLETACHE1') + ' ' + FTOBDet.GetString('ATA_LIBELLETACHE2'));
    TobAppel.PutValue('AFF_LIBELLE', fTOBDet.GetString('ATA_LIBELLETACHE1'))
  end else
  Begin
    BlocNote := THRichEditOLE.create(Application.mainForm);
    BlocNote.parent := Application.mainForm;
    BlocNote.Visible := false;
    StringToRich(BlocNote, FTOBDet.GetString('ATA_DESCRIPTIF'));
    TobAppel.PutValue('AFF_DESCRIPTIF', FTOBDet.GetString('ATA_DESCRIPTIF'));
    //gestion du libellé en fonction du descriptif de la tache
    DefiniLibelle(fTOBDet,BlocNote);
    //--
    TobAppel.PutValue('AFF_LIBELLE', fTOBDet.GetString('AFF_LIBELLE'));
    BlocNote.free;
  end;


  TobAppel.PutValue('AFF_PRIOCONTRAT', GetParamSocSecur('SO_BTIMPTACHE','1'));

  //gestion de l'adresse
  fTOBDet.AddChampSupValeur('ADR_LIBELLE',   '');
  fTOBDet.AddChampSupValeur('ADR_LIBELLE2',    '');
  fTOBDet.AddChampSupValeur('ADR_ADRESSE1',  '');
  fTOBDet.AddChampSupValeur('ADR_ADRESSE2',  '');
  fTOBDet.AddChampSupValeur('ADR_ADRESSE3',  '');
  fTOBDet.AddChampSupValeur('ADR_CODEPOSTAL','');
  fTOBDet.AddChampSupValeur('ADR_VILLE',     '');
  fTOBDet.AddChampSupValeur('CLIENTTELEPHONE', '');
  fTOBDet.AddChampSupValeur('CLIENTCONTACT',   '');
  
  //Vérification si adresse existe déjà pour l'appel
  LoadAdresseTache(True, fTobDet);
  NoAdresse := GetSetNumAdresse; //TOBDet.GetInteger('ATA_NUMEROADRESSE');

  TobAdrInt := Tob.create('ADRESSES',nil, -1);
  TobAdrIntB := Tob.create('BADRESSES',nil, -1);

  TobAdrInt.PutValue('ADR_NUMEROADRESSE',  NoAdresse);
  TobAdrInt.PutValue('ADR_NATUREAUXI', '');
  TobAdrInt.PutValue('ADR_REFCODE', CodeAppel);
  TobAdrInt.PutValue('ADR_TYPEADRESSE', 'INT');
  TobAdrInt.PutValue('ADR_JURIDIQUE', fTOBDet.GetString('T_JURIDIQUE'));
  TobAdrInt.PutValue('ADR_LIBELLE',   fTOBDet.GetString('ADR_LIBELLE'));
  TobAdrInt.PutValue('ADR_LIBELLE2',  fTOBDet.GetString('ADR_LIBELLE2'));
  TobAdrInt.PutValue('ADR_ADRESSE1',  fTOBDet.GetString('ADR_ADRESSE1'));
  TobAdrInt.PutValue('ADR_ADRESSE2',  fTOBDet.GetString('ADR_ADRESSE2'));
  TobAdrInt.PutValue('ADR_ADRESSE3',  fTOBDet.GetString('ADR_ADRESSE3'));
  TobAdrInt.PutValue('ADR_CODEPOSTAL',fTOBDet.GetString('ADR_CODEPOSTAL'));
  TobAdrInt.PutValue('ADR_VILLE',     fTOBDet.GetString('ADR_VILLE'));
  TobAdrInt.PutValue('ADR_PAYS', '');
  TobAdrInt.PutValue('ADR_EMAIL', '');
  TobAdrInt.PutValue('ADR_REGION', '');
  TobAdrInt.PutValue('ADR_LIVR', 'X');
  TobAdrInt.PutValue('ADR_FACT', 'X');
  TobAdrInt.PutValue('ADR_REGL', 'X');
  TobAdrInt.PutValue('ADR_TELEPHONE', fTOBDet.GetString('CLIENTTELEPHONE'));

  TobAdrInt.InsertOrUpdateDB(True);

  TobAdrIntB.PutValue('BA0_TYPEADRESSE', 'INT');
  TobAdrIntB.PutValue('BA0_NUMEROADRESSE',  NoAdresse);
  TobAdrIntB.PutValue('BA0_INT',  'X');
  TobAdrIntB.SetAllModifie(true); 
  TobAdrIntB.InsertOrUpdateDB(True);


  if Contact <> '' then
  Begin
    TobAppel.PutValue('AFF_NUMEROCONTACT', ftobDet.GetString('ATA_FAITPARQUI'));
    TobAdrInt.PutValue('ADR_CONTACT', ftobDet.GetString('ATA_FAITPARQUI'));
  end;

  //recherche du responsable dans la tob des ressources associée à la grille
  Ressource := RechRessPrinc(FTOBDet.GetString('ATA_AFFAIRE'), fTobDet.GetValue('ATA_NUMEROTACHE'));
  if Ressource <> '' then
  Begin
    TobAppel.PutValue('AFF_RESPONSABLE', Ressource);
    TobAppel.PutValue('AFF_ETATAFFAIRE', 'AFF');
  	//GenereEvenement(TobAppel, TobAdrInt.GetValue('ADR_NUMEROADRESSE'));
  end
  else
    TobAppel.PutValue('AFF_DATESIGNE', TobAppel.GetValue('AFF_DATEDEBUT'));

  TobAppel.InsertOrUpdateDB(true);

  TobAdrInt.free;
  TobAdrIntB.free;
  TobAppel.free;

end;

Function RechRessPrinc(Affaire, NoTache : String) : String;
Var TobTachesRes : Tob;
    vQr          : TQuery;
    vStSQL       : String;
Begin

    Result  := '';

    TobTachesRes := Tob.Create('TACHERESSOURCE', nil, -1);

    vStSQL := 'SELECT ATR_RESSOURCE FROM TACHERESSOURCE WHERE ATR_AFFAIRE="' + Affaire + '"';
    vStSQL := vStSQL + ' AND ATR_NUMEROTACHE=' + NoTache;

    vQr := nil;

    try
      vQR := OpenSQL(vStSQL, True);
      if not (vQR.EOF) then
         Begin
         TobTachesRes.LoadDetailDB('TACHERESSOURCE', '', '', vQR, True);
         Result := TobTachesRes.detail[0].GetValue('ATR_RESSOURCE');
         end;
    finally
      Ferme(vQR);
    end;

    TobTachesRes.free;

end;


end.

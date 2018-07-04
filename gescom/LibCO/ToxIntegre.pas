{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Intégration des fichiers TOX
Suite ........ : Module de contrôle et d'intégration des données
Mots clefs ... : TOX;CONTROLE;INTEGRATION
*****************************************************************}
unit ToxIntegre;
interface

uses {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hctrls, SysUtils, uTob, StdCtrls, Dialogs, FactComm, FactUtil, FactCpta,
  SaisUtil, HEnt1, FactNomen, ToxMessage, UtilPGI, Ent1, FactCalc, ParamSoc, EntGC,
  TarifUtil, stockutil,  UtoxClasses, FactTob, FactArticle, FactLotSerie,uEntCommun
{$IFDEF BTP}
	,FactOuvrage,FactTvaMilliem
{$ENDIF}
,UtilTOBPiece
  ;

//////////////////////////////////////////////////////////////////////////////
// Liste des fonctions permettant de contrôler ou d'intégrer un enregistrement
///////////////////////////////////////////////////////////////////////////////
function ImportTOBAcomptes(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBAdresses(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBAffaire(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArrondi(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArticle(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArticleCompl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArticleLie(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArticlePiece(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBArticleTiers(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBMemorisation (TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCatalogu(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBChancell(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBChoixcod(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBChoixExt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBClavierEcran(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCodepost(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCommentaire(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCommercial(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCompteurEtab(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBContact(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBCtrlCaisMt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBDepots(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBDevise(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBDimension(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBDimasque(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBDispo(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBEtabliss(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBExercice(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBFideliteEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBFideliteLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBFonction(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBJoursCaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBJoursEtab(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBJoursEtabEvt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBLiensOLE(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBLigne(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBLigneLot(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBLigneNomen(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
{$IFDEF BTP}
function ImportTOBPardoc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBLigneOuvrages(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBNaturePrest(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBMillieme(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTobPieceRg(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
{$ENDIF}
function ImportTOBListeInvEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBListeInvLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBMea(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBModeles(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBModedata(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBModeRegl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBModePaie(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBNomenEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBNomenLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBOperCaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParamSoc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParamTaxe(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParcaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParFidelite(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParRegleFid(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParSeuilFid(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBParPiecBil(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPays(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPiedBase(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPiedEche(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPiedPort(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPiece(TPST: TOB; Integre, IntegrationRapide, MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPieceAdresse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBPort(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBProfilArt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBRegion(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBSociete(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB; CodeSite: string): integer;
function ImportTOBStoxquerys(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTarif(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTarifMode(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
//function ImportTOByTarifs(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
//function ImportTOByTarifsFourchette(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTarifPer(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTarifTypMode(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTiers(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTiersCompl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTradDico(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTradTablette(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTraduc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTxcpttva(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTypeMasque(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBTypeRemise(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBUserGrp(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function ImportTOBUtilisat(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;

function ToxControleInfo(NomTable: string; var ToxInfo: TOB; Integre, IntegrationRapide, MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB; CodeSite: string):
  integer;
function ImportDocument(MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
function AnnuleStockPiece(TobPiece: TOB): integer;

////////////////////////////////////////////////////////////////////////////////
// TOB utilisées pour l'import des documents
////////////////////////////////////////////////////////////////////////////////
var
  Tob_Entete: TOB; // TOB entête de pièce
  Tob_Ligne: TOB; // TOB Lignes
  Tob_LigneNom: TOB; // TOB Lignes Nomenclatures
{$IFDEF BTP}
  Tob_LigneOuv: TOB; // TOB Lignes Ouvrages
  Tob_Millieme: TOB; // TOB Lignes Milliemes (repartition de tva)
  Tob_PieceRg: TOB;   // TOB pied bases RG
  Tob_PiedBaseRg: TOB;   // TOB pied bases RG
  Tob_Ouvrages: TOB; // TOB des ouvrages
{$ENDIF}
  Tob_LigneLot: TOB; // TOB Lignes des Lots
  Tob_PiedEche: TOB; // TOB PIEDECHE
  Tob_PiedBase: TOB; // TOB PIEDBASE
  Tob_PiedPort: TOB; // TOB PIEDPORT
  Tob_Acomptes: TOB; // TOB ACOMPTES
  Tob_Tiers: TOB; // TOB du tiers
  Tob_Article: TOB; // TOB des articles
  Tob_Catalogue: TOB; // TOB des catalogues
  Tob_Nomenclature: TOB; // TOB des nomenclatures

  // Import des pièces
  DernierTiers: string; // Derniers Tiers testé
  DerniereNature: string; // Dernière nature
  DernierEtab: string; // Dernier Etab
  DernierDepot: string;
  DernierModeRegl: string;
  DernierRegimeTaxe: string;
  DerniereDevise: string;
  DernierTarifArticle: string;
  DerniereCatTaxe: string;
  DerniereFamTaxe: string;
  DernierModePaie: string;
  DernierTypeRemise: string;

  // Import des articles
  DerniereFamArt1: string;
  DerniereFamArt2: string;
  DerniereFamArt3: string;
  DerniereFamTaxe1: string;
  DerniereFamTaxe2: string;
  DerniereCollection: string;
  DerniereTypeArt: string;

  // Gestion des suppressions de tablettes / claviers ....
  DernierTypeChoixcod: string;
  DernierTypeChoixext: string;
  DernierCaisseSupp: string;
  ModePaieSupp: boolean;
  ModeReglSupp: boolean;
  ParPiecBilSupp: boolean;

implementation

uses
{  FactPieceContainer, yTarifs_TOM, }uTom, wTom;


//
// Fonction permettant de remettre en forme une pièce dans la devise du site récepteur
//
procedure ToxCalculFacture(sDevise: string; sbascule: boolean);
var
	DEV: RDEVISE;
	CleDoc: R_CleDoc;
{  PieceContainer: TPieceContainer;}
begin
  //
  // Le site émetteur et le site récepteur sont ils gérés dans des devises différentes ?
  //
  if (sbascule <> VH^.TenueEuro) or (sDevise <> V_PGI.DeviseFongible) then
  begin
    // Cas 1 : le site récepteur n'a pas basculé en EURO
    if not VH^.TenueEuro then
    begin
      if not sBascule then // Le site émetteur n'a pas basculé en EURO
      begin
        if Tob_Entete.GetValue('GP_SAISIECONTRE') = 'X' then
        begin
          Tob_Entete.PutValue('GP_DEVISE', V_PGI.DevisePivot);
        end;
      end else
      begin // Le site émetteur a basculé en EURO
        if Tob_Entete.GetValue('GP_SAISIECONTRE') = 'X' then
        begin
          Tob_Entete.PutValue('GP_SAISIECONTRE', '-');
          Tob_Entete.PutValue('GP_DEVISE', sDevise);
        end else
        begin
          if Tob_Entete.GetValue('GP_DEVISE') = 'EUR' then
          begin
            Tob_Entete.PutValue('GP_SAISIECONTRE', 'X');
            Tob_Entete.PutValue('GP_DEVISE', V_PGI.DevisePivot);
          end;
        end;
      end;
    end else
      // Cas 2 : le site récepteur a basculé en EURO
    begin
      if not sBascule then // Le site émetteur n'a pas basculé en EURO
      begin
        if Tob_Entete.GetValue('GP_SAISIECONTRE') = 'X' then
        begin
          Tob_Entete.PutValue('GP_SAISIECONTRE', '-');
          Tob_Entete.PutValue('GP_DEVISE', V_PGI.DevisePivot);
        end else
        begin
          if Tob_Entete.GetValue('GP_DEVISE') = V_PGI.DeviseFongible then
          begin
            Tob_Entete.PutValue('GP_SAISIECONTRE', 'X');
            Tob_Entete.PutValue('GP_DEVISE', V_PGI.DevisePivot);
          end;
        end;
      end else
      begin
        if Tob_Entete.GetValue('GP_SAISIECONTRE') = 'X' then
        begin
          Tob_Entete.PutValue('GP_SAISIECONTRE', '-');
          Tob_Entete.PutValue('GP_DEVISE', sDevise);
        end else
        begin
          if Tob_Entete.GetValue('GP_DEVISE') = V_PGI.DeviseFongible then
          begin
            Tob_Entete.PutValue('GP_SAISIECONTRE', 'X');
            Tob_Entete.PutValue('GP_DEVISE', V_PGI.DevisePivot);
          end;
        end;
      end;
    end;
    // Recalcul de la pièce
{$IFNDEF BTP}
    PutValueDetail(Tob_Entete, 'GP_RECALCULER', 'X');
{$ENDIF}
{GC_NEWFACT_TP-Début}
  DEV.Code   := Tob_Entete.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  CleDoc     := TOB2CleDoc(Tob_Entete);
  DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, CleDoc.DatePiece);
{$IFDEF BTP}
	AppliqueMillieme (Tob_Millieme,Tob_Entete,Tob_PiedBase,Tob_Ouvrages);
//	AjouteLignesVirtuellesOuv (TOB_entete,Tob_LigneOuv,Tob_Article,nil,TOB_Tiers,TOB_Catalogue,nil,nil,DEV);
//	AjouteLignesVirtuellesOuv (TOB_entete,Tob_ouvrages,Tob_Article,nil,TOB_Tiers,TOB_Catalogue,nil,nil,DEV);
	PutValueDetail(TOB_entete,'GP_RECALCULER','X');
{$ENDIF}
  CalculFacture(Tob_Entete,nil,nil,Tob_PiedBase,nil,Tob_Tiers,Tob_Article,Tob_PiedPort,TOB_piecerg,TOB_PiedBaseRg,DEV);
{$IFDEF BTP}
//	DetruitLignesVirtuellesOuv (TOB_Entete,DEV);
{$ENDIF}
{GC_NEWFACT_TP-Fin}
  end;
end;

//
// Fonction permettant de remettre en forme une pièce dans la devise du site récepteur
//
procedure ToxCalculEcheance(var Tob_Echeance: TOB; sDevise: string; sbascule: boolean; TobDev: TOB);
var
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin
  //
  // Le site émetteur et le site récepteur sont ils gérés dans des devises différentes ?
  //
  if (sbascule <> VH^.TenueEuro) or (sDevise <> V_PGI.DeviseFongible) then
  begin
    Tob_Echeance.PutValue('GPE_TAUXDEV', Tob_Entete.GetValue('GP_TAUXDEV'));
    Tob_Echeance.PutValue('GPE_COTATION', Tob_Entete.GetValue('GP_COTATION'));
    //
    // Cas particulier des tickets : On repart de GPE_DEVISEESP et GPE_MONTANTENCAIS
    //
    if Tob_Echeance.GetValue('GPE_NATUREPIECEG') = 'FFO' then
    begin
      // Montant Echeance
      DevOrg := Tob_Echeance.GetValue('GPE_DEVISEESP');
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;
      MontantDev := 0; // initialisation bidon
      MontantOld := 0; // initialisation bidon
      Montant := Tob_Echeance.GetValue('GPE_MONTANTENCAIS');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Echeance.PutValue('GPE_MONTANTECHE', MontantDev);
      Tob_Echeance.PutValue('GPE_MONTANTDEV', MontantDev);
      // Ré-initialisation des devises
      Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
      Tob_Echeance.PutValue('GPE_SAISIECONTRE', '-');
    end else
    begin // Autres documents
      // Cas 1 : le site récepteur n'a pas basculé en EURO
      if not VH^.TenueEuro then
      begin
        if not sBascule then // Le site émetteur n'a pas basculé en EURO
        begin
          if Tob_Echeance.GetValue('GPE_SAISIECONTRE') = 'X' then
          begin
            Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
          end;
        end else
        begin // Le site émetteur a basculé en EURO
          if Tob_Echeance.GetValue('GPE_SAISIECONTRE') = 'X' then
          begin
            Tob_Echeance.PutValue('GPE_SAISIECONTRE', '-');
            Tob_Echeance.PutValue('GPE_DEVISE', sDevise);
          end else
          begin
            if Tob_Echeance.GetValue('GPE_DEVISE') = 'EUR' then
            begin
              Tob_Echeance.PutValue('GPE_SAISIECONTRE', 'X');
              Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
            end;
          end;
        end;
      end else
        // Cas 2 : le site récepteur a basculé en EURO
      begin
        if not sBascule then // Le site émetteur n'a pas basculé en EURO
        begin
          if Tob_Echeance.GetValue('GPE_SAISIECONTRE') = 'X' then
          begin
            Tob_Echeance.PutValue('GPE_SAISIECONTRE', '-');
            Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
          end else
          begin
            if Tob_Echeance.GetValue('GPE_DEVISE') = V_PGI.DeviseFongible then
            begin
              Tob_Echeance.PutValue('GPE_SAISIECONTRE', 'X');
              Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
            end;
          end;
        end else
        begin
          if Tob_Echeance.GetValue('GPE_SAISIECONTRE') = 'X' then
          begin
            Tob_Echeance.PutValue('GPE_SAISIECONTRE', '-');
            Tob_Echeance.PutValue('GPE_DEVISE', sDevise);
          end else
          begin
            if Tob_Echeance.GetValue('GPE_DEVISE') = V_PGI.DeviseFongible then
            begin
              Tob_Echeance.PutValue('GPE_SAISIECONTRE', 'X');
              Tob_Echeance.PutValue('GPE_DEVISE', V_PGI.DevisePivot);
            end;
          end;
        end;
      end;
      //
      // Conversion
      //
      // Devise d'origine ?
      //
      if Tob_Echeance.GetValue('GPE_SAISIECONTRE') = '-' then
      begin
        DevOrg := Tob_Echeance.GetValue('GPE_DEVISE')
      end else
      begin
        if VH^.TenueEuro then DevOrg := V_PGI.DeviseFongible
        else DevOrg := 'EUR';
      end;
      //
      // Devise destinataire
      //
      DevDest := V_PGI.DevisePivot;
      //
      // Calcul
      //
      Dateconv := NowH;
      Montant := 0; // initialisation bidon
      MontantOld := 0; // initialisation bidon
      MontantDev := Tob_Echeance.GetValue('GPE_MONTANTDEV');
      ToxConvertir(MontantDev, Montant, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Echeance.PutValue('GPE_MONTANTECHE', Montant);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ACOMPTES
Mots clefs ... : TOX;
*****************************************************************}
function ImportTOBAcomptes(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Acpt: TOB;
  SQL: string;
  Q: TQUERY;
  StRech: string;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ///////////////////////////////////////////////////////////////////////////
      // Mode de paiement
      ////////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GAC_MODEPAIE');
      if (StRech <> '') then
      begin
        SQL := 'Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrModePaie;
        Ferme(Q);
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
    //  PHASE D'INTEGRATION
    //  1 - Création de la TOB ACOMPTES
    //  2 - Conversion éventuelle en devise
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Acpt := TOB.Create('ACOMPTES', nil, -1);
    Tob_Acpt.Dupliquer(TPST, False, True, True);

    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Montants
      Montant := Tob_Acpt.GetValue('GAC_MONTANTDEV');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Acpt.PutValue('GAC_MONTANT', MontantDev);
    end;

    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  2 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
  //////////////////////////////////////////////////////////////////////////////////////
    Tob_Acpt.ChangeParent(Tob_Acomptes, -1);
    Tob_Acpt.SetAllModifie(True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Santucci
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB AFFAIRE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBAffaire(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Affaires: TOB;
  Affaire: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Affaires := TOB.Create('AFFAIRE', nil, -1);
    Tob_Affaires.Dupliquer(TPST, False, True, True);
    Tob_Affaires.SetAllModifie(True);
    Affaire := TPST.GetValue('AFF_AFFAIRE');

    SQL := 'Select AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="' + Affaire + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Affaires.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Affaires.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Affaires.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ADRESSES
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBAdresses(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Adresses: TOB;
  NumAdresse: integer;
  RefAdresse: string;
  TypeAdresse: string;
  NaturePiece: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      // Adresse d'une pièce
      // Contrôle 1 : Nature de pièce existe ?
      /////////////////////////////////////////////////////////////////
      if (TPST.GetValue('ADR_TYPEADRESSE') = 'PIE') then
      begin
        RefAdresse := TPST.GetValue('ADR_REFCODE');
        NaturePiece := ReadTokenSt(RefAdresse);

        SQL := 'Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="' + NaturePiece + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrNaturePiece;
        Ferme(Q);
      end
        ////////////////////////////////////////////////////////
        // Adresse d'un tiers
        ////////////////////////////////////////////////////////
      else if (TPST.GetValue('ADR_TYPEADRESSE') = 'TIE') then
      begin
        // ???? Le client n'est pas forcément créé
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Adresses := TOB.Create('ADRESSES', nil, -1);
    Tob_Adresses.Dupliquer(TPST, False, True, True);
    Tob_Adresses.SetAllModifie(True);

    NumAdresse := Tob_Adresses.GetValue('ADR_NUMEROADRESSE');
    TypeAdresse := Tob_Adresses.GetValue('ADR_TYPEADRESSE');
    SQL := 'Select ADR_REFCODE From ADRESSES WHERE ADR_NUMEROADRESSE="' + IntToStr(NumAdresse) + '" AND ADR_TYPEADRESSE="' + TypeAdresse + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Adresses.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Adresses.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Adresses.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARRONDI
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArrondi(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Arrondi: TOB;
  CodeArrondi: string;
  Rang: integer;
  SQL: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Arrondi := TOB.Create('ARRONDI', nil, -1);
    Tob_Arrondi.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Valeur Seuil d'application arrondi
      Montant := Tob_Arrondi.GetValue('GAR_VALEURSEUIL');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Arrondi.PutValue('GAR_VALEURSEUIL', MontantDev);

      // Constante prix psycologique
      Montant := Tob_Arrondi.GetValue('GAR_CONSTANTE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Arrondi.PutValue('GAR_CONSTANTE', MontantDev);
    end;
    //
    // On met à jour tous les champs
    //
    Tob_Arrondi.SetAllModifie(True);

    CodeArrondi := Tob_Arrondi.GetValue('GAR_CODEARRONDI');
    Rang := Tob_Arrondi.GetValue('GAR_RANG');
    SQL := 'Select GAR_LIBELLE From ARRONDI WHERE GAR_CODEARRONDI="' + CodeArrondi + '" AND GAR_RANG="' + IntToStr(Rang) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Arrondi.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Arrondi.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Arrondi.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARTICLE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArticle(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Article: TOB;
  CodeArticle: string;
  StRech: string;
  Zone: string;
  YxType: string;
  SQL: string;
  DevOrg, DevDest: string;
  Q: TQUERY;
  MontantDev: double;
  MontantOld: double;
  Montant: double;
  Dateconv: TDateTime;
  i: integer;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      // Contrôle 1	:	Famille 1
      // Contrôle 2 : Famille 2
      // Contrôle 3 : Famille 3
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_FAMILLENIV1');
        if (StRech <> '') and (StRech <> DerniereFamArt1) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="FN1" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamilleNiv1
          else DerniereFamArt1 := StRech;
          Ferme(Q);
        end;
      end;
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_FAMILLENIV2');
        if (StRech <> '') and (StRech <> DerniereFamArt2) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="FN2" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamilleNiv2
          else DerniereFamArt2 := StRech;
          Ferme(Q);
        end;
      end;
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_FAMILLENIV3');
        if (StRech <> '') and (StRech <> DerniereFamArt3) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="FN3" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamilleNiv3
          else DerniereFamArt3 := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 4	:	Famille taxe 1
      // Contrôle 5	:	Famille taxe 2
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_FAMILLETAXE1');
        if (StRech <> '') and (StRech <> DerniereFamTaxe1) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamilleTaxe1
          else DerniereFamTaxe1 := StRech;
          Ferme(Q);
        end;
      end;
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_FAMILLETAXE2');
        if (StRech <> '') and (StRech <> DerniereFamTaxe2) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TX2" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamilleTaxe2
          else DerniereFamTaxe2 := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 6	:	Collection
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_COLLECTION');
        if (StRech <> '') and (StRech <> DerniereCollection) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GCO" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrCollection
          else DerniereCollection := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 7	:	Tarif Article
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_TARIFARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TAR" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTarifArticle;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 8	:	Statut
      /////////////////////////////////////////////////////////////////
      // if (result=IntegOK) then
      //begin
      //  StRech := TPST.GetValue ('GA_STATUTART');
      //  if (StRech <> '') then
      //  begin
      //  	SQL:='Select CO_LIBELLE From COMMUN WHERE CO_TYPE="GDS" AND CO_CODE="'+StRech+'"';
      //  	Q:=OpenSQL(SQL,True) ;
      //		if Q.EOF then result := ErrStatutArticle;
      //Ferme(Q);
      //end;
      //end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 9	:	Type article
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_TYPEARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select CO_LIBELLE From COMMUN WHERE CO_TYPE="TYA" AND CO_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTypeArticle;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôle 10	:	Dimensions
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA_DIMMASQUE');
        if (StRech <> '') then
        begin
          SQL := 'Select * From DIMMASQUE WHERE GDM_MASQUE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDimMasque
          else
          begin
            StRech := TPST.GetValue('GA_GRILLEDIM1');
            if (StRech <> '') then
            begin
              if (StRech <> Q.FindField('GDM_TYPE1').AsString) then result := ErrGrilleDim1;
            end;
            StRech := TPST.GetValue('GA_GRILLEDIM2');
            if ((StRech <> '') and (result = IntegOK)) then
            begin
              if (StRech <> Q.FindField('GDM_TYPE2').AsString) then result := ErrGrilleDim2;
            end;
            StRech := TPST.GetValue('GA_GRILLEDIM3');
            if ((StRech <> '') and (result = IntegOK)) then
            begin
              if (StRech <> Q.FindField('GDM_TYPE3').AsString) then result := ErrGrilleDim3;
            end;
            StRech := TPST.GetValue('GA_GRILLEDIM4');
            if ((StRech <> '') and (result = IntegOK)) then
            begin
              if (StRech <> Q.FindField('GDM_TYPE4').AsString) then result := ErrGrilleDim4;
            end;
            StRech := TPST.GetValue('GA_GRILLEDIM5');
            if ((StRech <> '') and (result = IntegOK)) then
            begin
              if (StRech <> Q.FindField('GDM_TYPE5').AsString) then result := ErrGrilleDim5;
            end;
          end;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Contrôles 11 à 20	:	Zone table libre 1 à 10
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        for i := 1 to 10 do
        begin
          if (i < 10) then
          begin
            Zone := 'GA_LIBREART' + IntToStr(i);
            YxType := 'LA' + IntToStr(i);
          end
          else
          begin
            Zone := 'GA_LIBREARTA';
            YxType := 'LAA';
          end;

          StRech := TPST.GetValue(Zone);
          if (StRech <> '') then
          begin
            SQL := 'Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="' + YxType + '" AND YX_CODE="' + StRech + '"';
            Q := OpenSQL(SQL, True);
            if Q.EOF then
            begin
              if (i = 1) then result := ErrStatArt1
              else if (i = 2) then result := ErrStatArt2
              else if (i = 3) then result := ErrStatArt3
              else if (i = 4) then result := ErrStatArt4
              else if (i = 5) then result := ErrStatArt5
              else if (i = 6) then result := ErrStatArt6
              else if (i = 7) then result := ErrStatArt7
              else if (i = 8) then result := ErrStatArt8
              else if (i = 9) then result := ErrStatArt9
              else if (i = 10) then result := ErrStatArt10;
            end;
            Ferme(Q);
          end;
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Article := TOB.Create('ARTICLE', nil, -1);
    Tob_Article.Dupliquer(TPST, False, True, True);
    ////////////////////////////////////////////////////////////////
    // Conversion dans la devise du dossier
    ////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Dernier prix d'achat
      Montant := Tob_Article.GetValue('GA_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_DPA', MontantDev);
      // Dernier prix de revient
      Montant := Tob_Article.GetValue('GA_DPR');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_DPR', MontantDev);
      // Prix de vente HT
      Montant := Tob_Article.GetValue('GA_PVHT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PVHT', MontantDev);
      // Prix de vente TTC
      Montant := Tob_Article.GetValue('GA_PVTTC');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PVTTC', MontantDev);
      // Prix de base achat HT
      Montant := Tob_Article.GetValue('GA_PAHT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PAHT', MontantDev);
      // Prix de base revient HT
      Montant := Tob_Article.GetValue('GA_PRHT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PRHT', MontantDev);
      // Prix d'achat moyen pondéré
      Montant := Tob_Article.GetValue('GA_PMAP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PMAP', MontantDev);
      // Prix de revient moyen pondéré
      Montant := Tob_Article.GetValue('GA_PMRP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PMRP', MontantDev);
      // Minium de marge
      Montant := Tob_Article.GetValue('GA_MARGEMINI');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_MARGEMINI', MontantDev);
      // Prix simulation achat
      Montant := Tob_Article.GetValue('GA_PRIXSIMULACH');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Article.PutValue('GA_PRIXSIMULACH', MontantDev);
    end;

    Tob_Article.PutValue('GA_DATEINTEGR', NowH);
    Tob_Article.SetAllModifie(True);
    Tob_Article.SetDateModif(Tob_Article.GetValue('GA_DATEMODIF'));

    CodeArticle := Tob_Article.GetValue('GA_ARTICLE');
    SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + CodeArticle + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Article.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Article.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Article.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARTICLECOMPL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArticleCompl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ArticleCompl: TOB;
  CodeArticle: string;
  StRech: string;
  Cpt: integer;
  SQL: string;
  TypeChoix: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      //  Famille article 4 à 8
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        for cpt := 4 to 8 do
        begin
          StRech := TPST.GetValue('GA2_FAMILLENIV' + IntToStr(cpt));
          if (StRech <> '') then
          begin
            TypeChoix := 'FN' + IntToStr(cpt);
            SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="' + TypeChoix + '" AND CC_CODE="' + StRech + '"';
            Q := OpenSQL(SQL, True);
            if Q.EOF then
            begin
              if cpt = 4 then result := ErrFamilleNiv4
              else if cpt = 5 then result := ErrFamilleNiv5
              else if cpt = 6 then result := ErrFamilleNiv6
              else if cpt = 7 then result := ErrFamilleNiv7
              else if cpt = 8 then result := ErrFamilleNiv8;
            end;
            Ferme(Q);
          end;
        end;
      end;
      /////////////////////////////////////////////////////////////////
      //  Statistique 1 et 2
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA2_STATART1');
        if (StRech <> '') then
        begin
          TypeChoix := 'FNA';
          SQL := 'Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="' + TypeChoix + '" AND YX_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrStat9Art1;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      //  Statistique 1 et 2
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GA2_STATART2');
        if (StRech <> '') then
        begin
          TypeChoix := 'FNB';
          SQL := 'Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="' + TypeChoix + '" AND YX_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrStat9Art2;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ArticleCompl := TOB.Create('ARTICLECOMPL', nil, -1);
    Tob_ArticleCompl.Dupliquer(TPST, False, True, True);
    Tob_ArticleCompl.SetAllModifie(True);

    CodeArticle := Tob_ArticleCompl.GetValue('GA2_ARTICLE');
    SQL := 'Select GA2_CODEARTICLE From ARTICLECOMPL WHERE GA2_ARTICLE="' + CodeArticle + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ArticleCompl.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ArticleCompl.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ArticleCompl.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARTICLELIE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArticleLie(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ArticleLie: TOB;
  CodeArticle: string;
  StRech: string;
  TypeLien: string;
  Rang: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      //  Code Article de référence
      /////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GAL_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_TYPEARTICLE From ARTICLE WHERE GA_CODEARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticleRef;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////
      //  Code Article de rattachement
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GAL_ARTICLELIE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_TYPEARTICLE From ARTICLE WHERE GA_CODEARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticleRat;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ArticleLie := TOB.Create('ARTICLELIE', nil, -1);
    Tob_ArticleLie.Dupliquer(TPST, False, True, True);
    Tob_ArticleLie.SetAllModifie(True);

    TypeLien := Tob_ArticleLie.GetValue('GAL_TYPELIENART');
    CodeArticle := Tob_ArticleLie.GetValue('GAL_ARTICLE');
    Rang := Tob_ArticleLie.GetValue('GAL_RANG');
    SQL := 'Select GAL_LIBELLE From ARTICLELIE WHERE GAL_TYPELIENART="' + TypeLien + '" AND GAL_ARTICLE="' + CodeArticle + '" AND GAL_RANG="' + IntToStr(Rang) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ArticleLie.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ArticleLie.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ArticleLie.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARTICLEPIECE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArticlePiece(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ArticlePiece: TOB;
  StRech: string;
  CodeArticle: string;
  Nature: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      //  Code Article
      /////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GAP_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticle;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Nature de document
      /////////////////////////////////////////////////////////////////////////////
      if result = IntegOk then
      begin
        StRech := TPST.GetValue('GAP_NATUREPIECEG');
        SQL := 'Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrNaturePiece;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ArticlePiece := TOB.Create('ARTICLEPIECE', nil, -1);
    Tob_ArticlePiece.Dupliquer(TPST, False, True, True);
    Tob_ArticlePiece.SetAllModifie(True);

    CodeArticle := Tob_ArticlePiece.GetValue('GAP_ARTICLE');
    Nature := Tob_ArticlePiece.GetValue('GAP_NATUREPIECEG');
    SQL := 'Select GAP_LIBELLE From ARTICLEPIECE WHERE GAP_ARTICLE="' + CodeArticle + '" AND GAP_NATUREPIECEG="' + Nature + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ArticlePiece.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ArticlePiece.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ArticlePiece.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ARTICLETIERS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBArticleTiers(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ArticleTiers: TOB;
  StRech: string;
  CodeArticle: string;
  CodeTiers: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      //  Code Article
      /////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GAT_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_ARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticle;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Tiers
      /////////////////////////////////////////////////////////////////////////////
      if result = IntegOk then
      begin
        StRech := TPST.GetValue('GAT_REFTIERS');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiers;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ArticleTiers := TOB.Create('ARTICLETIERS', nil, -1);
    Tob_ArticleTiers.Dupliquer(TPST, False, True, True);
    Tob_ArticleTiers.SetAllModifie(True);

    CodeArticle := Tob_ArticleTiers.GetValue('GAT_ARTICLE');
    CodeTiers := Tob_ArticleTiers.GetValue('GAT_REFTIERS');
    SQL := 'Select GAT_REFARTICLE From ARTICLETIERS WHERE GAT_ARTICLE="' + CodeArticle + '" AND GAT_REFTIERS="' + CodeTiers + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ArticleTiers.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ArticleTiers.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ArticleTiers.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Santucci
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MEMORISATION
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBMemorisation (TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Memo: TOB;
  StRech: string;
  CodeMemo: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Memo := TOB.Create('BMEMORISATION', nil, -1);
    Tob_Memo.Dupliquer(TPST, False, True, True);
    Tob_Memo.SetAllModifie(True);

    CodeMemo := Tob_Memo.GetValue('BMO_CODEMEMO');
    SQL := 'Select BMO_CODEMEMO From BMEMORISATION WHERE BMO_CODEMEMO="' + CodeMemo + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Memo.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Memo.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Memo.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CATALOGU
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCatalogu(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Catalogu: TOB;
  StRech: string;
  Reference: string;
  CodeTiers: string;
  SQL: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////
      //  Code Article
      /////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GCA_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticle;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Tiers
      /////////////////////////////////////////////////////////////////////////////
      if result = IntegOk then
      begin
        StRech := TPST.GetValue('GCA_TIERS');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiers;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Catalogu := TOB.Create('CATALOGU', nil, -1);
    Catalogu.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Prix de base
      Montant := Catalogu.GetValue('GCA_PRIXBASE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Catalogu.PutValue('GCA_PRIXBASE', MontantDev);
      // Dernier prix d'achat
      Montant := Catalogu.GetValue('GCA_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Catalogu.PutValue('GCA_DPA', MontantDev);
      // Prix de vente public
      Montant := Catalogu.GetValue('GCA_PRIXVENTE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Catalogu.PutValue('GCA_PRIXVENTE', MontantDev);
    end;
    //
    // Mise à jour de tous les champs
    //
    Catalogu.SetAllModifie(True);

    Reference := Catalogu.GetValue('GCA_REFERENCE');
    CodeTiers := Catalogu.GetValue('GCA_TIERS');
    SQL := 'Select GCA_LIBELLE From CATALOGU WHERE GCA_REFERENCE="' + Reference + '" AND GCA_TIERS="' + CodeTiers + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Catalogu.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Catalogu.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Catalogu.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHANCELL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBChancell(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Chancell: TOB;
  CodeDevise: string;
  DateCours: TDateTime;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ///////////////////////////////////////////////////////////////////////////
      // Contrôle de l'existence de la devise
      ///////////////////////////////////////////////////////////////////////////
      CodeDevise := TPST.GetValue('H_DEVISE');
      SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + CodeDevise + '"';
      Q := OpenSQL(SQL, True);
      if Q.EOF then Result := ErrDevise;
      Ferme(Q);
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Chancell := TOB.Create('CHANCELL', nil, -1);
    Tob_Chancell.Dupliquer(TPST, False, True, True);
    Tob_Chancell.SetAllModifie(True);

    CodeDevise := Tob_Chancell.GetValue('H_DEVISE');
    DateCours := Tob_Chancell.GetValue('H_DATECOURS');
    SQL := 'Select H_TAUXREEL From CHANCELL WHERE H_DEVISE="' + CodeDevise + '" AND H_DATECOURS="' + UsDateTime(DateCours) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Chancell.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Chancell.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Chancell.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHOIXCOD
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBChoixcod(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Choixcod: TOB;
  CCType: string;
  CCCode: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ChoixCod := TOB.Create('CHOIXCOD', nil, -1);
    Tob_ChoixCod.Dupliquer(TPST, False, True, True);
    Tob_ChoixCod.SetAllModifie(True);

    CCType := Tob_ChoixCod.GetValue('CC_TYPE');
    CCCode := Tob_ChoixCod.GetValue('CC_CODE');

    ////////////////////////////////////////////////////////////////
    // MODIF LM 28112002 - Suppression des enregistrements CHOIXCOD
    ////////////////////////////////////////////////////////////////
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and (CCType <> DernierTypeChoixcod) then
    begin
      SQL := 'DELETE From CHOIXCOD WHERE CC_TYPE="' + CCType + '"';
      ExecuteSQL(SQL);
      // Sauvegarde du dernier type CHOIXCOD supprimé
      DernierTypeChoixcod := CCType;
    end;

    SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="' + CCType + '" AND CC_CODE="' + CCCode + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ChoixCod.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ChoixCod.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ChoixCod.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CHOIXEXT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBChoixExt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ChoixExt: TOB;
  YXType: string;
  YXCode: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ChoixExt := TOB.Create('CHOIXEXT', nil, -1);
    Tob_ChoixExt.Dupliquer(TPST, False, True, True);
    Tob_ChoixExt.SetAllModifie(True);

    YXType := Tob_ChoixExt.GetValue('YX_TYPE');
    YXCode := Tob_ChoixExt.GetValue('YX_CODE');

    //////////////////////////////////////////////////////////////////
    // MODIF LM 28/11/2002 - Suppression des enregistrements CHOIXEXT
    //////////////////////////////////////////////////////////////////
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and (YXType <> DernierTypeChoixext) then
    begin
      SQL := 'DELETE From CHOIXEXT WHERE YX_TYPE="' + YXType + '"';
      ExecuteSQL(SQL);
      // Sauvegarde du dernier type CHOIXCOD supprimé
      DernierTypeChoixext := YXType;
    end;

    SQL := 'Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="' + YXType + '" AND YX_CODE="' + YXCode + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ChoixExt.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ChoixExt.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ChoixExt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CLAVIERECRAN
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBClavierEcran(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ClavierEcran: TOB;
  Caisse: string;
  Numero: integer;
  Page: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ClavierEcran := TOB.Create('CLAVIERECRAN', nil, -1);
    Tob_ClavierEcran.Dupliquer(TPST, False, True, True);
    Tob_ClavierEcran.SetAllModifie(True);

    Caisse := Tob_ClavierEcran.GetValue('CE_CAISSE');
    Page := Tob_ClavierEcran.GetValue('CE_PAGE');
    Numero := Tob_ClavierEcran.GetValue('CE_NUMERO');

    ////////////////////////////////////////////////////////////////
     // MODIF LM 28112002 - Suppression des enregistrements CLAVIERECRAN
     ////////////////////////////////////////////////////////////////
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and (Caisse <> DernierCaisseSupp) then
    begin
      SQL := 'DELETE From CLAVIERECRAN WHERE CE_CAISSE="' + Caisse + '"';
      ExecuteSQL(SQL);
      DernierCaisseSupp := Caisse;
    end;

    SQL := 'Select CE_CAPTION From CLAVIERECRAN WHERE CE_CAISSE="' + Caisse + '" AND CE_PAGE="' + IntToStr(Page) + '" AND CE_NUMERO="' + IntToStr(Numero) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ClavierEcran.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ClavierEcran.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ClavierEcran.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CODEPOST
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCodepost(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Codepost: TOB;
  Postal: string;
  Ville: string;
  SQL: string;
  Q: TQUERY;
begin
  Result := IntegOK;
  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Codepost := TOB.Create('CODEPOST', nil, -1);
    Tob_Codepost.Dupliquer(TPST, False, True, True);
    Tob_Codepost.SetAllModifie(True);

    Postal := Tob_Codepost.GetValue('O_CODEPOSTAL');
    Ville := Tob_Codepost.GetValue('O_VILLE');
    SQL := 'Select O_PAYS From CODEPOST WHERE O_CODEPOSTAL="' + Postal + '" AND O_VILLE="' + Ville + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Codepost.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Codepost.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Codepost.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 28/06/2004
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB COMMENTAIRE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCommentaire(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Commentaire: TOB;
  Code: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Commentaire := TOB.Create('COMMENTAIRE', nil, -1);
    Tob_Commentaire.Dupliquer(TPST, False, True, True);
    Tob_Commentaire.SetAllModifie(True);

    Code := Tob_Commentaire.GetValue('GCT_CODE');
    SQL := 'SELECT GCT_CODE FROM COMMENTAIRE WHERE GCT_CODE="' + Code + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Commentaire.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Commentaire.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Commentaire.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB COMMERCIAL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCommercial(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Commercial: TOB;
  CodeCommercial: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////////
      // Type de commercial
      ////////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GCL_TYPECOMMERCIAL');
      if (StRech <> '') then
      begin
        SQL := 'Select CO_LIBELLE From COMMUN WHERE CO_TYPE="GTR" AND CO_CODE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrTypeCommercial;
        Ferme(Q);
      end;
      ////////////////////////////////////////////////////////////////////////////
      // Etablissement
      ////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GCL_ETABLISSEMENT');
        if (StRech <> '') then
        begin
          SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrEtablissement;
          Ferme(Q);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////////
      // Zone commerciale
      ////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GCL_ZONECOM');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GZC" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrZoneCommercial;
          Ferme(Q);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////////
      // Type de commissionnement
      ////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GCL_TYPECOM');
        if (StRech <> '') then
        begin
          SQL := 'Select CO_LIBELLE From COMMUN WHERE CO_TYPE="GTC" AND CO_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTypeCommission;
          Ferme(Q);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////////
      // Nature de document
      ////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GCL_NATUREPIECEG');
        if (StRech <> '') then
        begin
          if StRech = 'ZZ1' then // JTR - eQualité 12036
            SQL := 'Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG IN ("FAC", "AVS", "AVC", "FFO")'
            else
            SQL := 'Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrNaturePiece;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Commercial := TOB.Create('COMMERCIAL', nil, -1);
    Tob_Commercial.Dupliquer(TPST, False, True, True);
    Tob_Commercial.SetAllModifie(True);

    CodeCommercial := Tob_Commercial.GetValue('GCL_COMMERCIAL');
    SQL := 'Select GCL_COMMERCIAL From COMMERCIAL WHERE GCL_COMMERCIAL="' + CodeCommercial + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Commercial.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Commercial.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Commercial.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB COMPTEURETAB
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCompteurEtab(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_CompteurEtab: TOB;
  SQL: string;
  StRech: string;
  Etab: string;
  ID: integer;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////////
      // Etablissement
      ////////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GCE_ETABLISSEMENT');
      if (StRech <> '') and (StRech <> DernierEtab) then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_CompteurEtab := TOB.Create('COMPTEURETAB', nil, -1);
    Tob_CompteurEtab.Dupliquer(TPST, False, True, True);
    Tob_CompteurEtab.SetAllModifie(True);

    Etab := Tob_CompteurEtab.GetValue('GCE_ETABLISSEMENT');
    ID := Tob_CompteurEtab.GetValue('GCE_ID');
    SQL := 'Select GCE_ID From COMPTEURETAB WHERE GCE_ETABLISSEMENT="' + Etab + '" AND GCE_ID="' + IntToStr(ID) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_CompteurEtab.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_CompteurEtab.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_CompteurEtab.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CONTACT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBContact(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Contact: TOB;
  TypeContact: string;
  Auxiliaire: string;
  SQL: string;
  StRech: string;
  NumContact: integer;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Civilité
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('C_CIVILITE');
      if (StRech <> '') then
      begin
        SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="CIV" AND CC_CODE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrCivilite;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Lien de parenté
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('C_LIPARENT');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LIP" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrParente;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      //Sexe
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('C_SEXE');
        if (StRech <> '') then
        begin
          SQL := 'Select CO_LIBELLE From COMMUN WHERE CO_TYPE="JMF" AND CO_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrSexe;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Contact := TOB.Create('CONTACT', nil, -1);
    Tob_Contact.Dupliquer(TPST, False, True, True);
    Tob_Contact.SetAllModifie(True);

    TypeContact := Tob_Contact.GetValue('C_TYPECONTACT');
    Auxiliaire := Tob_Contact.GetValue('C_AUXILIAIRE');
    NumContact := Tob_Contact.GetValue('C_NUMEROCONTACT');
    SQL := 'Select C_NOM From CONTACT WHERE C_TYPECONTACT="' + TypeContact + '" AND C_AUXILIAIRE="' + Auxiliaire + '" AND C_NUMEROCONTACT="' + IntToStr(NumContact) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Contact.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Contact.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Contact.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB CTRLCAISMT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBCtrlCaisMt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_CtrlCaisMt: TOB;
  SQL: string;
  Caisse: string;
  ModePaie: string;
  NumZ: integer;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    // Pas de contrôle .......
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_CtrlCaisMt := TOB.Create('CTRLCAISMT', nil, -1);
    Tob_CtrlCaisMt.Dupliquer(TPST, False, True, True);
    Tob_CtrlCaisMt.SetAllModifie(True);

    Caisse := Tob_CtrlCaisMt.GetValue('GJM_CAISSE');
    ModePaie := Tob_CtrlCaisMt.GetValue('GJM_MODEPAIE');
    NumZ := Tob_CtrlCaisMt.GetValue('GJM_NUMZCAISSE');
    SQL := 'Select GJM_DEVISE From CTRLCAISMT WHERE GJM_CAISSE="' + Caisse + '" AND GJM_MODEPAIE="' + ModePaie + '" AND GJM_NUMZCAISSE="' + IntToStr(NumZ) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_CtrlCaisMt.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_CtrlCaisMt.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_CtrlCaisMt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DEPOTS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBDepots(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Depots: TOB;
  CodeDepot: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pays
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GDE_PAYS');
      if (StRech <> '') then
      begin
        SQL := 'Select PY_LIBELLE From PAYS WHERE PY_PAYS="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrPays;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Depots := TOB.Create('DEPOTS', nil, -1);
    Tob_Depots.Dupliquer(TPST, False, True, True);
    Tob_Depots.SetAllModifie(True);

    //
    // doit-on conserver la coche "Géré sur site" ?
    //
    CodeDepot := Tob_Depots.GetValue('GDE_DEPOT');
    {$IFNDEF NOMADE}
    if not ExisteSQL('Select GPK_CAISSE from PARCAISSE where GPK_DEPOT="' + CodeDepot + '"') then
    begin
      Tob_Depots.PutValue('GDE_SURSITE', '-');
    end;
    {$ENDIF}

    SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + CodeDepot + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Depots.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Depots.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Depots.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DEVISE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBDevise(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Devise: TOB;
  CodeDevise: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Devise := TOB.Create('DEVISE', nil, -1);
    Tob_Devise.Dupliquer(TPST, False, True, True);
    Tob_Devise.SetAllModifie(True);

    CodeDevise := Tob_Devise.GetValue('D_DEVISE');
    SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + CodeDevise + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Devise.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Devise.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Devise.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DIMENSION
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBDimension(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Dimension: TOB;
  TypeDim: string;
  GrilleDim: string;
  CodeDim: string;
  SQL: string;
  StRech: string;
  StType: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 1, 2, 3, 4, 5
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GDI_GRILLEDIM');
      if (StRech <> '') then
      begin
        TypeDim := TPST.GetValue('GDI_TYPEDIM');
        if TypeDim = 'DI1' then StType := 'GG1'
        else if TypeDim = 'DI2' then StType := 'GG2'
        else if TypeDim = 'DI3' then StType := 'GG3'
        else if TypeDim = 'DI4' then StType := 'GG4'
        else if TypeDim = 'DI5' then StType := 'GG5';
        SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="' + StType + '" AND CC_CODE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then
        begin
          if TypeDim = 'DI1' then result := ErrGrilleDim1
          else if TypeDim = 'DI2' then result := ErrGrilleDim2
          else if TypeDim = 'DI3' then result := ErrGrilleDim3
          else if TypeDim = 'DI4' then result := ErrGrilleDim4
          else if TypeDim = 'DI5' then result := ErrGrilleDim5;
        end;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Dimension := TOB.Create('DIMENSION', nil, -1);
    Tob_Dimension.Dupliquer(TPST, False, True, True);
    Tob_Dimension.SetAllModifie(True);

    TypeDim := Tob_Dimension.GetValue('GDI_TYPEDIM');
    GrilleDim := Tob_Dimension.GetValue('GDI_GRILLEDIM');
    CodeDim := Tob_Dimension.GetValue('GDI_CODEDIM');
    SQL := 'Select GDI_LIBELLE From DIMENSION WHERE GDI_TYPEDIM="' + TypeDim + '" AND GDI_GRILLEDIM="' + GrilleDim + '" AND GDI_CODEDIM="' + CodeDim + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Dimension.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Dimension.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Dimension.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. :
Suite ........ : Contrôles/Import TOB DIMASQUE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBDimasque(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Dimmasque: TOB;
  CodeMasque: string;
  TypeMasque: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  StRech := '' ;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 1
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GDM_TYPE1');
      if (StRech<>'') then
      begin
        SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GG1" AND CC_CODE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrGrilleDim1;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 2
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        /////////////////////////////////////////////////////////////////////////
        StRech := TPST.GetValue('GDM_TYPE2');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GG2" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrGrilleDim2;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 3
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        /////////////////////////////////////////////////////////////////////////
        StRech := TPST.GetValue('GDM_TYPE3');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GG3" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrGrilleDim3;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 4
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        /////////////////////////////////////////////////////////////////////////
        StRech := TPST.GetValue('GDM_TYPE4');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GG4" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrGrilleDim4;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Grille dimension 5
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        /////////////////////////////////////////////////////////////////////////
        StRech := TPST.GetValue('GDM_TYPE5');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GG5" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrGrilleDim5;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
     ////////////////////////////////////////////////////////////////
    Tob_Dimmasque := TOB.Create('DIMMASQUE', nil, -1);
    Tob_Dimmasque.Dupliquer(TPST, False, True, True);
    Tob_Dimmasque.SetAllModifie(True);

    CodeMasque := Tob_Dimmasque.GetValue('GDM_MASQUE');
    TypeMasque := Tob_Dimmasque.GetValue('GDM_TYPEMASQUE');
    SQL := 'Select GDM_MASQUE From DIMMASQUE WHERE GDM_MASQUE="' + CodeMasque + '" AND GDM_TYPEMASQUE="' + TypeMasque + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Dimmasque.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Dimmasque.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Dimmasque.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB DISPO
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBDispo(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Dispo: TOB;
  DateCloture: TDateTime;
  CodeArticle: string;
  CodeDepot: string;
  Cloture: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      if not (IntegrationRapide) then // PHASE DE CONTROLE
      begin
        /////////////////////////////////////////////////////////////////////////
        // Code Article
        /////////////////////////////////////////////////////////////////////////
        if (result = IntegOK) then
        begin
          StRech := TPST.GetValue('GQ_ARTICLE');
          if (StRech <> '') then
          begin
            SQL := 'Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
            Q := OpenSQL(SQL, True);
            if Q.EOF then result := ErrArticle;
            Ferme(Q);
          end;
        end;
        /////////////////////////////////////////////////////////////////////////
        // Code dépôt
        /////////////////////////////////////////////////////////////////////////
        if (result = IntegOK) then
        begin
          StRech := TPST.GetValue('GQ_DEPOT');
          if (StRech <> '') then
          begin
            SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
            Q := OpenSQL(SQL, True);
            if Q.EOF then result := ErrDepot;
            Ferme(Q);
          end;
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Dispo := TOB.Create('DISPO', nil, -1);
    Tob_Dispo.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Dernier prix d'achat
      Montant := Tob_Dispo.GetValue('GQ_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Dispo.PutValue('GQ_DPA', MontantDev);
      // Prix moyen pondéré d'achat
      Montant := Tob_Dispo.GetValue('GQ_PMAP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Dispo.PutValue('GQ_PMAP', MontantDev);
      // Dernier prix de revient
      Montant := Tob_Dispo.GetValue('GQ_DPR');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Dispo.PutValue('GQ_DPR', MontantDev);
      // Prix moyen pondéré de revient
      Montant := Tob_Dispo.GetValue('GQ_PMRP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Dispo.PutValue('GQ_PMRP', MontantDev);
    end;
    //
    // Mise à jour de tous les champs
    //
    Tob_dispo.SetAllModifie(True);

    Tob_Dispo.PutValue('GQ_DATEINTEGR', NowH);
    Tob_Dispo.SetDateModif(Tob_Dispo.GetValue('GQ_DATEMODIF'));

    CodeArticle := Tob_Dispo.GetValue('GQ_ARTICLE');
    CodeDepot := Tob_Dispo.GetValue('GQ_DEPOT');
    Cloture := Tob_Dispo.GetValue('GQ_CLOTURE');
    DateCloture := Tob_Dispo.GetValue('GQ_DATECLOTURE');
    SQL := 'Select GQ_STOCKMIN From DISPO WHERE GQ_ARTICLE="' + CodeArticle + '" AND GQ_DEPOT="' + CodeDepot + '" AND GQ_CLOTURE="' + Cloture + '" AND GQ_DATECLOTURE="' +
      UsDateTime(DateCloture) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Dispo.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Dispo.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Dispo.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ETABLISS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBEtabliss(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Etab: TOB;
  CodeEtab: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pays
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('ET_PAYS');
      if (StRech <> '') then
      begin
        SQL := 'Select PY_LIBELLE From PAYS WHERE PY_PAYS="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrPays;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('ET_LANGUE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrLangue;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Etab := TOB.Create('ETABLISS', nil, -1);
    Tob_Etab.Dupliquer(TPST, False, True, True);
    Tob_Etab.SetAllModifie(True);

    //
    // doit-on conserver la coche "Géré sur site" ?
    //
    CodeEtab := Tob_Etab.GetValue('ET_ETABLISSEMENT');
    {$IFNDEF NOMADE}
    if not ExisteSQL('Select GPK_CAISSE from PARCAISSE where GPK_ETABLISSEMENT="' + CodeEtab + '"') then
    begin
      Tob_Etab.PutValue('ET_SURSITE', '-');
    end;
    {$ENDIF}

    SQL := 'Select ET_ETABLISSEMENT From ETABLISS WHERE ET_ETABLISSEMENT="' + CodeEtab + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Etab.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Etab.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Etab.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB EXERCICE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBExercice(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Exer: TOB;
  CodeExer: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Exer := TOB.Create('EXERCICE', nil, -1);
    Tob_Exer.Dupliquer(TPST, False, True, True);
    Tob_Exer.SetAllModifie(True);

    CodeExer := Tob_Exer.GetValue('EX_EXERCICE');
    SQL := 'Select EX_EXERCICE From EXERCICE Where EX_EXERCICE="' + CodeExer + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Exer.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Exer.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Exer.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB Fidelité entête
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBFideliteEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Fident: TOB;
  DateFerme: TDateTime;
  CodeTiers: string;
  CodeEtab: string;
  FlagFerme: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      //
      // Contrôles du tiers
      //
      StRech := TPST.GetValue('GFE_TIERS');
      if Trim(StRech) = '' then Result := ErrTiersVide;

      if (result = IntegOK) then
      begin
        if (StRech <> DernierTiers) then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiers
          else DernierTiers := StRech;
          Ferme(Q);
        end;
      end;
      //
      // Controle de l'établissement
      //
      StRech := TPST.GetValue('GFE_ETABLISSEMENT');
      if (StRech <> '') and (StRech <> DernierEtab) then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement
        else DernierEtab := StRech;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Fident := TOB.Create('FIDELITEENT', nil, -1);
    Tob_Fident.Dupliquer(TPST, False, True, True);
    Tob_Fident.SetAllModifie(True);
    //
    // Conservation de la date de modification et Maj de la date d'intégration
    //
    Tob_Fident.SetDateModif(Tob_Fident.GetValue('GFE_DATEMODIF'));
    Tob_Fident.PutValue('GFE_DATEINTEGR', NowH);

    CodeTiers := Tob_Fident.GetValue('GFE_TIERS');
    CodeEtab := Tob_Fident.GetValue('GFE_ETABLISSEMENT');
    FlagFerme := Tob_Fident.GetValue('GFE_FERME');
    DateFerme := Tob_Fident.GetValue('GFE_DATEFERMETURE');

    SQL := 'Select GFE_TIERS From FIDELITEENT WHERE GFE_TIERS="' + CodeTiers + '" AND GFE_ETABLISSEMENT="' + CodeEtab + '" AND GFE_FERME="' + FlagFerme +
      '" AND GFE_DATEFERMETURE="' + UsDateTime(DateFerme) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Fident.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Fident.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Fident.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB Fidelité ligne
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBFideliteLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_FidLig: TOB;
  NumCarte: string;
  CodeEtab: string;
  NumLigne: integer;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      //
      // Controle de l'établissement
      //
      StRech := TPST.GetValue('GFI_ETABLISSEMENT');
      if (StRech <> '') and (StRech <> DernierEtab) then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement
        else DernierEtab := StRech;
        Ferme(Q);
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_FidLig := TOB.Create('FIDELITELIG', nil, -1);
    Tob_FidLig.Dupliquer(TPST, False, True, True);
    Tob_FidLig.SetAllModifie(True);
    //
    // Conservation de la date de modification et Maj de la date d'intégration
    //
    Tob_FidLig.SetDateModif(Tob_FidLig.GetValue('GFI_DATEMODIF'));
    Tob_FidLig.PutValue('GFI_DATEINTEGR', NowH);

    NumCarte := Tob_FidLig.GetValue('GFI_NUMCARTEINT');
    CodeEtab := Tob_FidLig.GetValue('GFI_ETABLISSEMENT');
    NumLigne := Tob_FidLig.GetValue('GFI_LIGNE');

    SQL := 'Select GFI_ETABLISSEMENT From FIDELITELIG WHERE GFI_NUMCARTEINT="' + NumCarte + '" AND GFI_ETABLISSEMENT="' + CodeEtab + '" AND GFI_LIGNE="' + IntToStr(NumLigne) +
      '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_FidLig.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_FidLig.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_FidLig.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB FONCTION
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBFonction(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Fonc: TOB;
  CodeFonc: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Fonc := TOB.Create('FONCTION', nil, -1);
    Tob_Fonc.Dupliquer(TPST, False, True, True);
    Tob_Fonc.SetAllModifie(True);

    CodeFonc := Tob_Fonc.GetValue('AFO_FONCTION');
    SQL := 'Select AFO_FONCTION From FONCTION WHERE AFO_FONCTION="' + CodeFonc + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Fonc.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Fonc.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Fonc.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB JOURSCAISSE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBJoursCaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_JoursCaisse: TOB;
  Caisse: string;
  NumOuverture: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôles ...............
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_JoursCaisse := TOB.Create('JOURSCAISSE', nil, -1);
    Tob_JoursCaisse.Dupliquer(TPST, False, True, True);
    Tob_JoursCaisse.SetAllModifie(True);

    Caisse := Tob_JoursCaisse.GetValue('GJC_CAISSE');
    NumOuverture := Tob_JoursCaisse.GetValue('GJC_NUMZCAISSE');

    SQL := 'Select GJC_ETAT From JOURSCAISSE WHERE GJC_CAISSE="' + Caisse + '" AND GJC_NUMZCAISSE="' + IntToStr(NumOuverture) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_JoursCaisse.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_JoursCaisse.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_JoursCaisse.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB JOURSETAB
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBJoursEtab(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_JoursEtab: TOB;
  CodeEta: string;
  Journee: TDateTime;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // ETABLISSEMENT
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GJE_ETABLISSEMENT');
      if (StRech <> '') then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Code météo
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GJE_METEO');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GME" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrMeteo;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_JoursEtab := TOB.Create('JOURSETAB', nil, -1);
    Tob_JoursEtab.Dupliquer(TPST, False, True, True);
    Tob_JoursEtab.SetAllModifie(True);

    CodeEta := Tob_JoursEtab.GetValue('GJE_ETABLISSEMENT');
    Journee := Tob_JoursEtab.GetValue('GJE_JOURNEE');

    SQL := 'Select GJE_METEO From JOURSETAB WHERE GJE_ETABLISSEMENT="' + CodeEta + '" AND GJE_JOURNEE="' + UsDateTime(Journee) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_JoursEtab.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_JoursEtab.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_JoursEtab.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB JOURSETABEVT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBJoursEtabEvt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_JoursEtabEvt: TOB;
  Etab: string;
  CodeEvt: string;
  NoEvent: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôles ...............
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_JoursEtabEvt := TOB.Create('JOURSETABEVT', nil, -1);
    Tob_JoursEtabEvt.Dupliquer(TPST, False, True, True);
    Tob_JoursEtabEvt.SetAllModifie(True);

    Etab := Tob_JoursEtabEvt.GetValue('GET_ETABLISSEMENT');
    NoEvent := Tob_JoursEtabEvt.GetValue('GET_NOEVENT');
    CodeEvt := Tob_JoursEtabEvt.GetValue('GET_CODEEVENT');

    SQL := 'Select GET_NOEVENT From JOURSETABEVT WHERE GET_ETABLISSEMENT="' + Etab + '" AND GET_NOEVENT="' + IntToStr(NoEvent) + '" AND GET_CODEEVENT="' + CodeEvt + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_JoursEtabEvt.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_JoursEtabEvt.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_JoursEtabEvt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIENSOLE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBLiensOLE(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_LiensOLE: TOB;
  Table: string;
  Ident: string;
  Rang: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_LiensOLE := TOB.Create('LIENSOLE', nil, -1);
    Tob_LiensOLE.Dupliquer(TPST, False, True, True);
    Tob_LiensOLE.SetAllModifie(True);

    Table := Tob_LiensOLE.GetValue('LO_TABLEBLOB');
    Ident := Tob_LiensOLE.GetValue('LO_IDENTIFIANT');
    Rang := Tob_LiensOLE.GetValue('LO_RANGBLOB');

    SQL := 'Select LO_EMPLOIBLOB From LIENSOLE WHERE LO_TABLEBLOB="' + Table + '" AND LO_IDENTIFIANT="' + Ident + '" AND LO_RANGBLOB="' + IntToStr(Rang) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_LiensOLE.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_LiensOLE.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_LiensOLE.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIGNE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBLigne(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  TobFilleArticle: TOB;
  CodeArt: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  ArticleDefaut: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////////////////////////////
      // Naturez de pièces, Tiers, tiers livré, payeur, facturé : on considère qu'ils sont identiques
      // à ceux de l'entête pour limiter les requètes de contrôles.
      ////////////////////////////////////////////////////////////////////////////////////////////////
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GL_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle
          else if TPST.GetValue('GL_CODEARTICLE') <> Q.Findfield('GA_CODEARTICLE').AsString then result := ErrArticleGen;
          Ferme(Q);

          // Doit on forcer l'article pour intégrer un ticket et récupérer le CA ?
          //if (result = ErrArticle) and (TPST.GetValue ('GL_NATUREPIECEG') = 'FFO') then
          // 05/02/2003 PCP - Activation pour toutes les natures de pièces
          if (result = ErrArticle) then
          begin
            if GetParamSoc('SO_GCTOXFORCEARTICLE') then
            begin
              ArticleDefaut := GetParamSoc('SO_GCTOXARTICLE');

              SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + ArticleDefaut + '"';
              Q := OpenSQL(SQL, True);
              if not Q.EOF then
              begin
                TPST.PutValue('GL_ARTICLE', ArticleDefaut);
                TPST.PutValue('GL_CODEARTICLE', Q.Findfield('GA_CODEARTICLE').AsString);
                result := IntegOK;
              end;
              Ferme(Q);
            end;
          end;
        end;
      end;
      /////////////////////////////////////////////////////////////////
      // Tarif Article
      /////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GL_TARIFARTICLE');
        if (StRech <> '') and (StRech <> DernierTarifArticle) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TAR" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTarifArticle
          else DernierTarifArticle := StRech;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
    //  PHASE D'INTEGRATION
    //  1 - Création de la TOB lignes
    //    - Conversion en devise
 //////////////////////////////////////////////////////////////////////////////////////
    Tob_Ligne := TOB.Create('LIGNE', nil, -1);
    Tob_Ligne.Dupliquer(TPST, False, True, True);

    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Taux et Cotation
      Tob_Ligne.PutValue('GL_TAUXDEV', Tob_Entete.GetValue('GP_TAUXDEV'));
      Tob_Ligne.PutValue('GL_COTATION', Tob_Entete.GetValue('GP_COTATION'));
      // Dernier prix d'achat
      Montant := Tob_Ligne.GetValue('GL_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_DPA', MontantDev);
      // Prix moyen pondéré d'achat
      Montant := Tob_Ligne.GetValue('GL_PMAP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_PMAP', MontantDev);
      // Prix moyen pondéré de revient
      Montant := Tob_Ligne.GetValue('GL_PMRP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_PMRP', MontantDev);
      // PMAP actualisé
      Montant := Tob_Ligne.GetValue('GL_PMAPACTU');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_PMAPACTU', MontantDev);
      // PMRP actualisé
      Montant := Tob_Ligne.GetValue('GL_PMRPACTU');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_PMRPACTU', MontantDev);
      // Dernier prix de revient
      Montant := Tob_Ligne.GetValue('GL_DPR');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Ligne.PutValue('GL_DPR', MontantDev);
    end;
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  2 - Rattachement de la TOB à l'entête
     //  3 - Chargement de la TOB Article associé ainsi que les fiches stocks
     //  4 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
  //////////////////////////////////////////////////////////////////////////////////////
    Tob_Ligne.ChangeParent(Tob_Entete, -1);
    Tob_Ligne.SetAllModifie(True);
    // MODIF LM 03/05/01
    AddLesSupLigne(Tob_Ligne, False);

    CodeArt := Tob_Ligne.GetValue('GL_ARTICLE');
    if CodeArt <> '' then
    begin
      SQL := 'Select * From ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE GA_ARTICLE="' + CodeArt + '"';
      Q := OpenSQL(SQL, True);
      if not Q.EOF then
      begin
        TobFilleArticle := CreerTOBArt(Tob_Article);
        TobFilleArticle.SelectDB('', Q);
        // MODIF LM 16/01/2002 pour optimisation : ajout du 3ième argument

        if (Tob_Entete.GetValue('GP_NATUREPIECEG') = 'TEM') or (Tob_Entete.GetValue('GP_NATUREPIECEG') = 'TRE') or (Tob_Entete.GetValue('GP_NATUREPIECEG') = 'TRV') then
        begin
          LoadTOBDispo(TobFilleArticle, True);
        end else
        begin
          LoadTOBDispo(TobFilleArticle, True, CreerQuelDepot(Tob_Entete));
        end;
      end;
      Ferme(Q);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIGNELOT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBLigneLot(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_lig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GLL_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_LigneLot
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Lig := TOB.Create('LIGNELOT', nil, -1);
    Tob_Lig.Dupliquer(TPST, False, True, True);
    Tob_Lig.ChangeParent(Tob_LigneLot, -1);
    Tob_Lig.SetAllModifie(True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LISTEINVENT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBListeInvEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ListeInvEnt: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  CodeListe: string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Depôt
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GIE_DEPOT');
        if (StRech <> '') then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepot;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_ListeInvEnt
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
       ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ListeInvEnt := TOB.Create('LISTEINVENT', nil, -1);
    Tob_ListeInvEnt.Dupliquer(TPST, False, True, True);
    Tob_ListeInvEnt.SetAllModifie(True);

    CodeListe := Tob_ListeInvEnt.GetValue('GIE_CODELISTE');

    SQL := 'Select GIE_LIBELLE From LISTEINVENT WHERE GIE_CODELISTE="' + CodeListe + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ListeInvEnt.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ListeInvEnt.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ListeInvEnt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LISTEINVLIG
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBListeInvLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ListeInvLig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  CodeListe: string;
  CodeArticle: string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Depôt
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GIL_DEPOT');
        if (StRech <> '') then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepot;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////
      //  Code Article
      /////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GIL_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticle;
        Ferme(Q);
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_ListeInvLig
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
       ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ListeInvLig := TOB.Create('LISTEINVLIG', nil, -1);
    Tob_ListeInvLig.Dupliquer(TPST, False, True, True);
    Tob_ListeInvLig.SetAllModifie(True);

    CodeListe := Tob_ListeInvLig.GetValue('GIL_CODELISTE');
    CodeArticle := Tob_ListeInvLig.GetValue('GIL_ARTICLE');

    SQL := 'Select GIL_LOT From LISTEINVLIG WHERE GIL_CODELISTE="' + CodeListe + '" AND GIL_ARTICLE="' + CodeArticle + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ListeInvLig.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ListeInvLig.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ListeInvLig.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIGNENOMEN
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBLigneNomen(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Lig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GLN_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle
          else if TPST.GetValue('GLN_CODEARTICLE') <> Q.Findfield('GA_CODEARTICLE').AsString then result := ErrArticleGen;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_LigneNom
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Lig := TOB.Create('LIGNENOMEN', nil, -1);
    Tob_Lig.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Dernier prix d'achat
      Montant := Tob_Lig.GetValue('GLN_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('GLN_DPA', MontantDev);
      // Prix moyen pondéré d'achat
      Montant := Tob_Lig.GetValue('GLN_PMAP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('GLN_PMAP', MontantDev);
      // Dernier prix de revient
      Montant := Tob_Lig.GetValue('GLN_DPR');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('GLN_DPR', MontantDev);
      // Prix moyen pondéré de revient
      Montant := Tob_Lig.GetValue('GLN_PMRP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('GLN_PMRP', MontantDev);
    end;
    //
    // Rattachement à la TOB mère
    //
    Tob_Lig.ChangeParent(Tob_LigneNom, -1);
    Tob_Lig.SetAllModifie(True);
  end;
end;

{$IFDEF BTP}
function ImportTOBMillieme(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Lig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
begin
  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB des milliemes de TVA
     //  2 - Rattachement de la TOB virtuelles Tob_Millieme
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Lig := TOB.Create('BTPIECEMILIEME', nil, -1);
    Tob_Lig.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    //
    // Rattachement à la TOB mère
    //
    Tob_Lig.ChangeParent(Tob_Millieme, -1);
    Tob_Lig.SetAllModifie(True);
  end;
end;

function ImportTOBPardoc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParDoc: TOB;
  Nature,Souche: string;
  Numero : integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ParDoc := TOB.Create('BTPARDOC', nil, -1);
    Tob_ParDoc.Dupliquer(TPST, False, True, True);
    Tob_ParDoc.SetAllModifie(True);

    Nature := Tob_ParDoc.GetValue('BPD_NATUREPIECE');
    Souche := Tob_ParDoc.GetValue('BPD_SOUCHE');
    Numero := Tob_ParDoc.GetValue('BPD_NUMPIECE');
    SQL := 'Select BPD_NUMPIECE From BTPARDOC WHERE BPD_NATUREPIECE="' + Nature+'" AND BPD_SOUCHE="'+Souche+'" AND BPD_NUMPIECE='+IntToStr(Numero);
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParDoc.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParDoc.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParDoc.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Santucci
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB LIGNEOUV
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBLigneOuvrages(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Lig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('BLO_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle
          else if TPST.GetValue('BLO_CODEARTICLE') <> Q.Findfield('GA_CODEARTICLE').AsString then result := ErrArticleGen;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_LigneOuv
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Lig := TOB.Create('LIGNEOUV', nil, -1);
    Tob_Lig.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Dernier prix d'achat
      Montant := Tob_Lig.GetValue('BLO_DPA');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('BLO_DPA', MontantDev);
      // Prix moyen pondéré d'achat
      Montant := Tob_Lig.GetValue('BLO_PMAP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('BLO_PMAP', MontantDev);
      // Dernier prix de revient
      Montant := Tob_Lig.GetValue('BLO_DPR');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('BLO_DPR', MontantDev);
      // Prix moyen pondéré de revient
      Montant := Tob_Lig.GetValue('BLO_PMRP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('BLO_PMRP', MontantDev);
    end;
    //
    // Rattachement à la TOB mère
    //
    Tob_Lig.ChangeParent(Tob_LigneOuv, -1);
    Tob_Lig.SetAllModifie(True);
  end;
end;

function ImportTOBNaturePrest(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Natureprest: TOB;
  Nature: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Natureprest := TOB.Create('NATUREPREST', nil, -1);
    Tob_Natureprest.Dupliquer(TPST, False, True, True);
    Tob_Natureprest.SetAllModifie(True);

    Nature := Tob_Natureprest.GetValue('BNP_NATUREPRES');
    SQL := 'Select BNP_NATUREPRES From NATUREPREST WHERE BNP_NATUREPRES="' + Nature+'"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Natureprest.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Natureprest.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Natureprest.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Santucci
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB Retenue de garantie
Mots clefs ... : TOX
*****************************************************************}
function ImportTobPieceRg(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Lig: TOB;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
    //
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB lignes
     //  2 - Rattachement de la TOB virtuelles Tob_piecerg
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Lig  := TOB.Create('PIECERG', nil, -1);
    Tob_Lig.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      Montant := Tob_Lig.GetValue('PRG_MTHTRG');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('PRG_DPA', MontantDev);

      Montant := Tob_Lig.GetValue('PRG_MTTTCRG');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('PRG_MTTTCRG', MontantDev);

      Montant := Tob_Lig.GetValue('PRG_CAUTIONMT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('PRG_CAUTIONMT', MontantDev);

      Montant := Tob_Lig.GetValue('PRG_CAUTIONMTU');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Lig.PutValue('PRG_CAUTIONMTU', MontantDev);
    end;
    //
    // Rattachement à la TOB mère
    //
    Tob_Lig.ChangeParent(Tob_PieceRg, -1);
    Tob_Lig.SetAllModifie(True);
  end;
end;

{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MEA
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBMea(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Mea: TOB;
  QualifMesure: string;
  Mesure: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Mea := TOB.Create('MEA', nil, -1);
    Tob_Mea.Dupliquer(TPST, False, True, True);
    Tob_Mea.SetAllModifie(True);

    QualifMesure := Tob_Mea.GetValue('GME_QUALIFMESURE');
    Mesure := Tob_Mea.GetValue('GME_MESURE');
    SQL := 'Select GME_LIBELLE From MEA WHERE GME_QUALIFMESURE="' + QualifMesure + '" AND GME_MESURE="' + Mesure + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Mea.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Mea.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Mea.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MODELES
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBModeles(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Modeles: TOB;
  StRech: string;
  TypeModele: string;
  Nature: string;
  Code: string;
  Langue: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('MO_LANGUE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrLangue;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Modeles := TOB.Create('MODELES', nil, -1);
    Tob_Modeles.Dupliquer(TPST, False, True, True);
    Tob_Modeles.SetAllModifie(True);

    TypeModele := Tob_Modeles.GetValue('MO_TYPE');
    Nature := Tob_Modeles.GetValue('MO_NATURE');
    Code := Tob_Modeles.GetValue('MO_CODE');
    Langue := Tob_Modeles.GetValue('MO_LANGUE');

    SQL := 'Select MO_LIBELLE From MODELES WHERE MO_TYPE="' + TypeModele + '" AND MO_NATURE="' + Nature + '" AND MO_CODE="' + Code + '" AND MO_LANGUE="' + Langue + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Modeles.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Modeles.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Modeles.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MODEDATA
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBModeData(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ModeData: TOB;
  Clef: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Contrôles ....
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ModeData := TOB.Create('MODEDATA', nil, -1);
    Tob_ModeData.Dupliquer(TPST, False, True, True);
    Tob_ModeData.SetAllModifie(True);

    Clef := Tob_ModeData.GetValue('MD_CLE');

    SQL := 'Select MD_CLE From MODEDATA WHERE MD_CLE="' + Clef + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ModeData.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ModeData.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ModeData.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MODEREGL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBModeRegl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ModeRegl: TOB;
  CodeRegl: string;
  StRech: string;
  Zone: string;
  SQL: string;
  Q: TQUERY;
  i: integer;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Mode de paiement 1 à 12
      /////////////////////////////////////////////////////////////////////////
      for i := 1 to 12 do
      begin
        Zone := 'MR_MP' + IntToStr(i);
        StRech := TPST.GetValue(Zone);
        if (StRech <> '') then
        begin
          SQL := 'Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then
          begin
            if (i = 1) then result := ErrModePaie1
            else if (i = 2) then result := ErrModePaie2
            else if (i = 3) then result := ErrModePaie3
            else if (i = 4) then result := ErrModePaie4
            else if (i = 5) then result := ErrModePaie5
            else if (i = 6) then result := ErrModePaie6
            else if (i = 7) then result := ErrModePaie7
            else if (i = 8) then result := ErrModePaie8
            else if (i = 9) then result := ErrModePaie9
            else if (i = 10) then result := ErrModePaie10
            else if (i = 11) then result := ErrModePaie11
            else if (i = 12) then result := ErrModePaie12;
          end;
          Ferme(Q);
        end;
        if Result <> IntegOK then break;
      end;
    end;
  end else
  begin // PHASE D'INTEGRATION
    //
    // Suppression des modes de règlement au 1ier enregistrement intégré
    // Permet de prendre en compte les suppressions d'enregistrements
    //
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and not (ModeReglSupp) then
    begin
      SQL := 'DELETE From MODEREGL';
      ExecuteSQL(SQL);
      ModeReglSupp := True;
    end;
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ModeRegl := TOB.Create('MODEREGL', nil, -1);
    Tob_ModeRegl.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;
      // Montant minimum requis
      Montant := Tob_ModeRegl.GetValue('MR_MONTANTMIN');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_ModeRegl.PutValue('MR_MONTANTMIN', MontantDev);
    end;
    //
    // Mise à jour de tous les champs
    //
    Tob_ModeRegl.SetAllModifie(True);

    CodeRegl := Tob_ModeRegl.GetValue('MR_MODEREGLE');
    SQL := 'Select MR_LIBELLE From MODEREGL WHERE MR_MODEREGLE="' + CodeRegl + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ModeRegl.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ModeRegl.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ModeRegl.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB MODEPAIE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBModePaie(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ModePaie: TOB;
  CodePaie: string;
  SQL: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    //
    // Suppression des modes de paiements au 1ier enregistrement intégré
    // Permet de prendre en compte les suppressions d'enregistrements
    //
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and not (ModePaieSupp) then
    begin
      SQL := 'DELETE From MODEPAIE';
      ExecuteSQL(SQL);
      ModePaieSupp := True;
    end;
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_ModePaie := TOB.Create('MODEPAIE', nil, -1);
    Tob_ModePaie.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;
      // Montant maximum
      Montant := Tob_ModePaie.GetValue('MP_MONTANTMAX');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_ModePaie.PutValue('MP_MONTANTMAX', MontantDev);
    end;
    //
    // Mise à jour de tous les champs
    //
    Tob_ModePaie.SetAllModifie(True);

    CodePaie := Tob_ModePaie.GetValue('MP_MODEPAIE');
    SQL := 'Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="' + CodePaie + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ModePaie.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ModePaie.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ModePaie.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB NOMENENT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBNomenEnt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_NomenEnt: TOB;
  CodeArt: string;
  CodeNom: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GNE_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_NomenEnt := TOB.Create('NOMENENT', nil, -1);
    Tob_NomenEnt.Dupliquer(TPST, False, True, True);
    Tob_NomenEnt.SetAllModifie(True);

    CodeArt := Tob_NomenEnt.GetValue('GNE_ARTICLE');
    CodeNom := Tob_NomenEnt.GetValue('GNE_NOMENCLATURE');
    SQL := 'Select GNE_LIBELLE From NOMENENT WHERE GNE_ARTICLE="' + CodeArt + '" AND GNE_NOMENCLATURE="' + CodeNom + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_NomenEnt.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_NomenEnt.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_NomenEnt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB NOMENLIG
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBNomenLig(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_NomenLig: TOB;
  CodeNom: string;
  SQL: string;
  StRech: string;
  NumLigne: integer;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GNL_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_NomenLig := TOB.Create('NOMENLIG', nil, -1);
    Tob_NomenLig.Dupliquer(TPST, False, True, True);
    Tob_NomenLig.SetAllModifie(True);

    CodeNom := Tob_NomenLig.GetValue('GNL_NOMENCLATURE');
    NumLigne := Tob_NomenLig.GetValue('GNL_NUMLIGNE');
    SQL := 'Select GNL_LIBELLE From NOMENLIG WHERE GNL_NOMENCLATURE="' + CodeNom + '" AND GNL_NUMLIGNE="' + IntToStr(NumLigne) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_NomenLig.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_NomenLig.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_NomenLig.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PARCAISSE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBOperCaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_OperCaisse: TOB;
  StRech: string;
  SQL: string;
  Nature: string;
  Souche: string;
  Prefixe: string;
  Numero: integer;
  Indice: integer;
  NumLigne: integer;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
    if not (IntegrationRapide) then
    begin
      /////////////////////////////////////////////////////////////////////////
      // Etablissement
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GOC_ETABLISSEMENT');
      if (StRech <> '') and (StRech <> DernierEtab) then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////
      //  Code Article de référence
      /////////////////////////////////////////////////////////////////
      if result = IntegOk then
      begin
        StRech := TPST.GetValue('GOC_ARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select GA_TYPEARTICLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrArticle;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Tiers
      /////////////////////////////////////////////////////////////////////////////
      if result = IntegOk then
      begin
        StRech := TPST.GetValue('GOC_TIERS');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiers;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_OperCaisse := TOB.Create('OPERCAISSE', nil, -1);
    Tob_OperCaisse.Dupliquer(TPST, False, True, True);
    Tob_OperCaisse.SetAllModifie(True);
    Tob_OperCaisse.SetDateModif(Tob_OperCaisse.GetValue('GOC_DATEMODIF'));
    Tob_OperCaisse.PutValue('GOC_DATEINTEGR', NowH);

    Nature := Tob_OperCaisse.GetValue('GOC_NATUREPIECEG');
    Souche := Tob_OperCaisse.GetValue('GOC_SOUCHE');
    Numero := Tob_OperCaisse.GetValue('GOC_NUMERO');
    Indice := Tob_OperCaisse.GetValue('GOC_INDICEG');
    NumLigne := Tob_OperCaisse.GetValue('GOC_NUMLIGNE');
    Prefixe := Tob_OperCaisse.GetValue('GOC_PREFIXE');

    SQL := 'Select GOC_SOUCHE From OPERCAISSE WHERE GOC_NATUREPIECEG="' + Nature + '" AND GOC_SOUCHE="' + Souche + '" AND GOC_NUMERO="' + IntToStr(Numero) + '" AND GOC_INDICEG="'
      + IntToStr(Indice) + '" AND GOC_NUMLIGNE="' + IntToStr(Numligne) + '" AND GOC_PREFIXE="' + Prefixe + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_OperCaisse.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_OperCaisse.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_OperCaisse.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PARCAISSE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParCaisse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParCaisse: TOB;
  Caisse: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
    if not (IntegrationRapide) then
    begin
      /////////////////////////////////////////////////////////////////////////
      // Etablissement
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GPK_ETABLISSEMENT');
      if (StRech <> '') then
      begin
        SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrEtablissement;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Dépôt
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GPK_DEPOT');
        if (StRech <> '') then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepot;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_ParCaisse := TOB.Create('PARCAISSE', nil, -1);
    Tob_ParCaisse.Dupliquer(TPST, False, True, True);
    Tob_Parcaisse.SetAllModifie(True);
    Tob_Parcaisse.SetDateModif(Tob_ParCaisse.GetValue('GPK_DATEMODIF'));
    Tob_ParCaisse.PutValue('GPK_DATEINTEGR', NowH);

    Caisse := Tob_ParCaisse.GetValue('GPK_CAISSE');
    SQL := 'Select GPK_LIBELLE From PARCAISSE WHERE GPK_CAISSE="' + Caisse + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParCaisse.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParCaisse.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParCaisse.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Santucci
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Import ParamSoc
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParamSoc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParSoc: TOB;
  ParamSoc: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_ParSoc := TOB.Create('PARAMSOC', nil, -1);
    Tob_ParSoc.Dupliquer(TPST, False, True, True);
    Tob_ParSoc.SetAllModifie(True);
    ParamSoc := Tob_ParSoc.GetValue('SOC_NOM');
    SQL := 'Select SOC_NOM From PARAMSOC WHERE SOC_NOM="' + ParamSoc + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParSoc.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParSoc.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParSoc.free;
  end;
end;

function ImportTOBParamTaxe(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParamTaxe: TOB;
  ParamTaxe: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_ParamTaxe := TOB.Create('PARAMTAXE', nil, -1);
    Tob_ParamTaxe.Dupliquer(TPST, False, True, True);
    Tob_ParamTaxe.SetAllModifie(True);
    ParamTaxe := Tob_ParamTaxe.GetValue('BPT_CATEGORIETAXE');
    SQL := 'Select BPT_CATEGORIETAXE From PARAMTAXE WHERE BPT_CATEGORIETAXE="' + ParamTaxe + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParamTaxe.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParamTaxe.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParamTaxe.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Suppression
Mots clefs ... : TOX
*****************************************************************}
procedure SupprimeParamFidelite(CodeFid: string);
var
  SQL: string;
begin
  SQL := 'DELETE From PARREGLEFID WHERE GFR_CODEFIDELITE="' + CodeFid + '"';
  ExecuteSQL(SQL);
  SQL := 'DELETE From PARSEUILFID WHERE GFS_REGLEFIDELITE="' + CodeFid + '"';
  ExecuteSQL(SQL);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ParFidelite
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParFidelite(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParFidelite: TOB;
  CodeFidelite: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
    if not (IntegrationRapide) then
    begin
      //
      // Mode de paiement
      //
      StRech := TPST.GetValue('GFO_MODEPAIE');
      if (StRech <> '') and (StRech <> DernierModePaie) then
      begin
        SQL := 'Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrModePaie
        else DernierModePaie := StRech;
        Ferme(Q);
      end;
      //
      // Type de démarque
      //
      if result = IntegOK then
      begin
        StRech := TPST.GetValue('GFO_TYPEREMISE');
        if (StRech <> '') and (StRech <> DernierTypeRemise) then
        begin
          SQL := 'Select GTR_LIBELLE From TYPEREMISE WHERE GTR_TYPEREMISE="' + StRech + '"';
          if not ExisteSQL(SQL) then result := ErrTypeDemarque
          else DernierTypeRemise := StRech;
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Suppression des paramètres de fidélité associés
    // Attention : On travaille en annule et remplace,
    //             les tables PARSEUILFID, PARREGLEFID doivent
    //             impérativement être en requêtes de sous-niveaux
    ////////////////////////////////////////////////////////////////
    CodeFidelite := TPST.GetValue('GFO_CODEFIDELITE');
    SupprimeParamFidelite(CodeFidelite);

    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_ParFidelite := TOB.Create('PARFIDELITE', nil, -1);
    Tob_ParFidelite.Dupliquer(TPST, False, True, True);
    Tob_ParFidelite.SetAllModifie(True);
    //
    // Conservation de la date de modification et Maj de la date d'intégration
    //
    Tob_ParFidelite.SetDateModif(Tob_ParFidelite.GetValue('GFO_DATEMODIF'));
    Tob_ParFidelite.PutValue('GFO_DATEINTEGR', NowH);

    CodeFidelite := Tob_ParFidelite.GetValue('GFO_CODEFIDELITE');

    SQL := 'Select GFO_FERME From PARFIDELITE WHERE GFO_CODEFIDELITE="' + CodeFidelite + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParFidelite.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParFidelite.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParFidelite.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB ParFidelite
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParRegleFid(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParRegleFid: TOB;
  CodeFidelite: string;
  RegleFidelite: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
    if not (IntegrationRapide) then
    begin
      // Pas de contrôle pour le moment
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
 ////////////////////////////////////////////////////////////////
    Tob_ParRegleFid := TOB.Create('PARREGLEFID', nil, -1);
    Tob_ParRegleFid.Dupliquer(TPST, False, True, True);
    Tob_ParRegleFid.SetAllModifie(True);
    //
    // Conservation de la date de modification et Maj de la date d'intégration
    //
    Tob_ParRegleFid.SetDateModif(Tob_ParRegleFid.GetValue('GFR_DATEMODIF'));

    CodeFidelite := Tob_ParRegleFid.GetValue('GFR_CODEFIDELITE');
    RegleFidelite := Tob_ParRegleFid.GetValue('GFR_REGLEFIDELITE');

    SQL := 'Select GFR_FERME From PARREGLEFID WHERE GFR_CODEFIDELITE="' + CodeFidelite + '" AND GFR_REGLEFIDELITE="' + RegleFidelite + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParRegleFid.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParRegleFid.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParRegleFid.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB Paramétrage des seuils de Fidelite
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParSeuilFid(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParSeuilFid: TOB;
  CodeFidelite: string;
  RegleFidelite: string;
  Rang: integer;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then // PHASE DE CONTROLE
  begin
    if not (IntegrationRapide) then
    begin
      // Pas de contrôle pour le moment
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_ParSeuilFid := TOB.Create('PARSEUILFID', nil, -1);
    Tob_ParSeuilFid.Dupliquer(TPST, False, True, True);
    Tob_ParSeuilFid.SetAllModifie(True);
    //
    // Conservation de la date de modification et Maj de la date d'intégration
    //
    Tob_ParSeuilFid.SetDateModif(Tob_ParSeuilFid.GetValue('GFS_DATEMODIF'));

    CodeFidelite := Tob_ParSeuilFid.GetValue('GFS_CODEFIDELITE');
    RegleFidelite := Tob_ParSeuilFid.GetValue('GFS_REGLEFIDELITE');
    Rang := Tob_ParSeuilFid.GetValue('GFS_RANG');

    SQL := 'Select GFS_RANG From PARSEUILFID WHERE GFS_CODEFIDELITE="' + CodeFidelite + '" AND GFS_REGLEFIDELITE="' + RegleFidelite + '" AND AND GFS_RANG="' + IntToStr(Rang) +
      '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParSeuilFid.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParSeuilFid.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParSeuilFid.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PARCPIECBIL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBParPiecBil(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_ParPiecBil: TOB;
  Devise: string;
  TypePB: string;
  PieceBillet: integer;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      StRech := TPST.GetValue('GPI_DEVISE');
      if (StRech <> '') then
      begin
        SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrDevise;
        Ferme(Q);
      end;
    end;
  end else
  begin // PHASE D'INTEGRATION
    //
    // Suppression des pièces et billets au 1ier enregistrement intégré
    // Permet de prendre en compte les suppressions d'enregistrements
    //
    if (GetParamSoc('SO_GCTOXANNULEPARAM')) and not (ParPiecBilSupp) then
    begin
      SQL := 'DELETE From PARPIECBIL';
      ExecuteSQL(SQL);
      ParPiecBilSupp := True;
    end;
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
 ////////////////////////////////////////////////////////////////
    Tob_ParPiecBil := TOB.Create('PARPIECBIL', nil, -1);
    Tob_ParPiecBil.Dupliquer(TPST, False, True, True);
    Tob_ParPiecBil.SetAllModifie(True);

    Devise := Tob_ParPiecBil.GetValue('GPI_DEVISE');
    TypePB := Tob_ParPiecBil.GetValue('GPI_TYPE');
    PieceBillet := Tob_ParPiecBil.GetValue('GPI_PIECEBILLET');

    SQL := 'Select GPI_LIBELLE From PARPIECBIL WHERE GPI_DEVISE="' + Devise + '" AND GPI_TYPE="' + TypePB + '" AND GPI_PIECEBILLET="' + IntToStr(PieceBillet) + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_ParPiecBil.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_ParPiecBil.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_ParPiecBil.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PAYS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPays(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Pays: TOB;
  CodePays: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Devise
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('PY_DEVISE');
      if (StRech <> '') then
      begin
        SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrDevise;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('PY_LANGUE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrLangue;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
   ////////////////////////////////////////////////////////////////
    Tob_Pays := TOB.Create('PAYS', nil, -1);
    Tob_Pays.Dupliquer(TPST, False, True, True);
    Tob_Pays.SetAllModifie(True);

    CodePays := Tob_Pays.GetValue('PY_PAYS');
    SQL := 'Select PY_LIBELLE From PAYS WHERE PY_PAYS="' + CodePays + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Pays.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Pays.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Pays.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PIEDBASE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPiedBase(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Base: TOB;
  SQL     : string;
  Q       : TQUERY;
  StRech  : string;
  StRech2 : string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////
      // Devise
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GPB_DEVISE');
      if (StRech <> '') and (StRech <> DerniereDevise) then
      begin
        SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrDevise
        else DerniereDevise := StRech;
        Ferme(Q);
      end;
      ////////////////////////////////////////////////////////////////////////
      // Catégorie de taxe
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GPB_CATEGORIETAXE');
      if (StRech <> '') and (StRech <> DerniereCatTaxe) then
      begin
        SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GCX" AND CC_CODE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrCategorieTaxe
        else DerniereCatTaxe := StRech;
        Ferme(Q);
      end;
      ///////////////////////////////////////////////////////////////////////
      // Famille de taxe
      /////////////////////////////////////////////////////////////////////////
      StRech  := TPST.GetValue('GPB_CATEGORIETAXE');
      StRech2 := TPST.GetValue('GPB_FAMILLETAXE');
      if (StRech <> '') and (StRech <> DerniereFamTaxe) then
      begin
        SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE="' + StRech2 + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrFamilleTaxe
        else DerniereFamTaxe := StRech;
        Ferme(Q);
      end;
    end;
  end
  else
  begin
    // Taux et Cotation
    if sDevise <> V_PGI.DevisePivot then
    begin
      TPST.PutValue('GPB_TAUXDEV', Tob_Entete.GetValue('GP_TAUXDEV'));
      TPST.PutValue('GPB_COTATION', Tob_Entete.GetValue('GP_COTATION'));
    end;

    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB PIEDBASE
     //  2 - Rattachement à la mère PIEDBASE du document
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Base := TOB.Create('PIEDBASE', nil, -1);
    Tob_Base.Dupliquer(TPST, False, True, True);
    Tob_Base.ChangeParent(Tob_PiedBase, -1);
    Tob_Base.SetAllModifie(True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PIEDECHE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPiedEche(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Eche: TOB;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////////
      // Mode de paiement
      ////////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GPE_MODEPAIE');
      if (StRech <> '') and (StRech <> DernierModePaie) then
      begin
        SQL := 'Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrModePaie
        else DernierModePaie := StRech;
        Ferme(Q);
      end;
      ///////////////////////////////////////////////////////////////////////////
      // Devise
      ///////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GPE_DEVISE');
        if (StRech <> '') and (StRech <> DerniereDevise) then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDevise
          else DerniereDevise := StRech;
          Ferme(Q);
        end;
      end;
      ///////////////////////////////////////////////////////////////////////////
      // Devise comptoir
      ///////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GPE_DEVISEESP');
        if (StRech <> '') and (StRech <> DerniereDevise) then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDeviseEsp
          else DerniereDevise := StRech;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    // Recalcul dans la devise dossier
    ToxCalculEcheance(TPST, sDevise, sBasculeEuro, TobDev);
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB PIEDECHE
     //  3 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Eche := TOB.Create('PIEDECHE', nil, -1);
    Tob_Eche.Dupliquer(TPST, False, True, True);
    Tob_Eche.ChangeParent(Tob_PiedEche, -1);
    Tob_Eche.SetAllModifie(True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PIEDPORT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPiedPort(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Port: TOB;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////////
      // Autres contrôles quand ça marchera dans PGI !!!!!!
      ////////////////////////////////////////////////////////////////////////////

      ///////////////////////////////////////////////////////////////////////////
      // Devise
      ///////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GPT_DEVISE');
        if (StRech <> '') and (StREch <> DerniereDevise) then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDevise
          else DerniereDevise := StRech;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    /////////////////////////////////////////////////////////////////////////////////////
     //  PHASE D'INTEGRATION
     //  1 - Création de la TOB PIEDPORT
     //  2 - La MAJ DB de fera plus tard lorsque l'ensemble de la pièce sera chargé en TOB
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Port := TOB.Create('PIEDPORT', nil, -1);
    Tob_Port.Dupliquer(TPST, False, True, True);
    Tob_Port.ChangeParent(Tob_PiedPort, -1);
    Tob_Port.SetAllModifie(True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Chargement de la TOB des Nomenclatures
Suite ........ : Cette fonction reprend les sources de la fct LoadNomen
Suite ........ : (facture.pas). Cette dernière n'était pas adaptée à notre cas
Suite ........ : car elle repartait de la table LIGNENOMEN. Or, dans notre
Suite ........ : cas, tous les éléments de la pièce ont été supprimés
Mots clefs ... : TOX
*****************************************************************}
procedure ToxLoadLesNomen(TOBPiece, TobLigNomen, TobNomen, TOBArticles: TOB);
var
  i, OldL, Lig, MaxNiv, k, Niv, IndiceNomen: integer;
  TOBL, TOBLN, TOBGroupeNomen, TOBLoc: TOB;
  OkN: boolean;
begin
  //
  // Faut-il charger les nomenclatures
  //
  OkN := False;
  OldL := -1;
  IndiceNomen := 1;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if ((TOBL.GetValue('GL_TYPEARTICLE') = 'NOM') and (TOBL.GetValue('GL_TYPENOMENC') = 'ASS')) then
    begin
    	TOBL.PutValue('GL_INDICENOMEN', 0);
    	OkN := True;
    end;
  end;
  if not OkN then Exit;

  for i := 0 to TobLigNomen.Detail.Count - 1 do
  begin
    TOBLN := TobLigNomen.Detail[i];
    Lig := TOBLN.GetValue('GLN_NUMLIGNE');
    if OldL <> Lig then
    begin
      TOBGroupeNomen := TOB.Create('', TobNomen, -1);
      MaxNiv := -1;
      TOBGroupeNomen.AddChampSup('UTILISE', False);
      TOBGroupeNomen.PutValue('UTILISE', '-');
      for k := i to TobLigNomen.Detail.Count - 1 do
      begin
        TOBLoc := TobLigNomen.Detail[k];
        if TOBLoc.GetValue('GLN_NUMLIGNE') <> Lig then Break;
        Niv := TOBLoc.GetValue('GLN_NIVEAU');
        if Niv > MaxNiv then MaxNiv := Niv;
      end;
      CreerLesTOBNomen(TOBGroupeNomen, TobLigNomen, TOBArticles, Lig, MaxNiv, i);
    end;
    OldL := Lig;
  end;

  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if ((TOBL.GetValue('GL_TYPEARTICLE') = 'NOM') and (TOBL.GetValue('GL_TYPENOMENC') = 'ASS')) then
    begin
      TOBL.PutValue('GL_INDICENOMEN', IndiceNomen);
      Inc(IndiceNomen);
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Chargement de la TOB des Nomenclatures
Suite ........ : Cette fonction reprend les sources de la fct LoadNomen
Suite ........ : (facture.pas). Cette dernière n'était pas adaptée à notre cas
Suite ........ : car elle repartait de la table LIGNENOMEN. Or, dans notre
Suite ........ : cas, tous les éléments de la pièce ont été supprimés
Mots clefs ... : TOX
*****************************************************************}
procedure ToxLoadLesouvrages(TOBPiece, TobLigouv, Tobouvrage, TOBArticles: TOB);
var
  i, OldL, Lig, MaxNiv, k, Niv, IndiceNomen: integer;
  TOBL, TOBLN, TOBGroupeNomen, TOBLoc: TOB;
  OkN: boolean;
begin
{$IFDEF BTP}
  //
  // Faut-il charger les nomenclatures
  //
  OkN := False;
  OldL := -1;
  for i := 0 to TobLigOuv.Detail.Count - 1 do
  begin
    TOBLN := TobLigOuv.Detail[i];
    Lig := TOBLN.GetValue('BLO_NUMLIGNE');
    if OldL <> Lig then
    begin
      TOBGroupeNomen := TOB.Create('', Tobouvrage, -1);
      MaxNiv := -1;
      TOBGroupeNomen.AddChampSup('UTILISE', False);
      TOBGroupeNomen.PutValue('UTILISE', 'X');
      for k := i to TobLigOuv.Detail.Count - 1 do
      begin
        TOBLoc := TobLigOuv.Detail[k];
        if TOBLoc.GetValue('BLO_NUMLIGNE') <> Lig then Break;
        Niv := TOBLoc.GetValue('BLO_NIVEAU');
        if Niv > MaxNiv then MaxNiv := Niv;
      end;
      CreerLesTOBOuv(TOBGroupeNomen, TobLigOuv, TOBArticles, Lig, i);
    end;
    OldL := Lig;
  end;
{$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Annulation du stock de la pièce avant suppression du
Suite ........ : contenu
Mots clefs ... : TOX;ANNULE ET REMPLACE
*****************************************************************}
function AnnuleStockPiece(TobPiece: TOB): integer;
var
  TobEntete: TOB;
  TOBDispo: TOB;
  TobDispoArt: TOB;
  TobNomenclature: TOB;
  TOBNomen: TOB;
  TobArticle: TOB;
  TobFilleArticle: TOB;
  Condition: string;
  SQL: string;
//  Nature, Souche: string;
//  Numero, indice: integer;
  Cpt: integer;
  Q: TQUERY;
  Cledoc : R_cledoc;
//  PieceContainer: TPieceContainer;
begin
//  PieceContainer := TPieceContainer.Create;
  try

    V_PGI.IoError := oeOk;
    result := 0;
    /////////////////////////////////////////////////////////////////////////
    // Chargement de la TOB des lignes du document
    // (qui doivent être fille de l'entête)
    /////////////////////////////////////////////////////////////////////////
    cledoc.NaturePiece := TobPiece.GetValue('GP_NATUREPIECEG');
    cledoc.Souche := TobPiece.GetValue('GP_SOUCHE');
    Cledoc.NumeroPiece := TobPiece.GetValue('GP_NUMERO');
    cledoc.Indice := TobPiece.GetValue('GP_INDICEG');

    TobEntete := TOB.Create('PIECE', nil, -1);
    AddLesSupEntete(TobEntete);
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // MODIF LM : Lecture entête et non plus duplication de TOb_Piece pour Pb des pièces en annule et
    //            et remplace ayant un dépôt différent
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
    Q := OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + cledoc.NaturePiece + '" AND GP_SOUCHE="' + cledoc.Souche  + '" AND GP_NUMERO="' + IntToStr(cledoc.NumeroPiece) + '" AND GP_INDICEG="' +
      IntToStr(cledoc.Indice) + '"', False);
    TobEntete.SelectDB('', Q);
    Ferme(Q);
//    PieceContainer.TCPiece    := TobEntete;
    /////////////////////////////////////////////////////////////////////////
    // Récupère la clé du document
    /////////////////////////////////////////////////////////////////////////
//    PieceContainer.CleDoc := TOB2CleDoc(TobEntete);

    Condition := 'GL_NATUREPIECEG="' + cledoc.NaturePiece  + '" AND GL_SOUCHE="' + cledoc.Souche  + '" AND GL_NUMERO="' + IntToStr(Cledoc.NumeroPiece) + '" AND GL_INDICEG="' + IntToStr(Cledoc.Indice) + '"';
    SQL := 'SELECT * From LIGNE WHERE' + ' ' + Condition;
    Q := OpenSQL(SQL, True);
    if not Q.EOF then TobEntete.LoadDetailDB('LIGNE', '', '', Q, True, True);
    Ferme(Q);
    /////////////////////////////////////////////////////////////////////////
    // Chargement de la TOB des articles + Fiches DISPO
    /////////////////////////////////////////////////////////////////////////
    //
    // Optimisation du chargement des articles 01/04/2003
    // Chargement des articles des lignes
    //
    TobArticle := TOB.CREATE('Les Articles', nil, -1);
//    Q := OpenSQL('SELECT ARTICLE.* FROM ARTICLE, LIGNE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdLigne, False) + ' AND GL_ARTICLE=GA_ARTICLE', True);
    Q := OpenSQL('SELECT ARTICLE.* FROM ARTICLE, LIGNE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE WHERE ' + WherePiece(CleDoc, ttdLigne, False) + ' AND GL_ARTICLE=GA_ARTICLE', True);
    TobArticle.LoadDetailDB('ARTICLE', '', '', Q, False, True);
    Ferme(Q);
//    PieceContainer.TCArticles := Tob_Article;
    //
    // Chargement des fiches stocks
    //
    TobDispo := TOB.CREATE('DISPO', nil, -1);
//    Q := OpenSQL('SELECT DISPO.* FROM DISPO, LIGNE WHERE ' + WherePiece(PieceContainer.CleDoc, ttdLigne, False) +
    Q := OpenSQL('SELECT DISPO.* FROM DISPO, LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, False) +
      ' AND GQ_CLOTURE="-" AND GL_DEPOT=GQ_DEPOT AND GL_ARTICLE=GQ_ARTICLE AND GL_TYPEARTICLE="MAR"', True);
    TobDispo.LoadDetailDB('DISPO', '', '', Q, False, True);
    Ferme(Q);
    //
    //  Ajout des champs supp MODE
    //
    DispoChampSupp(TobDispo);
    //
    // Rattachement des fiches stocks aux fiches articles
    //
    for cpt := 0 to TobArticle.detail.Count - 1 do
    begin
      //
      //  Ajout des champs supp
      //
      TobFilleArticle := TobArticle.Detail[cpt];
      TobFilleArticle.AddChampSup('REFARTBARRE', False);
      TobFilleArticle.PutValue('REFARTBARRE', '');
      TobFilleArticle.AddChampSup('REFARTTIERS', False);
      TobFilleArticle.PutValue('REFARTTIERS', '');
      TobFilleArticle.AddChampSup('REFARTSAISIE', False);
      TobFilleArticle.PutValue('REFARTSAISIE', '');
      TobFilleArticle.AddChampSup('UTILISE', False);
      TobFilleArticle.PutValue('UTILISE', '-');

      TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'], [TobFilleArticle.GetValue('GA_ARTICLE')], False);
      if TobDispoArt <> nil then
      begin
        TobDispoArt.Changeparent(TobFilleArticle, -1);
      end;
    end;
    TobDispo.free;

    /////////////////////////////////////////////////////////////////////////
    // Chargement des lignes nomenclatures
    ////////////////////////////////////////////////////////////////////////
    TOBNomen := TOB.Create('', nil, -1);
    Condition := 'GLN_NATUREPIECEG="' + cledoc.NaturePiece  + '" AND GLN_SOUCHE="' + Cledoc.Souche  + '" AND GLN_NUMERO="' + IntToStr(Cledoc.NumeroPiece ) + '" AND GLN_INDICEG="' + IntToStr(Cledoc.Indice) + '"';
    SQL := 'SELECT * From LIGNENOMEN WHERE ' + Condition;
    Q := OpenSQL(SQL, True);
    if not Q.EOF then TOBNomen.LoadDetailDB('LIGNENOMEN', '', '', Q, True, False);
    Ferme(Q);
    /////////////////////////////////////////////////////////////////////////
    // Chargement des nomenclatures
    /////////////////////////////////////////////////////////////////////////
    TobNomenclature := TOB.CREATE('Les Nomens', nil, -1);
    ToxLoadLesNomen(TobEntete, TOBNomen, TobNomenclature, TobArticle);
    if V_PGI.IoError <> oeOK then result := ErrLoadLesNomens;
    //////////////////////////////////////////////////////////////////////////////
    // MAJ du stock
    // Fiches stocks "DISPO" : chargées en TOB Fille de TobFilleArticle
    //   TobArticle -> TobFilleArticle -> TobDispo
    // Pour mettre à jour les stocks, il faut faire un updateDB de TobArticle
    //
    // Attention : Modif 06/08/2003
    //             Nouvelle Gestion des Reliquats !!!!!!
    //             On ré-initialise la quantité GL_QTESTOCK par GL_QTERESTE
    //////////////////////////////////////////////////////////////////////////////
    for Cpt := 0 to TobEntete.Detail.Count - 1 do
    begin
      TobEntete.detail[Cpt].PutValue('GL_QTESTOCK', TobEntete.detail[Cpt].GetValue('GL_QTERESTE'));
      TobEntete.detail[Cpt].PutValue('GL_QTEFACT', TobEntete.detail[Cpt].GetValue('GL_QTERESTE'));
    end;
    ////// FIN des modifications pour nouvelle gestion des reliquats /////////////

//    if result = 0 then InverseStock(PieceContainer);
    if result = 0 then InverseStock(TobEntete,TobArticle,nil,nil);
    if V_PGI.IoError <> oeOK then result := ErrCalculInverseDispo;
    //////////////////////////////////////////////////////////////////////////////
    // MAJ du stock (updateDB de dispo)
    //////////////////////////////////////////////////////////////////////////////
//    if result = 0 then ValideLesArticles(PieceContainer, TobEntete);
    if result = 0 then ValideLesArticles(TOBentete,TOBarticle);
    if V_PGI.IoError <> oeOK then result := ErrMAJInverseDispo;
    //////////////////////////////////////////////////////////////////////////////
    // Libération des TOB
    //////////////////////////////////////////////////////////////////////////////
    TobEntete.free;
    TobArticle.free;
    TobNomenclature.free;
    TOBNomen.free;
  finally
//    PieceContainer.Free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Suppression des élkéments d'une pièce
Mots clefs ... : TOX;SUPPRESSION PIECE
*****************************************************************}
function SupprimeElementsPiece(Nature, Souche: string; Numero, indice: integer): boolean;
var
  Condition: string;
  SQL: string;
  IdentifiantBlob : string;
begin
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des LIGNES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GL_NATUREPIECEG="' + Nature + '" AND GL_SOUCHE="' + Souche + '" AND GL_NUMERO="' + IntToStr(Numero) + '" AND GL_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNE WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des LIGNES NOMENCLATURES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GLN_NATUREPIECEG="' + Nature + '" AND GLN_SOUCHE="' + Souche + '" AND GLN_NUMERO="' + IntToStr(Numero) + '" AND GLN_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNENOMEN WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des LIGNES LOTS
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GLL_NATUREPIECEG="' + Nature + '" AND GLL_SOUCHE="' + Souche + '" AND GLL_NUMERO="' + IntToStr(Numero) + '" AND GLL_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNELOT WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession dans PIEDBASE
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GPB_NATUREPIECEG="' + Nature + '" AND GPB_SOUCHE="' + Souche + '" AND GPB_NUMERO="' + IntToStr(Numero) + '" AND GPB_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From PIEDBASE WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des PIEDECHE
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GPE_NATUREPIECEG="' + Nature + '" AND GPE_SOUCHE="' + Souche + '" AND GPE_NUMERO="' + IntToStr(Numero) + '" AND GPE_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From PIEDECHE WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des PIEDPORT
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GPT_NATUREPIECEG="' + Nature + '" AND GPT_SOUCHE="' + Souche + '" AND GPT_NUMERO="' + IntToStr(Numero) + '" AND GPT_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From PIEDPORT WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des ACOMPTES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GAC_NATUREPIECEG="' + Nature + '" AND GAC_SOUCHE="' + Souche + '" AND GAC_NUMERO="' + IntToStr(Numero) + '" AND GAC_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From ACOMPTES WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
{$IFDEF BTP}
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des Bloc notes entetes et pied
  ////////////////////////////////////////////////////////////////////////////
  IdentifiantBlob := Nature+':'+Souche+':'+IntToStr(Numero)+':'+IntToStr(Indice);
  Condition := 'LO_TABLEBLOB="GP" AND LO_QUALIFIANTBLOB="MEM" AND LO_IDENTIFIANT="' + IdentifiantBlob + '" ';
  SQL := 'Delete From LIENSOLE WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des LIGNES COMPLEMENTAIRES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'GLC_NATUREPIECEG="' + Nature + '" AND GLC_SOUCHE="' + Souche + '" AND GLC_NUMERO="' + IntToStr(Numero) + '" AND GLC_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNECOMPL WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des LIGNES D'OUVRAGES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'BLO_NATUREPIECEG="' + Nature + '" AND BLO_SOUCHE="' + Souche + '" AND BLO_NUMERO="' + IntToStr(Numero) + '" AND BLO_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNEOUV WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des REPARTITIONS EN MILLIEMES DES TVA
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'BPM_NATUREPIECEG="' + Nature + '" AND BPM_SOUCHE="' + Souche + '" AND BPM_NUMERO="' + IntToStr(Numero) + '" AND BPM_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From BTPIECEMILIEME WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des VARIABLES DOCUMENTS
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'BVD_NATUREPIECE="' + Nature + '" AND BVD_SOUCHE="' + Souche + '" AND BVD_NUMERO="' + IntToStr(Numero) + '" AND BVD_INDICE="' + IntToStr(Indice) + '"';
  SQL := 'Delete From BVARDOC WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des PARAMETRES D'EDITIONS DOCUMENT
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'BPD_NATUREPIECE="' + Nature + '" AND BPD_SOUCHE="' + Souche + '" AND BPD_NUMPIECE="' + IntToStr(Numero) + '"';
  SQL := 'Delete From BTPARDOC WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppression des LIENS AVEC LES PHASES
  ////////////////////////////////////////////////////////////////////////////
  Condition := 'BLP_NATUREPIECEG="' + Nature + '" AND BLP_SOUCHE="' + Souche + '" AND BLP_NUMERO="' + IntToStr(Numero) + '" AND BLP_INDICEG="' + IntToStr(Indice) + '"';
  SQL := 'Delete From LIGNEPHASES WHERE' + ' ' + Condition;
  ExecuteSQL(SQL);
{$ENDIF}
  result := True;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Chargement de la TOB des lots
Suite ........ : Cette fonction reprend les sources de la fct LoadLeslots
Suite ........ : (factutil.pas).  Cette dernière n'était pas adaptée à notre cas
Suite ........ : car elle repartait de la table LIGNELOT. Or, dans notre cas,
Suite ........ : tous les éléments de la pièce ont été supprimés
Mots clefs ... : TOX
*****************************************************************}
procedure ToxLoadLesLots(TOBPiece, TobLigLot, TOBDesLots: TOB);
{$IFNDEF STK}
var
  i, k, NumL, Indicelot: integer;
  CodeLot: string;
  TOBLigne, TOBLoc, TOBNoeud: TOB;
{$ENDIF STK}
begin
{$IFNDEF STK}
  for i := TobLigLot.Detail.Count - 1 downto 0 do
  begin
    {Lecture des LigneLot "en vrac"}
    TOBLoc := TobLigLot.Detail[i];
    NumL := TOBLoc.GetValue('GLL_NUMLIGNE');
    CodeLot := TOBLoc.GetValue('GLL_NUMEROLOT');
    TOBLigne := TOBPiece.FindFirst(['GL_NUMLIGNE'], [NumL], False);
    {recherche de la ligne de pièce concernée}
    if TOBLigne <> nil then
    begin
      {Recherche du noeud des lots de la ligne, si pas ok le créer}
      TOBNoeud := nil;
      IndiceLot := 0;
      for k := 0 to TOBDesLots.Detail.Count - 1 do
      begin
        if TOBDesLots.Detail[k].GetValue('NUMLIGNE') = NumL then
        begin
          TOBNoeud := TOBDeslots.Detail[k];
          IndiceLot := k + 1;
          Break;
        end;
      end;
      if TOBNoeud = nil then
      begin
        TOBNoeud := CreerNoeudLot(TOBPiece, TOBDesLots);
        TOBNoeud.PutValue('NUMLIGNE', NumL);
        IndiceLot := TOBDesLots.Detail.Count;
      end;
      TOBLigne.PutValue('GL_INDICELOT', IndiceLot);
      TOBLoc.ChangeParent(TOBNoeud, -1);                                              
    end;
  end;
{$ENDIF STK}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour des stocks
Mots clefs ... : TOX;IMPORT PIECE
*****************************************************************}
procedure MetAJourLesStocks(var TobEntete, TobArticle, TobCatalogue, TobNomenclature: TOB; MajStock: boolean);
var
  TobEnt: TOB;
//  Retour: integer;
  Cpt: integer;
//  PieceContainer: TPieceContainer;
begin
//  PieceContainer := TPieceContainer.Create;
  try
    //
    // Duplication de la TOB PIECE pour nepas l'endommager
    //
    TobEnt := TOB.Create('PIECE', nil, -1);
    TobEnt.Dupliquer(TobEntete, True, True, True);
    //
    // Initialisation de GL_QTESTOCK par GL_QTERESTE
    //
    for Cpt := 0 to TobEnt.Detail.Count - 1 do
    begin
      TobEnt.detail[Cpt].PutValue('GL_QTESTOCK', TobEnt.detail[Cpt].GetValue('GL_QTERESTE'));
      TobEnt.detail[Cpt].PutValue('GL_QTEFACT', TobEnt.detail[Cpt].GetValue('GL_QTERESTE'));
    end;
    //
    // Maj des stocks
    //
//    PieceContainer.TCPiece    := TobEnt;
//    PieceContainer.TCArticles := TobArticle;
//    ValideLesLignes(PieceContainer, False, False);
    ValideLesLignes(TOBEnt, TOBArticle, nil,nil,nil,nil,nil,true,false);
//    if MajStock then ValideLesArticles(PieceContainer, TobEnt);
//    if MajStock then ValideLesArticles(PieceContainer, TobEnt);
    if MajStock then ValideLesArticles(Tobent,TObarticle);
    //
    // Libération de la TOB
    //
    TobEnt.free;
  finally
//    PieceContainer.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Importation d'un document avec MAJ stock ...
Mots clefs ... : TOX;IMPORT PIECE
*****************************************************************}
function ImportDocument(MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  TobDesLots: TOB;
  NowFutur: TDateTime;
  i: integer;
  SumQteReste: integer;
  SumQteStock: integer;
  QteReste: integer;
  QteStock: integer;
begin
  try
    TobDesLots := nil;
    if Tob_Entete = nil then
    begin
      result := 0;
      exit;
    end;
    /////////////////////////////////////////////////////////////////////////
    // Nouvelle gestion des reliquats
    // Assure la compatibilité entre versions
    /////////////////////////////////////////////////////////////////////////
    SumQteReste := 0;
    SumQteStock := 0;
    for i := 0 to Tob_Entete.Detail.Count - 1 do
    begin
      ///////////////////////////////////////
      // Resynchro GL_VIVANTE / GP_VIVANTE
      ///////////////////////////////////////
      if Tob_Entete.detail[i].GetValue('GL_VIVANTE') <> Tob_Entete.GetValue('GP_VIVANTE') then
      begin
        Tob_Entete.detail[i].PutValue('GL_VIVANTE', Tob_Entete.GetValue('GP_VIVANTE'));
      end;
      //////////////////////////////////////////////////////////////
      // Transformation Ancienne Gestion des reliquats en Nouvelle
      //////////////////////////////////////////////////////////////
      if Tob_Entete.detail[i].GetValue('GL_TYPELIGNE') = 'ART' then
      begin
        QteReste := Tob_Entete.detail[i].GetValue('GL_QTERESTE');
        QteStock := Tob_Entete.detail[i].GetValue('GL_QTESTOCK');
        SumQteReste := sumQteReste + Abs(QteReste);
        SumQteStock := sumQteStock + Abs(QteStock);
      end;
    end;
    /////////////////////////////////////////////////////////////////////////////////////////
    // Cas 1 : une ancienne pièce remonte avec des reliquats, elle est vivante
    //         Sur ancienne version, QTERESTE est à 0, QTESTOCK correspond au reliquat
    //         Sur une nouvelle pièce, il faut initialiser QTERESTE avec QTESTOCK
    /////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Entete.GetValue('GP_VIVANTE') = 'X') and (SumQteReste = 0) and (SumQteStock <> 0) then
    begin
      for i := 0 to Tob_Entete.Detail.Count - 1 do
      begin
        Tob_Entete.detail[i].PutValue('GL_QTERESTE', Tob_Entete.detail[i].GetValue('GL_QTESTOCK'));
      end;
    end;
    ///////////////////////////////////////////////////////////////////////////////////////
    //  Cas 2 : si annonce de transfert ou annonce de livraison non vivante, alors
    //          on initialise les quantités restes à 0 pour ne plus avoir de reliquat
    //  PS : On se limite à ces 2 natures de pièce pour limiter les effets de bords
    //  ex : les tickets FFO sont non vivants, il faut conserver les valeurs de QL_QTERESTE
    ////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'ALF')
       or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'TRV')
       or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'CF')then
    begin
      if (Tob_Entete.GetValue('GP_VIVANTE') = '-') then
      begin
        for i := 0 to Tob_Entete.Detail.Count - 1 do
        begin
          Tob_Entete.detail[i].PutValue('GL_QTERESTE', 0);
        end;
      end;
    end;
    ///////////////////////////////////////////////////////////////////////////////////////
    //  Cas 3 : si réception fournisseur vivante, GL_QTESTOCK = GK_QTERESTE
    ////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'BLF')
        or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'TRE')
        or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'TEM')then
    begin
      if (Tob_Entete.GetValue('GP_VIVANTE') = 'X') then
      begin
        for i := 0 to Tob_Entete.Detail.Count - 1 do
        begin
          Tob_Entete.detail[i].PutValue('GL_QTERESTE', Tob_Entete.detail[i].GetValue('GL_QTESTOCK'));
        end;
      end;
    end;
    ///////////////////////////////////////////////////////////////////////////////////////
    //  Cas 4 : Les tickets et inventaires 4.2
    ////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'FFO')
        or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'INV') then
    begin
      for i := 0 to Tob_Entete.Detail.Count - 1 do
      begin
        Tob_Entete.detail[i].PutValue('GL_QTERESTE', Tob_Entete.detail[i].GetValue('GL_QTESTOCK'));
      end;
    end;


    V_PGI.IoError := oeOk;
    result := IntegOK;
    /////////////////////////////////////////////////////////////////////////
    // Lecture des nomenclatures
    /////////////////////////////////////////////////////////////////////////
    ToxLoadLesNomen(Tob_Entete, Tob_LigneNom, Tob_Nomenclature, Tob_Article);
    if V_PGI.IoError <> oeOK then result := ErrLoadLesNomens;
{$IFDEF BTP}
    /////////////////////////////////////////////////////////////////////////
    // lesture des  Ouvrages
    ////////////////////////////////////////////////////////////////////////
    ToxLoadLesouvrages(Tob_Entete, Tob_ligneouv, Tob_Ouvrages, Tob_Article);
    if V_PGI.IoError <> oeOK then result := ErrLoadLesOuvrages;
{$ENDIF}
    /////////////////////////////////////////////////////////////////////////
    // Lecture des lots
    /////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      TobDesLots := TOB.Create('LES LOTS', nil, -1);
      ToxLoadLesLots(Tob_Entete, Tob_LigneLot, TobDesLots);
      if V_PGI.IoError <> oeOK then result := ErrLoadLesLots;
    end;
    /////////////////////////////////////////////////////////////////////////
    // MAJ des "Top Modif" des TOB
    /////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      NowFutur := NowH;
      Tob_Entete.SetAllModifie(True);

      Tob_Entete.SetDateModif(Tob_Entete.GetValue('GP_DATEMODIF'));

      Tob_Entete.PutValue('GP_DATEINTEGR', NowFutur);
      Tob_LigneLot.SetAllModifie(True);
      Tob_LigneNom.SetAllModifie(True);
{$IFDEF BTP}
      Tob_LigneOuv.SetAllModifie(True);
      Tob_Millieme.SetAllModifie (True);
      Tob_PiedBaseRg.SetAllModifie(True);
{$ENDIF}
      Tob_PiedBase.SetAllModifie(True);
      Tob_PiedEche.SetAllModifie(True);
      Tob_Acomptes.SetAllModifie(True);
      Tob_PiedPort.SetAllModifie(True);
      Tob_Tiers.SetAllModifie(True);
      //TOBAnaP.SetAllModifie(True)    ;
    end;

    //////////////////////////////////////////////////////////////////////////////
    // Mise à jour de la pièce dans la devise du dossier
    //////////////////////////////////////////////////////////////////////////////
    ToxCalculFacture(sDevise, sbasculeEuro);
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Préparation des MAJ de stock
    //             -> Recalcul les quantités des fiches DISPO, TOB filles de Tob_Article
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then MetAJourLesStocks(Tob_Entete, Tob_Article, Tob_Catalogue, Tob_LigneNom, MajStock);
    //////////////////////////////////////////////////////////////////////////////////////////////
    // MAJ de la TOB Tiers
    //             -> MAJ date dernière pièce, dernier n° dicument et dernier CA
    //             -> UpdateDB de Tob_Tiers
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      // MODIF LM le 16/01/2002
      // On se passe de mettre à jour le Tiers :
      //    1 - Optimisation de l'intégration
      //    2 - On ne change pas la date de modification du Tiers, ce qui évite de redescendre
      //        le tiers dans toutes les boutiques

      //if result = 0 then ValideLeTiers (PieceContainer);
      //if V_PGI.IoError<>oeOK then result := ErrMAJTiersPiece;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des acomptes
    //            -> Reforce la clé des enregistrements Acomptes
    //            -> InsertDB de Tob_Acomptes
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideLesAcomptes(TOB_Entete,Tob_Acomptes);
      if V_PGI.IoError <> oeOK then result := ErrMAJAcomptes;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des lignes de ports
    //            -> Reforce la clé des enregistrements Ports
    //            -> Fait un InsertDB de Tob_PiedPort
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideLesPorcs(Tob_Entete,Tob_PiedPort);
      if V_PGI.IoError <> oeOK then result := ErrMAJPorts;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des Nomenclatures
    //             -> MAJ du flag UTILISE des nomenclatures
    //             -> Update de TOBNomenclature
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideLesNomen( Tob_LigneNom);
      if V_PGI.IoError <> oeOK then result := ErrMAJNomencl;
    end;
{$IFDEF BTP}
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des Ouvrages
    //             -> MAJ du flag UTILISE des nomenclatures
    //             -> Update de TOBOuvrage
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
//      if result = 0 then ValideLesOuv ( Tob_Ligneouv,Tob_entete);
			if result = 0 then TOB_LigneOuv.InsertDB (nil);
      if V_PGI.IoError <> oeOK then result := ErrMAJOuvrage;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des TOB TVAMILLIME
    //             -> MAJ du flag UTILISE des nomenclatures
    //             -> Update de TOBMilliemes
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideMillieme ( Tob_Millieme);
      if V_PGI.IoError <> oeOK then result := ErrMillieme;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des TOB RETENUE DE GARANTIE
    //             -> MAJ du flag UTILISE des nomenclatures
    //             -> Update de TOBPieceRg
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideLesRetenues ( Tob_entete,Tob_PieceRg);
      if V_PGI.IoError <> oeOK then result := ErrMAJPieceRG;
    end;
    //////////////////////////////////////////////////////////////////////////////////////////////
    // Validation des TOB Base TVA sur RETENUE DE GARANTIE
    //             -> MAJ du flag UTILISE des nomenclatures
    //             -> Update de TOBPiedBaseRg
    //////////////////////////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      if result = 0 then ValideLesBasesRG ( Tob_entete,Tob_PiedBaseRg);
      if V_PGI.IoError <> oeOK then result := ErrMAJPiedPieceRG;
    end;
{$ENDIF}
    /////////////////////////////////////////////////////////////////////////
    // Importation DB
    //////////////////////////////////////////////////////////////////////////
    if (result = IntegOK) then
    begin
      Tob_Entete.InsertOrUpdateDB(False);
      Tob_PiedEche.InsertDB(nil);
      Tob_PiedBase.InsertDB(nil);
    end;
    //////////////////////////////////////////////////////////////////////////
    // Libération des TOB
    //////////////////////////////////////////////////////////////////////////
    if Tob_Entete <> nil then FreeAndNil(Tob_Entete);
    if Tob_Article <> nil then FreeAndNil(Tob_Article);
    if Tob_PiedBase <> nil then FreeAndNil(Tob_PiedBase);
    if Tob_PiedEche <> nil then FreeAndNil(Tob_PiedEche);
    if Tob_PiedPort <> nil then FreeAndNil(Tob_PiedPort);
    if Tob_Acomptes <> nil then FreeAndNil(Tob_Acomptes);
    if Tob_LigneNom <> nil then FreeAndNil(Tob_LigneNom);
    if Tob_LigneLot <> nil then FreeAndNil(Tob_LigneLot);
    if TobDesLots <> nil then FreeAndNil(TobDesLots);
    if Tob_Catalogue <> nil then FreeAndNil(Tob_Catalogue);
    if Tob_Nomenclature <> nil then FreeAndNil(Tob_Nomenclature);
{$IFDEF BTP}
  	if Tob_LigneOuv <> nil then FreeAndNil (Tob_LigneOuv);
    if Tob_Millieme <> nil then FreeAndNil (Tob_Millieme);
    if Tob_PieceRg  <> nil then FreeAndNil (Tob_PieceRg);
    if Tob_PiedBaseRg <> nil then FreeAndNil (Tob_PiedBaseRg);
    if Tob_Ouvrages <> nil then FreeAndNil (Tob_Ouvrages);
{$ENDIF}

  finally
  //
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Annule une pièce existante avec Maj du stock
Mots clefs ... : TOX
*****************************************************************}
procedure AnnuleLaPieceEtLeStock(TobEntete: TOB; MajStock: boolean);
var
  Nature: string;
  Souche: string;
  Numero: integer;
  Indice: integer;
  Retour: integer;
begin
  ////////////////////////////////////////////////////
  // MODIF LM 27/08/2003
  // Annulation de la pièce via une transaction
  ////////////////////////////////////////////////////
  try
    BeginTrans;

    Nature := TobEntete.GetValue('GP_NATUREPIECEG');
    Souche := TobEntete.GetValue('GP_SOUCHE');
    Numero := TobEntete.GetValue('GP_NUMERO');
    Indice := TobEntete.GetValue('GP_INDICEG');

    if MajStock then Retour := AnnuleStockPiece(TobEntete) else Retour := 0;
    if Retour = 0 then SupprimeElementsPiece(Nature, Souche, Numero, Indice);

    CommitTrans ;
  except
    RollBack;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles TOB PIECE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPIECE(TPST: TOB; Integre, IntegrationRapide, MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  CodeTiers: string;
  Condition: string;
  Nature: string;
  Souche: string;
  Numero: integer;
  Indice: integer;
  SQL: string;
  StRech: string;
  PieceVivante : string ;
  Q: TQUERY;
  DEV: RDEVISE;
  DatePiece: TDateTime;
//  io: TIOErr;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////////
      // Nature de document
      /////////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GP_NATUREPIECEG');
      if StRech <> DerniereNature then
      begin
        SQL := 'Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then Result := ErrNaturePiece
        else DerniereNature := StRech;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Contrôles des Tiers
      /////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_TIERS');
        if Trim(StRech) = '' then Result := ErrTiersVide;
      end;
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_TIERS');
        if (StRech <> '') and (StRech <> DernierTiers) then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiers
          else DernierTiers := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Possibilité de remplacer le tiers (Ex : cas des franchises)
      /////////////////////////////////////////////////////////////////////////////
      if ((Result = ErrTiersVide) or (Result = ErrTiers)) then
      begin
        if GetParamSoc('SO_GCTOXFORCETIERS') then
        begin
          Result := IntegOK;
          CodeTiers := GetParamSoc('SO_GCTOXTIERS');
          if Trim(CodeTiers) = '' then Result := ErrTiersVide
          else
          begin
            Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeTiers + '"', False);
            if Q.EOF then Result := ErrTiers;
            Ferme(Q);
          end;
          if (result = IntegOk) then
          begin
            TPST.PutValue('GP_TIERS', CodeTiers);
            TPST.PutValue('GP_TIERSLIVRE', CodeTiers);
            TPST.PutValue('GP_TIERSFACTURE', CodeTiers);
            TPST.PutValue('GP_TIERSPAYEUR', CodeTiers);
          end;
        end;
      end;

      if (result = IntegOK) and (TPST.GetValue('GP_TIERSLIVRE') <> TPST.GetValue('GP_TIERS')) then
      begin
        StRech := TPST.GetValue('GP_TIERSLIVRE');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiersLivre;
          Ferme(Q);
        end;
      end;
      if (result = IntegOK) and (TPST.GetValue('GP_TIERSFACTURE') <> TPST.GetValue('GP_TIERS')) then
      begin
        StRech := TPST.GetValue('GP_TIERSFACTURE');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiersFacture;
          Ferme(Q);
        end;
      end;
      if (result = IntegOK) and (TPST.GetValue('GP_TIERSPAYEUR') <> TPST.GetValue('GP_TIERS')) then
      begin
        StRech := TPST.GetValue('GP_TIERSPAYEUR');
        if StRech <> '' then
        begin
          SQL := 'Select T_LIBELLE From TIERS WHERE T_TIERS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrTiersPayeur;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////////
      // Représentant
      /////////////////////////////////////////////////////////////////////////////
      //if (result=IntegOK) then
      //begin
      //  StRech     := TPST.GetValue ('GP_REPRESENTANT');
      //  if StRech <> '' then
      //  begin
      //    SQL:='Select GCL_LIBELLE From COMMERCIAL WHERE ((GCL_TYPECOMMERCIAL="REP") OR (GCL_TYPECOMMERCIAL="VEN")) AND GCL_COMMERCIAL="'+StRech+'"';
      //    Q:=OpenSQL(SQL,True) ;
      //    if Q.EOF then Result:=ErrRepresentant;
      //    Ferme(Q);
      //  end;
      //end;
      ////////////////////////////////////////////////////////////////////////////
      // Etablissement
      ////////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_ETABLISSEMENT');
        if (StRech <> '') and (StRech <> DernierEtab) then
        begin
          SQL := 'Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrEtablissement
          else DernierEtab := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Code dépôt
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_DEPOT');
        if (StRech <> '') and (StRech <> DernierDepot) then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepot
          else DernierDepot := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Code dépôt destinataire (transferts inter-boutiques)
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_DEPOTDEST');
        if (StRech <> '') and (StRech <> DernierDepot) then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepotDest
          else DernierDepot := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Mode de règlement
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_MODEREGLE');
        if (StRech <> '') and (StRech <> DernierModeRegl) then
        begin
          SQL := 'Select MR_LIBELLE From MODEREGL WHERE MR_MODEREGLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrModeRegle
          else DernierModeRegl := StRech;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Régime TVA
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_REGIMETAXE');
        if (StRech <> '') and (StRech <> DernierRegimeTaxe) then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrRegimeTVA
          else DernierRegimeTaxe := StRech;
          Ferme(Q);
        end;
      end;
      ///////////////////////////////////////////////////////////////////////////
      // Devise
      ///////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('GP_DEVISE');
        if (StRech <> '') and (StRech <> DerniereDevise) then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDevise
          else DerniereDevise := StRech;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin
    ////////////////////////////////////////////////////////////////
    // PHASE D'INTEGRATION
    //       1 - Initialisation des TOB de la pièce à nil avant tout traitement
    //       2 - Création de la TOB Tob_Ent
    //       3 - Création des TOB de la pièce : articles, pied .....
    //       4 - Suppression des éléments du document (lignes+pieds+acomptes)
    //           si déjà existant
   ////////////////////////////////////////////////////////////////
    Tob_Entete := nil;
    Tob_Ligne := nil;
    Tob_LigneNom := nil;
    Tob_LigneLot := nil;
    Tob_PiedBase := nil;
    Tob_PiedEche := nil;
    Tob_PiedPort := nil;
    Tob_Acomptes := nil;
    Tob_Tiers := nil;
    Tob_Article := nil;
    Tob_Catalogue := nil;
    Tob_Nomenclature := nil;
{$IFDEF BTP}
    // BTP
    Tob_LigneOuv:= nil;
    Tob_Millieme:= nil;
    Tob_PieceRg:= nil;
    Tob_PiedBaseRg:= nil;
    Tob_Ouvrages := nil;
    //
{$ENDIF}
    ////////////////////////////////////////////////////
    // Création TOB entête
    ////////////////////////////////////////////////////
    Tob_Entete := TOB.Create('PIECE', nil, -1);
    Tob_Entete.Dupliquer(TPST, False, True, True);
    AddLesSupEntete(Tob_Entete);
    //
    // Recalcul des taux entre la devise du document et la devise pivot
    //
    if sDevise <> V_PGI.DevisePivot then
    begin
      DatePiece := Tob_Entete.GetValue('GP_DATEPIECE');
      DEV.Code := Tob_Entete.GetValue('GP_DEVISE');
      GetInfosDevise(DEV);
      DEV.Taux := GetTaux(DEV.Code, DEV.DateTaux, DatePiece);
      Tob_Entete.PutValue('GP_TAUXDEV', DEV.Taux);
      AttribCotation(Tob_Entete);
    end;

    ////////////////////////////////////////////////////
    // Création et chargement TOB TIERS
    ////////////////////////////////////////////////////
    CodeTiers := Tob_Entete.GetValue('GP_TIERS');
    SQL := 'Select * From TIERS WHERE T_TIERS="' + CodeTiers + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.Create('TIERS', nil, -1);
      Tob_Tiers.SelectDB('', Q);
    end;
    Ferme(Q);
    /////////////////////////////////////////////////////
    // Création des TOB Virtuelles rattachées à la pièce
    /////////////////////////////////////////////////////
    Tob_LigneNom := TOB.CREATE('Lignes Nomenclatures', nil, -1);
    Tob_LigneLot := TOB.CREATE('Lignes Lots', nil, -1);
    Tob_Article := TOB.CREATE('Les Articles', nil, -1);
    Tob_PiedBase := TOB.CREATE('Les Taxes', nil, -1);
    Tob_PiedEche := TOB.CREATE('Les Echéances', nil, -1);
    Tob_PiedPort := TOB.CREATE('Les Ports', nil, -1);
    Tob_Acomptes := TOB.CREATE('Les Acomptes', nil, -1);
    Tob_Catalogue := TOB.CREATE('Les Catalogues', nil, -1);
    Tob_nomenclature := TOB.CREATE('Les Nomenclatures', nil, -1);
{$IFDEF BTP}
    // BTP
    Tob_LigneOuv:= TOB.Create ('LEs Lignes Ouvrages',nil,-1);
    Tob_Millieme:= TOB.Create ('LES MILLIEMES',nil,-1);
    Tob_PieceRg:= TOB.Create ('LES RG',nil,-1);
    Tob_PiedBaseRg:= TOB.Create ('LES BASES RG',nil,-1);
    Tob_Ouvrages := TOB.Create ('LES OUVRAGES',nil,-1);
{$ENDIF}
    //
    /////////////////////////////////////////////////////////////////
    // PHASE D'INTEGRATION : Les pièces sont intégrées en "annule et
    //                       remplace". Lors de l'intégration, la ligne
    //                       ne doit pas existée
    //////////////////////////////////////////////////////////////////
    Nature := Tob_Entete.GetValue('GP_NATUREPIECEG');
    Souche := Tob_Entete.GetValue('GP_SOUCHE');
    Numero := Tob_Entete.GetValue('GP_NUMERO');
    Indice := Tob_Entete.GetValue('GP_INDICEG');

    Condition := 'GP_NATUREPIECEG="' + Nature + '" AND GP_SOUCHE="' + Souche + '" AND GP_NUMERO="' + IntToStr(Numero) + '" AND GP_INDICEG="' + IntToStr(Indice) + '"';
    SQL := 'Select GP_VIVANTE From PIECE WHERE' + ' ' + Condition;
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      ////////////////////////////////////////////////////////////////////
      // MODIF LM 27/08/2003
      // Si la pièce existe et est non vivante, on conserve son état de
      // non modifiable
      // PS : on se limite aux annonces de livraisons et annonces de
      //      de transferts pour limiter les effets de bords
      ////////////////////////////////////////////////////////////////////
      if (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'ALF') or (Tob_Entete.GetValue ('GP_NATUREPIECEG') = 'TRV') then
      begin
        PieceVivante := Q.FindField('GP_VIVANTE').AsString ;
        if (PieceVivante = '-') and (Tob_Entete.GetValue ('GP_VIVANTE') = 'X') then Tob_Entete.PutValue ('GP_VIVANTE', '-');
      end;
      ////////////////////////////////////////////////////////////////////
      // La pièce existe déjà
      //    1 - Mise à jour du stock/compta pour annulation de la pièce
      //    2 - Suppression de son contenu
      //
      // MODIF LM 23/07/02
      // Mise à jour des stocks facultative
      //
      ////////////////////////////////////////////////////////////////////
      // io:=Transactions (AnnuleLaPieceEtLeStock, 0) ;
      // if io<>oeOk then Result:=ErrAnnulePiece;
      AnnuleLaPieceEtLeStock(Tob_Entete, MajStock);
    end;
    Ferme(Q);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PIECEADRESSES
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPieceAdresse(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_PieceAdresse: TOB;
  Nature: string;
  Souche: string;
  Numero: integer;
  Indice: integer;
  Numligne: integer;
  TypePieceAdr: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Pas de contrôle, la pièce est peut être en cours d'import et n'existe pas encore
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_PieceAdresse := TOB.Create('PIECEADRESSE', nil, -1);
    Tob_PieceAdresse.Dupliquer(TPST, False, True, True);
    Tob_PieceAdresse.SetAllModifie(True);

    Nature := Tob_PieceAdresse.GetValue('GPA_NATUREPIECEG');
    Souche := Tob_PieceAdresse.GetValue('GPA_SOUCHE');
    Numero := Tob_PieceAdresse.GetValue('GPA_NUMERO');
    Indice := Tob_PieceAdresse.GetValue('GPA_INDICEG');
    Numligne := Tob_PieceAdresse.GetValue('GPA_NUMLIGNE');
    TypePieceAdr := Tob_PieceAdresse.GetValue('GPA_TYPEPIECEADR');

    SQL := 'Select GPA_NATUREPIECEG From PIECEADRESSE WHERE GPA_NATUREPIECEG="' + Nature + '" AND GPA_SOUCHE="' + Souche + '" AND GPA_NUMERO="' + IntToStr(Numero) +
      '" AND GPA_INDICEG="' + IntToStr(Indice) + '" AND GPA_NUMLIGNE="' + IntToStr(Numligne) + '" AND GPA_TYPEPIECEADR="' + TypePieceAdr + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_PieceAdresse.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_PieceAdresse.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_PieceAdresse.free;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB PORT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBPort(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  SQL, StRech, CodePort: string;
  Tob_Port: TOB;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Devise
      StRech := TPST.GetValue('GPO_TYPEPORT');
      if (StRech <> '') then
      begin
        SQL := 'Select CO_CODE From COMMUN WHERE CO_TYPE="GTP" and CO_CODE="' + StRech + '"';
        if not ExisteSQL(SQL) then result := ErrTypePort;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    //
    // Création d'une TOB contenant  l'enregistrement courant
    //
    Tob_Port := TOB.Create('PORT', nil, -1);
    Tob_Port.Dupliquer(TPST, False, True, True);
    /////////////////////////////////////////////////////////////////
    // Conversion dans la devise dossier
    /////////////////////////////////////////////////////////////////
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Montant HT
      Montant := Tob_Port.GetValue('GPO_PVHT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Port.PutValue('GPO_PVHT', MontantDev);
      // Montant TTC
      Montant := Tob_Port.GetValue('GPO_PVTTC');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Port.PutValue('GPO_PVTTC', MontantDev);
      // Montant minimum pour franco
      Montant := Tob_Port.GetValue('GPO_MINIMUM');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Port.PutValue('GPO_MINIMUM', MontantDev);
    end;
    //
    // Mise à jour de tous les champs
    //
    Tob_Port.SetAllModifie(True);

    CodePort := Tob_Port.GetValue('GPO_CODEPORT');

    SQL := 'Select GPO_CODEPORT From PORT WHERE GPO_CODEPORT="' + CodePort + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_Port.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end else
    begin
      if not Tob_Port.InsertDB(nil) then Result := ErrInsertDB;
    end;
    Tob_Port.free;
  end;

end;

function ImportTOBProfilArt(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_profilArt: TOB;
  SQL: string;
  profilArt   : string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Pas de contrôle
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    // Création d'une TOB contenant  l'enregistrement courant
    Tob_profilArt := TOB.Create('PROFILART', nil, -1);
    Tob_profilArt.Dupliquer(TPST, False, True, True);
    Tob_profilArt.SetAllModifie(True);

    profilArt   := Tob_profilArt.GetValue('GPF_PROFILARTICLE');

    SQL := 'Select GPF_PROFILARTICLE From PROFILART WHERE GPF_PROFILARTICLE="' + profilArt + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_profilArt.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end else
    begin
      if not Tob_profilArt.InsertDB(nil) then Result := ErrInsertDB;
    end;
    Tob_profilArt.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB REGION
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBRegion(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Region: TOB;
  SQL: string;
  Pays   : string;
  Region : string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Pas de contrôle
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    // Création d'une TOB contenant  l'enregistrement courant
    Tob_Region := TOB.Create('REGION', nil, -1);
    Tob_Region.Dupliquer(TPST, False, True, True);
    Tob_Region.SetAllModifie(True);

    Pays   := Tob_Region.GetValue('RG_PAYS');
    Region := Tob_Region.GetValue('RG_REGION');

    SQL := 'Select RG_PAYS From REGION WHERE RG_PAYS="' + Pays + '" AND RG_REGION="' + Region + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_Region.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end else
    begin
      if not Tob_Region.InsertDB(nil) then Result := ErrInsertDB;
    end;
    Tob_Region.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB SOCIETE pour récupérer le numéro de version du site
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBSociete(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB; CodeSite: string): integer;
var
//  LesSites : TCollectionSites;
  LeSite   : TCollectionSite;
  Variable : TCollectionVariable;
  Tob_Vars : TOB;
  Q: Tquery;
  NumeroVersionTox: integer;
  NumeroVersionSiteDistant: integer;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Pas de contrôle
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    //
    // La variable $VERSION existe elle ?
    //
    Tob_Vars := TOB.Create('STOXVARS', nil, -1);

    Q := OpenSQL('Select * from STOXVARS Where SVA_NOMVARIABLE="$VERSION"', False);
    if not Q.EOF then
    begin
      Tob_Vars.SelectDB('', Q);
      Ferme(Q);
    end else
    begin
      Ferme(Q);
      Tob_Vars.PutValue('SVA_NOMVARIABLE', '$VERSION');
      Tob_Vars.PutValue('SVA_LIBELLE', 'Version PGI du site');
      Tob_Vars.PutValue('SVA_TYPEDATA', 'INT');
      Tob_Vars.PutValue('SVA_VDCHAR', '');
      Tob_Vars.PutValue('SVA_VDINT', 0);
      Tob_Vars.PutValue('SVA_VDBOOL', '-');
      Tob_Vars.PutValue('SVA_VDDAT', StrToDate('01/01/1900'));
      Tob_Vars.PutValue('SVA_DATECREATION', Nowh);
      Tob_Vars.PutValue('SVA_DATEMODIF', Nowh);
      Tob_Vars.PutValue('SVA_CREATEUR', V_PGI.USER);
      Tob_Vars.PutValue('SVA_UTILISATEUR', V_PGI.USER);

      Tob_Vars.InsertDB(nil);
      Tob_Vars.free;
    end;

    TX_LesVariables.Reload(nil) ;

    // Rechargement de tous les sites
    TX_LesSites.Reload() ;

    if CodeSite = '' then exit ;

    //
    // Numero de version du site distant
    //
    NumeroVersionSiteDistant := TPST.GetValue('SO_VERSIONBASE');
    //
    // Chargement du site
    //
    LeSite := TX_LesSites.Find(CodeSite);
    if LeSite <> nil then
    begin
      Variable := LeSite.LesVariables.Find('$VERSION');
      if Variable <> nil then
      begin
        NumeroVersionTox := ValeurI(Variable.SVA_VALUE);

        if NumeroVersionTox < NumeroVersionSiteDistant then
        begin
          Variable.SVA_VALUE := IntToStr(NumeroVersionSiteDistant);
          LeSite.Update();
        end;
      end
      else
      begin
        // Erreur !!
      end ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB STOXQUERYS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBStoxquerys(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Stoxquerys: TOB;
  SQL: string;
  TypeTrt: string;
  CodeRequete: string;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Pas de contrôle
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    // Création d'une TOB contenant  l'enregistrement courant
    Tob_Stoxquerys := TOB.Create('STOXQUERYS', nil, -1);
    Tob_Stoxquerys.Dupliquer(TPST, False, True, True);
    Tob_Stoxquerys.SetAllModifie(True);

    TypeTrt := Tob_Stoxquerys.GetValue('SQE_TYPETRT');
    CodeRequete := Tob_Stoxquerys.GetValue('SQE_CODEREQUETE');

    SQL := 'Select SQE_TYPEREQUETE From STOXQUERYS WHERE SQE_TYPETRT="' + TypeTrt + '" AND SQE_CODEREQUETE="' + CodeRequete + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_Stoxquerys.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end else
    begin
      if not Tob_Stoxquerys.InsertDB(nil) then Result := ErrInsertDB;
    end;
    Tob_Stoxquerys.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TARIF
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTarif(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Tarif: TOB;
  CodeTarif: string;
  NatureAuxi: string;
  StRech: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      ////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('GF_ARTICLE');
      if (StRech <> '') then
      begin
        SQL := 'Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrArticle;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Dépôt
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('GF_DEPOT');
        if (StRech <> '') then
        begin
          SQL := 'Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrDepot;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Devise
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('GF_DEVISE');
        if (StRech <> '') then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDevise;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Client ou Fournisseur
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('GF_TIERS');
        if (StRech <> '') then
        begin
          NatureAuxi := TPST.GetValue('GF_NATUREAUXI');
          if (NatureAuxi = 'CLI') or (NatureAuxi = 'FOU') then
          begin
            SQL := 'Select T_LIBELLE From TIERS WHERE T_NATUREAUXI="' + NatureAuxi + '" AND T_TIERS="' + StRech + '"';
            Q := OpenSQL(SQL, True);
            if Q.EOF then
            begin
              if NatureAuxi = 'CLI' then Result := ErrClient
              else Result := ErrFournisseur;
            end;
            Ferme(Q);
          end;
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Tarif Article
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('GF_TARIFARTICLE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TAR" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTarifArticle;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant  l'enregistrement courant
 ////////////////////////////////////////////////////////////////
    Tob_Tarif := TOB.Create('TARIF', nil, -1);
    Tob_Tarif.Dupliquer(TPST, False, True, True);
    Tob_Tarif.SetAllModifie(True);
    Tob_Tarif.PutValue('GF_DATEINTEGR', NowH);
    Tob_Tarif.SetDateModif(Tob_Tarif.GetValue('GF_DATEMODIF'));

    CodeTarif := Tob_Tarif.GetValue('GF_TARIF');
    SQL := 'Select GF_LIBELLE From TARIF WHERE GF_TARIF="' + CodeTarif + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Tarif.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Tarif.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Tarif.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MM JTR
Créé le ...... : 21/09/2004
Modifié le ... :   /  /
Description .. : Contrôles/Import yTarif
Mots clefs ... : TOX
*****************************************************************}
(*
function ImportTOByTarifs(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Tarif: TOB;
  CodeTarif: string;
  Q: TQUERY;
  TomYTS : Tom;
begin
  Result := IntegOK;
  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      TomYTS := CreateTOM('YTARIFS', nil, false, true);
      try
        { Paramètres }
        TOM_yTarifs(TomYTS).sAppel           := 'Tox';
        TOM_yTarifs(TomYTS).lFicheEntete     := False;
        TOM_yTarifs(TomYTS).lModifClefEntete := True;
        TOM_yTarifs(TomYTS).lModifClefLignes := True;
        TOM_yTarifs(TomYTS).lModifConditions := True;
        TOM_yTarifs(TomYTS).sClefEnteteTarif :=
             'YTS_DEVISE;YTS_DEPOT;YTS_TARIFDEPOT;YTS_TIERS;YTS_TARIFTIERS;YTS_TARIFSPECIAL;'
            +'YTS_AFFAIRE;YTS_TARIFAFFAIRE;YTS_SITE;'
            +'YTS_ARTICLE;YTS_TARIFARTICLE;YTS_FAMILLENIV1;YTS_FAMILLENIV2;YTS_FAMILLENIV3;'
            +'YTS_CIRCUIT;YTS_PHASE;YTS_TRAITEMENT;YTS_TTCOUHT;YTS_DATEDEBUT;YTS_DATEFIN';
        TOM_yTarifs(TomYTS).sClefLignesTarif := TOM_yTarifs(TomYTS).sClefEnteteTarif;
        TPST.AddChampSupValeur('IKC', 'C', false);
        if TomYTS.VerifTOB(TPST) then
        begin
        end;
        Result := TOM_yTarifs(TomYTS).LastError;
        if Result = 3 then Result := 64 //Tiers
        else if Result = 8 then Result := 95 //Affaire
        else if Result = 28 then Result := 40 //Dépôt
        else if Result = 14 then Result := 39 //Article
        else Result := 0;
      finally
        TomYTS.Free;
      end;
    end;
  end else
  begin // PHASE D'INTEGRATION
    Tob_Tarif := TOB.Create('YTARIFS', nil, -1);
    Tob_Tarif.Dupliquer(TPST, False, True, True);
    Tob_Tarif.SetAllModifie(True);
    Tob_Tarif.SetDateModif(Tob_Tarif.GetValue('YTS_DATEMODIF'));

    CodeTarif := Tob_Tarif.GetString('YTS_IDENTIFIANT');
    if ExisteSQL('Select 1 From YTARIFS WHERE YTS_IDENTIFIANT='+ CodeTarif ) then
    begin
      if not Tob_Tarif.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end else
    begin
      if not Tob_Tarif.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Tarif.free;
  end;
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 22/09/2004
Modifié le ... :   /  /
Description .. : Contrôles/Import yTarifsFourchette
Mots clefs ... : TOX
*****************************************************************}
(*
function ImportTOByTarifsFourchette(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var TobTarifDet : TOB;
    Id, IdEnTete : integer;
    StRech, SQL : string;
begin
  Result := IntegOK;
  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Test si l'en tête existe
      StRech := TPST.GetString('YTF_IDENTIFIANTYTS');
      if (StRech <> '') then
      begin
        if not ExisteSQL('Select 1 From YTARIFS WHERE YTS_IDENTIFIANT=' + StRech ) then
          result := ErrTarif;
      end;
    end;
  end else
  begin // PHASE D'INTEGRATION
    TobTarifDet := TOB.Create('YTARIFSFOURCHETTE', nil, -1);
    TobTarifDet.Dupliquer(TPST, False, True, True);
    TobTarifDet.SetAllModifie(True);
    Id := TobTarifDet.GetInteger('YTF_IDENTIFIANT');
    IdEnTete := TobTarifDet.GetInteger('YTF_IDENTIFIANTYTS');
    SQL := 'SELECT 1 FROM YTARIFSFOURCHETTE WHERE YTF_IDENTIFIANTYTS=' + IntToStr(IdEnTete)
    + ' AND YTF_IDENTIFIANT=' + IntToStr(Id);
    if ExisteSQL(SQL) then
    begin
      if not TobTarifDet.UpdateDB(FALSE) then
        Result := ErrUpdateDB;
    end else
      if not TobTarifDet.InsertDB(nil) then
        Result := ErrInsertDB;
    if TobTarifDet <> nil then FreeAndNil(TobTarifDet);
  end;
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TARIFMODE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTarifMode(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  SQL, StRech: string;
  CodeTarifMode: integer;
  Tob_TarifMode: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin

      // Devise
      StRech := TPST.GetValue('GFM_DEVISE');
      if (StRech <> '') then
      begin
        SQL := 'Select D_DEVISE From DEVISE WHERE D_DEVISE="' + StRech + '"';
        if not ExisteSQL(SQL) then result := ErrDevise;
      end;

      // Type d'arrondi
      if result = IntegOK then
      begin
        StRech := TPST.GetValue('GFM_ARRONDI');
        if (StRech <> '') then
        begin
          SQL := 'Select GAR_CODEARRONDI From ARRONDI WHERE GAR_CODEARRONDI="' + StRech + '"';
          if not ExisteSQL(SQL) then result := ErrTypeArrondi;
        end;
      end;

      // Type de démarque
      if result = IntegOK then
      begin
        StRech := TPST.GetValue('GFM_DEMARQUE');
        if (StRech <> '') then
        begin
          SQL := 'Select GTR_LIBELLE From TYPEREMISE WHERE GTR_TYPEREMISE="' + StRech + '"';
          if not ExisteSQL(SQL) then result := ErrTypeDemarque ;
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TarifMode := TOB.Create('TARIFMODE', nil, -1);
    Tob_TarifMode.Dupliquer(TPST, False, True, True);
    Tob_TarifMode.SetAllModifie(True);

    CodeTarifMode := Tob_TarifMode.GetValue('GFM_TARFMODE');

    SQL := 'Select GFM_TARFMODE From TARIFMODE WHERE GFM_TARFMODE="' + IntToStr(CodeTarifMode) + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_TarifMode.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
      if not Tob_TarifMode.InsertDB(nil) then Result := ErrInsertDB;
    Tob_TarifMode.free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TARIFPER
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTarifPer(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  SQL: string;
  StRech: string;
  CodePeriode: string;
  Etab: string;
  Tob_TarifPer: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin

      // Type d'arrondi
      if result = IntegOK then
      begin
        StRech := TPST.GetValue('GFP_ARRONDI');
        if (StRech <> '') then
        begin
          SQL := 'Select GAR_CODEARRONDI From ARRONDI WHERE GAR_CODEARRONDI="' + StRech + '"';
          if not ExisteSQL(SQL) then result := ErrTypeArrondi;
        end;
      end;

      // Type de démarque
      if result = IntegOK then
      begin
        StRech := TPST.GetValue('GFP_DEMARQUE');
        if (StRech <> '') then
        begin
          SQL := 'Select GTR_LIBELLE From TYPEREMISE WHERE GTR_TYPEREMISE="' + StRech + '"';
          if not ExisteSQL(SQL) then result := ErrTypeDemarque ;
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TarifPer := TOB.Create('TARIFPER', nil, -1);
    Tob_TarifPer.Dupliquer(TPST, False, True, True);
    Tob_TarifPer.SetAllModifie(True);

    CodePeriode := Tob_TarifPer.GetValue('GFP_CODEPERIODE');
    Etab := Tob_TarifPer.GetValue('GFP_ETABLISSEMENT');

    SQL := 'Select GFP_CODEPERIODE From TARIFPER WHERE GFP_CODEPERIODE="' + CodePeriode + '" AND GFP_ETABLISSEMENT="' + Etab + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_TarifPer.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
      if not Tob_TarifPer.InsertDB(nil) then Result := ErrInsertDB;
    Tob_TarifPer.free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TARIFTYPMODE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTarifTypMode(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  SQL, StRech, CodeType: string;
  Tob_TarifTypMode: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin

      // Devise
      StRech := TPST.GetValue('GFT_DEVISE');
      if (StRech <> '') then
      begin
        SQL := 'Select D_DEVISE From DEVISE WHERE D_DEVISE="' + StRech + '"';
        if not ExisteSQL(SQL) then result := ErrDevise;
      end
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TarifTypMode := TOB.Create('TARIFTYPMODE', nil, -1);
    Tob_TarifTypMode.Dupliquer(TPST, False, True, True);
    Tob_TarifTypMode.SetAllModifie(True);

    CodeType := Tob_TarifTypMode.GetValue('GFT_CODETYPE');

    SQL := 'Select GFT_CODETYPE From TARIFTYPMODE WHERE GFT_CODETYPE="' + CodeType + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_TarifTypMode.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
      if not Tob_TarifTypMode.InsertDB(nil) then Result := ErrInsertDB;
    Tob_TarifTypMode.free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TIERS
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTiers(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Tiers: TOB;
  Auxiliaire: string;
  SQL: string;
  StRech: string;
  Q: TQUERY;
  DevOrg: string;
  DevDest: string;
  DateConv: TDateTime;
  Montant: double;
  MontantOld: double;
  MontantDev: double;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pays
      /////////////////////////////////////////////////////////////////////////
      StRech := TPST.GetValue('T_PAYS');
      if (StRech <> '') then
      begin
        SQL := 'Select PY_LIBELLE From PAYS WHERE PY_PAYS="' + StRech + '"';
        Q := OpenSQL(SQL, True);
        if Q.EOF then result := ErrPays;
        Ferme(Q);
      end;
      /////////////////////////////////////////////////////////////////////////
      // Langue
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('T_LANGUE');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrLangue;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Devise
      /////////////////////////////////////////////////////////////////////////
      if Result = IntegOk then
      begin
        StRech := TPST.GetValue('T_DEVISE');
        if (StRech <> '') then
        begin
          SQL := 'Select D_LIBELLE From DEVISE WHERE D_DEVISE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then Result := ErrDevise;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Secteur d'activité
      /////////////////////////////////////////////////////////////////////////
      // if (result=IntegOK) then
      //begin
      //  StRech := TPST.GetValue ('T_SECTEUR');
      //  if (StRech <> '') then
      //  begin
      //		SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="SCC" AND CC_CODE="'+StRech+'"';
      //		Q:=OpenSQL(SQL,True) ;
      //		if Q.EOF then result := ErrSecteurAct;
      //    Ferme(Q);
      //  end;
      //end;
      /////////////////////////////////////////////////////////////////////////
      // Zone commerciale
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_ZONECOM');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GZC" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrZoneCom;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Code tarif tiers
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_TARIFTIERS');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TRC" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrTarifTiers;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Mode de règlement
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_MODEREGLE');
        if (StRech <> '') then
        begin
          SQL := 'Select MR_LIBELLE From MODEREGL WHERE MR_MODEREGLE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrModeRegle;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Régime TVA
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_REGIMETVA');
        if (StRech <> '') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrRegimeTVA;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      // Civilité
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_JURIDIQUE');
        if (StRech <> '') and (TPST.GetValue('T_PARTICULIER') = 'X') then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="CIV" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrCivilite;
          Ferme(Q);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////
      // Nationalité
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_NATIONALITE');
        if StRech <> '' then
        begin
          SQL := 'Select PY_LIBELLE From PAYS WHERE PY_PAYS="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrNationalite;
          Ferme(Q);
        end;
      end;
      ////////////////////////////////////////////////////////////////////////
      // Famille Comptable
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_COMPTATIERS');
        if StRech <> '' then
        begin
          SQL := 'Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GCT" AND CC_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrFamComptable;
          Ferme(Q);
        end;
      end;
      /////////////////////////////////////////////////////////////////////////
      //Sexe
      /////////////////////////////////////////////////////////////////////////
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('T_SEXE');
        if (StRech <> '') then
        begin
          SQL := 'Select CO_LIBELLE From COMMUN WHERE CO_TYPE="JMF" AND CO_CODE="' + StRech + '"';
          Q := OpenSQL(SQL, True);
          if Q.EOF then result := ErrSexe;
          Ferme(Q);
        end;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Tiers := TOB.Create('TIERS', nil, -1);
    Tob_Tiers.Dupliquer(TPST, False, True, True);
    //
    // Conversion dans la devise dossier
    //
    if (sDevise <> V_PGI.DeviseFongible) or (sBasculeEuro <> VH^.TenueEuro) then
    begin
      //
      // Devise d'origine : si le site a basculé en EURO, c'est de l'EURO
      //
      if sBasculeEuro then DevOrg := 'EUR'
      else DevOrg := sDevise;
      //
      // Devise destinataire : devise de tenue du dossier qui intègre
      //
      DevDest := V_PGI.DevisePivot;
      Dateconv := NowH;

      // Franco de port
      Montant := Tob_Tiers.GetValue('T_FRANCO');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_FRANCO', MontantDev);
      // Total débit
      Montant := Tob_Tiers.GetValue('T_TOTALDEBIT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTALDEBIT', MontantDev);
      // Total crédit
      Montant := Tob_Tiers.GetValue('T_TOTALCREDIT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTALCREDIT', MontantDev);
      // Débit dernier mouvement
      Montant := Tob_Tiers.GetValue('T_DEBITDERNMVT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_DEBITDERNMVT', MontantDev);
      // Crédit dernier mouvement
      Montant := Tob_Tiers.GetValue('T_CREDITDERNMVT');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_CREDITDERNMVT', MontantDev);
      // Coût horaire
      Montant := Tob_Tiers.GetValue('T_COUTHORAIRE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_COUTHORAIRE', MontantDev);
      // Total débit N-1
      Montant := Tob_Tiers.GetValue('T_TOTDEBP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDEBP', MontantDev);
      // Total crédit N-1
      Montant := Tob_Tiers.GetValue('T_TOTCREP');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTCREP', MontantDev);
      // Total débit N
      Montant := Tob_Tiers.GetValue('T_TOTDEBE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDEBE', MontantDev);
      // Total créit N
      Montant := Tob_Tiers.GetValue('T_TOTCREE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTCREE', MontantDev);
      // Total débit N+1
      Montant := Tob_Tiers.GetValue('T_TOTDEBS');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDEBS', MontantDev);
      // Total crédit N+1
      Montant := Tob_Tiers.GetValue('T_TOTCRES');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTCRES', MontantDev);
      // Total débit à nouveaux
      Montant := Tob_Tiers.GetValue('T_TOTDEBANO');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDEBANO', MontantDev);
      // Total crédit à nouveaux
      Montant := Tob_Tiers.GetValue('T_TOTCREANO');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTCREANO', MontantDev);
      // A nouveaux provisoire N+1
      Montant := Tob_Tiers.GetValue('T_TOTDEBANON1');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDEBANON1', MontantDev);
      // A nouveaux provisoire N+1
      Montant := Tob_Tiers.GetValue('T_TOTCREANON1');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTCREANON1', MontantDev);
      // Total HT dernière pièce
      Montant := Tob_Tiers.GetValue('T_TOTDERNPIECE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_TOTDERNPIECE', MontantDev);
      // Crédit demandé
      Montant := Tob_Tiers.GetValue('T_CREDITDEMANDE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_CREDITDEMANDE', MontantDev);
      // Crédit accordé
      Montant := Tob_Tiers.GetValue('T_CREDITACCORDE');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_CREDITACCORDE', MontantDev);
      // Plafond de crédit autorisé
      Montant := Tob_Tiers.GetValue('T_CREDITPLAFOND');
      ToxConvertir(Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, TobDev);
      Tob_Tiers.PutValue('T_CREDITPLAFOND', MontantDev);
    end;

    ////////////////////////////////////////////////////////////////
    // Faut il masquer les coordonnées des fournisseurs
    ///////////////////////////////////////////////////////////////
    if (Tob_Tiers.GetValue('T_NATUREAUXI') = 'FOU') and (GetParamSoc('SO_GCTOXCACHEFOU') = True) then
    begin
      Tob_Tiers.putValue('T_ADRESSE1', '');
      Tob_Tiers.putValue('T_ADRESSE2', '');
      Tob_Tiers.putValue('T_ADRESSE3', '');
      Tob_Tiers.putValue('T_CODEPOSTAL', '');
      Tob_Tiers.putValue('T_VILLE', '');
      Tob_Tiers.putValue('T_PAYS', '');
      Tob_Tiers.putValue('T_TELEPHONE', '');
      Tob_Tiers.putValue('T_FAX', '');
      Tob_Tiers.putValue('T_TELEX', '');
      Tob_Tiers.putValue('T_TELEPHONE2', '');
    end;

    Tob_Tiers.SetAllModifie(True);
    Tob_Tiers.PutValue('T_DATEINTEGR', NowH);
    Tob_Tiers.SetDateModif(Tob_Tiers.GetValue('T_DATEMODIF'));

    Auxiliaire := Tob_Tiers.GetValue('T_AUXILIAIRE');
    SQL := 'Select T_LIBELLE From TIERS WHERE T_AUXILIAIRE="' + Auxiliaire + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Tiers.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Tiers.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Tiers.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TIERSCOMPL
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTiersCompl(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_TiersCompl: TOB;
  Auxiliaire: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle
    end;
  end
  else
  begin // Phase d'intégration
    ///////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_TiersCompl := TOB.Create('TIERSCOMPL', nil, -1);
    Tob_TiersCompl.Dupliquer(TPST, False, True, True);
    Tob_TiersCompl.SetAllModifie(True);

    Auxiliaire := Tob_TiersCompl.GetValue('YTC_AUXILIAIRE');
    SQL := 'Select YTC_TIERS From TIERSCOMPL WHERE YTC_AUXILIAIRE="' + Auxiliaire + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_TiersCompl.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_TiersCompl.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_TiersCompl.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADICCO
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTradDico(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  SQL, CodeDico: string;
  Tob_TradDico: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // pas de contrôle
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TradDico := TOB.Create('TRADDICO', nil, -1);
    Tob_TradDico.Dupliquer(TPST, False, True, True);
    Tob_TradDico.SetAllModifie(True);

    CodeDico := Tob_TradDico.GetValue('DX_FRA');

    SQL := 'Select DX_FRA From TRADDICO WHERE DX_FRA="' + CodeDico + '"';
    if ExisteSQL(SQL) then
    begin
      if not Tob_TradDico.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
      if not Tob_TradDico.InsertDB(nil) then Result := ErrInsertDB;
    Tob_TradDico.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADTABLETTE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTradTablette(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  sLangue, sTypeTablette, sCodeTablette, StRech: string;
  Tob_TradTablette: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      //
      // Langue
      //
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('YTT_LANGUE');
        if (StRech <> '') then
          if not ExisteSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"') then
            result := ErrLangue;
      end;
      //
      // Code
      //
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('YTT_CODE');
        if (StRech <> '') then
          if not ExisteSQL('Select DO_COMBO From DECOMBOS WHERE DO_COMBO="' + StRech + '"') then
            result := ErrCodeTablette;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant  l'enregistrement courant
    Tob_TradTablette := TOB.Create('TRADTABLETTE', nil, -1);
    Tob_TradTablette.Dupliquer(TPST, False, True, True);
    Tob_TradTablette.SetAllModifie(True);

    sLangue := Tob_TradTablette.GetValue('YTT_LANGUE');
    sTypeTablette := Tob_TradTablette.GetValue('YTT_TYPE');
    sCodeTablette := Tob_TradTablette.GetValue('YTT_CODE');

    if ExisteSQL('Select YTT_LANGUE From TRADTABLETTE WHERE YTT_LANGUE="' + sLangue + '" AND YTT_TYPE="' + sTypeTablette + '" AND YTT_CODE="' + sCodeTablette + '"') then
    begin
      if not Tob_TradTablette.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
    begin
      if not Tob_TradTablette.InsertDB(nil) then Result := ErrInsertDB;
    end;

    Tob_TradTablette.free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADUC
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTraduc(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  sLangue, sForme, StRech: string;
  Tob_Traduc: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      //
      // Langue
      //
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('TR_LANGUE');
        if (StRech <> '') then
          if not ExisteSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="' + StRech + '"') then
            result := ErrLangue;
      end;
      //
      // Forme
      //
      if (result = IntegOK) then
      begin
        StRech := TPST.GetValue('TR_FORME');
        if (StRech <> '') then
          if not ExisteSQL('Select DFM_FORME From FORMES WHERE DFM_TYPEFORME="GC" AND DFM_FORME="' + StRech + '"') then
            result := ErrForme;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant l'enregistrement courant
    Tob_Traduc := TOB.Create('TRADUC', nil, -1);
    Tob_Traduc.Dupliquer(TPST, False, True, True);
    Tob_Traduc.SetAllModifie(True);

    sLangue := Tob_Traduc.GetValue('TR_LANGUE');
    sForme := Tob_Traduc.GetValue('TR_FORME');

    if ExisteSQL('Select TR_LANGUE From TRADUC WHERE TR_LANGUE="' + sLangue + '" AND TR_FORME="' + sForme + '"') then
    begin
      if not Tob_Traduc.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end
    else
    begin
      if not Tob_Traduc.InsertDB(nil) then Result := ErrInsertDB;
    end;

    Tob_Traduc.free;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TRADUC
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTxCptTva(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  StRech: string;
  TVAOUTPF: string;
  CodeTaux: string;
  Regime: string;
  Tob_TxCptTva: TOB;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      //
      // Régime fiscal
      //
      StRech := TPST.GetValue('TV_REGIME');
      if (StRech <> '') then
      begin
        if not ExisteSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="' + StRech + '"') then result := ErrRegime;
      end;
    end;
  end
  else
  begin // PHASE D'INTEGRATION

    // Création d'une TOB contenant l'enregistrement courant
    Tob_TxCptTva := TOB.Create('TXCPTTVA', nil, -1);
    Tob_TxCptTva.Dupliquer(TPST, False, True, True);
    Tob_TxCptTva.SetAllModifie(True);

    TVAOUTPF := Tob_TxCptTva.GetValue('TV_TVAOUTPF');
    CodeTaux := Tob_TxCptTva.GetValue('TV_CODETAUX');
    Regime := Tob_TxCptTva.GetValue('TV_REGIME');

    if ExisteSQL('Select TV_SOCIETE From TXCPTTVA WHERE TV_TVAOUTPF="' + TVAOUTPF + '" AND TV_CODETAUX="' + CodeTaux + '" AND TV_REGIME="' + Regime + '"') then
    begin
      if not Tob_TxCptTva.UpdateDB(FALSE) then Result := ErrUpdateDB;
    end else
    begin
      if not Tob_TxCptTva.InsertDB(nil) then Result := ErrInsertDB;
    end;

    Tob_TxCptTva.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TYPEMASQUE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTypeMasque(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_TypeMasque: TOB;
  TypeMasque: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
       // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_TypeMasque := TOB.Create('TYPEMASQUE', nil, -1);
    Tob_TypeMasque.Dupliquer(TPST, False, True, True);
    Tob_TypeMasque.SetAllModifie(True);

    TypeMasque := Tob_TypeMasque.GetValue('GMQ_TYPEMASQUE');
    SQL := 'Select GMQ_TYPEMASQUE From TYPEMASQUE WHERE GMQ_TYPEMASQUE="' + TypeMasque + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_TypeMasque.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_TypeMasque.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_TypeMasque.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L Meunier
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB TYPEMASQUE
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBTypeRemise(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_TypeRemise: TOB;
  TypeRemise: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
       // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_TypeRemise := TOB.Create('TYPEREMISEE', nil, -1);
    Tob_TypeRemise.Dupliquer(TPST, False, True, True);
    Tob_TypeRemise.SetAllModifie(True);

    TypeRemise := Tob_TypeRemise.GetValue('GTR_TYPEREMISE');
    SQL := 'Select GTR_TYPEREMISE From TYPEREMISE WHERE GTR_TYPEREMISE="' + TypeRemise + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_TypeRemise.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_TypeRemise.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_TypeRemise.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB USERGRP
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBUserGrp(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_UserGrp: TOB;
  Groupe: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      // Contrôle(s)
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
     // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_UserGrp := TOB.Create('USERGRP', nil, -1);
    Tob_UserGrp.Dupliquer(TPST, False, True, True);
    Tob_UserGrp.SetAllModifie(True);

    Groupe := Tob_UserGrp.GetValue('UG_GROUPE');
    SQL := 'Select UG_LIBELLE From USERGRP WHERE UG_GROUPE="' + Groupe + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_UserGrp.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_UserGrp.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_UserGrp.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : D. Carret
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : Contrôles/Import TOB UTILISAT
Mots clefs ... : TOX
*****************************************************************}
function ImportTOBUtilisat(TPST: TOB; Integre, IntegrationRapide: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB): integer;
var
  Tob_Utilisat: TOB;
  Utilisat: string;
  SQL: string;
  Q: TQUERY;
begin

  Result := IntegOK;

  if not (Integre) then
  begin
    if not (IntegrationRapide) then // PHASE DE CONTROLE
    begin
      /////////////////////////////////////////////////////////////////////////
      // Pas de contrôle
      /////////////////////////////////////////////////////////////////////////
    end;
  end
  else
  begin // PHASE D'INTEGRATION
    ////////////////////////////////////////////////////////////////
       // Création d'une TOB contenant l'enregistrement courant
    ////////////////////////////////////////////////////////////////
    Tob_Utilisat := TOB.Create('UTILISAT', nil, -1);
    Tob_Utilisat.Dupliquer(TPST, False, True, True);
    Tob_Utilisat.SetAllModifie(True);

    Utilisat := Tob_Utilisat.GetValue('US_UTILISATEUR');
    SQL := 'Select US_UTILISATEUR From UTILISAT WHERE US_UTILISATEUR="' + Utilisat + '"';
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
    begin
      if not Tob_Utilisat.UpdateDB(FALSE) then
      begin
        Result := ErrUpdateDB;
      end;
    end
    else
    begin
      if not Tob_Utilisat.InsertDB(nil) then
      begin
        Result := ErrInsertDB;
      end;
    end;
    Ferme(Q);
    Tob_Utilisat.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Fonction générique qui chaîne sur les fonctions
Suite ........ : spécifiques aux tables traitées
Mots clefs ... : TOX
*****************************************************************}
function ToxControleInfo(NomTable: string; var ToxInfo: TOB; Integre, IntegrationRapide, MajStock: boolean; sDevise: string; sBasculeEuro: boolean; TobDev: TOB; CodeSite: string):
  integer;
begin
  if (NomTable = 'ACOMPTES') then result := ImportTOBAcomptes(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'AFFAIRE') then result := ImportTOBAffaire(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ADRESSES') then result := ImportTOBAdresses(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARRONDI') then result := ImportTOBArrondi(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARTICLE') then result := ImportTOBArticle(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARTICLECOMPL') then result := ImportTOBArticleCompl(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARTICLELIE') then result := ImportTOBArticleLie(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARTICLEPIECE') then result := ImportTOBArticlePiece(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ARTICLETIERS') then result := ImportTOBArticleTiers(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
{$IFDEF BTP}
  else if (NomTable = 'BMEMORISATION') then result := ImportTOBMemorisation(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
{$ENDIF}
  else if (NomTable = 'CATALOGU') then result := ImportTOBCatalogu(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CHANCELL') then result := ImportTOBChancell(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CHOIXCOD') then result := ImportTOBChoixCod(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CHOIXEXT') then result := ImportTOBChoixExt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CLAVIERECRAN') then result := ImportTOBClavierEcran(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CODEPOST') then result := ImportTOBCodepost(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'COMMENTAIRE') then result := ImportTOBCommentaire(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'COMMERCIAL') then result := ImportTOBCommercial(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'COMPTEURETAB') then result := ImportTOBCompteurEtab(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CONTACT') then result := ImportTOBContact(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'CTRLCAISMT') then result := ImportTOBCtrlCaisMt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'DEPOTS') then result := ImportTOBDepots(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'DEVISE') then result := ImportTOBDevise(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'DIMMASQUE') then result := ImportTOBDimasque(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'DIMENSION') then result := ImportTOBDimension(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'DISPO') then result := ImportTOBDispo(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'ETABLISS') then result := ImportTOBEtabliss(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'EXERCICE') then result := ImportTOBExercice(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'FONCTION') then result := ImportTOBFonction(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'FIDELITEENT') then result := ImportTOBFideliteEnt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'FIDELITELIG') then result := ImportTOBFideliteLig(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'JOURSCAISSE') then result := ImportTOBJoursCaisse(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'JOURSETAB') then result := ImportTOBJoursEtab(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'JOURSETABEVT') then result := ImportTOBJoursEtabEvt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'LIENSOLE') then result := ImportTOBLiensOLE(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'LIGNE') then result := ImportTOBLigne(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'LIGNELOT') then result := ImportTOBLigneLot(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'LIGNENOMEN') then result := ImportTOBLigneNomen(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
{$IFDEF BTP}
  else if (NomTable = 'LIGNEOUV') then result := ImportTOBLigneOuvrages(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'NATUREPREST') then result := ImportTOBNaturePrest(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'BTPIECEMILIEME') then result := ImportTOBMillieme(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'BTPARDOC') then result := ImportTOBPardoc(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIECERG') then result := ImportTobPieceRg(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
{$ENDIF}
  else if (NomTable = 'LISTEINVENT') then result := ImportTOBListeinvent(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'LISTEINVLIG') then result := ImportTOBListeInvLig(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'MEA') then result := ImportTOBMea(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'MODELES') then result := ImportTOBModeles(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'MODEDATA') then result := ImportTOBModeData(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'MODEREGL') then result := ImportTOBModeRegl(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'MODEPAIE') then result := ImportTOBModePaie(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'NOMENENT') then result := ImportTOBNomenEnt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'NOMENLIG') then result := ImportTOBNomenLig(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'OPERCAISSE') then result := ImportTOBOperCaisse(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARAMSOC') then result := ImportTOBParamSoc(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARAMTAXE') then result := ImportTOBParamTaxe(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARCAISSE') then result := ImportTOBParCaisse(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARFIDELITE') then result := ImportTOBParFidelite(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARREGLEFID') then result := ImportTOBParRegleFid(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARSEUILFID') then result := ImportTOBParSeuilFid(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PARPIECBIL') then result := ImportTOBParPiecBil(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PAYS') then result := ImportTOBPays(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIEDBASE') then result := ImportTOBPiedBase(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIEDECHE') then result := ImportTOBPiedEche(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIEDPORT') then result := ImportTOBPiedPort(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIECE') then result := ImportTOBPiece(ToxInfo, Integre, IntegrationRapide, MajStock, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PORT') then result := ImportTOBPort(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PIECEADRESSE') then result := ImportTOBPieceAdresse(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'PROFILART') then result := ImportTOBProfilArt(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'REGION') then result := ImportTOBRegion(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'SOCIETE') then result := ImportTOBSociete(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev, CodeSite)
  else if (NomTable = 'STOXQUERYS') then result := ImportTOBStoxQuerys(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TARIF') then result := ImportTOBTarif(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TARIFMODE') then result := ImportTOBTarifMode(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TARIFPER') then result := ImportTOBTarifPer(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TARIFTYPMODE') then result := ImportTOBTarifTypMode(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
//  else if (NomTable = 'YTARIFS') then result := ImportTOByTarifs(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
//  else if (NomTable = 'YTARIFSFOURCHETTE') then result := ImportTOByTarifsFourchette(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TIERS') then result := ImportTOBTiers(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TIERSCOMPL') then result := ImportTOBTiersCompl(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TRADDICO') then result := ImportTOBTradDico(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TRADTABLETTE') then result := ImportTOBTradTablette(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TRADUC') then result := ImportTOBTraduc(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TXCPTTVA') then result := ImportTOBTxCptTva(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TYPEMASQUE') then result := ImportTOBTypeMasque(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'TYPEREMISE') then result := ImportTOBTypeRemise(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'USERGRP') then result := ImportTOBUserGrp(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else if (NomTable = 'UTILISAT') then result := ImportTOBUtilisat(ToxInfo, Integre, IntegrationRapide, sDevise, sBasculeEuro, TobDev)
  else result := TableNonReconnu;
end;

end.


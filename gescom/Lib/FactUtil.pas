 unit FactUtil;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  {$IFDEF NOMADE}
  {$IFDEF EAGLCLIENT}
  UToxClasses,
  {$ENDIF}
  {$ENDIF}
  {$IFDEF BTP}
  FactUtilBTP,HrichOle,
  {$ENDIF}
  SysUtils, Dialogs, Utiltarif, TarifUtil, UtilPGI, UtilGC, AGLInit, EntGC, SaisUtil, Windows,
  Forms, Classes, AGLInitGC, UtilArticle, UtilDimArticle, HDimension, HMsgBox,
  FactTOB, FactArticle, FactPiece,FactLigneBase,
  ParamSoc, HPanel  ,uEntCommun,UtilTOBPiece, UtilsMetresXLS;

const NbRowsInit: integer = 100; // GM le 24/05/2002
  NbRowsPlus: integer = 20;
  // ----

type //T_RechArt = (traAucun,traErreur,traErrFour,traOk,traGrille,traOkGene,traSus,traContreM) ;
  T_ActionTarifArt = (ataOK, ataAlerte, ataCancel);
  // MOdif BTP
  TModeAff = set of (tModeBlocNotes, tModeGridMode, tModeGridStd,tModeSaisieBordereau);
  T_CtrlRefIntExt = (crInterne, crExterne);


  // Modif BTP
type R_CPercent = record
    Niveau: Integer;
    Depart: Integer;
    Fin: Integer;
    Traite: Integer;
  end;

  // DEFINITION DES TYPES DE PRESENTATION DES DETAILS D'OUVRAGES
const DOU_AUCUN: integer = 0;
  DOU_CODE: integer = 1;
  DOU_LIBELLE: integer = 2;
  DOU_QTE: integer = 4;
  DOU_UNITE: integer = 8;
  DOU_PU: integer = 16;
  DOU_MONTANT: integer = 32;
  DOU_TOUS: integer = 63;
  DOU_NUMP : integer = 128;
  // --------

  Type T_EtatSoldeLigne = (eslNone, eslEnCours, eslSolde, eslVerrouille);

var MessageValid : string;
  // Opérations sur les champs ou les zones
function IsComptaStock(Nature: string): boolean;
//AC 18/08/03 NV GESTION COMPTA DIFF
function IsComptaPce(Nature: string): boolean;
function IsComptaAcc(Nature: string): boolean;
//Fin AC
function IsLigneDetail(TOBPiece, TOBLOC: TOB; Ligne: integer = -1): boolean;
function IsLigneDetailMode(TOBLOC: TOB): boolean; // Rajout Mode
function CreerQuelDepot(TOBPIECE: TOB): string;
// Affaires
procedure ChangeTzAffaireSaisie(CodeTiers: string);
function TestAffaire(TOBPiece: TOB): boolean; //mcd 14/05/03
// PL le 07/11/02 : ajout du paramètre contexte d'appel
function GetAffaireEntete(TAffaire, TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit; bProposition, bChangeStatut, bModele, bChangeTiers, BcreatAff:
  Boolean; sContexte: string = ''; bReduit: boolean = true; bOuvreSiUn: boolean = true; bItDatesMax : boolean = false): boolean;
function GetAffaireEnteteSt(Taffaire0,TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit; var CodeAffaire: string; bProposition, bChangeStatut, bModele,
  bChangeTiers, BCreatAff: Boolean; sContexte: string = ''; bReduit: boolean = false; bOuvreSiUn: boolean = true; bItDatesMax : boolean = false; EtatAffaire : string = ''): boolean;
////////////////////////
function GetAffaireLigne(GS: THGrid; Tiers, NaturePieceG: string; bAffReduit: Boolean): Boolean;
function GetAffaireLookUp(GS: THGrid; TitreSel, CodeTiers: string): boolean;
// PL le 07/11/02 : ajout du paramètre contexte d'appel
function GetAffaireMul(var CodeAff: string; Taffaire0,TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit; bProposition, bChangeStatut, bModele,
  bChangetiers,BCreatAff, bReduit: Boolean; sContexte: string = ''; bOuvreSiUn: boolean = true; bIntervalleMaxi : boolean = false; EtatAffaire : string= ''): boolean;
  ////////////////////////
procedure ZoomAffaire(CodeAff: string);
function RemplirTOBAffaire(CodeAffaire: string; TobAffaire: Tob): boolean;
procedure AffaireVersLigne(TOBPiece, TOBL, TOBAffaire: TOB);
procedure AffaireVersPiece(TOBPiece, TOBAffaire: TOB);
procedure MajAnaAffaire(TOBPiece, TOBCataLogu, TOBArticles, TOBCpta, TOBTiers, TOBAnaP: TOB; ARow: integer);
{$IFDEF CHR}
procedure RemplirTOBHrdossier(Hrdossier: string; TobHrdossier: Tob);
procedure EcheVersPiece(TOBEches, TOBPiece: TOB);
procedure GetSommeReglements(TOBEches: TOB; var XP, XD: Double);
{$ENDIF} // FIN CHR
// Articles
procedure ArticleVersLigne(TOBPiece, TOBA, TOBConds, TOBL, TOBTiers: TOB; AvecPrix: Boolean = true);
procedure ArticleDetailversLigne(refSais : String; TOBart,TOBL : TOB; SaContexte : TModeAff);
//  BBI, fiche correction 10410
{$IFDEF GPAO}
function GetCircuitMul(GS: THGrid; TobPiece: tob; Article: string): boolean; //Tobl:Tob ) : boolean ;
{$ENDIF GPAO}
procedure CalculePrixArticle(TOBA: TOB; Depot: string = '');
//  BBI, fin fiche correction 10410
function TrouverArticleSQL(NaturePiece, RefSais, DomainePiece, SelectFourniss: string; DatePiece: TDateTime; TOBArt, TOBTiers: TOB; RefArt : String=''): T_RechArt;
procedure ZoomArticle(RefUnique, TypeArt: string; Action: TActionFiche);
procedure ErreurCodeArticle(RefSais: string);
procedure ErreurCodeArticleFour(RefSais: string);
// Catalogue
procedure CatalogueVersLigne(TOBPiece, TOBCata, TOBL, TOBTiers, TOBArtRef: TOB);
procedure ZoomCatalogue(RefCata, RefFour: string; Action: TActionFiche);
// Dimension
procedure LigneVersDim(TOBL, TOBD: TOB; QteDejaSaisi, ReliqDejaSaisi: Double);
function SelectUneDimension(sCodeGene: string): string;
function SelectMultiDimensions(sCodeGene: string; GS: THGrid; DepotTrf: string; Action: TActionFiche): Boolean; //AC
procedure InitDim(TOBD: TOB; Article: string; QteDejaSaisi, ReliqDejaSaisi: Double);
function CalcQteDejaSaisie(TOBPiece: TOB; Article: string): Double;
Function CalcMtDejaSaisie(TOBPiece: TOB; Article: string): Double;

function EtudieEtat(TOBArt: TOB; NaturePiece: string): T_RechArt;
function GetCondRecherche(GS: THGrid; CondOrig, RefUnique: string; TOBConds: TOB): string;
// Représentant
function GetWhereRepresentant(TOBG: TOB; ZoneCom: string; ModeSql: Boolean): string;
function GetRepresentantEntete(TC: THCritMaskEdit; TitreSel, ZoneCom: string; TOBPiece: TOB): boolean;
function GetRepresentantSaisi(GS: THGrid; TitreSel, ZoneCom: string; Letl: TTypeLocate; TOBPiece: TOB): boolean;

//FV1 : 20/02/2014 - FS#898 - BAGE : Ajouter une zone pour indiquer le contact du bon de commande
function GetRessourceEntete(TC: THCritMaskEdit; TitreSel, ZoneCom: string; TOBPiece: TOB): boolean;
function GetRessourceLookUp(CC: TControl; TitreSel: string; Letl: TTypeLocate; TOBPiece: TOB): boolean;
function GetRessourceMul(CC: TControl; ZoneCom: string; TOBPiece: TOB): boolean;

//
procedure InitialiseComm(TOBL: TOB; Aff: boolean);
// Dépôts, Divers
function GetDepotSaisi(GS: THGrid; TitreSel: string; Letl: TTypeLocate): boolean;
function GetLibellesAutos(GS: THGrid; TitreSel: string): boolean;
procedure GetCommentaireLigne(var CodeCom,Libelle,Texte : string );
procedure ZoomDispo(RefUnique, RefFour, Client, Four, Depot: string; Contrem: Boolean);
// TOBS
procedure LoadTOBDispo(TOBA: TOB; Force: boolean; QuelDepot: string = '');
procedure MajFromCleDoc(LaTOB: TOB; CleDoc: R_CleDoc);
// Lignes
// Ces deux fonctions sont transférées dans FactTarifs.pas
// function TarifVersLigne(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif: TOB; ForceTous, Totale: boolean; DEV: RDEVISE): T_ActionTarifArt;
// function TrouveTarif(TOBA, TOBTiers, TOBL, TOBPiece, TOBTarif: TOB; ForceTous, EnHT: boolean): boolean; //JD
Function PreAffecteLigne ( TobPiece, TobL, TobLigneTarif, TobA, TobTiers, TobTarif, TobConds, TobAffaire, TobCata, TobArtRef: Tob; ARow: integer; DEV: RDEVISE; NewQte: double; CalcTarif : boolean = True;CTRLCatFrs : Boolean = true ): T_ActionTarifArt;
procedure InitLigneVide(TOBPiece, TOBL, TOBTiers, TOBAffaire: TOB; ARow: integer; NewQte: double);
procedure MemoriseNumLigne(TobPiece : TOB);
procedure PieceVersLigne(TOBPiece, TOBL: TOB; WithAffaire : boolean=true; WithDomaine : boolean=true);
procedure PieceVersLigneExi(TOBPiece, TOBL: TOB);
procedure PieceVersLigneRessource(TOBPiece, TOBL: TOB; OldRes: string);
procedure CalculeSousTotauxPiece(TOBPiece: TOB);
// Affichages
procedure MontreInfosLigne(TOBLig, TOBArt, TOBCata, TOBTiers: TOB; Info: THGrid; PPInfosLigne: TStrings);
procedure ControlsVisible(NomPanel: THPanel; Visible: boolean; SfControlName: array of string);
// Divers
// Modif BTP
function AffectePrixDefaut(Prix, PrixDef1, PrixDef2: double): double;
procedure AffectePrixValo(TOBL, TOBA: TOB; TOBOuvrage: TOB = nil);
// ---
procedure AffecteValoNomen(TOBN, TOBA: TOB; Depot: string);
// Catalogue
procedure MajCatalogueSaisie(TOBArticles: TOB);
// Conversions
function ConvertSaisieEF(XSais: Double; ModeOppose: boolean): Double;
function ConvertSaisieEFNA(XSais: Double; ModeOppose: boolean): Double;
// Divers lignes
procedure ZeroLigne(TOBL: TOB);
procedure ZeroLigneMontant(TOBL : TOB) ;
// Modif BTP
procedure ValideLesRetenues(TOBPiece, TOBpieceRg: TOB);
procedure ValideLesBasesRg(TOBPiece, TOBBasesRg: TOB);
procedure ZeroLignePourcent(TOBL: TOB);
// ------
procedure SommerLignes(TOBPiece: TOB; ARow, Niv: integer);
function FlagDetruitArriere(TOBPiece: TOB): boolean;
// Acomptes
procedure ValideLesAcomptes(TOBPiece, TOBAcomptes: TOB; TOBAcomptes_O: TOB = nil);
procedure GetSommeAcomptes(TOBAcomptes: TOB; var XP, XD: Double; Fournisseur : string=''; All : Boolean=false);
procedure LoadLesAcomptes(TOBPiece, TOBAcomptes: TOB; CleDoc: R_CleDoc; TOBAcptesIni: TOb = nil);
// MOdif BTP
procedure LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG: TOB; Action : TactionFiche= taCreat);
procedure LoadLesRetenuesPRE(TOBPiece, TOBRGPRE: TOB; Action : TactionFiche= taCreat);
procedure LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, TOBTRFPOC: TOB; DEV: Rdevise; saisieTypeAvanc: boolean; TheMetreDoc : TMetreArt);
procedure LoadLesBlocNotes(SaContexte: TModeAff; TOBLienOle: TOB; Cledoc: R_CleDoc);
// --
procedure AcomptesVersPiece(TOBAcomptes, TOBPiece: TOB);
procedure ReliquatAcomptes(TOBAcomptes, OldAcomptes, AcomptesRel: TOB);
// Tables libres pour statistiques
function SaisieTablesLibres(TOBPiece: TOB; Bouton: Boolean = False): boolean;
procedure MAJTablesLibres(TOBPiece: TOB; NaturePieceG, Infos: string);
function RecupArgTablesLibres(TOBPiece: TOB): string;
procedure StatPieceDuplic(TOBPiece: TOB; OldNat, NewNature: string);
procedure MAVideStatPiece(TOBPiece: TOB);
// lien planning-pièce
function GerePlanningPiece(NaturePiece: string): Boolean;
function QuelPrixBase(NaturePiece, Depot: string; TOBA, TOBL: TOB): double;
// Modif BTP
procedure AddPourcent(TOBPiece, TOBPourcent: TOB; Ligne: integer);
function DansChampsPourcent(Ligne: integer; TOBPOurcent: TOB): boolean;
procedure AffecteEcartPrixMarche(XX : TForm; TOBPIece, TOBTiers, TOBAffaire, TOBArticles: TOB; Ecart: double; Article: string; IndMaxMontant, IndDepart, IndFin,
  EtenduApplication: integer);
procedure AffecteLeDepotParag (FF: TForm; Ligne : integer; TOBPiece : TOB);
procedure AppliqueChangeTaxePara(TOBPIece, TOBOuvrage: TOB; ligne: integer; Taxe1,Taxe2,Taxe3,Taxe4,Taxe5: string; Niveau: integer);
procedure AppliqueChangeDateLivPara(TOBPIece, TOBOuvrage: TOB; Ligne: integer; TheCurDate: TDateTime; Niveau: integer);
function OkLigneCalculMarche(TOBLig: TOB; IndtypeLigne: integer): boolean;
procedure BeforeReCalculPR (TOBPiece,TOBOuvrage : TOB;EnHt : boolean;Indepart,IndFin : integer; InclusSTinFG : boolean=false);
procedure AfterCalculPR (TOBPiece,TOBOuvrage : TOB;EnHt : boolean;Indepart,IndFin : integer;CalculPv : boolean=true);
function BeforeReCalculPrixMarche(TOBPiece,TOBOuvrage: TOB; ModeGestion, IndDepart, Indfin: Integer;InclusStInFg : boolean=false): boolean;
// Modif BTP
procedure CalculValoAvanc(TOBPiece: TOB; Indice: integer; DEV: Rdevise; ModifAvanc : boolean=false); overload;
procedure CalculValoAvanc(TOBPiece, TOBL: TOB; DEV: Rdevise;ModifAvanc : boolean = false); overload;
procedure CalculAvancementPiece(TobPiece, TOBL: TOB; DEv: Rdevise; ModifAvanc : boolean = false);
function AcompteRegle(TOBPiece: TOB): boolean;
function ISAcompteSoldePartiel(TOBPiece: TOB): boolean;
procedure InitialiseLigne(TOBL: TOB; ARow: integer; NewQte: double);
// ---
// Modif MODE pour le multi-dépôt (CT)
function ListeDepotParEtablissement(etab: string): string;
function ListeEtablissementParDepot(depot: string): string;
function Etabl_du_depot(depot: string): string;
procedure TiersFraisVersPiedPort(TobPiece, TobTiers, TobPL: TOB);
procedure ChargeTobPiedPort(TobPiedPort, TobPort, TobPiece: TOB);
function RecherchePort(CodePort: string; TobPort: Tob): boolean;
procedure RazFrais(TobPL: Tob);
procedure RecalculPiedPort(Action: TActionFiche; TOBPiece, TOBPiedPort,TOBBases: Tob; MontantbaseFixe : double=0);
function PivotToSaisie(TheDEV: RDevise; Base, Taux, Quotite: Double; Decimale: integer): double;
procedure CalculMontantsPiedPort(TOBPiece, TOBPL,TOBBases: TOB);
procedure CalculBasePiedPort(TOBPL: TOB; EnHT: boolean);

// Recherche des prix d'achat
function GetPrixAchat(NaturePiece, CodeArt: string; TOBArt, TOBSto, TOBCat: TOB; FactureHT: boolean; RatioStock, RatioDoc: double): double;

// MODIF VARIANTE
function IsCommentaire (TOBL : TOB) : boolean;
function IsArticle (TOBL : TOB) : boolean;
function IsDebutParagraphe (TOBL : TOB) : boolean; overload;
function IsDebutParagraphe (TOBL : TOB ; niveau : integer) : boolean; overload;
function IsFinParagraphe (TOBL : TOB) : boolean;  overload;
function IsFinParagraphe (TOBL : TOB ; niveau : integer) : boolean; overload;
function IsParagraphe (TOBL : TOB) : boolean; overload;
function IsParagraphe (TOBL : TOB ; niveau : integer) : boolean; overload;
function TestDepotOblig : boolean;
function IsReferenceLivr (TOBL : TOB) : boolean;
function IsNewLigne (TOBL : TOB) : boolean;
function IsBLobVide (FF : TForm;TOBL : TOB; Champs : string) : boolean;
function IsFromExcel (TOBL : TOB) : boolean;
//
function IsDebutParagrapheStd (TOBL : TOB) : boolean; overload;
function IsDebutParagrapheStd (TOBL : TOB ; niveau : integer) : boolean; overload;
function IsFinParagrapheStd (TOBL : TOB) : boolean;  overload;
function IsFinParagrapheStd (TOBL : TOB ; niveau : integer) : boolean; overload;
function IsParagrapheStd (TOBL : TOB) : boolean; overload;
function IsParagrapheStd (TOBL : TOB ; niveau : integer) : boolean; overload;
function IsSousTotal (TOBL : TOB) : boolean;
function IsSousDetail (TOBL : TOB) : boolean;
function IsLignefacture (TOBL : TOB) : boolean;
//
{$IFDEF BTP}
procedure ReactualisePr (TOBA : TOB);
procedure ReactualisePv (TOBA : TOB);
procedure ReactualisePrPv (TOBA : TOB);
function IsExisteFraisChantier (TOBPiece : TOB) : boolean;
function ExisteCurrentFraisChantier (TOBporcs : TOB) : boolean;
{$ENDIF}
function PieceFacturation : string;
function IsTransformable(NewNature : string) : boolean;
procedure positionneCoefMarge (TOBL : TOB);
function ExisteAncienFraisChantier (var Cledoc : r_cledoc) : boolean;
function ExisteFraisChantier (TOBpiece : TOB; var cledoc : r_cledoc) : Boolean;
function IsprestationST(LaTOb : TOB) : boolean;
procedure ValideLesBases(TOBPiece,TobBases,TOBBasesL : TOb);
function GetEtatSolde(TobLigne: Tob): T_EtatSoldeLigne; overload;
function GetEtatSolde(const CleDoc: R_CleDoc): T_EtatSoldeLigne; overload;
procedure GPDoActionAfterUpdatePiece(const CleDoc: R_CleDoc); overload;
function GetNumSituation (TOBPiece : TOB) : integer;
procedure LoadLesSsTraitants(CleDoc: R_CleDoc; TOBSsTraitant : TOB);
procedure GetSommeReglement(TOBAcomptes: TOB; var XP, XD: Double; Fournisseur : string=''; All : Boolean=false);
procedure RecupTOBArticle (TOBPiece,TOBArticles,TOBAFormule,TOBConds : TOB; CledocAffaire,Cledoc : R_cledoc; VenteAchat : string);
function IsPieceAutoriseRemligne(TOBpiece : TOB) : boolean;
function SetSituationVivante (TOBPiece : TOB; VIVANTE : string) : Integer;
procedure AffecteLeDomaineParag (FF: TForm; Ligne : integer; TOBPiece : TOB);
function IsLigneFromCentralis ( TOBL : TOB) : boolean;
function IsExisteFraisRepartit (Cledoc : R_Cledoc) : boolean; overload;
function IsExisteFraisRepartit (TOBPorcs : TOB) : boolean; overload;
function CalculeQteSais (TOBL : TOB) : double;
function CalculeQteFact (TOBL : TOB) : double;
function ISDocumentChiffrage(TOBL : TOB) : boolean; overload;
function ISDocumentChiffrage(NaturePiece : string) : boolean; overload;
procedure InitConditionnements (TOBL : TOB);


implementation

uses FactComm, FactSpec, Facture, TiersUtil, FactCpta, NomenUtil, FactNomen, DepotUtil , FactRg,
  uTofPiedPort,
  utoflisteInv,
  {$IFDEF AFFAIRE}
  AffaireUtil,
  {$ENDIF}
  {$IFDEF BTP}
  FactAcompte, FactOuvrage, UtilSuggestion, CalcOleGenericBTP,BTPUtil,
  {$ENDIF}
  FactVariante,
  Tarifs,
  FactTarifs,
  StockUtil,
  wCommuns,
  Ucotraitance,
  PiecesRecalculs,
  FactRetenues ;


function IsLigneFromCentralis ( TOBL : TOB) : boolean;
begin
  result := false;
  if (TOBL.GetString('GL_TYPELIGNE')='ART') and (TOBL.getInteger('GL_LIGNELIEE')<>0) then result := true;
end;
  
{$IFDEF BTP}
procedure ReactualisePr (TOBA : TOB);
var DPR     : Double;
    MTPAF   : Double;
    PMAP    : Double;
    CoefFG  : double;
begin

  CoefFG  := TOBA.getvalue('GA_COEFFG');
  PMAP    := TOBA.getvalue('GA_PMAP');
  MTPAF   := TOBA.getvalue('GA_PAHT');
  DPR     := TOBA.getvalue('GA_DPR');

  if (CoefFG <> 0) and
  	 ((TOBA.getValue('GA_DPRAUTO')='X') OR (TOBA.GetValue('GA_DPR')=0)) and
     (TOBA.GetValue('GA_CALCPRIXPR')='PAA') then
  begin
		if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
    begin
      if PMAP <> 0 then
      begin
        DPR := Arrondi(PMAP * CoefFG, V_PGI.okdecP);
      	TOBA.PutValue('GA_DPR',DPR);
      end;
    end
    else
    begin
      DPR := Arrondi(MTPAF*CoefFG, V_PGI.okdecP);
      TOBA.PutValue('GA_DPR',DPR);
    end;
  end;

end;

procedure ReactualisePv (TOBA : TOB);
begin
  if (TOBA.getvalue('GA_COEFCALCHT') <> 0) and
  	 ((TOBA.getValue('GA_CALCAUTOHT')='X') OR (TOBA.getValue('GA_PVHT')=0)) and
     (TOBA.GetValue('GA_CALCPRIXHT')='DPR') then
  begin
    TOBA.PutValue('GA_PVHT',Arrondi(TOBA.getvalue('GA_DPR')*TOBA.getvalue('GA_COEFCALCHT'),V_PGI.okdecP));
  end;
end;

procedure ReactualisePrPv (TOBA : TOB);
begin
	ReactualisePr (TOBA);
  ReactualisePv(TOBA);
end;

function IsExisteFraisRepartit (Cledoc : R_Cledoc) : boolean; overload;
var QQ : TQuery;
    TOBPort : TOB;
    Indice : integer;
begin
	result := false;
  TOBPort := TOB.Create ('LES PORTS', nil,-1);
	QQ := OpenSql ('SELECT * FROM PIEDPORT WHERE '+WherePiece(cledoc,ttdPorc,false),true,-1, '', True);
  TRY
    if not QQ.eof then
    begin
      TOBPort.loadDetailDb ('PIEDPORT','','',QQ,false);
      for Indice := 0 to TOBport.detail.count -1 do
      begin
        if TOBPOrt.detail[Indice].GetValue('GPT_FRAISREPARTIS') = 'X' then
        Begin
          result := true;
          break;
        END;
      end;
    end;
  finally
  	Ferme (QQ);
    TOBPOrt.free;
  end;
end;


function IsExisteFraisRepartit (TOBPorcs : TOB) : boolean; overload;
var
    Indice : integer;
begin
	result := false;
  for Indice := 0 to TOBPorcs.detail.count -1 do
  begin
    if TOBPorcs.detail[Indice].GetValue('GPT_FRAISREPARTIS') = 'X' then
    Begin
      result := true;
      break;
    END;
  end;
end;


function FindPiecePrec (TOBPiece : TOB) : string;
var Indice : integer;
begin
	result := '';
  if TOBPiece.GetValue('GP_NATUREPIECEG') = 'DBT' Then
  begin
    for indice := 0 to tobpiece.detail.count -1 do
    begin
      if TOBPiece.detail[Indice].getValue('GL_PIECEORIGINE') <> '' then
      begin
        result := TOBPiece.detail[Indice].getValue('GL_PIECEORIGINE');
      end;
    end;
  end else
  begin
    for indice := 0 to tobpiece.detail.count -1 do
    begin
      if TOBPiece.detail[Indice].getValue('GL_PIECEPRECEDENTE') <> '' then
      begin
        result := TOBPiece.detail[Indice].getValue('GL_PIECEPRECEDENTE');
      end;
    end;
  end;
end;


function ExisteAncienFraisChantier (var Cledoc : r_cledoc) : boolean;
begin
  // ancien
  cledoc := cledoc;
  cledoc.naturepiece := 'FRC';

  result := (ExistePiece(CleDoc));
end;

function ExisteFraisChantier (TOBpiece : TOB; var cledoc : r_cledoc) : Boolean;
var ThePiecePrec : string;
    TobOldPiece : TOB;
    naturePiece : string;
    QQ : TQuery;
begin
	result := false;
  NaturePiece := TOBPIece.getValue('GP_NATUREPIECEG');
  if Pos(NaturePiece,'ETU;DBT;BCE;') < 0 then exit;

  // verif sur piece courante
  if TOBPiece.getValue ('GP_PIECEFRAIS')='' then
  begin
    // ancien
    DecodeRefPiece(EncodeRefPiece (TOBPiece),Cledoc);
    result := ExisteAncienFraisChantier (Cledoc);
  end else
  begin
  // Nouveau Frais de chantier
    DecodeRefPiece (TOBpiece.getValue('GP_PIECEFRAIS'),CleDoc);
    result := (ExistePiece(CleDoc) );
  end;

  if result then exit;

  ThePiecePrec := findPiecePrec (TOBPiece);
  if ThePiecePrec <> '' then
  begin
    DecodeRefPiece(ThePiecePrec, Cledoc);

    TobOldPiece := TOB.Create ('PIECE',nil,-1);
    QQ := OpenSql('SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_DATEPIECE,GP_INDICEG,GP_PIECEFRAIS FROM PIECE WHERE '+WherePiece (cledoc,ttdPiece,true),true,-1, '', True);
    if (not QQ.eof) then
    begin
      TOBOldPiece.selectDb ('',QQ);
      result := ExisteFraisChantier (TOBOldPiece,cledoc);
    end;
    ferme (QQ);
    TOBOldPiece.free;
  end;
end;


function ExisteAncienFraisCh (Cledoc : r_cledoc) : boolean;
var cledocf : r_cledoc;
begin
  // ancien
(*
	FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := TOBPiece.GetValue('GP_DATEPIECE');
  CleDoc.DatePiece := TOBPiece.GetValue('GP_DATEPIECE');
  CleDoc.Souche := TOBPiece.GetValue('GP_SOUCHE');
  CleDoc.NumeroPiece := TOBPiece.GetValue('GP_NUMERO');
  CleDoc.Indice := TOBPiece.GetValue('GP_INDICEG');
*)
  cledocF := cledoc;
  cledocF.naturepiece := 'FRC';

  result := (ExistePiece(CleDocF) or (IsExisteFraisRepartit(CleDoc)));
end;

function IsExisteFraisChPiece (TOBpiece : TOB) : Boolean;
var cledocF : r_cledoc;
		ThePiecePrec : string;
    TobOldPiece : TOB;
    naturePiece : string;
    QQ : TQuery;
begin
	result := false;
  NaturePiece := TOBPIece.getValue('GP_NATUREPIECEG');
  if Pos(NaturePiece,'ETU;DBT;BCE;') < 0 then exit;

  // verif sur piece courante
  if TOBPiece.getValue ('GP_PIECEFRAIS')='' then
  begin
    // ancien
    DecodeRefPiece(EncodeRefPiece (TOBPiece),CledocF);
    result := ExisteAncienFraisCh (Cledocf);
  end else
  begin
  // Nouveau Frais de chantier
    DecodeRefPiece (TOBpiece.getValue('GP_PIECEFRAIS'),CleDocF);
    result := (ExistePiece(CleDocF) or (IsExisteFraisRepartit(CleDocf)));
  end;

  if result then exit;



  ThePiecePrec := findPiecePrec (TOBPiece);
  if ThePiecePrec <> '' then
  begin
    DecodeRefPiece(ThePiecePrec, Cledocf);

    TobOldPiece := TOB.Create ('PIECE',nil,-1);
    QQ := OpenSql('SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_DATEPIECE,GP_INDICEG,GP_PIECEFRAIS  FROM PIECE WHERE '+WherePiece (cledocf,ttdPiece,true),true,-1, '', True);
    if (not QQ.eof) then
    begin
      TOBOldPiece.selectDb ('',QQ);
      result := IsExisteFraisChPiece (TOBOldPiece);
    end;
    ferme (QQ);
    TOBOldPiece.free;
  end;
end;


function IsExisteFraisChLigne (TOBPiece : TOB) : Boolean;
var cledoc : r_cledoc;
		QQ : TQuery;
    TOBPieceM : TOB;
    ThePiecePrec : string;
begin
	result := false;
  if Pos(TOBPiece.GetValue('GL_NATUREPIECEG'),'ETU;DBT;') < 0 then exit;

  TOBPieceM := TOB.create ('PIECE',nil,-1);
  DecodeRefPiece(EncodeRefPiece (TOBPiece,0,false),cledoc);
  QQ := OpenSql ('SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_DATEPIECE,GP_INDICEG,GP_PIECEFRAIS FROM PIECE WHERE '+WherePiece (cledoc,ttdPiece,true),True,-1, '', True);  // piece courante
  if not QQ.eof then
  begin
  	TOBPieceM.selectDb ('',QQ);
    result:= IsExisteFraisChPiece (TOBPieceM);
  end;
  ferme (QQ);

  if result then exit; // on a trouve...

  if TOBPiece.getValue('GL_PIECEPRECEDENTE') <> '' then
  begin

    ThePiecePrec := TOBPiece.getValue('GL_PIECEPRECEDENTE');

    DecodeRefPiece(ThePiecePrec, Cledoc);
    if Pos(Cledoc.NaturePiece,'ETU;DBT;') < 0 then exit;

    QQ := OpenSql ('SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_DATEPIECE,GP_INDICEG,GP_PIECEFRAIS FROM PIECE WHERE '+WherePiece (cledoc,ttdPiece,true),True,-1, '', True);  // piece precedente
    if not QQ.eof then
    begin
      TOBPieceM.selectDb ('',QQ);
      result:= IsExisteFraisChPiece (TOBPieceM);
    end;
    ferme (QQ);
  end;

end;

function IsExisteFraisChantier (TOBPiece : TOB) : boolean;
begin
	result := false;
  if TOBPiece.NomTable = 'PIECE' then
  begin
  	result := IsExisteFraisChPiece (TOBpiece);
  end else if TOBPiece.NomTable = 'LIGNE' then
  begin
  	result := IsExisteFraisChLigne (TOBpiece);
  end;
end;

function ExisteCurrentFraisChantier (TOBporcs : TOB) : boolean;
var i : integer;
begin
  for i := 0 to TOBPorcs.Detail.Count - 1 do
  begin
    if TOBPorcs.Detail[i].GetValue('GPT_FRAISREPARTIS') = 'X' then
    BEGIN
      result := true;
      break;
    end;
  end;
end;

{$ENDIF}

{======================== Opérations sur les champs ou les zones ==================}
//AC 18/08/03 NV GESTION COMPTA DIFF
function IsComptaPce(Nature: string): boolean;
var LienP,PassP : string;
begin
  Result := false;
  if GetParamSocSecur('SO_GCDESACTIVECOMPTA',false) then exit; // SI pas de compta on sort
  LienP:=GetInfoParPiece(Nature,'GPP_TYPEPASSCPTA') ;
  PassP:=GetInfoParPiece(Nature,'GPP_TYPEECRCPTA') ; if ((PassP='') or (PassP='RIE')) then Exit ;
  Result := true;

end;

function IsComptaAcc(Nature: string): boolean;
begin
  Result := True;
  if ctxMode in V_PGI.PGIContexte then
  begin
    if (GetParamSoc('SO_GCDESACTIVECOMPTA')) or
      (GetInfoParPiece(Nature, 'GPP_TYPEPASSACC') = 'DIF') then
      Result := False;
  end;
end;
// Fin AC

function IsComptaStock(Nature: string): boolean;
var LienP, PassP: string;
begin
  Result := False;
  if Nature = '' then Exit;
  {Paramètres généraux}
  if not GetParamSoc('SO_GCINVPERM') then Exit;
  if GetParamSoc('SO_GCJALSTOCK') = '' then Exit;
  {Nature de pièce}
  PassP := GetInfoParPiece(Nature, 'GPP_TYPEECRSTOCK');
  if PassP <> 'NOR' then Exit;
  LienP := GetInfoParPiece(Nature, 'GPP_TYPEPASSCPTA');
  if LienP <> 'REE' then Exit;
  Result := True;
end;

function IsLigneDetail(TOBPiece, TOBLOC: TOB; Ligne: integer = -1): boolean;
var TOBL: TOB;
begin
  result := false;
  if (Ligne = -1) and (TOBLOC = nil) then exit;

  if TOBLOC = nil then
  begin
    TOBL := GetTOBLigne(TOBPIECE, Ligne);
  end else TOBL := TOBLOC;
  if TOBL = nil then exit; // JT eQualité 10974
  // Modif VARIANTE
(*  if (TOBL.getValue('GL_TYPELIGNE') = 'COM')  and *)
    if IsCommentaire (TOBL) and
      (TOBL.GetVAlue('GL_INDICENOMEN') > 0) then result := true;
end;

function IsLigneDetailMode(TOBLOC: TOB): boolean;
begin
  result := false;
  if (TOBLOC.getValue('GL_TYPELIGNE') = 'ART') and (TOBLOC.GetVAlue('GL_TYPEDIM') = 'DIM') then result := true;
end;

{========================== Recherches et chargements =============================}
{=================================== AFFAIRES ================================}

procedure ChangeTzAffaireSaisie(CodeTiers: string);
var St, sWhere: string;
  ii: integer;
begin
  if CodeTiers <> '' then sWhere := 'AFF_TIERS="' + CodeTiers + '"' else sWhere := '';
  St := 'AFLAFFAIRESAISIE';
  ii := TTToNum(St);
  if ii > 0 then V_PGI.DECombos[ii].Where := sWhere;
end;

// PL le 07/11/02 : ajout du paramètre contexte d'appel

function GetAffaireEntete(TAffaire, TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit; bProposition, bChangeStatut, bModele, bChangeTiers,BCreatAff:
  Boolean; sContexte: string = ''; bReduit: boolean = true; bOuvreSiUn: boolean = true; bItDatesMax : boolean = false): boolean;
var St: string;
begin
  if TAffaire <> nil then St := TAffaire.Text else St := '';
  // PL le 07/11/02 : ajout du paramètre contexte d'appel
  //mcd 25/09/03 : ajout option Bcreat aff: n epermet pas bouton création
  GetAffaireMul(St, nil,TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers, bProposition, bChangeStatut, bModele, bChangetiers,BcreatAff, bReduit, sContexte, bOuvreSiUn, bItDatesMax);
  Result := (St <> TAffaire.Text);
  TAffaire.Text := St;
end;

// PL le 07/11/02 : ajout du paramètre contexte d'appel

function GetAffaireEnteteSt(Taffaire0,TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit; var CodeAffaire: string; bProposition, bChangeStatut, bModele,
  bChangeTiers, BCreatAff: Boolean; sContexte: string = ''; bReduit: boolean = false; bOuvreSiUn: boolean = true; bItDatesMax : boolean = false; EtatAffaire : string= ''): boolean;
var St: string;
begin
  result := false;
  St := CodeAffaire;
  // PL le 07/11/02 : ajout du paramètre contexte d'appel
  if GetAffaireMul(CodeAffaire,Taffaire0, TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers, bProposition, bChangeStatut, bModele, bChangetiers,
  					BCreatAff, bReduit, sContexte,bOuvreSiUn, bItDatesMax,EtatAffaire) then
  begin
  	Result := (St <> CodeAffaire);
  end;
end;

function GetAffaireLigne(GS: THGrid; Tiers, NaturePieceG: string; bAffReduit: Boolean): Boolean;
var Aff, StChamps, StArgument, stAff: string;
begin
  Result := False;
  Aff := GS.Cells[GS.Col, GS.Row];
  Stchamps := 'AFF_STATUTAFFAIRE=AFF';
  Stargument := 'NOCHANGESTATUT';
  if GetInfoParPiece(NaturePieceG, 'GPP_VENTEACHAT') = 'VEN' then
  begin
    Stargument := Stargument + ';NOCHANGETIERS';
    if Tiers <> '' then StChamps := StChamps + ';AFF_TIERS=' + Tiers;
  end;
  Aff := AGLLanceFiche('AFF', 'AFFAIRERECH_MUL', StChamps, '', StArgument);
  if Aff <> '' then
  begin
    Result := True;
    stAff := ReadTokenSt(Aff);
    {$IFDEF BTP}
    {$ELSE IF AFFAIRE}
    if bAffreduit then stAff := CodeAffaireAffiche(stAff);
    {$ENDIF}
    GS.Cells[GS.Col, GS.Row] := stAff;
  end;
end;

// PL le 07/11/02 : ajout du paramètre contexte d'appel

function GetAffaireMul(var CodeAff: string; TAffaire0, TAffaire1, TAffaire2, TAffaire3, TAvenant, Tiers: THCritMaskEdit;
  bProposition, bChangeStatut, bModele, bChangetiers,BCreatAff, bReduit: Boolean;
  sContexte: string = ''; bOuvreSiUn: boolean = true; bIntervalleMaxi : boolean = false; EtatAffaire : string =''): boolean;
var StChamps, tmp, StArgument: string;
    TheTypeAffaire : string;
begin
  StChamps := '';

  // chargement du type d'affaire (I=Contrat, W=Appel, P=Appel d'offre, A=Chantier)
  if Taffaire0 = nil then
    //FV1 : 07/02/2014 - FS#825 - DELABOUDINIERE : en transfert inter-chantiers, ajouter la recherche dans les appels et contrats
    //TheTypeAffaire := 'A'
    TheTypeAffaire := ''
  else
    TheTypeAffaire := Taffaire0.text;

  if bProposition then TheTypeAffaire := 'P';

  if TheTypeAffaire = 'A' then
  begin
    StChamps := 'AFF_STATUTAFFAIRE=AFF';
    if not bChangeStatut then
      StArgument := 'STATUT:AFF;NOCHANGESTATUT'
    else
      Stargument := 'STATUT:';
  end
  else if TheTypeAffaire = 'I' then
  begin
    StChamps := 'AFF_STATUTAFFAIRE=INT';
    if not bChangeStatut then
      StArgument := 'STATUT:INT;NOCHANGESTATUT'
    else
      stArgument := 'STATUT:';
  end
  else if TheTypeAffaire = 'W' then
  begin
    StChamps := 'AFF_STATUTAFFAIRE=APP';
    if not bChangeStatut then
      StArgument := 'STATUT:APP;NOCHANGESTATUT'
    else
      stArgument := 'STATUT:';
  end
	else if TheTypeAffaire = 'P' then
  Begin
    StChamps := 'AFF_STATUTAFFAIRE=PRO';
    if not bChangeStatut then
      StArgument := 'STATUT:PRO;NOCHANGESTATUT'
    else
      stArgument := 'STATUT:';
  end
  else
  //FV1 : 07/02/2014 - FS#825 - DELABOUDINIERE : en transfert inter-chantiers, ajouter la recherche dans les appels et contrats
  begin
    Stchamps := 'AFF_STATUTAFFAIRE=;';
    //if not bChangeStatut then StArgument := 'STATUT:AFF' else StArgument := 'STATUT:';
    StArgument := 'TOUS=OUI';
  end;

  If Not (BcreatAff) then StArgument := Stargument +';ACTION=CONSULTATION' ;//mcd 25/09/03 pour acceptation proposision (pas créat sur mul rech misison origine)
  //
  //if not (bChangeStatut) then Stargument := Stargument + ';NOCHANGESTATUT';
  //
  if bModele then Stargument := Stargument + ';MODELEONLY';

  If Assigned(Tiers) then
  begin
    if not (bChangeTiers) and (Tiers.Text <> '') then Stargument := Stargument + ';NOCHANGETIERS';
  end;

  if ((TAffaire0 <> nil) and (TAFFAIRE0.text<>'')) or (TAFFAIRE1.text<>'') or (TAFFAIRE2.text<>'') or (TAFFAIRE3.text<>'') then
  begin
    if TheTypeAffaire <> '' then StChamps := StChamps + ';AFF_AFFAIRE0=' + TheTypeAffaire;
    if TAFFAIRE1.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE1=' + TAFFAIRE1.Text;
    if TAFFAIRE2.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE2=' + TAFFAIRE2.Text;
    if TAFFAIRE3.text <> '' then StChamps := StChamps + ';AFF_AFFAIRE3=' + TAFFAIRE3.Text;
    if (TAvenant <> nil) then
       if (TAvenant.Text <> '00') or (TAvenant.Text <> '') then
          StChamps := StChamps + ';AFF_AVENANT=' + TAvenant.Text;
  end;

  if Assigned(tiers) then
  begin
    if Tiers.Text <> '' then
    begin
      StChamps := StChamps + ';AFF_TIERS=' + Tiers.Text;
      Stargument := Stargument + ';NOFILTRE'; // mcd 02/07/02
    end;
  end;

  {$IFDEF BTP}
  Stargument := Stargument + ';ACTION=RECH';

  if EtatAffaire <> '' then
  Begin
    if Pos(',',EtatAffaire)= 0 then
    begin
      StChamps := StChamps +';AFF_ETATAFFAIRE='+EtatAffaire;
    end else
    begin
      stargument := stargument + ';ETAT=' + EtatAffaire;
    end;
    StArgument := Stargument + ';NOAFFETAT'
  end;

  if GetParamSoc('SO_BTAFFAIRERECHLIEU') then
     tmp := AGLLanceFiche('BTP', 'BTAFFAIREINT_MUL', StChamps, '', StArgument)
  else
   	 tmp := AGLLanceFiche('BTP', 'BTAFFAIRE_MUL', StChamps, '', StArgument);
  {$ELSE}
  //MCD 15/01/02 ajout bReduit
  if bReduit then if VH_GC.UserInvite then StArgument := StArgument + ';REDUIT';

  Stargument := Stargument + ';ADMINGRAYED'; // mcd 10/10/02
  // PL le 07/11/02 : ajout du paramètre contexte d'appel
  if (sContexte <> '') then
    Stargument := Stargument + ';CTX=' + sContexte;
  /////////////////////////

  if (bOuvreSiUn = false) then
    Stargument := Stargument + ';PASSIUN'; // PL le 25/06/03

  if (bIntervalleMaxi = true) then
    Stargument := Stargument + ';ITDATESMAX'; // PL le 09/07/03

  tmp := AGLLanceFiche('AFF', 'AFFAIRERECH_MUL', StChamps, '', StArgument);
  {$ENDIF}

  if tmp <> '' then
  begin
    Result := True;
    CodeAff := ReadTokenSt(tmp);
    if tiers <> nil then Tiers.Text := tmp;
  end
  else Result := False;
end;

function GetAffaireLookUp(GS: THGrid; TitreSel, CodeTiers: string): boolean;
var sWhere: string;
begin
  if CodeTiers <> '' then sWhere := 'AFF_TIERS="' + CodeTiers + '"' else sWhere := '';
  Result := LookupList(GS, TitreSel, 'AFFAIRE', 'AFF_AFFAIRE', 'AFF_LIBELLE', sWhere, 'AFF_AFFAIRE', True, 5);
end;

procedure ZoomAffaire(CodeAff: string);
var Statut: string;
begin
  if CodeAff[1] = 'P' then Statut := 'PRO' else Statut := 'AFF';
  V_PGI.DispatchTT(5, taModif, CodeAff, '', 'MONOFICHE;STATUT:' + Statut); // mcd 17/12/02 mise tamodif au lieu taconsult  +monofiche
end;

function RemplirTOBAffaire(CodeAffaire: string; TobAffaire: Tob): boolean;
var Q: TQuery;
begin
  Result := True;
  if CodeAffaire = '' then TOBAffaire.InitValeurs(False)
  else if CodeAffaire <> TOBAffaire.GetValue('AFF_AFFAIRE') then
  begin
    Q := OpenSQL('SELECT * FROM AFFAIRE WHERE AFF_AFFAIRE="' + CodeAffaire + '"', True,-1, '', True);
    if (not Q.EOF) then
      Result := TOBAffaire.SelectDB('', Q)
    else Result := False;
    Ferme(Q);
  end;
end;

procedure AffaireVersPiece(TOBPiece, TOBAffaire: TOB);
var
  {$IFDEF AFFAIRE}
  TobPieceAff: TOB;
  i: integer;
  {$ENDIF}
  CodeAff: string;
  VenteAchat : string;
begin
  if (TobPiece = nil) or (TobAffaire = nil) then Exit;
  VenteAchat := GetInfoParPiece(TOBPiece.getValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  TOBPiece.PutValue('GP_APPORTEUR', TOBAffaire.GetValue('AFF_APPORTEUR'));
  TOBPiece.PutValue('GP_LIBREAFF1', TOBAffaire.GetValue('AFF_LIBREAFF1'));
  TOBPiece.PutValue('GP_LIBREAFF2', TOBAffaire.GetValue('AFF_LIBREAFF2'));
  TOBPiece.PutValue('GP_LIBREAFF3', TOBAffaire.GetValue('AFF_LIBREAFF3'));
  //TOBPiece.PutValue('GP_RESSOURCE',TOBAffaire.GetValue('AFF_RESPONSABLE')) ;
  // gm oter le 17/7/02 car incompatible avec facture sur activité
  if TobAffaire.GetValue('AFF_REFEXTERNE') <> '' then TOBPiece.PutValue('GP_REFEXTERNE', TOBAffaire.GetValue('AFF_REFEXTERNE'));
  CodeAff := TobAffaire.GetValue('AFF_AFFAIRE');
  {$IFDEF AFFAIRE}
  if (CodeAff <> '') and (ctxAffaire in V_PGI.PGIContexte) then
  begin
    TobPieceAff := Tob.create('PIECE', nil, -1);
    if GetTobPieceAffaire(CodeAff, TobAffaire.GetValue('AFF_STATUTAFFAIRE'), TobPieceAff) then
    begin
      for i := 1 to 9 do TOBPiece.PutValue('GP_LIBRETIERS' + IntToStr(i), TOBPieceAff.GetValue('GP_LIBRETIERS' + IntToStr(i)));
      TOBPiece.PutValue('GP_LIBRETIERSA', TOBPieceAff.GetValue('GP_LIBRETIERSA'));
      for i := 1 to 3 do TOBPiece.PutValue('GP_TIERSSAL' + IntToStr(i), TOBPieceAff.GetValue('GP_TIERSSAL' + IntToStr(i)));
    end;
    TobPieceAff.Free;
  end;
  {$ENDIF}
  if VenteAchat= 'VEN' then
  begin
    if (TOBAffaire.GetBoolean ('AFF_SSTRAITANCE')) and (TOBAffaire.GetValue('AFF_DATESSTRAIT')>StrToDate('01/01/2014')) then TOBPIece.SetBoolean('GP_AUTOLIQUID',true);
  end;
end;

procedure AffaireVersLigne(TOBPiece, TOBL, TOBAffaire: TOB);
begin
  if not (ctxaffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte) then Exit;
  if TOBAffaire = nil then Exit;
  if TOBL.FieldExists('COMPTAAFFAIRE') then
    TOBL.PutValue('COMPTAAFFAIRE', TobAffaire.GetValue('AFF_COMPTAAFFAIRE'));
end;

procedure MajAnaAffaire(TOBPiece, TOBCataLogu, TOBArticles, TOBCpta, TOBTiers, TOBAnaP: TOB; ARow: integer);
var TOBL, TOBC, TobA, TobCata: TOB;
begin
  TOBC := nil;
  TobA := nil;
  TOBCata := nil;
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL.GetValue('GL_TYPEREF') = 'CAT' then
  begin
    TOBCata := FindTOBCataRow(TOBPiece, TOBCataLogu, ARow);
  end else
  begin
    TOBA := FindTOBArtRow(TOBPiece, TOBArticles, ARow);
  end;
  if (TOBCata = nil) or (TobA = nil) then TOBC := ChargeAjouteCompta(TOBCpta, TOBPiece, TOBL, TOBA, TOBTiers, TOBCata, nil, True);
  if (TOBC <> nil) then PreVentileLigne(TOBC, TOBAnaP, nil, TOBL);
end;

{============================= Chr ==================================}
{$IFDEF CHR}

procedure RemplirTOBHrdossier(Hrdossier: string; TobHrdossier: Tob);
var Q: TQuery;
begin
  if Hrdossier = '' then TOBHrdossier.InitValeurs(False)
  else if (copy(Hrdossier, 12, 1) = 'P') then
  begin
    if Hrdossier <> TOBHrdossier.GetValue('HDR_DOSRES') then
      // modif CHR : RH le 18/11/02
      Q := OpenSQL('SELECT HDC_HRDOSSIER,HDC_LIBELLE1,HDC_TIERSAPPORT1,HDC_SEGMENT1,HDC_FAMREG,HDC_TIERSPAYEUR,HDC_TIERS,'
        + 'HDC_LIBELLEDEPART1,HDC_LIBELLEDEPART2,HDC_TYPDOS,'
        + 'HDR_HRDOSSIER,HDR_DOSRES,HDR_DATEARRIVEE, HDR_DATEDEPART,HDR_TIERS,HDR_RESSOURCE,HDR_NBPERSONNE1,HDR_TYPRESTARIF,'
        + 'HDR_CIVILITE1,HDR_NOMCLIENT1,HDR_HRTARIF,HDR_NOMBRERES,HDR_HRETATINTERNE FROM HRDOSSIER LEFT JOIN HRDOSRES ON HDC_HRDOSSIER =HDR_HRDOSSIER WHERE HDR_DOSRES="' +
        HRDossier + '"', True,-1, '', True);
  end else
  begin
    if Hrdossier <> TOBHrdossier.GetValue('HDC_HRDOSSIER') then
      Q := OpenSQL('SELECT * FROM HRDOSSIER WHERE HDC_HRDOSSIER="' + Hrdossier + '"', True,-1, '', True);
  end;
  if (Q <> Nil) and (not Q.EOF) then
    TobHrdossier.SelectDB('', Q);
  Ferme(Q);
end;

procedure EcheVersPiece(TOBEches, TOBPiece: TOB);
var TOTD, TOTP: Double;
begin
  GetSommeReglements(TOBEches, TOTP, TOTD);
  TOBPiece.PutValue('GP_ACOMPTE', TOTP);
  TOBPiece.PutValue('GP_ACOMPTEDEV', TOTD);
end;

procedure GetSommeReglements(TOBEches: TOB; var XP, XD: Double);
var i: integer;
  TOBECH: TOB;
begin
  XP := 0;
  XD := 0;
  if TOBEches = nil then Exit;
  for i := 0 to TOBEches.Detail.Count - 1 do
  begin
    TOBECH := TOBEches.Detail[i];
    XD := Arrondi(XD + TOBECH.GetValue('GPE_MONTANTDEV'), 6);
    XP := Arrondi(XP + TOBECH.GetValue('GPE_MONTANTECHE'), 6);
  end;
end;
{$ENDIF}

{================================== TOBS ======================================}

procedure MajFromCleDoc(LaTOB: TOB; CleDoc: R_CleDoc);
var Pref: string;
begin
  if LaTOB = nil then Exit;
  if      LaTOB.NomTable = 'LIGNE'      then Pref := 'GL'
  else if LaTOB.NomTable = 'LIGNOMEN'      then Pref := 'GLN'
  else if LaTOB.NomTable = 'BLIGNEMETRE'      then Pref := 'BLM'
  else if LaTOB.NomTable = 'PIECE'      then Pref := 'GP'
  else if LaTOB.NomTable = 'PIECEADRESSE'  then Pref := 'GPA'
  else if LaTOB.NomTable = 'PIECERG'      then Pref := 'PRG'
  else if LaTOB.NomTable = 'LIGNEBASE'   then Pref := 'BLB'
  else if LaTOB.NomTable = 'LIGNEOUV'      then Pref := 'BLO'
  else if LaTOB.NomTable = 'LIGNEOUVPLAT'  then Pref := 'BOP'
  else if LaTOB.NomTable = 'PIEDBASE'   then Pref := 'GPB'
  else if LaTOB.NomTable = 'PIEDBASERG'   then Pref := 'PBR'
  else if LaTOB.NomTable = 'PIEDECHE'   then Pref := 'GPE'
  else if LaTOB.NomTable = 'PIECETRAIT'    then Pref := 'BPE'
  else if LaTOB.NomTable = 'PIECEINTERV'   then Pref := 'BPI'
  else if LaTOB.NomTable = 'PIEDPORT'      then Pref := 'GPT'
  else if LaTOB.NomTable = 'LIGNETARIF' then Pref := 'GLT'
  else if LaTOB.NomTable = 'PIECEADRESSE'  then Pref := 'GPA'
  else if LaTOB.NomTable = 'LIENSOLE'      then Pref := 'LO'
  else if LaTOB.NomTable = 'BREVISIONS' then Pref := 'BRV'
  else exit;

  LaTOB.PutValue(Pref + '_NATUREPIECEG', CleDoc.NaturePiece);
  LaTOB.PutValue(Pref + '_SOUCHE', CleDoc.Souche);
  LaTOB.PutValue(Pref + '_NUMERO', CleDoc.NumeroPiece);
  LaTOB.PutValue(Pref + '_INDICEG', CleDoc.Indice);
  if LaTOB.FieldExists(pref+'_DATEPIECE') then
  begin
    LaTOB.PutValue(Pref + '_DATEPIECE', CleDoc.DatePiece);
  end;

end;

function CreerQuelDepot(TOBPIECE: TOB): string;
begin
  {if VH_GC.GCMultiDepots then Result:='' // JS, DBR : Depot unique chargé
    else } if TOBPiece.GetValue('GP_NATUREPIECEG') = 'TEM' then Result := '"' + TOBPiece.GetValue('GP_DEPOT') + '","' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'
  else if TOBPiece.GetValue('GP_NATUREPIECEG') = 'TRE' then Result := '"' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'
  else Result := '"' + TOBPiece.GetValue('GP_DEPOT') + '"';
end;

procedure LoadTOBDispo(TOBA: TOB; Force: boolean; QuelDepot: string = '');
var QQ: TQuery;
  i: integer;
  RefUnique, Depot, StWhereDispo, StWhereDispoLot: string;
  TOBDesLots, TOBDL, TOBDepot: TOB;
begin
  if TOBA = nil then Exit;
  if ((TOBA.Detail.Count > 0) and (not Force)) then Exit;
  {Chargement des dispos}
  TOBA.ClearDetail;
  RefUnique := TOBA.GetValue('GA_ARTICLE');
  if QuelDepot = '' then StWhereDispo := '' else StWhereDispo := ' AND GQ_DEPOT IN (' + QuelDepot + ')';
  QQ := OpenSQL('Select * from DISPO WHERE GQ_ARTICLE="' + RefUnique + '" AND GQ_CLOTURE="-"' + StWhereDispo, True,-1, '', True);
  if not QQ.EOF then
  begin
    TOBA.LoadDetailDB('DISPO', '', '', QQ, False);
    //$IFNDEF EAGLCLIENT fonctionne en CWAS 07/03/2003
    //DispoChampSupp(TOBA); déplacé
    //$ENDIF
  end;
  Ferme(QQ);
  {Chargement des lots}
  if TOBA.GetValue('GA_LOT') = 'X' then
  begin
    TOBDesLots := TOB.Create('', nil, -1);
    if QuelDepot = '' then StWhereDispoLot := '' else StWhereDispoLot := ' AND GQL_DEPOT IN (' + QuelDepot + ')';
    QQ := OpenSQL('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="' + RefUnique + '"' + StWhereDispoLot, True,-1, '', True);
    if not QQ.EOF then TOBDesLots.LoadDetailDB('DISPOLOT', '', '', QQ, False);
    Ferme(QQ);
    for i := TOBDesLots.Detail.Count - 1 downto 0 do
    begin
      TOBDL := TOBDesLots.Detail[i];
      Depot := TOBDL.GetValue('GQL_DEPOT');
      TOBDepot := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
      if TOBDepot <> nil then TOBDL.ChangeParent(TOBDepot, -1);
    end;
    TOBDesLots.Free;
  end;
  DispoChampSupp(TOBA);
end;

{================================== ARTICLES =================================}

procedure CalculePrixArticle(TOBA: TOB; Depot: string = '');
var CodeBase, CodeArrondi: string;
  BaseTarif, CoefCalc, PrixRef: double;
  //  BBI, fiche correction 10410
  TobD: TOB;
  //  BBI, fin fiche correction 10410
begin
  PrixRef := 0;
  if TOBA.GetValue('GA_CALCAUTOHT') = 'X' then
  begin
    CodeBase := TOBA.GetValue('GA_CALCPRIXHT');
    CoefCalc := Valeur(TOBA.GetValue('GA_COEFCALCHT'));
    CodeArrondi := TOBA.GetValue('GA_ARRONDIPRIX');

    if (CoefCalc <> 0) and (CodeBase <> '') then
    begin
      //  BBI, fiche correction 10410
      TobD := nil;
      if Depot <> '' then
        TobD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], True);
      if ((Copy(CodeBase, 1, 1) <> 'S') or
        (TobD = nil) or
        ((Copy(CodeBase, 1, 1) = 'S') and (Depot = ''))) then
      begin
        if Copy(CodeBase, 2, 2) = 'PA' then PrixRef := Valeur(TOBA.GetValue('GA_DPA'))
        else if Copy(CodeBase, 2, 2) = 'PR' then PrixRef := Valeur(TOBA.GetValue('GA_DPR'))
        else if CodeBase = 'PAA' then PrixRef := Valeur(TOBA.GetValue('GA_PAHT'))
        else if CodeBase = 'PRA' then PrixRef := Valeur(TOBA.GetValue('GA_PRHT'))
        else if Copy(CodeBase, 2, 2) = 'MA' then PrixRef := Valeur(TOBA.GetValue('GA_PMAP'))
        else if Copy(CodeBase, 2, 2) = 'MR' then PrixRef := Valeur(TOBA.GetValue('GA_PMRP'));
      end
      else
      begin
        if CodeBase = 'SPA' then PrixRef := Valeur(TOBD.GetValue('GQ_DPA'))
        else if CodeBase = 'SPR' then PrixRef := Valeur(TOBD.GetValue('GQ_DPR'))
        else if CodeBase = 'SMA' then PrixRef := Valeur(TOBD.GetValue('GQ_PMAP'))
        else if CodeBase = 'SMR' then PrixRef := Valeur(TOBD.GetValue('GQ_PMRP'));
      end;
      //  BBI, fin fiche correction 10410
      BaseTarif := PrixRef * CoefCalc;
      BaseTarif := ArrondirPrix(CodeArrondi, BaseTarif);
      TOBA.PutValue('GA_PVHT', BaseTarif);
    end;
  end;

  if TOBA.GetValue('GA_CALCAUTOTTC') = 'X' then
  begin
    CodeBase := TOBA.GetValue('GA_CALCPRIXTTC');
    CoefCalc := Valeur(TOBA.GetValue('GA_COEFCALCTTC'));
    CodeArrondi := TOBA.GetValue('GA_ARRONDIPRIXTTC');

    if (CoefCalc <> 0) and (CodeBase <> '') then
    begin
      if CodeBase = 'DPA' then PrixRef := Valeur(TOBA.GetValue('GA_DPA'))
      else if CodeBase = 'DPR' then PrixRef := Valeur(TOBA.GetValue('GA_DPR'))
      else if CodeBase = 'PAA' then PrixRef := Valeur(TOBA.GetValue('GA_PAHT'))
      else if CodeBase = 'PMA' then PrixRef := Valeur(TOBA.GetValue('GA_PMAP'))
      else if CodeBase = 'PMR' then PrixRef := Valeur(TOBA.GetValue('GA_PMRP'))
      else if CodeBase = 'HT' then PrixRef := Valeur(TOBA.GetValue('GA_PVHT'));
      BaseTarif := PrixRef * CoefCalc;
      BaseTarif := ArrondirPrix(CodeArrondi, BaseTarif);
      if ARRONDI(TOBA.GetValue('GA_PVTTC'), V_PGI.OkDecV) <> ARRONDI(BaseTarif, V_PGI.OkDecV) then
        TOBA.PutValue('GA_PVTTC', BaseTarif);
    end;
  end;
end;

function EtudieEtat(TOBArt: TOB; NaturePiece: string): T_RechArt;
var Etat: string;
  k: integer;
  ToutObli: boolean;
begin
  if TobArt = nil then
  begin
    Result := traAucun;
    Exit;
  end;
  Etat := TOBArt.GetValue('GA_STATUTART');
  if Etat = 'UNI' then Result := traOk else
  begin
    ToutObli := True;
    for k := 1 to 5 do if TOBArt.GetValue('GA_GRILLEDIM' + IntToStr(k)) <> '' then
        if GetInfoParPiece(NaturePiece, 'GPP_TYPEDIMOBLI' + IntToStr(k)) = '-' then ToutObli := False;
    if ToutObli then Result := traGrille else Result := traOkGene;
  end;
end;

function TrouverArticleSQL(NaturePiece, RefSais, DomainePiece, SelectFourniss: string; DatePiece: TDateTime; TOBArt, TOBTiers: TOB; RefArt : String=''): T_RechArt;
var Prio: string;
  Q,Q1: TQuery;
  i: integer;
  Bloque: boolean;
  WhereNat, CodeTiers, NatTiers, SQL: string;
  TOBGCA : TOB;
  RefUnique : string;

  procedure TrouveRefTiers(CodeArt, TypeTiers, CodeTiers, Prio : string);
  var TobRefTiers: TOB;
    RefTiers: string;
    Qry: TQuery;
    found : boolean;
  begin
    if (CodeArt = '') or (TypeTiers = '') or (CodeTiers = '') then exit;
    if not TOBArt.FieldExists('REFARTTIERS') then TOBArt.AddChampSup('REFARTTIERS', False);

    if ((NatTiers = 'CLI') or (NatTiers = 'AUD')) then
    begin
      Qry := OpenSQL('SELECT * FROM ARTICLETIERS WHERE GAT_REFTIERS="' + CodeTiers + '"' +
        'AND GAT_ARTICLE="' + CodeArt + '" ', True,-1, '', True);
      if not Qry.EOF then
      begin
        TobRefTiers := TOB.Create('ARTICLETIERS', nil, -1);
        TobRefTiers.SelectDB('', Qry);
        RefTiers := TobRefTiers.GetValue('GAT_REFARTICLE');
        TobRefTiers.Free;
      end else
      begin
        RefTiers := '';
      end;
    end else if ((NatTiers = 'FOU') or (NatTiers = 'AUC')) then
    begin

      found := false;
      if not found then
      begin
        SQL := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + CodeTiers + '" AND GCA_ARTICLE="' + CodeArt + '" ' +
          'AND GCA_DATESUP>="' + USDateTime(DatePiece) + '" AND GCA_DATEREFERENCE<="' + USDateTime(DatePiece) +'" ORDER BY GCA_DATEREFERENCE DESC';
        Qry := OpenSQL(SQL, True,-1, '', True);
        if not Q.EOF then
        begin
          TobRefTiers := TOB.Create('CATALOGU', nil, -1);
          TobRefTiers.SelectDB('', Qry);
          RefTiers := TobRefTiers.GetValue('GCA_REFERENCE');
          TobRefTiers.Free;
        end else
        begin
          RefTiers := '';
        end;
      end;
    end;
    Ferme(Qry);
    TOBArt.PutValue('REFARTTIERS', RefTiers)
  end;

begin
  TOBGCA := TOB.Create ('CATALOGU',nil,-1);
  Result := traAucun;
  if ((RefSais = '') or (NaturePiece = '') or (TOBArt = nil)) then Exit;
  if TOBTiers = nil then CodeTiers := '' else
  begin
    CodeTiers := TOBTiers.GetValue('T_TIERS');
    NatTiers := TOBTiers.GetValue('T_NATUREAUXI');
  end;
  WhereNat := FabricWhereNatArt(NaturePiece, DomainePiece, SelectFourniss);
  if WhereNat <> '' then WhereNat := ' AND ' + WhereNat;
  for i := 1 to 3 do
  begin
    Prio := GetInfoParPiece(NaturePiece, 'GPP_PRIORECHART' + IntToStr(i));
    Bloque := GetInfoParPiece(NaturePiece, 'GPP_CONTRECHART' + IntToStr(i)) = 'X';
    if Prio = 'AUC' then Break;
    if ((Prio = 'BAR') and (Length(RefSais) >= VH_GC.GCLgMinBarre)) then
    begin
      // Recherche via code barre (1er code barre)
    	Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEBARRE="'+RefSais+'"' + WhereNat,true,-1, '', True);

      //2nd code barre stocké dans la zone GA_CODEDOUANIER
      if (Q.EOF) and (ctxMode in V_PGI.PGIContexte) then
      begin
        if GetParamSoc('SO_ARTMULTICODEBARRE') then
        begin
          Ferme(Q);
          Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 		 '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEDOUANIER="'+RefSais+'"' + WhereNat,true,-1, '', True);
        end;
      end;

      if (Q.eof) and (NatTiers = 'FOU') then // recherche dans le cas d'un fournisseur sur le code barre
      begin
        SQL := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + CodeTiers + '" AND GCA_CODEBARRE="' + RefSais + '" ' +
          'AND GCA_DATESUP>="' + USDateTime(DatePiece) + '" AND GCA_DATEREFERENCE<="' + USDateTime(DatePiece) +'" ORDER BY GCA_DATEREFERENCE DESC';
        Q1:= OpenSQL(SQL, True,-1, '', True);
        if not Q1.EOF then
        begin
          TOBGCA.selectDb ('',Q1);
          if TOBGCA.getValue('GCA_ARTICLE')<> '' then
          begin
            Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 			  '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                           'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                           'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
                           'WHERE A.GA_ARTICLE="'+TOBGCA.getValue('GCA_ARTICLE')+
                           '" AND (GA_STATUTART="UNI" OR GA_STATUTART="GEN") ' + WhereNat,true,-1, '', True);
          end;
        end;
        ferme (Q1);
      end;
      if not Q.EOF then
      begin
        TOBArt.SelectDB('', Q);
        InitChampsSupArticle (TOBARt);
        TOBArt.PutValue('REFARTBARRE', RefSais);
        TrouveRefTiers(TOBArt.GetValue('GA_ARTICLE'), NatTiers, CodeTiers, Prio);
        Result := traOk;
      end;
      Ferme(Q);

      if ((Result <> traOk) and (Bloque)) then
      begin
        Result := traErreur;
        Break;
      end;
    end;
    if ((Prio = 'REF') and (CodeTiers <> '') and (Length(RefSais) >= VH_GC.GCLgMinReFTiers)) then
    begin
      // Recherche via code catalogue ou articletiers
      if ((NatTiers = 'CLI') or (NatTiers = 'AUD')) then
      begin
        Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 		'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE ' +
                       'WHERE A.GA_ARTICLE=(Select GAT_ARTICLE From ARTICLETIERS Where GAT_REFTIERS="' + CodeTiers + '" AND GAT_REFARTICLE="'+ RefSais + '") ' + WhereNat,true,-1, '', True);
        if not Q.EOF then
        begin
          TOBArt.SelectDB('', Q);
          InitChampsSupArticle (TOBART);
          TOBArt.PutValue('REFARTTIERS', RefSais);
          Result := traOk;
        end;
        Ferme(Q);
        if ((Result <> traOk) and (Bloque)) then
        begin
          Result := traErreur;
          Break;
        end;
      end;
      if ((NatTiers = 'FOU') or (NatTiers = 'AUC')) then
      begin
        SQL := 'Select A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
        			 '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
        			 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
               'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
        			 'Left Join CATALOGU CA ON A.GA_ARTICLE=CA.GCA_ARTICLE '+
               'Where GCA_TIERS="' + CodeTiers + '" AND GCA_REFERENCE="' + RefSais + '" ' +
               'AND GCA_DATESUP>="' + USDateTime(DatePiece) + '" AND GCA_DATEREFERENCE<="' + USDateTime(DatePiece) +'" ' + WhereNat +' ORDER BY GCA_DATEREFERENCE DESC';
        Q := OpenSQL(SQL, True,-1, '', True);
        if not Q.EOF then
        begin
          TOBArt.SelectDB('', Q);
          TOBArt.PutValue('REFARTTIERS', RefSais);
          Result := traOk;
        end;
        Ferme(Q);
        if ((Result <> traOk) and (Bloque)) then
        begin
          Result := traErreur;
          Break;
        end;
      end;
    end;
    if Prio = 'ART' then
    begin
      // Recherche via code unique ou générique
      if RefArt = '' then
      begin
        // Recherche via code unique ou générique
        Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 		'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
                       'WHERE A.GA_CODEARTICLE="'+RefSais+'" AND (GA_STATUTART="UNI" OR GA_STATUTART="GEN") ' + WhereNat,true,-1, '', True);
      end else
      begin
        // Recherche via code unique ou générique
        Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 		'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                       'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                       'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefART+'" '+ WhereNat,true,-1, '', True);
      end;
      if not Q.EOF then
      begin
        TOBArt.SelectDB('', Q);
        InitChampsSupArticle (TOBART);
        TrouveRefTiers(TOBArt.GetValue('GA_ARTICLE'), NatTiers, CodeTiers, prio);
        Result := EtudieEtat(TOBArt, NaturePiece);
      end;
      Ferme(Q);
    end;
    IF prio = 'BAF' then
    begin
      if (NatTiers = 'FOU') then // recherche dans le cas d'un fournisseur sur le code barre
      begin
        SQL := 'SELECT * FROM CATALOGU WHERE GCA_TIERS="' + CodeTiers + '" AND GCA_CODEBARRE="' + RefSais + '" ' +
          'AND GCA_DATESUP>="' + USDateTime(DatePiece) + '" AND GCA_DATEREFERENCE<="' + USDateTime(DatePiece) +'" ORDER BY GCA_DATEREFERENCE DESC';
        Q1:= OpenSQL(SQL, True,-1, '', True);
        if not Q1.EOF then
        begin
          TOBGCA.selectDb ('',Q1);
          if TOBGCA.getValue('GCA_ARTICLE')<> '' then
          begin
            Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE'+
      						 				'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                           'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                           'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
                           'WHERE A.GA_ARTICLE="'+TOBGCA.getValue('GCA_ARTICLE')+
                           '" AND (GA_STATUTART="UNI" OR GA_STATUTART="GEN") ' + WhereNat,true,-1, '', True);
          end;
        end;
        ferme (Q1);
      end else
      begin
        SQL := 'SELECT * FROM CATALOGU WHERE GCA_CODEBARRE="' + RefSais + '" ' +
          'AND GCA_DATESUP>="' + USDateTime(DatePiece) + '" AND GCA_DATEREFERENCE<="' + USDateTime(DatePiece) +'" ORDER BY GCA_DATEREFERENCE DESC';
        Q1:= OpenSQL(SQL, True,-1, '', True);
        if not Q1.EOF then
        begin
          TOBGCA.SelectDB ('',Q1);
          if TOBGCA.getValue('GCA_ARTICLE')<> '' then
          begin
            Q := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
      						 				'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
                           'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                           'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE '+
                           'WHERE A.GA_ARTICLE="'+TOBGCA.getValue('GCA_ARTICLE')+
                           '" AND (GA_STATUTART="UNI" OR GA_STATUTART="GEN") ' + WhereNat,true,-1, '', True);
      if not Q.EOF then
      begin
        TOBArt.SelectDB('', Q);
        InitChampsSupArticle (TOBARt);
        TOBArt.PutValue('REFARTBARRE', RefSais);
        TrouveRefTiers(TOBArt.GetValue('GA_ARTICLE'), NatTiers, CodeTiers, Prio);
        Result := traOk;
      end;
      Ferme(Q);
          end;
        end;
        ferme (Q1);
      end;

      if ((Result <> traOk) and (Bloque)) then
      begin
        Result := traErreur;
        Break;
      end;
      
    end;
    if Result <> traAucun then Break;
  end;
  TOBGCA.free;
end;

function SelectUneDimension(sCodeGene: string): string;
var TOBL: TOB;
begin
  Result := '';
  ///TheTOB:=TOB.Create('',Nil,-1) ;
  TOBL := TOB.Create('', nil, -1);
  TheTOB := TOBL;
  //TheTOB.Dupliquer(TOBArt,True,True) ;
  AglLanceFiche('GC', 'GCSELECTDIM', '', '', 'GA_ARTICLE=' + sCodeGene + ';ACTION=SELECT;CHAMP= ');
  if TheTOB <> nil then
  begin
    Result := TheTOB.Detail[0].GetValue('GA_ARTICLE');
    ///TheTOB.Free ; TheTOB:=Nil ;
    TheTOB := nil;
  end;
  if TOBL <> nil then TOBL.Free;
end;

procedure LigneVersDim(TOBL, TOBD: TOB; QteDejaSaisi, ReliqDejaSaisi: Double);
begin
  TOBD.AddChampSup('GA_ARTICLE', False);
  TOBD.PutValue('GA_ARTICLE', TOBL.GetValue('GL_ARTICLE'));
  TOBD.AddChampSup('GL_QTEFACT', False);
  TOBD.PutValue('GL_QTEFACT', TOBL.GetValue('GL_QTEFACT'));
  TOBD.AddChampSup('GL_QTERELIQUAT', False);
  TOBD.PutValue('GL_QTERELIQUAT', TOBL.GetValue('GL_QTERELIQUAT'));
  TOBD.AddChampSup('GL_PUHTDEV', False);
  TOBD.PutValue('GL_PUHTDEV', TOBL.GetValue('GL_PUHTDEV'));
  TOBD.AddChampSup('GL_PUTTCDEV', False);
  TOBD.PutValue('GL_PUTTCDEV', TOBL.GetValue('GL_PUTTCDEV'));
  TOBD.AddChampSup('GL_REMISELIGNE', False);
  TOBD.PutValue('GL_REMISELIGNE', TOBL.GetValue('GL_REMISELIGNE'));
  TOBD.AddChampSup('GL_CODEARRONDI', False);
  TOBD.PutValue('GL_CODEARRONDI', TOBL.GetValue('GL_CODEARRONDI'));
  { NEWPIECE }
  TOBD.AddChampSup('GL_PIECEPRECEDENTE', False);
  TOBD.PutValue('GL_PIECEPRECEDENTE', TOBL.GetValue('GL_PIECEPRECEDENTE'));
  TOBD.AddChampSup('GL_QTERESTE', False);
  TOBD.PutValue('GL_QTERESTE', TOBL.GetValue('GL_QTERESTE'));
  TOBD.AddChampSup('GL_SOLDERELIQUAT', False);
  TOBD.PutValue('GL_SOLDERELIQUAT', TOBL.GetValue('GL_SOLDERELIQUAT'));
  TOBD.AddChampSup('GL_MTRESTE', False);
  TOBD.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MTRESTE'));
  TOBD.AddChampSup('GL_MTRELIQUAT', False);
  TOBD.PutValue('GL_MTRELIQUAT', TOBL.GetValue('GL_MTRELIQUAT'));

  TOBD.AddChampSup('_QteDejaSaisi', False);
  TOBD.PutValue('_QteDejaSaisi', QteDejaSaisi);
  TOBD.AddChampSup('_ReliqDejaSaisi', False);
  TOBD.PutValue('_ReliqDejaSaisi', ReliqDejaSaisi);

end;

procedure InitDim(TOBD: TOB; Article: string; QteDejaSaisi, ReliqDejaSaisi: Double);
begin
  TOBD.AddChampSup('GA_ARTICLE', False);
  TOBD.PutValue('GA_ARTICLE', Article);
  TOBD.AddChampSup('GL_QTEFACT', False);
  TOBD.PutValue('GL_QTEFACT', 0);
  TOBD.AddChampSup('GL_QTERELIQUAT', False);
  TOBD.PutValue('GL_QTERELIQUAT', 0);
  TOBD.AddChampSup('GL_PUHTDEV', False);
  TOBD.PutValue('GL_PUHTDEV', 0);
  TOBD.AddChampSup('GL_PUTTCDEV', False);
  TOBD.PutValue('GL_PUTTCDEV', 0);
  TOBD.AddChampSup('GL_REMISELIGNE', False);
  TOBD.PutValue('GL_REMISELIGNE', 0);
  TOBD.AddChampSup('GL_CODEARRONDI', False);
  TOBD.PutValue('GL_CODEARRONDI', '');
  { NEWPIECE }
  TOBD.AddChampSup('GL_PIECEPRECEDENTE', False);
  TOBD.PutValue('GL_PIECEPRECEDENTE', '');
  TOBD.AddChampSup('GL_QTERESTE', False);
  TOBD.PutValue('GL_QTERESTE', 0);
  TOBD.AddChampSup('GL_MTRESTE', False);
  TOBD.PutValue('GL_MTRESTE', 0);
  TOBD.AddChampSup('GL_MTRELIQUAT', False);
  TOBD.PutValue('GL_MTRELIQUAT', 0);
  TOBD.AddChampSup('GL_SOLDERELIQUAT', False);
  TOBD.PutValue('GL_SOLDERELIQUAT', '-');
  TOBD.AddChampSup('_QteDejaSaisi', False);
  TOBD.PutValue('_QteDejaSaisi', QteDejaSaisi);
  TOBD.AddChampSup('_ReliqDejaSaisi', False);
  TOBD.PutValue('_ReliqDejaSaisi', ReliqDejaSaisi);
end;

function CalcQteDejaSaisie(TOBPiece: TOB; Article: string): Double;
var TOBAnc: TOB;
  SommeQte: Double;
  SommeMt : Double;
begin
  SommeQte := 0;
  TOBAnc := TOBPiece.findfirst(['GL_ARTICLE'], [Article], false);
  while TOBAnc <> nil do
  begin
    SommeQte := TOBAnc.GetValue('GL_QTERESTE') + SommeQte;
    TOBAnc := TOBPiece.FindNext(['GL_ARTICLE'], [Article], false);
  end;
  Result := SommeQte;
end;

function CalcMtDejaSaisie(TOBPiece: TOB; Article: string): Double;
var TOBAnc: TOB;
  SommeMt : Double;
begin
  SommeMt := 0;
  TOBAnc := TOBPiece.findfirst(['GL_ARTICLE'], [Article], false);
  while TOBAnc <> nil do
  begin
    SommeMt  := TOBAnc.GetValue('GL_QTERESTE') + SommeMt;
    TOBAnc   := TOBPiece.FindNext(['GL_ARTICLE'], [Article], false);
    SommeMt := TOBAnc.GetValue('GL_MTRESTE') + SommeMT;
  end;
  Result := SommeMt;
end;


function SelectMultiDimensions(sCodeGene: string; GS: THGrid; DepotTrf: string; Action: TActionFiche): Boolean; //AC
var Top, Left, Height, Width: Integer;
  CelluleEcran: Tpoint;
  AuDessus, ValAction: string;
begin
  if Action = taConsult then ValAction := 'CONSULT' else ValAction := 'SAISIE';
  Result := False;
  CelluleEcran := RetourneCoordonneeCellule(1, GS);
  if CelluleEcran.y >= 372 then AuDessus := 'X' else AuDessus := '-';
  Top := CelluleEcran.y;
  Left := CelluleEcran.x;
  Height := GS.Height - GS.RowHeights[0] - GS.RowHeights[1];
  Width := GS.Width - GS.ColWidths[0];
  V_PGI.FormCenter := False;
  AglLanceFiche('GC', 'GCMSELECTDIMDOC', '', '', 'GA_CODEARTICLE=' + sCodeGene + ';ACTION=' + ValAction + ';CHAMP= ;TOP=' + IntToStr(Top) + ';LEFT=' +
    IntToStr(Left) + ';OU=' + AuDessus + ';TYPEPARAM=' + TheTob.GetValue('GP_VENTEACHAT') + ';CODEPARAM=' + TheTob.GetValue('GP_NATUREPIECEG') + ';DEPOT=' +
    DepotTrf + ';HEIGTH=' + IntToStr(Height) + ';WIDTH=' + IntToStr(Width) + '');
  if TheTOB <> nil then Result := (TheTOB.Detail.Count > 0);
end;

function GetCondRecherche(GS: THGrid; CondOrig, RefUnique: string; TOBConds: TOB): string;
var TH: THCritMaskEdit;
  Coord: TRect;
  StCode: string;
  TOBC: TOB;
  Qte: Double;
  CoefN, NbArt, XQ: Double;
begin
  Result := CondOrig;
  Qte := Valeur(GS.Cells[GS.Col, GS.Row]);
  if Qte = 0 then Exit;
  TH := THCritMaskEdit.Create(TForm(GS.Owner));
  TH.Parent := GS;
  Coord := GS.CellRect(GS.Col, GS.Row);
  TH.Top := Coord.Top;
  TH.Left := Coord.Left;
  TH.Width := 0;
  TH.Visible := False;
  TH.Plus := RefUnique;
  if LookupList(TH, 'Conditionnement en cours : ' + CondOrig, 'CONDITIONNEMENT', 'GCO_CODECOND', 'GCO_LIBELLE', 'GCO_ARTICLE="' + RefUnique + '"',
    'GCO_CODECOND', True, 10) then
  begin
    StCode := TH.Text;
    if StCode = CondOrig then
    begin
      TH.Free;
      Exit;
    end;
    TOBC := TOBConds.FindFirst(['GCO_ARTICLE', 'GCO_CODECOND'], [RefUnique, StCode], False);
    if TOBC <> nil then
    begin
      CoefN := 1 - 2 * Ord(Qte < 0);
      Qte := Abs(Qte);
      NbArt := TOBC.GetValue('GCO_NBARTICLE');
      if NbArt = 0 then NbArt := 1;
      XQ := Qte / NbArt;
      if Trunc(XQ) <> XQ then
      begin
        XQ := Trunc(XQ) + 1;
        Qte := XQ * NbArt * CoefN;
        GS.Cells[GS.Col, GS.Row] := StrF00(Qte, V_PGI.OkDecQ);
      end;
    end;
    Result := StCode;
  end;
  TH.Free;
end;

procedure ZoomArticle(RefUnique, TypeArt: string; Action: TActionFiche);
//var StAct: string;
begin
  {$IFDEF BTP}
  if Action = taConsult then Action := taconsult else Action := taModif;
  V_PGI.DispatchTT(7, Action, RefUnique, '', 'TYPEARTICLE=' + TypeArt);
  {$ELSE}
  {$IFDEF GPAO}
  if Action = taConsult then Action := taconsult else Action := taModif;
  V_PGI.DispatchTT(7, Action, RefUnique, '', 'TYPEARTICLE=' + TypeArt);
  {$ELSE}
  {$IFDEF CHR}
  if Action = taConsult then StAct := 'ACTION=CONSULTATION' else StAct := 'ACTION=MODIFICATION';
  AGLLanceFiche('H', 'HRPRESTATIONS', '', RefUnique, StAct + ';TYPEARTICLE=' + TypeArt)
    {$ELSE}
  if Action = taConsult then StAct := 'ACTION=CONSULTATION' else StAct := 'ACTION=MODIFICATION';
  if ctxMode in V_PGI.PGIContexte then
  begin
    if Action = taConsult then Action := taconsult else Action := taModif;
    V_PGI.DispatchTT(7, Action, RefUnique, '', StAct + ';TYPEARTICLE=' + TypeArt)
  end
  else if (IsArticleSpecif(TypeArt) = 'FicheAffaire')
    then
  begin
    // mcd 06/03/02AGLLanceFiche('AFF','AFARTICLE','',RefUnique,StAct+';TYPEARTICLE='+TypeArt)
    if Action = taConsult then Action := taconsult else Action := taModif;
    V_PGI.DispatchTT(7, Action, RefUnique, '', 'TYPEARTICLE=' + TypeArt);
  end
  else AGLLanceFiche('GC', 'GCARTICLE', '', RefUnique, StAct + ';TYPEARTICLE=' + TypeArt);
  {$ENDIF}
  {$ENDIF}
  {$ENDIF}
end;

procedure ZoomCatalogue(RefCata, RefFour: string; Action: TActionFiche);
var StAct: string;
begin
  if Action = taConsult then StAct := 'ACTION=CONSULTATION' else StAct := 'ACTION=MODIFICATION';
  AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', RefCata + ';' + RefFour, StAct);
end;

procedure ErreurCodeArticle(RefSais: string);
begin
  ShowMessage('Le code article ' + RefSais + ' n''existe pas');
end;

procedure ErreurCodeArticleFour(RefSais: string);
begin
  HShowmessage('0;?caption?;L''article ' + RefSais + ' n''appartient pas au fournisseur de ce document;E;O;O;O', '', '');
end;

{==================================== Commercial ==================================}

function GetWhereRepresentant(TOBG: TOB; ZoneCom: string; ModeSql: Boolean): string;
var NaturePieceG, NomChamp, ValChamp, TypeCom, CodeC, Pref: string;
begin
  Result := '';
  // type de commercial
  TypeCom := '';
  Pref := '';
  NaturePieceG := '';
  if TOBG <> nil then
  begin
    if TOBG.NomTable = 'PIECE' then Pref := 'GP_' else
      if TOBG.NomTable = 'LIGNE' then Pref := 'GL_' else
      ;
  end;
  if Pref <> '' then
  begin
    NaturePieceG := TOBG.GetValue(Pref + 'NATUREPIECEG');
    TypeCom := GetInfoParPiece(NaturePieceG, 'GPP_TYPECOMMERCIAL');
  end;
{$IFNDEF BTP}
  if TypeCom = '' then
  begin
    TypeCom := 'REP';
    if ctxFO in V_PGI.PGIContexte then TypeCom := 'VEN';
    {$IFDEF GESCOM}
    if NaturePieceG = 'FFO' then TypeCom := 'VEN';
    {$ENDIF}
  end;
{$ENDIF}
  // Requête SQL ou syntaxe ALLanceFiche
  if TypeCom <> '' then
  begin
    if ModeSql then Result := 'GCL_TYPECOMMERCIAL="' + TypeCom + '" AND '
    else Result := 'GCL_TYPECOMMERCIAL=' + TypeCom+ ';';
  end;
  // date de fermeture
  if (ctxMode in V_PGI.PGIContexte) or (ctxFO in V_PGI.PGIContexte) then
  begin
    if ModeSql then Result := Result + 'GCL_DATESUPP>"' + USDateTime(Date) + '"'
    else Result := Result + 'GCL_DATESUPP=' + DateToStr(Date);
  end;
  // filtre par défaut sur la zone commerciale
  NomChamp := 'GCL_ZONECOM';
  ValChamp := ZoneCom;
  if NaturePieceG <> '' then
  begin
    CodeC := GetInfoParPiece(NaturePieceG, 'GPP_FILTRECOMM');
    if CodeC <> '' then // DBR Fiche 10745
      NomChamp := RechDom('GCINFOCOMMERCIAL', CodeC, True);
  end;
  if NomChamp <> 'GCL_ZONECOM' then
  begin
    CodeC := Pref + ExtractSuffixe(NomChamp);
    if TOBG.FieldExists(CodeC) then ValChamp := TOBG.GetValue(CodeC);
  end;
  if ValChamp <> '' then
  begin
    if ModeSql then Result := Result + ' AND ' + NomChamp + '="' + ValChamp + '"'
    else Result := Result + ';' + NomChamp + '=' + ValChamp + ';';
  end;
end;

function GetRepresentantLookUp(CC: TControl; TitreSel: string; Letl: TTypeLocate; TOBPiece: TOB): boolean;
begin
  Result := LookupList(CC, TitreSel, 'COMMERCIAL', 'GCL_COMMERCIAL', 'GCL_LIBELLE', GetWhereRepresentant(TOBPiece, '', True), 'GCL_COMMERCIAL', True, 9, '',
    Letl);
end;

function GetRepresentantMul(CC: TControl; ZoneCom: string; TOBPiece: TOB): boolean;
var StC, Repres: string;
  GG: THGrid;
begin
  Result := False;
  if CC is THGrid then
  begin
    GG := THGrid(CC);
    Repres := GG.Cells[GG.Col, GG.Row];
  end else
  begin
    Repres := THCritMaskEdit(CC).Text;
  end;
// DBR Fiche 10745
//  StC := GetWhereRepresentant(TOBPiece, ZoneCom, False) + ';GCL_COMMERCIAL=' + Repres;
  StC := GetWhereRepresentant(TOBPiece, ZoneCom, False);
  // Si false en dernier param de GetWhereRepresentant, ';' en fin de chaine retournée
  if Repres <> '' then StC := StC + 'GCL_COMMERCIAL=' + Repres;
  Repres := AGLLanceFiche('GC', 'GCCOMMERCIAL_RECH', StC, '', ''); // DBR fiche 10745
  if Repres <> '' then
  begin
    Result := True;
    if CC is THGrid then
    begin
      GG := THGrid(CC);
      GG.Cells[GG.Col, GG.Row] := Repres;
    end else
    begin
      THCritMaskEdit(CC).Text := Repres;
    end;
  end;
end;

function GetRepresentantEntete(TC: THCritMaskEdit; TitreSel, ZoneCom: string; TOBPiece: TOB): boolean;
begin
  if GetParamSocSecur('SO_GCRECHCOMMAV', False) then Result := GetRepresentantMul(TC, ZoneCom, TOBPiece)
  else Result := GetRepresentantLookUp(TC, TitreSel, tlLocate, TOBPiece);
end;

function GetRepresentantSaisi(GS: THGrid; TitreSel, ZoneCom: string; Letl: TTypeLocate; TOBPiece: TOB): boolean;
begin
  if GetParamSocSecur('SO_GCRECHCOMMAV', False) then Result := GetRepresentantMul(GS, ZoneCom, TOBPiece)
  else Result := GetRepresentantLookUp(GS, TitreSel, Letl, TOBPiece);
end;

//FV1 : 20/02/2014 - FS#898 - BAGE : Ajouter une zone pour indiquer le contact du bon de commande
function GetRessourceEntete(TC: THCritMaskEdit; TitreSel, ZoneCom: string; TOBPiece: TOB): boolean;
begin
  if GetParamSocSecur('SO_AFRECHRESAV', False) then Result := GetRessourceMul(TC, ZoneCom, TOBPiece)
  else Result := GetRessourceLookUp(TC, TitreSel, tlLocate, TOBPiece);
end;

//FV1 : 10/12/20014 - FS#1312 - SOCMA : en recherche sur contact, la loupe et l'ajout ouvrent des fenêtres sur les commerciaux
function GetRessourceLookUp(CC: TControl; TitreSel: string; Letl: TTypeLocate; TOBPiece: TOB): boolean;
begin
  Result := LookupList(CC, TitreSel, 'RESSOURCE', 'ARS_RESSOURCE', 'ARS_LIBELLE,ARS_LIBELLE2', 'ARS_TYPERESSOURCE="SAL"', 'ARS_RESSOURCE', True, 6, '', Letl);
end;

function GetRessourceMul(CC: TControl; ZoneCom: string; TOBPiece: TOB): boolean;
var StC       : String;
    Ressource : String;
    Action    : String;
    Range     : String;
    GG        : THGrid;
begin

  Result      := False;

  if CC is THGrid then
  begin
    GG        := THGrid(CC);
    Ressource := GG.Cells[GG.Col, GG.Row];
  end
  else
  begin
    Ressource  := THCritMaskEdit(CC).Text;
  end;

  Action := ';ACTION=RECH';

  StC := 'TYPERESSOURCE=SAL ';

  if Ressource <> '' then Range := 'ARS_RESSOURCE=' + Ressource;
  Ressource := AGLLanceFiche('BTP', 'BTRESSRECH_Mul',Range, '', Stc+ ';ACTION=RECH'); // DBR fiche 10745

  if Ressource <> '' then
  begin
    Result := True;
    if CC is THGrid then
    begin
      GG := THGrid(CC);
      GG.Cells[GG.Col, GG.Row] := Ressource;
    end else
    begin
      THCritMaskEdit(CC).Text := Ressource;
    end;
  end;
end;

{==================================== Dépôts ==================================}

function GetDepotLookUp(CC: TControl; TitreSel: string; Letl: TTypeLocate): boolean;
begin
  Result := LookUpList (CC, TitreSel, 'DEPOTS', 'GDE_DEPOT', 'GDE_LIBELLE', '', 'GDE_DEPOT', True, 0, '', Letl);
//  Result := LookupList(CC, TitreSel, 'CHOIXCOD', 'CC_CODE', 'CC_LIBELLE', 'CC_TYPE="DEP"', 'CC_CODE', True, 0, '', Letl);
// ce lookuplist ne correspond a rien du tout.
end;

function GetDepotSaisi(GS: THGrid; TitreSel: string; Letl: TTypeLocate): boolean;
begin
{$IFDEF EAGLCLIENT}     //JS 29102003
    if GetDepotLookUp(GS, TitreSel, Letl) then
    begin
      Result := (GS.Cells[GS.Col, GS.Row] <> TraduireMemoire('Erreur'));
      if not Result then GS.Cells[GS.Col, GS.Row] := '';
    end else Result := False;
{$ELSE}
    Result := GetDepotLookUp(GS, TitreSel, Letl);
{$ENDIF}
end;

function GetLibellesAutos(GS: THGrid; TitreSel: string): boolean;
begin
  Result := LookupList(GS, TitreSel, 'CHOIXEXT', 'YX_LIBELLE', 'YX_CODE', 'YX_TYPE="GCI"', 'YX_LIBELLE', True, 900, '', tlLocate);
end;

procedure GetCommentaireLigne(var CodeCom,Libelle,Texte : string );
var QQ: TQuery;
begin
  Libelle := '';
  Texte := '';
  if CodeCom = '' then CodeCom := AGLLanceFiche('BTP', 'BTMULCOMMENTAIRE', '', '', '');
  if CodeCom = '' then exit;
  QQ := OpenSql('Select GCT_LIBELLE,GCT_CONTENU From COMMENTAIRE WHERE GCT_CODE="' + CodeCom + '"', True,-1, '', True);
  if not QQ.Eof then
  begin
    Texte := QQ.FindField('GCT_CONTENU').AsString;
    Libelle := QQ.FindField('GCT_LIBELLE').AsString;
  end;
  Ferme(QQ);
end;

procedure ZoomDispo(RefUnique, RefFour, Client, Four, Depot: string; Contrem: Boolean);
var StAct, St1: string;
begin
  StAct := 'ACTION=CONSULTATION';
  if (not Contrem) then
  begin
    {$IFDEF GPAO}
    AGLLanceFiche('W', 'WDISPOART_MUL', 'GQ_ARTICLE=' + RefUnique, '', StAct);
    {$ELSE}
    AGLLanceFiche('GC', 'GCDISPOART_MUL', 'GQ_ARTICLE=' + RefUnique, '', StAct);
    {$ENDIF}
  end else
  begin
    //st1 := 'GQC_DEPOT='+Depot+';GQC_REFERENCE='+RefFour+';GQC_CLIENT='+Client+
    st1 := 'GQC_DEPOT=' + Depot + ';GQC_REFERENCE=' + RefFour + ';XX_WHERE=GQC_CLIENT IN("","' + Client + '")' +
      ';GQC_FOURNISSEUR=' + Four;
    AGLLanceFiche('GC', 'GCDISPOCONTR_MUL', st1, '', StAct);
  end;
end;

{==================================== Affichages ==================================}

function CalculMargeLigne(TOBLig, TOBA: TOB): Double;
var RefU, QualM: string;
  PxValo, PuNet, MC: double;
begin
  Result := 0;
  MC := 0;
  if TOBLig = nil then Exit;
  if TOBA = nil then Exit;
  RefU := TOBLig.GetValue('GL_ARTICLE');
  if RefU = '' then Exit;
  QualM := TOBA.GetValue('GA_QUALIFMARGE');
  if QualM = '' then Exit;
  if VH_GC.GCMargeArticle = 'PMA' then PxValo := TOBLig.GetValue('GL_PMAP') else
    if VH_GC.GCMargeArticle = 'PMR' then PxValo := TOBLig.GetValue('GL_PMRP') else
    if VH_GC.GCMargeArticle = 'DPA' then PxValo := TOBLig.GetValue('GL_DPA') else
    if VH_GC.GCMargeArticle = 'DPR' then PxValo := TOBLig.GetValue('GL_DPR') else Exit;
  PuNet := TOBLig.GetValue('GL_PUHTNET');
  if QualM = 'MO' then
  begin
    MC := PuNet - PxValo;
  end else if QualM = 'CO' then
  begin
    if TOBLig.GetValue('GL_FACTUREHT') = '-' then PuNet := TOBLig.GetValue('GL_PUTTCNET');
    if PxValo > 0 then MC := (PuNet / PxValo) else MC := 0;
  end else if QualM = 'PC' then
  begin
    if PuNet > 0 then MC := ((PuNet - PxValo) / PuNet) * 100 else MC := 0; // DBR * 100 pour faire un pourcentage 13112003
  end;
  Result := MC;
end;

function CalculMargeTotLigne(TOBLig: TOB): Double;
var RefU: string;
  PxValo, TotalLig, Qte, MC: double;
begin
  Result := 0;
  if TOBLig = nil then Exit;
  RefU := TOBLig.GetValue('GL_ARTICLE');
  if RefU = '' then Exit;
  if VH_GC.GCMargeArticle = 'PMA' then PxValo := TOBLig.GetValue('GL_PMAP') else
    if VH_GC.GCMargeArticle = 'PMR' then PxValo := TOBLig.GetValue('GL_PMRP') else
    if VH_GC.GCMargeArticle = 'DPA' then PxValo := TOBLig.GetValue('GL_DPA') else
    if VH_GC.GCMargeArticle = 'DPR' then PxValo := TOBLig.GetValue('GL_DPR') else Exit;
  TotalLig := TOBLig.GetValue('GL_TOTALHT');
  Qte := TOBLig.GetValue('GL_QTEFACT');
  MC := TotalLig - (PxValo * Qte);
  Result := MC;
end;

function MontreITiers(TOBTiers: TOB; Champ: string): string;
begin
  Result := '';
  if TOBTiers = nil then Exit;
  if Champ = '' then Exit;
  Result := VarAsType(TOBTiers.GetValue(Champ), VarString);
end;

function MontreIArticle(TOBArt: TOB; Champ: string): string;
begin
  Result := '';
  if TOBArt = nil then Exit;
  if Champ = '' then Exit;
  Result := VarAsType(TOBArt.GetValue(Champ), VarString);
  (* Dimensions
  GRILLEDIM1, tablette GCGRILLEDIM1
  CODEDIM1
  *)
end;

function MontreILigne(TOBLig, TobArt: TOB; Champ: string): string;
begin
  Result := '';
  if TOBLig = nil then Exit;
  if Champ = '' then Exit;
  if Champ = 'GL_DEPOT' then Result := RechDom('GCDEPOT', TOBLig.GetValue(Champ), False) else
    if Champ = 'GL_REPRESENTANT' then Result := RechDom('GCREPRESENTANT', TOBLig.GetValue(Champ), False) else
    if Champ = 'GL_MARGELIGNE' then Result := Strf00(CalculMargeLigne(TOBLig, TOBArt), V_PGI.OkDecV) else
    if Champ = 'GL_MARGETOTLIGNE' then Result := Strf00(CalculMargeTotLigne(TOBLig), V_PGI.OkDecV) else
    {$IFDEF DEBUG}
    if Champ = 'GL_AFFAIRE' then Result := TOBLig.GetValue('GL_AFFAIRE') else
    {$ELSE}
    {$IFDEF BTP}
    {$ELSE IF AFFAIRE}
    if Champ = 'GL_AFFAIRE' then Result := CodeAffaireAffiche(TOBLig.GetValue('GL_AFFAIRE')) else
    {$ENDIF}
    {$ENDIF}
    if Champ = 'GL_APPORTEUR' then Result := RechDom('AFAPPORTEUR', TOBLig.GetValue(Champ), False) else
    if Champ = 'GL_NONIMPRIMABLE' then if TOBLig.GetValue(Champ) = 'X' then Result := 'Non' else Result := 'Oui' else
    if Pos('GL_LIBREART', Champ) > 0 then Result := RechDom('GCLIBREART' + Copy(Champ, 12, 1), TOBLig.GetValue(Champ), False) else
    Result := VarAsType(TOBLig.GetValue(Champ), VarString);
end;

function MontreIStock(TOBArt, TOBLig: TOB; Champ: string): string;
var Depot: string;
  	TOBD: TOB;
		dResult,CoefUaUS,CoefUSUV,FUV,FUS, FUA : double;
    VenteAchat : string;
begin
  Result := '';
  dResult := 0;
  if TOBArt = nil then Exit;
  if TOBLig = nil then Exit;
  if Champ = '' then Exit;
  VenteAchat := GetInfoParPiece(TOBLig.GetString('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  //
  FUV := RatioMesure('PIE', TOBLig.getValue('GL_QUALIFUNITEVTE'));
  FUS := RatioMesure('PIE', TOBLig.getValue('GL_QUALIFUNITESTO'));
  FUA := RatioMesure('PIE', TOBLig.getValue('GL_QUALIFUNITEACH'));
  CoefUaUS := TOBLig.getDouble('GL_COEFCONVQTE');
  CoefUSUV := TOBLig.getDouble('GL_COEFCONVQTEVTE');
  //
  Depot := TOBLig.GetValue('GL_DEPOT');
  TOBD := TOBArt.FindFirst(['GQ_DEPOT'], [Depot], False);
  if TOBD <> nil then
  begin
    if Champ = 'GQ_PHYSIQUE' then dResult := TOBD.GetValue('GQ_PHYSIQUE') else
      if Champ = 'GQ_DISPONIBLE' then
      dResult := TOBD.GetValue('GQ_PHYSIQUE') - TOBD.GetValue('GQ_RESERVECLI') - TOBD.GetValue('GQ_PREPACLI') else
      if Champ = 'STOCNET' then
      dResult := (TOBD.GetValue('GQ_PHYSIQUE') - TOBD.GetValue('GQ_RESERVECLI') - TOBD.GetValue('GQ_PREPACLI')) + TOBD.GetValue('GQ_RESERVEFOU') else
      dResult := VarAsType(TOBD.GetValue(Champ), varDouble);
  end;

  if VenteAchat = 'VEN' then
  begin
    if CoefUSUV <> 0 then
    begin
    	dResult := dResult * CoefUSUV;
    end else
    begin
			dResult := dResult * FUS / FUV;
    end;
  end else if VenteAchat = 'ACH' then
  begin
    if CoefUaUS <> 0 then
    begin
    	dResult := dResult * CoefUAUS;
    end else
    begin
			dResult := dResult * FUS / FUA;
    end;
  end;
  result := StrF00(dResult, V_PGI.OkDecQ)
end;

function MontreIStockContreM(TOBCata, TOBLig: TOB; Champ: string): string;
var Depot: string;
  TOBD: TOB;
  QPhys, QResaCli, QPrepaCli, QResaFou: Double;
begin
  Result := '';
  if TOBCata = nil then Exit;
  if TOBLig = nil then Exit;
  if Champ = '' then Exit;
  Depot := TOBLig.GetValue('GL_DEPOT');
  QPhys := 0;
  QResaCli := 0;
  QPrepaCli := 0;
  QResaFou := 0;
  TOBD := TOBCata.FindFirst(['GQC_DEPOT'], [Depot], False);
  //if TOBD<>Nil then
  while TOBD <> nil do
  begin
    QPhys := QPhys + TOBD.GetValue('GQC_PHYSIQUE');
    QResaCli := QResaCli + TOBD.GetValue('GQC_RESERVECLI');
    QPrepaCli := QPrepaCli + TOBD.GetValue('GQC_PREPACLI');
    QResaFou := QResaFou + TOBD.GetValue('GQC_RESERVEFOU');
    TOBD := TOBCata.FindNext(['GQC_DEPOT'], [Depot], False);
  end;
  if Champ = 'GQ_PHYSIQUE' then Result := StrF00(QPhys, V_PGI.OkDecQ) else
    if Champ = 'GQ_DISPONIBLE' then Result := StrF00(QPhys - QResaCli - QPrepaCli, V_PGI.OkDecQ) else
    if Champ = 'STOCNET' then Result := StrF00((QPhys - QResaCli - QPrepaCli) + QResaFou, V_PGI.OkDecQ) else
    Result := VarAsType(TOBD.GetValue(Champ), VarString);
end;

function MontreIStockGen(TOBArt, TOBLig: TOB; Champ: string): string;
var Depot: string;
  QQ: TQuery;
begin
  Result := '';
  if TOBArt = nil then Exit;
  if TOBLig = nil then Exit;
  if Champ = '' then Exit;
  Depot := TOBLig.GetValue('GL_DEPOT');
  QQ :=
    OpenSQL('SELECT SUM(GQ_PHYSIQUE) AS PHYSIQUE, (SUM(GQ_PHYSIQUE)-SUM(GQ_RESERVECLI)) as DISPONIBLE, ((SUM(GQ_PHYSIQUE)-SUM(GQ_RESERVECLI))+SUM(GQ_RESERVEFOU)) as STOCNET'
    + ' FROM DISPO WHERE GQ_ARTICLE LIKE "' + Copy(TOBArt.GetValue('GA_ARTICLE'), 1, 18) + '%" AND GQ_DEPOT="' + Depot + '" AND GQ_CLOTURE="-"', True,-1, '', True);
  if not QQ.EOF then
  begin
    if Champ = 'GQ_PHYSIQUE' then Result := StrF00(QQ.FindField('PHYSIQUE').AsFloat, V_PGI.OkDecQ) else
      if Champ = 'GQ_DISPONIBLE' then Result := StrF00(QQ.FindField('DISPONIBLE').AsFloat, V_PGI.OkDecQ) else
      if Champ = 'STOCNET' then Result := StrF00(QQ.FindField('STOCNET').AsFloat, V_PGI.OkDecQ) else
  end;
  Ferme(QQ);
end;

procedure MontreInfosLigne(TOBLig, TOBArt, TOBCata, TOBTiers: TOB; Info: THGrid; PPInfosLigne: TStrings);
var i, Num, idim: integer;
  Pref, St, Lib, Champ, StChamp, Grille, CodeDim: string;
  // modif BTP
  PxValo, PuNet, MtMarge: double;
  EntPxValo, EntPuNet, TPSCumule: double;
  TOBP: TOB;
  QQ: Tquery;
  GestionMarq : boolean;
begin
  if TobLig = nil then Exit;
  TOBP := TOBLig.Parent;
  GestionMarq := GetParamSocSecur('SO_BTGESTIONMARQ', False);
  // Modif BTP
  // C'est nouveau on calcule la marge incluant la tva trop drole..
  (*
  if TOBLIG.GetValue('GL_FACTUREHT') = 'X' then PuNet := TOBLig.GetValue('GL_MONTANTHT')
  else PuNet := TOBLig.GetValue('GL_MONTANTTTC');
  *)
  PuNet := TOBLig.GetValue('GL_MONTANTHT');
  if VH_GC.GCMargeArticle = 'PMA' then PxValo := TOBLig.GetValue('GL_PMAP') else
    if VH_GC.GCMargeArticle = 'PMR' then PxValo := TOBLig.GetValue('GL_PMRP') else
    if VH_GC.GCMargeArticle = 'DPA' then PxValo := TOBLig.GetValue('GL_DPA') else
    if VH_GC.GCMargeArticle = 'DPR' then PxValo := TOBLig.GetValue('GL_DPR') else PxValo := 0;
  // C'est nouveau on calcule la marge incluant la tva trop drole..
  (*
  if TOBP.GetValue('GP_FACTUREHT') = 'X' then EntPuNet := TOBP.GetValue('GP_TOTALHT')
  else EntPuNet := TOBP.GetValue('GP_TOTALTTC');
  *)
  EntPuNet := TOBP.GetValue('GP_TOTALHT');
  if VH_GC.GCMargeArticle = 'PMA' then EntPxValo := TOBP.GetValue('GP_MONTANTPA') else
    if VH_GC.GCMargeArticle = 'PMR' then EntPxValo := TOBP.GetValue('GP_MONTANTPR') else
    if VH_GC.GCMargeArticle = 'DPA' then EntPxValo := TOBP.GetValue('GP_MONTANTPA') else
    if VH_GC.GCMargeArticle = 'DPR' then EntPxValo := TOBP.GetValue('GP_MONTANTPR') else EntPxValo := 0;

  // VARIANTE
  (* if copy(TOBLIG.GetValue('GL_TYPELIGNE'), 1, 2) <> 'TP' then *)
  if not ISFinParagraphe (TOBLIG) then
  begin
    PxValo := PXValo * TOBLig.getValue('GL_QTEFACT');
    {$IFDEF BTP}
    TPSCumule := TOBLIG.GetValue('GL_TPSUNITAIRE') * TOBLig.getValue('GL_QTEFACT');
    {$ENDIF}
  end else
  begin
    {$IFDEF BTP}
    TPSCumule := TOBLIG.GetValue('GL_TPSUNITAIRE');
    {$ENDIF}
  end;
  // --
  Info.RowCount := 1;
  Num := 0;
  for i := 1 to 8 do
  begin
    Info.Cells[0, i] := '';
    Info.Cells[1, i] := '';
    if i <= PPInfosLigne.Count then
    begin
      St := PPInfosLigne[i - 1];
      if St = '' then Continue;
      Lib := ReadTokenSt(St);
      Champ := ReadTokenSt(St);
      if Champ = 'DIMENSION' then
      begin
        if TOBArt <> nil then for iDim := 1 to MaxDimension do
          begin
            Grille := TOBArt.GetValue('GA_GRILLEDIM' + IntToStr(idim));
            CodeDim := TOBArt.GetValue('GA_CODEDIM' + IntToStr(idim));
            if ((Grille <> '') and (CodeDim <> '')) then
            begin
              Lib := RechDom('GCGRILLEDIM' + IntToStr(iDim), Grille, False);
              StChamp := GCGetCodeDim(Grille, CodeDim, iDim);
              Inc(Num);
              if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
              Info.Cells[0, Num - 1] := Lib;
              Info.Cells[1, Num - 1] := StChamp;
            end;
          end;
      end else
        if Champ = 'STOCNET' then
      begin
        if TOBLig.GetValue('GL_TYPEDIM') = 'GEN' then StChamp := MontreIStockGen(TOBArt, TOBLig, Champ)
        else StChamp := MontreIStock(TOBArt, TOBLig, Champ);
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'CLTAFF' then //mcd 27/05/03 AFFAIRE
      begin
        QQ := OPENSQL('SELECT T_LIBELLE From TIERS WHERE T_TIERS= (SELECT AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE="'
          + TOBLig.GetValue('GL_AFFAIRE') + '")', True,-1, '', True);
        if not QQ.EOF then StChamp := QQ.Fields[0].AsString
        else StChamp := '';
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
        Ferme(QQ);
      end else
        // Modif BTP
        if Champ = 'TPS_CUMUL' then
      begin
        if TpsCumule <> 0 then stChamp := StrF00(TpsCUmule, 3) + ' H'
        else stchamp := StrF00(0, 3);

        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        // Modif BTP
        if Champ = 'MARGE_POU' then
      begin
        if GestionMarq then
        begin
          if PuNet <> 0 then stChamp := StrF00(arrondi(((PuNet - PxValo) / PuNet) * 100, 2), V_PGI.OkDecV)
          else stchamp := StrF00(0, V_PGI.OkDecV);

          Inc(Num);
          if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        end else
        begin
        if Pxvalo <> 0 then stChamp := StrF00(arrondi(((PuNet - PxValo) / Pxvalo) * 100, 2), V_PGI.OkDecV)
        else stchamp := StrF00(0, V_PGI.OkDecV);

        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        end;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'MARGE_MONT' then
      begin
        MtMarge := (PuNet - PxValo);
        stChamp := STrF00(MtMarge, V_PGI.OkDecV);
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'MARGE_COEF' then
      begin
        if Pxvalo <> 0 then stChamp := StrF00(arrondi((PuNet / Pxvalo), 4), 4)
        else stchamp := StrF00(0, 4);
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'ENTMARGE_POU' then
      begin
        if EntPuNet <> 0 then stChamp := StrF00(arrondi(((EntPuNet - EntPxValo) / EntPuNet) * 100, 2), V_PGI.OkDecV)
        else stchamp := StrF00(0, V_PGI.OkDecV);

        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'ENTMARGE_MONT' then
      begin
        MtMarge := (EntPuNet - EntPxValo);
        stChamp := STrF00(MtMarge, V_PGI.OkDecV);
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'ENTMARGE_COEF' then
      begin
        if EntPxvalo <> 0 then stChamp := StrF00(arrondi((EntPuNet / EntPxvalo), 2), V_PGI.OkDecQ)
        else stchamp := StrF00(0, V_PGI.OkDecQ);
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        if Champ = 'MARGEPV_POU' then
      begin
        if PuNet <> 0 then stChamp := StrF00(arrondi(((PuNet - PxValo) / PuNet) * 100, 2), V_PGI.OkDecV)
        else stchamp := StrF00(0, V_PGI.OkDecV);

        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end else
        // ----
      begin
        Inc(Num);
        if Num > Info.RowCount then Info.RowCount := Info.RowCount + 1;
        Pref := Copy(Champ, 1, Pos('_', Champ) - 1);
        if Pref = 'GA' then StChamp := MontreIArticle(TOBArt, Champ) else
          if Pref = 'GL' then StChamp := MontreILigne(TOBLig, TOBArt, Champ) else
          if Pref = 'T' then StChamp := MontreITiers(TOBTiers, Champ) else
          if Pref = 'GQ' then
        begin
          if (TOBLig <> nil) and (TOBLig.GetValue('GL_TYPEDIM') = 'GEN') then StChamp := MontreIStockGen(TOBArt, TOBLig, Champ)
          else
          begin
            if (TOBLig.GetValue('GL_ENCONTREMARQUE') <> 'X') then
              StChamp := MontreIStock(TOBArt, TOBLig, Champ)
            else
              StChamp := MontreIStockContreM(TOBCata, TOBLig, Champ);
          end;
        end else
          StChamp := '';
        // reprise libellé paramétrable des stats articles à la ligne
        if Pos('GL_LIBREART', Champ) > 0 then GCTitreZoneLibre(Champ, Lib);
        Info.Cells[0, Num - 1] := Lib;
        Info.Cells[1, Num - 1] := StChamp;
      end;
    end;
  end;
  Info.Row := 0;
  info.refresh;
end;

procedure ControlsVisible(NomPanel: THPanel; Visible: boolean; SfControlName: array of string);
  function FindStInTab(Tab: array of string; St: string): boolean;
  var i: integer;
  begin
    Result := False;
    for i := Low(Tab) to High(Tab) do
      if UpperCase(Tab[i]) = UpperCase(St) then
      begin
        Result := True;
        Break;
      end;
  end;
var i: integer;
  TempControl: TControl;
begin
  for i := 0 to NomPanel.ControlCount - 1 do
  begin
    TempControl := NomPanel.Controls[i];
    if FindStInTab(SfControlName, TempControl.Name) then TempControl.visible := not Visible
    else TempControl.visible := Visible;
  end;
end;

{=================================== Tarifs ==================================}
{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 17/03/2003
Modifié le ... : 17/03/2003
Description .. : Fonction appelée depuis FactUtil et UTofSelectDimDoc.
Suite ........ : Recherche des prix d'achat :
Suite ........ : La recherche est effectuée en utilisant les priorités définies
Suite ........ : dans le paramétrage des pièces -> GPP_APPELPRIX,
Suite ........ : GPP_APPELPRIX2 et GPP_APPELPRIX3.
Suite ........ : Si le prix trouvé est inférieur ou égal à 0, la priorité suivante
Suite ........ : est utilisée.
Suite ........ : Par défaut, si aucun prix n'est trouvé :
Suite ........ : - si la priorité 1 est parmi :
Suite ........ :    - Stock - Dernier prix de revient
Suite ........ :    - Stock - Prix moyen de revient pondéré
Suite ........ :    - Article - Dernier prix de revient
Suite ........ :    - Article - Prix moyen de revient pondéré
Suite ........ : le prix de revient standard de la fiche article est utilisé,
Suite ........ : sinon le prix d'achat standard de la fiche article est utilisé.
Mots clefs ... : PRIX;ACHAT;RECHERCHE
*****************************************************************}

function GetPrixAchat(NaturePiece, CodeArt: string; TOBArt, TOBSto, TOBCat: TOB; FactureHT: boolean; RatioStock, RatioDoc: double): double;
var iAppelPrix: integer;
  Prix, PrixDefaut: Double;
  TOBDD: TOB;
  QQ: TQuery;
  CodeP, stSql: string;
begin
  TOBDD := nil; // Init. Tob Stock Dépôt par défaut
  iAppelPrix := 0;
  Prix := 0.0;
  PrixDefaut := 0.0;
  while ((Prix = 0.0) and (iAppelPrix < 3)) do
  begin
    inc(iAppelPrix);
    if iAppelPrix = 1 then
    begin
      CodeP := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX');
      if (CodeP = 'SDR') or (CodeP = 'SMR')
        then PrixDefaut := Arrondi(TOBArt.GetValue('GA_PRHT'), V_PGI.OkDecV)
      else PrixDefaut := Arrondi(TOBArt.GetValue('GA_PAHT'), V_PGI.OkDecV);
    end
    else if iAppelPrix = 2 then CodeP := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX2')
    else if iAppelPrix = 3 then CodeP := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX3');

    if (((CodeP = 'DDA') or (CodeP = 'DDR') or (CodeP = 'DMA') or (CodeP = 'DMR'))
      and (TOBDD = nil)) then
    begin
      // Chargement des prix du stock du dépôt par défaut
      stSql := 'select GQ_DPA,GQ_PMAP,GQ_DPR,GQ_PMRP from DISPO where GQ_ARTICLE="' + CodeArt +
        '" and GQ_DEPOT="' + GetParamSoc('SO_GCDEPOTDEFAUT') + '" and GQ_CLOTURE="-"';
      QQ := OpenSQL(stSql, True,-1, '', True);
      if not QQ.EOF then TOBDD.LoadDetailDB('DISPO', '', '', QQ, False);
      Ferme(QQ);
    end;

    if (CodeP = 'SDA') and (TOBSto <> nil) and TOBSto.FieldExists('GQ_DPA')
      then Prix := Arrondi(TOBSto.GetValue('GQ_DPA') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'SDR') and (TOBSto <> nil) and TOBSto.FieldExists('GQ_DPR')
      then Prix := Arrondi(TOBSto.GetValue('GQ_DPR') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'SMA') and (TOBSto <> nil) and TOBSto.FieldExists('GQ_PMAP')
      then Prix := Arrondi(TOBSto.GetValue('GQ_PMAP') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'SMR') and (TOBSto <> nil) and TOBSto.FieldExists('GQ_PMRP')
      then Prix := Arrondi(TOBSto.GetValue('GQ_PMRP') / RatioStock, V_PGI.OkDecV) else
      { Pour mémoire, prix de la fiche article : DPA, PMAP, DPR, PMRP
      if ( CodeP = 'ADA' ) then Prix := Arrondi ( TOBArt.GetValue ('GA_DPA') / RatioDoc, V_PGI.OkDecV ) else
      if ( CodeP = 'ADR' ) then Prix := Arrondi ( TOBArt.GetValue ('GA_DPR') / RatioDoc, V_PGI.OkDecV ) else
      if ( CodeP = 'AMA' ) then Prix := Arrondi ( TOBArt.GetValue ('GA_PMAP') / RatioDoc, V_PGI.OkDecV ) else
      if ( CodeP = 'AMR' ) then Prix := Arrondi ( TOBArt.GetValue ('GA_PMRP') / RatioDoc, V_PGI.OkDecV ) else
      }
      if (CodeP = 'DDA') and (TOBDD <> nil) then Prix := Arrondi(TOBDD.GetValue('GQ_DPA') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'DDR') and (TOBDD <> nil) then Prix := Arrondi(TOBDD.GetValue('GQ_DPR') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'DMA') and (TOBDD <> nil) then Prix := Arrondi(TOBDD.GetValue('GQ_PMAP') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'DMR') and (TOBDD <> nil) then Prix := Arrondi(TOBDD.GetValue('GQ_PMRP') / RatioStock, V_PGI.OkDecV) else
      if (CodeP = 'ASA') then Prix := Arrondi(TOBArt.GetValue('GA_PAHT') / RatioDoc, V_PGI.OkDecV) else
      if (CodeP = 'ASR') then Prix := Arrondi(TOBArt.GetValue('GA_PRHT') / RatioDoc, V_PGI.OkDecV) else
      if (CodeP = 'AVH') then
    begin
      if FactureHT
        then Prix := Arrondi(TOBArt.GetValue('GA_PVHT') / RatioDoc, V_PGI.OkDecV)
      else Prix := Arrondi(TOBArt.GetValue('GA_PVTTC') / RatioDoc, V_PGI.OkDecV);
    end else
      if (CodeP = 'AM1') then Prix := Arrondi(TOBArt.GetValue('GA_VALLIBRE1') / RatioDoc, V_PGI.OkDecV) else
      if (CodeP = 'AM2') then Prix := Arrondi(TOBArt.GetValue('GA_VALLIBRE2') / RatioDoc, V_PGI.OkDecV) else
      if (CodeP = 'AM3') then Prix := Arrondi(TOBArt.GetValue('GA_VALLIBRE3') / RatioDoc, V_PGI.OkDecV) else
      if (CodeP = 'CDA') and (TOBCat <> nil) and (TOBCat.FieldExists('GCA_DPA'))
      then Prix := TOBCat.GetValue('GCA_DPA') else
      if (CodeP = 'CAT') and (TOBCat <> nil) and (TOBCat.FieldExists('GCA_PRIXBASE'))
      then Prix := TOBCat.GetValue('GCA_PRIXBASE');

  end; // while ( ( Prix = 0.0 ) and ( iAppelPrix < 3 ) ) do

  if TOBDD <> nil then TOBDD.Free;
  if Prix <= 0.0 then Prix := PrixDefaut;
  Result := Prix;
end;

function QuelPrixBase(NaturePiece, Depot: string; TOBA, TOBL: TOB): double;
var CodeP: string;
  TOBD: TOB;
  Prix, RatioDoc, RatioStock: Double;
  DPACat, BaseCat: Double;
begin
  // Attention document TTC
  Result := 0;
  if TOBA = nil then Exit;
  CodeP := GetInfoParPiece(NaturePiece, 'GPP_APPELPRIX');
  Prix := 0;
  RatioDoc := GetRatio(TOBL, nil, trsDocument);
  RatioStock := GetRatio(TOBL, nil, trsStock);
{$IFDEF BTP}
  ReactualisePrPv (TOBA);
{$ENDIF}
  TOBD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
//  Calcul au Dernier prix achat
  if CodeP = 'DPA' then
  begin
    if TOBA.FieldExists('GCA_DPA') then DPACat := TOBA.GetValue('GCA_DPA') else DPACat := 0;
    if DPACat <> 0 then Prix := DPACat else
    begin
      if TOBD <> nil then
      begin
        Prix := Arrondi(TOBD.GetValue('GQ_DPA') / RatioStock, V_PGI.OkDecV);
        if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_DPA') / RatioDoc, V_PGI.OkDecV);
      end
      else Prix := Arrondi(TOBA.GetValue('GA_DPA') / RatioDoc, V_PGI.OkDecV);
    end;
  end else
//  Calcul au Dernier prix de revient
    if CodeP = 'DPR' then
  begin
    if TOBD <> nil then
    begin
      Prix := Arrondi(TOBD.GetValue('GQ_DPR') / RatioStock, V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_DPR') / RatioDoc, V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_DPR') / RatioDoc, V_PGI.OkDecV);
  end else
//  Calcul au Prix moyen achat pondéré
    if CodeP = 'PPA' then
  begin
    if TOBD <> nil then
    begin
      Prix := Arrondi(TOBD.GetValue('GQ_PMAP') / RatioStock, V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_PMAP') / RatioDoc, V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_PMAP') / RatioDoc, V_PGI.OkDecV);
  end else
//  Calcul au Prix moyen revient pondéré
    if CodeP = 'PPR' then
  begin
    if TOBD <> nil then
    begin
      Prix := Arrondi(TOBD.GetValue('GQ_PMRP') / RatioStock, V_PGI.OkDecV);
      if (Prix <= 0) then Prix := Arrondi(TOBA.GetValue('GA_PMRP') / RatioDoc, V_PGI.OkDecV);
    end
    else Prix := Arrondi(TOBA.GetValue('GA_PMRP') / RatioDoc, V_PGI.OkDecV);
  end else
//  Calcul au Prix achat standard
    if CodeP = 'PAS' then Prix := Arrondi(TOBA.GetValue('GA_PAHT') / RatioDoc, V_PGI.OkDecV) else
//  Calcul au Prix revient standard
    if CodeP = 'PRS' then Prix := Arrondi(TOBA.GetValue('GA_PRHT') / RatioDoc, V_PGI.OkDecV) else
//  Calcul au Prix base catalogue
    if CodeP = 'CAT' then
  begin
    if TOBA.FieldExists('GCA_PRIXBASE') then
      BaseCat := TOBA.GetValue('GCA_PRIXBASE')
    else if TOBA.FieldExists('GCA_DPA') then
      BaseCat := TOBA.GetValue('GCA_DPA')
    else
      BaseCat := 0;
    if BaseCat <> 0 then Prix := BaseCat else Prix := Arrondi(TOBA.GetValue('GA_DPA') / RatioDoc, V_PGI.OkDecV);
  end else
//  Calcul au Prix vente HT
    if CodeP = 'PUH' then
  begin
{$IFDEF BTP}
		ReactualisePv (TOBA);
{$ENDIF}
//    if TOBL.GetValue('GL_FACTUREHT') = 'X' then Prix := Arrondi(TOBA.GetValue('GA_PVHT') / RatioDoc, V_PGI.OkDecV)
//    else Prix := Arrondi(TOBA.GetValue('GA_PVTTC') / RatioDoc, V_PGI.OkDecV);
    if TOBL.GetValue('GL_FACTUREHT') = 'X' then Prix := Arrondi(TOBA.GetValue('GA_PVHT') / RatioDoc, V_PGI.OkDecP)
    else Prix := Arrondi(TOBA.GetValue('GA_PVTTC') / RatioDoc, V_PGI.OkDecP);
  end else
//  Calcul au Montant libre 1
    if CodeP = 'MT1' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE1') / RatioDoc, V_PGI.OkDecV) else
//  Calcul au Montant libre 2
    if CodeP = 'MT2' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE2') / RatioDoc, V_PGI.OkDecV) else
//  Calcul au Montant libre 3
    if CodeP = 'MT3' then Prix := Arrondi(TOBA.GetValue('GA_VALLIBRE3') / RatioDoc, V_PGI.OkDecV);
  if (Prix <= 0) and (CodeP <> 'PUH') then // DC - Valeur par défaut - Fonctionnement Mode généralisé
  begin
    if (CodeP = 'DPR') or (CodeP = 'PPR')
      then Prix := Arrondi(TOBA.GetValue('GA_PRHT') / RatioDoc, V_PGI.OkDecV) //JS 23/07/03 F10611
    else if GetInfoParPiece(NaturePiece, 'GPP_GESTIONGRATUIT') <> 'X' then Prix := Arrondi(TOBA.GetValue('GA_PAHT') / RatioDoc, V_PGI.OkDecV);
  end; // DC Fin
  Result := Prix;
end;

procedure CalculeSousTotauxPiece(TOBPiece: TOB);
var i, iTot, Niv: integer;
  TOBL: TOB;
  TypeL: string;
begin
  iTot := 0;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i = 0 then iTot := TOBL.GetNumChamp('GL_TYPELIGNE');
    if TOBL.GetValeur(iTot) = 'TOT' then SommerLignes(TOBPiece, i + 1, 0);
    TypeL := TOBL.GetValeur(iTot);
    // VARIANTE
    if ISFinParagraphe (TOBL) then
    (* if Copy(TypeL, 1, 2) = 'TP' then *)
    begin
      Niv := StrToInt(Copy(TypeL, 3, 1));
      SommerLignes(TOBPiece, i + 1, Niv);
    end;
  end;
end;

procedure PieceVersLigneExi(TOBPiece, TOBL: TOB);
var StE: string;
  i: integer;
begin
  {#TVAENC}
  if TOBPiece.GetValue('GP_TVAENCAISSEMENT') = 'TE' then StE := 'X' else StE := '-';
  if TOBL = nil then
  begin
    for i := 0 to TOBPiece.Detail.Count - 1 do TOBPiece.Detail[i].PutValue('GL_TVAENCAISSEMENT', StE);
  end else
  begin
    TOBL.PutValue('GL_TVAENCAISSEMENT', StE);
  end;
end;

procedure PieceVersLigneRessource(TOBPiece, TOBL: TOB; OldRes: string);
var NewRes: string;
  i: integer;
begin
  NewRes := TOBPiece.GetValue('GP_RESSOURCE');
  if NewRes = '' then Exit;

  if TOBL = nil then
  begin
    for i := 0 to TOBPiece.Detail.Count - 1 do
    begin
      if (TOBPiece.Detail[i].getValue('GL_RESSOURCE') = '') or (TOBPiece.Detail[i].getValue('GL_RESSOURCE') = OldRes) then
        TOBPiece.Detail[i].PutValue('GL_RESSOURCE', NewRes);
    end;
  end else
  begin
    if (TOBL.GetValue('GL_RESSOURCE') = '') or (TOBL.GetValue('GL_RESSOURCE') = OldRes) then
      TOBL.PutValue('GL_RESSOURCE', NewRes);
  end;
end;

procedure MemoriseNumLigne(TobPiece : TOB);
var ind, inl, ino: integer;
  TOBL: TOB;
begin
  inl := 0;  ino := 0;
  for ind := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, ind + 1);
    if ind = 0 then
    begin
      inl := TOBL.GetNumChamp('GL_NUMLIGNE');
      TOBL.AddChampSup('ANCIENNUMLIGNE', True);
      ino := TOBL.GetNumChamp('GL_NUMORDRE');     //Affaire-ONYX
      TOBL.AddChampSup('ANCIENNUMORDRE', True);
    end;
    TOBL.PutValue('ANCIENNUMLIGNE', TOBL.GetValeur(inl));
    TOBL.PutValue('ANCIENNUMORDRE', TOBL.GetValeur(ino));
  end;
end;

procedure PieceVersLigne(TOBPiece, TOBL: TOB; WithAffaire : boolean=true; WithDomaine : boolean=true);
begin
  TOBL.PutValue('GL_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
  TOBL.PutValue('GL_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TOBL.PutValue('GL_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
  if TOBPiece.GetValue('GP_NATUREPIECEG') = 'TRE' then TOBL.PutValue('GL_DEPOT', TOBPiece.GetValue('GP_DEPOTDEST'))
  else TOBL.PutValue('GL_DEPOT', TOBPiece.GetValue('GP_DEPOT'));
  TOBL.PutValue('GL_TIERS', TOBPiece.GetValue('GP_TIERS'));
  TOBL.PutValue('GL_TIERSFACTURE', TOBPiece.GetValue('GP_TIERSFACTURE'));
  TOBL.PutValue('GL_TIERSLIVRE', TOBPiece.GetValue('GP_TIERSLIVRE'));
  TOBL.PutValue('GL_TIERSPAYEUR', TOBPiece.GetValue('GP_TIERSPAYEUR'));
  TOBL.PutValue('GL_TARIFSPECIAL', TOBPiece.GetValue('GP_TARIFSPECIAL')); 
  TOBL.PutValue('GL_TARIFTIERS', TOBPiece.GetValue('GP_TARIFTIERS')); //mcd 20/01/03 oubli ...
  TOBL.PutValue('GL_DATEPIECE', TOBPiece.GetValue('GP_DATEPIECE'));
  TOBL.PutValue('GL_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TOBL.PutValue('GL_ETABLISSEMENT', TOBPiece.GetValue('GP_ETABLISSEMENT'));
  TOBL.PutValue('GL_FACTUREHT', TOBPiece.GetValue('GP_FACTUREHT'));
  TOBL.PutValue('GL_DEVISE', TOBPiece.GetValue('GP_DEVISE'));
  TOBL.PutValue('GL_TAUXDEV', TOBPiece.GetValue('GP_TAUXDEV'));
  TOBL.PutValue('GL_COTATION', TOBPiece.GetValue('GP_COTATION'));
  TOBL.PutValue('GL_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TOBL.PutValue('GL_REGIMETAXE', TOBPiece.GetValue('GP_REGIMETAXE'));
  TOBL.PutValue('GL_REPRESENTANT', TOBPiece.GetValue('GP_REPRESENTANT'));
  TOBL.PutValue('GL_APPORTEUR', TOBPiece.GetValue('GP_APPORTEUR'));
  TOBL.PutValue('GL_VIVANTE', TOBPiece.GetValue('GP_VIVANTE'));
  TOBL.PutValue('GL_ESCOMPTE', TOBPiece.GetValue('GP_ESCOMPTE'));
  TOBL.PutValue('GL_REMISEPIED', TOBPiece.GetValue('GP_REMISEPIED'));
  TOBL.PutValue('GL_SAISIECONTRE', TOBPiece.GetValue('GP_SAISIECONTRE'));
  if (WithDomaine) and (TOBPiece.GetValue('GP_DOMAINE')<>'') then
  begin
  	TOBL.PutValue('GL_DOMAINE', TOBPiece.GetValue('GP_DOMAINE'));
  end;
  {Affaires}
  if (WithAffaire) and (TOBPiece.GetValue('GP_AFFAIRE')<>'') then
  begin
    TOBL.PutValue('GL_AFFAIRE', TOBPiece.GetValue('GP_AFFAIRE'));
    TOBL.PutValue('GL_AFFAIRE1', TOBPiece.GetValue('GP_AFFAIRE1'));
    TOBL.PutValue('GL_AFFAIRE2', TOBPiece.GetValue('GP_AFFAIRE2'));
    TOBL.PutValue('GL_AFFAIRE3', TOBPiece.GetValue('GP_AFFAIRE3'));
    TOBL.PutValue('GL_AVENANT', TOBPiece.GetValue('GP_AVENANT'));
  end;
  {#TVAENC}
  PieceversLigneExi(TOBPiece, TOBL);
  {Ressource}
  PieceVersLigneRessource(TOBPiece, TOBL, '');
end;

// DC - Valeurs par défaut

function AffectePrixDefaut(Prix, PrixDef1, PrixDef2: double): double;
begin
  if Prix <= 0 then
  begin
    Prix := PrixDef1;
    if Prix <= 0 then Prix := PrixDef2;
  end;
  Result := Prix;
end;
// DC Fin

procedure AffectePrixValo(TOBL, TOBA: TOB; TOBOuvrage: TOB = nil);
var TOBD: TOB;
  Depot: string;
  OkDepot: boolean;
  VenteAchat,Ua: string;
  FUS,FUV ,Pmap: double;
  {$IFDEF BTP}
//  Prx : double;
  Fournisseur,Article : string;
//  TOBPiece : TOB;
  COEFFIXE,CoefPAPr,CoefUaUs : double;
  {$ENDIF}
  DEV : Rdevise;
begin

	CoefuaUs := 0;

  if TOBL = nil then Exit;

  DEV.Code := TOBL.GetValue('GL_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux  := TOBL.GetValue('GL_TAUXDEV');

  if TOBA = nil then Exit;

  COEFFIXE := 0;

  VenteAchat := GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');

  // Modif BTP
  if TOBL.GetValue('GL_TYPEARTICLE') = 'NOM' then Exit;

  if ((TOBL.GetValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP')) and (ApplicPrixPose (TOBL)) then
  begin
  	if TobOuvrage = nil then exit;
    AffectePrixValoBTP (TOBL,TOBOuvrage);
    exit;
  end;

  if (TOBL.GetValue('GL_TYPEARTICLE')='MAR') or
  	 ((TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') and (not ApplicPrixPose (TOBL))) then
  begin
    if VenteAchat = 'VEN' then
    begin
      Fournisseur := TOBL.GetValue('GL_FOURNISSEUR');
      Article := TOBL.GetValue('GL_ARTICLE');
      RecupTarifAch (Fournisseur,Article,Ua,CoefuaUs,TurVente,true,True,TOBA);
      TOBL.PutValue('GL_COEFCONVQTE',CoefUaus);
    end;
  end;

  //-----
  OkDepot := False;

  TOBD := nil;

  if VH_GC.GCValoArtDepot then
  begin
    Depot := TOBL.GetValue('GL_DEPOT');
    TOBD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
    if TOBD <> nil then OkDepot := True;
  end;

  if OkDepot then
  begin
    TOBL.PutValue('GL_DPA', TOBD.GetValue('GQ_PMAP'));
    TOBL.PutValue('GL_PMAP', TOBD.GetValue('GQ_PMAP'));
    TOBL.PutValue('GL_PMAPACTU', TOBD.GetValue('GQ_PMAP'));
    TOBL.PutValue('GL_PMRP', TOBD.GetValue('GQ_PMRP'));
    TOBL.PutValue('GL_PMRPACTU', TOBD.GetValue('GQ_PMRP'));
  end else
  begin
    {$IFDEF BTP}
    if (TOBL.GetValue('GL_TENUESTOCK')='X') and (pos(TOBL.GetValue('GL_NATUREPIECEG'),'BLC;LBT')>0) then
    begin
      Depot := TOBL.GetValue('GL_DEPOT');
      TOBD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
      if (TOBD <> nil) and (TOBD.getValue('GQ_PMAP') <> 0) then OkDepot := true;
    end ;
    if VenteAchat = 'VEN' then
    begin
    	if OkDepot then
      begin
        FUV := RatioMesure('PIE', TobA.getValue('GA_QUALIFUNITEVTE'));
        FUS := RatioMesure('PIE', TobA.getValue('GA_QUALIFUNITESTO'));

        if TOBA.GetDouble('GA_COEFCONVQTEVTE')<> 0 then
        	Pmap := TOBD.GetValue('GQ_PMAP') / TOBA.GetDouble('GA_COEFCONVQTEVTE')
        else
        	Pmap := (TOBD.GetValue('GQ_PMAP') / FUS)*FUV;
      	TOBL.PutValue('GL_DPA', Pmap);
      	TOBA.PutValue('GA_PAHT', Pmap);
      	TOBA.PutValue('GA_DPA', Pmap);
        TOBL.PutValue('GL_PMAP', TOBL.GetValue('GL_DPA'));
      end else
      begin
				if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
        begin
        	if TOBA.GetValue('GA_PMAP') <> 0 then
          begin
						TOBL.PutValue('GL_DPA', TOBA.GetValue('GA_PMAP'));
						ReactualisePrPv (TOBA);
          end else
          begin
          	if TOBA.GetValue('GA_PAHT') <> 0 then  TOBL.PutValue('GL_DPA', TOBA.GetValue('GA_PAHT'));
          end;
        end else
        begin
        	if (TOBL.GetDouble('GL_DPA')=0) and (TOBA.GetValue('GA_PAHT') <> 0) then TOBL.PutValue('GL_DPA', TOBA.GetValue('GA_PAHT'));
        end;
      end;
    end
    else
    begin
      TOBL.PutValue('GL_DPA', TOBA.GetValue('GA_DPA'));
    end;
    {$ELSE}
    TOBL.PutValue('GL_DPA', TOBA.GetValue('GA_DPA'));
    {$ENDIF}
    //ou les coef de la ligne sont à zéro... (before : if (TOBL.GetValue('GL_DOMAINE')='') or (TOBA.GetValue('GA_PRIXPASMODIF')='X') then)
    if (TOBL.GetValue('GL_DOMAINE')='') or (TOBA.GetValue('GA_PRIXPASMODIF')='X') then
    begin
      ReactualisePr (TOBA);
      if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
      begin
        if TOBA.GetValue('GA_PMAP') <> 0 then CoefPapr := TOBA.GetValue('GA_DPR')/TOBA.GetValue('GA_PMAP')
                                         else CoefPapr := 1;
      end
      else
      begin
        //FV1 : 02/03/2015 - Gestion des Prix d'achat dans fiche article
        if TOBA.GetValue('GA_PAHT') <> 0 then
          CoefPapr := TOBA.GetValue('GA_DPR')/TOBA.GetValue('GA_PAHT') //TOBA.GetValue('GA_PAHT')
        else
          CoefPapr := 1;
      end;
      TOBL.PutValue('GL_DPR', TOBL.GetValue('GL_DPA')*CoefPaPr);
    end else
    begin
			CalculeLigneAc (TOBL,DEV);
    end;
    TOBL.PutValue('GL_PMAP', TOBL.GetValue('GL_DPR'));
    TOBL.PutValue('GL_PMAPACTU', TOBL.GetValue('GL_DPR'));
    TOBL.PutValue('GL_PMRP', TOBL.GetValue('GL_DPR'));
    TOBL.PutValue('GL_PMRPACTU', TOBL.GetValue('GL_DPR'));
  end;

{$IFNDEF BTP}
  // DC - Valeurs par défaut
  TOBL.PutValue('GL_DPA', AffectePrixDefaut (TOBL.GetValue('GL_DPA'),     TOBA.GetValue('GA_DPA'),  TOBA.GetValue('GA_PAHT')));
  TOBL.PutValue('GL_DPR', AffectePrixDefaut (TOBL.GetValue('GL_DPR'),     TOBA.GetValue('GA_DPR'),  TOBA.GetValue('GA_PRHT')));
  TOBL.PutValue('GL_PMAP', AffectePrixDefaut(TOBL.GetValue('GL_PMAP'), TOBA.GetValue('GA_PMAP'), TOBA.GetValue('GA_PAHT')));
  if TOBL.GetValue('GL_PMAPACTU') <= 0 then TOBL.PutValue('GL_PMAPACTU', TOBL.GetValue('GL_PMAP'));
  TOBL.PutValue('GL_PMRP', AffectePrixDefaut(TOBL.GetValue('GL_PMRP'), TOBA.GetValue('GA_PMRP'), TOBA.GetValue('GA_PRHT')));
  if TOBL.GetValue('GL_PMRPACTU') <= 0 then TOBL.PutValue('GL_PMRPACTU', TOBL.GetValue('GL_PMRP'));
  // DC Fin
{$ELSE}
  if (TOBL.getValue('GL_DOMAINE') = '') or (TOBA.GetValue('GA_PRIXPASMODIF')='X') then
  begin
    if TOBL.GetValue('GL_DPR')=0    then TOBL.PutValue('GL_DPR',    TOBL.GetValue('GL_DPA')*TOBA.GetValue('GA_COEFFG'));
    if TOBL.GetValue('GL_DPA') <> 0 then TOBL.PutValue('GL_COEFFG',(TOBL.GetValue('GL_DPR')-TOBL.GetValue('GL_DPA'))/TOBL.GetValue('GL_DPA'));
  end;
{$ENDIF BTP}

end;

procedure AffecteValoNomen(TOBN, TOBA: TOB; Depot: string);
var TOBD: TOB;
  OkDepot: boolean;
begin
  if TOBN = nil then Exit;
  if TOBA = nil then Exit;
  OkDepot := False;
  TOBD := nil;
  if VH_GC.GCValoArtDepot then
  begin
    TOBD := TOBA.FindFirst(['GQ_DEPOT'], [Depot], False);
    if TOBD <> nil then OkDepot := True;
  end;
  if OkDepot then
  begin
    TOBN.PutValue('DPA', TOBD.GetValue('GQ_PMAP'));
    if VH_GC.GCIfDefCEGID then
       TOBN.PutValue('DPR', TOBA.GetValue('GA_DPR'))
    else
       TOBN.PutValue('DPR', TOBD.GetValue('GQ_DPR'));
    TOBN.PutValue('PMAP', TOBD.GetValue('GQ_PMAP'));
    TOBN.PutValue('PMRP', TOBD.GetValue('GQ_PMRP'));
  end else
  begin
    TOBN.PutValue('DPA', TOBA.GetValue('GA_DPA'));
    TOBN.PutValue('DPR', TOBA.GetValue('GA_DPR'));
    TOBN.PutValue('PMAP', TOBA.GetValue('GA_PMAP'));
    TOBN.PutValue('PMRP', TOBA.GetValue('GA_PMRP'));
  end;
end;

function GetLibTrad(RefArt, Lgu: string): string;
var Q: TQuery;
begin
  Result := '';
  if RefArt = '' then Exit;
  if ((Lgu = '') or (Lgu = V_PGI.LanguePrinc)) then Exit;
  Q := OpenSQL('SELECT GTA_LIBELLE FROM GTRADARTICLE WHERE GTA_ARTICLE="' + RefArt + '" AND GTA_LANGUE="' + Lgu + '"', True,-1, '', True);
  if not Q.EOF then Result := Q.Fields[0].AsString;
  Ferme(Q);
end;

procedure ArticleVersLigne(TOBPiece, TOBA, TOBConds, TOBL, TOBTiers: TOB; AvecPrix: Boolean = true);
var NaturePiece, Depot: string;
  Prix, PCB, Qte: Double;
  i, j: integer;
  TOBC: TOB;
  RefUnique, LguTiers, LibA, VenteAchat: string;
begin
  if TOBA = nil then Exit;
  LibA := '';
  RefUnique := TOBA.GetValue('GA_ARTICLE');

  if RefUnique <> '' then
  begin
    VenteAchat := GetInfoParPiece(TOBL.GetValue('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');

    if (TOBA.FieldExists('GCA_LIBELLE')) and (TOBL.GetValue('GL_REFARTSAISIE') <> TOBL.GetValue('GL_CODEARTICLE')) then
    begin
      LibA := TOBA.GetVAlue('GCA_LIBELLE');
      if LibA = '' then LibA := TOBA.GetValue('GA_LIBELLE');
    end
    else
    begin
      { Permet d'appeler cette procédure en passant la TobTiers à nil }
      // Avant : LguTiers:=TOBTiers.GetValue('T_LANGUE') ;
      if Assigned(TOBTiers) then
        LguTiers := TOBTiers.GetValue('T_LANGUE')
      else
        LguTiers := V_PGI.LanguePrinc;
      {Fin GPAO}
      if ((LguTiers <> '') and (LguTiers <> V_PGI.LanguePrinc)) then
      begin
        LibA := GetLibTrad(RefUnique, LguTiers);
        if LibA = '' then LibA := TOBA.GetValue('GA_LIBELLE');
      end
      else LibA := TOBA.GetValue('GA_LIBELLE');

      end;

    TOBL.PutValue('GL_LIBELLE', LibA);
    //
    TOBL.putValue('BNP_TYPERESSOURCE',TOBA.GetValue('BNP_TYPERESSOURCE'));
    //
    if VenteAchat = 'VEN' then
    begin
{$IFDEF BTP}
			InitValoArtNomen (TOBA);
{$ENDIF}
      TOBL.PutValue('GL_PRIXPOURQTE', TOBA.GetValue('GA_PRIXPOURQTE'));
      PCB := TOBA.GetValue('GA_PCB');
      TOBL.PutValue('GL_PCB', PCB);

      Qte := AjusteQte_MiniMultiple(TobL.GetValue('GL_QTESTOCK'), TobA.getValue('GA_QECOVTE'), TobA.GetValue('GA_PCB'));

      TOBL.PutValue('GL_QTESTOCK', Qte);
      TOBL.PutValue('GL_QTERESTE', Qte);
      TOBL.PutValue('GL_QTEFACT', Qte);
      if TOBA.GetBoolean('GA_RELIQUATMT') then TOBL.PutValue('GL_MTRESTE', Qte * TOBL.GetValue('GL_PUHTDEV'));
    end
    else if VenteAchat = 'ACH' then
    begin
      TOBL.PutValue('GL_PRIXPOURQTE', TOBA.GetValue('GA_PRIXPOURQTEAC'));

      Qte := AjusteQte_MiniMultiple(TobL.GetValue('GL_QTESTOCK'), TobA.getValue('GA_QECOACH'), TobA.GetValue('GA_QPCBACH'));

      TOBL.PutValue('GL_QTESTOCK', Qte);
      TOBL.PutValue('GL_QTERESTE', Qte);
      TOBL.PutValue('GL_QTEFACT' , Qte);
      if TOBA.GetBoolean('GA_RELIQUATMT') then TOBL.PutValue('GL_MTRESTE', Qte * TOBL.GetValue('GL_PUHTDEV'));
    end
    else TOBL.PutValue('GL_PRIXPOURQTE', 1);

    if TOBL.GetValue('GL_PRIXPOURQTE') < 1 then TOBL.PutValue('GL_PRIXPOURQTE', 1);

    TOBL.PutValue('GL_LIBCOMPL', TOBA.GetValue('GA_LIBCOMPL'));
    TOBL.PutValue('GL_REFARTBARRE', TOBA.GetValue('GA_CODEBARRE'));
    TOBL.PutValue('GL_ESCOMPTABLE', TOBA.GetValue('GA_ESCOMPTABLE'));
    TOBL.PutValue('GL_REMISABLEPIED', TOBA.GetValue('GA_REMISEPIED'));

    if IsPieceAutoriseRemligne (TOBPiece) then
      TOBL.PutValue('GL_REMISABLELIGNE',TOBA.GetValue('GA_REMISELIGNE'))
    else
      TOBL.PutValue('GL_REMISABLELIGNE', '-');
    //
    //
    TOBL.PutValue('GL_TENUESTOCK', TOBA.GetValue('GA_TENUESTOCK'));
    TOBL.PutValue('GL_TARIFARTICLE', TOBA.GetValue('GA_TARIFARTICLE'));
    TOBL.PutValue('GL_TYPELIGNE', 'ART');
    TOBL.PutValue('GL_NUMEROSERIE', TOBA.GetValue('GA_NUMEROSERIE'));
    TOBL.PutValue('GL_PAYSORIGINE', TOBA.GetValue('GA_PAYSORIGINE'));  //JS FI10226
    TOBL.PutValue('GL_CODEDOUANE', TOBA.GetValue('GA_CODEDOUANIER'));
    {Familles, collection, domaine}
    TOBL.PutValue('GL_FAMILLENIV1', TOBA.GetValue('GA_FAMILLENIV1'));
    TOBL.PutValue('GL_FAMILLENIV2', TOBA.GetValue('GA_FAMILLENIV2'));
    TOBL.PutValue('GL_FAMILLENIV3', TOBA.GetValue('GA_FAMILLENIV3'));
    TOBL.PutValue('GL_COLLECTION', TOBA.GetValue('GA_COLLECTION'));

    if TOBPiece.GetValue('GP_DOMAINE')='' then
    begin
    	TOBL.PutValue('GL_DOMAINE', TOBA.GetValue('GA_DOMAINE'));
    end;
    {Nomenclature}
    TOBL.PutValue('GL_TYPEARTICLE', TOBA.GetValue('GA_TYPEARTICLE'));
    TOBL.PutValue('GL_TYPENOMENC', TOBA.GetValue('GA_TYPENOMENC'));

    NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
    Depot := TOBL.GetValue('GL_DEPOT');

    for i := 1 to 5 do TOBL.PutValue('GL_FAMILLETAXE' + IntToStr(i), TOBA.GetValue('GA_FAMILLETAXE' + IntToStr(i)));
    {Unités de mesure}
    TOBL.PutValue('GL_QUALIFSURFACE', TOBA.GetValue('GA_QUALIFSURFACE'));
    TOBL.PutValue('GL_QUALIFVOLUME', TOBA.GetValue('GA_QUALIFVOLUME'));
    TOBL.PutValue('GL_QUALIFPOIDS', TOBA.GetValue('GA_QUALIFPOIDS'));
    TOBL.PutValue('GL_QUALIFLINEAIRE', TOBA.GetValue('GA_QUALIFLINEAIRE'));
    TOBL.SetDouble('GL_QTESAIS',TOBL.GetDouble('GL_QTEFACT'));
    if GetParamSocSecur('SO_BTGESTQTEAVANC',false) and (ISDocumentChiffrage(TOBPiece)) then
    begin
      if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
      begin
        TOBL.SetString('GL_QUALIFHEURE',     TOBA.GetString('GA_QUALIFHEURE'));
        TOBL.SetDouble('GL_RENDEMENT',       TOBA.GetDouble('GA_COEFPROD'));
        TOBL.SetDOuble('GL_PERTE',0);
      end else
      begin
        TOBL.SetString('GL_QUALIFHEURE','');
        TOBL.SetDouble('GL_PERTE',           TOBA.GetDouble('GA_PERTEPROP'));
        TOBL.SetDouble('GL_RENDEMENT',0);
      end;
      TOBL.SetDouble('GL_QTESTOCK',TOBL.GetDouble('GL_QTEFACT'));
      TOBL.SetDouble('GL_QTERESTE',TOBL.GetDouble('GL_QTEFACT'));
    end;
    TOBL.SetDouble('GL_QTEFACT',CalculeQteFact (TOBL));
    TOBL.PutValue('GL_SURFACE', TOBA.GetValue('GA_SURFACE'));
    TOBL.PutValue('GL_VOLUME', TOBA.GetValue('GA_VOLUME'));
    TOBL.PutValue('GL_POIDSBRUT', TOBA.GetValue('GA_POIDSBRUT'));
    TOBL.PutValue('GL_POIDSNET', TOBA.GetValue('GA_POIDSNET'));
    TOBL.PutValue('GL_POIDSDOUA', TOBA.GetValue('GA_POIDSDOUA'));
    TOBL.PutValue('GL_LINEAIRE', TOBA.GetValue('GA_LINEAIRE'));
    TOBL.PutValue('GL_HEURE', TOBA.GetValue('GA_HEURE'));
    TOBL.PutValue('GL_QUALIFQTESTO', TOBA.GetValue('GA_QUALIFUNITESTO'));
    TOBL.PutValue('GL_COEFCONVQTE', TOBA.GetValue('GA_COEFCONVQTEACH'));
    TOBL.PutValue('GL_COEFCONVQTEVTE', TOBA.GetValue('GA_COEFCONVQTEVTE'));
    TOBL.PutValue('GL_QUALIFQTEVTE', TOBA.GetValue('GA_QUALIFUNITEVTE'));
    TOBL.PutValue('GL_QUALIFQTEACH',    TOBA.GetValue('GA_QUALIFUNITEACH'));
    {Tables libres}
    for j := 1 to 9 do TOBL.PutValue('GL_LIBREART' + IntToStr(j), TOBA.GetValue('GA_LIBREART' + IntToStr(j)));
    TOBL.PutValue('GL_LIBREARTA', TOBA.GetValue('GA_LIBREARTA'));
    {$IFNDEF BTP}
    {Affaires}
    if ctxAffaire in V_PGI.PGIContexte then TOBL.PutValue('GL_BLOCNOTE', TOBA.GetValue('GA_COMMENTAIRE'));
    {$ENDIF}

    {Prix}
    if AvecPrix then
    begin
      Prix := QuelPrixBase(NaturePiece, Depot, TOBA, TOBL);
      if VenteAchat = 'ACH' then
      begin
        Prix := Prix * TOBL.GetValue('GL_PRIXPOURQTE');
      end;
      if TOBL.GetValue('GL_FACTUREHT') = 'X' then
      begin
        TOBL.PutValue('GL_PUHTDEV', Prix);
      end else
      begin
        TOBL.PutValue('GL_PUTTCDEV', Prix);
      end;
    end;

    {$IFDEF BTP}
    if (VenteAchat = 'VEN') and (TOBA.GetValue('GA_FOURNPRINC') <> '') then
    begin
      TOBL.PutValue('GL_FOURNISSEUR', TOBA.GetValue('GA_FOURNPRINC'));
    end;
    {$ENDIF}

    AffectePrixValo(TOBL, TOBA);

    //TOBL.PutValue('GL_QUALIFQTEACH', TOBA.GetValue('GA_QUALIFUNITEACH'));

    {Conditionnements}
    (*
    if TOBConds <> nil then
    begin
      TOBC := TOBConds.FindFirst(['GCO_ARTICLE'], [RefUnique], False);
      if TOBC <> nil then
      begin
        TOBL.PutValue('GL_CODECOND', TOBC.GetValue('GCO_CODECOND'));
        TOBL.PutValue('GL_QTESTOCK', TOBC.GetValue('GCO_NBARTICLE'));
        TOBL.PutValue('GL_QTEFACT', TOBC.GetValue('GCO_NBARTICLE'));
        { TP - NEWPIECE - Bug 10709 }
        TOBL.PutValue('GL_QTERESTE', TOBC.GetValue('GCO_NBARTICLE'));
      end;
    end;
    *)
    InitConditionnements (TOBL);

    {Bloc-Notes}
    if GetInfoParPiece(NaturePiece, 'GPP_BLOBART') = 'X' then TOBL.PutValue('GL_BLOCNOTE', TOBA.GetValue('GA_BLOCNOTE'));

    {Divers}
    if TOBL.getValue('GL_NIVEAUIMBRIC') = 0 then TOBL.PutValue('GL_DATELIVRAISON', TOBPiece.GetValue('GP_DATELIVRAISON'));
    TOBL.PutValue('GL_CAISSE', TOBPiece.GetValue('GP_CAISSE'));

    {Mode}
    if (ctxMode in V_PGI.PGIContexte) and ((VenteAchat = 'VEN') or (TOBL.GetValue('GL_NATUREPIECEG') = 'FCF')) then
      TOBL.PutValue('GL_FOURNISSEUR', TOBA.GetValue('GA_FOURNPRINC'));

    {$IFDEF CHR}
    TOBL.PutValue('GL_REGROUPELIGNE', TOBA.GetValue('GA_REGROUPELIGNE'));
    if (TOBA.GetValue('GA_REGROUPELIGNE') <> '') then
    begin
      if not TOBL.FieldExists('(REGROUPEMENT)') then
        TOBL.AddChampSup('(REGROUPEMENT)', False);
      TOBL.PutValue('(REGROUPEMENT)', RechDom('HRTREGROUPELIGNE', TOBA.GetValue('GA_REGROUPELIGNE'), False));
    end;
    if (TOBA.GetValue('GA_PRINCIPALEXTRA') = 'EXT') then
      TOBL.PutValue('GL_NOFOLIO', '2')
    else
      TOBL.PutValue('GL_NOFOLIO', '1');
    {$ENDIF}

    {$IFDEF BTP}
//    if ((TOBL.GetValue('GL_TYPEARTICLE') <> 'PRE') or (RenvoieTypeRes(TOBL.GetValue('GL_ARTICLE')) <> 'ST')) then
(*
    if ((TOBL.GetValue('GL_TYPEARTICLE') <> 'PRE') or (TOBL.GetValue('BNP_TYPERESSOURCE') <> 'ST')) then
    begin
    	TOBL.PutValue('TPSUNITAIRE', 1);
    end else
    begin
    	TOBL.PutValue('TPSUNITAIRE', 0);
    end;
*)
    {$ENDIF}
    TOBL.SetBoolean('GESTRELIQUAT',TOBA.GetBoolean('GA_RELIQUATMT'));
    // Modif BTP
    TOBL.PutValue('GL_RECALCULER', 'X');
  end;
end;

procedure CatalogueVersLigne(TOBPiece, TOBCata, TOBL, TOBTiers, TOBArtRef: TOB);
var NaturePiece, Depot, VenteAchat: string;
  Prix, PCB: Double;
  i, j: Integer;
begin

  if TOBCata = nil then EXIT;

  TOBL.PutValue('GL_TENUESTOCK', '-');
  TOBL.PutValue('GL_TYPELIGNE', 'ART');

  {Issu des lignes}
  NaturePiece := TOBL.GetValue('GL_NATUREPIECEG');
  Depot := TOBL.GetValue('GL_DEPOT');
  VenteAchat := GetInfoParPiece(NaturePiece, 'GPP_VENTEACHAT');

  {Issu du Catalogue}
  TOBL.PutValue('GL_LIBELLE', TOBCata.GetValue('GCA_LIBELLE'));
  for i := 1 to 5 do TOBL.PutValue('GL_FAMILLETAXE' + IntToStr(i), TOBCata.GetValue('GCA_FAMILLETAXE' + IntToStr(i)));

  TOBL.PutValue('GL_PAYSORIGINE', TOBCata.GetValue('GCA_PAYSORIGINE'));  //JS FI10226
  {Prix}
  if VenteAchat = 'VEN' then
    Prix := TOBCata.GetValue('GCA_PRIXVENTE')
  else
    Prix := TOBCata.GetValue('GCA_PRIXBASE');

  if TOBL.GetValue('GL_FACTUREHT') = 'X' then
  begin
    TOBL.PutValue('GL_PUHT', Prix);
    TOBL.PutValue('GL_PUHTDEV', Prix);
  end else
  begin
    TOBL.PutValue('GL_PUTTC', Prix);
    TOBL.PutValue('GL_PUTTCDEV', Prix);
  end;

  if TOBCata.GetValue('GCA_DPA') > 0 then
    TOBL.PutValue('GL_DPA', TOBCata.GetValue('GCA_DPA'))
  else
    TOBL.PutValue('GL_DPA', TOBCata.GetValue('GCA_PRIXBASE'));

  TOBL.PutValue('GL_QUALIFQTEACH', TOBCata.GetValue('GCA_QUALIFUNITEACH'));

  {Bloc-Notes}
  if GetInfoParPiece(NaturePiece, 'GPP_BLOBART') = 'X' then TOBL.PutValue('GL_BLOCNOTE', TOBCata.GetValue('GCA_BLOC_NOTE'));

  {Issu des pieces}
  {Divers}
  TOBL.PutValue('GL_DATELIVRAISON', TOBPiece.GetValue('GP_DATELIVRAISON'));
  TOBL.PutValue('GL_CAISSE', TOBPiece.GetValue('GP_CAISSE'));

  // Escomptable positionné à vrai par défaut
  TOBL.PutValue('GL_ESCOMPTABLE', 'X');
  TOBL.PutValue('GL_REMISABLEPIED', 'X');
  TOBL.PutValue('GL_REMISABLELIGNE', 'X');

  {Issu de l'article de reference}
  if TOBArtRef <> nil then
  begin
    PCB := TOBArtRef.GetValue('GA_PCB');
    TOBL.PutValue('GL_PCB', PCB);
    if PCB <> 0 then
    begin
      TOBL.PutValue('GL_QTESTOCK', PCB);
      TOBL.PutValue('GL_QTEFACT', PCB);
      TOBL.PutValue('GL_QTERESTE', PCB);
    end;
    // ContreM Traduction ??
    TOBL.PutValue('GL_ESCOMPTABLE', TOBArtRef.GetValue('GA_ESCOMPTABLE'));
    TOBL.PutValue('GL_REMISABLEPIED', TOBArtRef.GetValue('GA_REMISEPIED'));
    if IsPieceAutoriseRemligne (TOBPiece) then TOBL.PutValue('GL_REMISABLELIGNE', TOBArtRef.GetValue('GA_REMISELIGNE'))
    																			else TOBL.PutValue('GL_REMISABLELIGNE', '-');
    {Familles, collection}
    TOBL.PutValue('GL_FAMILLENIV1', TOBArtRef.GetValue('GA_FAMILLENIV1'));
    TOBL.PutValue('GL_FAMILLENIV2', TOBArtRef.GetValue('GA_FAMILLENIV2'));
    TOBL.PutValue('GL_FAMILLENIV3', TOBArtRef.GetValue('GA_FAMILLENIV3'));
    TOBL.PutValue('GL_COLLECTION', TOBArtRef.GetValue('GA_COLLECTION'));
    {Tables libres}
    for j := 1 to 9 do TOBL.PutValue('GL_LIBREART' + IntToStr(j), TOBArtRef.GetValue('GA_LIBREART' + IntToStr(j)));
    //   TOBL.PutValue('GL_LIBREARTA',TOBArtRef.GetValue('GA_LIBREARTA')) ;
  end;
end;


procedure InitialiseComm(TOBL: TOB; Aff: boolean);
var prefixe : string;
begin
  if TOBL.nomTable = 'LIGNE' Then prefixe := 'GL'
  else if TOBL.nomTable = 'LIGNEOUVPLAT' Then prefixe := 'BOP';
  TOBL.PutValue(prefixe+'_TYPECOM', '');
  TOBL.PutValue(prefixe+'_COMMISSIONR', 0);
  TOBL.PutValue(prefixe+'_COMMISSIONA', 0);
  if Aff then TOBL.PutValue(prefixe+'_VALIDECOM', 'AFF') else TOBL.PutValue(prefixe+'_VALIDECOM', 'NON');
end;

procedure InitialiseLigne(TOBL: TOB; ARow: integer; NewQte: double);
var prefixe : string;
//  TobPieceLocal, TobLigneTarifLocal: Tob;
begin
  if TOBL.nomTable = 'LIGNE' then prefixe:='GL'
  else if TOBL.nomTable = 'LIGNEOUVPLAT' Then prefixe := 'BOP';
  TOBL.PutValue(prefixe+'_NUMLIGNE', ARow);
  TOBL.PutValue('RECALCULTARIF', '-');
  TOBL.PutValue('RECALCULCOMM' , '-'); 
  TOBL.PutValue(prefixe+'_PRIXPOURQTE', 1);
  TOBL.PutValue(prefixe+'_CODECOND', '');
  TOBL.PutValue(prefixe+'_INDICENOMEN', 0);
  if Pos(TOBL.GetValue(prefixe+'_TYPELIGNE'),'SD;SDV') =0 then
  	TOBL.PutValue(prefixe+'_TYPELIGNE', 'COM');
  TOBL.PutValue(prefixe+'_TYPEDIM', 'NOR');
  TOBL.PutValue(prefixe+'_DATELIVRAISON', iDate1900);
  if prefixe = 'GL' then
  begin
    InitLesSupLigneCompl(TOBL, False);
  end;
  // Commissions
  InitialiseComm(TOBL, False);
  // Qtés
  if TOBL.GetValue(prefixe+'_ARTICLE') <> '' then
  begin
    TOBL.PutValue(prefixe+'_QTESTOCK', NewQte);
    TOBL.PutValue(prefixe+'_QTEFACT', NewQte);
    TOBL.PutValue(prefixe+'_QTERESTE', NewQte); { NEWPIECE }
    if prefixe = 'GL' then
    begin
      if not TOBL.FieldExists('QTECHANGE') then TOBL.AddChampSup('QTECHANGE', False);
      TOBL.PutValue('QTECHANGE', 'X');
      if not TOBL.FieldExists('SUPPRIME') then TOBL.AddChampSup('SUPPRIME', False);
      TOBL.PutValue('SUPPRIME', '-');
      TOBL.PutValue('GL_RECALCULER', 'X');
    end;
  end;
end;

procedure InitLigneVide(TOBPiece, TOBL, TOBTiers, TOBAffaire: TOB; ARow: integer; NewQte: double);
var Niv: integer;
  TypL: string;
  TheDateLiv : TdateTime;
  MaxNumOrdre : integer;
begin
  InitialiseLigne(TOBL, ARow, NewQte);
  // Modif BTP
  // initialisation du niveau d'imbrication de la ligne
  // en fonction de la ligne précédente sauf si dernière ligne.
  // zéro par défaut.
  if (ARow > 1) and (ARow < TOBPiece.Detail.count) then
  begin
    Niv := GetChampLigne(TOBPiece, 'GL_NIVEAUIMBRIC', ARow + 1);
    TypL := GetChampLigne(TOBPiece, 'GL_TYPELIGNE', ARow + 1);
    TheDateLiv := GetChampLigne(TOBPiece, 'GL_DATELIVRAISON', ARow + 1);

    // VARIANTE
    if IsDebutParagraphe (TOBPiece.detail[Arow]) then
    (* if (Copy(TypL, 1, 2) = 'DP') then *)
      Niv := Niv - 1;

    TOBL.PutValue('GL_NIVEAUIMBRIC', Niv);
    TOBL.PutValue('GL_DATELIVRAISON', TheDateLiv);
    // VARIANTE
    SetLigneCommentaire (TOBPiece,TOBL,Arow-1);
  end;
  // --------
  MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
  if TOBL.GetValue('GL_NUMORDRE')=0 then
  begin
    Inc(MaxNumOrdre);
    TOBL.PutValue('GL_NUMORDRE', MaxNumOrdre);
    EcrireMaxNumordre (TOBPiece,MaxNumOrdre);
  end;
  PieceVersLigne(TOBPiece, TOBL);
  AffaireVersLigne(TOBPiece, TOBL, TOBAffaire);
end;

function PreAffecteLigne(TobPiece, TobL, TobLigneTarif, TobA, TobTiers, TobTarif, TobConds, TobAffaire, TobCata, TobArtRef: Tob; ARow: integer; DEV: RDEVISE; NewQte: double; CalcTarif : boolean = True; CTRLCatFrs : boolean = True ): T_ActionTarifArt;
begin
//  Result:=ataOK ;
  InitLigneVide(TOBPiece, TOBL, TOBTiers, TobAffaire, ARow, NewQte);

  if TOBL.GetValue('GL_TYPEREF') = 'CAT' then
    CatalogueVersLigne(TOBPiece, TOBCata, TOBL, TOBTiers, TOBArtRef)
  else
    ArticleVersLigne(TOBPiece, TOBA, TOBConds, TOBL, TOBTiers);

  // VARIANTE
  if IsLigneInVariante (TOBpiece,TOBL,Arow) then TOBL.putValue('GL_TYPELIGNE','ARV');
  // --
  if CalcTarif and (not GetParamSoc('SO_PREFSYSTTARIF')) then
    Result := TarifVersLigne(TobA, TobTiers, TobL, TobligneTarif, TobPiece, TobTarif, False, True, DEV, CTRLCatFrs )
  else
    Result := ataOk;

  // si en achat -- > On recup la reference Article/fournisseur
end;

procedure MajCatalogueSaisie(TOBArticles: TOB);
var TOBA: TOB;
  i, j: integer;
  OldDP, NewDP: Double;
  TOBG, TOBC: TOB;
  Nam: string;
begin
  TOBG := TOB.Create('', nil, -1);
  for i := 0 to TOBArticles.Detail.Count - 1 do
  begin
    TOBA := TOBArticles.Detail[i];
    if TOBA.FieldExists('DPANEW') then
    begin
      OldDP := TOBA.GetValue('GCA_DPA');
      NewDP := TOBA.GetValue('DPANEW');
      if Arrondi(OldDP - NewDP, V_PGI.OkDecV) <> 0 then
      begin
        TOBC := TOB.Create('CATALOGU', TOBG, -1);
        for j := 1 to TOBC.NbChamps do
        begin
          Nam := TOBC.GetNomChamp(j);
          if TOBA.FieldExists(Nam) then TOBC.PutValue(Nam, TOBA.GetValue(Nam));
        end;
        TOBC.PutValue('GCA_DPA', NewDP);
      end;
    end;
  end;
  if TOBG.Detail.Count > 0 then TOBG.UpdateDB;
  TOBG.Free;
end;

function ConvertSaisieEF(XSais: Double; ModeOppose: boolean): Double;
begin
  if VH^.TenueEuro then
  begin
    if ModeOppose then Result := FrancToEuro(XSais) else Result := EuroToFranc(XSais);
  end else
  begin
    if ModeOppose then Result := EuroToFranc(XSais) else Result := FrancToEuro(XSais);
  end;
end;

function ConvertSaisieEFNA(XSais: Double; ModeOppose: boolean): Double;
begin
  if VH^.TenueEuro then
  begin
    if ModeOppose then Result := FrancToEuroNA(XSais) else Result := EuroToFrancNA(XSais);
  end else
  begin
    if ModeOppose then Result := EuroToFrancNA(XSais) else Result := FrancToEuroNA(XSais);
  end;
end;

function ZonesAchat(prefixe,Nam : string) : boolean;
var Zones : string;
begin
  if Prefixe = 'GL' then
  begin
	  Zones := 'GL_MONTANTFG;GL_MONTANTFR;GL_MONTANTFC;GL_MONTANTPAFR;GL_MONTANTPAFC;GL_MONTANTPAFG;GL_MONTANTPA;GL_MONTANTPR';
  end else
  begin
	  Zones := 'BOP_MONTANTFG;BOP_MONTANTFR;BOP_MONTANTFC;BOP_MONTANTPAFR;BOP_MONTANTPAFC;BOP_MONTANTPAFG;BOP_MONTANTPA;BOP_MONTANTPR';
  end;
	result := false;
	if pos (Nam,Zones) > 0 then result := true;
end;

procedure ZeroLigne(TOBL: TOB);
var i: integer;
  Nam,Prefixe: string;
begin
  prefixe := GetPrefixeTable (TOBL);
  for i := 1 to TOBL.NbChamps do
  begin
    Nam := TOBL.GetNomChamp(i);
    if prefixe = 'GL' then
    begin
      if Copy(Nam, 1, 6) = 'GL_TOT' then TOBL.PutValue(Nam, 0);
      if (Copy(Nam, 1, 10) = 'GL_MONTANT') and (not ZonesAchat(prefixe,Nam)) then TOBL.PutValue(Nam, 0);
    end else
    begin
      if Copy(Nam, 1, 7) = 'BOP_TOT' then TOBL.PutValue(Nam, 0);
      if (Copy(Nam, 1, 11) = 'BOP_MONTANT') and (not ZonesAchat(prefixe,Nam)) then TOBL.PutValue(Nam, 0);
    end;
  end;
  if not IsOuvrage(TOBL) then
  begin
    ZeroMtLigne(TOBL);
  end;
end;

//Modif BTP

procedure ZeroLigneMontant(TOBL : TOB) ;
var Indice : integer;
begin
  (*
	TOBL.PutValue('GL_TOTALHT',0);
	TOBL.PutValue('GL_TOTALHTDEV',0);
	TOBL.PutValue('GL_MONTANTHT',0);
	TOBL.PutValue('GL_MONTANTHTDEV',0);
	TOBL.PutValue('GL_MONTANTTTC',0);
	TOBL.PutValue('GL_MONTANTTTCDEV',0);
  for Indice := 0 to  do
  begin
	  TOBL.PutValue('GL_TOTALTAXE'+InttoStr(Indice),0);
	  TOBL.PutValue('GL_TOTALTAXEDEV'+InttoStr(Indice),0);
  end;
	TOBL.PutValue('GL_TOTALTTC',0);
	TOBL.PutValue('GL_TOTALTTCDEV',0);
	TOBL.PutValue('GL_TOTALREMPIED',0);
	TOBL.PutValue('GL_TOTALREMPIEDDEV',0);
	TOBL.PutValue('GL_TOREMLIGNE',0);
	TOBL.PutValue('GL_TOREMLIGNEDEV',0);
	TOBL.PutValue('GL_TOTESCLIGNE',0);
	TOBL.PutValue('GL_TOTESCLIGNEDEV',0);
  *)
  ZeroLigne(TOBL);
  //
  TOBL.PutValue('TOTREMLIGNETTCDEV',0) ;
  TOBL.PutValue('TOTREMLIGNETTC',0) ;
  TOBL.PutValue('TOTESCLIGNETTCDEV',0) ;
  TOBL.PutValue('TOTESCLIGNETTC',0) ;
  TOBL.PutValue('TOTREMPIEDTTCDEV',0) ;
  TOBL.PutValue('TOTREMPIEDTTC',0) ;
end;


procedure ZeroLignePourcent(TOBL: TOB);
var i: integer;
  Nam: string;
begin
  for i := 1 to TOBL.NbChamps do
  begin
    Nam := TOBL.GetNomChamp(i);
    if Copy(Nam, 1, 6) = 'GL_TOT' then TOBL.PutValue(Nam, 0);
    if Copy(Nam, 1, 10) = 'GL_MONTANT' then TOBL.PutValue(Nam, 0);
    if Copy(Nam, 1, 5) = 'GL_PU' then TOBL.PutValue(Nam, 0);
  end;
  TOBL.Putvalue('GL_DPA', 0);
  TOBL.Putvalue('GL_DPR', 0);
  TOBL.Putvalue('GL_PMAP', 0);
  TOBL.Putvalue('GL_PMRP', 0);
end;
//--------

procedure SommerLignes(TOBPiece: TOB; ARow, Niv: integer);
var TOBL, TOBS: TOB;
  i, k: integer;
  RefU, TypL, Nam, TypD: string;
  // VARIANTE
  variante : boolean;
  DEV : Rdevise;
begin
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
	//
  TOBS := GetTOBLigne(TOBPiece, ARow);
  if TOBS = nil then Exit;
  variante := IsVariante (TOBS);
  ZeroLigne(TOBS);
  TOBS.PutValue('GL_QTEFACT', 0);
  TOBS.PutValue('GL_QTESTOCK', 0);
  TOBS.PutValue('GL_QTERESTE', 0); { NEWPIECE }
  TOBS.PutValue('GL_MTRESTE', 0); { NEWPIECE }

  // Modif BTP
  TOBS.putValue('GL_PMAP', 0);
  TOBS.putValue('GL_PMRP', 0);
  TOBS.putValue('GL_DPA', 0);
  TOBS.putValue('GL_DPR', 0);
  TOBS.putValue('GL_MONTANTPA', 0);
  TOBS.putValue('GL_MONTANTPR', 0);

  TOBS.putValue('GL_MONTANTFG', 0);
  TOBS.putValue('GL_MONTANTFC', 0);

  TOBS.putValue('GL_MONTANTFR', 0);
  TOBS.putValue('GL_MONTANTFG', 0);
  TOBS.putValue('GL_MONTANTPAFG', 0);
  TOBS.putValue('GL_MONTANTPAFR', 0);
  TOBS.putValue('GL_MONTANTPAFC', 0);
  TOBS.putvalue('GL_TPSUNITAIRE', 0);
  TOBS.putvalue('GL_HEURE', 0);
  TOBS.putvalue('GL_TOTALHEURE', 0);
  if TOBS.FieldExists('BLF_MTPRODUCTION') then TOBS.putvalue('BLF_MTPRODUCTION', 0);
  //TOBS.putvalue('MONTANTSIT', 0);
  if (IsgestionAvanc(TOBpiece)) and (TOBS.FieldExists ('BLF_MTMARCHE')) then
  begin
    TOBS.putvalue('BLF_MTMARCHE', 0);
    TOBS.putvalue('BLF_MTDEJAFACT', 0);
    TOBS.putvalue('BLF_MTCUMULEFACT', 0);
    TOBS.putvalue('BLF_MTSITUATION', 0);
    TOBS.putvalue('BLF_POURCENTAVANC', 0);
  end;
  TOBS.PutValue('POURCENTMARQ',0);
  TOBS.putvalue('POURCENTMARG', 0);
  TOBS.putvalue('GL_COEFMARG', 0);
  // ---------
  for i := ARow - 1 downto 1 do
  begin
    TOBL := GetTOBLigne(TOBPiece, i);
    if TOBL = nil then Break;
    TypL := TOBL.GetValue('GL_TYPELIGNE');
    if IsLigneFromCentralis(TOBL) then continue;
    if (TOBL.GetValue('GL_NATUREPIECEG')<>'PBT') and (isVariante(TOBL)) and (not variante)  then continue;
    if Niv = 0 then
    begin
      if ((TypL = 'TOT') or (TypL = 'DEB')) then Break;
    end
    else
    begin
    // VARIANTE
     (* if ((TypL = 'DP' + IntToStr(Niv)) or (TypL = 'TP' + IntToStr(Niv))) then Break; *)
     if IsParagraphe (TOBL,niv) then break;
    end;
    // Modif BTP
    if (TYPl = 'RG') or (IsLigneFromCentralis(TOBL)) then continue;
    // ---
    TypD := TOBL.GetValue('GL_TYPEDIM');
    if TypD = 'GEN' then Continue;
    if (IsSousDetail(TOBL)) or (IsLigneDetail(nil, TOBL)) then Continue;
    RefU := TOBL.GetValue('GL_ARTICLE');
    if RefU = '' then Continue;
    for k := 1 to TOBL.NbChamps do
    begin
      Nam := TOBL.GetNomChamp(k);
      if Copy(Nam, 1, 6) = 'GL_TOT' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + Arrondi(TOBL.GetValue(Nam),DEV.decimale));
      if Copy(Nam, 1, 10) = 'GL_MONTANT' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + Arrondi(TOBL.GetValue(Nam),DEV.decimale));
      // Modif BTP
      if Nam = 'GL_PMAP' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + (TOBL.GetValue(Nam) * TOBL.GetValue('GL_QTEFACT')));
      if Nam = 'GL_PMRP' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + (TOBL.GetValue(Nam) * TOBL.GetValue('GL_QTEFACT')));
      if Nam = 'GL_DPA' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + (TOBL.GetValue(Nam) * TOBL.GetValue('GL_QTEFACT')));
      if Nam = 'GL_DPR' then TOBS.PutValue(Nam, TOBS.GetValue(Nam) + (TOBL.GetValue(Nam) * TOBL.GetValue('GL_QTEFACT')));
      // ---
    end;
    CumuleSuivantType (TOBS,TOBL);
    if (IsgestionAvanc(TOBpiece)) and (TOBS.FieldExists ('BLF_MTMARCHE')) then
    begin
      TOBS.putvalue('BLF_MTMARCHE',TOBS.Getvalue('BLF_MTMARCHE')+TOBL.Getvalue('BLF_MTMARCHE'));
      TOBS.putvalue('BLF_MTDEJAFACT', TOBS.Getvalue('BLF_MTDEJAFACT')+TOBL.Getvalue('BLF_MTDEJAFACT'));
      TOBS.putvalue('BLF_MTCUMULEFACT', TOBS.Getvalue('BLF_MTCUMULEFACT')+TOBL.Getvalue('BLF_MTCUMULEFACT'));
      TOBS.putvalue('BLF_MTSITUATION', TOBS.Getvalue('BLF_MTSITUATION')+TOBL.Getvalue('BLF_MTSITUATION'));
      if TOBS.Getvalue('BLF_MTMARCHE') <> 0 then
      begin
        TOBS.putvalue('BLF_POURCENTAVANC',Arrondi(TOBS.Getvalue('BLF_MTCUMULEFACT')/TOBS.Getvalue('BLF_MTMARCHE')*100,2));
      end else
      begin
        TOBS.putvalue('BLF_POURCENTAVANC',0);
      end;
    end;
    if TOBS.FieldExists('BLF_MTPRODUCTION') then
    begin
      TOBS.putvalue('BLF_MTPRODUCTION', TOBS.Getvalue('BLF_MTPRODUCTION')+TOBL.Getvalue('BLF_MTPRODUCTION'));
    end;
    {$IFDEF BTP}
    TOBS.PutValue('GL_TPSUNITAIRE', TOBS.GetValue('GL_TPSUNITAIRE') + (TOBL.GetValue('GL_TPSUNITAIRE') * TOBL.getValue('GL_QTEFACT')));
    TOBS.PutValue('GL_HEURE', TOBS.GetValue('GL_HEURE') + TOBL.GetValue('GL_HEURE'));
    TOBS.putvalue('GL_TOTALHEURE',TOBS.Getvalue('GL_HEURE'));
//    TOBS.PutValue('MONTANTSIT', TOBS.GetValue('MONTANTSIT') + TOBL.GetValue('MONTANTSIT'));
    {$ENDIF}
    if not (ctxAffaire in V_PGI.PGIContexte) then
    begin
      TOBS.PutValue('GL_QTEFACT', TOBS.GetValue('GL_QTEFACT') + TOBL.GetValue('GL_QTEFACT'));
      TOBS.PutValue('GL_QTESTOCK', TOBS.GetValue('GL_QTESTOCK') + TOBL.GetValue('GL_QTESTOCK'));
      TOBS.PutValue('GL_QTERESTE', TOBS.GetValue('GL_QTERESTE') + TOBL.GetValue('GL_QTERESTE'));
      TOBS.PutValue('GL_MTRESTE',  TOBS.GetValue('GL_MTRESTE')  + TOBL.GetValue('GL_MTRESTE'));
    end;
  end;
  if Niv = 0 then
  begin
    TOBS.PutValue('GL_TYPELIGNE', 'TOT');
    TOBS.PutValue('GL_TYPEDIM', 'NOR');
    if TOBS.GetValue('GL_LIBELLE') = '' then TOBS.PutValue('GL_LIBELLE', TraduireMemoire('Sous-total'));
  end;
end;

{============================ ACOMPTES ===============================}

procedure GetSommeReglement(TOBAcomptes: TOB; var XP, XD: Double; Fournisseur : string=''; All : Boolean=false);
var i: integer;
  TOBACC: TOB;
begin
  XP := 0;
  XD := 0;
  if TOBAcomptes = nil then Exit;
  for i := 0 to TOBAcomptes.Detail.Count - 1 do
  begin
    TOBACC := TOBAcomptes.Detail[i];
    if (TOBACC.GetString('GAC_FOURNISSEUR')=Fournisseur) or (All) then
    begin
		  if TOBACC.GetBoolean ('GAC_ISREGLEMENT') then
      begin
        XD := Arrondi(XD + TOBACC.GetValue('GAC_MONTANTDEV'), 6);
        XP := Arrondi(XP + TOBACC.GetValue('GAC_MONTANT'), 6);
      end;
    end;
  end;
end;

procedure GetSommeAcomptes(TOBAcomptes: TOB; var XP, XD: Double; Fournisseur : string=''; All : Boolean=false);
var i: integer;
  TOBACC: TOB;
begin
  XP := 0;
  XD := 0;
  if TOBAcomptes = nil then Exit;
  for i := 0 to TOBAcomptes.Detail.Count - 1 do
  begin
    TOBACC := TOBAcomptes.Detail[i];
    if (TOBACC.GetString('GAC_FOURNISSEUR')=Fournisseur) or (All) then
    begin
      XD := Arrondi(XD + TOBACC.GetValue('GAC_MONTANTDEV'), 6);
      XP := Arrondi(XP + TOBACC.GetValue('GAC_MONTANT'), 6);
    end;
  end;
end;

procedure ReliquatAcomptes(TOBAcomptes, OldAcomptes, AcomptesRel: TOB);
var i: integer;
  TOBACCNew, OldAcc, RelAcc: TOB;
  OldJalEcr: string;
  OldNumEcr: integer;
  OldMontantP, OldMontantD: Double;
  NewMontantP, NewMontantD: Double;
  DeltaP, DeltaD: Double;
begin
  if OldAcomptes.Detail.Count <= 0 then Exit;
  for i := 0 to OldAcomptes.Detail.Count - 1 do
  begin
    OldAcc := OldAcomptes.Detail[i];
    OldJalEcr := OldAcc.GetValue('GAC_JALECR');
    OldNumEcr := OldAcc.GetValue('GAC_NUMECR');
    OldMontantP := OldAcc.GetValue('GAC_MONTANT');
    OldMontantD := OldAcc.GetValue('GAC_MONTANTDEV');
    TOBACCNew := TOBAcomptes.FindFirst(['GAC_JALECR', 'GAC_NUMECR'], [OldjalEcr, OldNumEcr], False);
    if TOBACCNew = nil then
    begin
      RelAcc := TOB.Create('ACOMPTES', AcomptesRel, -1);
      RelAcc.Dupliquer(OldAcc, True, True);
    end else
    begin
      NewMontantP := TOBAccNew.GetValue('GAC_MONTANT');
      NewMontantD := TOBAccNew.GetValue('GAC_MONTANTDEV');
      if NewMontantP < OldMontantP then
      begin
        RelAcc := TOB.Create('ACOMPTES', AcomptesRel, -1);
        RelAcc.Dupliquer(OldAcc, True, True);
        DeltaP := Arrondi(OldMontantP - NewMontantP, 6);
        DeltaD := Arrondi(OldMontantD - NewMontantD, 6);
        RelAcc.PutValue('GAC_MONTANT', DeltaP);
        RelAcc.PutValue('GAC_MONTANTDEV', DeltaD);
      end;
    end;
  end;
end;

procedure AcomptesVersPiece(TOBAcomptes, TOBPiece: TOB);
var TOTD, TOTP: Double;
begin
  GetSommeAcomptes(TOBAcomptes, TOTP, TOTD,'',true);
  TOBPiece.PutValue('GP_ACOMPTE', TOTP);
  TOBPiece.PutValue('GP_ACOMPTEDEV', TOTD);
end;

procedure ValideLesAcomptes(TOBPiece, TOBAcomptes: TOB; TOBAcomptes_O: TOB = nil);
var i, Numero, Indice: integer;
  Nature, Souche: string;
  TOBACCINI, TOBACC, TOBA, TOBRACC, TOBRA: TOB;
  okok : boolean;
begin
  AcomptesVersPiece(TOBAcomptes, TOBPiece);
  if TOBAcomptes.Detail.Count <= 0 then Exit;
  if (TOBAcomptes_O <> nil) and (TOBAcomptes_O.detail.count <> 0) and (tobpiece.getString('GP_ATTACHEMENT')='') then
  begin
    for i := 0 to TOBAcomptes.Detail.Count - 1 do
    begin
      TOBACC := TOBAcomptes.Detail[i];
      TOBACC.PutValue('GAC_PIECEPRECEDENTE', EncoderefPiece (TOBACC));
    end;
//    
    TOBRACC := TOB.Create('LESACOMPTES', nil, -1);
    for i := 0 to TOBAcomptes_O.detail.count - 1 do
    begin
      TOBACCINI := TOBAcomptes_O.detail[I];
      TOBA := TOBAcomptes.FindFirst(['GAC_JALECR', 'GAC_NUMECR'], [TOBACCINI.GetValue('GAC_JALECR'), TOBACCINI.GetValue('GAC_NUMECR')], true);
      if TOBA <> nil then
      begin
        if TOBA.GetValue('GAC_MONTANT') <= TOBACCINI.GetValue('GAC_MONTANT') then
        begin
          // Si le montant de l'acompte constaté sur le nouveau doc est inférieur a celui du document précédent
          // Alors on cree un acompte sur le document précédent
          // correspondant au reliquat d'acompte
          TOBRA := TOB.create('ACOMPTES', TOBRACC, -1);
          TOBRA.dupliquer(TOBACCINI,true,true);
          TOBRA.PutValue('GAC_MONTANT', TOBACCINI.GetValue('GAC_MONTANT') - TOBA.GetValue('GAC_MONTANT'));
          TOBRA.PutValue('GAC_MONTANTDEV', TOBACCINI.GetValue('GAC_MONTANTDEV') - TOBA.GetValue('GAC_MONTANTDEV'));
        end;
      end else
      begin
        // Si l'on ne trouve plus l'acompte dans le nouveau doc alors on laisse le rattachement sur l'ancien document
        TOBRA := TOB.create('ACOMPTES', TOBRACC, -1);
        TOBRA.dupliquer(TOBACCINI, true, true);
      end;
    end;
    if TOBRACC.detail.count > 0 then
      begin
      TRY
        okok := TOBRACC.InsertOrUpdateDB(false);
      EXCEPT
        on E: Exception do
        begin
          PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour acomptes (1)');
        end;
      END;
      if not okok then
      begin
        V_PGI.IoError := oeUnknown;
        Exit;
      end;
    end;
    TOBRACC.free;
  end;
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  // AJOUT BTP LS le 26/11/2002 (purge)
  i := 0;
  if TOBAcomptes.detail.count > 0 then
  begin
    repeat
      if TOBAcomptes.detail[i].getValue('GAC_MONTANTDEV') = 0 then TOBAcomptes.detail[i].free
      else Inc(i);
    until i >= TOBAcomptes.detail.count - 1;
  end;
  if TOBAcomptes.Detail.Count <= 0 then Exit;
  // --
  for i := 0 to TOBAcomptes.Detail.Count - 1 do
  begin
    TOBACC := TOBAcomptes.Detail[i];
    TOBACC.PutValue('GAC_NATUREPIECEG', Nature);
    TOBACC.PutValue('GAC_SOUCHE', Souche);
    TOBACC.PutValue('GAC_NUMERO', Numero);
    TOBACC.PutValue('GAC_INDICEG', Indice);
  end;
  TRY
    okok := TOBAcomptes.InsertDB(nil);
  EXCEPT
    on E: Exception do
  begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour acomptes (2)');
    end;
  END;
  if not okok then
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;

procedure LoadLesAcomptes(TOBPiece, TOBAcomptes: TOB; CleDoc: R_CleDoc; TOBAcptesIni: TOb = nil);
var NaturePieceG: string;
  GereAcompte: Boolean;
  QQ: TQuery;
  TOBL: TOB;
  Indice: integer;
begin
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  GereAcompte := (GetInfoParPiece(NaturePieceG, 'GPP_ACOMPTE') = 'X');
  if ((not GereAcompte) and (TOBPiece.GetValue('GP_ACOMPTEDEV') = 0)) then Exit;
  QQ := OpenSQl('SELECT * FROM ACOMPTES WHERE ' + WherePiece(CleDoc, ttdAcompte, False), True,-1, '', True);
  if not QQ.EOF then TOBAcomptes.LoadDetailDB('ACOMPTES', '', '', QQ, False);
  Ferme(QQ);
  if (TOBAcptesIni <> nil) then
  begin
    for Indice := 0 to TOBAcomptes.detail.count - 1 do
    begin
      TOBL := TOB.Create('ACOMPTES', TOBAcptesIni, -1);
      TOBL.dupliquer(TOBAcomptes.detail[Indice], true, true);
    end;
  end;
end;

procedure LoadLesSsTraitants(CleDoc: R_CleDoc; TOBSsTraitant : TOB);
Var StSQL : string;
    QQ    : TQuery;
begin

  //chargement informations cotraitance au niveau de la piece
  StSQL := 'Select *,';
  StSQL := StSQL + '(SELECT BPI_TYPEPAIE FROM PIECEINTERV WHERE ';
  StSQL := StSQL + wherePiece(Cledoc, ttdPieceInterv, True) + ' AND BPI_TIERSFOU=BPE_FOURNISSEUR)';
  StSQl := StSQl + 'AS REGSSTRAIT, ';
  StSql := StSQL + '(SELECT T_LIBELLE    FROM TIERS       WHERE ';
  StSQl := StSQl + 'T_NATUREAUXI="FOU"    AND T_TIERS=BPE_FOURNISSEUR) ';
  StSQl := StSQl + 'AS LIBTIERS ';
  StSQl := StSQL + 'FROM PIECETRAIT WHERE ' + WherePiece(cledoc, ttdPieceTrait, True);

  TOBSsTraitant.LoadDetailFromSQL(StSQl);

end;

procedure LoadLesBlocNotes(SaContexte: TModeAff; TOBLienOle: TOB; Cledoc: R_CleDoc);
var Q: Tquery;
begin
(*
  if (tModeBlocNotes in SaContexte) then
  begin
*)
    // Lecture des textes debut et fin
    Q := OpenSQL('SELECT * FROM LIENSOLE WHERE ' + WherePiece(CleDoc, ttdLienOle, False), True,-1, '', True);
    TOBLIENOLE.LoadDetailDB('LIENSOLE', '', '', Q, False);
    Ferme(Q);
(*
  end;
*)
end;

function SetSituationVivante (TOBPiece : TOB; VIVANTE : string) : Integer;
var SQl : string;
		QQ : Tquery;
    Cledoc : r_cledoc;
begin
	result := 0;
  cledoc := TOB2CleDoc(TOBPiece);
  Sql := 'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+WherePiece(Cledoc,ttdSit,false);
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then
  begin
    result := QQ.Fields[0].AsInteger;
    SQl := 'UPDATE BSITUATIONS SET BST_VIVANTE="'+VIVANTE+'" WHERE '+WherePiece(Cledoc,ttdSit,false);
    ExecuteSQL(SQL);
  end;
  ferme (QQ);
end;

function GetNumSituation (TOBPiece : TOB) : integer;
var SQl : string;
		QQ : Tquery;
    Cledoc : r_cledoc;
begin
	result := 0;
  cledoc := TOB2CleDoc(TOBPiece);
  Sql := 'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+WherePiece(Cledoc,ttdSit,false);
  QQ := OpenSql (Sql,true,1,'',true);
  if not QQ.eof then result := QQ.findField('BST_NUMEROSIT').AsInteger;
  ferme (QQ);
end;

procedure LoadLesRetenuesPRE(TOBPiece, TOBRGPRE: TOB; Action : TactionFiche= taCreat);
var Req : string;
		TOBLOC,TOBL,TOBP : TOB;
    Indice : integer;
    isAvanc : boolean;
begin
	IsAvanc := (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC;')>0);

  TOBLOC := TOB.Create ('LES RG PRE',nil,-1);
  TRY
    if isAvanc then
    begin
      Req := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,PRG_NUMLIGNE,PRG_FOURN,'+
             'PRG_MTHTRG,PRG_MTHTRGDEV,PRG_MTTTCRGDEV,PRG_MTTTCRG,PRG_NUMCAUTION,PRG_CAUTIONMT,PRG_CAUTIONMTDEV,'+
             'GL_PIECEPRECEDENTE FROM BSITUATIONS '+
             'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=BST_NATUREPIECE AND '+
             'PRG_SOUCHE=BST_SOUCHE AND PRG_NUMERO=BST_NUMEROFAC AND PRG_INDICEG=0 '+
             'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BST_NATUREPIECE AND GL_SOUCHE=BST_SOUCHE AND '+
             'GL_NUMERO=BST_NUMEROFAC AND GL_INDICEG=0 AND GL_NUMORDRE=PRG_NUMLIGNE '+
             'WHERE BST_SSAFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NUMEROSIT < '+IntToStr(GetNumSituation (TOBpiece)) +' AND '+
             'BST_VIVANTE="X" '+
             'ORDER BY GL_PIECEPRECEDENTE,PRG_FOURN,BST_NUMEROSIT';
    end else
    begin
      Req := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,PRG_NUMLIGNE,PRG_FOURN,'+
             'PRG_MTHTRG,PRG_MTHTRGDEV,PRG_MTTTCRGDEV,PRG_MTTTCRG,PRG_NUMCAUTION,PRG_CAUTIONMT,PRG_CAUTIONMTDEV,'+
             'GL_PIECEPRECEDENTE FROM BSITUATIONS '+
             'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=BST_NATUREPIECE AND '+
             'PRG_SOUCHE=BST_SOUCHE AND PRG_NUMERO=BST_NUMEROFAC AND PRG_INDICEG=0 '+
             'LEFT JOIN LIGNE ON GL_NATUREPIECEG=BST_NATUREPIECE AND GL_SOUCHE=BST_SOUCHE AND '+
             'GL_NUMERO=BST_NUMEROFAC AND GL_INDICEG=0 AND GL_NUMORDRE=PRG_NUMLIGNE '+
             'WHERE BST_SSAFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_NUMEROFAC < '+IntToStr(TOBPiece.GetValue('GP_NUMERO')) +' AND '+
             'BST_VIVANTE="X" '+
             'ORDER BY GL_PIECEPRECEDENTE,PRG_FOURN,BST_NUMEROSIT';
    end;
    TOBLOC.LoadDetailDBFromSQL ('PIECERG',req,false);
    for Indice := 0 to TOBLOC.detail.count -1 do
    begin
      TOBL := TOBLOC.detail[Indice];
      TOBP := TOBRGPre.findFirst(['PIECEPRECEDENTE','PRG_FOURN'],
                [TOBL.GetValue('GL_PIECEPRECEDENTE'),TOBL.GetValue('PRG_FOURN')],true);
      if TOBP = nil then
      begin
        TOBP := TOB.Create ('PIECERG',TOBRGPre,-1);
        TOBP.Dupliquer(TOBL,false,true);
        TOBP.AddChampSupValeur('PIECEPRECEDENTE',TOBL.GetValue('GL_PIECEPRECEDENTE'));
        if TOBP.GetString('PRG_NUMCAUTION')='' then
        begin
          TOBP.SetDouble('PRG_CAUTIONMT',0);
          TOBP.SetDouble('PRG_CAUTIONMTDEV',0);
        end;
      end else
      begin
        TOBP.SetDouble ('PRG_MTHTRG',TOBP.GetDouble ('PRG_MTHTRG')+TOBL.GetDouble ('PRG_MTHTRG'));
        TOBP.SetDouble ('PRG_MTHTRGDEV',TOBP.GetDouble ('PRG_MTHTRGDEV')+TOBL.GetDouble ('PRG_MTHTRGDEV'));
        TOBP.SetDouble ('PRG_MTTTCRG',TOBP.GetDouble ('PRG_MTTTCRG')+TOBL.GetDouble ('PRG_MTTTCRG'));
        TOBP.SetDouble ('PRG_MTTTCRGDEV',TOBP.GetDouble ('PRG_MTTTCRGDEV')+TOBL.GetDouble ('PRG_MTTTCRGDEV'));
        if TOBP.GetString('PRG_NUMCAUTION')='' then
        begin
          TOBP.SetDouble('PRG_CAUTIONMT',0);
          TOBP.SetDouble('PRG_CAUTIONMTDEV',0);
        end;
      end;
    end;
  FINALLY
  	TOBLOC.Free;
  END;
end;

procedure LoadLesRetenues(TOBPiece, TOBPieceRG, TOBBasesRG: TOB; Action : TactionFiche);
var Q       : Tquery;
    i       : integer;
    IndiceRg,NumSituation: integer;
    TOBL, TOBB,TOBCautionsfac,TOBFacture: TOB;
    CautionU,CautionUdev,CautionApres,CautionApresDev : double;
    isAvanc : boolean;
    CleDoc  : R_cledoc;
begin

  //chargement du Cledoc
  CleDoc := TOB2CleDoc(TOBPiece);
  //
	IsAvanc := (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC;')>0);
  if (Action = TaModif) and (IsAvanc) then
  begin
		NumSituation := GetNumSituation (TOBpiece);
  end;
  IndiceRg := 1;
  Q := OpenSQL('SELECT * FROM PIECERG WHERE ' + WherePiece(CleDoc, ttdretenuG, False), True,-1, '', True);
  TOBPieceRg.loaddetailDB('PIECERG', '', '', Q, False, True);
  Ferme(Q);
  for i := 0 to TOBPieceRg.detail.count - 1 do
  begin
	  TOBPIeceRG.detail[i].addchampsupValeur('INDICERG', 0);
    TOBPieceRg.detail[i].addchampsupValeur('PIECEPRECEDENTE', '');
    TOBPieceRg.detail[i].AddChampSupValeur ('NUMORDRE',0);
    TOBPieceRg.detail[i].AddChampSupValeur ('CAUTIONUTIL',0);
    TOBPieceRg.detail[i].AddChampSupValeur ('CAUTIONUTILDEV',0);
    TOBPieceRg.detail[i].AddChampSupValeur ('CAUTIONAPRES',0);
    TOBPieceRg.detail[i].AddChampSupValeur ('CAUTIONAPRESDEV',0);
  end;
  if (TOBPiece.getString('GP_RGMULTIPLE')='') and (TOBPieceRG.detail.count > 0) then
  begin
  	if not IsRGSimple  (TOBPieceRG) then TOBPiece.PutValue('GP_RGMULTIPLE','X')
    			 													else TOBPiece.PutValue('GP_RGMULTIPLE','-');
  end;
  // Chargement de la tob des Bases de tva retenues de garantie
  Q := OpenSQL('SELECT * FROM PIEDBASERG WHERE ' + WherePiece(CleDoc, ttdBaseRG, False), True,-1, '', True);
  TOBBasesRg.loaddetaildb('PIEDBASERG', '', '', Q, False);
  Ferme(Q);
  for i := 0 to TOBBasesRG.detail.count - 1 do
  begin
    TOBBasesRG.detail[i].addchampsupValeur('INDICERG', 0);
  end;
  if TOBPieceRG.detail.Count > 0 then
  begin
    // Attribution du N° d'indice de RG
    TOBL := TOBPIECE.findfirst(['GL_TYPELIGNE'], ['RG'], true);
    while TOBL <> nil do
    begin
      // Calcul Retenue de garantie dans le cas de document regroupé
      if Not TOBL.FieldExists ('INDICERG') then TOBL.AddChampSupValeur('INDICERG',0);
      TOBL.Putvalue('INDICERG', IndiceRG);
      TOBB := TOBPIeceRG.findfirst(['PRG_NUMLIGNE'], [TOBL.GetValue('GL_NUMORDRE')], true);
      repeat
      	if Not TOBL.FieldExists ('INDICERG') then TOBL.AddChampSupValeur('INDICERG',0);
      	TOBB.PutValue('INDICERG', IndiceRG);
        TOBB.AddChampSupValeur ('PIECEPRECEDENTE',TOBL.GEtValue('GL_PIECEPRECEDENTE'));
        CautionU :=0; CautionUdev := 0;
        if Not TOBB.FieldExists ('CAUTIONUTIL') then TOBB.AddChampSupValeur ('CAUTIONUTIL',0);
        if Not TOBB.FieldExists ('CAUTIONUTILDEV') then TOBB.AddChampSupValeur ('CAUTIONUTILDEV',0);
        if Not TOBB.FieldExists ('CAUTIONAPRES') then TOBB.AddChampSupValeur ('CAUTIONAPRES',0);
        if Not TOBB.FieldExists ('CAUTIONAPRESDEV') then TOBB.AddChampSupValeur ('CAUTIONAPRESDEV',0);
        (*
        if (Pos(TOBPiece.getString('GP_NATUREPIECEG'),'FBT;FBP')>0) and (isAvanc) then
        begin
        	GetCautionUsedAfter (TOBPiece,TOBL.GEtValue('GL_PIECEPRECEDENTE'),NumSituation,CautionApres,CautionApresdev,TOBB.getString('PRG_FOURN'));
          TOBB.SetDouble ('CAUTIONAPRES',CautionApres);
          TOBB.SetDouble ('CAUTIONAPRESDEV',CautionApresdev);
        end;
        *)
        if (Action=taModif) or (Action = TaConsult) then
        begin
          if (TOBPiece.getString('GP_NATUREPIECEG')<>'DBT') then
          begin
            GetCautionAlreadyUsed (TOBPiece,TOBL,TOBL.GEtValue('GL_PIECEPRECEDENTE'),CautionU,CautionUdev,TOBB.getString('PRG_FOURN'));
            TOBB.AddChampSupValeur ('NUMORDRE',TOBL.GEtValue('GL_NUMORDRE'));
            TOBB.SetDouble ('CAUTIONUTIL',CautionU);
            TOBB.SetDouble ('CAUTIONUTILDEV',CautionUDev);
          end;
        end;
  	    TOBB := TOBPIeceRG.findnext(['PRG_NUMLIGNE'], [TOBL.GetValue('GL_NUMORDRE')], true);
      until TOBB = nil;
      TOBB := TOBBasesRG.findFirst(['PBR_NUMLIGNE'], [TOBL.GetValue('GL_NUMORDRE')], true);
      while TOBB <> nil do
      begin
        TOBB.PutValue('INDICERG', IndiceRG);
        TOBB := TOBBasesRG.findNext(['PBR_NUMLIGNE'], [TOBL.GetValue('GL_NUMORDRE')], true);
      end;
      Inc(IndiceRG);
      TOBL := TOBPIECE.findnext(['GL_TYPELIGNE'], ['RG'], true);
    end;
  end;
end;

procedure LoadLesLibDetail(TOBPiece, TOBNomenclature, TobOuvrage, TOBTiers, TOBAffaire, TOBTRFPOC: TOB; DEV: Rdevise; saisieTypeAvanc: boolean; TheMetreDoc : TMetreArt);
var ModifSousDetail :boolean;
begin
	ModifSousDetail :=  GetParamSocSecur('SO_BTMODIFSDETAIL',false);
  LoadLesLibDetailNomen(TOBPiece, TOBNomenclature, TOBTiers, TOBAffaire, DEV);
  {$IFDEF BTP}
//	if ((SaisieTypeAvanc) and (ModifSousDetail)) or (not SaisieTypeAvanc) then
  begin
  	LoadLesLibDetailOuvrages(TOBPiece, TOBOuvrage, TOBTiers, TOBAffaire,TOBTRFPOC, DEV, TheMetreDoc);
  end;
  {$ENDIF}
  NumeroteLignesGC(nil, TOBpiece);
end;

// Modif BTP
{=================== RETENUES DE GARANTIE ================================}

procedure ValideLesRetenues(TOBPiece, TOBpieceRg: TOB);
var i, Numero, IndiceRG, Indice: integer;
  Nature, Souche: string;
  TOBRG, TOBL: TOB;
  okok : boolean;
begin
  if TOBPieceRg = nil then exit;
  if TOBPieceRG.detail.count = 0 then exit;
  i := 0;

  repeat
    TOBRG := TOBPieceRG.detail[i];
    if TOBRG.GetValue('PRG_TAUXRG') = 0 then TOBRG.free else Inc(i);
  until i > TOBPiecERG.detail.count - 1;

  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  TOBL := TOBPiece.findFirst(['GL_TYPELIGNE'], ['RG'], true);
  while TOBL <> nil do
  begin
    IndiceRG := TOBL.GetValue('INDICERG');
    TOBRG := TOBPIECERG.findfirst(['INDICERG'], [IndiceRg], true);
    while TOBRG <> nil do
    begin
    	if TOBRG.fieldExists('PIECEPRECEDENTE') then
    			TOBL.putValue('GL_PIECEPRECEDENTE',TOBRG.GetValue('PIECEPRECEDENTE'));
      TOBRG.PutValue('PRG_NATUREPIECEG', Nature);
      TOBRG.PutValue('PRG_SOUCHE', Souche);
      TOBRG.PutValue('PRG_NUMERO', Numero);
      TOBRG.PutValue('PRG_INDICEG', Indice);
      TOBRG.PutValue('PRG_DATEPIECE', TOBPIece.getValue('GP_DATEPIECE'));
      TOBRG.Putvalue('PRG_NUMLIGNE', TOBL.getValue('GL_NUMORDRE'));
      TOBRG := TOBPIECERG.findnext(['INDICERG'], [IndiceRg], true);
    end;
    TOBL := TOBPiece.findNext(['GL_TYPELIGNE'], ['RG'], true);
  end;
  TRY
    okok :=  TOBPieceRG.InsertDB(nil);
  EXCEPT
    on E: Exception do
  begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour des RG');
    end;
  END;
  if not okok then
  begin
    V_PGI.IoError := oeUnknown;
  end;
end;

procedure ValideLesBasesRg(TOBPiece, TOBBasesRg: TOB);
var i, Numero, IndiceRG, Indice: integer;
  Nature, Souche: string;
  TOBB, TOBL: TOB;
begin
  if TOBBasesRg = nil then exit;
  if TOBBasesRG.detail.count = 0 then exit;
  i:=0;
  repeat // after me ??
    TOBB := TOBBasesRG.detail[I];
    if (TOBB.GEtValue('PBR_BASEDEV') = 0) or (TOBB.GetValue('PBR_VALEURDEV') = 0) then TOBB.free else Inc(i);
  until i > TOBBasesRG.detail.count - 1;

  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  Indice := TOBPiece.GetValue('GP_INDICEG');
  TOBL := TOBPiece.findFirst(['GL_TYPELIGNE'], ['RG'], true);
  while TOBL <> nil do
  begin
    IndiceRG := TOBL.GetValue('INDICERG');
    TOBB := TOBBasesRG.findfirst(['INDICERG'], [IndiceRg], true);
    while TOBB <> nil do
    begin
      TOBB.PutValue('PBR_NATUREPIECEG', Nature);
      TOBB.PutValue('PBR_SOUCHE', Souche);
      TOBB.PutValue('PBR_NUMERO', Numero);
      TOBB.PutValue('PBR_INDICEG', Indice);
      TOBB.Putvalue('PBR_NUMLIGNE', TOBL.getValue('GL_NUMORDRE'));
      if not TOBB.InsertDB(nil) then V_PGI.IoError := oeUnknown;
      TOBB := TOBBasesRG.findnext(['INDICERG'], [IndiceRg], true);
    end;
    TOBL := TOBPiece.findNext(['GL_TYPELIGNE'], ['RG'], true);
  end;
end;
// -----

function FlagDetruitArriere(TOBPiece: TOB): boolean; { NEWPIECE }
var
  i, iPrec: integer;
  TOBL: TOB;
  StPrec, StColis: string;
  CD: R_CleDoc;
  Q: TQuery;
  PasBon: boolean;
  iColis: integer;
begin
  PasBon := False;
  iPrec := 0;
  icolis := 0;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBL := TOBPiece.Detail[i];
    if i <= 0 then
    begin
      iPrec := TOBL.GetNumChamp('GL_PIECEPRECEDENTE');
      if VH_GC.GCIfDefCEGID then
        iColis := TOBL.GetNumChamp('GL_REFCOLIS');
    end;
    StPrec := TOBL.GetValeur(iPrec);
    StColis := '';
    if VH_GC.GCIfDefCEGID then
    begin
      StColis := TOBL.GetValeur(iColis);
      if StColis <> '' then
      begin
        PasBon := True;
        Break;
      end;
    end;
    while ((StPrec <> '') and (not PasBon)) do
    begin
      DecodeRefPiece(StPrec, CD);
      if CD.NumOrdre <= 0 then { NEWPIECE }
      begin
        PasBon := True;
        Break;
      end;
      Q := OpenSQL('SELECT GL_PIECEPRECEDENTE, GL_QUALIFMVT FROM LIGNE WHERE ' + WherePiece(CD, ttdLigne, True, True), False);
      if not Q.EOF then
      begin
        Q.Edit;
        Q.Fields[1].AsString := 'ANN';
        StPrec := Q.Fields[0].AsString;
        Q.Post;
      end
      else
      begin
        StPrec := '';
        PasBon := True;
        Break;
      end;
      Ferme(Q);
    end;
  end;
  Result := (not PasBon);
end;

{============================= Tables Libres ==================================}
// Si Bouton=True cela signifie que la fonction est appelée à partir du bouton donc
// la fiche GCSTATPIECE doit s'ouvrir

function SaisieTablesLibres(TOBPiece: TOB; Bouton: Boolean = False): boolean;
var Nature, Resultat, Arguments: string;
  VerifTablesLibres: boolean;
begin
  Result := True;
  Resultat := '';
  VerifTablesLibres := False;
  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  if GetInfoParPiece(Nature, 'GPP_AFFPIECETABLE') = 'X' then
  begin
    Arguments := 'NATUREPIECEG=' + Nature + RecupArgTablesLibres(TOBPiece);
    if Bouton then
    begin
      if ctxFO in V_PGI.PGIContexte then Resultat := AGLLanceFiche('MFO', 'FOSTATPIECE', '', '', Arguments)
      else Resultat := AGLLanceFiche('GC', 'GCSTATPIECE', '', '', Arguments);
    end
    else
    begin
      if TOBPiece.FieldExists('VALIDESTATPIECE') then
      begin
        if TOBPiece.GetValue('VALIDESTATPIECE') = 'TRUE' then VerifTablesLibres := True;
      end;
      if VerifTablesLibres then exit
      else
      begin
        if ctxFO in V_PGI.PGIContexte then Resultat := AGLLanceFiche('MFO', 'FOSTATPIECE', '', '', Arguments)
        else Resultat := AGLLanceFiche('GC', 'GCSTATPIECE', '', '', Arguments);
      end;
    end;
    if Trim(Resultat) <> '' then MAJTablesLibres(TOBPiece, Nature, Resultat) else Result := False;
  end;
end;

procedure MAJTablesLibres(TOBPiece: TOB; NaturePieceG, Infos: string);
var Critere, ValMul, ChampMul: string;
  x: integer;
begin
  repeat
    Critere := uppercase(Trim(ReadTokenSt(Infos)));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'VALIDESTATPIECE' then
        begin
          if not TOBPiece.FieldExists('VALIDESTATPIECE') then TOBPiece.AddChampSup('VALIDESTATPIECE', False);
          TOBPiece.PutValue('VALIDESTATPIECE', ValMul);
        end
        else TOBPiece.PutValue('GP_LIBREPIECE' + ChampMul, ValMul);
      end;
    end;
  until Critere = '';
end;

function RecupArgTablesLibres(TOBPiece: TOB): string;
var i: integer;
  ValArg: string;
begin
  Result := '';
  for i := 1 to 3 do
  begin
    ValArg := TOBPiece.GetValue('GP_LIBREPIECE' + IntToStr(i));
    if ValArg <> '' then Result := Result + ';' + IntToStr(i) + '=' + ValArg;
  end;
end;

procedure StatPieceDuplic(TOBPiece: TOB; OldNat, NewNature: string);
var i_ind: integer;
  PlusOld, PlusNew: string;
  Obligatoire, Diff: boolean;
begin
  if GetInfoParPiece(OldNat, 'GPP_AFFPIECETABLE') <> 'X' then exit;
  Diff := False;
  for i_ind := 1 to 3 do
  begin
    PlusNew := GetInfoParPiece(NewNature, 'GPP_PIECETABLE' + InttoStr(i_ind));
    PlusOld := GetInfoParPiece(OldNat, 'GPP_PIECETABLE' + InttoStr(i_ind));
    Obligatoire := GetInfoParPiece(OldNat, 'GPP_CODEPIECEOBL' + InttoStr(i_ind)) = 'X';
    if (PlusNew = '') or (PlusNew <> PlusOld) then
      TOBPiece.PutValue('GP_LIBREPIECE' + IntToStr(i_ind), '');
    if (PlusNew <> PlusOld) or (Obligatoire and (TOBPiece.GetValue('GP_LIBREPIECE' + IntToStr(i_ind)) = '')) then
      Diff := True;
  end;
  if not Diff then
  begin
    if not TOBPiece.FieldExists('VALIDESTATPIECE') then TOBPiece.AddChampSup('VALIDESTATPIECE', False);
    TOBPiece.PutValue('VALIDESTATPIECE', 'TRUE');
  end;
end;

procedure MAVideStatPiece(TOBPiece: TOB);
var i: integer;
begin
  for i := 1 to 3 do
  begin
    TOBPiece.PutValue('GP_LIBREPIECE' + IntToStr(i), '');
  end;
end;

// Lien planning-pièce

function GerePlanningPiece(NaturePiece: string): Boolean;
begin
  Result := False;
  if (VH_GC.GAPlanningSeria) and (GetInfoParPiece(NaturePiece, 'GPP_LIENTACHE') = 'X') then Result := True;
end;

// Modif BTP
//-----------

procedure AjouteChampPourcent(TOBPOu: TOB);
begin
  TOBPou.addchampSup('TP_DEPART', false);
  TOBPou.addchampSup('TP_FIN', false);
end;

procedure AlimenteChampPourcent(TOBPiece, TOBPOu: TOB; Ligne: integer);
var Niveau: Integer;
  TOBL: TOB;
begin
  TOBL := TOBPIece.detail[Ligne];
  Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
  TOBPou.putvalue('TP_DEPART', RecupDebutPourcent(TOBPIece, Ligne - 1, Niveau));
  TOBPou.putvalue('TP_FIN', Ligne - 1);
end;

procedure AddPourcent(TOBPiece, TOBPourcent: TOB; Ligne: integer);
var TOBPou: TOB;
begin
  TOBPOu := TOB.create('POULOC', TOBPourcent, -1);
  AjouteChampPourcent(TOBPOu);
  AlimenteChampPourcent(TOBPiece, TOBPou, Ligne);
end;

function DansChampsPourcent(Ligne: integer; TOBPOurcent: TOB): boolean;
var Indice: Integer;
  TOBPou: TOB;
begin
  result := false;
  for Indice := 0 to TOBPourcent.detail.count - 1 do
  begin
    TOBPou := TOBPourcent.detail[Indice];
    if (Ligne >= TOBPOu.getvalue('TP_DEPART')) and (Ligne <= TOBPOU.getvalue('TP_FIN')) then
    begin
      result := true;
      break;
    end;
  end;
end;

procedure AffecteEcartPmArticle(XX : TForm; TOBPIece, TOBTiers, TOBAffaire, TOBArticles: TOB; Ecart: double; Article: string; IndDepart, IndFin, EtenduApplication:
  integer);
var TOBL, TOBArt: TOB;
  Numl: Integer;
  ArtEcart, MMessage: string;
  QArt: TQuery;
  EnHt: boolean;
  TF : TFFacture;
begin
  TF := TFfacture(XX);
  EnHt := TOBPIece.getValue('GP_FACTUREHT') = 'X';
  ArtEcart := GetParamsoc('SO_BTECARTPMA');
  Qart := opensql('Select GA_ARTICLE from ARTICLE Where GA_CODEARTICLE="' + ArtEcart + '"', true,-1, '', True);
  if not Qart.eof then ArtEcart := Qart.fields[0].AsString
  else ArtEcart := Article;
  ferme(Qart);
  Qart := nil;
  if EtenduApplication = 0 then
  begin
    TOBL := NewTOBLigne(TOBPIece, 0);
    NumL := TOBPiece.Detail.Count;
  end else
  begin
    TOBL := NewTOBLigne(TOBPIece, IndFin + 1);
    NumL := IndFin + 1;
  end;
  InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, NumL);
  TOBArt := FindTOBArtSais(TOBArticles, ArtEcart);
  if TOBArt = nil then
  begin
    QArt := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+ArtEcart+'"',true,-1, '', True);
    if not QArt.EOF then
    begin
      TOBArt := CreerTOBArt(TOBArticles);
      TOBArt.SelectDB('', QArt);
      InitChampsSupArticle (TOBART);
    end;
  end;
  if TOBART = nil then
  begin
    MMessage := 'Aucun article d''écart n''a été paramétré' + #13#10;
    MMessage := MMessage + 'Vous n''obtiendrez donc pas le montant exact demandé' + #13#10;
    ShowMessage(MMessage);
    exit;
  end else ferme(Qart);
  // -- Pour ne pas recalculer a partir des PA * coef ..etc..
  TOBL.putValue('GLC_NONAPPLICFRAIS','X');
  TOBL.putValue('GLC_NONAPPLICFG','X');
  TOBL.putValue('GLC_NONAPPLICFC','X');
  //
  TOBArt.PutValue('REFARTSAISIE', copy(TOBArt.GetValue('GA_ARTICLE'), 1, 18));
  TOBL.PutValue('GL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
  TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
  TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
  TOBL.PutValue('GL_TYPEREF', 'ART');
  ArticleVersLigne(TOBPiece, TOBArt, nil, TOBL, TOBTiers);
  TOBL.PutValue('GL_QTEFACT', 1);
  TOBL.PutValue('GL_QTESTOCK', 1);
  TOBL.PutValue('GL_QTERESTE', 1);

  TOBL.PutValue('GL_MTRESTE', 0);

  TOBL.PutValue('GL_PRIXPOURQTE', 1);
  TOBL.Putvalue('GL_LIBELLE', 'Ecart sur Prix marché');
  if EnHt then
  begin
    TOBL.PutValue('GL_PUHT', ecart );
    TOBL.PutValue('GL_PUHTDEV', ecart);
  end else
  begin
    TOBL.PutValue('GL_PUTTC', ecart);
    TOBL.PutValue('GL_PUTTCDEV', ecart);
  end;
  TOBL.putValue('GL_DPA',0);
  TOBL.putValue('GL_COEFFG',0);
  TOBL.putValue('GL_COEFFC',0);
  TOBL.putValue('GL_DPR',0);
  TOBL.putValue('GL_COEFMARG',0);
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPIECE.PutValue('Gp_RECALCULER', 'X');
  TF.CalculeLaSaisie (-1,numl,false);
  MMessage := 'Un écart de calcul sur prix marché à été constaté' + #13#10;
  MMessage := MMessage + 'Celui à été imputé à l''article ' + trim(copy(Artecart, 1, 18)) + #13#10;
  if EtenduApplication = 0 then MMessage := MMessage + 'en fin de document'
  else MMessage := MMessage + 'en fin de paragraphe';

  ShowMessage(MMessage);
end;

procedure AffecteEcartPmLigne(XX : TForm ; TOBPIece, TOBTiers, TOBAffaire, TOBArticles: TOB; Ecart: double; IndMaxMontant: integer);
var TOBL: TOB;
    TF : TFFacture;
begin
  TF := TFfacture(XX);
  TOBL := TOBPiece.detail[IndMaxMontant];
  TOBL.Putvalue('GL_VALEURREMDEV', TOBL.Getvalue('GL_VALEURREMDEV') - Ecart);
  TOBL.PutValue('GL_RECALCULER', 'X');
  TOBPiece.PutValue('GP_RECALCULER', 'X');
  TF.CalculeLaSaisie (-1,TOBL.getValue('GL_NUMLIGNE'),false);
end;

procedure AffecteEcartPrixMarche(XX : TForm; TOBPIece, TOBTiers, TOBAffaire, TOBArticles: TOB; Ecart: double; Article: string; IndMaxMontant, IndDepart, IndFin,
  EtenduApplication: integer);
var retour: word;
begin
  //
  	(*
  if IndMaxMontant <> -1 then
  begin
    retour :=
      HShowMessage('0;Attribution de l''écart sur prix marché;Veuillez définir comment affecter l''écart;X;YNC;Y;C;;Beep;helpId;Article d''écart| Sur une ligne',
      '', '');
    if retour = Mryes then
    begin
      AffecteEcartPmArticle(XX,TOBPiece, TOBTiers, TOBAffaire, TOBArticles, Ecart, Article, IndDepart, IndFin, EtenduApplication);
    end else if retour = MrNo then
    begin
      AffecteEcartPmLigne(XX,TOBPiece, TOBTiers, TOBAffaire, TOBArticles, Ecart, IndMaxMontant);
    end;

  end else *)
  AffecteEcartPmArticle(XX,TOBPiece, TOBTiers, TOBAffaire, TOBArticles, Ecart, Article, IndDepart, IndFin, EtenduApplication);
end;

procedure AffecteLeDepotParag (FF: TForm; Ligne : integer; TOBPiece : TOB);
var TOBL: TOB;
  indice: Integer;
  TheDepot : string;
  XX : TFFActure;
  niveau : integer;
begin
  XX := TFFacture(FF);
  niveau :=TOBPiece.detail[ligne-1].getValue('GL_NIVEAUIMBRIC');
  TheDepot := TOBPiece.detail[ligne-1].getValue('GL_DEPOT');
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
    TOBL.Putvalue('GL_DEPOT', theDepot);
    XX.ChargeNouveauDepot(nil, TheTob, TheTob.GetValue ('GL_DEPOT'));
  end;
end;

procedure AffecteLeDomaineParag (FF: TForm; Ligne : integer; TOBPiece : TOB);
var TOBL: TOB;
  indice: Integer;
  TheDomaine : string;
  XX : TFFActure;
  niveau : integer;
begin
  XX := TFFacture(FF);
  niveau :=TOBPiece.detail[ligne-1].getValue('GL_NIVEAUIMBRIC');
  TheDomaine := TOBPiece.detail[ligne-1].getValue('GL_DOMAINE');
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
    TOBL.Putvalue('GL_DOMAINE', TheDomaine);
  end;
end;


procedure AppliqueChangeTaxePara(TOBPIece, TOBOuvrage: TOB; ligne: integer; Taxe1,Taxe2,Taxe3,Taxe4,Taxe5: string; Niveau: integer);
var TOBL: TOB;
  indice: Integer;
begin
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
    {$IFDEF BTP}
    if (TOBL.getValue('GL_TYPEARTICLE') = 'OUV') or (TOBL.getValue('GL_TYPEARTICLE') = 'ARP') then
      AppliqueChangeTaxeOuv(TOBPiece, TOBOuvrage, Indice + 1, Taxe1,Taxe2,taxe3,Taxe4,Taxe5);
    {$ENDIF}
    if not IsCommentaire(TOBL) then
    begin
      TOBL.Putvalue('GL_FAMILLETAXE1', Taxe1);
      TOBL.Putvalue('GL_FAMILLETAXE2', Taxe2);
      TOBL.Putvalue('GL_FAMILLETAXE3', Taxe3);
      TOBL.Putvalue('GL_FAMILLETAXE4', Taxe4);
      TOBL.Putvalue('GL_FAMILLETAXE5', Taxe5);
      TOBPIece.putvalue('GP_RECALCULER', 'X');
      TOBL.PutValue('GL_RECALCULER', 'X');
    end;
  end;
end;

procedure AppliqueChangeDateLivPara(TOBPIece, TOBOuvrage: TOB; Ligne: integer; TheCurDate: TDateTime; Niveau: integer);
var TOBL: TOB;
  indice: Integer;
begin
  for Indice := Ligne to TOBPiece.detail.count - 1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    TOBL.Putvalue('GL_DATELIVRAISON', TheCurDate);
    if (copy(TOBL.getvalue('GL_TYPELIGNE'), 1, 2) = 'TP') and
      (TOBL.getvalue('GL_NIVEAUIMBRIC') = niveau) then break;
  end;
end;

function OkLigneCalculMarche(TOBLig: TOB; IndtypeLigne: integer): boolean;
var TypeLigne: string;
		valTrav : double;
begin
  result := false;
  if TOBLig = nil then exit;
  TypeLigne := TOBLIG.GetValeur(IndTypeLigne);
  //
  ValTrav := VALEUR(TOBLIG.GetValue('GLC_NATURETRAVAIL'));
  if (Valtrav > 0) and (ValTrav < 10) then Exit;  // co traité ou sous traité
  //
  // VARIANTE
  (*if (TYpeLigne <> 'COM') and
    (TypeLigne <> 'TOT') and
    (Copy(TYpeLIgne, 1, 2) <> 'DP') and
    (Copy(TypeLigne, 1, 2) <> 'TP') then Result := true; *)
  if not (IsCommentaire(TOBLIG)) and
    (TypeLigne <> 'TOT') and not (Isparagraphe (TOBLIG)) and not IsVariante(TOBLIG) then Result := true;
end;

procedure BeforeReCalculPR (TOBPiece,TOBOuvrage : TOB;EnHt : boolean;Indepart,IndFin : integer; InclusSTinFG : boolean=false);
var Indice : integer;
		TOBL,TOBOuv : TOB;
    TypeLigne,TypeArticle,IndiceNomen : Integer;
begin
	for Indice := Indepart to IndFin do
  begin
    TOBL := TOBPiece.detail[Indice];
    TypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
    TypeArticle := TOBL.GetNumChamp ('GL_TYPEARTICLE');
    if (TOBL.GetValeur(TypeArticle) = 'EPO') then continue;
    if (TOBL.GetValeur(TypeArticle) = 'POU') then continue;
    if (OkLigneCalculMarche(TOBL, Typeligne)) then
    begin
//    	if not TOBL.FieldExists ('__COEFMARG') then TOBL.AddChampSupValeur ('__COEFMARG',0);
      if (TOBL.GetValue('GL_DPR') = 0) and (TOBL.GetValue('GL_DPA') = 0) and (TOBL.GetValue('GL_PUHTDEV')<> 0) then continue;
      if TOBL.GetValue('GL_DPR') = 0 then TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
      if TOBL.GetValue('GL_DPR') = 0 then
      begin
        TOBL.PutValue('GL_COEFMARG',1);
        TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
      end else
      begin
        if TOBL.GetValue('GL_COEFMARG') = 0 then
        begin
           {if EnHt then }TOBL.PutValue ('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_PUHT')/TOBL.GetValue('GL_DPR'),4));
//                   else TOBL.PutValue ('GL_COEFMARG',TOBL.GetValue('GL_PUTTCDEV')/TOBL.GetValue('GL_DPR'));
        end;
      end;
      if TOBL.GetValue('GL_PUHT') <> 0 then
      begin
        TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('GL_PUHT')- TOBL.GetValue('GL_DPR'))/TOBL.GetValue('GL_PUHT'))*100,2));
      end else
      begin
        TOBL.PutValue('POURCENTMARQ',0);
      end;
      //
      if ((TOBL.GetValue('GL_TYPEARTICLE')<>'PRE') or (TOBL.GetValue('BNP_TYPERESSOURCE') <> 'ST') OR (InclusSTinFG)) and
      	 (TOBL.GetValue('GL_TYPEARTICLE')<>'OUV') and
      	 (TOBL.GetValue('GL_ARTICLE')<>'') and
         (TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X') then
      //
      begin
      	TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
      end;
      if TOBL.GetValue('GL_BLOQUETARIF')='-' then
      begin
      	if (TOBL.GetValue('GL_COEFMARG') <> 0) then
        begin
          if EnHt then TOBL.PutValue('GL_PUHTDEV',Arrondi(TOBL.GEtValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OkdecP))
                  else TOBL.PutValue('GL_PUTTCDEV',Arrondi(TOBL.GEtValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OKdecP));
        end;
      	TOBL.PutValue('GL_RECALCULER','X');
{$IFDEF BTP}
        if (TOBL.GetValeur(TypeArticle) = 'OUV') or (TOBL.GetValeur(TypeArticle) = 'ARP') then
        begin
          IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
          if IndiceNomen > 0 then
          begin
            TOBOuv := TOBOuvrage.Detail[IndiceNomen - 1];
            BeforeRecalculPrOuv (TOBOuv,EnHt,(TOBL.GetValue('GL_BLOQUETARIF')='-'),InclusSTinFG);
          end;
        end;
{$ENDIF}
      end;
    end;
  end;
end;

procedure AfterCalculPR (TOBPiece,TOBOuvrage : TOB;EnHt : boolean;Indepart,IndFin : integer;CalculPv : boolean=true);
var Indice : integer;
		TOBL,TOBOuv : TOB;
    TypeLigne,TypeArticle,IndiceNomen : Integer;
begin
	for Indice := Indepart to IndFin do
  begin
    TOBL := TOBPiece.detail[Indice];
    TypeLigne := TOBL.GetNumChamp('GL_TYPELIGNE');
    TypeArticle := TOBL.GetNumChamp ('GL_TYPEARTICLE');
    if (TOBL.GetValeur(TypeArticle) = 'EPO') then continue;
    if (TOBL.GetValeur(TypeArticle) = 'POU') then continue;
    if (OkLigneCalculMarche(TOBL, Typeligne)) then
    begin
      if (TOBL.GetValue('GL_BLOQUETARIF')='-') AND (CalculPv)  then
      begin
      	if TOBL.GetValue('GL_COEFMARG') <> 0 then
        begin
          if EnHt then TOBL.PutValue('GL_PUHTDEV',Arrondi(TOBL.GEtValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OkdecP))
                  else TOBL.PutValue('GL_PUTTCDEV',Arrondi(TOBL.GEtValue('GL_DPR') * TOBL.GetValue('GL_COEFMARG'),V_PGI.OKdecP));
        end;
        TOBL.PutValue('GL_RECALCULER','X');
{$IFDEF BTP}
        if (TOBL.GetValeur(TypeArticle) = 'OUV') or (TOBL.GetValeur(TypeArticle) = 'ARP') then
        begin
          IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
          if IndiceNomen > 0 then
          begin
            TOBOuv := TOBOuvrage.Detail[IndiceNomen - 1];
            AftercalculPrOuv (TOBOuv,EnHt,(TOBL.GetValue('GL_BLOQUETARIF')='-'));
          end;
        end;
{$ENDIF}
			end else
      begin
				positionneCoefMarge (TOBL);
      end;
    end;
  end;
end;

function BeforeReCalculPrixMarche(TOBPiece,TOBOuvrage: TOB; ModeGestion, IndDepart, Indfin: Integer;InclusStInFg : boolean=false): boolean;
var
  Indice: Integer;
  TOBL: TOB;
  {revient,Vente : double;}
//  EnHt : boolean;
begin
  result := true;
//  EnHt := ( TOBPiece.GetValue ('GP_FACTUREHT') = 'X');
  if (Modegestion = 1)  then exit;
  if Modegestion = 3 then
  begin
   BeforeReCalculPR (TOBPiece,TOBOUvrage,(TOBPiece.getValue('GP_FACTUREHT')='X'),IndDepart,IndFin,InclusStInFg);
//   if (not InclusStInFg) then TOBPiece.putvalue('DPR',TOBPiece.getValue('DPA')+TOBPiece.Getvalue('DPR_ST'));
  end else
  begin
     for Indice := IndDepart to IndFin do
     begin
       TOBL := TOBpiece.detail[Indice];
       if TOBL.getValue('GL_TYPEARTICLE') = 'POU' then
       begin
         Result := false;
         break;
       end;
     end;
  end;
end;

procedure CalculValoAvanc(TOBPiece: TOB; Indice: integer; DEV: Rdevise; ModifAvanc : boolean=false); overload;
var TOBL: TOB;
begin
  TOBL := TOBPiece.detail[Indice]; if TOBL = nil then exit;
  CalculValoAvanc (TOBPiece,TOBL,DEV,ModifAvanc);
end;

procedure CalculValoAvanc(TOBPiece, TOBL: TOB; DEV: Rdevise;ModifAvanc : boolean = false);
begin
  if (TOBL.GetValue('GL_ARTICLE') = '') and (TOBL.GetValue('GL_REFCATALOGUE') = '') then exit;
  if (TOBL.GetValue('GL_TYPEARTICLE') = '$$$') or (TOBL.GetValue('GL_TYPEARTICLE') = '$#$') then exit; // on ne prend pas en compte les sous detail d'ouvrage
  if not ModifAvanc then
  begin
    TOBL.PUtValue('GL_TOTPREVAVANC', arrondi((TOBL.GetValue('GL_QTEPREVAVANC') * TOBL.GetValue('GL_PUHT')) / TOBL.GetValue('GL_PRIXPOURQTE'), DEV.Decimale));
    TOBL.PUtValue('GL_TOTPREVDEVAVAN', arrondi((TOBL.GetValue('GL_QTEPREVAVANC') * TOBL.GetValue('GL_PUHTDEV')) / TOBL.GetValue('GL_PRIXPOURQTE'),
      DEV.Decimale));
  end else
  begin
    TOBL.PUtValue('GL_TOTPREVAVANC', arrondi((TOBL.GetValue('GL_QTEFACT') * TOBL.GetValue('GL_PUHT')) / TOBL.GetValue('GL_PRIXPOURQTE'), DEV.Decimale));
    TOBL.PUtValue('GL_TOTPREVDEVAVAN', arrondi((TOBL.GetValue('GL_QTEFACT') *
                                              TOBL.GetValue('GL_PUHTDEV')) / TOBL.GetValue('GL_PRIXPOURQTE'),
                                              DEV.Decimale));
  end;
end;

procedure CalculAvancementPiece(TobPiece, TOBL: TOB; DEv: Rdevise; ModifAvanc : boolean = false);
var Montant: double;
begin
  (* attention lance aussi depuis le calcul de piece *)
  if (TOBL.GetValue('GL_ARTICLE') = '') and (TOBL.GetValue('GL_REFCATALOGUE') = '') then exit;
  // on ne prend pas en compte les sous detail d'ouvrage
  if (TOBL.GetValue('GL_TYPEARTICLE') = '$$$') or (TOBL.GetValue('GL_TYPEARTICLE') = '$#$') then exit;
  // VARIANTE
  if isVariante(TOBL) then exit;
  // --
  CalculValoAvanc(TOBPIece, TOBL, DEV, ModifAvanc);
  TOBPiece.putValue('AVANCSAISIE', TOBPIece.GetValue('AVANCSAISIE') + TOBL.GetValue('GL_TOTPREVDEVAVAN'));
  Montant := arrondi((TOBL.GetValue('GL_QTESIT') * TOBL.GetValue('GL_PUHTNETDEV')) / TOBL.GetValue('GL_PRIXPOURQTE'), DEV.decimale);
  TOBPIece.putvalue('AVANCPREC', TOBPiece.GetValue('AVANCPREC') + Montant);
end;

function ISAcompteSoldePartiel(TOBPiece: TOB): boolean;
var QQ: TQuery;
begin
  result := false;
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then exit;
  QQ := OpenSql('SELECT AFF_ACOMPTEREND FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBPiece.GetValue('GP_AFFAIREDEVIS') + '"', true,-1, '', True);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;
  if QQ.FindField('AFF_ACOMPTEREND').AsFloat > 0 then Result := True;
  ferme(QQ);
end;

function AcompteRegle(TOBPiece: TOB): boolean;
var QQ: TQuery;
begin
  result := false;
  if TOBPiece.GetValue('GP_AFFAIREDEVIS') = '' then exit;
  QQ := OpenSql('SELECT AFF_ACOMPTE,AFF_ACOMPTEREND FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBPiece.GetValue('GP_AFFAIREDEVIS') + '"', true,-1, '', True);
  if QQ.eof then
  begin
    Ferme(QQ);
    Exit;
  end;
  if QQ.FindField('AFF_ACOMPTE').AsFloat <= QQ.FindField('AFF_ACOMPTEREND').AsFloat then Result := True;
  ferme(QQ);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 21/02/2003
Modifié le ... : 21/02/2003
Description .. : fct qui permet, si pièce achat, et si Affaire gérée
Suite ........ : de verifier que le code affaire est renseigné au niveau de
Suite ........ : chaque ligne
Suite ........ : Ne le fait pas si la pièce ne passe pas en activité (pas lieu
Suite ........ : ..)
Mots clefs ... :
*****************************************************************}

function TestAffaire(TOBPiece: TOB): boolean;
var tb: tob;
  mess: string;
begin
  result := false;
  if not (ctxaffaire in V_PGI.PGIContexte) then exit;
  if (GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), 'GPP_ACHATACTIVITE') <> 'X') then exit;
  tb := TobPiece.findfirst(['GL_AFFAIRE', 'GL_TYPELIGNE'], ['', 'ART'], False);
  if tb <> nil then
  begin
    if ctxscot in V_PGI.PGIContexte then Mess := 'Vous avez des codes missions non renseignés. Confirmez vous la validation de la pièce? '
    else mess := 'Vous avez des codes affaires non renseignés. Confirmez vous la validation de la pièce? ';
    if (PgiAsk(mess)) = MrYes then result := False
    else result := true;
  end
end;

{**********************************************************
Auteur  ...... : C TARDY
Créé le ...... : 23/07/2002
Modifié le ... :
Description .. : Fonction permettant de réduire la liste des dépôts
Suite ........ : dans la tablette GCDEPOT pour qu'ils correspondent
Suite ........ : aux dépôts liés se trouvant dans ET_DEPOTLIE
Retour ....... : '' si ET_DEPOTLIE='' sinon la liste des dépôts
Retour ....... : formatée de la manière suivante "001","002" etc
Mots clefs ... : DEPOT;ETABLISSEMENT
*****************************************************************}

function ListeDepotParEtablissement(etab: string): string;
var Q: TQuery;
  ListeDepotLie: string;
begin
  result := '';
  if GetParamSoc('SO_GCMULTIDEPOTS') then
  begin
    Q := OpenSQL('select ET_DEPOTLIE from ETABLISS where ET_ETABLISSEMENT="' + etab + '"', true,-1, '', True);
    if not Q.EOF then ListeDepotLie := Q.FindField('ET_DEPOTLIE').AsString;
    ferme(Q);
    result := StringReplace(ListeDepotLie, ';', '","', [rfReplaceAll]);
    if result <> '' then result := '"' + result + '"';
  end;
end;

{**********************************************************
Auteur  ...... : C TARDY
Créé le ...... : 23/07/2002
Modifié le ... :
Description .. : Fonction permettant de définir l'ensemble des établissements
Suite ........ : dont le dépôt est lié la liste sera
Retour ....... : formatée de la manière suivante "001","002" etc
Mots clefs ... : DEPOT;ETABLISSEMENT
*****************************************************************}

function ListeEtablissementParDepot(depot: string): string;
var Q: TQuery;
  TOBEtab: TOB;
  ListeEtab: string;
  i_ind: integer;
begin
  result := '';
  if GetParamSoc('SO_GCMULTIDEPOTS') then
  begin
    Q := OpenSQL('select ET_ETABLISSEMENT from ETABLISS where (et_depotlie like "%' + depot + '%" or et_depotlie="")', true,-1, '', True);
    if not Q.EOF then
    begin
      TOBEtab := TOB.CREATE('', nil, -1);
      TOBEtab.LoadDetailDB('', '', '', Q, False);
      for i_ind := 0 to TOBEtab.Detail.count - 1 do
      begin
        if ListeEtab <> '' then ListeEtab := ListeEtab + ',';
        ListeEtab := ListeEtab + '"' + TOBEtab.detail[i_ind].GetValue('ET_ETABLISSEMENT') + '"';
      end;
      TOBEtab.free;
    end;
    ferme(Q);
    result := ListeEtab;
  end;
end;

{**********************************************************
Auteur  ...... : C TARDY
Créé le ...... : 19/08/2002
Modifié le ... :
Description .. : fonction renvoyant l'établissement par défaut du site
Suite ........ : s'il contient ce dépôt ou sinon le premier établissement
Suite ........ : qui contient ce dépôt
Mots clefs ... : DEPOT;ETABLISSEMENT
*****************************************************************}

function Etabl_du_depot(depot: string): string;
var Q: TQuery;
  etab: string;
begin
  result := '';
  etab := GetParamSoc('SO_ETABLISDEFAUT');
  Q := OpenSQL('select ET_ETABLISSEMENT from ETABLISS where ET_ETABLISSEMENT="' + etab + '" and (et_depotlie like "%' + depot + '%" or et_depotlie="")', true,-1, '', True);
  if not Q.EOF then result := Q.FindField('ET_ETABLISSEMENT').AsString
  else
  begin
    ferme(Q);
    Q := OpenSQL('select ET_ETABLISSEMENT from ETABLISS where (et_depotlie like "%' + depot + '%" or et_depotlie="")', true,-1, '', True);
    result := Q.FindField('ET_ETABLISSEMENT').AsString;
  end;
  ferme(Q);
end;

function IsArticle (TOBL : TOB) : boolean;
var prefixe,TheValue: string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='ART') or (TheValue='ARV') then result := true;
end;

function IsSousTotal (TOBL : TOB) : boolean;
var prefixe,TheValue: string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='TOT') then result := true;
end;

function IsDebutParagraphe (TOBL : TOB) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (copy(TheValue,1,2)='DP') or (copy(TheValue,1,2)='DV') then result := true;
end;

function IsDebutParagraphe (TOBL : TOB ; niveau : integer) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='DP'+inttostr(niveau)) or
     (TheValue='DV'+inttoStr(niveau)) then result := true;
end;

function IsFinParagraphe (TOBL : TOB) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (copy(TheValue,1,2)='TP') or (copy(TheValue,1,2)='TV') then result := true;
end;

function IsFinParagraphe (TOBL : TOB ; niveau : integer) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='TP'+inttostr(niveau)) or
     (TheValue='TV'+inttoStr(niveau)) then result := true;
end;

function IsDebutParagrapheStd (TOBL : TOB) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (copy(TheValue,1,2)='DP') then result := true;
end;

function IsDebutParagrapheStd (TOBL : TOB ; niveau : integer) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='DP'+inttostr(niveau)) then result := true;
end;

function IsFinParagrapheStd (TOBL : TOB) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (copy(TheValue,1,2)='TP') then result := true;
end;

function IsFinParagrapheStd (TOBL : TOB ; niveau : integer) : boolean;
var prefixe,TheValue : string;
begin
  result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if prefixe = '' then exit;
  TheValue := TOBL.GetValue(prefixe+'_TYPELIGNE');
  if (TheValue='TP'+inttostr(niveau)) then result := true;
end;

function IsParagraphe (TOBL : TOB) : boolean;
begin
  result := false;
  if isDebutParagraphe (TOBL) or IsFinParagraphe (TOBL) then result := true;
end;

function IsParagraphe (TOBL : TOB ; niveau : integer) : boolean;
begin
  result := false;
  if isDebutParagraphe (TOBL,niveau) or IsFinParagraphe (TOBL,niveau) then result := true;
end;

function IsParagrapheStd (TOBL : TOB) : boolean;
begin
  result := false;
  if isDebutParagrapheStd (TOBL) or IsFinParagrapheStd (TOBL) then result := true;
end;

function IsParagrapheStd (TOBL : TOB ; niveau : integer) : boolean;
begin
  result := false;
  if isDebutParagrapheStd (TOBL,niveau) or IsFinParagrapheStd (TOBL,niveau) then result := true;
end;

function IsCommentaire (TOBL : TOB) : boolean;
var TheValue : string;
begin
  result := false;
  TheValue := TOBL.GetValue('GL_TYPELIGNE');
  if (TheValue='COM') or (TheValue='COV') then
    result := true;
end;

function IsSousDetail (TOBL : TOB) : boolean;
var TheValue : string;
begin

  result := false;

  if TOBl = nil then exit;

  IF TOBL.NomTable='LIGNE' then
    TheValue := TOBL.GetValue('GL_TYPELIGNE')
  else
    TheValue := TOBL.GetValue('BLO_TYPELIGNE');

  if (TheValue='SDV') or (TheValue='SD') then
    result := true;
end;

function IsBLobVide (FF : TForm;TOBL : TOB; Champs : string) : boolean;
var texte : THRichEditOLE;
begin
	result := true;
  TRY
    Texte := THRichEditOLE.create (Application.MainForm);
    texte.name := 'MFFFF';
    texte.Parent := FF;
    Texte.Visible := false;
    StringToRich(Texte, TOBL.GetValue(Champs));
    if (Length(texte.Text) <> 0) and (texte.Text <> #$D#$A) then result := false;
  FINALLY
  	texte.Free;
  END;
end;

function IsLignefacture (TOBL : TOB) : boolean;
var Valeur : double;
		Nature : string;
begin

  //correction bug violation d'accès si affectation cotraitant 1ère ligne : 31052012 - FV
  if TOBL = nil then
  begin
    result := false;
    exit;
  end;

  if TOBL.NomTable = 'LIGNE' then
  begin
    Valeur := TOBL.GetValue('GL_QTEPREVAVANC');
    Nature := TOBL.GetValue('GL_NATUREPIECEG');
  end else
  begin
  	if VarIsNull(TOBL.getValue('BLF_QTEDEJAFACT')) or (VarAsType(TOBL.getValue('BLF_QTEDEJAFACT'), varString) = #0) then
    begin
      Valeur := 0;
    end else
    begin
    	Valeur := TOBL.getValue('BLF_QTEDEJAFACT');
  	end;
    Nature := TOBL.getValue('BLO_NATUREPIECEG');
  end;
  result := (Valeur <> 0) and (Nature='DBT');
end;

function IsReferenceLivr (TOBL : TOB) : boolean;
begin
  result := false;
  if (TOBL.GetValue('GL_TYPELIGNE')='RV') then result := true;
end;

function IsNewLigne (TOBL : TOB) : boolean;
begin
  result := false;
  if not TOBL.FieldExists ('NEWONE') then exit;
  if TOBL.GetValue('NEWONE')='X' Then result := true;
end;

{$IFDEF GPAO}
{***********A.G.L.***********************************************
Auteur  ...... : Frédéric BOURRAS
Créé le ...... : 20/11/2002
Modifié le ... :   /  /
Description .. : Appel de fiche de séléction du circuit
Mots clefs ... : CICRUIT
*****************************************************************}

function GetCircuitMul(GS: THGrid; TobPiece: Tob; Article: string): boolean; //TobL:TOB) : boolean ;
var sWhere, CodeCirc: string;
begin
  CodeCirc := GS.Cells[GS.Col, GS.Row];
  sWhere := 'QCI_CIRCUIT IN (SELECT QCI_CIRCUIT ' +
    'FROM WARTNAT ' +
    'LEFT JOIN QCIRCUIT Q1 ON QCI_CODITI=WAN_CODITI ' +
    'WHERE WAN_ARTICLE="' + Article + '" ' +
    'AND WAN_NATURETRAVAIL="FAB" ' +
    'AND NOT EXISTS (SELECT "X" ' +
    'FROM QDETCIRC Q2 ' +
    'LEFT JOIN QSITE ON Q2.QDE_SITE=QSI_SITE AND Q2.QDE_CTX=QSI_CTX ' +
    'WHERE Q2.QDE_CIRCUIT = Q1.QCI_CIRCUIT ' +
    'AND QSI_TIERS<>"' + TOBPiece.GetValue('GP_TIERS') + '" ' +
    'AND Q2.QDE_OPECIRC <> (SELECT MAX(Q3.QDE_OPECIRC) ' +
    'FROM QDETCIRC Q3 ' +
    'WHERE Q3.QDE_CTX=Q1.QCI_CTX ' +
    'AND   Q3.QDE_CIRCUIT = Q1.QCI_CIRCUIT)) ' +
    'AND EXISTS(SELECT "X" ' +
    'FROM QDETCIRC Q4 ' +
    'LEFT JOIN QSITE ON Q4.QDE_SITE=QSI_SITE AND Q4.QDE_CTX=QSI_CTX ' +
    'WHERE Q4.QDE_CIRCUIT = Q1.QCI_CIRCUIT ' +
    'AND QSI_DEPOT="' + TOBPiece.GetValue('GP_DEPOT') + '" ' +
    'AND Q4.QDE_OPECIRC = (SELECT MAX(Q5.QDE_OPECIRC) ' +
    'FROM QDETCIRC Q5 ' +
    'WHERE Q5.QDE_CTX=Q1.QCI_CTX ' +
    'AND   Q5.QDE_CIRCUIT = Q1.QCI_CIRCUIT))) ';
  CodeCirc := AGLLanceFiche('Q', 'QUFMSELECTCIRCUIT', '', '', sWhere);

  if CodeCirc <> '' then
  begin
    Result := True;
    GS.Cells[GS.Col, GS.Row] := CodeCirc;
  end
  else
    Result := False;
end;
{$ENDIF GPAO}

procedure TiersFraisVersPiedPort(TobPiece, TobTiers, TobPL: TOB);
var
  TheTob, TobTF, TobPort: Tob;
  Q: Tquery;
  i: integer;
  DEV : Rdevise;

  function GetTobTF: boolean; // Charge les frais associés au tiers (table TIERSFRAIS)
  begin
    result := false;
    Q := OpenSQL('SELECT * FROM TIERSFRAIS WHERE GTF_NATUREAUXI = "' + TOBTiers.GetValue('T_NATUREAUXI') + '" AND GTF_TIERS = "' + TobTiers.GetValue('T_TIERS')
      + '"', True,-1, '', True);
    try
      if not Q.eof then
      begin
        TobTF.loadDetailDB('TIERSFRAIS', '', '', Q, True, True);
        result := true;
      end;
    finally
      ferme(Q);
    end;
  end;
  { Main }
begin
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  DEV.DateTaux := TOBPiece.GetValue('GP_DATETAUXDEV');
  //
  TobTF := TOB.Create('TIERSFRAIS', nil, -1);
  TobPort := TOB.Create('PORT', nil, -1);

  try
    if (TobPL.Detail.Count <> 0) then
      RazFrais(TobPL); // Suppression des frais associés au tiers
    if GetTobTF then
    begin
      for i := 0 to TobTF.Detail.Count - 1 do
      begin
        if RecherchePort(TOBTF.Detail[i].GetValue('GTF_CODEPORT'), TobPort) then
        begin
          TheTob := Tob.create('PIEDPORT', TobPL, -1);

          TheTob.PutValue('GPT_CODEPORT', TobTF.Detail[i].GetValue('GTF_CODEPORT'));
          ChargeTobPiedPort(TheTob, TobPort, TobPiece);
          TheTob.AddChampSup('TIERSFRAIS', False);
        end;
      end;
    end;
  finally
    TobTF.Free;
    TobPort.Free;
  end;
end;

procedure RazFrais(TobPL: Tob);
var
  i: integer;
begin
  { Suppression des frais associés au tiers - Crées en automatique }
  for i := TobPL.Detail.Count - 1 downto 0 do
  begin
    if TobPL.Detail[i].FieldExists('TIERSFRAIS') then
      TobPL.Detail[i].free;
  end;
  { Pour les frais qui ont été crées manuellement }
  if TobPL.Detail.Count <> 0 then
  begin
    if PGIAsk('Des frais ont été affectés à la pièce. Voulez-vous les supprimer ?', 'Frais') = MrYes then
      TobPL.ClearDetail
  end;
end;

function RecherchePort(CodePort: string; TobPort: Tob): boolean;
var
  Q: Tquery;
  DateSup: TDateTime;
  sQl: string;
begin
  Result := false;
  sQl := 'Select * from PORT Where GPO_CODEPORT="' + CodePort + '"'
    ;
  Q := OpenSQL(sQL, True,-1, '', True);
  if not Q.EOF then
  begin
    if Q.FindField('GPO_DATESUP').AsString <> '' then
    begin
      DateSup := StrtoDate(Q.FindField('GPO_DATESUP').AsString);
      if ((DateSup < V_PGI.DateEntree) and (DateSup > iDate1900)) or (Q.FindField('GPO_FERME').AsString = 'X') then
        Result := false
      else
      begin
        TobPort.SelectDB('', Q);
        Result := True;
      end;
    end
    else
    begin
      Result := True;
      PGIBox('Anomalie : Code port : "' + CodePort + '" / Date de suppression logique non renseignée','RECHERCHEPORT');
    end
  end;
  Ferme(Q);
end;


procedure ChargeTobPiedPort(TobPiedPort, TobPort, TobPiece: TOB);
var
  TypePort: string;
  Base: Double;
  EnHT: boolean;
  DEV : Rdevise;
begin
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  TobPiedPort.PutValue('GPT_LIBELLE', TobPort.GetValue('GPO_LIBELLE'));
  TypePort := TobPort.GetValue('GPO_TYPEPORT');
  TobPiedPort.PutValue('GPT_TYPEPORT', TypePort);

  if TOBPiece.getString('GP_VENTEACHAT')='VEN' then
  begin
    TobPiedPort.SetString('GPT_COLLECTIF',TobPort.GetString('GPO_COLLECTIF'));
  end else if TOBPiece.getString('GP_VENTEACHAT')='ACH' then
  begin
    TobPiedPort.SetString('GPT_COLLECTIF',TobPort.GetString('GPO_COLLECTIFAC'));
  end;

  TobPiedPort.SetBoolean('GPT_RETENUEDIVERSE',TobPort.getBoolean('GPO_RETENUEDIVERSE'));
  TobPiedPort.PutValue('GPT_FAMILLETAXE1', TobPort.GetValue('GPO_FAMILLETAXE1'));
  TobPiedPort.PutValue('GPT_FAMILLETAXE2', TobPort.GetValue('GPO_FAMILLETAXE2'));
  TobPiedPort.PutValue('GPT_FAMILLETAXE3', TobPort.GetValue('GPO_FAMILLETAXE3'));
  TobPiedPort.PutValue('GPT_FAMILLETAXE4', TobPort.GetValue('GPO_FAMILLETAXE4'));
  TobPiedPort.PutValue('GPT_FAMILLETAXE5', TobPort.GetValue('GPO_FAMILLETAXE5'));
  TobPiedPort.PutValue('GPT_COMPTAARTICLE', TobPort.GetValue('GPO_COMPTAARTICLE'));
  TobPiedPort.AddChampSupValeur('DEJA_CALCULE', '-', False);
  TobPiedPort.PutValue('GPT_BLOQUEFRAIS', TobPort.GetValue('GPO_VERROU'));

  EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
  if EnHT then base := TobPort.GetValue('GPO_MINIMUM') else base := TobPort.GetValue('GPO_MINIMUMTTC');
  Base := Arrondi(PivotToSaisie(DEV, Base, DEV.Taux, DEV.Quotite, DEV.Decimale), DEV.decimale);
  TobPiedPort.PutValue('GPT_MINIMUM', Base);
  TobPiedPort.PutValue('GPT_FRANCO', TobPort.GetValue('GPO_FRANCO'));
  TobPiedPort.PutValue('GPT_MONTANTMINI', 0);
  TobPiedPort.PutValue('GPT_FRAISREPARTIS', TobPort.GetValue('GPO_FRAISREPARTIS'));
  //
  if (Pos(TypePort,'HT;PT;MI;MIC;;')>0) then
  begin
    TobPiedPort.PutValue('GPT_POURCENT', TobPort.GetValue('GPO_COEFF'));
  end;

  if (Pos(TypePort, 'MT;MTC')>0) then
  begin
    TobPiedPort.PutValue('GPT_POURCENT', 0);
  if TypePort = 'MT' then
  begin
      base := TobPort.GetValue('GPO_PVHT');
    end else
    begin
      base := TobPort.GetValue('GPO_PVTTC');
    end;
    Base := Arrondi(PivotToSaisie(DEV, Base, DEV.Taux, DEV.Quotite, DEV.Decimale), DEV.decimale);
    if TypePort = 'MT' then
    begin
      TobPiedPort.PutValue('GPT_TOTALHTDEV', Base);
  end else
    begin
      TobPiedPort.PutValue('GPT_TOTALTTCDEV', Base);
    end;
  end;
  
  if (Pos(TypePort, 'MI;MIC')>0) then
  begin
  if TypePort = 'MI' then
  begin
      base := TobPort.GetValue('GPO_PVHT');
    end else
    begin
      base := TobPort.GetValue('GPO_PVTTC');
    end;
    Base := Arrondi(PivotToSaisie(DEV, Base, DEV.Taux, DEV.Quotite, DEV.Decimale), DEV.decimale);
    TobPiedPort.PutValue('GPT_MONTANTMINI', Base);
  end;

  TobPiedPort.PutValue('GPT_NATUREPIECEG', TOBPiece.GetValue('GP_NATUREPIECEG'));
  TobPiedPort.PutValue('GPT_SOUCHE', TOBPiece.GetValue('GP_SOUCHE'));
  TobPiedPort.PutValue('GPT_NUMERO', TOBPiece.GetValue('GP_NUMERO'));
  TobPiedPort.PutValue('GPT_INDICEG', TOBPiece.GetValue('GP_INDICEG'));
  TobPiedPort.PutValue('GPT_DEVISE', TOBPiece.GetValue('GP_DEVISE'));
  TobPiedPort.PutValue('GPT_NUMPORT', 0);
  if (Pos(TypePort ,'MT;MTC')>0) then
  begin
    TobPiedPort.PutValue('GPT_BASEHT', 0);
    TobPiedPort.PutValue('GPT_BASETTC', 0);
    TobPiedPort.PutValue('GPT_BASEHTDEV', 0);
    TobPiedPort.PutValue('GPT_BASETTCDEV', 0);
  end else
  begin
      Base := TOBPiece.Somme('GL_TOTALHT', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTALESC');
      TobPiedPort.PutValue('GPT_BASEHT', Base);
      Base := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTALESCDEV');
      TobPiedPort.PutValue('GPT_BASEHTDEV', Base);
      Base := TOBPiece.Somme('GL_TOTALTTC', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTESCTTC');
      TobPiedPort.PutValue('GPT_BASETTC', Base);
      Base := TOBPiece.Somme('GL_TOTALTTCDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTESCTTCDEV');
      TobPiedPort.PutValue('GPT_BASETTCDEV', Base);
    end;
end;

function PivotToSaisie(TheDEV: RDevise; Base, Taux, Quotite: Double; Decimale: integer): double;
begin
  Result := Base;
  if (TheDEV.Code <> V_PGI.DevisePivot) then
    Result := EuroToDevise(Base, Taux, Quotite, Decimale)
end;

procedure RecalculPiedPort(Action: TActionFiche; TOBPiece, TOBPiedPort,TOBBases: Tob; MontantbaseFixe : double=0);
var
  i: integer;
  Base,baseTTc,BaseAch: Double;
  EnHT: boolean;
  TypePort: string;
  TobPied: Tob;
  Abandon: boolean;
  DEV : Rdevise;
begin
  //
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  if Action = taConsult then Exit;
  if (TOBPiece = nil) or (TOBPiedPort = nil) then exit;
  TOBPied := TOB.Create('', nil, -1);
  TOBPied.Dupliquer(TOBPiedPort, True, True);
  DEV.Code := TOBPiece.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  DEV.DateTaux := TOBPiece.GetValue('GP_DATETAUXDEV');
	EnHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
  BaseAch := TOBPiece.getValue('GP_MONTANTPA')+TOBPiece.getValue('GP_MONTANTFG')+TOBPiece.getValue('GP_MONTANTFC');
  for i := 0 to TobPied.Detail.count - 1 do
  begin
    Abandon := false;
    BaseTTC := TOBPiece.GetValue('GP_TOTALTTCDEV');
    TypePort := TOBPied.Detail[i].GetValue('GPT_TYPEPORT');
{$IFDEF BTP}
		if MontantBaseFixe <> 0 then
    begin
    	Base := MontantbaseFixe;
    end else
    begin
    	Base := TOBPiece.GetValue('GP_TOTALHTDEV');
    end;
{$ELSE}
    Base := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False);
{$ENDIF}
    // Recalcul base si ligne de frais non bloquée
    // ou pour les frais non encore calculés }
    if Action = TaCreat then
    begin
      if (TobPied.Detail[i].GetValue('GPT_BLOQUEFRAIS') = 'VER') then
      begin
            Abandon := True;
      end
    end
    else
      if (TobPied.Detail[i].GetValue('GPT_BLOQUEFRAIS') <> 'NON') then abandon := True;

    if not abandon then
    begin
      if (Pos(TypePort,'MT;MTC')=0) then
      begin
        if TobPied.detail[i].GetValue('GPT_FRAISREPARTIS')<>'X' then
        begin
          TobPied.detail[i].PutValue('GPT_BASEHTDEV', Base);
          TobPied.detail[i].PutValue('GPT_BASETTCDEV', BaseTTC);
          CalculBasePiedPort(TobPied.Detail[i], (TobPied.detail[i].GetValue('GPT_FRAISREPARTIS')<>'X'));
        end else
        begin
          // frais repartis
          TobPied.detail[i].PutValue('GPT_BASEHTDEV', BaseAch);
        end;
      end;
      //
      CalculMontantsPiedPort(TOBPiece, TOBPied.detail[i],TOBBases);
      if (Action = TaCreat) and (Base > 0) then
      begin
        if TobPied.Detail[i].FieldExists('DEJA_CALCULE') then
          TobPied.Detail[i].PutValue('DEJA_CALCULE', 'X')
        else
          TobPied.AddChampSupValeur('DEJA_CALCULE', 'X', False);
      end
    end
    else
    begin
      CalculMontantsPiedPort(TOBPiece, TOBPied.detail[i],TOBBases);
  end;
  end;
  TOBPiedPort.Dupliquer(TOBPied, True, True);
  TOBPied.Free;
end;

procedure CalculBasePiedPort(TOBPL: TOB; EnHT: boolean);
var Base: Double;
  TypePort: string;
  DEV : Rdevise;
begin
	DEV.Code := TOBPL.GetValue('GPT_DEVISE');
  GetInfosDevise(DEV);  if DEV.Taux = 0 then DEV.taux := 1; 
  //
  TypePort := TOBPL.GetValue('GPT_TYPEPORT');
  if (Pos(TypePort,'MT;MTC')=0) then
  begin
    if EnHT then
    begin
      Base := TOBPL.GetValue('GPT_BASEHTDEV');
      TOBPL.PutValue('GPT_BASEHT', DeviseToEuro(base, DEV.Taux, DEV.Quotite));
    end
    else
    begin
      Base := TOBPL.GetValue('GPT_BASETTCDEV');
      TOBPL.PutValue('GPT_BASETTC', DeviseToEuro(base, DEV.Taux, DEV.Quotite));
    end;
  end;
end;

procedure ArticleDetailversLigne(refSais : String; TOBart,TOBL : TOB; SaContexte : TModeAff);
Begin

  TOBArt.PutValue('REFARTSAISIE', RefSais);
  TOBL.PutValue('GL_ARTICLE', TOBArt.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBArt.GetValue('REFARTSAISIE'));
  TOBL.PutValue('GL_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_REFARTBARRE', TOBArt.GetValue('REFARTBARRE'));
  TOBL.PutValue('BNP_TYPERESSOURCE', TOBArt.GetValue('BNP_TYPERESSOURCE'));

  if not (tModeSaisieBordereau in SaContexte) then TOBL.PutValue('GL_REFARTTIERS', TOBArt.GetValue('REFARTTIERS'));
  TOBL.PutValue('GL_TYPEREF', 'ART');
{$IFNDEF BTP}
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    if (TOBArt.GetValue('REFARTSAISIE') = TOBArt.GetValue('REFARTBARRE')) and not SaisieCodeBarreAvecFenetre then
    begin
      if not TOBL.FieldExists('REGROUPE_CB') then TOBL.AddChampSup('REGROUPE_CB', False);
      if not TOBL.FieldExists('UNI_OU_DIM') then TOBL.AddChampSup('UNI_OU_DIM', False);
      TOBL.PutValue('REGROUPE_CB', '-');
      if TOBArt.GetValue('GA_STATUTART') = 'DIM' then TOBL.PutValue('UNI_OU_DIM', 'DIM')
      else TOBL.PutValue('UNI_OU_DIM', 'UNI');
    end;
  end;
{$ENDIF}
end;


procedure CalculMontantsPiedPort(TOBPiece, TOBPL,TOBBases: TOB);
var Base, Total, Pour, Montmini: Double;
  TypePort: string;
  PieceHT: Boolean;
  DEV : Rdevise;
  TOBBASesRD : TOB;
begin
  TOBBASesRD := TOB.Create ('LES BASES',nil,-1);
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  PieceHT := (TOBPiece.GetValue('GP_FACTUREHT') = 'X');
  TypePort := TOBPL.GetValue('GPT_TYPEPORT');
  if (TypePort = 'HT') or (TypePort='') then
  begin
    if TOBPL.GetValue('GPT_FRAISREPARTIS')='X' then
    begin
      // Dans le cas des frais repartis on prend le montant achat + Montant FG + Montant FC comme base
      TOBPL.putValue('GPT_BASEHTDEV',TOBPiece.getValue('GP_MONTANTPA')+TOBPiece.getValue('GP_MONTANTFG')+TOBPiece.getValue('GP_MONTANTFC'));
      Base := TOBPL.GetValue('GPT_BASEHTDEV');
      Pour := TOBPL.GetValue('GPT_POURCENT');
      Total := Arrondi(Base * Pour / 100.0, DEV.Decimale);
    end else
    begin
      Base := TOBPL.GetValue('GPT_BASEHTDEV');
      Pour := TOBPL.GetValue('GPT_POURCENT');
      Total := Arrondi(Base * Pour / 100.0, DEV.Decimale);
      Base := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTALESCDEV');
      if TOBPL.GetValue('GPT_FRANCO') = 'X' then
        if Base >= TOBPL.GetValue('GPT_MINIMUM') then Total := 0;
    end;
    TOBPL.PutValue('GPT_TOTALHTDEV', Total);
  end else if TypePort = 'PT' then
    begin
      Pour := TOBPL.GetValue('GPT_POURCENT');
      Base := TOBPiece.Somme('GL_TOTALTTCDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTESCTTCDEV');
    Total := Arrondi(Base * Pour / 100.0, DEV.Decimale);
    if TOBPL.GetValue('GPT_FRANCO') = 'X' then
      if Base >= TOBPL.GetValue('GPT_MINIMUM') then Total := 0;
    TOBPL.PutValue('GPT_TOTALTTCDEV', Total);
  end else if TypePort = 'MI' then
  begin
      Base := TOBPL.GetValue('GPT_BASEHTDEV');
      Pour := TOBPL.GetValue('GPT_POURCENT');
      Montmini := TOBPL.GetValue('GPT_MONTANTMINI');
      Total := Arrondi(Base * Pour / 100.0, DEV.Decimale);
      if Total < Montmini then Total := Montmini;
      Base := TOBPiece.Somme('GL_TOTALHTDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTALESCDEV');
    if TOBPL.GetValue('GPT_FRANCO') = 'X' then
      if Base >= TOBPL.GetValue('GPT_MINIMUM') then Total := 0;
    TOBPL.PutValue('GPT_TOTALHTDEV', Total);
  end else if TypePort = 'MIC' then
    begin
      Base := TOBPL.GetValue('GPT_BASETTCDEV');
      Pour := TOBPL.GetValue('GPT_POURCENT');
      Montmini := TOBPL.GetValue('GPT_MONTANTMINI');
      Total := Arrondi(Base * Pour / 100.0, DEV.Decimale);
      if Total < Montmini then Total := Montmini;
      Base := TOBPiece.Somme('GL_TOTALTTCDEV', ['GL_TYPELIGNE'], ['ART'], False);
      Base := Base + TOBPiece.GetValue('GP_TOTESCTTCDEV');
    if TOBPL.GetValue('GPT_FRANCO') = 'X' then
      if Base >= TOBPL.GetValue('GPT_MINIMUM') then Total := 0;
    TOBPL.PutValue('GPT_TOTALTTCDEV', Total);
  end else if (TypePort='MT') or (TypePort='MTC') then
  begin
    if TypePort = 'MT' then
    begin
      TOBPL.PutValue('GPT_TOTALHTDEV', TOBPL.GetValue('GPT_BASEHTDEV'));
    end else
    begin
      TOBPL.PutValue('GPT_TOTALTTCDEV', TOBPL.GetValue('GPT_BASETTCDEV'));
  end;
  end;

  if DEV.Code <> V_PGI.DevisePivot then
  begin
      Total := TOBPL.GetValue('GPT_TOTALHTDEV');
      TOBPL.PutValue('GPT_TOTALHT', Arrondi(DeviseToEuro(Total, DEV.Taux, DEV.Quotite), DEV.Decimale));
      Total := TOBPL.GetValue('GPT_TOTALTTCDEV');
      TOBPL.PutValue('GPT_TOTALTTC', Arrondi(DeviseToEuro(Total, DEV.Taux, DEV.Quotite), DEV.Decimale));
  end else
  begin
      Total := TOBPL.GetValue('GPT_TOTALHTDEV');
      TOBPL.PutValue('GPT_TOTALHT', Total);
      Total := TOBPL.GetValue('GPT_TOTALTTCDEV');
      TOBPL.PutValue('GPT_TOTALTTC', Total);
    end;
  //
  if TOBPL.GetBoolean('GPT_RETENUEDIVERSE') then
  begin
    ConstitueTVARetenue (TOBPiece,TOBPL,TOBBAses,TOBBASesRD,DEV);
    RecalcRetenue (TOBPL,TOBBasesRD);
  end;
  TOBBASesRD.free;
end;

function TestDepotOblig : boolean;
begin
  if (CtxScot in V_PGI.PGIContexte) or (CtxTempo in V_PGI.PGIContexte) or (CtxChr in V_PGI.PGIContexte) then
    Result := False
    else
    Result := True;
end;

function PieceFacturation : string;
begin
	result := 'FF;FAC;FBT;DAC;ABT;ABP;';
end;

function IsTransformable(NewNature : string) : boolean;
var ThePieceSuiv : string;
		UnePieceSUiv : string;
begin
	result := false;
	ThePieceSuiv := GetInfoParPiece(NewNature,'GPP_NATURESUIVANTE');
  if ThePieceSuiv = '' then exit;
  repeat
  	UnePieceSuiv := READTOKENST (ThePieceSUiv);
    if UnePieceSuiv <> '' then
    begin
    	if UnePieceSuiv <> NewNature then BEGIN result := true; break; END;
    end;
  until UnePieceSuiv = '';
end;


procedure positionneCoefMarge (TOBL : TOB);
var EnHt : boolean;
begin
	EnHt := (TOBL.GEtVAlue('GL_FACTUREHT')='X');
//  if not TOBL.FieldExists ('__COEFMARG') then TOBL.AddChampSupValeur ('__COEFMARG',0);
  if TOBL.GetValue('GL_DPR') = 0 then TOBL.PutValue('GL_DPR',TOBL.GetValue('GL_DPA'));
  if TOBL.GetValue('GL_DPR') = 0 then
  begin
//    TOBL.PutValue('GL_COEFMARG',0)
  end else
  begin
     {if EnHt then }TOBL.PutValue ('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_PUHTDEV')/TOBL.GetValue('GL_DPR'),4));
//             else TOBL.PutValue ('GL_COEFMARG',Arrondi(TOBL.GetValue('GL_PUTTCDEV')/TOBL.GetValue('GL_DPR'),9));
  end;
end;


function IsprestationST(LaTOb : TOB) : boolean;
begin
	result := false;
  if IsLigneSoustraite(LaTOb) then
  begin
    result := true;
    exit;
  end;
  if not LaTOB.FieldExists('BNP_TYPERESSOURCE') then exit;
	if LaTOb.NomTable = 'LIGNE' then
  begin
  	if (LaTOB.GetValue('GL_TYPEARTICLE')='PRE') and (LaTOB.GetValue('BNP_TYPERESSOURCE')='ST') then result := true;
  end else
  begin if LaTOb.NomTable = 'LIGNEOUV' then
  	if (LaTOB.GetValue('BLO_TYPEARTICLE')='PRE') and (LaTOB.GetValue('BNP_TYPERESSOURCE')='ST') then result := true;
  end;
end;

procedure ValideLesBases(TOBPiece,TobBases,TOBBasesL : TOb);
var Indice : integer;
    TOBB : TOB;
begin
  for Indice := 0 to TOBBases.detail.count -1 do
  begin
    TOBB := TOBBases.detail[Indice];
    TOBB.PutValue('GPB_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
    TOBB.PutValue('GPB_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
    TOBB.PutValue('GPB_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
    TOBB.PutValue('GPB_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
    TOBB.PutValue('GPB_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
    TOBB.PutValue('GPB_SAISIECONTRE',TOBPiece.GetValue('GP_SAISIECONTRE')) ;
  end;
  for Indice := 0 to TOBBasesL.detail.count -1 do
  begin
    TOBB := TOBBasesL.detail[Indice];
    if TOBB.detail.count > 0 then
    begin
      ValideLignesBases(TOBpiece,TOBB);
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 01/04/2003
Description .. : Etat de la ligne
*****************************************************************}
function GetEtatSolde(TobLigne: Tob): T_EtatSoldeLigne;
begin
  Result := eslNone;
  if Assigned(TobLigne) then
  begin
    { ici, seul GL_ETATSOLDE est nécessaire. En cas de changement, impacter le SELECT de la fonction GetEtatSolde(CleDoc) }
    if TobLigne.GetString('GL_ETATSOLDE') = 'ENC' then Result := eslEnCours
    else if TobLigne.GetString('GL_ETATSOLDE') = 'SOL' then Result := eslSolde
    else if TobLigne.GetString('GL_ETATSOLDE') = 'VER' then Result := eslVerrouille
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 14/06/2006
Description .. : Idem à la précédente mais par requête SQL
*****************************************************************}
function GetEtatSolde(const CleDoc: R_CleDoc): T_EtatSoldeLigne;
var
  TobLigne: Tob;
begin
  TobLigne := Tob.Create('GL', nil, -1);
  try
    if wSelectTobFromSQL('SELECT GL_ETATSOLDE FROM LIGNE WHERE ' + WherePiece(CleDoc, ttdLigne, True, True), TobLigne, False) then
      Result := GetEtatSolde(TobLigne)
    else
      Result := eslNone
  finally
    TobLigne.Free
  end
end;

procedure GPDoActionAfterUpdatePiece(const CleDoc: R_CleDoc);
begin
  {$IFDEF EDI}
    EDIAfterCreateOrUpdatePiece(CleDoc)
  {$ENDIF EDI}
end;

procedure AppliqueDomaineActivite (TOBL,TOBA : TOB);
begin
	if TOBL.GetValue('GL_DOMAINE') = '' then exit;
  
end;

function IsFromExcel (TOBL : TOB) : boolean;
begin
  result := false;
  if TOBL = nil then exit;
  if not TOBL.fieldExists('BLX_NUMLIGXLS') then exit;
  if VarIsNull(TOBL.GetValue('BLX_NUMLIGXLS')) or (VarAsType(TOBL.GetValue('BLX_NUMLIGXLS'), varString) = #0 ) then
  begin
  	TOBL.PutValue('BLX_NUMLIGXLS',0);
  end;
  if TOBL.GetValue('BLX_NUMLIGXLS')<>0 then result := true;
end;

procedure RecupTOBArticle (TOBPiece,TOBArticles,TOBAFormule,TOBConds : TOB; CledocAffaire,Cledoc : R_cledoc; VenteAchat : string);
var i, NbArt : integer;
  StSelect, StSelectDepot, StSelectDepotLot, StSelectCon, {StArticle,}
  StSelectCatalogu, StSelectFormule, {StCodeArticle,} stWhereLigne, ListeDepot,
  NomChamp, VteAch {, RefGen, Statut}: string;
  TOBDispo, TOBDispoLot, TOBArt, TobDispoArt, TobDispoLotArt, TOBCata, TobCataArt, TobAF: TOB;
  QArticle, QDepot, QDepotLot, QCon, QCata, QFormule : TQuery;
  {OkTab,} DoubleDepot: Boolean;
begin

//
  StSelect := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
              'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
              'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE';
//
//  StSelect := 'SELECT * FROM ARTICLE LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GA_ARTICLE';
  StSelectDepot := 'SELECT * FROM DISPO LEFT JOIN ARTICLE ON GQ_ARTICLE=GA_ARTICLE ';
  StSelectDepotLot := 'SELECT * FROM DISPOLOT';
  StSelectCon := 'SELECT * FROM CONDITIONNEMENT';
  StSelectCatalogu := 'SELECT * FROM CATALOGU';
  StSelectFormule := 'SELECT * FROM ARTICLEQTE';

  ListeDepot := '';
  DoubleDepot := IsTransfert(TOBPiece.GetValue('GP_NATUREPIECEG'));
  if DoubleDepot then
    ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '","' + TOBPiece.GetValue('GP_DEPOTDEST') + '"'
  else ListeDepot := '"' + TOBPiece.GetValue('GP_DEPOT') + '"';

  TOBDispo := TOB.CREATE('Les Dispos', nil, -1);
  TOBDispoLot := TOB.CREATE('Les Dispolots', nil, -1);
  TOBCata := TOB.CREATE('Les Catalogues', nil, -1);

{$IFDEF AFFAIRE}
// gestion de la saisie de facture depuis l'affaire avec reprise des lignes de l'affaire
// il faut charger les articles issus des lignes d'affaire  gm 01/07/03
  if (CleDocAffaire.NumeroPiece = 0) then
    stWhereLigne := WherePiece(CleDoc, ttdLigne, False)
  else
    stWhereLigne := WherePiece(CleDocAffaire, ttdLigne, False);
{$else}
  stWhereLigne := WherePiece(CleDoc, ttdLigne, False) ;
{$endif}

  QArticle := OpenSQL(StSelect + ' WHERE GA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ') ORDER BY GA_ARTICLE', True,-1, '', True);
  if not QArticle.EOF then TOBArticles.LoadDetailDB('ARTICLE', '', '', QArticle, True, True);
  Ferme(QArticle);

  if DoubleDepot then
    QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (' + ListeDepot + ') AND GQ_CLOTURE="-" AND GA_TENUESTOCK="X"', True,-1, '', True)
  else QDepot := OpenSQL(StSelectDepot + ' WHERE GQ_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GQ_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE ' + stWhereLigne + ') AND GQ_CLOTURE="-" AND GA_TENUESTOCK="X"', True,-1, '', True);
  if not (CtxMode in V_PGI.PGIContexte) then
  begin
    if DoubleDepot then
      QDepotLot := OpenSQL(StSelectDepotLot + ' WHERE GQL_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
        ') AND GQL_DEPOT IN (' + ListeDepot + ')', True,-1, '', True)
    else QDepotLot := OpenSQL(StSelectDepotLot + ' WHERE GQL_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
        ') AND GQL_DEPOT IN (SELECT DISTINCT GL_DEPOT FROM LIGNE WHERE ' + stWhereLigne + ')', True,-1, '', True);
    if not QDepotLot.EOF then TOBDispoLot.LoadDetailDB('DISPOLOT', '', '', QDepotLot, True, True);
    Ferme(QDepotLot);
  end;
  if not QDepot.EOF then TOBDispo.LoadDetailDB('DISPO', '', '', QDepot, True, True);
  Ferme(QDepot);

  (*
  //Formules de qté  JS 30/10/2003
  QFormule := OpenSQL(StSelectFormule + ' WHERE GAF_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ')', True,-1, '', True);
  if not QFormule.EOF then TOBAFormule.LoadDetailDB('ARTICLEQTE', '', '', QFormule, True, True);
  Ferme(QFormule);
  QCon := OpenSQL(StSelectCon + ' WHERE GCO_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne +
    ') ORDER BY GCO_NBARTICLE', True,-1, '', True);
  if not QCon.EOF then TOBConds.LoadDetailDB('CONDITIONNEMENT', '', '', QCon, True, True);
  Ferme(QCon);
  *)
  if (VenteAchat = 'ACH') then
  begin
    QCata := OpenSQL(StSelectCatalogu + ' WHERE GCA_TIERS IN (SELECT DISTINCT GL_TIERS FROM LIGNE WHERE ' + stWhereLigne +
      ') AND GCA_ARTICLE IN (SELECT GL_ARTICLE FROM LIGNE WHERE ' + stWhereLigne + ' AND GL_ARTICLE<>"" )', True,-1, '', True);
    if not QCata.EOF then TOBCata.LoadDetailDB('CATALOGU', '', '', QCata, True, True);
    Ferme(QCata);
  end;

  //Affecte les stocks aux articles sélectionnés
  for NbArt := 0 to TOBArticles.detail.Count - 1 do
  begin
    with TOBArticles.detail[NbArt] do
    begin
      AddChampSupValeur('REFARTSAISIE', '', False);
      AddChampSupvaleur('REFARTBARRE', '', False);
      AddChampSupValeur('REFARTTIERS', '', False);
      AddChampSupValeur('_FROMOUVRAGE', '-', False);
      AddChampSupValeur('SUPPRIME', '-', False);
      AddChampSupValeur('UTILISE', '-', False);
    end;
    // FIN MODIF DC

    TOBArt := TOBArticles.detail[NbArt];
    //
    InitChampsSupArticle (TOBART);
    //
    TobDispoArt := TOBDispo.FindFirst(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    //if TobDispoArt<>nil then //Modif du 18/07/02: Bug chez Negoce
    while TobDispoArt <> nil do
    begin
      TobDispoLotArt := TOBDispoLot.FindFirst(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TobDispoArt.GetValue('GQ_DEPOT')], False);
      while TobDispoLotArt <> nil do
      begin
        TobDispoLotArt.Changeparent(TobDispoArt, -1);
        TobDispoLotArt := TOBDispoLot.FindNext(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TobDispoArt.GetValue('GQ_DEPOT')], False);
      end;
      TobDispoArt.Changeparent(TOBArt, -1);
      TobDispoArt := TOBDispo.FindNext(['GQ_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
      if DoubleDepot then
      begin
        //TobDispoArt:=TOBDispo.FindNext(['GQ_ARTICLE'],[TOBArt.GetValue('GA_ARTICLE')],False);
        if TobDispoArt <> nil then
        begin
          TobDispoLotArt := TOBDispoLot.FindFirst(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TOBPiece.GetValue('GP_DEPOTDEST')],
            False);
          while TobDispoLotArt <> nil do
          begin
            TobDispoLotArt.Changeparent(TobDispoArt, -1);
            TobDispoLotArt := TOBDispoLot.FindNext(['GQL_ARTICLE', 'GQL_DEPOT'], [TobDispoArt.GetValue('GQ_ARTICLE'), TOBPiece.GetValue('GP_DEPOTDEST')],
              False);
          end;
          TobDispoArt.Changeparent(TOBArt, -1);
        end;
        break;
      end;
    end;
    //Affecte les catalogu aux articles sélectionnés
    TobCataArt := TOBCata.FindFirst(['GCA_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    while TobCataArt <> nil do
    begin
      for i := 1 to TobCataArt.NbChamps do
      begin
        NomChamp := TobCataArt.GetNomChamp(i);
        if NomChamp = '' then continue;
        TOBArt.AddChampSup(NomChamp, False);
        TOBArt.PutValue(NomChamp, TobCataArt.GetValue(NomChamp));
      end;
      TOBArt.AddChampSup('DPANEW', False);
      TOBArt.PutValue('DPANEW', TOBArt.GetValue('GCA_DPA'));
      TobCataArt := TOBCata.FindNext(['GCA_ARTICLE'], [TOBArt.GetValue('GA_ARTICLE')], False);
    end;
  end;

  //Formules de qté  JS 30/10/2003
  if VenteAchat='VEN' then VteAch:='VTE'
  else VteAch := VenteAchat;
  for i := 0 to TobAFormule.Detail.Count -1 do
  begin
    TobAF := TobAFormule.Detail[i];
    QFormule := OpenSql('SELECT * FROM ARTFORMULEVAR WHERE GAV_ARTICLE="'+
                   TobAF.GetValue('GAF_ARTICLE')+'" AND GAV_VENTEACHAT="'+VteAch+'"',True,-1, '', True);
    if not QFormule.Eof then TobAF.LoadDetailDB('ARTICLEQTE','','',QFormule,False) ;
    Ferme(QFormule);
  end;

  TOBDispo.Free;
  TOBDispoLot.Free;
  TOBCata.Free;
  DispoChampSupp(TOBArticles);
end;

function IsPieceAutoriseRemligne(TOBpiece : TOB) : boolean;
begin
	Result := True;
	if TOBpiece.GetString('GP_VENTEACHAT')='ACH' then Exit; // pour les Pièces d'achats c'est OK
  if (TOBpiece.GetString('GP_VENTEACHAT')='VEN') and (Pos(TOBPiece.getString('GP_NATUREPIECEG'),'DE;CC;BLC;FBC;ABC;BBO;')>0) then Exit;
  Result := False;
end;


function CalculeQteSais (TOBL : TOB) : double;
var prefixe : string;
begin
  Prefixe := GetPrefixeTable(TOBL);
  Result:= TOBL.GeTDouble(prefixe+'_QTEFACT');
  if not GetParamSocSecur('SO_BTGESTQTEAVANC',false) or ( not ISDocumentChiffrage(TOBL)) then exit;
  //
  if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST') > 0 then
  begin
    if TOBL.GetDouble(prefixe+'_RENDEMENT')<> 0 then
    begin
      result := TOBL.GeTDouble(prefixe+'_QTEFACT')* TOBL.GetDouble(prefixe+'_RENDEMENT');
    end else
    begin
     result := TOBL.GeTDouble(prefixe+'_QTEFACT');
    end;
  end else
  begin
    if TOBL.GetDouble(prefixe+'_PERTE')<> 0 then
    begin
      result := Arrondi(TOBL.GeTDouble(prefixe+'_QTEFACT')/ TOBL.GetDouble(prefixe+'_PERTE'),V_PGI.okdecQ);
    end else
    begin
      Result := TOBL.GeTDouble(prefixe+'_QTEFACT');
    end;
  end;
end;

function CalculeQteFact (TOBL : TOB) : double;
var prefixe : string;
begin
  Prefixe := GetPrefixeTable(TOBL);
  Result := TOBL.GetDouble(prefixe+'_QTESAIS');
  if not GetParamSocSecur('SO_BTGESTQTEAVANC',false) or ( not ISDocumentChiffrage(TOBL)) then exit;
  //
  if POS(TOBL.GetString('BNP_TYPERESSOURCE'),'SAL;INT;ST')> 0 then
  begin
    if TOBL.GetDouble(prefixe+'_RENDEMENT')<> 0 then
    begin
      Result := Arrondi(TOBL.GeTDouble(prefixe+'_QTESAIS')/ TOBL.GetDouble(prefixe+'_RENDEMENT'),V_PGI.okdecQ);
    end else
    begin
      Result := TOBL.GetDouble(prefixe+'_QTESAIS');
    end;
  end else
  begin
    if TOBL.GetDouble(prefixe+'_PERTE')<> 0 then
    begin
      result := Arrondi(TOBL.GeTDouble(prefixe+'_QTESAIS')* TOBL.GetDouble(prefixe+'_PERTE'),V_PGI.okdecQ);
    end else
    begin
      Result := TOBL.GetDouble(prefixe+'_QTESAIS');
    end;
  end;
end;

function ISDocumentChiffrage(NaturePiece : string) : boolean; overload;
begin
  result := (Pos(NaturePiece,'DBT;ETU;BCE')>0);
end;


function ISDocumentChiffrage(TOBL : TOB) : boolean; overload;
var prefixe : string;
begin
  prefixe := GetPrefixeTable(TOBL);
  result := ISDocumentChiffrage(TOBL.GetString(prefixe+'_NATUREPIECEG'));
end;

procedure InitConditionnements (TOBL : TOB);
var TheTiers : string;
    VenteAchat : string;
    QQ : TQuery;
begin
  VenteAchat := GetInfoParPiece(TOBL.GetString('GL_NATUREPIECEG'), 'GPP_VENTEACHAT');
  if VenteAchat = 'ACH' then TheTiers := TOBL.getString('GL_TIERS') else TheTiers := '';
  QQ:= OpenSql ('SELECT * FROM CONDITIONNEMENT '+
                'WHERE '+
                'GCO_TYPESFLUX="'+VenteAchat+'" AND GCO_TIERS="'+TheTiers+'" AND GCO_ARTICLE="'+TOBL.GetString('GL_ARTICLE')+'" AND GCO_DEFAUT="X"',True,1,'',true);
  if not QQ.eof then
  begin
    TOBL.SetString('GL_CODECOND',QQ.FindField('GCO_CODECOND').AsString);
    TOBL.SetDouble('GLC_QTECOND',1);
    TOBL.SetDouble('GLC_COEFCOND',QQ.FindField('GCO_NBARTICLE').asfloat);
    TOBL.SetDouble('GL_QTEFACT',ARRONDI(TOBL.GetDouble('GLC_COEFCOND'),V_PGI.okdecQ));
    TOBL.SetDouble('GL_QTESTOCK',TOBL.GetDouble('GL_QTEFACT'));
    TOBL.SetDouble('GL_QTERESTE',TOBL.GetDouble('GL_QTEFACT'));
  end;
  Ferme(QQ);
end;

end.

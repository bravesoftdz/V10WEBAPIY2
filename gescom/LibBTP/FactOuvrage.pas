unit FactOuvrage;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} EdtDoc, DB, Fe_Main,
{$ENDIF}
{$IFDEF BTP}
     BTPUtil,
{$ENDIF}
     FactTOB,FactArticle,FactPiece,forms,Windows,Graphics,
     SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil,factcalc,factcomm,factcpta,
     Math, StockUtil, EntGC, Classes, HMsgBox, Choix,NomenUtil,LigNomen,paramsoc,TarifUtil,utofbtanaldev,
     UtilArticle,OptimizeOuv,FactDomaines,uEntCommun,UtilTOBPiece, UtilsMetresXLS, UtilsCopierCollerMetresXLS;

var TZoneGlob : array [1..35] of string = ('DATEPIECE',
                                           'AFFAIRE',
                                           'AFFAIRE1',
                                           'AFFAIRE2',
                                           'AFFAIRE3',
                                           'AVENANT',
                                           'REPRESENTANT',
                                           'APPORTEUR',
                                           'COMMISSIONR',
                                           'COMMISSIONA',
                                           'TYPECOM',
                                           'VALIDECOM',
                                           'FOURNISSEUR',
                                           'PIECEPRECEDENTE',
                                           'PIECEORIGINE',
                                           'DEPOT',
                                           'TYPELIGNE',
                                           'TYPEREMISE',
                                           'DEVISE',
                                           'CREATEUR',
                                           'UTILISATEUR',
                                           'REGIMETAXE',
                                           'FACTUREHT',
                                           'TVAENCAISSEMENT',
                                           'TAUXDEV',
                                           'COTATION',
                                           'ESCOMPTE',
                                           'REMISEPIED',
                                           'REMISELIGNE',
                                           'REMISECASCADE',
                                           'DATECREATION',
                                           'DATEMODIF',
                                           'CODECOMPTAAFF',
                                           'CREERPAR',
                                           'SOCIETE');

var TZoneArt : array [1..24] of string = ('FAMILLETAXE1',
                                          'FAMILLETAXE2',
                                          'FAMILLETAXE3',
                                          'FAMILLETAXE4',
                                          'FAMILLETAXE5',
                                          'TYPEARTICLE',
                                          'TYPENOMENC',
                                          'TENUESTOCK',
                                          'FAMILLENIV1',
                                          'FAMILLENIV2',
                                          'FAMILLENIV3',
                                          'COLLECTION',
                                          'LIBREART1',
                                          'LIBREART2',
                                          'LIBREART3',
                                          'LIBREART4',
                                          'LIBREART5',
                                          'LIBREART6',
                                          'LIBREART7',
                                          'LIBREART8',
                                          'LIBREART9',
                                          'LIBREARTA',
                                          'PRIXPOURQTE',
                                          'PCB');

var TZoneLig : array [1..17] of string = ('SOUCHE',
                                          'TIERS',
                                          'TIERSFACTURE',
                                          'TIERSPAYEUR',
                                          'TIERSLIVRE',
                                          'REFCOLIS',
                                          'PIECEPRECEDENTE',
                                          'TARIF',
                                          'TARIFTIERS',
                                          'ETABLISSEMENT',
                                          'VIVANTE',
                                          'TAUXRFA',
                                          'NBTRANSMIS',
                                          'DATELIVRAISON',
                                          'NUMADRESSELIVR',
                                          'SAISIECONTRE',
                                          'DOMAINE');
{$IFDEF BTP}
var ArtPrec: string;
{$ENDIF}

procedure PositionneOuvrageOut (TOBPiece,TOBOuvrage,TOBOuvragesP : TOB; ARow : integer);
procedure CalculMontantHtDevLigOuv (TOBOuv:TOb;DEV:Rdevise);
procedure InsertionChampSupOuv (TOBLN: TOb;MultiNiveau:boolean);
Procedure RenseigneValoOuvrage ( TOBLN : TOB ;CodeDT: String; Qte,QteDuDetail: Double;DEV:RDEVISE;var Valeurs:T_Valeurs) ;
Procedure LoadLesOuvrages ( TOBPiece,TOBOuvrage,TOBArticles : TOB ; CleDoc : R_CleDoc ;Var IndiceNomen : integer; OptimizeOuv : TOptimizeOuv = nil; Forprint : boolean=false);
function TraiteLesOuvrages ( XX: Tform; TOBPiece,TOBArticles,TOBOuvrage,TOBRepart, TobMetres : TOB ; ARow : integer ; LePremier : boolean ;DEV :RDEVISE;OrigineExcel:boolean;OptimizeOuv : TOptimizeOuv = nil;WithDetails : boolean=true;FromBordereau : boolean=false; ModeSaisieAch : boolean= false) : boolean;
Procedure ValideLesOuv ( TOBOuv : TOB ; TOBPiece : TOB = nil) ;
Procedure CreerLesTOBOuv (TOBL,TOBGroupeNomen,TOBNomen,TOBArticles : TOB ; LaLig,idep : integer;OptimizeOuv : TOptimizeOuv = nil;AvecStock : boolean=true;WithValo : boolean=true ; Forprint : boolean=false; WithLigneFac : boolean=false) ;
function ReAffecteLigOuv (IndiceNomen : integer ; TOBLig,TOBOuvrage : TOB ) : boolean ;
procedure CalculeOuvrageDoc (TOBOuvrage : TOB;Qte,QteDetail:double;recursif:boolean;DEV:RDevise ; var valeurs:T_Valeurs;EnHt : boolean;AvecHt:boolean=true;WithArrondi : boolean=false);
function ReajusteMontantOuvrage (TOBARticles,TOBPiece,TOBL,TOBOuvrage:TOB;Montantcalc: double;var MontantPr : double;var MontantDemande:double;DEV:RDevise;EnHt : boolean;atributionEcart:boolean=true;EnPR : boolean = false;EnPa : boolean=false;ForcePaEqualPv:boolean=false) : boolean;
procedure CopieOuvFromArt (TOBPiece,TOBLND,TOBART : TOB;DEV:Rdevise; RecupPrix : string='PUH');
procedure CopieOuvFromLigne(TOBLND,TOBL: TOB);
procedure DefiniPrixLigneOuv (TOBouvrage,TOBNomen,TOBART,TOBPiece : TOb;DEV:Rdevise);
procedure CopieOuvFromRef(TOBOuv,TOBRef: TOB);
procedure AttribueEcart (TOBPiece,TOBL,TobOuvrage : TOB;ArtEcart:string;Montantecart: double;DEV:RDEVISE;EnHt:Boolean;EnPA : boolean=false;OnlyPa:boolean=false;PrixPourQte : double=0; LibEcart : string='';ForcePaEqualPv : Boolean=false);
procedure LoadLesLibDetOuvLig ( TOBPIECE,TOBOuvrage,TOBTiers,TOBAffaire,TOBL, TOBTRFPOC: TOB;var Indice: integer; DEV:RDevise; TheMetreDoc : TMetreArt; AffSousDetailUnitaire : boolean);
procedure LoadLesLibDetailOuvrages (TOBPIECE,TOBOuvrage,TOBTiers,TOBAffaire,TOBTRFPOC : TOB;DEV:RDevise; TheMetreDoc : TMetreArt);
procedure DropLesLibDetailouvrages (TOBPiece : TOB);
procedure MiseAPlatOuv (TOBPiece,TOBL,TOBOuv,TOBPlat : TOB; TousNiveaux:boolean; GardeReference : boolean=false; WithCalcul : boolean=true; SurLigneStd : boolean=false;PourCpta : boolean=false);
Procedure CalculeLigneHTOuv ( TOBOuv,TOBPiece : TOB ; DEV:RDevise ; Forcer:boolean=False ) ;
Procedure CalculeLigneTTCOuv ( TOBOuv,TOBPiece : TOB ; DEV:RDevise; Forcer:boolean=False) ;
procedure AppliqueChangeTaxeOuv (TOBPiece,TOBOuvrage : TOB;Ligne:integer;Taxe1,Taxe2,Taxe3,Taxe4,Taxe5 : string);
procedure AppliqueChangeRegimeOuv (TOBPiece,TOBOuvrage : TOB; Ligne : integer; newCode : string);
procedure AffecteValoFromOuv (TOBOuv : TOB;DEV : RDevise;var Valeurs : T_Valeurs);
procedure PutValueDetailOuv (TOBOuvrage : TOB;Champ:string; Valeur : variant; TOBarticles:TOb=nil);
//NewOne
procedure RecupValoDetailHrs (TOBRef: TOB;CalculPv : boolean=true);
procedure GetValoDetail (TOBOuv:TOB);
procedure InitValoDetail (TOBPiece,TOBOuv,TOBRef : TOB; DEV : Rdevise;CalculPv:boolean=true; RecupPrix : string='PUH');
procedure RecupValoDetail (TOBOuv,TOBRef : TOB);
procedure CalculOuvFromDetail (TOBOuv : TOB;DEV : RDevise ; var Valeur: T_Valeurs);
function LookUpFournisseur (CodFourn, CodArtic ,LibArt: string; SelectFourSiPasCataLog : boolean=false) : string;
procedure positionneEltTarif (TOBOUV,TOBREF: TOB);
procedure AjouteSousDetail (TOBL,TOBOuvrage,TOBArticle : TOB;MarchandiseUniquement : boolean=true);
procedure SuprimeSousDetail(TOBL,TOBOUvrage:TOB);
procedure AddLesReferences (TOBIns,TOBLoc : TOB);
procedure InitChampSupOuv (TOBLN : TOB);
procedure GestionDetailPrixPose (TOBDepart : TOB);
procedure TraiteRecupValoGrp ( TOBOuvrage,TOBArticles : TOB; var Valeur : T_Valeurs);
procedure PositionneLigneOuvVariante (TOBL,TOBOuvrage : TOB ; variante : boolean);
procedure PositionneVarianteOuv (TobOuv : TOB;variante : boolean);
function IsOuvrage (TOBL : TOB) : boolean;
function IsOuvrageComptabiliseDetail (TOBL : TOB) : boolean;
function IsArticlePrixPose (TOBL : TOB) : boolean;
function GetQteDetailOuv (TOBL,TOBOUvrages : TOB) : integer;
procedure PosQteDetailAndCalcule (TOBLig,TOBOUvrages : TOB;QteDetail : double;DEV: RDevise);
(* OPTIMIZATION *)
procedure RecupValorisationOUV (TOBOuv,TOBRef : TOB; OptimizeOuv : ToptimizeOuv);
(* ----------- *)
function ApplicPrixPose (TOBPiece : TOB) : boolean;
procedure BeforeRecalculPrOuv (TOBOuv : TOB; EnHt : boolean;CalculPv : boolean=true;InclusSTinFG : boolean=false);
Procedure CalculeLigneHTOuvMere ( TOBL,TOBBases,TOBPiece : TOB ; DEv : Rdevise;NbDec : integer;arrondiok:boolean;var MontantLignes:double) ;
Procedure CalculeLigneTTCOuvMere ( TOBL,TOBBases,TOBPiece : TOB ; DEV: Rdevise; NbDec : integer;ArrondiOk : boolean;var MontantLignes:double);
//--
procedure EssayeAffecterEcartQte (TOBL,TOBOuvrage : TOB; Ecart,EcartPr : double ; var SumLoc: double ; var SumPr : double;
																  IndPa,IndPR,IndPVHtDev,IndPVHt : integer; EnPA,EnPr,EnHt : boolean; DEV : Rdevise;
                                  TouslesNiveaux : boolean= false; ForcePaEqualPv : Boolean=false);
procedure AffecteDifferenceSurLigne (TOBL,TOBLOC : TOB;Ecart,EcartPr:double;
																		 IndPa,IndPR,IndPVHtDev,IndPVHt : integer;
                                     EnPa,EnPR,EnHt : boolean; DEv : RDevise;ForcePaEqualPv:Boolean=false);

procedure AftercalculPrOuv (TOBOuv : TOB; EnHt : boolean;CalculPv : boolean=true);
Procedure ChargeLesOuvrages ( TOBPiece,TOBOuvrage: TOB; CleDoc : R_CleDoc);
//procedure DocMetreBiblioDetail(TOBMere,TOBL:TOB);
//procedure DocMetreBiblioEntete(TOBL: TOB; CodeArt,TypeArt:String);
procedure recupInfoLigneOuv (TOBDEST,TOBRef : TOB);
//procedure VaChercherQuantitatifDetail (TObDepart : TOB);
procedure ReinitCoefPaPrOuv (TOBL,TOBouvrage : TOB);
procedure SetAffaireSousDetail (TOBPiece,TOBOUvrage : TOB);
procedure InformeErreurOuvrage (FF : TForm ; TOBL : TOB; Arect : Trect);
procedure CopieLigFromRef(TOBLig,TOBRef: TOB);
procedure MultipleLookUpFournisseur (TOBLiaison : TOB );
procedure CalculMontantsDocOuv(TOBPiece,TOBL,TOBOuvrage : TOB; CalculPv : boolean=true; FromImportBordereaux : boolean = false);
procedure GereCalculMontantOuv (TOBPiece,TOBL,TOBOuvrage : TOb; var valeur : T_Valeurs; calculpv : boolean=true);
procedure BeforeApplicFraisDetailOuv (TOBL,TOBOuvrage : TOB; RazFg : boolean=false);
procedure AppliqueModeGestionOuvFg (TOBL,TOBOuvrage : TOB ; ZoneAgerer : string; ModeGestion : boolean);
procedure CalculeLigneAcOuv (TOBOUv : TOB ; DEV : Rdevise;WithPuv : boolean=true; TOBART : TOB=nil);
procedure ValideLesArticlesFromOuv (TOBArticles,TOBOuvrages : TOB);
procedure GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers : TOB;DEv : Rdevise; PourCpta: boolean = true; SansEcart : Boolean=false);
function FindOuvragesPlat(TOBL, TOBOuvrageP: TOB ) : TOB;
function NewTobLignePlat (TOBmere: TOB; NumOrdre : integer; position : integer;SurLigneStd : boolean=false) : TOB;
procedure LoadLesOuvragesPlat(TOBPiece, TOBOuvragesP: TOB; Cledoc : R_CLEDOC );
procedure ValideLesOuvPlat(TOBOuvragesP, TOBPiece : TOB);
function AddMereLignePlat (TOBMere: TOB;NumOrdre : integer) : TOB;
Procedure MajLigOuv (TOBN,TOBLig : TOB) ;
Procedure NumeroteLigneOuv (TOBN,TOBLig : TOB; NiveauCur,LigPere1,LigPere2, LigPere3,LigPere4: integer ) ;
procedure ReinitCoefMarg (TOBL,TOBOuvrage : TOB);
procedure changeTaxeOuvInterne (TOBOuvrage : TOB;Taxe1,Taxe2,Taxe3,Taxe4,Taxe5 : string);
procedure changeRegimeOuvInterne (TOBOuvrage : TOB;regime : string);

procedure AppliqueCoefMargDetail (TOBL,TOBOuvrage: TOB; CoefMarg : double; DEV : Rdevise);
procedure AppliqueChangeFraisDetOuv (TOBPiece,TOBO : TOB; NewGestionfrais,NewgestionFC,NewgestionFA : string);
procedure recalculCoefligneMere (TOBOUV : TOB);
procedure CalculeLigneAcOuvCumul (TOBOUv : TOB);
procedure AffecteTempsPoseOuvrage (TOBArticles,TOBL: TOB ; temps : double; DEV : RDevise);
function GetQualifTps(TOBL,TOBOUvrage : TOB) : string;
procedure AjoutePrestationPrixPose (TOBPiece,TOBLigPiece,TOBOP,TOBArticles : TOB; PrestDefaut : string; DEV : RDevise );
function OuvrageNonDifferencie (TOBPiece,TOBOuvrage : TOB) : boolean;
procedure  OuvrageDifferencie (TOBpiece,TOBOuvrage : TOB; TraiteAvancOuv : boolean; DEV : RDevise);
function EncodeRefPieceOuv (TOBO : TOB) : string;
procedure DecodeRefPieceOUv(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
procedure SetAvancementOuvrage(TOBGroupeNomen,TOBOD,TOBL : TOB; DEV : RDevise);
procedure SetLignesFactureOuvDetail (TOBL,TOBOuvrage,TOBLFactures : TOB);
//FV1 : 29022012
Procedure RecalculeLigneOuv(TOBOP, TOBL : TOB; DEV:RDevise; EnHt : boolean);

function GetTvaOuvrage (TOBl,TOBOuvrage : TOB) : string;
function CalculePrixRevientOuv (TOBOUV,TOBART : TOB; PUA,COEFFG : double; Var MontantFg,MontantFc,MontantFr : double) : double;
procedure PositionneLigneOuvOrigine (TOBOuvrage : TOB);
//
procedure AplatOuvrage (TOBpiece,TOBL,TOBOuv,TOBPlat: TOB;Qte,QteDuDetail: double;TousNiveaux:boolean;GardeReference:boolean=false; WithCalcul : boolean=true; SurLigneStd : boolean=false; PourCpta : boolean=false);
procedure RecupLesOuvrages(TOBPiece,TOBOuvrages : TOB);
procedure SetInfoPrecOuv (TOBOuvrage : TOB);
procedure  AjouteSousDetailVide(TOBL, TOBOuvrage : TOB);
procedure SupprimeSDOuvVide (TOBL,TOBOuvrage : TOB);
procedure InformePTCOuvrage (FF : TForm ; TOBL : TOB; Arect : Trect);
//
procedure CumuleLesTypes (TOBL : TOB; var Valloc : T_Valeurs);
procedure StockeLesTypes (TOBL : TOB; Valloc : T_Valeurs);
//
Procedure DerouleOuvrage ( TOBN : TOB ; CodeNomen,RefUnique,Depot : String ; TOBArticles : TOB ; QteDuDetail : Double;DEV:RDevise;VenteAchat : string) ;
Procedure OuvVersLigOuv ( TOBPIece,TOBL,TOBNOMEN,TOBLN,TOBA : TOB ; Niv,OrdreCompo : integer ;DEV:RDevise;EnHt:Boolean;Prixtraite : string='PUH' ) ;
function FindLigneOuvFromUnique (TOBL,TOBOuvrages : TOB; UniqueBlo : Integer) : TOB;
//
implementation
uses factgrp,FactVariante,Facture,FactureBTP
{$IFDEF BTP}
//,UMetreArticle
//,UtilMetres,
 ,PiecesRecalculs
 ,BTStructChampSup
 ,Ucotraitance
 ,UspecifPOC
{$ENDIF}
  ,CbpMCD
  ,CbpEnumerator

;

procedure InitArticleFromLigneNomen (TOBL,TOBA : TOB);
begin
  TOBA.putValue('GA_ARTICLE',TOBL.getvalue('GNL_ARTICLE'));
  TOBA.putValue('GA_CODEARTICLE',TOBL.getvalue('GNL_CODEARTICLE'));
  TOBA.putValue('GA_LIBELLE',TOBL.getvalue('GNL_LIBELLE'));
  TOBA.putValue('SUPPRIME','X');
end;

procedure InitArticleFromLigneOUV (TOBL,TOBA : TOB);
begin
  TOBA.putValue('GA_ARTICLE',TOBL.getvalue('BLO_ARTICLE'));
  TOBA.putValue('GA_CODEARTICLE',TOBL.getvalue('BLO_CODEARTICLE'));
  TOBA.putValue('GA_LIBELLE',TOBL.getvalue('BLO_LIBELLE'));
  TOBA.putValue('GA_PAHT',TOBL.getvalue('BLO_DPA'));
  TOBA.putValue('GA_DPR',TOBL.getvalue('BLO_DPR'));
  TOBA.putValue('GA_PVHT',TOBL.getvalue('BLO_PUHT'));
  TOBA.putValue('GA_TENUESTOCK',TOBL.getvalue('BLO_TENUESTOCK'));
  TOBA.putValue('GA_ESCOMPTABLE',TOBL.getvalue('BLO_ESCOMPTABLE'));
  TOBA.putValue('SUPPRIME','X');
end;

procedure StockeLesTypes (TOBL : TOB; Valloc : T_Valeurs);
var Qte,QteDuDetail : double;
begin
  if TOBL.NomTable = 'LIGNE' then
  begin
    TOBL.PutValue('GLC_MTMOPA',ValLoc[23]);
    TOBL.PutValue('GLC_MTMOPR',ValLoc[24]);
    TOBL.PutValue('GLC_MTMOPV',ValLoc[25]);
    TOBL.PutValue('GLC_MTFOUPA',ValLoc[26]);
    TOBL.PutValue('GLC_MTFOUPR',ValLoc[27]);
    TOBL.PutValue('GLC_MTFOUPV',ValLoc[28]);
    TOBL.PutValue('GLC_MTINTPA',ValLoc[29]);
    TOBL.PutValue('GLC_MTINTPR',ValLoc[30]);
    TOBL.PutValue('GLC_MTINTPV',ValLoc[31]);
    TOBL.PutValue('GLC_MTLOCPA',ValLoc[32]);
    TOBL.PutValue('GLC_MTLOCPR',ValLoc[33]);
    TOBL.PutValue('GLC_MTLOCPV',ValLoc[34]);
    TOBL.PutValue('GLC_MTMATPA',ValLoc[35]);
    TOBL.PutValue('GLC_MTMATPR',ValLoc[36]);
    TOBL.PutValue('GLC_MTMATPV',ValLoc[37]);
    TOBL.PutValue('GLC_MTOUTPA',ValLoc[38]);
    TOBL.PutValue('GLC_MTOUTPR',ValLoc[39]);
    TOBL.PutValue('GLC_MTOUTPV',ValLoc[40]);
    TOBL.PutValue('GLC_MTSTPA',ValLoc[41]);
    TOBL.PutValue('GLC_MTSTPR',ValLoc[42]);
    TOBL.PutValue('GLC_MTSTPV',ValLoc[43]);
    TOBL.PutValue('GLC_MTAUTPA',ValLoc[44]);
    TOBL.PutValue('GLC_MTAUTPR',ValLoc[45]);
    TOBL.PutValue('GLC_MTAUTPV',ValLoc[46]);
  end else if TOBL.NomTable = 'LIGNEOUV' then
  begin
    QteDuDetail := TOBL.GetDouble('BLO_QTEDUDETAIL'); if QteDuDetail = 0 then QteDuDetail := 1;
    Qte := TOBL.GetDouble('BLO_QTEFACT'); Qte := Qte / QteDuDetail;
    TOBL.PutValue('BLO_MTMOPA',Arrondi(Qte*ValLoc[23],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTMOPR',Arrondi(Qte*ValLoc[24],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTMOPV',Arrondi(Qte*ValLoc[25],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTFOUPA',Arrondi(Qte*ValLoc[26],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTFOUPR',Arrondi(Qte*ValLoc[27],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTFOUPV',Arrondi(Qte*ValLoc[28],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTINTPA',Arrondi(Qte*ValLoc[29],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTINTPR',Arrondi(Qte*ValLoc[30],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTINTPV',Arrondi(Qte*ValLoc[31],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTLOCPA',Arrondi(Qte*ValLoc[32],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTLOCPR',Arrondi(Qte*ValLoc[33],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTLOCPV',Arrondi(Qte*ValLoc[34],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTMATPA',Arrondi(Qte*ValLoc[35],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTMATPR',Arrondi(Qte*ValLoc[36],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTMATPV',Arrondi(Qte*ValLoc[37],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTOUTPA',Arrondi(Qte*ValLoc[38],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTOUTPR',Arrondi(Qte*ValLoc[39],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTOUTPV',Arrondi(Qte*ValLoc[40],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTSTPA',Arrondi(Qte*ValLoc[41],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTSTPR',Arrondi(Qte*ValLoc[42],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTSTPV',Arrondi(Qte*ValLoc[43],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTAUTPA',Arrondi(Qte*ValLoc[44],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTAUTPR',Arrondi(Qte*ValLoc[45],V_PGI.okdecV));
    TOBL.PutValue('BLO_MTAUTPV',Arrondi(Qte*ValLoc[46],V_PGI.okdecV));
  end;
end;

procedure CumuleLesTypes (TOBL : TOB; var Valloc : T_Valeurs);
var Qte,QteDuDetail : double;
begin
  QteDuDetail := TOBL.GetDouble('BLO_QTEDUDETAIl'); if QTeDuDetail = 0 then QteDuDetail := 1;
  Qte := TOBL.GetDouble('BLO_QTEFACT') / QteDuDetail;
  ValLoc[23] := ValLoc[23] + Arrondi(TOBL.GetDouble('BLO_MTMOPA') * Qte,V_PGI.okdecP);
  ValLoc[24] := ValLoc[24] + Arrondi(TOBL.GetDouble('BLO_MTMOPR')  * Qte,V_PGI.okdecP);
  ValLoc[25] := ValLoc[25] + Arrondi(TOBL.GetDouble('BLO_MTMOPV')  * Qte,V_PGI.okdecP);
  ValLoc[26] := ValLoc[26] + Arrondi(TOBL.GetDouble('BLO_MTFOUPA') * Qte,V_PGI.okdecP);
  ValLoc[27] := ValLoc[27] + Arrondi(TOBL.GetDouble('BLO_MTFOUPR') * Qte,V_PGI.okdecP);
  ValLoc[28] := ValLoc[28] + Arrondi(TOBL.GetDouble('BLO_MTFOUPV') * Qte,V_PGI.okdecP);
  ValLoc[29] := ValLoc[29] + Arrondi(TOBL.GetDouble('BLO_MTINTPA') * Qte,V_PGI.okdecP);
  ValLoc[30] := ValLoc[30] + Arrondi(TOBL.GetDouble('BLO_MTINTPR') * Qte,V_PGI.okdecP);
  ValLoc[31] := ValLoc[31] + Arrondi(TOBL.GetDouble('BLO_MTINTPV') * Qte,V_PGI.okdecP);
  ValLoc[32] := ValLoc[32] + Arrondi(TOBL.GetDouble('BLO_MTLOCPA') * Qte,V_PGI.okdecP);
  ValLoc[33] := ValLoc[33] + Arrondi(TOBL.GetDouble('BLO_MTLOCPR') * Qte,V_PGI.okdecP);
  ValLoc[34] := ValLoc[34] + Arrondi(TOBL.GetDouble('BLO_MTLOCPV') * Qte,V_PGI.okdecP);
  ValLoc[35] := ValLoc[35] + Arrondi(TOBL.GetDouble('BLO_MTMATPA') * Qte,V_PGI.okdecP);
  ValLoc[36] := ValLoc[36] + Arrondi(TOBL.GetDouble('BLO_MTMATPR') * Qte,V_PGI.okdecP);
  ValLoc[37] := ValLoc[37] + Arrondi(TOBL.GetDouble('BLO_MTMATPV') * Qte,V_PGI.okdecP);
  ValLoc[38] := ValLoc[38] + Arrondi(TOBL.GetDouble('BLO_MTOUTPA') * Qte,V_PGI.okdecP);
  ValLoc[39] := ValLoc[39] + Arrondi(TOBL.GetDouble('BLO_MTOUTPR') * Qte,V_PGI.okdecP);
  ValLoc[40] := ValLoc[40] + Arrondi(TOBL.GetDouble('BLO_MTOUTPV') * Qte,V_PGI.okdecP);
  ValLoc[41] := ValLoc[41] + Arrondi(TOBL.GetDouble('BLO_MTSTPA') * Qte,V_PGI.okdecP);
  ValLoc[42] := ValLoc[42] + Arrondi(TOBL.GetDouble('BLO_MTSTPR') * Qte,V_PGI.okdecP);
  ValLoc[43] := ValLoc[43] + Arrondi(TOBL.GetDouble('BLO_MTSTPV') * Qte,V_PGI.okdecP);
  ValLoc[44] := ValLoc[44] + Arrondi(TOBL.GetDouble('BLO_MTAUTPA') * Qte,V_PGI.okdecP);
  ValLoc[45] := ValLoc[45] + Arrondi(TOBL.GetDouble('BLO_MTAUTPR') * Qte,V_PGI.okdecP);
  ValLoc[46] := ValLoc[46] + Arrondi(TOBL.GetDouble('BLO_MTAUTPV') * Qte,V_PGI.okdecP);
end;
//////////////////////////////////////////
{Fonction de calcul des lignes d'ouvrage}
//////////////////////////////////////////
procedure CalculMontantHtDevLigOuv (TOBOuv:TOb;DEV:Rdevise);
var MontantHt,MontantAch,MontantTTC,MontantPR,MontantHtDev,MontantTTCDev,DPR,PUHT,MontantPA,pa : double;
begin
  if TOBOUV.GetValue('BLO_PRIXPOURQTE') = 0 then TOBOUV.PutValue('BLO_PRIXPOURQTE',1);
  DPR := TobOuv.getValue('BLO_DPR');
  PUHT := TobOuv.getValue('BLO_PUHT');
  PA := TobOuv.getValue('BLO_DPA');
  MontantPA :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*PA)/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  MontantHt :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*PuHT)/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.OkdecP);
  MontantTTC :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*TobOuv.getValue('BLO_PUTTC'))/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  MontantHtDev :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*TobOuv.getValue('BLO_PUHTDEV'))/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  MontantTTCDev :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*TobOuv.getValue('BLO_PUTTCDEV'))/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  MontantAch :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*TobOuv.getValue('BLO_DPA'))/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  MontantPR :=Arrondi((TOBOuv.GetValue('BLO_QTEFACT')*DPR)/TOBOUV.GetValue('BLO_PRIXPOURQTE'),V_PGI.okdecP);
  TOBOuv.PutValue('BLO_MONTANTPA',    Arrondi(MontantPA,V_PGI.okdecP));
  TOBOuv.PutValue('BLO_MONTANTHT',    Arrondi(MontantHt,V_PGI.okdecP));
  TOBOuv.PutValue('BLO_MONTANTTTC',   Arrondi(MontantTTc,V_PGI.okdecP));
  TOBOuv.PutValue('BLO_MONTANTHTDEV', Arrondi(MontantHtDev,V_PGI.okdecP));
  TOBOuv.PutValue('BLO_MONTANTTTCDEV',Arrondi(MontantTTcDev,V_PGI.okdecP));
  //
  if TOBOUV.detail.count = 0 then
  begin
    StockeMontantTypeSurLigne (TOBOUV);
  end;
  //
  if DPR <> 0 then
  begin
    TOBOuv.PutValue('BLO_COEFMARG',   Arrondi(PuHT/DPR,4));
    TOBOuv.PutValue('POURCENTMARG',   Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
  end;
  if PUHT <> 0 then
  begin
    TOBOuv.PutValue('POURCENTMARQ',   Arrondi(((PUHT- DPR)/PUHT)*100,2));
  end else
  begin
    TOBOuv.PutValue('POURCENTMARQ',0);
  end;

end;


procedure InitChampSupOuv (TOBLN : TOB);
var Indice : integer;
begin
//TOBLN.PutValue ('UTILISE','X') ;
TOBLN.PutValue ('UTILISE','-') ;
TOBLN.PutValue ('ANCPV',0);
//TOBLN.PutValue ('MONTANTHTDEV',0);
//TOBLN.PutValue ('MONTANTTTCDEV',0);
//TOBLN.PutValue ('MONTANTACHAT',0);
//TOBLN.PutValue ('MONTANTPR',0);
TobLN.PutValue ('SOUSNOMEN','');
TobLN.PutValue ('LIBELLE','');
TOBLN.PutValue ('CODENOMEN','');
// NEWONE
TOBLN.PutValue ('TGA_FOURNPRINC','');
TOBLN.PutValue ('GA_FOURNPRINC','');
TOBLN.PutValue ('GCA_PRIXBASE',0);
TOBLN.PutValue ('GCA_QUALIFUNITEACH','');
TOBLN.PutValue ('GF_CALCULREMISE','');
TOBLN.PutValue ('GA_PAHT',0);
TOBLN.PutValue ('GA_PVHT',0);
TOBLN.PutValue ('GA_NATUREPRES','');
TOBLN.PutValue ('GA_ARTICLE','');
TOBLN.PutValue ('TGA_PRESTATION','');
TOBLN.PutValue ('GA_HEURE',0);
TOBLN.PutValue ('GF_PRIXUNITAIRE',0);
TOBLN.PutValue ('INDICELIENDEVCHA',0);
//TOBLN.PutValue ('TPSUNITAIRE',0);
TOBLN.PutValue ('ANCPA',0);
TOBLN.PutValue ('ANCPR',0);
  for Indice:= 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
     TOBLN.PutValue ('MILLIEME'+IntToStr(indice+1),0);
  end;
end;

procedure InsertionChampSupOuv (TOBLN: TOb;MultiNiveau:boolean);
var Indice : integer;
begin
  AddlesChampsSupOuv (TOBLN);
(*
if not TOBLN.FieldExists('BNP_TYPERESSOURCE') then
begin
	TOBLN.AddChampSupValeur ('BNP_TYPERESSOURCE','');
end else
begin
	if (VarIsNull (TOBLN.GetValue ('BNP_TYPERESSOURCE'))) or
  	 (VarAsType(TOBLN.GetValue ('BNP_TYPERESSOURCE'),VarString)=#0) then
  begin
  	TOBLN.putValue('BNP_TYPERESSOURCE','');
  end;
end;
if not TOBLN.FieldExists('BNP_LIBELLE') then
begin
	TOBLN.AddChampSupValeur ('BNP_LIBELLE','');
end else
begin
	if (VarIsNull (TOBLN.GetValue ('BNP_LIBELLE'))) or
  	 (VarAsType(TOBLN.GetValue ('BNP_LIBELLE'),VarString)=#0) then
  begin
  	TOBLN.putValue('BNP_LIBELLE','');
  end;
end;

if not TOBLN.FieldExists('LIBELLEFOU') then
begin
	TOBLN.AddChampSupValeur ('LIBELLEFOU','');
end else
begin
	if (VarIsNull (TOBLN.GetValue ('LIBELLEFOU'))) or
  	 (VarAsType(TOBLN.GetValue ('LIBELLEFOU'),VarString)=#0) then
  begin
  	TOBLN.putValue('LIBELLEFOU','');
  end;
end;
*)
//TOBLN.AddChampSupValeur ('UTILISE','X',Multiniveau) ;
TOBLN.AddChampSupValeur ('UTILISE','-',Multiniveau) ;
TOBLN.addchampsupValeur ('ANCPV',0,Multiniveau);
//TOBLN.addchampsupValeur ('MONTANTHTDEV',0,MultiNiveau);
//TOBLN.addchampsupValeur ('MONTANTTTCDEV',0,MultiNiveau);
//TOBLN.addchampsupValeur ('MONTANTTTCDEV',0,MultiNiveau);
//TOBLN.addchampsupValeur ('MONTANTACHAT',0,MultiNiveau);
//TOBLN.addchampsupValeur ('MONTANTPR',0,MultiNiveau);
TobLN.addchampsupValeur ('SOUSNOMEN','',MultiNiveau);
TobLN.addchampsupValeur ('LIBELLE','',MultiNiveau);
TOBLN.AddChampSupValeur ('CODENOMEN','',Multiniveau);
// NEWONE
TOBLN.AddChampSupValeur ('TGA_FOURNPRINC','',Multiniveau);
TOBLN.AddChampSupValeur ('GA_FOURNPRINC','',Multiniveau);
TOBLN.AddChampSupValeur ('GCA_PRIXBASE',0,Multiniveau);
TOBLN.AddChampSupValeur ('GCA_QUALIFUNITEACH','',Multiniveau);
TOBLN.AddChampSupValeur ('GF_CALCULREMISE','',Multiniveau);
TOBLN.AddChampSupValeur ('GA_PAHT',0,Multiniveau);
TOBLN.AddChampSupValeur ('GA_PVHT',0,Multiniveau);
TOBLN.AddChampSupValeur ('GA_NATUREPRES','',Multiniveau);
TOBLN.AddChampSupValeur ('GA_ARTICLE','',Multiniveau);
TOBLN.AddChampSupValeur ('TGA_PRESTATION','',Multiniveau);
TOBLN.AddChampSupValeur ('GA_HEURE',0,Multiniveau);
TOBLN.AddChampSupValeur ('GF_PRIXUNITAIRE',0,Multiniveau);
TOBLN.AddChampSupValeur ('INDICELIENDEVCHA',0,Multiniveau);
TOBLN.addchampsupValeur ('ANCPA',0,Multiniveau);
TOBLN.addchampsupValeur ('ANCPR',0,Multiniveau);
// --
//TOBLN.AddChampSupValeur ('TPSUNITAIRE',0,Multiniveau);
TOBLN.AddChampSupValeur ('INDICEMETRE',0,Multiniveau);
// ------
//TOBLN.AddChampSupValeur ('LIBCOMPL','',Multiniveau);
  for Indice:= 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
     TOBLN.addChampSupValeur ('MILLIEME'+IntToStr(indice+1),0);
  end;

end;

Procedure CalculeUneCategorieTaxeOuv ( VenteAchat,FamilleTaxe : String ; DEV : Rdevise; NumCat,NbDec : integer ; MontantBase : double ; TOBOuv,TOBTaxes,TOBPiece : TOB ; EnHT : boolean ) ;
Var Base,BaseDev,Taux,MontantTaxe,MontantTaxeDev : Double ;
    TOBA,TOBT : TOB ;
    RegimeTaxe,FormuleTaxe,CatTaxe : String ;
    TheMillieme,TauxDev : double;
BEGIN
  if FamilleTaxe='' then FamilleTaxe:=TOBOuv.GetValue('BLO_FAMILLETAXE'+IntToStr(NumCat)) ;
  RegimeTaxe:=TOBOuv.GetValue('BLO_REGIMETAXE') ; Taux:=0 ;
  TheMillieme := 0;
  if TOBOUV.FieldExists('MILLIEME'+IntToStr(NumCat)) then
  begin
    TheMillieme := valeur(TOBOUV.GetValue('MILLIEME'+IntToStr(NumCat)));
  end;
  TauxDev := TOBOUV.GetValue('BLO_TAUXDEV');
  CatTaxe := 'TX'+IntToStr(NumCat);
  TOBT:=VH^.LaTOBTVA.FindFirst(['TV_TVAOUTPF','TV_REGIME','TV_CODETAUX'],[CatTaxe,RegimeTaxe,FamilleTaxe],False) ;
  if TOBT<>Nil then
  BEGIN
    if VenteAchat='VEN' then Taux:=TOBT.GetValue('TV_TAUXVTE') else Taux:=TOBT.GetValue('TV_TAUXACH') ;
    FormuleTaxe:=TOBT.GetValue('TV_FORMULETAXE') ;
  END ;

  TOBA:=TOBTaxes.FindFirst(['BLB_CATEGORIETAXE','BLB_FAMILLETAXE'],[CatTaxe,FamilleTaxe],true);
  if TOBA = nil then
  begin
    TOBA:= TOB.Create('LIGNEBASE',TOBTaxes,-1);
    TOBA.putValue('BLB_CATEGORIETAXE',CatTaxe);
    TOBA.putValue('BLB_FAMILLETAXE',FamilleTaxe);
    TOBA.putValue('BLB_TAUXTAXE',Taux);
    TOBA.putValue('BLB_TAUXDEV',TOBOUV.GetValue('BLO_TAUXDEV'));
    TOBA.putValue('BLB_COTATION',TOBOUV.GetValue('BLO_COTATION'));
    TOBA.putValue('BLB_NUMORDRE',TOBOUV.GetValue('BLO_NUMORDRE'));
    TOBA.putValue('BLB_DEVISE',TOBOUV.GetValue('BLO_DEVISE'));
  end;
  BaseDev:=TOBA.GetValue('BLB_BASEDEV') ;
  if TheMillieme <> 0 then
  begin
    BaseDev:=BaseDev+Arrondi(MontantBase * (TheMillieme / 1000),NbDec);
  end else
  begin
    BaseDev:= BaseDev+MontantBase ;
  end;
  Base := DeviseToPivotEx(BaseDev,TauxDev,DEV.quotite,DEV.decimale);
  TOBA.PutValue('BLB_BASEDEV',BaseDev) ;
  TOBA.PutValue('BLB_BASETAXE',Base) ;
  if EnHT then
  begin
    MontantTaxeDev:=Arrondi(CalculeMontantTaxe(BaseDev,Taux,FormuleTaxe,TOBOUV),NbDec);
    MontantTaxe:=Arrondi(CalculeMontantTaxe(Base,Taux,FormuleTaxe,TOBOUV),NbDec);
  end else
  begin
    MontantTaxeDev:=Arrondi(CalculeMontantTaxeTTC(BaseDev,Taux,FormuleTaxe,TOBOUV),NbDec) ;
    MontantTaxe:=Arrondi(CalculeMontantTaxe(Base,Taux,FormuleTaxe,TOBOUV),NbDec);
  end;
END ;

Function CalculTaxesLigneOuv ( TOBOuv,TOBPiece : TOB ; DEV : RDevise; NbDec : integer ; Base : Double ; EnHT : Boolean ) : TOB ;
Var i : integer ;
    AvecTaxes : boolean ;
    TOBTaxes,TOBTaxe : TOB ;
    TotTPF,TotTVA,LaBase : Double ;
    FamilleTaxe,RegimeTaxe,NaturePiece,VenteAchat : String ;
BEGIN
Result:=Nil ;
// Informations PARPIECE
NaturePiece:=TOBOuv.GetValue('BLO_NATUREPIECEG') ;
VenteAchat:=GetInfoParPiece(NaturePiece,'GPP_VENTEACHAT') ;
AvecTaxes:=(GetInfoParPiece(NaturePiece,'GPP_TAXE')='X') ;
RegimeTaxe:=TOBOuv.GetValue('BLO_REGIMETAXE') ;
if ((Not AvecTaxes) or (VenteAchat='AUT')) then Exit ;
TOBTaxes:=TOB.Create('TAXES',Nil,-1) ;
TotTPF:=0 ; TotTVA:=0 ;
if EnHT then
   BEGIN
   {En HT, calculer les bases type TPF (sur HT) puis la TVA en dernier (sur HT+TPF)}
   for i:=5 downto 1 do
       BEGIN
       FamilleTaxe:=TOBOuv.GetValue('BLO_FAMILLETAXE'+IntToStr(i)) ;
       if i=1 then LaBase:=Base+TotTPF else LaBase:=Base ;
       if FamilleTaxe<>'' then
          BEGIN
          CalculeUneCategorieTaxeOuv(VenteAchat,FamilleTaxe,DEV, i,NbDec,LaBase,TOBOuv,TOBTaxes,TOBPiece,EnHT) ;
//          if i>1 then BEGIN TOBTaxe:=TOBTaxes.Detail[i-1] ; TotTPF:=TotTPF+TOBTaxe.GetValue('BLB_VALEURDEV') ; END ;
          END ;
       END ;
   END else
   BEGIN
   {En TTC, calculer la base type TVA (sur TTC) puis les TPF en dernier (sur TTC-TVA)}
   for i:=1 to 5 do
       BEGIN
       FamilleTaxe:=TOBOuv.GetValue('BLO_FAMILLETAXE'+IntToStr(i)) ;
       if i=1 then LaBase:=Base else LaBase:=Base-TotTVA ;
       if FamilleTaxe<>'' then
          BEGIN
          CalculeUneCategorieTaxeOuv(VenteAchat,FamilleTaxe,DEV,i,NbDec,LaBase,TOBOuv,TOBTaxes,TOBPiece,EnHT) ;
//          if i=1 then BEGIN TOBTaxe:=TOBTaxes.Detail[i-1] ; TotTVA:=TOBTaxe.GetValue('BLB_VALEURDEV') ; END ;
          END ;
       END ;
   END ;
Result:=TOBTaxes ;
END ;

Procedure CalculeLigneHTOuv ( TOBOuv,TOBPiece : TOB ; DEV:RDevise ; Forcer:boolean=False ) ;
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    TOBTaxes,TTLoc : TOB ;
    i        : integer ;
    Nbdec : Integer;
    PUHTDEV,PUHTNETDEV,MontantHTDev,MontantHTBrutDev,TotTaxes,LaTaxe,LaTaxeNR,TotTaxesNR : Double ;
    PUTTCDEV,(*PUTTCNETDEV,*)TaxeLoc : Double ;
    RemLigneDev,EscompteDev,TotalHTDev,TotalHTRemises,RemPiedDev : double ;
BEGIN
Nbdec := DEV.Decimale;
if not Forcer then
  begin
  if (TOBOuv.GetValue('BLO_TYPEREF') <> 'CAT') and (TOBOuv.GetValue('BLO_ARTICLE') = '') then Exit ;
  end;
QteF:=TOBOuv.GetValue('BLO_QTEFACT') ; // BRL : if QteF=0 then Exit ;
PUHTDEV:=TOBOuv.GetValue('BLO_PUHTDEV') ; // BRL : if PUHTDEV=0 then Exit ;
PQ:=TOBOuv.GetValue('BLO_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBOuv.PutValue('BLO_PRIXPOURQTE',PQ) ; END ;
RemLigne:=0 ; RemPied:=0 ; Escompte:=0 ; TOBTaxes:=Nil ;
// Informations issues de l'article
if TOBOuv.GetValue('BLO_REMISABLELIGNE')='X' then RemLigne:=TOBOuv.GetValue('BLO_REMISELIGNE')/100.0 ;
if TOBOuv.GetValue('BLO_REMISABLEPIED')='X' then RemPied:=TOBOuv.GetValue('BLO_REMISEPIED')/100.0 ;
if TOBOuv.GetValue('BLO_ESCOMPTABLE')='X' then Escompte:=TOBOuv.GetValue('BLO_ESCOMPTE')/100.0 ;
// Avec remise ligne
MontantHTBrutDev:=Arrondi(PUHTDEV*QteF/PQ,Nbdec) ;
MontantHTDev:=Arrondi(PUHTDEV*QteF*(1.0-RemLigne)/PQ,Nbdec) ;

RemLigneDev:=MontantHTBrutDev-MontantHTDev ;
MontantHTDev:=Arrondi(MontantHTDev,NbDec) ;
if QteF<>0 then PUHTNETDEV:=Arrondi(MontantHTDev/QteF/PQ,V_PGI.okdecP) else PUHTNETDEV:=Arrondi(PUHTDEV*(1.0-RemLigne),V_PGI.okdecP) ;
// Avec remise pied
TotalHTRemises:=Arrondi(PUHTDev*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,Nbdec) ;
RemPiedDev:=MontantHTDev-TotalHTRemises ;
// Avec escompte
TotalHTDev:=Arrondi(MontantHTDEV*(1.0-RemPied)*(1.0-Escompte),6) ;
TotalHTDev:=Arrondi(TotalHTDev,NbDec) ;
EscompteDev:=Arrondi(TotalHTRemises-TotalHTDev,Nbdec) ;
// Taxes
TOBTaxes:=CalculTaxesLigneOuv(TOBOuv,TOBPiece,DEV,NbDec,ToTalHtDev,True) ;
// Autres cumuls lignes
TotTaxes:=0 ; TotTaxesNR:=0 ;
if TOBTaxes<>Nil then
   BEGIN
   for i:=0 to TOBTaxes.detail.count -1 do
       BEGIN
       LaTaxe:=TOBTaxes.Detail[i].GetValue('BLB_VALEURDEV') ;
       if ((RemPied=1) or (Escompte=1)) then LaTaxeNR:=0 else LaTaxeNR:=LaTaxe/(1.0-RemPied)/(1.0-Escompte) ;
       LaTaxe:=Arrondi(LaTaxe,6) ; TotTaxes:=TotTaxes+LaTaxe ;
       LaTaxeNR:=Arrondi(LaTaxeNR,NbDec) ; TotTaxesNR:=TotTaxesNR+LaTaxeNR ;
       END ;
   END ;
if Arrondi(QteF*(1.0-RemLigne),6)<>0 then
   BEGIN
   PUTTCDEV:=Arrondi(MontantHTDev+TotTaxesNR,6)*PQ/(QteF*(1.0-RemLigne)) ;
   TOBOuv.PutValue('BLO_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.okdecP)) ;
   END else
   BEGIN
   TTLoc:=CalculTaxesLigneOuv(TOBOuv,TOBPiece,DEV,NbDec,TOBOuv.GetValue('BLO_PUHTDEV'),True) ;
   TaxeLoc:=0 ;
   if TTLoc<>Nil then for i:=0 to TOBTaxes.detail.count-1 do TaxeLoc:=TaxeLoc+TTLoc.Detail[i].GetValue('BLB_VALEURDEV') ;
   PUTTCDEV:=TOBOuv.GetValue('BLO_PUHTDEV')+TaxeLoc ; (*PUTTCNETDEV:=PUTTCDEV*(1.0-RemLigne) ;*)
   TOBOuv.PutValue('BLO_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.okdecP)) ;
   END ;
TOBOuv.PutValue('BLO_PUTTC',DeviseToPivotEx(TOBOUV.GetValue('BLO_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
CalculMontantHtDevLigOuv (TOBOUV,DEV);
TOBTaxes.Free ; TOBTaxes:=Nil ;
END ;

Procedure CalculeLigneTTCOuv ( TOBOuv,TOBPiece : TOB ; DEV:RDevise; Forcer:boolean=False) ;
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    TOBTaxes,TTLoc : TOB ;
    i        : integer ;
    Nbdec : Integer;
    PUHTDEV,PUHTNETDEV,MontantTTCDev,MontantTTCBrutDev,TotTaxes,LaTaxe,LaTaxeNR,TotTaxesNR : Double ;
    PUTTCDEV,PUTTCNETDEV,TaxeLoc,TotalTTCDev : Double ;
    RemLigneDevTTC,EscompteTTCDev,TotalTTCRemises,RemPiedTTCDev : double ;
BEGIN
Nbdec := DEV.Decimale ;
if not Forcer then
  begin
  if (TOBOuv.GetValue('BLO_TYPEREF') <> 'CAT') and (TOBOuv.GetValue('BLO_ARTICLE') = '') then Exit ;
  end;
QteF:=TOBOuv.GetValue('BLO_QTEFACT') ; if QteF=0 then Exit ;
PUTTCDEV:=TOBOuv.GetValue('BLO_PUTTCDEV') ; if PUTTCDEV=0 then Exit ;
PQ:=TOBOuv.GetValue('BLO_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBOuv.PutValue('BLO_PRIXPOURQTE',PQ) ; END ;
RemLigne:=0 ; RemPied:=0 ; Escompte:=0 ; TOBTaxes:=Nil ;
// Informations issues de l'article
if TOBOuv.GetValue('BLO_REMISABLELIGNE')='X' then RemLigne:=TOBOuv.GetValue('BLO_REMISELIGNE')/100.0 ;
if TOBOuv.GetValue('BLO_REMISABLEPIED')='X' then RemPied:=TOBOuv.GetValue('BLO_REMISEPIED')/100.0 ;
if TOBOuv.GetValue('BLO_ESCOMPTABLE')='X' then Escompte:=TOBOuv.GetValue('BLO_ESCOMPTE')/100.0 ;
// Avec remise ligne
MontantTTCBrutDev:=Arrondi(PUTTCDEV*QteF/PQ,6) ;
MontantTTCDev:=Arrondi(PUTTCDEV*QteF*(1.0-RemLigne)/PQ,6) ;
RemLigneDevTTC:=MontantTTCBrutDev-MontantTTCDev ;
if TOBPiece.GetValue('GP_ARRONDILIGNE')='X' then MontantTTCDev:=Arrondi(MontantTTCDev,NbDec) ;
if (ctxMode in V_PGI.PGIContexte) then
begin
     MontantTTCDEV:=ArrondirPrix(TOBOuv.GetValue('BLO_CODEARRONDI'),MontantTTCDEV) ;
     PUTTCNETDEV:=ArrondirPrix(TOBOuv.GetValue('BLO_CODEARRONDI'),PUTTCNETDEV) ;
end ;
if QteF<>0 then PUTTCNETDEV:=Arrondi(MontantTTCDev/QteF/PQ,6) else PUTTCNETDEV:=Arrondi(PUTTCDEV*(1.0-RemLigne),6) ;
// Avec remise pied
TotalTTCRemises:=Arrondi(PUTTCDev*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,6) ;
RemPiedTTCDev:=MontantTTCDev-TotalTTCRemises ;
// Avec escompte
TotalTTCDev:=Arrondi(MontantTTCDEV*(1.0-RemPied)*(1.0-Escompte),6) ;
if TOBPiece.GetValue('GP_ARRONDILIGNE')='X' then TotalTTCDev:=Arrondi(TotalTTCDev,NbDec) ;
EscompteTTCDev:=Arrondi(TotalTTCRemises-TotalTTCDev,6) ;
// Calcul des bases TTC
TOBTaxes:=CalculTaxesLigneOuv(TOBOuv,TOBPiece,DEV,NbDec,TotalTTCDev,False) ;
// Taxes
TotTaxes:=0 ; TotTaxesNR:=0 ;
if TOBTaxes<>Nil then
   BEGIN
   for i:=0 to TOBTaxes.detail.count-1 do
       BEGIN
       LaTaxe:=TOBTaxes.Detail[i].GetValue('BLB_VALEURDEV') ;
       if ((RemPied=1) or (Escompte=1)) then LaTaxeNR:=0 else LaTaxeNR:=LaTaxe/(1.0-RemPied)/(1.0-Escompte) ;
       LaTaxe:=Arrondi(LaTaxe,6) ; TotTaxes:=TotTaxes+LaTaxe ;
       LaTaxeNR:=Arrondi(LaTaxeNR,NbDec) ; TotTaxesNR:=TotTaxesNR+LaTaxeNR ;
       END ;
   END ;
TOBTaxes.Free ; TOBTaxes:=Nil ;
if Arrondi(QteF*(1.0-RemLigne),6)<>0 then
   BEGIN
   PUHTDEV:=Arrondi(MontantTTCDev-TotTaxesNR,6)*PQ/(QteF*(1.0-RemLigne)) ;
   TOBOuv.PutValue('BLO_PUHTDEV',Arrondi(PUHTDEV,V_PGI.OkdecP)) ;
   PUHTNETDEV:=Arrondi(MontantTTCDev-TotTaxesNR,6)*PQ/QteF ;
   END else
   BEGIN
   TTLoc:=CalculTaxesLigneOuv(TOBOuv,TOBPiece,DEV,NbDec,TOBOuv.GetValue('BLO_PUTTCDEV'),False) ;
   TaxeLoc:=0 ;
   if TTLoc<>Nil then for i:=0 to TTLoc.detail.count -1 do TaxeLoc:=TaxeLoc+TTLoc.Detail[i].GetValue('BLB_VALEURDEV') ;
   PUHTDEV:=TOBOuv.GetValue('BLO_PUTTCDEV')-TaxeLoc ; PUHTNETDEV:=PUHTDEV*(1.0-RemLigne) ;
   TOBOuv.PutValue('BLO_PUHTDEV',Arrondi(PUHTDEV,V_PGI.okdecP)) ;
   END ;
TOBOuv.PutValue('BLO_PUHT',DeviseToPivotEx(TOBOUV.GetValue('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
CalculMontantHtDevLigOuv (TOBOUV,DEV);
END ;

//////////////////////////////////////////
// FOnctions de calcul sur ligne document
//////////////////////////////////////////
Procedure CalculeLigneHTOuvMere ( TOBL,TOBBases,TOBPiece : TOB ; DEV : Rdevise; NbDec : integer;arrondiok:boolean;var MontantLignes:double) ;
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    TOBTaxes,TTLoc : TOB ;
    i        : integer ;
    PUHTDEV,PUHTNETDEV,MontantHTDev,MontantHTBrutDev,TotTaxes,LaTaxe,LaTaxeNR,TotTaxesNR : Double ;
    PUTTCDEV,PUTTCNETDEV,TaxeLoc : Double ;
    RemLigneDev,EscompteDev,TotalHTDev,TotalHTRemises,RemPiedDev : double ;
BEGIN
TOBTaxes := TOB.create('UNE LIGNE',nil,-1);
ZeroLigne(TOBL) ;
QteF:=TOBL.GetValue('GL_QTEFACT') ; if QteF=0 then Exit ;
PUHTDEV:=TOBL.GetValue('GL_PUHTDEV') ; if PUHTDEV=0 then Exit ;
PQ:=TOBL.GetValue('GL_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue('GL_PRIXPOURQTE',PQ) ; END ;
RemLigne:=0 ; RemPied:=0 ; Escompte:=0 ; TOBTaxes:=Nil ;

// Calcul montant achat ligne
{$IFNDEF V43}
//TOBL.PutValue('GL_MONTANTPA',Arrondi(TOBL.GetValue('GL_DPA')*QteF/PQ,NbDec)) ;
//TOBL.PutValue('GL_MONTANTPR',Arrondi(TOBL.GetValue('GL_DPR')*QteF/PQ,NbDec)) ;
{$ENDIF}

// Informations issues de l'article
if TOBL.GetValue('GL_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue('GL_REMISELIGNE')/100.0 ;
if TOBL.GetValue('GL_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue('GL_REMISEPIED')/100.0 ;
if TOBL.GetValue('GL_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue('GL_ESCOMPTE')/100.0 ;
// Avec remise ligne
   MontantHTBrutDev:=Arrondi(PUHTDEV*QteF/PQ,6) ;

if TOBL.GetValue('GL_VALEURREMDEV')<>0 then
   BEGIN
   RemLigneDEV:=TOBL.GetValue('GL_VALEURREMDEV') ;
   MontantHTDev:=MontantHTBrutDev-RemLigneDev ;
   END else
   BEGIN
   MontantHTDev:=Arrondi(PUHTDEV*QteF*(1.0-RemLigne)/PQ,6) ;
   RemLigneDev:=MontantHTBrutDev-MontantHTDev ;
   END ;
if arrondiOk then MontantHTDev:=Arrondi(MontantHTDev,NbDec) ;
TOBL.PutValue('GL_MONTANTHTDEV',MontantHTDEV) ;
MontantLignes:=MontantLignes+MontantHTDEV ;
TOBL.PutValue('GL_TOTREMLIGNEDEV',RemLigneDev) ;
if QteF<>0 then PUHTNETDEV:=Arrondi(MontantHTDev/QteF/PQ,6)
else PUHTNETDEV:=Arrondi(PUHTDEV*(1.0-RemLigne),6) ;
TOBL.PutValue('GL_PUHTNETDEV',PUHTNETDEV) ;
// Avec remise pied
TotalHTRemises:=Arrondi(PUHTDev*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,6) ;
RemPiedDev:=MontantHTDev-TotalHTRemises ;
TOBL.PutValue('GL_TOTREMPIEDDEV',RemPiedDev) ;
// Avec escompte
//TotalHTDev:=Arrondi(PUHTDev*QteF*(1.0-RemLigne)*(1.0-RemPied)*(1.0-Escompte)/PQ,6) ;
TotalHTDev:=Arrondi(MontantHTDEV*(1.0-RemPied)*(1.0-Escompte),6) ;
if TOBPiece.GetValue('GP_ARRONDILIGNE')='X' then TotalHTDev:=Arrondi(TotalHTDev,NbDec) ;
EscompteDev:=Arrondi(TotalHTRemises-TotalHTDev,6) ;
TOBL.PutValue('GL_TOTESCLIGNEDEV',EscompteDev) ;
TOBL.PutValue('GL_TOTALHTDEV',TotalHTDev) ;
// Taxes
CalculTaxesLigne(TOBL,TOBPiece,Nil,TOBtaxes,DEV,NbDec,TOBL.GetValue('GL_TOTALHTDEV'),True) ;
//if TOBTaxes<>Nil then CumuleLesBases(TOBBases,TOBTaxes,NbDec) ;
// Autres cumuls lignes
TotTaxes:=0 ; TotTaxesNR:=0 ;
if TOBTaxes<>Nil then
   BEGIN
   for i:=0 to tobTaxes.detail.count -1 do
       BEGIN
       LaTaxe:=TOBTaxes.Detail[i].GetValue('BLB_VALEURDEV') ;
       if ((RemPied=1) or (Escompte=1)) then LaTaxeNR:=0 else LaTaxeNR:=LaTaxe/(1.0-RemPied)/(1.0-Escompte) ;
       LaTaxe:=Arrondi(LaTaxe,6) ; TotTaxes:=TotTaxes+LaTaxe ;
       LaTaxeNR:=Arrondi(LaTaxeNR,NbDec) ; TotTaxesNR:=TotTaxesNR+LaTaxeNR ;
//     TOBL.PutValue('GL_TOTALTAXEDEV'+IntToStr(i),LaTaxe) ;
       END ;
   END ;
TOBTaxes.Free ; TOBTaxes:=Nil ;
TOBL.PutValue('GL_TOTALTTCDEV',Arrondi(TotalHTDev+TotTaxes,6)) ;
TOBL.PutValue('GL_MONTANTTTCDEV',Arrondi(MontantHTDev+TotTaxesNR,6)) ;
if Arrondi(QteF*(1.0-RemLigne),6)<>0 then
   BEGIN
   PUTTCDEV:=TOBL.GetValue('GL_MONTANTTTCDEV')*PQ/(QteF*(1.0-RemLigne)) ;
   TOBL.PutValue('GL_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.okdecP)) ;
   PUTTCNETDEV:=TOBL.GetValue('GL_MONTANTTTCDEV')*PQ/QteF ;
   TOBL.PutValue('GL_PUTTCNETDEV',Arrondi(PUTTCNETDEV,V_PGI.OkdecP)) ;
   END else
   BEGIN
   TTLoc := TOB.create ('UNE LIGNE',nil,-1);
   CalculTaxesLigne(TOBL,TOBPiece,Nil,TTLoc,DEV, NbDec,TOBL.GetValue('GL_PUHTDEV'),True) ;
   TaxeLoc:=0 ;
   if TTLoc<>Nil then for i:=1 to 5 do TaxeLoc:=TaxeLoc+TTLoc.Detail[i-1].GetValue('BLB_VALEURDEV') ;
   PUTTCDEV:=TOBL.GetValue('GL_PUHTDEV')+TaxeLoc ; PUTTCNETDEV:=PUTTCDEV*(1.0-RemLigne) ;
   TOBL.PutValue('GL_PUTTCDEV',Arrondi(PUTTCDEV,V_PGI.OkdecP)) ;
   TOBL.PutValue('GL_PUTTCNETDEV',Arrondi(PUTTCNETDEV,V_PGI.OkdecP)) ;
   TTLoc.free;
   END ;
TOBTaxes.free;
END ;

Procedure CalculeLigneTTCOuvMere ( TOBL,TOBBases,TOBPiece : TOB ; DEV: Rdevise; NbDec : integer;ArrondiOk : boolean;var MontantLignes:double);
Var QteF,PQ,RemLigne,RemPied,Escompte : Double ;
    TOBTaxes,TTLoc : TOB ;
    i        : integer ;
    PUHTDEV,PUHTNETDEV,MontantTTCDev,MontantTTCBrutDev,TotTaxes,LaTaxe,LaTaxeNR,TotTaxesNR : Double ;
    PUTTCDEV,PUTTCNETDEV,TaxeLoc,TotalTTCDev : Double ;
    RemLigneDevTTC,EscompteTTCDev,TotalTTCRemises,RemPiedTTCDev : double ;
BEGIN
TOBTaxes := TOB.create ('UNE LIGNE',nil,-1);
ZeroLigne(TOBL) ; InitLesSupMontants(TOBL) ;
QteF:=TOBL.GetValue('GL_QTEFACT') ; if QteF=0 then Exit ;
PUTTCDEV:=TOBL.GetValue('GL_PUTTCDEV') ; if PUTTCDEV=0 then Exit ;
PQ:=TOBL.GetValue('GL_PRIXPOURQTE') ; if PQ<=0 then BEGIN PQ:=1.0 ; TOBL.PutValue('GL_PRIXPOURQTE',PQ) ; END ;
RemLigne:=0 ; RemPied:=0 ; Escompte:=0 ; TOBTaxes:=Nil ;

// Calcul montant achat ligne
{$IFNDEF V43}
//TOBL.PutValue('GL_MONTANTPA',Arrondi(TOBL.GetValue('GL_DPA')*QteF/PQ,NbDec)) ;
//TOBL.PutValue('GL_MONTANTPR',Arrondi(TOBL.GetValue('GL_DPR')*QteF/PQ,NbDec)) ;
{$ENDIF}

// Informations issues de l'article
if TOBL.GetValue('GL_REMISABLELIGNE')='X' then RemLigne:=TOBL.GetValue('GL_REMISELIGNE')/100.0 ;
if TOBL.GetValue('GL_REMISABLEPIED')='X' then RemPied:=TOBL.GetValue('GL_REMISEPIED')/100.0 ;
if TOBL.GetValue('GL_ESCOMPTABLE')='X' then Escompte:=TOBL.GetValue('GL_ESCOMPTE')/100.0 ;
// Avec remise ligne
MontantTTCBrutDev:=Arrondi(PUTTCDEV*QteF/PQ,6) ;
if TOBL.GetValue('GL_VALEURREMDEV')<>0 then
   BEGIN
   RemLigneDevTTC:=TOBL.GetValue('GL_VALEURREMDEV') ;
   MontantTTCDev:=MontantTTCBrutDev-RemLigneDevTTC ;
   END else
   BEGIN
   MontantTTCDev:=Arrondi(PUTTCDEV*QteF*(1.0-RemLigne)/PQ,6) ;
   RemLigneDevTTC:=MontantTTCBrutDev-MontantTTCDev ;
   END ;
if ArrondiOk then MontantTTCDev:=Arrondi(MontantTTCDev,NbDec) ;
TOBL.PutValue('GL_MONTANTTTCDEV',MontantTTCDEV) ;
MontantLignes:=MontantLignes+MontantTTCDEV ;
TOBL.PutValue('TOTREMLIGNETTCDEV',RemLigneDevTTC) ;
if QteF<>0 then PUTTCNETDEV:=Arrondi(MontantTTCDev/QteF/PQ,6)
else PUTTCNETDEV:=Arrondi(PUTTCDEV*(1.0-RemLigne),6) ;
TOBL.PutValue('GL_PUTTCNETDEV',PUTTCNETDEV) ;
// Avec remise pied
TotalTTCRemises:=Arrondi(PUTTCDev*QteF*(1.0-RemLigne)*(1.0-RemPied)/PQ,6) ;
RemPiedTTCDev:=MontantTTCDev-TotalTTCRemises ;
TOBL.PutValue('TOTREMPIEDTTCDEV',RemPiedTTCDev) ;
// Avec escompte
TotalTTCDev:=Arrondi(MontantTTCDEV*(1.0-RemPied)*(1.0-Escompte),6) ;
if TOBPiece.GetValue('GP_ARRONDILIGNE')='X' then TotalTTCDev:=Arrondi(TotalTTCDev,NbDec) ;
EscompteTTCDev:=Arrondi(TotalTTCRemises-TotalTTCDev,6) ;
TOBL.PutValue('TOTESCLIGNETTCDEV',EscompteTTCDev) ;
TOBL.PutValue('GL_TOTALTTCDEV',TotalTTCDev) ;
// Calcul des bases TTC
CalculTaxesLigne(TOBL,TOBPiece,Nil,TOBTaxes,DEV,NbDec,TOBL.GetValue('GL_TOTALTTCDEV'),False) ;
//if TOBTaxes<>Nil then CumuleLesBases(TOBBases,TOBTaxes,NbDec) ;
// Taxes
TotTaxes:=0 ; TotTaxesNR:=0 ;
if TOBTaxes<>Nil then
   BEGIN
   for i:=1 to 5 do
       BEGIN
       LaTaxe:=TOBTaxes.Detail[i-1].GetValue('BLB_VALEURDEV') ;
       if ((RemPied=1) or (Escompte=1)) then LaTaxeNR:=0 else LaTaxeNR:=LaTaxe/(1.0-RemPied)/(1.0-Escompte) ;
       LaTaxe:=Arrondi(LaTaxe,6) ; TotTaxes:=TotTaxes+LaTaxe ;
       LaTaxeNR:=Arrondi(LaTaxeNR,NbDec) ; TotTaxesNR:=TotTaxesNR+LaTaxeNR ;
       TOBL.PutValue('GL_TOTALTAXEDEV'+IntToStr(i),LaTaxe) ;
       END ;
   END ;
TOBTaxes.Free ; TOBTaxes:=Nil ;

TOBL.PutValue('GL_TOTALHTDEV',Arrondi(TotalTTCDev-TotTaxes,6)) ;
TOBL.PutValue('GL_MONTANTHTDEV',Arrondi(MontantTTCDev-TotTaxesNR,6)) ;
if Arrondi(QteF*(1.0-RemLigne),6)<>0 then
   BEGIN
   PUHTDEV:=TOBL.GetValue('GL_MONTANTHTDEV')*PQ/(QteF*(1.0-RemLigne)) ;
   TOBL.PutValue('GL_PUHTDEV',Arrondi(PUHTDEV,V_PGI.OkdecP)) ;
   PUHTNETDEV:=TOBL.GetValue('GL_MONTANTHTDEV')*PQ/QteF ;
   TOBL.PutValue('GL_PUHTNETDEV',Arrondi(PUHTNETDEV,V_PGI.OkdecP)) ;
   END else
   BEGIN
   TTLOC := TOB.Create('UNE LIGNE',nil,-1);
   CalculTaxesLigne(TOBL,TOBPiece,Nil,TTLOc,DEV,NbDec,TOBL.GetValue('GL_PUTTCDEV'),False) ;
   TaxeLoc:=0 ;
   if TTLoc<>Nil then for i:=1 to 5 do TaxeLoc:=TaxeLoc+TTLoc.Detail[i-1].GetValue('BLB_VALEURDEV') ;
   PUHTDEV:=TOBL.GetValue('GL_PUTTCDEV')-TaxeLoc ; PUHTNETDEV:=PUHTDEV*(1.0-RemLigne) ;
   TOBL.PutValue('GL_PUHTDEV',Arrondi(PUHTDEV,V_PGI.OkdecP)) ;
   TOBL.PutValue('GL_PUHTNETDEV',Arrondi(PUHTNETDEV,V_PGI.OkdecP)) ;
   TTLOc.free;
   END ;
TOBTaxes.free;
END ;

//////////////////////////////////////////

Procedure AffecteValoOuv ( TOBN,TOBA : TOB ; Depot : String;DEV:RDevise ) ;
Var TOBD : TOB ;
    OkDepot : boolean ;
BEGIN
  if TOBN=Nil then Exit ;
  if TOBA=Nil then Exit ;
  OkDepot:=False ; TOBD:=Nil ;
  if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
  begin
    if TOBA.GetValue('GA_PMAP') <> 0 then
    begin
      TOBA.PutValue('GA_PAHT', TOBA.GetValue('GA_PMAP'))
    end;
  end;
// --- OMFG -----
	ReactualisePrPv (TOBA);
// -------------------

  if VH_GC.GCValoArtDepot then
  BEGIN
    TOBD:=TOBA.FindFirst(['GQ_DEPOT'],[Depot],False) ;
    if TOBD<>Nil then OkDepot:=True ;
  END ;
  if OkDepot then
  BEGIN
    TOBN.PutValue('DPA',valeur(strs(TOBD.GetValue('GQ_DPA'),V_PGI.OkDecP))) ;
    TOBN.PutValue('DPR',valeur(strs(TOBD.GetValue('GQ_DPR'),V_PGI.OkDecP))) ;
    TOBN.PutValue('PMAP',valeur(strs(TOBD.GetValue('GQ_PMAP'),V_PGI.OkDecP))) ;
    TOBN.PutValue('PMRP',valeur(strs(TOBD.GetValue('GQ_PMRP'),V_PGI.OkDecP))) ;
  END else
  BEGIN
    //   TOBN.PutValue('DPA',valeur(strs(TOBA.GetValue('GA_PAHT'),V_PGI.OkDecP))) ;
    TOBN.PutValue('DPA', TOBA.GetValue('GA_PAHT'));
    TOBN.PutValue('DPR',valeur(strs(TOBA.GetValue('GA_DPR'),V_PGI.OkDecP))) ;
    TOBN.PutValue('PMAP',valeur(strs(TOBA.GetValue('GA_PMAP'),V_PGI.OkDecP))) ;
    TOBN.PutValue('PMRP',valeur(strs(TOBA.GetValue('GA_PMRP'),V_PGI.OkDecP))) ;
  END ;

  TOBN.PutValue('PVHT',pivottodevise(TOBA.GetValue('GA_PVHT'),DEV.Taux,DEV.quotite,V_PGI.OkdecP)) ;
  TOBN.PutValue('PVTTC',pivottodevise(TOBA.GetValue('GA_PVTTC'),DEV.Taux ,DEV.quotite,V_PGI.OkdecP) ) ;
END ;

Procedure DerouleOuvrage ( TOBN : TOB ; CodeNomen,RefUnique,Depot : String ; TOBArticles : TOB ; QteDuDetail : Double;DEV:RDevise;VenteAchat : string) ;
Var Q,QQ : TQuery ;
    TOBL,TOBArt : TOB ;
    i    : integer ;
    CodeN,RefArticle  : String ;
    StSQL : String;
BEGIN

// Nomenclatures
Q:=OpenSQL('SELECT NOMENLIG.*,GNE_QTEDUDETAIL,GNE_DOMAINE FROM NOMENLIG LEFT OUTER JOIN NOMENENT ON GNE_NOMENCLATURE=GNL_NOMENCLATURE WHERE GNL_NOMENCLATURE="'+CodeNomen+'" ORDER BY GNL_NUMLIGNE',True,-1, '', True) ;
TOBN.LoadDetailDB('NOMENLIG','','',Q,False,False) ;
Ferme(Q) ;
{// Entete
Q:=OpenSQL('SELECT GNE_QTEDUDETAIL FROM NOMENENT WHERE GNE_NOMENCLATURE="'+CodeNomen+'"',True) ;
if Q.Fields[0].AsFloat <> 0 then QteDuDetail := Q.Fields[0].AsFloat ;
Ferme(Q) ;}
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    TOBL.AddChampSup('COMPOSE',False) ; TOBL.AddChampSup('TENUESTOCK',False) ;
    TOBL.AddChampSup('QUALIFQTEACH',False) ; TOBL.AddChampSup('QUALIFQTEVTE',False) ; TOBL.AddChampSup('QUALIFQTESTO',False) ;
    TOBL.AddChampSup('DPA',False) ; TOBL.AddChampSup('DPR',False) ; TOBL.AddChampSup('PMAP',False) ; TOBL.AddChampSup('PMRP',False) ;
    TOBL.AddChampSup('PVTTC',False) ; TOBL.AddChampSup('PVHT',False) ;
    TOBL.AddChampSup ('QTEDUDETAIL',false);
    TOBL.addChampSUp ('PRIXPOURQTE',false);
    TOBL.addChampSUp ('TYPEARTICLE',false);
    TOBL.AddChampSup ('LIBCOMPL',false);
    TOBL.AddChampSupValeur('LIBELLEFOU','');
    (* OPTIMIZATION *)
		TOBN.addchampsupValeur('NOMENCLATURE',CodeNomen,false);
    TOBL.AddChampSupValeur ('PRXACHBASE',0,false);
    TOBL.AddChampSupValeur ('REMISES','',false);
    (* ---------- *)
    TOBL.Putvalue('DPA',valeur('0'));TOBL.Putvalue('DPR',valeur('0'));
    TOBL.Putvalue('PMAP',valeur('0'));TOBL.Putvalue('PMRP',valeur('0'));
    TOBL.Putvalue('PVHT',valeur('0'));TOBL.Putvalue('PVTTC',valeur('0'));
    TOBL.PutValue('COMPOSE',RefUnique) ;
    TOBL.PutValue('QTEDUDETAIL',TOBL.Getvalue('GNE_QTEDUDETAIL'));
    TOBL.PutValue('QUALIFQTEACH','');
    // fv 02032004
    //if TOBL.GetValue('QTEDUDETAIL') = 0 then TOBL.PutValue('QTEDUDETAIL',1);
    QteDuDetail := TOBL.GetValue('QTEDUDETAIL');
    // Articles
    RefArticle:=TOBL.GetValue('GNL_ARTICLE') ; TOBArt:=TOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
    if TOBArt=Nil then
       BEGIN
       TOBArt:=CreerTOBArt(TOBArticles) ;
//       TOBArt.SelectDB('"'+RefArticle+'"',Nil) ;
       StSQL := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
       							  '(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=A.GA_FOURNPRINC) AS LIBELLEFOU FROM ARTICLE A '+
    	  					    'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							    'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="' + RefArticle + '"';
       QQ := OpenSql (StSQl,true,-1, '', True);
  		 if not QQ.eof then
       begin
        TOBArt.SelectDB('',QQ);
       end else
       begin
        InitArticleFromLigneNomen (TOBL,TOBART); // grmfff
       end;
       ferme (QQ);
       InitChampsSupArticle (TOBART);
       LoadTOBDispo(TOBArt,False) ;
       END ;
    TOBL.Putvalue ('PRIXPOURQTE',TOBArt.GetValue ('GA_PRIXPOURQTE'));
    TOBL.PutValue('TENUESTOCK',TOBArt.GetValue('GA_TENUESTOCK')) ;
    TOBL.PutValue('QUALIFQTESTO',TOBArt.GetValue('GA_QUALIFUNITESTO')) ;
//    TOBL.PutValue('QUALIFQTEACH',TOBArt.GetValue('GA_QUALIFUNITEACH')) ;
    TOBL.PutValue('QUALIFQTEVTE',TOBArt.GetValue('GA_QUALIFUNITEVTE')) ;
    TOBL.PutValue('TYPEARTICLE',TOBArt.GetValue('GA_TYPEARTICLE')) ;
    TOBL.PutValue('LIBCOMPL',TOBArt.GetValue('GA_LIBCOMPL')) ;
    if TOBL.FieldExists('LIBELLEFOU') then TOBL.PutValue('LIBELLEFOU',TOBArt.GetValue('LIBELLEFOU')) ;
//
    InitValoArtNomen (TOBArt,VenteAchat);
//
    TOBL.PutValue('QUALIFQTEACH',TOBArt.GetValue('QUALIFQTEACH')) ;
    TOBL.PutValue ('PRXACHBASE',TOBArt.GetValue('PRXACHBASE'));
    TOBL.putValue ('REMISES',TOBArt.GetValue('REMISES'));
//
    AffecteValoOuv(TOBL,TOBArt,Depot,DEV) ;
    // -- Recup des coefs en fonction du domaine d'activite
    if (TOBL.GetValue('GNE_DOMAINE')<>'') and (TOBART.getValue('GA_PRIXPASMODIF')<>'X') then
    begin
      AppliqueValoDomaineAct (TOBL,TOBART);
    end;
    //
    // Nomenclatures
    CodeN:=TOBL.GetValue('GNL_SOUSNOMEN') ;
    if CodeN<>'' then DerouleOuvrage(TOBL,CodeN,TOBL.GetValue('GNL_ARTICLE'),Depot,TOBArticles,QteDuDetail,DEV,VenteAchat) ;
    END ;
END ;

Function ChoixOuvrage ( RefUnique,Depot: string ;var Libelle : String ; TOBArticles : TOB ; LePremier : boolean;DEV:RDevise;VenteAchat : string ) : TOB ;
Var TOBN : TOB ;
    Q    : TQuery ;
    TT   : TStrings ;
    CodeNomen,St : String ;
BEGIN
Result:=Nil ; TOBN:=Nil ; TT:=Nil ; CodeNomen:='' ;
Q:=OpenSQL('SELECT GNE_NOMENCLATURE, GNE_LIBELLE FROM NOMENENT WHERE GNE_ARTICLE="'+RefUnique+'"',True,-1, '', True) ;
if Not Q.EOF then
   BEGIN
   TT:=TStringList.Create ;
   While Not Q.EOF do
      BEGIN
      TT.Add(Q.Fields[0].AsString+';'+Q.Fields[1].AsString) ;
      if LePremier then Break else Q.Next ;
      END ;
   if TT.Count>1 then
      BEGIN
      CodeNomen:=Choisir('Choix d''un ouvrage','NOMENENT','GNE_LIBELLE','GNE_NOMENCLATURE','GNE_ARTICLE="'+RefUnique+'"','') ;
      END else
      BEGIN
      St:=TT[0] ; CodeNomen:=ReadTokenSt(St) ;
      END ;
   TT.Clear ; TT.Free ;
   END ;
Ferme(Q) ;
if CodeNomen='' then Exit ;
Q:=OpenSQL('SELECT GNE_LIBELLE FROM NOMENENT WHERE GNE_ARTICLE="'+RefUnique+'" AND GNE_NOMENCLATURE="'+CodeNomen+'"',True,-1, '', True) ;
if not q.eof then libelle:=Q.Fields[0].AsString;
ferme(Q);
TOBN:=TOB.Create('',Nil,-1) ;
// METRES
{$IFDEF BTP}
TOBN.addchampsupValeur('NOMENCLATURE',CodeNomen,false);
{$ENDIF}
DerouleOuvrage(TOBN,CodeNomen,RefUnique,Depot,TOBArticles,1,DEV,VenteAchat) ;
Result:=TOBN ;
END ;

Procedure RenseigneValoOuvrage ( TOBLN : TOB ;CodeDT: String; Qte,QteDuDetail: Double;DEV:RDEVISE;var Valeurs:T_Valeurs) ;
Var TOBPD : TOB ;
    i,IndPou      : integer ;
    TypeArtOk : string;
    QTeLoc,QteDuDet,PrixPourQte,pxht,Pxttc : Double;
    ValLoc,ValPou : T_Valeurs;
BEGIN
InitTableau (ValPou);
TypeArtOk := '';
IndPou := 0;

// fv 02032004
//if QteDudetail = 0 then QteDudetail := 1;

TOBPD:=TOBLN.findfirst(['TYPEARTICLE'],['POU'],false);
if TOBPD<> nil  then TypeArtOk := TOBPD.getValue('LIBCOMPL');

for i:=0 to TOBLN.Detail.Count-1 do
    BEGIN
    TOBPD:=TOBLN.Detail[i] ;
    if TOBPD.GetValue('TYPEARTICLE') = 'POU' then
       begin
       IndPou := i;
       continue;
       end;
    PrixPourQte := TOBPD.GetValue('PRIXPOURQTE');
    if PrixPourQte <= 0 then PrixPourQte := 1;
    QteLoc := Qte * TOBPD.GetValue('GNL_QTE');
    if TOBPD.detail.count > 0 then
       BEGIN
       InitTableau (ValLoc);
       QteDuDet:= QteDuDetail * TOBPD.GetValue('QTEDUDETAIL');
       RenseigneValoOuvrage (TOBPD,CodeDt,QTe,QteDuDet,DEV,ValLoc);
       TOBPD.Putvalue ('DPA',ValLoc[0]);
       TOBPD.Putvalue ('DPR',ValLoc[1]);
       TOBPD.Putvalue ('PVHT',ValLoc[2]);
       TOBPD.Putvalue ('PVTTC',ValLoc[3]);
       END;
    Valeurs [0] := Valeurs [0] + arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('DPA')),V_PGI.okdecP);
    Valeurs [1] := Valeurs [1] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('DPR')),V_PGI.OkdecP);
    if TOBPD.GetValue ('GNL_PRIXFIXE') <> 0 then
       pxHt :=(QteLoc/(QteDudetail*PrixPourQte)) * valeur(strs(TOBPD.GetValue ('GNL_PRIXFIXE'),V_PGI.OkDecP))
    else
       PxHt := (QteLoc/(QteDudetail*PrixPourQte)) * TOBPD.GetValue ('PVHT');
    pxTtc :=(QteLoc/(QteDudetail*PrixPourQte)) * TOBPD.GetValue ('PVTTC');
    Valeurs [2] := Valeurs [2] + Arrondi(PxHt,V_PGI.okdecP);
    Valeurs [3] := Valeurs [3] + Arrondi(PxTTc,V_PGI.OkdecP);
    Valeurs [6] := Valeurs [6] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('PMAP')),V_PGI.OkdecP);
    Valeurs [7] := Valeurs [7] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('PMRP')),V_PGI.okdecP);
    if  ArticleOKInPOUR (TOBPD.getValue('TYPEARTICLE'),TypeartOk) then
        begin
        Valpou [0] := Valpou [0] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('DPA')),V_PGI.OKDecP);
        Valpou [1] := Valpou [1] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('DPR')),V_PGI.OKdecP);
        Valpou [2] := Valpou [2] + Arrondi(PxHt,V_PGI.OkDecP);
        Valpou [3] := Valpou [3] + Arrondi(PxTTc,V_PGI.OkdecP);
        Valpou [6] := Valpou [6] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('PMAP')),V_PGI.OKDecP);
        Valpou [7] := Valpou [7] + Arrondi(((QteLoc/QteDudetail) * TOBPD.GetValue ('PMRP')),V_PGI.OkdecP);
        end;
    END ;
    if IndPou <> 0 then
       begin
       TOBLN.detail[IndPou].Putvalue('DPA',Valpou[0]);TOBLN.detail[IndPou].Putvalue('DPR',Valpou[1]);
       TOBLN.detail[IndPou].Putvalue('PMAP',Valpou[6]);TOBLN.detail[IndPou].Putvalue('PMRP',Valpou[7]);
       TOBLN.detail[IndPou].Putvalue('PVHT',Valpou[2]);TOBLN.detail[IndPou].Putvalue('PVTTC',Valpou[3]);
       ValLoc := CalculSurTableau ('*',Valpou,TOBLN.detail[IndPou].getvalue('GNL_QTE')/100);
       FormatageTableau (ValLoc,V_PGI.OkdecP);
       Valeurs := CalculSurTableau ('+',Valeurs,ValLoc);
       end;
    FormatageTableau (valeurs,V_PGI.OkdecP);
END ;

function CalculePrixRevientOuv (TOBOUV,TOBART : TOB; PUA,COEFFG : double; Var MontantFg,MontantFc,MontantFr : double) : double;
var TOBA : TOB;
    MontantPA, MontantCharge: double;
    Qte,QteDuDetail : double;
begin
  //
  MontantFg := 0;
  MontantFc := 0;
  MontantFR := 0;
  //
  result := 0;
  Qte := TOBOUV.GetValue('BLO_QTEFACT');
  if Qte = 0 then exit;
  if TOBART = nil then
  begin
   TOBA := TOB.Create ('ARTICLE',nil,-1);
   TOBA.selectDb ('"'+TOBOUV.getValue('BLO_ARTICLE')+'"',nil);
  end else TOBA := TOBART;

  if TOBOUV.GetValue('BLO_QTEDUDETAIL') = 0 then TOBOUV.PutValue('BLO_QTEDUDETAIL',1) ;
  QteDuDetail := TOBOUV.GetValue('BLO_QTEDUDETAIL');

  montantPa := Arrondi((Qte*PUA)/QteDudetail,V_PGI.okDecP);
  //
  if (CoefFg <> 0) or ((TOBA <> nil) and (TOBA.GetValue('GA_DPRAUTO')='X')) then
  begin
   if TOBOUV.GetValue('BLO_NONAPPLICFG')<>'X' then  MontantFg := arrondi(MontantPa*(CoefFg-1),4);
   MontantCharge := MontantPA+MontantFG;
  end;
  //
  if TOBOUV.GetValue('BLO_NONAPPLICFC')<>'X' then MontantFC := arrondi(MontantCharge*TOBOUV.GetValue('BLO_COEFFC'),V_PGI.okDecP);
  MontantCharge := MontantPa+MontantFG+MontantFC;
  //
  if TOBOUV.GetValue('BLO_NONAPPLICFRAIS')<>'X' then MontantFr := arrondi(MontantCharge*TOBOUV.GetValue('BLO_COEFFR'),V_PGI.okdecP);
  MontantCharge := MontantPa+MontantFG+MontantFc+MontantFr;
  //
  result := arrondi((MontantCharge/Qte)*QteDudetail,V_PGI.okdecP) ;
  if TOBARt = nil then TOBA.Free;

end;

procedure CalculeLigneAcOuv (TOBOUv : TOB ; DEV : Rdevise;WithPuv : boolean=true; TOBART : TOB=nil);
var MontantCharge,CoefMarg,Qte,QteDudetail : double;
		MontantPr,MontantPAU,MontantFGU,MontantFCU,MontantFRU,MontantPRU : double;
		QQ : TQuery;
    TOBA : TOB;
begin
 if TOBART = nil then
 begin
   TOBA := TOB.Create ('ARTICLE',nil,-1);
   TOBA.selectDb ('"'+TOBOUV.getValue('BLO_ARTICLE')+'"',nil);
 end else TOBA := TOBART;
 Qte := TOBOUV.GetValue('BLO_QTEFACT');
 if TOBOUV.GetValue('BLO_QTEDUDETAIL') = 0 then TOBOUV.PutValue('BLO_QTEDUDETAIL',1) ;
 QteDuDetail := TOBOUV.GetValue('BLO_QTEDUDETAIL');
 // ----------------------------------------------------------------
 // initialisation
 // ----------------------------------------------------------------
 MontantPAu := 0;
 MontantFGU := 0;
 MontantFCU := 0;
 MontantFRU := 0;
 MontantPRU := 0;
 // ----------------------------------------------------------------
 // calcul de DPR et PUV en mode Unitaire
 // ----------------------------------------------------------------
 MontantPAU := Arrondi(TOBOUV.GetValue('BLO_QTEFACT')*TOBOUV.GetValue('BLO_DPA'),V_PGI.okDecP);
 if (TOBOUV.GetValue('BLO_COEFFG') <> 0) or ((TOBA <> nil) and (TOBA.GetValue('GA_DPRAUTO')='X')) then
 begin
 	MontantFGU := Arrondi(MontantPAU*TOBOUV.GetValue('BLO_COEFFG'),V_PGI.okDecP);
 end;
 MontantCharge := MontantPAU + MontantFGU;
 if TOBOUV.GetValue('BLO_NONAPPLICFC')<>'X' then
 begin
 	MontantFCU := Arrondi(MontantCharge *TOBOUV.GetValue('BLO_COEFFC'),V_PGI.okDecP);
 end;
 MontantCharge := MontantCharge + MontantFCU;
 if TOBOUV.GetValue('BLO_NONAPPLICFRAIS')<>'X' then
 begin
 	MontantFRU := Arrondi(MontantCharge *TOBOUV.GetValue('BLO_COEFFR'),V_PGI.okDecP);
 end;
 MontantPRU := MontantCharge + MontantFRU;
 if QTe = 0 then Qte := 1;
 TOBOUV.setDouble('BLO_DPR',Arrondi(MontantPRU / Qte,V_PGI.okdecP));
 if WithPuv then
 begin
   if (TOBOUV.GetValue('BLO_COEFMARG') <> 0) and (TOBOUV.GetValue('BLO_DPR')<>0 ) and (TOBA.GetValue('GA_CALCAUTOHT')='X') then
   begin
   	TOBOUV.putValue('BLO_PUHT',arrondi(TOBOUV.GetValue('BLO_DPR')*TOBOUV.GetValue('BLO_COEFMARG'),V_PGI.okdecP));
   end else
   begin
   	TOBOUV.putValue('BLO_PUHT',arrondi(TOBOUV.GetValue('BLO_PUHT'),V_PGI.okdecP));
   	TOBOUV.PutValue('BLO_COEFMARG',0);
   end;
//                                           else TOBOUV.putValue('BLO_PUHT',TOBOUV.GetValue('BLO_DPR'));
   TOBOuv.PutValue ('BLO_PUHTDEV',pivottodevise(TobOuv.GetValue('BLO_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
 end else
 begin
    if TOBOUV.GetValue('BLO_DPR') <> 0 then
    begin
      TOBOUV.putValue('BLO_COEFMARG',arrondi(TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR'),4));
      TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
 end;
 end;
  if TOBOUV.GetValue('BLO_PUHT') <> 0 then
  begin
    TOBOuv.PutValue('POURCENTMARQ',Arrondi(((TOBOUV.GetValue('BLO_PUHT')- TOBOUV.GetValue('BLO_DPR'))/TOBOUV.GetValue('BLO_PUHT'))*100,2));
  end else
  begin
    TOBOuv.PutValue('POURCENTMARQ',0);
  end;

 // ----------------------------------------------------------------
 // Calcul des montant avec prise en charge des detail exprime pour
 // ----------------------------------------------------------------
 TOBOUV.putValue('BLO_MONTANTPA',Arrondi(MontantPAU/QteDudetail,V_PGI.okDecP));
 //
 TOBOUV.putValue('BLO_MONTANTPAFG',TOBOUV.GetValue('BLO_MONTANTPA'));
 TOBOUV.putValue('BLO_MONTANTPAFC',TOBOUV.GetValue('BLO_MONTANTPA'));
 TOBOUV.putValue('BLO_MONTANTPAFR',TOBOUV.GetValue('BLO_MONTANTPA'));
 //
 TOBOUV.putValue('BLO_MONTANTFG',Arrondi(MontantFGU/QteDudetail,V_PGI.okDecP));
 TOBOUV.putValue('BLO_MONTANTFC',Arrondi(MontantFCU/QteDudetail,V_PGI.okDecP));
 TOBOUV.putValue('BLO_MONTANTFR',Arrondi(MontantFRU/QteDudetail,V_PGI.okDecP));
 TOBOUV.putValue('BLO_MONTANTPR',Arrondi(MontantPRU/QteDudetail,V_PGI.okDecP));
 //
 if TOBARt = nil then TOBA.Free;
end;

procedure CalculeLigneAcOuvCumul (TOBOUv : TOB);
var MontantCharge,CoefMarg,Qte : double;
begin
  Qte := TOBOUV.GetValue('BLO_QTEFACT');

  TOBOUV.putValue('BLO_MONTANTPA',Arrondi((TOBOUV.GetValue('BLO_QTEFACT')*TOBOUV.GetValue('BLO_DPA')),V_PGI.okdecP));
  //
  TOBOUV.putValue('BLO_MONTANTPAFG',TOBOUV.GetValue('BLO_MONTANTPA'));
  TOBOUV.putValue('BLO_MONTANTPAFC',TOBOUV.GetValue('BLO_MONTANTPA'));
  TOBOUV.putValue('BLO_MONTANTPAFR',TOBOUV.GetValue('BLO_MONTANTPA'));
  //
  MontantCharge := TOBOUV.GetValue('BLO_MONTANTPA')+TOBOUV.GetValue('BLO_MONTANTFG')+TOBOUV.GetValue('BLO_MONTANTFC')+TOBOUV.GetValue('BLO_MONTANTFR');
  //
  TOBOUV.putValue('BLO_MONTANTPR',MontantCharge);

  if TOBOUV.GetValue('BLO_MONTANTPAFG') <> 0 then
  begin
    TOBOUV.PutValue('BLO_COEFFG',arrondi(TOBOUV.GetValue('BLO_MONTANTFG')/TOBOUV.GetValue('BLO_MONTANTPAFG'),4));
  end;
  //
  if TOBOUV.GetValue('BLO_MONTANTPAFC') <> 0 then
  begin
    TOBOUV.PutValue('BLO_COEFFC',arrondi(TOBOUV.GetValue('BLO_MONTANTFC')/TOBOUV.GetValue('BLO_MONTANTPAFC'),4));
  end;

  if TOBOUV.GetValue('BLO_MONTANTPAFR') <> 0 then
  begin
    TOBOUV.PutValue('BLO_COEFFR',arrondi(TOBOUV.GetValue('BLO_MONTANTFR')/TOBOUV.GetValue('BLO_MONTANTPAFR'),4));
  end;

  if TOBOUV.GetValue('BLO_DPR') <> 0 then
  begin
    TOBOUV.PutValue('BLO_COEFMARG',arrondi(TOBOUV.GetValue('BLO_PUHT')/TOBOUV.GetValue('BLO_DPR'),4));
    TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));
  end else TOBOUV.PutValue('BLO_COEFMARG',0);

  if TOBOUV.GetValue('BLO_PUHT') <> 0 then
  begin
    TOBOuv.PutValue('POURCENTMARQ',Arrondi(((TOBOUV.GetValue('BLO_PUHT')- TOBOUV.GetValue('BLO_DPR'))/TOBOUV.GetValue('BLO_PUHT'))*100,2));
  end else
  begin
    TOBOuv.PutValue('POURCENTMARQ',0);
  end;

end;

Procedure OuvVersLigOuv ( TOBPIece,TOBL,TOBNOMEN,TOBLN,TOBA : TOB ; Niv,OrdreCompo : integer ;DEV:RDevise;EnHt:Boolean;Prixtraite : string='PUH' ) ;
Var i : integer ;
    TOBNLD,TOBLND,TOBART,TOBDetOuv,TOBD : TOB ;
    prix,QTe,QteDuPv,QteDuDetail : double;
    pvinit,pvfin,PrixPr,valrech : double;
    VenteAchat,CodeArt : string;
    Reajuste,OKGestQteAvance : boolean;
    Depot : string;
    StINFG,StInFc : boolean;
    TypeNomenc,RecupPrix : string;
    QteSit,MtSit : double;
BEGIN
  RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');
  TypeNomenc := TOBL.GetValue('GL_TYPENOMENC');
  VenteAchat := GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  reajuste :=  isExistsArticle (trim(GetParamsoc('SO_BTECARTPMA')));
  StInFg := (TOBPiece.getValue('GP_APPLICFGST')='X');
  StInFC := (TOBPiece.getValue('GP_APPLICFCST')='X');
  OKGestQteAvance := (GetParamSocSecur('SO_BTGESTQTEAVANC',false))  and (ISDocumentChiffrage(TOBPiece)) ;

  for i:=0 to TOBNOMEN.Detail.Count-1 do
    BEGIN
    TOBNLD:=TOBNOMEN.Detail[i] ;
    TOBART := TOBA.FIndfirst (['GA_ARTICLE'],[TOBNLD.GetValue('GNL_ARTICLE')],false);
    if TOBART = nil then continue;
    CodeArt := TOBART.getString('GA_ARTICLE');
    TOBLND:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
    InsertionChampSupOuv (TOBLND,false);
    {
    TOBLND.addchampsup ('ANCPV',false);
    }
    //
    (*
    if (TOBART.GetValue('GA_TYPEARTICLE')='PRE') and (not StInFg) and (Pos(TOBART.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)=0) then TOBLND.PutValue('BLO_NONAPPLICFRAIS','X') ;
    if (TOBART.GetValue('GA_TYPEARTICLE')='PRE') and (not StInFC) and (Pos(TOBART.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)=0) then TOBLND.PutValue('BLO_NONAPPLICFC','X') ;
    *)
    //
    TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
    TOBLND.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    // --
    if TOBL.GetValue('GL_DOMAINE')<> '' then
    begin
    	TOBLND.putValue('BLO_DOMAINE',TOBL.GetValue('GL_DOMAINE'));
    end else
    begin
    	TOBLND.PutValue('BLO_DOMAINE',TOBNLD.GetValue('GNE_DOMAINE')) ;
    end;
    TOBLND.PutValue('BLO_ARTICLE',TOBNLD.GetValue('GNL_ARTICLE')) ;
    TOBLND.PutValue('BLO_CODEARTICLE',TOBNLD.GetValue('GNL_CODEARTICLE')) ;
    TOBLND.PutValue('BLO_REFARTSAISIE',TOBNLD.GetValue('GNL_CODEARTICLE')) ;
    TOBLND.PutValue('BLO_NUMORDRE',TOBNLD.GetValue('GNL_NUMLIGNE')) ;
    TOBLND.PutValue('BLO_LIBELLE',TOBNLD.GetValue('GNL_LIBELLE')) ;
    // -
    if (OKGestQteAvance) then
    begin
      TOBLND.SetSTring('BLO_QUALIFHEURE',TOBNLD.GetString('GNL_QUALIFUNITEACH'));
      TOBLND.SetDouble('BLO_RENDEMENT',  TOBNLD.GetDouble('GNL_RENDEMENT'));
      TOBLND.SetDouble('BLO_PERTE', TOBNLD.GetDouble('GNL_PERTE'));
      TOBLND.SetDouble('BLO_QTESAIS',TOBNLD.GetDouble('GNL_QTESAIS'));
    end else
    begin
      TOBLND.SetDouble('BLO_QTESAIS',TOBNLD.GetDouble('GNL_QTE'));
    end;
    // -
    TOBLND.PutValue('BLO_QTEFACT',TOBNLD.GetValue('GNL_QTE')) ;
    TOBLND.PutValue('BLO_COMPOSE',TOBNLD.GetValue('COMPOSE')) ;
    TOBLND.PutValue('BLO_TENUESTOCK',TOBNLD.GetValue('TENUESTOCK')) ;
    TOBLND.PutValue('BLO_QUALIFQTEVTE',TOBNLD.GetValue('QUALIFQTEVTE')) ;
    TOBLND.PutValue('LIBELLEFOU',TOBNLD.GetValue('LIBELLEFOU')) ;
    if (Pos (TOBArt.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0)  and (not IsLigneExternalise(TOBLND)) then
    begin
      TOBLND.PutValue('BLO_TPSUNITAIRE',1);
    end else
    begin
      TOBLND.PutValue('BLO_TPSUNITAIRE',0);
    end;
    if (TOBLND.GetValue('BLO_TENUESTOCK')='X') and (pos(TOBLND.GetValue('BLO_NATUREPIECEG'),'BLC;LBT')>0) then
    begin
      Depot := TOBLND.GetValue('BLO_DEPOT');
    	TOBD := TOBArt.FindFirst(['GQ_DEPOT'], [Depot], False);
    	if (TOBD <> nil) and (TOBD.getValue('GQ_PMAP')<>0) then
      begin
    		TOBLND.PutValue('BLO_DPA',TOBD.GetValue('GQ_PMAP')) ;
      end else
      begin
    		TOBLND.PutValue('BLO_DPA',TOBNLD.GetValue('DPA')) ;
      end;
    end else
    begin
    	TOBLND.PutValue('BLO_DPA',TOBNLD.GetValue('DPA')) ;
    end;
    //    TOBLND.PutValue('BLO_DPR',TOBNLD.GetValue('DPR')) ;
    TOBLND.PutValue('BLO_DPR',TOBNLD.GetValue('DPA')+(TOBNLD.GetValue('DPA')*TOBLND.GetValue('BLO_COEFFG'))) ;
    TOBLND.PutValue('BLO_PMAP',TOBNLD.GetValue('PMAP')) ;
    TOBLND.PutValue('BLO_PMRP',TOBNLD.GetValue('PMRP')) ;
    (* OPTIMIZATION *)
    TOBLND.PutValue('BLO_QUALIFQTEACH',TOBNLD.GetValue('QUALIFQTEACH')) ;
    TOBLND.PutValue('BLO_PRXACHBASE',TOBNLD.GetValue('PRXACHBASE')) ;
    TOBLND.PutValue('BLO_REMISES',TOBNLD.GetValue('REMISES')) ;
    TOBLND.PutValue('BLO_NOMENCLATURE',TOBNLD.GetValue('GNL_NOMENCLATURE')) ;
    (* ------------- *)
    TOBLND.PutValue('BLO_QTEDUDETAIL',TOBNLD.getValue('QTEDUDETAIL'));
    //    TOBLND.PutValue('BLO_LIBCOMPL',TOBNLD.getValue('LIBCOMPL'));
    // Provenance ligne
    TOBLND.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    TOBLND.PutValue('BLO_REGIMETAXE',TOBL.GetValue('GL_REGIMETAXE')) ;
    TOBLND.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    TOBLND.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    TOBLND.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    TOBLND.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    TOBLND.PutValue('BLO_NIVEAU',Niv) ;
    TOBLND.PutValue('BLO_ORDRECOMPO',OrdreCompo) ;
    CopieOuvFromLigne(TOBLND,TOBL); // recup des infos en provenance de la ligne de doc.
    TOBLND.putvalue ('BLO_TYPELIGNE','ART');
    //    TOBART := TOBA.FIndfirst (['GA_ARTICLE'],[TOBNLD.GetValue('GNL_ARTICLE')],false);
    CopieOuvFromArt (TobPiece,TOBLND,TOBART,DEV,RecupPrix);
    if TypeNomenc = 'OU1' then TOBLND.PutValue('BLO_FAMILLETAXE1',TOBL.getValue('GL_FAMILLETAXE1'));
    //
    if (IsPrestationSt(TOBLND)) and (not StInFg) then TOBLND.PutValue('BLO_NONAPPLICFRAIS','X') ;
    if (IsPrestationSt(TOBLND)) and (not StInFC) then TOBLND.PutValue('BLO_NONAPPLICFC','X') ;
    //
    if TOBArt.GetValue('GA_TYPEARTICLE')='ARP' then
    begin
      if (TOBLND.GetValue('BLO_TENUESTOCK')='X') and (pos(TOBLND.GetValue('BLO_NATUREPIECEG'),'BLC;LBT')>0) then
      begin
        Depot := TOBLND.GetValue('BLO_DEPOT');
        TOBD := TOBArt.FindFirst(['GQ_DEPOT'], [Depot], False);
        if TOBD <> nil then
        begin
          TOBLND.PutValue('BLO_DPA',TOBD.GetValue('GQ_PMAP')) ;
        end else
        begin
          TOBLND.PutValue('BLO_DPA',TOBArt.GetValue('GA_PAHT')) ;
        end;
      end else
      begin
        TOBLND.PutValue('BLO_DPA',TOBArt.GetValue('GA_PAHT')) ;
      end;
      TOBLND.PutValue('BLO_DPR',TOBArt.GetValue('GA_DPR')) ;
      TOBNLD.PutValue('PVHT',TOBART.GetValue('GA_PVHT'));
    end;
    // Gestion des prix
    if Prixtraite = 'PUH' then
    begin
      if TOBNLD.GetValue ('GNL_PRIXFIXE') <> 0 then
      begin
         prix := pivottodevise(TOBNLD.GetValue('GNL_PRIXFIXE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP );
         TOBLND.putvalue ('BLO_COEFMARG',0);
      end else prix := TOBNLD.GetValue ('PVHT');
    end else if PrixTraite = 'DPR' then
    begin
    	prix := TOBArt.GetValue('GA_DPR');
    end else if PrixTraite = 'DPA' then
    begin
    	prix := TOBArt.GetValue('GA_PAHT');
    end;
    //
    if TOBL.Getvalue('GL_SAISIECONTRE')='X' then
       BEGIN
       if VH^.TenueEuro then Prix:=EuroToFranc(Prix) else Prix:=FrancToEuro(Prix) ;
       END ;
    //
    TOBLND.putvalue ('BLO_PUHTDEV',prix);
    TOBLND.putvalue ('BLO_PUTTCDEV',TOBNLD.getValue('PVTTC'));
    TOBLND.putvalue ('BLO_PUHT',devisetopivotEx(prix,DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBLND.putvalue ('BLO_PUTTC',devisetopivotEx ( TOBNLD.GetValue('PVTTC'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBLND.putvalue ('BLO_PUHTBASE',TOBLND.GetValue('BLO_PUHT'));
    TOBLND.putvalue ('BLO_PUTTCBASE',TOBLND.GetValue('BLO_PUTTC'));
    //
    //
    if TOBART.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBLND,'A',RecupPrix);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBLND,DEV);
    end;
    CalculeLigneAcOuv (TOBLND,DEV,true,TOBART);
    if EnHT then
       BEGIN
       CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END else
       BEGIN
       if TOBLND.Getvalue ('BLO_PUTTCDEV') <> 0 then
      begin
          CalculeLigneTTCOuv(TOBLND,TOBPiece,DEV)
      end else CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END;
		TOBLND.putvalue ('ANCPA',TOBLND.Getvalue ('BLO_DPA'));
		TOBLND.putvalue ('ANCPR',TOBLND.Getvalue ('BLO_DPR'));
    TOBLND.PutValue('GA_PVHT',TOBLND.GetValue('BLO_PUHTDEV'));
    if EnHt then TOBLND.putvalue ('ANCPV',TOBLND.Getvalue ('BLO_PUHTDEV'))
    				else TOBLND.putvalue ('ANCPV',TOBLND.Getvalue ('BLO_PUTTCDEV'));
    Qte := TOBLND.GetValue('BLO_QTEFACT');
    QteDuPv := TOBLND.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPV := 1;
	  if TOBLND.GetValue('BLO_QTEDUDETAIL') = 0 then TOBLND.PutValue('BLO_QTEDUDETAIL',1) ;
    QteDuDetail := TOBLND.GetValue ('BLO_QTEDUDETAIL');
    // fv 02032004
    if QteDudetail = 0 then QteDuDetail := 1;
    {$IFDEF BTP}
    //if TheMetreShare <> nil then
    //begin
    //  IF (TOBL.GetValue('GL_TYPEARTICLE')='OUV') and
    //     (TOBArt.GetValue('GA_TYPEARTICLE')<>'ARP') and
    //     (VenteAchat = 'VEN') then
    //     if TheMetreShare.AutorisationMetre(TOBL.GetValue('GL_NATUREPIECEG')) then
    //        DocMetreBiblioDetail (TOBNLD,TOBLND);
    //end;
    {$ENDIF}
    OuvVersLigOuv(TOBPiece,TOBL,TOBNLD,TOBLND,TOBA,Niv+1,TOBLND.GetValue('BLO_NUMORDRE'),DEV,EnHt) ;

    if EnHT then
       BEGIN
       CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END else
       BEGIN
       if TOBLND.Getvalue ('BLO_PUTTCDEV') <> 0 then
      begin
          CalculeLigneTTCOuv(TOBLND,TOBPiece,DEV)
      end else CalculeLigneHTOuv(TOBLND,TOBPiece,DEV) ;
       END;
    if (TOBLND.nomtable = 'LIGNEOUV') and
       (TOBL.GetString('GL_PIECEPRECEDENTE')='') and
       (TOBLND.FieldExists('BLF_QTESITUATION')) and
       (Niv = 1) then
    begin
      QteSit := arrondi(TOBLND.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBLND.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      MtSit := arrondi(TOBLND.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBLND.getValue('BLO_QTEDUDETAIL')*TOBLND.getValue('BLO_PUHTDEV'),DEV.Decimale);
      TOBLND.PutValue('BLF_QTESITUATION', QteSit); { NEWPIECE }
      TOBLND.PutValue('BLF_MTSITUATION',MtSit);
      TOBLND.PutValue('BLF_POURCENTAVANC',100);
    end;
    //    CalculMontantHtDevLigOuv (TOBLND,DEV);
    END ;
   // Controle et ajustement pour prix forfait
  //
  if TOBLN.nomtable = 'LIGNEOUV' then
  begin
    TOBLN.PutValue('BLO_MONTANTHTDEV',0);
    TOBLN.PutValue('BLO_MONTANTPA',0);
    TOBLN.PutValue('BLO_MONTANTPR',0);
    TOBLN.PutValue('BLO_MONTANTTTCDEV',0);
  end;
  //
  for i:=0 to TOBLN.Detail.Count-1 do
    begin
    TOBDetOuv := TOBLN.detail[I];
    pvinit := pivottodevise (TOBDetOuv.getValue('ANCPV') ,DEV.taux,DEV.quotite,V_PGI.OkdecP);
    PrixPr := TOBDetOuv.getValue('BLO_DPR');
    if EnHt then pvfin := TOBDetOuv.getValue('BLO_PUHTDEV')
    else pvfin := TOBDetOuv.getValue('BLO_PUTTCDEV');

    if (pvinit <> pvfin) and (TOBDetOuv.detail.count > 0) and (reajuste) then
       begin
       ReajusteMontantOuvrage (TOBA,TOBPiece,TOBL,TOBDetOuv,pvinit,PrixPr,pvfin,DEV,EnHt,True);
       end;
    if tobLN.nomtable = 'LIGNEOUV' then
    begin
      TOBLN.PutValue('GA_PVHT',TOBLN.GetValue('BLO_PUHTDEV'));
      TOBLN.PutValue('BLO_MONTANTHTDEV',TOBLN.GetValue('BLO_MONTANTHTDEV')+TOBDetOUV.GetValue('BLO_MONTANTHTDEV'));
      TOBLN.PutValue('BLO_MONTANTPA',TOBLN.GetValue('BLO_MONTANTPA')+TOBDetOUV.GetValue('BLO_MONTANTPA'));
      TOBLN.PutValue('BLO_MONTANTPR',TOBLN.GetValue('BLO_MONTANTPR')+TOBDetOUV.GetValue('BLO_MONTANTPR'));
      TOBLN.PutValue('BLO_MONTANTTTCDEV',TOBLN.GetValue('BLO_MONTANTTTCDEV')+TOBDetOUV.GetValue('BLO_MONTANTTTCDEV'));
    end;
    end;
END ;

procedure DefiniPrixLigneOuv (TOBouvrage,TOBNomen,TOBART,TOBPiece : TOb;DEV:Rdevise);
var prix: double;
begin
if TOBNomen.GetValue ('GNL_PRIXFIXE') <> 0 then
begin
   prix := TOBnomen.GetValue('GNL_PRIXFIXE');
end
else prix := TOBArt.GetValue ('GA_PVHT');
if TOBPiece.GetValue('GL_DEVISE') <> V_PGI.DevisePivot then
   BEGIN
   TOBOuvrage.putvalue ('BLO_PUHTDEV',pivottodevise (prix,DEV.taux,DEV.quotite,V_PGI.OkdecP));
   TOBOuvrage.putvalue ('BLO_PUTTCDEV',pivottodevise (TOBART.getValue('GA_PVTTC'),DEV.taux,DEV.quotite,V_PGI.OkdecP));
   END ELSE
   BEGIN
   TOBOuvrage.putvalue ('BLO_PUHTDEV',prix);
   TOBOuvrage.putvalue ('BLO_PUTTCDEV',TOBART.getValue('GA_PVTTC'));
   END;
TOBOuvrage.putvalue ('BLO_PUHT',prix);
TOBOuvrage.putvalue ('BLO_PUHTBASE',TOBART.GetValue('GA_PVHT'));
TOBOuvrage.putvalue ('BLO_PUTTCBASE',TOBART.GetValue('GA_PVTTC'));
TOBOuvrage.putvalue ('BLO_PUTTC',TOBART.GetValue('GA_PVTTC'));
CalculMontantHtDevLigOuv (TOBOUVrage,DEV);
end;

procedure recupInfoLigneOuv (TOBDEST,TOBRef : TOB);
var i : integer;
    nomChamp : string;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  FieldList := MCD.GetTable(TOBRef.NomTable).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomChamp := (FieldList.Current as IFieldCOM).name;
    if (Nomchamp <> '') and (Pos(NomChamp,'BLO_N1;BLO_N2;BLO_N3;BLO_N4;BLO_N5')=0) then
    begin
    	TOBDEST.PutValue (NomChamp,TOBREF.Getvalue(NomChamp));
    end;
  end;
  TOBDest.Putvalue ('UTILISE',TOBREF.GetValue('UTILISE')) ;
  TOBDest.Putvalue ('ANCPV',TOBREF.GetValue('ANCPV')) ;
  TOBDest.Putvalue ('ANCPA',TOBREF.GetValue('ANCPA')) ;
  TOBDest.Putvalue ('ANCPR',TOBREF.GetValue('ANCPR')) ;
  (*
  TOBDest.Putvalue ('MONTANTHTDEV',TOBREF.GetValue('MONTANTHTDEV')) ;
  TOBDest.Putvalue ('MONTANTTTCDEV',TOBREF.GetValue('MONTANTTTCDEV')) ;
  TOBDest.Putvalue ('MONTANTTTCDEV',TOBREF.GetValue('MONTANTTTCDEV')) ;
  TOBDest.Putvalue ('MONTANTACHAT',TOBREF.GetValue('MONTANTACHAT')) ;
  TOBDest.Putvalue ('MONTANTPR',TOBREF.GetValue('MONTANTPR')) ;
  *)
  TOBDest.Putvalue ('SOUSNOMEN',TOBREF.GetValue('SOUSNOMEN')) ;
  TOBDest.Putvalue ('LIBELLE',TOBREF.GetValue('LIBELLE')) ;
  TOBDest.Putvalue ('CODENOMEN',TOBREF.GetValue('CODENOMEN')) ;
  (* OPTIMIZATION *)
  TOBDest.Putvalue ('BLO_FOURNISSEUR',TOBREF.GetValue('GA_FOURNPRINC')) ;
  TOBDest.Putvalue ('BLO_REMISES',TOBREF.GetValue('GF_CALCULREMISE')) ;
  TOBDest.Putvalue ('BLO_QUALIFQTEACH',TOBREF.GetValue('GCA_QUALIFUNITEACH')) ;
  TOBDest.Putvalue ('BLO_NOMENCLATURE',TOBREF.GetValue('BLO_NOMENCLATURE')) ;
  TOBDest.Putvalue ('BLO_PRXACHBASE',TOBREF.GetValue('GCA_PRIXBASE')) ;
  (* ------------ *)
  // NEWONE
  TOBDest.Putvalue ('TGA_FOURNPRINC',TOBREF.GetValue('TGA_FOURNPRINC')) ;
  TOBDest.Putvalue ('GA_FOURNPRINC',TOBREF.GetValue('GA_FOURNPRINC')) ;
  TOBDest.Putvalue ('GCA_PRIXBASE',TOBREF.GetValue('GCA_PRIXBASE')) ;
  TOBDest.Putvalue ('GCA_QUALIFUNITEACH',TOBREF.GetValue('GCA_QUALIFUNITEACH')) ;
  TOBDest.Putvalue ('GF_CALCULREMISE',TOBREF.GetValue('GF_CALCULREMISE')) ;
  TOBDest.Putvalue ('GA_PAHT',TOBREF.GetValue('GA_PAHT')) ;
  TOBDest.Putvalue ('GA_PVHT',TOBREF.GetValue('GA_PVHT')) ;
  TOBDest.Putvalue ('GA_NATUREPRES',TOBREF.GetValue('GA_NATUREPRES')) ;
  TOBDest.Putvalue ('GA_ARTICLE',TOBREF.GetValue('GA_ARTICLE')) ;
  TOBDest.Putvalue ('TGA_PRESTATION',TOBREF.GetValue('TGA_PRESTATION')) ;
  TOBDest.Putvalue ('GA_HEURE',TOBREF.GetValue('GA_HEURE')) ;
  TOBDest.Putvalue ('GF_PRIXUNITAIRE',TOBREF.GetValue('GF_PRIXUNITAIRE')) ;
  TOBDest.Putvalue ('BLO_TPSUNITAIRE',TOBREF.GetValue('BLO_TPSUNITAIRE')) ;
  TOBDest.Putvalue ('BLO_TYPEARTICLE',TOBREF.GetValue('BLO_TYPEARTICLE')) ;
  //
  TOBDest.Putvalue ('BLO_ARTICLE',TOBREF.GetValue('BLO_ARTICLE')) ;
  TOBDest.Putvalue ('BLO_CODEARTICLE',TOBREF.GetValue('BLO_CODEARTICLE')) ;
  TOBDest.Putvalue ('BLO_REFARTSAISIE',TOBREF.GetValue('BLO_REFARTSAISIE')) ;
  TOBDest.Putvalue ('BLO_AFFAIRE1',TOBREF.GetValue('BLO_AFFAIRE1')) ;
  TOBDest.Putvalue ('BLO_AFFAIRE2',TOBREF.GetValue('BLO_AFFAIRE2')) ;
  TOBDest.Putvalue ('BLO_AFFAIRE3',TOBREF.GetValue('BLO_AFFAIRE3')) ;
  TOBDest.Putvalue ('BLO_AVENANT',TOBREF.GetValue('BLO_AVENANT')) ;
  TOBDest.Putvalue ('BLO_LIBELLE',TOBREF.GetValue('BLO_LIBELLE')) ;
  TOBDest.Putvalue ('BLO_FOURNISSEUR',TOBREF.GetValue('BLO_FOURNISSEUR')) ;
end;


procedure GestionDetailPrixPose (TOBDepart : TOB);
var indice : integer;
    TOBL : TOB;
begin
  if TobDepart.detail.count = 0 then exit;
  For Indice := 0 to TOBDEPART.detail.count -1 do
  begin
    TOBL := TOBDepart.detail[Indice];
    if (TOBL.GetValue ('BLO_TYPEARTICLE') = 'ARP') and (TOBL.detail.count > 0) then GetValoDetail (TOBL)
    else if TOBL.GetValue('BLO_TYPEARTICLE') = 'OUV' then GestionDetailPrixPose (TOBL);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Cr le ...... : 10/09/2003
Modifi le ... :   /  /
Description .. : Rcupration de la zone PGI-RESULTAT dans les fichier
Suite ........ : XLS qui compose le mtr de l'ouvrage
Mots clefs ... : GESTION DES MTRS
*****************************************************************}
{*procedure VaChercherQuantitatifDetail (TObDepart : TOB);
var Indice : integer;
    TOBP : TOB;
    ValRet : variant;
    Prefixe : String;
begin

  for Indice := 0 to TOBDepart.detail.count -1 do
  begin
    TOBP := TOBDepart.detail[Indice];

    if TOBP.detail.count > 0 then
    begin
      VaChercherQuantitatifDetail (TOBP)
    end;

    if TOBP.NomTable = 'LIGNE' then
       Prefixe := 'GL'
    else if TOBP.NomTable = 'LIGNEOUV' then
       prefixe := 'BLO'
    else
       exit;

    ValRet := 0;
    if (TheMetreShare <> nil) and (TheMetreShare.AutorisationMetre(TOBP.GetValue(Prefixe + '_NATUREPIECEG'))) then
    begin
    	ValRet := TheMetreShare.GetValeurMetre (TOBP,false);
    end;

    if ValRet = 0 then continue;

    TOBP.Putvalue(prefixe+'_QTEFACT',Arrondi(VarasType(valret,vardouble),V_PGI.OkDecQ));
    if Prefixe = 'GL' then
    begin
      TOBP.Putvalue(prefixe+'_QTERESTE',Arrondi(VarasType(valret,vardouble),V_PGI.OkDecQ));
      TOBP.Putvalue(prefixe+'_QTESTOCK',Arrondi(VarasType(valret,vardouble),V_PGI.OkDecQ));
      if CtrlOkReliquat(TOBP, Prefixe) then TOBP.Putvalue(prefixe+'_QTERESTE',Arrondi(VarasType(valret,vardouble),V_PGI.OkDecP));
    end;
  end;

end;
*}

function TraiteLesOuvrages ( XX: Tform; TOBPiece,TOBArticles,TOBOuvrage,TOBRepart,TobMetres : TOB ; ARow : integer ; LePremier : boolean ;DEV :RDEVISE;OrigineExcel:boolean;OptimizeOuv : TOptimizeOuv = nil;WithDetails : boolean=true;FromBordereau : boolean=false; ModeSaisieAch : boolean= false) : boolean;
Var RefUnique,TypeArt,TypeNomenc,Depot : String ;
    TOBL,TOBNomen,TOBLN,TOBTMP,TOBDEPART : TOB ;
    IndiceNomen : integer ;
    conversion,arrondiOk : boolean;
    Valeurs : T_Valeurs;
    CodeDevCli : String;
    EnHt : Boolean;
    Bidon : double;
    Affichage,NiveauDepart : Integer;
    SaisieContreVal : boolean;
    Libelle : string;
    VenteAchat,RecupPrix : string;
    PrxVenteSto,ValPr : double;
    CalculAutorise : boolean;
BEGIN

  if not assigned(Tobpiece) then exit;
  CalculAutorise := true;
  if XX <> nil then CalculAutorise := TFFacture(XX).CalculPieceAutorise;
  //
  VenteAchat := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT');
  RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');

  NiveauDepart := 1;
  result := false;

  SaisieContreVal := (TOBPiece.Getvalue('GP_SAISIECONTRE')='X');

  conversion := false;
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  //if (OrigineEXCEL) or (not WithDetails) then
  if (not WithDetails) then
  begin
    Affichage := DOU_AUCUN;
  end else
  begin
    Affichage:=GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'),'GPP_TYPEPRESENT');
  end;
  RefUnique:=GetCodeArtUnique(TOBPiece,ARow) ; if RefUnique='' then Exit ;
  TOBL:=GetTOBLigne(TOBPiece,ARow) ; if TOBL=Nil then Exit ;
  // Correction LS si provenance de la recherche article
  if TOBL.GetValue('GL_INDICENOMEN') > 0 then exit;
  // --
  CodeDevCli:= TOBL.GetValue ('GL_DEVISE');
  TypeArt:=TOBL.GetValue('GL_TYPEARTICLE') ;
  if (TypeArt<>'OUV') and (TypeArt <> 'ARP') then
  BEGIN
    result := true;
    Exit ;
  END;
  TypeNomenc:=TOBL.GetValue('GL_TYPENOMENC') ;
  if (TypeNomenc<>'OUV') and (TypeNomenc <> 'OU1') and (TypeArt <> 'ARP') then Exit ;

  if FromBordereau then
  begin
    TOBL.putValue('GLC_FROMBORDEREAU','X');
    TOBL.putValue('GL_BLOQUETARIF','X');
    PrxVenteSto := TOBL.GetValue('GL_PUHTDEV');
  end;

  // Correction pour les pieces d'achat
  // On ne peut pas acheter des ouvrages ou la prestation associe au prix pose
  //if (VenteAchat = 'ACH') or (pos (TOBPiece.getValue('GP_NATUREPIECEG'),'DBT;ETU;FBT;ABT;FRC') = 0) then
  if Not(ApplicPrixPose (TOBpiece)) then
  BEGIN
    result := true;
    exit;
  END;
  // --
  //Depot:=TOBL.GetValue('GL_DEPOT') ; if Depot='' then Exit ;
  TOBNomen:=ChoixOuvrage(RefUnique,Depot,Libelle,TOBArticles,False,DEV,VenteAchat) ;
  if TOBNomen = nil then exit;


  //Description .. : Appel de la recherche du fichier mtrs associ  l'ouvrage
  {$IFDEF BTP}
  //if TheMetreShare <> nil then
  //begin
  //  if (TypeArt='OUV') and (VenteAchat = 'VEN')  then
  //  if TheMetreShare.AutorisationMetre(TOBL.GetValue('GL_NATUREPIECEG')) then
  //  DocMetreBiblioEntete (TOBL, TOBL.GETVALUE('GL_ARTICLE'),TOBL.getValue('GL_TYPEARTICLE'));
  //end;
  {$ENDIF}

  TOBLN:=TOB.Create('',TOBOuvrage,-1) ;
  InsertionChampSupOuv (TOBLN,false);
  if ((TOBNomen<>Nil) and (TOBNomen.Detail.Count>0)) then
  BEGIN
    InitTableau (Valeurs);
    RenseigneValoOuvrage(TOBNomen,CodeDevCli,1,1,DEV,Valeurs) ;
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      TOBDEPART := TOB.Create('LIGNEOUV',TOBLN,-1) ;
      InsertionChampSupOuv (TOBDEPART,false);
      inc (NiveauDepart);
    end else
    begin
      TOBDEPART := TOBLN;
      if Not FromBordereau then TOBL.putvalue('GL_LIBELLE', Libelle);
    end;
    OuvVersLigOuv(TOBPiece,TOBL,TOBNomen,TOBDEPART,TOBArticles,NiveauDepart,0,DEV,EnHt) ;
    NumeroteLigneOuv (TOBLN,TOBL,1,1,0,0,0);
    {$IFDEF BTP}
    //if TheMetreShare <> nil then
    //begin
    //  if (TypeArt='OUV') and (VenteAchat = 'VEN') then
    //  if TheMetreShare.AutorisationMetre(TOBL.GetValue('GL_NATUREPIECEG')) then
    //    TheMetreShare.SaisieVar(TOBL);
    //end;
    {$ENDIF}
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      recupInfoLigneOuv (TOBDEPART,TOBDEPART.Detail[0]);
      TOBDEPART.PutValue('BLO_NATUREPIECEG',TOBDEPART.detail[0].GetValue('BLO_NATUREPIECEG'));
      TOBDEPART.PutValue('BLO_QTEFACT',1) ;
      TOBDEPART.PutValue('BLO_NIVEAU',1) ;
      TOBDEPART.PutValue('BLO_NUMORDRE',1) ;
      TOBDEPART.PutValue('BLO_QTEDUDETAIL',1) ;
      TOBDEPART.PutValue('BLO_PRIXPOURQTE',1) ;
      GetValoDetail (TOBDEPART);
    end;
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBDEPART,1,1,true,DEV,valeurs,EnHt,true,true);
    StockeLesTypes(TOBL,Valeurs);
    result := true;
  END else
  BEGIN
    TOBTMP:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
    InsertionChampSupOuv (TOBTMP,false);
    TOBDEPART := TOBLN;
    TOBTMP.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
    TOBTMP.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
    TOBTMP.PutValue('BLO_DATEPIECE',TOBL.GetValue('GL_DATEPIECE')) ;
    TOBTMP.PutValue('BLO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE')) ;
    TOBTMP.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
    TOBTMP.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
    TOBTMP.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
    TOBTMP.PutValue('BLO_AVENANT',TOBL.GetValue('GL_AVENANT')) ;
    TOBTMP.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
    TOBTMP.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
    TOBTMP.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
    TOBTMP.PutValue('BLO_QTEDUDETAIL',1) ;
    TOBTMP.PutValue('BLO_COMPOSE',RefUnique) ;
    InitTableau (Valeurs);
    result := Entree_LigneOuv(XX,nil,TOBLN,TobArticles,TobL,TOBRepart,TobMetres,False,1,0,DEV,TOBL.Getvalue('GL_QUALIFQTEVTE'),valeurs,tamodif,EnHt,false,ModeSaisieAch,false) ;
    if not result then TOBLN.free;
    if (TypeArt='OUV') then TOBL.putvalue('GL_LIBELLE', Libelle);
  END ;
  if result then
  begin
    if (Libelle <> TOBL.getValue('GL_LIBELLE')) and (TOBL.GetValue('GLC_FROMBORDEREAU')<>'X') then
    begin
      TOBL.putValue('GL_LIBELLE',libelle);
    end;
    if TobNomen <> nil then TOBNomen.Free ;
    // Modif Metre => fv
    Result := false;
    //if VenteAchat = 'VEN' then
    //begin
    //  if (TheMetreShare <> nil) and (TheMetreShare.AutorisationMetre(TOBL.GetValue('GL_NATUREPIECEG'))) then
    //  begin
    //    result := TheMetreShare.MetreSousDetail(TOBL);
    //  end;
    //end;

    if result then
    begin
      //VaChercherQuantitatifDetail (TObDepart);
      InitTableau (Valeurs);
      CalculeOuvrageDoc (TOBDEPART,1,1,true,DEV,valeurs,EnHt);
      StockeLesTypes(TOBL,Valeurs);
    end;
    if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      TOBDEPART.putvalue ('BLO_DPA',valeurs[0]);
      TOBDEPART.putvalue ('BLO_DPR',valeurs[1]);
      TOBDEPART.PutValue('BLO_PMAP',valeurs[6]);
      TOBDEPART.PutValue('BLO_PMRP',valeurs[7]);
      TOBDEPART.putvalue ('BLO_PUHTDEV',valeurs[2]);
      TOBDEPART.putvalue ('BLO_PUTTCDEV',valeurs[3]);
      TOBDEPART.putvalue ('BLO_PUHT',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
      TOBDEPART.putvalue ('BLO_PUTTC',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
      TOBDEPART.putvalue ('BLO_PUHTBASE',TOBDEPART.GetValue('BLO_PUHT'));
      TOBDEPART.putvalue ('BLO_PUTTCBASE',TOBDEPART.getValue('BLO_PUTTC'));
      TOBDEPART.putvalue ('BLO_TPSUNITAIRE',valeurs[9]);
      TOBDEPART.Putvalue ('ANCPA',TOBDEPART.GetValue('BLO_DPA')) ;
      TOBDEPART.Putvalue ('ANCPR',TOBDEPART.GetValue('BLO_DPR')) ;
      //
      TOBDEPART.PutValue('GA_HEURE',valeurs[9]);
      TOBDEPART.PutValue('BLO_MONTANTPAFG',valeurs[10]);
      TOBDEPART.PutValue('BLO_MONTANTPAFR',valeurs[11]);
      TOBDEPART.PutValue('BLO_MONTANTPAFC',valeurs[12]);
      TOBDEPART.PutValue('BLO_MONTANTFG',valeurs[13]);
      TOBDEPART.PutValue('BLO_MONTANTFR',valeurs[14]);
      TOBDEPART.PutValue('BLO_MONTANTFC',valeurs[15]);
      TOBDEPART.PutValue('BLO_MONTANTPA',valeurs[16]);
      TOBDEPART.PutValue('BLO_MONTANTPR',valeurs[17]);
      //

      if EnHt then TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUHTDEV'))
              else TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUTTCDEV'));
      CalculMontantHtDevLigOuv (TOBDEPART,DEV);
    end else
    begin
      GestionDetailPrixPose (TOBDepart);
    end;
    IndiceNomen:=TOBOuvrage.Detail.Count ;
    TOBL.PutValue('GL_INDICENOMEN',IndiceNomen) ;
    if RecupPrix = 'PUH' Then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
      TOBL.Putvalue('GL_PUHTNETDEV',valeurs[2]);
      TOBL.Putvalue('GL_PUTTCNETDEV',valeurs[3]);
    end else if RecupPrix = 'DPR' then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[1]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[1]);
      TOBL.Putvalue('GL_PUHTNETDEV',valeurs[1]);
      TOBL.Putvalue('GL_PUTTCNETDEV',valeurs[1]);
    end else if (RecupPrix = 'DPA') or (RecupPrix = 'PAS')  Then
    begin
      TOBL.Putvalue('GL_PUHTDEV',valeurs[0]);
      TOBL.Putvalue('GL_PUTTCDEV',valeurs[0]);
      TOBL.Putvalue('GL_PUHTNETDEV',valeurs[0]);
      TOBL.Putvalue('GL_PUTTCNETDEV',valeurs[0]);
    end;
    TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
    TOBL.Putvalue('GL_PUHTNET',TOBL.getValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCNET',TOBL.GetValue('GL_PUTTC'));
    TOBL.Putvalue('GL_DPA',valeurs[0]);
    TOBL.Putvalue('GL_DPR',valeurs[1]);
    TOBL.Putvalue('GL_PMAP',valeurs[6]);
    TOBL.Putvalue('GL_PMRP',valeurs[7]);
    TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
    //
    TOBL.PutValue('GL_MONTANTPAFG',valeurs[10]);
    TOBL.PutValue('GL_MONTANTPAFR',valeurs[11]);
    TOBL.PutValue('GL_MONTANTPAFC',valeurs[12]);
    TOBL.PutValue('GL_MONTANTFG',valeurs[13]);
    TOBL.PutValue('GL_MONTANTFR',valeurs[14]);
    TOBL.PutValue('GL_MONTANTFC',valeurs[15]);
    TOBL.PutValue('GL_MONTANTPA',valeurs[16]);
    TOBL.PutValue('GL_MONTANTPR',valeurs[17]);
    //

    if TOBL.getValue('GL_TYPEARTICLE') <> 'ARP' then TOBL.Putvalue('GL_TYPEPRESENT', Affichage)
                                                else TOBL.Putvalue('GL_TYPEPRESENT', 0);
    arrondiOk := (TOBPiece.GetValue('GP_ARRONDILIGNE')='X');
    //   SommationAchatDoc (TOBPiece,TOBL,False);
    if (PrxVenteSto <> TOBL.GetValue('GL_PUHTDEV')) and (FromBordereau) then
    begin
      ValPr := TOBL.GetValue('GL_DPR');
      TOBL.PutValue('GL_BLOQUETARIF','-');
      TOBL.PutValue('GL_RECALCULER','X');
      TFFacture(XX).AppliquePrixOuvrage(TOBL,PrxVenteSto);
      TOBL.PutValue('GL_BLOQUETARIF','X');
    end;
    //
    TOBDEPART := TOBOUvrage.detail[IndiceNomen-1];
    NumeroteLigneOuv (TOBDEPART,TOBL,1,1,0,0,0);
    //
    (*
    positionneCoefMarge (TOBL);
    ZeroLigneMontant (TOBL);
    *)
    result := true;
  end;
END ;

procedure RecupValoDetailHrs (TOBRef: TOB;CalculPv : boolean);
var TOBPrestation : TOB;
    PVHT : double;
    QQ : TQuery;
    Req : String;
begin
if TOBRef = nil then Exit;
if TOBRef.getValue('BLO_TYPEARTICLE') <> 'ARP' then exit;
TOBPrestation := TOB.Create ('ARTICLE',nil,-1);
Req := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+TOBRef.detail[1].GetValue('BLO_ARTICLE')+'"';
TRY
   QQ := OpenSql (req,true,-1, '', True);
   TOBPrestation.selectDb ('',QQ);
   if not QQ.eof then
      begin
      if TOBPrestation.GetValue ('GA_ARTICLE') <> '' then
         begin
         if TOBPrestation.getValue ('GA_PRIXPOURQTE')=0 then TOBPrestation.PutValue('GA_PRIXPOURQTE',1);
      //
         PVHT := TOBPrestation.GetValue('GA_PVHT')/ TOBPrestation.getValue('GA_PRIXPOURQTE');
      // Rajustement au prix pour qte de l'article de base
         PVHT := PVHT * TOBRef.GetValue('BLO_PRIXPOURQTE');

      // Rcupration de la prestation associ
         TOBRef.putvalue('TGA_PRESTATION',TobPrestation.GetValue('GA_LIBELLE')) ;
         TOBRef.putvalue('GA_NATUREPRES',TOBPrestation.GetValue('GA_CODEARTICLE')) ;
         TOBRef.putvalue('GA_ARTICLE',TOBPrestation.GetValue('GA_ARTICLE')) ;
         TOBRef.putvalue('GA_HEURE',TOBRef.detail[1].GetValue('BLO_QTEFACT')) ;
         TOBRef.PutValue('GF_PRIXUNITAIRE',TOBRef.detail[1].getValue('BLO_PUHTDEV'));
         TOBRef.Putvalue ('BLO_TPSUNITAIRE',TOBRef.detail[1].GetValue('BLO_TPSUNITAIRE')) ;
         end;
      end else
      begin
         TOBRef.putvalue('TGA_PRESTATION','') ;
         TOBRef.putvalue('GA_NATUREPRES','') ;
         TOBRef.putvalue('GA_ARTICLE','') ;
         TOBRef.putvalue('GA_HEURE',TOBRef.detail[1].GetValue('BLO_QTEFACT')) ;
         TOBRef.PutValue('GF_PRIXUNITAIRE',TOBRef.detail[1].getValue('BLO_PUHTDEV'));
         TOBRef.Putvalue ('BLO_TPSUNITAIRE',TOBRef.detail[1].GetValue('BLO_TPSUNITAIRE')) ;
      end;
FINALLY
    ferme (QQ);
    TOBPrestation.free;
END;
end;

procedure CalculOuvFromDetail (TOBOuv : TOB;DEV : RDevise ; var Valeur: T_Valeurs);
var TOBDetOuv: TOB;
    Indice,IndPou : integer;
    valloc : T_Valeurs;
    Qte,QteDuPv,QteP : double;
begin
IndPou := -1;
InitTableau (Valeur);
for Indice:= 0 to TOBOuv.detail.count -1 do
    begin
  	QteP := TOBOuv.getValue('BLO_QTEFACT');
    TobDetOuv := TOBOuv.detail[Indice];
    if TobDetOuv.GetValue('CODENOMEN') <> '' then
       begin
       CalculOuvFromDetail (TOBDetOuv,DEV,valloc);
       TOBDetOuv.Putvalue ('BLO_DPA',valloc[0]);
       TOBDetOuv.Putvalue ('BLO_DPR',valloc[1]);
       TOBDetOuv.Putvalue ('BLO_PUHT',Valloc[2]);
       TOBDetOuv.Putvalue ('BLO_PUHTBASE',Valloc[2]);
       TOBDetOuv.Putvalue ('BLO_PUHTDEV',pivottodevise(Valloc[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
       TOBDetOuv.Putvalue ('BLO_PUTTC',Valloc[3]);
       TOBDetOuv.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valloc[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
       TOBDetOuv.Putvalue ('BLO_PMAP',Valloc[6]);
       TOBDetOuv.Putvalue ('BLO_PMRP',VAlloc[7]);
       TOBDetOuv.PutValue('BLO_TPSUNITAIRE',ValLoc[9]);
       TOBDetOuv.PutValue('GA_HEURE',ValLoc[9]);
       //
       TOBDetOuv.PutValue('BLO_MONTANTFG',Arrondi(ValLoc[13]*Qte,4));
       TOBDetOuv.PutValue('BLO_MONTANTFC',Arrondi(ValLoc[15]*Qte,4));
       TOBDetOuv.PutValue('BLO_MONTANTFR',Arrondi(ValLoc[14]*Qte,4));
       StockeLesTypes (TOBDetOuv,Valloc);
  		 CalculeLigneAcOuvCumul (TOBDETOUV);
       CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
       end;
    Qte := TOBDetOuv.getValue('BLO_QTEFACT');
  	if TOBDetOuv.GetValue('BLO_QTEDUDETAIL') = 0 then TOBDetOuv.PutValue('BLO_QTEDUDETAIL',1) ;
    QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE') * TOBDetOuv.getValue('BLO_QTEDUDETAIL');
    if QteDuPv = 0 then QteDuPv := 1;
    CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
    if TobDetOuv.Getvalue('BLO_TYPEARTICLE') <> 'POU' then
       begin
       valeur[0] := valeur [0] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA')),V_PGI.okdecP);
       valeur[1] := valeur [1] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR')),V_PGI.okdecP);
       valeur[2] := valeur [2] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT')),V_PGI.okdecP);
       valeur[3] := valeur [3] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC')),V_PGI.okdecP);
       valeur[6] := valeur [6] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP')),V_PGI.okdecP);
       valeur[7] := valeur [7] +Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP')),V_PGI.okdecP);
        //
        if ((Pos(TobDetOuv.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0)  and
        		(not IsLigneExternalise(TobDetOuv)) ) or (TobDetOuv.GetValue('BLO_TYPEARTICLE')='OUV') then
        begin
          valeur[9] := valeur[9] + Arrondi(((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_TPSUNITAIRE')),V_PGI.OkDecQ);
        end;
        valeur[13] := valeur[13]+Arrondi(TOBDetOuv.GetValue('BLO_MONTANTFG')/QteDuPv,4);
        valeur[14] := valeur[14]+Arrondi(TOBDetOuv.GetValue('BLO_MONTANTFR')/QteDuPv,4);
        valeur[15] := valeur[15]+Arrondi(TOBDetOuv.GetValue('BLO_MONTANTFC')/QteDuPv,4);
        CumuleLesTypes (TobDetOuv,valeur);
       end else
       begin
       indpou := 1;
       end;
    end;
FormatageTableau (valeur,V_PGI.OkdecP);
// calcule des eventuels pourcentages
if indpou >= 0 then
   begin
   for Indice := 0 to TOBOuv.detail.count -1 do
       begin
       TOBDetOuv := TOBOuv.detail[indice];
       if TOBDetOuv.Getvalue('BLO_TYPEARTICLE') <> 'POU' then continue;
       TOBDetOuv.Putvalue ('BLO_DPA',valeur[0]);
       TOBDetOuv.Putvalue ('BLO_DPR',valeur[1]);
       TOBDetOuv.Putvalue ('BLO_PUHT',valeur[2]);
       TOBDetOuv.Putvalue ('BLO_PUHTBASE',valeur[2]);
       TOBDetOuv.Putvalue ('BLO_PUHTDEV',pivottodevise(valeur[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
       TOBDetOuv.Putvalue ('BLO_PUTTC',valeur[3]);
       TOBDetOuv.Putvalue ('BLO_PUTTCDEV',pivottodevise(valeur[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
       TOBDetOuv.Putvalue ('BLO_PMAP',valeur[6]);
       TOBDetOuv.Putvalue ('BLO_PMRP',valeur[7]);
       Qte := TOBDetOuv.getValue('BLO_QTEFACT');
       CalculMontantHtDevLigOuv (TOBDetOuv,DEV);
       QteDuPv := TOBDetOuv.getValue('BLO_PRIXPOURQTE'); if QteDuPv = 0 then QteDuPv := 1;
       valeur[0] := valeur [0] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPA'),V_PGI.okdecP);
       valeur[1] := valeur [1] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_DPR'),V_PGI.okdecP);
       valeur[2] := valeur [2] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUHT'),V_PGI.okdecP);
       valeur[3] := valeur [3] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PUTTC'),V_PGI.okdecP);
       valeur[6] := valeur [6] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMAP'),V_PGI.okdecP);
       valeur[7] := valeur [7] + Arrondi((Qte/QteDuPv) * TOBDetOuv.GetValue ('BLO_PMRP'),V_PGI.okdecP);
       FormatageTableau (valeur,V_PGI.OkdecP );
       end;
   end;
end;

procedure GetValoDetail (TOBOuv:TOB);
var TypeArticle : string;
begin
  TypeArticle:= TOBOuv.GetValue('BLO_TYPEARTICLE');

  if TypeArticle = 'PRE' Then
  begin
   TOBOuv.putvalue('GA_HEURE',TOBOUV.GetValue('BLO_QTEFACT')) ;
//   if (RenvoieTypeRes(TOBOuv.GetValue('BLO_ARTICLE')) <> 'ST') then
(*
   if (TOBOuv.GetValue('BNP_TYPERESSOURCE') <> 'ST') then
   begin
   	TOBOuv.putvalue('TPSUNITAIRE',1);
   end else
   begin
   	TOBOuv.putvalue('TPSUNITAIRE',0);
   end;
*)
   //TOBOuv.PutValue('GF_PRIXUNITAIRE',TOBOuv.getValue('BLO_PUHTDEV'));
  end;

  if (TypeArticle <> 'ARP') then exit;
  if TobOuv.detail.count > 0 then
  begin
    TOBOuv.PutValue ('GCA_PRIXBASE',TOBOUV.Detail[0].GetValue('GCA_PRIXBASE'));
    TOBOuv.PutValue ('TGA_FOURNPRINC',TOBOUV.Detail[0].GetValue('TGA_FOURNPRINC'));
    TOBOuv.PutValue ('GA_FOURNPRINC',TOBOUV.Detail[0].GetValue('GA_FOURNPRINC'));
    TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBOUV.Detail[0].GetValue('GCA_QUALIFUNITEACH'));
    TOBOuv.PutValue ('GF_CALCULREMISE',TOBOUV.Detail[0].GetValue('GF_CALCULREMISE'));
    TOBOuv.PutValue ('GA_PAHT',TOBOUV.Detail[0].GetValue('GA_PAHT'));
    TOBOuv.PutValue ('GA_PVHT',TOBOUV.Detail[0].GetValue('GA_PVHT'));
(* OPTIMIZATION *)
    TOBOuv.PutValue ('BLO_FOURNISSEUR',TOBOUV.Detail[0].GetValue('GA_FOURNPRINC'));
    TOBOuv.PutValue ('BLO_REMISES',TOBOUV.Detail[0].GetValue('GF_CALCULREMISE'));
    TOBOuv.PutValue ('BLO_QUALIFQTEACH',TOBOUV.Detail[0].GetValue('GCA_QUALIFUNITEACH'));
		TOBOuv.Putvalue ('BLO_NOMENCLATURE',TOBOUV.Detail[0].GetValue('BLO_NOMENCLATURE')) ;
    TOBOuv.PutValue ('BLO_PRXACHBASE',TOBOUV.Detail[0].GetValue('GCA_PRIXBASE'));
(* ------------ *)
  end;
  if TOBOUV.detail.count > 1 then
  begin
   // dans le cas ou la prestation n'a pas ete defini
   TOBOuv.putvalue('TGA_PRESTATION',TOBOUV.Detail[1].GetValue('BLO_LIBELLE')) ;
   TOBOuv.putvalue('GA_NATUREPRES',TOBOUV.Detail[1].GetValue('BLO_CODEARTICLE')) ;
   TOBOuv.putvalue('GA_ARTICLE',TOBOUV.Detail[1].GetValue('BLO_ARTICLE')) ;
//   if (RenvoieTypeRes(TOBOUV.Detail[1].GetValue('BLO_ARTICLE')) <> 'ST') then
(*
   if (TOBOUV.Detail[1].GetValue('BNP_TYPERESSOURCE') <> 'ST') then
   begin
   	 TOBOuv.putvalue('BLO_TPSUNITAIRE',TOBOUV.Detail[1].GetValue('BLO_QTEFACT')) ;
   end else
   begin
   	 TOBOuv.putvalue('BLO_TPSUNITAIRE',0) ;
   end;
*)
   TOBOuv.putvalue('GA_HEURE',TOBOUV.Detail[1].GetValue('BLO_QTEFACT')) ;
   TOBOuv.PutValue('GF_PRIXUNITAIRE',TOBOUV.Detail[1].getValue('BLO_PUHTDEV'));
  end;
end;



procedure InitValoDetail (TOBPiece,TOBOuv,TOBRef : TOB; DEV : Rdevise;CalculPv:boolean=true; RecupPrix : string='PUH');
var  TOBCatalog,TOBTarif,TOBTiers : TOB;
     QQ : TQuery;
     MTPAF : double;
     TypeArticle :string;
     Fournisseur : string;
     PrixPourQte : double;
     IsUniteAchat : boolean;
begin
IsUniteAchat := true; // par dfaut
TypeArticle:= TOBOuv.GetValue('BLO_TYPEARTICLE');
Fournisseur := TOBOuv.GetValue('BLO_FOURNISSEUR');
if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') and (TypeArticle <> 'FRA') then exit;
if not TOBRef.FieldExists ('QUALIFQTEACH') then TOBREF.AddChampSupValeur ('QUALIFQTEACH','',false);
if not TOBRef.FieldExists ('PRXACHBASE') then TOBREF.AddChampSupValeur ('PRXACHBASE',0,false);
if not TOBRef.FieldExists ('REMISES') then TOBREF.AddChampSupValeur ('REMISES','',false);
if not TOBRef.FieldExists ('DEJA CALCULE') then TOBREF.AddChampSupValeur ('DEJA CALCULE','-',false);

TOBTarif := TOB.Create ('TARIF',nil,-1);
TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
TOBTiers := TOB.Create ('TIERS',nil,-1);
TRY
   if Fournisseur <> '' then
   begin
     QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+fournisseur+'"',true,-1, '', True);
     TOBTiers.SelectDB ('',QQ);
     ferme (QQ);
     // ON ne récupère pas les tarifs si on est en mode .. je recup le pmap
     if (VH_GC.ModeValoPa <> 'PMA') then
     begin
       chargeTOBs (TOBRef,TOBCatalog,TOBTiers,TOBTarif);
       if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
       begin
        MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
        TOBOuv.PutValue ('GCA_PRIXBASE',arrondi(MTPAF,V_PGI.okdecP));
       end else
       begin
        PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
        if PrixPourQte = 0 then PrixPourQte := 1;
        if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
                                                    else MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
        TOBOuv.PutValue ('GCA_PRIXBASE',arrondi(MTPAF,V_PGI.okdecP));
       end;
       if MTPAF = 0 then
       begin
         if TOBRef.GetValue ('DEJA CALCULE')<>'X' then
         begin
           MTPAF := TobRef.getValue('GA_PAHT');
           TOBOuv.PutValue ('GCA_PRIXBASE',arrondi(MTPAF,V_PGI.okdecP));
         end;
         IsUniteAchat := false;
       end;
       MTPAF := arrondi(TOBOuv.GetValue ('GCA_PRIXBASE') * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
       // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
       if IsUniteAchat then MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
       //
       if (MtPaf <> 0) and (mtpaf <> TobRef.getValue('GA_PAHT')) then
       begin
          TOBREF.putValue('GA_PAHT',MTPAF);
          RecalculPrPV (TOBRef,TOBCatalog);
       end;
     end;
     TOBOuv.PutValue ('BLO_QUALIFQTEACH',TOBCatalog.getValue('GCA_QUALIFUNITEACH'));
     TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBCatalog.getValue('GCA_QUALIFUNITEACH'));
     TOBOuv.PutValue ('GF_CALCULREMISE',TOBTarif.GetValue('GF_CALCULREMISE'));
   end else
   begin
      if TOBRef.GetString('GA_QUALIFUNITEACH')<> '' then
      begin
        TOBOuv.PutValue ('BLO_QUALIFQTEACH',TOBRef.getValue('GA_QUALIFUNITEACH'));
        TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBRef.getValue('GA_QUALIFUNITEACH'));
        TOBOuv.PutValue ('GCA_PRIXBASE',TOBRef.getValue('GA_PAUA'));
   end;
   end;
   TOBOuv.PutValue ('TGA_FOURNPRINC',TOBTIers.GetValue('T_LIBELLE'));
   TOBOuv.PutValue ('GA_FOURNPRINC',TOBTIers.GetValue('T_TIERS'));
//   TOBOuv.PutValue ('GA_PAHT',TOBREf.GetValue('GA_PAHT'));
   if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBRef.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
   begin
    if TOBREf.GetValue('GA_PMAP')<>0 then
    begin
   		TOBOuv.PutValue ('GA_PAHT',TOBREf.GetValue('GA_PMAP'));
    end else
    begin
   		TOBOuv.PutValue ('GA_PAHT',TOBREf.GetValue('GA_PAHT'));
    end;
   end else
   begin
   	TOBOuv.PutValue ('GA_PAHT',TOBREf.GetValue('GA_PAHT'));
   end;
   (* OPTIMIZATION *)
   TOBOuv.PutValue ('BLO_FOURNISSEUR',TOBTIers.GetValue('T_TIERS'));
   TOBOuv.PutValue ('BLO_REMISES',TOBTarif.GetValue('GF_CALCULREMISE'));
   TOBOuv.PutValue ('BLO_PRXACHBASE',TOBOuv.GetValue('GCA_PRIXBASE'));
   TOBOuv.PutValue ('BLO_COEFCONVQTE',TOBCatalog.getValue('GCA_COEFCONVQTEACH'));
   TOBOuv.PutValue ('BLO_COEFCONVQTEVTE',TOBREF.getValue('GA_COEFCONVQTEVTE'));
   if RecupPrix = 'PUH' then
   begin
     TOBOuv.PutValue ('BLO_COEFMARG',TOBRef.GetValue('GA_COEFCALCHT'));
     TOBOuv.PutValue('POURCENTMARG',Arrondi((TOBOuv.GetValue('BLO_COEFMARG')-1)*100,2));

     (* ------------- *)
     if CalculPv then
        begin
        if TypeArticle <> 'PRE' Then TOBOuv.PutValue ('GA_PVHT',TOBRef.GetValue('GA_PVHT'));
        TOBOUv.putvalue ('BLO_PUHT',TOBRef.GetValue('GA_PVHT'));
        TOBOUv.putvalue ('BLO_PUTTC',TOBRef.GetValue('GA_PVTTC'));
        TOBOuv.PutValue ('BLO_PUHTDEV',pivottodevise(TobOuv.GetValue('BLO_PUHT'),DEV.Taux,DEV.quotite,V_PGI.okdecP ));
        TOBOuv.PutValue ('BLO_PUTTCDEV',pivottodevise(TobOuv.GetValue('BLO_PUTTC'),DEV.Taux,DEV.quotite,V_PGI.OKdecP));
        end;
      if TOBOUV.GetValue('BLO_PUHT') <> 0 then
      begin
        TOBOuv.PutValue('POURCENTMARQ',Arrondi(((TOBOUV.GetValue('BLO_PUHT')- TOBOUV.GetValue('BLO_DPR'))/TOBOUV.GetValue('BLO_PUHT'))*100,2));
      end else
      begin
        TOBOuv.PutValue('POURCENTMARQ',0);
      end;
     // Mise en place des valorisations dans la tob Ouvrage
     TOBOUv.putvalue ('BLO_DPA',TOBOUv.GetValue('GA_PAHT'));
     TOBOUv.putvalue ('BLO_DPR',TOBRef.GetValue('GA_DPR'));
     if TypeArticle = 'PRE' Then
     begin
       TOBOuv.putvalue('GA_PVHT',TOBRef.GetValue('GA_PVHT')) ;
       TOBOuv.putvalue('GA_HEURE',TOBOUV.GetValue('BLO_QTEFACT')) ;
       TOBOUv.putvalue ('GF_PRIXUNITAIRE',TOBRef.GetValue('GA_PVHT'));
     end;
     // Modif LS OPTIMISATION
     if TOBRef.GetValue('GA_DPRAUTO')='X' then
     begin
      if TobRef.getValue('GA_COEFFG') <> 0 then TOBOUv.putvalue ('BLO_COEFFG',TobRef.getValue('GA_COEFFG')-1);
     end else
     begin
      if TOBOUV.GetValue('BLO_DPA') = 0 then
      begin
        TOBOUv.putvalue ('BLO_DPA',TOBOUv.getvalue ('BLO_DPR'));
        TOBOUv.putvalue ('BLO_COEFFG',0);
      end else
      begin
  //   		TOBOUv.putvalue ('BLO_COEFFG',Arrondi(TOBOUV.GetValue('BLO_DPR')/TOBOUV.GetValue('BLO_DPA'),4)-1);
        TOBOUv.putvalue ('BLO_COEFFG',0);
      end;
     end;
     // --
     if TOBPiece <> nil then
     begin
       // if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),'SAL;INT;')>0) then
       if IsPrestationSt(TOBOUV) then
       begin
        if TOBPiece.GetValue('GP_APPLICFGST')<>'X' then TOBOUV.PutValue('BLO_NONAPPLICFRAIS','X');
        if TOBPiece.GetValue('GP_APPLICFCST') <> 'X' then TOBOUV.PutValue('BLO_NONAPPLICFC','X');
       end;
     end;
   end else  if RecupPrix = 'DPR' then
   begin
     TOBOUv.putvalue ('BLO_DPA',TOBOUv.GetValue('GA_PAHT'));
     TOBOUv.putvalue ('BLO_DPR',TOBRef.GetValue('GA_DPR'));
     // Modif LS OPTIMISATION
     if TOBRef.GetValue('GA_DPRAUTO')='X' then
     begin
      if TobRef.getValue('GA_COEFFG') <> 0 then TOBOUv.putvalue ('BLO_COEFFG',TobRef.getValue('GA_COEFFG')-1);
     end else
     begin
      if TOBOUV.GetValue('BLO_DPA') = 0 then
      begin
        TOBOUv.putvalue ('BLO_DPA',TOBOUv.getvalue ('BLO_DPR'));
        TOBOUv.putvalue ('BLO_COEFFG',0);
      end else
      begin
  //   		TOBOUv.putvalue ('BLO_COEFFG',Arrondi(TOBOUV.GetValue('BLO_DPR')/TOBOUV.GetValue('BLO_DPA'),4)-1);
        TOBOUv.putvalue ('BLO_COEFFG',0);
      end;
     end;
     // --
     if TOBPiece <> nil then
     begin
       // if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),'SAL;INT;')>0) then
       if IsPrestationSt(TOBOUV) then
       begin
        if TOBPiece.GetValue('GP_APPLICFGST')<>'X' then TOBOUV.PutValue('BLO_NONAPPLICFRAIS','X');
        if TOBPiece.GetValue('GP_APPLICFCST') <> 'X' then TOBOUV.PutValue('BLO_NONAPPLICFC','X');
       end;
     end;

   end else if ((RecupPrix='PAS') or (RecupPrix = 'DPA')) then
   begin
     TOBOUv.putvalue ('BLO_DPA',TOBOUv.GetValue('GA_PAHT'));
     TOBOUv.putvalue ('BLO_DPR',TOBRef.GetValue('GA_PAHT'));
     TOBOUv.putvalue ('BLO_COEFFG',0);
     TOBOUv.putvalue ('BLO_COEFMARG',0);
     if TypeArticle = 'PRE' Then
     begin
       TOBOuv.putvalue('GA_PVHT',TOBRef.GetValue('GA_PVHT')) ;
       TOBOuv.putvalue('GA_HEURE',TOBOUV.GetValue('BLO_QTEFACT')) ;
       TOBOUv.putvalue ('GF_PRIXUNITAIRE',TOBRef.GetValue('GA_PVHT'));
     end;
      TOBOUV.PutValue('BLO_NONAPPLICFRAIS','X');
      TOBOUV.PutValue('BLO_NONAPPLICFC','X');
   end;

FINALLY
   TOBTarif.free;
   TOBCatalog.free;
   TOBTiers.free;
end;
end;

procedure positionneEltTarif (TOBOUV,TOBREF : TOB);
begin
TOBREF.PutValue ('GA_PAHT',TOBOUV.GetValue('BLO_DPA'));
TOBREF.PutValue ('GA_DPR',TOBOUV.GetValue('BLO_DPR'));
TOBREF.PutValue ('GA_PVHT',TOBOUV.GetValue('BLO_PUHT'));
//TOBREF.PutValue ('GA_FOURNPRINC',TOBOUV.GetValue('BLO_FOURNISSEUR'));
end;

procedure RecupValorisationOUV (TOBOuv,TOBRef : TOB; OptimizeOuv : ToptimizeOuv);
var TypeArticle,Fournisseur : string;
		TOBTiers : TOB;
begin
  TypeArticle:= TOBOuv.GetValue('BLO_TYPEARTICLE');
  Fournisseur := TOBOuv.GetValue('BLO_FOURNISSEUR');
  if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') then exit;
  // initialisation
  TOBTiers := OptimizeOuv.TOBFournisseur.FindFirst (['T_TIERS'],[Fournisseur],true);

  TOBOuv.PutValue ('GCA_PRIXBASE',TOBOuv.getValue('BLO_PRXACHBASE'));
  if TOBTiers <> nil then TOBOuv.PutValue ('TGA_FOURNPRINC',TOBTiers.GetValue('T_LIBELLE'));
  TOBOuv.PutValue ('GA_FOURNPRINC',Fournisseur);
  TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBOuv.getValue('BLO_QUALIFQTEACH'));
  TOBOuv.PutValue ('GF_CALCULREMISE',TOBOuv.getValue('BLO_REMISES'));
  TOBOuv.PutValue ('GA_PAHT',TOBOUV.GetValue('BLO_DPA'));
  TOBOuv.PutValue ('GA_PVHT',TOBOUV.GetValue('BLO_PUHTDEV'));
  if (TypeArticle='PRE') then
  begin
    //recup des elements Main oeuvre
    TOBOUV.putvalue('TGA_PRESTATION',TobRef.GetValue('GA_LIBELLE')) ;
    TOBOUV.putvalue('GA_NATUREPRES',TOBRef.GetValue('GA_CODEARTICLE')) ;
    TOBOUV.putvalue('GA_ARTICLE',TOBRef.GetValue('GA_ARTICLE')) ;
    TOBOUV.putvalue('GA_HEURE',TOBOuv.GetValue('BLO_QTEFACT')) ;
    TOBOUV.PutValue('GF_PRIXUNITAIRE',TOBOuv.getValue('BLO_PUHTDEV'));
(*    if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),'SAL;INT;')>0) then
    begin
    	TOBOUV.putvalue('TPSUNITAIRE',1) ;
    end; *)
  end;
end;

procedure RecupValoDetail (TOBOuv,TOBRef : TOB);
var  TOBCatalog,TOBTarif,TOBTiers : TOB;
     QQ : TQuery;
     MTPAF : double;
     TypeArticle :string;
     Fournisseur : string;
     PrixPourQte : double;
     IsUniteAchat : boolean;

begin
IsUniteAchat := true;
TypeArticle:= TOBOuv.GetValue('BLO_TYPEARTICLE');
Fournisseur := TOBOuv.GetValue('BLO_FOURNISSEUR');
if (TypeArticle<> 'MAR') and (TypeArticle <> 'ARP') and (TypeArticle <> 'PRE') then exit;
// initialisation
TOBTarif := TOB.Create ('TARIF',nil,-1);
TOBCatalog := TOB.Create ('CATALOGU',nil,-1);
TOBTiers := TOB.Create ('TIERS',nil,-1);
TRY

   QQ := OpenSql ('SELECT * FROM TIERS WHERE T_TIERS="'+fournisseur+'"',true,-1, '', True);
   TOBTiers.SelectDB ('',QQ);
   ferme (QQ);

   positionneEltTarif (TOBOUV,TOBREF);
   chargeTOBs (TOBRef,TOBCatalog,TOBTiers,TOBTarif);

   if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
   begin
      MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
   end else
   begin
      PrixPourQte := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
      if PrixPourQte = 0 then PrixPourQte := 1;
      if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then MTPAF :=TOBCatalog.GetValue('GCA_PRIXBASE')/PrixPourQte
                                                  else MTPAF :=TOBCatalog.GetValue('GCA_PRIXVENTE')/PrixPourQte;
   end;
   if MTPaf = 0 Then BEGIN MTPAF := TOBRef.GetValue('GA_PAHT'); IsUniteAchat := false; END;
   TOBOuv.PutValue ('GCA_PRIXBASE',MTPAF);
   MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
   // voila voila voila ..le seul hic c'est que ce prix est en UA..donc passage de l'UA en UV
   if IsUniteAchat then MTPAF := PassageUAUV (TOBRef,TOBCatalog,MTPAF);
   //
   TOBOuv.PutValue ('TGA_FOURNPRINC',TOBTIers.GetValue('T_LIBELLE'));
   TOBOuv.PutValue ('GA_FOURNPRINC',TOBTIers.GetValue('T_TIERS'));
   if TOBRef.GetString('GA_QUALIFUNITEACH')<> '' then
   begin
      TOBOuv.PutValue ('BLO_QUALIFQTEACH',TOBRef.getValue('GA_QUALIFUNITEACH'));
      TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBRef.getValue('GA_QUALIFUNITEACH'));
      TOBOuv.PutValue ('GCA_PRIXBASE',TOBRef.getValue('GA_PAUA'));
   end else
   begin
   TOBOuv.PutValue ('GCA_QUALIFUNITEACH',TOBCatalog.getValue('GCA_QUALIFUNITEACH'));
   end;
   TOBOuv.PutValue ('GF_CALCULREMISE',TOBTarif.GetValue('GF_CALCULREMISE'));
(* OPTIMIZATION *)
   TOBOuv.PutValue ('BLO_FOURNISSEUR',TOBTIers.GetValue('T_TIERS'));
   TOBOuv.PutValue ('BLO_REMISES',TOBTarif.GetValue('GF_CALCULREMISE'));
   TOBOuv.PutValue ('BLO_PRXACHBASE',MTPAF);
(* ------------ *)
   TOBOuv.PutValue ('GA_PAHT',TOBOUV.GetValue('BLO_DPA'));
   TOBOuv.PutValue ('GA_PVHT',TOBOUV.GetValue('BLO_PUHTDEV'));
   if (TypeArticle='PRE') then
      begin
      //recup des elements Main oeuvre
      TOBOUV.putvalue('TGA_PRESTATION',TobRef.GetValue('GA_LIBELLE')) ;
      TOBOUV.putvalue('GA_NATUREPRES',TOBRef.GetValue('GA_CODEARTICLE')) ;
      TOBOUV.putvalue('GA_ARTICLE',TOBRef.GetValue('GA_ARTICLE')) ;
      TOBOUV.putvalue('GA_HEURE',TOBOuv.GetValue('BLO_QTEFACT')) ;
      TOBOUV.PutValue('GF_PRIXUNITAIRE',TOBOuv.getValue('BLO_PUHTDEV'));
(*      if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),'SAL;INT;')>0) then
      begin
      	TOBOUV.putvalue('TPSUNITAIRE',1) ;
      end; *)
      end;
FINALLY
   TOBTarif.free;
   TOBCatalog.free;
   TOBTiers.free;
end;
end;

procedure TraiteRecupValoGrp ( TOBOuvrage,TOBArticles : TOB; var Valeur : T_Valeurs);
var TOBOuv : TOB;
    Indice : integer;
    TOBRef : TOB;
    TypeArticle: string;
    ValOuv : T_Valeurs;
begin
TOBRef := TOB.Create ('ARTICLE',nil,-1);
TRY
for Indice := 0 to TOBOuvrage.detail.count -1 do
begin
    TOBOuv := TOBOuvrage.detail[Indice];
    TypeArticle:= TOBOuv.GetValue('BLO_TYPEARTICLE');
    if ((TypeArticle = 'ARP') and (TOBOUV.Detail.count > 0)) or (TypeArticle = 'PRE') then
    begin
       GetValoDetail (TOBOUv);
       if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and
       	  (not IsLigneExternalise(TOBOuv)) then
       begin
       	Valeur[9] := Valeur[9] + (TOBOuv.getValue('BLO_TPSUNITAIRE')*TOBOuv.getValue('BLO_QTEFACT'));
       end;
    end;
    if (TypeArticle='OUV') then
    begin
       TraiteRecupValoGrp (TOBOUV,TOBArticles,ValOuv);
       if (Pos(TOBOuv.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and
       	  (not IsLigneExternalise(TOBOuv)) then
       begin
       		Valeur[9] := Valeur[9] + (TOBOuv.getValue('BLO_TPSUNITAIRE')*TOBOuv.getValue('BLO_QTEFACT'));
       end;
    end;
end;
FINALLY
   TOBRef.free;
end;

end;

Procedure LoadLesOuvrages ( TOBPiece,TOBOuvrage,TOBArticles : TOB ; CleDoc : R_CleDoc ;Var IndiceNomen : integer; OptimizeOuv : TOptimizeOuv = nil; Forprint : boolean=false);
Var i,OldL,Lig,II : integer ;
    TOBL,TOBLN,TOBNomen,TOBGroupeNomen,TOBPP : TOB ;
    OkN  : boolean ;
    Q    : TQuery ;
    valeurs : T_Valeurs;
    Requete : string;
    WithLigneFac,IsAvanc : boolean;
    DEV : RDevise;
BEGIN
	DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  IsAvanc := (Pos(RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS')),'AVA;DAC;')>0);

  //
  (* OPTIMIZATION *)
  if (OptimizeOuv <> nil) and (not Forprint) then
  begin
  	OptimizeOuv.initialisation;
  	OptimizeOuv.ChargeLesTOBS (Cledoc);
  end;
  (* ------------ *)
  if not NaturepieceOKPourOuvrage (TOBPiece) then exit;
  WithLigneFac := (pos(cledoc.NaturePiece,'FBT;DAC;FBP;BAC')>0);

  OkN:=False ; OldL:=-1 ;
  for i:=0 to TOBPiece.Detail.Count-1 do
  BEGIN
    TOBL:=TOBPiece.Detail[i] ; TOBL.PutValue('GL_INDICENOMEN',0) ;
    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') then OkN:=True ;
  END ;
  if Not OkN then Exit ;

  TOBNomen:=TOB.Create('',Nil,-1) ;
  (*
  requete :='SELECT O.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE,'+
  					'(SELECT T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="FOU" AND T_TIERS=BLO_FOURNISSEUR) AS LIBELLEFOU FROM LIGNEOUV O '+
            'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=O.BLO_ARTICLE) '+
            'WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
            ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';
  *)
  requete := MakeSelectLigneOuvBtp (WithLigneFac,TOBpiece.getString('GP_NATUREPIECEG'));
  requete := requete + ' WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
            					 ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';

//
  Q:=OpenSQL(requete,True,-1, '', True) ;
  TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
  Ferme(Q) ;
  if TOBNomen.Detail.Count=0 then Exit;
  i := 0; OldL := 0;
  repeat
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if OldL<>Lig then
    BEGIN
      TOBGroupeNomen:=TOB.Create('',TOBOuvrage,-1) ;
      TOBGroupeNomen.AddChampSup('UTILISE',False) ;
//      TOBGroupeNomen.PutValue('UTILISE','X') ;
      TOBGroupeNomen.PutValue('UTILISE','-') ;
      TOBGroupeNomen.AddChampSupValeur ('AVANCSOUSDET','-') ;
      IndiceNomen := TobOuvrage.detail.count;
      TOBl := TOBpiece.detail[Lig-1];
//      TobL := TobPiece.FindFirst(['GL_NUMLIGNE'], [Lig], False);
      if (TOBL <> nil) and (IsOuvrage(TOBL)) then
      begin
        TOBL.PutValue('GL_INDICENOMEN',IndiceNomen);
      end;
      (* OPTIMIZATION *)
      CreerLesTOBOuv(TOBL,TOBGroupeNomen,TOBNomen,TOBArticles,Lig,i,OPtimizeOuv,ForPrint,true,false,WithLigneFac);
      (* ------------ *)
      if (TOBL <> nil) and (TOBL.GetValue('GL_TYPEARTICLE')='ARP') then
      begin
        TOBPP := TOBGroupeNomen.detail[0];
        if TOBPP.detail.count > 0 then
        begin
          TOBPP.PutValue ('GCA_PRIXBASE',TOBPP.Detail[0].GetValue('GCA_PRIXBASE'));
          TOBPP.PutValue ('TGA_FOURNPRINC',TOBPP.Detail[0].GetValue('TGA_FOURNPRINC'));
          TOBPP.PutValue ('GA_FOURNPRINC',TOBPP.Detail[0].GetValue('GA_FOURNPRINC'));
          TOBPP.PutValue ('GCA_QUALIFUNITEACH',TOBPP.Detail[0].GetValue('GCA_QUALIFUNITEACH'));
          TOBPP.PutValue ('GF_CALCULREMISE',TOBPP.Detail[0].GetValue('GF_CALCULREMISE'));
          TOBPP.PutValue ('GA_PAHT',TOBPP.Detail[0].GetValue('GA_PAHT'));
          TOBPP.PutValue ('GA_PVHT',TOBPP.Detail[0].GetValue('GA_PVHT'));
        (* OPTIMIZATION *)
          TOBPP.PutValue ('BLO_FOURNISSEUR',TOBPP.Detail[0].GetValue('GA_FOURNPRINC'));
          TOBPP.PutValue ('BLO_REMISES',TOBPP.Detail[0].GetValue('GF_CALCULREMISE'));
          TOBPP.PutValue ('BLO_QUALIFQTEACH',TOBPP.Detail[0].GetValue('GCA_QUALIFUNITEACH'));
          TOBPP.Putvalue ('BLO_NOMENCLATURE',TOBPP.Detail[0].GetValue('BLO_NOMENCLATURE')) ;
          TOBPP.PutValue ('BLO_PRXACHBASE',TOBPP.Detail[0].GetValue('GCA_PRIXBASE'));
        (* ------------ *)
        end;

        if TOBPP.detail.count > 1 then
        begin
          TOBPP.putvalue('GA_ARTICLE',TOBPP.detail[1].getValue('GA_ARTICLE'));
          TOBPP.putvalue('GA_NATUREPRES',TOBPP.detail[1].getValue('GA_NATUREPRES'));
          TOBPP.putvalue('GA_HEURE',TOBPP.Detail[1].GetValue('BLO_QTEFACT')) ;
          TOBPP.PutValue('GF_PRIXUNITAIRE',TOBPP.Detail[1].getValue('BLO_PUHTDEV'));
        end;
      end;
      OldL := Lig;
    END ;
  until i>=TOBNomen.detail.count ;
  TOBNomen.Free ;
  if (WithLigneFac) and (IsAvanc) then
  begin
    For I := 0 to TOBOuvrage.detail.count -1 do
    begin
    	TOBLN:=TOBOuvrage.Detail[i] ;
      for II := 0 to TOBLN.detail.count - 1 do
      begin
    		Lig:=TOBLN.detail[II].GetValue('BLO_NUMLIGNE') ;
      	TOBl := TOBpiece.detail[Lig-1];
        //
      	SetAvancementOuvrage(TOBLN,TOBLN.detail[II],TOBL,DEV);
      end;
    end;
  end;
//
END ;

Procedure ValideLesOuv ( TOBOuv : TOB ; TOBPiece : TOB = nil) ;
Var i,II,IBlonumligne,iIndiceNomen : integer ;
    TOBL,TOBGO,TOBO : TOB;
    TOBLL : TOB;
BEGIN
  IBlonumligne := -1;
  iIndiceNomen := -1;
  if TOBOUV = nil then exit;
  if TOBPiece = nil then exit;
  i := 0;
  if TOBOuv.detail.count = 0 then exit;
  repeat
    TOBGO := TOBOuv.Detail[i];
  	if TOBGO.GetValue('UTILISE')<>'X' then
    begin
  		TOBGO.Free ;
    end else
    begin
      if TOBGO.detail.count = 0 then BEGIN inc(I); continue; END;
  		if TOBPiece <> nil then
      begin
        TOBO := TOBGO.detail[0];
        if IBlonumligne = -1 then
        begin
          IBlonumligne := TOBO.GetNumChamp('BLO_NUMLIGNE');
        end;

        if TOBO <> nil then
        begin
          TOBL := GetTOBLIgne (TOBPiece,TOBO.GetInteger(IBlonumligne));
          if iIndiceNomen = -1 then
          begin
            iIndiceNomen := TOBL.GetNumChamp('GL_INDICENOMEN');
          end;
          if TOBL <> nil then TOBL.SetInteger(iIndiceNomen,I+1);
        end;
      end;
      Inc(i);
    end;
  until i >= TOBOuv.Detail.Count;

  if Not TOBOuv.InsertDBByNivel(false) then
  begin
    MessageValid := 'Erreur écriture des ouvrages';
    PgiError(MessageValid);
    V_PGI.IoError:=oeUnknown ;
  end;

END ;

Procedure MajLigOuv (TOBN,TOBLig : TOB) ;
Var i : integer ;
    TOBL : TOB ;
BEGIN
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    TOBL.PutValue('BLO_NATUREPIECEG',TOBLig.GetValue('GL_NATUREPIECEG')) ;
    TOBL.PutValue('BLO_SOUCHE',TOBLig.GetValue('GL_SOUCHE')) ;
    TOBL.PutValue('BLO_NUMERO',TOBLig.GetValue('GL_NUMERO')) ;
    TOBL.PutValue('BLO_INDICEG',TOBLig.GetValue('GL_INDICEG')) ;
    TOBL.PutValue('BLO_NUMLIGNE',TOBLig.GetValue('GL_NUMLIGNE')) ;
    // Protection coef
    if TOBL.GetValue('BLO_COEFMARG')>99999 then TOBL.putValue('BLO_COEFMARG',0); // protection du coef de marge
    if TOBL.GetValue('BLO_COEFFG')>99999 then TOBL.putValue('BLO_COEFG',0); // protection du coef de frais généraux
    if TOBL.GetValue('BLO_COEFFC')>99999 then TOBL.putValue('BLO_COEFC',0); // protection du coef de frais généraux
    if TOBL.GetValue('BLO_COEFFR')>99999 then TOBL.putValue('BLO_COEFR',0); // protection du coef de frais répartis
    // --
    MajLigOuv(TOBL,TOBLig) ;
    END ;
END ;

Procedure NumeroteLigneOuv (TOBN,TOBLig : TOB; NiveauCur,LigPere1,LigPere2, LigPere3,LigPere4: integer ) ;
Var i : integer ;
    TOBL : TOB ;
    LigNiv1,LigNiv2,LigNiv3,LigNiv4,LigNiv5,LigPers1,LigPerS2,LigPerS3,LigPerS4 : integer;
    PLigniveauCur : ^Integer;
    NiveauSuiv : integer;
BEGIN
if NiveauCur = 1 then
   begin
   LigNiv1 := 1;
   LigNiv2 := 0;
   LigNIv3 := 0;
   LigNIv4 := 0;
   LigNiv5 := 0;
   PLigniveauCur := @LigNiv1;
   end else if NiveauCur = 2 then
       begin
       LigNiv1 := LigPere1;
       LigNiv2 := 1;
       LigNIv3 := 0;
       LigNIv4 := 0;
       LigNiv5 := 0;
       PLigniveauCur := @LigNiv2;
       end else if NiveauCur = 3 then
           begin
           LigNiv1 := LigPere1;
           LigNiv2 := LigPere2;
           LigNIv3 := 1;
           LigNIv4 := 0;
           LigNiv5 := 0;
           PLigniveauCur := @LigNiv3;
           end else if NiveauCur = 4 then
               begin
               LigNiv1 := LigPere1;
               LigNiv2 := LigPere2;
               LigNiv3 := LigPere3;
               LigNIv4 := 1;
               LigNiv5 := 0;
               PLigniveauCur := @LigNiv4;
               end else BEGin
                        LigNiv1 := LigPere1;
                        LigNiv2 := LigPere2;
                        LigNiv3 := LigPere3;
                        LigNiv4 := LigPere4;
                        LigNiv5 := 1;
                        PLigniveauCur := @LigNiv5;
                        END;
NiveauSuiv := NiveauCur +1;
for i:=0 to TOBN.Detail.Count-1 do
    BEGIN
    TOBL:=TOBN.Detail[i] ;
    TOBL.PutValue('BLO_N1',LigNiv1);
    TOBL.PutValue('BLO_N2',LigNiv2);
    TOBL.PutValue('BLO_N3',LigNiv3);
    TOBL.PutValue('BLO_N4',LigNiv4);
    TOBL.PutValue('BLO_N5',LigNiv5);
    TOBL.PutValue('BLO_NIVEAU',NiveauCur);
    LigPerS1 := LigNiv1;
    LigPerS2 := LigNiv2;
    LigPerS3 := LigNiv3;
    LigPerS4 := LigNiv4;
    inc (PLigniveauCur^);
    NumeroteLigneOuv(TOBL,TOBLig,NiveauSuiv,LigPerS1,LigPerS2,LigPerS3,LigPerS4) ;
    END ;
END ;

function ReAffecteLigOuv (IndiceNomen : integer ; TOBLig,TOBOuvrage : TOB ) : boolean ;
Var TOBN : TOB ;
BEGIN

result := true;

if TOBOuvrage=Nil then Exit ;

//FV1 : 28/07/2016
if TobOuvrage.detail.count = 0 then exit;

if IndiceNomen<=0 then Exit ;

if TOBOuvrage.Detail.Count-1 < IndiceNomen-1 then   //?????
begin
  result := false;
  Exit ;
end;

TOBN:=TOBOuvrage.Detail[IndiceNomen-1] ; TOBN.PutValue('UTILISE','X') ;
NumeroteLigneOuv (TOBN,TOBlig,1,1,0,0,0);
MajLigOuv(TOBN,TOBLig) ;
END ;

Procedure CreerLesTOBOuv (TOBL,TOBGroupeNomen,TOBNomen,TOBArticles : TOB ; LaLig,idep : integer;OptimizeOuv : TOptimizeOuv = nil;AvecStock : boolean=true;WithValo : boolean=true; Forprint : boolean=false; WithLigneFac : boolean=false ) ;

  function Findpere (TOBGroupeNomen: TOB;Ligne1,Ligne2,Ligne3,Ligne4 : integer;niveau : integer) : TOB;
  begin
    if Niveau = 5 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2','BLO_N3','BLO_N4'],[Ligne1,Ligne2,Ligne3,Ligne4],true)
    else if Niveau = 4 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2','BLO_N3'],[Ligne1,Ligne2,Ligne3],true)
    else if Niveau = 3 then result := TOBGroupeNomen.findfirst(['BLO_N1','BLO_N2'],[Ligne1,Ligne2],true)
    else if Niveau = 2 then result := TOBGroupeNomen.findfirst(['BLO_N1'],[Ligne1],true)
    else result := nil;
    if (result = nil) and (Niveau > 1) then
    begin
      result := FindPere (TOBGroupeNomen,Ligne1,Ligne2,Ligne3,Ligne4,Niveau-1);
    end;
  end;

Var i,Lig,II : integer ;
    TOBLN,TOBPere,TOBLoc,TOBArt,TOBRef : TOB ;
    RefArticle,TypeArticle : String ;
    DEV:Rdevise;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
    IndN1,IndN2,IndN3,IndN4,IndN5,IndNumLigne,IndNart : integer;
    TOBpere4,TOBpere3,TOBpere2,TOBpere1 : TOB;
    LaTOBArticles : TOB;
    QQ : TQuery;
    QteDuDetail : double;
    valeurs : T_Valeurs;
BEGIN
(* OPTIMIZATION *)
	if (not ForPrint) then
  begin
		if OPtimizeOuv <> nil then LaTOBArticles := OPtimizeOuv.TOBArticles
                        	else laTobArticles := TOBArticles;
  end;
  (* ------------- *)
  // OPTIMISATIONS LS
//  for i:=idep to TOBNomen.Detail.Count-1 do
  i := 0;
  //
  IndNumLigne := TOBNomen.detail[i].GetNumChamp('BLO_NUMLIGNE') ;
  IndN1 := TOBNomen.detail[0].GetNumChamp('BLO_N1');
  IndN2 := TOBNomen.detail[0].GetNumChamp('BLO_N2');
  IndN3 := TOBNomen.detail[0].GetNumChamp('BLO_N3');
  IndN4 := TOBNomen.detail[0].GetNumChamp('BLO_N4');
  IndN5 := TOBNomen.detail[0].GetNumChamp('BLO_N5');
  IndNArt := TOBNomen.detail[0].GetNumChamp('BLO_CODEARTICLE');
  //
  repeat
    TOBLN:=TOBNomen.Detail[i] ;
    RefArticle := TOBLN.GetValeur(IndNart);
    //
    Lig:=TOBLN.GetValeur(IndNumLigne) ;
    if lig <> LaLig then break;
    LigneN1 := TOBLN.GetValeur(IndN1);
    LigneN2 := TOBLN.GetValeur(indN2);
    LigneN3 := TOBLN.GetValeur(IndN3);
    LigneN4 := TOBLN.GetValeur(IndN4);
    LigneN5 := TOBLN.GetValeur(IndN5);
    //
    if LigneN2 = 0 then
    begin
    	TOBPere:=TOBGroupeNomen;
      // Stockage du pere de niveau 1
      TOBpere1 := TOBLN;
    end else if LigneN3 = 0 then
    begin
      // Stockage du pere de niveau 2
      TOBpere := TOBpere1;
      TOBpere2 := TOBLN;
    end else if LigneN4 = 0 then
    begin
      TOBpere := TOBpere2;
      // Stockage du pere de niveau 3
      TOBpere3 := TOBLN;
    end else if LigneN5=0 then
    begin
      TOBpere := TOBpere3;
      // Stockage du pere de niveau 4
      TOBpere4 := TOBLN;
    end else TOBpere := TOBPere4;
    //
    if TOBPere = nil then
    begin
      if LigneN2 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,2)
      else if LigneN3 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,3)
      else if LigneN4 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,4)
      else if LigneN5 = 0 then Findpere (TOBGroupeNomen,LigneN1,LigneN2,LigneN3,LigneN4,5);
      if TOBPere = nil then TOBpere := TOBGroupeNomen;
    end;
    if (TOBPere<>Nil) {and (TOBLN.getValue('BLO_ARTICLE')<>'')} then
    BEGIN
      // OPTIMISATIONS LS
      InsertionChampSupOuv (TOBLN,false);
      TOBLoc := TOBLN;
      //
      if OptimizeOuv <> nil then
      begin
        if (TOBLN.GetDouble('BLO_MONTANTPA') <> 0)  then
        begin
          if (TOBLN.GetString('BLO_TYPEARTICLE')='MAR') and (TOBLN.GetDouble('BLO_MTFOUPA')=0) then
          begin
            OptimizeOuv.Arecalc := True;
            StockeMontantTypeSurLigne (TOBLN);
          end else if (TOBLN.GetString('BLO_TYPEARTICLE')='PRE') then
          begin
            if (TOBLN.GetDouble('BLO_MTMOPA')+TOBLN.GetDouble('BLO_MTSTPA')+
                TOBLN.GetDouble('BLO_MTINTPA')+TOBLN.GetDouble('BLO_MTAUTPA')+
                TOBLN.GetDouble('BLO_MTLOCPA')+TOBLN.GetDouble('BLO_MTMATPA')+
                TOBLN.GetDouble('BLO_MTOUTPA')  =0) then
            begin
              StockeMontantTypeSurLigne (TOBLN);
              OptimizeOuv.Arecalc := True;
            end;
          end;
        end;
      end;
      if TOBL.FieldExists('BLF_POURCENTAVANC') and (TOBLN.FieldExists('BLF_POURCENTAVANC')) and
         (TOBL.getValue('BLF_POURCENTAVANC')<> TOBLN.Getvalue('BLF_POURCENTAVANC')) and
         (TOBLN.Getvalue('BLF_POURCENTAVANC')<>0) AND ((TOBL.Getvalue('BLF_POURCENTAVANC')<>0)) and
         (TOBLN.Getvalue('BLO_N2')=0)then
      begin
        TOBGroupeNomen.putValue('AVANCSOUSDET','X');
      end;
      //
      TOBLoc.ChangeParent (TOBpere,-1);
      DEV.Code:=TOBLOC.GetValue('BLO_DEVISE') ; GetInfosDevise(DEV) ;
      CalculMontantHtDevLigOuv (TOBLOC,DEV);
      RefArticle:=TOBLoc.GetValue('BLO_ARTICLE') ;
      TypeArticle := TOBLoc.GetValue('BLO_TYPEARTICLE');
      if (not ForPrint) then
      begin
        (* OPTIMIZATION *)
        TOBArt:=LaTOBArticles.FindFirst(['GA_ARTICLE'],[RefArticle],False) ;
        (* ------------ *)
        if TOBArt=Nil then
        BEGIN
          (* OPTIMZATION *)
          (* -------- *)
          TOBArt:=CreerTOBArt(LaTOBArticles) ;
          QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
          'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
          'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+RefArticle+'"',true,-1, '', True);
          if not QQ.eof then TOBArt.SelectDB('',QQ)
                        else InitArticleFromLigneOuv (TOBLoc,TOBArt);

          Ferme (QQ);
          InitChampsSupArticle (TOBART);
          if AvecStock then LoadTOBDispo(TOBArt,False) ;
        END ;
        if WithValo then
        begin
          if (TypeArticle = 'MAR') or (TypeArticle='PRE') or ((TypeArticle='ARP') and (TOBLoc.Detail.count = 0)) then
          begin
            TOBRef := TOB.Create ('ARTICLE',nil,-1);
            TOBRef.dupliquer (TOBART,false,true);
            TRY
              (* OPTIMIZATION *)
              (* ------------- *)
              IF OptimizeOuv = nil then RecupValoDetail (TOBLoc,TOBRef)
                                   else RecupValorisationOUV (TOBLoc,TOBRef,OptimizeOuv);
            FINALLY
              TOBREF.free;
            END;
          end;
          TOBLoc.PutValue('GA_HEURE',TOBLoc.getValue('BLO_TPSUNITAIRE'));

          TOBLoc.Putvalue('GA_PAHT',TOBLOC.getvalue('BLO_DPA'));
          TOBLoc.Putvalue('GA_PVHT',TOBLOC.getvalue('BLO_PUHTDEV'));
          TOBLoc.Putvalue('ANCPV',TOBLOC.getvalue('BLO_PUHTDEV'));
          TOBLoc.Putvalue('ANCPA',TOBLOC.getvalue('BLO_DPA'));
          TOBLoc.Putvalue('ANCPR',TOBLOC.getvalue('BLO_DPR'));
          if TOBPere.FieldExists('BLO_TYPEARTICLE') and (TOBPere.GetValue('BLO_TYPEARTICLE')='ARP') and (TOBLoc.GetValue('BLO_TYPEARTICLE')='PRE') then
          begin
             TOBPere.putvalue('GA_NATUREPRES',TOBLoc.GetValue('BLO_CODEARTICLE')) ;
             TOBPere.putvalue('GA_HEURE',TOBLoc.GetValue('BLO_QTEFACT')) ;
             TOBPere.PutValue('GF_PRIXUNITAIRE',TOBLoc.getValue('BLO_PUHTDEV'));
             QteDudetail := TOBLoc.getValue('BLO_PRIXPOURQTE'); if QteDudetail = 0 then QteduDetail := 1;
//             TOBPere.Putvalue ('BLO_TPSUNITAIRE',TOBLoc.GetValue('BLO_TPSUNITAIRE')*(TOBLoc.getValue('BLO_QTEFACT')/QteDuDetail)) ;
          end;
          if TOBPere.FieldExists('BLO_TYPEARTICLE') and (TOBPere.GetValue('BLO_TYPEARTICLE')='ARP') and (TOBLoc.GetValue('BLO_TYPEARTICLE')='ARP') then
          begin
            TOBPere.Putvalue('GA_PAHT',TOBLOC.getvalue('GA_PAHT'));
            TOBPere.Putvalue('GA_PVHT',TOBLOC.getvalue('GA_PVHT'));
          end;
        end;
    	end;
      if TOBPere.FieldExists('BLO_TPSUNITAIRE') then
      begin
      	TOBPere.Putvalue ('GA_HEURE',TOBPere.GetValue('BLO_TPSUNITAIRE')) ;
      end;
      //
    END else
    begin
      TOBLN.free;
    end;
  until i>=TOBnomen.detail.count;
  //FV1 - 25/04/2018 : FS#3084 - ESPACS : Violation d'accès sur certaines factures en Saisie de réglements
  if (OptimizeOuv <> nil) And (OptimizeOuv.Arecalc) then
  begin
    for II := 0 to TOBgroupenomen.detail.count -1 do
    begin
      if TOBGroupeNomen.detail[II].detail.count > 0 then
        CalculOuvFromDetail (TOBGroupeNomen.detail[II],DEV,valeurs);
    end;
  end;

END ;

procedure CalculeOuvrageDoc (TOBOuvrage : TOB;Qte,QteDetail:double;recursif:boolean;DEV:RDevise ; var valeurs:T_Valeurs;EnHt : boolean;AvecHt:boolean=true;WithArrondi : boolean=false);
var Indice,IndPou : Integer;
    TOBL : TOB;
    QteLoc,QteDudetail,QteDupv,QteL,QTEDEt : Double;
    Valloc,ValPou,MtPou : T_Valeurs;
    ArticleOk : string;
    Article : string;
    PPQ,X : Double;
begin
  if TOBOuvrage = nil then exit;
  InitTableau (Valpou);
  InitTableau (Mtpou);
  ArticleOk := '';
  IndPou := -1 ;
  TOBL:=TOBOuvrage.findfirst(['BLO_TYPEARTICLE'],['POU'],false);
  if TOBL <> nil then ArticleOk := TOBL.GetValue('BLO_LIBCOMPL');
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBL := TOBOuvrage.detail [Indice];
    Article := TOBL.GetValue('BLO_CODEARTICLE');
    // VARIANTE
    if isVariante (TOBL) then continue;
    // --
    if TOBL.GetValue('BLO_TYPEARTICLE') = 'POU' then
    begin
      IndPou := indice;
      continue;
    end;
    QteDuPv := TOBL.GetValue('BLO_PRIXPOURQTE');
    if QteDuPV = 0 then QteDuPv := 1;

    if Qtedetail <= 0 then Qtedetail := 1;
    QteLoc := Qte * TOBL.GetValue('BLO_QTEFACT');
    QTEDEt := TOBL.getDouble('BLO_QTEDUDETAIL'); if QteDet = 0 then QteDet := 1;
    QteL := TOBL.GetValue('BLO_QTEFACT') / QteDet;
    if (TOBL.detail.count > 0 ) and (recursif) then
    begin
      //      Qtedudetail := Qtedetail * TOBL.GetValue('BLO_QTEDUDETAIL');
      Qtedudetail := 1;
      InitTableau (Valloc);
      CalculeOuvrageDoc (TOBL,1,Qtedudetail,true,DEV,valloc,EnHt,AvecHt);
      TOBL.PutValue ('GA_PAHT',valloc[0]);
      TOBL.putvalue ('BLO_DPA',valloc[0]);
      TOBL.putvalue ('BLO_DPR',valloc[1]);
      TOBL.PutValue('BLO_PMAP',ValLoc[6]);
      TOBL.PutValue('BLO_PMRP',ValLoc[7]);
      // --
      TOBL.PutValue('BLO_TPSUNITAIRE',ValLoc[9]);
      TOBL.PutValue('GA_HEURE',ValLoc[9]);
      TOBL.PutValue('BLO_MONTANTPAFG',Arrondi(ValLoc[10] * QteL,V_PGI.OKDecV) );
      TOBL.PutValue('BLO_MONTANTPAFR',Arrondi(ValLoc[11] * QteL,V_PGI.OKDecV));
      TOBL.PutValue('BLO_MONTANTPAFC',Arrondi(ValLoc[12] * QteL,V_PGI.OKDecV));
      TOBL.PutValue('BLO_MONTANTFG',  Arrondi(ValLoc[13] * QteL,V_PGI.OKDeCV));
      TOBL.PutValue('BLO_MONTANTFR',  Arrondi(ValLoc[14] * QteL,V_PGI.OKDecV));
      TOBL.PutValue('BLO_MONTANTFC',  Arrondi(ValLoc[15] * QteL,V_PGI.OKDecV));
      TOBL.PutValue('BLO_MONTANTPA',  Arrondi(ValLoc[16] * QteL,V_PGI.OKDecV));
      TOBL.PutValue('BLO_MONTANTPR',  Arrondi(ValLoc[17] * QteL,V_PGI.OKDecV));
      if (AvecHt) then
      begin
        TOBL.PutValue('GA_PVHT',valloc[2]);
        TOBL.putvalue ('BLO_PUHTDEV',valloc[2]);
        TOBL.putvalue ('BLO_PUTTCDEV',valloc[3]);
        TOBL.putvalue ('BLO_PUHT',devisetopivotEx(valloc[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
        TOBL.putvalue ('BLO_PUTTC',devisetopivotEx(valloc[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
        TOBL.putvalue ('BLO_PUHTBASE',TOBL.GetValue('BLO_PUHT'));
        TOBL.putvalue ('BLO_PUTTCBASE',TOBL.getValue('BLO_PUTTC'));
        TOBL.Putvalue('ANCPA',TOBL.Getvalue('BLO_DPA'));
        TOBL.Putvalue('ANCPR',TOBL.Getvalue('BLO_DPR'));
        if EnHt then TOBL.Putvalue('ANCPV',TOBL.Getvalue('BLO_PUHTDEV'))
                else TOBL.Putvalue('ANCPV',TOBL.Getvalue('BLO_PUTTCDEV'));
        CalculMontantHtDevLigOuv (TOBL,DEV);
      end;
      //
      StockeLesTypes (TOBL,ValLoc);
      //
      GetValoDetail (TOBL); // ajouts LS pour els articles en prix posés
      if Valloc[1] <> 0 then
      begin
        TOBL.PutValue('BLO_COEFMARG',arrondi(ValLoc[2]/Valloc[1],4));
        TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('BLO_COEFMARG')-1)*100,2));
    end;
      if TOBL.GetValue('BLO_PUHT') <> 0 then
      begin
        TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
      end else
      begin
        TOBL.PutValue('POURCENTMARQ',0);
      end;
    end;
    if TOBL.GetValue('BLO_QTEDUDETAIL') = 0 then TOBL.PutValue('BLO_QTEDUDETAIL',1) ;
    Qtedudetail := TOBL.GetValue('BLO_QTEDUDETAIL');
    // fv 02032004
    if QteDudetail = 0 then  QteDuDetail := 1;
    if WithArrondi then
    begin
      Valeurs[0] := Valeurs[0] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPA'))),V_PGI.OkDecP);
      Valeurs[1] := Valeurs[1] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPR'))),V_PGI.OkDecP);
      Valeurs[6] := Valeurs[6] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMAP'))),V_PGI.OkDecP);
      Valeurs[7] := Valeurs[7] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMRP'))),V_PGI.OkDecP);
      if ((Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and (not IsLigneExternalise(TOBL))) or
      (TOBL.GetValue('BLO_TYPEARTICLE')='OUV') then
      begin
        Valeurs[9] := Valeurs[9] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_TPSUNITAIRE'))),V_PGI.OkDecQ);
      end;
      Valeurs[10] := Valeurs[10] + Arrondi(TOBL.GetValue ('BLO_MONTANTPAFG'),V_PGI.OkDecP);
      Valeurs[11] := Valeurs[11] + Arrondi(TOBL.GetValue ('BLO_MONTANTPAFR'),V_PGI.OkDecP);
      Valeurs[12] := Valeurs[12] + Arrondi(TOBL.GetValue ('BLO_MONTANTPAFC'),V_PGI.OkDecP);
      Valeurs[13] := Valeurs[13] + Arrondi(TOBL.GetValue ('BLO_MONTANTFG'),V_PGI.OkDecP);
      Valeurs[14] := Valeurs[14] + Arrondi(TOBL.GetValue ('BLO_MONTANTFR'),V_PGI.OkDecP);
      Valeurs[15] := Valeurs[15] + Arrondi(TOBL.GetValue ('BLO_MONTANTFC'),V_PGI.OkDecP);
      Valeurs[16] := Valeurs[16] + Arrondi(TOBL.GetValue ('BLO_MONTANTPA'),V_PGI.OkDecP);
      Valeurs[17] := Valeurs[17] + Arrondi(TOBL.GetValue ('BLO_MONTANTPR'),V_PGI.OkDecP);
    end else
    begin
      Valeurs[0] := Valeurs[0] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPA')));
      Valeurs[1] := Valeurs[1] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPR')));
      Valeurs[6] := Valeurs[6] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMAP')));
      Valeurs[7] := Valeurs[7] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMRP')));
      if ((Pos(TOBL.GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0)  and (not IsLigneExternalise(TOBL))) or
          (TOBL.GetValue('BLO_TYPEARTICLE')='OUV') then
      begin
        X:=RatioMesure('TEM',TOBL.GetValue('BLO_QUALIFQTEFACT')); // pour revenir a l'unite de base 
        Valeurs[9] := Valeurs[9] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_TPSUNITAIRE') * X));
      end;
      Valeurs[10] := Valeurs[10] + TOBL.GetValue ('BLO_MONTANTPAFG');
      Valeurs[11] := Valeurs[11] + TOBL.GetValue ('BLO_MONTANTPAFR');
      Valeurs[12] := Valeurs[12] + TOBL.GetValue ('BLO_MONTANTPAFC');
      Valeurs[13] := Valeurs[13] + TOBL.GetValue ('BLO_MONTANTFG');
      Valeurs[14] := Valeurs[14] + TOBL.GetValue ('BLO_MONTANTFR');
      Valeurs[15] := Valeurs[15] + TOBL.GetValue ('BLO_MONTANTFC');
      Valeurs[16] := Valeurs[16] + TOBL.GetValue ('BLO_MONTANTPA');
      Valeurs[17] := Valeurs[17] + TOBL.GetValue ('BLO_MONTANTPR');
    end;
    if AvecHt then
    begin
      if WithArrondi then
      begin
        Valeurs[2] := Valeurs[2] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail*QteDuPv) * TOBL.GetValue ('BLO_PUHTDEV'))),V_PGI.OKdecP);
        Valeurs[3] := Valeurs[3] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PUTTCDEV'))),V_PGI.OKdecP);
      end else
      begin
        Valeurs[2] := Valeurs[2] + ((QteLoc/(Qtedetail*QteDuDetail*QteDuPv) * TOBL.GetValue ('BLO_PUHTDEV')));
        Valeurs[3] := Valeurs[3] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PUTTCDEV')));
      end;
    end;
    // Pour les totaux en fonction du type (MAT,ST,MO,etc..)
    CumuleLesTypes (TOBL,Valeurs);
    //
    if ArticleOKInPOUR (TOBL.GetValue('BLO_TYPEARTICLE'),ArticleOK) then
    begin
      if WIthArrondi then
      begin
        ValPou[0] := ValPou[0] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPA'))),V_PGI.okdecP);
        ValPou[1] := ValPou[1] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPR'))),V_PGI.okdecP);
        ValPou[6] := ValPou[6] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMAP'))),V_PGI.okdecP);
        ValPou[7] := ValPou[7] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMRP'))),V_PGI.okdecP);
        if AvecHt then
        begin
          ValPou[2] := ValPou[2] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail*QteDuPv) * TOBL.GetValue ('BLO_PUHTDEV'))),V_PGI.okdecP);
          ValPou[3] := ValPou[3] + Arrondi(((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PUTTCDEV'))),V_PGI.okdecP);
        end;
      end else
      begin
        ValPou[0] := ValPou[0] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPA')));
        ValPou[1] := ValPou[1] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_DPR')));
        ValPou[6] := ValPou[6] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMAP')));
        ValPou[7] := ValPou[7] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PMRP')));
        if AvecHt then
        begin
          ValPou[2] := ValPou[2] + ((QteLoc/(Qtedetail*QteDuDetail*QteDuPv) * TOBL.GetValue ('BLO_PUHTDEV')));
          ValPou[3] := ValPou[3] + ((QteLoc/(Qtedetail*QteDuDetail) * TOBL.GetValue ('BLO_PUTTCDEV')));
        end;
      end;
    	CalculMontantHtDevLigOuv (TOBL,DEV);
    end;
  end;
  if IndPou >= 0 then
  begin
    PPQ := TOBOuvrage.detail[IndPou].getDouble ('BLO_PRIXPOURQTE');
    Qte := TOBOuvrage.detail[IndPou].GetDouble ('BLO_QTEFACT');
    //
    TOBOuvrage.detail[IndPou].putvalue ('BLO_DPA',ValPou[0]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_DPR',ValPou[1]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHTDEV',ValPou[2]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTCDEV',ValPou[3]);
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHT',DevisetopivotEx(ValPou[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUHTBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUHT'));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTC',DevisetopivotEx(ValPou[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBOuvrage.detail[IndPou].putvalue ('BLO_PUTTCBASE',TOBOuvrage.detail[IndPou].Getvalue ('BLO_PUTTC'));
    TOBOuvrage.detail[IndPou].PutValue('BLO_PMAP',ValPou[6]);
    TOBOuvrage.detail[IndPou].PutValue('BLO_PMRP',ValPou[7]);
    CalculMontantHtDevLigOuv (TOBOuvrage.detail[IndPou],DEV);
    //
    Mtpou[0] := ValPou[0] * (Qte/PPQ);
    Mtpou[1] := ValPou[1] * (Qte/PPQ) ;
    Mtpou[2] := arrondi (ValPou[2] * (Qte/PPQ),V_PGI.OkDecP ) ;
    Mtpou[3] := ValPou[3] * (Qte/PPQ);
    Mtpou[6] := ValPou[6] * (Qte/PPQ);
    Mtpou[7] := ValPou[7] * (Qte/PPQ) ;
    FormatageTableau (Mtpou,V_PGI.okdecP);
    //
    Valeurs := CalculSurTableau ('+',Valeurs,MtPou);
  end;
  FormatageTableau (Valeurs,V_PGI.OkdecP);
end;

procedure AttribueEcart (TOBPiece,TOBL,TobOuvrage : TOB;ArtEcart:string;Montantecart: double;DEV:RDEVISE;EnHt:Boolean;EnPA : boolean=false;OnlyPa:boolean=false;PrixPourQte : double=0; LibEcart : string=''; ForcePaEqualPv : Boolean=false);
var TOBEcart,TOBART,TOBRef : TOB;
    Q : Tquery;
    Sql,RecupPrix : String;
    Indice : integer;
    QteDuDetail : Double;
begin
	RecupPrix := GetInfoParPiece(TOBPIece.GetValue('GP_NATUREPIECEG'), 'GPP_APPELPRIX');
  QteDuDetail := TOBOuvrage.detail[TOBOuvrage.detail.count -1].GetDouble ('BLO_QTEDUDETAIL');
  if QteDuDetail = 0 then QteDuDetail := 1;
  if PrixPourQte = 0 then prixPourQte := 1;
  TOBRef := TOBOuvrage.detail[TOBOuvrage.detail.count -1]; // On se positionne sur la derniere fille
  TOBART := TOB.Create ('ARTICLE',nil,-1);
  TobEcart:=TOB.Create('LIGNEOUV',TOBOuvrage,-1) ;
  InsertionChampSupOuv (TOBEcart,false);

  MontantEcart := Arrondi(MontantEcart,V_PGI.okdecP);
  SQl := 'SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
        'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
        'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEARTICLE="'+ArtEcart+'"';

  Q := opensql ( sql,true,-1, '', True);
  if not Q.Eof then
  begin
    TOBART.selectdb ('',Q);
    InitChampsSupArticle (TOBART);
    TobEcart.PutValue('BNP_TYPERESSOURCE',TOBArt.GetValue('BNP_TYPERESSOURCE')) ;
    TobEcart.PutValue('BLO_ARTICLE',TOBArt.GetValue('GA_ARTICLE')) ;
    TobEcart.PutValue('BLO_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE')) ;
    TobEcart.PutValue('BLO_REFARTSAISIE',TOBArt.GetValue('GA_CODEARTICLE')) ;
    CopieOuvFromArt (nil,TobEcart,TOBART,DEV,RecupPrix);
    CopieOuvFromRef (TOBEcart,TOBRef);
    TobEcart.PutValue('BLO_NUMORDRE',TOBRef.GetValue('BLO_NUMORDRE')+1) ;
    TobEcart.PutValue('BLO_N1',TOBRef.GetValue('BLO_N1')+1) ;
    TobEcart.PutValue('BLO_LIBELLE',TOBArt.GetValue('GA_LIBELLE')) ;
    if LibEcart <> '' then TobEcart.PutValue('BLO_LIBELLE',LibEcart);
    //FV1 : 17/01/2017 - FS#2282 - ALES - Divison par zéro en virgule flottante incorrecte
    If TOBEcart.GetValue('BLO_QTEDUDETAIL') = 0 then TobEcart.PutValue('BLO_QTEDUDETAIL', 1);
    //
    TobEcart.PutValue('BLO_QTEFACT',QteDuDetail) ;
    TobEcart.PutValue('BLO_COMPOSE',TOBRef.GetValue('BLO_COMPOSE')) ;
    TobEcart.PutValue('BLO_TENUESTOCK',TOBArt.GetValue('GA_TENUESTOCK')) ;
    TobEcart.PutValue('BLO_QUALIFQTEVTE',TOBArt.GetValue('GA_QUALIFUNITEVTE')) ;
    TobEcart.PutValue('BLO_DPA',0) ;
    TobEcart.PutValue('BLO_DPR',0) ;
    TobEcart.PutValue('BLO_PMAP',0) ;
    TobEcart.PutValue('BLO_PMRP',0) ;
    // Provenance ligne
    TobEcart.PutValue('BLO_NATUREPIECEG',TOBRef.GetValue('BLO_NATUREPIECEG')) ;
    TobEcart.PutValue('BLO_REGIMETAXE',TOBref.GetValue('BLO_REGIMETAXE')) ;
    TobEcart.PutValue('BLO_SOUCHE',TOBRef.GetValue('BLO_SOUCHE')) ;
    TobEcart.PutValue('BLO_NUMERO',TOBRef.GetValue('BLO_NUMERO')) ;
    TobEcart.PutValue('BLO_INDICEG',TOBRef.GetValue('BLO_INDICEG')) ;
    TobEcart.PutValue('BLO_NUMLIGNE',TOBRef.GetValue('BLO_NUMLIGNE')) ;
    TobEcart.PutValue('BLO_NIVEAU',TobRef.getValue('BLO_NIVEAU')) ;
    TobEcart.PutValue('BLO_ORDRECOMPO',TOBRef.GetValue('BLO_ORDRECOMPO')) ;
    if EnPa Then
    begin
      TobEcart.putvalue ('BLO_DPA',MontantEcart);
      TobEcart.putvalue ('BLO_DPR',MontantEcart);
    end;
    if (not Onlypa) then
    begin
	    TobEcart.putvalue ('BLO_PUHTDEV',MontantEcart);
      if ForcePaEqualPv then
      begin
	      TobEcart.putvalue ('BLO_DPA',MontantEcart);
	      TobEcart.putvalue ('BLO_DPR',MontantEcart);
	      TobEcart.putvalue ('BLO_COEFFG',0);
	      TobEcart.putvalue ('BLO_COEFFC',0);
	      TobEcart.putvalue ('BLO_COEFFR',0);
        TobEcart.SetDouble('BLO_COEFMARG',1);
      end;
    end else
    begin
	    TobEcart.putvalue ('BLO_PUHTDEV',0);
    end;
    //    TobEcart.putvalue ('BLO_PUHT',DeviseToPivotEx (MontantEcart,DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TobEcart.putvalue ('ANCPA',TOBEcart.Getvalue ('BLO_DPA'));
    TobEcart.putvalue ('ANCPR',TOBEcart.Getvalue ('BLO_DPR'));
    if EnHt then
	    TobEcart.putvalue ('ANCPV',TOBEcart.Getvalue ('BLO_PUHTDEV'))
    else
  	  TobEcart.putvalue ('ANCPV',TOBEcart.Getvalue ('BLO_PUTTCDEV'));
    TobEcart.putvalue ('BLO_TYPELIGNE','ART');
    TobEcart.putvalue ('BLO_QUALIFQTEVTE','');
    TobEcart.PutValue('BLO_PRIXPOURQTE',1);
    TobEcart.PutValue ('BLO_PRIXPOURQTE',QteDuDetail);
    if EnHt then CalculeLigneHTOuv(TobEcart, TOBPiece, DEV)
    				else CalculeLigneTTCOuv(TobEcart, TOBPiece, DEV);
    TobEcart.PutValue ('BLO_PUTTC',DeviseToPivotEx(TobEcart.GetValue('BLO_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TobEcart.putvalue ('BLO_PUHT',DeviseToPivotEx(TobEcart.GetValue('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TobEcart.putvalue ('BLO_PUHTBASE',TOBEcart.getValue('BLO_PUHT'));
    TobEcart.PutValue ('GA_PAHT',TobEcart.GetValue('BLO_DPA'));
    TobEcart.PutValue ('GA_PVHT',TobEcart.GetValue('BLO_PUHTDEV'));
    TobEcart.putvalue ('BLO_PUTTCBASE',TOBEcart.GetValue('BLO_PUTTC'));

    CalculMontantHtDevLigOuv (ToBEcart,DEV);
    TOBARt.free;
  end
  else if TOBRef.GetValue('BLO_TYPEARTICLE') = 'MAR' then
  begin
    TOBEcart.Dupliquer (TOBRef,true,true);
    TobEcart.PutValue('BLO_QTEFACT',QteDuDetail) ;
    TOBEcart.SetString('BLO_COMPOSE','');
    TobEcart.PutValue('BLO_NUMORDRE',TOBRef.GetValue('BLO_NUMORDRE')+1) ;
    TobEcart.PutValue('BLO_N1',TOBRef.GetValue('BLO_N1')+1) ;
    TobEcart.PutValue('BLO_N2',0) ;
    TobEcart.PutValue('BLO_N3',0) ;
    TobEcart.PutValue('BLO_N4',0) ;
    TobEcart.PutValue('BLO_N5',0) ;
    TobEcart.putvalue ('BLO_TYPELIGNE','ART');
    TOBEcart.putvalue ('BLO_TYPENOMENC','');
    TOBEcart.putvalue ('BLO_TYPEARTICLE','MAR');
    TobEcart.PutValue ('BLO_PRIXPOURQTE',PrixPourQte);
    TobEcart.PutValue ('BLO_QTEDUDETAIL',QteDuDetail);
    if EnPa Then
    begin
      TobEcart.putvalue ('BLO_DPA',MontantEcart);
      TobEcart.putvalue ('BLO_DPR',MontantEcart);
    end;
    if not Onlypa then
    begin
    	TobEcart.putvalue ('BLO_PUHTDEV',MontantEcart);
      if ForcePaEqualPv then
      begin
	      TobEcart.putvalue ('BLO_DPA',MontantEcart);
	      TobEcart.putvalue ('BLO_DPR',MontantEcart);
	      TobEcart.putvalue ('BLO_COEFFG',0);
	      TobEcart.putvalue ('BLO_COEFFC',0);
	      TobEcart.putvalue ('BLO_COEFFR',0);
        TobEcart.putvalue ('BLO_COEFMARG',1);
      end;
    end else
    begin
    	TobEcart.putvalue ('BLO_PUHTDEV',0);
    end;
    TobEcart.putvalue ('ANCPA',TOBEcart.Getvalue ('BLO_DPA'));
    TobEcart.putvalue ('ANCPR',TOBEcart.Getvalue ('BLO_DPR'));
    if EnHt then CalculeLigneHTOuv(TobEcart, TOBPiece, DEV)
    else CalculeLigneTTCOuv(TobEcart, TOBPiece, DEV);

    TobEcart.putvalue ('BLO_PUHT',DeviseToPivotEx(TobEcart.GetValue('BLO_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TobEcart.putvalue ('BLO_PUHTBASE',TOBEcart.getValue('BLO_PUHT'));
    TobEcart.PutValue ('BLO_PUTTC',DeviseToPivotEx(TobEcart.GetValue('BLO_PUTTCDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    if EnHt then TobEcart.putvalue ('ANCPV',TOBEcart.Getvalue ('BLO_PUHTDEV'))
    				else TobEcart.putvalue ('ANCPV',TOBEcart.Getvalue ('BLO_PUTTCDEV'));
    TobEcart.putvalue ('BLO_PUTTCBASE',TOBEcart.GetValue('BLO_PUTTC'));
    TobEcart.putvalue ('BLO_QUALIFQTEVTE','');
    TobEcart.PutValue ('GA_PAHT',TobEcart.GetValue('BLO_DPA'));
    TobEcart.PutValue ('GA_PVHT',TobEcart.GetValue('BLO_PUHTDEV'));
    //
    // Correction FQ 11984 --> tva incorrecte sur article d'ecart
    TobEcart.PutValue ('BLO_REGIMETAXE',TOBRef.GetValue('BLO_REGIMETAXE'));
    TobEcart.PutValue ('BLO_TVAENCAISSEMENT',TOBRef.GetValue('BLO_TVAENCAISSEMENT'));
    for Indice := 1 to 5 do TobEcart.PutValue ('BLO_FAMILLETAXE'+InttoStr(Indice),TOBRef.GetValue('BLO_FAMILLETAXE'+InttoStr(Indice)));
    //
    CalculMontantHtDevLigOuv (ToBEcart,DEV);
    TOBEcart.SetAllModifie (true);
  end;
  ferme (Q);
  NumeroteLigneOuv (TobOuvrage,TOBL,TOBOUvrage.GetINteger('BLO_NIVEAU')+1,TOBOUvrage.GetINteger('BLO_N1'),TOBOUvrage.GetINteger('BLO_N2'),TOBOUvrage.GetINteger('BLO_N3'),TOBOUvrage.GetINteger('BLO_N4'));

  TOBRef := nil;
end;

//-----------

procedure AffecteDifferenceSurLigne (TOBL,TOBLOC : TOB;Ecart,EcartPr:double;
																		 IndPa,IndPR,IndPVHtDev,IndPVHt : integer;
                                     EnPa,EnPR,EnHt : boolean; DEv : RDevise;ForcePaEqualPv:Boolean=false);
var Valeurs: T_Valeurs;
		CoefPaPr,CoefPrPv : double;
begin
  if TOBLOc.getValeur(indPa) > 0 then CoefpaPr := TOBLOc.getValeur(indPr) / TOBLOc.getValeur(indPa)
                                 else CoefpaPr := 1;
  if TOBLOc.getValeur(indPr) > 0 then CoefprPv := TOBLOc.getValeur(indPvHtDev) / TOBLOc.getValeur(indpr)
                                 else CoefprPv := 1;
  if EnPa then
  begin
    if (TOBLOC.GetValue('BLO_TYPEARTICLE') = 'ARP') and (TOBLOC.detail.count > 0) then
    begin
      TOBLOC.detail[0].Putvaleur(IndPa, arrondi(TOBLOC.detail[0].GetValeur(indpa) + (Ecart), V_PGI.okDecP));
      TOBLOC.detail[0].PutValue('GA_PAHT', TOBLOC.detail[0].GetValue('BLO_DPA'));
      CalculeLigneAcOuv (TOBLOC.detail[0],DEV,(TOBL.GetValue('GL_BLOQUETARIF')='-'));
//      TOBLOC.detail[0].Putvaleur(IndPr, arrondi(TOBLOC.detail[0].GetValeur(indpa) * coefPaPr, V_PGI.okDecP));
      if TOBL.GetValue('GL_BLOQUETARIF')='-' then
      begin
//        TOBLOC.detail[0].Putvaleur(IndPvHtDev, arrondi(TOBLOC.detail[0].GetValeur(indpr) * coefPrPv, V_PGI.okdecP));
//        TOBLOC.detail[0].Putvaleur(IndPvHt, devisetopivotEx(TOBLOC.detail[0].getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
        TOBLOC.detail[0].PutValue('GA_PVHT', TOBLOC.detail[0].Getvaleur(IndPvHtdev));
      end;
      InitTableau(Valeurs);
      CalculeOuvrageDoc(TOBLOC, 1, 1, true, DEV, valeurs, EnHt);
      TOBLOC.putvalue('BLO_DPA', valeurs[0]);
      TOBLOC.putvalue('BLO_DPR', valeurs[1]);
      TOBLOC.PutValue('BLO_PMAP', valeurs[6]);
      TOBLOC.PutValue('BLO_PMRP', valeurs[7]);
      if TOBL.GetValue('GL_BLOQUETARIF')='-' then
      begin
        TOBLOC.putvalue('BLO_PUHTDEV', valeurs[2]);
        TOBLOC.putvalue('BLO_PUTTCDEV', valeurs[3]);
        TOBLOC.putvalue('BLO_PUHT', devisetopivotEx(valeurs[2], DEV.Taux, DEV.quotite,V_PGI.okdecP));
        TOBLOC.putvalue('BLO_PUTTC', devisetopivotEx(valeurs[3], DEV.Taux, DEV.quotite,V_PGI.OkdecP));
      end;
      TOBLOC.putvalue('BLO_TPSUNITAIRE', valeurs[9]);
//
      TOBLOC.PutValue('GA_HEURE',valeurs[9]);
      TOBLOC.PutValue('BLO_MONTANTPAFG',valeurs[10]);
      TOBLOC.PutValue('BLO_MONTANTPAFR',valeurs[11]);
      TOBLOC.PutValue('BLO_MONTANTPAFC',valeurs[12]);
      TOBLOC.PutValue('BLO_MONTANTFG',valeurs[13]);
      TOBLOC.PutValue('BLO_MONTANTFR',valeurs[14]);
      TOBLOC.PutValue('BLO_MONTANTFC',valeurs[15]);
      TOBLOC.PutValue('BLO_MONTANTPA',valeurs[16]);
      TOBLOC.PutValue('BLO_MONTANTPR',valeurs[17]);
      //
      StockeLesTypes (TOBLOC,valeurs);
      //
    end else
    begin
      TOBLOC.Putvaleur(IndPa, arrondi(TobLoc.GetValeur(indpa) + (Ecart), V_PGI.okDecP));
      TOBLOC.PutValue('GA_PAHT', TOBLOC.GetValue('BLO_DPA'));
//      TOBLOC.Putvaleur(IndPr, arrondi(TobLoc.GetValeur(indpa) * coefPaPr, V_PGI.okDecP));
        CalculeLigneAcOuv (TOBLoc,DEV,(TOBL.GetValue('GL_BLOQUETARIF')='-'));
      if TOBL.GetValue('GL_BLOQUETARIF')='-' then
      begin
//        TOBLOC.Putvaleur(IndPvHtDev, arrondi(TobLoc.GetValeur(indpr) * coefPrPv, V_PGI.okdecP));
//        TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(TOBLOc.getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
        TOBLOC.PutValue('GA_PVHT', TOBLOC.Getvaleur(IndPvHtdev));
        if TOBLOC.GetValue('BLO_TYPEARTICLE') = 'PRE' then
          TOBLOC.PutValue('GF_PRIXUNITAIRE', TOBLOC.Getvaleur(IndPvHtdev));
      end;
    end;
  end else
  begin
    if EnPR then
    begin
      //-------------------
      if (TOBLOC.GetValue('BLO_TYPEARTICLE') = 'ARP') and (TOBLOC.detail.count > 0) then
      begin
        TOBLOC.detail[0].Putvaleur(IndPr, arrondi(TOBLOC.detail[0].GetValeur(indPr) + (EcartPr), V_PGI.okdecP));
        TOBLOC.detail[0].Putvaleur(IndPvHtdev, arrondi(TOBLOC.detail[0].GetValeur(indpr) * CoefPrPv, V_PGI.okdecP));
        TOBLOC.detail[0].Putvaleur(IndPvHt, devisetopivotEx(TOBLOC.detail[0].getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
        InitTableau(Valeurs);
        CalculeOuvrageDoc(TOBLOC, 1, 1, true, DEV, valeurs, EnHt);
        TOBLOC.putvalue('BLO_DPA', valeurs[0]);
        TOBLOC.putvalue('BLO_DPR', valeurs[1]);
        TOBLOC.PutValue('BLO_PMAP', valeurs[6]);
        TOBLOC.PutValue('BLO_PMRP', valeurs[7]);
        if TOBL.GetValue('GL_BLOQUETARIF')='-' then
        begin
          TOBLOC.putvalue('BLO_PUHTDEV', valeurs[2]);
          TOBLOC.putvalue('BLO_PUTTCDEV', valeurs[3]);
          TOBLOC.putvalue('BLO_PUHT', devisetopivotEx(valeurs[2], DEV.Taux, DEV.quotite,V_PGI.okdecP));
          TOBLOC.putvalue('BLO_PUTTC', devisetopivotEx(valeurs[3], DEV.Taux, DEV.quotite,V_PGI.okdecP));
          if TOBLoc.GetValue('BLO_DPR')<>0 then
          begin
            TOBLoc.putValue('BLO_COEFMARG',arrondi(TOBLoc.GetValue('BLO_PUHT')/TOBLoc.GetValue('BLO_DPR'),4));
            TOBLoc.PutValue('POURCENTMARG',Arrondi((TOBLoc.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
          if TOBLoc.GetValue('BLO_PUHT') <> 0 then
          begin
            TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
          end else
          begin
            TOBLoc.PutValue('POURCENTMARQ',0);
          end;
        end;
        TOBLOC.putvalue('BLO_TPSUNITAIRE', valeurs[9]);
        StockeLesTypes (TOBLOC,valeurs);
      end else
      begin
        //-------------------
        //             BRL : inutile car dj fait plus haut et plante si diviseur = 0
        //      			 CoefprPv := TOBLOc.getValeur(indPvHtDev)/TOBLOc.getValeur(indpr);
        TOBLOC.Putvaleur(IndPr, arrondi(TobLoc.GetValeur(indPr) + (EcartPr), V_PGI.okdecP));
        if TOBL.GetValue('GL_BLOQUETARIF')='-' then
        begin
          TOBLOC.Putvaleur(IndPvHtdev, arrondi(TobLoc.GetValeur(indpr) * CoefPrPv, V_PGI.okdecP));
          TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(TOBLOc.getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
          if TOBLoc.GetValue('BLO_DPR')<>0 then
          begin
            TOBLoc.putValue('BLO_COEFMARG',arrondi(TOBLoc.GetValue('BLO_PUHT')/TOBLoc.GetValue('BLO_DPR'),4));
            TOBloc.PutValue('POURCENTMARG',Arrondi((TOBloc.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
          if TOBLoc.GetValue('BLO_PUHT') <> 0 then
          begin
            TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
          end else
          begin
            TOBLoc.PutValue('POURCENTMARQ',0);
      end;

        end;
      end;
    end else
    begin
      //--------
      //-------------------
      if (TOBLOC.GetValue('BLO_TYPEARTICLE') = 'ARP') and (TOBLOC.Detail.count > 0) then
      begin
        TOBLOC.detail[0].Putvaleur(IndPvHtdev, arrondi(TOBLOC.detail[0].GetValeur(indpvhtdev) + (Ecart), V_PGI.okdecP));
        TOBLOC.detail[0].Putvaleur(IndPvHt, devisetopivotEx(TOBLOC.detail[0].getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
        if ForcePaEqualPv then
        begin
          TOBLOC.detail[0].SetDouble('BLO_DPA',TOBLOC.detail[0].GetDouble('BLO_PUHT'));
          TOBLOC.detail[0].SetDouble('BLO_DPR',TOBLOC.detail[0].GetDouble('BLO_DPA'));
          TOBLOC.detail[0].SetDouble('BLO_COEFFG',0);
          TOBLOC.detail[0].SetDouble('BLO_COEFFC',0);
          TOBLOC.detail[0].SetDouble('BLO_COEFFR',0);
          TOBLOC.detail[0].SetDouble('BLO_COEFMARG',1);
        end;
        if TOBLoc.detail[0].GetValue('BLO_DPR')<>0 then
        begin
          TOBLoc.detail[0].putValue('BLO_COEFMARG',arrondi(TOBLoc.detail[0].GetValue('BLO_PUHT')/TOBLoc.detail[0].GetValue('BLO_DPR'),4));
          TOBLoc.detail[0].PutValue('POURCENTMARG',Arrondi((TOBLoc.detail[0].GetValue('BLO_COEFMARG')-1)*100,2));
        end;
        if TOBLoc.detail[0].GetValue('BLO_PUHT') <> 0 then
        begin
          TOBLoc.detail[0].PutValue('POURCENTMARQ',Arrondi(((TOBLoc.detail[0].GetValue('BLO_PUHT')- TOBLoc.detail[0].GetValue('BLO_DPR'))/TOBLoc.detail[0].GetValue('BLO_PUHT'))*100,2));
        end else
        begin
          TOBLoc.detail[0].PutValue('POURCENTMARQ',0);
        end;
        InitTableau(Valeurs);
        CalculeOuvrageDoc(TOBLOC, 1, 1, true, DEV, valeurs, EnHt);
        TOBLOC.putvalue('BLO_DPA', valeurs[0]);
        TOBLOC.putvalue('BLO_DPR', valeurs[1]);
        TOBLOC.PutValue('BLO_PMAP', valeurs[6]);
        TOBLOC.PutValue('BLO_PMRP', valeurs[7]);
        if TOBL.GetValue('GL_BLOQUETARIF')='-' then
        begin
          TOBLOC.putvalue('BLO_PUHTDEV', valeurs[2]);
          TOBLOC.putvalue('BLO_PUTTCDEV', valeurs[3]);
          TOBLOC.putvalue('BLO_PUHT', devisetopivotEx(valeurs[2], DEV.Taux, DEV.quotite,V_PGI.okdecP));
          TOBLOC.putvalue('BLO_PUTTC', devisetopivotEx(valeurs[3], DEV.Taux, DEV.quotite,V_PGI.okdecP));
          if TOBLoc.GetValue('BLO_DPR')<>0 then
          begin
            TOBLoc.putValue('BLO_COEFMARG',arrondi(TOBLoc.GetValue('BLO_PUHT')/TOBLoc.GetValue('BLO_DPR'),4));
            TOBLoc.PutValue('POURCENTMARG',Arrondi((TOBLoc.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
          if TOBLoc.GetValue('BLO_PUHT') <> 0 then
          begin
            TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
          end else
          begin
            TOBLoc.PutValue('POURCENTMARQ',0);
          end;
        end;
        TOBLOC.PutValue('GA_HEURE',valeurs[9]);
        TOBLOC.PutValue('BLO_MONTANTPAFG',valeurs[10]);
        TOBLOC.PutValue('BLO_MONTANTPAFR',valeurs[11]);
        TOBLOC.PutValue('BLO_MONTANTPAFC',valeurs[12]);
        TOBLOC.PutValue('BLO_MONTANTFG',valeurs[13]);
        TOBLOC.PutValue('BLO_MONTANTFR',valeurs[14]);
        TOBLOC.PutValue('BLO_MONTANTFC',valeurs[15]);
        TOBLOC.PutValue('BLO_MONTANTPA',valeurs[16]);
        TOBLOC.PutValue('BLO_MONTANTPR',valeurs[17]);
      end else
      begin
        //-------------------
        //--------
        if TOBL.GetValue('GL_BLOQUETARIF')='-' then
        begin
          TOBLOC.Putvaleur(IndPvHtdev, arrondi(TobLoc.GetValeur(indpvhtdev) + (Ecart), V_PGI.okdecP));
          TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(TOBLOc.getValeur(IndPvHtdev), DEV.Taux, DEV.quotite, V_PGI.okdecP));
          if ForcePaEqualPv then
          begin
            TOBLOC.SetDouble('BLO_DPA',TOBLoc.GetDouble('BLO_PUHT'));
            TOBLOC.SetDouble('BLO_DPR',TOBLoc.GetDouble('BLO_DPA'));
            TOBLOC.SetDouble('BLO_COEFFG',0);
            TOBLOC.SetDouble('BLO_COEFFC',0);
            TOBLOC.SetDouble('BLO_COEFFR',0);
            TOBLOC.SetDouble('BLO_COEFMARG',1);
          end;
          if TOBLoc.GetValue('BLO_DPR')<>0 then
          begin
            TOBLoc.putValue('BLO_COEFMARG',arrondi(TOBLoc.GetValue('BLO_PUHT')/TOBLoc.GetValue('BLO_DPR'),4));
            TOBLoc.PutValue('POURCENTMARG',Arrondi((TOBLoc.GetValue('BLO_COEFMARG')-1)*100,2));
        end;
          if TOBLoc.GetValue('BLO_PUHT') <> 0 then
          begin
            TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
          end else
          begin
            TOBLoc.PutValue('POURCENTMARQ',0);
      end;
    end;
  end;
    end;
  end;
  CalculMontantHtDevLigOuv(ToBloc, DEV);
//  positionneCoefMarge (TOBL);
end;

procedure EssayeAffecterEcartQte (TOBL,TOBOuvrage : TOB; Ecart,EcartPr : double ; var SumLoc: double ; var SumPr : double;
																  IndPa,IndPR,IndPVHtDev,IndPVHt : integer; EnPA,EnPr,EnHt : boolean; DEV : Rdevise; TouslesNiveaux : boolean= false; ForcePaEqualPv : Boolean=false);
var Indice : integer;
    restitue,restituePr,EnCoursPr,EnCours ,MontantLoc,MontantLocPr, LocSumLoc,LocSumLocPr,LastMontantLoc,LastMontantLocPr: double;
    TOBO : TOB;
    CoefpaPr,CoefPrPV,ValTrav : double;
    TOBPiece : TOB;
    CoefFGfixe : boolean;
    ValPou : T_Valeurs;
    IndPou : integer;
    ArticleOk : string;
begin
  TOBO := TOBOuvrage.findfirst(['BLO_TYPEARTICLE'], ['POU'], false);
  if TOBO <> nil then
  begin
    exit;
  end;
	InitTableau (Valpou);

	CoeffGFixe := false;
  TOBPiece := TOBL.Parent;
  if (TOBPiece <> nil) and
  	 (TOBPiece.FieldExists('COEFFGFORCE')) and
     (TOBPiece.getValue('COEFFGFORCE')<> 0) then CoefFgFixe := true;
	SumLoc := 0;
  SumPr := 0;
	Restitue := Ecart;
  restituePr := EcartPr;
  IndPou := -1;
  ArticleOk := '';
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
  	TOBO := TOBOuvrage.detail[Indice];
    if TOBO.GetValue('BLO_QTEFACT') = 0 then continue;
    if TOBO.GetValue('BLO_TYPEARTICLE') = 'POU' then continue;
    //
    ValTrav := VALEUR(TOBO.GetValue('BLO_NATURETRAVAIL'));
    if (ValTrav > 0) then Continue;   // cotraité ou sous traité
    //
    EnCours :=  Arrondi (Restitue/ TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKdecP);
    EnCoursPr :=  Arrondi (RestituePr/TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKdecP);
    if (TousLesNiveaux) and
    	 (TOBO.GetValue('BLO_TYPEARTICLE')='OUV') and
       (not Isvariante(TOBO)) and
       (EnCOurs<>0) then
    begin
      if EnPa then LastMontantLoc := TOBO.getValue('BLO_MONTANTPA')
      else if EnHt then LastMontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
      else LastMontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');
      if TOBO.getValeur(indPa) > 0 then CoefpaPr := TOBO.getValeur(indPr) / TOBO.getValeur(indPa)
                                   else CoefpaPr := 1;
      if TOBO.getValeur(indPr) > 0 then CoefprPv := TOBO.getValeur(indPvHtDev) / TOBO.getValeur(indpr)
      														 else CoefprPv := 1;

      LastMontantLocPr := TOBO.getValue('BLO_MONTANTPR');
      LocSumLoc := 0;
      LocSumLocPr := 0;

    	EssayeAffecterEcartQte (TOBL,TOBO,Encours,EncoursPr,LocSumLoc,LocSumLocPr,IndPa,IndPr,IndPVHtdev,IndPvHt,EnPa,EnPR,EnHt,DEV,TousLesNiveaux,ForcePaEqualPv);

      if ENPA then
      begin
        TOBO.Putvaleur(IndPa, LocSumLoc);
        TOBO.PutValue('GA_PAHT', TOBO.GetValue('BLO_DPA'));
        if not CoefFgFixe then TOBO.Putvaleur(IndPr, Arrondi(LocSumLoc * CoefPaPr, V_PGI.okDecP));
        if (TOBL.GetValue('GL_BLOQUETARIF')='-')  then
        begin
          TOBO.Putvaleur(IndPvHtDev, Arrondi(TOBO.getValeur(IndPr) * CoefPrPv, V_PGI.okdecP));
          TOBO.Putvaleur(IndPvHt, devisetopivotEx(TOBO.Getvaleur(IndPvHtDev), DEV.Taux, DEV.quotite, V_Pgi.okdecP));
          TOBO.PutValue('GA_PVHT', TOBO.Getvaleur(IndPvHtdev));
        end;
        if TOBO.GetValue('BLO_TYPEARTICLE') = 'PRE' then
          TOBO.PutValue('GF_PRIXUNITAIRE', TOBO.Getvaleur(IndPvHtdev));
      end else
      begin
        TOBO.Putvaleur(IndPvHtDev, LocSumLoc);
        TOBO.Putvaleur(IndPvHt, devisetopivotEx(LocSumLoc, DEV.Taux, DEV.quotite, V_PGI.okdecP));
        if ForcePaEqualPv then
        begin
          TOBO.putvalue ('BLO_DPA',LocSumLoc);
          TOBO.PutValue('GA_PAHT', TOBO.GetValue('BLO_DPA'));
          TOBO.putvalue ('BLO_DPR',LocSumLoc);
          TOBO.putvalue ('BLO_COEFFG',0);
          TOBO.putvalue ('BLO_COEFFC',0);
          TOBO.putvalue ('BLO_COEFFR',0);
          TOBO.SetDouble('BLO_COEFMARG',1);
      end;
      end;
      CalculMontantHtDevLigOuv(TOBO, DEV);
      if EnPa then MontantLoc := TOBO.getValue('BLO_MONTANTPA')
      else if EnHt then MontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
      else MontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

      MontantLocPr := TOBO.getValue('BLO_MONTANTPR');

      restitue := ARRONDI(restitue - (MontantLoc - lastMontantLoc),V_PGI.OkDecP);
      restituePr := ARRONDI(restituePr - (MontantLocPr - LastMontantLocPr),V_PGI.OkDecP);

    end else
    begin
      if (TOBO.GEtValue('BLO_TYPEARTICLE')<> 'OUV' ) and
         (TOBO.GEtValue('BLO_TYPEARTICLE')<> 'POU') and (not IsVariante(TOBO)) and
         (EnCours <> 0) then
      begin
        if EnPa then LastMontantLoc := TOBO.getValue('BLO_MONTANTPA')
        else if EnHt then LastMontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
        else LastMontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

        LastMontantLocPr := TOBO.getValue('BLO_MONTANTPR');
        AffecteDifferenceSurLigne (TOBL,TOBO,EnCours,EnCoursPr,IndPa,IndPr,IndPVHtDev,IndPVHt,EnPa,EnPr,EnHt,DEV,ForcePaEqualPv);
        CalculMontantHtDevLigOuv(TOBO, DEV);
        if EnPa then MontantLoc := TOBO.getValue('BLO_MONTANTPA')
        else if EnHt then MontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
        else MontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

        MontantLocPr := TOBO.getValue('BLO_MONTANTPR');

        restitue := ARRONDI(restitue - (MontantLoc - lastMontantLoc),V_PGI.OkDecP);
        restituePr := arrondi(restituePr - (MontantLocPr - LastMontantLocPr),V_PGI.OkDecP);

      end else
      begin
        if EnPa then MontantLoc := TOBO.getValue('BLO_MONTANTPA')
        else if EnHt then MontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
        else MontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

        MontantLocPr := TOBO.getValue('BLO_MONTANTPR');
      end;
    end;
    if ArticleOKInPOUR(TOBO.GetValue('BLO_TYPEARTICLE'), ArticleOk) then
    begin
      ValPou[0] := ValPou[0] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_DPA'));
      ValPou[1] := ValPou[1] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_DPR'));
      ValPou[6] := ValPou[6] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_PMAP'));
      ValPou[7] := ValPou[7] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_PMRP'));
      ValPou[2] := ValPou[2] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_PUHTDEV'));
      ValPou[3] := ValPou[3] + ((TOBO.GetValue ('BLO_QTEFACT')/TOBO.GetValue ('BLO_QTEDUDETAIL')) * TOBO.GetValue ('BLO_PUTTCDEV'));
    end;
    // --
    SumLoc := ARRONDI(SumLoc + MontantLoc,V_PGI.okdecP);
    SumPr := ARRONDI(SumPr + MontantLocPr,V_PGI.OkDecP);
  end;

  if (Indpou >= 0) then
  begin
  	FormatageTableau (ValPou,V_PGI.okdecP);

    TOBO := TOBOUvrage.detail[Indpou];
    if ENPA then
    begin
      TOBO.Putvaleur(IndPa, Valpou[0]);
      TOBO.PutValue('GA_PAHT', TOBO.GetValue('BLO_DPA'));
      TOBO.Putvaleur(IndPr, Arrondi(ValPou[1], V_PGI.okDecP));
      if (TOBL.GetValue('GL_BLOQUETARIF')<>'X')  then
      begin
        TOBO.Putvaleur(IndPvHtDev, Arrondi(ValPou[2], V_PGI.okdecP));
        TOBO.Putvaleur(IndPvHt, devisetopivotEx(ValPou[2], DEV.Taux, DEV.quotite, V_Pgi.okdecP));
        TOBO.PutValue('GA_PVHT', TOBO.Getvaleur(IndPvHtdev));
      end;
      if TOBO.GetValue('BLO_TYPEARTICLE') = 'PRE' then
        TOBO.PutValue('GF_PRIXUNITAIRE', TOBO.Getvaleur(IndPvHtdev));
    end else
    begin
      TOBO.Putvaleur(IndPvHtDev, ValPou[2]);
      TOBO.Putvaleur(IndPvHt, devisetopivotEx(ValPou[2], DEV.Taux, DEV.quotite, V_PGI.okdecP));
      if ForcePaEqualPv then
      begin
        TOBO.putvalue ('BLO_DPA',TOBO.Getvaleur(IndPvHtDev));
        TOBO.PutValue('GA_PAHT', TOBO.GetValue('BLO_DPA'));
        TOBO.putvalue ('BLO_DPR',TOBO.Getvaleur(IndPvHtDev));
        TOBO.putvalue ('BLO_COEFFG',0);
        TOBO.putvalue ('BLO_COEFFC',0);
        TOBO.putvalue ('BLO_COEFFR',0);
        TOBO.SetDouble('BLO_COEFMARG',1);
    end;
    end;
    CalculMontantHtDevLigOuv(TOBO, DEV);
    if EnPa then MontantLoc := TOBO.getValue('BLO_MONTANTPA')
    else if EnHt then MontantLoc := TOBO.getValue('BLO_MONTANTHTDEV')
    else MontantLoc := TOBO.getValue('BLO_MONTANTTTCDEV');

    MontantLocPr := TOBO.getValue('BLO_MONTANTPR');
    // --
    SumLoc := SumLoc + MontantLoc;
    SumPr := SumPr + MontantLocPr;
  end;
  SumLoc := ARRONDI(SumLoc,V_PGI.OkDecP);
  SumPr := ARRONDI(SumPr,V_PGI.OkDecP);
end;
//-------------


function ReajusteMontantOuvrage(TOBARticles,TOBPiece, TOBL, TOBOuvrage: TOB; Montantcalc: double; var MontantPr: double; var MontantDemande: double; DEV: RDevise; EnHt: boolean; atributionEcart: boolean = true; EnPR: boolean = false; EnPa: boolean = false;ForcePaEqualPv:boolean=false): boolean;
var
  Indice, IndPa, IndPr, IndPvHt, IndPvhtdev, IndPvBase, IndQte, IndQteDuDetail, IndQteDuPv: Integer;
  MontantLoc, MontantLocPr, Ppq, CoefApplic, PrixInterm, QTe, QteDuDetail, QteDuPv, SumLoc, MontantVentil, MontantVentilPrec: Double;
  MontantLocDet, SumLocDet, IndiceEcart, IndPou, IndTypeArt,IndNatureTravail: Integer;
  PuMax, CoefMG, PrixPr: Double;
  TOBLOC,TOBA: TOB;
  Valinit, valfin: double;
  ArtEcart, ArticleOk: string;
  IndiceFixe: boolean;
  MontantBasePou: double;
  CoefPaPr, CoefPrPv: double;
  SumPr: double;
  Valeurs,ValPou : T_Valeurs;
  CoefFgFixe : boolean;
  ValTrav : double;
  InitMontantdemande,InitMontantPR : double;

begin
  InitMontantDemande := MontantDemande;
  InitMontantPR := MontantPR;
  //
  InitTableau (Valpou);

	CoefFgFixe := false;
  if (TOBPiece <> nil) and
  	 (TOBPiece.FieldExists('COEFFGFORCE')) and
     (TOBPiece.getValue('COEFFGFORCE')<> 0) then CoefFgFixe := true;
  result := true;
  MontantBasePou := 0;
  IndPou := -1;
  IndiceFixe := false;
  IndiceEcart := -1;
  PuMax := 0;
  SUmLoc := 0;
  ArtEcart := trim(GetParamsoc('SO_BTECARTPMA'));
  //
  if EnPR then atributionEcart := false;
  //
  if not isExistsArticle(ArtEcart) then
  begin
    PgiBox(TraduireMemoire('L''article d''cart est invalide ou non renseign#13#10Veuillez le dfinir'), Traduirememoire('Gestion d''cart'));
    result := false;
    exit;
  end;
  QteDudetail := TOBOuvrage.detail[0].getvalue('BLO_QTEDUDETAIL');
  if QteDuDetail = 0 then QteDuDetail := 1;
  // fv 02032004
  //if QteDuDetail =0 then qtedudetail := 1;
  if EnPr then MontantDemande := MontantPr * QteDudetail
  				else MontantDemande := MontantDemande * QteDudetail;
  MontantCalc := MontantCalc * QteDuDetail; // prcdente valeur
  if EnPr then MontantVentil := MontantPr // Montant  ventiler sur les lignes
  				else MontantVentil := MontantDemande; // Montant  ventiler sur les lignes
  MontantVentilPrec := MontantCalc; // Montant ventil prcdent


  TOBLOC := TOBOuvrage.findfirst(['BLO_TYPEARTICLE'], ['POU'], false);
  if TOBLOC <> nil then
  begin
  	ArticleOk := TOBLOc.GetValue('BLO_LIBCOMPL');
    IndPou := TOBLoc.GetIndex;
  end;

  for Indice := 0 to TOBOuvrage.detail.Count - 1 do
  begin
    TOBLOC := TOBOuvrage.detail[Indice];
    if trim(TOBLOc.GetValue('BLO_CODEARTICLE')) = ArtEcart then
    begin
      IndiceEcart := Indice;
      IndiceFixe := true;
    end;
  end;

  (*
  if IndPou >= 0 then
  begin
    TOBLOC := TOBOuvrage.detail[IndPou];
    Ppq := TOBLOc.GetValue('BLO_PRIXPOURQTE');
    if Ppq = 0 then Ppq := 1;
    MontantVentil := arrondi(MontantDemande / (1 + (TOBLOc.getvalue('BLO_QTEFACT') / Ppq)), DEV.Decimale);
    MontantVentilPrec := arrondi(MontantCalc / (1 + (TOBLOc.getvalue('BLO_QTEFACT') / Ppq)), DEV.Decimale);
  end;
  *)
  MontantLoc := 0;
  SumLoc := 0;
  if MontantVentilPrec = 0 then Coefapplic := MontantVentil / TOBOuvrage.detail.Count
  												 else CoefApplic := MontantVentil / MontantVentilPrec;
  if CoefApplic=1 then
  begin
    // ca evitera de partir en vrille
    MontantDemande := InitMontantDemande;
    MontantPR := InitMontantPR;
  	exit; // si on ne change rien pas la peine de faire des calculs
  end;
//  if CoefApplic <> 0 then
  begin
    if EnHT then
    begin
      IndPvBase := TOBOuvrage.detail[0].GetNumChamp('BLO_PUHTBASE');
      IndPvHtDev := TOBOuvrage.detail[0].GetNumChamp('BLO_PUHTDEV');
      IndPvHt := TOBOuvrage.detail[0].GetNumChamp('BLO_PUHT');
    end else
    begin
      IndPvBase := TOBOuvrage.detail[0].GetNumChamp('BLO_PUTTCBASE');
      IndPvHtDev := TOBOuvrage.detail[0].GetNumChamp('BLO_PUTTCDEV');
      IndPvHt := TOBOuvrage.detail[0].GetNumChamp('BLO_PUTTC');
    end;
    IndPa := TOBOuvrage.detail[0].GetNumChamp('BLO_DPA');
    IndPr := TOBOuvrage.detail[0].GetNumChamp('BLO_DPR');
    IndQte := TOBOuvrage.detail[0].GetNumChamp('BLO_QTEFACT');
    IndQteDuDetail := TOBOuvrage.detail[0].GetNumChamp('BLO_QTEDUDETAIL');
    IndTypeArt := TOBOuvrage.detail[0].GetNumChamp('BLO_TYPEARTICLE');
    IndQteDuPv := TOBOuvrage.detail[0].GetNumChamp('BLO_PRIXPOURQTE');
    IndNatureTravail := TOBOuvrage.detail[0].GetNumChamp('BLO_NATURETRAVAIL');
    for Indice := 0 to TOBOuvrage.detail.Count - 1 do
    begin
      TOBLOC := TOBOuvrage.detail[Indice];
      if TOBLOC.getvaleur(IndTypeart) = 'POU' then continue;
      // Rajout LS le 17/10/2003 -- gestion des variantes
      if IsVariante(TOBLoc) then continue;
      Qte := TOBLoc.getvaleur(IndQte);
    	QteDuDetail := TOBLoc.GetValue('BLO_QTEDUDETAIL');
    	if QteDuDetail = 0  then QteDuDetail := 1;

      // --
      ValTrav := VALEUR(TOBloc.GetValeur(IndNatureTravail));
      if (ValTrav = 0) or (ValTrav >= 10) then
      begin
        // --
        PrixPr := TOBLOc.getValeur(IndPr);
        // --
        if EnPa then
        begin
          (*
            if TOBLOc.getValeur(indPa) > 0 then CoefpaPr := TOBLOc.getValeur(indPr) / TOBLOc.getValeur(indPa)
                                           else CoefpaPr := 1;
            if TOBLOc.getValeur(indPr) > 0 then CoefprPv := TOBLOc.getValeur(indPvHtDev) / TOBLOc.getValeur(indpr)
                                           else CoefprPv := 1;
          *)
          if TOBLoc.getValeur(IndPa) = 0 then
          begin
            Qte := TobLoc.getValue('BLO_QTEFACT');
            if Qte = 0 then Qte := 1;
            PrixInterm := arrondi(CoefApplic / Qte, V_PGI.okdecP);
          end else
          begin
            PrixInterm := arrondi(TOBLOc.getValeur(Indpa) * CoefApplic, V_PGI.okdecP);
          end;
        end else if EnPR then // affectation coefficient sauf pour sous-traitance
        begin
  //        if (TOBLoc.getValue('BLO_TYPEARTICLE') = 'PRE') and (RenvoieTypeRes(TOBLOc.GetValue('BLO_ARTICLE')) = 'ST') then
          if (TOBLoc.getValue('BLO_TYPEARTICLE') = 'PRE') and (TOBLOc.GetValue('BNP_TYPERESSOURCE') = 'ST') then
          begin
            PrixInterm := TOBLOc.getValeur(IndPvHtDev);
            PrixPr := TOBLOc.getValeur(IndPr);
          end else IndPa := TOBLoc.GetNumChamp('BLO_DPA');
  (*          if TOBLoc.getvaleur (IndPr) <> 0 then CoefMG:=TOBLoc.getvaleur (IndPvHtDev)/TOBLoc.getvaleur (IndPr)
          else CoefMG:=0; *)
          if not CoefFgFixe then
            TOBLoc.PutValeur(IndPr, arrondi(TOBLoc.getvaleur(IndPr) * Coefapplic, V_PGI.okdecP));
          PrixPr := TOBLoc.getValeur(IndPr);
          PrixInterm := arrondi(TOBLOc.getValeur(IndPr) * TOBLOc.Getvalue('BLO_COEFMARG'), V_PGI.okdecP);
        end else
        begin
          if TOBLoc.getValeur(IndPvHtDev) = 0 then
          begin
            Qte := TobLoc.getValue('BLO_QTEFACT');
            if Qte = 0 then Qte := 1;
            PrixInterm := arrondi(CoefApplic / Qte, V_PGI.okdecP);
          end else
          begin
            PrixInterm := arrondi(TOBLOc.getValeur(IndPvHtDev) * CoefApplic, V_PGI.okdecP);
          end;
        end;
        // --
        if EnPR then
        begin
          if (TobLoc.getvalue('ANCPR') <> PrixPr) and
            (TobLoc.Detail.count > 0) then
            ReajusteMontantOuvrage(TOBArticles,TOBPiece, TOBL,TOBLoc, TOBLOc.getValue('ANCPR'), PrixPr, PrixInterm, DEV, EnHt, atributionEcart, EnPr, EnPa);
        end else if EnPa then
        begin
          if (TobLoc.getvalue('ANCPA') <> PrixInterm) and
            (TobLoc.Detail.count > 0) then
            ReajusteMontantOuvrage(TOBArticles,TOBPiece, TOBL, TOBLoc, TOBLOc.getValue('ANCPA'), PrixPr, PrixInterm, DEV, EnHt, atributionEcart, EnPr, EnPa);
        end else
        begin
          if (TobLoc.getvalue('ANCPV') <> PrixInterm) and
            (TobLoc.Detail.count > 0) then
            ReajusteMontantOuvrage(TOBArticles,TOBPiece,TOBL, TOBLoc, TOBLOc.getValue('ANCPV'), PrixPr, PrixInterm, DEV, EnHt, atributionEcart, EnPr,false,ForcePaEqualPv);
        end;
        // --
        if ENPA then
        begin
          if TOBARticles = nil then
          begin
            // on fait rien --> on passe
            TOBA := nil;
          end else if TOBARticles.NomTable = 'ARTICLE' then
          begin
            TOBA := TOBARticles.findFirst(['GA_ARTICLE'],[TOBLOC.getValue('BLO_ARTICLE')],true);
          end else
          begin
            TOBA := TOBArticles;
          end;

          TOBLOC.Putvaleur(IndPa, PrixInterm);
          TOBLOC.PutValue('GA_PAHT', TOBLOC.GetValue('BLO_DPA'));
          CalculeLigneAcOuv (TOBLoc,DEV,(TOBL.GetValue('GL_BLOQUETARIF')='-'),TOBA);
          if TOBL.GetValue('GL_BLOQUETARIF')='-' then
          begin
  //          TOBLOC.Putvaleur(IndPvHtDev, Arrondi(TOBLoc.getValeur(IndPr) * CoefPrPv, V_PGI.okdecP));
  //          TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(TOBLOC.Getvaleur(IndPvHtDev), DEV.Taux, DEV.quotite, V_Pgi.okdecP));
            TOBLOC.PutValue('GA_PVHT', TOBLOC.Getvaleur(IndPvHtdev));
            if TOBLOC.GetValue('BLO_TYPEARTICLE') = 'PRE' then
              TOBLOC.PutValue('GF_PRIXUNITAIRE', TOBLOC.Getvaleur(IndPvHtdev));
          end;
        end else
        begin
          if (TOBL.GetValue('GL_BLOQUETARIF') = '-') and (TobLoc.getvalue('ANCPV') <> PrixInterm)then
          begin
            TOBLOC.Putvaleur(IndPvHtDev, PrixInterm);
            TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(PrixInterm, DEV.Taux, DEV.quotite, V_PGI.okdecP));
            if ForcePaEqualPv then
            begin
              TOBLOC.Putvaleur(IndPa, PrixInterm);
              TOBLOC.Putvaleur(IndPr, PrixInterm);
              TOBLOC.Setdouble('BLO_COEFFG', 0);
              TOBLOC.Setdouble('BLO_COEFFC', 0);
              TOBLOC.Setdouble('BLO_COEFFR', 0);
              TOBLOC.Setdouble('BLO_COEFMARG', 1);
            end;
            if EnHT then
            begin
              TOBLOC.PutValue('GA_PVHT',PrixInterm);
              TOBLOC.PutValue('GF_PRIXUNITAIRE',PrixInterm);
            end;
          end;
        end;
        Qte := TOBLoc.getvaleur(IndQte);
        QteDuPv := TOBLOc.getValeur(IndQteDupv);
        {Correction suite a erreur dans le calcul prix march-- pb d'arrondi }
        // MontantLoc := (Qte/QteDuPv) * TOBLOC.GetValeur (IndPvHtDev);

        if EnHt then CalculeLigneHTOuv(TOBLoc, TOBPiece, DEV)
                else CalculeLigneTTCOuv(TOBLoc, TOBPiece, DEV);

      end;
      // Correction LS       CalculMontantHtDevLigOuv (ToBLOC,DEV);
      if EnPa then MontantLoc := TOBLoc.getValue('BLO_MONTANTPA')
      else if enHt then MontantLoc := TOBLoc.getValue('BLO_MONTANTHTDEV')
      else MontantLoc := TOBLoc.getValue('BLO_MONTANTTTCDEV');

      MontantLocPr := TOBLoc.getValue('BLO_MONTANTPR');
      // --
      if ArticleOKInPOUR(TOBLOC.GetValeur(IndTypeArt), ArticleOk) then
      begin
        ValPou[0] := ValPou[0] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_DPA'));
        ValPou[1] := ValPou[1] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_DPR'));
        ValPou[6] := ValPou[6] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_PMAP'));
        ValPou[7] := ValPou[7] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_PMRP'));
        ValPou[2] := ValPou[2] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_PUHTDEV'));
        ValPou[3] := ValPou[3] + ((Qte/QteDuDetail) * TOBLoc.GetValue ('BLO_PUTTCDEV'));
      end;
      SumLoc := SumLoc + MontantLoc;
      SumPr := SumPr + MontantLocPr;
      if (not IndiceFixe) and (Qte = 1) and (PuMax < MontantLoc) and
        ((TOBLOC.detail.count = 0) or (TOBLOC.GetValue('BLO_TYPEARTICLE') = 'ARP')) then
      begin
        // on autorise maintenant l'application de l'ecart sur les article en prix pose
        PuMax := MontantLoc;
        IndiceEcart := Indice;
      end;
      //
      if (TOBLOc.getValue('BLO_TYPEARTICLE')='ARP') and (TOBLOC.detail.count > 0) then
      begin
      	GetValoDetail (TOBloc);
      end;
      //
    end;
    SumPr := arrondi(sumPr, V_PGI.okdecP);
    sumloc := arrondi(sumloc, V_PGI.okdecP);
    montantBasePou := arrondi(MontantBasePou, V_PGI.okdecP);

    if (Indpou >= 0) then
    begin
      TOBLOc := TOBOUvrage.detail[Indpou];
      Qte := TOBLoc.getvaleur(IndQte);
      QteDuPv := TOBLOc.getValeur(IndQteDupv);
    	QteDuDetail := TOBLoc.GetValue('BLO_QTEDUDETAIL');
      if EnPA then
      begin
        TOBLOC.Putvaleur(IndPa, ValPou[0]);
        TOBLOC.PutValue('GA_PAHT', ValPou[0]);
        if not CoefFgFixe then TOBLOC.Putvaleur(IndPr, ValPou[1]);
        if TOBL.GetValue('GL_BLOQUETARIF') <> 'X' then
        begin
          TOBLOC.Putvaleur(IndPVhtDEV, arrondi(ValPou[2], V_PGI.okdecP));
          TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(ValPou[2], DEV.Taux, DEV.quotite, V_PGI.okdecP));
          TOBLOC.PutValue('GA_PVHT', TOBLOC.Getvaleur(IndPvHtdev));
          if TOBLOC.GetValue('BLO_TYPEARTICLE') = 'PRE' then
            TOBLOC.PutValue('GF_PRIXUNITAIRE', TOBLOC.Getvaleur(IndPvHtdev));
        end;
      end else
      begin
        if TOBL.GetValue('GL_BLOQUETARIF') <> 'X' then
        begin
          TOBLOC.Putvaleur(IndPvHtdev, arrondi(ValPou[2], V_PGI.okdecP));
          TOBLOC.Putvaleur(IndPvHt, devisetopivotEx(arrondi(ValPou[2], V_PGI.okdecP), DEV.Taux, DEV.quotite,V_PGI.okdecP));
          if ForcePaEqualPv then
          begin
            TOBLOC.Putvaleur(IndPa, PrixInterm);
            TOBLOC.Putvaleur(IndPr, PrixInterm);
            TOBLOC.Setdouble('BLO_COEFFG', 0);
            TOBLOC.Setdouble('BLO_COEFFC', 0);
            TOBLOC.Setdouble('BLO_COEFFR', 0);
            TOBLOC.Setdouble('BLO_COEFMARG', 1);
        end;
      end;
      end;
      CalculMontantHtDevLigOuv(ToBLoc, DEV);
      {Correction suite a erreur dans le calcul prix march-- pb d'arrondi }
      //      MontantLoc := arrondi((Qte/QteDuPv) * MontantBasePou,DEV.Decimale );
      if EnPa then MontantLoc := TOBLoc.getValue('BLO_MONTANTPA')
      else if EnHt then MontantLoc := TOBLoc.getValue('BLO_MONTANTHTDEV')
      else MontantLoc := TOBLoc.getValue('BLO_MONTANTTTCDEV');

      MontantLocPr := TOBLoc.getValue('BLO_MONTANTPR');
      if TOBLOc.GetValue('BLO_DPR')<>0 then
      begin
        TOBLOC.Putvalue('BLO_COEFMARG', arrondi(TOBLoc.getValue('BLO_MONTANTHTDEV')/TOBLoc.getValue('BLO_DPR'),4));
        TOBLOC.PutValue('POURCENTMARG',Arrondi((TOBLOC.GetValue('BLO_COEFMARG')-1)*100,2));
      end;
      if TOBLoc.GetValue('BLO_PUHT') <> 0 then
      begin
        TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
      end else
      begin
        TOBLoc.PutValue('POURCENTMARQ',0);
      end;
      // --
      SumLoc := SumLoc + MontantLoc;
      SumPr := SumPr + MontantLocPr;
    end;
    //
    SumLoc := arrondi(sumloc, V_PGI.okdecP);
    SumPr := arrondi(sumPr, V_PGI.okdecP);
    // TEST LS
    if (((not EnPR) and (Arrondi(MontantDemande - SumLoc,V_PGI.OKdecP) <> 0)) or
    										((EnPr) and (Arrondi(MontantPr - SumPr,V_PGI.OKdecP) <> 0))) then
    begin
      EssayeAffecterEcartQte (TOBL,TOBOUvrage,Arrondi(MontantDemande-Sumloc,V_PGI.OkdecP),Arrondi(MontantPr-SumPr,V_PGI.OKDecp),Sumloc,SumPr,
                              IndPa,IndPr,IndPVHTDev,IndPvHt,
                              EnPa,EnPR,EnHt,DEV,false,ForcePaEqualPv);
    end;
    //
    if (atributionEcart) and
      (((not EnPR) and (Arrondi(MontantDemande - SumLoc,V_PGI.OKDecP) <> 0)) or
      ((EnPr) and (Arrondi(MontantPr - SumPr,V_PGI.OKdecP) <> 0))) then
    begin
      if IndiceEcart >= 0 then
      begin
        TOBLOC := TOBOUvrage.detail[IndiceEcart];
        AffecteDifferenceSurLigne (TOBL,TOBLOC,Arrondi(Montantdemande - Sumloc,V_PGI.Okdecp),MontantPr-SumPr,
                                  IndPa,IndPr,IndPVHTDev,IndPvHt,
                                  EnPa,EnPR,EnHt,DEV,ForcePaEqualPv);
      end else
      begin
        if (TobLoc.Detail.count > 0) and
          (((not EnPr) and (Abs(Arrondi(MontantDemande - SumLoc,V_PGI.OKdecP)) > 0.01)) or ((EnPr) and (Abs(MontantPr - SumPr) > 0.01))) then
        begin
          if EnPr then
          begin
            SumLoc := Sumpr / QteDuDetail;
            MontantLoc := MontantPr / QteDuDetail;
            if Qte <> 0 then
            begin
              SumLoc := arrondi(SumLoc / Qte, V_PGI.okdecP);
              MontantLoc := arrondi(MontantLoc / Qte, V_PGI.okdecP);
              // CORRECTIF LS
              (* ReajusteMontantOuvrage (TOBOuvrage,SumLoc,MontantLoc,DEV,EnHT,False); *)
              //                ReajusteMontantOuvrage (TOBPiece,TOBLoc,SumLoc,MontantPr,MontantLoc,DEV,EnHT,False,True,EnPa);
              ReajusteMontantOuvrage(TOBArticles,TOBPiece, TOBL,TOBLoc, SumLoc, MontantPr, MontantLoc, DEV, EnHT, atributionEcart, True, EnPa);
              SumLoc := MontantLoc * QteDuDetail * Qte;
              SumPr := MontantPr * QteDuDetail * Qte;
            end
          end else
          begin
            Sumloc := SumLoc / QteDuDetail;
            MontantLoc := MontantDemande / QteDuDetail;
            if Qte <> 0 then
            begin
              SumLoc := arrondi(SumLoc / Qte, V_PGI.okdecP);
              MontantLoc := arrondi(MontantLoc / Qte, V_PGI.okdecP);
              // CORRECTIF LS
              (* ReajusteMontantOuvrage (TOBOuvrage,SumLoc,MontantLoc,DEV,EnHT,False); *)
              //                ReajusteMontantOuvrage (TOBPiece,TOBLoc,SumLoc,MontantPr,MontantLoc,DEV,EnHT,False,false,EnPa);
              ReajusteMontantOuvrage(TOBArticles,TOBPiece, TOBL,TOBLoc, SumLoc, MontantPr, MontantLoc, DEV, EnHT, atributionEcart, false, EnPa);
              SumLoc := MontantLoc * QteDuDetail * Qte;
              SumPr := MontantPr * QteDuDetail * Qte;
            end;
          end;

        end;
        // On affecte la diffrence sur l'article d'ecart
        if (Arrondi(MontantDemande - SumLoc,V_PGI.OKdecP) <> 0) and (atributionEcart) then
          AttribueEcart(TobPiece,TOBL,TobOuvrage, ArtEcart, MontantDemande - Sumloc, DEV, ENHt, EnPA);
      end;
    end;
    SumLoc := 0;
    SumPr := 0;
    MontantBasePou := 0;
    if (TOBOuvrage.detail.count > 0) and (TOBOuvrage.FieldExists('BLO_MONTANTPA')) then
    begin
      TOBOuvrage.Putvalue('BLO_MONTANTPA', 0);
      TOBOuvrage.Putvalue('BLO_MONTANTHTDEV', 0);
      TOBOuvrage.Putvalue('BLO_MONTANTHT', 0);
      TOBOuvrage.Putvalue('BLO_MONTANTPR', 0);
      TOBOuvrage.Putvalue('BLO_MONTANTTTCDEV', 0);
      TOBOuvrage.Putvalue('BLO_MONTANTTTC', 0);
    end;
    for Indice := 0 to TOBOuvrage.detail.Count - 1 do
    begin
      TOBLoc := TOBOuvrage.detail[Indice];
      if IsVariante(TOBLoc) then continue;

      Qte := TOBLoc.getvaleur(IndQte);
      QteDuPv := TOBLOc.getValeur(IndQteDupv);
      //
      {Correction suite a erreur dans le calcul prix march-- pb d'arrondi }
      //MontantLoc := arrondi((Qte/QteDuPv) * TOBLOC.GetValeur (IndPvHtdev),DEV.Decimale );
      CalculMontantHtDevLigOuv(ToBLOC, DEV);
      if EnPA then MontantLoc := TOBLoc.getValue('BLO_MONTANTPA')
      else if EnHt then MontantLoc := TOBLoc.getValue('BLO_MONTANTHTDEV')
      else MontantLoc := TOBLoc.getValue('BLO_MONTANTTTCDEV');
      MontantPr := TOBLoc.getValue('BLO_MONTANTPR');
      //
      SumLoc := SumLoc + MontantLoc;
      SumPr := SumPr + MontantPr;
      // Stockage de la valeur finale
      TOBLOC.Putvalue('ANCPV', TOBLOC.GetValeur(IndPvHtDev));
      TOBLOC.Putvalue('ANCPA', TOBLOC.GetValeur(IndPa));
      if TOBLOC.GetValeur(IndPr) <> 0 then
      begin
        TOBLOC.Putvalue('BLO_COEFMARG', arrondi(TOBLOC.GetValeur(IndPvHt)/TOBLOC.GetValeur(IndPr),4));
        TOBLOC.PutValue('POURCENTMARG',Arrondi((TOBLOC.GetValue('BLO_COEFMARG')-1)*100,2));
      end;
      if TOBLoc.GetValue('BLO_PUHT') <> 0 then
      begin
        TOBLoc.PutValue('POURCENTMARQ',Arrondi(((TOBLoc.GetValue('BLO_PUHT')- TOBLoc.GetValue('BLO_DPR'))/TOBLoc.GetValue('BLO_PUHT'))*100,2));
      end else
      begin
        TOBLoc.PutValue('POURCENTMARQ',0);
      end;

      // Mise  jour des cumuls
      if TOBOuvrage.fieldExists('BLO_MONTANTPA') then
      begin
        TOBOuvrage.Putvalue('BLO_MONTANTPA', TOBOuvrage.Getvalue('BLO_MONTANTPA') + TOBLOC.GetValue('BLO_MONTANTPA'));
        TOBOuvrage.Putvalue('BLO_MONTANTHTDEV', TOBOuvrage.Getvalue('BLO_MONTANTHTDEV') + TOBLOC.GetValue('BLO_MONTANTHTDEV'));
        TOBOuvrage.Putvalue('BLO_MONTANTHT', TOBOuvrage.Getvalue('BLO_MONTANTHT') + TOBLOC.GetValue('BLO_MONTANTHT'));
        TOBOuvrage.Putvalue('BLO_MONTANTTTCDEV', TOBOuvrage.Getvalue('BLO_MONTANTTTCDEV') + TOBLOC.GetValue('BLO_MONTANTTTCDEV'));
        TOBOuvrage.Putvalue('BLO_MONTANTTTC', TOBOuvrage.Getvalue('BLO_MONTANTTTC') + TOBLOC.GetValue('BLO_MONTANTTTC'));
        TOBOuvrage.Putvalue('BLO_MONTANTPR', TOBOuvrage.Getvalue('BLO_MONTANTPR') + TOBLOC.GetValue('BLO_MONTANTPR'));
      end;
      // --
      if (trim(TOBLoc.getValue('BLO_CODEARTICLE')) = ArtECART) and (TOBLOC.GetValeur(indPvHtDev) = 0) then
        TOBLOC.Free;
    end;
    (*
    if TOBOuvrage.FieldExists('BLO_MONTANTPR') then
    begin
      if TOBOuvrage.Getvalue('BLO_MONTANTPR') <> 0 then
      begin
        TOBOUVRAGE.Putvalue('BLO_COEFMARG', TOBOUVRAGE.GetValue('BLO_MONTANTHT')/TOBOuvrage.Getvalue('BLO_MONTANTPR'));
      end;
    end;
    *)
    MontantDemande := SumLoc;
    MontantPr := SumPr;
  end;
//  MontantDemande := arrondi(MontantDemande / QteDudetail, V_PGI.okdecP);
//  MontantPr := arrondi(MontantPr / QteDudetail, V_PGI.okdecP);
  MontantDemande := ARRONDI(MontantDemande / QteDudetail, V_PGI.OkDecP);
  MontantPr :=  Arrondi(MontantPr / QteDudetail,V_PGI.OkDecP);
end;

procedure CopieOuvFromArt (TOBPiece,TOBLND,TOBART : TOB;DEV:Rdevise; RecupPrix : string='PUH');
var IZone : Integer;
begin
  if TOBART <> nil then
  begin
    for IZone:=1 to length (TZoneArt) do
    begin
       TOBLND.PutValue('BLO_'+TZoneArt[IZone],TOBArt.GetValue('GA_'+TZoneArt[IZone])) ;
    end;
    TOBLND.PutValue('BLO_COEFFG',TOBArt.GetValue('GA_COEFFG')-1) ;
    TOBLND.PutValue('BLO_REMISABLELIGNE','-');
    TOBLND.PutValue('BLO_REMISABLEPIED',TOBArt.GetValue('GA_REMISEPIED')) ;
    TOBLND.PutValue('BLO_ESCOMPTABLE',TOBArt.GetValue('GA_ESCOMPTABLE')) ;
    TOBLND.PutValue('BLO_QUALIFQTEVTE',TOBArt.GetValue('GA_QUALIFUNITEVTE')) ;
    TOBLND.PutValue('BLO_QUALIFQTESTO',TOBArt.GetValue('GA_QUALIFUNITESTO')) ;
    TOBLND.PutValue('BLO_LIBCOMPL',TOBArt.GetValue('GA_LIBCOMPL')) ;
    TOBLND.PutValue('BLO_FOURNISSEUR',TOBArt.GetValue('GA_FOURNPRINC')) ;
    TOBLND.PutValue('BNP_TYPERESSOURCE',TOBArt.GetValue('BNP_TYPERESSOURCE')) ;
    TOBLND.SetBoolean('GESTRELIQUAT',TOBArt.GetBoolean('GA_RELIQUATMT'));

    InitValoDetail (TOBPiece,TOBLND,TOBART,DEV,true,RecupPrix);
  end;
end;

procedure CopieOuvFromLigne(TOBLND,TOBL: TOB);
var IZone : Integer;
begin
for IZone:=1 to length (TZoneGlob) do
    begin
    TOBLND.PutValue('BLO_'+TZoneGlob[IZone],TOBL.GetValue('GL_'+TZoneGlob[IZone])) ;
    end;
TOBLND.PutValue('BLO_DOMAINE',TOBL.GetValue('GL_DOMAINE')) ;
// Cas particulier d'un ouvrage sur le document en variante
// --- Dans ce cas les sous dtail ne peuvent pas tre en variante
  if (IsVariante(TOBL)) and (TOBL.GetValue('GL_TYPEARTICLE')='OUV') then
  begin
    TOBLND.putValue('BLO_TYPELIGNE','ART');
  end;
// --
end;

procedure CopieOuvFromRef(TOBOuv,TOBRef: TOB);
var IZone : Integer;
begin
for IZone:=1 to length (TZoneGlob) do
    begin
    TOBOuv.PutValue('BLO_'+TZoneGlob[IZone],TOBRef.GetValue('BLO_'+TZoneGlob[IZone])) ;
    end;
end;

procedure CopieLigFromRef(TOBLig,TOBRef: TOB);
var IZone : Integer;
    prefixe : string;
begin
  if TOBLig.nomTable = 'LIGNE' then prefixe := 'GL' else prefixe := 'BOP';
  for IZone:=1 to length (TZoneLig) do
  begin
    if TOBLig.fieldExists(prefixe+'_'+TZoneLig[IZone]) then
    begin
      TOBLig.PutValue(prefixe+'_'+TZoneLig[IZone],TOBRef.GetValue('GL_'+TZoneLig[IZone])) ;
    end;
  end;
end;

procedure LoadLesLibDetOuvLig ( TOBPIECE,TOBOuvrage,TOBTiers,TOBAffaire,TOBL,TOBTRFPOC : TOB;var Indice: integer; DEV:RDevise; TheMetreDoc : TMetreArt;  AffSousDetailUnitaire : boolean);
var IndOuv,IndiceNomen,IndiceLig,TypePresent,NiveauImbric : integer;
    TOBL1,TOBOuv,TOBLoc,TT,TL : TOB;
    QteDuDetail,QteDUPv,Montant,MontantUni,MontantAchat,MontantAchatUni,CumMontant,CumMontantUNI,MontantOuv,MontantLig,MontantPaOuv,MontantLPa,QteFact : double;
    IndicePou ,RowRef: Integer;
    EnHt : boolean;
    ArticleOk : string;
    Variante : boolean;
    TheTOBref : TOB;
    ModifSousDetail,ModifAvanc,AfficheSousDetail,WithAvanc,AffSousDetailForModif : boolean;
begin
  WithAvanc := ( Pos(TOBPiece.getValue('GP_NATUREPIECEG'),'FBT;AVC;DAC;FBP;BAC')>0);
  ModifSousDetail :=  (GetParamSocSecur('SO_BTMODIFSDETAIL',false)) ;
  AffSousDetailForModif := ModifSousDetail and (TOBL.GetValue('GLC_VOIRDETAIL')='X');
  AfficheSousDetail := (not ModifSousDetail) AND (TOBL.GetValue('GL_TYPEPRESENT')>DOU_AUCUN);
  ModifAvanc := (Pos(TOBPiece.GetValue('AFF_GENERAUTO'),'AVA;DAC')>0);
  IndicePou := -1;
  EnHT:=(TOBPiece.GetValue('GP_FACTUREHT')='X') ;
  CumMontant := 0; MontantPaOuv := 0; MontantLpa := 0; CumMontantUNI := 0;
  TheTOBRef := nil;
  //
  if (not AffSousDetailForModif) and (not AfficheSousDetail) then BEGIN Inc(Indice);exit; END;
  //
  if (TOBL.GetValue('GL_TYPEARTICLE') = 'OUV') and ((modifSousDetail) or (AfficheSousDetail)) then
  begin
    TypePresent := TOBL.GetValue('GL_TYPEPRESENT');
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    NiveauImbric := TOBL.Getvalue('GL_NIVEAUIMBRIC');
    // VARIANTE
    Variante := IsVariante (TOBL);
    // --
    if IndiceNomen = 0 then
    begin
      inc (indice);
      exit;
    end;
    if (TobOuvrage.detail.count = 0) or ((IndiceNomen -1) > TobOuvrage.detail.count -1) then
    begin
      inc(Indice);
      exit;
    end;
    TOBLoc := TOBOuvrage.detail[IndiceNomen -1];
    if TOBLoc = nil then
    begin
      // Surement une ligne nomenclature
      inc (indice);
      exit;
    end;
    if EnHt then MontantOuv := TOBL.GetValue('GL_MONTANTHTDEV')
            else MontantOuv := TOBL.GetValue('GL_MONTANTTTCDEV');
    MontantPaOuv := TOBL.GetValue('GL_MONTANTPA');
    RowRef := -1;
    MontantLig := 0;
    IndiceLig := Indice+1;
    TOBOUv := TOBLOC.FindFirst (['BLO_TYPEARTICLE'],['POU'],false);
    if TOBOUV <> nil then ArticleOk := TOBOUV.GetValue('BLO_LIBCOMPL');
    for IndOuv := 0 to TOBLoc.detail.count -1 do
    begin
      TOBOUV := TOBLOC.detail[Indouv];
        QteFact := TOBOUV.GetValue('BLO_QTEFACT');
      //end;
      QteDudetail :=TOBOUV.GetValue('BLO_QTEDUDETAIL');
      // fv 02032004
      if QteDUdetail = 0 then qteduDetail := 1;
      QteDuPv :=TOBOUV.GetValue('BLO_PRIXPOURQTE'); if QteDUpv = 0 then qtedupv := 1;
      TOBL1:=NewTobLigne(TOBPiece,IndiceLig+indOuv+1);
      if IndOuv = 0 then
      begin
        TheTOBref := TOBL1;
      end;
      TOBL1.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
      InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,IndiceLig+indOuv+1) ;
      if ModifSousDetail then
      begin
        if (Variante) or (IsVariante(TOBOuv)) then TOBL1.PutValue('GL_TYPELIGNE','SDV')
                                              else TOBL1.PutValue('GL_TYPELIGNE','SD') ;
      end else
      begin
        if (Variante) or (IsVariante(TOBOuv)) then TOBL1.PutValue('GL_TYPELIGNE','COV')
                                              else TOBL1.PutValue('GL_TYPELIGNE','COM') ;
      end;
      TOBL1.PutValue('BNP_TYPERESSOURCE', TOBOUV.GetValue('BNP_TYPERESSOURCE'));
      TOBL1.PutValue('GL_FAMILLETAXE1', TOBOUV.GetValue('BLO_FAMILLETAXE1'));
      TOBL1.PutValue('GL_QUALIFQTEVTE', TOBOUV.GetValue('BLO_QUALIFQTEVTE'));
      TOBL1.PutValue('GL_TYPEARTICLE', TOBOUV.GetValue('BLO_TYPEARTICLE'));
      TOBL1.PutValue('GL_CODEARTICLE', TOBOUV.GetValue('BLO_CODEARTICLE'));
      TOBL1.PutValue('GL_ARTICLE', TOBOUV.GetValue('BLO_ARTICLE'));
      TOBL1.PutValue('GL_COEFMARG', TOBOUV.GetValue('BLO_COEFMARG'));
      TOBL1.PutValue('GL_TYPENOMENC', TOBOUV.GetValue('BLO_TYPENOMENC'));
      TOBL1.PutValue('POURCENTMARG', TOBOUV.GetValue('POURCENTMARG'));
      TOBL1.PutValue('POURCENTMARQ', TOBOUV.GetValue('POURCENTMARQ'));
      TOBL1.PutValue('GL_INDICENOMEN',IndiceNomen);
      TOBL1.PutValue('GL_NIVEAUIMBRIC',NiveauImbric);
      TOBL1.PutValue('UNIQUEBLO', TOBOUV.GetValue('BLO_UNIQUEBLO'));
      TOBL1.PutValue('GLC_NATURETRAVAIL', TOBOUV.GetValue('BLO_NATURETRAVAIL'));
      TOBL1.PutValue('GL_PIECEPRECEDENTE', TOBOUV.GetValue('BLO_PIECEPRECEDENTE'));
      TOBL1.PutValue('GL_TPSUNITAIRE', TOBOUV.GetValue('BLO_TPSUNITAIRE'));
      TOBL1.SetString('GL_RENDEMENT',TOBOUV.GetValue('BLO_RENDEMENT'));
      TOBL1.SetString('GL_PERTE',TOBOUV.GetValue('BLO_PERTE'));
      TOBL1.SetString('GL_RENDEMENT',TOBOUV.GetValue('BLO_RENDEMENT'));
      TOBL1.SetString('GL_QUALIFHEURE', TOBOUV.GetValue('BLO_QUALIFHEURE'));
      // -- POC --
      TOBL1.PutValue('NUMTRANSFERT', TOBOUV.GetValue('NUMTRANSFERT'));
      TOBL1.PutValue('MTTRANSFERT', TOBOUV.GetValue('MTTRANSFERT'));
      //
      if TOBTRFPOC <> nil then
      begin
        TL := TOBTRFPOC.FIndfirst(['BT3_UNIQUEBLO'],[TOBOUV.GetInteger('BLO_UNIQUEBLO')],true);
        if TL <> nil then
        begin
          TL.Data := TOBL1;
        end;
      end;
      if TOBL.getValue('GL_BLOQUETARIF')='X' then
      begin
        TOBL1.putValue('GL_BLOQUETARIF','X');
      end;
      if TOBOUV.GetValue('BLO_NATURETRAVAIL') <> '' then TOBL1.PutValue('LIBELLEFOU', TOBOUV.GetValue('LIBELLEFOU'));
      if TOBOUV.GetValue('BLO_TYPEARTICLE')='POU' then
      begin
        IndicePou := IndiceLig+indOuv+1;
      end else
      begin
        MontantAchat := arrondi(TOBOUV.GetValue('BLO_DPA')* TOBL.getvalue('GL_QTEFACT') * (QteFact/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP);
        MontantAchatUni  := arrondi(TOBOUV.GetValue('BLO_DPA')* (QteFact/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP);
        if EnHt then
        begin
          	Montant := arrondi(TOBOUV.GetValue('BLO_PUHTDEV')* TOBL.getvalue('GL_QTEFACT') * (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
          MontantUni := arrondi(TOBOUV.GetValue('BLO_PUHTDEV')* (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
        end else
        begin
          Montant := arrondi(TOBOUV.GetValue('BLO_PUTTCDEV')* TOBL.getvalue('GL_QTEFACT') * (QteFact/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
          MontantUni := arrondi(TOBOUV.GetValue('BLO_PUTTCDEV')* (QteFact/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
        end;
        // VARIANTE
        if ((not IsVariante (TOBL1)) and (not IsVariante (TOBL))) or (IsVariante (TOBL1) and (IsVariante (TOBL)))then
        begin
          MontantLig := MontantLig + Montant;
          MontantLPA := MontantLPa + MontantAchat;
        end;
        // --
        if ArticleOKInPOUR (TOBOUV.getValue('BLO_TYPEARTICLE'),ArticleOk) then
        begin
          CumMontant := CumMontant + Montant;
          CumMontantUNI := CumMontantUNI + MontantUNI;
      end;
      end;
      if ((TypePresent and DOU_CODE) = DOU_CODE) or (ModifSousDetail)  then
      BEGIN
        TOBL1.PutValue('GL_REFARTSAISIE',TOBOUV.GetValue('BLO_CODEARTICLE'));
      END;
      TOBL1.PutValue('GL_REFARTTIERS',TOBOuv.GetValue('BLO_ARTICLE'));
      if ((TypePresent and DOU_LIBELLE) = DOU_LIBELLE) or (ModifSousDetail)  then TOBL1.PutValue('GL_LIBELLE',TOBOUV.GetValue('BLO_LIBELLE'));
      if ((typepresent and DOU_QTE) = DOU_QTE) or (ModifSousDetail)  then
      begin
        if TOBOUV.GetValue('BLO_TYPEARTICLE')='POU' then
          TOBL1.PutValue('GL_QTEFACT',TOBOUV.GetValue('BLO_QTEFACT')/QteDuDetail)
        else
        begin
          if AffSousDetailUnitaire then
          begin
	          TOBL1.PutValue('GL_QTEFACT',TOBOUV.GetValue('BLO_QTEFACT'));
	          TOBL1.PutValue('GL_QTESAIS',TOBOUV.GetValue('BLO_QTESAIS'));
          end else
          begin
	          TOBL1.PutValue('GL_QTEFACT',(TOBOUV.GetValue('BLO_QTEFACT')* TOBL.GetValue('GL_QTEFACT'))/QteDuDetail);
	          TOBL1.PutValue('GL_QTESAIS',(TOBOUV.GetValue('BLO_QTESAIS')* TOBL.GetValue('GL_QTEFACT'))/QteDuDetail);
				end;
				end;
        // --- GUINIER ---
        TOBL1.PutValue('GL_QTERESTE',TOBL1.getValue('GL_QTEFACT'));
        if CtrlOkReliquat(TOBL1, 'GL') then TOBL1.PutValue('GL_MTRESTE',TOBL1.getValue('GL_MONTANTHTDEV'));

        //
        TOBL1.PutValue('GL_QTESTOCK',TOBL1.getValue('GL_QTEFACT'));
      end;
      if ((typepresent and DOU_UNITE) = DOU_UNITE) or (ModifSousDetail)  then TOBL1.PutValue('GL_QUALIFQTEVTE',TOBOUV.GetValue('BLO_QUALIFQTEVTE'));
      if ((typepresent and DOU_PU) = DOU_PU) or (ModifSousDetail) then
      begin
        if TOBOUV.GetValue('BLO_TYPEARTICLE')<>'POU' then
        begin
          TOBL1.PutValue('GL_DPA', TOBOuv.GetValue('BLO_DPA') / QteDupv);
          if EnHt then TOBL1.PutValue('GL_PUHTDEV',Arrondi(TOBOUV.GetValue('BLO_PUHTDEV')/QteDupv,V_PGI.okdecP))
                  else TOBL1.PutValue('GL_PUTTCDEV',Arrondi(TOBOUV.GetValue('BLO_PUTTCDEV')/QteDupv,V_PGI.okdecP));
        end;
      end;
      if ((typepresent and DOU_MONTANT) = DOU_MONTANT) or (ModifSousDetail) then
      begin
        if TOBOUV.GetValue('BLO_TYPEARTICLE')<>'POU' then
        begin
          if AffSousDetailUnitaire then
          begin
            TOBL1.PUtValue('GL_MONTANTPA',MontantAchatUNI);
            if EnHt then TOBL1.PUtValue('GL_MONTANTHTDEV',MontantUNI)
                    else TOBL1.PUtValue('GL_MONTANTTTCDEV',MontantUNI);
          end else
          begin
          TOBL1.PUtValue('GL_MONTANTPA',MontantAchat);
          if EnHt then TOBL1.PUtValue('GL_MONTANTHTDEV',Montant)
                  else TOBL1.PUtValue('GL_MONTANTTTCDEV',Montant);
        end;
      end;
      end;
      if WithAvanc then
      begin
        TOBL1.putValue('BLF_NATUREPIECEG',TOBOUV.getValue('BLF_NATUREPIECEG'));
        TOBL1.putValue('BLF_MTMARCHE',TOBOUV.getValue('BLF_MTMARCHE'));
        TOBL1.putValue('BLF_MTDEJAFACT',TOBOUV.getValue('BLF_MTDEJAFACT'));
        TOBL1.putValue('BLF_MTSITUATION',TOBOUV.getValue('BLF_MTSITUATION'));
        TOBL1.putValue('BLF_MTCUMULEFACT',TOBOUV.getValue('BLF_MTCUMULEFACT'));
        TOBL1.putValue('BLF_QTEMARCHE',TOBOUV.getValue('BLF_QTEMARCHE'));
        TOBL1.putValue('BLF_QTEDEJAFACT',TOBOUV.getValue('BLF_QTEDEJAFACT'));
        TOBL1.putValue('BLF_QTESITUATION',TOBOUV.getValue('BLF_QTESITUATION'));
        TOBL1.putValue('BLF_QTECUMULEFACT',TOBOUV.getValue('BLF_QTECUMULEFACT'));
        TOBL1.putValue('BLF_POURCENTAVANC',TOBOUV.getValue('BLF_POURCENTAVANC'));
      end;
    	TOBL1.putvalue('GL_HEURE',Arrondi(TOBL1.getValue('GL_TPSUNITAIRE')*TOBL1.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
    	TOBL1.putvalue('GL_TOTALHEURE',TOBL1.Getvalue('GL_HEURE'));
      RecupSousDetailSurLigne (TOBL1,TOBOUV);
      //la logique voudrait que l'on colle la copie des métrés ici !!!!!!!
      //problème on ne peut faire des appels circulaire...
      //il faut donc dupliquer dans ce source ce qui a été fait dans copiercollerutil...
      //c'est pas super malin mais pas d'autres solution !!!!!!!!!!
      //if TheMetreDoc <> nil then
      //begin
      //  ColleTobVarDoc(TOBL1, TOBOUV, TOBL1, TheMetreDoc.TTobVarDoc);
      //  ColleTobMetres(TOBL1, TOBOUV, TOBL1, TheMetreDoc);
      //end;
      //
    end;
    indice := Indice + TOBLoc.detail.count+1;
  end else
  begin
  	inc (indice);
  end;
  if IndicePou <> -1 then
  begin
    TOBOUv := TOBLOC.FindFirst (['BLO_TYPEARTICLE'],['POU'],false);
    TOBL1 := GetTOBLIgne (TOBPiece,IndicePou);
    Montant := arrondi(CumMontant * (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
    MontantUNI := arrondi(CumMontantUNI * (TOBOUV.GetValue('BLO_QTEFACT')/ (QteDuPv * QteDuDetail)),V_PGI.OkdecP );
    MontantLig := MontantLig + Montant;
    if ((typepresent and DOU_PU) = DOU_PU) or (ModifSousDetail) then
    begin
      if AffSousDetailUnitaire then
      begin
        if EnHt then TOBL1.PutValue('GL_PUHTDEV',CumMontantUNI)
                else TOBL1.PutValue('GL_PUTTCDEV',CumMontantUNI);
      end else
      begin
        if AffSousDetailUnitaire then
        begin
          if EnHt then TOBL1.PutValue('GL_PUHTDEV',CumMontantUNI)
                  else TOBL1.PutValue('GL_PUTTCDEV',CumMontantUNI);
        end else
        begin
    if EnHt then TOBL1.PutValue('GL_PUHTDEV',CumMontant)
            else TOBL1.PutValue('GL_PUTTCDEV',CumMontant);
        end;
      end;
    end;
    if ((typepresent and DOU_MONTANT) = DOU_MONTANT) or (ModifSousDetail) then
    begin
      if AffSousDetailUnitaire then
      begin
        if EnHt then TOBL1.PUtValue('GL_MONTANTHTDEV',MontantUNI)
                else TOBL1.PUtValue('GL_MONTANTTTCDEV',MontantUNI);
      end else
      begin
      if EnHt then TOBL1.PUtValue('GL_MONTANTHTDEV',Montant)
              else TOBL1.PUtValue('GL_MONTANTTTCDEV',Montant);
    end;
  end;
  end;
  // Rajustement total des lignes / Montant Ouvrage
  if (TheTOBRef <> nil) and ((MontantLig <> MontantOuv) or (MontantLPa <> MontantPaOuv))and ((typepresent and DOU_MONTANT) = DOU_MONTANT) and  (not ModifSousDetail) and (not AffSousDetailUnitaire) then
  begin
    TOBL1 := TheTOBref;
    if MontantOuv <> 0 then
    begin
      if EnHt then TOBL1.PutValue ('GL_MONTANTHTDEV',TOBL1.GetValue('GL_MONTANTHTDEV')+(MontantOuv-MontantLig))
              else TOBL1.PutValue ('GL_MONTANTTTCDEV',TOBL1.GetValue('GL_MONTANTTTCDEV')+(MontantOuv-MontantLig)) ;
    end;
    if MontantPaOuv <> 0 then
    begin
      TOBL1.PutValue ('GL_MONTANTPA',TOBL1.GetValue('GL_MONTANTPA')+(MontantPaOuv-MontantLPa))
    end;
  end;

end;

procedure LoadLesLibDetailOuvrages (TOBPIECE,TOBOuvrage,TOBTiers,TOBAffaire,TOBTRFPOC : TOB;DEV:RDevise; TheMetreDoc : TmetreArt);
var indice : integer;
    TOBL : TOB;
    VoirDetail : boolean;
    TypePresent : integer;
    ModifSousDetail,ModifLeSousDetail,AffSousDetail,fAffSousDetailUnitaire : boolean;
begin
  ModifSousDetail :=  (GetParamSocSecur('SO_BTMODIFSDETAIL',false));
  fAffSousDetailUnitaire := GetParamSocSecur('SO_BTAFFSDUNITAIRE', False);

  if TOBPiece.detail.count = 0 then Exit;
  Indice := 0;
  repeat
    TOBL := TOBPiece.detail[indice];
    if TOBL.GetValue('GL_INDICENOMEN')= 0 then
    begin
      inc(indice);
      Continue;
    end;
    ModifLeSousDetail := (ModifSousDetail) and (TOBL.getValue('GLC_VOIRDETAIL')='X');
    TypePresent := TOBL.GetValue('GL_TYPEPRESENT');
    VoirDetail := (TOBL.GetValue('GLC_VOIRDETAIL')='X');
    AffSousDetail := (not ModifSousDetail) and (TypePresent <> DOU_AUCUN);
    if (not ModifLeSousDetail) and (not AffSousDetail) then begin inc(Indice); continue; end;
{    if (not ModifLeSousDetail) and (not VoirDetail) then begin inc(Indice); continue; end;}
    LoadLesLibDetOuvLig ( TOBPIECE,TOBOuvrage,TOBTiers,TOBAffaire,TOBL,TOBTRFPOC ,Indice,DEV, TheMetreDoc,fAffSousDetailUnitaire);
  until indice > TOBPiece.detail.count -1;
end;

procedure DropLesLibDetailouvrages (TOBPiece : TOB);
var indice : integer;
    TOBL : TOB;
    IndiceNomen : Integer;
    TypeLigne : string;
begin
  if TOBPiece.detail.count = 0 then Exit;
  Indice := 0;
  repeat
    TOBL := TOBPiece.detail[indice];
    TypeLigne := TOBL.GetValue('GL_TYPELIGNE');
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    if (Pos(TypeLigne,'SD;SDV;COM;COV')>0) and (IndiceNomen>0) then
    begin
    	TOBL.Free;
    end else inc(indice);
  until indice > TOBPiece.detail.count -1;
end;


procedure OuvrageVersLigne (TOBL,TOBINS,TOBOuv : TOB;Qte,QteDuDetail : double);
var IndZone,ITableLig,Ind : Integer;
    NomZone,NomChamp : String;
    PMAP,PMRP,DPA,DPR : double;
    prefixe : string;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  prefixe := TableToPrefixe(TOBINS.NomTable);
  if TOBIns.NomTable = 'LIGNE' Then iTableLig := PrefixeToNum ('GL')
  else if TOBIns.NomTable = 'LIGNEOUVPLAT' Then iTableLig := PrefixeToNum ('BOP');
  FieldList := MCD.GetTable(TOBINS.NomTable).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomZone := (FieldList.Current as IFieldCOM).name;
    ind := Pos ('_',NomZone);
    NomChamp := copy (NomZone,ind+1,255);
    if TOBOuv.FieldExists ('BLO_'+Nomchamp) then
      TOBIns.PutValue(NomZone,TOBOuv.getvalue('BLO_'+NomChamp));
  end;
  // Copie du reste des infos par l'intermdiaire de la TOBL
  CopieLigFromRef (TOBIns,TOBL);
  // --
  TOBIns.putValue('BNP_TYPERESSOURCE',TOBOUV.getValue('BNP_TYPERESSOURCE'));
  for Ind:= 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    if TOBOUV.fieldExists('MILLIEME'+IntToStr(Ind+1)) then
    begin
      TOBIns.PutValue ('MILLIEME'+IntToStr(Ind+1),TOBOUV.GetValue('MILLIEME'+IntToStr(Ind+1)));
    end;
  end;
  if Prefixe = 'GL' then
  begin
    TOBIns.putValue('GLC_NONAPPLICFRAIS',TOBOUV.getValue('BLO_NONAPPLICFRAIS'));
    TOBIns.putValue('GLC_NONAPPLICFC',TOBOUV.getValue('BLO_NONAPPLICFC'));
    TOBIns.putValue('GLC_NONAPPLICFG',TOBOUV.getValue('BLO_NONAPPLICFG'));
    TOBIns.putValue('GLC_NATURETRAVAIL',TOBOUV.getValue('BLO_NATURETRAVAIL'));
  end;
  // --
end;

procedure AddLesReferences (TOBIns,TOBLoc : TOB);
var prefixe : string;
begin
  prefixe := GetprefixeTable (TOBINs);
  if TOBLOC.FieldExists ('BLO_N1') then
  begin
    (*
    TOBINS.addChampsupValeur ('N1',TOBLOC.GetValue('BLO_N1'));
    TOBINS.addChampsupValeur ('N2',TOBLOC.GetValue('BLO_N2'));
    TOBINS.addChampsupValeur ('N3',TOBLOC.GetValue('BLO_N3'));
    TOBINS.addChampsupValeur ('N4',TOBLOC.GetValue('BLO_N4'));
    TOBINS.addChampsupValeur ('N5',TOBLOC.GetValue('BLO_N5'));
    *)
    TOBINS.PutValue (prefixe+'_N1',TOBLOC.GetValue('BLO_N1'));
    TOBINS.PutValue (prefixe+'_N2',TOBLOC.GetValue('BLO_N2'));
    TOBINS.PutValue (prefixe+'_N3',TOBLOC.GetValue('BLO_N3'));
    TOBINS.PutValue (prefixe+'_N4',TOBLOC.GetValue('BLO_N4'));
    TOBINS.PutValue (prefixe+'_N5',TOBLOC.GetValue('BLO_N5'));
  end else
  begin
    TOBINS.PutValue (prefixe+'_N1',0);
    TOBINS.PutValue (prefixe+'_N2',0);
    TOBINS.PutValue (prefixe+'_N3',0);
    TOBINS.PutValue (prefixe+'_N4',0);
    TOBINS.PutValue (prefixe+'_N5',0);
  end;
end;

procedure AplatOuvrage (TOBpiece,TOBL,TOBOuv,TOBPlat: TOB;Qte,QteDuDetail: double;TousNiveaux:boolean;GardeReference:boolean=false; WithCalcul : boolean=true; SurLigneStd : boolean=false; PourCpta : boolean=false);

	procedure  AjusteMontantFacOuv (TOBOP : TOB;MontantGlob : Double; DEV : Rdevise; SurLigneStd,EnHt : boolean);
  var indice : Integer;
  		MtMax,MtCum,mtDif : double;
      iMtMax : Integer;
      Prefixe : string;
      TOBOPD : TOB;
  begin
    if SurLigneStd then
    begin
      prefixe := 'GL';
    end else
    begin
      prefixe := 'BOP';
    end;

    MtMax := 0;
    iMtMax := -1;
		for Indice := 0 to TOBOP.detail.Count -1 do
    begin
			TOBOPD := TOBOP.detail[Indice];
      if TOBOPD.GetDouble('BLF_MTSITUATION') > MtMax then
      begin
        MtMax := TOBOPD.GetDouble('BLF_MTSITUATION');
        iMtMax := indice;
      end;
      MtCum := MtCum + TOBOPD.GetDouble('BLF_MTSITUATION');
    end;
    if iMtMAx >= 0 then
    begin
      mtDif := ARRONDI(MontantGlob,DEV.Decimale) - ARRONDI(MtCum,DEV.decimale);
      if mtDif <> 0 then
      begin
        TOBOPD := TOBOP.detail[ImtMax];
  			TOBOPD.SetDouble('BLF_MTSITUATION',TOBOPD.GetDouble('BLF_MTSITUATION')+MtDif);
        if EnHt then
        begin
          TOBOPD.Putvalue (prefixe+'_MONTANTHTDEV',TOBOPD.getvalue('BLF_MTSITUATION'));
        end else
        begin
          TOBOPD.Putvalue (prefixe+'_MONTANTTTCDEV',TOBOPD.getvalue('BLF_MTSITUATION'));
        end;
      end;
    end;
  end;

  procedure ReintegreDetailOuvPlat(TOBPlat,TOBOP : TOB);
  begin
		repeat
      TOBOP.detail[0].ChangeParent(TOBPlat,-1);
    until TOBOP.Detail.count =0;
  end;

var Indice,ind1 : Integer;
    TOBLOc,TOBIns,TOBOP : TOB;
    QTESuite,QTEDuDetailSuite:double;
    DEV : Rdevise;
    NumOrdre : integer;
    prefixe : string;
    EnHt : boolean;
    isBLF,ForceEnregBLF,IsAvanc : boolean;
begin
  if TobOuv = nil then exit;
  isAvanc := (Pos(TOBpiece.GetString('AFF_GENERAUTO'),'AVA;DAC')>0);
  ForceEnregBLF := false;
  TOBOP := TOB.Create ('UN OUVRAGE A PLAT',nil,-1);
  //
  TRY
    EnHt := (TOBPiece.GetValue('GP_FACTUREHT')='X');
  //  TOBPiece := TOBL.Parent;
    DEV.Code:=TOBL.GetValue('GL_DEVISE') ; GetInfosDevise(DEV) ;
    for Indice := 0 to TOBOuv.detail.count -1 do
    begin
      TOBLOC := TOBOuv.detail[Indice];
      if TOBLOc.Parent.NomTable = 'LIGNEOUV' then
      begin
        ForceEnregBLF := True;
      end;
      //
      if (TOBLOC.FieldExists ('BLF_NATUREPIECEG')) and (TOBLOC.getvalue('BLF_NATUREPIECEG')<> '') then
      begin
        isBlf := True;
        QteDuDetailSuite := 1;
        QteSuite := TOBLOC.getvalue('BLF_QTESITUATION');
      end else
      begin
        isBLF := false;
        if TOBLOC.GetValue('BLO_QTEDUDETAIL') <> 0 then
          QteDuDetailSuite := QteDudetail * TOBLOC.GetValue('BLO_QTEDUDETAIL')
        else
          QteDuDetailSuite := QteDudetail;

        QteSuite := Qte * TOBLOc.GetValue('BLO_QTEFACT');
      end;
      //
      if (TOBLoc.detail.count > 0) and (TousNiveaux) then
      begin
        TOBOP.ClearDetail;
        // Le sous dtail est lui mme un ouvrage
        APlatOuvrage (TOBPiece,TOBL,TOBLoc,TOBOP,QTeSuite,QteDuDetailSuite,TousNiveaux,GardeReference,WithCalcul,SurLigneStd,pourCpta);
        if isBLF then
        begin
          AjusteMontantFacOuv (TOBOP,TOBLOc.GetDouble('BLF_MTSITUATION'),DEV,SurLigneStd,EnHt);
        end;
        ReintegreDetailOuvPlat(TOBPlat,TOBOP);
      end else
      begin
        if SurLigneStd then
        begin
          prefixe := 'GL';
        end else
        begin
          prefixe := 'BOP';
        end;
        NumOrdre := TOBL.getValue('GL_NUMORDRE');
        TOBIns:=NewTobLignePlat(TOBPLat,NumOrdre,-1,SurLigneSTd);
        InitialiseLigne (TOBIns,TOBL.GetValue('GL_NUMLIGNE'),1);
        OuvrageVersLigne (TOBL,TOBINS,TOBLoc,Qte,QteDuDetail);
        TOBIns.SetInteger('UNIQUEBLO',TOBLoc.GetInteger('BLO_UNIQUEBLO'));

        if (TOBLOC.FieldExists ('BLF_NATUREPIECEG')) and
           (TOBLOC.getvalue('BLF_NATUREPIECEG')<> '') then
        begin
          TOBIns.putValue('BLF_NATUREPIECEG',TOBLoc.getvalue('BLF_NATUREPIECEG'));
        end;
        //
        if (TOBLOC.FieldExists ('BLF_NATUREPIECEG')) and (TOBLOC.getvalue('BLF_NATUREPIECEG')<> '') then
        begin
          //
          if (PourCpta) and (valeur(TOBLoc.GetValue('BLO_NATURETRAVAIL')) >0) then
          begin
            // Pour les sous traitants, les avancements sont en PA... il faut donc recalculer en PV pour la compta
            if IsAvanc then
            begin
            	TOBINS.putValue('BLF_MTSITUATION',Arrondi(TOBLOc.getDouble('BLF_MTMARCHE') * TOBLoc.getDouble('BLF_POURCENTAVANC') /100-TOBLOc.getDouble('BLF_MTDEJAFACT'),DEV.decimale));
            end else
            begin
            	TOBINS.putValue('BLF_MTSITUATION',Arrondi(TOBLOc.getDouble('BLF_MTMARCHE') * TOBLoc.getDouble('BLF_POURCENTAVANC') /100,DEV.decimale));
            end;
            TOBINS.putValue('BLF_QTESITUATION',TOBLOC.getvalue('BLF_QTESITUATION'));
            TOBIns.Putvalue (prefixe+'_QTEFACT',TOBLOC.getvalue('BLF_QTESITUATION'));
            if EnHt then
            begin
              TOBIns.Putvalue (prefixe+'_MONTANTHTDEV',TOBINS.getvalue('BLF_MTSITUATION'));
            end else
            begin
              TOBIns.Putvalue (prefixe+'_MONTANTTTCDEV',TOBINS.getvalue('BLF_MTSITUATION'));
            end;
          end else
          begin
            TOBINS.putValue('BLF_MTSITUATION',TOBLOC.getvalue('BLF_MTSITUATION'));
            TOBINS.putValue('BLF_QTESITUATION',TOBLOC.getvalue('BLF_QTESITUATION'));
            TOBIns.Putvalue (prefixe+'_QTEFACT',TOBLOC.getvalue('BLF_QTESITUATION'));
            if EnHt then
            begin
              TOBIns.Putvalue (prefixe+'_MONTANTHTDEV',TOBLOC.getvalue('BLF_MTSITUATION'));
            end else
            begin
              TOBIns.Putvalue (prefixe+'_MONTANTTTCDEV',TOBLOC.getvalue('BLF_MTSITUATION'));
            end;
          end;
        end else
        begin
          TOBIns.Putvalue (prefixe+'_QTEFACT',QteSuite/QTeDudetailSuite);
        end;
        //
        if not SurLigneStd then
        begin
          TOBIns.putValue('BOP_NUMORDRE',TOBL.GetValue('GL_NUMORDRE'));
          TOBIns.putValue('BOP_TYPELIGNE',TOBL.GetValue('GL_TYPELIGNE'));
          TOBIns.putValue('GLC_DOCUMENTLIE',TOBL.GetValue('GLC_DOCUMENTLIE'));
        end;
        // tentative de reolution du probleme des decimales
        TOBIns.SetDouble (prefixe+'_PUHTDEV',Arrondi(TOBIns.GetDouble (prefixe+'_PUHTDEV'),DEV.decimale));
        TOBIns.PutValue(prefixe+'_PUHT',Arrondi(TOBIns.GetValue(prefixe+'_PUHT'),V_PGI.okdecP));
        // Correction anomalie li a la gestion des dimensions de mode
        TOBIns.PutValue(prefixe+'_TYPEDIM','NOR');
        // ---------------
        // On reprend pour les lignes de sous-dtails la remise pied
        // et l'escompte de la ligne ouvrage
        if not SurLigneStd then
        begin
          TOBIns.Putvalue ('BOP_REMISEPIED',TOBL.GetValue('GL_REMISEPIED'));
          TOBIns.Putvalue ('BOP_ESCOMPTE',TOBL.GetValue('GL_ESCOMPTE'));
        end;
        // recalcul des montant de frais en fonction dela quantite de l'ouvrage
        CalculeLigneAc (TOBIns,DEV,false);
        // correction FQ 12010
        TOBINS.putValue('BOP_ESCOMPTABLE',TOBL.GetValue('GL_ESCOMPTABLE'));
        TOBINS.putValue('BOP_REMISABLELIGNE',TOBL.GetValue('GL_REMISABLELIGNE'));
        TOBINS.putValue('BOP_REMISELIGNE',TOBL.GetValue('GL_REMISELIGNE'));
        TOBINS.putValue('BOP_REMISABLEPIED',TOBL.GetValue('GL_REMISABLEPIED'));
        // correction FQ 12005 --> mise en place de la tva de l'ouvrage dans le cas ou celui ci est comptabilise au niveau ouvrage
        if (TOBL.GetValue('GL_TYPENOMENC')='OU1') then
        begin
          TOBINS.putValue('BOP_REGIMETAXE',TOBL.GetValue('GL_REGIMETAXE'));
          for Ind1 := 1 to 5 do
          begin
            TOBINS.putValue('BOP_FAMILLETAXE'+InttOStr(Ind1),TOBL.GetValue('GL_FAMILLETAXE'+InttOStr(Ind1)));
          end;
        end;
        if GardeReference then
        begin
          AddLesReferences (TOBIns,TOBLoc);
        end;
        //TOBIns.putValue('GL_RECALCULER','X');
        if not SurLIgneStd then
        begin
          TOBIns.putValue('BOP_BLOQUETARIF',TOBL.GetValue('GL_BLOQUETARIF'));
        end;
        if WithCalcul then
        begin
          if EnHt then CalculeLigneHT (TOBIns,nil,TOBPiece,DEV,DEV.decimale)
                  else CalculeLigneTTC (TOBIns,nil,TOBPiece,DEV,DEV.decimale);
        end;
        if ForceEnregBLF then
        begin
          if (TOBLoc.parent.FieldExists('BLF_NATUREPIECEG')) AND (TobIns.FieldExists('BLF_NATUREPIECEG')) Then TOBIns.putValue('BLF_NATUREPIECEG',TOBLoc.parent.getvalue('BLF_NATUREPIECEG'));
          //
          if EnHt then
          begin
            if TobIns.FieldExists('BLF_MTSITUATION') then	TOBIns.Putvalue('BLF_MTSITUATION',TOBINS.getvalue(prefixe+'_MONTANTHTDEV') );
          end
          else
          begin
            if TobIns.FieldExists('BLF_MTSITUATION') then	TOBIns.Putvalue('BLF_MTSITUATION',TOBINS.getvalue(prefixe+'_MONTANTTTCDEV') );
        end;
      end;
        StockeMontantTypeSurLigne (TOBIns);
    end;
    end;
  FINALLY
  	TOBOP.Free;
  END;
end;

procedure MiseAPlatOuv (TOBpiece,TOBL,TOBOuv,TOBPlat : TOB; TousNiveaux:boolean; GardeReference : boolean=false; WithCalcul : boolean=true; SurLigneStd : boolean=false;PourCpta : boolean=false);
var IndiceNomen: Integer;
    TOBLoc : TOB;
    Qte: double;
begin
IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
if IndiceNomen = 0 then exit;
//FV1 : 28/07/2016
if TOBOuv.detail.count = 0 then exit;
if Indicenomen-1 > TOBOuv.detail.count - 1 then exit;
//
TOBLoc := TOBOuv.detail[IndiceNomen -1];
Qte:= TOBL.GetValue('GL_QTEFACT');
APlatOuvrage (TOBPiece,TOBL,TOBLoc,TOBPlat,Qte,1,TousNiveaux,GardeReference,WithCalcul,SurLigneStd,PourCpta);
end;

procedure copyLigne (TOBIns,TOBLoc : TOB);
var itableLig,Indzone,Ind : integer;
    NomZone,NomCHamp : string;
    prefixeDest,PrefixeSrc : string;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  prefixeSrc := GetPrefixeTable (TOBLoc);
  prefixeDest := GetPrefixeTable (TOBIns);
  //
  FieldList := MCD.GetTable(TOBINS.NomTable).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    NomZone := (FieldList.Current as IFieldCOM).name;
    ind := Pos ('_',NomZone);
    NomChamp := prefixeDest +'_'+ copy (NomZone,ind+1,255);
    if TOBIns.FieldExists (NomChamp) then
      TOBIns.PutValue(NomChamp,TOBLoc.getvalue(NomZone));
  end;
//
  for Ind:= 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
  	TOBIns.PutValue ('MILLIEME'+IntToStr(Ind+1),TOBLoc.GetValue('MILLIEME'+IntToStr(Ind+1)));
  end;
  TOBIns.PutValue ('BNP_TYPERESSOURCE',TOBLoc.GetValue('BNP_TYPERESSOURCE'));
  TOBIns.PutValue (PrefixeDest+'_MONTANTPA',TOBLoc.GetValue(PrefixeSrc+'_MONTANTPA'));
  if (TOBInS.FieldExists (PrefixeDest+'_MONTANTPR')) and (TOBInS.FieldExists (PrefixeSrc+'_MONTANTPR')) then
  begin
    TOBIns.PutValue (PrefixeDest+'_MONTANTPR',TOBLoc.GetValue(PrefixeSrc+'_MONTANTPR'));
  end;
  if (TOBLoc.NomTable = 'LIGNE') and (TOBIns.NomTable = 'LIGNE') then
  begin
    TOBIns.PutValue ('GLC_NONAPPLICFRAIS',TOBLoc.GetValue('GLC_NONAPPLICFRAIS'));
    TOBIns.PutValue ('GLC_NONAPPLICFC',TOBLoc.GetValue('GLC_NONAPPLICFC'));
    TOBIns.PutValue ('GLC_NONAPPLICFG',TOBLoc.GetValue('GLC_NONAPPLICFG'));
    TOBIns.PutValue ('GLC_NATURETRAVAIL',TOBLoc.GetValue('GLC_NATURETRAVAIL'));
  end;
end;


procedure changeRegimeOuvInterne (TOBOuvrage : TOB;Regime : string);
var indice : Integer;
    TOBOuv : TOB;
begin
if TOBOuvrage = nil then exit;
TOBOuvrage.SetString('GL_REGIMETAXE',regime);
if TOBOuvrage.detail.count = 0 then exit;
for Indice := 0 To TOBOuvrage.detail.count - 1 do
    begin
    TOBOuv:=TOBOuvrage.Detail[Indice] ;
    changeRegimeOuvInterne (TOBOuv,Regime);
    end;
end;

procedure changeTaxeOuvInterne (TOBOuvrage : TOB;Taxe1,Taxe2,Taxe3,Taxe4,Taxe5 : string);
var indice : Integer;
    TOBOuv : TOB;
begin
if TOBOuvrage = nil then exit;
if TOBOuvrage.FieldExists ('BLO_FAMILLETAXE1') then TOBOuvrage.PutValue ('BLO_FAMILLETAXE1',Taxe1);
if TOBOuvrage.FieldExists ('BLO_FAMILLETAXE2') then TOBOuvrage.PutValue ('BLO_FAMILLETAXE2',Taxe2);
if TOBOuvrage.FieldExists ('BLO_FAMILLETAXE3') then TOBOuvrage.PutValue ('BLO_FAMILLETAXE3',Taxe3);
if TOBOuvrage.FieldExists ('BLO_FAMILLETAXE4') then TOBOuvrage.PutValue ('BLO_FAMILLETAXE4',Taxe4);
if TOBOuvrage.FieldExists ('BLO_FAMILLETAXE5') then TOBOuvrage.PutValue ('BLO_FAMILLETAXE5',Taxe5);
if TOBOuvrage.detail.count = 0 then exit;
for Indice := 0 To TOBOuvrage.detail.count - 1 do
    begin
    TOBOuv:=TOBOuvrage.Detail[Indice] ;
    changeTaxeOuvInterne (TOBOuv,Taxe1,Taxe2,Taxe3,Taxe4,Taxe5);
    end;
end;

procedure AppliqueChangeRegimeOuv (TOBPiece,TOBOuvrage : TOB; Ligne : integer; newCode : string);
var TOBL,TOBOuv : TOB;
    indicenomen : Integer;
begin
  TOBL := GetTOBLIgne (TOBPiece,Ligne);
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if (TOBOuvrage = nil) or (IndiceNomen = 0) then exit;
  //FV1 : 28/07/2016
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen-1 > TobOuvrage.detail.count - 1 then exit;

  TOBOuv:=TOBOuvrage.Detail[IndiceNomen-1] ;
  changeRegimeOuvInterne (TOBOuv,NewCode);
end;

procedure AppliqueChangeTaxeOuv (TOBPiece,TOBOuvrage : TOB;Ligne:integer;Taxe1,Taxe2,Taxe3,Taxe4,Taxe5 : string);
var TOBL,TOBOuv : TOB;
    indicenomen : Integer;
begin
  TOBL := GetTOBLIgne (TOBPiece,Ligne);
  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if (TOBOuvrage = nil) or (IndiceNomen = 0) then exit;

  //FV1 : 28/07/2016
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen -1 > TobOuvrage.detail.count - 1 then exit;

  TOBOuv:=TOBOuvrage.Detail[IndiceNomen-1] ;
  changeTaxeOuvInterne (TOBOuv,Taxe1,Taxe2,Taxe3,Taxe4,Taxe5);
end;

procedure AffecteValoFromOuv (TOBOuv : TOB;DEV : RDevise;var Valeurs : T_Valeurs);
var indice : Integer;
    TOBLoc: TOB;
    Req : TQuery;
    Valloc : T_Valeurs;
begin
for Indice := 0 To TOBOuv.detail.count - 1 do
    begin
    TOBLoc := TOBOUV.Detail[Indice];
    if TOBLOC.GetValue('BLO_TYPEARTICLE') = 'OUV' Then
       begin
       AffecteValoFromOuv (TOBLoc,DEV,Valloc);
       continue;
       end;
    Req := OpenSql ('SELECT GA_PMAP,GA_PAHT,GA_DPR,GA_PMAP,GA_PMRP FROM ARTICLE WHERE GA_ARTICLE = "' + TOBLOc.getValue('BLO_ARTICLE') + '"',true,-1, '', True);
    if (not Req.eof) then
    begin
       if (VH_GC.ModeValoPa = 'PMA') and (pos(Req.findfield('GA_TYPEARTICLE').AsString,'MAR;ARP')>0) then
       begin
        if Req.findfield('GA_PMAP').AsFloat<>  0 then
        begin
       		TOBLoc.Putvalue('BLO_DPA',Req.findfield('GA_PMAP').AsFloat );

        end else
        begin
       		TOBLoc.Putvalue('BLO_DPA',Req.findfield('GA_PAHT').AsFloat );
        end;
       end else
       begin
       	TOBLoc.Putvalue('BLO_DPA',Req.findfield('GA_PAHT').AsFloat );
       end;
       TOBLOC.Putvalue('BLO_DPR',Req.findfield('GA_DPR').AsFloat );
       TOBLOC.Putvalue('BLO_PMAP',Req.findfield('GA_PMAP').AsFloat );
       TOBLOC.Putvalue('BLO_PMRP',Req.findfield('GA_PMRP').AsFloat );
    end;
    ferme(Req);
    end;
CalculeOuvrageDoc (TOBOUV,1,1,true,DEV,valeurs,true,false);
StockeLesTypes (TOBOUV,Valeurs);

end;

procedure PutValueDetailOuv (TOBOuvrage : TOB;Champ:string; Valeur : variant; TOBARTICLES : TOB=nil);
var Indice,il : Integer;
    TOBL,TOBART : TOB;
    prixpasModif : boolean;
    QQ : TQuery;
begin
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBL := TOBOuvrage.detail[Indice];
    if TOBL.detail.Count > 0 then PutValueDetailOuv (TOBL,Champ,Valeur);
    if indice=0 then il:=TOBL.GetNumChamp(Champ) ;
    if il>0 then
    begin
      TOBL.PutValeur(il,Valeur) ;
      if (Champ = 'BLO_DOMAINE') and (TOBL.GetValue('BLO_ARTICLE')<>'') then
      begin
        if TOBArticles <> nil then
        begin
			    TOBArt :=TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.getValue('BLO_ARTICLE')],true);
          prixpasModif :=  (TOBART.getValue('GA_PRIXPASMODIF')='X');
        end;
        if (TOBART = nil) then
        begin
          TOBART := TOB.Create('ARTICLE',nil,-1);
          QQ := OPenSql('SELECT GA_PRIXPASMODIF FROM ARTICLE WHERE GA_ARTICLE="'+TOBL.getValue('BLO_ARTICLE')+'"',true,1,'',true);
          if not QQ.eof then
          begin
            TOBART.selectDB('',QQ);
            prixpasModif :=  (TOBART.getValue('GA_PRIXPASMODIF')='X');
          end;
          ferme(QQ);
          TOBARt.free;
        end;
        if not PrixPasModif then
        begin
      	  AppliqueCoefDomaineActOuv (TOBL);
          if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL,DEV);
        end;
      end;
    end;
  end;
end;

function LookUpFournisseur (CodFourn, CodArtic ,LibArt: string; SelectFourSiPasCataLog : boolean=false) : string;
var TOBFourn,TOBFournFille,TOBARt,TOBTiers,TOBTARIf,TOBCataFouPrinc : TOB ;
    QFourn,QQ: TQuery ;
    SQL : string ;
    MaCle,CodRefer : String ;
    TypeRech : String ;
    ExisteCatalogue : boolean ;
    NbCatalogue : integer ;
    BMonoFournisseur : boolean;
    TarifComplet : boolean;
    MTPAF : double;
    CodeArticleGen,StLEQUEL,SousFamilletarif : string;
    PrixPourQTe : double;
Begin
  TarifComplet := GetParamSoc('SO_BTRECHTARIFFOU');
  if TarifComplet then
  begin
    TOBART := TOB.Create ('ARTICLE',nil,-1);
    //
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+CodArtic+'"',true,-1, '', True);
    //
    TOBArt.SelectDb('',QQ);
    ferme (QQ);
    InitChampsSupArticle (TOBArt);
    //
    SousFamilletarif := TOBArt.getValue('GA2_SOUSFAMTARART');
    //
    TOBTiers := TOB.Create ('TIERS',nil,-1);
    TOBTArif := TOB.Create ('TARIF',nil,-1);
    TOBCataFouPrinc := TOB.Create ('CATALOGU',nil,-1);
    if TOBART.GetValue('GA_FOURNPRINC') <> '' then
    begin
      QQ := OpenSql ('SELECT * FROM CATALOGU WHERE GCA_TIERS="'+TOBART.GEtVAlue('GA_FOURNPRINC')+'" AND GCA_ARTICLE="'+CodArtic+'"',True,-1, '', True);
      if not QQ.eof Then TOBCataFouPrinc.SelectDB ('',QQ);
      ferme (QQ);
    end;
  end;
  result := '';
  if (ctxMode in V_PGI.PGIContexte) then BMonoFournisseur := GetParamsoc('SO_MONOFOURNISS')
  else BMonoFournisseur := False;
  TOBFourn:=nil ;
  TRY
  // Recherche s'il existe au moins un catalogue fournisseur pour l'article
    TypeRech:='1';
    ExisteCatalogue:=ExisteSQL('Select GCA_REFERENCE from CATALOGU where GCA_ARTICLE="'+CodArtic+'"');

    if (BMonoFournisseur = True) and
       ((Not ExisteCatalogue) or (Trim(Copy(CodArtic,1,18))='')) then
    begin
      // Recherche s'il existe dans le catalogue, des rfrences du fournisseur non
      // rattach  un article
      TypeRech:='2';
      CodArtic:='';
      CodArtic:=CodeArticleUnique2(CodArtic,'') ;
      ExisteCatalogue:=ExisteSQL('Select GCA_REFERENCE from CATALOGU where (GCA_ARTICLE="'+CodArtic+'") and (GCA_TIERS="'+CodFourn+'")');
    end;

    if ExisteCatalogue=True then
      begin
        // Chargement en TOB des catalogues fournisseurs
        SQL:= 'Select GCA_TIERS, GCA_REFERENCE, GCA_PRIXBASE, GCA_PRIXPOURQTEAC,GCA_PRIXVENTE, GCA_DELAILIVRAISON, T1.T_LIBELLE , '
                       +'T1.T_AUXILIAIRE from CATALOGU '
                       +' left outer join TIERS T1 on GCA_TIERS=T_TIERS '
                       +' where (GCA_ARTICLE="'+CodArtic+'")' ;
        if TypeRech='2' then SQL:=SQL + ' and (GCA_TIERS="'+CodFourn+'")';  // Slection restrictive sur le fournisseur de l'article
        NbCatalogue := 0;
        QFourn:=OpenSql(SQL,True,-1, '', True) ;
        if not QFourn.EOF then
        begin
          TOBFourn:= TOB.Create('les fournisseurs',NIL,-1);
          TOBFourn.AddChampSupValeur ('MODE',TarifComplet,false);
          While not QFourn.EOF do
          Begin
            TOBFournFille:=Tob.create('un fournisseur',TOBFourn,-1);
            TOBFournFille.AddChampSup('GCA_TIERS',False) ;
            TOBFournFille.AddChampSup('GCA_REFERENCE',False) ;
            TOBFournFille.AddChampSup('T_LIBELLE',False) ;
            TOBFournFille.AddChampSup('GCA_PRIXBASE',False) ;
            TOBFournFille.AddChampSupValeur('GF_PRIXF',0,False) ;
            TOBFournFille.AddChampSupValeur('GF_REMISE','',False) ;
            TOBFournFille.AddChampSupValeur('GF_TARIF',0,False) ;
            TOBFournFille.AddChampSup('GCA_DELAILIVRAISON',False) ;
            TOBFournFille.AddChampSup('T_AUXILIAIRE',False) ;
            TOBFournFille.AddChampSupValeur('GCA_ARTICLE',CodArtic,false) ;

            CodFourn := QFourn.Findfield('GCA_TIERS').AsString ;
            TOBFournFille.PutValue('GCA_TIERS', CodFourn) ;
            CodRefer := QFourn.Findfield('GCA_REFERENCE').AsString ;
            //
            if TarifComplet then
            begin
               TOBTiers.InitValeurs;
               TOBTArif.InitValeurs;
               TOBTIers.PutValue ('T_AUXILIAIRE',QFourn.Findfield('T_AUXILIAIRE').AsString );
               TOBTIERS.loaddb();
               GetTarifGlobal (CodArtic,TOBArt.getValue('GA_TARIFARTICLE'),SousFamilletarif,'ACH',TOBArt,TOBTiers,TOBTarif,true);
               if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
               begin
                  MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
               end else
               begin
                  prixPourQte := QFourn.Findfield('GCA_PRIXPOURQTEAC').Asfloat;
                  if PrixPourQte = 0 then PrixPourQte := 1;
                  if (QFourn.Findfield('GCA_PRIXVENTE').Asfloat = 0) and
                     (QFourn.Findfield('GCA_PRIXBASE').Asfloat = 0) then
                  begin
                    if TOBCataFouPrinc.GetValue('GCA_PRIXBASE') = 0 then
                    begin
                      MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXVENTE')/prixPourQte;
                    end else
                    begin
                      MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXBASE')/prixPourQte;
                    end;
                  end else
                  begin
                    if QFourn.Findfield('GCA_PRIXBASE').Asfloat = 0 then
                    begin
                     MTPAF :=QFourn.Findfield('GCA_PRIXVENTE').Asfloat/PrixPourQte;
                    end else
                    begin
                     MTPAF :=QFourn.Findfield('GCA_PRIXBASE').Asfloat/PrixPourQTe;
                    end;
                  end;
               end;
               TOBFournFille.PutValue('GCA_PRIXBASE',MTPAF);
               TOBFournFille.PutValue('GF_PRIXF',MTPAF);
               TOBFournFille.PutValue('GF_REMISE',TOBTarif.getValue('GF_CALCULREMISE'));
               MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
               TOBFournFille.PutValue('GF_TARIF',MTPAF);
            end else TOBFournFille.PutValue('GCA_PRIXBASE',QFourn.Findfield('GCA_PRIXBASE').Asfloat);
            //
            TOBFournFille.PutValue('GCA_REFERENCE', CodRefer) ;
            TOBFournFille.PutValue('T_LIBELLE',QFourn.Findfield('T_LIBELLE').AsString);
            TOBFournFille.PutValue('GCA_DELAILIVRAISON',QFourn.Findfield('GCA_DELAILIVRAISON').AsInteger);
            inc ( NbCatalogue );
            QFourn.next ;
          End ;
        end;
        Ferme(QFourn) ;
        if (BMonoFournisseur = True) and (NbCatalogue = 1) and (TypeRech = '1') then
        begin
          // Acces direct en modification de la fiche Catalogue du fournisseur de l'article
          MaCle:=AGLLanceFiche('GC','GCCATALOGU_SAISI3','',CodRefer+';'+CodFourn,'ACTION=MODIFICATION') ;
        end else
        begin
          // Affichage de la fiche Grid proposant les diffrents choix du catalogue fournisseur
          TheTob:=TOBFourn ;
          MaCle:=AGLLanceFiche('BTP','BTGRIDFOURNISSEUR','','','ART='+CodArtic+';FOURN='+CodFourn+';LIBART='+libart+';TYPAFF='+TypeRech) ;
        end;
      end else
      begin
        if not SelectFourSiPasCataLog then
        begin
          // Acces en cration de la fiche Catalogue du fournisseur principal de l'article
          MaCle:=AGLLanceFiche('GC','GCCATALOGU_SAISI3','','','ACTION=CREATION;LIB='+libart+';ARTICLE='+CodArtic) ;
        end else
        begin
          MaCle := AGLLanceFiche('GC', 'GCFOURNISSEUR_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION')
        end;
      end ;

    if (BMonoFournisseur = False) then
      begin
        If (MaCle<>'') And (MaCle<>'Abandon') then result:=ReadTokenst(MaCle) ;
      end;
  FINALLY
    if TOBFourn<>nil then TOBFourn.free ;
    if TarifComplet then
    begin
       TOBART.free;
       TOBTiers.free;;
       TOBTArif.free;
       TOBCataFouPrinc.free;
    end;
  END;
End;

//

procedure MultipleLookUpFournisseur (TOBLiaison : TOB );
var TOBFourn,TOBFournFille,TOBARt,TOBTiers,TOBTARIf,TOBCataFouPrinc : TOB ;
    QFourn,QQ: TQuery ;
    SQL : string ;
    MaCle,CodRefer : String ;
    TypeRech : String ;
    ExisteCatalogue : boolean ;
    NbCatalogue,Indice : integer ;
    TarifComplet : boolean;
    MTPAF : double;
    CodeArticleGen,StLEQUEL,SousFamilletarif : string;
    PrixPourQTe : double;
    CodFourn, CodArtic ,LibArt: string;
Begin
  if TOBLiaison = nil then exit;
  if not TOBLiaison.FieldExists ('FOURNISSEUR') then exit;
  CodFourn := TOBLiaison.getValue ('FOURNISSEUR');
  CodArtic := TOBLiaison.getValue ('ARTICLE');
  LibArt := TOBLiaison.getValue ('LIBELLEART');
  //
  TarifComplet := GetParamSoc('SO_BTRECHTARIFFOU');
  if TarifComplet then
  begin
    TOBART := TOB.Create ('ARTICLE',nil,-1);
    TOBTiers := TOB.Create ('TIERS',nil,-1);
    TOBTArif := TOB.Create ('TARIF',nil,-1);
    TOBCataFouPrinc := TOB.Create ('CATALOGU',nil,-1);
    //
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
    							 'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
    							 'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+CodArtic+'"',true,-1, '', True);
    //
    TOBArt.SelectDb('',QQ);
    ferme (QQ);
    InitChampsSupArticle (TOBArt);
    //
    SousFamilletarif := TOBArt.getValue('GA2_SOUSFAMTARART');
    //
    if TOBART.GetValue('GA_FOURNPRINC') <> '' then
    begin
      QQ := OpenSql ('SELECT * FROM CATALOGU WHERE GCA_TIERS="'+TOBART.GEtVAlue('GA_FOURNPRINC')+'" AND GCA_ARTICLE="'+CodArtic+'"',True,-1, '', True);
      if not QQ.eof Then TOBCataFouPrinc.SelectDB ('',QQ);
      ferme (QQ);
    end;
  end;
  TOBFourn:=nil ;
  TRY
  // Recherche s'il existe au moins un catalogue fournisseur pour l'article
    TypeRech:='1';
    ExisteCatalogue:=ExisteSQL('Select GCA_REFERENCE from CATALOGU where GCA_ARTICLE="'+CodArtic+'"');

    if ExisteCatalogue=True then
      begin
        // Chargement en TOB des catalogues fournisseurs
        SQL:= 'Select GCA_TIERS, GCA_REFERENCE, GCA_PRIXBASE, GCA_PRIXPOURQTEAC,GCA_PRIXVENTE, GCA_DELAILIVRAISON, T1.T_LIBELLE , '
                       +'T1.T_AUXILIAIRE from CATALOGU '
                       +' left outer join TIERS T1 on GCA_TIERS=T_TIERS '
                       +' where (GCA_ARTICLE="'+CodArtic+'")' ;
        if TypeRech='2' then SQL:=SQL + ' and (GCA_TIERS="'+CodFourn+'")';  // Slection restrictive sur le fournisseur de l'article
        NbCatalogue := 0;
        QFourn:=OpenSql(SQL,True,-1, '', True) ;
        if not QFourn.EOF then
        begin
          TOBFourn:= TOB.Create('les fournisseurs',NIL,-1);
          TOBFourn.AddChampSupValeur ('MODE',TarifComplet,false);
          While not QFourn.EOF do
          Begin
            TOBFournFille:=Tob.create('un fournisseur',TOBFourn,-1);
            TOBFournFille.AddChampSup('GCA_TIERS',False) ;
            TOBFournFille.AddChampSup('GCA_REFERENCE',False) ;
            TOBFournFille.AddChampSup('T_LIBELLE',False) ;
            TOBFournFille.AddChampSup('GCA_PRIXBASE',False) ;
            TOBFournFille.AddChampSupValeur('GF_PRIXF',0,False) ;
            TOBFournFille.AddChampSupValeur('GF_REMISE','',False) ;
            TOBFournFille.AddChampSupValeur('GF_TARIF',0,False) ;
            TOBFournFille.AddChampSup('GCA_DELAILIVRAISON',False) ;
            TOBFournFille.AddChampSup('T_AUXILIAIRE',False) ;
            TOBFournFille.AddChampSupValeur('GCA_ARTICLE',CodArtic,false) ;

            CodFourn := QFourn.Findfield('GCA_TIERS').AsString ;
            TOBFournFille.PutValue('GCA_TIERS', CodFourn) ;
            CodRefer := QFourn.Findfield('GCA_REFERENCE').AsString ;
            //
            if TarifComplet then
            begin
               TOBTiers.InitValeurs;
               TOBTArif.InitValeurs;
               TOBTIers.PutValue ('T_AUXILIAIRE',QFourn.Findfield('T_AUXILIAIRE').AsString );
               TOBTIERS.loaddb();
               GetTarifGlobal (CodArtic,TOBArt.getValue('GA_TARIFARTICLE'),SousFamilletarif,'ACH',TOBArt,TOBTiers,TOBTarif,true);
               if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
               begin
                  MTPAF :=TOBTarif.GetValue('GF_PRIXUNITAIRE');
               end else
               begin
                  prixPourQte := QFourn.Findfield('GCA_PRIXPOURQTEAC').Asfloat;
                  if PrixPourQte = 0 then PrixPourQte := 1;
                  if (QFourn.Findfield('GCA_PRIXVENTE').Asfloat = 0) and
                     (QFourn.Findfield('GCA_PRIXBASE').Asfloat = 0) then
                  begin
                    if TOBCataFouPrinc.GetValue('GCA_PRIXBASE') = 0 then
                    begin
                      MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXVENTE')/prixPourQte;
                    end else
                    begin
                      MTPAF :=TOBCataFouPrinc.GetValue('GCA_PRIXBASE')/prixPourQte;
                    end;
                  end else
                  begin
                    if QFourn.Findfield('GCA_PRIXBASE').Asfloat = 0 then
                    begin
                     MTPAF :=QFourn.Findfield('GCA_PRIXVENTE').Asfloat/PrixPourQte;
                    end else
                    begin
                     MTPAF :=QFourn.Findfield('GCA_PRIXBASE').Asfloat/PrixPourQTe;
                    end;
                  end;
               end;
               TOBFournFille.PutValue('GCA_PRIXBASE',MTPAF);
               TOBFournFille.PutValue('GF_PRIXF',MTPAF);
               TOBFournFille.PutValue('GF_REMISE',TOBTarif.getValue('GF_CALCULREMISE'));
               MTPAF := arrondi(MTPAF * (1-(TOBTarif.GetValue('GF_REMISE')/100)),V_PGI.OkDecP );
               TOBFournFille.PutValue('GF_TARIF',MTPAF);
            end else TOBFournFille.PutValue('GCA_PRIXBASE',QFourn.Findfield('GCA_PRIXBASE').Asfloat);
            //
            TOBFournFille.PutValue('GCA_REFERENCE', CodRefer) ;
            TOBFournFille.PutValue('T_LIBELLE',QFourn.Findfield('T_LIBELLE').AsString);
            TOBFournFille.PutValue('GCA_DELAILIVRAISON',QFourn.Findfield('GCA_DELAILIVRAISON').AsInteger);
            inc ( NbCatalogue );
            QFourn.next ;
          End ;
        end;
        Ferme(QFourn) ;
        // Affichage de la fiche Grid proposant les diffrents choix du catalogue fournisseur
        TheTob:=TOBFourn ;
        TheTOB.data := TOBLiaison;
        AGLLanceFiche('BTP','BTMULGRIDCATALOG','','','ART='+CodArtic+';FOURN='+CodFourn+';LIBART='+libart+';TYPAFF='+TypeRech) ;
      end;
  FINALLY
    if TOBFourn<>nil then TOBFourn.free ;
    if TarifComplet then
    begin
       TOBART.free;
       TOBTiers.free;;
       TOBTArif.free;
       TOBCataFouPrinc.free;
    end;
  END;
End;
//

procedure recupInfoGenerique (TOBDEST,TOBRef : TOB);
var i : integer;
    PrefixeRef,PrefixeDest,suffixe : string;
    ChampDest : string;
    Mcd : IMCDServiceCOM;
    FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  PrefixeRef := Mcd.TableToPrefixe(TOBRef.Nomtable);
  PrefixeDest := Mcd.TableToPrefixe(TobDest.Nomtable);
  //
  FieldList := MCD.GetTable(TOBREF.NomTable).Fields;
  FieldList.Reset();
  While FieldList.MoveNext do
  begin
    Suffixe := copy((FieldList.Current as IFieldCOM).name,Length(PrefixeRef)+2,255);
    if Suffixe <> '' then
    begin
      ChampDest := PrefixeDest+'_'+suffixe;
      if TOBDest.FieldExists (ChampDest) then TOBDEST.PutValue (PrefixeDest+'_'+suffixe,TOBREF.Getvaleur(i));
    end;
  end;
end;

procedure AjouteSousDetailVide(TOBL, TOBOuvrage : TOB);
var TOBLN,TOBTMP : TOB;
begin
  TOBLN:=TOB.Create('',TOBOuvrage,-1) ;
  TOBL.SetInteger('GL_INDICENOMEN',TOBOuvrage.detail.count); 
  InsertionChampSupOuv (TOBLN,false);
  TOBTMP:=TOB.Create('LIGNEOUV',TOBLN,-1) ;
  InsertionChampSupOuv (TOBTMP,false);
  TOBTMP.PutValue('BLO_NATUREPIECEG',TOBL.GetValue('GL_NATUREPIECEG')) ;
  TOBTMP.PutValue('BLO_SOUCHE',TOBL.GetValue('GL_SOUCHE')) ;
  TOBTMP.PutValue('BLO_DATEPIECE',TOBL.GetValue('GL_DATEPIECE')) ;
  TOBTMP.PutValue('BLO_AFFAIRE',TOBL.GetValue('GL_AFFAIRE')) ;
  TOBTMP.PutValue('BLO_AFFAIRE1',TOBL.GetValue('GL_AFFAIRE1')) ;
  TOBTMP.PutValue('BLO_AFFAIRE2',TOBL.GetValue('GL_AFFAIRE2')) ;
  TOBTMP.PutValue('BLO_AFFAIRE3',TOBL.GetValue('GL_AFFAIRE3')) ;
  TOBTMP.PutValue('BLO_AVENANT',TOBL.GetValue('GL_AVENANT')) ;
  TOBTMP.PutValue('BLO_NUMERO',TOBL.GetValue('GL_NUMERO')) ;
  TOBTMP.PutValue('BLO_INDICEG',TOBL.GetValue('GL_INDICEG')) ;
  TOBTMP.PutValue('BLO_NUMLIGNE',TOBL.GetValue('GL_NUMLIGNE')) ;
  TOBTMP.PutValue('BLO_QTEDUDETAIL',1) ;
  TOBTMP.PutValue('BLO_COMPOSE','') ;
end;

procedure SupprimeSDOuvVide (TOBL,TOBOuvrage : TOB);

  procedure DeleteLigneVide (TOBOUv : TOB);
  var II : integer;
  begin
    II := 0;
    if TOBOUV.detail.count = 0 then exit;
    repeat
      if TOBOUV.detail[II].getString('BLO_CODEARTICLE')='' then TOBOUV.detail[II].free
                                                           else Inc(II);
    until II >= TOBOUV.detail.count;
  end;

var TOBOUV : TOB;
    IndiceNomen : integer;
begin
  IndiceNomen := TOBL.getInteger('GL_INDICENOMEN');
  if IndiceNomen = 0 then exit;
  //FV1 : 28/07/2016
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen -1 > TobOuvrage.detail.count - 1 then exit;

  TOBOUV := TOBOUvrage.detail[IndiceNomen-1]; if TOBOUv = nil then exit;
  DeleteLigneVide (TOBOUV);
  if TOBOUV.detail.count = 0 then
  begin
    TOBL.SetInteger('GL_INDICENOMEN',0);
  end;
end;




procedure AjouteSousDetail (TOBL,TOBOuvrage,TOBArticle : TOB;MarchandiseUniquement : boolean);
var TOBTMP,TOBARt,TOBREf,TOBGroupeN : TOB;
    TypeRessource : string;
begin
TOBRef := TOB.Create ('ARTICLE',nil,-1);
TRY
TOBArt := TOBArticle.findfirst (['GA_ARTICLE'],[TOBL.GetValue('GL_ARTICLE')],false);
if TOBART = nil then BEGIN TOBREF.free; TOBREF := nil; Exit; END;
if (TOBL.GetValue('GL_TYPEARTICLE')='PRE') then
   begin
   TypeRessource := GetTypeRessourceFromNature (TOBArt.GetValue('GA_NATUREPRES'));
   if ((TypeRessource = 'SAL') or (TypeREssource = 'MAT')) and
      (MarchandiseUniquement) then
      BEGIN
      TOBRef.Free;
      TOBREf := nil;
      Exit;
      END;
   end;
TOBRef.Dupliquer (TOBART,false,true);
TOBRef.putValue ('GA_FOURNPRINC',TOBL.GetValue('GL_FOURNISSEUR'));
TOBGroupeN:=TOB.Create('LIGNEOUV',TOBOuvrage,-1) ;
TOBTMP:=TOB.Create('LIGNEOUV',TOBGroupeN,-1) ;
TOBL.PutValue ('GL_INDICENOMEN', TOBOuvrage.detail.count);
InsertionChampSupOuv (TOBTMP,false);
recupInfoGenerique (TOBTmp,TOBL);
TOBTmp.putvalue('BLO_PUHTBASE',TOBL.GetValue('GL_PUHTDEV'));
TOBTMP.PutValue('BLO_NIVEAU',1) ;
TOBTMP.PutValue('BLO_PRIXPOURQTE',1) ;
TOBTMP.PutValue('BLO_QTEDUDETAIL',1) ;
TOBTMP.PutValue('BLO_COMPOSE',TOBL.GetValue('GL_ARTICLE')) ;
TOBTMP.PutValue('BLO_QTEFACT',1) ;
RecupValoDetail (TOBTmp,TOBRef);
FINALLY
if TOBREF <> nil then TOBREF.free;
END;
end;

procedure SuprimeSousDetail(TOBL,TOBOUvrage:TOB);
var TOBN : TOB;
    IndiceNomen : Integer;
begin

  //FV1 : 28/07/2016
  If Tobl = nil then Exit;

  IndiceNomen := TOBL.GetValue ('GL_INDICENOMEN')-1;

  if Indicenomen < 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen > TobOuvrage.detail.count  then exit;

  TOBN:=TOBOuvrage.Detail[IndiceNomen] ;
  if TOBN = nil then exit;
  // Recup de l'info manquante
  TOBL.PutValue ('GL_FOURNISSEUR',TOBN.detail[0].GetValue('BLO_FOURNISSEUR'));
  // --
  TOBN.free;
  TOBL.Putvalue ('GL_INDICENOMEN',0);
end;

procedure PositionneVarianteOuv (TobOuv : TOB;variante : boolean);
var indice : integer;
    TOBl : TOB;
begin
  for Indice := 0 to TOBOuv.detail.count -1 do
  begin
    TOBL := TOBOuv.detail[Indice];
    if TOBL.fieldExists ('BLO_TYPELIGNE') then SetTypeLigne (TOBL,Variante);
    if TOBL.detail.count > 0 then PositionneVarianteOuv (TOBL,variante);
  end;
end;

procedure PositionneLigneOuvVariante (TOBL,TOBOuvrage : TOB ; variante : boolean);
var indiceNomen : integer;
    TOBLoc : TOB;
begin
  //FV1 : 28/07/2016
  If Tobl = nil then Exit;

  IndiceNomen := TOBL.GetValue('GL_INDICENOMEN') - 1 ;

  if IndiceNomen < 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen > TobOuvrage.detail.count then exit;

  TOBLoc := TOBOuvrage.detail[IndiceNomen];

  PositionneVarianteOuv (TobLoc,variante);

end;

function IsArticlePrixPose (TOBL : TOB) : boolean;
var TypeArticle : string;
		prefixe : string;
begin
result := false;
prefixe := GetPrefixeTable(TOBL);
TypeArticle := TOBL.getValue(Prefixe+'_TYPEARTICLE');
if (TypeArticle = 'ARP') or (TypeArticle = 'ARV') then result := true;
end;


function IsOuvrage (TOBL : TOB) : boolean;
var TypeArticle : string;
		prefixe : string;
begin
result := false;
prefixe := GetPrefixeTable(TOBL);
TypeArticle := TOBL.getValue(Prefixe+'_TYPEARTICLE');
if (TypeArticle = 'ARP') or (TypeArticle = 'OUV') then result := true;
end;

function IsOuvrageComptabiliseDetail (TOBL : TOB) : boolean;
var TypeArticle : string;
	  Typenomen : string;
		prefixe : string;
    NatureTravail : Double;
begin
result := false;
prefixe := GetPrefixeTable(TOBL);
TypeArticle := TOBL.getValue(Prefixe+'_TYPEARTICLE');
TypeNomen := TOBL.getValue(Prefixe+'_TYPENOMENC');
if TOBL.FieldExists('BLF_NATURETRAVAIL') then NatureTravail := Valeur(TOBL.GetValue('BLF_NATURETRAVAIL'))
																				 else NatureTravail := 0;

if  (TypeArticle = 'OUV') and ((TypeNomen='OUV') or (NatureTravail >= 10)) then result := true;
end;

//Vrification si qte exprim pour dans dclinaison <> 1.
function GetQteDetailOuv (TOBL,TOBOUvrages : TOB) : integer;
var NumOuvrage : integer;
    TOBO : TOB;
begin
  result := 1;
  if TOBL.GetValue('GL_TYPEARTICLE') <> 'OUV' Then Exit;
  NumOuvrage := TOBL.GetValue('GL_INDICENOMEN');
  if NumOuvrage = 0 then exit;
  TOBO := TOBOuvrages.Detail[NumOuvrage-1];
  result := TOBO.detail[0].getValue('BLO_QTEDUDETAIL');
  if Result = 0 then result := 1;
end;

procedure PosQteDetailAndCalcule (TOBLig,TOBOUvrages : TOB;QteDetail : double;DEV: RDevise);
var
  Indice: Integer;
  TOBL,TOBO: TOB;
  TOBOuvrage: TOB;
  NumOuvrage : integer;
  EnHt : boolean;
  Valeurs : T_Valeurs;
  ValLoc : T_Valeurs;
begin
  if TOBOuvrages = nil then exit;
  InitTableau  (Valeurs);
  NumOuvrage := TOBLig.GetValue('GL_INDICENOMEN');
  EnHt := (TOBLIG.GetValue('GL_FACTUREHT')='X');
  if NumOuvrage = 0 then exit;
  TOBOuvrage := TOBOuvrages.Detail[NumOuvrage-1];
  //
  TOBL:=TOBOuvrage.findfirst(['BLO_TYPEARTICLE'],['POU'],false);
  //
  InitTableau (valeurs);
  //
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBO := TOBOuvrage.detail [Indice];
    // VARIANTE
    if isVariante (TOBO) then continue;
    // --
    if TOBO.GetValue('BLO_TYPEARTICLE') = 'POU' then
    begin
      continue;
    end;
    //
    TOBO.Putvalue ('BLO_QTEDUDETAIL',QteDetail);
  end;
  //
  CalculeOuvrageDoc (TOBOuvrage,1,1,true,DEV,valeurs,EnHt);
  //
  TOBLig.Putvalue('GL_PUHT',DeviseToPivotEx(valeurs[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
  TOBLig.Putvalue('GL_PUTTC',DevisetoPivotEx(valeurs[3],DEV.taux,DEV.quotite,V_PGI.okdecP));
  TOBLig.Putvalue('GL_PUHTDEV',valeurs[2]);
  TOBLig.Putvalue('GL_PUTTCDEV',valeurs[3]);
  TOBLig.Putvalue('GL_PUHTBASE',TOBLIG.GetValue('GL_PUHT'));
  TOBLig.Putvalue('GL_PUTTCBASE',TOBLIG.GetValue('GL_PUTTC'));
  TOBLig.Putvalue('GL_DPA',valeurs[0]);
  TOBLig.Putvalue('GL_DPR',valeurs[1]);
  TOBLig.Putvalue('GL_PMAP',valeurs[6]);
  TOBLig.Putvalue('GL_PMRP',valeurs[7]);
  TOBLig.putvalue('GL_TPSUNITAIRE',valeurs[9]);
  positionneCoefMarge (TOBL);
  StockeInfoTypeLigne (TOBLIG,valeurs);
end;

function ApplicPrixPose (TOBPiece : TOB) : boolean;
var venteAchat,Prefixe : string;
begin
	result := true;
	if TOBPiece.NomTable = 'PIECE' then Prefixe := 'GP_';
  if TOBPiece.nomTable = 'LIGNE' then Prefixe := 'GL_';
  if TOBPiece.nomTable = 'LIGNEOUV' then Prefixe := 'BLO_';
	VenteAchat:=GetInfoParPiece(TOBPiece.getValue(Prefixe+'NATUREPIECEG'),'GPP_VENTEACHAT') ;
  // Seules ces pices peuvent accueillir le detail d'un ouvrage
  if (not NaturepieceOKPourOuvrage(TOBPiece)) or
  	 (VenteAchat='ACH') then result := false;
end;

procedure ImpactePrixOuv (TOBO : TOb;CalculPv,EnHt : boolean; InclusStInFg : boolean=false);
begin
//   if not TOBO.FieldExists ('__COEFMARG') then TOBO.AddChampSupValeur ('__COEFMARG',0);
   if (not TOBO.FieldExists ('BLO_DPR')) then exit;
   //
   if (TOBO.GetValue('BLO_DPR') = 0) and (TOBO.GetValue('BLO_DPA') = 0) and (TOBO.GetValue('BLO_PUHTDEV')<>0) then exit;
   //
   if TOBO.GetValue('BLO_DPR') = 0 then TOBO.PutValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
   if TOBO.GetValue('BLO_DPR') = 0 then
   begin
      TOBO.PutValue('BLO_COEFMARG',1);
      TOBO.PutValue('POURCENTMARG',Arrondi((TOBO.GetValue('BLO_COEFMARG')-1)*100,2));
   end else
   begin
      if TOBO.GetValue('BLO_COEFMARG')= 0 then
      begin
         if EnHt then TOBO.PutValue('BLO_COEFMARG',arrondi(TOBO.GetValue('BLO_PUHTDEV')/TOBO.GetValue('BLO_DPR'),4))
                 else TOBO.PutValue('BLO_COEFMARG',arrondi(TOBO.GetValue('BLO_PUTTCDEV')/TOBO.GetValue('BLO_DPR'),4));
        TOBO.PutValue('POURCENTMARG',Arrondi((TOBO.GetValue('BLO_COEFMARG')-1)*100,2));
      end;
   end;

   if ((TOBO.GetValue('BLO_TYPEARTICLE')<>'PRE') OR
//   		(RenvoieTypeRes (TOBO.GetValue('BLO_ARTICLE')) <> 'ST') then
   		(TOBO.GetValue('BNP_TYPERESSOURCE') <> 'ST') OR (InclusStInFg) )  and
      (TOBO.GetValue('BLO_NONAPPLICFRAIS')<>'X') then
   begin
			TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
   end;
   TOBO.putValue('ANCPR',TOBO.GetValue('BLO_DPR'));
   if CalculPv then
   begin
   if EnHt then TOBO.PutValue('BLO_PUHTDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG'),V_PGI.OkDECP))
           else TOBO.PutValue('BLO_PUTTCDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG'),V_PGI.OKdecP));
   end;

  if TOBO.GetValue('BLO_PUHT') <> 0 then
  begin
    TOBO.PutValue('POURCENTMARQ',Arrondi(((TOBO.GetValue('BLO_PUHT')- TOBO.GetValue('BLO_DPR'))/TOBO.GetValue('BLO_PUHT'))*100,2));
  end else
  begin
    TOBO.PutValue('POURCENTMARQ',0);
  end;

end;

procedure BeforeRecalculPrOuv (TOBOuv : TOB; EnHt : boolean;CalculPv : boolean=true;InclusSTinFG : boolean=false);
var Indice : integer;
		TOBO : TOB;
begin
  ImpactePrixOuv (TOBOuv,CalculPv,EnHt,InclusSTinFG);
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBO := TOBOUV.detail[Indice];
    ImpactePrixOuv (TOBO,CalculPv,EnHt,InclusSTinFG);
    if TOBO.detail.count > 0 then BeforeRecalculPrOuv (TOBO,EnHt,InclusSTinFG);
  end;
end;

procedure AftercalculPrOuv (TOBOuv : TOB; EnHt : boolean;CalculPv : boolean=true);
var Indice : integer;
		TOBO : TOB;
begin
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBO := TOBOUV.detail[Indice];
    if CalculPv then
    begin
      if TOBO.GetValue('BLO_COEFMARG') <> 0 then
      begin
         if EnHt then TOBO.PutValue('BLO_PUHTDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG'),V_PGI.OkDECP))
                 else TOBO.PutValue('BLO_PUTTCDEV',Arrondi(TOBO.GetValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG'),V_PGI.OKdecP));
         if TOBO.detail.count > 0 then AfterCalculPrOuv (TOBO,EnHt);
      end;
  	end;
  end;
end;


Procedure DefinirLesOuv (TOBGroupeNomen,TOBNomen : TOB ; LaLig,idep : integer) ;
Var i,Lig : integer ;
    TOBLN,TOBPere,TOBLoc,TOBArt,TOBRef : TOB ;
    RefArticle,TypeArticle : String ;
    DEV:Rdevise;
    LigPere1,LigPere2,LigPere3,LigPere4 : integer;
    LigneN1,LigneN2,LigneN3,LigneN4,LigneN5: integer;
BEGIN
for i:=idep to TOBNomen.Detail.Count-1 do
    BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if lig <> LaLig then break;
    LigPere1 := 0; LigneN1 := TOBLN.GetValue('BLO_N1');
    LigPere2 := 0; LigneN2 := TOBLN.GetValue('BLO_N2');
    LigPere3 := 0; LigneN3 := TOBLN.GetValue('BLO_N3');
    LigPere4 := 0; LigneN4 := TOBLN.GetValue('BLO_N4');
    LigneN5 := TOBLN.GetValue('BLO_N5');
    if LigneN5 > 0 then
       begin
       // recherche du pere au niveau 4
       TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,LigneN4,0],True) ;
       end else if TOBLN.GetValue('BLO_N4') > 0 then
           begin
           // recherche du pere au niveau 3
           TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,LigneN3,0,0],True) ;
           end else if TOBLN.GetValue('BLO_N3') > 0 then
               begin
               // recherche du pere au niveau 2
               TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,LigneN2,0,0,0],True) ;
               end else if TOBLN.GetValue('BLO_N2') > 0 then
                   begin
                   // recherche du pere au niveau 1
                   TOBPere:=TOBGroupeNomen.FindFirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],[Lig,LigneN1,0,0,0,0],True) ;
                   end else TOBPere:=TOBGroupeNomen;

    if TOBPere<>Nil then
       BEGIN
       TOBLoc:=TOB.Create('LIGNEOUV',TOBPere,-1) ;
       TOBLoc.Dupliquer(TOBLN,False,True) ;
       InsertionChampSupOuv (TOBLOC,false);
       DEV.Code:=TOBLOC.GetValue('BLO_DEVISE') ; GetInfosDevise(DEV) ;
       CalculMontantHtDevLigOuv (TOBLOC,DEV);
       END ;
    END ;
END ;

Procedure ChargeLesOuvrages ( TOBPiece,TOBOuvrage: tob ; CleDoc : R_CleDoc );
Var i,OldL,Lig : integer ;
    TOBL,TOBLN,TOBNomen,TOBGroupeNomen : TOB ;
    OkN  : boolean ;
    Q    : TQuery ;
    valeurs : T_Valeurs;
    DEV : RDevise;
    IndiceNomen : integer;
    requete : string;
BEGIN
if not NaturepieceOKPourOuvrage (TOBPiece)  then exit;
OkN:=False ; OldL:=-1 ;
for i:=0 to TOBPiece.Detail.Count-1 do
    BEGIN
    TOBL:=TOBPiece.Detail[i] ;
    TOBL.PutValue('GL_INDICENOMEN',0) ;
    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') then OkN:=True ;
    END ;
if Not OkN then Exit ;

TOBNomen:=TOB.Create('',Nil,-1) ;
requete :='SELECT O.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM LIGNEOUV O '+
          'LEFT JOIN NATUREPREST N ON BNP_NATUREPRES=(SELECT GA_NATUREPRES FROM ARTICLE WHERE GA_ARTICLE=O.BLO_ARTICLE) '+
          'WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+
          ' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5';

//Q:=OpenSQL('SELECT * FROM LIGNEOUV WHERE '+WherePiece(CleDoc,ttdOuvrage,False)+' ORDER BY BLO_NUMLIGNE,BLO_N1, BLO_N2, BLO_N3, BLO_N4,BLO_N5',True,-1, '', True) ;
Q:=OpenSQL(Requete,True,-1, '', True) ;
TOBNomen.LoadDetailDB ('LIGNEOUV','','',Q,True,true);
Ferme(Q) ;
for i:=0 to TOBNomen.Detail.Count-1 do
    BEGIN
    TOBLN:=TOBNomen.Detail[i] ;
    InsertionChampSupOuv (TOBLN,false);
    TOBLN.putvalue('ANCPV',TOBLN.Getvalue('BLO_PUHTDEV'));
    TOBLN.putvalue('ANCPA',TOBLN.Getvalue('BLO_DPA'));
    TOBLN.putvalue('ANCPR',TOBLN.Getvalue('BLO_DPR'));
    Lig:=TOBLN.GetValue('BLO_NUMLIGNE') ;
    if OldL<>Lig then
       BEGIN
       TOBGroupeNomen:=TOB.Create('',TOBOuvrage,-1) ;
       TOBGroupeNomen.AddChampSup('UTILISE',False) ;
       TOBGroupeNomen.PutValue('UTILISE','X') ;
       IndiceNomen := TobOuvrage.detail.count;
       TobL := TobPiece.FindFirst(['GL_NUMLIGNE'], [Lig], False);
//       TOBL := GetTobLigne (TOBPiece,Lig);
       if TOBL <> nil then
          begin
          TOBL.PutValue('GL_INDICENOMEN',IndiceNomen);
          end;
//      CreerLesTOBOuv(TOBGroupeNomen,TOBNomen,TOBArticles,Lig,i) ;
			(* OPTIMIZATION *)
      DefinirLesOuv(TOBGroupeNomen,TOBNomen,Lig,i);
      (* ------------ *)
       END ;
    OldL:=Lig ;
    END ;
TOBNomen.Free ;
// FInalisation
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.GetValue('GP_TAUXDEV');
  for i:= 0 to TobOuvrage.detail.count -1 do
  begin
    TOBLN := TOBOuvrage.detail[I];
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBLN,1,1,true,DEV,valeurs,(TOBPiece.getValue('GP_FACTUREHT')='X'));
  end;
//
END ;

procedure PositionneOuvrageOut (TOBPiece,TOBOuvrage,TOBOuvragesP : TOB; ARow : integer);
var IndiceNomen : Integer;
		TOBL,TOBOuv,TOBOP : TOB;
begin
  // Correction LS
  TOBL := GetTOBLigne (TOBPiece,Arow);
  IndiceNomen := TOBL.GEtValue('GL_INDICENOMEN');
  //FV1 - 28/07/2016
  if IndiceNomen = 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen -1 > TobOuvrage.detail.count - 1 then exit;

  TOBOUV := TOBOUvrage.detail[IndiceNomen-1];
  TOBOuv.PutValue('UTILISE','-');
  // --
  TOBOP := FindOuvragesPlat (TOBL,TOBouvragesP);
  if TOBOP <> nil then TOBOP.free;
end;

procedure ReinitCoefPaPrOuvDetail (TOBL,TOBOUV : TOB);
var TOBO : TOB;
		Indice : integer;
begin
  if TOBOuv.fieldExists ('BLO_DPR') then
  begin
    TOBOUV.PutValue('BLO_COEFFG',0);
    TOBOUV.PutValue('BLO_COEFFC',0);
    TOBOUV.PutValue('BLO_DPR',TOBOUV.GetValue('BLO_DPA'));
    if TOBL.GetValue('GLC_FROMBORDEREAU')<>'X' then
    begin
      TOBOUV.PutValue('BLO_PUHTDEV',Arrondi(TOBOUV.GetValue('BLO_DPA')*TOBOUV.GetValue('BLO_COEFMARG'),V_PGI.okdecP));
      TOBOUV.PutValue('BLO_PUTTCDEV',Arrondi(TOBOUV.GetValue('BLO_DPA'),V_PGI.okdecP));
    end;
  end;
  //
  for Indice := 0 to TOBOuv.detail.count -1 do
  begin
    TOBO := TOBOUV.detail[Indice];
    ReinitCoefPaPrOuvDetail (TOBL,TOBO);
  end;
end;

procedure ReinitCoefPaPrOuv (TOBL,TOBouvrage : TOB);
var TOBOuv : TOB;
		IndiceNomen : Integer;
begin

  IndiceNomen := TOBL.GEtValue('GL_INDICENOMEN');

  //FV1 - 28/07/2016
  if IndiceNomen = 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen -1 > TobOuvrage.detail.count - 1 then exit;

  TOBOUV := TOBOUvrage.detail[IndiceNomen-1];

  if (TOBOUv <> nil) and (TOBOUV.detail.count > 0) then
  begin
  	//
    ReinitCoefPaPrOuvDetail (TOBL,TOBOUV);
  end;
end;

procedure SetAffaireSouDetailOuv (TOBPiece,TOBOuvrage : TOB) ;
var Indice: integer;
		TOBO : TOB;
begin
	for indice:= 0 to TOBOUvrage.detail.count -1 do
  begin
    TOBO := TOBOUvrage.detail[Indice];
    TOBO.PutValue('BLO_AFFAIRE',TOBPiece.GetVAlue('GP_AFFAIRE'));
    TOBO.PutValue('BLO_AFFAIRE1',TOBPiece.GetVAlue('GP_AFFAIRE1'));
    TOBO.PutValue('BLO_AFFAIRE2',TOBPiece.GetVAlue('GP_AFFAIRE2'));
    TOBO.PutValue('BLO_AFFAIRE3',TOBPiece.GetVAlue('GP_AFFAIRE3'));
    TOBO.PutValue('BLO_AVENANT',TOBPiece.GetVAlue('GP_AVENANT'));
    TOBO.PutValue('BLO_AVENANT',TOBPiece.GetVAlue('GP_AVENANT'));
    TOBO.putValue('BLO_DOMAINE',TOBPiece.GetVAlue('GP_DOMAINE'));
    if TOBO.detail.count > 0 then
    begin
    	SetAffaireSouDetailOuv (TOBPiece,TOBO);
    end;
  end;
end;

procedure ReinitCoefMarg (TOBL,TOBOuvrage : TOB);

  procedure ReinitCOefMargDetail (TOBOUV : TOB);
  var Indice : integer;
      TOBO : TOB;
  begin
    for Indice := 0 to TOBOuv.detail.count -1 do
    begin
      TOBO := TOBOUV.detail[Indice];
      TOBO.putvalue('BLO_COEFMARG',0);
      if TOBO.detail.count > 0 then  ReinitCOefMargDetail (TOBO);
    end;
  end;

var IndiceNomen : integer;
		TOBOUv : TOB;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN');

  //FV1 - 28/07/2016
  if IndiceNomen = 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen -1 > TobOuvrage.detail.count - 1 then exit;

  TOBOuv := TOBOuvrage.detail[IndiceNomen-1];
  if TOBOUv = nil then exit;
  TOBOUV.putvalue('GL_COEFMARG',0);
  if TOBOUV.detail.count > 0 then ReinitCOefMargDetail (TOBOUV);

end;

procedure SetAffaireSousDetail (TOBPiece,TOBOUvrage : TOB);
var Indice : integer;
		TOBL,TOBOUv : TOB;
begin
	for Indice := 0 To TOBPIece.detail.count -1 do
  begin
  	TOBL := TOBPIEce.detail[Indice];
  	if (TOBL.getValue('GL_INDICENOMEN')> 0) then
    begin
    	TOBOUV := TOBOUVrage.detail[TOBL.getValue('GL_INDICENOMEN')-1];
      if TOBOUV <> nil then SetAffaireSouDetailOuv (TOBPiece,TOBOuv);
    end;
  end;
end;

procedure InformeErreurOuvrage (FF : TForm ; TOBL : TOB; Arect : Trect);
var FFact : TFFacture;
  LastBrush, LastPen: Tcolor;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    LastBrush := GS.Canvas.Brush.Color;
    LastPen := GS.Canvas.Pen.Color;
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Brush.Color := clblack;
    GS.Canvas.Pen.Color := clblack;
    GS.Canvas.Polygon([Point(Arect.right -4, Arect.top), point(Arect.right, Arect.top), Point(Arect.right, Arect.top + 4)]);
    GS.Canvas.Brush.Color := LastBrush;
    GS.Canvas.Pen.Color := LastPen;
  end;
end;


procedure InformePTCOuvrage (FF : TForm ; TOBL : TOB; Arect : Trect);
var FFact : TFFacture;
  LastBrush, LastPen: Tcolor;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    LastBrush := GS.Canvas.Brush.Color;
    LastPen := GS.Canvas.Pen.Color;
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Brush.Color := clblack;
    GS.Canvas.Pen.Color := clRed;
    GS.Canvas.Polygon([Point(Arect.left, Arect.top), point(Arect.left, Arect.top+4), Point(Arect.left+4, Arect.top)]);
    GS.Canvas.Brush.Color := LastBrush;
    GS.Canvas.Pen.Color := LastPen;
  end;
end;




procedure CalculMontantsDocOuv(TOBPiece,TOBL,TOBOuvrage : TOB; CalculPv : boolean=true; FromImportBordereaux : boolean = false);
var IndiceNomen : integer;
		Valeur : T_Valeurs;
    Qte : double;
    TOBO : TOB;
    DEV : RDevise;
begin

	if TOBL = nil then exit;
  //
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  //
  //FV1 : 28/07/2016
  if IndiceNomen <= 0 then exit;
  if TobOuvrage.detail.count = 0 then exit;
  if Indicenomen-1 > TobOuvrage.detail.count - 1 then exit;
  //
  DEV.Code := TOBPiece.getValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.getValue('GP_TAUXDEV');
  //
  QTe := TOBL.GetValue ('GL_QTEFACT');

  InitTableau (Valeur);

  TOBO := TOBOuvrage.detail[IndiceNomen-1];

  GereCalculMontantOuv (TOBPiece,TOBL,TOBO,valeur,calculPv);
  TOBL.PutValue('GL_MONTANTPAFG', arrondi(Valeur[10] * Qte,4));
  TOBL.PutValue('GL_MONTANTPAFR', arrondi(Valeur[11] * Qte,4));

  TOBL.PutValue('GL_MONTANTPAFC', arrondi(Valeur[12] * Qte,4));

  TOBL.PutValue('GL_MONTANTPA', arrondi(Valeur[16] * Qte,V_PGI.okdecP));
	TOBL.PutValue('GL_MONTANTPR', arrondi(Valeur[17] * Qte,V_PGI.okdecP));
  if TOBL.GetValue('GL_QTEFACT')<>0 then
  begin
    TOBL.PutValue('GL_DPA', Arrondi(TOBL.GetValue('GL_MONTANTPA')/TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP));
    TOBL.PutValue('GL_DPR', Arrondi(TOBL.GetValue('GL_MONTANTPR')/TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecP));
  end;

	TOBL.PutValue('GL_MONTANTFG', arrondi(Valeur[13] * Qte,4));
	TOBL.PutValue('GL_MONTANTFR', arrondi(Valeur[14] * Qte,4));
	TOBL.PutValue('GL_MONTANTFC', arrondi(Valeur[15] * Qte,4));
	TOBL.PutValue('GL_TPSUNITAIRE', Valeur[9]);
  if (TOBL.GetValue('GL_BLOQUETARIF')='-') and (CalculPv) and (not FromImportBordereaux) then
  begin
    TOBL.PutValue('GL_PUHTDEV', arrondi(Valeur[2],V_PGI.okdecP));
    TOBL.PutValue('GL_PUTTCDEV', arrondi(Valeur[3],V_PGI.okdecP));
    TOBL.putvalue ('GL_PUHT',DEVISETOPIVOTEx(valeur[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
    TOBL.putvalue ('GL_PUTTC',DEVISETOPIVOTEx(valeur[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));

    TOBL.PutValue('GL_RECALCULER', 'X');
  end;
end;


procedure GereCalculMontantOuv (TOBPiece,TOBL,TOBOuvrage : TOb; var valeur : T_Valeurs; calculpv : boolean=true);
var TOBO : TOB;
		Indice : integer;
    Valloc : T_Valeurs;
    QteDuDetail,Qte,Coefdev,coefmarg : double;
    DEV : RDevise;
    MontantCharge : double;
begin
  TOBPIece := TOBL.Parent;
  DEV.Code := TOBPiece.getValue('GP_DEVISE');
  GetInfosDevise(DEV);
  DEV.Taux := TOBPiece.getValue('GP_TAUXDEV');
	for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
  	TOBO := TOBOuvrage.detail[Indice];
    Qte := TOBO.GetValue('BLO_QTEFACT');
    if Indice = 0 then QteDuDetail := TOBO.GetValue('BLO_QTEDUDETAIL');
    if QteDuDetail = 0  then QteDuDetail := 1;
    if TOBO.Detail.count > 0 then
    begin
		  InitTableau (Valloc);
    	GereCalculMontantOuv (TOBPiece,TOBL,TOBO,Valloc,calculPv);
      if calculPv then
      begin
        TOBO.putvalue ('BLO_PUHTDEV',arrondi(valloc[2],V_PGI.okdecP));
        TOBO.putvalue ('BLO_PUTTCDEV',arrondi(valloc[3],V_PGI.okdecP));
        //
        TOBO.putvalue ('BLO_PUHT',DEVISETOPIVOTEx(valloc[2],DEV.Taux,DEV.quotite,V_PGI.okdecP));
        TOBO.putvalue ('BLO_PUTTC',DEVISETOPIVOTEx(valloc[3],DEV.Taux,DEV.quotite,V_PGI.okdecP));
      end;
      //
      TOBO.PutValue('BLO_MONTANTPAFG',arrondi(Qte*Valloc[10],4));
      TOBO.PutValue('BLO_MONTANTPAFR',arrondi(Qte*Valloc[11],4));
      TOBO.PutValue('BLO_MONTANTPAFC',arrondi(Qte*Valloc[12],4));
      //
      TOBO.PutValue('BLO_MONTANTPA',arrondi(Qte*Valloc[16],V_PGI.okdecP));
      TOBO.PutValue('BLO_MONTANTPR',arrondi(Qte*Valloc[17],V_PGI.okdecP));
      //
      TOBO.PutValue('BLO_MONTANTFG',arrondi(Qte*Valloc[13],4));
      TOBO.PutValue('BLO_MONTANTFR',arrondi(Qte*Valloc[14],4));
      TOBO.PutValue('BLO_MONTANTFC',arrondi(Qte*Valloc[15],4));
      //
      if TOBO.GetValue('BLO_QTEFACT') <> 0 then
      begin
      	TOBO.putValue('BLO_DPA',Arrondi(TOBO.GetValue('BLO_MONTANTPA')/TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKDecP));
      	TOBO.putValue('BLO_DPR',Arrondi(TOBO.GetValue('BLO_MONTANTPR')/TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKDecP));
      end else
      begin
      	TOBO.putValue('BLO_DPA',TOBO.GetValue('BLO_MONTANTPA'));
      	TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
      end;

      TOBO.PutValue('BLO_TPSUNITAIRE',valloc[9]);
      if TOBPiece.getValue('GP_FACTUREHT')='X' then
      begin
        CalculeLigneHTOuv (TOBO,TOBPiece,DEV,true);
      end else
      begin
        CalculeLigneTTCOuv (TOBO,TOBPiece,DEV,true);
      end;
    end else
    begin
      Coefdev := TOBL.GetValue('GL_TAUXDEV');
      //
      TOBO.putValue('BLO_MONTANTPA',arrondi(TOBO.GetValue('BLO_DPA')*TOBO.GetValue('BLO_QTEFACT'),V_PGI.okdecP));
      TOBO.putValue('BLO_MONTANTPAFG',TOBO.GetValue('BLO_MONTANTPA'));
      TOBO.putValue('BLO_MONTANTPAFR',TOBO.GetValue('BLO_MONTANTPA'));
      TOBO.putValue('BLO_MONTANTPAFG',TOBO.GetValue('BLO_MONTANTPA'));
      if TOBO.GetValue('BLO_NONAPPLICFG')<> 'X' then TOBO.putValue('BLO_MONTANTFG',arrondi(TOBO.GetValue('BLO_MONTANTPAFG')*TOBO.GetValue('BLO_COEFFG'),4))
                                                else TOBO.putValue('BLO_MONTANTFG',0);
      MontantCharge := TOBO.getValue('BLO_MONTANTPA')+TOBO.GetValue('BLO_MONTANTFG');
      //
      if TOBO.GetValue('BLO_NONAPPLICFC')<> 'X' then TOBO.putValue('BLO_MONTANTFC',arrondi(MontantCharge*TOBO.GetValue('BLO_COEFFC'),4))
      																					else TOBO.putValue('BLO_MONTANTFC',0);
      MontantCharge := TOBO.getValue('BLO_MONTANTPA')+TOBO.GetValue('BLO_MONTANTFG')+TOBO.GetValue('BLO_MONTANTFC');
      //
      if TOBO.GetValue('BLO_NONAPPLICFRAIS')<> 'X' then TOBO.putValue('BLO_MONTANTFR',arrondi(MontantCharge*TOBO.GetValue('BLO_COEFFR'),4))
      																					   else TOBO.putValue('BLO_MONTANTFR',0);
      MontantCharge := TOBO.getValue('BLO_MONTANTPA')+TOBO.GetValue('BLO_MONTANTFG')+TOBO.GetValue('BLO_MONTANTFC')+TOBO.GetValue('BLO_MONTANTFR');
    	TOBO.putValue('BLO_MONTANTPR',MontantCharge);
      if TOBO.GetValue('BLO_QTEFACT') <> 0 then TOBO.putValue('BLO_DPR',Arrondi(TOBO.GetValue('BLO_MONTANTPR')/TOBO.GetValue('BLO_QTEFACT'),V_PGI.OKDecP))
                                           else TOBO.putValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
      //
      if (TOBL.GetValue('GL_BLOQUETARIF')='-') and (CalculPv) then
      begin
        if (TOBO.GetValue('BLO_TYPEARTICLE')<>'POU') and (TOBO.GetValue('BLO_COEFMARG') <> 0) and (TOBO.getValue('BLO_DPR')<>0) then
        begin
          TOBO.putValue('BLO_PUHT',Arrondi(TOBO.getValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG'),V_PGI.okdecP));
          TOBO.putValue('BLO_PUHTDEV',Arrondi(TOBO.getValue('BLO_DPR')*TOBO.GetValue('BLO_COEFMARG')*CoefDev,V_PGI.okdecP));
//          TOBO.putValue('BLO_RECALCULER','X');
        end;
        if TOBPiece.getValue('GP_FACTUREHT')='X' then
        begin
          CalculeLigneHTOuv (TOBO,TOBPiece,DEV,true);
        end else
        begin
          CalculeLigneTTCOuv (TOBO,TOBPiece,DEV,true);
        end;
      end else
      begin
        if (TOBO.GetValue('BLO_TYPEARTICLE')<>'POU') and (TOBO.getValue('BLO_DPR') <> 0) then
        begin
          CoefMarg := Arrondi(TOBO.GetValue('BLO_PUHT')/TOBO.getValue('BLO_DPR'),4);
          TOBO.putValue('BLO_COEFMARG',CoefMarg);
          TOBO.PutValue('POURCENTMARG',Arrondi((TOBO.GetValue('BLO_COEFMARG')-1)*100,2));
          if TOBO.GetValue('BLO_PUHT') <> 0 then
          begin
            TOBO.PutValue('POURCENTMARQ',Arrondi(((TOBO.GetValue('BLO_PUHT')- TOBO.GetValue('BLO_DPR'))/TOBO.GetValue('BLO_PUHT'))*100,2));
          end else
          begin
            TOBO.PutValue('POURCENTMARQ',0);
          end;
          //
          if (TOBL.getValue('GL_PIECEPRECEDENTE')<>'') and (Pos(TOBO.getValue('BLO_NATUREPIECEG'),'FBT;ABT;ABP;FPR;FAC;AVC;FBP;BAC')>0) then
          begin
            if Arrondi(TOBO.getValue('BLO_DPR')*CoefMarg,V_PGI.okdecP)<>TOBO.GetValue('BLO_PUHT') then
            begin
              TOBL.PutValue('GL_BLOQUETARIF','X');
            end;
          end else
          begin
            if (TOBL.GetValue('GL_BLOQUETARIF')='-') and (CalculPv) then
            begin
              TOBO.putValue('BLO_PUHT',Arrondi(TOBO.getValue('BLO_DPR')*coefmarg,V_PGI.okdecP));
              TOBO.putValue('BLO_PUHTDEV',Arrondi(TOBO.getValue('BLO_DPR')*CoefMarg*CoefDev,V_PGI.okdecP));
            end;
          end;

        end;
        if TOBPiece.getValue('GP_FACTUREHT')='X' then
        begin
          CalculeLigneHTOuv (TOBO,TOBPiece,DEV,true);
        end else
        begin
          CalculeLigneTTCOuv (TOBO,TOBPiece,DEV,true);
        end;
      end;
    end;

//    if (TOBL.GetValue('GL_BLOQUETARIF')='-') then
    begin
      //
      (*
      Valeur[2] := Valeur[2] + (TOBO.Getvalue ('BLO_MONTANTHTDEV')/QTeDuDEtail);
      Valeur[3] := Valeur[3] + (TOBO.Getvalue ('BLO_MONTANTTTCDEV')/QTeDuDEtail);
      *)
      //
      Valeur[2] := Valeur[2] + Arrondi((TOBO.GetValue('BLO_QTEFACT')*TobO.getValue('BLO_PUHTDEV'))/(TOBO.GetValue('BLO_PRIXPOURQTE')),V_PGI.okdecP);
      Valeur[3] := Valeur[3] + Arrondi((TOBO.GetValue('BLO_QTEFACT')*TobO.getValue('BLO_PUTTCDEV'))/(TOBO.GetValue('BLO_PRIXPOURQTE')),V_PGI.okdecP);
    end;
    //
    Valeur[16] := Valeur[16]+(TOBO.getValue('BLO_MONTANTPA'));
    Valeur[17] := Valeur[17]+(TOBO.getValue('BLO_MONTANTPR'));
    //
    Valeur[9] := Valeur[9]+(TOBO.GetValue('BLO_TPSUNITAIRE')*TOBO.getValue('BLO_QTEFACT'));
    //
    Valeur[13] := Valeur[13]+(TOBO.getValue('BLO_MONTANTFG'));
    Valeur[14] := Valeur[14]+(TOBO.getValue('BLO_MONTANTFR'));
    Valeur[15] := Valeur[15]+(TOBO.getValue('BLO_MONTANTFC'));
    //
    Valeur[10] := Valeur[10]+(TOBO.getValue('BLO_MONTANTPAFG'));
    Valeur[11] := Valeur[11]+(TOBO.getValue('BLO_MONTANTPAFR'));
    Valeur[12] := Valeur[12]+(TOBO.getValue('BLO_MONTANTPAFC'));
  end;
  //
  if QteDuDetail = 0  then QteDuDetail := 1;
  Valeur[2] := Arrondi(Valeur[2] /qteDudetail,V_PGI.okdecP);
  Valeur[3] := Arrondi(Valeur[3] /qteDudetail,V_PGI.okdecP);
  //
  Valeur[16] := Arrondi(Valeur[16] /qteDudetail,V_PGI.okdecP);
  Valeur[17] := Arrondi(Valeur[17] /qteDudetail,V_PGI.okdecP);
  //
  Valeur[9] := Arrondi(Valeur[9] /qteDudetail,V_PGI.okdecP);
  //
  Valeur[13] := Arrondi(Valeur[13] /qteDudetail,V_PGI.okdecP);
  Valeur[14] := Arrondi(Valeur[14] /qteDudetail,V_PGI.okdecP);
  Valeur[15] := Arrondi(Valeur[15] /qteDudetail,V_PGI.okdecP);
  //
  Valeur[10] := Arrondi(Valeur[10] /qteDudetail,V_PGI.okdecP);
  Valeur[11] := Arrondi(Valeur[11] /qteDudetail,V_PGI.okdecP);
  Valeur[12] := Arrondi(Valeur[12] /qteDudetail,V_PGI.okdecP);

end;


procedure BeforeApplicFraisDetailOuv (TOBL,TOBOuvrage : TOB; RazFg : boolean=false);

var IndiceNomen : integer;
		Valeur : T_Valeurs;
    Qte : double;
    TOBO : TOB;
    DEV : Rdevise;
begin

	if TOBL = nil then exit;
  QTe := TOBL.GetValue ('GL_QTEFACT');
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');

  if IndiceNomen <= 0 then exit;
  DEV.Code := TOBL.GetString('GL_DEVISE');
  GetInfosDevise (DEV);

  InitTableau (Valeur);

  TOBO := TOBOuvrage.detail[IndiceNomen-1];
  GereTraitementOuvrage (TOBO,valeur,DEV,false,RazFG);
  TOBL.PutValue('GL_TPSUNITAIRE', Valeur[9] );
	TOBL.putvalue('GL_HEURE',Arrondi(TOBL.getValue('GL_TPSUNITAIRE')*TOBL.GetValue('GL_QTEFACT'),V_PGI.okdecQ));
  TOBL.putvalue('GL_TOTALHEURE',TOBL.Getvalue('GL_HEURE'));

  TOBL.PutValue('GL_MONTANTPA',   arrondi(Valeur[16] * Qte,V_PGI.okdecP));

  TOBL.PutValue('GL_MONTANTPAFG', arrondi(Valeur[10] * Qte,V_PGI.okdecP));
  TOBL.PutValue('GL_MONTANTPAFR', arrondi(Valeur[11] * Qte,V_PGI.okdecP));
  TOBL.PutValue('GL_MONTANTPAFC', arrondi(Valeur[12] * Qte,V_PGI.okdecP));
  TOBL.PutValue('GL_MONTANTFG', arrondi(Valeur[13] * Qte,V_PGI.okdecP));

end;


procedure AppliqueModeGestionDetOuvFrais (TOBOUV : TOB;  ZoneAgerer : string ; ModeGestion : boolean);
var TheModeGestion : string;
    Indice : integer;
begin
  if ModeGestion then TheModeGestion := '-' else TheModeGestion := 'X';
  
  if ZoneAgerer = 'FR' then
  begin
    if IsprestationSt (TOBOUV) then
    begin
      TOBOUV.putValue('BLO_NONAPPLICFRAIS',TheModegestion);
    end;
  end else if ZoneAgerer = 'FC' then
  begin
    if IsprestationST (TOBOUV) then
    begin
      TOBOUV.putValue('BLO_NONAPPLICFC',TheModegestion);
    end;
  end;

  if TOBOUV.detail.count > 0 then
  begin
    For Indice := 0 to TOBOUV.detail.count -1 do
    begin
      AppliqueModeGestionDetOuvFrais (TOBOUV.detail[Indice],ZoneAgerer, ModeGestion);
    end;
  end;
end;

procedure AppliqueModeGestionOuvFg (TOBL,TOBOuvrage : TOB ; ZoneAgerer : string; ModeGestion : boolean);
var TOBO : TOB;
		IndiceNomen,Indice : integer;
    TheModeGestion : string;
begin

	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
  if IndiceNomen=0 then exit;
  TOBO := TOBOuvrage.detail[IndiceNomen-1];
  for Indice := 0 TO TOBO.detail.count -1 do
  begin
    AppliqueModeGestionDetOuvFrais (TOBO.detail[Indice], ZoneAgerer,ModeGestion);
  end;
end;

procedure ValideLesArticlesFromOuv (TOBArticles,TOBOuvrages : TOB);

  procedure ValideLesArticlesFromDetailOuv (TOBArticles,TOBOuvrage : TOB);
  var TOBA,TOBO : TOB;
      Article : string;
      Indice : integer;
  begin
    if TOBOUvrage.FieldExists ('BLO_ARTICLE') then
    begin
      Article := TOBOuvrage.GetValue('BLO_ARTICLE');
      TOBA := TOBArticles.FindFirst(['GA_ARTICLE'], [Article],true);
      if (TOBA <> nil) then
      begin
        if Not TOBA.FieldExists('UTILISE') then TOBA.AddChampSupValeur ('UTILISE','-');
      	if (TOBA.GetValue('SUPPRIME')<>'X') then TOBA.PutValue('UTILISE','X');
      end;
    end;
    if TOBOuvrage.detail.count > 0 then
    begin
      for Indice := 0 to TOBOUvrage.detail.count -1 do
      begin
        TOBO := TOBOuvrage.Detail[Indice];
        ValideLesArticlesFromDetailOuv (TOBArticles,TOBO);
      end;
    end;
  end;


var TOBOuvrage : TOB;
    Indice : integer;
begin
  for Indice := 0 to TOBOuvrages.detail.count -1 do
  begin
    TOBOuvrage := TOBOuvrages.detail[Indice];
    ValideLesArticlesFromDetailOuv (TOBArticles,TOBOuvrage);
  end;
end;

procedure GetOuvragePlat (TOBpiece,TOBL,TOBOuvrages,TOBPlat,TOBTiers : TOB;DEv : Rdevise; PourCpta: boolean = true; SansEcart : Boolean=false);
var Indice, IndiceEcart,NumOrdre,NumLigne : integer;
    MontantOuv,MontantLignes,MontantOuvAc,MontantLignesAc,puHtDev : double;
    EnHt : boolean;
    TOBLOc,TOBIns,TOBEC : TOB;
    Avancement : boolean;
    ArtEcart : String;
//    Qart : TQuery;
begin
  IndiceEcart := 0;
	if TOBOUvrages = nil then exit; // protection ..
  NumOrdre := TOBL.GetValue('GL_NUMORDRE');
  NumLigne := TOBL.GetValue('GL_NUMLIGNE');
  EnHt := TOBL.GetValue('GL_FACTUREHT')='X';
  Avancement := IsgestionAvanc (TOBpiece);
//  CalculeLigneHT(TOBL,nil,TOBPiece,DEV,DEV.Decimale ,False,TOBTiers) ;

  MontantLignes := 0; MontantLignesAc := 0;
  MiseAPlatOuv (TOBpiece,TOBL,TOBouvrages,TOBPlat,true,true,TRUE,false,PourCpta);
  if (EnHt) then
  begin
    if {(avancement) and }(TOBL.FieldExists('BLF_NATUREPIECEG') AND (TOBL.GetValue('BLF_NATUREPIECEG')<>'')) then
    begin
      TOBL.putValue('GL_MONTANTHTDEV',TOBL.GetValue('BLF_MTSITUATION'));
      MontantOuv := TOBL.GetValue('BLF_MTSITUATION');
    end else
    begin
      MontantOuv := Arrondi(TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_PUHTDEV')/ TOBL.GetValue('GL_PRIXPOURQTE'),DEv.decimale);
      TOBL.putValue('GL_MONTANTHTDEV',MontantOUv);
    end;
  end else
  begin
    MontantOuv := Arrondi(TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_PUTTCDEV')/ TOBL.GetValue('GL_PRIXPOURQTE'),DEv.decimale);
    TOBL.putValue('GL_MONTANTTTCDEV',MontantOUv);
  end;
  MontantOuvAc := Arrondi(TOBL.GetValue('GL_QTEFACT')*TOBL.GetValue('GL_DPA'),V_PGI.okdecP);
  TOBL.putValue('GL_MONTANTPA',MontantOUvAc);
  //
  if (TOBPlat=nil) Or (TOBPlat.detail.count = 0) then Exit;
  //
  for Indice := 0 to TOBPlat.detail.count -1 do
  begin
    TOBLoc := TOBPlat.detail[Indice];
    // CalculeLigneHT(TOBLOC,nil,nil,TOBPiece,DEV, DEV.decimale,true,TOBTiers) ;
    if EnHT then MontantLignes := MontantLignes + TOBLoc.GetValue('BOP_MONTANTHTDEV')
            else MontantLignes := MontantLignes + TOBLoc.GetValue('BOP_MONTANTTTCDEV');
    MontantLignesAc := MontantLignesAc + TOBLoc.GetValue('BOP_MONTANTPA');
    If (IndiceEcart=0) and (TOBLoc.GetValue('BOP_TYPEARTICLE') <> 'PRE') then IndiceEcart := Indice;
  end;
  //
  MontantLignes := Arrondi(MontantLignes,DEV.decimale);
  MontantLignesAc := Arrondi(MontantLignesAc,V_PGI.OkdecP);

  if (not SansEcart ) and ((MontantOuv <> MontantLignes) or (MontantOuvAc <> MontantLignesAc)) then
  begin
    TOBLOc := TOBPlat.detail[IndiceEcart];
    TOBIns:=NewTobLignePlat(TOBPLat,NumOrdre,-1);
    copyLigne (TOBIns,TOBLoc);
    ZeroLigne (TOBIns);
    TOBIns.putValue('BOP_PUHTDEV',  0);
    TOBIns.putValue('BOP_PUTTCDEV', 0);
    TOBIns.putValue('BOP_DPA',      0);
    if (valeur(TOBL.getValue('GLC_NATURETRAVAIL'))>0) and (valeur(TOBL.getValue('GLC_NATURETRAVAIL'))<10) then
    begin
    	TOBIns.putValue('BOP_NATURETRAVAIL',TOBL.getvalue('GLC_NATURETRAVAIL'));
    	TOBIns.putValue('BOP_FOURNISSEUR',  TOBL.getvalue('GL_FOURNISSEUR'));
    end else
    begin
    	TOBIns.putValue('BOP_NATURETRAVAIL','');
    	TOBIns.putValue('BOP_FOURNISSEUR','');
    end;
    TOBIns.putValue('BOP_REMISABLELIGNE','-');
    TOBIns.putValue('BOP_REMISELIGNE',0);
    TOBIns.putValue('BOP_PRIXPOURQTE',1);
    TOBIns.putValue('BOP_PUHT',0); TOBIns.putValue('BOP_PUTTC',0);
    TOBIns.putValue('BOP_COEFFG',0);
    TOBIns.putValue('BOP_COEFFC',0);
    TOBIns.putValue('BOP_COEFFR',0);
    TOBIns.putValue('BOP_COEFMARG',0);
    TOBIns.putValue('BOP_MONTANTFG',0);
    TOBIns.putValue('BOP_MONTANTFC',0);
    TOBIns.putValue('BOP_MONTANTFR',0);
    TOBIns.putValue('BOP_MONTANTPAFG',0);
    TOBIns.putValue('BOP_MONTANTPAFC',0);
    TOBIns.putValue('BOP_MONTANTPAFR',0);
    TOBIns.putValue('BOP_MONTANTPA',0);
    TOBIns.putValue('BOP_MONTANTPR',0);
    TOBIns.putValue('BOP_DPA',0);
    TOBIns.putValue('BOP_DPR',0);

    if (MontantOuv <> MontantLignes) then
    begin
      PuHTDev := arrondi(MontantOuv-MontantLignes,V_PGI.okdecP);
      if EnHt then
      begin
        TOBIns.putvalue ('BOP_PUHTDEV',PuHtDev);
        TOBIns.putvalue ('BOP_PUHT',DEVISETOPIVOTEx(PuHTDEV,TOBIns.GetValue('BOP_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
        TOBIns.putvalue ('BOP_PUHTNETDEV',PuHtDev);
        TOBIns.putvalue ('BOP_PUHTNET',DEVISETOPIVOTEx(PuHTDEV,TOBIns.GetValue('BOP_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
      end else
      begin
        TOBIns.putvalue ('BOP_PUTTCDEV',PuHtDev);
        TOBIns.putvalue ('BOP_PUTTC',DEVISETOPIVOTEx(PuHTDEV,TOBIns.GetValue('BOP_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
        TOBIns.putvalue ('BOP_PUTTCNETDEV',PuHtDev);
        TOBIns.putvalue ('BOP_PUTTCNET',DEVISETOPIVOTEx(PuHTDEV,TOBIns.GetValue('BOP_TAUXDEV'),DEV.quotite,V_PGI.okdecP));
      end;
    end;
    // Correction FQ 11984 --> tva incorrecte sur article d'ecart
    TobIns.PutValue ('BLO_REGIMETAXE',TOBLoc.GetValue('BLO_REGIMETAXE'));
    Tobins.PutValue ('BLO_TVAENCAISSEMENT',TOBloc.GetValue('BLO_TVAENCAISSEMENT'));
    for Indice := 1 to 5 do Tobins.PutValue ('BLO_FAMILLETAXE'+InttoStr(Indice),TOBloc.GetValue('BLO_FAMILLETAXE'+InttoStr(Indice)));
    //
    TOBIns.putvalue ('BOP_PRIXPOURQTE',1);
    TOBIns.putvalue ('BOP_QTEFACT',1);
    if (MontantOuvAc <> MontantLignesAc) then
    begin
      TOBIns.putvalue ('BOP_DPA',MontantOuvAc-MontantLignesAc);
      CalculeLigneAc (TOBIns, DEV, False);
    end;
    if TOBIns.fieldExists ('BOP_PMAP') then
    begin
      TOBIns.putvalue ('BOP_PMAP',0);
      TOBIns.putvalue ('BOP_PMRP',0);
    end;

    // Traitement affectation écart sur article d'écart si mise à plat non comptable (ex : minute)
    ArtEcart := GetParamsoc('SO_BTECARTPMA');
    TOBEC := GetInfoArtEcart (ArtEcart);
    if (Not PourCpta) and (TOBEC <> nil) then
    begin
//      Qart := opensql('Select GA_ARTICLE, GA_LIBELLE from ARTICLE Where GA_CODEARTICLE="' + ArtEcart + '"', true,-1, '', True);
//      if not Qart.eof then
//      begin
        TobIns.putvalue('BOP_CODEARTICLE',ArtEcart);
        TobIns.putvalue('BOP_REFARTSAISIE',ArtEcart);
//        ArtEcart := Qart.fields[0].AsString;
        TobIns.putvalue('BOP_ARTICLE',TOBEC.GetString('GA_ARTICLE'));
//        ArtEcart := Qart.fields[1].AsString;
        TobIns.putvalue('BOP_LIBELLE',TOBEC.GetString('GA_LIBELLE'));
//        ferme (Qart);
//      end;
    end;
    StockeMontantTypeSurLigne (TOBIns);

  end;
end;

function FindOuvragesPlat(TOBL, TOBOuvrageP: TOB ) : TOB;
var prefixe : string;
begin
	if TOBouvrageP = nil then BEGIN result := nil;exit; END;
  prefixe := GetPrefixeTable (TOBL);
  Result := TOBOuvrageP.findFirst(['NUMORDRE'],[TOBL.GetValue(prefixe+'_NUMORDRE')],true);
end;

procedure LoadLesOuvragesPlat(TOBPiece, TOBOuvragesP: TOB; Cledoc : R_CLEDOC );
var NaturePieceG: string;
  GereAcompte: Boolean;
  QQ: TQuery;
  TOBL,TOBRef: TOB;
  Indice,IndNumOrdre: integer;
  NumOrdreCur : Integer;
begin
  NumOrdreCur := 0;
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  QQ := OpenSQl('SELECT * FROM LIGNEOUVPLAT WHERE ' + WherePiece(CleDoc, ttdOuvrageP, False)+' ORDER BY BOP_NUMORDRE', True,-1, '', True);
  if not QQ.EOF then TOBOuvragesP.LoadDetailDB('LIGNEOUVPLAT', '', '', QQ, False);
  Ferme(QQ);
  if TOBOuvragesP.detail.count = 0 then exit;
  Indice := 0;
  repeat
    TOBL := TOBOuvragesP.detail[Indice];
    if Indice = 0 then
    begin
      IndNumOrdre := TOBL.GetNumChamp('BOP_NUMORDRE');
    end;
    if TOBL.NomTable = 'LIGNEOUVPLAT' then
    begin
      if TOBL.GetValeur(IndNumOrdre) <> NumOrdreCur then
      begin
        TOBRef := AddMereLignePlat (TOBOuvragesP,TOBL.GetValue('BOP_NUMORDRE'));
        NumOrdreCur := TOBL.GetValeur(IndNumOrdre);
      end;
      TOBL.ChangeParent (TOBRef,-1);
      AddLesSupLigne(TOBL, False);
      InitLesSupLigne(TOBL);
    end else break;
  until indice > TOBOuvragesP.detail.count -1;
end;

function AddMereLignePlat (TOBMere: TOB;NumOrdre : integer) : TOB;
begin
  Result := TOB.create ('UNE LIGNE',TOBmere,-1);
  Result.AddChampSupValeur ('NUMORDRE',NumOrdre);
  Result.AddChampSupValeur ('OKOK','-');
end;

function NewTobLignePlat (TOBmere: TOB; NumOrdre : integer; position : integer;SurLigneStd : boolean=false) : TOB;
var TOBLigne : TOB;
begin
  if SurLigneStd then
  begin
    result := TOB.create('LIGNE',TOBMere,position);
  end else
  begin
    result := TOB.create('LIGNEOUVPLAT',TOBMere,position);
  end;
  NewTOBLigneFille(result);
  AddLesSupLigne(result, False);
  InitLesSupLigne(result);
  result.AddChampSupValeur('UNIQUEBLO',0)
end;

procedure ValideLesOuvPlat(TOBOuvragesP, TOBPiece : TOB);

  procedure ValideLesOuvPlatDet(TOBP, TOBPiece : TOB);
  var Indice : integer;
    TOBB : TOB;
  begin
    for Indice := 0 to TOBP.detail.count -1 do
    begin
      TOBB := TOBP.detail[Indice];
        TOBB.PutValue('BOP_NATUREPIECEG',TOBPiece.GetValue('GP_NATUREPIECEG')) ;
        TOBB.PutValue('BOP_DATEPIECE',TOBPiece.GetValue('GP_DATEPIECE')) ;
        TOBB.PutValue('BOP_SOUCHE',TOBPiece.GetValue('GP_SOUCHE')) ;
        TOBB.PutValue('BOP_NUMERO',TOBPiece.GetValue('GP_NUMERO')) ;
        TOBB.PutValue('BOP_INDICEG',TOBPiece.GetValue('GP_INDICEG')) ;
      if TOBB.GetValue('BOP_COEFMARG')>99999 then TOBB.putValue('BOP_COEFMARG',0); // protection du coef de marge
      if TOBB.GetValue('BOP_COEFFG')>99999 then TOBB.putValue('BOP_COEFG',0); // protection du coef de frais généraux
      if TOBB.GetValue('BOP_COEFFC')>99999 then TOBB.putValue('BOP_COEFC',0); // protection du coef de frais généraux
      if TOBB.GetValue('BOP_COEFFR')>99999 then TOBB.putValue('BOP_COEFR',0); // protection du coef de frais répartis
      end;
  end;

var Indice,II : integer;
    TOBB,TOBL : TOB;
begin

  for Indice := 0 to TOBpiece.detail.count -1 do
    begin
    TOBL := TOBpiece.detail[Indice];
    if IsOuvrage(TOBL) then
      begin
      For II := 0 to TOBOuvragesP.detail.count -1 do
      begin
        TOBB := TOBOuvragesP.detail[II];
        if TOBB.GetInteger('NUMORDRE')=TOBL.GetInteger('GL_NUMORDRE') then
        begin
          TOBB.SetString('OKOK','X');
          break;
      end;
    end;
  end;
  end;

  if (TOBOuvragesP.detail.count > 0) then
  begin
    Indice := 0;
    repeat
      TOBB := TOBOuvragesP.detail[Indice];

      if TOBB.GetString('OKOK')<>'X' then
      begin
        TOBB.free;
        continue;
      end;

      if (TOBB.detail.count > 0) then
      begin
        ValideLesOuvPlatDet (TOBB,TOBPiece);
      end;
      Inc(Indice);
    until Indice >= TOBOuvragesP.detail.count;
  end;

end;

procedure AppliqueCoefMargDetail (TOBL,TOBOuvrage: TOB; CoefMarg : double; DEV : Rdevise);
var Indice : integer;
begin
  if TOBOuvrage.NomTable = 'LIGNEOUV' then
  begin
    TOBOUVrage.putValue('BLO_COEFMARG',CoefMarg);
    TOBOUVrage.PutValue('POURCENTMARG',Arrondi((TOBOUVrage.GetValue('BLO_COEFMARG')-1)*100,2));
    if TOBOUVrage.GetValue('BLO_PUHT') <> 0 then
    begin
      TOBOUVrage.PutValue('POURCENTMARQ',Arrondi(((TOBOUVrage.GetValue('BLO_PUHT')- TOBOUVrage.GetValue('BLO_DPR'))/TOBOUVrage.GetValue('BLO_PUHT'))*100,2));
    end else
    begin
      TOBOUVrage.PutValue('POURCENTMARQ',0);
    end;
  end;
  if TOBOUvrage.detail.count > 0 then
  begin
    for Indice := 0 to TOBOUvrage.detail.count -1 do
    begin
      AppliqueCoefMargDetail(TOBL,TOBouvrage.detail[Indice],Coefmarg,DEV);
    end;
  end;
  if TOBOuvrage.NomTable = 'LIGNEOUV' then CalculeLigneAcOuv(TOBOuvrage,DEV,True); 
end;

procedure AppliqueChangeFraisDetOuv (TOBPiece,TOBO : TOB; NewGestionfrais,NewgestionFC,NewgestionFA : string);
var indice : integer;
    TOBD  : TOB;
    ApplicationFrais : string;
begin
  if IsPrestationSt(TOBO) then
  begin
    if (NewGestionFrais='-') and (TOBPiece.getValue('GP_APPLICFGST')='X') then ApplicationFrais:= '-'
                                                                          else ApplicationFrais:= 'X';
    TOBO.putValue('BLO_NONAPPLICFRAIS',ApplicationFrais);
    if (NewGestionFC='-') and (TOBPiece.getValue('GP_APPLICFCST')='X') then ApplicationFrais:= '-'
                                                                       else ApplicationFrais:= 'X';
    TOBO.putValue('BLO_NONAPPLICFC',ApplicationFrais);
  end else
  begin
    TOBO.putValue('BLO_NONAPPLICFRAIS',Newgestionfrais);
    TOBO.putValue('BLO_NONAPPLICFC',NewgestionfC);
  end;
  TOBO.putValue('BLO_NONAPPLICFG',NewgestionfA);
//    if newGestionFrais = 'X' then TOBO.PutValue('BLO_DPR',TOBO.GetValue('BLO_DPA'));
  if TOBO.Detail.count > 0 then
  begin
    for Indice := 0 to TOBO.detail.count -1 do
    begin
      TOBD := TOBO.detail[Indice];
      AppliqueChangeFraisDetOuv (TOBPiece,TOBD,NewGestionfrais,NewgestionFC,NewgestionFA);
    end;
  end;
end;

procedure recalculCoefligneMere (TOBOUV : TOB);
begin

end;

procedure AffecteTempsPoseOuvrage (TOBArticles,TOBL: TOB ; temps : double; DEV : RDevise);
var TOBA : TOB;
		Valeurs : T_Valeurs;
    Qte : double;
begin
	Qte := TOBL.GetValue('BLO_QTEFACT');
	TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValue('BLO_ARTICLE')],true);
  TOBL.detail[1].putValue('BLO_QTEFACT',temps);
  if ((Pos(TOBL.detail[1].GetValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne)>0) and (not IsLigneExternalise(TOBL))) then
  begin
  	TOBL.detail[1].putValue('BLO_TPSUNITAIRE',1);
  end else
  begin
  	TOBL.detail[1].putValue('BLO_TPSUNITAIRE',0);
  end;
  TOBL.Putvalue('GA_HEURE',temps);
  TOBL.Putvalue('BLO_TPSUNITAIRE',temps);
  CalculeLigneAcOuv (TOBL.detail[1],DEV,true,TOBA);
  //RecupValoDetailHrs (TOBLIG);
  CalculOuvFromDetail (TOBL,DEV,Valeurs);
  TOBL.Putvalue ('BLO_DPA',Valeurs[0]);
  TOBL.Putvalue ('BLO_DPR',Valeurs[1]);
  TOBL.Putvalue ('BLO_PUHT',Valeurs[2]);
  TOBL.Putvalue ('BLO_PUHTBASE',Valeurs[2]);
  TOBL.Putvalue ('BLO_PUHTDEV',pivottodevise(Valeurs[2],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
  TOBL.Putvalue ('BLO_PUTTC',Valeurs[3]);
  TOBL.Putvalue ('BLO_PUTTCDEV',pivottodevise(Valeurs[3],DEV.Taux,DEV.quotite,V_PGI.OkdecP));
  TOBL.PutValue('ANCPV', TOBL.GetValue('BLO_PUHTDEV'));
  TOBL.Putvalue ('BLO_PMAP',Valeurs[6]);
  TOBL.Putvalue ('BLO_PMRP',Valeurs[7]);
  //
  TOBL.PutValue('BLO_TPSUNITAIRE',Valeurs[9]);
  //TOBL.PutValue('GA_HEURE',Valeurs[9]);
  TOBL.PutValue('BLO_MONTANTFG',Arrondi(Valeurs[13]*Qte,4));
  TOBL.PutValue('BLO_MONTANTFC',Arrondi(Valeurs[15]*Qte,4));
  TOBL.PutValue('BLO_MONTANTFR',Arrondi(Valeurs[14]*Qte,4));
  //
  CalculeLigneAcOuvCumul (TOBL);
  CalculMontantHtDevLigOuv (TOBL,DEV);

end;

function GetQualifTps(TOBL,TOBOUvrage : TOB) : string;
var TOBOP,TOBO : TOB;
begin
	result := '';
	if TOBL.getValue('GL_INDICENOMEN')=0 then exit;
  TOBO := TOBOUvrage.detail[TOBL.getValue('GL_INDICENOMEN')-1];
  TOBOP := TOBO.detail[0];
  if TOBOP.detail.count > 1 then
    result := TOBOP.detail[1].GetValue('BLO_QUALIFQTEVTE');
end;

procedure AjoutePrestationPrixPose (TOBPiece,TOBLigPiece,TOBOP,TOBArticles : TOB; PrestDefaut : string; DEV : Rdevise);

	procedure RenseigneTOBOuv ( TOBPiece,TOBL,TOBA : TOB; DEV : Rdevise);
  begin
   	CopieOuvFromLigne (TOBL,TOBLigPiece);
    //
    TOBL.PutValue('BLO_TENUESTOCK',TobA.GetValue('GA_TENUESTOCK'));
    TOBL.PutValue('BLO_QUALIFQTEVTE',TobA.GetValue('GA_QUALIFUNITEVTE'));
    TOBL.PutValue('BLO_QUALIFQTESTO',TobA.GetValue('GA_QUALIFUNITESTO'));
    TOBL.PutValue('BLO_PRIXPOURQTE',TobA.GetValue('GA_PRIXPOURQTE'));
//    TOBL.PutValue('BLO_DPA',TobA.GetValue('GA_PAHT'));
    if (VH_GC.ModeValoPa = 'PMA') and (pos(TOBA.getString('GA_TYPEARTICLE'),'MAR;ARP')>0) then
    begin
			if TOBL.GetDouble('GA_PMAP')<>0 then
      begin
    		TOBL.PutValue('BLO_DPA',TobA.GetValue('GA_PMAP'));
      end else
      begin
    		TOBL.PutValue('BLO_DPA',TobA.GetValue('GA_PAHT'));
      end;
    end else
    begin
    	TOBL.PutValue('BLO_DPA',TobA.GetValue('GA_PAHT'));
    end;
    if TOBA.GetValue('GA_DPRAUTO')='X' then
    begin
    	TOBL.PutValue('BLO_COEFFG',TobA.GetValue('GA_COEFFG'));
    end else TOBL.PutValue('BLO_COEFFG',0);

    if TOBA.GetValue('GA_CALCAUTOHT')='X' then
    begin
    	TOBL.PutValue('GL_COEFMARG', TOBA.GetValue('GA_COEFCALCHT'));
      TOBL.PutValue('POURCENTMARG',Arrondi((TOBL.GetValue('GL_COEFMARG')-1)*100,2));
    end else TOBL.PutValue('GL_COEFMARG',0);

    TOBL.PutValue('GL_DOMAINE', TOBA.GetValue('GA_DOMAINE'));
    if TOBPiece.getValue('GP_DOMAINE')<>'' then TOBL.putValue('GL_DOMAINE',TOBPIece.getValue('GP_DOMAINE'));
    TOBL.PutValue('BLO_PMAP',TobA.GetValue('GA_PMAP'));
    CopieOuvFromArt (TOBPiece,TOBL,TobA,DEV);
    TOBL.PutValue('BLO_PUHT',TobA.GetValue('GA_PVHT'));
    TOBL.PutValue('BLO_PUHTBASE',TobA.GetValue('GA_PVHT'));
    TOBL.PutValue('BLO_PUHTDEV',pivottodevise(TobL.GetValue('BLO_PUHTBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
    TOBL.PutValue('ANCPV',TobL.GetValue('BLO_PUHTDEV'));
    TOBL.PutValue('BLO_PUTTC',TobA.GetValue('GA_PVTTC'));
    TOBL.PutValue('BLO_PUTTCBASE',TobA.GetValue('GA_PVTTC'));
    TOBL.PutValue('BLO_PUTTCDEV',pivottodevise(TOBL.GetValue('BLO_PUTTCBASE'),DEV.Taux,DEV.quotite,V_PGI.OkdecP ));
   	TOBL.putValue('BNP_TYPERESSOURCE',TOBA.getValue('BNP_TYPERESSOURCE'));
    //
    if (Pos (TOBL.getValue('BNP_TYPERESSOURCE'),VH_GC.BTTypeMoInterne) > 0)  and (not IsLigneExternalise(TOBL)) then
    begin
      TOBL.PutValue('BLO_TPSUNITAIRE',1);
    end else
    begin
      TOBL.PutValue('BLO_TPSUNITAIRE',0);
    end;
    //
    if (IsPrestationSt(TOBL)) then
    begin
      if TOBPiece.getValue('GP_APPLICFGST')='X' then
      begin
        TOBL.putValue('BLO_NONAPPLICFRAIS','-');
      end else
      begin
        TOBL.putValue('BLO_NONAPPLICFRAIS','X');
      end;
      if TOBPiece.getValue('GP_APPLICFCST')='X' then
      begin
        TOBL.putValue('BLO_NONAPPLICFC','-');
      end else
      begin
        TOBL.putValue('BLO_NONAPPLICFC','X');
      end;
    end else
    begin
      TOBL.putValue('BLO_NONAPPLICFRAIS','-');
      TOBL.putValue('BLO_NONAPPLICFC','-');
    end;
    if TobA.getValue('GA_PRIXPASMODIF')<>'X' Then
    begin
      AppliqueCoefDomaineActOuv (TOBL);
      if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBL,DEV);
    end;
    CalculeLigneAcOuv (TOBL,DEV,true,TOBA);
    if TOBL.GetValue('BLO_PUHT') <> 0 then
    begin
      TOBL.PutValue('POURCENTMARQ',Arrondi(((TOBL.GetValue('BLO_PUHT')- TOBL.GetValue('BLO_DPR'))/TOBL.GetValue('BLO_PUHT'))*100,2));
    end else
    begin
      TOBL.PutValue('POURCENTMARQ',0);
  end;

  end;

var TOBL : TOB;
		TOBA : TOB;
    QQ : Tquery;
begin
  TOBL := TOB.create('LIGNEOUV',TOBOP,-1);
  InsertionChampSupOuv (TOBL,false);
  TOBA := TOBArticles.FindFirst(['GA_CODEARTICLE'],[PrestDefaut],true);
  if TOBA = nil then
  begin
    QQ := OPenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                    'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                    'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE GA_CODEARTICLE ="'+PrestDefaut+'"',True,-1, '', True);
    TOBA:= TOB.create ('ARTICLE',TOBArticles,-1);
    TOBA.selectDB ('',QQ);
    ferme (QQ);
    InitChampsSupArticle (TOBA);
  end;
  // mise en place de la nouvelle prestation
  TOBL.putValue('BLO_ARTICLE',TOBA.GetValue('GA_ARTICLE'));
  TOBL.putValue('BLO_CODEARTICLE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.putValue('BLO_REFARTSAISIE',TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.putValue('BLO_LIBELLE',TOBA.GetValue('GA_LIBELLE'));
  RenseigneTOBOuv (TOBpiece,TOBL,TOBA,DEV);
	CalculeLigneAcOuv (TOBL,DEV,true,TOBA);
  NumeroteLigneOuv (TOBOP,nil,2,1,0,0,0);
  TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
  TOBL.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
end;


function OuvrageNonDifferencie (TOBPiece,TOBOuvrage : TOB) : boolean;
begin
	result :=  (TOBPiece.getValue('GP_UNIQUEBLO')=1) and (TOBOuvrage.detail.count > 0);
end;

procedure OuvrageDifferencie (TOBpiece,TOBOuvrage : TOB; TraiteAvancOuv : boolean; DEV : RDevise);

	procedure  IndiceUniqueBlo (TOBouv : TOB; var  UniqueBlo : integer);
  var Indice : integer;
  		TOBL : TOB;
  begin

    For Indice := 0 to TOBouv.detail.count -1 do
    begin
    	Inc(UniqueBlo);
      //
      TOBOUV.detail[Indice].putValue('BLO_UNIQUEBLO',UniqueBlo);
      if TOBOUV.detail[Indice].detail.count> 0 then
      begin
      	IndiceUniqueBlo (TOBOUV.detail[Indice],uniqueBLO);
      end;
    end;
  end;

var Indice,UniqueBLO : integer;
begin
	Indice := 0; UniqueBLO := 0;
  for Indice := 0 to TOBouvrage.detail.count -1 do
  begin
  	IndiceUniqueBlo (TOBouvrage.detail[Indice], UniqueBlo);
  end;
  TOBPiece.putvalue('GP_UNIQUEBLO',UniqueBlo);
end;

function EncodeRefPieceOuv (TOBO : TOB) : string;
var DD : TDateTime;
		Std : string;
begin
  DD := TOBO.GetValue('BLO_DATEPIECE');
  StD := FormatDateTime('ddmmyyyy', DD);
  Result := StD + ';' + TOBO.GetValue('BLO_NATUREPIECEG') + ';' + TOBO.GetValue('BLO_SOUCHE') + ';'
            + IntToStr(TOBO.GetValue('BLO_NUMERO')) + ';' + IntToStr(TOBO.GetValue('BLO_INDICEG')) + ';';
  result := result + IntToStr(TOBO.GetValue('BLO_UNIQUEBLO')) + ';'; { NEWPIECE }
end;

procedure DecodeRefPieceOUv(St: string; var CleDoc: R_CleDoc); { NEWPIECE }
var
  StC, StL: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  StC := St;
  CleDoc.DatePiece := EvalueDate(ReadTokenSt(StC));
  CleDoc.NaturePiece := ReadTokenSt(StC);
  CleDoc.Souche := ReadTokenSt(StC);
  CleDoc.NumeroPiece := StrToInt(ReadTokenSt(StC));
  CleDoc.Indice := StrToInt(ReadTokenSt(StC));
  StL := ReadTokenSt(StC);
  if StL <> '' then
  begin
    CleDoc.UniqueBLO := StrToInt(StL);
  end;
end;


procedure SetAvancementOuvrage(TOBGroupeNomen,TOBOD,TOBL : TOB; DEV : RDevise);
var QteDet,MtDetDev,avancprec,QteDejaFact,MtDejafact : double;
		AvancSousDetailPartiel : boolean;
begin
	if TOBOD.GetValue('BLO_QTEDUDETAIL') = 0 then TOBOD.PutValue('BLO_QTEDUDETAIL',1) ;
	if TOBOD.getValue('BLF_NATUREPIECEG')='' then
  begin
  	AvancSousDetailPartiel := (TOBGroupeNomen.getValue('AVANCSOUSDET')='X');
    if AvancSousDetailPartiel then
    begin
      QteDet := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEFACT') /TOBOD.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      QteDejaFact := 0;
      MtDetDev := arrondi(QteDet*TOBOD.getValue('BLO_PUHTDEV'),DEV.Decimale);
      MtDejaFact := 0;
      avancPrec := 0;
    end else
    begin
    // cas de figure ou il n'y a pas d'avancement au niveau détail d'ouvrage --> mais bien au niveau ligne de piece
      QteDet := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEPREVAVANC') /TOBOD.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      if TOBL.Getvalue('BLF_MTMARCHE') <> 0 then
      begin
//        avancprec := (TOBL.Getvalue('BLF_MTDEJAFACT')- TOBL.Getvalue('BLF_MTSITUATION'))/TOBL.Getvalue('BLF_MTMARCHE'); // ratio pour deja facturé
        avancprec := TOBL.Getvalue('BLF_MTDEJAFACT')/TOBL.Getvalue('BLF_MTMARCHE'); // ratio pour deja facturé
      end else
      begin
        Avancprec := 0;
      end;
      QteDejaFact := Arrondi(QteDet * Avancprec,V_PGI.okdecQ);
      MtDetDev := arrondi(QteDet*TOBOD.getValue('BLO_PUHTDEV'),DEV.Decimale);
      MtDejafact := arrondi(QteDejafact*TOBOD.getValue('BLO_PUHTDEV'),DEV.Decimale);
    end;
    TOBOD.putvalue('BLF_UNIQUEBLO',TOBOD.getValue('BLO_UNIQUEBLO'));
		//
    if AvancSousDetailPartiel then
    begin
      TOBOD.putValue('BLF_MTMARCHE',MtDetDev);
      TOBOD.putValue('BLF_QTEMARCHE',QteDet);
    end else
    begin
    // cas de figure ou il n'y a pas d'avancement au niveau détail d'ouvrage --> mais bien au niveau ligne de piece
      TOBOD.putValue('BLF_POURCENTAVANC',TOBL.Getvalue('BLF_POURCENTAVANC'));
      //
      QteDejaFact := arrondi(TOBOD.getValue('BLO_QTEFACT')*TOBL.getvalue('GL_QTEPREVAVANC') /TOBOD.getValue('BLO_QTEDUDETAIL'),V_PGI.okdecQ);
      //
      TOBOD.putValue('BLF_MTMARCHE',MtDetDev);
      TOBOD.putValue('BLF_MTCUMULEFACT',arrondi(MtDetDev*TOBL.Getvalue('BLF_POURCENTAVANC')/100,DEV.Decimale ));
      TOBOD.putValue('BLF_MTDEJAFACT',arrondi(MtDetDev*avancprec,DEV.decimale));
      TOBOD.putValue('BLF_MTSITUATION',arrondi(TOBOD.getValue('BLF_MTCUMULEFACT')-TOBOD.getValue('BLF_MTDEJAFACT'),DEV.Decimale));
      TOBOD.putValue('BLF_QTEMARCHE',QteDet);
      TOBOD.putValue('BLF_QTECUMULEFACT',arrondi(QteDet*TOBOD.Getvalue('BLF_POURCENTAVANC')/100,V_PGI.okdecQ));
      TOBOD.putValue('BLF_QTEDEJAFACT',arrondi(QteDet*avancprec,V_PGI.okdecQ));
      TOBOD.putValue('BLF_QTESITUATION',arrondi(TOBOD.GetValue('BLF_QTECUMULEFACT')-TOBOD.GetValue('BLF_QTEDEJAFACT'),V_PGI.okdecQ));
    end;
    TOBOD.putValue('BLF_NATUREPIECEG',TOBL.Getvalue('GL_NATUREPIECEG'));
    TOBOD.putValue('BLF_SOUCHE',TOBL.Getvalue('GL_SOUCHE'));
    TOBOD.putValue('BLF_NUMERO',TOBL.Getvalue('GL_NUMERO'));
    TOBOD.putValue('BLF_INDICEG',TOBL.Getvalue('GL_INDICEG'));
    TOBOD.putValue('BLF_NUMORDRE',0);
    TOBOD.putValue('BLF_UNIQUEBLO',TOBOD.Getvalue('BLO_UNIQUEBLO'));
    TOBOD.putValue('BLF_NATURETRAVAIL',TOBOD.Getvalue('BLO_NATURETRAVAIL'));
    TOBOD.putValue('BLF_FOURNISSEUR',TOBOD.Getvalue('BLO_FOURNISSEUR'));
    //
  end;
end;

procedure SetLignesFactureOuvDetail (TOBL,TOBOuvrage,TOBLFactures : TOB);
var indiceNomen,II : integer;
		TOBOUV,TOBOO,TOBD : TOB;
begin
	IndiceNomen := TOBL.GetValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrage.detail[IndiceNomen-1]; if TOBOUV = nil then exit;
  for II := 0 to TOBOUV.detail.count -1 do
  begin
  	TOBOO := TOBOUV.detail[II];
    if (TOBOO.GetValue('BLF_MTDEJAFACT')=0) and (TOBOO.GetValue('BLF_MTSITUATION')=0) then continue;
  	TOBD := TOB.Create ('LIGNEFAC',   TOBLFactures,-1);
    TOBD.putValue('BLF_NATUREPIECEG', TOBL.GetValue('GL_NATUREPIECEG'));
    TOBD.putValue('BLF_SOUCHE',       TOBL.GetValue('GL_SOUCHE'));
    TOBD.putValue('BLF_DATEPIECE',    TOBL.GetValue('GL_DATEPIECE'));
    TOBD.putValue('BLF_NUMERO',       TOBL.GetValue('GL_NUMERO'));
    TOBD.putValue('BLF_INDICEG',      TOBL.GetValue('GL_INDICEG'));
    TOBD.putValue('BLF_NUMORDRE',0);
    TOBD.putValue('BLF_UNIQUEBLO',    TOBOO.GetValue('BLO_UNIQUEBLO'));
    TOBD.putValue('BLF_MTMARCHE',     TOBOO.GetValue('BLF_MTMARCHE'));
    TOBD.putValue('BLF_MTDEJAFACT',   TOBOO.GetValue('BLF_MTDEJAFACT'));
    TOBD.putValue('BLF_MTCUMULEFACT', TOBOO.GetValue('BLF_MTCUMULEFACT'));
    TOBD.putValue('BLF_MTSITUATION',  TOBOO.GetValue('BLF_MTSITUATION'));
    TOBD.putValue('BLF_QTEMARCHE',    TOBOO.GetValue('BLF_QTEMARCHE'));
    TOBD.putValue('BLF_QTEDEJAFACT',  TOBOO.GetValue('BLF_QTEDEJAFACT'));
    TOBD.putValue('BLF_QTECUMULEFACT',TOBOO.GetValue('BLF_QTECUMULEFACT'));
    TOBD.putValue('BLF_QTESITUATION', TOBOO.GetValue('BLF_QTESITUATION'));
    TOBD.putValue('BLF_POURCENTAVANC',TOBOO.GetValue('BLF_POURCENTAVANC'));
    TOBD.putValue('BLF_NATURETRAVAIL',TOBOO.GetValue('BLF_NATURETRAVAIL'));
    TOBD.putValue('BLF_FOURNISSEUR',  TOBOO.GetValue('BLF_FOURNISSEUR'));
    TOBD.putValue('BLF_TOTALTTCDEV',  TOBOO.GetValue('BLF_TOTALTTCDEV'));
    TOBD.putValue('BLF_MTPRODUCTION', 0);
    //
    TOBD.AddChampSupValeur('PIECEPRECEDENTE', TOBL.GetValue('GL_PIECEPRECEDENTE'));
    //
    TOBD.SetAllModifie (true);
  end;
end;

function GetTvaOuvrageDetail (TOBOUV : TOB; var Valeur : string) : string;
var Indice : integer;
		TOBO : TOB;
begin
	for Indice :=0  to TOBOUV.detail.count -1 do
  begin
    if TOBOUV.detail[Indice].detail.count > 0 then
    begin
			result := GetTvaOuvrageDetail(TOBOUV.detail[Indice],Valeur);
      if result = '***' then break;
    end else
    begin
      if valeur = '' then
      begin
        result := TOBOUV.detail[Indice].getString('BLO_FAMILLETAXE1');
        valeur := result;
      end else if result <> TOBOUV.detail[Indice].getString('BLO_FAMILLETAXE1') then
      begin
        result := '***';
        break;
      end;
    end;
  end;
end;

function GetTvaOuvrage (TOBl,TOBOuvrage : TOB) : string;
var TOBOUV : TOB;
    IndiceNomen : integer;
    valeur : string;
begin
  result := TOBL.GetString('GL_FAMILLETAXE1');
  if (TOBL.GetString('GL_TYPENOMENC')='OU1') then exit;
  IndiceNomen := TOBL.GetInteger('GL_INDICENOMEN');
  if IndiceNomen=0 then exit;
  valeur := '';
  TOBOUV := TOBOuvrage.Detail[IndiceNomen-1];
  Result := GetTvaOuvrageDetail (TOBOUV,Valeur);
end;

Procedure RecalculeLigneOuv(TOBOP, TOBL : TOB; DEV:RDevise; EnHt : boolean);
var valeurs : T_Valeurs;
    Qte     : double;
begin
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOP,1,1,true,DEV,valeurs, EnHT);
    GetValoDetail (TOBOP); // pour le cas des Article en prix posés
    Qte := TOBL.Getvalue('GL_QTEFACT');
    TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
    TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
    TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
    TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);
    TOBL.Putvalue('GL_MONTANTPA',Arrondi((Qte * TOBL.GetValue('GL_DPA')),V_PGI.okdecV));
    TOBL.Putvalue('GL_MONTANTPR',Arrondi((Qte * TOBL.GetValue('GL_DPR')),V_PGI.okdecV));

    TOBL.Putvalue('GL_DPA',valeurs[16]);
    TOBL.Putvalue('GL_DPR',valeurs[17]);
    //
    TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
    TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
    TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
    TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
    TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
    TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
    TOBL.Putvalue('GL_DPA',valeurs[0]);
    TOBL.Putvalue('GL_DPR',valeurs[1]);
    TOBL.Putvalue('GL_PMAP',valeurs[6]);
    TOBL.Putvalue('GL_PMRP',valeurs[7]);
    TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
    TOBL.PutValue('GL_RECALCULER', 'X');
    StockeInfoTypeLigne (TOBL,valeurs);
end;

procedure PositionneLigneOuvOrigine (TOBOuvrage : TOB);
var Indice : Integer;
begin
	for Indice := 0 to TOBOuvrage.detail.Count -1 do
  begin
  	TOBOuvrage.detail[Indice].SetString ('UTILISE','-') ;
  end;
end;

procedure RecupLesOuvrages(TOBPiece,TOBOuvrages : TOB);
begin

end;

procedure SetInfoPrecOuv (TOBOuvrage : TOB);
var II : Integer;
    TOBO : TOB;
begin
  for II := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBO := TOBOUvrage.detail[II];
    if TOBO.NomTable = 'LIGNEOUV' then
    begin
      if TOBO.Getvalue('BLO_PIECEORIGINE',) = '' then TOBO.putvalue('BLO_PIECEORIGINE',EncodeRefPieceOuv(TOBO));
      TOBO.putvalue('BLO_PIECEPRECEDENTE',EncodeRefPieceOuv(TOBO));
    end;
    if TOBO.Detail.count > 0 then
    begin
      SetInfoPrecOuv(TOBO);
    end;
  end;
end;

function FindLigneOuvFromUnique (TOBL,TOBOuvrages : TOB; UniqueBlo : Integer) : TOB;
var IndiceNomen : Integer;
    TOBOUV : TOB;
    II : Integer;
begin
  Result := nil;
  IndiceNomen := TOBL.GetInteger('GL_INDICENOMEN') -1;
  if IndiceNomen < TOBOuvrages.Detail.count then
  begin
    TOBOUV := TOBOuvrages.detail[IndiceNOmen];
    for II := 0 to TOBOUV.detail.Count -1 do
    begin
      if TOBOUV.detail[II].GetInteger('BLO_UNIQUEBLO')=UniqueBlo then
      begin
        Result := TOBOUV.detail[II];
        break;
      end;
    end;
  end;
end;

end.

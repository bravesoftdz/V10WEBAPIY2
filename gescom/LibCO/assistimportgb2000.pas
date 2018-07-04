{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001                                                                         
Modifié le ... :   /  /
Description .. : Reprise des données d'un dossier GB 2000
Mots clefs ... : REPRISE;GB 2000
*****************************************************************}
unit AssistImportGB2000;
// DEVISES - Reflexion à mener sur la reprise des taux de devises (4 types de taux dans GB 2000)
// ARTICLES - Articles de port : Ppr
// PARAMETRES - Motifs des entrees-sorties à gérer lors de la reprise des E/S
// PARAMETRES - Table des MODES DE REGLEMENT/PAIEMENT de GB 2000
// CLIENTS - par d'archivage des CLIRG ni des CLIAL
// CLIENTS - problème du multi-boutiques à revoir
// CLIENTS - reprise de l'état comptable            

//////////////////////////// Reprise du 25-05-2000
// CLIENTS - plus de T_COMMERCIAL/VENDEUR/REPRESENTANT                                                
// VENDEURS - intégration en commentaire (modif. de la table de l'époque)
// REPRESENTANT - intégration en commentaire (modif. de la table de l'époque)
// CAISSIER - intégration en commentaire (modif. de la table de l'époque)       f
//////////////////////////// Reprise du 28-06-2000
// VENDEURS REPRESENTANT CAISSIER : Alimentation de la table COMMERCIAL
interface
                                                                                                                                 
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ComCtrls, ExtCtrls,
  HPanel, UIutil, Spin, Hent1, ent1, HRichEdt, HRichOLE, UTob, Grids, DBTables,
  ImgList, HStatus, Mask, UTOM, RecupGBParamGen, RecupGBParamFrn,
  RecupGBParamCli, RecupGBParamArt, RecupGBDelete, RecupGBParamMvt , SaisieTableCorresp,
  FactUtil, FactCalc, Facture, FactComm, SaisUtil,
  UTomTiers, HDebug , UtilPGI, ParamSoc, EntGC, VoirTob, TarifUtil, SaisieCorrespGPAO,
  Menus, FactTob, FactPiece, FactAdresse;



type
  TFRecupGB2000 = class(TFAssist)
    Lancement: TTabSheet;
    TabSheet1: TTabSheet;
    HLabel8: THLabel;
    HLabel9: THLabel;
    Repfich: THCritMaskEdit;
    NumPremBtq: TEdit;
    Panel1: TPanel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    CompteRendu: TMemo;                                                                            
    bRecup: TToolbarButton97;
    BStop: TToolbarButton97;
    Article: TTabSheet;
    Label1: TLabel;
    HLabel46: THLabel;
    NumDerBtq: TEdit;
    Label3: TLabel;
    FichAno: THCritMaskEdit;
    OpenFichAno: TOpenDialog;
    GroupBox2: TGroupBox;
    RepParamGen: TCheckBox;
    RepParamBtq: TCheckBox;
    RepFournisseurs: TCheckBox;
    RepClients: TCheckBox;
    RepArticles: TCheckBox;
    GroupBox3: TGroupBox;
    RepCommandes: TCheckBox;
    RepAnnonces: TCheckBox;
    RepReceptions: TCheckBox;
    RepFactures: TCheckBox;
    RepTransferts: TCheckBox;
    RepESStock: TCheckBox;
    RepOD: TCheckBox;
    RepDetaxes: TCheckBox;
    RepMatricules: TCheckBox;
    RepStatut: TCheckBox;
    RepStock: TCheckBox;
    RepReglements: TCheckBox;
    RepVentes: TCheckBox;
    ButtonParamGen: TButton;
    ButtonParamFrn: TButton;
    ButtonParamClients: TButton;
    ButtonParamArticles: TButton;
    GroupBox4: TGroupBox;
    DeleteTable: TCheckBox;
    ButtonParamDelete: TButton;
    ButtonParamMvt: TButton;
    Label2: TLabel;
    FichTrt: THCritMaskEdit;
    OpenFichTrt: TOpenDialog;
    ButtonCorresp: TButton;
    RepNegC: TCheckBox;
    RepNEGF: TCheckBox;
    PopupCorresp: TPopupMenu;
    ModeRegle: TMenuItem;
    Pays1: TMenuItem;
    Devise1: TMenuItem;
    Famille1: TMenuItem;
    ModePaie: TMenuItem;
    TVA1: TMenuItem;
    procedure bFinClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure NumPremBtqExit(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bRecupClick(Sender: TObject);
    procedure bStopButton(Sender: TObject);
    procedure NumDerBtqExit(Sender: TObject);
    procedure RepArticlesClick(Sender: TObject);
    procedure RepClientsClick(Sender: TObject);
    procedure RepFournisseursClick(Sender: TObject);
    procedure FichAnoElipsisClick(Sender: TObject);
    procedure RepParamGenClick(Sender: TObject);
    procedure DeleteTableClick(Sender: TObject);
    procedure ButtonParamGenClick(Sender: TObject);
    procedure ButtonParamFrnClick(Sender: TObject);
    procedure ButtonParamClientsClick(Sender: TObject);
    procedure ButtonParamArticleClick(Sender: TObject);
    procedure ButtonParamDeleteClick(Sender: TObject);
    procedure RepCommandesClick(Sender: TObject);
    procedure ButtonParamMvtClick(Sender: TObject);
    procedure FichTrtElipsisClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bVerifLesPieces(Sender: TObject);
    procedure ModeRegleClick(Sender: TObject);
    procedure Pays1Click(Sender: TObject);
    procedure Devise1Click(Sender: TObject);
    procedure Famille1Click(Sender: TObject);
    procedure ModePaieClick(Sender: TObject);
    procedure ButtonCorrespMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure TVA1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private   { Déclarations privées }
    TOB_Tailles     	: TOB ;
    Tob_Article     	: TOB ;
    Tob_ArticleCompl	: TOB ;
    Tob_Tiers       	: TOB ;
    Tob_Tiers_Defaut	: TOB ;
    Tob_Tiers_Compl 	: TOB ;
    Tob_Adresses    	: TOB ;
    Tob_Mere_Dispo   	: TOB ;
    Tob_Dispo		: TOB ;
    Tob_Mere_Tarif	: TOB ;
    Tob_Piece       	: TOB ;
    Tob_Ligne       	: TOB ;
    Tob_Piedbase	: TOB ;
    Tob_PiedEche	: TOB ;
    Tob_PiedPort        : TOB ;
    Tob_Acomptes    	: TOB ;
    Tob_TarifMode_TTC	: TOB ;
    Tob_TarifMode_Sol   : TOB ;
    Tob_TarifMode_Ach	: TOB ;
    Tob_DevisePGI       : TOB ;
    ArretRecup      	: Boolean ;
    FichierAno      	: textfile;
    NbrAno          	: Integer ;
    FichierTrt      	: textfile;
    Compteur	    	: integer ;
    CompteurParam   	: integer ;
    CompteurClient  	: integer ;
    CompteurArticle   	: integer ;
    CompteurCodeBarre	: integer ;
    //
    // Intégration des tables de paramètres
    //
    DerTableLibreArg  : String;
    DerTableLibreAff  : Integer;
    DerTableLibreArg2 : String;
    DerTypeParam      : string ;
    DerTypeParam2     : string ;
    //
    // Intégration des articles
    //
    OptimisationArticle : string ;
    RecodifieFamille  : string ;
    FNiv1             : string ;
    FNiv2             : string ;
    FNiv3             : string ;
    Famille           : string ;
    Modele            : string ;
    Matiere           : string ;
    Coloris           : string ;
    Finition          : string ;
    Compo1            : string ;
    Compo2            : string ;
    CatDim            : string ;
    TVA               : string ;
    Categorie         : string ;
    Formule           : string ;
    Structure         : string ;
    TypeArt           : string ;
    Libcol            : string ;

    TauxTVADefaut     : double ;
    CodeArticle       : string ;     // Code Article traité dans la reprise des mouvements
    Designation       : string ;     // Désignation de l'article
    DeviseSiteCentral : string ;
    //
    // Intégration des documents
    //
    Numligne          : integer ;    // Numéro de ligne pour la reprise des documents
    CleDoc            : R_CleDoc ;
    DEV               : RDEVISE ;
    NbErreurs  	      : integer ;    // Nombre d'erreur rencontrée lors de l'intégration
    QualifMvt	      : string ;     // qualifiant mouvement de la nature de pièce
    CategorieDimension: string  ;
    // Gestion de la ligne commentaire dans document
    LigneCommACreer   : boolean;
    NumLigneComment   : integer;     // Numéro de la ligne commentaire
    TotQte	          : double ;     // Total quantité
    TotQteReste       : double ;     // Total quantité
    TotPrix	          : double ;     // Total Prix
    // intégration des lignes de ventes
    NumEtabVte	      : string ;    // Code Etablissement en cours
    NumPieceVte	      : string ;    // Numéro de ticket en cours
    DatePieceVte      : string ;    // Date de la pièce en cours
    QuantiteVendueGB  : double ;    // Qté de la vente en cours GB (Valeur Absolue)
    QuantiteVenduePGI : double ;    // Qté de la vente en cours PGI (Valeur signée)
    CoutVentePAMP     : double ;    // Qté de la vente en cours GB (Valeur Absolue)
    // Intégration des règlements
    NatureVente       : string ;
    SoucheVente       : string ;
    NumeroVente       : integer;
    IndiceVente       : integer;
    // Intégration des commandes
    CommandePartiel   : boolean;
    Commandesoldee    : boolean;
    // Intégration des réceptions
    Tob_Commande      : TOB    ;
    ReceptSurCommande : boolean;
    RetourFrn	      : Boolean;
    BtqCommandeGB2000 : string ;
    NumCommandeGB2000 : string ;
    RefCommandeGB2000 : string ;
    // intégration des E/S de stock
    EntreeStock	      : boolean;
    NatureEntree      : string ;
    NatureSortie      : string ;
    Motif_ES          : string ;
    // intégration des stocks et tarifs
    CompteurDispo     : integer;
    QteVendue 	      : double ;
    QteRecue	      : double ;
    MaxTarif	      : integer;
    DeviseTarif	      : string ;
    CodeArtTarif      : string ;
    DernierArticlePA  : string ;
    DernierArticlePV  : string ;
    DernierArticlePR  : string ;
    DernierArticlePVHT: string ;
    // Intégration des factures NEGOCE
    FactureSurCommande  : boolean ;
    FactureSurLivraison : boolean ;
    AvoirNegoce         : boolean ;
    PieceAGenerer       : boolean ;
    Tob_AdrFact         : TOB     ;
    BtqNeg2000          : string ;
    NumNeg2000          : string ;
    NumligPort          : integer;
    EtatComptable       : string ;
    PrixUnique          : boolean;
    PrixVenteNeg        : double;
    LigneAReprendre     : boolean;
    Procedure RecupGrille (st : string);
    Procedure RecupCivilite (st : string);
    Procedure RecupArticle1 (st : string);
    Procedure RecupArticle2 (st : string);
    Procedure RecupArticle3 (st : string);
    Procedure RecupArticleAB (st : string);
    Procedure RecupArtTaille (st : string);
    procedure ImportArticle ;
    procedure ImportMereArtTaille;
    procedure RecupParametres ;
    Procedure RecupFamille (st : string);
    procedure RecupParamBtq ;
    Procedure RecupFournisseurs ;
    procedure RecupArticles ;
    Procedure RecupClients ;
    Procedure RecupCommandes ;
    Procedure RecupCommande (st : String);
    Procedure RecupAnnonces ;
    Procedure RecupReceptions ;
    Procedure RecupReception (st:string) ;
    Procedure RecupFactures;
    Procedure RecupTransferts ;
    Procedure RecupTransfert (st:string) ;
    Procedure RecupESStocks ;
    Procedure RecupESStock (st:string);
    Procedure RecupOD ;
    Procedure RecupVentes;
    Procedure RecupVente (st:string) ;
    Procedure RecupReglements ;
    Procedure RecupReglement (st:string) ;
    Procedure RecupNegFactures ;
    Procedure RecupNegFacture (st:string) ;
    Procedure RecupNegCommandes;
    Procedure RecupNegCommande  (st : string);
    Procedure RecupStock ;
    Procedure RecupStatuts;
    Procedure RecupMatricules;
    Procedure RecupDetaxes;
    Procedure RecupCLI (st : string);
    Procedure RecupFOU (st : String);
    procedure ImportTiers (VerifTom : boolean) ;
    procedure ImportMereTiers ;
    procedure ImportTiersCompl;
    procedure ImportMereTiersCompl;
    Procedure RecupSTD (st : String);
    procedure ImportDispo ;
    procedure ImportTarif ;
    procedure RecupMaxTarif;
    procedure ChargeTobQualifiantTarif (var Tob_Tarif : TOB; TypeTarif : string; TypePeriode : string);
    Procedure VerifPiece ;
    Procedure RecupVenteDiverse (st : string);
    procedure MajPrixAchatArticle   (CodeArticle : string ; PrixArticle : double; Devise : string) ;
    procedure MajPrixVenteArticle   (CodeArticle : string ; PrixArticle : double; Devise : string) ;
    procedure MajPrixVenteHTArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
    procedure MajPrixRevientArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;

    procedure CreeLigneCommentaire;
    procedure ImportPiece ;
    procedure ImportLivraison ;
    procedure ImportTransfert ;
    procedure ImportESStock ;
    procedure ImportFacCdeLiv ;
    procedure ImportReglement ;
    procedure ImportAdresseTiers;
    procedure SupprimeTables ;
    procedure DeleteTablesLibres(typ : string; num : Integer);
    procedure RecupTablesLibresArt(st : string);
    procedure RecupTablesLibresStatArt(st : string);
    procedure RecupTablesLibresCli(st : string);
    procedure RecupTablesLibresFou(st : string);
    procedure RecupRegroupeFamille (st : string);
    procedure CtrlTableTypeArt ;
    procedure CtrlTableEtatCpta ;
    procedure EcrireFichAno (Chtmp : String) ;
    procedure EcrireFichTrt (Chtmp : String) ;
    procedure EcrireInfo (Chtmp : String; Cpte, Trt, Ano : boolean);
    procedure ChargeTobTaille;
    procedure RecupTauxTVADefaut;
    Procedure ChargeParamArticle ;
    Procedure ChargeDeviseCentral ;
    Procedure ChargeDevisePGI ;
    function  NombreErreurs (nberr : integer) : Integer;
    procedure SavTobPaiement;
    procedure SavTobPays;
    procedure SavTobDevise;
    procedure SavTobFamille;
    function  ChargeLaCommande (NaturePiece, RefInterne : string) : boolean ;
public    { Déclarations publiques }
  end;

/////////////////////////////////////////////////////////////////////////////
// TOB MERE pour récupération des paramètres, clients, articles tailles, ...
/////////////////////////////////////////////////////////////////////////////
var	Tob_Parametre         : TOB ;
   	Tob_Mere_Tiers        : TOB ;
        Tob_Mere_Tiers_Compl  : TOB ;
        Tob_Mere_Art_Taille   : TOB ;

        Tob_Param      	      : TOB ;  // TOB utilisée pour stocker le paramétrage de la récup
        Tob_Reglement         : TOB ;  // TOB de correspondance des modes de règlements
        Tob_ModePaie          : TOB ;  // TOB de correspondance des modes de paiements
        Tob_Pays              : TOB ;  // TOB de correspondance des pays
        Tob_Devise            : TOB ;  // TOB des correspondance des devises
        Tob_Famille           : TOB ;  // TOB des correspondances des pays
        Tob_TVA               : TOB ;  // TOB des correspondances des TVA
        CodeInterne           : string ;
        DateCloture           : TDateTime ;

Type R_IdentCol = RECORD
                  ColName, ColVal : String ;
                  END ;

const
 ////////////////////////////////////////////////////////////////////
 // Noms des champs de la TOB Tob_param, avec leur valeur par défaut
 /////////////////////////////////////////////////////////////////////
 ChampTobParam : array[1..105,1..2] of String 	= (
   ('OPTIMISATIONARTICLE', '0'),
   ('RECODIFIEFAMILLE'  , '0'),
   ('FAMNIV1'           , '3'),
   ('FAMNIV2'           , '3'),
   ('FAMNIV3'           , '0'),
   ('REPFAMILLE'        , ''),
   ('REPMOD'            , ''),
   ('REPMAT'            , ''),
   ('REPCOL'            , ''),
   ('REPFIN'            , ''),
   ('REPSTR'            , ''),
   ('REPLIBCOL'         , ''),
   ('REPCAT'            , ''),
   ('REPFOR'            , ''),
   ('REPCP1'            , ''),
   ('REPCP2'            , ''),
   ('REPTYP'            , ''),
   ('CATEGORIEDIM'      , ''),
   ('GRILLEUNI'         , 'UNI'),
   ('REPCODEBARRE'      , ''),
   ('REPTVA'            , ''),
   ('REPREG'            , ''),
   ('CODECLI'           , '0'),
   ('REPCLIPROMAIL'     , ''),
   ('CLIEXPORT'         , ''),
   ('REPETACPTA'        , ''),
   ('FRNFAMCOMPTA'      , ''),
   ('FRNREGIMETVA'      , ''),
   ('NATCDE'            , ''),
   ('NATREC'            , ''),
   ('NATENT'            , ''),
   ('NATSOR'            , ''),
   ('NATTRFEM'          , ''),
   ('NATTRFRE'          , ''),
   ('NATNEGCDE'         , ''),
   ('NATNEGLIV'         , ''),
   ('NATNEGFAC'         , ''),
   ('NATNEGAVO'         , ''),
   ('NATVTE'            , ''),
   ('MODEREGLVTE'       , ''),
   ('CODECLIVTE'        , ''),
   ('VALSTOCK'          , '0'),
   ('TARIF'             , '0'),
   ('TARIFHT'           , '0'),
   ('TYPETARIFTTC'      , '' ),
   ('PERIODETARIFTTC'   , '' ),
   ('TYPETARIFSOLDE'    , '' ),
   ('PERIODETARIFSOLDE' , '' ),
   ('TARIFACH'          , '0'),
   ('TYPETARIFACH'      , '' ),
   ('PERIODETARIFACH'   , '' ),
   ('DELBOUTIQUES'      , ''),
   ('DELDEVISES'        , '0'),
   ('DELARRONDIS'       , '0'),
   ('DELPAYS'           , '0'),
   ('DELCODESPOSTAUX'   , '0'),
   ('DELFAMILLES'       , '0'),
   ('DELCOLLECTIONS'    , '0'),
   ('DELPOIDS'          , '0'),
   ('DELDEMARQUES'      , '0'),
   ('DELCIVILITES'      , '0'),
   ('DELREPRESENTANTS'  , '0'),
   ('DELCATEGORIES'     , '0'),
   ('DELGRILLEDIM1'     , '0'),
   ('DELDIM1'           , '0'),
   ('DELGRILLEDIM2'     , '0'),
   ('DELDIM2'           , '0'),
   ('DELGRILLEDIM3'     , '0'),
   ('DELDIM3'           , '0'),
   ('DELGRILLEDIM4'     , '0'),
   ('DELDIM4'           , '0'),
   ('DELGRILLEDIM5'     , '0'),
   ('DELDIM5'           , '0'),
   ('DELMASQUEDIM'      , '0'),
   ('DELTABLEA1'        , '0'),
   ('DELTABLEA2'        , '0'),
   ('DELTABLEA3'        , '0'),
   ('DELTABLEA4'        , '0'),
   ('DELTABLEA5'        , '0'),
   ('DELTABLEA6'        , '0'),
   ('DELTABLEA7'        , '0'),
   ('DELTABLEA8'        , '0'),
   ('DELTABLEA9'        , '0'),
   ('DELTABLEA10'       , '0'),
   ('DELTABLEC1'        , '0'),
   ('DELTABLEC2'        , '0'),
   ('DELTABLEC3'        , '0'),
   ('DELTABLEC4'        , '0'),
   ('DELTABLEC5'        , '0'),
   ('DELTABLEC6'        , '0'),
   ('DELTABLEC7'        , '0'),
   ('DELTABLEC8'        , '0'),
   ('DELTABLEC9'        , '0'),
   ('DELTABLEC10'       , '0'),
   ('DELARTICLES'       , '0'),
   ('DELCLIENTS'        , '0'),
   ('DELFOURNISSEURS'   , '0'),
   ('STOCKCLOTURE'      , 'N'),
   ('DTECLOTURE'        , '01/01/1900'),
   ('REPPAARTICLE'      , 'N'),
   ('REPPVARTICLE'      , 'N'),
   ('REPPVHTARTICLE'    , 'N'),
   ('REPPRARTICLE'      , 'N'),
   ('REPDEVISE'         , '' ),
   ('CODEARTVTE'        , '' ));

Procedure Assist_ImportGB2000 ;
Procedure ArticleCode (Tob_Art : TOB) ;      // Reconstitue GA_ARTICLE à partir des autres zones

//
// Chargement en TOB des tables de correspondances
//
procedure DefineChampTobParam (TobParametrage : TOB) ;
procedure ChargeTobParam (TobParametrage : TOB) ;

implementation
{$R *.DFM}

uses GenereBE;

////////////////////////////////////////////////////////////
// Sauvegarde la tob Paiement
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.SavTobPaiement;
begin
   // Sauvegarde de la TOB
  Tob_ModePaie.InsertOrUpdateDB(True) ;
end;

///////////////////////////////////////////////////////////
// Sauvegarde la tob Pays
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.SavTobPays;
begin
  // Sauvegarde de la TOB
  Tob_Pays.InsertOrUpdateDB(True) ;
end;

///////////////////////////////////////////////////////////
// Sauvegarde la tob Devise
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.SavTobDevise;
begin
  // Sauvegarde de la TOB
  Tob_Devise.InsertOrUpdateDB(True) ;
end;

///////////////////////////////////////////////////////////
// Sauvegarde la tob Famille
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.SavTobFamille;
begin
  // Sauvegarde de la TOB
  Tob_Famille.InsertOrUpdateDB(True) ;
end;


///////////////////////////////////////////////////////////////////////////////////////////////
// Retourne le code interne PGI en fonction du rang GB 2000
///////////////////////////////////////////////////////////////////////////////////////////////
function CalculTaillePGI ( Taille : integer) : string ;
begin
  case Taille of
        1  : result := '001';
        2  : result := '002';
        3  : result := '003';
        4  : result := '004';
        5  : result := '005';
        6  : result := '006';
        7  : result := '007';
        8  : result := '008';
        9  : result := '009';
        10 : result := '00A';
        11 : result := '00B';
        12 : result := '00C';
        13 : result := '00D';
        14 : result := '00E';
        15 : result := '00F';
        16 : result := '00G';
        17 : result := '00H';
        18 : result := '00I';
        19 : result := '00J';
        20 : result := '00K';
        21 : result := '00L';
        22 : result := '00M';
        23 : result := '00N';
        24 : result := '00O';
        25 : result := '00P';
        26 : result := '00Q';
        27 : result := '00R';
        28 : result := '00S';
        29 : result := '00T';
        30 : result := '00U';
        else result := '???';
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////////////
// Détermine la référence de l'article dimensionné ou non : GA_ARTICLE
///////////////////////////////////////////////////////////////////////////////////////////////
procedure ArticleCode (Tob_Art : TOB) ;
Var Chtmp     : String ;
    StBlanc   : string ;
    Article   : string ;
    longueur  : integer;
begin
// MODIF LM
Article := Tob_art.GetValue ('GA_CODEARTICLE');
Article := Trim   (Article);
longueur:= length (Article);
StBlanc := '                  ';

Chtmp:=Article+Copy(StBlanc, 1, 18-longueur)+     // MODIF LM : complete par des blancs à 18 caractères
               Tob_Art.GetValue('GA_CODEDIM1')+
               Tob_Art.GetValue('GA_CODEDIM2')+
               Tob_Art.GetValue('GA_CODEDIM3')+
               Tob_Art.GetValue('GA_CODEDIM4')+
               Tob_Art.GetValue('GA_CODEDIM5')+
               'X';    // MODIF LM : Ajout X pour fin référence

Tob_Art.PutValue('GA_ARTICLE',Chtmp);
end ;

////////////////////////////////////////////////////////////
// Importation des tables de paramètres
////////////////////////////////////////////////////////////
procedure ImportParametre;
begin
    Tob_Parametre.SetAllModifie(TRUE) ;
    Tob_Parametre.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
    Tob_Parametre.free;
    Tob_Parametre:=nil;
    committrans;
end ;

////////////////////////////////////////////////////////////
// Pco - REPRISE DES TABLES LIBRES
////////////////////////////////////////////////////////////
procedure ImportTablesLibres ( T : TOB );       // Mettre des controles de cohérence
begin
    T.SetAllModifie(TRUE) ;
    T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

////////////////////////////////////////////////////////////
// Gère le nombre d'erreur
////////////////////////////////////////////////////////////
function TFRecupGB2000.NombreErreurs (nberr : integer) : Integer;
begin

	if (nberr = 0) then NbErreurs := 0
  else if (nberr > 0) then NbErreurs:=NbErreurs+1;

  result:=NbErreurs;
end;

////////////////////////////////////////////////////////////
// Intégration des tables libres articles
////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTablesLibresArt (st : string);
var TPST: TOB;
    Chtmp,Chtmp2 : string;
    cpt : integer;
begin
  if DerTableLibreArg<>copy(st,1,3) then           // recherche de la table concernée
  begin
    Chtmp         := '' ;
    Chtmp2        := '';
    DerTypeParam  := '' ;
    DerTypeParam2 := '' ;

    if      Copy(st,1,3)='Pco' then Chtmp:=Tob_Param.GetValue ('REPCOL')
    else if Copy(st,1,3)='Pfs' then Chtmp:=Tob_Param.GetValue ('REPFAMILLE')
    else if Copy(st,1,3)='Pmo' then Chtmp:=Tob_Param.GetValue ('REPMOD')
    else if Copy(st,1,3)='Pma' then Chtmp:=Tob_Param.GetValue ('REPMAT')
    else if Copy(st,1,3)='Pfi' then Chtmp:=Tob_Param.GetValue ('REPFIN')
    else if Copy(st,1,3)='Pcg' then Chtmp:=Tob_Param.GetValue ('REPCAT')
    else if Copy(st,1,3)='Pfo' then Chtmp:=Tob_Param.GetValue ('REPFOR')
    else if Copy(st,1,3)='Pcp' then
    begin
      if Tob_Param.GetValue ('REPCP1')=''    then Chtmp:=Tob_Param.GetValue ('REPCP2')
      else begin
        Chtmp:=Tob_Param.GetValue ('REPCP1');
        if Tob_Param.GetValue ('REPCP2')<>'' then Chtmp2:=Tob_Param.GetValue ('REPCP2');
      end;
    end;

    if Chtmp<>''then
    begin
      if copy (Chtmp, 1, 3) = 'GA_' then
      begin
        // Table libre 1 à 10
        for cpt:=1 to 10 do
        begin
          if (cpt=10) and (Chtmp='GA_LIBREARTA') then break
          else if Chtmp='GA_LIBREART'+InttoStr(cpt) then break ;
        end ;

        if cpt<11 then
        begin
          DerTypeParam    := 'LA' + IntToStr(cpt);
          DerTableLibreArg:=copy(st,1,3);
        end ;
      end else
      begin
        if copy (Chtmp, 1, 3) = 'GA2' then
        begin
          // Famille de niveau 4 à 8
          if copy (Chtmp, 5, 3) = 'FAM' then
          begin
            for cpt:=4 to 8 do if Chtmp='GA2_FAMILLENIV'+InttoStr(cpt) then break ;
            if cpt < 9 then
            begin
              DerTypeParam    := 'FN' + IntToStr(cpt);
              DerTableLibreArg:=copy(st,1,3);
            end;
          end else
          begin
            // statistique article 1 à 2
            if Chtmp='GA2_STATART1'       then DerTypeParam := 'FNA'
            else if Chtmp='GA2_STATART2'  then DerTypeParam := 'FNB';
             DerTableLibreArg := copy(st,1,3);
          end;
        end;
      end;
    end ;

    if Chtmp2<>'' then
    begin
      if copy (Chtmp2, 1, 3) = 'GA_' then
      begin
        for cpt:=1 to 10 do
        begin
          if (cpt=10) and (Chtmp2='GA_LIBREARTA') then break
          else if Chtmp2='GA_LIBREART'+InttoStr(cpt) then break ;
        end ;
        if cpt<11 then
        begin
          DerTypeParam2     := 'LA' + IntToStr(cpt);
          DerTableLibreArg2 :=copy(st,1,3);
        end ;
      end else
      begin
        if copy (Chtmp2, 1, 3) = 'GA2' then
        begin
          // Famille de niveau 4 à 8
          if copy (Chtmp2, 5, 3) = 'FAM' then
          begin
            for cpt:=4 to 8 do if Chtmp2='GA2_FAMILLENIV'+InttoStr(cpt) then break ;
            if  cpt < 9 then
            begin
              DerTypeParam2     := 'FN' + IntToStr(cpt);
              DerTableLibreArg2 := copy(st,1,3);
            end;
          end else
          begin
            // statistique article 1 à 2
            if Chtmp2='GA2_STATART1'      then DerTypeParam2 := 'FNA'
            else if Chtmp2='GA2_STATART1' then DerTypeParam2 := 'FNB';
            DerTableLibreArg2 := copy(st,1,3);
          end;
        end;
      end;
    end;
  end;

  if DerTypeParam <> '' then                      // si affectation trouvee
  begin
    // Famille niv 4 à 8
    if (copy (DerTypeParam, 1, 3) >= 'FN4') and (copy (DerTypeParam, 1, 3) <= 'FN8') then
    begin
      TPST := TOB.CREATE ('CHOIXCOD', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('CC_TYPE', DerTypeParam) ;
      if      Copy (st, 1, 3) = 'Pco' then TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfs' then TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfi' then TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pcp' then TPST.PutFixedStValue('CC_CODE', st, 9, 2, tctrim, true)
      else if Copy (st, 1, 3) = 'Pcg' then TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfo' then TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true)
      else TPST.PutFixedStValue('CC_CODE', st, 9, 3, tctrim, true);
      if Copy (st, 1, 3) = 'Pfs' then
      begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 27, 25, tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE' , st, 27, 17, tcChaine, False);
      end else
      begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE', st, 15, 17, tcChaine, False);
      end;
      ImportTablesLibres (TPST);
      TPST.free ;
    end else
    begin
      TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('YX_TYPE', DerTypeParam) ;

      if Copy (st, 1, 3) = 'Pco' then TPST.PutFixedStValue('YX_CODE', st, 9, 5, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfi' then TPST.PutFixedStValue('YX_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfs' then TPST.PutFixedStValue('YX_CODE', st, 9, 6, tctrim, true)
      else if Copy (st, 1, 3) = 'Pcp' then TPST.PutFixedStValue('YX_CODE', st, 9, 2, tctrim, true)
      else if Copy (st, 1, 3) = 'Pcg' then TPST.PutFixedStValue('YX_CODE', st, 9, 3, tctrim, true)
      else if Copy (st, 1, 3) = 'Pfo' then TPST.PutFixedStValue('YX_CODE', st, 9, 3, tctrim, true)
      else TPST.PutFixedStValue('YX_CODE', st, 9, 6, tctrim, true);
      if Copy (st, 1, 3) = 'Pfs' then
      begin
        TPST.PutFixedStValue('YX_LIBELLE', st, 27, 25, tcChaine, False);
        TPST.PutFixedStValue('YX_ABREGE' , st, 27, 17, tcChaine, False);
      end else
      begin
        TPST.PutFixedStValue('YX_LIBELLE', st, 15, 25, tcChaine, False);
        TPST.PutFixedStValue('YX_ABREGE' , st, 15, 17, tcChaine, False);
      end;

      ImportTablesLibres (TPST);
      TPST.free ;
    end;
  end ;

  if DerTypeParam2 <> '' then                      // si affectation trouvee
  begin
    // Famille niv 4 à 8
     if (copy (DerTypeParam, 1, 3) >= 'FN4') and (copy (DerTypeParam, 1, 3) <= 'FN8') then
    begin
      TPST := TOB.CREATE ('CHOIXCOD', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('CC_TYPE', DerTypeParam2) ;
      TPST.PutFixedStValue('CC_CODE'   , st,  9,  2, tctrim, true)  ;
      TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcChaine, False);
      TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcChaine, False);
      ImportTablesLibres (TPST);
      TPST.free ;
    end else
    begin
      TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('YX_TYPE', DerTypeParam2) ;
      TPST.PutFixedStValue('YX_CODE'   , st,  9,  2, tctrim, true);
      TPST.PutFixedStValue('YX_LIBELLE', st, 15, 25, tcChaine, False);
      TPST.PutFixedStValue('YX_ABREGE' , st, 15, 17, tcChaine, False);
      ImportTablesLibres (TPST);
      TPST.free ;
    end;
  end ;
end;

////////////////////////////////////////////////////////////
// Ecriture d'un message dans le fichier des anomalies
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.EcrireFichAno (Chtmp : String) ;
begin
Writeln(FichierAno,Chtmp);
NbrAno:=NbrAno+1;
end;

////////////////////////////////////////////////////////////
// Ecriture d'un message dans le fichier de compte rendu
////////////////////////////////////////////////////////////
procedure TFRecupGB2000.EcrireFichTrt (Chtmp : String) ;
begin
Writeln(FichierTrt,Chtmp);
end;

//////////////////////////////////////////////////////////////////////////////////////////////
// Gestion des messages d'intégration : Affichage Ecran, Fichier Compte rendu, Fichier anomalies
//////////////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.EcrireInfo (Chtmp : String; Cpte, Trt, Ano : boolean);
begin
	if (Cpte = True) then CompteRendu.lines.add(Chtmp);
  if (Trt  = True) then EcrireFichTrt (Chtmp);
  if (Ano  = True) then EcrireFichAno (Chtmp);
end;

////////////////////////////////////////////////////////////
// Intégration des statistiques articles
////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTablesLibresStatArt (st : string);
var TPST: TOB;
    Chtmp : string;
    cpt : integer;
begin
  if DerTableLibreArg<>copy(st,1,3) then           // recherche de la table concernée
  begin
    Chtmp:='' ;
    DerTableLibreAff:=0;
    if Copy(st,1,3)='P2A' then Chtmp:=Tob_Param.GetValue ('REPSTATA01')
    else if Copy(st,1,3)='P2B' then Chtmp:=Tob_Param.GetValue ('REPSTATA02')
    else if Copy(st,1,3)='P2C' then Chtmp:=Tob_Param.GetValue ('REPSTATA03')
    else if Copy(st,1,3)='P2D' then Chtmp:=Tob_Param.GetValue ('REPSTATA04')
    else if Copy(st,1,3)='P2E' then Chtmp:=Tob_Param.GetValue ('REPSTATA05')
    else if Copy(st,1,3)='P2F' then Chtmp:=Tob_Param.GetValue ('REPSTATA06')
    else if Copy(st,1,3)='P2G' then Chtmp:=Tob_Param.GetValue ('REPSTATA07')
    else if Copy(st,1,3)='P2H' then Chtmp:=Tob_Param.GetValue ('REPSTATA08')
    else if Copy(st,1,3)='P2I' then Chtmp:=Tob_Param.GetValue ('REPSTATA09')
    else if Copy(st,1,3)='P2J' then Chtmp:=Tob_Param.GetValue ('REPSTATA10')
    else if Copy(st,1,3)='P2K' then Chtmp:=Tob_Param.GetValue ('REPSTATA11')
    else if Copy(st,1,3)='P2L' then Chtmp:=Tob_Param.GetValue ('REPSTATA12')
    else if Copy(st,1,3)='P2M' then Chtmp:=Tob_Param.GetValue ('REPSTATA13')
    else if Copy(st,1,3)='P2N' then Chtmp:=Tob_Param.GetValue ('REPSTATA14')
    else if Copy(st,1,3)='P2O' then Chtmp:=Tob_Param.GetValue ('REPSTATA15')
    else if Copy(st,1,3)='P2P' then Chtmp:=Tob_Param.GetValue ('REPSTATA16')
    else if Copy(st,1,3)='P2Q' then Chtmp:=Tob_Param.GetValue ('REPSTATA17')
    else if Copy(st,1,3)='P2R' then Chtmp:=Tob_Param.GetValue ('REPSTATA18')
    else if Copy(st,1,3)='P2S' then Chtmp:=Tob_Param.GetValue ('REPSTATA19')
    else if Copy(st,1,3)='P2T' then Chtmp:=Tob_Param.GetValue ('REPSTATA20');

    if Chtmp<>''then
    begin
      if copy (Chtmp, 1, 3) = 'GA_' then
      begin
        // Table libre 1 à 10
        for cpt:=1 to 10 do
        begin
          if (cpt=10) and (Chtmp='GA_LIBREARTA') then break
          else if Chtmp='GA_LIBREART'+InttoStr(cpt) then break ;
        end ;

        if cpt<11 then
        begin
          if cpt<10 then DerTypeParam    := 'LA' + IntToStr(cpt)
          else if cpt=10 then DerTypeParam    := 'LAA';
          DerTableLibreArg:=copy(st,1,3);
        end ;
      end else
      begin
        if copy (Chtmp, 1, 3) = 'GA2' then
        begin
          // Famille de niveau 4 à 8
          if copy (Chtmp, 5, 3) = 'FAM' then
          begin
            for cpt:=4 to 8 do if Chtmp='GA2_FAMILLENIV'+InttoStr(cpt) then break ;
            if cpt < 9 then
            begin
              DerTypeParam    := 'FN' + IntToStr(cpt);
              DerTableLibreArg:=copy(st,1,3);
            end;
          end else
          begin
            // statistique article 1 à 2
            if Chtmp='GA2_STATART1'       then DerTypeParam := 'FNA'
            else if Chtmp='GA2_STATART2'  then DerTypeParam := 'FNB';
             DerTableLibreArg := copy(st,1,3);
          end;
        end;
      end;
    end ;
  end;

  if DerTypeParam <> '' then                      // si affectation trouvee
  begin
    // Famille niv 4 à 8
    if (copy (DerTypeParam, 1, 3) >= 'FN4') and (copy (DerTypeParam, 1, 3) <= 'FN8') then
    begin
      TPST := TOB.CREATE ('CHOIXCOD', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('CC_TYPE', DerTypeParam) ;
      TPST.PutFixedStValue('CC_CODE'    , st, 9, 3, tctrim, true);
      TPST.PutFixedStValue('CC_LIBELLE' , st, 15, 25, tcChaine, False);
      TPST.PutFixedStValue('CC_ABREGE'  , st, 15, 17, tcChaine, False);
      ImportTablesLibres (TPST);
      TPST.Free;
    end else
    begin
      TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue('YX_TYPE', DerTypeParam) ;

      TPST.PutFixedStValue('YX_CODE'   , st, 9, 6, tctrim, true);
      TPST.PutFixedStValue('YX_LIBELLE', st, 15, 25, tcChaine, False);
      TPST.PutFixedStValue('YX_ABREGE' , st, 15, 17, tcChaine, False);
      ImportTablesLibres (TPST);
      TPST.Free;
    end;
  end ;
end;

////////////////////////////////////////////////////////////
// Intégration des statistiques clients
////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTablesLibresCli (st : string);
var TPST: TOB;
    Chtmp : string;
    cpt : integer;
begin
  if DerTableLibreArg<>copy(st,1,3) then           // recherche de la table concernée
  begin
    Chtmp:='' ;
    DerTableLibreAff:=0;
    if Copy(st,1,3)='P1A' then Chtmp:=Tob_Param.GetValue ('REPSTATC01')
    else if Copy(st,1,3)='P1B' then Chtmp:=Tob_Param.GetValue ('REPSTATC02')
    else if Copy(st,1,3)='P1C' then Chtmp:=Tob_Param.GetValue ('REPSTATC03')
    else if Copy(st,1,3)='P1D' then Chtmp:=Tob_Param.GetValue ('REPSTATC04')
    else if Copy(st,1,3)='P1E' then Chtmp:=Tob_Param.GetValue ('REPSTATC05')
    else if Copy(st,1,3)='P1F' then Chtmp:=Tob_Param.GetValue ('REPSTATC06')
    else if Copy(st,1,3)='P1G' then Chtmp:=Tob_Param.GetValue ('REPSTATC07')
    else if Copy(st,1,3)='P1H' then Chtmp:=Tob_Param.GetValue ('REPSTATC08')
    else if Copy(st,1,3)='P1I' then Chtmp:=Tob_Param.GetValue ('REPSTATC09')
    else if Copy(st,1,3)='P1J' then Chtmp:=Tob_Param.GetValue ('REPSTATC10')
    else if Copy(st,1,3)='P1K' then Chtmp:=Tob_Param.GetValue ('REPSTATC11')
    else if Copy(st,1,3)='P1L' then Chtmp:=Tob_Param.GetValue ('REPSTATC12')
    else if Copy(st,1,3)='P1M' then Chtmp:=Tob_Param.GetValue ('REPSTATC13')
    else if Copy(st,1,3)='P1N' then Chtmp:=Tob_Param.GetValue ('REPSTATC14')
    else if Copy(st,1,3)='P1O' then Chtmp:=Tob_Param.GetValue ('REPSTATC15')
    else if Copy(st,1,3)='P1P' then Chtmp:=Tob_Param.GetValue ('REPSTATC16')
    else if Copy(st,1,3)='P1Q' then Chtmp:=Tob_Param.GetValue ('REPSTATC17')
    else if Copy(st,1,3)='P1R' then Chtmp:=Tob_Param.GetValue ('REPSTATC18')
    else if Copy(st,1,3)='P1S' then Chtmp:=Tob_Param.GetValue ('REPSTATC19')
    else if Copy(st,1,3)='P1T' then Chtmp:=Tob_Param.GetValue ('REPSTATC20');

    if Chtmp<>'' then
    begin
      for cpt:=1 to 10 do
      begin
        if (cpt=10) and (Chtmp='YTC_TABLELIBRETIERSA') then break
        else if Chtmp='YTC_TABLELIBRETIERS'+InttoStr(cpt) then break ;
      end ;
      if cpt<11 then
      begin
        DerTableLibreAff:=cpt;
        DerTableLibreArg:=copy(st,1,3);
      end ;
    end;
  end;

  if DerTableLibreAff<>0 then                      // si affectation trouvee
  begin
    TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);
    //TPST.initValeurs;
    if DerTableLibreAff=10 then TPST.PutValue('YX_TYPE', 'LTA')
    else TPST.PutValue('YX_TYPE', 'LT'+InttoStr(DerTableLibreAff));
    TPST.PutFixedStValue('YX_CODE', st, 9, 6, tctrim, true);
    TPST.PutFixedStValue('YX_LIBELLE', st, 15, 25, tcChaine, False);
    TPST.PutFixedStValue('YX_ABREGE', st, 15, 17, tcChaine, False);
    ImportTablesLibres (TPST);
    TPST.Free;
  end ;
end;

////////////////////////////////////////////////////////////
// Intégration des statistiques Fournisseurs
////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTablesLibresFou (st : string);
var TPST: TOB;
    Chtmp : string;
    cpt : integer;
begin
  if DerTableLibreArg<>copy(st,1,3) then           // recherche de la table concernée
  begin
    Chtmp:='' ;
    DerTableLibreAff:=0;
    if Copy(st,1,3)='P4A'      then Chtmp:=Tob_Param.GetValue ('REPSTATF01')
    else if Copy(st,1,3)='P4B' then Chtmp:=Tob_Param.GetValue ('REPSTATF02')
    else if Copy(st,1,3)='P4C' then Chtmp:=Tob_Param.GetValue ('REPSTATF03')
    else if Copy(st,1,3)='P4D' then Chtmp:=Tob_Param.GetValue ('REPSTATF04')
    else if Copy(st,1,3)='P4E' then Chtmp:=Tob_Param.GetValue ('REPSTATF05')
    else if Copy(st,1,3)='P4F' then Chtmp:=Tob_Param.GetValue ('REPSTATF06')
    else if Copy(st,1,3)='P4G' then Chtmp:=Tob_Param.GetValue ('REPSTATF07')
    else if Copy(st,1,3)='P4H' then Chtmp:=Tob_Param.GetValue ('REPSTATF08')
    else if Copy(st,1,3)='P4I' then Chtmp:=Tob_Param.GetValue ('REPSTATF09')
    else if Copy(st,1,3)='P4J' then Chtmp:=Tob_Param.GetValue ('REPSTATF10')
    else if Copy(st,1,3)='P4K' then Chtmp:=Tob_Param.GetValue ('REPSTATF11')
    else if Copy(st,1,3)='P4L' then Chtmp:=Tob_Param.GetValue ('REPSTATF12')
    else if Copy(st,1,3)='P4M' then Chtmp:=Tob_Param.GetValue ('REPSTATF13')
    else if Copy(st,1,3)='P4N' then Chtmp:=Tob_Param.GetValue ('REPSTATF14')
    else if Copy(st,1,3)='P4O' then Chtmp:=Tob_Param.GetValue ('REPSTATF15')
    else if Copy(st,1,3)='P4P' then Chtmp:=Tob_Param.GetValue ('REPSTATF16')
    else if Copy(st,1,3)='P4Q' then Chtmp:=Tob_Param.GetValue ('REPSTATF17')
    else if Copy(st,1,3)='P4R' then Chtmp:=Tob_Param.GetValue ('REPSTATF18')
    else if Copy(st,1,3)='P4S' then Chtmp:=Tob_Param.GetValue ('REPSTATF19')
    else if Copy(st,1,3)='P4T' then Chtmp:=Tob_Param.GetValue ('REPSTATF20') ;

    if Chtmp<>'' then
    begin
      for cpt:=1 to 3 do
      begin
        if Chtmp='YTC_TABLELIBREFOU'+InttoStr(cpt) then break ;
      end ;
      if cpt<4 then
      begin
        DerTableLibreAff:=cpt;
        DerTableLibreArg:=copy(st,1,3);
      end ;
    end;
  end ;

  if DerTableLibreAff<>0 then                      // si affectation trouvee
  begin
    TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);
    //TPST.initValeurs;
    if DerTableLibreAff=10 then TPST.PutValue('YX_TYPE', 'LTA')
    else TPST.PutValue('YX_TYPE', 'LF'+InttoStr(DerTableLibreAff));
    TPST.PutFixedStValue('YX_CODE', st, 9, 6, tctrim, true);
    TPST.PutFixedStValue('YX_LIBELLE', st, 15, 25, tcChaine, False);
    TPST.PutFixedStValue('YX_ABREGE', st, 15, 17, tcChaine, False);
    ImportTablesLibres (TPST);
    TPST.Free;
  end ;
end;

/////////////////////////////////////////////////////////////////////////////
// Par - REPRISE DES ARRONDIS
/////////////////////////////////////////////////////////////////////////////
procedure ImportArrondi ( T : TOB );       // Mettre des controles de cohérence
begin
 T.SetAllModifie(TRUE) ;
 T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure RecupArrondi (st : string);
var TPST: TOB;
    poids : double;
    methode : string;
begin

  // Création de la TOB MERE si inexistante
	if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('ARRONDI', Tob_Parametre, -1);          // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;

  TPST.PutFixedStValue('GAR_CODEARRONDI', st, 9, 1, tcChaine, true);
  TPST.PutValue ('GAR_RANG', 1);
  TPST.PutValue ('GAR_VALEURSEUIL', 1);
  TPST.PutFixedStValue('GAR_LIBELLE', st, 15, 15, tcChaine, false);
  case st [30] of
      'S' : methode := 'S';
      'I' : methode := 'I';
      else methode := 'P';
  end;
  TPST.PutValue ('GAR_METHODE', methode);
  case st [31] of
     '0' : poids := 0.20;
     '1' : poids := 0.10;
     '2' : poids := 0.50;
     '4' : poids := 5;
     '5' : poids := 9;
     '6' : poids := 10;
     '7' : poids := 50;
     '8' : poids := 99;
     '9' : poids := 100;
     'A' : poids := 500;
     'B' : poids := 1000;
     'C' : poids := 5000;
     'D' : poids := 10000;
     else poids := 1;
  end;
  TPST.PutValue ('GAR_POIDSARRONDI', poids);
  // ImportArrondi (TPST);
  // TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
//******* Pre-ve-ca - REPRISE DES REPRESENTANTS/VENDEURS/CAISSIERS
/////////////////////////////////////////////////////////////////////////////
procedure ImportVendeur ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure RecupRepresentant (st : string);
var TPST: TOB;
begin
TPST := TOB.CREATE  ('COMMERCIAL', NIL, -1);
//TPST.initValeurs;
TPST.PutValue       ('GCL_TYPECOMMERCIAL', 'REP');
TPST.PutFixedStValue('GCL_COMMERCIAL', st, 27, 17, tcChaine, true);
TPST.PutFixedStValue('GCL_ETABLISSEMENT', st, 9, 3, tcChaine, false);
TPST.PutFixedStValue('GCL_LIBELLE', st, 27, 20, tcChaine, true);
TPST.PutFixedStValue('GCL_PRENOM', st, 47, 15, tcChaine, true);
TPST.PutFixedStValue('GCL_COMMISSION', st, 15, 12, tcDouble1000, false);
TPST.PutValue       ('GCL_TYPECOM', 'CA');
TPST.PutValue       ('GCL_DATESUPP', iDate2099);
ImportVendeur (TPST);
TPST.Free;
end;

Procedure RecupVendeur (st : string);
var TPST           : TOB    ;
    CodeVendeur    : string ;
begin
TPST := TOB.CREATE ('COMMERCIAL', NIL, -1);
//TPST.initValeurs;
TPST.PutValue       ('GCL_TYPECOMMERCIAL', 'VEN');
CodeVendeur := copy (st, 9, 3) + copy (st, 12, 2);
TPST.PutValue       ('GCL_COMMERCIAL', CodeVendeur);
//TPST.PutFixedStValue('GCL_COMMERCIAL'    , st, 27, 17, tcChaine, true);
TPST.PutFixedStValue('GCL_ETABLISSEMENT' , st, 9, 3, tcChaine, false);
TPST.PutFixedStValue('GCL_LIBELLE'       , st, 27, 20, tcChaine, true);
TPST.PutFixedStValue('GCL_PRENOM'        , st, 47, 15, tcChaine, true);
TPST.PutFixedStValue('GCL_COMMISSION'    , st, 15, 12, tcDouble1000, false);
TPST.PutValue       ('GCL_TYPECOM'       , 'CA');
TPST.PutValue       ('GCL_DATESUPP'      , iDate2099);
ImportVendeur (TPST);
TPST.Free;
end;

Procedure RecupCaissier (st : string);
var TPST         : TOB    ;
    CodeCaissier : string ;
begin
TPST := TOB.CREATE ('COMMERCIAL', NIL, -1);
//TPST.initValeurs;
TPST.PutValue       ('GCL_TYPECOMMERCIAL', 'CAI');
CodeCaissier := copy (st, 9, 3) + copy (st, 12, 2);
TPST.PutValue       ('GCL_COMMERCIAL', CodeCaissier);
//TPST.PutFixedStValue('GCL_COMMERCIAL', st, 15, 17, tcChaine, true);
TPST.PutFixedStValue('GCL_ETABLISSEMENT', st, 9, 3, tcChaine, false);
TPST.PutFixedStValue('GCL_LIBELLE'      , st, 15, 20, tcChaine, true);
TPST.PutFixedStValue('GCL_PRENOM'       , st, 35, 15, tcChaine, true);
TPST.PutValue       ('GCL_TYPECOM'      , 'CA');
TPST.PutValue       ('GCL_DATESUPP'     , iDate2099);
ImportVendeur (TPST);
TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
//******* Pbo - REPRISE DES BOUTIQUES
/////////////////////////////////////////////////////////////////////////////
procedure ImportBoutique ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure RecupBoutique (st : string);
var TPST: TOB;
begin
	//////////////////////////////
  // Création de l'établissement
  //////////////////////////////
  TPST := TOB.CREATE ('ETABLISS', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;

  TPST.PutFixedStValue('ET_ETABLISSEMENT', st,  9,  3, tcChaine, true);
  TPST.PutFixedStValue('ET_LIBELLE'      , st, 15, 15, tcChaine, false);
  TPST.PutFixedStValue('ET_ABREGE'       , st, 15, 15, tcChaine, false);
  TPST.PutFixedStValue('ET_PAYS'         , st, 69,  3, tcChaine, false);
  TPST.PutFixedStValue('ET_LANGUE'       , st, 57,  3, tcChaine, false);
  ImportBoutique (TPST);
  TPST.Free;

  /////////////////////////////////////////////////////////////////
  // Création du dépôt associé : Code Dépôt = Code Etablissement
  /////////////////////////////////////////////////////////////////
  TPST := TOB.CREATE ('DEPOTS', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;

  TPST.PutFixedStValue('GDE_DEPOT'  , st,  9,  3, tcChaine, true);
  TPST.PutFixedStValue('GDE_LIBELLE', st, 15, 15, tcChaine, false);
  TPST.PutFixedStValue('GDE_ABREGE' , st, 15, 15, tcChaine, false);
  TPST.PutFixedStValue('GDE_PAYS'   , st, 69,  3, tcChaine, false);
  ImportBoutique (TPST);
  TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Chargement en TOB pour saisie d'une correspondance
/////////////////////////////////////////////////////////////////////////////
Procedure RecupTableCorrespondance ( TOBPrinc : TOB ; NomChamp, ValeurGPAO, LibelleGPAO : String ) ;
var TobFille : TOB ;
begin
  if (TOBPrinc = Nil) or (NomChamp = '') or (ValeurGPAO = '') then Exit ;

  TobFille :=  TOBPrinc.FindFirst(['GRO_TYPEREPRISE','GRO_NOMCHAMP','GRO_VALEURGPAO'],
                                  [TypeRepriseGB2000, ExtractSuffixe(NomChamp), ValeurGPAO], False) ;
  if TobFille = Nil then
  begin
    TobFille := TOB.Create('REPRISECOGPAO', TOBPrinc, -1) ;
    //TobFille.InitValeurs ;
    TobFille.PutValue('GRO_TYPEREPRISE', TypeRepriseGB2000) ;
    TobFille.PutValue('GRO_NOMCHAMP', ExtractSuffixe(NomChamp)) ;
    TobFille.PutValue('GRO_VALEURGPAO', ValeurGPAO) ;
    TobFille.PutValue('GRO_LIBELLEGPAO', LibelleGPAO) ;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
//******* Pmr - REPRISE DES MODES DE PAIEMENTS
// Chargement en TOB pour saisie d'une correspondance
/////////////////////////////////////////////////////////////////////////////
Procedure RecupModePaiement (st : string);
var CodePaie   : string ;
    Libelle    : string ;
begin

  CodePaie   := Trim (copy (st, 9 , 3));
  Libelle    := Trim (copy (st, 15,25));
  RecupTableCorrespondance (Tob_ModePaie, 'GPE_MODEPAIE', CodePaie, Libelle) ;

end;

/////////////////////////////////////////////////////////////////////////////
//******* Pde - REPRISE DES DEVISES
// Chargement en TOB pour saisie d'une correspondance
/////////////////////////////////////////////////////////////////////////////
Procedure RecupDevise (st : string);
var CodeDev    : string ;
    Libelle    : string ;
begin

  CodeDev   := Trim (copy (st,  9, 3));
  Libelle   := Trim (copy (st, 15,25));
  RecupTableCorrespondance (Tob_Devise, 'T_DEVISE', CodeDev, Libelle) ;

end;

/////////////////////////////////////////////////////////////////////////////
//******* Ppa - REPRISE DES PAYS
// Chargement en TOB pour saisie d'une correspondance
/////////////////////////////////////////////////////////////////////////////
Procedure RecupPays (st : string);
var CodePays   : string ;
    Libelle    : string ;
begin

  CodePays   := Trim (copy (st,  9, 3));
  Libelle    := Trim (copy (st, 15,25));
  RecupTableCorrespondance (Tob_Pays, 'T_PAYS', CodePays, Libelle) ;

end;

/////////////////////////////////////////////////////////////////////////////
//******* PL1 - REPRISE DES CODES POSTAUX
/////////////////////////////////////////////////////////////////////////////
Procedure ImportCodePostaux ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure RecupCodePostaux (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CODEPOST', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;

  TPST.PutFixedStValue('O_PAYS'      , st,  8,  3, tcChaine, true);
  TPST.PutFixedStValue('O_CODEPOSTAL', st, 11,  7, tcChaine, false);
  TPST.PutFixedStValue('O_VILLE'     , st, 23, 32, tcChaine, true);
end;

/////////////////////////////////////////////////////////////////////////////
//******* Ppo - REPRISE DES POIDS
/////////////////////////////////////////////////////////////////////////////
procedure ImportMea ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure RecupPoids (st : string);
var TPST: TOB;
    Chtmp : String;
    Quotite : Double;
begin
 // Création de la TOB MERE si inexistante
 if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

 TPST := TOB.CREATE ('MEA', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
 //TPST.initValeurs;
 TPST.PutValue('GME_QUALIFMESURE', 'POI');
 TPST.PutFixedStValue('GME_MESURE', st, 9, 3, tcTrim, true);
 TPST.PutFixedStValue('GME_LIBELLE', st, 15, 25, tcTrim, False);
 if Length(st)>40 then Chtmp:=Trim(copy(st,41,3))
                  else Chtmp:='';
 if Chtmp='KGR' then Quotite:=1
 else if Chtmp='T'   then Quotite:=1000
 else if Chtmp='DAG' then Quotite:=0.01
 else if Chtmp='CG'  then Quotite:=0.00001
 else if Chtmp='HG'  then Quotite:=0.1
 else if Chtmp='G'   then Quotite:=0.001
 else if Chtmp='DG'  then Quotite:=0.0001
 else Quotite:=0;
 TPST.PutValue('GME_QUOTITE', Quotite);
 //ImportMea (TPST);
 //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Pfs - REPRISE DES FAMILLES
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupFamille (st : string);
var TPST          : TOB     ;
    Famille       : string  ;
    Libelle       : string  ;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  if RecodifieFamille = 'X' then
  begin
    Famille    := Trim (copy (st,  9, 6));
    Libelle    := Trim (copy (st, 27,25));
    RecupTableCorrespondance (Tob_Famille, '$$_CODEFAMSF', Famille, Libelle) ;
  end else
  begin
    if StrToInt (Fniv1) > 0 then
    begin
      TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
      //TPST.initValeurs;
      TPST.PutFixedStValue('CC_TYPE', 'FN1', 1, 3, tcChaine, true);
      TPST.PutFixedStValue('CC_CODE',   st , 9, StrToInt (Fniv1), tcChaine, true);
      if (StrToInt (Fniv2) = 0) then
      begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 27, 25, tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE' , st, 27, 17, tcChaine, False);
      end
      else begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 9, StrToInt (Fniv1), tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE' , st, 9, StrToInt (Fniv1), tcChaine, False);
      end;
      //ImportChoixCod (TPST);
    end;
    if StrToInt (Fniv2) > 0 then
    begin
      TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
      //TPST.initValeurs;
      TPST.PutFixedStValue('CC_TYPE', 'FN2', 1, 3, tcChaine, true);
      TPST.PutFixedStValue('CC_CODE', st, 9+StrToInt (Fniv1), StrToInt (Fniv2), tcChaine, true);
      if (StrToInt (Fniv3) = 0) then
      begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 27, 25, tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE' , st, 27, 17, tcChaine, False);
      end
      else begin
        TPST.PutFixedStValue('CC_LIBELLE', st, 9+StrToInt (Fniv1), StrToInt (Fniv2), tcChaine, False);
        TPST.PutFixedStValue('CC_ABREGE' , st, 9+StrToInt (Fniv1), StrToInt (Fniv2), tcChaine, False);
      end;
      //ImportChoixCod (TPST);
    end;
    if StrToInt (Fniv3) > 0 then
    begin
     TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
     //TPST.initValeurs;
     TPST.PutFixedStValue('CC_TYPE'   , 'FN3', 1, 3, tcChaine, true);
     TPST.PutFixedStValue('CC_CODE'   , st, 9+StrToInt (Fniv1) + StrToInt(Fniv2), StrToInt (Fniv3), tcChaine, true);
     TPST.PutFixedStValue('CC_LIBELLE', st, 27, 25, tcChaine, False);
     TPST.PutFixedStValue('CC_ABREGE' , st, 27, 17, tcChaine, False);
    end;
  end;

  ////////////////////////////////////////////////////////////////////
  // MODIF LM 29/07/02
  // Alimente éventuellement une table statistique
  ////////////////////////////////////////////////////////////////////
  RecupTablesLibresArt (st) ;


end;

/////////////////////////////////////////////////////////////////////////////
// Prf :  Reprise des regroupements de familles
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupRegroupeFamille (st : string);
var TPST: TOB;
    Chtmp: String;
    FamilleNiv1, FamilleNiv2 : integer;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  FamilleNiv1 := Tob_Param.GetValue ('FAMNIV1');
  FamilleNiv2 := Tob_Param.GetValue ('FAMNIV2');
  
  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;

  Chtmp:=Copy(st,9,6);
  if (Length(Trim(Chtmp))=FamilleNiv1) and (FamilleNiv2>0) then
  begin
    TPST.PutFixedStValue('CC_TYPE'   , 'FN1', 1, 3, tcChaine, true);
    TPST.PutFixedStValue('CC_CODE'   , st, 9, FamilleNiv1, tcChaine, true);
    TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcChaine, False);
    TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcChaine, False);
    //ImportChoixCod (TPST);
  end;

  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Psa :  Reprise des saisons
/////////////////////////////////////////////////////////////////////////////
Procedure RecupCollection (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'GCO', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   ,  st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE',  st, 45, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' ,  st, 45, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Pcc : Reprise des catégories clients
/////////////////////////////////////////////////////////////////////////////
Procedure RecupCategorieClient (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'TRC', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 27, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 27, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Plg : Reprise des langues
/////////////////////////////////////////////////////////////////////////////
Procedure RecupLangue (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE  ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'LGU', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Pfj : Reprise des formes juridiques
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupCivilite (st : string);
var TPST: TOB;
begin
  if copy(st,40,1)='I' then
  begin
    // Création de la TOB MERE si inexistante
    if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

    TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
    //TPST.initValeurs;
    TPST.PutFixedStValue('CC_TYPE', 'CIV', 1, 3, tcChaine, true);
    TPST.PutFixedStValue('CC_CODE', st, 9, 3, tcTrim, true);
    if length(trim(copy(st,9,4)))>3 then EcrireFichAno('Civilité '+copy(st,9,4)+' tronquée à '+copy(st,9,3));
    TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcTrim, False);
    TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcTrim, False);
    //ImportChoixCod (TPST);
    //TPST.Free;
  end ;

  // Alimentation des formes juridiques ????
  {TPST := TOB.CREATE ('COMMUN', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  TPST.initValeurs;
  TPST.PutFixedStValue('CO_TYPE', 'JUR', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CO_CODE', st, 9, 3, tcTrim, true);
  if length(trim(copy(st,9,4)))>3 then EcrireFichAno('Forme Juridique '+copy(st,9,4)+' tronquée à '+copy(st,9,3));
  TPST.PutFixedStValue('CO_LIBELLE', st, 15, 25, tcTrim, False);
  TPST.PutFixedStValue('CO_ABREGE', st, 15, 17, tcTrim, False);
  TPST.InsertOrUpdateDB (FALSE);
  TPST.Free; }

end;

/////////////////////////////////////////////////////////////////////////////
// Pdm : Reprise des démarques
/////////////////////////////////////////////////////////////////////////////
Procedure RecupDemarque (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'GTM', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st,  9,  1, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Pdm : Reprise des démarques
/////////////////////////////////////////////////////////////////////////////
Procedure RecupMotif_ES (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'GMM', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st,  9,  3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Prg : Reprise des motifs de remise
/////////////////////////////////////////////////////////////////////////////
Procedure RecupRemise (st : string);
var TPST: TOB;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE ('CHOIXCOD', Tob_Parametre, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'QFR', 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 15, 25, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 15, 17, tcTrim, False);
  //ImportChoixCod (TPST);
  //TPST.Free;
end;

/////////////////////////////////////////////////////////////////////////////
// Récupération du taux de TVA par défaut
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTauxTVADefaut ;
var
  SQL         : string ;
  Q           : TQUERY ;
  TVA, Regime : string ;
begin
  TVA     := Tob_Param.GetValue ('REPTVA');
  Regime  := Tob_Param.GetValue ('REPREG');

  SQL:='Select TV_TAUXVTE From TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="'+TVA+'" AND TV_REGIME="'+Regime+'" ';
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then TauxTVADefaut := Q.FindField('TV_TAUXVTE').AsFloat
  else TauxTVADefaut:=19.6;
  Ferme(Q) ;
end;

/////////////////////////////////////////////////////////////////////////////
// Récupération du paramétrage de reprise des articles
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.ChargeParamArticle ;
begin
  //
  // ARFC1 ...
  //
  OptimisationArticle := Tob_Param.GetValue ('OPTIMISATIONARTICLE');
  RecodifieFamille    := Tob_Param.GetValue ('RECODIFIEFAMILLE');
  FNiv1               := Tob_Param.GetValue ('FAMNIV1');
  FNiv2               := Tob_Param.GetValue ('FAMNIV2');
  FNiv3               := Tob_Param.GetValue ('FAMNIV3');
  Famille             := Tob_Param.GetValue ('REPFAMILLE');
  Modele              := Tob_Param.GetValue ('REPMOD');
  Matiere             := Tob_Param.GetValue ('REPMAT');
  Coloris             := Tob_Param.GetValue ('REPCOL');
  Finition            := Tob_Param.GetValue ('REPFIN');
  CatDim              := Tob_Param.GetValue ('CATEGORIEDIM');
  Compo1              := Tob_Param.GetValue ('REPCP1');
  Compo2              := Tob_Param.GetValue ('REPCP2');
  TVA                 := Tob_Param.GetValue ('REPTVA');
  DeviseSiteCentral   := Tob_Param.GetValue ('REPDEVISE') ;
  //
  // ARFC2 .....
  Categorie:= Tob_Param.GetValue ('REPCAT');
  Formule  := Tob_Param.GetValue ('REPFOR');
  Structure:= Tob_Param.GetValue ('REPSTR');
  TypeArt  := Tob_Param.GetValue ('REPTYP');
  LibCol   := Tob_Param.GetValue ('REPLIBCOL');

end;

/////////////////////////////////////////////////////////////////////////////
// Récupération du paramétrage de reprise des articles
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.ChargeDeviseCentral ;
begin
  DeviseSiteCentral := Tob_Param.GetValue ('REPDEVISE') ;
end;

/////////////////////////////////////////////////////////////////////////////
// chargmenent en TOB des devises PGI pour optimisation
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.ChargeDevisePGI ;
var Q : TQUERY ;
begin
  //
  // Chargement de la TOB des devises permettant de convertir les montants en devise
  //
  Tob_DevisePGI := TOB.Create('LES DEVISES', Nil, -1) ;
  Q             := OpenSQL('SELECT * FROM DEVISE',True) ;
  Tob_DevisePGI.LoadDetailDB('DEVISE','','',Q,False) ;
  Ferme(Q) ;
end;

/////////////////////////////////////////////////////////////////////////////
// Chargement de la TOB des grilles de tailles pour intégrer les articles
/////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.ChargeTobTaille ;
var TPST : TOB;
     Q : TQuery ;
     CatDim : string;
begin
  CatDim := Tob_Param.GetValue ('CATEGORIEDIM');

  TPST := TOB.CREATE ('DIMENSION', Tob_Tailles, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  Q := OpenSQL('SELECT * FROM DIMENSION WHERE GDI_TYPEDIM ="'+CatDim+'"',True) ;
  TPST.LoadDetailDB('DIMENSION','','',Q,False) ;
  Ferme(Q) ;

end;

//////////////////////////////////////////////////////////
// Suppression des éléments d'une pièce avant intégration
//////////////////////////////////////////////////////////
function SupprimeElementsPiece (Nature, Souche : string; Numero, indice : integer) : boolean ;
var Condition : string ;
    SQL       : string ;
begin
  ////////////////////////////////////////////////////////////////////////////
  // Suppession de PIECE
  ////////////////////////////////////////////////////////////////////////////
  Condition:= 'GP_NATUREPIECEG="'+Nature+'" AND GP_SOUCHE="'+Souche+'" AND GP_NUMERO="'+IntToStr(Numero)+'" AND GP_INDICEG="'+IntToStr(Indice)+'"';
  SQL:='Delete From PIECE WHERE'+ ' ' + Condition ;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des LIGNES
  ////////////////////////////////////////////////////////////////////////////
  Condition:= 'GL_NATUREPIECEG="'+Nature+'" AND GL_SOUCHE="'+Souche+'" AND GL_NUMERO="'+IntToStr(Numero)+'" AND GL_INDICEG="'+IntToStr(Indice)+'"';
  SQL:='Delete From LIGNE WHERE'+ ' ' + Condition ;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession dans PIEDBASE
  ////////////////////////////////////////////////////////////////////////////
  Condition:= 'GPB_NATUREPIECEG="'+Nature+'" AND GPB_SOUCHE="'+Souche+'" AND GPB_NUMERO="'+IntToStr(Numero)+'" AND GPB_INDICEG="'+IntToStr(Indice)+'"';
  SQL:='Delete From PIEDBASE WHERE'+ ' '+ Condition ;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des PIEDECHE
  ////////////////////////////////////////////////////////////////////////////
  Condition:= 'GPE_NATUREPIECEG="'+Nature+'" AND GPE_SOUCHE="'+Souche+'" AND GPE_NUMERO="'+IntToStr(Numero)+'" AND GPE_INDICEG="'+IntToStr(Indice)+'"';
  SQL:='Delete From PIEDECHE WHERE'+ ' '+ Condition ;
  ExecuteSQL(SQL);
  ////////////////////////////////////////////////////////////////////////////
  // Suppession des ACOMPTES
  ////////////////////////////////////////////////////////////////////////////
  Condition:= 'GAC_NATUREPIECEG="'+Nature+'" AND GAC_SOUCHE="'+Souche+'" AND GAC_NUMERO="'+IntToStr(Numero)+'" AND GAC_INDICEG="'+IntToStr(Indice)+'"';
  SQL:='Delete From ACOMPTES WHERE'+ ' '+ Condition ;
  ExecuteSQL(SQL);
  result:=True;
end;


//
// Elimination des commandes et réceptions "parasites"
//
Procedure TFRecupGB2000.VerifPiece ;
var TPST     : TOB;
    TobFille : TOB ;
     Q       : TQuery  ;
     i       : integer;
     TotalQteFact : double ;
     TotalQteSto  : double ;
     TotalHt      : double ;
     TotalTTC     : double ;
     Nature       : string ;
	 Souche       : string ;
     Numero       : integer;
     Indice       : integer;
begin
  //
  // Verification des commandes fournisseurs
  //
  TPST := TOB.CREATE ('Les pieces', nil, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  Q := OpenSQL('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="CF"',True) ;
  TPST.LoadDetailDB('PIECE','','',Q,False) ;
  Ferme(Q) ;

  for i:=0 to TPST.detail.count-1 do
  begin
    TobFille     := TPST.detail [i];
    TotalQteFact := TobFille.GetValue ('GP_TOTALQTEFACT');
    TotalQteSto  := TobFille.GetValue ('GP_TOTALQTESTOCK');
    TotalHt      := TobFille.GetValue ('GP_TOTALHTDEV');
    TotalTTC     := TobFille.GetValue ('GP_TOTALTTCDEV');

    if (TotalQteFact = 0) and (TotalQteSto = 0) and (TotalHt = 0) and (TotalTTC = 0) then
    begin
      Nature   := TobFille.GetValue ('GP_NATUREPIECEG');
      Souche   := TobFille.GetValue ('GP_SOUCHE');
      Numero   := TobFille.GetValue ('GP_NUMERO');
      Indice   := TobFille.GetValue ('GP_INDICEG');
      SupprimeElementsPiece (Nature, Souche, Numero, Indice);
    end;
  end;

  TPST.free ;
  //
  // Verification des récptions fournisseurs
  //
  TPST := TOB.CREATE ('Les pieces', nil, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  Q := OpenSQL('SELECT * FROM PIECE WHERE GP_NATUREPIECEG="BLF"',True) ;
  TPST.LoadDetailDB('PIECE','','',Q,False) ;
  Ferme(Q) ;

  for i:=0 to TPST.detail.count-1 do
  begin
    TobFille := TPST.detail [i];
    TotalQteFact := TobFille.GetValue ('GP_TOTALQTEFACT');
    TotalQteSto  := TobFille.GetValue ('GP_TOTALQTESTOCK');
    TotalHt      := TobFille.GetValue ('GP_TOTALHTDEV');
    TotalTTC     := TobFille.GetValue ('GP_TOTALTTCDEV');

    if (TotalQteFact = 0) and (TotalQteSto = 0) and (TotalHt = 0) and (TotalTTC = 0) then
    begin
      Nature   := TobFille.GetValue ('GP_NATUREPIECEG');
      Souche   := TobFille.GetValue ('GP_SOUCHE');
      Numero   := TobFille.GetValue ('GP_NUMERO');
      Indice   := TobFille.GetValue ('GP_INDICEG');
      SupprimeElementsPiece (Nature, Souche, Numero, Indice);
    end;
  end;
  TPST.free ;
end;

/////////////////////////////////////////////////////////////////////////////
// Pgr :  Récupération des grilles de tailles
/////////////////////////////////////////////////////////////////////////////
procedure ImportMasqueDim ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

procedure ImportChoixCod ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

Procedure TFRecupGB2000.RecupGrille (st : string);
var TPST: TOB;
    i, posdebut, fin, rang : integer;
    nbtaille, chtmp : string;
    CategorieDimension : string;
begin
  if ( st[5] = '1' ) then
  begin
    posdebut :=  17 - 4;
    nbtaille := copy (st, 15, 2) ;
    fin := strtoint (nbtaille) ;
    if fin > 27 then fin := 27 ;
    rang := 0;
  end else
  begin
    posdebut :=  15 - 4;
    fin := 3 ;
    rang := 27;
  end ;

  CategorieDimension := Tob_Param.GetValue ('CATEGORIEDIM');
  // Création de la dimension dans la tablette
  TPST := TOB.CREATE ('CHOIXCOD', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('CC_TYPE'   , 'GG'+CategorieDimension[3], 1, 3, tcChaine, true);
  TPST.PutFixedStValue('CC_CODE'   , st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('CC_LIBELLE', st, 9, 3, tcTrim, False);
  TPST.PutFixedStValue('CC_ABREGE' , st, 9, 3, tcTrim, False);
  ImportChoixCod (TPST);
  TPST.Free;

  // Création du masque de la dimension
  TPST := TOB.CREATE ('DIMMASQUE', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
  //TPST.initValeurs;
  TPST.PutFixedStValue('GDM_MASQUE' , st, 9, 3, tcTrim, true);
  TPST.PutFixedStValue('GDM_LIBELLE', st, 9, 3, tcTrim, False);
  for i:=49 to 53 do
  begin
    if chr(i)=CategorieDimension[3] then
    begin
      TPST.PutFixedStValue('GDM_TYPE'+chr(i), st, 9, 3, tcTrim, true);
      TPST.PutValue('GDM_POSITION'+chr(i), 'CO1');
    end else
    begin
      TPST.PutValue('GDM_TYPE'+chr(i), '   ');
      TPST.PutValue('GDM_POSITION'+chr(i), '   ');
    end ;
  end ;
  TPST.PutValue ('GDM_TYPEMASQUE', '...');
  ImportMasqueDim (TPST);
  TPST.Free;

  // Création du contenu de la table dimension
  for i := 1 to fin do
  begin
    if length (st) >= (posdebut + (i * 4) - 3) then
    begin
      TPST := TOB.CREATE ('DIMENSION', Tob_Tailles, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
      TPST.PutValue('GDI_TYPEDIM', CategorieDimension);
      TPST.PutFixedStValue('GDI_GRILLEDIM', st, 9, 3, tcTrim, true);
      if (rang+i)<10  then chtmp:='00'+ chr(48 + rang + i)
                      else chtmp:='00'+ chr(65 + rang + i - 10);
      TPST.PutValue('GDI_CODEDIM', chtmp);
      TPST.PutFixedStValue('GDI_LIBELLE', st, posdebut + (i * 4), 4, tcTrim, false);
      TPST.PutValue('GDI_RANG', rang + i);
    end;
  end;
end;

procedure CodeCodeInterne (var Code: String);
Var i : integer ;
    CL : String ;
BEGIN
CL:=uppercase(Trim(Code)) ;
if Length(CL)<3 then Code:='001' else
   BEGIN
   i:=3 ; While CL[i]='Z' do BEGIN CL[i]:='0' ; Dec(i) ; END ;
   if Ord(CL[i])=57 then CL[i]:='A' else CL[i]:=Succ(CL[i]) ;
   Code:=CL ;
   END ;
END ;


///////////////////////////////////////////////////////////////////////////////////
// Pvd : Récupération des ventes diverses -> Donne un article de type "Prestation"
///////////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupVenteDiverse (st : string);
Var TPST       : TOB    ;
	RefArtPre  : string ;
    Montant    : double ;
    MontantDev : double ;
    MontantOld : double ;
    DateConv   : TDateTime;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE    ('ARTICLE', Tob_Parametre, -1);
  //TPST.initValeurs;
  RefArtPre:= copy (st, 9, 2) + '          ' + '      ' + '   ' + '   ' + '   ' + '   ' + '   ' + 'X';
  TPST.PutValue         ('GA_ARTICLE'    , RefArtPre);
  TPST.PutFixedStValue  ('GA_CODEARTICLE', st, 9, 2, tcTrim, true);
  TPST.PutValue         ('GA_TYPEARTICLE', 'PRE');
  TPST.PutValue         ('GA_STATUTART'  , 'UNI');
  TPST.PutValue         ('GA_TENUESTOCK' , '-');
  TPST.PutValue         ('GA_PCB'        , '1');
  TPST.PutValue         ('GA_REMISELIGNE', 'X')  ;
  TPST.PutValue         ('GA_REMISEPIED' , 'X')  ;
  TPST.PutValue         ('GA_CALCPRIXHT' , 'AUC')  ;
  TPST.PutValue         ('GA_CALCPRIXTTC', 'AUC')  ;
  TPST.PutFixedStValue  ('GA_LIBELLE'    , st, 15, 15, tcTrimL, false);
  TPST.PutValue         ('GA_PRIXUNIQUE' , 'X');
  TPST.PutValue         ('GA_DATESUPPRESSION', iDate2099);
  TPST.PutValue         ('GA_DATECREATION'   , Date);

  Montant := StrToFloat (copy (st, 31, 12));
  Montant := Montant/100 ;

  if DeviseSiteCentral = V_PGI.DevisePivot then
  begin
    TPST.PutValue ('GA_PAHT', Montant);
    TPST.PutValue ('GA_PMAP', Montant);
    TPST.PutValue ('GA_DPA' , Montant);
    TPST.PutValue ('GA_PRHT', Montant);
    TPST.PutValue ('GA_PMRP', Montant);
    TPST.PutValue ('GA_DPR' , Montant);
  end else
  begin
    DateConv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    ToxConvertir (Montant, MontantDev, MontantOld, DeviseSiteCentral, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;

    TPST.PutValue ('GA_PAHT' , MontantDev);
    TPST.PutValue ('GA_PMAP' , MontantDev);
    TPST.PutValue ('GA_DPA'  , MontantDev);
    TPST.PutValue ('GA_PRHT' , MontantDev);
    TPST.PutValue ('GA_PMRP' , MontantDev);
    TPST.PutValue ('GA_DPR'  , MontantDev);
  end;
  //TPST.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
  //TPST.Free;
end ;

////////////////////////////////////////////////////////////////////////////////////////
// Poc : Récupération des opérations de caisses -> Donne un article de type "Financier"
///////////////////////////////////////////////////////////////////////////////////////
Procedure RecupOperationCaisse (st : string);
Var TPST      : TOB    ;
		RefArtPre : string ;
begin
  // Création de la TOB MERE si inexistante
  if (Tob_Parametre = nil) then Tob_Parametre := TOB.CREATE ('Les paramètres', NIL, -1);

  TPST := TOB.CREATE    ('ARTICLE', Tob_Parametre, -1);
  //TPST.initValeurs;
  RefArtPre:= copy (st, 9, 2) + '          ' + '      ' + '   ' + '   ' + '   ' + '   ' + '   ' + 'X';
  TPST.PutValue         ('GA_ARTICLE'    , RefArtPre);
  TPST.PutFixedStValue  ('GA_CODEARTICLE', st, 9, 2, tcTrim, true);
  TPST.PutValue         ('GA_TYPEARTICLE', 'FI');
  TPST.PutValue         ('GA_STATUTART'  , 'UNI');
  TPST.PutValue         ('GA_TENUESTOCK' , '-');
  TPST.PutValue         ('GA_REMISELIGNE', 'X')  ;
  TPST.PutValue         ('GA_REMISEPIED' , 'X')  ;
  TPST.PutValue         ('GA_PCB'        , '1')  ;
  TPST.PutValue         ('GA_CALCPRIXHT' , 'AUC')  ;
  TPST.PutValue         ('GA_CALCPRIXTTC', 'AUC')  ;
  TPST.PutFixedStValue  ('GA_LIBELLE'    , st, 15, 15, tcTrimL, false);
  TPST.PutValue         ('GA_DATECREATION'   , Date);
  TPST.PutValue         ('GA_DATESUPPRESSION', iDate2099);

  //TPST.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
  //TPST.Free;
end ;

/////////////////////////////////////////////////////////////////////////////
//******* ARFC1-2-3-A-B - REPRISE DES ARTICLES
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportArticle ;       // Mettre des controles de cohérence
Var TA, TG , TACompl  : TOB    ;
    Tob_Articles      : TOB    ;
    Tob_ArticlesCompl : TOB    ;
    //TOB_Catalogue     : TOB    ;
    Cpt               : integer;
    i                 : integer;
    CreeArticleCompl  : boolean;
begin
  /////////////////////////////////////////////////
  // Doit on créer les articles compl ???
  /////////////////////////////////////////////////
  Tob_ArticlesCompl := Nil;
  CreeArticleCompl := False;
  
  for i:=4 to 8 do
  begin
    if Tob_ArticleCompl.GetValue ('GA2_FAMILLENIV'+IntToStr (i)) <> '' then
    begin
      CreeArticleCompl := True;
      break ;
    end;
  end;
  if not CreeArticleCompl then
  begin
   for i:=1 to 2 do
    begin
      if Tob_ArticleCompl.GetValue ('GA2_STATART'+IntToStr (i))<> '' then
      begin
        CreeArticleCompl := True;
        break ;
      end;
    end;
  end;


  TOB_Articles :=TOB.Create('Liste des Articles',Nil,-1) ;

  if CreeArticleCompl then TOB_ArticlesCompl :=TOB.Create('Liste des Articles Compl',Nil,-1) ;
  //////////////////////////////////////////////////
  // Recherche de la grille taille associée
  // Création des fiches articles tailles associées
  //////////////////////////////////////////////////
  TG:=Tob_Tailles.FindFirst(['GDI_TYPEDIM','GDI_GRILLEDIM'],[CatDim,Tob_Article.GetValue('GA_GRILLEDIM'+CatDim[3])],TRUE) ;
  While (TG<>Nil) and (not ArretRecup) do
  BEGIN
   TA:=TOB.Create('ARTICLE',Tob_Articles,-1) ;
   //TA.Assign(TOB_Article) ;
   TA.Dupliquer (TOB_Article, False, True, True);

   For Cpt:=49 to 53 do
   begin
      if CatDim[3]=chr(Cpt)
      then TA.PutValue('GA_CODEDIM'+chr(Cpt),TG.GetValue('GDI_CODEDIM'))
      else TA.PutValue('GA_CODEDIM'+chr(Cpt),'   ') ;
   end ;
   TA.PutValue('GA_CODEDIM'+CatDim[3],TG.GetValue('GDI_CODEDIM')) ;
   ArticleCode (TA) ;
   TA.PutValue('GA_STATUTART', 'DIM');  // MODIF LM : Création d'un article dimensionné
   TA.PutValue('GA_DATESUPPRESSION', iDate2099);

   if CreeArticleCompl then
   begin
     TACompl:=TOB.Create('ARTICLECOMPL',Tob_ArticlesCompl,-1) ;
     //TACompl.Assign(TOB_ArticleCompl) ;
     TACompl.Dupliquer (TOB_ArticleCompl, False, True, True);

     TACompl.PutValue ('GA2_ARTICLE'    , TA.GetValue ('GA_ARTICLE'));
     TACompl.PutValue ('GA2_CODEARTICLE', TA.GetValue ('GA_CODEARTICLE'));
   end;

   TG:=Tob_Tailles.FindNext(['GDI_TYPEDIM','GDI_GRILLEDIM'],[CatDim,Tob_Article.GetValue('GA_GRILLEDIM'+CatDim[3])],TRUE) ;
  END ;

  ////////////////////////////////////////////////////
  // Création de l'article générique
  ////////////////////////////////////////////////////
  TA:=TOB.Create('ARTICLE',Tob_Articles,-1) ;
  //TA.Assign  (TOB_Article) ;
  TA.Dupliquer (TOB_Article, False, True, True);

  TA.PutValue('GA_CODEDIM1'  , '   ');
  TA.PutValue('GA_CODEDIM2'  , '   ');
  TA.PutValue('GA_CODEDIM3'  , '   ');
  TA.PutValue('GA_CODEDIM4'  , '   ');
  TA.PutValue('GA_CODEDIM5'  , '   ');
  TA.PutValue('GA_GRILLEDIM1', '   ');
  TA.PutValue('GA_GRILLEDIM2', '   ');
  TA.PutValue('GA_GRILLEDIM3', '   ');
  TA.PutValue('GA_GRILLEDIM4', '   ');
  TA.PutValue('GA_GRILLEDIM5', '   ');
  TA.PutValue('GA_STATUTART', 'GEN');  // MODIF LM : Création d'un article générique
  TA.PutValue('GA_DATESUPPRESSION', iDate2099);
  ArticleCode (TA) ;

  if CreeArticleCompl then
   begin
     TACompl:=TOB.Create('ARTICLECOMPL',Tob_ArticlesCompl,-1) ;
     TACompl.Assign(TOB_ArticleCompl) ;
     TACompl.PutValue ('GA2_ARTICLE'    , TA.GetValue ('GA_ARTICLE'));
     TACompl.PutValue ('GA2_CODEARTICLE', TA.GetValue ('GA_CODEARTICLE'));
   end;

  /////////////////////////////////////////////////
  // Alimentation du catalogue ARTICLE/FOURNISSEUR
  /////////////////////////////////////////////////
  //TOB_Catalogue :=TOB.Create ('CATALOGU', nil, -1) ;
  //if (TA.GetValue ('GA_COMMENTAIRE') <> '') then TOB_Catalogue.PutValue('GCA_REFERENCE', TA.GetValue('GA_COMMENTAIRE'))
  //else TOB_Catalogue.PutValue('GCA_REFERENCE', TA.GetValue('GA_CODEARTICLE'));
  //TOB_Catalogue.PutValue ('GCA_TIERS'        , TA.GetValue('GA_FOURNPRINC'));
  //TOB_Catalogue.PutValue ('GCA_ARTICLE'      , TA.GetValue('GA_ARTICLE'));
  //TOB_Catalogue.PutValue ('GCA_LIBELLE'      , TA.GetValue('GA_LIBELLE'));

  //Tob_Catalogue.SetAllModifie(TRUE) ;
  //Tob_Catalogue.InsertOrUpdateDB(FALSE);
  //Tob_Catalogue.Free ;

  /////////////////////////////////////////////////
  // MAJ DB
  /////////////////////////////////////////////////
  Tob_Articles.SetAllModifie(TRUE) ;

  if OptimisationArticle = 'X' then Tob_Articles.InsertDB(nil)
  else Tob_Articles.InsertOrUpdateDB(FALSE);
  Tob_Articles.Free ;

  Tob_Article  := nil ;

  if CreeArticleCompl then
  begin
    Tob_ArticlesCompl.SetAllModifie(TRUE) ;
    if OptimisationArticle = 'X' then Tob_ArticlesCompl.InsertDB(nil)
    else Tob_ArticlesCompl.InsertOrUpdateDB(FALSE);
  end;
  Tob_ArticlesCompl.Free ;
end ;


// Reprise ARFC1
Procedure TFRecupGB2000.RecupArticle1 (st : string);
var DateConv            : TDateTime ;
    Refart              : string    ;
    Debut               : string    ;
    CodeFamGB           : string    ;
    CodeFamPGI          : string    ;
    CodeFam1PGI         : string    ;
    CodeFam2PGI         : string    ;
    CodeFam3PGI         : string    ;
    Montant             : double    ;
    MontantDev          : double    ;
    MontantOld          : double    ;
    longueur            : integer   ;
begin
	// Import de la fiche article précédente
  //
  if Tob_Article <> Nil then ImportArticle;
  //
  // Création des TOB Articles
  //
  Tob_Article:= TOB.CREATE ('ARTICLE', NIL, -1);
  //Tob_Article.initValeurs;
  //
  // Statistiques articles
  //
  Tob_ArticleCompl:= TOB.CREATE ('ARTICLECOMPL', NIL, -1);
  //Tob_ArticleCompl.initValeurs;
  //
  // MODIF LM : ajout 6 blancs + X pour définir l'article générique
  //
  Refart :=  Copy(st, 7, 12)+ '      ' + '               X';
  //
  // Initialisation de la fiche article PGI
  //
  Tob_Article.PutValue            ('GA_ARTICLE'     , Refart);
  Tob_Article.PutFixedStValue     ('GA_CODEARTICLE' , st, 7, 12, tcTrim, true);

  Tob_ArticleCompl.PutValue       ('GA2_ARTICLE'    , Refart);
  Tob_Articlecompl.PutFixedStValue('GA2_CODEARTICLE', st, 7, 12, tcTrim, true);

  Tob_Article.PutValue             ('GA_TYPEARTICLE', 'MAR');
  Tob_Article.PutValue             ('GA_TENUESTOCK' , 'X')  ;
  Tob_Article.PutValue             ('GA_REMISELIGNE', 'X')  ;
  Tob_Article.PutValue             ('GA_REMISEPIED' , 'X')  ;
  Tob_Article.PutValue             ('GA_PCB'        , '1')  ;
  Tob_Article.PutValue             ('GA_CALCPRIXHT' , 'AUC')  ;
  Tob_Article.PutValue             ('GA_CALCPRIXTTC', 'AUC')  ;
  Tob_Article.PutValue             ('GA_QUALIFCODEBARRE', Tob_Param.GetValue ('REPCODEBARRE'));
  //
  // Saison -> Collection
  //
  Tob_Article.PutFixedStValue      ('GA_COLLECTION' , st, 19, 3, tcTrim, true);
  //
  // Fournisseur
  //
  Tob_Article.PutFixedStValue      ('GA_FOURNPRINC' , st, 22, 5, tcTrim, false);
  //
  // Familles sous-familles
  //
  if RecodifieFamille = 'X' then
  begin
    CodeFamGB := Trim (copy (st, 27, 6));
    CodeFamPGI := TraduitGPAOenPGI(TypeRepriseGB2000, '$$_CODEFAMSF', CodeFamGB, Tob_Famille);
    if CodeFamPGI <> '' then
    begin
      CodeFam1PGI   := ReadTokenSt(CodeFamPGI);
      CodeFam2PGI   := ReadTokenSt(CodeFamPGI);
      CodeFam3PGI   := ReadTokenSt(CodeFamPGI);
    end else
    begin
      CodeFam1PGI := '999' ;
      CodeFam2PGI := '999' ;
      CodeFam3PGI := '999' ;
    end;
    Tob_Article.PutValue ('GA_FAMILLENIV1', CodeFam1PGI );
    Tob_Article.PutValue ('GA_FAMILLENIV2', CodeFam2PGI );
    Tob_Article.PutValue ('GA_FAMILLENIV3', CodeFam3PGI );
  end else
  begin
    Tob_Article.PutFixedStValue      ('GA_FAMILLENIV1', st, 27                                , strtoint(FNiv1), tcTrim, true);
    Tob_Article.PutFixedStValue      ('GA_FAMILLENIV2', st, 27+strtoint(FNiv1)                , strtoint(FNiv2), tcTrim, true);
    Tob_Article.PutFixedStValue      ('GA_FAMILLENIV3', st, 27+strtoint(FNiv1)+strtoint(FNiv2), strtoint(FNiv3), tcTrim, true);
  end;
  //
  // Famille
  //
  if Famille<>'' then
  begin
    Debut := copy (Famille, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Famille, st, 27, 6, tcTrim, true)
    else begin
      Debut := copy (Famille, 1, 7);
      if Debut = 'GA2_FAM' then longueur := 3
      else longueur := 6 ;
      Tob_ArticleCompl.PutFixedStValue  (Famille, st, 27, longueur, tcTrim, true);
    end;
  end;

  //
  // Modèle
  //
  if Modele<>'' then
  begin
    Debut := copy (Modele, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Modele, st, 33, 6, tcTrim, true)
    else begin
      Debut := copy (Modele, 1, 7);
      if Debut = 'GA2_FAM' then longueur := 3
      else longueur := 6 ;
      Tob_ArticleCompl.PutFixedStValue  (Modele, st, 33, longueur, tcTrim, true);
    end;
  end;
  //
  // Matière
  //
  if Matiere<>'' then
  begin
    Debut := copy (Matiere, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Matiere, st, 39, 6, tcTrim, true)
    else begin
      Debut := copy (Matiere, 1, 7);
      if Debut = 'GA2_FAM' then longueur := 3
      else longueur := 6 ;
      Tob_ArticleCompl.PutFixedStValue (Matiere, st, 39, longueur, tcTrim, true);
    end;
  end;
  //
  // Coloris
  //
  if Coloris<>'' then
  begin
    Debut := copy (Coloris, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Coloris, st, 45, 5, tcTrim, true)
    else begin
      Debut := copy (Coloris, 1, 7);
      if Debut = 'GA2_FAM' then longueur := 3
      else longueur := 5 ;
      Tob_ArticleCompl.PutFixedStValue (Coloris, st, 45, longueur, tcTrim, true);
    end;
  end;
  //
  // Finition spéciale
  //
  if Finition<>'' then
  begin
    Debut := copy (Finition, 1, 3);
    if Debut = 'GA_' then  Tob_Article.PutFixedStValue(Finition, st, 50, 3, tcTrim, true)
    else begin
      Tob_ArticleCompl.PutFixedStValue(Finition, st, 50,3, tcTrim, true);
    end;
  end;
  //
  // Compositions
  //
  if Compo1 <>'' then
  begin
    Debut := copy (Compo1, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Compo1, st, 72, 2, tcTrim, true)
    else begin
      Tob_ArticleCompl.PutFixedStValue (Compo1, st, 72, 2, tcTrim, true);
    end;
  end;
  if Compo2 <>'' then
  begin
    Debut := copy (Compo1, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Compo2, st, 74, 2, tcTrim, true)
    else begin
      Tob_ArticleCompl.PutFixedStValue (Compo2, st, 74, 2, tcTrim, true);
    end;
  end;

  //
  // Grille de tailles
  //
  Tob_Article.PutFixedStValue('GA_GRILLEDIM'+CatDim[3], st, 53, 3, tcTrim, true);
  Tob_Article.PutFixedStValue('GA_DIMMASQUE'  , st, 53, 3, tcTrim, true);

  Tob_Article.PutFixedStValue('GA_LIBELLE'    , st, 56, 15, tcTrimL, false);
  // MODIF LM 11/09/02
  if Trim (Tob_Article.GetValue ('GA_LIBELLE'))='' then Tob_Article.PutValue ('GA_LIBELLE', Tob_Article.GetValue ('GA_CODEARTICLE'));

  Tob_Article.PutFixedStValue('GA_PRIXUNIQUE' , st, 71, 1, tcBooleanON, false);

  Tob_Article.PutFixedStValue('GA_DATECREATION', st, 87,   8, tcDate8AMJ, false);

  //
  // Récupération du PA
  //
  Montant := StrToFloat (copy (st, 107, 12));
  Montant := Montant/100 ;

  if DeviseSiteCentral = V_PGI.DevisePivot then
  begin
    Tob_Article.PutValue ('GA_PAHT', Montant);
    Tob_Article.PutValue ('GA_PMAP', Montant);
    Tob_Article.PutValue ('GA_DPA' , Montant);
  end else
  begin
    DateConv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    ToxConvertir (Montant, MontantDev, MontantOld, DeviseSiteCentral, V_PGI.DevisePivot, DateConv, nil) ;
    Tob_Article.PutValue ('GA_PAHT', MontantDev);
    Tob_Article.PutValue ('GA_PMAP', MontantDev);
    Tob_Article.PutValue ('GA_DPA' , MontantDev);
  end;

  if TVA<>'' then Tob_Article.PutFixedStValue('GA_FAMILLETAXE1', TVA, 1, 3, tcTrim, true);
end;

// Reprise ARFC2
Procedure TFRecupGB2000.RecupArticle2 (st : string);
var TPST     : TOB       ;
    DateConv : TDateTime ;
    Design   : string    ;
    Chtmp    : String    ;
    Debut    : string    ;
    PxTTC    : double    ;
    PxHT     : double    ;
    PxTTCDev : double    ;
    PxTTCOld : double    ;
    Cpt      : Integer   ;
begin
  //
  // Récupération du PV TTC
  //
  PxTTC := StrToFloat (copy (st, 19, 12));
  PxTTC := PxTTC/100 ;

  if DeviseSiteCentral = V_PGI.DevisePivot then Tob_Article.PutValue('GA_PVTTC', PxTTC)
  else begin
    Dateconv := NowH ;
    PxTTCDev := 0 ;  // initialisation bidon
    PxTTCOld := 0 ;  // initialisation bidon
    ToxConvertir (PxTTC, PxTTCDev, PxTTCOld, DeviseSiteCentral, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
    Tob_Article.PutValue('GA_PVTTC', PxTTCDev);
  end;
  //
  // Calcul du PV HT à partir du PV TTC
  //
  PxTTC := Tob_Article.GetValue  ('GA_PVTTC');
  PxHT  := (PxTTC*100)/(100+TauxTVADefaut);
  Tob_Article.PutValue ('GA_PVHT', PxHT);
  //
  // Catégorie GB 2000
  //
  if Categorie <>'' then
  begin
    Debut := copy (Categorie, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Categorie, st, 31, 3, tcTrim, true)
    else begin
      Tob_ArticleCompl.PutFixedStValue (Categorie, st, 31, 3, tcTrim, true);
    end;
  end;
  //
  // Formule GB 2000
  //
  if Formule   <>'' then
  begin
    Debut := copy (Formule, 1, 3);
    if Debut = 'GA_' then Tob_Article.PutFixedStValue (Formule, st, 34, 3, tcTrim, true)
    else begin
      Tob_ArticleCompl.PutFixedStValue (Formule, st, 34, 3, tcTrim, true);
    end;
  end;
  //
  // Structure GB 2000
  //
  if (Structure<>'') and (Trim(Copy(st,49,10))<>'') then
  begin
    Chtmp:=Trim(Copy(st,49,10));
    for cpt:=1 to 10 do
    begin
      if (cpt=10) and (Structure='GA_LIBREARTA') then break
      else if Structure='GA_LIBREART'+InttoStr(cpt) then break ;
    end ;
    if cpt<11 then
    begin
      Chtmp:=Trim(Copy(st,49,6));
      TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
      //TPST.initValeurs;
      if cpt=10 then TPST.PutValue('YX_TYPE', 'LAA')
      else TPST.PutValue('YX_TYPE', 'LA'+InttoStr(Cpt));
      TPST.PutValue('YX_CODE'   , Chtmp);
      TPST.PutValue('YX_LIBELLE', Chtmp);
      TPST.PutValue('YX_ABREGE' , Chtmp);
      ImportTablesLibres (TPST);
      TPST.Free;
    end ;
    Tob_Article.PutValue(Structure, Chtmp);
  end ;
  //
  // Type d'article GB 2000
  //
  if TypeArt<>'' then Tob_Article.PutFixedStValue(TypeArt, st, 37,  1, tcTrim, true);
  //
  // Libellé spécifique de l'article
  //
  if LibCol <>'' then Tob_Article.PutFixedStValue(LibCol , st, 60, 25, tcTrim, false);

  Tob_Article.PutFixedStValue   ('GA_CODEDOUANIER', st, 85, 15, tcTrim, false);
  Tob_Article.PutFixedStValue   ('GA_PAYSORIGINE' , st, 100, 3, tcTrim, false);
  Tob_Article.PutFixedStValue   ('GA_QUALIFPOIDS' , st, 103, 3, tcTrim, false);
  Tob_Article.PutFixedStValue   ('GA_FERME'       , st, 106, 1, tcBooleanON, false);
  Tob_Article.PutFixedStValue   ('GA_NUMEROSERIE' , st, 110, 1, tcBooleanON, false);
  Design := Tob_Article.GetValue('GA_LIBELLE') + copy (st, 111, 10) ;
  Tob_Article.PutValue          ('GA_LIBELLE', trim (Design));
end;

// Reprise ARFC3
Procedure TFRecupGB2000.RecupArticle3 (st : string);
var DateConv   : TDateTime;
    Montant    : double ;
    MontantDev : double ;
    MontantOld : double ;
begin
  //
  // Prix de revient de l'article
  //
  Montant := StrToFloat (copy (st, 19, 12));
  Montant := Montant/100 ;

  if DeviseSiteCentral = V_PGI.DevisePivot then
  begin
    Tob_Article.PutValue('GA_PRHT', Montant);
    Tob_Article.PutValue('GA_PMRP', Montant);
    Tob_Article.PutValue('GA_DPR' , Montant);
  end else
  begin
    Dateconv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    ToxConvertir (Montant, MontantDev, MontantOld, DeviseSiteCentral, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;

    Tob_Article.PutValue ('GA_PRHT'   , MontantDev);
    Tob_Article.PutValue ('GA_PMRP'   , MontantDev);
    Tob_Article.PutValue ('GA_DPR'    , MontantDev);
  end;
  //
  // Référence de l'article chez le fournisseurs
  //
  Tob_Article.PutFixedStValue('GA_COMMENTAIRE' , st, 34, 25, tcTrim, false);
end;

// Reprise ARFCA et ARFCB
Procedure TFRecupGB2000.RecupArticleAB (st : string);
var chp, typchp, val : string ;
    Debut            : string ;
    longueur, i      : integer;
begin
  chp:='';
  for i := 0 to 17 do
  begin
    if length (st) > (12 + (i * 6)) then
    begin
      case i of
        0  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATA01')
             else chp := Tob_Param.GetValue ('REPSTATA19');
        1  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATA02')
             else chp := Tob_Param.GetValue ('REPSTATA20');
        2  : chp := Tob_Param.GetValue ('REPSTATA03') ;
        3  : chp := Tob_Param.GetValue ('REPSTATA04') ;
        4  : chp := Tob_Param.GetValue ('REPSTATA05') ;
        5  : chp := Tob_Param.GetValue ('REPSTATA06') ;
        6  : chp := Tob_Param.GetValue ('REPSTATA07') ;
        7  : chp := Tob_Param.GetValue ('REPSTATA08') ;
        8  : chp := Tob_Param.GetValue ('REPSTATA09') ;
        9  : chp := Tob_Param.GetValue ('REPSTATA10') ;
        10 : chp := Tob_Param.GetValue ('REPSTATA11') ;
        11 : chp := Tob_Param.GetValue ('REPSTATA12') ;
        12 : chp := Tob_Param.GetValue ('REPSTATA13') ;
        13 : chp := Tob_Param.GetValue ('REPSTATA14') ;
        14 : chp := Tob_Param.GetValue ('REPSTATA15') ;
        15 : chp := Tob_Param.GetValue ('REPSTATA16') ;
        16 : chp := Tob_Param.GetValue ('REPSTATA17');
        17 : chp := Tob_Param.GetValue ('REPSTATA18');
      end;

      if chp<>'' then
      begin
        val    := copy(st, 19 + (i * 6), 6);
        val    := trim(val);
        typchp := ChampToType (chp);

        Debut := copy (Chp, 1, 3);
        if Debut = 'GA_' then longueur := 6          // Stat articles
        else                  longueur := 3 ;        // Fam articles

        if typchp = 'DOUBLE'  then
        begin
          if longueur = 6 then Tob_Article.PutValue(chp, valeur (val))
          else                 Tob_ArticleCompl.PutFixedStValue (Chp, st, 1, 3, tcDouble, true);
        end
        else if typchp = 'DATE'    then
        begin
          if longueur = 6 then Tob_Article.PutValue(chp, Str6ToDate(val, 90));
        end
        else if typchp = 'BOOLEAN' then
        begin
          if longueur = 6 then
          begin
            if val = 'O' then Tob_Article.PutValue(chp,'X') else Tob_Article.PutValue(chp,'-') ;
          end else
          begin
             if val = 'O' then Tob_ArticleCompl.PutValue(chp,'X') else Tob_ArticleCompl.PutValue(chp,'-') ;
          end;
        end
        else begin
          if Longueur = 6 then Tob_Article.PutValue(chp, val)
          else            Tob_ArticleCompl.PutValue(chp, copy (val, 1, 3));
        end;
      end ;
    end;
  end;
end;


///////////////////////////////////////////////////////////////////////////////////////
// Importation de la TOB Mere des articles par tailles
////////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportMereArtTaille;
begin
  Tob_Mere_Art_Taille.SetAllModifie (True);
  Tob_Mere_Art_Taille.UpdateDB (FALSE);            // Mise à jour de l'enregistrement table
  Tob_Mere_Art_Taille.Free;
  Tob_Mere_Art_Taille:=nil;
end ;

// Reprise ARTC1
Procedure TFRecupGB2000.RecupArtTaille (st : string);
var codeart            : string ;
    rang               : string ;
    CategorieDim       : string ;
    CodeDim1, CodeDim2 : string ;
    CodeDim3, CodeDim4 : string ;
    CodeDim5 		   : string ;
    Taille             : integer;
    codeBarre          : string ;
    PA                 : double ;
    PV                 : double ;
    PR                 : double ;
    ST_PA, ST_PV, ST_PR: string ;
    ST1                : string ; 
begin
  //////////////////////////////////////////////////////////////////////////////////////
  // Création de la TOB Mère si inexistante
  //////////////////////////////////////////////////////////////////////////////////////
  //if (Tob_Mere_Art_Taille = nil) then Tob_Mere_Art_Taille := TOB.CREATE ('Les articles tailles', NIL, -1);
  //////////////////////////////////////////////////////////////////////////////////////
  // Récupération du paramétrage
  //////////////////////////////////////////////////////////////////////////////////////
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

   Taille := StrToInt (copy (st, 19 ,2));
   rang   := CalculTaillePGI (Taille);

  CodeDim1     := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
  CodeDim4     := '   '; CodeDim5 := '   ';
  if (CategorieDim [3] = '1') then Codedim1 :=rang
  else if (CategorieDim [3] = '2') then Codedim2 :=rang
  else if (CategorieDim [3] = '3') then Codedim3 :=rang
  else if (CategorieDim [3] = '4') then Codedim4 :=rang
  else if (CategorieDim [3] = '5') then Codedim5 :=rang;
  /////////////////////////////////////////////////////////////////////////////////////////
  // Référence de l'article dimensionné
  // La fiche article par taille doit forcément existée dans PGI suite à la récup de l'ARF
  // Si ce n'est pas le cas, on laisse tomber.
  /////////////////////////////////////////////////////////////////////////////////////////
  Codeart := Copy (st, 7, 12) + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X';

  CodeBarre := Trim (Copy (st, 58, 13));
  PA        := StrToFloat (copy (st, 22, 12));  PA := PA/100 ;
  PV        := StrToFloat (copy (st, 34, 12));  PV := PV/100 ;
  PR        := StrToFloat (copy (st, 71, 12));  PR := PR/100 ;

  if (PA <> 0.0) then ST_PA := ',GA_DPA='+StrfPoint(PA)+', GA_PAHT='+StrfPoint(PA)+' ' else ST_PA := '' ;
  if (PV <> 0.0) then ST_PV := ',GA_PVTTC='+StrfPoint(PV)+' ' else ST_PV := '' ;
  if (PR <> 0.0) then ST_PR := ',GA_DPR='+StrfPoint(PR)+', GA_PRHT='+StrfPoint(PR)+' ' else ST_PR := '' ;

  St1 := 'UPDATE ARTICLE SET GA_CODEBARRE="'+CodeBarre+'" ' + ST_PA + ST_PV + ST_PR + ' WHERE GA_ARTICLE="'+CodeArt+'"' ;
  //ExecuteSQL('UPDATE ARTICLE SET GA_CODEBARRE="'+CodeBarre+'" WHERE GA_ARTICLE="'+CodeArt+'"') ;
  //ExecuteSQL('UPDATE ARTICLE SET GA_CODEBARRE="'+CodeBarre+'" ' + ST_PA + ST_PV + ST_PR + ' WHERE GA_ARTICLE="'+CodeArt+'"') ;
  ExecuteSQL(ST1);
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des contacts
///////////////////////////////////////////////////////////////////////////////////////
procedure ImportContact ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des RIB
///////////////////////////////////////////////////////////////////////////////////////
procedure ImportRib ( T : TOB );       // Mettre des controles de cohérence
begin
  T.SetAllModifie(TRUE) ;
  T.InsertOrUpdateDB (FALSE);            // Mise à jour de l'enregistrement table
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des tiers (utilisée pour les fournisseurs)
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportTiers (VerifTom : boolean) ;       // Mettre des controles de cohérence
var TomTiers : TOM_TIERS;
begin
  /////////////////////////////////////////////////////////
  // Vérification du client par la TOM
  // Attention : TOM TIERS modifié pour bon fonctionnement
  /////////////////////////////////////////////////////////
  if VerifTom then
  begin
{$IFDEF AGL560}
    TomTiers:=TOM_TIERS(CreateTOM('TIERS',Nil,False,False)) ;
{$ELSE}
    TomTiers:=TOM_TIERS.Create(Nil,Nil,FALSE) ;
{$ENDIF}

    if TomTiers.VerifTOB(Tob_Tiers) then
    begin
      Tob_Tiers.SetAllModifie(TRUE) ;
      Tob_Tiers.InsertOrUpdateDB (FALSE);
      CommitTrans;
    end else EcrireInfo ('    ATTENTION PROBLEME TOM -> ' + TomTiers.LastErrorMsg, True, True, True);

    TomTiers.Free ;
  end else
  begin
    Tob_Tiers.SetAllModifie(TRUE) ;
    Tob_Tiers.InsertOrUpdateDB (FALSE);
  end;

  Tob_Tiers.Free;
  Tob_Tiers:=nil;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des tiers (utilisée pour les clients) : Gestion d'une TOB Mère
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportMereTiers ;       // Mettre des controles de cohérence
var TomTiers  : TOM_TIERS;
    TobRejet  : TOB      ;
    TobClient : TOB      ;
    Tobcontact: TOB	 ;
    CptClient : integer  ;

begin
  //GcVoirTob (Tob_Mere_Tiers);
  TobRejet:=nil;
  /////////////////////////////////////////////////////////
	// Vérification du client par la TOM
  // Attention : TOM TIERS modifié pour bon fonctionnement
  /////////////////////////////////////////////////////////
{$IFDEF AGL560}
  TomTiers:=TOM_TIERS(CreateTOM('TIERS',Nil,False,False)) ;
{$ELSE}
  TomTiers:=TOM_TIERS.Create(Nil,Nil,FALSE) ;
{$ENDIF}
  for CptClient:=0 to Tob_Mere_Tiers.Detail.Count-1 do
  begin
    TobClient:=Tob_Mere_Tiers.Detail[CptClient];
    if TomTiers.VerifTOB(TobClient)=False then
    begin
      if TobRejet=nil then TOB.Create ('Clients Rejetés', nil, -1);
      TobClient.ChangeParent (TobRejet, -1);
      EcrireInfo ('    ATTENTION PROBLEME TOM -> ' + TomTiers.LastErrorMsg + ' pour le client ' + TobClient.GetValue('T_AUXILIAIRE'), True, True, True);
    end else
    begin
      Tobcontact:=TOB.Create('CONTACT', NIL, -1);  // Création de la TOB des contacts
      Tobcontact.PutValue   ('C_TYPECONTACT'   , 'T') ;
      Tobcontact.PutValue   ('C_AUXILIAIRE'    , TobClient.GetValue('T_AUXILIAIRE')) ;
      Tobcontact.PutValue   ('C_NUMEROCONTACT' ,'1') ;
      Tobcontact.PutValue   ('C_NATUREAUXI'    ,'CLI') ;
      Tobcontact.PutValue   ('C_CIVILITE'      , TobClient.GetValue('T_JURIDIQUE')) ;
      Tobcontact.PutValue   ('C_NOM'           , TobClient.GetValue('T_LIBELLE')) ;
      Tobcontact.PutValue   ('C_PRENOM'        , TobClient.GetValue('T_PRENOM')) ;
      Tobcontact.PutValue   ('C_TELEPHONE'     , TobClient.GetValue('T_TELEPHONE')) ;
      Tobcontact.PutValue   ('C_TELEX'         , TobClient.GetValue('T_TELEX')) ;
      Tobcontact.PutValue   ('C_FAX'           , TobClient.GetValue('T_FAX')) ;
      Tobcontact.PutValue   ('C_RVA'           , TobClient.GetValue('T_RVA')) ;
      Tobcontact.PutValue   ('C_JOURNAIS'      , TobClient.GetValue('T_JOURNAISSANCE')) ;
      Tobcontact.PutValue   ('C_MOISNAIS'      , TobClient.GetValue('T_MOISNAISSANCE')) ;
      Tobcontact.PutValue   ('C_ANNEENAIS'     , TobClient.GetValue('T_ANNEENAISSANCE')) ;
      Tobcontact.PutValue   ('C_SEXE'          , TobClient.GetValue('T_SEXE')) ;
      Tobcontact.PutValue   ('C_PRINCIPAL'     ,'X') ;
      Tobcontact.SetAllModifie(TRUE) ;
      Tobcontact.InsertOrUpdateDB (FALSE);
      Tobcontact.free;
    end;
  end;

  ///////////////////////////////////////////////////////////
  // MAJ de la DB
  ///////////////////////////////////////////////////////////
  //GcVoirTob (Tob_Mere_Tiers);
  Tob_Mere_Tiers.SetAllModifie(TRUE) ;
  Tob_Mere_Tiers.InsertOrUpdateDB (FALSE);
  CommitTrans;

  TomTiers.Free ;
  Tob_Mere_Tiers.Free;
  Tob_Mere_Tiers:=nil;
  if TobRejet<>nil then	TobRejet.Free;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des tiers complémentaire (utilisée pour les fournisseurs)
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportTiersCompl ;       // Mettre des controles de cohérence
begin
  Tob_Tiers_Compl.SetAllModifie(TRUE) ;
  Tob_Tiers_Compl.InsertOrUpdateDB (FALSE);
  Tob_Tiers_Compl.Free;
  Tob_Tiers_Compl:=nil;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des tiers (utilisée pour les clients) : Gestion d'une TOB Mère
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportMereTiersCompl ;       // Mettre des controles de cohérence
begin
  Tob_Mere_Tiers_Compl.SetAllModifie(TRUE) ;
  Tob_Mere_Tiers_Compl.InsertOrUpdateDB (FALSE);
  Tob_Mere_Tiers_Compl.Free;
  Tob_Mere_Tiers_Compl:=nil;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation des adresses (utilisée dans la récupération des fournisseurs)
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportAdresseTiers ;       // Mettre des controles de cohérence
begin
  Tob_Adresses.SetAllModifie(TRUE) ;
  Tob_Adresses.InsertOrUpdateDB (FALSE);
  Tob_Adresses.Free;
  Tob_Adresses:=Nil;
end ;



///////////////////////////////////////////////////////////////////////////////////////
// Création d'une ligne commentaire, contenant la référence de l'article générique
// pour la gestion spécifique MODE des lignes de documents
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.CreeLigneCommentaire ;
var prix           : double;
    GL_CodeArticle : string ;
begin
		////////////////////////////////////////////////////////////////
    // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    ////////////////////////////////////////////////////////////////
    Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
    //Tob_Ligne.initValeurs;
    ////////////////////////////////////////////////////////////////
    // Initialise les champs de la ligne à partir de l'entête
    ////////////////////////////////////////////////////////////////
    PieceVersLigne (Tob_Piece, Tob_Ligne);
    Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
    Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
    ////////////////////////////////////////////////////////////////
    // Ajoute des champs supplémentaires pour le calcul des cumuls
    ////////////////////////////////////////////////////////////////
    AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
    //////////////////////////////////////////////
    // Numéro de ligne
    //////////////////////////////////////////////
    Tob_Ligne.PutValue ('GL_NUMLIGNE' , NumLigneComment);
    Tob_Ligne.PutValue ('GL_NUMORDRE' , NumLigneComment);

    Tob_Ligne.PutValue ('GL_LIBELLE'  , Designation);
    Tob_Ligne.PutValue ('GL_TYPELIGNE', 'COM');		// Ligne commentaire

    Tob_Ligne.PutValue ('GL_QTESTOCK' , TotQte);   // Total des quantités
    Tob_Ligne.PutValue ('GL_QTEFACT'  , TotQte);
    Tob_Ligne.PutValue ('GL_QTERESTE' , TotQteReste);
    ////////////////////////////////////////////
    // Prix unitaire
    ///////////////////////////////////////////
    if (TotQte<>0) then Prix:=Arrondi(TotPrix/TotQte,DEV.Decimale) else Prix:=0.0;
    if Tob_Ligne.getValue    ('GL_FACTUREHT')='X' then Tob_Ligne.PutValue('GL_PUHTDEV',Prix)
    else Tob_Ligne.PutValue  ('GL_PUTTCDEV'   ,Prix) ;
    
    if Tob_Ligne.getValue    ('GL_FACTUREHT')='X' then Tob_Ligne.PutValue('GL_MONTANTHTDEV',ToTPrix)
    else Tob_Ligne.PutValue  ('GL_MONTANTTTCDEV', ToTPrix) ;

    Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);

    GL_CodeArticle := Trim (CodeArticle);
    Tob_Ligne.PutValue       ('GL_CODESDIM'    , GL_CodeArticle);
    Tob_Ligne.PutValue       ('GL_REFARTSAISIE', GL_CodeArticle);

    Tob_Ligne.PutValue       ('GL_FOURNISSEUR' , Tob_Piece.GetValue ('GP_TIERS'));

    Tob_Ligne.PutValue       ('GL_TYPEDIM'    , 'GEN');
    Tob_Ligne.PutValue       ('GL_QUALIFMVT'  , QualifMvt) ;
    Tob_Ligne.PutValue       ('GL_VALIDECOM'  , 'AFF') ;
    Tob_Ligne.PutValue       ('GL_MOTIFMVT'   , Motif_ES) ;

    if PieceAGenerer then
    begin
      if FactureSurCommande or FactureSurLivraison then
      begin
        if FactureSurCommande then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'C')
        else if FactureSurLivraison then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'L') ;

        Tob_Ligne.AddChampSupValeur('$$_CODEBTQ' , BtqNeg2000);
        Tob_Ligne.AddChampSupValeur('$$_NUMDOC'  , NumNeg2000);
      end;
    end;
end;


///////////////////////////////////////////////////////////////////////////////////////
// Imporation des réglements clients
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportReglement ;       // Mettre des controles de cohérence
begin
  //////////////////////////////////////////////////////
  // Création des TOB
  //////////////////////////////////////////////////////
  if (NombreErreurs (-1) = 0) then
  begin
    TOB_Acomptes:=TOB.Create('LES ACOMPTES',Nil,-1) ;
    //////////////////////////////////////////////////////////////////////////////////////////////////
    // Génération des PIEDECHE
    // PB 2 : Multi-devise
    //				Si Règlement en devise, il faut que :
    //					GPE_MONTANTECHE, GPE_MONTANTDEV doit être exprimé en francs (ou plutôt devise de tenue)
    //					GPE_DEVISE dans la devise de tenue du dossier
    //					GPE_DEVISEESP dans la devise étrangère
    //					GPE_MONTANTENCAIS en devise étrangère
    /////////////////////////////////////////////////////////////////////////////////////////////////////
    GereEcheancesGC (Tob_Piece, Tob_Tiers, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
    ///////////////////////////////////////////////////////////////
    // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES
    ///////////////////////////////////////////////////////////////
    //Tob_Piece.InsertOrUpdateDB (FALSE);
    //Tob_PiedBase.InsertOrUpdateDB (FALSE);
    Tob_PiedEche.SetAllModifie(TRUE) ;
    Tob_PiedEche.InsertOrUpdateDB (FALSE);
    //Tob_Adresses.InsertOrUpdateDB (FALSE);
  end;

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece   <> nil)	then Tob_Piece.free;	    // Libère la TOB du fournisseurs                                   // Libère TOB Pièce+Lignes
  if (Tob_Tiers   <> nil)	then Tob_Tiers.free;	    // Libère la TOB du fournisseurs
  if (Tob_PiedEche<> nil)	then Tob_PiedEche.free;   // Libère la TOB PiedEche
  if (Tob_Acomptes<> nil)	then Tob_Acomptes.free;   // Libère la TOB Acomptes

  Tob_Piece    := nil;
  Tob_Tiers    := nil;
  Tob_PiedEche := nil;
  Tob_Acomptes := nil;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Traitement des enregistrements CLI....
///////////////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupCLI (st : string);
var Q			: TQUERY ;
    SQL			: string ;
    chp, typchp, val    : string ;
    Codecli, RepEtaCpta : string ;
    RepCliProMail       : string ;
    Civilite		: string ;
    CodePaysGB          : string ;
    CodePaysPGI         : string ;
    CodeClientPGI       : string ;
    i, Reste            : integer;
begin
  Codecli       := Tob_Param.GetValue ('CODECLI');
  RepEtaCpta    := Tob_Param.GetValue ('REPETACPTA');
  RepCliProMail := Tob_Param.GetValue ('REPCLIPROMAIL');

  if ArretRecup then exit;
  if Copy(st,1,5)='CLIC1' then
  begin
    EtatComptable := '';
    //////////////////////////////////////////////////////////////////////////////////////
    // Création de la TOB Mère si inexistante
    //////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Mere_Tiers = nil) then Tob_Mere_Tiers := TOB.CREATE ('Les clients', NIL, -1);
    //////////////////////////////////////////////////////////////////////////////////////
    // MAJ du Compte rendu d'intégration
    //////////////////////////////////////////////////////////////////////////////////////
    EcrireInfo ('Client ' + Copy(st,10 ,6) + ' de la boutique ' + copy (st, 7, 3), False, True, False);
    Compteur := Compteur+1;
    Reste := Compteur Mod 10000;
    if (Reste = 0) then EcrireInfo (' Le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) + ' : ' + IntToStr (Compteur) + ' clients intégrés', True, False, False);
    //////////////////////////////////////////////////////////////////////////////////////
    // Création de la TOB Fille Tiers
    //////////////////////////////////////////////////////////////////////////////////////
    Tob_Tiers := TOB.CREATE ('TIERS', Tob_Mere_Tiers, -1);
    //Tob_Tiers:= TOB.CREATE ('TIERS', NIL, -1);
    //Tob_Tiers.initValeurs;
    if  CodeCli='0' then Chp:=copy(st,7,9)
    else if CodeCli='1' then chp:=copy(st,10,6);

    //////////////////////////////////////////////////////////////////////////////////////
    // Client particulier ou non ????
    //	1 - Si civilité à blanc 																		-> Client particulier
    //	2 - Si civilité renseignée, et civilité connue dans PGI 		-> Client particulier
    //      Si civilité renseignée, et civilité inconnue dans PGI   -> Client entreprise
    //																																 Si FJ connue dans PGI, on la reprend dans la fiche tiers
    //////////////////////////////////////////////////////////////////////////////////////
    if copy (st, 30, 3) <> '   ' then
    begin
      Civilite := copy (st, 30, 3);
      SQL:='Select * From CHOIXCOD WHERE CC_TYPE="CIV" AND CC_CODE="'+Civilite+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        Tob_Tiers.PutFixedStValue('T_PARTICULIER', 'O', 1, 1, tcBooleanON, false);
        Tob_Tiers.PutFixedStValue('T_JURIDIQUE'  , st, 30,  3, tcTrim, false);
        Ferme (Q);
      end else
      begin
	Ferme (Q);
        Tob_Tiers.PutFixedStValue('T_PARTICULIER', 'N', 1, 1, tcBooleanON, false);
        SQL:='Select * From COMMUN WHERE CO_TYPE="JUR" AND CO_CODE="'+Civilite+'"';
	Q:=OpenSQL(SQL,True) ;
  	if Not Q.EOF then Tob_Tiers.PutFixedStValue('T_JURIDIQUE'  , st, 30,  3, tcTrim, false);
        Ferme (Q);
        Tob_Tiers.PutValue ('T_FACTUREHT', 'X');
      end;
    end else Tob_Tiers.PutFixedStValue('T_PARTICULIER', 'O', 1, 1, tcBooleanON, false);

    Tob_Tiers.PutValue       ('T_AUXILIAIRE' , chp);
    Tob_Tiers.PutValue       ('T_TIERS'      , chp);
    Tob_Tiers.PutValue       ('T_FACTURE'    , chp);
    Tob_Tiers.PutValue       ('T_NATUREAUXI' , 'CLI');
    //Tob_Tiers.PutFixedStValue('T_DEPOT'      , st,  7,  3, tcTrim, false);
    //
    // Ajout LM 03/12/01 : Récupération du régime fiscal paramétré
    //
    Tob_Tiers.PutValue       ('T_REGIMETVA'  , Tob_Param.GetValue ('CLIEXPORT'));

    Tob_Tiers.PutFixedStValue('T_LIBELLE'    , st, 34, 32, tcTrim, false);
    if (Tob_Tiers.GetValue   ('T_LIBELLE') = '') then
    begin
      Tob_Tiers.PutValue('T_LIBELLE', Tob_Tiers.GetValue ('T_AUXILIAIRE'));
      EcrireInfo ('   Attention : Le nom du client ' + Tob_Tiers.GetValue ('T_AUXILIAIRE') + ' n''est pas renseigné', True, False, False);
    end;
    Tob_Tiers.PutFixedStValue('T_ABREGE'    , st,  34, 17, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_PRENOM'    , st,  66, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_CODEPOSTAL', st,  98,  7, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_TELEPHONE' , st, 105, 20, tcTrim, false);

    EtatComptable := copy (st, 126, 1);

    if RepEtaCpta<>'' then
    begin
       if (Tob_Mere_Tiers_Compl = nil) then  Tob_Mere_Tiers_Compl := TOB.CREATE ('Les Stats Clients', NIL, -1);

       Tob_Tiers_Compl:= TOB.CREATE ('TIERSCOMPL', Tob_Mere_Tiers_Compl, -1);
       //Tob_Tiers_Compl.initValeurs;

       if CodeCli='0' then Chp:=copy(st,7,9)
       else if CodeCli='1' then chp:=copy(st,10,6);

       Tob_Tiers_Compl.PutValue ('YTC_AUXILIAIRE', chp);
       Tob_Tiers_Compl.PutValue ('YTC_TIERS'     , chp);
       Tob_Tiers_Compl.PutFixedStValue(RepEtaCpta, st, 126, 1, tcTrim, true);
     end;

  end
  else if Copy(st,1,5)='CLIC2' then
  begin
    Tob_Tiers.PutFixedStValue('T_ADRESSE1'      , st,  16, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_ADRESSE2'      , st,  48, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_ADRESSE3'      , st,  80, 32, tcTrim, false);
    // Recherche du Pays PGI dans la table de correcpondance
    CodePaysGB := Trim (copy (st, 112, 3));
    CodePaysPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'T_PAYS', CodePaysGB, Tob_Pays);
    Tob_Tiers.PutValue('T_PAYS', CodePaysPGI);

    // Recherche du Pays PGI dans la table de correcpondance pour nationnalité
    CodePaysGB := Trim (copy (st, 115, 3));
    CodePaysPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'T_PAYS', CodePaysGB, Tob_Pays);
    Tob_Tiers.PutValue('T_NATIONALITE', CodePaysPGI);

    Tob_Tiers.PutFixedStValue('T_ANNEENAISSANCE', st, 118,  4, tcEntier, false);
    Tob_Tiers.PutFixedStValue('T_MOISNAISSANCE' , st, 122,  2, tcEntier, false);
    Tob_Tiers.PutFixedStValue('T_JOURNAISSANCE' , st, 124,  2, tcEntier, false);
  end
  else if Copy(st,1,5)='CLIC3' then
  begin
    Tob_Tiers.PutFixedStValue('T_VILLE'         , st, 16, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_DATECREATION'  , st, 50,  8, tcDate8AMJ, false);
    Tob_Tiers.PutFixedStValue('T_DATEMODIF'     , st, 58,  8, tcDate8AMJ, false);
    {if RepCliProMail<>'Booléen libre n° 1' then Tob_Tiers.PutFixedStValue('T_BOOLLIBRE1', st, 82, 1, tcBooleanON, True)
    else if RepCliProMail<>'Booléen libre n° 2' then Tob_Tiers.PutFixedStValue('T_BOOLLIBRE2', st, 82, 1, tcBooleanON, True)
    else if RepCliProMail<>'Booléen libre n° 3' then Tob_Tiers.PutFixedStValue('T_BOOLLIBRE3', st, 82, 1, tcBooleanON, True);}
    if copy(st,86,1)='P' then Tob_Tiers.PutValue('T_NATUREAUXI', 'PRO');
    Tob_Tiers.PutFixedStValue('T_TARIFTIERS'    , st,  87,  3, tcTrim, True);
    Tob_Tiers.PutFixedStValue('T_DEVISE'        , st,  90,  3, tcTrim, True);
    Tob_Tiers.PutFixedStValue('T_TELEX'         , st,  93, 20, tcTrim, false);
    Tob_Tiers.PutFixedStValue('T_LANGUE'        , st, 113,  3, tcTrim, false);
  end
  else if (Copy(st,1,5)='CLICA') or (Copy(st,1,5)='CLICB') then
  begin
    Tob_Tiers_Compl := nil;

    if (RepEtaCpta<>'') and (EtatComptable<>'') then
    begin
       if (Tob_Mere_Tiers_Compl = nil) then  Tob_Mere_Tiers_Compl := TOB.CREATE ('Les Stats Clients', NIL, -1);

       if Tob_Tiers_Compl=nil then
       begin
         Tob_Tiers_Compl:= TOB.CREATE ('TIERSCOMPL', Tob_Mere_Tiers_Compl, -1);
   	     //Tob_Tiers_Compl.initValeurs;
       end;

       if CodeCli='0' then Chp:=copy(st,7,9)
       else if CodeCli='1' then chp:=copy(st,10,6);

       Tob_Tiers_Compl.PutValue ('YTC_AUXILIAIRE', chp);
       Tob_Tiers_Compl.PutValue ('YTC_TIERS'     , chp);
       Tob_Tiers_Compl.PutFixedStValue(RepEtaCpta, st, 126, 1, tcTrim, true);
    end;

    for i := 0 to 17 do
    begin
      chp:='';
      if length (st) > (16 + (i * 6)) then
      begin
        case i of
            0  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATC01')
                else chp := Tob_Param.GetValue ('REPSTATC19');
            1  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATC02')
                else chp := Tob_Param.GetValue ('REPSTATC20');
            2  : chp := Tob_Param.GetValue ('REPSTATC03') ;
            3  : chp := Tob_Param.GetValue ('REPSTATC04') ;
            4  : chp := Tob_Param.GetValue ('REPSTATC05') ;
            5  : chp := Tob_Param.GetValue ('REPSTATC06') ;
            6  : chp := Tob_Param.GetValue ('REPSTATC07') ;
            7  : chp := Tob_Param.GetValue ('REPSTATC08') ;
            8  : chp := Tob_Param.GetValue ('REPSTATC09') ;
            9  : chp := Tob_Param.GetValue ('REPSTATC10') ;
            10 : chp := Tob_Param.GetValue ('REPSTATC11');
            11 : chp := Tob_Param.GetValue ('REPSTATC12');
            12 : chp := Tob_Param.GetValue ('REPSTATC13');
            13 : chp := Tob_Param.GetValue ('REPSTATC14');
            14 : chp := Tob_Param.GetValue ('REPSTATC15');
            15 : chp := Tob_Param.GetValue ('REPSTATC16');
            16 : chp := Tob_Param.GetValue ('REPSTATC17');
            17 : chp := Tob_Param.GetValue ('REPSTATC18');
        end;

        if chp<>'' then
        begin
      	  if (Tob_Mere_Tiers_Compl = nil) then  Tob_Mere_Tiers_Compl := TOB.CREATE ('Les Stats Clients', NIL, -1);

          if Tob_Tiers_Compl = nil then
          begin
            Tob_Tiers_Compl:= TOB.CREATE ('TIERSCOMPL', Tob_Mere_Tiers_Compl, -1);
   	        //Tob_Tiers_Compl.initValeurs;
          end;

      	  if CodeCli='0' then CodeClientPGI:=copy(st,7,9)
          else if CodeCli='1' then CodeClientPGI:=copy(st,10,6);

          Tob_Tiers_Compl.PutValue ('YTC_AUXILIAIRE', CodeClientPGI);
          Tob_Tiers_Compl.PutValue ('YTC_TIERS'     , CodeClientPGI);

          val := copy(st, 16 + (i * 6), 6);
          val := trim(val);

          typchp := ChampToType (chp);
          if typchp = 'DOUBLE'  then Tob_Tiers_Compl.PutValue(chp, valeur(val))
          else if typchp = 'DATE'    then Tob_Tiers_Compl.PutValue(chp, Str6ToDate(val, 90))
          else if typchp = 'BOOLEAN' then
          begin
            if val = 'O' then Tob_Tiers_Compl.PutValue(chp,'X')
            else Tob_Tiers_Compl.PutValue(chp,'-') ;
          end
          else Tob_Tiers_Compl.PutValue(chp, val);
        end;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Traitement des enregistrements FOU....
///////////////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupFOU (st : string);
var TPST             : TOB;
    chp, typchp, val : string;
    CodePaysGB       : string ;
    CodePaysPGI      : string ;
    CodeDevGB        : string ;
    CodeDevPGI       : string ;
    CodeModGB        : string ;
    CodeModPGI       : string ;
    SQL              : string ;
    Q                : TQuery ;
    i, NumeroContact : integer;
    numero           : integer;
begin

  if ArretRecup then exit;
  if Copy(st,1,5)='FOUC1' then
  begin
    if (Tob_Tiers<>Nil)        then ImportTiers (True) ;
    if (Tob_Adresses<>Nil)     then ImportAdresseTiers ;      // Adresse de facturation
    if (Tob_Tiers_Compl<>Nil)  then ImportTiersCompl ;

    if ArretRecup then exit;
    EcrireInfo ('Fournisseur ' + Copy(st,7 ,5), False, True, False);

    Tob_Tiers:= TOB.CREATE ('TIERS', NIL, -1);
    //Tob_Tiers.initValeurs;

    Tob_Tiers.PutFixedStValue ('T_AUXILIAIRE', st, 7, 5, TcTrim, True);
    Tob_Tiers.PutFixedStValue ('T_TIERS', st, 7, 5, TcTrim, True);
    Tob_Tiers.PutValue        ('T_NATUREAUXI', 'FOU');
    Tob_Tiers.PutFixedStValue ('T_LIBELLE', st, 19, 25, tcTrim, false);

    Tob_Tiers.PutFixedStValue ('T_ABREGE', st, 12, 7, tcTrim, false);
    Tob_Tiers.PutFixedStValue ('T_ADRESSE1', st, 44, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue ('T_ADRESSE2', st, 76, 32, tcTrim, false);
    Tob_Tiers.PutFixedStValue ('T_CODEPOSTAL', st, 108, 7, tcTrim, false);
    ///////////////////////////////////////////////////////////
    // Valeurs par défaut
    //////////////////////////////////////////////////////////
    Tob_Tiers.PutValue        ('T_COLLECTIF'  , VH^.DefautFou);
    Tob_Tiers.PutValue        ('T_REGIMETVA'  , Tob_Param.GetValue ('FRNREGIMETVA'));
    Tob_Tiers.PutValue        ('T_COMPTATIERS', Tob_Param.GetValue ('FRNFAMCOMPTA'));
    Tob_Tiers.PutValue        ('T_FACTUREHT'  , 'X');
  end
  else if Copy(st,1,5)='FOUC2' then
  begin
    Tob_Tiers.PutFixedStValue ('T_VILLE', st, 12, 32, tcTrim, false);
    // Recherche du Pays PGI dans la table de correcpondance
    CodePaysGB  := Trim (copy (st, 44, 3));
    CodePaysPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'T_PAYS', CodePaysGB, Tob_Pays);
    Tob_Tiers.PutValue ('T_PAYS', CodePaysPGI);
    //////////////////////////////////////////////////////////////
    // Ajout du nom du responsable dans les CONTACT
    //////////////////////////////////////////////////////////////
    NumeroContact:=0;
    if trim(copy(st,47,20))<>'' then
    begin
    	NumeroContact:=NumeroContact+1;
      TPST := TOB.CREATE ('CONTACT', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue        ('C_TYPECONTACT'  , 'T');
      TPST.PutFixedStValue ('C_AUXILIAIRE'   , st, 7, 5, tcTrim, false);
      TPST.PutValue        ('C_NUMEROCONTACT', NumeroContact);
      TPST.PutValue        ('C_NATUREAUXI'   , 'FOU');
      TPST.PutFixedStValue ('C_PRINCIPAL'    , 'O', 1, 1, tcBooleanON, false);
      TPST.PutFixedStValue ('C_NOM'          , st, 47, 20, tcChaine, false);
      TPST.PutFixedStValue ('C_TELEPHONE'    , st, 87, 20, tcChaine, false);
      TPST.PutFixedStValue ('C_TELEX'        , st, 107, 20, tcChaine, false);
      ImportContact (TPST);
      TPST.Free;
   end ;
   //////////////////////////////////////////////////////////////
   // Ajout du nom du représentant dans les CONTACT
   //////////////////////////////////////////////////////////////
   if trim(copy(st,67,20))<>'' then
    begin
      NumeroContact:=NumeroContact+1;
      TPST := TOB.CREATE ('CONTACT', NIL, -1);
      //TPST.initValeurs;
      TPST.PutValue        ('C_TYPECONTACT'  , 'TIE');
      TPST.PutFixedStValue ('C_AUXILIAIRE'   , st, 7, 5, tcTrim, false);
      TPST.PutValue        ('C_NUMEROCONTACT', NumeroContact);
      TPST.PutValue        ('C_NATUREAUXI'   , 'FOU');
      if (NumeroContact=1) then TPST.PutFixedStValue ('C_PRINCIPAL'    , 'O', 1, 1, tcBooleanON, false)
      else TPST.PutFixedStValue ('C_PRINCIPAL'    , 'N', 1, 1, tcBooleanON, false);
      TPST.PutFixedStValue ('C_NOM'          , st, 67, 20, tcChaine, false);
      TPST.PutFixedStValue ('C_TELEPHONE'    , st, 87, 20, tcChaine, false);
      TPST.PutFixedStValue ('C_TELEX'        , st, 107, 20, tcChaine, false);
      ImportContact (TPST);
      TPST.Free;
   end ;
   Tob_Tiers.PutFixedStValue('T_TELEPHONE', st, 87, 20, tcTrim, false);
   Tob_Tiers.PutFixedStValue('T_RVA', st, 107, 20, tcTrim, false);
  end
  else if Copy(st,1,5)='FOUC3' then
  begin
    Tob_Tiers.PutFixedStValue('T_FAX', st, 12, 20, tcTrim, false);
    // Recherche de la devise PGI dans la table de correcpondance
    CodeDevGB  := Trim (copy (st, 49, 3));
    CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'T_DEVISE', CodeDevGB, Tob_Devise);
    Tob_Tiers.PutValue ('T_DEVISE', CodeDevPGI);
    //Tob_Tiers.PutFixedStValue('T_DEVISE', st, 49, 3, tcTrim, True);

    // Recherche du mode de règlement PGI dans la table de correcpondance
    CodeModGB  := Trim (copy (st, 52, 3));
    CodeModPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'T_MODEREGLE', CodeModGB, Tob_Reglement);
    Tob_Tiers.PutValue ('T_MODEREGLE', CodeModPGI);
    //Tob_Tiers.PutFixedStValue('T_MODEREGLE', st, 52, 3, tcTrim, True);
    // ???????????????????
    //if trim(copy(st,61,10))<>'' then Tob_Tiers.PutFixedStValue('T_AUXILIAIRE', st, 61, 10, tcTrim, false);

    ///////////////////////////////////////////////////////////
    // Adresse de facturation
    ///////////////////////////////////////////////////////////
    if trim(copy(st,71,32))<>'' then
    begin
      Tob_Adresses := TOB.CREATE ('ADRESSES', NIL, -1);
      //Tob_Adresses.InitValeurs;
      // Recherche du numéro interne dispo
      numero := 0;
      SQL:='Select * From ADRESSES ORDER BY ADR_NUMEROADRESSE';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        while not Q.EOF do
        begin
          if (numero < Q.FindField('ADR_NUMEROADRESSE').AsInteger) then
            numero := Q.FindField('ADR_NUMEROADRESSE').AsInteger;
          Q.Next;
        end;
      end;
      Ferme(Q);
      numero:=numero+1;
      Tob_Adresses.PutValue        ('ADR_NUMEROADRESSE', numero);
      Tob_Adresses.PutFixedStValue ('ADR_REFCODE'      , st, 7, 5, tcTrim, True);
      Tob_Adresses.PutValue        ('ADR_TYPEADRESSE'  , 'TIE');
      Tob_Adresses.PutFixedStValue ('ADR_LIBELLE'      , st, 71, 32, tcTrim, True);
      Tob_Adresses.PutFixedStValue ('ADR_CODEPOSTAL'   , st, 103, 5, tcTrim, True);
    end;
  end
  else if Copy(st,1,5)='FOUC4' then
  begin
    if (Tob_Adresses<>Nil)  then
    begin
      Tob_Adresses.PutFixedStValue('ADR_ADRESSE1', st, 12, 32, tcTrim, True);
      Tob_Adresses.PutFixedStValue('ADR_ADRESSE2', st, 44, 32, tcTrim, True);
      Tob_Adresses.PutFixedStValue('ADR_VILLE'   , st, 76, 32, tcTrim, True);
      Tob_Adresses.PutFixedStValue('ADR_PAYS'    , st, 108, 3, tcTrim, True);
    end;
  end
  else if Copy(st,1,5)='FOUC5' then
  begin
    if trim(copy(st,12,25))<>'' then
    begin
      TPST := TOB.CREATE ('RIB', NIL, -1);
      //TPST.initValeurs;
      TPST.PutFixedStValue('R_AUXILIAIRE'    , st, 7, 5, tcTrim, false);
      TPST.PutValue       ('R_NUMERORIB'     , '1');
      TPST.PutFixedStValue('R_PRINCIPAL'     , 'O', 1, 1, tcBooleanON, false);
      TPST.PutFixedStValue('R_ETABBQ'        , st, 37, 5, tcTrim, false);
      TPST.PutFixedStValue('R_GUICHET'       , st, 42, 5, tcTrim, false);
      TPST.PutFixedStValue('R_NUMEROCOMPTE'  , st, 47, 11, tcTrim, false);
      TPST.PutFixedStValue('R_CLERIB'        , st, 58, 2, tcTrim, false);
      TPST.PutFixedStValue('R_DOMICILIATION' , st, 12, 25, tcTrim, false);
      TPST.PutValue       ('R_DEVISE'        , Tob_Tiers.GetValue ('T_DEVISE')) ;
      TPST.PutValue       ('R_VILLE'         , Tob_Tiers.GetValue ('T_VILLE')) ;
      TPST.PutValue       ('R_PAYS'          , Tob_Tiers.GetValue ('T_PAYS')) ;
      ImportRIB (TPST);
      TPST.Free;
    end;
    Tob_Tiers.PutFixedStValue('T_REMISE', st, 99, 12, tcDouble100, false);
    if (Copy(st,98,1)='O') then Tob_Tiers.PutFixedStValue('T_SOUMISTPF', 'N', 1, 1, tcBooleanON, false)
    else Tob_Tiers.PutFixedStValue('T_SOUMISTPF', 'O', 1, 1, tcBooleanON, false);
    Tob_Tiers.PutFixedStValue('T_NIF', st, 76, 17, tcTrim, false);
  end
  else if Copy(st,1,5)='FOUC6' then
  begin
    Tob_Tiers.PutFixedStValue('T_ESCOMPTE', st, 38, 12, tcDouble100, false);
  end
  else if (Copy(st,1,5)='FOUCA') or (Copy(st,1,5)='FOUCB') then
  begin
    if (Copy(st,1,5)='FOUCA') then
    begin
         Tob_Tiers_Compl:= TOB.CREATE ('TIERSCOMPL', NIL, -1);
         //Tob_Tiers_Compl.initValeurs;
         Tob_Tiers_Compl.PutFixedStValue('YTC_AUXILIAIRE', st, 7, 5, TcTrim, True);
         Tob_Tiers_Compl.PutFixedStValue('YTC_TIERS', st, 7, 5, TcTrim, True);
    end;

    for i := 0 to 17 do
    begin
      chp:='';
      if length (st) > (12 + (i * 6)) then
      begin
        case i of
            0  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATF01')
                 else chp := Tob_Param.GetValue ('REPSTATF19');
            1  : if st[5] = 'A' then chp := Tob_Param.GetValue ('REPSTATF02')
                 else chp := Tob_Param.GetValue ('REPSTATF20');
            2  : chp := Tob_Param.GetValue ('REPSTATF03') ;
            3  : chp := Tob_Param.GetValue ('REPSTATF04') ;
            4  : chp := Tob_Param.GetValue ('REPSTATF05') ;
            5  : chp := Tob_Param.GetValue ('REPSTATF06') ;
            6  : chp := Tob_Param.GetValue ('REPSTATF07') ;
            7  : chp := Tob_Param.GetValue ('REPSTATF08') ;
            8  : chp := Tob_Param.GetValue ('REPSTATF09') ;
            9  : chp := Tob_Param.GetValue ('REPSTATF10') ;
            10 : chp := Tob_Param.GetValue ('REPSTATF11') ;
            11 : chp := Tob_Param.GetValue ('REPSTATF12') ;
            12 : chp := Tob_Param.GetValue ('REPSTATF13') ;
            13 : chp := Tob_Param.GetValue ('REPSTATF14') ;
            14 : chp := Tob_Param.GetValue ('REPSTATF15') ;
            15 : chp := Tob_Param.GetValue ('REPSTATF16') ;
            16 : chp := Tob_Param.GetValue ('REPSTATF17') ;
            17 : chp := Tob_Param.GetValue ('REPSTATF18') ;
      end;
      if chp<>'' then
      begin
        val := copy(st, 12 + (i * 6), 6);
        val := trim(val);
        typchp := ChampToType (chp);
        if typchp = 'DOUBLE'  then Tob_Tiers_Compl.PutValue(chp, valeur(val)) else
        if typchp = 'DATE'    then Tob_Tiers_Compl.PutValue(chp, Str6ToDate(val, 90)) else
        if typchp = 'BOOLEAN' then
        begin
          if val = 'O' then Tob_Tiers_Compl.PutValue(chp,'X') else Tob_Tiers_Compl.PutValue(chp,'-') ;
          end else Tob_Tiers_Compl.PutValue(chp, val);
        end ;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Traitement des enregistrements STD....   : Stock GB 2000
// Attention : Reste à faire les conversions des montants en contre-valeur.
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportDispo ;       // Mettre des controles de cohérence
var  Tob_Dispo       : TOB    ;
     QteRec , QteTrf : double ;
     QteVte          : double ;
     QteES           : double ;
     Ecart           : double ;
     CptDispo	     : integer;
begin
  //////////////////////////////////////////////////////////////////////
  // Recalcul du stock physique
  /////////////////////////////////////////////////////////////////////
  if Tob_Mere_Dispo <> nil then
  begin
    for CptDispo:=0 to Tob_Mere_Dispo.Detail.Count-1 do
    begin
      Tob_Dispo:=Tob_Mere_Dispo.Detail[CptDispo];

      QteRec  := Tob_Dispo.GetValue ('GQ_LIVREFOU');      // Qté réceptionnée fournisseur
      QteTrf  := Tob_Dispo.GetValue ('GQ_TRANSFERT');     // Qté transférée
      QteES   := Tob_Dispo.GetValue ('GQ_ENTREESORTIES'); // Qté E/S
      QteVte  := Tob_Dispo.GetValue ('GQ_VENTEFFO');      // Qté vendue
      Ecart   := Tob_Dispo.GetValue ('GQ_STOCKINITIAL');  // Ecart inventaire stocké momentanément dans GQ_STOCKINITIAL
      Tob_Dispo.PutValue            ('GQ_STOCKINITIAL', 0);

      if Tob_Param.GetValue ('STOCKCLOTURE')='X' then
      begin
        Tob_Dispo.PutValue ('GQ_PHYSIQUE'   , QteRec+QteTrf+QteES+Ecart-QteVte);
        Tob_Dispo.PutValue ('GQ_CLOTURE'    , 'X');
        Tob_Dispo.putValue ('GQ_DATECLOTURE', DateCloture);
      end else
      begin
        Tob_Dispo.PutValue ('GQ_PHYSIQUE', QteRec+QteTrf+QteES+Ecart-QteVte);
      end;
    end;

    Tob_Mere_Dispo.SetAllModifie (True);
    Tob_Mere_Dispo.InsertOrUpdateDB (FALSE);
    Tob_Mere_Dispo.Free;
    Tob_Mere_Dispo:=nil;
  end;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Récupération du dernier chrono des tarifs
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupMaxTarif ;
var QMax    : TQUERY ;
    CodeMax : string ;
begin
  MaxTarif := 0;

  QMax := OpenSQL ('SELECT MAX(GF_TARIF) FROM TARIF',True);
  if Not QMax.EOF then
  begin
    CodeMax:=QMax.Fields[0].AsString;
    if CodeMax <> '' then MaxTarif := StrToInt(CodeMax)+1;
  end;
  Ferme(QMax) ;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Chargement en TOB des qualifiants tarifs
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ChargeTobQualifiantTarif(var Tob_Tarif : TOB; TypeTarif : string; TypePeriode : string);
var Tob_TypeTar    : TOB    ;
    Tob_Periode    : TOB    ;
    CodeMax        : string ;
    Q		   : TQUERY ;
    QMax           : TQUERY ;
    CodeMaxInt     : integer;
    Trouve_Periode : boolean;
    Trouve_TypeTar : boolean;
begin

  if (TypeTarif <> '') and (TypePeriode <> '') then
  begin
    Trouve_Periode := False ;
    Trouve_TypeTar := False ;

    Q:=OpenSql('Select * from TARIFMODE where GFM_TYPETARIF="'+TypeTarif+'" AND GFM_PERTARIF="'+TypePeriode+'"',True) ;
    if not Q.EOF then
    begin
      Tob_Tarif.SelectDB('',Q);
      Ferme (Q) ;
    end else
    begin
      //
      // L'enregistrement dans la table TARIFMODE n'existe pas, il faut le créer
      ///
      Ferme (Q) ;
      //
      // Chargement en TOB de la période de tarif
      //
      Tob_Periode := nil ;
      Tob_TypeTar := nil ;
      Q:=OpenSql('Select * from TARIFPER where GFP_CODEPERIODE="'+TypePeriode+'"',True) ;
      if not Q.EOF then
      begin
  	    Tob_Periode := TOB.Create('TARIFPER',NIL,-1);
        //Tob_Periode.initValeurs;
        Tob_Periode.SelectDB('',Q);
        Trouve_Periode := True ;
      end;
      Ferme (Q);
      //
      // Chargement en TOB du type de tarif
      //
      Q:=OpenSql('Select * from TARIFTYPMODE where GFT_CODETYPE="'+TypeTarif+'"',True) ;
      if not Q.EOF then
      begin
  	   Tob_TypeTar := TOB.Create('TARIFPER',NIL,-1);
        //Tob_TypeTar.initValeurs;
     	if Tob_TypeTar.SelectDB('', Q) then Tob_TypeTar.LoadDB (FALSE);
        Trouve_TypeTar := True ;
      end;
      Ferme (Q);
      //
      // Chargement en TOB et création de l'enregistrement dans la table TARIFMODE
      if (trouve_Periode) and (Trouve_TypeTar) then
      begin
        CodeMaxInt := 0 ;
        QMax := OpenSQL ('SELECT MAX(GFM_TARFMODE) FROM TARIFMODE',True);
        if Not QMax.EOF then
        begin
          CodeMax:=QMax.Fields[0].AsString;
          if CodeMax <> '' then CodeMaxInt := StrToInt(CodeMax)+1;
        end ;
        Ferme(QMax) ;

        Tob_Tarif.PutValue('GFM_TARFMODE'    , CodeMaxInt);
        Tob_Tarif.PutValue('GFM_TYPETARIF'   , Tob_TypeTar.GetValue('GFT_CODETYPE'));
        Tob_Tarif.PutValue('GFM_PERTARIF'    , Tob_Periode.GetValue('GFP_CODEPERIODE'));
        Tob_Tarif.PutValue('GFM_LIBELLE'     , Tob_TypeTar.GetValue('GFT_LIBELLE') + '-' + Tob_Periode.GetValue('GFP_LIBELLE'));
        Tob_Tarif.PutValue('GFM_DATEDEBUT'   , Tob_Periode.GetValue('GFP_DATEDEBUT'));
        Tob_Tarif.PutValue('GFM_DATEFIN'     , Tob_Periode.GetValue('GFP_DATEFIN'));
        Tob_Tarif.PutValue('GFM_DEMARQUE'    , Tob_Periode.GetValue('GFP_DEMARQUE'));
        Tob_Tarif.PutValue('GFM_ARRONDI'     , Tob_Periode.GetValue('GFP_ARRONDI'));
        Tob_Tarif.PutValue('GFM_TARIFBASE'   , Tob_Periode.GetValue('GFP_TARIFBASE'));
        Tob_Tarif.PutValue('GFM_COEF'        , Tob_TypeTar.GetValue('GFT_COEF'));
        Tob_Tarif.PutValue('GFM_DEVISE'      , Tob_TypeTar.GetValue('GFT_DEVISE'));
        Tob_Tarif.PutValue('GFM_NATURETYPE'  , Tob_TypeTar.GetValue('GFT_NATURETYPE'));
        // MAJ DB
        Tob_Tarif.InsertDB (nil, False);
        // Libération des TOB
        Tob_TypeTar.free;
        Tob_Periode.free;
      end;
    end;
  end;
end ;

procedure CalculPriorite(var Tob_Tarif : TOB) ;
var Priorite : Integer ;
begin
Priorite:=0 ;
if Tob_Tarif.GetValue ('GF_ARTICLE')      <> ''    then Priorite:=Priorite+40 else
if Tob_Tarif.GetValue ('GF_TARIFARTICLE') <> ''    then Priorite:=Priorite+10 ;
if Tob_Tarif.GetValue ('GF_TIERS')        <> ''    then Priorite:=Priorite+40 ;
if Tob_Tarif.GetValue ('GF_TARIFTIERS')   <> ''    then Priorite:=Priorite+5 ;
if Tob_Tarif.GetValue ('GF_DEVISE')       <> ''    then Priorite:=Priorite+2 ;
if Tob_Tarif.GetValue ('GF_DEPOT')        <> ''    then Priorite:=Priorite+1 ;
if Tob_Tarif.GetValue ('GF_CASCADEREMISE') = 'FOR' then Priorite:=Priorite+200 ;
if Tob_Tarif.GetValue ('GF_CASCADEREMISE') = 'CAS' then Priorite:=Priorite+100 ;
if Tob_Tarif.GetValue ('GF_FERME')         = 'X'   then Priorite:=0 ;

Tob_Tarif.PutValue    ('GF_PRIORITE', Priorite) ;
end ;

{=======================Conversion prix contrevaleur ====================}
function ConvertPrix (Prix : Double) : Double;
BEGIN
if VH^.TenueEuro then Result:=EuroToFranc(Prix) else Result:=FrancToEuro(Prix) ;
END;

///////////////////////////////////////////////////////////////////////////////////////
// Traitement des enregistrements STD....   : Tarifs GB 2000
// Problématique : Il faut créer pour chaque article générique une ligne tarif
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportTarif ;       // Mettre des controles de cohérence
var cpt            : integer;
    Tob_fille      : TOB    ;
    Tob_New        : TOB    ;
    TarifMode      : integer;
    TarifOld       : integer;
    CodeArticle    : string ;
    CodeArticleGen : string ;
    CodeArticleTai : string ;
begin
  CodeArticle := '';
  TarifOld    := -1;

  for cpt:=0 to Tob_Mere_Tarif.detail.count-1 do
  begin
    Tob_Fille      := Tob_Mere_Tarif.detail [cpt];
    CodeArticleTai := Tob_Fille.GetValue ('GF_ARTICLE');
    CodeArticleGen := copy (CodeArticleTai, 1, 18) + '               X';
    TarifMode      := Tob_Fille.GetValue ('GF_TARFMODE');

    if (TarifMode<>TarifOld) or (CodeArticleGen <> CodeArticle) then
    begin
      CodeArticle := CodeArticleGen;
      TarifOld    := TarifMode     ;

      Tob_New := Tob.create ('TARIF', Tob_Mere_Tarif, -1);
      Tob_New.Dupliquer (Tob_Fille, True, True, True);

      MaxTarif:=MaxTarif+1;
      Tob_New.PutValue ('GF_TARIF'  , MaxTarif);
      Tob_New.putValue ('GF_ARTICLE', CodeArticleGen);
    end;
  end;

  Tob_Mere_Tarif.SetAllModifie    (TRUE) ;
  Tob_Mere_Tarif.InsertOrUpdateDB (FALSE);
  Tob_Mere_Tarif.Free;
  Tob_Mere_Tarif := nil ;
end ;

//
// Recopie d'un prix d'achat dans la fiche article/
//
procedure TFRecupGB2000.MajPrixAchatArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeArticleGen : string    ;
    requete        : string    ;
    Montant        : double    ;
begin
  Montant    := PrixArticle ;
  //
  // Dernier prix d'achat
  //
  if Montant <> 0.0 then
  begin
    //
    // Si devise différente de la devise dossier -> Conversion
    //
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;
    //
    // MAJ de l'article dimensionné
    //
    ExecuteSQL ('UPDATE ARTICLE SET GA_DPA="'+StrFPoint(Montant)+'" , GA_PAHT="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');
    //
    // L'article générique a t il un prix d'achat ?
    //
    CodeArticleGen := Copy(CodeArticle, 1, 18) + '               X';
    if CodeArticleGen <> DernierArticlePA then
    begin
      Requete := 'UPDATE ARTICLE SET GA_DPA="'+StrFPoint(Montant)+'", GA_PAHT="'+StrFPoint(Montant)+'"' ;
      Requete := Requete + ' WHERE GA_ARTICLE="'+CodeArticleGen+'"';
      ExecuteSQL (Requete);
    end;
    DernierArticlePA := CodeArticleGen ;
  end;
end;

//
// Recopie d'un prix d'achat dans la fiche article/
//
procedure TFRecupGB2000.MajPrixRevientArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeArticleGen : string    ;
    requete        : string    ;
    Montant        : double    ;
begin
  Montant        := PrixArticle ;
  //
  // Dernier prix d'achat
  //
  if Montant <> 0.0 then
  begin
    //
    // Si devise différente de la devise dossier -> Conversion
    //
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;
    //
    // MAJ de l'article dimensionné
    //
    ExecuteSQL ('UPDATE ARTICLE SET GA_DPR="'+StrFPoint(Montant)+'" , GA_PRHT="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');
    //
    // L'article générique a t il un prix d'achat ?
    //
    CodeArticleGen := Copy(CodeArticle, 1, 18) + '               X';
    if CodeArticleGen <> DernierArticlePR then
    begin
      Requete := 'UPDATE ARTICLE SET GA_DPR="'+StrFPoint(Montant)+'", GA_PRHT="'+StrFPoint(Montant)+'"' ;
      Requete := Requete + ' WHERE GA_ARTICLE="'+CodeArticleGen+'"';
      ExecuteSQL (Requete);
    end;
    DernierArticlePR := CodeArticleGen ;
  end;
end;


//
// Recopie d'un prix d'achat dans la fiche article/
//
procedure TFRecupGB2000.MajPrixVenteArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeArticleGen : string    ;
    requete        : string    ;
    Montant        : double    ;
begin
  Montant        := PrixArticle ;
  //
  // Dernier prix d'achat
  //
  if Montant <> 0.0 then
  begin
    //
    // Si devise différente de la devise dossier -> Conversion
    //
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;
    //
    // MAJ de l'article dimensionné
    //
    ExecuteSQL ('UPDATE ARTICLE SET GA_PVTTC="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');
    //
    // L'article générique a t il un prix d'achat ?
    //
    CodeArticleGen := Copy(CodeArticle, 1, 18) + '               X';
    if CodeArticleGen <> DernierArticlePV then
    begin
      Requete := 'UPDATE ARTICLE SET GA_PVTTC="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticleGen+'" ' ;
      ExecuteSQL (Requete);
    end;
    DernierArticlePV := CodeArticleGen ;
  end;
end;

//
// Recopie d'un prix d'achat dans la fiche article/
//
procedure TFRecupGB2000.MajPrixVenteHTArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeArticleGen : string    ;
    requete        : string    ;
    Montant        : double    ;
begin
  Montant        := PrixArticle ;
  //
  // Dernier prix d'achat
  //
  if Montant <> 0.0 then
  begin
    //
    // Si devise différente de la devise dossier -> Conversion
    //
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;
    //
    // MAJ de l'article dimensionné
    //
    ExecuteSQL ('UPDATE ARTICLE SET GA_PVHT="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');
    //
    // L'article générique a t il un prix d'achat ?
    //
    CodeArticleGen := Copy(CodeArticle, 1, 18) + '               X';
    if CodeArticleGen <> DernierArticlePVHT then
    begin
      Requete := 'UPDATE ARTICLE SET GA_PVHT="'+StrFPoint(Montant)+'"' ;
      Requete := Requete + ' WHERE GA_ARTICLE="'+CodeArticleGen+'"';
      ExecuteSQL (Requete);
    end;
    DernierArticlePVHT := CodeArticleGen ;
  end;
end;

// Traitement des enregistrements STD....
Procedure TFRecupGB2000.RecupSTD (st : string);
var Tob_Tarif		 : TOB   ;
    rang                 : string;
    CategorieDim         : string;
    CodeDim1, CodeDim2   : string;
    CodeDim3, CodeDim4   : string;
    CodeDim5             : string;
    SQL			 : string;
    Q                    : TQUERY;
    qte1, qte2           : double;
    QteEntree, QteSortie : double;
    Montant              : double;
    MontantDev           : double;
    MontantOld           : double;
    EcartInventaire      : double;
    Taille               : integer;
    DateConv             : TDateTime;
begin

  if ArretRecup then exit;

  if Copy(st,1,5)='STDC1' then
  begin
    //
    // Initialise le nombre d'erreur
    //
    NombreErreurs (0);

    CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

    Taille := StrToInt (copy (st, 22 ,2));
    rang   := CalculTaillePGI (Taille);

    CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
    CodeDim4 := '   '; CodeDim5 := '   ';

    if (CategorieDim [3] = '1') then Codedim1 :=rang
    else if (CategorieDim [3] = '2') then Codedim2 := rang
    else if (CategorieDim [3] = '3') then Codedim3 := rang
    else if (CategorieDim [3] = '4') then Codedim4 := rang
    else if (CategorieDim [3] = '5') then Codedim5 := rang;

    CodeArtTarif := Copy (st, 10, 12) + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

    SQL:='Select GA_FOURNPRINC From ARTICLE WHERE GA_ARTICLE="'+CodeArtTarif+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo ('   Erreur -> Problème pour retrouver l''article ' + CodeArtTarif, True, True, True);
      NombreErreurs (1);
    end;
    Ferme(Q) ;
  end;

  if NombreErreurs (-1) = 0 then
  begin
    if Copy(st,1,5)='STDC1' then
    begin
      if ArretRecup then exit;

      //EcrireInfo ('Fiche Stock de l''article ' + Copy(st,10 ,12) + ' dans la boutique ' + copy (st, 7, 3), False, True, False);

      ////////////////////////////////////////////////////////////////////////////////////
      // Initialisation
      ////////////////////////////////////////////////////////////////////////////////////
      DeviseTarif     := copy (st, 122, 3);
      QteVendue       := 0;
      QteRecue        := 0;
      /////////////////////////////////////////////////////////////////////////////////////
      // CAS 1 : Reprise des stocks
      ////////////////////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('VALSTOCK')='X' then
      begin
	    if (Tob_Mere_Dispo = nil) then Tob_Mere_Dispo := TOB.CREATE ('Les stocks', NIL, -1);

    	Tob_Dispo:= TOB.CREATE ('DISPO', Tob_Mere_Dispo, -1);
    	//Tob_Dispo.initValeurs;

    	Tob_Dispo.PutValue       ('GQ_ARTICLE'   , CodeArtTarif);
    	Tob_Dispo.PutFixedStValue('GQ_DEPOT'     , st     , 7, 3, TcTrim, True);
      	// Stock minimum
	    Tob_Dispo.PutFixedStValue('GQ_STOCKMIN'  , st     , 24, 5, tcDouble, false);
      	// Quantité commandée PGI = Encours de commande : Cde initiale GB + Cde réassort GB - quantité reçue
    	Qte1      := StrToFloat (copy (st, 29, 5));  // Qte commandée initiale
    	Qte2      := StrToFloat (copy (st, 34, 5));  // Qte commandée réassort
        QteRecue  := StrtoFloat (copy (st, 39,5));
    	Tob_Dispo.PutValue      ('GQ_RESERVEFOU', Qte1+Qte2-QteRecue);
      	// Quantité transférée PGI = Qté transféré reçue GB - Qté transféré émise GB
    	Qte1 := StrToFloat (copy (st, 44, 5));  // Qte transférée reçue
    	Qte2 := StrToFloat (copy (st, 49, 5));  // Qte transférée émise
    	Tob_Dispo.PutValue       ('GQ_TRANSFERT', Qte1-Qte2);
        // Sauvegarde de la quantité vendue et reçue
        QteVendue := StrtoFloat (copy (st, 54,5));
      	QteRecue  := StrtoFloat (copy (st, 39,5));

      	// Prix d'achat
        //
        Montant := StrToFloat (copy (st, 71, 12));
        Montant := Montant/100 ;

        if DeviseTarif = V_PGI.DevisePivot then
        begin
          Tob_Dispo.PutValue('GQ_DPA' , Montant);
          Tob_Dispo.PutValue('GQ_PMAP', Montant);
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DeviseTarif, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Dispo.PutValue('GQ_DPA'    , MontantDev);
          Tob_Dispo.PutValue('GQ_PMAP'   , MontantDev);
        end;
      end;

      /////////////////////////////////////////////////////////////////////////////////////
      // CAS 2 : Reprise des tarifs
      ////////////////////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('TARIF')='X' then
      begin
      	if (Tob_Mere_Tarif = nil) then Tob_Mere_Tarif := TOB.CREATE ('Les tarifs', NIL, -1);

        //////////////////////////////////////////////////////////////////////
        // Reprise Tarif PV TTC = Prix de vente GB 2000
        //////////////////////////////////////////////////////////////////////
        if Tob_TarifMode_TTC <> nil then
        begin
          if Tob_TarifMode_TTC.GetValue ('GFM_DEVISE') = DeviseTarif then
          begin
   	        Tob_Tarif:= TOB.CREATE ('TARIF', Tob_Mere_Tarif, -1);
    	    //Tob_Tarif.initValeurs;

    	    MaxTarif:=MaxTarif+1;
      	    Tob_Tarif.PutValue       ('GF_TARIF'        , MaxTarif);
      	    Tob_Tarif.PutValue       ('GF_ARTICLE'      , CodeArtTarif);
            //Tob_Tarif.PutFixedStValue('GF_DEPOT'        , st, 7, 3, tcTrim, false);
      	    //Tob_Tarif.PutFixedStValue('GF_SOCIETE'      , st, 7, 3, tcTrim, false);
	        Tob_Tarif.PutFixedStValue('GF_PRIXUNITAIRE' , st, 83, 12, tcDouble100, false);
            // Initialisation à partir de la TOB TARIFMODE
            Tob_Tarif.PutValue       ('GF_TARFMODE'     , Tob_TarifMode_TTC.GetValue ('GFM_TARFMODE'));
            Tob_Tarif.PutValue       ('GF_LIBELLE'      , Tob_TarifMode_TTC.GetValue ('GFM_LIBELLE'));
            Tob_Tarif.PutValue       ('GF_DATEDEBUT'    , Tob_TarifMode_TTC.GetValue ('GFM_DATEDEBUT'));
            Tob_Tarif.PutValue       ('GF_DATEFIN'      , Tob_TarifMode_TTC.GetValue ('GFM_DATEFIN'));
            Tob_Tarif.PutValue       ('GF_DEVISE'       , Tob_TarifMode_TTC.GetValue ('GFM_DEVISE'));
            Tob_Tarif.PutValue       ('GF_DEMARQUE'     , Tob_TarifMode_TTC.GetValue ('GFM_DEMARQUE'));
            Tob_Tarif.PutValue       ('GF_ARRONDI'      , Tob_TarifMode_TTC.GetValue ('GFM_ARRONDI'));
	        // Initialisation en "dur"
  	        Tob_Tarif.PutValue       ('GF_QUALIFPRIX'   , 'GRP');
    	    Tob_Tarif.PutValue       ('GF_CASCADEREMISE', 'MIE');
      	    Tob_Tarif.PutValue       ('GF_MODECREATION' , 'MAN');
      	    Tob_Tarif.PutValue       ('GF_REGIMEPRIX'   , 'TTC' );
            Tob_Tarif.PutValue       ('GF_NATUREAUXI'   , 'CLI' );
            Tob_Tarif.PutValue       ('GF_BORNEINF'     , -999999);
            Tob_Tarif.PutValue       ('GF_BORNESUP'     , 999999);
            Tob_Tarif.PutValue       ('GF_QUANTITATIF'  , '-');
            // Calcul de la priorité
            CalculPriorite(Tob_Tarif) ;
          end;
        end;

        //////////////////////////////////////////////////////////////////////
        // Reprise Tarif PV SOLDES = Prix de soldes GB 2000
        //////////////////////////////////////////////////////////////////////
        if Tob_TarifMode_Sol <> nil then
        begin
          if Tob_TarifMode_Sol.GetValue ('GFM_DEVISE') = DeviseTarif then
          begin
   	        Tob_Tarif:= TOB.CREATE ('TARIF', Tob_Mere_Tarif, -1);
    	    //Tob_Tarif.initValeurs;

    	    MaxTarif:=MaxTarif+1;
      	    Tob_Tarif.PutValue       ('GF_TARIF'        , MaxTarif);
      	    Tob_Tarif.PutValue       ('GF_ARTICLE'      , CodeArtTarif);
            Tob_Tarif.PutFixedStValue('GF_DEPOT'        , st, 7, 3, tcTrim, false);
      	    Tob_Tarif.PutFixedStValue('GF_SOCIETE'      , st, 7, 3, tcTrim, false);
	        Tob_Tarif.PutFixedStValue('GF_PRIXUNITAIRE' , st, 107, 12, tcDouble100, false);
            // Initialisation à partir de la TOB TARIFMODE
            Tob_Tarif.PutValue       ('GF_TARFMODE'     , Tob_TarifMode_Sol.GetValue ('GFM_TARFMODE'));
            Tob_Tarif.PutValue       ('GF_LIBELLE'      , Tob_TarifMode_Sol.GetValue ('GFM_LIBELLE'));
            Tob_Tarif.PutValue       ('GF_DATEDEBUT'    , Tob_TarifMode_Sol.GetValue ('GFM_DATEDEBUT'));
            Tob_Tarif.PutValue       ('GF_DATEFIN'      , Tob_TarifMode_Sol.GetValue ('GFM_DATEFIN'));
            Tob_Tarif.PutValue       ('GF_DEVISE'       , Tob_TarifMode_Sol.GetValue ('GFM_DEVISE'));
            Tob_Tarif.PutValue       ('GF_DEMARQUE'     , Tob_TarifMode_Sol.GetValue ('GFM_DEMARQUE'));
            Tob_Tarif.PutValue       ('GF_COEFFICIENT'  , Tob_TarifMode_Sol.GetValue ('GFM_COEFFICIENT'));
            Tob_Tarif.PutValue       ('GF_ARRONDI'      , Tob_TarifMode_Sol.GetValue ('GFM_ARRONDI'));
	        // Initialisation en "dur"
  	        Tob_Tarif.PutValue       ('GF_QUALIFPRIX'   , 'GRP');
    	    Tob_Tarif.PutValue       ('GF_CASCADEREMISE', 'MIE');
      	    Tob_Tarif.PutValue       ('GF_MODECREATION' , 'MAN');
      	    Tob_Tarif.PutValue       ('GF_REGIMEPRIX'   , 'TTC' );
            Tob_Tarif.PutValue       ('GF_NATUREAUXI'   , 'CLI' );
            Tob_Tarif.PutValue       ('GF_BORNEINF'     , -999999);
            Tob_Tarif.PutValue       ('GF_BORNESUP'     , 999999);
            Tob_Tarif.PutValue       ('GF_QUANTITATIF'  , '-');
            // Calcul de la priorité
            CalculPriorite(Tob_Tarif) ;
          end;
        end;
      end;

      if Tob_Param.GetValue ('TARIFACH')='X' then
      begin
      	if (Tob_Mere_Tarif = nil) then Tob_Mere_Tarif := TOB.CREATE ('Les tarifs', NIL, -1);
        //////////////////////////////////////////////////////////////////////
        // Reprise Tarif Achat HT = Prix d'achat GB 2000
        //////////////////////////////////////////////////////////////////////
        if Tob_TarifMode_Ach <> nil then
        begin
          if Tob_TarifMode_Ach.GetValue ('GFM_DEVISE') = DeviseTarif then
          begin
   	        Tob_Tarif:= TOB.CREATE ('TARIF', Tob_Mere_Tarif, -1);
    	    //Tob_Tarif.initValeurs;

    	    MaxTarif:=MaxTarif+1;
      	    Tob_Tarif.PutValue       ('GF_TARIF'        , MaxTarif);
      	    Tob_Tarif.PutValue       ('GF_ARTICLE'      , CodeArtTarif);
            //Tob_Tarif.PutFixedStValue('GF_DEPOT'        , st, 7, 3, tcTrim, false);
      	    //Tob_Tarif.PutFixedStValue('GF_SOCIETE'      , st, 7, 3, tcTrim, false);
	        Tob_Tarif.PutFixedStValue('GF_PRIXUNITAIRE' , st, 71, 12, tcDouble100, false);
            // Initialisation à partir de la TOB TARIFMODE
            Tob_Tarif.PutValue       ('GF_TARFMODE'     , Tob_TarifMode_TTC.GetValue ('GFM_TARFMODE'));
            Tob_Tarif.PutValue       ('GF_LIBELLE'      , Tob_TarifMode_TTC.GetValue ('GFM_LIBELLE'));
            Tob_Tarif.PutValue       ('GF_DATEDEBUT'    , Tob_TarifMode_TTC.GetValue ('GFM_DATEDEBUT'));
            Tob_Tarif.PutValue       ('GF_DATEFIN'      , Tob_TarifMode_TTC.GetValue ('GFM_DATEFIN'));
            Tob_Tarif.PutValue       ('GF_DEVISE'       , Tob_TarifMode_TTC.GetValue ('GFM_DEVISE'));
            Tob_Tarif.PutValue       ('GF_DEMARQUE'     , Tob_TarifMode_TTC.GetValue ('GFM_DEMARQUE'));
            Tob_Tarif.PutValue       ('GF_ARRONDI'      , Tob_TarifMode_TTC.GetValue ('GFM_ARRONDI'));
	        // Initialisation en "dur"
  	        Tob_Tarif.PutValue       ('GF_QUALIFPRIX'   , 'GRP');
    	    Tob_Tarif.PutValue       ('GF_CASCADEREMISE', 'MIE');
      	    Tob_Tarif.PutValue       ('GF_MODECREATION' , 'MAN');
      	    Tob_Tarif.PutValue       ('GF_REGIMEPRIX'   , 'HT' );
            Tob_Tarif.PutValue       ('GF_NATUREAUXI'   , 'FOU' );
            Tob_Tarif.PutValue       ('GF_BORNEINF'     , -999999);
            Tob_Tarif.PutValue       ('GF_BORNESUP'     , 999999);
            Tob_Tarif.PutValue       ('GF_QUANTITATIF'  , '-');
            // Calcul de la priorité
            CalculPriorite(Tob_Tarif) ;
          end;
        end;
      end;
      //////////////////////////////////////////////////////////////////////
      // Report du PA dans la fiche article
      //////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('REPPAARTICLE')='X' then
      begin
        Montant := StrToFloat (copy (st, 71, 12));
        Montant := Montant/100 ;
        MajPrixAchatArticle (CodeArtTarif, Montant, DeviseTarif) ;
      end;
      //////////////////////////////////////////////////////////////////////
      // Report du PA dans la fiche article
      //////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('REPPVARTICLE')='X' then
      begin
        Montant := StrToFloat (copy (st, 83, 12));
        Montant := Montant/100 ;
        if Montant = 0.0 then
        begin
          Montant := StrToFloat (copy (st, 95, 12));
          Montant := Montant/100 ;
        end ;
        MajPrixVenteArticle (CodeArtTarif, Montant, DeviseTarif) ;
      end;
    end
    else if Copy(st,1,5)='STDC2' then
    begin
      ////////////////////////////////////////////////////////////////////////////////////
      // CAS 1 : Reprise des stocks
      ////////////////////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('VALSTOCK')='X' then
      begin
    	// Quantité livrée fournisseur PGI = Reçu GB - Retour FRN GB
    	Qte1 := StrToFloat (copy (st, 76, 5));  // Qte retour fournisseur
      	Tob_Dispo.PutValue ('GQ_LIVREFOU'  , QteRecue-Qte1);
      	// Quantité Vente FO PGI = Qté vendue GB - Qté Retour GB
    	Qte1 := StrToFloat (copy (st, 81, 5));  // Qte retour comptoir
    	Tob_Dispo.PutValue ('GQ_VENTEFFO' , QteVendue-Qte1);
     	// Quantité réservée client = Qté réservée client NEGOCE
    	//Tob_Dispo.PutFixedStValue('GQ_RESERVECLI', st, 71, 5, tcDouble, false);
      	// Cumul des entrées de stock = Entrées de stock GB
        QteEntree := StrToFloat (copy (st, 61, 5));      // Qte entrée
        QteSortie := StrToFloat (copy (st, 66, 5));      // Qte sortie
        Tob_Dispo.PutValue ('GQ_ENTREESORTIES' , QteEntree-QteSortie);
        // Sauvegarde de l'écart d'inventaire temporairement dans GQ_STOCKINITIAL pour calcul du stock physique.
        EcartInventaire := StrtoFloat (copy (st, 86,5));
        Tob_Dispo.PutValue ('GQ_STOCKINITIAL' , EcartInventaire);
      end;

      ////////////////////////////////////////////////////////////////////////////////////
      // CAS 2 : Reprise des tarifs NEGOCE 2000
      ////////////////////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('TARIFHT')='X' then
      begin
     	//////////////////////////////////////////////////////////////////////
    	// Reprise Tarif PV HT = Prix de vente NEGOCE 2000
        //////////////////////////////////////////////////////////////////////
	    Tob_Tarif:= TOB.CREATE ('TARIF', Tob_Mere_Tarif, -1);
     	//Tob_Tarif.initValeurs;

    	MaxTarif:=MaxTarif+1;
      	Tob_Tarif.PutValue       ('GF_TARIF'        , MaxTarif);
      	Tob_Tarif.PutValue       ('GF_ARTICLE'      , CodeArtTarif);
	    Tob_Tarif.PutFixedStValue('GF_PRIXUNITAIRE' , st, 49, 12, tcDouble100, false);
        Tob_Tarif.PutValue       ('GF_DEVISE'       , DeviseTarif);
    	//Tob_Tarif.PutFixedStValue('GF_DEPOT'        , st, 7, 3, tcTrim, false);
      	//Tob_Tarif.PutFixedStValue('GF_SOCIETE'      , st, 7, 3, tcTrim, false);
      	// Initialisation en "dur"
	    Tob_Tarif.PutValue       ('GF_QUALIFPRIX'   , 'GRP');
   	    Tob_Tarif.PutValue       ('GF_CASCADEREMISE', 'MIE');
     	Tob_Tarif.PutValue       ('GF_MODECREATION' , 'MAN');
     	Tob_Tarif.PutValue       ('GF_REGIMEPRIX'   , 'HT' );
        Tob_Tarif.PutValue       ('GF_BORNEINF'     , -999999);
        Tob_Tarif.PutValue       ('GF_BORNESUP'     , 999999);
        Tob_Tarif.PutValue       ('GF_QUANTITATIF'  , '-');

     	// Initialisation à partir du qualifiant tarif MODE
 	    Tob_Tarif.PutValue       ('GF_LIBELLE'      , 'Prix HT NEGOCE 2000');
        Tob_Tarif.PutFixedStValue('GF_DATEDEBUT'    , '01011900' , 1, 8, tcDate8JMA, false);
        Tob_Tarif.PutFixedStValue('GF_DATEFIN'      , '31122099' , 1, 8, tcDate8JMA, false);

        CalculPriorite(Tob_Tarif) ;
      end;
      //////////////////////////////////////////////////////////////////////
      // Report du PA dans la fiche article
      /////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('REPPVHTARTICLE')='X' then
      begin
        Montant := StrToFloat (copy (st, 49, 12));
        Montant := Montant/100 ;
        MajPrixVenteHTArticle (CodeArtTarif, Montant, DeviseTarif) ;
      end;
    end
    else if Copy(st,1,5)='STDC3' then
    begin
      ////////////////////////////////////////////////////////////////////////////////////
      // CAS 1 : Reprise des stocks
      ////////////////////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('VALSTOCK')='X' then
      begin
      	Montant := StrToFloat (copy (st, 24, 12));
        Montant := Montant/100 ;

        if DeviseTarif = V_PGI.DevisePivot then
        begin
          Tob_Dispo.PutValue('GQ_DPR' , Montant);
          Tob_Dispo.PutValue('GQ_PMRP', Montant);
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DeviseTarif, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Dispo.PutValue('GQ_DPR'    , MontantDev);
          Tob_Dispo.PutValue('GQ_PMRP'   , MontantDev);
        end;
      end;
      /////////////////////////////////////////////////////////////////////
      // Report du PR dans la fiche article
      /////////////////////////////////////////////////////////////////////
      if Tob_Param.GetValue ('REPPRARTICLE')='X' then
      begin
        Montant := StrToFloat (copy (st, 24, 12));
        Montant := Montant/100 ;
        MajPrixRevientArticle (CodeArtTarif, Montant, DeviseTarif) ;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
// Importation d'une pièce PGI
// 		-> Création des pieds de document
//		-> Recalcul de la pièce (entête, ligne et pied)
//		-> Import DB (piéce, lignes, pieds, adresses, ....
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportPiece ;       // Mettre des controles de cohérence
var i : integer ;
begin
	//////////////////////////////////////////////////////
  // Création des TOB
  //////////////////////////////////////////////////////
  if (NombreErreurs (-1) = 0) then
  begin
  	//////////////////////////////////////////////////////////////////
    // On force le recalcul de la pièce
 	//////////////////////////////////////////////////////////////////
    PutValueDetail (Tob_Piece,'GP_RECALCULER','X');

  	//////////////////////////////////////////////////////////////////
    // Création de la dernière ligne commentaire
    //////////////////////////////////////////////////////////////////
    if (LigneCommACreer = True) then
    begin
      LigneCommACreer:=False;
      CreeLigneCommentaire;
    end;

    //////////////////////////////////////////////////////////////////
    // MAJ des remises lignes
    //////////////////////////////////////////////////////////////////
    if Tob_Piece.GetValue ('GP_REMISEPIED') <> 0 then
    begin
      for i:=0 to Tob_Piece.detail.count-1 do
      begin
        Tob_Piece.detail [i].PutValue ('GL_REMISEPIED', Tob_Piece.GetValue ('GP_REMISEPIED'));
      end;
    end;

	Tob_PiedBase:= TOB.CREATE ('Les taxes', NIL, -1);
  	Tob_PiedEche:= TOB.CREATE ('Les règlements', NIL, -1);
    Tob_Acomptes:=TOB.Create('LES ACOMPTES',Nil,-1) ;
    //
    // MODIF LM 15/05/02 : On enlève les adresses
    //
  	//Tob_Adresses:= TOB.Create ('LES ADRESSES',Nil,-1) ;
    //TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Livraison}
	//TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Facturation}

  	//////////////////////////////////////////////////////////
    //
    // MODIF LM 15/05/02 : On enlève les adresses
    //
  	// Génération des adresses de livraison et de facturation
  	//////////////////////////////////////////////////////////
  	//ValideLesAdresses ( Tob_Piece, Tob_Piece, Tob_Adresses) ;
  	//////////////////////////////////////////////////////////
  	// Calcul des cumuls de la pièce
  	// Calcul de PIEDBASE
  	//////////////////////////////////////////////////////////
  	CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article, Tob_PiedPort,nil,nil, DEV) ;
  	//////////////////////////////////////////////////////////
  	// Génération des PIEDECHE
  	//////////////////////////////////////////////////////////
  	GereEcheancesGC (Tob_Piece, Tob_Tiers, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
  	//////////////////////////////////////////////////////////
  	// Sauvegarde Chrono
  	//////////////////////////////////////////////////////////
  	IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce
  	///////////////////////////////////////////////////////////////
  	// MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES
  	///////////////////////////////////////////////////////////////

 {   Tob_Piece.SetAllModifie(TRUE) ;
    Tob_PiedBase.SetAllModifie(TRUE) ;
    Tob_PiedEche.SetAllModifie(TRUE) ;
    //Tob_Adresses.SetAllModifie(TRUE) ;
    if Tob_PiedPort<>nil then Tob_PiedPort.SetAllModifie(TRUE);}

  	Tob_Piece.InsertDB(nil);
  	Tob_PiedBase.InsertDB(nil);
  	Tob_PiedEche.InsertDB(nil);
  	//Tob_Adresses.InsertorUpdateDB (False);
    //if Tob_PiedPort<>nil then Tob_PiedPort.InsertDBTable (nil);

  end;

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece   <> nil) then Tob_Piece.Free;      // Libère TOB Pièce+Lignes
  if (Tob_Article <> nil)	then Tob_Article.free;		// Libère la TOB des articles de la pièce
  if (Tob_Tiers   <> nil)	then Tob_Tiers.free;	    // Libère la TOB du fournisseurs
  if (Tob_PiedBase<> nil)	then Tob_PiedBase.free;   // Libère la TOB PiedBase
  if (Tob_PiedEche<> nil)	then Tob_PiedEche.free;   // Libère la TOB PiedEche
  if (Tob_Acomptes<> nil)	then Tob_Acomptes.free;   // Libère la TOB Acomptes
  if (Tob_PiedPort<> nil)	then Tob_PiedPort.free;   // Libère la TOB des ports
  //if (Tob_Adresses<> nil)	then Tob_Adresses.free;   // Libère la TOB Adresses

  Tob_Piece    := nil;
  Tob_Article  := nil;
  Tob_Tiers    := nil;
  Tob_PiedBase := nil;
  Tob_PiedEche := nil;
  Tob_Acomptes := nil;
  Tob_PiedPort := nil;
  Tob_Adresses := nil;
end ;

///////////////////////////////////////////////////////////////////////////////////////
// Importation d'une pièce PGI
// 		-> Création des pieds de document
//		-> Recalcul de la pièce (entête, ligne et pied)
//		-> Import DB (piéce, lignes, pieds, adresses, ....
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportESStock;       // Mettre des controles de cohérence
var i : integer ;
begin
	//////////////////////////////////////////////////////
  // Création des TOB
  //////////////////////////////////////////////////////
  if (NombreErreurs (-1) = 0) then
  begin
  	//////////////////////////////////////////////////////////////////
    // On force le recalcul de la pièce
 	//////////////////////////////////////////////////////////////////
    PutValueDetail (Tob_Piece,'GP_RECALCULER','X');

  	//////////////////////////////////////////////////////////////////
    // Création de la dernière ligne commentaire
    //////////////////////////////////////////////////////////////////
    if (LigneCommACreer = True) then
    begin
      LigneCommACreer:=False;
      CreeLigneCommentaire;
    end;

    //////////////////////////////////////////////////////////////////
    // MAJ des remises lignes
    //////////////////////////////////////////////////////////////////
    if Tob_Piece.GetValue ('GP_REMISEPIED') <> 0 then
    begin
      for i:=0 to Tob_Piece.detail.count-1 do
      begin
        Tob_Piece.detail [i].PutValue ('GL_REMISEPIED', Tob_Piece.GetValue ('GP_REMISEPIED'));
      end;
    end;

	Tob_PiedBase:= TOB.CREATE ('Les taxes', NIL, -1);
  	Tob_PiedEche:= TOB.CREATE ('Les règlements', NIL, -1);
    Tob_Acomptes:=TOB.Create('LES ACOMPTES',Nil,-1) ;
    //
    // MODIF LM 15/05/02 : On enlève les adresses
    //
  	//Tob_Adresses:= TOB.Create ('LES ADRESSES',Nil,-1) ;
    //TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Livraison}
	//TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Facturation}

  	//////////////////////////////////////////////////////////
    //
    // MODIF LM 15/05/02 : On enlève les adresses
    //
  	// Génération des adresses de livraison et de facturation
  	//////////////////////////////////////////////////////////
  	//ValideLesAdresses ( Tob_Piece, Tob_Piece, Tob_Adresses) ;
  	//////////////////////////////////////////////////////////
  	// Calcul des cumuls de la pièce
  	// Calcul de PIEDBASE
  	//////////////////////////////////////////////////////////
  	CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers_Defaut, Tob_Article, Tob_PiedPort,nil,nil, DEV) ;
  	//////////////////////////////////////////////////////////
  	// Génération des PIEDECHE
  	//////////////////////////////////////////////////////////
  	GereEcheancesGC (Tob_Piece, Tob_Tiers_Defaut, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
  	//////////////////////////////////////////////////////////
  	// Sauvegarde Chrono
  	//////////////////////////////////////////////////////////
  	IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce
  	///////////////////////////////////////////////////////////////
  	// MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES
  	///////////////////////////////////////////////////////////////

 {   Tob_Piece.SetAllModifie(TRUE) ;
    Tob_PiedBase.SetAllModifie(TRUE) ;
    Tob_PiedEche.SetAllModifie(TRUE) ;
    //Tob_Adresses.SetAllModifie(TRUE) ;
    if Tob_PiedPort<>nil then Tob_PiedPort.SetAllModifie(TRUE);}

  	Tob_Piece.InsertDB(nil);
  	Tob_PiedBase.InsertDB(nil);
  	Tob_PiedEche.InsertDB(nil);
  	//Tob_Adresses.InsertorUpdateDB (False);
    //if Tob_PiedPort<>nil then Tob_PiedPort.InsertDBTable (nil);

  end;

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece   <> nil) then Tob_Piece.Free;      // Libère TOB Pièce+Lignes
  if (Tob_Article <> nil)	then Tob_Article.free;		// Libère la TOB des articles de la pièce
  //if (Tob_Tiers   <> nil)	then Tob_Tiers.free;	    // Libère la TOB du fournisseurs
  if (Tob_PiedBase<> nil)	then Tob_PiedBase.free;   // Libère la TOB PiedBase
  if (Tob_PiedEche<> nil)	then Tob_PiedEche.free;   // Libère la TOB PiedEche
  if (Tob_Acomptes<> nil)	then Tob_Acomptes.free;   // Libère la TOB Acomptes
  //if (Tob_PiedPort<> nil)	then Tob_PiedPort.free;   // Libère la TOB des ports
  //if (Tob_Adresses<> nil)	then Tob_Adresses.free;   // Libère la TOB Adresses

  Tob_Piece    := nil;
  Tob_Article  := nil;
  Tob_Tiers    := nil;
  Tob_PiedBase := nil;
  Tob_PiedEche := nil;
  Tob_Acomptes := nil;
  Tob_PiedPort := nil;
  Tob_Adresses := nil;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI avec solde de la pièce précédente
Suite ........ :  1 - Chargement des différents documents précédents
Suite ........ :  2 - Calcul des reliquats
Suite ........ :  2 - Mise à jour des pièces d'origine, suivante et de la nouvelle
Suite ........ :  3 - Import DB de la pièce précédente et de la nouvelle
Mots clefs ... : IMPORT PIECE
*****************************************************************}
function TFRecupGB2000.ChargeLaCommande (NaturePiece, RefInterne : string) : boolean ;
var CleDoc_Prec     : R_CLeDoc ;
    Q               : TQUERY   ;
begin
  //
  // chargement entête
  //
  Tob_Commande := TOB.CREATE ('PIECE', NIL, -1);
  AddLesSupEntete (Tob_Commande) ;
  // Lecture entête
  Q:=OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_REFINTERNE="' + RefInterne + '"',False) ;
  if Not Q.EOF then
  begin
    Tob_Commande.SelectDB('',Q) ;
    Ferme(Q) ;
    CleDoc_Prec := TOB2CleDoc (Tob_Commande);
    //
    // Chargement des Lignes
    //
    Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc_Prec,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
    Tob_Commande.LoadDetailDB('LIGNE','','',Q,False,True) ;
    Ferme(Q) ;
    PieceAjouteSousDetail(Tob_Commande);

    Result := True ;
  end else
  begin
    Ferme(Q) ;
    Result := False ;
  end
end;

///////////////////////////////////////////////////////////////////////////////////////
// Importation d'une réception sur commandes dans  PGI
//		-> Importation de la réception fournisseur
//		-> Recherche de la commande d'origine dans PGI pour récupérer le dernier indice
//		-> Génération d'une nouvelle commande d'indice supérieur, reprennant les qtés de la réception
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportLivraison;       // Mettre des controles de cohérence
var Tob_Ligne_Cde : TOB    ;
    Tob_Ligne_Rec : TOB    ;
    NatureCommande: string ;
    RefCommande	  : string ;
    SoucheCdePGI  : string ;
    PieceSuivante : string ;
    NumeroCdePGI  : integer;
    NbLigne       : integer;
begin

  TOB_Acomptes  := nil ;

  if (NombreErreurs (-1) = 0) then
  begin
    ////////////////////////////////////////////////////////////////
    //  1 - Création de la réception
    //////////////////////////////////////////////////////////////////
    //////////////////////////////////////////////////////////////////
    // On force le recalcul de la pièce
 	//////////////////////////////////////////////////////////////////
    PutValueDetail (Tob_Piece,'GP_RECALCULER','X');

    //////////////////////////////////////////////////////////////////
    // Création de la dernière ligne commentaire
    //////////////////////////////////////////////////////////////////
    if (LigneCommACreer = True) then
    begin
      LigneCommACreer:=False;
      CreeLigneCommentaire;
    end;

    Tob_PiedBase := TOB.CREATE ('LE PIED BASE', NIL, -1);
    Tob_PiedEche := TOB.CREATE ('LE PIED ECHE', NIL, -1);
    TOB_Acomptes := TOB.Create ('LES ACOMPTES',Nil,-1) ;
    //////////////////////////////////////////////////////////
    // Calcul des cumuls de la pièce
    // Calcul de PIEDBASE
    //////////////////////////////////////////////////////////
    CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article, nil,nil,nil, DEV) ;
    //////////////////////////////////////////////////////////
    // Génération des PIEDECHE
    //////////////////////////////////////////////////////////
    GereEcheancesGC (Tob_Piece, Tob_Tiers, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
    //////////////////////////////////////////////////////////
    // Sauvegarde Chrono
    //////////////////////////////////////////////////////////
    IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce

    if (Tob_Piece.GetValue ('GP_TOTALQTEFACT') <> 0) and (Tob_Piece.GetValue ('GP_TOTALHTDEV') <> 0) then
    begin
      //////////////////////////////////////////////////////////////////
      //  2.1 Recherche et MAJ de la commande dans PGI
      //////////////////////////////////////////////////////////////////
      NatureCommande := Tob_Param.GetValue ('NATCDE');
      RefCommande    := 'Btq ' + BtqCommandeGB2000 + 'N°' + NumCommandeGB2000 + ' Réf ' + RefCommandeGB2000;

      Tob_Commande   := TOB.Create ('PIECE', nil, -1);
      if ChargeLaCommande (NatureCommande, RefCommande) then
      begin
        ////////////////////////////////////////////////////////////////////////
        // Mise à jour de l'entête de la commande
        ////////////////////////////////////////////////////////////////////////
        SoucheCdePGI  := Tob_Commande.GetValue ('GP_SOUCHE') ;
        NumeroCdePGI  := Tob_Commande.GetValue ('GP_NUMERO') ;
        PieceSuivante := EncodeRefPiece(Tob_Piece) ;
        ExecuteSQL ('UPDATE PIECE SET GP_DEVENIRPIECE="'+PieceSuivante+'" WHERE GP_NATUREPIECEG="'+NatureCommande+'" AND GP_SOUCHE="'+SoucheCdePGI+'" AND GP_NUMERO="'+IntToStr(NumeroCdePGI)+'"');
        ////////////////////////////////////////////////////////////////////////
        //
        ////////////////////////////////////////////////////////////////////////
        // MAJ des lignes de cdes et réceptions
        for NbLigne:=0 to Tob_Commande.Detail.Count-1 do
        begin
          Tob_Ligne_Cde:=Tob_Commande.Detail[NbLigne];
          Tob_Ligne_Rec:=Tob_Piece.Detail[NbLigne];
          Tob_Ligne_Rec.PutValue ('GL_PIECEPRECEDENTE', EncodeRefPiece(Tob_Ligne_Cde)) ;
        end;
      end else
      begin
        EcrireInfo ('Impossible de retrouver la commande référence ' + RefCommandeGB2000, True, True, True);
      end;
      Tob_Commande.free ; Tob_Commande := nil ;
      //////////////////////////////////////////////////////////////////////////
      // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  RECEPTION
      /////////////////////////////////////////////////////////////////////////
      Tob_Piece.SetAllModifie(TRUE) ;
      Tob_PiedBase.SetAllModifie(TRUE) ;
      Tob_PiedEche.SetAllModifie(TRUE) ;

      Tob_Piece.InsertDB (nil);
      Tob_PiedBase.InsertDB (nil);
      Tob_PiedEche.InsertDB (nil);
    end;
  end;
  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece    <> nil) then Tob_Piece.Free;       // Libère TOB Pièce+Lignes réceptions
  if (Tob_Article  <> nil) then Tob_Article.free;     // Libère la TOB des articles de la pièce
  if (Tob_Tiers    <> nil) then Tob_Tiers.free;	      // Libère la TOB du fournisseurs
  if (Tob_PiedBase <> nil) then Tob_PiedBase.free;    // Libère la TOB PiedBase réceptions
  if (Tob_PiedEche <> nil) then Tob_PiedEche.free;    // Libère la TOB PiedEche réceptions
  if (Tob_Acomptes <> nil) then Tob_Acomptes.free;    // Libère la TOB Acomptes réceptions

  Tob_Piece    := nil;
  Tob_Article  := nil;
  Tob_Tiers    := nil;
  Tob_PiedBase := nil;
  Tob_PiedEche := nil;
  Tob_Acomptes := nil;
end ;

//////////////////////////////////////////////////////////////////////////////////////
// Importation d'un transfert inter-boutique GB 2000 dans  PGI
//		-> Duplication du transfert émis pour obtenir le transfert reçu
//		-> Importation des transferts émis et reçus
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportTransfert;
begin

  if (NombreErreurs (-1) = 0) then
  begin
    //////////////////////////////////////////////////////////////////
    // On force le recalcul de la pièce
 	//////////////////////////////////////////////////////////////////
    PutValueDetail (Tob_Piece,'GP_RECALCULER','X');

    //////////////////////////////////////////////////////////////////
    // 1 - Création du transfert émis
    //////////////////////////////////////////////////////////////////
    if (LigneCommACreer = True) then
    begin
      LigneCommACreer:=False;
      CreeLigneCommentaire;
    end;

    Tob_PiedBase:= TOB.CREATE ('LE PIEDBASE', NIL, -1);
    Tob_PiedEche:= TOB.CREATE ('LE PIEDECHE', NIL, -1);
    //Tob_Adresses:= TOB.Create ('LES ADRESSES',Nil,-1) ;
    //TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Livraison}
    //TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Facturation}
    TOB_Acomptes:=TOB.Create('LES ACOMPTES',Nil,-1) ;
    //////////////////////////////////////////////////////////
    // Génération des adresses de livraison et de facturation
    //////////////////////////////////////////////////////////
    //ValideLesAdresses ( Tob_Piece, Tob_Piece, Tob_Adresses) ;
    //////////////////////////////////////////////////////////
    // Calcul des cumuls de la pièce
    // Calcul de PIEDBASE
    //////////////////////////////////////////////////////////
    CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article,nil, nil,nil, DEV) ;
    //////////////////////////////////////////////////////////
    // Génération des PIEDECHE
    //////////////////////////////////////////////////////////
    GereEcheancesGC (Tob_Piece, Tob_Tiers_Defaut, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
    //////////////////////////////////////////////////////////
    // Sauvegarde Chrono
    //////////////////////////////////////////////////////////
    IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce

    //////////////////////////////////////////////////////////////////
    // 2 - Création du transfert recu
    //////////////////////////////////////////////////////////////////
    // LM NatureTrfRecu := Tob_Param.GetValue ('NATTRFRE');
    // LM Etab          := Tob_Piece.GetValue ('GP_DEPOTDEST');

    ///////////////////////////////////////////////////////////////////
    // Détermination de la clé d'accès
    //////////////////////////////////////////////////////////////////
    // LM SoucheTrfPGI := GetSoucheG (NatureTrfRecu, Etab,'') ;
    // LM NumeroTrfPGI := GetNumSoucheG (SoucheTrfPGI);
    // LM IncNumSoucheG (SoucheTrfPGI) ;  	// Incrémentation du numéro de pièce

    ////////////////////////////////////////////////////////////////////////////////
    //  2.3 Duplication des TOB du transfert émis pour créer les TOB du transferts reçu
    ////////////////////////////////////////////////////////////////////////////////
    // LM Tob_Piece_Recu := TOB.Create ('PIECE', nil, -1);
    // LM Tob_Piece_Recu.Dupliquer (Tob_Piece, True, True, True);
    // LM Tob_Piece_Recu.PutValue  ('GP_NATUREPIECEG', NatureTrfRecu);
    // LM Tob_Piece_Recu.PutValue  ('GP_SOUCHE'      , SoucheTrfPGI );
    // LM Tob_Piece_Recu.PutValue  ('GP_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    // LM Tob_Piece_Recu.PutValue  ('GP_INDICEG'     , 0 );
    // LM Tob_Piece_Recu.PutValue  ('GP_VIVANTE'     , '-');       // Pièce vivante    : Non

    // MAJ des lignes de cdes et réceptions
    // LM for NbLigne:=0 to Tob_Piece_Recu.Detail.Count-1 do
    // LM begin
    // LM   Tob_Ligne_Recu:=Tob_Piece_Recu.Detail[NbLigne];
    // LM   Tob_Ligne_Recu.PutValue ('GL_NATUREPIECEG', NatureTrfRecu);
    // LM   Tob_Ligne_Recu.PutValue ('GL_SOUCHE'      , SoucheTrfPGI  );
    // LM   Tob_Ligne_Recu.PutValue ('GL_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    // LM   Tob_Ligne_Recu.PutValue ('GL_INDICEG'     , 0 );
    // LM   Tob_Ligne_Recu.PutValue ('GL_DEPOT'       , Tob_Piece.GetValue ('GP_DEPOTDEST'));
    // LM end;

    // MAJ PIEDBASE
    // LM Tob_PiedB_Recu := TOB.Create ('LE PIED BASE', nil, -1);
    // LM Tob_PiedB_Recu.Dupliquer (Tob_PiedBase, True, True, True);
    // MAJ des lignes de piedbase
    // LM for NbLigne:=0 to Tob_PiedB_Recu.Detail.Count-1 do
    // LM begin
    // LM   Tob_PiedF_Recu := Tob_PiedB_Recu.Detail[NbLigne];
    // LM   Tob_PiedF_Recu.PutValue ('GPB_NATUREPIECEG', NatureTrfRecu);
    // LM   Tob_PiedF_Recu.PutValue ('GPB_SOUCHE'      , SoucheTrfPGI  );
    // LM   Tob_PiedF_Recu.PutValue ('GPB_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    // LM   Tob_PiedF_Recu.PutValue ('GPB_INDICEG'     , 0  );
    // LM end;

    // MAJ PIEDECHE
    // LM Tob_PiedE_Recu := TOB.Create ('LE PIED ECHE', nil, -1);
    // LM Tob_PiedE_Recu.Dupliquer (Tob_PiedEche, True, True, True);
    // MAJ des lignes de piedeche
    // LM for NbLigne:=0 to Tob_PiedE_Recu.Detail.Count-1 do
    // LM begin
    // LM   Tob_PiedF_Recu := Tob_PiedE_Recu.Detail[NbLigne];
    // LM   Tob_PiedF_Recu.PutValue ('GPE_NATUREPIECEG', NatureTrfRecu);
    // LM   Tob_PiedF_Recu.PutValue ('GPE_SOUCHE'      , SoucheTrfPGI  );
    // LM   Tob_PiedF_Recu.PutValue ('GPE_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    // LM   Tob_PiedF_Recu.PutValue ('GPE_INDICEG'     , 0  );
    // LM end;

    /////////////////////////////////////////////////////////////////////////
    // On fait le lien entre le transfert émis et le transfert reçu
    /////////////////////////////////////////////////////////////////////////
    // LM Tob_Piece.PutValue ('GP_DEVENIRPIECE', EncodeRefPiece(Tob_Piece_Recu)) ;
    //////////////////////////////////////////////////////////////////////////
    // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  RECEPTION
    /////////////////////////////////////////////////////////////////////////
    // LM Tob_Piece.SetAllModifie(TRUE) ;
    // LM Tob_PiedBase.SetAllModifie(TRUE) ;
    // LM Tob_PiedEche.SetAllModifie(TRUE) ;
    // LM Tob_Adresses.SetAllModifie(TRUE) ;

    Tob_Piece.InsertDB (nil);
    Tob_PiedBase.InsertDB (nil);
    Tob_PiedEche.InsertDB (nil);
    // LM Tob_Adresses.InsertorUpdateDB (False);

    //////////////////////////////////////////////////////////////////////////
    // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  COMMANDE
    /////////////////////////////////////////////////////////////////////////
    //Tob_Piece_Recu.SetAllModifie(TRUE) ;
    //Tob_PiedB_Recu.SetAllModifie(TRUE) ;
    //Tob_PiedE_Recu.SetAllModifie(TRUE) ;

    //Tob_Piece_Recu.InsertDB (nil);
    //Tob_PiedB_Recu.InsertorUpdateDB (False);
    //Tob_PiedE_Recu.InsertorUpdateDB (False);
  end;

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece     <> nil) then Tob_Piece.Free;       // Libère TOB Pièce+Lignes réceptions
  if (Tob_Article   <> nil) then Tob_Article.free;     // Libère la TOB des articles de la pièce
  //if (Tob_Tiers     <> nil) then Tob_Tiers.free;       // Libère la TOB du fournisseurs
  if (Tob_PiedBase  <> nil) then Tob_PiedBase.free;    // Libère la TOB PiedBase réceptions
  if (Tob_PiedEche  <> nil) then Tob_PiedEche.free;    // Libère la TOB PiedEche réceptions
  //if (Tob_Adresses  <> nil) then Tob_Adresses.free;    // Libère la TOB Adresses réceptions
  if (Tob_Acomptes  <> nil) then Tob_Acomptes.free;    // Libère la TOB Acomptes réceptions
  //if (Tob_Piece_Recu<> nil) then Tob_Piece_Recu.free;  // Libère la TOB Piece+Ligne commande
  //if (Tob_PiedB_Recu<> nil) then Tob_PiedB_Recu.free;  // Libère la TOB PiedBase commande
  //if (Tob_PiedE_Recu<> nil) then Tob_PiedE_Recu.free;  // Libère la TOB PiedEche commande

  Tob_Piece    := nil;
  Tob_Article  := nil;
  Tob_Tiers    := nil;
  Tob_PiedBase := nil;
  Tob_PiedEche := nil;
  //Tob_Adresses := nil;
  Tob_Acomptes := nil;
end ;


///////////////////////////////////////////////////////////////////////////////////////
// Importation d'une facture NEGOCE sur commandes ou sur livraison dans  PGI
//		-> Importation de la facture clients
//		-> Recherche de la commande ou de la livraison d'origine dans PGI pour récupérer le dernier indice
//		-> Génération d'une nouvelle commande ou une nouvelle livraison d'indice supérieur,
//       reprennant les qtés de la réception
///////////////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.ImportFacCdeLiv ;       // Mettre des controles de cohérence
var TobFille      : TOB    ;
    Tob_Piece_Bis : TOB    ;
    Tob_Ligne_Bis : TOB    ;
    Tob_PiedB_Bis : TOB    ;
    Tob_PiedE_Bis : TOB	   ;
    Tob_Adres_Bis : TOB    ;
    Tob_Acomp_Bis : TOB    ;
    Q  		  : TQUERY ;
    SQL		  : string ;
    NatureBis     : string ;
    RefBis    	  : string ;
    SoucheBisPGI  : string ;
    NumeroBisPGI  : integer;
    IndiceBisPGI  : integer;
    Indice	  : integer;
    TrouveBisPGI  : boolean;
    TypePiece     : string ;
    Boutique      : string ;
    Numero        : string ;
    TypePieceOld  : string ;
    BoutiqueOld   : string ;
    NumeroOld     : string ;
    i , j         : integer;
    NumLigne      : integer;
    MinLigne      : integer;
begin
  Tob_Piece_Bis := nil ;
  NumeroBisPGI :=  0;
  IndiceBisPGI := -1;
  
  if (NombreErreurs (-1) = 0) then
  begin
    //////////////////////////////////////////////////////////////////
    // On force le recalcul de la pièce
 	//////////////////////////////////////////////////////////////////
    PutValueDetail (Tob_Piece,'GP_RECALCULER','X');
    //////////////////////////////////////////////////////////////////
    // 1 - Création de la facture
    //////////////////////////////////////////////////////////////////

    //////////////////////////////////////////////////////////////////
    // Création de la dernière ligne commentaire
    //////////////////////////////////////////////////////////////////
    if (LigneCommACreer = True) then
    begin
      LigneCommACreer:=False;
      CreeLigneCommentaire;
    end;

    //////////////////////////////////////////////////////////////////
    // MAJ des remises lignes
    //////////////////////////////////////////////////////////////////
    if Tob_Piece.GetValue ('GP_REMISEPIED') <> 0 then
    begin
      for i:=0 to Tob_Piece.detail.count-1 do Tob_Piece.detail [i].PutValue ('GL_REMISEPIED', Tob_Piece.GetValue ('GP_REMISEPIED'));
    end;

		Tob_PiedBase:= TOB.CREATE ('LE PIED BASE', NIL, -1);
  	Tob_PiedEche:= TOB.CREATE ('LE PIED ECHE', NIL, -1);
  	Tob_Adresses:= TOB.Create ('LES ADRESSES',Nil,-1) ;
		TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Livraison}
		TOB.Create('ADRESSES', Tob_Adresses,-1) ; {Facturation}
  	TOB_Acomptes:=TOB.Create('LES ACOMPTES',Nil,-1) ;
    if Tob_PiedPort = nil then TOB_PiedPort:=TOB.Create('LES PORTS',Nil,-1) ;
  	//////////////////////////////////////////////////////////
  	// Génération des adresses de livraison et de facturation
  	//////////////////////////////////////////////////////////
  	ValideLesAdresses ( Tob_Piece, Tob_Piece, Tob_Adresses) ;
  	//////////////////////////////////////////////////////////
  	// Calcul des cumuls de la pièce
  	// Calcul de PIEDBASE
  	//////////////////////////////////////////////////////////
  	CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article, Tob_PiedPort,nil,nil, DEV) ;
  	//////////////////////////////////////////////////////////
  	// Génération des PIEDECHE
  	//////////////////////////////////////////////////////////
  	GereEcheancesGC (Tob_Piece, Tob_Tiers, Tob_PiedEche, Tob_Acomptes,nil, taCreat, DEV, False);
  	//////////////////////////////////////////////////////////
  	// Sauvegarde Chrono
  	//////////////////////////////////////////////////////////
  	IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce

  	//////////////////////////////////////////////////////////////////////////
  	// MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  RECEPTION
  	/////////////////////////////////////////////////////////////////////////
    Tob_Piece.SetAllModifie(TRUE) ;
    Tob_PiedBase.SetAllModifie(TRUE) ;
    Tob_PiedEche.SetAllModifie(TRUE) ;
    Tob_Adresses.SetAllModifie(TRUE) ;
    Tob_PiedPort.SetAllModifie(TRUE) ;

  	Tob_Piece.InsertDB (nil);
  	Tob_PiedBase.InsertorUpdateDB (False);
  	Tob_PiedEche.InsertorUpdateDB (False);
  	Tob_Adresses.InsertorUpdateDB (False);
    Tob_PiedPort.InsertorUpdateDB (False) ;


    //////////////////////////////////////////////////////////////////
    // 2 - Création des commandes ou de la livraisons
  	//////////////////////////////////////////////////////////////////
    if PieceAGenerer then
    begin
      For i:=0 to Tob_Piece.detail.count-1 do
      begin
        TobFille := Tob_Piece.detail [i];

        if TobFille.FieldExists ('$$_TYPEDOC') then
        begin
          TypePiece := TobFille.GetValue ('$$_TYPEDOC');
          Boutique  := TobFille.GetValue ('$$_CODEBTQ');
          Numero    := TobFille.GetValue ('$$_NUMDOC');

          if (TypePiece <> TypePieceOld) or (Boutique <> BoutiqueOld) or (Numero <> NumeroOld) then
          begin
            TypePieceOld := TypePiece ;
            BoutiqueOld  := Boutique  ;
            NumeroOld    := Numero    ;

            ////////////////////////////////////////////////////////////////
            // Import du document précédent
            ////////////////////////////////////////////////////////////////
            if Tob_Piece_Bis <> nil then
            begin
              //////////////////////////////////////////////////////////////////
              // MAJ des remises lignes
              //////////////////////////////////////////////////////////////////
              if Tob_Piece_Bis.GetValue ('GP_REMISEPIED') <> 0 then
              begin
                for j:=0 to Tob_Piece_Bis.detail.count-1 do Tob_Piece_Bis.detail [j].PutValue ('GL_REMISEPIED', Tob_Piece_Bis.GetValue ('GP_REMISEPIED'));
              end;

              //////////////////////////////////////////////////////////////////
              // La pièce commence à la ligne 1
              //////////////////////////////////////////////////////////////////
              MinLigne := 9999999 ;
              for j:=0 to Tob_Piece_Bis.detail.count-1 do
              begin
                if MinLigne > Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE') then MinLigne := Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE');
              end;

              if MinLigne > 1 then
              begin
                for j:=0 to Tob_Piece.detail.count-1 do
                begin
                  NumLigne := Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE');
                  Tob_Piece_Bis.detail [j].PutValue ('GL_NUMLIGNE', NumLigne - MinLigne + 1 );
                  Tob_Piece_Bis.detail [j].PutValue ('GL_NUMORDRE', NumLigne - MinLigne + 1 );
                end;
              end;

              Tob_PiedB_Bis := TOB.CREATE ('LE PIED BASE', NIL, -1);
            	Tob_PiedE_Bis := TOB.CREATE ('LE PIED ECHE', NIL, -1);
  	          Tob_Adres_Bis := TOB.Create ('LES ADRESSES',Nil,-1) ;
          		TOB.Create('ADRESSES', Tob_Adres_Bis,-1) ; {Livraison}
          		TOB.Create('ADRESSES', Tob_Adres_Bis,-1) ; {Facturation}
  	          TOB_Acomp_Bis :=TOB.Create('LES ACOMPTES',Nil,-1) ;
              //////////////////////////////////////////////////////////
  	          // Génération des adresses de livraison et de facturation
  	          //////////////////////////////////////////////////////////
  	          ValideLesAdresses ( Tob_Piece_Bis, Tob_Piece_Bis, Tob_Adres_Bis) ;
  	          //////////////////////////////////////////////////////////
  	          // Calcul des cumuls de la pièce
  	          // Calcul de PIEDBASE
  	          //////////////////////////////////////////////////////////
  	          CalculFacture (Tob_Piece_Bis, Tob_PiedB_Bis, Tob_Tiers, Tob_Article, nil,nil,nil, DEV) ;
  	          //////////////////////////////////////////////////////////
  	          // Génération des PIEDECHE
  	          //////////////////////////////////////////////////////////
  	          GereEcheancesGC (Tob_Piece_Bis, Tob_Tiers, Tob_PiedE_Bis, Tob_Acomp_Bis,nil, taCreat, DEV, False);
  	          //////////////////////////////////////////////////////////
  	          // Sauvegarde Chrono
  	          //////////////////////////////////////////////////////////
  	          IncNumSoucheG (Tob_Piece_Bis.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce

  	          //////////////////////////////////////////////////////////////////////////
  	          // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  RECEPTION
  	          /////////////////////////////////////////////////////////////////////////
              Tob_Piece_Bis.SetAllModifie(TRUE) ;
              Tob_PiedB_Bis.SetAllModifie(TRUE) ;
              Tob_PiedE_Bis.SetAllModifie(TRUE) ;
              Tob_Adres_Bis.SetAllModifie(TRUE) ;

  	          Tob_Piece_Bis.InsertDB (nil);
           	  Tob_PiedB_Bis.InsertorUpdateDB (False);
  	          Tob_PiedE_Bis.InsertorUpdateDB (False);
  	          Tob_Adres_Bis.InsertorUpdateDB (False);

              if (Tob_Piece_Bis<> nil)	then Tob_Piece_Bis.free;   // Libère la TOB Piece+Ligne commande
              if (Tob_PiedB_Bis<> nil)	then Tob_PiedB_Bis.free;   // Libère la TOB PiedBase commande
              if (Tob_PiedE_Bis<> nil)	then Tob_PiedE_Bis.free;   // Libère la TOB PiedEche commande
              if (Tob_Acomp_Bis<> nil)	then Tob_Acomp_Bis.free;   // Libère la TOB PiedEche commande
              if (Tob_Adres_Bis<> nil)	then Tob_Adres_Bis.free;   // Libère la TOB PiedEche commande
            end;

            if TypePiece = 'C' then NatureBis := Tob_Param.GetValue ('NATNEGCDE')
            else NatureBis := Tob_Param.GetValue ('NATNEGLIV');

            RefBis  := 'Btq ' + Boutique + 'N° ' + Numero ;

            NumeroBisPGI :=  0;
            IndiceBisPGI := -1;
            TrouveBisPGI := False;
            //
            // La pièce existe déjà dans PGI ?
            // Si oui, on recherche l'indice le + élevé
            //
   	        SQL:='Select GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO, GP_INDICEG From PIECE WHERE GP_NATUREPIECEG="'+NatureBis+'" AND GP_REFINTERNE="'+RefBis+'"';
            Q:=OpenSQL(SQL,True) ;
            if Not Q.EOF then
            begin
              while not Q.EOF do
              begin
      	        TrouveBisPGI := True;
      	        SoucheBisPGI := Q.FindField('GP_SOUCHE').AsString ;
                NumeroBisPGI := Q.FindField('GP_NUMERO').AsInteger;
                Indice := Q.FindField('GP_INDICEG').AsInteger;
                if Indice>IndiceBisPGI then IndiceBisPGI:=Indice;
                Q.Next;
     	        end;
            end;
            Ferme(Q);

            //////////////////////////////////////////////////////////////////
            //  2.2  Détermination de la clef d'acces de la commande
            //////////////////////////////////////////////////////////////////
            if TrouveBisPGI then
            begin
              Inc (IndiceBisPGI);
            end else
            begin
              IndiceBisPGI:=0;
              /////////////////////////
    	        // Souche, Numéro
    	        ////////////////////////
    	        SouchebisPGI := GetSoucheG(NatureBis, Boutique,'') ;
    	        NumeroBisPGI := GetNumSoucheG(SouchebisPGI) ;
              IncNumSoucheG (SoucheBisPGI) ;  	// Incrémentation du numéro de pièce
            end;

            ////////////////////////////////////////////////////////////////////////////////
            //  2.3 Duplication des TOB de la réception pour créer les TOB de la commande
            //      et MAJ des Tob de la commande
            ////////////////////////////////////////////////////////////////////////////////
            Tob_Piece_Bis := TOB.Create ('PIECE', nil, -1);
            Tob_Piece_Bis.Dupliquer (Tob_Piece, False, True, True);
            Tob_Piece_Bis.PutValue ('GP_NATUREPIECEG', NatureBis);
            Tob_Piece_Bis.PutValue ('GP_SOUCHE'      , SoucheBisPGI  );
            Tob_Piece_Bis.PutValue ('GP_NUMERO'      , NumeroBisPGI  );
            Tob_Piece_Bis.PutValue ('GP_INDICEG'     , IndiceBisPGI  );
            Tob_Piece_Bis.PutValue ('GP_REFINTERNE'  , RefBis        ) ;
            Tob_Piece_Bis.PutValue ('GP_DEVENIRPIECE',EncodeRefPiece(Tob_Piece)) ;
            Tob_Piece_Bis.PutValue ('GP_VIVANTE'       , '-');       // Pièce vivante    : Non
          end;

          Tob_Ligne_Bis := TOB.Create ('LIGNE', nil, -1);
          Tob_Ligne_Bis.Dupliquer    (TobFille, True, True, True);
          Tob_Ligne_Bis.ChangeParent (Tob_Piece_Bis, -1);

          Tob_Ligne_Bis.PutValue ('GL_NATUREPIECEG', NatureBis);
          Tob_Ligne_Bis.PutValue ('GL_SOUCHE'      , SoucheBisPGI  );
          Tob_Ligne_Bis.PutValue ('GL_NUMERO'      , NumeroBisPGI  );
    	  Tob_Ligne_Bis.PutValue ('GL_INDICEG'     , IndiceBisPGI  );

          //
          // Détermine le nouveau numéro de ligne
          //
          // Inc (NewNumLigne);
          //Tob_Ligne_Bis.PutValue ('GL_NUMLIGNE'    , NewNumLigne  );
          //Tob_Ligne_Bis.PutValue ('GL_NUMORDRE'    , NewNumLigne  );

          //
          // MAJ de la ligne facture
          //
          TobFille.PutValue ('GL_PIECEPRECEDENTE', EncodeRefPiece(Tob_Ligne_Bis)) ;
        end ;
      end;

      ////////////////////////////////////////////////////////////////
      // Import du document précédent
      ////////////////////////////////////////////////////////////////
      if Tob_Piece_Bis <> nil then
      begin
        //////////////////////////////////////////////////////////////////
        // MAJ des remises lignes
        //////////////////////////////////////////////////////////////////
        if Tob_Piece_Bis.GetValue ('GP_REMISEPIED') <> 0 then
        begin
          for j:=0 to Tob_Piece_Bis.detail.count-1 do Tob_Piece_Bis.detail [j].PutValue ('GL_REMISEPIED', Tob_Piece_Bis.GetValue ('GP_REMISEPIED'));
        end;

        //////////////////////////////////////////////////////////////////
        // La pièce commence à la ligne 1
        //////////////////////////////////////////////////////////////////
        MinLigne := 9999999 ;
        for j:=0 to Tob_Piece_Bis.detail.count-1 do
        begin
          if MinLigne > Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE') then MinLigne := Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE');
        end;

        if MinLigne > 1 then
        begin
          for j:=0 to Tob_Piece_Bis.detail.count-1 do
          begin
            NumLigne := Tob_Piece_Bis.detail [j].GetValue ('GL_NUMLIGNE');
            Tob_Piece_Bis.detail [j].PutValue ('GL_NUMLIGNE', NumLigne - MinLigne + 1 );
            Tob_Piece_Bis.detail [j].PutValue ('GL_NUMORDRE', NumLigne - MinLigne + 1 );
          end;
        end;

        Tob_PiedB_Bis := TOB.CREATE ('LE PIED BASE', NIL, -1);
       	Tob_PiedE_Bis := TOB.CREATE ('LE PIED ECHE', NIL, -1);
  	    Tob_Adres_Bis := TOB.Create ('LES ADRESSES',Nil,-1) ;
        TOB.Create('ADRESSES', Tob_Adres_Bis,-1) ; {Livraison}
        TOB.Create('ADRESSES', Tob_Adres_Bis,-1) ; {Facturation}
  	    TOB_Acomp_Bis :=TOB.Create('LES ACOMPTES',Nil,-1) ;
        //////////////////////////////////////////////////////////
  	    // Génération des adresses de livraison et de facturation
  	    //////////////////////////////////////////////////////////
  	    ValideLesAdresses ( Tob_Piece_Bis, Tob_Piece_Bis, Tob_Adres_Bis) ;
  	    //////////////////////////////////////////////////////////
  	    // Calcul des cumuls de la pièce
  	    // Calcul de PIEDBASE
  	    //////////////////////////////////////////////////////////
  	    CalculFacture (Tob_Piece_Bis, Tob_PiedB_Bis, Tob_Tiers, Tob_Article,nil, nil,nil, DEV) ;
  	    //////////////////////////////////////////////////////////
  	    // Génération des PIEDECHE
  	    //////////////////////////////////////////////////////////
  	    GereEcheancesGC (Tob_Piece_Bis, Tob_Tiers, Tob_PiedE_Bis, Tob_Acomp_Bis,nil, taCreat, DEV, False);
  	    //////////////////////////////////////////////////////////
  	    // Sauvegarde Chrono
  	    //////////////////////////////////////////////////////////
  	    IncNumSoucheG (Tob_Piece_Bis.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce
        //////////////////////////////////////////////////////////////////////////
  	    // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES  RECEPTION
  	    /////////////////////////////////////////////////////////////////////////
        Tob_Piece_Bis.SetAllModifie(TRUE) ;
        Tob_PiedB_Bis.SetAllModifie(TRUE) ;
        Tob_PiedE_Bis.SetAllModifie(TRUE) ;
        Tob_Adres_Bis.SetAllModifie(TRUE) ;

  	    Tob_Piece_Bis.InsertDB (nil);
        Tob_PiedB_Bis.InsertorUpdateDB (False);
  	    Tob_PiedE_Bis.InsertorUpdateDB (False);
  	    Tob_Adres_Bis.InsertorUpdateDB (False);

        if (Tob_Piece_Bis<> nil)	then Tob_Piece_Bis.free;   // Libère la TOB Piece+Ligne commande
        if (Tob_PiedB_Bis<> nil)	then Tob_PiedB_Bis.free;   // Libère la TOB PiedBase commande
        if (Tob_PiedE_Bis<> nil)	then Tob_PiedE_Bis.free;   // Libère la TOB PiedEche commande
        if (Tob_Acomp_Bis<> nil)	then Tob_Acomp_Bis.free;   // Libère la TOB PiedEche commande
        if (Tob_Adres_Bis<> nil)	then Tob_Adres_Bis.free;   // Libère la TOB PiedEche commande
      end;
    end;
  end;

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  if (Tob_Piece    <> nil) then Tob_Piece.Free;        // Libère TOB Pièce+Lignes réceptions
  if (Tob_Article  <> nil)	then Tob_Article.free;		 // Libère la TOB des articles de la pièce
  if (Tob_Tiers    <> nil)	then Tob_Tiers.free;	     // Libère la TOB du fournisseurs
  if (Tob_PiedBase <> nil)	then Tob_PiedBase.free;    // Libère la TOB PiedBase réceptions
  if (Tob_PiedEche <> nil)	then Tob_PiedEche.free;    // Libère la TOB PiedEche réceptions
  if (Tob_Adresses <> nil)	then Tob_Adresses.free;    // Libère la TOB Adresses réceptions
  if (Tob_Acomptes <> nil)	then Tob_Acomptes.free;    // Libère la TOB Acomptes réceptions

  Tob_Piece    := nil;
  Tob_Article  := nil;
  Tob_Tiers    := nil;
  Tob_PiedBase := nil;
  Tob_PiedEche := nil;
  Tob_Adresses := nil;
  Tob_Acomptes := nil;
end ;


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Récupération des commandes fournisseurs GB 2000
// 3 cas de figure :  1 - Commande en cours (non réceptionnée)
//                        -> récupération totale dans PGI
//										2 - Commande soldée (totalement réceptionnée)
//												-> Pas de récupération, elle sera créée lors de la récupération des livraison
//										3 - Commande partiellement livrée
//												-> la partie "livrée" sera récupérée en même temps que les réceptions
//												-> Seule la partie "non livrée" est prise en compte.
//
//										Remarque : Si une commande est partiellement livrée, on doit forcément retrouver la commande
//															 d'origine dans PGI pour pouvoir inétgrer le reliquat.
//															 Cela signifie que les réceptions fournisseurs doivent être récupérées
//															 AVANT les commandes fournisseurs
//
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupCommande(st : string);
var DateConv         : TDateTime;
    DD   	         : TDateTime;
    Q                : TQUERY   ;
    CodeArt 	     : string   ;
    Rang    	     : string   ;
    CategorieDim     : string   ;
    NaturePiece      : string   ;
    Etab             : string   ;
    SQL              : string   ;
    CodeDim1,CodeDim2: string   ;
    CodeDim3,CodeDim4: string   ;
    CodeDim5 	     : string   ;
    RefCde	         : string   ;		// Référence commande = N° GB + Réf Cde GB
    CodeFrn	         : string   ;
    GL_CodeArticle   : string   ;
    DevisePiece      : string   ;
    CodeDevPGI       : string   ;
    CodeDevGB        : string   ;
    Qte, Prix	     : double   ;
    QteCde	         : double   ;
    QteLivre         : double   ;
    QteReste         : double   ;
    Montant          : double   ;
    MontantDev       : double   ;
    MontantOld       : double   ;
    Per,Sem          : integer  ;
    Taille, i 	     : integer  ;
begin

  NaturePiece := Tob_Param.GetValue ('NATCDE');
  Etab := copy (st, 7, 3);

  if ArretRecup then exit;
  if Copy(st,1,5)='CDEC1' then
  begin
    if (Tob_Piece <> Nil) then ImportPiece ;

    if (copy (st, 65, 1) = 'T')  then
    begin
      CommandeSoldee  := True ;
      CommandePartiel := False ;
    end else
    begin
      CommandeSoldee  := False ;
      if copy (st, 65, 1) = 'C' then CommandePartiel:=False
      else CommandePartiel:=True;
    end;
    EcrireInfo ('Commande ' + Copy(st,10 ,6) + ' de la boutique ' + copy (st, 7, 3), False, True, False);

    //////////////////////////////////////////////////////////////////////
    // Initialisations diverses et varièes
    //////////////////////////////////////////////////////////////////////
    TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0 ;
    Numligne := 0;            // Initialisation du premier numéro de ligne
    NombreErreurs(0);					// Initialisation du nombre d'erreurs
    QualifMvt:=GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
    LigneCommACreer:=False;

    Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
    //Tob_Piece.initValeurs;
    AddLesSupEntete (Tob_Piece) ;
    InitTobPiece (Tob_Piece);

	///////////////////////////
  	// Nature de pièce
    /////////////////////////
    Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
    ////////////////////////
    // Souche, Numéro, .....
    ////////////////////////
    CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
    CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
    CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');
    PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
    PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
   	PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;

    ////////////////////////////
    // Date
    ////////////////////////////
    Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 16, 8, tcDate8AMJ, false);
    Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 16, 8, tcDate8AMJ, false);
    //////////////////////////////////////////////////////
    // Référence de commande
    //////////////////////////////////////////////////////
    RefCde := 'Btq ' + copy (st, 7, 3) + 'N°' + copy (st, 10, 6) + ' Réf ' + copy (st, 29, 8);
    PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefCde) ;
    ///////////////////////
    // Fournisseur
    ///////////////////////
    CodeFrn := copy (st, 24, 5);
    Tob_Piece.PutFixedStValue('GP_TIERS'       , st, 24, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSLIVRE'  , st, 24, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSFACTURE', st, 24, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSPAYEUR' , st, 24, 5, tcTrim, false);
    //////////////////////////////////////
    // Chargement en TOB du fournisseur
    //////////////////////////////////////
    Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeFrn + '" AND T_NATUREAUXI = "FOU" ',False) ;
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      Tob_Tiers.SelectDB('',Q);
      ///////////////////////////////////////////////////////////////
      // Initialise les zones de l'entête en fonction du fournisseur
      ////////////////////////////////////////////////////////////////
   	  Tob_Piece.PutValue('GP_FACTUREHT'	      , 'X');                  // Commande GB 2000 forcément en HT
      Tob_Piece.PutValue('GP_ESCOMPTE'        , Tob_Tiers.GetValue('T_ESCOMPTE')) ;
	  Tob_Piece.PutValue('GP_MODEREGLE'       , Tob_Tiers.GetValue('T_MODEREGLE')) ;
	  Tob_Piece.PutValue('GP_REGIMETAXE'      , Tob_Tiers.GetValue('T_REGIMETVA')) ;
      Tob_Piece.PutValue('GP_TVAENCAISSEMENT' , PositionneExige   (Tob_Tiers));
	  Tob_Piece.PutValue('GP_QUALIFESCOMPTE'  , Tob_Tiers.GetValue('T_QUALIFESCOMPTE')) ;
	  Tob_Piece.PutValue('GP_TARIFTIERS'      , Tob_Tiers.GetValue('T_TARIFTIERS')) ;
    end else
    begin
      EcrireInfo ('   Erreur : -> Fournisseur '+ Copy(st,24,5)+ ' inexistant dans PGI', True, True, True);
      NombreErreurs (1);
    end;
    Ferme(Q) ;

    if (NombreErreurs (-1) = 0) then
    begin
   	  Tob_Piece.PutFixedStValue('GP_ETABLISSEMENT', st, 7, 3, tcTrim, false);
      Tob_Piece.PutFixedStValue('GP_DEPOT', st, 7, 3, tcTrim, false);

      //////////////////////////////////////////////
      // Champs par  défaut
      //////////////////////////////////////////////
      if CommandeSoldee then Tob_Piece.PutValue ('GP_VIVANTE'       , '-')
      else Tob_Piece.PutValue ('GP_VIVANTE'       , 'X');
      Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');      // Pièce visée      : Non
      Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
      Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'ACH');  // Document         : Achat
      Tob_Piece.PutValue ('GP_EDITEE'        , 'X');        // Pièce éditée     : Oui
      Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
      Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
      Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
	  Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_DATECREATION'   , Date) ;
      Tob_Piece.PutValue ('GP_HEURECREATION'  , NowH) ;
      Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;

      /////////////////////////////////////////////
      // Devise du document
   	  /////////////////////////////////////////////
      CodeDevGB  := Trim (copy (st, 59, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      DevisePiece := CodeDevPGI;
      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      PutValueDetail(Tob_Piece,'GP_TAUXDEV',DEV.Taux) ;
      AttribCotation(Tob_Piece) ;
	  PutValueDetail(Tob_Piece,'GP_DATETAUXDEV',DEV.DateTaux) ;

      DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
	  Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
	  Tob_Piece.PutValue('GP_PERIODE',Per) ;
      Tob_Piece.PutValue('GP_SEMAINE',Sem) ;
    end;
  end
  else if Copy(st,1,5)='CDLC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      ////////////////////////////////////////////////////////////////
      // Création de la ligne commentaire précédente ???
      /////////////////////////////////////////////////////////////////
      if (LigneCommACreer = True) then
      begin
      	LigneCommACreer:=False;
      	CreeLigneCommentaire;
      end;
      TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
      ////////////////////////////////////////////////
      // Formatage de la référence article générique
      ////////////////////////////////////////////////
      CodeArticle := Copy (st, 28, 12) ;
      codeart     := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
      //////////////////////////////////////////
      // Récupération du libellé de l'article
      //////////////////////////////////////////
      SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then	Designation := Q.FindField('GA_LIBELLE').AsString
      else begin
       	EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,28,12), True, True, True);
       	NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;
    if (NombreErreurs (-1) = 0) then
    begin
      /////////////////////////////////////////////////////////////////
      // gestion d'une ligne commentaire avec l'article générique ????
      /////////////////////////////////////////////////////////////////
      NumLigne        := NumLigne+1;
      NumLigneComment := Numligne;
      LigneCommACreer := True;
    end;
  end else if Copy(st,1,5)='CDLC2' then
  begin
    // Rien à faire
  end else if Copy(st,1,5)='CDTC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      ////////////////////////////////////////////////////////////////
      // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
      ////////////////////////////////////////////////////////////////
      Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
      //Tob_Ligne.initValeurs;
      ////////////////////////////////////////////////////////////////
      // Initialise les champs de la ligne à partir de l'entête
      ////////////////////////////////////////////////////////////////
      PieceVersLigne (Tob_Piece, Tob_Ligne);
      Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
      Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
      //////////////////////////////////////////////
      // Nouveau numéro de ligne
      //////////////////////////////////////////////
      NumLigne := NumLigne+1;
      Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
      Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
      ////////////////////////////////////////////////////////////////
      // Ajoute des champs supplémentaires pour le calcul des cumuls
      ////////////////////////////////////////////////////////////////
      AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
      //////////////////////////////////////////////////////////
      // Détermination de l'article taille
      /////////////////////////////////////////////////////////
      CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

      Taille := StrToInt (copy (st, 28 ,2));
      rang   := CalculTaillePGI (Taille);

      CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
      CodeDim4 := '   '; CodeDim5 := '   ';
      if (CategorieDim [3] = '1') then Codedim1 :=rang
      else if (CategorieDim [3] = '2') then Codedim2 :=rang
      else if (CategorieDim [3] = '3') then Codedim3 :=rang
      else if (CategorieDim [3] = '4') then Codedim4 :=rang
      else if (CategorieDim [3] = '5') then Codedim5 :=rang;
      Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        if (Tob_Article=nil) then Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
        /////////////////////////////////////////////////////////////////////////////
        // Initialisation des champs de la ligne à partir de l'article dimensionné
        /////////////////////////////////////////////////////////////////////////////
        Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
        Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
        Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
        Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
        Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
        Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
        Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
        Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
        Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
        Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
        Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
        Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
        Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
        Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
        Tob_Ligne.PutValue ('GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
        Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
        Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
        Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
        Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
        for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
        for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
        Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
        /////////////////////////////////////////////////////////////////////////////
        // Ajout de l'article dans la TOB Article
        /////////////////////////////////////////////////////////////////////////////
        Tob_Article.SelectDB('',Q);
      end
      else begin
        EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
        NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Ligne.PutValue       ('GL_ARTICLE'     , Codeart);
      Tob_Ligne.PutValue       ('GL_LIBELLE'     , Designation);

      GL_CodeArticle := Trim (CodeArticle);
      Tob_Ligne.PutValue       ('GL_CODEARTICLE' , GL_CodeArticle);
      Tob_Ligne.PutValue       ('GL_REFARTSAISIE', GL_CodeArticle);

      Tob_Ligne.PutValue       ('GL_TARIF'       , 0);
      Tob_Ligne.PutValue       ('GL_TYPELIGNE'   , 'ART');
      Tob_Ligne.PutValue       ('GL_TYPEARTICLE' , 'MAR');
      Tob_Ligne.PutValue       ('GL_VALIDECOM'   , 'NON') ;
      Tob_Ligne.PutValue       ('GL_TYPEREF'     , 'ART');

      QteCde   := StrToFloat (copy (st, 30,4)) ;
      QteLivre := StrToFloat (copy (st, 58,4)) ;
      QteReste := QteCde-QteLivre ;
      if QteReste < 0 then QteReste := 0 ;

      if Commandesoldee then
      begin
        Tob_Ligne.PutValue('GL_QTESTOCK', QteCde);
        Tob_Ligne.PutValue('GL_QTEFACT' , QteCde);
        Tob_Ligne.PutValue('GL_QTERESTE', 0);
      end else
      begin
       Tob_Ligne.PutValue('GL_QTESTOCK', QteCde);
       Tob_Ligne.PutValue('GL_QTEFACT' , QteCde);
       Tob_Ligne.PutValue('GL_QTERESTE', QteReste);
      end;

      //////////////////////////////////////////////////////////
      // récupération des PA et PR
      //////////////////////////////////////////////////////////
      // -> devise des prix ?
      DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
      begin
        if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
        else                  DevisePiece := 'EUR';
      end;
      //
      // PA de l'article
      //
      Montant := StrToFloat (copy (st, 34, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPA'      , Montant);
        Tob_Ligne.PutValue('GL_PMAP'     , Montant);
        Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPA'         , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAPACTU'    , MontantDev);
      end;
      //
      // PR de l'article
      //
      Montant := StrToFloat (copy (st, 77, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPR'      , Montant);
        Tob_Ligne.PutValue('GL_PMRP'     , Montant);
        Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPR'          , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRPACTU'    , MontantDev);
      end;
      //
      // Valorisation de la commande
      //
      Tob_Ligne.PutFixedStValue('GL_PUHT'    , st, 34, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTDEV' , st, 34, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTNET' , st, 34, 12, tcDouble100, false);
      Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);
      Tob_Ligne.PutValue       ('GL_TYPEDIM', 'DIM');
      Tob_Ligne.PutValue       ('GL_QUALIFMVT', QualifMvt) ;

      Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
      Prix:=Tob_Ligne.GetValue('GL_PUHTDEV') ;
      TotQte      :=TotQte+QteCde;
      TotQteReste :=TotQteReste+QteReste;
      TotPrix     :=TotPrix+Qte*Prix;
     end;
  end;
end;

///////////////////////////////////////////////////////////////////
// Récupération des réceptions fournisseurs GB 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupReception(st : string);
var CodeArt 	     : string;
    Rang    	     : string;
    CategorieDim     : string;
    NaturePiece      : string;
    Etab             : string;
    SQL              : string;
    CodeDim1,CodeDim2: string;
    CodeDim3,CodeDim4: string;
    CodeDim5 	     : string;
    RefRec   	     : string;		// Référence commande = N° GB + Réf Cde GB
    CodeFrn	         : string;
    GL_CodeArticle   : string;
    DevisePiece      : string;
    CodeDevPGI       : string;
    CodeDevGB        : string;
    Qte, Prix	     : double;
    qteRec	         : double;
    Montant          : double;
    MontantDev       : double;
    MontantOld       : double;
    Q                : TQUERY ;
    Per,Sem          : integer;
    i		         : integer;
    Taille           : integer;
    DD   	         : TDateTime ;
    DateConv         : TDateTime;
begin

  NaturePiece := Tob_Param.GetValue ('NATREC');
  Etab := copy (st, 7, 3);

  if ArretRecup then exit;

  if Copy(st,1,5)='REEC1' then
  begin
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Importation de la réception précédente
    // 2 cas :	1 - Réception sans commande -> facile, importation directe
    //		2 - Réception dsur commande -> Recherche de la commande d'origine dans PGI (pour récup de l'indice)
    //						    												 Génération et Importation de la commande correspondante
    //																				 Importation de la réception
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Piece <> Nil) then
    begin
      if ReceptSurCommande then ImportLivraison
      else ImportPiece;
    end;

    EcrireInfo ('   ' + 'Boutique ' + Copy(st,7,3) + ' Réception ' + Copy (st,11, 6) , True, True, False);

    //////////////////////////////////////////////////////////////////////
    // Initialisations diverses et varièes
    //////////////////////////////////////////////////////////////////////
    TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
    Numligne := 0;            // Initialisation du premier numéro de ligne
    NombreErreurs(0);					// Initialisation du nombre d'erreurs
    QualifMvt:=GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
    LigneCommACreer:=False;

    Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
    AddLesSupEntete ( Tob_Piece);
    ////////////////////////////////
    // Réception sur commandes ????
    ////////////////////////////////
    if (copy (st, 52, 1) = '0') then
    begin
      ReceptSurCommande := True;
      BtqCommandeGB2000 := copy (st,  7, 3);
      NumCommandeGB2000 := copy (st, 38, 6);
      RefCommandeGB2000 := copy (st, 30, 8);
      RetourFrn         := False ;
    end	else
    begin
      ReceptSurCommande := False;
      if copy (st, 52, 1) = '6' then RetourFrn:=True
      else RetourFrn:=False;
    end;

    ////////////////////////////
    // Nature de pièce + Date
    ////////////////////////////
    Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
    Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 17, 8, tcDate8AMJ, false);
    Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 17, 8, tcDate8AMJ, false);
    /////////////////////////
    // Souche, Numéro, .....
    ////////////////////////
    CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
    CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
    CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');

    PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
    PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
    PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
    //////////////////////////////////////////////////////
		// Référence de commande
    //////////////////////////////////////////////////////
    RefRec := 'Btq ' + copy (st, 7, 3) + 'N°' + copy (st, 11, 6) + ' Réf ' + copy (st, 30, 8);
    PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefRec) ;
    ///////////////////////
    // Fournisseur
    ///////////////////////
    CodeFrn := copy (st, 25, 5);
    Tob_Piece.PutFixedStValue('GP_TIERS'       , st, 25, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSLIVRE'  , st, 25, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSFACTURE', st, 25, 5, tcTrim, false);
    Tob_Piece.PutFixedStValue('GP_TIERSPAYEUR' , st, 25, 5, tcTrim, false);
    // Chargement en TOB du fournisseur
    Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeFrn + '" AND T_NATUREAUXI = "FOU" ',False) ;
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      //Tob_Tiers.initValeurs;
      Tob_Tiers.SelectDB('',Q);
      ////////////////////////////////////////////////////////////////
      // Initialise les zones de l'entête en fonction du fournisseur
      ////////////////////////////////////////////////////////////////
      Tob_Piece.PutValue('GP_FACTUREHT'	     , 'X');  // Réception GB 2000 en HT
      Tob_Piece.PutValue('GP_ESCOMPTE'       , Tob_Tiers.GetValue('T_ESCOMPTE')) ;
      Tob_Piece.PutValue('GP_MODEREGLE'      , Tob_Tiers.GetValue('T_MODEREGLE')) ;
      Tob_Piece.PutValue('GP_REGIMETAXE'     , Tob_Tiers.GetValue('T_REGIMETVA')) ;
      Tob_Piece.PutValue('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers));
      Tob_Piece.PutValue('GP_QUALIFESCOMPTE' , Tob_Tiers.GetValue('T_QUALIFESCOMPTE')) ;
      Tob_Piece.PutValue('GP_TARIFTIERS'     , Tob_Tiers.GetValue('T_TARIFTIERS')) ;
    end else
    begin
      EcrireInfo ('   Erreur -> Fournisseur '+ Copy(st,25,5)+ ' inexistant dans PGI', True, True, True);
      NombreErreurs (1);
    end;
    Ferme(Q) ;

    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Piece.PutFixedStValue('GP_ETABLISSEMENT', st, 7, 3, tcTrim, false);
      Tob_Piece.PutFixedStValue('GP_DEPOT'        , st, 7, 3, tcTrim, false);

      ///////////////////////////////////////////////
      // Champs par  défaut
      //////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_VIVANTE'       , 'X');       // Pièce vivante    : Oui
      Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');      // Pièce visée      : Non
      Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
      Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'ACH');  // Document         : Achat
      Tob_Piece.PutValue ('GP_EDITEE'        , 'X');        // Pièce éditée     : Oui
      Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
      Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
      Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
      Tob_Piece.PutValue('GP_DATECREATION'   , Date) ;
      Tob_Piece.PutValue('GP_HEURECREATION'  , NowH) ;
      Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;

      ///////////////////////////////////////////////////////////////
      // Devise du document
      // Recherche de la devise PGI dans la table de correcpondance
      ///////////////////////////////////////////////////////////////
      CodeDevGB  := Trim (copy (st, 74, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      DevisePiece:=CodeDevPGI;
      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
           Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
           Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece ;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      PutValueDetail(Tob_Piece, 'GP_TAUXDEV'    ,DEV.Taux) ;
      AttribCotation(Tob_Piece) ;
      PutValueDetail(Tob_Piece, 'GP_DATETAUXDEV',DEV.DateTaux) ;

      DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
      Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
      Tob_Piece.PutValue('GP_PERIODE',Per) ;
      Tob_Piece.PutValue('GP_SEMAINE',Sem) ;
    end;
  end
  else if Copy(st,1,5)='RELC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      /////////////////////////////////////////////////////////////////
      // Création de la ligne commentaire précédente ???
      /////////////////////////////////////////////////////////////////
      if (LigneCommACreer = True) then
      begin
      	LigneCommACreer:=False;
        CreeLigneCommentaire;
      end;
      /////////////////////////////////////////////////////////////////
      // gestion d'une ligne commentaire avec l'article générique ????
      /////////////////////////////////////////////////////////////////
      NumLigne        := NumLigne+1;
      NumLigneComment := Numligne;
      LigneCommACreer := True;

      TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
      ////////////////////////////////////////////////
      // formatage de la référence article générique
      ////////////////////////////////////////////////
      CodeArticle := Copy (st, 31, 12) ;
      CodeArt     := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
      //////////////////////////////////////////
      // Récupération du libellé de l'article
      //////////////////////////////////////////
      SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then Designation := Q.FindField('GA_LIBELLE').AsString
      else begin
        EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,31,12), True, True, True);
        NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;
  end
  else if Copy(st,1,5)='RELC2' then
  begin
    // Rien à faire
  end
  else if Copy(st,1,5)='RETC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      ////////////////////////////////////////////////////////////////
      // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
      ////////////////////////////////////////////////////////////////
      Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
      ////////////////////////////////////////////////////////////////
      // Initialise les champs de la ligne à partir de l'entête
      ////////////////////////////////////////////////////////////////
      PieceVersLigne (Tob_Piece, Tob_Ligne);
      Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
      Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
      //////////////////////////////////////////////
      // Nouveau numéro de ligne
      //////////////////////////////////////////////
      NumLigne := NumLigne+1;
      Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
      Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
      ////////////////////////////////////////////////////////////////
      // Ajoute des champs supplémentaires pour le calcul des cumuls
      ////////////////////////////////////////////////////////////////
      AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
      //////////////////////////////////////////////////////////
      // Détermination de l'article taille
      /////////////////////////////////////////////////////////
      CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

      Taille := StrToInt (copy (st, 21 ,2));
      rang   := CalculTaillePGI (Taille);

      CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
      CodeDim4 := '   '; CodeDim5 := '   ';
      if (CategorieDim [3] = '1') then Codedim1 :=rang
      else if (CategorieDim [3] = '2') then Codedim2 :=rang
      else if (CategorieDim [3] = '3') then Codedim3 :=rang
      else if (CategorieDim [3] = '4') then Codedim4 :=rang
      else if (CategorieDim [3] = '5') then Codedim5 :=rang;
      Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        if (Tob_Article=nil) then
      	begin
       	  Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     	  //Tob_Article.initValeurs;
      	end;
	    /////////////////////////////////////////////////////////////////////////////
      	// Initialisation des champs de la ligne à partir de l'article dimensionné
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
        Tob_Ligne.PutValue ('GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
        Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
        Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	    Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	    Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
      	for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
        for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
        Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
      	/////////////////////////////////////////////////////////////////////////////
      	// Ajout de l'article dans la TOB Article
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Article.SelectDB('',Q);
      end
      else begin
      	EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
	  Tob_Ligne.PutValue ('GL_ARTICLE'     , Codeart);
  	  Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

      GL_CodeArticle := Trim (CodeArticle);
      Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	  Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

  	  Tob_Ligne.PutValue ('GL_TARIF'       , 0);
      Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	  Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
      Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON');
      Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

      QteRec := StrtoFloat (copy (st, 23, 4));
      if RetourFrn then QteRec:=-QteRec;
  	  Tob_Ligne.PutValue ('GL_QTESTOCK', QteRec);
   	  Tob_Ligne.PutValue ('GL_QTEFACT' , QteRec);
      Tob_Ligne.PutValue ('GL_QTERESTE', QteRec);

      //////////////////////////////////////////////////////////
  	  // récupération des PA et PR
      //////////////////////////////////////////////////////////
      // -> devise des prix ?
      DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
      begin
        if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
        else                    DevisePiece := 'EUR';
      end;
      //
      // PA de l'article
      //
      Montant := StrToFloat (copy (st, 27, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPA'      , Montant);
        Tob_Ligne.PutValue('GL_PMAP'     , Montant);
        Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPA'         , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAPACTU'    , MontantDev);
      end;
      //
      // PR de l'article
      //
      Montant := StrToFloat (copy (st, 80, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPR'      , Montant);
        Tob_Ligne.PutValue('GL_PMRP'     , Montant);
        Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPR'          , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRPACTU'    , MontantDev);
      end;
      //
      // Valorisation de la réception
      //
	  Tob_Ligne.PutFixedStValue('GL_PUHT'    , st, 27, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTDEV' , st, 27, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTNET' , st, 27, 12, tcDouble100, false);
	  Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);
  	  Tob_Ligne.PutValue       ('GL_TYPEDIM'    , 'DIM');
      Tob_Ligne.PutValue       ('GL_QUALIFMVT'  , QualifMvt) ;

	  Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
 	  if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then Prix:=Tob_Ligne.GetValue('GL_PUHTDEV')
      else Prix:=Tob_Ligne.GetValue('GL_PUTTCDEV') ;
  	  TotQte      := TotQte+Qte;
      TotQteReste := TotQteReste+Qte;
	  TotPrix     := TotPrix+Qte*Prix;
	end;
  end;
end;

///////////////////////////////////////////////////////////////////
// Récupération des transferts inter-boutiques GB 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupTransfert(st : string);
var CodeArt 	     : string    ;
    Rang    	     : string    ;
    CategorieDim     : string    ;
    NaturePiece      : string    ;
    Etab             : string    ;
    SQL              : string    ;
    CodeDim1,CodeDim2: string    ;
    CodeDim3,CodeDim4: string    ;
    CodeDim5 	     : string    ;
    RefRec   	     : string    ;		// Référence commande = N° GB + Réf Cde GB
    CodeTiers	     : string    ;
    GL_CodeArticle   : string    ;
    DevisePiece      : string    ;
    CodeDevGB        : string    ;
    CodeDevPGI       : string    ;
    Qte, Prix	     : double    ;
    qteRec	         : double    ;
    Montant          : double    ;
    MontantDev       : double    ;
    MontantOld       : double    ;
    Q                : TQUERY    ;
    Per,Sem          : integer   ;
    i		         : integer   ;
    DD , DateConv    : TDateTime ;
    Taille           : integer   ;
begin

  NaturePiece := Tob_Param.GetValue ('NATTRFEM');
  Etab := copy (st, 7, 3);

  if ArretRecup then exit;

  if Copy(st,1,5)='REEC1' then
  begin
    if (Tob_Piece <> Nil) then ImportTransfert ;

    EcrireInfo ('   ' + 'Boutique ' + Copy(st,7,3) + ' Transfert ' + Copy (st,11, 6), True, True, False);

    //////////////////////////////////////////////////////////////////////
    // Initialisations diverses et varièes
    //////////////////////////////////////////////////////////////////////
    TotQte:=0.0; TotPrix:=0.0;TotQteReste:=0.0;
    Numligne := 0;            // Initialisation du premier numéro de ligne
    NombreErreurs(0);					// Initialisation du nombre d'erreurs
    QualifMvt:=GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
    LigneCommACreer:=False;

    Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
    AddLesSupEntete (Tob_Piece);
    //Tob_Piece.initValeurs;

    ////////////////////////////
    // Nature de pièce + Date
    ////////////////////////////
    Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
    Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 17, 8, tcDate8AMJ, false);
    Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 17, 8, tcDate8AMJ, false);
    /////////////////////////
    // Souche, Numéro, .....
    ////////////////////////
    CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
    CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
    CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');

    PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
    PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
    PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
    //////////////////////////////////////////////////////
    // Référence du transferts
    //////////////////////////////////////////////////////
    RefRec := 'Btq ' + copy (st, 7, 3) + ' N°' + copy (st, 11, 6);
    PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefRec) ;
    //////////////////////////////////////////////////////////////
    // Code Tiers : Code tiers PGI par défaut pour les transferts
    //////////////////////////////////////////////////////////////
    CodeTiers := GetParamSoc ('SO_GCTIERSDEFAUT');
    Tob_Piece.PutValue ('GP_TIERS'       , CodeTiers);
    Tob_Piece.PutValue ('GP_TIERSLIVRE'  , CodeTiers);
    Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeTiers);
    Tob_Piece.PutValue ('GP_TIERSPAYEUR' , CodeTiers);
    ////////////////////////////////////////////////////////////////
    // Initialise les zones de l'entête en fonction du fournisseur
    ////////////////////////////////////////////////////////////////
    Tob_Piece.PutValue('GP_FACTUREHT'	   , 'X');   // Les transferts GB 2000 sont valorisés en HT ....
    Tob_Piece.PutValue('GP_ESCOMPTE'       , Tob_Tiers_Defaut.GetValue('T_ESCOMPTE')) ;
    Tob_Piece.PutValue('GP_MODEREGLE'      , Tob_Tiers_Defaut.GetValue('T_MODEREGLE')) ;
    Tob_Piece.PutValue('GP_REGIMETAXE'     , Tob_Tiers_Defaut.GetValue('T_REGIMETVA')) ;
    Tob_Piece.PutValue('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers_Defaut ));;
    Tob_Piece.PutValue('GP_QUALIFESCOMPTE' , Tob_Tiers_Defaut.GetValue('T_QUALIFESCOMPTE')) ;
    Tob_Piece.PutValue('GP_TARIFTIERS'     , Tob_Tiers_Defaut.GetValue('T_TARIFTIERS')) ;

    if (NombreErreurs (-1) = 0) then
    begin
      ///////////////////////////////////////////////////////////////////////
      // Emetteur
      ///////////////////////////////////////////////////////////////////////
      Tob_Piece.PutFixedStValue('GP_ETABLISSEMENT', st, 7, 3, tcTrim, false);
      Tob_Piece.PutFixedStValue('GP_DEPOT'        , st, 7, 3, tcTrim, false);
      ///////////////////////////////////////////////////////////////////////
      // Récepteur
      ///////////////////////////////////////////////////////////////////////
      Tob_Piece.PutFixedStValue('GP_DEPOTDEST', st, 27, 3, tcTrim, false);
      ///////////////////////////////////////////////
      // Champs par  défaut
      //////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_VIVANTE'       , '-');       // Pièce vivante    : Oui
      Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
      Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
      Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'TRF');     // Document         : Transfert
      Tob_Piece.PutValue ('GP_EDITEE'        , 'X');       // Pièce éditée     : Oui
      Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
      Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
      Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_SOCIETE'       , V_PGI.CodeSociete);
      Tob_Piece.PutValue ('GP_DATECREATION'   , Date) ;
      Tob_Piece.PutValue ('GP_HEURECREATION'  , NowH) ;
      //Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;

      ////////////////////////////////////////////
      // Devise du document
      /////////////////////////////////////////////
      CodeDevGB  := Trim (copy (st, 74, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      DevisePiece:=CodedevPGI ;

      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      PutValueDetail (Tob_Piece, 'GP_TAUXDEV'    ,DEV.Taux) ;
      AttribCotation (Tob_Piece) ;
      PutValueDetail (Tob_Piece, 'GP_DATETAUXDEV',DEV.DateTaux) ;

      DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
      Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
      Tob_Piece.PutValue ('GP_PERIODE',Per) ;
      Tob_Piece.PutValue ('GP_SEMAINE',Sem) ;
    end;
  end
  else if Copy(st,1,5)='RELC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      /////////////////////////////////////////////////////////////////
      // Création de la ligne commentaire précédente ???
      /////////////////////////////////////////////////////////////////
      if (LigneCommACreer = True) then
      begin
      	LigneCommACreer:=False;
        CreeLigneCommentaire;
      end;
     	/////////////////////////////////////////////////////////////////
     	// gestion d'une ligne commentaire avec l'article générique ????
     	/////////////////////////////////////////////////////////////////
     	NumLigne        := NumLigne+1;
     	NumLigneComment := Numligne;
     	LigneCommACreer := True;

        TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
     	////////////////////////////////////////////////
     	// formatage de la référence article générique
     	////////////////////////////////////////////////
     	CodeArticle := Copy (st, 31, 12) ;
     	codeart := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
     	//////////////////////////////////////////
     	// Récupération du libellé de l'article
     	//////////////////////////////////////////
     	//SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
     	//Q:=OpenSQL(SQL,True) ;
     	//if Not Q.EOF then	Designation := Q.FindField('GA_LIBELLE').AsString
     	//else begin
       	//  EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,31,12), True, True, True);
      	//  NombreErreurs (1);
     	//end;
    	//Ferme(Q) ;
    end;
  end
  else if Copy(st,1,5)='RELC2' then
  begin
    // Rien à faire
  end
  else if Copy(st,1,5)='RETC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      ////////////////////////////////////////////////////////////////
      // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
      ////////////////////////////////////////////////////////////////
      Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
      //Tob_Ligne.initValeurs;
      ////////////////////////////////////////////////////////////////
      // Initialise les champs de la ligne à partir de l'entête
      ////////////////////////////////////////////////////////////////
      PieceVersLigne (Tob_Piece, Tob_Ligne);
      Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
      Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
      //////////////////////////////////////////////
      // Nouveau numéro de ligne
      //////////////////////////////////////////////
      inc (NumLigne);
      Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
      Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
      ////////////////////////////////////////////////////////////////
      // Ajoute des champs supplémentaires pour le calcul des cumuls
      ////////////////////////////////////////////////////////////////
      AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
      //////////////////////////////////////////////////////////
      // Détermination de l'article taille
      /////////////////////////////////////////////////////////
      CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

      Taille := StrToInt (copy (st, 21 ,2));
      rang   := CalculTaillePGI (Taille);

      CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
      CodeDim4 := '   '; CodeDim5 := '   ';
      if (CategorieDim [3] = '1') then Codedim1 :=rang
      else if (CategorieDim [3] = '2') then Codedim2 :=rang
      else if (CategorieDim [3] = '3') then Codedim3 :=rang
      else if (CategorieDim [3] = '4') then Codedim4 :=rang
      else if (CategorieDim [3] = '5') then Codedim5 :=rang;
      Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        if (Tob_Article=nil) then
      	begin
       	  Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     	  //Tob_Article.initValeurs;
      	end;
	/////////////////////////////////////////////////////////////////////////////
      	// Initialisation des champs de la ligne à partir de l'article dimensionné
      	/////////////////////////////////////////////////////////////////////////////
        Tob_Ligne.PutValue ('GL_LIBELLE'       , Q.FindField('GA_LIBELLE').AsString);
      	Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
        Tob_Ligne.PutValue ('GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
        Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
        Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	    Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	    Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
      	for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
        for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
        Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
      	/////////////////////////////////////////////////////////////////////////////
      	// Ajout de l'article dans la TOB Article
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Article.SelectDB('',Q);
    	end
    	else begin
      	  EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	  NombreErreurs (1);
     	end;
    	Ferme(Q) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Ligne.PutValue       ('GL_ARTICLE'     , Codeart);
      //Tob_Ligne.PutValue       ('GL_LIBELLE'     , Designation);

      GL_CodeArticle := Trim (CodeArticle);
      Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
      Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

      Tob_Ligne.PutValue ('GL_TARIF'       , 0);
      Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
      Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
      Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
      Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

      QteRec := StrtoFloat (copy (st, 23, 4));
      Tob_Ligne.PutValue ('GL_QTESTOCK', QteRec);
      Tob_Ligne.PutValue ('GL_QTEFACT' , QteRec);
      Tob_Ligne.PutValue ('GL_QTERESTE', QteRec);

      //////////////////////////////////////////////////////////
      // récupération des PA et PR
      //////////////////////////////////////////////////////////
      // -> devise des prix ?
      DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
      begin
        if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
        else                  DevisePiece := 'EUR';
      end;
      //
      // PA de l'article
      //
      Montant := StrToFloat (copy (st, 27, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        // rien à faire
      end else
      begin
        if  DevisePiece = V_PGI.Devisefongible then
        begin
          if VH^.TenueEuro then
          begin
            MontantOld := Montant ;
            Montant    := PivotToEuro (MontantOld)
          end ;
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Montant := MontantDev ;
        end;
      end;
      Tob_Ligne.PutValue('GL_DPA'      , Montant);
      Tob_Ligne.PutValue('GL_PMAP'     , Montant);
      Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
      Tob_Ligne.PutValue('GL_DPR'      , Montant);
      Tob_Ligne.PutValue('GL_PMRP'     , Montant);
      Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
      //
      // Valorisation du transfert
      //
	  //Tob_Ligne.PutFixedStValue('GL_PUHT'    , st, 27, 12, tcDouble100, false);
  	  Tob_Ligne.PutFixedStValue('GL_PUHTDEV' , st, 27, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTNET' , st, 27, 12, tcDouble100, false);

	  Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);
  	  Tob_Ligne.PutValue       ('GL_TYPEDIM'    , 'DIM');
      Tob_Ligne.PutValue       ('GL_QUALIFMVT'  , QualifMvt) ;
      //
      // Cumul pour ligne commentaire
      //
	  Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
 	  Prix:=Tob_Ligne.GetValue('GL_PUHTDEV') ;

  	  TotQte      :=TotQte+Qte;
      TotQteReste :=TotQteReste+Qte;
	  TotPrix     :=TotPrix+Qte*Prix;
	end;
  end;
end;


///////////////////////////////////////////////////////////////////
// Récupération des E/S de stocks GB 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupESStock (st : string);
var CodeArt 	     : string;
    Rang    	     : string;
    CategorieDim     : string;
    NaturePiece      : string;
    Etab             : string;
    SQL              : string;
    CodeDim1,CodeDim2: string;
    CodeDim3,CodeDim4: string;
    CodeDim5 	     : string;
    RefRec   	     : string;		// Référence commande = N° GB + Réf Cde GB
    CodeFrn	         : string;
    GL_CodeArticle   : string;
    DevisePiece      : string;
    CodeDevPGI       : string;
    CodeDevGB        : string;
    Montant          : double;
    MontantOld       : double;
    MontantDev       : double;
    Qte, Prix	     : double;
    QteES 	         : double;
    Q                : TQUERY ;
    Per,Sem          : integer;
    i		         : integer;
    Taille           : integer;
    DD   	         : TDateTime ;
    DateConv         : TDateTime;
begin

  NatureEntree := Tob_Param.GetValue ('NATENT');
  NatureSortie := Tob_Param.GetValue ('NATSOR');
  Etab := copy (st, 7, 3);

  if ArretRecup then exit;
  if Copy(st,1,5)='REEC1' then
  begin
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Importation de l'E/S précédente
    ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    if (Tob_Piece <> Nil) then ImportESStock;

    EcrireInfo ('   ' + 'Boutique ' + Copy(st,7,3) + ' E/S de stock ' + Copy (st,11, 6), True, True, False);

    ////////////////////////////////
    // Entrée ou sortie ?
    ////////////////////////////////
    if (copy (st, 52, 1) = '4') then
    begin
    	EntreeStock := True;
      NaturePiece := NatureEntree;
    end
    else if (copy (st, 52, 1) = '5') then
    begin
      EntreeStock := False;
      NaturePiece := NatureSortie;
    end
    else begin
      EcrireInfo ('   Cette pièce n''est ni une entrée, ni une sortie de stock', True, True, True);
      NombreErreurs (1);
    end;
    ///////////////////////////////
    // Récupération du motif E/S
    ///////////////////////////////
    Motif_ES := copy (st, 80, 3) ;

    if NombreErreurs (-1) = 0 then
    begin
      //////////////////////////////////////////////////////////////////////
      // Initialisations diverses et varièes
      //////////////////////////////////////////////////////////////////////
      TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
      Numligne := 0;            // Initialisation du premier numéro de ligne
      NombreErreurs(0);					// Initialisation du nombre d'erreurs
      QualifMvt:=GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
      LigneCommACreer:=False;

      Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
      AddLesSupEntete ( Tob_Piece);
      //Tob_Piece.initValeurs;
      ////////////////////////////
      // Nature de pièce + Date
      ////////////////////////////
      Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
      Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 17, 8, tcDate8AMJ, false);
      Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 17, 8, tcDate8AMJ, false);
      /////////////////////////
      // Souche, Numéro, .....
      ////////////////////////
      CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
      CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
      CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');

      PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
      PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
      PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
      //////////////////////////////////////////////////////
      // Référence de commande
      //////////////////////////////////////////////////////
      RefRec := 'Btq ' + copy (st, 7, 3) + 'N°' + copy (st, 11, 6);
      PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefRec) ;
      //////////////////////////////////////////////////////////////
      // Code Tiers : Code tiers PGI par défaut pour les transferts
      //////////////////////////////////////////////////////////////
      CodeFrn := Tob_Tiers_Defaut.GetValue('T_TIERS') ;
      Tob_Piece.PutValue ('GP_TIERS'       , CodeFrn);
      Tob_Piece.PutValue ('GP_TIERSLIVRE'  , CodeFrn);
      Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeFrn);
      Tob_Piece.PutValue ('GP_TIERSPAYEUR' , CodeFrn);

      ////////////////////////////////////////////////////////////////
      // Initialise les zones de l'entête en fonction du fournisseur
      ////////////////////////////////////////////////////////////////
      Tob_Piece.PutValue('GP_FACTUREHT'	   , 'X');   // Les transferts GB 2000 sont valorisés en HT ....
      Tob_Piece.PutValue('GP_ESCOMPTE'       , Tob_Tiers_Defaut.GetValue('T_ESCOMPTE')) ;
      Tob_Piece.PutValue('GP_MODEREGLE'      , Tob_Tiers_Defaut.GetValue('T_MODEREGLE')) ;
      Tob_Piece.PutValue('GP_REGIMETAXE'     , Tob_Tiers_Defaut.GetValue('T_REGIMETVA')) ;
      Tob_Piece.PutValue('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers_Defaut ));;
      Tob_Piece.PutValue('GP_QUALIFESCOMPTE' , Tob_Tiers_Defaut.GetValue('T_QUALIFESCOMPTE')) ;
      Tob_Piece.PutValue('GP_TARIFTIERS'     , Tob_Tiers_Defaut.GetValue('T_TARIFTIERS')) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Piece.PutFixedStValue('GP_ETABLISSEMENT', st, 7, 3, tcTrim, false);
      Tob_Piece.PutFixedStValue('GP_DEPOT', st, 7, 3, tcTrim, false);

      //////////////////////////////////////////////
      // Champs par  défaut
      //////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_VIVANTE'       , 'X');       // Pièce vivante    : Oui
      Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
      Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
      Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'ACH');     // Document         : Achat
      Tob_Piece.PutValue ('GP_EDITEE'        , 'X');       // Pièce éditée     : Oui
      Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
      Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
      Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
//      Tob_Piece.PutValue ('GP_SOCIETE'       ,V_PGI.NomSociete);
      Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers_Defaut.GetValue('T_ZONECOM'))) ;


      ////////////////////////////////////////////
      // Devise du document
      /////////////////////////////////////////////
      CodeDevGB  := Trim (copy (st, 74, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      DevisePiece:=CodeDevPGI;
      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
           Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
           Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      PutValueDetail(Tob_Piece, 'GP_TAUXDEV'    ,DEV.Taux) ;
      AttribCotation(Tob_Piece) ;
      PutValueDetail(Tob_Piece, 'GP_DATETAUXDEV',DEV.DateTaux) ;

      DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
      Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
      Tob_Piece.PutValue('GP_PERIODE',Per) ;
      Tob_Piece.PutValue('GP_SEMAINE',Sem) ;
    end;
  end
  else if Copy(st,1,5)='RELC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      /////////////////////////////////////////////////////////////////
      // Création de la ligne commentaire précédente ???
      /////////////////////////////////////////////////////////////////
      if (LigneCommACreer = True) then
      begin
      	LigneCommACreer:=False;
        CreeLigneCommentaire;
      end;
     	/////////////////////////////////////////////////////////////////
     	// gestion d'une ligne commentaire avec l'article générique ????
     	/////////////////////////////////////////////////////////////////
     	NumLigne        := NumLigne+1;
     	NumLigneComment := Numligne;
     	LigneCommACreer := True;

        TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
     	////////////////////////////////////////////////
     	// formatage de la référence article générique
     	////////////////////////////////////////////////
     	CodeArticle := Copy (st, 31, 12) ;
     	codeart := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
     	//////////////////////////////////////////
     	// Récupération du libellé de l'article
     	//////////////////////////////////////////
     	SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
     	Q:=OpenSQL(SQL,True) ;
     	if Not Q.EOF then Designation := Q.FindField('GA_LIBELLE').AsString
     	else begin
       	EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,31,12), True, True, True);
      	NombreErreurs (1);
     	end;
    	Ferme(Q) ;
    end;
  end
  else if Copy(st,1,5)='RELC2' then
  begin
    // Rien à faire
  end
  else if Copy(st,1,5)='RETC1' then
  begin
    if (NombreErreurs (-1) = 0) then
    begin
      ////////////////////////////////////////////////////////////////
      // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
      ////////////////////////////////////////////////////////////////
      Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
      //Tob_Ligne.initValeurs;
      ////////////////////////////////////////////////////////////////
      // Initialise les champs de la ligne à partir de l'entête
      ////////////////////////////////////////////////////////////////
      PieceVersLigne (Tob_Piece, Tob_Ligne);
      Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
      Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
      //////////////////////////////////////////////
      // Nouveau numéro de ligne
      //////////////////////////////////////////////
      NumLigne := NumLigne+1;
      Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
      Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
      ////////////////////////////////////////////////////////////////
      // Ajoute des champs supplémentaires pour le calcul des cumuls
      ////////////////////////////////////////////////////////////////
      AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
      //////////////////////////////////////////////////////////
      // Détermination de l'article taille
      /////////////////////////////////////////////////////////
      CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

      Taille := StrToInt (copy (st, 21 ,2));
      rang   := CalculTaillePGI (Taille);

      CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
      CodeDim4 := '   '; CodeDim5 := '   ';
      if (CategorieDim [3] = '1') then Codedim1 :=rang
      else if (CategorieDim [3] = '2') then Codedim2 :=rang
      else if (CategorieDim [3] = '3') then Codedim3 :=rang
      else if (CategorieDim [3] = '4') then Codedim4 :=rang
      else if (CategorieDim [3] = '5') then Codedim5 :=rang;
      Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        if (Tob_Article=nil) then
      	begin
       	  Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     	  //Tob_Article.initValeurs;
      	end;
	/////////////////////////////////////////////////////////////////////////////
      	// Initialisation des champs de la ligne à partir de l'article dimensionné
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
        Tob_Ligne.PutValue( 'GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
        Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
        Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	    Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	    Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
      	for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
      	/////////////////////////////////////////////////////////////////////////////
      	// Ajout de l'article dans la TOB Article
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Article.SelectDB('',Q);
      end
      else begin
      	EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
	  Tob_Ligne.PutValue       ('GL_ARTICLE'     , Codeart);
  	  Tob_Ligne.PutValue       ('GL_LIBELLE'     , Designation);

      GL_CodeArticle := Trim (CodeArticle);
      Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	  Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

  	  Tob_Ligne.PutValue ('GL_TARIF'       , 0);
      Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	  Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
      Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
      Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');
      Tob_Ligne.PutValue ('GL_MOTIFMVT'    , Motif_ES) ;

      // Entrée ou sortie ?????,
      QteES := StrToFloat (copy (st, 23, 4));
      //if EntreeStock = False then QteES := -QteES;
  	  Tob_Ligne.PutValue('GL_QTESTOCK', QteES);
      Tob_Ligne.PutValue('GL_QTEFACT' , QteES);
      Tob_Ligne.PutValue('GL_QTERESTE', QteES);

	    //////////////////////////////////////////////////////////
  	  // récupération des PA et PR
    	//////////////////////////////////////////////////////////
      // -> devise des prix ?
      DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
      begin
        if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
        else                    DevisePiece := 'EUR';
      end;
      //
      // PA de l'article
      //
      //
      // PA de l'article
      //
      Montant := StrToFloat (copy (st, 27, 12));
      Montant := Montant/100 ;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        // Rien à faire
      end else
      begin
        if  DevisePiece = V_PGI.Devisefongible then
        begin
          if VH^.TenueEuro then
          begin
            MontantOld := Montant ;
            Montant    := PivotToEuro (MontantOld)
          end;
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Montant := MontantDev ;
        end;
      end;
      Tob_Ligne.PutValue('GL_DPA'      , Montant);
      Tob_Ligne.PutValue('GL_PMAP'     , Montant);
      Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
      Tob_Ligne.PutValue('GL_DPR'      , Montant);
      Tob_Ligne.PutValue('GL_PMRP'     , Montant);
      Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
      //
      // Valorisation du transfert
      //
	  //Tob_Ligne.PutFixedStValue('GL_PUHT'    , st, 27, 12, tcDouble100, false);
  	  Tob_Ligne.PutFixedStValue('GL_PUHTDEV' , st, 27, 12, tcDouble100, false);
      Tob_Ligne.PutFixedStValue('GL_PUHTNET' , st, 27, 12, tcDouble100, false);

	  Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);
  	  Tob_Ligne.PutValue       ('GL_TYPEDIM'    , 'DIM');
      Tob_Ligne.PutValue       ('GL_QUALIFMVT'  , QualifMvt) ;
      //
      // cumul des prix et quantité pour ligne commentaire
      //
	  Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
 	  Prix:=Tob_Ligne.GetValue('GL_PUHTDEV') ;

  	  TotQte      :=TotQte+Qte;
      TotQteReste :=TotQteReste+Qte;
	  TotPrix     :=TotPrix+Qte*Prix;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////
// Récupération des ventes CASH 2000 ou GB 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupVente (st : string);
var CodeArt 	     : string;
    Rang    	     : string;
    NaturePiece      : string;
    Etab             : string;
    DatePiece	     : string;
    NumTicket	     : string;
    SQL              : string;
    CodeDim1,CodeDim2: string;
    CodeDim3,CodeDim4: string;
    CodeDim5 	     : string;
    RefVte	         : string;		// Référence commande = N° GB + Réf Cde GB
    RefCli	         : string;
    CodeCli	         : string;
    CodeDevGB        : string ;
    CodeDevPGI       : string ;
    CodeTVAGB        : string ;
    CodeTVAPGI       : string ;
    RegimePGI        : string ;
    CodeVendeur      : string ;
    GL_CodeArticle   : string;
    DevisePiece      : string ;
    Montant          : double;
    MontantDev       : double;
    MontantOld       : double;
    Q                : TQUERY ;
    Per,Sem          :integer ;
    i		         : integer ;
    Taille           : integer ;
    DD, DateConv     : TDateTime;
begin

  if ArretRecup then exit;

  NaturePiece := Tob_Param.GetValue ('NATVTE');
  Etab        := copy (st,  7, 3);
  DatePiece   := copy (st, 10, 8);
  NumTicket   := copy (st, 18, 6);

  if Copy(st,1,5)='VTEC1' then
  begin
    /////////////////////////////////////////////////////////////////////////
    // Intégration du ticket précédent
    ////////////////////////////////////////////////////////////////////////
    if ( (Tob_Piece <> Nil) and ( (NumEtabVte<>Etab) or (DatePieceVte<>DatePiece) or (NumPieceVte<>NumTicket))) then
    begin
      ImportPiece ;
      //EcrireInfo ('   '+'Boutique '+ Copy(st,7,3) + ' Journée '+Copy (st,16,2)+'/'+copy (st,14,2)+'/'+copy (st,10,4)+ ' Ticket ' + copy (st,18,6), False, True, False);
    end;

    Inc(Compteur);
    if Compteur=10000 then
    begin
      EcrireInfo (' Le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) + ' : ' + IntToStr (Compteur) + ' ventes intégrées', True, False, False);
      Compteur:=0 ;
    end ;

    NumEtabVte   := Etab;
    DatePieceVte := DatePiece;
    NumPieceVte  := Numticket;

    if (Tob_Piece = nil) then
    begin
      //////////////////////////////////////////////////////////////////////
      // Nouveau Ticket : Initialisations diverses et varièes
      //////////////////////////////////////////////////////////////////////
      //TotQte:=0.0; TotPrix:=0.0;
      Numligne := 0;            // Initialisation du premier numéro de ligne
      NombreErreurs(0);					// Initialisation du nombre d'erreurs

      LigneCommACreer:=False;

      Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
//      Tob_Piece.initValeurs;
      AddLesSupEntete ( Tob_Piece);
   	  InitTobPiece (Tob_Piece);			// Init date de création, modif, utilisateur ......

      Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
      Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 10, 8, tcDate8AMJ, false);
      Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 10, 8, tcDate8AMJ, false);
      /////////////////////////
      // Souche, Numéro, .....
      ////////////////////////
      CleDoc.Souche      := GetSoucheG(NaturePiece, Etab,'') ;
      CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche) ;
      CleDoc.DatePiece   := Tob_Piece.GetValue('GP_DATEPIECE');

      PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
      PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
      PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
      ////////////////////////////////////////////////////////////////////////
		  // Référence de la vente : Boutique + Date de vente + N° ticket
      // Attention : Si cette référence doit être modifiée, il faudra penser
      // 						 à modifier la recherche de la vente dans lé récupération
      //						 des règlements
      ///////////////////////////////////////////////////////////////////////
      RefVte := 'Btq ' + copy (st, 7,3) + ' le ' + copy (st, 16, 2) + '/' + copy (st, 14, 2) + '/' + copy (st, 10,4) + ' Ticket ' + copy (st, 18, 6);
      PutValueDetail(Tob_Piece, 'GP_REFINTERNE', RefVte) ;
      ///////////////////////////////////////////////////////////////
      // Client
      // Référence client PGI = Code BTQ + Code Client ou Code Client
      // Si pas de client GB 2000, utilisation du client par défaut
      ////////////////////////////////////////////////////////////////
      if copy (st, 38, 6) <> '      ' then
      begin
	    RefCli := Tob_Param.GetValue ('CODECLI');
  	    if RefCli='0' then CodeCli:=copy(st, 7, 3) + copy(st,38,6)
    	else if RefCli='1' then CodeCli:=copy(st,38,6);

        ///////////////////////////////////////////////////////////////
        // Chargement en TOB du client
        ///////////////////////////////////////////////////////////////
      	Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCli + '" AND T_NATUREAUXI = "CLI" ',False) ;
	    if not Q.EOF then
      	begin
     	  Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
//          Tob_Tiers.initValeurs;
    	  Tob_Tiers.SelectDB('',Q);
          Ferme(Q) ;
        end else
      	begin
	      Ferme(Q) ;
    	  EcrireInfo ('   Erreur -> Problème pour retrouver le client '+ CodeCli, True, True, True);
          EcrireInfo ('          -> La vente est rattachée au client par défaut ' + Tob_Param.GetValue ('CODECLIVTE'), True, True, False);
          Tob_Tiers := TOB.Create ('TIERS', nil, -1);
   	      Tob_Tiers.Dupliquer (Tob_Tiers_Defaut, False, True, True);
          CodeCli := Tob_Param.GetValue ('CODECLIVTE');
      	end;
      end else 				// Client GB 2000 non renseigné, on prend le client par défaut
      begin
      	CodeCli := Tob_Param.GetValue ('CODECLIVTE');
        Tob_Tiers := TOB.Create ('TIERS', nil, -1);
   	    Tob_Tiers.Dupliquer (Tob_Tiers_Defaut, False, True, True);
      end;

      Tob_Piece.PutValue ('GP_TIERS'        , CodeCli) ;
      Tob_Piece.PutValue ('GP_TIERSLIVRE'   , CodeCli) ;
      Tob_Piece.PutValue ('GP_TIERSFACTURE' , CodeCli) ;
      Tob_Piece.PutValue ('GP_TIERSPAYEUR'  , CodeCli) ;
      ////////////////////////////////////////////////////////////////
      // Initialise les zones de l'entête en fonction du client
      ////////////////////////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_ESCOMPTE'       , Tob_Tiers.GetValue('T_ESCOMPTE')) ;
      ///////////////////////////////////////////////////////////////////////////////////
      // Plutôt que d'initialiser le mode de règlement avec celui du client (qui est en
      // fait le mode de règlement du dossier, donc pas terrible), on initialise avec
      // le mode de règlement paramétré dans la récupération GB 2000
      ///////////////////////////////////////////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_MODEREGLE'      , Tob_Param.GetValue ('MODEREGLVTE'));
      CodeTVAGB := Copy (st, 70, 1) ;
      CodeTVAPGI := TraduitGPAOenPGI(TypeRepriseGB2000, '$$_CODETVA', CodeTVAGB, Tob_TVA);
      RegimePGI := ReadTokenST(CodeTVAPGI) ;
      if RegimePGI = '' then RegimePGI := Tob_Tiers.GetValue('T_REGIMETVA') ;
      Tob_Piece.PutValue ('GP_REGIMETAXE'     , RegimePGI) ;
      Tob_Piece.PutValue ('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers));
      Tob_Piece.PutValue ('GP_QUALIFESCOMPTE' , Tob_Tiers.GetValue('T_QUALIFESCOMPTE')) ;
      Tob_Piece.PutValue ('GP_TARIFTIERS'     , Tob_Tiers.GetValue('T_TARIFTIERS')) ;

      Tob_Piece.PutFixedStValue('GP_ETABLISSEMENT', st, 7, 3, tcTrim, false);
      Tob_Piece.PutFixedStValue('GP_DEPOT'        , st, 7, 3, tcTrim, false);
      ///////////////////////////////////////////////
      // Champs par  défaut
      //////////////////////////////////////////////
      Tob_Piece.PutValue ('GP_FACTUREHT'	   , '-');      // Facturation TTC
      Tob_Piece.PutValue ('GP_VIVANTE'       , '-');       // Pièce vivante    : Oui
      Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
      Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
      Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'VEN');     // Document         : Achat
      Tob_Piece.PutValue ('GP_EDITEE'        , 'X');       // Pièce éditée     : Oui
      Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
      Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
      if copy (st, 68, 2) <> '  ' then
      begin
        CodeVendeur := copy (st, 7, 3) + copy (st, 68, 2);
        Tob_Piece.PutValue ('GP_REPRESENTANT'  , CodeVendeur) ;
      end else Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;
      Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
      Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
//      Tob_Piece.PutValue ('GP_SOCIETE'       ,V_PGI.NomSociete);

      DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
      Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
      Tob_Piece.PutValue ('GP_PERIODE', Per) ;
      Tob_Piece.PutValue ('GP_SEMAINE', Sem) ;
    end;

    ///////////////////////////////////////////////
    // Formatage de la référence article générique
    // Cas 1 : Vente normale avec un article
    // Cas 2 : Vente diverse ou opération de caisse
    ////////////////////////////////////////////////
    if copy (st, 66, 1) = 'N' then
    begin
    	CodeArticle := Copy (st, 53, 12) ;
    	CodeArt     := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
    end else  // Vente diverse ou opération de caisse
    begin
      CodeArticle := Copy (st, 56, 2) ;
      CodeArt     := CodeArticle + '          ' + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
    end;
    //////////////////////////////////////////////////////////////////////////////
    // Test si l'article générique existe et récupération du libellé de l'article
    //////////////////////////////////////////////////////////////////////////////
    // SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
    // Q:=OpenSQL(SQL,True) ;
    // if Not Q.EOF then Designation := Q.FindField('GA_LIBELLE').AsString
    // else begin
    //   EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ CodeArt, True, True, True);
    //   NombreErreurs (1);
    // end;
    // Ferme(Q) ;

    ////////////////////////////////////////////////////////////////
    // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    // avec l'article dimensionné
    ////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
//      Tob_Ligne.initValeurs;
      ////////////////////////////////////////////////////////////////
      // Initialise les champs de la ligne à partir de l'entête
      ////////////////////////////////////////////////////////////////
      PieceVersLigne (Tob_Piece, Tob_Ligne);
      Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
      Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
      //////////////////////////////////////////////
      // Nouveau numéro de ligne
      //////////////////////////////////////////////
      Inc(NumLigne);// := NumLigne+1;
      Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
      Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
      ////////////////////////////////////////////////////////////////
      // Ajoute des champs supplémentaires pour le calcul des cumuls
      ////////////////////////////////////////////////////////////////
      AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
      ///////////////////////////////////////////////////////////////
      // Détermination de l'article taille :
      //	Vente normal                         : Référence + Taille
      //  Vente diverse ou Opération de caisse : Référence
      //////////////////////////////////////////////////////////////
      if St[66]='N' then
      begin
	    //CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

        Taille := StrToInt (copy (st, 77 ,2));
        rang   := CalculTaillePGI (Taille);

        CodeDim1     := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
        CodeDim4     := '   '; CodeDim5 := '   ';
        if      (CategorieDimension [3] = '1') then Codedim1 :=rang
        else if (CategorieDimension [3] = '2') then Codedim2 :=rang
        else if (CategorieDimension [3] = '3') then Codedim3 :=rang
        else if (CategorieDimension [3] = '4') then Codedim4 :=rang
        else if (CategorieDimension [3] = '5') then Codedim5 :=rang;
        Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X' // Code article PGI
      end else CodeArt := CodeArticle + '          ' + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';

      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';

      Q:=OpenSQL(SQL,True) ;

      ///////////////////////////////////////////////////////////////////////////////////////////
      // MODIF LM 2707/02
      // L'article n'existe plus, on le remplace par l'article par défaut pour assurer un bon CA
      ///////////////////////////////////////////////////////////////////////////////////////////
      if Q.EOF then
      begin
        Ferme (Q) ;
        CodeArt     := Tob_Param.GetValue ('CODEARTVTE');
        CodeArticle := copy (CodeArt, 1, 18) ;

        SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
        Q:=OpenSQL(SQL,True) ;
      end;
      // FIN DES MODIFS LM 29/07/02

      if Not Q.EOF then
      begin
    	if (Tob_Article=nil) then
      	begin
       	  Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
//     	  Tob_Article.initValeurs;
      	end;
     	/////////////////////////////////////////////////////////////////////////////
      	// Initialisation des champs de la ligne à partir de l'article dimensionné
      	/////////////////////////////////////////////////////////////////////////////
        Tob_Ligne.PutValue ('GL_LIBELLE'       , Q.FindField('GA_LIBELLE').AsString);
      	Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
        Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	    Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	    Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
        Tob_Ligne.PutValue( 'GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
        Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;

      	for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
      	Tob_Ligne.PutValue ('GL_LIBREARTA'  , Q.FindField('GA_LIBREARTA').AsString);

        ///////////////////////////////////////////////////////////////////////////
        // MODIF Temporaire LM 05/12/2001 pour MANCINI
        // intégration des ventes US hors taxe -> Code TVA GB 2000 = 0
        // Prévoir une table de correspondance des taxes GB / PGI
        ///////////////////////////////////////////////////////////////////////////
        //if copy (st, 70, 1) <> '0' then
        //begin
        //  for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
        //end;
        if CodeTVAPGI = '' then CodeTVAPGI := Q.FindField('GA_FAMILLETAXE1').asString ;
        Tob_Ligne.PutValue('GL_FAMILLETAXE1', CodeTVAPGI) ;

      	/////////////////////////////////////////////////////////////////////////////
      	// Ajout de l'article dans la TOB Article
      	/////////////////////////////////////////////////////////////////////////////
      	Tob_Article.SelectDB('',Q);
      end
      else begin
       	EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article '+ Codeart, True, True, True);
       	NombreErreurs (1);
      end;
      Ferme(Q) ;
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      Tob_Ligne.PutValue ('GL_ARTICLE'     , CodeArt);
      //Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

      if copy (st, 66, 1) = 'N' then GL_CodeArticle := copy (st, 53, 12)
      else GL_CodeArticle := copy (st, 56, 2);
      GL_CodeArticle      := Trim (CodeArticle);
      Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
      Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

      Tob_Ligne.PutValue ('GL_TARIF'       , 0);
      Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');

      if copy (st, 66, 1) = 'D' then Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'PRE')
      else if copy (st, 66, 1) = 'O' then Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'FI')    else
      Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR') ;

      Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'AFF') ;
      Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

      // Gestion de la quantité vendue
      QuantiteVendueGB := StrToFloat (copy (st, 79, 4));
      if copy (st, 65, 1) = '-' then QuantiteVenduePGI := -QuantiteVendueGB
      else QuantiteVenduePGI := QuantiteVendueGB;
      Tob_Ligne.PutValue('GL_QTESTOCK' , QuantiteVenduePGI);
      Tob_Ligne.PutValue('GL_QTEFACT'  , QuantiteVenduePGI);
      Tob_Ligne.PutValue('GL_QTERESTE' , QuantiteVenduePGI);

      //////////////////////////////////////////////////////////
      // récupération des PA = Coût de la vente au PAMP
      //////////////////////////////////////////////////////////
      CoutVentePAMP := StrToFloat (copy (st, 107, 12));
      CoutVentePAMP := CoutVentePAMP/100 ;

      //////////////////////////////////////////////////////////
      // Récupération des PV
      //////////////////////////////////////////////////////////
      Montant := StrToFloat (copy (st, 95, 12));
      Montant := Montant/100 ;
      if (QuantiteVendueGB <> 0)then Montant:=Montant/QuantiteVendueGB	// Détermination du PA au moment de la vente
      else Montant:=0.0;

      Tob_Ligne.PutValue('GL_PUTTCDEV', Montant) ;

      Tob_Ligne.PutValue ('GL_PRIXPOURQTE', 1);
      Tob_Ligne.PutValue ('GL_TYPEDIM', 'NOR');
      Tob_Ligne.PutValue ('GL_QUALIFMVT', QualifMvt) ;

      TotQte :=TotQte+QuantiteVenduePGI;
      TotPrix:=TotPrix+QuantiteVenduePGI*Montant;
    end
  end
  else if Copy(st,1,5)='VTEC2' then
  begin
    /////////////////////////
    // Devise du document
    ////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      //
      // Recherche de la devise PGI dans la table de correcpondance
      //
      CodeDevGB  := Trim (copy (st, 39, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      ////////////////////////////////////////////
      // Devise du document
      /////////////////////////////////////////////
      DevisePiece:=CodeDevPGI ;
      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
           Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
           Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece ;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      // Devise dans entête ....
      PutValueDetail(Tob_Piece,'GP_TAUXDEV'      , DEV.Taux) ;
      AttribCotation(Tob_Piece) ;
      PutValueDetail(Tob_Piece,'GP_DATETAUXDEV'  , DEV.DateTaux) ;
      // Devise dans ligne ...
      PutValueDetail(Tob_Ligne, 'GL_DEVISE'      , Tob_Piece.GetValue ('GP_DEVISE'));
      PutValueDetail(Tob_Ligne, 'GL_SAISIECONTRE', Tob_Piece.GetValue ('GP_SAISIECONTRE'));
      PutValueDetail(Tob_Ligne, 'GL_TAUXDEV'     , DEV.Taux) ;
      PutValueDetail(Tob_Ligne, 'GL_COTATION'    , DEV.Taux) ;
      //PutValueDetail(Tob_Ligne, 'GL_DATETAUXDEV', DEV.DateTaux) ;

      //////////////////////////////////////////////////////////
      // récupération des PR = Coût de la vente au PRMP
      //////////////////////////////////////////////////////////
      // -> devise des prix ?
      DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
      begin
        if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
        else                    DevisePiece := 'EUR';
      end;
      //
      // PA de l'article
      //
      Montant := CoutVentePAMP;
      if (QuantiteVendueGB <> 0)then Montant:=Montant/QuantiteVendueGB	// Détermination du PA au moment de la vente
      else Montant:=0.0;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPA'      , Montant);
        Tob_Ligne.PutValue('GL_PMAP'     , Montant);
        Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPA'         , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMAPACTU'    , MontantDev);
      end;
      //
      // PR de l'article
      //
      Montant := StrToFloat (copy (st, 64, 12));
      Montant := Montant/100 ;
      if (QuantiteVendueGB <> 0)then Montant:=Montant/QuantiteVendueGB	// Détermination du PA au moment de la vente
      else Montant:=0.0;

      if  DevisePiece = V_PGI.DevisePivot then
      begin
        Tob_Ligne.PutValue('GL_DPR'      , Montant);
        Tob_Ligne.PutValue('GL_PMRP'     , Montant);
        Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
      end else
      begin
        DateConv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
        Tob_Ligne.PutValue ('GL_DPR'         , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRP'        , MontantDev);
        Tob_Ligne.PutValue ('GL_PMRPACTU'    , MontantDev);
      end;
  	end;
  end;
end;

///////////////////////////////////////////////////////////////////
// Récupération des règlements clients CASH 2000 ou GB 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupReglement (st : string);
var NaturePiece      : string   ;
    Etab             : string   ;
    DatePiece	     : string   ;
    NumTicket	     : string   ;
    RefVte           : string   ;		// Référence vente = Boutique + Date de vente + Numero de ticket
    CodeCli	         : string   ;
    RefCli	         : string   ;
    ModePaie 	     : string   ;
    CodeModGB        : string   ;
    CodeDevGB        : string   ;
    CodeDevPGI       : string   ;
    DevOrg           : string   ;
    DevDest          : string   ;
    DateVente        : TDateTime;
    Q                : TQUERY   ;
    DD   	         : TDateTime;
    Tob_Regl         : TOB      ;
    Reste	         : integer  ;
    Montant          : double   ;
    MontantDev       : double   ;
    MontantOld       : double   ;
    DateConv         : double   ;
begin

  if ArretRecup then exit;
  if ( (Copy(st,1,3)='RGL')	and (copy (st, 10, 1) = 'V') ) then
  begin
    NaturePiece := Tob_Param.GetValue ('NATVTE');   // Nature de pièce
    Etab        := copy (st,  7, 3);                // Etablissement
    DatePiece   := copy (st, 11, 8);                // Date de pièce
    NumTicket   := copy (st, 19, 6);                // Numéro de ticket

    //EcrireInfo ('   '+'Boutique '+ Copy(st,7,3) + ' Règlement de la Journée '+Copy (st,17,2)+'/'+copy (st,15,2)+'/'+copy (st,11,4)+ ' Ticket '+copy (st,19,6), False, True, False);
    inc (Compteur);
    Reste := Compteur Mod 10000;
    if (Reste = 0) then EcrireInfo (' Le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) + ' : ' + IntToStr (Compteur) + ' règlements intégrés', True, False, False);

    ///////////////////////////////////////////////////////////////////////////////////////
    // On change de ticket de vente, il faut retrouver l'entête
    ///////////////////////////////////////////////////////////////////////////////////////
    if (NumEtabVte<>Etab) or (DatePieceVte<>DatePiece) or (NumPieceVte<>NumTicket) then
    begin
      if (Tob_PiedEche <> Nil) then
      begin
    	ImportReglement ;
      end;

      NombreErreurs (0);
      Numligne     := 0;
      NumEtabVte   := Etab;
      DatePieceVte := DatePiece;
      NumPieceVte  := Numticket;

      RefVte := 'Btq ' + copy (st, 7, 3) + ' le ' + copy (st, 17, 2) + '/' + copy (st, 15, 2) + '/' + copy (st, 11,4) + ' Ticket ' + copy (st, 19, 6);

      DateVente := StrToDate (Copy (st,17,2) + '/' +copy (st,15,2) + '/ '+ copy (st,11,4));

      Q := OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_DATEPIECE="'+UsDateTime(DateVente)+'" AND GP_REFINTERNE="' + RefVte + '" ',False) ;
      if not Q.EOF then
      begin
      	NatureVente := Q.FindField ('GP_NATUREPIECEG').AsString;
        SoucheVente := Q.FindField ('GP_SOUCHE').AsString;
        NumeroVente := Q.FindField ('GP_NUMERO').AsInteger;
        IndiceVente := Q.FindField ('GP_INDICEG').AsInteger;

        Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
        AddLesSupEntete ( Tob_Piece);
	    //Tob_Piece.initValeurs;
       	Tob_Piece.SelectDB('',Q);
      end else
      begin
      	NombreErreurs (1);
        EcrireInfo ('   Erreur -> Impossible de retrouver le ticket PGI correspondant', True, True, True);
      end;
      Ferme (Q);
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      //////////////////////////////////////////////////////////////////////////
      // Recherche du mode de paiement PGI
      //////////////////////////////////////////////////////////////////////////
      CodeModGB  := Trim (copy (st, 32, 3));
      ModePaie   := TraduitGPAOenPGI(TypeRepriseGB2000, 'GPE_MODEPAIE', CodeModGB, Tob_ModePaie);
      Q := OpenSQL('Select MP_MODEPAIE from MODEPAIE Where MP_MODEPAIE="' + ModePaie + '" ',False) ;
      if Q.EOF then
      begin
        NombreErreurs (1);
        EcrireInfo ('   Erreur -> Le mode de paiement ' + ModePaie + ' est inexistant', True, True, True);
    	end;
      Ferme (Q);
    end;

    ////////////////////////////////////////////////////////////////
    // Recherche du tiers
    ////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      if (Tob_Tiers = nil) then
      begin
        if copy (st, 26, 6) <> '      ' then
        begin
       	  RefCli := Tob_Param.GetValue ('CODECLI');
    	  if RefCli='0' then CodeCli:=copy(st, 7, 3) + copy(st, 26 ,6)
    	  else if RefCli='1' then CodeCli:=copy(st, 26, 6);
        end else CodeCli:=Tob_Piece.GetValue ('GP_TIERS');

    	// Chargement en TOB du client
    	Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCli + '" AND T_NATUREAUXI = "CLI" ',False) ;
	    if not Q.EOF then
    	begin
      	  Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      	  //Tob_Tiers.initValeurs;
      	  Tob_Tiers.SelectDB('',Q);
          Ferme (Q);
        end else
    	begin
          Ferme (Q);
          EcrireInfo ('   Erreur -> Le tiers ' + Codecli + ' est introuvable, on reprend le tiers de la vente', True, True, True);
          CodeCli:=Tob_Piece.GetValue ('GP_TIERS');
          Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCli + '" AND T_NATUREAUXI = "CLI" ',False) ;
	      if not Q.EOF then
  	      begin
     	    Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      	    //Tob_Tiers.initValeurs;
      	    Tob_Tiers.SelectDB('',Q);
            Ferme (Q);
          end else
          begin
            NombreErreurs (1);
            EcrireInfo ('   Erreur -> Le tiers ' + Codecli + ' de la vente est introuvable', True, True, True);
          end;
        end;
      end;
    end;

    //////////////////////////////////////////////////////////////////////////
    // Chargement infos devise
    //////////////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      DD  := Tob_Piece.GetValue('GP_DATEPIECE');
      DEV.Code:=copy (st, 59, 3);
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, DD) ;
    end;
    //////////////////////////////////////////////////////////////////////////
    // Création de la TOB Reglement
    //////////////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      if (Tob_PiedEche = nil) then
      begin
      	Tob_PiedEche:= TOB.CREATE ('Les reglements', nil, -1);
    	//Tob_PiedEche.initValeurs;
      end;

      Tob_Regl := TOB.CREATE ('PIEDECHE', Tob_PiedEche, -1);
      //Tob_Regl.initValeurs;

      // Initialisation
      Tob_Regl.PutValue        ('GPE_NATUREPIECEG' , NaturePiece);
      Tob_Regl.PutFixedStValue ('GPE_DATEPIECE'    , st, 11, 8, tcDate8AMJ, false);
      Tob_Regl.PutValue        ('GPE_SOUCHE'       , SoucheVente);
      Tob_Regl.PutValue        ('GPE_NUMERO'       , NumeroVente);
      Tob_Regl.PutValue        ('GPE_INDICEG'      , IndiceVente);
      NumLigne := Numligne+1;
      Tob_Regl.PutValue        ('GPE_NUMECHE'      , NumLigne);

      Tob_Regl.PutValue        ('GPE_MODEPAIE'     , ModePaie);
      Tob_Regl.PutFixedStValue ('GPE_DATEECHE'     , st, 11, 8 , tcDate8AMJ, false);
      Tob_Regl.PutValue        ('GPE_TAUXDEV'      , Tob_Piece.GetValue ('GP_TAUXDEV'));
      Tob_Regl.PutValue        ('GPE_COTATION'     , Tob_Piece.GetValue ('GP_COTATION'));
      Tob_Regl.PutValue        ('GPE_DATETAUXDEV'  , Tob_Piece.GetValue ('GP_DATETAUXDEV'));

      //
      // Conversion devise des montants
      //
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Recherche de la devise PGI dans la table de correcpondance
        //
        CodeDevGB  := Trim (copy (st, 59, 3));
        CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GPE_DEVISE', CodeDevGB, Tob_Devise);
        //
        // devise de règlement
        //
        Tob_Regl.PutValue ('GPE_DEVISEESP', CodeDevPGI);
        //
        // Montant encaissé
        //
        Montant := StrToFloat (copy (st, 75, 12));
        Montant := Montant/100 ;
        Tob_Regl.PutValue ('GPE_MONTANTENCAIS', Montant);
        //
        // Conversion des autres montants en devise dossier
        //
        DevOrg     := Tob_Regl.GetValue ('GPE_DEVISEESP') ;
        DevDest    := V_PGI.DevisePivot;
        Dateconv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        Montant    := Tob_Regl.GetValue ('GPE_MONTANTENCAIS') ;
        ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, Tob_DevisePGI) ;
        Tob_Regl.PutValue ('GPE_MONTANTECHE', MontantDev) ;
        Tob_Regl.PutValue ('GPE_MONTANTDEV' , MontantDev) ;
        // Ré-initialisation des devises
        Tob_Regl.PutValue ('GPE_DEVISE'      , V_PGI.DevisePivot) ;
        Tob_Regl.PutValue ('GPE_SAISIECONTRE', '-') ;
      end;
    end;
  end
  else if ( (Copy(st,1,3)='RGL')	and (copy (st, 10, 1) = 'F') ) then
  begin
    NaturePiece := Tob_Param.GetValue ('NATNEGFAC');	// Nature de pièce
    Etab        := copy (st,  7, 3);                    // Etablissement
    DatePiece   := copy (st, 11, 8);                    // Date de pièce
    NumTicket   := copy (st, 19, 6);                    // Numéro de ticket

    EcrireInfo ('   '+'Boutique '+ Copy(st,7,3) + ' Règlement de la facture ' + copy (st,19,6), False, True, False);
    Compteur := Compteur+1;
    Reste := Compteur Mod 500;
    if (Reste = 0) then EcrireInfo (' Le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) + ' : ' + IntToStr (Compteur) + ' règlements intégrés', True, False, False);

    ///////////////////////////////////////////////////////////////////////////////////////
    // On change de ticket de vente, il faut retrouver l'entête
    ///////////////////////////////////////////////////////////////////////////////////////
    if (NumEtabVte<>Etab) or (DatePieceVte<>DatePiece) or (NumPieceVte<>NumTicket) then
    begin
      if (Tob_PiedEche <> Nil) then
      begin
    	ImportReglement ;
      end;

      NombreErreurs (0);
      Numligne     := 0;
      NumEtabVte   := Etab;
      DatePieceVte := DatePiece;
      NumPieceVte  := Numticket;

      RefVte := 'Btq ' + copy (st, 7, 3) + ' N°' + copy (st, 19, 6);
      PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefVte) ;

      DateVente := StrToDate (Copy (st,17,2) + '/' +copy (st,15,2) + '/ '+ copy (st,11,4));

      Q := OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_DATEPIECE="'+UsDateTime(DateVente)+'" AND GP_REFINTERNE="' + RefVte + '" ',False) ;
      if not Q.EOF then
      begin
      	NatureVente := Q.FindField ('GP_NATUREPIECEG').AsString;
        SoucheVente := Q.FindField ('GP_SOUCHE').AsString;
        NumeroVente := Q.FindField ('GP_NUMERO').AsInteger;
        IndiceVente := Q.FindField ('GP_INDICEG').AsInteger;

        Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
        AddLesSupEntete ( Tob_Piece);
	    //Tob_Piece.initValeurs;
       	Tob_Piece.SelectDB('',Q);
      end else
      begin
      	NombreErreurs (1);
        EcrireInfo ('   Erreur -> Impossible de retrouver la facture NEGOCE '+ copy (st, 19, 6) , True, True, True);
      end;
      Ferme (Q);
    end;

    if (NombreErreurs (-1) = 0) then
    begin
      //////////////////////////////////////////////////////////////////////////
      // Recherche du mode de paiement PGI
      //////////////////////////////////////////////////////////////////////////
      CodeModGB  := Trim (copy (st, 32, 3));
      ModePaie   := TraduitGPAOenPGI(TypeRepriseGB2000, 'GPE_MODEPAIE', CodeModGB, Tob_ModePaie);
      Q := OpenSQL('Select MP_MODEPAIE from MODEPAIE Where MP_MODEPAIE="' + ModePaie + '" ',False) ;
      if Q.EOF then
      begin
        NombreErreurs (1);
        EcrireInfo ('   Erreur -> Le mode de paiement ' + ModePaie + ' est inexixtant', True, True, True);
      end;
      Ferme (Q);
    end;

    ////////////////////////////////////////////////////////////////
    // Recherche du tiers
    ////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      if (Tob_Tiers = nil) then
      begin
        if copy (st, 26, 6) <> '      ' then
        begin
       	  RefCli := Tob_Param.GetValue ('CODECLI');
    	  if RefCli='0' then CodeCli:=copy(st, 7, 3) + copy(st, 26 ,6)
    	  else if RefCli='1' then CodeCli:=copy(st, 26, 6);
        end else CodeCli:=Tob_Piece.GetValue ('GP_TIERS');

    	// Chargement en TOB du client
    	Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCli + '" AND T_NATUREAUXI = "CLI" ',False) ;
	    if not Q.EOF then
    	begin
      	  Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      	  //Tob_Tiers.initValeurs;
      	  Tob_Tiers.SelectDB('',Q);
          Ferme (Q);
    	end else
    	begin
          Ferme (Q);

          EcrireInfo ('   Erreur -> Le tiers ' + Codecli + ' est introuvable, on reprend le tiers de la vente', True, True, True);

          CodeCli:=Tob_Piece.GetValue ('GP_TIERS');
          Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCli + '" AND T_NATUREAUXI = "CLI" ',False) ;
	      if not Q.EOF then
  	      begin
     	    Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      	    //Tob_Tiers.initValeurs;
      	    Tob_Tiers.SelectDB('',Q);
            Ferme (Q);
          end else
          begin
            NombreErreurs (1);
            EcrireInfo ('   Erreur -> Le tiers ' + Codecli + ' de la facture est introuvable', True, True, True);
          end;
        end;
      end;
    end;

    //////////////////////////////////////////////////////////////////////////
    // Chargement infos devise
    //////////////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      DD  := Tob_Piece.GetValue('GP_DATEPIECE');
      DEV.Code:=copy (st, 59, 3);
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, DD) ;
    end;
    //////////////////////////////////////////////////////////////////////////
    // Création de la TOB Reglement
    //////////////////////////////////////////////////////////////////////////
    if (NombreErreurs (-1) = 0) then
    begin
      if (Tob_PiedEche = nil) then
      begin
      	Tob_PiedEche:= TOB.CREATE ('Les reglements', nil, -1);
    	//Tob_PiedEche.initValeurs;
      end;

      Tob_Regl := TOB.CREATE ('PIEDECHE', Tob_PiedEche, -1);
      //Tob_Regl.initValeurs;

      // Initialisation
      Tob_Regl.PutValue        ('GPE_NATUREPIECEG' , NaturePiece);
      Tob_Regl.PutFixedStValue ('GPE_DATEPIECE'    , st, 11, 8, tcDate8AMJ, false);
      Tob_Regl.PutValue        ('GPE_SOUCHE'       , SoucheVente);
      Tob_Regl.PutValue        ('GPE_NUMERO'       , NumeroVente);
      Tob_Regl.PutValue        ('GPE_INDICEG'      , IndiceVente);
      NumLigne := Numligne+1;
      Tob_Regl.PutValue        ('GPE_NUMECHE'      , NumLigne);

      Tob_Regl.PutValue        ('GPE_MODEPAIE'     , ModePaie);
      if copy (st, 40, 8) <> '        ' then Tob_Regl.PutFixedStValue ('GPE_DATEECHE'     , st, 40, 8 , tcDate8AMJ, false)
      else Tob_Regl.PutFixedStValue ('GPE_DATEECHE'     , st, 11, 8 , tcDate8AMJ, false)  ;
      Tob_Regl.PutValue        ('GPE_TAUXDEV'      , Tob_Piece.GetValue ('GP_TAUXDEV'));
      Tob_Regl.PutValue        ('GPE_COTATION'     , Tob_Piece.GetValue ('GP_COTATION'));
      Tob_Regl.PutValue        ('GPE_DATETAUXDEV'  , Tob_Piece.GetValue ('GP_DATETAUXDEV'));

      //
      // Conversion devise des montants
      //
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Recherche de la devise PGI dans la table de correcpondance
        //
        CodeDevGB  := Trim (copy (st, 59, 3));
        CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GPE_DEVISE', CodeDevGB, Tob_Devise);
        //
        // devise de règlement
        //
        Tob_Regl.PutValue ('GPE_DEVISEESP', CodeDevPGI);
        //
        // Montant encaissé
        //
        Montant := StrToFloat (copy (st, 63, 12));
        Montant := Montant/100 ;
        Tob_Regl.PutValue ('GPE_MONTANTENCAIS', Montant);
        //
        // Conversion des autres montants en devise dossier
        //
        DevOrg     := Tob_Regl.GetValue ('GPE_DEVISEESP') ;
        DevDest    := V_PGI.DevisePivot;
        Dateconv   := NowH ;
        MontantDev := 0 ;  // initialisation bidon
        MontantOld := 0 ;  // initialisation bidon
        Montant    := Tob_Regl.GetValue ('GPE_MONTANTENCAIS') ;
        ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, Tob_DevisePGI) ;
        Tob_Regl.PutValue ('GPE_MONTANTECHE', MontantDev) ;
        Tob_Regl.PutValue ('GPE_MONTANTDEV' , MontantDev) ;
        // Ré-initialisation des devises
        Tob_Regl.PutValue ('GPE_DEVISE'      , V_PGI.DevisePivot) ;
        Tob_Regl.PutValue ('GPE_SAISIECONTRE', '-') ;
      end;
    end;
  end;
end;

///////////////////////////////////////////////////////////////////
// Récupération des commandes clients NEGOCE 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupNegCommande (st : string);
var CodeArt 	     : string;
    Rang    	     : string;
    CategorieDim     : string;
    NaturePiece      : string;
    NaturePieceFac   : string;
    NaturePieceAvo   : string;
    Etab             : string;
    SQL              : string;
    CodeDim1,CodeDim2: string;
    CodeDim3,CodeDim4: string;
    CodeDim5 	     : string;
    RefRec   	     : string;		// Référence commande = N° GB + Réf Cde GB
    CodeClient	     : string;
    ModeRegl         : string;
    CodeModGB        : string;
    GL_CodeArticle   : string;
    Refcli           : string;
    Qte, Prix	     : double;
    qteRec	     : double;
    Q                : TQUERY ;
    Per,Sem          : integer;
    i		     : integer;
    Taille           : integer;
    DD, DateConv     : TDateTime ;
    TobFille         : TOB ;
    TobPort          : TOB ;
    CodePort         : string ;
    TypePort         : string ;
    DevisePiece      : string ;
    CodeDevGB        : string ;
    CodeDevPGI       : string ;
    Base             : double ;
    Remise           : double ;
    QteTotCde        : double ;
    QteTotLiv        : double ;
    QteArtCde        : double ;
    QteArtLiv        : double ;
    QteCdeTai        : double ;
    QteRecTai        : double ;
    Montant          : double ;
    MontantDev       : double ;
    MontantOld       : double ;
begin

  NaturePiece := Tob_Param.GetValue ('NATNEGCDE');
  RefCli      := Tob_Param.GetValue ('CODECLI');
  Etab := copy (st, 8, 3);

  if ArretRecup then exit;

  if Copy (st, 7, 1) = 'C' then
  begin
    if Copy(st,1,5)='DENC1' then
    begin
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      // Importation de la réception précédente
      // 2 cas :	1 - Réception sans commande -> facile, importation directe
      //					2 - Réception dsur commande -> Recherche de la commande d'origine dans PGI (pour récup de l'indice)
      //						    												 Génération et Importation de la commande correspondante
      //																				 Importation de la réception
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if (Tob_Piece <> Nil) then
      begin
        ImportPiece;
      end;

      //////////////////////////////////////////////////////////////////////
      // Initialisations diverses et varièes
      //////////////////////////////////////////////////////////////////////
      NombreErreurs(0);      	// Initialisation du nombre d'erreurs
      TotQte              := 0;
      TotQteReste         := 0;
      TotPrix             := 0;
      Numligne            := 0;            // Initialisation du premier numéro de ligne
      NumLigPort          := 0;
      QualifMvt           := GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
      LigneCommACreer     := False;
      PieceAGenerer       := False;
      FacturesurCommande  := False;
      FactureSurLivraison := False;
      AvoirNegoce         := False;
      PrixUnique          := False;
      PrixVenteNeg        := 0.0;

      Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
      AddLesSupEntete ( Tob_Piece);
      //Tob_Piece.initValeurs;

      ////////////////////////////
      // Nature de pièce + Date
      ////////////////////////////
      Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
      Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 17, 8, tcDate8AMJ, false);
      Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 17, 8, tcDate8AMJ, false);
      /////////////////////////
      // Souche, Numéro, .....
      ////////////////////////
      CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
      CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
      CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');

      PutValueDetail(Tob_Piece, 'GP_SOUCHE' , CleDoc.Souche) ;
      PutValueDetail(Tob_Piece, 'GP_NUMERO' , CleDoc.NumeroPiece) ;
      PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
      //////////////////////////////////////////////////////
      // Référence de la facture
      //////////////////////////////////////////////////////
      RefRec := 'Btq ' + copy (st, 8, 3) + ' N°' + copy (st, 11, 6);
      PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefRec) ;
      ///////////////////////
      // Clients
      ///////////////////////
      CodeClient := copy (st, 25, 6);

      if RefCli='0' then CodeClient:=copy(st, 8, 3) + copy(st,25,6)
      else if RefCli='1' then CodeClient:=copy(st,25,6);

      Tob_Piece.PutValue ('GP_TIERS'       , CodeClient);
      Tob_Piece.PutValue ('GP_TIERSLIVRE'  , CodeClient);
      Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeClient);
      Tob_Piece.PutValue ('GP_TIERSPAYEUR' , CodeClient);
      // Chargement en TOB du client
      Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeClient + '" AND T_NATUREAUXI = "CLI" ',False) ;
      if not Q.EOF then
      begin
        Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
        //Tob_Tiers.initValeurs;
    	Tob_Tiers.SelectDB('',Q);
        ////////////////////////////////////////////////////////////////
        // Initialise les zones de l'entête en fonction du fournisseur
        ////////////////////////////////////////////////////////////////
        Tob_Piece.PutValue('GP_FACTUREHT'	     , Tob_Tiers.GetValue('T_FACTUREHT'));
	    Tob_Piece.PutValue('GP_ESCOMPTE'       , Tob_Tiers.GetValue('T_ESCOMPTE')) ;
	    Tob_Piece.PutValue('GP_MODEREGLE'      , Tob_Tiers.GetValue('T_MODEREGLE')) ;
	    Tob_Piece.PutValue('GP_REGIMETAXE'     , Tob_Tiers.GetValue('T_REGIMETVA')) ;
        Tob_Piece.PutValue('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers));
	    Tob_Piece.PutValue('GP_QUALIFESCOMPTE' , Tob_Tiers.GetValue('T_QUALIFESCOMPTE')) ;
	    Tob_Piece.PutValue('GP_TARIFTIERS'     , Tob_Tiers.GetValue('T_TARIFTIERS')) ;
      end else
      begin
        EcrireInfo ('   Erreur -> Client '+ Copy(st,25,5)+ ' inexistant dans PGI', True, True, True);
        NombreErreurs (1);
      end;
      Ferme(Q) ;

      if (NombreErreurs (-1) = 0) then
      begin
	Tob_Piece.PutFixedStValue ('GP_ETABLISSEMENT', st, 8, 3, tcTrim, false);
  	Tob_Piece.PutFixedStValue ('GP_DEPOT'        , st, 8, 3, tcTrim, false);

    	///////////////////////////////////////////////
    	// Champs par  défaut
    	//////////////////////////////////////////////
    	Tob_Piece.PutValue ('GP_VIVANTE'       , 'X');       // Pièce vivante    : Non
    	Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
    	Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
    	Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'VEN');     // Document         : Achat
    	Tob_Piece.PutValue ('GP_EDITEE'        , 'X');        // Pièce éditée     : Oui
    	Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
    	Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
        Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
  	Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
//        Tob_Piece.PutValue ('GP_SOCIETE'       , V_PGI.NomSociete);
        Tob_Piece.PutValue ('GP_DATECREATION'   , Date) ;
	Tob_Piece.PutValue ('GP_HEURECREATION'  , NowH) ;
        Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;

    	DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
	Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
	Tob_Piece.PutValue ('GP_PERIODE',Per) ;
    	Tob_Piece.PutValue ('GP_SEMAINE',Sem) ;
        //
        // Création d'une adresse de facturation si besoin
        //
        Tob_AdrFact := TOB.Create ('ADRESSES', nil, -1);
        Tob_AdrFact.PutValue       ('ADR_TYPEADRESSE', 'PIE');
        Tob_AdrFact.PutFixedStValue('ADR_JURIDIQUE'  , st, 37,  4, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_LIBELLE'    , st, 43, 32, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_ADRESSE1'   , st, 75, 32, tcTrim, false);
      end;
    end
    else if Copy(st,1,5)='DENC2' then
    begin
       Tob_AdrFact.PutFixedStValue('ADR_ADRESSE2'   , st, 17, 32, tcTrim, false);
       Tob_AdrFact.PutFixedStValue('ADR_ADRESSE1'   , st, 49, 32, tcTrim, false);
       Tob_AdrFact.PutFixedStValue('ADR_CODEPOSTAL' , st, 81,  7, tcTrim, false);
    end
    else if Copy(st,1,5)='DENC3' then
    begin
      Tob_AdrFact.PutFixedStValue('ADR_VILLE'     , st, 17, 32, tcTrim, false);
      Tob_AdrFact.PutFixedStValue('ADR_PAYS'      , st, 49,  3, tcTrim, false);
      Tob_AdrFact.PutFixedStValue('ADR_TELEPHONE' , st, 52, 20, tcTrim, false);
      //
      // Reprise du code client à facturer
      //
      if copy (st, 80, 6) <> '      ' then
      begin
        CodeClient := copy (st, 80, 6);

        if RefCli='0' then CodeClient:=copy(st, 8, 3) + copy(st,80,6)
  	else if RefCli='1' then CodeClient:=copy(st,80,6);

        Tob_Piece.PutValue('GP_TIERSFACTURE', CodeClient);
      end;
    end
    else if Copy(st,1,5)='DENC4' then
    begin
      //
      // Prénom du client
      //
      Tob_AdrFact.PutFixedStValue('ADR_LIBELLE2', st, 56, 32, tcTrim, false);
      //
      // Devise du document
      //
      ////////////////////////////////////////////
      // Devise du document
      /////////////////////////////////////////////
      CodeDevGB  := Trim (copy (st, 23, 3));
      CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

      DevisePiece := CodedevPGI;
      if VH^.TenueEuro then
      begin
        if DevisePiece = V_PGI.DeviseFongible then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end else
      begin
        if DevisePiece = 'EUR' then
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
          Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
        end else
        begin
          Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
        end;
      end;

      DEV.Code:=DevisePiece;
      GetInfosDevise(DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      AttribCotation(Tob_Piece) ;
      PutValueDetail(Tob_Piece, 'GP_DATETAUXDEV',DEV.DateTaux) ;
      //
      // Facture ou avoir ?????
      //
      AvoirNegoce := False ;
      ///
      // Recherche du mode de paiement PGI
      ///
      CodeModGB  := Trim (copy (st, 38, 3));
      ModeRegl   := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_MODEREGLE', CodeModGB, Tob_Reglement);
      Q := OpenSQL('Select MR_MODEREGLE from MODEREGL Where MR_MODEREGLE="' + ModeRegl + '" ',False) ;
      if Q.EOF then
      begin
        NombreErreurs (1);
        EcrireInfo ('   Erreur -> Le mode de règlement ' + ModeRegl + ' est inexistant', True, True, True);
      end else
      begin
        Tob_Piece.PutValue ('GP_MODEREGLE', ModeRegl);
      end;
      Ferme (Q);
      //
      // Facture TTC ou HT ?????????
      //
      if copy (st, 51, 1) = 'T' then Tob_Piece.PutValue ('GP_FACTUREHT', '-')
      else Tob_Piece.PutValue ('GP_FACTUREHT', 'X') ;

      if copy (st, 20, 3) <> '   ' then  Tob_Piece.PutFixedStValue('GP_REPRESENTANT', st, 20, 3, tcTrim, false);

    end
    else if Copy(st,1,5)='DENC5' then
    begin
      /// ?????????
      QteTotCde := StrToFloat (copy (st, 33, 8));
      if QteTotCde = 0 then
      begin
        EcrireInfo ('   Attention -> Commande ' + copy (st, 11, 6) + ' : Total des quantités à 0 -> Commande non reprise', True, True, True);
        NombreErreurs (1);
      end;
      QteTotLiv := StrToFloat (copy (st, 41, 8));
      if QteTotCde<=QteTotLiv then
      begin
         EcrireInfo ('   Attention -> Commande ' + copy (st, 11, 6) + ' soldée : Commande non reprise', True, True, True);
        NombreErreurs (1);
      end;
    end
    else if Copy(st,1,5)='DLAC1' then
    begin
  	  if (NombreErreurs (-1) = 0) then
      begin
        LigneAReprendre := True ;
        //
     	  // Création de la ligne commentaire précédente ???
        //
        if (LigneCommACreer = True) then
        begin
          LigneCommACreer:=False;
          CreeLigneCommentaire;
        end;
     	  //
     	  // Gestion d'une ligne commentaire avec l'article générique ????
     	  //
        NumLigne        := NumLigne+1;
     	  NumLigneComment := Numligne;
     	  LigneCommACreer := True;

          TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
     	  //
          // formatage de la référence article générique
     	  //
     	  CodeArticle := Copy (st, 23, 12) ;
     	  CodeArt     := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
     	  //
     	  // Récupération du libellé de l'article
     	  //
     	  SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
     	  Q:=OpenSQL(SQL,True) ;
     	  if Not Q.EOF then	Designation := Q.FindField('GA_LIBELLE').AsString
     	  else begin
       	  EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,23,12), True, True, True);
      	  NombreErreurs (1);
        end;
    	  Ferme(Q) ;
        //
        // Prix unique ????
        //
        if copy (st, 85, 1) = 'O' then PrixUnique := True
        else PrixUnique := False;
        PrixVenteNeg := 0;
      end;
    end
    else if Copy(st,1,5)='DLAC2' then
    begin
        QteArtCde := StrToFloat (copy (st, 30, 8));
        QteArtLiv := StrToFloat (copy (st, 37, 8));
        if QteArtCde <= QteArtLiv then LigneAReprendre := False;
    end
    else if Copy(st,1,5)='DLAC3' then
    begin
         // Rien à faire
    end
    else if Copy(st,1,5)='DTIC1' then
    begin
      if (NombreErreurs (-1) = 0) and LigneAReprendre then
      begin
        QteCdeTai := StrtoFloat (copy (st, 30, 5));
        QteRecTai := StrtoFloat (copy (st, 35, 5));
        if QteCdeTai > QteRecTai then
        begin
          ////////////////////////////////////////////////////////////////
    	    // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    	    ////////////////////////////////////////////////////////////////
    	    Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
    	    //Tob_Ligne.initValeurs;
    	    ////////////////////////////////////////////////////////////////
    	    // Initialise les champs de la ligne à partir de l'entête
    	    ////////////////////////////////////////////////////////////////
    	    PieceVersLigne (Tob_Piece, Tob_Ligne);
    	    Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
    	    Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
    	    //////////////////////////////////////////////
    	    // Nouveau numéro de ligne
    	    //////////////////////////////////////////////
    	    NumLigne := NumLigne+1;
    	    Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
    	    Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
    	    ////////////////////////////////////////////////////////////////
    	    // Ajoute des champs supplémentaires pour le calcul des cumuls
    	    ////////////////////////////////////////////////////////////////
    	    AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
    	    //////////////////////////////////////////////////////////
    	    // Détermination de l'article taille
    	    /////////////////////////////////////////////////////////
    	    CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

          Taille := StrToInt (copy (st, 23 ,2));
          rang   := CalculTaillePGI (Taille);

    	    CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
    	    CodeDim4 := '   '; CodeDim5 := '   ';
    	    if (CategorieDim [3] = '1') then Codedim1 :=rang
    	    else if (CategorieDim [3] = '2') then Codedim2 :=rang
    	    else if (CategorieDim [3] = '3') then Codedim3 :=rang
    	    else if (CategorieDim [3] = '4') then Codedim4 :=rang
    	    else if (CategorieDim [3] = '5') then Codedim5 :=rang;
    	    Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

    	    SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
    	    Q:=OpenSQL(SQL,True) ;
    	    if Not Q.EOF then
    	    begin
    		    if (Tob_Article=nil) then
      	    begin
              Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     		  //Tob_Article.initValeurs;
            end;
				    /////////////////////////////////////////////////////////////////////////////
      	    // Initialisation des champs de la ligne à partir de l'article dimensionné
      	    /////////////////////////////////////////////////////////////////////////////
      	    Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	    Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	    Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	    Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	    Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	    Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	    Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	    Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	    Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	    Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	    Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	    Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	    Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	    Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
            Tob_Ligne.PutValue ('GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
            Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
            Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	   		Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  		   	Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
        	for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
            for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
            Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
      	    /////////////////////////////////////////////////////////////////////////////
      	    // Ajout de l'article dans la TOB Article
      	    /////////////////////////////////////////////////////////////////////////////
      	    Tob_Article.SelectDB('',Q);
    	    end
    	    else begin
      	    EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	    NombreErreurs (1);
          end;
    	    Ferme(Q) ;

          if (NombreErreurs (-1) = 0) then
          begin
	        Tob_Ligne.PutValue ('GL_ARTICLE'     , Codeart);
  	        Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

            GL_CodeArticle := Trim (CodeArticle);
      	    Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	        Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

    	    Tob_Ligne.PutValue ('GL_TARIF'       , 0);
    	    Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	        Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
            Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
            Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

            QteCdeTai := StrtoFloat (copy (st, 30, 5));
            QteRecTai := StrtoFloat (copy (st, 35, 5));

            QteRec := QteCdeTai - QteRecTai;       // On récupère que l'encours

  	        Tob_Ligne.PutValue('GL_QTESTOCK', QteRec);
    	    Tob_Ligne.PutValue('GL_QTEFACT' , QteRec);
            Tob_Ligne.PutValue('GL_QTERESTE', QteRec);

            DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
            if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
            begin
              if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
              else                    DevisePiece := 'EUR';
            end;
            //
            // PA de l'article
            //
            Montant := StrToFloat (copy (st, 40, 8));
            Montant := Montant/100 ;

            if  DevisePiece <> V_PGI.DevisePivot then
            begin
              DateConv   := NowH ;
              MontantDev := 0 ;  // initialisation bidon
              MontantOld := 0 ;  // initialisation bidon
              ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
              Tob_Ligne.PutValue ('GL_DPA', MontantDev);
            end;
            //
            // PAMP de l'article
            //
            Montant := StrToFloat (copy (st, 48, 8));
            Montant := Montant/100 ;

            if  DevisePiece = V_PGI.DevisePivot then
            begin
              Tob_Ligne.PutValue('GL_PMAP'     , Montant);
              Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
            end else
            begin
              DateConv   := NowH ;
              MontantDev := 0 ;  // initialisation bidon
              MontantOld := 0 ;  // initialisation bidon
              ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
              Tob_Ligne.PutValue ('GL_PMAP'        , MontantDev);
              Tob_Ligne.PutValue ('GL_PMAPACTU'    , MontantDev);
            end;
            //
            // PR de l'article
            //
            Montant := StrToFloat (copy (st, 72, 8));
            Montant := Montant/100 ;

            if  DevisePiece = V_PGI.DevisePivot then Tob_Ligne.PutValue('GL_DPR', Montant)
            else begin
              DateConv   := NowH ;
              MontantDev := 0 ;  // initialisation bidon
              MontantOld := 0 ;  // initialisation bidon
              ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
              Tob_Ligne.PutValue ('GL_DPR', MontantDev);
            end;

            //
            // PRMP de l'article
            //
            Montant := StrToFloat (copy (st, 80, 8));
            Montant := Montant/100 ;

            if  DevisePiece = V_PGI.DevisePivot then
            begin
              Tob_Ligne.PutValue('GL_PMRP'     , Montant);
              Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
            end else
            begin
              DateConv   := NowH ;
              MontantDev := 0 ;  // initialisation bidon
              MontantOld := 0 ;  // initialisation bidon
              ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
              Tob_Ligne.PutValue ('GL_PMRP'    , MontantDev);
              Tob_Ligne.PutValue ('GL_PMRPACTU', MontantDev);
            end;

            //
            // Gestion du prix unique
            //
            if PrixUnique = True then
            begin
              if PrixVenteNeg = 0 then
              begin
                PrixVenteNeg := StrToFloat (copy (st, 64,8));
                PrixVenteNeg := PrixVenteNeg/100;
              end;

              if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
              begin
                Tob_Ligne.PutValue('GL_PUHTDEV'   , PrixVenteNeg);
    	        Tob_Ligne.PutValue('GL_PUHTNETDEV', PrixVenteNeg);
              end else
              begin
                Tob_Ligne.PutValue('GL_PUTTCDEV'   , PrixVenteNeg);
    	        Tob_Ligne.PutValue('GL_PUTTCNETDEV', PrixVenteNeg);
              end;
            end else
            begin
              if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
              begin
                Tob_Ligne.PutFixedStValue('GL_PUHTDEV'   , st, 64, 8, tcDouble100, false);
    	        Tob_Ligne.PutFixedStValue('GL_PUHTNETDEV', st, 64, 8, tcDouble100, false);
              end else
              begin
                Tob_Ligne.PutFixedStValue('GL_PUTTCDEV'   , st, 64, 8, tcDouble100, false);
    	        Tob_Ligne.PutFixedStValue('GL_PUTTCNETDEV', st, 64, 8, tcDouble100, false);
              end;
            end;
  	        Tob_Ligne.PutValue         ('GL_PRIXPOURQTE', 1);
            Tob_Ligne.PutValue         ('GL_PCB'        , 1);
  	        Tob_Ligne.PutValue         ('GL_TYPEDIM'    , 'DIM');
    	    Tob_Ligne.PutValue         ('GL_QUALIFMVT'  , QualifMvt) ;

            if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then Prix:=Tob_Ligne.GetValue('GL_PUHTDEV')
      	    else Prix:=Tob_Ligne.GetValue('GL_PUTTCDEV') ;

            Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
    	    TotQte      :=TotQte+Qte;
            TotQteReste :=TotQteReste+Qte;
	        TotPrix     :=TotPrix+Qte*Prix;
          end;
        end;
      end;
    end
    else if Copy(st,1,5)='DLRC1' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;

        //  Gestion de la remise
        Remise := StrToFloat (copy (st, 72, 12));
        Remise := Remise/100 ;
        if Remise <> 0 then
        begin
          Tob_Piece.PutValue ('GP_REMISEPIED', Remise);
        end;
      end;
    end
    else if Copy(st,1,5)='DLXC1' then             // Ligne de Port : ne marche pas en V530
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;
        //
        // Recherche du port dans la table des ports
        //
        CodePort := copy (st, 43, 3);
        SQL:='Select * From PORT WHERE GPO_CODEPORT="'+CodePort+'"';
    	  Q:=OpenSQL(SQL,True) ;

    	  if Not Q.EOF then
    	  begin
          //
          // Création de la TOB des ports
          //
    	  if (Tob_PiedPort=nil) then
      	  begin
            Tob_PiedPort := TOB.CREATE ('Les ports', NIL, -1);
     	    //Tob_PiedPort.initValeurs;
          end;
          //
          //  Chargement en TOB du descriptif du Port
          //
          TobPort := TOB.Create ('PORT', nil, -1);
          TobPort.SelectDB ('', Q);
          //
          // Alimentation de la TOB PIEDPORT
          //
          TobFille := Tob.create ('PIEDPORT', Tob_PiedPort, -1);
          //TobFille.InitValeurs ;

          TobFille.PutValue ('GPT_CODEPORT'     , CodePort);
          TobFille.PutValue ('GPT_LIBELLE', TobPort.GetValue('GPO_LIBELLE'));

          TypePort := TobPort.GetValue('GPO_TYPEPORT') ;
          TobFille.PutValue ('GPT_TYPEPORT'     , TypePort);
          TobFille.PutValue ('GPT_FAMILLETAXE1' , TobPort.GetValue('GPO_FAMILLETAXE1'));
          TobFille.PutValue ('GPT_FAMILLETAXE2' , TobPort.GetValue('GPO_FAMILLETAXE2'));
          TobFille.PutValue ('GPT_FAMILLETAXE3' , TobPort.GetValue('GPO_FAMILLETAXE3'));
          TobFille.PutValue ('GPT_FAMILLETAXE4' , TobPort.GetValue('GPO_FAMILLETAXE4'));
          TobFille.PutValue ('GPT_FAMILLETAXE5' , TobPort.GetValue('GPO_FAMILLETAXE5'));
          TobFille.PutValue ('GPT_COMPTAARTICLE', TobPort.GetValue('GPO_COMPTAARTICLE'));

          Base    := TobPort.GetValue('GPO_MINIMUM') ;
          TobFille.PutValue ('GPT_MINIMUM', Base);
          TobFille.PutValue ('GPT_FRANCO' , TobPort.GetValue('GPO_FRANCO'));
          TobFille.PutValue ('GPT_MONTANTMINI', 0);
          if TypePort='MT' then
          begin
            TobFille.PutValue      ('GPT_POURCENT', 0) ;
            Base:=TobPort.GetValue ('GPO_PVHT') ;
            TobFille.PutValue      ('GPT_MONTANTMINI', Base) ;
          end else
          begin
            TobFille.PutValue ('GPT_POURCENT', TobPort.GetValue('GPO_COEFF'));
          end;
          if TypePort='MI' then
          begin
            Base:=TobPort.GetValue('GPO_PVHT') ;
            TobFille.PutValue ('GPT_MONTANTMINI', Base) ;
          end ;
          TobFille.PutValue ('GPT_NATUREPIECEG' , Tob_Piece.GetValue ('GP_NATUREPIECEG'));
          TobFille.PutValue ('GPT_SOUCHE'       , Tob_Piece.GetValue ('GP_SOUCHE'));
          TobFille.PutValue ('GPT_NUMERO'       , Tob_Piece.GetValue ('GP_NUMERO'));
          TobFille.PutValue ('GPT_INDICEG'      , Tob_Piece.GetValue ('GP_INDICEG'));
          TobFille.PutValue ('GPT_DEVISE'       , Tob_Piece.GetValue ('GP_DEVISE'));
          Inc (NumLigPort);
          TobFille.PutValue ('GPT_NUMPORT'      , NumLigPort);

          //
          // Récupération des montants dans le fichier
          //
          //TobFille.PutFixedStValue('GPT_BASEHT'    , st, 73, 12, tcDouble100, false);
          //TobFille.PutFixedStValue('GPT_BASEHTDEV' , st, 73, 12, tcDouble100, false);
    	    //TobFille.PutFixedStValue('GPT_BASETTC'   , st, 97, 12, tcDouble100, false);
          //TobFille.PutFixedStValue('GPT_BASETTCDEV', st, 97, 12, tcDouble100, false);
          DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
          if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
          begin
            if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
            else                    DevisePiece := 'EUR';
          end;
          //
          // Prix du port
          //
          Montant := StrToFloat (copy (st, 73, 12));
          Montant := Montant/100 ;

          if  DevisePiece = V_PGI.DevisePivot then
          begin
            Tob_Ligne.PutValue('GPT_TOTALHT'      , Montant);
            Tob_Ligne.PutValue('GPT_TOTALHTDEV'   , Montant);
          end else
          begin
            DateConv   := NowH ;
            MontantDev := 0 ;  // initialisation bidon
            MontantOld := 0 ;  // initialisation bidon
            ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
            Tob_Ligne.PutValue ('GPT_TOTALHT'      , Montant);
            Tob_Ligne.PutValue ('GPT_TOTALHTDEV'   , MontantDev);
          end;
        end else
        begin
          EcrireInfo ('   Erreur -> Problème pour récupérer le PORT '+ CodePort, True, True, True);
       	  NombreErreurs (1);
        end;
        Ferme (Q);
      end;
    end
    else if Copy(st,1,5)='DLPC1' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;
        //
     	  // Création de la ligne commentaire précédente ???
        //
        if (LigneCommACreer = True) then
        begin
          LigneCommACreer:=False;
          CreeLigneCommentaire;
        end;

        ////////////////////////////////////////////////////////////////
    	  // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    	  ////////////////////////////////////////////////////////////////
    	  Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
    	  //Tob_Ligne.initValeurs;
    	  ////////////////////////////////////////////////////////////////
    	  // Initialise les champs de la ligne à partir de l'entête
    	  ////////////////////////////////////////////////////////////////
    	  PieceVersLigne (Tob_Piece, Tob_Ligne);
    	  Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
    	  Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
    	  //////////////////////////////////////////////
    	  // Nouveau numéro de ligne
    	  //////////////////////////////////////////////
    	  NumLigne := NumLigne+1;
    	  Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
    	  Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
    	  ////////////////////////////////////////////////////////////////
    	  // Ajoute des champs supplémentaires pour le calcul des cumuls
    	  ////////////////////////////////////////////////////////////////
    	  AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
    	  //////////////////////////////////////////////////////////
    	  // Détermination de l'article taille
    	  /////////////////////////////////////////////////////////
        CodeArticle := copy (st, 43, 3) + '         ' ;

        Codeart := CodeArticle + '      ' + '   ' + '   ' + '   ' + '   ' + '   ' + 'X'; // Code article PGI

    	  SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
    	  Q:=OpenSQL(SQL,True) ;
    	  if Not Q.EOF then
    	  begin
            if (Tob_Article=nil) then
      	    begin
              Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     	      //Tob_Article.initValeurs;
            end;
	   /////////////////////////////////////////////////////////////////////////////
      	  // Initialisation des champs de la ligne à partir de l'article dimensionné
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	  Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	  Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	  Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	  Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
          Tob_Ligne.PutValue ('GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
          Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
          Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	      Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	      Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
          for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
          for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
          Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
          Tob_Ligne.PutValue( 'GL_LIBELLE'     , Q.FindField('GA_LIBELLE').AsString) ;
      	  /////////////////////////////////////////////////////////////////////////////
      	  // Ajout de l'article dans la TOB Article
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Article.SelectDB('',Q);
    	  end
    	  else begin
      	  EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	  NombreErreurs (1);
        end;
    	  Ferme(Q) ;
      end;

      if (NombreErreurs (-1) = 0) then
      begin
	Tob_Ligne.PutValue ('GL_ARTICLE'     , Codeart);
  	//Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

        GL_CodeArticle := Trim (CodeArticle);
      	Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

    	Tob_Ligne.PutValue ('GL_TARIF'       , 0);
    	Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	//Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'PRE');
        Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
        Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
        Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

        QteRec := 1.0;                         // Quantité Port : forcément 1
        if AvoirNegoce then QteRec:=-QteRec;
        Tob_Ligne.PutValue('GL_QTESTOCK', QteRec);
    	Tob_Ligne.PutValue('GL_QTEFACT' , QteRec);
        Tob_Ligne.PutValue('GL_QTERESTE', QteRec);

	//////////////////////////////////////////////////////////
  	// Récupération des PA et PR
    	//////////////////////////////////////////////////////////
        DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
        if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
        begin
          if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
          else                    DevisePiece := 'EUR';
        end;
        //
        // Prix du port
        //
        Montant := StrToFloat (copy (st, 73, 12));
        Montant := Montant/100 ;

        if  DevisePiece = V_PGI.DevisePivot then
        begin
          Tob_Ligne.PutValue('GL_DPA'        , Montant) ;
  	      Tob_Ligne.PutValue('GL_PMAP'       , Montant) ;
    	  Tob_Ligne.PutValue('GL_PMAPACTU'   , Montant) ;
   	      Tob_Ligne.PutValue('GL_DPR'        , Montant) ;
  	      Tob_Ligne.PutValue('GL_PMRP'       , Montant) ;
    	  Tob_Ligne.PutValue('GL_PMRPACTU'   , Montant) ;
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Ligne.PutValue('GL_DPA'        , MontantDev) ;
  	      Tob_Ligne.PutValue('GL_PMAP'       , MontantDev) ;
    	  Tob_Ligne.PutValue('GL_PMAPACTU'   , MontantDev) ;
  	      Tob_Ligne.PutValue('GL_DPR'        , MontantDev) ;
  	      Tob_Ligne.PutValue('GL_PMRP'       , MontantDev) ;
    	  Tob_Ligne.PutValue('GL_PMRPACTU'   , MontantDev) ;;
        end;

        if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
        begin
          Tob_Ligne.PutFixedStValue('GL_PUHTDEV'   , st, 73, 12, tcDouble100, false);
    	  Tob_Ligne.PutFixedStValue('GL_PUHTNETDEV', st, 73, 12, tcDouble100, false);
        end else
        begin
          Tob_Ligne.PutFixedStValue('GL_PUTTCDEV'   , st, 97, 12, tcDouble100, false);
    	  Tob_Ligne.PutFixedStValue('GL_PUTTCNETDEV', st, 97, 12, tcDouble100, false);
        end;
        Tob_Ligne.PutValue         ('GL_PRIXPOURQTE'    , 1);
        Tob_Ligne.PutValue         ('GL_PCB'            , 1);
        Tob_Ligne.PutValue         ('GL_TVAENCAISSEMENT', 'X');
  	    Tob_Ligne.PutValue         ('GL_TYPEDIM'        , 'NOR');
    	Tob_Ligne.PutValue         ('GL_QUALIFMVT'      , QualifMvt) ;
        Tob_Ligne.PutValue         ('GL_DEPOT'          , '') ;

        //if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then Prix:=Tob_Ligne.GetValue('GL_PUHTDEV')
    	//else Prix:=Tob_Ligne.GetValue('GL_PUTTCDEV') ;

        ////////////////////////////////////////////////////////////////////////////////////////////
    	// Ajoute des champs supplémentaires pour la génération des commandes ou livraison initiales
    	////////////////////////////////////////////////////////////////////////////////////////////
        //if FactureSurCommande then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'C')
        //else if FactureSurLivraison then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'L') ;

        //Tob_Ligne.AddChampSupValeur('$$_CODEBTQ' , BtqNeg2000);
        //Tob_Ligne.AddChampSupValeur('$$_NUMDOC'  , NumNeg2000);
      end;
    end;
  end;
end;



///////////////////////////////////////////////////////////////////
// Récupération des factures clients NEGOCE 2000
///////////////////////////////////////////////////////////////////
Procedure TFRecupGB2000.RecupNegFacture (st : string);
var CodeArt 	     : string;
    Rang    	     : string;
    CategorieDim     : string;
    NaturePiece      : string;
    NaturePieceFac   : string;
    NaturePieceAvo   : string;
    Etab             : string;
    SQL              : string;
    CodeDim1,CodeDim2: string;
    CodeDim3,CodeDim4: string;
    CodeDim5 	     : string;
    RefRec   	     : string;		// Référence commande = N° GB + Réf Cde GB
    CodeClient	     : string;
    ModeRegl         : string;
    CodeModGB        : string;
    GL_CodeArticle   : string;
    CodeDevPGI       : string;
    CodeDevGB        : string;
    DevisePiece      : string;
    RefCli           : string;
    Qte, Prix	     : double;
    qteRec	     : double;
    Q                : TQUERY ;
    Per,Sem          : integer;
    i		     : integer;
    Taille           : integer;
    DD, DateConv     : TDateTime ;
    TobFille         : TOB ;
    TobPort          : TOB ;
    CodePort         : string ;
    TypePort         : string ;
    Base             : double ;
    Remise           : double ;
    QteTotFac        : double ;
    Montant          : double ;
    MontantDev       : double ;
    MontantOld       : double ;
begin

  NaturePieceFac := Tob_Param.GetValue ('NATNEGFAC');
  NaturePieceAvo := Tob_Param.GetValue ('NATNEGAVO');
  RefCli         := Tob_Param.GetValue ('CODECLI');

  Etab := copy (st, 8, 3);

  if ArretRecup then exit;

  if Copy (st, 7, 1) = 'F' then
  begin
    if Copy(st,1,5)='DENC1' then
    begin
  	  ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      // Importation de la réception précédente
      // 2 cas :	1 - Réception sans commande -> facile, importation directe
  	  //					2 - Réception dsur commande -> Recherche de la commande d'origine dans PGI (pour récup de l'indice)
		  //						    												 Génération et Importation de la commande correspondante
      //																				 Importation de la réception
      ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
      if (Tob_Piece <> Nil) then
      begin
        if AvoirNegoce then ImportPiece
        else begin
          if PieceAGenerer then ImportFacCdeLiv
          else ImportPiece;
        end;
      end;

      //
      // Facture ou avoir ???
      //
      if copy (st, 107, 1) = 'A' then
      begin
        EcrireInfo ('   ' + 'Boutique ' + Copy(st,7,3) + ' Avoir ' + Copy (st,11, 6) , True, True, False);
        NaturePiece := NaturePieceAvo;
        AvoirNegoce := True ;
      end else
      begin
        EcrireInfo ('   ' + 'Boutique ' + Copy(st,8,3) + ' Facture ' + Copy (st,11, 6) , True, True, False);
        NaturePiece := NaturePieceFac;
        AvoirNegoce := False ;
      end;

      //////////////////////////////////////////////////////////////////////
      // Initialisations diverses et varièes
      //////////////////////////////////////////////////////////////////////
      NombreErreurs(0);      	// Initialisation du nombre d'erreurs
      TotQte              := 0;
      TotQteReste         := 0;
      TotPrix             := 0;
      Numligne            := 0;            // Initialisation du premier numéro de ligne
      NumLigPort          := 0;
      QualifMvt           := GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
      LigneCommACreer     := False;
      PieceAGenerer       := False;
      FacturesurCommande  := False;
      FactureSurLivraison := False;
      AvoirNegoce         := False;
      PrixUnique          := False;
      PrixVenteNeg        := 0.0;

      Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
      AddLesSupEntete ( Tob_Piece);
      //Tob_Piece.initValeurs;

      ////////////////////////////
      // Nature de pièce + Date
      ////////////////////////////
      Tob_Piece.PutValue       ('GP_NATUREPIECEG' , NaturePiece);
      Tob_Piece.PutFixedStValue('GP_DATEPIECE'    , st, 17, 8, tcDate8AMJ, false);
      Tob_Piece.PutFixedStValue('GP_DATELIVRAISON', st, 17, 8, tcDate8AMJ, false);
      /////////////////////////
      // Souche, Numéro, .....
      ////////////////////////
      CleDoc.Souche     := GetSoucheG(NaturePiece, Etab,'') ;
      CleDoc.NumeroPiece:= GetNumSoucheG(CleDoc.Souche) ;
      CleDoc.DatePiece  := Tob_Piece.GetValue('GP_DATEPIECE');

      PutValueDetail(Tob_Piece, 'GP_SOUCHE' , CleDoc.Souche) ;
      PutValueDetail(Tob_Piece, 'GP_NUMERO' , CleDoc.NumeroPiece) ;
      PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
      //////////////////////////////////////////////////////
		  // Référence de la facture
      //////////////////////////////////////////////////////
      RefRec := 'Btq ' + copy (st, 8, 3) + ' N°' + copy (st, 11, 6);
      PutValueDetail(Tob_Piece,'GP_REFINTERNE', RefRec) ;
      ///////////////////////
      // Clients
      ///////////////////////
      CodeClient := copy (st, 25, 6);

      if RefCli='0' then CodeClient:=copy(st, 8, 3) + copy(st,25,6)
  	  else if RefCli='1' then CodeClient:=copy(st,25,6);

      Tob_Piece.PutValue ('GP_TIERS'       , CodeClient);
      Tob_Piece.PutValue ('GP_TIERSLIVRE'  , CodeClient);
      Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeClient);
      Tob_Piece.PutValue ('GP_TIERSPAYEUR' , CodeClient);
      // Chargement en TOB du client
      Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeClient + '" AND T_NATUREAUXI = "CLI" ',False) ;
		  if not Q.EOF then
      begin
        Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
        //Tob_Tiers.initValeurs;
    	Tob_Tiers.SelectDB('',Q);
        ////////////////////////////////////////////////////////////////
        // Initialise les zones de l'entête en fonction du fournisseur
        ////////////////////////////////////////////////////////////////
			  Tob_Piece.PutValue('GP_FACTUREHT'	     , Tob_Tiers.GetValue('T_FACTUREHT'));
			  Tob_Piece.PutValue('GP_ESCOMPTE'       , Tob_Tiers.GetValue('T_ESCOMPTE')) ;
			  Tob_Piece.PutValue('GP_MODEREGLE'      , Tob_Tiers.GetValue('T_MODEREGLE')) ;
			  Tob_Piece.PutValue('GP_REGIMETAXE'     , Tob_Tiers.GetValue('T_REGIMETVA')) ;
        Tob_Piece.PutValue('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers));
			  Tob_Piece.PutValue('GP_QUALIFESCOMPTE' , Tob_Tiers.GetValue('T_QUALIFESCOMPTE')) ;
			  Tob_Piece.PutValue('GP_TARIFTIERS'     , Tob_Tiers.GetValue('T_TARIFTIERS')) ;
      end else
      begin
        EcrireInfo ('   Erreur -> Client '+ CodeClient + ' inexistant dans PGI', True, True, True);
        NombreErreurs (1);
      end;
      Ferme(Q) ;

      if (NombreErreurs (-1) = 0) then
      begin
	      Tob_Piece.PutFixedStValue ('GP_ETABLISSEMENT', st, 8, 3, tcTrim, false);
  	    Tob_Piece.PutFixedStValue ('GP_DEPOT'        , st, 8, 3, tcTrim, false);

    	  ///////////////////////////////////////////////
    	  // Champs par  défaut
    	  //////////////////////////////////////////////
    	  Tob_Piece.PutValue ('GP_VIVANTE'       , '-');       // Pièce vivante    : Non
    	  Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
    	  Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
    	  Tob_Piece.PutValue ('GP_VENTEACHAT'    , 'VEN');     // Document         : Achat
    	  Tob_Piece.PutValue ('GP_EDITEE'        , 'X');        // Pièce éditée     : Oui
    	  Tob_Piece.PutValue ('GP_NUMADRESSELIVR', -1);
    	  Tob_Piece.PutValue ('GP_NUMADRESSEFACT', -1);
        Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
  		  Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
//        Tob_Piece.PutValue ('GP_SOCIETE'       , V_PGI.NomSociete);
        Tob_Piece.PutValue ('GP_DATECREATION'   , Date) ;
			  Tob_Piece.PutValue ('GP_HEURECREATION'  , NowH) ;
        Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('REP',Tob_Tiers.GetValue('T_ZONECOM'))) ;

    	  DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
			  Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
			  Tob_Piece.PutValue ('GP_PERIODE',Per) ;
    	  Tob_Piece.PutValue ('GP_SEMAINE',Sem) ;
        //
        // Création d'une adresse de facturation si besoin
        //
        Tob_AdrFact := TOB.Create ('ADRESSES', nil, -1);
        Tob_AdrFact.PutValue       ('ADR_TYPEADRESSE', 'PIE');
        Tob_AdrFact.PutFixedStValue('ADR_JURIDIQUE'  , st, 37,  4, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_LIBELLE'    , st, 43, 32, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_ADRESSE1'   , st, 75, 32, tcTrim, false);
      end;
    end
    else if Copy(st,1,5)='DENC2' then
    begin
       if (NombreErreurs (-1) = 0) then
       begin
         Tob_AdrFact.PutFixedStValue('ADR_ADRESSE2'   , st, 17, 32, tcTrim, false);
         Tob_AdrFact.PutFixedStValue('ADR_ADRESSE1'   , st, 49, 32, tcTrim, false);
         Tob_AdrFact.PutFixedStValue('ADR_CODEPOSTAL' , st, 81,  7, tcTrim, false);
       end;
    end
    else if Copy(st,1,5)='DENC3' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        Tob_AdrFact.PutFixedStValue('ADR_VILLE'     , st, 17, 32, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_PAYS'      , st, 49,  3, tcTrim, false);
        Tob_AdrFact.PutFixedStValue('ADR_TELEPHONE' , st, 52, 20, tcTrim, false);
        //
        // Reprise du code client à facturer
        //
        if copy (st, 80, 6) <> '      ' then
        begin
          CodeClient := copy (st, 80, 6);

          if RefCli='0' then CodeClient:=copy(st, 8, 3) + copy(st,80,6)
    	    else if RefCli='1' then CodeClient:=copy(st,80,6);

          Tob_Piece.PutFixedStValue('GP_TIERSFACTURE', st, 80, 6, tcTrim, false);
        end;
      end;
    end
    else if Copy(st,1,5)='DENC4' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Prénom du client
        //
        Tob_AdrFact.PutFixedStValue('ADR_LIBELLE2', st, 56, 32, tcTrim, false);
        //
        // Devise du document
        //
        ////////////////////////////////////////////
      	// Devise du document
      	/////////////////////////////////////////////
        CodeDevGB  := Trim (copy (st, 23, 3));
        CodeDevPGI := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_DEVISE', CodeDevGB, Tob_Devise);

        DevisePiece := CodeDevPGI;
        if VH^.TenueEuro then
        begin
          if DevisePiece = V_PGI.DeviseFongible then
          begin
            Tob_Piece.PutValue ('GP_DEVISE'      , 'EUR');
            Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
          end else
          begin
            Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
          end;
        end else
        begin
          if DevisePiece = 'EUR' then
          begin
            Tob_Piece.PutValue ('GP_DEVISE'      , V_PGI.DevisePivot);
            Tob_Piece.PutValue ('GP_SAISIECONTRE', 'X');
          end else
          begin
            Tob_Piece.PutValue ('GP_DEVISE'      , DevisePiece);
          end;
        end;

        DEV.Code:=DevisePiece;
       	GetInfosDevise(DEV) ;
      	DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
      	PutValueDetail(Tob_Piece,'GP_TAUXDEV',DEV.Taux) ;
        AttribCotation(Tob_Piece) ;
			  PutValueDetail(Tob_Piece,'GP_DATETAUXDEV',DEV.DateTaux) ;

        //
        // Facture ou avoir ?????
        //
        if copy (st, 32, 1) = 'A' then
        begin
          AvoirNegoce := True ;
        end;
        ///
        // Recherche du mode de paiement PGI
        ///
        if AvoirNegoce then
        begin
          ModeRegl := 'EMA' ;
          Tob_Piece.PutValue('GP_MODEREGLE', ModeRegl);
          Tob_Tiers.PutValue('T_MODEREGLE' , ModeRegl) ;
        end
        else begin
          CodeModGB  := Trim (copy (st, 38, 3));

          if CodeModGB <> '' then
          begin
            ModeRegl := TraduitGPAOenPGI(TypeRepriseGB2000, 'GP_MODEREGLE', CodeModGB, Tob_Reglement);
          end else
          begin
            // Si le mode de règlement n'est pas renseigné, on prend celui par défaut
            ModeRegl := Tob_Param.GetValue ('MODEREGLVTE');
          end;

          Q := OpenSQL('Select MR_MODEREGLE from MODEREGL Where MR_MODEREGLE="' + ModeRegl + '" ',False) ;
	  if Q.EOF then
          begin
            NombreErreurs (1);
            EcrireInfo ('   Erreur -> Le mode de règlement ' + ModeRegl + ' est inexixtant', True, True, True);
      	  end else
          begin
            Tob_Piece.PutValue ('GP_MODEREGLE', ModeRegl);
          end;
          Ferme (Q);
        end;
        //
        // Facture TTC ou HT ?????????
        //
        if copy (st, 51, 1) = 'T' then Tob_Piece.PutValue ('GP_FACTUREHT', '-')
        else Tob_Piece.PutValue ('GP_FACTUREHT', 'X') ;

        if copy (st, 20, 3) <> '   ' then  Tob_Piece.PutFixedStValue('GP_REPRESENTANT', st, 20, 3, tcTrim, false);
      end;
    end
    else if Copy(st,1,5)='DENC5' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        /// ?????????
        QteTotFac := StrToFloat (copy (st, 33, 8));
        if QteTotFac = 0 then
        begin
          EcrireInfo ('   Attention -> Facture ' + copy (st, 11, 6) + ' : Total des quantités à 0 -> Facture non reprise', True, True, True);
          NombreErreurs (1);
        end;
      end;
    end
    else if Copy(st,1,5)='DLAC1' then
    begin
  	  if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;

        //
     	  // Création de la ligne commentaire précédente ???
        //
        if (LigneCommACreer = True) then
        begin
          LigneCommACreer:=False;
          CreeLigneCommentaire;
        end;
     	  //
     	  // Gestion d'une ligne commentaire avec l'article générique ????
     	  //
        NumLigne        := NumLigne+1;
     	NumLigneComment := Numligne;
     	LigneCommACreer := True;

        TotQte:=0.0; TotPrix:=0.0; TotQteReste:=0.0;
     	//
        // formatage de la référence article générique
     	//
     	CodeArticle := Copy (st, 23, 12) ;
     	CodeArt     := CodeArticle + '      ' + '   ' + '   ' + '   '+ '   ' + '   ' + 'X';
     	//
     	// Récupération du libellé de l'article
     	//
     	SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
     	Q:=OpenSQL(SQL, True) ;
     	if Not Q.EOF then Designation := Q.FindField('GA_LIBELLE').AsString
     	else begin
       	  EcrireInfo ('   Erreur -> Problème pour récupérer la désignation de l''article générique '+ Copy(st,23,12), True, True, True);
      	  NombreErreurs (1);
        end;
        Ferme(Q) ;
        //
        // Provenance de la facture ?????????
        //
        FactureSurCommande  := False ;
        FactureSurLivraison := False ;
        if copy (st, 89, 1) = 'C' then FactureSurCommande := True
        else if copy (st, 89, 1) = 'B' then FactureSurLivraison := True ;

        if FactureSurCommande or FactureSurLivraison then
        begin
          PieceAGenerer := True;
          BtqNeg2000 := copy (st, 90, 3);
          NumNeg2000 := copy (st, 93, 6);
        end else
        begin
          BtqNeg2000 := '';
          NumNeg2000 := '';
        end;
        //
        // Prix unique ????
        //
        if copy (st, 85, 1) = 'O' then PrixUnique := True
        else PrixUnique := False;
        PrixVenteNeg := 0;
      end;
    end
    else if Copy(st,1,5)='DLAC2' then
    begin
         // Rien à faire
    end
    else if Copy(st,1,5)='DLAC3' then
    begin
         // Rien à faire
    end
    else if Copy(st,1,5)='DTIC1' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;
        ////////////////////////////////////////////////////////////////
    	// Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    	////////////////////////////////////////////////////////////////
    	Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
    	//Tob_Ligne.initValeurs;
    	////////////////////////////////////////////////////////////////
    	// Initialise les champs de la ligne à partir de l'entête
    	////////////////////////////////////////////////////////////////
    	PieceVersLigne (Tob_Piece, Tob_Ligne);
    	Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
    	Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
    	//////////////////////////////////////////////
    	// Nouveau numéro de ligne
    	//////////////////////////////////////////////
    	NumLigne := NumLigne+1;
    	Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
    	Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
    	////////////////////////////////////////////////////////////////
    	// Ajoute des champs supplémentaires pour le calcul des cumuls
    	////////////////////////////////////////////////////////////////
    	AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
    	//////////////////////////////////////////////////////////
    	// Détermination de l'article taille
    	/////////////////////////////////////////////////////////
    	CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');

        Taille := StrToInt (copy (st, 23 ,2));
        rang   := CalculTaillePGI (Taille);

    	CodeDim1 := '   '; CodeDim2 := '   '; CodeDim3 := '   ';
    	CodeDim4 := '   '; CodeDim5 := '   ';
    	if (CategorieDim [3] = '1') then Codedim1 :=rang
    	else if (CategorieDim [3] = '2') then Codedim2 :=rang
    	else if (CategorieDim [3] = '3') then Codedim3 :=rang
    	else if (CategorieDim [3] = '4') then Codedim4 :=rang
    	else if (CategorieDim [3] = '5') then Codedim5 :=rang;
    	Codeart := CodeArticle + '      ' + CodeDim1 + CodeDim2 + CodeDim3 + CodeDim4 + CodeDim5 + 'X'; // Code article PGI

    	SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
    	Q:=OpenSQL(SQL,True) ;
    	if Not Q.EOF then
    	begin
    	  if (Tob_Article=nil) then
      	  begin
            Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
     	    //Tob_Article.initValeurs;
          end;
	  /////////////////////////////////////////////////////////////////////////////
      	  // Initialisation des champs de la ligne à partir de l'article dimensionné
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	  Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	  Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	  Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	  Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
          Tob_Ligne.PutValue( 'GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
          Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
          Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	      Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	      Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
          for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
          for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
          Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
      	  /////////////////////////////////////////////////////////////////////////////
      	  // Ajout de l'article dans la TOB Article
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Article.SelectDB('',Q);
    	  end
    	  else begin
      	  EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	  NombreErreurs (1);
        end;
        Ferme(Q) ;
      end;

      if (NombreErreurs (-1) = 0) then
      begin
	   Tob_Ligne.PutValue ('GL_ARTICLE'     , Codeart);
  	   Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

        GL_CodeArticle := Trim (CodeArticle);
      	Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	    Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

    	Tob_Ligne.PutValue ('GL_TARIF'       , 0);
    	Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	    Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
        Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
        Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

        QteRec := StrtoFloat (copy (st, 30, 5));
        if AvoirNegoce then QteRec:=-QteRec;
    	Tob_Ligne.PutValue('GL_QTESTOCK' , QteRec);
    	Tob_Ligne.PutValue('GL_QTEFACT'  , QteRec);
        Tob_Ligne.PutValue('GL_QTERESTE' , QteRec);

	    DevisePiece := Tob_Piece.GetValue ('GP_DEVISE');
        if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then
        begin
          if VH^.TenueEuro then DevisePiece := V_PGI.DeviseFongible
          else                  DevisePiece := 'EUR';
        end;
        //
        // PA de l'article
        //
        Montant := StrToFloat (copy (st, 40, 8));
        Montant := Montant/100 ;

        if  DevisePiece = V_PGI.DevisePivot then Tob_Ligne.PutValue('GL_DPA', Montant)
        else begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Ligne.PutValue ('GL_DPA', MontantDev);
        end;
        //
        // PAMP de l'article
        //
        Montant := StrToFloat (copy (st, 48, 8));
        Montant := Montant/100 ;

        if  DevisePiece = V_PGI.DevisePivot then
        begin
          Tob_Ligne.PutValue('GL_PMAP'     , Montant);
          Tob_Ligne.PutValue('GL_PMAPACTU' , Montant);
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Ligne.PutValue ('GL_PMAP'    , MontantDev);
          Tob_Ligne.PutValue ('GL_PMAPACTU', MontantDev);
        end;
        //
        // PR de l'article
        //
        Montant := StrToFloat (copy (st, 72, 8));
        Montant := Montant/100 ;

        if  DevisePiece = V_PGI.DevisePivot then Tob_Ligne.PutValue('GL_DPR', Montant)
        else begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Ligne.PutValue ('GL_DPR', MontantDev);
        end;

        //
        // PRMP de l'article
        //
        Montant := StrToFloat (copy (st, 80, 8));
        Montant := Montant/100 ;

        if  DevisePiece = V_PGI.DevisePivot then
        begin
          Tob_Ligne.PutValue('GL_PMRP'     , Montant);
          Tob_Ligne.PutValue('GL_PMRPACTU' , Montant);
        end else
        begin
          DateConv   := NowH ;
          MontantDev := 0 ;  // initialisation bidon
          MontantOld := 0 ;  // initialisation bidon
          ToxConvertir (Montant, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, Tob_DevisePGI) ;
          Tob_Ligne.PutValue ('GL_PMRP'     , MontantDev);
          Tob_Ligne.PutValue ('GL_PMRPACTU' , MontantDev);
        end;

        //
        // Gestion du prix unique
        //
        if PrixUnique = True then
        begin
          if PrixVenteNeg = 0 then
          begin
            PrixVenteNeg := StrToFloat (copy (st, 64,8));
            PrixVenteNeg := PrixVenteNeg/100;
          end;

          if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
          begin
            Tob_Ligne.PutValue('GL_PUHTDEV'   , PrixVenteNeg);
    	    Tob_Ligne.PutValue('GL_PUHTNETDEV', PrixVenteNeg);
          end else
          begin
            Tob_Ligne.PutValue('GL_PUTTCDEV'   , PrixVenteNeg);
    	    Tob_Ligne.PutValue('GL_PUTTCNETDEV', PrixVenteNeg);
          end;
        end else
        begin
          if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
          begin
            Tob_Ligne.PutFixedStValue('GL_PUHTDEV'   , st, 64, 8, tcDouble100, false);
    	    Tob_Ligne.PutFixedStValue('GL_PUHTNETDEV', st, 64, 8, tcDouble100, false);
          end else
          begin
            Tob_Ligne.PutFixedStValue('GL_PUTTCDEV'   , st, 64, 8, tcDouble100, false);
    	    Tob_Ligne.PutFixedStValue('GL_PUTTCNETDEV', st, 64, 8, tcDouble100, false);
          end;
        end;
  	    Tob_Ligne.PutValue         ('GL_PRIXPOURQTE', 1);
        Tob_Ligne.PutValue         ('GL_PCB'        , 1);
  	    Tob_Ligne.PutValue         ('GL_TYPEDIM'    , 'DIM');
    	Tob_Ligne.PutValue         ('GL_QUALIFMVT'  , QualifMvt) ;

        if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then Prix:=Tob_Ligne.GetValue('GL_PUHTDEV')
    	else Prix:=Tob_Ligne.GetValue('GL_PUTTCDEV') ;

        Qte:=Arrondi(Tob_Ligne.GetValue('GL_QTEFACT'),6) ;
        TotQteReste :=TotQteReste+Qte;
  	    TotQte      :=TotQte+Qte;
	    TotPrix     :=TotPrix+Qte*Prix;

        ////////////////////////////////////////////////////////////////////////////////////////////
    	// Ajoute des champs supplémentaires pour la génération des commandes ou livraison initiales
    	////////////////////////////////////////////////////////////////////////////////////////////
        if FactureSurCommande then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'C')
        else if FactureSurLivraison then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'L') ;

        Tob_Ligne.AddChampSupValeur('$$_CODEBTQ' , BtqNeg2000);
        Tob_Ligne.AddChampSupValeur('$$_NUMDOC'  , NumNeg2000);
      end;
    end
    else if Copy(st,1,5)='DLRC1' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;

        //  Gestion de la remise
        Remise := StrToFloat (copy (st, 72, 12));
        Remise := Remise/100 ;
        if Remise <> 0 then
        begin
          Tob_Piece.PutValue ('GP_REMISEPIED', Remise);
        end;
      end;
    end
    else if Copy(st,1,5)='DLXC1' then             // Ligne de Port : ne marche pas en V530
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;
        //
        // Recherche du port dans la table des ports
        //
        CodePort := copy (st, 43, 3);
        SQL:='Select * From PORT WHERE GPO_CODEPORT="'+CodePort+'"';
    	Q:=OpenSQL(SQL,True) ;

    	  if Not Q.EOF then
    	  begin
          //
          // Création de la TOB des ports
          //
    		  if (Tob_PiedPort=nil) then
      	  begin
            Tob_PiedPort := TOB.CREATE ('Les ports', NIL, -1);
     	    //Tob_PiedPort.initValeurs;
          end;
          //
          //  Chargement en TOB du descriptif du Port
          //
          TobPort := TOB.Create ('PORT', nil, -1);
          TobPort.SelectDB ('', Q);
          //
          // Alimentation de la TOB PIEDPORT
          //
          TobFille := Tob.create ('PIEDPORT', Tob_PiedPort, -1);
          //TobFille.InitValeurs ;

          TobFille.PutValue ('GPT_CODEPORT'     , CodePort);
          TobFille.PutValue ('GPT_LIBELLE', TobPort.GetValue('GPO_LIBELLE'));

          TypePort := TobPort.GetValue('GPO_TYPEPORT') ;
          TobFille.PutValue ('GPT_TYPEPORT'     , TypePort);
          TobFille.PutValue ('GPT_FAMILLETAXE1' , TobPort.GetValue('GPO_FAMILLETAXE1'));
          TobFille.PutValue ('GPT_FAMILLETAXE2' , TobPort.GetValue('GPO_FAMILLETAXE2'));
          TobFille.PutValue ('GPT_FAMILLETAXE3' , TobPort.GetValue('GPO_FAMILLETAXE3'));
          TobFille.PutValue ('GPT_FAMILLETAXE4' , TobPort.GetValue('GPO_FAMILLETAXE4'));
          TobFille.PutValue ('GPT_FAMILLETAXE5' , TobPort.GetValue('GPO_FAMILLETAXE5'));
          TobFille.PutValue ('GPT_COMPTAARTICLE', TobPort.GetValue('GPO_COMPTAARTICLE'));

          Base    := TobPort.GetValue('GPO_MINIMUM') ;
          TobFille.PutValue ('GPT_MINIMUM', Base);
          TobFille.PutValue ('GPT_FRANCO' , TobPort.GetValue('GPO_FRANCO'));
          TobFille.PutValue ('GPT_MONTANTMINI', 0);
          if TypePort='MT' then
          begin
            TobFille.PutValue      ('GPT_POURCENT', 0) ;
            Base:=TobPort.GetValue ('GPO_PVHT') ;
            TobFille.PutValue      ('GPT_MONTANTMINI', Base) ;
          end else
          begin
            TobFille.PutValue ('GPT_POURCENT', TobPort.GetValue('GPO_COEFF'));
          end;
          if TypePort='MI' then
          begin
            Base:=TobPort.GetValue('GPO_PVHT') ;
            TobFille.PutValue ('GPT_MONTANTMINI', Base) ;
          end ;
          TobFille.PutValue ('GPT_NATUREPIECEG' , Tob_Piece.GetValue ('GP_NATUREPIECEG'));
          TobFille.PutValue ('GPT_SOUCHE'       , Tob_Piece.GetValue ('GP_SOUCHE'));
          TobFille.PutValue ('GPT_NUMERO'       , Tob_Piece.GetValue ('GP_NUMERO'));
          TobFille.PutValue ('GPT_INDICEG'      , Tob_Piece.GetValue ('GP_INDICEG'));
          TobFille.PutValue ('GPT_DEVISE'       , Tob_Piece.GetValue ('GP_DEVISE'));
          Inc (NumLigPort);
          TobFille.PutValue ('GPT_NUMPORT'      , NumLigPort);

          //
          // Récupération des montants dans le fichier
          //
          //TobFille.PutFixedStValue('GPT_BASEHT'    , st, 73, 12, tcDouble100, false);
          //TobFille.PutFixedStValue('GPT_BASEHTDEV' , st, 73, 12, tcDouble100, false);
          //TobFille.PutFixedStValue('GPT_BASETTC'   , st, 97, 12, tcDouble100, false);
          //TobFille.PutFixedStValue('GPT_BASETTCDEV', st, 97, 12, tcDouble100, false);
          TobFille.PutFixedStValue('GPT_TOTALHT'    , st, 73, 12, tcDouble100, false);
          TobFille.PutFixedStValue('GPT_TOTALHTDEV' , st, 73, 12, tcDouble100, false);
        end else
        begin
          EcrireInfo ('   Erreur -> Problème pour récupérer le PORT '+ CodePort, True, True, True);
       	  NombreErreurs (1);
        end;
        Ferme (Q);
      end;
    end
    else if Copy(st,1,5)='DLPC1' then
    begin
      if (NombreErreurs (-1) = 0) then
      begin
        //
        // Facture ou avoir ???
        //
        if AvoirNegoce then NaturePiece := NaturePieceAvo
        else NaturePiece := NaturePieceFac;
        //
     	  // Création de la ligne commentaire précédente ???
        //
        if (LigneCommACreer = True) then
        begin
          LigneCommACreer:=False;
          CreeLigneCommentaire;
        end;

        ////////////////////////////////////////////////////////////////
        // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
        ////////////////////////////////////////////////////////////////
        Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, -1);
        //Tob_Ligne.initValeurs;
        ////////////////////////////////////////////////////////////////
        // Initialise les champs de la ligne à partir de l'entête
        ////////////////////////////////////////////////////////////////
        PieceVersLigne (Tob_Piece, Tob_Ligne);
        Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
        Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
        //////////////////////////////////////////////
        // Nouveau numéro de ligne
        //////////////////////////////////////////////
        NumLigne := NumLigne+1;
        Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
        Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
        ////////////////////////////////////////////////////////////////
        // Ajoute des champs supplémentaires pour le calcul des cumuls
        ////////////////////////////////////////////////////////////////
        AddLesSupLigne  (Tob_Ligne, False) ; // Ajout Champs spécifiques pour calcul de la pièce
        //////////////////////////////////////////////////////////
        // Détermination de l'article taille
        /////////////////////////////////////////////////////////
        CodeArticle := copy (st, 43, 3) + '         ' ;

        Codeart := CodeArticle + '      ' + '   ' + '   ' + '   ' + '   ' + '   ' + 'X'; // Code article PGI

        SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+Codeart+'"';
        Q:=OpenSQL(SQL,True) ;
        if Not Q.EOF then
        begin
          if (Tob_Article=nil) then
          begin
            Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
            //Tob_Article.initValeurs;
          end;
	  /////////////////////////////////////////////////////////////////////////////
      	  // Initialisation des champs de la ligne à partir de l'article dimensionné
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Ligne.PutValue ('GL_PRIXPOURQTE'   , Q.FindField('GA_PRIXPOURQTE').AsString);
      	  Tob_Ligne.PutValue ('GL_ESCOMPTABLE'   , Q.FindField('GA_ESCOMPTABLE').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLEPIED' , Q.FindField('GA_REMISEPIED').AsString);
      	  Tob_Ligne.PutValue ('GL_REMISABLELIGNE', Q.FindField('GA_REMISELIGNE').AsString);
      	  Tob_Ligne.PutValue ('GL_TENUESTOCK'    , Q.FindField('GA_TENUESTOCK').AsString);
      	  Tob_Ligne.PutValue ('GL_TARIFARTICLE'  , Q.FindField('GA_TARIFARTICLE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFSURFACE' , Q.FindField('GA_QUALIFSURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFVOLUME'  , Q.FindField('GA_QUALIFVOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFPOIDS'   , Q.FindField('GA_QUALIFPOIDS').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFLINEAIRE', Q.FindField('GA_QUALIFLINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFHEURE'   , Q.FindField('GA_QUALIFHEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_SURFACE'       , Q.FindField('GA_SURFACE').AsString);
      	  Tob_Ligne.PutValue ('GL_VOLUME'        , Q.FindField('GA_VOLUME').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSBRUT'     , Q.FindField('GA_POIDSBRUT').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSNET'      , Q.FindField('GA_POIDSNET').AsString);
      	  Tob_Ligne.PutValue ('GL_POIDSDOUA'     , Q.FindField('GA_POIDSDOUA').AsString);
      	  Tob_Ligne.PutValue ('GL_LINEAIRE'      , Q.FindField('GA_LINEAIRE').AsString);
      	  Tob_Ligne.PutValue ('GL_HEURE'         , Q.FindField('GA_HEURE').AsString);
      	  Tob_Ligne.PutValue ('GL_QUALIFQTESTO'  , Q.FindField('GA_QUALIFUNITESTO').AsString);
          Tob_Ligne.PutValue( 'GL_COLLECTION'    , Q.FindField('GA_COLLECTION').AsString) ;
          Tob_Ligne.PutValue( 'GL_FOURNISSEUR'   , Q.FindField('GA_FOURNPRINC').AsString) ;
          Tob_Ligne.PutValue ('GL_FAMILLENIV1'   , Q.FindField('GA_FAMILLENIV1').AsString) ;
	      Tob_Ligne.PutValue ('GL_FAMILLENIV2'   , Q.FindField('GA_FAMILLENIV2').AsString) ;
  	      Tob_Ligne.PutValue( 'GL_FAMILLENIV3'   , Q.FindField('GA_FAMILLENIV3').AsString) ;
          for i:=1 to 5 do Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),Q.FindField('GA_FAMILLETAXE'+IntToStr(i)).asString) ;
          for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),Q.FindField('GA_LIBREART'+IntToStr(i)).asString) ;
          Tob_Ligne.PutValue( 'GL_LIBREARTA'   , Q.FindField('GA_LIBREARTA').AsString) ;
          Tob_Ligne.PutValue( 'GL_LIBELLE'     , Q.FindField('GA_LIBELLE').AsString) ;
      	  /////////////////////////////////////////////////////////////////////////////
      	  // Ajout de l'article dans la TOB Article
      	  /////////////////////////////////////////////////////////////////////////////
      	  Tob_Article.SelectDB('',Q);
    	  end
    	  else begin
      	  EcrireInfo ('   Erreur -> Problème pour récupérer la fiche article par taille '+ Codeart, True, True, True);
       	  NombreErreurs (1);
        end;
    	Ferme(Q) ;
      end;

      if (NombreErreurs (-1) = 0) then
      begin
	Tob_Ligne.PutValue ('GL_ARTICLE'     , Codeart);
  	//Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

        GL_CodeArticle := Trim (CodeArticle);
      	Tob_Ligne.PutValue ('GL_CODEARTICLE' , GL_CodeArticle);
	Tob_Ligne.PutValue ('GL_REFARTSAISIE', GL_CodeArticle);

    	Tob_Ligne.PutValue ('GL_TARIF'       , 0);
    	Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
	//Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'PRE');
        Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
        Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'NON') ;
        Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');

        QteRec := 1.0;                         // Quantité Port : forcément 1
        if AvoirNegoce then QteRec:=-QteRec;
     	Tob_Ligne.PutValue('GL_QTESTOCK', QteRec);
    	Tob_Ligne.PutValue('GL_QTEFACT' , QteRec);
        Tob_Ligne.PutValue('GL_QTERESTE', QteRec);

  	    //////////////////////////////////////////////////////////
  	    // Récupération des PA et PR
    	//////////////////////////////////////////////////////////
  	    Tob_Ligne.PutFixedStValue('GL_DPA'     , st, 73, 12, tcDouble100, false);
  	    Tob_Ligne.PutFixedStValue('GL_PMAP'    , st, 73, 12, tcDouble100, false);
    	Tob_Ligne.PutFixedStValue('GL_PMAPACTU', st, 73, 12, tcDouble100, false);
	    Tob_Ligne.PutFixedStValue('GL_DPR'     , st, 73, 12, tcDouble100, false);
  	    Tob_Ligne.PutFixedStValue('GL_PMRP'    , st, 73, 12, tcDouble100, false);
    	Tob_Ligne.PutFixedStValue('GL_PMRPACTU', st, 73, 12, tcDouble100, false);

        if Tob_Ligne.GetValue('GL_FACTUREHT')='X'  then
        begin
          Tob_Ligne.PutFixedStValue('GL_PUHTDEV'   , st, 73, 12, tcDouble100, false);
    	  Tob_Ligne.PutFixedStValue('GL_PUHTNETDEV', st, 73, 12, tcDouble100, false);
        end else
        begin
          Tob_Ligne.PutFixedStValue('GL_PUTTCDEV'   , st, 97, 12, tcDouble100, false);
    	  Tob_Ligne.PutFixedStValue('GL_PUTTCNETDEV', st, 97, 12, tcDouble100, false);
        end;
        Tob_Ligne.PutValue         ('GL_PRIXPOURQTE'    , 1);
        Tob_Ligne.PutValue         ('GL_PCB'            , 1);
        Tob_Ligne.PutValue         ('GL_TVAENCAISSEMENT', 'X');
  	    Tob_Ligne.PutValue         ('GL_TYPEDIM'        , 'NOR');
    	Tob_Ligne.PutValue         ('GL_QUALIFMVT'      , QualifMvt) ;
        Tob_Ligne.PutValue         ('GL_DEPOT'          , '') ;

        //if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then Prix:=Tob_Ligne.GetValue('GL_PUHTDEV')
    	//else Prix:=Tob_Ligne.GetValue('GL_PUTTCDEV') ;

        ////////////////////////////////////////////////////////////////////////////////////////////
    	// Ajoute des champs supplémentaires pour la génération des commandes ou livraison initiales
    	////////////////////////////////////////////////////////////////////////////////////////////
        //if FactureSurCommande then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'C')
        //else if FactureSurLivraison then Tob_Ligne.AddChampSupValeur('$$_TYPEDOC' , 'L') ;

        //Tob_Ligne.AddChampSupValeur('$$_CODEBTQ' , BtqNeg2000);
        //Tob_Ligne.AddChampSupValeur('$$_NUMDOC'  , NumNeg2000);
      end;
    end;
  end;
end;


procedure TFRecupGB2000.bFinClick(Sender: TObject);
begin
inherited;
// Enregistrement dans la registry du nom du repertoire utilisé
SaveSynRegKey('IMPORTGB2000REPFICH', RepFich.Text, True) ;
if not isInside(Self) then close ;
end;
 
procedure TFRecupGB2000.bSuivantClick(Sender: TObject);
begin
inherited;
if (P.ActivePage.PageIndex = P.PageCount - 1)
   then bFin.enabled := True;
end;

procedure TFRecupGB2000.bPrecedentClick(Sender: TObject);
begin
inherited;
bFin.enabled := False;
end;

procedure TFRecupGB2000.DeleteTablesLibres(typ : string; num : Integer);
begin
if typ='A' then
   begin
   CompteRendu.lines.add('   Suppression de la table libre '+InttoStr(num)+' des articles');
   if num=10 then ExecuteSQL ('delete from CHOIXEXT where YX_TYPE="LAA"')
             else ExecuteSQL ('delete from CHOIXEXT where YX_TYPE="LA'+InttoStr(num)+'"');
   end
else if typ='C' then
   begin
   CompteRendu.lines.add('   Suppression de la table libre '+InttoStr(num)+' des tiers');
   if num=10 then ExecuteSQL ('delete from CHOIXEXT where YX_TYPE="LTA"')
             else ExecuteSQL ('delete from CHOIXEXT where YX_TYPE="LT'+InttoStr(num)+'"');
   end
end;

procedure TFRecupGB2000.NumPremBtqExit(Sender: TObject);
var Num : Integer;
begin
inherited;
Num := StrToInt (NumPremBtq.Text);
NumPremBtq.Text := FormatFloat ('000', Num);
if (NumDerBtq.Text<NumPremBtq.Text) then NumDerBtq.Text:=NumPremBtq.text;
end;

procedure TFRecupGB2000.NumDerBtqExit(Sender: TObject);
var Num : Integer;
begin
inherited;
Num := StrToInt (NumDerBtq.Text);
NumDerBtq.Text := FormatFloat ('000', Num);
if (NumDerBtq.Text<NumPremBtq.Text) then NumDerBtq.Text:=NumPremBtq.text;
end;

procedure TFRecupGB2000.FichAnoElipsisClick(Sender: TObject);
begin
inherited;
if OpenFichAno.execute then FichAno.text:=OpenFichAno.FileName;
end;

procedure TFRecupGB2000.bStopButton(Sender: TObject);
begin
inherited;
ArretRecup := TRUE;
BStop.visible := False;
EcrireInfo ('Récupération interrompue !', True, True, False);
Application.ProcessMessages ;
end;

procedure TFRecupGB2000.bVerifLesPieces(Sender: TObject);
begin
inherited;
  VerifPiece;
end;

procedure TFRecupGB2000.RepArticlesClick(Sender: TObject);
begin
inherited;
  if RepArticles.Checked then ButtonParamArticles.Enabled:=TRUE
  else ButtonParamArticles.Enabled:=FALSE;
end;

procedure TFRecupGB2000.RepClientsClick(Sender: TObject);
begin
inherited;
  if RepClients.Checked then ButtonParamClients.Enabled:=TRUE
  else ButtonParamClients.Enabled:=FALSE;
end;

procedure TFRecupGB2000.RepFournisseursClick(Sender: TObject);
begin
inherited;
  if RepFournisseurs.Checked then ButtonParamFrn.enabled:=TRUE
  else ButtonParamFrn.Enabled:=FALSE;
end;

procedure TFRecupGB2000.CtrlTableTypeArt ;
var TPST: TOB;
    Cpt: Integer;
    RepTypArticle : string;
begin

  RepTypArticle := Tob_Param.GetValue ('REPTYP');

  if RepTypArticle='' then exit;
  for cpt:=1 to 10 do
    begin
    if (cpt=10) and (RepTypArticle='GA_LIBREARTA') then break
    else if RepTypArticle='GA_LIBREART'+InttoStr(cpt) then break ;
  end ;
  if cpt<11 then
  begin
    TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
    //TPST.initValeurs;
    if cpt=10 then TPST.PutValue('YX_TYPE', 'LAA')
    else TPST.PutValue('YX_TYPE', 'LA'+InttoStr(Cpt));
    TPST.PutValue('YX_CODE', 'N');
    TPST.PutValue('YX_LIBELLE', 'Article Normal');
    TPST.PutValue('YX_ABREGE',  'Article Normal');
    ImportTablesLibres (TPST);
    TPST.PutValue('YX_CODE', 'P');
    TPST.PutValue('YX_LIBELLE', 'Article Permanent');
    TPST.PutValue('YX_ABREGE',  'Article Permanent');
    ImportTablesLibres (TPST);
    TPST.PutValue('YX_CODE', 'S');
    TPST.PutValue('YX_LIBELLE', 'Article Suivi');
    TPST.PutValue('YX_ABREGE',  'Article Suivi');
    ImportTablesLibres (TPST);
    TPST.Free;
  end;         
end;

procedure TFRecupGB2000.CtrlTableEtatCpta ;
var TPST: TOB;
    Cpt: Integer;
    EtatCpta : string ;
begin
  EtatCpta := Tob_Param.GetValue ('REPETACPTA');
  if EtatCpta='' then exit;
  for cpt:=1 to 10 do
  begin
    if (cpt=10) and (EtatCpta='YTC_TABLELIBRETIERSA') then break
    else if EtatCpta='YTC_TABLELIBRETIERS'+InttoStr(cpt) then break ;
  end ;
  if cpt<11 then
  begin
    TPST := TOB.CREATE ('CHOIXEXT', NIL, -1);  // Nom de table - Parent - Indice (pour insertion à position donnée)
    //TPST.initValeurs;
    if cpt=10 then TPST.PutValue('YX_TYPE', 'LTA')
    else TPST.PutValue('YX_TYPE', 'LT'+InttoStr(Cpt));
    TPST.PutValue('YX_CODE', 'V');
    TPST.PutValue('YX_LIBELLE', 'Vert');
    TPST.PutValue('YX_ABREGE',  'Vert');
    ImportTablesLibres (TPST);
    TPST.PutValue('YX_CODE', 'O');
    TPST.PutValue('YX_LIBELLE', 'Orange');
    TPST.PutValue('YX_ABREGE',  'Orange');
    ImportTablesLibres (TPST);
    TPST.PutValue('YX_CODE', 'R');
    TPST.PutValue('YX_LIBELLE', 'Rouge');
    TPST.PutValue('YX_ABREGE',  'Rouge');
    ImportTablesLibres (TPST);
    TPST.Free;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Suppression des tables avant intégration
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.SupprimeTables ;
var StDim1, Stdim2, Stdim3, Stdim4, StDim5 : string;
begin
StDim1:=RechDom('GCCATEGORIEDIM','DI1',FALSE) ;
StDim2:=RechDom('GCCATEGORIEDIM','DI2',FALSE) ;
StDim3:=RechDom('GCCATEGORIEDIM','DI3',FALSE) ;
StDim4:=RechDom('GCCATEGORIEDIM','DI4',FALSE) ;
StDim5:=RechDom('GCCATEGORIEDIM','DI5',FALSE) ;

CompteRendu.lines.add(#13#10'Suppression des tables :');
if ((Tob_Param.GetValue ('DELARRONDIS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des arrondis');
   ExecuteSQL ('delete from ARRONDI');
   end;
if  ((Tob_Param.GetValue ('DELBOUTIQUES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des boutiques');
   ExecuteSQL ('delete from ETABLISS');
   end;
if ((Tob_Param.GetValue ('DELPAYS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des pays');
   ExecuteSQL ('delete from PAYS');
   end;
if ((Tob_Param.GetValue ('DELCODESPOSTAUX')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des codes postaux');
   ExecuteSQL ('delete from CODEPOST');
   end;
if ((Tob_Param.GetValue ('DELDEVISES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des devises');
   ExecuteSQL ('delete from DEVISE');
   end;
if ((Tob_Param.GetValue ('DELFAMILLES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des familles');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="FN1"');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="FN2"');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="FN3"');
   end;
if ((Tob_Param.GetValue ('DELTABLEA1')='1') and not (ArretRecup)) then DeleteTablesLibres('A',1);
if ((Tob_Param.GetValue ('DELTABLEA2')='1') and not (ArretRecup)) then DeleteTablesLibres('A',2);
if ((Tob_Param.GetValue ('DELTABLEA3')='1') and not (ArretRecup)) then DeleteTablesLibres('A',3);
if ((Tob_Param.GetValue ('DELTABLEA4')='1') and not (ArretRecup)) then DeleteTablesLibres('A',4);
if ((Tob_Param.GetValue ('DELTABLEA5')='1') and not (ArretRecup)) then DeleteTablesLibres('A',5);
if ((Tob_Param.GetValue ('DELTABLEA6')='1') and not (ArretRecup)) then DeleteTablesLibres('A',6);
if ((Tob_Param.GetValue ('DELTABLEA7')='1') and not (ArretRecup)) then DeleteTablesLibres('A',7);
if ((Tob_Param.GetValue ('DELTABLEA8')='1') and not (ArretRecup)) then DeleteTablesLibres('A',8);
if ((Tob_Param.GetValue ('DELTABLEA9')='1') and not (ArretRecup)) then DeleteTablesLibres('A',9);
if ((Tob_Param.GetValue ('DELTABLEA10')='1') and not (ArretRecup)) then DeleteTablesLibres('A',10);
if ((Tob_Param.GetValue ('DELTABLEC1')='1') and not (ArretRecup)) then DeleteTablesLibres('C',1);
if ((Tob_Param.GetValue ('DELTABLEC2')='1') and not (ArretRecup)) then DeleteTablesLibres('C',2);
if ((Tob_Param.GetValue ('DELTABLEC3')='1') and not (ArretRecup)) then DeleteTablesLibres('C',3);
if ((Tob_Param.GetValue ('DELTABLEC4')='1') and not (ArretRecup)) then DeleteTablesLibres('C',4);
if ((Tob_Param.GetValue ('DELTABLEC5')='1') and not (ArretRecup)) then DeleteTablesLibres('C',5);
if ((Tob_Param.GetValue ('DELTABLEC6')='1') and not (ArretRecup)) then DeleteTablesLibres('C',6);
if ((Tob_Param.GetValue ('DELTABLEC7')='1') and not (ArretRecup)) then DeleteTablesLibres('C',7);
if ((Tob_Param.GetValue ('DELTABLEC8')='1') and not (ArretRecup)) then DeleteTablesLibres('C',8);
if ((Tob_Param.GetValue ('DELTABLEC9')='1') and not (ArretRecup)) then DeleteTablesLibres('C',9);
if ((Tob_Param.GetValue ('DELTABLEC10')='1') and not (ArretRecup)) then DeleteTablesLibres('C',10);
if ((Tob_Param.GetValue ('DELMASQUEDIM')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des masques de dimensions');
   ExecuteSQL ('delete from DIMMASQUE');
   end ;
if ((Tob_Param.GetValue ('DELGRILLEDIM1')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des grilles de dimensions '+StDim1);
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GG1"');
   end ;
if ((Tob_Param.GetValue ('DELGRILLEDIM2')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des grilles de dimensions '+StDim2);
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GG2"');
   end ;
if ((Tob_Param.GetValue ('DELGRILLEDIM3')='1')  and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des grilles de dimensions '+StDim3);
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GG3"');
   end ;
if ((Tob_Param.GetValue ('DELGRILLEDIM4')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des grilles de dimensions '+StDim4);
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GG4"');
   end ;
if ((Tob_Param.GetValue ('DELGRILLEDIM5')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des grilles de dimensions '+StDim5);
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GG5"');
   end ;
if ((Tob_Param.GetValue ('DELDIM1')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des dimensions '+StDim1);
   ExecuteSQL ('delete from DIMENSION where GDI_TYPEDIM="DI1"');
   end ;
if ((Tob_Param.GetValue ('DELDIM2')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des dimensions '+StDim2);
   ExecuteSQL ('delete from DIMENSION where GDI_TYPEDIM="DI2"');
   end ;
if ((Tob_Param.GetValue ('DELDIM3')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des dimensions '+StDim3);
   ExecuteSQL ('delete from DIMENSION where GDI_TYPEDIM="DI3"');
   end ;
if ((Tob_Param.GetValue ('DELDIM4')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des dimensions '+StDim4);
   ExecuteSQL ('delete from DIMENSION where GDI_TYPEDIM="DI4"');
   end ;
if ((Tob_Param.GetValue ('DELDIM5')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des dimensions '+StDim5);
   ExecuteSQL ('delete from DIMENSION where GDI_TYPEDIM="DI5"');
   end ;
if ((Tob_Param.GetValue ('DELCOLLECTIONS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des collections');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="GCO"');
   end ;
if ((Tob_Param.GetValue ('DELPOIDS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des unités de poids');
   ExecuteSQL ('delete from MEA where GME_QUALIFMESURE="POI"');
   end ;
if ((Tob_Param.GetValue ('DELDEMARQUES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des motifs de démarque');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="QFR"');
   end ;
if ((Tob_Param.GetValue ('DELCIVILITES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des civilités');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="CIV"');
   end ;
if ((Tob_Param.GetValue ('DELREPRESENTANTS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des représentants');
   {ExecuteSQL ('delete from VENDEUR where GVD_TYPEVENDEUR="REP"');}
   end ;
if ((Tob_Param.GetValue ('DELCATEGORIES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add('   Suppression des catégories clients');
   ExecuteSQL ('delete from CHOIXCOD where CC_TYPE="TRC"');
   end ;
if ((Tob_Param.GetValue ('DELARTICLES')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add(#13#10'   Suppression des ARTICLES');
   ExecuteSQL ('delete from ARTICLE');
   end ;
if ((Tob_Param.GetValue ('DELCLIENTS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add(#13#10'   Suppression des CLIENTS');
   ExecuteSQL ('delete from TIERS where T_NATUREAUXI="CLI"');
   end ;
if ((Tob_Param.GetValue ('DELFOURNISSEURS')='1') and not (ArretRecup)) then
   begin
   CompteRendu.lines.add(#13#10'   Suppression des FOURNISSEURS');
   ExecuteSQL ('delete from TIERS where T_NATUREAUXI="FOU"');
   end ;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXPST......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupParametres ;
var Fichier         : textfile;
    Chtmp, AncChtmp : string  ;
    CategorieDim    : string  ;
begin
inherited;
  if ArretRecup then exit;

  CodeInterne       := '';
  DerTableLibreArg  := '';
  DerTableLibreArg2 := '';

  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
    exit;
  end;

  EcrireInfo ('Récupération des PARAMETRES GENERAUX le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
  Chtmp:=RepFich.text+'\EXPST'+NumPremBtq.text;
  if FileExists (Chtmp) then
  begin
    AssignFile(Fichier, Chtmp);
    Reset (Fichier);                                 // Ouverture du fichier
    InitMove (1000, '');
    AncChtmp:='';
    Tob_Parametre:=nil;
    CompteurParam := 0;

    while not EOF(Fichier) and not ArretRecup do
    begin
      MoveCur (False);
      readln(Fichier,Chtmp);
      Application.ProcessMessages ;

      CompteurParam:=CompteurParam+1;
      if CompteurParam>=1000 then
      begin
        if Tob_Parametre<>nil then ImportParametre;
        CompteurParam:=0;
      end;

      if Copy(Chtmp,1,3)='Par' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
          CompteRendu.lines.add('   Récupération des arrondis');
        end;
        RecupArrondi (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pbo' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
        	if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des boutiques');
        end;
        //RecupBoutique (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pde' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
          CompteRendu.lines.add('   Récupération des devises -> Alimentation des tables de correspondances');
        end;
        RecupDevise (Chtmp);
        SavTobDevise ;
      end
      else if Copy(Chtmp,1,3)='Ppa' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des pays -> Alimentation des tables de correspondances');
        end;
        RecupPays (Chtmp);
        SavTobPays;
      end
      else if Copy(Chtmp,1,3)='Pmr' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des modes de paiement -> Alimentation des tables de correspondances');
        end;
        RecupModePaiement (Chtmp);
        SavTobPaiement;
      end
      else if Copy(Chtmp,1,3)='PL1' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des codes postaux');
        end;
        RecupCodePostaux (Chtmp);
      end
      else if Copy(Chtmp,1,3)='PL2' then     // Affectation TVA
      else if Copy(Chtmp,1,2)='PL' then      // Traduction des paramétres NON ENCORE PREVU
      else if Copy(Chtmp,1,3)='Pfs' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
          if RecodifieFamille = 'X' then	CompteRendu.lines.add('   Récupération des familles -> Alimentation des tables de correspondances')
          else CompteRendu.lines.add('   Récupération des familles ' ) ;
        end;
        RecupFamille (Chtmp);
        SavTobFamille ;
      end
      else if Copy(Chtmp,1,3)='Prf' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des regroupements de familles');
        end;
        RecupRegroupeFamille (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pgr' then
        begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des grilles de tailles');
        end;
         RecupGrille (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pco' then
      begin
       if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
       begin
        if Tob_Parametre<>nil then ImportParametre;
       	CompteRendu.lines.add('   Récupération des coloris');
       end;
       RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pmo' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des modèles');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pma' then
        begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des matières');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pfi' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des finitions');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pcg' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des catégories');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pfo' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des formules');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pcp' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des compositions');
        end;
        RecupTablesLibresArt (Chtmp);
      end
      else if Copy(Chtmp,1,2)='P1' then
      begin
         if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('      Attention : Pas de récupération des statistiques clients');
      end
      else if Copy(Chtmp,1,2)='P2' then       // définition des statistiques articles
      begin
         if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('      Attention : Pas de récupération des statistiques articles');
      end
      else if Copy(Chtmp,1,2)='P4' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('      Attention : Pas de récupération des statistiques fournisseurs');
      end
      else if Copy(Chtmp,1,3)='Psa' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des saisons/collections');
        end;
        RecupCollection (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Plg' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des langues');
        end;
        RecupLangue (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pfj' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des civilités');
        end;
        RecupCivilite (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pdm' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des motifs de démarques');
        end;
        RecupDemarque (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pes' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
          CompteRendu.lines.add('   Récupération des motifs d''entrées/sorties');
        end;
        RecupMotif_ES (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Prg' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des motifs de remises');
       	end;
        RecupRemise (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Ppo' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
	        CompteRendu.lines.add('   Récupération des unités de poids');
        end;
        RecupPoids (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pvd' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des ventes diverses');
        end;
        RecupVenteDiverse (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Poc' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des opérations de caisse');
        end;
        RecupOperationCaisse (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pcc' then
      begin
        if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
        begin
          if Tob_Parametre<>nil then ImportParametre;
        	CompteRendu.lines.add('   Récupération des catégories clients');
        end;
        RecupCategorieClient (Chtmp);
      end
      else if Copy(Chtmp,1,3)='Pbt' then       // Info complémentaires des boutiques
      else if Copy(Chtmp,1,3)='Pbb' then       // Info complémentaires des boutiques
      else if Copy(Chtmp,1,3)='Pbe' then       // Info complémentaires des boutiques
      else if Copy(Chtmp,1,3)='Pb4' then       // Paramétrage comptable
      else if Copy(Chtmp,1,3)='Pme' then       // Modèles des étiquettes
      else if Copy(Chtmp,1,3)='Pz ' then       // Données dans les modèles des étiquettes
      else if Copy(Chtmp,1,3)='Ps1' then       // Statuts de stock attendu
      else if Copy(Chtmp,1,3)='Ps2' then       // Statuts de stock disponible
      else if Copy(Chtmp,1,3)='Ps3' then       // Statuts de stock vendu
      else if Copy(Chtmp,1,3)='Pzc' then       // Formatage des zones
      else if Copy(Chtmp,1,3)='Pab' then       // Chronos
      else if Copy(Chtmp,1,3)='Pad' then       // Chronos
      else if Copy(Chtmp,1,3)='Pbc' then       // Chronos
      else if Copy(Chtmp,1,3)='Pcb' then       // Paramétrage des codes à barres
      else if Copy(Chtmp,1,3)='Pmc' then       // Mots clefs présents sur une étiquette
      else if Copy(Chtmp,1,3)='Prm' then       // Regroupement de modes de reglements
      else if Copy(Chtmp,1,3)='FOU' then       // Fournisseurs traités ailleurs
      else if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
      begin
        CompteRendu.Font.Style:=[FsBold];
        CompteRendu.lines.add('      Pas de récupération des enregistrements '+Copy(Chtmp,1,3));
        CompteRendu.Font.Style:=[];
      end;
      AncChtmp:=Chtmp;
    end;
    // MAJ de la dernière table encore chargée
    if Tob_Parametre<>nil then ImportParametre;
    // MAJ des grilles de tailles
    TOB_Tailles.SetAllModifie(TRUE) ;
    TOB_Tailles.InsertOrUpdateDB(FALSE) ;
    CloseFile(Fichier);
    FiniMove ();
  end  ;
  EcrireInfo ('Fin de la Récupération des PARAMETRES GENERAUX le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXPAR......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupParamBtq ;
var Fichier: textfile;
    Chtmp, AncChtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
begin
inherited;
  if ArretRecup then exit;
  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);
  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXpar'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des PARAMETRES de la boutique ' + FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) , True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      AncChtmp:='';
      while not EOF(Fichier) and not ArretRecup do
         begin
         MoveCur (False);
         readln(Fichier,Chtmp);
         Application.ProcessMessages ;
         if Copy(Chtmp,1,3)='Pre' then
            begin
            if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('   Récupération des représentants');
            RecupRepresentant (Chtmp);
            end
         else if Copy(Chtmp,1,3)='Pve' then
            begin
            if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('   Récupération des vendeurs');
            RecupVendeur (Chtmp);
            end
         else if Copy(Chtmp,1,3)='Pca' then
            begin
            if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then CompteRendu.lines.add('   Récupération des caissiers');
            RecupCaissier (Chtmp);
            end
         else if Copy(Chtmp,1,3)='CAP' then       // Paramétrage du CA progressif
         else if Copy(Chtmp,1,3)='PSV' then       // Paramétrage du CA progressif par vendeur
         else if Copy(Chtmp,1,3)<>Copy(AncChtmp,1,3) then
            begin
            CompteRendu.Font.Style:=[FsBold];
            CompteRendu.lines.add('      Pas de récupération de : '+Copy(Chtmp,1,3));
            CompteRendu.Font.Style:=[];
            end ;
         AncChtmp:=Chtmp;
         end;
      CloseFile(Fichier);
       EcrireInfo ('Fin de la récupération des PARAMETRES de la boutique ' + FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time) , True, True, False);
    end  ;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXART....
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupArticles ;
var Fichier: textfile;
    st : string;
    CategorieDim    : string  ;
begin
inherited;
  if ArretRecup then exit;

  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;

  EcrireInfo ('Récupération des ARTICLES' + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
  st := RepFich.text+'\EXart'+NumPremBtq.text;
  Tob_Article:=Nil ;

  if not ArretRecup and FileExists (st) then
  begin
    AssignFile(Fichier, st);
    Reset (Fichier);                                 // Ouverture du fichier
    CtrlTableTypeArt;
    // initialisations diverses
    Tob_Article        :=nil;
    Tob_Mere_Art_Taille:=nil;
    CompteurArticle    :=0  ;
    CompteurCodeBarre  :=0  ;

    while not EOF(Fichier) and not ArretRecup do
    begin
      readln(Fichier,st);
      Application.ProcessMessages ;
      if Copy (st, 1, 5) = 'ARFC1' then
      begin
         EcrireInfo ('   Article ' + Copy (st, 7, 12), False, True, False);
         RecupArticle1 (st);
      end
      else if Copy (st, 1, 5) = 'ARFC2' then RecupArticle2 (st)
      else if Copy (st, 1, 5) = 'ARFC3' then RecupArticle3 (st)
      else if Copy (st, 1, 5) = 'ARFCA' then RecupArticleAB (st)
      else if Copy (st, 1, 5) = 'ARFCB' then RecupArticleAB (st)
      else if Copy (st, 1, 5) = 'ARTC1' then
      begin
          // Import des derniers articles
          if Tob_Article<>Nil then ImportArticle;
          inc (CompteurArticle);
          inc (CompteurCodeBarre);
      	  if CompteurArticle>=1000 then
      	  begin
            CompteRendu.lines.add('      ' + IntToStr(CompteurCodeBarre) + ' codes à barres intégrés' );
            CompteurArticle:=0;
      	  end;
          RecupArtTaille (st);
      end
      else if Copy (st, 1, 2) = 'P2' then RecupTablesLibresStatArt (st)
      else CompteRendu.lines.add('      Pas de récupération de : ' + Copy(st,1,5));
    end;
    CloseFile(Fichier);
  end  ;

  if Tob_Article<>Nil then
  begin
  	ImportArticle ;
  	Tob_Article.Free ;
    Tob_Article:=nil;
  end;
  
 	if Tob_Mere_Art_Taille<>nil then ImportMereArtTaille;

  EcrireInfo ('Fin de récupération des ARTICLES' + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXCLI...
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupClients ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
begin
inherited;
  if ArretRecup then exit;
  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);
  Tob_Tiers:=Nil ;
  Tob_Tiers_Compl:=Nil;
  CtrlTableEtatCpta ;

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXcli'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
    	EcrireInfo ('Récupération des CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      // Initialisations diverses....
      Compteur      := 0;
      CompteurClient:=0;
      Tob_Tiers           :=Nil;
      Tob_Mere_Tiers      :=nil;
      Tob_Mere_Tiers_Compl:=nil;

      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;

        if Copy (Chtmp,1,5)='CLIC1' then
        begin
        	CompteurClient:=CompteurClient+1;
      		if CompteurClient>=100 then
      		begin
            // MAJ des TOB Mere
        		if Tob_Mere_Tiers<>nil then ImportMereTiers;
          	if Tob_Mere_Tiers_Compl<>nil then ImportMereTiersCompl;
        		CompteurClient:=0;
      		end;
        end;

        if Copy (Chtmp,1,3)='CLI' then RecupCLI (Chtmp)
        else if Copy (Chtmp, 1, 2) = 'P1' then RecupTablesLibresCli(Chtmp)
//      else if Copy (Chtmp, 1, 3) = 'CRG' then RecupCRG (Chtmp) non terminé
        else if Copy(Chtmp,1,3)='CLS' then
        else if Copy(Chtmp,1,3)='CLT' then
        else CompteRendu.lines.add('      Pas de récupération de : '+Copy(Chtmp,1,5));
      end;
      if Tob_Mere_Tiers<>nil then ImportMereTiers;
		  if Tob_Mere_Tiers_Compl<>nil then ImportMereTiersCompl;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end  ;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXPST...... pour les fournisseurs
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupFournisseurs ;
var Fichier: textfile;
    Chtmp : string;
begin
inherited;
  if ArretRecup then exit;
  Tob_Tiers:=Nil ;
  Chtmp:=RepFich.text+'\EXpst'+NumPremBtq.text;
  if FileExists (Chtmp) then
  begin
	  EcrireInfo ('Récupération des FOURNISSEURS' + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    AssignFile(Fichier, Chtmp);
    Reset (Fichier);                                 // Ouverture du fichier
    while not EOF(Fichier) and not ArretRecup do
    begin
      readln(Fichier,Chtmp);
      Application.ProcessMessages ;
      if Copy (Chtmp,1,3)='FOU' then RecupFOU (Chtmp)
      else if Copy (Chtmp, 1, 2) = 'P4' then RecupTablesLibresFou(Chtmp)
    end;
    CloseFile(Fichier);
    EcrireInfo ('Fin de la Récupération des FOURNISSEURS' + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
  end;
  if Tob_Tiers<>Nil then ImportTiers (False);
  if (Tob_Adresses<>Nil) then ImportAdresseTiers ;      // Adresse de facturation
  if (Tob_Tiers_Compl<>Nil) then ImportTiersCompl ;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXCDE.....
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupCommandes ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
    CategorieDim    : string  ;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////////
  // Renseignements minimums :  la dimension choisie pour les articles GB 2000
  //                         :	la nature des commandes
  //////////////////////////////////////////////////////////////////////////////
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  if Tob_Param.GetValue ('NATCDE')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des commandes fournisseurs');
    exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXCDE'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des COMMANDES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
        if Copy (Chtmp, 1, 3)='CDE' then RecupCommande (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='CDL') or (Copy (Chtmp, 1, 3)='CDT')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupCommande (Chtmp);
        end
      end  ;
      if (Tob_Piece<>Nil) then ImportPiece ;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des COMMANDES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXANN....
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupAnnonces ;
begin
  inherited;
  EcrireInfo ('ATTENTION : Pas de récupération des ANNONCES DE LIVRAISON', True, True, True);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXREC......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupReceptions ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////////
  // Renseignements minimums :  la dimension choisie pour les articles GB 2000
  //                         :	la nature des commandes
  //												 :	la nature des réceptions
  //////////////////////////////////////////////////////////////////////////////
  if Tob_Param.GetValue ('CATEGORIEDIM')='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  if Tob_Param.GetValue ('NATCDE')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des commandes fournisseurs');
    exit;
  end;
  if Tob_Param.GetValue ('NATREC')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des réceptions fournisseurs');
    exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXREC'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des RECEPTIONS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
        if Copy (Chtmp, 1, 3)='REE' then RecupReception (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='REL') or (Copy (Chtmp, 1, 3)='RET')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupReception(Chtmp);
        end ;
      end  ;

      if (Tob_Piece <> Nil) then
	    begin
  	    if ReceptSurCommande then ImportLivraison
    	  else ImportPiece;
    	end;

      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des RECEPTIONS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXFAC......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupFactures ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
    CategorieDim    : string  ;
begin
  inherited;
  if ArretRecup then exit;

  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);
  
  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXFAC'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
   		EcrireInfo ('Récupération des FACTURES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
      end  ;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des FACTURES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXTRF......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupTransferts ;
var Fichier  : textfile;
    Chtmp    : string;
    NumDeb, NumFin, NumBtq : Integer;
    CategorieDim    : string  ;
    CodeTiers : string ;
    Q : TQUERY ;
begin
  inherited;
  if ArretRecup then exit;

  //
  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  //
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
    exit;
  end;
  //
  //
  if Tob_Param.GetValue ('NATTRFEM')='' then
  begin
	ShowMessage ('Attention : vous devez renseigner la nature PGI des transferts émis');
    exit;
  end;
  //
  //
  if Tob_Param.GetValue ('NATTRFRE')='' then
  begin
	ShowMessage ('Attention : vous devez renseigner la nature PGI des transferts reçus');
    exit;
  end;
  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client du client par défaut
  ///////////////////////////////////////////////////////////////
  CodeTiers := GetParamSoc ('SO_GCTIERSDEFAUT');

  Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeTiers+ '" AND T_NATUREAUXI = "CLI" ',False) ;
  if not Q.EOF then
  begin
    Tob_Tiers_Defaut := TOB.CREATE ('TIERS', NIL, -1);
    //Tob_Tiers_Defaut.initValeurs;
    Tob_Tiers_Defaut.SelectDB('',Q);
    Ferme(Q) ;
  end else
  begin
    Ferme(Q) ;
    EcrireInfo ('   '+'Impossible de charger en TOB le client par défaut ' +  CodeTiers, True, True, True);
    exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXTRF'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo('Récupération des TRANSFERTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);
                                       // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
        if Copy (Chtmp, 1, 3)='REE' then RecupTransfert (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='REL') or (Copy (Chtmp, 1, 3)='RET')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupTransfert(Chtmp);
        end ;
      end  ;

      if (Tob_Piece <> Nil) then ImportTransfert;

      CloseFile(Fichier);
      EcrireInfo('Fin de la récupération des TRANSFERTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Fin du traitement : libération de la TOB du Client par défaut
  /////////////////////////////////////////////////////////////////
  Tob_Tiers_Defaut.free;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXESS...
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupESStocks ;
var Fichier      : textfile;
    Chtmp        : string;
    NumDeb, NumFin, NumBtq : Integer;
    CategorieDim : string  ;
    CodeTiers    : string ;
    Q            : TQUERY ;
begin
  inherited;
  if ArretRecup then exit;

  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  if Tob_Param.GetValue ('NATENT')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des entrées de stock');
    exit;
  end;
  if Tob_Param.GetValue ('NATSOR')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des sorties de stock');
    exit;
  end;

  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client du client par défaut
  ///////////////////////////////////////////////////////////////
  CodeTiers := GetParamSoc ('SO_GCTIERSMVSTK');

  Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeTiers+ '" AND T_NATUREAUXI = "CLI" ',False) ;
  if not Q.EOF then
  begin
    Tob_Tiers_Defaut := TOB.CREATE ('TIERS', NIL, -1);
    //Tob_Tiers_Defaut.initValeurs;
    Tob_Tiers_Defaut.SelectDB('',Q);
    Ferme(Q) ;
  end else
  begin
    Ferme(Q) ;
    EcrireInfo ('   '+'Impossible de charger en TOB le client par défaut ' +  CodeTiers, True, True, True);
    exit;
  end;


  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXEST'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des E/S DE STOCK de la boutique ' + FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier

      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
        if Copy (Chtmp, 1, 3)='REE' then RecupESSTock (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='REL') or (Copy (Chtmp, 1, 3)='RET')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupESSTock (Chtmp);
        end
      end  ;

      if (Tob_Piece <> Nil) then ImportESStock;

      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des E/S DE STOCK de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Fin du traitement : libération de la TOB du Client par défaut
  /////////////////////////////////////////////////////////////////
  Tob_Tiers_Defaut.free;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXOPD......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupOD ;
begin
  inherited;
  EcrireInfo ('ATTENTION : Pas de récupération des OPERATIONS DIVERSES', True, True, True);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXVTE......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupVentes ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
    CodeCliDefaut : string ;
    Q : TQUERY ;

    I: Integer ;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////////
  // Renseignements minimums :  la dimension choisie pour les articles GB 2000
  //                         :	la nature des ventes
  //////////////////////////////////////////////////////////////////////////////
  CategorieDimension := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDimension='' then
  begin
    ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
    exit;
  end;
  if Tob_Param.GetValue ('NATVTE')='' then
  begin
    ShowMessage ('Attention : vous devez renseigner la nature PGI des lignes de ventes');
    exit;
  end;

  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client du client par défaut
  ///////////////////////////////////////////////////////////////
  CodeCliDefaut := Tob_Param.GetValue ('CODECLIVTE');
  Q := OpenSQL('Select * from Tiers Where T_TIERS="' + CodeCliDefaut+ '" AND T_NATUREAUXI = "CLI" ',False) ;
  if not Q.EOF then
  begin
    Tob_Tiers_Defaut := TOB.CREATE ('TIERS', NIL, -1);
    //Tob_Tiers_Defaut.initValeurs;
    Tob_Tiers_Defaut.SelectDB('',Q);
    Ferme(Q) ;
  end else
  begin
    Ferme(Q) ;
    EcrireInfo ('   '+'Impossible de charger en TOB le client par défaut ' +  CodeCliDefaut, True, True, True);
    exit;
  end;
  ///////////////////////////////////////////////////////////////
  // Qualifiant du mouvement
  ///////////////////////////////////////////////////////////////
  QualifMvt:=GetInfoParPiece(Tob_Param.GetValue ('NATVTE'),'GPP_QUALIFMVT') ;

  ///////////////////////////////////////////////////////////////
  // Lancement du traitement
  ///////////////////////////////////////////////////////////////
  NumDeb   := StrToInt (NumPremBtq.Text);
  NumFin   := StrToInt (NumDerBtq.Text);

  I:=0 ;
  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXVTE'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des VENTES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      Compteur := 0 ;
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        if i=10 then
          begin
          Application.ProcessMessages ;
          I:=0;
          end else Inc(i) ;
        if Copy (Chtmp, 1, 3)='VTE' then RecupVente (Chtmp);
      end  ;
      if (Tob_Piece<>Nil) then ImportPiece ;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des VENTES de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Fin du traitement : libération de la TOB du Client par défaut
  /////////////////////////////////////////////////////////////////
  Tob_Tiers_Defaut.free;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXRGL....
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupReglements ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
    CategorieDim    : string  ;
begin
  inherited;
  if ArretRecup then exit;

  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
    ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
    exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXRGL'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des REGLEMENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      Compteur := 0;
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;
    	if Copy (Chtmp, 1, 3)='RGL' then RecupReglement (Chtmp);
      end  ;
      if (Tob_PiedEche<>Nil) then ImportReglement;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des REGLEMENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXSTO.....
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupStock ;
var Fichier        : textfile;
    Chtmp          : string ;
    CategorieDim   : string ;
    Qualif         : string ;
    Periode        : string ;
    NumDeb, NumFin : Integer;
    NumBtq         : Integer;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////
  // Renseignement minimum : la dimension choisie pour les articles GB 2000
  //////////////////////////////////////////////////////////////////////////
  CategorieDim := Tob_Param.GetValue ('CATEGORIEDIM');
  if CategorieDim='' then
  begin
      ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  
  /////////////////////////////////////////////////////////////////////////
  // Chargement en TOB des qualifiants TARIFS VTE
  /////////////////////////////////////////////////////////////////////////
  if Tob_Param.GetValue ('TARIF')='X' then
  begin
    Qualif := Tob_Param.GetValue ('TYPETARIFTTC');
    if Qualif = '' then
    begin
      ShowMessage ('Attention, pas de type de tarif TTC : Les prix de ventes TTC ne seront pas repris');
    end;
    Periode := Tob_Param.GetValue ('PERIODETARIFTTC');
    if Periode = '' then
    begin
      ShowMessage ('Attention, pas de période de tarif TTC : Les prix de ventes TTC ne seront pas repris');
    end;
    //
    // Chargement de la TOB TARIFMODE TTC
    //
    if (Qualif <> '') and (Periode <> '') then
    begin
      Tob_TarifMode_TTC := TOB.Create ('TARIFMODE', nil, -1);
      ChargeTobQualifiantTarif (Tob_TarifMode_TTC, Qualif, Periode);
    end;

    Qualif := Tob_Param.GetValue ('TYPETARIFSOLDE');
    if Qualif = '' then
    begin
      ShowMessage ('Attention, pas de type de tarif Soldes : Les prix de soldes ne seront pas repris');
    end;
    Periode := Tob_Param.GetValue ('PERIODETARIFSOLDE');
    if Periode = '' then
    begin
      ShowMessage ('Attention, pas de période de tarif de soldes : Les prix de soldes ne seront pas repris');
    end;
    //
    // Chargement de la TOB TARIFMODE SOLDE
    //
    if (Qualif <> '') and (Periode <> '') then
    begin
      Tob_TarifMode_Sol := TOB.Create ('TARIFMODE', nil, -1);
      ChargeTobQualifiantTarif (Tob_TarifMode_Sol, Qualif, Periode);
    end;
  end;

  /////////////////////////////////////////////////////////////////////////
  // Chargement en TOB des qualifiants TARIFS VTE
  /////////////////////////////////////////////////////////////////////////
  if Tob_Param.GetValue ('TARIFACH')='X' then
  begin
    Qualif := Tob_Param.GetValue ('TYPETARIFACH');
    if Qualif = '' then
    begin
      ShowMessage ('Attention, pas de type de tarif Achat : Les prix d''achat ne seront pas repris');
    end;
    Periode := Tob_Param.GetValue ('PERIODETARIFTTC');
    if Periode = '' then
    begin
      ShowMessage ('Attention, pas de période de tarif achat : Les prix d''achat ne seront pas repris');
    end;
    //
    // Chargement de la TOB TARIFMODE Achat
    //
    if (Qualif <> '') and (Periode <> '') then
    begin
      Tob_TarifMode_Ach := TOB.Create ('TARIFMODE', nil, -1);
      ChargeTobQualifiantTarif (Tob_TarifMode_Ach, Qualif, Periode);
    end;
  end;

  ///////////////////////////////////////////////////////
  // Récupération du dernier chrono tarif
  ///////////////////////////////////////////////////////
  if Tob_Param.GetValue ('TARIF')='X' then
  begin
    RecupMaxTarif;
  end;

  // Initialisations
  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);
  CompteurDispo:=0;
  Tob_Mere_Dispo:=nil;
  Tob_Mere_Tarif:=nil;

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXSTO'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des STOCKS/TARIFS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;

        if Copy (Chtmp, 1, 5)='STDC1' then
        begin

          CompteurDispo:=CompteurDispo+1;
      	  if CompteurDispo>=100 then
      	  begin
            // MAJ des TOB Mere
            if Tob_Mere_Dispo<>nil then ImportDispo;
            if Tob_Mere_Tarif<>nil then ImportTarif;
            CompteurDispo:=0;
          end;
        end;

        if Copy (Chtmp, 1, 3)='STD' then RecupSTD (Chtmp);
      end  ;
      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des STOCKS/TARIFS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
  if Tob_Mere_Dispo <> Nil then ImportDispo ;
  if Tob_Mere_Tarif <> Nil then ImportTarif ;

  if Tob_TarifMode_TTC <> Nil then Tob_TarifMode_TTC.free ;
  if Tob_TarifMode_Sol <> Nil then Tob_TarifMode_Sol.free ;
  if Tob_TarifMode_Ach <> Nil then Tob_TarifMode_Ach.free ;

  Tob_Mere_Dispo    := nil ;
  Tob_Mere_Tarif    := nil ;
  Tob_TarifMode_TTC := nil ;
  Tob_TarifMode_Sol := nil ;
  Tob_TarifMode_Ach := nil ;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXSTU......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupStatuts ;
begin
  inherited;
  EcrireInfo ('ATTENTION : Pas de récupération des STATUTS DE STOCKS', True, True, True);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXMTR......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupMatricules ;
begin
  inherited;
  EcrireInfo ('ATTENTION : Pas de récupération des MATRICULES', True, True, True);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXDTA......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupDetaxes ;
begin
  inherited;
  EcrireInfo ('ATTENTION : Pas de récupération des DETAXES', True, True, True);
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXNFA......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupNegFactures ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////////
  // Renseignements minimums :  la dimension choisie pour les articles GB 2000
  //                         :	la nature des commandes
  //												 :	la nature des réceptions
  //////////////////////////////////////////////////////////////////////////////
  if Tob_Param.GetValue ('CATEGORIEDIM')='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  if Tob_Param.GetValue ('NATNEGCDE')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des commandes clients');
    exit;
  end;
  if Tob_Param.GetValue ('NATNEGLIV')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des livraisons clients');
    exit;
  end;
  if Tob_Param.GetValue ('NATNEGFAC')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des factures clients');
    exit;
  end;

  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXNFA'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des FACTURES CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier

      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;

        if Copy (Chtmp, 1, 3)='DEN' then RecupNegFacture (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='DLA') or (Copy (Chtmp, 1, 3)='DTI') or (Copy (Chtmp, 1, 3)='DLP') or (Copy (Chtmp, 1, 3)='DLR')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupNegFacture(Chtmp);
        end ;
      end  ;

      if (Tob_Piece <> Nil) then
      begin
        if AvoirNegoce then ImportPiece
        else begin
          if PieceAGenerer then ImportFacCdeLiv
          else ImportPiece;
        end;
      end;

      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des FACTURES CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Intégration des fichiers EXNCD......
/////////////////////////////////////////////////////////////////////////////
procedure TFRecupGB2000.RecupNegCommandes ;
var Fichier: textfile;
    Chtmp : string;
    NumDeb, NumFin, NumBtq : Integer;
begin
  inherited;
  if ArretRecup then exit;

  //////////////////////////////////////////////////////////////////////////////
  // Renseignements minimums :  la dimension choisie pour les articles GB 2000
  //                         :	la nature des commandes
  //												 :	la nature des réceptions
  //////////////////////////////////////////////////////////////////////////////
  if Tob_Param.GetValue ('CATEGORIEDIM')='' then
  begin
    	ShowMessage ('Attention : vous devez renseigner au mnimum la dimension choisie pour les articles GB 2000');
      exit;
  end;
  if Tob_Param.GetValue ('NATNEGCDE')='' then
  begin
	  ShowMessage ('Attention : vous devez renseigner la nature PGI des commandes clients');
    exit;
  end;


  NumDeb := StrToInt (NumPremBtq.Text);
  NumFin := StrToInt (NumDerBtq.Text);

  For NumBtq:=NumDeb to NumFin do if not ArretRecup then
  begin
    Chtmp:=RepFich.text+'\EXNCD'+FormatFloat('000',NumBtq);
    if FileExists (Chtmp) then
    begin
      EcrireInfo ('Récupération des COMMANDES CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
      AssignFile(Fichier, Chtmp);
      Reset (Fichier);                                 // Ouverture du fichier
      while not EOF(Fichier) and not ArretRecup do
      begin
        readln(Fichier,Chtmp);
        Application.ProcessMessages ;

        if Copy (Chtmp, 1, 3)='DEN' then RecupNegCommande (Chtmp)
        else if ((Copy (Chtmp, 1, 3)='DLA') or (Copy (Chtmp, 1, 3)='DTI') or (Copy (Chtmp, 1, 3)='DLP') or (Copy (Chtmp, 1, 3)='DLR')) then
        begin
        	if (NombreErreurs (-1) = 0)	then RecupNegCommande(Chtmp);
        end ;

      end  ;

      if (Tob_Piece <> Nil) then ImportPiece;


      CloseFile(Fichier);
      EcrireInfo ('Fin de la récupération des COMMANDES CLIENTS de la boutique '+FormatFloat('000'+' :',NumBtq) + ' le ' + DateToStr(Date) + ' à ' + TimeToStr(Time), True, True, False);
    end;
  end;
end;

procedure TFRecupGB2000.bRecupClick(Sender: TObject);
begin
inherited;
AssignFile(FichierAno, FichAno.text);
if FileExists(FichAno.text) then Append(FichierAno) else ReWrite(FichierAno);
Writeln(FichierAno,'--- Reprise du '+DateToStr(Date)+' à '+TimeToStr(Time));

AssignFile(FichierTrt, FichTrt.text);
ReWrite(FichierTrt);
Writeln(FichierTrt,'--- Reprise du '+DateToStr(Date)+' à '+TimeToStr(Time));

NbrAno:=0;
BStop.visible := True;
Application.ProcessMessages ;
ArretRecup := False;
//TOB_Tailles:=TOB.Create('Liste des tailles',Nil,-1) ;
EcrireInfo ('Transfert d''un dossier GB 2000', True, True, False);
If DeleteTable.Checked then SupprimeTables;

//
// récupération du paramétrage de reprise des articles
//
ChargeParamArticle;

ChargeDeviseCentral;
ChargeDevisePGI;

If RepParamGen.Checked then
begin
  //  Création de la TOB des grilles de tailles
  TOB_Tailles:=TOB.Create('Liste des tailles',Nil,-1);
  RecupParametres ;
  TOB_Tailles.Free;
  TOB_Tailles := nil;
end;
If RepParamBtq.Checked then RecupParambtq ;
If RepFournisseurs.Checked then RecupFournisseurs ;
If RepArticles.Checked then
begin
  //
  //  Création de la TOB des grilles de tailles et Chargement
  //
  TOB_Tailles:=TOB.Create('Liste des tailles',Nil,-1) ;
  //
  // Chargement en TOB des grilles de tailles
  //
  ChargeTobTaille;
  //
  // Récupération Taux TVA par défaut
  //
  RecupTauxTVADefaut;
  //
  // Début de récupératrion des articles
  //
  RecupArticles ;
  //
  // Libération des TOB
  //
  TOB_Tailles.Free ;
  TOB_Tailles := nil;
end;
If RepClients.Checked then RecupClients ;
If RepReceptions.Checked then RecupReceptions ;
If RepAnnonces.Checked then RecupAnnonces ;
If RepCommandes.Checked then RecupCommandes ;
If RepFactures.Checked then RecupFactures ;
If RepTransferts.Checked then RecupTransferts ;
If RepESSTock.Checked then RecupESStocks ;
If RepOD.Checked then RecupOD ;
If RepVentes.Checked then RecupVentes ;
If RepReglements.Checked then RecupReglements ;
If RepStock.Checked then RecupStock ;
If RepStatut.Checked then RecupStatuts ;
If RepMatricules.Checked then RecupMatricules ;
If RepDetaxes.Checked then RecupDetaxes ;
If RepNegF.Checked then RecupNegFactures ;
If RepNegC.Checked then RecupNegCommandes ;

BStop.visible := False;
if Not ArretRecup then
begin
  EcrireInfo ('Récupération teminée.', True, True, False);
end;

if NbrAno=1 then CompteRendu.lines.add(#13#10'ATTENTION : 1 anomalie dans le fichier.')
else if NbrAno>1 then CompteRendu.lines.add(#13#10'ATTENTION : '+IntToStr(NbrAno)+' anomalies dans le fichier.');

//
// Libération des TOB des devises PGI
//
Tob_DevisePGI.Free ;
Tob_DevisePGI := nil;

CloseFile(FichierAno);
CloseFile(FichierTrt);
end;

Procedure Assist_ImportGB2000 ;
Var PP : THPanel ;
    X : TFRecupGB2000 ;
begin
  X:=TFRecupGB2000.Create(Application) ;

  // Création des TOB contenant le paramétrage des récupération
  Tob_Param := TOB.CREATE ('Les Paramètres', NIL, -1);
  Tob_Param.initValeurs;
  // Création de la TOB de correspondance des règlements
  Tob_Reglement := TOB.CREATE ('Les Reglements', NIL, -1);
  Tob_Reglement.initValeurs;
  // Création de la TOB de correspondance des modes de paiements
  Tob_ModePaie := TOB.CREATE ('Les Paiements', NIL, -1);
  Tob_ModePaie.initValeurs;
  // Création de la TOB de correspondance des pays
  Tob_Pays := TOB.CREATE ('Les Pays', NIL, -1);
  Tob_Pays.initValeurs;
  // Création de la TOB de correspondance des pays
  Tob_Devise := TOB.CREATE ('Les Devises', NIL, -1);
  Tob_Devise.initValeurs;
  // Création de la TOB de correspondance des familles
  Tob_Famille := TOB.CREATE ('Les Familles', NIL, -1);
  Tob_Famille.initValeurs;
  // Création de la TOB de correspondance des familles
  Tob_TVA := TOB.CREATE ('Les TVA', NIL, -1);
  Tob_TVA.initValeurs;
  // Définition des champs des TOB des paramètrage
  DefineChampTobParam(Tob_Param) ;
  // Chargement de la TOB des paramètrages à partir du fichier ASCII
  ////ChargeTobParam (Tob_Param) ;

  PP:=FindInsidePanel ;
  if (PP=Nil)or true then
  begin
    Try
      X.ShowModal;
    Finally
      X.Free;
    end;
  end else
  begin
   InitInside(X,PP) ;
   X.Show ;
  end ;

  // Libération de la TOB de paramétrage
  Tob_Param.free;
  Tob_Reglement.free;
  Tob_ModePaie.free;
  Tob_Pays.free;
  Tob_Devise.free;
  Tob_Famille.free;
  Tob_TVA.free;
end ;

//
// Initialisation des champs de la TOB de paramétrage
//
//procedure TFRecupGB2000.DefineChampTobParam ;
procedure DefineChampTobParam (TobParametrage : TOB) ;
var cpt : integer;
    Chtmp : string;
begin
  // Création des champs dans la TOB
  for cpt:=Low(ChampTobParam) to High(ChampTobParam) do
  begin
    if (ChampTobParam[cpt, 1]<>'') and not (TobParametrage.FieldExists(ChampTobParam[cpt, 1])) then
    begin
      TobParametrage.AddChampSupValeur(ChampTobParam[cpt, 1], ChampTobParam[cpt, 2]);
    end ;
  end ;

  // Pour les articles
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATA'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

  // Pour les clients
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATC'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

  // Pour les Fournisseurs
  for cpt:=1 to 20 do
  begin
    Chtmp := 'REPSTATF'+ Format('%2.2d', [cpt]) ;
    if not TobParametrage.FieldExists(Chtmp) then TobParametrage.AddChampSupValeur(Chtmp, '') ;
  end;

end;

//
// Chargement de la TOB de paramétrage
//
procedure ChargeTobParam (TobParametrage : TOB) ;
var Stg     : string;
    TOBCorr : TOB ;
    TOBL    : TOB ;
    Ind     : Integer ;
begin
  // l'ancienne structure de la TOB de paramétrage est conservée
  TOBCorr := TOB.Create('', Nil, -1) ;
  Stg := 'SELECT GRO_NOMCHAMP,GRO_OPTIONS FROM REPRISECOGPAO'
       + ' WHERE GRO_TYPEREPRISE="'+ TypeRepriseParam +'" AND GRO_VALEURGPAO="..."' ;
  TOBCorr.LoadDetailFromSQL(Stg) ;
  for Ind := 0 to TOBCorr.Detail.Count -1 do
  begin
    TOBL := TOBCorr.Detail[Ind] ;
    Stg := TOBL.GetValue('GRO_NOMCHAMP') ;
    if TobParametrage.FieldExists(Stg) then
    begin
      TobParametrage.PutValue(Stg, TOBL.GetValue('GRO_OPTIONS')) ;
    end;
  end;
  TOBCorr.Free ;
end;

//
// Enregistrement de la TOB de paramétrage
//
procedure SauveTobParam (TobParametrage : TOB) ;
var TOBCorr : TOB ;
    Ind     : Integer ;
    //////////////////////////////////////////////////////////////////////////////
    procedure AjoutTOB ( NomChamp : String ) ;
    var Stg  : string;
        TOBL : TOB ;
    begin
      if (NomChamp <> '') and (Tob_Param.FieldExists(NomChamp)) then
      begin
        TOBL := TOB.Create('REPRISECOGPAO', TOBCorr, -1) ;
        TOBL.InitValeurs ;
        TOBL.PutValue('GRO_TYPEREPRISE', TypeRepriseParam) ;
        TOBL.PutValue('GRO_NOMCHAMP', NomChamp) ;
        TOBL.PutValue('GRO_VALEURGPAO', '...') ;
        Stg := VarAsType(Tob_Param.GetValue(NomChamp), VarOleStr) ;
        if Stg = #0 then Stg := '' ;
        TOBL.PutValue('GRO_OPTIONS', stg) ;
      end ;
    end ;
    //////////////////////////////////////////////////////////////////////////////
begin
  TOBCorr := TOB.Create('', Nil, -1) ;

  // conversion de l'ancienne structure (un seul TOB)
  for Ind := Low(ChampTobParam) to High(ChampTobParam) do AjoutTOB(ChampTobParam[Ind, 1]) ;
  // Statistiques des articles
  for Ind := 1 to 20 do AjoutTOB('REPSTATA'+Format('%2.2d', [Ind])) ;
  // Statistiques des clients
  for Ind := 1 to 20 do AjoutTOB('REPSTATC'+Format('%2.2d', [Ind])) ;
  // Statistiques des Fournisseurs
  for Ind := 1 to 20 do AjoutTOB('REPSTATF'+Format('%2.2d', [Ind])) ;

  // mise à jour de la base
  Try
    BeginTrans ;
    ExecuteSQL('DELETE REPRISECOGPAO WHERE GRO_TYPEREPRISE="'+ TypeRepriseParam +'" AND GRO_VALEURGPAO="..."') ;
    TOBCorr.SetAllModifie(TRUE) ;
    TOBCorr.InsertDB(Nil) ;
    CommitTrans ;
  Except
    RollBack ;
  End ;

  TOBCorr.Free ;
end;

procedure TFRecupGB2000.RepParamGenClick(Sender: TObject);
begin
  inherited;
  if RepParamGen.Checked then ButtonParamGen.enabled:=TRUE
  else ButtonParamGen.Enabled:=FALSE;
end;

procedure TFRecupGB2000.DeleteTableClick(Sender: TObject);
begin
  inherited;
  if DeleteTable.Checked then ButtonParamDelete.enabled:=TRUE
  else ButtonParamDelete.Enabled:=FALSE;
end;

procedure TFRecupGB2000.ButtonParamGenClick(Sender: TObject);
begin
  inherited;
  // Reprise de l'ancien paramétrage
  //ChargeOldTobParam ;
  // Affichage de la fenêtre de saisie
  AppelRecupGBParamGen (Tob_Param);
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.ButtonParamFrnClick(Sender: TObject);
begin
  inherited;
  // Affichage de la fenêtre de saisie
  AppelRecupGBParamFrn (Tob_Param) ;
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.ButtonParamClientsClick(Sender: TObject);
begin
  inherited;
  // Affichage de la fenêtre de saisie
  AppelRecupGBParamCli (Tob_Param) ;
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.ButtonParamArticleClick(Sender: TObject);
begin
  inherited;
  // Affichage de la fenêtre de saisie
  AppelRecupGBParamArt (Tob_Param) ;
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.ButtonParamDeleteClick(Sender: TObject);
begin
  inherited;
  // Affichage de la fenêtre de saisie
  AppelDeleteParam (Tob_Param) ;
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.RepCommandesClick(Sender: TObject);
var check : boolean;
begin
  inherited;
  check := FALSE ;
  if RepCommandes.Checked then check := TRUE
  else if RepAnnonces.Checked then check := TRUE
  else if RepReceptions.Checked then check := TRUE
  else if RepFactures.Checked then check := TRUE
  else if RepTransferts.Checked then check := TRUE
  else if RepESSTock.Checked then check := TRUE
  else if RepOD.Checked then check := TRUE
  else if RepNegC.Checked then check := TRUE
  else if RepNegF.Checked then check := TRUE
  else if RepVentes.Checked then check := TRUE
  else if RepReglements.Checked then check := TRUE
  else if RepStock.Checked then check := TRUE
  else if RepStatut.Checked then check := TRUE
  else if RepMatricules.Checked then check := TRUE
  else if RepDetaxes.Checked then check := TRUE;

  if (check = TRUE) then ButtonParamMvt.enabled:=TRUE
  else ButtonParamMvt.Enabled:=FALSE;
end;

procedure TFRecupGB2000.ButtonParamMvtClick(Sender: TObject);
var CodeArticleVente : string ;
begin
  inherited;
  // Affichage de la fenêtre de saisie
  DateCloture      := StrToDate (Tob_Param.GetValue ('DTECLOTURE'));
  CodearticleVente := Tob_Param.GetValue ('CODEARTVTE') ;
  //gcvoirtob(Tob_Param);
  AppelRecupGBParamMvt (Tob_Param, DateCloture, CodeArticleVente) ;
  Tob_Param.PutValue ('DTECLOTURE', DateToStr(DateCloture));
  Tob_Param.PutValue ('CODEARTVTE', CodeArticleVente);
  // Sauvegarde de la TOB paramètres
  SauveTobParam (Tob_Param) ;
end;

procedure TFRecupGB2000.FichTrtElipsisClick(Sender: TObject);
begin
  inherited;
  if OpenFichTrt.execute then FichTrt.text:=OpenFichTrt.FileName;
end;

procedure TFRecupGB2000.FormShow(Sender: TObject);
begin
  inherited;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;
  RepFich.DataType := 'DIRECTORY';
  RepFich.text     := GetParamSoc ('SO_GCREPORLI');
  RepFich.Enabled  := False ;

  FichAno.text     :=GetParamSoc ('SO_GCREPORLI') + '\AnoGB2000.txt';
  FichTrt.text     :=GetParamSoc ('SO_GCREPORLI') + '\TrtGB2000.txt';
  //
  // Chargement des TOB de paramétrage
  //
  ChargeTobParam (Tob_Param);
end;

procedure TFRecupGB2000.ButtonCorrespMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var ALeft, ATop : Integer ;
begin
  inherited;
  ALeft := X ;
  ATop  := Y ;
  if Sender is TWinControl then
  begin
    Inc(ALeft, TWinControl(Sender).ClientOrigin.X) ;
    Inc(ATop, TWinControl(Sender).ClientOrigin.Y) ;
  end;
  PopupCorresp.Popup(ALeft, ATop) ;
end;

procedure TFRecupGB2000.ModeRegleClick(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, 'T_MODEREGLE') ;
end;

procedure TFRecupGB2000.Pays1Click(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, 'T_PAYS') ;
end;

procedure TFRecupGB2000.Devise1Click(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, 'T_DEVISE') ;
end;

procedure TFRecupGB2000.Famille1Click(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, '$$_CODEFAMSF') ;
end;

procedure TFRecupGB2000.ModePaieClick(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, 'MP_MODEPAIE') ;
end;

procedure TFRecupGB2000.TVA1Click(Sender: TObject);
begin
  inherited;
  AppelSaisieCorrespGPAO(taModif, TypeRepriseGB2000, '$$_CODETVA') ;
end;

procedure TFRecupGB2000.FormDestroy(Sender: TObject);
begin
  inherited ; 
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

end.

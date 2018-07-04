{***********UNITE*************************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Outil générique d'import de données
Mots clefs ... : IMPORT;GPAO
*****************************************************************}                                   
unit AssistImportGPAO;

interface
    
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFDEF EAGLCLIENT}
  UtileAGL,
  {$ELSE}
  MajTable,
  {$ENDIF}
  assist, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ComCtrls, ExtCtrls,
  HPanel, UIutil, Spin, Hent1, ent1, HRichEdt, HRichOLE, UTob, Grids, DBTables,
  ImgList, HStatus, Mask, UTOM, SaisUtil, HDebug , UtilPGI, ParamSoc, UTomTiers,
  FactUtil, FactCalc, Facture, FactComm, EntGC, VoirTob, FileCtrl, formule, M3VM,
  TarifUtil, GcInfoFichier, SaisieCorrespGPAO, ArtPhoto_tof, ToxIntegre, UtilArticle,
  StockUtil, UFidelite, FactTob, FactPiece, FactAdresse, FactArticle,
  Buttons, UtilGPAO, AssistTransformeGPAO;

type
  TTImportGPAO = class(TFAssist)
    Lancement: TTabSheet;
    TabSheet1: TTabSheet;
    HLabel8: THLabel;
    FichTrt: THCritMaskEdit;
    Panel1: TPanel;
    TITREASSISTANT: THLabel;
    CompteRendu: TMemo;
    bRecup: TToolbarButton97;
    BStop: TToolbarButton97;
    Label1: TLabel;
    Label3: TLabel;
    FichAno: THCritMaskEdit;
    OpenFichAno: TOpenDialog;
    Label2: TLabel;
    FichCpt: THCritMaskEdit;
    OpenFichCpt: TOpenDialog;
    Provenance: THValComboBox;
    HLabel2: THLabel;
    OpenDialog2: TOpenDialog;
    OpenFichGPAO: TOpenDialog;
    BoutonDetope: TButton;
    LU: TLabel;
    Traitees: TLabel;
    Erreur: TLabel;
    NbreLu: TEdit;
    NbreTraite: TEdit;
    NbreRejet: TEdit;
    IntegAuto: TCheckBox;
    ListeFichier: TFileListBox;
    Timer1: TTimer;
    IntegBoucle: TCheckBox;
    TempsBoucle: TSpinEdit;
    Label4: TLabel;
    NbreTotal: TEdit;
    Bevel1: TBevel;
    Bevel2: TBevel;
    TransfoType: THValComboBox;
    LabelTransfo: TLabel;
    procedure bPrecedentClick(Sender: TObject);
    procedure bSuivantClick(Sender: TObject);
    procedure bRecupClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure FichAnoElipsisClick(Sender: TObject);
    procedure FichCptElipsisClick(Sender: TObject);
    procedure FichTrtElipsisClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BoutonDetopeClick(Sender: TObject);
    procedure IntegAutoClick(Sender: TObject);
    procedure ListeFichierDblClick(Sender: TObject);
    procedure TempsBoucleChange(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure ProvenanceChange(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private   { Déclarations privées }

    Inventaire   : boolean ;// Provenance de la donnée

    Tob_InfoGPAO    : TOB ; // Tob mère contenant les info à intégrer dans la base
    Tob_RepriseGPAO : TOB ; // Tob contenant le descriptif du fichier à reprendre
    Tob_CorrespGPAO : TOB ; // Tob contenant les tables de correspondance

    // Gestion du fichier d'intégration
    Fichier           : file of char;    // Pointeur sur le fichier d'intégration
    FichierTexte      : textfile   ;     // Pointeur sur le fichier d'intégration sans entête - Fichier texte
    NomTable          : string ;         // Nom de la table en cours de traitement
    PosTope           : Longint ;        // Position début de lecture pour topage ultérieur
    PosTopePiece      : Longint ;        // Position début de pièce pour topage ultérieur
    PosLecture        : Longint ;        // Position courante de lecture

    // Gestion des compteurs d'enregistrements
    CptEnregTot       : integer ;        // Compteur pour enregistrements lus (à traiter ou pas)
    CptEnreg          : integer ;        // Compteur pour enregistrements lus à traiter
    CptMAJ            : integer ;        // Compteur pour MAJ par paquet
    CptTrt            : integer ;        // Compteur des enregistrements traités
    CptErr            : integer ;        // Compteur des enregistrement en erreurs
    CptIntegre        : integer ;        // Compteur des enregistrement en erreurs
    ArretRecup        : Boolean ;        // Arret de l'import
    FichierAno        : textfile;        // Pointeur sur le fichier des anomalies
    FichierTrt        : textfile;        // Pointeur sur le fichier d'entrée
    NbrAno            : Integer ;

    // Reprise des documents
    Tob_Tiers         : TOB ;            // TOB du tiers
    Tob_Piece         : TOB ;            // TOB entête de pièce
    Tob_Ligne         : TOB ;            // TOB lignes
    Tob_Article       : TOB ;            // TOB des articles de la pièce
    Tob_PiedBase      : TOB ;            // TOB PiedBase
    Tob_PiedEche      : TOB ;            // TOB PiedEche
    Tob_Adresses      : TOB ;            // TOB adresses
    Tob_Acomptes      : TOB ;            // TOB Acomptes
    Tob_Catalogue     : TOB ;            // TOB Catalogues
    Tob_Noment        : TOB ;            // TOB Nomenclatures

    Tob_Piece_Prec_Old    : TOB ;            // TOB entête de pièce précédente à solder
    Tob_Article_Prec_Old  : TOB ;            // TOB des articles de la pièce précédente d'indice supérieur
    Tob_PiedBase_Prec_Old : TOB ;            // TOB PiedBase précédente à solder
    Tob_PiedEche_Prec_Old : TOB ;            // TOB PiedEche précédente à solder
    Tob_Adresses_Prec_Old : TOB ;            // TOB Adresses précédente à solder
    Tob_Article_Prec_New  : TOB ;            // TOB des articles de la pièce précédente d'indice supérieur
    EncoreReliquat        : boolean ;        // Y-a-t-il encore du reliquat ?
    
    CleDoc            : R_CleDoc ;       // Structure document
    DEV               : RDEVISE ;        // Structure devise
    QualifMvt	      : string ;         // qualifiant mouvement de la nature de pièce
    VenteAchat	      : string ;         // qualifiant mouvement ACH/VTE/TRF
    Numligne          : integer;         // Numéro de la dernière ligne créée
    QuantiteVendueGB  : double ;         // Qté de la vente en cours GB (Valeur Absolue)
    QuantiteVenduePGI : double ;         // Qté de la vente en cours PGI (Valeur signée)
    OldRefArticle     : string ;         // Code article récupéré
    Designation       : string ;         // Designation
    NatureTrfRecu     : string ;         // Nature du transfert dupliqué
    MajStockPiece     : boolean;         // MAJ du stock ?
    CtrlPrixNul       : boolean;         // MAJ du stock ?
    GenereEcheances   : boolean;         // Génération des échéances ?
    RegroupeLignes    : boolean;         // Regroupement des lignes avec même article ?
    DernierClientCree : string ;         // Code du dernier client créé (cas du 999999 dans CASH 2000)
    ModePaiement      : string ;         // Mode de paiement MAGESTEL
    OldNaturePiece    : string ;         // Code article récupéré
    OldRefInterne     : string ;         // Code article récupéré
    DateLivraison     : string ;         // Date de livraison

    ValTiersDefaut    : string ;         // Valeur du tiers par défaut
    Tob_TiersDefaut   : Tob ;            // Tob du tiers par défaut

    // Gestion de la ligne commentaire dans document
    LigneCommACreer   : boolean;
    NumLigneComment   : integer;	 // Numéro de la ligne commentaire
    TotQte	      : double ;         // Total quantité
    TotPrix	      : double ;         // Total Prix
    TotPrixNet        : double ;         // Total prix net
    RefInterneLigneCom: string ;
    NatureLigneCom    : string ;

    // Gestion de la reprise des tarifs
    Tob_Tarif         : TOB    ;         // Tob contenant le descriptif de TARIFTYPMODE
    Periode           : string ;         // Période de tarif pour PV
    CodeTarif         : string ;         // Code tarif pour PV
    MaxTarif          : integer;         // Numéro du tarif à créer
    DernierArticleTar : string ;

    // Gestion de la reprise des inventaires
    CodeTrans         : string ;         // Code transmission
    NumLigneInv       : integer;         // Ligne Inventaire
    EnteteInvCree     : boolean;

    // Optimisation de l'intégration des tarifs
    DernierDepot               : string ;
    DerniereDevise             : string ;
    DernierArticlePA           : string ;
    DernierArticlePV           : string ;
    DernierArticleFlagPxUnique : string ;

    // Optimisation de l'intégration des articles
    DernierFournisseur     : string ;
    DernierTiers           : string ;
    DerniereCollection     : string ;
    DerniereFamille        : array [1..8]  of string ;
    DerniereStatArt        : array [1..10] of string ;
    DerniereTaxe           : array [1..2]  of string ;
    DerniereGrille         : array [1..5]  of string ;
    DerniereStartCompl     : array [1..2]  of string ;
    DernierTypeTarif       : string ;
    DernierStatut          : string ;
    DernierTypeArticle     : string ;
    DernierArticleGen      : string ;
    DernierArticleGenCompl : string ;

    // Optimisation de l'intégration des tiers
    DernierPays            : string ;
    DerniereLangue         : string ;
    DerniereZoneCom        : string ;
    DernierTarifTiers      : string ;
    DernierModeRegle       : string ;
    DernierRegimeTVA       : string ;
    DernierJuridique       : string ;
    DerniereNationalite    : string ;
    DernierComptaTiers     : string ;

    // Optimisation de l'intégration des pieces
    DebutFichier           : boolean;
    IntegrationPiece       : boolean;
    PieceASolder           : boolean;
    DerniereNature         : string ;
    DernierEtab            : string ;
    DernierRefInterne      : string ;
    LibellePiece           : string ;
    IdentArticleImport     : string ;
    TiersPiecePrecedente   : string ;
    ReprePiecePrecedente   : string ;
    DerniereNaturePrec     : string ;
    DerniereReferencePrec  : string ;

    // Optimisation de l'intégration de la fidélité
    CodeProgramme          : string ;
    DernierCodeProgramme   : string ;
    TypeCumulFid           : string ;
    //CptFid                 : integer;
    // Déconnexion/reconnexion à la base
    NomDossier             : string ;
    MultiUserLogin         : boolean ;
    ConnexionBase          : boolean ;
    DateConnexion          : TDateTime;
    DateDeconnexion        : TDateTime;

    // Procédure principale
    procedure RecupInfoGPAO     (Nomfichier : string) ;
    procedure RecupInfoGPAOFixe (Nomfichier : string) ;
    procedure RecupInfoGPAOAuto (Entete     : boolean);
    procedure InitImportGPAO ;
    procedure FinImportGPAO  ;

    // Gestion des comptes rendu
    procedure EcrireFichAno (Chtmp : String) ;
    procedure EcrireFichCpt (Chtmp : String) ;
    procedure EcrireInfo    (Chtmp : String; Cpte, Trt, Ano : boolean);
    procedure AfficheCompteur      ;
    procedure AfficheFinTraitement ;

    // Gestion du fichier d'intégration
    function  LectureEnregistrement : string ;
    procedure TopageEnregistrement (CharTop : char) ;
    procedure TopagePiece (CharTop : char);

    // Traitement des formules
    function TraiteFormule (Valeur, ValDefaut : string; TOBFilleRecup, TobFilleParam : TOB) : String;

    // Transfert fichier -> TOB
    procedure FileToTob(Info: string; TOBFilleRecup: TOB; Entete : boolean);

    // Gestion des paramètres, articles,.....
    function  IntegreInfoGPAO (var Info : string; Entete : boolean) : boolean;
    procedure ImportTobInfoGPAO;

    // Recherche des articles PGI
    function  ReconstitueIdentArticle (TobInfo : TOB) : string ;

    // Gestion des documents
    function  InitEntetePiece (TobInfo : TOB) : boolean;
    function  InitLignePiece (TobInfo : TOB) : boolean;
    function  NumeroteLignes : integer ;
    procedure DupliqueTrfRecu ;
    procedure CreeLigneCommentaire ;
    procedure IntegrePiece (var Info : string) ;
    procedure ImportPiece ;
    procedure ControlePrixNul ;
    procedure TraiteLesreliquats (Tob_Piece, Tob_Article : TOB) ;
    procedure ChargeLesArticles  (TobPiece, TobArticle : TOB);
//    procedure ChargeLesArticles2 (TobPiece, TobArticle : TOB);
    procedure ChargeLesArticles3 (TobPiece, TobArticle : TOB);
    procedure ChargeLaPiecePrecedente (NaturePiece, Refinterne : string) ;
    procedure MetAJourQteReste (RefInterne : string);
    function  AnnuleStockPieceASolder ( TobPieceSolde : TOB ; TobArticle : TOB) : integer ;
    function  MajStockPieceASolder ( TobPieceSolde : TOB ; TobArticle : TOB) : integer ;
    procedure RecalculLigneGenerique (TobPieceSolde : TOB) ;

    // Contrôle des informations
    function  ControleInfoOK     (NomTable : string; var TobInfo : TOB) : boolean;
    function  CtrlTOBChoixCod    (var TobInfo : TOB) : boolean;
    function  CtrlTOBChoixExt    (var TobInfo : TOB) : boolean;
    function  CtrlTOBDimension   (var TobInfo : TOB) : boolean;
    function  CtrlTOBArticle     (var TobInfo : TOB) : boolean;
    function  CtrlTOBDispo       (var TobInfo : TOB) : boolean;
    function  CtrlTOBTarif       (var TobInfo : TOB) : boolean;
    function  CtrlTOBPiece       (var TobInfo : TOB) : boolean;
    function  CtrlTOBTiers       (var TobInfo : TOB) : boolean;
    function  CtrlTOBPiedEche    (var TobInfo : TOB) : boolean;
    function  CtrlTOBAcomptes    (var TobInfo : TOB) : boolean;
    function  CtrlTOBInventaire  (var TobInfo : TOB) : boolean;
    function  CtrlTOBFidelite    (var TobInfo : TOB) : boolean;
    function  CtrlTOBProspects   (var TobInfo : TOB) : boolean;

    // MAJ spécifique des TOB avant import
    procedure MAJSpecifInfo               (NomTable : string; var TobInfo : TOB) ;
    procedure MAJSpecifDimension          (var TobInfo : TOB) ;
    procedure MAJSpecifArticle            (var TobInfo : TOB) ;
    procedure AjouteChampSpecifArticle    (var Tob_Art : TOB) ;
    procedure GestionDimensionArticle     (var Tob_Art : TOB) ;
    procedure ArticleCodeDim              (var Tob_Art : TOB) ;
    procedure ConvPrixMajPartielleArticle (var Tob_Art : TOB) ;
    procedure MAJSpecifDevise             (var TobInfo : TOB) ;
    procedure MAJSpecifTarifTypMode       (var TobInfo : TOB) ;
    procedure MAJSpecifTarifPer           (var TobInfo : TOB) ;
    procedure MAJSpecifDispo              (var TobInfo : TOB) ;
    procedure MAJSpecifTarif              (var TobInfo : TOB) ;
    procedure MAJSpecifTiers              (var TobInfo : TOB) ;
    procedure MAJSpecifInventaire         (var TobInfo : TOB) ;
    procedure MAJSpecifFidelite           (var TobInfo : TOB) ;

    procedure IntegreGPAO (var Info : string) ;
    procedure IntegreGPAOFixe (var Info : string) ;
    function  ChargeParamReprise (var Info : string) : boolean ;
    function  ChargeParamRepriseFixe   : boolean ;
    function  ReprisefichierEnteteFixe : boolean ;
    function  ChargeTobQualifiantTarif(var Tob_Tarif : TOB; TypeTarif : string; TypePeriode : string) : boolean ;
    procedure RecupMaxTarif ;
    procedure MajPrixAchatArticle     (CodeArticle : string ; PrixArticle : double; Devise : string) ;
    procedure MajPrixVenteArticle     (CodeArticle : string ; PrixArticle : double; Devise : string) ;
    procedure MajPrixVenteToutArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;

    procedure AfficheRepORLI ;
    procedure IntegrationGPAO;

    // Déconnexion/Reconnexion à la base
    procedure DeConnexion;
    procedure Connexion;
    procedure InitDateConnect ;

  public    { Déclarations publiques }

  end;

//  objet pour le traitement des formules
Type TGPAOFormule = class(TObject)
  private
    Tob_Formule   : TOB ;
    Valeur        : String ;
    ValDefaut     : String ;

  public
    Constructor Create (TOBChamp : TOB ; ValChamp, ValDefChamp : String) ; Overload ;
    Destructor  Destroy ; Override ;
    function    PrendChamp (Champ : String) : Variant ;

  end;

Procedure Assist_ImportGPAO (Inventaire : boolean = False) ;

implementation

{$R *.DFM}

///////////////////////////////////////////////////////////////////////////////////////
//  Create : création de l'objet de traitement des formules
///////////////////////////////////////////////////////////////////////////////////////
constructor TGPAOFormule.Create (TOBChamp : TOB ; ValChamp, ValDefChamp : String) ;
begin
  Inherited Create ;
  Tob_Formule := TOBChamp;
  Valeur := ValChamp;
  ValDefaut := ValDefChamp;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Destroy : destruction de l'objet de traitement des formules
///////////////////////////////////////////////////////////////////////////////////////
destructor TGPAOFormule.Destroy ;
begin
  Tob_Formule := Nil;
  Inherited Destroy ;
end;

/////////////////////////////////////////////////////////////////////////////
// Fonction de recherche de la valeur d'un champ pour une formule
/////////////////////////////////////////////////////////////////////////////
function TGPAOFormule.PrendChamp ( Champ : String ) : Variant ;
begin
   Champ := Uppercase(Trim(Champ));
        if Champ = 'VALEUR'    then Result := Valeur
   else if Champ = 'VALDEFAUT' then Result := ValDefaut
   else if Copy(Champ, 1, 3) = 'SO_' then Result := GetParamSoc(Champ)
   else if (Tob_Formule <> Nil) and (Tob_Formule.FieldExists(Champ)) then
     Result := VarAsType(Tob_Formule.GetValue (Champ), varString)
   else Result := #0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Module d'import de données
Mots clefs ... :
*****************************************************************}
Procedure Assist_ImportGPAO (Inventaire : boolean = False) ;
Var PP : THPanel ;
    X : TTImportGPAO ;
begin
  X:=TTImportGPAO.Create(Application) ;

  // Initialisation de la provenance de la donnée
  X.Inventaire := Inventaire;
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
end ;

{***********A.G.L.***********************************************
Auteur  ...... : L. Meunier
Créé le ...... : 27/11/2003
Modifié le ... :   /  /    
Description .. : Affiche le contenu du répertoire par défaut
Mots clefs ... : 
*****************************************************************}
procedure TTImportGPAO.AfficheRepORLI  ;
Var Repertoire: String ;
begin
  Repertoire := FichTrt.Text ;
  if (Repertoire <> '') and (DirectoryExists (Repertoire)) then
  begin
    ListeFichier.Directory  := Repertoire + '\';
    ListeFichier.Mask := 'AE*.*;CE*.*;FE*.*;DT*.*' ;
    ListeFichier.ShowGlyphs := True ;
    ListeFichier.Update ;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Génère la référence de l'article dimensionné à partir de la
Suite ........ : référence générique et des code dimension
Mots clefs ... : REFERENCE ARTICLE DIMENSIONNÉ
*****************************************************************}
procedure TTImportGPAO.ArticleCodeDim (var Tob_Art : TOB) ;
Var Chtmp     : String ;
    StBlanc   : string ;
    Article   : string ;
    CodeDim1  : string ;
    CodeDim2  : string ;
    CodeDim3  : string ;
    CodeDim4  : string ;
    CodeDim5  : string ;
    longueur  : integer;
begin

  if (Tob_Art.GetValue ('GA_CODEARTICLE') <> '') and (Tob_Art.GetValue ('GA_ARTICLE') = '') then
  begin
    // Etape 1 : Formatage de la référence article dimensionné PGI
    Article := Tob_art.GetValue ('GA_CODEARTICLE');
    Article := Trim   (Article);
    longueur:= length (Article);
    StBlanc := '                  ';

    if Tob_Art.GetValue ('GA_CODEDIM1') <> '' then CodeDim1 := Tob_Art.GetValue ('GA_CODEDIM1') else CodeDim1 := '   ';
    if Tob_Art.GetValue ('GA_CODEDIM2') <> '' then CodeDim2 := Tob_Art.GetValue ('GA_CODEDIM2') else CodeDim2 := '   ';
    if Tob_Art.GetValue ('GA_CODEDIM3') <> '' then CodeDim3 := Tob_Art.GetValue ('GA_CODEDIM3') else CodeDim3 := '   ';
    if Tob_Art.GetValue ('GA_CODEDIM4') <> '' then CodeDim4 := Tob_Art.GetValue ('GA_CODEDIM4') else CodeDim4 := '   ';
    if Tob_Art.GetValue ('GA_CODEDIM5') <> '' then CodeDim5 := Tob_Art.GetValue ('GA_CODEDIM5') else CodeDim5 := '   ';

    Chtmp := Article + Copy(StBlanc, 1, 18-longueur) + CodeDim1 + CodeDim2+ CodeDim3 + CodeDim4 + CodeDim5 + 'X';
    Tob_Art.PutValue('GA_ARTICLE',Chtmp);
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Détermine la référence de l'article générique
Mots clefs ... : REFERENCE ARTICLE GENERIQUE
*****************************************************************}
procedure ArticleCodeGen (var Article : string) ;
Var StBlanc   : string ;
    longueur  : integer;
begin
  Article  := Trim   (Article);
  longueur := length (Article);
  StBlanc  := '                  ';
  Article  := Article + Copy(StBlanc, 1, 18-longueur) + '               ' + 'X' ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Création d'un article générique à partir d'un article 
Suite ........ : dimensionné
Mots clefs ... : CREATION ARTICLE GENERIQUE
*****************************************************************}
procedure CreateArticleGenerique (TobArtDim : TOB);
var
  TobArtGen      : TOB    ;
  CodeArticleGen : string ;
begin
  // Création et dupliqcation de la TOB contenant l'article dimensionné
  TobArtGen := Tob.Create ('ARTICLE', nil, -1);
  TobArtGen.dupliquer (TobArtDim, True, True, True);

  // MAJ spécifique pour l'article générique
  CodeArticleGen := TobArtGen.GetValue ('GA_CODEARTICLE');
  ArticleCodeGen (CodeArticleGen);
  TobArtGen.PutValue ('GA_ARTICLE'   , CodeArticleGen);
  TobArtGen.PutValue ('GA_STATUTART' , 'GEN');
  TobArtGen.PutValue ('GA_CODEBARRE' , '');
  TobArtGen.PutValue ('GA_CODEDIM1'  , '');
  TobArtGen.PutValue ('GA_CODEDIM2'  , '');
  TobArtGen.PutValue ('GA_CODEDIM3'  , '');
  TobArtGen.PutValue ('GA_CODEDIM4'  , '');
  TobArtGen.PutValue ('GA_CODEDIM5'  , '');
  TobArtGen.PutValue ('GA_GRILLEDIM1', '');
  TobArtGen.PutValue ('GA_GRILLEDIM2', '');
  TobArtGen.PutValue ('GA_GRILLEDIM3', '');
  TobArtGen.PutValue ('GA_GRILLEDIM4', '');
  TobArtGen.PutValue ('GA_GRILLEDIM5', '');
  // Insertion dans la base
  TobArtGen.InsertDB (nil);
  FreeAndNil( TobArtGen );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 27/09/2001
Modifié le ... :   /  /
Description .. : Envoi le PAHT d'un article
Mots clefs ... :
*****************************************************************}
function RecherchePAHT (Article : string) : double ;
var Q    : TQuery ;
    Prix : double ;
begin
  Prix := 0.0 ;
  Q := OpenSQL ('SELECT GA_PAHT FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"', True) ;
  if not Q.EOF then Prix := Q.FindField('GA_PAHT').AsFloat ;
  Ferme (Q);
  result := Prix ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 27/09/2001
Modifié le ... :   /  /
Description .. : Envoi le PVHT d'un article
Mots clefs ... :
*****************************************************************}
function RecherchePVHT (Article : string) : double ;
var Q    : TQuery ;
    Prix : double ;
begin
  Prix := 0.0 ;
  Q := OpenSQL ('SELECT GA_PVHT FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"', True) ;
  if not Q.EOF then Prix := Q.FindField('GA_PVHT').AsFloat ;
  Ferme (Q);
  result := Prix ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 27/09/2001
Modifié le ... :   /  /
Description .. : Envoi le PATTC d'un article
Mots clefs ... :
*****************************************************************}
function RecherchePVTTC (Article : string) : double ;
var Q    : TQuery ;
    Prix : double ;
begin
  Prix := 0.0 ;
  Q := OpenSQL ('SELECT GA_PVTTC FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"', True) ;
  if not Q.EOF then Prix := Q.FindField('GA_PVTTC').AsFloat ;
  Ferme (Q);
  result := Prix ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 27/09/2001
Modifié le ... :   /  /
Description .. : Mies à jour du flag prix unique
Mots clefs ... :
*****************************************************************}
procedure MiseAjourFlagPrixUnique (Article : string) ;
Var CodeArticle : string ;
begin
  CodeArticle := Trim (Copy(Article, 1, 18));
  ExecuteSQL ('UPDATE ARTICLE SET GA_PRIXUNIQUE="-" WHERE GA_STATUTART IN ("DIM","GEN","UNI") AND GA_CODEARTICLE="'+CodeArticle+'" AND GA_PRIXUNIQUE="X"');
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 17/10/2001
Modifié le ... :   /  /    
Description .. : Maj des PA des la fiche article
Mots clefs ... : 
*****************************************************************}
procedure TTImportGPAO.MajPrixAchatArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeArticleGen : string    ;
    requete        : string    ;
    Montant        : double    ;
begin
  Montant        := PrixArticle ;

  // Dernier prix d'achat
  if Montant <> 0.0 then
  begin

    // Si devise différente de la devise dossier -> Conversion
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;

    // MAJ de l'article dimensionné
    ExecuteSQL ('UPDATE ARTICLE SET GA_DPA="'+StrFPoint(Montant)+'" , GA_PAHT="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');

    // L'article générique a t il un prix d'achat ?
    CodeArticleGen := Copy(CodeArticle, 1, 18) + '               X';
    if CodeArticleGen <> DernierArticlePA then
    begin
      Requete := 'UPDATE ARTICLE SET GA_DPA="'+StrFPoint(Montant)+'", GA_PAHT="'+StrFPoint(Montant)+'"' ;
      Requete := Requete + ' WHERE GA_ARTICLE="'+CodeArticleGen+'"';
      ExecuteSQL (Requete);

      DernierArticlePA := CodeArticleGen ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 17/10/2001
Modifié le ... : 17/10/2001
Description .. : Maj des PV des la fiche article
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MajPrixVenteArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    Montant        : double    ;
begin
  Montant := PrixArticle ;

  // Dernier prix d'achat
  if Montant <> 0.0 then
  begin
    // Si devise différente de la devise dossier -> Conversion
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;

    // MAJ de l'article dimensionné
    ExecuteSQL ('UPDATE ARTICLE SET GA_PVTTC="'+StrFPoint(Montant)+'" WHERE GA_ARTICLE="'+CodeArticle+'"');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 17/10/2001
Modifié le ... : 09/11/2001
Description .. : Maj des PV des toutes les  fiches articles
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MajPrixVenteToutArticle (CodeArticle : string ; PrixArticle : double; Devise : string) ;
var Date           : TDateTime ;
    DEV            : RDEVISE   ;
    CodeToutArticle: string    ;
    Montant        : double    ;
begin
  Montant := PrixArticle ;

  // Dernier prix d'achat
  if Montant <> 0.0 then
  begin
    // Si devise différente de la devise dossier -> Conversion
    if Devise <> V_PGI.DevisePivot then
    begin
      Date         := NowH ;
      DEV.Code     := Devise ;
      DEV.DateTaux := NowH ;
      GetInfosDevise(DEV) ;
      DEV.Taux := GetTaux (DEV.Code, DEV.DateTaux, Date) ;
      Montant  := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
    end;

    // Récup du code article
    CodeToutArticle := Trim (Copy(CodeArticle, 1, 18));

    // MAJ des fiches articles
    ExecuteSQL ('UPDATE ARTICLE SET GA_PVTTC="'+StrFPoint(Montant)+'" WHERE GA_STATUTART IN ("DIM","GEN","UNI") AND GA_CODEARTICLE="'+CodeToutArticle+'"');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Détermine le prochain code interne de stockage de la
Suite ........ : dimension
Mots clefs ... : CODE INTERNE
*****************************************************************}
procedure CodeCodeInterne (var Code: String);
Var i : integer ;
    CL : String ;
BEGIN
  CL:=uppercase(Trim(Code)) ;
  if Length(CL)<3 then Code:='001'
  else begin
    i:=3 ; While CL[i]='Z' do BEGIN CL[i]:='0' ; Dec(i) ; END ;
    if Ord(CL[i])=57 then CL[i]:='A' else CL[i]:=Succ(CL[i]) ;
    Code:=CL ;
  end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Création d'un nouveau masque de dimension
Mots clefs ... : MASQUE DIMENSION
*****************************************************************}
function CreateNewMasque (GrilleDim1, GrilleDim2, GrilleDim3, GrilleDim4, GrilleDim5 : string) : string;
Var SQL           : string  ;
    Q             : TQUERY  ;
    TobMasque     : TOB     ;
    CodeMasqueMax : String  ;
    LibMasque     : string  ;
    TypeMasque    : string  ;
    MultiEtab     : string  ;
    ExisteMasque  : boolean ;
    LI1, LI2      : boolean ;
    CO1, CO2      : boolean ;
    ON1           : boolean ;
begin
  LI1 := False ;
  LI2 := False ;
  CO1 := False ;
  CO2 := False ;
  ON1 := False ;

  // Recherche du dernier masque créé
  CodeMasqueMax := '001';
  ExisteMasque  := True ;
  while ExisteMasque do
  begin
    Q := OpenSQL ('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="'+CodeMasqueMax+'"', True) ;
    if not Q.EOF then CodeCodeInterne (CodeMasqueMax)
    else ExisteMasque := False ;
    Ferme (Q);
  end;

  // Création du nouveau masque
  SQL:='Select GMQ_TYPEMASQUE, GMQ_MULTIETAB From TYPEMASQUE';
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then
  begin
    while not Q.EOF do
    begin
       TypeMasque := Q.FindField('GMQ_TYPEMASQUE').AsString ;
       MultiEtab  := Q.FindField('GMQ_MULTIETAB').AsString  ;

       // Création d'un masque
       TobMasque := Tob.create ('DIMMASQUE', nil, -1);
       TobMasque.PutValue ('GDM_MASQUE'   , CodeMasqueMax);

       LibMasque := 'Grilles ';
       if GrilleDim1 <> '' then LibMasque := LibMasque + GrilleDim1;
       if GrilleDim2 <> '' then LibMasque := LibMasque + '-' + GrilleDim2;
       if GrilleDim3 <> '' then LibMasque := LibMasque + '-'  + GrilleDim3;
       if GrilleDim4 <> '' then LibMasque := LibMasque + '-'  + GrilleDim4;
       if GrilleDim5 <> '' then LibMasque := LibMasque + '-'  + GrilleDim5;

//       TobMasque.PutValue ('GDM_LIBELLE'  , 'Masque GPAO ' + CodeMasqueMax);
       TobMasque.PutValue ('GDM_LIBELLE'  , LibMasque);
       if GrilleDim1 <> '' then
       begin
         TobMasque.PutValue ('GDM_TYPE1'    , GrilleDim1);
         TobMasque.PutValue ('GDM_POSITION1', 'CO1');
         CO1 := True ;
       end;
       if GrilleDim2 <> '' then
       begin
         TobMasque.PutValue ('GDM_TYPE2'    , GrilleDim2);
         TobMasque.PutValue ('GDM_POSITION2', 'LI1');
         LI1 := True ;
       end;
       if GrilleDim3 <> '' then
       begin
         TobMasque.PutValue ('GDM_TYPE3'    , GrilleDim3);
         if CO1 then
         begin
           TobMasque.PutValue ('GDM_POSITION3', 'CO2');
           CO2 := True ;
         end else
         begin
           TobMasque.PutValue ('GDM_POSITION3', 'CO1');
           CO1 := True ;
         end;
       end;
       if GrilleDim4 <> '' then
       begin
         TobMasque.PutValue ('GDM_TYPE4'    , GrilleDim4);
         if LI1 then
         begin
           TobMasque.PutValue ('GDM_POSITION4', 'LI2') ;
           LI2 := True ;
         end else
         begin
           TobMasque.PutValue ('GDM_POSITION4', 'LI1');
           LI1 := True ;
         end;
       end;
       if GrilleDim5 <> '' then
       begin
         TobMasque.PutValue ('GDM_TYPE5'    , GrilleDim5);
         TobMasque.PutValue ('GDM_POSITION5', 'ON1');
         ON1 := True;
       end;

       // Type du masque
       TobMasque.PutValue ('GDM_TYPEMASQUE', TypeMasque);

       // 6ième position
       if MultiEtab = 'X' then
       begin
              if not LI1 then TobMasque.PutValue ('GDM_POSITION6', 'LI1')
         else if not LI2 then TobMasque.PutValue ('GDM_POSITION6', 'LI2')
         else if not CO1 then TobMasque.PutValue ('GDM_POSITION6', 'CO1')
         else if not CO2 then TobMasque.PutValue ('GDM_POSITION6', 'CO2')
         else if not ON1 then TobMasque.PutValue ('GDM_POSITION6', 'ON1');
       end;

       // Insertion dans DB
       TobMasque.InsertDB (nil);
       FreeAndNil( TobMasque );

       // Type de masque suivant
       Q.Next;
    end;
  end;
  Ferme(Q);
  // Retourne le code du masque créé
  result := CodeMasqueMax;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Recherche de la référence article à partir du code
Suite ........ : barres ou du code article et des codes dimensions GPAO
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.ReconstitueIdentArticle (TobInfo : TOB) : string ;
var Q              : TQUERY ;
    TypeDim        : string ;
    CodeGPAO       : string ;
    GrilleDim      : string ;
    CodeArticleGen : string ;
    Chtmp          : string ;
    StBlanc        : string ;
    CodeInterne1   : string ;
    CodeInterne2   : string ;
    CodeInterne3   : string ;
    CodeInterne4   : string ;
    CodeInterne5   : string ;
    Masque         : string ;
    CodeBarre      : string ;
    Longueur       : integer;
    i              : integer;
    LectureArticle : boolean;
    TrouveArticle  : boolean;
begin
  // Cas prioritaire : on travaille sur le GENCOD
  TrouveArticle := False ;
  result        := '';

  if (TobInfo.FieldExists ('$$_CODEBARRE')) and (TobInfo.GetValue ('$$_CODEBARRE') <> '') then
  begin
    IdentArticleImport := TobInfo.GetValue ('$$_CODEBARRE');
    CodeBarre          := TobInfo.GetValue ('$$_CODEBARRE');
    Q := OpenSQL ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEBARRE="'+CodeBarre+'"' ,True) ;
    //Multi code barres spécifique DEFI MODE
    if Q.EOF then
    begin
      if GetParamSoc('SO_ARTMULTICODEBARRE') then
      begin
        Ferme(Q);
        Q:=OpenSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEDOUANIER="'+CodeBarre+'"',True);
      end;
    end;
    if not Q.EOF then
    begin
      result := Q.FindField('GA_ARTICLE').asString;
    end else  EcrireInfo (TraduireMemoire('  Attention : le code barre ') + CodeBarre + TraduireMemoire(' ne correspond à aucun article !'), True, True, True);
    Ferme (Q);
  end else
  begin
    if (TobInfo.FieldExists ('$$_CODEARTICLE')) and (TobInfo.GetValue ('$$_CODEARTICLE') <> '') then
    begin
      IdentArticleImport := TobInfo.GetValue ('$$_CODEARTICLE');
      LectureArticle := False ;
      for i:=1 to 5 do
      begin
        if not TobInfo.FieldExists ('$$_GRILLEGPAODIM'+IntToStr(i)) then
        begin
          TobInfo.AddChampSup ('$$_GRILLEGPAODIM'+IntToStr(i), False);
          TobInfo.PutValue    ('$$_GRILLEGPAODIM'+IntToStr(i), '');
        end;
        if not TobInfo.FieldExists ('$$_CODEGPAODIM'+IntToStr(i)) then
        begin
          TobInfo.AddChampSup ('$$_CODEGPAODIM'+IntToStr(i), False);
          TobInfo.PutValue    ('$$_CODEGPAODIM'+IntToStr(i), '');
        end;
        if not TobInfo.FieldExists ('$$_CODEINTERNE'+IntToStr(i)) then
        begin
          TobInfo.AddChampSup ('$$_CODEINTERNE'+IntToStr(i), False);
          TobInfo.PutValue    ('$$_CODEINTERNE'+IntToStr(i), '');
        end;
        /////////////////////////////////////////////////////////////
        // MODIF LM 07/06/2002 pour DEFI MODE
        ////////////////////////////////////////////////////////////
        if (TobInfo.GetValue ('$$_CODEGPAODIM'+IntToStr(i)) <> '') and (TobInfo.GetValue ('$$_CODEINTERNE'+IntToStr(i)) = '') then LectureArticle := True;
      end;

      // Lecture de la fiche article pour récupérer le masque puis les grilles de tailles (si besoin)
      if LectureArticle then
      begin
        CodeArticleGen := TobInfo.GetValue ('$$_CODEARTICLE');
        ArticleCodeGen (CodeArticleGen);

        Q := OpenSQL ('SELECT GA_DIMMASQUE FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticleGen+'"' ,True) ;
        if not Q.EOF then
        begin
          TrouveArticle := True ;
          Masque        := Q.FindField('GA_DIMMASQUE').AsString;
          Ferme (Q);

          if Masque <> '' then
          begin
            Q := OpenSQL ('SELECT GDM_TYPE1, GDM_TYPE2, GDM_TYPE3, GDM_TYPE4, GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+Masque+'"' ,True) ;
            if not Q.EOF then
            begin
              for i:=1 to 5 do
              begin
                if Q.FindField('GDM_TYPE'+IntToStr(i)).AsString <> '' then TobInfo.PutValue ('$$_GRILLEGPAODIM'+IntToStr(i), Q.FindField('GDM_TYPE'+IntToStr(i)).AsString);
              end;
            end else EcrireInfo (TraduireMemoire('  Attention : Impossible trouver la masque ') + Masque + TraduireMemoire(' rattaché à l''article ' + CodeArticleGen), True, True, True);
            Ferme (Q);
          end;
        end else
        begin
          Ferme (Q);
          EcrireInfo (TraduireMemoire('  Attention : l''article ') + CodeArticleGen + TraduireMemoire(' n''existe pas !'), True, True, True);
        end;

        // Recherche des codes internes pour reconstituer l'identifiant article
        if TrouveArticle then
        begin
          for i:=1 to 5 do
          begin
            if (TobInfo.GetValue ('$$_GRILLEGPAODIM'+IntToStr(i)) <> '') and (TobInfo.GetValue ('$$_CODEGPAODIM'+IntToStr(i)) <> '') then
            begin
              TypeDim   := 'DI'+IntToStr(i);
              GrilleDim := TobInfo.GetValue ('$$_GRILLEGPAODIM'+IntToStr(i));

              if (TobInfo.FieldExists ('$$_RECHGPAOLIBDIM'+IntToStr(i))) and (TobInfo.GetValue ('$$_RECHGPAOLIBDIM'+IntToStr(i))='X') then
              begin
                CodeGPAO  := TobInfo.GetValue ('$$_CODEGPAODIM'+IntToStr(i));
                Q := OpenSQL ('SELECT GDI_CODEDIM FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+Grilledim+'" AND GDI_LIBELLE="'+CodeGPAO+'"' ,True) ;
                if not Q.EOF then
                begin
                  TobInfo.PutValue ('$$_CODEINTERNE'+IntToStr(i), Q.FindField('GDI_CODEDIM').AsString)
                end else
                begin
                   TrouveArticle := False ;
                   EcrireInfo ('  Erreur : Impossible de retrouver le libellé ' + CodeGPAO + ' dans la grille ' + Grilledim + ' de dimension ' + IntToStr(i), True, True, True);
                end;
                Ferme (Q);
              end else
              begin
                CodeGPAO  := TobInfo.GetValue ('$$_CODEGPAODIM'+IntToStr(i));
                Q := OpenSQL ('SELECT GDI_CODEDIM FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+Grilledim+'" AND GDI_DIMORLI="'+CodeGPAO+'"' ,True) ;
                if not Q.EOF then
                begin
                  TobInfo.PutValue ('$$_CODEINTERNE'+IntToStr(i), Q.FindField('GDI_CODEDIM').AsString)
                end else
                begin
                  TrouveArticle := False ;
                  EcrireInfo ('  Erreur : Impossible de retrouver le code GPAO ' + CodeGPAO + ' dans la grille ' + Grilledim + ' de dimension ' + IntToStr(i), True, True, True);
                end;
                Ferme (Q);
              end;
            end;
          end;
        end;
      end else TrouveArticle := True;
    end;

    // Reconstitution de l'identifiant article PGI (GQ_ARTICLE)
    if TrouveArticle then
    begin
      Chtmp    := TobInfo.GetValue ('$$_CODEARTICLE');
      Chtmp    := Trim   (Chtmp);
      longueur := length (Chtmp);
      StBlanc := '                  ';

      if TobInfo.GetValue ('$$_CODEINTERNE1') <> '' then CodeInterne1 := TobInfo.GetValue ('$$_CODEINTERNE1') else CodeInterne1 := '   ';
      if TobInfo.GetValue ('$$_CODEINTERNE2') <> '' then CodeInterne2 := TobInfo.GetValue ('$$_CODEINTERNE2') else CodeInterne2 := '   ';
      if TobInfo.GetValue ('$$_CODEINTERNE3') <> '' then CodeInterne3 := TobInfo.GetValue ('$$_CODEINTERNE3') else CodeInterne3 := '   ';
      if TobInfo.GetValue ('$$_CODEINTERNE4') <> '' then CodeInterne4 := TobInfo.GetValue ('$$_CODEINTERNE4') else CodeInterne4 := '   ';
      if TobInfo.GetValue ('$$_CODEINTERNE5') <> '' then CodeInterne5 := TobInfo.GetValue ('$$_CODEINTERNE5') else CodeInterne5 := '   ';

      Result := Chtmp + Copy(StBlanc, 1, 18-longueur) + CodeInterne1 + CodeInterne2 + CodeInterne3 + CodeInterne4 + CodeInterne5 + 'X';

      // Ultime tests : l'article dimensionné existe ?
      Q := OpenSQL ('SELECT GA_DIMMASQUE FROM ARTICLE WHERE GA_ARTICLE="'+Result+'"' ,True) ;
      if Q.EOF then
      begin
        EcrireInfo ((TraduireMemoire('  Attention : l''article générique ') + Result + TraduireMemoire (' les dimensions existent, mais pas l''article dimensionné !')), True, True, True);
        result := '' ;
      end;
      Ferme (Q);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Création d'un nouveau masque de dimension
Mots clefs ... : MASQUE DIMENSION
*****************************************************************}
function NewCodeTrans (Depot : string): string;
Var NextCode  : integer ;
begin
  // Recherche du dernier code inventaire créé
  NextCode:=1 ;
  while ExisteSQL('Select GIT_CODETRANS from TRANSINVENT '+
                  'where GIT_DEPOT="'+Depot+'" and '+
                  'GIT_CODETRANS="'+Format('%.3d',[NextCode])+'"') do inc(NextCode);
  Result:=Format('%.3d',[NextCode]) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Ecriture d'un message dans le fichier des anomalies
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.EcrireFichAno (Chtmp : String) ;
begin
  Writeln(FichierAno,Chtmp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Ecriture d'un message dans le fichier compte rendu
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.EcrireFichCpt (Chtmp : String) ;
begin
  Writeln(FichierTrt,Chtmp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Gestion des messages d'intégration
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.EcrireInfo (Chtmp : String; Cpte, Trt, Ano : boolean);
begin
  if (Cpte = True) then CompteRendu.lines.add(Chtmp);
  if (Trt  = True) then EcrireFichCpt (Chtmp);
  if (Ano  = True) then EcrireFichAno (Chtmp);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 17/10/2001
Modifié le ... :   /  /
Description .. : Affichage des compteurs
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.AfficheCompteur ;
begin
  NbreLu.Text     := IntToStr (CptEnregTot) ;
  NbreTraite.Text := IntToStr (CptTrt) ;
  NbreRejet.Text  := IntToStr (CptErr) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 17/10/2001
Modifié le ... :   /  /
Description .. : Affichage des infos (compteurs) en fin de traitement
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.AfficheFinTraitement ;
begin
  EcrireInfo ('  ', True, False, False);
  EcrireInfo (TraduireMemoire(' Fin de la Récupération des Informations le ') + DateToStr(Date) + TraduireMemoire(' à ') + TimeToStr(Time), True, True, True);

  if CptEnregTot > 1        then EcrireInfo ('   -> ' + IntToStr (CptEnregTot) + TraduireMemoire(' enregistrements lus ') , True, False, False)
  else                           EcrireInfo ('   -> ' + IntToStr (CptEnregTot) + TraduireMemoire(' enregistrement lu') , True, False, False);

  if CptTrt   > 1           then EcrireInfo ('   -> ' + IntToStr (CptTrt) + TraduireMemoire(' enregistrements traités ') , True, False, False)
  else                           EcrireInfo ('   -> ' + IntToStr (CptTrt) + TraduireMemoire(' enregistrement traité ') , True, False, False) ;

  if CptIntegre > 1         then EcrireInfo ('   -> ' + IntToStr (CptIntegre) + TraduireMemoire(' enregistrements intégrés ') , True, False, False)
  else if CptIntegre=1      then                        EcrireInfo ('   -> ' + IntToStr (CptIntegre) + TraduireMemoire(' enregistrement intégré ') , True, False, False)
  else if CptIntegre=0      then                        EcrireInfo (TraduireMemoire('   -> Aucun enregistrement intégré ') , True, False, False) ;

  if CptErr   > 1           then EcrireInfo ('   -> ' + IntToStr (CptErr) + TraduireMemoire(' enregistrements comportant des erreurs ') , True, False, False)
  else if CptErr = 1        then EcrireInfo ('   -> ' + IntToStr (CptErr) + TraduireMemoire(' enregistrement comportant des erreurs ') , True, False, False)
  else if CptErr = 0        then EcrireInfo (TraduireMemoire('   -> Aucune erreur d''intégration') , True, False, False);

  if CptEnreg-CptTrt > 2    then EcrireInfo ('   -> ' + IntToStr (CptEnreg-CptTrt) + TraduireMemoire(' enregistrements non repris ') , True, False, False)
  else if CptEnreg-CptTrt=1 then EcrireInfo ('   -> ' + IntToStr (CptEnreg-CptTrt) + TraduireMemoire(' enregistrement non repris ') , True, False, False)
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Lecture d'un enregistrement dans le fichier d'entrée
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.LectureEnregistrement : string ;
var Tampon         : char    ;         // Caractère lu dans le fichier
    NumRead        : longint ;         // Nombre de caractères lus
    Chtmp          : string  ;
    EnregAIntegrer : boolean ;
    Top            : char    ;
begin
  EnregAIntegrer := False ;
  // Position courante avant lecture de la nouvelle ligne
  PosTope := FilePos (Fichier);

  repeat
    if DebutFichier then PosTope := FilePos (Fichier);
    Chtmp := '';

    repeat
       BlockRead(Fichier, Tampon, 1, NumRead);
       if Tampon <> #10 then Chtmp := Chtmp + Tampon;
    until (NumRead = 0) or (Tampon = #13);

    if Length (Chtmp) > 7 then
    begin
      Top := Chtmp [6];
      if Top <> 'I' then
      begin
        DebutFichier   := False ;
        EnregAIntegrer := True  ;
        Inc (CptEnregTot) ;
      end;
    end;
  until (NumRead = 0) or (EnregAIntegrer) ;

  // Position courante de la lecture
  PosLecture := FilePos (Fichier);

  // On retourne l'enregistrement à intégrer
  if EnregAIntegrer then
  begin
    Inc (CptEnreg) ;
    result := Chtmp ;
  end else result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Topage d'un enregistrement dans le fichier d'entrée
Mots clefs ... : TOPAGE
*****************************************************************}
procedure TTImportGPAO.TopageEnregistrement (CharTop : char) ;
var NumWrite : longint ;         // Nombre de caractères lus
begin

  // Topage de la ligne : Repositionnement dans le fichier
  if PosTope = 0 then PosTope := PosTope + 5
  else PosTope := PosTope+6;
  seek (Fichier, PosTope);

  // Ecriture du top
  BlockWrite(Fichier, CharTop, 1, NumWrite);

  // Repositionnement pour suite du traitement
  seek (Fichier, PosLecture);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Topage de toutes les lignes d'une pièce
Mots clefs ... : TOPAGE
*****************************************************************}
procedure TTImportGPAO.TopagePiece (CharTop : char);
var PosPtr      : longint ;
    PosCourante : longint ;
    Tampon      : char    ;
    Chtmp       : string  ;
    NumRead     : integer ;
    NumWrite    : integer ;
begin
  // Position courante avant lecture de la nouvelle ligne
  seek (Fichier, PosTopePiece);
  repeat
    PosPtr := FilePos (Fichier);
    Chtmp  := '';

    repeat
      BlockRead(Fichier, Tampon, 1, NumRead);
      if Tampon <> #10 then Chtmp := Chtmp + Tampon;
    until (NumRead = 0) or (Tampon = #13);

    PosCourante := FilePos (Fichier);

    // Topage de la ligne : Repositionnement dans le fichier
    if PosPtr = 0 then PosPtr := PosPtr + 5
    else PosPtr := PosPtr+6;
    seek (Fichier, PosPtr);

    // Ecriture du top
    BlockWrite(Fichier, CharTop, 1, NumWrite);
    if CharTop = 'I' then Inc (CptIntegre) ;

    // Repositionnement pour suite du traitement
    seek (Fichier, PosCourante);
  until PosCourante >= PosLecture ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Chargement du descriptif du fichier d'entrée à partir de la
Suite ........ : provenance du fichier
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.ChargeParamReprise (var Info : string) : boolean ;
var IdentInfo   : string ;
    Source      : string ;
    TypeReprise : string ;
    TypeDonnees : string ;
    Q           : TQUERY ;
begin
  FreeAndNil( Tob_RepriseGPAO );

  Tob_RepriseGPAO := TOB.CREATE ('REPRISEGPAO', NIL, -1);

  Source    := Provenance.Value;
  IdentInfo := copy (Info, 1, 3);

  Q := OpenSQL ('SELECT * FROM REPRISEGPAO WHERE GRE_TYPEREPRISE="'+Source+'" AND GRE_DEBUTENREG="'+IdentInfo+'"' ,True) ;
  if not Q.EOF then
  begin
    // Chargement entête
    result := True ;
    Tob_RepriseGPAO.PutValue ('GRE_TYPEREPRISE' , Q.FindField('GRE_TYPEREPRISE' ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_TYPEDONNEES' , Q.FindField('GRE_TYPEDONNEES' ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_NOMTABLE'    , Q.FindField('GRE_NOMTABLE'    ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_DEBUTENREG'  , Q.FindField('GRE_DEBUTENREG'  ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_LONGUEURFIXE', Q.FindField('GRE_LONGUEURFIXE').AsString);
    Tob_RepriseGPAO.PutValue ('GRE_SEPARENREG'  , Q.FindField('GRE_SEPARENREG'  ).AsString);
    Ferme (Q);

    // Chargement des lignes "détail"
    TypeReprise := Tob_RepriseGPAO.GetValue ('GRE_TYPEREPRISE');
    TypeDonnees := Tob_RepriseGPAO.GetValue ('GRE_TYPEDONNEES');
    Q := OpenSQL ('SELECT * FROM REPRISECHGPAO WHERE GRC_TYPEREPRISE="'+TypeReprise+'" AND GRC_TYPEDONNEES="'+TypeDonnees+'"' ,True) ;
    Tob_RepriseGPAO.LoadDetailDB ('REPRISECHGPAO', '', '', Q, False);
    Ferme (Q);
  end else
  begin
    Ferme (Q);
    result := False ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Chargement du descriptif du fichier d'entrée à partir de la
Suite ........ : provenance du fichier
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.ChargeParamRepriseFixe : boolean ;
var Source      : string ;
    TypeReprise : string ;
    TypeDonnees : string ;
    Q           : TQUERY ;
    Combien     : integer;
begin
  result := False ;
  FreeAndNil( Tob_RepriseGPAO );
  Source := Provenance.Value;

  // Vérification que c'est bien une reprise sans entête fixe et que c'est unique
  Q:=OpenSQL('Select count(*) from REPRISEGPAO where GRE_TYPEREPRISE="'+Source+'"', True) ;
  if Q.EOF then
  begin
    PGIBox ('Impossible de récupérer le format du fichier !', '');
    Ferme(Q) ;
    exit ;
  end else
  begin
    Combien := Q.Fields[0].AsInteger ;
    Ferme(Q) ;
    if (Combien >1)  then
    begin
      PGIBox ('Sans entête de ligne, vous ne pouvez définir qu''une seule reprise par provenance de données !', '');
      PGIBox ('Pour cette provenance, vous avez ' + IntToStr (combien) +  ' reprises possibles !', '');
      result := False ;
      exit ;
    end;
  end;

  Tob_RepriseGPAO := TOB.CREATE ('REPRISEGPAO', NIL, -1);

  Q := OpenSQL ('SELECT * FROM REPRISEGPAO WHERE GRE_TYPEREPRISE="'+Source+'"' ,True) ;
  if not Q.EOF then
  begin
    Tob_RepriseGPAO.PutValue ('GRE_TYPEREPRISE' , Q.FindField('GRE_TYPEREPRISE' ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_TYPEDONNEES' , Q.FindField('GRE_TYPEDONNEES' ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_NOMTABLE'    , Q.FindField('GRE_NOMTABLE'    ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_DEBUTENREG'  , Q.FindField('GRE_DEBUTENREG'  ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_LONGUEURFIXE', Q.FindField('GRE_LONGUEURFIXE').AsString);
    Tob_RepriseGPAO.PutValue ('GRE_SEPARENREG'  , Q.FindField('GRE_SEPARENREG'  ).AsString);
    Tob_RepriseGPAO.PutValue ('GRE_ENTETELIG'   , Q.FindField('GRE_ENTETELIG'   ).AsString);
    Ferme (Q);

    // Chargement des lignes "détail"
    TypeReprise := Tob_RepriseGPAO.GetValue ('GRE_TYPEREPRISE');
    TypeDonnees := Tob_RepriseGPAO.GetValue ('GRE_TYPEDONNEES');
    Q := OpenSQL ('SELECT * FROM REPRISECHGPAO WHERE GRC_TYPEREPRISE="'+TypeReprise+'" AND GRC_TYPEDONNEES="'+TypeDonnees+'"' ,True) ;
    Tob_RepriseGPAO.LoadDetailDB ('REPRISECHGPAO', '', '', Q, False);
    result := True ;
    Ferme (Q);
  end else Ferme (Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Chargement du descriptif du fichier d'entrée à partir de la
Suite ........ : provenance du fichier
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.RepriseFichierEnteteFixe : boolean ;
var Source      : string ;
    Entete      : string ;
    Q           : TQUERY ;
    Combien     : integer;
begin
  result := True ;
  Source := Provenance.Value;

  // Vérification que c'est bien une reprise sans entête fixe et que c'est unique
  Q:=OpenSQL('Select count(*) from REPRISEGPAO where GRE_TYPEREPRISE="'+Source+'"', True) ;
  if Q.EOF then
  begin
    PGIBox ('Impossible de récupérer le format du fichier ! Provenance de donnée inconnue ', '');
    Ferme(Q) ;
    exit ;
  end else
  begin
    Combien := Q.Fields[0].AsInteger ;
    Ferme(Q) ;
    if (Combien =1)  then
    begin
      Q:=OpenSQL('Select * from REPRISEGPAO where GRE_TYPEREPRISE="'+Source+'"', True) ;
      Entete := Q.FindField('GRE_ENTETELIG').AsString;
      Ferme (Q) ;
      if Entete <> '-' then Entete:='X' ;
      if Entete = '-' then result := False ;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import DB de la TOB
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.ImportTobInfoGPAO;
begin
  // Import si informations dans la TOB
  if (Tob_InfoGPAO <> nil) and (Tob_InfoGPAO.detail.count > 0) then
  begin
//      Tob_InfoGPAO.SetAllModifie (True);
    Tob_InfoGPAO.InsertOrUpdateDB (FALSE);   // Mise à jour de l'enregistrement table
  end;
  // Libération de la TOB
  FreeAndNil( Tob_InfoGPAO );
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Recherche des codes internes à partir de la dimension
Suite ........ : ORLI (ou GPAO)
Mots clefs ... : CODE INTERNE;DIMORLI
*****************************************************************}
procedure TTImportGPAO.MAJSpecifDimension (var TobInfo : TOB) ;
var Q         : TQUERY ;
    TypeDim   : string ;
    GrilleDim : string ;
    CodeDim   : string ;
    CodeGPAO  : string ;
    CodeMax   : string ;
    Libelle   : string ;
    RangMax   : integer;
begin
  // Suppression des blancs dans le libellé
  Libelle := Trim(TobInfo.GetValue ('GDI_LIBELLE'));
  TobInfo.PutValue ('GDI_LIBELLE', Libelle);

  if (TobInfo.FieldExists ('$$_GCCATEGORIEDIM')) and (TobInfo.GetValue ('$$_GCCATEGORIEDIM') <> '') then
    TobInfo.PutValue ('GDI_TYPEDIM', TobInfo.GetValue ('$$_GCCATEGORIEDIM'));

  Typedim   := TobInfo.GetValue ('GDI_TYPEDIM');
  Grilledim := TobInfo.GetValue ('GDI_GRILLEDIM');
  Codedim   := TobInfo.GetValue ('GDI_CODEDIM');
  CodeGPAO  := TobInfo.GetValue ('GDI_DIMORLI');

  if (Trim(CodeDim) = '') and (Trim(CodeGPAO) <> '') then
  begin
    Q := OpenSQL ('SELECT * FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+Grilledim+'" AND GDI_DIMORLI="'+CodeGPAO+'"' ,True) ;
    if not Q.EOF then
    begin
      // On a trouvé la taille qui existe déjà
      // On récupère le code interne et le rang de présentation
      TobInfo.PutValue ('GDI_CODEDIM', Q.FindField('GDI_CODEDIM').AsString);
      TobInfo.PutValue ('GDI_RANG'   , Q.FindField('GDI_RANG').AsInteger);
      Ferme (Q);
    end else
    begin
      Ferme (Q);
      // La taille n'existe pas, on recherche le dernier code interne et le dernier rang de présentation
      CodeMax := '000';
      RangMax := 0 ;
      Q := OpenSQL ('SELECT * FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+Grilledim+'"' ,True) ;
      Q.First ;
      while not Q.EOF do
      begin
        if CodeMax < Q.FindField('GDI_CODEDIM').AsString then CodeMax := Q.FindField('GDI_CODEDIM').AsString ;
        if RangMax < Q.FindField('GDI_RANG').AsInteger   then RangMax := Q.FindField('GDI_RANG').AsInteger   ;
        Q.Next ;
      end;
      Ferme (Q);
      CodeCodeInterne (CodeMax); // Incrémente le code interne
      TobInfo.PutValue ('GDI_CODEDIM', CodeMax);
      TobInfo.PutValue ('GDI_RANG'   , RangMax+1);
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Ajout des champs spécifiques pour la suite des traitements
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.AjouteChampSpecifArticle (var Tob_Art : TOB) ;
var i  : integer;
begin
  // Collection de base
  if not Tob_Art.FieldExists ('$$_COLLECTIONBAS') then
  begin
    Tob_Art.AddChampSup ('$$_COLLECTIONBAS', False);
    Tob_Art.PutValue    ('$$_COLLECTIONBAS', '');
  end;

  // Familles 4 à 8
  for i:=4 to 8 do
  begin
    if not Tob_Art.FieldExists ('$$_FAMILLENIV'+IntToStr(i)) then
    begin
      Tob_Art.AddChampSup ('$$_FAMILLENIV'+IntToStr(i), False);
      Tob_Art.PutValue    ('$$_FAMILLENIV'+IntToStr(i), '');
    end;
  end;

  // Statistiques complémentaires
  for i:=1 to 2 do
  begin
    if not Tob_Art.FieldExists ('$$_STATART'+IntToStr(i)) then
    begin
      Tob_Art.AddChampSup ('$$_STATART'+IntToStr(i), False);
      Tob_Art.PutValue    ('$$_STATART'+IntToStr(i), '');
    end;
  end;

  // Codes dans GPAO
  for i:=1 to 5 do
  begin
    if not Tob_Art.FieldExists ('$$_CODEGPAODIM'+IntToStr(i)) then
    begin
      Tob_Art.AddChampSup ('$$_CODEGPAODIM'+IntToStr(i), False);
      Tob_Art.PutValue    ('$$_CODEGPAODIM'+IntToStr(i), '');
    end;
  end;

  // Devise des PA
  if not Tob_Art.FieldExists ('$$_CODEDEVISEPA') then
  begin
    Tob_Art.AddChampSup ('$$_CODEDEVISEPA', False);
    Tob_Art.PutValue    ('$$_CODEDEVISEPA', V_PGI.DevisePivot);
  end;

  // Devise de PR
  if not Tob_Art.FieldExists ('$$_CODEDEVISEPR') then
  begin
    Tob_Art.AddChampSup ('$$_CODEDEVISEPR', False);
    Tob_Art.PutValue    ('$$_CODEDEVISEPR', V_PGI.DevisePivot);
  end;

  // Devise des PV
  if not Tob_Art.FieldExists ('$$_CODEDEVISEPV') then
  begin
    Tob_Art.AddChampSup ('$$_CODEDEVISEPV', False);
    Tob_Art.PutValue    ('$$_CODEDEVISEPV', V_PGI.DevisePivot);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Gestion des dimensions articles :
Suite ........ : 1 - Détermination du statut de l'article
Suite ........ : 2 - Création du masque su inexistant
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.GestionDimensionArticle (var Tob_Art : TOB) ;
var TobGrille     : TOB      ;
    TobDimension  : TOB      ;
    Q             : TQUERY   ;
    SQL           : string   ;
    TypeDim       : string   ;
    StType        : string   ;
    CodeGrille    : string   ;
    LibGrille     : string   ;
    GrilleDim1    : string   ;
    GrilleDim2    : string   ;
    GrilleDim3    : string   ;
    GrilleDim4    : string   ;
    GrilleDim5    : string   ;
    NewMasque     : string   ;
    CodeChamp     : string   ;
    LibChamp      : string   ;
    CodeGPAO      : string   ;
    codeMax       : string   ;
    CodeBarre     : string   ;
    TypeCAB       : string   ;
    cpt           : integer  ;
    RangMax       : integer  ;
    ErreurGrille  : boolean  ;
begin
  ErreurGrille := False ;

  // Détermination du statut de l'article : DIM/UNI/GEN
  if (Tob_Art.GetValue ('GA_GRILLEDIM1')='') and (Tob_Art.GetValue ('GA_GRILLEDIM2')='') and (Tob_Art.GetValue ('GA_GRILLEDIM3')='') and (Tob_Art.GetValue ('GA_GRILLEDIM4')='') and (Tob_Art.GetValue ('GA_GRILLEDIM5')='') then
    Tob_Art.PutValue ('GA_STATUTART', 'UNI');

  // Calcul automatique de la clé EAN 13 si demandé
  if (Tob_Art.FieldExists ('$$_CALCULCLECB')) and (Tob_Art.GetValue ('$$_CALCULCLECB')='X') then
  begin
    TypeCAB   := Tob_Art.GetValue ('GA_QUALIFCODEBARRE');
    CodeBarre := Tob_Art.GetValue ('GA_CODEBARRE');
    if (TypeCAB='E13') or (TypeCAB='EA8') or (TypeCAB='39C') or (TypeCAB='ITC') or (TypeCAB='ITF') then
    begin
      CodeBarre := CodeBarre + CalculCleCodeBarre(CodeBarre+'0', TypeCAB );
      Tob_Art.PutValue ('GA_CODEBARRE', CodeBarre);
    end;
  end;

  // Contrôle de l'existence de la grille de dimension
  // Création de la grille de dimension si besoin
  if (Tob_Art.GetValue ('GA_STATUTART') = 'DIM') then
  begin
    for cpt:=1 to 5 do
    begin
      CodeChamp  := 'GA_GRILLEDIM' + IntToStr(cpt);
      LibChamp   := '$$_LIBGRILLE' + IntToStr(cpt);
      StType     := 'GG' + IntToStr(cpt);
      CodeGrille := Tob_Art.GetValue (CodeChamp) ;

      if (CodeGrille <> '') and (CodeGrille <> DerniereGrille [cpt]) then
      begin
        Q := OpenSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+StType+'" AND CC_CODE="'+CodeGrille+'"',True) ;
        if Q.EOF then
        begin
          Ferme (Q);
          // La grille n'existe pas , le libellé a été envoyé, on la crée
          if (Tob_Art.FieldExists (LibChamp)) and (Tob_Art.GetValue (LibChamp)<>'') then
          begin
            LibGrille := Tob_Art.GetValue (LibChamp) ;
            TobGrille := Tob.create ('CHOIXCOD', nil, -1);
            TobGrille.PutValue      ('CC_TYPE'   , StType);
            TobGrille.PutValue      ('CC_CODE'   , CodeGrille);
            TobGrille.PutValue      ('CC_LIBELLE', LibGrille);
            TobGrille.PutValue      ('CC_ABREGE' , copy (LibGrille, 1, 17));
            // Insertion dans DB
            TobGrille.InsertDB (nil);
            FreeAndNil( TobGrille );
            DerniereGrille [cpt] := CodeGrille ;
          end else
          begin
            ErreurGrille := True ;
          end;
        end else
        begin
          DerniereGrille [cpt] := CodeGrille ;
          Ferme (Q);
        end;
      end;
    end;

    // Contrôle de l'existence de la dimension
    // Création de la dimension si besoin
    if not ErreurGrille then
    begin
      // Etape 1 : Recherche des codes internes PGI à partir des codes dimensions GPAO
      for cpt:=1 to 5 do
      begin
        if Tob_Art.GetValue ('$$_CODEGPAODIM'+IntToStr(cpt)) <> '' then
        begin
          TypeDim   := 'DI'+IntToStr(cpt);
          CodeGrille:= Tob_Art.GetValue ('GA_GRILLEDIM'+IntToStr(cpt));
          CodeGPAO  := Tob_Art.GetValue ('$$_CODEGPAODIM'+IntToStr(cpt));
          Q := OpenSQL ('SELECT * FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+CodeGrille+'" AND GDI_DIMORLI="'+CodeGPAO+'"' ,True) ;
          if not Q.EOF then
          begin
            Tob_Art.PutValue ('GA_CODEDIM'+IntToStr(cpt), Q.FindField('GDI_CODEDIM').AsString);
            Ferme (Q) ;
          end else
          begin
            Ferme (Q);
            // La dimension n'existe pas , le libellé a été envoyé, on la crée
            if (Tob_Art.FieldExists ('$$_LIBGPAODIM'+IntToStr(cpt))) and (Tob_Art.GetValue ('$$_LIBGPAODIM'+IntToStr(cpt))<>'') then
            begin
              TobDimension := Tob.create ('DIMENSION', nil, -1);
              TobDimension.PutValue      ('GDI_TYPEDIM'   , TypeDim);
              TobDimension.PutValue      ('GDI_GRILLEDIM' , CodeGrille);
              TobDimension.PutValue      ('GDI_LIBELLE'   , Tob_Art.GetValue ('$$_LIBGPAODIM'+IntToStr(cpt)));
              TobDimension.PutValue      ('GDI_DIMORLI'   , Tob_Art.GetValue ('$$_CODEGPAODIM'+IntToStr(cpt)));

              Q := OpenSQL ('SELECT MAX(GDI_CODEDIM), MAX(GDI_RANG) FROM DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+CodeGrille+'"' ,True) ;
              CodeMax := Q.Fields[0].AsString ;
              CodeCodeInterne (CodeMax); // Incrémente le code interne
              RangMax := Q.Fields[1].AsInteger ;
              Ferme (Q);

              TobDimension.PutValue ('GDI_CODEDIM', CodeMax);
              TobDimension.PutValue ('GDI_RANG'   , RangMax+1);

              // Initialisation du code interne de l'article
              Tob_Art.PutValue ('GA_CODEDIM'+IntToStr(cpt), CodeMax);

              // Insertion dans DB
              TobDimension.InsertDB (nil);
              FreeAndNil( TobDimension );
            end else
            begin
              ErreurGrille := True ;
            end;
          end;
        end;
      end;
    end;

    // Détermination du masque de dimension s'il n'est pas renseigné
    // Création du masque au besoin
    if not ErreurGrille then
    begin
      if ((Tob_Art.GetValue ('GA_STATUTART') = 'DIM') or (Tob_Art.GetValue ('GA_STATUTART') = 'GEN')) and (Tob_Art.GetValue ('GA_DIMMASQUE') = '') then
      begin
        GrilleDim1 := Tob_Art.GetValue ('GA_GRILLEDIM1');
        GrilleDim2 := Tob_Art.GetValue ('GA_GRILLEDIM2');
        GrilleDim3 := Tob_Art.GetValue ('GA_GRILLEDIM3');
        GrilleDim4 := Tob_Art.GetValue ('GA_GRILLEDIM4');
        GrilleDim5 := Tob_Art.GetValue ('GA_GRILLEDIM5');

        SQL := 'SELECT * FROM DIMMASQUE WHERE GDM_TYPE1="'+GrilleDim1+'" AND GDM_TYPE2="'+GrilleDim2+'" AND GDM_TYPE3="'+GrilleDim3+'" AND GDM_TYPE4="'+GrilleDim4+'" AND GDM_TYPE5="'+GrilleDim5+'"' ;
        Q := OpenSQL (SQL ,True) ;
        if not Q.EOF then
        begin
          Tob_Art.PutValue ('GA_DIMMASQUE', Q.FindField('GDM_MASQUE').AsString);
          Ferme (Q);
        end else
        begin
          Ferme (Q);
          NewMasque := CreateNewMasque (GrilleDim1, GrilleDim2, GrilleDim3, GrilleDim4, GrilleDim5);
          Tob_Art.PutValue ('GA_DIMMASQUE', NewMasque);
        end;
      end;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 20/09/2001
Modifié le ... :   /  /
Description .. : Conversion des prix de la fiche article dans la devise de
Suite ........ : tenue du dossier
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.ConvPrixMajPartielleArticle (var Tob_Art : TOB) ;
Var
    Date          : TDateTime;
    Q             : TQUERY   ;
    SQL           : string   ;
    CodeDevisePR  : string   ;
    CodeDevisePA  : string   ;
    CodeDevisePV  : string   ;
    CodeArticle   : string   ;
    NomChamp      : string   ;
    Recup_DPA     : boolean  ;
    Recup_PAHT    : boolean  ;
    Recup_PMAP    : boolean  ;
    Recup_DPR     : boolean  ;
    Recup_PRHT    : boolean  ;
    Recup_PMRP    : boolean  ;
    Recup_PVTTC   : boolean  ;
    Recup_PVHT    : boolean  ;
    TobFilleParam : TOB      ;
    Tob_Article   : TOB      ;
    Montant       : double   ;
    MontantDev    : double   ;
    MontantOld    : double   ;
    DPR           : double   ;
    PRHT          : double   ;
    PMRP          : double   ;
    i             : integer  ;
begin
  // Initialisation
  Recup_DPA   := False ;
  Recup_PAHT  := False ;
  Recup_PMAP  := False ;
  Recup_DPR   := False ;
  Recup_PRHT  := False ;
  Recup_PMRP  := False ;
  Recup_PVTTC := False ;
  Recup_PVHT  := False ;
  Date    := NowH ;

  // Chargement des taux
  CodeDevisePR := Tob_Art.GetValue ('$$_CODEDEVISEPR');
  if CodeDevisePR = '' then CodeDevisePR := V_PGI.DevisePivot;

  CodeDevisePA := Tob_Art.GetValue ('$$_CODEDEVISEPA');
  if CodeDevisePA = '' then CodeDevisePA := V_PGI.DevisePivot;

  CodeDevisePV := Tob_Art.GetValue ('$$_CODEDEVISEPV');
  if CodeDevisePV = '' then CodeDevisePV := V_PGI.DevisePivot;

  // Conversion des PA, PR et PV si besoin
  // Calcul des montants en contre valeur

  // Dernier prix d'achat
  Montant := Tob_Art.GetValue ('GA_DPA');
  if Montant <> 0.0 then
  begin
    recup_DPA  := True ;
    if CodeDevisePA <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_DPA' , MontantDev);
    end ;
  end;

  // Prix de base d'achat HT
  Montant := Tob_Art.GetValue ('GA_PAHT');
  if Montant <> 0.0 then
  begin
    recup_PAHT := True ;
    if CodeDevisePA <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PAHT'    , MontantDev);
    end ;
  end;

  // Prix d'achat moyen pondéré
  Montant := Tob_Art.GetValue ('GA_PMAP');
  if Montant <> 0.0 then
  begin
    recup_PMAP := True ;
    if CodeDevisePA <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PMAP'    , MontantDev);
    end ;
  end;

  // Prix de revient
  Montant := Tob_Art.GetValue ('GA_DPR');
  if Montant <> 0.0 then
  begin
    recup_DPR := True ;
    if CodeDevisePR <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_DPR'    , MontantDev);
    end ;
  end;

  // Prix de base revient HT
  Montant := Tob_Art.GetValue ('GA_PRHT');
  if Montant <> 0.0 then
  begin
    recup_PRHT := True ;
    if CodeDevisePR <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PRHT', MontantDev);
    end ;
  end;

  // Prix de revient moyen pondéré
  Montant := Tob_Art.GetValue ('GA_PMRP');
  if Montant <> 0.0 then
  begin
    recup_PMRP := True ;
    if CodeDevisePR <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PMRP', MontantDev);
    end ;
  end;

  // PV TTC
  Montant := Tob_Art.GetValue ('GA_PVTTC');
  if Montant <> 0.0 then
  begin
    recup_PVTTC   := True ;
    if CodeDevisePV <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PVTTC', MontantDev);
    end ;
  end;

  // PV HT
  Montant := Tob_Art.GetValue ('GA_PVHT');
  if Montant <> 0.0 then
  begin
    recup_PVHT := True ;
    if CodeDevisePV <> V_PGI.DevisePivot then
    begin
      ToxConvertir  (Montant, MontantDev, MontantOld, CodeDevisePR, V_PGI.DevisePivot, Date, nil) ;
      Tob_Art.PutValue ('GA_PVHT', MontantDev);
    end ;
  end;

  // Faut-il mettre à jour partiellement la fiche article ?
  if (Tob_Art.FieldExists ('$$_MAJPARTIELLE')) and (Tob_Art.GetValue ('$$_MAJPARTIELLE')='X') then
  begin
    CodeArticle := Tob_Art.GetValue ('GA_ARTICLE');
    SQL := 'SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+CodeArticle+'"' ;
    Q   := OpenSQL (SQL ,True) ;
    if not Q.EOF then
    begin
      Tob_Article := TOB.CREATE ('ARTICLE', NIL, -1);
      Tob_Article.SelectDB('',Q);
      Ferme(Q) ;
      // MODIF JACADI 23/05/03 : conservation des DPR/PRHT
      DPR     := Tob_Article.GetValue ('GA_DPR');
      PRHT    := Tob_Article.GetValue ('GA_PRHT');
      PMRP    := Tob_Article.GetValue ('GA_PMRP');
      // on bascule tous les champs récupérés dans Tob_Info dans Tob_Article
      for i:=0 to Tob_RepriseGPAO.detail.count-1 do
      begin
        TobFilleParam := Tob_RepriseGPAO.Detail [i];
        NomChamp      := TobFilleParam.GetValue ('GRC_NOMCHAMP');
        // Création des éventuels champs supplémentaires
        if not Tob_Article.FieldExists (NomChamp) then
        begin
          Tob_Article.AddChampSup (NomChamp, False);
          Tob_Article.PutValue (NomChamp, '');
        end;
        Tob_Article.PutValue (NomChamp, Tob_Art.GetValue (NomChamp));
      end;

      // Initialisation des prix
      if Recup_DPA = True   then Tob_Article.PutValue ('GA_DPA'   , Tob_Art.GetValue ('GA_DPA'));
      if Recup_PAHT = True  then Tob_Article.PutValue ('GA_PAHT'  , Tob_Art.GetValue ('GA_PAHT'));
      if Recup_PMAP = True  then Tob_Article.PutValue ('GA_PMAP'  , Tob_Art.GetValue ('GA_PMAP'));
      if Recup_DPR = True   then Tob_Article.PutValue ('GA_DPR'   , Tob_Art.GetValue ('GA_DPR'));
      // MODIF JACADI 23/05/2003 pour ne pas écraser à 0 le DPR
      if (Tob_Article.GetValue ('GA_DPR')    = 0) and (DPR    <> 0) then Tob_Article.PutValue ('GA_DPR'   , DPR);
      if Recup_PRHT = True  then Tob_Article.PutValue ('GA_PRHT'  , Tob_Art.GetValue ('GA_PRHT'));
      // MODIF JACADI 23/05/2003 pour ne pas écraser à 0 le PRHT
      if (Tob_Article.GetValue ('GA_PRHT')    = 0) and (PRHT    <> 0) then Tob_Article.PutValue ('GA_PRHT'   , DPR);
      if Recup_PMRP = True  then Tob_Article.PutValue ('GA_PMRP'  , Tob_Art.GetValue ('GA_PMRP'));
      // MODIF JACADI 23/05/2003 pour ne pas écraser à 0 le PMRP
      if (Tob_Article.GetValue ('GA_PMRP')    = 0) and (PMRP    <> 0) then Tob_Article.PutValue ('GA_PMRP'   , PMRP);
      if Recup_PVTTC = True then Tob_Article.PutValue ('GA_PVTTC'   , Tob_Art.GetValue ('GA_PVTTC'));
      if Recup_PVHT = True  then Tob_Article.PutValue ('GA_PVHT'   , Tob_Art.GetValue ('GA_PVHT'));

      Tob_Art.dupliquer (Tob_article, False, True, True);

      // Rajout des champs spécifiques
      AjouteChampSpecifArticle (Tob_Art) ;
      FreeAndNil( Tob_Article );
    end else Ferme (Q);
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique d'une TOB article
Mots clefs ... : IMPORT ARTICLE;
*****************************************************************}
procedure TTImportGPAO.MAJSpecifArticle (var TobInfo : TOB) ;
begin
  // Ajout des champs spécifiques pour la suite des traitements
  AjouteChampSpecifArticle (TobInfo) ;

  // Gestion des dimensions
  GestionDimensionArticle (TobInfo) ;

  // Détermination de l'identifiant article (GA_ARTICLE) = Code Article + Codes internes + X
  ArticleCodeDim (TobInfo) ;

  // Conversion des PA, PR et PV dans la devise de tenue du dossier
  // MAJ partielle de la fiche article
  ConvPrixMajPartielleArticle (TobInfo) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique d'une TOB stock (DISPO)
Mots clefs ... : IMPORT STOCK
*****************************************************************}
procedure TTImportGPAO.MAJSpecifDispo (var TobInfo : TOB) ;
var Date          : TdateTime ;
    Q             : TQUERY    ;
    DEV           : RDEVISE   ;
    Tob_Dispo     : TOB       ;
    TobFilleParam : TOB       ;
    SQL           : string    ;
    Chtmp         : string    ;
    CodeArticle   : string    ;
    CodeDepot     : string    ;
    Cloture       : string    ;
    CodeDevise    : string    ;
    NomChamp      : string    ;
    DateCloture   : TDateTime ;
    Montant       : double    ;
    Recup_DPA     : boolean   ;
    Recup_DPR     : boolean   ;
    Recup_PMRP    : boolean   ;
    Recup_PMAP    : boolean   ;
    i             : integer   ;
begin
  if (TobInfo.FieldExists ('$$_RAZSTOCK')) and (TobInfo.GetValue ('$$_RAZSTOCK')='X') then
  begin
    // Suppression des stocks d'un dépôt
  end else
  begin
    Recup_DPA  := False ;
    Recup_DPR  := False ;
    Recup_PMRP := False ;
    Recup_PMAP := False ;
    //
    // Forcage des éléments de la cléf = Zones obligatoires
    // Attention : les clôtures de stock ne sont pas repris
    //
    // Pour le démarrage LMV, reprise du stock de démarrage en tant que "stock cloturé"
    // Test de sécurité : Si Stock non clôturé, la date est forcée au 01/01/1900
    //
    //TobInfo.PutValue ('GQ_DATECLOTURE', StrToDate('01/01/1900'));
    //TobInfo.PutValue ('GQ_CLOTURE'    , '-');
    if TobInfo.GetValue ('GQ_CLOTURE') = '-' then TobInfo.PutValue ('GQ_DATECLOTURE', StrToDate('01/01/1900'));

    // Récupération des 5 codes internes en fonction des codes dimensions GPAO
    if TobInfo.GetValue ('GQ_ARTICLE') = '' then
    begin
      Chtmp := ReconstitueIdentArticle (TobInfo);
      TobInfo.PutValue('GQ_ARTICLE', Chtmp);
    end;

    // Conversion des prix en devise dans la devise de tenue du dossier
    Date:= NowH ;
    if not TobInfo.FieldExists ('$$_CODEDEVISEPA') then
    begin
      TobInfo.AddChampSup ('$$_CODEDEVISEPA', False);
      TobInfo.PutValue    ('$$_CODEDEVISEPA', V_PGI.DevisePivot);
    end;
    if not TobInfo.FieldExists ('$$_CODEDEVISEPR') then
    begin
      TobInfo.AddChampSup ('$$_CODEDEVISEPR', False);
      TobInfo.PutValue    ('$$_CODEDEVISEPR', V_PGI.DevisePivot);
    end;
    // PA
    Montant := TobInfo.GetValue ('GQ_DPA');
    if Montant <> 0.0 then
    begin
      recup_DPA  := True ;
      CodeDevise := TobInfo.GetValue ('$$_CODEDEVISEPA');
      if CodeDevise <> V_PGI.DevisePivot then
      begin
        DEV.Code:=CodeDevise ;
        DEV.DateTaux := NowH ;
      	GetInfosDevise(DEV) ;
    	DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
        Montant := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
        TobInfo.PutValue ('GQ_DPA', Montant);
      end;
    end;
    // PR
    Montant := TobInfo.GetValue ('GQ_DPR');
    if Montant <> 0.0 then
    begin
      recup_DPR  := True ;
      CodeDevise := TobInfo.GetValue ('$$_CODEDEVISEPR');
      if CodeDevise <> V_PGI.DevisePivot then
      begin
        DEV.Code:=CodeDevise ;
        DEV.DateTaux := NowH ;
      	GetInfosDevise(DEV) ;
    	DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
        Montant := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
        TobInfo.PutValue ('GQ_DPR', Montant);
      end;
    end;
    // PMAP
    Montant := TobInfo.GetValue ('GQ_PMAP');
    if Montant <> 0.0 then
    begin
      recup_PMAP := True ;
      CodeDevise := TobInfo.GetValue ('$$_CODEDEVISEPA');
      if CodeDevise <> V_PGI.DevisePivot then
      begin
        DEV.Code:=CodeDevise ;
        DEV.DateTaux := NowH ;
        GetInfosDevise(DEV) ;
        DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
        Montant := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
        TobInfo.PutValue ('GQ_PMAP', Montant);
      end;
    end;
    // PMRP
    Montant := TobInfo.GetValue ('GQ_PMRP');
    if Montant <> 0.0 then
    begin
      recup_PMRP := True ;
      CodeDevise := TobInfo.GetValue ('$$_CODEDEVISEPR');
      if CodeDevise <> V_PGI.DevisePivot then
      begin
        DEV.Code:=CodeDevise ;
        DEV.DateTaux := NowH ;
        GetInfosDevise(DEV) ;
        DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, Date) ;
        Montant := DeviseToPivot (Montant, DEV.Taux, DEV.Quotite);
        TobInfo.PutValue ('GQ_PMRP', Montant);
      end;
    end;

    if (TobInfo.FieldExists ('$$_MAJPARTIELLE')) and (TobInfo.GetValue ('$$_MAJPARTIELLE')='X') then
    begin
      CodeArticle := TobInfo.GetValue ('GQ_ARTICLE');
      CodeDepot   := TobInfo.GetValue ('GQ_DEPOT');
      Cloture     := TobInfo.GetValue ('GQ_CLOTURE');
      DateCloture := TobInfo.GetValue ('GQ_DATECLOTURE');

      SQL:='Select GQ_STOCKMIN From DISPO WHERE GQ_ARTICLE="'+CodeArticle+'" AND GQ_DEPOT="'+CodeDepot+'" AND GQ_CLOTURE="'+Cloture+'" AND GQ_DATECLOTURE="'+UsDateTime(DateCloture)+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        // Chargement de l'article déjà existant en TOB
        // Rq : le selectDB ne fonctionne pas !!!?????
        //      on charge avec LoadDetailDB et in travaille sur la fille
        Tob_Dispo := TOB.CREATE ('DISPO', NIL, -1);
        Tob_Dispo.SelectDB('',Q);
        Ferme(Q) ;

        // on bascule tous les champs récupérés dans Tob_Info dans Tob_Article
        for i:=0 to Tob_RepriseGPAO.detail.count-1 do
        begin
          TobFilleParam := Tob_RepriseGPAO.Detail [i];
          NomChamp      := TobFilleParam.GetValue ('GRC_NOMCHAMP');

          // Création des éventuels champs supplémentaires
          if not Tob_Dispo.FieldExists (NomChamp) then
          begin
            Tob_Dispo.AddChampSup (NomChamp, False);
            Tob_Dispo.PutValue (NomChamp, '');
          end;
          Tob_Dispo.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
        end;

        // Initialisation des prix
        if Recup_DPA = True  then Tob_Dispo.PutValue ('GQ_DPA'  , TobInfo.GetValue ('GQ_DPA'));
        if Recup_DPR = True  then Tob_Dispo.PutValue ('GQ_DPR'  , TobInfo.GetValue ('GQ_DPR'));
        if Recup_PMAP = True then Tob_Dispo.PutValue ('GQ_PMAP' , TobInfo.GetValue ('GQ_PMAP'));
        if Recup_PMRP = True then Tob_Dispo.PutValue ('GQ_PMRP' , TobInfo.GetValue ('GQ_PMRP'));

        TobInfo.Dupliquer( Tob_Dispo, True, True, True);
        FreeAndNil( Tob_Dispo );
      end else Ferme (Q);
    end;
    // MAJ Date de modif
    if TobInfo <> nil then TobInfo.SetDateModif(NowH);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique d'un inventaire
Mots clefs ... : IMPORT INVENTAIRES
*****************************************************************}
procedure TTImportGPAO.MAJSpecifInventaire (var TobInfo : TOB) ;
var CodeArticleGen : string ;
    Chtmp          : string ;
begin

  // Ajout de champs supplémentaires
  if not TobInfo.FieldExists ('$$_LIBELLE') then
  begin
    TobInfo.AddChampSup ('$$_LIBELLE', False);
    TobInfo.PutValue    ('$$_LIBELLE', '');
  end;

  // Ajout de champs supplémentaires
  if not TobInfo.FieldExists ('$$_CODEARTICLE') then
  begin
    TobInfo.AddChampSup ('$$_CODEARTICLE', False);
    TobInfo.PutValue    ('$$_CODEARTICLE', TobInfo.GetValue ('GIN_CODEARTICLE'));
  end;
  if not TobInfo.FieldExists ('$$_CODEBARRE') then
  begin
    TobInfo.AddChampSup ('$$_CODEBARRE', False);
    TobInfo.PutValue    ('$$_CODEBARRE', TobInfo.GetValue ('GIN_CODEBARRE'));
  end;

  // Initialisation des articles des codes articles
  if TobInfo.GetValue ('GIN_ARTICLE') = '' then
  begin
    // Code article
    Chtmp := ReconstitueIdentArticle (TobInfo);
    TobInfo.PutValue('GIN_ARTICLE', Chtmp);
    // Code article générique
    CodeArticleGen := Copy(Chtmp, 1, 18) ;
    TobInfo.PutValue('GIN_CODEARTICLE', CodeArticleGen);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique d'un inventaire
Mots clefs ... : IMPORT INVENTAIRES
*****************************************************************}
procedure TTImportGPAO.MAJSpecifFidelite (var TobInfo : TOB) ;
begin
  // Ajout de champs supplémentaires
  if not TobInfo.FieldExists ('$$_TIERS') then
  begin
    TobInfo.AddChampSup ('$$_TIERS', False);
    TobInfo.PutValue    ('$$_TIERS', '');
  end;
  if not TobInfo.FieldExists ('$$_NUMCARTEEXT') then
  begin
    TobInfo.AddChampSup ('$$_NUMCARTEEXT', False);
    TobInfo.PutValue    ('$$_NUMCARTEEXT', '');
  end;
  if not TobInfo.FieldExists ('$$_DATEOUVERTURE') then
  begin
    TobInfo.AddChampSup ('$$_DATEOUVERTURE', False);
    TobInfo.PutValue    ('$$_DATEOUVERTURE', iDate1900);
  end;
  if not TobInfo.FieldExists ('$$_FERME') then
  begin
    TobInfo.AddChampSup ('$$_FERME', False);
    TobInfo.PutValue    ('$$_FERME', '-');
  end;
  if not TobInfo.FieldExists ('$$_DATEFERMETURE') then
  begin
    TobInfo.AddChampSup ('$$_DATEFERMETURE', False);
    TobInfo.PutValue    ('$$_DATEFERMETURE', iDate1900);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Chargement en TOB du qualifiant tarif en fonction du code
Suite ........ : tarif et de la période
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.ChargeTobQualifiantTarif(var Tob_Tarif : TOB; TypeTarif : string; TypePeriode : string) : boolean ;
var Tob_TypeTar    : TOB    ;
    Tob_Periode    : TOB    ;
    CodeMax        : string ;
    Libelle        : string ;
	Q  		       : TQUERY ;
    QMax           : TQUERY ;
    CodeMaxInt     : integer;
    Trouve_Periode : boolean;
    Trouve_TypeTar : boolean;
begin
  result := True ;
  Tob_Periode := nil ;
  Tob_TypeTar := nil ;
  CodeMaxInt  := 0   ;

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
      // L'enregistrement dans la table TARIFMODE n'existe pas, il faut le créer
      Ferme (Q) ;
      // Chargement en TOB de la période de tarif
      Q:=OpenSql('Select * from TARIFPER where GFP_CODEPERIODE="'+TypePeriode+'" AND GFP_ETABLISSEMENT="..."',True) ;
      if not Q.EOF then
      begin
        Tob_Periode := TOB.Create('TARIFPER',NIL,-1);
        Tob_Periode.SelectDB('',Q);
        Trouve_Periode := True ;
      end else
      begin
        EcrireInfo (TraduireMemoire('  TABLE TARIF : La période ') + TypePeriode + TraduireMemoire(' n''existe pas'), True, True, True);
        result := False ;
      end;
      Ferme (Q);
      // Chargement en TOB du type de tarif
      Q:=OpenSql('Select * from TARIFTYPMODE where GFT_CODETYPE="'+TypeTarif+'"',True) ;
      if not Q.EOF then
      begin
        Tob_TypeTar := TOB.Create('TARIFTYPMODE',NIL,-1);
      	Tob_TypeTar.SelectDB('', Q);
        Trouve_TypeTar := True ;
      end else
      begin
        EcrireInfo (TraduireMemoire('  TABLE TARIF : Le code tarif ') + TypeTarif + TraduireMemoire(' n''existe pas'), True, True, True) ;
        result := False ;
      end;
      Ferme (Q);
      // Chargement en TOB et création de l'enregistrement dans la table TARIFMODE
      if (trouve_Periode) and (Trouve_TypeTar) then
      begin
        QMax := OpenSQL ('SELECT MAX(GFM_TARFMODE) FROM TARIFMODE',True);
        if Not QMax.EOF then
        begin
          CodeMax:=QMax.Fields[0].AsString;
        end ;
        Ferme(QMax) ;

        if CodeMax <> '' then CodeMaxInt := StrToInt(CodeMax)+1;
        If CodeMaxInt = 0 then CodeMaxInt := 1;

        Tob_Tarif.PutValue('GFM_TARFMODE'    , CodeMaxInt);
        Tob_Tarif.PutValue('GFM_TYPETARIF'   , Tob_TypeTar.GetValue('GFT_CODETYPE'));
        Tob_Tarif.PutValue('GFM_PERTARIF'    , Tob_Periode.GetValue('GFP_CODEPERIODE'));
        Libelle := Tob_TypeTar.GetValue('GFT_LIBELLE') + '-' + Tob_Periode.GetValue('GFP_LIBELLE');
        Tob_Tarif.PutValue('GFM_LIBELLE'     , copy (Libelle, 1, 35));
        Tob_Tarif.PutValue('GFM_DATEDEBUT'   , Tob_Periode.GetValue('GFP_DATEDEBUT'));
        Tob_Tarif.PutValue('GFM_DATEFIN'     , Tob_Periode.GetValue('GFP_DATEFIN'));
        Tob_Tarif.PutValue('GFM_PROMO'       , Tob_Periode.GetValue('GFP_PROMO'));
        Tob_Tarif.PutValue('GFM_COEF'        , Tob_TypeTar.GetValue('GFT_COEF'));
        Tob_Tarif.PutValue('GFM_ARRONDI'     , Tob_Periode.GetValue('GFP_ARRONDI'));
        Tob_Tarif.PutValue('GFM_DEMARQUE'    , Tob_Periode.GetValue('GFP_DEMARQUE'));
        Tob_Tarif.PutValue('GFM_DEVISE'      , Tob_TypeTar.GetValue('GFT_DEVISE'));
        Tob_Tarif.PutValue('GFM_ETABLISREF'  , Tob_TypeTar.GetValue('GFT_ETABLISREF'));
        Tob_Tarif.PutValue('GFM_CASCADE'     , Tob_Periode.GetValue('GFP_CASCADE'));
        Tob_Tarif.PutValue('GFM_NATURETYPE'  , Tob_TypeTar.GetValue('GFT_NATURETYPE'));
        // MAJ DB
        Tob_Tarif.InsertDB (nil, False);
        // Libération des TOB
        FreeAndNil( Tob_TypeTar );
        FreeAndNil( Tob_Periode );
      end else FreeAndNil( Tob_Tarif );
    end;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 24/07/2001
Description .. : Récupération du numéro maximum de tarif
Mots clefs ... : NUMERO TARIF
*****************************************************************}
procedure TTImportGPAO.RecupMaxTarif ;
var
  QMax    : TQUERY ;
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

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Mise à jour spécifique de la TOB Tarif.
Suite ........ : Attention ; si tarif PV, il faut obligatoirement travailler avec
Suite ........ : un code tarif et une période
Mots clefs ... : IMPORT TARIF
*****************************************************************}
procedure TTImportGPAO.MAJSpecifTarif (var TobInfo : TOB) ;
begin
  // Existence des champs spécifiques période et code tarif
  if not TobInfo.FieldExists ('$$_CODEPERIODE') then
  begin
    TobInfo.AddChampSup ('$$_CODEPERIODE', False);
    TobInfo.PutValue    ('$$_CODEPERIODE', '');
  end;

  if not TobInfo.FieldExists ('$$_CODETYPE') then
  begin
    TobInfo.AddChampSup ('$$_CODETYPE', False);
    TobInfo.PutValue    ('$$_CODETYPE', '');
  end;

  if not TobInfo.FieldExists ('$$_PRIXARTICLE') then
  begin
    TobInfo.AddChampSup ('$$_PRIXARTICLE', False);
    TobInfo.PutValue    ('$$_PRIXARTICLE', '-');
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/09/2001
Modifié le ... :   /  /
Description .. : Fonction spécifique pour la reprise des devises
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MAJSpecifDevise (var TobInfo : TOB) ;
var Q             : TQUERY  ;
    Tob_Devise    : TOB     ;
    Tob_Fille     : TOB     ;
    SQL           : string  ;
    CodeDevise    : string  ;
    NomChamp      : string  ;
    i             : integer ;
begin
  // MAJ partielle de la devise dans PGI pour ne pas écraser des informations comprométantes
  CodeDevise := TobInfo.GetValue ('D_DEVISE');

  SQL:='Select * From DEVISE WHERE D_DEVISE="'+CodeDevise+'"';
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then
  begin
    Tob_Devise := TOB.CREATE ('DEVISE', NIL, -1);
    Tob_Devise.SelectDB('',Q);
    Ferme(Q) ;

    // on bascule tous les champs récupérés dans Tob_Info dans Tob_Devise
    for i:=0 to Tob_RepriseGPAO.detail.count-1 do
    begin
      Tob_Fille := Tob_RepriseGPAO.Detail [i];
      NomChamp  := Tob_Fille.GetValue ('GRC_NOMCHAMP');

      // Création des éventuels champs supplémentaires
      if not Tob_Devise.FieldExists (NomChamp) then
      begin
        Tob_Devise.AddChampSup (NomChamp, False);
        Tob_Devise.PutValue (NomChamp, '');
      end;

      Tob_Devise.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
    end;
    //TobInfo := Tob_Devise;
    TobInfo.Dupliquer (Tob_Devise, False, True, True);

    FreeAndNil( Tob_Devise );
  end else Ferme (Q);

  if TobInfo.GetValue ('D_SYMBOLE') = '' then TobInfo.PutValue ('D_SYMBOLE', TobInfo.GetValue ('D_DEVISE'));
  if TobInfo.GetValue ('D_QUOTITE') = 0  then TobInfo.PutValue ('D_QUOTITE', 1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/09/2001
Modifié le ... :   /  /
Description .. : Fonction spécifique pour la reprise des devises
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MAJSpecifTarifTypMode (var TobInfo : TOB) ;
var Q                : TQUERY  ;
    Tob_TarifTypMode : TOB     ;
    Tob_Fille        : TOB     ;
    SQL              : string  ;
    CodeTarif        : string  ;
    NomChamp         : string  ;
    i                : integer ;
begin
  // MAJ partielle de la devise dans PGI pour ne pas écraser des informations comprométantes
  CodeTarif := TobInfo.GetValue ('GFT_CODETYPE');

  SQL:='Select * From TARIFTYPMODE WHERE GFT_CODETYPE="'+CodeTarif+'"';
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then
  begin
    Tob_TarifTypMode := TOB.CREATE ('TARIFTYPMODE', NIL, -1);
    Tob_TarifTypMode.SelectDB('',Q);
    Ferme(Q) ;

    // on bascule tous les champs récupérés dans Tob_Info dans Tob_TarifTypMode
    for i:=0 to Tob_RepriseGPAO.detail.count-1 do
    begin
      Tob_Fille := Tob_RepriseGPAO.Detail [i];
      NomChamp  := Tob_Fille.GetValue ('GRC_NOMCHAMP');

      // Création des éventuels champs supplémentaires
      if not Tob_TarifTypMode.FieldExists (NomChamp) then
      begin
        Tob_TarifTypMode.AddChampSup (NomChamp, False);
        Tob_TarifTypMode.PutValue (NomChamp, '');
      end;

      Tob_TarifTypMode.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
    end;
    TobInfo.Dupliquer (Tob_TarifTypMode, False, True, True);
    FreeAndNil( Tob_TarifTypMode );
  end else Ferme (Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/09/2001
Modifié le ... :   /  /
Description .. : Fonction spécifique pour la reprise des devises
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MAJSpecifTarifPer (var TobInfo : TOB) ;
var Q            : TQUERY  ;
    Tob_TarifPer : TOB     ;
    Tob_Fille    : TOB     ;
    SQL          : string  ;
    CodePeriode  : string  ;
    NomChamp     : string  ;
    DateDebut    : TDateTime  ;
    DateFin      : TDateTime  ;
    i            : integer ;
begin
  // MAJ partielle de la devise dans PGI pour ne pas écraser des informations comprométantes
  CodePeriode := TobInfo.GetValue ('GFP_CODEPERIODE');

  SQL:='Select * From TARIFPER WHERE GFP_CODEPERIODE="'+CodePeriode+'"';
  Q:=OpenSQL(SQL,True) ;
  if Not Q.EOF then
  begin
    Tob_TarifPer := TOB.CREATE ('TARIFPER', NIL, -1);
    Tob_TarifPer.SelectDB('',Q);
    Ferme(Q) ;

    DateDebut := Tob_TarifPer.GetValue('GFP_DATEDEBUT');
    DateFin   := Tob_TarifPer.GetValue('GFP_DATEFIN');
    // on bascule tous les champs récupérés dans Tob_Info dans Tob_TarifTypMode
    for i:=0 to Tob_RepriseGPAO.detail.count-1 do
    begin
      Tob_Fille := Tob_RepriseGPAO.Detail [i];
      NomChamp  := Tob_Fille.GetValue ('GRC_NOMCHAMP');

      // Création des éventuels champs supplémentaires
      if not Tob_TarifPer.FieldExists (NomChamp) then
      begin
        Tob_TarifPer.AddChampSup (NomChamp, False);
        Tob_TarifPer.PutValue (NomChamp, '');
      end;

      Tob_TarifPer.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
    end;
    if (DateDebut<>Tob_TarifPer.GetValue('GFP_DATEDEBUT')) or (DateFin<>Tob_TarifPer.GetValue('GFP_DATEFIN')) then
    begin
      // Si la période d'application est modifiée
      SQL:='UPDATE TARIFMODE SET GFM_DATEDEBUT="'+ UsDateTime(Tob_TarifPer.GetValue('GFP_DATEDEBUT')) +
           '", GFM_DATEFIN="'+ UsDateTime(Tob_TarifPer.GetValue('GFP_DATEFIN')) +
           '" WHERE GFM_PERTARIF="'+CodePeriode+'"';
      ExecuteSQL( SQL );
      SQL:='UPDATE TARIF SET GF_DATEDEBUT="'+ UsDateTime(Tob_TarifPer.GetValue('GFP_DATEDEBUT')) +
           '", GF_DATEFIN="'+ UsDateTime(Tob_TarifPer.GetValue('GFP_DATEFIN')) +
           '" WHERE GF_TARFMODE IN (SELECT GFM_TARFMODE FROM TARIFMODE WHERE GFM_PERTARIF="'+CodePeriode+'")';
      ExecuteSQL( SQL );
    end;
    TobInfo.Dupliquer (Tob_TarifPer, False, True, True);
    FreeAndNil( Tob_TarifPer );
  end else Ferme (Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique de la TOB Tiers
Suite ........ :
Suite ........ :
Mots clefs ... : IMPORT TIERS
*****************************************************************}
procedure TTImportGPAO.MAJSpecifTiers (var TobInfo : TOB) ;
var i : integer ;
begin
  for i:=1 to 10 do
  begin
    if not TobInfo.FieldExists ('$$_TABLELIBRETIERS'+IntToStr(i)) then
    begin
      TobInfo.AddChampSup ('$$_TABLELIBRETIERS'+IntToStr(i), False);
      TobInfo.PutValue    ('$$_TABLELIBRETIERS'+IntToStr(i), '');
    end;
  end;
  for i:=1 to 3 do
  begin
    if not TobInfo.FieldExists ('$$_TEXTELIBRE'+IntToStr(i)) then
    begin
      TobInfo.AddChampSup ('$$_TEXTELIBRE'+IntToStr(i), False);
      TobInfo.PutValue    ('$$_TEXTELIBRE'+IntToStr(i), '');
    end;
    if not TobInfo.FieldExists ('$$_VALLIBRE'  +IntToStr(i)) then
    begin
      TobInfo.AddChampSup ('$$_VALLIBRE'+IntToStr(i), False);
      TobInfo.PutValue    ('$$_VALLIBRE'+IntToStr(i), 0);
    end;
    if not TobInfo.FieldExists ('$$_DATELIBRE' +IntToStr(i)) then
    begin
      TobInfo.AddChampSup ('$$_DATELIBRE'+IntToStr(i), False);
      TobInfo.PutValue    ('$$_DATELIBRE'+IntToStr(i), '01011900');
    end;
    if not TobInfo.FieldExists ('$$_BOOLLIBRE' +IntToStr(i)) then
    begin
      TobInfo.AddChampSup ('$$_BOOLLIBRE'+IntToStr(i), False);
      TobInfo.PutValue    ('$$_BOOLLIBRE'+IntToStr(i), '-');
    end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique des TOB à importer
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.MAJSpecifInfo (NomTable : string; var TobInfo : TOB) ;
begin
  if      (NomTable = 'DIMENSION'   ) then MAJSpecifDimension   (TobInfo)
  else if (NomTable = 'ARTICLE'     ) then MAJSpecifArticle     (TobInfo)
  else if (NomTable = 'DISPO'       ) then MAJSpecifDispo       (TobInfo)
  else if (NomTable = 'TARIF'       ) then MAJSpecifTarif       (TobInfo)
  else if (NomTable = 'DEVISE'      ) then MAJSpecifDevise      (TobInfo)
  else if (NomTable = 'TARIFTYPMODE') then MAJSpecifTarifTypMode(TobInfo)
  else if (NomTable = 'TARIFPER'    ) then MAJSpecifTarifPer    (TobInfo)
  else if (NomTable = 'TIERS'       ) then MAJSpecifTiers       (TobInfo)
  else if (NomTable = 'TRANSINVLIG' ) then MAJSpecifInventaire  (TobInfo)
  else if (NomTable = 'FIDELITELIG' ) then MAJSpecifFidelite    (TobInfo);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB CHOIXCOD avant import
Mots clefs ... : IMPORT CHOIXCOD
*****************************************************************}
function TTImportGPAO.CtrlTOBChoixCod (var TobInfo : TOB) : boolean;
begin
  result := True;

  if trim( TobInfo.GetValue ('CC_TYPE')) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE CHOIXCOD : le champ CC_TYPE doit être renseigné !'), True, True, True);
  end;

  if trim ( TobInfo.GetValue ('CC_CODE')) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE CHOIXCOD : le champ CC_CODE doit être renseigné !'), True, True, True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB CHOIXEXT avant import
Mots clefs ... : IMPORT CHOIXEXT
*****************************************************************}
function TTImportGPAO.CtrlTOBChoixExt (var TobInfo : TOB) : boolean;
begin
  result := True;

  if trim( TobInfo.GetValue ('YX_TYPE')) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE CHOIXEXT : le champ YX_TYPE doit être renseigné !'), True, True, True);
  end;

  if trim( TobInfo.GetValue ('YX_CODE')) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE CHOIXEXT : le champ YX_CODE doit être renseigné !'), True, True, True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB DIMENSION avant import
Mots clefs ... : IMPORT DIMENSION
*****************************************************************}
function TTImportGPAO.CtrlTOBDimension (var TobInfo : TOB) : boolean;
var Q       : TQUERY ;
    Test    : string ;
    TypeDim : string ;
    StType  : string ;
begin
  result := True;

  // Type de dimension : de 1 à 5
  Typedim := TobInfo.GetValue ('GDI_TYPEDIM');
  if trim (Typedim) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE DIMENSION : le type de dimension n''est pas renseigné !'), True, True, True);
  end;

  if (Typedim <> 'DI1')  and (Typedim <> 'DI2') and (Typedim <> 'DI3') and (Typedim <> 'DI4') and (Typedim <> 'DI5') then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE DIMENSION : le type de dimension doit être : DI1, DI2, DI3, DI4 ou DI5!'), True, True, True);
  end;

  // La taille doit appartenir à une grille
  test := TobInfo.GetValue ('GDI_GRILLEDIM');
  if  Trim( test) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE DIMENSION : la grille de dimension n''est pas renseignée'), True, True, True);
  end else
  begin
    if TypeDim = 'DI1' then StType := 'GG1' else
    if TypeDim = 'DI2' then StType := 'GG2' else
    if TypeDim = 'DI3' then StType := 'GG3' else
    if TypeDim = 'DI4' then StType := 'GG4' else
    if TypeDim = 'DI5' then StType := 'GG5';

    Q := OpenSQL('Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+StType+'" AND CC_CODE="'+Test+'"',True) ;
    if Q.EOF then
    begin
      EcrireInfo (TraduireMemoire('  TABLE DIMENSION : la grille de dimension "') + Test + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False ;
    end;
    Ferme(Q);
  end;

  // La taille doit avoir un code interne
  Test := TobInfo.GetValue ('GDI_CODEDIM');
  if (Trim(Test) = '') or (Test = '000') then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE DIMENSION : le code interne "') + Test + TraduireMemoire('" est incorrect'), True, True, True);
  end;

  // Le libellé doit être renseigné
  if Trim( TobInfo.GetValue ('GDI_LIBELLE')) = '' then
  begin
    result := False ;
    EcrireInfo (TraduireMemoire('  TABLE DIMENSION : le libellé de la taille n''est pas renseigné'), True, True, True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Contrôle de la TOB ARTICLE avant import
Mots clefs ... : IMPORT ARTICLE
*****************************************************************}
function TTImportGPAO.CtrlTOBArticle (var TobInfo : TOB) : boolean;
var TobPourCreer     : TOB    ;
    Tob_ArtCtrl      : TOB    ;
    Tob_ArtCompl     : TOB    ;
    Tob_ArtComplGen  : TOB    ;
    CodeArticle      : string ;
    CodeArticleGen   : string ;
    TypeChoix        : string ;
    TypeDim          : string ;
    GrilleDim        : string ;
    CodeDim          : string ;
    StRech	     : string ;
    Zone, LibZone    : string ;
    SQL              : string ;
    MessArticle      : string ;
    CollectionBase   : string ;
    Q                : TQUERY ;
    i		     : integer;
    ModifArticleGen  : boolean;
    CreeArticleCompl : boolean;
    IdentArticle     : string ;
    LaBellePhoto     : string ;
    TobPhoto         : TOB    ;
    PathPhoto        : string ;
    NomPhoto         : string ;
begin
  result      := True;
  MessArticle := TraduireMemoire('  ARTICLE ');

  // Comment identifier l'article ?
  if TobInfo.GetValue ('GA_CODEBARRE') <> '' then IdentArticle := TobInfo.GetValue ('GA_CODEBARRE')
  else                                            IdentArticle := TobInfo.GetValue ('GA_CODEARTICLE') ;
  /////////////////////////////////////////////////////////////////
  // Contrôle 1	:	Fournisseur principal
  /////////////////////////////////////////////////////////////////
  StRech :=  TobInfo.GetValue ('GA_FOURNPRINC');
  if (StRech <> '') and (StRech <> DernierFournisseur) then
  begin
    SQL:='Select T_LIBELLE From TIERS WHERE T_TIERS="'+StRech+'" AND T_NATUREAUXI="FOU"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le fournisseur "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierFournisseur := StRech;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 2	:	Collection
  /////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GA_COLLECTION');
  if (StRech <> '') and (StRech <> DerniereCollection) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GCO" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      Ferme (Q);
      // Peut on créer la collection ?
      if TobInfo.FieldExists ('$$_LIBCOLLECTION') then
      begin
        //
        // Si libellé récupéré et non renseigné, on y met le code
        //
        if TobInfo.GetValue ('$$_LIBCOLLECTION')='' then TobInfo.PutValue ('$$_LIBCOLLECTION', StRech);
        TobPourCreer := Tob.create ('CHOIXCOD', nil, -1);
        TobPourCreer.PutValue      ('CC_TYPE'   , 'GCO');
        TobPourCreer.PutValue      ('CC_CODE'   , StRech);
        TobPourCreer.PutValue      ('CC_LIBELLE', TobInfo.GetValue ('$$_LIBCOLLECTION'));
        TobPourCreer.PutValue      ('CC_ABREGE' , copy (TobInfo.GetValue ('$$_LIBCOLLECTION'), 1, 17));
        // Insertion dans DB
        TobPourCreer.InsertDB (nil);
        FreeAndNil( TobPourCreer );
        DerniereCollection := StRech ;
      end else
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : la collection "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
        result := False;
      end;
    end else
    begin
      DerniereCollection := StRech ;
      Ferme(Q);
    end;
  end;

  /////////////////////////////////////////////////////////////////
  // Contrôle 3 :	Les familles
  /////////////////////////////////////////////////////////////////
  for i:=1 to 8 do
  begin
    if i <= 3 then StRech := TobInfo.GetValue ('GA_FAMILLENIV'+IntToStr(i))
    else begin
      if TobInfo.Fieldexists ('$$_FAMILLENIV'+IntToStr(i)) then StRech := TobInfo.GetValue ('$$_FAMILLENIV'+IntToStr(i))
      else StRech := '';
    end;
    if (StRech <> '') and (StRech <> DerniereFamille [i]) then
    begin
      TypeChoix := 'FN' + IntToStr (i);
      SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+TypeChoix+'" AND CC_CODE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        Ferme (Q);
        // Peut on créer la famille ?
        if TobInfo.FieldExists ('$$_LIBFAMILLE'+IntToStr(i)) then
        begin
          // Si libellé récupéré et non renseigné, on y met le code
          if TobInfo.GetValue ('$$_LIBFAMILLE'+IntToStr(i))='' then TobInfo.PutValue ('$$_LIBFAMILLE'+IntToStr(i), StRech);
          TobPourCreer := Tob.create ('CHOIXCOD', nil, -1);
          TobPourCreer.PutValue      ('CC_TYPE'   , TypeChoix);
          TobPourCreer.PutValue      ('CC_CODE'   , StRech);
          TobPourCreer.PutValue      ('CC_LIBELLE', TobInfo.GetValue ('$$_LIBFAMILLE'+IntToStr(i)));
          TobPourCreer.PutValue      ('CC_ABREGE' , copy (TobInfo.GetValue ('$$_LIBFAMILLE'+IntToStr(i)), 1, 17));
          // Insertion dans DB
          TobPourCreer.InsertDB (nil);
          FreeAndNil( TobPourCreer );
          DerniereFamille [i] := StRech ;
        end else
        begin
          EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : la famille de niveau ') + IntToStr(i) + ' "' + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
          result := False;
        end;
      end else
      begin
        DerniereFamille [i] := StRech ;
        Ferme (Q);
      end;
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôles 4	:	Zones tables libres 1 à 10
  /////////////////////////////////////////////////////////////////
  for i:=1 to 10 do
  begin
    if (i < 10) then
    begin
      Zone      := 'GA_LIBREART' + IntToStr(i);
      LibZone   := '$$_LIBLIBREART' + IntToStr(i);
      TypeChoix := 'LA' + IntToStr (i);
    end
    else begin
      Zone      := 'GA_LIBREARTA';
      LibZone   := '$$_LIBLIBREARTA';
      TypeChoix := 'LAA';
    end;

    StRech := TobInfo.GetValue (Zone);
    if (StRech <> '') and (StRech <> DerniereStatArt [i]) then
    begin
      SQL:='Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="'+TypeChoix+'" AND YX_CODE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        Ferme (Q);
        // Peut on créer la stat ?
        if TobInfo.FieldExists (LibZone) then
        begin
          //
          // Si libellé récupéré et non renseigné, on y met le code
          //
          if TobInfo.GetValue (LibZone)='' then TobInfo.PutValue (LibZone, StRech);

          TobPourCreer := Tob.create ('CHOIXEXT', nil, -1);
          TobPourCreer.PutValue      ('YX_TYPE'   , TypeChoix);
          TobPourCreer.PutValue      ('YX_CODE'   , StRech);
          TobPourCreer.PutValue      ('YX_LIBELLE', TobInfo.GetValue (LibZone));
          TobPourCreer.PutValue      ('YX_ABREGE' , copy (TobInfo.GetValue (LibZone), 1, 17));
          // Insertion dans DB
          TobPourCreer.InsertDB (nil);
          FreeAndNil( TobPourCreer );
          DerniereStatArt [i] := StRech ;
        end else
        begin
          EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : la statistique libre "') + StRech + TraduireMemoire('" n''existe pas en table ') + IntToStr(i), True, True, True);
          result := False;
        end;
      end else
      begin
        DerniereStatArt [i] := StRech ;
        Ferme(Q);
      end;
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôles 4b	:	Zones statistiques complémentaires
  /////////////////////////////////////////////////////////////////
  for i:=1 to 2 do
  begin
    Zone      := '$$_STATART' + IntToStr(i);
    StRech := TobInfo.GetValue (Zone);
    if (StRech <> '') and (StRech <> DerniereStartCompl [i]) then
    begin
      LibZone := '$$_LIBSTATART' + IntToStr(i);
      if i = 1 then TypeChoix := 'FNA'
      else          TypeChoix := 'FNB';

      SQL:='Select YX_LIBELLE From CHOIXEXT WHERE YX_TYPE="'+TypeChoix+'" AND YX_CODE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        Ferme (Q);
        // Peut on créer la stat ?
        if TobInfo.FieldExists (LibZone) then
        begin
          //
          // Le libellé est il renseigné ?
          //
          if TobInfo.GetValue (libzone) = '' then TobInfo.PutValue (libzone, StRech);

          TobPourCreer := Tob.create ('CHOIXEXT', nil, -1);
          TobPourCreer.PutValue      ('YX_TYPE'   , TypeChoix);
          TobPourCreer.PutValue      ('YX_CODE'   , StRech);
          TobPourCreer.PutValue      ('YX_LIBELLE', TobInfo.GetValue (libzone) );
          TobPourCreer.PutValue      ('YX_ABREGE' , copy (TobInfo.GetValue (libzone), 1, 17) );
          // Insertion dans DB
          TobPourCreer.InsertDB (nil);
          FreeAndNil( TobPourCreer );
          DerniereStartCompl [i] := StRech ;
        end else
        begin
          EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le statistique libre "') + StRech + TraduireMemoire('" n''existe pas en table ') + IntToStr(i), True, True, True);
          result := False;
        end;
      end else
      begin
        DerniereStartCompl [i] := StRech ;
        Ferme(Q);
      end;
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 5	:	Les familles de taxe
  /////////////////////////////////////////////////////////////////
  for i:=1 to 2 do
  begin
    StRech := TobInfo.GetValue ('GA_FAMILLETAXE'+IntToStr (i));
    if (StRech <> '') and (StRech <> DerniereTaxe [i]) then
    begin
      TypeChoix := 'TX' + IntToStr (i);
      SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="'+TypeChoix+'" AND CC_CODE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : la famille de taxe ') + IntToStr(i) +  '"' + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
        result := False;
      end else DerniereTaxe [i] := StRech ;
      Ferme(Q);
    end;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 6 :	Tarif Article
  /////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GA_TARIFARTICLE');
  if (StRech <> '') and (StRech <> DernierTypeTarif) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TAR" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le code tarif "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierTypeTarif := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 7	:	Statut
  /////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GA_STATUTART');
  if (StRech <> '') then
  begin
    if (StRech <> DernierStatut) then
    begin
      SQL:='Select CO_LIBELLE From COMMUN WHERE CO_TYPE="GSA" AND CO_CODE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le statut article "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
        result := False;
      end else DernierStatut := StRech ;
      Ferme(Q);
    end;
  end else
  begin
    EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le statut article n''est pas renseigné'), True, True, True);
    result := False;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 8	:	Type article
  /////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GA_TYPEARTICLE');
  if (StRech <> '') and (StRech <> DernierTypeArticle) then
  begin
    SQL:='Select CO_LIBELLE From COMMUN WHERE CO_TYPE="TYA" AND CO_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le type d''article "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierTypeArticle := StRech;
    Ferme(Q);
  end;
  ///////////////////////////////////////////////////////////////////////////////
  // Contrôle 9	:	Gestion du masque de dimensions et des grilles de dimensions
  ///////////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GA_DIMMASQUE');
  if (StRech ='' ) and ((TobInfo.GetValue ('GA_STATUTART') = 'DIM') or (TobInfo.GetValue ('GA_STATUTART') = 'GEN')) then
  begin
    EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le masque de dimension n''est pas renseigné'), True, True, True);
    result := False;
  end;
  /////////////////////////////////////////////////////////////////
  // Contrôle 10 : contrôle de la photo transmise
  /////////////////////////////////////////////////////////////////
  if ((TobInfo.FieldExists ('$$_CHEMINPHOTO')) and (TobInfo.GetValue ('$$_CHEMINPHOTO') <> '')) then
  begin
    LaBellePhoto := TobInfo.GetValue ('$$_CHEMINPHOTO') ;
    if not FileExists(LaBellePhoto) then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : Impossible d''accéder à la photo ') + LaBellePhoto, True, True, True);
      result := True ;  // MODIF LM 14/05/03 pour LACOSTE
      TobInfo.PutValue ('$$_CHEMINPHOTO', '');
    end ;
    if ((result) and (TobInfo.GetValue ('$$_CHEMINPHOTO') <> '')) then
    begin
      StRech :=  ExtractFileExt (LaBellePhoto) ;
      if UpperCase (StRech) <> '.JPG'  then
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : Seules les photos au format JPEG sont reprises !') , True, True, True);
        result := False;
      end;
    end;
  end;
  ////////////////////////////////////////////////////////////////////////
  // Contrôle des codes internes
  ////////////////////////////////////////////////////////////////////////
  CodeArticle := TobInfo.GetValue ('GA_ARTICLE');
  for i:=1 to 5 do
  begin
    GrilleDim := TobInfo.GetValue ('GA_GRILLEDIM'+IntToStr(i));
    CodeDim   := TobInfo.GetValue ('GA_CODEDIM'+IntToStr(i));

    if (GrilleDim <> '') and (CodeDim = '') then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : la grille ou la dimension ') + IntToStr(i) + TraduireMemoire(' n''existe pas'), True, True, True);
      result := False;
    end;

    if (GrilleDim <> '') and (CodeDim <> '') then
    begin
      TypeDim := 'DI' + IntToStr(i);
      SQL:='Select * From DIMENSION WHERE GDI_TYPEDIM="'+TypeDim+'" AND GDI_GRILLEDIM="'+GrilleDim+'" AND GDI_CODEDIM="'+CodeDim+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : le code interne "') + CodeDim + TraduireMemoire('" n''existe pas dans la grille ') + GrilleDim + TraduireMemoire(' de la dimension ') + IntToStr(i), True, True, True);
        result := False;
      end;
      Ferme (Q);
    end;
  end;
  //////////////////////////////////////////////////////////////////////////////
  // Contrôle de cohérence des grilles de tailles et du masque
  //////////////////////////////////////////////////////////////////////////////
  CodeArticle := TobInfo.GetValue ('GA_ARTICLE');
  SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+CodeArticle+'"';
  Q:=OpenSQL(SQL,True) ;
  if not Q.EOF then
  begin
    Tob_ArtCtrl := TOB.CREATE ('ARTICLE', NIL, -1);
    Tob_ArtCtrl.SelectDB('',Q);

    for i:=1 to 5 do
    begin
      if TobInfo.GetValue ('GA_GRILLEDIM'+IntToStr(i)) <> Tob_ArtCtrl.GetValue ('GA_GRILLEDIM'+IntToStr(i)) then
      begin
        EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : Vous ne pouvez pas changer la grille de dimansion ') + IntToStr(i), True, True, True);
        result := False;
      end;
    end;

    if TobInfo.GetValue ('GA_DIMMASQUE') <> Tob_ArtCtrl.GetValue ('GA_DIMMASQUE') then
    begin
      EcrireInfo (MessArticle + IdentArticle + TraduireMemoire(' : Vous ne pouvez pas changer le masque de dimansion '), True, True, True);
      result := False;
    end;
    FreeAndNil( Tob_ArtCtrl );
  end;
  Ferme(Q) ;

  //////////////////////////////////////////////////////////////////////////////
  // L'article générique existe ??????
  // Si tous les contrôles sont OK, et que l'article générique n'existe pas,
  // on le crée
  // S'il existe, on vérifie les éléments principaux pour une cohérence
  // Article Générique - Article Dimensionné
  //////////////////////////////////////////////////////////////////////////////
  if result = True then
  begin
    CodeArticleGen := TobInfo.GetValue ('GA_CODEARTICLE');
    ArticleCodeGen (CodeArticleGen);

    if (TobInfo.GetValue ('GA_STATUTART') = 'DIM') then
    begin
      if CodeArticleGen <> DernierArticleGen then
      begin
        SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+CodeArticleGen+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then
        begin
          Ferme (Q);
          CreateArticleGenerique (TobInfo);
          DernierArticleGen := CodeArticleGen ;
        end else
        begin
          //
          // Sauvegarde du dernier article générique contrôlé pour éviter de multiples contrôles
          //
          DernierArticleGen := CodeArticleGen ;

          Tob_ArtCtrl := TOB.CREATE ('ARTICLE', NIL, -1);
          Tob_ArtCtrl.initValeurs;
          Tob_ArtCtrl.SelectDB('',Q);
          Ferme (Q);

          ModifArticleGen := False ;
          //
          // Collection
          //
          if TobInfo.GetValue ('GA_COLLECTION') <> Tob_ArtCtrl.GetValue ('GA_COLLECTION') then
          begin
            Tob_ArtCtrl.PutValue ('GA_COLLECTION', TobInfo.GetValue ('GA_COLLECTION')) ;
            ModifArticleGen := True ;
          end;
          //
          // Famille
          //
          for i:=1 to 3 do
          begin
            if TobInfo.GetValue ('GA_FAMILLENIV'+IntToStr(i)) <> Tob_ArtCtrl.GetValue ('GA_FAMILLENIV'+IntToStr(i)) then
            begin
              Tob_ArtCtrl.PutValue ('GA_FAMILLENIV'+IntToStr(i), TobInfo.GetValue ('GA_FAMILLENIV'+IntToStr(i))) ;
              ModifArticleGen := True ;
            end;
          end;
          //
          // Statistiques
          //
          for i:=1 to 10 do
          begin
            if (i < 10) then Zone := 'GA_LIBREART' + IntToStr(i)
            else             Zone := 'GA_LIBREARTA';

            if TobInfo.GetValue (Zone) <> Tob_ArtCtrl.GetValue (Zone) then
            begin
              Tob_ArtCtrl.PutValue (Zone, TobInfo.GetValue (Zone)) ;
              ModifArticleGen := True ;
            end;
          end;
          //
          // Date libre
          //
          for i:=1 to 3 do
          begin
            Zone := 'GA_DATELIBRE' + IntToStr(i)  ;
            if TobInfo.GetValue (Zone) <> Tob_ArtCtrl.GetValue (Zone) then
            begin
              Tob_ArtCtrl.PutValue (Zone, TobInfo.GetValue (Zone)) ;
              ModifArticleGen := True ;
            end;
            Zone := 'GA_CHARLIBRE' + IntToStr(i)  ;
            if TobInfo.GetValue (Zone) <> Tob_ArtCtrl.GetValue (Zone) then
            begin
              Tob_ArtCtrl.PutValue (Zone, TobInfo.GetValue (Zone)) ;
              ModifArticleGen := True ;
            end;
            Zone := 'GA_BOOLLIBRE' + IntToStr(i)   ;
            if TobInfo.GetValue (Zone) <> Tob_ArtCtrl.GetValue (Zone) then
            begin
              Tob_ArtCtrl.PutValue (Zone, TobInfo.GetValue (Zone)) ;
              ModifArticleGen := True ;
            end;
          end;
          //
          // libellés ....
          //
          if TobInfo.GetValue ('GA_LIBELLE') <> Tob_ArtCtrl.GetValue ('GA_LIBELLE') then
          begin
            Tob_ArtCtrl.PutValue ('GA_LIBELLE', TobInfo.GetValue ('GA_LIBELLE')) ;
            ModifArticleGen := True ;
          end;
          if TobInfo.GetValue ('GA_LIBCOMPL') <> Tob_ArtCtrl.GetValue ('GA_LIBCOMPL') then
          begin
            Tob_ArtCtrl.PutValue ('GA_LIBCOMPL', TobInfo.GetValue ('GA_LIBCOMPL')) ;
            ModifArticleGen := True ;
          end;
          if TobInfo.GetValue ('GA_COMMENTAIRE') <> Tob_ArtCtrl.GetValue ('GA_COMMENTAIRE') then
          begin
            Tob_ArtCtrl.PutValue ('GA_COMMENTAIRE', TobInfo.GetValue ('GA_COMMENTAIRE')) ;
            ModifArticleGen := True ;
          end;
          //
          // Poids ....
          //
          if TobInfo.GetValue ('GA_POIDSNET') <> Tob_ArtCtrl.GetValue ('GA_POIDSNET') then
          begin
            Tob_ArtCtrl.PutValue ('GA_POIDSNET', TobInfo.GetValue ('GA_POIDSNET')) ;
            ModifArticleGen := True ;
          end;
          if TobInfo.GetValue ('GA_POIDSBRUT') <> Tob_ArtCtrl.GetValue ('GA_POIDSBRUT') then
          begin
            Tob_ArtCtrl.PutValue ('GA_POIDSBRUT', TobInfo.GetValue ('GA_POIDSBRUT')) ;
            ModifArticleGen := True ;
          end;

          //  DPR pour JACADI
          if TobInfo.GetValue ('GA_DPR') <> Tob_ArtCtrl.GetValue ('GA_DPR') then
          begin
            if (TobInfo.GetValue ('GA_DPR') <> 0) then
            begin
              Tob_ArtCtrl.PutValue ('GA_DPR', TobInfo.GetValue ('GA_DPR')) ;
              ModifArticleGen := True ;
            end;
          end;

          //  PRHT pour JACADI
          if TobInfo.GetValue ('GA_PRHT') <> Tob_ArtCtrl.GetValue ('GA_PRHT') then
          begin
            if (TobInfo.GetValue ('GA_PRHT') <> 0) then
            begin
              Tob_ArtCtrl.PutValue ('GA_PRHT', TobInfo.GetValue ('GA_PRHT')) ;
              ModifArticleGen := True ;
            end;
          end;

          // Mise à jour de la fiche article dimensionné
          if ModifArticleGen then
          begin
            Tob_ArtCtrl.SetAllModifie(TRUE) ;
            Tob_ArtCtrl.UpdateDB (True);
          end;

          // Libération de la TOB de contrôle
          FreeAndNil( Tob_ArtCtrl );
        end;
      end;
    end else
    begin
      DernierArticleGen := CodeArticleGen ;
    end;

    ////////////////////////////////////////////////////////
    // Faut-il y mettre une belle photo ?
    ////////////////////////////////////////////////////////
    if ((TobInfo.FieldExists ('$$_CHEMINPHOTO')) and (TobInfo.GetValue ('$$_CHEMINPHOTO') <> '')) then
    begin
      LaBellePhoto := TobInfo.GetValue ('$$_CHEMINPHOTO') ;
      PathPhoto    := ExtractFilePath  (LaBellePhoto) ;
      NomPhoto     := ExtractFileName  (LaBellePhoto) ;

      TobPhoto := AttacheLaPhoto(nil,DernierArticleGen,PathPhoto, NomPhoto,1);

      SQL:='Select LO_LIBELLE From LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="'+DernierArticleGen+'" AND LO_RANGBLOB=0';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        Ferme (Q);
        TobPhoto.InsertDB(nil);
      end else
      begin
        Ferme (Q);
        TobPhoto.UpdateDB(False);
      end;
    end;
  end;
  //////////////////////////////////////////////////////////////////////////////
  // Si tout va bien : création de l'enregistrement ARTICLE_COMPL
  //////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    CreeArticleCompl := False ;
    if TobInfo.GetValue ('$$_COLLECTIONBAS') <> '' then  CreeArticleCompl := True ;
    for i:=4 to 8 do
    begin
      if TobInfo.GetValue ('$$_FAMILLENIV'+IntToStr(i)) <> '' then  CreeArticleCompl := True ;
    end;
    for i:=1 to 2 do
    begin
      if TobInfo.GetValue ('$$_STATART'+IntToStr(i)) <> '' then  CreeArticleCompl := True ;
    end;

    if CreeArticleCompl then
    begin
      Tob_ArtCompl := TOB.CREATE ('ARTICLECOMPL', NIL, -1);
      Tob_ArtCompl.PutValue ('GA2_ARTICLE'       , TobInfo.GetValue ('GA_ARTICLE'));
      Tob_ArtCompl.PutValue ('GA2_CODEARTICLE'   , TobInfo.GetValue ('GA_CODEARTICLE'));
      Tob_ArtCompl.PutValue ('GA2_COLLECTIONBAS' , TobInfo.GetValue ('$$_COLLECTIONBAS'));
      for i:=4 to 8 do
      begin
        Tob_ArtCompl.PutValue ('GA2_FAMILLENIV'+IntToStr(i) , TobInfo.GetValue ('$$_FAMILLENIV'+IntToStr(i)));
      end;
      for i:=1 to 2 do
      begin
        Tob_ArtCompl.PutValue ('GA2_STATART'+IntToStr(i) , TobInfo.GetValue ('$$_STATART'+IntToStr(i)));
      end;
      //
      // MAJ DB
      //
//      Tob_ArtCompl.SetAllModifie(TRUE) ;
      Tob_ArtCompl.InsertOrUpdateDB (True);

      //
      // Création de l'enregistrement pour l'article générique
      //
      CodeArticleGen := TobInfo.GetValue ('GA_CODEARTICLE');
      ArticleCodeGen (CodeArticleGen);

      if CodeArticleGen <> DernierArticleGenCompl then
      begin
        //
        // Sauvegarde du dernier article générique contrôlé pour éviter de multiples contrôles
        //
        DernierArticleGenCompl := CodeArticleGen ;
        Tob_ArtComplGen := Tob.Create ('ARTICLE', nil, -1);
        Tob_ArtComplGen.dupliquer (Tob_ArtCompl, True, True, True);
        //
        // MAJ spécifique pour l'article générique
        //
        Tob_ArtComplGen.PutValue ('GA2_ARTICLE' , CodeArticleGen);
        //
        // Sauvegarde de la collection de base d'origine pour le générique
        //
        SQL:='Select GA2_COLLECTIONBAS from ARTICLECOMPL WHERE GA2_ARTICLE="'+CodeArticleGen+'"';
        Q:=OpenSQL(SQL,True) ;
        if not Q.EOF then
        begin
          CollectionBase := Q.FindField('GA2_COLLECTIONBAS').AsString ;
          if CollectionBase <> '' then  Tob_ArtComplGen.PutValue ('GA2_COLLECTIONBAS' , CollectionBase);
        end ;
        Ferme (Q);

//        Tob_ArtComplGen.SetAllModifie(TRUE) ;
        Tob_ArtComplGen.InsertOrUpdateDB (True);
        FreeAndNil( Tob_ArtComplGen );
      end;
      FreeAndNil( Tob_ArtCompl );
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB DISPO avant import
Mots clefs ... : IMPORT DISPO
*****************************************************************}
function TTImportGPAO.CtrlTOBDispo (var TobInfo : TOB) : boolean;
var StRech : string ;
    SQL    : string ;
    Statut : string ;
    Q      : TQUERY ;
begin
  result := True;

  /////////////////////////////////////////////////////////////////////////
  // Code dépôt
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GQ_DEPOT');
  if (StRech <> '') then
  begin
    SQL:='Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (TraduireMemoire('  TABLE DISPO : le dépôt "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end;
    Ferme(Q);
  end else
  begin
    EcrireInfo (TraduireMemoire('  TABLE DISPO : le dépôt n''est pas renseigné'), True, True, True);
    result := False;
  end;

  /////////////////////////////////////////////////////////////////////////////////////
  // Attention : Zone sensible
  // Suppresion de toutes les fiches stocks du dépôt
  /////////////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    if (TobInfo.FieldExists ('$$_RAZSTOCK')) and (TobInfo.GetValue ('$$_RAZSTOCK')='X') then
    begin
      // Suppression des stocks d'un dépôt
      StRech := TobInfo.GetValue ('GQ_DEPOT');
      EcrireInfo (TraduireMemoire('  ---> Suppression du stock du dépôt ') + StRech, True, True, True);
      ExecuteSQL ('DELETE FROM DISPO WHERE GQ_DEPOT="'+StRech+'" AND GQ_CLOTURE="-"');
      FreeAndNil( TobInfo );
    end else
    begin
      /////////////////////////////////////////////////////////////////////////
      // Code Article
      /////////////////////////////////////////////////////////////////////////
      StRech := TobInfo.GetValue ('GQ_ARTICLE');
      if (StRech <> '') then
      begin
        SQL:='Select GA_STATUTART From ARTICLE WHERE GA_ARTICLE="'+StRech+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then
        begin
          EcrireInfo (TraduireMemoire('  TABLE DISPO : l''article "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
          result := False;
        end else
        begin
          Statut := Q.FindField('GA_STATUTART').AsString ;
          if Statut = 'GEN' then
          begin
            EcrireInfo (TraduireMemoire('  TABLE DISPO : Article ') + StRech + TraduireMemoire (' : Impossible d''intégrer le stock d''un article générique'), True, True, True);
            result := False;
          end;
        end;
        Ferme(Q);
      end else
      begin
        EcrireInfo (TraduireMemoire('  TABLE DISPO : l''article n''est pas renseigné'), True, True, True);
        result := False;
      end;
    end;
  end;
  TobInfo.SetAllModifie (True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB TARIF avant import
Mots clefs ... : IMPORT TARIF
*****************************************************************}
function TTImportGPAO.CtrlTOBTarif (var TobInfo : TOB) : boolean;
var Q                         : TQUERY ;
    SQL                       : string ;
    StRech                    : string ;
    Code_Tarif                : string ;
    Periode_Tarif             : string ;
    CodeArticleTai            : string ;
    CodeArticleGen            : string ;
    CodeDepot                 : string ;
    CodeDevise                : string ;
    Chtmp                     : string ;
    TarfMode                  : integer;
    NumTarif                  : integer;
    PrixAncien                : double ;
    PrixDimension             : double ;
    PrixGenerique             : double ;
    MajTarifArticleDimension  : boolean;
    CreeTarifArticleDimension : boolean;
    CreeTarifArticleGenerique : boolean;
    PrixArticle               : double ;
    Devise                    : string ;
    IdentArticle              : string ;
    MessTarif                 : string ;
begin
  result     := True;
  Numtarif   := 0 ;
  PrixAncien := 0.0 ;
  MessTarif  := TraduireMemoire('  TARIF ');

  // Comment identifier l'article ?
  if (TobInfo.FieldExists ('$$_CODEBARRE')) and (TobInfo.GetValue ('$$_CODEBARRE') <> '') then IdentArticle := TobInfo.GetValue ('$$_CODEBARRE')
  else IdentArticle := TobInfo.GetValue ('$$_CODEARTICLE') ;
  /////////////////////////////////////////////////////////////////////////
  // Code article
  /////////////////////////////////////////////////////////////////////////
  if TobInfo.GetValue ('GF_ARTICLE') = '' then
  begin
    Chtmp := ReconstitueIdentArticle (TobInfo);
    if Chtmp <> '' then TobInfo.PutValue('GF_ARTICLE', Chtmp)
    else result := False ;
  end else
  begin
    StRech := TobInfo.GetValue ('GF_ARTICLE');
    if (StRech <> '') then
    begin
      SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        EcrireInfo (MessTarif + IdentArticle + TraduireMemoire(' : l''article "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
        result := False;
      end;
      Ferme(Q);
    end;
  end;

  /////////////////////////////////////////////////////////////////////////
  // Code dépôt
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GF_DEPOT');
  if (StRech <> '') and (StRech <> DernierDepot) then
  begin
    SQL:='Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTarif + IdentArticle + TraduireMemoire(' : le dépôt "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierDepot := StRech ;
    Ferme(Q);
  end;

  /////////////////////////////////////////////////////////////////////////
  // Code devise
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('GF_DEVISE');
  if (StRech <> '') and (StRech <> DerniereDevise) then
  begin
    SQL:='Select D_LIBELLE From DEVISE WHERE D_DEVISE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTarif + IdentArticle + TraduireMemoire(' : la devise "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DerniereDevise := StRech ;
    Ferme (Q);
  end;

  // Récupération de la période et du code tarif (obligatoires pour PV)
  if result then
  begin

    Periode_Tarif := TobInfo.GetValue ('$$_CODEPERIODE');
    Code_Tarif    := TobInfo.GetValue ('$$_CODETYPE');

    if (Periode_Tarif <> '') and (Code_Tarif <> '') then
    begin

      // Si on change de tarif, il faut recharger la TOB "Tob_Tarif"
      if (Periode_Tarif <> Periode) or (Code_Tarif <> CodeTarif) then
      begin
        // Charge la TOb et créé éventuellement le croisement Periode/Tarif dans TARIFMODE
        FreeAndNil( Tob_Tarif );
        Tob_Tarif := TOB.Create ('TARIFMODE', nil, -1);

        result := ChargeTobQualifiantTarif (Tob_Tarif, Code_Tarif, Periode_Tarif);

        if result then
        begin
          Periode   := Periode_Tarif;
          CodeTarif := Code_Tarif   ;
          TobInfo.PutValue ('GF_TARFMODE', Tob_Tarif.GetValue ('GFM_TARFMODE'));
        end;
      end else
      begin
        // MODIF LACOSTE
        // Ajout récup du tarfmode
        TobInfo.PutValue ('GF_TARFMODE', Tob_Tarif.GetValue ('GFM_TARFMODE'));
      end;

      // Pour l'instant, on ne met à jour que les tarifs PV
      if result then
      begin
        // Initialisation de la devise
        CodeDevise := TobInfo.GetValue ('GF_DEVISE');
        if CodeDevise = '' then
        begin
          CodeDevise := Tob_Tarif.GetValue ('GFM_DEVISE');
          TobInfo.PutValue ('GF_DEVISE', CodeDevise);
        end else
        begin
          if CodeDevise <> Tob_Tarif.GetValue ('GFM_DEVISE') then
          begin
            EcrireInfo (MessTarif  + IdentArticle + TraduireMemoire(' : Dépôt ') + TobInfo.GetValue ('GF_DEPOT') + TraduireMemoire(' -> Devise incompatible avec le tarif ') + Tob_Tarif.GetValue ('GFM_DEVISE'), True, True, True);
            result := False ;
          end;
        end;
      end;

      if result then
      begin
        PrixDimension := TobInfo.GetValue ('GF_PRIXUNITAIRE');
        //
        // Doit-on créer réellement le tarif ?
        //
        // Cas 1 - Si le tarif existe, on le met à jour
        // Cas 2 - Sinon
        //         Si le tarif pour l'article générique n'existe pas, on le crée
        //         et laisse tomber le tarif pour la taille. Le tarif appliqué pour
        //         l'article dimensionné sera celui de l'article générique
        //
        //          Si le tarif de l'article générique existe, et que le tarif de l'article
        //         dimensionné est différent, alors on créé la tarif de l'article dimensionné.
        //         S'il est identique à celui de l'article générique, on laisse tomber.
        //
        MajTarifArticleDimension  := False ;
        CreeTarifArticleDimension := False ;
        CreeTarifArticleGenerique := False ;

        CodeDepot      := TobInfo.GetValue ('GF_DEPOT');
        CodeArticleTai := TobInfo.GetValue ('GF_ARTICLE');
        CodeArticleGen := copy (CodeArticleTai, 1, 18) + '               X';
        TarfMode       := TobInfo.GetValue ('GF_TARFMODE');

        // Recherche du tarif à la taille
        SQL:='Select GF_TARIF, GF_PRIXANCIEN From TARIF WHERE GF_DEPOT="'+CodeDepot+'" AND GF_ARTICLE="'+CodeArticleTai+'" AND GF_TARFMODE="'+IntToStr(TarfMode)+'"';
        Q:=OpenSQL(SQL,True) ;
        if Not Q.EOF then
        begin
          PrixAncien := Q.FindField('GF_PRIXANCIEN').AsFloat;
          NumTarif   := Q.FindField('GF_TARIF').AsInteger;
          MajTarifArticleDimension := True ;
          Ferme (Q);
        end else
        begin
          Ferme (Q);

          // Recherche le tarif à l'article générique
          CodeDepot      := TobInfo.GetValue   ('GF_DEPOT');
          TarfMode       := Tob_Tarif.GetValue ('GFM_TARFMODE');

          SQL:='Select * From TARIF WHERE GF_DEPOT="'+CodeDepot+'" AND GF_ARTICLE="'+CodeArticleGen+'" AND GF_TARFMODE="'+IntToStr(TarfMode)+'"';
          Q:=OpenSQL(SQL,True) ;
          if Not Q.EOF then
          begin
            PrixGenerique := Q.FindField('GF_PRIXUNITAIRE').AsFloat;
            Ferme (Q);
            if PrixDimension <> PrixGenerique then CreeTarifArticleDimension := True ;
          end else
          begin
            Ferme (Q);
            CreeTarifArticleGenerique := True ;
          end;
        end;

        if (MajTarifArticleDimension) or (CreeTarifArticleDimension) or (CreeTarifArticleGenerique) then
        begin

          // Numéro de tarif
          if MajTarifArticleDimension then
          begin
            TobInfo.PutValue ('GF_TARIF'      , NumTarif);
            TobInfo.PutValue ('GF_PRIXANCIEN' , PrixAncien);
          end else
          begin
            RecupMaxTarif;
            Inc (MaxTarif);
            TobInfo.PutValue ('GF_TARIF'      , MaxTarif);
          end;

          // Code article : s'il faut créer un tarif pour l'article générique, il faut ré-initialiser la ref article
          if CreeTarifArticleGenerique then TobInfo.PutValue('GF_ARTICLE', CodeArticleGen) ;

          // S'il faut créer le tarif pour l'article dimensionné, il faut mettre à jour les champs 'GA_PRIXUNIQUE'
          // L'article générique et tous les articles dimensionnés sont flagés
          if CreeTarifArticleDimension then
          begin
            if CodeArticleGen <> DernierArticleFlagPxUnique then
            begin
              MiseAJourFlagPrixUnique (CodeArticleGen);
              DernierArticleFlagPxUnique := CodeArticleGen ;
            end;
          end;

          // Initialisations diverses et varièes
          TobInfo.PutValue ('GF_SOCIETE'  , TobInfo.GetValue ('GF_DEPOT'));
          TobInfo.PutValue ('GF_TARFMODE' , Tob_Tarif.GetValue ('GFM_TARFMODE'));
          TobInfo.PutValue ('GF_LIBELLE'  , Tob_Tarif.GetValue ('GFM_LIBELLE'));

          // Modif LM 05/06/01 : On peux désormais avoir une période de tarif par dépôt
          Q:=OpenSql('Select * from TARIFPER where GFP_CODEPERIODE="'+Periode_Tarif+'" AND GFP_ETABLISSEMENT="'+TobInfo.GetValue('GF_DEPOT')+'"',True) ;
	  if not Q.EOF then
          begin
            TobInfo.PutValue ('GF_DATEDEBUT' , Q.FindField ('GFP_DATEDEBUT').AsDateTime);
            TobInfo.PutValue ('GF_DATEFIN'   , Q.FindField ('GFP_DATEFIN').AsDateTime);
          end else
          begin
            TobInfo.PutValue ('GF_DATEDEBUT' , Tob_Tarif.GetValue ('GFM_DATEDEBUT'));
            TobInfo.PutValue ('GF_DATEFIN'   , Tob_Tarif.GetValue ('GFM_DATEFIN'));
          end;
          Ferme (Q);

          TobInfo.PutValue   ('GF_DEVISE'       , Tob_Tarif.GetValue ('GFM_DEVISE'));
          TobInfo.PutValue   ('GF_DEMARQUE'     , Tob_Tarif.GetValue ('GFM_DEMARQUE'));
          TobInfo.PutValue   ('GF_ARRONDI'      , Tob_Tarif.GetValue ('GFM_ARRONDI'));
          // Initialisation en "dur"
          TobInfo.PutValue   ('GF_QUALIFPRIX'   , 'GRP');
          TobInfo.PutValue   ('GF_CASCADEREMISE', 'MIE');
    	  TobInfo.PutValue   ('GF_MODECREATION' , 'MAN');
          TobInfo.PutValue   ('GF_BORNEINF'     , -999999);
          TobInfo.PutValue   ('GF_BORNESUP'     , 999999);
          TobInfo.PutValue   ('GF_QUANTITATIF'  , '-');
          // Calcul de la priorité
          CalculPriorite(TobInfo) ;
          if Tob_Tarif.GetValue ('GFM_NATURETYPE') = 'VTE' then
          begin
            TobInfo.PutValue ('GF_NATUREAUXI' , 'CLI') ;
            TobInfo.PutValue ('GF_REGIMEPRIX' , 'TTC') ;
          end else
          begin
            TobInfo.PutValue ('GF_NATUREAUXI' , 'FOU') ;
            TobInfo.PutValue ('GF_REGIMEPRIX' , 'HT' ) ;
          end;
        end;

        // MODIF LACOSTE : MAJ de la fiche article en mise à jour
        // Si problème lors du premier import , on ne met jamais à jour la
        // fiche article.
        // Faut -il mettre à jour le PV de la fiche article ?
        if Tob_Tarif.GetValue ('GFM_NATURETYPE') = 'VTE' then
        begin
          if (TobInfo.GetValue ('$$_PRIXARTICLE')='O') or ((Tob_Tarif.GetValue ('GFM_TYPETARIF')=GetParamSoc ('SO_GCTARPV')) and (Tob_Tarif.GetValue ('GFM_PERTARIF')=GetParamSoc ('SO_GCPERPV'))) then
          begin
            PrixArticle := TobInfo.GetValue ('GF_PRIXUNITAIRE');
            Devise      := TobInfo.GetValue ('GF_DEVISE');

            if CreeTarifArticleGenerique then
            begin
              MajPrixVenteToutArticle (CodeArticleGen, PrixArticle, Devise);
            end else
            begin
              MajPrixVenteArticle (CodeArticleTai, PrixArticle, Devise);
              if CodeArticleGen <> DernierArticleTar then      // MODIF LACOSTE
              begin
                MajPrixVenteArticle (CodeArticleGen, PrixArticle, Devise);
                DernierArticleTar := CodeArticleGen;         // MODIF LACOSTE
              end;
            end;
          end;
        end else if Tob_Tarif.GetValue ('GFM_NATURETYPE') = 'ACH' then
        begin
          if (TobInfo.GetValue ('$$_PRIXARTICLE')='O') or ((Tob_Tarif.GetValue ('GFM_TYPETARIF')=GetParamSoc ('SO_GCTARPA')) and (Tob_Tarif.GetValue ('GFM_PERTARIF')=GetParamSoc ('SO_GCPERPA'))) then
          begin
            //CodeArticleTai := TobInfo.GetValue ('GF_ARTICLE');
            PrixArticle    := TobInfo.GetValue ('GF_PRIXUNITAIRE');
            Devise         := TobInfo.GetValue ('GF_DEVISE');

            MajPrixAchatArticle (CodeArticleTai, PrixArticle, Devise);
          end;
        end;

        // Rien à faire. Le prix est le même que celui de l'article générique
        if (not MajTarifArticleDimension) and (not CreeTarifArticleDimension) and (not CreeTarifArticleGenerique) then
          FreeAndNil( TobInfo );
      end;
    end else
    begin
      EcrireInfo (MessTarif + IdentArticle + TraduireMemoire(' : Le code tarif et la période de tarif sont obligatoires ! '), True, True, True);
      result := False;
    end;
  end;
  ////////////////////////////////////////////////////////////
  // On force la date de modification
  //////////////////////////////////////////////////////////////
   if TobInfo <> nil then TobInfo.SetDateModif(NowH);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... : 15/12/2003
Description .. : Contrôle de la TOB PROSPECTS avant import
Mots clefs ... : IMPORT PROSPECTS 
*****************************************************************}
function TTImportGPAO.CtrlTOBProspects (var TobInfo : TOB) : boolean;
var
    SQL            : string    ;
    Auxiliaire     : string    ;
    NomChamp       : string    ;
    Q              : TQUERY    ;
    Tob_Prospects  : TOB       ;
    TobFilleParam  : TOB       ;
    i              : integer   ;
begin
  /////////////////////////////////////////////////////////////////////////
  // Faut-il mettre à jour partiellement la fiche prospects ?
  /////////////////////////////////////////////////////////////////////////
  if (TobInfo.FieldExists ('$$_MAJPARTIELLE')) and (TobInfo.GetValue ('$$_MAJPARTIELLE') ='X') then
  begin
    Auxiliaire := TobInfo.GetValue ('RPR_AUXILIAIRE');
    SQL := 'SELECT * FROM PROSPECTS WHERE RPR_AUXILIAIRE="'+Auxiliaire+'"' ;
    Q   := OpenSQL (SQL ,True) ;
    if not Q.EOF then
    begin
      Tob_Prospects := TOB.CREATE ('PROSPECTS', NIL, -1);
      Tob_Prospects.SelectDB('',Q);
      Ferme(Q) ;

      // on bascule tous les champs récupérés dans Tob_Info dans Tob_Prospects
      for i:=0 to Tob_RepriseGPAO.detail.count-1 do
      begin
        TobFilleParam := Tob_RepriseGPAO.Detail [i];
        NomChamp      := TobFilleParam.GetValue ('GRC_NOMCHAMP');
        //
        // Création des éventuels champs supplémentaires
        //
        if not Tob_Prospects.FieldExists (NomChamp) then
        begin
          Tob_Prospects.AddChampSup (NomChamp, False);
          Tob_Prospects.PutValue    (NomChamp, '');
        end;
        Tob_Prospects.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
      end;

      TobInfo.dupliquer (Tob_Prospects, False, True, True);
      FreeAndNil( Tob_Prospects );
    end;
  end;
  Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB TIERS avant import
Mots clefs ... : IMPORT TIERS
*****************************************************************}
function TTImportGPAO.CtrlTOBTiers (var TobInfo : TOB) : boolean;
var StRech         : string    ;
    SQL            : string    ;
    IdentTiers     : string    ;
    Q              : TQUERY    ;
    TomTiers       : TOM_TIERS ;
    CreeTiersCompl : boolean   ;
    CreeCliCash    : boolean   ;
    i              : integer   ;
    Tob_TiersCompl : TOB       ;
    Tob_Tiers      : TOB       ;
    TobFilleParam  : TOB       ;
    Auxiliaire     : string    ;
    NomChamp       : string    ;
    MessTiers      : string    ;
begin
  result    := True;
  MessTiers := TraduireMemoire('  TIERS ') ;
  ///////////////////////////////////////////////////////////////
  // Cas du client 999999 du CASH                 //NA 07/03/2002
  ///////////////////////////////////////////////////////////////
  CreeCliCash := False;
  if (TobInfo.FieldExists ('$$_CREECLICASH')) and (TobInfo.GetValue ('$$_CREECLICASH') = 'X') then
  begin
    IdentTiers := TobInfo.GetValue ('T_TIERS');
    if (IdentTiers = '999999') or (Copy(IdentTiers, 4, 6) = '999999') then
    begin
      TobInfo.PutValue ('T_TIERS', '');
      CreeCliCash := True;
    end;
  end ;

  if TobInfo.GetValue ('T_TIERS') = '' then IdentTiers := TobInfo.GetValue ('T_LIBELLE')
  else IdentTiers := TobInfo.GetValue ('T_TIERS');

  /////////////////////////////////////////////////////////////////////////
  // Pays
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_PAYS');
  if (StRech <> '') and (StRech <> DernierPays) then
  begin
    SQL:='Select PY_LIBELLE From PAYS WHERE PY_PAYS="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : le pays "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierPays := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Langue
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_LANGUE');
  if (StRech <> '') and (StRech <> DerniereLangue) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="LGU" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la langue "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DerniereLangue := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Devise
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_DEVISE');
  if (StRech <> '') and (StRech <> DerniereDevise) then
  begin
    SQL:='Select D_LIBELLE From DEVISE WHERE D_DEVISE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la devise "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      Result:=False;
    end else DerniereDevise := StRech ;
    Ferme (Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Zone commerciale
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_ZONECOM');
  if (StRech <> '') and (StRech <> DerniereZoneCom) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GZC" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la zone commerciale "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DerniereZoneCom := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Code tarif tiers
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_TARIFTIERS');
  if (StRech <> '') and (StRech <> DernierTarifTiers) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="TRC" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : le code tairf tiers "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierTarifTiers := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Mode de règlement
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_MODEREGLE');
  if (StRech <> '') and (StRech <> DernierModeRegle) then
  begin
    SQL:='Select MR_LIBELLE From MODEREGL WHERE MR_MODEREGLE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : le mode règlement "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierModeRegle := StRech ;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Régime TVA
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_REGIMETVA');
  if (StRech <> '') and (StRech <> DernierRegimeTVA) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="RTV" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : le régime de TVA "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierRegimeTVA := StRech;
    Ferme(Q);
  end;
  /////////////////////////////////////////////////////////////////////////
  // Civilité
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_JURIDIQUE');
  if (StRech <> '') and (TobInfo.GetValue ('T_PARTICULIER')='X') and (StRech <> DernierJuridique) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="CIV" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la civilité "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierJuridique := StRech;
    Ferme(Q);
  end;
  ////////////////////////////////////////////////////////////////////////
  // Nationalité
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_NATIONALITE');
  if (StRech <> '') and (StRech <> DerniereNationalite) then
  begin
    SQL:='Select PY_LIBELLE From PAYS WHERE PY_PAYS="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la nationalité "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else  DerniereNationalite := StRech ;
    Ferme(Q);
  end;
  ////////////////////////////////////////////////////////////////////////
  // Famille Comptable
  /////////////////////////////////////////////////////////////////////////
  StRech := TobInfo.GetValue ('T_COMPTATIERS');
  if (StRech <> '') and (StRech <> DernierComptaTiers) then
  begin
    SQL:='Select CC_LIBELLE From CHOIXCOD WHERE CC_TYPE="GCT" AND CC_CODE="'+StRech+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessTiers + IdentTiers + TraduireMemoire(' : la famille comptable "') + StRech + TraduireMemoire('" n''existe pas'), True, True, True);
      result := False;
    end else DernierComptaTiers:= StRech;
    Ferme(Q);
  end;

  ////////////////////////////////////////////////////////////////////////
  // Appel de la TOM Tiers
  /////////////////////////////////////////////////////////////////////////
{$IFDEF AGL560}
    TomTiers:=TOM_TIERS(CreateTOM('TIERS',Nil,False,False)) ;
{$ELSE}
    TomTiers:=TOM_TIERS.Create(Nil,Nil,FALSE) ;
{$ENDIF}

  if not TomTiers.VerifTOB(TobInfo) then
  begin
    EcrireInfo (MessTiers + IdentTiers + ' : ' + TomTiers.LastErrorMsg, True, True, True);
    result := False ;
  end;
  FreeAndNil( TomTiers );

  /////////////////////////////////////////////////////////////////////////
  // Faut-il mettre à jour partiellement la fiche tiers ?
  /////////////////////////////////////////////////////////////////////////
  if ((TobInfo.FieldExists ('$$_MAJPARTIELLE')) and (TobInfo.GetValue ('$$_MAJPARTIELLE') ='X') or
      (TobInfo.FieldExists ('$$_CREECLICASH')) and (TobInfo.GetValue ('$$_CREECLICASH') = 'X')) then
  begin
    Auxiliaire := TobInfo.GetValue ('T_AUXILIAIRE');
    SQL := 'SELECT * FROM TIERS WHERE T_AUXILIAIRE="'+Auxiliaire+'"' ;
    Q   := OpenSQL (SQL ,True) ;
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      Tob_Tiers.SelectDB('',Q);
      Ferme(Q) ;

      // Ajout des champs supplémentaires
      MAJSpecifTiers (Tob_Tiers);

      // Y a t il des stat ?
      SQL := 'SELECT * FROM TIERSCOMPL WHERE YTC_AUXILIAIRE="'+Auxiliaire+'"' ;
      Q   := OpenSQL (SQL ,True) ;
      if not Q.EOF then
      begin
        Tob_TiersCompl := TOB.CREATE ('TIERSCOMPL', NIL, -1);
        Tob_TiersCompl.SelectDB('',Q);
        Ferme(Q) ;

        for i:=1 to 10 do
        begin
          if i<10 then
          begin
            if Tob_TiersCompl.GetValue ('YTC_TABLELIBRETIERS'+IntToStr(i)) <> '' then
              Tob_Tiers.PutValue ('$$_TABLELIBRETIERS'+IntToStr(i), Tob_TiersCompl.GetValue ('YTC_TABLELIBRETIERS'+IntToStr(i)));
          end else
          begin
            if Tob_TiersCompl.GetValue ('YTC_TABLELIBRETIERSA') <> '' then
              Tob_Tiers.PutValue ('$$_TABLELIBRETIERS'+IntToStr(i), Tob_TiersCompl.GetValue ('YTC_TABLELIBRETIERSA'));
          end;
        end;
        for i:=1 to 3 do
        begin
          if Tob_TiersCompl.GetValue ('YTC_TEXTELIBRE'+IntToStr(i)) <> '' then
            Tob_Tiers.PutValue ('$$_TEXTELIBRE'+IntToStr(i), Tob_TiersCompl.GetValue ('YTC_TEXTELIBRE'+IntToStr(i)));
          if Tob_TiersCompl.GetValue ('YTC_VALLIBRE'  +IntToStr(i)) <> 0 then
            Tob_Tiers.PutValue ('$$_VALLIBRE'  +IntToStr(i), Tob_TiersCompl.GetValue ('YTC_VALLIBRE'  +IntToStr(i)));
          if IsValidDate(Tob_TiersCompl.GetValue ('YTC_DATELIBRE' +IntToStr(i))) then
            Tob_Tiers.PutValue ('$$_DATELIBRE' +IntToStr(i), Tob_TiersCompl.GetValue ('YTC_DATELIBRE' +IntToStr(i)));
          if Tob_TiersCompl.GetValue ('YTC_BOOLLIBRE' +IntToStr(i)) <> '' then
            Tob_Tiers.PutValue ('$$_BOOLLIBRE' +IntToStr(i), Tob_TiersCompl.GetValue ('YTC_BOOLLIBRE' +IntToStr(i)));
        end;
        FreeAndNil( Tob_TiersCompl );
      end;

      // on bascule tous les champs récupérés dans Tob_Info dans Tob_Article
      for i:=0 to Tob_RepriseGPAO.detail.count-1 do
      begin
        TobFilleParam := Tob_RepriseGPAO.Detail [i];
        NomChamp      := TobFilleParam.GetValue ('GRC_NOMCHAMP');
        //
        // Création des éventuels champs supplémentaires
        //
        if not Tob_Tiers.FieldExists (NomChamp) then
        begin
          Tob_Tiers.AddChampSup (NomChamp, False);
          Tob_Tiers.PutValue    (NomChamp, '');
        end;
        Tob_Tiers.PutValue (NomChamp, TobInfo.GetValue (NomChamp));
      end;

      TobInfo.dupliquer (Tob_Tiers, False, True, True);
      FreeAndNil( Tob_Tiers );
    end else
    begin
      ///////////////////////////////////////////////////////////////
      // Cas du client 999999 du CASH
      ///////////////////////////////////////////////////////////////
      if CreeCliCash then DernierClientCree := TobInfo.GetValue ('T_TIERS');
      Ferme (Q);
    end;
  end;
  //////////////////////////////////////////////////////////////////////////////
  // Gestion de la date de création et la date de fermeture
  //////////////////////////////////////////////////////////////////////////////
  if DateToStr (TobInfo.GetValue ('T_DATECREATION')) = '01/01/1900' then
    TobInfo.PutValue ('T_DATECREATION', Now);
  if DateToStr (TobInfo.GetValue ('T_DATEOUVERTURE')) = '01/01/1900' then
    TobInfo.PutValue ('T_DATEOUVERTURE', Now);

  //////////////////////////////////////////////////////////////////////////////
  // Si tout va bien : Création éventuel de l'enregistrement TIERSCOMPL
  //////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    CreeTiersCompl := False ;
    for i:=1 to 10 do
    begin
      if TobInfo.GetValue ('$$_TABLELIBRETIERS'+IntToStr(i)) <> '' then  CreeTiersCompl := True ;
    end;
    for i:=1 to 3 do
    begin
      if TobInfo.GetValue ('$$_TEXTELIBRE'+IntToStr(i)) <> ''        then  CreeTiersCompl := True ;
      if TobInfo.GetValue ('$$_VALLIBRE'  +IntToStr(i)) <> 0         then  CreeTiersCompl := True ;
      if IsValidDate(TobInfo.GetValue ('$$_DATELIBRE' +IntToStr(i))) then  CreeTiersCompl := True ;
      if TobInfo.GetValue ('$$_BOOLLIBRE' +IntToStr(i)) <> ''        then  CreeTiersCompl := True ;
    end;

    if CreeTiersCompl then
    begin
      Tob_TiersCompl := TOB.CREATE ('TIERSCOMPL', NIL, -1);
      Tob_TiersCompl.PutValue ('YTC_AUXILIAIRE' , TobInfo.GetValue ('T_AUXILIAIRE'));
      Tob_TiersCompl.PutValue ('YTC_TIERS'      , TobInfo.GetValue ('T_TIERS'));
      for i:=1 to 10 do
      begin
        if TobInfo.GetValue ('$$_TABLELIBRETIERS'+IntToStr(i)) <> '' then Tob_TiersCompl.PutValue ('YTC_TABLELIBRETIERS'+IntToStr(i), TobInfo.GetValue ('$$_TABLELIBRETIERS'+IntToStr(i)));
      end;
      for i:=1 to 3 do
      begin
        if TobInfo.GetValue ('$$_TEXTELIBRE'+IntToStr(i)) <> ''        then  Tob_TiersCompl.PutValue ('YTC_TEXTELIBRE'+IntToStr(i), TobInfo.GetValue ('$$_TEXTELIBRE'+IntToStr(i)));
        if TobInfo.GetValue ('$$_VALLIBRE'  +IntToStr(i)) <> 0         then  Tob_TiersCompl.PutValue ('YTC_VALLIBRE'  +IntToStr(i), TobInfo.GetValue ('$$_VALLIBRE'  +IntToStr(i)));
        if IsValidDate(TobInfo.GetValue ('$$_DATELIBRE' +IntToStr(i))) then  Tob_TiersCompl.PutValue ('YTC_DATELIBRE' +IntToStr(i), TobInfo.GetValue ('$$_DATELIBRE' +IntToStr(i)));
        if TobInfo.GetValue ('$$_BOOLLIBRE' +IntToStr(i)) <> ''        then  Tob_TiersCompl.PutValue ('YTC_BOOLLIBRE' +IntToStr(i), TobInfo.GetValue ('$$_BOOLLIBRE' +IntToStr(i)));
      end;

      // MAJ DB
//      Tob_TiersCompl.SetAllModifie(TRUE) ;
      Tob_TiersCompl.InsertOrUpdateDB (True);
      FreeAndNil( Tob_TiersCompl );
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique de la TOB Piedeche
Suite ........ :
Suite ........ :
Mots clefs ... : IMPORT PIEDECHE
*****************************************************************}
function TTImportGPAO.CtrlTOBPiedEche (var TobInfo : TOB) : boolean;
var NaturePiece     : string   ;
    IdentReglement  : string   ;
    DeviseReglement : string   ;
    ModeReglement   : string   ;
    SouchePiece     : string   ;
    NumLigReglement : integer  ;
    NumeroPiece     : integer   ;
    IndicePiece     : integer   ;
    SQL             : string   ;
    Q               : TQUERY  ;
    TobPiece        : TOB      ;
    DD   		    : TDateTime;
    DEV             : RDEVISE  ;
    Montant         : double   ;
    MontantDev      : double   ;
    MontantOld      : double   ;
    DevOrg          : string   ;
    DevDest         : string   ;
    MessReglt       : string   ;
    Dateconv        : TDateTime;
begin
  result      := True ;
  TobPiece    := nil  ;
  NumeroPiece := 0    ;
  IndicePiece := 0    ;
  MessReglt   := TraduireMemoire('  Règlement ');

  if TobInfo.FieldExists ('$$_REFINTERNE') then IdentReglement := TobInfo.GetValue ('$$_REFINTERNE')
  else IdentReglement := '';

  //////////////////////////////////////////////////////////////////////////
  // Identification du règlement
  /////////////////////////////////////////////////////////////////////////
  if IdentReglement = '' then
  begin
    EcrireInfo (TraduireMemoire('  Règlement : la référence interne permet de retrouver la pièce, elle est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Nature de la pièce
  /////////////////////////////////////////////////////////////////////////
  NaturePiece := TobInfo.GetValue ('GPE_NATUREPIECEG');
  if result and (NaturePiece = '') then
  begin
    EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : la nature de pièce est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Devise de règlement
  /////////////////////////////////////////////////////////////////////////
  DeviseReglement := TobInfo.GetValue ('GPE_DEVISE');
  if result and (DeviseReglement = '') then
  begin
    EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : la devise de règlement est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Mode de règlement
  /////////////////////////////////////////////////////////////////////////
  ModeReglement := TobInfo.GetValue ('GPE_MODEPAIE');
  if result and (ModeReglement = '') then
  begin
    EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : le mode de paiement est obligatoire'), True, True, True);
    result := False;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Mode de paiement
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    SQL:='Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="'+ModeReglement+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : le mode de paiement ') + ModeReglement + TraduireMemoire('n''existe pas'), True, True, True);
      result := False;
    end;
    Ferme (Q);
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Numéro d'échéance
  ////////////////////////////////////////////////////////////////////////////
  NumLigReglement := TobInfo.GetValue ('GPE_NUMECHE');
  if result and (NumLigReglement = 0) then
  begin
    EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : le numéro d''échéance doit être supérieur à 0'), True, True, True);
    result := False;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Recherche de la pièce
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    Q := OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_REFINTERNE="' + IdentReglement + '" ',False) ;
    if not Q.EOF then
    begin
      SouchePiece := Q.FindField ('GP_SOUCHE').AsString;
      NumeroPiece := Q.FindField ('GP_NUMERO').AsInteger;
      IndicePiece := Q.FindField ('GP_INDICEG').AsInteger;

      TobPiece:= TOB.CREATE ('PIECE', NIL, -1);
      AddLesSupEntete (TobPiece) ;
      TobPiece.SelectDB('',Q);
    end else
    begin
      EcrireInfo (MessReglt + IdentReglement + TraduireMemoire(' : Impossible de retrouver la pièce correspondante'), True, True, True);
      result := False;
    end;
    Ferme (Q);
  end;

  if result then
  begin
    //////////////////////////////////////////////////////////////////////////
    // Chargement infos devise
    //////////////////////////////////////////////////////////////////////////
    if (result) then
    begin
      DD       := TobPiece.GetValue('GP_DATEPIECE');
      DEV.Code := DeviseReglement;
      GetInfosDevise (DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, DD) ;
    end;

    // Initialisation
    TobInfo.PutValue        ('GPE_NATUREPIECEG' , NaturePiece);
    TobInfo.PutValue        ('GPE_DATEPIECE'    , TobPiece.GetValue('GP_DATEPIECE'));
    TobInfo.PutValue        ('GPE_SOUCHE'       , SouchePiece);
    TobInfo.PutValue        ('GPE_NUMERO'       , NumeroPiece);
    TobInfo.PutValue        ('GPE_INDICEG'      , IndicePiece);
    TobInfo.PutValue        ('GPE_NUMECHE'      , NumLigReglement);
    TobInfo.PutValue        ('GPE_TIERS'        , TobPiece.GetValue('GP_TIERS'));
    TobInfo.PutValue        ('GPE_CAISSE'       , TobPiece.GetValue('GP_CAISSE'));
    TobInfo.PutValue        ('GPE_NUMZCAISSE'   , TobPiece.GetValue('GP_NUMZCAISSE'));

    TobInfo.PutValue        ('GPE_MODEPAIE'     , ModeReglement);
    TobInfo.PutValue        ('GPE_DATEECHE'     , TobPiece.GetValue('GP_DATEPIECE'));
    TobInfo.PutValue        ('GPE_TAUXDEV'      , DEV.Taux);
    TobInfo.PutValue        ('GPE_DATETAUXDEV'  , TobPiece.GetValue ('GP_DATETAUXDEV'));
    AttribCotation          (TobInfo) ;   // Calcul GP_COTATION
    TobInfo.PutValue        ('GPE_DEVISEESP'    , TobInfo.GetValue ('GPE_DEVISE'));
    //////////////////////////////////////////////////////////////////////////
    // Conversion des montants en devise
    //////////////////////////////////////////////////////////////////////////
    Montant := TobInfo.GetValue ('GPE_MONTANTDEV');
    TobInfo.PutValue ('GPE_MONTANTENCAIS', Montant);
    //
    // Conversion des autres montants en devise dossier
    //
    DevOrg  := TobInfo.GetValue ('GPE_DEVISE') ;
    DevDest := TobPiece.GetValue('GP_DEVISE') ;

    if  TobPiece.GetValue ('GP_SAISIECONTRE')= 'X' then
    begin
      if not VH^.TenueEuro then DevDest := 'EUR'
      else DevDest := V_PGI.DeviseFongible ;
    end;

    //
    // Calcul du montant dans la devise de la pièce
    //
    Dateconv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    Montant    := TobInfo.GetValue ('GPE_MONTANTENCAIS') ;

    ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, nil) ;
    TobInfo.PutValue ('GPE_MONTANTDEV' , MontantDev   ) ;
    //
    // Calcul du montant dans la devise dossier
    //
    DevDest := V_PGI.DevisePivot ;
    Dateconv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    Montant    := TobInfo.GetValue ('GPE_MONTANTENCAIS') ;

    ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, nil) ;
    TobInfo.PutValue ('GPE_MONTANTECHE', MontantDev) ;

    //
    // On force la devise du règlement avec la devise de la pièce
    //
    TobInfo.PutValue ('GPE_SAISIECONTRE',  TobPiece.GetValue ('GP_SAISIECONTRE'));
    TobInfo.PutValue ('GPE_DEVISE'      ,  TobPiece.GetValue ('GP_DEVISE'));
  end;
  FreeAndNil( TobPiece );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique de la TOB Acomptes
Suite ........ :
Suite ........ :
Mots clefs ... : IMPORT Acomptes
*****************************************************************}
function TTImportGPAO.CtrlTOBAcomptes (var TobInfo : TOB) : boolean;
var NaturePiece 		: string   ;
    IdentReglement  : string   ;
    DeviseReglement : string   ;
    DevisePiece     : string   ;
    ModeReglement   : string   ;
    SouchePiece     : string   ;
    NumLigReglement : integer  ;
    NumeroPiece     : integer  ;
    IndicePiece     : integer  ;
    SQL             : string   ;
    Q               : TQUERY   ;
    TobPiece        : TOB      ;
    DD  		 	: TDateTime;
    DEV             : RDEVISE  ;
    Montant         : double   ;
    MontantDev      : double   ;
    MontantOld      : double   ;
    DevOrg          : string   ;
    DevDest         : string   ;
    MessAcompte     : string   ;
    Dateconv        : TDateTime;
begin
  result      := True ;
  TobPiece    := nil  ;
  NumeroPiece := 0    ;
  IndicePiece := 0    ;
  DD          := iDate1900;
  MessAcompte := TraduireMemoire('  Acompte ');

  if TobInfo.FieldExists ('$$_REFINTERNE') then IdentReglement := TobInfo.GetValue ('$$_REFINTERNE')
  else IdentReglement := '';

  //////////////////////////////////////////////////////////////////////////
  // Identification du règlement
  /////////////////////////////////////////////////////////////////////////
  if IdentReglement = '' then
  begin
    EcrireInfo (TraduireMemoire('  Acompte : la référence interne permet de retrouver la pièce, elle est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Nature de la pièce
  /////////////////////////////////////////////////////////////////////////
  NaturePiece := TobInfo.GetValue ('GAC_NATUREPIECEG');
  if result and (NaturePiece = '') then
  begin
    EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : la nature de pièce est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Devise de règlement
  /////////////////////////////////////////////////////////////////////////
  if TobInfo.FieldExists ('$$_DEVISE') then DeviseReglement := TobInfo.GetValue ('$$_DEVISE')
  else DeviseReglement := '';
  if result and (DeviseReglement = '') then
  begin
    EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : la devise de règlement est obligatoire'), True, True, True);
    result := False;
  end;
  //////////////////////////////////////////////////////////////////////////
  // Mode de règlement
  /////////////////////////////////////////////////////////////////////////
  ModeReglement := TobInfo.GetValue ('GAC_MODEPAIE');
  if result and (ModeReglement = '') then
  begin
    EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : le mode de paiement est obligatoire'), True, True, True);
    result := False;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Mode de paiement
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    SQL:='Select MP_LIBELLE From MODEPAIE WHERE MP_MODEPAIE="'+ModeReglement+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : le mode de paiement ') + ModeReglement + TraduireMemoire('n''existe pas'), True, True, True);
      result := False;
    end;
    Ferme (Q);
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Numéro d'échéance
  ////////////////////////////////////////////////////////////////////////////
  NumLigReglement := TobInfo.GetValue ('GAC_NUMECR');
  if result and (NumLigReglement = 0) then
  begin
    EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : le numéro de ligne doit être supérieur à 0'), True, True, True);
    result := False;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Recherche de la pièce
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    Q := OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_REFINTERNE="' + IdentReglement + '" ',False) ;
    if not Q.EOF then
    begin
      SouchePiece := Q.FindField ('GP_SOUCHE').AsString;
      NumeroPiece := Q.FindField ('GP_NUMERO').AsInteger;
      IndicePiece := Q.FindField ('GP_INDICEG').AsInteger;
      DD          := Q.FindField ('GP_DATEPIECE').AsDateTime;

      TobPiece:= TOB.CREATE ('PIECE', NIL, -1);
      AddLesSupEntete (TobPiece) ;
      TobPiece.SelectDB('',Q);
    end else
    begin
      EcrireInfo (MessAcompte + IdentReglement + TraduireMemoire(' : Impossible de retrouver la pièce correspondante'), True, True, True);
      result := False;
    end;
    Ferme (Q);
  end;

  if result then
  begin
    //////////////////////////////////////////////////////////////////////////
    // Chargement infos devise
    //////////////////////////////////////////////////////////////////////////
    if (result) then
    begin
      DEV.Code := DeviseReglement;
      GetInfosDevise (DEV) ;
      DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, DD) ;
    end;

    // Initialisation
    TobInfo.PutValue        ('GAC_NATUREPIECEG' , NaturePiece);
    TobInfo.PutValue        ('GAC_SOUCHE'       , SouchePiece);
    TobInfo.PutValue        ('GAC_NUMERO'       , NumeroPiece);
    TobInfo.PutValue        ('GAC_INDICEG'      , IndicePiece);
    TobInfo.PutValue        ('GAC_NUMECR'       , NumLigReglement);
    TobInfo.PutValue        ('GAC_MODEPAIE'     , ModeReglement);

    //////////////////////////////////////////////////////////////////////////
    // Conversion des montants en devise
    //////////////////////////////////////////////////////////////////////////
    Montant := TobInfo.GetValue ('GAC_MONTANTDEV');

    // Conversion des autres montants en devise dossier
    DevOrg     := TobInfo.GetValue ('$$_DEVISE') ;
    DevDest    := V_PGI.DevisePivot;
    Dateconv   := NowH ;
    MontantDev := 0 ;  // initialisation bidon
    MontantOld := 0 ;  // initialisation bidon
    ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, nil) ;
    TobInfo.PutValue ('GAC_MONTANT'    , MontantDev) ;
    TobInfo.PutValue ('GAC_MONTANTDEV' , MontantDev) ;
    ////////////////////////////////////////////////////////////////////////////
    // MAJ des champs acomptes dans la PIECE
    ////////////////////////////////////////////////////////////////////////////
    DevisePiece := TobPiece.GetValue ('GP_DEVISE');
    if VH^.TenueEuro  then
    begin
      if (TobPiece.GetValue ('GP_SAISIECONTRE') = 'X') then DevisePiece := V_PGI.DeviseFongible;
    end else
    begin
      if (TobPiece.GetValue ('GP_SAISIECONTRE') = 'X') then DevisePiece := 'EUR';
    end;
    Montant := TobInfo.GetValue ('GAC_MONTANT') ;
    DevOrg     := V_PGI.DevisePivot ;        // Les acomptes sont toujours exprimés dans
    DevDest    := DevisePiece;
    ToxConvertir (Montant, MontantDev, MontantOld, DevOrg, DevDest, DateConv, nil) ;
    if NumLigReglement = 1 then
    begin
      TobPiece.PutValue ('GP_ACOMPTE'   , Montant);
      TobPiece.PutValue ('GP_ACOMPTEDEV', MontantDev);
    end else
    begin
      TobPiece.PutValue ('GP_ACOMPTE'   , Montant    + TobPiece.GetValue ('GP_ACOMPTE'));
      TobPiece.PutValue ('GP_ACOMPTEDEV', MontantDev + TobPiece.GetValue ('GP_ACOMPTEDEV'));
    end;
    TobPiece.UpdateDB(False);
  end;
  /////////////////////////////////////////////////////////////////////////////
  // Libération de la TOb Piece
  /////////////////////////////////////////////////////////////////////////////
  FreeAndNil( TobPiece );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique de la TOB Acomptes
Suite ........ :
Suite ........ :
Mots clefs ... : IMPORT Acomptes
*****************************************************************}
function TTImportGPAO.CtrlTOBInventaire (var TobInfo : TOB) : boolean;
var Tob_TransInvEnt : TOB    ;
    IdentArticle    : string ;
    Depot           : string ;
    Article         : string ;
    SQL             : string ;
    MessInvent      : string ;
    Q               : TQUERY ;
begin
  result     := True ; // Optimiste
  MessInvent := TraduireMemoire('  Inventaire ') ;

  // Identifiant de la ligne
  if TobInfo.GetValue ('GIN_CODEBARRE') <> '' then IdentArticle := TobInfo.GetValue ('GIN_CODEBARRE')
  else IdentArticle := TobInfo.GetValue ('GIN_CODEARTICLE');
  ////////////////////////////////////////////////////////////////////////////
  // Depôt / Etablissement
  ////////////////////////////////////////////////////////////////////////////
  Depot := TobInfo.GetValue ('GIN_DEPOT');
  ////////////////////////////////////////////////////////////////////////////
  // Depôt
  ////////////////////////////////////////////////////////////////////////////
  if (result) and (Depot <> '') then
  begin
    SQL:='Select GDE_LIBELLE From DEPOTS WHERE GDE_DEPOT="'+Depot+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessInvent + IdentArticle + TraduireMemoire(' : le dépôt ') + Depot + TraduireMemoire(' n''existe pas'), True, True, True);
      result := False;
    end;
    Ferme(Q);
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Article
  ////////////////////////////////////////////////////////////////////////////
  Article := TobInfo.GetValue ('GIN_ARTICLE');
  if (result) and (Article = '') then
  begin
    EcrireInfo (MessInvent + IdentArticle + TraduireMemoire(' : Problème pour retrouver l''article correspondant'), True, True, True);
    result := False;
  end;
  if (result) then
  begin
    SQL:='Select GA_LIBELLE From ARTICLE WHERE GA_ARTICLE="'+Article+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      EcrireInfo (MessInvent + IdentArticle + TraduireMemoire(' : l''article ') + Article + TraduireMemoire(' n''existe pas'), True, True, True);
      result := False;
    end;
    Ferme(Q);
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Libelle
  ////////////////////////////////////////////////////////////////////////////
  if TobInfo.GetValue ('$$_LIBELLE') = '' then
  begin
   //TobInfo.PutValue ('$$_LIBELLE', 'RECUPERATION INVENTAIRE DU ' + DateToStr(Date) + ' A ' + TimeToStr(Time));
   TobInfo.PutValue ('$$_LIBELLE', 'RECUP');
  end;

  ////////////////////////////////////////////////////////////////////////////
  // Contrôle OK : On crée une entête
  ////////////////////////////////////////////////////////////////////////////
  If not EnteteInvCree then
  begin
    Tob_TransInvEnt := TOB.CREATE ('TRANSINVENT', NIL, -1);
    CodeTrans := NewCodeTrans (Depot) ;

    Tob_TransInvEnt.PutValue ('GIT_DEPOT'         , Depot);
    Tob_TransInvEnt.PutValue ('GIT_CODETRANS'     , CodeTrans);
    Tob_TransInvEnt.PutValue ('GIT_LIBELLE'       , TobInfo.GetValue('$$_LIBELLE'));
    Tob_TransInvEnt.PutValue ('GIT_ETABLISSEMENT' , Depot);
    Tob_TransInvEnt.PutValue ('GIT_CREATEUR'      , V_PGI.User);
    Tob_TransInvEnt.PutValue ('GIT_UTILISATEUR'   , V_PGI.User);
    Tob_TransInvEnt.PutValue ('GIT_INTEGRATION'   , '-');

    Tob_TransInvEnt.InsertDB (nil);
    FreeAndNil( Tob_TransInvEnt );

    EnteteInvCree := True ;
  end;
  //
  // Contrôle OK On crée la ligne
  //
  Inc (NumLigneInv);
  TobInfo.PutValue ('GIN_CODETRANS'  , CodeTrans);
  TobInfo.PutValue ('GIN_NUMLIGNE'   , NumligneInv);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Mise à jour spécifique de la TOB fidelite
Suite ........ : Attention, l'assistant d'import ne crée que des lignes
Suite ........ : fidelite de type "REPRISE"
Mots clefs ... : IMPORT Acomptes
*****************************************************************}
function TTImportGPAO.CtrlTOBFidelite (var TobInfo : TOB) : boolean;
var Tob_FideliteEnt : TOB    ;
    Tob_FideliteLig : TOB    ;
    StRech          : string ;
    Tiers           : string ;
    Programme       : string ;
    SQL             : string ;
    MessFidelite    : string ;
    CodeMax         : string ;
    NumLigne        : integer;
    Q               : TQUERY ;
begin
  result     := True ; // Optimiste
  MessFidelite := TraduireMemoire('  Carte fidélité du client ') ;
  /////////////////////////////////////////////////////////////////
  // Contrôle 1	:	Tiers
  /////////////////////////////////////////////////////////////////
  StRech :=  TobInfo.GetValue ('$$_TIERS');
  if (StRech = '') then
  begin
    EcrireInfo (MessFidelite + ' : ' + TraduireMemoire(' : le tiers n''est pas renseigné'), True, True, True);
    result := False
  end else
  begin
    if (StRech <> '') and (StRech <> DernierTiers) then
    begin
      SQL:='Select T_LIBELLE From TIERS WHERE T_TIERS="'+StRech+'"';
      Q:=OpenSQL(SQL,True) ;
      if Q.EOF then
      begin
        EcrireInfo (MessFidelite + TobInfo.GetValue ('$$_TIERS') + ' : ' + TraduireMemoire(' : le tiers est inconnu'), True, True, True);
        result := False ;
      end else DernierTiers := StRech;
      Ferme(Q);
    end;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Etablissement
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    StRech := TobInfo.GetValue ('GFI_ETABLISSEMENT');
    if (StRech = '') then
    begin
      EcrireInfo (MessFidelite + TobInfo.GetValue ('$$_TIERS') + ' : ' + TraduireMemoire(' : l''établissement n''est pas renseigné'), True, True, True);
      result := False ;
    end else
    begin
      if (StRech <> DernierEtab) then
      begin
        SQL:='Select ET_PROGRAMME From ETABLISS WHERE ET_ETABLISSEMENT="'+StRech+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then
        begin
          EcrireInfo (MessFidelite + TobInfo.GetValue ('$$_TIERS') + ' : ' + TraduireMemoire(' : l''établissement n''existe pas'), True, True, True);
          result := False ;
        end
        else begin
          CodeProgramme := Q.Findfield('ET_PROGRAMME').AsString;
          DernierEtab := StRech ;
        end;
        Ferme(Q);
      end;
    end;
  end;
  ////////////////////////////////////////////////////////////////////////////
  // Programme
  ////////////////////////////////////////////////////////////////////////////
  if result then
  begin
    StRech := TobInfo.GetValue ('GFI_PROGRAMME');
    if (StRech = '') then
    begin
      TobInfo.PutValue ('GFI_PROGRAMME', CodeProgramme) ;
      EcrireInfo (MessFidelite + TobInfo.GetValue ('$$_TIERS') + ' : ' + TraduireMemoire(' : Attention, le code programme est repris de l''établissement'), True, True, True);
    end else
    begin
      if (StRech <> DernierCodeProgramme) then
      begin
        SQL:='Select GFO_LIBELLE, GFO_TYPECUMULFID From PARFIDELITE WHERE GFO_CODEFIDELITE="'+StRech+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then
        begin
          EcrireInfo (MessFidelite + TobInfo.GetValue ('$$_TIERS') + ' : ' + TraduireMemoire(' : Le programme de fidélité n''existe pas'), True, True, True);
          result := False ;
        end else
        begin
          TypeCumulFid         := Q.Findfield('GFO_TYPECUMULFID').AsString;
          DernierCodeProgramme := StRech ;
        end;
        Ferme(Q);
      end;
    end;
  end;

  ////////////////////////////////////////////////////////////////////////////
  // Le client a t il déjà une carte ouverte ?
  ////////////////////////////////////////////////////////////////////////////
  Tiers     := TobInfo.GetValue ('$$_TIERS') ;
  Programme := TobInfo.GetValue ('GFI_PROGRAMME') ;

  SQL:='Select * From FIDELITEENT WHERE GFE_TIERS="'+Tiers+'" AND GFE_PROGRAMME="'+Programme+'" AND GFE_FERME="-"';
  Q:=OpenSQL(SQL,True) ;
  if Q.EOF then
  begin
    Ferme(Q) ;
    ////////////////////////////////////////////////////////////////////////////
    // Contrôle OK : On crée une entête
    ////////////////////////////////////////////////////////////////////////////
    If result then
    begin
      Tob_FideliteEnt := TOB.CREATE ('FIDELITEENT', NIL, -1);
      Tob_FideliteEnt.PutValue('GFE_TIERS'        , TobInfo.GetValue ('$$_TIERS'));
      Tob_FideliteEnt.PutValue('GFE_ETABLISSEMENT', TobInfo.GetValue ('GFI_ETABLISSEMENT'));
      Tob_FideliteEnt.PutValue('GFE_FERME'        , TobInfo.GetValue ('$$_FERME'));
      Tob_FideliteEnt.PutValue('GFE_DATEFERMETURE', TobInfo.GetValue ('$$_DATEFERMETURE'));
      Tob_FideliteEnt.PutValue('GFE_NUMCARTEINT'  , CreationNumeroCarteInterne(TobInfo.GetValue ('GFI_ETABLISSEMENT')));
      //inc (CptFid);
      //Tob_FideliteEnt.PutValue('GFE_NUMCARTEINT'  , InttoStr(Cptfid));

      Tob_FideliteEnt.PutValue('GFE_NUMCARTEEXT'  , TobInfo.GetValue ('$$_NUMCARTEEXT'));
      Tob_FideliteEnt.PutValue('GFE_PROGRAMME'    , TobInfo.GetValue ('GFI_PROGRAMME'));
      Tob_FideliteEnt.PutValue('GFE_DATECREATION' , NowH);
      Tob_FideliteEnt.PutValue('GFE_DATEMODIF'    , NowH);
      Tob_FideliteEnt.PutValue('GFE_DATEINTEGR'   , iDate1900);
      Tob_FideliteEnt.PutValue('GFE_CREATEUR'     , V_PGI.User);
      Tob_FideliteEnt.PutValue('GFE_UTILISATEUR'  , V_PGI.User);
      if TobInfo.GetValue ('$$_DATEOUVERTURE') = iDate1900 then
        Tob_FideliteEnt.PutValue('GFE_DATEOUVERTURE', TobInfo.GetValue ('$$_DATEOUVERTURE'))
      else
        Tob_FideliteEnt.PutValue('GFE_DATEOUVERTURE', NowH);
      //
      // Initailisation des autres champs de la ligne
      //
      TobInfo.PutValue ('GFI_NUMCARTEINT'    ,  Tob_FideliteEnt.GetValue('GFE_NUMCARTEINT'));
      TobInfo.PutValue ('GFI_LIGNE'          ,  1);
      TobInfo.PutValue ('GFI_DATECREATION'   , NowH);
      TobInfo.PutValue ('GFI_DATEMODIF'      , NowH);
      TobInfo.PutValue ('GFI_DATEINTEGR'     , iDate1900);
      TobInfo.PutValue ('GFI_CREATEUR'       , V_PGI.User);
      TobInfo.PutValue ('GFI_UTILISATEUR'    , V_PGI.User);
      TobInfo.PutValue ('GFI_TYPELIGNEFIDEL' , '999');
      TobInfo.PutValue ('GFI_TYPECUMULFID'   , TypeCumulFid);
      //
      // Mise à jour de la base
      //
      Tob_FideliteEnt.InsertDB(nil);
      FreeAndNil( Tob_FideliteEnt );
     end;
   end else
   begin
     Tob_FideliteEnt := TOB.CREATE ('FIDELITEENT', NIL, -1);
     Tob_FideliteEnt.SelectDB('',Q);
     Ferme(Q) ;

     SQL:='Select * From FIDELITELIG WHERE GFI_NUMCARTEINT="'+Tob_FideliteEnt.GetValue('GFE_NUMCARTEINT')+'" AND GFI_ETABLISSEMENT="'+Tob_FideliteEnt.GetValue('GFE_ETABLISSEMENT')+'" AND GFI_TYPELIGNEFIDEL="999"';
     Q:=OpenSQL(SQL,True) ;
     if Q.EOF then
     begin
       Ferme(Q) ;

       SQL:='Select MAX(GFI_LIGNE) From FIDELITELIG WHERE GFI_NUMCARTEINT="'+Tob_FideliteEnt.GetValue('GFE_NUMCARTEINT')+'" AND GFI_ETABLISSEMENT="'+Tob_FideliteEnt.GetValue('GFE_ETABLISSEMENT')+'"';
       Q:=OpenSQL(SQL,True) ;
       NumLigne :=1 ;
       if Not Q.EOF then
       begin
         CodeMax:=Q.Fields[0].AsString;
         if CodeMax <> '' then NumLigne := StrToInt(CodeMax)+1;
       end;
       Ferme (Q);
       TobInfo.PutValue ('GFI_NUMCARTEINT'    , Tob_FideliteEnt.GetValue('GFE_NUMCARTEINT'));
       TobInfo.PutValue ('GFI_LIGNE'          , NumLigne);
       TobInfo.PutValue ('GFI_DATECREATION'   , NowH);
       TobInfo.PutValue ('GFI_DATEMODIF'      , NowH);
       TobInfo.PutValue ('GFI_DATEINTEGR'     , iDate1900);
       TobInfo.PutValue ('GFI_CREATEUR'       , V_PGI.User);
       TobInfo.PutValue ('GFI_UTILISATEUR'    , V_PGI.User);
       TobInfo.PutValue ('GFI_TYPELIGNEFIDEL' , '999');
       TobInfo.PutValue ('GFI_TYPECUMULFID'   , TypeCumulFid);
     end else
     begin
       Tob_FideliteLig := TOB.CREATE ('FIDELITELIG', NIL, -1);
       Tob_FideliteLig.SelectDB('',Q);
       Ferme(Q) ;

       Tob_FideliteLig.PutValue ('GFI_VALEUR', TobInfo.GetValue ('GFI_VALEUR'));
       TobInfo.Dupliquer (Tob_FideliteLig, True, True, True);
       FreeAndNil( Tob_FideliteLig );
     end;
     FreeAndNil( Tob_FideliteEnt );
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI
Suite ........ :  1 - Chargement les articles d'une pièce en TOB + DISPO
               : Fonction non optimisée , on fait un SELECT sur chaque fiche article et stock
               : A utiliser lors d'une création de pièces
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.ChargeLesArticles (TobPiece, TobArticle : TOB);
var TobFilleArticle : TOB      ;
    Q               : TQUERY   ;
    CodeArt         : string   ;
    SQL             : string   ;
    cpt             : integer  ;
begin

  for Cpt:=0 to TobPiece.Detail.Count-1 do
  begin
    CodeArt := TobPiece.Detail[Cpt].GetValue ('GL_ARTICLE');
    if CodeArt <> '' then
    begin
      SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+CodeArt+'"';
      Q:=OpenSQL(SQL,True) ;
      if Not Q.EOF then
      begin
        TobFilleArticle := CreerTOBArt(TobArticle);
        TobFilleArticle.SelectDB('',Q);
        // Chargement des stocks
        LoadTOBDispo (TobFilleArticle, True, CreerQuelDepot(TobPiece)) ;
      end;
      Ferme (Q);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI
Suite ........ :  1 - Chargement les articles d'une pièce en TOB + DISPO
Mots clefs ... : IMPORT PIECE
*****************************************************************}
{procedure TTImportGPAO.ChargeLesArticles2 (TobPiece, TobArticle : TOB);
var TobFilleArticle : TOB      ;
    Q               : TQUERY   ;
    cpt             : integer  ;
    CleDoc_Piece    : R_CLeDoc ;
begin

  // Idée 1 : Charger les articles à partir des lignes en 1 seule requête
  CleDoc_Piece := TOB2CleDoc (TobPiece);
  //
  // Chargement des articles des lignes
  //
  Q:=OpenSQL('SELECT ARTICLE.* FROM ARTICLE, LIGNE WHERE ' + WherePiece(CleDoc_Piece,ttdLigne,False) + ' AND GL_ARTICLE=GA_ARTICLE',True) ;
  TobArticle.LoadDetailDB('ARTICLE','','',Q,False,True) ;
  Ferme(Q) ;
  //
  // Ajout des champs supplémentaires
  //
  for Cpt:=0 to TobArticle.Detail.Count-1 do
  begin
    TobFilleArticle := TobArticle.Detail [cpt] ;
    TobFilleArticle.AddChampSup('REFARTBARRE' ,False) ; TobFilleArticle.PutValue('REFARTBARRE' ,'') ;
    TobFilleArticle.AddChampSup('REFARTTIERS' ,False) ; TobFilleArticle.PutValue('REFARTTIERS' ,'') ;
    TobFilleArticle.AddChampSup('REFARTSAISIE',False) ; TobFilleArticle.PutValue('REFARTSAISIE','') ;
    TobFilleArticle.AddChampSup('UTILISE'     ,False) ; TobFilleArticle.PutValue('UTILISE'     ,'-');
    //
    // Chargement des stocks
    //
    LoadTOBDispo (TobFilleArticle, True, CreerQuelDepot(TobPiece)) ;
  end;
end;
}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI
Suite ........ :  1 - Chargement les articles d'une pièce en TOB + DISPO
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.ChargeLesArticles3 (TobPiece, TobArticle : TOB);
var TobFilleArticle : TOB      ;
    TOBDispo        : TOB      ;
    TobDispoArt     : TOB      ;
    Q               : TQUERY   ;
    cpt             : integer  ;
    CleDoc_Piece    : R_CLeDoc ;
begin

  // Idée 1 : Charger les articles à partir des lignes en 1 seule requête
  CleDoc_Piece := TOB2CleDoc (TobPiece);
  //
  // Chargement des articles des lignes
  //
  Q:=OpenSQL('SELECT ARTICLE.* FROM ARTICLE, LIGNE WHERE ' + WherePiece(CleDoc_Piece,ttdLigne,False) + ' AND GL_ARTICLE=GA_ARTICLE',True) ;
  TobArticle.LoadDetailDB('ARTICLE','','',Q,False,True) ;
  Ferme(Q) ;
  //
  // Chargement des fiches stocks
  //
  TobDispo :=  TOB.CREATE ('DISPO', NIL, -1);
  Q:=OpenSQL('SELECT DISPO.* FROM DISPO, LIGNE WHERE ' + WherePiece(CleDoc_Piece,ttdLigne,False) + ' AND GQ_CLOTURE="-" AND GL_DEPOT=GQ_DEPOT AND GL_ARTICLE=GQ_ARTICLE AND GL_TYPEARTICLE="MAR"',True) ;
  TobDispo.LoadDetailDB('DISPO','','',Q,False,True) ;
  Ferme(Q) ;
  //
  //  Ajout des champs supp MODE
  //
  DispoChampSupp(TobDispo);
  //
  // Rattachement des fiches stocks aux fiches articles
  //
  for cpt:=0 to TobArticle.detail.Count-1 do
  begin
    //
    //  Ajout des champs supp
    //
    TobFilleArticle := TobArticle.Detail [cpt] ;
    TobFilleArticle.AddChampSup('REFARTBARRE' ,False) ; TobFilleArticle.PutValue('REFARTBARRE' ,'') ;
    TobFilleArticle.AddChampSup('REFARTTIERS' ,False) ; TobFilleArticle.PutValue('REFARTTIERS' ,'') ;
    TobFilleArticle.AddChampSup('REFARTSAISIE',False) ; TobFilleArticle.PutValue('REFARTSAISIE','') ;
    TobFilleArticle.AddChampSup('UTILISE'     ,False) ; TobFilleArticle.PutValue('UTILISE'     ,'-');

    TobDispoArt:=TOBDispo.FindFirst(['GQ_ARTICLE'],[TobFilleArticle.GetValue('GA_ARTICLE')],False) ;
    if TobDispoArt<>nil then TobDispoArt.Changeparent(TobFilleArticle,-1);
  end;
  FreeAndNil( TobDispo );
end;

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
procedure TTImportGPAO.ChargeLaPiecePrecedente (NaturePiece, RefInterne : string);
var CleDoc_Prec     : R_CLeDoc ;
    Q               : TQUERY   ;
begin
  // chargement entête
  Tob_Piece_Prec_Old := TOB.CREATE ('PIECE', NIL, -1);
  AddLesSupEntete (Tob_Piece_Prec_Old) ;
  // Lecture entête
  Q:=OpenSQL('Select * from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_REFINTERNE="' + RefInterne + '" AND GP_VIVANTE="X" ',False) ;
  Tob_Piece_Prec_Old.SelectDB('',Q) ;
  Ferme(Q) ;
  CleDoc_Prec := TOB2CleDoc (Tob_Piece_Prec_Old);

  // Chargement des Lignes
  Q:=OpenSQL('SELECT * FROM LIGNE WHERE '+WherePiece(CleDoc_Prec,ttdLigne,False)+' ORDER BY GL_NUMLIGNE',True) ;
  Tob_Piece_Prec_Old.LoadDetailDB('LIGNE','','',Q,False,True) ;
  Ferme(Q) ;
  PieceAjouteSousDetail(Tob_Piece_Prec_Old);

  // Lecture bases
  Tob_PiedBase_Prec_Old := TOB.CREATE ('Les Taxes' , nil, -1);
  // ZZZZZZZZZZZZZZZZZZZZZ
  Q:=OpenSQL('SELECT * FROM PIEDBASE WHERE '+WherePiece(CleDoc_Prec,ttdPiedBase,False),True) ;
  Tob_PiedBase_Prec_Old.LoadDetailDB('PIEDBASE','','',Q,False) ;
  Ferme(Q) ;

  // Lecture échéances
  Tob_PiedEche_Prec_Old := TOB.CREATE ('Les Echéances' , nil, -1);
  Q:=OpenSQL('SELECT * FROM PIEDECHE WHERE '+WherePiece(CleDoc_Prec,ttdeche,False),True) ;
  Tob_PiedEche_Prec_Old.LoadDetailDB('PIEDECHE','','',Q,False) ;
  Ferme(Q) ;

  // Lecture Adresses
  Tob_Adresses_Prec_Old := TOB.CREATE ('Les Echéances' , nil, -1);
  Q:=OpenSQL('SELECT * FROM PIECEADRESSE WHERE ' + WherePiece(CleDoc_Prec, ttdPieceAdr, False) + ' ORDER BY GPA_TYPEPIECEADR', True);
  Tob_Adresses_Prec_Old.LoadDetailDB('PIECEADRESSE','','',Q,False) ;
  Ferme(Q) ;

  /////////////////////////////////////////////////////////////////////////
  // Chargement de la TOB des articles + Fiches DISPO
  /////////////////////////////////////////////////////////////////////////
  Tob_Article_Prec_Old := TOB.CREATE ('Les Articles', nil, -1);
  ChargeLesArticles3 (Tob_Piece_Prec_Old, Tob_Article_Prec_Old);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI avec solde de la pièce précédente
Suite ........ :  1 - Balayage de la commande
Suite ........ :  2 - Recherche de la ligne de réception
Suite ........ :  2 - Calcul du reliquat
Suite ........ :  3 - MAJ des lignes de cdes et de réceptions
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.MetAJourQteReste (RefInterne : string);
var TOBL_Cde    : TOB      ;
    TOBL_Rec    : TOB      ;
    TOBL_RecGen : TOB      ;
    CleDoc_Prec : R_CLeDoc ;
    i           : integer  ;
    Ecart       : double   ;
begin
  // Annule l'impact de cette pièce sur le stock avant modification
  if MajStockPiece then AnnuleStockPieceASolder (Tob_Piece_Prec_Old, Tob_Article_Prec_Old);

  // Maj du champ "pièce suivante" du document initial
  if Tob_Piece_Prec_Old.GetValue ('GP_DEVENIRPIECE') = '' then
  begin
    CleDoc_Prec := TOB2CleDoc (Tob_Piece);
    Tob_Piece_Prec_Old.PutValue ('GP_DEVENIRPIECE', EncodeRefPiece(Tob_Piece,CleDoc_Prec.NumeroPiece));
  end;

  // On balaye les lignes de la pièce d'origine, on recherche les lignes
  // dans la nouvelle pièce et on calcule le reliquat
  TobL_Rec := nil ;
  for i:=0 to Tob_Piece_Prec_Old.Detail.Count-1 do
  begin
    // chargement de la ligne de cde à traiter
    Tobl_Cde := Tob_Piece_Prec_Old.detail [i] ;

    // Recherche de la ligne de réception correspondante
    if Tobl_Cde.GetValue ('GL_TYPELIGNE') = 'ART' then
      TOBL_Rec := Tob_Piece.FindFirst(['$$_REFINTERNE','GL_ARTICLE'],[RefInterne,Tobl_Cde.GetValue ('GL_ARTICLE')],False)
    else
      continue ;

    // Premier cas : la ligne n'est pas reçu, elle ne change pas et on a encore du reliquat
    if TOBL_Rec = nil then
    begin
      if Tobl_Cde.GetValue ('GL_QTERESTE') <> 0 then EncoreReliquat := True ;
      Continue ;
    end;

    // Calcul du reste à livrer
    if not (TOBL_Rec.FieldExists ('$$_DEJAUTILISE')) then
    begin
      TOBL_Rec.AddChampSup ('$$_DEJAUTILISE'      , False);
      TOBL_Rec.PutValue    ('$$_DEJAUTILISE'      , 'X');
      TOBL_Rec.AddChampSup ('$$_VALEUR_A_AFFECTER', False);
      TOBL_Rec.PutValue    ('$$_VALEUR_A_AFFECTER', TOBL_Rec.GetValue('GL_QTERESTE'));
    end;

    Ecart:=Arrondi (Tobl_Cde.GetValue('GL_QTERESTE')-TOBL_Rec.GetValue('$$_VALEUR_A_AFFECTER'),6) ;

    // 1ier cas  : Plus de commandé que de reçu
    //             On conserve un reliquat sur la ligne de commande
    //             La quantité reçu est utilisé en totalité pour solder la ligne de commande
    //
    // 2ième cas : Le reçu est supérieur au commandé
    //             On ne gère pas de reliquat négatif
    //             On solde la ligne de commande et pas plus
    //             On conserve le reste à affecter dans le cas
    //             où l'article serait dans une autre ligne du document
    if (Tobl_Cde.GetValue('GL_QTERESTE') >= 0) then
    begin
      if Ecart >= 0 then
      begin
        Tobl_Cde.PutValue ('GL_QTERESTE'         , Ecart);
        TOBL_Rec.PutValue ('$$_VALEUR_A_AFFECTER', 0);
        if (Ecart>0) then EncoreReliquat := True ;
      end else
      begin
        Tobl_Cde.PutValue ('GL_QTERESTE'         , 0);
        TOBL_Rec.PutValue ('$$_VALEUR_A_AFFECTER', -Ecart);
      end;
    end else
    begin
      if Ecart <= 0 then
      begin
        Tobl_Cde.PutValue ('GL_QTERESTE'         , Ecart);
        TOBL_Rec.PutValue ('$$_VALEUR_A_AFFECTER', 0);
        if (Ecart<0) then EncoreReliquat := True ;
      end else
      begin
        Tobl_Cde.PutValue ('GL_QTERESTE'         , 0);
        TOBL_Rec.PutValue ('$$_VALEUR_A_AFFECTER', -Ecart);
      end;
    end;

    // Encode la pièce précédente
    Tobl_Rec.PutValue ('GL_PIECEPRECEDENTE', EncodeRefPiece(Tobl_Cde)) ;
    if Tobl_Rec.GetValue ('GL_TYPELIGNE') = 'ART' then Tobl_Rec.PutValue ('GL_PIECEORIGINE', EncodeRefPiece(Tobl_Cde)) ;

    // Recalcul des montants sur les lignes génériques
    // O'NEIL 05012003
    if Tobl_Rec.GetValue ('GL_TYPELIGNE') = 'ART' then
    begin
      if Tobl_Rec.GetValue ('GL_FACTUREHT')='X' then
      begin
        if (Tobl_Rec.GetValue ('GL_PUHTDEV') = 0) and (Tobl_Cde.GetValue ('GL_PUHTDEV') <> 0) then
        begin
          Tobl_Rec.PutValue('GL_PUHTDEV'     , Tobl_Cde.GetValue ('GL_PUHTDEV'    )) ;
          Tobl_Rec.PutValue('GL_PUHTNETDEV'  , Tobl_Cde.GetValue ('GL_PUHTNETDEV' )) ;
          Tobl_Rec.PutValue('GL_REMISELIGNE' , Tobl_Cde.GetValue ('GL_REMISELIGNE')) ;
          Tobl_Rec.PutValue('GL_REMISEPIED'  , Tobl_Cde.GetValue ('GL_REMISEPIED' )) ;
          Tobl_Rec.PutValue('GL_ESCOMPTE'    , Tobl_Cde.GetValue ('GL_ESCOMPTE'   )) ;
          Tobl_Rec.PutValue('GL_DPA'         , Tobl_Cde.GetValue ('GL_DPA'        )) ;
          Tobl_Rec.PutValue('GL_PMAP'        , Tobl_Cde.GetValue ('GL_PMAP'       )) ;
          Tobl_Rec.PutValue('GL_PMRP'        , Tobl_Cde.GetValue ('GL_PMRP'       )) ;
          Tobl_Rec.PutValue('GL_PMAPACTU'    , Tobl_Cde.GetValue ('GL_PMAPACTU'   )) ;
          Tobl_Rec.PutValue('GL_PMRPACTU'    , Tobl_Cde.GetValue ('GL_PMRPACTU'   )) ;
          Tobl_Rec.PutValue('GL_DPR'         , Tobl_Cde.GetValue ('GL_DPR'        )) ;

          // Mise à jour de la ligne générique
          if Tobl_Rec.GetValue ('GL_TYPEDIM') = 'DIM' then
          begin
            TOBL_RecGen := Tob_Piece.FindFirst(['$$_REFINTERNE','GL_CODESDIM'],[RefInterne,Tobl_Rec.GetValue ('GL_CODEARTICLE')],False) ;
            if TOBL_RecGen<>Nil then
            begin
              TOBL_RecGen.PutValue('GL_PUHTDEV'      , Tobl_Rec.GetValue   ('GL_PUHTDEV')) ;
              TOBL_RecGen.PutValue('GL_PUHTNETDEV'   , Tobl_Rec.GetValue   ('GL_PUHTNETDEV')) ;
              TOBL_RecGen.PutValue('GL_PUHT'         , Tobl_Rec.GetValue   ('GL_PUHT')) ;
              TOBL_RecGen.PutValue('GL_PUHTNET'      , Tobl_Rec.GetValue   ('GL_PUHTDEV')) ;
              TOBL_RecGen.PutValue('GL_MONTANTHTDEV' , TOBL_RecGen.GetValue('GL_MONTANTHTDEV') + Tobl_Rec.GetValue ('GL_PUHTNETDEV') * Tobl_Rec.GetValue ('GL_QTEFACT'));
              TOBL_RecGen.PutValue('GL_MONTANTHT'    , TOBL_RecGen.GetValue('GL_MONTANTHT')    + Tobl_Rec.GetValue ('GL_PUHTNET')    * Tobl_Rec.GetValue ('GL_QTEFACT'));
           end;
          end;
        end;
      end else
      begin
        if (Tobl_Rec.GetValue ('GL_PUTTCDEV') = 0) and (Tobl_Cde.GetValue ('GL_PUTTCDEV') <> 0) then
        begin
          Tobl_Rec.PutValue('GL_PUTTCDEV'    , Tobl_Cde.GetValue ('GL_PUTTCDEV'   )) ;
          Tobl_Rec.PutValue('GL_PUTTCNETDEV' , Tobl_Cde.GetValue ('GL_PUTTCNETDEV')) ;
          Tobl_Rec.PutValue('GL_REMISELIGNE' , Tobl_Cde.GetValue ('GL_REMISELIGNE')) ;
          Tobl_Rec.PutValue('GL_REMISEPIED'  , Tobl_Cde.GetValue ('GL_REMISEPIED' )) ;
          Tobl_Rec.PutValue('GL_ESCOMPTE'    , Tobl_Cde.GetValue ('GL_ESCOMPTE'   )) ;
          Tobl_Rec.PutValue('GL_DPA'         , Tobl_Cde.GetValue ('GL_DPA'        )) ;
          Tobl_Rec.PutValue('GL_PMAP'        , Tobl_Cde.GetValue ('GL_PMAP'       )) ;
          Tobl_Rec.PutValue('GL_PMRP'        , Tobl_Cde.GetValue ('GL_PMRP'       )) ;
          Tobl_Rec.PutValue('GL_PMAPACTU'    , Tobl_Cde.GetValue ('GL_PMAPACTU'   )) ;
          Tobl_Rec.PutValue('GL_PMRPACTU'    , Tobl_Cde.GetValue ('GL_PMRPACTU'   )) ;
          Tobl_Rec.PutValue('GL_DPR'         , Tobl_Cde.GetValue ('GL_DPR'        )) ;

          // Mise à jour de la ligne générique
          if Tobl_Rec.GetValue ('GL_TYPEDIM') = 'DIM' then
          begin
            TOBL_RecGen := Tob_Piece.FindFirst(['$$_REFINTERNE','GL_CODESDIM'],[RefInterne,Tobl_Rec.GetValue ('GL_CODEARTICLE')],False) ;
            if TOBL_RecGen<>Nil then
            begin
              TOBL_RecGen.PutValue('GL_PUTTCDEV'      , Tobl_Rec.GetValue ('GL_PUTTCDEV')) ;
              TOBL_RecGen.PutValue('GL_PUTTCNETDEV'   , Tobl_Rec.GetValue ('GL_PUTTCNETDEV')) ;
              TOBL_RecGen.PutValue('GL_PUTTC'         , Tobl_Rec.GetValue ('GL_PUTTC')) ;
              TOBL_RecGen.PutValue('GL_PUTTCNET'      , Tobl_Rec.GetValue ('GL_PUTTCDEV'));
              TOBL_RecGen.PutValue('GL_MONTANTTTCDEV' , TOBL_RecGen.GetValue('GL_MONTANTTTCDEV') + Tobl_Rec.GetValue ('GL_PUTTCNETDEV') * Tobl_Rec.GetValue ('GL_QTEFACT'));
              TOBL_RecGen.PutValue('GL_MONTANTTTC'    , TOBL_RecGen.GetValue('GL_MONTANTTTC')    + Tobl_Rec.GetValue ('GL_PUTTCNET')    * Tobl_Rec.GetValue ('GL_QTEFACT'));
            end;
          end;
        end;
      end;
    end;
  end;

  // Il faut recalculer le GL_QTERESTE des lignes génériques de la pièce à solder.
  RecalculLigneGenerique (Tob_Piece_Prec_Old);

  // Recalcul de la pièce à solder -> Inutile avec la nouvelle gestion
  // des reliquats, la pièce initiale conserve ses montants
  // CalculFacture (Tob_Piece_Prec_Old, Tob_PiedBase_Prec_Old, Tob_Tiers, Tob_Article_Prec_New, nil, nil, nil, DEV) ;
  // Mise à jour du stock si reliquat
  if EncoreReliquat then
  begin
    if MajStockPiece then
    begin
      // Rechargement des articles + stock
      Tob_Article_Prec_New := TOB.CREATE ('Les Articles', nil, -1);
      ChargeLesArticles (Tob_Piece_Prec_Old, Tob_Article_Prec_New);
      MajStockPieceASolder (Tob_Piece_Prec_Old, Tob_Article_Prec_New);
      FreeAndNil( Tob_Article_Prec_New );
    end;
  end else
  begin
    Tob_Piece_Prec_Old.PutValue ('GP_VIVANTE', '-') ;
    for i:=0 to Tob_Piece_Prec_Old.Detail.Count-1 do
    begin
      Tob_Piece_Prec_Old.detail [i].PutValue ('GL_VIVANTE', '-') ;
    end;
  end;

  // MAJ DB de la pièce soldée
  Tob_Piece_Prec_Old.UpdateDB (False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Annulation du stock de la pièce avant suppression du
Suite ........ : contenu
Mots clefs ... : TOX;ANNULE ET REMPLACE
*****************************************************************}
function TTImportGPAO.AnnuleStockPieceASolder ( TobPieceSolde : TOB ; TobArticle : TOB) : integer ;
var TobEntete       : TOB      ;
    TobNomenclature : TOB      ;
    i               : integer  ;
begin
  V_PGI.IoError:=oeOk;
  result := 0 ;
  /////////////////////////////////////////////////////////////////////////
  // Duplication de la TOB piece pour traitement
  // -> on récupère l'entête
  /////////////////////////////////////////////////////////////////////////
  TobEntete := TOB.Create ('PIECE', nil, -1);
  TobEntete.Dupliquer (TobPieceSolde, True, True, True);
  /////////////////////////////////////////////////////////////////////////
  // Ré-initialisation du QTESTOCK par QTERESTE pour bonne MAJ de dispo
  /////////////////////////////////////////////////////////////////////////
  for i:=0 to TobEntete.Detail.Count-1 do
  begin
    TobEntete.detail [i].PutValue ('GL_QTESTOCK', TobEntete.detail [i].GetValue ('GL_QTERESTE'));
    TobEntete.detail [i].PutValue ('GL_QTEFACT' , TobEntete.detail [i].GetValue ('GL_QTERESTE'));
  end;
  /////////////////////////////////////////////////////////////////////////
  // Chargement des nomenclatures
  /////////////////////////////////////////////////////////////////////////
  TobNomenclature := TOB.CREATE ('Les Nomens', nil, -1);
  ////////////////////////////////////////////////////////////////////////////
  // Calcul la mise à jour de stock
  ////////////////////////////////////////////////////////////////////////////
  InverseStock(TobEntete, TobArticle, Nil, TobNomenclature) ;

  if V_PGI.IoError<>oeOK then
    EcrireInfo (TraduireMemoire('  Attention : problème pour mettre à jour le stock à partir de la pièce solder !'), True, False, False)
  else
    ValideLesArticles(TobEntete, TOBArticle) ;

  //////////////////////////////////////////////////////////////////////////////
  // Libération des TOB
  //////////////////////////////////////////////////////////////////////////////
  FreeAndNil( TobEntete );
  FreeAndNil( TobNomenclature );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Annulation du stock de la pièce avant suppression du
Suite ........ : contenu
Mots clefs ... : TOX;ANNULE ET REMPLACE
*****************************************************************}
function TTImportGPAO.MajStockPieceASolder ( TobPieceSolde : TOB ; TobArticle : TOB) : integer ;
var TobEntete  : TOB      ;
    Tob_Cat    : TOB      ;
    Tob_Nom    : TOB      ;
    i          : integer  ;
begin
  V_PGI.IoError:=oeOk;
  result := 0 ;
  /////////////////////////////////////////////////////////////////////////
  // Duplication de la TOB piece pour traitement
  // -> on récupère l'entête
  /////////////////////////////////////////////////////////////////////////
  TobEntete := TOB.Create ('PIECE', nil, -1);
  TobEntete.Dupliquer (TobPieceSolde, True, True, True);
  /////////////////////////////////////////////////////////////////////////
  // Ré-initialisation du QTESTOCK par QTERESTE pour bonne MAJ de dispo
  /////////////////////////////////////////////////////////////////////////
  for i:=0 to TobEntete.Detail.Count-1 do
  begin
    TobEntete.detail [i].PutValue ('GL_QTESTOCK', TobEntete.detail [i].GetValue ('GL_QTERESTE'));
    TobEntete.detail [i].PutValue ('GL_QTEFACT' , TobEntete.detail [i].GetValue ('GL_QTERESTE'));
  end;
  /////////////////////////////////////////////////////////////////////////
  // MAJ des stocks
  /////////////////////////////////////////////////////////////////////////
  Tob_Cat := TOB.Create ('LE CAT'   , nil, -1);
  Tob_Nom := TOB.Create ('LES NOMEN', nil, -1);
  ValideLesLignes   (TobEntete, TobArticle, Tob_Cat, Tob_Nom, nil, nil, nil, False, False) ;
  ValideLesArticles (TobEntete, TobArticle) ;
  //////////////////////////////////////////////////////////////////////////////
  // Libération des TOB
  //////////////////////////////////////////////////////////////////////////////
  FreeAndNil( TobEntete );
  FreeAndNil( Tob_Nom   );
  FreeAndNil( Tob_Cat   );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : Recalcul la zone GL_QTERESTE des lignes génériques
Suite ........ : contenu
Mots clefs ... : TOX;ANNULE ET REMPLACE
*****************************************************************}
procedure TTImportGPAO.RecalculLigneGenerique ( TobPieceSolde : TOB )  ;
var  Tob_Ligne    : TOB     ;
     Tob_LigneDim : TOB     ;
     CodeArticle  : string  ;
     rc, i        : integer ;
     NumLigne     : integer ;
     QteReste     : double  ;
begin

  for i:=0 to TobPieceSolde.Detail.Count-1 do
  begin
    Tob_Ligne := TobPieceSolde.Detail[i];

    if Tob_Ligne.GetValue('GL_TYPEDIM') = 'GEN' then
    begin
      CodeArticle := Tob_Ligne.GetValue('GL_CODESDIM');
      Numligne    := Tob_Ligne.GetValue('GL_NUMLIGNE');
      rc          := 0 ;
      QteReste    := 0   ;

      while (rc = 0) do
      begin
        inc (NumLigne);
//      Tob_LigneDim := nil ;

        Tob_LigneDim := TobPieceSolde.FindFirst(['GL_NUMLIGNE'],[NumLigne],False) ;

        if (Tob_LigneDim <> Nil) and (Tob_LigneDim.GetValue('GL_TYPEDIM') = 'DIM') and (Tob_LigneDim.GetValue('GL_CODEARTICLE')=CodeArticle) then
         QteReste := QteReste + Tob_LigneDim.GetValue ('GL_QTERESTE')
        else
        begin
          Tob_Ligne.PutValue('GL_QTERESTE', QteReste);
          rc := -1;
        end;
      end;
     end;
  end;
end;


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
procedure TTImportGPAO.TraiteLesreliquats (Tob_Piece, Tob_Article : TOB);
var NaturePiece    : string ;
    RefInterne     : string ;
    OldNaturePiece : string ;
    OldRefInterne  : string ;
    i              : integer;
begin

  for i:=0 to Tob_Piece.detail.Count-1 do
  begin
    NaturePiece := Tob_Piece.detail[i].GetValue ('$$_NATUREPIECEG');
    RefInterne  := Tob_Piece.detail[i].GetValue ('$$_REFINTERNE');

    if ((NaturePiece <> '') and (NaturePiece <> OldNaturePiece)) or ((Refinterne<>'') and (RefInterne <> OldRefInterne)) then
    begin
      OldNaturePiece := NaturePiece ;
      OldRefInterne  := RefInterne  ;
      EncoreReliquat := False       ;

      // Charge en TOB la pièce précédente, les pieds, les articles, les stocks , ....
      EcrireInfo (TraduireMemoire('  Chargement de la pièce à solder référence ') + RefInterne, True, False, False);
      ChargeLaPiecePrecedente (NaturePiece, RefInterne);

      // Mise à jour de la zone GL_QTERESTE = GL_QTESTOCK-Qté reçue
      EcrireInfo (TraduireMemoire('  Calcul du reliquat de la pièce référence ') + RefInterne, True, False, False);
      MetAJourQteReste (RefInterne);

      // Recopie Zone Adresse de Pièce Précédente sur nouvelle pièce
      Tob_Piece.PutValue( 'GP_NUMADRESSELIVR', Tob_Piece_Prec_Old.GetValue('GP_NUMADRESSELIVR'));
      Tob_Piece.PutValue( 'GP_NUMADRESSEFACT', Tob_Piece_Prec_Old.GetValue('GP_NUMADRESSEFACT'));

      // Recopie Zone Libre de Pièce Précédente sur nouvelle pièce
      Tob_Piece.PutValue( 'GP_LIBRETIERS1', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS1'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS2', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS2'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS3', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS3'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS4', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS4'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS5', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS5'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS6', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS6'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS7', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS7'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS8', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS8'));
      Tob_Piece.PutValue( 'GP_LIBRETIERS9', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERS9'));
      Tob_Piece.PutValue( 'GP_LIBRETIERSA', Tob_Piece_Prec_Old.GetValue('GP_LIBRETIERSA'));
//      Tob_Piece.PutValue( 'GP_MAJLIBRETIERS', Tob_Piece_Prec_Old.GetValue('GP_MAJLIBRETIERS'));
      Tob_Piece.PutValue( 'GP_LIBREPIECE1', Tob_Piece_Prec_Old.GetValue('GP_LIBREPIECE1'));
      Tob_Piece.PutValue( 'GP_LIBREPIECE2', Tob_Piece_Prec_Old.GetValue('GP_LIBREPIECE2'));
      Tob_Piece.PutValue( 'GP_LIBREPIECE3', Tob_Piece_Prec_Old.GetValue('GP_LIBREPIECE3'));
      Tob_Piece.PutValue( 'GP_DATELIBREPIECE1', Tob_Piece_Prec_Old.GetValue('GP_DATELIBREPIECE1'));
      Tob_Piece.PutValue( 'GP_DATELIBREPIECE2', Tob_Piece_Prec_Old.GetValue('GP_DATELIBREPIECE2'));
      Tob_Piece.PutValue( 'GP_DATELIBREPIECE3', Tob_Piece_Prec_Old.GetValue('GP_DATELIBREPIECE3'));
      Tob_Piece.PutValue( 'GP_LIBREAFF1', Tob_Piece_Prec_Old.GetValue('GP_LIBREAFF1'));
      Tob_Piece.PutValue( 'GP_LIBREAFF2', Tob_Piece_Prec_Old.GetValue('GP_LIBREAFF2'));
      Tob_Piece.PutValue( 'GP_LIBREAFF3', Tob_Piece_Prec_Old.GetValue('GP_LIBREAFF3'));


      // Libération des TOB
      FreeAndNil( Tob_Piece_Prec_Old    );
      FreeAndNil( Tob_Article_Prec_Old  );
      FreeAndNil( Tob_PiedBase_Prec_Old );
      FreeAndNil( Tob_PiedEche_Prec_Old );
      FreeAndNil( Tob_Article_Prec_New  );
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI
Suite ........ :  1 - Création des pieds de document
Suite ........ :  2 - Recalcul de la pièce
Suite ........ :  3 - Import DB
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.ControlePrixNul ;       // Mettre des controles de cohérence
var i       : integer ;
    Prix    : double  ;
    Tob_Gen : TOB     ;
begin

  for i:=0 to Tob_Piece.Detail.Count-1 do
  begin
    if Tob_Piece.detail [i].GetValue ('GL_TYPELIGNE') = 'ART' then
    begin
      Prix := 0 ;
      if (Tob_Piece.GetValue ('GP_FACTUREHT') = 'X') then
      begin
        if Tob_Piece.detail [i].GetValue ('GL_PUHTDEV') = 0.0 then
        begin
          if VenteAchat = 'ACH' then Prix := RecherchePAHT (Tob_Piece.detail [i].GetValue ('GL_ARTICLE'))
          else if VenteAchat = 'VEN' then Prix := RecherchePVHT (Tob_Piece.detail [i].GetValue ('GL_ARTICLE'));
          if Prix <> 0 then
          begin
            Tob_Piece.detail [i].PutValue ('GL_PUHTDEV', Prix) ;
            //
            // Mise à jour de la ligne générique
            //
            Tob_Gen := Tob_Piece.FindFirst(['GL_CODESDIM'],[Tob_Piece.detail [i].GetValue ('GL_CODEARTICLE')],False) ;
            if Tob_Gen <> nil then
            begin
              Numligne := Tob_Gen.GetValue ('GL_NUMLIGNE');
              Tob_Gen.PutValue('GL_PUHTDEV'      , Tob_Piece.detail [i].GetValue ('GL_PUHTDEV')) ;
              Tob_Gen.PutValue('GL_PUHTNETDEV'   , Tob_Piece.detail [i].GetValue ('GL_PUHTNETDEV')) ;
              Tob_Gen.PutValue('GL_PUHT'         , Tob_Piece.detail [i].GetValue ('GL_PUHT')) ;
              Tob_Gen.PutValue('GL_PUHTNET'      , Tob_Piece.detail [i].GetValue ('GL_PUHTDEV')) ;
              Tob_Gen.PutValue('GL_MONTANTHTDEV' , Tob_Gen.GetValue ('GL_MONTANTHTDEV') + Tob_Piece.detail [i].GetValue ('GL_PUHTNETDEV') * Tob_Piece.detail [i].GetValue ('GL_QTEFACT'));
              Tob_Gen.PutValue('GL_MONTANTHT'    , Tob_Gen.GetValue ('GL_MONTANTHT')    + Tob_Piece.detail [i].GetValue ('GL_PUHTNET')    * Tob_Piece.detail [i].GetValue ('GL_QTEFACT'));
            end;
          end;
        end;
      end else
      begin
        if Tob_Piece.detail [i].GetValue ('GL_PUTTCDEV') = 0.0 then
        begin
          if VenteAchat = 'VEN' then Prix := RecherchePVTTC (Tob_Piece.detail [i].GetValue ('GL_ARTICLE')) ;
          if Prix <> 0 then
          begin
            Tob_Piece.detail [i].PutValue ('GL_PUTTCDEV', Prix) ;
            //
            // Mise à jour de la ligne générique
            //
            Tob_Gen := Tob_Piece.FindFirst(['GL_CODESDIM'],[Tob_Piece.detail [i].GetValue ('GL_CODEARTICLE')],False) ;
            if Tob_Gen <> nil then
            begin
              Numligne := Tob_Gen.GetValue ('GL_NUMLIGNE');
              if Tob_Gen <> nil then
              begin
                Numligne := Tob_Gen.GetValue ('GL_NUMLIGNE');
                Tob_Gen.PutValue('GL_PUTTCDEV'      , Tob_Piece.detail [i].GetValue ('GL_PUTTCDEV')) ;
                Tob_Gen.PutValue('GL_PUTTCNETDEV'   , Tob_Piece.detail [i].GetValue ('GL_PUTTCNETDEV')) ;
                Tob_Gen.PutValue('GL_PUTTC'         , Tob_Piece.detail [i].GetValue ('GL_PUTTC')) ;
                Tob_Gen.PutValue('GL_PUTTCNET'      , Tob_Piece.detail [i].GetValue ('GL_PUTTCDEV')) ;
                Tob_Gen.PutValue('GL_MONTANTTTCDEV' , Tob_Gen.GetValue ('GL_MONTANTTTCDEV') + Tob_Piece.detail [i].GetValue ('GL_PUHTNETDEV') * Tob_Piece.detail [i].GetValue ('GL_QTEFACT'));
                Tob_Gen.PutValue('GL_MONTANTTTC'    , Tob_Gen.GetValue ('GL_MONTANTTTC')    + Tob_Piece.detail [i].GetValue ('GL_PUHTNET')    * Tob_Piece.detail [i].GetValue ('GL_QTEFACT'));
              end;
            end;
          end;
        end;
      end;
    end;
  end;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Import d'une pièce PGI
Suite ........ :  1 - Création des pieds de document
Suite ........ :  2 - Recalcul de la pièce
Suite ........ :  3 - Import DB
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.ImportPiece ;       // Mettre des controles de cohérence
var i : integer ;
begin
  //////////////////////////////////////////////////////////////////
  // Création de la dernière ligne commentaire
  //////////////////////////////////////////////////////////////////
  if (LigneCommACreer = True) then
  begin
    LigneCommACreer:=False;
    CreeLigneCommentaire;
  end;
  //////////////////////////////////////////////////////////
  // Sauvegarde Chrono
  //////////////////////////////////////////////////////////
  // IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce
  //////////////////////////////////////////////////////////
  // Gestion des reliquats sur les pièces initiales
  //////////////////////////////////////////////////////////
  if PieceASolder then TraiteLesreliquats (Tob_Piece, Tob_Article) ;
  //////////////////////////////////////////////////////////
  // Contrôle éventuel des prix unitaires
  //////////////////////////////////////////////////////////
  if CtrlPrixNul then ControlePrixNul ;
  //////////////////////////////////////////////////////////
  // Calcul des cumuls de la pièce
  // Calcul de PIEDBASE
  //////////////////////////////////////////////////////////
  PutValueDetail(Tob_Piece, 'GP_RECALCULER', 'X') ;
  CalculFacture (Tob_Piece, Tob_PiedBase, Tob_Tiers, Tob_Article, nil, nil, nil, DEV) ;
  //////////////////////////////////////////////////////////
  // Génération des PIEDECHE
  //////////////////////////////////////////////////////////
  if GenereEcheances then GereEcheancesGC (Tob_Piece, Tob_Tiers, Tob_PiedEche, Tob_Acomptes, nil, taCreat, DEV, False);
  //////////////////////////////////////////////////////////////////////////////
  // Cas spécifique des tickets FO : il faut renseigner la zone GPE_MONTANTENCAIS
  //////////////////////////////////////////////////////////////////////////////
  if Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'FFO' then
  begin
    for i:=0 to Tob_PiedEche.detail.Count-1 do
    begin
      Tob_PiedEche.detail[i].PutValue ('GPE_MONTANTENCAIS', Tob_PiedEche.detail[i].GetValue ('GPE_MONTANTDEV'));
      if ModePaiement <> '' then Tob_PiedEche.detail[i].PutValue ('GPE_MODEPAIE', ModePaiement) ;
    end;
  end;

  ///////////////////////////////////////////////////////////////
  // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES
  ///////////////////////////////////////////////////////////////
//  Tob_Piece.SetAllModifie(TRUE) ;
//  Tob_PiedBase.SetAllModifie(TRUE) ;
//  Tob_PiedEche.SetAllModifie(TRUE) ;
  //////////////////////////////////////////////////////////////////
  // Gestion des adresses
  //////////////////////////////////////////////////////////////////
  Tob_Adresses:= TOB.Create ('LES ADRESSES',Nil,-1) ;
  if PieceASolder then
  begin
    // Reprise des adresses de la pièce précédente
    Tob_Adresses.Dupliquer( Tob_Adresses_Prec_Old, True, True, False );
    for i:=0 to Tob_Adresses.Detail.Count-1 do
    begin
      Tob_Adresses.Detail[i].PutValue( 'GPA_NATUREPIECEG', Tob_Piece.GetValue( 'GP_NATUREPIECEG' ));
      Tob_Adresses.Detail[i].PutValue( 'GPA_SOUCHE'      , Tob_Piece.GetValue( 'GP_SOUCHE'  ));
      Tob_Adresses.Detail[i].PutValue( 'GPA_NUMERO'      , Tob_Piece.GetValue( 'GP_NUMERO'  ));
      Tob_Adresses.Detail[i].PutValue( 'GPA_INDICEG'     , Tob_Piece.GetValue( 'GP_INDICEG' ));
      Tob_Adresses.Detail[i].PutValue( 'GPA_NUMLIGNE'    , 0);
    end;
  end else
  begin
    TOB.Create('PIECEADRESSE', Tob_Adresses,-1) ; {Livraison}
    TOB.Create('PIECEADRESSE', Tob_Adresses,-1) ; {Facturation}
    ValideLesAdresses ( Tob_Piece, Tob_Piece, Tob_Adresses) ;
    // for i:=0 to Tob_Adresses.Detail.Count-1 do GetAdrFromTOB (Tob_Adresses.Detail[i], TOB_Tiers) ;
    for i:=0 to Tob_Adresses.Detail.Count-1 do GetAdrFromCode(Tob_Adresses.Detail[i], Tob_Tiers.GetValue ('T_TIERS'), False);
  end;

  ////////////////////////////////////////////////////////////////////
  // MAJ du stock en cas de besoin
  ////////////////////////////////////////////////////////////////////
  if MajStockPiece then
  begin
    ValideLesLignes   (Tob_Piece, Tob_Article, Tob_Catalogue, Tob_Noment, nil, nil, nil, False, False) ;
    ValideLesArticles (Tob_Piece, Tob_Article) ;
  end;
  //////////////////////////////////////////////////////////////
  // Duplication de la piece
  //////////////////////////////////////////////////////////////
  if NatureTrfRecu <> '' then DupliqueTrfRecu ;

  /////////////////////////////////////////////////////////////////////////////////////////////////
  // MAJ des TOB -> DB de la piece
  ///////////////////////////////////////////////////////////////////////////////////////////////
  Tob_Piece.InsertDB (nil);
  Tob_PiedBase.InsertorUpdateDB (False);
  Tob_PiedEche.InsertorUpdateDB (False);
  Tob_Adresses.InsertorUpdateDB (False);

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  FreeAndNil( Tob_Piece     );  // Libère TOB Pièce+Lignes
  FreeAndNil( Tob_Article   );  // Libère la TOB des articles de la pièce
  FreeAndNil( Tob_Tiers     );  // Libère la TOB du fournisseurs
  FreeAndNil( Tob_PiedBase  );  // Libère la TOB PiedBase
  FreeAndNil( Tob_PiedEche  );  // Libère la TOB PiedEche
  FreeAndNil( Tob_Adresses  );  // Libère la TOB Adresses
  FreeAndNil( Tob_Acomptes  );  // Libère la TOB Acomptes
  FreeAndNil( Tob_Catalogue );  // Libère la TOB catalogues
  FreeAndNil( Tob_Noment    );  // Libère la TOB des nomenclatures
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /    
Description .. : Initialisation de la TOB entête de document
Mots clefs ... : 
*****************************************************************}
function TTImportGPAO.InitEntetePiece (TobInfo : TOB) : boolean;
var NaturePiece : string   ;
    Etab        : string   ;                              
    CodeTiers   : string   ;
    DevisePiece : string   ;
    CodeTva     : string   ;
    Q           : TQUERY   ;
    DD          : TDateTime;
    Per,Sem,i   : integer  ;

    procedure RechTiers;
    begin
      Q := OpenSQL('Select * from Tiers Where T_TIERS="'+CodeTiers+'"' ,False) ;
    end;

begin
  result := True ;
  //////////////////////////////////////////////////////////////////////
  // Nouvelle pièce :  Initialisations diverses et varièes
  //////////////////////////////////////////////////////////////////////
  TotQte       := 0.0;
  TotPrix      := 0.0;
  TotPrixNet   := 0.0 ;
  Numligne     := 0  ;            // Initialisation du premier numéro de ligne
  PieceASolder := False ;

  NaturePiece := TobInfo.GetValue ('GP_NATUREPIECEG');
  Etab        := TobInfo.GetValue ('GP_ETABLISSEMENT');

  QualifMvt          := GetInfoParPiece(NaturePiece,'GPP_QUALIFMVT') ;
  VenteAchat         := GetInfoParPiece(NaturePiece,'GPP_VENTEACHAT') ;
  LigneCommACreer    := False;
  OldRefArticle      := '' ;      //NA 28/03/2002
  NatureLigneCom     := '' ;
  RefInterneLigneCom := '' ;
  OldNaturePiece     := '' ;
  OldRefInterne      := '' ;

  Tob_Piece:= TOB.CREATE ('PIECE', NIL, -1);
  AddLesSupEntete (Tob_Piece) ;
  InitTobPiece (Tob_Piece);			// Init date de création, modif, utilisateur ......
  Tob_Piece.PutValue  ('GP_NATUREPIECEG' , NaturePiece);
  Tob_Piece.PutValue  ('GP_DATEPIECE'    , TobInfo.GetValue ('GP_DATEPIECE'));

  if TobInfo.GetValue ('GP_DATELIVRAISON')=iDate1900 then Tob_Piece.PutValue ('GP_DATELIVRAISON', TobInfo.GetValue ('GP_DATEPIECE'))
  else Tob_Piece.PutValue ('GP_DATELIVRAISON', TobInfo.GetValue ('GP_DATELIVRAISON'));
  /////////////////////////////
  // Modification pour JACADI
  ///////////////////////////
  if (TobInfo.FieldExists ('$$_MAJDATEINT')) and (TobInfo.GetValue ('$$_MAJDATEINT') = 'X') then
  begin
    Tob_Piece.PutValue('GP_DATEINTEGR', NowH);
  end;

  /////////////////////////
  // Souche, Numéro, .....
  ////////////////////////
  CleDoc.Souche      := GetSoucheG(NaturePiece, Etab, '') ;
  CleDoc.NumeroPiece := GetNumSoucheG(CleDoc.Souche) ;
  CleDoc.DatePiece   := Tob_Piece.GetValue('GP_DATEPIECE');

  PutValueDetail(Tob_Piece, 'GP_SOUCHE' ,CleDoc.Souche) ;
  PutValueDetail(Tob_Piece, 'GP_NUMERO' ,CleDoc.NumeroPiece) ;
  PutValueDetail(Tob_Piece, 'GP_INDICEG', 0) ;
  //
  // MODIF LM 20/05/2003 pour O'NEIL et conflit de numéros
  // Dès le départ, on réserve le numéro de pièce
  //
  IncNumSoucheG (Tob_Piece.GetValue ('GP_SOUCHE')) ;  	// Incrémentation du numéro de pièce
  ////////////////////////////////////////////////////////////////////////
  // Référence de la pièce
  ///////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_REFINTERNE', TobInfo.GetValue ('GP_REFINTERNE'));
  Tob_Piece.PutValue ('GP_REFEXTERNE', TobInfo.GetValue ('GP_REFEXTERNE'));
  /////////////////////////////////////////////////////////////////////////
  // Remise pied / Escompte
  /////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_REMISEPIED', TobInfo.GetValue ('GP_REMISEPIED'));
  Tob_Piece.PutValue ('GP_ESCOMPTE'  , TobInfo.GetValue ('GP_ESCOMPTE'));
  /////////////////////////////////////////////////////////////////////////
  // Etablissement / Dépôt
  /////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_ETABLISSEMENT', TobInfo.GetValue ('GP_ETABLISSEMENT'));
  Tob_Piece.PutValue ('GP_DEPOT'        , TobInfo.GetValue ('GP_ETABLISSEMENT'));

  ///////////////////////////////////////////////////////////////
  // Cas du client 999999 du CASH
  ///////////////////////////////////////////////////////////////
  CodeTiers := TobInfo.GetValue ('GP_TIERS');
  if ((CodeTiers = '999999') or (CodeTiers = Etab + '999999')) and
     (TobInfo.FieldExists ('$$_CREECLICASH')) and (TobInfo.GetValue ('$$_CREECLICASH') = 'X') then
  begin
    CodeTiers := DernierClientCree ;
  end ;

  ///////////////////////////////////////////////////////////////
  // Chargement en TOB du client
  ///////////////////////////////////////////////////////////////
  if (CodeTiers = ValTiersDefaut) and (Tob_TiersDefaut <> nil) then
  begin
    Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
    Tob_Tiers.Dupliquer( Tob_TiersDefaut, False, True, True );
  end else
  begin
    RechTiers;
    //  Q := OpenSQL('Select * from Tiers Where T_TIERS="'+CodeTiers+'"' ,False) ;
    if not Q.EOF then
    begin
      Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
      Tob_Tiers.SelectDB('',Q);
      if CodeTiers = ValTiersDefaut then
      begin
        Tob_TiersDefaut := TOB.CREATE ('TIERS', NIL, -1);
        Tob_TiersDefaut.Dupliquer( Tob_Tiers, False, True, True );
      end;
      Ferme(Q) ;
    end else
    begin
      Ferme(Q) ;

      // Le tiers n'existe pas, peut-on l'initialiser avec le tiers par défaut ??
      if (TobInfo.FieldExists ('$$_FORCETIERS')) and (TobInfo.GetValue ('$$_FORCETIERS') = 'X') then
      begin
        CodeTiers := GetParamSoc('SO_GCTOXTIERS') ;
        RechTiers;
        //Q := OpenSQL('Select * from Tiers Where T_TIERS="'+CodeTiers+'"' ,False) ;
        if not Q.EOF then
        begin
          EcrireInfo (TraduireMemoire('   PIECE référence ') + Tob_Piece.GetValue('GP_REFINTERNE') + TraduireMemoire(' -> Attention, le tiers est inconnu. On reprend le tiers par défaut ')+ CodeTiers, True, True, True);
          Tob_Tiers := TOB.CREATE ('TIERS', NIL, -1);
          Tob_Tiers.SelectDB('',Q);
          Ferme(Q) ;
        end else
        begin
          Ferme(Q) ;
          EcrireInfo (TraduireMemoire('   PIECE référence ') + Tob_Piece.GetValue('GP_REFINTERNE') + TraduireMemoire(' -> Attention, le tiers est inconnu et impossible de charger le tiers par défaut ')+ CodeTiers, True, True, True);
          result := False ;
        end;
      end else
      begin
        EcrireInfo (TraduireMemoire('   PIECE référence ') + Tob_Piece.GetValue('GP_REFINTERNE') + TraduireMemoire(' -> Problème pour retrouver le tiers ')+ CodeTiers, True, True, True);
        result := False ;
      end;
    end;
  end;

  if result = True then
  begin
    Tob_Piece.PutValue      ('GP_TIERS', CodeTiers);
    if (TobInfo.FieldExists ('GP_TIERSLIVRE')) and (TobInfo.GetValue ('GP_TIERSLIVRE')<>'') then Tob_Piece.PutValue ('GP_TIERSLIVRE', TobInfo.GetValue ('GP_TIERSLIVRE'))
    else Tob_Piece.PutValue ('GP_TIERSLIVRE', CodeTiers) ;
    if (TobInfo.FieldExists ('GP_TIERSFACTURE')) and (TobInfo.GetValue ('GP_TIERSFACTURE')<>'') then Tob_Piece.PutValue ('GP_TIERSFACTURE', TobInfo.GetValue ('GP_TIERSFACTURE'))
    else Tob_Piece.PutValue ('GP_TIERSFACTURE', CodeTiers) ;
    if (TobInfo.FieldExists ('GP_TIERSPAYEUR')) and (TobInfo.GetValue ('GP_TIERSPAYEUR')<>'') then Tob_Piece.PutValue ('GP_TIERSPAYEUR', TobInfo.GetValue ('GP_TIERSPAYEUR'))
    else Tob_Piece.PutValue ('GP_TIERSPAYEUR', CodeTiers) ;
    ////////////////////////////////////////////////////////////////////////////
    // Initialise les zones de l'entête en fonction du client
    ////////////////////////////////////////////////////////////////////////////
    if Tob_Piece.GetValue ('GP_ESCOMPTE')=0 then Tob_Piece.PutValue ('GP_ESCOMPTE', Tob_Tiers.GetValue ('T_ESCOMPTE')) ;
    Tob_Piece.PutValue ('GP_MODEREGLE'      , Tob_Tiers.GetValue ('T_MODEREGLE')) ;
    Tob_Piece.PutValue ('GP_TVAENCAISSEMENT', Tob_Tiers.GetValue ('T_TVAENCAISSEMENT')) ;
    Tob_Piece.PutValue ('GP_QUALIFESCOMPTE' , Tob_Tiers.GetValue ('T_QUALIFESCOMPTE')) ;
    Tob_Piece.PutValue ('GP_TARIFTIERS'     , Tob_Tiers.GetValue ('T_TARIFTIERS')) ;
    /////////////////////////////////////////////////////////////////////////
    // Gestion de la TVA
    /////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_CODETVA') then
    begin
       CodeTva := TobInfo.GetValue ('$$_CODETVA') ;
       Tob_Piece.PutValue ('GP_REGIMETAXE', ReadTokenSt (CodeTva)) ;
    end ;
    if Tob_Piece.GetValue ('GP_REGIMETAXE') = '' then Tob_Piece.PutValue ('GP_REGIMETAXE', Tob_Tiers.GetValue ('T_REGIMETVA')) ;
    Tob_Piece.PutValue ('GP_TVAENCAISSEMENT', PositionneExige ( Tob_Tiers));
    /////////////////////////////////////////////////////////////////////////////
    // Facturation HT ou TTC
    /////////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_FACTUREHT') then Tob_Piece.PutValue ('GP_FACTUREHT', TobInfo.GetValue ('$$_FACTUREHT'))
    else Tob_Piece.PutValue ('GP_FACTUREHT', Tob_Tiers.GetValue ('T_FACTUREHT'));
    /////////////////////////////////////////////////////////////////////////////
    // Représentant
    /////////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_REPRESENTANT') then Tob_Piece.PutValue ('GP_REPRESENTANT', TobInfo.GetValue ('$$_REPRESENTANT')) ;
    //NA 07/03/2002 else Tob_Piece.PutValue ('GP_REPRESENTANT'  , ChoixCommercial('VEN',Tob_Tiers.GetValue('T_ZONECOM'))) ;

    ///////////////////////////////////////////////
    // Champs par  défaut
    //////////////////////////////////////////////
    // MODIF LM pour 12/11/2001
    //if Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'FFO' then Tob_Piece.PutValue ('GP_VIVANTE' , '-')
    //else Tob_Piece.PutValue ('GP_VIVANTE'       , 'X');     // Pièce vivante    : Oui
    Tob_Piece.PutValue ('GP_VIVANTE', TobInfo.GetValue ('GP_VIVANTE')) ;    //NA 07/03/2002

    Tob_Piece.PutValue ('GP_ETATVISA'      , 'NON');     // Pièce visée      : Non
    Tob_Piece.PutValue ('GP_CREEPAR'       , 'SAI');     // Mode de creation : Saisie
    //Tob_Piece.PutValue ('GP_VENTEACHAT', TobInfo.GetValue ('GP_VENTEACHAT')) ;    // Document         : Vente/Achat
    if TobInfo.FieldExists ('GP_VENTEACHAT') then VenteAchat := TobInfo.GetValue ('GP_VENTEACHAT') else VenteAchat := '' ;
    if VenteAchat = '' then VenteAchat := GetInfoParPiece(NaturePiece,'GPP_VENTEACHAT');
    Tob_Piece.PutValue ('GP_VENTEACHAT', VenteAchat);                               //NA 07/03/2002
    Tob_Piece.PutValue ('GP_EDITEE'        , 'X');       // Pièce éditée     : Oui
    Tob_Piece.PutValue ('GP_NUMADRESSELIVR', 0);
    Tob_Piece.PutValue ('GP_NUMADRESSEFACT', 0);
    Tob_Piece.PutValue ('GP_UTILISATEUR'   , V_PGI.User) ;
    Tob_Piece.PutValue ('GP_CREATEUR'      , V_PGI.User) ;
    Tob_Piece.PutValue ('GP_SOCIETE'       , V_PGI.CodeSociete);
    Tob_Piece.PutValue('GP_DATECREATION'   , Date) ;
    Tob_Piece.PutValue('GP_HEURECREATION'  , NowH) ;
    if TobInfo.FieldExists ('GP_CAISSE') then Tob_Piece.PutValue ('GP_CAISSE', TobInfo.GetValue ('GP_CAISSE')) ;          //NA 07/03/2002
    if TobInfo.FieldExists ('GP_DEPOTDEST') then Tob_Piece.PutValue ('GP_DEPOTDEST', TobInfo.GetValue ('GP_DEPOTDEST')) ; //NA 07/03/2002

    /////////////////////////
    // Devise du document
    ////////////////////////
    if not TobInfo.FieldExists ('$$_DEVISE') then
    begin
      TobInfo.AddChampSup ('$$_DEVISE', False);
      TobInfo.PutValue    ('$$_DEVISE', V_PGI.DevisePivot);
    end;

    DevisePiece := TobInfo.GetValue ('$$_DEVISE');
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

    DEV.Code:=Tob_Piece.GetValue ('GP_DEVISE');
    GetInfosDevise(DEV) ;
    DEV.Taux:=GetTaux (DEV.Code, DEV.DateTaux, CleDoc.DatePiece) ;
    PutValueDetail(Tob_Piece,'GP_TAUXDEV',DEV.Taux) ;
    PutValueDetail(Tob_Piece,'GP_DATETAUXDEV',DEV.DateTaux) ;
    AttribCotation(Tob_Piece) ;   // Calcul GP_COTATION

    DD:=Tob_Piece.GetValue ('GP_DATEPIECE') ;
    Per:=GetPeriode(DD) ; Sem:=NumSemaine(DD) ;
    Tob_Piece.PutValue ('GP_PERIODE', Per) ;
    Tob_Piece.PutValue ('GP_SEMAINE', Sem) ;

    /////////////////////////////////////////////////////////////////////////////
    // Impacte le stock ????
    /////////////////////////////////////////////////////////////////////////////
    if (TobInfo.FieldExists ('$$_MAJSTOCK')) and (TobInfo.GetValue('$$_MAJSTOCK') = 'X') then MajStockPiece := True
    else MajStockPiece := False ;
    /////////////////////////////////////////////////////////////////////////////
    // Contrôle des prix unitaire à 0
    /////////////////////////////////////////////////////////////////////////////
    if (TobInfo.FieldExists ('$$_CTRLPRIXNUL')) and (TobInfo.GetValue('$$_CTRLPRIXNUL') = 'X') then CtrlPrixNul := True
    else CtrlPrixNul := False ;
    /////////////////////////////////////////////////////////////////////////////
    // Génération des échéances ????                             //NA 07/03/2002
    /////////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_GEREECHEANCES') then
    begin
       if TobInfo.GetValue('$$_GEREECHEANCES') = 'N' then GenereEcheances := False
       else GenereEcheances := True ;
    end else
    begin
       if GetInfoParPiece(NaturePiece,'GPP_GEREECHEANCE') = 'SAN' then GenereEcheances := False
       else GenereEcheances := True ;
    end;

    /////////////////////////////////////////////////////////////////////////////
    // Regroupement des lignes avec même article ????            //NA 07/03/2002
    /////////////////////////////////////////////////////////////////////////////
    if (TobInfo.FieldExists ('$$_REGROUPELIGNES')) and (TobInfo.GetValue('$$_REGROUPELIGNES') = 'X') then RegroupeLignes := True
    else RegroupeLignes := False ;

    /////////////////////////////////////////////////////////////////////////////
    // Nature du transfert dupliqué                              //NA 07/03/2002
    /////////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_DUPLICPIECE') then NatureTrfRecu := TobInfo.GetValue('$$_DUPLICPIECE')
    else NatureTrfRecu := '' ;

    /////////////////////////////////////////////////////////////////////////////
    //Ajout de champ spécifique
    /////////////////////////////////////////////////////////////////////////////
    for i:=1 to 5 do
    begin
      if not TobInfo.FieldExists ('$$_FAMILLETAXE'+IntToStr(i)) then
      begin
        TobInfo.AddChampSup ('$$_FAMILLETAXE'+IntToStr(i), False);
        TobInfo.PutValue    ('$$_FAMILLETAXE'+IntToStr(i), '');
      end;
      Tob_Piece.AddChampSup ('$$_FAMILLETAXE'+IntToStr(i), False);
      Tob_Piece.PutValue    ('$$_FAMILLETAXE'+IntToStr(i), TobInfo.GetValue('$$_FAMILLETAXE'+IntToStr(i)));
    end;

  end else FreeAndNil( Tob_Piece );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Création d'une ligne commentaire, contenant la référence
Suite ........ : de l'article générique pour la gestion spécifique MODE des
Suite ........ : lignes de documents
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.CreeLigneCommentaire ;
var prix    : double;
    prixnet :double ;
begin
    ////////////////////////////////////////////////////////////////
    // Création d'une TOB Fille Ligne rattachée à l'entête de pièce
    ////////////////////////////////////////////////////////////////
    Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, (NumLigneComment -1));     //NA 07/02/2002
    ////////////////////////////////////////////////////////////////
    // Initialise les champs de la ligne à partir de l'entête
    ////////////////////////////////////////////////////////////////
    PieceVersLigne (Tob_Piece, Tob_Ligne);
    Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
    Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
    ////////////////////////////////////////////////////////////////
    // Ajoute des champs supplémentaires pour le calcul des cumuls
    ////////////////////////////////////////////////////////////////
    AddLesSupLigne  (Tob_Ligne, True) ; // Ajout Champs spécifiques pour calcul de la pièce
    //////////////////////////////////////////////
    // Numéro de ligne
    //////////////////////////////////////////////
    Tob_Ligne.PutValue ('GL_NUMLIGNE'    , NumLigneComment);
    Tob_Ligne.PutValue ('GL_NUMORDRE'    , NumLigneComment);

    OldRefArticle := Trim (OldRefArticle);
    Tob_Ligne.PutValue ('GL_REFARTSAISIE', OldRefArticle);
    Tob_Ligne.PutValue ('GL_CODESDIM'    , OldRefArticle);
    Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);
    Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'COM');		// Ligne commentaire

    Tob_Ligne.PutValue ('GL_QTESTOCK'    , TotQte);   // Total des quantités
    Tob_Ligne.PutValue ('GL_QTEFACT'     , TotQte);
    Tob_Ligne.PutValue ('GL_QTERESTE'    , TotQte);   // MODIF RELIQUAT 24/06/03
    ///////////////////////////////////////////
    // Date de livraison
    ///////////////////////////////////////////
    Tob_Ligne.PutValue ('GL_DATELIVRAISON', StrToDate(DateLivraison));
    ////////////////////////////////////////////
    // Prix unitaire
    ///////////////////////////////////////////
    if (TotQte<>0) then
    begin
      Prix    := Arrondi(TotPrix/TotQte,DEV.Decimale);
      PrixNet := Arrondi(TotPrixNet/TotQte,DEV.Decimale);
    end else
    begin
      Prix   := 0.0;
      PrixNet:= 0.0;
    end;

    if Tob_Ligne.getValue    ('GL_FACTUREHT')='X' then
    begin
      Tob_Ligne.PutValue('GL_PUHTDEV'      , Prix   ) ;
      Tob_Ligne.PutValue('GL_PUHTNETDEV'   , PrixNet) ;
      Tob_Ligne.PutValue('GL_PUHT'         , Prix   ) ;
      Tob_Ligne.PutValue('GL_PUHTNET'      , PrixNet) ;
      Tob_Ligne.PutValue('GL_MONTANTHTDEV' , TotPrix) ;
      Tob_Ligne.PutValue('GL_MONTANTHT'    , TotPrixNet) ;
    end else
    begin
      Tob_Ligne.PutValue('GL_PUTTCDEV'      , Prix   ) ;
      Tob_Ligne.PutValue('GL_PUTTCNETDEV'   , PrixNet) ;
      Tob_Ligne.PutValue('GL_PUTTC'         , Prix   ) ;
      Tob_Ligne.PutValue('GL_PUTTCNET'      , PrixNet) ;
      Tob_Ligne.PutValue('GL_MONTANTTTCDEV' , TotPrix) ;
      Tob_Ligne.PutValue('GL_MONTANTTTC'    , TotPrixNet) ;
    end;

    Tob_Ligne.PutValue       ('GL_PRIXPOURQTE', 1);
    Tob_Ligne.PutValue       ('GL_TYPEDIM'    , 'GEN');
    Tob_Ligne.PutValue       ('GL_QUALIFMVT'  , QualifMvt) ;
    Tob_Ligne.PutValue       ('GL_VALIDECOM'  , 'AFF') ;
    ///////////////////////////////////////////////////////
    // Gestion des reliquats
    //////////////////////////////////////////////////////
    if RefInterneLigneCom <> '' then
    begin
      Tob_Ligne.AddChampSup ('$$_NATUREPIECEG', False);
      Tob_Ligne.AddChampSup ('$$_REFINTERNE'  , False);
      Tob_Ligne.PutValue    ('$$_NATUREPIECEG', NatureLigneCom);
      Tob_Ligne.PutValue    ('$$_REFINTERNE'  , RefInterneLigneCom);
    end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Initialisation d'une la TOB ligne de document
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.InitLignePiece (TobInfo : TOB) : boolean;
var TobFilleArticle  : TOB       ;
    DateConv         : TDateTime ;
    ReferenceArtDim  : string    ;
    ReferenceArtGen  : string    ;
    SQL              : string    ;
    DevisePiece      : string    ;
    CodeTva          : string    ;
    FamilleTva       : string    ;
    OperationFinanc  : string    ;
    NaturePiecePrec  : string    ;
    RefinternePrec   : string    ;
    Q                : TQUERY    ;
    MajPxAch         : boolean   ;
    MajPxRev         : boolean   ;
    OnChangePiecePrec: boolean   ;
    IndiceFils       : integer   ;
    PrixAch, PrixVte : double    ;
    PrixVteNet       : double    ;
    PrixVteRem       : double    ;
    Remise           : double    ;
    PrixRev          : double    ;
    MontantDev       : double    ;
    MontantOld       : double    ;
    i                : integer;
begin
  result            := True ;
  OnChangePiecePrec := False ;
  //////////////////////////////////////////////////////////////////////////////////////
  // On retrouve l'article dimensionné
  //////////////////////////////////////////////////////////////////////////////////////
  if TobInfo.FieldExists ('$$_OPCAISSE') then
  begin
    OperationFinanc  := TobInfo.GetValue ('$$_OPCAISSE') ;
    ReferenceArtDim  := ReadTokenSt (OperationFinanc)    ;
    ModePaiement     := ReadTokenSt (OperationFinanc)    ;
    ReferenceArtGen  := copy (ReferenceArtDim, 1, 18)    ;
  end else
  begin
    ReferenceArtDim := TobInfo.GetValue ('$$_ARTICLE');
    ReferenceArtGen := Trim (TobInfo.GetValue ('$$_CODEARTICLE'));
    ModePaiement    := '';
  end ;

  ////////////////////////////////////////////////////////////////
  // Change-t-on de pièce à solder ?
  // Si oui, inutile de faire des regroupements
  // Si oui, il faut impérativement créé de nouvelles lignes
  ////////////////////////////////////////////////////////////////
  NaturePiecePrec := TobInfo.GetValue ('$$_NATUREPIECEG') ;
  RefInternePrec  := TobInfo.GetValue ('$$_REFINTERNE');

  if (NaturePiecePrec <> OldNaturePiece) or (RefInternePrec <> OldRefInterne) then
  begin
    OnChangePiecePrec := True;
    OldNaturePiece    := NaturePiecePrec ;
    OldRefInterne     := RefInternePrec  ;
  end;

  ////////////////////////////////////////////////////////////////
  // Regroupement des lignes avec même article     //NA 07/03/2002
  //
  // MODIF LM le 08/01/2003
  // Si on solde une pièce, il faut regrouper les lignes en fonctions
  // du document d'origine à solder.
  ////////////////////////////////////////////////////////////////
  if not OnChangePiecePrec then
  begin
    if RegroupeLignes then
    begin
      Tob_Ligne := Tob_Piece.FindFirst(['GL_ARTICLE','$$_NATUREPIECEG','$$_REFINTERNE'], [ReferenceArtDim, NaturePiecePrec, RefInternePrec], False) ;
      if Tob_Ligne <> Nil then
      begin
        ////////////////////////////////////////////////////////////////
        // Ajout de la quantité sur la ligne trouvée
        ////////////////////////////////////////////////////////////////
        QuantiteVendueGB := TobInfo.GetValue('$$_QTESTOCK');
        if (TobInfo.FieldExists ('$$_SIGNE')) and ((TobInfo.GetValue ('$$_SIGNE') = '-') or (TobInfo.GetValue ('$$_SIGNE') = 'NEG')) then
           QuantiteVenduePGI := Tob_Ligne.GetValue('GL_QTEFACT') - QuantiteVendueGB
        else QuantiteVenduePGI := Tob_Ligne.GetValue('GL_QTEFACT') + QuantiteVendueGB;
        Tob_Ligne.PutValue('GL_QTESTOCK' , QuantiteVenduePGI);
        Tob_Ligne.PutValue('GL_QTEFACT'  , QuantiteVenduePGI);
        Tob_Ligne.PutValue('GL_QTERESTE' , QuantiteVenduePGI);  // MODIF RELIQUAT 24/06/03

        ////////////////////////////////////////////////////////////////
        // Recherche du prix de vente
        ////////////////////////////////////////////////////////////////
        if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then
        begin
           PrixVte    := Tob_Ligne.GetValue('GL_PUHTDEV') ;
           PrixVteNet := Tob_Ligne.GetValue('GL_PUHTNETDEV') ;
        end else
        begin
          PrixVte    := Tob_Ligne.GetValue('GL_PUTTCDEV') ;
          PrixVteNet := Tob_Ligne.GetValue('GL_PUTTCNETDEV') ;
        end;

        ////////////////////////////////////////////////////////////////
        // Cumul des quantités et des montants
        ////////////////////////////////////////////////////////////////
        TotQte     := TotQte+QuantiteVenduePGI;
        TotPrix    := TotPrix+(QuantiteVenduePGI*PrixVte);
        TotPrixNet := TotPrixNet+(QuantiteVenduePGI*PrixVteNet);
        ////////////////////////////////////////////////////////////////
        // Traitement de la ligne suivante (fin de InitLignePiece)
        ////////////////////////////////////////////////////////////////
        Exit ;
      end ;
    end;
  end;

  /////////////////////////////////////////////////////////////////
  // Création de la ligne commentaire précédente ???
  /////////////////////////////////////////////////////////////////
  if (ReferenceArtGen <> OldRefArticle) or (OnChangePiecePrec)  then
    begin
      if (LigneCommACreer = True) then CreeLigneCommentaire;

      OldRefArticle := ReferenceArtGen;
      TotQte:=0.0; TotPrix:=0.0;

      //////////////////////////////////////////////////////////////////
      // a modifier ultérieurement quand on aura le temps .........
      /////////////////////////////////////////////////////////////////
      if (Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'FFO') then LigneCommACreer := False
      else begin
        SQL:='Select GA_STATUTART From ARTICLE WHERE GA_ARTICLE="'+ReferenceArtDim+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.Findfield('GA_STATUTART').AsString='UNI' then LigneCommACreer := False
        else begin
          NumLigne        := NumLigne+1;
          NumLigneComment := Numligne;
          LigneCommACreer := True;
        end;
        Ferme (Q);
      end ;
    end;

  ////////////////////////////////////////////////////////////////
  // Regroupement des lignes avec même article     //NA 07/03/2002
  ////////////////////////////////////////////////////////////////
  IndiceFils := -1;
  if not OnChangePiecePrec then
  begin
    if RegroupeLignes then
    begin
      /////////////////////////////////////////////////////////////////
      // Recherche de la ligne commentaire
      /////////////////////////////////////////////////////////////////
      Tob_Ligne := Tob_Piece.FindFirst(['GL_CODESDIM','GL_TYPELIGNE','$$_NATUREPIECEG','$$_REFINTERNE'], [Trim (ReferenceArtGen),'COM', NaturePiecePrec, RefInternePrec], False) ;
      if Tob_Ligne <> Nil then
      begin
        IndiceFils := Tob_Ligne.GetValue ('GL_NUMLIGNE');
        LigneCommACreer := False ;
      end;
    end;
  end;

  ////////////////////////////////////////////////////////////////
  // création d'une nouvelle ligne
  ////////////////////////////////////////////////////////////////
  Tob_Ligne:= TOB.CREATE ('LIGNE', Tob_Piece, IndiceFils);
  ////////////////////////////////////////////////////////////////
  // Initialise les champs de la ligne à partir de l'entête
  ////////////////////////////////////////////////////////////////
  PieceVersLigne (Tob_Piece, Tob_Ligne);
  Tob_Ligne.PutValue('GL_PERIODE', Tob_Piece.GetValue('GP_PERIODE')) ;
  Tob_Ligne.PutValue('GL_SEMAINE', Tob_Piece.GetValue('GP_SEMAINE')) ;
  /////////////////////////////////////////////
  // Nouveau numéro de ligne
  //////////////////////////////////////////////
  inc (NumLigne);
  Tob_Ligne.PutValue ('GL_NUMLIGNE', NumLigne);
  Tob_Ligne.PutValue ('GL_NUMORDRE', NumLigne);
  if IndiceFils <> -1 then NumLigne := NumeroteLignes ;
  ////////////////////////////////////////////////////////////////
  // Ajoute des champs supplémentaires pour le calcul des cumuls
  ////////////////////////////////////////////////////////////////
  AddLesSupLigne  (Tob_Ligne, True) ; // Ajout Champs spécifiques pour calcul de la pièce

  if (TobInfo.FieldExists ('$$_NATUREPIECEG')) and (TobInfo.GetValue ('$$_NATUREPIECEG') <> '') and (TobInfo.FieldExists ('$$_REFINTERNE')) and (TobInfo.GetValue ('$$_REFINTERNE') <> '') then
  begin
    PieceASolder := True ;
    Tob_Ligne.AddChampSup ('$$_NATUREPIECEG', False);
    Tob_Ligne.AddChampSup ('$$_REFINTERNE'  , False);
    Tob_Ligne.PutValue    ('$$_NATUREPIECEG', TobInfo.GetValue ('$$_NATUREPIECEG'));
    Tob_Ligne.PutValue    ('$$_REFINTERNE'  , TobInfo.GetValue ('$$_REFINTERNE')  );
    if LigneCommACreer then
    begin
      NatureLignecom     := TobInfo.GetValue ('$$_NATUREPIECEG');
      RefInterneLigneCom := TobInfo.GetValue ('$$_REFINTERNE');
    end;
  end;

  ///////////////////////////////////////////////////////////////
  // Détermination de l'identifiant article
  //////////////////////////////////////////////////////////////
  TobFilleArticle := Tob_Article.FindFirst(['GA_ARTICLE'], [ReferenceArtDim], False) ;   //NA 07/03/2002
  if TobFilleArticle = Nil then
  begin
    // Recherche de l'article
    SQL:='Select * From ARTICLE WHERE GA_ARTICLE="'+ReferenceArtDim+'"';
    Q:=OpenSQL(SQL,True) ;
    if Q.EOF then
    begin
      Ferme (Q);
      EcrireInfo (TraduireMemoire('   Erreur -> Problème pour charger la fiche article ') + ReferenceArtDim, True, True, True);
      result := False ;
    end;
    ////////////////////////////////////////////////////////////////////////////////////////////////
    // Chargement de l'article dans la TOB des articles de la piece
    ////////////////////////////////////////////////////////////////////////////////////////////////
    if result then
    begin
      TobFilleArticle := CreerTOBArt(Tob_Article);
      TobFilleArticle.SelectDB('',Q);
      Ferme(Q) ;
    end;
  end;

  if TobFilleArticle <> Nil then
  begin
    /////////////////////////////////////////////////////////////////////////////
    // Initialisation des champs de la ligne à partir de l'article dimensionné
    /////////////////////////////////////////////////////////////////////////////
    ArticleVersLigne(Tob_Piece, TobFilleArticle, Nil, Tob_Ligne, Tob_Tiers) ;     //NA 07/03/2002

    // Libéllé spécifique pour l'article
    if (TobInfo.FieldExists ('$$_LIBELLELIGNE')) and (TobInfo.GetValue ('$$_LIBELLELIGNE') <> '') then
    begin
      Designation := TobInfo.GetValue ('$$_LIBELLELIGNE') ;
      Tob_Ligne.PutValue('GL_LIBELLE', Designation);
    end else Designation := TobFilleArticle.GetValue('GA_LIBELLE');

    // Famille niveau 1 spécifique pour l'article
    if (TobInfo.FieldExists ('$$_FAMILLENIV1')) and (TobInfo.GetValue ('$$_FAMILLENIV1') <> '') then
      Tob_Ligne.PutValue('GL_FAMILLENIV1', TobInfo.GetValue ('$$_FAMILLENIV1'));

    // Famille niveau 2 spécifique pour l'article
    if (TobInfo.FieldExists ('$$_FAMILLENIV2')) and (TobInfo.GetValue ('$$_FAMILLENIV2') <> '') then
      Tob_Ligne.PutValue('GL_FAMILLENIV2', TobInfo.GetValue ('$$_FAMILLENIV2'));

    // Famille niveau 3 spécifique pour l'article
    if (TobInfo.FieldExists ('$$_FAMILLENIV3')) and (TobInfo.GetValue ('$$_FAMILLENIV3') <> '') then
      Tob_Ligne.PutValue('GL_FAMILLENIV3', TobInfo.GetValue ('$$_FAMILLENIV3'));

    // Collection spécifique pour l'article
    if (TobInfo.FieldExists ('$$_COLLECTION')) and (TobInfo.GetValue ('$$_COLLECTION') <> '') then
      Tob_Ligne.PutValue('GL_COLLECTION', TobInfo.GetValue ('$$_COLLECTION'));

    /////////////////////////////////////////////////////////////////////////////
    //Ajout de champs spécifiques
    /////////////////////////////////////////////////////////////////////////////
    for i:=1 to 5 do
    begin
      if (TobInfo.FieldExists ('$$_FAMILLETAXE'+IntToStr(i))) and (TobInfo.GetValue ('$$_FAMILLETAXE'+IntToStr(i)) <> '') then
        Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),TobInfo.GetValue ('$$_FAMILLETAXE'+IntToStr(i)))
      else
        Tob_Ligne.PutValue('GL_FAMILLETAXE'+IntToStr(i),TobFilleArticle.GetValue('GA_FAMILLETAXE'+IntToStr(i))) ;
    end;
    for i:=1 to 9 do Tob_Ligne.PutValue('GL_LIBREART'+IntToStr(i),TobFilleArticle.GetValue('GA_LIBREART'+IntToStr(i))) ;
    Tob_Ligne.PutValue ('GL_LIBREARTA'  , TobFilleArticle.GetValue('GA_LIBREARTA'));

    /////////////////////////////////////////////////////////////////////////
    // Gestion de la date de livraison à la ligne (O'NEIL)
    /////////////////////////////////////////////////////////////////////////
    if TobInfo.GetValue ('GP_DATELIVRAISON')=iDate1900 then Tob_Ligne.PutValue ('GL_DATELIVRAISON', Tob_Piece.GetValue ('GP_DATEPIECE'))
    else begin
      Tob_Ligne.PutValue ('GL_DATELIVRAISON', TobInfo.GetValue ('GP_DATELIVRAISON'));
      //
      // On tente d'initialiser la date de livraison de l'entête avec la date de livraison
      // la plus proche
      //
      if (TobInfo.GetValue ('GP_DATELIVRAISON') < Tob_Piece.GetValue ('GP_DATELIVRAISON')) then
        Tob_Piece.PutValue ('GP_DATELIVRAISON', TobInfo.GetValue ('GP_DATELIVRAISON'));
    end;
    DateLivraison := Tob_Ligne.GetValue ('GL_DATELIVRAISON');
    /////////////////////////////////////////////////////////////////////////
    // Gestion de la TVA
    /////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_CODETVA') then
    begin
       FamilleTva := GetParamSoc('SO_GCDEFCATTVA') ;
       if FamilleTva = 'TX5'      then FamilleTva := 'GL_FAMILLETAXE5'
       else if FamilleTva = 'TX4' then FamilleTva := 'GL_FAMILLETAXE4'
       else if FamilleTva = 'TX3' then FamilleTva := 'GL_FAMILLETAXE3'
       else if FamilleTva = 'TX2' then FamilleTva := 'GL_FAMILLETAXE2'
       else FamilleTva := 'GL_FAMILLETAXE1' ;
       CodeTva := TobInfo.GetValue ('$$_CODETVA') ;
       Tob_Ligne.PutValue('GL_REGIMETAXE', ReadTokenSt (CodeTva)) ;
       Tob_Ligne.PutValue(FamilleTva, ReadTokenSt (CodeTva)) ;
    end ;

    ////////////////////////////////////////////////////////////////////////////////////////////////
    // Type de ligne : Dépend de l'article
    ////////////////////////////////////////////////////////////////////////////////////////////////
    if Tob_Ligne.GetValue ('GL_NATUREPIECEG') = 'FFO' then Tob_Ligne.PutValue ('GL_TYPEDIM' , 'NOR')
    else begin
      if TobFilleArticle.GetValue('GA_STATUTART') = 'DIM' then Tob_Ligne.PutValue ('GL_TYPEDIM'     , 'DIM')
      else Tob_Ligne.PutValue ('GL_TYPEDIM'     , 'NOR')
    end;

    ///////////////////////////////////////////////////////////////////////////////////////////////
    // Chargement éventuel des fiches stocks pour MAJ du stock
    ///////////////////////////////////////////////////////////////////////////////////////////////
    if MajStockPiece then
    begin
      //
      // MODIF LM 04/06/2002
      //
      // LoadTOBDispo (TobFilleArticle, True) ;
      if (Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'TEM') or (Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'TRE') or (Tob_Piece.GetValue ('GP_NATUREPIECEG') = 'TRV') then
        LoadTOBDispo (TobFilleArticle, True)
      else
        LoadTOBDispo (TobFilleArticle, True, CreerQuelDepot(Tob_Piece)) ;
    end;
  end else
  begin
     EcrireInfo (TraduireMemoire('   Erreur -> Problème pour récupérer la fiche article ') + ReferenceArtDim, True, True, True);
     result := False ;
  end;

  if result then
  begin
    Tob_Ligne.PutValue ('GL_ARTICLE'     , ReferenceArtDim);
    ReferenceArtGen := Trim (ReferenceArtGen);
    Tob_Ligne.PutValue ('GL_CODEARTICLE' , ReferenceArtGen);
    Tob_Ligne.PutValue ('GL_REFARTSAISIE', ReferenceArtGen);
    Tob_Ligne.PutValue ('GL_LIBELLE'     , Designation);

    Tob_Ligne.PutValue ('GL_TARIF'       , 0);
    Tob_Ligne.PutValue ('GL_TYPELIGNE'   , 'ART');
    Tob_Ligne.PutValue ('GL_TYPEREF'     , 'ART');
    //NA 28/03/2002 Tob_Ligne.PutValue ('GL_TYPEARTICLE' , 'MAR');
    Tob_Ligne.PutValue ('GL_VALIDECOM'   , 'AFF') ;
    Tob_Ligne.PutValue ('GL_RECALCULER'  , 'X') ;

    /////////////////////////////////////////////////////////////////////////////
    // Représentant de la ligne (par défaut reprise de celui de la pièce)      //NA 07/03/2002
    /////////////////////////////////////////////////////////////////////////////
    if TobInfo.FieldExists ('$$_REPRESENTANT') then Tob_Ligne.PutValue ('GL_REPRESENTANT', TobInfo.GetValue ('$$_REPRESENTANT'));
    ////////////////////////////////////////////////////////////////////////////
    // Récupération de la quantité
    ////////////////////////////////////////////////////////////////////////////
    QuantiteVendueGB := TobInfo.GetValue('$$_QTESTOCK');
    if (TobInfo.FieldExists ('$$_SIGNE')) and ((TobInfo.GetValue ('$$_SIGNE') = '-') or (TobInfo.GetValue ('$$_SIGNE') = 'NEG')) then
         QuantiteVenduePGI := -QuantiteVendueGB
    else QuantiteVenduePGI := QuantiteVendueGB;
    Tob_Ligne.PutValue('GL_QTESTOCK' , QuantiteVenduePGI);
    Tob_Ligne.PutValue('GL_QTEFACT'  , QuantiteVenduePGI);
    Tob_Ligne.PutValue('GL_QTERESTE' , QuantiteVenduePGI);  // MODIF RELIQUAT 24/06/03
    /////////////////////////////////////////////////////////
    // Récupération des valorisation en PA
    //////////////////////////////////////////////////////////
    MajPxAch := True ;
    if TobInfo.FieldExists ('$$_PRIXACHUNITAIRE') then PrixAch:=TobInfo.GetValue ('$$_PRIXACHUNITAIRE')
    else if TobInfo.FieldExists ('$$_PRIXACHTOTAL') then
    begin
      PrixAch := TobInfo.GetValue ('$$_PRIXACHTOTAL') ;
      if QuantiteVendueGB <> 0 then PrixAch := PrixAch / QuantiteVendueGB ;
    end else
    begin
      PrixAch := 0;
      MajPxAch := False ;       //NA 07/03/2002
    end;
    //////////////////////////////////////////////////////////
    // Récupération des valorisation en PR
    //////////////////////////////////////////////////////////
    MajPxRev := True ;
    if TobInfo.FieldExists ('$$_PRIXREVUNITAIRE') then PrixRev:=TobInfo.GetValue ('$$_PRIXREVUNITAIRE')
    else if TobInfo.FieldExists ('$$_PRIXREVTOTAL') then
    begin
      PrixRev := TobInfo.GetValue ('$$_PRIXREVTOTAL') ;
      if QuantiteVendueGB <> 0 then PrixRev := PrixRev / QuantiteVendueGB ;
    end else
    begin
      PrixRev  := PrixAch;
      MajPxRev := MajPxAch ;    //NA 07/03/2002
    end;
    //////////////////////////////////////////////////////////
    // Récupération du prix de la pièce
    //////////////////////////////////////////////////////////
    //
    // Prix de l'article avant remise
    //
    if TobInfo.FieldExists ('$$_PRIXVTEUNITAIRE') then PrixVte:=TobInfo.GetValue ('$$_PRIXVTEUNITAIRE')
    else if TobInfo.FieldExists ('$$_PRIXVTETOTAL') then
    begin
      PrixVte := TobInfo.GetValue ('$$_PRIXVTETOTAL') ;
      if QuantiteVendueGB <> 0 then PrixVte := PrixVte / QuantiteVendueGB ;
      if PrixVte < 0 then PrixVte := -PrixVte ;
    end else PrixVte := 0;
    //
    // Prix de l'article après remise
    //
    PrixVteRem := PrixVte ;   // Au départ, c'est le même
    if TobInfo.FieldExists ('$$_REMISE') then
    begin
      Remise:=TobInfo.GetValue ('$$_REMISE');
      if PrixVte <> 0 then PrixVteRem := PrixVte * (1 - Remise/100) ;
    end else
    begin
      if TobInfo.FieldExists ('$$_PRIXVTEUNIREM') then PrixVteRem:=TobInfo.GetValue ('$$_PRIXVTEUNIREM')
      else if TobInfo.FieldExists ('$$_PRIXVTETOTREM') then
      begin
        PrixVteRem := TobInfo.GetValue ('$$_PRIXVTETOTREM') ;
        if QuantiteVendueGB <> 0 then PrixVteRem := PrixVteRem / QuantiteVendueGB ;
        if PrixVteRem < 0 then PrixVteRem := -PrixVteRem ;
      end;
      //
      // Calcul du pourcentage de remise
      //
      if PrixVte <> 0 then Remise := (100 * (PrixVte-PrixVteRem)) / PrixVte
      else Remise := 0;
    end;
    //
    // Initialisation de la remise
    //
    Tob_Ligne.PutValue('GL_REMISELIGNE', Remise);

    /////////////////////////////////////////////////////////
    // Motif de démarque
    //////////////////////////////////////////////////////////
    if remise <> 0 then
    begin
      if (TobInfo.FieldExists ('$$_TYPEREMISE')) and (TobInfo.GetValue ('$$_TYPEREMISE') <> '') then
        Tob_Ligne.PutValue ('GL_TYPEREMISE', TobInfo.GetValue ('$$_TYPEREMISE'));
    end;

    /////////////////////////////////////////////////////////
    // Motif pour E/S de stock
    //////////////////////////////////////////////////////////
    if (TobInfo.FieldExists ('$$_MOTIFMVT')) and (TobInfo.GetValue ('$$_MOTIFMVT') <> '') then
      Tob_Ligne.PutValue ('GL_MOTIFMVT', TobInfo.GetValue ('$$_MOTIFMVT'));

    ///////////////////////////////////////////////////////////////////////////
    // Conversion en devise dossier si besoin
    ///////////////////////////////////////////////////////////////////////////
    DevisePiece := Tob_Ligne.GetValue ('GL_DEVISE');
    if VH^.TenueEuro then
    begin
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then DevisePiece := V_PGI.DeviseFongible ;
    end else
    begin
      if Tob_Piece.GetValue ('GP_SAISIECONTRE') = 'X' then DevisePiece := 'EUR' ;
    end;

    if  DevisePiece = V_PGI.DevisePivot then
    begin
      if MajPxAch then       //NA 07/03/2002
      begin
         Tob_Ligne.PutValue('GL_DPA'         , PrixAch);
         Tob_Ligne.PutValue('GL_PMAP'        , PrixAch);
         Tob_Ligne.PutValue('GL_PMAPACTU'    , PrixAch);
      end ;

      if MajPxRev then       //NA 07/03/2002
      begin
         Tob_Ligne.PutValue('GL_DPR'         , PrixRev);
         Tob_Ligne.PutValue('GL_PMRP'        , PrixRev);
         Tob_Ligne.PutValue('GL_PMRPACTU'    , PrixRev);
      end ;
    end else
    begin
      DateConv   := NowH ;
      if MajPxAch then
      begin
         MontantDev := 0 ;  // initialisation bidon
         MontantOld := 0 ;  // initialisation bidon
         ToxConvertir (PrixAch, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, nil) ;

         Tob_Ligne.PutValue ('GL_DPA'         , MontantDev);
         Tob_Ligne.PutValue ('GL_PMAP'        , MontantDev);
         Tob_Ligne.PutValue ('GL_PMAPACTU'    , MontantDev);
      end ;

      if MajPxRev then
      begin
         MontantDev := 0 ;  // initialisation bidon
         MontantOld := 0 ;  // initialisation bidon
         ToxConvertir (PrixRev, MontantDev, MontantOld, DevisePiece, V_PGI.DevisePivot, DateConv, nil) ;

         Tob_Ligne.PutValue ('GL_DPR'         , MontantDev);
         Tob_Ligne.PutValue ('GL_PMRP'        , MontantDev);
         Tob_Ligne.PutValue ('GL_PMRPACTU'    , MontantDev);
      end ;
    end;

    // PV ........
    if Tob_Ligne.GetValue('GL_FACTUREHT')='X' then
    begin
      Tob_Ligne.PutValue('GL_PUHTDEV'   , PrixVte);
      Tob_Ligne.PutValue('GL_PUHTNETDEV', PrixVteRem);
    end else
    begin
      Tob_Ligne.PutValue('GL_PUTTCDEV'   , PrixVte) ;
      Tob_Ligne.PutValue('GL_PUTTCNETDEV', PrixVteRem) ;
    end;

    Tob_Ligne.PutValue ('GL_PRIXPOURQTE' , 1);

    Tob_Ligne.PutValue ('GL_QUALIFMVT'   , QualifMvt) ;

    TotQte     := TotQte+QuantiteVenduePGI;
    TotPrix    := TotPrix+(QuantiteVenduePGI*PrixVte);
    TotPrixNet := TotPrixNet+(QuantiteVenduePGI*PrixVteRem);
  end else FreeAndNil( Tob_Ligne );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Renumérotation des lignes d'une pièce
Mots clefs ... : IMPORTPIECE
*****************************************************************}
function TTImportGPAO.NumeroteLignes : integer ;
var ind, inl, ino  : integer ;
    TOBL           : TOB ;
begin
  inl := 0 ; ino := 0 ;
  for ind := 0 to Tob_Piece.Detail.Count -1 do
  begin
    TOBL := Tob_Piece.Detail[Ind] ;
    if ind = 0 then
    begin
      inl := TOBL.GetNumChamp('GL_NUMLIGNE');
      ino := TOBL.GetNumChamp('GL_NUMORDRE');
    end;
    TOBL.PutValeur(inl, (ind +1));
    TOBL.PutValeur(ino, (ind +1));
  end;
  Result := Tob_Piece.Detail.Count ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Duplication du transfert émis pour obtenir le transfert reçu
Mots clefs ... : IMPORTPIECE
*****************************************************************}
procedure TTImportGPAO.DupliqueTrfRecu ;
var Tob_Piece_Recu : TOB    ;
    Tob_Ligne_Recu : TOB    ;
    Tob_PiedB_Recu : TOB    ;
    Tob_PiedF_Recu : TOB    ;
    Tob_PiedE_Recu : TOB    ;
    SoucheTrfPGI   : string ;
    Etab	       : string ;
    NbLigne        : integer;
begin
  //////////////////////////////////////////////////////////////////
  // Création du transfert recu
  //////////////////////////////////////////////////////////////////
  Etab := Tob_Piece.GetValue ('GP_DEPOTDEST');

  ///////////////////////////////////////////////////////////////////
  // Détermination de la clé d'accès
  //////////////////////////////////////////////////////////////////
  SoucheTrfPGI := GetSoucheG (NatureTrfRecu, Etab,'') ;
  //NA 28/03/2002 NumeroTrfPGI := GetNumSoucheG (SoucheTrfPGI);
  //NA 28/03/2002 IncNumSoucheG (SoucheTrfPGI) ;  	// Incrémentation du numéro de pièce

  ////////////////////////////////////////////////////////////////////////////////
  //  Duplication des TOB du transfert émis pour créer les TOB du transferts reçu
  ////////////////////////////////////////////////////////////////////////////////
  Tob_Piece_Recu := TOB.Create ('PIECE', nil, -1);
  Tob_Piece_Recu.Dupliquer (Tob_Piece, True, True, True);
  Tob_Piece_Recu.PutValue  ('GP_NATUREPIECEG', NatureTrfRecu);
  Tob_Piece_Recu.PutValue  ('GP_SOUCHE'      , SoucheTrfPGI );
  Tob_Piece_Recu.PutValue  ('GP_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
  Tob_Piece_Recu.PutValue  ('GP_INDICEG'     , 0 );
  Tob_Piece_Recu.PutValue  ('GP_VIVANTE'     , Tob_Piece.GetValue ('GP_VIVANTE'));       // Pièce vivante    : Non

  ////////////////////////////////////////////////////////////////////
  // MAJ des lignes de transfert
  ////////////////////////////////////////////////////////////////////
  for NbLigne:=0 to Tob_Piece_Recu.Detail.Count-1 do
  begin
    Tob_Ligne_Recu:=Tob_Piece_Recu.Detail[NbLigne];
    Tob_Ligne_Recu.PutValue ('GL_NATUREPIECEG', NatureTrfRecu);
    Tob_Ligne_Recu.PutValue ('GL_SOUCHE'      , SoucheTrfPGI  );
    Tob_Ligne_Recu.PutValue ('GL_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    Tob_Ligne_Recu.PutValue ('GL_INDICEG'     , 0 );
    Tob_Ligne_Recu.PutValue ('GL_DEPOT'       , Tob_Piece.GetValue ('GP_DEPOTDEST'));
  end;

  ////////////////////////////////////////////////////////////////////
  // MAJ du stock en cas de besoin
  ////////////////////////////////////////////////////////////////////
  if MajStockPiece then
  begin
    ValideLesLignes   (Tob_Piece_Recu, Tob_Article, Tob_Catalogue, Tob_Noment, nil, nil, nil, False, False) ;
    ValideLesArticles (Tob_Piece_Recu, Tob_Article) ;
  end;

  ////////////////////////////////////////////////////////////////////
  // MAJ PIEDBASE
  ////////////////////////////////////////////////////////////////////
  Tob_PiedB_Recu := TOB.Create ('LE PIED BASE', nil, -1);
  Tob_PiedB_Recu.Dupliquer (Tob_PiedBase, True, True, True);
  // MAJ des lignes de piedbase
  for NbLigne:=0 to Tob_PiedB_Recu.Detail.Count-1 do
  begin
    Tob_PiedF_Recu := Tob_PiedB_Recu.Detail[NbLigne];
    Tob_PiedF_Recu.PutValue ('GPB_NATUREPIECEG', NatureTrfRecu);
    Tob_PiedF_Recu.PutValue ('GPB_SOUCHE'      , SoucheTrfPGI  );
    Tob_PiedF_Recu.PutValue ('GPB_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    Tob_PiedF_Recu.PutValue ('GPB_INDICEG'     , 0  );
  end;

  ////////////////////////////////////////////////////////////////////
  // MAJ PIEDECHE
  ////////////////////////////////////////////////////////////////////
  Tob_PiedE_Recu := TOB.Create ('LE PIED ECHE', nil, -1);
  Tob_PiedE_Recu.Dupliquer (Tob_PiedEche, True, True, True);
  // MAJ des lignes de piedeche
  for NbLigne:=0 to Tob_PiedE_Recu.Detail.Count-1 do
  begin
    Tob_PiedF_Recu := Tob_PiedE_Recu.Detail[NbLigne];
    Tob_PiedF_Recu.PutValue ('GPE_NATUREPIECEG', NatureTrfRecu);
    Tob_PiedF_Recu.PutValue ('GPE_SOUCHE'      , SoucheTrfPGI  );
    Tob_PiedF_Recu.PutValue ('GPE_NUMERO'      , Tob_Piece.GetValue ('GP_NUMERO'));
    Tob_PiedF_Recu.PutValue ('GPE_INDICEG'     , 0  );
  end;

  //////////////////////////////////////////////////////////////////////////
  // On fait le lien entre le transfert émis et le transfert reçu
  /////////////////////////////////////////////////////////////////////////
  Tob_Piece.PutValue ('GP_DEVENIRPIECE', EncodeRefPiece(Tob_Piece_Recu)) ;

  //////////////////////////////////////////////////////////////////////////
  // MAJ de la base : PIECE, LIGNE, PIEDBASE, PIEDECHE, ADRESSES
  /////////////////////////////////////////////////////////////////////////
//  Tob_Piece_Recu.SetAllModifie(TRUE) ;
//  Tob_PiedB_Recu.SetAllModifie(TRUE) ;
//  Tob_PiedE_Recu.SetAllModifie(TRUE) ;

  Tob_Piece_Recu.InsertDB (nil);
  Tob_PiedB_Recu.InsertorUpdateDB (False);
  Tob_PiedE_Recu.InsertorUpdateDB (False);

  //////////////////////////////////////////////////////////////
  // Libération des listes mémoires
  //////////////////////////////////////////////////////////////
  FreeAndNil( Tob_Piece_Recu );  // Libère la TOB Piece+Ligne
  FreeAndNil( Tob_PiedB_Recu );  // Libère la TOB PiedBase
  FreeAndNil( Tob_PiedE_Recu );  // Libère la TOB PiedEche
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Contrôle de la TOB PIECE avant import
Mots clefs ... : IMPORTPIECE
*****************************************************************}
function TTImportGPAO.CtrlTOBPiece (var TobInfo : TOB) : boolean;
var ReferencePiece       : string ;
    ArticleDefaut        : string ;
    SQL                  : string ;
    ArticleDimension     : string ;
    ReferenceArtGen      : string ;
    StRech               : string ;
    NaturePiece          : string ;
    RefInterne           : string ;
    MessagePiece         : string ;
    Nature               : string ;
    Q                    : TQuery ;
    Resultat             : boolean;
    ArticleOk            : boolean;
    RemisePied           : double ;
    Escompte             : double ;
begin
  Resultat   := True;
  RemisePied := 0.0 ;
  Escompte   := 0.0 ;
  //
  // Contrôle 1 : il faut la référence de la pièce
  //
  if not TobInfo.FieldExists ('GP_REFINTERNE') then
  begin
    EcrireInfo (TraduireMemoire('  TABLE PIECE : la référence de la pièce n''est pas récupérée'), True, True, True);
    Resultat := False;
  end else
  begin
    if TobInfo.GetValue ('GP_REFINTERNE') = '' then
    begin
      EcrireInfo (TraduireMemoire('  TABLE PIECE : la référence de la pièce n''est pas renseignée'), True, True, True);
      Resultat := False;
    end;
  end;
  if Resultat then ReferencePiece := TobInfo.GetValue ('GP_REFINTERNE') else ReferencePiece := '';
  //
  // Contrôle : il faut une nature de document
  //
  if Resultat then
  begin
    MessagePiece := TraduireMemoire ('  TABLE PIECE référence ') + ReferencePiece ;

    if not TobInfo.FieldExists ('GP_NATUREPIECEG') then
    begin
      EcrireInfo (MessagePiece +  TraduireMemoire(' : la nature de document est obligatoire'), True, True, True);
      Resultat := False;
    end else
    begin
      if TobInfo.GetValue ('GP_NATUREPIECEG') = '' then
      begin
        EcrireInfo (MessagePiece + TraduireMemoire(' : la nature de document n''est pas renseignée'), True, True, True);
        Resultat := False;
      end else
      begin
        StRech := TobInfo.GetValue ('GP_NATUREPIECEG');
        if StRech <> DerniereNature then
        begin
          SQL:='Select GPP_LIBELLE From PARPIECE WHERE GPP_NATUREPIECEG="'+StRech+'"';
          Q:=OpenSQL(SQL,True) ;
          if Q.EOF then
          begin
            EcrireInfo (MessagePiece + TraduireMemoire(' : la nature de document ') + StRech + TraduireMemoire(' n''existe pas '), True, True, True);
            Resultat:=False ;
          end else
          begin
           LibellePiece := Q.Findfield('GPP_LIBELLE').AsString;
           DerniereNature := StRech ;
          end;
          Ferme(Q);
        end;
      end;
    end;
  end;
  //
  // Message affiché
  //
  MessagePiece := '  '  + LibellePiece + traduireMemoire (' référence ') ;
  //
  // Contrôle de l'unicité de la pièce si demandé
  //
  if Resultat then
  begin
    if (ReferencePiece <> '') and (ReferencePiece <> DernierRefInterne) then
    begin
      Nature := TobInfo.GetValue ('GP_NATUREPIECEG') ;
      SQL:='Select GP_NUMERO From PIECE WHERE GP_NATUREPIECEG="'+Nature+'" and GP_REFINTERNE="'+ReferencePiece+'"';
      Q:=OpenSQL(SQL,True) ;
      if not Q.EOF then
      begin
        if (TobInfo.FieldExists ('$$_PIECEUNIQUE')) and (TobInfo.GetValue ('$$_PIECEUNIQUE')='X') then
        begin
          EcrireInfo (MessagePiece + ReferencePiece , True, True, True);
          EcrireInfo (TraduireMemoire ('  Cette pièce existe déjà et ne peut être ré-intégrée') , True, True, True);
          Resultat := False;
        end else
        begin
          EcrireInfo (MessagePiece + ReferencePiece , True, True, True);
          EcrireInfo (TraduireMemoire ('  Attention, une pièce avec la même référence existe déjà') , True, True, True);
        end;
      end;
      Ferme(Q);
      DernierRefInterne := ReferencePiece ;
    end;
  end;
  //
  // Contrôle 2 : il faut un établissement
  //
  if not TobInfo.FieldExists ('GP_ETABLISSEMENT') then
  begin
    EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire (' : l''établissement de la pièce n''est pas récupéré'), True, True, True);
    Resultat := False;
  end else
  begin
    if TobInfo.GetValue ('GP_ETABLISSEMENT') = '' then
    begin
      EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : l''établissement de la pièce n''est pas renseigné'), True, True, True);
      Resultat := False;
    end else
    begin
      StRech := TobInfo.GetValue ('GP_ETABLISSEMENT');
      if (StRech <> '') and (StRech <>DernierEtab) then
      begin
        SQL:='Select ET_LIBELLE From ETABLISS WHERE ET_ETABLISSEMENT="'+StRech+'"';
        Q:=OpenSQL(SQL,True) ;
        if Q.EOF then
        begin
          EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : l''établissement ') + StRech + TraduireMemoire(' n''existe pas'), True, True, True);
          resultat := False ;
        end else DernierEtab := StRech ;
        Ferme(Q);
      end;
    end;
  end;

  //
  // Contrôle 8 : S'il faut solder une pièce, on la recherche
  // Attention : cette pièce doit être vivante.
  //
  if not TobInfo.FieldExists ('$$_NATUREPIECEG') then
  begin
    TobInfo.AddChampSup ('$$_NATUREPIECEG', False);
    TobInfo.PutValue    ('$$_NATUREPIECEG', '');
  end;
  if not TobInfo.FieldExists ('$$_REFINTERNE') then
  begin
    TobInfo.AddChampSup ('$$_REFINTERNE'  , False);
    TobInfo.PutValue    ('$$_REFINTERNE'  , '');
  end;

  if (TobInfo.GetValue ('$$_NATUREPIECEG') <> '') and (TobInfo.GetValue ('$$_REFINTERNE') <> '') then
  begin
    if Resultat then
    begin
      NaturePiece := TobInfo.GetValue ('$$_NATUREPIECEG') ;
      RefInterne  := TobInfo.GetValue ('$$_REFINTERNE') ;

      if (NaturePiece <> DerniereNaturePrec) or (RefInterne <> DerniereReferencePrec) then
      begin
        Q := OpenSQL('Select GP_TIERS, GP_REPRESENTANT, GP_REMISEPIED, GP_ESCOMPTE from PIECE Where GP_NATUREPIECEG="' + NaturePiece + '" AND GP_REFINTERNE="' + RefInterne + '" AND GP_VIVANTE="X"',False) ;
	    if Q.EOF then
        begin
          EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : impossible de retrouver la pièce vivante à solder référence ' + RefInterne), True, True, True);
          Resultat := False;
        end else
        begin
          DerniereNaturePrec    := NaturePiece ;
          DerniereReferencePrec := RefInterne  ;

          TiersPiecePrecedente := Q.FindField('GP_TIERS').AsString ;
          ReprePiecePrecedente := Q.FindField('GP_REPRESENTANT').AsString ;
          RemisePied           := Q.FindField('GP_REMISEPIED').AsFloat ;
          Escompte             := Q.FindField('GP_ESCOMPTE').AsFloat ;
        end;
        Ferme (Q) ;
      end;

      if Resultat then
      begin
        //
        // Récupère les remises pieds et escompte
        //
        TobInfo.PutValue ('GP_REMISEPIED', RemisePied);
        TobInfo.PutValue ('GP_ESCOMPTE'  , Escompte  );

        //
        // Récupération du tiers si pas fournit
        //
        if TobInfo.GetValue ('GP_TIERS') <> '' then
        begin
          if  TiersPiecePrecedente <> TobInfo.GetValue ('GP_TIERS') then
          begin
            EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : Le tiers de la pièce précédente est différent !'), True, True, True);
            Resultat := False;
          end;
        end else
        begin
          TobInfo.PutValue ('GP_TIERS'        , TiersPiecePrecedente);
          TobInfo.PutValue ('GP_TIERSPAYEUR'  , TiersPiecePrecedente);
          TobInfo.PutValue ('GP_TIERSLIVRE'   , TiersPiecePrecedente);
          TobInfo.PutValue ('GP_TIERSFACTURE' , TiersPiecePrecedente);
        end;
        //
        // Récupération du représentant si pas fournit
        //
        if not TobInfo.FieldExists ('$$_REPRESENTANT') then
        begin
          TobInfo.AddChampSup ('$$_REPRESENTANT', False);
          TobInfo.PutValue    ('$$_REPRESENTANT', '');
        end;
        if  TobInfo.GetValue ('$$_REPRESENTANT') = '' then
        begin
          TobInfo.PutValue ('$$_REPRESENTANT' , ReprePiecePrecedente);
        end;
      end;
    end;
  end;
  //
  // Contrôle 4 : il faut un tiers
  //
  if not TobInfo.FieldExists ('GP_TIERS') then
  begin
    EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : le tiers de la pièce n''est pas récupéré'), True, True, True);
    Resultat := False;
  end else
  begin
    if TobInfo.GetValue ('GP_TIERS') = '' then
    begin
      EcrireInfo (TraduireMemoire('  TABLE PIECE : le tiers de la pièce n''est pas renseigné'), True, True, True);
      Resultat := False;
    end else
    begin
      if ((TobInfo.GetValue ('GP_TIERS') = '999999') or (TobInfo.GetValue ('GP_TIERS') = TobInfo.GetValue ('GP_ETABLISSEMENT') + '999999')) and
          (TobInfo.FieldExists ('$$_CREECLICASH')) and (TobInfo.GetValue ('$$_CREECLICASH') = 'X') then
      begin
        if DernierClientCree = ''  then
        begin
          EcrireInfo (TraduireMemoire('  TABLE PIECE : le tiers de la pièce 999999 n''est pas connu'), True, True, True);
          Resultat := False;
        end;
      end;
    end;
  end;
  //
  // Contrôle 5 : il faut une date de pièce
  //
  if not TobInfo.FieldExists ('GP_DATEPIECE') then
  begin
    EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : la date de la pièce n''est pas récupérée'), True, True, True);
    Resultat := False;
  end ;

  //
  // Contrôle 6 : il faut un article ou un code à barre
  //
  if Resultat then
  begin
    ArticleOK := True ;
    //
    // A-t-on récupéré une info pour l'article ?
    //
    if (not TobInfo.FieldExists ('$$_CODEARTICLE')) and (not TobInfo.FieldExists ('$$_CODEBARRE')) and (not TobInfo.FieldExists ('$$_OPCAISSE')) then
    begin
      EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : L''article ou le code à barres ne sont pas récupérés'), True, True, True);
      ArticleOK := False;
    end ;
    //
    // Travail en CB en priorité
    //      Si travail sur CB, il faut qu'il soit renseigné
    //      Si travail sur code article, il faut qu'il soit renseigné
    //
    if (TobInfo.FieldExists ('$$_CODEBARRE')) and (TobInfo.GetValue ('$$_CODEBARRE') = '') then
    begin
      EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : le code à barre n''est pas renseigné'), True, True, True);
      ArticleOK := False;
    end else
    begin
      if (TobInfo.FieldExists ('$$_CODEARTICLE')) and (TobInfo.GetValue ('$$_CODEARTICLE') = '') then
      begin
        EcrireInfo (TraduireMemoire('   Erreur -> La référence article n''est pas renseignée ! '), True, True, True);
        ArticleOK := False;
      end else
      begin
        //
        // Dernière option : L'opération financière est elle renseignée ??
        //
        if (TobInfo.FieldExists ('$$_OPCAISSE')) and (TobInfo.GetValue ('$$_OPCAISSE') = '') then
        begin
          EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : l''opération financière n''est pas renseignée'), True, True, True);
          ArticleOK := False;
        end;
      end;
    end;
    //
    // Contrôle de l'existence ce l'article
    //
    if not TobInfo.FieldExists ('$$_OPCAISSE') then
    begin
      //
      // Contrôle et recherche de l'article
      //
      ArticleDimension   := ReconstitueIdentArticle (TobInfo);
      if ArticleDimension = '' then ArticleOK := False
      else
      begin
        if not TobInfo.FieldExists ('$$_ARTICLE')     then TobInfo.AddChampSup ('$$_ARTICLE', False);
        if not TobInfo.FieldExists ('$$_CODEARTICLE') then TobInfo.AddChampSup ('$$_CODEARTICLE', False);
        TobInfo.PutValue ('$$_ARTICLE', ArticleDimension);
        ReferenceArtGen := Trim (copy (ArticleDimension, 1, 18));
        TobInfo.PutValue ('$$_CODEARTICLE', ReferenceArtGen);
      end;
    end;
    //
    // Cas où l'article n'est pas OK
    // On a encore une chance si on veut le remplacer par une article "défaut"
    //
    if not ArticleOk then
    begin
      if (TobInfo.FieldExists ('$$_FORCEARTICLE')) and (TobInfo.GetValue ('$$_FORCEARTICLE') = 'X') then
      begin
        ArticleDefaut := GetParamSoc('SO_GCTOXARTICLE') ;
        SQL:='Select GA_CODEARTICLE From ARTICLE WHERE GA_ARTICLE="'+ArticleDefaut+'"';
        Q:=OpenSQL(SQL,True) ;
        if not Q.EOF then
        begin
          EcrireInfo (TraduireMemoire('               -> L''article est initialisé avec ') + ArticleDefaut, True, True, True);
          //
          // MAJ du code article
          //

          if not TobInfo.FieldExists ('$$_ARTICLE')      then TobInfo.AddChampSup ('$$_ARTICLE'     , False);
          if not TobInfo.FieldExists ('$$_CODEARTICLE')  then TobInfo.AddChampSup ('$$_CODEARTICLE' , False);
          if not TobInfo.FieldExists ('$$_LIBELLELIGNE') then TobInfo.AddChampSup ('$$_LIBELLELIGNE', False);

          TobInfo.PutValue ('$$_ARTICLE'      , ArticleDefaut);
          TobInfo.PutValue ('$$_CODEARTICLE'  , Q.Findfield('GA_CODEARTICLE').AsString);
          TobInfo.PutValue ('$$_LIBELLELIGNE' , IdentArticleImport);

          ArticleOk := True ;
        end else
        begin
          EcrireInfo (TraduireMemoire('   Erreur -> Attention, l''article défaut ') + ArticleDefaut + TraduireMemoire(' est incorrect !! '), True, True, True);
        end;
        Ferme (Q);
      end;
    end;
    Resultat := ArticleOk ;
  end;
  //
  // Contrôle 7 : il faut une quantité
  //
  if not TobInfo.FieldExists ('$$_QTESTOCK') then
  begin
    EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : la quantité n''est pas récupérée'), True, True, True);
    Resultat := False;
  end else
  begin
    if TobInfo.GetValue ('$$_QTESTOCK') = 0 then
    begin
      EcrireInfo (MessagePiece + ReferencePiece + TraduireMemoire(' : la quantité n''est pas renseignée'), True, True, True);
    end;
  end;

  result := Resultat ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Appelle les fonctions de contrôles avant intégration des
Suite ........ : données
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.ControleInfoOK (NomTable : string; var TobInfo : TOB) : boolean;
var LePointDuTraitement : integer ;
begin
  //
  // MAJ des compteurs + Affichage
  //
  inc (CptTrt) ;
  if CptTrt < 100 then AfficheCompteur
  else if CptTrt < 2000 then
  begin
    LePointDuTraitement :=  CptTrt mod 100;
    if LePointDuTraitement = 0 then AfficheCompteur ;
  end else
  begin
    LePointDuTraitement :=  CptTrt mod 500;
    if LePointDuTraitement = 0 then AfficheCompteur ;
  end;
  //
  // Traitement
  //
  if      (NomTable = 'CHOIXCOD')    then result := CtrlTOBChoixCod  (TobInfo)
  else if (NomTable = 'CHOIXEXT')    then result := CtrlTOBChoixExt  (TobInfo)
  else if (NomTable = 'DIMENSION')   then result := CtrlTOBDimension (TobInfo)
  else if (NomTable = 'ARTICLE')     then result := CtrlTOBArticle   (TobInfo)
  else if (NomTable = 'DISPO')       then result := CtrlTOBDispo     (TobInfo)
  else if (NomTable = 'TARIF')       then result := CtrlTOBTarif     (TobInfo)
  else if (NomTable = 'PIECE')       then result := CtrlTOBPiece     (TobInfo)
  else if (NomTable = 'TIERS')       then result := CtrlTOBTiers     (TobInfo)
  else if (NomTable = 'PIEDECHE')    then result := CtrlTOBPiedEche  (TobInfo)
  else if (NomTable = 'ACOMPTES')    then result := CtrlTOBAcomptes  (TobInfo)
  else if (NomTable = 'TRANSINVLIG') then result := CtrlTOBInventaire(TobInfo)
  else if (NomTable = 'FIDELITELIG') then result := CtrlTOBFidelite  (TobInfo)
  else if (NomTable = 'PROSPECTS')   then result := CtrlTOBProspects (TobInfo)
  else result := True;
end;


/////////////////////////////////////////////////////////////////////////////
// Procedure d'intégration des éléments d'une GPAO
/////////////////////////////////////////////////////////////////////////////
procedure TTImportGPAO.IntegreGPAO (var Info : string) ;
begin
  NomTable := Tob_RepriseGPAO.GetValue ('GRE_NOMTABLE');

  if (NomTable <> 'PIECE') then
  begin
    IntegrationPiece := False ;
    if not IntegreInfoGPAO (Info, True) then
    begin
      TopageEnregistrement ('R');
      Inc (CptErr);
    end else
    begin
      ImportTobInfoGPAO;
      inc (CptMAJ) ;
      Inc (CptIntegre);
      TopageEnregistrement ('I');
    end;
  end else
  begin
    IntegrationPiece := True ;
    IntegrePiece (Info);
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// Procedure d'intégration des éléments d'une GPAO
/////////////////////////////////////////////////////////////////////////////
procedure TTImportGPAO.IntegreGPAOFixe (var Info : string) ;
begin
  NomTable := Tob_RepriseGPAO.GetValue ('GRE_NOMTABLE');

  IntegrationPiece := False ;
  if not IntegreInfoGPAO (Info, False) then
  begin
    Inc (CptErr);
  end else
  begin
    ImportTobInfoGPAO;
    inc (CptMAJ) ;
    Inc (CptIntegre);
    Inc (CptEnregTot);
  end;
end;

/////////////////////////////////////////////////////////////////////////////
// function de traitement de la formule d'un champ
/////////////////////////////////////////////////////////////////////////////
function TTImportGPAO.TraiteFormule (Valeur, ValDefaut : string; TOBFilleRecup, TobFilleParam : TOB) : String;
var Formule, Expr : string;
    XX            : TGPAOFormule;
    VM            : TM3VM ;
begin
  Result := Valeur;
  if TobFilleParam.FieldExists ('GRC_FORMULE') then Formule := TobFilleParam.GetValue ('GRC_FORMULE')
                                               else Formule := '';
  if Formule <> '' then
  begin
    // Recherche des champs utilisés dans la formule
    XX := TGPAOFormule.Create (TOBFilleRecup, Valeur, ValDefaut) ;
    Expr := ChampFormule (Formule, XX.PrendChamp, Nil, 1) ;
    FreeAndNil( XX );

    if Pos ('(', Expr) > 0 then
    begin
      // Evaluation de l'expression
      VM := TM3VM.Create ;
      Result := VarAsType (VM.GetExprValue (Expr), varString) ;
      FreeAndNil( VM );
    end else
    begin
      Result := Expr;
    end;
    if Result = #0 then Result := '';
    Result := Trim (Result);
    end;
end;

/////////////////////////////////////////////////////////////////////////////
// Procédure d'intégration des éléments d'une GPAO
/////////////////////////////////////////////////////////////////////////////
procedure TTImportGPAO.FileToTob (Info : string; TOBFilleRecup : TOB; Entete : boolean) ;
var TobFilleParam : TOB    ;
    NomChamp      : string ;
    TypeChamp     : string ;
    ValDefaut     : string ;
    Valeur        : string ;
    TypeDuChamp   : TFixedTypeChamp ;
    Compteur      : integer;
    RecupInfo     : boolean;
    PosDebut      : integer;
    Longueur      : integer;
    LongueurMax   : integer;
begin
  TypeDuChamp   :=  tcTrim ;   // Type par défaut
  LongueurMax   := length (Info);

  for Compteur:=0 to Tob_RepriseGPAO.detail.count-1 do
  begin
    TobFilleParam := Tob_RepriseGPAO.Detail [Compteur];
    NomChamp      := TobFilleParam.GetValue ('GRC_NOMCHAMP');
    TypeChamp     := TobFilleParam.GetValue ('GRC_TYPECHAMP');
    PosDebut      := TobFilleParam.GetValue ('GRC_POSDEBUT');
    Longueur      := TobFilleParam.GetValue ('GRC_LONGUEUR');
    ValDefaut     := TobFilleParam.GetValue ('GRC_VALDEFAUT');
    Valdefaut     := Trim (Valdefaut);
    If (NomChamp = 'GP_TIERS') and (Valdefaut <> '') then
      ValTiersDefaut := Valdefaut;

    if Entete then  PosDebut := PosDebut + 7  else PosDebut := PosDebut + 1;
    //
    // Création des éventuels champs supplémentaires
    //
    if not TobFilleRecup.FieldExists (NomChamp) then
    begin
      TobFilleRecup.AddChampSup (NomChamp, False);
      if copy (TypeChamp, 1, 7) = 'VARCHAR' then TobFilleRecup.PutValue (NomChamp, '')
      else if TypeChamp = 'COMBO'           then TobFilleRecup.PutValue (NomChamp, '')
      else if TypeChamp = 'BOOLEAN'         then TobFilleRecup.PutValue (NomChamp, 'N')
      else if TypeChamp = 'DATE'            then TobFilleRecup.PutValue (NomChamp, StrToDate('01/01/1900'))
      else if TypeChamp = 'INTEGER'         then TobFilleRecup.PutValue (NomChamp, 0)
      else if TypeChamp = 'DOUBLE'          then TobFilleRecup.PutValue (NomChamp, 0.0)
      else if TypeChamp = 'RATE'            then TobFilleRecup.PutValue (NomChamp, 0.0);
    end;
    //
    // Récupère t on l'info dans le fichier ?
    //
    RecupInfo := False ;
    // if (PosDebut >= 0) and (Longueur > 0) and (Posdebut < LongueurMax) and (PosDebut+Longueur <= LongueurMax+1) then
    if (PosDebut >= 0) and (Longueur > 0) and (Posdebut < LongueurMax) then
    begin
      Valeur := copy (Info, Posdebut, Longueur);

      // Traitement de la formule
      Valeur := TraiteFormule (Valeur, ValDefaut, TOBFilleRecup, TobFilleParam);

      // Traduction en fonction de la table de correspondance
      if (copy (TypeChamp, 1, 7) = 'VARCHAR') or (TypeChamp = 'COMBO') then
      begin
        if TobFilleParam.GetValue ('GRC_GERETABLECORR') = 'X' then
        begin
          Valeur := TraduitGPAOenPGI (TobFilleParam.GetValue ('GRC_TYPEREPRISE'), NomChamp, Valeur, Tob_CorrespGPAO);
        end;
      end;

      if trim (Valeur) <> '' then RecupInfo := True;
    end;

    if copy (TypeChamp, 1, 7) = 'VARCHAR' then TypeDuChamp := tcTrim
    else if TypeChamp = 'COMBO'           then TypeDuChamp := tcTrim
    else if TypeChamp = 'BOOLEAN'         then TypeDuChamp := tcBooleanON
    else if TypeChamp = 'DATE'            then TypeDuChamp := tcDate8JMA
    else if TypeChamp = 'INTEGER'         then TypeDuChamp := tcEntier
    else if TypeChamp = 'DOUBLE'          then TypeDuChamp := tcDouble
    else if TypeChamp = 'RATE'            then TypeDuChamp := tcDouble ;

    if (RecupInfo) then TobFilleRecup.PutFixedStValue (NomChamp, Valeur, 1, Length(Valeur), TypeDuChamp, false)
    else if Valdefaut <> '' then
    begin
      if TypeChamp = 'DATE' then  TobFilleRecup.PutValue (NomChamp, StrToDate (Valdefaut))
      else TobFilleRecup.PutValue (NomChamp, Valdefaut);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Procédure d'intégration d'une pièce
Mots clefs ... : IMPORT PIECE
*****************************************************************}
procedure TTImportGPAO.IntegrePiece (var Info : string) ;
var io             : TIOErr ;
    DebutLigne     : string  ;
    DebutLigneNext : string  ;
    RefPiece       : string  ;
    RefInterne     : string  ;
    FinPiece       : boolean ;
    Resultat       : boolean ;
    First          : boolean ;
    TobRecup       : TOB     ;
    NbreErrPiece   : integer ;
begin
  //Resultat := False ;
  /////////////////////////////////////////////////////
  // Création des TOB Virtuelles rattachées à la pièce
  /////////////////////////////////////////////////////
  Tob_Article   := TOB.CREATE ('Les Articles', nil, -1);
  Tob_PiedBase  := TOB.CREATE ('Les Taxes' , nil, -1);
  Tob_PiedEche  := TOB.CREATE ('Les Echéances', nil, -1);
  Tob_Acomptes  := TOB.CREATE ('Les Acomptes', nil, -1);
  Tob_Catalogue := TOB.CREATE ('Les Catalogues', nil, -1);
  Tob_noment    := TOB.CREATE ('Les Noments', nil, -1);

  NomTable     := Tob_RepriseGPAO.GetValue ('GRE_NOMTABLE');
  NbreErrPiece := 0 ;
  //
  // Sauvegarde de la position de début pour topage ultérieur des lignes
  //
  PosTopePiece := PosTope ;
  DebutLigne   := copy (Info, 1, 3);
  FinPiece     := False ;
  First        := True ;

  repeat
    //
    // Création de la TOB de Récupération
    //
    TobRecup := TOB.Create ('PIECE', nil, -1);
    //
    // Initialisation de la TOB à partir du fichier en tenant compte du paramétrage de la récupération
    //
    FileToTob (Info, TobRecup, True) ;

    if first then
    begin
      RefInterne := TobRecup.GetValue ('GP_REFINTERNE');
      first      := False ;
    end;

    //
    // Change-t-on de pièce ?
    //
    RefPiece := TobRecup.GetValue ('GP_REFINTERNE');
    if RefPiece <> RefInterne then FinPiece := True ;

    if not finpiece then
    begin
      //
      // Lancement des contrôles
      //
      if ControleInfoOK (NomTable, TobRecup) then
      begin
        Resultat := True ;
        if (Tob_Piece = nil) and (NbreErrPiece = 0) then
        begin
          if InitEntetePiece (TobRecup) then Resultat := InitLignePiece (TobRecup)
          else Resultat := False;
        end
        else begin
          if NbreErrPiece = 0 then Resultat   := InitLignePiece (TobRecup);
        end;

        if not Resultat then
        begin
          Inc (CptErr);                     // Nombre d'erreurs dans le fichier
          inc (NbreErrPiece);               // Nombre d'erreurs de la pièce
        end;
      end
      else begin
        Inc (CptErr);                     // Nombre d'erreurs dans le fichier
        inc (NbreErrPiece);               // Nombre d'erreurs de la pièce
      end;
    end;

    // Libération de la TOB de récupération des infos fichiers
    FreeAndNil( TobRecup );

    if not FinPiece then
    begin
      //
      // Topage temporaire de l'enregistrement
      //
      TopageEnregistrement ('T');
      //
      // Lecture de l'enregistrement suivant
      //
      Info           := LectureEnregistrement ;
      DebutLigneNext := copy (Info, 1, 3) ;

      if DebutLigneNext <> DebutLigne then
      begin
        if DebutLigneNext <> '' then
        begin
          Dec (CptEnregTot) ;
          Dec (CptEnreg) ;
          //Dec (CptTrt) ;
        end;
        FinPiece := True;
      end;
    end else
    begin
      Dec (CptEnregTot) ;
      Dec (CptEnreg) ;
      //Dec (CptTrt) ;
    end;
  until FinPiece;

  //
  // Repositionnement sur le prochain enregistremebt non traité
  //
  PosLecture := PosTope ;
  //
  // Si TOB chargée et pas d'erreur, on importe
  //
  if (Tob_Piece <> nil) and (NbreErrPiece = 0) then
  begin
    PosLecture := PosTope;
    io:=Transactions (ImportPiece, 0) ;
    if io<>oeOk then
    begin
      EcrireInfo ('' , True, False, False);
      EcrireInfo (TraduireMemoire('  Attention, problème transactionel lors de l''import de la pièce') , True, False, False);
      EcrireInfo ('' , True, False, False);
//      NbreErrPiece := 1 ;
      inc (CptErr) ;
      TopagePiece ('R') ;
    end else TopagePiece ('I');
  end else
  begin
    TopagePiece ('R');
  end;

  // Libération des TOB
  FreeAndNil( Tob_Piece     );
  FreeAndNil( Tob_Article   );
  FreeAndNil( Tob_PiedBase  );
  FreeAndNil( Tob_PiedEche  );
  FreeAndNil( Tob_Acomptes  );
  FreeAndNil( Tob_Tiers     );
  FreeAndNil( Tob_Adresses  );
  FreeAndNil( Tob_Catalogue );
  FreeAndNil( Tob_Noment    );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Intégration de données
Mots clefs ... :
*****************************************************************}
function TTImportGPAO.IntegreInfoGPAO (var Info : string; Entete : boolean) : boolean ;
var TobFilleRecup : TOB ;
begin

  if Tob_InfoGPAO = nil then Tob_InfoGPAO := TOB.Create ('INFO GPAO', nil, -1);
  //
  // Création de la TOB Fille contenant l'enregistrement
  //
  NomTable      := Tob_RepriseGPAO.GetValue ('GRE_NOMTABLE');
  TobFilleRecup := TOB.Create (NomTable, nil, -1);
  //
  // Initialisation de la TOB à partir du fichier en tenant compte du paramétrage de la récupération
  //
  FileToTob (info, TobFilleRecup, Entete);
  //
  // MAJ / Ajout specifique des champs
  //
  MAJSpecifInfo (NomTable, TobFilleRecup);
  //
  // Contrôle des données
  //
  if ControleInfoOK (NomTable, TobFilleRecup) then
  begin
    if TobFilleRecup <> nil then
    begin
      TobFilleRecup.ChangeParent (Tob_InfoGPAO, -1);
    end;
    //GCVoirTob (Tob_InfoGPAO);
    Result := True ;
  end else
  begin
    FreeAndNil( TobFilleRecup );
    Result := False ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Initialisation des variables avant import
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.InitImportGPAO ;
var i : integer ;
begin
  CptEnregTot   := 0 ;
  CptEnreg      := 0 ;
  CptMAJ        := 0 ;
  CptErr        := 0 ;
  CptTrt        := 0 ;
  CptIntegre    := 0 ;
  // Import des tarifs
  Tob_Tarif     := nil ;
  Periode       := ''  ;
  CodeTarif     := ''  ;
  MaxTarif      := 0   ;
  DernierArticleFlagPxUnique := '';
  // TOB
  Tob_Piece     := nil;
  Tob_Article   := nil;
  Tob_Tiers     := nil;
  Tob_PiedBase  := nil;
  Tob_PiedEche  := nil;
  Tob_Adresses  := nil;
  Tob_Acomptes  := nil;
  Tob_Catalogue := nil ;
  Tob_Noment    := nil ;
  // Tests pour articles
  DernierFournisseur   := '' ;
  DernierTiers         := '' ;
  DerniereCollection   := '' ;
  DernierCodeProgramme := '' ;
  //CptFid:=1;
  for i:=1 to 8  do DerniereFamille   [i] := '' ;
  for i:=1 to 10 do DerniereStatArt   [i] := '' ;
  for i:=1 to 2  do DerniereTaxe      [i] := '' ;
  for i:=1 to 5  do DerniereGrille    [i] := '' ;
  for i:=1 to 2  do DerniereStartCompl[i] := '' ;
  DernierTypeTarif       := '' ;
  DernierStatut          := '' ;
  DernierTypeArticle     := '' ;
  DernierArticleGen      := '' ;
  DernierArticleGenCompl := '' ;
  // Tests pour tarifs
  DernierDepot           := '' ;
  DerniereDevise         := '' ;
  DernierArticlePA       := '' ;
  DernierArticlePV       := '' ;
  // Tests pour les tiers
  DernierPays            := '' ;
  DerniereLangue         := '' ;
  DerniereZoneCom        := '' ;
  DernierTarifTiers      := '' ;
  DernierModeRegle       := '' ;
  DernierRegimeTVA       := '' ;
  DernierJuridique       := '' ;
  DerniereNationalite    := '' ;
  DernierComptaTiers     := '' ;
  // Inventaires
  CodeTrans              := '' ;
  NumLigneInv            := 0  ;
  EnteteInvCree          := False;
  // Pièces
  DernierClientCree      := '' ;
  DerniereNature         := '' ;
  DernierEtab            := '' ;
  DernierRefInterne      := '' ;
  LibellePiece           := '' ;
  IdentArticleImport     := '' ;
  TiersPiecePrecedente   := '' ;
  ReprePiecePrecedente   := '' ;
  DerniereNaturePrec     := '' ;
  DerniereReferencePrec  := '' ;
  // Table de correspondance
  Tob_CorrespGPAO := TOB.Create ('Tables de correspondance', Nil, -1) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Libération des TOB en fin d'impôrt
Mots clefs ... :
//*****************************************************************}
procedure TTImportGPAO.FinImportGPAO ;
begin

  // TOB stockant le format d'import
  if Tob_InfoGPAO <> nil then ImportTobInfoGPAO;
  FreeAndNil( Tob_InfoGPAO    );
  FreeAndNil( Tob_RepriseGPAO );
  FreeAndNil( Tob_CorrespGPAO );

  // Libération des TOB de l'objet recup
  FreeAndNil( Tob_Piece       );
  FreeAndNil( Tob_Tiers       );
  FreeAndNil( Tob_Article     );
  FreeAndNil( Tob_PiedBase    );
  FreeAndNil( Tob_PiedEche    );
  FreeAndNil( Tob_Adresses    );
  FreeAndNil( Tob_Acomptes    );
  FreeAndNil( Tob_Catalogue   );
  FreeAndNil( Tob_Noment      );
  FreeAndNil( Tob_TiersDefaut );
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Traitement du fichier d'entrée
Mots clefs ... : IMPORT FICHIER
*****************************************************************}
procedure TTImportGPAO.RecupInfoGPAO (Nomfichier : string) ;
var Chtmp        : string  ;
    IdentInfo    : string  ;
    OldIdentInfo : string  ;
    Info         : string  ;
    SQL          : string  ;
    Q            : TQuery  ;
    ChargeParam  : boolean ;
    NbEnregTot   : integer ;
    Transfo      : TTTransformeGPAO;
begin
inherited;

  Chtmp := Nomfichier;
  if FileExists (Chtmp) then
  begin

    EcrireInfo ('  ', True, False, False);
    EcrireInfo (TraduireMemoire(' Début de la récupération des informations le ') + DateToStr(Date) + TraduireMemoire(' à ') + TimeToStr(Time), True, True, True);

    //Transformation du fichier si nécessaire
    if TransfoType.Text <> TraduireMemoire('<<Aucune>>') then
    begin
      EcrireInfo (TraduireMemoire(' Fichier transformé : ') + Nomfichier, True, True, True);
      EcrireInfo ('  ', True, False, False);
      SQL := 'SELECT * FROM REPRISETFGPAO WHERE GRT_TYPETRANSFO="'+TransfoType.Value+'"' ;
      Q   := OpenSQL (SQL ,True) ;
      Transfo := TTTransformeGPAO.Create( Application );
      try
        if not Transfo.TransformationGPAO( TransfoType.Value, Chtmp, Q.FindField('GRT_IDENTETE').AsString, Q.FindField('GRT_IDLIGNE').AsString, '', Q.FindField('GRT_IDOUT').AsString, false) then
        begin
          EcrireInfo ('  La transformation a échouée', True, True, True);
          exit;
        end;
      finally
        FreeAndNil( Transfo );
        Ferme(Q);
      end;
    end;

    EcrireInfo (TraduireMemoire(' Fichier intégré : ') + Nomfichier, True, True, True);
    EcrireInfo ('  ', True, False, False);

    // Initialisation de la barre de progression
    NbEnregTot := GetNbEnregInFile( Chtmp );
    InitMove( NbEnregTot, '');
    NbreTotal.Text := IntToStr(NbEnregTot) ;

    // Ouverture du fichier d'entrée en mode RW
    AssignFile(Fichier, Chtmp);
    FileMode:=2;
    Reset (Fichier);
    try
      // Initialisation des variables
      InitImportGPAO;
      OldIdentInfo  := '';
      ChargeParam   := False;
      DebutFichier  := True ;

      while not EOF(Fichier) and not ArretRecup do
      begin
        MoveCur (False);
        // Lecture du prochain enregistrement à intégrer
        Info := LectureEnregistrement ;

        Application.ProcessMessages ;

        if Info <> '' then
        begin
          IdentInfo := copy (Info, 1, 3);

          // Chargement du paramétrage de la récup si changement d'enregistrement
          if (OldIdentInfo <> IdentInfo) then
          begin
            ChargeParam := ChargeParamReprise (Info);
            if not ChargeParam then EcrireInfo (TraduireMemoire('  ATTENTION : les enregistrements de type "') + IdentInfo + TraduireMemoire('" ne sont pas repris'), True, False, False);
          end;

          // Traitement de l'enregistrement
          if ChargeParam then IntegreGPAO (Info)
          else TopageEnregistrement ('R');

          OldIdentInfo := IdentInfo ;
        end;
      end;
    finally
      // Libération des TOB
      FinImportGPAO ;

      // Fermeture du fichier d'entrée
      CloseFile(Fichier);
      FiniMove ();
    end;
    // MAJ des compteurs
    AfficheCompteur;

    // Affichage du compte rendu d'intégration
    AfficheFinTraitement ;
  end else EcrireInfo (TraduireMemoire('Le Fichier ') + Chtmp + TraduireMemoire(' n''existe pas'), True, True, True);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Traitement du fichier d'entrée
Mots clefs ... : IMPORT FICHIER
*****************************************************************}
procedure TTImportGPAO.RecupInfoGPAOAuto (Entete : boolean)  ;
var Repertoire          : string  ;
    RepertoireArch      : string  ;
    RepertoireRejet     : string  ;
    i                   : integer ;
    FichierTraite       : string  ;
    FichierTraiteCompl  : string  ;
begin
inherited;
  Repertoire      := FichTrt.Text ;
  RepertoireArch  := GetParamSoc ('SO_GCREPORLIARCH');
  RepertoireRejet := GetParamSoc ('SO_GCREPORLIREJET');

  if (Repertoire <> '') and (DirectoryExists (Repertoire)) then
  begin
    // Affiche la liste des fichiers à intégrer
    AfficheRepORLI ;
    for i:=0 to ListeFichier.items.Count-1 do
    begin
      FichierTraite      := ListeFichier.Items.Strings[i];
      FichierTraiteCompl := Repertoire + '\' + FichierTraite;

      if Entete then RecupInfoGPAO (FichierTraiteCompl)
      else RecupInfoGPAOFixe (FichierTraiteCompl);

      if not (ArretRecup) then
        //
        // Si pas d'erreur, on archive le fichier
        //
        if (CptErr = 0) then
        begin
          // on copie le fichier sélectionné dans la corbeille d'archivage
          if CopyFile ( pchar(FichierTraiteCompl), pchar(RepertoireArch + '\' +ExtractFileName(FichierTraite)), False) then
          // et on le supprime de la corbeille
          DeleteFile(pchar(Repertoire + '\' + FichierTraite)) ;
        end
        else
        begin
          // on copie le fichier sélectionné dans la corbeille de rejet
          if CopyFile ( pchar(FichierTraiteCompl), pchar(RepertoireRejet + '\' +ExtractFileName(FichierTraite)), False) then
          // et on le supprime de la corbeille
          DeleteFile(pchar(Repertoire + '\' + FichierTraite)) ;
        end
      else break ;
    end;

    // Affiche la liste des fichiers restant à intégrer
    AfficheRepORLI ;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Traitement du fichier d'entrée
Mots clefs ... : IMPORT FICHIER
*****************************************************************}
procedure TTImportGPAO.RecupInfoGPAOFixe (Nomfichier  : string) ;
var Chtmp         : string  ;
    Info          : string  ;
    FichierFinTrt : string  ;
begin
inherited;

  Chtmp := Nomfichier;
  if FileExists (Chtmp) then
  begin
    EcrireInfo ('  ', True, False, False);
    EcrireInfo (TraduireMemoire(' Début de la récupération des informations le ') + DateToStr(Date) + TraduireMemoire(' à ') + TimeToStr(Time), True, False, False);
    EcrireInfo (TraduireMemoire(' Fichier intégré sans topage : ') + Nomfichier, True, False, False);
    EcrireInfo ('  ', True, False, False);

    // Chargement du format du fichier
    if ChargeParamReprisefixe then
    begin
      // Ouverture du fichier d'entrée en mode texte (lecture seule, pas de topage)
      AssignFile(FichierTexte, Chtmp);
      Reset (FichierTexte);                                 // Ouverture du fichier
      try
        // Initialisation des variables
        InitImportGPAO;
        DebutFichier  := True ;

        while not EOF(FichierTexte) and not ArretRecup do
        begin
          MoveCur (False);
          readln(FichierTexte, Info);
          Application.ProcessMessages ;
          if Info <> '' then IntegreGPAOFixe (Info);
        end;
      finally
        // Fermeture du fichier d'entrée
        CloseFile(FichierTexte);
      end;
    end;

    // MAJ de la dernière table et libération des TOB
    FinImportGPAO ;

    // MAJ des compteurs
    FiniMove ();
    AfficheCompteur;

    // Affichage du compte rendu d'intégration
    AfficheFinTraitement ;

    // Le fichier est renommé en .old.
    FichierFinTrt := Nomfichier + '.old' ;
    EcrireInfo (' ', True, True, True);
    EcrireInfo (TraduireMemoire(' Le Fichier ') + Nomfichier + TraduireMemoire(' est sauvegardé en ' + FichierFinTrt), True, True, True);
    if CopyFile ( pchar(Nomfichier), pchar(FichierFinTrt), False) then
    // et on le supprime de la corbeille
    DeleteFile(pchar(Nomfichier)) ;

  end else EcrireInfo (TraduireMemoire(' Le Fichier ') + Chtmp + TraduireMemoire(' n''existe pas'), True, True, True);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Lancement de l'intégration
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.IntegrationGPAO;
begin
inherited;

  NbrAno:=0;
  BRecup.enabled := False;
  BoutonDetope.enabled := False;
  BStop.enabled := True;
  Application.ProcessMessages ;
  ArretRecup := False;

  //  Création de la TOB des grilles de tailles et lancement de la récupération
  if IntegAuto.checked = False then
  begin
    if RepriseFichierEnteteFixe then RecupInfoGPAO (FichTrt.text)
    else RecupInfoGPAOFixe (FichTrt.text);
  end else RecupInfoGPAOAuto (RepriseFichierEnteteFixe);

  BRecup.enabled := True;
  if (not IntegAuto.checked) and (ReprisefichierEnteteFixe) then BoutonDetope.enabled := True;
  BStop.enabled := False;

  if NbrAno=1 then CompteRendu.lines.add(#13#10'ATTENTION : 1 fichier n''a pas été intégré.')
  else if NbrAno>1 then CompteRendu.lines.add(#13#10'ATTENTION : '+IntToStr(NbrAno)+' fichiers n''ont pas été intégrés.');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2003
Modifié le ... : 23/07/2003
Description .. : Déconnexion de la base
Mots clefs ... : FO
*****************************************************************}

procedure TTImportGPAO.Deconnexion ;
begin
  if not ConnexionBase then Exit ;

  NomDossier := V_PGI.CurrentAlias ;
  MultiUserLogin := V_PGI.MultiUserLogin ;
  EcrireInfo (TraduireMemoire('Déconnexion de la société ')+ NomDossier + TraduireMemoire(' le ') + DateToStr(Date)+ TraduireMemoire(' à ') + TimeToStr(Time), True, False, False) ;
  DeconnecteHalley ;
  ConnexionBase := False ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2003
Modifié le ... : 23/07/2003
Description .. : Reconnexion à la base
Mots clefs ... : FO
*****************************************************************}
procedure TTImportGPAO.Connexion ;
begin
  if ConnexionBase then Exit ;

  EcrireInfo (TraduireMemoire('Connexion de la société ')+ NomDossier + TraduireMemoire(' le ') + DateToStr(Date)+ TraduireMemoire(' à ') + TimeToStr(Time), True, True, True) ;
  V_PGI.MultiUserLogin := True ;
  if ConnecteHalley(NomDossier, False, nil, nil, nil, nil, False, False) then
    ConnexionBase := True
  else
    EcrireInfo (TraduireMemoire('Erreur de connexion n°')+ IntToStr(V_PGI.ErreurLogin), True, True, True) ;
  V_PGI.MultiUserLogin := MultiUserLogin ;
  InitDateConnect ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Détermine la date de déconnexion et celle de reconnexion
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.InitDateConnect ;
var
  Heure : TDateTime ;
begin
  // heure de reconnexion
  Heure := GetParamSoc('SO_GCTOXCONNECT') ;
  if Heure <= 0        then DateConnexion := iDate2099
  else if Heure < Time then DateConnexion := Date + 1 + Heure
  else DateConnexion := Date + Heure ;

  // heure de déconnexion
  Heure := GetParamSoc('SO_GCTOXDECONNECT') ;
  if Heure <= 0        then DateDeconnexion := iDate2099
  else if Heure < Time then DateDeconnexion := Date + 1 + Heure
  else DateDeconnexion := Date + Heure ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Lancement de l'intégration
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.bRecupClick(Sender: TObject);
begin
inherited;

  // Ouverure du fichier des anomalies
  AssignFile(FichierAno, FichAno.text);
  if FileExists(FichAno.text) then Append(FichierAno) else ReWrite(FichierAno);
  Writeln(FichierAno,'--- Reprise du '+DateToStr(Date)+' à '+TimeToStr(Time));

  // Ouverture du fichier Compte rendu
  AssignFile(FichierTrt, FichCpt.text);
  ReWrite(FichierTrt);
  Writeln(FichierTrt,'--- Reprise du '+DateToStr(Date)+' à '+TimeToStr(Time));

  try
    if IntegAuto.checked and IntegBoucle.checked then
    begin
      if NowH >= DateDeconnexion then Deconnexion;
      if NowH >= DateConnexion   then Connexion;
      Timer1.Enabled := False ;
      if ConnexionBase then IntegrationGPAO;
      Timer1.Enabled := True ;
    end else IntegrationGPAO;
  finally
    CloseFile(FichierAno);
    CloseFile(FichierTrt);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Interruption de l'intégration
Mots clefs ... :
*****************************************************************}
procedure TTImportGPAO.BStopClick(Sender: TObject);
begin
inherited;
  ArretRecup := TRUE;
  BStop.enabled := False;
  Timer1.Enabled := False ;
  EcrireInfo (TraduireMemoire(' La récupération a été interrompue !'), True, True, False);
  Application.ProcessMessages ;
end;

procedure TTImportGPAO.bSuivantClick(Sender: TObject);
begin
inherited;
  if (P.ActivePage.PageIndex = P.PageCount - 1)
    then bFin.enabled := True;
end;

procedure TTImportGPAO.bPrecedentClick(Sender: TObject);
begin
inherited;
  bFin.enabled := False;
end;

procedure TTImportGPAO.FichAnoElipsisClick(Sender: TObject);
begin
inherited;
  OpenfichAno.InitialDir:=GetParamSoc ('SO_GCREPORLI');
  if OpenFichAno.execute then FichAno.text:=OpenFichAno.FileName;
end;

procedure TTImportGPAO.FichCptElipsisClick(Sender: TObject);
begin
inherited;
  OpenfichCpt.InitialDir:=GetParamSoc ('SO_GCREPORLI');
  if OpenFichCpt.execute then FichCpt.text:=OpenFichCpt.FileName;
end;

procedure TTImportGPAO.FichTrtElipsisClick(Sender: TObject);
begin
inherited;
  OpenfichGPAO.InitialDir:=GetParamSoc ('SO_GCREPORLI');
  if IntegAuto.Checked then
  begin
    FichTrt.text:=GetParamSoc ('SO_GCREPORLI');
    FichTrt.DataType     := 'DIRECTORY';
    FichTrt.Enabled      := False ;
    BoutonDetope.Enabled := False ;
  end else
  begin
    FichTrt.DataType     := 'FILE' ;
    FichTrt.Enabled      := True ;
    BoutonDetope.Enabled := True ;
    if OpenFichGPAO.execute then FichTrt.text:=OpenFichGPAO.FileName;
  end;
end;

procedure TTImportGPAO.bFinClick(Sender: TObject);
begin
inherited;
  if not isInside(Self) then close ;
end;

procedure TTImportGPAO.FormShow(Sender: TObject);
begin
inherited;
  ConnexionBase := True ;
  InitDateConnect ;

  // Appel de la fonction de dépilage dans la liste des fiches
  AglEmpileFiche(Self) ;

  // Lecture du répertoire indiqué dans les paramsoc
  FichTrt.Text := GetParamSoc ('SO_GCREPORLI') ;
  FichAno.Text := GetParamSoc ('SO_GCREPORLI') + '\GPAOAno.txt' ;
  FichCpt.Text := GetParamSoc ('SO_GCREPORLI') + '\GPAOCpt.txt' ;

  if Inventaire then
  begin
    Provenance.value   := '...' ;
    Provenance.enabled := False ;
    TitreAssistant.Caption := TraduireMemoire ('      REPRISE D''UN FICHIER INVENTAIRE');
    Caption := TraduireMemoire ('Reprise inventoriste') ;
  end else Provenance.value := GetParamSoc ('SO_IMPORTGPAOPRO') ;
  if Provenance.value='' then Provenance.value := 'DEF' ;

  IntegAuto.Checked := GetParamSoc ('SO_IMPORTGPAOAUTO');
  IntegBoucle.Checked := GetParamSoc ('SO_IMPORTGPAOBOUCLE');
  TempsBoucle.Text := GetParamSoc ('SO_IMPORTGPAOTEMPS');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : L. Meunier
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Detopage du fichier d'entrée avant import
Mots clefs ... : DETOPAGE;GPAO
*****************************************************************}
procedure TTImportGPAO.BoutonDetopeClick(Sender: TObject);
var Fichier    : file of char;
    Chtmp      : string  ;
    Info       : string  ;
    PosTope    : longint ;
    PosLecture : longint ;
    CharTope   : char    ;
    Tampon     : char    ;
    NumRead    : integer ;
    NumWrite   : integer ;
    Cpt        : integer ;
begin
inherited;

  BRecup.enabled := False;
  BStop.enabled  := False;
  if Timer1.enabled then Timer1.enabled := False;

  Cpt := 0 ;
  Chtmp:=FichTrt.text;
  if FileExists (Chtmp) then
  begin
    EcrireInfo (TraduireMemoire(' Début lecture du fichier') , True, False, False);

    AssignFile(Fichier, Chtmp);
    FileMode:=2;
    Reset (Fichier);
    try
      while not EOF(Fichier) and not ArretRecup do
      begin
        Info := '';

        PosTope := FilePos (Fichier);

        repeat
          BlockRead(Fichier, Tampon, 1, NumRead);
          if Tampon <> #10 then Info := info + Tampon;
        until (NumRead = 0) or (Tampon = #13);

        if length (info) > 6 then
        begin
          Inc (Cpt);
          PosLecture := FilePos (Fichier);

          if PosTope = 0 then PosTope := PosTope + 5
          else PosTope := PosTope+6;
          seek (Fichier, PosTope);
          CharTope := ' ';

          BlockWrite(Fichier, CharTope, 1, NumWrite);
          seek (Fichier, PosLecture);
        end;
      end;
    finally
      CloseFile(Fichier);
    end;
    EcrireInfo (' ----> ' + IntToStr (Cpt) + TraduireMemoire(' enregistrements lus') , True, False, False);
    EcrireInfo (TraduireMemoire(' Fin lecture du fichier') , True, False, False);
  end;

  BRecup.enabled := True;
  BStop.enabled  := False;
end;

procedure TTImportGPAO.IntegAutoClick(Sender: TObject);
begin
inherited;
  if IntegAuto.Checked then
  begin
    IntegBoucle.Enabled  := True ;
    TempsBoucle.Enabled  := True ;
    Timer1.Enabled       := False ;
    Timer1.Interval      := TempsBoucle.Value * 60 * 1000 ;

    FichTrt.DataType     := 'DIRECTORY';
    FichTrt.text:=GetParamSoc ('SO_GCREPORLI');
    FichTrt.Enabled      := False ;
    BoutonDetope.Enabled := False ;

    CompteRendu.Height   := 210  ;
    CompteRendu.Left     := 136  ;
    CompteRendu.Width    := 409  ;
    CompteRendu.Top      := 8    ;
    ListeFichier.Visible := True ;

    AfficheRepORLI ;
  end else
  begin
    IntegBoucle.Enabled  := False ;
    TempsBoucle.Enabled  := False ;
    Timer1.Enabled       := False ;
    Timer1.Interval      := TempsBoucle.Value * 60 * 1000 ;

    FichTrt.DataType     := 'FILE' ;
    FichTrt.Enabled      := True ;
    BoutonDetope.Enabled := True ;

    CompteRendu.Height   := 210 ;
    CompteRendu.Left     := 0   ;
    CompteRendu.Width    := 545 ;
    CompteRendu.Top      := 8   ;
    ListeFichier.Visible := False ;
  end;
end;

procedure TTImportGPAO.ListeFichierDblClick(Sender: TObject);
var F     : file of byte ;
    iSize : integer ;
begin
inherited;
  AssignFile (F, TFileListBox (Sender).FileName);
  Reset(F);
  iSize := FileSize(F);
  CloseFile (F) ;
  InfoTox ( ExtractFileName (TFileListBox(Sender).FileName), iSize ) ;
end;

procedure TTImportGPAO.TempsBoucleChange(Sender: TObject);
begin
inherited;
  Timer1.Enabled := False ;
  Timer1.Interval := TempsBoucle.Value * 60 * 1000 ;
end;

procedure TTImportGPAO.Timer1Timer(Sender: TObject);
begin
inherited;
  bRecupClick(Sender);
end;

procedure TTImportGPAO.FormDestroy(Sender: TObject);
begin
  inherited ;
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

procedure TTImportGPAO.ProvenanceChange(Sender: TObject);
begin
  inherited;

  //
  // Reprise bridée pour inventaire
  //
  if inventaire then
  begin
    IntegAuto.Checked   := False;
    IntegAuto.Enabled   := False;
    IntegBoucle.Checked := False;
    IntegBoucle.Enabled := False;
    TempsBoucle.Enabled := False ;

    Timer1.Enabled       := False ;
    Timer1.Interval      := TempsBoucle.Value * 60 * 1000 ;

    FichTrt.DataType     := 'FILE' ;
    FichTrt.Enabled      := True ;

    CompteRendu.Height   := 210 ;
    CompteRendu.Left     := 0   ;
    CompteRendu.Width    := 545 ;
    CompteRendu.Top      := 8   ;
    ListeFichier.Visible := False ;
    //
    // Bouton détopage actif ou pas ?
    //
    if not ReprisefichierEnteteFixe then BoutonDetope.Enabled := False
    else BoutonDetope.Enabled := True ;
  end else
  begin
    if not ReprisefichierEnteteFixe then
    begin
      IntegAuto.Checked   := False;
      IntegAuto.Enabled   := False;
      IntegBoucle.Checked := False;
      IntegBoucle.Enabled := False;
      TempsBoucle.Enabled := False ;

      Timer1.Enabled       := False ;
      Timer1.Interval      := TempsBoucle.Value * 60 * 1000 ;

      FichTrt.DataType     := 'FILE' ;
      FichTrt.Enabled      := True ;
      BoutonDetope.Enabled := False ;           // Impossible de la détoper

      CompteRendu.Height   := 210 ;
      CompteRendu.Left     := 0   ;
      CompteRendu.Width    := 545 ;
      CompteRendu.Top      := 8   ;
      ListeFichier.Visible := False ;
    end else
    begin
      // Lecture du dernièr répertoire utilisé dans la registry
      TempsBoucle.value   := 120   ;

      if IntegAuto.Checked then
      begin
        IntegAuto.Enabled    := True ;
        IntegBoucle.Enabled  := True ;
        TempsBoucle.Enabled  := True ;
        Timer1.Enabled := False ;
        Timer1.Interval := TempsBoucle.Value * 60 * 1000 ;

        FichTrt.DataType     := 'DIRECTORY';
        FichTrt.text:=GetParamSoc ('SO_GCREPORLI');
        FichTrt.Enabled      := False ;
        BoutonDetope.Enabled := False ;

        CompteRendu.Height   := 210  ;
        CompteRendu.Left     := 136  ;
        CompteRendu.Width    := 409  ;
        CompteRendu.Top      := 8    ;
        ListeFichier.Visible := True ;
        AfficheRepORLI ;
      end else
      begin
        IntegAuto.Enabled    := True ;
        Timer1.Enabled       := False ;
        Timer1.Interval      := TempsBoucle.Value * 60 * 1000 ;
        IntegBoucle.Enabled  := False ;
        TempsBoucle.Enabled  := False ;

        FichTrt.DataType     := 'FILE' ;
        FichTrt.Enabled      := True ;
        BoutonDetope.Enabled := True ;

        CompteRendu.Height   := 210 ;
        CompteRendu.Left     := 0   ;
        CompteRendu.Width    := 545 ;
        CompteRendu.Top      := 8   ;
        ListeFichier.Visible := False ;
      end;
    end;
  end;
end;

procedure TTImportGPAO.FormCloseQuery(Sender: TObject;
var CanClose: Boolean);
begin
inherited;
  Connexion;
end;

end.

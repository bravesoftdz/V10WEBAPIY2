{***********UNITE*************************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 02/02/2000
Modifié le ... : 10/09/2002
Description .. : TOM pour ARTICLE et tables associées :
Suite ........ :    - CATALOGU
Suite ........ :    - ARTICLELIE
Suite ........ :    - ARTICLEPIECE
Suite ........ :    - CODECPTA
Suite ........ :    - CONDITIONNEMENT
Suite ........ :    - PROFILART
Suite ........ :
Suite ........ : Arguments TOM_ARTICLE :
Suite ........ :    TARIF=N  : pas d'appel aux fiches tarifs
Suite ........ :    DUPLICATION=CodeArticle : ouvre la fiche en insert
Suite ........ : avec initialisation avec les éléments de la fiche de
Suite ........ : CodeArticle
Suite ........ :
Suite ........ :
Suite ........ :
Suite ........ :
Mots clefs ... : ARTICLE;TOM;CONDITIONNEMENT;PROFILART;CODECPTA;ARTICLELIE;CATALOGU;ARTICLEPIECE
*****************************************************************}
unit UTomArticle;

interface

uses {$IFDEF VER150} variants,{$ENDIF} Classes, Utom, UTob, UTof, HDimension, UDimArticle, sysutils, LookUp, HEnt1, HCtrls,
  UtilArticle, FACTUTIL, menus, Entgc, AglInit, UtilGC, nomenlig,
  {$IFDEF GPAO}
  wCommuns,
  wConversionUnite,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  UtileAGL, eFiche, eFichList, maineagl, Spin,
  {$ELSE}
  DBCtrls, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FichList, Fiche, Fe_Main, MAJTable, HDB, AglIsoflex,
  {$IFDEF V530}
  EdtEtat,
  {$ELSE}
  EdtREtat,
  {$ENDIF}
  {$ENDIF}

  UtilTarif,TarifCliart,olectnrs, UtilXlsBTP,FileCtrl,

  {$IFDEF AFFAIRE}
  afUtilArticle,
  {$ENDIF}
  graphics, HsysMenu, Ent1,
  {$IFNDEF CCS3}
  UtilConfid,
  {$ENDIF}
  HMsgBox, forms, extctrls, ComCtrls, choix, UtilPGI, controls, AglInitGC,
  StdCtrls, TarifUtil, M3FP, UtilDimArticle, UtofMBOParamDim, HTB97, windows,
  HQRY, ARTPHOTO_TOF, HRichOLE, UTobDebug,
  UtilsMetresXLS;

type
  TOM_Catalogu = class(TOM)
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnDeleteRecord; override;
    procedure OnArgument(st: string); override;
    procedure AffectationReference;
    procedure CodeUnique;
    procedure SetLastError(Num: integer; ou: string);
  private
    ArticleAffecte, OldArticle: string;
    ArticleParam: string;
    FournParam : string;
    UniteParam : string;
    prixParam : Double;
    Connexion: boolean;
    BMonoFournisseur: boolean;
    //
    _ARTICLE_GEN    : THEdit;

    function FindUniteSto(Article: string): string;
    procedure AfficheCoefConv;
    procedure CacheAndZeroCoefConv;
    procedure ARTGEN_OnElipsisClick(Sender: TObject);
    procedure ARTGEN_OnExit(Sender: TObject);
    procedure BTCONDITClick (Sender : TObject);
  end;

  {$IFNDEF CCS3}
  TOM_ArticlePiece = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnNewRecord; override;
  end;
  {$ENDIF}

  TOM_ArticleLie = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
  end;

  TOM_Article = class(TOM)
    DimensionsArticle: TODimArticle;
  private
    Classeur        : TPageControl;
    OldRemplacement : string;
    OldSubstitution : string;
    OldFournPrinc   : string;
    OldCodeCAB      : string;
    AfficheTarif    : boolean;
    ModifparLot     : Boolean;
    StSQL           : string; //   Modif par lot
    //
    ClickActuPrix   : Boolean;
    //
    // Modif BTP
    //
    TabTarif : TTabSheet;
    //
    BCODEBARRE      : TToolBarButton97;
    BOuvrage        : TToolbarButton97;
    MnTariffou      : TMenuItem;
    MnFourn         : TMenuItem;
    Prefixe         : string;
    TOBCatalog      : TOB;
    TOBTarif        : TOB;
    TOBEDetail      : TOB;
    TOBDetail       : TOB;
    TOBPrestation   : TOB;
    TOBTiers        : TOB;
    TPrestation     : THEdit;
    DQtePrest       : THNumEdit;
    //
    // Gestion des Métrés
    MetreArticle    : TMetreArt;
    BAppelVariable  : TToolbarButton97;
    BExcel          : TToolbarButton97;
    // Gestion des nomenclatures
    BNomenclature   : TToolbarButton97;
    CodeArt : string;
    Ok_MetreActif : boolean;
    OK_MetreDoc		: Boolean;
    TEPrixPourQte : THEdit;
    CodeSousTarif : THValComboBox;
    ModifChampPr    : Boolean;
    SelTypeArticle  : THMultiValComboBox;
    //
    ModifChampHT    : Boolean;
    ModifChampTTC   : Boolean;
    FaireCalcul     : Boolean;
    CalculEnCours   : Boolean;

    ArtToDuplic     : string;
    DuplicFromFiche : Boolean;
    DuplicFromMenu  : boolean;
    DuplicArt       : string;
    Consultation    : boolean;

    TypeArticle     : string;
    TypeNomenc      : string;
    //AfficheToutesTailles : string;

    //FV1 : 08/01/2018 - FS#2929 - SES_ETANCHEITE SERVICE : En création article, ne prend pas 3 décimales pour le PU en UA
  	DecQte				  : String;
    DecPrix 			  : string;
   	NBDecPrix			  : Integer;
    NbdecQte 			  : integer;


    OldDimMasque      : string; // Ancien GA_DIMMASQUE : test si modif.
    TypeMasque        : string; // Type de masque utilisé passé en paramètre : "TYPEMASQ="
    Top, Left         : integer;
    FicheDimensionne  : Boolean; // Fiche article dimensionné
    bZeroDim          : Boolean; // Existe-t-il des articles dimensionnés pour l'article générique ?
    FicheStock        : Boolean; // Gestion des dimensions dans la fiche stock
    FicheStoVte       : Boolean; // Gestion des dimensions dans la fiche stock + lignes de vente
    TobArt            : TOB; // Tob des champs de la fiche article
    TobArtLoad        : TOB; // Tob des champs de la fiche article au OnLoadRecord
    TobZonelibre      : TOB; //Tob pour gérer les zones de la table ARTICLECOMPL
    TOBInfoArt        : TOB; //Création article à partir d'une photo
    NavigClick        : boolean; // Fiche article dimensionné: navigateur utilisé Oui/Non
    BMonoFournisseur  : boolean;
    CreatArticlePhoto : boolean; //Création article à partir d'une photo
    FournDocAch       : string;
    DerniereCreate    : string; // Contournement bug création+modif
    CalcPAPR          : string;
    TobLiensOle       : Tob; // Duplication article : table LIENSOLE
    SpecifGescom      : TOF;
    TOBARTICLEPARC    : TOB;
    { GPAO }

    IsLine                : boolean;
    bCalculPrixDuplicHT   : Boolean;
    bCalculPrixDuplicTTC  : Boolean; // Pour ne pas mettre de message 37 à l'ouverture de la fiche en duplication

    BACTUPRIX   : TToolBarButton97;
    CALCAUTOHT  : TCheckBox;
    CALCAUTOTTC : TCheckBox;
    DPRAUTO     : TCheckBox;

    RepServeur  : String;
    RepPoste    : String;
    ExistCatalog : Boolean;


    procedure AfterFormShow;
    procedure SetArguments(StSQL: string); //modif par lot
    procedure ArticleVoir(Sender: TObject);
    procedure ArticleActionDim(ItemDim: THDimensionItem);
    procedure InitDimensionsArticle;
    {$IFDEF EAGLCLIENT}
    procedure eAglInsertionOn; // DBR Fiche 10093
    {$ELSE}
    procedure InsertionOn(DataSet: TDataSet);
    procedure InsertionOff(DataSet: TDataSet);
    {$ENDIF}
    procedure OnDefaire(DataSet: TDataSet);
    procedure AffichePhoto;
    procedure AffecteDimension;
    procedure LookUpFournisseur;
    procedure PoseEvenementFournisseur;
    procedure OnElipsisClick_Fournisseur(Sender: TObject);
    procedure OnExit_Fournisseur(Sender: TObject);
    function  ControleModifFournisseur(AffBox: boolean): boolean;
    procedure CalculBaseTarifHT(Click: boolean);
    procedure CalculBaseTarifTTC(Click: boolean);
    procedure CalculCoefTarifHT(Click: boolean);
    procedure CalculCoefTarifTTC(Click: boolean);
    procedure TestFaireCalcul;
    procedure CodeUniqueArtLie(TypeLien: string);
    procedure SetLastError(Num: integer; ou: string);
    procedure FormateCAB(bCtrl: boolean);
    procedure ErgoGCS3;
    // spécif prestations
    procedure OnUpdatePrestation;
    // Modif BTP
    //
    procedure CalculBaseTarifPR(Click: boolean);
    procedure CalculCoefTarifPR(Click: boolean);
    function  TrimZoneCritique: boolean;
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SetFicheModifie;
    //
    //--- FV : Externalisation des onarguments pour lecture plus Facile !!!
    //
    procedure OnArgumentsCCS3(Arguments: String);
    procedure OnArgumentsBTP(Arguments: String);
    procedure OnArgumentsLine(Arguments: String);
    procedure ControleCritere(ChampMul, ValMul : String);

    //--- FV : Gestion des codes Barres
    procedure BCODEBARREOnClick(Sender: Tobject);
    procedure ChargeCodeBarre;
    //
    // ------
    procedure LibresArtToLibresLigne;
    function IsChampStockAffiche: boolean;
    procedure OnChangeOnglet(Sender: TObject);
    procedure ZonesLibArtIsModified;
    procedure ChangeStatArt(CodeStat: string);
    //
    // Gestion Isoflex
    //
    procedure GereIsoflex;

    procedure EntreeDetOuv(Sender: Tobject);
    procedure AfficheMarge(Indice: integer; PrixAchat, PrixVente, Prixrevient: double);
    procedure OuvreNomenclature(Sender: Tobject);
    procedure OuvreVariable(Sender: Tobject);
    procedure SousFamTarArtChange (sender : TObject);
    procedure DetermineValorisation(var pa, pr, pv: double);
    procedure GetInfosPrixTarif(CodeFournisseur, CodeArticle, TarifArticle, SousTarifArticle: string; TOBcatalog, TOBTarif: TOB);
    procedure SetInfosPrixTarif;
    procedure SetDetailPrixPose(Article, Libelle: string; QtePrest: double; TOBPrestation: TOB);
    procedure AffichageInfosPrixPose;
    procedure ChangePrestation(Prestation: string; Qte: double);
    procedure TPrestationElipsClick(Sender: TOBject);
    procedure TPrestationChange(Sender: TOBject);
    procedure ChangeQtePrest(Sender: TObject);
    procedure ReajusteTOBDetail;
    procedure MnTariffouClick(sender: Tobject);
    procedure CodeTarifArticleChange(Sender: Tobject);
    procedure LoadTobArticlecompl;
    procedure LoadTobArticlecompParc;
  	procedure EcritTobArticleCompl;
  	procedure EcritTobArticleCompParc;
  	procedure SetChangeCodeTarif (CodeTarif : string);
    procedure GestionFicheNormale;
    procedure OuvreNomenParc(sender: Tobject);
    procedure EntreeDetParcNomen(Sender: Tobject);
    procedure CODEINTERNEChange (Sender : TObject);
    procedure NbMoisGarantieChange (Sender : Tobject);
    procedure IsVersionClick (Sender : TObject);
    procedure ARTPARCCLick (Sender : TObject);
    procedure DuplicFromArt (Sender : TObject);
    procedure CalculPrixRef(CodeBasePR: String; var PrixRef: double);
    function RechercheQuotite(Unite: string): double;
    //FV1 : 12/02/2015 - Recalcul du Prix d'achat en UV à partir du Prix d'achat en UA
    function CalculPAHT: Double;
    procedure GestionFournisseurPrincipal;
    //FV1 : 09/04/2015
    procedure BACTUPRIX_OnClick(Sender: Tobject);
    procedure CALCAUTOHT_OnClick(Sender: Tobject);
    procedure CALCAUTOTTC_OnClick(Sender: Tobject);
    procedure DPRAUTO_OnClick(Sender: Tobject);
    procedure ExcelOnClick(Sender: Tobject);
    procedure DefiniAffichage;

    {$IFDEF GPAO}
    { Evènements des objets de la fiche "WARTICLE_FIC" }
    TobUnAutorizedFields: Tob;
    OnlyProfil: Boolean;
    procedure MnUniteConversion_OnClick(Sender: TObject);
    procedure MnUniteInit_OnClick(Sender: tObject);
    procedure MnPrixDeRevient_OnClick(Sender: TObject);
    procedure MnProperties_OnClick(Sender: TObject);
    procedure wSetPlusUniteStock;
    procedure wGetRetour;
    { Applique le profil ARTICLE }
    procedure WProfilArticleState;
    procedure GA_ProfilArticleOnChange(Sender: TObject);
    procedure OnArgumentsGPAO(Arguments: String);
    {$ENDIF}

    {$IFNDEF GPAO}
    function ChampsObligatoires: boolean;
    {$ENDIF}

    {$IFDEF GPAOLIGHT}
    function  BtArtNatCanBeActif: Boolean;
    procedure MnNatureTravail_OnClick(Sender: TObject);
    procedure SetBtArtNatState;
    {$ENDIF}

    {$IFDEF AFFAIRE}
    procedure AFFormuleVariable; //Affaire-ONYX
    procedure AFFormuleVariableClick(sender: Tobject);
    {$ENDIF}
    procedure SelFourn (Sender : Tobject);
    procedure SelCatalog (Sender : Tobject);

  public
    //           SauveParamGrille : string[1];
    ArticleDimFocus: string; // Code article dimensionné recevant le focus
    Depot, Cloture: string; // Appel depuis fiche Stock : dépôt et clôture de l'article
    ListeNatPiece, DatePieceDeb, DatePieceFin: string; // Appel depuis fiche Stock et ventes : Natures, dates début et fin de pièces
    procedure OnChangeDetail;
    procedure OnNavigClick;
    procedure OnClickPrixUnique;
    procedure OnUpdatePrix(stPrix: string);
    procedure OnClickCalcPrix(ga_calcprix, ga_coefcalc, ga_calcauto, ga_pv, t_stockpv: string);
    procedure VerifModifDimension;
    procedure SaisiTarifBase(Etat: string);
    procedure OnArgument(Arguments: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnClose; override; //DC
    procedure OnCancelRecord; override;
    procedure AppliqueProfilArticle(Confirm: boolean; NewRecord: boolean = False);
    {$IFDEF BTP}
    procedure AppliqueProfilBTP(Confirm: boolean);
    procedure PrixPourQteExit (Sender : TObject);
    {$ENDIF}
    procedure EditArtLie(TypArtLie: string);
    procedure Duplication(CleDuplic: string); //PAUL
    procedure ControlNomenc; //AC
    procedure SetTypeArticle(typeArt: string); // PL
    procedure CreateItem(ItemDim: THDimensionItem);
    procedure DeleteItem(ItemDim: THDimensionItem);
    procedure OnChangeItem(Sender: TObject); //Dim
    procedure AjusterTailleTHDim; // repris depuis AdapterFiche de UTofSelectDimDoc
    procedure OnUpdateXXChamp(FieldName: string);
    procedure AffComboEtab(bForceState: boolean = False);
    procedure MAJFicheArticleWithTOB;
    procedure CreationPhotoLiensOLE;
//    function ChampModif(NomChamp: string; PrixUnique: Boolean=True; ArticleDim: Boolean=True): boolean;
  end;

  TOM_Conditionnement = class(TOM)
    procedure OnNewRecord; override;
    procedure OnLoadRecord; override;
  end;

  TOM_CodeCpta = class(TOM)
    procedure OnNewRecord; override;
  end;

  TOM_ProfilArt = class(TOM)
  protected
    procedure OnArgument(Arguments: string); override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnClose; override;
  private
  	{$IFDEF BTP}
    procedure TarifChange (Sender : TObject);
    {$ENDIF}
    {$IFDEF GPAO}
    TobChamp: Tob;
    MontrerCodeChamp: Boolean;
    procedure ActiveOnglets;
    procedure BtAddField_OnClick(Sender: tObject);
    procedure BtDelField_OnClick(Sender: tObject);
    procedure BtMajArticle_OnClick(Sender: tObject);
    procedure CBShowCode_OnCLick(Sender: tObject);
    function GetCBShowCode: TCheckBox;
    function GetListException: tListBox;
    function GetListChamp: tListBox;
    function IsAdvancedFilter: Boolean;
    procedure KillListChampInException(TSChamp, TSException: TStrings);
    procedure ListChamp_DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure ListException_DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
    procedure LoadListChamp;
    procedure LoadListException(TS: TStrings);
    procedure MasqueOnglets;
    procedure PutTobChampInListChamp(TS: TStrings);
    procedure SetLabelInclureExclure;
    procedure SetBtnMajArticleState;
    procedure SaveListException;
    procedure BPARAMPROFILART_OnClick(Sender: tObject);
    procedure GPF_MODECOPIEChange(Sender: TObject);
    procedure GPF_MODEREMPLACEChange(Sender: TObject);
    {$ENDIF}
  end;

procedure AGLLookUpfournisseur(parms: array of variant; nb: integer);
procedure AGLCalculBaseTarifHT(parms: array of variant; nb: integer);
procedure AGLCalculBaseTarifTTC(parms: array of variant; nb: integer);
procedure AFLanceFiche_ArticleGA(lequel, argument: string);
procedure AFLanceFiche_ArticlePouGA(lequel, argument: string);
function  ChampModif(NomChamp: string; PrixUnique: Boolean=True; ArticleDim: Boolean=True): boolean;

const
  // libellés des messages
  TexteMessage: array[1..61] of string = (
    {1} 'Vous devez renseigner un fournisseur' //Catalogue
    {2}, 'Vous devez renseigner la référence' //Catalogue
    {3}, 'La nature de la pièce suivante ne peut être égale à la nature de pièce courante'
    {4}, 'Le prix de vente est inférieur au prix d''achat'
    {5}, 'Le code barre et le qualifiant code barre doivent être renseignés  '
    {6}, 'Le nombre de caractères est incorrect'
    {7}, 'Le code barre est incorrect'
    {8}, 'Le code barre contient des caractères non autorisés'
    {9}, 'Le code barre doit être numérique'
    {10}, 'Le libellé doit être renseigné'
    {11}, 'L''article de remplacement n''existe pas'
    {12}, 'L''article de substitution n''existe pas'
    {13}, 'Vous ne pouvez pas sélectionner l''article courant comme article de remplacement'
    {14}, 'Vous ne pouvez pas sélectionner l''article courant comme article de substitution'
    {15}, 'Erreur de création de l''article dimensionné'
    {16}, 'La première famille de taxe est obligatoire '
    {17}, 'Suppression impossible, cet article est utilisé dans une pièce'
    {18}, 'Le type prestation/frais/fourniture est obligatoire' // Affaire
    {19}, 'Le code prestation/frais/fourniture est obligatoire' // Affaire
    {20}, 'Le code article est obligatoire'
    {21}, 'L''article n''existe pas' //Catalogue
    {22}, 'Impossible de supprimer un fournisseur principal' //Catalogue
    {23}, 'Le fournisseur n''existe pas' //Catalogue
    {24}, 'Votre base tarif HT fait référence au fournisseur principal'
    {25}, 'Fournisseur principal inexistant'
    {26}, 'L''article lié doit être différent de l''article courant'
    {27}, 'Cet article lié existe déjà'
    {28}, 'Le calcul de la base tarif se fait uniquement à partir des bases de l''article'
    {29}, 'L''article doit être tenu en stock, car il est disponible'
    {30}, 'Vous devez renseigner le libellé' //Catalogue
    {31}, 'Suppression impossible, cet article est utilisé sur une ligne d''activité' // Affaire
    {32}, 'Le fournisseur principal n''existe pas'
    {33}, 'Vous n''avez pas calculé la base tarif'
    {34}, 'Le type d''article financier est obligatoire'
    {35}, 'Article en mouvement, impossible de modifier sa gestion par lot'
    {36}, 'Impossible de dé référencer un fournisseur principal' //Catalogue
    {37}, 'Calcul du coefficient impossible, le prix de référence est nul'
    {38}, 'Une déclinaison de même code existe déjà'
    {39}, 'Ce code barres est déjà affecté à un autre article'
    {40}, 'Cet article est déjà mouvementé, vous ne pouvez pas modifier le fournisseur'
    {41}, 'Version de démonstration : le nombre d''articles excède le nombre autorisé.'
    {42}, 'Opération non disponible, il existe du stock sur cet article'
    {43}, 'Opération non disponible, il existe un tarif détail sur cet article'
    {44}, 'Les prix des articles dimensionnés vont être mis à jour !'
    {45}, 'Confirmation'
    {46}, 'Cette TVA n''existe pas'
    {47}, 'Cette TPF n''existe pas'
    {48}, 'Choix d''un type d''ouvrage'
    {49}, 'Choix d''un type de nomenclature'
    {50}, 'Confirmez-vous la suppression du catalogue sur l''ancien fournisseur?'
    {51}, 'Le code article est obligatoire sur le catalogue'
    {52}, 'La saisie du champs suivant est obligatoire : '
    {53}, 'Article indisponible, vous devez auparavant valider votre saisie'
    {54}, 'Cette collection n''existe pas'
    {55}, 'Cette valeur n''existe pas'
    {56}, 'La nature de prestation n''existe pas'
    {57}, 'Plusieurs utilisateurs essayent de créer le même article, veuillez valider de nouveau'
    {58}, 'Renseignez la prestation associée !'
    {59}, 'Ce code existe déjà'
    {60}, 'Les caractères spéciaux dans le code article ont été remplacés par -'
    {61}, 'Ce code barre n''est pas associé à cet article. Voulez-vous l''associer'
    );

  // Libellés des messages GPAO
  TexteMessageGPAO: array[1..6] of string = (
    {1}'Cet article est déjà utilisé dans un autre profil avancé',
    {2}'La limite du nombre maximum de champ est atteinte',
    {3}'Ce profil avancé est utilisé dans des articles',
    {4}'L''article de référence n''existe pas',
    {5}'Aucun champ à inclure - Vérifier le mode de copie ou renseigner les champs à inclure.',
    {6}'La mise à jour des articles utilisant un profil avancé qui préserve les différences '
        + 'est impossible depuis la fiche profil.' + #13 + ' ' + #13 + ' Pour se faire il faut autoriser la mise à jour du profil depuis la fiche article et,'
        + #13 + ' après modification d''un article profil, valider la modification des articles associés.'
    );
  // Const MaxTaxe = 5;

  ErrArticleProfil = 1;   
  ErrNbChampsLimite = 2;
  ErrProfilUtilise = 3;
  ErrArticleProfilExist = 4;
  ErrInclureChamps = 5;
  ErrMajArticles = 6;

var MaxTaxe: integer; // Cette valeur est égale à 1 pour Affaire
  ArticleFourn: boolean; // gestion nouveau fournisseur principal dans lookup
  ArticlePrin, FournPrinc, ArtLib: string; // gestion fournisseur principal
  TobChampModif: TOB; // Champs dimensionnés modifiables

const
  MaxDimChamp = 10;
  {$IFDEF GPAO}
  ErrChampObligatoire = 666;
  {$ENDIF}

implementation

uses
  {$IFDEF GPAO}
    wInitChamp,
  {$ENDIF}
{$IFDEF GPAOLIGHT}
    wArtNat,
{$ENDIF GPAOLIGHT}
{$IFDEF BTP}
    NomenUtil,
{$ENDIF}
	UTofListeInv,
  ParamSoc,UtilSoc,
  BTPUtil,
  StrUtils,
  UTilFonctionCalcul,
  UtilFichiers,
  SaisUtil
  ,UFonctionsCBP, Types;

var ReferArtFourn: string; // Référence article chez le fournisseur


function PlusCalcLine : string;
begin
result := '';
//uniquement en line
//	result := ' AND (CO_CODE <> "FOU")';
end;

  /////////////////////////////////////////////
  // ****** TOM Catalogu ***********************
  /////////////////////////////////////////////

procedure TOM_Catalogu.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

procedure TOM_Catalogu.OnNewRecord;
begin
  inherited;
  //if connexion then exit ;
  THLabel(GetControl('TGA_LIBNATURE')).caption := '';

  SetField('GCA_DATEREFERENCE', Date);
  SetField('GCA_CREERPAR', 'SAI');
  SetField('GCA_DATESUP', iDate2099);

  OldArticle := '';

  if Ecran.Name <> 'GCCATALOGU_SAISI3' then exit;

  if ArticleFourn then
  begin
    // Création d'un nouveau catalogue, à partir du lookup fournisseur de la fiche Article
    if BMonoFournisseur = True then
      SetField('GCA_TIERS', FournPrinc)
    else
      SetField('GCA_REFERENCE', trim(copy(ArticlePrin, 1, 18))); // Référence article chez le fournisseur par défaut
    //
    SetField('GCA_ARTICLE', ArticlePrin);
    SetControlText('_ARTICLE_GEN', trim(copy(ArticlePrin, 1, 18)));
  end;

  if ArtLib <> '' then SetField('GCA_LIBELLE', ArtLib);

  if BMonoFournisseur = True then
    SetFocusControl('GCA_REFERENCE')
  else
    SetFocusControl('GCA_TIERS');

end;

procedure TOM_Catalogu.OnUpdateRecord;
var CodeFourn: string;
    CodeBarre,TypeCodeBarre : string;
begin
  inherited;
  if (ArticleFourn) and (BMonoFournisseur = False) then FournPrinc := GetControlText('GCA_TIERS');
  CodeFourn := GetField('GCA_TIERS');
  if (CodeFourn = '') then
  begin
    SetLastError(1, 'GCA_TIERS');
    exit;
  end;

  if (not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_NATUREAUXI="FOU" and T_TIERS="' + CodeFourn + '"')) then
  begin
    SetLastError(23, 'GCA_TIERS');
    exit;
  end;

  if GetField('GCA_REFERENCE') = '' then
  begin
    SetLastError(2, 'GCA_REFERENCE');
    exit;
  end
  else if (DS.State in [dsInsert]) then
    SetField('GCA_REFERENCE', Trim(GetField('GCA_REFERENCE')));

  if Ecran.Name = 'GCCATALOGU_SAISI3' then
  begin
    if GetControlText('_ARTICLE_GEN') <> '' then
    begin
      // ???? pas compris pourquoi on prenait l'article sans terminaison ==> ça plante !!!!
      //if not LookupValueExist(THEDIT(GetControl('_ARTICLE_GEN'))) then
      if not LookupValueExist(THEDIT(GetControl('GCA_ARTICLE'))) then
      begin
        SetLastError(21, '_ARTICLE_GEN');
        exit;
      end;
    end;

    if (GetControlText('_ARTICLE_GEN') <> Trim(copy(OldArticle, 1, 18))) and (ExisteSQL('SELECT * FROM  ARTICLE WHERE GA_FOURNPRINC ="' +
      GetControlText('GCA_TIERS') + '" AND GA_ARTICLE="' + OldArticle + '"')) then
    begin
      SetLastError(36, '');
      SetControlText('_ARTICLE_GEN', Trim(copy(OldArticle, 1, 18)));
      SetFocusControl('GCA_DATEREFERENCE');
      SetFocusControl('_ARTICLE_GEN');
      exit;
    end;
    if (ArticleFourn) and (GetControlText('_ARTICLE_GEN') = '') then
    begin
      // Si le catalogue est non affecté à un article, celui-ci le sera automatiquement
      // après validation de la fiche Article.
      // Pour celà on doit mémoriser la référence article chez le fournisseur, pour pouvoir
      // accéder ensuite en mise à jour sur le catalogue.
      if GetField('GCA_TIERS') = FournPrinc then ReferArtFourn := GetField('GCA_REFERENCE');
    end;
  end;

  if GetField('GCA_LIBELLE') = '' then
  begin
    SetLastError(30, 'GCA_LIBELLE');
    exit;
  end;

  if ctxAffaire in V_PGI.PGIContexte then
  begin
    if GetField('GCA_ARTICLE') = '' then
    begin
      SetLastError(51, '');
      Exit;
    end;
  end;
  ReferArtFourn := '';
//
  if (Getfield('GCA_CODEBARRE') <> '') or (GetField('GCA_QUALIFCODEBARRE') <> '') then
  begin
    CodeBarre := Getfield('GCA_CODEBARRE');
    TypeCodeBarre := Getfield('GCA_QUALIFCODEBARRE');
    if ((CodeBarre <> '') and (TypeCodeBarre = '')) or
      ((CodeBarre = '') and (TypeCodeBarre <> '')) then
    begin
      SetLastError( 5,''); Exit;
    end
    else
    begin
      case ControlCodeBarre(CodeBarre, TypeCodeBarre) of
        1 : BEGIN SetLastError (6,''); END;
        2 : BEGIN SetLastError (7,''); END;
        3 : BEGIN SetLastError (8,''); END;
        4 : BEGIN SetLastError (9,''); END;
      end;
    end;
  end;
//
end;

function TOM_Catalogu.FindUniteSto (Article : string) : string;
var QQ : TQuery;
begin
  result := 'Non défini';
  QQ := OpenSql ('SELECT GA_QUALIFUNITESTO,GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
  if not QQ.eof then
  begin
  	if QQ.findField('GA_QUALIFUNITESTO').AsString <> '' then
    begin
    	result := QQ.findField('GA_QUALIFUNITESTO').AsString
    end else
    begin
      if QQ.findField('GA_QUALIFUNITEVTE').AsString <> '' then
      begin
    		result := QQ.findField('GA_QUALIFUNITEVTE').AsString
      end;
    end;
  end;
  ferme (QQ);
end;


procedure TOM_Catalogu.AfficheCoefConv;
begin

  THlabel(getControl('TGCA_COEFCONVQTEACH')).Caption := 'Coef. conversion '+
          getCOntroltext('GCA_QUALIFUNITEACH') +'/'+FindUniteSto(GetControltext('GCA_ARTICLE'));
  THlabel(getControl('TGCA_COEFCONVQTEACH')).visible := true;
  THDbEdit(getControl('GCA_COEFCONVQTEACH')).visible := true;

end;

procedure TOM_Catalogu.CacheAndZeroCoefConv;
begin
  THlabel(getControl('TGCA_COEFCONVQTEACH')).visible := false;
  THDbEdit(getControl('GCA_COEFCONVQTEACH')).visible := false;
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  SetField('GCA_COEFCONVQTEACH',1);
end;

procedure TOM_Catalogu.OnChangeField(F: TField);
begin
  inherited;
  if (F.FieldName = 'GCA_TIERS') then
  begin
    if GetField('GCA_TIERS') <> '' then
    begin
      THLabel(GetControl('LIBFOURNISSEUR')).Caption := RechDom('GCTIERSFOURN', GetField('GCA_TIERS'), False);
      if THLabel(GetControl('LIBFOURNISSEUR')).Caption = 'Error' then THLabel(GetControl('LIBFOURNISSEUR')).Caption := '';
    end
    else THLabel(GetControl('LIBFOURNISSEUR')).Caption := '';
  end;


  if (F.FieldName = 'GCA_ARTICLE') then
  begin
    SetControlVisible('LIBARTICLE', Boolean(GetField('GCA_ARTICLE') <> ''));
    SetControlVisible('TTITRE', Boolean(GetField('GCA_ARTICLE') <> ''));
    SetControlVisible('TGCA_ARTICLE', Boolean(GetField('GCA_ARTICLE') <> ''));
    if (getCOntroltext('GCA_QUALIFUNITEACH')<>'') and (getField('GCA_QUALIFUNITEACH')<>'') then
    begin
    	AfficheCoefConv;
    end else
    begin
    	if (getField('GCA_ARTICLE')='') and (getField('GCA_QUALIFUNITEACH')='') then
    		CacheAndZeroCoefConv;
    end;
    if (Ecran is TFFiche) then
    begin
      SetControlVisible('TGCA_ARTICLE', True);
    end;
  end;

  if F.fieldName = 'GCA_QUALIFUNITEACH' then
  begin
    if (getCOntroltext('GCA_QUALIFUNITEACH')<>'') and (getField('GCA_QUALIFUNITEACH')<>'') then
    begin
    	AfficheCoefConv;
    end else
    begin
    	if (getField('GCA_ARTICLE')='') and (getField('GCA_QUALIFUNITEACH')='') then
    		CacheAndZeroCoefConv;
    end;
  end;

  // Modif BTP
  if (F.FieldName = 'GCA_PRIXBASE') then
  begin
    {$IFDEF BTP}
    if (TFFiche(Ecran).fTypeAction <> taModif) then
    begin
      SetField('GCA_DPA', F.Value);
    end;
    {$ENDIF}
  end;
  // ---------

end;


procedure TOM_Catalogu.OnLoadRecord;

begin
  inherited;

  THLabel(GetControl('TGA_LIBNATURE')).caption := '';

  if GetField('GCA_TIERS') <> ''
    then Ecran.Caption := 'Catalogue fournisseur: ' + GetField('GCA_TIERS') + '  ' + RechDom('GCTIERS', GetField('GCA_TIERS'), False);

  {$IFNDEF EAGLCLIENT}
  if connexion then while not DS.EOF do if DS.FindField('GCA_TIERS').AsString = FournPrinc then Break else DS.Next;
  {$ENDIF}

  if BMonoFournisseur = True then
    SetControlEnabled('GCA_TIERS', False)
  else
    SetControlEnabled('GCA_TIERS', (DS.State in [dsInsert]));

  //FV1 : 24/02/2014 - FS#897 - BAGE : en saisie catalogue, permettre la modification de la référence fournisseur
  SetControlEnabled('GCA_REFERENCE', True);

  OldArticle := GetControlText('GCA_ARTICLE');
  connexion := False;
  if Ecran.Name <> 'GCCATALOGU_SAISI3' then exit;
  {$IFDEF BTP}
  if (DS.state = dsInsert) then
  begin
    if FournParam <> '' then SetField('GCA_TIERS', FournParam);
    if UniteParam <> '' then SetField('GCA_QUALIFUNITEACH', UniteParam);
    if prixParam <> 0 then SetField('GCA_PRIXBASE', prixParam);
    if (ArticleParam <> '') then
    begin
    ArticleFourn := true;
    SetField('GCA_ARTICLE', ArticleParam);
    SetControlText('_ARTICLE_GEN', trim(copy(ArticleParam, 1, 18)));
    CodeUnique;
  end;
  end;
  {$ENDIF}
  SetControlText('_ARTICLE_GEN', Trim(Copy(GetControlText('GCA_ARTICLE'), 1, 18)));
  if GetControlText('_ARTICLE_GEN') <> '' then
    THLabel(GetControl('LIBARTICLE')).Caption := LibelleArticleGenerique(Copy(OldArticle, 1, 18))
  else if (ArticleFourn) and (not (DS.State = dsInsert)) then
  begin
    // Affectation de l'article courant à la référence fournisseur du catalogue qui est non encore affecté
    SetControlText('_ARTICLE_GEN', trim(copy(ArticlePrin, 1, 18)));
    CodeUnique;
  end;
  if BMonoFournisseur = True then SetControlEnabled('_ARTICLE_GEN', False);
  if (GetControltext('GCA_ARTICLE') <> '') and (getCOntroltext('GCA_QUALIFUNITEACH')<>'') then
  begin
  	AfficheCoefConv;
  end;

  if DS.State = dsInsert then
  begin
    if not ArticleFourn then
    begin
      SetControlText('_ARTICLE_GEN', '');
      Exit;
    end;
  end
  else if ArticleFourn then SetFocusControl('GCA_DATEREFERENCE');
end;

procedure TOM_Catalogu.OnArgument(st: string);
var Critere, ChampMul, ValMul: string;
  x: integer;
    FF : string ;
    ii : integer ;
begin
  inherited;
  if TToolbarButton97 (GetControl('BTCONDIT')) <> nil then
  begin
    TToolbarButton97 (GetControl('BTCONDIT')).OnClick := BTCONDITClick;
  end;

  if THDBEdit(getControl('GA_ARTICLEPARC')) <> nil then
  begin
  	THDbEDit(GetControl('GA_ARTICLEPARC')).visible := false;
    THLabel(GetControl('TGA_ARTICLEPARC')).visible := false;
  end;

   FF:='#,##0.';
   for ii:=1 to GetParamSoc('SO_DECPRIX') do FF:=FF+'0';
   SetControlProperty('GCA_PRIXBASE','DisplayFormat',FF);
   SetControlProperty('GCA_PRIXVENTE','DisplayFormat',FF);
   SetControlProperty('GCA_DPA','DisplayFormat',FF);

  if (ctxMode in V_PGI.PGIContexte) then BMonoFournisseur := GetParamsoc('SO_MONOFOURNISS')
  else BMonoFournisseur := False;

  if (Ecran is TFFicheListe) then
  begin
    connexion := True; // pour se positionner sur le fournisseur principal
    ArticleAffecte := TFFicheListe(Ecran).FRange;
  end;

  if TFFicheListe(Ecran).FTypeAction = taConsult then SetControlEnabled('B_CATALOGUREF', false);

  // Récupération libellé par défaut passé en paramètre
  ArtLib := '';
  repeat
    Critere := Trim(ReadTokenSt(st));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        if ChampMul = 'LIB' then ArtLib := ValMul;
        {$IFDEF BTP}
        if ChampMul = 'ARTICLE' then ArticleParam := ValMul;
        if ChampMul = 'TIERS' then FournParam := ValMul;
        if ChampMul = 'UNITE' then UniteParam := ValMul;
        if ChampMul = 'PRIX' then prixParam := valeur(ValMul);
        {$ENDIF}
      end;
    end;
  until Critere = '';

  // Par affaire pas de cotation ni accès au catalogue fournisseur
  if ctxAffaire in V_PGI.PGIContexte then
  begin
    SetControlVisible('GCA_COTE', False);
    SetControlVisible('TGCA_COTE', False);
    SetControlVisible('B_CATALOGUREF', False);
  end;

  {$IFDEF EAGL}
  THVAlCOMBOBOX(GetConTrol('GCA_QUALIFUNITEACH')).Datatype := 'GCQUALUNITTOUS';
  THVAlCOMBOBOX(GetConTrol('GCA_QUALIFUNITEACH')).reload;
  {$ELSE}
  THDBVAlCOMBOBOX(GetConTrol('GCA_QUALIFUNITEACH')).Datatype := 'GCQUALUNITTOUS';
  THDBVAlCOMBOBOX(GetConTrol('GCA_QUALIFUNITEACH')).reload;
  {$ENDIF}

  _ARTICLE_GEN := THEdit(GetControl('_ARTICLE_GEN'));
  _ARTICLE_GEN.OnElipsisClick := ARTGEN_OnElipsisClick;
  _ARTICLE_GEN.OnExit := ARTGEN_OnExit;

{$IFNDEF V10}
  SetControlVisible('BTCONDIT',false);             
{$ENDIF}

//uniquement en line
{*
  SetControlVisible('GCA_CODEBARRE',false);
  SetControlVisible('TGCA_CODEBARRE',false);
  SetControlVisible('GCA_QUALIFCODEBARRE',false);
*}
end;

//FV1 : 13/02/2014 - FS#818 - Fiche catalogue : la recherche sur article ne doit faire apparaître que des articles
Procedure TOM_Catalogu.ARTGEN_OnElipsisClick(Sender: TObject);
Var CodeArt : String;
    StChamps: String;
begin

  {$IFDEF BTP}
  CodeArt := GetControlText('_ARTICLE_GEN');
  StChamps:= '';

  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';

  StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP"';

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps + ';ACTION=CONSULTATION');

  if codeArt <> '' then SetControlText('_ARTICLE_GEN', CodeArt)
  {$ENDIF BTP}

end;

Procedure TOM_Catalogu.ARTGEN_OnExit(Sender: TObject);
begin

  CodeUnique()

end;

procedure TOM_Catalogu.OnDeleteRecord;
var WhereSt: string;
begin
  inherited;
  if (BMonoFournisseur = False) and
    (ExisteSQL('SELECT GA_ARTICLE FROM  ARTICLE WHERE GA_FOURNPRINC ="' + GetControlText('GCA_TIERS') + '" AND GA_ARTICLE="' + GetControlText('GCA_ARTICLE') +
      '"')) then
  begin
    SetLastError(22, '');
    exit;
  end
  else
  begin
    if Ecran.Name = 'GCCATALOGU_SAISI3' then SetControlText('_ARTICLE_GEN', '');
    THLabel(GetControl('LIBFOURNISSEUR')).Caption := '';
    THLabel(GetControl('LIBARTICLE')).Caption := '';
  end;
  WhereSt := ' GL_REFCATALOGUE="' + GetField('GCA_REFERENCE') + '" and GL_ENCONTREMARQUE="X" and GL_VIVANTE="X" ';
  if ExisteSQL('SELECT GL_REFCATALOGUE FROM LIGNE WHERE ' + WhereSt) then
  begin
    SetLastError(17, '');
    exit;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 31/01/2000
Modifié le ... : 31/01/2000
Description .. : Affectation d'un article à une référence article fournisseur
Mots clefs ... : FOURNISSEUR; ARTICLE; REFERENCE
*****************************************************************}

procedure TOM_Catalogu.AffectationReference;
var MaCle, CodeFournisseur, Reference: string;
begin
  MaCle := AGLLanceFiche('GC', 'GCCATALOGU_AFFECT', '', '', '');
  Reference := ReadTokenst(MaCle);
  CodeFournisseur := ReadTokenSt(MaCle);
  if Reference = '' then exit;
  begin
    //if HShowMessage('0;Confirmation;Confirmez vous l''affectation de cette référence?;Q;YN;N;N;','','')<>mrYes then exit;
    ExecuteSQL('UPDATE CATALOGU SET GCA_ARTICLE="' + ArticleAffecte + '" WHERE GCA_REFERENCE="' + Reference + '" AND GCA_TIERS="' + CodeFournisseur + '"');
    DS.DisableControls;
    DS.Refresh;
    TFFicheliste(Ecran).BFirstClick(nil);
    while not DS.EOF do if DS.FindField('GCA_REFERENCE').AsString = Reference then Break else DS.Next;
  end;
  DS.EnableControls;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Agnès CATHELINEAU
Créé le ...... : 04/02/2000
Modifié le ... : 04/02/2000
Description .. : Gestion du code article dans le catalogue
Mots clefs ... : CODE;ARTICLE;GENERIQUE;UNIQUE;
*****************************************************************}

procedure TOM_Catalogu.CodeUnique;
var CodeUnique    : String;
    CodeGenerique : String;
    CodeDim       : String;
    CodeArticle   : String;
    LibArticle    : String;
begin

  LibArticle := '';
  //
  CodeArticle := GetControlText('_ARTICLE_GEN');
  CodeDim := '';
  CodeUnique := CodeArticleUnique2(CodeArticle, '');

  if CodeArticleUnique2(CodeArticle, '') <> OldArticle then
    if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit

  if CodeArticle = Trim(Copy(GetField('GCA_ARTICLE'), 1, 18)) then exit;

  if CodeArticle = '' then
  begin
    SetField('GCA_ARTICLE', '');
    THLabel(GetControl('LIBARTICLE')).Caption := '';
    exit;
  end;

  if not BMonoFournisseur then
  begin
    // Sélection éventuelle d'un article dimensionné
    if ExisteSQL('Select GA_ARTICLE from ARTICLE where GA_CODEARTICLE="' + CodeArticle + '" AND GA_STATUTART="GEN"')
      then CodeDim := SelectUneDimension(CodeUnique);
  end;

  if CodeDim <> '' then
    SetField('GCA_ARTICLE', CodeDim)
  else
    SetField('GCA_ARTICLE', CodeUnique);

  LibArticle := LibelleArticleGenerique(Copy(CodeUnique, 1, 18));

  if LibArticle = 'Error' then SetControlText('LIBARTICLE', '');

  if CodeArticle <> '' then
  begin
    THLabel(GetControl('LIBARTICLE')).Visible := True;
    THLabel(GetControl('TGCA_ARTICLE')).Visible:= True;
    //
    THLabel(GetControl('LIBARTICLE')).Caption := LibArticle;
    UpdateCaption(ecran);
  end;


end;

/////////////////////////////////////////////
//********* TOM ArticlePiece ****************
/////////////////////////////////////////////
{$IFNDEF CCS3}

procedure TOM_ArticlePiece.OnChangeField(F: TField);
var NaturePiece: string;
begin
  // on initialise le libellé avec le libellé de la nature de pièce
  if (F.FieldName = 'GAP_NATUREPIECEG') then
  begin
    if (GetField('GAP_LIBELLE') = '') then
    begin
      if GetField('GAP_NATUREPIECEG') = '' then exit;
      NaturePiece := RechDom('GCNATUREPIECEG', GetField('GAP_NATUREPIECEG'), FALSE);
      SetField('GAP_LIBELLE', NaturePiece);
    end;
  end;
  
end;

procedure TOM_ArticlePiece.OnNewRecord;
begin
  SetField('GAP_DATEUTILISAT', Date);
end;
{$ENDIF}

/////////////////////////////////////////////
// ****** TOM ArticleLie ********************
/////////////////////////////////////////////

procedure TOM_ArticleLie.OnLoadRecord;
var CodeUnique: string;
begin
  SetControlText('LIBARTICLE', LibelleArticleGenerique(GetField('GAL_ARTICLE')));
  SetControlText('LIBARTICLELIE', LibelleArticleGenerique(GetField('GAL_ARTICLELIE')));
  CodeUnique := CodeArticleUnique2(GetControlText('GAL_ARTICLELIE'), ''); // ajout OT
  SetControlText('_CODEARTICLEUNIQUE', CodeUnique); // ajout OT
  
{$IFDEF GIGI}
  // PL le 16/10/03 : en GI/GA on ne veut pas les choix sur le tarif et la quantité pour l'instant
  if (GetControl('GAL_TARIFREF') <> nil) then
    SetControlProperty ('GAL_TARIFREF', 'visible', False);
  if (GetControl('GAL_QTEREF') <> nil) then
    SetControlProperty ('GAL_QTEREF', 'visible', False);
{$ENDIF}
end;

procedure TOM_ArticleLie.OnChangeField(F: TField);
var CodeUnique: string;
begin
  // on initialise le libellé avec le libellé de la nature de pièce
  if (F.FieldName = 'GAL_ARTICLELIE') and (GetField('GAL_ARTICLELIE') <> '') then
  begin
    SetControlEnabled('BZOOM', True);
    SetControlText('LIBARTICLELIE', LibelleArticleGenerique(GetField('GAL_ARTICLELIE')));
    if (GetField('GAL_LIBELLE') = '') then SetField('GAL_LIBELLE', getControlText('LIBARTICLELIE'));
    // Modif BTP
    CodeUnique := CodeArticleUnique2(GetControlText('GAL_ARTICLELIE'), '');
    SetControlText('_CODEARTICLEUNIQUE', CodeUnique);
  end;
  if GetField('GAL_ARTICLELIE') = '' then SetControlEnabled('BZOOM', False);
end;

procedure TOM_ArticleLie.OnNewRecord;
var QQ: TQuery;
  Article: string;
begin
  SetField('GAL_TYPELIENART', 'LIE');
  SetField('GAL_OBLIGATOIRE', '-');
  SetField('GAL_TARIFPREF', '-');
  SetField('GAL_RANG', '10');
  Article := GetField('GAL_ARTICLE');
  QQ := OpenSQL('SELECT MAX(GAL_RANG) AS MAXRANG FROM ARTICLELIE WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="' + Article + '"', TRUE);
  if not QQ.EOF then
  begin
    SetField('GAL_RANG', (QQ.Findfield('MAXRANG').AsInteger) + 10);
  end;
  Ferme(QQ);

end;

procedure TOM_ArticleLie.OnUpdateRecord;
begin
  if (GetControlText('GAL_ARTICLELIE') <> '') then
  begin
    if GetControlText('GAL_ARTICLELIE') = GetField('GAL_ARTICLE') then
    begin
      SetFocusControl('GAL_ARTICLELIE');
      LastError := 26;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;

    if (DS.State in [dsInsert]) and ExisteSQL('SELECT GAL_ARTICLELIE FROM ARTICLELIE WHERE GAL_ARTICLE="' + GetField('GAL_ARTICLE') + '" AND GAL_ARTICLELIE="'
      + GetControlText('GAL_ARTICLELIE') + '"') then
    begin
      SetFocusControl('GAL_ARTICLELIE');
      LastError := 27;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
  end;
end;

/////////////////////////////////////////////
// ****** TOM Article ***********************
/////////////////////////////////////////////

procedure TOM_Article.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end;

procedure TOM_Article.MAJFicheArticleWithTOB;
var CC: TImage;
  EmplacementPhoto: string;
begin
  TOBInfoArt.PutEcran(Ecran);
  SetControlVisible('P_PHOTO', false);
  CC := TImage(GetControl('LAPHOTO'));
  if CC = nil then exit;
  EmplacementPhoto := TOBInfoArt.GetValue('PathPhoto') + TOBInfoArt.GetValue('NamePhoto');
  if not FileExists(EmplacementPhoto) then exit;
  CC.Picture.LoadFromFile(EmplacementPhoto);
  SetControlVisible('P_PHOTO', True);
  SetActiveTabSheet('P_PHOTO');
end;

//
// Modif BTP
//
{$IFDEF BTP}
procedure TOM_Article.EntreeDetParcNomen(Sender: Tobject);
var Article, Libelle, domaine: string;
begin
  Article := Getfield('GA_ARTICLE');
  Libelle := Getfield('GA_LIBELLE');
  EntreeParcDetail([Article, Libelle, TFFiche(Ecran).FTypeAction], 3);
end;

procedure TOM_Article.EntreeDetOuv(Sender: Tobject);
var Article, Libelle, domaine: string;
begin
  Article := Getfield('GA_ARTICLE');
  Libelle := Getfield('GA_LIBELLE');
  Domaine := Getfield('GA_DOMAINE');
  EntreeOuvrageDetail([Article, Libelle, Domaine, TFFiche(Ecran).FTypeAction], 3);
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 19/09/2003
Modifié le ... : 19/09/2003
Description .. : NomChamp est-il modifiable pour un article dimensionné ?
Suite ........ : PrixUnique : utilisé si le champ est un prix
Suite ........ : ArticleDim : question posée pour les articles
Suite ........ : dimensionnés ou non
Mots clefs ... : ARTICLE;MODIF;DIM
*****************************************************************}
//function TOM_Article.ChampModif(NomChamp: string; PrixUnique: Boolean=True; ArticleDim: Boolean=True): boolean;
function ChampModif(NomChamp: string; PrixUnique: Boolean=True; ArticleDim: Boolean=True): boolean;
begin
  // DCA - FQ MODE 10813
  if not ArticleDim then result := True
  else if ( (NomChamp = 'GA_CODEBARRE') or (NomChamp = 'GA_CODEDOUANIER') or
{$IFDEF BTP}
            (NomChamp = 'GA_LIBELLE') or
{$ENDIF}
            (NomChamp = 'GA_FAMILLETAXE1') or (NomChamp = 'GA_FAMILLETAXE2') or
            (NomChamp = 'GA_POIDSBRUT') or (NomChamp = 'GA_POIDSNET') or
            (NomChamp = 'GA_POIDSDOUA') or (NomChamp = 'GA2_COLLECTIONBAS') or
            ((NomChamp = 'GA_COLLECTION') and (GetParamSocSecur('SO_GCDIMCOLLECTION', False))) )
  then result := True
  else result := PrixModifiable(NomChamp, PrixUnique, False);
end;

procedure TOM_Article.ErgoGCS3;
{$IFDEF CCS3}
var TM: TMenuItem;
  {$ENDIF}
begin
  {$IFDEF CCS3}
  TM := TMenuItem(Ecran.FindComponent('mnReferencement'));
  if TM <> nil then TM.Visible := False;
  TM := TMenuItem(Ecran.FindComponent('mnTradLibelle'));
  if TM <> nil then TM.Visible := False;
  TM := TMenuItem(Ecran.FindComponent('mnArtPiece'));
  if TM <> nil then TM.Visible := False;
  if (copy(Ecran.Name, 1, 9) = 'AFARTICLE') then
  begin
    if (getcontrol('GB_VALEURLIBRE')  <> Nil) then SetControlVisible('GB_VALEURLIBRE', False);
    if (getcontrol('GB_DATELIBRE')    <> Nil) then SetControlVisible('GB_DATELIBRE', False);
    if (getcontrol('GB_BOOLEANLIBRE') <> Nil) then SetControlVisible('GB_BOOLEANLIBRE', False);
    if (getcontrol('GB_TEXTELIBRE')   <> Nil) then SetControlVisible('GB_TEXTELIBRE', False);
  end;
  {$ENDIF}
end;

procedure TOM_Article.OnArgument(Arguments: string);
var Critere				: String;
		NomChamp			: String;
    StArg					: string;
    ChampMul			: String;
    ValMul				: string;
    PrixUnique		: boolean;
	  x							: Integer;
    i							: Integer;
    iCol					: Integer;
    pos1					: integer;
  {$IFDEF EAGLCLIENT}
  	Chp_Fournisseur	: THEdit;
  {$ELSE}
  	Chp_Fournisseur	: THDBEdit;
  {$ENDIF}
  	QQ						: TQuery;
  	GB1						: TGroupBox;
    GB2						: TGroupBox;
  	Button 				: THButton;
begin
  inherited;

  AppliqueFontDefaut (THRichEditOle(GetControl('GA_BLOCNOTE')));

  IsLine := False;
  //
  TabTarif  := TTabSheet (getControl('P_TARIF'));
  NBDecPrix := GetParamSoc('SO_DECPRIX');
  NBDecQte  := GetParamSoc('SO_DECQTE');
  //
  StArg := Arguments;
  BMonoFournisseur := (ctxMode in V_PGI.PGIContexte) and GetParamsoc('SO_MONOFOURNISS');
  //
  Consultation      := False;
  ArticleFourn      := False; // fournisseur principal
  AfficheTarif      := TRUE; //par defaut
  ArtToDuplic       := '';
  DuplicFromMenu    := False;
  DuplicFromFiche   := False;
  MaxTaxe           := 5;
  FicheDimensionne  := False; //DC
  PrixUnique        := True;
  NavigClick        := False;
  FicheStock        := False;
  FicheStoVte       := False;
  ArticleDimFocus   := '';
  Depot             := '';
  Cloture           := '-';
  ListeNatPiece     := '';
  DatePieceDeb      := '';
  DatePieceFin      := '';
  TypeMasque        := VH_GC.BOTypeMasque_Defaut; // Si pas de type masque précisé en param, utilisation du type masque défaut.
  CreatArticlePhoto := False;
  //

  // DBR : ARTICLES_LIES que si paramsoc activé - fiche 10804
  if not VH_GC.GCArticlesLies then SetControlProperty ('mnArtLie', 'Visible', False);
  //
  DecQte := '';
  for i := 1 to NBDecQte do DecQte := DecQte + '0';
  //
  DecPrix := '';
  for i := 1 to NBDecPrix do DecPrix := DecPrix + '0';

  SetControlProperty('GA_POIDSNET', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_POIDSBRUT', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_VOLUME', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_SURFACE', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_LINEAIRE', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_POIDSDOUA', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QTEDEFAUT', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QECOACH', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QPCBACH', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QECOVEN', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QECOPROD', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_QPCBPROD', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_PRIXPOURQTE', 'DisplayFormat', '#,##0.' + DecQte);
  SetControlProperty('GA_PRIXPOURQTEAC', 'DisplayFormat', '#,##0.' + DecQte);
  //FV1 : 08/02/2018 - FS#2929 - SES_ETANCHEITE SERVICE : En création article, ne prend pas 3 décimales pour le PU en UA
  SetControlProperty('GA_PAUA','DisplayFormat', '#,##0.' + DecPrix);
  //
  SetControlProperty('GA_DPA', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_DPR', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PVHT', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PVTTC', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PAHT', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PRHT', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PMAP', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PMRP', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_PRIXSIMULACH', 'DisplayFormat', '#,##0.' + DecPrix);
  if Assigned(GetControl('GA_QUALIFUNITESTO')) then SetControlEnabled ('GA_QUALIFUNITESTO',false);
  //SetControlEnabled('TGA_QUALIFUNITESTO',false);

  if assigned(GetControl('GA_COEFCONVQTEVTE')) then
  begin
    SetControlEnabled('GA_COEFCONVQTEVTE',false);
    THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := true;
  end;

  TEPrixPourQte := THEdit(GetControl('GA_PRIXPOURQTE'));
  if TEPrixPourQte <> Nil then TEPrixPourQte.OnExit := PrixPourQteExit;
  // Gestion des decimale de prix et quantité
  if (copy (ecran.name,1,13)<> 'BTARTPOURCENT') and (copy(ecran.name,1,9)<>'BTARTPARC') then
  begin
    SetControlProperty('CMARGE','Decimals',4);
    THNumEdit(GetControl('CMARGE')).NumericType := ntDecimal;
    SetControlProperty('MMARGE','Decimals',NBDecPrix);
    THNumEdit(GetControl('MMARGE')).NumericType := ntDecimal;
  End;

  if copy (ecran.name,1,9)= 'BTARTICLE' then
  begin
    SetControlProperty('TPAMO','Decimals',NBDecPrix);
    THNumEdit(GetControl('TPAMO')).NumericType := ntDecimal;
    SetControlProperty('TPRMO','Decimals',NBDecPrix);
    THNumEdit(GetControl('TPRMO')).NumericType := ntDecimal;
    SetControlProperty('TPVHTMO','Decimals',NBDecPrix);
    THNumEdit(GetControl('TPVHTMO')).NumericType := ntDecimal;
    SetControlProperty('TPVTTCMO','Decimals',NBDecPrix);
    THNumEdit(GetControl('TPVTTCMO')).NumericType := ntDecimal;
    //
    SetControlProperty('GA_PAHT1','DisplayFormat', '#,##0.' + DecPrix);
    SetControlProperty('GA_DPR1','DisplayFormat', '#,##0.' + DecPrix);
    SetControlProperty('GA_PVHT1','DisplayFormat', '#,##0.' + DecPrix);
    SetControlProperty('GA_PVTTC1','DisplayFormat', '#,##0.' + DecPrix);
    //
    //FV1 : 08/02/2018 - FS#2929 - SES_ETANCHEITE SERVICE : En création article, ne prend pas 3 décimales pour le PU en UA
    SetControlProperty('GA_DPR_',     'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_DPR_')).NumericType := ntDecimal;
    SetControlProperty('GA_PVHT_',    'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_PVHT_')).NumericType := ntDecimal;
    SetControlProperty('GA_PVTTC_',   'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_PVTTC_')).NumericType := ntDecimal;
    SetControlProperty('TOTALPA_',    'Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPA_')).NumericType := ntDecimal;
    SetControlProperty('TOTALPR_',    'Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPR_')).NumericType := ntDecimal;
    SetControlProperty('TOTALPVHT_',  'Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPVHT_')).NumericType := ntDecimal;
    SetControlProperty('TOTALPVTTC_', 'Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPVTTC_')).NumericType := ntDecimal;
    //
    SetControlProperty('GA_DPR_1',    'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_DPR_1')).NumericType := ntDecimal;
    SetControlProperty('GA_PVHT_1',   'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_PVHT_1')).NumericType := ntDecimal;
    SetControlProperty('GA_PVTTC_1',  'Decimals',NBDecPrix);
    THNumEdit(GetControl('GA_PVTTC_1')).NumericType := ntDecimal;
    SetControlProperty('TOTALPA_1',   'Decimals',NBDecPrix);
    //
    SetControlProperty('TOTALPA','Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPA')).NumericType := ntDecimal;
    SetControlProperty('TOTALPR','Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPR')).NumericType := ntDecimal;
    SetControlProperty('TOTALPVHT','Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPVHT')).NumericType := ntDecimal;
    SetControlProperty('TOTALPVTTC','Decimals',NBDecPrix);
    THNumEdit(GetControl('TOTALPVTTC')).NumericType := ntDecimal;
    SetControlProperty('DQTEPREST','Decimals',NBDecQte);
    THNumEdit(GetControl('DQTEPREST')).NumericType := ntDecimal;
    //FV1 : 08/02/2018 - FS#2929 - SES_ETANCHEITE SERVICE : En création article, ne prend pas 3 décimales pour le PU en UA
    SetControlProperty('GCA_PRIXBASE','DisplayFormat', '#,##0.' + DecPrix);    //
    SetControlProperty('GF_PRIXUNITAIRE','DisplayFormat', '#,##0.' + DecPrix);    //
    //
    SetControlProperty('GCA_PRIXBASE','Decimals',NBDecPrix);
    //THNumEdit(GetControl('GCA_PRIXBASE')).NumericType := ntDecimal;
    SetControlProperty('GF_PRIXUNITAIRE','Decimals',NBDecPrix);
    //THNumEdit(GetControl('GF_PRIXUNITAIRE')).NumericType := ntDecimal;
  end;

  SetControlProperty('GA_PCB', 'DisplayFormat', '#,##0');

	{$IFNDEF BTP}
  SetControlProperty('GA_COEFCALCHT', 'DisplayFormat', '#,##0.' + DecPrix);
  SetControlProperty('GA_COEFCALCTTC', 'DisplayFormat', '#,##0.' + DecPrix);
	{$ENDIF}

  {MODE DCA - Passage en mode consultation dans le FO dans la fiche article dimensionné }
  if (ctxFO in V_PGI.PgiContexte) and (Ecran.Name = 'GCARTICLE') then
  	 begin
     if Arguments <> '' then Arguments := Arguments + ';' ;
     Arguments := Arguments + 'ACTION=CONSULTATION' ;
  	 end;

  if (copy(ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(ecran.name, 1, 13) = 'AFARTPOURCENT') then
  	 begin
     SelTypeArticle := THMultiValComboBox(ecran.FindComponent('SELTYPEART'));
     critere := ' AND (CO_CODE NOT LIKE "F%") AND (CO_CODE NOT LIKE "PRI%") AND (CO_CODE <> "EPO") AND (CO_CODE <> "POU")';
     Critere := Critere + ' AND (CO_CODE <> "NOM") AND (CO_CODE <> "CTR")';
     SelTypeArticle.Plus := Critere;
  	 end;

  if (copy(ecran.name, 1, 9) = 'BTARTPARC') then
  	 begin
     SelTypeArticle := THMultiValComboBox(ecran.FindComponent('SELTYPEART'));
     critere := ' AND (CO_CODE LIKE "PA%")';
     SelTypeArticle.Plus := Critere;
  	 end;

  if THValComboBox(GetControl('GA_TARIFARTICLE')) <> nil then
     THValComboBox(GetControl('GA_TARIFARTICLE')).OnChange := CodeTarifArticleChange;

  {$IFDEF GPAOLIGHT}
  if Assigned(Ecran) and (Ecran.Name = 'WARTICLE_FIC') then
  	 begin
     if Assigned(TToolbarButton97(GetControl('BT_ARTNAT'))) then TToolbarButton97(GetControl('BT_ARTNAT')).OnClick := MnNatureTravail_OnClick;
     if Assigned(TMenuItem(GetControl('MnNatureTravail'))) then TMenuItem(GetControl('MnNatureTravail')).OnClick := MnNatureTravail_OnClick;
     end;
  {$ENDIF}

  OK_MetreDoc		:= GetParamSocSecur('SO_BTMETREDOC', false);

  //MODIF PAR LOT
  i := pos('MODIFLOT', Arguments);
  ModifparLot := i <> 0;

  if ModifParLot then
	   begin
     TFfiche(Ecran).MonoFiche := true;
     StSQL := copy(Arguments, i + 9, length(Arguments));
     end
  else
    {récup des arguments }
     begin
     repeat
     Critere := uppercase(Trim(ReadTokenSt(Arguments)));
     if Critere <> '' then
        begin
        x := pos('=', Critere);
        if x <> 0 then
           begin
           ChampMul := copy(Critere, 1, x - 1);
           ValMul := copy(Critere, x + 1, length(Critere));
           if ChampMul = 'FICHDIM' then
              begin
              FicheDimensionne := True;
              PrixUnique := Boolean(ValMul = 'X');
              end
           Else
	            ControleCritere(ChampMul, Valmul);
	         end;
   	    end;
     until Critere = '';
     {$IFDEF GPAO}
	   OnArgumentsGPAO(Arguments);
  	 {$ENDIF}
		 //
	   {$IFDEF CCS3}
     OnArgumentsCCS3(Arguments);
   	 {$ENDIF}
  	 //
     {$IFDEF BTP}
     OnArgumentsBTP(Arguments);
     {$ENDIF}
     //
     //uniquement en line
//   OnArgumentsLine(Arguments);
     GestionFicheNormale;
     end;

  // Affichage Domaine
  if (GetControl('TMPDOMAINE') <> nil) and // DCA - FQ MODE 10814
     (copy(Ecran.Name, 1, 9) = 'GCARTICLE') or (copy(Ecran.Name, 1, 9) = 'AFARTICLE') then
  	 begin
     if THValComboBox(GetControl('TMPDOMAINE')).Values.Count > 1 then
    		begin
        SetControlVisible('GA_DOMAINE', True);
        SetControlVisible('TGA_DOMAINE', True);
        SetControlProperty('GA_DOMAINE', 'TabStop', True);
        end;
     end;

  // DCA - FQ MODE 10814
  if (Ecran.Name = 'GCARTICLE') then
  	 begin
     SpecifGescom := CreateTOF(Ecran.Name, Ecran, False, False);
     if SpecifGescom <> nil then SpecifGescom.Argument(StArg);
  	 end;

  // Paramétrage des libellés des tables libres
  GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GA_LIBREART', 10, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_CHARLIBRE', 3, '');
  GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '');

  // Chargement des champs dimensionnés modifiables en TOB
  if not FicheDimensionne then
		 begin
     TobChampModif := TOB.Create('Champs modifiables', nil, -1);
     if TobChampModif <> nil then
        begin
        QQ := OpenSQL('select CO_ABREGE from COMMUN where CO_TYPE="GDC" and CO_LIBRE like "%ART%"' + ' and not CO_ABREGE like "%LIBRE%"', True);
        TobChampModif.LoadDetailDB('COMMUN', '', '', QQ, False);
        Ferme(QQ);
        end;
     // Gestion du OnChangeOnglet
     Classeur := TPageControl(GetControl('PAGES'));
     if Classeur <> nil then Classeur.OnChange := OnChangeOnglet;
     end
  else
     begin
     if (ctxMode in V_PGI.PGIContexte) then
        begin
        TobChampModif := TheTOB; // Tob passé en paramètre entre fiche article GEN et fiche article DIM
        if TobChampModif <> nil then
           begin
           FicheReadOnly(ECRAN); // Désactive la fiche
           // FicheReadOnly désactive les champs -> réactivation.
           if GetControl('TGA_DATECREATION') <> nil then SetControlProperty('TGA_DATECREATION', 'Enabled', True);
           if GetControl('GA_DATECREATION') <> nil then SetControlProperty('GA_DATECREATION', 'Enabled', True);
           if GetControl('TGA_CREATEUR') <> nil then SetControlProperty('TGA_CREATEUR', 'Enabled', True);
           if GetControl('GA_CREATEUR') <> nil then SetControlProperty('GA_CREATEUR', 'Enabled', True);
           if GetControl('TGA_DATEMODIF') <> nil then SetControlProperty('TGA_DATEMODIF', 'Enabled', True);
           if GetControl('GA_DATEMODIF') <> nil then SetControlProperty('GA_DATEMODIF', 'Enabled', True);
           if GetControl('TGA_UTILISATEUR') <> nil then SetControlProperty('TGA_UTILISATEUR', 'Enabled', True);
           if GetControl('GA_UTILISATEUR') <> nil then SetControlProperty('GA_UTILISATEUR', 'Enabled', True);
           // FicheReadOnly désactive les boutons -> réactivation.
           if GetControl('BInsert') <> nil then SetControlProperty('BInsert', 'Enabled', True);
           if GetControl('BAnnuler') <> nil then SetControlProperty('BAnnuler', 'Enabled', True);
           if GetControl('BDelete') <> nil then SetControlProperty('BDelete', 'Enabled', True);
           if GetControl('BValider') <> nil then SetControlProperty('BValider', 'Enabled', True);
           // Active les seuls champs modifiables d'un article dimensionné
           if (TobChampModif <> nil) then
              for i := 0 to TobChampModif.Detail.Count - 1 do
                  begin
                  NomChamp := TobChampModif.Detail[i].GetValue('CO_ABREGE');
                  SetControlProperty(NomChamp, 'Enabled', ChampModif(NomChamp, PrixUnique, True));        // DCA - FQ MODE 10813
                  end;
           end;
        end;
        // Gestion du OnChangeOnglet
        Classeur := TPageControl(GetControl('PAGES'));
        if Classeur <> nil then Classeur.OnChange := OnChangeOnglet;
     end;

  // Paramétrage des libellés pour la MODE
  if (ctxMode in V_PGI.PGIContexte) then
		 begin
     SetControlCaption('TGA_PVHT',        TraduireMemoire('Prix Négoce (HT)'));
     SetControlCaption('TGA_PVTTC',       TraduireMemoire('Prix Déta&il (TTC)'));
     SetControlCaption('TGA_CALCPRIXHT',  TraduireMemoire('Prix Négoce'));
     SetControlCaption('TGA_CALCPRIXTTC', TraduireMemoire('Prix Dé&tail'));
     SetControlCaption('TGA_COEFCALCHT',  TraduireMemoire('Coef. de calcul'));
     SetControlCaption('TGA_COEFCALCTTC', TraduireMemoire('Coef. de ca&lcul'));
     SetControlCaption('GA_CALCAUTOHT',   TraduireMemoire('&Recalcul automatique'));
     SetControlCaption('GA_CALCAUTOTTC',  TraduireMemoire('Recalc&ul automatique'));
     SetControlCaption('TGA_PCB',         TraduireMemoire('&Géré par multiple de'));
     SetControlVisible('XXTYPEMASQUE',    True);
     SetControlVisible('LXXTYPEMASQUE',   True);
     // Paramétrage libellés Coef + recalcul automatique PA/PR
     CalcPAPR := GetParamsoc('SO_CALCPAPR');
     if CalcPAPR = 'AR' then
        begin
        SetControlCaption('TGA_COEFFG',   TraduireMemoire('C&oefficient PA->PR'));
        SetControlCaption('GA_DPRAUTO',   TraduireMemoire('&Recalcul automatique du PR'));
        end
     else if CalcPAPR = 'RA' then
        begin
        SetControlCaption('TGA_COEFFG',   TraduireMemoire('C&oefficient PR->PA'));
        SetControlCaption('GA_DPRAUTO',   TraduireMemoire('&Recalcul automatique du PA'));
        // Modif présentation : Group Box PR puis PA
        GB1 := TGroupBox(GetControl('GB_PRIX_ACHAT'));
        GB2 := TGroupBox(GetControl('GB_PRIX_REVIENT'));
        if (GB1 <> nil) and (GB2 <> nil) then
           begin
           pos1 := GB1.Top;
           GB1.Top := GB2.Top;
           GB1.TabOrder := 2;
           GB2.Top := pos1;
           GB2.TabOrder := 0;
           end;
        end;
     {$IFDEF MODES3}
     SetControlVisible('BPROFIL', False);
     {$ENDIF}
     // Mise en forme des zones familles et stat. article
     for iCol := 1 to 3 do
         Begin
         ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), ecran);
         end;
     ChangeLibre2('TGA_COLLECTION', ecran);
     if GetPresentation = ART_ORLI then
     		begin
        for iCol := 4 to 8 do ChangeLibre2('TGA2_FAMILLENIV' + InttoStr(iCol), ecran);
        for iCol := 1 to 2 do ChangeLibre2('TGA2_STATART' + InttoStr(iCol), ecran);
    	  end;
  	 end
  else
  	 begin
     SetControlVisible('XXTYPEMASQUE', False);
     SetControlVisible('LXXTYPEMASQUE', False);
  	 end;

  SetActiveTabSheet('PGeneral');

  // Définition d'un évênement sur le OnExit du champ FOURNPRINC
  {$IFDEF EAGLCLIENT}
  Chp_Fournisseur := THEdit(GetControl('GA_FOURNPRINC'));
  if Chp_Fournisseur <> nil then Chp_Fournisseur.OnExit := OnExit_Fournisseur;
  {$ELSE}
  Chp_Fournisseur := THDBEdit(GetControl('GA_FOURNPRINC'));
  if Chp_Fournisseur <> nil then Chp_Fournisseur.OnExit := OnExit_Fournisseur;
  {$ENDIF}

  ErgoGCS3;

  if (ecran <> nil) then TFFiche(Ecran).OnKeyDown := FormKeyDown;

  {JLD Confidentialité concepts}
  if TFFiche(Ecran).fTypeAction <> taConsult then
  	 begin
     if not ExJaiLeDroitConcept(TConcept(gcArtCreat), False) then
    		begin
        SetControlVisible('BINSERT', False);
        SetControlVisible('BINSERTNOMENC', False);
        SetControlVisible('B_DUPLICATION', False);
    		end;
  	 end;

  // TODO
  if TypeArticle = 'MAR' then
  begin
    SetControlVIsible('GA_PERTEPROP',true);
    SetControlVIsible('TGA_PERTEPROP',true);
  end;
  if TypeArticle = 'OUV' then
  	 begin
  	 THDbValComboBox(GetControl('GA_FAMILLENIV1')).DataType := 'BTFAMILLEOUV1';
   	 THDbValComboBox(GetControl('GA_FAMILLENIV2')).DataType := 'BTFAMILLEOUV2';
  	 THDbValComboBox(GetControl('GA_FAMILLENIV3')).DataType := 'BTFAMILLEOUV3';
     THButton(GetControl('BPARAMOU2')).visible := true;
     THButton(GetControl('BPARAMOU3')).visible := true;
     THButton(GetControl('BPARAMNV2')).visible := false;
     THButton(GetControl('BPARAMNV3')).visible := false;
     TDBCHeckBox(GetControl('GA_PRIXPASMODIF')).visible := false;
     SetControlVisible('GA_REMISELIGNE',false);
  end else if copy(TypeArticle,1,2) = 'PA' then  // articles parcs
  	 begin
  	 THDbValComboBox(GetControl('GA_FAMILLENIV1')).DataType := 'BTFAMILLEPARC1';
   	 THDbValComboBox(GetControl('GA_FAMILLENIV2')).DataType := 'BTFAMILLEPARC2';
  	 THDbValComboBox(GetControl('GA_FAMILLENIV3')).DataType := 'BTFAMILLEPARC3';
     THButton(GetControl('BPARAMPA2')).visible := true;
     THButton(GetControl('BPARAMPA3')).visible := true;
     THButton(GetControl('BPARAMNV2')).visible := false;
     THButton(GetControl('BPARAMNV3')).visible := false;
     SetControlVisible('GA_REMISELIGNE',false);
  end else
  	 begin
		 if (GetControl('BPARAMOU2')) <> nil then THButton(GetControl('BPARAMOU2')).visible := false;
		 if (GetControl('BPARAMOU3')) <> nil then THButton(GetControl('BPARAMOU3')).visible := false;
 		 if (GetControl('BPARAMNV2')) <> nil then THButton(GetControl('BPARAMNV2')).visible := true;
		 if (GetControl('BPARAMNV3')) <> nil then THButton(GetControl('BPARAMNV3')).visible := true;
  	 end;
  //
  TFFiche(ecran).OnAfterFormShow := AfterFormShow;
  //
  //FV1 : 27/10/20014-FS#1280 - SOCMA : la gestion automatique des codes articles et prestations ne fonctionne plus
  if GetParamSoc('SO_GCNUMARTAUTO') then
    ThEdit(GetControl('GA_CODEARTICLE')).Enabled := False
  Else
    ThEdit(GetControl('GA_CODEARTICLE')).Enabled := true;

end;

Procedure TOM_Article.GestionFicheNormale;
Var	DimMasque			: THDBValComboBox;
	  LibTypArticle	: String;
    ArtSpecif			: string;
    i							: Integer;
    X							: Integer;
  {$IFDEF EAGLCLIENT}
	  ControlCommiss: TCheckBox;
  {$ELSE}
	  ControlCommiss: TDBCheckBox;
  {$ENDIF}
	  pop						: TPopupMenu;
    CodeAxe : string;
    NumAxe : integer;
Begin

  if (copy(ecran.name, 1, 9) = 'BTARTICLE') then // ecran des articles prix posés
     Begin
	   if (TypeArticle = 'MAR') or (TypeArticle = 'ARP') or (TypeArticle = 'PRE') then
        begin
          MnTariffou := TmenuItem(Getcontrol('mnTarifFou'));
          MnTariffou.OnClick := MnTariffouClick;
          //
          MnFourn := TmenuItem(Getcontrol('mnFourn'));
          Mnfourn.OnClick := OnElipsisClick_Fournisseur;
          if TypeArticle <> 'PRE' then THCheckbox(GetControl('GA_RELIQUATMT')).Visible := False;
//Non en line
          if (Pos(typeArticle,'MAR;ARP') <> 0) and (JaiLeDroitTag(323210)) then
          begin
            THlabel(GetControl('TGA_ARTICLEPARC')).visible  := true;
            THDBEdit(GetControl('GA_ARTICLEPARC')).visible  := true;
            THDBEdit(GetControl('GA_ARTICLEPARC')).OnElipsisClick := ARTPARCCLick;
          end;
        end;
	   end
  Else if (copy(ecran.name, 1, 12) = 'BTPRESTATION') then
     Begin
     if (TypeArticle = 'PRE') then
     begin
       THCheckbox(GetControl('GA_RELIQUATMT')).Visible := True;
     end;
     if (TypeArticle = 'FRA') then
     begin
      THCheckbox(GetControl('GA_RELIQUATMT')).Visible := False;
      SetControlProperty('BAppelVariable', 'Visible', False);
     end;
     CodeAxe := GetparamSocSecur('SO_BTAXEANALSTOCK','TX1');
      if VH^.LiaisonY2ViaShare then
      begin
        THDbEdit(GetControl('GA_SECTION')).DataType := 'TZCSECTION'+Copy(CodeAxe,2,1);
      end else
      begin
        THDbEdit(GetControl('GA_SECTION')).DataType := 'TZSECTION'+Copy(CodeAxe,2,1);
      end;
     end
  Else if (copy(ecran.name, 1, 13) = 'BTARTPOURCENT') then
     Begin
     end;

  {JLD Confidentialité concepts}
  if ((TFFiche(Ecran).fTypeAction = taModif) and (not ExJaiLeDroitConcept(TConcept(gcArtModif), False))) then
     begin
     TFFiche(Ecran).fTypeAction := taConsult;
   	 SetControlVisible('B_DUPLICATION', False); //mcd  le bouton n'est pas cacher en auto comme le Binsert...
   	 SetControlVisible('BINSERTNOMENC', False); // DBR Fiche 10191
   	 SetControlEnabled('BPARAMNV2',False);
     SetControlEnabled('BPARAMNV3',False);
     end;

  { ne prendre en compte que les masques de dimension non fermés lors de la création d'un nouvel article }
  if (TFFiche(Ecran).fTypeAction = taCreat) then
     begin
     DimMasque := THDBValComboBox(ecran.FindComponent('GA_DIMMASQUE'));
     if DimMasque <> nil then DimMasque.Plus := 'AND GDM_FERMER <> "X" ';
     end;

  {Affichage de la gestion documentaire isoflex}
  GereIsoflex;

  ArtSpecif := IsArticleSpecif(TypeArticle);

  // SPECIF AFFAIRE ////////////////////////
  if ArtSpecif = 'FicheAffaire' then
     begin
     MaxTaxe := 1;
     if (TypeArticle = 'PRE') then SetControlVisible('GA_ACTIVITEEFFECT', True);
     // La famille comptable est visible
     if (GetParamSocSecur('SO_GCVENTCPTAART', False) = False) then
        begin
        SetControlVisible('GA_COMPTAARTICLE', False);
        SetControlVisible('TGA_COMPTAARTICLE', False);
        end;
     // SPECIF SCOT ////////////////////////
     if (ctxScot in V_PGI.PGIContexte) then
        begin
        if not (GetParamSOc('SO_AFGESTIONTARIF')) then
           begin
           SetControlVisible('GA_REMISELIGNE', False);
           SetControlVisible('GA_TARIFARTICLE', False);
           SetControlVisible('TGA_TARIFARTICLE', False);
           end;
        SetControlVisible('GA_PRIXPOURQTE', False);
        SetControlVisible('TGA_PRIXPOURQTE', False);
        SetControlVisible('GA_PCB', False);
        SetControlVisible('TGA_PCB', False);
        end;
{$IFDEF AFFAIRE}
     SelTypeArticle := THMultiValComboBox(ecran.FindComponent('SELTYPEART'));
     if SelTypeArticle <> nil then
        begin
        SelTypeArticle.Plus := PlusTYpeArticle;
        i := Pos(' OR CO_CODE="POU"', SelTypeArticle.plus);
        SelTypeArticle.plus := Copy(SelTypeArticle.plus, 1, i - 1) + Copy(SelTypeArticle.plus, i + 17, Strlen(Pchar(SelTypeArticle.plus)));
        end;
{$ENDIF}
     end;

  if (ArtSpecif = 'FicheModeArt') or (ArtSpecif = 'FicheModePre') then
     MaxTaxe := 2;

  if (ArtSpecif = 'FicheAffaire') or (ArtSpecif = 'FicheBtp') then   // Gestion du libellé du type article en titre du form et sur les libellés de certains champs
     begin
     if TypeArticle <> '' then
        begin
        LibTypArticle := RechDom('GCTYPEARTICLE', TypeArticle, FALSE);
        {$IFDEF EAGLCLIENT}
        ControlCommiss := TCheckBox(GetControl('GA_COMMISSIONNABLE'));
        if (ControlCommiss <> nil) then TCheckBox(ControlCommiss).Caption := LibTypArticle + ' c&ommissionnable';
        {$ELSE}
        ControlCommiss := TDBCheckBox(GetControl('GA_COMMISSIONNABLE'));
        if (ControlCommiss <> nil) then TDBCheckBox(ControlCommiss).Caption := LibTypArticle + ' c&ommissionnable';
        {$ENDIF}
        //
        {$IFDEF BTP}
        if (LibTypArticle = 'Marchandise') then LibTypArticle := TraduireMemoire('Article');
        if (LibTypArticle = 'Nomenclature') then
           begin
           LibTypArticle := TraduireMemoire('Ouvrage');
           SetControlProperty('B_DUPLICATION', 'Hint', TraduireMemoire('Dupliquer l''ouvrage'));
           end;
        {$ELSE}
        if (LibTypArticle = 'Marchandise') then LibTypArticle := 'Fourniture';
        {$ENDIF}
        Ecran.Caption := LibTypArticle + ' :';
        updatecaption(Ecran);
        end;
     end;
//
  if (ctxChr in V_PGI.PGIContexte) then // CHR
     begin
     if TypeArticle <> 'FI' then
        begin
        SetControlVisible('GA_TYPEARTFINAN', False);
        SetControlVisible('TGA_TYPEARTFINAN', False);
        end;
     if TypeArticle <> 'NOM' then
        begin
        SetControlVisible('BNOMENC', False);
        end
     else
        begin
        SetControlVisible('GB_COMPTA', False);
        SetControlVisible('GA_REMISELIGNE', False);
        end;
     end;
//
  pop := TPopupMenu(GetControl('POPMENU'));
  if (pop <> nil) and // DCA - FQ MODE 10814
     ((copy(Ecran.Name, 1, 9) = 'GCARTICLE') or
      (copy(Ecran.Name, 1, 7) = 'ARTICLE') or
      (copy(Ecran.Name, 1, 9) = 'AFARTICLE') or
      (copy(Ecran.name, 1, 9) = 'BTARTICLE') or
      (copy(Ecran.name, 1, 12) = 'BTPRESTATION')) then // Modif BTP
      for x := 0 to pop.items.count - 1 do
  	      begin
	        if pop.items[x].name = 'mnTarif' then pop.items[x].visible := AfficheTarif;
		      end;

//MM a prévoir les menus de l'écran fiche article GC
  if (Ecran.Name = 'WARTICLE_FIC') or ( GetParamSoc('SO_PREFSYSTTARIF')) then
     begin
     for x := 0 to pop.items.count - 1 do
         begin
         if      (pop.items[x].name = 'mnTarif')      then pop.items[x].visible := True
         else if (pop.items[x].name = 'mnTarifAchat') then pop.items[x].visible := AfficheTarif
         else if (pop.items[x].name = 'mnTarifVente') then pop.items[x].visible := AfficheTarif
         else if (pop.items[x].name = 'mnTarifHT'   ) then pop.items[x].visible := False
         else if (pop.items[x].name = 'mnTarifTTC'  ) then pop.items[x].visible := False
         end;
     end;

//--- Modif BTP
  if ((TypeArticle = 'OUV') or (TypeArticle = 'POU')) then
     SetControlVisible('BPOPCOMPLEMENT', False)
  else if (copy(Ecran.Name, 1, 9) = 'GCARTICLE') then
     begin
     pop := TPopupMenu(GetControl('POPCOMPLEMENT')); // DCA - FQ MODE 10814
     if (pop <> nil) then for x := 0 to pop.items.count - 1 do
     		begin
        if (TFFiche(Ecran).ftypeaction = taConsult) and (pop.items[x].name = 'mnReferencement') then
           pop.items[x].enabled := false;
        if (uppercase(pop.items[x].name) = 'MNMULTIBLOB') then
           if (not GetParamSoc('SO_GCCDUSEIT')) then pop.items[x].enabled := false;
        end;
     end;

//---
  {$IFDEF EAGLCLIENT}
  if TFFiche(Ecran).FTypeAction = taCreat then eAglInsertionOn; // DBR Fiche 10093
  {$ELSE}
  DS.AfterInsert := InsertionOn;
  DS.AfterPost := InsertionOff;
  DS.AfterCancel := OnDefaire; // pour rinitialiser les champs non DB aprés un Cancel
  {$ENDIF}

end;


Procedure TOM_Article.ControleCritere(ChampMul, ValMul : String);
Begin

	if ChampMul = 'TARIF' then
     AfficheTarif := (ValMul = 'O')
  else if ChampMul = 'DUPLICATION' then // Duplication
     begin
     ArtToDuplic := ValMul;
     DuplicFromMenu := True;
     SetControlEnabled('B_DUPLICATION', FALSE);
     end
  else if ChampMul = 'ACTION' then // Action
     begin
     if (ValMul = 'CONSULTATION') then
        begin
        Consultation := True;
        SetControlEnabled('B_DUPLICATION', False);
        end;
     end
  else if ChampMul = 'TYPENOMENC' then TypeNomenc := ValMul // LS
  else if ChampMul = 'TYPEARTICLE' then TypeArticle := ValMul // AC:traitment mar ou nom
  else if ChampMul = 'TOP' then Top := StrToInt(ValMul)
  else if ChampMul = 'LEFT' then Left := StrToInt(ValMul)
  else if ChampMul = 'STOCK' then FicheStock := Boolean(ValMul = 'X')
  else if ChampMul = 'STOCKVTE' then FicheStoVte := Boolean(ValMul = 'X')
  else if ChampMul = 'ARTDIM' then ArticleDimFocus := ValMul
  else if ChampMul = 'DEPOT' then Depot := ValMul
  else if ChampMul = 'CLOTURE' then Cloture := ValMul
  else if ChampMul = 'TYPEMASQ' then TypeMasque := ValMul
  else if ChampMul = 'NATPIECE' then ListeNatPiece := ValMul
  else if ChampMul = 'DATEPIECEDEB' then DatePieceDeb := ValMul
  else if ChampMul = 'DATEPIECEFIN' then DatePieceFin := ValMul
  else if ChampMul = 'FOURNDOCACH' then FournDocAch := ValMul //AC renseigner FournPrinc si nouveau ds doc Ach et monofournisseur
  else if ChampMul = 'ARTICLEPHOTO' then CreatArticlePhoto := ValMul = 'X';

end;

{$IFDEF GPAO}
Procedure TOM_Article.OnArgumentsGPAO(Arguments : String);
Begin

  if Assigned(Ecran) and (Ecran.Name = 'WARTICLE_FIC') then
  	 begin
     { Récupération du type Article / Article profil}
     OnlyProfil := GetArgumentValue(Arguments, 'ONLYPROFIL') = wTrue;
     if OnlyProfil then
        begin
        TForm(Ecran).Caption := TraduireMemoire('Article Profil') + ' : ';
        UpdateCaption(Ecran);
        end;
     { Récupération des évènements }
     if GetParamSoc('SO_WCUCONVERSION') then
        begin
        if Assigned(GetControl('MnUniteConversion')) then TMenuItem(GetControl('MnUniteConversion')).OnClick := MnUniteConversion_OnClick
        end
     else
        begin { Les conversions d'unités avancés ne sont pas activées }
        SetControlEnabled('MnUniteConversion', False);
        SetControlEnabled('GA_CODEFORME', False);
        end;
     if Assigned(TMenuItem(GetControl('MnUniteInit'))) then TMenuItem(GetControl('MnUniteInit')).OnClick := MnUniteInit_OnClick;
     if Assigned(TMenuItem(GetControl('MnProperties'))) then TMenuItem(GetControl('MnProperties')).OnClick := MnProperties_OnClick;
     if Assigned(TMenuItem(GetControl('MnPrixDeRevient'))) then TMenuItem(GetControl('MnPrixDeRevient')).OnClick := MnPrixDeRevient_OnClick;
     { Tob des champs non duplicables dans les profils avancés }
     TobUnAutorizedFields := nil;
     { Active les onglets }
     SetActiveTabSheet('PGENERAL');
     { Libellé des familles }{ TP GPAO 2}
     for iCol := 1 to 3 do ChangeLibre2('TGA_FAMILLENIV' + InttoStr(iCol), ecran);
     { Gestion du domaine d'activité }
     if THValComboBox(GetControl('TMPDOMAINE')).Values.Count > 1 then
        begin
        SetControlVisible('GA_DOMAINE', True);
        SetControlVisible('TGA_DOMAINE', True);
        SetControlProperty('GA_DOMAINE', 'TabStop', True);
        end;
     if Assigned(GetControl('GA_PROFILARTICLE')) then ThDbValComboBox(GetControl('GA_PROFILARTICLE')).OnChange := GA_ProfilArticleOnChange;
     { Controles des boutons en fonction des paramètres d'ouverture de la fiche }
     SetControlProperty('BINSERTNOMENC', 'VISIBLE', not OnlyProfil);
     end;

end;
{$ENDIF}

Procedure TOM_Article.OnArgumentsCCS3(Arguments : String);
Var x	: Integer;
Begin

{$IFDEF PREMIUM}
  if not VH_GC.GCAchatStockSeria then
	  begin
    SetControlEnabled('GA_TENUESTOCK',False);
    SetControlEnabled('GA_CONTREMARQUE',False);
    SetControlVisible('GA_FOURNPRINC',False);
    SetControlVisible('TGA_FOURNPRINC',False);
    SetControlEnabled('GA_DPA',False);
    SetControlEnabled('TGA_DPA',False);
    SetControlEnabled('GA_PMAP',False);
    SetControlEnabled('TGA_PMAP',False);
    SetControlEnabled('GA_DPR',False);
    SetControlEnabled('TGA_DPR',False);
    SetControlEnabled('GA_PMRP',False);
    SetControlEnabled('TGA_PMRP',False);
    if assigned(GetControl('GA_QUALIFUNITESTO')) then SetControlEnabled('GA_QUALIFUNITESTO',False);
    //SetControlEnabled('TGA_QUALIFUNITESTO',False);
    if assigned(GetControl('GA_COEFCONVQTEVTE')) then
    begin
    	SetControlEnabled('GA_COEFCONVQTEVTE',false);
    	THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := true;
    end;
    SetControlVisible('GA_TYPENOMENC', False);
  	end;
//
  pop := TPopupMenu(GetControl('POPMENU'));
  if (pop <> nil) and // DCA - FQ MODE 10814
     ((copy(Ecran.Name, 1, 9) = 'GCARTICLE') or
      (copy(Ecran.Name, 1, 7) = 'ARTICLE') or
      (copy(Ecran.Name, 1, 9) = 'AFARTICLE') or
      (copy(Ecran.name, 1, 9) = 'BTARTICLE') or
      (copy(Ecran.name, 1, 12) = 'BTPRESTATION')) then // Modif BTP
      for x := 0 to pop.items.count - 1 do
          if pop.items[x].name = 'mnDispo' then
             pop.items[x].visible := VH_GC.GCAchatStockSeria;

  if (copy(Ecran.Name, 1, 9) = 'GCARTICLE') then
     begin
     pop := TPopupMenu(GetControl('POPCOMPLEMENT')); // DCA - FQ MODE 10814
     if (pop <> nil) then for x := 0 to pop.items.count - 1 do
     		begin
        if pop.items[x].name = 'mnFourn' then  pop.items[x].visible := VH_GC.GCAchatStockSeria;
        end;
     end;
{$ENDIF}

end;

Procedure TOM_Article.OnArgumentsBTP(Arguments : String);
Begin

  if (Not IsMasterOnShare('GCFAMILLENIV1')) then
  begin
  	SetControlCaption('GB_FAMILLE','Familles à renseigner dans la base principale');
  	SetControlEnabled('GB_FAMILLE', False);
  end else
  begin
    if TypeArticle='OUV' then
    begin
      setcontrolProperty('TGA_FAMILLENIV1', 'Caption', RechDom('BTLIBOUVRAGE', 'BO1', false));
      setcontrolProperty('TGA_FAMILLENIV2', 'Caption', RechDom('BTLIBOUVRAGE', 'BO2', false));
      setcontrolProperty('TGA_FAMILLENIV3', 'Caption', RechDom('BTLIBOUVRAGE', 'BO3', false));
    end else if (copy(TypeArticle,1,2)='PA') then
    begin
      setcontrolProperty('TGA_FAMILLENIV1', 'Caption', RechDom('BTLIBARTPARC', 'PA1', false));
      setcontrolProperty('TGA_FAMILLENIV2', 'Caption', RechDom('BTLIBARTPARC', 'PA2', false));
      setcontrolProperty('TGA_FAMILLENIV3', 'Caption', RechDom('BTLIBARTPARC', 'PA3', false));
    end else
    begin
      setcontrolProperty('TGA_FAMILLENIV1', 'Caption', RechDom('GCLIBFAMILLE', 'LF1', false));
      setcontrolProperty('TGA_FAMILLENIV2', 'Caption', RechDom('GCLIBFAMILLE', 'LF2', false));
      setcontrolProperty('TGA_FAMILLENIV3', 'Caption', RechDom('GCLIBFAMILLE', 'LF3', false));
    end;
    if GetParamSoc('SO_GCFAMHIERARCHIQUE') = True then
  	begin
      SetControlProperty('GA_FAMILLENIV2', 'DataTypeParametrable', False);
      SetControlVisible('BPARAMNV2', True);
      SetControlProperty('GA_FAMILLENIV3', 'DataTypeParametrable', False);
      SetControlVisible('BPARAMNV3', True);
  	end
    else
  	begin
      SetControlProperty('GA_FAMILLENIV2', 'DataTypeParametrable', true);
      SetControlVisible('BPARAMNV2', false);
      SetControlProperty('GA_FAMILLENIV3', 'DataTypeParametrable', true);
      SetControlVisible('BPARAMNV3', False);
  	end;
  end;

  if (copy(ecran.name, 1, 9) = 'BTARTICLE') then
  Begin
    if (TypeArticle <> 'NOM') and (TypeArticle <> 'OUV') then TGroupBox(Getcontrol('G_MARGE')).Visible := true;
    // gestion des nomenclatures : déclaration bouton
    BNomenclature := TToolbarButton97(ecran.FindComponent('BNOMENC'));
    BNomenclature.onclick := OuvreNomenclature;
    // gestion des Ouvrages : déclaration bouton
    BOuvrage := TToolbarButton97(ecran.FindComponent('BNOMENC1'));
    BOuvrage.onclick := EntreeDetOuv;
    //0
    BCodeBarre := TToolbarButton97(ecran.FindComponent('BCODEBARRE'));
    BCodeBarre.onclick := BCODEBARREOnClick;
    //
    //BCodeBarre.Visible := False;
    //
    MaxTaxe := 1;
    //
    TTabSheet(GetControl('TTARIFACH')).TabVisible := false;
    //Création des TOBs
    TOBTarif := TOB.Create('TARIF', nil, -1);
    TOBCatalog := TOB.Create('CATALOGU', nil, -1);
    TOBDetail := TOB.Create('LES LIGNES', nil, -1);
    TOBEDetail := TOB.Create('NOMENENT', nil, -1);
    TOBPrestation := TOB.Create('ARTICLE', nil, -1);
    TOBTiers := TOB.Create('TIERS', nil, -1);
    //
    TPrestation := THEdit(Getcontrol('TPRESTATION'));
    DQtePrest := THNumEdit(GetControl('DQTEPREST'));
    //
    if TPrestation <> nil then
       begin
       TPrestation.OnElipsisClick := TPrestationElipsClick;
       TPrestation.OnDblClick := TPrestationElipsClick;
       TPrestation.OnExit := TPrestationChange;
       end;
    //
    if DQtePrest <> nil then DQtePrest.OnExit := ChangeQtePrest;
    TmenuItem(GetControl('MnChFourn')).onclick := SelFourn;
    TmenuItem(GetControl('MnChCatalogu')).onclick := SelCatalog;
  end
  Else if (copy(ecran.name, 1, 12) = 'BTPRESTATION') then
  begin
    TGroupBox(Getcontrol('G_MARGE')).Visible := true;
    TTAbSheet(GetControl('P_STATS')).TabVisible := False;
  End
  Else if (copy(ecran.name, 1, 9) = 'BTARTPARC') then
  begin
    if TFFiche(Ecran).fTypeAction = taCreat then
    begin
    	TToolbarButton97(ecran.FindComponent('BDUPLICART')).OnClick := DuplicFromArt;
    end;
    if TypeArticle='PA2' then
    begin
      BNomenclature := TToolbarButton97(ecran.FindComponent('BNOMENC1'));
      BNomenclature.onclick := EntreeDetParcNomen;
    end else
    begin
    	TToolbarButton97(ecran.FindComponent('BNOMENC1')).visible := false;
    end;
  End
  Else if (copy(ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(ecran.name, 1, 13) = 'AFARTPOURCENT') then
	begin
  end;

  if TYPEARTICLE = 'MAR' then
     prefixe := trim(GetParamsoc('SO_GCPREFIXEART'))
  else if (TYPEARTICLE = 'ARP') then
     Begin
     prefixe := trim(GetParamsoc('SO_BTPREFIXEPOS'));
     TTabSheet(GetControl('TTARIFACH')).TabVisible := true;
     End
  else if (TYPEARTICLE = 'PRE') or (TYPEARTICLE = 'CTR') then
     prefixe := trim(GetParamsoc('SO_GCPREFIXEPRE'))
  else if (TYPEARTICLE = 'NOM') or (TYPEARTICLE = 'OUV') then
     prefixe := trim(GetParamsoc('SO_GCPREFIXENOM'))
  else
     prefixe := '';

  if (TypeArticle = 'MAR') or (TypeArticle='PRE') or (TypeArticle = 'ARP') then
  	 begin
  	 CodeSousTarif := THValComboBox(GetCOntrol('GA2_SOUSFAMTARART'));
     if CodeSousTarif <> nil then
				CodeSousTarif.OnChange := SousFamTarArtChange;
  	 end;

  //FV1 : 09/04/2015 - FS#1468 - SOCMA : Problème après MAJ tarif artciles
  BACTUPRIX   := TToolBarButton97(Getcontrol('BACTUPRIX'));
  if BACTUPRIX <> nil then BACTUPRIX.OnClick   := BACTUPRIX_OnClick;
  CALCAUTOHT  := TCheckBox(Getcontrol('GA_CALCAUTOHT'));
  if CALCAUTOHT <> nil then  CALCAUTOHT.OnClick  := CALCAUTOHT_OnClick;
  CALCAUTOTTC := TCheckBox(Getcontrol('GA_CALCAUTOTTC'));
  if CALCAUTOTTC <> nil then CALCAUTOTTC.OnClick := CALCAUTOTTC_OnClick;
  DPRAUTO     := TCheckBox(Getcontrol('GA_DPRAUTO'));
  if DPRAUTO <> nil then DPRAUTO.OnClick     := DPRAUTO_OnClick;

  if TToolbarButton97(ecran.FindComponent('BAPPELVARIABLE')) <> nil then
  begin
    // gestion des Variables : déclaration bouton
    BAppelVariable := TToolbarButton97(ecran.FindComponent('BAPPELVARIABLE'));
    BappelVariable.Visible := True;
    BAppelVariable.onclick := OuvreVariable;
  end;

  if TToolbarButton97(ecran.FindComponent('BTEXCEL')) <> nil then
  begin
    BExcel         := TToolbarButton97(GetControl('BTEXCEL'));
    BExcel.visible := false;
    BExcel.onclick := ExcelOnClick;
  end;


End;

Procedure TOM_Article.OnArgumentsLine(Arguments : String);
var Button : THButton;
Begin

	IsLine := True;

  if THLabel(GetCOntrol('TGA_NATUREPRES')) 	<> nil then  THLabel(GetCOntrol('TGA_NATUREPRES')).visible := False;
  if THEdit(GetCOntrol('TGA_NATUREPRES')) <> nil then THDbEDIt(GetCOntrol('GA_NATUREPRES')).visible := False;

  if THDbEdit(Getcontrol('GA_CODEARTICLE')) <> nil then
     THDbEdit(Getcontrol('GA_CODEARTICLE')).MaxLength := GetParamSoc('SO_GCLGNUMART');

  if (copy(ecran.name, 1, 12) = 'BTARTICLE_S1')then
  	 begin
     // Gestion des métrés
     BExcel := TToolbarButton97(GetControl('BTEXCEL'));
     // Modified by f.vautrain 14/11/2017 14:36:54 - MICHEL SA : Violation d'accès sur ouverture article pourcentage
     If Assigned(BExcel) then BExcel.Visible := false;
     // gestion des nomenclatures : déclaration bouton
     BNomenclature := TToolbarButton97(ecran.FindComponent('BNOMENC'));
     BNomenclature.onclick := OuvreNomenclature;
     // gestion des Ouvrages : déclaration bouton
     BOuvrage := TToolbarButton97(ecran.FindComponent('BNOMENC1'));
     BOuvrage.onclick := EntreeDetOuv;
     //
	   TPrestation := THEdit(Getcontrol('TPRESTATION'));
     DQtePrest := THNumEdit(GetControl('DQTEPREST'));
     //
     if TPrestation <> nil then
        begin
        TPrestation.OnElipsisClick := TPrestationElipsClick;
        TPrestation.OnDblClick := TPrestationElipsClick;
        TPrestation.OnExit := TPrestationChange;
        end;
     if DQtePrest <> nil then DQtePrest.OnExit := ChangeQtePrest;
  	 end;

  //Gestion des métrés
  //if (TypeArticle = 'MAR') or (TypeArticle = 'ARP')  or (TypeArticle = 'PRE') then
  //   OpenOleContainer;

  if TypeArticle = 'PRE' then
	   begin
	   THLabel(GetCOntrol('TGA_NATUREPRES')).visible := true;
		 THDbEDIt(GetCOntrol('GA_NATUREPRES')).visible := true;
		 THLabel(GetCOntrol('TGA_LIBNATURE')).visible := true;
		 TToolBarButton97(GetCOntrol('BPROFIL')).visible := false;
     end;

  //if (GetControl('GA_NATUREPRES') <> nil) then
 	//   THLabel(GetControl('GA_NATUREPRES')).visible := False;

  {JLD Confidentialité concepts}
  if TFFiche(Ecran).fTypeAction <> taConsult then
  begin
    if not ExJaiLeDroitConcept(TConcept(gcArtCreat), False) then
    begin
      SetControlVisible('BINSERT', False);
      SetControlVisible('BINSERTNOMENC', False);
      SetControlVisible('B_DUPLICATION', False);
    end;
  end;
  // TODO
  if copy(TypeArticle,1,2) = 'PA' then
  begin
  	THDbValComboBox(GetControl('GA_FAMILLENIV1')).DataType := 'BTFAMILLEPARC1';
   	THDbValComboBox(GetControl('GA_FAMILLENIV2')).DataType := 'BTFAMILLEPARC2';
  	THDbValComboBox(GetControl('GA_FAMILLENIV3')).DataType := 'BTFAMILLEPARC3';
    THButton(GetControl('BPARAMPA2')).visible := true;
    THButton(GetControl('BPARAMPA3')).visible := true;
    THButton(GetControl('BPARAMNV2')).visible := false;
    THButton(GetControl('BPARAMNV3')).visible := false;
    TDBCHeckBox(GetControl('GA_PRIXPASMODIF')).visible := false;
  end else
  if TypeArticle = 'OUV' then
  begin
  	THDbValComboBox(GetControl('GA_FAMILLENIV1')).DataType := 'BTFAMILLEOUV1';
   	THDbValComboBox(GetControl('GA_FAMILLENIV2')).DataType := 'BTFAMILLEOUV2';
  	THDbValComboBox(GetControl('GA_FAMILLENIV3')).DataType := 'BTFAMILLEOUV3';
    THButton(GetControl('BPARAMOU2')).visible := true;
    THButton(GetControl('BPARAMOU3')).visible := true;
    THButton(GetControl('BPARAMNV2')).visible := false;
    THButton(GetControl('BPARAMNV3')).visible := false;
    TDBCHeckBox(GetControl('GA_PRIXPASMODIF')).visible := false;
  end else
  begin
		Button := THButton(GetControl('BPARAMOU2'));
    if button <> nil then
    begin
		THButton(GetControl('BPARAMOU2')).visible := false;
    THButton(GetControl('BPARAMOU3')).visible := false;
    THButton(GetControl('BPARAMNV2')).visible := true;
    THButton(GetControl('BPARAMNV3')).visible := true;
    end;
  end;
  //
  TFFiche(ecran).OnAfterFormShow := AfterFormShow;

End;

{$IFDEF EAGLCLIENT} // pour l'instant c'est les meme que pas eAgl mais
// c'est pour différencier si nécessaire : DBR Fiche 10093

procedure TOM_Article.eAglInsertionOn; //désactiver les boutons lors d'une création d'article
begin

  SetControlEnabled('BINSERTNOMENC', False);
  SetControlEnabled('BNOMENC', False);

  {$IFDEF BTP}
  SetControlEnabled('BAPPELVARIABLE', False);
  SetControlEnabled('BNOMENC1', False);
  {$ENDIF}

  SetControlEnabled('BPOPMENU', False);
  SetControlEnabled('BPOPCOMPLEMENT', False);
  SetControlEnabled('BMEMO', False);
  SetControlEnabled('B_DUPLICATION', False);
  SetControlEnabled('BCODEBARRE', False);

end;
// eAglInsertionOff n'est pas nécessaire car en validation apres création
// on sort systématiquement de la fiche article

{$ELSE}

procedure TOM_Article.InsertionOn(DataSet: TDataSet); //désactiver les boutons lors d'une création d'article
begin
  SetControlEnabled('BINSERTNOMENC', False);
  SetControlEnabled('BNOMENC', False);

  {$IFDEF BTP}
  SetControlEnabled('BAPPELVARIABLE', False);
  SetControlEnabled('BNOMENC1', False);
  {$ENDIF}

  SetControlEnabled('BPOPMENU', False);
  SetControlEnabled('BPOPCOMPLEMENT', False);
  SetControlEnabled('BMEMO', False);
  SetControlEnabled('B_DUPLICATION', False);
  SetControlEnabled('BCODEBARRE', False);

end;

procedure TOM_Article.InsertionOff(DataSet: TDataSet); //réactiver les boutons après validation d'un nouvel article
var ArtError, Sql: string;
    QQ: TQuery;
    ii : integer;
begin
  // Nouvelle gestion : création des dimensions après la création de l'article générique !
  if (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') <> '') and
    (DimensionsArticle.TableDim <> nil) and (DimensionsArticle.TableDim.Count > 0) then
  begin
    LastError := DimensionsArticle.UpdateDB(TobArt, TobZoneLibre, ArtError);
    if LastError <> 0 then
    begin
      // DCA - FQ MODE 10444
      // Recherche des libellés des dimensions de l'article non mis à jour
      Sql := 'select DI1.GDI_LIBELLE,DI2.GDI_LIBELLE,DI3.GDI_LIBELLE,DI4.GDI_LIBELLE,DI5.GDI_LIBELLE ' +
             'from ARTICLE ' +
             'left join DIMENSION DI1 on DI1.GDI_TYPEDIM="DI1" and DI1.GDI_GRILLEDIM=GA_GRILLEDIM1 and DI1.GDI_CODEDIM=GA_CODEDIM1 ' +
             'left join DIMENSION DI2 on DI2.GDI_TYPEDIM="DI2" and DI2.GDI_GRILLEDIM=GA_GRILLEDIM2 and DI2.GDI_CODEDIM=GA_CODEDIM2 ' +
             'left join DIMENSION DI3 on DI3.GDI_TYPEDIM="DI3" and DI3.GDI_GRILLEDIM=GA_GRILLEDIM3 and DI3.GDI_CODEDIM=GA_CODEDIM3 ' +
             'left join DIMENSION DI4 on DI4.GDI_TYPEDIM="DI4" and DI4.GDI_GRILLEDIM=GA_GRILLEDIM4 and DI4.GDI_CODEDIM=GA_CODEDIM4 ' +
             'left join DIMENSION DI5 on DI5.GDI_TYPEDIM="DI5" and DI5.GDI_GRILLEDIM=GA_GRILLEDIM5 and DI5.GDI_CODEDIM=GA_CODEDIM5 ' +
             'where GA_ARTICLE like "' + ArtError + '%"' ;
      QQ := OpenSql (Sql, True);
      if not QQ.Eof then
      begin
        LastErrorMsg := TraduireMemoire('Article ' + ArtError + ' Dimension : ') ;
        for ii := 0 to 4 do if QQ.Fields[ii].AsString <> '' then LastErrorMsg := LastErrorMsg + QQ.Fields[ii].AsString + ' - ';
        LastErrorMsg := LastErrorMsg + TexteMessage[LastError];
        PGIBox (LastErrorMsg, 'Modification non enregistrée');
      end;
      Ferme (QQ);
      if ArtError <> '' then itemDimSetFocus(DimensionsArticle, nil, ArtError, 1, 'GA_CODEBARRE');
    end;
  end;

  SetControlEnabled('BINSERTNOMENC', True);
  SetControlEnabled('BNOMENC', True);
  {$IFDEF BTP}
  SetControlEnabled('BAPPELVARIABLE', True);
  SetControlEnabled('BNOMENC1', true);
  {$ENDIF}
  SetControlEnabled('BPOPMENU', True);
  SetControlEnabled('BPOPCOMPLEMENT', True);
  SetControlEnabled('BMEMO', True);
  SetControlEnabled('B_DUPLICATION', True);
  SetControlEnabled('BCODEBARRE', true);

  {$IFDEF GPAOLIGHT}
    SetBtArtNatState; 
  {$ENDIF}
end;
{$ENDIF}

procedure TOM_Article.OnDefaire(DataSet: TDataSet);
begin
  {$IFNDEF EAGLCLIENT}
  InsertionOff(DS);
  {$ENDIF}
  OnLoadRecord;
end;

// Modif par lot

procedure TOM_ARTICLE.SetArguments(StSQL: string);
var Critere, ChampMul, ValMul: string;
  x, y: integer;
  Ctrl: TControl;
  Fiche: TFFiche;
begin
  SetControlVisible('BSTOP', TRUE);

  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
  Fiche := TFFiche(ecran);
  repeat
    Critere := uppercase(Trim(ReadTokenPipe(StSQL, '|')));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        ChampMul := copy(Critere, 1, x - 1);
        ValMul := copy(Critere, x + 1, length(Critere));
        y := pos(',', ValMul);
        if y <> 0 then ValMul := copy(ValMul, 1, length(ValMul) - 1);
        if copy(ValMul, 1, 1) = '"' then ValMul := copy(ValMul, 2, length(ValMul));
        if copy(ValMul, length(ValMul), 1) = '"' then ValMul := copy(ValMul, 1, length(ValMul) - 1);
        SetField(ChampMul, ValMul);
        Ctrl := TControl(Fiche.FindComponent(ChampMul));
        if Ctrl = nil then exit;
        {$IFDEF EAGLCLIENT}
        if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) or (Ctrl is TCustomEdit) then TEdit(Ctrl).Font.Color := clRed
        else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color := clRed;
        {$ELSE}
        if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color := clRed
        else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color := clRed;
        {$ENDIF}
      end;
    end;
  until Critere = '';
end;

procedure TOM_Article.OnChangeField(F: TField);
var TValMarge: THEdit;
  TRefMarge: THValComboBox;
  Pa, Pv, Pr: double;
  IsTenueStock: boolean;
  RecalcAutoHT, RecalcAutoTTC: boolean;
  ArtSpecif, NomChamps: string;
  // Modif BTP
  //Requete : String;
  RecalcAutoPR: boolean;
  {$IFDEF BTP}
  Requete : String;
  QQ: TQuery;
  {$IFDEF EAGLCLIENT}
  Naturepres: THEdit;
  {$ELSE}
  Naturepres: THDBEdit;
  {$ENDIF}
  {$ENDIF}
  CodeSSTarif : string;
  ValeurPlus : string;
  CodeArticle : String;
begin
	CodeSSTArif := '';

  if CodeSousTarif <> nil then CodeSSTarif := CodeSousTarif.Value;

  if CalculEnCours then exit; // Reaffectation manuelle, pas de OnChangeField.

  if copy(ECRAN.name, 1, 12) = 'GCARTICLEDIM' then exit;

  SetControlEnabled('BINSERTNOMENC', not (DS.State in [dsEdit, dsInsert]));

  NomChamps := F.FieldName;

  //
  // Modif BTP 06/02/2001 + Mode 07/01/2002
  //
  if DPRAUTO <> nil then RecalcAutoPR  := DPRAUTO.Checked;
  if CALCAUTOHT <> nil then RecalcAutoHT  := CALCAUTOHT.Checked;
  if CALCAUTOTTC <> nil then RecalcAutoTTC := CALCAUTOTTC.Checked;
  {$IFDEF GPAOLIGHT}
  SetBtArtNatState;
  {$ENDIF}

  if (NomChamps = 'GA_DIMMASQUE') and
    (GetField('GA_DIMMASQUE') <> OldDimMasque) then
  begin
    OldDimMasque := GetField('GA_DIMMASQUE');
    if (F.asString <> '') then InitDimensionsArticle
    else if DimensionsArticle <> nil then
    begin
      DimensionsArticle.free;
      DimensionsArticle := nil;
    end;
  end

  else if (NomChamps = 'GA_CODEARTICLE') then
  begin
{$IFDEF BTP}
    if (GetField('GA_CODEARTICLE') <> Prefixe) then
{$ELSE}
    if (GetField('GA_CODEARTICLE') <> '') then
{$ENDIF}
    begin
      SetControlEnabled('Pages', TRUE);
      if (DS <> nil) and (DS.State in [dsInsert]) then
      begin
        SetField('GA_ARTICLE', CodeArticleUnique2(GetField('GA_CODEARTICLE'), ''));
{$IFDEF BTP}
        if ExisteSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + GetControlText ('GA_CODEARTICLE') + '"') then
        begin
          SetField('GA_CODEARTICLE', Prefixe);
          SetLastError (59, 'GA_CODEARTICLE');
          exit;
        end;
        //Modif FV - FQ FS#134 - VERRE & METAL : Pb si codification article avec le signe % : 28082012
        CodeArticle := GetField('GA_CODEARTICLE');
        if SupprimeCaracteresSpeciaux(CodeArticle,'-',False,true) then
        begin
          SetField('GA_CODEARTICLE', CodeArticle);
          SetLastError (60, 'GA_CODEARTICLE');
          exit;
        end;
        //
        if (GetField('GA_TYPEARTICLE') = 'OUV') then
        begin
          if ExisteSql ('SELECT GNE_NOMENCLATURE FROM NOMENENT WHERE GNE_NOMENCLATURE="' + GetControlText ('GA_CODEARTICLE') + '"') then
          begin
            SetField('GA_CODEARTICLE', Prefixe);
            SetLastError (38, 'GA_CODEARTICLE');
            exit;
          end;
        end;
{$ENDIF}
      end;
    end;
  end;

  // SPECIF AFFAIRE ////////////////////////
  ArtSpecif := IsArticleSpecif(TypeArticle);
  if ArtSpecif = 'FicheAffaire' then
  begin
    FaireCalcul := false; // gm le 21/09/01 corrige le message"vous n'avez pas calculé la base tarif"
  end
    //////////////////////////////////////////
  else if (copy(Ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(ecran.name,1,9)='BTARTPARC') or (copy(Ecran.name, 1, 13) = 'AFARTPOURCENT') then
  begin
  end
    //////////////////////////////////////////
    // SPECIF MODE ////////////////////////
  else if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    {$IFDEF CHR}
  else if (ecran.name = 'HRPRESTATIONS') then
  begin
  end
    {$ENDIF}
    //////////////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else
  begin

  //uniquement en line
  {*
    if (copy(ecran.name, 1, 12) = 'BTARTICLE_S1') and (NomChamps = 'GA_NATUREPRES') then
  *}
    if (copy(ecran.name, 1, 12) = 'BTPRESTATION') and (NomChamps = 'GA_NATUREPRES') then
    begin
      //SetControlEnabled('BPROFIL', true);
      {$IFDEF EAGLCLIENT}
      Naturepres := THEdit(GetControl('GA_NATUREPRES'));
      {$ELSE}
      Naturepres := THDBEdit(GetControl('GA_NATUREPRES'));
      {$ENDIF}
      THLabel(GetControl('TGA_LIBNATURE')).caption := '';
      if Naturepres <> nil then
      begin
        Requete := 'Select BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES="' + Getfield('GA_NATUREPRES') + '"';
        QQ := OpenSQL(requete, TRUE);
        if not QQ.EOF then THLabel(GetControl('TGA_LIBNATURE')).caption := QQ.findfield('BNP_LIBELLE').AsString;
        ferme(QQ);
      end;
    end;
    //
    if (NomChamps = 'GA_PAHT') or (NomChamps = 'GA_PVHT') or (NomChamps = 'GA_DPR') or (NomChamps = 'REFMARGE') then
    begin
      // Modif BTP
      //if (NomChamps = 'GA_PAHT') and ((TFFiche(Ecran).fTypeAction = taCreat) or duplicFromFiche) then SetField('GA_DPA', Getfield('GA_PAHT'));
      // ----
      TRefMarge := THValComboBox(GetControl('REFMARGE'));
      TValMarge := THEDIT(GetControl('VALEURMARGE'));
      if (TRefMarge.ItemIndex = (-1)) then TRefMarge.ItemIndex := 4;
      //if (TRefMarge.ItemIndex = (-1)) then TRefMarge.ItemIndex := 3;
      if GetField('GA_PAHT') <> Null then Pa := GetField('GA_PAHT') // Mettre en place en fonction du vrai pa (dpa ou pAht)
      else
      begin
        Pa := 0;
        SetField('GA_PAHT', Pa);
      end;
      if GetField('GA_DPR') <> Null then Pr := GetField('GA_DPR')
      else
      begin
        Pr := 0;
        SetField('GA_DPR', Pr);
      end;
      if GetField('GA_PVHT') <> Null then Pv := GetField('GA_PVHT')
      else
      begin
        Pv := 0;
        SetField('GA_PVHT', Pv);
      end;
      DetermineValorisation(pa, pr, pv);
      AfficheMarge(TRefMarge.ItemIndex, pa, pv, pr);
      //TValMarge.Text := FloatToStrF(CalculMarge(TRefMarge.ItemIndex, Pa, Pr, Pv), ffFixed, 7, 2);
    end;
    if (NomChamps = 'GA_PAUA')           Or
       (NomChamps = 'GA_QUALIFUNITEACH') or
       (NomChamps = 'GA_COEFCONVQTEACH') or
       (Nomchamps = 'GA_QUALIFUNITEVTE') or
       (Nomchamps = 'GA_COEFCONVQTEVTE') then
    begin
      SetField('GA_PAHT', CalculPAHT);
      SetInfosPrixTarif;
      DefiniAffichage;
      (*
      if (NomChamps ='GA_QUALIFUNITEACH') then
      begin
        if GetField('GA_QUALIFUNITEACH')<>'' then
        BEGIN
          SetControlEnabled('GA_COEFCONVQTEACH',True);
          SetControlEnabled('GA_PAUA',True);
        END else
        BEGIN
          SetControlEnabled('GA_COEFCONVQTEACH',false);
          SetControlEnabled('GA_PAUA',false);
        end;
      end;
      *)
    end;

  end;

  // Mise en forme de la fiche suivant la valeur de certain champ
  // Modif BTP ////////////////////////
  if (copy(Ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(Ecran.name, 1, 13) = 'AFARTPOURCENT') then
  begin
  end else
    // SPECIF AFFAIRE ////////////////////////
    if ArtSpecif = 'FicheAffaire' then
  begin
  end else
    //////////////////////////////////////////
    // SPECIF MODE ////////////////////////
    if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    //////////////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else
  begin
    if (NomChamps = 'GA_TENUESTOCK') then
    begin
      IsTenueStock := (GetField('GA_TENUESTOCK') = 'X');
      SetControlEnabled('GA_LOT', (IsTenueStock and not Consultation));
      SetControlEnabled('GA_QUALIFUNITESTO', (IsTenueStock and not Consultation));
{$IFDEF CCS3}
{$IFDEF PREMIUM}
      if VH_GC.GCAchatStockSeria then
{$ENDIF}
{$ENDIF}
       SetControlEnabled('GA_CONTREMARQUE', (not IsTenueStock and not Consultation));
      //SetControlEnabled('GA_NUMEROSERIE', (IsTenueStock and Consultation)) ;
      SetControlProperty('GA_PMRP', 'ReadOnly', IsTenueStock);
      SetControlProperty('GA_PMRP', 'TabStop', not IsTenueStock);
      if (GetField('GA_TYPEARTICLE') = 'MAR') and (not IsTenueStock) then // Lecture seule pour les nomenclature
      begin
        SetControlProperty('GA_PMAP', 'ReadOnly', False);
        SetControlProperty('GA_PMAP', 'TabStop', True);
        SetControlProperty('GA_PMAP', 'color', clWindow);
      end;
      if not IsTenueStock then
      begin
        {$IFDEF EAGLCLIENT}
        //
        // Modif BTP du 06/02/2001
        //
        {$IFDEF BTP}
        if (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S') then
        begin
          SetField('GA_CALCPRIXPR', 'AUC');
          THValComboBox(GetControl('GA_CALCPRIXPR')).OnClick(nil);
        end;
        {$ENDIF}
        // -------------------------------------
        if copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S' then
        begin
          SetField('GA_CALCPRIXHT', 'AUC');
          THValComboBox(GetControl('GA_CALCPRIXHT')).OnClick(nil);
        end;
        if copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S' then
        begin
          SetField('GA_CALCPRIXTTC', 'AUC');
          THValComboBox(GetControl('GA_CALCPRIXTTC')).OnClick(nil);
        end;
        {$IFDEF BTP}
        THValComboBox(GetControl('GA_CALCPRIXPR')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        {$ENDIF}
        // --------------------
        THValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        THValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        {$ELSE}
        // Modif BTP du 06/02/2001
        {$IFDEF BTP}
        if (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S') then
        begin
          SetField('GA_CALCPRIXPR', 'AUC');
          THDBValComboBox(GetControl('GA_CALCPRIXPR')).OnClick(nil);
        end;
        {$ENDIF}
        // -------------------------------------
        if copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S' then
        begin
          SetField('GA_CALCPRIXHT', 'AUC');
          THDBValComboBox(GetControl('GA_CALCPRIXHT')).OnClick(nil);
        end;
        if copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S' then
        begin
          SetField('GA_CALCPRIXTTC', 'AUC');
          THDBValComboBox(GetControl('GA_CALCPRIXTTC')).OnClick(nil);
        end;
        //
        // Modif BTP du 06/02/2001
        //
        {$IFDEF BTP}
        if assigned(GetControl('GA_QUALIFUNITESTO'))  then SetControlEnabled ('GA_QUALIFUNITESTO',false);
        if assigned(GetControl('TGA_QUALIFUNITESTO')) then SetControlEnabled('TGA_QUALIFUNITESTO',false);
        if assigned(GetControl('GA_COEFCONVQTEVTE'))  then
        begin
    			SetControlEnabled('GA_COEFCONVQTEVTE',false);
    			THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := true;
        end;
        THDBValComboBox(GetControl('GA_CALCPRIXPR')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        AvertirTable ('BTCALCPRIXPR');
        {$ENDIF}
        // --------------------
        THDBValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        THDBValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
        AvertirTable ('BTCALCPRIXHT');
        AvertirTable ('BTCALCPRIXTTC');
        {$ENDIF}
        SetControlProperty('GA_PMRP', 'color', clWindow);
      end;
      if IsTenueStock then
      begin
        SetControlProperty('GA_PMAP', 'ReadOnly', True);
        SetControlProperty('GA_PMAP', 'TabStop', False);
        SetControlProperty('GA_PMRP', 'Color', clbtnFace);
        SetControlProperty('GA_PMAP', 'Color', clbtnFace);
        {$IFDEF EAGLCLIENT}
        //
        // Modif BTP du 06/02/2001
        //
        {$IFDEF BTP}
        THValComboBox(GetControl('GA_CALCPRIXPR')).Plus := '';
        {$ENDIF}
        // -------------------------------------------------------
        THValComboBox(GetControl('GA_CALCPRIXHT')).Plus := '';
        THValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := '';
        {$ELSE}
        //
        // Modif BTP du 06/02/2001
        //
        {$IFDEF BTP}
        THDBValComboBox(GetControl('GA_CALCPRIXPR')).Plus := '';
        {$ENDIF}
        // -------------------------------------------------------
        if assigned(GetControl('GA_QUALIFUNITESTO'))  then SetControlEnabled ('GA_QUALIFUNITESTO',True);
        if assigned(GetControl('TGA_QUALIFUNITESTO')) then SetControlEnabled('TGA_QUALIFUNITESTO',True);
        if assigned(GetControl('GA_COEFCONVQTEVTE'))  then
        begin
  			  SetControlEnabled('GA_COEFCONVQTEVTE', GetControltext('GA_QUALIFUNITESTO')<>'');
  			  THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := (GetControltext('GA_QUALIFUNITESTO')='');
        end;
        THDBValComboBox(GetControl('GA_CALCPRIXHT')).Plus := '';
        THDBValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := '';
        {$ENDIF}
      end;
    end;
  end;
  if NomChamps = 'GA_DPRAUTO' then
  begin
    if DPRAUTO <> nil then
    begin
      if DPRAUTO.Checked      then SetControlProperty('GA_DPR', 'Color', clbtnFace)
                              else SetControlProperty('GA_DPR', 'Color', clWindow);
    end;
  end else
  if NomChamps = 'GA_CALCAUTOHT' then
  begin
    if CALCAUTOHT <> nil then
    begin
      if CALCAUTOHT.checked   then SetControlProperty('GA_PVHT', 'Color', clbtnFace)
                              else SetControlProperty('GA_PVHT', 'Color', clWindow);
    end;
  end else
  if NomChamps = 'GA_CALCAUTOTTC' then
  begin
    if CALCAUTOTTC <> nil then
    begin
      if CALCAUTOTTC.checked  then SetControlProperty('GA_PVTTC', 'Color', clbtnFace)
                              else SetControlProperty('GA_PVTTC', 'Color', clWindow);
    end;
  end else
  if (NomChamps = 'GA_DIMMASQUE') then
  begin
    // Modif BTP // attention, ne fonctionne pas pour les articles dimensionnes
    // if (GetField('GA_DIMMASQUE')<>'')
    if (GetField('GA_STATUTART') <> 'DIM') then
    begin
      if (GetField('GA_DIMMASQUE') <> '') then
        SetField('GA_STATUTART', 'GEN')
      else
        SetField('GA_STATUTART', 'UNI');
    end;
  end;

  // DCA - FQ MODE  10792
  if (NomChamps = 'GA_FOURNPRINC') and
     (ArtSpecif <> 'FicheModePre') and (ArtSpecif <> 'FicheModeFi') then
  begin
    GestionFournisseurPrincipal;
  end;
  //
  // Modif BTP du 06/02/2001
  //
  {$IFDEF BTP}
  if (NomChamps = 'GA_CALCPRIXPR') then
  begin
    if (GetField('GA_CALCPRIXPR') = 'AUC') or (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S') then
    begin
      SetControlProperty('GA_COEFFG', 'color', clbtnFace);
      SetControlProperty('GA_DPR', 'Color', clWindow);
      FaireCalcul := False;
    end else
    begin
      if (DS.State in [DsEdit, DsInsert]) and (RecalcAutoPR) then CalculBaseTarifPR(False);
      if (DS.State in [dsEdit]) then
      begin
        FaireCalcul := True;
        ClickActuPrix := False;
        ModifChampPR := not RecalcAutoPR; // control pour le calcul
      end else if DS.State in [dsInsert] then TestFaireCalcul;
      SetControlProperty('GA_COEFFG', 'color', clWindow);
      SetControlEnabled('GA_COEFFG', True);
      SetControlEnabled('GA_DPRAUTO', True);
      if (GetField('GA_DPRAUTO') = 'X') then SetControlProperty('GA_DPR', 'Color', clbtnFace)
      																	else SetControlProperty('GA_DPR', 'Color', clWindow);
    end;
  end;
  {$ENDIF}
  // ------------------------------------------------

  // Mise en forme de la fiche suivant la valeur du champ Base prix HT
  // SPECIF MODE ////////////////////////
  if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    //////////////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else if (NomChamps = 'GA_CALCPRIXHT') then
  begin
    if (GetField('GA_CALCPRIXHT') = 'FOU') and (GetControlText('GA_FOURNPRINC') = '') then
    begin
      if not DuplicFromFiche then SetLastError(25, 'GA_CALCPRIXHT');
      exit;
    end
    else if (GetField('GA_CALCPRIXHT') = 'AUC') or (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') then
    begin
      SetControlEnabled('GA_COEFCALCHT',false);
      SetControlEnabled('GA_CALCAUTOHT',false);
      //
      if (DS.State in [DsEdit, DsInsert]) then
      begin
        THEdit(GetControl('GA_COEFCALCHT')).text := '0.00';
      end;
      if not (ctxMode in V_PGI.PGIContexte) then
      begin
        //  BBI, fiche correction 10410
        // SetControlText('GA_CALCAUTOHT','X');
        // SetControlEnabled('GA_CALCAUTOHT', False);
        //  BBI, fin fiche correction 10410
        SetControlProperty('GA_PVHT', 'Color', clWindow);
        FaireCalcul := False;
      end;
    end
    else
    begin
      if (DS.State in [DsEdit, DsInsert]) and (RecalcAutoHT) then CalculBaseTarifHT(False);
      if (DS.State in [dsEdit]) then
      begin
        FaireCalcul := True;
        ClickActuPrix := False;
        ModifChampHT := not RecalcAutoHT; // control pour le calcul
      end
      else if DS.State in [dsInsert] then TestFaireCalcul;
      SetControlEnabled('GA_COEFCALCHT',true);
      SetControlEnabled('GA_CALCAUTOHT',true);

      if not (ctxMode in V_PGI.PGIContexte) then
      begin
        //  BBI, fiche correction 10410
        //       SetControlProperty('GA_COEFCALCHT','color',clWindow);
        //       SetControlProperty('GA_COEFCALCHT', 'color', clWindow);
        //  BBI, fin fiche correction 10410
        if CALCAUTOHT.checked then SetControlProperty('GA_PVHT', 'Color', clbtnFace)
        											else SetControlProperty('GA_PVHT', 'Color', clWindow);
      end;
    end;
  end;

  // Mise en forme de la fiche suivant la valeur du champ Base prix TTC
  // SPECIF MODE ////////////////////////
  if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    //////////////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else if (NomChamps = 'GA_CALCPRIXTTC') then
  begin
    if (GetField('GA_CALCPRIXTTC') = 'AUC') or (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S') then
    begin
      if not (ctxMode in V_PGI.PGIContexte) then
      begin
        //  BBI, fiche correction 10410
        //  SetControlProperty('GA_COEFCALCTTC', 'color', clbtnFace);
        //  BBI, fin fiche correction 10410
        SetControlProperty('GA_PVTTC', 'Color', clWindow);
        FaireCalcul := False;
      end;
    end else
    begin
      if (DS.State in [DsEdit, DsInsert]) and (RecalcAutoTTC) then CalculBaseTarifTTC(False);
      if (DS.State in [dsEdit]) then
      begin
        FaireCalcul := True;
        ClickActuPrix := False;
        ModifChampTTC := not RecalcAutoTTC; // control pour le calcul
      end
      else if DS.State in [dsInsert] then TestFaireCalcul;
      if not (ctxMode in V_PGI.PGIContexte) then
      begin
        //  BBI, fiche correction 10410
        //  SetControlProperty('GA_COEFCALCTTC', 'color', clWindow);
        //  BBI, fin fiche correction 10410
        if CALCAUTOHT.Checked   then SetControlProperty('GA_PVTTC', 'Color', clbtnFace)
        												else SetControlProperty('GA_PVTTC', 'Color', clWindow);
      end;
    end;
  end;
  //
  // Modif BTP du 06/02/2001
  //
  if BACTUPRIX <> nil then
  begin
    if ((Getfield('GA_CALCPRIXPR') = 'AUC')  and
        (GetField('GA_CALCPRIXTTC') = 'AUC') and
        (GetField('GA_CALCPRIXHT')  = 'AUC')) or
       ((copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') and
        (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S') and
        (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S')) then BACTUPRIX.Enabled:=False
    else if (GetField('GA_TYPENOMENC') <> 'MAC')       then BACTUPRIX.enabled:= not Consultation;
  end;

  {*
  $ELSE
  if ((GetField('GA_CALCPRIXTTC') = 'AUC') and
      (GetField('GA_CALCPRIXHT') = 'AUC')) or
     ((copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') and
    (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S')) then BACTUPRIX.enabled := False
  else if (GetField('GA_TYPENOMENC') <> 'MAC') then SetControlEnabled('BACTUPRIX', (not Consultation));
  $ENDIF
  *}
  // ----

  // Calcul automatique sur modification d'un des éléments de calcul
  // SPECIF BTP ////////////////////////
  if (copy(Ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(Ecran.name, 1, 13) = 'AFARTPOURCENT') then
  begin
  end
  else
    //////////////////////////////////////////
    // SPECIF MODE ////////////////////////
    if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    //////////////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else
  begin
    if (NomChamps = 'GA_DPA') or (NomChamps = 'GA_DPR') or (NomChamps = 'GA_PAHT') or (NomChamps = 'GA_PRHT') or (NomChamps = 'GA_PMAP') or (NomChamps =
      'GA_PMRP') then
    begin
      if (DS.State in [dsEdit, dsInsert]) then
      begin
        //
        // Modif BTP
        //
        {$IFDEF BTP}
        if (RecalcAutoPR) and ((NomChamps = 'GA_DPA') or (NomChamps = 'GA_PAHT') or (NomChamps = 'GA_PMAP')) then
          CalculBaseTarifPR(False);
        {$ENDIF}
        // ---------------------------------------------
        if (CtxMode in V_PGI.PGIContexte) and (RecalcAutoPR) then
        begin
          if ((NomChamps = 'GA_PAHT') and (CalcPAPR = 'AR')) or
            ((NomChamps = 'GA_PRHT') and (CalcPAPR = 'RA'))
            then CalculBaseTarifPR(False);
        end;
        if RecalcAutoHT then CalculBaseTarifHT(False);
        if RecalcAutoTTC then CalculBaseTarifTTC(False);
      end;
      if DS.State in [DsInsert] then TestFaireCalcul
      else if DS.State in [DsEdit] then
      begin
        FaireCalcul   := True;
        ClickActuPrix := False;
        ModifChampHT  := not RecalcAutoHT;
        ModifChampTTC := not RecalcAutoTTC;
      end;
    end;
  end;

  //
  // Modif BTP 06/02/2001
  //
  {$IFDEF BTP}
  if (NomChamps = 'GA_COEFFG') OR (NomChamps = 'GA_DPRAUTO') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoPR then CalculBaseTarifPR(False);
    if RecalcAutoHT then CalculBaseTarifHT(False);
    if RecalcAutoTTC then CalculBaseTarifTTC(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampPR := not RecalcAutoPR;
    end;
  end;

  if NomChamps = 'GA_QUALIFUNITEVTE' then
  begin
    if (GetField('GA_TENUESTOCK') <> 'X') then
    begin
      SetField('GA_QUALIFUNITESTO', GetControlText ('GA_QUALIFUNITEVTE'));
      if assigned(GetControl('GA_COEFCONVQTEVTE')) then
      begin
  		  SetControlEnabled('GA_COEFCONVQTEVTE',false);
        SetField('GA_COEFCONVQTEVTE',0);
      end;
    end else
    begin
      if assigned(GetControl('GA_COEFCONVQTEVTE')) then
      begin
  		  SetControlEnabled('GA_COEFCONVQTEVTE',GetControltext('GA_QUALIFUNITESTO')<>'');
        THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := (GetControltext('GA_QUALIFUNITESTO')='');
      end;
    end;
  end;

  if NomChamps = 'GA_QUALIFUNITESTO' then
  begin
    if GetControltext('GA_QUALIFUNITESTO')<>'' then
    begin
      if assigned(GetControl('GA_COEFCONVQTEVTE')) then SetControlEnabled('GA_COEFCONVQTEVTE',GetControltext('GA_QUALIFUNITESTO')<>'');
    end else
    begin
	    if assigned(GetControl('GA_COEFCONVQTEVTE')) then
      begin
        SetControlEnabled('GA_COEFCONVQTEVTE',false);
        SetField('GA_COEFCONVQTEVTE',0);
      end;
    end;
  end;

  if assigned(GetControl('GA_COEFCONVQTEVTE')) then
  begin
    if NomChamps = 'GA_COEFCONVQTEVTE' then
    begin
      SetInfosPrixTarif;
    end;
  end;

  {$ENDIF}

  if (CtxMode in V_PGI.PGIContexte) then
    if (NomChamps = 'GA_COEFFG') then
    begin
      if (DS.State in [dsEdit, dsInsert]) and RecalcAutoPR then CalculBaseTarifPR(False);
      if RecalcAutoHT then CalculBaseTarifHT(False);
      if RecalcAutoTTC then CalculBaseTarifTTC(False);
      if DS.State in [DsInsert] then TestFaireCalcul
      else if DS.State in [DsEdit] then
      begin
        FaireCalcul := True;
        ClickActuPrix := False;
        ModifChampPR := not RecalcAutoPR;
      end;
    end;

  // ---------------------------------------

  if (NomChamps = 'GA_COEFCALCHT') or (NomChamps = 'GA_CALCAUTOHT') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoHT then CalculBaseTarifHT(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampHT := not RecalcAutoHT;
    end;
  end;

  if (NomChamps = 'GA_COEFCALCTTC') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoTTC then CalculBaseTarifTTC(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampTTC := not RecalcAutoTTC;
    end;
  end;

  //
  // Modif BTP 06/02/2001
  //
  {$IFDEF BTP}
  if (NomChamps = 'GA_DPR') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoPR then CalculCoefTarifPR(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampPR := not RecalcAutoPR;
    end;
  end;
  {$ENDIF}

  if (CtxMode in V_PGI.PGIContexte) then
  begin
    if ((CalcPAPR = 'AR') and (NomChamps = 'GA_PRHT')) or
      ((CalcPAPR = 'RA') and (NomChamps = 'GA_PAHT')) then
    begin
      if (DS.State in [dsEdit, dsInsert]) and RecalcAutoPR then CalculCoefTarifPR(False);
      if DS.State in [DsInsert] then TestFaireCalcul
      else if DS.State in [DsEdit] then
      begin
        FaireCalcul := True;
        ClickActuPrix := False;
        ModifChampPR := not RecalcAutoPR;
      end;
    end;
  end;

  if (NomChamps = 'GA_PVHT') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoHT then CalculCoefTarifHT(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampHT := not RecalcAutoHT;
    end;
  end
  else if (NomChamps = 'GA_PVTTC') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and RecalcAutoTTC then CalculCoefTarifTTC(False);
    if DS.State in [DsInsert] then TestFaireCalcul
    else if DS.State in [DsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampTTC := not RecalcAutoTTC;
    end;
    {$IFDEF BTP}
    if GetField(NomChamps) = 0 then SetControlText('TGA_PVTTC', TraduireMemoire('Base Tarif T.T.C (non déterminé)'))
    else SetControlText('TGA_PVTTC', TraduireMemoire('Base Tarif T.T.C'));
    {$ENDIF}
  end;

  // DEBUT CHR le 26/02/02
  if (ctxChr in V_PGI.PGIContexte) then //CHR
  begin
    if (NomChamps = 'GA_TYPEARTICLE') then
      setcontrolProperty('GA_TYPEARTICLE', 'Plus', 'AND (CO_CODE = "PRE" OR CO_CODE = "FI" OR CO_CODE ="NOM" OR CO_CODE ="MAR")');
  end;
  // FIN CHR le 26/02/02

  if (NomChamps = 'GA_ARRONDIPRIX') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and (RecalcAutoHT) then CalculBaseTarifHT(False);
    if Ds.State in [dsInsert] then TestFaireCalcul
    else if Ds.State in [dsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampHT := not RecalcAutoHT;
    end;
  end
  else if (NomChamps = 'GA_ARRONDIPRIXTTC') then
  begin
    if (DS.State in [dsEdit, dsInsert]) and (RecalcAutoTTC) then CalculBaseTarifTTC(False);
    if Ds.State in [dsInsert] then TestFaireCalcul
    else if Ds.State in [dsEdit] then
    begin
      FaireCalcul := True;
      ClickActuPrix := False;
      ModifChampTTC := not RecalcAutoTTC;
    end;
  end
  else if (NomChamps = 'GA_QUALIFMARGE') then
  begin
    {$IFDEF GPAO}
    SetControlEnabled('GA_MARGEMINI', not (GetField('GA_QUALIFMARGE') = ''));
    {$ENDIF}
    SetControlProperty('GA_MARGEMINI', 'TabStop', not (GetField('GA_QUALIFMARGE') = ''));
    SetControlProperty('GA_MARGEMINI', 'ReadOnly', (GetField('GA_QUALIFMARGE') = ''));
    {$IFDEF BTP}
    if GetField('GA_QUALIFMARGE') = '' then SetField('GA_QUALIFMARGE', 'CO');
    {$ENDIF}
    if GetField('GA_QUALIFMARGE') = '' then SetControlProperty('GA_MARGEMINI', 'Color', clbtnFace)
    else SetControlProperty('GA_MARGEMINI', 'Color', clWindow);
  end
  else if (NomChamps = 'GA_FAMILLENIV1') and (GetParamSoc('SO_GCFAMHIERARCHIQUE') = True) then
  begin
    if Getfield('GA_FAMILLENIV1') = '' then
    begin
      Setfield('GA_FAMILLENIV2', '');
      Setfield('GA_FAMILLENIV3', '');
      SetControlEnabled('GA_FAMILLENIV2', false);
      SetControlEnabled('GA_FAMILLENIV3', false);
    end
    else SetControlEnabled('GA_FAMILLENIV2', true);
  end
  else if (NomChamps = 'GA_FAMILLENIV2') and (GetParamSoc('SO_GCFAMHIERARCHIQUE') = True) then
  begin
    if Getfield('GA_FAMILLENIV2') = '' then
    begin
      Setfield('GA_FAMILLENIV3', '');
      SetControlEnabled('GA_FAMILLENIV3', false);
    end
    else
      SetControlEnabled('GA_FAMILLENIV3', true);
  end;

  {$IFDEF GPAO}
  if (NomChamps = 'GA_ESTPROFIL') and (DS.State in [dsInsert])
    and (GetField('GA_ESTPROFIL') = 'X')
    and (not DuplicFromFiche) and (not DuplicFromMenu) then
  begin
    SetField('GA_PROFILARTICLE', '');
    SetField('GA_ARTICLEPROFIL', '');
    wProfilArticleState;
  end
  else if (F.FieldName = 'GA_PROFILARTICLE') then
    wProfilArticleState
  else if (F.FieldName = 'GA_QUALIFUNITESTO') then
    wSetPlusUniteStock;
  {$ENDIF}

end;


{$IFDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 04/02/2003
Description .. : Initialise la propriété plus des unités de l'article
Description .. : avec le qualifiant de l'unité de stock, si les
Description .. : conversion d'unités avancées ne sont pas activée
*****************************************************************}
procedure TOM_Article.wSetPlusUniteStock;
var
  s, Plus: string;
  T: Tob;
begin
  { Initialise la propriété plus des unités comme celle de l'unité de stock,
  si les conversions avancées ne sont pas utilisées... }
  if (not GetParamSoc('SO_WCUCONVERSION')) and (getField('GA_QUALIFUNITESTO') <> '') then
  begin
    { Récupère le qualifiant de l'unité de stock }
    s := GetQualifiantMEA(GetField('GA_QUALIFUNITESTO'));
    if s <> '' then
    begin
      { Forme le Plus }
      Plus := 'GME_QUALIFMESURE="' + s + '"';
      { et l'applique aux autres unités }
      SetControlProperty('GA_UNITEQTEVTE', 'Plus', Plus);
      SetControlProperty('GA_QUALIFUNITEVTE', 'Plus', Plus);
      SetControlProperty('GA_UNITEQTEACH', 'Plus', Plus);
      SetControlProperty('GA_UNITEPROD', 'Plus', Plus);
      SetControlProperty('GA_UNITECONSO', 'Plus', Plus);
    end;
  end;
end;
{$ENDIF GPAO}

{$IFDEF GPAO}

procedure TOM_Article.wGetRetour;
begin
  if Assigned(Ecran) then
    TfFiche(Ecran).Retour := wGetValueClef1('ARTICLE', TfFiche(Ecran));
end;
{$ENDIF}

/////////////////////////////////////////////////////////////////////////////

procedure TOM_Article.TestFaireCalcul;
begin
  ClickActuPrix := False;
  //
  // Modif BTP 06/02/2001
  //
  {$IFDEF BTP}
  if (GetField('GA_CALCPRIXPR') = 'DPA') and (GetField('GA_DPA') <> 0) then
  begin
    ModifChampPR := not DPRAUTO.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXPR') = 'PAA') and (GetField('GA_PAHT') <> 0) then
  begin
    ModifChampPR := not DPRAUTO.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXPR') = 'PMA') and (GetField('GA_PMAP') <> 0) then
  begin
    ModifChampPR := not DPRAUTO.Checked;
    FaireCalcul := True;
  end;
  {$ENDIF}
  if (GetField('GA_CALCPRIXHT') = 'DPA') and (GetField('GA_DPA') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXHT') = 'DPR') and (GetField('GA_DPR') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXHT') = 'PAA') and (GetField('GA_PAHT') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXHT') = 'PRA') and (GetField('GA_PRHT') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXHT') = 'PMA') and (GetField('GA_PMAP') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXHT') = 'PMR') and (GetField('GA_PMRP') <> 0) then
  begin
    ModifChampHT := not CALCAUTOHT.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'DPA') and (GetField('GA_DPA') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'DPR') and (GetField('GA_DPR') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'PAA') and (GetField('GA_PAHT') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'PRA') and (GetField('GA_PRHT') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'PMA') and (GetField('GA_PMAP') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.Checked;
    FaireCalcul := True;
  end;
  if (GetField('GA_CALCPRIXTTC') = 'PMR') and (GetField('GA_PMRP') <> 0) then
  begin
    ModifChampTTC := not CALCAUTOTTC.Checked;
    FaireCalcul := True;
  end;
end;

// Fonction d'attribution automatique d'un code à barres

procedure TOM_Article.FormateCAB(bCtrl: boolean);
var TypeCAB: string;
  CodeCAB: string;
begin
  CodeCAB := '';
  TypeCAB := '';
  // Recherche du paramétrage du code à barres
  if RechParamCAB_Auto(GetField('GA_FOURNPRINC'), TypeCAB) then
  begin
    // Attribution du nouveau code à barres
    CodeCAB := AttribNewCodeCAB(bCtrl, GetField('GA_FOURNPRINC'));
    SetField('GA_CODEBARRE', CodeCAB);
    SetField('GA_QUALIFCODEBARRE', TypeCAB);
  end;
end;

function TOM_Article.TrimZoneCritique: boolean;
var Chiffre: integer;
begin
  Result := True;
  if Trim(GetField('GA_FOURNPRINC')) <> GetField('GA_FOURNPRINC') then
    SetField('GA_FOURNPRINC', Trim(GetField('GA_FOURNPRINC')));
  if Trim(GetField('GA_COLLECTION')) <> GetField('GA_COLLECTION') then
    SetField('GA_COLLECTION', Trim(GetField('GA_COLLECTION')));
  for Chiffre := 1 to 2 do
    if Trim(GetField('GCSTATART' + IntToStr(Chiffre))) <> GetField('GCSTATART' + IntToStr(Chiffre)) then
      SetField('GCSTATART' + IntToStr(Chiffre), Trim(GetField('GCSTATART' + IntToStr(Chiffre))));
  for Chiffre := 1 to $A do
  begin
    if Trim(GetField('GA_LIBREART' + format('%x', [Chiffre]))) <> GetField('GA_LIBREART' + format('%x', [Chiffre])) then
      SetField('GA_LIBREART' + format('%x', [Chiffre]), Trim(GetField('GA_LIBREART' + format('%x', [Chiffre]))));
    if (GetField('GA_LIBREART' + format('%x', [Chiffre])) <> '')
      and not LookUpValueExist(THEdit(Getcontrol('GA_LIBREART' + format('%x', [Chiffre])))) then
    begin
      SetActiveTabSheet('P_DIVERS');
      SetLastError(55, 'GA_LIBREART' + format('%x', [Chiffre]));
      Result := False;
      Break;
    end;
  end;
end;

procedure TOM_Article.OnUpdateRecord;
var QQ: TQuery;
  Nbr, iItem, NbrMax: integer;
  SQL, NomChamp, ValeurNomChamp: string;
  PrixUnique, ArtDimModifie: Boolean;
  CodeArt, {ArtError,} NumChrono, NewCodeCAB, ArtSpecif, TypeCAB: string;
  ArtDim, TobArtCompl: TOB;
  OngletDimension: TTAbSheet;
  PanelDimension: TPanel;
//  Err: Boolean; 
  st: string;
  pv, ht, tx: double;
  // Modif BTP
  {$IFDEF BTP}
  LibTypNomenc: string;
  lngCodeArt: integer;
  {$IFDEF EAGLCLIENT}
  Naturepres: THEdit;
  {$ELSE}
  Naturepres: THDBEdit;
  {$ENDIF}
  {$ENDIF}
  {$IFDEF GPAO}
  ErrorFieldName: string;
  TypeInitAction: TInitAction;
  Ok: Boolean;
  {$ENDIF}
  stArticle, stChronoArt, stSql: string;
  bOkChrono: boolean;
  iRecupChrono: integer;

begin

  if (ds <> nil) and (DS.State in [dsInsert]) and (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') = '') then
    SetField('GA_STATUTART', 'UNI');

  if (Pos(getField('GA_TYPEARTICLE'),'MAR;ARP')>0) and(GetField('GA_ARTICLEPARC')<> '') then
  begin
  	QQ := OpenSql ('SELECT GA_CODEARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+
    								getField('GA_ARTICLEPARC')+'" AND GA_TYPEARTICLE LIKE "PA%"',
                    true,1,'',true);
    if QQ.eof then
    begin
    	ferme (QQ);
      LastError := 41;
      LastErrorMsg := TexteMessage[LastError];
      exit;
    end;
    ferme (QQ);
  end else
  begin
  	SetField('GA_ARTICLEPARC','');
  end;

  if (GetField('GA_STATUTART') = 'DIM') then
  begin
    NewCodeCAB := GetField('GA_CODEBARRE');
    LastError := ControleCAB(GetField('GA_ARTICLE'), GetField('GA_STATUTART'), OldCodeCAB, NewCodeCAB, GetField('GA_QUALIFCODEBARRE'));
    if LastError <> 0 then
    begin
      SetLastError(LastError, 'GA_CODEBARRE');
      exit
    end;
    if NewCodeCAB <> GetField('GA_CODEBARRE') then SetField('GA_CODEBARRE', NewCodeCAB); // Maj CAB modifié par fonction de contrôle
    TobArt.GetEcran(Ecran); // Maj TobArt pour retour à l'article générique

    {$IFDEF BTP}
    if GetField('GA_TYPEARTICLE') = 'ARP' then
    begin
      ReajusteTOBDetail;
      ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_NOMENCLATURE="' + GetField('GA_ARTICLE') + '"');
      ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="' + GetField('GA_ARTICLE') + '"');
      TOBEDetail.InsertDb(nil);
      TOBDetail.InsertDb(nil);
    end;

    // Gestion des métrés
    //FV1 - 08/07/2016
    //controle si métré ok ou non pour affichage bouton excel dans fiche

    {$ENDIF BTP}
    exit;
  end;
  if (ds <> nil) and (DS.State in [dsInsert]) then
  begin
    if V_PGI.VersionDemo then
    begin
      if (ctxMode in V_PGI.PGIContexte) then
      begin // 10 articles génériques ou uniques
        QQ := OpenSQL('SELECT COUNT(*) FROM ARTICLE WHERE GA_TYPEARTICLE="MAR" AND GA_STATUTART <> "DIM"', True);
        NbrMax := 10;
      end else
      begin // 100 articles
        QQ := OpenSQL('SELECT COUNT(*) FROM ARTICLE', True);
        NbrMax := 100;
      end;
      if not QQ.EOF then Nbr := QQ.Fields[0].AsInteger else Nbr := 0;
      Ferme(QQ);
      if (Nbr >= NbrMax) then
      begin
        LastError := 41;
        LastErrorMsg := TexteMessage[LastError];
        exit;
      end;
    end;
  end;

  // SPECIF AFFAIRE
  ArtSpecif := IsArticleSpecif(TypeArticle);
  if ArtSpecif = 'FicheAffaire' then ModifChampTTC := False;

  // Modif BTP 06/02/2001
  {$IFDEF BTP}
  // COrrection à apporter
  if (DS <> nil) and (DS.State in [dsInsert]) then SetField('GA_REMISELIGNE','X');
  // ---
  if ((GetField('GA_CALCPRIXPR') = 'AUC') and (GetField('GA_CALCPRIXTTC') = 'AUC') and (GetField('GA_CALCPRIXHT') = 'AUC')) or
    ((copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') and (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S')) then FaireCalcul := False;
  //If (not ClickActuPrix) And ( ModifChampPr Or ModifChampHT Or ModifChampTTC)
  //   And FaireCalcul And (Not ModifparLot) then Begin SetLastError(33,'GA_CALCPRIXPR') ; exit ; end ;
  if FaireCalcul and (not ClickActuPrix) and (not ModifparLot) then
  begin
    if (ModifChampPR and (GetField('GA_CALCPRIXPR') <> 'AUC') and (getField('GA_DPRAUTO')='X')) then
    begin
      SetLastError(33, 'GA_CALCPRIXPR');
      exit
    end
    else if (ModifChampHT and (GetField('GA_CALCPRIXHT') <> 'AUC') and CALCAUTOHT.Checked) then
    begin
      SetLastError(33, 'GA_CALCPRIXHT');
      exit
    end
    else if (ModifChampTTC and (GetField('GA_CALCPRIXTTC') <> 'AUC') and CALCAUTOTTC.Checked) then
    begin
      SetLastError(33, 'GA_CALCPRIXTTC');
      exit
    end;
  end;
  if GetField('GA_DPA') = 0 then SetField('GA_DPA', GetField('GA_PAHT'));
  {$ELSE}
  if ((GetField('GA_CALCPRIXTTC') = 'AUC') and (GetField('GA_CALCPRIXHT') = 'AUC')) or ((copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') and
    (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S')) then FaireCalcul := False;
  if FaireCalcul and (not ClickActuPrix) and (not ModifparLot) then
  begin
    if (ModifChampHT and (GetField('GA_CALCPRIXHT') <> 'AUC')) then
    begin
      SetLastError(33, 'GA_CALCPRIXHT');
      exit
    end
    else if (ModifChampTTC and (GetField('GA_CALCPRIXTTC') <> 'AUC')) then
    begin
      SetLastError(33, 'GA_CALCPRIXTTC');
      exit
    end;
  end;
  {$ENDIF}

  // Mise à jour du fournisseur principal
  if (GetField('GA_CALCPRIXHT') = 'FOU') and (GetField('GA_FOURNPRINC') = '') then
  begin
    SetLastError(24, 'GA_CALCPRIXHT');
    exit;
  end;

  // SPECIF AFFAIRE
  if ArtSpecif = 'FicheAffaire' then
  begin
    if (GetField('GA_TYPEARTICLE') = '') then
    begin
      SetLastError(18, 'GA_TYPEARTICLE');
      exit;
    end;
  end;
  ///////////////////////

  // SPECIF MODE   le type d'article financier est obligatoire
  if (ArtSpecif = 'FicheModeFi') then
  begin
    if (GetField('GA_TYPEARTFINAN') = '') then
    begin
      SetLastError(34, 'GA_TYPEARTFINAN');
      exit;
    end;
  end;
  ///////////////////////
  {$IFDEF GESCOM}
  if (ArtSpecif = 'FicheGescomFi') then
  begin
    if (GetField('GA_TYPEARTFINAN') = '') then
    begin
      SetLastError(34, 'GA_TYPEARTFINAN');
      exit;
    end;
  end;
  {$ENDIF}

  if (GetField('GA_LIBELLE') = '') then
  begin
    SetLastError(10, 'GA_LIBELLE');
    exit;
  end;

  if (DS <> nil) and (DS.State in [dsInsert]) then SetField('GA_CODEARTICLE', Trim(GetField('GA_CODEARTICLE')));

  //
  // Modif BTP
  //
  //************************ Gestion des chronos Article automatiques *****************
  {$IFDEF BTP}
  //programmation de merde ++> oN CASSE TOUT ET ON RECOMMENCE !!!!!!!!!!!!!!
  //FS#1280 - SOCMA : la gestion automatique des codes articles et prestations ne fonctionne plus
  CodeArt := Getfield('GA_CODEARTICLE');

  if (TFFiche(Ecran).fTypeAction = taCreat) then
  begin
    if GetParamsoc('SO_GCNUMARTAUTO') then //Numérotation automatique
    begin
      if ((GetField('GA_CODEARTICLE') = '') or (copy(GetField('GA_CODEARTICLE'), length(prefixe) + 1, 255) = '')) then
      begin
        lngCodeArt := GetParamsoc('SO_GCLGNUMART'); //longueur du compte
        if GetParamsoc('SO_GCTYPSUFART') then //Utilisation du Libellé comme suffixe
        begin
          CodeArt := AttribNewCodeB('ARTICLE', 'GA_CODEARTICLE', lngCodeArt, prefixe, Getfield('GA_LIBELLE'));
        end
        else // Récupération du chrono article : gestion multi-utilisateurs
        begin
          iRecupChrono := 0;
          repeat
            inc(iRecupChrono);
            stChronoArt := GetParamSoc ( 'SO_GCCOMPTEURART', True ) ;
            CodeArt := AttribNewCode('ARTICLE', 'GA_CODEARTICLE', lngCodeArt, Prefixe, stChronoArt, '', GetParamsoc('SO_GCCABARTICLE'));
            //SAUVEGARDE DES PARAMETRE SOCIETE AVEC NOUVEAU CHRONO ARTICLE
            NumChrono := ExtraitChronoCode(CodeArt);
            StSql := 'UPDATE PARAMSOC SET SOC_DATA="' + NumChrono + '" WHERE SOC_NOM="SO_GCCOMPTEURART" AND SOC_DATA="' + stChronoArt + '"';
            bOkChrono := (ExecuteSQL(stSQL) = 1);
          until ((bOkChrono) or (iRecupChrono > 20));
          if not bOkChrono then
          begin
            SetLastError(57, CodeArt);
            Exit;
          end;
        end;
        if (Length(Prefixe) > 0) and (copy(CodeArt, 1,Length(prefixe)) <> prefixe) then //Numérotation auto avec prefixe
        begin
          CodeArt := copy(Prefixe + CodeArt, 1, lngCodeArt);
        end
      end;
    end;
    SetField('GA_CODEARTICLE', CodeArt);
    if (GetField('GA_STATUTART') = 'GEN') then DimensionsArticle.CodeArticle := CodeArt;
  end;       
  {$ENDIF}
  //---------

  if (GetField('GA_CODEARTICLE') = '') then
  begin
    SetFocusControl('GA_CODEARTICLE');
    // SPECIF AFFAIRE ////////////////////////
    if ArtSpecif = 'FicheAffaire' then LastError := 19
      //////////////////////////////////////////
    else LastError := 20;
    LastErrorMsg := TexteMessage[LastError];
    exit;
  end;

  if (GetField('GA_ARTICLE') = '') then SetField('GA_ARTICLE', CodeArticleUnique2(GetField('GA_CODEARTICLE'), '')); // A revoir avec les dimensions

  SetField('GA_CODEBARRE', Trim(GetField('GA_CODEBARRE')));
  if (DS <> nil) and (DS.State in [dsInsert]) then
  begin
    //************************ Gestion des Codes à barres automatiques *****************
    if (GetField('GA_CODEBARRE') = '') then
    begin
      // DCA - FQ MODE 10322 - Ajout nomenclature
      {$IFDEF MODE}
      if ((TypeArticle = 'MAR') or (TypeArticle = 'NOM')) and (GetParamsoc('SO_GCNUMCABAUTO')) then
      {$ELSE}
      if (TypeArticle = 'MAR') and (GetParamsoc('SO_GCNUMCABAUTO')) then
      {$ENDIF}
      begin
        if (GetField('GA_STATUTART') = 'GEN') then
        begin
          // Maj du qualifiant code à barres
          if RechParamCAB_Auto(GetField('GA_FOURNPRINC'), TypeCAB) then SetField('GA_QUALIFCODEBARRE', TypeCAB);
        end
        else FormateCAB(GetParamsoc('SO_GCCABARTICLE'));
      end
      else if ((TypeArticle = 'PRE') and (GetParamsoc('SO_GCPRCABAUTO')))
        then FormateCAB(GetParamsoc('SO_GCPRCABPRESTATION'));
    end;
  end;
  NewCodeCAB := GetField('GA_CODEBARRE');
  LastError := ControleCAB(GetField('GA_ARTICLE'), GetField('GA_STATUTART'), OldCodeCAB, NewCodeCAB, GetField('GA_QUALIFCODEBARRE'));
  if LastError <> 0 then
  begin
    SetLastError(LastError, 'GA_CODEBARRE');
    exit
  end;
  if NewCodeCAB <> GetField('GA_CODEBARRE') then SetField('GA_CODEBARRE', NewCodeCAB); // Maj CAB modifié par fonction de contrôle

  // DCA - FQ MODE 10793
  if (GetControl('REMPLACEMENT') <> nil) and (GetControlText('REMPLACEMENT') <> '') then
  begin
    if GetControlText('REMPLACEMENT') = GetField('GA_CODEARTICLE') then
    begin
      SetLastError(13, 'REMPLACEMENT');
      exit;
    end;

    if not existeSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + GetControlText('REMPLACEMENT') + '"') then
    begin
      SetLastError(11, 'REMPLACEMENT');
      exit;
    end;
  end;

  // DCA - FQ MODE 10793
  if (GetControl('SUBSTITUTION') <> nil) and (GetControlText('SUBSTITUTION') <> '') then
  begin
    if GetControlText('SUBSTITUTION') = GetField('GA_CODEARTICLE') then
    begin
      SetLastError(14, 'SUBSTITUTION');
      exit;
    end;

    if not existeSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + GetControlText('SUBSTITUTION') + '"') then
    begin
      SetLastError(12, 'SUBSTITUTION');
      exit;
    end;
  end;

  if (CtxMode in V_PGI.PGIContexte) then
    if not TrimZoneCritique then Exit;

  {$IFDEF GPAO}
  {$ELSE}
  if (GetField('GA_FAMILLETAXE1') = '') and (GetField('GA_TYPEARTICLE') <> 'FI') then
  begin
    SetActiveTabSheet('P_CARACTERISTIQUE');
    SetLastError(16, 'GA_FAMILLETAXE1');
    exit;
  end;
  {$ENDIF}

  if (GetField('GA_TENUESTOCK') <> 'X') then
  begin
    SetField('GA_LOT', '-');
    //SetField('GA_NUMEROSERIE','-');
    // DCA - Eviter msg en vision SAV si GA_ARTICLE n'existe pas (cas des articles financiers)
    if (GetField('GA_STATUTART') <> 'GEN') and (GetField('GA_TYPEARTICLE') <> 'PRE') and (Getfield('GA_TYPEARTICLE') <> 'CTR') and
       (Getfield('GA_TYPEARTICLE') <> 'FRA') and (GetControl('GA_ARTICLE') <> Nil) then
    begin
      QQ := OpenSQL('Select (GQ_PHYSIQUE - GQ_RESERVECLI) DISPONIBLE FROM DISPO WHERE GQ_ARTICLE="' + GetControlText('GA_ARTICLE') + '"', True);
      if not (QQ.EOF) then
        if QQ.FindField('DISPONIBLE').AsFloat > 0 then
        begin
          SetLastError(29, 'GA_TENUESTOCK');
          ferme(QQ);
          exit;
        end;
      ferme(QQ);
    end;
  end else
  begin
    SetField('GA_CONTREMARQUE', '-');
  end;

  if GetField('GA_QUALIFMARGE') = '' then SetField('GA_MARGEMINI', 0);

  if GetField('GA_TYPEARTICLE') = 'NOM' then ControlNomenc;
  {$IFDEF BTP}
  if (GetField('GA_TYPEARTICLE') = 'OUV') then ControlNomenc;
  {$ENDIF}
  // maj des champs spécifiques aux prestations
  if (GetField('GA_TYPEARTICLE') = 'PRE') or (Getfield('GA_TYPEARTICLE') = 'CTR') then OnUpdatePrestation;

  if (GetField('GA_COLLECTION') <> '') and (not LookUpValueExist(THEdit(Getcontrol('GA_COLLECTION')))) then
  begin
    SetLastError(54, 'GA_COLLECTION');
    exit;
  end;
  /// Fournisseur principal de l'article
  if GetField('GA_FOURNPRINC') <> '' then
  begin
    if BMonoFournisseur = True then
    begin
      if (not ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_NATUREAUXI="FOU" and T_TIERS="' + GetField('GA_FOURNPRINC') + '"')) then
      begin
        SetLastError(23, 'GA_FOURNPRINC');
        exit;
      end;
(*  modif brl 26/05/05
    end else
    begin
      if not ExisteSQL('Select GCA_TIERS from catalogu where GCA_TIERS="' + GetField('GA_FOURNPRINC') + '"') then
      begin
        SetLastError(32, 'GA_FOURNPRINC');
        exit;
      end;
*)
    end;
  end;
  if ControleModifFournisseur(False) = False then
  begin
    SetLastError(40, 'GA_FOURNPRINC');
    exit;
  end;

  {$IFNDEF GPAO}
  if ChampsObligatoires then exit;
  {$ELSE}
  if (ArtSpecif = 'FicheGpao') then
  begin
    { Vérification des champs obligatoires }
    ErrorFieldName := '';
    TypeInitAction := iaVerif;
    if Assigned(DS) then
      Ok := wInitFieldsbyDS(DS, 'ARTICLE', TypeInitAction, ErrorFieldName)
    else
      Ok := wInitFields(fTob, 'ARTICLE', TypeInitAction, ErrorFieldName);
    if not Ok then
    begin
      LastError := ErrChampObligatoire;
      LastErrorMsg := TraduireMemoire('Le champ') + ' : ' + wGetLibChamp(ErrorFieldName) + ' (' + ErrorFieldName + ') ' + TraduireMemoire('est obligatoire.');
      if Assigned(Ecran) then
        SetFocusControl(ErrorFieldName)
      else
        fTob.AddChampSupValeur('Error', LastErrorMsg, false);
      EXIT; ////////
    end;
  end;
  {$ENDIF}

  if ReferArtFourn <> '' then
  begin
    // Mise à jour du catalogue fournisseur, pour lui affecter l'article
    if ExisteSQL('SELECT GCA_REFERENCE FROM CATALOGU WHERE GCA_REFERENCE="' + ReferArtFourn + '" AND GCA_TIERS="' + GetField('GA_FOURNPRINC') + '"') then
      ExecuteSQL('UPDATE CATALOGU SET GCA_ARTICLE="' + GetField('GA_ARTICLE') + '" WHERE GCA_REFERENCE="' + ReferArtFourn + '" AND GCA_TIERS="' +
        GetField('GA_FOURNPRINC') + '"');
    ReferArtFourn := '';
  end;

  /////////////
  /// SI OK création des tables liées
  ////////////
  // DCA - FQ MODE 10793
  if (GetControl('REMPLACEMENT') <> nil) and (GetControlText('REMPLACEMENT') <> '') then
  begin
    st := GetControlText('TREMPLACEMENT');
    EntreCote(st, true);
    st := copy(st, 1, ChampToLength('GAL_LIBELLE')); // ENtreCote ajoute une quote si présence d'une quote dans le libellé ....
    QQ := OpenSQL('SELECT GAL_ARTICLELIE FROM ARTICLELIE WHERE GAL_TYPELIENART="REM"'
      + ' AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"'
      + ' AND GAL_RANG=1', TRUE);
    if not QQ.EOF then
    begin
      if QQ.Findfield('GAL_ARTICLELIE').AsString <> GetControlText('GA_REMPLACEMENT') then
      begin
        ExecuteSQL('UPDATE ARTICLELIE SET GAL_ARTICLELIE="' + GetControlText('GA_REMPLACEMENT')
          + '",GAL_LIBELLE="' + st + '"'
          + ' WHERE GAL_TYPELIENART="REM" AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"'
          + ' AND GAL_RANG=1');
      end;
    end else
    begin
      ExecuteSQL('INSERT INTO ARTICLELIE (GAL_TYPELIENART,GAL_ARTICLE,GAL_RANG,GAL_ARTICLELIE,GAL_LIBELLE,GAL_OBLIGATOIRE,GAL_TARIFREF) VALUES ("REM","'
        + string(GetField('GA_ARTICLE')) + '","1","' + GetControlText('GA_REMPLACEMENT')
        + '","' + st + '","-","")');
    end;
    ferme(QQ);
  end else
  begin
    ExecuteSQL('DELETE FROM ARTICLELIE WHERE GAL_TYPELIENART="REM" AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '" AND GAL_ARTICLELIE="' +
      OldRemplacement + '"');
  end;

  // DCA - FQ MODE 10793
  if (GetControl('GA_SUBSTITUTION') <> nil) and (GetControlText('GA_SUBSTITUTION') <> '') then
  begin
    st := GetControlText('TSUBSTITUTION');
    EntreCote(st, true);
    st := copy(st, 1, ChampToLength('GAL_LIBELLE'));
    QQ := OpenSQL('SELECT GAL_ARTICLELIE FROM ARTICLELIE WHERE GAL_TYPELIENART="SUB"'
      + ' AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"'
      + ' AND GAL_RANG=1', TRUE);
    if not QQ.EOF then
    begin
      if QQ.Findfield('GAL_ARTICLELIE').AsString <> GetControlText('GA_SUBSTITUTION') then
      begin
        ExecuteSQL('UPDATE ARTICLELIE SET GAL_ARTICLELIE="' + GetControlText('GA_SUBSTITUTION')
          + '",GAL_LIBELLE="' + st + '"'
          + ' WHERE GAL_TYPELIENART="SUB" AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"'
          + ' AND GAL_RANG=1');
      end;
    end else
    begin
      ExecuteSQL('INSERT INTO ARTICLELIE (GAL_TYPELIENART,GAL_ARTICLE,GAL_RANG,GAL_ARTICLELIE,GAL_LIBELLE,GAL_OBLIGATOIRE,GAL_TARIFREF) VALUES ("SUB","'
        + string(GetField('GA_ARTICLE')) + '","1","' + GetControlText('GA_SUBSTITUTION')
        + '","' + st + '","-","")');

    end;
    ferme(QQ);
  end else
  begin
    ExecuteSQL('DELETE FROM ARTICLELIE WHERE GAL_TYPELIENART="SUB" AND GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '" AND GAL_ARTICLELIE="' +
      OldSubstitution + '"');
  end;

  if (CtxMode in V_PGI.PGIContexte) and ((GetPresentation = ART_ORLI) or (GetField('GA_TYPEARTICLE') = 'FI')) then
  begin
    // DCA - FQ MODE 10793
    if (GetControl('GA2_STATART1') <> nil) and (GetControlText('GA2_STATART1') <> '') and (not LookUpValueExist(THEdit(Getcontrol('GA2_STATART1')))) then
    begin
      SetLastError(55, 'GA2_STATART1');
      exit;
    end;
    // DCA - FQ MODE 10793
    if (GetControl('GA2_STATART2') <> nil) and (GetControlText('GA2_STATART2') <> '') and (not LookUpValueExist(THEdit(Getcontrol('GA2_STATART2')))) then
    begin
      SetLastError(55, 'GA2_STATART2');
      exit;
    end;
    // DCA - Eviter pb sur articles dimensionnés (CHR)
    if TobZoneLibre <> nil then
    begin
      TobZoneLibre.GetEcran(TFfiche(Ecran), nil);
      if TobZoneLibre.IsOneModifie then
      begin
        // Mise à jour de la table ARTICLECOMPL pour les articles dimensionnés
        // DCA - FQ MODE 10506 - Gestion insertion dans ARTICLECOMPL si inexistant
        //SQL := 'select * from ARTICLECOMPL where GA2_CODEARTICLE="' + GetField('GA_CODEARTICLE') + '"';
        // Cette nouvelle requête charge les enregistrements des articles dimensionnés même si ceux-ci n'existent pas !
        SQL := 'select GA_ARTICLE as GA2_ARTICLE, GA_CODEARTICLE as GA2_CODEARTICLE, ' +
               'GA2_FAMILLENIV4, GA2_FAMILLENIV5, GA2_FAMILLENIV6, GA2_FAMILLENIV7, ' +
               'GA2_FAMILLENIV8, GA2_STATART1, GA2_STATART2, GA2_RATTACHECLI ' +
               'from ARTICLE left join ARTICLECOMPL on GA2_ARTICLE = GA_ARTICLE ' +
               'where GA_ARTICLE like "' + GetField('GA_CODEARTICLE') + '%"';
        QQ := OpenSQL(SQL, True); //False);
        TobArtCompl := TOB.Create('ARTICLECOMPL', nil, -1);
        TobArtCompl.LoadDetailDB('ARTICLECOMPL', '', '', QQ, False);
        Ferme(QQ);
        for iItem := 0 to TobArtCompl.Detail.Count - 1 do
          TOBCopyFieldValues(TobZoneLibre, TobArtCompl.Detail[iItem],
            ['GA2_FAMILLENIV4', 'GA2_FAMILLENIV5', 'GA2_FAMILLENIV6', 'GA2_FAMILLENIV7',
            'GA2_FAMILLENIV8', 'GA2_STATART1', 'GA2_STATART2', 'GA2_RATTACHECLI']);
        // Maj infos article générique
        for iItem := 4 to 8 do TobArtCompl.AddChampSup('GA2_FAMILLENIV' + IntToStr(iItem), False);
        for iItem := 1 to 2 do TobArtCompl.AddChampSup('GA2_STATART' + IntToStr(iItem), False);
        TobArtCompl.AddChampSup('GA2_ARTICLE', False);
        TobArtCompl.AddChampSup('GA2_CODEARTICLE', False);
        TOBCopyFieldValues(TobZoneLibre, TobArtCompl,
          ['GA2_FAMILLENIV4', 'GA2_FAMILLENIV5', 'GA2_FAMILLENIV6', 'GA2_FAMILLENIV7',
          'GA2_FAMILLENIV8', 'GA2_STATART1', 'GA2_STATART2', 'GA2_RATTACHECLI']);
        TobArtCompl.PutValue('GA2_ARTICLE', GetField('GA_ARTICLE'));
        TobArtCompl.PutValue('GA2_CODEARTICLE', GetField('GA_CODEARTICLE'));
        TobArtCompl.InsertOrUpdateDB(False);
        TobArtCompl.Free;
        TobZoneLibre.SetAllModifie(False);
      end;
    end;
  end;

  if TobLiensOle <> nil then with TobLiensOle do
    begin
      stArticle := GetField('GA_ARTICLE');
      for iItem := 0 to Detail.Count - 1 do Detail[iItem].PutValue('LO_IDENTIFIANT', stArticle);
      SetAllModifie(True);
      InsertOrUpdateDB(False);
      Free;
      TobLiensOle := nil;
    end;

  if (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') <> '') and
    (DimensionsArticle.TableDim <> nil) and (DimensionsArticle.TableDim.Count > 0) then
  begin
    TobArt.GetEcran(Ecran);

    if (DS <> nil) and not (DS.State in [dsInsert]) then
    begin
      // Maj des champs communs aux articles dimensionnés :
      // - parcours des champs de la fiche article,
      //     si le champ n'est pas un champ modifiable des articles dimensionnés
      //     alors report de la modif du champ pour les articles dimensionnés.
      prixUnique := Boolean(GetField('GA_PRIXUNIQUE') = 'X');
      QQ := OpenSQL('select DH_NOMCHAMP,DH_TYPECHAMP from DECHAMPS where DH_PREFIXE="GA"', True);
      SQL := '';
      while not QQ.EOF do
      begin
        nomChamp := QQ.FindField('DH_NOMCHAMP').AsString;
        if (GetControl(nomChamp) <> nil) and // Le champ est present dans la fiche
        (nomChamp <> 'GA_ARTICLE') and (nomChamp <> 'GA_CODEARTICLE') and // Pas de maj des codes article
        (nomChamp <> 'GA_BLOCNOTE') and // Pas de maj du bloc note !!
        (nomChamp <> 'GA_DATEMODIF') and (nomChamp <> 'GA_UTILISATEUR') and // Maj forcée
        (pos('DIM', nomChamp) = 0) and // Pas de maj des champs nommés '%DIM%'
        (TobArt.GetValue(nomChamp) <> TobArtLoad.Getvalue(nomChamp)) and // Maj des champs modifiés uniquement
        // DCA - FQ MODE 10813
        {Anciennes lignes
          ((TobChampModif.FindFirst(['CO_ABREGE'], [nomChamp], True) = nil) or
          (not PrixModifiable(NomChamp, PrixUnique, True))) then
        }
        (not ChampModif (NomChamp, PrixUnique, True)) then
          // Le champ n'est pas modifiable pour les articles dimensionnés ou article en prix unique
        begin
          if SQL = '' then SQL := 'update ARTICLE set ' else SQL := SQL + ',';
          ValeurNomChamp := ChampTobToSQL(NomChamp, TobArt);
          SQL := SQL + NomChamp + '=' + ValeurNomChamp;
          // DCA - FQ MODE 10761
          // Plus de maj des champs contrevaleur !
          //if ((nomChamp = 'GA_DPA') or (nomChamp = 'GA_DPR') or (nomChamp = 'GA_PVHT') or
          //  (nomChamp = 'GA_PVTTC') or (nomChamp = 'GA_PAHT') or (nomChamp = 'GA_PRHT') or
          //  (nomChamp = 'GA_PMAP') or (nomChamp = 'GA_PMRP'))
          //  then SQL := SQL + ',' + NomChamp + 'CON=' + StrFPoint(ConvertPrix(TobArt.GetValue(NomChamp)));

          // Report des modifs des champs communs article générique et articles dimensionnés
          // dans les TOB de DimensionsArticle.TableDim.Items[]
          if (DimensionsArticle.TableDim <> nil) and
            (DimensionsArticle.TOBArticleDim <> nil) and (DimensionsArticle.TOBArticleDim.FieldExists(NomChamp)) then
          begin
            for iItem := 0 to DimensionsArticle.TableDim.count - 1 do
            begin
              if DimensionsArticle.TableDim.Items[iItem] <> nil then
              begin
                ArtDim := TOB(THDimensionItem(DimensionsArticle.TableDim.Items[iItem]).data);
                if (ArtDim <> nil) then
                begin
                  ArtDimModifie := ArtDim.modifie; // svg avant PutValue
                  ArtDim.PutValue(NomChamp, ValeurNomChamp);
                  ArtDim.modifie := ArtDimModifie; // False ; // restauration : valeur inchangée
                  // La maj base est en cours en 1 seule requete pour les champs communs.
                  // La maj specifique aux dimensions est faite par DimensionsArticle.UpdateDB(TobArt) ;
                end;
              end;
            end;
          end;
        end;
        QQ.next;
      end;
      ferme(QQ);
      if SQL <> '' then
      begin
        SQL := SQL + ',GA_DATEMODIF="' + USTime(NowH);
        SQL := SQL + '",GA_UTILISATEUR="' + V_PGI.User;
        SQL := SQL + '" where GA_CODEARTICLE="' + GetControlText('GA_CODEARTICLE');
        SQL := SQL + '" and GA_STATUTART="DIM"';
        ExecuteSQL(SQL);

        // Charge les valeurs des TOB des items dans itemDim pour maj ecran
        DimensionsArticle.ChangeChampDimMul(True);
      end;
    end;

    if (DS <> nil) and (DS.State in [dsInsert]) then
    begin
      if ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_STATUTART="DIM" AND GA_CODEARTICLE="' + CodeArt + '"') then
      begin
        SetField('GA_CODEARTICLE', '');
        SetLastError(57, 'GA_CODEARTICLE');
        Exit;
      end;
    end;

    { DEPLACE DANS INSERTIONOFF POUR INSERTION DES DIMENSIONS APRES L'ARTICLE GENERIQUE
    LastError:=DimensionsArticle.UpdateDB(TobArt,TobZoneLibre,ArtError) ;
    if LastError<>0 then
    begin
       LastErrorMsg:=TraduireMemoire('Article Dimension : ')+TexteMessage[LastError] ;
       if ArtError<>'' then itemDimSetFocus( DimensionsArticle, nil, ArtError, 1, 'GA_CODEBARRE' ) ;
    end;
    }
  end;

  if (ctxMode in V_PGI.PGIContexte) and
    (LastError = 0) and
    (DS <> nil) and
    (DS.State in [dsInsert]) and
    (GetParamSoc('SO_GCPERBASETARIF') <> '') and
    (GetParamSoc('SO_GCCREERTARIFBASE') = 'TRUE') and
    (GetField('GA_TYPEARTICLE') = 'MAR') then
  begin
    TobArt.GetEcran(Ecran);
    SaisiTarifBase('nouveau');
  end;

  if (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') <> '') then
  begin
    // Si validation d'un nouvel article, code article et masque deviennent non modifiables
    if DimensionsArticle.Dim.Action = taCreat then
    begin
      SetControlEnabled('GA_CODEARTICLE', False);
      SetControlEnabled('GA_DIMMASQUE', False);
      SetControlEnabled('CBDETAIL', True);
      {$IFNDEF BTP}
      SetControlEnabled('BPARAM', True);
      {$ENDIF}
    end;
  end;

  ArticlePrin := GetField('GA_ARTICLE');
  FournPrinc := GetField('GA_FOURNPRINC');
  OldFournPrinc := FournPrinc;
  OldCodeCAB := GetField('GA_CODEBARRE');

  // En creation, rend (in)visible l'onglet dimension, le panel dimension et le checkbox prix unique
  if (DS <> nil) and (DS.State in [dsInsert]) and (GetField('GA_STATUTART') <> 'GEN') then
  begin
    // affichage onglet dimension suivant type d'article
    OngletDimension := TTAbSheet(GetControl('P_DIMENSION'));
    if OngletDimension <> nil then OngletDimension.TabVisible := (GetField('GA_STATUTART') = 'GEN');
    PanelDimension := TPANEL(GetControl('PDIMENSION'));
    if PanelDimension <> nil then PanelDimension.Visible := (GetField('GA_STATUTART') = 'DIM');
    // affichage checkbox prix unique
    if (GetControl('GA_PRIXUNIQUE') <> nil) then SetControlVisible('GA_PRIXUNIQUE', Boolean(GetField('GA_STATUTART') = 'GEN'));
  end;

  if (DS <> nil) and (DS.State in [dsInsert]) and (GetField('GA_TYPEARTICLE') = 'ARP') then
  begin
    // affichage onglet dimension suivant type d'article
    TTAbSheet(GetControl('P_DIMENSION')).TabVisible := false;
    TPANEL(GetControl('PDIMENSION')).Visible := false;
  end;

  // SPECIF AFFAIRE////////////////////////////
  if (ctxAffaire in V_PGI.PGIContexte) then
  begin
    SetControlVisible('BINSERTNOMENC', False);
    if ArtSpecif <> 'FicheAffaire' then SetControlCaption('TTYPEARTICLE', 'Fourniture');
  end;
  //
  // Modif BTP 06/02/2001
  //
  {$IFDEF BTP}
  if (DS <> nil) and (DS.State in [dsInsert]) and (GetField('GA_STATUTART') = 'UNI') and (TypeArticle = 'ARP') then
    TTabSheet(GetControl('TTARIFACH')).TabVisible := true;

  if (IsArticleSpecif(TypeArticle) = 'FicheBtp') then
  begin
    if (TypeArticle = 'MAR') then SetControlCaption('TTYPEARTICLE', 'Article')
    else if (TypeArticle = 'PRE') then SetControlCaption('TTYPEARTICLE', 'Prestation')
    else if (TypeArticle = 'NOM') or (TypeArticle = 'OUV') then
    begin
      LibTypNomenc := RechDom('BTTYPENOMENC', getfield('GA_TYPENOMENC'), FALSE);
      SetControlCaption('TTYPEARTICLE', LibTypNomenc);
    end;
  end;
  SetControlEnabled('BNOMENC1', true);
  SetControlEnabled('BNOMENC', true);
  {$IFDEF EAGLCLIENT}
  Naturepres := THEdit(GetControl('GA_NATUREPRES'));
  {$ELSE}
  Naturepres := THDBEdit(GetControl('GA_NATUREPRES'));
  {$ENDIF}
  if (Naturepres <> nil) AND (Naturepres.visible) then
  begin
    SQL := 'Select BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES="' + Getfield('GA_NATUREPRES') + '"';
    QQ := OpenSQL(SQL, TRUE);
    if QQ.EOF then
    begin
      SetLastError(56, 'GA_NATUREPRES');
      exit;
    end;
    ferme(QQ);
  end;

  {$ELSE}
  // Modif par lot
  if ModifparLot then TFFiche(ecran).BFermeClick(nil);
  {$ENDIF}

  if (ctxMode in V_PGI.PGIContexte) and (GetField('GA_STATUTART') = 'GEN')
    then AffComboEtab(True); // Gestion affichage combo établissement
  if TobArtLoad <> nil then TobArtLoad.GetEcran(Ecran);

  {$IFDEF BTP}
  if (DuplicFromMenu or (DuplicArt <> '')) and (TypeArticle = 'OUV') then
  begin
    BTPDuplicOuvrage(DuplicArt, GetField('GA_ARTICLE'));
    DuplicArt := ''
  end;
  {$ENDIF}

  {$IFDEF GPAO}
  if (DS <> nil) and (DS.State in [dsEdit]) then
  begin
    { Applique le profil aux articles associés }
    if (GetField('GA_ESTPROFIL') = 'X') then wSetAdvancedProfil(paMajArticle, DS);
    { Désactive le bouton de configuration }
    SetControlEnabled('BCFXARTICLE', False);
  end;
  if (DuplicFromMenu or (DuplicArt <> '')) then
  begin
    wDuplicWAN(DuplicArt, GetField('GA_ARTICLE'));
    DuplicArt := ''
  end;
  SetBtArtNatState;
  wGetRetour;
  {$ENDIF}

  // DEBUT CHR le 26/02/02
  if (ctxChr in V_PGI.PGIContexte) then
  begin
    pv := GetField('GA_PVTTC');
    QQ := OpenSql('Select TV_TAUXVTE from TXCPTTVA Where TV_TVAOUTPF="TX1" And ' +
      'TV_CODETAUX="' + getfield('GA_FAMILLETAXE1') + '" And TV_REGIME="FRA"', True);
    if not QQ.EOF then
    begin
      tx := QQ.Fields[0].AsFloat;
      ht := Arrondi(pv / (1.0 + (tx / 100)), 2);
      SetField('GA_PVHT', ht);
    end;
    Ferme(QQ);
    if (GetField('GA_TYPEARTICLE') = 'NOM') then
    begin
      SetControlVisible('BNOMENC', true);
    end;
  end;
  // FIN CHR le 26/02/02

  // Contrôle cohérence des familles hiérarchiques
  // Correction FQ;032;12594
  if  (IsMasterOnShare('GCFAMILLENIV1') and GetParamSoc('SO_GCFAMHIERARCHIQUE')) then
  // --
//  if GetParamSoc('SO_GCFAMHIERARCHIQUE') = True then
  begin
  	if (copy(GetField('GA_TYPEARTICLE'),1,2)='PA') then
    begin
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BP3" AND CC_CODE = "' + getfield('GA_FAMILLENIV3') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        getfield('GA_FAMILLENIV2') + '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV3', '');
      Ferme(QQ);
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BP2" AND CC_CODE = "' + getfield('GA_FAMILLENIV2') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV2', '');
      Ferme(QQ);
    end else
  	if (GetField('GA_TYPEARTICLE')='OUV') then
    begin
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BO3" AND CC_CODE = "' + getfield('GA_FAMILLENIV3') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        getfield('GA_FAMILLENIV2') + '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV3', '');
      Ferme(QQ);
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "BO2" AND CC_CODE = "' + getfield('GA_FAMILLENIV2') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV2', '');
      Ferme(QQ);
    end else
    begin
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "FN3" AND CC_CODE = "' + getfield('GA_FAMILLENIV3') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        getfield('GA_FAMILLENIV2') + '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV3', '');
      Ferme(QQ);
      St := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE = "FN2" AND CC_CODE = "' + getfield('GA_FAMILLENIV2') + '" AND CC_LIBRE like "%' + getfield('GA_FAMILLENIV1') +
        '%"';
      QQ := OpenSql(St, True);
      if QQ.Eof then SetField('GA_FAMILLENIV2', '');
      Ferme(QQ);
    end;
  end;

  if ArtSpecif = 'FicheModePre' then exit;

{$IFDEF BTP}
  // Correction fiche qualité 10166
  if GetField('GA_TYPEARTICLE') = 'OUV' Then
    CreationEnteteOuv (GetField('GA_ARTICLE'),GetField('GA_LIBELLE'),GetField('GA_DOMAINE'));

  if GetField('GA_TYPEARTICLE') = 'ARP' then
  begin
    ReajusteTOBDetail;
    ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_NOMENCLATURE="' + GetField('GA_ARTICLE') + '"');
    ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="' + GetField('GA_ARTICLE') + '"');
    if GetField('GA_STATUTART') <> 'GEN' then
    begin
      TOBEDetail.InsertDb(nil);
      TOBDetail.InsertDb(nil);
    end;
  end;
  //
  if (TypeArticle = 'ARP') or (TypeArticle='MAR') or (TypeArticle ='PRE') then
  begin
  	EcritTobArticleCompl;
  end;
  if copy(TypeArticle,1,2) = 'PA' then
  begin
  	EcritTobArticleCompParc;
  end;
  //
  // Gestion des métrés
  //FV1 - 08/07/2016
  //Validation du fichier métré
  if MetreArticle <> nil then
  Begin
    MetreArticle.FermerMetreXLS;
    MetreArticle.CopieLocaltoServeur(RepPoste, RepServeur);
  end;
  //
  //Gestion de la duplication d'un Article et de son métré
  if DuplicArt <> '' then
  begin
    DuplicationArticleMetre(TypeArticle, Duplicart, GetField('GA_ARTICLE'));
  end;
  //
  {$ENDIF}

  // Contournement bug création+modif => repositionnement sur 1er élément.
  //Non en line
  if (DS <> nil) and (DS.State in [dsInsert]) then
     DerniereCreate := GetField('GA_ARTICLE')
  else if (DerniereCreate = GetField('GA_ARTICLE')) then
      Ecran.Close; // le bug arrive on se casse !!!

  //JD Création article à partir d'une photo
  if CreatArticlePhoto then CreationPhotoLiensOLE;

  // DBR : pour ne pas pouvoir modifier GA_CODEARTICLE après validation si enreg inexistant - Fiche 10549
  if (DS <> nil) and (DS.State in [dsInsert]) then
  begin
    if ExisteSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + GetControlText ('GA_CODEARTICLE') + '"') then
    begin
      SetLastError (59, 'GA_CODEARTICLE');
      exit;
    end else
    begin
      SetControlEnabled ('GA_CODEARTICLE', False);
    end;
  end;

end;

procedure TOM_Article.CreationPhotoLiensOLE;
var TOBLIENSOLE: TOB;
begin
  with TOBInfoArt do
  begin
    TOBLIENSOLE := AttacheLaPhoto(nil, GetField('GA_ARTICLE'), GetValue('PathPhoto'), GetValue('NamePhoto'), GetValue('Extension'));
    TOBLIENSOLE.InsertDB(nil);
  end;
end;

// Fonction générique pour champs HT ou champs TTC

procedure TOM_Article.OnClickCalcPrix(ga_calcprix, ga_coefcalc, ga_calcauto, ga_pv, t_stockpv: string);
{$IFDEF BTP}
var IsNomenc, PaVisi: boolean;
  {$ENDIF}
begin
  {$IFDEF BTP}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM') or (GetField('GA_TYPEARTICLE') = 'OUV');
  PAVisi := (not Isnomenc) and (not (copy(ecran.name, 1, 12) = 'BTPRESTATION'));
  {$ENDIF}
  if GetField(ga_calcprix) = 'AUC' then
  begin
    {$IFDEF BTP}
    SetField(ga_coefcalc, 0);
    {$ENDIF}
    SetControlEnabled(ga_coefcalc, False);
    SetControlEnabled(ga_calcauto, False);
    SetControlText(ga_calcauto, '-');
    if not (ctxMode in V_PGI.PGIContexte) then
    begin
      SetControlProperty(ga_pv, 'ReadOnly', False);
      SetControlProperty(ga_pv, 'TabStop', True);
    end;
    SetControlVisible(ga_pv, True);
    SetControlVisible(t_stockpv, False);
  end
  else if copy(GetField(ga_calcprix), 1, 1) = 'S' then
  begin
    SetControlVisible(ga_pv, False);
    SetControlVisible(t_stockpv, True);
    //  BBI, fiche correction 10410
    //  SetControlEnabled(ga_coefcalc,False) ;
    SetControlText(ga_calcauto,'X');
    SetControlEnabled(ga_calcauto, False);
    //  BBI, fin fiche correction 10410
  end
  else
  begin
    SetControlEnabled(ga_coefcalc, True);
    SetControlEnabled(ga_calcauto, True);
    if not (ctxMode in V_PGI.PGIContexte) then
    begin
      SetControlProperty(ga_pv, 'ReadOnly', True);
      SetControlProperty(ga_pv, 'TabStop', False);
    end;
    SetControlVisible(ga_pv, True);
    SetControlVisible(t_stockpv, False);
  end;
  {$IFDEF BTP}
  if (GetControlText('GA_CALCPRIXHT') <> 'FOU') and (GetControlText('GA_CALCPRIXPR') <> 'FOU') and (GetControlText('GA_CALCPRIXTTC') <> 'FOU') then
  begin
    if not IsLine then SetControlvisible('GA_DPA', Pavisi);
    if not IsLine then SetControlVisible('PAFOUPRINC', false);
    if not IsLine then SetControlvisible('TGA_DPA', Pavisi);
    if not IsLine then SetControlVisible('TPAFOUPRINC', false);
  end;
  {$ENDIF}
end;

{$IFNDEF GPAO}
function TOM_Article.ChampsObligatoires: boolean;
var NomChamp: string;
begin
  result := false;
  NomChamp := '';
  {$IFNDEF CCS3}
  TypeArticle := GetField('GA_TYPEARTICLE');
  //  BBI correction fiche 10360
  if TypeArticle = 'NOM' then
    TypeArticle := TypeArticle + GetField('GA_TYPENOMENC');
  //  BBI fin correction fiche 10360
  NomChamp := VerifierChampsObligatoires(Ecran, TypeArticle);
  if NomChamp <> '' then
  begin
    NomChamp := ReadTokenSt(NomChamp);
    SetFocusControl(NomChamp);
    LastError := 52;
    LastErrorMsg := TexteMessage[LastError] + champToLibelle(NomChamp);
    result := True;
  end;
  {$ENDIF}
end;
{$ENDIF}

procedure TOM_Article.OnChangeDetail;
begin
  DimensionsArticle.HideUnUsed := boolean(GetControlText('CBDETAIL') <> 'X');
  if DimensionsArticle.bChargtOptimise and (not DimensionsArticle.HideUnUsed) and GetParamSoc('SO_CHARGEDIMDEGRADE') then
  begin
    DimensionsArticle.bChargtOptimise := False;
    DimensionsArticle.LoadDimensions(False);
    OnChangeAfficheChampDimMul(DimensionsArticle, DimensionsArticle.natureDoc);
  end;
  DimensionsArticle.Dim.HideUnUsed := DimensionsArticle.HideUnUsed;
end;

procedure TOM_Article.OnNavigClick;
begin
  NavigClick := True;
end;

procedure TOM_Article.OnClickPrixUnique;
var bAbandon: boolean;
begin
  if not (DS.State in [dsInsert, dsEdit]) then exit;
  if CreatArticlePhoto then exit;
  if (GetField('GA_STATUTART') = 'GEN') and (not DuplicFromMenu) and (not DuplicFromFiche) then
  begin
    DimensionsArticle.PrixUnique := Boolean(GetField('GA_PRIXUNIQUE') = 'X');
    if (DimensionsArticle.PrixUnique) then
    begin
      bAbandon := True;
      // Contrôle non existence dans le dispo
      if ExisteSQL('select GQ_ARTICLE from DISPO left join ARTICLE on GA_ARTICLE=GQ_ARTICLE where GA_STATUTART="DIM" and GA_CODEARTICLE="' +
        GetField('GA_CODEARTICLE') + '"')
        then PGIBox(TexteMessage[42], Ecran.Caption)
        // Contrôle non existence dans les tarifs
      else if ExisteSQL('select GF_ARTICLE from TARIF left join ARTICLE on GA_ARTICLE=GF_ARTICLE where GA_STATUTART="DIM" and GA_CODEARTICLE="' +
        GetField('GA_CODEARTICLE') + '"')
        then PGIBox(TexteMessage[43], Ecran.Caption)
        // Message de confirmation : oui -> maj prix des fiches articles dimensionnés
      else if PGIAsk(TexteMessage[44], TexteMessage[45]) = mrYes then
      begin
        TobArt.GetEcran(ECRAN);
        DimensionsArticle.MajPrixUniques(TobArt);
        DimensionsArticle.ChangeChampDimMul(True);
        DimensionsArticle.RefreshGrid(DimensionsArticle.HideUnUsed);
        bAbandon := False;
      end;
      if bAbandon then
      begin
        SetField('GA_PRIXUNIQUE', '-');
        DimensionsArticle.PrixUnique := False;
      end;
    end else
    begin
      DimensionsArticle.MiseEnForme;
      DimensionsArticle.RefreshGrid(DimensionsArticle.HideUnUsed);
    end;
  end;
end;

procedure TOM_Article.OnUpdatePrix(stPrix: string);
var prixAffiche: boolean;
  iChamp: integer;
begin
  if not (DS.State in [dsInsert, dsEdit]) then exit;
  if (GetField('GA_STATUTART') = 'GEN') then
  begin
    DimensionsArticle.PrixUnique := Boolean(GetField('GA_PRIXUNIQUE') = 'X');
    if (DimensionsArticle.PrixUnique) then
    begin
      prixAffiche := False;
      iChamp := 1;
      while (iChamp < MaxDimChamp) and (not prixAffiche) do
      begin
        if (DimensionsArticle.NewDimChamp[iChamp] = stPrix) then prixAffiche := True;
        inc(iChamp);
      end;
      // Report automatique du prix modifié sur les articles dimensionnés (Tob)
      TobArt.GetEcran(ECRAN);
      DimensionsArticle.MajPrixUniques(TobArt);
      if (prixAffiche) then
      begin
        // Report du prix modifié (Tob) dans le THDim + réaffichage
        DimensionsArticle.ChangeChampDimMul(True);
        DimensionsArticle.RefreshGrid(DimensionsArticle.HideUnUsed);
      end;
    end;
  end;
end;

procedure TOM_Article.VerifModifDimension;
begin
  if TFFiche(Ecran).fTypeAction = taConsult then Exit;
  if (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') <> '') then
  begin
    if DimensionsArticle.IsModified then
      if not (DS.State in [dsInsert, dsEdit])
        then DS.edit // pour passer DS.state en mode dsEdit
      else if DS.State = dsInsert then SetControlEnabled('GA_DIMMASQUE', False);
  end;
end;

procedure TOM_Article.SaisiTarifBase(Etat: string);
var creer: Boolean;
begin
  if (GetParamSoc('SO_GCPERBASETARIF') = '') or not (ExisteSQL('Select GFT_CODETYPE from TARIFTYPMODE')) then exit;
  creer := False;
  if TobArt = nil then
  begin
    creer := true;
    TobArt := TOB.Create('ARTICLE', nil, -1);
    TobArt.GetEcran(Ecran);
  end;
  TheTob := TOBArt;
  AGLLanceFiche('MBO', 'TARIFSAISIBASE', '', '', 'CodeArticle=' + GetField('GA_ARTICLE') + ';Etat=' + Etat);
  TheTob := nil;
  if creer then
  begin
    tobArt.Free;
    tobArt := nil;
  end;
end;

procedure TOM_Article.ArticleVoir(Sender: TObject);
var QMasque: TQuery;
  ArticleDim: string;
  ItemDim: THDimensionItem;
  ArtDim: TOB;
  iChamp: integer;
  NomChamp, Acces: string;
  ToRefresh: boolean;
begin
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then exit;
  if ItemDim.ToCreate then
  begin
    PGIInfo(TexteMessage[53], Ecran.Caption);
    exit
  end;

  if ctxMode in V_PGI.PGIContexte then
  begin
    QMasque := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + GetField('GA_DIMMASQUE') + '" and GDM_TYPEMASQUE="' + TypeMasque + '"', TRUE);
    if (QMasque.EOF) and (TypeMasque <> VH_GC.BOTypeMasque_Defaut) then
      QMasque := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + GetField('GA_DIMMASQUE') + '" and GDM_TYPEMASQUE="' + VH_GC.BOTypeMasque_Defaut + '"',
      TRUE);
  end
  else QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + GetField('GA_DIMMASQUE') + '"', TRUE);

  if not QMasque.EOF then
  begin
    ArticleDim := CodeArticleUnique(GetField('GA_CODEARTICLE'), AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim), AffecteDim(QMasque, 3,
      ItemDim), AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim));
    ferme(Qmasque);
    {$IFDEF EAGLCLIENT}
    {AFAIREEAGL}
    {$ELSE}
    if Not (ctxMode in V_PGI.PGIContexte) then
    begin
      if TheMulq = nil then TheMulQ := TQuery.Create(Application); // pour passer la clause where à la fiche
      TheMulQ.SQL.Clear;
      TheMulQ.SQL.Add('SELECT * From Article where (GA_STATUTART="DIM") AND GA_CODEARTICLE="' + GetField('GA_CODEARTICLE') + '"');
      ChangeSQL(TheMulQ);
    end;
    {$ENDIF}
    TheTob := TobChampModif;
    V_PGI.FormCenter := False;
    if Consultation then Acces := 'ACTION=CONSULTATION' else Acces := 'ACTION=MODIFICATION';
    if (ctxMode in V_PGI.PGIContexte)
      then
        AGLLanceFiche('GC', 'GCARTICLEDIM', '', ArticleDim, Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
        GetField('GA_PRIXUNIQUE')+';MONOFICHE')
      // MOdif BTP
      {$IFDEF BTP}
    else if consultation then
      V_PGI.DispatchTT(7, taconsult, ArticleDim, '', Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
        GetField('GA_PRIXUNIQUE'))
    else
      V_PGI.DispatchTT(7, tamodif, ArticleDim, '', Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
        GetField('GA_PRIXUNIQUE'));
    {$ELSE}
      {$IFDEF GPAO}
    else if consultation then
      V_PGI.DispatchTT(7, taconsult, ArticleDim, '', Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
        GetField('GA_PRIXUNIQUE'))
    else
      V_PGI.DispatchTT(7, tamodif, ArticleDim, '', Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
        GetField('GA_PRIXUNIQUE'));
    {$ELSE}
    else
      AGLLanceFiche('GC', 'GCARTICLE', '', ArticleDim, Acces + ';TOP=' + IntToStr(Ecran.Top + 20) + ';LEFT=' + IntToStr(Ecran.Left + 20) + ';FICHDIM=' +
      GetField('GA_PRIXUNIQUE'));
    {$ENDIF}
    {$ENDIF}

    // En retour TheTob pointe sur la Tob de l'article dimensionné -> Maj de l'item de THDim.
    if TheTob <> nil then
    begin
      ToRefresh := False;
      ArtDim := TOB(ItemDim.data);
      if ArtDim <> nil then for iChamp := 1 to MaxDimChamp do
        begin
          NomChamp := DimensionsArticle.NewDimChamp[iChamp];
          if (NomChamp <> '') and (TheTOB.isFieldModified(NomChamp)) then
          begin
            ArtDim.PutValue(NomChamp, TheTob.GetValue(NomChamp));
            ItemDim.Valeur[iChamp] := TheTob.GetValue(NomChamp);
            ToRefresh := True;
          end;
        end;
      TheTob.free;
      TheTob := nil;
      if ToRefresh then DimensionsArticle.refresh(ItemDim);
    end
    else
    begin // Rechargement complet de THDim
      ArticleDimFocus := TOB(ItemDim.data).GetValue('GA_ARTICLE');
      initDimensionsArticle;
      ItemDimSetFocus(DimensionsArticle, nil, ArticleDimFocus);
    end;
  end
  else ferme(Qmasque);
end;

procedure TOM_Article.ArticleActionDim(ItemDim: THDimensionItem);
var NomChamp: string;
  iChamp: integer;
  ToRefresh: boolean;
  ArtDim: TOB;
begin
  if ItemDim = nil then exit;
  ArtDim := TOB(ItemDim.data);
  if ArtDim <> nil then
  begin
    ToRefresh := False;
    for iChamp := 1 to MaxDimChamp do
    begin
      NomChamp := DimensionsArticle.NewDimChamp[iChamp];
      // Traitement des champs de la fiche article uniquement.
      if (copy(NomChamp, 1, 3) = 'GA_') and (NomChamp <> 'GA_CODEBARRE') and
        (GetField(NomChamp) <> ItemDim.Valeur[iChamp]) then
      begin
        ItemDim.Valeur[iChamp] := GetField(NomChamp);
        ArtDim.PutValue(NomChamp, ItemDim.valeur[iChamp]);
        // Correction bug maj par click droit dans la fiche article
        ArtDim.PutValue('_MAJARTICLE', True) ;
        ToRefresh := True;
      end;
    end;
    if ToRefresh then DimensionsArticle.Refresh(ItemDim);
  end;
end;

procedure TOM_Article.CreateItem(ItemDim: THDimensionItem);
var TobArtGen: TOB;
  TomArt: TOM;
begin
  if ItemDim = nil then exit;
  TobArtGen := TOB.Create('ARTICLE', nil, -1);
  TomArt := CreateTOM('ARTICLE', nil, False, False);
  TomArt.InitTOB(TobArtGen);
  TobArtGen.GetEcran(ECRAN);
  if TomArt <> nil then TomArt.Free;
  DimensionsArticle.CreateItemDim(ItemDim, TobArtGen);
  if TobArtGen <> nil then TobArtGen.Free;
end;

procedure TOM_Article.DeleteItem(ItemDim: THDimensionItem);
var Article: string;
//    TOBTT : TOB;
begin
  if ItemDim = nil then exit;
//  TOBTT := TOB(ItemDim.data);
  Article := TOB(ItemDim.data).GetValue('GA_ARTICLE');
  // Si Suppression d'un item a créer, on le supprime de la liste tout de suite.
  // DCA - FQ MODE 10682
  ItemDim.ToAbort := not DimensionsArticle.DeleteItemDim(ItemDim, Article);
end;

procedure TOM_Article.OnChangeItem(Sender: TObject);
var ItemDim: THDimensionItem;
  iChamp, iContreV: Integer;
  NomChamp, NomContreV: string;
begin
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then exit;
  for iChamp := 1 to MaxDimChamp do
  begin
    NomChamp := DimensionsArticle.OldDimChamp[iChamp];
    if NomChamp = 'GA_FAMILLETAXE1' then
    begin
      ItemDim.Valeur[iChamp] := UpperCase(ItemDim.Valeur[iChamp]);
      if not ExisteSQL('Select CC_CODE from CHOIXCOD where CC_TYPE="TX1" and CC_CODE="' + ItemDim.Valeur[iChamp] + '"') and (ItemDim.Valeur[iChamp] <> '') then
      begin
        DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
        PGIInfo(TexteMessage[46], Ecran.Caption);
      end;
    end
    else if NomChamp = 'GA_FAMILLETAXE2' then
    begin
      ItemDim.Valeur[iChamp] := UpperCase(ItemDim.Valeur[iChamp]);
      if not ExisteSQL('Select CC_CODE from CHOIXCOD where CC_TYPE="TX2" and CC_CODE="' + ItemDim.Valeur[iChamp] + '"') and (ItemDim.Valeur[iChamp] <> '') then
      begin
        DimensionsArticle.Dim.FocusDim(ItemDim, iChamp);
        PGIInfo(TexteMessage[47], Ecran.Caption);
      end;
    end
    else if ((nomChamp = 'GA_DPA') or (nomChamp = 'GA_DPR') or (nomChamp = 'GA_PVHT') or
      (nomChamp = 'GA_PVTTC') or (nomChamp = 'GA_PAHT') or (nomChamp = 'GA_PRHT') or
      (nomChamp = 'GA_PMAP') or (nomChamp = 'GA_PMRP')) then
    begin
      // Recherche de l'éventuel prix de contrevaleur
      NomContreV := NomChamp + 'CON';
      iContreV := 1;
      while (DimensionsArticle.NewDimChamp[iContreV] <> NomContreV) and (iContreV < 10) do inc(iContreV);
      // Prix contrevaleur également affiché -> maj
      if DimensionsArticle.NewDimChamp[iContreV] = NomContreV
        then ItemDim.Valeur[iContreV] := ConvertPrix(ItemDim.Valeur[iChamp]);
    end;
  end;
end;

procedure TOM_Article.AjusterTailleTHDim;
var CoordDernCol, CoordDernLign: TRect;
  GridDim: THGrid;
  ValOnglet, MaxHeight, MaxWidth, GridTop, GridLeft: Integer;
begin
  MaxHeight := 313;
  MaxWidth := 616;
  GridTop := THDimension(GetControl('FDIM')).Top;
  GridLeft := THDimension(GetControl('FDIM')).Left;
  // Maximisation de FDIM pour que VisibleColCount et VisibleRowCount fournissent les dernières cellules 'visibles'
  SetControlProperty('FDIM', 'Width', 1000);
  SetControlProperty('FDIM', 'Height', 1000);

  GridDim := DimensionsArticle.Dim.GridDim;
  if DimensionsArticle.Dim.DimOngl = nil then ValOnglet := 22 else ValOnglet := 4;
  CoordDernCol := GridDim.CellRect(GridDim.VisibleColCount + GridDim.FixedCols, 1);
  CoordDernLign := GridDim.CellRect(2, GridDim.VisibleRowCount + GridDim.FixedRows);
  GridDim.ScrollBars := ssNone;
  // Ajustement taille en fonction des dernières cellules (non prise en compte taille éventuel ascenseur)
  if CoordDernLign.Bottom + GridTop > MaxHeight then
  begin
    SetControlProperty('FDIM', 'Height', MaxHeight);
    GridDim.ScrollBars := ssVertical;
  end
  else SetControlProperty('FDIM', 'Height', CoordDernLign.Bottom + GridTop - ValOnglet);
  if CoordDernCol.Right + GridLeft > MaxWidth then
  begin
    SetControlProperty('FDIM', 'Width', MaxWidth);
    if GridDim.ScrollBars = ssVertical then GridDim.ScrollBars := ssBoth else GridDim.ScrollBars := ssHorizontal;
  end
  else SetControlProperty('FDIM', 'Width', CoordDernCol.Right + GridLeft + 8);
  // Ajustement taille si ajout d'un ascenseur
  if (GridDim.ScrollBars = ssVertical) or (GridDim.ScrollBars = ssBoth) then
  begin
    if CoordDernCol.Right + GridLeft + 16 > MaxWidth
      then SetControlProperty('FDIM', 'Width', MaxWidth)
    else SetControlProperty('FDIM', 'Width', CoordDernCol.Right + GridLeft + 8 + 16);
  end;
  if (GridDim.ScrollBars = ssHorizontal) or (GridDim.ScrollBars = ssBoth) then
  begin
    if CoordDernLign.Bottom + GridTop + 16 > MaxHeight
      then SetControlProperty('FDIM', 'Height', MaxHeight)
    else SetControlProperty('FDIM', 'Height', CoordDernLign.Bottom + GridTop - ValOnglet + 16);
  end;
end;

procedure TOM_Article.OnUpdateXXChamp(FieldName: string);
begin
  if FieldName = 'XXTYPEMASQUE' then
  begin
    TypeMasque := GetControlText(FieldName);
    OnLoadRecord();
  end
  else if FieldName = '_ETABLISS' then
  begin
    Depot := THValComboBox(GetControl(FieldName)).Value;
    OnLoadRecord();
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paul Chapuis
Créé le ...... : 19/06/2000
Modifié le ... :   /  /
Description .. : Affichage photo sur une fiche
Mots clefs ... : PHOTO;IMAGE;LIENSOLE
*****************************************************************}

procedure TOM_Article.AffichePhoto;
var CC: TImage;
  IsJpeg: boolean;
  //{$IFDEF EAGLCLIENT}
  QQ: TQuery;
  SQL: string;
  //{$ELSE}
  //    TLIENSOLE : THTable;
  //    StQualif : string;
  //{$ENDIF}
begin
  SetControlVisible('P_PHOTO', false);
  CC := TImage(GetControl('LAPHOTO'));
  if CC = nil then exit;
  if VH_GC.GCPHOTOFICHE = '' then exit;
  //{$IFDEF EAGLCLIENT}	{ GPAO V1.0 }
  SQL := 'SELECT * from LIENSOLE where LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="' + getField('GA_ARTICLE') + '"';
  SQL := SQL + ' AND (LO_QUALIFIANTBLOB="PHO" OR  LO_QUALIFIANTBLOB="PHJ" OR  LO_QUALIFIANTBLOB="VIJ") AND LO_EMPLOIBLOB="' + VH_GC.GCPHOTOFICHE + '"';
  QQ := OpenSQL(SQL, true);
  if not QQ.EOF then
  begin
    IsJpeg := ((QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'PHJ') or (QQ.Findfield('LO_QUALIFIANTBLOB').asString = 'VIJ'));
    LoadBitMapFromChamp(QQ, 'LO_OBJET', CC, IsJpeg);
    if CC.Picture <> nil then SetControlVisible('P_PHOTO', True);
  end;
  Ferme(QQ);
  (*
  {$ELSE}
    //JD
    TLIENSOLE := THTable.Create(nil);
    try
      with TLIENSOLE do
        begin
        DatabaseName := DBSOC.DatabaseName;
        Tablename:= 'LIENSOLE'; Open;
        if Locate('LO_TABLEBLOB;LO_IDENTIFIANT;LO_EMPLOIBLOB',VarArrayOf(['GA',getField('GA_ARTICLE'),VH_GC.GCPHOTOFICHE]),[loCaseInsensitive]) then
          begin
          StQualif := Findfield('LO_QUALIFIANTBLOB').asString;
          if (StQualif='PHO') OR (StQualif='PHJ') OR (StQualif='VIJ') then
            begin
            IsJpeg := ((StQualif='PHJ') or (StQualif='VIJ'))  ;
            LoadBitMapFromChamp(TLIENSOLE,'LO_OBJET',CC,IsJpeg) ;
            if CC.Picture<>Nil then SetControlVisible('P_PHOTO',True);
            end;
          end;
        Close;
        end;
    Finally
      TLIENSOLE.Free;
      end;
  {$ENDIF} *)
end;

//// Affiche les dimensions en fonction du parametrage /////////////////

procedure TOM_Article.AffecteDimension;
var TLIB, TLIB2: THLabel;
  CHPS, CHPS2: THEdit;
  i_ind, i_Dim: integer;
  b_Dim: Boolean;
  PanelDimension: TPanel;
  LibCodeDim, GrilleDim, CodeDim: string;
begin
  i_Dim := 1;
  b_Dim := False;
  for i_ind := 1 to MaxDimension do
  begin
    TLIB := THLabel(Ecran.FindComponent('GRILLEDIM' + IntToStr(i_Dim)));
    CHPS := THEdit(Ecran.FindComponent('CODEDIM' + IntToStr(i_Dim)));
    TLIB2 := THLabel(Ecran.FindComponent('GRILLEDIM' + IntToStr(i_ind)));
    CHPS2 := THEdit(Ecran.FindComponent('CODEDIM' + IntToStr(i_ind)));
    TLIB2.Caption := '';
    CHPS2.Text := '';
    CHPS2.Visible := False;
    GrilleDim := GetField('GA_GRILLEDIM' + IntToStr(i_ind));
    CodeDim := GetField('GA_CODEDIM' + IntToStr(i_ind));
    if GrilleDim <> '' then
    begin
      TLIB.Caption := RechDom('GCGRILLEDIM' + IntToStr(i_ind), GrilleDim, FALSE);
      LibCodeDim := GCGetCodeDim(GrilleDim, CodeDim, i_ind);
      if LibCodeDim <> '' then
      begin
        CHPS.Text := LibCodeDim;
        CHPS.Visible := True;
        b_Dim := True;
        Inc(i_Dim);
      end else
      begin
        TLIB.Caption := '';
        CHPS.Text := '';
        CHPS.Visible := False;
      end;
    end;
  end;
  PanelDimension := TPANEL(GetControl('PDIMENSION'));
  if PanelDimension <> nil then
  begin
    if not b_Dim then PanelDimension.Visible := False else PanelDimension.Visible := True;
  end;
end;

//////////////////////////////////////////////////////////////////////

procedure TOM_Article.OnLoadRecord;
var
  // Modif BTP
  OngletCaracteristique : TTabSheet;
  OngletTarif           : TTabSheet;
  OngletDimension       : TTAbSheet;
  //
  PanelCaracteristique  : TPanel;
  PanelTarif            : TPanel;
  PanelDimension        : TPanel;
  //
  NbDeclinaison : integer;
  I             : Integer;
  x             : Integer;

  IsMacro       : Boolean;
  IsNomenc      : Boolean;
  IsAucOuStock  : Boolean;
  IsParc        : Boolean;
  IsNomenParc   : Boolean;
  IsEnabled     : boolean;
  //
  stTemp        : string;
  ArtSpecif     : string;
  Requete       : string;
  CodeSSTarif   : string;
  LibTypNomenc  : string;
  //
  QQ            : TQuery;
begin
  inherited;
  if CodeSousTarif <> nil then CodeSSTarif := CodeSousTarif.value else CodeSStarif := '';
  if not (DS.State in [dsInsert]) then
  begin
    bCalculPrixDuplicHT := True; // DBR pour pas avoir le message 37 la premiere fois
    bCalculPrixDuplicTTC := True;
  end;
  if ECRAN.name = 'GCARTICLEDIM' then // Fiche article dimensionné
  begin
    // Décalage fenêtre de l'article dimensionné
    Ecran.Top := Top;
    Ecran.Left := Left;
    InitCaption(Ecran, GetField('GA_CODEARTICLE'), GetField('GA_LIBELLE'));
    AffecteDimension;

    if TobArt <> nil then
    begin
      TobArt.free;
      TobArt := nil;
    end;
    TobArt := TOB.Create('ARTICLE', nil, -1);
    TobArt.GetEcran(Ecran);
    TobArt.SetAllModifie(False);

    // Mise en forme des captions de l'écran
    if (GetField('GA_QUALIFCODEBARRE') <> '')
      then stTemp := RechDom('YYYBARTYPE', GetField('GA_QUALIFCODEBARRE'), False)
    else stTemp := '';
    SetControlProperty('TGA_QUALIFCODEBARRE', 'Caption', stTemp);
    if (GetField('GA_QUALIFPOIDS') <> '')
      then stTemp := ' (' + GetField('GA_QUALIFPOIDS') + ')'
    else stTemp := '';
    SetControlProperty('TGA_POIDSBRUT', 'Caption', 'Poids brut' + stTemp);
    SetControlProperty('TGA_POIDSNET', 'Caption', 'Poids net' + stTemp);
    SetControlProperty('TGA_POIDSDOUA', 'Caption', 'Poids douanier' + stTemp);
    OldCodeCAB := GetField('GA_CODEBARRE');

    exit;
  end; // if ECRAN.name='GCARTICLEDIM' then

  // Modif BTP
  if (copy(ecran.name, 1, 12) = 'BTPRESTATION') then
  	begin
    // Modif LS pour article de type frais //
    if TypeArticle = 'PRE' Then
    begin
      THLabel(GetControl('TGA_LIBNATURE')).caption := '';
      if getfield('GA_NATUREPRES') <> '' then
      begin
        Requete := 'Select BNP_LIBELLE FROM NATUREPREST WHERE BNP_NATUREPRES="' + Getfield('GA_NATUREPRES') + '"';
        QQ := OpenSQL(requete, TRUE);
        if not QQ.EOF then THLabel(GetControl('TGA_LIBNATURE')).caption := QQ.findfield('BNP_LIBELLE').AsString;
        ferme(QQ);
        //QQ := nil; inutile
      end;
    end
    else if TypeArticle = 'FRA' then
    	begin
      SetCOntrolVisible ('TGA_NATUREPRES',false);
      SetCOntrolVisible ('GA_NATUREPRES',false);
      SetCOntrolVisible ('BPARAMNATURE',false);
      SetCOntrolVisible ('TGA_LIBNATURE',false);
//MODIF BRL 10/07/06      SetCOntrolVisible ('GB_COMPTA',false);
      SetCOntrolVisible ('G_MARGE',false);
      SetCOntrolVisible ('BPROFIL',false);
      SetCOntrolVisible ('TGA_TARIFARTICLE',false);
      SetCOntrolVisible ('GA_TARIFARTICLE',false);
      SetCOntrolVisible ('GA_ESCOMPTABLE',false);
      SetCOntrolVisible ('GA_REMISEPIED',false);
      SetCOntrolVisible ('GA_REMISELIGNE',false);
      SetCOntrolVisible ('GA_COMMISSIONNABLE',false);
    	end;
    // fin modif ls
  end;

{$IFDEF BTP}
  if (copy(ecran.name, 1, 9) = 'BTARTICLE') then
  begin
    if (GetField('GA_STATUTART') = 'GEN') then TTabSheet(GetControl('TTARIFACH')).TabVisible := false;

    // gestion sous famille tarifaire
    if (TypeArticle = 'ARP') or (TypeArticle='MAR') or (TypeARticle ='PRE') then
    begin
    	LoadTobArticlecompl;
    end;


    if TypeArticle = 'ARP' then
    begin
      // chargement des tobs + recalcul du paht
      if (TFFiche(Ecran).fTypeAction <> TaCreat) and (not DuplicFromMenu) then
        GetDetailPrixPose(GetField('GA_ARTICLE'), TOBPrestation, TOBEdetail, TOBDetail);
    end;


    // Recup des elements de tarification fournisseur
    if not DuplicFromMenu then
    begin
      GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'),
        GetField('GA_TARIFARTICLE'), CodeSSTarif, TOBcatalog, TOBTarif);
    end;
    ExistCatalog := (TOBCatalog.GetString('GCA_ARTICLE')<>'');
    //
    SetInfosPrixTarif;
    // Remise a jour des infos si besoin est
    if TypeArticle = 'ARP' then
    begin
      SetDetailPrixPose(GetField('GA_ARTICLE'), GetField('GA_LIBELLE'), -1, TOBPrestation);
      AffichageInfosPrixPose;
    end;
    //
    SetControlText('GCA_PAUA',StrS(0,V_PGI.okdecP) );
    //
    SetInfosPrixTarif;
    DefiniAffichage;
      end;
  // Gestion du parc
  if copy(TypeArticle,1,2)='PA' then
  begin
    LoadTobArticlecompParc;
  end;
     
  // Gestion des métrés
  if GetField('GA_ARTICLE') <> '' then
  begin
    MetreArticle := TMetreArt.CreateArt(TypeArticle, GetField('GA_ARTICLE'));
    if not MetreArticle.ControleMetre then
    begin
      // Modified by f.vautrain 14/11/2017 14:36:54 - MICHEL SA : Violation d'accès sur ouverture article pourcentage
      If Assigned(BExcel) then BExcel.visible := False;
      BAppelVariable.Visible := false;
    end
    else
    begin
      // Modified by f.vautrain 14/11/2017 14:36:54 - MICHEL SA : Violation d'accès sur ouverture article pourcentage
      If Assigned(BExcel) then BExcel.visible          := True;
      BAppelVariable.Visible  := True;
      RepServeur              := MetreArticle.RepMetre + MetreArticle.fFileName;
      RepPoste                := MetreArticle.RepMetreLocal + MetreArticle.fFileName;
    end;
  end;
  {$ENDIF}
  // -------------

  if DS.State <> dsInsert then SetControlEnabled('GA_CODEARTICLE', False);
  // Taxes
  for i := 1 to MaxTaxe do
  begin
    if GetControl ('LTAXE' + IntToStr (i)) <> nil then
      SetControlText('LTAXE' + IntToStr(i), RechDom('GCCATEGORIETAXE', 'TX' + IntToStr(i), False));
  end;

  //Specifique Mode pour gérer la table ARTICLECOMPL
  if (CtxMode in V_PGI.PGIContexte) and ((GetPresentation = ART_ORLI) or (GetField('GA_TYPEARTICLE') = 'FI')) then
  begin
    if TobZonelibre = nil then TobZonelibre := TOB.Create('ARTICLECOMPL', nil, -1);
    if not TobZonelibre.SelectDB('"' + GetField('GA_ARTICLE') + '"', nil) then TobZonelibre.InitValeurs;
    if (GetField('GA_TYPEARTICLE') = 'FI') and
      (TobZonelibre.FieldExists('GA2_RATTACHECLI')) and
      (TobZonelibre.GetValue('GA2_RATTACHECLI') = '') then TobZonelibre.PutValue('GA2_RATTACHECLI', 'FAC');
    TobZonelibre.PutEcran(TFFiche(Ecran));
    for i := 1 to 2 do ChangeStatArt(IntToStr(i));
  end;

  TypeArticle := GetField('GA_TYPEARTICLE');
  ArtSpecif := IsArticleSpecif(TypeArticle);
  if (ArtSpecif = 'FicheModePre') then
  begin
    SetControlEnabled('GA_TENUESTOCK', False);
    SetControlEnabled('GA_LOT', False);
    if (ctxFO in V_PGI.PGIContexte) then SetControlVisible('GB_PRIX', False);
    exit;
  end;

  CalculEnCours := False; // Pas de recalcul automatique en cours de prix ou coefficients
  if ModifparLot then FaireCalcul := False;
  if Consultation then SetControlVisible('BINSERTNOMENC', False);
  if GetField('GA_TENUESTOCK') <> 'X' then SetControlEnabled('SUBSTITUTION', False);
  ArticleFourn := False;
  SetFocusControl('GA_CODEARTICLE');
  ArticlePrin := GetField('GA_ARTICLE');
  FournPrinc := GetField('GA_FOURNPRINC'); //pour se positionner sur le fournisseur principal (FicheListe)
  OldFournPrinc := FournPrinc;
  OldCodeCAB := GetField('GA_CODEBARRE');
  ReferArtFourn := '';
  //
  // Modif BTP 06/02/2001
  //
  ModifChampPR := False;
  // ----------------------
  ModifChampHT := False;
  ModifChampTTC := False;
  ClickActuPrix := True; //Gestion calculPrix

  {$IFNDEF CCS3}
  AppliquerConfidentialite(Ecran, TypeArticle);
  {$ENDIF}

  // Si l'article est en mouvement, interdire la modification de GA_LOT et Unité de stock
  // DCA - FQ MODE 10852 - Pas de requête sur LIGNE
  if (CtxMode in V_PGI.PGIContexte) then IsEnabled := True
  else if (TypeArticle <> 'PRE') and (TypeArticle <> 'FI') and (GetControlText('GA_ARTICLE') <> '') and (GetField('GA_TENUESTOCK') = 'X')
    then IsEnabled := not ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE="' + GetControlText('GA_ARTICLE') + '" AND GL_VIVANTE="X"')
  else IsEnabled := True;
  SetControlProperty('GA_LOT', 'ReadOnly', not IsEnabled);
{$IFDEF PREMIUM}
{$IFDEF CCS3}
  if VH_GC.GCAchatStockSeria then
{$ENDIF}
{$ENDIF}
  SetControlEnabled('GA_QUALIFUNITESTO', IsEnabled);
  if assigned(GetControl('GA_COEFCONVQTEVTE')) then
  begin
    SetControlEnabled('GA_COEFCONVQTEVTE',(IsEnabled) and (GetControltext('GA_QUALIFUNITESTO')<>''));
    THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := ((not IsEnabled) or (GetControltext('GA_QUALIFUNITESTO')=''));
  end;

  // affichage onglet dimension suivant type d'article
  OngletDimension := TTAbSheet(GetControl('P_DIMENSION'));
  if OngletDimension <> nil then OngletDimension.TabVisible := (GetField('GA_STATUTART') = 'GEN') and (not ISLine) ;
  PanelDimension := TPANEL(GetControl('PDIMENSION'));
  if PanelDimension <> nil then PanelDimension.Visible := (GetField('GA_STATUTART') = 'DIM') and (Not IsLIne);
  if GetField('GA_STATUTART') = 'DIM' then
  begin
    // DCA - FQ MODE 10813
    SetControlEnabled('GA_COLLECTION', ChampModif('GA_COLLECTION', False, True));
    // Décalage fenêtre de l'article dimensionné
    if (Top <> 0) and (Left <> 0) then
    begin
      Ecran.Top := Top;
      Ecran.Left := Left
    end;
  end;

  if (GetField('GA_TYPEARTICLE') = 'ARP') then
  begin
    // affichage onglet dimension suivant type d'article
    TTAbSheet(GetControl('P_DIMENSION')).TabVisible := false;
    TPANEL(GetControl('PDIMENSION')).Visible := false;
  end;

  OldDimMasque := GetField('GA_DIMMASQUE');
  if GetField('GA_STATUTART') = 'GEN' then
  begin
    bZeroDim := not ExisteSQL('select GA_CODEARTICLE from ARTICLE where GA_STATUTART="DIM" and GA_CODEARTICLE="' + GetControlText('GA_CODEARTICLE') + '"');
    SetControlEnabled('GA_DIMMASQUE', bZeroDim);
  end
  else SetControlEnabled('GA_DIMMASQUE', (DS.State in [dsInsert]));
  SetControlVisible('GA_PRIXUNIQUE', (DS.State in [dsInsert]) or Boolean(GetField('GA_STATUTART') = 'GEN'));
  SetControlEnabled('BPROFIL', (DS.State in [dsInsert]));

  // affichage fiche suivant le type des article nomenclature et marchandise et tenue en stock
  // Modif BTP
  {$IFDEF BTP}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM') or (GetField('GA_TYPEARTICLE') = 'OUV');
  {$ELSE}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM');
  {$ENDIF}
  IsMacro := (GetField('GA_TYPENOMENC') = 'MAC');
  IsParc := (copy(GetField('GA_TYPEARTICLE'),1,2) = 'PA');
  IsNomenParc := (GetField('GA_TYPEARTICLE') = 'PA2');

  SetControlVisible('GA_PAHT', not IsNomenc);
  SetControlVisible('T_PAHT', IsNomenc);
  if not IsLine then SetControlVisible('GA_DPA', not IsNomenc);
  if not IsLine then SetControlVisible('T_DPA', IsNomenc);
  SetControlVisible('GA_PMAP', not IsNomenc);
  SetControlVisible('T_PMAP', IsNomenc);
  SetControlVisible('GA_PRHT', not IsNomenc);
  SetControlVisible('GA_DPR', not IsNomenc);
  if VH_GC.GCIfDefCEGID then
    if not IsNomenc then
    begin
      SetControlProperty('GA_DPR', 'ReadOnly', False);
      SetControlProperty('GA_DPR', 'Color', clWindow);
      SetControlProperty('GA_DPR', 'TabStop', True);
    end;

  SetControlVisible('T_PRHT', IsNomenc);
  {$IFNDEF BTP}
  SetControlVisible('GA_PMRP', not IsNomenc);
  SetControlVisible('T_PMRP', IsNomenc);
  {$ENDIF}
  {$IFDEF CCS3}
  SetControlVisible('GA_LOT', False);
  SetControlVisible('GA_NUMEROSERIE', False);
  {$ELSE}
  {$IFNDEF BTP}
  SetControlVisible('GA_LOT', not IsNomenc);
  SetControlVisible('GA_NUMEROSERIE', not IsNomenc);
  SetControlVisible('GA_CONTREMARQUE', not IsNomenc);
  if (ctxAffaire in V_PGI.PGIContexte) then SetControlVisible('GA_CONTREMARQUE', False); // mcd 09/02/2002 pas de ctremarque (champ existe article %)
  {$ENDIF}
  {$ENDIF}

  // Nomenclature
  //SetControlProperty('P_DIMENSION','TabVisible',not IsNomenc) ;
  //
  // Modif BTP
  //
  {$IFDEF BTP}
  SetControlVisible('BNOMENC1', IsNomenc or IsNomenParc);
  SetControlVisible('GA_NUMEROSERIE', False);
  {$ENDIF}
  // -----------------
  SetControlVisible('BNOMENC', IsNomenc );
  SetControlProperty('GA_PRHT', 'ReadOnly', IsNomenc);
  SetControlProperty('GA_PRHT', 'TabStop', not IsNomenc);
  SetControlProperty('GA_PAHT', 'ReadOnly', IsNomenc);
  SetControlProperty('GA_PAHT', 'TabStop', not IsNomenc);
  SetControlProperty('GA_PMAP', 'ReadOnly', IsNomenc);
  SetControlProperty('GA_PMAP', 'TabStop', not IsNomenc);

{$IFDEF PREMIUM}
{$IFDEF CCS3}
  if VH_GC.GCAchatStockSeria then
  begin
   SetControlEnabled('GA_TENUESTOCK', (not IsNomenc and not Consultation));
   SetControlEnabled('GA_CONTREMARQUE', (not IsNomenc and not Consultation));
   SetControlVisible('TGA_FOURNPRINC', not IsNomenc);
   SetControlVisible('GA_FOURNPRINC', not IsNomenc);
   SetControlVisible('FOURNPRINC', not IsNomenc);
  end;
{$ENDIF}
{$ELSE}
  SetControlEnabled('GA_TENUESTOCK', ((not IsNomenc) and (not Consultation) and (not IsLine)));
  SetControlEnabled('GA_CONTREMARQUE', ((not IsNomenc) and (not Consultation) and (not IsLine)));
  SetControlVisible('TGA_FOURNPRINC', (not IsNomenc) and (not IsLine));
  SetControlVisible('GA_FOURNPRINC', (not IsNomenc) and (not IsLine));
  SetControlVisible('FOURNPRINC', (not IsNomenc) and (not IsLine));
{$ENDIF}

  //
  //
  // Modif BTP du 06/02/2001
  {$IFNDEF BTP}
  if ((TypeArticle = 'PRE') or (TypeArticle = 'CTR')) then
    SetControlEnabled('GA_TENUESTOCK', (not IsNomenc and not Consultation));
  {$ENDIF}

  if ((TypeArticle = 'NOM') or (TypeArticle = 'OUV')) then
  begin
    {$IFDEF BTP}
    SetControlVisible('TGA_TYPENOMENC', false);
    SetControlVisible('GA_TYPENOMENC', false);
    {$ELSE}
    SetControlVisible('TGA_TYPENOMENC', IsNomenc);
    SetControlVisible('GA_TYPENOMENC', IsNomenc);
    {$ENDIF}
  end;
  {$IFDEF CHR}
  SetControlVisible('GA_TYPENOMENC', False);
  {$ENDIF}

  if IsNomenc then
  begin
    SetControlProperty('GA_PMAP', 'color', clbtnFace);
    SetControlProperty('GA_PAHT', 'color', clbtnFace);
    SetControlProperty('GA_PRHT', 'color', clbtnFace);
    SetControlProperty('GA_TYPENOMENC', 'TabStop', False);
    SetControlProperty('GA_TYPENOMENC', 'color', clbtnFace);
    SetField('GA_TENUESTOCK', '-');
    SetField('GA_CONTREMARQUE', '-');
    {$IFDEF EAGLCLIENT}
    //
    // Modif BTP 06/02/2001
    //
    {$IFDEF BTP}
    THValComboBox(GetControl('GA_CALCPRIXPR')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    SetControlVisible('B_DUPLICATION', not isnomenc);
    {$ENDIF}
    // ----------------
    THValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    THValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")'; // AND (CO_CODE NOT LIKE "HT")';
    {$ELSE}
    //
    // Modif BTP 06/02/2001
    //
    {$IFDEF BTP}
    THDBValComboBox(GetControl('GA_CALCPRIXPR')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    SetControlVisible('B_DUPLICATION', not isnomenc);
    {$ENDIF}
    // ----------------
    THDBValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    THDBValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")'; // AND (CO_CODE NOT LIKE "HT")';
    {$ENDIF}
  end
    // SPECIF pas MODE
  else if (ArtSpecif <> 'FicheModePre') and (ArtSpecif <> 'FicheModeFi') and (ArtSpecif <> 'FicheGescomFi') and
					(copy(ecran.name,1,9)<>'BTARTPARC') and
    			(copy(ecran.name, 1, 13) <> 'BTARTPOURCENT') and (copy(ecran.name, 1, 13) <> 'AFARTPOURCENT') then
  begin
    SetControlProperty('GA_PMAP', 'color', clWindow);
    SetControlProperty('GA_PAHT', 'color', clWindow);
    SetControlProperty('GA_PRHT', 'color', clWindow);
    {$IFDEF EAGLCLIENT}
    //
    // Modif BTP 06/02/2001
    {$IFDEF BTP}
    THValComboBox(GetControl('GA_CALCPRIXPR')).Plus := ''+PlusCalcLine;

    {$ENDIF}
    // ---------------------------
    THValComboBox(GetControl('GA_CALCPRIXHT')).Plus := ''+PlusCalcLine;
    THValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := ''+PlusCalcLine;
    {$ELSE}
    //
    // Modif BTP 06/02/2001
    {$IFDEF BTP}
    THDBValComboBox(GetControl('GA_CALCPRIXPR')).Plus := ''+PlusCalcLine;
    AvertirTable ('BTCALCPRIXPR');
    {$ENDIF}
    // ---------------------------
    THDBValComboBox(GetControl('GA_CALCPRIXHT')).Plus := ''+PlusCalcLine;
    THDBValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := ''+PlusCalcLine;
    AvertirTable ('BTCALCPRIXHT');
    AvertirTable ('BTCALCPRIXTTC');
    {$ENDIF}
  end;

  if (TypeArticle = 'PRE') then
  begin
    SetControlEnabled('GA_TENUESTOCK', False);
    SetControlEnabled('GA_LOT', False);
  end;

  if IsNomenc then TFFiche(Ecran).HelpContext := 110000019 else TFFiche(Ecran).HelpContext := 110000018;

  if IsArticleSpecif(TypeArticle) = 'FicheAffaire' then TFFiche(Ecran).HelpContext := 120000019;
  // Nomenclature type macro
  SetControlProperty('P_PHOTO', 'TabVisible', (not IsMacro) and (not IsLine));
  {$IFDEF BTP}
  SetControlProperty('P_DIVERS', 'TabVisible', ((not IsNomenc) and (not IsLine)));
  {$ELSE}
  SetControlProperty('P_DIVERS', 'TabVisible', ((not IsMacro) and (not IsLine));
  {$ENDIF}
  {$IFDEF GPAO}
  // Voir pour statuer sur le sort de l'onglet P_CARACTERISTIQUE suivant les types de fiche utilisé
  {$ELSE}
  SetControlProperty('P_CARACTERISTIQUE', 'TabVisible', (not IsMacro) and (not IsLine));
  {$ENDIF}
  if (ctxFO in V_PGI.PGIContexte)
    then SetControlProperty('P_TARIF', 'TabVisible', False)
  else SetControlProperty('P_TARIF', 'TabVisible', (not IsMacro));
  //MODIF CHR
  if (ECRAN.name <> 'MBOARTFINANCIER') and (ECRAN.name <> 'HRPRESTATIONS') then // Fiche Article Financier Mode
    SetControlVisible('Pgeneral', not IsMacro);
  if isMacro then TFFiche(Ecran).Pages.ActivePage := TTabSheet(Ecran.FindComponent('PGeneral'));
  //if (TToolBarButton97(GetControl('B_DUPLICATION')).Enabled) then
  //    SetControlEnabled('B_DUPLICATION',Not IsMacro) ;
  SetControlEnabled('B_DUPLICATION', (not IsMacro and not Consultation));
  SetControlVisible('GB_PROFIL', (not IsMacro) and (not Isparc));
  SetControlVisible('GA_BLOCNOTE_MAC', IsMacro);
  if IsMacro then
  begin
    BACTUPRIX.enabled := False;
    SetControlProperty('GA_BLOCNOTE_MAC', 'TabOrder', 9);
    SetControlEnabled('BPROFIL', False);
  end;

  if (copy(ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(ecran.name, 1, 13) = 'AFARTPOURCENT') then
    Begin
      SelTypearticle.text := Getfield('GA_LIBCOMPL');
    /////////////////////////////////
    end
  else
    // SPECIF MODE  pas de gestion des bases tarif HT
    if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
  end
    /////////////////////////////////
    {$IFDEF GESCOM}
  else if (ArtSpecif = 'FicheGescomFi') then
  begin
  end
    {$ENDIF}
  else
    // gestion des bases tarif HT
  begin
    //
    // Modif BTP 06/02/2001
    //
    {$IFDEF BTP}
    if (GetField('GA_CALCPRIXPR') = 'AUC') or (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S') then IsAucOuStock := False else
      IsAucOuStock := True; // si aucun ou stock
    SetControlEnabled('GA_COEFFG', (IsAucOuStock and not Consultation));
    SetControlEnabled('GA_DPRAUTO', (IsAucOuStock and not Consultation));
    if not IsAucOuStock then
    begin
      SetField('GA_DPRAUTO', '-');
      SetFocusControl('GA_CALCPRIXHT');
    end;
    SetControlProperty('GA_DPR', 'Visible', (not (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S')) and (not IsNomenc));
    SetControlProperty('T_PRHT', 'Visible', (copy(GetField('GA_CALCPRIXPR'), 1, 1) = 'S') or IsNomenc);
    if IsAucOuStock then SetFocusControl('GA_COEFFG');
    {$ENDIF}

    //
    if (GetField('GA_CALCPRIXHT') = 'AUC') or (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') then IsAucOuStock := False else
      IsAucOuStock := True; // si aucun ou stock
    //  BBI, fiche correction 10410
      SetControlEnabled('GA_COEFCALCHT',(IsAucOuStock and Not Consultation)) ;
      CALCAUTOHT.Enabled := (IsAucOuStock and Not Consultation) ;
    //  BBI, fin fiche correction 10410
    if not (ctxMode in V_PGI.PGIContexte) then
    begin
      SetControlProperty('GA_PVHT', 'ReadOnly', IsAucOuStock and not ChampModif('GA_PVHT', GetField('GA_PRIXUNIQUE')='X', False));
      SetControlProperty('GA_PVHT', 'TabStop', not (IsAucOuStock and not ChampModif('GA_PVHT', GetField('GA_PRIXUNIQUE')='X', False)));
    end;
    if not IsAucOuStock then
    begin
      //  BBI, fiche correction 10410
      CALCAUTOHT.Checked := False;
      //  BBI, fin fiche correction 10410
      SetFocusControl('GA_CALCPRIXTTC');
    end;
    SetControlProperty('GA_PVHT', 'Visible', not (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S'));
    SetControlProperty('T_STOCKPVHT', 'Visible', (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S'));
    //  BBI, fiche correction 10410
        If IsAucOuStock then SetFocusControl('GA_COEFCALCHT') ;
    //  BBI, fin fiche correction 10410

    // gestion des bases tarif TTC
    if (GetField('GA_CALCPRIXTTC') = 'AUC') or (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S') then IsAucOuStock := False else
      IsAucOuStock := True; // si aucun ou stock
    SetControlEnabled('GA_COEFCALCTTC', (IsAucOuStock and not Consultation));
    if CALCAUTOTTC <> nil then CALCAUTOTTC.Enabled := (IsAucOuStock and not Consultation);
    if not (ctxMode in V_PGI.PGIContexte) then
    begin
      SetControlProperty('GA_PVTTC', 'ReadOnly', IsAucOuStock and not ChampModif('GA_PVTTC', GetField('GA_PRIXUNIQUE')='X', False));
      SetControlProperty('GA_PVTTC', 'TabStop', not (IsAucOuStock and not ChampModif('GA_PVTTC', GetField('GA_PRIXUNIQUE')='X', False)));
    end;
    if not IsAucOuStock then
    begin
      SetField('GA_CALCAUTOTTC', '-');
        //mcd 25/09/03 ajout test présence zone.. n'existe pas en GI/GA
      If THValComboBox(GetControl('GA_ARRONDIPRIX')) <> nil
         then SetFocusControl('GA_ARRONDIPRIX'); // aller sur la prochaine zone de saisie
    end;
    SetControlProperty('GA_PVTTC', 'Visible', not (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S'));
    SetControlProperty('T_STOCKPVTTC', 'Visible', (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S'));
    if IsAucOuStock then SetFocusControl('GA_COEFCALCTTC');
  end;

  // Taxes
  for i := 1 to MaxTaxe do
  begin
    if GetControl ('LTAXE' + IntToStr (i)) <> nil then
      SetControlText('LTAXE' + IntToStr(i), RechDom('GCCATEGORIETAXE', 'TX' + IntToStr(i), FALSE));
  end;

  // SPECIF AFFAIRE /////////////////
  if ArtSpecif = 'FicheAffaire' then
  begin
    if (TypeArticle <> 'MAR') and (TypeArticle <> 'NOM') and (TypeArticle <> 'OUV') then
      // Champs achats visible sur les articles de type marchandises uniquement
    begin
      SetControlVisible('GA_FOURNPRINC', False);
      SetControlVisible('TGA_FOURNPRINC', False);
      //      SetControlVisible ('GB_PRIXACHAT',False);  gm le 14/01/02
      SetControlVisible('BPOPCOMPLEMENT', False); // Bouton complément reprend la liste des fournisseurs
      if ( ( (ctxScot in V_PGI.PGIContexte) or (TypeArticle = 'CTR') ) and ( not VH_GC.GCIfDefCEGID ) ) or
             ( (ctxScot in V_PGI.PGIContexte) and ( VH_GC.GCIfDefCEGID ) )then 
        SetControlVisible('GB_PRIXACHAT', False) // GB des prix d'achats
      else // on ne cache que les DPA et DPR - on conserve les PA en GA
      if copy(ecran.name, 1, 12) <> 'BTPRESTATION' then
      begin
        SetControlVisible('GA_DPA', False);
        SetControlVisible('TGA_DPA', False);
        SetControlVisible('GA_DPR', False);
        SetControlVisible('TGA_DPR', False);
      end;
      if VH_GC.GCIfDefCEGID then
      begin
        SetControlVisible('GA_PMRP', False);
        SetControlVisible('TGA_PMRP', False);
      end;
    end;
    // Plus sur méthode de calcul du PV change en fonction du type d'articles ( prix achats ou pas)
    // Modif BTP
    {$IFNDEF BTP}
    // -----
    if (ctxScot in V_PGI.PGIContexte) or ((TypeArticle <> 'MAR') and (TypeArticle <> 'NOM')) then
      SetControlProperty('GA_CALCPRIXHT', 'Plus', 'AND CO_LIBRE="DEF"');
    {$ENDIF}
  end else
  begin
    ////// article de remplacement //////////
    //JT - eQualité n° 10671 - Test si GA_REMPLACEMENT est présent dans la fiche
    if Ecran.FindChildControl('GA_REMPLACEMENT') <> nil then
    begin
      if GetControlText('GA_REMPLACEMENT') <> '' then
      begin
        OldRemplacement := GetControlText('GA_REMPLACEMENT');
        SetControlText('REMPLACEMENT', Trim(Copy(GetControlText('GA_REMPLACEMENT'), 1, 18)));
        THLabel(GetControl('TREMPLACEMENT')).Caption := LibelleArticleGenerique(Copy(OldRemplacement, 1, 18));
      end else
      begin
        SetControlText('REMPLACEMENT', '');
        SetControlText('TREMPLACEMENT', '');
      end;
    end;

    //JT - eQualité n° 10671 - Test si GA_SUBSTITUTION est présent dans la fiche
    if Ecran.FindChildControl('GA_SUBSTITUTION') <> nil then
    begin
      if GetControlText('GA_SUBSTITUTION') <> '' then
      begin
        OldSubstitution := GetControlText('GA_SUBSTITUTION');
        SetControlText('SUBSTITUTION', trim(Copy(GetControlText('GA_SUBSTITUTION'), 1, 18)));
        THLabel(GetControl('TSUBSTITUTION')).Caption := LibelleArticleGenerique(Copy(OldSubstitution, 1, 18));
      end else
      begin
        SetControlText('SUBSTITUTION', '');
        SetControlText('TSUBSTITUTION', '');
      end;
    end;
  end;

  // affiche Photo
  if VH_GC.GCPHOTOFICHE <> '' then AffichePhoto else SetControlVisible('P_PHOTO', false);

  ///////// Dimensions //////////////////
  if (GetField('GA_STATUTART') = 'DIM') then // article dimensionné
  begin
    AffecteDimension;
  end; // fin article dimensionné

  if (GetField('GA_STATUTART') = 'GEN') and (not DuplicFromMenu) and (not DuplicFromFiche) then // article générique (dimension)
  begin
    InitDimensionsArticle;
  end;

  if (copy(Ecran.Name, 1, 9) = 'GCARTICLE') then
  begin
    if SpecifGescom <> nil then SpecifGescom.Load;
  end;

  if (ctxMode in V_PGI.PGIContexte) and (DS.State in [dsInsert]) and ((DuplicFromMenu) or (DuplicFromFiche))
    then SetControlEnabled('B_DUPLICATION', False);
  if (DuplicFromMenu) then // duplication
  begin
    SetControlEnabled('GA_COEFFG', True);
    DPRAUTO.Enabled := True;
    SetControlEnabled('GA_COEFCALCHT', True);
    CALCAUTOHT.enabled := True;
    SetControlEnabled('GA_COEFCALCTTC', True);
    CALCAUTOTTC.Enabled := True;
    Duplication(ArtToDuplic);
    ArtToDuplic := '';
  end;

  /////////// SPECIF MODE ///////////
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    SetControlProperty('PMESURE2', 'Visible', False);
    SetControlProperty('BMESURE_MODE', 'Visible', True);
    SetControlProperty('BMESURE', 'Visible', False);
    // Présentation ORLI
    if GetPresentation = ART_ORLI then
    begin
      // Onglet Général
      SetControlProperty('B_FAMILLE3', 'Visible', False);
      SetControlProperty('B_FAMILLE8', 'Visible', True);
      SetControlProperty('HP_FAMILLE8', 'Visible', True);
      SetControlProperty('HP_FERMEGEN', 'Visible', False);
      SetControlProperty('HP_STATART', 'Visible', True);
      SetControlProperty('GB_PROFILGEN', 'Visible', False);
      // Onglet Caractéristique
      SetControlProperty('HP_FERMECAR', 'Visible', True);
      SetControlProperty('TGA_TYPEEMPLACE', 'Top', 270);
      SetControlProperty('GA_TYPEEMPLACE', 'Top', 266);
      SetControlProperty('GB_PROFILCAR', 'Visible', True);
    end else
    begin
      SetControlProperty('GA_INVISIBLEWEB1', 'Visible', V_PGI.LaSerie > S3);
    end;
  end;

  if (not (ctxMode in V_PGI.PGIContexte) and (TypeArticle <> 'FI')) then
    for x := 0 to TPopupMenu(GetControl('POPMENU')).items.count - 1 do
    begin
      if TPopupMenu(GetControl('POPMENU')).items[x].name = 'mnTarifBase' then
      begin
        TPopupMenu(GetControl('POPMENU')).items[x].visible := False;
        break;
      end;
      {$IFDEF BTP}
      if TPopupMenu(GetControl('POPMENU')).items[x].name = 'mnDispo' then
      begin
        // MODIF LS 
        TPopupMenu(GetControl('POPMENU')).items[x].visible := True;
        break;
      end;
      {$ENDIF}
    end;

  if (ctxFO in V_PGI.PGIContexte) then
  begin
    SetControlVisible('BPOPCOMPLEMENT', False); // Bouton complément reprend la liste des fournisseurs
    SetControlVisible('BPOPMENU', False);
  end;
  ///////////////////////////////////

  /////////// SPECIF AFFAIRE /////////
  if (ctxAffaire in V_PGI.PGIContexte) then
  begin
    {$IFNDEF BTP}
    SetControlProperty('TGA_QUALIFUNITEACT', 'Visible', True);
    SetControlProperty('GA_QUALIFUNITEACT', 'Visible', True);
    {$ENDIF}
    if (DS.State in [dsInsert]) then
      SetControlEnabled('B_DUPLICATION', False); // Spécif Affaire modif PL du 31/05/2000
    SetControlVisible('BINSERTNOMENC', False);
    // Modif BTP
    if ArtSpecif = 'FicheBtp' then
    	if (TypeArticle = 'PRE') then SetControlCaption('TTYPEARTICLE', 'Prestation')
			else SetControlCaption('TTYPEARTICLE', 'Article')
    else
      // --------
      if ArtSpecif <> 'FicheAffaire' then
      SetControlCaption('TTYPEARTICLE', 'Fourniture'); // PL 20/06/2000
    {$IFDEF AFFAIRE}
    AFFormuleVariable; //Affaire-ONYX
    {$ENDIF}
  end;

  if FicheStock or FicheStoVte then
  begin
    TFFiche(Ecran).Pages.ActivePage := TTabSheet(Ecran.FindComponent('P_Dimension'));
    AjusterTailleTHDim;
    if ArticleDimFocus <> '' then ItemDimSetFocus(DimensionsArticle, nil, ArticleDimFocus); // Positionnement sur l'item sélectionné = ArticleDimFocus
  end
  else
  begin
    if TobArt <> nil then
    begin
      TobArt.free;
      TobArt := nil;
    end;
    TobArt := TOB.Create('ARTICLE', nil, -1);
    {$IFDEF EAGLCLIENT}
    TobArt.dupliquer(DS.CurrentFille, False, True, False);
    {$ELSE}
    TobArt.GetEcran(Ecran);
    {$ENDIF}
    if TobArtLoad <> nil then
    begin
      TobArtLoad.Free;
      TobArtLoad := nil;
    end;
    TobArtLoad := TOB.Create('ARTICLE', nil, -1);
    TobArtLoad.dupliquer(TobArt, True, True, True);
  end;

  if ModifparLot then
  begin
    SetArguments(StSQL);
    TobArt.GetEcran(Ecran);
  end;

  if not (DS.State in [dsInsert]) then DerniereCreate := ''; // Contournement bug création+modif
  //
  // Modif BTP du 02/06/2001
  //

  {$IFDEF BTP}
  if (isnomenc) then
  begin
    SetControlProperty('GA_PRHT', 'color', clbtnFace);

    LibTypNomenc := RechDom('BTTYPENOMENC', getfield('GA_TYPENOMENC'), FALSE);
    SetControlCaption('TTYPEARTICLE', LibTypNomenc);

    if copy(Getfield('GA_TYPENOMENC'), 1, 2) = 'OU' then
    begin
(*      if Getfield('GA_TYPENOMENC') = 'OUV' then
      begin *)
        Ongletcaracteristique := TTAbSheet(GetControl('P_CARACTERISTIQUE'));
        if Ongletcaracteristique <> nil then Ongletcaracteristique.TabVisible := false;
        Panelcaracteristique := TPANEL(GetControl('P_CARACTERISTIQUE'));
        if Panelcaracteristique <> nil then Panelcaracteristique.Visible := false;
      // MODIF LS LE 28/05/2004
      if Getfield('GA_TYPENOMENC') = 'OUV' then TGroupBox(GetControl('GB_CEE')).visible := false;
      // --
      SetControlVisible('BPROFIL', false);
      onglettarif := TTAbSheet(GetControl('P_TARIF'));
      if onglettarif <> nil then OngletTarif.tabVisible := false;
      PanelTarif := TPANEL(GetControl('P_TARIF'));
      if PanelTarif <> nil then PanelTarif.Visible := false;
      SetControlProperty('GA_FOURNPRINC', 'visible', false);
      SetControlProperty('TGA_FOURNPRINC', 'visible', false);
      SetControlProperty('FOURNPRINC', 'visible', false);
      SetControlProperty('TREMPLACEMENT_', 'visible', false);
      SetControlProperty('REMPLACEMENT', 'visible', false);
      SetControlProperty('TSUBSTITUTION_', 'visible', false);
      SetControlProperty('SUBSTITUTION', 'visible', false);
      SetControlProperty('TGA_QUALIFUNITEFACT1', 'visible', true);
      SetControlProperty('GA_QUALIFUNITEVTE1', 'visible', true);
      SetControlProperty('TGA_DOMAINE', 'visible', true);
      SetControlProperty('GA_DOMAINE', 'visible', true);
      SetControlVisible('TGA_CODEBARRE', false);
      SetControlVisible('GA_CODEBARRE', false);
      SetControlVisible('GA_QUALIFCODEBARRE', false);
    end;
  end else
  begin
    // Pas de gestion de stock pour l'instant
    //SetControlEnabled('GA_TENUESTOCK', false);
    // --------------------
    CalculBaseTarifPR(true);
    CalculBaseTarifHT(true);
    CalculBaseTarifTTC(true);
    if (Pos(Getfield('GA_TYPEARTICLE'),'MAR;PRE;ARP;POU')>0) then
    begin
      SetControlProperty('TGA_DOMAINE', 'visible', true);
      SetControlProperty('GA_DOMAINE', 'visible', true);
    end;
  end;
  {$ENDIF}

  //JD : Création d'articles à partir d'un repertoire de photos
  if CreatArticlePhoto then
  begin
    TOBInfoArt := TheTOB;
    if TOBInfoArt <> nil then MAJFicheArticleWithTOB;
  end;

  if TypeArticle <> 'FI' then
  begin
    if GetParamSoc('SO_GCFAMHIERARCHIQUE') = True then
    begin
      if (THValComboBox(GetControl('GA_FAMILLENIV1')).value = '') then
      begin
        SetControlEnabled('GA_FAMILLENIV2', false);
        SetControlEnabled('GA_FAMILLENIV3', false);
      end
      else
      begin
        SetControlEnabled('GA_FAMILLENIV2', true);
        stTemp := '';
        stTemp := stTemp + THValComboBox(GetControl('GA_FAMILLENIV1')).value;
        if stTemp <> '' then THValComboBox(GetControl('GA_FAMILLENIV2')).Plus := 'and CC_LIBRE like "%' + stTemp + '%"';
        if stTemp <> '' then THValComboBox(GetControl('GA_FAMILLENIV3')).Plus := 'and CC_LIBRE like "%' + stTemp + '%"';
        if (THValComboBox(GetControl('GA_FAMILLENIV2')).value = '') then
          SetControlEnabled('GA_FAMILLENIV3', false)
        else
        begin
          SetControlEnabled('GA_FAMILLENIV3', true);
          stTemp := stTemp + THValComboBox(GetControl('GA_FAMILLENIV2')).value;
          if stTemp <> '' then THValComboBox(GetControl('GA_FAMILLENIV3')).Plus := 'and CC_LIBRE like "%' + stTemp + '%"';
        end;
      end;
    end;
  end;
  {$IFDEF GPAO}
  if (Ecran <> nil) and (Ecran.Name = 'WARTICLE_FIC') then
  begin
    WProfilArticleState;
    { Accès à l'initialisation des unités }
    if TMenuItem(GetControl('MNUNITEINIT')) <> nil then
      TMenuItem(GetControl('MNUNITEINIT')).Enabled := TFFiche(Ecran).FTypeAction <> taConsult;
    { Accès aux conversions d'unités }
    if TMenuItem(GetControl('MNUNITECONVERSION')) <> nil then
      TMenuItem(GetControl('MNUNITECONVERSION')).Enabled := ((GetField('GA_CODEFORME') <> '') and (TFFiche(Ecran).fTypeAction <> TaConsult));
    SetBtArtNatState;
    if not (DS.State in [dsInsert]) then wGetRetour;
  end;
  {$ENDIF}
  {$IFNDEF CCS3}
  AppliquerConfidentialite(Ecran, TypeArticle);
  {$ENDIF}

  TobLiensOle := nil; // Duplication article : table LIENSOLE

  if GetField('GA_TENUESTOCK') = 'X' then
  begin
    if (CtxMode in V_PGI.PGIContexte) then
    begin
      SetControlProperty('GB_PRIX_ACHAT', 'Caption', 'Prix d''achat ' + RechDom('GCDEPOTABR', VH_GC.GCDepotDefaut, False));
      SetControlProperty('GB_PRIX_REVIENT', 'Caption', 'Prix de revient ' + RechDom('GCDEPOTABR', VH_GC.GCDepotDefaut, False));
    end else
    begin        //mcd 26/09/05 ajout du test fiche.. en GI/GA terme spécifique sur ce champ, il ne doit pas être changé en duplic
      If (Ecran.name) <> 'AFARTICLE' then SetControlProperty('TGA_PMRP', 'Caption', 'PMRP ' + RechDom('GCDEPOT', VH_GC.GCDepotDefaut, False));
//      SetControlProperty('TGA_PMAP', 'Caption', 'PMAP ' + RechDom('GCDEPOT', VH_GC.GCDepotDefaut, False));
    end;
  end;

  {$IFDEF BTP}
  if GetField('GA_TENUESTOCK')='X' Then
  begin
    SetControlEnabled ('GA_QUALIFUNITESTO',True);
    SetControlEnabled('TGA_QUALIFUNITESTO',True);
    if assigned(GetControl('GA_COEFCONVQTEVTE')) then
    begin
   	  SetControlEnabled('GA_COEFCONVQTEVTE',GetControltext('GA_QUALIFUNITESTO')<>'');
      THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := (GetControltext('GA_QUALIFUNITESTO')='');
    end;
  end else
  begin
    SetControlEnabled ('GA_QUALIFUNITESTO',False);
    SetControlEnabled('TGA_QUALIFUNITESTO',False);
    if assigned(GetControl('GA_COEFCONVQTEVTE')) then
    begin
      SetControlEnabled('GA_COEFCONVQTEVTE',false);
      THEdit(GetControl('GA_COEFCONVQTEVTE')).ReadOnly := true;
    end;
  end;
	TPageControl(GetControl('Pages')).ActivePageIndex := 0;

  if (Pos(TypeArticle, 'ARP;MAR;PRE;') > 0) and (CodeSousTarif <> nil) then
  begin
		CodeSousTarif.Plus := 'AND BSF_FAMILLETARIF="'+GetField('GA_TARIFARTICLE')+'"';
  end;
  {$ENDIF}

//uniquement en line
{*
	if TToolBarButton97(GetControl('BPOPCOMPLEMENT')) <> nil then TToolBarButton97(GetControl('BPOPCOMPLEMENT')).visible := false;
	if TToolBarButton97(GetControl('BPOPMENU')) <> nil then TToolBarButton97(GetControl('BPOPMENU')).visible := false;
  if TToolBarButton97(GetControl('BPARAMNATURE')) <> nil then TToolBarButton97(GetControl('BPARAMNATURE')).visible := false;
  if TToolBarButton97(GetControl('BACTUPRIX')) <> nil then TToolBarButton97(GetControl('BACTUPRIX')).visible := false;
  if TToolBarButton97(GetControl('BNOMENC')) <> nil then TToolBarButton97(GetControl('BNOMENC')).visible := false;

	SetControlvisible('GB_PROFIL', False);
	SetControlvisible('GA_ESCOMPTABLE', False);
	SetControlvisible('GA_REMISEPIED', False);
	SetControlvisible('GA_COMMISSIONNABLE', False);

  if (GetparamsocSecur('SO_GCDESACTIVECOMPTA','X') = True) then
  	begin
    SetControlVisible('GA_COMPTAARTICLE', False);
    SetControlVisible('TGA_COMPTAARTICLE', False);
  end;

	if (copy(ecran.name, 1, 13) = 'BTARTPOURCENT') or (copy(ecran.name, 1, 9) = 'BTARTPARC') then
  	begin
		SetControlvisible('TGA_LIBCOMPL', False);
		SetControlvisible('SELTYPEART', False);
  	end
  else
  	begin
  	TTabSheet(ecran.FindComponent('P_CARACTERISTIQUE')).TabVisible := false;
  	TTabSheet(ecran.FindComponent('P_DIVERS')).TabVisible := false;
  	TTabSheet(ecran.FindComponent('P_DIMENSION')).TabVisible := false;
  	TTabSheet(ecran.FindComponent('P_TARIF')).TabVisible := false;
		end;
*}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Si click sur onglet Dimension,
Suite ........ : positionnement du focus sur un item du THDIM
Mots clefs ... : FOCUS;THDIM
*****************************************************************}

procedure TOM_Article.OnChangeOnglet;
var item: THDimensionItem;
begin
  if (GetActiveTabSheet('PAGES').Name = 'P_DIMENSION') and not (DS.State in [dsInsert]) then
  begin
    FocusedFirstItem(THDimension(GetControl('FDIM')));
  end;

{$IFDEF BTP}
   if (GetActiveTabSheet('PAGES').Name = 'TTARIFACH') then
   begin
	  AffichageInfosPrixPose;
   end;
{$ENDIF}

end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Retourne True si un champ de la fiche stock est affiché
Suite ........ : dans le THDIM ( -> NewDimChamp[1..10] )
Mots clefs ... :
*****************************************************************}

function TOM_Article.IsChampStockAffiche: boolean;
var iChamp: integer;
  bExist: boolean;
begin
  bexist := False;
  iChamp := 1;
  while (not bExist) and (iChamp < 11) do
  begin
    bExist := (copy(DimensionsArticle.NewDimChamp[iChamp], 1, 3) = 'GQ_');
    inc(iChamp);
  end;
  Result := bExist;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Affiche la combo Etablissement si un champ de la fiche
Suite ........ : stock est affiché dans le THDIM ( -> permet de visualiser à
Suite ........ : quel établissement se rapporte les données affichées ).
Mots clefs ... :
*****************************************************************}

procedure TOM_Article.AffComboEtab(bForceState: boolean = False);
var CB: THValComboBox;
begin
  if (bForceState or not (DS.State in [DsInsert])) and
    (DimensionsArticle.TypeMasqueDim = VH_GC.BOTypeMasque_Defaut) and (IsChampStockAffiche) then
  begin
    SetControlVisible('_ETABLISS', True);
    SetControlVisible('TETABLISS', True);
    CB := THValComboBox(GetControl('_ETABLISS'));
    if CB <> nil then
    begin
      if CB.Value = '' then CB.Value := Depot;
      if CB.Value = '' then CB.Value := VH_GC.GCDepotDefaut;
    end;
  end else
  begin
    SetControlVisible('_ETABLISS', False);
    SetControlVisible('TETABLISS', False);
  end;
end;

procedure TOM_Article.InitDimensionsArticle;
var detail, natureDoc, naturePiece: string;
begin
  if DimensionsArticle <> nil then DimensionsArticle.free;

  if FicheStock then
  begin
    natureDoc := NAT_STOCK;
    naturePiece := NATPIE_STOCK
  end // Fiche stock
  else if FicheStoVte then
  begin
    natureDoc := NAT_STOVTE;
    naturePiece := NAT_STOVTE
  end // TobViewer stock / Ventes
  else
  begin
    natureDoc := NAT_ARTICLE;
    naturePiece := NATPIE_ARTICLE
  end; // Fiche article
  { DC : Pas d'ajout de '_' : -> caractère spécial SQL !!!
  if ((DuplicFromMenu) or (DuplicFromFiche)) and (Not GetParamsoc('SO_GCNUMARTAUTO'))
      then SetField('GA_CODEARTICLE',GetField('GA_CODEARTICLE')+'_') ;}
  DimensionsArticle := TODimArticle.Create(THDimension(GetControl('FDIM')), GetField('GA_CODEARTICLE'),
    GetField('GA_DIMMASQUE'), '', 'GCDIMCHAMP', natureDoc, naturePiece, Depot, Cloture,
    Boolean(GetField('GA_PRIXUNIQUE') = 'X'), ListeNatPiece, DatePieceDeb, DatePieceFin, bZeroDim);
  DimensionsArticle.Dim.OnVoir := ArticleVoir;
  DimensionsArticle.Dim.OnDblClick := ArticleVoir;
  if (TFFiche(Ecran).FTypeAction <> taConsult) then
  begin
    DimensionsArticle.Dim.OnAction := ArticleActionDim;
    DimensionsArticle.Dim.OnCreateItem := CreateItem;
    DimensionsArticle.Dim.OnDeleteItem := DeleteItem;
    DimensionsArticle.Dim.OnChange := OnChangeItem;
    DimensionsArticle.Dim.PopActionItem.Caption := TraduireMemoire('Mettre à Jour');
    DimensionsArticle.Dim.PopActionItem.visible := True;
  end else
  begin
    DimensionsArticle.Dim.PopUp.Items[2].Visible := False; // Menu Existant invisible
    DimensionsArticle.Dim.PopUp.Items[3].Visible := False; // Menu Inexistant invisible
    // DCA - FQ MODE 10681
    // Rend invisible l'option "Rendre existant" du "collage spéciale" de l'objet dimension (Option disponible en saisie de pièce / saisie de transferts)
    DimensionsArticle.Dim.Options.RendreExistanteVisible := False;
    // Rend invisible l'option "Rendre existant" du "Aller à" de l'objet dimension (Option disponible en saisie de transferts)
    DimensionsArticle.Dim.Options.AllerARendreExistanteVisible := False;
  end;
  DimensionsArticle.Dim.Action := TFFiche(Ecran).FTypeAction;

  {Affichage des champs par défaut de l'utilisateur}
  if DS.State in [DsInsert] then
  begin
    DimensionsArticle.HideUnUsed := False; // Si nouvel article dimensionné, voir détail de la grille.
    // En création, le type de masque est forcé au type de masque par défaut
    if ctxMode in V_PGI.PGIContexte then DimensionsArticle.TypeMasqueDim := VH_GC.BOTypeMasque_Defaut;
  end;
  detail := CheckToString(not DimensionsArticle.HideUnUsed);
  AfficheUserPref(DimensionsArticle, natureDoc, naturePiece);
  SetControlEnabled('CBDETAIL', not (DS.State in [DsInsert]));
  SetControlVisible('CBDETAIL', (natureDoc = NAT_ARTICLE));
  {$IFNDEF BTP}
  SetControlEnabled('BPARAM', not (DS.State in [DsInsert]));
  SetControlVisible('BPARAM', NaturePieceGeree(naturePiece));
  {$ENDIF}
  SetControlText('CBDETAIL', detail);
  AffComboEtab(False);
end;

procedure TOM_Article.OnNewRecord;
var TypeCAB, ArtSpecif: string;
  Cpt, LgNumCAB, Nbr: integer;
  QQ: TQuery;
  {$IFDEF GPAO}
  s: string;
  TypeInitAction: TInitAction;
  {$ENDIF}
begin
  bCalculPrixDuplicHT := False; // DBR pour pas avoir le message 37 la premiere fois
  bCalculPrixDuplicTTC := False;
  if Modifparlot = True then
  begin
    SetField('GA_LIBELLE', TraduireMemoire('Modification par lot d''article'));
    Exit;
  end;
  {$IFDEF GPAO}
  s := ''; 
  TypeInitAction := iaInit;
  if Assigned(DS) then
    wInitFieldsbyDS(DS, 'ARTICLE', TypeInitAction, s)
  else
    wInitFields(fTob, 'ARTICLE', TypeInitAction, s);

  wSetPlusUniteStock;
  {$ENDIF}
  if V_PGI.VersionDemo then
  begin
    QQ := OpenSQL('SELECT COUNT(*) FROM ARTICLE', True);
    if not QQ.EOF then Nbr := QQ.Fields[0].AsInteger else Nbr := 0;
    Ferme(QQ);
    if (Nbr >= 100) then PgiBox(TexteMessage[41], TraduireMemoire('Fiche Article'));
  end;

  ClickActuPrix := True;
  ModifChampHT := False;
  ModifChampTTC := False;
  // Modif BTP
  ModifChampPR := false;

  {$IFDEF GPAOLIGHT}
  SetField('GA_UNITEPROD', 'UNI');
  SetField('GA_UNITECONSO', 'UNI');
  SetField('GA_MODECONSO', 'LAN');
  {$ENDIF GPAOLIGHT}
  // Par défaut
  SetField('GA_GEREANAL', 'X');
  // ---
  if ((TypeArticle = 'NOM') or (TypeArticle = 'OUV')) then
  begin
    //
    // Modif BTP
    //
    {$IFDEF BTP}    
    if TypeNomenc = '' then
      TypeNomenc := CHOISIR(TexteMessage[48], 'COMMUN', 'CO_LIBELLE', 'CO_CODE', 'CO_TYPE="BTN"', '');
    if Typenomenc = '' then TypeNomenc := 'OUV';
    {$ELSE}

    {$IFDEF CCS3}
    TypeNomenc := 'ASS';
    {$ELSE}
    {$IFDEF CHR}
    TypeNomenc := 'MAC';
    {$ELSE}
    TypeNomenc := CHOISIR(TexteMessage[49], 'COMMUN', 'CO_LIBELLE', 'CO_CODE', 'CO_TYPE="GTN"', '');
    {$ENDIF}
    {$ENDIF}

    {$ENDIF} // FIN DU IFDEF BTP

    if TypeNomenc <> '' then SetField('GA_TYPENOMENC', TypeNomenc) else exit;
    end
  {$IFDEF BTP}
  Else if (TypeArticle = 'PRE') then
      Begin    
      end;
  {$ENDIF}

  SetControlEnabled('BNOMENC', False);
  // Modif BTP
  {$IFDEF BTP}
  SetControlEnabled('BNOMENC1', False);
  {$ENDIF}
  SetField('GA_CODEBARRE', '');
  SetField('GA_QUALIFCODEBARRE', '');
  if TypeArticle = '' then SetField('GA_TYPEARTICLE', 'MAR')
  else SetField('GA_TYPEARTICLE', TypeArticle);
  if (((TypeArticle = '') or (TypeArticle = 'MAR')) and (not GetParamsoc('SO_GCNUMCABAUTO')))
    or ((TypeArticle = 'PRE') and (not GetParamsoc('SO_GCPRCABAUTO'))) then
  begin
    // Dans le cas d'une saisie manuelle du code à barres, on propose automatiquement
    // le type ou la qualification du code à barres,si celui-ci est paramétré dans la table CODEBARRES.
    if RechParamCAB_Manu('', TypeCAB, LgNumCAB) = True then
      SetField('GA_QUALIFCODEBARRE', TypeCAB);
  end;
  SetField('GA_FERME', '-');
  SetField('GA_TENUESTOCK', 'X');
  SetField('GA_INVISIBLEWEB', 'X');
  SetField('GA_REMISELIGNE', 'X');
  SetField('GA_REMISEPIED', 'X');
  SetField('GA_ESCOMPTABLE', 'X');
  SetField('GA_FAMILLETAXE1', VH_GC.GCFAMILLETAXE1);
  SetField('GA_PCB', VH_GC.GCPCB);
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    // Gestion spécifique des articles de type PRESTATIONS - Ventes diverses
    if (TypeArticle = 'PRE') then
    begin
      // Reprise des paramètres société : Gestion commerciale / Prestation
      SetField('GA_FAMILLENIV1', GetParamsoc('SO_GCPRFAMNIV1'));
      SetField('GA_FAMILLENIV2', GetParamsoc('SO_GCPRFAMNIV2'));
      SetField('GA_FAMILLENIV3', GetParamsoc('SO_GCPRFAMNIV3'));
      SetField('GA_FAMILLETAXE1', GetParamsoc('SO_GCPRFAMTAXE1'));
      SetField('GA_FAMILLETAXE2', GetParamsoc('SO_GCPRFAMTAXE2'));
      SetField('GA_COMPTAARTICLE', GetParamsoc('SO_GCPRCOMPTAART'));
      SetField('GA_TENUESTOCK', CheckToString(GetParamsoc('SO_GCPRTENUESTOCK')));
      SetField('GA_ESCOMPTABLE', CheckToString(GetParamsoc('SO_GCPRESCOMPTE')));
      SetField('GA_LOT', CheckToString(GetParamsoc('SO_GCPRLOT')));
      SetField('GA_COMMISSIONNABLE', CheckToString(GetParamsoc('SO_GCPRCOMMISSION')));
      SetField('GA_REMISEPIED', CheckToString(GetParamsoc('SO_GCPRREMPIED')));
      SetField('GA_REMISELIGNE', CheckToString(GetParamsoc('SO_GCPRREMLIGNE')));
      //SetField('GA_QUALIFCODEBARRE', GetParamsoc('SO_GCPRQUALIFCAB'));
      SetField('GA_PCB', VH_GC.GCPCB);
      SetControlEnabled('GA_TENUESTOCK', False);
      SetControlEnabled('GA_LOT', False);
    end;
    // Gestion spécifique des articles de type FINANCIER - Opérations de caisses
  end; // DBR - pour prendre en compte aussi dans la GC - Fiche 10755
  if (TypeArticle = 'FI') then
  begin
    SetField('GA_TENUESTOCK', '-');
    SetField('GA_REMISELIGNE', '-');
    SetField('GA_REMISEPIED', '-');
    SetField('GA_ESCOMPTABLE', '-');
    for cpt := 1 to 5 do SetField('GA_FAMILLETAXE' + IntToStr(cpt), '');
    SetField('GA_PCB', '1,00');
  end;
//  end; DBR - pour prendre en compte aussi dans la GC - Fiche 10755
  // SetField('GA_PRIXPOURQTE',VH_GC.GCPRIXPOURQTE) ;
  // Voir MR SetField('GA_PRIXPOURQTE',GetParamSoc('SO_GCPPCB')) ;
  SetField('GA_CREATEUR', V_PGI.User);
  SetField('GA_CREERPAR', 'SAI'); // saisie
  SetField('GA_DATESUPPRESSION', iDate2099);
  //
  // Modif BTP
  //
  {$IFDEF BTP}
  if (typeArticle = 'NOM') or (typeArticle = 'OUV') or (TypeArticle = 'PRE') or
    (TypeArticle = 'CTR') or (TypeArticle = 'POU') or (TypeArticle = 'FRA') or (copy(Typearticle,1,2)='PA') then
    SetField('GA_STATUTART', 'UNI')
  else
    SetField('GA_STATUTART', 'GEN');
  {$ELSE}

  {$IFDEF CCS3}
  SetField('GA_STATUTART', 'UNI');
  {$ELSE}
  if typeArticle = 'NOM' then SetField('GA_STATUTART', 'UNI') else SetField('GA_STATUTART', 'GEN');
  {$ENDIF}

  {$ENDIF}

  if (not DuplicFromFiche) then
  begin
    //
    // Modif BTP
    //
    {$IFDEF BTP}
    SetField('GA_CALCPRIXPR', 'AUC');
    {$ENDIF}
    SetField('GA_CALCPRIXHT', 'AUC');
    SetField('GA_CALCPRIXTTC', 'AUC');
  end;

  {$IFDEF GPAO}
  if (not OnlyProfil) and (GetParamSoc('SO_GCPROFILART') <> '') and (GetParamSoc('SO_GCPROFILART') <> '...') then
  begin
    SetField('GA_PROFILARTICLE', GetParamSoc('SO_GCPROFILART'));
    if GetParamsoc('SO_GCPROFILARTAUTO') then AppliqueProfilArticle(False, True);
  end;
  {$ENDIF}

  /////// Forme ///////////////
  //SetControlEnabled('GA_CODEARTICLE', TRUE);

  //SetControlEnabled('Pages',FALSE) ;
  // gm 18/08/03 attention certaine fiche article "métier" ne contiennent pas certains Control
  if (GetControl ('REMPLACEMENT') <> nil)  then  SetControlText('REMPLACEMENT', '');
  if (GetControl ('SUBSTITUTION') <> nil)  then  SetControlText('SUBSTITUTION', '');

  OldRemplacement := '';
  OldSubstitution := '';
  //
  // Modif BTP 06/02/2001
  //
  if TypeArticle = 'POU' then setfield('GA_PRIXPOURQTE', 100)
  else SetField('GA_PRIXPOURQTE', 1);

  ArtSpecif := IsArticleSpecif(TypeArticle);
  if ArtSpecif = 'FicheBtp' then
  begin
    SetField('GA_TENUESTOCK', '-');
    SetField('GA_PCB', 0);
    SetField('GA_ESCOMPTABLE', 'X');
    SetField('GA_COMMISSIONNABLE', 'X');
    SetField('GA_REMISEPIED', 'X');
    SetField('GA_REMISELIGNE','X');
    SetField('GA_ACTIVITEREPRISE', 'F');
    // Gestion de marge a l'identique que dans la gamme II
    SetField('GA_QUALIFMARGE', 'CO');
    SetControlProperty('REFMARGE', 'ItemIndex', 4);
    // ------------------
    if (TypeArticle = 'PRE') or (TypeArticle = 'CTR') then
    begin
      TcheckBox(GetControl('GA_TENUESTOCK')).Enabled := false;
      SetField('GA_QUALIFUNITEVTE', VH_GC.AFMesureActivite);
      SetField('GA_QUALIFUNITEACT', VH_GC.AFMesureActivite);
    end;
  end;

  // SPECIF AFFAIRE ///////////////////////
  //mcd 26/09/03 ajout test origine; ne pas effacer si on est en duplication
  if (ArtSpecif = 'FicheAffaire') and (not DuplicFromFiche) and (not DuplicFromMenu) then
  begin
    SetField('GA_TENUESTOCK', '-');
    SetField('GA_CALCPRIXHT', 'AUC');
    SetField('GA_CALCPRIXTTC', 'AUC');
    // mcd 06/03/02 ajout test POU ....
    if TypeArticle = 'POU' then setfield('GA_PRIXPOURQTE', 100)
    else SetField('GA_PRIXPOURQTE', 1);
    SetField('GA_PCB', 0);
    // mcd 16/04/02 prise en compte option des paramtre prestation
    SetField('GA_ESCOMPTABLE', CheckToString(GetParamsoc('SO_GCPRESCOMPTE')));
    SetField('GA_COMMISSIONNABLE', CheckToString(GetParamsoc('SO_GCPRCOMMISSION')));
    SetField('GA_REMISEPIED', CheckToString(GetParamsoc('SO_GCPRREMPIED')));
    SetField('GA_REMISELIGNE', CheckToString(GetParamsoc('SO_GCPRREMLIGNE')));
    {      SetField('GA_ESCOMPTABLE','X');
          SetField('GA_COMMISSIONNABLE','X');
          SetField('GA_REMISEPIED','X');     }
    SetField('GA_ACTIVITEREPRISE', 'F');
    if (TypeArticle = 'PRE') or (TypeArticle = 'CTR') then
    begin
      SetField('GA_QUALIFUNITEVTE', VH_GC.AFMesureActivite);
      SetField('GA_QUALIFUNITEACT', VH_GC.AFMesureActivite);
    end;
    SetField('GA_STATUTART', 'UNI'); // pas de dimension pour AFFAIRE  20/06/2000 PL
  end;
  // mcd 09/01/02 l'affectation est faite systématiquement, même si appel fct pas en saisie
  SetField('GA_ACTIVITEEFFECT', CheckToString(TypeArticle = 'PRE'));

  ////////////////////////////////////

  //SPECIF MODE
  if (ArtSpecif = 'FicheModePre') or (ArtSpecif = 'FicheModeFi') then
  begin
    SetField('GA_STATUTART', 'UNI'); // pas de dimension pour article specif MODE
    TypeMasque := VH_GC.BOTypeMasque_Defaut;
  end;
  {$IFDEF GESCOM}
  if (ArtSpecif = 'FicheGescomFi') then
  begin
    SetField('GA_STATUTART', 'UNI');
  end;
  {$ENDIF}
  // Si nouveau depuis DOC ACH et monofournisseur saisie auto fournprinc
  // DCA - FQ MODE 10482 - Ne pas écraser le fournisseur principal du profil article par défaut
  if FournDocAch<>'' then SetField('GA_FOURNPRINC', FournDocAch);
  ////////////////////////////////////
  {$IFDEF GPAO}
  SetControlEnabled('MNNATURETRAVAIL', False);
  SetControlEnabled('MNPrixDeRevient', False);
  SetControlEnabled('BT_ARTNAT', False);
  { Activation du bouton de configuration }
  SetControlEnabled('BCFXARTICLE', True); { GPAO V1.0 }

  // Gestion des Articles Profils - DS le 20/01/02
  SetField('GA_ESTPROFIL', iif(OnlyProfil, 'X', '-'));
  SetControlChecked('GA_ESTPROFIL', OnlyProfil);
  SetControlEnabled('GA_ESTPROFIL', False);
  SetControlChecked('GA_FERME', OnlyProfil);
  {$ENDIF}
  {$IFDEF BTP}
  SetField('GA_CODEARTICLE', prefixe);
  {$ENDIF}
  SetActiveTabSheet('PGeneral');
  SetFocusControl('GA_CODEARTICLE');
  {$IFDEF EAGLCLIENT}
  eAglInsertionOn; // DBR Fiche 10093
  {$ENDIF} // EAGLCLIENT
//uniquement en line
{*
  SetField('GA_CALCPRIXPR', 'PAA');
  SetField('GA_CALCPRIXHT', 'DPR');
  SetField('GA_CALCAUTOHT', 'X');
  SetField('GA_DPRAUTO', 'X');
*}
  //FV1 : 22/05/2017 - FS#2343 - Pb en application du profil article par défaut
  if (TypeArticle = 'MAR') Or (TypeArticle = 'ARP') then
  begin
    if (GetParamSoc('SO_GCPROFILART') <> '') and (GetParamSoc('SO_GCPROFILART') <> '...') then
      // DCA - FQ MODE 11011 Prise en compte du paramètre société "Appliquer automatiquement à la création"
      if GetParamsoc('SO_GCPROFILARTAUTO') then AppliqueProfilArticle(False, True);
  end;

end;

procedure TOM_Article.OnDeleteRecord;
begin

  if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE="' + GetField('GA_ARTICLE') + '"') then
  begin
    SetLastError(17, '');
    exit;
  end;

  // SPECIF AFFAIRE ///////////////////
  if (ctxAffaire in V_PGI.PGIContexte) then
    if ExisteSQL('SELECT ACT_ARTICLE FROM ACTIVITE WHERE ACT_ARTICLE="' + GetField('GA_ARTICLE') + '"') then
    begin
      SetLastError(31, '');
      exit;
    end;
  /////////////////////////////////////

  ////// Mise à jour catalogue fournisseur /////////
  ExecuteSQL('UPDATE CATALOGU SET GCA_ARTICLE="" WHERE GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '"');
  ExecuteSQL('DELETE FROM ARTICLELIE WHERE GAL_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"');
  ExecuteSQL('DELETE FROM CONDITIONNEMENT WHERE GCO_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"');
  ExecuteSQL('DELETE FROM ARTICLEPIECE WHERE GAP_ARTICLE="' + string(GetField('GA_ARTICLE')) + '"');
  {$IFDEF CHR}
  ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_NOMENCLATURE="' + string(GetField('GA_ARTICLE')) + '"');
  ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="' + string(GetField('GA_ARTICLE')) + '"');
  {$ENDIF}

  // DCA - FQ MODE 10793
  if GetControl ('REMPLACEMENT') <> nil then
    begin
    SetControlText('REMPLACEMENT', '');
    THLabel(GetControl('TREMPLACEMENT')).Caption := '';
    end;
  if GetControl ('SUBSTITUTION') <> nil then
    begin
    SetControlText('SUBSTITUTION', '');
    THLabel(GetControl('TSUBSTITUTION')).Caption := '';
    end;

  if TobZonelibre <> nil then
  begin
    TobZonelibre.free;
    TobZonelibre := nil
  end;

  // Gestion des métrés
  if MetreArticle <> nil then
  begin
    //Fermeture d'excel
    MetreArticle.CloseXLS;
    //Suppression du fichier dans répertoire poste de Travail
    MetreArticle.SupprimeFichierXLS(RepPoste);
    //suppression du fichier dans répertoire serveur
    MetreArticle.SupprimeFichierXLS(RepServeur);
    FreeAndNil(MetreArticle);
  end;

end;

procedure TOM_Article.OnCancelRecord;

begin
  inherited;
  //Specifique Mode pour gérer la table ARTICLECOMPL
  if (CtxMode in V_PGI.PGIContexte) and ((GetPresentation = ART_ORLI) or (GetField('GA_TYPEARTICLE') = 'FI')) then
    if TFfiche(Ecran) <> nil then TobZonelibre.PutEcran(TFfiche(Ecran));
end;

procedure TOM_Article.OnClose;
{$IFDEF BTP}
var  NbDeclinaison : integer;
     FichierArticle: string;
{$ENDIF}
begin
  // Si l'on est dans la fiche dimensionné et pas de navigation,
  // retour de la TOB article dimensionné (pour maj THDIM)
  // sinon, en retour a la fiche generique, il faudra recharger entierement THDIM.
  if (FicheDimensionne) and (not NavigClick) then TheTob := TobArt else TheTob := nil;
  if TobArtLoad <> nil then
  begin
    TobArtLoad.Free;
    TobArtLoad := nil
  end;
  if not FicheDimensionne then
  begin
    if TobChampModif <> nil then
    begin
      TobChampModif.Free;
      TobChampModif := nil
    end;
    if (ECRAN.name <> 'GCARTICLEDIM') and (TobArt <> nil) then
    begin
      TobArt.Free;
      TobArt := nil
    end;
  end;
  if (GetField('GA_STATUTART') = 'GEN') and (DimensionsArticle <> nil) then DimensionsArticle.Destroy;
  if TobZonelibre <> nil then TobZonelibre.free;
  if (copy(Ecran.Name, 1, 9) = 'GCARTICLE') then
  begin
    if SpecifGescom <> nil then SpecifGescom.Free;
  end;
  //Detruit la TOB contenant le profil article
  if TOBInfoArt <> nil then
  begin
    TOBInfoArt.Free;
    TOBInfoArt := nil;
  end;
  ArticleFourn := False;
  {$IFDEF GPAO}
  if (TobUnAutorizedFields <> nil) then TobUnAutorizedFields.Free;
  {$ENDIF}
  //
  {$IFDEF BTP}
  if TOBTarif <> nil then TOBTarif.free;
  if TOBDetail <> nil then TOBDetail.free;
  if TOBCatalog <> nil then TOBCatalog.free;
  if TOBEDetail <> nil then TOBEDetail.free;
  if TOBPrestation <> nil then TOBPrestation.free;
  if TOBTiers <> nil then TOBTiers.free;

  // Gestion des métrés
  if MetreArticle <> nil then
  begin
    FichierArticle := RepPoste;
    DeleteFichier(FichierArticle);
    FreeAndNil(MetreArticle);
  end;
  // ecran des articles prix posés
  {$ENDIF}

  if TobLiensOle <> nil then TobLiensOle.Free;
  if assigned(TOBARTICLEPARC) then freeAndNil (TOBARTICLEPARC);
end;

//si modif dans cette fonction : penser à MAJ la function CreerArticle() dans TOF_ARTPHOTO

procedure TOM_Article.AppliqueProfilArticle(Confirm: boolean; NewRecord: boolean = False);
var
  QQ: TQuery;
  SQL: String;
  {$IFNDEF GPAO}
    Niv1, Niv2, Niv3: string;
  {$ELSE}
    CodeProfil: string;
    RangeProfil: string;
  {$ENDIF}
  Profil: string;
  ProfilMul, IsAucOuStock, Abandon : boolean;
  AppliqueProfil: boolean;
  SelectProfil: string;
  i: integer;
begin
  Profil := '';
  {$IFDEF GPAO}
  if (Ecran.Name = 'WARTICLE_FIC') and (GetField('GA_PROFILARTICLE') <> '') then
    SQL := 'SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + GetField('GA_PROFILARTICLE') + '"'
  else
  begin
    RangeProfil := '';
    if Getfield('GC_FAMILLENIV1') <> '' then
      RangeProfil := RangeProfil + 'GPF_FAMILLENIV1=' + Getfield('GA_FAMILLENIV1');

    if Getfield('GC_FAMILLENIV2') <> '' then
    begin
      if length(RangeProfil) > 0 then RangeProfil := RangeProfil + ';';
      RangeProfil := RangeProfil + 'GPF_FAMILLENIV2=' + Getfield('GA_FAMILLENIV2');
    end;

    if Getfield('GC_FAMILLENIV3') <> '' then
    begin
      if length(RangeProfil) > 0 then RangeProfil := RangeProfil + ';';
      RangeProfil := RangeProfil + 'GPF_FAMILLENIV3=' + Getfield('GA_FAMILLENIV3');
    end;

    CodeProfil := AglLanceFiche('W', 'WPROFILART_MUL', RangeProfil, '', '');
    if CodeProfil <> '' then
      SQL := 'SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + CodeProfil + '"'
    else
      EXIT; { Pas de CodeProfil sélectionné = on dégage }
  end;
  {$ELSE}
  // Modification Marché Mode OT
  if NewRecord then
  begin
      SQL := 'SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + GetParamSoc('SO_GCPROFILART') + '"';
  end else
  begin
    Niv1 := GetField('GA_FAMILLENIV1');
    Niv2 := GetField('GA_FAMILLENIV2');
    Niv3 := GetField('GA_FAMILLENIV3');
    SQL := 'SELECT * FROM PROFILART';
    if Niv1 <> '' then
      SQL := SQL + ' WHERE GPF_FAMILLENIV1="' + Niv1 + '"';
    if (Niv2 <> '') then
      if (SQL <> '') then
        SQL := SQL + ' AND GPF_FAMILLENIV2="' + Niv2 + '"' else
        SQL := SQL + ' WHERE GPF_FAMILLENIV2="' + Niv2 + '"';
    if (Niv3 <> '') then
      if (SQL <> '') then
        SQL := SQL + ' AND GPF_FAMILLENIV3="' + Niv3 + '"' else
        SQL := SQL + ' WHERE GPF_FAMILLENIV3="' + Niv3 + '"';
    SQL := SQL + ' ORDER BY  GPF_FAMILLENIV2 DESC, GPF_FAMILLENIV3 DESC';
  end;
{$ENDIF}

  QQ := OpenSQL(SQL, TRUE);

  if CtxMode in V_PGI.PGIContexte then
  begin
    if QQ.RecordCount > 1 then
    begin
      ferme(QQ);
      if Getfield('GC_FAMILLENIV1') <> '' then
        profil := profil + 'GPF_FAMILLENIV1=' + Getfield('GA_FAMILLENIV1');

      if Getfield('GC_FAMILLENIV2') <> '' then
      begin
        if length(profil) > 0 then profil := profil + ';';
        profil := profil + 'GPF_FAMILLENIV2=' + Getfield('GA_FAMILLENIV2');
      end;

      if Getfield('GC_FAMILLENIV3') <> '' then
      begin
        if length(profil) > 0 then profil := profil + ';';
        profil := profil + 'GPF_FAMILLENIV3=' + Getfield('GA_FAMILLENIV3');
      end;

      QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + AglLanceFiche('GC', 'GCPROFILART_MUL', profil, '', '') + '"', true);
    end;
  end;

  ProfilMul := False;
  Abandon := False;
  SelectProfil := '';
  if QQ.EOF then
  begin
    ferme(QQ);
    {$IFNDEF GPAO}
    QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + AglLanceFiche('GC', 'GCPROFILART_MUL', '', '', '') + '"', true);
    {$ELSE}
    SelectProfil := AglLanceFiche('W', 'WPROFILART_MUL', '', '', '');
    if SelectProfil <> '' then
      QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + SelectProfil + '"', True);
    {$ENDIF}
    ProfilMul := True;
  end else
  begin
    if Confirm then
      Case PGIAskCancel('Appliquer le profil ' + QQ.Findfield('GPF_LIBELLE').AsString + ' ?','Profil article') of
        mrNo : begin
                 ferme(QQ);
                 QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + AglLanceFiche('GC', 'GCPROFILART_MUL', '', '', '') + '"', true);
                 ProfilMul := True;
               end;
        mrCancel : Abandon := True;
      end ;
  end;

  if not QQ.EOF then
  begin
    if not Abandon then
    begin
      AppliqueProfil := True;

      {$IFDEF GPAO}
      if QQ.FindField('GPF_TYPEPROFILART').AsString = 'ADV' then { Profil avancé }
      begin
        if wSetAdvancedProfilToArticle(DS, QQ) then
        begin
          { Mise à jour du code profil dans l'article }
          if GetField('GA_PROFILARTICLE') = '' then
            SetField('GA_PROFILARTICLE', QQ.Findfield('GPF_PROFILARTICLE').AsString);
        end;
        AppliqueProfil := False;
      end;
      {$ENDIF}

      if AppliqueProfil then
      begin
        if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
        SetField('GA_COMPTAARTICLE',QQ.Findfield('GPF_COMPTAARTICLE').AsString);
        SetField('GA_TENUESTOCK', QQ.Findfield('GPF_TENUESTOCK').AsString);
        {$IFDEF CCS3}
        SetField('GA_LOT', '-');
        SetField('GA_NUMEROSERIE', '-');
        {$ELSE}
        SetField('GA_LOT', QQ.Findfield('GPF_LOT').AsString);
        SetField('GA_NUMEROSERIE', QQ.Findfield('GPF_NUMEROSERIE').AsString);
        {$ENDIF}
        SetField('GA_CONTREMARQUE', QQ.Findfield('GPF_CONTREMARQUE').AsString);
        SetField('GA_REMISEPIED', QQ.Findfield('GPF_REMISEPIED').AsString);
        SetField('GA_REMISELIGNE', QQ.Findfield('GPF_REMISELIGNE').AsString);
        SetField('GA_ESCOMPTABLE', QQ.Findfield('GPF_ESCOMPTABLE').AsString);
        SetField('GA_FAMILLETAXE1', QQ.Findfield('GPF_CODETAXE').AsString); //AC fiche bug 343
        SetField('GA_COMMISSIONNABLE', QQ.Findfield('GPF_COMMISSIONNABL').AsString);
        // mise à jour des champs famille si lorsque l'on vient du mul
        if ProfilMul or NewRecord then
        begin
          SetField('GA_FAMILLENIV1', QQ.Findfield('GPF_FAMILLENIV1').AsString);
          SetField('GA_FAMILLENIV2', QQ.Findfield('GPF_FAMILLENIV2').AsString);
          SetField('GA_FAMILLENIV3', QQ.Findfield('GPF_FAMILLENIV3').AsString);
        end;
        SetField('GA_CALCPRIXHT', QQ.Findfield('GPF_CALCPRIXHT').AsString);
        SetField('GA_CALCPRIXTTC', QQ.Findfield('GPF_CALCPRIXTTC').AsString);

        // le test est a remettre en cause si correction faite dans l'AGL
        if QQ.FindField('GPF_COEFCALCHT').AsString = '' then SetField('GA_COEFCALCHT', '0')
        else SetField('GA_COEFCALCHT', QQ.Findfield('GPF_COEFCALCHT').AsString);

        if QQ.FindField('GPF_COEFCALCTTC').AsString = '' then
          SetField('GA_COEFCALCTTC', '0')
        else
          SetField('GA_COEFCALCTTC',QQ.Findfield('GPF_COEFCALCTTC').AsString);
        //
        SetField('GA_CALCAUTOHT', QQ.Findfield('GPF_CALCAUTOHT').AsString);
        SetField('GA_CALCAUTOTTC', QQ.Findfield('GPF_CALCAUTOTTC').AsString);
        SetField('GA_TARIFARTICLE', QQ.Findfield('GPF_TARIFARTICLE').AsString);
        SetField('GA_PAYSORIGINE', QQ.Findfield('GPF_PAYSORIGINE').AsString);
        SetField('GA_COEFFG', QQ.Findfield('GPF_COEFCALCPR').AsFloat);
        //FV1 : 22/05/2017 -  FS#2343 - Pb en application du profil article par défaut
        SetField('GA_CALCPRIXPR', QQ.Findfield('GPF_CALCPRIXPR').AsString);
        SetField('GA_DPRAUTO', QQ.Findfield('GPF_CALCAUTOPR').AsString);
        //
        SetField('GA_ARRONDIPRIX', QQ.Findfield('GPF_ARRONDIPRIX').AsString);
        SetField('GA_ARRONDIPRIXTTC', QQ.Findfield('GPF_ARRONDIPRIXTTC').AsString);
        SetField('GA_PRIXUNIQUE', QQ.Findfield('GPF_PRIXUNIQUE').AsString);
        //JD
        SetField('GA_COLLECTION', QQ.Findfield('GPF_COLLECTION').AsString);
        SetField('GA_FOURNPRINC', QQ.Findfield('GPF_FOURNPRINC').AsString);

        //FV1 : 06/06/2017 - FS#2343 - Pb en application du profil article par défaut
        CodeTarifArticleChange(Self);
        //CodeSousTarif.value := QQ.Findfield('GPF_SOUSFAMTARART').AsString;
        if QQ.Findfield('GPF_SOUSFAMTARART').AsString = '' then
          CodeSousTarif.text := '<<Aucun>>'
        else
          CodeSousTarif.text := QQ.Findfield('GPF_SOUSFAMTARART').AsString;
        SetField('GA2_SOUSFAMTARART', CodeSousTarif.value);
        //
        for i := 1 to $A do SetField('GA_LIBREART' + format('%x', [i]), QQ.Findfield('GPF_LIBREART' + format('%x', [i])).AsString);
        for i := 1 to 3 do SetField('GA_VALLIBRE' + IntToStr(i), QQ.Findfield('GPF_VALLIBRE' + IntToStr(i)).AsFloat);
        for i := 1 to 3 do SetField('GA_DATELIBRE' + IntToStr(i), QQ.Findfield('GPF_DATELIBRE' + IntToStr(i)).AsDateTime);
        for i := 1 to 3 do SetField('GA_CHARLIBRE' + IntToStr(i), QQ.Findfield('GPF_CHARLIBRE' + IntToStr(i)).AsString);
        for i := 1 to 3 do SetField('GA_BOOLLIBRE' + IntToStr(i), QQ.Findfield('GPF_BOOLLIBRE' + IntToStr(i)).AsString);
        { Init des champs GPAO }
        if (CtxGPAO in V_PGI.PGIContexte) then
        begin
          { Onglet PRODUCTION}
          SetField('GA_UNITEPROD', QQ.FindField('GPF_UNITEPROD').AsString);
          SetField('GA_QECOPROD', QQ.FindField('GPF_QECOPROD').AsString);
          SetField('GA_QPCBPROD', QQ.FindField('GPF_QPCBPROD').AsString);
          SetField('GA_DELAIPROD', QQ.FindField('GPF_DELAIPROD').AsString);
          SetField('GA_METHPROD', QQ.FindField('GPF_METHPROD').AsString);
          SetField('GA_UNITECONSO', QQ.FindField('GPF_UNITECONSO').AsString);
          SetField('GA_MODECONSO', QQ.FindField('GPF_MODECONSO').AsString);
          SetField('GA_PERTEPROP', QQ.FindField('GPF_PERTEPROP').AsString);
          SetField('GA_CODEFORME', QQ.FindField('GPF_CODEFORME').AsString);
          { Onglet GENERAL}
          SetField('GA_EAN', QQ.FindField('GPF_EAN').AsString);
        end;
      end;
    end;
  end
  else
  begin
    if (SelectProfil <> '') then
      HShowMessage('0;' + 'Profil article' + ';Aucun profil correspondant;E;O;O;O', '', '');
  end;
  Ferme(QQ);

  if (GetField('GA_CALCPRIXHT') = 'AUC') or (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') then IsAucOuStock := False else
    IsAucOuStock := True; // si aucun ou stock
  //  BBI, fiche correction 10410
  SetControlEnabled('GA_COEFCALCHT',(IsAucOuStock and Not Consultation)) ;
  SetControlEnabled('GA_CALCAUTOHT',(IsAucOuStock and Not Consultation)) ;
  //  BBI, fin fiche correction 10410
  if (GetField('GA_CALCPRIXTTC') = 'AUC') or (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S') then IsAucOuStock := False else
    IsAucOuStock := True; // si aucun ou stock
  SetControlEnabled('GA_COEFCALCTTC', (IsAucOuStock and not Consultation));
  SetControlEnabled('GA_CALCAUTOTTC', (IsAucOuStock and not Consultation));
end;

{$IFDEF BTP}
//si modif dans cette fonction : penser à MAJ la function CreerArticle() dans TOF_ARTPHOTO

procedure TOM_Article.AppliqueProfilBTP(Confirm: boolean);
var QQ: TQuery;
  SQL, Requete: string;
  Profil: string;
  ProfilMul, IsAucOuStock: boolean;
  i: integer;
begin
  if GetField('GA_TYPEARTICLE') <> 'PRE' then
  begin
    SQL := 'SELECT * FROM PROFILART WHERE GPF_FAMILLENIV1="' + GetField('GA_FAMILLENIV1') + '"';
    SQL := SQL + ' AND (GPF_FAMILLENIV2="' + GetField('GA_FAMILLENIV2') + '"' + ' OR GPF_FAMILLENIV2="") ';
    SQL := SQL + ' AND (GPF_FAMILLENIV3="' + GetField('GA_FAMILLENIV3') + '"' + ' OR GPF_FAMILLENIV3="") ';
    SQL := SQL + ' AND (GPF_FAMILLENIV2<>"" OR GPF_FAMILLENIV3="") ';
    SQL := SQL + ' ORDER BY  GPF_FAMILLENIV2 DESC, GPF_FAMILLENIV3 DESC';
    QQ := OpenSQL(SQL, TRUE);

    //Traitement des familles hierarchiques
    if QQ.RecordCount > 1 then
    begin
      ferme(QQ);
      if Getfield('GC_FAMILLENIV1') <> '' then
        profil := profil + 'GPF_FAMILLENIV1=' + Getfield('GA_FAMILLENIV1');
      if Getfield('GC_FAMILLENIV2') <> '' then
      begin
        if length(profil) > 0 then profil := profil + ';';
        profil := profil + 'GPF_FAMILLENIV2=' + Getfield('GA_FAMILLENIV2');
      end;

      if Getfield('GC_FAMILLENIV3') <> '' then
      begin
        if length(profil) > 0 then profil := profil + ';';
        profil := profil + 'GPF_FAMILLENIV3=' + Getfield('GA_FAMILLENIV3');
      end;

      QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + AglLanceFiche('BTP', 'BTPROFILART_MUL', profil, '', '') + '"', true);
    end;

    ProfilMul := False;

    if QQ.EOF then
    begin
      ferme(QQ);
      QQ := OpenSQL('SELECT * FROM PROFILART WHERE GPF_PROFILARTICLE="' + AglLanceFiche('BTP', 'BTPROFILART_MUL', '', '', '') + '"', true);
      ProfilMul := True;
    end;

    if not QQ.EOF then
    begin
      if (not Confirm) or (Confirm and (HShowMessage('0;' + 'Profil article' + ';Appliquer le profil ' + QQ.Findfield('GPF_LIBELLE').AsString + ' ? ;E;OC;O;C',
        '', '') = mrOk)) then
      begin
        if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
        SetField('GA_COMPTAARTICLE', QQ.Findfield('GPF_COMPTAARTICLE').AsString);
        SetField('GA_TENUESTOCK', QQ.Findfield('GPF_TENUESTOCK').AsString);
        {$IFDEF CCS3}
        SetField('GA_LOT', '-');
        SetField('GA_NUMEROSERIE', '-');
        {$ELSE}
        SetField('GA_LOT', QQ.Findfield('GPF_LOT').AsString);
        SetField('GA_NUMEROSERIE', QQ.Findfield('GPF_NUMEROSERIE').AsString);
        {$ENDIF}
        SetField('GA_CONTREMARQUE', QQ.Findfield('GPF_CONTREMARQUE').AsString);
        SetField('GA_REMISEPIED', QQ.Findfield('GPF_REMISEPIED').AsString);
        SetField('GA_REMISELIGNE', QQ.Findfield('GPF_REMISELIGNE').AsString);
        SetField('GA_ESCOMPTABLE', QQ.Findfield('GPF_ESCOMPTABLE').AsString);
        SetField('GA_FAMILLETAXE1', QQ.Findfield('GPF_CODETAXE').AsString); //AC fiche bug 343
        SetField('GA_COMMISSIONNABLE', QQ.Findfield('GPF_COMMISSIONNABL').AsString);
        // mise à jour des champs famille si lorsque l'on vient du mul
(*        if ProfilMul then
        begin *) // Correction suite a fiche de transmission
          SetField('GA_FAMILLENIV1', QQ.Findfield('GPF_FAMILLENIV1').AsString);
          SetField('GA_FAMILLENIV2', QQ.Findfield('GPF_FAMILLENIV2').AsString);
          SetField('GA_FAMILLENIV3', QQ.Findfield('GPF_FAMILLENIV3').AsString);
(*        end; *)
        SetField('GA_CALCPRIXHT', QQ.Findfield('GPF_CALCPRIXHT').AsString);
        SetField('GA_CALCPRIXTTC', QQ.Findfield('GPF_CALCPRIXTTC').AsString);
        if QQ.FindField('GPF_COEFCALCHT').AsString = '' then SetField('GA_COEFCALCHT', '0')
        else SetField('GA_COEFCALCHT', QQ.Findfield('GPF_COEFCALCHT').AsString);
        if QQ.FindField('GPF_COEFCALCTTC').AsString = '' then SetField('GA_COEFCALCTTC', '0')
        else SetField('GA_COEFCALCTTC', QQ.Findfield('GPF_COEFCALCTTC').AsString);
        SetField('GA_CALCAUTOHT', QQ.Findfield('GPF_CALCAUTOHT').AsString);
        SetField('GA_CALCAUTOTTC', QQ.Findfield('GPF_CALCAUTOTTC').AsString);
        SetField('GA_TARIFARTICLE', QQ.Findfield('GPF_TARIFARTICLE').AsString);

        //FV1 : 06/06/2017 - FS#2343 - Pb en application du profil article par défaut
        CodeTarifArticleChange(Self);
        CodeSousTarif.value := QQ.Findfield('GPF_SOUSFAMTARART').AsString;
        SetField('GA2_SOUSFAMTARART', CodeSousTarif.value);

        SetField('GA_PAYSORIGINE', QQ.Findfield('GPF_PAYSORIGINE').AsString);
        SetField('GA_CALCPRIXPR', QQ.Findfield('GPF_CALCPRIXPR').AsString);
        if QQ.FindField('GPF_COEFCALCPR').AsString = '' then SetField('GA_COEFFG', '0')
        else SetField('GA_COEFFG', QQ.Findfield('GPF_COEFCALCPR').AsString);
        SetField('GA_DPRAUTO', QQ.Findfield('GPF_CALCAUTOPR').AsString);
        ProfilMul := False;
      end;
    end else
      HShowMessage('0;' + 'Profil article' + ';Aucun profil correspondant;E;O;O;O', '', '');
    Ferme(QQ);
  end
  else if getfield('GA_NATUREPRES') <> '' then
  begin
    Requete := 'Select * FROM NATUREPREST WHERE BNP_NATUREPRES="' + Getfield('GA_NATUREPRES') + '"';
    QQ := OpenSQL(requete, TRUE);
    if not QQ.EOF then
    begin
      if (HShowMessage('0;' + 'Nature de prestation' + ';Appliquer la nature ' + QQ.Findfield('BNP_LIBELLE').AsString + ' ? ;E;OC;O;C', '', '') = mrOk) then
      begin
        THLabel(GetControl('TGA_LIBNATURE')).caption := QQ.findfield('BNP_LIBELLE').AsString;
        Setfield('GA_FAMILLETAXE1', QQ.findfield('BNP_CODETAXE').AsString);
        Setfield('GA_COMPTAARTICLE', QQ.findfield('BNP_COMPTAARTICLE').AsString);
        Setfield('GA_COMMISSIONNABLE', QQ.findfield('BNP_COMMISSIONNABL').AsString);
        Setfield('GA_ESCOMPTABLE', QQ.findfield('BNP_ESCOMPTABLE').AsString);
        Setfield('GA_REMISELIGNE', QQ.findfield('BNP_REMISELIGNE').AsString);
        Setfield('GA_REMISEPIED', QQ.findfield('BNP_REMISEPIED').AsString);
        // PV HT
        Setfield('GA_CALCPRIXHT', QQ.findfield('BNP_CALCPRIXHT').AsString);
        Setfield('GA_COEFCALCHT', QQ.findfield('BNP_COEFCALCHT').AsFloat);
        Setfield('GA_CALCAUTOHT', QQ.findfield('BNP_CALCAUTOHT').AsString);
        // PR
        Setfield('GA_CALCPRIXPR', QQ.findfield('BNP_CALCPRIXPR').AsString);
        Setfield('GA_COEFFG', QQ.findfield('BNP_COEFCALCPR').AsFloat);
        Setfield('GA_DPRAUTO', QQ.findfield('BNP_CALCAUTOPR').AsString);
        // PV TTC
        Setfield('GA_CALCPRIXTTC', QQ.findfield('BNP_CALCPRIXTTC').AsString);
        Setfield('GA_COEFCALCTTC', QQ.findfield('BNP_COEFCALCTTC').AsFloat);
        Setfield('GA_CALCAUTOTTC', QQ.findfield('BNP_CALCAUTOTTC').AsString);
      end;
    end;
    ferme(QQ);
  end;

  if (GetField('GA_CALCPRIXHT') = 'AUC') or (copy(GetField('GA_CALCPRIXHT'), 1, 1) = 'S') then IsAucOuStock := False else
    IsAucOuStock := True; // si aucun ou stock
  SetControlEnabled('GA_COEFCALCHT',(IsAucOuStock and Not Consultation)) ;
  SetControlEnabled('GA_CALCAUTOHT',(IsAucOuStock and Not Consultation)) ;
  if (GetField('GA_CALCPRIXTTC') = 'AUC') or (copy(GetField('GA_CALCPRIXTTC'), 1, 1) = 'S') then IsAucOuStock := False else
    IsAucOuStock := True; // si aucun ou stock
  SetControlEnabled('GA_COEFCALCTTC', (IsAucOuStock and not Consultation));
  SetControlEnabled('GA_CALCAUTOTTC', (IsAucOuStock and not Consultation));
end;
{$ENDIF}

procedure TOM_Article.EditArtLie(TypArtLie: string);
begin
  if (TypArtLie = 'REM') and (CodeArticleUnique2(GetControlText('REMPLACEMENT'), '') <> OldRemplacement) then
    if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
  if (TypArtLie = 'SUB') and (CodeArticleUnique2(GetControlText('SUBSTITUTION'), '') <> OldSubstitution) then
    if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
end;

procedure TOM_Article.Duplication(CleDuplic: string);
var TobForm: TOB;
  i: integer;
  StTable, StCleDuplic: string;
  CC: Tcontrol;
  G_CodeArticle: THCritMaskEdit;
  CC_RO, bZoneLibreOk: Boolean;
  Sql: string;
  QQ: TQuery;
begin
  DuplicFromFiche := True;
  StCleDuplic := CleDuplic;
  DuplicArt := StCleDuplic;
  // StTable := GetTableNameFromDataSet(DS);
  StTable := 'ARTICLE';
  TobForm := TOB.Create(StTable, nil, -1);
  if TobForm <> nil then
  begin
    bZoneLibreOk := False;
    if STCleDuplic = '' then
    begin
      if (DS.State = dsInsert) then // si nouvelle fiche  recherche record à dupliquer
      begin
        G_CodeArticle := THCritMaskEdit.Create(nil);
        DispatchRecherche(G_CodeArticle, 1, '', '', '');
        StCleDuplic := G_CodeArticle.Text;
        DuplicArt := StCleDuplic;
        if not TobForm.SelectDB('"' + StCleDuplic + '"', nil) then StCleDuplic := '';
        G_CodeArticle.Free;
      end else
      begin
        { Solution non satisfaisante :
          problématique = quid des modifs non enregistrées avant duplication ?
          Bouge() demande confirmation de maj alors que TobForm déjà chargée !
          L'ajout d'un GetEcran ne résout pas le pb si les modifs ne sont pas enreg.}
        {
        if not TobForm.SelectDB ('',TFFiche(Ecran).QFiche) then StCleDuplic:='' ;
        stCleDuplic:='*' ;
        TobForm.GetEcran(Ecran) ;  // Maj des champs modifiés non enregistrés avant duplication
        TFFiche (Ecran).Bouge (NbInsert);
        }
        stCleDuplic := GetField('GA_ARTICLE');
        DuplicArt := stCleDuplic;
        TFFiche(Ecran).Bouge(NbInsert);
        if not TobForm.SelectDB('"' + stCleDuplic + '"', nil) then StCleDuplic := '';
        if GetPresentation = ART_ORLI then
        begin // Chargement de ARTICLE et ARTICLECOMPL !
          if not TobZonelibre.SelectDB('"' + stCleDuplic + '"', nil) then TobZonelibre.InitValeurs;
          bZoneLibreOk := True;
        end;
      end;
    end else
    begin
      if not TobForm.SelectDB('"' + StCleDuplic + '"', nil) then StCleDuplic := '';
    end;
    if StCleDuplic <> '' then
    begin
      for i := 1 to TobForm.NbChamps do
      begin
        CC := TControl(Ecran.findcomponent(TobForm.GetNomChamp(i)));
        {$IFDEF EAGLCLIENT}
        CC_RO := ((CC is TEdit) and (TEdit(CC).ReadOnly));
        {$ELSE}
        CC_RO := (((CC is TEdit) and (TEdit(CC).ReadOnly)) or ((CC is TDBEdit) and (TDBEdit(CC).ReadOnly)));
        {$ENDIF}
        {Duplication de tous les champs
         sauf des champs GA_DPA, GA_PMAP, GA_DPR, GA_PMRP,
         sauf du code barre ;
         Cas du champ GA_CODEARTICLE : il est dupliqué si pas de chrono article automatique.}
        if (CC <> nil) and (CC.enabled and (not CC_RO) and cc.Visible) and
{$IFDEF BTP}
          (TobForm.GetNomChamp(i) <> 'GA_CODEARTICLE') and
{$ELSE}
          ((not GetParamsoc('SO_GCNUMARTAUTO')) or (TobForm.GetNomChamp(i) <> 'GA_CODEARTICLE')) and
{$ENDIF}
          (TobForm.GetNomChamp(i) <> 'GA_CODEBARRE') then
        begin
          if (TobForm.GetNomChamp(i) <> 'GA_DPA') and
            (TobForm.GetNomChamp(i) <> 'GA_DPR') and
            (TobForm.GetNomChamp(i) <> 'GA_PMAP') and
            (TobForm.GetNomChamp(i) <> 'GA_PMRP')
            then SetField(TobForm.GetNomChamp(i), TobForm.GetValeur(i))
          else if (TobForm.GetNomChamp(i) = 'GA_PMRP') and (Ecran.name='AFARTICLE')
                 then SetField(TobForm.GetNomChamp(i), TobForm.GetValeur(i)) // mcd 26/09/03 dans la saisie GI/GA des prestation, ce champ n'est pas a remettre a zéro
                 else SetField(TobForm.GetNomChamp(i), 0.0);
        end;
      end;
      if GetPresentation = ART_ORLI then
      begin
        if not bZoneLibreOk then
        begin
          if not TobZonelibre.SelectDB('"' + GetField('GA_ARTICLE') + '"', nil) then TobZonelibre.InitValeurs;
          //if not TobZonelibre.SelectDB('"'+stCleDuplic+'"',Nil) then TobZonelibre.InitValeurs ;
        end;
        TobZonelibre.PutEcran(TFfiche(Ecran));
        TobZoneLibre.SetAllModifie(True);
      end;

      // Duplication des photos et mémos
      SQL := 'select * from LIENSOLE where LO_TABLEBLOB="GA" and LO_IDENTIFIANT="' + StCleDuplic + '"';
      QQ := OpenSQL(SQL, True);
      TobLiensOle := TOB.Create('LIENSOLE', nil, -1);
      TobLiensOle.LoadDetailDB('LIENSOLE', '', '', QQ, False);
      Ferme(QQ);
    end;
  end;

  TobForm.Free;

  SetActiveTabSheet(TFFiche(Ecran).pages.Pages[0].name);


  //FV1 - 25/06/2015 - FS#1535 - SGTA (NCN) - problème de duplication des prestations à l'intérieur de la fiche
  if GetParamsoc('SO_GCNUMARTAUTO') then
    SetFocusControl('GA_LIBELLE')
  else
  begin
    SetControlEnabled('GA_CODEARTICLE',True);
    SetFocusControl('GA_CODEARTICLE');
  end;

  // surcharge
  // On conserve le masque de l'article dupliqué SetField ('GA_DIMMASQUE', '') ;
  SetField('GA_CREATEUR', V_PGI.User);
  SetField('GA_UTILISATEUR', V_PGI.User);
  SetField('GA_CREERPAR', 'SAI'); // saisie
  SetField('GA_DATECREATION', Date);
  SetField('GA_DATEMODIF', Date);
  SetField('GA_DATESUPPRESSION', iDate2099);
  if not (CtxMode in V_PGI.PGIContexte) then SetField('GA_FOURNPRINC', '');
  DuplicFromFiche := False;
  DuplicFromMenu := False; //JS 29102003
  {$IFDEF CHR}
  if GetField('GA_CALCPRIXHT') = '' then
    SetField('GA_CALCPRIXHT', 'AUC');
  if GetField('GA_CALCPRIXTTC') = '' then
    SetField('GA_CALCPRIXTTC', 'AUC');
  if GetField('GA_CALCPRIXRT') = '' then
    SetField('GA_CALCPRIXRT', 'AUC');
  {$ENDIF}
end;

//
// Fonction contrôlant la modification du fournisseur dans la fiche Article
//

function TOM_Article.ControleModifFournisseur(AffBox: boolean): boolean;
var QQ: TQuery;
begin
  result := True;

  if BMonoFournisseur = True then
  begin
    if (OldFournPrinc <> GetField('GA_FOURNPRINC')) and (OldFournPrinc <> '') and
      (not (DS.State in [dsInsert])) then
    begin
      // Le fournisseur ne peut être changé, si cet article figure dans les lignes des documents
      if ExisteSQL('SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE LIKE "' + Copy(GetField('GA_ARTICLE'), 1, 18) + '%"') then
      begin
        if AffBox = True then PGIBox(TexteMessage[40], Ecran.Caption);
        SetField('GA_FOURNPRINC', OldFournPrinc);
        SetFocusControl('GA_FOURNPRINC');
        Result := False;
        exit;
      end;

      // Test si un catalogue de cet article existe déjà pour l'ancien fournisseur
      QQ := OpenSQL('Select GCA_REFERENCE from CATALOGU where GCA_ARTICLE="' +
        GetField('GA_ARTICLE') + '" and GCA_TIERS="' + OldFournPrinc + '"', TRUE);
      if not QQ.EOF then
      begin
        if HShowMessage('0;' + TexteMessage[45] + ';' + TexteMessage[50] + ';Q;YN;N;N;', '', '') = mrYes then
          ExecuteSQL('DELETE FROM CATALOGU WHERE GCA_REFERENCE="' +
            QQ.Findfield('GCA_REFERENCE').AsString + '" and GCA_TIERS="' + OldFournPrinc + '"');
      end;
      Ferme(QQ);
      OldFournPrinc := GetField('GA_FOURNPRINC');
    end;
  end;

  IF GetField('GA_FOURNPRINC')='' then
  begin
//    SetField('GA_QUALIFUNITEACH','');
//    SetField('GA_COEFCONVQTEACH',0.0);
    SetField('GA_PAUA',0.0);
    SetField('GCA_PAUA',0.0);
    //SetInfosPrixTarif;
  end;

    SetInfosPrixTarif;

  if TypeArticle = 'ARP' then AffichageInfosPrixPose;

  DefiniAffichage;

end;

//
// Procedure permettant de poser ou retirer un évênement sur le OnElipsisClick du
// champ GA_FOURNPRINC, en fonction de son contenu et du paramètre societe SO_MONOFOURNISS.
//

procedure TOM_Article.PoseEvenementFournisseur;
var Existe_Fournisseur: boolean;
  {$IFDEF EAGLCLIENT}
  Chp_Fournisseur: THEdit;
  {$ELSE}
  Chp_Fournisseur: THDBEdit;
  {$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  Chp_Fournisseur := THEdit(GetControl('GA_FOURNPRINC'));
  {$ELSE}
  Chp_Fournisseur := THDBEdit(GetControl('GA_FOURNPRINC'));
  {$ENDIF}
  if Chp_Fournisseur <> nil then
  begin
    Chp_Fournisseur.OnElipsisClick := OnElipsisClick_Fournisseur;
    if BMonoFournisseur = True then
    begin
      // ----- Gestion Mono-Fournisseur -----
      // L'utilisateur sélectionne d'abord le fournisseur, avant d'accéder au catalogue.
      // Si le fournisseur est inexistant, l'évènement sur le OnElipsisClick est retiré,
      // et dans ce cas la tablette Fournisseur s'affiche.
      Existe_Fournisseur := False;
      if (GetField('GA_FOURNPRINC') <> '') then
        if ExisteSQL('Select T_AUXILIAIRE from TIERS where T_TIERS="' + GetField('GA_FOURNPRINC') + '"')
          then Existe_Fournisseur := True;
      if Existe_Fournisseur = False then Chp_Fournisseur.OnElipsisClick := nil;
    end;
  end;
end;

// Procédure exécutée suite à l'évênement OnElipsisClick du champ GA_FOURNPRINC

procedure TOM_Article.OnElipsisClick_Fournisseur(Sender: TObject);
var CodeSStarif : string;
    Rect : Trect;
    X,Y : integer;
    Pos : Tpoint;
begin

  if ControleModifFournisseur(True) = False then Exit;
  //
  Rect := ThEdit(GetControl('GA_FOURNPRINC')).ClientRect;
  X := Rect.Left + ThEdit(GetControl('GA_FOURNPRINC')).Width + 1 ;
  Y := Rect.Top;
  POs := ThEdit(GetControl('GA_FOURNPRINC')).ClientToScreen(POint(X,Y));
  TPopupMenu(GetControl('POPCHOIXFOUR')).Popup(Pos.X,Pos.Y);
  //
  (* -----------
  if CodeSousTarif <> nil then CodeSStarif := CodeSousTarif.value else CodeSSTarif := '';
  LookupFournisseur;
  {$IFDEF BTP}
  if (copy(ecran.name, 1, 9) = 'BTARTICLE') then
  begin
    GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'),
      GetField('GA_TARIFARTICLE'), CodeSStarif, TOBcatalog, TOBTarif);
    SetInfosPrixTarif;
    DefiniAffichage;
    if Typearticle = 'ARP' then AffichageInfosPrixPose;
  end;
  {$ENDIF}
  ---------------- *)
end;

// LookUp pour la gestion d'un catalogue article Mono-fournisseur ou Multi-fournisseurs

procedure TOM_Article.LookUpFournisseur;
var TOBFourn, TOBFournFille,TOBCataFouPrinc: TOB;
  QFourn: TQuery;
  SQL: string;
  MaCle, CodFourn, CodArtic, CodRefer: string;
  TypeRech: string;
  ExisteCatalogue: boolean;
  NbCatalogue: integer;
  CodeSSTarif,SousFamilletarif  : string;
    MTPAF : double;
    TarifComplet : boolean;
    PrixPourQTe : double;
begin
	if CodeSousTarif <> nil then CodeSStarif := CodeSousTarif.value else CodeSStarif := '';
  TOBFourn := nil;
  ArticleFourn := True;
  CodFourn := GetField('GA_FOURNPRINC');
  CodArtic := GetField('GA_ARTICLE');
  TRY
  // Recherche s'il existe au moins un catalogue fournisseur pour l'article
  TypeRech := '1';
  ExisteCatalogue := ExisteSQL('Select GCA_REFERENCE from CATALOGU where GCA_ARTICLE="' + CodArtic + '"');
  if (BMonoFournisseur = True) and
    ((not ExisteCatalogue) or (Trim(Copy(CodArtic, 1, 18)) = '')) then
  begin
    // Recherche s'il existe dans le catalogue, des références du fournisseur non
    // rattaché à un article
    TypeRech := '2';
    CodArtic := '';
    CodArtic := CodeArticleUnique2(CodArtic, '');
    ExisteCatalogue := ExisteSQL('Select GCA_REFERENCE from CATALOGU where (GCA_ARTICLE="' + CodArtic + '") and (GCA_TIERS="' + CodFourn + '")');
  end;

  if ExisteCatalogue = True then
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
    if (BMonoFournisseur = True) and (NbCatalogue = 1) and (TypeRech = '1') and
      (GetField('GA_FOURNPRINC') = CodFourn) then
    begin
      // Acces direct en modification de la fiche Catalogue du fournisseur de l'article
      MaCle := AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', CodRefer + ';' + CodFourn, 'ACTION=MODIFICATION');
    end else
    begin
      // Affichage de la fiche Grid proposant les différents choix du catalogue fournisseur
      TheTob := TOBFourn;
  //      MaCle := AGLLanceFiche('GC', 'GCGRIDFOURNISSEUR', '', '', 'FOURN=' + CodFourn + ';LIBART=' + GetField('GA_LIBELLE') + ';TYPAFF=' + TypeRech);
        MaCle:=AGLLanceFiche('BTP','BTGRIDFOURNISSEUR','','','ART='+CodArtic+';FOURN='+CodFourn+';LIBART='+GetField('GA_LIBELLE')+';TYPAFF='+TypeRech) ;
    end;
  end else
  begin
    // Acces en création de la fiche Catalogue du fournisseur principal de l'article
    MaCle := AGLLanceFiche('GC', 'GCCATALOGU_SAISI3', '', '', 'ACTION=CREATION;LIB=' + GetField('GA_LIBELLE'));
  end;

  if (BMonoFournisseur = False) then
  begin
    if (MaCle = 'Abandon') or (MaCle = '') and (GetControlText('GA_FOURNPRINC') = '') then FournPrinc := '';
    if (MaCle <> '') and (MaCle <> 'Abandon') then FournPrinc := ReadTokenst(MaCle);
    if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
    SetControlText('GA_FOURNPRINC', FournPrinc);
    {$IFDEF BTP}
    if copy(ecran.name, 1, 9) = 'BTARTICLE' then
    begin
      GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'),
        GetField('GA_TARIFARTICLE'), CodeSStarif, TOBcatalog, TOBTarif);
      SetInfosPrixTarif;
        DefiniAffichage;
      if TypeArticle = 'ARP' then AffichageInfosPrixPose;
    end;
    {$ENDIF}
  end;
  FINALLY
  if TOBFourn <> nil then TOBFourn.free;
  END;
end;

// Procédure exécutée lors de la sortie du champ GA_FOURNPRINC

procedure TOM_Article.OnExit_Fournisseur(Sender: TObject);
var TypeCAB: string;
  LgNumCAB: integer;
begin

  if (DS <> nil) and (DS.State in [dsInsert]) then
  begin
    if (GetField('GA_CODEBARRE') = '') and (GetField('GA_FOURNPRINC') <> '') and
      (GetField('GA_FOURNPRINC') <> OldFournPrinc) and (TypeArticle = 'MAR') then
    begin
      // Dans le cas d'une saisie manuelle du code à barres, on recherche s'il existe
      // une qualification du code à barres paramétrée pour ce fournisseur
      if RechParamCAB_Manu(GetField('GA_FOURNPRINC'), TypeCAB, LgNumCAB) = True then
        SetField('GA_QUALIFCODEBARRE', TypeCAB);
    end;
  end;

  // Contrôle sur la modification du fournisseur de la fiche Article
  ControleModifFournisseur(True);
  
end;

//
// Modif BTP 06/02/2001
// Modif MODE 07/01/2002 - Calcul PA->PR ou PR->PA
// Procedure de calcul des bases tarif PR

procedure TOM_Article.CalculBaseTarifPR(Click: boolean); // click pour bouton actualisation des prix
var CodeBasePR: string;
  BaseTarifPR, CoefCalcPR, TarifFournisseur, QuotiteAchat, QuotiteVente, PrixRef: double;
  QQ, QQUALIFACH, QQUALIFVTE: TQuery;
  {$IFDEF BTP}
  PAvisi, isnomenc: Boolean;
  pa, pr, pv: double;
  {$ENDIF}
begin
  {$IFDEF BTP}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM') or (GetField('GA_TYPEARTICLE') = 'OUV');
  PAVisi := (not Isnomenc) and (not (copy(ecran.name, 1, 12) = 'BTPRESTATION'));
  {$ENDIF}
  // Non car sinon pas de remise à jour des champs en contrevaleur suite aux maj des champs principaux.
  //  CalculEnCours:=True ;
  PrixRef := 0;
  ModifChampHT := False;
  FaireCalcul := False;
  CoefCalcPR := Valeur(GetControlText('GA_COEFFG'));
  if CtxMode in V_PGI.PGIContexte then
  begin
    if CalcPAPR = 'AR' then CodeBasePR := 'PAA'
    else if CalcPAPR = 'RA' then CodeBasePR := 'PRA'
  end
  else CodeBasePR := GetControlText('GA_CALCPRIXPR');

  if (CoefCalcPR <> 0) and (CodeBasePR <> '') then
  begin
    CalculPrixRef(CodeBasePR, PrixRef);
    //
    //Spécificités BTP
    if (CodeBasePr <> 'FOU') and (GetControlText('GA_CALCPRIXHT') <> 'FOU') and (GetControlText('GA_CALCPRIXTTC') <> 'FOU') then
    begin
      if not IsLine then SetControlvisible('GA_DPA', Pavisi);
      if not IsLine then SetControlVisible('PAFOUPRINC', false);
      if not IsLine then SetControlvisible('TGA_DPA', Pavisi);
      if not IsLine then SetControlVisible('TPAFOUPRINC', false);
    end;
    //
    BaseTarifPR := PrixRef * CoefCalcPR;
    //
    if CtxMode in V_PGI.PGIContexte then
    begin
      if CalcPAPR = 'AR' then
        SetField('GA_PRHT', FloatToStr(BaseTarifPR))
      else if CalcPAPR = 'RA' then
        SetField('GA_PAHT', FloatToStr(BaseTarifPR))
    end
    else if GetField('GA_DPRAUTO')='X' then
      SetField('GA_DPR', FloatToStr(BaseTarifPR));
      if (GetField('GA_DPRAUTO') = 'X') then
        SetControlProperty('GA_DPR', 'Color', clbtnFace)
      else
        SetControlProperty('GA_DPR', 'Color', clWindow);
    //
    ClickActuPrix := Click;
    //
  end;
  {$IFDEF BTP}
  if (copy(Ecran.name, 1, 13) <> 'BTARTPOURCENT') and (copy(ecran.name,1,9)<>'BTARTPARC') then
  begin
    DetermineValorisation(pa, pr, pv);
    AfficheMarge(THValComboBox(GetControl('REFMARGE')).ItemIndex, pa, pv, pr);
  end;
  {$ENDIF}
  // Non car sinon pas de remise à jour des champs en contrevaleur suite aux maj des champs principaux.
  //CalculEnCours:=False ;  // Pas de recalcul automatique en cours de prix ou coefficients

end;
// -----------------------------------------------

// Procedure de calcul des bases tarif HT

Procedure TOM_Article.CalculPrixRef(CodeBasePR : String; Var PrixRef : double);
Var StSQl         : String;
    QQ            : TQuery;
    UA            : string;
    UV            : string;
    QuotiteAchat  : Double;
    QuotiteVente  : double;
    TarifFournisseur  : Double;
begin

  if CodeBasePR = 'DPA' then
  begin
    PrixRef := Valeur(GetControlText('GA_DPA'));
  end
  else if CodeBasePR = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
  else if CodeBasePR = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
  else if CodeBasePR = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
  else if CodeBasePR = 'FOU' then
  begin
    StSQL := 'Select GCA_QUALIFUNITEACH, GCA_DPA From Catalogu where GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '" And GCA_TIERS="' +      GetControlText('GA_FOURNPRINC') + '"';
    QQ := OpenSql(StSQl, True);
    if not QQ.Eof then
    begin
      UA := QQ.FindField('GCA_QUALIFUNITEACH').AsString;
      UV := GetControlText('GA_QUALIFUNITEVTE');
      TarifFournisseur := Valeur(QQ.FindField('GCA_DPA').AsString);
      //
      Ferme(QQ);
      //
      StSQL := 'Select GME_QUOTITE From MEA where GME_MESURE="' + UA + '"';
      QQ    := OpenSql(StSql, True);
      if not QQ.Eof then
        QuotiteAchat := Valeur(QQ.FindField('GME_QUOTITE').AsString)
      else
        QuotiteAchat := 1;
      //
      Ferme(QQ);

      StSQl := 'Select GME_QUOTITE From MEA where GME_MESURE="' + UV + '"';
      QQ    := OpenSql(StSQL, True);
      if not QQ.Eof then
        QuotiteVente := Valeur(QQ.FindField('GME_QUOTITE').AsString)
      else
        QuotiteVente := 1;
      Ferme(QQ);
      //
      PrixRef := ((TarifFournisseur * QuotiteVente / QuotiteAchat) * Valeur(GetControlText('GA_PRIXPOURQTE')));
      //
    end;
    Ferme(QQ);

    //spécificté BTP
    THNumEdit(GetControl('PAFOUPRINC')).Value := PrixRef;
    if not IsLine then SetControlvisible('GA_DPA', false);
    if not IsLine then SetControlVisible('PAFOUPRINC', true);
    if not IsLine then SetControlvisible('TGA_DPA', false);
    if not IsLine then SetControlVisible('TPAFOUPRINC', True);
  end;

end;

procedure TOM_Article.CalculBaseTarifHT(Click: boolean); // click pour bouton actualisation des prix
var CodeBaseHT, CodeArrondi: string;
  BaseTarifHT, CoefCalcHT, TarifFournisseur, QuotiteAchat, QuotiteVente, PrixRef: double;
  QQ, QQUALIFACH, QQUALIFVTE: TQuery;
  {$IFDEF BTP}
  pa, pr, pv: double;
  PAvisi, isnomenc: Boolean;
  {$ENDIF}
begin
  //
  if not boolean(GetField('GA_CALCAUTOHT') = 'X') then Exit;
  //
  CalculEnCours := True;
  PrixRef := 0;
  ModifChampHT := False;
  FaireCalcul := False;
  CodeBaseHT := GetControlText('GA_CALCPRIXHT');
  CoefCalcHT := Valeur(GetControlText('GA_COEFCALCHT'));
    //mcd 25/09/03 ajout tets présence champ
  If THValComboBox(GetControl('GA_ARRONDIPRIX')) = nil then CodeArrondi:=''
   else CodeArrondi := GetControlText('GA_ARRONDIPRIX');
  {$IFDEF BTP}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM') or (GetField('GA_TYPEARTICLE') = 'OUV');
  PAVisi := (not Isnomenc) and (not (copy(ecran.name, 1, 12) = 'BTPRESTATION'));
  {$ENDIF}

  if (CoefCalcHT <> 0) and (CodeBaseHT <> '') then
  begin
    if CodeBaseHT = 'DPA' then
    begin
      PrixRef := Valeur(GetControlText('GA_DPA'));
    end
    else if CodeBaseHT = 'DPR' then PrixRef := Valeur(GetControlText('GA_DPR'))
    else if CodeBaseHT = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
//    else if CodeBaseHT = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
//    else if CodeBaseHT = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
//    else if CodeBaseHT = 'PMR' then PrixRef := Valeur(GetControlText('GA_PMRP'))
    else if CodeBaseHT = 'FOU' then
    begin
      QQ := OpenSql('Select GCA_QUALIFUNITEACH, GCA_DPA From Catalogu where GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '" And GCA_TIERS="' +
        GetControlText('GA_FOURNPRINC') + '"', True);
      if not QQ.Eof then
      begin
        QQUALIFACH := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + QQ.FindField('GCA_QUALIFUNITEACH').AsString + '"', True);
        QQUALIFVTE := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + GetControlText('GA_QUALIFUNITEVTE') + '"', True);
        if not QQUALIFACH.Eof then QuotiteAchat := Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString)
        else QuotiteAchat := 1;
        if not QQUALIFVTE.Eof then QuotiteVente := Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString)
        else QuotiteVente := 1;
        TarifFournisseur := Valeur(QQ.FindField('GCA_DPA').AsString);
        PrixRef := ((TarifFournisseur * QuotiteVente / QuotiteAchat) * Valeur(GetControlText('GA_PRIXPOURQTE')));
        Ferme(QQUALIFACH);
        Ferme(QQUALIFVTE);
      end;
      Ferme(QQ);
      {$IFDEF BTP}
      THNumEdit(GetControl('PAFOUPRINC')).Value := PrixRef;
      if not IsLine then SetControlvisible('GA_DPA', false);
      if not IsLine then SetControlVisible('PAFOUPRINC', true);
      if not IsLine then SetControlvisible('TGA_DPA', false);
      if not IsLine then SetControlVisible('TPAFOUPRINC', true);
      {$ENDIF}
    end;
    BaseTarifHT := PrixRef * CoefCalcHT;
    BaseTarifHT := ArrondirPrix(CodeArrondi, BaseTarifHT);
  	if GetField('GA_CALCAUTOHT')='X' then SetField('GA_PVHT', FloatToStr(BaseTarifHT));

    // Ajout LS
    SetField('GA_PVHT', arrondi(GetField('GA_PVHT'), NbDecPrix));

    //SetField('GA_PVHT', arrondi(GetField('GA_PVHT'), V_PGI.OkDecV));
    // --
    ClickActuPrix := Click;
  end;
  {$IFDEF BTP}
  if (copy(Ecran.name, 1, 13) <> 'BTARTPOURCENT') and (copy(ecran.name,1,9)<>'BTARTPARC') then
  begin
    DetermineValorisation(pa, pr, pv);
    AfficheMarge(THValComboBox(GetControl('REFMARGE')).ItemIndex, pa, pv, pr);
  end;
  {$ENDIF}
  CalculEnCours := False; // Pas de recalcul automatique en cours de prix ou coefficients
end;

// Procedure de calcul des bases tarif TTC

procedure TOM_Article.CalculBaseTarifTTC(Click: boolean); // bouton actualisation true ou false
var CodeBaseTTC, CodeArrondi: string;
  BaseTarifTTC, CoefCalcTTC, PrixRef: double;
  QQ, QQUALIFACH, QQUALIFVTE: TQuery;
  TarifFournisseur, QuotiteAchat, QuotiteVente: double;
  {$IFDEF BTP}
  pa, pr, pv: double;
  PAvisi, isnomenc: Boolean;
  {$ENDIF}
begin
  CalculEnCours := True;
  PrixRef := 0;
  ModifChampTTC := False;
  FaireCalcul := False;
  CodeBaseTTC := GetControlText('GA_CALCPRIXTTC');
  CoefCalcTTC := Valeur(GetControlText('GA_COEFCALCTTC'));
  CodeArrondi := GetControlText('GA_ARRONDIPRIXTTC');
  {$IFDEF BTP}
  IsNomenc := (GetField('GA_TYPEARTICLE') = 'NOM') or (GetField('GA_TYPEARTICLE') = 'OUV');
  PAVisi := (not Isnomenc) and (not (copy(ecran.name, 1, 12) = 'BTPRESTATION'));
  {$ENDIF}

  if (CoefCalcTTC <> 0) and (CodeBaseTTC <> '') and (CodeBaseTTC <> 'AUC') then
  begin
    if CodeBaseTTC = 'DPA' then PrixRef := Valeur(GetControlText('GA_DPA'))
    else if CodeBaseTTC = 'DPR' then PrixRef := Valeur(GetControlText('GA_DPR'))
    else if CodeBaseTTC = 'HT' then PrixRef := Valeur(GetControlText('GA_PVHT'))
    else if CodeBaseTTC = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
//    else if CodeBaseTTC = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
//    else if CodeBaseTTC = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
//    else if CodeBaseTTC = 'PMR' then PrixRef := Valeur(GetControlText('GA_PMRP'))
      // Modif BTP
    else if CodeBaseTTC = 'FOU' then
    begin
      QQ := OpenSql('Select GCA_QUALIFUNITEACH, GCA_DPA From Catalogu where GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '" And GCA_TIERS="' +
        GetControlText('GA_FOURNPRINC') + '"', True);
      if not QQ.Eof then
      begin
        QQUALIFACH := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + QQ.FindField('GCA_QUALIFUNITEACH').AsString + '"', True);
        QQUALIFVTE := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + GetControlText('GA_QUALIFUNITEVTE') + '"', True);
        if not QQUALIFACH.Eof then QuotiteAchat := Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString)
        else QuotiteAchat := 1;
        if not QQUALIFVTE.Eof then QuotiteVente := Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString)
        else QuotiteVente := 1;
        TarifFournisseur := Valeur(QQ.FindField('GCA_DPA').AsString);
        PrixRef := ((TarifFournisseur * QuotiteVente / QuotiteAchat) * Valeur(GetControlText('GA_PRIXPOURQTE')));
        Ferme(QQUALIFACH);
        Ferme(QQUALIFVTE);
      end;
      Ferme(QQ);
      {$IFDEF BTP}
      THNumEdit(GetControl('PAFOUPRINC')).Value := PrixRef;
      if not IsLine then SetControlvisible('GA_DPA', false);
      if not IsLine then SetControlVisible('PAFOUPRINC', true);
      if not IsLine then SetControlvisible('TGA_DPA', false);
      if not IsLine then SetControlVisible('TPAFOUPRINC', true);
      {$ENDIF}
    end;
    // --
    BaseTarifTTC := PrixRef * CoefCalcTTC;
    BaseTarifTTC := ArrondirPrix(CodeArrondi, BaseTarifTTC);
    SetField('GA_PVTTC', FloatToStr(BaseTarifTTC));
    // Ajout LS
    if GetField('GA_CALCAUTOTTC')='X' then  SetField('GA_PVTTC', arrondi(GetField('GA_PVTTC'), NBDecPrix));
    if (GetField('GA_CALCAUTOTTC') = 'X') then SetControlProperty('GA_PVTTC', 'Color', clbtnFace)
                                      	  else SetControlProperty('GA_PVTTC', 'Color', clWindow);
//    SetField('GA_PVTTC', arrondi(GetField('GA_PVTTC'), V_PGI.OkDecV));
    // --
    ClickActuPrix := Click;
    // Modif BTP
    {$IFDEF BTP}
    if BasetarifTTC = 0 then SetControlText('TGA_PVTTC', TraduireMemoire('Base Tarif T.T.C (non déterminé)'))
    else SetControlText('TGA_PVTTC', TraduireMemoire('Base Tarif T.T.C'));
  	if (copy(Ecran.name, 1, 13) <> 'BTARTPOURCENT') and (copy(ecran.name,1,9)<>'BTARTPARC') then
    begin
      DetermineValorisation(pa, pr, pv);
      AfficheMarge(THValComboBox(GetControl('REFMARGE')).ItemIndex, pa, pv, pr);
    end;
    {$ENDIF}
  end;
  CalculEnCours := False; // Pas de recalcul automatique en cours de prix ou coefficients
end;

//
// Modif BTP 06/02/2001
//

procedure TOM_Article.CalculCoefTarifPR(Click: boolean); // click pour bouton actualisation des prix
var CodeBasePR: string;
  BaseTarifPR, CoefCalcPR, TarifFournisseur, QuotiteAchat, QuotiteVente, PrixRef: double;
  QQ, QQUALIFACH, QQUALIFVTE: TQuery;
begin
  CalculEnCours := True;
  ModifChampPR := False;
  FaireCalcul := False;
  BaseTarifPR := 0;
  if CtxMode in V_PGI.PGIContexte then
  begin
    if CalcPAPR = 'AR' then
    begin
      CodeBasePR := 'PAA';
      BaseTarifPR := Valeur(GetControlText('GA_PRHT'));
    end
    else if CalcPAPR = 'RA' then
    begin
      CodeBasePR := 'PRA';
      BaseTarifPR := Valeur(GetControlText('GA_PAHT'));
    end
  end else
  begin
    CodeBasePR := GetControlText('GA_CALCPRIXPR');
    BaseTarifPR := Valeur(GetControlText('GA_DPR'));
  end;

  PrixRef := 0;

  if (BaseTarifPR <> 0) and (CodeBasePR <> '') then
  begin
    if CodeBasePR = 'DPA' then PrixRef := Valeur(GetControlText('GA_DPA'))
    else if CodeBasePR = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
    else if CodeBasePR = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
    else if CodeBasePR = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
    else if CodeBasePR = 'FOU' then
    begin
      QQ := OpenSql('Select GCA_QUALIFUNITEACH, GCA_DPA From Catalogu where GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '" And GCA_TIERS="' +
        GetControlText('GA_FOURNPRINC') + '"', True);
      if not QQ.Eof then
      begin
        QQUALIFACH := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + QQ.FindField('GCA_QUALIFUNITEACH').AsString + '"', True);
        QQUALIFVTE := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + GetControlText('GA_QUALIFUNITEVTE') + '"', True);
        if not QQUALIFACH.Eof then QuotiteAchat := Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString)
        else QuotiteAchat := 1;
        if not QQUALIFVTE.Eof then QuotiteVente := Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString)
        else QuotiteVente := 1;
        TarifFournisseur := Valeur(QQ.FindField('GCA_DPA').AsString);
        PrixRef := ((TarifFournisseur * QuotiteVente / QuotiteAchat) * Valeur(GetControlText('GA_PRIXPOURQTE')));
        Ferme(QQUALIFACH);
        Ferme(QQUALIFVTE);
      end;
      Ferme(QQ);
    end;
    if PrixRef <> 0 then
    begin
      CoefCalcPR := BaseTarifPR / PrixRef;
      SetField('GA_COEFFG', FloatToStr(CoefCalcPR));
      if not (CtxMode in V_PGI.PGIContexte) then SetField('GA_DPR', FloatToStr(BaseTarifPR));
      ClickActuPrix := Click;
    end
    else SetLastError(37, 'GA_CALCPRIXPR');
  end;
  CalculEnCours := False;
end;
// -------------------------------------

// Procedure de calcul du coefficient du tarif HT

procedure TOM_Article.CalculCoefTarifHT(Click: boolean); // click pour bouton actualisation des prix
var CodeBaseHT, CodeArrondi: string;
  BaseTarifHT, CoefCalcHT, TarifFournisseur, QuotiteAchat, QuotiteVente, PrixRef: double;
  QQ, QQUALIFACH, QQUALIFVTE: TQuery;
begin
  CalculEnCours := True;
  ModifChampHT := False;
  FaireCalcul := False;
  CodeBaseHT := GetControlText('GA_CALCPRIXHT');
  BaseTarifHT := Valeur(GetControlText('GA_PVHT'));
    If THValComboBox(GetControl('GA_ARRONDIPRIX')) = nil then CodeArrondi :='' //mcd 25/09/03 champ inexistant en GI/GA
     else CodeArrondi := GetControlText('GA_ARRONDIPRIX');
  BaseTarifHT := ArrondirPrix(CodeArrondi, BaseTarifHT);
  PrixRef := 0;

  if (BaseTarifHT <> 0) and (CodeBaseHT <> '') then
  begin
    if CodeBaseHT = 'DPA' then PrixRef := Valeur(GetControlText('GA_DPA'))
    else if CodeBaseHT = 'DPR' then PrixRef := Valeur(GetControlText('GA_DPR'))
    else if CodeBaseHT = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
    else if CodeBaseHT = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
    else if CodeBaseHT = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
    else if CodeBaseHT = 'PMR' then PrixRef := Valeur(GetControlText('GA_PMRP'))
    else if CodeBaseHT = 'FOU' then
    begin
      QQ := OpenSql('Select GCA_QUALIFUNITEACH, GCA_DPA From Catalogu where GCA_ARTICLE="' + GetControlText('GA_ARTICLE') + '" And GCA_TIERS="' +
        GetControlText('GA_FOURNPRINC') + '"', True);
      if not QQ.Eof then
      begin
        QQUALIFACH := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + QQ.FindField('GCA_QUALIFUNITEACH').AsString + '"', True);
        QQUALIFVTE := OpenSql('Select GME_QUOTITE From MEA where GME_MESURE="' + GetControlText('GA_QUALIFUNITEVTE') + '"', True);
        if not QQUALIFACH.Eof then QuotiteAchat := Valeur(QQUALIFACH.FindField('GME_QUOTITE').AsString)
        else QuotiteAchat := 1;
        if not QQUALIFVTE.Eof then QuotiteVente := Valeur(QQUALIFVTE.FindField('GME_QUOTITE').AsString)
        else QuotiteVente := 1;
        TarifFournisseur := Valeur(QQ.FindField('GCA_DPA').AsString);
        PrixRef := ((TarifFournisseur * QuotiteVente / QuotiteAchat) * Valeur(GetControlText('GA_PRIXPOURQTE')));
        Ferme(QQUALIFACH);
        Ferme(QQUALIFVTE);
      end;
      Ferme(QQ);
    end;
    if PrixRef <> 0 then
    begin
      CoefCalcHT := BaseTarifHT / PrixRef;
      SetField('GA_COEFCALCHT', FloatToStr(CoefCalcHT));
      SetField('GA_PVHT', FloatToStr(BaseTarifHT));
      ClickActuPrix := Click;
    end
    else if bCalculPrixDuplicHT then SetLastError(37, 'GA_CALCPRIXHT')
    else bCalculPrixDuplicHT := True;
  end;
  CalculEnCours := False;
end;

// Procedure de calcul du coefficient du tarif TTC

procedure TOM_Article.CalculCoefTarifTTC(Click: boolean); // click pour bouton actualisation des prix
var CodeBaseTTC, CodeArrondi: string;
  BaseTarifTTC, CoefCalcTTC, PrixRef: double;
begin
  CalculEnCours := True;
  PrixRef := 0;
  ModifChampTTC := False;
  FaireCalcul := False;
  CodeBaseTTC := GetControlText('GA_CALCPRIXTTC');
  BaseTarifTTC := Valeur(GetControlText('GA_PVTTC'));
  CodeArrondi := GetControlText('GA_ARRONDIPRIXTTC');
  BaseTarifTTC := ArrondirPrix(CodeArrondi, BaseTarifTTC);

  if (BaseTarifTTC <> 0) and (CodeBaseTTC <> '') and (CodeBaseTTC <> 'AUC') then
  begin
    if CodeBaseTTC = 'DPA' then PrixRef := Valeur(GetControlText('GA_DPA'))
    else if CodeBaseTTC = 'DPR' then PrixRef := Valeur(GetControlText('GA_DPR'))
    else if CodeBaseTTC = 'HT'  then PrixRef := Valeur(GetControlText('GA_PVHT'))
    else if CodeBaseTTC = 'PAA' then PrixRef := Valeur(GetControlText('GA_PAHT'))
    else if CodeBaseTTC = 'PRA' then PrixRef := Valeur(GetControlText('GA_PRHT'))
    else if CodeBaseTTC = 'PMA' then PrixRef := Valeur(GetControlText('GA_PMAP'))
    else if CodeBaseTTC = 'PMR' then PrixRef := Valeur(GetControlText('GA_PMRP'));
    if PrixRef <> 0 then
    begin
      CoefCalcTTC := BaseTarifTTC / PrixRef;
      SetField('GA_PVTTC', FloatToStr(BaseTarifTTC));
      SetField('GA_COEFCALCTTC', FloatToStr(CoefCalcTTC));
      ClickActuPrix := Click;
    end
    else if bCalculPrixDuplicTTC then SetLastError(37, 'GA_CALCPRIXTTC')
    else bCalculPrixDuplicTTC := True;
  end;
  CalculEnCours := False; // Pas de recalcul automatique en cours de prix ou coefficients
end;

procedure TOM_Article.SetTypeArticle(typeArt: string);
begin
  TypeArticle := typeArt;
end;

procedure TOM_Article.ControlNomenc;
begin
  if (GetControlText('GA_TYPENOMENC') = '') then
  begin
    if (Typenomenc = 'ASS') then SetControlText('GA_TYPENOMENC', 'ASS')
      //
      // Modif BTP
      //
    else if (TypeNomenc = 'OUV') then SetControlText('GA_TYPENOMENC', 'OUV')
    else if (TypeNomenc = 'OU1') then SetControlText('GA_TYPENOMENC', 'OU1')
      //
    else SetControlText('GA_TYPENOMENC', 'MAC');
  end;
  // Modif BTP
  {$IFDEF BTP}
  SetControlText('GA_CALCPRIXHT', 'AUC');
  SetControlText('GA_CALCPRIXTTC', 'AUC');
  SetControlText('GA_CALCPRIXPR', 'AUC');
  {$ELSE}
  if GetField('GA_CALCPRIXHT') = '' then SetControlText('GA_CALCPRIXHT', 'AUC');
  if GetField('GA_CALCPRIXTTC') = '' then SetControlText('GA_CALCPRIXTTC', 'AUC');
  {$ENDIF}
  if GetField('GA_TENUESTOCK') = 'X' then SetField('GA_TENUESTOCK', '-');
  if (copy(GetControlText('GA_CALCPRIXHT'), 1, 1) = 'S') then
  begin
    LastError := 28;
    LastErrorMsg := TexteMessage[LastError];
    {$IFDEF EAGLCLIENT}
    THValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    {$ELSE}
    THDBValComboBox(GetControl('GA_CALCPRIXHT')).Plus := 'AND ((CO_CODE NOT LIKE "S%") AND (CO_CODE<>"FOU"))';
    {$ENDIF}
    SetActiveTabSheet('P_TARIF');
    SetFocusControl('GA_CALCPRIXHT');
    Exit;
  end;
  if (copy(GetControlText('GA_CALCPRIXTTC'), 1, 1) = 'S') then
  begin
    LastError := 28;
    LastErrorMsg := TexteMessage[LastError];
    {$IFDEF EAGLCLIENT}
    THValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")';
    {$ELSE}
    THDBValComboBox(GetControl('GA_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")';
    {$ENDIF}
    SetActiveTabSheet('P_TARIF');
    SetFocusControl('GA_CALCPRIXTTC');
    Exit;
  end;
end;

// gestion du code unique des articles lié

procedure TOM_Article.CodeUniqueArtLie(TypeLien: string);
var CodeUniqueRem, CodeUniqueSub, CodeDimRem, CodeDimSub: string;
begin
  if TypeLien = 'REM' then
  begin
    if IsCodeArticleUnique(GetControlText('REMPLACEMENT')) then SetControlText('REMPLACEMENT', Trim(Copy(GetControlText('REMPLACEMENT'), 1, 18)));
    if GetControlText('REMPLACEMENT') = Trim(Copy(GetField('GA_REMPLACEMENT'), 1, 18)) then exit;
    if GetControlText('REMPLACEMENT') = '' then
    begin
      SetField('GA_REMPLACEMENT', '');
      THLabel(GetControl('TREMPLACEMENT')).Caption := '';
      exit;
    end;
    CodeUniqueRem := CodeArticleUnique2(GetControlText('REMPLACEMENT'), '');
    //article dimensionné
    if not (Ctxmode in V_PGI.PGIContexte) then
    begin
      if ExisteSQL('Select GA_ARTICLE from ARTICLE where GA_CODEARTICLE="' + GetControlText('REMPLACEMENT') + '" AND GA_STATUTART="GEN"')
        then CodeDimRem := SelectUneDimension(CodeUniqueRem);
    end;
    if CodeDimRem <> '' then SetField('GA_REMPLACEMENT', CodeDimRem) else
      SetField('GA_REMPLACEMENT', CodeUniqueRem); // si pas de dim selectionne: article GEN
    THLabel(GetControl('TREMPLACEMENT')).Caption := RechDom('GCARTICLEGENERIQUE', GetControlText('REMPLACEMENT'), False);
  end else
  begin
    if IsCodeArticleUnique(GetControlText('SUBSTITUTION')) then SetControlText('SUBSTITUTION', Trim(Copy(GetControlText('SUBSTITUTION'), 1, 18)));
    if GetControlText('SUBSTITUTION') = Trim(Copy(GetField('GA_SUBSTITUTION'), 1, 18)) then exit;
    if GetControlText('SUBSTITUTION') = '' then
    begin
      SetField('GA_SUBSTITUTION', '');
      THLabel(GetControl('TSUBSTITUTION')).Caption := '';
      exit;
    end;
    CodeUniqueSub := CodeArticleUnique2(GetControlText('SUBSTITUTION'), '');
    if not (Ctxmode in V_PGI.PGIContexte) then
    begin
      if ExisteSQL('Select GA_ARTICLE from ARTICLE where GA_CODEARTICLE="' + GetControlText('SUBSTITUTION') + '" AND GA_STATUTART="GEN"')
        then CodeDimSub := SelectUneDimension(CodeUniqueSub);
    end;
    if CodeDimSub <> '' then SetField('GA_SUBSTITUTION', CodeDimSub) else
      SetField('GA_SUBSTITUTION', CodeUniqueSub); // si pas de dim, article GEN
    THLabel(GetControl('TSUBSTITUTION')).Caption := RechDom('GCARTICLEGENERIQUE', GetControlText('SUBSTITUTION'), False);
  end;
end;

//************** Spécif prestations ********************************************

procedure TOM_Article.OnUpdatePrestation;
var PaHT: double;
begin
  if (GetField('GA_TYPEARTICLE') <> 'PRE') and (GetField('GA_TYPEARTICLE') <> 'CTR') then Exit;
  PaHT := GetField('GA_PAHT');
  if PaHT <> 0 then
  begin // maj de l'ensemble des prix d'achats sur la valeur de base
    SetField('GA_DPA', PaHT);
    // Modif BTP
    {$IFNDEF BTP}
    SetField('GA_PRHT', PaHT);
    SetField('GA_DPR', PaHT);
    {$ENDIF}
    SetField('GA_PMAP', PaHT);
    if GetField('GA_PMRP') = 0 then
    begin
      SetField('GA_PMRP', PaHT);
    end;
  end;
end;

procedure TOM_Article.LibresArtToLibresLigne;
var SQL, IdentArt, ValT: string;
  i: integer;
  Okok: boolean;
  DD: TDateTime;
begin
  if TFFiche(Ecran).fTypeAction <> taModif then Exit;
  if VH_GC.GCIfDefCEGID then
  begin
    IdentArt := GetField('GA_ARTICLE');
    if IdentArt = '' then Exit;
    SQL := 'UPDATE LIGNE SET';
    Okok := False;
    DD := EncodeDate(2000, 01, 01);
    for i := 1 to 9 do
    begin
      ValT := GetField('GA_LIBREART' + IntToStr(i));
      if ValT = '' then Continue;
      SQL := SQL + ' GL_LIBREART' + IntToStr(i) + '="' + ValT + '",';
      Okok := True;
    end;
    ValT := GetField('GA_LIBREARTA');
    if ValT <> '' then
    begin
      SQL := SQL + ' GL_LIBREARTA="' + ValT + '",';
      Okok := True;
    end;
    if not Okok then Exit;
    Delete(SQL, Length(SQL), 1);
    SQL := SQL + ' WHERE GL_ARTICLE="' + IdentArt + '" AND GL_DATEPIECE>="' + USDATETIME(DD) + '"';
    ExecuteSQL(SQL);
    end;
end;

procedure TOM_Article.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F10: if (Shift = [ssAlt]) then
      begin
        Key := 0;
        LibresArtToLibresLigne;
      end;
    VK_F12: if (Shift = [ssAlt]) then
      begin
        Key := 0;
        AGLLanceFiche('GC', 'GCJPEGMEMO', '', '', GetField('GA_ARTICLE'));
      end;
  end;
  if (ecran <> nil) then TFFiche(ecran).FormKeyDown(Sender, Key, Shift);
end;

procedure TOM_Article.ZonesLibArtIsModified;
begin
  if (CtxMode in V_PGI.PGIContexte) and ((GetPresentation = ART_ORLI) or (GetField('GA_TYPEARTICLE') = 'FI')) then
  begin
    TobZoneLibre.GetEcran(TFfiche(Ecran), nil);
    if TobZoneLibre.IsOneModifie and not (DS.State in [dsInsert, dsEdit]) then
    begin
      DS.edit; // pour passer DS.state en mode dsEdit
      {$IFDEF EAGLCLIENT}
      TFFiche(Ecran).QFiche.CurrentFille.Modifie := true;
      {$ELSE}
      SetField('GA_ARTICLE', GetControlText('GA_ARTICLE'));
      {$ENDIF}
    end;
  end;
end;

{Force la maj du libellé des champs stat. articles non issus du DataSet}

procedure TOM_Article.ChangeStatArt(CodeStat: string);
begin
  if Ecran.FindComponent('GA2_STATART' + CodeStat) <> Nil then
    SetControlText('TSTATART' + CodeStat, RechDom('GCSTATART' + CodeStat, GetControlText('GA2_STATART' + CodeStat), False));
end;

{$IFDEF GPAO}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Karine BORGHETTI
Créé le ...... : 30/01/2002
Modifié le ... : 30/01/2002
Description .. : Lancement de l'écran des conversions d'unités.
Mots clefs ... : CONVERSION;UNITE;
*****************************************************************}
procedure TOM_Article.MnUniteConversion_OnClick(Sender: TObject);
var
  OConv: TWConversionUnite;
begin
  OConv := TWConversionUnite.Create;
  try
    OConv.ConvUniteByDSAndShowFiche(wcuGA, GetField('GA_ARTICLE'), DS);
  finally
    OConv.Free;
  end;
end;
{$ENDIF}

{$IFDEF GPAOLIGHT}
function TOM_Article.BtArtNatCanBeActif: Boolean;
{ Vérifie si on peut accéder aux natures de travail }
var
  StatutArt, ModeEdit: Boolean;
begin
  Result := (GetField('GA_STATUTART') = 'GEN') or (GetField('GA_STATUTART') = 'UNI');
  if Result then
    Result := not (DS.State in [dsInsert, dsEdit]);
end;
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
procedure TOM_Article.SetBtArtNatState;
{ Etat du bouton d'accès aux natures de travail }
begin
  if Ecran <> nil then
  begin
    SetControlEnabled('BT_ARTNAT', BtArtNatCanBeActif);
    if Assigned(GetControl('MnNatureTravail')) then
      TMenuItem(GetControl('MnNatureTravail')).Enabled := BtArtNatCanBeActif;
  end;
end;
{$ENDIF GPAOLIGHT}

{$IFDEF GPAOLIGHT}
procedure TOM_Article.MnNatureTravail_OnClick(Sender: tObject);
begin
  if BtArtNatCanBeActif then
    wCallWAN(GetField('GA_ARTICLE'), TFFiche(Ecran).fTypeAction, '')
  else
    PgiInfo(TraduireMemoire('Vous n''avez pas validé votre saisie'));
end;
{$ENDIF GPAOLIGHT}

{$IFDEF GPAO}

procedure TOM_Article.MnPrixDeRevient_OnClick(Sender: tObject);
begin
  if DS.State = DSBrowse then
    AGLLanceFiche('W', 'WPDR_MUL', 'WPE_CODEARTICLE=' + GetField('GA_CODEARTICLE'), '', 'ACTION=CREATION    ;DROIT=CDAMV');
end;
{$ENDIF}

{$IFDEF GPAO}

procedure TOM_Article.WProfilArticleState;
var
  CodeProfil: string;
  ValuesProfil: MyArrayValue;
  ModeCreation: boolean;
begin
  if Ecran.Name <> 'WARTICLE_FIC' then
    Exit;
  ValuesProfil := nil;

  ModeCreation := DS.State = DsInsert;
  SetControlProperty('GA_ARTICLEPROFIL', 'VISIBLE', not OnlyProfil);
  SetControlProperty('TGA_ARTICLEPROFIL', 'VISIBLE', not OnlyProfil);
  SetControlProperty('GA_ESTPROFIL', 'VISIBLE', OnlyProfil);
  { GPAO DS le 25/23 }
  SetControlProperty('GA_ESTPROFIL', 'ENABLED', (OnlyProfil) and (not wExistArticleInGPF(GetField('GA_ARTICLE')))
    and (not wExistArticleProfilInGA(GetField('GA_ARTICLE'))));

  CodeProfil := GetField('GA_PROFILARTICLE');
  if CodeProfil = '' then // Pour récup. code profil si GetField ne ramène rien
    CodeProfil := ThDBValCombobox(GetControl('GA_PROFILARTICLE')).Value;

  if CodeProfil <> '' then
    ValuesProfil := wGetSqlFieldsValues(['GPF_TYPEPROFILART', 'GPF_ARTICLE'], 'PROFILART', 'GPF_PROFILARTICLE="' + CodeProfil + '"');

  if (CodeProfil = '') or ((CodeProfil <> '')
    and (ValuesProfil <> nil)
    and (((ValuesProfil[0] = 'SIM') or ((ValuesProfil[0] = 'ADV') and (ValuesProfil[1] = '')))
    )) then
    SetField('GA_ARTICLEPROFIL', '');

  if GetField('GA_ESTPROFIL') <> 'X' then { L'article n'est pas un profil }
  begin
    SetControlEnabled('GA_ARTICLEPROFIL', False);
    if (CodeProfil <> '') and (ValuesProfil <> nil) then
    begin
      if (ValuesProfil[0] = 'ADV') then
        if (ValuesProfil[1] = '') then
        begin
          SetControlEnabled('GA_ARTICLEPROFIL', True);
        end
        else SetField('GA_ARTICLEPROFIL', ValuesProfil[1]);
    end;
  end
  else { L'article est un profil }
  begin
    SetControlEnabled('GA_ARTICLEPROFIL', False);
  end;

end;
{$ENDIF}

{$IFDEF GPAO}

procedure TOM_Article.MnUniteInit_OnClick(Sender: tObject);
var
  InitUnite: twInitUniteArticle;
begin
  InitUnite := tWInitUniteArticle.Create;
  try
    { Copie le dataSet dans la tob }
    InitUnite.MakeTobFromDS(DS);
    { Montre la fiche }
    InitUnite.OpenForm;
    { Remise à jour du DataSet à partir de la Tob }
    if InitUnite.Valid then InitUnite.PutTobInDS(DS);
  finally
    InitUnite.free;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_Article.MnProperties_OnClick(Sender: TObject);
var
  Prefixe: string;
begin
  Prefixe := TableToPrefixe('ARTICLE');
  wCallProperties(Prefixe, GetField(Prefixe + '_ARTICLE'), wMakeFieldString('ARTICLE', '~'), IntToStr(LongInt(Ecran)));
end;
{$ENDIF}

{$IFDEF GPAO}

procedure TOM_Article.GA_ProfilArticleOnChange(Sender: TObject);
begin
  wProfilArticleState;
end;
{$ENDIF}

/////////////////////////////////////////////
// ****** TOM Conditionnement ***************
/////////////////////////////////////////////

procedure TOM_Conditionnement.OnNewRecord;
begin
  SetControlEnabled('GCO_CODECOND', TRUE);
  SetControlChecked('GCO_ARRONDIINF', False);
  SetControlText('GCO_COEFCONV', '1');
end;

procedure TOM_Conditionnement.OnLoadRecord;
var F: TForm;
  CouleurTrue, CouleurFalse: TColor;
begin
  inherited;
  F := TForm(Ecran);
  CouleurTrue := TEdit(F.FindComponent('GCO_LIBELLE')).Color;
  CouleurFalse := TGroupBox(F.FindComponent('GB_SAISIE')).Color;
  {$IFDEF EAGLCLIENT}
  SetControlText('GCO_NBARTICLE', getfield('GCO_NBARTICLE'));
  if TCheckBox(F.FindComponent('GCO_CHANGEUNITE')).Checked = True then
    {$ELSE}
  if TDBCheckBox(F.FindComponent('GCO_CHANGEUNITE')).Checked = True then
    {$ENDIF}

  begin
    SetControlEnabled('GCO_UNITEAFFICHE', True);
    SetControlEnabled('GCO_COEFCONV', True);
    SetControlEnabled('GCO_ARRONDIINF', True);
    SetControlProperty('GCO_UNITEAFFICHE', 'Color', CouleurTrue);
    SetControlProperty('GCO_COEFCONV', 'Color', CouleurTrue);
  end else
  begin
    SetControlEnabled('GCO_UNITEAFFICHE', False);
    SetControlEnabled('GCO_COEFCONV', False);
    SetControlEnabled('GCO_ARRONDIINF', False);
    SetControlProperty('GCO_UNITEAFFICHE', 'Color', CouleurFalse);
    SetControlProperty('GCO_COEFCONV', 'Color', CouleurFalse);
  end;
end;

/////////////////////////////////////////////
// ****** TOM CodeCPTA ***************
/////////////////////////////////////////////

procedure TOM_CodeCPTA.OnNewRecord;
begin
end;

/////////////////////////////////////////////
// ****** TOM ProfilArt***************
/////////////////////////////////////////////

procedure TOM_ProfilArt.OnArgument(Arguments: string);
var sCalcPAPR: string;
begin
  inherited;
  if (ctxMode in V_PGI.PGIContexte) then
  begin
    // Maj libellés coef et recalcul automatique PA/PR
    sCalcPAPR := GetParamsoc('SO_CALCPAPR');
    if sCalcPAPR = 'AR' then
    begin
      SetControlText('TGPF_COEFCALCPR', TraduireMemoire('C&oef. PA->PR'));
      SetControlText('GPF_CALCAUTOPR', TraduireMemoire('&Recalcul automatique du PR'));
    end
    else if sCalcPAPR = 'RA' then
    begin
      SetControlText('TGPF_COEFCALCPR', TraduireMemoire('C&oef. PR->PA'));
      SetControlText('GPF_CALCAUTOPR', TraduireMemoire('&Recalcul automatique du PA'));
    end;
    // Paramétrage des libellés des tables libres
    GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GPF_LIBREART', 10, '');
    GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GPF_VALLIBRE', 3, '');
    GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GPF_DATELIBRE', 3, '');
    GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GPF_CHARLIBRE', 3, '');
    GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GPF_BOOLLIBRE', 3, '');
  end;
  {$IFDEF CCS3}
  SetControlVisible('GPF_LOT', False);
  SetControlVisible('GPF_NUMEROSERIE', False);
  {$ENDIF}
  {$IFDEF GPAO}
  if (ECRAN <> nil) and (ECRAN.Name = 'WPROFILART') then
  begin
    { Montre ou cache les champs spécifiques de la GP }
    if (THdbValComboBox(GetControl('GPF_TYPEPROFILART')) <> nil) then THdbValComboBox(GetControl('GPF_TYPEPROFILART')).Visible := True;
    if (TTabSheet(GetControl('P_GP')) <> nil) then TTabSheet(GetControl('P_GP')).TabVisible := True;
    { Active premier onglet }
    SetActiveTabSheet('PFAMILLE');
    { }
    if TCheckBox(GetControl('CBSHOWCODE')) <> nil then TCheckBox(GetControl('CBSHOWCODE')).OnClick := CBShowCode_OnCLick;
    { Charge les champs de la table article }
    TobChamp := Tob.Create('WCHAMP', nil, -1);
    { Initialise les listes des champs disponibles }
    if GetListException <> nil then GetListException.OnDrawItem := ListException_DrawItem;
    if GetListChamp <> nil then
    begin
      LoadListChamp;
      PutTobChampInListChamp(GetListChamp.Items);
      GetListChamp.OnDrawItem := ListChamp_DrawItem;
    end;
    { Boutons d'ajout et de Suppression de champ }
    if TToolBarButton97(GetControl('BTADDFIELD')) <> nil then TToolBarButton97(GetControl('BTADDFIELD')).OnClick := BtAddField_OnClick;
    if TToolBarButton97(GetControl('BTDELFIELD')) <> nil then TToolBarButton97(GetControl('BTDELFIELD')).OnClick := BtDelField_OnClick;
    { Boutons de mise à jour des articles associés au profil }
    if TToolBarButton97(GetControl('BTMAJARTICLE')) <> nil then TToolBarButton97(GetControl('BTMAJARTICLE')).OnClick := BtMajArticle_OnClick;
    { bouton d'appel de paramétrage du profil article }
    if GetControl('BPARAMPROFILART') <> nil then TToolBarButton97(GetControl('BPARAMPROFILART')).OnClick := BPARAMPROFILART_OnClick;
    if Assigned(GetControl('GPF_MODECOPIE')) then
      ThDbValComboBox(GetControl('GPF_MODECOPIE')).OnChange := GPF_MODECOPIEChange;
    if Assigned(GetControl('GPF_MODEREMPLACE')) then
      ThDbValComboBox(GetControl('GPF_MODEREMPLACE')).OnChange := GPF_MODEREMPLACEChange;
  end;
  {$ENDIF}
  {$IFDEF BTP}
  THDbValComboBox(GetCOntrol('GPF_TARIFARTICLE')).OnChange := TarifChange;
  {$ENDIF}
   // GPAO DS le 20/03
  // Paramétrage des libellés des tables libres
  GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GPF_LIBREART', 10, '');

  //uniquement en line
  {*
    if (TTabSheet(GetControl('P_CALC_PRIX')) <> nil) then TTabSheet(GetControl('P_CALC_PRIX')).TabVisible := False;
  	if (GetparamsocSecur('SO_GCDESACTIVECOMPTA','X') = True) then
  	begin
    	SetControlVisible('GPF_COMPTAARTICLE', False);
    	SetControlVisible('TGPF_COMPTAARTICLE', False);
  	end;
  *}
end;

procedure TOM_ProfilArt.OnNewRecord;
begin
  SetField('GPF_TYPEPROFILART', 'SIM'); { Profil simple par défaut }
  {$IFDEF GPAO}
    SetField('GPF_MODECOPIE', 'EXC'); { Mode de copie par défaut }
    SetField('GPF_MODEREMPLACE', 'SAF'); { Mode de remplacement par défaut : préserve les différence }
  {$ENDIF}
  {$IFDEF CCS3}
  SetField('GPF_LOT', '-');
  SetField('GPF_NUMEROSERIE', '-');
  {$ENDIF}
  SetField('GPF_REMISELIGNE', 'X');

  //uniquement en line
  {*
  SetField('GPF_COMMISSIONNABL', 'X');
  SetField('GPF_ESCOMPTABLE', 'X');
  SetField('GPF_REMISEPIED', 'X');
  SetField('GPF_TENUESTOCK', '-');
  SetField('GPF_CALCPRIXPR', 'PAA');
  SetField('GPF_CALCAUTOPR', 'X');
  SetField('GPF_CALCPRIXHT', 'DPR');
  SetField('GPF_CALCAUTOHT', 'X');
  SetField('GPF_CALCPRIXTTC', 'AUC');
  *}
end;

procedure TOM_ProfilArt.OnLoadRecord;
begin
  inherited;
  {$IFDEF GPAO}
  { Cache les onglets en fonction du type de profil }
  MasqueOnglets;
  { Active l'onglet par défaut }
  ActiveOnglets;
  SetFocusControl('GPF_PROFILARTICLE');
  { Mise à jour label Inclure/Exclure en fonction du mode de copie }
  SetLabelInclureExclure;
  GPF_MODECOPIEChange(Self);
  { Etat du bouton de mise à jour des champs }
  SetBtnMajArticleState;
  { Recharge la table des champs disponibles }
  if GetListChamp <> nil then PutTobChampInListChamp(GetListChamp.Items);
  { Charge les champs en exception }
  if (GetListException <> nil) then LoadListException(GetListException.Items);
  { Supprime de LISTCHAMP les champs figurant dans LISTEXCEPTION }
  if (GetListChamp <> nil) and (GetListException <> nil) then KillListChampInException(GetListChamp.Items, GetListException.Items);
  SetControlVisible('BPARAMPROFILART', V_PGI.SAV and (GetField('GPF_TYPEPROFILART') = 'ADV'));
  {$ENDIF}

  //THDbValComboBox(GetCOntrol('GPF_SOUSFAMTARART')).Plus := ' AND CC_LIBRE="'+GetField('GPF_TARIFARTICLE')+'"';
  THDbValComboBox(GetCOntrol('GPF_SOUSFAMTARART')).Plus := ' AND BSF_FAMILLETARIF="'+GetField('GPF_TARIFARTICLE')+'"';
  //uniquement en line
  {*
  if TDbCheckBox(GetCOntrol('GPF_TENUESTOCK')) <> nil then TDbCheckBox(GetCOntrol('GPF_TENUESTOCK')).visible := false;
  THValComboBox(GetControl('GPF_CALCPRIXPR')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
  THValComboBox(GetControl('GPF_CALCPRIXHT')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
  THValComboBox(GetControl('GPF_CALCPRIXTTC')).Plus := 'AND (CO_CODE NOT LIKE "S%")'+PlusCalcLine;
  *}

end;

procedure TOM_ProfilArt.OnChangeField(F: TField);
begin
  if (F.FieldName = 'GPF_TENUESTOCK') then
  begin
    if (GetField('GPF_TENUESTOCK') <> 'X') then
    begin
      SetField('GPF_LOT', '-');
      SetField('GPF_NUMEROSERIE', '-');
      SetControlEnabled('GPF_LOT', FALSE);
    end else
    begin
      SetControlEnabled('GPF_LOT', TRUE);
    end;
  end;

  {$IFDEF GPAO}
  if (F.FieldName = 'GPF_TYPEPROFILART') then
    MasqueOnglets;
  {$ENDIF}
end;

procedure TOM_ProfilArt.OnUpdateRecord;
begin
  inherited;
  SetField('GPF_PROFILARTICLE', Trim(GetField('GPF_PROFILARTICLE')));
  {$IFDEF GPAO}
  if GetField('GPF_ARTICLE') <> '' then
  begin
    if not existeSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="' + GetControlText('GPF_ARTICLE') + '"'
      + ' AND GA_ESTPROFIL = "X"') then
      LastError := ErrArticleProfilExist;
  end;
  if lastError = 0 then
  begin
    if ThDBValCombobox(GetControl('GPF_MODECOPIE')).Value = 'INC' then
      if (GetListException <> nil) and (TListBox(GetControl('LISTEXCEPTION')).Items.Count <= 0) then LastError := ErrInclureChamps;
  end;

  if LastError <> 0 then
  begin
    LastErrorMsg := TexteMessageGPAO[LastError];
    SetFocusControl('GPF_ARTICLE');
  end;
  SetControlProperty('BTMAJARTICLE', 'ENABLED', True);
  {$ENDIF}
end;

procedure TOM_ProfilArt.OnClose;
begin
  inherited;
  {$IFDEF GPAO}
  if Assigned(TobChamp) then
    TobChamp.Free;
  {$ENDIF}
end;

{$IFDEF GPAO}
function TOM_ProfilArt.IsAdvancedFilter: Boolean;
begin
  Result := (GetField('GPF_TYPEPROFILART') = 'ADV');
end;
{$ENDIF}

{$IFDEF GPAO}
function TOM_ProfilArt.GetListChamp: tListBox;
begin
  Result := TListBox(GetControl('LISTCHAMP'));
end;
{$ENDIF}

{$IFDEF GPAO}
function TOM_ProfilArt.GetListException: tListBox;
begin
  Result := TListBox(GetControl('LISTEXCEPTION'));
end;
{$ENDIF}

{$IFDEF GPAO}
function TOM_ProfilArt.GetCBShowCode: TCheckBox;
begin
  Result := TCheckBox(GetControl('CBSHOWCODE'));
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.MasqueOnglets;
{ Affiche ou masque les onglets en fonction du type de profil }
var
  AdvancedFilter: Boolean;
begin
  AdvancedFilter := IsAdvancedFilter;
  if (TTabSheet(GetControl('PFAMILLE')) <> nil) then TTabSheet(GetControl('PFAMILLE')).TabVisible := True;
  if (TTabSheet(GetControl('PVALEURPARDEFAUT')) <> nil) then TTabSheet(GetControl('PVALEURPARDEFAUT')).TabVisible := not AdvancedFilter;
  if (TTabSheet(GetControl('P_CALC_PRIX')) <> nil) then TTabSheet(GetControl('P_CALC_PRIX')).TabVisible := not AdvancedFilter;
  if (TTabSheet(GetControl('P_INFO1')) <> nil) then TTabSheet(GetControl('P_INFO1')).TabVisible := not AdvancedFilter;
  if (TTabSheet(GetControl('P_INFO2')) <> nil) then TTabSheet(GetControl('P_INFO2')).TabVisible := not AdvancedFilter;
  if (TTabSheet(GetControl('P_GP')) <> nil) then TTabSheet(GetControl('P_GP')).TabVisible := not AdvancedFilter;
  if (TTabSheet(GetControl('P_AVANCE')) <> nil) then TTabSheet(GetControl('P_AVANCE')).TabVisible := AdvancedFilter;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.ActiveOnglets;
begin
  if (TTabSheet(GetControl('PFAMILLE')) <> nil) then SetActiveTabSheet('PFAMILLE');
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.SetBtnMajArticleState;
begin
  SetControlProperty('BTMAJARTICLE', 'VISIBLE', IsAdvancedFilter);
  SetControlProperty('BTMAJARTICLE', 'ENABLED', DS.State = dsBrowse);
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.GPF_MODECOPIEChange(Sender: tObject);
begin
  SetLabelInclureExclure;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.GPF_MODEREMPLACEChange(Sender: tObject);
begin
  SetBtnMajArticleState;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.SetLabelInclureExclure;
var
  i: Integer;
  Msg: string;
begin
  i := ThDBValCombobox(GetControl('GPF_MODECOPIE')).ItemIndex;
  if (i <> -1) then
    Msg := ThDBValCombobox(GetControl('GPF_MODECOPIE')).Items[i]
  else
    Msg := '';
  if Assigned(GetControl('GBEXCEPTIONCHAMP')) then TgroupBox(GetControl('GBEXCEPTIONCHAMP')).Caption := Msg;
  if Assigned(GetControl('LBEXCEPTIONCHAMP')) then SetControlCaption('LBEXCEPTIONCHAMP', Msg);
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.LoadListChamp;
{ Charge la liste des champs article disponibles }
var
  sRequete: string;
  Q: tQuery;
begin
  TobChamp.ClearDetail;
  sRequete := 'SELECT * FROM WCHAMP WHERE WCA_CONTEXTEPROFIL = "PRO" AND WCA_NOMTABLE="ARTICLE"';
  if ExisteSQL(sRequete) then
  begin
    Q := OpenSql(sRequete, True);
    try
      TobChamp.LoadDetailDB('WCHAMP', '', '', Q, False);
    finally
      Ferme(Q);
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.PutTobChampInListChamp(TS: TStrings);
var
  i: Integer;
begin
  if TS <> nil then
  begin
    TS.Clear;
    for i := 0 to TobChamp.Detail.Count - 1 do
    begin
      { Vérifie si le champ figure dans DECHAMPS et ajoute dans la liste de visu. }
      if WGetLibChamp(TobChamp.Detail[i].GetValue('WCA_NOMCHAMP')) <> '' then TS.Add(TobChamp.Detail[i].GetValue('WCA_NOMCHAMP'));
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.LoadListException(TS: TStrings);
{ Charge la liste des exceptions depuis le champ GPF_LISTECHAMPS }
var
  s, s8: string;
  i: Integer;
begin
  if TS <> nil then
  begin
    TS.Clear;
    s := '';
    for i := 1 to 5 do s := s + GetField('GPF_LISTECHAMPS' + IntToStr(i)) + ';';
    while (s <> '') do
    begin
      s8 := ReadTokenSt(s);
      if (WGetLibChamp(s8) <> '') then TS.Add(s8);
    end;
  end;
end;

{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.KillListChampInException(TSChamp, TSException: TStrings);
var
  i, j: Integer;
begin
  if (TSChamp <> nil) and (TSException <> nil) then
  begin
    for i := 0 to TSException.Count - 1 do
    begin
      j := TSCHamp.IndexOF(TSException[i]);
      if j <> -1 then TSChamp.Delete(j);
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.BtAddField_OnClick(Sender: tObject);
{ Ajoute un champ }
var
  s: string;
  i: Integer;
  ListChamp, ListException: TListBox;
begin
  ListChamp := GetListChamp;
  if (ListChamp <> nil) and (ListChamp.ItemIndex <> -1) then
  begin
    s := ListChamp.Items[ListChamp.ItemIndex];
    ListException := TListBox(GetControl('LISTEXCEPTION'));
    if ListException.Items.IndexOf(s) = -1 then
    begin
      i := ListChamp.ItemIndex;
      ListException.Items.Add(s);
      ListChamp.Items.Delete(ListChamp.ItemIndex);
      if ListChamp.Items.Count > 0 then
      begin
        if i < ListChamp.Items.Count then ListChamp.ItemIndex := i
        else ListChamp.ItemIndex := ListChamp.Items.Count - 1;
      end;
    end;
    SaveListException;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.BtDelField_OnClick(Sender: tObject);
{ Supprime un champ }
var
  s: string;
  i: Integer;
  ListChamp, ListException: TListBox;
begin
  ListException := GetListException;
  if (ListException <> nil) and (ListException.ItemIndex <> -1) then
  begin
    i := ListException.ItemIndex;
    ListChamp := GetListChamp;
    ListChamp.Items.Add(ListException.Items[i]);
    ListException.Items.Delete(i);
    if ListException.Items.Count > 0 then
    begin
      if i < ListException.Items.Count then ListException.ItemIndex := i
      else ListException.ItemIndex := ListException.Items.Count - 1;
    end;
    SaveListException;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}

procedure TOM_ProfilArt.CBShowCode_OnCLick(Sender: tObject);
begin
  GetListChamp.Refresh;
  GetListException.Refresh;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.BtMajArticle_OnClick(Sender: tObject);
{ Appel de la mise à jour des articles liés au profil }
var
  TobArticle: Tob;
  NbrArt: Integer;
begin
  if GetField('GPF_MODEREMPLACE') = 'SAF' then
  begin
    PGIInfo(TexteMessageGPAO[6], 'Profil article');
  end
  else
  begin
    TobArticle := Tob.Create('ARTICLE', nil, -1);
    try
      { Charge la TobArticle }
      if (GetField('GPF_ARTICLE') <> '') and (ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="' + GetField('GPF_ARTICLE') + '"')) then
      begin
        TobArticle.SelectDB('"' + GetField('GPF_ARTICLE') + '"', nil);
        { Appel de la mise à jour }
        NbrArt := wSetAdvancedProfil2(paMajArticleFromProfil, TobArticle);
        if NbrArt < 0 then PGIInfo('Mise à jour des articles effectuée. (' + IntToStr(-1 * NbrArt) + ' article(s) mis à jour)', 'Profil article');
      end
      else
        PGIInfo(TexteMessageGPAO[3], 'Profil article');
    finally
      TobArticle.Free;
    end;
  end;
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.BPARAMPROFILART_OnClick(Sender: tObject);
begin
  AGLLanceFIche('W', 'WCHAMP_FIC', '', '', 'ACTION=MODIFICATION;TABLE=ARTICLE');
  LoadListChamp;
  PutTobChampInListChamp(GetListChamp.Items);
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.ListChamp_DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  sLib: string;
begin
  if (GetListChamp <> nil) and (GetCBShowCode <> nil) and (not GetCBShowCode.Checked) then
    sLib := WGetLibChamp(GetListChamp.Items[Index])
  else
    sLib := GetListChamp.Items[Index];
  GetListChamp.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top, sLib);
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.ListException_DrawItem(Control: TWinControl; Index: Integer; Rect: TRect; State: TOwnerDrawState);
var
  sLib: string;
begin
  if (GetListException <> nil) and (GetCBShowCode <> nil) and (not GetCBShowCode.Checked) then
    sLib := WGetLibChamp(GetListException.Items[Index])
  else
    sLib := GetListException.Items[Index];
  GetListException.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top, sLib);
end;
{$ENDIF}

{$IFDEF GPAO}
procedure TOM_ProfilArt.SaveListException;
var
  sChamps: array[0..4] of string;
  ListException: TListBox;
  i, j: Integer;
  Ok: Boolean;
begin
  ListException := GetListException;
  if ListException <> nil then
  begin
    for i := 0 to 4 do sChamps[i] := '';
    i := -1;
    j := 0;
    Ok := True;
    while (i < ListException.Items.Count - 1) and Ok do
    begin
      Inc(i);
      if (Length(sChamps[j]) + Length(ListException.Items[i]) > 250) then
      begin
        if j < 4 then
          Inc(j)
        else
          Ok := False;
      end;
      if Ok then
      begin
        if sChamps[j] <> '' then sChamps[j] := sChamps[j] + ';';
        sChamps[j] := sChamps[j] + ListException.Items[i];
      end;
    end;
    if Ok then
    begin
      if DS.State = dsBrowse then DS.Edit;
      for i := 1 to 5 do SetField('GPF_LISTECHAMPS' + IntToStr(i), sChamps[i - 1]);
    end
    else
      PGIError(TexteMessageGpao[2], Ecran.Caption);
  end;
end;
{$ENDIF}

procedure AGLLookupFournisseur(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).LookupFournisseur else exit;
end;

//
// Modif BTP 06/02/2001
//

procedure AGLCalculBaseTarifPR(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).CalculBaseTarifPR(Parms[1]) else exit;
end;
//

procedure AGLCalculBaseTarifHT(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  //
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).CalculBaseTarifHT(Parms[1]) else exit;
  //
end;

procedure AGLCalculBaseTarifTTC(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).CalculBaseTarifTTC(Parms[1]) else exit;
end;

procedure AGLCodeUniqueArtLie(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).CodeUniqueArtLie(Parms[1]) else exit;
end;

procedure AGLSetTypeArticle(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).SetTypeARticle(Parms[1]) else exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Patrice ARANEGA
Créé le ...... : 11/02/2000
Modifié le ... : 11/02/2000
Description .. : Cet Article utilise-t-il l'interface affaire de saisie d'articles
Mots clefs ... : ARTICLE;AFFAIRE
*****************************************************************}

function AGLIsFicheAFArticle(parms: array of variant; nb: integer): variant;
var
  TypeArticle: string;
begin
  TypeArticle := string(Parms[0]);
  Result := 0;
  if IsArticleSpecif(TypeArticle) = 'FicheAffaire' then Result := 1;
end;

procedure AGLAffichePhotoArticle(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).AffichePhoto else exit;
end;

////// Dimensions Pour le Script ////////////////////

procedure AGLOnClickParamGrille(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
  ItemDim: THDimensionItem;
  detail: string;
  natureDoc, naturePiece: string;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then
  begin
    with TOM_Article(OM) do
    begin
      if FicheStock then
      begin
        natureDoc := NAT_STOCK;
        naturePiece := NATPIE_STOCK
      end // Fiche stock
      else if FicheStoVte then
      begin
        natureDoc := NAT_STOVTE;
        naturePiece := ''
      end // Fiche stock
      else
      begin
        natureDoc := NAT_ARTICLE;
        naturePiece := NATPIE_ARTICLE
      end; // Fiche article
      itemDim := DimensionsArticle.Dim.CurrentItem;
      ParamGrille(DimensionsArticle, natureDoc, naturePiece);
      detail := CheckToString(not DimensionsArticle.HideUnUsed);
      SetControlText('CBDETAIL', detail);
      AffComboEtab(False);
      if FicheStock or FicheStoVte then AjusterTailleTHDim;
      if itemDim <> nil then ItemDimSetFocus(DimensionsArticle, itemDim, '');
    end
  end;
end;

procedure AGLOnChangeDetail(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).OnChangeDetail()
  else exit;
end;

procedure AGLOnNavigClick(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).OnNavigClick()
  else exit;
end;

procedure AGLOnClickPrixUnique(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then TOM_Article(OM).OnClickPrixUnique()
  else exit;
end;

procedure AGLOnUpdatePrix(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then TOM_Article(OM).OnUpdatePrix(Parms[1])
  else exit;
end;

//
// Modif BTP 06/02/2001
//

procedure AGLOnClickCalcPrixPR(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then
  begin
    TOM_Article(OM).OnClickCalcPrix('GA_CALCPRIXPR', 'GA_COEFFG', 'GA_DPRAUTO', 'GA_DPR', 'T_PRHT');
    // Pour rattraper le coup du onclickcalcprix
    THEDIT(F.FindComponent('GA_DPR')).Visible := not ((THEDIt(F.findcomponent('GA_TYPEARTICLE')).Text = 'NOM') or
      (THEDIt(F.findcomponent('GA_TYPEARTICLE')).Text = 'OUV'));
    THEDIT(F.FindComponent('T_PRHT')).Visible := (THEDIt(F.findcomponent('GA_TYPEARTICLE')).Text = 'NOM') or (THEDIt(F.findcomponent('GA_TYPEARTICLE')).Text =
      'OUV');
  end
  else exit;
end;
//

procedure AGLOnClickCalcPrixHT(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then TOM_Article(OM).OnClickCalcPrix('GA_CALCPRIXHT', 'GA_COEFCALCHT', 'GA_CALCAUTOHT', 'GA_PVHT', 'T_STOCKPVHT')
  else exit;
end;

procedure AGLOnClickCalcPrixTTC(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then TOM_Article(OM).OnClickCalcPrix('GA_CALCPRIXTTC', 'GA_COEFCALCTTC', 'GA_CALCAUTOTTC', 'GA_PVTTC', 'T_STOCKPVTTC')
  else exit;
end;

// Saisi des tarif de base AC

procedure AGLSaisiTarifBase(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article)
    then TOM_Article(OM).SaisiTarifBase(Parms[1])
  else exit;
end;

procedure AGLOnUpdateXXChamp(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_Article) then TOM_Article(OM).OnUpdateXXChamp(Parms[1]) else exit;
end;

procedure AGLLanceEtatArticle(parms: array of variant; nb: integer);
var F: TForm;
  Article, refmarge, valeurmarge, collection, stWhere, Modele: string;
begin
  F := TForm(Longint(Parms[0]));

  // Modif BTP
  {$IFDEF BTP}
  TFFiche(F).BImprimerClick(nil);
  exit;
  {$ENDIF}
  // -----

  if (F.Name <> 'GCARTICLE') and (F.Name <> 'ARTICLE') then exit;
  Article := Trim(Copy(string(Parms[1]), 1, 18));
  refmarge := string(Parms[2]);
  valeurmarge := string(Parms[3]);
  collection := string(Parms[4]);
  stWhere := ' GA_CODEARTICLE="' + Article + '"';
  Modele := GetParamSoc('SO_GCETATARTICLE');
  if Modele <> '' then
    LanceEtat('E', 'GFA', Modele, True, False, False, nil, stWhere, '', False, 0, 'XX_REFMARGE=' + refmarge + '`XX_VALEURMARGE=' + valeurmarge +
      '`XX_COLLECTION=' + collection)
  else
    TFFiche(F).BImprimerClick(nil);
end;

function CataloguAffectationReference(parms: array of variant; nb: integer): variant;
var F: TForm;
  Titre, Ref, Tiers: string;
begin
  Result := '';
  F := TForm(Longint(Parms[0]));
  if F.Name <> 'GCCATALOGU_AFFECT' then exit;
  Titre := F.Caption;
  Ref := Trim(string(Parms[1]));
  Tiers := string(Parms[2]);
  if Ref = '' then exit;
  if HShowMessage('0;' + Titre + ';Confirmez vous l''affectation de cette référence?;Q;YN;N;N;', '', '') <> mrYes then exit;
  Result := Ref + ';' + Tiers;
end;

procedure TOM_ArticlePiece_LanceEtat(parms: array of variant; nb: integer);
var F: TForm;
  Article, stWhere, Modele: string;
begin
  F := TForm(Longint(Parms[0]));
  if F.Name <> 'GCARTICLEPIECE' then exit;
  Article := string(Parms[1]);
  stWhere := ' GAP_ARTICLE="' + Article + '"';
  Modele := 'EXC';
  LanceEtat('E', 'GFA', Modele, True, False, False, nil, stWhere, '', False);
end;

procedure TOM_GtradArticle_LanceEtat(parms: array of variant; nb: integer);
var F: TForm;
  Article, stWhere, Modele: string;
begin
  F := TForm(Longint(Parms[0]));
  if F.Name <> 'GCTRADARTICLE' then exit;
  Article := string(Parms[1]);
  stWhere := ' GTA_ARTICLE="' + Article + '"';
  Modele := 'TRA';
  LanceEtat('E', 'GFA', Modele, True, False, False, nil, stWhere, '', False);
end;

{$IFDEF BTP}

procedure TOM_Article.AfficheMarge(Indice: integer; PrixAchat, PrixVente, Prixrevient: double);
begin
  THNumEdit(Getcontrol('CMARGE')).Value := valeur('0');
  THNumEdit(Getcontrol('PMARGE')).Value := valeur('0');
  THNumEdit(Getcontrol('MMARGE')).Value := valeur('0');
  if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
  begin
    if Prixrevient <> 0 then THNumEdit(Getcontrol('CMARGE')).Value := PrixVente / prixrevient;
    if PrixVente<> 0 then THNumEdit(Getcontrol('PMARGE')).Value := Arrondi(((PrixVente - prixrevient) / PrixVente) * 100,4);
    THNumEdit(Getcontrol('MMARGE')).Value := PrixVente - prixrevient;
  end else
  begin
    if ((Indice = 0) or (Indice = 2) or (Indice = 3)) and (PrixAchat > 0) then
    begin
      THNumEdit(Getcontrol('CMARGE')).Value := PrixVente / prixAchat;
      if Indice = 3 then
        THNumEdit(Getcontrol('PMARGE')).Value := ((PrixVente - prixAchat) / PrixVente) * 100
      else
        THNumEdit(Getcontrol('PMARGE')).Value := ((PrixVente - prixAchat) / PrixAchat) * 100;
      THNumEdit(Getcontrol('MMARGE')).Value := PrixVente - prixAchat;
    end else if Prixrevient > 0 then
    begin
      THNumEdit(Getcontrol('CMARGE')).Value := PrixVente / prixrevient;
      if Indice = 5 then
        THNumEdit(Getcontrol('PMARGE')).Value := ((PrixVente - prixrevient) / Prixvente) * 100
      else
        THNumEdit(Getcontrol('PMARGE')).Value := ((PrixVente - prixrevient) / Prixrevient) * 100;
      THNumEdit(Getcontrol('MMARGE')).Value := PrixVente - prixrevient;
    end;
  end;
end;
{$ENDIF}

procedure AGLZonesLibArtIsModified(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM
  else if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_ARTICLE) then TOM_ARTICLE(OM).ZonesLibArtIsModified else exit;
end;

{$IFDEF BTP}

procedure AGLAppliqueProfilBTP(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM
  else if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_ARTICLE) then TOM_ARTICLE(OM).AppliqueProfilBTP(parms[1]) else exit;
end;
{$ENDIF}

procedure AGLOnChangeStatArt(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM
  else if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_ARTICLE) then TOM_ARTICLE(OM).ChangeStatArt(Parms[1]) else exit;
end;

// *****************************************************************************
// ********************** gestion Isoflex **************************************
// *****************************************************************************

procedure TOM_Article.GereIsoflex;
var bIso: Boolean;
  MenuIso: TMenuItem;
begin
  MenuIso := TMenuItem(GetControl('mnSGED'));
  {$IFNDEF EAGLCLIENT}
  bIso := AglIsoflexPresent;
  {$ELSE}
  bIso := False;
  {$ENDIF}
  if MenuIso <> nil then MenuIso.Visible := bIso;
end;

procedure TOM_ARTICLE_SetFicheModifie (parms : array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM
  else if (F is TFFicheListe) then OM := TFFicheListe(F).OM else exit;
  if (OM is TOM_ARTICLE) then TOM_ARTICLE(OM).SetFicheModifie else exit;
end;

procedure TOM_Article_AppelIsoFlex(parms: array of variant; nb: integer);
var F: TForm;
  Article: string;
begin
  {$IFNDEF EAGLCLIENT}
  F := TForm(Longint(Parms[0]));
  if (F.Name <> 'BTARTICLE') then exit;
  Article := string(Parms[1]);
  AglIsoflexViewDoc(NomHalley, F.Name, 'ARTICLE', 'GA_CLE1', 'GA_ARTICLE', Article, '');
  {$ENDIF}
end;

{$IFDEF BTP}

procedure TOM_Article.DetermineValorisation(var pa, pr, pv: double);
var TypeValo: string;
begin
  {$IFDEF EAGL}
  TypeValo := THValComboBox(GetControl('GA_CALCPRIXHT')).Value;
  {$ELSE}
  TypeValo := THDBValComboBox(GetControl('GA_CALCPRIXHT')).Value;
  {$ENDIF}
  if TypeValo = 'DPA' then
  begin
    pa := THNumEdit(GetControl('GA_DPA')).Value;
    THValComboBox(GetControl('REFMARGE')).ItemIndex := 0;
  end else if TypeValo = 'PAA' then
  begin
    pa := THNumEdit(Getcontrol('GA_PAHT')).Value;
    THValComboBox(GetControl('REFMARGE')).ItemIndex := 0;
  end else if TypeValo = 'FOU' then
  begin
    pa := THNumEdit(GetControl('PAFOUPRINC')).value;
    THValComboBox(GetControl('REFMARGE')).ItemIndex := 0;
  end else if TypeValo = 'DPR' then
  begin
    THValComboBox(GetControl('REFMARGE')).ItemIndex := 1;
  end else
  begin
    pa := THNumedit(GetControl('GA_PAHT')).value;
    THValComboBox(GetControl('REFMARGE')).ItemIndex := 0;
  end;
  pr := GetField('GA_DPR');
  pv := Getfield('GA_PVHT');
end;

procedure TOM_ARTICLE.SetInfosPrixTarif;
var Pu            : Double;
    PRIXPOURQTEAC : Double;
    PrixPourQte   : Double;
    pa            : Double;
    pr            : Double;
    pv            : Double;
    pvc           : Double;
    //
    UA            : String;
    //
    TRefMarge     : THValComboBox;
begin
  if copy(ecran.name,1,9) <> 'BTARTICLE' then exit;
  //
	THlabel(getControl('TUA')).Caption := '';
  //
  PrixPourQteAC := TOBCatalog.GetValue('GCA_PRIXPOURQTEAC');
  if PRixPourQteAC = 0 then PrixPourQteAC := 1;
  //
  PrixPourQte := GetField('GA_PRIXPOURQTE');
  if PrixPourQTe = 0 then PrixPourQTe := 1;
  //
  if ExistCatalog then
  begin
  UA := TOBCatalog.GetString('GCA_QUALIFUNITEACH');
  end;
  
  IF UA = '' then UA := Getfield('GA_QUALIFUNITEACH');

  if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then
    THNumEdit(GetControl('GCA_PRIXBASE')).Value := (TOBCatalog.GetValue('GCA_PRIXBASE') / PrixPourQteAC)
  else
    THNumEdit(GetControl('GCA_PRIXBASE')).Value := (TOBCatalog.GetValue('GCA_PRIXVENTE') / PrixPourQteAC) ;
  //
  THNumEdit(GetControl('GF_PRIXUNITAIRE')).Value := TOBTarif.GetValue('GF_PRIXUNITAIRE');

  if TOBTarif.GetValue('GF_PRIXUNITAIRE') <> 0 then
    Pu := TOBTarif.GetValue('GF_PRIXUNITAIRE')
  else
  begin
    if TOBCatalog.GetString('GCA_ARTICLE')<> '' then
    begin
    if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then
      Pu := (TOBCatalog.GetValue('GCA_PRIXBASE') / prixPourQteAC)
    else
      Pu := (TOBCatalog.GetValue('GCA_PRIXVENTE') / PrixPourQTeAC);
    end else
    begin
      Pu := GeTField('GA_PAUA');
  end;
  end;
  //
  if THlabel(getControl('TUA')) <> nil then
  begin
    SetControlVisible('TUA',true);
    THlabel(getControl('TUA')).Caption := '/'+TOBCatalog.GetValue('GCA_QUALIFUNITEACH');
  end;
  THEdit(GetControl('GF_CALCULREMISE')).text := TOBTarif.GetValue('GF_CALCULREMISE');

  if Pu <> 0 then THEdit(GetCOntrol('GA_PAHT')).Enabled := false
  else if (TFFiche(Ecran).fTypeAction <> TaConsult) then THEdit(GetCOntrol('GA_PAHT')).Enabled := true;

  Pu := PU * (1 - (TOBTarif.GetValue('GF_REMISE') / 100));
  //FV1: 18/02/2015 - Gestion Prix D'achat dans fiche article
  if ExistCatalog then
  begin
  SetControlText('GCA_PAUA',StrS(Pu,V_PGI.okdecP) );
  Pu := PassageUAUV(TOBCatalog, Pu, PrixPourQTe,TOBCatalog.GetValue('GCA_COEFCONVQTEACH'),GetField('GA_COEFCONVQTEVTE'), UA, GetField('GA_QUALIFUNITESTO'), GetField('GA_QUALIFUNITEVTE'));
  end else
  begin
    Pu := PassageUAUV(TOBCatalog, Pu, PrixPourQTe,GetFIeld('GA_COEFCONVQTEACH'),GetField('GA_COEFCONVQTEVTE'), UA, GetField('GA_QUALIFUNITESTO'), GetField('GA_QUALIFUNITEVTE'));
  end;

  //  if (Pu <> 0) and (Pu <> GetField('GA_PAHT')) then
  if (Pu <> 0) {and (Pu <> valeur(GetControltext('GA_PAHT_')))} then
  begin
    if (TFFiche(Ecran).fTypeAction <> TaConsult) and (not (DS.State in [dsInsert, dsEdit])) then
    begin
      DS.EDit;
      SetField('GA_PAHT', Pu);
    end else
    begin
      {$IFDEF BTP}
      if (DS.State in [dsInsert, dsEdit]) then SetField('GA_PAHT', Pu);

      if THNumEdit(GetControl('GA_PAHT_')) <> nil then
      begin
        THNumEdit(GetControl('GA_PAHT_')).visible := true;
        THEdit(GetControl('GA_PAHT')).visible := false;
        THNumEdit(GetControl('GA_PAHT_')).Value := Pu;
      end;

      if THNumEdit(GetControl('GA_PAHT_1')) <> nil then
      begin
        THNumEdit(GetControl('GA_PAHT_1')).visible := true;
        THEdit(GetControl('GA_PAHT1')).visible := false;
        THNumEdit(GetControl('GA_PAHT_1')).Value := Pu;
      end;

      TRefMarge := THValComboBox(GetControl('REFMARGE'));

      if THNumEdit(GetControl('GA_DPR_')) <> nil then
      begin
        if GetField('GA_COEFFG') <> 0 then Pr := pu * GetField('GA_COEFFG')
                                      else Pr := GetField('GA_DPR');
        THNumEdit(GetControl('GA_DPR_')).visible := true;
        THEdit(GetControl('GA_DPR')).visible := false;
        THNumEdit(GetControl('GA_DPR_')).Value := Pr;
      end;

      if THNumEdit(GetControl('GA_DPR_1')) <> nil then
      begin
        THNumEdit(GetControl('GA_DPR_1')).visible := true;
        THEdit(GetControl('GA_DPR1')).visible := false;
        THNumEdit(GetControl('GA_DPR_1')).Value := Pr;
      end;

      if THNumEdit(GetControl('GA_PVHT_')) <> nil then
      begin
        Pv := GetField('GA_PVHT');
        if CALCAUTOHT.Checked then
        begin
          if GetField ('GA_COEFCALCHT') <> 0 then Pv := pr * GetField('GA_COEFCALCHT');
        end;

        THNumEdit(GetControl('GA_PVHT_')).visible := true;
        THEdit(GetControl('GA_PVHT')).visible := false;
        THNumEdit(GetControl('GA_PVHT_')).Value := Pv;
      end;

      if THNumEdit(GetControl('GA_PVHT_1')) <> nil then
      begin
        THNumEdit(GetControl('GA_PVHT_1')).visible := true;
        THEdit(GetControl('GA_PVHT1')).visible := false;
        THNumEdit(GetControl('GA_PVHT_1')).Value := Pv;
      end;

//
      if THNumEdit(GetControl('GA_PVTTC_')) <> nil then
      begin
        if GetField ('GA_COEFCALCTTC') <> 0 then PvC := pv * GetField('GA_COEFCALCTTC')
                                           else PvC := GetField('GA_PVTTC');
        THNumEdit(GetControl('GA_PVTTC_')).visible := true;
        THEdit(GetControl('GA_PVTTC')).visible := false;
        THNumEdit(GetControl('GA_PVTTC_')).Value := PvC;
      end;

      if THNumEdit(GetControl('GA_PVTTC_1')) <> nil then
      begin
        THNumEdit(GetControl('GA_PVTTC_1')).visible := true;
        THEdit(GetControl('GA_PVTTC1')).visible := false;
        THNumEdit(GetControl('GA_PVTTC_1')).Value := Pvc;
      end;
//
      AfficheMarge(TRefMarge.ItemIndex, pu, pv, pr);
      {$ENDIF}
    end;
  end;

end;

procedure TOM_ARTICLE.GetInfosPrixTarif(CodeFournisseur, CodeArticle, TarifArticle, SousTarifArticle: string; TOBcatalog, TOBTarif: TOB);
var Req: string;
  QQ: Tquery;
  TOBART: TOB;
begin
  //if not(DS.State in [dsEdit]) and (TFFiche(Ecran).fTypeAction=taCreat) then exit;
  //if TFFiche(Ecran).fTypeAction = TaCreat then exit;

  TOBCatalog.InitValeurs;

  TOBTarif.initValeurs; // Tarif fournisseur principal

  TOBArt := TOB.Create('ARTICLES', nil, -1);
  TOBArt.GetEcran(Ecran);

  Req := 'SELECT * FROM CATALOGU WHERE GCA_ARTICLE="' + CodeArticle + '" AND GCA_TIERS="' + CodeFournisseur + '"';
  QQ := OpenSql(Req, true);
  if not QQ.eof then
  begin
    TOBCatalog.SelectDB('', QQ);
  end;
  ferme(QQ);

  Req := 'SELECT * FROM TIERS WHERE T_TIERS="' + CodeFournisseur + '"';
  QQ := OpenSql(Req, true);
  TOBTiers.SelectDB('', QQ);
  ferme(QQ);
  TobArt.GetEcran(Ecran);
  //
  GetTarifGlobal(CodeArticle, TarifArticle,SousTarifArticle, 'ACH', TOBART, TOBTiers, TOBTarif, true);
  TOBART.free;
  ExistCatalog := (TOBCatalog.GetString('GCA_ARTICLE')<>'');
end;

procedure TOM_Article.SetDetailPrixPose(Article, Libelle: string; QtePrest: double; TOBPrestation: TOB);
var TOBD: TOB;
  indice: integer;
begin
  if DS.state = dsInsert then exit;
  if TOBEDetail.GetValue('GNE_NOMENCLATURE') = '' then
  begin
    TOBEDetail.putValue('GNE_NOMENCLATURE', Article);
    TOBEDetail.putValue('GNE_LIBELLE', Libelle);
    TOBEDetail.putValue('GNE_ARTICLE', Article);
    TOBEDetail.putValue('GNE_DATECREATION', V_PGI.DateEntree);
    TOBEDetail.putValue('GNE_QTEDUDETAIL', 1);
  end;
  if TOBDetail.detail.count = 0 then
  begin
    // Ligne Article
    TOBD := TOB.Create('NOMENLIG', TOBDetail, -1);
    TOBD.putvalue('GNL_NOMENCLATURE', Article);
    TOBD.putvalue('GNL_NUMLIGNE', 1);
    TOBD.putValue('GNL_LIBELLE', Libelle);
    TOBD.PutValue('GNL_CODEARTICLE', Copy(Article, 1, 18));
    TOBD.PutValue('GNL_QTE', 1);
    TOBD.PutValue('GNL_JOKER', 'N');
    TOBD.PutValue('GNL_ARTICLE', Article);
    TOBD.PutValue('GNL_DATECREATION', V_PGI.DateEntree);
    TOBD.PutValue('GNL_CREATEUR', V_PGI.User);
    TOBD.PutValue('GNL_UTILISATEUR', V_PGI.User);
  end;

  // Gestion de la prestation associée
  if (TOBPrestation.GetValue('GA_ARTICLE') = '') and (TOBDetail.detail.count = 2) then
  begin
    TOBDetail.detail[1].free; // on enleve la ligne de prestation si plus de prestation associée
  end;
  if (TOBPrestation.GetValue('GA_ARTICLE') <> '') and (TOBDetail.detail.count = 1) then
  begin
    TOBD := TOB.Create('NOMENLIG', TOBDetail, -1);
    TOBD.putvalue('GNL_NOMENCLATURE', Article);
    TOBD.putvalue('GNL_NUMLIGNE', 2);
    TOBD.putValue('GNL_LIBELLE', TOBPrestation.GetValue('GA_LIBELLE'));
    TOBD.PutValue('GNL_CODEARTICLE', Copy(TOBPrestation.GetValue('GA_ARTICLE'), 1, 18));
    TOBD.PutValue('GNL_QTE', 0);
    TOBD.PutValue('GNL_JOKER', 'N');
    TOBD.PutValue('GNL_ARTICLE', TOBPrestation.GetValue('GA_ARTICLE'));
    TOBD.PutValue('GNL_DATECREATION', V_PGI.DateEntree);
    TOBD.PutValue('GNL_CREATEUR', V_PGI.User);
    TOBD.PutValue('GNL_UTILISATEUR', V_PGI.User);
  end;
  if (TOBDetail.detail.count = 2) then
  begin
    if (QtePrest >= 0) then TOBDetail.detail[1].putvalue('GNL_QTE', QtePrest);
    TOBDetail.detail[1].putValue('GNL_LIBELLE', TOBPrestation.GetValue('GA_LIBELLE'));
    TOBDetail.detail[1].PutValue('GNL_CODEARTICLE', Copy(TOBPrestation.GetValue('GA_ARTICLE'), 1, 18));
    TOBDetail.detail[1].PutValue('GNL_DATECREATION', V_PGI.DateEntree);
  end;
end;

procedure TOM_ARTICLE.AffichageInfosPrixPose;
var PAMO, PRMO, PVHTMO, PVTTCMO: double;
    PAART,PRART,PVART,PCART : double;
    ZonePa,ZonePR,ZONEPVHT,ZonePVTTC : THNumedit;
begin
  if (TFFiche(Ecran).fTypeAction <> TaConsult) then
  begin
(*    PAART := VALEUR(GetField('GA_PAHT'));
    PRART := valeur(GetField('GA_DPR'));
    PVART := valeur(GetField('GA_PVHT'));
    PCART := valeur(GetField('GA_PVTTC'));
*)
    PAART := Valeur(GetControlText ('GA_PAHT1'));
    PRART := Valeur(getControlText('GA_DPR1'));
    PVART := Valeur(getControlText('GA_PVHT1'));
    PCART := Valeur(getControlText('GA_PVTTC1'));
    //
    ZONEPA := THNumEdit(GetControl('TOTALPA'));
    ZONEPR := THNumEdit(GetControl('TOTALPR'));
    ZONEPVHT := THNumEdit(GetControl('TOTALPVHT'));
    ZONEPVTTC := THNumEdit(GetControl('TOTALPVTTC'));
  end else
  begin
    PAART := ThNumedit(getControl('GA_PAHT_')).value;
    PRART := ThNumedit(getControl('GA_DPR_')).value;
    PVART := ThNumedit(getControl('GA_PVHT_')).value;
    PCART := ThNumedit(getControl('GA_PVTTC_')).value;
    ZONEPA := THNumEdit(GetControl('TOTALPA_'));
    ZONEPR := THNumEdit(GetControl('TOTALPR_'));
    ZONEPVHT := THNumEdit(GetControl('TOTALPVHT_'));
    ZONEPVTTC := THNumEdit(GetControl('TOTALPVTTC_'));
    THNumEdit(GetControl('TOTALPA')).visible := false;
    THNumEdit(GetControl('TOTALPR')).visible := false;
    THNumEdit(GetControl('TOTALPVHT')).visible := false;
    THNumEdit(GetControl('TOTALPVTTC')).visible := false;
    ZONEPA.Visible := true;
    ZONEPR.Visible := true;
    ZONEPVHT.Visible := true;
    ZONEPVTTC.Visible := true;
  end;

  THLabel(GetControl('LPRESTATION')).Caption := '';
  THLabel(GetControl('LTPS')).Caption := '';
  // Prestation
  THNumEdit(GetControl('TPAMO')).Value := TOBPrestation.getValue('GA_PAHT');
  PAMO := TOBPrestation.getValue('GA_PAHT');
  THNumEdit(GetControl('TPRMO')).Value := TOBPrestation.getValue('GA_DPR');
  PRMO := TOBPrestation.getValue('GA_DPR');
  THNumEdit(GetControl('TPVHTMO')).Value := TOBPrestation.getValue('GA_PVHT');
  PVHTMO := TOBPrestation.getValue('GA_PVHT');
  THNumEdit(GetControl('TPVTTCMO')).Value := TOBPrestation.getValue('GA_PVTTC');
  PVTTCMO := TOBPrestation.getValue('GA_PVTTC');

  if PAMO > 0 then THNumEdit(GetControl('TCOEFFGMO')).value := PRMO / PAMO;
  if PRMO > 0 then THNumEdit(GetControl('TCOEFMARGMO')).value := PVHTMO / PRMO;

  THLabel(GetControl('LPRESTATION')).Caption := TOBPrestation.GetValue('GA_LIBELLE');
  TPrestation.Text := TOBPrestation.GetValue('GA_CODEARTICLE');
  TPrestation.Refresh;
  if TOBPrestation.GetValue('GA_QUALIFUNITEVTE') <> '' then
    THLabel(GetControl('LTPS')).Caption := rechdom('GCQUALUNITTOUS', TOBPrestation.GetValue('GA_QUALIFUNITEVTE'), false);
  DQTEPREST.value := 0;
  if TOBDetail.detail.count = 2 then DQTEPREST.value := TOBDetail.detail[1].GetValue('GNL_QTE');

  ZonePA.value := (PAMO * DQTEPrest.value) + PAART;
  ZONEPR.value := (PRMO * DQTEPrest.Value) + PRART;
  ZONEPVHT.value := (PVHTMO * DQTEPrest.Value) + PVART;
  ZONEPVTTC.value := (PVTTCMO * DQTEPrest.Value) + PCART;
  if ZONEPA.value > 0 then
  begin
    THNumEdit(GetControl('COEFFGTOT')).value := ZONEPR.value / ZONEPA.value;
  end;
  if ZONEPR.value > 0 then
  begin
    THNumEdit(GetControl('COEFMARGTOT')).value := ZONEPVHT.value / ZONEPR.value;
  end;
end;

procedure TOM_Article.TPrestationChange(Sender: TOBject);
var TOBART: TOB;
  QQ: TQuery;
begin
  TOBART := TOB.Create('ARTICLE', nil, -1);
  QQ := OpenSql('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="' + GetControlText('TPRESTATION') + '"', true);
  TOBART.selectdb('', QQ);
  ferme(QQ);
  ChangePrestation(TOBART.GetValue('GA_ARTICLE'), DQTEPREST.Value);
  TOBART.free;
end;

procedure TOM_Article.TPrestationElipsClick(Sender: TOBject);
var MaCle, Prestation: string;
begin
  MaCle := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', 'ACTION=MODIFICATION;XX_WHERE= AND GA_TYPEARTICLE="PRE"');
  if (MaCle = 'Abandon') or (MaCle = '') then Prestation := '';
  if (MaCle <> '') and (MaCle <> 'Abandon') then Prestation := ReadTokenst(MaCle);
  TPrestation.Text := Prestation;
  ChangePrestation(Prestation, DQTEPREST.Value);
end;

procedure TOM_Article.ChangeQtePrest(Sender: TObject);
begin
  SetDetailPrixPose(GetField('GA_ARTICLE'), GetField('GA_LIBELLE'), DQtePrest.Value, TobPrestation);
  AffichageInfosPrixPose;
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
end;

procedure TOM_Article.ChangePrestation(Prestation: string; Qte: double);
begin
  chargeTobPrestation(Prestation, TOBPrestation);
  SetDetailPrixPose(GetField('GA_ARTICLE'), GetField('GA_LIBELLE'), Qte, TobPrestation);
  AffichageInfosPrixPose;
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
end;

procedure TOM_Article.ReajusteTOBDetail;
var NEWTOB: TOB;
begin
  TOBEDetail.PutValue('GNE_NOMENCLATURE', GetField('GA_ARTICLE'));
  TOBEDetail.PutValue('GNE_LIBELLE', GetField('GA_LIBELLE'));
  TOBEDetail.PutValue('GNE_ARTICLE', GetField('GA_ARTICLE'));
  TOBEDetail.PutValue('GNE_QTEDUDETAIL', 1);
  //
  TOBDEtail.ClearDetail;
  //
  NewTOB := TOB.Create('NOMENLIG', TOBDetail, -1);
  NewTOB.PutValue('GNL_NOMENCLATURE', GetField('GA_ARTICLE'));
  NewTOB.PutValue('GNL_NUMLIGNE', 1);
  NewTOB.PutValue('GNL_LIBELLE', GetField('GA_LIBELLE'));
  NewTOB.PutValue('GNL_QTE', 1);
  NewTOB.PutValue('GNL_CODEARTICLE', GetField('GA_CODEARTICLE'));
  NewTOB.PutValue('GNL_ARTICLE', GetField('GA_ARTICLE'));
  NewTOB.PutValue('GNL_JOKER', 'N');
  NewToB.PutValue('GNL_DATECREATION', V_PGI.DateEntree);
  NewToB.PutValue('GNL_CREATEUR', V_PGI.User);
  NewToB.PutValue('GNL_UTILISATEUR', V_PGI.User);
  //
  if TOBPrestation.GetValue('GA_ARTICLE') <> '' then
  begin
    NewTOB := TOB.Create('NOMENLIG', TOBDetail, -1);
    NewTOB.PutValue('GNL_NOMENCLATURE', GetField('GA_ARTICLE'));
    NewTOB.PutValue('GNL_NUMLIGNE', 2);
    NewTOB.PutValue('GNL_LIBELLE', TOBPrestation.GetValue('GA_LIBELLE'));
    NewTOB.PutValue('GNL_QTE', DQtePrest.Value);
    NewTOB.PutValue('GNL_CODEARTICLE', TOBPrestation.GetValue('GA_CODEARTICLE'));
    NewTOB.PutValue('GNL_ARTICLE', TOBPrestation.GetValue('GA_ARTICLE'));
    NewTOB.PutValue('GNL_JOKER', 'N');
    NewToB.PutValue('GNL_DATECREATION', V_PGI.DateEntree);
    NewToB.PutValue('GNL_CREATEUR', V_PGI.User);
    NewToB.PutValue('GNL_UTILISATEUR', V_PGI.User);
  end;
end;

procedure TOM_Article.CodeTarifArticleChange(Sender: Tobject);
var CodeSStarif : string;
begin
  if (copy(ecran.name, 1, 9) = 'BTARTICLE') and
//    ((TypeArticle = 'MAR') or (TypeArticle = 'ARP')) then
	(Pos(TypeArticle,'ARP;MAR;PRE;') > 0) then
  begin
  	if CodeSOusTarif <> NIL then CodeSSTarif := CodeSousTarif.value else CodeSSTarif := '';
  	SetChangeCodeTarif (THValComboBox(GetControl('GA_TARIFARTICLE')).Value);
    GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'), THValComboBox(GetControl('GA_TARIFARTICLE')).Value,
      CodeSStarif, TOBcatalog, TOBTarif);
    SetInfosPrixTarif;
    DefiniAffichage;
    if Typearticle = 'ARP' then AffichageInfosPrixPose;
  end;
end;

procedure TOM_Article.MnTariffouClick(sender: Tobject);
var ActionSuiv: TActionFiche;
		CodeSStarif : string;
begin
  if (DS.State in [dsInsert]) then exit;
  if (GetField('GA_FOURNPRINC') = '') then exit;
  if TFFiche(Ecran).fTypeAction <> taConsult then ActionSuiv := taModif
  else ActionSuiv := taConsult;
  SaisieTarifCliArtKnown('FOU', GetField('GA_FOURNPRINC'), Getfield('GA_ARTICLE'), ActionSuiv, false);
  if (copy(ecran.name, 1, 9) = 'BTARTICLE') and
	   (Pos(TypeArticle,'ARP;MAR;PRE;') > 0) then
  begin
	  if CodeSousTarif <> nil then CodeSStarif := CodeSousTarif.value else CodeSSTarif := '';
    GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'),
      GetField('GA_TARIFARTICLE'),CodeSStarif, TOBcatalog, TOBTarif);
    SetInfosPrixTarif;
    DefiniAffichage;
    if Typearticle = 'ARP' then AffichageInfosPrixPose;
  end;
end;

{$ENDIF} //FIN IFDEF BTP

{$IFDEF AFFAIRE}
procedure TOM_Article.AFFormuleVariable; //Affaire-ONYX
{$IFDEF EAGLCLIENT}
var ED_Formule: THEdit;
  {$ELSE}
var ED_Formule: THDBEdit;
  {$ENDIF}
begin
  if (copy(ecran.name, 1, 9) = 'AFARTICLE') then
  begin
    {$IFDEF EAGLCLIENT}
    ED_Formule := THEdit(GetControl('GA_FORMULEVAR'));
    {$ELSE}
    ED_Formule := THDBEdit(GetControl('GA_FORMULEVAR'));
    {$ENDIF}
    if (IsArticleSpecif(TypeArticle) = 'FicheAffaire') and GetParamSoc('SO_AFVARIABLES') then
    begin
      if ED_Formule <> nil then
      begin
        ED_Formule.visible := true;
        SetControlVisible('TGA_FORMULEVAR', true);
        ED_Formule.OnElipsisClick := AFFormuleVariableClick;
      end;
    end
    else if ED_Formule <> nil then
    begin
      ED_Formule.visible := false;
      SetControlVisible('TGA_FORMULEVAR', false);
    end;
  end;
end;

procedure TOM_Article.AFFormuleVariableClick(sender: Tobject);
var StFormule: string;
begin
  StFormule := GetControlText('GA_FORMULEVAR');
  if StFormule = '' then
  begin
    StFormule := AGLLanceFiche('AFF', 'AFORMULEVAR_RECH', '', StFormule, '');
    if StFormule <> '' then
    begin
      DS.edit;
      SetField('GA_FORMULEVAR', StFormule);
    end;
  end else
    AglLanceFiche('AFF', 'AFORMULEVAR', '', '', 'AVF_FORMULEVAR=' + StFormule + ';ACTION=MODIFICATION');
end;
{$ENDIF}

procedure AFLanceFiche_ArticleGA(lequel, argument: string);
begin
  AGLLanceFiche('AFF', 'AFARTICLE', '', Lequel, Argument);
end;

procedure AFLanceFiche_ArticlePouGA(lequel, argument: string);
begin
  AGLLanceFiche('AFF', 'AFARTPOURCENT', '', Lequel, Argument);
end;

// Gestion des métrés
{$IFDEF BTP}
procedure TOM_Article.OuvreNomenclature(Sender: Tobject);
var NbDeclinaison : integer;
begin

  if not FileExists(RepPoste) then
  begin
    if FileExists(RepServeur) then
      MetreArticle.CopieLocaltoServeur(RepServeur, RepPoste)
    else
      MetreArticle.CopieLocaltoServeur(MetreArticle.fFichierVide, RepPoste);
  end;           

  AGLLanceFiche('GC', 'GCNOMENENT', Getfield ('GA_ARTICLE'), '', 'ARTICLE=' + Getfield('GA_ARTICLE')+';'+actionToString(TFFiche(Ecran).FTypeAction));

end;

procedure TOM_ARTICLE.OuvreNomenParc (sender : Tobject);
begin
	AGLLanceFiche('GC', 'GCNOMENENT', Getfield ('GA_ARTICLE'), '', 'ARTICLE=' + Getfield('GA_ARTICLE')+';'+actionToString(TFFiche(Ecran).FTypeAction));
end;

{***********A.G.L.***********************************************
Auteur  ...... : franck Vautrain
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Ouverture de la fiche des variables/constantes pour la
Suite ........ : gestion de métrés
Mots clefs ... : VARIABLES/CONSTANTES
*****************************************************************}
procedure TOM_Article.OuvreVariable(Sender: Tobject);
begin

  AGLLanceFiche('BTP', 'BTVARIABLE', 'B;'+ Getfield ('GA_ARTICLE'),'', 'LIBELLEARTICLE=' + Getfield('GA_LIBELLE') + '; TYPEARTICLE=' + Getfield('GA_TYPEARTICLE'));
  
end;

procedure TOM_Article.ExcelOnClick(Sender: Tobject);
begin

  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit

  //On copie le fichier de l'article se trouvant sur le serveur sur le poste Local
  if not FileExists(RepPoste) then
  begin
    if FileExists(RepServeur) then
      MetreArticle.CopieLocaltoServeur(RepServeur, RepPoste)
    else
      MetreArticle.CopieLocaltoServeur(MetreArticle.fFichierVide, RepPoste);     //On charge un fichier vide...
  end;

  MetreArticle.OuvrirMetreXLs;

end;

procedure TOM_Article.PrixPourQteExit(Sender: TObject);
begin
  SetInfosPrixTarif ;
  DefiniAffichage;
end;

{$ENDIF}

// suppression des bouton magneto alors que monofiche = false
procedure TOM_Article.AfterFormShow;
begin
//

  TToolbarButton97(GetControl('BLast')).visible := false;
  TToolbarButton97(GetControl('BFirst')).visible := false;
  TToolbarButton97(GetControl('BNext')).visible := false;
  TToolbarButton97(GetControl('BPrev')).visible := false;

end;

{$IFDEF BTP}
procedure TOM_Article.LoadTobArticlecompl;
begin
	if CodeSousTarif = nil then exit;
	if TobZonelibre = nil then TOBZoneLibre := TOB.Create ('ARTICLECOMPL',nil,-1);
  if not TobZonelibre.SelectDB('"' + GetField('GA_ARTICLE') + '"', nil) then TobZonelibre.InitValeurs;
  CodeSousTarif.Value := TOBZoneLibre.GetValue('GA2_SOUSFAMTARART');
end;

procedure TOM_Article.EcritTobArticleCompl;
begin
	if TOBZoneLibre = nil then exit;
  if TobZonelibre.IsFieldModified ('GA2_SOUSFAMTARART') then
  begin
  	TOBZoneLibre.putvalue('GA2_ARTICLE',GetField('GA_ARTICLE'));
    TOBZoneLibre.PutValue('GA2_CODEARTICLE', GetField('GA_CODEARTICLE'));
    TOBZoneLibre.SetAllModifie (true);
    TobZonelibre.InsertOrUpdateDB;
  	if (GetField('GA_STATUTART') = 'GEN') and (GetField('GA_DIMMASQUE') <> '')  then
    begin
      ExecuteSql ('UPDATE ARTICLECOMPL SET GA2_SOUSFAMTARART="'+TOBZoneLibre.getValue('GA2_SOUSFAMTARART')+'" WHERE '+
      						'GA2_CODEARTICLE="'+GetField('GA_CODEARTICLE')+'"');
    end;

  end;
end;

procedure TOM_Article.EcritTobArticleCompParc;
begin
	TOBARTICLEPARC.GetEcran(ecran);
  if TOBARTICLEPARC.IsOneModifie (true) then
  begin
    if TOBARTICLEPARC.Getvalue('BCP_ARTICLE')='' then
    begin
      TOBARTICLEPARC.PutValue('BCP_ARTICLE',GetField ('GA_ARTICLE'));
      TOBARTICLEPARC.PutValue('BCP_CODEARTICLE',GetField ('GA_CODEARTICLE'));
    end;
    TOBARTICLEPARC.InsertOrUpdateDB;
  end;
end;

procedure TOM_Article.SetChangeCodeTarif(CodeTarif: string);
begin

	if CodeSousTarif = nil then exit;
  //CodeSousTarif.value := '';
  CodeSousTarif.Plus := ' AND BSF_FAMILLETARIF="'+CodeTarif+'"';

end;

{$ENDIF}

procedure TOM_Article.SousFamTarArtChange(sender: TObject);
var CodeSStarif : string;
begin
	if CodeSousTarif = nil then exit;
	TobZonelibre.PutValue('GA2_SOUSFAMTARART',GetControltext('GA2_SOUSFAMTARART'));
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
	if Pos(TypeArticle,'ARP;MAR;PRE;') > 0 then
  begin
	  if CodeSousTarif <> nil then CodeSStarif := CodeSousTarif.value else CodeSSTarif := '';
    GetInfosPrixTarif(GetField('GA_FOURNPRINC'),
      GetField('GA_ARTICLE'),
      THValComboBox(GetControl('GA_TARIFARTICLE')).Value,
      CodeSStarif,
      TOBcatalog,
      TOBTarif);
    SetInfosPrixTarif;
    DefiniAffichage;
    if Typearticle = 'ARP' then AffichageInfosPrixPose;
  end;
end;

{$IFDEF BTP}
procedure TOM_ProfilArt.TarifChange (Sender : Tobject);
var TheValeur : string;
begin
	TheValeur := THDbValComboBox(GetCOntrol('GPF_TARIFARTICLE')).value;
  THDbValComboBox(GetCOntrol('GPF_SOUSFAMTARART')).value := '';
  //THDbValComboBox(GetCOntrol('GPF_SOUSFAMTARART')).Plus := ' AND CC_LIBRE="'+TheValeur+'"';
  THDbValComboBox(GetCOntrol('GPF_SOUSFAMTARART')).Plus := ' AND BSF_FAMILLETARIF="'+TheValeur+'"';
end;
{$ENDIF}

procedure TOM_Article.LoadTobArticlecompParc;
	procedure InitArticleParc (OneTOB : TOB);
  begin
  	TOBARTICLEPARC.InitValeurs;
  	TOBARTICLEPARC.PutValue('BCP_ARTICLE',GetField ('GA_ARTICLE'));
  	TOBARTICLEPARC.PutValue('BCP_CODEARTICLE',GetField ('GA_CODEARTICLE'));
  	TOBARTICLEPARC.PutValue('BCP_ISVERSION','-');
  	TOBARTICLEPARC.PutValue('BCP_NBMOISGARANTIE',12);
  	TOBARTICLEPARC.PutValue('BCP_REFFABRICANT','');
  	TOBARTICLEPARC.PutValue('BCP_QTEZERO','-');
  end;
var QQ : Tquery;
begin
	if TOBARTICLEPARC = nil then TOBARTICLEPARC  := TOB.Create ('ARTICLECOMPPARC',nil,-1)
  												else TOBARTICLEPARC.InitValeurs;
	if TFFiche(Ecran).fTypeAction <> taCreat then
  begin
  	QQ := OpenSql ('SELECT * FROM ARTICLECOMPPARC WHERE BCP_ARTICLE="'+GetField ('GA_ARTICLE')+'"',true,1,'',true);
    if not QQ.eof then
    begin
  		TOBARTICLEPARC.SelectDB('"' + GetField('GA_ARTICLE') + '"', nil);
    end else
    begin
			SetFicheModifie;
    	InitArticleParc (TOBARTICLEPARC);
    end;
  end else
  begin
		SetFicheModifie;
  	InitArticleParc (TOBARTICLEPARC);
  end;
  THEDit(GetControl('BCP_CODEINTERNE')).Text  := TOBARTICLEPARC.GetValue('BCP_CODEINTERNE');
  THEDit(GetControl('BCP_CODEINTERNE')).OnChange := CODEINTERNEChange;
  THSpinEdit(GetControl('BCP_NBMOISGARANTIE')).Value := TOBARTICLEPARC.GetValue('BCP_NBMOISGARANTIE');
  THSpinEdit(GetControl('BCP_NBMOISGARANTIE')).OnChange  := NbMoisGarantieChange;
  ThCheckBox(GetControl('BCP_ISVERSION')).Checked  := (TOBARTICLEPARC.GetValue('BCP_ISVERSION')='X');
  ThCheckBox(GetControl('BCP_ISVERSION')).OnClick  := IsVersionClick;
  ThCheckBox(GetControl('BCP_QTEZERO')).Checked  := (TOBARTICLEPARC.GetValue('BCP_QTEZERO')='X');
  ThCheckBox(GetControl('BCP_QTEZERO')).OnClick  := IsVersionClick;
  ThEdit(GetControl('BCP_REFFABRICANT')).text  := TOBARTICLEPARC.GetValue('BCP_REFFABRICANT');
  ThEdit(GetControl('BCP_REFFABRICANT')).OnChange := CODEINTERNEChange;
end;

procedure TOM_Article.SetFicheModifie;
begin
	if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
end;

procedure TOM_Article.CODEINTERNEChange(Sender: TObject);
begin
  SetFicheModifie;
end;

procedure TOM_Article.NbMoisGarantieChange(Sender: Tobject);
begin
  SetFicheModifie;
end;

procedure TOM_Article.IsVersionClick(Sender: TObject);
begin
  SetFicheModifie;
end;

procedure TOM_Article.ARTPARCCLick(Sender: TObject);
var retour : string;
begin
	retour := AglLanceFiche('BTP','BTARTPARC_RECH','','','ACTION=MODIFICATION');
  if retour <> '' then
  begin
    SetFicheModifie;
    SetField('GA_ARTICLEPARC',copy(retour,1,18));
  end;
end;

procedure TOM_Article.DuplicFromArt(Sender: TObject);
var ART : THCritMaskEdit;
    Article,stfiche,stWhere : string;
    QQ : TQuery;
begin
	Article := '';
	ART := THCritMaskEdit.create(ecran);
  ART.Parent := ecran;
  ART.Width := 3; ART.Visible := False;
  ART.DataType := 'GCARTICLE';
  ART.Text := '';
  StFiche := '';
  StWhere := GetTypeArticleBTP;
  StWhere := 'XX_WHERE=AND '+stWhere;
  DispatchRecherche (ART, 1, '',stWhere, stFiche);
  if ART.Text <> '' then Article := ART.Text;
  ART.free;
  if Article = '' then exit;
  QQ := OpenSql ('SELECT * FROM ARTICLE WHERE GA_ARTICLE="'+Article+'"',true,1,'',true);
  if not QQ.eof then
  begin
  	SetField ('GA_CODEARTICLE','P'+QQ.findField('GA_CODEARTICLE').AsString);
  	SetField ('GA_LIBELLE',QQ.findField('GA_LIBELLE').AsString);
  	SetField ('GA_BLOCNOTE',QQ.findField('GA_BLOCNOTE').AsString);
  end;
  ferme (QQ);
end;

function TOM_ARTICLE.RechercheQuotite(Unite : string) : double;
Var StSql : string;
    QQ    : TQuery;
begin

  Result := 1;

  StSQL := 'Select GME_QUOTITE From MEA where GME_MESURE="' + Unite + '"';

  QQ    := OpenSql(StSql, True);

  if not QQ.Eof then Result := Valeur(QQ.FindField('GME_QUOTITE').AsString);
  //
  Ferme(QQ);

end;

//FV1 : 12/02/2015 - Recalcul du Prix d'achat en UV à partir du Prix d'achat en UA
Function TOM_ARTICLE.CalculPAHT : Double;
Var UV        : String;
    UA        : String;
    US        : string;
    //
    MtPAF     : Double;
    PAUA      : Double;
    PAUV      : double;
    CoefUAUS  : Double;
    CoefUSUV  : Double;
    PrixPourQte : double;
begin

  Result := 0;
  (*
  If GetField('GA_FOURNPRINC') <> '' then
  begin
    Result := Valeur(GetField('GA_PAHT'));
    exit;
  end;
  *)
  PAUA := Valeur(GetField('GA_PAUA'));
  PAUV := Valeur(Getfield('GA_PAHT'));
  //
  UV   := Getfield('GA_QUALIFUNITEVTE');
  US   := Getfield('GA_QUALIFUNITESTO');
  //
  if (ExistCatalog) and (GetField('GA_FOURNPRINC')<> '') then
  begin
    UA        := TOBCatalog.GetString('GCA_QUALIFUNITEACH');
    CoefUAUS  := TOBCatalog.GetValue('GCA_COEFCONVQTEACH');
  end Else
  begin
    UA        := Getfield('GA_QUALIFUNITEACH');
    CoefUAUS  := Getfield('GA_COEFCONVQTEACH');
  end;
  //
  CoefUSUV  := Getfield('GA_COEFCONVQTEVTE');
  //
  PrixPourQte := GetField('GA_PRIXPOURQTE');
  //
  MTPAF := PAUA;

  Result := PassageUAUV(TOBCatalog, MTPAF, PrixPourQte, CoefUAUS, CoefUSUV, UA, US, UV);
  if Result = 0 then result := PAUV;

end;

Procedure TOM_ARTICLE.GestionFournisseurPrincipal;
Var CodeFrs     : String;
    CodeArt     : String;
    CodeTar     : string;
    CodeSSTarif : string;
begin

  if GetControl('GA_FOURNPRINC') = nil then Exit;

  CodeFrs := GetField('GA_FOURNPRINC');
  CodeArt := GetField('GA_ARTICLE');
  CodeTar := GetField('GA_TARIFARTICLE');

  if CodeSousTarif <> nil then CodeSSTarif := CodeSousTarif.Value;


  if CodeFrs = '' then
  begin
    //FV1 : remise à zéro de l'ensemble des zones associées àu fournisseur principal
    TOBCatalog.InitValeurs(false);
    SetControlText('FOURNPRINC', '');
    SetControlText('GF_PRIXUNITAIRE','0.00');
    SetControlText('GCA_PRIXBASE','0.00');
    SetControlText('TUA','');
    SetControlText('GF_CALCULREMISE','0.00');
    //FV1 : 19/02/2015 - Gestion du prix d'achat au niveau de la fiche Article
    //SetControlEnabled('GA_QUALIFUNITEACH', True);
    SetField('GA_PAHT', CalculPAHT);
  end;
  if GetControlText('GA_QUALIFUNITEACH')<>'' then
  begin
    SetControlEnabled('GA_COEFCONVQTEACH', True);
  end;
  // DCA - Gestion Commerciale PGI 10790
  SetControlText('FOURNPRINC', RechDom('GCTIERS', CodeFrs, False));

  if (copy(ecran.name, 1, 9) = 'BTARTICLE')  then
  begin
    if not DuplicFromMenu   then GetInfosPrixTarif(CodeFrs, CodeArt, CodeTar,  CodeSSTarif, TOBcatalog, TOBTarif);
    SetInfosPrixTarif;
    if Typearticle = 'ARP'  then AffichageInfosPrixPose;
    if ExistCatalog then
    begin
      SetControlEnabled('GA_PAUA', False);
      //SetControlEnabled('GA_QUALIFUNITEACH', False);
      //SetControlEnabled('GA_COEFCONVQTEACH', False);
    end else
    begin
      if GetControlText('FOURNPRINC') <> '' then
      begin
        SetControlEnabled('GA_PAUA', True);
        //SetControlEnabled('GA_QUALIFUNITEACH', TRUE);
        //SetControlEnabled('GA_COEFCONVQTEACH', tRUE);
  end;
    end;
  end;
  ArticlePrin := GetField('GA_ARTICLE');

  // DCA - Gestion Commerciale PGI 10790
  if BMonoFournisseur then FournPrinc := CodeFrs;

  PoseEvenementFournisseur; // Gestion de l'évênement OnElipsisClick

end;


//FV1 : 09/03/2015

procedure TOM_Article.BCODEBARREOnClick(Sender : Tobject);
var trouve      : Boolean;
    CodeArticle : string;
    Nature      : String;
    Action      : string;
begin

  CodeArticle := Getfield('GA_ARTICLE');

  if not ExisteSQL('SELECT BCB_IDENTIFCAB FROM BTCODEBARRE WHERE BCB_NATURECAB="' + TypeArticle + '" and BCB_IDENTIFCAB="'+ CodeArticle +'"') then
    Action := 'ACTION=CREATION'
  else
    Action := 'ACTION=MODIFICATION';

  AGLLanceFiche('BTP','BTCODEBARRE','','', Action + ';ID='+TypeArticle+';CODE='+ CodeArticle);

  ChargeCodeBarre;

end;

Procedure TOM_Article.ChargeCodeBarre;
var StSQL  : string;
    QQ     : TQuery;
begin

  //Au retour on vérifie si un code principal et on l'affiche
  StSQL := 'SELECT * FROM BTCODEBARRE ';
  StSQL := StSQL + ' WHERE BCB_NATURECAB = "' +  TypeArticle + '" ';
  StSQL := StSQL + ' AND   BCB_IDENTIFCAB= "' +  Getfield('GA_ARTICLE') + '" ';
  StSQL := StSQL + ' AND   BCB_CABPRINCIPAL= "X" ';

  QQ:=OpenSql(StSQL, True) ;

  if not QQ.Eof then
  begin
    SetField('GA_CODEBARRE',       QQ.FindField('BCB_CODEBARRE').AsString);
    Setfield('GA_QUALIFCODEBARRE', QQ.FindField('BCB_QUALIFCODEBARRE').AsString);
  end;

  Ferme(QQ) ;

end;

//FV1 : 09/04/2015

procedure TOM_Article.BACTUPRIX_OnClick(Sender : Tobject);
begin

  If DPRAUTO.checked      then CalculBaseTarifPR(True);
  If CALCAUTOHT.Checked   then CalculBaseTarifHT(True);
  If CALCAUTOTTC.checked  then CalculBaseTarifTTC(True);

end;

procedure TOM_Article.DPRAUTO_OnClick(Sender : Tobject);
begin

  If DPRAUTO.checked      then CalculBaseTarifPR(False);
  If CALCAUTOHT.Checked   then CalculBaseTarifHT(False);
  If CALCAUTOTTC.checked  then CalculBaseTarifTTC(False);

end;

procedure TOM_Article.CALCAUTOHT_OnClick(Sender : Tobject);
begin

  If CALCAUTOHT <> nil then if CALCAUTOHT.Checked   then CalculBaseTarifHT(False);
  If CALCAUTOTTC <> nil then If CALCAUTOTTC.checked  then CalculBaseTarifTTC(False);

end;

procedure TOM_Article.CALCAUTOTTC_OnClick(Sender : Tobject);
begin
  If CALCAUTOTTC.checked  then CalculBaseTarifTTC(False);

end;

procedure TOM_Article.SelCatalog(Sender: Tobject);
begin
  LookUpFournisseur;
end;

procedure TOM_Article.SelFourn(Sender: Tobject);
var MaCle : String;
    CodeSStarif,CodFourn,CodArtic : string;
begin
	if CodeSousTarif <> nil then CodeSStarif := CodeSousTarif.value else CodeSStarif := '';
  CodFourn := GetField('GA_FOURNPRINC');
  CodArtic := GetField('GA_ARTICLE');
  Macle := DispatchRecherche(THEdit(GetControl('GA_FOURNPRINC')),2,'T_NATUREAUXI = "FOU"','','');
  if (MaCle = 'Abandon') or (MaCle = '') and (GetControlText('GA_FOURNPRINC') = '') then FournPrinc := '';
  if (MaCle <> '') and (MaCle <> 'Abandon') then FournPrinc := ReadTokenst(MaCle);
  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
  SetControlText('GA_FOURNPRINC', FournPrinc);
  {$IFDEF BTP}
  if copy(ecran.name, 1, 9) = 'BTARTICLE' then
  begin
    GetInfosPrixTarif(GetField('GA_FOURNPRINC'), GetField('GA_ARTICLE'),
      GetField('GA_TARIFARTICLE'), CodeSStarif, TOBcatalog, TOBTarif);
    SetInfosPrixTarif;
    if TypeArticle = 'ARP' then AffichageInfosPrixPose;
  end;
  {$ENDIF}
  DefiniAffichage;
end;

procedure TOM_Article.DefiniAffichage;
begin
  SetControlProperty('TLBLUNITEACH', 'visible', false);
  SetControlProperty('TLBLCOEFACH', 'visible', false);
  if GetField('GA_FOURNPRINC')<>'' then
  begin
    if ExistCatalog then
    begin
      SetControlProperty('GA_PAUA', 'visible', False);
      SetControlProperty('GCA_PAUA', 'visible', true);
      SetControlEnabled ('GCA_PAUA', false);
      //
      //SetControlProperty('GA_QUALIFUNITEACH', 'visible', False);
      //SetControlProperty('GA_COEFCONVQTEACH', 'visible', False);
      SetControlProperty('TLBLUNITEACH', 'visible', true);
      SetControlProperty('TLBLCOEFACH', 'visible', true);
      if TOBCatalog.GetString('GCA_QUALIFUNITEACH') <> '' then
      begin
        SetControlText('TLBLUNITEACH', RechDom('GCQUALUNITTOUS',TOBCatalog.GetString('GCA_QUALIFUNITEACH'),false));
      end;
      {*
      if TOBCatalog.GetValue('GCA_PRIXBASE') <> 0 then
        SetControlText('GCA_PAUA',StrS(TOBCatalog.GetDouble('GCA_PRIXBASE'),V_PGI.okdecP) )
      else
        SetControlText('GCA_PAUA',StrS(TOBCatalog.GetDouble('GCA_PRIXVENTE'),V_PGI.okdecP) );
      *}
      //
      SetControltext('TLBLCOEFACH', StrS(TOBCatalog.getDouble('GCA_COEFCONVQTEACH'),4));
    end else
    begin
      SetControlProperty('GA_PAUA', 'visible', true);
      SetControlProperty('GCA_PAUA', 'visible', false);
      //
      //SetControlProperty('GA_QUALIFUNITEACH', 'enabled', true);
      //SetControlProperty('GA_QUALIFUNITEACH', 'visible', true);
      SetControlProperty('GA_COEFCONVQTEACH', 'visible', True);
      SetControlProperty('TLBLUNITEACH', 'visible', false);
      SetControlProperty('TLBLCOEFACH', 'visible', false);
    end;
  end else
  begin
    SetControlText('TLBLUNITEACH', '<aucun>');
    SetControltext('TLBLCOEFACH', StrS(0,4));
//    SetControlProperty('GA_QUALIFUNITEACH', 'visible', true);
//    SetControlProperty('GA_COEFCONVQTEACH', 'visible', True);
//    SetControlProperty('GA_QUALIFUNITEACH', 'enabled', False);
//    SetControlProperty('GA_COEFCONVQTEACH', 'enabled', False);
    SetControlProperty('TLBLUNITEACH', 'visible', false);
    SetControlProperty('TLBLCOEFACH', 'visible', false);
    SetControlEnabled ('GA_PAUA',false);
    SetControlProperty('GCA_PAUA', 'visible', false);
    SetControlVisible ('GA_PAUA',true);
  end;
end;

procedure TOM_Catalogu.BTCONDITClick(Sender: TObject);
begin
  if GetField('GCA_ARTICLE')= '' then Exit;
  if GetField('GCA_TIERS')= '' then Exit;
  AGLLanceFiche('BTP','BTCONDITIONNEMENT','ACH;'+GetField('GCA_TIERS')+';'+GetField('GCA_ARTICLE'),'','ACTION=MODIFICATION');
end;

initialization
  {$IFNDEF CCS3}
  registerclasses([TOM_ArticlePiece]);
  {$ENDIF}
  registerclasses([TOM_Catalogu, TOM_ProfilArt, TOM_ArticleLie]);
  registerclasses([TOM_Article, TOM_Conditionnement {, TOM_Dimension}]);
  registerclasses([TOM_CodeCpta]);
  RegisterAglProc('LookupFournisseur', TRUE, 0, AGLLookupFournisseur);
  RegisterAglProc('CalculBaseTarifHT', TRUE, 1, AGLCalculBaseTarifHT);
  RegisterAglProc('CalculBaseTarifTTC', TRUE, 1, AGLCalculBaseTarifTTC);
  RegisterAglProc('CodeUniqueArtLie', TRUE, 1, AGLCodeUniqueArtLie);
  RegisterAglProc('SetTypeARticle', TRUE, 1, AGLSetTypeARticle);
  RegisterAglFunc('IsFicheAFArticle', False, 1, AGLIsFicheAFArticle);
  RegisterAglProc('AffichePhotoArticle', TRUE, 0, AGLAffichePhotoArticle);
  RegisterAglProc('OnClickParamGrille', TRUE, 0, AGLOnClickParamGrille);
  RegisterAglProc('OnChangeDetail', TRUE, 0, AGLOnChangeDetail);
  RegisterAglProc('LanceEtatArticle', TRUE, 1, AGLLanceEtatArticle);
  RegisterAglFunc('CataloguAffectReference', True, 2, CataloguAffectationReference); // Catalogu_AFFECT
  RegisterAglProc('TOM_ArticlePiece_LanceEtat', TRUE, 1, TOM_ArticlePiece_LanceEtat);
  RegisterAglProc('TOM_GtradArticle_LanceEtat', TRUE, 1, TOM_GtradArticle_LanceEtat);
  RegisterAglProc('OnNavigClick', TRUE, 1, AGLOnNavigClick);
  RegisterAglProc('OnClickPrixUnique', TRUE, 0, AGLOnClickPrixUnique);
  RegisterAglProc('OnUpdatePrix', TRUE, 0, AGLOnUpdatePrix);
  RegisterAglProc('OnClickCalcPrixHT', TRUE, 0, AGLOnClickCalcPrixHT);
  RegisterAglProc('OnClickCalcPrixTTC', TRUE, 0, AGLOnClickCalcPrixTTC);
  RegisterAglProc('OnUpdateXXChamp', TRUE, 0, AGLOnUpdateXXChamp);
  RegisterAglProc('SaisiTarifBase', TRUE, 1, AGLSaisiTarifBase);
  RegisterAglProc('TOM_Article_AppelIsoFlex', TRUE, 1, TOM_Article_AppelIsoFlex);
  RegisterAglProc('SetFicheModifie',True,0,TOM_ARTICLE_SetFicheModifie);
  //
  // Modif BTP 06/02/2001
  //
  RegisterAglProc('CalculBaseTarifPR', TRUE, 1, AGLCalculBaseTarifPR);
  RegisterAglProc('OnClickCalcPrixPR', TRUE, 0, AGLOnClickCalcPrixPR);
  //---------
  RegisterAglProc('ZonesLibArtIsModified', TRUE, 0, AGLZonesLibArtIsModified);
  RegisterAglProc('OnChangeStatArt', TRUE, 0, AGLOnChangeStatArt);
  {$IFDEF BTP}
  RegisterAglProc('AppliqueProfilBTP', TRUE, 2, AGLAppliqueProfilBTP);
  {$ENDIF}
end.


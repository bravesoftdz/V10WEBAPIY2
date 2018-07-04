unit SaisieParamGpao;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, HTB97, ExtCtrls, DicoAf, Spin, ComCtrls, Hent1, UTOB, ed_formu, PrintDBG,
  Grids, Mask, HmsgBox, DBTables, SaisieValeurGPAO, HSysMenu, Hctrls, UIUtil, HPanel,
  SaisieCorrespGPAO;

type
  TFParamGPAO = class(TForm)
    Panel1: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    GRE_TYPEREPRISE: THValComboBox;
    GRE_TYPEDONNEES: THValComboBox;
    Panel3: TPanel;
    GChamp: THGrid;
    BFlecheDroite: TToolbarButton97;
    BFlecheGauche: TToolbarButton97;
    Panel4: TPanel;
    BValider: TToolbarButton97;
    TLChamp: THGrid;
    Label3: TLabel;
    GRE_NOMTABLE: THCritMaskEdit;                                                                             
    LibNomTable: TLabel;
    Identifiant: TLabel;
    GRE_DEBUTENREG: TEdit;
    GRE_LONGUEURFIXE: TCheckBox;
    Label4: TLabel;
    GRE_SEPARENREG: TEdit;
    LibelleNomTable: TLabel;
    TButtonDefaut: TToolbarButton97;
    ButtonDupliq: TToolbarButton97;
    BFerme: TToolbarButton97;
    BFlecheHaut: TToolbarButton97;
    BFlecheBas: TToolbarButton97;
    BFormule: TToolbarButton97;
    FindLigne: TFindDialog;
    BChercher: TToolbarButton97;
    BTableCorr: TToolbarButton97;
    BImprime: TToolbarButton97;
    GRE_ENTETELIG: TCheckBox;
    procedure BValiderClick(Sender: TObject);
    procedure GRE_NOMTABLEExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BFlecheDroiteClick(Sender: TObject);
    procedure BFlecheGaucheClick(Sender: TObject);
    procedure GRE_TYPEDONNEESExit(Sender: TObject);
    procedure TButtonDefautClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure ButtonDupliqClick(Sender: TObject);
    procedure BFlecheHautClick(Sender: TObject);
    procedure BFlecheBasClick(Sender: TObject);
    procedure BFormuleClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BTableCorrClick(Sender: TObject);
    procedure GChampRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BImprimeClick(Sender: TObject);
    procedure GRE_ENTETELIGClick(Sender: TObject);
  private
    Procedure InitAffichage ;
    function  ControleDonnees : boolean ;
    procedure InitialiseListeChamps;
    Procedure ChargeParametrage ;
    Procedure AfficheParametrage ;
    Procedure SauvegardeParametrage ;
    Procedure SupprimeAncienParametrage ;
    procedure AjoutChampSpecifiqueTable (NomTable : string; TTOB_REPRISE : TOB);
    { Déclarations privées }
  public
    Action      : TActionFiche ;
    FindFirst   : Boolean ;
    TypeReprise : string ;
    TypeDonnees : string ;
    Inventaire  : boolean;
    Prefixe     : string ;
    TobDeChamps : TOB    ;
    Tob_Entete  : TOB    ;
    Tob_Reprise : TOB    ;
    { Déclarations publiques }
  end;


Procedure AppelParamGPAO (Action : TActionFiche; TypeReprise, TypeDonnees : string; Inventaire : boolean);
Procedure SupprimeParamGPAO (TypeReprise, TypeDonnees : string);
//
// Définition des colonnes du THGRID
//
Const   COL_Champ     = 0;
        COL_PosDebut  = 1;
        COL_Longueur  = 2;
        COL_TableCorr = 3;


///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques pour la table ARTICLE
///////////////////////////////////////////////////////////////////////////////
ChampTobArticle : array[1..50,1..3] of String 	= (
   ('$$_LIBCOLLECTION'   , 'Libellé collection'                 , 'VARCHAR(35)'  ),
   ('$$_FAMILLENIV4'     , 'Famille niveau 4'                   , 'COMBO'        ),
   ('$$_FAMILLENIV5'     , 'Famille niveau 5'                   , 'COMBO'        ),
   ('$$_FAMILLENIV6'     , 'Famille niveau 6'                   , 'COMBO'        ),
   ('$$_FAMILLENIV7'     , 'Famille niveau 7'                   , 'COMBO'        ),
   ('$$_FAMILLENIV8'     , 'Famille niveau 8'                   , 'COMBO'        ),
   ('$$_LIBFAMILLE1'     , 'Libellé famille niveau 1'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE2'     , 'Libellé famille niveau 2'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE3'     , 'Libellé famille niveau 3'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE4'     , 'Libellé famille niveau 4'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE5'     , 'Libellé famille niveau 5'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE6'     , 'Libellé famille niveau 6'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE7'     , 'Libellé famille niveau 7'           , 'VARCHAR(35)'  ),
   ('$$_LIBFAMILLE8'     , 'Libellé famille niveau 8'           , 'VARCHAR(35)'  ),
   ('$$_LIBGRILLE1'      , 'Libellé grille dimension 1'         , 'VARCHAR(35)'  ),
   ('$$_LIBGRILLE2'      , 'Libellé grille dimension 2'         , 'VARCHAR(35)'  ),
   ('$$_LIBGRILLE3'      , 'Libellé grille dimension 3'         , 'VARCHAR(35)'  ),
   ('$$_LIBGRILLE4'      , 'Libellé grille dimension 4'         , 'VARCHAR(35)'  ),
   ('$$_LIBGRILLE5'      , 'Libellé grille dimension 5'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM1'    , 'Code dimension 1 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM2'    , 'Code dimension 2 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM3'    , 'Code dimension 3 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM4'    , 'Code dimension 4 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM5'    , 'Code dimension 5 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_LIBGPAODIM1'     , 'Libellé dimension 1 dans GPAO'      , 'VARCHAR(35)'  ),
   ('$$_LIBGPAODIM2'     , 'Libellé dimension 2 dans GPAO'      , 'VARCHAR(35)'  ),
   ('$$_LIBGPAODIM3'     , 'Libellé dimension 3 dans GPAO'      , 'VARCHAR(35)'  ),
   ('$$_LIBGPAODIM4'     , 'Libellé dimension 4 dans GPAO'      , 'VARCHAR(35)'  ),
   ('$$_LIBGPAODIM5'     , 'Libellé dimension 5 dans GPAO'      , 'VARCHAR(35)'  ),
   ('$$_STATART1'        , 'Statistique complémentaire 1'       , 'VARCHAR(9)'   ),
   ('$$_LIBSTATART1'     , 'Libellé Statistique compl. 1'       , 'VARCHAR(35)'  ),
   ('$$_STATART2'        , 'Statistique complémentaire 2'       , 'VARCHAR(9)'   ),
   ('$$_LIBSTATART2'     , 'Libellé Statistique compl. 3'       , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART1'    , 'Libellé statistique article 1'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART2'    , 'Libellé statistique article 2'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART3'    , 'Libellé statistique article 3'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART4'    , 'Libellé statistique article 4'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART5'    , 'Libellé statistique article 5'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART6'    , 'Libellé statistique article 6'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART7'    , 'Libellé statistique article 7'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART8'    , 'Libellé statistique article 8'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREART9'    , 'Libellé statistique article 9'      , 'VARCHAR(35)'  ),
   ('$$_LIBLIBREARTA'    , 'Libellé statistique article 10'     , 'VARCHAR(35)'  ),
   ('$$_COLLECTIONBAS'   , 'Collection de base'                 , 'VARCHAR(3)'   ),
   ('$$_CODEDEVISEPV'    , 'Code devise des PV'                 , 'VARCHAR(3)'   ),
   ('$$_CODEDEVISEPA'    , 'Code devise des PA'                 , 'VARCHAR(3)'   ),
   ('$$_CODEDEVISEPR'    , 'Code devise des PR'                 , 'VARCHAR(3)'   ),
   ('$$_CHEMINPHOTO'     , 'Répertoire et nom de la photo JPEG' , 'VARCHAR(70)'  ),
   ('$$_CALCULCLECB'     , 'Calcul de la clé du code à barres'  , 'BOOLEAN'      ),
   ('$$_MAJPARTIELLE'    , 'Mise à jour partielle'              , 'BOOLEAN'      ) );

///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques pour la table DISPO
///////////////////////////////////////////////////////////////////////////////
ChampTobDispo : array[1..16,1..3] of String 	= (
   ('$$_CODEBARRE'       , 'Code à barres'                      , 'VARCHAR(18)'  ),
   ('$$_CODEARTICLE'     , 'Code article générique'             , 'VARCHAR(18)'  ),
   ('$$_CODEGPAODIM1'    , 'Code dimension 1 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM2'    , 'Code dimension 2 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM3'    , 'Code dimension 3 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM4'    , 'Code dimension 4 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM5'    , 'Code dimension 5 dans GPAO'         , 'VARCHAR(35)'  ),
   ('$$_RECHGPAOLIBDIM1' , 'Codes GPAO 1 dans les libellés'     , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM2' , 'Codes GPAO 2 dans les libellés'     , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM3' , 'Codes GPAO 3 dans les libellés'     , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM4' , 'Codes GPAO 4 dans les libellés'     , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM5' , 'Codes GPAO 5 dans les libellés'     , 'BOOLEAN'      ),
   ('$$_CODEDEVISEPA'    , 'Code devise des PA'                 , 'VARCHAR(3)'   ),
   ('$$_CODEDEVISEPR'    , 'Code devise des PR'                 , 'VARCHAR(3)'   ),
   ('$$_RAZSTOCK'        , 'Supprime les stocks d''un dépôt'    , 'BOOLEAN'      ),
   ('$$_MAJPARTIELLE'    , 'Mise à jour partielle'              , 'BOOLEAN'      ) );


///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques pour la table TARIF
///////////////////////////////////////////////////////////////////////////////
ChampTobTarif : array[1..15,1..3] of String 	= (
   ('$$_CODEBARRE'       , 'Code à barres'                        , 'VARCHAR(18)'  ),
   ('$$_CODETYPE'        , 'Code du tarif'                        , 'VARCHAR(6)'   ),
   ('$$_CODEPERIODE'     , 'Période de tarif'                     , 'VARCHAR(6)'   ),
   ('$$_CODEARTICLE'     , 'Code article générique'               , 'VARCHAR(18)'  ),
   ('$$_CODEGPAODIM1'    , 'Code dimension 1 dans GPAO'           , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM2'    , 'Code dimension 2 dans GPAO'           , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM3'    , 'Code dimension 3 dans GPAO'           , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM4'    , 'Code dimension 4 dans GPAO'           , 'VARCHAR(35)'  ),
   ('$$_CODEGPAODIM5'    , 'Code dimension 5 dans GPAO'           , 'VARCHAR(35)'  ),
   ('$$_RECHGPAOLIBDIM1' , 'Codes GPAO dans les libellés'         , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM2' , 'Codes GPAO 2 dans les libellés'       , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM3' , 'Codes GPAO 3 dans les libellés'       , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM4' , 'Codes GPAO 4 dans les libellés'       , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM5' , 'Codes GPAO 5 dans les libellés'       , 'BOOLEAN'      ),
   ('$$_PRIXARTICLE'     , 'Prix à reporter dans la fiche article', 'BOOLEAN'      ));


///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques pour la table TIERS
///////////////////////////////////////////////////////////////////////////////
ChampTobTiers : array[1..24,1..3] of String 	= (
   ('$$_MAJPARTIELLE'     , 'Mise à jour partielle'              , 'BOOLEAN'      ),
   ('$$_TABLELIBRETIERS1' , 'Code table libre tiers 1'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS2' , 'Code table libre tiers 2'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS3' , 'Code table libre tiers 3'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS4' , 'Code table libre tiers 4'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS5' , 'Code table libre tiers 5'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS6' , 'Code table libre tiers 6'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS7' , 'Code table libre tiers 7'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS8' , 'Code table libre tiers 8'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS9' , 'Code table libre tiers 9'           , 'VARCHAR(6)'   ),
   ('$$_TABLELIBRETIERS10', 'Code table libre tiers 10'          , 'VARCHAR(6)'   ),
   ('$$_TEXTELIBRE1'      , 'Texte libre tiers 1'                , 'VARCHAR(35)'  ),
   ('$$_TEXTELIBRE2'      , 'Texte libre tiers 2'                , 'VARCHAR(35)'  ),
   ('$$_TEXTELIBRE3'      , 'Texte libre tiers 3'                , 'VARCHAR(35)'  ),
   ('$$_VALLIBRE1'        , 'Valeur libre tiers 1'               , 'DOUBLE'       ),
   ('$$_VALLIBRE2'        , 'Valeur libre tiers 2'               , 'DOUBLE'       ),
   ('$$_VALLIBRE3'        , 'Valeur libre tiers 3'               , 'DOUBLE'       ),
   ('$$_DATELIBRE1'       , 'Date libre tiers 1'                 , 'DATE'         ),
   ('$$_DATELIBRE2'       , 'Date libre tiers 2'                 , 'DATE'         ),
   ('$$_DATELIBRE3'       , 'Date libre tiers 3'                 , 'DATE'         ),
   ('$$_BOOLLIBRE1'       , 'Booléan libre tiers 1'              , 'BOOLEAN'      ),
   ('$$_BOOLLIBRE2'       , 'Booléan libre tiers 2'              , 'BOOLEAN'      ),
   ('$$_BOOLLIBRE3'       , 'Booléan libre tiers 3'              , 'BOOLEAN'      ),
   ('$$_CREECLICASH'      , 'Client 999999 CASH 2000'            , 'BOOLEAN'      ));

///////////////////////////////////////////////////////////////////////////////
// Définition des champs spécifiques pour la table dimension
///////////////////////////////////////////////////////////////////////////////
ChampTobDimension : array[1..1,1..3] of String 	= (
   ('$$_GCCATEGORIEDIM'     , 'Numéro de la dimension'             , 'COMBO' ));

////////////////////////////////////////////////////////////////////
// Noms des champs permettant de récupérer un document
// Un document, c'est une entête, des lignes, des pieds ....
// Bref, c'est trop complexes et il y a trop de champs pour s'en sortir
// -> La reprise est donc simplifiée
/////////////////////////////////////////////////////////////////////
ChampTobDocument : array[1..72,1..3] of String 	= (
   ('GP_NATUREPIECEG'    , 'Nature de document'                , 'COMBO'        ),
   ('GP_DATEPIECE'       , 'Date de la pièce'                  , 'DATE'         ),
   ('GP_DATELIVRAISON'   , 'Date de livraison'                 , 'DATE'         ),
   ('GP_REFINTERNE'      , 'Référence interne'                 , 'VARCHAR (35)' ),
   ('GP_REFEXTERNE'      , 'Référence externe'                 , 'VARCHAR (35)' ),
   ('GP_TIERS'           , 'Tiers'                             , 'VARCHAR (17)' ),
   ('GP_TIERSLIVRE'      , 'Tiers livré'                       , 'VARCHAR (17)' ),
   ('GP_TIERSFACTURE'    , 'Tiers facturé'                     , 'VARCHAR (17)' ),
   ('GP_TIERSPAYEUR'     , 'Tiers payeur'                      , 'VARCHAR (17)' ),
   ('GP_ETABLISSEMENT'   , 'Etablissement'                     , 'COMBO'        ),
   ('GP_VENTEACHAT'      , 'Type de document'                  , 'COMBO'        ),
   ('GP_VIVANTE'         , 'Pièce vivante'                     , 'BOOLEAN'      ),
   ('GP_REGIMETAXE'      , 'Régime fiscal'                     , 'COMBO'        ),
   ('GP_CAISSE'          , 'Code caisse'                       , 'COMBO'        ),
   ('GP_DEPOTDEST'       , 'Dépot destinataire'                , 'COMBO'        ),
   ('$$_CODEBARRE'       , 'Code à barres'                     , 'VARCHAR (18)' ),
   ('$$_CODEARTICLE'     , 'Code article'                      , 'VARCHAR (18)' ),
   ('$$_CODEGPAODIM1'    , 'Code dimension 1 dans GPAO'        , 'VARCHAR (35)' ),
   ('$$_CODEGPAODIM2'    , 'Code dimension 2 dans GPAO'        , 'VARCHAR (35)' ),
   ('$$_CODEGPAODIM3'    , 'Code dimension 3 dans GPAO'        , 'VARCHAR (35)' ),
   ('$$_CODEGPAODIM4'    , 'Code dimension 4 dans GPAO'        , 'VARCHAR (35)' ),
   ('$$_CODEGPAODIM5'    , 'Code dimension 5 dans GPAO'        , 'VARCHAR (35)' ),
   ('$$_CODEINTERNE1'    , 'Code dimension interne 1'          , 'VARCHAR (3)'  ),
   ('$$_CODEINTERNE2'    , 'Code dimension interne 2'          , 'VARCHAR (3)'  ),
   ('$$_CODEINTERNE3'    , 'Code dimension interne 3'          , 'VARCHAR (3)'  ),
   ('$$_CODEINTERNE4'    , 'Code dimension interne 4'          , 'VARCHAR (3)'  ),
   ('$$_CODEINTERNE5'    , 'Code dimension interne 5'          , 'VARCHAR (3)'  ),
   ('$$_RECHGPAOLIBDIM1' , 'Codes GPAO 1 dans les libellés'    , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM2' , 'Codes GPAO 2 dans les libellés'    , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM3' , 'Codes GPAO 3 dans les libellés'    , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM4' , 'Codes GPAO 4 dans les libellés'    , 'BOOLEAN'      ),
   ('$$_RECHGPAOLIBDIM5' , 'Codes GPAO 5 dans les libellés'    , 'BOOLEAN'      ),
   ('$$_OPCAISSE'        , 'Opération financière'              , 'VARCHAR (35)' ),
   ('$$_LIBELLELIGNE'    , 'Libellé spécifique ligne'          , 'VARCHAR (70)' ),
   ('$$_FAMILLENIV1'     , 'Code famille niv1 spécif. ligne'   , 'COMBO'        ),
   ('$$_FAMILLENIV2'     , 'Code famille niv2 spécif. ligne'   , 'COMBO'        ),
   ('$$_FAMILLENIV3'     , 'Code famille niv3 spécif. ligne'   , 'COMBO'        ),
   ('$$_COLLECTION'      , 'Code collection spécif. ligne'     , 'COMBO'        ),
   ('$$_FORCEARTICLE'    , 'Article défaut si article inconnu' , 'BOOLEAN'      ),
   ('$$_FORCETIERS'      , 'Tiers défaut si tiers inconnu'     , 'BOOLEAN'      ),
   ('$$_QTESTOCK'        , 'Quantité en valeur absolue'        , 'DOUBLE'       ),
   ('$$_SIGNE'           , 'Signe du mouvement'                , 'VARCHAR(1)'   ),
   ('$$_MAJSTOCK'        , 'Mise à jour des stocks'            , 'BOOLEAN'      ),
   ('$$_PRIXACHUNITAIRE' , 'Prix d''achat unitaire'            , 'DOUBLE'       ),
   ('$$_PRIXREVUNITAIRE' , 'Prix de revient unitaire'          , 'DOUBLE'       ),
   ('$$_PRIXVTEUNITAIRE' , 'Prix unitaire'                     , 'DOUBLE'       ),
   ('$$_PRIXVTEUNIREM'   , 'Prix unitaire remisé'              , 'DOUBLE'       ),
   ('$$_REMISE'          , 'Remise en pourcentage'             , 'DOUBLE'       ),
   ('$$_PRIXACHTOTAL'    , 'Prix d''achat total'               , 'DOUBLE'       ),
   ('$$_PRIXREVTOTAL'    , 'Prix de revient total'             , 'DOUBLE'       ),
   ('$$_PRIXVTETOTAL'    , 'Prix total ligne'                  , 'DOUBLE'       ),
   ('$$_PRIXVTETOTREM'   , 'Prix total ligne remisé'           , 'DOUBLE'       ),
   ('$$_FAMILLETAXE1'    , 'Famille de taxe 1'                 , 'COMBO'        ),
   ('$$_FAMILLETAXE2'    , 'Famille de taxe 2'                 , 'COMBO'        ),
   ('$$_FAMILLETAXE3'    , 'Famille de taxe 3'                 , 'COMBO'        ),
   ('$$_FAMILLETAXE4'    , 'Famille de taxe 4'                 , 'COMBO'        ),
   ('$$_FAMILLETAXE5'    , 'Famille de taxe 5'                 , 'COMBO'        ),
   ('$$_REPRESENTANT'    , 'Représentant'                      , 'VARCHAR (17)' ),
   ('$$_GEREECHEANCES'   , 'Génération des échéances'          , 'BOOLEAN'      ),
   ('$$_FACTUREHT'       , 'Facture HT'                        , 'BOOLEAN'      ),
   ('$$_CODETVA'         , 'Code TVA'                          , 'VARCHAR (3)'  ),
   ('$$_TYPEREMISE'      , 'Motif de démarque'                 , 'COMBO'        ),
   ('$$_MOTIFMVT'        , 'Motif pour E/S de stock'           , 'COMBO'        ),
   ('$$_DEVISE'          , 'Devise'                            , 'COMBO'        ),
   ('$$_DUPLICPIECE'     , 'Nature duplication'                , 'COMBO'        ),
   ('$$_NATUREPIECEG'    , 'Nature de la pièce à solder'       , 'COMBO'        ),
   ('$$_REFINTERNE'      , 'Réf Int de la pièce à solder'      , 'VARCHAR (35)' ),
   ('$$_PIECEUNIQUE'     , 'Référence interne unique'          , 'BOOLEAN'      ),
   ('$$_CTRLPRIXNUL'     , 'Initialisation des PU nuls'        , 'BOOLEAN'      ),
   ('$$_REGROUPELIGNES'  , 'Regroupement des lignes'           , 'BOOLEAN'      ),
   ('$$_MAJDATEINT'      , 'MAJ de la date d''intégration'     , 'BOOLEAN'      ),
   ('$$_CREECLICASH'     , 'Client 999999 CASH 2000'           , 'BOOLEAN'      ));

////////////////////////////////////////////////////////////////////
// Noms des champs permettant de récupérer un règlement FFO
/////////////////////////////////////////////////////////////////////
ChampTobReglement : array[1..6,1..3] of String 	= (
   ('GPE_NATUREPIECEG'   , 'Nature de document'        , 'COMBO'        ),
   ('$$_REFINTERNE'      , 'Référence interne'         , 'VARCHAR (35)' ),
   ('GPE_NUMECHE'        , 'Numéro de ligne'           , 'INTEGER'      ),
   ('GPE_MODEPAIE'       , 'Mode de paiement'          , 'COMBO'        ),
   ('GPE_MONTANTDEV'     , 'Montant en devise'         , 'DOUBLE'       ),
   ('GPE_DEVISE'         , 'Devise de règlement'       , 'COMBO'        ));

////////////////////////////////////////////////////////////////////
// Noms des champs permettant de récupérer un règlement autre que FFO
/////////////////////////////////////////////////////////////////////
ChampTobAcomptes : array[1..2,1..3] of String 	= (
   ('$$_REFINTERNE'      , 'Référence interne'         , 'VARCHAR (35)' ),
   ('$$_DEVISE'          , 'Devise de règlement'       , 'COMBO'       ) );


////////////////////////////////////////////////////////////////////
// Noms des champs permettant de récupérer un inventaire
/////////////////////////////////////////////////////////////////////
ChampTobInventaire : array[1..6,1..3] of String 	= (
   ('$$_LIBELLE'         , 'Libellé'                   , 'VARCHAR (70)' ),
   ('$$_CODEGPAODIM1'    , 'Code dimension 1 dans GPAO', 'VARCHAR (9)'  ),
   ('$$_CODEGPAODIM2'    , 'Code dimension 2 dans GPAO', 'VARCHAR (9)'  ),
   ('$$_CODEGPAODIM3'    , 'Code dimension 3 dans GPAO', 'VARCHAR (9)'  ),
   ('$$_CODEGPAODIM4'    , 'Code dimension 4 dans GPAO', 'VARCHAR (9)'  ),
   ('$$_CODEGPAODIM5'    , 'Code dimension 5 dans GPAO', 'VARCHAR (9)'  ));

////////////////////////////////////////////////////////////////////
// Noms des champs permettant de récupérer la fidélité
/////////////////////////////////////////////////////////////////////
ChampTobFidelite : array[1..5,1..3] of String 	= (
   ('$$_TIERS'           , 'Tiers'                     , 'VARCHAR (17)' ),
   ('$$_NUMCARTEEXT'     , 'Numéro de carte externe'   , 'VARCHAR (17)' ),
   ('$$_DATEOUVERTURE'   , 'Date d''ouverture carte'   , 'DATE'         ),
   ('$$_DATEFERMETURE'   , 'Date de fermeture carte'   , 'DATE'         ),
   ('$$_FERME'           , 'Carte de fidélité fermé'   , 'BOOLEAN'      ));

implementation

{$R *.DFM}


Procedure AppelParamGPAO (Action : TActionFiche; TypeReprise, TypeDonnees : string ; Inventaire : boolean) ;
var X  : TFParamGPAO ;
    PP : THPanel ;
begin
  PP:=FindInsidePanel ;
  X := TFParamGPAO.Create ( Application ) ;

  X.Action      := Action ;
  X.TypeReprise := TypeReprise;
  X.TypeDonnees := TypeDonnees;
  X.Inventaire  := Inventaire;
  X.InitAffichage;

  if Action = TaModif then
  begin
    X.Tob_Entete.PutEcran(X) ;
    X.AfficheParametrage ;
  end;

  if PP=Nil then
   BEGIN
    try
      X.ShowModal ;
    finally
      X.Free ;
    end ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
end;

////////////////////////////////////////////////////////////////////////////////////////
// Suppression d'un code paramétrage
////////////////////////////////////////////////////////////////////////////////////////
Procedure SupprimeParamGPAO (TypeReprise, TypeDonnees : string) ;
var SQL : string ;
begin
  SQL:='Delete From REPRISEGPAO WHERE GRE_TYPEREPRISE="'+TypeReprise+'" AND GRE_TYPEDONNEES="'+TypeDonnees+'"';
  ExecuteSQL(SQL);
  SQL:='Delete From REPRISECHGPAO WHERE GRC_TYPEREPRISE="'+TypeReprise+'" AND GRC_TYPEDONNEES="'+TypeDonnees+'"';
  ExecuteSQL(SQL);
  SQL:='Delete From REPRISECOGPAO WHERE GRO_TYPEREPRISE="'+TypeReprise+'" AND GRO_TYPEDONNEES="'+TypeDonnees+'"';
  ExecuteSQL(SQL);
end;

//////////////////////////////////////////////////////////////////////////////
// Affichage du paramétrage
//////////////////////////////////////////////////////////////////////////////
Procedure TFParamGPAO.AfficheParametrage ;
begin
   TobDeChamps.PutGridDetail (TLChamp, False, False, 'DH_NOMCHAMP;DH_TYPECHAMP;DH_LIBELLE', True);
   Tob_Reprise.PutGridDetail (GChamp , False, False, 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR', True);
end;

//////////////////////////////////////////////////////////////////////////////
// Modification d'un paramétrage
//   1 - Chargement des TOB
//   2 - Affichage des TOB
///////////////////////////////////////////////////////////////////////////////
Procedure TFParamGPAO.ChargeParametrage ;
var Q               : TQuery ;
    RechTob         : TOB    ;
    TobChamp        : TOB    ;
    ChampTobDoc     : TOB    ;
    TobFilleReprise : TOB    ;
    cpt             : integer;
begin
  //
  // Chargement de la TOB "entête"
  //
  Q := OpenSQL ('SELECT * FROM REPRISEGPAO WHERE GRE_TYPEREPRISE="'+TypeReprise+'" AND GRE_TYPEDONNEES="'+TypeDonnees+'"' ,True) ;
  if Q.EOF then
  begin
    PGIBox ('Problème pour récupérer le paramétrage ! ', '');
    Ferme (Q);
    exit ;
  end
  else begin
    Tob_Entete.PutValue ('GRE_TYPEREPRISE' , Q.FindField('GRE_TYPEREPRISE').AsString);
    Tob_Entete.PutValue ('GRE_TYPEDONNEES' , Q.FindField('GRE_TYPEDONNEES').AsString);
    Tob_Entete.PutValue ('GRE_NOMTABLE'    , Q.FindField('GRE_NOMTABLE').AsString);
    Tob_Entete.PutValue ('GRE_DEBUTENREG'  , Q.FindField('GRE_DEBUTENREG').AsString);
    Tob_Entete.PutValue ('GRE_LONGUEURFIXE', Q.FindField('GRE_LONGUEURFIXE').AsString);
    Tob_Entete.PutValue ('GRE_SEPARENREG'  , Q.FindField('GRE_SEPARENREG').AsString);
    // Nouveau champ : entêtre de colonne
    Tob_Entete.PutValue ('GRE_ENTETELIG'   , Q.FindField('GRE_ENTETELIG').AsString);
    if Tob_Entete.GetValue ('GRE_ENTETELIG') <> '-' then Tob_Entete.PutValue ('GRE_ENTETELIG' , 'X')
    else Tob_Entete.PutValue ('GRE_ENTETELIG' , '-');
    Ferme (Q);
  end;
  //
  // Chargement de la TOB "détail" de la reprise
  //
  Q := OpenSQL ('SELECT * FROM REPRISECHGPAO WHERE GRC_TYPEREPRISE="'+TypeReprise+'" AND GRC_TYPEDONNEES="'+TypeDonnees+'" ORDER BY GRC_PRESENTATION' ,True) ;
  Q.First ;
  While not Q.Eof do
  begin
    TobFilleReprise := TOB.Create ('REPRISECHGPAO', Tob_Reprise, -1) ;
    TobFilleReprise.PutValue ('GRC_TYPEREPRISE'   , Q.FindField ('GRC_TYPEREPRISE').AsString);
    TobFilleReprise.PutValue ('GRC_TYPEDONNEES'   , Q.FindField ('GRC_TYPEDONNEES').AsString);
    TobFilleReprise.PutValue ('GRC_NOMCHAMP'      , Q.FindField ('GRC_NOMCHAMP').AsString);
    TobFilleReprise.PutValue ('GRC_TYPECHAMP'     , Q.FindField ('GRC_TYPECHAMP').AsString);
    TobFilleReprise.PutValue ('GRC_LIBELLECHAMP'  , Q.FindField ('GRC_LIBELLECHAMP').AsString);
    TobFilleReprise.PutValue ('GRC_POSDEBUT'      , Q.FindField ('GRC_POSDEBUT').AsInteger);
    TobFilleReprise.PutValue ('GRC_LONGUEUR'      , Q.FindField ('GRC_LONGUEUR').AsInteger);
    TobFilleReprise.PutValue ('GRC_VALDEFAUT'     , Q.FindField ('GRC_VALDEFAUT').AsString);
    TobFilleReprise.PutValue ('GRC_FORMULE'       , Q.FindField ('GRC_FORMULE').AsString);
    TobFilleReprise.PutValue ('GRC_GERETABLECORR' , Q.FindField ('GRC_GERETABLECORR').AsString);
    Q.Next ;
  end;
  Ferme (Q);

  //
  // Chargement de la TOB des champs pouvant être repris
  // Attention : il faut épurer les champs déjà présent dans la TOB "détail"
  //     1 - Récupération du préfixe de la table
  //     2 - Balayage des champs et scrutation de la TOB détail
  //
  if (Tob_Entete.GetValue ('GRE_NOMTABLE') <> 'PIECE') and (Tob_Entete.GetValue ('GRE_NOMTABLE') <> 'PIEDECHE') then
  begin
    Q := OpenSQL ('SELECT DT_PREFIXE FROM DETABLES WHERE DT_NOMTABLE="'+Tob_Entete.GetValue('GRE_NOMTABLE')+'"',True) ;
    if Q.EOF then
    begin
      PGIBox ('Problème pour récupérer le préfixe de la table ! ', '');
      Ferme (Q);
      exit ;
    end
    else begin
      Prefixe := Q.FindField('DT_PREFIXE').AsString;
      Ferme (Q);
    end;

    if TobDeChamps <> nil then TobDeChamps.free;
    TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);

    Q := OpenSQL ('SELECT DH_NOMCHAMP, DH_LIBELLE, DH_TYPECHAMP FROM DECHAMPS WHERE DH_PREFIXE="'+Prefixe+'" ORDER BY DH_NOMCHAMP' ,True) ;
    Q.First ;
    While not Q.Eof do
    begin
      RechTob:=Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, Q.FindField('DH_NOMCHAMP').AsString],TRUE) ;
      if RechTob = nil then
      begin
        TobChamp := Tob.Create ('DECHAMPS', TobDeChamps, -1);
        TobChamp.PutValue ('DH_NOMCHAMP' , Q.FindField ('DH_NOMCHAMP').AsString);
        TobChamp.PutValue ('DH_TYPECHAMP', Q.FindField ('DH_TYPECHAMP').AsString);
        TobChamp.PutValue ('DH_LIBELLE'  , Q.FindField ('DH_LIBELLE').AsString);
      end;
      Q.Next ;
    end;
    Ferme (Q);
  end else
  begin
    //
    // Cas spécifique des pièces
    //
    if Tob_Entete.GetValue ('GRE_NOMTABLE') = 'PIECE' then
    begin
      if TobDeChamps <> nil then TobDeChamps.free;
      TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);

      for cpt:=Low(ChampTobDocument) to High(ChampTobDocument) do
      begin
        if (ChampTobDocument[cpt, 1]<>'') then
        begin
          RechTob:=Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobDocument[cpt, 1]],TRUE) ;
          if RechTob = nil then
          begin
            ChampTobDoc := Tob.Create ('DECHAMPS', TobDeChamps, -1);
            if ChampTobDoc.FieldExists ('DH_NOMCHAMP') then ChampTobDoc.AddChampSup ('DH_NOMCHAMP', True);
            ChampTobDoc.PutValue ('DH_NOMCHAMP', ChampTobDocument[cpt, 1]);
            if ChampTobDoc.FieldExists ('DH_LIBELLE') then ChampTobDoc.AddChampSup ('DH_LIBELLE', True);
            ChampTobDoc.PutValue ('DH_LIBELLE', ChampTobDocument[cpt, 2]);
            if ChampTobDoc.FieldExists ('DH_TYPECHAMP') then ChampTobDoc.AddChampSup ('DH_TYPECHAMP', True);
            ChampTobDoc.PutValue ('DH_TYPECHAMP', ChampTobDocument[cpt, 3]);
          end ;
        end;
      end;
    end;
    //
    // Cas spécifique des pièces
    //
    if Tob_Entete.GetValue ('GRE_NOMTABLE') = 'PIEDECHE' then
    begin
      if TobDeChamps <> nil then TobDeChamps.free;
      TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);

      for cpt:=Low(ChampTobReglement) to High(ChampTobReglement) do
      begin
        if (ChampTobReglement[cpt, 1]<>'') then
        begin
          RechTob:=Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobReglement[cpt, 1]],TRUE) ;
          if RechTob = nil then
          begin
            ChampTobDoc := Tob.Create ('DECHAMPS', TobDeChamps, -1);
            if ChampTobDoc.FieldExists ('DH_NOMCHAMP') then ChampTobDoc.AddChampSup ('DH_NOMCHAMP', True);
            ChampTobDoc.PutValue ('DH_NOMCHAMP', ChampTobReglement[cpt, 1]);
            if ChampTobDoc.FieldExists ('DH_LIBELLE') then ChampTobDoc.AddChampSup ('DH_LIBELLE', True);
            ChampTobDoc.PutValue ('DH_LIBELLE', ChampTobReglement[cpt, 2]);
            if ChampTobDoc.FieldExists ('DH_TYPECHAMP') then ChampTobDoc.AddChampSup ('DH_TYPECHAMP', True);
            ChampTobDoc.PutValue ('DH_TYPECHAMP', ChampTobReglement[cpt, 3]);
          end ;
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques
  //
  AjoutChampSpecifiqueTable (Tob_Entete.GetValue('GRE_NOMTABLE'), Tob_Reprise);
end;

///////////////////////////////////////////////////////////////////////////////
// Initialisation de l'affichage au lancement de la fenêtre
///////////////////////////////////////////////////////////////////////////////
Procedure TFParamGPAO.InitAffichage ;
begin
  //
  // Création de la TOB de reprise "entête"
  //
  Tob_Entete := TOB.CREATE ('REPRISEGPAO', NIL, -1);
  //
  // Création de la TOB de reprise "détail"
  //
  Tob_Reprise := TOB.CREATE ('Le détail', NIL, -1);

  GChamp.ColLengths  [COL_Champ]     := -1;
  GChamp.ColLengths  [COL_PosDebut]  :=  4;
  GChamp.ColTypes    [COL_PosDebut]  := 'R';
  GChamp.ColFormats  [COL_PosDebut]  := '###0';
  GChamp.ColAligns   [COL_PosDebut]  := taRightJustify;
  GChamp.ColLengths  [COL_Longueur]  :=  4;
  GChamp.ColTypes    [COL_Longueur]  := 'R';
  GChamp.ColFormats  [COL_Longueur]  := '###0';
  GChamp.ColAligns   [COL_Longueur]  := taRightJustify;
  GChamp.ColLengths  [COL_TableCorr] :=  1;
  GChamp.ColTypes    [COL_TableCorr] := 'B';
  GChamp.ColFormats  [COL_TableCorr] := IntToStr(Integer(csCoche));
  GChamp.ColAligns   [COL_TableCorr] := taCenter;
  GChamp.ColEditables[COL_TableCorr] := False;
  GChamp.FlipBool := True;

  if Action = TaModif then
  begin
    ChargeParametrage;
    GRE_TYPEREPRISE.Enabled := False;
    GRE_TYPEDONNEES.Enabled := False;
    GRE_NOMTABLE.Enabled    := False;
    if not Inventaire then ButtonDupliq.Enabled := True
    else ButtonDupliq.Enabled := False ;
    GRE_ENTETELIG.Enabled   := True ;
    if GRE_ENTETELIG.checked then
    begin
      GRE_DEBUTENREG.Enabled := True ;
    end else
    begin
     GRE_DEBUTENREG.Enabled := False ;
    end;
    //ChargeParametrage;
  end
  else begin
    GRE_TYPEREPRISE.Enabled := True ;
    GRE_TYPEDONNEES.Enabled := True ;
    GRE_NOMTABLE.Enabled    := True ;
    ButtonDupliq.Enabled    := False;
    GRE_ENTETELIG.Enabled   := True ;
    GRE_ENTETELIG.checked   := True ;
  end;
end;

//////////////////////////////////////////////////////////////////////////////
// Contrôle des données avant sauvegarde
//////////////////////////////////////////////////////////////////////////////
function TFParamGPAO.ControleDonnees : boolean ;
var QMax    : TQuery ;
    Code    : string ;
    Combien : integer ;
begin
  result := False ;
  if Action = TaCreat then
  begin
    if ExisteSQL ('Select GRE_NOMTABLE from REPRISEGPAO where GRE_TYPEREPRISE="'+GRE_TYPEREPRISE.Value+'" AND GRE_TYPEDONNEES="'+GRE_TYPEDONNEES.Value+'"') then
    begin
      PGIBox ('Cette reprise existe déjà !', '');
      GRE_TYPEDONNEES.Value := '';
      GRE_TYPEDONNEES.SetFocus;
      exit ;
    end;
  end;
  //
  // Type de reprise
  //
  Code := GRE_TYPEREPRISE.Value ;
  if Code = '' then
  begin
    PGIBox('Vous devez renseigner le type de reprise !','');
    exit ;
  end;
  if not ExisteSQL ('Select CC_LIBELLE from CHOIXCOD where CC_TYPE = "GT1" AND CC_CODE = "'+ Code +'"') then
  begin
    PGIBox ('Le type de reprise est incorrect ! ', '');
    GRE_TYPEREPRISE.SetFocus;
    exit;
  end;
  //
  // Type de donnée
  //
  Code := GRE_TYPEDONNEES.Value ;
  if Code = '' then
  begin
    PGIBox ('Vous devez renseigner le type de donnée ! ', '');
    exit ;
  end;
  if not ExisteSQL ('Select CC_LIBELLE from CHOIXCOD where CC_TYPE = "GT2" AND CC_CODE = "'+ Code +'"') then
  begin
    PGIBox ('Le type de donnée est incorrect ! ', '');
    GRE_TYPEDONNEES.SetFocus;
    exit;
  end;
  //
  // Nom de la table
  //
  Code := GRE_NOMTABLE.Text ;
  if Code = '' then
  begin
    PGIBox ('Vous devez renseigner le nom de la table ! ', '');
    exit ;
  end;
  if not ExisteSQL ('Select DT_LIBELLE from DETABLES where DT_NOMTABLE = "'+ Code +'"') then
  begin
    PGIBox ('Le nom de la table est incoorect ! ', '');
    GRE_NOMTABLE.SetFocus;
    exit;
  end;
  //
  // Entête de colonne
  //
  if not GRE_ENTETELIG.Checked then
  begin
    QMax:=OpenSQL('Select count(*) from REPRISEGPAO where GRE_TYPEREPRISE="'+GRE_TYPEREPRISE.Value+'"', True) ;
    Combien := QMax.Fields[0].AsInteger ;
    Ferme(QMax) ;
    if ((Action = TaCreat) and (Combien >=1)) or ((Action = TaModif) and (Combien>1))  then
    begin
      PGIBox ('Sans entête de ligne, vous ne pouvez définir qu''une seule reprise par provenance de données !', '');
      exit ;
    end;
  end else
  begin
    if GRE_DEBUTENREG.Text = '' then
    begin
      PGIBox ('Vous devez renseigner l''identifiant ! ', '');
      GRE_DEBUTENREG.SetFocus;
      exit;
    end;
  end;
  result := True ;
end;


//////////////////////////////////////////////////////////////////////////////
// Supprime l'ancien paramétrage avant MAJ du détail
// Objectif : Faire un annule et rezmplace du paramétrage
///////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.SupprimeAncienParametrage ;
var SQL : string ;
begin
  SQL:='Delete From REPRISECHGPAO WHERE GRC_TYPEREPRISE="'+TypeReprise+'" AND GRC_TYPEDONNEES="'+TypeDonnees+'"';
  ExecuteSQL(SQL);
end;

//////////////////////////////////////////////////////////////////////////////
// Sauvegarde DB du paramétrage
///////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.SauvegardeParametrage ;
var TobFille : TOB     ;
    cpt      : integer ;
begin
  //
  // Sauvegarde de l'enregistrement entête
  //
  if Action = TaCreat then Tob_Entete.InsertDB (nil, True)
  else Tob_Entete.UpdateDB (True);
  //
  // Vérifie la clef de la TOB "détail"
  //
  for cpt:=0 to Tob_Reprise.detail.count-1 do
  begin
    TobFille := Tob_Reprise.detail [cpt] ;
    if TobFille.GetValue ('GRC_TYPEREPRISE') <> GRE_TYPEREPRISE.Value then
    begin
      TobFille.PutValue ('GRC_TYPEREPRISE', GRE_TYPEREPRISE.Value);
    end;
    if TobFille.GetValue ('GRC_TYPEDONNEES') <> GRE_TYPEDONNEES.Value then
    begin
      TobFille.PutValue ('GRC_TYPEDONNEES', GRE_TYPEDONNEES.Value);
    end;
    TobFille.PutValue ('GRC_PRESENTATION', cpt+1);
  end;
  //
  // Sauvegarde des éléments "détails"
  //
  if Action = TaModif then SupprimeAncienParametrage;

  Tob_Reprise.InsertDB (nil, True);
end;

//////////////////////////////////////////////////////////////////////////////
// Sauvegarde du paramétrage saisi
///////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.BValiderClick(Sender: TObject);
begin
  //
  // Mise à jour de la Tob "entête" avec les zones de saisie
  //
  if ControleDonnees then
  begin
    Tob_Entete.GetEcran(Self) ;
    //
    // Récupération de la TOB des champs à reprendre
    //
    if Tob_Reprise.Detail.count > 0 then Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
    //
    // Sauvegarde de la saisie
    //
    SauvegardeParametrage ;

    Action := taModif ;
    GRE_TYPEREPRISE.Enabled := False;
    GRE_TYPEDONNEES.Enabled := False;
    GRE_NOMTABLE.Enabled    := False;
    ButtonDupliq.Enabled    := True;
  end;
end;

//////////////////////////////////////////////////////////////////////////////
// Ajout des champs spécifique en fonction des tables
//////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.AjoutChampSpecifiqueTable (NomTable : string; TTOB_REPRISE : TOB);
var TobFille : TOB ;
    RechTob  : TOB ;
    cpt      : integer;
begin
  //
  // Ajout des champs spécifiques pour la table ARTICLE
  //
  if (NomTable = 'ARTICLE') then
  begin
    for cpt:=Low(ChampTobArticle) to High(ChampTobArticle) do
    begin
      if (ChampTobArticle[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobArticle[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobArticle[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobArticle[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobArticle[cpt, 3]);
        end;
      end;
    end;
    //
    // Spécificité n°ARTICLE : Les types des codes dimensions sont passés de COMBO en VARCHAR(3)
    //                         En effet, ces codes sont les codes internes et dépendent de la grille de dimension,
    //                         il est donc difficile (voir impossible) d'afficher la tablette correspondante
    //
    for cpt := 0 to TobDeChamps.detail.count-1 do
    begin
      TobFille := TobDeChamps.detail[cpt] ;
      if (TobFille.GetValue('DH_NOMCHAMP')='GA_CODEDIM1') or (TobFille.GetValue('DH_NOMCHAMP')='GA_CODEDIM2') or (TobFille.GetValue('DH_NOMCHAMP')='GA_CODEDIM3') or
         (TobFille.GetValue('DH_NOMCHAMP')='GA_CODEDIM4') or (TobFille.GetValue('DH_NOMCHAMP')='GA_CODEDIM5') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques DISPO
  //
  if NomTable = 'DISPO' then
  begin
    for cpt:=Low(ChampTobDispo) to High(ChampTobDispo) do
    begin
      if (ChampTobDispo[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobDispo[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobDispo[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobDispo[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobDispo[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques TARIF
  //
  if NomTable = 'TARIF' then
  begin
    for cpt:=Low(ChampTobTarif) to High(ChampTobTarif) do
    begin
      if (ChampTobTarif[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobTarif[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobTarif[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobTarif[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobTarif[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques TIERS
  //
  if NomTable = 'TIERS' then
  begin
    for cpt:=Low(ChampTobTiers) to High(ChampTobTiers) do
    begin
      if (ChampTobTiers[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobTiers[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobTiers[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobTiers[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobTiers[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques ACOMPTES
  //
  if NomTable = 'ACOMPTES' then
  begin
    for cpt:=Low(ChampTobAcomptes) to High(ChampTobAcomptes) do
    begin
      if (ChampTobAcomptes[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobAcomptes[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobAcomptes[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobAcomptes[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobAcomptes[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques TRANSINVLIG
  //
  if NomTable = 'TRANSINVLIG' then
  begin
    for cpt:=Low(ChampTobInventaire) to High(ChampTobInventaire) do
    begin
      if (ChampTobInventaire[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobInventaire[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobInventaire[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobInventaire[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobInventaire[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques TRANSINVLIG
  //
  if NomTable = 'TRANSINVLIG' then
  begin
    for cpt:=Low(ChampTobInventaire) to High(ChampTobInventaire) do
    begin
      if (ChampTobInventaire[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobInventaire[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobInventaire[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobInventaire[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobInventaire[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Ajout des champs spécifiques FIDELITELIG
  //
  if NomTable = 'FIDELITELIG' then
  begin
    for cpt:=Low(ChampTobFidelite) to High(ChampTobFidelite) do
    begin
      if (ChampTobFidelite[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobFidelite[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobFidelite[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobFidelite[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobFidelite[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Spécificité table CHOIXCOD
  //
  if (NomTable = 'CHOIXCOD') then
  begin
    for cpt := 0 to TobDeChamps.detail.count-1 do
    begin
      TobFille := TobDeChamps.detail[cpt] ;
      if (TobFille.GetValue('DH_NOMCHAMP')='CC_CODE') or (TobFille.GetValue('DH_NOMCHAMP')='CC_TYPE') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
    end;
  end;
  //
  // Spécificité table CHOIXEXT
  //
  if (NomTable = 'CHOIXEXT') then
  begin
    for cpt := 0 to TobDeChamps.detail.count-1 do
    begin
      TobFille := TobDeChamps.detail[cpt] ;
      if (TobFille.GetValue('DH_NOMCHAMP')='YX_TYPE') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
    end;
  end;
  //
  // Spécificité table DIMENSION
  //
  if (NomTable = 'DIMENSION') then
  begin
    for cpt := 0 to TobDeChamps.detail.count-1 do
    begin
      TobFille := TobDeChamps.detail[cpt] ;
      if (TobFille.GetValue('DH_NOMCHAMP')='GDI_GRILLEDIM') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
      if (TobFille.GetValue('DH_NOMCHAMP')='GDI_TYPEDIM') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
    end;

    for cpt:=Low(ChampTobDimension) to High(ChampTobDimension) do
    begin
      if (ChampTobDimension[cpt, 1]<>'') then
      begin
        if TTOB_REPRISE = nil then RechTob := nil
        else RechTob := Tob_Reprise.FindFirst(['GRC_TYPEREPRISE', 'GRC_TYPEDONNEES', 'GRC_NOMCHAMP'],[TypeReprise, TypeDonnees, ChampTobDimension[cpt, 1]],TRUE) ;

        if RechTob = nil then
        begin
          TobFille := TOB.Create ('DECHAMPS', TobDeChamps, -1);
          TobFille.PutValue ('DH_NOMCHAMP'  , ChampTobDimension[cpt, 1]);
          TobFille.PutValue ('DH_LIBELLE'   , ChampTobDimension[cpt, 2]);
          TobFille.PutValue ('DH_TYPECHAMP' , ChampTobDimension[cpt, 3]);
        end;
      end;
    end;
  end;
  //
  // Spécificité table TARIFPER
  //
  if (NomTable = 'TARIFPER') then
  begin
    for cpt := 0 to TobDeChamps.detail.count-1 do
    begin
      TobFille := TobDeChamps.detail[cpt] ;
      if (TobFille.GetValue('DH_NOMCHAMP')='GFP_ETABLISSEMENT') then
      begin
        TobFille.PutValue ('DH_TYPECHAMP', 'VARCHAR(3)');
      end;
    end;
  end;
end;

//////////////////////////////////////////////////////////////////////////////
// Charge la liste des champs pouvant être repris
//////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.InitialiseListeChamps;
var Q           : TQuery  ;
    Prefixe     : string  ;
    ChampTOBDoc : TOB     ;
    cpt         : integer ;
begin
  //
  // Vidage de la liste pécédentes des champs
  //
  TLChamp.VidePile (False);
  //
  // Vidage de la liste des champs à reprendre
  //
  GChamp.VidePile (False);
  if Tob_Reprise <> nil then Tob_Reprise.free;
  Tob_Reprise := TOB.CREATE ('Le détail', NIL, -1);
  //
  // Récupération du préfixe de la table sélectionnée
  //
  if (GRE_NOMTABLE.text <> 'PIECE') and (GRE_NOMTABLE.text <> 'PIEDECHE') then
  begin
    Q := OpenSQL ('SELECT DT_PREFIXE FROM DETABLES WHERE DT_NOMTABLE="'+GRE_NOMTABLE.Text+'"',True) ;
    if Q.EOF then
    begin
      PGIBox ('Problème pour récupérer le préfixe de la table ! ', '');
      Ferme (Q);
      exit ;
    end else Prefixe := Q.FindField('DT_PREFIXE').AsString;
    Ferme (Q);
    //
    // Chargement du TListView avec les champs de la table sélectionnée
    //
    if TobDeChamps <> nil then TobDeChamps.free;
    Q := OpenSQL ('SELECT DH_NOMCHAMP, DH_LIBELLE, DH_TYPECHAMP FROM DECHAMPS WHERE DH_PREFIXE="'+Prefixe+'" ORDER BY DH_NOMCHAMP',True) ;
    TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);
    TobDeChamps.LoadDetailDB ('DECHAMPS', '', '', Q, False);
    Ferme (Q);
  end else
  begin
    //
    // Chargement spécifique des champs pour la reprise de documents
    //
    if GRE_NOMTABLE.text = 'PIECE' then
    begin
      TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);
      for cpt:=Low(ChampTobDocument) to High(ChampTobDocument) do
      begin
        if (ChampTobDocument[cpt, 1]<>'') then
        begin
          ChampTobDoc := Tob.Create ('DECHAMPS', TobDeChamps, -1);
          ChampTobDoc.PutValue ('DH_NOMCHAMP' , ChampTobDocument[cpt, 1]);
          ChampTobDoc.PutValue ('DH_LIBELLE'  , ChampTobDocument[cpt, 2]);
          ChampTobDoc.PutValue ('DH_TYPECHAMP', ChampTobDocument[cpt, 3]);
        end ;
      end;
    end;
    //
    // Chargement spécifique des champs pour la reprise des règlements
    //
    if GRE_NOMTABLE.text = 'PIEDECHE' then
    begin
      TobDeChamps := Tob.Create ('Liste des champs', Nil, -1);
      for cpt:=Low(ChampTobReglement) to High(ChampTobReglement) do
      begin
        if (ChampTobDocument[cpt, 1]<>'') then
        begin
          ChampTobDoc := Tob.Create ('DECHAMPS', TobDeChamps, -1);
          ChampTobDoc.PutValue ('DH_NOMCHAMP' , ChampTobReglement[cpt, 1]);
          ChampTobDoc.PutValue ('DH_LIBELLE'  , ChampTobReglement[cpt, 2]);
          ChampTobDoc.PutValue ('DH_TYPECHAMP', ChampTobReglement[cpt, 3]);
        end ;
      end;
    end;
  end ;
  //
  // Ajout des champs spécifiques en fonction de la table sélectionnée
  //
  AjoutChampSpecifiqueTable (GRE_NOMTABLE.text, nil);
  //
  // Affichage
  //
  TobDeChamps.PutGridDetail (TLChamp, False, False, 'DH_NOMCHAMP;DH_TYPECHAMP;DH_LIBELLE', True);
end;

////////////////////////////////////////////////////////////////////////////////
// La table est sélectionnée, on affiche les champs pouvant être repris
////////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.GRE_NOMTABLEExit(Sender: TObject);
begin
  GRE_NOMTABLE.Text := uppercase (GRE_NOMTABLE.Text);
  LibelleNomTable.visible := True;
  InitialiseListeChamps;
end;

procedure TFParamGPAO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  //
  // Libération des TOB
  //
  if TobDeChamps <> nil then TobDeChamps.free;
  if Tob_Entete  <> nil then Tob_Entete.free;
  if Tob_Reprise <> nil then Tob_Reprise.free;
end;

procedure TFParamGPAO.BFlecheDroiteClick(Sender: TObject);
var TobChampSelect  : TOB     ;
    TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
begin
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if TLChamp.Row < 0 then exit;
  if TobDeChamps.detail.count <= 0 then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  if Tob_Reprise.detail.count > 0 then Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow := TLChamp.Row;
  TobChampSelect := TobDeChamps.Detail [IndiceChampRow-1];
  //
  // Création d'une ligne dans la TOB de reprise
  //
  TobFilleReprise := TOB.Create ('REPRISECHGPAO', Tob_Reprise, -1) ;
  TobFilleReprise.PutValue ('GRC_TYPEREPRISE'   , GRE_TYPEREPRISE.Value);
  TobFilleReprise.PutValue ('GRC_TYPEDONNEES'   , GRE_TYPEDONNEES.Value);
  TobFilleReprise.PutValue ('GRC_NOMCHAMP'      , TobChampSelect.GetValue ('DH_NOMCHAMP'));
  TobFilleReprise.PutValue ('GRC_TYPECHAMP'     , TobChampSelect.GetValue ('DH_TYPECHAMP'));
  TobFilleReprise.PutValue ('GRC_LIBELLECHAMP'  , TobChampSelect.GetValue ('DH_LIBELLE'));

  //
  // Affichage de la TOB de reprise
  //
  Tob_Reprise.PutGridDetail (GChamp, False, False, 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR', True);
  GChamp.Row := Tob_Reprise.detail.count;
  //
  // Suppression de la ligne sélectionnée dans la liste des champs
  //
  TobDeChamps.Detail [IndiceChampRow-1].free ;
  //
  // Réaffichage de la liste des champs pouvant être repris
  //
  TobDeChamps.PutGridDetail (TLChamp, False, False, 'DH_NOMCHAMP;DH_TYPECHAMP;DH_LIBELLE', True);
  if IndiceChampRow > 1 then
  begin
    if IndiceChampRow > TobDeChamps.detail.count then  TLChamp.Row := TobDeChamps.detail.count
    else TLChamp.Row := IndiceChampRow;
  end;

end;

procedure TFParamGPAO.BFlecheGaucheClick(Sender: TObject);
var TobChamp        : TOB     ;
    TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
begin
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if GChamp.Row < 0 then exit;
  if Tob_Reprise.detail.count <= 0 then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow  := GChamp.Row;
  TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow-1];
  //
  // Ajout d'une ligne dans la TOB des champs pouvant être repris
  //
  TobChamp := TOB.Create ('DECHAMPS', TobDeChamps, -1) ;
  TobChamp.PutValue ('DH_NOMCHAMP' , TobFilleReprise.GetValue ('GRC_NOMCHAMP'));
  TobChamp.PutValue ('DH_TYPECHAMP', TobFilleReprise.GetValue ('GRC_TYPECHAMP'));
  TobChamp.PutValue ('DH_LIBELLE'  , TobFilleReprise.GetValue ('GRC_LIBELLECHAMP'));
  //
  // Affichage de la TOB des champs pouvant être repris
  //
  TobDeChamps.PutGridDetail (TLChamp, False, False, 'DH_NOMCHAMP;DH_TYPECHAMP;DH_LIBELLE', True);
  TLChamp.Row := TobDeChamps.detail.count;
  //
  // Suppression de la ligne sélectionnée dans la liste des champs
  //
  Tob_Reprise.Detail [IndiceChampRow-1].free ;
  //
  // Réaffichage de la liste des champs pouvant être repris
  //
  Tob_Reprise.PutGridDetail (GChamp, False, False, 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR', True);
  if IndiceChampRow > 1 then
  begin
    if IndiceChampRow > Tob_Reprise.detail.count then  GChamp.Row := Tob_Reprise.detail.count
    else GChamp.Row := IndiceChampRow;
  end;
end;

procedure TFParamGPAO.GRE_TYPEDONNEESExit(Sender: TObject);
begin
  if Action = TaCreat then
  begin
    if ExisteSQL ('Select GRE_NOMTABLE from REPRISEGPAO where GRE_TYPEREPRISE="'+GRE_TYPEREPRISE.Value+'" AND GRE_TYPEDONNEES="'+GRE_TYPEDONNEES.Value+'"') then
    begin
      PGIBox ('Cette reprise existe déjà !', '');
      GRE_TYPEDONNEES.Value := '';
      GRE_TYPEDONNEES.SetFocus;
    end;
  end;
end;


procedure TFParamGPAO.TButtonDefautClick(Sender: TObject);
begin
  if Tob_Reprise.detail.count > 0 then
  begin
    Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
    AppelSaisieValeurGPAO (Tob_Reprise);
  end;
end;

procedure TFParamGPAO.BFermeClick(Sender: TObject);
begin
  Self.Close;
end;

procedure TFParamGPAO.ButtonDupliqClick(Sender: TObject);
begin
  if Action = TaModif then
  begin
    if (GRE_TYPEREPRISE.Value <> '') AND (GRE_TYPEDONNEES.Value <> '') then
    begin
      Action                  := TaCreat ;
      GRE_TYPEREPRISE.Enabled := True;
      GRE_TYPEDONNEES.Enabled := True;
      GRE_NOMTABLE.Enabled    := True;
      GRE_TYPEREPRISE.Value   := '' ;
      GRE_TYPEDONNEES.Value   := '' ;
      ButtonDupliq.Enabled    := False;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
// Déplace un élement repris dans la liste (vers le haut)
///////////////////////////////////////////////////////////////////////////////
procedure TFParamGPAO.BFlecheHautClick(Sender: TObject);
var TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
begin
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if GChamp.Row < 0 then exit;
  if Tob_Reprise.detail.count <= 0 then exit ;
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow  := GChamp.Row;
  if IndiceChampRow <= 1 then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // Sauvegarde de la TOB (=la ligne) à déplacer
  TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow-2];
  //
  // Inversemment des lignes
  //
  Tob_Reprise.Detail [IndiceChampRow-2] := Tob_Reprise.Detail [IndiceChampRow-1] ;
  Tob_Reprise.Detail [IndiceChampRow-1] := TobFilleReprise ;
  //
  // Réaffichage de la liste des champs pouvant être repris
  //
  Tob_Reprise.PutGridDetail (GChamp, False, False, 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR', True);
  //
  // Placement du focus
  //
  GChamp.Row := IndiceChampRow-1;
end;

procedure TFParamGPAO.BFlecheBasClick(Sender: TObject);
var TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
begin
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if GChamp.Row < 0 then exit;
  if Tob_Reprise.detail.count <= 0 then exit ;
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow  := GChamp.Row;
  if IndiceChampRow >= Tob_Reprise.detail.count then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // Sauvegarde de la TOB (=la ligne) à déplacer
  TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow];
  //
  // Inversemment des lignes
  //
  Tob_Reprise.Detail [IndiceChampRow]   := Tob_Reprise.Detail [IndiceChampRow-1] ;
  Tob_Reprise.Detail [IndiceChampRow-1] := TobFilleReprise ;
  //
  // Réaffichage de la liste des champs pouvant être repris
  //
  Tob_Reprise.PutGridDetail (GChamp, False, False, 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR', True);
  //
  // Placement du focus
  //
  GChamp.Row := IndiceChampRow+1;
end;

procedure TFParamGPAO.BFormuleClick(Sender: TObject);
var TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
    Formule         : string ;
begin
  if Action = taConsult then exit ;
  if not BFormule.Enabled then exit ;
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if GChamp.Row < 0 then exit;
  if Tob_Reprise.detail.count <= 0 then exit ;
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow  := GChamp.Row -1;
  if (IndiceChampRow < 0) or (IndiceChampRow > Tob_Reprise.detail.count) then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // Lecture de la formule associée au champ
  //
  TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow];
  Formule := TobFilleReprise.GetValue('GRC_FORMULE') ;
  //
  // Saisie de la formule
  //
  Formule := EditeFormule(Nil, Formule);
  //
  // Mise en forme et enregistrement de la formule (en TOB)
  //
  Formule := Uppercase(Trim(Formule));
  if (Formule = ';') or (Formule = '[SI(;;)]') then Formule := '';
  TobFilleReprise.PutValue('GRC_FORMULE', Formule) ;
end;

procedure TFParamGPAO.BTableCorrClick(Sender: TObject);
var TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
    NomChamp        : string ;
begin
  if Action = taConsult then exit ;
  if not BTableCorr.Enabled then exit ;
  //
  // Y a t il quelque chose de sélectionné ?
  //
  if GChamp.Row < 0 then exit;
  if Tob_Reprise.detail.count <= 0 then exit ;
  //
  // récupération de la ligne sélectionnée
  //
  IndiceChampRow  := GChamp.Row -1;
  if (IndiceChampRow < 0) or (IndiceChampRow > Tob_Reprise.detail.count) then exit ;
  //
  // récupération des éventuelles valeurs saisies ou modifié
  //
  Tob_Reprise.GetGridDetail (GChamp, Tob_Reprise.Detail.count, 'REPRISECHGPAO', 'GRC_NOMCHAMP;GRC_POSDEBUT;GRC_LONGUEUR;GRC_GERETABLECORR');
  //
  // saisie de la table de correspondance
  //
  TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow];
  NomChamp := TobFilleReprise.GetValue('GRC_NOMCHAMP') ;
  AppelSaisieCorrespGPAO(Action, TypeReprise, NomChamp);
end;

procedure TFParamGPAO.BImprimeClick(Sender: TObject);
begin
  PrintDBGrid(GChamp, Panel1, Caption, '');
end;

procedure TFParamGPAO.BChercherClick(Sender: TObject);
begin
  if TLChamp.RowCount <= 2 then Exit ;
  FindFirst := True ;
  FindLigne.Execute ;
end;

procedure TFParamGPAO.FindLigneFind(Sender: TObject);
begin
  if (ActiveControl = GChamp) and (GChamp.RowCount > GChamp.FixedRows) then
  begin
    Rechercher(GChamp, FindLigne, FindFirst) ;
  end else
  begin
    Rechercher(TLChamp, FindLigne, FindFirst) ;
  end;
end;

procedure TFParamGPAO.GChampRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var TobFilleReprise : TOB     ;
    IndiceChampRow  : integer ;
    TypeChamp       : string ;
    Actif           : boolean ;
begin
  if Action = taConsult then Exit ;
  //
  // récupération de la ligne sélectionnée
  //
  Actif := False ;
  IndiceChampRow  := GChamp.Row -1;
  if (IndiceChampRow >= 0) and (IndiceChampRow < Tob_Reprise.detail.count) then
  begin
     TobFilleReprise := Tob_Reprise.Detail [IndiceChampRow];
     TypeChamp := TobFilleReprise.GetValue('GRC_TYPECHAMP') ;
     if (copy (TypeChamp, 1, 7) = 'VARCHAR') or (TypeChamp = 'COMBO') then Actif := True ;
  end;
  BTableCorr.Enabled := Actif ;
end;

procedure TFParamGPAO.GRE_ENTETELIGClick(Sender: TObject);
begin
  if GRE_ENTETELIG.Checked then
  begin
    GRE_DEBUTENREG.Enabled := True ;
  end else
  begin
     GRE_DEBUTENREG.Enabled := False ;
  end;
end;

end.

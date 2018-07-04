{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 17/02/2004
Modifié le ... : 17/12/2004
Description .. : Source TOM de la TABLE : IMMO (IMMO)
Suite ........ : - FQ 12514 - CA 16/08/2004
Suite ........ : - FQ 15080 - CA - 17/12/2004 - Au cas où sous ORACLE
Suite ........ : pou réviter les message 'conversion de variant'
Suite......... : - FQ 15625 - MBO 21:07:2005 - modif base HT - impact s/base déduct. et quote part
Suite......... : - FQ 16101 - MBO 25:07:2005 - coche dégrevement accessible quelque soit le mode
                                               d'amortissement (éco ou dégressif)
Suite......... : - FQ 16258 - MBO 26/07/2005 - pb affichage montant TTC
Suite......... : - FQ 16171 - BTY 07/05 Comparer le cumul éco antérieurs avec la base amortt et non avec le mtt HT
Suite......... : - FQ 13234 - MBO 31/08/2005 - remise à 0 base TP si modif compte immo corporelle en incorporelle
Suite......... : - FQ 16385 - MBO 01/09/2005 - suite annulation opération revenir sur liste des immos
Suite......... : - FQ 14277 - MBO 06/0
9/2005 - pas de modification de base sur contrat location longue durée
Suite......... : - FQ 12508 - MBO 06/09/2005 - Retouche le 26.09.2005 -Gestion de la touche F6
Mots clefs ... : TOM;IMMO

// TGA 14/09/2005 Appel de la dépréciation d'actif
Suite......... : - FQ 14363 - BTY 09/05 - Bouton Infos bloqué si le regroupement n'est pas un modèle
Suite......... : -          - PGR 09/2005    - Ajout valeur théorique sur Base amortissement fiscal
Suite......... : - FQ 16756 - MBO 17/09/2005 - borner base fiscalement déductible à la base éco
Suite......... : -          - PGR 03/10/2005 - I_DUREEFISC initialisé à 1 au lieu de 0 pour éviter des messages intempestifs
Suite......... : -          - PGR 26/10/2005 - Ajout coche de calcul du plan futur avec la VNC
Suite......... : -          - PGR 4/11/2005 - Pb affichage AmortDerog avec ou sans décimale
                            - TGA 09/11/2005 Test si immo déjà dépréciée
Suite..........: - FQ 16804 - MBO 14/11/2005 - autoriser exceptionnel sur immo avec plan VAR
Suite..........: - FQ 16755 - MBO 28/11/2005 - modif message 11
Suite..........: - FQ 17056 - PGR 28/11/2005 - ouverture systématique du champ Organisme
Suite..........: - FQ 17056 - PGR 28/11/2005 - ouverture systématique du champ Organisme
Suite..........: - FQ 17100 - BTY 11/05 - Perte du fiscal en validation de fiche après une cession
Suite..........: - FQ 17120 - MBO 06/12/2005 - calcul du nombre de jours écoulés sur la base de 360
                                               et prise en compte des dates suivant le mode d'amortissement
Suite..........: - FQ 17156 - BTY 12/2005 - Plantage en validation d'opération après avoir modifié la quotepart de l'immo
                   La modif de quotepart ajoutait un nouveau plan d'amortissement dans IMMOAMOR sans enreg dans IMMOLOG
Suite..........: - FQ 17183 - BTY 12/2005    - Sortir de la fiche immo après une dépréciation
Suite..........: - FQ 17184 - PGR 21/12/2005 - Validation coche VNC sans devoir tabuler
Suite..........: - FQ 17215 - TGA 21/12/2005 - GetParamSoc => GetParamSocSecur
Suite......... : - FQ 17259 - BTY - 01/06    - Nouveau top dépréciation dans IMMO
Suite......... : - FQ 17339 - MBO 20/01/2006 - on affiche date fin contrat = i_datesuspcb si alimentée
                                               (cette zone est alimentée par la levée d'option)
Suite......... : - FQ 17362 - MBO 25/01/2006 - déblocage de la date suspension contrat et activation des
                                               contrôles.
Suite......... : - FQ 17339 - MBO 25/01/2006 - stockage de la date fin de contrat dans une nouvelle zone
                 - TGA 26/01/2006 - Ajout init I_DOCGUID
                 - TGA 26/01/2006 - stockage du cumul dépréciation voir I_REPRISEDEP
Suite......... : - MBO 10/02/2006 - positionnement flag pour Recalcul du plan d'amortissement
                                    si modif i_reprisedep
Suite......... : - MBO 16/02/2006 - Modif TraiteDotationAnterieures pour que l'antérieur dépréciation
                                    n'impacte pas le calcul théorique des antérieurs sur les 2 plans
Suite......... : - TGA 03/03/2006 - Pas de cumul dépréciation pour crédit bail
                 - TGA 24/03/2006 - changement du bouton regroupement pour appel de la nouvelle saisie regroupement
                 - BTY 03/06      - FQ 17596 - Violation d'accès en création immo en série
Suite......... : - MBO 03/04/2006 - FQ 17704 - si tabu avant sur un radical compte (bourre avec des 0) et que le compte
                                    n'existait pas ---> boucle sur la liste des comptes
                 - TGA 04/04/2006 - init en création de I_REPRISEDEP
                 - BTY 04/06      - FQ 17516 Changement de regroupement devient une opération
                 - TGA 07/04/06   - Initialisation de champs IMMO, IMMOLOG
                 - BTY 04/06      - FQ 17826 Eclatement ne doit pas être ouvert sur une acquisition
Suite......... : -          - XVI 13/04/2006 - Ajustement du TagDispacht pour la rechercher des tiers
Suite......... : - XVI 18/04/2006 FD 3978 Ajoute des l'analytiques aux Immos S1
Suite......... : - XVI 25/04/2006 FD 3978 Ajoute de l'analytique mono ventilation aux Immos S1
                 - BTY 03/06      - FQ 17596 - Violation d'accès en création immo en série
                 - MBO 20/04/2006 - FQ 13047 - pouvoir saisir une date de mise en service sur un exercice non défini
                 - MBO 20/04/2006 - FQ 17941 - réafficher taux et antérieurs si changt méthode fiscale en création plan fiscal
                 - MBO 25/04/2006 - FQ 17923 - gestion nvelle opération : modif date de mise en service
                 - MBO 27/04/2006 - FQ 17569 - alim des zones datedebeco et datedebfis en création
                                               + correction sur alim date debut en méthode variable
                 - BTY 05/06      - FQ 18119 - Positionner les indicateurs de modif de compta dans PARAMSOC
                 - MBO 23/05/2006 - FQ 18218 - modif du montant HT et forçage de la base éco
                 - MBO 31.05.2006 - FQ 18261 - sauvegarde de la date d'achat en entrée fdateorigine pour mise à jour immolog dateop
                 - MBO 08.06.2006 - FQ 18330 - message sur montant HT a tort en création de fiche
                 - MVG 09.06.2006 - FQ 17569 - Suite
                 - MVG 09.06.2006 - FQ 18341 - Message d'erreur si on fait 2 fois F10 pour valider une ficher
                 - MVG 12.06.2006 - FQ 18367 - Si on corrige une fiche en erreur, il faut valider 2 fois par F10
                 - MBO 13.06.2006 - FQ 17569 - Suite : si plan VAR : si modif date deb eco ne pas appeler la saisie
                                               du plan UO si nb uO = 0
Suite......... : - BTY 06/06      - FQ 18385 - Exceptionnel bloqué par message si date début amortissement ECO > date fin exo en cours
Suite......... : - MBO 16/06/2006 - FQ 15625 - suite : décocher coche réintégration fiscale si remise à 0 de la base fiscalt déductible
Suite......... : - BTY 06/06      - FQ 18393 - Cause paramètre variable, modif appel éclatement/révision plan/exceptionnel/modif bases
Suite......... : - MBO 19/06/2006 - FQ 18430 - ajouter l'antérieur dépréciation aux antérieurs pour tester le minimum légal et
                                               déporter le test sur la validation de la fiche
Suite......... : - MBO 19/06/2006 - FQ 17500 - correction des tests anterieurs + deprciat > base amortissement
Suite......... : - Mbo 22/06/2006 - FQ 17569 - suite : modif éco en variable --> aligner la date du fiscal sur la date de mise service
Suite......... : - MBO 26/06/2006 - FQ 14054 - si création d'un cb avec dépôt de garantie : en afficheage n° cb tronqué en
                                               consultation de la fiche du dépôt de garantie
Suite......... :   BTY 06/06 version 7.10 Paramètres pour affectation DPI
Suite......... : - MBO 27/06/2006 - FQ 14324 - message d'avertissement en suppression d'une immo cloturée
Suite......... :   MVG 03/07/2006 - FQ 18514 - L'accès au tableau d'amortissement force le fait que la fiche est modifiée
                                               d'ou pb de recalcul après
// TGA - 04/07/2006 suppression des enregs de immomvtd
//XVI 12/07/2006 FQ (BL) 12746 contrôle sur immo ventilable et pas de ventilation paramétrée.
Suite......... : - MBO 12/07/2006 - FQ 15274 - autoriser la saisie d'une immo avec antérieurs en mode VARIABLE
Suite......... : - MBO 13/07/2006 - FQ 18582 - controler montant négatif sur reprise eco , fiscale et dépréciation
//XVI 13/09/2006 FQ (BL) 12952 Controle coherence ventilation analytique entre fiche IMMO et Compte général
//XVI 15/09/2006 FQ (BL) 12746 récupération de la correction faite le 12/07/2006 qui a disparue.
//XVI 20/09/2006 FQ (BL) 13092 correction d'orthographe au message de S1_VerifieVentilation
//XVI 27/09/2006 FQ (BL) 13303 Message d'avertissement lors la création d'une fiche IMMO sur exercice clôturé.
//XVI 15/01/2007 FQ (BL) 12746 Lors l'enregistrement on ne vérifie que ventilation analytique impossible.
//XVI 16/01/2007 FQ (BL) 12746 On testait le Paramsoc qui ne correspond pas....  
Suite......... : - BTY 07/06 Nouveaux paramètres pour affectation DPI
Suite......... : - BTY 07/06 Saisie Base ECO ou Affectation DPI antérieure à une immo
Suite......... : - BTY 07/06 Bloquer modif de dates si l'immo utilise une DPI
Suite......... : - MBO 25/08/2006 - correction conseil de compilation
Suite......... : - TGA 30/08/2006 - modif oépration en cours + suppression d'une dpi
Suite......... : - BTY 09/06 - Appel derrière le suivi DPI, bloquer les opérations
Suite......... : - BTY 09/06 - Appel derrière le suivi DPI, réactualiser les soldes disponibles au cas où une immo vient d'être traitée
Suite......... : - MBO 09/06 - Ajout de l'opération 'Prime d'équipement'
Suite......... : - BTY 09/06 - FQ 16149 Appliquer les restrictions utilisateur à l'Etablissement
Suite......... : - MBO 09/06 - suite prime d'équipement : ajout popup PopZoomPlan
Suite......... : - BTY 10/06 - FQ 18963 Montant erronés en retour d'annulation de cession en Web Access uniquement :
Suite......... : - Cause : en retour de toute opération, le pgm détecte une modif, demande enregistrer les modifs ? OUI/NON et réenregistre la fiche avec les anciennes données
Suite......... :   BTY 10/06 - Nouvelle opération : réduction de la prime d'équipement
Suite......... :   BTY 10/06 - FQ 18976 Revenir sur les modifs FQ 18963 car la demande enregistrer les modifs OUI/NON
Suite......... :   est causée par un pb dans opeencours.pas dans l'annulation de cession, et non par un BFerme.Click insuffisant
Suite......... :   BTY 10/06 - FQ 18968 Saisie antérieurs DPI possible sur une immo
Suite......... :   datant d'avant N-5 mais non encore amortie
Suite......... :   BTY 10/06 - FQ 18992 1- Saisie DPI sur immo acquise : solde dispo doit cumuler ceux des exos précédant l'exo en cours et
Suite......... :   non des exos précédant la ligne en cours du tableau de suivi d'où l'on vient
Suite......... :   FQ 18992 2- Erreur similaire pour le stockage : prendre les DPI dispo de N-5 à N-1
Suite......... :   BTY 10/06 Transformer la mention DPI stDPI en bouton bInfoDPI qui ouvre l'onglet Informations
Suite......... :   MBO 10/06 - Branchement de la saisie subvention
Suite......... :   BTY 10/06 Ne pas saisir de prime ou subvention sur une immo de type simulation
Suite......... :   BTY 10/06 Bouton Opération en cours caché après clôture des immos
Suite......... :   BTY 10/06 FQ 19013 Cacher la saisie de la base éco d'une immo non amortissable
Suite......... :   BTY 10/06 FQ 19015 Pas de modif base sur une immo non amortissable
Suite......... :   BTY 10/06 FQ 19019 En modif de fiche Suppression des DPI est accessible alorsque toutes les opérations sont disable
Suite......... :   BTY 11/06 Ajout réduction de la subvention d'investissement
Suite......... :   BTY FQ 19151 Date de fin de contrat non enregistrée sur un crédit-bail ou location longue durée
Suite......... :   BTY FQ 19260 12/06 Interdire la révision de plan à une immo NAM avec prime ou subvention
Suite......... :   MBO FQ 17512 03/07 gestion fiscale
Suite......... :   MBO FQ 19888 27/03/2007 - ne pas autoriser la modif plan UO si opérations
Suite......... :   MBO FQ 14013 05/04/2007 - masquer zones liées à la section onglet champs libres si pas S1
Suite......... :   MBO 16/04/2007 - branchement remplacement de composant
Suite......... :   MBO 27/04/2007 - pas de modif sur un composant - utilisation boolean fComposant
Suite......... :   BTY FQ 19511 Lancer directement la fonction derrière menu POPZOOMPLAN si fonction unique
Suite ........ :   BTY 05/07 FQ 20310 Empêcher I_DUREEREPRISE d'être négatif car décale le tracé export de l'enreg amortissement éco
Suite ........ :   BTY 05/07 FQ 20256 Enregistrer le type de dérogatoire dans i_typederoglia
Suite ........ :   BTY 05/07 FQ 20005 Ouvrir certains paramètres du contrat de CB/LOC après clôture ou opération
                   autre que levée option. Attention :  pas encore fini
Suite ........ :   MBO 07/06/07 modif condition accès à option de remplacement (présence plan fiscal)
Suite ........ :   MBO 07/06/07 si les 2 plans sont identiques mais gestion fiscale : on garde le fiscal
Suite ........ :   MBO 08/06/07 pas de contrôle sur un composant remplacé ou un composant de remplacement
Suite ........ :   MBO 11/06/07 FQ 20660 - correction d'un plantage que l'on ne voyait qu'en debug
                                (modif du mode linéaire en mode variable)
Suite ........ :   MBO 15/06/07 FQ 20555 - différence de comportement entre cwas et 2/3
                                (en cwas le setfield est actif même en mode browse)
Suite ........ :   MBO 22/06/07 FQ 20811 - Correction erreur en retour de modif de sections analytiques
Suite ........ :   BTY 26/06/07 FQ 20005 - Fin - Saisir des échéances au lieu de tranches
Suite ........ :   MBO 27/06/07 pas d'enregistrement modèle de regroupement sur immo issue d'un remplacement
Suite ........ :   MBO 29/06/07 composant de 2ème catégorie : afficher ce libellé dans type base fiscale
Suite ........ :   MBO 29/06/07 Minimum légal : si g° fiscale on le calcule avec durée la + longue
Suite ........ :   MBO 05/07/07 Correction sur popoup des opérations : affichage ou non annulation opération
Suite ........ :   MBO 06/07/07 pas d'enregistrement modèle de regroupement sur immo remplacée
Suite ........ :   MBO 10/07/07 Après cloture : le type base fiscale doit être 'rempl.composant' sur immo remplaçante
Suite ........ :   MBO 13/07/07 on ne teste plus I_remplace mais I_string1 pour immo remplaçante
Suite ........ :   MBO 19/07/07 FQ 20555 - pb du spinbutton durée fiscale avec perte de focus sur la fenêtre
Suite ........ :   MBO 20/07/07 FQ 21153 - plantage si saisie durée fiscale >1200 et validation sans sortir du champ
Suite ........ :   MBO 25/07/07 FQ 19620 - avertissement si durée > 180 et méthode dégressive
Suite ........ :   BTY 08/07 FQ 12503 - Onglet Comptable en cas de pb sur comptes associés
Suite ........ :   BTY 09/07 FQ 21340 - Complément FQ 20005 griser la zone de frais compris dans l'échéance si levée option
Suite ........ :   BTY 09/07 FQ 20005 - Suite. En modification d'échéancier, saisir des échéances sans pouvoir toucher aux dates
Suite ........ :   => La modif de dates se fait dans la fiche immo en changeant le nb d'échéances
Suite ........ :   BTY 09/07 FQ 21410 - Valeur base éco incorrecte avec mht=1790, tvar=350.84 et tvaR=350.84
Suite ........ :   MBO 17/09/07 FQ 21423 - si méth éco = VAR prédisposer la durée fiscale = durée éco
Suite ........ :   BTY 09/07 FQ 16101 incomplète cause ELIGILETP au lieu de ELIGIBLETP
Suite ........ :   BTY 09/07 FQ 21514 ELIGILETP => ELIGIBLETP
Suite ........ :   BTY 09/07 FQ 21502 Type base fiscale présente Rempl. Composant => remplacer items.delete par la propriété Plus qui filtre
Suite ........ :    Va avec le paramétrage &#@ de la tablette TITYPEBASEFISC
Suite ........ :   BTY 10/07 FQ 21646 Création auto lieu géographique en saisie directe fiche immo via la compta
Suite ........ :   BTY 11/07 FQ 21799 Ouvrir échéances suivantes et valeur résiduelle en création d'immo à loyers constants
Suite ........ :   BTY 11/07 Afficher indication si immo issue d'un bien au forfait (Agricoles) (i_pfr = 'X')
*****************************************************************}

unit IMMO_TOM;

interface

uses StdCtrls,
  windows,
  Controls,
  Classes,
  Math,
  menus,
  dialogs,
{$IFNDEF EAGLCLIENT}
  db,
  hdb,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fiche,
  fe_main,
  CodePost,
{$ELSE}
  eFiche,
  MaineAGL,
  eCodePost,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTob,UtobDebug,
  ParamSoc,
  HRichOle,
  LookUp,
  HTB97,
  AGLInit, // TheData
  { Comptabilité }
  // uLibExercice,
  {$IFDEF SERIE1}
  {$ELSE}
  Ventil,
  {$ENDIF}

  {$IFDEF VER150}
   Variants,
  {$ENDIF}

  { Amortissement}
  ImEnt,
  ImOutGen,
  ImContra,
  Outils,
  ImPlan,
  ImOuPlan,
  IMMOCPTE_TOM,
  PlanEche,
  ImSortie,
  ImEclate,
  ImMutati,
  ChanPlan,
  ImLevOpt,
  OpEnCour,
  ImModBas,
  PlanAmor,
  ListeEch,
  AMUNITEOEUVRE_TOF,
  AmPlanUO,
  AMREGROUPEMENT_TOF,
  TIMOTIFCESSION_TOT // Pour ne pas l'oublier dans le projet IMMO
  //PGR - 09/2005 - Changement de méthode
  //TGA - 09/2005 - Dépréciation d'actif
  ,AMCHGTMETHODE_TOF
  ,AMDEPRECIATION_TOF
  ,AMDEPRECGROUPE_TOF
  ,AMSREGRO_TOF
{$IFNDEF CMPGIS35}
  ,AMSYNTHESEDPI_TOF
{$ENDIF}
  ,AMMODIFSERVICE_TOF   // FQ 17923 - mbo - 04/2006
{$IFNDEF CMPGIS35}
  ,AMNDEDUCTION_TOF
  ,AmDPI
  ,AMBaseEco_TOF
  ,AmDPIAnterieurs_TOF
  ,AMImmoDPI_TOF   // FQ 17923 - mbo - 04/2006
  ,AMPrime_TOF
  ,PlanPrime
  ,AMReducPrime_TOF
  ,AMSubvention_TOF
  ,AMRemplacement_TOF
  ,PlanSbv
  ,AMReducSBV_TOF
  ,PlanEcheEche
{$ENDIF}
  {$IFNDEF SERIE1}
  ,Ent1
  {$ENDIF}
  ;

type
(*  TInfoLog = record  //YCP 25/08/05 déplacé
    TVARecuperable: double;
    TVARecuperee: double;
  end;

  TBaseImmo = record
    CompteImmo: string;
    Fournisseur: string;
    Reference: string;
    CodeEtab: string;
    DateAchat: TDateTime;
    MontantHT: double;
    MontantTVA: double;
    Quantite: double;
    Libelle: string;
  end;*)
  TOM_IMMO = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    fLeMode: TActionFiche;
    fContrat: TImContrat;
    fPlanAmor: TPlanAmort;
    fLog: TOB;
    fPlanUO: TPlanUO;
    fNature: string;
    FDPI: boolean;
    FDPIANT: boolean;
    fOperationEnCours: boolean;
    bChargementEnCours: boolean;
    fDateAchatIni: TDateTime;
    fDateOrigine: TDateTime;  // fq 18261 mbo 31.05.2006
    fDateMSini: TDateTime;
    CumEcoSav: double;
    CumFiscSav: double;
    ImmoDepotGarantieExiste: boolean;
    fDateFin: TDateTime;
    fCodeImmoLie: string;
    bPlanModifie: boolean;
    fDisableChangeField: boolean;
//    fFichier: TextFile;
    fCodeImmoDepotGarantie: string;
    fCurNbEcheances: integer;
    wTableLib: array[0..10] of string;
    fCodeSuppression : string;
    UpdateRecordEncours : boolean; //FQ 18341
    fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:
      TShiftState) of object;
    // 7.10 06/06 Paramètres pour affectation DPI
    fDateDebExoDPI : TDateTime;
    fDateFinExoDPI : TDateTime;
    fMontantMaxiDPI : double;
    // 07/06 BTY Paramètres DPI
{$IFNDEF CMPGIS35}
    fAmDPI : TAmDPI;
{$ENDIF}
    fDPIUtilisee : boolean;

    // TGA mt DPI a affecté
    FmontantDPI : Double;

    //mbo pour la gestion de remplacement de composant
    fComposant : boolean;

    {$IFNDEF SERIE1}
    FComboEtab : THValComboBox;
    function  GetComboEtab : THValComboBox;
    procedure GereEtablissement;
    procedure GereModeleEtablissement  (ModeleEtabt : string);
    {$IFNDEF EAGLSERVER}
    procedure PositionneEtabUser;
    {$ENDIF}
    property  ComboEtab : THValComboBox read GetComboEtab write FComboEtab;
    {$ENDIF}

    procedure OnCompteElipsisClick(Sender: TObject);
    procedure OnCompteDepotElipsisClick(Sender: TObject);
    procedure OnClickAmortissementFiscal(Sender: TObject);
    procedure OnClickReintegrationFiscale(Sender: TObject);
    procedure OnClickSurAmortissement(Sender: TObject);
    // TGA 06/09/2006
    procedure OnClickSaisieDPI(Sender: TObject);
    //PGR 26/10/2005 Ajout coche de calcul du plan futur avec la VNC
    procedure OnClickSurCalculAvecVnc(Sender: TObject);
    procedure OnClickSaisieEcheancier(Sender: TObject);
    procedure OnClickUniteOeuvre(Sender: TObject);
    procedure OnChangeOnglet(Sender: TObject);
    procedure MajComptesAssocies;
    procedure RazComptesAssocies;
    procedure EnableOngletsFiche;
    procedure EnableChamps;
    procedure EnableChampsCommuns;
    procedure EnableChampsLibre;
    procedure EnableChampsOngletContrat;
    procedure EnableChampsOngletGeneral;
    procedure EnableChampsOngletInfo;
    procedure EnableChampsOngletPlan;
    procedure InitEnregExistant;
    procedure AfficheMontantTTC;
    procedure AfficheCoutContrat;
    procedure ChargeHistoriqueOperations;
    procedure TraiteOrigineImmo;
    procedure GereSuramortissement;
    procedure EnableBoutonsInfos;
    procedure GereEco;
    procedure GereFiscal;
    procedure GerePlanVar; //ajout chantier fiscal fq 17512
    // mbo fq 17569 procedure GereReprise;
    procedure GereRepriseEco;
    procedure GereRepriseFisc;
    procedure GereSpinDureeEco(Methode: string);
    procedure GereSpinDureeFisc(Methode: string);   // fq 20555 - mbo
    procedure GereDepotGarantie;
    procedure GereGroupeEcheance;
    procedure ChargeLog;
    procedure AfficheNombreEcheance;
    procedure EnableBoutonBaseEco;  // FQ 19013

    // mbo fq 17569 function  IsRepriseVisible: boolean; remplacée par les 2 suivantes
    function  RepriseEcoVisible: boolean;
    function  RepriseFiscVisible:boolean;

    // Tga 15/06/2006
    function  Test_Reprise_Depreciation(Stchamp:String):boolean;

    procedure AfficheBaseTaxe;
    procedure AfficheTauxEco;
    procedure AfficheTauxFisc;
    //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
    procedure AfficheBaseFisc;
    procedure TraiteDotationAnterieures;
    procedure GereNatureBien;
    function EstUnChampCompte(St: string): boolean;
    function ControleMiniLegal: integer;  // fq 18430 controle minimum légal
    function ControleImmo: integer;
    function ControleCode: integer;
    function ControleComptes: integer;
    function ControleChamps: integer;
    function ControleLoyer: integer;
    function PlanAmortValide(var stMethode: string): boolean;
    function ControleDate: integer;
    function IsValideComptesAssocies: integer;
    function EnregistrePlan: boolean;
    function CreationImmoDepotGarantie: boolean;
    procedure EnregDepotGarantie;
    function GetPeriode: integer;
    procedure OnCessionClick(Sender: TObject);
    procedure OnEclatementClick(Sender: TObject);
    procedure OnMutationClick(Sender: TObject);
    procedure OnModeleRegroupementClick ( Sender : TObject );
    procedure OnModifBasesClick(Sender: TObject);
    procedure OnChangementContrat(stField: string);
    procedure GereNonAmortissable;
    procedure OnDepotGarantieClick(Sender: TObject);
    procedure AnnulationOperationClick(Sender: TObject);
    procedure DureeMethodeClick(Sender: TObject);
    //PGR - 09/2005 - Changement de méthode
    procedure ChangementMethodeClick(Sender: TObject);
    //TGA - 09/2005 - Dépréciation d'actif
    procedure OnDepreciationActifClick(Sender: TObject);
    procedure ModifMiseServiceClick(Sender: TObject);     //MBO - 04/2006 - fq 17923

    procedure SaisiePrimeClick(Sender: TObject);
    procedure SaisieSbvClick(Sender: TObject);
    procedure SaisieRemplacementClick(Sender: TObject);  // mbo - 04/2007

    procedure ElementExceptionnelClick(Sender: TObject);
    procedure LeveeOptionClick(Sender: TObject);
    // BTY 10/06 Réduction de la prime d'équipement
    procedure OnReductionPrimeClick(Sender: TObject);
    // BTY 11/06 Réduction de la subvention d'investissement
    procedure OnReductionSBVClick(Sender: TObject);
    procedure OnPopZoomActionPopup(Sender: TObject);
    procedure OnPopZoomPlan(Sender: TObject);  // ajout mbo pour popzoomPlan suite gestion prime
    procedure NbEcheancesEnter(Sender: TObject);
    procedure NbEcheancesExit(Sender: TObject);
    procedure OnLieuDblClick(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure OnVisuDepotGarantieClick(Sender: TObject);
    procedure OnOpeEnCoursClick(Sender: TObject);
    // BTY 10/06
    procedure OnInfoDPIClick(Sender: TObject);
    procedure OnAccesImmoLie(Sender: TObject);
    procedure OnChangementMethode(stField: string);
    procedure OnChangementDureeAmort(stField: string; bReinit : boolean);
    procedure OnPlanAmortClick(Sender: TObject);

    procedure OnPlanPrimeClick(Sender: TObject);   // ajout mbo pour prime
    procedure OnPlanSbvClick(Sender: TObject);   // ajout mbo pour subvention

    procedure OnPlanRemboursementClick(Sender: TObject);
    procedure OnShowZLS1(Sender: TObject);
    procedure OnClickVENTILABLE(Sender: TObject);
    procedure OnClickELIGIBLETP(Sender : TObject);
    procedure CompleteChampsVides;
    procedure OnChangeRegroupement ( Sender : TObject );
    procedure OnInformationRegroupementClick ( Sender : TObject );
    procedure OnDetailRegroupementClick ( Sender : TObject );
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift:
      TShiftState);
    //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
    procedure OnExitTypeBaseFisc ( Sender : TObject );
    procedure OnAfterFormShowEcran;
    procedure OnMouseDownI_BlocNote(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    // BTY 09:05 FQ14363
    procedure OnPopRegroupementPopup(Sender: TObject);
    // mbo - fq 17362
    procedure OnClickSuspensionCB(Sender: TObject);
    // TGA 24/03/2006
    procedure OnClickSaisieRegroupement(Sender: TObject);
    // mbo fq 17569
    function bPlanFiscal:boolean;
    function DateDebValide(queldate:string; affmess:boolean):boolean;
    // BTY 07/06 DPI
    procedure OnClickBaseEco(Sender: TObject);
    // TGA 21/07/2006
    function MiseAJourDPI(codeimmo:String):boolean;
    // TGA 30/08/2006
    procedure OnSuppressionDPIClick(Sender: TObject);
    // MBO 26/02/2007 pour chantier fiscal fq 17512
    Procedure AfficheDerogatoire;
    procedure OnClickGestionFiscale(Sender: TObject);  // chantier fiscal fq 17512
    procedure GereCocheVNF;
    procedure OnClickSurCalculAvecVNF(Sender: TObject);
    function ControleDerogatoire: integer;
    function ControleGestionFiscale: integer;
   {$IFDEF SERIE1}
    procedure SectionS1OnExit ( Sender : TObject) ; //XVI 25/04/2006 FD 3978 Analytique Mono ventilation
    procedure S1_VerifieVentilation(FaireMAJ : Boolean ) ; //XVI 13/09/2006 FQ 12952
    procedure TABLELIBRE1Change(Sender: TObject);
    Function S1_VerifCOherenceVentilation : integer ; //XVI 15/09/2006 FQ 12746
    {$ELSE}
    {$ENDIF !SERIE1}
    function LancerZoomPlanUnique  : Boolean;
  end;

procedure AMLanceFiche_FicheImmobilisation(Lequel: string; LeMode: TActionFiche;
  Nature: string; ParamDPI : string = ''); // 7.10 06/06 Appel pour affectation DPI );
function AMLanceFiche_FicheImmobilisationEnSaisie(T: TOB): string;
procedure EnregLogAcquisition(CodeImmo: string; DateOpe: TDateTime; PlanActif:
  integer; infos: TInfoLog);
procedure UpdateLogAcquisition(CodeImmo: string; DateOpe: TDateTime; infos:
  TInfoLog);
function CreationSaisieCompta(BaseImmo: TBaseImmo): string;

var
  gTobImmo : TOB;  // variable globale pour passer des paramètres à la fiche (création depuis la saisie comptable)

const
{$IFDEF SERIE1}
  TAGTIERS=8 ;
  TAGGENERAUX=3 ;
{$ELSE}
  TAGTIERS=2 ;
  TAGGENERAUX=1 ;
{$ENDIF SERIE1}
implementation

uses
  messages,   // pour PostMessage
  AMLISTE_TOF
  {$IFDEF SERIE1}
  , S1Util
  , VentilY//XVI 18/04/2006 FD 3978
  , ut0 //XVI 13/09/2006 FQ 12952
  , utY_ //XVI 15/09/2006 FQ 12746
  {$ELSE}
  {$ENDIF !SERIE1}
;
const
  MSG_DATEACHAT =
    'La date de début d''amortissement doit être postérieure ou égale à la date d''achat.';
  MSG_DATESERVICE =
    'La date de début d''amortissement doit être postérieure ou égale à la date de mise en service.';
  MSG_ECHDERPREM =
    'La date de première échéance doit être antérieure à la date de dernière échéance.';
  MSG_ECHPREMLIV =
    'La date de première échéance devrait être au moins égale à la date de livraison.';
  ASK_DATEACQDEBEX =
    'La date d''acquisition est antérieure à la date de début d''exercice sur un dossier clos.#10#13S''agit-il d''une opération exceptionnelle qui permet de rectifier la situation initiale ?';
  MSG_MODIFACHAT =
    'Vous avez modifié la date d''achat. Il convient de vérifier les dates de début d''amortissement de l''immobilisation.';
  MSG_MODIFSERVICE =
    'Vous avez modifié la date de mise en service. Il convient de vérifier les dates de début d''amortissement de l''immobilisation.';
  MSG_MODIFMONTANTHT =
    'Vous avez modifié le montant HT. Il convient de vérifier la base d''amortissement.';
  MSG_ERREURBASE =
    'La valeur de la base d''amortissement est incorrecte.';

  // modif mbo - fq 16756 - HMC: array[1..37] of string =
  // PGR -11/2005 - Ajout 39
  // mbo - 25/11/2005 ancien 11 = {11}'Vous devez renseigner la valeur de la réintégration fiscale.',
  // BTY FQ 18385
  // BTY 07/06 DPI
  HMC: array[1..53] of string =
  {1}('Le code immobilisation est incorrect. Vous devez le modifier.',
    {2}'Le code que vous avez saisi existe déjà. Vous devez le modifier.',
    {3}'Le code que vous avez saisi a déjà été utilisé. Vous devez le modifier.',
    {4}'Le compte d''immobilisation n''est pas renseigné.',
    {5}'Le compte d''immobilisation n''est pas renseigné correctement.',
    {6}'Le compte de charge n''est pas renseigné correctement.',
    {7}'L''organisme n''est pas renseigné correctement.',
    {8}'Compte associé non renseigné correctement.',
    {9}'Vous devez renseigner le lieu géographique.',
    {10}'Vous devez renseigner le plan d''amortissement.',
    {11}'Vous devez renseigner le plafond de calcul pour la réintégration fiscale.',
    {12}'Le taux d''amortissement économique ne peut être supérieur au taux fiscal.',
    {13}'Le cumul fiscal doit être supérieur au cumul économique.',
    {14}'La saisie n''est pas autorisée sur un exercice inexistant.',
    {15}'Vous devez renseigner la valeur du dépôt de garantie.',
    {16}'Le code immobilisation du dépôt de garantie existe déjà.',
    {17}'Le compte du dépôt de garantie n''est pas renseigné correctement.',
    {18}'La date de dépôt de garantie doit être supérieure à la date de début de contrat.',
    {19}'La date de suspension des effets du contrat n''est pas correcte.',
    {20}'La date de suspension des effets du contrat doit être antérieure à la date de dernière échéance.',
    {21}'Problème au niveau des comptes associés.',
    {22}'Le montant des frais est supérieur au montant de l''échéance.',
    {23}'La somme des loyers doit être au moins égale au montant HT. Voulez-vous continuer ?',
    {24}'L''amortissement économique n''est pas renseigné.',
    {25}'Le montant ne peut être supérieur au montant HT.',
    {26}'La quantité est nulle. Voulez-vous continuer l''enregistrement ?',
    {27}'La date de première échéance devrait être au moins égale à la date de livraison.',
    {28}'La date de première échéance doit être antérieure à la date de dernière échéance.',
    {29}'La reprise ne peut être supérieure à la valeur de base.',
    {30}'L''amortissement antérieur pratiqué est inférieur à l''amortissement linéaire.'
    + #10#13'Pour respecter le principe de l''amortissement minimum, il convient peut-être de vérifier' +
    ' la valeur saisie.#10#13Voulez-vous modifier la valeur ?',
    {31}'Vous devez renseigner le libellé.',
    {32}'Le montant HT ne peut être nul. Veuillez saisir une valeur.',
    {33}'La date de mise en service ne peut être antérieure à la date d''achat.',
    {34}'La valeur de la base d''amortissement est incorrecte.',
    {35}'Le nombre d''unités d''oeuvre doit être renseigné.',
    {36}'Le nombre total d''unités d''oeuvre saisi est incorrect.',
    {37}'L''amortissement ne peut être variable dans le cas où la date de début d''amortissement correspond à un exercice antérieur.',
    {38}'La base fiscalement déductible ne doit pas être supérieure à la base d''amortissement économique.',
    {39}'Le cumul économique plus le cumul des dépréciations ne peut être supérieur à la base amortissable.',
    {40}'Le cumul fiscal plus le cumul des dépréciations ne peut être supérieur à la base amortissable.',
    {41}'La date de début d''amortissement économique doit être postérieure ou égale à la date d''amortissement fiscal.',
    {42}'La date de début d''amortissement doit être postérieure ou égale à la plus petite des dates entre achat et mise en service.',
    {43}'Il n''est pas possible de saisir un élément exceptionnel ' +
        'si l''immobilisation n''est pas amortie sur l''exercice en cours.',
    {44}'Affectation non autorisée, les DPI utilisables ont déjà été réintégrées.' ,
    {45}'Confirmez-vous la suppression des DPI affectées ?',
    {46}'Affectation non autorisée, du fait de la présence d''une base fiscalement déductible.',
    {47}'(Dérogatoire - Réintégration/déduction) différent de (cumul fiscal - cumul économique).',
    {48}'(Dérogatoire - Réintégration/déduction) différent de (cumul fiscal - cumul économique - cumul dépréciation).',
    {49}'Affectation non autorisée sur une immobilisation ayant une gestion fiscale.',
    {50}'Il n''est pas possible d''enregistrer un antérieur dépréciation sur une immobilisation ayant une gestion fiscale.',
    {51}'La durée d''amortissement doit être inférieure ou égale à 1200 mois.',
    {52}'La durée d''amortissement économique est supérieure à 180 mois (15 ans) et le mode d''amortissement est dégressif.'
    + #10#13'Il convient peut-être de modifier la durée saisie' +
    ' .#10#13Voulez vous modifier la durée économique ?',
    {53}'La durée d''amortissement fiscale est supérieure à 180 mois (15 ans) et le mode d''amortissement est dégressif.'
    + #10#13'Il convient peut-être de modifier la durée saisie' +
    ' .#10#13Voulez vous modifier la durée fiscale ?');
    // mbo - fq 14054 - modif du message 6 car trop long et n° immo tronqué
  HM2: array[0..10] of string =
  ('Eclatement de ',
    'Levée d''option sur contrat',
    'Mutation de l''immobilisation',
    'Option levée',
    'Eclatement vers',
    'Dépot de garantie versé',
    'Versement suite à un contrat crédit-bail ou location',
    'Plan d''amortissement',
    'Plan d''amortissement théorique',
    'Durée déjà amortie : %d mois - %d jours',
    'Le montant de la TVA récupérée ne peut en aucun cas être supérieur au montant de la TVA récupérable.'
    );

function CreationSaisieCompta(BaseImmo: TBaseImmo): string;
var
  t: tob ;
begin
  t:=Tob.create('IMMO',nil,-1) ;
  t.putValue('I_DATEPIECEA', BaseImmo.DateAchat) ;
  t.putValue('I_COMPTEIMMO', BaseImmo.CompteImmo);
  t.putValue('I_REFINTERNEA', BaseImmo.Reference);
  t.putValue('I_QUANTITE', BaseImmo.Quantite) ;
  t.putValue('I_LIBELLE', BaseImmo.Libelle) ;
  t.putValue('I_MONTANTHT', BaseImmo.MontantHT) ;
  {$IFDEF SERIE1}
  t.putValue('I_VENTILABLE', Hent1.BoolToStr_(BaseImmo.Ventilable)) ; //XVI 18/04/2006 FD 3978
  {$ELSE}
  {$ENDIF !SERIE1}
  Result := AMLanceFiche_FicheImmobilisationEnSaisie(t) ;
  t.free ;
end;

procedure AMLanceFiche_FicheImmobilisation(Lequel: string; LeMode: TActionFiche;
  Nature: string; ParamDPI : string = ''); // 7.10 06/06 Appel pour affectation DPI );
var
  stAction: string;
begin
  gTobImmo := nil;
  if LeMode = taCreat then
    stAction := 'ACTION=CREATIONSERIE'
  else if LeMode = taCreatOne then
    stAction := 'ACTION=CREATION'
  else if LeMode = taModif then
    stAction := 'ACTION=MODIFICATION'
  else
    stAction := 'ACTION=CONSULTATION';
  stAction := stAction + ';' + Lequel + ';' + Nature;
  // 7.10 06/06
  if ParamDPI <> '' then
     stAction := stAction + ';' + ParamDPI;
  //

  AGLLanceFiche('AM', 'AMFICHE', '', Lequel, stAction);
end;

function AMLanceFiche_FicheImmobilisationEnSaisie(T: TOB): string;
begin
  if GetParamSocSecur('SO_EXOCLOIMMO','') = VHImmo^.Encours.Code then
  {$IFDEF SERIE1}
  begin
     PGIInfo('Vous ne pouvez pas créer de fiche immobilisation sur un exercice cloturé') ;
     exit;
  end ;
  {$ELSE}
    exit;
  {$ENDIF !SERIE1}
  if ImBlocage(['nrCloture'], False, 'nrAucun') then
    Exit;
  CreationLieuParDefaut; // FQ 21646
  gTobImmo := T;
  AGLLanceFiche('AM', 'AMFICHE', '', '', 'ACTION=CREATION');
  Result := gTobImmo.GetValue('I_IMMO');
  gTobImmo := nil;
end;


procedure TOM_IMMO.OnNewRecord;
var Exo : TImExoDate;
begin
  inherited;
  CumEcoSav := 0;
  CumFiscSav := 0;
  SetField('I_IMMO', NouveauCodeImmo);
  SetField('I_NATUREIMMO', fNature);
  SetField('I_ETAT', 'OUV');
  SetField('I_QUALIFIMMO', 'R'); // Par défaut Qualif Immo=réèlle.
  SetField('I_TYPELOYERCB', 'LCO'); // Par défaut type de loyer=loyer constant
  SetField('I_PERIODICITE', 'MEN'); // Par défaut Période=1 mois
  SetField('I_VERSEMENTCB', 'AVA'); // Par défaut Versement Avance
  SetField('I_QUANTITE', 1); // Par défaut quantité=1;
  SetField('I_PAYS', 'FRA'); // France par défaut
  SetField('I_NATUREBIEN', 'GEB'); // Généralités des biens par défaut
  SetField('I_DEVISE', V_PGI.DevisePivot);

  // MVG 07/06/07 FQ 20615
  // InitOpeEnCoursImmo(TQuery(DS), '-', '-', '-', '-', '-', '-', '-', '-', '-', '-','-');
  SetField('I_OPEMUTATION','-');
  SetField('I_OPEECLATEMENT','-');
  SetField('I_OPECESSION','-');
  SetField('I_OPECHANGEPLAN','-');
  SetField('I_OPELIEUGEO','-');
  SetField('I_OPEETABLISSEMENT','-');
  SetField('I_OPELEVEEOPTION','-');
  SetField('I_OPEMODIFBASES','-');
  SetField('I_OPEDEPREC','-');
  SetField('I_OPEREG','-');
  SetField('I_OPERATION','-');

  SetControLChecked('bAmortissementFiscal', false);

  SetControLChecked('bGestionFiscale', false);  // chantier fiscal fq 17512
  SetControLChecked('bCalculAvecVNF', false);
  SetControlEnabled('bGestionFiscale', false);
  SetControlEnabled('bCalculAvecVNF', false);

  if (fNature = 'CB') or (fNature = 'LOC') then
    SetControlChecked('bDepotGarantie', false);
  fPlanAmor.Free;
  // CA - 18/12/2002 - Suppression systématique de plan : on part sur des bases saines !!!
  fPlanAmor := TPlanAmort.Create(true);
  fPlanAmor.Charge(TQuery(DS));
  SetField('I_ETABLISSEMENT', VHImmo^.EtablisDefaut);

  if (THValCombobox(GetControl('I_LIEUGEO')).Values.Count > 0) then
    SetField('I_LIEUGEO', THValCombobox(GetControl('I_LIEUGEO')).Values[0]);

  SetControlProperty('I_LIEUGEO', 'ItemIndex', 0);
  SetControlEnabled('I_NATUREIMMO',((fLeMode = taCreatEnSerie) and (fNature='')) or ( gTobImmo <> nil ));
  SetControlEnabled('I_BASEECO', true);
  SetControlEnabled('I_DATEPIECEA', true);
  SetControlEnabled('I_DATEAMORT', true);
  SetControlEnabled('I_MONTANTHT', true);
  SetControlEnabled('I_TVARECUPERABLE', true);
  SetControlEnabled('I_TVARECUPEREE', true);
  SetControlEnabled('I_QUANTITE', true);
  SetControlEnabled('I_BASETAXEPRO', true);
  SetField('I_DATEPIECEA', VHImmo^.Encours.Deb);

  // Ajout mbo fq 17569
  fDateAchatIni := GetField('I_DATEPIECEA');
  fDateMSini := GetField('I_DATEPIECEA');
  fDateOrigine:= GetField('I_DATEPIECEA'); // fq 18261 mbo 31.05.2006
  SetField('I_DATEDEBECO', VHImmo^.Encours.Deb);
  SetField('I_DATEDEBFIS', VHImmo^.Encours.Deb);

  if gTobImmo <> nil then
  begin
    { FQ 15755 - Valeur de nature par défaut si création depuis la saisie }
    if ExisteSQL ('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="'+gTobImmo.GetValue('I_COMPTEIMMO')+
          '" AND G_GENERAL<="'+VHImmo^.CpteFinSup+'" AND G_GENERAL>="' +VHImmo^.CpteFinInf + '"') then
      fNature := 'FI' else fNature := 'PRO';
    SetField('I_NATUREIMMO', fNature);
    SetField('I_DATEAMORT', gTobImmo.GetValue('I_DATEAMORT'));
    SetField('I_DATEPIECEA', gTobImmo.GetValue('I_DATEPIECEA'));
    SetField('I_COMPTEIMMO', gTobImmo.GetValue('I_COMPTEIMMO'));
    SetField('I_MONTANTHT', gTobImmo.GetValue('I_MONTANTHT'));
    SetField('I_TVARECUPERABLE', gTobImmo.GetValue('I_TVARECUPERABLE'));
    SetField('I_TIERSA', gTobImmo.GetValue('I_TIERSA'));
    SetField('I_REFINTERNEA', gTobImmo.GetValue('I_REFINTERNEA'));
    SetField('I_VALEURACHAT', gTobImmo.GetValue('I_VALEURACHAT'));
    SetField('I_QUANTITE', gTobImmo.GetValue('I_QUANTITE'));
    SetField('I_ETABLISSEMENT', gTobImmo.GetValue('I_ETABLISSEMENT'));
    SetField('I_LIBELLE', gTobImmo.GetValue('I_LIBELLE'));
    // Ajout mbo fq 17569
    SetField('I_DATEDEBECO', gTobImmo.GetValue('I_DATEDEBECO'));
    SetField('I_DATEDEBFIS', gTobImmo.GetValue('I_DATEDEBFIS'));

    // mbo fq 17512 chantier fiscal fq 17512
    SetField('I_REPRISEDR', gTobImmo.GetValue('I_REPRISEDR'));
    SetField('I_REPRISEFEC', gTobImmo.GetValue('I_REPRISEFEC'));
    {$IFDEF SERIE1}
    SetField('I_VENTILABLE', gTobImmo.GetValue('I_VENTILABLE')); //XVI 06/06/2006
    {$ELSE}
    {$ENDIF !SERIE1}
  end
  else
  begin
    SetField('I_DATEAMORT', VHImmo^.Encours.Deb);
    SetField('I_REPRISEECO',0);
    SetField('I_MONTANTHT' ,0);

  end;
  SetField('I_DATEDEBUTECH', GetField('I_DATEPIECEA'));
  { FQ 15752 - Date de fin d'échéance par défaut renseignée correctement }
  if ImQuelDateDeExo (ImQuelExoDt(GetField('I_DATEPIECEA')), Exo ) then
    SetField('I_DATEFINECH', Exo.Fin)
  else SetField('I_DATEFINECH', VHImmo^.Encours.Fin);
  SetField('I_DATEDEPOTGAR', iDate1900);
  SetField('I_DATESUSPCB', iDate1900);
  SetField('I_DATEFINCB', iDate1900);  // ajout mbo 25.01.06

  // ajout tga 25.01.06
  SetField('I_REPRISEDEP',0);
  SetField('I_ECCLECR','');
  SetField('I_DOCGUID','');
  // ajout tga 04.04.06
  SetField('I_REPRISEDEPCEDEE',0);
  // ajout tga 07.04.06
  SetField('I_OPEREG','-');
  SetField('I_SUSDEF','A');
  SetField('I_REGLECESSION','NOR');
  SetField('I_NONDED','-');
  SetField('I_REPRISEINT',0);
  SetField('I_REPRISEINTCEDEE',0);
  SetField('I_DPI','-');
  SetField('I_DPIEC','-');
  SetField('I_CORRECTIONVR',0);
  SetField('I_CORVRCEDDE',0);
  SetField('I_SUBVENTION','NON');
  SetField('I_SBVPRI',0);
  SetField('I_SBVMTC',0);
  SetField('I_SBVPRIC',0);
  SetField('I_SBVEC','C');
  SetField('I_SBVDATE',iDate1900);
  SetField('I_CPTSBVR','');
  SetField('I_CPTSBVB','');
  //SetField('I_PRF','-');
  SetField('I_PFR','-');
  SetField('I_COEFDEG',0);
  SetField('I_AMTFOR',0);
  SetField('I_AMTFORC',0);
  SetField('I_ACHFOR',iDate1900);
  SetField('I_PRIXACFORC',0);
  SetField('I_VNCFOR',0);
  SetField('I_DURRESTFOR',0);
  SetField('I_DATEDEBECO',GetField('I_DATEAMORT'));
  SetField('I_DATEDEBFIS',GetField('I_DATEAMORT')); // mbo - 18.04.06

  //ajout pour chantier fiscal fq 17512
  SetField('I_NONDED','-');
  SetField('I_FUTURVNFISC','-');
  SetField('I_REPRISEFEC',0);
  SetField('I_REPRISEDR',0);
  SetField('I_REPRISEFECCEDEE',0);
  SetField('I_REPRISEFDRCEDEE',0);


  if (fContrat <> nil) and ((fNature = 'CB') or (fNature = 'LOC')) then
  begin
    fContrat.RazEcheances;
    fContrat.Charge(TQuery(DS));
  end
  else if ((fNature = 'CB') or (fNature = 'LOC')) then fContrat := TImContrat.Create;
  GereGroupeEcheance;
  SetControlEnabled('SaisieEcheancier', (not
    GetControlEnabled('GroupeEcheance')));
  SetControlChecked('cbI_DATESUSPCB', false);
  ImmoDepotGarantieExiste := false;
  AfficheMontantTTC;
  EnableChamps;
  SetControlVisible('lOrigineImmo', False);
  SetControlVisible('OrigineImmo', True);
  SetControlCaption('OrigineImmo', '');
  if (fNature = 'CB') or (fNature = 'LOC') then
  begin
    AfficheNombreEcheance;
    AfficheCoutContrat;
  end;
  fDateAchatIni := iDate1900;
  fDateMSini := iDate1900;
  fDateOrigine := iDate1900;  // fq 18261 mbo 31.05.2006
  if fPlanUO <> nil then fPlanUO.Raz
  else fPlanUO := TPlanUO.Create;
  {$IFNDEF SERIE1}
  // BTY 09/06 FQ 16149
  if DS.State=dsInsert then
     begin
     FComboEtab := nil;
     GetComboEtab;
     GereEtablissement;
     end;
  //if DS.State=dsInsert then PositionneEtabUser(THValComboBox(GetControl('I_ETABLISSEMENT')));
  SetField('I_ETABLISSEMENT',GetControlText('I_ETABLISSEMENT'));
  {$ENDIF}

  // ajout mbo 29.06.07
  // FQ 21502 Items.Delete met du désordre dans les valeurs
  //THValComboBox(GetControl('TYPEBASEFISC')).Items.Delete(THValComboBox(GetControl('TYPEBASEFISC')).Items.IndexOf(RechDom('TITYPEBASEFISC', 'RC2', FALSE)));
  //THValComboBox(GetControl('TYPEBASEFISC')).Items.Delete(THValComboBox(GetControl('TYPEBASEFISC')).Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE)));
  SetControlProperty('TYPEBASEFISC','Plus',' AND CO_CODE<>"TR2"');
end;

procedure TOM_IMMO.OnDeleteRecord;
begin
  (*  if not ConfirmationSuppressionImmo ( GetField('I_IMMO'), GetField('I_ETAT') ) then
  begin
    LastError := -1;
    LastErrorMsg := '';
  end;
  if not JaiLeDroitSupprimerImmo(GetField('I_IMMO'),GetField('I_CHANGECODE')) then
  begin
    LastError := -1;
    LastErrorMsg := '';
  end else*)

  // fq 14324 - ajout du test si immo a fait l'objet d'une cloture (confirmationsuppressionimmo)
  if ValideSuppressionImmo ( GetField('I_IMMO'), false ) then
  begin
     if not ConfirmationSuppressionImmo ( GetField('I_IMMO'), GetField('I_ETAT') ) then
     begin
       LastError := -1;
       LastErrorMsg := '';
     end else
     begin
       fCodeSuppression := GetField('I_IMMO');
       //Tga 06/09/2006 Maj Immomvtd suppression de l'immobilisation
{$IFNDEF CMPGIS35}
       AM_MAJ_IMMOMVTD('S',fCodeSuppression,'',0);
{$ENDIF}
     end;
  end else
  begin
     LastError := -1;
     LastErrorMsg := '';
  end;
  inherited;
end;

function TOM_IMMO.EnregistrePlan: boolean;
var
  BaseAmort: double;
  //Qlog : TQuery;
  i:integer;
  BaseEco, BaseFisc:double;
begin
  Result := True;
  BaseEco := 0;
  BaseFisc := 0;
  if ((GetField('I_NATUREIMMO') <> 'CB') and (GetField('I_NATUREIMMO') <> 'PRO'))
    then
    exit;

  //mbo fq 17569 ajout du test sur les dates
  //mbo chantier fiscal fq 17512 - ajout du test sur les calculs avec valeur nette F ou C
  //mbo 07.06.07 ajout du test sur gestion fiscale
  //            (si les 2 plans sont identiques et que la gestion fiscale est cochée : on garde le fiscal
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked)
    and (GetField('I_BASEFISC') = GetField('I_BASEECO'))
    and (GetField('I_DUREEFISC') = GetField('I_DUREEECO'))
    and (GetField('I_METHODEFISC') = GetField('I_METHODEECO'))
    and (GetField('I_REPRISEFISCAL') = GetField('I_REPRISEECO'))
    and (GetField('I_DATEDEBFIS') = GetField('I_DATEDEBECO'))
    and (GetCheckBoxState('bGestionFiscale')<> cbChecked)
    and (GetCheckBoxState('bCalculavecVNF')= GetCheckBoxState('bCalculavecVNC')) then
  begin

    // BTY 11/05 Ne pas supprimer le fiscal si base ECO distincte de base FISC AVANT CESSION
    if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'" AND I_OPECESSION="X"') then
      begin
      // fLog contient déjà les enregs IMMOLOG
      if fLog <> nil then
         begin
         if fLog.Detail.Count > 0 then
            begin
            for i := 0 to fLog.Detail.Count - 1 do
                begin
                if  (fLog.Detail[i].GetValue('IL_DATEOP') >= VHImmo^.Encours.Deb)
                and (fLog.Detail[i].GetValue('IL_TYPEOP') = 'CES') then
                     begin
                     BaseEco := fLog.Detail[i].GetValue('IL_BASEECOAVMB');
                     BaseFisc := fLog.Detail[i].GetValue('IL_BASEFISCAVMB');
                     break;
                     end;
                end;
            end;
         end;
      end;

    if (BaseEco = BaseFisc) then
      begin
      //

      SetField('I_BASEFISC', 0);
      SetField('I_DUREEFISC', 1);   //PGR -03/10/2005 1 au lieu de 0 pour éviter messages intempestifs
      SetField('I_TAUXFISC', 0);
      SetField('I_REPRISEFISCAL', 0);
      SetField('I_METHODEFISC', '');
      SetField('I_FUTURVNFISC', '-');  // ajout chantier fiscal fq 17512
      SetField('I_NONDED','-');

      SetControlText('TYPEBASEFISC', ''); //PGR - 09/2005 - Valeur théorique sur Base amort fiscal
      bPlanModifie := True;
      SetControlChecked('bAmortissementFiscal', false);

      SetField('I_REPRISEDR', 0); // chantier fiscal  fq 17512
      SetField('I_REPRISEFEC', 0);
      SetControlChecked('bGestionFiscale', false);
      SetControlChecked('bCalculavecVNF', false);
      SetControlEnabled('bGestionFiscale', false);
      SetControlEnabled('bCalculavecVNF', false);
      end;
  end;

  // TGA 26/04/2006 + mbo fq 17569 saisie des 2 dates 03.05.2006

  BaseAmort := GetFIeld('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
               Getfield('I_TVARECUPEREE');
  SetField('I_BASEAMORDEBEXO', BaseAmort);
  SetField('I_BASEAMORFINEXO', BaseAmort);

  //ajout pour chantier fiscal  fq 17512
  if (GetCheckBoxState('bGestionFiscale') = cbChecked) then
       SetField('I_NONDED', 'X')
  else SetField('I_NONDED', '-');

  if (GetCheckBoxState('bCalculavecVNF') = cbChecked) then
       SetField('I_FUTURVNFISC', '***')
  else SetField('I_FUTURVNFISC', '');


  if (bPlanModifie) or (fLeMode in [taCreat..taCreatOne]) or (DS.State =
    dsInsert) then
  begin
    // Suppression du plan de la base si plan déjà créé ==> évite d'avoir plsrs plan CRE.
    if (fPlanAmor.NumSeq > 0) and (fPlanAmor.TypeOpe = 'CRE') then
    begin
      fPlanAmor.Delete;
      fPlanAmor.NumSeq := fPlanAmor.NumSeq - 1;
      SetField('I_PLANACTIF', GetField('I_PLANACTIF') - 1);
    end;
    fPlanAmor.Recalcul(TQuery(DS));
    Result := fPlanAmor.Sauve;
    bPlanModifie := false;
    SetField('I_PLANACTIF', fPlanAmor.NumSeq);
  end
  else
    Result := True;
//  SetField('I_REPRISEECO', fPlanAmor.AmortEco.Reprise);
  SetField('I_DUREEREPRISE', fPlanAmor.NbMoisReprise);
//  SetField('I_REPRISEFISCAL', fPlanAmor.AmortFisc.Reprise);
  SetField('I_BASEAMORDEBEXO', fPlanAmor.BaseDebutAmort);
  SetField('I_BASEAMORFINEXO', fPlanAmor.BaseFinAmort);
  SetField('I_DATEDERMVTECO', fPlanAmor.GetDateFinAmortEx(fPlanAmor.AmortEco));

  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
      SetField('I_DATEDERNMVTFISC', fPlanAmor.GetDateFinAmortEx(fPlanAmor.AmortFisc));

  {modif pour chantier fiscal fq 17512
  // ajout pour fq 15274 - mbo
  if (GetField('I_METHODEECO') = 'VAR') and
     (GetField('I_REPRISEFISCAL')< GetField('I_REPRISEECO')) then
       SetControlProperty('AmortDerog','Value', 0)
  else
       SetControlText('AmortDerog', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
                      Getfield('I_REPRISEECO'), 13, V_PGI.OkDecV, '', TRUE));
  }
  //AfficheDerogatoire;

  // FQ 20256
  SetField('I_TYPEDEROGLIA', TypeDerogatoire(nil,TQuery(DS)));
end;

procedure TOM_IMMO.CompleteChampsVides;
begin
  if ((GetField('I_NATUREIMMO') = 'LOC') or (GetField('I_NATUREIMMO') = 'FI'))
    then
  begin
    if GetField('I_METHODEECO') = '' then
    begin
      SetField('I_METHODEECO', 'NAM');
      SetField('I_DUREEECO', 0);
      SetField('I_TAUXECO', 0);
    end;
  end;
  // FQ 19151 Création ou modif, enregistrer la date fin de contrat
  if ((GetField('I_NATUREIMMO') = 'CB') or (GetField('I_NATUREIMMO') = 'LOC')) then
     SetField ('I_DATEFINCB', StrToDate(GetControlText ('dtFinContrat')));
end;

procedure TOM_IMMO.OnUpdateRecord;
var
  iErreur: integer;
  stNature: string;
  Infos: TInfoLog;
Begin
  // FQ 18341
  if UpdateRecordEncours then
  begin
    LastError := -1;
    LastErrorMsg := '';
    UpdateRecordEncours:=false;
    exit
  end;
  UpdateRecordEncours:=true;
  // FQ 18341
  inherited;
  CompleteChampsVides;
  iErreur := ControleImmo;
  LastError := iErreur;
  if iErreur <> 0 then UpdateRecordEncours:=false; //MVG 18367
  if iErreur > 0 then
    LastErrorMsg := HMC[iErreur]
  else if iErreur < 0 then
    LastErrorMsg := '';

  stNature := GetField('I_NATUREIMMO');
  if (iErreur = 0) then
  begin
      //PGR 26/10/2005 Ajout coche de calcul du plan futur avec la VNC
     if (GetCheckBoxState('bCalculAvecVnc') = cbChecked) then
       SetField('I_JOURNALA', '***')
     else SetField('I_JOURNALA', '');
    { Enregistrement du plan d'unités d'oeuvre }
    if GetField('I_METHODEECO') = 'VAR' then
    begin
      fPlanUO.SetCodeImmo(GetField('I_IMMO'));
      fPlanUO.Enregistre;
    end;

    //ajout pour chantier fiscal  fq 17512
    if (GetCheckBoxState('bGestionFiscale') = cbChecked) then
       SetField('I_NONDED', 'X')
     else SetField('I_NONDED', '-');

    if (GetCheckBoxState('bCalculavecVNF') = cbChecked) then
       SetField('I_FUTURVNFISC', '***')
     else SetField('I_FUTURVNFISC', '');

    if EnregistrePlan then
    begin

      // TGA 21/07/2006 Mise à jour de la DPI=======================================
{$IFNDEF CMPGIS35}
      IF (FDPI=True) OR (FDPIANT=True) Then
         Begin
           IF MiseAJourDPI(GetField('I_IMMO')) = True Then
             Begin
                SetField('I_DPI', 'X');
                // TGA 30/08/2006 SetField('I_DPIEC', 'X');
                SetParamSoc('SO_IMMODPIEC', 'X');
                if fDPIAnt and (fAmDPI.St = IntToStr(2)) then
                   // La base fiscalement déductible doit contenir la DPI => l'ajouter
                   SetField('I_REINTEGRATION',
                      Arrondi(GetField('I_REINTEGRATION')+fAmDPI.Montant, V_PGI.OkDecV));
           End;
         END;
      //============================================================================

{$ENDIF}

      if ((stNature = 'PRO') or (stNature = 'FI')) then
        SetField('I_COMPTEREF', GetField('I_COMPTEIMMO'))
      else if ((stNature = 'CB') or (stNature = 'LOC')) then
        SetField('I_COMPTEREF', GetField('I_COMPTELIE'));

      if (DS.State = dsInsert) then
        EnregLogAcquisition(GetField('I_IMMO'), GetField('I_DATEPIECEA'),fPlanAmor.NumSeq, Infos);

      // fq 18261 mbo 31.05.2006 if ((DS.State = dsEdit) and (fDateAchatIni <> GetField('I_DATEPIECEA')))
      if ((DS.State = dsEdit) and (fDateOrigine <> GetField('I_DATEPIECEA')))
        then
      begin
        Infos.TVARecuperable := GetField('I_TVARECUPERABLE');
        Infos.TVARecuperee := GetField('I_TVARECUPEREE');
        UpdateLogAcquisition(GetField('I_IMMO'), GetField('I_DATEPIECEA'),
          Infos);
        ChargeLog;
      end;
    end;
  end
  // FQ 12503 Pb comptes associés => ouvrir l'onglet Comptable
  else
  begin
    if (iErreur = 21) and GetControlVisible('PComptable')  then
       SetActiveTabSheet('PComptable');
  end;
end;

procedure TOM_IMMO.OnAfterUpdateRecord;
begin
  inherited;
  { Enregistrement du contrat de crédit-bail }
  if (GetField('I_NATUREIMMO') = 'CB') or (GetField('I_NATUREIMMO') = 'LOC') then
  begin
    CreationImmoDepotGarantie;
    if fContrat.ListeEcheances.Count > 0 then
      fContrat.MajTableEcheances;
  end;
  if (GetField('I_NATUREIMMO') = 'FI') and (GetField('I_IMMOLIEGAR') <> '') then
  begin
    ExecuteSQL('UPDATE IMMO SET I_DEPOTGARANTIE=' +
      strfpoint(GetField('I_MONTANTHT'))
      + ' WHERE I_IMMO="' + GetField('I_IMMOLIEGAR') + '"');
  end;
  if (GetControlVisible('PContrat')) then
  begin
    SetControlEnabled('GroupeDepotGarantie', false);
    ImmoDepotGarantieExiste := (GetField('I_IMMOLIEGAR') <> '');
    GereDepotGarantie;
  end;
  bPlanModifie := false;
  UpdateRecordEncours:=false; // FQ 18341
  { si appel depuis la saisie, on renseigne le code de l'immo créée }
  if gTobImmo <> nil then
    gTobImmo.PutValue('I_IMMO', GetField('I_IMMO'));
  { Mise à jour des flags de modification }
  VHImmo^.ChargeOBImmo := True;
  ImMarquerPublifi(True);
  {$IFDEF EAGLCLIENT}
  // Dans ce cas, ça ferme la fiche tout seul !
  {$ELSE}
//  if (not (fLeMode = taCreatEnSerie)) and (fLeMode=taCreatOne ) then Close;
// la fermeture ne fonctionne pas ...
  {$ENDIF}
  if (fLeMode = taCreatEnSerie) then
  begin
    SetActiveTabSheet('PGeneral');
    CumFiscSav := 0;
    CumEcoSav := 0;
    // BTY 03/06 FQ 17596
    // Le clic sur le bouton Insert renvoie les évènts UpdateRecord et AfterUpdateRecord
    // => d'où violation d'accès dans les tables immolog et immoamor cause enregs doublonnés
    // => L'enregistrement de la fiche + ouverture nelle fiche se font automatiquement
    // si on a la propriété de la fiche TypeAction = taCreatEnSerie
    // TToolbarButton97(GetControl('BInsert')).Click;
    TFFiche(Ecran).Typeaction :=taCreatEnSerie;
    TFFiche(Ecran).Monofiche :=False;
    //
  end
  else
    TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.OnLoadRecord;
var
  QGar,Q : TQuery;
  NatureFiche: string;
  montant : double;
  TmpPlan: TPlanAmort;
begin

  fDisableChangeField := True;
  inherited;

  { Chargement du contrat }
  if fContrat <> nil then
    fContrat.RazEcheances;

  if (GetField('I_NATUREIMMO') = 'CB') or (GetField('I_NATUREIMMO') = 'LOC')
    then
  begin
    if fContrat = nil then
      fContrat := TImContrat.Create;
    fContrat.Charge(TQuery(DS));
    fContrat.ChargeTableEcheance;
    fContrat.CalculEcheances;
  end;
  { Chargement du plan d'unités d'oeuvre }
  if GetField('I_METHODEECO') = 'VAR' then
  begin
    if (fPlanUO = nil) then
      fPlanUO := TPlanUO.Create;
//    fPlanUO.Charge(GetField('I_IMMO'),GetField('I_DUREEECO'),GetField('I_UNITEOEUVRE'),GetField('I_DATEAMORT'));
  end;
  SetControlVisible('BUNITEOEUVRE', (GetField('I_METHODEECO') = 'VAR'));

  bChargementEnCours := true;

  if (TTabSheet(GetActiveTabSheet('Pages')).TabVisible = False) then
    SetActiveTabSheet('PGeneral');

  if (not (fLeMode in [taCreat..taCreatOne]) or (DS.State <> dsInsert)) then
    InitEnregExistant

  else
    begin
    fOperationEnCours := false;
    fDPIUtilisee := false;  // BTY 07/06 DPI
    end;

  AfficheMontantTTC;
  AfficheCoutContrat;

  if GetField('I_IMMOLIEGAR') <> '' then
  begin
    QGar := OpenSQL('SELECT I_COMPTEIMMO FROM IMMO WHERE I_IMMO="' +
      GetField('I_IMMOLIEGAR') + '"', false);
    if not QGar.Eof then
      SetControlText('COMPTE_DEPOT', QGar.FindField('I_COMPTEIMMO').AsString)
    else
      SetControlText('COMPTE_DEPOT', '');
    Ferme(QGar);
  end;

  if GetControlVisible('OpeEnCours') then // BTY 10/06 Bouton déjà caché depuis OnArgument
     SetControlVisible('OpeEnCours', IsOpeEnCours(nil, GetField('I_IMMO'), True));

  // TGA 29/08/2006 ============================================================
  // BTY 10/06 Complément
  If (fDateDebExoDPI = iDate1900) AND (GetField('I_DPI') = 'X') Then
    Begin
      // utilisation d'une DPI appelé par liste immo
      SetControlVisible('FE_DPI', True);
      SetControlVisible('bInfoDPI', True);  // stDPI // BTY 10/06
      montant:=0;
      Q :=OpenSQL('SELECT IZ_MONTANT FROM IMMOMVTD WHERE IZ_IMMO="'+
                  GetField('I_IMMO')+'" AND IZ_NATURED="IMM"', FALSE) ;
      try
        While Not Q.Eof do
          begin
            Montant := montant + Q.FindField('IZ_MONTANT').AsFloat;
            Q.Next ;
          End;
      finally
        Ferme(Q);
      End;
      SetControlText('MT_DPI', StrFMontant(montant,13,V_PGI.OkDecV,'',TRUE));
      SetControlVisible('MT_DPI', True);
    End
  ELSE
    // non utilisation d'une DPI
    Begin
      SetControlVisible('FE_DPI', False);
      SetControlVisible('MT_DPI', False);
      SetControlVisible('bInfoDPI', False); // stDPI // BTY 10/06
    End;
  // TGA 29/08/2006 ============================================================

  // BTY 10/06 Affichage infos Prime d'équipement ---------------------------
  if (not (fLeMode in [taCreat..taCreatOne]) and (DS.State <> dsInsert))
  and ((GetField('I_SBVPRI') <> 0) or (GetField('I_SBVPRIC') <> 0)) then
     begin
     TmpPlan := TPlanAmort.Create(true);
     try
       TmpPlan.Charge(TQuery(DS));
       SetControlVisible('stPrime', True);
       SetControlVisible('MtPrime', True);
       SetControlVisible('stBasePrime', True);
       SetControlVisible('MtBasePrime', True);
       if (GetField('I_CORVRCEDDE') <> 0) then
         begin
         SetControlVisible('stDureePrime', True);
         SetControlEnabled('DureePrime', False);
         SetControlVisible('DureePrime', True);
         SetControlText ('DureePrime', IntToStr (TmpPlan.PRIDuree) + ' mois');
         end
       else
         begin
         SetControlVisible('stDureePrime', False);
         SetControlVisible('DureePrime', False);
         end;
       SetControlText ('MtPrime', StrFMontant(TmpPlan.MNTPrime,13,V_PGI.OkDecV,'',TRUE));
       SetControlText ('MtBasePrime', StrFMontant((TmpPlan.MNTPrime)/2,13,V_PGI.OkDecV,'',TRUE));
     finally
       TmpPlan.free;
     end;
     end
  else
     begin
     SetControlVisible('StDureePrime', False);
     SetControlVisible('DureePrime', False);
     SetControlVisible('stPrime', False);
     SetControlVisible('MtPrime', False);
     SetControlVisible('stBasePrime', False);
     SetControlVisible('MtBasePrime', False);
     end;
  // -------------------------------------------------------------------------

  // BTY 10/06 Affichage infos Subvention d'investissement -------------------
  if (not (fLeMode in [taCreat..taCreatOne]) and (DS.State <> dsInsert))
  and ((GetField('I_SBVMT') <> 0) or (GetField('I_SBVMTC') <> 0)) then
  begin
     TmpPlan := TPlanAmort.Create(true);
     try
       TmpPlan.Charge(TQuery(DS));
       SetControlVisible('stSBV', True);
       SetControlVisible('MtSBV', True);
       SetControlVisible('stDateSBV', True);
       SetControlVisible('DateSBV', True);
       SetControlEnabled('DateSBV', False);
       SetControlVisible('stCompteSBV', True);
       SetControlVisible('CompteSBV', True);
       SetControlEnabled('CompteSBV', False);
       // Si durée affichée pour la prime d'équipement, ne pas afficher celle de la subvention
       if GetControlVisible('DureePrime') then
          begin
          SetControlVisible('stDureeSBV', False);
          SetControlVisible('DureeSBV', False);
          end
       else
          begin
          SetControlVisible('stDureeSBV', True);
          SetControlVisible('DureeSBV', True);
          SetControlEnabled('DureeSBV', False);
          SetControlText ('DureeSBV', IntToStr (TmpPlan.SBVDuree) + ' mois');
          if (GetField('I_METHODEECO') = 'NAM' ) then
            SetControlText ('stDureeSBV', 'Durée d''inaliénabilité')
          else
            SetControlText ('stDureeSBV', 'Durée d''échelonnement')
          end;
       SetControlText ('MtSBV', StrFMontant(TmpPlan.MNTSBV,13,V_PGI.OkDecV,'',TRUE));
       SetControlText ('DateSBV', DateToStr(TmpPlan.SBVDateSBV));
       SetControlText ('CompteSBV', GetField ('I_CPTSBVB'));
     finally
       TmpPlan.free;
     end;
  end else
  begin
       SetControlVisible('stSBV', False);
       SetControlVisible('MtSBV', False);
       SetControlVisible('stDateSBV', False);
       SetControlVisible('DateSBV', False);
       SetControlVisible('stDureeSBV', False);
       SetControlVisible('DureeSBV', False);
       SetControlVisible('stCompteSBV', False);
       SetControlVisible('CompteSBV', False);
  end;
  // -------------------------------------------------------------------------

  // BTY 11/07 Affichage infos Bien antérieur au passage au régime réel pour les agricoles
  SetControlVisible('stREGFR',
      (((fLeMode = taModif) or (fLeMode = taConsult)) and (GetField ('I_PFR') = 'X')));
  // -------------------------------------------------------------------------

  bChargementEnCours := false;

  if (GetField('I_NATUREIMMO') <> '') then
    NatureFiche := GetField('I_NATUREIMMO')
  else
    NatureFiche := fNature;

  SetControlEnabled('I_NATUREIMMO', GetControlEnabled('I_NATUREIMMO') and
    (fLeMode in [taCreat..taCreatOne]));

  if (not (fLeMode in [taCreat..taCreatOne]) and (DS.State <> dsInsert)) then
    TraiteOrigineImmo
  else
    SetControlVisible('AccesImmoLie', false);

  ChargeHistoriqueOperations; // CA le 01/02/1999
  fDateAchatIni := GetField('I_DATEPIECEA');
  fDateOrigine := fDateAchatIni; // fq 18261 mbo 31.05.2006
  GereSurAmortissement;
  fDisableChangeField := False;
  SetControlVisible ('BREGROUPEMENT',GetField('I_GROUPEIMMO')<>'');
  TCheckBox(GetControl('ELIGIBLETP')).OnClick := nil;
  if GetField('I_TABLE9')='X' then SetControlChecked('ELIGIBLETP',True);
  TCheckBox(GetControl('ELIGIBLETP')).OnClick := OnClickELIGIBLETP;
  {$IFDEF SERIE1}
  //XVI 25/04//2006 FD 3978 On ne le fait que si on a l'analytique... début
  if VS1Cpta.Y_Saisie then
  begin
     if (ds.State=dsBrowse) then
        S1_VerifieVentilation(FALSE) ; //XVI 13/09/2006 FQ 12952
     if not VS1Cpta.Y_AnalytiqueAdvanced then
        SetControltext('SECTIONS1',GetField('I_TABLE9')) ;
  end ;
  //XVI 25/04/2006 FD 3978 fin
  {$ELSE}
  // BTY 09/06 FQ 16149
  if DS.State=dsInsert then
     begin
     FComboEtab := nil;
     GetComboEtab;
     GereEtablissement;
     end;
  //  if DS.State=dsInsert then PositionneEtabUser(THValComboBox(GetControl('I_ETABLISSEMENT')));
  //  SetControlEnabled('I_ETABLISSEMENT',not VH^.ProfilUserC[prEtablissement].ForceEtab);
  {$ENDIF}

  // TGA 09/02/2006
  if (GetField('I_REPRISEDEP') <> 0) then
  begin
     SetControlChecked('bCalculAvecVnc', true);
     SetControlEnabled('bCalculAvecVnc', false);

     GereCocheVNF  // chantier fiscal  fq 17512
  end;

  // ajout mbo 29.06.07 + ajout test sur i_string1 pour après cloture 10.07.07
  // FQ 21502 Items.Delete met du désordre dans les valeurs
  if (GetField('I_STRING1')='') then
  begin
  //   THValComboBox(GetControl('TYPEBASEFISC')).Items.Delete(THValComboBox(GetControl('TYPEBASEFISC')).Items.IndexOf(RechDom('TITYPEBASEFISC', 'RC2', FALSE)));
  //   THValComboBox(GetControl('TYPEBASEFISC')).Items.Delete(THValComboBox(GetControl('TYPEBASEFISC')).Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE)));
     SetControlProperty('TYPEBASEFISC','Plus',' AND CO_CODE<>"TR2"');
  end;
end; //MVG !!!  29/03/07

{$IFDEF SERIE1}
procedure TOM_IMMO.TABLELIBRE1Change(Sender: TObject);
var
  i: integer ;
  st1, st2: string ;
begin
i:=1 ;
if Sender=GetControl('TABLELIBRE1') then i:=1
else if Sender=GetControl('TABLELIBRE2') then i:=2
else if Sender=GetControl('TABLELIBRE3') then i:=3 ;
st1:=GetField('I_TABLE'+intToStr(i-1)) ;
st2:=GetControlText('TABLELIBRE'+intToStr(i)) ;
if (i>0) and (st1 <> st2) then
  begin
  if DS.State in [dsBrowse] then DS.edit;
  SetField('I_TABLE'+intToStr(i-1), st2);
  end ;
end;
{$ELSE}
{$ENDIF SERIE1}

procedure TOM_IMMO.OnChangeField(F: TField);
var
  stCompte: string;
  stMsg: string;
  // fq 18430 dMinimumLineaire: double;
  ok:boolean;
  {$IFDEF SERIE1}
  idx : integer ; //XVI 18/04/2006 FD 3978
  {$ELSE}
  {$ENDIF !SERIE1}
var
{$IFDEF SERIE1}  //YCP 25/08/05
{$ELSE}
  wName, wNum, wVal: string;
{$ENDIF SERIE1}
  stNomChamp: string;
begin
  if DS.state in [dsBrowse] then
    exit;
  if fDisableChangeField then
    exit;
  fDisableChangeField := True;
  try
//    Append(Ffichier);
//    Writeln(fFichier, F.FieldName);
//    CloseFile(fFichier);
    inherited;
    stNomChamp := F.FieldName;

    //--------------------------------------------------------------------------
    if stNomChamp = 'I_IMMO' then
    begin
      SetField('I_ABREGE', Copy(GetField('I_IMMO'), 1, 17));
      if fPlanUO <> nil then fPlanUO.SetCodeImmo ( GetField('I_IMMO'));
    end
    //--------------------------------------------------------------------------
    else if stNomChamp = 'I_MONTANTHT' then
    begin
      bPlanModifie := true;
      if DS.state in [dsEdit, dsInsert] then
      begin
        AfficheMontantTTC;
        AfficheBaseTaxe;
        AfficheTauxEco;
        AfficheTauxFisc;
        // fq 17569 GereReprise;
        GereRepriseEco;
        if bPlanFiscal then GereRepriseFisc;

        TraiteDotationAnterieures;
        SetField('I_REPRISEFISCAL', CumFiscSav);
        SetField('I_REPRISEECO', CumEcoSav);

        { chantier fiscal  fq 17512
        // ajout pour fq 15274 - mbo
        if (CumFiscSav < CumEcoSav) and (GetField('I_METHODEECO') = 'VAR') then
           SetControlProperty('AmortDerog','Value', 0)
        else
           SetControlProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);
        }
        AfficheDerogatoire;

        SetField('I_VALEURACHAT', GetField('I_MONTANTHT'));
      end;
    end
    //--------------------------------------------------------------------------
    else if ((stNomChamp = 'I_DATEFINECH') or (stNomChamp = 'I_DATEDEBUTECH') or
      (stNomChamp = 'I_MONTANTPREMECHE')) then
    begin
      if ((stNomChamp = 'I_DATEFINECH') or (stNomChamp = 'I_DATEDEBUTECH')) then
      begin
        if GetField('I_DATEDEBUTECH') < GetField('I_DATEPIECEA') then
          PGIBox(MSG_ECHPREMLIV);
        if GetField('I_DATEFINECH') < GetField('I_DATEDEBUTECH') then
        begin
          PGIBox(MSG_ECHDERPREM);
          SetField('I_DATEFINECH', GetField('I_DATEDEBUTECH'));
          SetFocusControl('I_DATEFINECH');
          exit;
        end;
        OnChangementContrat(stNomChamp);
      end
      else if (stNomChamp = 'I_MONTANTPREMECHE') then
      begin
        if (GetField('I_MONTANTPREMECHE') <> 0.0) and
          (GetField('I_MONTANTSUIVECHE') = 0.0) then
          SetField('I_MONTANTSUIVECHE', GetField('I_MONTANTPREMECHE'));
      end;
      if (fContrat <> nil) then
      begin
        fContrat.Charge(TQuery(DS));
        fContrat.CalculEcheances;
        AfficheCoutContrat;
        AfficheNombreEcheance;
        EnableBoutonsInfos;
        OnChangementContrat('I_MONTANTPREMECHE');
      end;
    end
    //--------------------------------------------------------------------------
    else if stNomChamp = 'I_DATEPIECEA' then
    begin
      bPlanModifie := true;
      if (GetField('I_DATEPIECEA') < VHImmo^.Encours.Deb) then // FQ 12514
      begin
        if PGIAsk(ASK_DATEACQDEBEX) <> mrYes then
        begin
          // ajout fq 17569
          if (GetField('I_METHODEECO') <> '') and (fDateAchatIni <> GetField('I_DATEPIECEA')) then
             PGIInfo(MSG_MODIFACHAT);

          SetFocusControl('I_DATEPIECEA');
          exit;
        end;
      end;

      // ajout fq 17569
      if (GetField('I_METHODEECO') <> '') and (fDateAchatIni <> GetField('I_DATEPIECEA')) then
             PGIInfo(MSG_MODIFACHAT);

      SetField('I_DATEAMORT', GetField('I_DATEPIECEA'));
      fDateMSIni := GetField('I_DATEPIECEA');  // pour éviter d'avoir le message de modif de date
      fDateAchatIni := GetField('I_DATEPIECEA');

      // fq 17569
      if GetField('I_METHODEECO') = 'DEG' then
         SetField('I_DATEDEBECO',GetField('I_DATEPIECEA'))
      else
         SetField('I_DATEDEBECO',GetField('I_DATEAMORT'));

      If GetField('I_METHODEFISC')='DEG' Then
         SetField('I_DATEDEBFIS',GetField('I_DATEPIECEA'))
      Else
         SetField('I_DATEDEBFIS',GetField('I_DATEAMORT'));
        // fin fq 17569

      // 17569 if GetField('I_METHODEECO')='VAR' then fPlanUO.DateDebut := GetField('I_DATEPIECEA');
      if GetField('I_METHODEECO')='VAR' then fPlanUO.DateDebut := GetField('I_DATEAMORT');

      if (GetField('I_NATUREIMMO') = 'CB') then
        SetField('I_DATEDEBUTECH', GetField('I_DATEPIECEA'));

      AfficheTauxEco;
      AfficheTauxFisc;

      GereRepriseEco;  // fq 17569 GereReprise;
      if bPlanFiscal then GereRepriseFisc;

      TraiteDotationAnterieures;
      SetField('I_REPRISEFISCAL', CumFiscSav);
      SetField('I_REPRISEECO', CumEcoSav);

      // raz de la reprise dépréciation si date saisie dans l'exercice
      // modif fq 17569
      //if (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb) then SetField('I_REPRISEDEP', 0);
      if (GetField('I_DATEDEBECO') >= VHImmo^.Encours.Deb) then SetField('I_REPRISEDEP', 0);


      // ajout mbo - fq 15274
      {if (GetField('I_METHODEECO') = 'VAR') and (CumFiscSav < CumEcoSav) then
          SetControlProperty('AmortDerog', 'Value',0)
      else
         SetControlProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);}

      AfficheDerogatoire; // chantier fiscal  fq 17512

      GereSuramortissement;
    end
    //--------------------------------------------------------------------------
    else if EstUnChampCompte(stNomChamp) then
    begin
      if ((stNomChamp = 'I_COMPTEIMMO') and (GetField('I_COMPTEIMMO') = '')) then
          RazComptesAssocies;
      if (stNomChamp <> 'I_ORGANISMECB') then
      begin
        if (GetField(stNomChamp)<>'') then
        begin
          stCompte := ImBourreEtLess(GetField(stNomChamp), ImGeneTofb);
          if Presence('GENERAUX', 'G_GENERAL', stCompte) then
          begin
            SetField(stNomChamp, stCompte);
            if ((stNomChamp = 'I_COMPTEIMMO') and (GetField('I_NATUREIMMO') =
              'PRO')) then
            begin
              SetControlText('COMPTEIMMO',GetField('I_COMPTEIMMO'));
              MajComptesAssocies;
              {$IFDEF SERIE1}
              if ds.state<>dsBrowse then
                 setfield('I_VENTILABLE',Hent1.BoolToStr_(existeSQL('select G_GENERAL from GENERAUX where G_GENERAL="'+stCompte+'" and G_VENTILABLE="X"'))) ; //XVI 18/04/2006 FD 3978
              {$ELSE}
              {$ENDIF !SERIE1}
            end;
          { FQ 15753 - Ne pas ouvrir le choix des comptes si on n'est pas en saisie }
          end else (*if (GetActiveTabSheet('Pages')= TTabSheet(GetControl('PPlan'))) then*)
          begin
            OnCompteElipsisClick(GetControl(stNomChamp));
            if not Presence('GENERAUX', 'G_GENERAL', GetField(stNomChamp)) then
            begin
              SetFocusControl(stNomChamp);
              fDisableChangeField := False;
              stCompte := GetField(stNomChamp);
              stCompte := '' ; // ajout mbo 03.04.2006 fq 17704
              SetField(stNomChamp,'');
              NextPrevControl(Ecran);
              SetField(stNomChamp,stCompte);
            end else
            begin
              if ((stNomChamp = 'I_COMPTEIMMO') and (GetField('I_NATUREIMMO') =
                'PRO')) then
              begin
                SetControlText('COMPTEIMMO',GetField('I_COMPTEIMMO'));
                MajComptesAssocies;
              end;
            end;
          end;
        end;
      end;

      { FQ 12423 - CA - 02/07/2003 - si on modifie le compte, il faut réengistrer le plan pour répercuter la
          modification du compte sur la table IMMOAMOR }
      if stNomChamp = 'I_COMPTEIMMO' then
        { FQ 13234 - MBO - 31/08/2005 - si on modifie un compte d'immo corporelle en
        compte d'immo incorporelle ---> base taxe pro est remise à 0}
        afficheBaseTaxe;
        //
        bPlanModifie := True;
    end
    //--------------------------------------------------------------------------------------
    else if stNomChamp = 'I_DATEAMORT' then
    begin
      bPlanModifie := true;
      if DS.state in [dsEdit, dsInsert] then
      begin
        // FQ 17569
        if (GetField('I_METHODEECO')<> '') and (fDateMSini <> GetField('I_DATEAMORT')) then
             PGIInfo(MSG_MODIFSERVICE);

        fDateMSini := GetField('I_DATEAMORT');
        if (GetField('I_METHODEECO') <> 'DEG') then
           SetField('I_DATEDEBECO', Getfield('I_DATEAMORT'))
        else
           SetField('I_DATEDEBECO', Getfield('I_DATEPIECEA'));

        if bPlanFiscal then
        begin
          If GetField('I_METHODEFISC')='DEG' Then
             SetField('I_DATEDEFIS',GetField('I_DATEPIECEA'))
          Else
             SetField('I_DATEDEBFIS',GetField('I_DATEAMORT'));
        end;

        {If (GetField('I_METHODEECO')='VAR') and
           not (GetCheckBoxState('bGestionFiscale') = cbChecked) Then  //ajout pour chantier fiscal
        begin
           SetControlEnabled('I_DATEDEBFIS', False);
        end;}

        if GetField('I_METHODEECO') = 'NAM' then
           SetControlEnabled('I_DATEDEBECO', false);
        // fin fq 17569

        AfficheTauxEco;
        AfficheTauxFisc;
        GereRepriseEco;
        If bPlanFiscal then GereRepriseFisc;
        TraiteDotationAnterieures;
        SetField('I_REPRISEFISCAL', CumFiscSav);
        SetField('I_REPRISEECO', CumEcoSav);

        {chantier fiscal  fq 17512
        // ajout pour fq 15274 - mbo
        if (CumFiscSav < CumEcoSav) and (GetField('I_METHODEECO') = 'VAR') then
           SetControlProperty('AmortDerog','Value', 0)
        else
           SetControlProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);
        }
        AfficheDerogatoire;

        // Ajout fq 17569
        if (GetField('I_DATEDEBECO') >= VHImmo^.Encours.Deb) then SetField('I_REPRISEDEP', 0);
        // ajout fq 17569

        if (GetField('I_METHODEECO')='VAR') then
        begin
           fPlanUO.DateDebut := GetField('I_DATEDEBECO');
           OnClickUniteOeuvre(GetControl('I_DATEAMORT'));
       end;
      end;
    end
    //--------------------------------------------------------------------------------
    else if stNomChamp = 'I_TVARECUPERABLE' then
    begin
      if DS.state in [dsEdit, dsInsert] then
      begin
        SetField('I_TVARECUPEREE', GetField('I_TVARECUPERABLE'));
        AfficheMontantTTC;
        AfficheBaseTaxe;
        OnChangementContrat('I_TVARECUPERABLE');
      end;
      bPlanModifie := true;
    end
    //--------------------------------------------------------------------------------
    else if stNomChamp = 'I_TVARECUPEREE' then
    begin
      if (DS.state in [dsEdit, dsInsert]) and
        (GetField('I_TVARECUPEREE') > GetField('I_TVARECUPERABLE')) then
      begin
        PGIBox(HM2[10]);
        SetField('I_TVARECUPEREE', GetField('I_TVARECUPERABLE'));
      end;
      OnChangementContrat('I_TVARECUPEREE');
      AfficheBaseTaxe;
      bPlanModifie := true;
    end
    //--------------------------------------------------------------------------------
    else if stNomChamp = 'I_METHODEECO' then
    begin
      SetControlVisible('PComptable', (GetField('I_NATUREIMMO') = 'PRO') and
        (not (GetField('I_METHODEECO') = 'NAM')));

      SetControlEnabled('bAmortissementFiscal', ( (GetField('I_METHODEECO') <>'NAM')
                                           and    (GetField('I_NATUREIMMO') = 'PRO')
                                           and    (GetField('I_METHODEECO') <> 'VAR')));

      SetControlVisible('BUNITEOEUVRE', (GetField('I_METHODEECO') = 'VAR')
                                    and (GetField('I_UNITEOEUVRE')>0));

      //chantier fiscal
      SetControlChecked('bGestionFiscale', false);

      // modif mbo 25.07.2005 FQ 16101
      //SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0) and (GetField('I_METHODEECO')='DEG'));
      // mbo 09.07.07SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0));

      // Tga 18/07/2005 correction fiche 15829
      // mbo 25.07.2005  FQ 16101 SetControlEnabled('ELIGIBLETP',(GetField('I_METHODEECO')= 'DEG'));

      if (GetField('I_METHODEECO') = 'NAM') then
         GereNonAmortissable;

      // FQ 19013
      EnableBoutonBaseEco;

      if (GetField('I_METHODEECO')='VAR') then
      begin
        if fPlanUO = nil then
          fPlanUO := TPlanUO.Create;
        { Si amortissement variable ==>
                 on propose par défaut un plan fiscal en linéaire }

        SetControlChecked('bAmortissementFiscal', True);
        SetField('I_METHODEFISC','LIN');
        SetField('I_DUREEFISC',GetField('I_DUREEECO'));
        SetControlText('TYPEBASEFISC','ECO'); //PGR - 09/2005 - Valeur théorique sur Base amort fiscal
        SetField('I_BASEFISC',GetField('I_BASEECO'));
        fPlanUO.DateDebut := GetField('I_DATEDEBECO'); // fq 17569 GetField('I_DATEPIECEA');
        fPlanUO.Duree := GetField('I_DUREEECO');
        SetControlEnabled('bGestionfiscale', true);
        SetControlEnabled('AmortFisc', true);
      end else
      begin
        { Si on passe de variable à un autre type d'amortissement on supprime le fiscal }
        if (GetField('I_UNITEOEUVRE')<>0) then
        begin
          SetControlChecked('bAmortissementFiscal',False);
          //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
          SetControlText('TYPEBASEFISC', '');

          SetField('I_METHODEFISC','');
          SetField('I_DUREEFISC', 1); //PGR -03/10/2005 1 au lieu de 0 pour éviter messages intempestifs
          SetField('I_BASEFISC',0);

          SetControlChecked('bGestionFiscale', false); // Chantier fiscal
          SetControlChecked('bCalculavecVNF', false);
        end;

        SetField('I_UNITEOEUVRE',0);
        SetField('I_INFOUO','');
        ExecuteSQL('DELETE FROM IMMOUO WHERE IUO_IMMO="'+GetField('I_IMMO')+'"');


      if (GetField('I_METHODEECO') = 'DEG') then
         SetField('I_DATEDEBECO', GetField('I_DATEPIECEA'))
      else
         SetField('I_DATEDEBECO', GetField('I_DATEAMORT'))
      end;
      GereEco;
      GereFiscal;
      EnableBoutonsInfos;
      OnChangementMethode('I_METHODEECO');
    end
    //--------------------------------------------------------------------------------
    else if stNomChamp = 'I_METHODEFISC' then
    begin
      // chantier fiscal
      GereCocheVNF;

      GereEco;
      GereFiscal;
      GereRepriseEco;
      GereRepriseFisc;
      //fq 17941 - mbo - 20.04.06  déplacement du onchangementmethode
      OnChangementMethode('I_METHODEFISC');

      if (not (fLeMode in [taCreat..taCreatOne]) and (DS.State <> dsInsert)) then
        begin
          RecuperePlanRepriseTheorique(TQuery(DS), CumEcoSav, CumFiscSav);
          // fq 17941 - OnChangementMethode('I_METHODEFISC');

        end;
      // Tga 18/07/2005 correction fiche 15829
      // mbo 25/07/2005 FQ 16101 SetControlEnabled('ELIGIBLETP',(GetField('I_METHODEFISC')='DEG'));
      SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0));

    end
 //--------------------------------------------------------------------------------
    else if stNomChamp = 'I_DUREEECO' then
    begin
      if GetField('I_DUREEECO') = 0 then
      begin
        SetField('I_METHODEECO', 'NAM');
      end;

      if (GetField('I_METHODEECO')='VAR') then
      begin
        fPlanUO.Duree := GetField('I_DUREEECO');

        //if not(GetCheckBoxState('bGestionFiscale') = cbChecked) then // ajout du controle chantier fiscal
        //FQ 21423 - mbo - 17/09/07 - on réaligne durée fiscale sur la durée éco avec possibilité de la modifier
        SetField('I_DUREEFISC',GetField('I_DUREEECO'));

        if (GetField('I_UNITEOEUVRE')>0) then OnClickUniteOeuvre(GetControl('I_DUREEECO'));
      end;

      if not MinDureeAmortissement(GetField('I_METHODEECO'),
        GetField('I_DUREEECO'), stMsg) then
      begin
        PGIBox(stMsg);
        SetActiveTabSheet('PPlan');
        SetFocusControl('I_DUREEECO');
      end else
      begin
         OnChangementDureeAmort('I_DUREEECO', True);
      end;
      //OnChangementDureeAmort('I_DUREEECO', True);
    end
  //-----------------------------------------------------------------------------------
    else if stNomChamp = 'I_DUREEFISC' then
    begin
      //PGR -03/10/2005 Ctl si amortissement fiscal uniquement, pour éviter des messages intempestifs
      if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
      begin
        if not MinDureeAmortissement(GetField('I_METHODEFISC'),
          GetField('I_DUREEFISC'), stMsg) then
        begin
          PGIBox(stMsg);
          SetActiveTabSheet('PPlan');
          
          SetFocusControl('I_DUREEFISC');
        end else
        begin
          AfficheTauxFisc;
          OnChangementDureeAmort('I_DUREEFISC', True);

        end;
        //AfficheTauxFisc;
        //OnChangementDureeAmort('I_DUREEFISC', True);
      end;
    end
  //-----------------------------------------------------------------------------------
    // ajout fq 17569
    else if stNomChamp = 'I_DATEDEBECO' then
    begin
      bPlanModifie := true;
      if DS.state in [dsEdit, dsInsert] then
      begin
        if not (DateDebValide('ECO',true)) then
        begin
          SetActiveTabSheet('PPlan');
          SetFocusControl('I_DATEDEBECO');
          exit;
        end;
      end;

      AfficheTauxEco;
      AfficheTauxFisc;
      GereRepriseEco;
      if bPlanFiscal then GereRepriseFisc;
      TraiteDotationAnterieures;
      SetField('I_REPRISEFISCAL', CumFiscSav);
      SetField('I_REPRISEECO', CumEcoSav);

      {chantier fiscal fq 17512
      // ajout pour fq 15274 - mbo
      if (CumFiscSav < CumEcoSav) and (GetField('I_METHODEECO') = 'VAR') then
          SetControlProperty('AmortDerog','Value', 0)
      else
         SetControlProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);
      }
      AfficheDerogatoire;

      // Ajout fq 17569
      if (GetField('I_DATEDEBECO') >= VHImmo^.Encours.Deb) then SetField('I_REPRISEDEP', 0);
      // ajout fq 17569

      if (GetField('I_METHODEECO')='VAR') then
      begin
         fPlanUO.DateDebut := GetField('I_DATEDEBECO');
         // ajout 13.06.2006 suite fq 17569
         if (GetField('I_UNITEOEUVRE')>0) then
            OnClickUniteOeuvre(GetControl('I_DATEAMORT'));
      end;
    end
  //-----------------------------------------------------------------------------------
    // ajout fq 17569
    else if stNomChamp = 'I_DATEDEBFIS' then
    begin
      if bPlanFiscal then
      begin
         bPlanModifie := true;
         if DS.state in [dsEdit, dsInsert] then
         begin
           if not (DateDebValide('FISC',true)) then
           begin
             SetActiveTabSheet('PPlan');
             SetFocusControl('I_DATEDEBFIS');
             exit;
           end;
         end;

         AfficheTauxEco;
         AfficheTauxFisc;
         GereRepriseEco;
         GereRepriseFisc;
         TraiteDotationAnterieures;
         SetField('I_REPRISEFISCAL', CumFiscSav);
         SetField('I_REPRISEECO', CumEcoSav);
         {chantier fiscal  fq 17512
         // ajout pour fq 15274 - mbo
         if (CumFiscSav < CumEcoSav) and (GetField('I_METHODEECO') = 'VAR') then
            SetControlProperty('AmortDerog','Value', 0)
         else
            SetControlProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);
         }
         AfficheDerogatoire;
       end;
    end
//-----------------------------------------------------------------------------
    else if (stNomChamp = 'I_UNITEOEUVRE') then
    begin
      if (GetField('I_METHODEECO')='VAR') then
      begin
        // ajout pour correction anomalie 01.03.07
        if (fPlanUO.TotalUO <> GetField('I_UNITEOEUVRE')) then
        begin
           fPlanUO.TotalUo := GetField('I_UNITEOEUVRE');
           fPlanUO.DateDebut := GetField('I_DATEDEBECO');  // pour forcer un recalcul
           OnClickUniteOeuvre(GetControl('I_UNITEOEUVRE'));
        end;

        fPlanUO.TotalUO := GetField('I_UNITEOEUVRE');
        OnChangementDureeAmort('I_UNITEOEUVRE', True);
        SetControlVisible('BUNITEOEUVRE', (GetField('I_METHODEECO') = 'VAR') and (GetField('I_UNITEOEUVRE')>0));

        // ajout mbo chantier fiscal fq 17512
        bPlanModifie := true;
        TraiteDotationAnterieures;
        AfficheDerogatoire;
      end;
    end
 //--------------------------------------------------------------------------------------------
    else if ((stNomChamp = 'I_BASEECO') or (stNomChamp = 'I_BASEFISC')) then
    begin
      if GetField('I_BASEECO') > (GetField('I_MONTANTHT') +
        GetField('I_TVARECUPERABLE') - GetfIeld('I_TVARECUPEREE')) then
      begin
        PGIBox(MSG_ERREURBASE);
        SetActiveTabSheet('PPlan');
        // BTY 08/06 Saisie par le bouton sBaseEcoOuDPI => I_BASEECO bloqué
        // FQ 19013
        if GetControlVisible('sBAseEcoOuDPI') and GetControlEnabled('sBAseEcoOuDPI') then
           SetFocusControl('sBAseEcoOuDPI')
        else //
           SetFocusControl('I_BASEECO');
        exit;
      end;

      bPlanModifie := true;
      if DS.state in [dsEdit, dsInsert] then
      begin
        if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
          begin
            //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
            AfficheBaseFisc;
          end;

        AfficheTauxEco;
        TraiteDotationAnterieures;

        if stNomChamp = 'I_BASEFISC' then
        begin
          SetField('I_REPRISEFISCAL', CumFiscSav);
         //PGR 4/11/2005 - Pb affichage avec ou sans décimale
         // SetControlText('AmortDerog', GetField('I_REPRISEFISCAL') -
         //   GetField('I_REPRISEECO'));

         {chantier fiscal fq 17512
         // ajout pour fq 15274 - mbo
          if (GetField('I_METHODEECO') = 'VAR') and
             (GetField('I_REPRISEFISCAL')< GetField('I_REPRISEECO')) then
              SetControlProperty('AmortDerog','Value', 0)
          else
              SetControlText('AmortDerog', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
              Getfield('I_REPRISEECO'), 13, V_PGI.OkDecV, '', TRUE));
          }
          AfficheDerogatoire;
        end
        else if stNomChamp = 'I_BASEECO' then
        begin
          SetField('I_REPRISEECO', CumEcoSav);
          //PGR 4/11/2005 - Pb affichage avec ou sans décimale
          //SetControlText('AmortDerog', GetfIeld('I_REPRISEFISCAL') -
          //  GetField('I_REPRISEECO'));

          {chantier fiscal  fq 17512
          // ajout pour fq 15274 - mbo
          if (GetField('I_METHODEECO') = 'VAR') and
             (GetField('I_REPRISEFISCAL')< GetField('I_REPRISEECO')) then
              SetControlProperty('AmortDerog','Value', 0)
          else
              SetControlText('AmortDerog', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
              Getfield('I_REPRISEECO'), 13, V_PGI.OkDecV, '', TRUE));
          }
          AfficheDerogatoire;
        end;
      end;
    end
  //-------------------------------------------------------------------------------
    else if stNomChamp = 'I_TYPELOYERCB' then
    begin
      if DS.state in [dsEdit, dsInsert] then
      begin
        GereGroupeEcheance;
        SetControlEnabled('SaisieEcheancier', not
          GetControlEnabled('GroupeEcheance'));
        if (GetField('I_TYPELOYERCB') <> 'LCO') then
        begin
          SetField('I_MONTANTPREMECHE', 0);
          SetField('I_MONTANTSUIVECHE', 0);
          SetField('I_FRAISECHE', 0);
          SetField('I_DATEDEBUTECH', GetField('I_DATEPIECEA'));
          SetField('I_DATEFINECH', GetField('I_DATEPIECEA'));
        end;
        SetControlEnabled('I_DATEDEBUTECH', (GetField('I_TYPELOYERCB') =
          'LCO'));
        SetControlEnabled('I_DATEFINECH', (GetField('I_TYPELOYERCB') = 'LCO'));
        SetControlEnabled('NbEcheances', (GetField('I_TYPELOYERCB') = 'LCO'));
        if (fContrat <> nil) then
        begin
          fContrat.RazEcheances;
          fContrat.Charge(TQuery(DS));
          fContrat.CalculEcheances;
        end;
      end;
    end
    else if ((stNomChamp = 'I_RESIDUEL') or (stNomChamp = 'I_DATESUSPCB') or
      (stNomChamp = 'I_MONTANTPREMECHE')
      or (stNomChamp = 'I_MONTANTSUIVECHE') or (stNomChamp = 'I_FRAISECHE') or
        (stNomChamp = 'I_PERIODICITE')
      or (stNomChamp = 'I_VERSEMENTCB')) then
    begin
      //mbo - fq 17362 -
      {if stNomChamp = 'I_DATESUSPCB' then
      begin
        if bChargementEnCours then
          exit;
        if DS.State in [dsBrowse] then
          DS.Edit;
        SetCOntrolEnabled('I_DATESUSPCB', GetCheckBoxState('cbI_DATESUSPCB') =
          cbChecked);
        if not (GetCheckBoxState('cbI_DATESUSPCB') = cbChecked) then
          SetField('I_DATESUSPCB', iDate1900)
        else
          SetField('I_DATESUSPCB', StrToDate(GetControlText('dtFinContrat')));
      end; }
      OnChangementContrat(stNomChamp);
    end
 //-------------------------------------------------------------------------------------------------
    else if (stNomChamp = 'I_NATUREIMMO') then
    begin
      fNature := GetField('I_NATUREIMMO');
      SetControlEnabled('I_COMPTELIE', GetField('I_NATUREIMMO') = 'CB');
      SetControlEnabled('tI_COMPTELIE', GetField('I_NATUREIMMO') = 'CB');
      if not GetControlEnabled('I_COMPTELIE') then
        SetField('I_COMPTELIE', '');
      SetControlEnabled('I_COMPTEIMMO', true);
      SetControlEnabled('tI_COMPTEIMMO', true);
      if DS.State = dsInsert then
        EnableChamps;
      if ((GetField('I_NATUREIMMO')='CB') or (GetField('I_NATUREIMMO')='LOC')) then
        if fContrat = nil then fContrat := TImContrat.Create;
    end
  //---------------------------------------------------------------------------------------------
    else if (stNomChamp = 'I_REPRISEDEP') then // mbo fq 17500 OR (stNomChamp = 'I_REPRISEFISCAL')Then
      // TGA 30/01/2006 test de la reprise dépréciation
      Begin
        Ok := Test_Reprise_Depreciation(stNomChamp);

        IF ok=False Then
        Begin
//         SetFocusControl(stNomChamp);
           Ecran.ModalResult := mrNone;
           exit;
        end else
        begin
          if GetField('I_REPRISEDEP')<>0 Then
          begin
             SetControlChecked('bCalculAvecVnc', true);
             SetControlEnabled('bCalculAvecVnc', false);
          end else
          begin
             SetControlChecked('bCalculAvecVnc', false);
             SetControlEnabled('bCalculAvecVnc', true);
             // ajout pour fq 17512 calcul sur vn fiscale
             if (GetCheckBoxState('bCalculAvecVnf') = cbChecked) then
                 SetControlChecked('bCalculAvecVnf', false);
          end;

          GereCocheVNF;  //ajout chantier fiscal  fq 17512
           // ajout mbo 10.02.06 pour recalculer le plan d'amortissement

          AfficheDerogatoire;
          bPlanModifie := true;
        end;
      end
 //--------------------------------------------------------------------------------
    else if ((stNomChamp = 'I_REPRISEECO') or (stNomChamp = 'I_REPRISEFISCAL'))then
    begin
      if (DS.state in [dsInsert, dsEdit]) and (GetField(stNomChamp) = 0) then
      begin
          GereRepriseEco;
          if bPlanFiscal then GereRepriseFisc;
          TraiteDotationAnterieures;
          if (stNomChamp = 'I_REPRISEFISCAL') then
          begin
             SetField('I_REPRISEFISCAL', CumFiscSav);
             // ajout pour fq 15274  + modif chantier fiscal fq 17512
             AfficheDerogatoire;
          end
          else if (stNomChamp = 'I_REPRISEECO') then
          begin
            SetField('I_REPRISEECO', CumEcoSav);

            // ajout pour fq 15274 + chantier fiscal fq 17512
            AfficheDerogatoire;
          end;
      end;

      // ajout mbo 19.06.06 fq 17500
      Ok := Test_Reprise_Depreciation(stNomChamp);

      IF ok=False Then
      Begin
//           SetFocusControl(stNomChamp);
         Ecran.ModalResult := mrNone;
         exit;
      end;

      // BTY 07/05 FQ 16171 Test sur la base amortissement et non le montant HT
      // if (GetField(stNomChamp) > GetField('I_MONTANTHT')) then
      if (GetField(stNomChamp) > GetField('I_BASEECO')) then
      begin
        PGIBox(HMC[29]);
        SetFocusControl(stNomChamp);
        Ecran.ModalResult := mrNone;
        exit;
      end;
      bPlanModifie := true;
      // 18430 mettre en commentaire le pavé
      {dMinimumLineaire := GetMinimumLineaire(TQuery(DS));
      if ((GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
        (GetField('I_REPRISEFISCAL') < dMinimumLineaire))
        or ((not (GetCheckBoxState('bAmortissementFiscal') = cbChecked)) and
          (GetField('I_REPRISEECO') < dMinimumLineaire)) then
      begin
        if PGIAsk(HMC[30]) = mrYes then
        begin
          SetActiveTabSheet('PPlan');
          if GetCheckBoxState('bAmortissementFiscal') = cbChecked then
            SetFocusControl('I_REPRISEFISCAL')
          else
            SetFocusControl('I_REPRISEECO');
          Ecran.ModalResult := mrNone;
          exit;
        end;
        TraiteDotationAnterieures;
      end;  }

      // ajout pour fq 15274 + chantier fiscal fq 17512
      //AfficheDerogatoire;
      EnableBoutonsInfos;
    end
 //---------------------------------------------------------------------------------------
    else if (stNomChamp = 'I_BASETAXEPRO') then
    begin
      GereNatureBien;
      // MBO 25.07.05 FQ 16101 SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0) and (GetField('I_METHODEECO')='DEG'));
      SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0));
      if ((GetField('I_BASETAXEPRO')=0) and (GetCheckBoxState('ELIGIBLETP')=cbChecked)) then
      begin
        SetControlChecked ('ELIGIBLETP',False);
        SetField('I_TABLE9','-')
      end;
    end
  //---------------------------------------------------------------------------------
    else if (stNomChamp = 'I_REINTEGRATION') then
    begin
      bPlanModifie := true;
    end
  //-------------------------------------------------------------------------------
    else if (stNomChamp = 'I_QUOTEPART') then
    begin
      // BTY 12/05 FQ 17156 bPlanModifie=True <=> ajout plan d'amortissement
      // Or la modif de quotepart n'impacte que IMMO et pas IMMOAMOR
      //(Zone IA_QUOTEPART de IA_IMMOAMOR inutilisée)
      //bPlanModifie := true;
    end
  //-------------------------------------------------------------------------------
    else if stNomChamp = 'TABLELIBRE1' then
    begin
      if GetField('I_TABLE0') <> GetControlText('TABLELIBRE1') then
      begin
        if DS.State in [dsBrowse] then
          DS.edit;
        SetField('I_TABLE0', GetControlText('TABLELIBRE1'));
      end;
    end
    else if stNomChamp = 'TABLELIBRE2' then
    begin
      if GetField('I_TABLE1') <> GetControlText('TABLELIBRE2') then
      begin
        if DS.State in [dsBrowse] then
          DS.edit;
        SetField('I_TABLE1', GetControlText('TABLELIBRE2'));
      end;
    end
    else if stNomChamp = 'TABLELIBRE3' then
    begin
      if GetField('I_TABLE2') <> GetControlText('TABLELIBRE3') then
      begin
        if DS.State in [dsBrowse] then
          DS.edit;
        SetField('I_TABLE2', GetControlText('TABLELIBRE3'));
      end;
    end
    {else if stNomChamp = 'SECTIONS1' then
    begin
      if GetField('I_TABLE9') <> GetControlText('SECTIONS1') then
      begin
        if DS.State in [dsBrowse] then
          DS.edit;
        SetField('I_TABLE9', GetControlText('SECTIONS1'));
      end;
    end} //XVI 25/04/2006 FD 3978 On ne va jamais y passer par là, le controle ne fait pas partie de la Table....
    else if (Copy(stNomChamp, 1, 7) = 'I_TABLE') then
    begin
{$IFDEF SERIE1}
{$ELSE}
      begin
        if stNomChamp='I_TABLE9' then exit;
        wName := stNomChamp;
        wNum := wName[length(wName)];
        wVal := GetField(stNomChamp);
        if wVal <> wTableLib[strtoint(wNum)] then
          wTableLib[strtoint(wNum)] := wVal
        else
          exit; //YCP
      end;
      if (wVal <> '') and not
        ExisteSQL('SELECT NT_NATURE FROM NATCPTE WHERE NT_TYPECPTE="I0' + wNum +
        '" AND NT_NATURE="' + wVal + '"') then
      begin
        if
          PgiAsk('ATTENTION, le code de table libre n''existe pas. Désirez vous continuer ?', Ecran.caption) <> mrYes then
        begin
          Ecran.ModalResult := mrNone;
          SetActiveTabSheet('ZL');
          SetFocusControl(stNomChamp);
        end;
      end;
{$ENDIF}
    end

    else if stNomChamp = 'I_SURAMORT' then
    begin
      if fLeMode = taConsult then
        exit;
      if DS.State in [dsInsert, dsEdit] then
      begin
        if bChargementEnCours then
          exit;
        if ((GetCheckBoxState('I_SURAMORT') = cbChecked) and
          ((GetField('I_DATEPIECEA') >= StrToDate('01/04/2002')) and
          (GetField('I_DATEPIECEA') <= StrToDate('31/12/2003')))) then
        begin
          stMsg := 'La mesure temporaire s''applique aux biens ayant fait l''objet pendant la période du 17/10/2001 au 31/03/2002'
            + #10#13 +
            'd''une commande ferme assortie du versement d''acomptes au moins égaux à 10 % du montant total de la commande' + #10#13
              +
            'et dont l''acquisition ou la fabrication intervient avant le 31/12/2003';
          if PGIAsk(stMsg, Ecran.Caption) <> mrYes then
            SetControlChecked('I_SURAMORT', False);
        end;
        bPlanModifie := True;
        GereRepriseEco;
        if bPlanFiscal then GereRepriseFisc;
        TraiteDotationAnterieures;
        SetField('I_REPRISEFISCAL', CumFiscSav);
        SetField('I_REPRISEECO', CumEcoSav);
        {chantier fiscal fq 17512
        // ajout pour fq 15274 - mbo
        if (CumFiscSav < CumEcoSav) and (GetField('I_METHODEECO') = 'VAR') then
           SetControlProperty('AmortDerog','Value', 0)
        else
           SetCOntrolProperty('AmortDerog', 'Value', CumFiscSav - CumEcoSav);
        }
        AfficheDerogatoire;
      end;
    end
    else if stNomChamp = 'I_QUANTITE' then
    begin
      if VarIsNull(GetField('I_QUANTITE'))  then SetField('I_QUANTITE',0);
    {$IFDEF SERIE1}
    end
    else if (stNomChamp = 'I_VENTILABLE') and (ds.State<>dsBrowse) then
    begin
       for idx:=1 to VS1Cpta.Y_NBSections do
          setfield('I_VENTILABLE'+inttostr(idx),F.AsString) ;
    {$ELSE}
    {$ENDIF !SERIE1} //XVI 18/04/2006 FD 3978
    end;
  finally
    fDisableChangeField := False;
  end;
end;


function Tom_immo.Test_Reprise_Depreciation(StChamp:String ):boolean;
Begin
  // test de la reprise dépréciation
  Result:=True;
  // ajout fq 18582 - mbo - 13/07/2006
  if GetField(StChamp) < 0 then
  begin
     PGIInfo('Montant négatif !','Attention');
     Result:=False;
     SetField(StChamp,0);
     SetFocusControl(StChamp);
  end
  else
  begin
  IF StChamp = 'I_REPRISEDEP' THEN
     Begin
       // mbo correction du test 19.06.06 fq 17500
       if (GetField('I_REPRISEECO')<>0) AND (GetField('I_REPRISEECO')+ GetField('I_REPRISEDEP')>GetField('I_BASEECO')) Then
         Begin
           PGIBox(HMC[39]);
           Result:=False;
           SetFocusControl(StChamp);
         end
     End
  Else IF StChamp = 'I_REPRISEFISCAL' THEN
     Begin
       // mbo correction du test 19.06.06 fq 17500
       if (GetField('I_REPRISEFISCAL')<>0) AND (GetField('I_REPRISEFISCAL')+ GetField('I_REPRISEDEP')>GetField('I_BASEFISC')) Then
         Begin
           PGIBox(HMC[40]);
           Result:=False;
           SetFocusControl(StChamp);
         end;
     End;
  end;
end;



procedure TOM_IMMO.OnArgument(S: string);
var
  St: string;
  i: integer;
begin
  inherited;
  fDisableChangeField := False;
  // AssignFile(fFichier, 'c:\test.log');
  // Rewrite(fFichier);
  // CloseFile(fFichier);
  { Initialisation }
  SetControlVisible('BINSERT', True);
  SetControlVisible('BDELETE', True);
  fDisableChangeField := False;
  fLog := TOB.Create('', nil, -1);
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then
    fLeMode := taModif
  else if St = 'ACTION=CREATION' then
    fLeMode := taCreatOne
  else if St = 'ACTION=CREATIONSERIE' then
    fLeMode := taCreatEnSerie
  else
    fLeMode := taConsult;
  St := ReadTokenSt(S);
  St := ReadTokenSt(S);
  fNature := St;

  // Paramètres DPI
  fDateDebExoDPI := iDate1900;
  fDateFinExoDPI := iDate1900;
  fMontantMaxiDPI:= 0.00;
  St := ReadTokenSt(S);
  if St <> '' then fDateDebExoDPI := StrToDate (St);
  St := ReadTokenSt(S);
  if St <> '' then fDateFinExoDPI := StrToDate (St);
  St := ReadTokenSt(S);
  if St <> '' then fMontantMaxiDPI := StrToFloat (St);

  // BTY 07/06 Paramètres DPI : dates 5 exos N-1 à N-5
{$IFNDEF CMPGIS35}
  fAmDPI := TAmDPI.Create;
  St := ReadTokenSt(S);
  for i:= 1 to 5 do
     begin
     if St <> '' then fAmDPI.Exercice[i].Deb := StrToDate (St);
     St := ReadTokenSt(S);
     if St <> '' then fAmDPI.Exercice[i].Fin := StrToDate (St);
     St := ReadTokenSt(S);
     end;
  // BTY 07/06 Paramètres DPI : soldes 5 exos N-1 à N-5
  for i:= 1 to 5 do
     begin
     if St <> '' then fAmDPI.Solde[i] := StrToFloat(St);
     St := ReadTokenSt(S);
     end;
{$ENDIF}

  { Evènements }
{$IFDEF EAGLCLIENT}
  THEdit(GetControl('I_COMPTEIMMO')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_COMPTELIE')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_COMPTEAMORT')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_COMPTEDOTATION')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_REPEXPLOIT')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_DOTATIONEXC')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_REPEXCEP')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_VACEDEE')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_AMORTCEDE')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_VAOACEDEE')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_COMPTEDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_REPRISEDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_PROVISDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THEdit(GetControl('I_ORGANISMECB')).OnElipsisClick := OnCompteElipsisClick;
{$ELSE}
  THDBEdit(GetControl('I_COMPTEIMMO')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_COMPTELIE')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_COMPTEAMORT')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_COMPTEDOTATION')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_REPEXPLOIT')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_DOTATIONEXC')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_REPEXCEP')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_VACEDEE')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_AMORTCEDE')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_VAOACEDEE')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_COMPTEDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_REPRISEDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_PROVISDEROG')).OnElipsisClick := OnCompteElipsisClick;
  THDBEdit(GetControl('I_ORGANISMECB')).OnElipsisClick := OnCompteElipsisClick;
{$ENDIF}

  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFFiche(Ecran).OnAfterFormShow := OnAfterFormShowEcran;
  THEdit(GetControl('COMPTE_DEPOT')).OnElipsisClick := OnCompteDepotElipsisClick;
  THEdit(GetControl('NbEcheances')).OnEnter := NbEcheancesEnter;
  THEdit(GetControl('NbEcheances')).OnExit := NbEcheancesExit;
  THValComboBox(GetControl('I_GROUPEIMMO')).OnChange  := OnChangeRegroupement;
  TMenuItem(GetControl('BINFOREG')).OnClick := OnInformationRegroupementClick;
  TMenuItem(GetControl('BDETAILREG')).OnClick := OnDetailRegroupementClick;
  TCheckBox(GetControl('bAmortissementFiscal')).OnClick := OnClickAmortissementFiscal;
  TCheckBox(GetControl('bReintegrationFiscale')).OnClick := OnClickReintegrationFiscale;
  //PGR 26/10/2005 Ajout coche de calcul du plan futur avec la VNC
  TCheckBox(GetControl('bCalculAvecVnc')).OnClick := OnClickSurCalculAvecVnc;

  // chantier fiscal fq 17512
  TCheckBox(GetControl('bCalculAvecVNF')).OnClick := OnClickSurCalculAvecVNF;
  TCheckBox(GetControl('bGestionFiscale')).OnClick := OnClickGestionFiscale;

  TCheckBox(GetControl('I_SURAMORT')).OnClick := OnClickSurAmortissement;
  TCheckBox(GetControl('I_VENTILABLE')).OnClick := OnClickVENTILABLE;
  TCheckBox(GetControl('ELIGIBLETP')).OnClick := OnClickELIGIBLETP;
  TToolbarButton97(GetControl('SaisieEcheancier')).OnClick := OnClickSaisieEcheancier;
  TToolbarButton97(GetControl('BUNITEOEUVRE')).OnClick := OnClickUniteOeuvre;
  TPageControl(GetControl('Pages')).OnChange := OnChangeOnglet;
  TPopUpMenu(GetControl('POPZOOMACTION', True)).OnPopup := OnPopZoomActionPopup;
  // ajout pour pop zoomplan - mbo suite gestion prime
  TPopUpMenu(GetControl('POPZOOMPLAN', True)).OnPopup := OnPopZoomPlan;

  TMenuItem(GetControl('Mutation')).OnClick := OnMutationClick;
  TMenuItem(GetControl('Cession')).OnClick := OnCessionClick;
  TMenuItem(GetControl('Eclatement')).OnClick := OnEclatementClick;
  TMenuItem(GetControl('AnnulationOperation')).OnClick := AnnulationOperationClick;
  TMenuItem(GetControl('DureeMethode')).OnClick := DureeMethodeClick;
  TMenuItem(GetControl('ElementExceptionnel')).OnClick := ElementExceptionnelClick;
  TMenuItem(GetControl('LeveeOption')).OnClick := LeveeOptionClick;
  TMenuItem(GetControl('ModifBases')).OnClick := OnModifBasesClick;
  TMenuItem(GetControl('BMODELEREG')).OnClick := OnModeleRegroupementClick;
  TCheckBox(GetControl('bDepotGarantie')).OnClick := OnDepotGarantieClick;
  // BTY 09/05 fq 14363
  TPopUpMenu(GetControl('POPREGROUPEMENT', True)).OnPopup := OnPopRegroupementPopup;
  //PGR - 09/2005 - Changement de méthode
  TMenuItem(GetControl('ChangementMethode')).OnClick := ChangementMethodeClick;
  //TGA - 09/2005 - Dépréciation d'actif
  TMenuItem(GetControl('DepActif')).OnClick := OnDepreciationActifClick;
  //MBO - 04/2006 - FQ 17923 - changement date de mise en service
  TMenuItem(GetControl('ModifService')).OnClick := ModifMiseServiceClick;
  //TGA - 08/2006 - Suppression de la DPI
  TMenuItem(GetControl('SupDPI')).OnClick := OnSuppressionDPIClick;
  //MBO - 09/2006 - saisie d'une prime d'équipement
  TMenuItem(GetControl('Prime')).OnClick := SaisiePrimeClick;
  //MBO - 10/2006 - saisie d'une subvention d'équipement
  TMenuItem(GetControl('Sbv')).OnClick := SaisieSbvClick;
  //MBO - 04/2007 - saisie d'un Remplacement de composant
  TMenuItem(GetControl('Remplacement')).OnClick := SaisieRemplacementClick;
  //BTY 10/06 Réduction de la prime d'équipement
  TMenuItem(GetControl('ReductionPrime')).OnClick := OnReductionPrimeClick;
  //BTY 11/06 Réduction de la subvention
  TMenuItem(GetControl('ReductionSubvention')).OnClick := OnReductionSBVClick;
  //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
  THValComboBox(GetControl('TYPEBASEFISC')).OnExit  := OnExitTypeBaseFisc;
  //mbo - 25.01.06 - FQ 17362
  TCheckBox(GetControl('cbI_DATESUSPCB')).OnClick := OnClickSuspensionCB;

{$IFDEF EAGLCLIENT}
  THEdit(GetControl('I_CODEPOSTAL')).OnDblClick := OnLieuDblClick;
  THEdit(GetControl('I_VILLE')).OnDblClick := OnLieuDblClick;
{$ELSE}
  THDBEdit(GetControl('I_CODEPOSTAL')).OnDblClick := OnLieuDblClick;
  THDBEdit(GetControl('I_VILLE')).OnDblClick := OnLieuDblClick;
{$ENDIF}
  TToolbarButton97(GetControl('BVentil')).OnClick := BVentilClick;
  TToolbarButton97(GetControl('VisuDepotGarantie')).OnClick :=
    OnVisuDepotGarantieClick;
  TToolbarButton97(GetControl('OpeEnCours')).OnClick := OnOpeEnCoursClick;
  TToolbarButton97(GetControl('AccesImmoLie')).OnClick := OnAccesImmoLie;

  // modif pour prime - TToolbarButton97(GetControl('BPlanAMort')).OnClick := OnPlanAmortClick;
  TMenuItem(GetControl('ZoomPlanAmort')).OnClick := OnPlanAmortClick;
  TMenuItem(GetControl('ZoomPlanTheo')).OnClick := OnPlanAmortClick;
  TMenuItem(GetControl('ZoomPlanPrime')).OnClick := OnPlanPrimeClick;
  TMenuItem(GetControl('ZoomPlanSbv')).OnClick := OnPlanSbvClick;
  TToolbarButton97(GetControl('BPlanRemboursement')).OnClick :=
    OnPlanRemboursementClick;

  // TGA 24/03/2006
  TToolbarButton97(GetControl('SRegroupement')).OnClick := OnClickSaisieRegroupement;

  {$IFDEF EAGLCLIENT}
  THRichEditOle(GetControl('I_BLOCNOTE', True)).OnMouseDown := OnMouseDownI_BlocNote;
  {$ELSE}
  THDBRichEditOle(GetControl('I_BLOCNOTE', True)).OnMouseDown := OnMouseDownI_BlocNote;
  {$ENDIF}
  TTabSheet(GetControl('ZLS1')).OnShow := OnShowZLS1;
  // BTY 07/06 DPI
  TToolbarButton97(GetControl('SBaseEcoOuDPI')).OnClick := OnClickBaseEco;

  // TGA 06/06/2006
  //if fDateDebExoDPI <> iDate1900 then
  //  // appel derrière le tableau de suivi des DPI
  //  TToolbarButton97(GetControl('SBaseEcoOuDPI')).Hint := 'Utiliser une DPI'
  //else
  //TToolbarButton97(GetControl('SBaseEcoOuDPI')).Hint := 'Indiquer la base';

  // TGA 06/09/2006
  TToolbarButton97(GetControl('BDPI')).OnClick := OnClickSaisieDPI;
  // Bouton visible si fiche immo appelée par tableau DPI
  // BTY 09/06 Boutons de saisie Base Eco et DPI distincts
  if (fDateDebExoDPI <> iDate1900) then
     begin
     SetControlVisible('SBaseEcoOuDPI', False);
     SetControlVisible('BDPI', True);
     TToolbarButton97(GetControl('BDPI')).Hint := 'Utiliser une DPI';
     end
  ELSE
     begin
     SetControlVisible('BDPI', False);
     SetControlVisible('SBaseEcoOuDPI', True);
     TToolbarButton97(GetControl('SBaseEcoOuDPI')).Hint := 'Indiquer la base';
     end;
  SetControlEnabled('SBaseEcoOuDPI', False);
  SetControlEnabled('BDPI', False);
  // Mention DPI remplacée par un bouton qui ouvre l'onglet Informations
  TToolbarButton97(GetControl('bInfoDPI')).OnClick := OnInfoDPIClick;


  { Initialisation des zones }

{$IFDEF SERIE1}
  {SetControlVisible('BVentil', false);
  SetControlVisible('I_VENTILABLE', true);} //XVI fd18/04/2006 FD 3978
  SetControlVisible('I_ETABLISSEMENT', false);
  SetControlVisible('tI_ETABLISSEMENT', false);
  SetControlVisible('ZL', false);
  SetControlVisible('CreationLieu', false);
  //XVI 25/04/2006 FD 3978 Analytioque mono-ventilation début
  SetControlVisible('SECTIONS1', VS1Cpta.Y_Saisie and not VS1CPta.Y_analytiqueAdvanced);
  SetControlVisible('tSECTIONS1', GetControlVisible('SECTIONS1'));
  SetControlProperty('SECTIONS1', 'DataType', 'TTSECTION');
  THCritMaskedit(GetControl('SECTIONS1')).OnExit:=SectionS1OnExit ;
  THValComboBox(GetControl('TABLELIBRE1')).onChange:=TableLibre1Change ;
  THValComboBox(GetControl('TABLELIBRE2')).onChange:=TableLibre1Change ;
  THValComboBox(GetControl('TABLELIBRE3')).onChange:=TableLibre1Change ;
  if (fLeMode in [taCreat..taCreatOne]) then
    Ecran.HelpContext := 511010
  else if fLeMode = tamodif then
    Ecran.HelpContext := 511020
  else if fLeMode = taconsult then
    Ecran.HelpContext := 511030;
{$ELSE}
  SetControlVisible('ZLS1', false);
  // FQ 14013 - mbo 05/04/2007 nouvelles zones section onglet champs libres
  SetControlVisible('tSECTIONS1', false);
  SetControlVisible('SECTIONS1', false);
{$ENDIF}

  if (fLeMode in [taCreat..taCreatOne]) then
    SetFocusControl('I_IMMO');
  SetControlVisible('BFirst', not (fLeMode = taCreatEnSerie));
  SetControlVisible('BPrev', not (fLeMode = taCreatEnSerie));
  SetControlVisible('BNext', not (fLeMode = taCreatEnSerie));
  SetControlVisible('BLast', not (fLeMode = taCreatEnSerie));
//  SetControlVisible('BInsert', False);
  SetControlVisible('BDefaire', not (fLeMode = taCreatEnSerie));
  SetControlVisible('BZoomAction', (not (fLeMode = taCreatEnSerie)) and (fLeMode
    = taModif));
  // BTY 10/06 Bouton Opération en cours caché après clôture des immos
   SetControlVisible('OpeEncours', (not (fLeMode = taCreatEnSerie)) and (fLeMode
    = taModif));

  // Ajout mbo pour gestion prime
  SetControlVisible('BPlanAmort', (not (fLeMode = taCreatEnSerie)) and (fLeMode
    = taModif));

  THValComboBox(GetControl('I_METHODEFISC')).Items.Delete(THValComboBox(GetControl('I_METHODEFISC')).Items.IndexOf('Non amortissable'));

  { Mise à jour des zones de saisie des tables libres }
{$IFDEF SERIE1}
  ImLibellesTableLibre(TTabSheet(GetControl('ZLS1')), 'TT_TABLELIBREIMMO', '', 'I');
{$ELSE}
  ImLibellesTableLibre(TTabSheet(GetControl('ZL')), 'TI_TABLE', 'I_TABLE', 'I');
  { Table libre 9 utilisée pour le stockage 'éligible TP' }
  SetControlProperty ('I_TABLE9','Visible',False);
  SetControlVisible ('TI_TABLE9',False);
{$ENDIF}

  { Supression de la méthode variable du choix des méthodes }
  //SetControlProperty('I_METHODEECO','Plus',' AND CO_CODE<>"VAR"');

  // ajout mbo pour la gestion de remplacement de composant
  SetControlVisible('I_REMPLACE', false);  // le composant
  SetControlVisible('I_REMPLACEE', false);  // l'immo d'origine
end;

// TGA 06/09/2006 Saisie de la DPI ==============================================
procedure TOM_IMMO.OnClickSaisieDPI(Sender: TObject);
var
  resultat, stParam : string;
  PlanAmor2 : TPlanAmort;
  DateFinAmortissement : TDateTime;
  BaseAmort: double;
begin
  FDPI := False;
  FDPIANT := False;

  // Assurer le déplacement du focus
  NextPrevControl(Ecran);

  // On est derrière le tableau de suivi des DPI
{$IFNDEF CMPGIS35}
  if fDateDebExoDPI <> iDate1900 then
      begin
      // Soldes chargés à partir du tableau de suivi ont bougé
      // si on vient de faire une affectation DPI à une immo
      // => Recharger les soldes à partir de immomvtd
      fAmDPI.ChargeDateEtSoldeDPI;
      // => Recharger le solde maxi disponible à investir à fDateFinExoDPI
      // FQ 18992 Recharger le solde max à investir à date fin de l'exo N encours
      //fAmDPI.ChargeSoldeDisponibleDate (fDateFinExoDPI);
      fAmDPI.ChargeSoldeDisponibleDate (VHImmo^.Encours.Fin);
      fMontantMaxiDPI := fAmDPI.SoldeDispoDate;

      // Faire pointer la variable globale TheData accessible partout, sur l'objet DPI
      TheData := fAmDPI;

      if (GetField('I_DATEDEBECO') >= VHImmo^.Encours.Deb) and
         (GetField('I_DATEDEBECO') <= VHImmo^.Encours.Fin) then
         begin
           if GetField('I_REINTEGRATION') <> 0 then
              PGIBox(HMC[46])
           else if (GetCheckBoxState('bGestionFiscale') = cbChecked) then
                 PGIBox(HMC[49])   // ajout mbo FQ 17512
           else
              begin
                stParam := FloatToStr (GetField('I_REINTEGRATION')) + ';' +
                           FloatToStr (GetField('I_BASEECO')) + ';' +
                           FloatToStr (fMontantMaxiDPI) + ';' +
                           DateToStr (fDateFinExoDPI);
                resultat := AMLanceFiche_MontantImmoDPI(stParam);
                IF resultat<>'ECHAP' Then
                   BEGIN
                     //FmontantDPI := StrToFloat(resultat);
                     FmontantDPI := Valeur(resultat);
                     FDPI:= True;
                     //TGA 25/08/2006 Provoque une modif pour enregistrer et fermer la fiche sur BValider
                     DS.Edit;
                     TToolbarButton97(GetControl('BValider')).Click;
                   End;
              end;
         end

      else if GetField('I_DATEDEBECO') < VHImmo^.Encours.Deb then
         begin
//           FQ 18968
//           if GetField('I_DATEDEBECO') < fAmDPI.Exercice[5].Deb then
//              PGIBox(HMC[44])
//           else

         // Calculer la date de fin d'amortissement éco en l'état actuel des données de la fiche
         PlanAmor2 := TPlanAmort.create(true) ;
         PlanAmor2.Charge(TQuery(DS));
         BaseAmort := GetField('I_MONTANTHT')+GetField('I_TVARECUPERABLE')-Getfield('I_TVARECUPEREE');
         SetField('I_BASEAMORDEBEXO', BaseAmort);
         SetField('I_BASEAMORFINEXO', BaseAmort);
         PlanAmor2.Calcul(TQuery(DS), VHImmo^.Encours.Deb);

         DateFinAmortissement := PlanAmor2.GetDateFinAmortEx(PlanAmor2.AmortEco);
         PlanAmor2.Free;

         if (DateFinAmortissement < VHImmo^.Encours.Deb)  then
             PGIBox(HMC[44])
         else
             begin
                 stParam := DateToStr (GetField('I_DATEDEBECO')) + ';' +
                            FloatToStr (GetField('I_REINTEGRATION')) + ';' +
                            FloatToStr (GetField('I_BASEECO')) + ';' +
                            FloatToStr (GetField('I_MONTANTHT')) + ';' +
                            FloatToStr (GetField('I_TVARECUPERABLE')) + ';' +
                            FloatToStr (GetField('I_TVARECUPEREE'));
                 // Faire pointer la variable globale TheData accessible partout, sur l'objet DPI
                 // remontée plus haut TheData := fAmDPI;
                 if AMLanceFiche_DPIAnterieurs ( stParam )= mrYes then
                   begin
                     // TheData/fAmDPI contiennent les infos modifiées par AMLanceFiche_DPIAnterieurs
                     FDPIANT:=True;
                     // Provoque une modif pour enregistrer et fermer la fiche sur BValider
                     DS.Edit;
                     TToolbarButton97(GetControl('BValider')).Click;
                   end;
                 // déplacé plus bas TheData := nil;
             end;
         end;
      TheData := nil; // BTY 10/06
      End;
{$ENDIF}
end;




procedure TOM_IMMO.OnClickSaisieRegroupement(Sender: TObject);
begin
  AMLanceFiche_SaisieRegroupement('IGI');
  {$IFDEF EAGLCLIENT}
     THValComboBox(GetControl('I_GROUPEIMMO')).Reload;
  {$ELSE}
     THDbValComboBox(GetControl('I_GROUPEIMMO')).Reload;
  {$ENDIF}
end;

// ==========================================================================
procedure TOM_IMMO.OnClickBaseEco(Sender: TObject);
var
 stParam : string;
 TobW : TOB;
begin


 // Fiche immo normale
   if (fLeMode in [taCreat..taCreatOne]) or (DS.State = dsInsert) then
        stParam := 'ACTION=CREATION'
   else stParam := 'ACTION=MODIFICATION';

   stParam := stParam + ';' + FloatToStr(Arrondi (GetFIeld('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
                              Getfield('I_TVARECUPEREE'), V_PGI.OkDecV));

   // Stocker base éco dans la TOB globale TheTob accessible partout
   TobW := TOB.Create('', nil, -1);
   TobW.AddChampSupValeur('BASEECO', GetField ('I_BASEECO'));
   TheTob := TobW;

{$IFNDEF CMPGIS35}
   if AMLanceFiche_BASEECO ( stParam ) = mrYes then
      begin
      // TheTob systématiquement réinitialisé à nil => TobW contient la base éco modifiée
      if fLeMode = taConsult then  exit;
      if (DS.State = dsBrowse) then  DS.Edit;
      if (DS.State in [dsInsert, dsEdit]) then
         SetField('I_BASEECO', TobW.GetValue ('BASEECO'));
      end;
{$ENDIF}
   TobW.Free;
end;


procedure TOM_IMMO.OnClose;
begin
  if fPlanUO <> nil then
    FreeAndNil(fPlanUO);
  if fContrat <> nil then
    FreeAndNil(fContrat);
  if fPlanAmor <> nil then
    FreeAndNil(fPlanAmor);
  if fLog <> nil then
    FreeAndNil(fLog);
{$IFNDEF CMPGIS35}
  if fAmDPI <> nil then // BTY 07/06
     fAmDPI.Free;
{$ENDIF}
  inherited;
end;

procedure TOM_IMMO.OnCancelRecord;
begin
  inherited;
  EnableChamps;
  // ajout mbo 26.07.05  FQ 16258
  AfficheMontantTTC;
end;

procedure TOM_IMMO.MajComptesAssocies;
var
  Q: TQuery;
  Prefix: string;
begin
  if ControleCompteImmo(GetField('I_COMPTEIMMO'), GetField('I_NATUREIMMO')) then
  begin
{$IFDEF SERIE1}
    Prefix := 'IC';
{$ELSE}
    Prefix := 'PC';
{$ENDIF}
    if DS.State in [dsInsert, dsEdit] then
    begin
      Q := OpenSQL('SELECT * FROM IMMOCPTE WHERE ' + Prefix + '_COMPTEIMMO="' +
        GetField('I_COMPTEIMMO') + '"', TRUE);
      if not Q.EOF then
      begin
        SetField('I_COMPTEAMORT', Q.FindField(Prefix +
          '_COMPTEAMORT').AsString);
        SetField('I_COMPTEDOTATION', Q.FindField(Prefix +
          '_COMPTEDOTATION').AsString);
        SetField('I_COMPTEDEROG', Q.FindField(Prefix +
          '_COMPTEDEROG').AsString);
        SetField('I_REPRISEDEROG', Q.FindField(Prefix +
          '_REPRISEDEROG').AsString);
        SetField('I_PROVISDEROG', Q.FindField(Prefix +
          '_PROVISDEROG').AsString);
        SetField('I_DOTATIONEXC', Q.FindField(Prefix +
          '_DOTATIONEXC').AsString);
        SetField('I_VACEDEE', Q.FindField(Prefix + '_VACEDEE').AsString);
        SetField('I_AMORTCEDE', Q.FindField(Prefix + '_AMORTCEDE').AsString);
        SetField('I_VAOACEDEE', Q.FindField(Prefix + '_VOACEDE').AsString);
        SetField('I_REPEXPLOIT', Q.FindField(Prefix + '_REPEXPLOIT').AsString);
        SetField('I_REPEXCEP', Q.FindField(Prefix + '_REPEXCEP').AsString);
      end
      else
      begin
        SetField('I_COMPTEAMORT', CAssAmortissement(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_COMPTEAMORT')) then
          SetField('I_COMPTEAMORT', '');
        SetField('I_COMPTEDOTATION', CAssDotation(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_COMPTEDOTATION')) then
          SetField('I_COMPTEDOTATION', '');
        SetField('I_COMPTEDEROG', CAssDerog(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_COMPTEDEROG')) then
          SetField('I_COMPTEDEROG', '');
        SetField('I_REPRISEDEROG', CAssRepriseDerog(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_REPRISEDEROG')) then
          SetField('I_REPRISEDEROG', '');
        SetField('I_PROVISDEROG', CAssProvisDerog(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_PROVISDEROG')) then
          SetField('I_PROVISDEROG', '');
        SetField('I_DOTATIONEXC',
          CAssDotationExc(GetField('I_COMPTEDOTATION')));
        if not VerifCompteAssocie(GetControl('I_DOTATIONEXC')) then
          SetField('I_DOTATIONEXC', '');
        SetField('I_VACEDEE', CAssVaCedee(GetField('I_COMPTEIMMO')));
        if not VerifCompteAssocie(GetControl('I_VACEDEE')) then
          SetField('I_VACEDEE', '');
        SetField('I_AMORTCEDE', GetField('I_COMPTEAMORT'));
        if not VerifCompteAssocie(GetControl('I_AMORTCEDE')) then
          SetField('I_AMORTCEDE', '');
        SetField('I_VAOACEDEE', GetField('I_COMPTEIMMO'));
        if not VerifCompteAssocie(GetControl('I_VAOACEDEE')) then
          SetField('I_VAOACEDEE', '');
        SetField('I_REPEXPLOIT', CAssRepExploit(GetField('I_COMPTEDOTATION')));
        if not VerifCompteAssocie(GetControl('I_REPEXPLOIT')) then
          SetField('I_REPEXPLOIT', '');
        SetField('I_REPEXCEP', CAssRepExc(GetField('I_DOTATIONEXC')));
        if not VerifCompteAssocie(GetControl('I_REPEXCEP')) then
          SetField('I_REPEXCEP', '');
      end;
      Ferme(Q);
    end;
  end
  else
    RazComptesAssocies;
end;

procedure TOM_IMMO.RazComptesAssocies;
begin
  SetField('I_COMPTEAMORT', '');
  SetField('I_COMPTEDOTATION', '');
  SetField('I_COMPTEDEROG', '');
  SetField('I_REPRISEDEROG', '');
  SetField('I_PROVISDEROG', '');
  SetField('I_DOTATIONEXC', '');
  SetField('I_VACEDEE', '');
  SetField('I_AMORTCEDE', '');
  SetField('I_VAOACEDEE', '');
  SetField('I_REPEXPLOIT', '');
  SetField('I_REPEXCEP', '');
end;

procedure TOM_IMMO.EnableOngletsFiche;
begin
  SetControlVisible('PComptable', GetField('I_NATUREIMMO') = 'PRO');
  SetControlVisible('PContrat', (GetField('I_NATUREIMMO') = 'CB') or
    (GetField('I_NATUREIMMO') = 'LOC'));
  SetControlVisible('PPlan', (GetField('I_NATUREIMMO') = 'PRO') or
    (GetField('I_NATUREIMMO') = 'CB'));
  if GetControlVisible('PContrat') then
    EnableChampsOngletContrat;
  if GetControlVisible('PPlan') then
    EnableChampsOngletPlan;
end;

procedure TOM_IMMO.EnableChampsLibre;
begin
  SetControlEnabled('TABLELIBRE1', (fLeMode <> taConsult));
  SetControlEnabled('TT_TABLELIBREIMMO1', GetControlEnabled('TABLELIBRE1'));
  SetControlEnabled('TABLELIBRE2', (fLeMode <> taConsult));
  SetControlEnabled('TT_TABLELIBREIMMO2', GetControlEnabled('TABLELIBRE2'));
  SetControlEnabled('TABLELIBRE3', (fLeMode <> taConsult));
  SetControlEnabled('TT_TABLELIBREIMMO3', GetControlEnabled('TABLELIBRE3'));
end;

procedure TOM_IMMO.EnableChampsCommuns;
var
  bOuvert: boolean;
  sEtat: string;
begin
  sEtat := GetField('I_ETAT');
  bOuvert := (sEtat <> 'CLO') and (sEtat <> 'FER') and (sEtat <> 'SUS')
    and (fLeMode <> taConsult) and (not fOperationEnCours) and (not
    ((GetField('I_IMMOLIEGAR') <> '') and (GetField('I_NATUREIMMO') = 'FI')))
    and (not fDPIUtilisee) and (not fComposant) ; // BTY 07/06 DPI + mbo 27/04/2007 pour remplacement

  if DS.State = dsInsert then
  begin
    if (( fLeMode in [taCreat..taCreatOne]) and (fNature <> '') and  ( gTobImmo = nil )) then
      SetControlEnabled('I_NATUREIMMO', false)
    else
    begin
      if GetField('I_NATUREIMMO') = '' then
        SetField('I_NATUREIMMO', 'PRO');
      SetControlEnabled('I_NATUREIMMO', bOuvert);
    end;
  end
  else
  SetControlEnabled('I_NATUREIMMO', False);
  SetControlEnabled('I_LIBELLE', True);
  SetControlEnabled('I_QUALIFIMMO', bOuvert);
  // BTY 09/06 FQ 16149
  //SetControlEnabled('I_ETABLISSEMENT', bOuvert);
  {$IFNDEF SERIE1}
  {$IFNDEF EAGLSERVER}
  if DS.State <> dsInsert then
     SetControlEnabled('I_ETABLISSEMENT',
     ((bOuvert) and (VHImmo^.EtablisCpta) and (not VH^.ProfilUserC[prEtablissement].ForceEtab)));
  {$ENDIF}
  {$ENDIF}

  SetControlEnabled('I_LIEUGEO', bOuvert);
  SetControlEnabled('CreationLieu', bOuvert);
  SetControlEnabled('I_PAYS', bOuvert);
  SetControlEnabled('I_CODEPOSTAL', bOuvert);
  SetControlEnabled('I_VILLE', bOuvert);
  SetControlEnabled('I_DATEPIECEA', bOuvert);
  // BTY 04/06 FQ 17516
  SetControlEnabled('I_GROUPEIMMO', bOuvert);
  SetControlEnabled('SREGROUPEMENT', bOuvert);
end;

procedure TOM_IMMO.EnableChampsOngletGeneral;
var
  bOuvert, bEtatChamp, bEtatLibelle, bDepot: boolean;
  Nat: string;
begin
  Nat := GetField('I_NATUREIMMO');

  bOuvert := (GetField('I_ETAT') <> 'FER') and (GetField('I_ETAT') <> 'CLO')
    and (fLeMode <> taConsult) and (not fOperationEnCours)
    and (not fDPIUtilisee) and (not fComposant);  // BTY 07/06 DPI
  bDepot := (GetField('I_NATUREIMMO') = 'FI') and (GetField('I_IMMOLIEGAR') <>
    '');
  bEtatChamp := (Nat <> 'LOC') and (bOuvert) and (not bDepot) and (not fComposant);
  bEtatLibelle := not (Nat = 'LOC');
  SetControlEnabled('I_MONTANTHT', (Nat <> 'LOC') and (bOuvert));
  SetControlEnabled('tI_MONTANTHT', bEtatLibelle);
  SetControlEnabled('I_TVARECUPERABLE', bEtatChamp);
  SetControlEnabled('tI_TVARECUPERABLE', bEtatLibelle);
  SetControlEnabled('I_TVARECUPEREE', bEtatChamp);
  SetControlEnabled('tI_TVARECUPEREE', bEtatLibelle);
  SetControlEnabled('MontantTTC', bEtatChamp);
  SetControlEnabled('tMontantTTC', bEtatLibelle);

  SetControlEnabled('I_QUANTITE', (GetField('I_ETAT') <> 'FER')
    and (fLeMode <> taConsult)
    and (not fOperationEnCours)
    and (not fDPIUtilisee) and (not fComposant) );   // BTY 07/06 DPI

  SetControlEnabled('tI_QUANTITE', bEtatLibelle);
  SetControlEnabled('I_BASETAXEPRO', (Nat <> 'FI') or (Nat <> 'LOC'));
  SetControlEnabled('tI_BASETAXEPRO', GetControlEnabled('I_BASETAXEPRO'));
  SetControlEnabled('I_NATUREBIEN', bOuvert and ((Nat = 'LOC') or
    (GetField('I_BASETAXEPRO') <> 0.0)) and (Nat <> 'FI'));
  SetControlEnabled('tI_NATUREBIEN', GetControlEnabled('I_NATUREBIEN'));
  SetControlEnabled('I_COMPTELIE', GetControlVisible('PContrat') and bOuvert);
  SetControlEnabled('tI_COMPTELIE', GetCOntrolVisible('PContrat'));
  SetControlEnabled('I_COMPTEIMMO', ((bOuvert) or (Nat = 'LOC') or (Nat = 'CB'))
    and (not bDepot));
  SetControlEnabled('I_DATEPIECEA', (bOuvert) and (not bDepot));
  SetControlEnabled('I_DATEAMORT', ((Nat = 'PRO') or (Nat = 'CB')) and bOuvert);
  SetControlEnabled('tI_DATEAMORT', (Nat = 'PRO') or (Nat = 'CB'));
{$IFDEF SERIE1}
  SetControlVisible('I_VENTILABLE',VS1Cpta.Y_Saisie);  //XVI 25/04/2006 FQ 3978 analytoque monoventilation
  SetControlVisible('BVentil', VS1Cpta.Y_analytiqueAdvanced) ; //XVI 25/04/2006 FQ 3978 analytoque monoventilation
{$ELSE}
  SetControlVisible('I_VENTILABLE', (GetParamSocSecur('SO_CPPCLSANSANA','') = False));
  SetControlVisible('BVentil', (GetParamSocSecur('SO_CPPCLSANSANA','') = False));
{$ENDIF}
  SetControlEnabled('I_VENTILABLE', (Nat <> 'FI') and (fLeMode <> taConsult));
  // CA - 19/06/2003 - FQ 12379 - analytique toujours accessible
  {$IFDEF SERIE1}
  //XVI 25/04/2006 FQ 3978 analytoque monoventilation Début
  SetControlEnabled('BVentil', (GetField('I_VENTILABLE') = 'X') and VS1Cpta.Y_analytiqueAdvanced);
   //XVI 25/04/2006 FQ 3978 analytoque monoventilation fin
   //XVI 18/04/2006 FD 3978 S1 suite les même règles....
  {$ELSE}
  SetControlEnabled('BVentil', (GetField('I_VENTILABLE') = 'X'));
  {$ENDIF}
  EnableBoutonsInfos;
end;

procedure TOM_IMMO.EnableChampsOngletPlan;
var
  bOuvert, bSaisieDPI : boolean;
begin
  SetControlChecked('bReintegrationFiscale', FALSE);

  if (GetField('I_NATUREIMMO') = 'LOC') or (GetField('I_NATUREIMMO') = 'FI') then
    begin
    SetControlVisible('sBaseEcoOuDPI', False); // BTY 07/06 DPI
    exit;
    end;

  bOuvert := (GetField('I_ETAT') <> 'FER') and (GetField('I_ETAT') <> 'CLO')
    and (fLeMode <> taConsult) and (not fOperationEnCours)
    and (not fDPIUtilisee) and (not fComposant) ;  // BTY 07/06 DPI

  SetControlChecked('bAmortissementFiscal', not ((GetField('I_TAUXFISC') = 0.0)
    and (GetField('I_METHODEFISC') = '')));

  SetControlChecked('bReintegrationFiscale', not (GetField('I_REINTEGRATION') =
    0.0));
  SetControlEnabled('I_REINTEGRATION', GetCheckBoxState('bReintegrationFiscale')
    = cbChecked);

  SetControlEnabled('I_REPRISEFISCAL', bOuvert);
  SetControlEnabled('I_REPRISEECO', bOuvert);
  //chantier fiscal fq 17512
  SetControlEnabled('I_REPRISEDR', bOuvert);
  SetControlEnabled('I_REPRISEFEC', bOuvert);
  SetControlEnabled('bGestionFiscale', bOuvert);
  SetControlEnabled('bCalculAvecVNF', bOuvert);

  // TGA 26/01/2006
  IF GetField('I_NATUREIMMO')='CB' Then
    Begin
      SetControlVisible('tI_REPRISEDEP', False);
      SetControlVisible('I_REPRISEDEP', False);
      SetControlEnabled('I_REPRISEDEP', False);
    end
  ELSE
     SetControlEnabled('I_REPRISEDEP', bOuvert);

  SetControlEnabled('I_SURAMORT', bOuvert);
  SetControlEnabled('I_BASEECO', bOuvert);

  // BTY 07/06 Bloquer si Utilisation DPI
  //SetControlEnabled('I_QUOTEPART', True);
  SetControlEnabled('I_QUOTEPART', (not fDPIUtilisee));

  if not fOperationEnCours then bSaisieDPI := True
  else  bSaisieDPI := (GetField('I_QUANTITE')<>0);

  if fDateDebExoDPI <> iDate1900 then
     begin
     // On est derrière le tableau de suivi DPI
     SetControlVisible('sBaseEcoOuDPI', False);
     SetControlVisible('bDPI', True);
     SetControlEnabled('bDPI', bSaisieDPI and (not fDPIUtilisee) );
     end
  else
     if (GetField('I_METHODEECO') = 'NAM')  then
     // FQ 19013
       begin
       SetControlVisible('bDPI', False);
       SetControlEnabled('sBaseEcoOuDPI', False);
       SetControlVisible('SBaseEcoOuDPI', False);
       end
     else
     // Ouverture fiche normale
       begin
       SetControlVisible('bDPI', False);
       SetControlVisible('sBaseEcoOuDPI', True);
       SetControlEnabled('sBaseEcoOuDPI', (bOuvert));
       end;

  SetControlEnabled('bAmortissementFiscal', (bOuvert) and
    (GetField('I_NATUREIMMO') <> 'CB') and
    (GetField('I_METHODEECO') <> 'NAM') and
    (GetField('I_METHODEECO') <> 'VAR'));

  SetControlEnabled('AmortFisc', ((GetCheckBoxState('bAmortissementFiscal') = cbChecked) and (bOuvert)));

  SetControlEnabled('bReintegrationFiscale', bOuvert);
  SetControlEnabled('AmortEco', bOuvert);
  GereEco;
  GereFiscal;
  GereRepriseEco;
  GereRepriseFisc;
  GereSpinDureeEco(GetField('I_METHODEECO'));

  if (not (fLeMode in [taCreat..taCreatOne]) and (DS.State <> dsInsert)) then
  begin
    RecuperePlanRepriseTheorique(TQuery(DS), CumEcoSav, CumFiscSav);
  end;
end;

procedure TOM_IMMO.EnableChampsOngletInfo;
var
  stNature: string;
begin
  stNature := GetField('I_NATUREIMMO');
  SetControlVisible('I_DATEDERMVTECO', (stNature = 'PRO') or (stNature = 'CB'));
  SetControlVisible('tI_DATEDERMVTECO', (stNature = 'PRO') or (stNature =
    'CB'));
  SetControlVisible('I_DATEDERNMVTFISC', (stNature = 'PRO') or
    (GetField('I_METHODEFISC') <> ''));
  SetControlVisible('tI_DATEDERNMVTFISC', (stNature = 'PRO') or
    (GetField('I_METHODEFISC') <> ''));
end;

procedure TOM_IMMO.EnableChampsOngletContrat;
var
  stNature: string;
  bOuvert: boolean;
  bParamOuvert, Creation : boolean;
begin
  stNature := GetField('I_NATUREIMMO');
  if (stNature = 'PRO') or (stNature = 'FI') then
    exit;
  bOuvert := (GetField('I_ETAT') <> 'FER') and (GetField('I_ETAT') <> 'CLO')
    and (fLeMode <> taConsult) and (not fOperationEnCours)
    and (not fDPIUtilisee); // BTY 07/06 DPI

  bParamOuvert := (GetField('I_ETAT') <> 'FER')
    and (fLeMode <> taConsult)
    and (not fDPIUtilisee)
    and ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="'+GetField('I_IMMO')+'" AND I_OPELEVEEOPTION<>"X"') ;

  // FQ 21799
  Creation := ((fLeMode = taCreat) or (fLeMode = taCreatOne) or (fLeMode = taCreatEnSerie));

  SetControlEnabled('tCoutContrat', stNature <> 'LOC');
  SetControlEnabled('bDepotGarantie', (not ImmoDepotGarantieExiste) and
    bOuvert);
  ImmoDepotGarantieExiste := not (GetField('I_IMMOLIEGAR') = '');
  SetControlEnabled('VisuDepotGarantie', ImmoDepotGarantieExiste);
  SetControlChecked('bDepotGarantie', ImmoDepotGarantieExiste);
  if (GetControlEnabled('bDepotGarantie')) then
  begin
    if (GetCheckBoxState('bDepotGarantie') = cbChecked) then
      SetControlEnabled('GroupeDepotGarantie', bOuvert);
    SetControlEnabled('bDepotGarantie', bOuvert);
  end;
  SetControlEnabled('I_NUMCONTRATCB', bOuvert);
  //PGR - FQ 17056 28/11/2005 - ouverture systématique du champ Organisme
  SetControlEnabled('I_ORGANISMECB', true);
  SetControlEnabled('I_TYPELOYERCB', bOuvert);
  SetControlEnabled('I_PERIODICITE', bOuvert);
  SetControlEnabled('I_VERSEMENTCB', bOuvert);
  SetControlEnabled('I_RESIDUEL', bOuvert and (stNature <> 'LOC'));
  SetControlEnabled('tI_RESIDUEL', stNature <> 'LOC');
  SetControlEnabled('I_DATEDEBUTECH', bOuvert and (GetField('I_TYPELOYERCB') =
    'LCO'));
  //SetControlEnabled('I_DATEFINECH', bOuvert and (GetField('I_TYPELOYERCB') = 'LCO'));
  SetControlEnabled('I_MONTANTPREMECHE', bOuvert);
  //SetControlEnabled('I_MONTANTSUIVECHE', bOuvert);
  //SetControlEnabled('NbEcheances', bOuvert and (GetField('I_TYPELOYERCB') = 'LCO'));
  SetControlEnabled('cbI_DATESUSPCB', GetField('I_OPELEVEEOPTION') <> 'X');
  SetControlChecked('cbI_DATESUSPCB', (GetField('I_DATESUSPCB') > iDate1900));
  SetControlEnabled('I_DATESUSPCB', GetCheckBoxState('cbI_DATESUSPCB') =
    cbChecked);

  // FQ 20005
  SetControlEnabled('I_DATEFINECH', bParamOuvert and (GetField('I_TYPELOYERCB')='LCO'));
  SetControlEnabled('I_MONTANTSUIVECHE', (bParamOuvert) or Creation); // FQ 21799
  SetControlEnabled('NbEcheances', bParamOuvert and (GetField('I_TYPELOYERCB')='LCO'));
  // FQ 20005 Suite Ouvrir nb échéances en loyers variables
  if GetField('I_TYPELOYERCB')<>'LCO' then
     if fContrat <> nil then
        SetControlEnabled('NbEcheances', (bParamOuvert and (fContrat.ListeEcheances.Count > 0) ));
  // FQ 21799
  //SetControlEnabled('I_RESIDUEL', bParamOuvert and (stNature <> 'LOC'));
  SetControlEnabled('I_RESIDUEL', ((bParamOuvert or Creation) and (stNature <> 'LOC')));
  SetControlEnabled('SaisieEcheancier', bParamOuvert and (GetField('I_TYPELOYERCB')<>'LCO'));
  // FQ 21340 complément 20005 CB acquis et ayant subi levée option, griser le champ frais
  if (GetControlEnabled('I_FRAISECHE')) and
     (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb) and
     (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin) and
     (GetField('I_OPELEVEEOPTION') = 'X') then
     SetControlEnabled ('I_FRAISECHE', False);
end;

procedure TOM_IMMO.EnableChamps;
begin
  if (TTabSheet(GetActiveTabSheet('Pages')).TabVisible = False) then
    SetActiveTabSheet('PGeneral');
  EnableOngletsFiche;
  EnableChampsCommuns;
  EnableChampsOngletGeneral;
  EnableChampsOngletPlan;
  EnableChampsOngletInfo;
  EnableChampsOngletContrat;
  EnableChampsLibre;
  if fLeMode <> taConsult then
  begin
    GereEco;
    GereFiscal;
    GereDepotGarantie;
  end;
  GereRepriseEco;

  if bPlanFiscal then GereRepriseFisc;
  GereGroupeEcheance;
  // FQ 20005 vient d'être fait dans EnableChampsOngletContrat
  //SetControlEnabled('SaisieEcheancier', not
  //  GetControlEnabled('GroupeEcheance'));

  // BTY 09/06 FQ 16149
  //{$IFNDEF SERIE1}
  //if DS.State=dsInsert then PositionneEtabUser(THValComboBox(GetControl('I_ETABLISSEMENT')));
  //SetControlEnabled('I_ETABLISSEMENT',not VH^.ProfilUserC[prEtablissement].ForceEtab);
  //{$ENDIF}

  // BTY 08/06 Saisie par le bouton sBaseEcoOuDPI => I_BASEECO bloqué
  if GetControlEnabled('sBAseEcoOuDPI') then
     SetControlEnabled('I_BASEECO', False);
end;

procedure TOM_IMMO.InitEnregExistant;
begin
  { FQ 12378 - CA - 03/06/2003 - Création systématique du plan }
  if fPlanAmor <> nil then
    fPlanAmor.Free;
  fPlanAmor := TPlanAmort.Create(true);
  ChargeLog;
  EnableChamps;
  if DS.State <> dsInsert then
  begin
    if ((GetField('I_NATUREIMMO') = 'CB') or (GetField('I_NATUREIMMO') = 'PRO'))
      then
    begin
      fPlanAmor.Charge(TQuery(DS));
      fPlanAmor.Recupere(GetField('I_IMMO'), GetField('I_PLANACTIF'));
      //PGR 4/11/2005 - Pb affichage avec ou sans décimale
      //SetControlText('AmortDerog', GetField('I_REPRISEFISCAL') -
      //  GetField('I_REPRISEECO'));

      {chantier fiscal fq 17512
      // ajout pour fq 15274
      if (GetField('I_METHODEECO') = 'VAR') and
         (GetField('I_REPRISEFISCAL') < GetField('I_REPRISEECO')) then
          SetControlProperty('AmortDerog', 'Value', 0)
      else
         SetControlText('AmortDerog', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
           Getfield('I_REPRISEECO'), 13, V_PGI.OkDecV, '', TRUE));
      }
      //AfficheDerogatoire;
    end;
    if (GetField('I_METHODEECO')='VAR') then fPlanUO.Charge(string(GetField('I_IMMO')));
  end;
  AfficheMontantTTC;
  if ((GetField('I_NATUREIMMO') = 'CB') or (GetField('I_NATUREIMMO') = 'LOC'))
    then
  begin
    AfficheCoutContrat;
    AfficheNombreEcheance;
  end;
  if GetField('I_NATUREIMMO') = 'PRO' then
    SetControlText('COMPTEIMMO', GetField('I_COMPTEIMMO'));
  // MBO 25.07.05 FQ 16101 SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0) and (GetField('I_METHODEECO')='DEG'));
  SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0)); // FQ 21514 au lieu de ELIGILETP

  //Ajout mbo 29.06.07 + ajout test sur i_string1 pour après cloture 10.07.07
  if (GetField('I_STRING1')<> '') then
  begin
     SetControlText('TYPEBASEFISC', 'TR2');
  end;

end;

procedure TOM_IMMO.AfficheMontantTTC;
begin
  SetControlText('MontantTTC', StrFMontant(GetField('I_MONTANTHT') +
    GetField('I_TVARECUPERABLE'), 13, V_PGI.OkDecV, '', TRUE));
end;

procedure TOM_IMMO.AfficheCoutContrat;
var
  SommeDesLoyersTTC, dHTNet, dCoutTTC, dHT, dNet, SommeDesLoyersHT, dCoutFin:
  double;
  sTVADefault: string;
  Q: TQuery;
begin
  SommeDesLoyersHT := 0;
  if fContrat = nil then
  begin
    // if QFicheI_DATEPIECEA.AsString='' then exit;
    SetControlText('dtDebutContrat', GetField('I_DATEPIECEA'));
    SetControlText('dtFinContrat', GetField('I_DATEPIECEA'));
    SetControlText('CoutContrat', StrFMontant(0, 13, V_PGI.OkDecV, '', TRUE));
    SetControlText('CoutFinancementHT', StrFMontant(0, 13, V_PGI.OkDecV, '',
      TRUE));
  end
  else
  begin
    SetControlText('dtDebutContrat', DateToStr(fContrat.GetDateDebutContrat));
    SetControlText('dtFinContrat', DateToStr(fContrat.GetDateFinContrat));
    // BTY FQ 19151 I_DATEFINCB contient la bonne valeur si levée d'option
    // ou si immo dans l'historique
    // ajout mbo - fq 17339
    //if (fContrat.GetDateFinContrat <> GetField('I_DATEFINCB'))
    //and (GetField('I_DATEFINCB') <> iDate1900)
    if  (GetField('I_DATEFINCB') <> iDate1900)
    and ((GetField('I_OPELEVEEOPTION') = 'X') or (GetField('I_ETAT')='FER')) then
        SetControlText('dtFinContrat', GetField('I_DATEFINCB'));

    Q := OpenSQL('SELECT G_TVA FROM GENERAUX WHERE (G_GENERAL="' +
      GetField('I_COMPTELIE') + '")', TRUE);
    if not Q.EOF then
      sTVADefault := Q.FindField('G_TVA').AsString
    else
      sTVADefault := '';
    Ferme(Q);
    if (fContrat <> nil) then SommeDesLoyersHT := fContrat.SommeDesLoyers;
    dNet := GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
      GetField('I_TVARECUPEREE');
    dHTNet := CalculEchTva(TQuery(DS), sTVADefault, dNet, dHT);
    dCoutFin := SommeDesLoyersHT - dHTNet;

    SommeDesLoyersTTC := AffecteTvaEch(TQuery(DS), sTVADefault,
      SommeDesLoyersHT);
    dCoutTTC := SommeDesLoyersTTC + GetField('I_RESIDUEL');
    SetControlText('CoutContrat', StrFMontant(dCoutTTC, 13, V_PGI.OkDecV, '',
      TRUE));
    SetControlText('CoutFinancementHT', StrFMontant(dCoutFin, 13, V_PGI.OkDecV,
      '', TRUE));
  end;
end;

procedure TOM_IMMO.ChargeHistoriqueOperations;
var
  Info: string;
  i: integer;
begin
  THRichEditOLE(GetControl('I_HISTORIQUE')).Lines.Clear;
  if (GetField('I_ETAT') = 'OUV') then
    exit;
  for i := 0 to fLog.Detail.Count - 1 do
  begin
    if (fLog.Detail[i].GetValue('IL_DATEOP') < VHImmo^.Encours.Deb) then
    begin
      Info := DateToStr(fLog.Detail[i].GetValue('IL_DATEOP')) + ' - ' +
        RechDom('TIOPEAMOR', fLog.Detail[i].GetValue('IL_TYPEOP'), False) +
        ' - ' + fLog.Detail[i].GetValue('IL_TYPEMODIF');
      THRichEditOLE(GetControl('I_HISTORIQUE')).Lines.Add(Info);
    end;
  end;
end;

procedure TOM_IMMO.TraiteOrigineImmo;
var
  QLog: TQuery;
  bMutation: boolean;
  Where, Requete: string;
  NumMess: integer;
  i: integer;
  stMessage: string;
begin
  bMutation := false;
  SetControlVisible('lOrigineImmo', false);
  SetControlVisible('AccesImmoLie', false);
  stMessage := '';
  SetControlCaption('OrigineImmo', stMessage);
  Requete := 'SELECT IL_IMMO,IL_CODEECLAT,IL_CODECB FROM IMMOLOG ';
  // CA - 29/06/99 - Gestion sur 2 exercices
  Where := 'WHERE  (IL_DATEOP>="' + USDateTime(VHImmo^.Encours.Deb) +
    '" AND IL_DATEOP<="' + USDateTime(fDateFin) + '") AND (IL_CODECB="' +
    GetField('I_IMMO') + '" OR ' +
    'IL_CODEECLAT="' + GetField('I_IMMO') + '") ORDER BY IL_ORDRE DESC';
  QLog := OpenSQL(Requete + Where, TRUE);
  if not QLog.EOF then
  begin
    QLog.First;
    NumMess := -1;
    if QLog.FindField('IL_CODEECLAT').AsString <> '' then
      NumMess := 23
    else if QLog.FindField('IL_CODECB').AsString <> '' then
      NumMess := 24;
    if NumMess <> -1 then //EPZ 23/11/00
    begin
      stMessage := HM2[NumMess] + QLog.FindField('IL_IMMO').AsString;
      SetControlCaption('OrigineImmo', stMessage);
      if (stMessage <> '') then
        fCodeImmoLie := QLog.FindField('IL_IMMO').AsString;
    end;
    Ferme(QLog);
  end
  else
  begin
    // test si mutation
    NumMess := -1;
    Ferme(QLog);
    for i := 0 to fLog.Detail.Count - 1 do
    begin
      if (fLog.Detail[i].GetValue('IL_DATEOP') >= VHImmo^.Encours.Deb)
        and (fLog.Detail[i].GetValue('IL_DATEOP') <= fDateFin) then
      begin
        if fLog.Detail[i].GetValue('IL_CODEMUTATION') <> '' then
        begin
          fCodeImmoLie := fLog.Detail[i].GetValue('IL_CODEMUTATION');
          NumMess := 2;
          bMutation := true;
        end
        else if fLog.Detail[i].GetValue('IL_CODEECLAT') <> '' then
        begin
          fCodeImmoLie := fLog.Detail[i].GetValue('IL_CODEECLAT');
          NumMess := 3;
        end
        else if fLog.Detail[i].GetValue('IL_CODECB') <> '' then
        begin
          fCodeImmoLie := fLog.Detail[i].GetValue('IL_CODECB');
          NumMess := 4;
        end;
      end;
    end;
    if NumMess <> -1 then
    begin
      stMessage := HM2[NumMess] + fCodeImmoLie;
      SetControlCaption('OrigineImmo', stMessage);
    end;
  end;
  if GetField('I_IMMOLIEGAR') <> '' then
  begin
    if GetField('I_NATUREIMMO') = 'FI' then
      NumMess := 6
    else if (GetField('I_NATUREIMMO') = 'CB') then
      NumMess := 5
    else if (GetField('I_NATUREIMMO') = 'LOC') then
      NumMess := 5;
    stMessage := HM2[NumMess] + ' ' + GetField('I_IMMOLIEGAR');
    SetControlCaption('OrigineImmo', stMessage);
    if (stMessage <> '') then
      fCodeImmoLie := GetField('I_IMMOLIEGAR');
  end;
  SetControlVisible('lOrigineImmo', (stMessage <> '') and (not bMutation));
  SetControlVisible('AccesImmoLie', GetControlVisible('lOrigineImmo'));
end;

procedure TOM_IMMO.GereSuramortissement;
begin
  // CA - 14/05/2002 - Suramortissement
  SetControlVisible('I_SURAMORT', (GetField('I_DATEPIECEA') >=
    StrToDate('17/10/2001'))
    and (GetField('I_DATEPIECEA') <= StrToDate('31/03/2003'))
    and ((GetField('I_METHODEECO') = 'DEG') or
    (GetField('I_METHODEFISC') = 'DEG')));
  if not GetControlVisible('I_SURAMORT') then
    SetControlChecked('I_SURAMORT', False);
  // CA - 14/05/2002
end;

procedure TOM_IMMO.EnableBoutonsInfos;
var
  stNature: string;
  //okplan : boolean;
begin
  stNature := GetField('I_NATUREIMMO');
  { Modif mbo suite gestion prime : le bouton plan a maintenant un menu déroulant
    l'ouverture des options est géré dans onPopZoomPlan
  SetControlEnabled('BPlanAmort', ImmoAmortie(GetField('I_METHODEECO'),
    GetField('I_METHODEFISC'), GetField('I_NATUREIMMO')));
  }
  SetControlEnabled('BPlanRemboursement', false);

  if (stNature = 'CB') or (stNature = 'LOC') then
    SetControlEnabled('BPlanRemboursement', (fContrat <> nil) and
      (fContrat.ListeEcheances <> nil) and (FContrat.ListeEcheances.Count > 0));

  if stNature = 'CB' then
    SetControlProperty('BPlanAmort', 'Hint', HM2[8])
  else
    SetControlProperty('BPlanAmort', 'Hint', HM2[7]);

  // ajout mbo - fq 15274 - zone antérieurs fiscaux non accessible en mode VARIABLE
  //chantier fiscal fq 17512
  //IF GetField('I_METHODEECO')= 'VAR' then
  //   SetControlEnabled('I_REPRISEFISCAL', false);
  GerePlanVar;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Particularité du plan fiscal si eco = VAR V8.00
Suite ........ : Si coche gestion fiscale non cochée
Suite ........ :     le plan fiscal est non accessible (linéaire)
Suite ........ : sinon
Suite ........ :     le plan fiscal est ouvert
Mots clefs ... :
*****************************************************************}
Procedure TOM_IMMO.GerePlanVar;

var
   sEtat : string;
   bOuvert : boolean;

begin
  if (GetField('I_METHODEECO') = 'VAR') then
  begin
    SetControlChecked('bAmortissementFiscal', True);
    SetControlEnabled('bAmortissementFiscal', false);


    TraiteDotationAnterieures;
    AfficheDerogatoire;

    sEtat := GetField('I_ETAT');
    bOuvert := (sEtat <> 'CLO') and (sEtat <> 'FER') and (sEtat <> 'SUS')
    and (fLeMode <> taConsult) and (not fOperationEnCours) and (not
    ((GetField('I_IMMOLIEGAR') <> '') and (GetField('I_NATUREIMMO') = 'FI')))
    and (not fDPIUtilisee) ; // BTY 07/06 DPI

    SetControlEnabled('I_METHODEFISC', bOuvert);
    SetControlEnabled('I_DUREEFISC', bOuvert);
    SetControlEnabled('I_DATEDEBFIS', bOuvert);
    SetControlEnabled('TYPEBASEFISC', bOuvert);
    SetControlEnabled('I_BASEFISC', bOuvert);
    SetControlEnabled('I_REPRISEFISC', bOuvert);
    SetControlEnabled('I_REPRISEDR', bOuvert);
    SetControlEnabled('I_REPRISEFEC', bOuvert);
    SetControlEnabled('bGestionFiscale', bOuvert);
    SetControlEnabled('bCalculAvecVNF', bOuvert);
    SetControlEnabled('AmortFisc', bOuvert);
    SetControlEnabled('GBreprisef', bOuvert);

  end;
end;


procedure TOM_IMMO.GereFiscal;
var (*couleur : TColor;*)
  bEnable: boolean;
  bOuvert : boolean;
  sEtat : string;
begin
  if GetField('I_METHODEECO') = 'VAR' then
  begin
     SetControlChecked('bAmortissementFiscal', true);
     SetControlEnabled('bAmortissementFiscal', false);
  end;

  sEtat := GetField('I_ETAT');
  bOuvert := (sEtat <> 'CLO') and (sEtat <> 'FER') and (sEtat <> 'SUS')
             and (fLeMode <> taConsult) and (not fOperationEnCours) and (not
             ((GetField('I_IMMOLIEGAR') <> '') and (GetField('I_NATUREIMMO') = 'FI')))
             and (not fDPIUtilisee) and (not fComposant) ; // BTY 07/06 DPI

  bEnable := (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and bOuvert;

  SetControlEnabled('AmortFisc', bEnable);
  SetControlEnabled('I_METHODEFISC', bEnable);

  //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
 if GetCheckBoxState('bAmortissementFiscal') = cbChecked then
    begin
      if GetField('I_BASEFISC') = GetField('I_BASEECO') then
        SetControlText('TYPEBASEFISC','ECO')
      else
        SetControlText('TYPEBASEFISC', 'THE');
    end
  else
    begin
      SetControlText('TYPEBASEFISC', '');
      SetControlText('I_BASEFISC', StrFMontant(0, 13, V_PGI.OkDecV, '', TRUE));
      SetField('I_BASEFISC',0);
    end;

  SetControlEnabled('TYPEBASEFISC', bEnable);

  SetControlEnabled('I_DUREEFISC', bEnable);
  SetControlEnabled('I_BASEFISC', false);

  //ajout mbo fq 17569
  // FQ 20555 - différence de comportement entre cwas et 2/3 - ajout du test sur ds.STATE
  if (bEnable) and (DS.State=dsInsert) then
  begin
      if GetField('I_METHODEFISC') = 'DEG' then
      begin
        SetField('I_DATEDEBFIS', GetField('I_DATEPIECEA'));
        SetControlEnabled('I_DATEDEBFIS', false);
      end else
      begin
        SetField('I_DATEDEBFIS', GetField('I_DATEAMORT'));
        SetControlEnabled('I_DATEDEBFIS', true);
      end;
  end;

  if not bEnable then
      SetControlEnabled('I_DATEDEBFIS', false);

end;

procedure TOM_IMMO.GereEco;
(*var couleur : TColor;*)
var
  stMethode: string;
  bOuvert: boolean;
begin
  stMethode := GetField('I_METHODEECO');

  SetControlEnabled('I_DUREEECO', GetControlEnabled('AmortEco') and (not
    (stMethode = 'NAM')));

  // BTY 08/06 Saisie par le bouton sBaseEcoOuDPI => I_BASEECO bloqué
  if GetControlEnabled('sBAseEcoOuDPI') then
     SetControlEnabled('I_BASEECO', False)
  else //
     SetControlEnabled('I_BASEECO', GetControlEnabled('AmortEco') and (not
        (stMethode = 'NAM')));
  SetControlVisible('I_UNITEOEUVRE', (stMethode = 'VAR'));
  SetControlVisible('I_INFOUO', (stMethode = 'VAR'));
  SetControlVisible('TI_UNITEOEUVRE', (stMethode = 'VAR'));
  SetControlVisible('I_TAUXECO', not (stMethode = 'VAR'));
  SetControlVisible('TI_TAUXECO', not (stMethode = 'VAR'));

  //PGR 26/10/2005 Ajout coche de calcul du plan futur avec la VNC
  bOuvert := (GetField('I_ETAT') <> 'FER') and (GetField('I_ETAT') <> 'CLO')
    and (fLeMode <> taConsult) and (not fOperationEnCours)
    and (not fDPIUtilisee) and (not fComposant) ; // BTY 07/06 DPI

  SetControlEnabled('bCalculAvecVnc', (bOuvert) and
    ((stMethode = 'LIN') or (stMethode = 'VAR')));

  if DS.state in [dsBrowse] then
  begin
    SetControlChecked('bCalculAvecVnc', (GetField('I_JOURNALA') = '***'));
    //ajout chantier fiscal fq 17512
    SetControlChecked('bCalculAvecVNF', (GetField('I_FUTURVNFISC') = '***'));
    SetControlChecked('bGestionFiscale', (GetField('I_NONDED') = 'X'));
  end;

  // ajout mbo fq 17569
  SetControlEnabled('I_DATEDEBECO', GetControlEnabled('AmortEco') and (not
    (stMethode = 'NAM')));
end;

{ mbo fq 17569 procedure TOM_IMMO.GereReprise remplacée par 2 procedures : GereRepriseEco et GereRepriseFisc}

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 05/05/2006
Modifié le ... :   /  /
Description .. : fonction gérant le groupbox antérieurs éco
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.GereRepriseEco;
var
  RepriseVisible: boolean;
  PremMois, PremAnnee, NbMois: Word;
  nJour: integer;
  bDerog: boolean;
  datedebut: integer;  // ajout mbo 17120
begin
  if VarIsNull(GetField('I_DATEDEBECO')) then exit;  // FQ 15080 - CA - 17/12/2004 - Au cas où sous ORACLE, on en continue pas le traitement
  RepriseVisible := RepriseEcoVisible;
  NOMBREMOIS(GetField('I_DATEDEBECO'), VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);

  datedebut := GetField('I_DATEDEBECO');

  if ((NbMois - 1) > GetField('I_DUREEECO')) then
    nJour := GetField('I_DUREEECO') * 30
  else
    // fq 17120 - mbo - 6.12.2005 nJour := StrToInt(FloatToStr(VHImmo^.Encours.Deb - GetField('I_DATEAMORT')));
    nJour := NombreJour360(datedebut,VHImmo^.Encours.Deb-1);

  if nJour < 0 then nJour := 0;
  NOMBREMOIS(datedebut, VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);

  if ((NbMois - 1) > GetField('I_DUREEECO')) then
     NbMois := GetField('I_DUREEECO') + 1;

  SetControlCaption('tDUREEAMORTIE', Format(HM2[9], [NbMois - 1, nJour]));
  SetControlVisible('tDUREEAMORTIE', RepriseVisible);
  SetControlVisible('GBRepriseE', RepriseVisible);

  // TGA 26/01/2006   03/03/06 ajout test immo non cb
   if (GetField('I_NATUREIMMO') <> 'CB') then
   Begin
      SetControlVisible('I_REPRISEDEP', RepriseVisible);
      SetControlVisible('tI_REPRISEDEP', RepriseVisible);
   End;

  // gestion du gb antérieurs fiscaux
  bDerog := (bPlanFiscal) and (RepriseFiscVisible);
  //if bDerog then
    {chantier fiscal fq 17512
    // ajout pour fq 15274
    if (GetField('I_METHODEECO') = 'VAR') and
       (GetField('I_REPRISEFISCAL') < GetField('I_REPRISEECO')) then
       SetControlProperty('AmortDerog', 'Value', 0)
    else
       SetControlText('AmortDerog', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
         Getfield('I_REPRISEECO'), 13, V_PGI.OkDecV, '', TRUE));
    }
    //AfficheDerogatoire;
  SetControlVisible('GBrepriseF', bDerog);

  EnableBoutonsInfos;

end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 05/05/2006
Modifié le ... :   /  /
Description .. : gestion du GB antérieurs fiscaux
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.GereRepriseFisc;
var
  RepriseVisible: boolean;
  PremMois, PremAnnee, NbMois: Word;
  nJour: integer;
  bDerog: boolean;
  datedebut: integer;  // ajout mbo 17120
begin
  if not (bPlanFiscal) then exit;
  if VarIsNull(GetField('I_DATEDEBFIS')) then exit;  // FQ 15080 - CA - 17/12/2004 - Au cas où sous ORACLE, on en continue pas le traitement
  RepriseVisible := RepriseFiscVisible;
  NOMBREMOIS(GetField('I_DATEDEBFIS'), VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);

  datedebut := GetField('I_DATEDEBFIS');

  if ((NbMois - 1) > GetField('I_DUREEFISC')) then
    nJour := GetField('I_DUREEFISC') * 30
  else
    // fq 17120 - mbo - 6.12.2005 nJour := StrToInt(FloatToStr(VHImmo^.Encours.Deb - GetField('I_DATEAMORT')));
    nJour := NombreJour360(datedebut,VHImmo^.Encours.Deb-1);

  if nJour < 0 then nJour := 0;
  NOMBREMOIS(datedebut, VHImmo^.Encours.Deb, PremMois, PremAnnee,NbMois);

  if ((NbMois - 1) > GetField('I_DUREEFISC')) then
     NbMois := GetField('I_DUREEFISC') + 1;

  SetControlCaption('tDURAMORTF', Format(HM2[9], [NbMois - 1, nJour]));
  SetControlVisible('tDURAMORTF', RepriseVisible);
  SetControlVisible('GBRepriseF', RepriseVisible);

  // gestion du Dérogatoire
  bDerog := RepriseVisible;
  if bDerog then
  begin
    {chantier fiscal fq 17512
    // ajout pour fq 15274 - mbo
    MontantDerog := Arrondi((GetField('I_REPRISEFISCAL') -
                    GetField('I_REPRISEECO')), V_PGI.OkDecV);
    if (GetField('I_METHODEECO') = 'VAR') and (MontantDerog < 0) then
       MontantDerog := 0;

    SetControlText('AmortDerog', StrFMontant(MontantDerog, 13, V_PGI.OkDecV, '', TRUE));
    }
    //AfficheDerogatoire;
  end;
  SetControlVisible('GBrepriseF', bDerog);

  EnableBoutonsInfos;

end;


procedure TOM_IMMO.GereSpinDureeEco(Methode: string);
begin
  if Methode = 'NAM' then
  begin
    if GetField('I_DUREEECO') <> 0 then
    begin
      if DS.State in [dsBrowse] then
        DS.Edit;
      SetField('I_DUREEECO', 0);
    end;
    SetControlProperty('I_DUREEECO', 'MinValue', 0);
    SetControlProperty('I_DUREEECO', 'MaxValue', 0);
    SetControlChecked('bAmortissementFiscal', false);
  end
  else
  begin
    if GetField('I_DUREEECO') = 0 then
      SetField('I_DUREEECO', 1);
    SetControlProperty('I_DUREEECO', 'MinValue', 1);
    SetControlProperty('I_DUREEECO', 'MaxValue', 1200);
  end;
end;


// fq 20555 - mbo - 19.07.07 - permet de rétablir le focus ds la fenêtre
procedure TOM_IMMO.GereSpinDureeFisc(Methode: string);
begin
  if GetField('I_DUREEFISC') <= 1 then
    SetField('I_DUREEFISC', 1);

  //SetControlProperty('I_DUREEFISC', 'MinValue', 1);
  //SetControlProperty('I_DUREEFISC', 'MaxValue', 1200);

end;


// FQ 19013
procedure TOM_IMMO.EnableBoutonBaseEco;
var
  bOuvert: boolean;
begin
  bOuvert := (GetField('I_ETAT') <> 'FER') and (GetField('I_ETAT') <> 'CLO')
              and (fLeMode <> taConsult) and (not fOperationEnCours)
              and (not fDPIUtilisee) and (not fComposant);

  if fDateDebExoDPI <> iDate1900 then
     begin
     // On est derrière le tableau de suivi DPI
     SetControlVisible('sBaseEcoOuDPI', False);
     end
  else
     // Immo non amortissable
     if (GetField('I_METHODEECO') = 'NAM')  then
       begin
       SetControlVisible('SBaseEcoOuDPI', False);
       SetControlEnabled('sBaseEcoOuDPI', False);
       end
     else
     // Ouverture fiche normale
       begin
       SetControlVisible('sBaseEcoOuDPI', True);
       SetControlEnabled('sBaseEcoOuDPI', (bOuvert));
       end;

  if GetControlEnabled('sBAseEcoOuDPI') then
     SetControlEnabled('I_BASEECO', False);
end;


procedure TOM_IMMO.GereDepotGarantie;
var (*couleur : TColor;*)
  bDepotGarantie: boolean;
begin
  bDepotGarantie := (GetCheckBoxState('bDepotGarantie') = cbChecked) and (not
    ImmoDepotGarantieExiste);
  SetControlEnabled('GroupeDepotGarantie', bDepotGarantie);
  SetControlEnabled('bDepotGarantie', (not (GetCheckBoxState('bDepotGarantie') =
    cbChecked)) or (not ImmoDepotGarantieExiste));
  (*     if GroupeDepotGarantie.enabled then couleur:=clBlack
  else couleur:=clGrayText;
  tI_DEPOTGARANTIE.Font.Color:=couleur;
  tI_DATEDEPOTGAR.Font.Color:=couleur;
  tI_IMMOLIEGAR.Font.Color:=couleur;
  I_DEPOTGARANTIE.Font.Color:=couleur;
  I_DATEDEPOTGAR.Font.Color:=couleur;
  I_IMMOLIEGAR.Font.Color:=couleur;
  COMPTE_DEPOT.Font.Color:=couleur;
  tCOMPTE_DEPOT.Font.Color:=couleur;*)
  SetControlEnabled('I_IMMOLIEGAR', bDepotGarantie);
  SetControlEnabled('I_DEPOTGARANTIE', bDepotGarantie);
  SetControlEnabled('I_DATEDEPOTGAR', bDepotGarantie);
  SetControlEnabled('COMPTE_DEPOT', bDepotGarantie);
  SetControlEnabled('VisuDepotGarantie', ImmoDepotGarantieExiste);
end;

procedure TOM_IMMO.GereGroupeEcheance;
begin
  SetControlEnabled('GroupeEcheance', (GetField('I_TYPELOYERCB') = 'LCO'));
  (*  if GroupeEcheance.enabled then couleur:=clBlack else couleur:=clGrayText;
    for i:=0 to GroupeEcheance.ControlCount - 1 do
    begin
      if GroupeEcheance.Controls[i].ClassName='THDBEdit' then
        THDBEdit(GroupeEcheance.Controls[i]).Font.Color:=couleur
      else THLabel(GroupeEcheance.Controls[i]).Font.Color:=couleur;
    end;*)
end;

procedure TOM_IMMO.ChargeLog;
var
  Q: TQuery;
  i: integer;
begin
  fOperationEnCours := False;
  fDPIUtilisee := (GetField('I_DPI')='X');  // BTY 07/06 DPI

  //Ajout mbo pour la gestion du remplacement de composant
  fComposant := ((GetField('I_STRING1')<> '') or (GetField('I_REMPLACEE')<> ''));


  fLog.ClearDetail;
  Q := OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="' + GetField('I_IMMO') +
    '" ORDER BY IL_DATEOP', True);
  fLog.LoadDetailDB('IMMOLOG', '', '', Q, False);
  Ferme(Q);
  if fLog.Detail.Count > 0 then
  begin
    for i := 0 to fLog.Detail.Count - 1 do
    begin
      if (fLog.Detail[i].GetValue('IL_DATEOP') >= VHImmo^.Encours.Deb)
        and (fLog.Detail[i].GetValue('IL_TYPEOP') <> 'ACQ')
        and (fLog.Detail[i].GetValue('IL_TYPEOP') <> 'CLO') then
      begin
        fOperationEnCours := True;
        break;
      end;
    end;
  end;
end;

procedure TOM_IMMO.AfficheNombreEcheance;
begin
  if (GetField('I_TYPELOYERCB') = 'LVA') and (fContrat.ListeEcheances.Count <> 0)
    then
  begin
    if (fContrat <> nil) and (not bChargementEnCours) then
    begin
      SetField('I_DATEDEBUTECH', fContrat.GetDateDebutEcheance);
      SetField('I_DATEFINECH', fContrat.GetDateFinEcheance);
    end;
  end;
  if fContrat <> nil then
    SetControlProperty('NbEcheances', 'Value', fContrat.ListeEcheances.Count)
  else
    SetControlProperty('NbEcheances', 'Value', 0);
end;

{ mbo fq 17569 fonction function TOM_IMMO.isRepriseVisible remplacée par 2 fonctions RepriseEcoVisible et RepriseFiscVisible }

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 05/05/2006
Modifié le ... :   /  /
Description .. : fonction permettant de gérer le GB antérieurs économiques
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.RepriseEcoVisible: boolean;
begin
    result:= (VHImmo^.Encours.Deb > GetField('I_DATEDEBECO')) and
             (GetField('I_ETAT') = 'OUV') and (GetField('I_METHODEECO') <> 'NAM');
    // ajout mbo nam fq 17569
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 05/05/2006
Modifié le ... :   /  /
Description .. : fonction permettant de gérer le GB des antérieurs fiscaux
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.RepriseFiscVisible: boolean;
begin
    result := (bPlanFiscal) and
              (VHImmo^.Encours.Deb > GetField('I_DATEDEBFIS')) and
              (GetField('I_ETAT') = 'OUV') and (GetField('I_METHODEECO') <> 'NAM');

end;


procedure TOM_IMMO.AfficheBaseTaxe;
var
  Base: double;
  PrefCompte: string;
begin
  // Calcul des base Taxe pro, base eco. et base fisc.
  if DS.state in [dsEdit, dsInsert] then
  begin
    Base := GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
      GetField('I_TVARECUPEREE');
    if (GetField('I_NATUREIMMO') <> 'LOC') then
    begin
      // si compte de la classe 20,26 ou 27, base TP par défaut=0.0
      PrefCompte := Copy(GetField('I_COMPTEIMMO'), 1, 2);
      if ((PrefCompte = '20') or (PrefCompte = '26') or (PrefCompte = '27'))
        then
        SetField('I_BASETAXEPRO', 0)
      else
        SetField('I_BASETAXEPRO', Base);
        // MBO 25.07.05 FQ 16101 SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0) and (GetField('I_METHODEECO')='DEG'));
        SetControlEnabled('ELIGIBLETP',(GetField('I_BASETAXEPRO')<>0));
    end;
    GereNatureBien;

     // fq 18218 on ne réinitialise que si base éco est supérieur au prix d'achat
    // fq 18330 - mbo -07.06.2006 if fLeMode = taModif then
    if DS.state in [dsEdit] then
    begin
     PGIInfo(MSG_MODIFMONTANTHT);
     if GetField('I_BASEECO') > Base then SetField('I_BASEECO', Base)
    end
    else
    begin
     SetField('I_BASEECO', Base);
    end;

    //ajout mbo test sur reintegation FQ 15625   21.07.05
    IF GetField('I_MONTANTHT')<= GetField('I_REINTEGRATION') then
    begin
       SetField('I_REINTEGRATION', 0);
       //SetField('I_QUOTEPART', 0);
       SetControlChecked('bReintegrationFiscale', FALSE); // ajout mbo 16.06.06
    end;

    if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
      SetField('I_BASEFISC', Base);
    SetField('I_MONTANTBASEAMORT', Base);
  end;
end;

procedure TOM_IMMO.AfficheTauxEco;
begin
  // mbo fq 17569
  if DS.state in [dsEdit, dsInsert] then
  begin
    //if ((not VarIsNull (GetField('I_METHODEECO')))
    //   and (not VarIsNull (GetField('I_DATEAMORT'))) and (not VarIsNull (GetField('I_DATEPIECEA'))) and (not VarIsNull (GetField('I_DUREEECO')))) then  // FQ 15080 - CA - 17/12/2004 - Au cas où sous ORACLE

    if ((not VarIsNull (GetField('I_METHODEECO')))
       and (not VarIsNull (GetField('I_DATEDEBECO'))) and (not VarIsNull (GetField('I_DUREEECO')))) then  // FQ 15080 - CA - 17/12/2004 - Au cas où sous ORACLE

    begin
      SetField('I_TAUXECO', GetTaux(GetField('I_METHODEECO')
        , GetField('I_DATEDEBECO')
        , Getfield('I_DATEDEBECO')
        , GetField('I_DUREEECO')));
    end;
  end;
end;

procedure TOM_IMMO.AfficheTauxFisc;
begin
  // en attente mbo 09.05.06
  // ne rien modifier car le taux en dégressif diffère selon la date d'achat
  // question posée à CJ

  if ((DS.state in [dsEdit, dsInsert]) and
    (GetCheckBoxState('bAmortissementFiscal') = cbChecked)) then
  begin
   {MBO FQ 17569 if (not VarIsNull(GetField('I_DATEAMORT'))) and (not VarIsNull(GetField('I_DATEPIECEA')))
      and (not VarIsNull(GetField('I_DUREEFISC'))) then }

    if (not VarIsNull(GetField('I_DATEDEBFIS')))
      and (not VarIsNull(GetField('I_DUREEFISC'))) then
    begin
        SetField('I_TAUXFISC', GetTaux(GetField('I_METHODEFISC')
          , GetField('I_DATEDEBFIS')
          , Getfield('I_DATEDEBFIS')
          , GetField('I_DUREEFISC')));
    end;
  end;
end;

procedure TOM_IMMO.TraiteDotationAnterieures;
var
  TmpPlan: TPlanAmort;
  PremMois, PremAnnee, NbMois: Word;
  EcoOk, FiscalOk, presenceFiscal: boolean;
begin
  { fq 21153 _ mbo - 20/07/07
  EcoOk := (GetField('I_METHODEECO') <> '');
  FiscalOk := ((GetField('I_METHODEFISC') <> '') and
    (GetCheckBoxState('bAmortissementFiscal') = cbChecked));
  }
  EcoOk := ((GetField('I_METHODEECO') <> '') and (GetField('I_DUREEECO')<= 1200));
  FiscalOk := ((GetField('I_METHODEFISC') <> '') and
               (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
               (GetField('I_DUREEFISC') <= 1200)) or
               (GetCheckBoxState('bAmortissementFiscal') <> cbChecked);

  PresenceFiscal := ((GetField('I_METHODEFISC') <> '') and
               (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
               (GetField('I_DUREEFISC') <= 1200));

  if (EcoOk) and (FiscalOk) then
  begin
    NOMBREMOIS(GetField('I_DATEAMORT'), VHImmo^.Encours.Deb, PremMois,
      PremAnnee, NbMois);
    SetField('I_DUREEREPRISE', IntToStr(MaxIntValue([0, NbMois - 1])));
    TmpPlan := TPlanAmort.Create(true);
    try
      TmpPlan.Charge(TQuery(DS));
      // ajout mbo 16/02/06 pour un recalcul théorique sans tenir cpte antérieur dépréciation
      TmpPlan.AmortEco.RepriseDep := 0;
      TmpPlan.AmortFisc.RepriseDep := 0;
      // fin ajout
      TmpPlan.Suramort := (GetCheckBoxState('I_SURAMORT') = cbChecked);
      if (GetField('I_METHODEECO')='VAR') then
      begin
         // FQ 20660 ajout test fplanUO <> nil - mbo 11.06.07 -
         // le chargeplanUO plantait en debug
         if fPlanUO <> nil then
            TmpPlan.ChargePlanUO ( fPlanUO.LePlan );

      end;
      // Sinon, sur check de la zone, suramort pas mis à jour
      TmpPlan.CalculReprises;
      CumEcoSav := TmpPlan.AmortEco.Reprise;
      CumFiscSav := TmpPlan.AmortFisc.Reprise;
      SetField('I_DATEDERMVTECO', TmpPlan.GetDateFinAmortEx(TmpPlan.AmortEco));
      SetField('I_DATEDERNMVTFISC',
        TmpPlan.GetDateFinAmortEx(TmpPlan.AmortFisc));

      // ajout pour chantier fiscal fq 17512
      // 20.07.07 if (FiscalOk) then
      if presenceFiscal then
      begin
         if (GetField('I_METHODEECO')<>'VAR') then
         begin

            if (TmpPlan.Determine_gestion_fiscale(tQuery(DS),(CumEcoSav+GetField('I_REPRISEDEP')),
                CumFiscSav)) = true then
            begin
               SetControlChecked('bGestionFiscale',true);
               SetControlEnabled('bGestionFiscale', false);
            end else
               SetControlEnabled('bGestionFiscale', true);

         end else
         begin
             SetControlEnabled('bGestionFiscale', true);
         end;
      end else
          SetControlEnabled('bGestionFiscale', false);
    finally
      TmpPlan.free;
    end;
  {
  end else
  begin
     if GetField('I_DUREEECO') > 1200 then
     begin
        PGIBox(HMC[51]);
        SetActiveTabSheet('PPlan');
        SetFocusControl('I_DUREEECO');
     end else
     begin
        if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
           (GetField('I_DUREEFISC') > 1200) then
        begin
            PGIBox(HMC[51]);
            SetActiveTabSheet('PPlan');
            SetFocusControl('I_DUREEFISC');
        end;
     end;   }
  end;
end;

procedure TOM_IMMO.GereNatureBien;
begin
  SetControlEnabled('I_NATUREBIEN', (not (GetField('I_BASETAXEPRO') = 0) and
    (GetField('I_NATUREIMMO') <> 'LOC')));
  SetControlEnabled('tI_NATUREBIEN', (not (GetField('I_BASETAXEPRO') = 0) and
    (GetField('I_NATUREIMMO') <> 'LOC')));
end;

function TOM_IMMO.ControleCode: integer;
var
  stCode: string;
begin
  Result := 0;
  if ((fLeMode in [taCreat..taCreatOne]) and (DS.State = dsInsert)) then
  begin
    stCode := GetField('I_IMMO');
    if ImExisteCarInter(stCode) then
      Result := 1
    else if ExisteSQL('SELECT I_IMMO FROM IMMO WHERE I_IMMO="' + stCode + '"')
      then
      Result := 2
    else if ExisteSQL('SELECT I_CHANGECODE FROM IMMO WHERE I_CHANGECODE="' +
      stCode + '"') then
      Result := 3;
  end;
  if Result <> 0 then
  begin
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_IMMO');
  end;
  if (GetField('I_LIBELLE') = '') then
  begin
    Result := 31;
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_LIBELLE');
  end;
end;

function TOM_IMMO.ControleComptes: integer;
var
  stNature: string;
  stField: string;
begin
  Result := 0;
  stNature := GetField('I_NATUREIMMO');
  if GetField('I_COMPTEIMMO') = '' then
  begin
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_COMPTEIMMO');
    Result := 4
  end
  else if (not ControleCompteImmo(GetField('I_COMPTEIMMO'), stNature)) then
  begin
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_COMPTEIMMO');
    Result := 5;
  end
  else if (GetControlEnabled('I_COMPTELIE') and (not
    ControleCompteCharge(GetField('I_COMPTELIE'), stNature))) then
  begin
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_COMPTELIE');
    Result := 6;
  end
  else if (((stNature = 'LOC') or (stNature = 'CB')) and (not
    ControleOrganisme(GetfIeld('I_ORGANISMECB')))) then
  begin
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_ORGANISMECB');
    Result := 7;
  end
  else if (stNature = 'PRO') then
  begin
    stField := '';
    if (not VerifCompteAssocie(GetControl('I_COMPTEAMORT'))) then
      stField := 'I_COMPTEAMORT'
    else if (not VerifCompteAssocie(GetControl('I_COMPTEDOTATION'))) then
      stField := 'I_COMPTEDOTATION'
    else if (not VerifCompteAssocie(GetControl('I_REPEXPLOIT'))) then
      stField := 'I_REPEXPLOIT'
    else if (not VerifCompteAssocie(GetControl('I_DOTATIONEXC'))) then
      stField := 'I_DOTATIONEXC'
    else if (not VerifCompteAssocie(GetControl('I_REPEXCEP'))) then
      stField := 'I_REPEXCEP'
    else if (not VerifCompteAssocie(GetControl('I_VACEDEE'))) then
      stField := 'I_VACEDEE'
    else if (not VerifCompteAssocie(GetControl('I_AMORTCEDE'))) then
      stField := 'I_AMORTCEDE'
    else if (not VerifCompteAssocie(GetControl('I_VAOACEDEE'))) then
      stField := 'I_VAOACEDEE'
    else if (not VerifCompteAssocie(GetControl('I_COMPTEDEROG'))) then
      stField := 'I_COMPTEDEROG'
    else if (not VerifCompteAssocie(GetControl('I_REPRISEDEROG'))) then
      stField := 'I_REPRISEDEROG'
    else if (not VerifCompteAssocie(GetControl('I_PROVISDEROG'))) then
      stField := 'I_PROVISDEROG';
    if stField <> '' then
    begin
      SetActiveTabSheet('PComptable');
      SetFocusControl(stField);
      Result := 8;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 11/10/2004
Modifié le ... :   /  /
Description .. : 11/10/2004 - CA - FQ 14773 : contrôle de la base
Suite ........ : d'amortissement pa rapport au montant HT
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.ControleChamps: integer;
var
  //Exo: TImExoDate;
  stMsg: string;
  stMethode: string;
  wBaseEco : double;
begin
  Result := 0;
  wBaseEco := Arrondi (GetField('I_BASEECO'), V_PGI.OkDecV);
  if (GetField('I_LIEUGEO') = '') then
  begin
    Result := 9;
    SetActiveTabSheet('PComptable');
    SetFocusControl('I_LIEUGEO');
  end
  else if (not PlanAmortValide(stMethode)) then
  begin
    Result := 10;
    SetActiveTabSheet('PPlan');
    SetFocusControl(stMethode);
  end else if ((GetField('I_METHODEECO')='VAR') and (GetField('I_UNITEOEUVRE')=0)) then
  begin
    Result := 35;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_UNITEOEUVRE');
  end else if ((GetField('I_METHODEECO')='VAR') and (not fPlanUO.Valide)) then
  begin
    Result := 36;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_UNITEOEUVRE');
  //fq 17569 end else if ((GetField('I_METHODEECO')='VAR') and (GetField('I_DATEPIECEA')<VHImmo^.Encours.Deb) and (GetField('I_ETAT')='OUV')) then
  { FQ 15274 - gestion du mode var avec antérieurs
  end else if ((GetField('I_METHODEECO')='VAR') and (GetField('I_DATEDEBECO')<VHImmo^.Encours.Deb) and (GetField('I_ETAT')='OUV')) then
  begin
    Result := 37;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_METHODEECO');
  }
  end
  else if ((GetCheckBoxState('bReintegrationFiscale') = cbChecked) and
    (GetField('I_REINTEGRATION') = 0)) then
  begin
    Result := 11;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_REINTEGRATION');
  end
  // ajout mbo FQ 16756 17.10.05
  else if ((GetCheckBoxState('bReintegrationFiscale') = cbChecked) and
    (GetField('I_REINTEGRATION') > GetField('I_BASEECO'))) then
  begin
    Result := 38;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_REINTEGRATION');
  end
  //  (GetField('I_TAUXECO') > GetField('I_TAUXFISC'))) then
  //begin
  //  Result := 12;
  //  SetActiveTabSheet('PPlan');
  //  SetFocusControl('I_DUREEECO');
  //end
  //else if ((GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
  //  (GetControlVisible('I_REPRISEECO')) and
  //  (GetField('I_REPRISEECO') > GetField('I_REPRISEFISCAL')) and (GetField('I_METHODEECO')<>'VAR')) then
  //begin
  //  Result := 13;
  //  SetActiveTabSheet('PPlan');
  //  SetFocusControl('I_REPRISEECO');
  //end

  {fq 13047 - mbo - le 20.04.2006
  else if (GetField('I_DATEAMORT') > VHImmo^.Encours.Fin) and (not
       ImCQuelExercice(GetField('I_DATEAMORT'), Exo)) then
   begin
   Result := 14;
   SetActiveTabSheet('PGeneral');
   SetFocusControl('I_DATEAMORT');
  }
   (* FQ 10517 - pas de contrôle de la date de mise en service
  end else if (GetField('I_DATEAMORT') < GetField('I_DATEPIECEA')) then
  begin
    Result := 33;
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DATEAMORT');*)
  // fq 13047 end

  else if ((GetControlVisible('PContrat')) and
    (GetControlEnabled('GroupeDepotGarantie')) and (GetField('I_DEPOTGARANTIE') =
    0)) then
  begin
    Result := 15;
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_DEPOTGARANTIE');
  end
  else if ((GetControlVisible('PContrat') and
    GetControlEnabled('GroupeDepotGarantie')) and
    ((GetField('I_IMMOLIEGAR') = GetField('I_IMMO'))
    or (ExisteCodeImmo(GetField('I_IMMOLIEGAR'))))) then
  begin
    Result := 16;
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_IMMOLIEGAR');
  end
  else if ((GetControlVisible('PContrat')) and
    (GetControlEnabled('GroupeDepotGarantie')) and
    (not ControleCompteDepotGarantie(GetControlText('COMPTE_DEPOT')))) then
  begin
    Result := 17;
    SetActiveTabSheet('PContrat');
    SetFocusControl('COMPTE_DEPOT');
  end
  else if (GetField('I_DUREEECO') <> 0) and
    (not MinDureeAmortissement(GetField('I_METHODEECO'), GetField('I_DUREEECO'),
      stMsg)) then
  begin
    Result := -1;
    PGIBox(stMsg);
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_DUREEECO');
  end
  // fq 19620 mbo 25.07.07
  else if (GetField('I_DUREEECO') >180) and
    (GetField('I_METHODEECO')= 'DEG') and (fLeMode in [taCreat..taCreatOne]) then
  begin
    if PGIAsk(HMC[52]) = mrYes then
    begin
      SetActiveTabSheet('PPlan');
      SetFocusControl('I_DUREEECO');
      result:= -1;
    end;
  end
  else if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
    (not MinDureeAmortissement(GetField('I_METHODEFISC'),
    GetField('I_DUREEFISC'), stMsg)) then
  begin
    Result := -1;
    PGIBox(stMsg);
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_DUREEFISC');
  end
  //fq 21553 - mbo - 20/07/07
  else if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
    (GetField('I_DUREEFISC')>1200) then
  begin
     Result := 51;
     SetActiveTabSheet('PPlan');
     SetFocusControl('I_DUREEFISC');
  end
  // fq 19620 mbo 25.07.07
  else if (GetField('I_DUREEFISC') >180) and
    (GetField('I_METHODEFISC')= 'DEG') and (fLeMode in [taCreat..taCreatOne]) then
  begin
    if PGIAsk(HMC[53]) = mrYes then
    begin
      SetActiveTabSheet('PPlan');
      SetFocusControl('I_DUREEFISC');
      result:= -1;
    end;
  end
  else if (GetField('I_METHODEECO') = '') and (GetField('I_NATUREIMMO') = 'PRO')
    then
  begin
    Result := 24;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_METHODEECO');
  end
  // BTY 07/05 FQ 16171
  //else if (GetField('I_REPRISEECO') > GetField('I_MONTANTHT')) then
  else if (GetField('I_REPRISEECO') > GetField('I_BASEECO')) then

  begin
    Result := 29;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_REPRISEECO');
  end
  else if (GetField('I_REINTEGRATION') > GetField('I_MONTANTHT')) then
  begin
    Result := 25;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_REINTEGRATION');
  end
    { FQ 14773 - FQ 14922 }
  else
  if ( wBaseEco >   Arrondi((GetField('I_MONTANTHT') +
     GetField('I_TVARECUPERABLE') - GetfIeld('I_TVARECUPEREE')), V_PGI.OkDecV))
      then
{  FQ 21410
   if ((GetField('I_BASEECO') > Arrondi((GetField('I_MONTANTHT') +
        GetField('I_TVARECUPERABLE') - GetfIeld('I_TVARECUPEREE')), V_PGI.OkDecV) ))  then }
  begin
    Result := 34;
    SetActiveTabSheet('PPlan');
    SetFocusControl('I_BASEECO');
  end
  else
    { Si on peut saisir le montant HT et qu'il est nul ==> message - FQ 12499 - CA - 10/05/2004 }
    if
      ((GetControlEnabled('I_MONTANTHT')) and (GetField('I_MONTANTHT') = 0)) then
  begin
      Result := 32;
      SetActiveTabSheet('PGeneral');
      SetFocusControl('I_MONTANTHT');
  end
  else if ((GetField('I_METHODEECO')='VAR') and (GetField('I_UNITEOEUVRE')=0)) then
  begin
      Result := 35;
      SetActiveTabSheet('PPlan');
      SetFocusControl('I_UNITEOEUVRE');
  end
      //  else if ((fLeMode in [taCreat..taCreatOne]) and (GetField('I_QUANTITE') = 0)) then
  else if (GetField('I_QUANTITE') = 0) then // CA - 29/10/2004 - FQ 14510
     Begin
      if PGIAsk(HMC[26]) <> mrYes then
        Result := -1;
     end
  // Tga 15/06/2006
  else if Test_Reprise_Depreciation('I_REPRISEDEP')=False Then
      Result := -1
  else if Test_Reprise_Depreciation('I_REPRISEFISCAL')=False Then
      Result := -1;

end;

function TOM_IMMO.ControleDate: integer;
begin
  Result := 0;
  if (GetField('I_DATEPIECEA') < (date - 75 * 365)) then
  begin
    if (PGIAsk('Cette date semble incohérente. Voulez-vous la modifier ?',
      ChampToLibelle('I_DATEPIECEA')) = mrYes) then
    begin
      Result := -1;
      SetActiveTabSheet('PGeneral');
      SetFocusControl('I_DATEPIECEA');
    end;
  end
  // ajout mbo 17569
  else  if bPlanFiscal then
  begin
     if ((GetField('I_DATEDEBECO')) < (GetField('I_DATEDEBFIS'))) then
     begin
       Result := 41;
       SetActiveTabSheet('PPlan');
       SetFocusControl('I_DATEDEBECO');
     end;
  end
  else if (GetControlVisible('PContrat')) and (GetCheckBoxState('bDepotGarantie')
    = cbChecked) and
    (GetField('I_DATEDEPOTGAR') < StrToDate(GetControlText('dtDebutContrat')))
      then
  begin
    Result := 18;
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DATEDEPOTGAR');
  end
  else if (GetControlVisible('PContrat')) and (GetCheckBoxState('cbI_DATESUSPCB')  // fq 17362
    = cbChecked) and (GetField('I_DATESUSPCB') <
    StrToDate(GetControlText('dtDebutContrat'))) then
  begin
    Result := 19;
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_DATESUSPCB');
  end;
  if (GetControlVisible('PContrat')) and (GetCheckBoxState('cbI_DATESUSPCB') =
    cbChecked) and (GetField('I_DATESUSPCB') > GetField('I_DATEFINECH')) then
  begin
    Result := 20;
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_DATESUSPCB');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 16/06/2006
Modifié le ... :   /  /
Description .. : FQ 18430 - ne faire le controle du minimum legal
Suite ........ : qu'en validation de la fiche pour pouvoir prendre
Suite ........ : en compte l'antérieur dépréciation
Suite......... : le controle est à faire uniquement si l'immo est modifiable (22.06.06)
Suite......... : fq 15274 - on ne fait pas de controle minimum légal en mode VARIABLE
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.ControleMiniLegal: integer;
var
  anterieurfisc : double;
  anterieureco  : double;
  dMInimumLineaire : double;
  bOUvert : boolean;
  sEtat: string;
  duree : integer;
begin
  Result := 0;
  { mbo - 25.08.06 - correction de conseil de compilation
  bOUvert := true;
  dMinimumLineaire := 0;
  anterieurfisc := 0;
  anterieureco := 0;   }

  // fq 15274 - mbo 07.06.06 - on ne fait pas de controle minimum légal en plan VAR
  if GetField('I_METHODEECO') <> 'VAR' then
  begin
    // mbo 22.06.06 ajout du test pour ne faire le controle que si l'immo est
    // ouverte en modif plan
    sEtat := GetField('I_ETAT');

    bOuvert := (sEtat <> 'CLO') and (sEtat <> 'FER') and (sEtat <> 'SUS')
      and (fLeMode <> taConsult) and (not fOperationEnCours) and (not
      ((GetField('I_IMMOLIEGAR') <> '') and (GetField('I_NATUREIMMO') = 'FI'))
      and (not fDPIUtilisee) and (not fComposant) ); // BTY 07/06 DPI

    if bOuvert = true then
    begin

     anterieurfisc :=  GetField('I_REPRISEFISCAL') + GetField('I_REPRISEDEP');
     anterieurfisc := Arrondi(anterieurfisc, V_PGI.OkDecV);

     anterieureco :=  GetField('I_REPRISEECO') + GetField('I_REPRISEDEP');
     anterieureco := Arrondi(anterieureco, V_PGI.OkDecV);

     // ajout mbo 29.06.07
     if (GetCheckBoxState('bGestionFiscale')<> cbChecked) then
        duree := GetField('I_DUREEECO')
     else
        duree := Max(GetField('I_DUREEECO'), GetField('I_DUREEFISC'));

     dMinimumLineaire := GetMinimumLineaire(TQuery(DS), duree);
     if ((GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
         (anterieurfisc < dMinimumLineaire))
     or ((not (GetCheckBoxState('bAmortissementFiscal') = cbChecked)) and
         (anterieureco < dMinimumLineaire)) then
     begin
       if PGIAsk(HMC[30]) = mrYes then
       begin
         SetActiveTabSheet('PPlan');
         if GetCheckBoxState('bAmortissementFiscal') = cbChecked then
         begin
            SetFocusControl('I_REPRISEFISCAL');
            result:= -1;
         end else
         begin
            SetFocusControl('I_REPRISEECO');
            result:=-1;
         end;
         Ecran.ModalResult := mrNone;
         exit;
      end;
     end;
    end; // ajout pour fq 15274 - 06.07.06 - mbo
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 07/03/2007
Modifié le ... :   /  /
Description .. : fonction vérifiant l'équilibre entre les reprises et le
Suite ........ : dérogatoire et/ou la réintégration saisis
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.ControleDerogatoire: integer;

begin
// chantier fiscal
  Result := 0;

  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
          (GetField('I_METHODEECO') <> 'VAR') then
  begin
     if arrondi((GetField('I_REPRISEFISCAL')- GetField('I_REPRISEECO')), V_PGI.OkDecV) <>
         arrondi((GetField('I_REPRISEDR') - GetField('I_REPRISEFEC')), V_PGI.OkDecV) then
      begin
        Result := 47;
        SetActiveTabSheet('PPlan');
        SetFocusControl('I_REPRISEDR');
      end;
  end
  else if (GetCheckBoxState('GestionFiscale') = cbChecked) and
          (GetField('I_METHODEECO') = 'VAR') then
  begin
      if arrondi((GetField('I_REPRISEFISCAL')- GetField('I_REPRISEECO')-GetField('I_REPRISEDEP')), V_PGI.OkDecV) <>
            arrondi((GetField('I_REPRISEDR') - GetField('I_REPRISEFEC')), V_PGI.OkDecV) then
      begin
      Result := 48;
      SetActiveTabSheet('PPlan');
      SetFocusControl('I_REPRISEDR');
      end;
  end


end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 29/03/2007
Modifié le ... :   /  /    
Description .. : on ne peut pas enregistrer un antérieur dépréciation si 
Suite ........ : gestion fiscale du plan d'amortissement
Mots clefs ... : 
*****************************************************************}
function TOM_IMMO.ControleGestionFiscale: integer;

begin
// chantier fiscal
  Result := 0;

  if (GetCheckBoxState('bGestionFiscale') = cbChecked) and
          (GetField('I_REPRISEDEP') <> 0) then
  begin
     Result := 50;
        SetActiveTabSheet('PPlan');
        SetFocusControl('I_REPRISEDEP');
  end;

end;
//****************************************************

function TOM_IMMO.ControleImmo: integer;
var res : boolean;
begin
  // ajout mbo 07.06.07 on ne fait pas de controle sur un composant
  if (GetField('I_REMPLACEE') = '') and (COPY(GetField('I_STRING1'),1,1) = '') then
  begin

     Result := ControleCode;
     if Result = 0 then
       Result := ControleComptes;
     if Result = 0 then
       Result := ControleChamps;
     if Result = 0 then
       Result := ControleDate;
     if Result = 0 then
       Result := ControleGestionFiscale;  // fq 17512 gestion fisc + antérieurs dépréciation impossible
     if Result = 0 then
       Result := ControleDerogatoire;  // ajout pour chantier fiscal fq 17512
     if (Result = 0) then  // ajout mbo fq 18430
       Result := ControleMiniLegal;

     // MVG 09/06/06
     if Result = 0 then
       begin
       Res := DateDebValide('ECO',false);
       if not res then result :=42
       end;
     if Result = 0 then
       begin
       Res := DateDebValide('FISC',false);
       if not res then result :=42
       end;

     if Result = 0 then
       Result := IsValideComptesAssocies;
     if Result = 0 then
       Result := ControleLoyer;
     //XVI 15/09/2006 FQ 12746 début
     {$IFDEF SERIE1}
     if result=0 then
     begin
        //S1_VerifieVentilation(TRUE) ; //XVI 13/09/2006 FQ 12592 //XVI 15/01/2007 FQ 12746
        result:=S1_VerifCOherenceVentilation ;
     end ;
     {$ELSE}
     {$ENDIF !SERIE1}
     //XVI 15/09/2006 FQ 12746 fin
  end else
     result := 0;
end;

//----------------------------------------------------------------------------------------------

function TOM_IMMO.EstUnChampCompte(St: string): boolean;
begin
  St := UpperCase(St);
  Result := (St = 'I_COMPTEIMMO') or (St = 'I_COMPTELIE') or (St =
    'I_ORGANISMECB')
    or (St = 'I_COMPTEAMORT') or (St = 'I_COMPTEDOTATION') or (St =
      'I_REPEXPLOIT')
    or (St = 'I_DOTATIONEXC') or (St = 'I_REPEXCEP') or (St = 'I_VACEDEE')
    or (St = 'I_AMORTCEDE') or (St = 'I_VAOACEDEE') or (St = 'I_COMPTEDEROG')
    or (St = 'I_REPRISEDEROG') or (St = 'I_PROVISDEROG');
end;

procedure TOM_IMMO.OnCompteElipsisClick(Sender: TObject);
var
  stNature, stWhere: string;
  stEditName: string;
  bLookup: boolean;
begin
  bLookup := False;
{$IFDEF EAGLCLIENT}
  stEditName := THEdit(Sender).Name;
{$ELSE}
  stEditName := THDBEdit(Sender).Name;
{$ENDIF}
  stNature := GetField('I_NATUREIMMO');
  if stEditName = 'I_COMPTEIMMO' then
  begin
    if (stNature = 'CB') or (stNature = 'PRO') or (stNature = 'LOC') then
      stWhere := 'G_GENERAL<="' + VHImmo^.CpteImmoSup + '" AND G_GENERAL>="' +
        VHImmo^.CpteImmoInf + '"'
    else if (stNature = 'FI') then
      stWhere := 'G_GENERAL<="' + VHImmo^.CpteFinSup + '" AND G_GENERAL>="' +
        VHImmo^.CpteFinInf + '"';
    stWhere := stWhere + ' AND G_NATUREGENE="IMO"';
      // FQ 12060 - Activé pour S5 et S1 - CA - 03/03/2004
  end
  else if stEditName = 'I_COMPTELIE' then
  begin
    if (stNature = 'CB') then
      stWhere := 'G_GENERAL<="' + VHImmo^.CpteCBSup + '" AND G_GENERAL>="' +
        VHImmo^.CpteCBInf + '"'
    else if (stNature = 'LOC') then
      stWhere := 'G_GENERAL<="' + VHImmo^.CpteLocSup + '" AND G_GENERAL>="' +
        VHImmo^.CpteLocInf + '"';
{$IFDEF SERIE1}
    stWhere := stWhere + ' AND G_NATUREGENE="CHA"';
{$ELSE}
{$ENDIF}
  end
  else if stEditName = 'I_ORGANISMECB' then
  begin
    bLookup := True;
    if fLeMode <> taConsult then
    begin
      if DS.State in [dsBrowse] then
        DS.Edit;
    LookUpList(TControl(Sender), TraduireMemoire('Auxiliaire'), 'TIERS',
        'T_AUXILIAIRE', 'T_LIBELLE', 'T_NATUREAUXI="FOU"', 'T_AUXILIAIRE', True,
        TAGTIERS);
    end;
  end
  else
  begin
    bLookup := True;
    OnEllipsisCompteImmoGene(Sender);
  end;
  if not bLookup then
  begin
    LookupList(TControl(Sender), TraduireMemoire('Comptes'), 'GENERAUX',
      'G_GENERAL', 'G_LIBELLE', stWhere, 'G_GENERAL', True, TAGGENERAUX);
  end;
end;

function TOM_IMMO.PlanAmortValide(var stMethode: string): boolean;
begin
  stMethode := '';
  Result := True;
  if (GetField('I_NATUREIMMO') = 'PRO') or (GetField('I_NATUREIMMO') = 'CB')
    then
  begin
    if (GetField('I_METHODEECO') = '') then
      stMethode := 'I_METHODEECO'
    else if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
      (GetField('I_METHODEFISC') = '') then
      stMethode := 'I_METHODEFISC';
    if (stMethode <> '') then
      Result := false;
  end
  else
    Result := True;
end;

function TOM_IMMO.IsValideComptesAssocies: integer;
var
  CpteAssocies: TCompteAss;
begin
  Result := 0;
  if (GetField('I_NATUREIMMO') = 'PRO') and (GetField('I_METHODEECO') <> 'NAM')
    then
  begin
    RecupereComptesAssocies(TQuery(DS), '', CpteAssocies);
    if (not VerifEtCreationComptesAssocies(CpteAssocies)) then
    begin
      Result := 21;
    end;
  end;
end;

function TOM_IMMO.ControleLoyer: integer;
begin
  Result := 0;
  if ((GetField('I_FRAISECHE') > GetField('I_MONTANTPREMECHE')) or
    (GetField('I_FRAISECHE') > GetField('I_MONTANTSUIVECHE'))) and
    (GetField('I_TYPELOYERCB') = 'LCO') and (GetControlVisible('PContrat')) then
  begin
    Result := 22;
    SetActiveTabSheet('PContrat');
    SetFocusControl('I_FRAISECHE');
  end
  else if (GetField('I_NATUREIMMO') = 'CB')
    and ((fContrat.SommeDesLoyers + GetField('I_RESIDUEL')) <
      GetField('I_MONTANTHT')) then
  begin
    if PGIAsk(HMC[23]) <> mrYes then
    begin
      Result := -1;
      SetActiveTabSheet('PContrat');
    end;
  end;
end;

function TOM_IMMO.CreationImmoDepotGarantie: boolean;
begin
  Result := False;
  if (not ImmoDepotGarantieExiste) and (GetControlEnabled('GroupeDepotGarantie'))
    then
  begin
    GereDepotGarantie;
    fCodeImmoDepotGarantie := GetField('I_IMMOLIEGAR');
    if Transactions(EnregDepotGarantie, 1) <> oeOK then
      MessageAlerte('L''enregistrement de l''immobilisation a échoué.')
    else
    begin
      Result := true;
      ImmoDepotGarantieExiste := true;
    end;
  end
  else
    result := true;
end;

procedure EnregLogAcquisition(CodeImmo: string; DateOpe: TDateTime; PlanActif:
  integer; infos: TInfoLog);
var
  TLog: TOB;
begin
  // Création nouvel enregistrement ImmoLog
  TLog := TOB.Create('IMMOLOG', nil, -1);
  try
    TLog.PutValue('IL_TYPEOP', 'ACQ');
    TLog.PutValue('IL_IMMO', CodeImmo);
    TLog.PutValue('IL_DATEOP', DateOpe);
    TLog.PutValue('IL_ORDRE', 1);
    TLog.PutValue('IL_ORDRESERIE', 1);
    TLog.PutValue('IL_PLANACTIFAV', 0);
    TLog.PutValue('IL_PLANACTIFAP', PlanActif);
    TLog.PutValue('IL_LIBELLE', RechDom('TIOPEAMOR', 'ACQ', FALSE) + ' ' +
      DateToStr(DateOpe));
    TLog.PutValue('IL_TYPEMODIF', AffecteCommentaireOperation('ACQ'));
    TLog.PutValue('IL_TVARECUPERABLE', Infos.TVARecuperable);
    TLog.PutValue('IL_TVARECUPEREE', Infos.TVARecuperee);
    // TGA 07.04.2006
    TLog.PutValue('IL_TAUX',0);
    TLog.InsertDB(nil);
  finally
    TLog.Free;
  end;
end;

procedure UpdateLogAcquisition(CodeImmo: string; DateOpe: TDateTime; infos:
  TInfoLog);
var
  TLog: TOB;
  QLog: TQuery;
begin
  QLog := OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="' + CodeImmo +
    '" AND IL_TYPEOP="ACQ"', FALSE);
  try
    if not QLog.Eof then
    begin
      TLog := TOB.Create('', nil, -1);
      try
        TLog.LoadDetailDB('IMMOLOG', '', '', QLog, False);
        TLog.Detail[0].PutValue('IL_DATEOP', DateOpe);
        TLog.Detail[0].PutValue('IL_TVARECUPERABLE', Infos.TVARecuperable);
        TLog.Detail[0].PutValue('IL_TVARECUPEREE', Infos.TVARecuperee);
        TLog.UpdateDB;
      finally
        TLog.Free;
      end;
    end;
  finally
    Ferme(QLog);
  end;
end;

procedure TOM_IMMO.EnregDepotGarantie;
var
  TLog: TOB;
  infos: TInfoLog;
begin
  TLog := TOB.Create('IMMO', nil, -1);
  try
    TLog.PutValue('I_IMMO', fCodeImmoDepotGarantie);
    TLog.PutValue('I_ABREGE', Copy(fCodeImmoDepotGarantie, 1, 17));
    TLog.PutValue('I_CHANGECODE', '');
    TLog.PutValue('I_NATUREIMMO', 'FI');
    TLog.PutValue('I_LIBELLE', TraduireMemoire('Dépôt de garantie') + ' ' +
      GetField('I_LIBELLE'));
    TLog.PutValue('I_MONTANTHT', GetField('I_DEPOTGARANTIE'));
    TLog.PutValue('I_QUANTITE', 1);
    TLog.PutValue('I_TVARECUPERABLE', 0);
    TLog.PutValue('I_TVARECUPEREE', 0);
    TLog.PutValue('I_BASETAXEPRO', 0);
    TLog.PutValue('I_METHODEECO', 'NAM');
    TLog.PutValue('I_DEVISE', V_PGI.DevisePivot);
    TLog.PutValue('I_BASEECO', GetField('I_DEPOTGARANTIE'));
    TLog.PutValue('I_MONTANTBASEAMORT', GetField('I_DEPOTGARANTIE'));
    TLog.PutValue('I_VALEURACHAT', GetField('I_DEPOTGARANTIE'));
    TLog.PutValue('I_QUALIFIMMO', GetField('I_QUALIFIMMO'));
    TLog.PutValue('I_COMPTEIMMO', GetControlText('COMPTE_DEPOT'));
    TLog.PutValue('I_COMPTEREF', GetControlText('COMPTE_DEPOT'));
    TLog.PutValue('I_DATEPIECEA', GetField('I_DATEDEPOTGAR'));
    TLog.PutValue('I_DATEAMORT', GetField('I_DATEDEPOTGAR'));
    // Ajout mbo fq 17569
    TLog.PutValue('I_DATEDEBECO', GetField('I_DATEDEPOTGAR'));
    TLog.PutValue('I_DATEDEBFIS', GetField('I_DATEDEPOTGAR'));

    TLog.PutValue('I_CODEPOSTAL', GetField('I_CODEPOSTAL'));
    TLog.PutValue('I_VILLE', GetField('I_VILLE'));
    TLog.PutValue('I_PAYS', GetField('I_PAYS'));
    TLog.PutValue('I_ETABLISSEMENT', GetField('I_ETABLISSEMENT'));
    TLog.PutValue('I_LIEUGEO', GetField('I_LIEUGEO'));
    TLog.PutValue('I_IMMOLIEGAR', GetField('I_IMMO'));
    TLog.PutValue('I_ETAT', 'OUV');
    // BTY 01/06 FQ 17259 Nouveau top dépréciation
    // BTY 04/06 FQ 17516 Nouveau top changement de regroupemnt
    //InitOpeEnCoursImmo(TLog, '-', '-', '-', '-', '-', '-', '-', '-', '-');
    //InitOpeEnCoursImmo(TLog, '-', '-', '-', '-', '-', '-', '-', '-', '-', '-');
    InitOpeEnCoursImmo(TLog, '-', '-', '-', '-', '-', '-', '-', '-', '-', '-','-');


    //ATTENTION : le champ immolie est utilisé pour identifier l'immo d'origine pour un depot de gar.
    TLog.PutValue('I_IMMOLIE', GetField('I_IMMO'));
    TLog.PutValue('I_TIERSA', GetField('I_TIERSA'));
    TLog.PutValue('I_REFINTERNEA', GetField('I_REFINTERNEA'));
{$IFDEF SERIE1}
{$ELSE}
    TLog.PutValue('I_VENTILABLE', GetField('I_VENTILABLE'));
    TLog.PutValue('I_VENTILABLE1', GetField('I_VENTILABLE1'));
    TLog.PutValue('I_VENTILABLE2', GetField('I_VENTILABLE2'));
    TLog.PutValue('I_VENTILABLE3', GetField('I_VENTILABLE3'));
    TLog.PutValue('I_VENTILABLE4', GetField('I_VENTILABLE4'));
    TLog.PutValue('I_VENTILABLE5', GetField('I_VENTILABLE5'));
    ImDupliqueVentil(GetField('I_IMMO'), fCodeImmoDepotGarantie);
{$ENDIF}
    TLog.InsertDB(nil);
    infos.TVARecuperable := 0;
    infos.TVARecuperee := 0;
    EnregLogAcquisition(fCodeImmoDepotGarantie, GetField('I_DATEDEPOTGAR'), 0,
      infos);
  finally
    TLog.Free;
  end;
end;

procedure TOM_IMMO.OnClickAmortissementFiscal(Sender: TObject);
var
  PremMois, PremAnnee, NbMois: Word;
begin
  if bChargementEnCours then
    exit;
  if fLeMode = taConsult then
    exit;
  if DS.State = dsBrowse then
    DS.edit;
  if DS.State in [dsInsert, dsEdit] then
  begin
    //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
    AfficheBaseFisc;

    if (not (GetCheckBoxState('bAmortissementFiscal') = cbChecked)) then
    begin
      //PGR -03/10/2005 Forcé à 1 au lieu de 0 pour éviter des message de controle
      SetField('I_DUREEFISC', 1);
      //SetField('I_DUREEFISC', 0);
      SetField('I_TAUXFISC', 0);
      SetField('I_REPRISEFISCAL', 0);
      SetField('I_METHODEFISC', '');
      //PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
      SetControlText('TYPEBASEFISC', '');
      SetField('I_BASEFISC', 0);
      // mbo fq 17569
      SetField('I_DATEDEBFIS', Getfield('I_DATEDEBECO'));
      SetControlChecked('bGestionFiscale',false);
      SetControlEnabled('bGestionFiscale', false);
      SetControlVisible('I_GBrepriseF', false);

    end
    else
    begin
      { 19.07.07 - fq 20555 - la check est forcément cochée
       SetControlEnabled('AmortFisc', (GetCheckBoxState('bAmortissementFiscal') =
                 cbChecked)); }
       SetControlEnabled('AmortFisc',true);
      //if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
        SetField('I_DUREEFISC', 1);

       SetField('I_DATEDEBFIS', GetField('I_DATEDEBECO'));
    end;
    GereFiscal;
    NOMBREMOIS(Getfield('I_DATEAMORT'), VHImmo^.Encours.Deb, PremMois,
      PremAnnee, NbMois);
    //SetField('I_DUREEREPRISE', IntToStr(NbMois - 1)); FQ 20310
    SetField('I_DUREEREPRISE', IntToStr(MaxIntValue([0, NbMois - 1])));
    GereRepriseEco;
    GereRepriseFisc;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Gestion de la coche 'gestion fiscale'
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnClickGestionFiscale(Sender: TObject);

begin
  if bChargementEnCours then
    exit;
  if fLeMode = taConsult then
    exit;
  if DS.State = dsBrowse then
    DS.edit;
  if DS.State in [dsInsert, dsEdit] then
  begin
      bPlanModifie := true;
      TraiteDotationAnterieures;
      AfficheDerogatoire;
  end;

   { if not (GetCheckBoxState('bGestionFiscale') = cbChecked) then
    begin

      if GetField('I_METHODEECO') = 'VAR' then
      begin
         // on est sur plan variable classique --> linéaire sur durée éco etc...

         SetField('I_METHODEFISC', 'LIN');
         SetField('I_DUREEFISC', GetField('I_DUREEECO'));
         SetField('I_DATEDEBFIS',GetField('I_DATEAMORT'));
         SetControlText('TYPEBASEFISC','ECO');
         SetField('I_BASEFISC',GetField('I_BASEECO'));
         SetField('I_DUREEREPRISEF', GetField('I_DUREEREPRISE'));

         AfficheTauxFisc;
         GereRepriseFisc;
         TraiteDotationAnterieures;
         SetField('I_REPRISEFISCAL', CumFiscSav);

         AfficheDerogatoire;
         SetControlEnabled('bAmortissementFiscal', false);
         SetControlEnabled('AmortFisc', false);
      end else
      begin
         GereFiscal;
         TraiteDotationAnterieures;
         AfficheDerogatoire;
      end;
    end else
    begin
      if GetField('I_METHODEECO') = 'VAR' then

         SetControlEnabled('bAmortissementFiscal', true);
         SetControlChecked('bAmortissementFiscal', true);
         SetControlEnabled('bAmortissementFiscal', false);
         SetControlChecked('bGestionFiscale',true);
         SetControlEnabled('AmortFisc', true);
      GereFiscal;
      TraiteDotationAnterieures;
      AfficheDerogatoire;
    end;
  end;}
end;

//================================================================================
procedure TOM_IMMO.OnClickSaisieEcheancier(Sender: TObject);
var i: integer;
    wEcheance, wRecord :PEcheance;
    wListeEcheances : TList;
begin
  // FQ 20005  Création / màj sans avoir eu de clôture  => saisir des tranches
  // FQ 20005 09/07 Création échéancier => toujours saisir des tranches
  //if (fContrat.ListeTranches.Count = 0) or (GetField('I_ETAT') <> 'CLO')  then
  if (fContrat.ListeTranches.Count = 0) or ((fLeMode in [taCreat..taCreatOne]) or
   (DS.State = dsInsert)) then
  begin
    if FicheSaisieEcheancier(GetFIeld('I_DATEPIECEA'), GetPeriode,
       GetField('I_VERSEMENTCB'), fContrat.ListeTranches) then
    begin
      if DS.State in [dsBrowse] then
         DS.Edit;

      fContrat.ConvertTrancheIntoEcheance; // Ou fContrat.CalculEcheances
      SetField('I_DATEDEBUTECH', fContrat.GetDateDebutEcheance);
      SetField('I_DATEFINECH', fContrat.GetDateFinEcheance);
      AfficheNombreEcheance;
      AfficheCoutContrat;
    end;

  end else
  begin
    // FQ 20005 Modif échéancier sur immo clôturée => saisir des échéances
    // FQ 20005 Suite 09/07 Modif échéancier => toujours saisir des échéances

    // Faire une copie de la Tlist des échéances
    wListeEcheances := TList.Create;
    wListeEcheances.Clear;
    if fContrat.ListeEcheances <> nil then
    begin
      for i := 0 to (fContrat.ListeEcheances.Count - 1) do
        begin
        wEcheance := fContrat.ListeEcheances.Items[i];
        New (wRecord);
        wRecord^.Date := wEcheance^.Date;
        wRecord^.Montant := wEcheance^.Montant;
        wRecord^.Frais := wEcheance^.Frais;
        wListeEcheances.Add (wRecord);
        end;
    end;

{$IFNDEF CMPGIS35}
    if FicheSaisieEcheancierEche(GetFIeld('I_DATEPIECEA'), GetPeriode,
       GetField('I_VERSEMENTCB'), wListeEcheances) then
    begin
      // RAZ anciennes échéances
      fContrat.ClearListeEcheances;

      // Recopier la Tlist des nelles échéances dans fContrat
      if wListeEcheances <> nil then
      begin
        for i := 0 to (wListeEcheances.Count - 1) do
          begin
          wEcheance := wListeEcheances.Items[i];
          New (wRecord);
          wRecord^.Date := wEcheance^.Date;
          wRecord^.Montant := wEcheance^.Montant;
          wRecord^.Frais := wEcheance^.Frais;
          fContrat.ListeEcheances.Add (wRecord);
        end;
      end;

      // Attention : Remettre en tranches car
      // setfield=>onchangefield=>onchangementcontrat=> calculecheances qui recharge à partir des tranches
      fContrat.ConvertEcheanceIntoTranche;

      if DS.State in [dsBrowse] then
         DS.Edit;
      SetField('I_DATEDEBUTECH', fContrat.GetDateDebutEcheance);
      SetField('I_DATEFINECH', fContrat.GetDateFinEcheance);
      AfficheNombreEcheance;
      AfficheCoutContrat;
    end;

    // Libérer mémoire
    if wListeEcheances <> nil then
    begin
      for i := 0 to (wListeEcheances.Count - 1) do
      begin
        wRecord := wListeEcheances.Items[i];
        Dispose(wRecord);
      end;
      wListeEcheances.Free;
    end;
{$ENDIF}
  end; // fcontrat.listetranches.count
end;

function TOM_IMMO.GetPeriode: integer;
var
  Periode: integer;
begin
  if GetField('I_PERIODICITE') = 'ANN' then
    Periode := 12
  else if GetField('I_PERIODICITE') = 'SEM' then
    Periode := 6
  else if GetField('I_PERIODICITE') = 'TRI' then
    Periode := 3
  else if GetField('I_PERIODICITE') = 'MEN' then
    Periode := 1
  else
    Periode := 1;
  Result := Periode;
end;

procedure TOM_IMMO.OnChangeOnglet(Sender: TObject);
begin
  if (GetActiveTabSheet('Pages') = TTabSheet(GetControl('PPlan'))) then
  begin
    GereRepriseEco;
    if bPlanFiscal then GereRepriseFisc;
  end
  else
  begin
     if (GetActiveTabSheet('Pages') = TTabSheet(GetControl('PContrat'))) then
         AfficheCoutContrat;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 30/09/2004
Modifié le ... :   /  /
Description .. : Sortie de l'immobilisation
Suite ........ : 30/09/2004 - CA - FQ 14425 : si l'immobilisation fait partie
Suite ........ : d'un regroupement, on propose la sortie du regroupement
Suite ........ : complet.
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnCessionClick(Sender: TObject);
var stMessage : string;
    bRegroupement : boolean;
    Mr : TModalResult;
begin
  bRegroupement := False;
  if ExisteSQL ('SELECT I_IMMO FROM IMMO WHERE I_IMMO<>"'+GetField('I_IMMO')+'" AND I_GROUPEIMMO<>"" AND I_GROUPEIMMO="'+GetField('I_GROUPEIMMO')+'" AND I_OPECESSION<>"X"') then
  begin
    { Plusieurs immobilisations dans le regroupement }
    stMessage :=
      TraduireMemoire('Cette immobilisation est associée au regroupement ');
    stMessage := stMessage + RechDom('AMREGROUPEMENT', GetField('I_GROUPEIMMO'), False);
    stMessage := stMessage + '.#10#13' +
      TraduireMemoire('La sortie concerne-t-elle l''ensemble du regroupement ?');
    if PGIAsk(stMessage) = mrYes then bRegroupement := True;
  end;
  if bRegroupement then
    Mr := ExecuteCessionRegroupement(GetField('I_GROUPEIMMO'))
  else Mr := ExecuteCession(GetField('I_IMMO'));
  if Mr = mrYes then
     // FQ 18963 ATTENTION ICI il arrive que le pgm détecte qu'on a fait une modif
     // et pose la question Modif OUI / NON ?
     // => Forcer la sortie sans la question et sans enregistrer la fiche
     //TToolbarButton97(GetControl('BFerme')).Click;
   begin
     // FQ 18976 conserver les instructions, message de modif ne vient pas du pgm opeencours.pas
     TFFiche(Ecran).Typeaction :=taConsult;
     TFFiche(Ecran).Monofiche :=True;
     TToolbarButton97(GetControl('BFerme')).Click;
   end;
end;

procedure TOM_IMMO.OnEclatementClick(Sender: TObject);
var Code : string;
begin
  // FQ 18393
  //if ExecuteEclatement(GetField('I_IMMO')) = mrYes then
  Code := GetField('I_IMMO');
  if ExecuteEclatement(Code) = mrYes then
//    Close;
  TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.OnSuppressionDPIClick(Sender: TObject);
Var Code: String;
begin
  if PGIAsk(HMC[45]) = mrYes then
    Begin
       Code := GetField('I_IMMO');
      // Mise à jour de la table Immo
      ExecuteSQL ('UPDATE IMMO SET I_DPI="-" WHERE I_IMMO="'+ code +'"');
      // Suppression dans Immomvtd
      ExecuteSQL ('DELETE FROM IMMOMVTD WHERE IZ_IMMO="'+ code +'"');
      // Fermeture
      TToolbarButton97(GetControl('BFerme')).Click;
    End;
end;


procedure TOM_IMMO.OnMutationClick(Sender: TObject);
begin
  if ExecuteMutation(GetField('I_IMMO')) = mrYes then
//    Close;
  TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.GereNonAmortissable;
begin
  SetField('I_BASEECO', (GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE')
    - GetFIeld('I_TVARECUPEREE')));
  SetField('I_DUREEECO', 0);
  SetField('I_TAUXECO', 0);
  SetField('I_REPRISEECO', 0);
  SetControlEnabled('bAmortissementFiscal', false);
  GereFiscal;
  // ajout mbo fq 17569
  SetControlVisible('I_GBRepriseE', false);
  SetControlVisible('I_GBRepriseF', false);
end;

procedure TOM_IMMO.OnChangementContrat(stField: string);
begin
  if fContrat = nil then
    exit;
  if (GetField('I_NATUREIMMO') = 'PRO') then
    exit;
  if DS.state in [dsEdit, dsInsert] then
  begin
    if (stField = 'I_DATEFINECH') or (stField = 'I_DATEDEBUTECH') then
    begin
      if GetField('I_DATEDEBUTECH') < GetField('I_DATEPIECEA') then
        PGIBox(HMC[27]);
      if GetField('I_DATEFINECH') < Getfield('I_DATEDEBUTECH') then
      begin
        PGIBox(HMC[28]);
        SetField('I_DATEFINECH', GetField('I_DATEDEBUTECH'));
        SetFocusControl('I_DATEFINECH');
        exit;
      end;
    end
    else if (stField = 'I_MONTANTPREMECHE') then
      if (GetField('I_MONTANTPREMECHE') <> 0) and (GetField('I_MONTANTSUIVECHE')
        = 0) then
        SetField('I_MONTANTSUIVECHE', GetFIeld('I_MONTANTPREMECHE'));
    if fContrat <> nil then
    begin
      fContrat.Charge(TQuery(DS));
      fContrat.CalculEcheances;
    end;
    AfficheCoutContrat;
    AfficheNombreEcheance;
    EnableBoutonsInfos;
  end;
end;

procedure TOM_IMMO.OnDepotGarantieClick(Sender: TObject);
var
  CodeImmoGar: string;
begin
  GereDepotGarantie;
  if bChargementEnCours then
    exit;
  begin
    if (not (GetCheckBoxState('bDepotGarantie') = cbChecked)) then
    begin
      SetField('I_DATEDEPOTGAR', iDate1900);
      SetField('I_DEPOTGARANTIE', 0);
      SetField('I_IMMOLIEGAR', '');
      SetControlText('COMPTE_DEPOT', '');
    end
    else
    begin
      if GetControlEnabled('GroupeDepotGarantie') then
      begin
        if DS.State in [dsBrowse] then
          DS.Edit;
        SetField('I_DATEDEPOTGAR', GetfIeld('I_DATEPIECEA'));
        CodeImmoGar := NouveauCodeImmo;
        if (CodeImmoGar <> '') and (CodeImmoGar = GetField('I_IMMO')) then
          CodeImmoGar := Format('%.10d', [StrToInt(CodeImmoGar) + 1]);
        SetField('I_IMMOLIEGAR', CodeImmoGar);
      end;
    end;
    SetControlEnabled('GroupeDepotGarantie', GetCheckBoxState('bDepotGarantie')
      = cbChecked);
  end;
end;

procedure TOM_IMMO.DureeMethodeClick(Sender: TObject);
var Code : string;
begin
  // FQ 18393
  //if ExecuteChangePlan(GetField('I_IMMO'), 1) = mrYes then
  Code := GetField('I_IMMO');
  if ExecuteChangePlan(Code, 1) = mrYes then
//    Close;
  TToolbarButton97(GetControl('BFerme')).Click;
end;

//PGR - 09/2005 - Changement de méthode
procedure TOM_IMMO.ChangementMethodeClick(Sender: TObject);
begin
  if AmLanceFiche_FicheChgtMethode (GetField('I_IMMO')) = mrYes then
    TToolbarButton97(GetControl('BFerme')).Click;
end;

//MBO - 21/04/2006 - FQ 17923 - ajout de l'opération changement date de mise en service
procedure TOM_IMMO.ModifMiseServiceClick(Sender: TObject);
begin
  if AmLanceFiche_AMMODIFSERVICE (GetField('I_IMMO')) = mrYes then
    TToolbarButton97(GetControl('BFerme')).Click;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 13/09/2006
Modifié le ... :   /  /
Description .. :  Ajout de l'apération saisie d'une prime d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.SaisiePrimeClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  if AmLanceFiche_AMPRIME (GetField('I_IMMO')) = mrYes then
    TToolbarButton97(GetControl('BFerme')).Click;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. :  Ajout de l'apération saisie d'une subvention d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.SaisieRemplacementClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  if AmLanceFiche_FicheRemplacement(GetField('I_IMMO')) = mrYes then
    TToolbarButton97(GetControl('BFerme')).Click;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 16/04/2007
Modifié le ... :   /  /
Description .. : Remplacement d'un composant de 2ème catégorie
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.SaisieSbvClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  if AmLanceFiche_AMSUBVENTION(GetField('I_IMMO')) = mrYes then
    TToolbarButton97(GetControl('BFerme')).Click;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2006
Modifié le ... :   /  /
Description .. : Opération de réduction de la prime d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnReductionPrimeClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  if AMLanceFiche_AMReducPrime (GetField('I_IMMO')) = mrYes then
     TToolbarButton97(GetControl('BFerme')).Click;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. : Opération de réduction de la subvention d'investissement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnReductionSBVClick(Sender: TObject);
begin
{$IFNDEF CMPGIS35}
  if AMLanceFiche_AMReducSubvention (GetField('I_IMMO')) = mrYes then
     TToolbarButton97(GetControl('BFerme')).Click;
{$ENDIF}
end;

// BTY 09/05 Ajout de la dépréciation d'un regroupement
procedure TOM_IMMO.OnDepreciationActifClick(Sender: TObject);
var stCode, stGroupeImmo : string;
    mr : integer;
begin
  stCode := GetField('I_IMMO');
  stGroupeImmo := GetField('I_GROUPEIMMO');
  mr := mrNo;

  if RegroupementADeprecier (stCode,stGroupeImmo) then
  begin
     // Dépréciation du regroupement si pas déjà déprécié dans l'exo
     If not RegroupementDejaDeprecie (stGroupeImmo) then
        begin
        // Dépréciation du regroupement
        // BTY 12/05 FQ 17183
        //AMLanceFiche_DeprecGroupe(stGroupeImmo);
        mr := StrToInt (AMLanceFiche_DeprecGroupe(stGroupeImmo));
        end;
  end
  else
  begin
       // Dépréciation individuelle
       // AMLanceFiche_DepreciationActif(stcode);
       // BTY 12/05 FQ 17183
       mr := StrToInt (AMLanceFiche_DepreciationActif(stcode));
  end;

  // BTY 12/05 FQ 17183
  if (mr = mrYes) then TToolbarButton97(GetControl('BFerme')).Click;
end;



procedure TOM_IMMO.ElementExceptionnelClick(Sender: TObject);
var Code : string;
begin
  // FQ 18385
  if GetField('I_DATEDEBECO') > VHImmo^.Encours.Fin then
     PGIBox(HMC[43])
  else
    begin
    // FQ 18393
    //if ExecuteChangePlan(GetField('I_IMMO'), 2) = mrYes then
    Code := GetField ('I_IMMO');
    if ExecuteChangePlan(Code, 2) = mrYes then
//    Close;
      TToolbarButton97(GetControl('BFerme')).Click;
    end;
end;

procedure TOM_IMMO.LeveeOptionClick(Sender: TObject);
begin
  if ExecuteLeveeOption(GetField('I_IMMO')) = mrYes then
//    Close;
  TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.AnnulationOperationClick(Sender: TObject);
var ModeOpe : TActionFiche;
begin
  if DS.State=dsBrowse then ModeOpe := taModif else ModeOpe := taConsult;
  if OperationsEnCours(GetField('I_IMMO'), GetField('I_LIBELLE'), true, ModeOpe)
    then
//    Close;
  // ajout mbo 01.09.05 FQ 16385 pour revenir sur la liste des immos
     //close;
     //TToolbarButton97(GetControl('BFerme')).Click;

     // FQ 18963 ATTENTION ICI il arrive que le pgm détecte qu'on a fait une modif
     // et pose la question Modif OUI / NON ?
     // => Forcer la sortie sans la question et sans enregistrer la fiche
     // ATTENTION ICI instructions close + Bferme ne sont pas encadrées par begin end
     // => Bferme actionné même si l'annulation de l'opération n'a pas été réalisée
   begin
     // FQ 18976 Pb dans la fonction OperationsEnCours
     //TFFiche(Ecran).Typeaction :=taConsult;           1e version pour sortir sans valider
     //TFFiche(Ecran).Monofiche :=True;                 1e version   "
     //TToolbarButton97(GetControl('BDefaire')).Click;  2e version   "
     //TToolbarButton97(GetControl('BFerme')).Click;    2e version   "
     TToolbarButton97(GetControl('BFerme')).Click;
   end;
end;

// BTY 09/05 fq 14363
procedure TOM_IMMO.OnPopRegroupementPopup(Sender: TObject);
var CodeModele : string;
    TModele : TOB;
 begin
  CodeModele := '&#@MODELE_' +  THValComboBox(GetControl('I_GROUPEIMMO')).Value;
  TModele := TOB.Create ('',nil,-1);
  try
    TModele.LoadDetailFromSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeModele+'"');
    SetControlVisible('BINFOREG', (TModele.Detail.Count > 0));
  finally
    TModele.Free;
  end;
end;
//

procedure TOM_IMMO.OnPopZoomActionPopup(Sender: TObject);
var
  stNature: string;
  bDepotGar, bAcquisitionAnnee: boolean;
  bCedee: boolean;
  bdep: boolean;
  stcode: string;
  bDejaPrime, bPrimeEncours, bDejaSbv, bSBVEncours : boolean;
  bRevisionPlan : boolean;
begin
  bAcquisitionAnnee := (GetField('I_DATEPIECEA') >= VHImmo^.Encours.Deb)
    and (GetField('I_DATEPIECEA') <= VHImmo^.Encours.Fin);
  bCedee := (GetField('I_QUANTITE') = 0);
  stNature := GetField('I_NATUREIMMO');

  if DS.State = dsBrowse then
  begin

    // TGA 09/11/2005 Test si immo déjà dépréciée
    // Pas d'éclatement, de cession, de modification de base, d'exceptionnel
    // de changement de méthode si bdep=True
    Stcode := GetField('I_IMMO');
    if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_REVISIONECO<>0') then
      bdep:=True
    Else
      bdep:=False;

    if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_SBVPRI<>0') then
      bDejaPrime:=True
    Else
      bDejaPrime:=False;

    if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND I_SBVMT<>0') then
      bDejaSbv:=True
    Else
      bDejaSbv:=False;

    // 10/06 Test opération saisie prime d'équipement OU réduction de prime dans l'exo en cours
    bPrimeEncours := ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="' + Stcode +
                           '" AND (IL_TYPEOP = "PRI" OR IL_TYPEOP = "RPR") AND (IL_DATEOP>="' +
                           USDateTime(VHImmo^.Encours.Deb) + '" AND IL_DATEOP<="' +
                           USDateTime(VHImmo^.Encours.Fin) + '")');
    // 11/06 Test opération saisie subvention OU réduction de subvention dans l'exo en cours
    bSBVEncours := ExisteSQL ('SELECT IL_IMMO FROM IMMOLOG WHERE IL_IMMO="' + Stcode +
                           '" AND (IL_TYPEOP = "SBV" OR IL_TYPEOP = "RSB") AND (IL_DATEOP>="' +
                           USDateTime(VHImmo^.Encours.Deb) + '" AND IL_DATEOP<="' +
                           USDateTime(VHImmo^.Encours.Fin) + '")');

    // TGA 30/08/2006 Autorisation de la suppression de la DPI si non perspective DPI
    SetControlVisible('SupDPI', (fDateDebExoDPI=iDate1900) and (GetField('I_DPI')='X'));

    // mbo - fq 17923 - autorisation modif date de mise en service
    SetControlVisible('ModifService',  (stNature = 'PRO')
                                    and (not bCedee)
                                    and (not bdep)
                                    and (not (GetField('I_OPELEVEEOPTION') = 'X'))
                                    and (GetField('I_STRING1')='')  // mbo 28.05.07
                                    and (not fDPIUtilisee)   // BTY 07/06
                                    );

    // mbo - 13/09/2006 Autorisation saisie d'une prime d'équipement
    // BTY - 10/06 Pas de prime sur une simulation
    SetControlVisible('Prime', (stNature = 'PRO')
                      and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
                      and (not bDejaPrime) and (GetField('I_QUALIFIMMO') = 'R')
                      and (not (GetField('I_NONDED') = 'X'))); // ajout fq 17512

    // mbo - 20/10/2006 Autorisation saisie d'une subvention d'équipement
    // BTY - 10/06 Pas de subvention sur une simulation
    SetControlVisible('Sbv', (stNature = 'PRO')
                      and (not bCedee) and (not (GetField('I_OPELEVEEOPTION') = 'X'))
                      and (not bDejaSbv) and (GetField('I_QUALIFIMMO') = 'R')
                      and (not (GetField('I_NONDED') = 'X'))); // ajout fq 17512

    // mbo - 16/04/2007 - remplacement de composant
    // mbo - 07/06/2007 - si présence d'un plan fiscal
    SetControlVisible('Remplacement', (stNature = 'PRO')
                      and (not bCedee) and (not (GetField('I_OPEREMPL') = 'X'))
                      and (not bDejaPrime)
                      and (GetField('I_METHODEFISC') <> ''));


    // BTY 10/06 Opération de réduction de prime d'équipement :
    // ouverte si une saisie de prime a été effectuée dans un exo antérieur à l'exo en cours
    SetControlVisible('ReductionPrime', (bDejaPrime and (not bPrimeENcours)));

    // BTY 10/06 Opération de réduction de subvention d'investissement :
    // ouverte si une saisie de subvention a été effectuée dans un exo antérieur à l'exo en cours
    SetControlVisible('ReductionSubvention', (bDejaSBV and (not bSBVENcours)));

    // Autorisation Cession
    SetControlVisible('Cession', ((stNature = 'PRO') or (stNature = 'FI'))
      and (GetField('I_QUANTITE') <> 0) and (not bCedee) and (not bdep));

    // Autorisation Eclatement
    bDepotGar := (stNature = 'FI') and (GetField('I_IMMOLIE') <> '') and (not
      bCedee);
    SetControlVisible('Eclatement',  (stNature <> 'CB')
                                    and (stNature <> 'LOC')
                                    and (not bAcquisitionAnnee)  // BTY FQ 17826
                                    and (not bDepotGar)
                                    and (not bCedee)
                                    and (not bdep)
                                    and (not (GetField('I_OPELEVEEOPTION') = 'X'))
                                    and (not (
                                          (GetField('I_METHODEECO')='VAR')
                                          and
                                          (GetField('I_ETAT')<>'OUV')
                                          ))
                                    );
    // Autorisation Mutation
    SetControlVisible('Mutation', (not bCedee) and (not
      (GetField('I_OPELEVEEOPTION') = 'X')) and (not bAcquisitionAnnee));
    // Autorisation Changement de plan
    // BTY 19260
    bRevisionPlan := True;
    if ((GetField('I_METHODEECO')='NAM') and (bDejaPrime or bDejaSBV)) then
       bRevisionPlan := False;

    SetControlVisible('DureeMethode', (stNature = 'PRO') and (not
      bAcquisitionAnnee) and (not bCedee) and (not (GetField('I_OPECHANGEPLAN') =
      'X')) and (GetField('I_METHODEECO')<>'VAR') and (bRevisionPlan));

    // Autorisation exceptionnel - modif mbo 14.11.2005
    {SetControlVisible('ElementExceptionnel', (stNature = 'PRO') and (not bCedee)
      and (not bdep) and (not (Getfield('I_OPECHANGEPLAN') = 'X')) and
      (not (GetField('I_METHODEECO')='NAM')) and (GetField('I_METHODEECO')<>'VAR'));}

    SetControlVisible('ElementExceptionnel', (stNature = 'PRO') and (not bCedee)
      and (not bdep) and (not (Getfield('I_OPECHANGEPLAN') = 'X')) and
      (not (GetField('I_METHODEECO')='NAM')));

    // Autorisation Levée option
    SetControlVisible('LeveeOption', (stNature = 'CB') and (GetField('I_IMMOLIE')
      = '') and (not bCedee));

    // Autorisation modification des bases - FQ 14277 non autorisée pour location longue durée
    SetControlVisible('ModifBases', (not bAcquisitionAnnee) and (not bCedee) and
      (stNature <> 'LOC')  and (not bdep) and
      (not (GetField('I_OPELEVEEOPTION') = 'X')) and (stNature <> 'FI') and (GetField('I_METHODEECO')<>'VAR')
       and (GetField('I_METHODEECO')<>'NAM'));     // FQ 19015

    // mbo 05/07/07
    // on peut supprimer l'opération de remplacement sur une immo remplacée à condition qu'il n'y
    // ait pas d'opération sur l'immo remplaçante
    SetControlVisible('AnnulationOperation', ((fOperationEnCours) and (SupRemplacement(stCode))));


    SetControlVisible('SeparAnnulOpe', fOperationEnCours);
    { Modèle de regroupement accessible uniquement si un regroupement est affecté à la fiche
      + si l'immo n'est pas issue d'un remplacement de composant (ajout mbo 27.06.07)
      + si immo n'est pas remplacée (mbo 06.07.07)}
    SetControlVisible('BMODELEREG', (GetField('I_GROUPEIMMO')<>'') and
                       (GetField('I_STRING1')='') and (GetField('I_REMPLACEE')=''));

    //PGR - 09/2005 - Changement de méthode
    SetControlVisible('ChangementMethode', False);
    // TGA 08/09/2005
    SetControlVisible('DepActif', false);

    //PGR - 06/12/2005 - Autorisation Changement de méthode
    SetControlVisible('ChangementMethode', (stNature = 'PRO') and (not bAcquisitionAnnee)
      and not (GetField('I_METHODEECO')='VAR') and (not bdep) and (not bCedee)
      and (not (Getfield('I_OPECHANGEPLAN') = 'X'))
      and (not (Getfield('I_OPEMODIFBASES') = 'X'))
      and (not (Getfield('I_NONDED') = 'X')));  // Ajout mbo fq 17512

    // TGA 08/09/2005
    // Autorisation Dépréciation Actif
    SetControlVisible('DepActif', stNature = 'PRO');
  end;
  // CA - 03/09/2002 - Si fiche en édition, pas d'opération possible
  // ==> on grise les menus
  // BTY 09/06 Même chose si on est derrière le tableau suivi DPI
  SetControlEnabled('Cession', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('Cession'));
  SetControlEnabled('Eclatement', (DS.State = dsBrowse) and (fDateDebExoDPI =iDate1900)
                                and GetControlVisible('Eclatement'));
  SetControlEnabled('Mutation', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('Mutation'));
  SetControlEnabled('DureeMethode', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('DureeMethode'));
  SetControlEnabled('ElementExceptionnel', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('ElementExceptionnel'));
  SetControlEnabled('LeveeOption', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('LeveeOption'));
  SetControlEnabled('ModifBases', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and  GetControlVisible('ModifBases'));
  SetControlEnabled('AnnulationOperation', (DS.State = dsBrowse) and GetControlVisible('AnnulationOperation'));
  SetControlVisible('SeparAnnulOpe', GetControlVisible('AnnulationOperation'));
  SetControlEnabled('BMODELEREG', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('BMODELEREG'));

  //PGR - 09/2005 - Changement de méthode
  SetControlEnabled('ChangementMethode', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('ChangementMethode'));
  //TGA - 09/2005 - Dépréciation actif
  SetControlEnabled('DepActif', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('DepActif'));
  //MBO - 04/2006 - fq 17923
  SetControlEnabled('ModifService', (DS.State = dsBrowse) and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('ModifService'));

  // ajout mbo pour prime
  SetControlEnabled('Prime', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('Prime'));
  // BTY 10/06 Réduction de prime
  SetControlEnabled('ReductionPrime', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('ReductionPrime'));
  // ajout mbo pour subvention
  SetControlEnabled('Sbv', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('Sbv'));
  // BTY 11/06 Réduction de subvention
  SetControlEnabled('ReductionSubvention', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('ReductionSubvention'));
  // FQ 19019
  SetControlEnabled('SupDPI', (DS.State = dsBrowse)  and (fDateDebExoDPI = iDate1900)
                                and GetControlVisible('SupDPI'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 22/09/2006
Modifié le ... :   /  /
Description .. : gestion d'une popup pour visu plan amortissement ou plan amort théorique pour CB
Suite ........ : et suramortissement (prime d'équipement)
Suite ........ : et plan amortissement subvention
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnPopZoomPlan(Sender: TObject);
var
  stNature: string;
  stcode: string;
  bDejaPrime : boolean;
  bDejaSbv : boolean;
begin
  if DS.State = dsBrowse then
  begin

    Stcode := GetField('I_IMMO');
    if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND ((I_SBVPRI<>0) or (I_SBVPRIC<>0))') then
      bDejaPrime:=True
    Else
      bDejaPrime:=False;
    SetControlVisible('ZoomPlanPrime', bDejaPrime);

    if ExisteSQL ('SELECT * FROM IMMO WHERE I_IMMO="'+stCode+'" AND ((I_SBVMT<>0) or (I_SBVMTC<>0))') then
      bDejaSbv:=True
    Else
      bDejaSbv:=False;
    SetControlVisible('ZoomPlanSbv', bDejaSbv);



   stNature := GetField('I_NATUREIMMO');

   if stNature = 'CB' then
   begin
    SetControlVisible('ZoomPlanAmort',false);
    SetcontrolVisible('ZoomPlanTheo', true);
    SetControlEnabled('ZoomPlanTheo', ImmoAmortie(GetField('I_METHODEECO'),
                      GetField('I_METHODEFISC'), GetField('I_NATUREIMMO')));

   end else
   begin
    SetControlVisible('ZoomPlanAmort',true);
    SetcontrolVisible('ZoomPlanTheo', false);
    SetControlEnabled('ZoomPlanAmort', ImmoAmortie(GetField('I_METHODEECO'),
                      GetField('I_METHODEFISC'), GetField('I_NATUREIMMO')));

   end;

 end else
 begin
   SetControlVisible('ZoomPlanTheo', false);
   SetControlVisible('ZoomPlanAmort', true);
   SetControlEnabled('ZoomPlanAmort', false);
   SetControlVisible('ZoomPlanPrime', false);
   SetControlVisible('ZoomPlanSbv', false);

 end;

end;




procedure TOM_IMMO.NbEcheancesEnter(Sender: TObject);
begin
  fCurNbEcheances := THSpinEdit(GetControl('NbEcheances')).Value;
end;

procedure TOM_IMMO.NbEcheancesExit(Sender: TObject);
var
  DateDebutEch: TDateTime;
  NbEcheances: THSpinEdit;
begin
  inherited;
  NbEcheances := THSpinEdit(GetControl('NbEcheances'));
  if (nbEcheances.Value <> fCurNbEcheances) then
  begin
    // Recalcul de la date de dernière échéance
    if DS.State in [dsBrowse] then
       DS.Edit;
    DateDebutEch := GetField('I_DATEDEBUTECH');
    if (NbEcheances.Value > 0) then
    begin
     // FQ 20005 Suite 09/07 Ouvrir le NB Echéances en loyers variables
     // Le faire ICI AVANT le SetField(i_datefinech) qui fait onchangefield->onchangementcontrat->calculecheances
     // qui recharge à partir des tranches
     if GetField('I_TYPELOYERCB') <> 'LCO' then
     begin
       fContrat.MAJListeEcheances (PlusMois(DateDebutEch, (-1+Trunc(NbEcheances.Value)) * GetPeriode),
               (nbEcheances.Value - fCurNbEcheances));
       fContrat.ConvertEcheanceIntoTranche;  // pour CalculEcheances qui convertit les tranches
     end;
     SetField('I_DATEFINECH', PlusMois(DateDebutEch, (-1+Trunc(NbEcheances.Value)) * GetPeriode));
    end
    else
      SetField('I_DATEFINECH', DateDebutEch);
    if (fContrat<> nil) then
    begin
      fContrat.Charge(TQuery(DS));
      fContrat.CalculEcheances;
    end;
    AfficheCoutContrat;
  end;
end;

procedure TOM_IMMO.OnClickReintegrationFiscale(Sender: TObject);
begin
  SetControlEnabled('I_REINTEGRATION', GetCheckBoxState('bReintegrationFiscale')
    = cbChecked);
  if not bChargementEnCours then
  begin
    if DS.State in [dsBrowse] then
      DS.Edit;
    SetField('I_REINTEGRATION', 0);
    bPlanModifie := true;
  end;
  if ((GetActiveTabSheet('Pages') = TTabSheet(GetControl('PPlan'))) and
    (GetControlEnabled('I_REINTEGRATION'))) then
    SetFocusControl('I_REINTEGRATION');
end;

//PGR 26/10/2005 Ajout coche de calcul du plan futur avec la VNC
procedure TOM_IMMO.OnClickSurCalculAvecVnc(Sender: TObject);
begin
  if fLeMode = taConsult then
    exit;
  //PGR -21/12/2005- FQ 17184 Validation coche VNC sans devoir tabuler
  if (DS.State in [dsBrowse]) and (not bChargementEncours) then
    DS.Edit;

  if (DS.State in [dsEdit]) and (not bChargementEncours) then
  begin
    if (GetCheckBoxState('bCalculAvecVnc') = cbChecked) then
      SetField('I_JOURNALA', '***')
    else SetField('I_JOURNALA', '');

    bPlanModifie := true;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : click sur plan futur sur la VNF
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnClickSurCalculAvecVNF(Sender: TObject);
begin
  if fLeMode = taConsult then
    exit;
  if (DS.State in [dsBrowse]) and (not bChargementEncours) then
    DS.Edit;

  if (DS.State in [dsEdit]) and (not bChargementEncours) then
  begin
    if (GetCheckBoxState('bCalculAvecVNF') = cbChecked) then
      SetField('I_FUTURVNFISC', '***')
    else SetField('I_FUTURVNFISC', '');

    bPlanModifie := true;
  end;

end;


//*******************************************************************************
//mbo - FQ 17362 - click sur coche date de suspension pour les contrats
procedure TOM_IMMO.OnClickSuspensionCB(Sender: TObject);
begin
  if not bChargementEnCours then
  begin
    if DS.State in [dsBrowse] then
      DS.Edit;

    SetCOntrolEnabled('I_DATESUSPCB', GetCheckBoxState('cbI_DATESUSPCB') =
          cbChecked);
    if not (GetCheckBoxState('cbI_DATESUSPCB') = cbChecked) then
        SetField('I_DATESUSPCB', iDate1900)
    else
        SetField('I_DATESUSPCB', StrToDate(GetControlText('dtFinContrat')));
  end;
end;


procedure TOM_IMMO.BVentilClick(Sender: TObject);
var
  Axes: string;
  TVentil: TOB;
  Q: TQuery;
  i : integer;
  Nature: string;
  Test_ventil : boolean;
  {$IFDEF SERIE1}
  {$ELSE}
  j: integer;
  {$ENDIF !SERIE1}
begin
{$IFDEF SERIE1}
//XVI 25/04/2006 FQ 3978 analytoque monoventilation début
  if not VS1Cpta.Y_AnalytiqueAdvanced then
     exit ;
//XVI 25/04/2006 FQ 3978 analytoque monoventilation fin
{$ELSE}
{$ENDIF !SERIE1}
  if DS.State = dsInsert then
    if (ControleCode <> 0) then
      Exit;
  Axes := '';


(*
  if (GetCheckBoxState('I_Ventilable1') = cbChecked) then
    Axes := Axes + '1';
  if (GetCheckBoxState('I_Ventilable2') = cbChecked) then
    Axes := Axes + '2';
  if (GetCheckBoxState('I_Ventilable3') = cbChecked) then
    Axes := Axes + '3';
  if (GetCheckBoxState('I_Ventilable4') = cbChecked) then
    Axes := Axes + '4';
  if (GetCheckBoxState('I_Ventilable5') = cbChecked) then
    Axes := Axes + '5';
    *)

  if (GetField('I_Ventilable1')='X') then Axes := Axes + '1';
  if (GetField('I_Ventilable2')='X') then Axes := Axes + '2';
  if (GetField('I_Ventilable3')='X') then Axes := Axes + '3';
  if (GetField('I_Ventilable4')='X') then Axes := Axes + '4';
  if (GetField('I_Ventilable5')='X') then Axes := Axes + '5';

  // Si le compte possède une ventilation on propose de la déverser sur l'immo.
  if not
    ExisteSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "IM%" AND V_COMPTE="' +
    GetField('I_IMMO') + '"') then
  begin
    Q := OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE LIKE "GE%" AND V_COMPTE="'
      + GetField('I_COMPTEIMMO') + '" ORDER BY V_NATURE', True);
    if not Q.Eof then
    begin
      if
        PGIAsk('Voulez-vous rapatrier la ventilation du compte sur cette immobilisation ?', Ecran.Caption) = mrYes then
      begin
        TVentil := TOB.Create('', nil, -1);
        try
          TVentil.LoadDetailDB('VENTIL', '', '', Q, False);
          for i := 0 to TVentil.Detail.Count - 1 do
          begin
            Nature := TVentil.Detail[i].GetValue('V_NATURE');
            {$IFDEF SERIE1}
            TVentil.Detail[i].PutValue('V_NATURE', 'IMM') ;
            {$ELSE}
            TVentil.Detail[i].PutValue('V_NATURE', 'IM' +
              Nature[Length(Nature)]);
            {$ENDIF !SERIE1} //XVI 18/04/2006 FD 3978
            TVentil.Detail[i].PutValue('V_COMPTE', GetField('I_IMMO'));
          end;
          TVentil.InsertDB(nil);
        finally
          TVentil.Free;
        end;
      end;
    end;
    Ferme(Q);
  end;
  {$IFDEF SERIE1}
  PreVentilation (nil,GetField('I_IMMO'),fLeMode,'IMM') ;
  {$ELSE}
  ParamVentil('IM', GetField('I_IMMO'), Axes, fLeMode, False);
  {$ENDIF !SERIE1}
  { Mise à jour du détail de la ventilation (champs I_VENTILABLEx) - CA - 22/06/2004}
  {$IFDEF SERIE1}
  {$ELSE}
  TVentil := TOB.Create ('',nil,-1);
  try
    TVentil.LoadDetailFromSQL('SELECT DISTINCT(V_NATURE) V_NATURE FROM  VENTIL WHERE V_COMPTE="'+
            GetField('I_IMMO')+'" AND V_NATURE LIKE "IM%"');
    for i:=0 to TVentil.Detail.Count - 1 do
    begin
      for j:=1 to 5 do
      begin
        // modif mbo - FQ 20811 - 22/06/2007 - erreur en cwas
        //if (TVentil.Detail[i].Getboolean('I_VENTILABLE'+IntToStr(j))<>
        test_ventil := (GetField('I_VENTILABLE'+IntToStr(j)) = 'X');

        if (test_ventil) <>
           (TVentil.Detail[i].GetString('V_NATURE')='IM'+IntToStr(j)) then

        begin
          if (TVentil.Detail[i].GetString('V_NATURE')='IM'+IntToStr(j)) then
            SetField('I_VENTILABLE'+IntToStr(j), 'X')
          else SetField('I_VENTILABLE'+IntToStr(j), '-');
        end;
      end;
    end;
  finally
    TVentil.Free;
  end;
  {$ENDIF !SERIE1} //XVI 18/04/2006 FD 3978

  if DS.State = dsInsert then
    SetControlEnabled('I_IMMO', not
      ExisteSql('Select V_COMPTE FROM VENTIL WHERE V_COMPTE="' + GetField('I_IMMO')
      + '" And V_NATURE like "IM%"')); //XVI 18/04/2006 Je crois que là il y a un Bug, car la valeur "IM" ne s'enregistre jamais!!! + '" And V_NATURE="IM"'));
end;

procedure TOM_IMMO.OnLieuDblClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  VerifCodePostal(TQuery(DS), THEdit(GetControl('I_CODEPOSTAL')),
    THEdit(GetControl('I_VILLE')), False);
{$ELSE}
  VerifCodePostal(TQuery(DS), THDBEdit(GetControl('I_CODEPOSTAL')),
    THDBEdit(GetControl('I_VILLE')), False);
{$ENDIF}
end;

procedure TOM_IMMO.OnVisuDepotGarantieClick(Sender: TObject);
begin
  AMLanceFiche_FicheImmobilisation(GetField('I_IMMOLIEGAR'), fLeMode, '');
end;

procedure TOM_IMMO.OnOpeEnCoursClick(Sender: TObject);
var ModeOpe : TActionFiche;
begin
  if DS.State=dsBrowse then ModeOpe := taModif else ModeOpe := taConsult;

  // mbo - 05/07/07 on supprimer le remplacement de composant sur la mère à condition qu'il n'y
  // ait pas d'opération sur la fille
  if SupRemplacement(GetField('I_IMMO')) = false then
     ModeOpe := taConsult;

  if OperationsEnCours(GetFIeld('I_IMMO'), GetField('I_LIBELLE'), false, ModeOpe)
    then
//    close;
     // FQ 18963 ATTENTION ICI il arrive que le pgm détecte qu'on a fait une modif
     // et pose la question Modif OUI / NON ?
     // => Forcer la sortie sans la question et sans enregistrer la fiche
     //TToolbarButton97(GetControl('BFerme')).Click;
   begin
     // FQ 18976 Pb dans la fonction OperationsEnCours
     //TFFiche(Ecran).Typeaction :=taConsult;           1e version pour sortir sans valider
     //TFFiche(Ecran).Monofiche :=True;                 1e version  "
     //TToolbarButton97(GetControl('BDefaire')).Click;  2e version  "
     //TToolbarButton97(GetControl('BFerme')).Click;    2e version  "
     TToolbarButton97(GetControl('BFerme')).Click;
   end;
end;


// Clic sur le bouton ouvre l'onglet Informations de la fiche
procedure TOM_IMMO.OnInfoDPIClick(Sender: TObject);
begin
  if (TTabSheet(GetActiveTabSheet('Pages')).TabVisible = True) then
     SetActiveTabSheet('PInformations');
end;


procedure TOM_IMMO.OnAccesImmoLie(Sender: TObject);
begin
  AMLanceFiche_FicheImmobilisation(fCodeImmoLie, taConsult, '');
end;


procedure TOM_IMMO.OnChangementMethode(stField: string);
begin
  bPlanModifie := true;
  if stField = 'I_METHODEECO' then
  begin
     GereSpinDureeEco(GetField('I_METHODEECO'));
     // mbo - 22/06/2006 - fq 17569 suite
     if (GetField('I_METHODEECO') = 'VAR') then
        SetField('I_DATEDEBFIS',GetField('I_DATEAMORT'));
  end;

  if stField = 'I_METHODEFISC' then
  begin
     //ajout mbo 19.07.07 - fq 20555 - pour rétablir le focus en sortie du mode fiscal
     GereSpinDureeFisc(GetField('I_METHODEFISC'));

     if GetField('I_METHODEFISC') = 'DEG' then
     begin
        SetField('I_DATEDEBFIS', GetField('I_DATEPIECEA'));
        setControlenabled('I_DATEDEBFIS', false);
     end
     else
     begin
        SetField('I_DATEDEBFIS', GetField('I_DATEAMORT'));
        SetControlenabled('I_DATEDEBFIS', true);
     end;
  end;

  AfficheTauxEco;
  AfficheTauxFisc;
  TraiteDotationAnterieures;
  GereRepriseEco;
  if bPlanFiscal then GereRepriseFisc;
  SetCOntrolVisible('GBRepriseE', RepriseEcoVisible);
  SetControlVisible('GBRepriseF', RepriseFiscVisible);
  SetField('I_REPRISEFISCAL', CumFiscSav);
  SetField('I_REPRISEECO', CumEcoSav);
  {Chantier fiscal fq 17512
  // ajout pour fq 15274
  if (GetField('I_METHODEECO') = 'VAR') and
     (GetField('I_REPRISEFISCAL') < GetField('I_REPRISEECO')) then
     SetControlProperty('AmortDerog', 'Value', 0)
 else
     SetCOntrolProperty('AmortDerog', 'Value', GetField('I_REPRISEFISCAL') -
       GetField('I_REPRISEECO'));
  }
  AfficheDerogatoire;
  EnableBoutonsInfos;
  GereSurAmortissement;
end;

procedure TOM_IMMO.OnChangementDureeAmort(stField: string; bReinit : boolean);
begin
  bPlanModifie := true;
  AfficheTauxEco;
  AfficheTauxFisc;
  GereRepriseEco;
  if bPlanFiscal then GereRepriseFisc;

  TraiteDotationAnterieures;

  SetField('I_REPRISEFISCAL', CumFiscSav);
  SetField('I_REPRISEECO', CumEcoSav);
  {chantier fiscal fq 17512
  // ajout pour fq 15274
  if (GetField('I_METHODEECO') = 'VAR') and
     (GetField('I_REPRISEFISCAL') < GetField('I_REPRISEECO')) then
     SetControlProperty('AmortDerog', 'Value', 0)
  else
      SetControlProperty('AmortDerog', 'Value', GetField('I_REPRISEFISCAL') -
      GetField('I_REPRISEECO'));
  }
  AfficheDerogatoire;
  EnableBoutonsInfos;
end;

procedure TOM_IMMO.OnModifBasesClick(Sender: TObject);
var Code : string;
begin
  // FQ 18393
  //if ExecuteModificationBases(GetField('I_IMMO')) = mrYes then
  Code := GetField('I_IMMO');
  if ExecuteModificationBases(Code) = mrYes then
//    Close;
  TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.OnPlanAmortClick(Sender: TObject);
var
  BaseAmort: double;
  //Exo: TImExoDate;
begin
  { FQ 10517 - CA - 02/06/2003 - Saisie possible sur N +++++ ....}
  {fq 13047 - on autorise date de mise en service sur un exercice inexistant
  if (GetField('I_DATEAMORT') > VHImmo^.Encours.Fin) and (not
    ImCQuelExercice(GetField('I_DATEAMORT'), Exo)) then
  begin
    MessageAlerte('La saisie n''est pas autorisée sur un exercice inexistant.');
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DATEAMORT');
    exit;
  end; }

  {chantier fiscal fq 17512
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
    (GetField('I_TAUXECO') > GetField('I_TAUXFISC')) then
  begin
    PGIBox(HMC[12]);
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DUREEECO');
    exit;
  end; }
  if (fLeMode in [taCreat..taCreatOne]) or (DS.State = dsInsert) then
  begin
    BaseAmort := GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
      GetField('I_TVARECUPEREE');
    SetField('I_BASEAMORDEBEXO', BaseAmort);
    SetField('I_BASEAMORFINEXO', BaseAmort);
    bPlanModifie := true; // MVG 03/07/2006 FQ 18514
  end;
//   bPlanModifie := true; MVG 03/07/2006 FQ 18514
  AffichePlanAmortissement(fPlanAmor);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 22/09/2006
Modifié le ... :   /  /
Description .. : menu popupPlan : visualisation suramortissement prime
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnPlanPrimeClick(Sender: TObject);
var
  BaseAmort: double;
begin
  {chantier fiscal   fq 17512
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
    (GetField('I_TAUXECO') > GetField('I_TAUXFISC')) then
  begin
    PGIBox(HMC[12]);
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DUREEECO');
    exit;
  end; }
  if (fLeMode in [taCreat..taCreatOne]) or (DS.State = dsInsert) then
  begin
    BaseAmort := GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
      GetField('I_TVARECUPEREE');
    SetField('I_BASEAMORDEBEXO', BaseAmort);
    SetField('I_BASEAMORFINEXO', BaseAmort);
    bPlanModifie := true; // MVG 03/07/2006 FQ 18514
  end;
{$IFNDEF CMPGIS35}
  AffichePlanPrime(fPlanAmor);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 20/10/2006
Modifié le ... :   /  /
Description .. : menu popupPlan : visualisation amortissement subvention
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnPlanSbvClick(Sender: TObject);
var
  BaseAmort: double;
begin
  {chantier fiscal  fq 17512
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
    (GetField('I_TAUXECO') > GetField('I_TAUXFISC')) then
  begin
    PGIBox(HMC[12]);
    SetActiveTabSheet('PGeneral');
    SetFocusControl('I_DUREEECO');
    exit;
  end;
  }
  if (fLeMode in [taCreat..taCreatOne]) or (DS.State = dsInsert) then
  begin
    BaseAmort := GetField('I_MONTANTHT') + GetField('I_TVARECUPERABLE') -
      GetField('I_TVARECUPEREE');
    SetField('I_BASEAMORDEBEXO', BaseAmort);
    SetField('I_BASEAMORFINEXO', BaseAmort);
    bPlanModifie := true; // MVG 03/07/2006 FQ 18514
  end;
{$IFNDEF CMPGIS35}
  AffichePlanSbv(fPlanAmor);
{$ENDIF}
end;


procedure TOM_IMMO.OnPlanRemboursementClick(Sender: TObject);
begin
  ListeDesEcheances(fContrat.ListeEcheances, GetField('I_RESIDUEL'));
end;

procedure TOM_IMMO.OnCompteDepotElipsisClick(Sender: TObject);
var
  stWhere: string;
begin
  stWhere := 'G_GENERAL<="' + VHImmo^.CpteDepotSup + '" AND G_GENERAL>="' +
    VHImmo^.CpteDepotInf + '"';
  LookupList(TControl(Sender), TraduireMemoire('Comptes'), 'GENERAUX',
    'G_GENERAL', 'G_LIBELLE', stWhere, 'G_GENERAL', True, TAGGENERAUX);
end;

procedure TOM_IMMO.OnShowZLS1(Sender: TObject);
var
  i: integer;
begin
{$IFDEF EAGLCLIENT}
  for i := 0 to 9 do
    wTableLib[i] := THEdit(GetControl('I_TABLE' + intToStr(i))).Text;
{$ELSE}
  for i := 0 to 9 do
    wTableLib[i] := THDBEdit(GetControl('I_TABLE' + intToStr(i))).Text;
{$ENDIF}
{$IFDEF SERIE1}
  if GetControlText('TABLELIBRE1') <> GetField('I_TABLE0') then
    SetControlText('TABLELIBRE1', GetField('I_TABLE0'));
  if GetControlText('TABLELIBRE2') <> GetField('I_TABLE1') then
    SetControlText('TABLELIBRE2', GetField('I_TABLE1'));
  if GetControlText('TABLELIBRE3') <> GetField('I_TABLE2') then
    SetControlText('TABLELIBRE3', GetField('I_TABLE2'));
{$ELSE}
{$ENDIF}
end;

procedure TOM_IMMO.OnClickVENTILABLE(Sender: TObject);
var CheckVentil : string;
begin
  if not (bChargementEncours) then
  begin

    if (GetCheckBoxState('I_VENTILABLE')=cbChecked) then CheckVentil := 'X'
    else CheckVentil := '-';

    SetControlChecked('I_VENTILABLE1', GetCheckBoxState('I_VENTILABLE')=cbChecked);
    SetControlChecked('I_VENTILABLE2', GetCheckBoxState('I_VENTILABLE')=cbChecked);
    SetControlChecked('I_VENTILABLE3', GetCheckBoxState('I_VENTILABLE')=cbChecked);
    SetControlChecked('I_VENTILABLE4', GetCheckBoxState('I_VENTILABLE')=cbChecked);
    SetControlChecked('I_VENTILABLE5', GetCheckBoxState('I_VENTILABLE')=cbChecked);

    SetField('I_VENTILABLE1',CheckVentil);
    SetField('I_VENTILABLE2',CheckVentil);
    SetField('I_VENTILABLE3',CheckVentil);
    SetField('I_VENTILABLE4',CheckVentil);
    SetField('I_VENTILABLE5',CheckVentil);
  end;
  SetControlEnabled('BVentil', (GetCheckBoxState('I_VENTILABLE') = cbChecked));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : Affichage du plan d'unités d'oeuvre
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnClickUniteOeuvre(Sender: TObject);
var T : TOB;
  non_modifiable : boolean;

begin
  // fq 19888 - mbo - 27.03.2007 ajout des tests sur mode consul et opération en cours
  non_modifiable := false;
  if (GetField('I_ETAT')<>'OUV') or (fLeMode = taConsult) or
     (fOperationEnCours) then
  begin
    PGIInfo ('Le plan d''unités d''oeuvre n''est accessible qu''en consultation.');
    non_modifiable := true;
  end;

  T := TheTob;

  fPlanUO.LePlan.AddChampSupValeur('TOTALUO',fPlanUO.TotalUO);
  // fq 19888 fPlanUO.LePlan.AddChampSupValeur('ETAT',GetField('I_ETAT'));
  if (non_modifiable) then
     fPlanUO.LePlan.AddChampSupValeur('ETAT','FER')
  else
     fPlanUO.LePlan.AddChampSupValeur('ETAT',GetField('I_ETAT'));

  TheTOB := fPlanUO.LePlan;

  AMLanceFiche_UniteOeuvre;


 if TheTob <> nil then
  begin
        { Si modification du plan d'unité d'oeuvre, on passe en modification }
        if DS.State in [dsBrowse] then DS.Edit;
        fPlanUO.MajPlan(TheTob);
        OnChangementDureeAmort('I_UNITEOEUVRE', False);
        TheTob.Free;
  end;
  TheTob := T;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 28/05/2004
Modifié le ... :   /  /
Description .. : Association des informations de la fiche immo courante à un
Suite ........ : regroupement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnModeleRegroupementClick(Sender: TObject);

      procedure AmEnregistreCommeModele ( pCodeImmo, pCodeRegroupement : string );
      var TModele : TOB;
          stCodeModele : string;
      begin
        TModele := TOB.Create ('IMMO',nil,-1);
        if TModele.SelectDB('"'+pCodeImmo+'"', nil, True) then
        begin
          with TOB.Create('IMMO',nil,-1) do
          begin
            stCodeModele := '&#@MODELE_' + pCodeRegroupement;
            { Suppression du modèle le cas échéant }
            ExecuteSQL ('DELETE FROM IMMO WHERE I_IMMO="'+stCodeModele+'"');

            //Tga 28/06/2006 Maj Immomvtd suppression de l'immobilisation
{$IFNDEF CMPGIS35}
            AM_MAJ_IMMOMVTD('S',stCodeModele,'',0);
{$ENDIF}

            { Mise à jour de l'enregistrement modèle }
            Dupliquer(TModele,True,True);
            PutValue('I_IMMO',stCodeModele);
            PutValue('I_QUALIFIMMO','REG');
            PutValue('I_DATECREATION',Date);
            PutValue('I_DATEMODIF',Date);
            PutValue('I_GROUPEIMMO','');
            { Enregistrement du modèle }
            InsertDB(nil);
            Free;
          end;
        end;
        TModele.Free;
      end;
begin
  if (GetField('I_GROUPEIMMO')<>'') then
  begin
    AmEnregistreCommeModele ( GetField('I_IMMO'), GetField('I_GROUPEIMMO') );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 02/06/2004
Modifié le ... : 18/09/2006
Description .. : Evènement de changement de regroupement
Suite ........ : - CA - 26/08/2004 - on ne déverse le regroupement qu'en
Suite ........ : création
Suite ........ : BTY 09/06 FQ 16149 Récup Etablissement d'un modèle :
Suite ........ : - dossier sans gestion d'établissement => Etabt par défaut
Suite ........ : disable
Suite ........ : - sinon
Suite ........ :      utilisateur avec forçage sur un étabt => Etabt utilisateur
Suite ........ : disable
Suite ........ :      utilisateur sans forçage sur un étabt => Etabt utilisateur enable
Suite ........ :      sinon Etabt du modèle enable
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnChangeRegroupement(Sender: TObject);
var CodeModele : string;
    TModele : TOB;
begin
  if not (DS.State=dsInsert) then exit;  // FQ 14374
  { on déverse dans l'immobilisation en cours les informations du modèle }
  CodeModele := '&#@MODELE_' +  THValComboBox(Sender).Value;
  TModele := TOB.Create ('',nil,-1);
  try
    TModele.LoadDetailFromSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeModele+'"');
    if (TModele.Detail.Count > 0) then
    begin
      SetField('I_COMPTEIMMO',TModele.Detail[0].GetValue('I_COMPTEIMMO'));
      SetField('I_LIEUGEO',TModele.Detail[0].GetValue('I_LIEUGEO'));
      SetField('I_CODEPOSTAL',TModele.Detail[0].GetValue('I_CODEPOSTAL'));
      SetField('I_VILLE',TModele.Detail[0].GetValue('I_VILLE'));
      SetField('I_PAYS',TModele.Detail[0].GetValue('I_PAYS'));

      // BTY 09/06 FQ 16149
      //SetField('I_ETABLISSEMENT',TModele.Detail[0].GetValue('I_ETABLISSEMENT'));
      {$IFNDEF SERIE1}
      GereModeleEtablissement (TModele.Detail[0].GetValue('I_ETABLISSEMENT'));
      {$ENDIF}
      SetField('I_METHODEECO',TModele.Detail[0].GetValue('I_METHODEECO'));
      SetField('I_DUREEECO',TModele.Detail[0].GetValue('I_DUREEECO'));
      SetField('I_DATEPIECEA',TModele.Detail[0].GetValue('I_DATEPIECEA'));
      SetField('I_DATEAMORT',TModele.Detail[0].GetValue('I_DATEAMORT'));
      SetField('I_QUOTEPART',TModele.Detail[0].GetValue('I_QUOTEPART'));
      SetField('I_REINTEGRATION',TModele.Detail[0].GetValue('I_REINTEGRATION'));
      SetField('I_TABLE0',TModele.Detail[0].GetValue('I_TABLE0'));
      SetField('I_TABLE1',TModele.Detail[0].GetValue('I_TABLE1'));
      // ajout mbo fq 17569
      SetField('I_DATEDEBECO',TModele.Detail[0].GetValue('I_DATEDEBECO'));
      SetField('I_DATEDEBFIS',TModele.Detail[0].GetValue('I_DATEDEBFIS'));

    end;
  finally
    TModele.Free;
  end;
  { Affichage du bouton d'accès aux informations du regroupement }
  SetControlVisible ('BREGROUPEMENT',GetField('I_GROUPEIMMO')<>'');
end;

//PGR - 09/2005 - Valeur théorique sur Base amortissement fiscal
{****************************************************************
Auteur  ...... : PGR
Créé le ...... : 29/09/2005
Modifié le ... :
Description .. : Evènement de changement de méthode d'amortissement fiscal
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnExitTypeBaseFisc(Sender: TObject);
begin
  if DS.State in [dsBrowse] then
    DS.Edit;
  AfficheBaseFisc;
end;

procedure TOM_IMMO.AfficheBaseFisc;
var
  BaseAmort: double;
begin
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) then
  begin
    if GetControlText('TYPEBASEFISC') = 'THE' then
      BaseAmort := GetField('I_MONTANTHT') +
            GetField('I_TVARECUPERABLE') - GetField('I_TVARECUPEREE')
    else
      BaseAmort := GetField('I_BASEECO');
    SetField('I_BASEFISC', BaseAmort);
  end
  else SetField('I_BASEFISC', 0);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : Affichage des informations du regroupement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnInformationRegroupementClick(Sender: TObject);
begin
  AmLanceFiche_FicheRegroupement ('&#@MODELE_' +  GetField('I_GROUPEIMMO'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/06/2004
Modifié le ... :   /  /
Description .. : Affichage de la liste des immobilisation qui composent le
Suite ........ : regroupement
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.OnDetailRegroupementClick(Sender: TObject);
begin
  AMLanceFiche_ListeDesImmobilisations ( '', False, taConsult, GetField('I_GROUPEIMMO'));
end;

procedure TOM_IMMO.OnAfterDeleteRecord;
begin
  inherited;
  SupprimeFicheImmo(fCodeSuppression, False) ;
  // 05/06 FQ 18119 Positionner les indicateurs de modif
  VHImmo^.ChargeOBImmo := True;
  ImMarquerPublifi (True);

  TToolbarButton97(GetControl('BFerme')).Click;
end;

procedure TOM_IMMO.OnKeyDownEcran(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F11:
      begin
        Key := 0;
        if (GetControlVisible('BZOOMACTION') and GetControlEnabled('BZOOMACTION')) then
          TPopupMenu(GetControl('POPZOOMACTION')).Popup(Mouse.CursorPos.x,
            Mouse.CursorPos.y);
      end;

    // FQ 12508 - ajout mbo 06/09/2005 + modif 26/09/2005
    VK_F6:
      begin
        Key := 0;
        if GetControlEnabled('BPlanRemboursement') then
           OnPlanRemboursementClick(nil)
        else
           // modif mbo pour prime
           //if GetControlEnabled('BPlanAmort') then OnPlanAmortClick(nil)
           if not LancerZoomPlanUnique then // FQ 19511
              if (GetControlVisible('BPlanAmort') and GetControlEnabled('BPlanAmort')) then
                 TPopupMenu(GetControl('POPZOOMPLAN')).Popup(Mouse.CursorPos.x,
                 Mouse.CursorPos.y);
        end;
  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

// FQ 19511 Lancer directement la fonction au lieu du menu PopZoomPlan, si elle est unique
Function TOM_IMMO.LancerZoomPlanUnique : boolean;
var nb : integer;
begin
  nb := 0;
  result := False;
  OnPopZoomPlan (nil);

  if GetControlVisible('ZoomPlanAmort') and GetControlEnabled('ZoomPlanAmort') then
     nb := nb + 1;
  if GetControlVisible('ZoomPlanTheo') and GetControlEnabled('ZoomPlanTheo') then
     nb := nb + 1;
  if GetControlVisible('ZoomPlanPrime') and GetControlEnabled('ZoomPlanPrime') then
     nb := nb + 1;
  if GetControlVisible('ZoomPlanSbv') and GetControlEnabled('ZoomPlanSbv') then
     nb := nb + 1;

  if (nb = 1) then
     begin
     result := True;
     if GetControlVisible('ZoomPlanAmort') and GetControlEnabled('ZoomPlanAmort') then
        OnPlanAmortClick (nil)
     else
        if GetControlVisible('ZoomPlanTheo') and GetControlEnabled('ZoomPlanTheo') then
           OnPlanAmortClick (nil)
        else
           if GetControlVisible('ZoomPlanPrime') and GetControlEnabled('ZoomPlanPrime') then
             OnPlanPrimeClick (nil)
           else
              OnPlanSbvClick(nil);
     end;
end;


procedure TOM_IMMO.OnAfterFormShowEcran;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 29/10/2004
Modifié le ... :   /  /
Description .. : Mise à jour du format d'affichage des montants en fonctions
Suite ........ : du nombre de décimales de la devise du dossier
Mots clefs ... :
*****************************************************************}
  procedure _MajDisplayFormat;
  begin
    SetControlProperty('I_MONTANTHT', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('I_TVARECUPERABLE', 'DisplayFormat',
      StrfMask(V_PGI.OkDecV, '', True));
    SetControlProperty('I_TVARECUPEREE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
// FQ 16258 MBO 26/07/2005    SetControlProperty('MONTANTTTC', 'Decimals', V_PGI.OkDecV);
//    SetControlProperty('I_MONTANTTTC', 'Decimals', V_PGI.OkDecV);
//    SetControlProperty('MONTANTTTC', 'NumericType', ntDecimal); ==> pb affichage
    SetControlProperty('I_BASETAXEPRO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('I_BASEECO', 'DisplayFormat', StrfMask(V_PGI.OkDecV, '',
      True));
    SetControlProperty('I_BASEFISC', 'DisplayFormat', StrfMask(V_PGI.OkDecV, '',
      True));
    SetControlProperty('I_REINTEGRATION', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('I_REPRISEECO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    //tga 30/01/2006
    SetControlProperty('I_REPRISEDEP', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('I_REPRISEFISCAL', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
//    SetControlProperty('I_REPRISEECO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
//      '', True));
    //SetControlProperty('AMORTDEROG', 'Decimals', V_PGI.OkDecV);
    //SetControlProperty('AMORTDEROG', 'NumericType', ntDecimal);
    //mbo chantier fiscal fq 17512
    SetControlProperty('I_REPRISEDR', 'DisplayFormat',
      StrfMask(V_PGI.OkDecV, '', True));
    SetControlProperty('I_REPRISEFEC', 'DisplayFormat',
      StrfMask(V_PGI.OkDecV, '', True));

    SetControlProperty('I_MONTANTPREMECHE', 'DisplayFormat',
      StrfMask(V_PGI.OkDecV, '', True));
    SetControlProperty('I_MONTANTSUIVECHE', 'DisplayFormat',
      StrfMask(V_PGI.OkDecV, '', True));
    SetControlProperty('I_FRAISECHE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('COUTCONTRAT', 'Decimals', V_PGI.OkDecV);
    SetControlProperty('COUTCONTRAT', 'NumericType', ntDecimal);
    SetControlProperty('CoutFinancementHT', 'Decimals', V_PGI.OkDecV);
    SetControlProperty('CoutFinancementHT', 'NumericType', ntDecimal);
    SetControlProperty('I_DEPOTGARANTIE', 'DisplayFormat', StrfMask(V_PGI.OkDecV,
      '', True));
    SetControlProperty('I_RESIDUEL', 'DisplayFormat', StrfMask(V_PGI.OkDecV, '',
      True));

    // FQ 16282 TGA 27/07/2005
    SetControlProperty('I_UNITEOEUVRE', 'DisplayFormat', StrfMask(V_PGI.OkDecV, '',
      True));
  end;

begin
  _MajDisplayFormat;
end;

procedure TOM_IMMO.OnClickSurAmortissement(Sender: TObject);
begin
  if DS.State=dsBrowse then exit;
  NextPrevControl(Ecran);
  if not bChargementEnCours then
    OnChangementMethode('I_METHODEECO');
end;

procedure TOM_IMMO.OnClickELIGIBLETP(Sender: TObject);
begin
  if not bChargementEncours then
  begin
    if DS.State in [dsBrowse] then
          DS.Edit;

    if (GetCheckBoxState('ELIGIBLETP')=cbChecked) then SetField('I_TABLE9','X')
    else SetField('I_TABLE9','-');
  end;
end;

procedure TOM_IMMO.OnMouseDownI_BlocNote(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if fLeMode = taConsult then exit;
  if DS.State = dsBrowse then
    DS.Edit;
end;

function TOM_IMMO.bPlanFiscal: boolean;
begin

    result:= (GetCheckBoxState('bAmortissementFiscal') = cbChecked);
end;

{***********A.G.L.***********************************************
Auteur  ...... : mbo
Créé le ...... : 05/05/2006
Modifié le ... :   /  /
Description .. : tester la validité soit de la date de debut d'amortissement
Suite ........ : éco, soit de la date de début d'amortissement fiscal
Suite ........ : Cette date doit être supérieure ou égale à la
Suite ........ : plus petite des dates entre date achat et date mise en
Suite ........ : service.
Suite......... : de + la date éco ne doit pas être antérieure à la date fiscale
Mots clefs ... :
*****************************************************************}
function TOM_IMMO.DateDebValide(queldate:string;affmess:boolean):boolean;
var
  newdate:tDateTime;
begin
   result:= true;

   if queldate = 'ECO' then
      newdate:= GetField('I_DATEDEBECO')
   else
      newdate:= GetField('I_DATEDEBFIS');

   if GetField('I_DATEPIECEA') <= GetField('I_DATEAMORT') then
   begin
      if newdate <  GetField('I_DATEPIECEA') then
      begin
         //MVG 09/06/06
         if affmess then PGIBox(MSG_DATEACHAT);
         result:=false;
      end;
   end else
   begin
      if newdate <  GetField('I_DATEAMORT') then
      begin
         //MVG 09/06/06
         if affmess then PGIBox(MSG_DATESERVICE);
         result:=false;
      end;
   end;
end;

//XVI 25/04/2006 FD 3978 Analytique Mono ventilation  Début
{$IFDEF SERIE1}
procedure TOM_IMMO.SectionS1OnExit ( Sender : TObject) ;
begin
  if fLeMode<>taConsult then
  begin
     ForceUpdate ;
     SetField('I_TABLE9',GetControltext('SECTIONS1')) ;
  end ;
end ;
{$ELSE}
{$ENDIF !SERIE1}
//XVI 25/04/2006 FD 3978 Analytique Mono ventilation  fin

//XVI 13/09/2006 FQ 12952 début
{$IFDEF SERIE1}
procedure TOM_IMMO.S1_VerifieVentilation(FaireMAJ : Boolean ) ;
var IsVentilable : Boolean ;
begin
  if VS1Cpta.Y_Saisie then
  begin
     isVentilable:=existeSQL('select G_GENERAL from GENERAUX where G_GENERAL="'+getField('I_COMPTEIMMO')+'" and G_VENTILABLE="X"') ;
     if (isVentilable<>Hent1.StrToBool_(getField('I_VENTILABLE'))) then
     begin
        if (not IsVentilable) or
           (PGIAsk('Le compte associé à cette fiche d''immobilisation est ventilable en analytique.#13 Désirez-vous mettre à jour cette fiche?')=mryes) then
        begin
           if Not IsVentilable then
           begin
              PGIinfo('Le compte d''immobilisation associé n''est pas ventilable.'+si(FaireMAJ,'#13 La ventilation sera enlevée de cette fiche.','')) ; //XVI 20/09/2006 FQ 13092
              if FaireMAJ then
              executeSQL('delete from VENTIL where V_NATURE="IMM" and V_COMPTE="'+getField('I_IMMO')+'"') ;
           end ;
           if (FaireMAJ or IsVentilable) then
           begin
              forceupdate ;
              SetField('I_VENTILABLE',Hent1.BoolToStr_(IsVentilable)) ;
           end ;
        end ;
     end ;
  end ;
End ;

//XVI 15/09/2006 FQ 12746 début
Function TOM_IMMO.S1_VerifCOherenceVentilation : integer ;
var I_IMMO, I_COMPTEIMMO : String ;
begin
  Result:=0 ;
  if VS1Cpta.Y_Saisie then
  begin
     if (Hent1.StrToBool_(getField('I_VENTILABLE'))) then
     begin
        I_IMMO:=Getfield('I_IMMO') ;
        I_COMPTEIMMO:=Getfield('I_COMPTEIMMO') ;
        if ((VS1CPta.Y_analytiqueAdvanced) and (not Y_Existe(I_IMMO,FALSE)) and (not Y_Existe(I_COMPTEIMMO,TRUE))) or
           ((not VS1CPta.Y_analytiqueAdvanced) and (GetField('I_TABLE9')=''))                                                                  then
        begin
           //if VS1Cpta.Y_AutoPropose then
           if not VS1Cpta.Y_SAISIEobligatoire then //XVI 16/01/2007 FQ 12746
              Result:=-45*ord(PGIAsk('Vous avez choisi de ventiler votre immobilisation, la ventilation analytique sera effectuée à 100 % sur la section '
                                    +'d''attente'+si(VS1Cpta.Y_analytiqueAdvanced,' de chacun des niveaux','')+'.#13 Confirmez vous ?')=mrNo)
           else
              result:=47
        end ;
     end ;
  end ;
end ;
//XVI 15/09/2006 FQ 12746 fin
{$ELSE}
{$ENDIF !SERIE1}
//XVI 13/09/2006 FQ 12952 fin

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 26/02/2007
Modifié le ... : 26/02/2007
Description .. : Nouveau pour version 8.OO
Suite ........ : Affichage du dérogatoire
Suite ........ : Remplace le code
Suite ........ : 
Suite ........ : if (GetField('I_METHODEECO') = 'VAR') and
Suite ........ :      (GetField('I_REPRISEFISCAL')< GetField('I_REPRISEECO')) then
Suite ........ :        SetControlProperty('AmortDerog','Value', 0)
Suite ........ :   else
Suite ........ :        SetControlText('AmortDerog',StrFMontant(GetfIeld('I_REPRISEFISCAL') -
Suite ........ :                    Getfield('I_REPRISEECO'), 13,V_PGI.OkDecV, '', TRUE));
Suite ........ :
Mots clefs ... :
*****************************************************************}
Procedure TOM_IMMO.AfficheDerogatoire;

var
  CumAntEco: double;

begin
if (GetCheckBoxState('bAmortissementFiscal') = cbChecked ) then
begin
  SetControlEnabled('I_REPRISEDR', true);

  if (GetField('I_METHODEECO') = 'VAR') then
     CumAntEco := GetField('I_REPRISEECO') + GetField('I_REPRISEDEP')
  else
     CumAntEco := GetField('I_REPRISEECO');


  if (GetCheckBoxState('bGestionFiscale') <> cbChecked ) then
  begin
     if (GetField('I_REPRISEFISCAL')< CumAntEco) then
     begin
        SetField('I_REPRISEDR', 0);
        SetField('I_REPRISEFEC', 0);
     end else
     begin
        SetField('I_REPRISEFEC', 0);
        SetField('I_REPRISEDR', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
                      CumAntEco, 13, V_PGI.OkDecV, '', TRUE));
     end;

  end else
  begin
     begin
        if (GetField('I_METHODEECO') = 'VAR') then
        begin
           if (GetField('I_REPRISEFISCAL')< CumAntEco) then
           begin
              SetField('I_REPRISEDR', 0);
              SetField('I_REPRISEFEC', StrFMontant(CumAntEco -
                         Getfield('I_REPRISEFISCAL'), 13, V_PGI.OkDecV, '', TRUE));
              //SetField('I_REPRISEDR',(GetField('I_REPRISEECO') - GetField('I_REPRISEFISC')));
           end else
           begin
              SetField('I_REPRISEDR', StrFMontant(GetfIeld('I_REPRISEFISCAL') -
                         CumAntEco, 13, V_PGI.OkDecV, '', TRUE));
              //SetField('I_REPRISEDR',(GetField('I_REPRISEFISCAL') - GetField('I_REPRISEECO')));
              SetField('I_REPRISEFEC',0);
           end;
        end else
        begin
           // on ne gère plus de dérogatoire : tout est en réintégration/Déduction
           if (GetField('I_REPRISEFISCAL')< CumAntEco) then
           begin
              SetField('I_REPRISEDR', 0);
              SetField('I_REPRISEFEC', StrFMontant(CumAntEco -
                         Getfield('I_REPRISEFISCAL'), 13, V_PGI.OkDecV, '', TRUE));
              //SetField('I_REPRISEDR',(GetField('I_REPRISEECO') - GetField('I_REPRISEFISC')));
           end else
           begin
              SetField('I_REPRISEFEC', StrFMontant(CumAntEco -
                         Getfield('I_REPRISEFISCAL'), 13, V_PGI.OkDecV, '', TRUE));
              //SetField('I_REPRISEDR',(GetField('I_REPRISEFISCAL') - GetField('I_REPRISEECO')));
              SetField('I_REPRISEDR',0);
           end;
           SetControlEnabled('I_REPRISEDR', false);
        end;
     end;
  end;
end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse Boudin
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : Gestion du calcul du plan fiscal futur sur la VNF
Mots clefs ... :
*****************************************************************}
Procedure TOM_IMMO.GereCocheVNF;
begin
  if (GetCheckBoxState('bAmortissementFiscal') = cbChecked) and
      (GetField('I_METHODEFISC') = 'LIN') then
  begin
    if (GetField('I_METHODEECO')<>'VAR') then
    begin
       if GetField('I_REPRISEDEP') <> 0 then
       begin
          SetControlChecked('bCalculavecVNF', true);
          SetControlEnabled('bCalculavecVNF', true);
       end else
          SetControlEnabled('bCalculAvecVNF', true);
    end else
    begin
       // la dépréciation n'a pas d'impact sur le plan fiscal si plan eco = var
       SetControlChecked('bCalculAvecVNF', false);  // comme avant
       SetControlEnabled('bCalculAvecVNF', true);
    end;
  end else
  begin
    SetControlChecked('bCalculAvecVNF', false);
    SetControlEnabled('bCalculAvecVNF', false);
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 21/07/2006
Modifié le ... :   /  /
Description .. : Mise à jour de la DPI
Mots clefs ... :
En entrée  ... : montant de la DPI     => FmontantDPI
                 date de la DPI        => FDateFinExoDPI
                 Montant max de la dpi => FMontantMaxiDPI
                 fAmDPI
*****************************************************************}
Function TOM_IMMO.MiseAJourDPI (codeimmo:String):boolean;

var montant : Double;
    MTdpi : Double;
    MTaffecte : Double;
    MTdispo : Double;
    Q : Tquery;
    sequence : integer;
    I : integer;
    TLog : TOB;
begin
   Inherited ;

   // ajout mbo pour avertissement en compilation - 25.08.06
   Result := false;

{$IFNDEF CMPGIS35}
   IF FDPI=True Then
     Begin
       Result := True;
       Montant := FmontantDPI;

       BEGINTRANS ;
       try

         // Suppression des affectations sur l'immo  dans IMMOMVTD si existant
         executeSQL('DELETE FROM IMMOMVTD WHERE IZ_IMMO="'+ codeimmo +'" ');

         // Mise à jour des montants affectés sur l'immo
         for I := 5 downto 1 do  //0 do attention indice 0 non renseigné
           Begin
             // Affectation sur les exercices antérieurs
             // FQ 18992  sur les exercices antérieurs à l'exo N en cours
             //IF (montant<>0) AND (fAmDPI.Exercice[i].Fin<> idate1900) AND (fAmDPI.Exercice[i].Fin < FDateFinExoDPI) Then
             IF  (montant<>0) AND (fAmDPI.Exercice[i].Fin<> idate1900)
             and (fAmDPI.Exercice[i].Fin < VHImmo^.Encours.Fin) Then
               Begin

                 // Recherche du montant de la DPI pour l'exercice concerné
                 Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(fAmDPI.Exercice[i].Fin)+'"  AND ' +
                             'IZ_NATURED="DPI"', FALSE) ;
                 If Not Q.EOF Then
                    MTdpi := Q.FindField('IZ_MONTANT').AsFloat
                 else
                    MTdpi :=0;

                 // Recherche des affectations de la DPI pour l'exercice concerné
                 MTaffecte := AM_MONTANT_DPI_UTILISEE(fAmDPI.Exercice[i].Fin);

                 // Dpi disponible
                 MTdispo := MTdpi - MTaffecte;

                 // Affectation si montant disponible
                 IF MTdispo<>0 THEN
                   Begin
                     // Recherche dernier numéro de séquence pour la date de fin d'exercice
                     sequence := 0 ;
                     Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(fAmDPI.Exercice[i].Fin)+'"', FALSE) ;
                     try
                       While Not Q.Eof do
                         begin
                           IF Q.FindField('IZ_SEQUENCE').AsINTEGER > sequence THEN
                             sequence := Q.FindField('IZ_SEQUENCE').AsINTEGER;
                           Q.Next ;
                         End;
                     finally
                       Ferme(Q);
                     End;

                     // Création de l'affectation
                     TLog := TOB.Create ('IMMOMVTD',nil,-1);
                     try
                       TLog.PutValue('IZ_IMMO',codeimmo) ;
                       TLog.PutValue('IZ_SEQUENCE', sequence + 1) ;
                       TLog.PutValue('IZ_DATE',fAmDPI.Exercice[i].Fin) ;
                       IF montant > Mtdispo Then
                         Begin
                           TLog.PutValue('IZ_MONTANT',Mtdispo) ;
                           montant := montant - MTdispo
                         End
                       Else
                         Begin
                           TLog.PutValue('IZ_MONTANT',Montant) ;
                           montant := 0 ;
                         end;
                       TLog.PutValue('IZ_NATURED','IMM') ;
                       // TGA iz_libelle modifie 30/08/2006
                       TLog.PutValue('IZ_LIBELLE', DateToStr(GetField ('I_DATEDEBECO')));
                       TLog.InsertDB(nil);
                     finally
                       TLog.Free;
                     End;
                   End;
               End;
           End;

       COMMITTRANS ;
       Except
         Result := False;
         ROLLBACK ;
       End;
     End

   Else IF FDPIANT=True Then
     begin
     Result := False;

     BEGINTRANS ;
     try
       // Suppression des affectations immo dans IMMOMVTD
       ExecuteSQL('DELETE FROM IMMOMVTD WHERE IZ_IMMO="'+ codeimmo +'" ');

       // Affectation immo utilisant des DPI antérieures à N-5
       if fAmDPI.Anterieur <> 0 then
         begin
         // n° de séquence
         sequence := 0 ;
         Q := OpenSQL
         ('SELECT MAX(IZ_SEQUENCE) SEQMAX FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(iDate1900)+'"', True);
         try
           if not Q.Eof then
              sequence := Q.FindField('SEQMAX').AsInteger;
         finally
           ferme(Q);
         end;
         // écriture base
         TLog := TOB.Create ('IMMOMVTD',nil,-1);
         try
           TLog.PutValue('IZ_IMMO', codeimmo) ;
           TLog.PutValue('IZ_SEQUENCE', sequence + 1) ;
           TLog.PutValue('IZ_DATE', iDate1900) ;
           TLog.PutValue('IZ_MONTANT', fAmDPI.Anterieur) ;
           TLog.PutValue('IZ_NATURED', 'IMM') ;
           TLog.PutValue('IZ_LIBELLE', DateToStr(GetField ('I_DATEDEBECO')));
           TLog.InsertDB(nil);
           Result := True;
         finally
           TLog.Free;
         end;
         end;

       // Affectation immo utilisant des DPI de N-2 à N-5
       for i := 5 downto 2 do
         if fAmDPI.Solde[i] <> 0 then
            begin
            // n° de séquence
            sequence := 0 ;
            Q := OpenSQL
            ('SELECT MAX(IZ_SEQUENCE) SEQMAX FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(fAmDPI.Exercice[i].Fin)+'"', True);
            try
              if not Q.Eof then
                 sequence := Q.FindField('SEQMAX').AsInteger;
            finally
              ferme(Q);
            end;
            // écriture base
            TLog := TOB.Create ('IMMOMVTD',nil,-1);
            try
              TLog.PutValue('IZ_IMMO', codeimmo) ;
              TLog.PutValue('IZ_SEQUENCE', sequence + 1) ;
              TLog.PutValue('IZ_DATE', fAmDPI.Exercice[i].Fin) ;
              TLog.PutValue('IZ_MONTANT', fAmDPI.Solde[i]) ;
              TLog.PutValue('IZ_NATURED', 'IMM') ;
              TLog.PutValue('IZ_LIBELLE', DateToStr(GetField ('I_DATEDEBECO')));
              TLog.InsertDB(nil);
              Result := True;
            finally
              TLog.Free;
            end;
            end;

     COMMITTRANS ;
     except
       Result := False;
       ROLLBACK ;
     end;

     end;
{$ENDIF}
End;


{$IFNDEF SERIE1}
{BTY 09/06 : Récupération du combo gérant les établissements  FQ 16149
{---------------------------------------------------------------------------------------}
function TOM_IMMO.GetComboEtab : THValComboBox;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  if FComboEtab = nil then begin
    for n := 0 to Ecran.ComponentCount - 1 do
      if (Pos('ETAB', Ecran.Components[n].Name) > 0) and
         (Ecran.Components[n] is THValComboBox) then begin
        FComboEtab := THValComboBox(Ecran.Components[n]);
        Break;
      end;
  end;
  Result := FComboEtab;
end;

// BTY 09/06 Restrictions utilisateur sur les établissements cf FQ 16149
{---------------------------------------------------------------------------------------}
procedure TOM_IMMO.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ComboEtab) then
    begin
    {Si l'on ne gère pas les établissements ...}
    if not VHImmo^.EtablisCpta  then
       begin
       {... on affiche l'établissement par défaut}
       SetControlText(ComboEtab.Name, VHImmo^.EtablisDefaut);
       {... on désactive la zone}
       SetControlEnabled(ComboEtab.Name, False);
       end

    {On gère l'établisement, donc ...}
    else
       begin
      {... On commence par regarder les restrictions utilisateur}
      {$IFNDEF EAGLSERVER}
      PositionneEtabUser;
      {$ENDIF}
      {... s'il n'y a pas de restrictions, on reprend le paramSoc}
      if GetControlText(ComboEtab.Name) = '' then
        begin
        {... on affiche l'établissement par défaut}
        SetControlText(ComboEtab.Name, VHImmo^.EtablisDefaut);
        {... on active la zone}
        SetControlEnabled(ComboEtab.Name, True);
        end;
      end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 18/09/2006
Modifié le ... :   /  /
Description .. : BTY 09/06 FQ 16149 Récup Etablissement d'un modèle :
Suite ........ : - dossier sans gestion d'établissement => Etabt par défaut disable
Suite ........ : - sinon
Suite ........ :       utilisateur avec forçage sur un étabt => Etabt utilisateur disable
Suite ........ :       utilisateur sans forçage sur un étabt => Etabt utilisateur enable
Suite ........ :       sinon Etabt du modèle enable
Mots clefs ... :
*****************************************************************}
procedure TOM_IMMO.GereModeleEtablissement (ModeleEtabt : string);
{---------------------------------------------------------------------------------------}
{$IFNDEF EAGLSERVER}
var Etab : String ;
    Forcer : Boolean ;
{$ENDIF}
begin
  if Assigned(ComboEtab) then
    begin
    {Si l'on ne gère pas les établissements ...}
    if not VHImmo^.EtablisCpta  then
       begin
       {... on affiche l'établissement par défaut}
       SetControlText(ComboEtab.Name, VHImmo^.EtablisDefaut);
       {... on désactive la zone}
       SetControlEnabled(ComboEtab.Name, False);
       end

    {On gère l'établisement, donc ...}
    else
       begin
      {... On commence par regarder les restrictions utilisateur}
      {$IFNDEF EAGLSERVER}
       if (not ComboEtab.Visible) then
          exit;
       Etab := VH^.ProfilUserC[prEtablissement].Etablissement;
       Forcer := VH^.ProfilUserC[prEtablissement].ForceEtab;
       if ComboEtab.Values.IndexOf(Etab) < 0 then
          exit ;
       if Etab <> '' then
          // Etablissement utilisateur
          SetControlProperty(ComboEtab.Name, 'Value', Etab)
       else
          // Etablissement du modèle
          SetControlProperty(ComboEtab.Name, 'Value', ModeleEtabt);
       ComboEtab.enabled := (not Forcer);
      {$ENDIF}
       end;
    end;
end;

{$IFNDEF EAGLSERVER}
procedure TOM_IMMO.PositionneEtabUser;
var Etab : String ;
    Forcer : Boolean ;
begin

  if Assigned(ComboEtab) then
    begin
    if Not VHImmo^.EtablisCpta then
       exit;
    if (not ComboEtab.Visible) then
       exit;

    Etab := VH^.ProfilUserC[prEtablissement].Etablissement;
    if Etab='' then
       exit;
    Forcer := VH^.ProfilUserC[prEtablissement].ForceEtab;

    if ComboEtab.Values.IndexOf(Etab) < 0 then
       exit ;
    SetControlProperty(ComboEtab.Name, 'Value', Etab);
    ComboEtab.enabled := (not Forcer);
    end;
end;
{$ENDIF}   // EAGLSERVER
{$ENDIF}   // SERIE1


initialization
  registerclasses([TOM_IMMO]);
end.



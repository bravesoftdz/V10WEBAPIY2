{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 03/04/2001
Modifié le ... : 03/04/2001
Description .. : Définition des constantes pour le Front Office
Mots clefs ... : FO
*****************************************************************}
unit FODefi;

interface

///////////////////////////////////////////////////////////////////////////////////////
//  Etats d'une journée de caisse
///////////////////////////////////////////////////////////////////////////////////////
const
  ETATJOURCAISSE: array[1..4] of string = (
    {1}'NON', // Non Ouverte
    {2}'OUV', // Ouverte
    {3}'CPR', // Cloturée provisoirement
    {4}'CDE' // Cloturée définitivement
    );

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Types d'opérations de caisse
  ///////////////////////////////////////////////////////////////////////////////////////
const
  TYPEOPCBONACHAT: string = 'ABA'; // Acquisition d'un bon d'achat
  TYPEOPCECART: string = 'ECA'; // Ecart de caisse
  TYPEOPCCHQDIF: string = 'ECD'; // Encaissement de chèque différé
  TYPEOPCENCCREDIT: string = 'ECR'; // Encaissement de crédit
  TYPEOPCENTREE: string = 'END'; // Entrée diverse
  TYPEOPCFDCAISSE: string = 'FCA'; // Fond de caisse
  TYPEOPCPRELEV: string = 'PRE'; // Prélévement
  TYPEOPCREMBARRHES: string = 'RAR'; // Remboursement d'arrhes
  TYPEOPCREMBAVOIR: string = 'RAV'; // Remboursement d'avoir
  TYPEOPCREMBBONACH: string = 'RBA'; // Remboursement de bon d'achat
  TYPEOPCSORTIE: string = 'SOD'; // Sortie diverse
  TYPEOPCARRHES: string = 'VAR'; // Versement d'arrhes

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Types de modes de paiement
  ///////////////////////////////////////////////////////////////////////////////////////
const
  TYPEPAIEESPECE: string = '001'; // Espéce
  TYPEPAIEAVOIR: string = '002'; // Avoir
  TYPEPAIECHEQUE: string = '003'; // Chèque
  TYPEPAIECHQDIFF: string = '004'; // Chèque différé
  TYPEPAIECB: string = '005'; // Carte bancaire
  TYPEPAIEARRHES: string = '006'; // Arrhes déjà versés
  TYPEPAIERESTEDU: string = '007'; // Reste dû
  TYPEPAIEBONACHAT: string = '008'; // Bon d'achat
  TYPEPAIEAUTRE: string = '009'; // Autre type

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Devise EURO
  ///////////////////////////////////////////////////////////////////////////////////////
const
  SIGLEEURO: string = '€'; // Sigle monétaire de l'Euro
  CODEEURO: string = 'EUR'; // Code de la devise Euro

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Contrôle de caisse
  ///////////////////////////////////////////////////////////////////////////////////////
const
  FINPIECE: string = 'FINPIE'; // Marque de fin des libellés des pièces
  FINBILLET: string = 'FINBIL'; // Marque de fin des libellés des billets

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Nature du modéle d'impression
  ///////////////////////////////////////////////////////////////////////////////////////
  //    Table : MODELES
  //                      MO_TYPE    MO_NATURE
  //            Ticket      "T"        "VTC"      Tickets de caisse
  //            Etat        "E"        "GPJ"      Tickets de caisse format état
  //            Document    "L"        "GPI"      Tickets de caisse format document
  //            Etiquette   "E"        "GEA"
  //
const
  NATUREMODTICKET: string = 'GFT';   // Tickets de caisse
  NATUREMODTICETAT: string = 'GPJ';  // Tickets de caisse format état
  NATUREMODTICDOC: string = 'GPI';   // Tickets de caisse format document
  NATUREMODCHEQUE: string = 'GFQ';   // Chèques
  NATUREMODTICKETZ: string = 'GFC';  // Tickets X et Z
  NATUREMODSTATSGP: string = 'GSP';  // JTR - Nvelles stats - Statistiques pièces
  NATUREMODSTATSGPE: string = 'GSE'; // JTR - Nvelles stats - Statistiques mode de paiment
  NATUREMODSTATSGPV: string = 'GSV'; // JTR - Nvelles stats - Statistiques vendeurs
  NATUREMODSTATSGPZ: string = 'GSZ'; // JTR - Spécifiques
  // Type des états imprimés sur le Front Office en mode ticket
type
  TTypeEtatFO = (efoTicket, efoTicketX, efoTicketZ, efoRecapVend, efoListeRegle, efoTicketTest,
    efoCheque, efoRemiseBq, efoBonAchat, efoBonAvoir, efoBonArrhes, efoTransfert,
    efoCommande, efoReception, efoStatFam, efoStatRem, efoFdCaisse, efoListeArtVendu,
    efoStatsGP, efoStatsGPE, efoStatsGPV, efoStatsGPZ);

  {==============================================================================================}
  {=================================== CLÉS DE REGISTRE =========================================}
  {==============================================================================================}
  ///////////////////////////////////////////////////////////////////////////////////////
  //  Fermeture de journée
  ///////////////////////////////////////////////////////////////////////////////////////
const
  REGJOURNEE: string = 'FOJournee';
  REGFERMETICKETZ: string = 'FermeTicketZ'; // Enchaîner sur l'impression du ticket Z
  REGFERMERECAPVD: string = 'FermeRecapVd'; // Enchaîner sur l'impression du récapitulatif vendeurs
  REGFERMEFDCAISSE: string = 'FermeFdCaisse'; // Enchaîner sur l'impression du fond de caisse
  REGFERMELISTEART: string = 'FermeListeArt'; // Enchaîner sur l'impression de la liste des articles vendus
  REGFERMELISTEREG: string = 'FermeListeReg'; // Enchaîner sur l'impression de la liste des règlements
  REGFERMEFOBO: string = 'FermeFOBO'; // Enchaîner sur le démarrage des liaisons FO-BO
  REGTMODEBIMP: string = 'TMODeButImprime'; // Temps d'attente pour lancer la 1ère impression
  REGTMOFINIMP: string = 'TMOFinImprime'; // Temps d'attente pour les impressions suivantes
  REGALERTEALF: string = 'AlerteAnnonces'; // Message d'alerte si annonces non validées
  
  ///////////////////////////////////////////////////////////////////////////////////////
  //  Retransmission d'une journée
  ///////////////////////////////////////////////////////////////////////////////////////
const
  REGTRANSJOUR: string = 'ReTransJour';
  REGENVOIIMMEDIAT: string = 'EnvoiImmediat'; // Envoi immédiat
  REGCODECENTRAL: string = 'CodeSiteCentral'; // Code du site central
  REGLISTEREQUETES: string = 'ListeRequetes'; // Liste des requêtes à associer
  REGCODEEVENEMENT: string = 'CodeEvenement'; // Code de l'événement à générer
  REGTMODEMARRAGE: string = 'TMODemarrage'; // Temps d'attente pour lancer les échanges
  REGTMOATTENTEFIN: string = 'TMOAttenteFin'; // Temps d'attente pour vérifier la fin des échanges

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Situation Flash
  ///////////////////////////////////////////////////////////////////////////////////////
const
  REGSITUFLASH: string = 'SituationFlash';
  REGCATEGORIEPARDEFAUT: string = 'DefaultCategorie';
  REGSTATSPARDEFAUT: string = 'DefaultStats';
  REGVENDEURSPARDEFAUT: string = 'DefaultVendeurs';
  REGTYPECATEGPARDEFAUT: string = 'DefaultTypeCateg';
  REGREGLTSPARDEFAUT: string = 'DefaultReglements';
  REGDETAILOPCAISSE: string = 'DetailOpCaisse';
  REGDETAILPRESTATION: string = 'DetailPrestation';
  REGCAISSEBTQ: string = 'CaisseBoutique';
  STATSINI: Integer = 0;
  STATSCATEG: Integer = 1;
  VENDEURSINI: Integer = 0;
  VENDEURSACTIFS: Integer = 1;
  TYPECATEGINI: Integer = 0;
  TYPECATEGTOUT: Integer = 1;
  REGLTSINI: Integer = 0;
  REGLTSTOUS: Integer = 1;
  REGLTSAUCUN: Integer = 2;

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Mode de connexion à la caisse
  ///////////////////////////////////////////////////////////////////////////////////////
const
  REGCONNEXCAIS: string = 'FOConnexionCaisse';
  REGCHAQUEFOIS: string = 'ChaqueFois'; // A chaque choix d'une ligne de menu
  REGTOUTESLIGNES: string = 'ToutesLignes'; // Pour toutes les lignes du menu
  REGVERIFVERROU: string = 'VerifVerrou'; // Vérifie si la caisse est déjà en cours d'utilisation.

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Choix de la caisse
  ///////////////////////////////////////////////////////////////////////////////////////
const
  CAISSEPARDEFAUT: string = 'DefaultCash'; // Nom du paramètre de la registry pour stocker la caisse choisie
  CAISSEMONOPOSTE: string = 'CaisseMonoPoste'; // Caisse isolée

  ///////////////////////////////////////////////////////////////////////////////////////
  //  Saisie de ticket
  ///////////////////////////////////////////////////////////////////////////////////////
const
  REGSAISIETIC: string = 'FOSaisieTicket';
  REGTMOSIMUCLAV: string = 'TMOSimuleClavier'; // Temps attente avant de simuler la frappe d'une touche
  REGAUTODIALTPE: string = 'AutoDialTpe'; // Lancement automatique du dialogue avec le TPE
  REGQTEMAXFFO: string = 'QteMaxFFO'; // Quantité maximum autorisée pour une ligne de ticket
  REGPRIXMAXFFO: string = 'PrixMaxFFO'; // Prix maximum autorisé pour une ligne de ticket

implementation

end.

{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 24/07/2001
Modifié le ... :   /  /
Description .. : Initialisation des DUCS : sélection de la période,
Suite ........ : du type de DUCS, des établissements
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}
{
 PT1 : 10/10/2001 : V562  MF
                             1- Calcul des effectifs
                             2- Récupération du n° MSA
 PT2 : 15/10/2001 : V562  MF : Modification : la table INSTITUTIONPAYE remplace
                               la table INSTITUTIONPAIE

 PT4 : 07/11/2001 : V562  MF :
                             1- correction décompte d'effectif
                             2- bouton sélection de l'ensemble de la liste
                                désactivé après traitement.
 PT5 : 09/11/2001 : V562  MF : Modification : suite a modification des valeurs
                               de la tablette PGMONNAIE, correction alimentation
                               de la monnaie
 PT6 : 26/11/2001 : V569  MF : Modification du calcul des effectifs (calcul
                                 sur le champ PSA_CATDADS)
 PT7 : 05/12/2001 : V569  MF : Correction erreur SQL mise en évidence avec une
                               base ORACLE.(guillemets manquants)
 PT8 : 03/01/2002 : V571  MF :
                             1-Correction contrôle de date, résout PB qd date
                               incorrecte. Le sablier restait affiché.
                             2-Correction du calcul des effectifs. Il faut tester
                               aussi la date d'entrée.
                             3-Concerne les Ducs Néant. Alimentation des dates
                               d'exigibilité, limite de dépôt et de règlement
                               par rapport à la date de fin de période car la
                               date de versement des salaires n'est pas renseignée.
 PT9 : 08/01/2002 : V571  MF :
                             1- Ajout du champ Déclarant (émetteur) sur mulicritère
                                Le premier de la table EMETTEURSOCIAL est affiché
                                par défaut. Possibilité de sélection d'un autre
                                déclarant.
                             2- Edition des ducs Néant avec le bon fond de page
                             3- Correction fiche de bug n° 417 Concerne le calcul
                                du nbre de mois de la période sélectionnée.
                                Affichage d'un message d'alerte.
                             4- Correction fiche de bug n° 417 . Concerne la
                                prise en compte dans les ducs des salaires à
                                cheval sur deux périodes. Modification de la
                                requête de chargement des cotisation : la date de
                                fin doit être comprise entre les dates de début
                                et de fin de période sélectionnées.
                             5- Correction du traitement des forfaits (type de
                                cotisation = 'Q') = Mt forfait * Effectif
                             6- Traitement du champ LIBELLESUITE de DUCSDETAIL
                                Doit être éditer sur la ligne suivante
 PT10 : 13/02/2002 : V571  MF
                             1- Mise en place de la gestion des ruptures
                             2- Correction du traitement des taux AT (pas
                                uniquement le cas général)
 PT11 : 19/03/2002 : V571  MF
                             1- Correction des requêtes  de sélection des
                                institutions à rectifie le cas des ducs IRC
                                regroupées (bases et effectifs erronés).
                             2- Modification de la requête ActiveWhere pour ne
                                proposer que les organismes destinataires quand
                                code regroupement renseigné.
                             3- modification des requêtes de recherche sur
                                DUCSPARAM pour récupérer en priorité la
                                codification de type DOS sinon STD sinon CEGID
 PT12 : 29/03/2002 : V571 MF    Modification pour traiter le cas d'une ducs
                                "pernonnalisée"

 PT13 : 14/05/2002 : V575 MF    Modification: Alimentation des champs
                                PDU_DUCSDOSSIER et PDU_EMETTSOC

 PT14 : 15/05/2002 : V582 MF
                                1- Rectification des requêtes pour compatibilité
                                ORACLE.
                                2- Suite tests rupture : des Ducs "Néant" étaient
                                créées à tort
 PT15 : 23/05/2002 : V582 MF
                                1- Mise au point traitement IRC multi centre
                                payeur, multi institutions pour IRC EDI .
                                (Ducs Dossier,rupture siret, regroupement de caisses)
 PT15 BIS : 01/07/2002 : V582 MF
                                1- On épure les affectations non
                                uitles au dossier. (corrige le cas plusieurs
                                codifications regroupant les mêmes rubriques)
                                2- Correction anomalie : la codification était
                                forcée à tort à '1A' qd il s'agissait de Taux AT
                                3- Correction anomalie : Le n° d'ordre de Taux AT
                                alimentant la position 3 de la codification n'était
                                pas remis à zéro lors du changement de codif.
                                4- Correction du traitement concernant la création
                                d'une ligne de S/T quand on gère les ruptures
                                (siret, ape, groupe interne, n° interne)
                                5- Correction de l'alimentation du libellé de S/T
                                par le code et libellé de l'institution (on prenait
                                à tort celui du premier établissement)
                                6- Initialisation TOB RupMere : résoud acces Vio
                                qd on lançait une DUCS sans rupture après une DUCS
                                avec rupture.

 PT16 : 18/06/2002 : V582 MF
                                1- Correction alimentation Siret, Ape et groupe
                                sur DUCSDETAIL quand aucune rupture n'est gérée.
                                On reprend les mêmes valeurs que sur DUCSENTETE.
                                (corrige : les champs n'étaient pas alimentés
                                sur DUCSDETAIL, la suppression de DUCS ne pouvait
                                pas être contrôlée correctement)
 PT17 : 21/06/2002 : V585 MF    Traitement de la VLU pour les entreprises
                                en paiement groupé.
                                (Annexe à la DUCS)
 PT18 : 11/07/2002 : V582 MF
                                1- Fiche qualité n° 10180
                                Concaténation de la "raison sociale" de l'émetteur
                                De 40 acr ds EMETTEURSOCIAL à 35 car ds DUCSENTETE
                                2- Correction Access Vio Sur V 4.0.1.582.08 qd Néant
                                3- L'abrégé de période erroné qd année < 2000
                                4- Correction mauvais ordre de présentation qd
                                le code institution est renseigné sur un URSSAF
                                par exemple : le code institution dans DUCSDETAIL
                                n'est renseigné que pour les IRC
 PT19 : 05/08/2002 : V585 MF
                                1 - Les champs Ligne optique , N° MSA et N° Interne
                                sont tronqués pour alimenter DUCSENTETE.
                                2 - La ligne optique a pu être déjà renseignée
                                au niveau du paramètrage de l'organisme. Permet
                                de gérer les cas IRC non prévus dans le cahier des
                                charges. L'alimentation de ce champ ce fait selon
                                l'accord établi avec l'organisme.
                                3- Mise au point cas des regroupements de caisses
 PT20 : 03/10/2002 : V585 MF    Diverses modifications pour mise au point du
                                traitement des ruptures, du regroupement de DUCS
                                (mono et multi établissement), des DUCS IRC
                                persionnalisées.
                                NB : suppression des divers commentaires de
                                modification venant surcharger la lecture du code
                                sauvegarde du source sur
 PT21 : 14/10/2002 : V585 MF    Modification du calcul des dates d'exigibilité,
                                limite de dépôt, de règlement. Ces dates ne sont
                                plus calculées par rapport à la date de versement
                                des salaires mais à partir de la date de fin de
                                période de la déclaration. Règle le problème
                                rencontré lorsqu'il n'y a pas eu de bulletin pour
                                le dernier mois de la période.
 PT22 : 17/10/2002 : V585 MF
                                1-Rectification du traitement de la TOB TCot2 :
                                Les doubles sont arrondis, Le nom des champs ne
                                comporte pas de préfixe de table - Corrige pb
                                effectif ligne si bulletin complémentaire
                                2-Initialisation des doubles à 0.00, arrondi sur
                                les doubles
                                3-Pour limiter les requêtes,la recherche de la
                                codification taux n'est plus systématique.
 PT23 : 22/10/2002 : V585 MF    1- Optimisation du traitement : changement de
                                méthode pour le traitement des TOB.
 PT24 : 08/11/2002 : V591 MF    Correction : message indice de liste hors limite
                                quand ducs Néant
 PT25 : 06/01/2003 : V591 MF
                                1- Correction des avertissements de compile
 PT26 : 09/01/2003 : V591 MF
                                1- Correction : Ajout tests des champs des
                                TOB qui pourraient être Null avant récupération
                                par GetValue
 PT27 : 04/02/2003 : V591 MF    Alimentation de l'abrégé de période à 0000 quand
                                période semestrielle.

 PT29 : 14/02/2003 : V595 MF    Correction calcul des dates
 PT30 : 21/02/2003 : V 4.2.0 MF
                                1-Suppression de la notion d'établissement
                                principal
                                Il faut utiliser "Caisse destinataire"
                                pour une "ducs Dossier" et un "regroupement de
                                caisses"
                                2-On récupère systématiquement le code
                                institution dans la liste dee cotisations
                                3-Pour IRC (standardisée)- La codification est
                                modifiée en fct de la rupture/institution afin de
                                ne pas créer de codification duppliquée.
                                Incrémentation des 2 premières positions du code
 PT31 : 27/02/2003 : V  4.2.0 MF
                                1- Suite mise en place MEMCHECK - rectification
                                du traitement de libération des TOB
 PT32 : 12/03/03 :  V 4.2.0 MF
                                1- Cas ducs dossier : calcul des effectifs, il
                                   faut ne sélectionner que les établissements
                                   pour lesquels l'organisme traité est coché
                                   Ducs Dossier
                                2- Correction du calcul du nbre de cadres
 PT33 : 13/03/2003 : V 4.2.0 MF
                                1- Fiche qualité 10551 :Cas regroupement : Ne
                                sélectionner que les organismes ayant une même
                                nature Ducs.
                                2- Fiche qualité 10513 : si le n° siret
                                établissement n'est pas renseigné la ducs n'est
                                pas générée.
 PT34 : 14/03/2003 : V 4.2.0 MF
                                1- Fiche qualité 10411 : IRC contrôle de la
                                numéricité des positions 1 et 2 de la
                                codification lors du traitement des lignes de ST
 PT35 : 14/04/2003 : V 4.2.0 MF
                                1- correction du traitement pour cas ducs multi
                                etab avec rupture. Quand l'ensemble des ducs
                                étaient néant ne générait que la dernière.
 PT36 : 28/04/2003 : V 4.2.0 MF
                                1- AGFF :  un sous-total institution particulier
                                --> les codifications commençant par 8 sont
                                traitées spécifiquement.
 PT37 : 14/05/2003 ; V 4.2.0 MF
                                1- Alimentation du n° émetteur quand changement
                                d'émetteur sur critères. Résout pb ducs EDI,
                                tests IRC cas avec tiers déclarant.
 PT38 : 02/07/2003 :  V_421 MF  Mise au point CWAS
 PT39 : 16/07/2003 :  V_421 MF  suite mise en place Loi Fillon les codifications
                                de type montant sont alimentées à 0 en base
                                et taux afin de n'avoir qu'une seule ligne de
                                cotisation. (remarque : les rub. alimentant la
                                codif sont de la forme base*taux. Le champ taux
                                est en fait un coeff calculé et donc il peut y
                                avoir autant de coeff que de salarié)
 PT40 : 21/07/2003 :  V_421 MF  Caisses IRC : Gestion du Centre Payeur
 PT41 : 06/08/2003 :  V_421 MF  Correction FQ 10773 : Anomalie quand bulletin
                                complémentaire.
 PT42 : 15/09/2003 :  V_421 MF  Correction FQ 10788 : Traitement des raccourcis
                                clavier (Dates de période)
 PT43 : 18/09/2003 :  V_421 MF  PgiBox remplace MessageAlerte
 PT44 : 23/09/2003 :  V_421 MF  FQ 10829 : pour les envois DUCS sélection possible
                             de l'émetteur.
 PT45 : 06/10/2003 :  V_421 MF  FQ 10874 : calcul des effectifs spécifique pour
                                les cotisation de type Montant (forfait)
 PT46 : 17/12/2003 :  V_50  MF  FQ 10957 : prise en compte de bulletin à cheval
                                sur pls mois
                                FQ 11011 : correction de l'effectif pour les
                                codifications de type "Montant" (cas de la
                                réduction fillon)
 PT47 : 04/02/2004 :  V_50  SB  Reprise PT46 sur FQ 10957
                                Tenir compte de l'intervalle de début via la date de fin de paie
 PT48 : 09/02/2004 :  V_50  MF  FQ 11087 :  Le champ PDU_NBSALHAP (effectifs
                                déclaration hors apprentis) systématiquement à 0
                                car inutile
 PT49 : 23/02/2004 :  V_50  MF  1-FQ 10648 :  Mise en place des éditoons DUCS
                                dans les états chainés.
                                2-FQ 11065 : Correction du formatage du n°
                                d'institution (début de la codif) pour régler
                                erreur lors de la ducs EDI.
 PT50 : 22/03/2004 : V_50   MF  1- FQ 11201 : Correction ducs IRC, traitement
                                des sous-totaux, de la spécificité AGFF

 PT51 : 23/03/2004 : V_50   MF  FQ 11177 : correction rupture sur groupe interne
 PT52 : 25/07/2005 : V_604  MF  Compatibilité avec Socref 703 (prépa ducs edi v4.2)
 PT53 : 03/02/2006 : V_65   MF  FQ 12252 : DUCS IRC : le code institution alimente
                                chaque ligne même pour une ducs personnalisée.
 PT54 : 03/02/2006 : V_65   MF  FQ 12248 : vérification de l'existence d'un
                                émetteur par défaut. Si ce n'est pas le cas,
                                affichage d'un message d'info.
                                + Afin de pouvoir sélectionner parmis plusieurs
                                contacts DUCS mise en place d'un THEDIT sur la fiche
                                pour afficher N° et nom émetteur lors du choix
 PT55 : 08/02/2006 : V_65   MF  FQ 11357 :  une ligne détail n'est créée que si le
                                montant est différent de 0 ou
                                S'il s'agit d'une ligne de type Intitulé ou
                                S'il s'agit d'une ligne de type Montant et que taux
                                et montant sont différent de zéro  (cas loi Fillon)
 PT56 : 09/02/2006 : V_65   MF  FQ 12906 :  modification du calcul des dates
                                d'exigibilité, limite de dépôt, de règlement
 PT57 : 09/02/2006 : V_65   MF  DUCS EDI V4.2
 PT58 : 13/02/2006 : V_65   MF  FQ 11793 : Alimentation automatique Conditions spéciales
                                de cotisation
 PT59 : 29/03/2006 : V_65   MF  correction : règle PB base de cot absente pour un salarié (TDetCot.free)
 PT60 : 07/07/2006 : V_70   MF  DUCS EDI V4.2 : modifications suite évolution DRA 2005
 PT61 : 05/01/2007 : V_72   MF  FQ 13546  : correction :Qd rupture groupe interne on
                                n'éditait pas les bonnes coordonées établissement.
 PT62 : 29/01/2007 : V_72   MF  Modifs. Ducs V4.2
 PT63 : 02/02/2007 : V_72   MF  FQ 13070 : Traitement des codifications ALSACE MOSELLE

 PT64 : 09/02/2007 : V_72   FC Suivant un paramètre société, contrer l'habilitation qui est faite automatique
                            en effaçant le contenu du champ XX_WHERE2
 PT65 : 31/03/2007 : V_72   MF modifications liées aux éditions chainés , Process Server
 PT66 : 03/04/2007 : V_72   MF FQ 14095 : correction pb ducs IRC personnalisée
 PT67 : 13/04/2007 : V_72   MF FQ 14044 : Bouton Bouvrir replacer par bouton Blance
 PT68 : 16/04/2007 : V_72   MF FQ 14032 : Modif. du calcul date exigibilité....
 PT69 : 18/04/2007 : V_72   MF FQ 13969 : modif traitement des effectifs.
 PT70 : 27/06/2007 : V_72   MF FQ 14473 : alimentation des type de bordereau quand non renseignés
 PT71 : 16/07/2007 : V_72   MF FQ 14032 : complément au calcul des dates
 PT72 : 01/08/2007 : V_72   MF FQ 14632 : correction traitement des rubriques de type "présentation par taux AT"
 PT73 : 10/10/2007 : V_80   MF mise en place de la loi TEPA
 PT74 : 12/10/2007 : V_80   MF FQ 14846 : mise à jour CodifAlsece sur ducsparam et ducaffect
 PT78 : 17/10/2007 : V_80   MF FQ 14840 : correction Erreur Maj Table DUCSDETAIL, si apprenti
                            sans taux AT et apprenti avec taux AT. Plus généralement quand il y a une
                            cotisation sur le déplafonné et par sur le plaflonné et inversement
 PT79 : 25/10/2007 : V_80   MF on revoit le message concernant l'absence d'émetteur
 PT80 : 25/10/2007 : V_72   MF FQ 14842 et 14840 : correction "erreur de maj ducsdetail"
                               + indice de liste hors limite
                               + correction différence ducs/etat des charges
 PT81 : 20/11/2007 : V_80   MF Correction cas où le salarié n'a pas de cotisation AT.
 PT82 : 10/01/2008 : V_802  MF  FQ 15036 - codifs Alsace-Moselle
 PT83 : 13/01/2008 : V_802  MF  FQ 15051 - Comptabilisation de l'effectifs des ouvriers
 PT84 : 13/01/2008 : V_802  MF  FQ 15040 - Exclusion des salariés MSA du calcul des effectifs
 PT85 : 23/01/2008 : V_802  MF  FQ 15155 - Correction : cas utilisation rubrique de régul. (modif montant)
 PT87 : 26/02/2008 : V_802  MF  FQ 15199 - Bulletin complémentaire
 PT88 : 26/02/2008 : V_802  MF  FQ 15238 - Rub AT mal présentée
                            + modif pour traiter Rub de régul. (Erreur Maj Table DUCSDETAIL)
 PT89 : 29/02/2008 : V_802  MF  FQ 15222 - ajout message lorsqu'il existe des rubriques de régularisation
 PT94 : 20/06/2008 : V_810  MF CWAS : on vide MaTob et MATobAT
 PT96 : 07/07/2008 : V_810  MF  FQ 15620 - correction de l'alimentation de la
                                           base qd la base M+1 annule la base N.
                                           Anomalie apportée par le PT87
                                           (Bulletin complémentaire)  pour lequel
                                           il faut cumuler les bases d'une même
                                           rubriques pour un salarié qd celles-ci
                                           sont différentes.
 PT98 : 15/09/2008  : MF  La méthode de regroupement des cotisation par Taux est revue.
                          Pour résoudre Pb rencontré avec les Paies au contrat :
                          des ruptures étaient générées à tort sur les lignes de
                          type "présentation par taux AT".
 }
unit UTofPG_MulDucsInit;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
{$IFDEF EAGLCLIENT}
  UtileAGL, eMul,
{$ELSE}
  HDB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Mul,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
{$ENDIF}
{$ENDIF}
  UTOF, HCtrls, ComCtrls, Hqry, PgOutils, Pgoutils2, sysutils, ParamDat, Classes, P5Util, Hent1,
  ParamSoc, HTB97, hmsgbox, ed_tools, UTob, HStatus, PGEdtEtat, EntPaie,
// d PT65
  P5Def,        // pour maj jnal événements
  UficheJob,    // pour progammation Process Server
  ULibEditionPaie;
type
// PT65  TOF_PGMULDUCSINIT = class(TOF_PGEtatMul)  les états chainés ne lancent plus la fiche MUL_INITDUCS
  TOF_PGMULDUCSINIT = class(TOF)
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;

  private
    DD, DF: TDateTime; // Etats chaînés
//    ND, NF, ED, EF  : string; // PT65
    WW: THEdit;
    DateDeb, DateFin: THEdit;
    Q_Mul: THQuery;
    DebPer, FinPer, PeriodEtat, NatDeb, NatFin, EtabDeb, EtabFin: string;
    Etab, Organisme, EtSiret, EtApe, PogNumInterne: string;
    NbreMois, CodifPosDebut, CodifLongEdit: Integer;
    JExigible, JLimDepot, JReglement: Integer;
{$IFDEF DUCS41}
    EffTot, EffHom, EffFem, EffApp, EffCAd, EffOuv, EffEtam : Integer;
{$ENDIF}
    InstEnCours, CodifEnCours, CodifST, BaseArr: string;
    WBaseArr, MontArr, WMontArr: string;
    DebTrim, FinTrim, DebSem, FinSem, DatePaye: TDateTime;
    DebExer, FinExer: TDateTime;
    TDucs, TInst, TCot2: TOB;
    Declarant, Emetteur, RaisonSoc, ContactDecla, TelDeclarant, FaxDeclarant: string;
    IndNeant, IndDucs: boolean;
    IndVLU: boolean;
    Siretlu, Apelu, NumIntlu, GpIntlu: string;
    WSiret, WApe, WNumInt, WGpInt: string;
    Rupture, RupSiret, RupApe, RupNumInt, RupGpInt, IndNeant2: Boolean;
    TRupMere, TRupFille, TRup: TOB;
    SiretEnCours, ApeEnCours, EtabEnCours, OrganismeEnCours: string;
    NumIntEnCours, GpIntEnCours: string;
    NumDucs, NoMere: integer;
    StWhere, StWhereDet, StWhereEd: string;
    PogLgOptique: string;
    LGOptiqueEnCours: string;
    ErrRupGpInt: integer;
// d  DUCS EDI V4.2
{$IFNDEF DUCS41}
    PogTypBordereau                                         : string;
    EffAppH, EffAppF, EffCadH, EffCadF                      : double;
    EffTot, EffHom, EffFem, EffApp, EffCAd, EffOuv, EffEtam : double;
//PT62    Eff65H, Eff65F, EffCesH, EffCesF                        : double;
    Eff65H, Eff65F, EffProfH, EffProfF                        : double;   // PT62
    EffCDIH, EffCDIF                                        : double;
    EffCDDH, EffCDDF, EffCNEH, EffCNEF                      : double;
    EffCadH22, EffCadH23, EffCadH24, EffCadH25              : double; // PT62
    EffCadF22, EffCadF23, EffCadF24, EffCadF25              : double; // PT62
{$ENDIF}
// f  DUCS EDI V4.2
    CondSpec                                                : Boolean; 
    TCot3                                                   : TOB;     //PT62
//    Initialisation : Boolean; // PT65 quand = True, on lance linitialisation seulement menu 42301
    procedure ActiveWhere(Okok: Boolean);
    procedure DateElipsisclick(Sender: TObject);
    procedure CalculerClick(Sender: TObject);
    procedure ChargeDucsAfaire();
    procedure ChargeListeInstit(DucsDoss: Boolean; Regroupt, NatureDucs: string); 
// PT89 FQ 15222    procedure ChargeCot(DucsDoss: Boolean; Regroupt: string);
    procedure ChargeCot(DucsDoss: Boolean; Regroupt: string; TraceE : TStringList);
    procedure InitVarTrait(DucsDoss: Boolean; Regroupt: string);
    procedure ChargeWheres(Regroupt: string); // PT15-1
    procedure MajDucs(DucsDoss: Boolean; Regroupt: string; TraceE : TStringList); // PT65 +Trace dans jnal des événements
    procedure AlimDetail(var TOB_FilleLignes, TOB_FilleTete: TOB; var ligne: Integer; Regroupt: string);
    procedure Change(Sender: TObject);
    procedure ChercheClick(Sender: TObject);
    procedure ChangeEmetteur(Sender: TObject);
    function fctRupture(TDetail: TOB; Ind: integer): Boolean;
    procedure RechRupture(Organisme: string; DucsDoss: Boolean; Regroupt: string); //  PT15-1
    procedure LibereTobs();
//PT84   procedure CalculEffectifs(DucsDoss: Boolean; Regroupt: string);
    procedure CalculEffectifs(DucsDoss: Boolean; Regroupt, PogNature: string);
    procedure EditerDucs(Pages: TPageControl);
// d PT65 : Pour programmation Process server
    procedure ParamInitDucs(Sender: TObject);
    procedure OnChangeNat(Sender: TObject);   // voir si utile ??
    procedure OnChangeEtab(Sender: TObject);  // voir si utile ??
// f PT65
// d PT70
{$IFNDEF DUCS41}
    Function  AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
{$ENDIF}
// f PT70
  end;

implementation

uses
  PGVLU;
//d  PT63
var
    AlsaceMoselle : boolean;
// f PT63
    INIT : boolean; // PT65 = True quand Initiatialisation seulement
{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 25/07/2001
Modifié le ... : 25/07/2001
Description .. : OnArgument :
Suite ........ : Récupération de l'établissement par défaut.
Suite ........ : Détermination de la Date par défaut
Suite ........ : Initialisations diverses
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.OnArgument(Arguments: string);
var
  ExerPerEncours, MoisE, AnneeE: string;
  Ouvrir, Cherche, bSelectAll: TToolbarButton97; // PT65 pour gérer le bouton bSelectAll
  WDeclarant: THValComboBox;
  Q: Tquery;

// d PT65
{$IFDEF EAGLCLIENT}
  Liste: THGrid;
{$ELSE}
  Liste: THDBGrid;
{$ENDIF}
  BTNPROCESS: TToolbarButton97;        // bouton de programmation du process server
  NATURED, NATUREF, ETABD,ETABF : THValComboBox;  // voir si utile ??
// f PT65
begin
  inherited;
  DD := idate1900; DF := idate1900;

// d PT65    Qd Initialisation seulement (appel menu 42301 Toutes les elts de la liste sont sélectionnés
//           le bouton bSelectall est invisible
  INIT := False;

  if trim(Arguments) = 'INIT' then
      INIT := True
  else
      INIT := False;

  if (INIT) then
  begin
    TFMul(Ecran).Caption := 'Initialisation des DUCS';
    UpdateCaption(TFMul(Ecran));

    bSelectAll :=  TToolbarButton97(GetControl('bSelectAll'));
    bSelectAll.Visible := False;
{$IFDEF EAGLCLIENT}
    Liste := THGrid(GetControl('FListe'));
{$ELSE}
    Liste := THDBGrid(GetControl('FListe'));
{$ENDIF}
    Liste.AllSelected := TRUE;
{$IFNDEF EAGLCLIENT}
    Liste.MultiSelection := False;
{$ENDIF}
  end;
// f PT65

  WW := THEdit(GetControl('XX_WHERE'));
  DateDeb := THEdit(getcontrol('XX_VARIABLED'));
  DateFin := THEdit(getcontrol('XX_VARIABLED_'));
  SetControlText('EMETSOC2', GetParamSocSecur('SO_PGEMETTEUR','')); 
  Emetteur := GetControlText('EMETSOC2');

  WDeclarant := THValComboBox(getcontrol('EMETSOC2'));
  Q := OpenSQL('SELECT PET_EMETTSOC,PET_RAISONSOC,PET_CONTACTDUCS,' +
    'PET_TELDUCS,PET_FAXDUCS FROM EMETTEURSOCIAL' +
    ' where PET_EMETTSOC="' + Emetteur + '"', True); 
  if not Q.EOF then
  begin
    RaisonSoc := copy(Q.FindField('PET_RAISONSOC').AsString, 1, 35);
    ContactDecla := Q.FindField('PET_CONTACTDUCS').AsString;
    TelDeclarant := Q.FindField('PET_TELDUCS').AsString;
    FaxDeclarant := Q.FindField('PET_FAXDUCS').AsString;
  end;
  ferme(Q);
  WDeclarant.Text := Emetteur;
  if (WDeclarant <> nil) then
    WDeclarant.OnExit := ChangeEmetteur;

  if (DateDeb <> nil) and (DateFin <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateDeb.OnExit := Change;
    DateFin.OnElipsisClick := DateElipsisclick;
    DateFin.OnExit := Change;
  end;

// d PT65
  if ((DD = idate1900) and (DF = idate1900)) then
  // si les dates ne sont pas renseignées
  begin
// f PT65
  { Date par défaut : La date proposée est le trimestre en cours si la période
    en cours correspond à une fin de trimestre, sinon c'est le mois en cours.}
    if RendExerSocialEnCours(MoisE, AnneeE, ExerPerEncours, DebExer, FinExer) = True then
    begin
      RendPeriodeEnCours(ExerPerEnCours, DebPer, FinPer);
      RendTrimestreEnCours(StrToDate(DebPer), DebExer, FinExer, DebTrim, FinTrim, DebSem, FinSem);
      if FindeMois(StrToDate(DebPer)) <> FinTrim then
      { mois en cours}
      begin
        if DateDeb <> nil then DateDeb.text := DateToStr(StrToDate(DebPer));
        if DateFin <> nil then DateFin.text := DateToStr(FindeMois(StrToDate(DebPer)));
      end
      else
      { trimestre en cours}
      begin
        if DateDeb <> nil then DateDeb.text := DateToStr(DebTrim);
        if DateFin <> nil then DateFin.text := DateToStr(FinTrim);
      end;
    end;
// d PT65
  end
  else
  // les dates sont renseignées
  begin
    DebPer := DateToStr(DD);
    FinPer := DateToStr(DF);
  end;
// f PT65

// d PT65 voir si utile ??
  NATURED := ThValComboBox(GetControl('POG_NATUREORG'));
  if NATURED <> nil then NATURED.OnChange := OnChangeNat;
  NATUREF := ThValComboBox(GetControl('POG_NATUREORG_'));
  if NATUREF <> nil then NATUREF.OnChange := OnChangeNat;
  ETABD := ThValComboBox(GetControl('POG_ETABLISSEMENT'));
  if ETABD <> nil then ETABD.OnChange := OnChangeEtab;
  ETABF := ThValComboBox(GetControl('POG_ETABLISSEMENT_'));
  if ETABF <> nil then ETABF.OnChange := OnChangeEtab;
// f PT65

// PT67  Ouvrir := TToolbarButton97(GetControl('BOUVRIR'));
  Ouvrir := TToolbarButton97(GetControl('BLANCE'));
  if Ouvrir <> nil then
  begin
    Ouvrir.Visible := True;
    Ouvrir.OnClick := CalculerClick;
  end;

  Cherche := TToolbarButton97(GetControl('BCHERCHE'));
  if Cherche <> nil then
  begin
    Cherche.OnClick := ChercheClick;
  end;

// d PT65  C'est à partir du traitement d'initialisation (menu 42301) que l'on peut
//         programmer le Process server
  if (not INIT) then
  begin
    SetControlVisible ('BPROG', false);
    SetControlEnabled ('BPROG', false);
  end
  else
  begin
    BTNPROCESS := TToolbarButton97(GetControl('BPROG'));
    if BTNPROCESS <> nil then BTNPROCESS.OnClick := ParamInitDucs;
  end;
// f PT65
// d PT74
 Q := OpenSQL('select pdp_codification, pdp_codifAlsace from ducsparam where pdp_Codification  not like"1%" and pdp_codifAlsace =""', True);
  if not Q.EOF then
  begin
    ExecuteSQL('update ducsparam  set pdp_Codifalsace = pdp_Codification  where pdp_Codification  not like"1%"');
  end;
  ferme(Q);
 Q := OpenSQL('select pdf_codification, pdf_codifAlsace from ducsaffect where pdf_Codification  not like"1%" and pdf_codifAlsace =""', True);
  if not Q.EOF then
  begin
    ExecuteSQL('update ducsaffect  set pdf_Codifalsace = pdf_Codification  where pdf_Codification  not like"1%"');
  end;
  ferme(Q);
// f PT74

end; { fin OnArgument}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 25/07/2001
Modifié le ... : 26/07/2001
Description .. : OnLoad : Chargement de la fiche
Suite ........ : On détermine la périodicité des DUCS à initialiser
Suite ........ : On lance l'ActiveWhere
Suite ........ : On récupère les critères de sélection.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.OnLoad;
var
  Okok: Boolean;
  NbreJour: Integer;
  Q: TQuery;
begin
  inherited;

{ Récup. des critères de sélection}
  DebPer := GetControlText('XX_VARIABLED');
  FinPer := GetControlText('XX_VARIABLED_');
  NatDeb := GetControlText('POG_NATUREORG');
  NatFin := GetControlText('POG_NATUREORG_');
  EtabDeb := GetControlText('POG_ETABLISSEMENT');
  EtabFin := GetControlText('POG_ETABLISSEMENT_');

  Declarant := GetControlText('EMETSOC2');
  Q := OpenSQL('SELECT PET_EMETTSOC,PET_RAISONSOC,PET_CONTACTDUCS,PET_TELDUCS' +
    ',PET_FAXDUCS FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' +
    Declarant + '"', True);
  if not Q.EOF then
  begin
    RaisonSoc := copy(Q.FindField('PET_RAISONSOC').AsString, 1, 35);
    ContactDecla := Q.FindField('PET_CONTACTDUCS').AsString;
    TelDeclarant := Q.FindField('PET_TELDUCS').AsString;
    FaxDeclarant := Q.FindField('PET_FAXDUCS').AsString;
  end;
  ferme(Q);

{ On détermine la périodicité des DUCS à initialiser
  Mensuelle, Trimestrielle, Semestrielle, Annuelle ou Quelconque
  Remarque : Il faut que la période corresponde à un mois ou un trimestre
  ou un semestre de l'exercice (Par exemple Février/mars/avril n'est pas un
  trimestre) sinon elle sera considérée comme quelconque (PeriodEtat = '')}
  if ((IsValidDate(DebPer)) and (IsValidDate(FinPer))) then
  begin 
    DiffMoisJour(StrToDate(DebPer), StrToDate(FinPer), NbreMois, NbreJour);
    NbreMois := NbreMois + 1;

    if (NbreMois <> 1) and
      (NbreMois <> 3) and
      (NbreMois <> 6) and
      (NbreMois <> 12) then
    begin
      MessageAlerte('La période sélectionnée n''est pas un mois, ' +
        '#13#10 ni un trimestre, ni un semestre, ni une année ');
    end;

    PeriodEtat := 'M';
    if (NbreMois = 3) then PeriodEtat := 'T';
    if (NbreMois = 6) then PeriodEtat := 'S';
    if (NbreMois = 12) then PeriodEtat := 'A';

    Okok := TRUE;
    ActiveWhere(Okok);
  end; 
end; { fin OnLoad}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 25/07/2001
Modifié le ... :   /  /
Description .. : Affichage calendrier sur champs dates
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDUCSINIT.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end; {fin DateElipsisclick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 26/07/2001
Modifié le ... : 26/07/2001
Description .. : ActiveWhere : Compléments aux critères de sélection de la
Suite ........ : fiche.
Suite ........ : on recherche les DUCS à initialser sur la période
Suite ........ : sélectionnée, quand il s'agit de DUCS dossier on ne
Suite ........ : propose que l'établissement principal (établissement par
Suite ........ : défaut des paramètres société)
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ActiveWhere(Okok: Boolean);
var
  St: string;
  Where2:THEdit;//PT64
begin
  WW.Text := '';
  st := '';
  St := '((POG_PERIODICITDUCS="' + PeriodEtat + '") OR ' +
    '((POG_AUTREPERIODUCS<>"") AND (POG_AUTREPERIODUCS="' + PeriodEtat + '"))) ' +
    'AND' +
  '((POG_CAISSEDESTIN="X") OR ' +
    '((POG_CAISSEDESTIN<>"X") AND (POG_DUCSDOSSIER<>"X")))' +
    ' AND ' +
  '((POG_REGROUPEMENT="") OR (POG_REGROUPEMENT IS NULL)OR (POG_REGROUPEMENT<>"" AND POG_CAISSEDESTIN="X"))';

  if St <> '' then WW.Text := st;
  if Q_Mul <> nil then
  begin
    TFMul(Ecran).SetDBListe('PGDUCSINIT');
  end;

  //DEB PT64
  if GetParamSocSecur('SO_PGDRTVISUETAB',True) then
  begin
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Where2 <> nil then SetControlText('XX_WHERE2', '');
  end;
  //FIN PT64
end; {fin ActiveWhere}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 26/07/2001/
Modifié le ... : 26/07/2001
Description .. : CalculClick : Lance l'initialisation des DUCS sélectionnées.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.CalculerClick(Sender: TObject);
var

{$IFDEF EAGLCLIENT}
  Liste: THGrid;
{$ELSE}
  Liste: THDBGrid;
{$ENDIF}
   //TProgress: TQRProgressForm ; PORTAGECWAS
  Regroupt: string;
  i, IndOR: integer;
  DucsDoss, PaieGroupe: Boolean;
  Pages: TPageControl;
  NatureDucs: string;
  Trace, TraceE: TStringList; // PT65 - Pour trace dans jnal des événements
  StSQL : string; // PT65 - paramètre d'appel à EditVLU

begin
// PT65 - Pour trace dans jnal des événements
  Trace := TStringList.Create;
  TraceE := TStringList.Create;

  Trace.Add ('Traitement interactif des déclarations du '+ DebPer + ' au '+ FinPer);
  if (declarant = '') then
  begin
//PT79    PgiBox('Aucun émetteur par défaut n''est renseigné. Information obligatoire pour la Ducs EDI', Ecran.Caption);
//PT79  TraceE.add ('Aucun émetteur par défaut n''est renseigné.I nformation obligatoire pour la Ducs EDI');
    PgiBox('Aucun déclarant (émetteur) par défaut n''est renseigné. '+
           #13#10+' Cette information doit être présente sur la ducs papier.'+
           #13#10+' Information obligatoire pour la Ducs EDI.', Ecran.Caption);
    TraceE.add ('Aucun déclarant (émetteur) par défaut n''est renseigné.'+
                #13#10+' Cette information doit être présente sur la ducs papier.'+
                #13#10+' Information obligatoire pour la Ducs EDI.');
  end;
// f PT65

  IndNeant := FALSE;
  IndDucs := FALSE;
  IndVLU := FALSE;

  Pages := TPageControl(GetControl('Pages'));

  i := 0;
{$IFDEF EAGLCLIENT}
  Liste := THGrid(GetControl('FListe'));
{$ELSE}
  Liste := THDBGrid(GetControl('FListe'));
{$ENDIF}
// PT65  if Liste <> nil then
  if (Liste <> nil)  and (not INIT) then
  // si traitement Edition des DUCS (menu 42354)
    if (Liste.NbSelected = 0) and (not Liste.AllSelected) then
    begin
      PgiBox('Aucun élément sélectionné', Ecran.Caption);
      exit;
    end;

  if (INIT) then
  // si tratement Initialisation DUCS (menu 42301)
  begin
    Liste.AllSelected := TRUE;
{$IFNDEF EAGLCLIENT}
    Liste.MultiSelection := False;
{$ENDIF}
  end;
// f PT65

  { Chargt. tob TDucs - ducs à initialiser à partir de la table ORGANISMEPAIE}
  ChargeDucsAfaire();

  StWhereEd := 'WHERE (';
  StWhere := 'WHERE (';
  StWhereDet := 'WHERE (';
  IndOR := 0;

  { Sélection des DUCS à Initialiser}
  if Liste <> nil then
  begin
    if (Liste.AllSelected = TRUE) then
    begin
// d PT65
      if (not INIT) then
      //si traitement Edition des DUCS (menu 42354)
      begin
        InitMoveProgressForm(nil, 'Initialisation en cours',
                                  'Veuillez patienter SVP ...', i, FALSE, TRUE);

        InitMove(TFmul(Ecran).Q.RecordCount, '');
      end;
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin

      if (not INIT) then
      // si tratement Initialisation DUCS (menu 42301)
          MoveCur(False);
// f PT65

        DucsDoss := TFmul(Ecran).Q.FindField('POG_DUCSDOSSIER').asstring = 'X';
        PaieGroupe := TFmul(Ecran).Q.FindField('POG_PAIEGROUPE').asstring = 'X';
        Regroupt := TFmul(Ecran).Q.FindField('POG_REGROUPEMENT').asstring;
        Etab := TFmul(Ecran).Q.FindField('POG_ETABLISSEMENT').asstring;
        Organisme := TFmul(Ecran).Q.FindField('POG_ORGANISME').asstring;
        NatureDucs := TFmul(Ecran).Q.FindField('POG_NATUREORG').asstring;
// d PT65 Maj jnal des événements
        Trace.Add ('');
        Trace.Add ('Etablissement '+ Etab + ':' +
                   'Déclaration '+ TFmul(Ecran).Q.FindField('POG_LIBELLE').asstring);
// f PT65
        if (organisme = '002') and
          (DucsDoss = True) and
          (PaieGroupe = True) then IndVLU := TRUE;

        {On établit la liste des institutions rattachées
         à l'organisme destinataire}
        ChargeListeInstit(DucsDoss, Regroupt, NatureDucs);

        {Initialisation variables de traitement}
        InitVarTrait(DucsDoss, Regroupt);

        { Mémorisation en vue d'édition et de mise à jour des critères de
          sélection des différentes DUCS initialisées}
        if (IndOR <> 0) then
        begin
          StWhere := 'WHERE (';
          StWhereDet := 'WHERE (';
          StWhereEd := StWhereEd + 'OR ';
        end;

        {Initialisation des strings "where" en fonctions des ruptures
         paramètrées en vue de la suppression des ducs déjà existantes
         pour les mêmes critères}
        ChargeWheres(Regroupt);

        IndOR := 1;

        {Liste des cotisations regroupées
         par codification , par tauxat, par taux global et
         par condition de rupture
         pour chaque salarié par base identique pour une période.}
        ErrRupGpInt := 0;
// PT89 FQ 15222        ChargeCot(DucsDoss, Regroupt);
        ChargeCot(DucsDoss, Regroupt,TraceE);
        if (ErrRupGpInt <> 0) then break;


        {Mise à jour des lignes détail à partir de la TOB TCot  }
        MajDucs(DucsDoss, Regroupt, TraceE);  // PT65 +Trace dans jnal des événements

        { free des TOBs utilisées}
        LibereTobs();

        if (not INIT) then // PT65
        //si traitement Edition des DUCS (menu 42354)
          MoveCurProgressForm(Organisme);

        TFmul(Ecran).Q.Next;
      end;

      Liste.AllSelected := False;

// d PT65
      if (not INIT) then
      //si traitement Edition des DUCS (menu 42354)
      begin
        FiniMove;
        FiniMoveProgressForm;

        TFMul(Ecran).bSelectAll.Down := False;
      end;
// f PT65

    end
    else
    begin
      InitMoveProgressForm(nil, 'Initialisation en cours',
        'Veuillez patienter SVP ...', i, FALSE, TRUE);

      InitMove(Liste.NbSelected, '');
      for i := 0 to Liste.NbSelected - 1 do
      begin
        Liste.GotoLeBOOKMARK(i);
        MoveCur(False);
{$IFDEF EAGLCLIENT}
        TFmul(Ecran).Q.TQ.Seek(TFmul(Ecran).FListe.Row - 1);
{$ENDIF}
        DucsDoss := TFmul(Ecran).Q.FindField('POG_DUCSDOSSIER').asstring = 'X';
        PaieGroupe := TFmul(Ecran).Q.FindField('POG_PAIEGROUPE').asstring = 'X';
        Regroupt := TFmul(Ecran).Q.FindField('POG_REGROUPEMENT').asstring;
        Etab := TFmul(Ecran).Q.FindField('POG_ETABLISSEMENT').asstring;
        Organisme := TFmul(Ecran).Q.FindField('POG_ORGANISME').asstring;
        NatureDucs := TFmul(Ecran).Q.FindField('POG_NATUREORG').asstring;

        if (organisme = '002') and
          (DucsDoss = True) and
          (PaieGroupe = True) then IndVLU := TRUE;

        {On établit la liste des institutions rattachées
         à l'organisme destinataire}
        ChargeListeInstit(DucsDoss, Regroupt, NatureDucs);

        {Initialisation variables de traitement}
        InitVarTrait(DucsDoss, Regroupt);

        {Mémorisation en vue d'édition des critères de sélection des différentes
         DUCS initialisé}
        if (IndOR <> 0) then
        begin
          StWhere := 'WHERE (';
          StWhereDet := 'WHERE (';
          StWhereEd := StWhereEd + 'OR ';
        end;

        {Initialisation des strings "where" en fonctions des ruptures
         paramètrées en vue de la suppression des ducs déjà existantes
         pour les mêmes critères}
        ChargeWheres(Regroupt);

        IndOR := 1;

        {Liste des cotisations regroupées
         par codification , par tauxat, par taux global et
         pour chaque salarié par base identique pour une période.}
        ErrRupGpInt := 0;
// PT89  FQ 15222      ChargeCot(DucsDoss, Regroupt);
        ChargeCot(DucsDoss, Regroupt,TraceE);
        if (ErrRupGpInt <> 0) then break;

        {Mise à jour des lignes détail à partir de la TOB TCot}
        MajDucs(DucsDoss, Regroupt, TraceE);    // PT65 +Trace dans jnal des événements

        { free des TOBs utilisées}
        LibereTobs();

        MoveCurProgressForm(Organisme);
      end;

      Liste.ClearSelected;
      FiniMove;
      FiniMoveProgressForm;
    end;
  end;

  if Tducs <> nil then
  begin
    TDucs.Free;
  end;

  if (ErrRupGpInt <> 0) then exit;

// d PT65
  if (not INIT) then
  //si traitement Edition des DUCS (menu 42354)
  begin
    { Edition effective des DUCS Précédemment sélectionnées.}
    EditerDucs(Pages);

    if (IndVLU = TRUE) and
      (IndNeant = False) then
    { VLU}
    begin
      EditVLU(StrToDate(Debper), StrToDate(FinPer), Organisme, Pages, False, True, StSQL); // PT65
    end;
  end;
    Trace.Add ('Fin du traitement interactif d''initialisation des DUCS');
    CreeJnalEvt('001', '020', 'OK', nil, nil, Trace, TraceE);
  Trace.Free;
  TraceE.Free;
  Pgiinfo('Traitement terminé...', Ecran.Caption );
// f PT65
end; {fin CalculerClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 26/07/2001
Modifié le ... : 01/08/2001
Description .. : On charge la tob TDucs des ducs à initialiser à partir de la
Suite ........ : table ORGANISMEPAIE
Suite ........ : Un organisme est mis en table Si l'établisst. fait partie de la
Suite ........ : fourchettes des établissts., si sa nature fait partie de la
Suite ........ : fourchettes des natures de DUCS, si une des périodicités
Suite ........ : paramétrées correspond à la période choisie.
Suite ........ : Si la DUCS est une DUCS dossier on ne charge que
Suite ........ : l'organisme de l'établissement principal
Suite ........ :
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ChargeDucsAfaire();
var
  st: string;
  QRechDucs: TQuery;

begin
  st := ' ';
  st := 'SELECT POG_AUTPERCALCUL,POG_AUTREPERIODUCS,POG_BASETYPARR,' +
    'POG_CAISSEDESTIN,POG_DUCEXIGIBILITE,POG_DUCLIMITEDEPOT,' +
    'POG_DUCSDOSSIER,POG_DUCSREGLEMENT,POG_ETABLISSEMENT,POG_INSTITUTION,' +
    'POG_LONGEDITABLE,POG_LONGTOTAL,POG_LONGTOTALE,POG_MTTYPARR,' +
    'POG_NATUREORG,POG_NUMINTERNE,POG_ORGANISME,POG_PERIODCALCUL,' +
    'POG_PERIODICITDUCS,POG_POSDEBUT,POG_POSTOTAL,POG_REGROUPEMENT,' +
    'POG_RUPTAPE,POG_RUPTGROUPE,POG_RUPTNUMERO,POG_RUPTSIRET,POG_SOUSTOTDUCS,' +
    'POG_CONDSPEC,'+  
    'POG_BASETYPARR,POG_MTTYPARR ' +
    'FROM ORGANISMEPAIE WHERE ';
  if (EtabDeb <> '') and (EtabFin <> '') then
    st := st + '(POG_ETABLISSEMENT >="' + EtabDeb + '" ' +
      'AND POG_ETABLISSEMENT <="' + EtabFin + '")';
  if (NatDeb <> '') and (NatFin <> '') then
  begin
    if (EtabDeb <> '') and (EtabFin <> '') then
      st := st + ' AND ';

    st := st + '(POG_NATUREORG >= "' + NatDeb + '" ' +
      'AND POG_NATUREORG <="' + NatFin + '")';
  end;
  if ((EtabDeb <> '') and (EtabFin <> '')) or
    ((NatDeb <> '') and (NatFin <> '')) then
    st := st + ' AND ';

  st := st + '((POG_PERIODICITDUCS="' + PeriodEtat + '") OR ' +
    '((POG_AUTREPERIODUCS<>"") AND (POG_AUTREPERIODUCS="' + PeriodEtat + '"))) ' +
    'AND' +
    '((POG_CAISSEDESTIN="X") OR ' +
    '((POG_CAISSEDESTIN<>"X") AND (POG_DUCSDOSSIER<>"X")))';

  QRechDucs := OpenSql(st, TRUE);
  TDucs := TOB.Create('Les DUCS', nil, -1);
  TDucs.LoadDetailDB('ORGANISMEPAIE', '', '', QRechDucs, False);
  Ferme(QRechDucs);
end; {fin ChargeDucsAfaire}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 26/07/2001
Modifié le ... : 26/07/2001
Description .. : On établit la TOB des organismes ayant même code
Suite ........ : regroupement
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ChargeListeInstit(DucsDoss: Boolean; Regroupt, NatureDucs: string);
var
  st: string;
  QRechInstit: TQuery;
begin
  st := ' ';
  if (Regroupt <> '') then
  begin
    if (DucsDoss = TRUE) then
    {Ducs Dossier}
    begin
      st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME,POG_INSTITUTION, ' +
        'POG_NUMINTERNE ' +
        'FROM ORGANISMEPAIE' +
        ' WHERE POG_REGROUPEMENT ="' + Regroupt + '"' +
        ' AND POG_NATUREORG ="' + NatureDucs + '"';
    end
    else
    { Ducs Etablissement }
    begin
      st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME,POG_INSTITUTION, ' +
        'POG_NUMINTERNE ' +
        'FROM ORGANISMEPAIE' +
        ' WHERE POG_REGROUPEMENT="' + Regroupt + '"' +
        ' AND POG_ETABLISSEMENT="' + Etab + '"' +
        ' AND POG_NATUREORG ="' + NatureDucs + '"'; 
    end;
  end
  else
  begin
    st := 'SELECT POG_ETABLISSEMENT, POG_ORGANISME, POG_INSTITUTION, ' +
      'POG_NUMINTERNE ' +
      'FROM ORGANISMEPAIE' +
      ' WHERE POG_ORGANISME="' + Organisme + '"' +
      ' AND POG_ETABLISSEMENT="' + Etab + '"';
  end;

  QRechInstit := OpenSql(st, TRUE);
  TInst := TOB.Create('Les institutions', nil, -1);
  TInst.LoadDetailDB('TABLEORGANISMEPAIE', '', '', QRechInstit, False);
  Ferme(QRechInstit);
end; {fin ChargeListeInstit}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 26/07/2001
Modifié le ... :   /  /
Description .. : Initialisation des variables de traitement
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDUCSINIT.InitVarTrait(DucsDoss: Boolean; Regroupt: string);
var
  QQ: Tquery;
  StQQ: string;
  TD: TOB;

begin
  { Traitement des ruptures}
  WSiret := '';
  WApe := '';
  WNumInt := '';
  WGpInt := '';

  StQQ := '';

  InstEnCours := '';
  CodifEnCours := '';

  EffTot := 0;
  EffHom := 0;
  EffFem := 0;
  EffApp := 0;
{$IFNDEF DUCS41}
  EffAppH  := 0;
  EffAppF  := 0;
{$ENDIF}
  EffCad := 0;
{$IFNDEF DUCS41}
  EffCadh := 0;
  EffCadF := 0;
{$ENDIF}
  EffEtam := 0;
  EffOuv := 0;

  {Récupératon des paramètres Organisme}
  TD := TDucs.FindFirst(['POG_ETABLISSEMENT', 'POG_ORGANISME'],
    [Etab, Organisme], TRUE);

  {Paramètres d'édition de la codification}
  CodifPosDebut := TD.GetValue('POG_POSDEBUT');
  CodifLongEdit := TD.GetValue('POG_LONGEDITABLE');

  {Booléen condition spéciale de cotisation alimentation auto O/N}
  if (TD.GetValue('POG_CONDSPEC') =  'X') then
     CondSpec :=True
  else
     CondSpec := False;

  {Paramètres pour le traitement des ruptures}
  NoMere := 0;
  RupSiret := FALSE;
  RupApe := FALSE;
  RupNumInt := FALSE;
  RupGpInt := FALSE;
  Rupture := FALSE;

  if (TD.GetValue('POG_RUPTGROUPE') = 'X') then
  begin
    RupGpint := TRUE;
    if (NoMere = 0) then NoMere := 1;
    Rupture := TRUE;
  end;

  //debut PT15-1
  {Quand Mono établissement on ne traite pas les ruptures Siret, Ape, N° interne}
  if (TD.GetValue('POG_DUCSDOSSIER') = 'X') then
  begin
    if (TD.GetValue('POG_RUPTAPE') = 'X') then
    begin
      RupApe := TRUE;
      if (NoMere = 0) then NoMere := 3;
      Rupture := TRUE;
    end;

    if (TD.GetValue('POG_RUPTSIRET') = 'X') then
    begin
      RupSiret := TRUE;
      if (NoMere = 0) then NoMere := 4;
      Rupture := TRUE;
    end;
  end;

  if (TD.GetValue('POG_RUPTNUMERO') = 'X') then
  begin
    RupNumInt := TRUE;
    if (NoMere = 0) then NoMere := 2;
    Rupture := TRUE;
  end;

  { Paramètres d'arrondi des bases et montants}
  BaseArr := TD.GetValue('POG_BASETYPARR');
  MontArr := TD.GetValue('POG_MTTYPARR');
  WBaseArr := BaseArr;
  WMontArr := MontArr;

  {paramètres de calcul des dates}
  JExigible  := TD.GetValue('POG_DUCEXIGIBILITE');
  JLimDepot := TD.GetValue('POG_DUCLIMITEDEPOT');
  JReglement := TD.GetValue('POG_DUCSREGLEMENT');

  {On établit les listes des valeurs possibles pour chaque critère de rupture}
  if (Rupture = TRUE) then
    RechRupture(Organisme, DucsDoss, Regroupt); // PT15-1

  StQQ := '';
  StQQ := '((PPU_ETABLISSEMENT ="' + Etab + '") AND ' +
    '(PPU_DATEDEBUT >="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
    '(PPU_DATEFIN <= "' + UsDateTime(StrToDate(FinPer)) + '")) ';
  QQ := OpenSQL('SELECT MAX(PPU_PAYELE) FROM PAIEENCOURS WHERE ' +
    StQQ, True);
  if not QQ.EOF then
  begin
    DatePaye := QQ.Fields[0].AsDateTime;
  end;
  Ferme(QQ);
  StQQ := '';

end; {fin InitVarTrait}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 31/07/2001
Modifié le ... : 01/08/2001
Description .. : ChargeCot : Chargement de la TOB TCot des cotisations
Suite ........ : regroupées par codification , par tauxat, par taux global et
Suite ........ : pour chaque salarié par base identique pour une période.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

//  PT89 FQ 15222 procedure TOF_PGMULDUCSINIT.ChargeCot(DucsDoss: Boolean; Regroupt: string);
procedure TOF_PGMULDUCSINIT.ChargeCot(DucsDoss: Boolean; Regroupt: string; TraceE :TStringList);
var
  QRechCot: TQuery;
  LesInst, TDetCot, TCot2Fille, TCot, TDetCod: TOB;
  StOrder, StOrder2, StOrder3, StOrder4, WsiretEnCours: string;
  WRubLu, WRubEnCours, WSalLu, WSalEnCours, WsiretLu: string;
  WNoIntLu, WNoIntEnCours, WApeLu, WApeEnCours, WGpLu: string;
  WCodEnCours, WCodLu, WOrgEnCours, WOrgLu, WGpEnCours: string;
  WInstLu, WInstEnCours, st, Organ, WcodPrec: string;
  WTxATLu, WTxATEncours, WTxGlLu, WTxGlEnCours: double;
  WBaseEnCours, WBaselu, TxGlobal, MtGlobal: double;
  WDateLu, WDateEnCours: TDateTime;
  WDateLuFin, WDateFinEnCours: TDateTime;
  IndFree, Personnalisee, WPrecExiste: Boolean;
  I, NbInst: integer;
  WMtGlobalLu, WMtGlobalEnCours : double; 
{* d PT62 }
   IndicD, IndicP, NbCodif                : integer;
   TauxglobalEnCours, TauxGlobalLu        : double;
   CodifW , SalW                          : string;
   DateDebW                               : TdateTime;
   TauxAtW                                : double;
   TCotA, TCotAFille,TCot3Fille    : TOB;
   TCot2Bis,TCot2BisFille                 : TOB;        // PT72
{* f PT62 }
   DateFinW                               : TdateTime;  // PT80
// PT98   BaseCotW                               : double;     // PT80
   Tcot3FilleBis                          : tob;        // PT81
   WBaseMaj                               : double;     // PT87
   IndCotRegul                            : boolean;    // PT89 FQ 15222
   RegulNegative                          : boolean;    // PT96
// d PT98
   SalEnCours,SalLu,WCodifEnCours         : string;
   BaseEnCours, Baselue,TauxAtEnCours     : double;
   TCotABis, TCotABisFille,TCotABisFilleP : TOB;
   INDIC                                  : integer;
// f PT98
begin
  WcodPrec := '';
  WPrecExiste := false;

  Personnalisee := False;

  {CHARGEMENT DE LA REQUETE DE SELECTION DES COTISATIONS A TRAITER}
  st := ' ';

// PT63   st := 'SELECT PHB_SALARIE,PDF_CODIFICATION,PHB_RUBRIQUE,PDF_PREDEFINI,' +
   st := 'SELECT PHB_SALARIE,PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,' +
         'PHB_COTREGUL,' + // PT89 FQ 15222
         'PDF_NODOSSIER,PHB_BASECOT,PHB_TAUXAT,' +
         'PHB_DATEFIN, ' +
         'PHB_DATEDEBUT,PHB_ORGANISME, ';

  st := st + 'POG_POSTOTAL, POG_LONGTOTAL, ';
  st := st + 'POG_INSTITUTION, ';
  st := st + 'ETB_REGIMEALSACE, '; //PT63

  {traitement des ruptures : il faut récupérer les champs sur lesquels une
   rupture est possible
  }
  if (RupSiret = TRUE) then
    st := st + 'ET_SIRET, ';
  if (RupApe = TRUE) then
    st := st + 'ET_APE, ';
  if (RupNumInt = TRUE) then
    st := st + 'POG_NUMINTERNE, ';
  if (RupGpInt = TRUE) then
    st := st + 'PGI_GROUPE, ';

  { Il est nécessaire de récuperer le n° d'institution pour pouvoir gérer
    les sous-totaux par institution
  }
  st := st + 'SUM(DISTINCT(PHB_TAUXSALARIAL+PHB_TAUXPATRONAL)) AS TAUXGLOBAL, ' +
    'SUM(DISTINCT(PHB_MTSALARIAL+PHB_MTPATRONAL)) AS MTGLOBAL ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN DUCSAFFECT ON PDF_RUBRIQUE=HISTOBULLETIN.PHB_RUBRIQUE ';

  { Traitement ruptures Siret ou APe:}
  if (RupSiret = TRUE) or (RupApe = TRUE) then
    st := st + ' LEFT JOIN ETABLISS ON ' +
      'ET_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';
// d PT63
    st := st + ' LEFT JOIN ETABCOMPL ON ' +
      'ETB_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';
// f PT63
  st := st + ' LEFT JOIN ORGANISMEPAIE ON ' +
    'POG_ORGANISME=HISTOBULLETIN.PHB_ORGANISME ' +
    ' AND POG_ETABLISSEMENT=HISTOBULLETIN.PHB_ETABLISSEMENT ';

  {Traitement rupture sur Groupe interne}
  if (RupGpInt = TRUE) then
  begin
    st := st + ' LEFT JOIN SALARIES ON PSA_SALARIE= HISTOBULLETIN.PHB_SALARIE ';
    st := st + ' LEFT JOIN GROUPEINTERNE ON PGI_CATEGORIECRC = ' +
               'SALARIES.PSA_DADSPROF ' +
               'AND PGI_INSTITUTION=ORGANISMEPAIE.POG_INSTITUTION ';
  end;

{* d PT63
    st := st + ' WHERE ##PDF_PREDEFINI## AND PHB_NATURERUB="COT" AND ' +
               'PHB_DATEFIN >="' + UsDateTime(StrToDATe(DebPer)) + '" AND ' +
               'PHB_DATEFIN<="' + UsDateTime(StrToDATe(FinPer)) + '" AND  ' +
               'PDF_CODIFICATION IS NOT NULL AND PDF_CODIFICATION <> "" AND ' +
               'PDF_CODIFICATION <> "       " AND ';  *}
    st := st + ' WHERE ##PDF_PREDEFINI## AND PHB_NATURERUB="COT" AND ' +
               'PHB_DATEFIN >="' + UsDateTime(StrToDATe(DebPer)) + '" AND ' +
               'PHB_DATEFIN<="' + UsDateTime(StrToDATe(FinPer)) + '" AND  ' +
               '((ETB_REGIMEALSACE = "X" AND  '+
               'PDF_CODIFALSACE IS NOT NULL AND PDF_CODIFALSACE <> "" AND ' +
               'PDF_CODIFALSACE <> "       " ) OR ' +
               '(ETB_REGIMEALSACE <> "X" AND  '+
               'PDF_CODIFICATION IS NOT NULL AND PDF_CODIFICATION <> "" AND ' +
               'PDF_CODIFICATION <> "       ")) AND ';
// f PT63

  if (DucsDoss <> TRUE) then
  {Ducs établissement}
    st := st + 'PHB_ETABLISSEMENT="' + Etab + '" AND '

  else
  {Ducs dossier
   exclure les établissements qui ne font pas partie de la Ducs Dossier}
    st := st + ' POG_DUCSDOSSIER="X" AND ';

  if (Regroupt <> '') then
  {Ducs regroupement d'organisme, il est nécessaire de récupérer les
   cotisations des différents caisses participant au groupe}
  begin
    for NbInst := 0 to TInst.Detail.Count - 1 do
    begin
      LesInst := TInst.Detail[NbInst];
      Organ := LesInst.GetValue('POG_ORGANISME');
      if (NbInst <> 0) then
      begin
        if (Pos('PHB_ORGANISME="' + Organ + '"', st) = 0) then
          st := st + ' OR '
      end
      else
        st := st + '(';

      if (Pos('PHB_ORGANISME="' + Organ + '"', st) = 0) then
        st := st + 'PHB_ORGANISME="' + Organ + '"';
    end;
    st := st + ')';
  end
  else
    st := st + 'PHB_ORGANISME="' + Organisme + '" ';

// d PT63
{*   st := st + ' AND (PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR ' +
               'PHB_MTPATRONAL <> 0) ' +
               'GROUP BY PHB_BASECOT,PHB_TAUXAT,PHB_DATEDEBUT,PHB_DATEFIN,PHB_SALARIE,' +
               'PDF_CODIFICATION,PHB_RUBRIQUE,PDF_PREDEFINI,PDF_NODOSSIER,' + *}
   st := st + ' AND (PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR ' +
               'PHB_MTPATRONAL <> 0) ' +
               'GROUP BY PHB_BASECOT,PHB_TAUXAT,PHB_DATEDEBUT,PHB_DATEFIN,PHB_SALARIE,' +
               'PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,PDF_NODOSSIER,' +
               'PHB_ORGANISME ';
// f PT63

  st := st + ',POG_INSTITUTION ';

  {traitement des ruptures : il faut regrouper les cotisations en fonctions des
   ruptures possibles}
  if (RupSiret = TRUE) then
    st := st + ',ET_SIRET ';
  if (RupApe = TRUE) then
    st := st + ',ET_APE ';
  if (RupNumInt = TRUE) then
    st := st + ',POG_NUMINTERNE ';
  if (RupGpInt = TRUE) then
    st := st + ',PGI_GROUPE ';

  st := st + ',POG_POSTOTAL, POG_LONGTOTAL ';
  st := st + ',ETB_REGIMEALSACE '; // PT63
  st := st + ',PHB_COTREGUL ';   // PT89 FQ 15222

  st := st + 'ORDER BY ';
  StOrder := '';
  StOrder4 := '';

  {traitement des ruptures}
  if (RupNumInt = TRUE) then
  begin
    st := st + 'POG_NUMINTERNE,';
    StOrder := StOrder + 'POG_NUMINTERNE;';
    StOrder4 := StOrder4 + 'NUMINTERNE;';
  end;
  if (RupSiret = TRUE) then
  begin
    st := st + 'ET_SIRET,';
    StOrder := StOrder + 'ET_SIRET;';
    StOrder4 := StOrder4 + 'SIRET;';
  end;
  if (RupApe = TRUE) then
  begin
    st := st + 'ET_APE,';
    StOrder := StOrder + 'ET_APE;';
    StOrder4 := StOrder4 + 'APE;';
  end;
  if (RupGpInt = TRUE) then
  begin
    st := st + 'PGI_GROUPE, ';
    StOrder := StOrder + 'PGI_GROUPE;';
    StOrder4 := StOrder4 + 'GROUPE;';
  end;

  StOrder2 := StOrder;
  StOrder3 := StOrder;

//d PT63
{*    st := st + 'PHB_ORGANISME,PDF_CODIFICATION,PHB_RUBRIQUE,PDF_PREDEFINI,' +
              'PDF_NODOSSIER,PHB_TAUXAT,TAUXGLOBAL,PHB_SALARIE';*}
   st := st + 'PHB_ORGANISME,PDF_CODIFICATION,PDF_CODIFALSACE,PHB_RUBRIQUE,PDF_PREDEFINI,' +
              'PDF_NODOSSIER,PHB_TAUXAT,TAUXGLOBAL,PHB_SALARIE';

{*    StOrder := StOrder + 'PHB_ORGANISME;PDF_CODIFICATION;PHB_RUBRIQUE;' +
                         'PDF_PREDEFINI;PDF_NODOSSIER;' +
                         'PHB_TAUXAT;TAUXGLOBAL;PHB_SALARIE';*}
   StOrder := StOrder + 'PHB_ORGANISME;PDF_CODIFICATION;PDF_CODIFALSACE;PHB_RUBRIQUE;' +
                        'PDF_PREDEFINI;PDF_NODOSSIER;' +
                        'PHB_TAUXAT;TAUXGLOBAL;PHB_SALARIE';

{*    StOrder2 := StOrder2 + 'PHB_ORGANISME;' +
                           'PHB_TAUXAT;TAUXGLOBAL;MTGLOBAL;PHB_SALARIE;' +
                           'PHB_DATEDEBUT;PHB_DATEFIN;PRED;PDF_CODIFICATION';*}
  StOrder2 := StOrder2 + 'PHB_ORGANISME;' +
                         'PHB_TAUXAT;TAUXGLOBAL;MTGLOBAL;PHB_SALARIE;' +
                         'PHB_DATEDEBUT;PHB_DATEFIN;PRED;PDF_CODIFICATION;PDF_CODIFALSACE';
{*    StOrder3 := StOrder3 + 'PHB_ORGANISME;PDF_CODIFICATION;PHB_TAUXAT;PHB_SALARIE;' +
                            'PHB_BASECOT;PHB_DATEDEBUT';*}

{* PT87 FQ 15199
    StOrder3 := StOrder3 + 'PHB_ORGANISME;PDF_CODIFICATION;PDF_CODIFALSACE;PHB_TAUXAT;PHB_SALARIE;' +
                            'PHB_BASECOT;PHB_DATEDEBUT';}
    StOrder3 := StOrder3 + 'PHB_ORGANISME;PDF_CODIFICATION;PDF_CODIFALSACE;PHB_TAUXAT;PHB_SALARIE;' +
                            'PHB_RUBRIQUE;PHB_BASECOT;PHB_DATEDEBUT';
// f PT63
  StOrder4 := StOrder4 + 'ORGANISME';

  if (Regroupt <> '') then
  begin
    st := st + ',POG_INSTITUTION';
    StOrder := StOrder + ';POG_INSTITUTION';
    StOrder2 := StOrder2 + ';POG_INSTITUTION';
    StOrder3 := StOrder3 + ';POG_INSTITUTION';
    StOrder4 := StOrder4 + ';INSTITUTION';
  end;

  StOrder4 := StOrder4 + ';CODIFICATION;TAUXAT;SALARIE;TAUXGLOBAL;' +	// PT98
              'BASECOT;DATEDEBUT';

  QRechCot := OpenSql(st, TRUE);
  TCot := TOB.Create('Les Cotisations', nil, -1);
  TCot.LoadDetailDB('TABHISTOBULLETIN', '', '', QRechCot, False);
  Ferme(QRechCot);

  if (RupGpInt = TRUE) then
  begin
    ErrRupGpInt := 0;
    for I := 0 to TCot.Detail.Count - 1 do
    begin
      TDetCod := TCot.Detail[I];
      if (TDetCod.GetValue('POG_INSTITUTION') = '') or
        (TDetCod.GetValue('POG_INSTITUTION') = NULL) then
      begin
        ErrRupGpInt := 1;
        break;
      end;
      if (TDetCod.GetValue('PGI_GROUPE') = '') or
        (TDetCod.GetValue('PGI_GROUPE') = NULL) then
      begin
        ErrRupGpInt := 2;
        break;
      end;
    end;
    if (ErrRupGpInt <> 0) then
    begin
      if (ErrRupGpInt = 1) then
        PGIError('Organisme sans code institution.', 'Rupture sur groupe interne impossible');
      if (ErrRupGpInt = 2) then
        PGIError('Certains groupes internes ne sont pas créés.', 'Rupture sur groupe interne impossible');
      exit;
    end;
  end;

  {On épure les codifications non utiles au dossier}
  for I := 0 to TCot.Detail.Count - 1 do
  {On attribue une priorité aux affectations}
  begin
    TDetCod := TCot.Detail[I];
    TDetCod.AddChampSup('PRED', FALSE);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'DOS') then
      TDetCod.PutValue('PRED', 1);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'STD') then
      TDetCod.PutValue('PRED', 2);
    if (TDetCod.GetValue('PDF_PREDEFINI') = 'CEG') then
      TDetCod.PutValue('PRED', 3);
  end;

  TCot.Detail.Sort('PHB_RUBRIQUE;' + StOrder2);

  TDetCod := TCot.FindFirst([''], [''], TRUE);

  WRubEnCours := '';
  WSalEnCours := '';
  WsiretEnCours := '';
  WNoIntEnCours := '';
  WApeEnCours := '';
  WGpEnCours := '';
  WInstEnCours := '';
  WTxATEncours := 0.00;
  WTxGlEnCours := 0.00;
  WDateEnCours := IDate1900;
  WDateFinEnCours := IDate1900;
  WMtGlobalEncours := 0.00;
  for I := 0 to Tcot.Detail.Count - 1 do
  {On épure les affectations qui ne seront pas utilisées}
  begin
    WRubLu := TDetCod.GetValue('PHB_RUBRIQUE');
    WSalLu := TDetCod.GetValue('PHB_SALARIE');
    if (RupGpInt = True) or
      ((Regroupt <> '') and
      (RupGpInt = False)) then
      if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
        WInstLu := TDetCod.GetValue('POG_INSTITUTION');
    if (RupSiret = TRUE) then
      if (TDetCod.GetValue('ET_SIRET') <> Null) then
        WsiretLu := TDetCod.GetValue('ET_SIRET');
    if (RupNumInt = TRUE) then
      if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
        WNoIntLu := TDetCod.GetValue('POG_NUMINTERNE');
    if (RupApe = TRUE) then
      if (TDetCod.GetValue('ET_APE') <> Null) then
        WApeLu := TDetCod.GetValue('ET_APE');
    if (RupGpInt = TRUE) then
      if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
        WGpLu := TDetCod.GetValue('PGI_GROUPE');
    WTxATLu := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
    WTxGlLu := Arrondi(TDetCod.GetValue('TAUXGLOBAL'), 5);
    WDateLu := TDetCod.GetValue('PHB_DATEDEBUT');
    WDateLuFin := TDetCod.GetValue('PHB_DATEFIN');
    WMtGlobalLu := TDetCod.GetValue('MTGLOBAL');
    IndFree := TRUE;
    if (WRubLu <> WRubEnCours) then IndFree := False;
    if (WSalLu <> WSalEnCours) then IndFree := False;
    if (WInstLu <> WInstEnCours) then IndFree := False;
    if (WTxATLu <> WTxATEnCours) then IndFree := False;
    if (WTxGlLu <> WTxGlEnCours) then IndFree := False;
    if (WMtGlobalLu <> WMtGlobalEnCours) then IndFree := False;
    if (WDateLu <> WDateEnCours) then IndFree := False;
    if (WDateLuFin <> WDateFinEnCours) then IndFree := False;
    if (RupSiret = TRUE) and (WSiretLU <> WsiretEnCours) then IndFree := False;
    if (RupNumInt = TRUE) and (WNoIntLu <> WNoIntEnCours) then IndFree := False;
    if (RupApe = TRUE) and (WApeLu <> WApeEnCours) then IndFree := False;
    if (RupGpInt = TRUE) and (WGpLu <> WGpEnCours) then IndFree := False;

    if (IndFree = True) then
      TDetCod.Free;

    WRubEnCours := WRubLu;
    WSalEnCours := WSalLu;
    WInstEnCours := WInstLu;
    WsiretEnCours := WsiretLu;
    WNoIntEnCours := WNoIntLu;
    WApeEnCours := WApeLu;
    WGpEnCours := WGpLU;
    WTxATEncours := WTxATLu;
    WTxGlEnCours := WTxGlLu;
    WDateEnCours := WDateLu;
    WMtGlobalEnCours := WMtGlobalLu ;
    WDateFinEnCours := WDateLuFin;

    TDetCod := TCot.FindNext([''], [''], TRUE);
  end;

  {Création de la TOB Finale - cumul TAUXGLOBAL et MTGLOBAL
   par salarié, date de début de période, base, codification}
  TCot.Detail.Sort(StOrder3);
//TCot.SaveToFile('c:\tmp\TCot', False, TRUE, TRUE);

  TCot2 := TOB.Create('les codif', nil, -1);

  if (TCot.Detail.Count <> 0) then
    TDetCod := TCot.Detail[0];
  if (TDetCod <> nil) then
  begin
    WSalEnCours := TDetCod.GetValue('PHB_SALARIE');
    WRubEnCours := '';  // PT87 FQ15199

    if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
      WInstEnCours := TDetCod.GetValue('POG_INSTITUTION');

    WDateEnCours := TDetCod.GetValue('PHB_DATEDEBUT');
    WDateFinEnCours := TDetCod.GetValue('PHB_DATEFIN');
// d PT63
    if (TDetCod.GetValue('ETB_REGIMEALSACE')='X') then
      WCodEnCours := TDetCod.GetValue('PDF_CODIFALSACE')
    else
      WCodEnCours := TDetCod.GetValue('PDF_CODIFICATION');
// f PT63
    WBaseEnCours := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);
    WTxATEnCours := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
    WOrgEnCours := TDetCod.GetValue('PHB_ORGANISME');

    if (RupSiret = True) then
      if (TDetCod.GetValue('ET_SIRET') <> Null) then
        WsiretEnCours := TDetCod.GetValue('ET_SIRET');
    if (RupNumInt = True) then
      if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
        WNoIntEnCours := TDetCod.GetValue('POG_NUMINTERNE');
    if (RupApe = True) then
      if (TDetCod.GetValue('ET_APE') <> Null) then
        WApeEnCours := TDetCod.GetValue('ET_APE');
    if (RupGpInt = True) then
      if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
        WGpEnCours := TDetCod.GetValue('PGI_GROUPE');

    TxGlobal := 0.00;
    MtGlobal := 0.00;

    WBaseMaj := 0.0;  // PT87
    RegulNegative := False;   // PT96

    I := 0;
    IndCotRegul := false; // PT89 FQ 15222
    while (I < TCot.Detail.Count) do
    begin
      TDetCod := TCot.Detail[I];
// d  PT89 FQ 15222
      if (TDetCod.GetValue('PHB_COTREGUL') =  'REG')  then
        IndCotRegul := true;
// f  PT89 FQ 15222
// d PT63
    if (TDetCod.GetValue('ETB_REGIMEALSACE') =  'X') then
      AlsaceMoselle := true
    else
      AlsaceMoselle := false;
// f PT63
      
      I := I + 1;
      WSalLu := TDetCod.GetValue('PHB_SALARIE');
      WRubLu := TDetCod.GetValue('PHB_RUBRIQUE');  // PT87 FQ15199

      if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
        WInstLu := TDetCod.GetValue('POG_INSTITUTION');

      WDateLu := TDetCod.GetValue('PHB_DATEDEBUT');
      WDateLuFin := TDetCod.GetValue('PHB_DATEFIN');
// d PT63
      if (AlsaceMoselle) then
        WCodLu := TDetCod.GetValue('PDF_CODIFALSACE')
      else
        WCodLu := TDetCod.GetValue('PDF_CODIFICATION');
// f PT63
      WBaseLu := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);

      if (RupSiret = True) then
        if (TDetCod.GetValue('ET_SIRET') <> Null) then
          WsiretLu := TDetCod.GetValue('ET_SIRET');
      if (RupNumInt = True) then
        if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
          WNoIntLu := TDetCod.GetValue('POG_NUMINTERNE');
      if (RupApe = True) then
        if (TDetCod.GetValue('ET_APE') <> Null) then
          WApeLu := TDetCod.GetValue('ET_APE');
      if (RupGpInt = True) then
        if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
          WGpLu := TDetCod.GetValue('PGI_GROUPE');

      if (((RupSiret = True) and (WsiretLu = WsiretEnCours)) or
        (RupSiret = False)) and
        (((RupApe = True) and (WApeLu = WApeEnCours)) or
        (RupApe = False)) and
        (((RupNumInt = True) and (WNoIntLu = WNoIntEnCours)) or
        (RupNumInt = False)) and
        (((RupGpInt = True) and (WGpLu = WGpEnCours)) or
        (RupGpInt = False)) and
        (WSalLu = WSalEnCours) and
        (WInstLu = WInstEnCours) and
// d PT87  FQ 15199
//       (WdateLu = WDateEnCours) and
//        (WdateLuFin = WDateFinEnCours) and
        (WCodLu = WCodEnCours) and
//         (WBaseLu = WBaseEnCours) then
        ((WBaseLu = WBaseEnCours) or (WRubLu = WRubEnCours)) then
      begin
        if (WRubLu <> WRubEnCours) then
        begin
          TxGlobal := TxGlobal + Arrondi(TDetCod.GetValue('TAUXGLOBAL'), 5);
          WBaseMaj := 0.00;
        end;
        MtGlobal := MtGlobal + Arrondi(TDetCod.GetValue('MTGLOBAL'), 5);

        if (WRubLu = WRubEnCours) then
        begin
        if (WBaseMaj = 0.00) then
            WBaseMaj:=  WBaseEnCours;
          WBaseMaj :=  WBaseMaj  + WBaseLu;
// d PT96
        if (WbaseEnCours = WbaseLu *-1) then
          RegulNegative := True;
// f PT96
        end;

        WRubEnCours := TDetCod.GetValue('PHB_RUBRIQUE');
// f PT87 FQ 15199

      end
      else
      begin
        WRubEnCours := '';  // PT87 FQ15199

        I := I - 1;
        TCot2Fille := TOB.Create('', TCot2, -1);
        TCot2Fille.AddChampSupValeur('SALARIE', '', FALSE);
        TCot2Fille.AddChampSupValeur('CODIFICATION', '', FALSE);
        TCot2Fille.AddChampSupValeur('BASECOT', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('TAUXAT', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('DATEDEBUT', iDate1900, FALSE);
        TCot2Fille.AddChampSupValeur('DATEFIN', iDate1900, FALSE);
        TCot2Fille.AddChampSupValeur('ORGANISME', '', FALSE);
        TCot2Fille.AddChampSupValeur('INSTITUTION', '', FALSE);
        TCot2Fille.AddChampSupValeur('TAUXGLOBAL', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('MTGLOBAL', 0.00, FALSE);
        TCot2Fille.AddChampSupValeur('SIRET', '', FALSE);
        TCot2Fille.AddChampSupValeur('APE', '', FALSE);
        TCot2Fille.AddChampSupValeur('NUMINTERNE', '', FALSE);
        TCot2Fille.AddChampSupValeur('GROUPE', '', FALSE); 
//PT78        TCot2Fille.AddChampSupValeur('INSER', 0, FALSE);   //PT63
        TCot2Fille.AddChampSupValeur('REGIMEALSACE', '-', FALSE);   //PT63

        if (TDetCod.GetValue('POG_POSTOTAL') = 0) and
          (TDetCod.GetValue('POG_LONGTOTAL') = 0) then
        { ducs standardisée - on a besoin du n° institution}
        // TCot2Fille.AddChampSup('POG_INSTITUTION',FALSE)
        else
          {ducs personnalisée}
          Personnalisee := True;

        TCot2Fille.PutValue('SALARIE', WSalEnCours);
        TCot2Fille.PutValue('CODIFICATION', WCodEnCours);
        TCot2Fille.PutValue('BASECOT', Arrondi(WBaseEnCours, 5));
// D PT87 FQ 15199
        if (WBaseMaj <> 0.0) or (RegulNegative)  then // PT96
        begin
          TCot2Fille.PutValue('BASECOT', Arrondi(WBaseMaj, 5));
          WBaseMaj := 0.0;
        end;
// F PT87 FQ 15199
        TCot2Fille.PutValue('DATEDEBUT', WDateEnCours);
// PT87 TCot2Fille.PutValue('DATEFIN', WDateFinEnCours);    on force à la date de fin de période pour traiter les bull. compl.
        TCot2Fille.PutValue('DATEFIN', StrToDATe(FinPer));
        if (Copy(WCodEnCours, 1, 6) <> WcodPrec) then
        begin
          WCodPrec := Copy(WCodEnCours, 1, 6);
// d PT63
          if (((AlsaceMoselle) and
               (ExisteSQL('SELECT PDP_CODIFALSACE FROM DUCSPARAM ' +
                          'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                          'PDP_CODIFALSACE LIKE "' +
                          Copy(WCodEnCours, 1, 6) + '%"'))) or
              ((not AlsaceMoselle) and
               (ExisteSQL('SELECT PDP_CODIFICATION FROM DUCSPARAM ' +
                          'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                          'PDP_CODIFICATION LIKE "' +
                           Copy(WCodEnCours, 1, 6) + '%"')))) then
// f PT63
          begin
            TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5));
            WPrecExiste := True;
          end
          else
          begin
            TCot2Fille.PutValue('TAUXAT', 0.00);
            WPrecExiste := False
          end;
        end
        else
        begin
          if (WPrecExiste = True) then
            TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5))
          else
            TCot2Fille.PutValue('TAUXAT', 0.00)
        end;
        TCot2Fille.PutValue('ORGANISME', WOrgEnCours);

        {ducs standardisée - on a besoin du n° institution}
        TCot2Fille.PutValue('INSTITUTION', WInstEnCours);

        TCot2Fille.PutValue('TAUXGLOBAL', Arrondi(TxGlobal, 5));
        TCot2Fille.PutValue('MTGLOBAL', Arrondi(MtGlobal, 5));
        TCot2Fille.PutValue('REGIMEALSACE',TDetCod.GetValue('ETB_REGIMEALSACE'));   //PT63

        if (RupSiret = True) then TCot2Fille.PutValue('SIRET', WsiretEnCours);
        if (RupApe = True) then TCot2Fille.PutValue('APE', WApeEnCours);
        if (RupNumInt = True) then TCot2Fille.PutValue('NUMINTERNE', WNoIntEnCours);
        if (RupGpInt = True) then TCot2Fille.PutValue('GROUPE', WGpEnCours);

        WSalEnCours := TDetCod.GetValue('PHB_SALARIE');

        if (RupGpInt = True) or
          ((Regroupt <> '') and (RupGpInt = False)) then
          if (TDetCod.GetValue('POG_INSTITUTION') <> Null) then
            WInstEnCours := TDetCod.GetValue('POG_INSTITUTION');
        WDateEnCours := TDetCod.GetValue('PHB_DATEDEBUT');
        WDateFinEnCours := TDetCod.GetValue('PHB_DATEFIN');
// d PT63
        if (alsaceMoselle) then
          WCodEnCours := TDetCod.GetValue('PDF_CODIFALSACE')
        else
          WCodEnCours := TDetCod.GetValue('PDF_CODIFICATION');
// f PT63
        WBaseEnCours := Arrondi(TDetCod.GetValue('PHB_BASECOT'), 5);
        WTxATEnCours := Arrondi(TDetCod.GetValue('PHB_TAUXAT'), 5);
        WOrgEnCours := TDetCod.GetValue('PHB_ORGANISME');
        RegulNegative := False; // PT96
        
        if (RupSiret = True) then
          if (TDetCod.GetValue('ET_SIRET') <> Null) then
            WsiretEnCours := TDetCod.GetValue('ET_SIRET');
        if (RupNumInt = True) then
          if (TDetCod.GetValue('POG_NUMINTERNE') <> Null) then
            WNoIntEnCours := TDetCod.GetValue('POG_NUMINTERNE');
        if (RupApe = True) then
          if (TDetCod.GetValue('ET_APE') <> Null) then
            WApeEnCours := TDetCod.GetValue('ET_APE');
        if (RupGpInt = True) then
          if (TDetCod.GetValue('PGI_GROUPE') <> Null) then
            WGpEnCours := TDetCod.GetValue('PGI_GROUPE');
        TxGlobal := 0.00;
        MtGlobal := 0.00;
      end;
    end;

    { dernier élément}
    TCot2Fille := TOB.Create('', TCot2, -1);
    TCot2Fille.AddChampSupValeur('SALARIE', '', FALSE);
    TCot2Fille.AddChampSupValeur('CODIFICATION', '', FALSE);
    TCot2Fille.AddChampSupValeur('BASECOT', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('TAUXAT', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('DATEDEBUT', iDate1900, FALSE);
    TCot2Fille.AddChampSupValeur('DATEFIN', iDate1900, FALSE);
    TCot2Fille.AddChampSupValeur('ORGANISME', '', FALSE);
    TCot2Fille.AddChampSupValeur('INSTITUTION', '', FALSE);
    TCot2Fille.AddChampSupValeur('TAUXGLOBAL', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('MTGLOBAL', 0.00, FALSE);
    TCot2Fille.AddChampSupValeur('SIRET', '', FALSE);
    TCot2Fille.AddChampSupValeur('APE', '', FALSE);
    TCot2Fille.AddChampSupValeur('NUMINTERNE', '', FALSE);
    TCot2Fille.AddChampSupValeur('GROUPE', '', FALSE);
//PT78    TCot2Fille.AddChampSupValeur('INSER', 0, FALSE);   //PT63
    TCot2Fille.AddChampSupValeur('REGIMEALSACE', '-', FALSE);   //PT63

    TCot2Fille.PutValue('SALARIE', WSalEnCours);
    TCot2Fille.PutValue('CODIFICATION', WCodEnCours);
    TCot2Fille.PutValue('BASECOT', Arrondi(WBaseEnCours, 5));
    TCot2Fille.PutValue('DATEDEBUT', WDateEnCours);
// D PT87 FQ 15199
    if (WBaseMaj <> 0.0) and (not RegulNegative)  then // PT96
    begin
      TCot2Fille.PutValue('BASECOT', Arrondi(WBaseMaj, 5));
      WBaseMaj := 0.0;
    end;
// F PT87 FQ 15199
//PT87  TCot2Fille.PutValue('DATEFIN', WDateFinEnCours);    on force à la date de fin de période pour traiter les bull. compl.
        TCot2Fille.PutValue('DATEFIN', StrToDATe(FinPer));

    if (Copy(WCodEnCours, 1, 6) <> WcodPrec) then
    begin
      WCodPrec := Copy(WCodEnCours, 1, 6);
// d PT63
      if (((AlsaceMoselle) and
           (ExisteSQL('SELECT PDP_CODIFALSACE FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFALSACE LIKE "' +
                      Copy(WCodEnCours, 1, 6) + '%"'))) or
          ((not AlsaceMoselle) and
           (ExisteSQL('SELECT PDP_CODIFICATION FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFICATION LIKE "' +
                      Copy(WCodEnCours, 1, 6) + '%"')))) then
// f PT63
      begin
        TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5));
      end
      else
      begin
        TCot2Fille.PutValue('TAUXAT', 0.00);
      end;
    end
    else
    begin
      if (WPrecExiste = True) then
        TCot2Fille.PutValue('TAUXAT', Arrondi(WTxATEnCours, 5))
      else
        TCot2Fille.PutValue('TAUXAT', 0.00);
    end;

    TCot2Fille.PutValue('ORGANISME', WOrgEnCours);

//PT66    if (personnalisee = false) then
      TCot2Fille.PutValue('INSTITUTION', WInstEnCours);

    TCot2Fille.PutValue('TAUXGLOBAL', Arrondi(TxGlobal, 5));
    TCot2Fille.PutValue('MTGLOBAL', Arrondi(MtGlobal, 5));
    TCot2Fille.PutValue('REGIMEALSACE',TDetCod.GetValue('ETB_REGIMEALSACE'));   //PT63

    if (RupSiret = True) then TCot2Fille.PutValue('SIRET', WsiretEnCours);
    if (RupApe = True) then TCot2Fille.PutValue('APE', WApeEnCours);
    if (RupNumInt = True) then TCot2Fille.PutValue('NUMINTERNE', WNoIntLu);
    if (RupGpInt = True) then TCot2Fille.PutValue('GROUPE', WGpEnCours);

  end;

  TCot2.Detail.Sort(StOrder4);
//TCot2.SaveToFile('c:\tmp\TCot2', False, TRUE, TRUE);

{* d PT62*}
  IndicD := 0;
  IndicP := 0;
  TauxGlobalEnCours := 0;
// d PT98
  SalEnCours := '';
  BaseEnCours := 0;
  WMtGlobalEnCours := 0;
  WCodifEnCours := '';
  TauxAtEnCours := 0;

  TCotA := TOB.Create ('Les codifications type AT', nil, -1);
  TCot2bis := TOB.Create ('Les codifications non AT', nil, -1);
  // on éclate TCot2 en 2 TOB :
  //    - TCotA, tob des rubrique de type "présentation par taux AT"
  //    - TCot2Bis, tob des autres rubriques

  for i := 0 to TCot2.Detail.Count-1 do
  begin
    TCot2Fille := TCot2.Detail[i];
    if (ARRONDI(TCot2Fille.GetValue('TAUXAT'),5) <> 0) then
    // si Taux AT renseigné - TOB TCotA des rubriques de type "présentation par taux AT"
    begin
      Baselue := ARRONDI(TCot2Fille.GetValue('BASECOT'),5);
      CodifW := TCot2Fille.GetValue('CODIFICATION');
      SalW :=  TCot2Fille.GetValue('SALARIE');
      DateDebW := TCot2Fille.GetValue('DATEDEBUT');
      DateFinW := TCot2Fille.GetValue('DATEFIN');
      TauxAtW :=  TCot2Fille.GetValue('TAUXAT');
      WMtGlobalLu   := ARRONDI(TCot2Fille.GetValue('MTGLOBAL'),5);
      TauxGlobalLu := ARRONDI(TCot2Fille.GetValue('TAUXGLOBAL'),5);

      if ((SalW = SalEnCours) or (SalEnCours = '')) and
         ((CodifW = WCodifEnCours) or (WCodifEnCours = ''))then
      begin
      // Pour un salarié, une codification --> Cumul des Bases à Taux identiques
        if (TauxGlobalLu = TauxGlobalEnCours) or (TauxGlobalEnCours = 0) then
        begin
          BaseEnCours := BaseEnCours + Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalEnCours + WMtGlobalLu;
          WCodifEnCours := CodifW;
          SalEnCours := SalW;
          TauxAtEnCours := TauxAtW;
        end
        else
        begin
          if (SalEnCours <> '') and  (WCodifEnCours <> '') and (TauxAtEnCours <> 0) then
          // Test pour traiter le cas où l'élément précédent a un taux à 0
          // (c.a.d mis en Tob TCot2Bis) et que cependant sa codification est
          // de type "présentation par taux AT"
          begin
            // Changement de taux : mise en TOB de l'élément en cours
            TCotAFille := TOB.Create ('',TCotA,-1);
            TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
            TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
            TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
            TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
            TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
            TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
            TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
            TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
            TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
            TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
            TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
            TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
            TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
            TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
            TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));
          end;

          BaseEnCours :=  Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
          TauxAtEnCours := TauxAtW;
        end;
      end
      else
      begin
        // Changement de salarié ou de codification : mise en TOB de l'élément en cours
        TCotAFille := TOB.Create ('',TCotA,-1);
        TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
        TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
        TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
        TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
        TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
        TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
        TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
        TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));

        BaseEnCours :=  Baselue;
        TauxGlobalEnCours := TauxGlobalLu;
        SalEnCours := SalW;
        WMtGlobalEnCours := WMtGlobalLu;
        WCodifEnCours := CodifW;
        TauxAtEnCours := TauxAtW;
      end;
    end
    else
    // taux AT = 0 - TOB TCot2Bis des autres rubriques
    begin
      if (WCodifEnCours <> '') and (  SalEnCours <>'') then
      // TOB TCotA : mise en Tob du dernier élément (élt en cours)
      begin
        TCotAFille := TOB.Create ('',TCotA,-1);
        TCotAFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotAFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotAFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotAFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotAFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotAFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
        TCotAFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
        TCotAFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
        TCotAFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotAFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotAFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
        TCotAFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
        TCotAFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
        TCotAFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
        TCotAFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE')); //PT63

        SalEnCours := '';
        BaseEnCours := 0;
        WMtGlobalEnCours := 0;
        WCodifEnCours := '';
        TauxAtEnCours := 0;

      end;
      TCot2BisFille := TOB.Create ('',TCot2Bis,-1);
      TCot2BisFille.AddChampSupValeur ('SALARIE',TCot2Fille.GetValue('SALARIE'));
      TCot2BisFille.AddChampSupValeur ('CODIFICATION',TCot2Fille.GetValue('CODIFICATION'));
      TCot2BisFille.AddChampSupValeur ('BASECOT',TCot2Fille.GetValue('BASECOT'));
      TCot2BisFille.AddChampSupValeur ('TAUXAT',TCot2Fille.GetValue('TAUXAT'));
      TCot2BisFille.AddChampSupValeur ('DATEDEBUT',TCot2Fille.GetValue('DATEDEBUT'));
      TCot2BisFille.AddChampSupValeur ('DATEFIN',TCot2Fille.GetValue('DATEFIN'));
      TCot2BisFille.AddChampSupValeur ('ORGANISME',TCot2Fille.GetValue('ORGANISME'));
      TCot2BisFille.AddChampSupValeur ('INSTITUTION',TCot2Fille.GetValue('INSTITUTION'));
      TCot2BisFille.AddChampSupValeur ('TAUXGLOBAL',TCot2Fille.GetValue('TAUXGLOBAL'));
      TCot2BisFille.AddChampSupValeur ('MTGLOBAL',TCot2Fille.GetValue('MTGLOBAL'));
      TCot2BisFille.AddChampSupValeur ('SIRET',TCot2Fille.GetValue('SIRET'));
      TCot2BisFille.AddChampSupValeur ('APE',TCot2Fille.GetValue('APE'));
      TCot2BisFille.AddChampSupValeur ('NUMINTERNE',TCot2Fille.GetValue('NUMINTERNE'));
      TCot2BisFille.AddChampSupValeur ('GROUPE',TCot2Fille.GetValue('GROUPE'));
      TCot2BisFille.AddChampSupValeur ('REGIMEALSACE',TCot2Fille.GetValue('REGIMEALSACE'));
    end;
  end;
//TCotA.SaveToFile('c:\tmp\TCotA', False, TRUE, TRUE);
  if (TCotA.Detail.count <> 0) then
  begin
    TCotA.Detail.Sort('TAUXAT;SALARIE;CODIFICATION;TAUXGLOBAL;BASECOT;DATEDEBUT;');

    NbCodif := 0;

    TCotABis := TOB.Create ('Les codifications type AT', nil, -1);
    TauxGlobalEnCours := 0;
    SalEnCours := '';
    BaseEnCours := 0;
    TauxAtW := 0;
    WMtGlobalLu := 0;
    for i := 0 to TCotA.Detail.Count-1 do
    begin
      TCotAFille := TCotA.Detail[i];
      Baselue := ARRONDI(TCotAFille.GetValue('BASECOT'),5);
      TauxGlobalLu := ARRONDI(TCotAFille.GetValue('TAUXGLOBAL'),5);
      SalW := TCotAFille.GetValue('SALARIE');
      TauxAtW :=  TCotAFille.GetValue('TAUXAT');
      WMtGlobalLu := ARRONDI(TCotAFille.GetValue('MTGLOBAL'),5);
      CodifW := TCotAFille.GetValue('CODIFICATION');

      if ((SalW = SalEnCours) or (SalEnCours = '')) and ((CodifW = WCodifEnCours) or (WCodifEnCours = '')) then
      // Pour un salarié, une codification --> Cumul des Taux à Bases identiques
      begin
        if (Baselue = BaseEnCours) or (BaseEnCours = 0) then
        begin
          BaseEnCours := Baselue;
          TauxGlobalEnCours := TauxGlobalEnCours + TauxGlobalLu;
          SalEnCours := SalW;
          TauxAtEnCours := TauxAtW;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
        end
        else
        begin
          // Changement de base : mise en TOB de l'élément en cours
          TCotABisFille := TOB.Create ('',TCotABis,-1);
          Nbcodif := NbCodif+1;
          TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
          TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
          TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
          TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
          TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
          TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
          TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
          TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
          TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
          TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
          TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
          TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
          TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
          TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
          TCotABisFille.AddChampSupValeur ('INDIC',0);
          TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

          BaseEnCours :=  Baselue;
          TauxGlobalEnCours := TauxGlobalLu;
          WMtGlobalEnCours := WMtGlobalLu;
          WCodifEnCours := CodifW;
          TauxAtEnCours := TauxAtW;
         end;
      end
      else
      begin
        // Changement de salarié ou de codification : mise en TOB de l'élément en cours
        TCotABisFille := TOB.Create ('',TCotABis,-1);
        Nbcodif := NbCodif+1;
        TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
        TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
        TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
        TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
        TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
        TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
        TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
        TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
        TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
        TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
        TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
        TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
        TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
        TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
        TCotABisFille.AddChampSupValeur ('INDIC',0);
        TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

        BaseEnCours :=  Baselue;
        TauxGlobalEnCours := TauxGlobalLu;
        SalEnCours  := SalW;
        WMtGlobalEnCours := WMtGlobalLu;
        WCodifEnCours := CodifW;
        TauxAtEnCours := TauxAtW;
      end;
    end;
    // TOB TCotABis : mise en Tob du dernier élément (élt en cours)
    TCotABisFille := TOB.Create ('',TCotABis,-1);
    Nbcodif := NbCodif+1;
    TCotABisFille.AddChampSupValeur ('SALARIE',SalEnCours);
    TCotABisFille.AddChampSupValeur ('CODIFICATION',WCodifEnCours);
    TCotABisFille.AddChampSupValeur ('BASECOT',BaseEnCours);
    TCotABisFille.AddChampSupValeur ('TAUXAT',TauxAtEnCours);
    TCotABisFille.AddChampSupValeur ('DATEDEBUT',DateDebW);
    TCotABisFille.AddChampSupValeur ('DATEFIN',TCotAFille.GetValue('DATEFIN'));
    TCotABisFille.AddChampSupValeur ('ORGANISME',TCotAFille.GetValue('ORGANISME'));
    TCotABisFille.AddChampSupValeur ('INSTITUTION',TCotAFille.GetValue('INSTITUTION'));
    TCotABisFille.AddChampSupValeur ('TAUXGLOBAL',TauxGlobalEnCours);
    TCotABisFille.AddChampSupValeur ('MTGLOBAL',WMtGlobalEnCours);
    TCotABisFille.AddChampSupValeur ('SIRET',TCotAFille.GetValue('SIRET'));
    TCotABisFille.AddChampSupValeur ('APE',TCotAFille.GetValue('APE'));
    TCotABisFille.AddChampSupValeur ('NUMINTERNE',TCotAFille.GetValue('NUMINTERNE'));
    TCotABisFille.AddChampSupValeur ('GROUPE',TCotAFille.GetValue('GROUPE'));
    TCotABisFille.AddChampSupValeur ('INDIC',0);
    TCotABisFille.AddChampSupValeur ('REGIMEALSACE',TCotAFille.GetValue('REGIMEALSACE'));

    TCotABis.Detail.Sort('CODIFICATION;TAUXAT;TAUXGLOBAL;BASECOT;DATEDEBUT;');

    INDIC := 0;
    TauxAtEnCours := 0;
    TauxGlobalEnCours := 0;
    WCodifEnCours := '';

    for i := 0 to TCotABis.Detail.Count-1 do
    begin
      TCotABisFille := TCotABis.Detail[i];
      TauxAtW := TCotABisFille.GetValue('TAUXAT');
      TauxGlobalLu := Arrondi(TCotABisFille.GetValue('TAUXGLOBAL'),5);
      CodifW := TCotABisFille.GetValue('CODIFICATION');
      SalW := TCotABisFille.GetValue('SALARIE');
      if (copy(CodifW,7,1) = 'D') then
      begin
        if ((CodifW = WCodifEnCours) or (WCodifEnCours= '')) and
           ((TauxGlobalLu = TauxGlobalEnCours) or (TauxGlobalEnCours=0))and
           ((TauxAtW = TauxAtEnCours) or (TauxAtEnCours=0)) then
        begin
          TCotABisFille.PutValue ('INDIC',INDIC);
        end
        else
        begin
          INDIC := INDIC+1;
          TCotABisFille.PutValue ('INDIC',INDIC);
        end;
        TauxAtEnCours := TauxAtW;
        TauxGlobalEnCours := TauxGlobalLu;
        WCodifEnCours := CodifW;
      end;
    end;

    for i := 0 to TCotABis.Detail.Count-1 do
    begin
      TCotABisFille := TCotABis.Detail[i];
      TauxAtW := TCotABisFille.GetValue('TAUXAT');
      SalW := TCotABisFille.GetValue('SALARIE');
      CodifW := TCotABisFille.GetValue('CODIFICATION');
      if (copy(CodifW,7,1) = 'D') then
      begin
        TCotABisFilleP :=  TCotABis.FindFirst(['CODIFICATION','SALARIE','TAUXAT'],
                                         [copy(CodifW,1,6)+'P', SalW, TauxAtW], TRUE);
        while (TCotABisFilleP <> nil)  do
        begin
          TCotABisFilleP.AddChampSupValeur('INDIC',TCotABisFille.GetValue ('INDIC'));

          TCotABisFilleP :=  TCotABis.FindNext(['CODIFICATION','SALARIE','TAUXAT'],
                             [copy(CodifW,1,6)+'P', SalW, TauxAtW], TRUE);
        end
      end;
    end;
  end; // fin TCotA.Detail.Count <> 0

  // Création de la TOB TCot3 (détail des rubriques par codification et salarié)
  TCot3 := TOB.Create('les codifs à éditer', nil, -1);
  for i := 0 to NbCodif-1 do
  // on récupère d'abord les rubriquee de type "présentation par taux AT" (TCotA)
  begin
    TCotABisFille := TCotABis.Detail[i];

    TCot3Fille := TOB.Create ('',TCot3,-1);

    TCot3Fille.AddChampSupValeur ('SALARIE',TCotABisFille.GetValue('SALARIE'));
    TCot3Fille.AddChampSupValeur ('CODIFICATION',TCotABisFille.GetValue('CODIFICATION'));
    TCot3Fille.AddChampSupValeur ('BASECOT',TCotABisFille.GetValue('BASECOT'));
    TCot3Fille.AddChampSupValeur ('TAUXAT',TCotABisFille.GetValue('TAUXAT'));
    TCot3Fille.AddChampSupValeur ('DATEDEBUT',TCotABisFille.GetValue('DATEDEBUT'));
    TCot3Fille.AddChampSupValeur ('DATEFIN',TCotABisFille.GetValue('DATEFIN'));
    TCot3Fille.AddChampSupValeur ('ORGANISME',TCotABisFille.GetValue('ORGANISME'));
    TCot3Fille.AddChampSupValeur ('INSTITUTION',TCotABisFille.GetValue('INSTITUTION'));
    TCot3Fille.AddChampSupValeur ('TAUXGLOBAL',TCotABisFille.GetValue('TAUXGLOBAL'));
    TCot3Fille.AddChampSupValeur ('MTGLOBAL',TCotABisFille.GetValue('MTGLOBAL'));
    TCot3Fille.AddChampSupValeur ('SIRET',TCotABisFille.GetValue('SIRET'));
    TCot3Fille.AddChampSupValeur ('APE',TCotABisFille.GetValue('APE'));
    TCot3Fille.AddChampSupValeur ('NUMINTERNE',TCotABisFille.GetValue('NUMINTERNE'));
    TCot3Fille.AddChampSupValeur ('GROUPE',TCotABisFille.GetValue('GROUPE'));
    TCot3Fille.AddChampSupValeur ('INDIC',TCotABisFille.GetValue('INDIC'));
    TCot3Fille.AddChampSupValeur ('REGIMEALSACE',TCotABisFille.GetValue('REGIMEALSACE'));

  end;

  for i:= 0 to TCot2Bis.Detail.Count-1 do
  // on récupère ensuite les rubriquee de type "cas normal" (TCot2Bis)
  begin
      TCot2BisFille := TCot2Bis.Detail[i];

    TCot3Fille := TOB.Create ('',TCot3,-1);
    TCot3Fille.AddChampSupValeur ('SALARIE',TCot2BisFille.GetValue('SALARIE'));
    TCot3Fille.AddChampSupValeur ('CODIFICATION',TCot2BisFille.GetValue('CODIFICATION'));
    TCot3Fille.AddChampSupValeur ('BASECOT',TCot2BisFille.GetValue('BASECOT'));
    TCot3Fille.AddChampSupValeur ('TAUXAT',TCot2BisFille.GetValue('TAUXAT'));
    TCot3Fille.AddChampSupValeur ('DATEDEBUT',TCot2BisFille.GetValue('DATEDEBUT'));
    TCot3Fille.AddChampSupValeur ('DATEFIN',TCot2BisFille.GetValue('DATEFIN'));
    TCot3Fille.AddChampSupValeur ('ORGANISME',TCot2BisFille.GetValue('ORGANISME'));
    TCot3Fille.AddChampSupValeur ('INSTITUTION',TCot2BisFille.GetValue('INSTITUTION'));
    TCot3Fille.AddChampSupValeur ('TAUXGLOBAL',TCot2BisFille.GetValue('TAUXGLOBAL'));
    TCot3Fille.AddChampSupValeur ('MTGLOBAL',TCot2BisFille.GetValue('MTGLOBAL'));
    TCot3Fille.AddChampSupValeur ('SIRET',TCot2BisFille.GetValue('SIRET'));
    TCot3Fille.AddChampSupValeur ('APE',TCot2BisFille.GetValue('APE'));
    TCot3Fille.AddChampSupValeur ('NUMINTERNE',TCot2BisFille.GetValue('NUMINTERNE'));
    TCot3Fille.AddChampSupValeur ('GROUPE',TCot2BisFille.GetValue('GROUPE'));
    TCot3Fille.AddChampSupValeur ('INDIC',999);
    TCot3Fille.AddChampSupValeur ('REGIMEALSACE',TCot2BisFille.GetValue('REGIMEALSACE'));

// Le TAUX AT est à zéro : exite t'il dans TCotA, tob des rubriques de
// type "présentation par taux AT", une codification identique.
// Correspond au cas où le salarié n'a de cotisation AT.

    TCotABisFille := TCotA.FindFirst (['CODIFICATION','DATEDEBUT','DATEFIN'],
                                       [TCot2BisFille.GetValue('CODIFICATION'),
                                        TCot2BisFille.GetValue('DATEDEBUT'),
                                        TCot2BisFille.GetValue('DATEFIN')], TRUE);
    if TCotABisFille <> nil then
    begin
        Tcot3FilleBis :=  TCot3.FindFirst (['CODIFICATION','DATEDEBUT','DATEFIN'],
                                       [TCot2BisFille.GetValue('CODIFICATION'),
                                        TCot2BisFille.GetValue('DATEDEBUT'),
                                        TCot2BisFille.GetValue('DATEFIN')], TRUE);
        if Tcot3FilleBis <> nil then
          TCot3Fille.AddChampSupValeur ('INDIC',Tcot3FilleBis.GetValue('INDIC'));
    end;

   end;
// f PT98

// d PT63
    StOrder := '';
    if (RupSiret = True)  then  StOrder := StOrder + 'SIRET;';
    if (RupApe = True)    then  StOrder := StOrder + 'APE;';
    if (RupNumInt = True) then  StOrder := StOrder + 'NUMINTERNE;';
    if (RupGpInt = True)  then  StOrder := StOrder + 'GROUPE;';

//PT81    Storder := StOrder+'INDIC;CODIFICATION;TAUXGLOBAL;SALARIE;BASECOT;DATEDEBUT' ;
//PT85    Storder := StOrder+'INDIC;CODIFICATION;TAUXAT;TAUXGLOBAL;SALARIE;BASECOT;DATEDEBUT' ;
//PT88    Storder := StOrder+'CODIFICATION;TAUXAT;TAUXGLOBAL;SALARIE;BASECOT;DATEDEBUT' ;
    Storder := StOrder+'INDIC;CODIFICATION;TAUXAT;TAUXGLOBAL;SALARIE;BASECOT;DATEDEBUT' ;

    TCot3.Detail.Sort(StOrder);
//Tcot3.SaveToFile('c:\tmp\TCot3', False, TRUE, TRUE);

// f PT63
  FreeAndNil(TCotA);
  FreeAndNil(TCot2);
  FreeAndNil(TCot2Bis); // PT72
{* f PT62 *}
  FreeAndNil(TCotABis); // PT98

// d PT89 fq 15222
  if (IndCotRegul) then
  begin
      PGIBox('Etablissement '+ Etab +#13#10+
             ' Attention : des rubriques de régularisation de cotisation .R ont été calculées'+#13#10+
             ' dans les bulletins de la période sélectionnée et peuvent par conséquent entrainer'+#13#10+
             ' une rupture sur les codifications DUCS', Ecran.Caption);
      TraceE.Add('Etablissement '+ Etab +#13#10+
                 ' Des rubriques de régularisation de cotisation .R ont été calculées'+#13#10+
                 ' dans les bulletins de la période sélectionnée et peuvent par conséquent entrainer'+#13#10+
                 ' une rupture sur les codifications DUCS');
  end;
// f PT89 fq 15222

  if TCot <> nil then
  begin
    TCot.Free;
  end;
end; { fin ChargeCot}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 02/08/2001
Modifié le ... : 02/08/2001
Description .. : MajDucs
Suite ........ : Mise à jour des tables DUCSENTETE et DUCSDETAIL
Suite ........ : On Supprime la DUCS (entete et détail) si existe déjà
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.MajDucs(DucsDoss: Boolean; Regroupt: string; TraceE :TStringList); // PT65 + maj jnal événements
var
  TDetail, TOB_Lignes, TOB_FilleTete, TOB_FilleLignes: TOB;
  TAT, TobSTExiste: TOB;
  BaseCodif, BaseLigne, TauxEnCours, ATEnCours, TauxLu, ATLu: double;
  MontForfait, MontCodif, MontST, MontLigne: double;
  Effectif, ligne, i, NoAT, NoATP, NoInstEnCours: Integer;
//PT62   ligne, i, NoAT, NoATP, NoInstEnCours               : Integer;
  PogSTPos, PogSTLong, PosDeb, LongEdit, j, NbJour: Integer;
  PogBoolST, PogTypDecl1, PogTypDecl2, IndRup: Boolean;
  Personnalisee, IndCalcSt: Boolean;
  QDucsParam, QInst, QEtab, QOrg, Qdetail, QQ: TQuery;
  aa, mm, jj: Word;
  DatExigible, DatLimDepot, DatReglt: TDateTime;
  CodifEnCours, CodifLue, SalEnCours, SalLu, StrDelDetail: string;
  OrgLu, OrgEnCours, NoInst, LibInst, StrDelTete: string;
  PogPeriod1, PogPeriod2, PogNature, Cod100, AbregePeriod, StQQ: string;
  Monnaie, LigneOptique, BaseArrCod, MontArrCod, TypeCotis: string;
  CodifAT, Predef, ZEDE, WSiretLu, WNoIntLu, WApeLu, WGpLu: string;
  NumMsa, CodInstLu, CodInstEnCours: string;
  NomChamp: array[1..2] of string;
  ValeurChamp: array[1..2] of variant;
  ValChST: array[1..1] of variant;
  NomChampST: array[1..1] of string;

  Anomalie      : string;

{$IFDEF EAGLCLIENT}
  KK, SortDos: Integer;
  MaTob: TOB;
  MaTobAt: TOB;
{$ENDIF}

  QDucsParamAT: TQuery;

  PogCpInfo, PogCentrePayeur: string;
  TEffM, TEffMFille: TOB;

  CSC : string;
{$IFNDEF DUCS41}
  PogTypBordereau1, PogTypBordereau2, PogDucsDoss : string; //PT70
{$ENDIF}
begin
  Anomalie := '';
  IndRup := False;
  ZEDE := 'ZZZZZZZ';
  PosDeb := 0;
  LongEdit := 0;
  NoInstEnCours := 0;
  PogBoolST := FALSE;
  PogSTPos := 0;
  PogSTLong := 0;
  {Récupération informations établissement}
  PogNumInterne := 'Non Renseigné';
  EtSiret := 'Non Renseigné';
  EtApe := 'Inco';
  QEtab := OpenSql('SELECT ET_SIRET,ET_APE' +
    ' FROM ETABLISS ' +
    ' WHERE ET_ETABLISSEMENT = "' + Etab + '"', True);
  if not QEtab.eof then
  begin
    ForceNumerique(QEtab.FindField('ET_SIRET').AsString, EtSiret);
    EtApe := QEtab.FindField('ET_APE').AsString;
  end;
  Ferme(QEtab);

  {Récupération informations organisme}

  QOrg := OpenSql('SELECT POG_NUMINTERNE, POG_NATUREORG,' +
    'POG_SOUSTOTDUCS, POG_POSTOTAL, POG_LONGTOTAL,' +
    'POG_PERIODICITDUCS, POG_AUTREPERIODUCS,' +
    'POG_PERIODCALCUL, POG_AUTPERCALCUL,' +
    'POG_LGOPTIQUE' +
    ',POG_CPINFO,POG_CENTREPAYEUR' +
{$IFNDEF DUCS41}
                  ',POG_TYPEPERIOD,POG_TYPEAUTPERIOD'+
{$ENDIF}
    ' FROM ORGANISMEPAIE ' +
    ' WHERE POG_ETABLISSEMENT= "' + Etab + '" AND ' +
    ' POG_ORGANISME = "' + Organisme + '"', True);
  if not QOrg.eof then
  begin
    PogNumInterne := copy(QOrg.FindField('POG_NUMINTERNE').AsString, 1, 30);
    PogNature := QOrg.FindField('POG_NATUREORG').AsString;
    PogBoolST := QOrg.FindField('POG_SOUSTOTDUCS').AsString = 'X';
    PogSTPos := QOrg.FindField('POG_POSTOTAL').AsInteger;
    PogSTLong := QOrg.FindField('POG_LONGTOTAL').AsInteger;
    PogPeriod1 := QOrg.FindField('POG_PERIODICITDUCS').AsString;
    PogPeriod2 := QOrg.FindField('POG_AUTREPERIODUCS').AsString;
    PogTypDecl1 := QOrg.FindField('POG_PERIODCALCUL').AsString = 'X';
    PogTypDecl2 := QOrg.FindField('POG_AUTPERCALCUL').AsString = 'X';
    PogLGOptique := QOrg.FindField('POG_LGOPTIQUE').AsString;
    PogCPInfo := QOrg.FindField('POG_CPINFO').AsString;
    PogCentrePayeur := QOrg.FindField('POG_CENTREPAYEUR').AsString;
{$IFNDEF DUCS41}
// d PT70
    PogTypBordereau1 := QOrg.FindField('POG_TYPEPERIOD').AsString;
    PogTypBordereau2 := QOrg.FindField('POG_TYPEAUTPERIOD').AsString;
    if (DucsDoss) then
      PogDucsDoss := 'X'
    else
      PogDucsDoss := '-';

   if (PogTypBordereau1 = '') then
   begin
     try
     begintrans;
      PogTypBordereau1 := AffectTypeBordereau (PogNature, PogPeriod1,
                                         DucsDoss);
      ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_TYPEPERIOD = "'+PogTypBordereau1+'" '+
                 'WHERE POG_NATUREORG = "'+PogNature+'" AND '+
                 'POG_TYPEPERIOD = "" AND '+
                 'POG_PERIODICITDUCS = "'+PogPeriod1+'" AND '+
                 'POG_DUCSDOSSIER ="'+PogDucsDoss+'"');

     Committrans;
     except
     Rollback;
     end;
   end;
   if (PogTypBordereau2 = '') then
   begin
     try
     begintrans;
      PogTypBordereau2 := AffectTypeBordereau (PogNature, PogPeriod2,
                                         DucsDoss);
      ExecuteSQL('UPDATE ORGANISMEPAIE SET POG_TYPEAUTPERIOD = "'+PogTypBordereau2+'" '+
                 'WHERE POG_NATUREORG = "'+PogNature+'" AND '+
                 'POG_TYPEAUTPERIOD ="" AND '+
                 'POG_AUTREPERIODUCS = "'+PogPeriod2+'" AND '+
                 'POG_DUCSDOSSIER ="'+PogDucsDoss+'"');
     Committrans;
     except
     Rollback;
     end;
   end;
// f PT70
// récupération du type de bordereau
   if (PeriodEtat = PogPeriod1) then
//     PogTypBordereau := QOrg.FindField('POG_TYPEPERIOD').AsString
     PogTypBordereau := PogTypBordereau1
   else
     if (PeriodEtat = PogPeriod2) then
//       PogTypBordereau := QOrg.FindField('POG_TYPEAUTPERIOD').AsString
       PogTypBordereau := PogTypBordereau2
     else
       PogTypBordereau := '';
{$ENDIF}
  end;
  Ferme(QOrg);

  {ducs personnalisée}
  personnalisee := False;
  if (PogNature = '300') and
    (PogBoolST = True) and
    (PogSTPos <> 0) and
    (PogSTLong <> 0) then
    personnalisee := True;

  if PogNature = '600' then
  {MSA  : Récup du N° MSA}
  begin
    QEtab := OpenSql('SELECT ETB_NUMMSA' +
      ' FROM ETABCOMPL ' +
      ' WHERE ETB_ETABLISSEMENT = "' + Etab + '"', True);
    if not QEtab.eof then
    begin
      ForceNumerique(QEtab.FindField('ETB_NUMMSA').AsString, NumMsa);
    end;
    EtSiret := copy(NumMsa, 1, 14);
    Ferme(QEtab);
  end;

  {Formatage code période abrégé}
  DecodeDate(StrToDate(FinPer), aa, mm, jj);
  if (aa < 2000) then
    aa := aa - 1900
  else
    aa := aa - 2000;

  if (VH_Paie.PGDecalage = TRUE) or (VH_Paie.PGDecalagePetit = TRUE) then
  {Décalage de paie}
  begin
    if (mm = 12) then aa := aa + 1;
    if (aa = 100) then aa := 0;
    mm := mm + 1;
    if (mm = 13) then mm := 1;
  end;

  AbregePeriod := '0000';

  if (PeriodEtat = 'M') then
    AbregePeriod := FormatFloat('00', aa) +
      FormatFloat('0', ((mm - 1) div 3) + 1) +
      FormatFloat('0', (((mm - 1) - (((mm - 1) div 3) * 3) + 1)));
  if (PeriodEtat = 'T') then
    AbregePeriod := FormatFloat('00', aa) +
      FormatFloat('0', ((mm - 1) div 3) + 1) + '0';
  if (PeriodEtat = 'A') then
    AbregePeriod := FormatFloat('00', aa) + '00';

// d PT56
  { Calcul Date exigibilité }
  DecodeDate(StrToDate(Finper), aa, mm, jj);
{ d PT68
  if ( jj >= JExigible) then
  begin
     DatExigible := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
     DatExigible := PlusDate(DatExigible,JExigible-1,'J');
  end
  else
     DatExigible := PlusDate(DebutdeMois(StrToDate(Finper)),JExigible-1,'J');}

  DatExigible := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');

  if (JExigible = 30) or (JExigible = 31) then
  // fin de mois suivant la date de fin de période
    DatExigible := FinDeMois(DatExigible)
  else
  begin
     DecodeDate (FindeMois(DatExigible), aa, mm, jj);
     if (mm = 2) and ((JExigible = 28) or (JExigible = 29)) then
       DatExigible := FinDeMois(DatExigible)
     else
//PT71
     begin
       if (JExigible > 31) then
       // après la fin de mois suivant
       begin
         DatExigible := FinDeMois(DatExigible) ;
         JExigible := JExigible-31+1;
       end;
       DatExigible := PlusDate(DatExigible,JExigible-1,'J');
     end;
//PT71
  end;

  { Calcul Date limite de dépôt }
{
  if ( jj >= JLimDepot) then
  begin
     DatLimDepot := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
     DatLimDepot := PlusDate(DatLimDepot,JLimDepot-1,'J');
  end
  else
     DatLimDepot := PlusDate(DebutdeMois(StrToDate(Finper)),JLimDepot-1,'J'); }
  DatLimDepot := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');

  if (JLimDepot = 30) or (JLimDepot = 31) then
  // fin de mois suivant la date de fin de période
  DatLimDepot := FinDeMois(DatLimDepot)
  else
  begin
     DecodeDate (FindeMois(DatLimDepot), aa, mm, jj);
     if (mm = 2) and ((JLimDepot = 28) or (JLimDepot = 29)) then
       DatLimDepot := FinDeMois(DatLimDepot)
     else
//PT71
     begin
       if (JLimDepot > 31) then
       // après la fin de mois suivant
       begin
         DatLimDepot := FinDeMois(DatLimDepot) ;
         JLimDepot := JLimDepot-31+1;
       end;
       DatLimDepot := PlusDate(DatLimDepot,JLimDepot-1,'J');
     end;
//PT71
  end;

  {Calcul date de règlement}
{
  if ( jj >= JReglement) then
  begin
     DatReglt := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');
     DatReglt := PlusDate(DatReglt,JReglement-1,'J');
  end
  else
     DatReglt := PlusDate(DebutdeMois(StrToDate(Finper)),JReglement-1,'J');}
  DatReglt := PlusDate(DebutdeMois(StrToDate(Finper)),1,'M');

  // fin de mois suivant la date de fin de période
  if (JReglement = 30) or (JReglement = 31) then
    DatReglt := FinDeMois(DatReglt)
  else
  begin
     DecodeDate (FindeMois(DatReglt), aa, mm, jj);
     if (mm = 2) and ((JReglement = 28) or (JReglement = 29)) then
       DatReglt := FinDeMois(DatReglt)
     else
//PT71
     begin
       if (JReglement > 31) then
       // après la fin de mois suivant
       begin
         DatReglt := FinDeMois(DatReglt) ;
         JReglement := JReglement-31+1;
       end;
       DatReglt := PlusDate(DatReglt,JReglement-1,'J');
     end;
//PT71
  end;
// f PT68


{  DatExigible := IDate1900;
  if (JExigible <> 0) and (DatePaye <> 0) then
  begin
    if (DatePaye = FindeMois(DatePaye)) then}
    {La date de versement des salaires est un fin de mois}
{      DatExigible := PlusDate(DatePaye, JExigible, 'J')
    else
    begin}
    {La date de versement des salaires est en cours de mois}
{      if (JExigible < NbJour) then}
      {Date d'exigibilité dans le mois de versement des salaires}
{        DatExigible := PlusDate(DatePaye, JExigible, 'J')
      else}
      {Date d'exigibilité après le mois de versement des salaires}
{        DatExigible := PlusDate(FindeMois(DatePaye), JExigible - NbJour, 'J');
    end;
  end
  else}
  {cas des ducs néant où la date de versement des salaires
   n'est pas renseignéee
  }
{  begin
    if ((PogNature = '300') or (PogNature = '700')) and
      (JExigible <> 0) and (DatePaye = 0) then
    begin
      DatExigible := StrToDate(Finper);}
      {base de calcul : 1 mois = 31 jours}
{      if (JExigible > 31) then
      begin
        DatExigible := FindeMois(PlusMois(DatExigible, 1));
        JExigible := JExigible - 31;
      end;
      DatExigible := PlusDate(DatExigible, JExigible, 'J');
    end;
  end;}

  { Calcul Date limite de dépôt }
{  DatLimDepot := IDate1900;
  if (JLimDepot <> 0) and (DatePaye <> 0) then
  begin
    if (DatePaye = FindeMois(DatePaye)) then}
      {La date de versement des salaires est un fin de mois}
{      DatLimDepot := PlusDate(DatePaye, JLimDepot, 'J')
    else
    begin}
    {La date de versement des salaires est en cours de mois}
{      if (JLimDepot < NbJour) then}
      {Date limite de dépôt dans le mois de versement des salaires}
{        DatLimDepot := PlusDate(DatePaye, JLimDepot, 'J')
      else}
      {Date limite de dépôt après le mois de versement des salaires}
{        DatLimDepot := PlusDate(FindeMois(DatePaye), JLimDepot - NbJour, 'J')
    end;
  end
  else
  begin}
  { cas des ducs néant où la date de versement des salaires
    n'est pas renseigné
  }
{    if ((PogNature = '300') or (PogNature = '700') or (PogNature = '600')) and
      (JLimDepot <> 0) and (DatePaye = 0) then
    begin
      DatLimDepot := StrToDate(Finper);}
      {base de calcul : 1 mois = 31 jours}
{      if (JLimDepot > 31) then
      begin
        DatLimDepot := FindeMois(PlusMois(DatLimDepot, 1));
        JLimDepot := JLimDepot - 31;
      end;
      DatLimDepot := PlusDate(DatLimDepot, JLimDepot, 'J');
    end;
  end;}

  {Calcul date de règlement}
{  DatReglt := IDate1900;
  if (JReglement <> 0) and (DatePaye <> 0) then
  begin
    if (DatePaye = FindeMois(DatePaye)) then}
    {La date de versement des salaires est un fin de mois}
{      DatReglt := PlusDate(DatePaye, JReglement, 'J')
    else
    begin}
    {La date de versement des salaires est en cours de mois}
{      if (JReglement < NbJour) then}
       {Date limite de dépôt dans le mois de versement des salaires}
{        DatReglt := PlusDate(DatePaye, JReglement, 'J')
      else}
       {Date limite de dépôt après le mois de versement des salaires}
{        DatReglt := PlusDate(FindeMois(DatePaye), JReglement - NbJour, 'J')
    end;
  end
  else}
  {cas des ducs néant où la date de versement des salaires
   n'est pas renseignée
  }
{  begin
    if (JReglement <> 0) and (DatePaye = 0) then
    begin
      DatReglt := StrToDate(Finper);}
      {base de calcul : 1 mois = 31 jours}
{      if (JReglement > 31) then
      begin
        DatReglt := FindeMois(PlusMois(DatReglt, 1));
        JReglement := JReglement - 31;
      end;
      DatReglt := PlusDate(DatReglt, JReglement, 'J');
    end;
  end;}
  // f PT56

  if (DatePaye = 0) then DatePaye := IDate1900;

  {Récup. de la monnaie de tenue}
  if VH_Paie.PGTenueEuro = FALSE then
    Monnaie := 'FRF' // Franc
  else
    Monnaie := 'EUR'; // Euro

  {Formatage ligne optique (elle peut être déjà renseignée au niveau
   du paramètrage de l'organisme)}
  LigneOptique := '';
  if (PogLgOptique = '') then
  begin
    if (PogNature = '100') then
    { URSSAF }
    begin
      LigneOptique := Copy(PogNumInterne, 1, length(PogNumInterne));
      if (length(PogNumInterne) < 30) then
      begin
        for i := length(PogNumInterne) + 1 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end;

    if (PogNature = '200') then
    { ASSEDIC }
    begin
      LigneOptique := 'S2' + Copy(PogNumInterne, 1, length(PogNumInterne));
      if ((length(PogNumInterne) + 2) < 30) then
      begin
        for i := length(PogNumInterne) + 3 to 30 do
        begin
          LigneOptique := LigneOptique + '0';
        end;
      end;
    end;
  end
  else
  begin
    LigneOptique := PogLgOptique;
    if (length(PogLgOptique) < 30) then
    begin
      for i := length(PogLgOptique) + 1 to 30 do
      begin
        LigneOptique := LigneOptique + '0';
      end;
    end;
  end;

  {ALIMENTATION TABLES DUCS}
  BaseCodif := 0.00;
  MontCodif := 0.00;
  MontForfait := 0.00;
  MontST := 0.00;
  IndCalcSt := false; { ducs personnalisée}
  NoAT := -1;

  NumDucs := 0;
  ligne := 0;
  EtabEnCours := Etab;
  OrganismeEnCours := Organisme;
//PT62  TDetail := TCot2.FindFirst([''], [''], TRUE);
  TDetail := TCot3.FindFirst([''], [''], TRUE);
  { Gestion des ruptures}
  if (Rupture = TRUE) then
  begin
    TRup := TRupMere.FindFirst([''], [''], TRUE);
    if (TRup <> nil) then
    begin
      IndRup := False;
      if (TRup.GetValue('ET_SIRET') <> Null) then
        SiretEnCours := TRup.GetValue('ET_SIRET');
      if (TRup.GetValue('ET_APE') <> Null) then
        ApeEnCours := TRup.GetValue('ET_APE');
      EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
      OrganismeEnCours := TRup.GetValue('POG_ORGANISME');
      if (TRup.GetValue('POG_NUMINTERNE') <> Null) then
        NumIntEnCours := copy(TRup.GetValue('POG_NUMINTERNE'), 1, 30);
      if (TRup.GetValue('POG_LGOPTIQUE') <> Null) then
        LGOptiqueEncours := TRup.GetValue('POG_LGOPTIQUE');
      if (RupGpInt = TRUE) then
      begin
        if (TRup.GetValue('PGI_GROUPE') <> Null) then
          GpIntEnCours := TRup.GetValue('PGI_GROUPE');
      end;
    end;
  end; { fin if (Rupture = TRUE)}

  if (Tdetail <> nil) then
  begin
    if (RupSiret = TRUE) then
      if (TDetail.GetValue('SIRET') <> Null) then
        Wsiret := TDetail.GetValue('SIRET');
    if (RupApe = TRUE) then
      if (TDetail.GetValue('APE') <> Null) then
        WApe := TDetail.GetValue('APE');
    if (RupNumInt = TRUE) then
      if (TDetail.GetValue('NUMINTERNE') <> Null) then
        WNumInt := TDetail.GetValue('NUMINTERNE');
    if (RupGpInt = TRUE) then
      if (TDetail.GetValue('GROUPE') <> Null) then
        WGpInt := TDetail.GetValue('GROUPE');
  end; { fin if (Tdetail <> NIl)}

  while ((fctRupture(TDetail, 0) = TRUE) or (IndNeant2 = TRUE)) and
    ((Rupture = False) or ((Rupture = TRUE) and (TRup <> nil))) do
  begin
  {Alimentation DUCSENTETE}
    if (NumDucs = 0) then
    begin
      TOB_Lignes := TOB.Create('La Déclaration', nil, -1);
    end;
    {Recherche s'il n'existe pas déjà une ducs avec la même clé
     dans ce cas on incrémente le n° de ducs pour ne pas rencontrer de
     problème de clé duppliquée. (cas de l'utilisation des ruptures)}
    while ExisteSQL('SELECT PDU_ETABLISSEMENT,PDU_ORGANISME,PDU_DATEDEBUT,' +
      'PDU_DATEFIN,PDU_NUM FROM DUCSENTETE WHERE ' +
      'PDU_ETABLISSEMENT="' + EtabEnCours + '" AND ' +
      'PDU_ORGANISME="' + Organisme + '" AND ' +
// les ruptures sur l'état ne tiennent pas comptes des dates
// en Impression DUCS (menu 42302) certaines Ducs n'étaient pas éditées.
// PT65      'PDU_DATEDEBUT="' + UsDateTime(StrToDate(DebPer)) + '" AND ' +
// PT65      'PDU_DATEFIN="' + UsDateTime(StrToDate(FinPer)) + '" AND ' +
      'PDU_NUM=' + IntToStr(NumDucs) + '') do // DB2
    begin
      NumDucs := NumDucs + 1;
      ligne := 0;
    end;

    // Centre Payeur
    QOrg := OpenSql('SELECT' +
      ' POG_CPINFO,POG_CENTREPAYEUR' +
      ' FROM ORGANISMEPAIE ' +
      ' WHERE POG_ETABLISSEMENT= "' + EtabEnCours + '" AND ' +
      ' POG_ORGANISME = "' + Organisme + '"', True);
    if not QOrg.eof then
    begin
      PogCPInfo := QOrg.FindField('POG_CPINFO').AsString;
      PogCentrePayeur := QOrg.FindField('POG_CENTREPAYEUR').AsString;
    end;
    Ferme(QOrg);

    TOB_FilleTete := TOB.create('DUCSENTETE', TOB_Lignes, -1);
    TOB_FilleTete.PutValue('PDU_ETABLISSEMENT', EtabEnCours);
    TOB_FilleTete.PutValue('PDU_ORGANISME', Organisme);
    TOB_FilleTete.PutValue('PDU_DATEDEBUT', StrToDate(DebPer));
    TOB_FilleTete.PutValue('PDU_DATEFIN', StrToDate(FinPer));
    TOB_FilleTete.PutValue('PDU_NUM', NumDucs);

    if ((Rupture = TRUE) and (SiretEnCours = '')) or
      ((Rupture <> TRUE) and (EtSiret = '')) then
      Anomalie := 'Non Renseigné';

    if (Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_SIRET', SiretEnCours)
    else
      TOB_FilleTete.PutValue('PDU_SIRET', EtSiret);

    if (Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_APE', ApeEnCours)
    else
      TOB_FilleTete.PutValue('PDU_APE', EtApe);

    if (RupGpInt = TRUE) then
      TOB_FilleTete.PutValue('PDU_GROUPE', GpIntEnCours)
    else
      TOB_FilleTete.PutValue('PDU_GROUPE', '');

    if (Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_NUMERO', NumIntEnCours)
    else
      TOB_FilleTete.PutValue('PDU_NUMERO', PogNumInterne);

    TOB_FilleTete.PutValue('PDU_ABREGEPERIODE', AbregePeriod);
    TOB_FilleTete.PutValue('PDU_DATEEXIGIBLE', DatExigible);
    TOB_FilleTete.PutValue('PDU_DATELIMDEPOT', DatLimDepot);
    TOB_FilleTete.PutValue('PDU_DATEREGLEMENT', DatReglt);

    if (IndNeant2 = TRUE) then
      TOB_FilleTete.PutValue('PDU_PAIEMENT', IDate1900)
    else
      TOB_FilleTete.PutValue('PDU_PAIEMENT', DatePaye);

    TOB_FilleTete.PutValue('PDU_CESSATION', IDate1900);
    TOB_FilleTete.PutValue('PDU_CONTINUATION', IDate1900);
    TOB_FilleTete.PutValue('PDU_SUSPENSION', '-');
    TOB_FilleTete.PutValue('PDU_MAINTIENT', '-');
    TOB_FilleTete.PutValue('PDU_DECLARANT', RaisonSoc);
    TOB_FilleTete.PutValue('PDU_DECLARANTSUITE', ContactDecla);
    TOB_FilleTete.PutValue('PDU_TELEPHONEDECL', TelDeclarant);
    TOB_FilleTete.PutValue('PDU_FAXDECLARANT', FaxDeclarant);
    if (Rupture = TRUE) then
      TOB_FilleTete.PutValue('PDU_LIGNEOPTIQUE', LGOptiqueEnCours)
    else
      TOB_FilleTete.PutValue('PDU_LIGNEOPTIQUE', LigneOptique);
    TOB_FilleTete.PutValue('PDU_ACOMPTES', 0);
    TOB_FilleTete.PutValue('PDU_REGULARISATION', 0);
    TOB_FilleTete.PutValue('PDU_MONNAIETENUE', Monnaie);

//PT84
//  CalculEffectifs(DucsDoss, Regroupt);
    CalculEffectifs(DucsDoss, Regroupt, PogNature);

    if (IndNeant2 = TRUE) then
      TOB_FilleTete.PutValue('PDU_NBSALFPE', 0)
    else
      TOB_FilleTete.PutValue('PDU_NBSALFPE', EffTot);
    TOB_FilleTete.PutValue('PDU_NBSALHAP', 0);
{$IFNDEF DUCS41}
    TOB_FilleTete.PutValue ('PDU_NBSALQ920',EffHom);
    TOB_FilleTete.PutValue ('PDU_NBSALQ921',EffFem);
    TOB_FilleTete.PutValue ('PDU_NBSALQ922',EffCDIH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ923',EffCDIF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ924',EffAppH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ925',EffAppF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ926',Eff65H);
    TOB_FilleTete.PutValue ('PDU_NBSALQ927',Eff65F);
    TOB_FilleTete.PutValue ('PDU_NBSALQ928',EffCadH);
    TOB_FilleTete.PutValue ('PDU_NBSALQ929',EffCadF);
    TOB_FilleTete.PutValue ('PDU_NBSALQ934',EffProfH);    // PT62
    TOB_FilleTete.PutValue ('PDU_NBSALQ935',EffProfF);    // PT62
    TOB_FilleTete.PutValue ('PDU_NBSALQ936', EffCDDH);    // effectif 930 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ937', EffCDDF);    // effectif 931 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ962',EffCNEH);     // effectif 932 du CdC
    TOB_FilleTete.PutValue ('PDU_NBSALQ963',EffCNEF);     // effectif 933 du CdC
{$ELSE}
    TOB_FilleTete.PutValue('PDU_NBSALQ920', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ921', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ922', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ923', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ924', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ925', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ926', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ927', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ928', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ929', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ934', 0);
    TOB_FilleTete.PutValue('PDU_NBSALQ935', 0);
{$ENDIF}
    TOB_FilleTete.PutValue('PDU_TOTHOMMES', EffHom);
    TOB_FilleTete.PutValue('PDU_TOTFEMMES', EffFem);
    TOB_FilleTete.PutValue('PDU_TOTAPPRENTI', EffApp);
    TOB_FilleTete.PutValue('PDU_TOTOUVRIER', EffOuv);
    TOB_FilleTete.PutValue('PDU_TOTETAM', EffEtam);
    TOB_FilleTete.PutValue('PDU_TOTCADRES', EffCad);
    TOB_FilleTete.PutValue('PDU_DATEPREVEL', IDate1900);
    TOB_FilleTete.PutValue('PDU_EMETTSOC', Emetteur);
{$IFNDEF DUCS41}
    TOB_FilleTete.PutValue ('PDU_NBSALQ944',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ945',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ960',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ961',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ964',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ965',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ966',0);
    TOB_FilleTete.PutValue ('PDU_NBSALQ967',0);
    TOB_FilleTete.PutValue ('PDU_ECARTZE1','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE2','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE3','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE4','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE5','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE6','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE7','-');
    TOB_FilleTete.PutValue ('PDU_ECARTZE8','-');
{$ENDIF}
    if (DucsDoss = True) then
      TOB_FilleTete.PutValue('PDU_DUCSDOSSIER', 'X')
    else
      TOB_FilleTete.PutValue('PDU_DUCSDOSSIER', '-');

    //Centre Payeur
    if (PogNature <> '300') then
      TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', '')
    else
    begin
      if (PogCpInfo = '') or (PogCpInfo = '001') or (PogCpInfo = '002') then
      { type info = libre ou n° affiliation ou n° interne }
        TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', PogCentrePayeur)
      else
        if (PogCpInfo = '003') then
        { type info = siret }
        begin
          if (Rupture = TRUE) then
            TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', SiretEnCours)
          else
            TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', EtSiret);
        end
        else
          if (PogCpInfo = '004') then
          { type info = ape }
          begin
            if (Rupture = TRUE) then
              TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', ApeEnCours)
            else
              TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', EtApe);
          end
          else
            if (PogCpInfo = '005') then
            { type info = groupe interne }
            begin
              if (RupGpInt = TRUE) then
                TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', GpIntEnCours)
              else
                TOB_FilleTete.PutValue('PDU_CENTREPAYEUR', '');
            end;
    end;

{$IFNDEF DUCS41}
    TOB_FilleTete.PutValue ('PDU_TYPBORDEREAU',PogTypBordereau);
{$ENDIF}

    if (Tdetail = nil) or (IndNeant2 = TRUE) then
    { il s'agit d'une Ducs Néant}
      IndNeant := TRUE;

  {Alimentation DUCSDETAIL}
  {-----------------------}
  {
    Traitement des ruptures
  }

    while (TDetail <> nil) and
      (IndNeant2 = FALSE) and
      (fctRupture(TDetail, 0) = TRUE) do
    begin
      IndDucs := TRUE;
    {arrondi : récup. paramétrage organisme}
      BaseArr := WBaseArr;
      MontArr := WMontArr;
    {affectation des variables de traitement des lignes de détail}
      CodifLue := TDetail.GetValue('CODIFICATION');
      CodifEnCours := TDetail.GetValue('CODIFICATION');
// loi fillon
// d PT63
      if (TDetail.GetValue('REGIMEALSACE') =  'X') then
        AlsaceMoselle := true
      else
        AlsaceMoselle := false;

      if (AlsaceMoselle) then
        QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
        'PDP_TYPECOTISATION, PDP_BASETYPARR,PDP_MTTYPARR, PDP_CONDITION' +
        ' FROM DUCSPARAM ' +
        ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' + CodifEnCours + '"' +
        ' ORDER BY PDP_PREDEFINI', True)
      else
        QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
        'PDP_TYPECOTISATION, PDP_BASETYPARR,PDP_MTTYPARR, PDP_CONDITION' +
        ' FROM DUCSPARAM ' +
        ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifEnCours + '"' +
        ' ORDER BY PDP_PREDEFINI', True);
// f PT63
      Predef := '';
{$IFDEF EAGLCLIENT}
      SortDos := 0;
      KK := 0;
{$ENDIF}
      while not QDucsParam.eof do
      begin
        Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
        if (Predef = 'DOS') then
        begin
{$IFDEF EAGLCLIENT}
          SortDos := 1;
{$ENDIF}
          break;
        end;
{$IFDEF EAGLCLIENT}
        KK := KK + 1;
{$ENDIF}
        QDucsParam.Next;
      end;
{$IFDEF EAGLCLIENT}
      if SortDos = 1 then
        MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
      else
        MaTob := QDucsParam.Detail[KK - 1];
{$ENDIF}
      if (Predef <> '') then
      begin
{$IFNDEF EAGLCLIENT}
        TypeCotis := QDucsParam.FindField('PDP_TYPECOTISATION').AsString;
{$ELSE}
        TypeCotis := MaTob.GetValue('PDP_TYPECOTISATION');
{$ENDIF}
      end;
//    ferme (QDucsParam);
      TauxLu := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);

// loi fillon
      if (TypeCotis = 'M') then
        TauxLu := 0.00;
// d PT73
// loi TEPA
      if (TypeCotis = 'C') then       // base & montant
        TauxLu := 0.00;
      if (TypeCotis = 'D') then       // Quantité en base & mt
        TauxLu := 0.00;
// f PT73

      TauxEnCours := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);
// loi fillon
      if (TypeCotis = 'M') then
        TauxEnCours := 0.00;
// d PT73
// loi TEPA
      if (TypeCotis = 'C') then       // base & montant
        TauxEnCours := 0.00;

      if (TypeCotis = 'D') then       // quantité en base & mt
        TauxEnCours := 0.00;
// f PT73

      ATLu := Arrondi(TDetail.GetValue('TAUXAT'), 5);
      ATEnCours := Arrondi(TDetail.GetValue('TAUXAT'), 5);
      BaseLigne := Arrondi(TDetail.GetValue('BASECOT'), 5);
// loi fillon
      if (TypeCotis = 'M') then
        BaseLigne := 0.00;

      MontLigne := Arrondi(TDetail.GetValue('MTGLOBAL'), 5);
      SalLu := TDetail.GetValue('SALARIE');
      OrgLu := TDetail.GetValue('ORGANISME');


      if (personnalisee = False) then
        if (TDetail.GetValue('INSTITUTION') <> Null) then
          CodInstLu := TDetail.GetValue('INSTITUTION');
      if (RupSiret = True) then
        if (TDetail.GetValue('SIRET') <> Null) then
          WSiretLu := TDetail.GetValue('SIRET');
      if (RupNumInt = TRUE) then
        if (TDetail.GetValue('NUMINTERNE') <> Null) then
          WNoIntLu := TDetail.GetValue('NUMINTERNE');
      if (RupApe = TRUE) then
        if (TDetail.GetValue('APE') <> Null) then
          WApeLu := TDetail.GetValue('APE');
      if (RupGpInt = TRUE) then
        if (TDetail.GetValue('GROUPE') <> Null) then
          WGpLu := TDetail.GetValue('GROUPE');
      OrgEnCours := TDetail.GetValue('ORGANISME');
        if (TDetail.GetValue('INSTITUTION') <> Null) then
          CodInstEnCours := TDetail.GetValue('INSTITUTION');

      if (PogNature = '300') then
      {IRC : récup de l'institution encours }
        if (copy(IntToStr(NoInstEnCours), 1, 2) <> copy(CodifLue, 1, 2)) or
          ((NoInstEnCours = 0) and
          (copy(IntToStr(NoInstEnCours), 1, 1) = copy(CodifLue, 1, 1))) then
        begin
          if IsNumeric(copy(CodifLue, 1, 2)) then
            NoInstEnCours := StrToInt(copy(CodifLue, 1, 2))
          else
          begin
//PT65      MessageAlerte('Vérifier l''affectation des rubriques pour cet organisme');
            Anomalie := 'erreur d''affectation';
          end;
        end;
      Effectif := 0;
      FreeAndNil(TEffM);
      SalEnCours := ' ';
// d PT63
      if (AlsaceMoselle) then
        QDucsParamAT := OpenSQL('SELECT PDP_PREDEFINI,PDP_CODIFALSACE FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFALSACE LIKE "' + Copy(CodifEnCours, 1, 6) + '%"' +
                      ' ORDER BY PDP_PREDEFINI',
                      True)
      else
        QDucsParamAT := OpenSQL('SELECT PDP_PREDEFINI,PDP_CODIFICATION FROM DUCSPARAM ' +
                      'WHERE ##PDP_PREDEFINI## PDP_TYPECOTISATION="A" AND ' +
                      'PDP_CODIFICATION LIKE "' + Copy(CodifEnCours, 1, 6) + '%"' +
                      ' ORDER BY PDP_PREDEFINI',
                      True);
// f PT63
      CodifAT := '';
      while not QDucsParamAT.Eof do
      begin
// d PT63
        if (AlsaceMoselle) then
          CodifAT := QDucsparamAT.FindField('PDP_CODIFALSACE').AsString
        else
          CodifAT := QDucsparamAT.FindField('PDP_CODIFICATION').AsString;
// f PT63
        Predef := QDucsparamAT.FindField('PDP_PREDEFINI').AsString;
        if (Predef = 'DOS') then break;
        QDucsParamAT.Next;
      end;
      Ferme(QDucsParamAT);

      while ((IndRup = False) and
        (CodifLue = CodifEnCours) and
        (ARRONDI(TauxLu, 5) = ARRONDI(TauxEnCours, 5)) and
        (((CodifAT <> '') and (ARRONDI(ATLu, 5) = ARRONDI(ATEnCours, 5)) and
        (Copy(CodifEnCours, 1, 6) = Copy(CodifAT, 1, 6))) or
        (Copy(CodifEnCours, 1, 6) <> Copy(CodifAT, 1, 6))) and
        (TDetail <> nil)) do
    {pour une même codification: traitement d'une ligne}
      begin
    {(Cumuls)}
        BaseCodif := BaseCodif + BaseLigne;
        MontCodif := MontCodif + MontLigne;
        MontForfait := MontLigne;
        if (SalEnCours <> SalLu) then
      {(Incrémentation effectif)}
        begin
          if (TypeCotis = 'M') then
          begin
            if (TEffM = nil) then
              TEffM := TOB.Create('effectif forfait', nil, -1);

            TEffMFille := TEffM.FindFirst(['SALARIE'], [SalLu], TRUE);
            if (TEffMFille = nil) then
            begin
              TEffMFille := TOB.Create('', TEffM, -1);
              TEffMFille.AddChampSupValeur('SALARIE', SalLu, FALSE);
            end;
            if (MontLigne <> 0) then
              effectif := TEffM.Detail.Count;
          end
          else
            Effectif := Effectif + 1;

          SalEnCours := SalLu;
        end;

      {ligne de cotisation suivante}
//PT62        TDetail := TCot2.FindNext([''], [''], TRUE);
        TDetail := TCot3.FindNext([''], [''], TRUE);
        if (TDetail <> nil) then
      {affectation des variables de traitement des lignes de détail}
        begin
          CodifLue := TDetail.GetValue('CODIFICATION');
          TauxLu := Arrondi(TDetail.GetValue('TAUXGLOBAL'), 5);

// loi fillon
          if (TypeCotis = 'M') then
            TauxLu := 0.00;
// d PT73
// loi TEPA
          if (TypeCotis = 'C') then     //  base & montant
             TauxLu := 0.00;

          if (TypeCotis = 'D') then     // quantité & montant
             TauxLu := 0.00;
// f PT73

          ATLu := Arrondi(TDetail.GetValue('TAUXAT'), 5);
          BaseLigne := Arrondi(TDetail.GetValue('BASECOT'), 5);
// loi fillon
          if (TypeCotis = 'M') then
            BaseLigne := 0.00;

          MontLigne := Arrondi(TDetail.GetValue('MTGLOBAL'), 5);
          Sallu := TDetail.GetValue('SALARIE');
          OrgLu := TDetail.GetValue('ORGANISME');
          if (personnalisee = False) then
            if (TDetail.GetValue('INSTITUTION') <> Null) then
              CodInstLu := TDetail.GetValue('INSTITUTION');
          IndRup := false;
          if (RupSiret = True) then
          begin
            if (TDetail.GetValue('SIRET') <> Null) then
              WSiretLu := TDetail.GetValue('SIRET');
            if (WSiretLu <> SiretEnCours) then IndRup := True;
          end;
          if (RupNumInt = TRUE) then
          begin
            if (TDetail.GetValue('NUMINTERNE') <> Null) then
              WNoIntLu := TDetail.GetValue('NUMINTERNE');
            if (WNoIntLu <> NumIntEnCours) then IndRup := True;
          end;
          if (RupApe = TRUE) then
          begin
            if (TDetail.GetValue('APE') <> Null) then
              WApeLu := TDetail.GetValue('APE');
            if (WApeLu <> ApeEnCours) then IndRup := True;
          end;
          if (RupGpInt = TRUE) then
          begin
            if (TDetail.GetValue('GROUPE') <> Null) then
              WGpLu := TDetail.GetValue('GROUPE');
            if (WGpLu <> GpIntEnCours) then IndRup := True;
          end;
        end;
      end; { fin pour une même codification: traitement d'une ligne}

    {Mise à jour  les lignes}
      if (CodifEnCours = Copy(CodifAT, 1, 6) + 'D') then
    {Traitement TAUX AT}
      begin
        NoAT := NoAT + 1;
        if (NoAT > 9) then NoAT := 0; { Au cas où + de 10 tx AT Différent}
// d PT63
        if (AlsaceMoselle) then
          QDucsParamAT := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' + CodifAT + '" ' +
                                  'ORDER BY PDP_PREDEFINI', True)
        else
          QDucsParamAT := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifAT + '" ' +
                                  'ORDER BY PDP_PREDEFINI', True);
// f PT63
        Predef := '';
{$IFDEF EAGLCLIENT}
        SortDos := 0;
        KK := 0;
{$ENDIF}
        while not QDucsParamAT.eof do
        begin
          Predef := QDucsparamAT.FindField('PDP_PREDEFINI').AsString;
          if (Predef = 'DOS') then
          begin
{$IFDEF EAGLCLIENT}
            SortDos := 1;
{$ENDIF}
            break;
          end;
{$IFDEF EAGLCLIENT}
          KK := KK + 1;
{$ENDIF}
          QDucsParamAT.Next;
        end;

        if (Predef <> '') then
        begin
         AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
          TOB_FilleLignes.PutValue('PDD_CODIFICATION', Copy(CodifAT, 1, 2) +
            IntToStr(NoAT) + Copy(CodifAT, 4, 4));

          Cod100 := Copy(CodifAT, 1, 2) + IntToStr(NoAT) + Copy(CodifAT, 4, 4);
{$IFDEF EAGLCLIENT}
          if SortDos = 1 then MaTobAT := QDucsParamAT.Detail[QDucsParam.Detail.count - 1]
          else MaTobAT := QDucsParamAT.Detail[KK - 1];
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            MaTobAT.GetValue('PDP_TYPECOTISATION'));
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            MaTobAT.GetValue('PDP_LIBELLE'));
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            MaTobAT.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            QDucsParamAT.FindField('PDP_TYPECOTISATION').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            QDucsParamAT.FindField('PDP_LIBELLE').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            QDucsParamAT.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
          TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Arrondi(AtEnCours, 5));
          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', 0.00);
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
            copy(Cod100, CodifPosDebut, CodifLongEdit));
        end;
// d PT94
{$IFDEF EAGLCLIENT}
        if assigned(MaTobAT) then
          FreeAndNil (MaTobAT);
{$ENDIF}
// f PT94
        Ferme(QDucsParamAT);
      end; {If (CodifEnCours = '1A0100D')}
    {fin traitement Taux AT}

      if (Predef <> '') then
      begin
        if (MontCodif <> 0.00) or
          ((MontCodif = 0.00) and
           ((Arrondi(AtEnCours, 5) <> 0.00) and
{$IFNDEF EAGLCLIENT}
          (QDucsParam.FindField('PDP_TYPECOTISATION').AsString = 'M')) or
          (QDucsParam.FindField('PDP_TYPECOTISATION').AsString = 'I'))
{$ELSE}
          (MaTob.GetValue('PDP_TYPECOTISATION') = 'M')) or
          (MaTob.GetValue('PDP_TYPECOTISATION') = 'I'))
{$ENDIF}
        then
      {Une ligne n'est éditée que si le mt est différent de zéro ou s'il
       est à zéro et que le taux est différent de 0 et
       s'il s'agit d'une ligne "M"ontant ou  'I"ntitulé}
        begin
          AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
          if (Copy(CodifEnCours, 1, 6) = Copy(CodifAT, 1, 6)) then
        {CAS GENERAL URSSAF}
          begin
            if (Copy(CodifEnCours, 7, 1) = 'D') then
          { 1AxyyyD }
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                Copy(CodifEnCours, 1, 2) +
                IntToStr(NoAT) +
                Copy(CodifEnCours, 4, 7))

            else
            begin
          {1AxyyyP}
              TAT := TOB_FilleTete.FindFirst(['PDD_TAUXCOTISATION',
                'PDD_TYPECOTISATION',
                  'PDD_ETABLISSEMENT',
                  'PDD_ORGANISME',
                  'PDD_NUM',
                  'PDD_DATEDEBUT',
                  'PDD_DATEFIN'],
                  [ATEnCours,
                'A',
                  EtabEnCours,
                  organisme,
                  NumDucs, // DB2
                  StrToDate(DebPer),
                  StrToDate(FinPer)],
                  TRUE);

              if (TAT <> nil) then
                if (copy(TAT.GetValue('PDD_CODIFICATION'), 3, 1)) = '' then
                  NoATP := 0
                else
//PT62                  NoATP := StrToInt(Copy(TAT.GetValue('PDD_CODIFICATION'), 3, 1));
                  NoATP := NOAT;
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                Copy(CodifEnCours, 1, 2) +
                IntToStr(NoATP) +
                Copy(CodifEnCours, 4, 7));

            end;
          end
          else
            if (PogNature = '300') then
         {Codif IRC : La codif est modifée en fct de la rupture
                             sur institution }
            begin
              NomChamp[1] := 'PDD_CODIFICATION';
              ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5);
              NomChamp[2] := 'PDD_NUMORDRE';
              ValeurChamp[2] := ligne;

              TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              while (TobStExiste <> nil) do
              begin
                ValChST[1] := ColleZeroDevant(NoInstEnCours + 1, 2) + 'ZZZZZ';
                NomChampST[1] := 'PDD_CODIFICATION';
                if (Personnalisee = False) or
                  ((Personnalisee = True) and
                  (TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE) <> nil)) then
                begin
                  NoInstEnCours := NoInstEnCours + 1;
                  ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5);
                  TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
                end
                else
                begin
                  ligne := ligne + 1;
                  ValeurChamp[2] := ligne;
                  TOB_FilleLignes.PutValue('PDD_NUMORDRE', ligne);
                  TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
                end;
              end;
              TOB_FilleLignes.PutValue('PDD_CODIFICATION',
                ColleZeroDevant(NoInstEnCours, 2) + Copy(CodifEnCours, 3, 5))
            end
            else
         {autres natures de DUCS}
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifEnCours);
{$IFDEF EAGLCLIENT}
          if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
          else MaTob := QDucsParam.Detail[KK - 1];
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            MaTob.GetValue('PDP_TYPECOTISATION'));
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            MaTob.GetValue('PDP_LIBELLE'));
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            MaTob.GetValue('PDP_LIBELLESUITE'));
           // Type de cotisation
          TypeCotis := MaTob.GetValue('PDP_TYPECOTISATION');
          CSC := 'SCS';
          if (CondSpec) then
            CSC := MaTob.GetValue('PDP_CONDITION');
          if (CSC = '') then
            CSC := 'SCS';

           // arrondi de la codification
          BaseArrCod := MaTob.GetValue('PDP_BASETYPARR');
          MontArrCod := MaTob.GetValue('PDP_MTTYPARR');
{$ELSE}
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
            QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLE',
            QDucsParam.FindField('PDP_LIBELLE').AsString);
          TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
            QDucsParam.FindField('PDP_LIBELLESUITE').AsString);

           {Type de cotisation}
          TypeCotis := QDucsParam.FindField('PDP_TYPECOTISATION').AsString;
          CSC := 'SCS';
          if (CondSpec) then
            CSC := QDucsParam.FindField('PDP_CONDITION').AsString;
          if (CSC = '') then
            CSC := 'SCS';

           {arrondi de la codification}
          BaseArrCod := QDucsParam.FindField('PDP_BASETYPARR').AsString;
          MontArrCod := QDucsParam.FindField('PDP_MTTYPARR').AsString;
{$ENDIF}
// d PT94
{$IFDEF EAGLCLIENT}
          if assigned(MaTob) then
            FreeAndNil (MaTob);
{$ENDIF}
// f PT94

          Ferme(QDucsParam);

         {détermination de l'arrondi appliqué}
          if (BaseArrCod <> '') then BaseArr := BaseArrCod;
          if (MontArrCod <> '') then MontArr := MontArrCod;

         {arrondi de la base}
          if (BaseArr <> '') and (BaseArr <> 'P') then
          begin
            BaseCodif := Int(BaseCodif);
            MontForfait := Int(MontForfait);
            if (BaseArr = 'S') then BaseCodif := BaseCodif + 1;
          end;
          if (BaseArr = 'P') then
          begin
            BaseCodif := Arrondi(BaseCodif, 0);
            MontForfait := Arrondi(MontForfait, 0);
          end;

         {Calcul montant de la cotisation}
          if (TypeCotis = 'T') then
         {Base * Taux}
            MontCodif := (BaseCodif * TauxEnCours) / 100.0;
          if (TypeCotis = 'Q') then
          begin
         {Quantité (Base * Qté)}
            MontCodif := (MontForfait * Effectif);
            BaseCodif := MontForfait;
          end;

         {Arrondi du Montant}
          if (MontArr <> '') and (MontArr <> 'P') then
          begin
            MontCodif := Int(MontCodif);
            if (MontArr = 'S') then MontCodif := MontCodif + 1;
          end;
          if (MontArr = 'P') then
          begin
            MontCodif := Arrondi(MontCodif, 0);
          end;

         { Cumul SOUS TOTAL }
          MontSt := MontST + MontCodif;
          IndCalcSt := true;
          TOB_FilleLignes.PutValue('PDD_EFFECTIF', Effectif);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION',
            Arrondi(BaseCodif, 5));

          if (TypeCotis <> 'Q') then
            TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Arrondi(TauxEnCours, 5))
          else
            TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', Effectif);

          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontCodif, 5));
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', CSC);
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
            copy(CodifEnCours, CodifPosDebut, CodifLongEdit));
          if (PogNature = '300') then
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', CodInstEnCours)
          else
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');

          NoInst := CodInstEnCours;
        end { fin if MonCodif<>0}
        else
          Ferme(QDucsParam);
      end;

    {TRAITEMENT DES TYPES DE CODIFICATION PARTICULIERS}
    {SOUS TOTAL}
      if (PogBoolST = TRUE) then
      begin
        if (MontSt <> 0.00) or ((PogNature = '300') and
          (PogSTPos <> 0) and (PogSTLong <> 0)) then
        begin
          if (PogNature = '200') or
            (PogNature = '600') or
            (PogNature = '300') then
        {ASSEDIC ou MSA  ou IRC personnalisée}
          begin
            if (Copy(CodifLue, PogSTPos, PogStLong) <>
              Copy(CodifEnCours, PogSTPos, PogStLong)) then
            begin
              if (PogNature = '200') then
            {ASSEDIC}
                CodifST := Copy(CodifEnCours, 1, 5) + 'ZZ';
              if (PogNature = '600') then
            {MSA}
                CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
              if (PogNature = '300') then
            {IRC personnalisée}
                if (PogSTPos = 1) then
                begin
                  CodifST := Copy(CodifEnCours, 1, PogSTLong) +
                    Copy(ZEDE, 1, (7 - PogSTLong));
                end
                else
                begin
                  CodifST := Copy(CodifEnCours, 1, PogSTPos - 1) +
                    Copy(CodifEnCours, PogSTPos, PogSTLong) +
                    Copy(ZEDE, 1, (7 - (PogStPos + PogSTLong - 1)));
                end;
// d PT63
              if (AlsaceMoselle) then
                QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                      'PDP_LIBELLESUITE,' +
                                      'PDP_TYPECOTISATION' +
                                      ' FROM DUCSPARAM ' +
                                      ' WHERE ##PDP_PREDEFINI## ' +
                                      'PDP_CODIFALSACE = "' + CodifST + '"' +
                                      ' ORDER BY PDP_PREDEFINI', True)
              else
                QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                      'PDP_LIBELLESUITE,' +
                                      'PDP_TYPECOTISATION' +
                                      ' FROM DUCSPARAM ' +
                                      ' WHERE ##PDP_PREDEFINI## ' +
                                      'PDP_CODIFICATION = "' + CodifST + '"' +
                                      ' ORDER BY PDP_PREDEFINI', True);
// f PT63
              Predef := '';
{$IFDEF EAGLCLIENT}
              SortDos := 0;
              KK := 0;
{$ENDIF}
              while not QDucsParam.eof do
              begin
                Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
                if (Predef = 'DOS') then
                begin
{$IFDEF EAGLCLIENT}
                  SortDos := 1;
{$ENDIF}
                  break;
                end;
{$IFDEF EAGLCLIENT}
                KK := KK + 1;
{$ENDIF}
                QDucsParam.Next;
              end;
              if (Predef <> '') then
              begin
                ValChST[1] := CodifST;
                NomChampST[1] := 'PDD_CODIFICATION';
                TOB_FilleLignes := TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE);
                if (Personnalisee = TRUE) and
                  (TOB_FilleLignes <> nil) then
                begin
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT',
                    TOB_FilleLignes.GetValue('PDD_MTCOTISAT') +
                    Arrondi(MontST, 5));
                end
                else
                begin
                  AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
                  TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
                  if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
                  else MaTob := QDucsParam.Detail[KK - 1];
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    MaTob.GetValue('PDP_TYPECOTISATION'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    MaTob.GetValue('PDP_LIBELLE'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    QDucsParam.FindField('PDP_LIBELLE').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT94
{$IFDEF EAGLCLIENT}
                  if assigned(MaTob) then
                    FreeAndNil (MaTob);
{$ENDIF}
// f PT94
                  Ferme(QDucsParam);
                  TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
                  TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
                  TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
                  TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
                  if (PogNature = '200') then
              { ASSEDIC }
                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, 5, 1) + ' ');
                  if (PogNature = '600') then
              { MSA }
                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, 3, 2));

                  if (PogNature = '300') then
                  begin
              {IRC personnalisée}
                    if (PogSTPos < 3) then
                    begin
                      PosDeb := 3;
                      LongEdit := PogSTLong - 2;
                    end
                    else
                    begin
                      PosDeb := PogStPos;
                      LongEdit := PogStLong;
                    end;

                    TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                      copy(CodifST, PosDeb, LongEdit));
                    TOB_FilleLignes.PutValue('PDD_INSTITUTION',
                      CodInstEnCours);
                  end;
                end;
              {RAZ Cumul Sous Total}
                MontST := 0.00;
                IndCalcST := false;
              end; {if(Predef <> '')}
            end; { if (Copy(CodifLue,PogSTPos,PogStLong)...}
          end; {if (PogNature = '200') OR (PogNature = '600')}

          if (PogNature = '300') and (((PogSTPos = 0) and (PogSTLong = 0))) then
        {IRC  standardisée}
          begin
            if ((((RupSiret = True) and (WsiretLu <> siretEnCours)) or
              ((RupApe = True) and (WApeLu <> ApeEnCours)) or
              ((RupNumInt = True) and (WNoIntLu <> NumIntEnCours)) or
              ((RupGpInt = True) and (WGpLu <> GpIntEnCours))) or
              ((Orglu <> OrgEnCours) or
              (CodInstLu <> CodInstEnCours)) or
              ((CodifLue <> CodifEnCours) and
              (Copy(CodifLue, 1, 1) = '8') and
              (Copy(CodifEnCours, 1, 1) <> '8'))) then
            begin
              NomChamp[1] := 'PDD_CODIFICATION';
              ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
              NomChamp[2] := '';
              ValeurChamp[2] := '';

              TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              while (TobStExiste <> nil) do
              begin
                NoInstEnCours := NoInstEnCours + 1;
                ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
                TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
              end;
              CodifST := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';

              AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION', 'S');

             {recup Code Institution et Libellé}
              QInst := OpenSql('SELECT PIP_ABREGE ' +
                ' FROM INSTITUTIONPAYE ' +
                ' WHERE PIP_INSTITUTION = "' +
                NoInst + '"', True);

              if not QInst.eof then
              begin
                LibInst := 'ST ' + NoInst + ' ' +
                  Copy(QInst.FindField('PIP_ABREGE').AsString, 1, 7);
                TOB_FilleLignes.PutValue('PDD_LIBELLE', LibInst);
              end;
              Ferme(QInst);

              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
             {Le code n'est pas édité}
// ci-après les lignes seront à activer si le terme "AGFF" doit apparaître sur
// la ligne de Sous-total
//             if (Copy(CodifEnCours,1,1) = '8') then
//               TOB_FilleLignes.PutValue ('PDD_CODIFEDITEE','8')
//             else
              TOB_FilleLignes.PutValue('PDD_CODIFEDITEE', ' ');
              if (PogNature = '300') then
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', NoInst)
              else
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');
             {RAZ Cumul Sous Total}
              MontST := 0.00;
              IndCalcSt := false;
            end; { if (Orglu <> OrgEnCours)}
          end; { if (PogNature = '300')}

          if (PogNature = '700') then
        {BTP}
          begin
            if ((PogPeriod1 = PeriodEtat) and (PogTypDecl1 = TRUE)) or
              ((PogPeriod2 = PeriodEtat) and (PogTypDecl2 = TRUE)) then
          {Déclaration avec calcul}
            begin
              if (Copy(CodifLue, PogSTPos, PogStLong) <>
                Copy(CodifEnCours, PogSTPos, PogStLong)) then
              begin
                CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
// d PT63
                if (AlsaceMoselle) then
                  QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                        'PDP_LIBELLESUITE,' +
                                        'PDP_TYPECOTISATION' +
                                        ' FROM DUCSPARAM ' +
                                        ' WHERE ##PDP_PREDEFINI##' +
                                        ' PDP_CODIFALSACE = "' + CodifST + '"' +
                                        ' ORDER BY PDP_PREDEFINI', True)
                else
                  QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, ' +
                                        'PDP_LIBELLESUITE,' +
                                        'PDP_TYPECOTISATION' +
                                        ' FROM DUCSPARAM ' +
                                        ' WHERE ##PDP_PREDEFINI##' +
                                        ' PDP_CODIFICATION = "' + CodifST + '"' +
                                        ' ORDER BY PDP_PREDEFINI', True);
// f PT63
                Predef := '';
{$IFDEF EAGLCLIENT}
                SortDos := 0;
                KK := 0;
{$ENDIF}
                while not QDucsParam.eof do
                begin
                  Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
                  if (Predef = 'DOS') then
                  begin
{$IFDEF EAGLCLIENT}
                    SortDos := 1;
{$ENDIF}
                    break;
                  end;
{$IFDEF EAGLCLIENT}
                  KK := KK + 1;
{$ENDIF}
                  QDucsParam.Next;
                end;

                if (Predef <> '') then
                begin
                  AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
                  TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
                  if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
                  else MaTob := QDucsParam.Detail[KK - 1];
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    MaTob.GetValue('PDP_TYPECOTISATION'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    MaTob.GetValue('PDP_LIBELLE'));
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
                  TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                    QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLE',
                    QDucsParam.FindField('PDP_LIBELLE').AsString);
                  TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                    QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT94
{$IFDEF EAGLCLIENT}
                  if assigned(MaTob) then
                    FreeAndNil (MaTob);
{$ENDIF}
// f PT94
                  Ferme(QDucsParam);

                  TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
                  TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
                  TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
                  TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
                  TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
                  TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                    copy(CodifST, 3, 2));
                {RAZ Cumul Sous Total}
                  MontST := 0.00;
                  IndCalcSt := False;
                end; { if (Predef <> '')}
              end; {if (Copy(CodifLue,PogS}
            end; {if ((PogPeriod1 = PeriodEtat)}
          end; {if (PogNature = '700'}
        end; {if (MontST <> 0)}
      end; {if (PogBoolST = TRUE)}

      if (CodifLue <> CodifEnCours) then ligne := 0;
//PT81      if (copy(CodifLue, 1, 6) <> copy(CodifEnCours, 1, 6)) then NoAT := -1;
      BaseCodif := 0.00;
      MontCodif := 0.00;

    end; { fin while (tDetail....}
    FreeAndNil(TEffM);

    if (TDetail = nil) and (Rupture = False) then
    {il n'y a plus de détail}
      fctRupture(TDetail, 1);

  {deb traitement dernière ligne de sous total}

  {SOUS TOTAL}
    if (PogBoolST = TRUE) and (IndCalcSt = true) then
    begin
      if (MontSt <> 0.00) or ((PogNature = '300') and
        (PogSTPos <> 0) and (PogSTLong <> 0)) then
      begin
        if (PogNature = '200') or
          (PogNature = '600') or
          (PogNature = '300') then
      {ASSEDIC ou MSA  ou IRC personnalisée}
        begin
          if (PogNature = '200') then
        {ASSEDIC}
            CodifST := Copy(CodifEnCours, 1, 5) + 'ZZ';

          if (PogNature = '600') then
        {MSA}
            CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';

          if (PogNature = '300') then
        {IRC personnalisée}
            if (PogSTPos = 1) then
            begin
              CodifST := Copy(CodifEnCours, 1, PogSTLong) +
                Copy(ZEDE, 1, (7 - PogSTLong));
            end
            else
            begin
              CodifST := Copy(CodifEnCours, 1, PogSTPos - 1) +
                Copy(CodifEnCours, PogSTPos, PogSTLong) +
                Copy(ZEDE, 1, (7 - (PogStPos + PogSTLong - 1)));
            end;

// d  PT63
          if (AlsaceMoselle) then
            QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE = "' +
                                   CodifST + '"' +
                                  ' ORDER BY PDP_PREDEFINI', True)
          else
            QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                  'PDP_TYPECOTISATION' +
                                  ' FROM DUCSPARAM ' +
                                  ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' +
                                   CodifST + '"' +
                                  ' ORDER BY PDP_PREDEFINI', True);
// f PT63                                  
          Predef := '';
{$IFDEF EAGLCLIENT}
          SortDos := 0;
          KK := 0;
{$ENDIF}
          while not QDucsParam.eof do
          begin
            Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
            if (Predef = 'DOS') then
            begin
{$IFDEF EAGLCLIENT}
              SortDos := 1;
{$ENDIF}
              break;
            end;
{$IFDEF EAGLCLIENT}
            KK := KK + 1;
{$ENDIF}
            QDucsParam.Next;
          end;
          if (Predef <> '') then
          begin
            ValChST[1] := CodifST;
            NomChampST[1] := 'PDD_CODIFICATION';
            TOB_FilleLignes := TOB_FilleTete.FindFirst(NomChampST, ValChST, TRUE);
            if (Personnalisee = TRUE) and
              (TOB_FilleLignes <> nil) then
            begin
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT',
                TOB_FilleLignes.GetValue('PDD_MTCOTISAT') +
                Arrondi(MontST, 5));
            end
            else
            begin
              AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
          If SortDos = 1 then
            MaTob := QDucsParam.Detail [QDucsParam.Detail.count-1]
          else
            MaTob := QDucsParam.Detail [KK-1] ;
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                MaTob.GetValue('PDP_TYPECOTISATION'));
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                MaTob.GetValue('PDP_LIBELLE'));
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                QDucsParam.FindField('PDP_LIBELLE').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT94
{$IFDEF EAGLCLIENT}
              if assigned(MaTob) then
                FreeAndNil (MaTob);
{$ENDIF}
// f PT94
              Ferme(QDucsParam);
              
              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');

              if (PogNature = '200') then
          {ASSEDIC}
                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, 5, 1) + ' ');
              if (PogNature = '600') then
          {MSA}
                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, 3, 2));

              if (PogNature = '300') then
              begin
          {IRC personnalisée}
                if (PogSTPos < 3) then
                begin
                  PosDeb := 3;
                  LongEdit := PogSTLong - 2;
                end
                else
                begin
                  PosDeb := PogStPos;
                  LongEdit := PogStLong;
                end;

                TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                  copy(CodifST, PosDeb, LongEdit));
                TOB_FilleLignes.PutValue('PDD_INSTITUTION', CodInstEnCours);
              end;
            end;
          {RAZ Cumul Sous Total}
            MontST := 0.00;
            IndCalcSt := false;
          end; { if (Predef <> '')}
        end; { if (PogNature = '200') OR (PogNature = '600')}


        if (PogNature = '300') and (((PogSTPos = 0) and (PogSTLong = 0))) then
      {IRC  standardisée}
        begin
          NomChamp[1] := 'PDD_CODIFICATION';
          ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
          NomChamp[2] := '';
          ValeurChamp[2] := '';

          TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
          while (TobStExiste <> nil) do
          begin
            NoInstEnCours := NoInstEnCours + 1;
            ValeurChamp[1] := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
            TobSTExiste := TOB_FilleTete.FindFirst(NomChamp, ValeurChamp, TRUE);
          end;
          CodifST := ColleZeroDevant(NoInstEnCours, 2) + 'ZZZZZ';
          AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
          TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
          TOB_FilleLignes.PutValue('PDD_TYPECOTISATION', 'S');
        {recup Code Institution et Libellé}
          QInst := OpenSql('SELECT PIP_ABREGE' +
            ' FROM INSTITUTIONPAYE ' +
            ' WHERE PIP_INSTITUTION = "' + NoInst + '"', True);
          if not QInst.eof then
          begin
            LibInst := 'ST ' + NoInst + ' ' +
              Copy(QInst.FindField('PIP_ABREGE').AsString, 1, 7);
            TOB_FilleLignes.PutValue('PDD_LIBELLE', LibInst);
          end;
          Ferme(QInst);

          TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
          TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
          TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
          TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
          TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
        {Le code n'est pas édité}
// ci-après les lignes seront à activer si le terme "AGFF" doit apparaître sur
// la ligne de Sous-total
//        if (Copy(CodifEnCours,1,1)='8') then
//          TOB_FilleLignes.PutValue ('PDD_CODIFEDITEE','8')
//        else
          TOB_FilleLignes.PutValue('PDD_CODIFEDITEE', ' ');
          if (PogNature = '300') then
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', NoInst)
          else
            TOB_FilleLignes.PutValue('PDD_INSTITUTION', '');
        {RAZ Cumul Sous Total}
          MontST := 0.00;
          IndCalcSt := false;
        end; { if (PogNature = '300')}

        if (PogNature = '700') then
      {BTP}
        begin
          if ((PogPeriod1 = PeriodEtat) and (PogTypDecl1 = TRUE)) or
            ((PogPeriod2 = PeriodEtat) and (PogTypDecl2 = TRUE)) then
        {Déclaration avec calcul}
          begin
            CodifST := Copy(CodifEnCours, 1, 4) + 'ZZZ';
// d @à
            if (AlsaceMoselle) then
              QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                    'PDP_TYPECOTISATION' +
                                    ' FROM DUCSPARAM ' +
                                    ' WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE= "' + CodifST + '"' +
                                    ' ORDER BY PDP_PREDEFINI', True)
            else
              QDucsParam := OpenSql('SELECT PDP_PREDEFINI,PDP_LIBELLE, PDP_LIBELLESUITE,' +
                                    'PDP_TYPECOTISATION' +
                                    ' FROM DUCSPARAM ' +
                                    ' WHERE ##PDP_PREDEFINI## PDP_CODIFICATION = "' + CodifST + '"' +
                                    ' ORDER BY PDP_PREDEFINI', True);
// f PT63
            Predef := '';
{$IFDEF EAGLCLIENT}
            SortDos := 0;
            KK := 0;
{$ENDIF}
            while not QDucsParam.eof do
            begin
              Predef := QDucsparam.FindField('PDP_PREDEFINI').AsString;
              if (Predef = 'DOS') then
              begin
{$IFDEF EAGLCLIENT}
                SortDos := 1;
{$ENDIF}
                break;
              end;
{$IFDEF EAGLCLIENT}
             KK := KK + 1;
{$ENDIF}
              QDucsParam.Next;
            end;

            if (Predef <> '') then
            begin
              AlimDetail(TOB_FilleLignes, TOB_FilleTete, ligne, Regroupt);
              TOB_FilleLignes.PutValue('PDD_CODIFICATION', CodifST);
{$IFDEF EAGLCLIENT}
              if SortDos = 1 then MaTob := QDucsParam.Detail[QDucsParam.Detail.count - 1]
              else MaTob := QDucsParam.Detail[KK - 1];
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                MaTob.GetValue('PDP_TYPECOTISATION'));
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                MaTob.GetValue('PDP_LIBELLE'));
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                MaTob.GetValue('PDP_LIBELLESUITE'));
{$ELSE}
              TOB_FilleLignes.PutValue('PDD_TYPECOTISATION',
                QDucsParam.FindField('PDP_TYPECOTISATION').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLE',
                QDucsParam.FindField('PDP_LIBELLE').AsString);
              TOB_FilleLignes.PutValue('PDD_LIBELLESUITE',
                QDucsParam.FindField('PDP_LIBELLESUITE').AsString);
{$ENDIF}
// d PT94
{$IFDEF EAGLCLIENT}
              if assigned(MaTob) then
                FreeAndNil (MaTob);
{$ENDIF}
// f PT94
              Ferme(QDucsParam);

              TOB_FilleLignes.PutValue('PDD_EFFECTIF', 0);
              TOB_FilleLignes.PutValue('PDD_BASECOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_TAUXCOTISATION', 0.00);
              TOB_FilleLignes.PutValue('PDD_MTCOTISAT', Arrondi(MontST, 5));
              TOB_FilleLignes.PutValue('PDD_COMURBAINE', '    ');
              TOB_FilleLignes.PutValue('PDD_CONDITION', 'SCS');
              TOB_FilleLignes.PutValue('PDD_CODIFEDITEE',
                copy(CodifST, 3, 2));
            {RAZ Cumul Sous Total}
              MontST := 0.00;
              IndcalcSt := false;
            end; {if (Predef <> '')}
          end; {if ((PogPeriod1 = PeriodEtat)}
        end; {if (PogNature = '700'}
      end; {if (MontSt <> 0)}
    end; {if (PogBoolST = TRUE)}
    {fin traitement dernière ligne détail}

    if (Rupture = TRUE) then
    begin
      TRup := TRupMere.FindNext([''], [''], TRUE);
      if (TRup <> nil) then
      begin
        IndRup := False;
        if (TRup.GetValue('ET_SIRET') <> Null) then
          SiretEnCours := TRup.GetValue('ET_SIRET');
        if (TRup.GetValue('ET_APE') <> Null) then
          ApeEnCours := TRup.GetValue('ET_APE');
        EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
        OrganismeEnCours := TRup.GetValue('POG_ORGANISME');
        if (TRup.GetValue('POG_NUMINTERNE') <> Null) then
          NumIntEnCours := copy(TRup.GetValue('POG_NUMINTERNE'), 1, 30);
        if (TRup.GetValue('POG_LGOPTIQUE') <> Null) then
          LGOptiqueEncours := TRup.GetValue('POG_LGOPTIQUE');
        if (RupGpInt = TRUE) then
        begin
          if (TRup.GetValue('PGI_GROUPE') <> Null) then
            GpIntEnCours := TRup.GetValue('PGI_GROUPE');
        end;
      end;
    end;
  { traitement des ruptures }
    NumDucs := NumDucs + 1;
    ligne := 0;
    NoAT := -1;

  end; {fin while (fctRupture = true)}

  if (Anomalie = '') then
  begin
  {MISE A JOUR DES TABLES DUCSENTETE ET DUCSDETAIL}
  {-----------------------------------------------}
    try
      begintrans;
      if TOB_Lignes <> nil then
        TOB_Lignes.SetAllModifie(TRUE);
      StrDelDetail := 'DELETE FROM DUCSDETAIL ';
      if (Rupture = FALSE) then
        StrDelDetail := StrDelDetail + 'WHERE ' +
          '(PDD_ETABLISSEMENT ="' + etab + '") AND ' +
          '(PDD_ORGANISME ="' + organisme + '") AND ' +
          '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
          '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")'
      else
        StrDelDetail := StrDelDetail + StWhereDet + ')';
      StWhereDet := 'WHERE (';

      ExecuteSQL(StrDelDetail);

      StrDelTete := 'DELETE FROM DUCSENTETE ';
      if (Rupture = FALSE) then
        StrDelTete := StrDelTete + 'WHERE ' +
          '(PDU_ETABLISSEMENT ="' + etab + '") AND ' +
          '(PDU_ORGANISME ="' + organisme + '") AND ' +
          '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
          '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")'
      else
        StrDelTete := StrDelTete + StWhere + ')';
      StWhere := 'WHERE (';
      ExecuteSQL(StrDelTete);

      TOB_Lignes.InsertDB(nil, FALSE);
      Committrans;
    except
      Rollback;
// d PT65 maj jnal des événements
      PGIBox('Etablissement '+ Etab +'! Erreur maj table DUCSDETAIL', Ecran.Caption);
      TraceE.Add('Etablissement '+ Etab +'! Erreur maj table DUCSDETAIL');
// f PT65
    end;
  end
  else
  begin
// d PT65 maj jnal des événements
    if (Anomalie = 'Non Renseigné') then
    begin
      MessageAlerte('Le Siret établissement '+Etab+' n''est pas renseigné. Le traitement pour cet établissement est abandonné');
      TraceE.Add ('Le Siret de l''établissement '+Etab+' n''est pas renseigné.Le traitement pour cet établissement est abandonné');
    end
    else 
    if (Anomalie = 'erreur d''affectation') then
    begin
      MessageAlerte('Etablissement '+ Etab+' : Vérifier l''affectation des rubriques pour cet organisme');
      TraceE.Add ('Etablissement '+ Etab+' : Vérifier l''affectation des rubriques pour cet organisme');
    end;
// f PT65
  end;


  if TOB_Lignes <> nil then
  begin
    TOB_Lignes.Free;
  end;
end; { fin MajDucs}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 24/09/2001
Modifié le ... :   /  /
Description .. : Création des TOB filles (lignes détail) + alimentation de
Suite ........ : champs, commune quelque soit le type de ligne.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.AlimDetail(var TOB_FilleLignes, TOB_FilleTete: TOB; var ligne: Integer; Regroupt: string);
begin

  TOB_FilleLignes := TOB.create('DUCSDETAIL', TOB_FilleTete, -1);
  ligne := ligne + 1;
  TOB_FilleLignes.PutValue('PDD_ETABLISSEMENT', EtabEnCours);
// deb PT15-1
  if (Regroupt <> '') then
    TOB_FilleLignes.PutValue('PDD_ORGANISME', Organisme)
  else
    TOB_FilleLignes.PutValue('PDD_ORGANISME', OrganismeEnCours);

  TOB_FilleLignes.PutValue('PDD_DATEDEBUT', StrToDate(DebPer));
  TOB_FilleLignes.PutValue('PDD_DATEFIN', StrToDate(FinPer));
  TOB_FilleLignes.PutValue('PDD_NUM', NumDucs);

  if (Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_SIRET', SiretEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_SIRET', EtSiret);

  if (Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_APE', ApeEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_Ape', EtApe);

  if (RupGpInt = TRUE) then
    TOB_FilleLignes.PutValue('PDD_GROUPE', GpIntEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_GROUPE', '');

  if (Rupture = TRUE) then
    TOB_FilleLignes.PutValue('PDD_NUMERO', NumIntEnCours)
  else
    TOB_FilleLignes.PutValue('PDD_NUMERO', PogNumInterne);

  TOB_FilleLignes.PutValue('PDD_NUMORDRE', ligne);
// d PT63
  if (AlsaceMoselle) then
    TOB_FilleLignes.PutValue('PDD_REGIMEALSACE', 'X')
  else
    TOB_FilleLignes.PutValue('PDD_REGIMEALSACE', '-')
// f PT63
end; { fin AlimDetail }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 03/01/2002
Modifié le ... : 03/01/2002
Description .. : Contrôle sur les dates de période modifiées
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDUCSINIT.Change(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED')) then
  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de début est erronée.', Ecran.caption);
    SetControlText('XX_VARIABLED', DatetoStr(Date));

  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_')) then

  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de fin est erronée.', Ecran.caption);
    SetControlText('XX_VARIABLED_', DatetoStr(Date));
  end;
end; {fin Change}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 03/01/2002
Modifié le ... : 03/01/2002
Description .. : Contrôle des dates de période sur action d BCherche
Mots clefs ... :
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ChercheClick(Sender: TObject);
begin
  if not IsValidDate(GetControlText('XX_VARIABLED')) then
  {Pour générer message erreur si date erronnée }
  begin
    PGIBox('La date de début est erronée.', Ecran.caption);
    SetControlText('XX_VARIABLED', DatetoStr(Date));
  end;
  if not IsValidDate(GetControlText('XX_VARIABLED_')) then
  {Pour générer message erreur si date erronnée}
  begin
    PGIBox('La date de fin est erronée.', Ecran.caption);
    SetControlText('XX_VARIABLED_', DatetoStr(Date));
  end;
  TFMul(Ecran).BChercheClick(nil);
end; {fin ChercheClick}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 08/01/2002
Modifié le ... : 08/01/2002
Description .. : Récupération du déclarant par défaut (émetteur)
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ChangeEmetteur(Sender: TObject);
var
  Q: TQuery;
begin
  Declarant := GetControlText('EMETSOC2');
  Q := OpenSQL('SELECT PET_EMETTSOC,PET_RAISONSOC,PET_CONTACTDUCS,PET_TELDUCS,PET_FAXDUCS ' +
    'FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Declarant + '"', True);
  if not Q.EOF then
  begin
    Emetteur := Q.FindField('PET_EMETTSOC').AsString;
    RaisonSoc := copy(Q.FindField('PET_RAISONSOC').AsString, 1, 35);
    ContactDecla := Q.FindField('PET_CONTACTDUCS').AsString;
    TelDeclarant := Q.FindField('PET_TELDUCS').AsString;
    FaxDeclarant := Q.FindField('PET_FAXDUCS').AsString;
  end;
  ferme(Q);
end; { fin ChangeEmetteur}

{ Traitement des ruptures }

function TOF_PGMULDUCSINIT.fctRupture(TDetail: TOB; Ind: integer): Boolean;

begin
  Siretlu := '';
  Apelu := '';
  NumIntlu := '';
  GpIntlu := '';

  Result := TRUE;
  IndNeant2 := FALSE;

  if (Ind = 1) then
  {fin des lignes détail}
  begin
    WSiret := 'findetail';
  end
  else
    if (TDetail = nil) and (WSiret <> 'findetail') then
    {Pas de ligne détail - ducs néant}
    begin
      IndNeant2 := TRUE;
      exit;
    end;

  if (WSiret = 'findetail') then
  {fin du traitement}
  begin
    Result := FALSE;
    IndNeant2 := FALSE;
    exit;
  end
  else
  begin
    if (RupSiret = TRUE) then
      if (TDetail.GetValue('SIRET') <> Null) then
        Siretlu := TDetail.GetValue('SIRET');
    if (RupApe = TRUE) then
      if (TDetail.GetValue('APE') <> Null) then
        Apelu := TDetail.GetValue('APE');
    if (RupNumInt = TRUE) then
      if (TDetail.GetValue('NUMINTERNE') <> Null) then
        NumIntlu := TDetail.GetValue('NUMINTERNE');
    if (RupGpInt = TRUE) then
      if (TDetail.GetValue('GROUPE') <> Null) then
        GpIntlu := TDetail.GetValue('GROUPE');

    {Y-a-t'il rupture de Siret ?}
    if (RupSiret = TRUE) and (WSiret <> Siretlu) then
    begin
      Result := FALSE;
      WSiret := Siretlu;
    end;

    if (RupSiret = TRUE) then
      if (WSiret <> SiretEnCours) then
      {il n'y a aucune ligne pour le Siret en cours : c'est une ducs néant}
        IndNeant2 := TRUE
      else
        IndNeant2 := FALSE;

    {Y-a-t'il rupture de code Ape}
    if (RupApe = TRUE) and (WApe <> Apelu) then
    begin
      Result := FALSE;
      WApe := Apelu;
    end;

    if (RupApe = TRUE) then
      if (WApe <> ApeEnCours) then
      {il n'y a aucune ligne pour le code Ape en cours : c'est une ducs néant}
        IndNeant2 := TRUE
      else
        IndNeant2 := FALSE;

    {Y-a-t'il rupture de n° interne}
    if (RupNumInt = TRUE) and (WNumInt <> NumIntlu) then
    begin
      Result := FALSE;
      WNumInt := NumIntlu;
    end;

    if (RupNumInt = TRUE) and (IndNeant2 = FALSE) then
      if (WNumInt <> NumIntEnCours) then
      {il n'y a aucune ligne pour le n° interne en cours : c'est une ducs néant}
        IndNeant2 := TRUE
      else
        IndNeant2 := FALSE;

    if (RupGpInt = TRUE) and (WGpInt <> GpIntlu) then
    begin
      Result := FALSE;
      WGpInt := GpIntlu;
    end;
    if (RupGpInt = TRUE) and (IndNeant2 = FALSE) then
      if (WGpInt <> GpIntEnCours) then
      {il n'y a aucune ligne pour le groupe interne en cours : c'est une ducs néant}
        IndNeant2 := TRUE
      else
        IndNeant2 := FALSE;
  end;

  if (WSiret = '') then WSiret := Siretlu;
  if (WApe = '') then WApe := Apelu;
  if (WNumInt = '') then WNumInt := NumIntlu;
  if (WGpInt = '') then WGpInt := GpIntlu;
end; { fin fctRupture }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 13/02/2002
Modifié le ... : 13/02/2002
Description .. :
Suite ........ : Création des différentes TOB utilisées pour gérer
Suite ........ : les ruptures
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.RechRupture(Organisme: string; DucsDoss: boolean; Regroupt: string); // PT15-1
var
  st, Organ: string;
  QRech: TQuery;
  TobExiste, LesInst: TOB;
  NbInst, i: Integer;
// d PT61
//  NomChamp: array[1..3] of string;
//  ValeurChamp: array[1..3] of variant;
{*
  NomChamp: array[1..4] of string;
  ValeurChamp: array[1..4] of variant;
  Siret     : string;
  CreateTRupFille : boolean;  *}
  NomChamp: array[1..3] of string; // PT61
  ValeurChamp: array[1..3] of variant; //  PT61
// f PT61

begin
  st := ' ';
  {Création de la TOB utilisée en fonction du paramètrage des ruptures}
  case NoMere of
    1: begin
         { TOB - groupe interne}
        st := 'SELECT distinct(PGI_GROUPE),PGI_INSTITUTION, PSA_DADSPROF, ' +
          'ET_ETABLISSEMENT, ET_SIRET, ET_APE, POG_ORGANISME, ' +
          'POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+ // PT61
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+ //PT63
          'FROM HISTOBULLETIN ' +
          'LEFT JOIN ETABLISS ON ET_ETABLISSEMENT=PHB_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=PHB_ETABLISSEMENT ' +  //PT63
          'LEFT JOIN SALARIES ON PSA_SALARIE=HISTOBULLETIN.PHB_SALARIE ' +
          'LEFT JOIN GROUPEINTERNE ON PGI_CATEGORIECRC=SALARIES.PSA_DADSPROF ' +
          'LEFT JOIN ORGANISMEPAIE ON POG_INSTITUTION=PGI_INSTITUTION ' +
          'AND POG_ORGANISME=PHB_ORGANISME AND POG_ETABLISSEMENT=PHB_ETABLISSEMENT ' +
          'WHERE PSA_DADSPROF<>"" AND ';

        if (Regroupt <> '') then { Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR '
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME="' + Organisme + '" ';

        st := st + ' AND POG_RUPTGROUPE= "X"';
        if (DucsDoss = TRUE) then
          st := st + ' AND POG_DUCSDOSSIER="X" '
        else
          st := st + ' AND ET_ETABLISSEMENT="' + Etab + '"';

        st := st + ' ORDER BY ';
         {order by dans le même ordre que pour TCot2}
        if (RupSiret = True) then
          st := st + 'ET_SIRET,';
        if (RupApe = True) then
          st := st + 'ET_APE,';
        if (RupNumInt = true) then
          st := st + 'POG_NUMINTERNE,';
        if (RupSiret = False) then
          st := st + ' PGI_GROUPE, ET_SIRET'
        else
          st := st + ' PGI_GROUPE';

        QRech := OpenSql(st, TRUE);
      TRupMere := TOB.Create('Les groupes internes', nil, -1);
      end;

    2: begin
         { TOB - N° interne}
        st := 'SELECT DISTINCT(POG_NUMINTERNE), ET_ETABLISSEMENT, ET_SIRET, ' +
          'ET_APE, POG_ORGANISME,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+ //
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+ //PT63
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = ET_ETABLISSEMENT '+ //PT63
          'WHERE ';

        if (Regroupt <> '') then
         { Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR '
            end
            else
              st := st + '(';
            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME="' + Organisme + '" ';

        st := st + ' AND POG_RUPTNUMERO= "X" AND POG_NUMINTERNE <> ""';

        if (DucsDoss = TRUE) then
          st := st + ' AND POG_DUCSDOSSIER="X" '
        else
          st := st + ' AND POG_ETABLISSEMENT="' + Etab + '"';

        if (RupSiret = True) then
          st := st + 'ET_SIRET,';
        if (RupApe = True) then
          st := st + 'ET_APE,';

        if (RupSiret = False) then
          st := st + ' ORDER BY POG_NUMINTERNE, ET_SIRET'
        else
          st := st + ' ORDER BY POG_NUMINTERNE';

        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les numéros internes', nil, -1);
      end;

    3: begin
         {TOB - Code Ape}
        st := 'SELECT DISTINCT(ET_APE), ET_ETABLISSEMENT, ET_SIRET, POG_ORGANISME, POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+ // PT61
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+ //PT63
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT = ET_ETABLISSEMENT  ' +  //PT63
          ' WHERE ';

        if (Regroupt <> '') then
         {Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := TInst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');

            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR ';
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME ="' + Organisme + '"';

        st := st + ' AND POG_DUCSDOSSIER="X" AND POG_RUPTAPE= "X" ' +
          ' ORDER BY ET_APE';
        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les ape', nil, -1);
      end;

    4: begin
         {TOB - N° Siret}
        st := 'SELECT DISTINCT(ET_SIRET),ET_ETABLISSEMENT, ET_APE, ' +
          'POG_ORGANISME, POG_NUMINTERNE,' +
          'POG_REGROUPEMENT,POG_INSTITUTION, ' +
          'POG_CAISSEDESTIN, '+ // PT61
          'POG_LGOPTIQUE,POG_NATUREORG ' +
          ',ETB_REGIMEALSACE '+ //PT63
          'FROM ETABLISS ' +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON ET_ETABLISSEMENT=ORGANISMEPAIE.POG_ETABLISSEMENT ' +
          'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT=ET_ETABLISSEMENT ' +
          'WHERE '; // PT15-1

// debut PT15-1
        if (Regroupt <> '') then
         {Ducs regroupement d'organisme}
        begin
          for NbInst := 0 to TInst.Detail.Count - 1 do
          begin
            LesInst := Tinst.Detail[NbInst];
            Organ := LesInst.GetValue('POG_ORGANISME');
            if (NbInst <> 0) then
            begin
              if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
                st := st + ' OR ';
            end
            else
              st := st + '(';

            if (Pos('POG_ORGANISME="' + Organ + '"', st) = 0) then
              st := st + 'POG_ORGANISME="' + Organ + '"';
          end;
          st := st + ')';
        end
        else
          st := st + 'POG_ORGANISME ="' + Organisme + '"';

        st := st + ' AND POG_DUCSDOSSIER="X" AND POG_RUPTSIRET= "X"' +
          ' ORDER BY ET_SIRET';

        QRech := OpenSql(st, TRUE);
        TRupMere := TOB.Create('Les siret', nil, -1);
// fin PT15-1
      end;
  end;

  {Recherche si élément pas déjà en TOB}
  NomChamp[1] := '';
  ValeurChamp[1] := '';
  NomChamp[2] := '';
  ValeurChamp[2] := '';
  NomChamp[3] := '';
  ValeurChamp[3] := '';
// d PT61
{*
  NomChamp[4] := '';
  ValeurChamp[4] := '' ; *}
// f PT61

  while not QRech.EOF do
  begin
    if (RupGpInt = TRUE) then
    begin
      NomChamp[1] := 'PGI_GROUPE';
      ValeurChamp[1] := QRech.FindField('PGI_GROUPE').AsString;
// d PT61
{*
      if (DucsDoss = true) then
      begin
        NomChamp[4] := 'CAISSEDESTIN';
        if (QRech.FindField('POG_CAISSEDESTIN').AsString = 'X') then
          ValeurChamp[4] := '0'
        else
          ValeurChamp[4] := '1'
      end;*}
// f PT61

    end;
    if (RupNumInt = TRUE) then
    begin
      if (NomChamp[1] = '') then
      begin
        NomChamp[1] := 'POG_NUMINTERNE';
        ValeurChamp[1] := QRech.FindField('POG_NUMINTERNE').AsString;
      end
      else
      begin
        NomChamp[2] := 'POG_NUMINTERNE';
        ValeurChamp[2] := QRech.FindField('POG_NUMINTERNE').AsString;
      end;
    end;
    if (RupApe = TRUE) then
    begin
      if (NomChamp[2] = '') then
      begin
        if (NomChamp[1] = '') then
        begin
          NomChamp[1] := 'ET_APE';
          ValeurChamp[1] := QRech.FindField('ET_APE').AsString;
        end
        else
        begin
          NomChamp[2] := 'ET_APE';
          ValeurChamp[2] := QRech.FindField('ET_APE').AsString;
        end;
      end
      else
      begin
        NomChamp[3] := 'ET_APE';
        ValeurChamp[3] := QRech.FindField('ET_APE').AsString;
      end;
    end;

    if (RupSiret = TRUE) then
    begin
      if (NomChamp[2] = '') then
      begin
        if (NomChamp[1] = '') then
        begin
          NomChamp[1] := 'ET_SIRET';
          ValeurChamp[1] := QRech.FindField('ET_SIRET').AsString;
        end
        else
        begin
          NomChamp[2] := 'ET_SIRET';
          ValeurChamp[2] := QRech.FindField('ET_SIRET').AsString;
        end;
      end
      else
      begin
        NomChamp[3] := 'ET_SIRET';
        ValeurChamp[3] := QRech.FindField('ET_SIRET').AsString;
      end;
    end;

    TobExiste := TRupMere.FindFirst(NomChamp, ValeurChamp, TRUE);
    if (TobExiste = nil) then
    {Pas déjà en TOB - Création d'une TOB fille}
    begin
// d PT61
{*
      // on ne conserve que la caisse destinataire qd rupture sur groupe interne
      CreateTRupFille :=  true;
      if (DucsDoss = TRUE) and (not RupSiret) then
      begin
        if (ValeurChamp[4] = '0') then
        begin
          ValeurChamp[4] := '';
          NomChamp[4] := '';
          TobExiste := TRupMere.FindFirst(NomChamp, ValeurChamp, TRUE);
          if (TobExiste <> nil) then
            TobExiste.free;
        end;
        if (ValeurChamp[4] = '1') then
        begin
          ValeurChamp[4] := '';
          NomChamp[4] := '';
          TobExiste := TRupMere.FindFirst(NomChamp, ValeurChamp, TRUE);
          if (TobExiste <> nil) then
            CreateTRupFille := false;
        end;
      end;*}
// f PT61
// PT15-1        TRupFille := TOB.Create('ETABLISS', TRupMere, -1);
{* 
      if (CreateTRupFille) then
      begin*}
      TRupFille := TOB.Create('TABLEETABLISS', TRupMere, -1);
      TRupFille.InitValeurs;
      if (RupGpInt = true) then
      begin
        TRupFille.AddChampSup('PGI_GROUPE', FALSE);
        TRupFille.PutValue('PGI_GROUPE', QRech.FindField('PGI_GROUPE').AsString);
      end;
      TRupFille.AddChampSup('POG_NUMINTERNE', FALSE);
      TRupFille.PutValue('POG_NUMINTERNE', QRech.FindField('POG_NUMINTERNE').AsString);
      TRupFille.AddChampSup('ET_APE', FALSE);
      TRupFille.PutValue('ET_APE', QRech.FindField('ET_APE').AsString);
      TRupFille.AddChampSup('POG_ETABLISSEMENT', FALSE);
      TRupFille.PutValue('POG_ETABLISSEMENT', QRech.FindField('ET_ETABLISSEMENT').AsString);
      TRupFille.AddChampSup('ET_SIRET', FALSE);
      TRupFille.PutValue('ET_SIRET', QRech.FindField('ET_SIRET').AsString);
      TRupFille.AddChampSup('POG_ORGANISME', FALSE);
      TRupFille.PutValue('POG_ORGANISME', QRech.FindField('POG_ORGANISME').AsString);

      {Formatage ligne optique (elle peut être déjà renseignée au niveau
       du paramètrage de l'organisme}
      LigneOptique := '';
      if (QRech.FindField('POG_LGOPTIQUE').AsString = '') then
      begin
        if (QRech.FindField('POG_NATUREORG').AsString = '100') then
        {URSSAF}
        begin
          LigneOptique := Copy(QRech.FindField('POG_NUMINTERNE').AsString, 1,
            length(QRech.FindField('POG_NUMINTERNE').AsString));
          if (length(QRech.FindField('POG_NUMINTERNE').AsString) < 30) then
          begin
            for i := length(QRech.FindField('POG_NUMINTERNE').AsString) + 1 to 30 do
            begin
              LigneOptique := LigneOptique + '0';
            end;
          end;
        end;
        if (QRech.FindField('POG_NATUREORG').AsString = '200') then
        {ASSEDIC}
        begin
          LigneOptique := 'S2' + Copy(QRech.FindField('POG_NUMINTERNE').AsString, 1,
            length(QRech.FindField('POG_NUMINTERNE').AsString));
          if ((length(QRech.FindField('POG_NUMINTERNE').AsString) + 2) < 30) then
          begin
            for i := length(QRech.FindField('POG_NUMINTERNE').AsString) + 3 to 30 do
            begin
              LigneOptique := LigneOptique + '0';
            end;
          end;
        end;
      end
      else
      begin
        LigneOptique := QRech.FindField('POG_LGOPTIQUE').AsString;
        if (length(QRech.FindField('POG_LGOPTIQUE').AsString) < 30) then
        begin
          for i := length(QRech.FindField('POG_LGOPTIQUE').AsString) + 1 to 30 do
          begin
            LigneOptique := LigneOptique + '0';
          end;
        end;
      end;

      TRupFille.AddChampSup('POG_LGOPTIQUE', FALSE);
      TRupFille.PutValue('POG_LGOPTIQUE', LigneOptique);
// d PT61
//    Pour placer la caisse destinataire en début de TOB
{*
      if (DucsDoss = True) then
      begin
        TRupFille.AddChampSup('CAISSEDESTIN', FALSE);
        if (QRech.FindField('POG_CAISSEDESTIN').AsString = 'X')  then
          TRupFille.PutValue('CAISSEDESTIN', '0')
        else
          TRupFille.PutValue('CAISSEDESTIN', '1');
      end;
      end;*}
// f PT61
    end {fin if (TobExiste = NIL)}
// d PT62
    else
    begin
      // TobExiste <> NIL : élément déjà existant en TOB (cas pls Gps Internes avec
      // même code et ducs dossier) on récupère les infos de la caisse destinataire)
      if (QRech.FindField('POG_CAISSEDESTIN').AsString = 'X') then
      begin
        TRupFille.PutValue('POG_NUMINTERNE', QRech.FindField('POG_NUMINTERNE').AsString);
        TRupFille.PutValue('ET_APE', QRech.FindField('ET_APE').AsString);
        TRupFille.PutValue('POG_ETABLISSEMENT', QRech.FindField('ET_ETABLISSEMENT').AsString);
        TRupFille.PutValue('ET_SIRET', QRech.FindField('ET_SIRET').AsString);
        TRupFille.PutValue('POG_ORGANISME', QRech.FindField('POG_ORGANISME').AsString);
      end;
    end;
// f PT62
// d PT63
    TRupFille.AddChampSup('ETB_REGIMEALSACE', FALSE);
    TRupFille.PutValue('ETB_REGIMEALSACE', QRech.FindField('ETB_REGIMEALSACE').AsString);
// f PT63
    QRech.Next;

    NomChamp[1] := '';
    ValeurChamp[1] := '';
    NomChamp[2] := '';
    ValeurChamp[2] := '';
    NomChamp[3] := '';
    ValeurChamp[3] := '';
// d PT61
{*
    NomChamp[4] := '';
    ValeurChamp[4] := '';*}
// f PT61
  end; {fin while not QRech.EOF}


  Ferme(QRech);
// d PT61
{*
  if (DucsDoss = True) then
  begin
    NomChamp[1] := 'CAISSEDESTIN';
    ValeurChamp[1] := '0';

    TobExiste := TRupMere.FindFirst(NomChamp, ValeurChamp, TRUE);
    if (TobExiste <> nil) then
    begin
      Siret :=  TobExiste.GetValue('ET_SIRET');
      for i := 0 to TRupMere.Detail.Count-1 do
      begin
        TRupFille := TRupMere.Detail[i] ;
        TRupFille.PutValue('ET_SIRET',Siret );
      end;
    end;
  end;*}
// f PT61

end; { fin RechRupture}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Mise en forme des "Where" uitlisés pour l'édition
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.ChargeWheres(Regroupt: string);
begin
  {  Gestion des ruptures }
  if (rupture = TRUE) then
  begin
    TRup := TRupMere.FindFirst([''], [''], TRUE);
    while (TRup <> nil) do
    begin
      if (TRup.GetValue('ET_SIRET') <> Null) then
        SiretEnCours := TRup.GetValue('ET_SIRET');
      if (TRup.GetValue('POG_ETABLISSEMENT') <> Null) then
        EtabEnCours := TRup.GetValue('POG_ETABLISSEMENT');
      if (TRup.GetValue('ET_APE') <> Null) then
        ApeEnCours := TRup.GetValue('ET_APE');
// deb PT15-1
      if (Regroupt <> '') then
        OrganismeEnCours := organisme
      else
        OrganismeEnCours := TRup.GetValue('POG_ORGANISME');

      { en-tête }
      StWhere := StWhere +
        '((PDU_ETABLISSEMENT ="' + EtabEnCours + '") AND ' +
        '(PDU_ORGANISME ="' + organismeEnCours + '") AND ' +
        '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
        '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';
      StWhereEd := StWhereEd +
        '((PDU_ETABLISSEMENT ="' + EtabEnCours + '") AND ' +
        '(PDU_ORGANISME ="' + organismeEnCours + '") AND ' +
        '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
        '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

      { détail}
      StWhereDet := StWhereDet +
        '((PDD_ETABLISSEMENT ="' + EtabEnCours + '") AND ' +
        '(PDD_ORGANISME ="' + organismeEnCours + '") AND ' +
        '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
        '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

      TRup := TRupMere.FindNext([''], [''], TRUE);
      if (TRup <> nil) then
      {il y a d'autres Siret}
      begin
        StWhere := StWhere + 'OR ';
        StWhereEd := StWhereEd + 'OR ';
        StWhereDet := StWhereDet + 'OR ';
      end;
    end; { finwhile (TRup <> NIL) }
  end; {fin if (rupture = TRUE)}

  if (Rupture <> TRUE) then
  { Cas sans rupture}
  begin
    {en-tête}
    StWhere := StWhere +
      '((PDU_ETABLISSEMENT ="' + Etab + '") AND ' +
      '(PDU_ORGANISME ="' + organisme + '") AND ' +
      '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
      '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';
    StWhereEd := StWhereEd +
      '((PDU_ETABLISSEMENT ="' + Etab + '") AND ' +
      '(PDU_ORGANISME ="' + organisme + '") AND ' +
      '(PDU_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
      '(PDU_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';

    { détail}
    StWhereDet := StWhereDet +
      '((PDD_ETABLISSEMENT ="' + Etab + '") AND ' +
      '(PDD_ORGANISME ="' + organisme + '") AND ' +
      '(PDD_DATEDEBUT ="' + UsDateTime(StrToDate(DebPer)) + '") AND ' +
      '(PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '")) ';
  end; { fin if (Rupture <> TRUE) }
end; { fin ChargeWheres}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Libération des différentes TOB
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.LibereTobs();
begin
  if TInst <> nil then
  begin
    TInst.Free;
  end;

{PT62  if TCot2 <> nil then
  begin
    TCot2.Free;
  end;}
  FreeAndNil(TCot3); // PT62

  if (TRupMere <> nil) then
  begin
    TRupMere.Free;
    TRupMere := nil; // PT15 BIS -6
  end;
end; { fin LibereTobs}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 13/02/2002
Modifié le ... : 13/02/2002
Description .. : Calcul des effectifs
Suite ........ : Effectifs des salariés concernés par la déclaration:
Suite ........ : recherche sur HISTOBULLETIN
Suite ........ : Effectif des salariés présents en fin de période :
Suite ........ : recherche sur SALARIES
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

//PT84
//procedure TOF_PGMULDUCSINIT.CalculEffectifs(DucsDoss: Boolean; Regroupt: string);
procedure TOF_PGMULDUCSINIT.CalculEffectifs(DucsDoss: Boolean; Regroupt, PogNature: string);
var
  LesInst: TOB;
  NbInst: Integer;
  StQQ, StRupt, StJoinRupt, Organ, PogNumInt: string;
  QQ: Tquery;
  StJoin: string; 
{$IFNDEF DUCS41}
   TEffectifs , LeSal                        : TOB;
   I                                         : integer;
   DateNaiss                                 : TDateTime;
{$ENDIF}
   WEff                                      : double;  // PT69
begin
  STQQ := '';
  StJoin := ''; 

  {CALCUL DU NBRE DE SALARIES CONCERNES PAR LA DECLARATION}
  {(ayant des cotisations pour la période)                }
  {-------------------------------------------------------}
  if (DucsDoss = False) then
    { ducs établissement }
    StQQ := 'PHB_ETABLISSEMENT="' + Etab + '" AND '
  else
    { ducs dossier}
    if (RupNumInt = FALSE) and (RupGpInt = FALSE) then
    begin
    { On ne tient compte que des établissements pour lesquels "Ducs Dossier" est
      coché sur l'organisme traité
      dans ce cas là, le Left Join n'est pas fait par la suite }
      StJoin := 'LEFT JOIN ORGANISMEPAIE ON PHB_ETABLISSEMENT = POG_ETABLISSEMENT ';
      StQQ := 'POG_DUCSDOSSIER = "X" AND ';
    end;
  StQQ := StQQ + 'PHB_NATURERUB="COT" AND ';

  if (Regroupt <> '') then
  {Ducs regroupement d'organismes }
  begin
    for NbInst := 0 to TInst.Detail.Count - 1 do
    begin
      LesInst := TInst.Detail[NbInst];
      Organ := LesInst.GetValue('POG_ORGANISME');
      if (NbInst <> 0) then
      begin
        if (Pos('PHB_ORGANISME="' + Organ + '"', stQQ) = 0) then
          stQQ := stQQ + ' OR ';
      end
      else
        stQQ := stQQ + '(';

      if (Pos('PHB_ORGANISME="' + Organ + '"', stQQ) = 0) then
        stQQ := stQQ + 'PHB_ORGANISME="' + Organ + '"';
    end;
    stQQ := stQQ + ')';
  end
  else
  {Ducs sans regroupement d'organismes}
    StQQ := StQQ + 'PHB_ORGANISME="' + Organisme + '" ';

  StQQ := StQQ + 'AND ' +
    'PHB_DATEFIN >="' + UsDateTime(StrToDate(DebPer)) + '" AND ' +
    'PHB_DATEFIN <="' + UsDateTime(StrToDate(FinPer)) + '" AND ' +
    '(PHB_BASECOT <> 0 OR PHB_MTSALARIAL <> 0 OR PHB_MTPATRONAL <> 0)';

  if (Rupture = FALSE) then
  {PAS DE RUPTURE}
  begin
    QQ := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN ' + StJoin + 'WHERE ' + StQQ, True);
  end
  else
  {EXISTENCE DE RUPTURE }
  begin
    StRupt := '';
    StJoinRupt := '';
    if (RupSiret = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PHB_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_SIRET="' + SiretEnCours + '" ';
    end;
    if (RupApe = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PHB_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_APE="' + ApeEnCours + '" ';
    end;
    if (RupNumInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ORGANISMEPAIE ' +
        'ON PHB_ETABLISSEMENT=POG_ETABLISSEMENT AND ' +
        'PHB_ORGANISME=POG_ORGANISME';
      if (RupSiret = TRUE) or (RupApe = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';

      StRupt := StRupt + 'POG_NUMINTERNE="' + NumIntEnCours + '"';
    end;

    if (RupGpInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN SALARIES ' +
        'ON PHB_SALARIE=PSA_SALARIE ' +
        'LEFT JOIN GROUPEINTERNE ' +
        'ON PGI_CATEGORIECRC=PSA_DADSPROF ';
      if (RupNumInt = FALSE) then
        StJoinRupt := StJoinRupt +
          ' LEFT JOIN ORGANISMEPAIE ' +
          'ON PHB_ORGANISME=POG_ORGANISME ';
      if (RupSiret = TRUE) or (RupApe = TRUE) or (RupNumInt = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := ' AND ';

      StRupt := StRupt + 'PGI_GROUPE="' + GpIntEnCours + '" ';

      if (DucsDoss = TRUE) then
        StRupt := StRupt + 'AND POG_DUCSDOSSIER="X"';
    end;

    QQ := OpenSQL('SELECT Count(DISTINCT PHB_SALARIE) AS NOMBRE FROM HISTOBULLETIN ' +
      StJoin + ' ' + StJoinRupt + ' WHERE ' + StQQ + StRupt, True);
  end; { fin EXISTENCE DE RUPTURE}
  if not QQ.EOF then // PortageCWAS
    EffTot := QQ.FindField('NOMBRE').AsInteger;
  Ferme(QQ);

 {FIN -CALCUL DU NOMBRE DE SALARIES CONCERNES PAR LA DECLARATION}
  {--------------------------------------------------------------}

  {CALCUL DES EFFECTIFS PRESENTS EN FIN DE PERIODE}
  {-----------------------------------------------}
  StQQ := '';
  StJoin := '';

  if (DucsDoss = False) then
    { ducs établissement }
    StQQ := 'PSA_ETABLISSEMENT="' + Etab + '" AND '
  else
    { ducs dossier}
    if (RupNumInt = FALSE) and (RupGpInt = FALSE) then
    begin
    { On ne tient compte que des établissements pour lesquels "Ducs Dossier" est
      coché sur l'organisme traité
      dans ce cas là, le Left Join n'est pas fait par la suite }
      StJoin := 'LEFT JOIN ORGANISMEPAIE ON PSA_ETABLISSEMENT = POG_ETABLISSEMENT ';
      StQQ := 'POG_DUCSDOSSIER="X" AND POG_ORGANISME ="' + Organisme + '" AND';
    end;

  StQQ := StQQ + '(PSA_DATESORTIE="' + UsDateTime(IDate1900) + '"' +
    ' OR  PSA_DATESORTIE IS NULL ' +
    ' OR PSA_DATESORTIE>="' + UsDateTime(StrToDate(FinPer)) + '")' +
    ' AND PSA_DATEENTREE<="' + UsDateTime(StrToDate(FinPer)) + '"';

// d ducs edi v4.2
  if (Rupture = FALSE) then
  {PAS DE RUPTURE}
  begin
{$IFDEF DUCS41}
    {nbre d'hommes}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND PSA_SEXE="M"', True);
    if not QQ.EOF then
      EffHom := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {nbre de femmes}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND PSA_SEXE="F"', True);
    if not QQ.EOF then
      EffFem := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {Nbre d'apprentis}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND PSA_CATDADS="003"', True);
    if not QQ.EOF then
      EffApp := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {Nbre de cadres  (cadre ou dirigeant)}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND (PSA_CATDADS="001" OR PSA_CATDADS="002")', True);
    if not QQ.EOF then
      EffCad := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    {Nbre d'ouvriers}

    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND PSA_CATDADS="004"', True);
    if not QQ.EOF then
      EffOuv := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {nbre d'Etam }
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoin + ' WHERE ' + StQQ +
      ' AND PSA_CATDADS="005"', True);
    if not QQ.EOF then
      EffEtam := QQ.Fields[0].AsInteger;
    Ferme(QQ);
{$ELSE}
    StRupt := '' ;
    StJoinRupt := StJoin;
{$ENDIF}
  end {FIN PAS DE RUPTURE}
  else
  {EXISTENCE DE RUPTURE }
  begin
    StRupt := '';
    StJoinRupt := '';
    if (RupSiret = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_SIRET="' + SiretEnCours + '" ';
    end;
    if (RupApe = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ETABLISS ON PSA_ETABLISSEMENT=ET_ETABLISSEMENT ';
      StRupt := 'AND ET_APE="' + ApeEnCours + '" ';
    end;
    if (RupNumInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN ORGANISMEPAIE ' +
        'ON PSA_ETABLISSEMENT=POG_ETABLISSEMENT AND ';
      if (RupSiret = TRUE) or (RupApe = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';
      if (Regroupt <> '') then
      {Ducs regroupement d'organismes}
      begin
        for NbInst := 0 to Tinst.Detail.Count - 1 do
        begin
          LesInst := Tinst.Detail[NbInst];
          if (LesInst.GetValue('POG_NUMINTERNE') <> Null) then
            PogNumInt := LesInst.GetValue('POG_NUMINTERNE');
          if (NbInst <> 0) then
            StRupt := StRupt + ' OR '
          else
            StRupt := StRupt + '(';
          StRupt := StRupt + 'POG_NUMINTERNE="' + PogNumInt + '"';
        end;
        StRupt := StRupt + ')';
      end
      else
        StRupt := StRupt + 'POG_NUMINTERNE="' + NumIntEnCours + '"';
    end;

    if (RupGpInt = TRUE) then
    begin
      StJoinRupt := StJoinRupt +
        ' LEFT JOIN GROUPEINTERNE ' +
        'ON PSA_DADSPROF=PGI_CATEGORIECRC ';
      if (RupNumInt = False) then
        StJoinRupt := StJoinRupt +
          'LEFT JOIN ORGANISMEPAIE ' +
          'ON PSA_ETABLISSEMENT = POG_ETABLISSEMENT ';
      if (RupSiret = TRUE) or (RupApe = TRUE) or (RupNumInt = TRUE) then
        StRupt := StRupt + ' AND '
      else
        StRupt := 'AND ';

      StRupt := StRupt + 'PGI_GROUPE="' + GpIntEnCours + '"  ' +
        'AND POG_ORGANISME="' + OrganismeEnCours + '" ';

      if (DucsDoss = TRUE) then
        StRupt := StRupt + 'AND POG_DUCSDOSSIER="X"';
    end;
    StJoinRupt := StJoin + StJoinRupt;
{$IFDEF DUCS41}
    {nbre d'hommes}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_SEXE="M"', True);
    if not QQ.EOF then
      EffHom := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {nbre de femmes}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_SEXE="F"', True);
    if not QQ.EOF then
      EffFem := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {Nbre d'apprentis}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_CATDADS="003"', True);
    if not QQ.EOF then
      EffApp := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {Nbre de cadres  (cadre ou dirigeant)}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_CATDADS="001" OR PSA_CATDADS="002"', True);
    if not QQ.EOF then
      EffCad := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    {Nbre d'ouvriers}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_CATDADS="004"', True);
    if not QQ.EOF then
      EffOuv := QQ.Fields[0].AsInteger;
    Ferme(QQ);

    { nbre d'Etam}
    QQ := OpenSQL('SELECT Count(*) FROM SALARIES ' + StJoinRupt + ' WHERE ' + StQQ + StRupt +
      ' AND PSA_CATDADS="005"', True);
    if not QQ.EOF then
      EffEtam := QQ.Fields[0].AsInteger;
    Ferme(QQ);
{$ENDIF}
  end;{FIN EXISTENCE DE RUPTURE}

{$IFNDEF DUCS41}
   StQQ := StQQ +' and ((pci_debutcontrat <= "' + UsDateTime(StrToDate(FinPer)) + '" and '+
                 'pci_fincontrat >= "' + UsDateTime(StrToDate(DebPer)) + '" or '+
                 'pci_fincontrat = "' + UsDateTime(IDate1900) + '") or '+
                 '(pci_typecontrat is  NULL) or (psa_salarie not in '+
                 '(select pci_Salarie from contrattravail where '+
                 '(pci_debutcontrat <= "' + UsDateTime(StrToDate(FinPer)) + '" and '+
                 'pci_fincontrat >= "' + UsDateTime(StrToDate(DebPer)) + '" or '+
                 'pci_fincontrat = "' + UsDateTime(IDate1900) + '"))))';
// d PT84
 if PogNature <> '600' then
  StQQ := StQQ + ' and psa_salarie not in (select pse_salarie from deportsal where '+
                 'pse_msa="X")'
 else
   StQQ := StQQ + ' and psa_salarie in (select pse_salarie from deportsal where '+
                 'pse_msa="X")';
// f PT84

//PT69  QQ := OpenSQL('SELECT PSA_SALARIE,PSA_SEXE,PSA_CATDADS,PSA_DATENAISSANCE,'+
  QQ := OpenSQL('SELECT DISTINCT (PSA_SALARIE),PSA_SEXE,PSA_CATDADS,PSA_DATENAISSANCE,'+
                'PSA_PRISEFFECTIF, PSA_UNITEPRISEFF,'+ //PT69
                'PSA_DADSPROF,PSA_PROFIL, PCI_TYPECONTRAT FROM SALARIES '+
                  StJoinRupt+
                  ' left join contrattravail on psa_salarie = pci_salarie WHERE '+
                  StQQ+StRupt,True) ;
  TEffectifs := TOB.Create('Les effectifs', NIL, -1);
  TEffectifs.LoadDetailDB('SAL','','',QQ,False);
    Ferme(QQ);

  if (TEffectifs.Detail.Count <> 0) then
  begin
    DateNaiss := PlusMois(StrToDate(FinPer),-780);

    for I := 0 to TEffectifs.Detail.Count-1 do
    begin
      LeSal := TEffectifs.Detail[I];
      if (LeSal.GetValue('PSA_DATENAISSANCE') <=  DateNaiss) then
        LeSal.AddChampSupValeur('+65ANS','X')
      else
         LeSal.AddChampSupValeur('+65ANS','-');
// d PT62
    QQ := OpenSQL('SELECT PPS_PROFIL FROM PROFILSPECIAUX WHERE PPS_CODE ="'+
                   LeSal.GetValue('PSA_SALARIE')+'"',True);
    if QQ.Eof then
      LeSal.AddChampSupValeur('PROFILSPECIAL','')
    else
      LeSal.AddChampSupValeur('PROFILSPECIAL',QQ.Fields[0].AsString);
    Ferme(QQ);
// f PT62
    end;
// d PT69
    if (PogTypBordereau = '915') or (PogTypBordereau = '916') then
    // Pour le TR URSSAF un salarié à temps partiel est pris en compte au
    // prorata de son temps de travail (PSA_UNITEPRISEFF)
    begin
      {nbre d'hommes}
      WEff := TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_SEXE','PSA_PRISEFFECTIF'],['M','X'],TRUE, TRUE);
      EffHom := Int(Weff);
      if (EffHom <> Weff) then EffHom := EffHom+1;
      {nbre de femmes}
      WEff := TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_SEXE','PSA_PRISEFFECTIF'],['F','X'],TRUE, FALSE);
      EffFem := Int(Weff);
      if (EffFem <> Weff) then EffFem := EffFem+1;
      {Nbre d'apprentis}
      Weff:= TEffectifs.Somme('PSA_UNITEPRISEFF',['PSA_CATDADS','PSA_PRISEFFECTIF'],['003','X'],TRUE, TRUE);
      EffApp := Int(Weff);
      if (EffApp <> Weff) then EffApp := EffApp+1;
    end
    else
    begin
      {nbre d'hommes}
      EffHom := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PRISEFFECTIF'],['M','X'],TRUE, TRUE);
      {nbre de femmes}
      EffFem := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PRISEFFECTIF'],['F','X'],TRUE, TRUE);
      {Nbre d'apprentis}
      EffApp:= TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['003','X'],TRUE, TRUE);
    end;
    {Nbre d'apprentis hommes}
    EffAppH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS','PSA_PRISEFFECTIF'],['M','003','X'],TRUE, TRUE);
    {Nbre d'apprentis femmes}
    EffAppF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_CATDADS','PSA_PRISEFFECTIF'],['F','003','X'],TRUE, TRUE);
    {Nbre de cadres  (cadre ou dirigeant)}
    EffCad := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['001','X'],TRUE, TRUE)+
              TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['002','X'],TRUE, TRUE);
    {Nbre de cadres hommes (cadre ou dirigeant)}
    EffCadH := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','22','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','23','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','24','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['M','25','X'],TRUE, TRUE);
    {Nbre de cadres femmes (cadre ou dirigeant)}
    EFFCadF := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','22','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','23','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','24','X'],TRUE, TRUE)+
               TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_DADSPROF','PSA_PRISEFFECTIF'],['F','25','X'],TRUE, TRUE);
    {Nbre d'ouvriers}
//PT83
//  EffOuv := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['004'],TRUE, TRUE);
    EffOuv := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['004','X'],TRUE, TRUE);

    {nbre d'Etam }
    EffEtam := TEffectifs.Somme('PSA_CATDADS',['PSA_CATDADS','PSA_PRISEFFECTIF'],['005','X'],TRUE, TRUE);

    {nbre salariés hommes de + de 65 ans}
    Eff65H := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS','PSA_PRISEFFECTIF'],['M','X','X'],TRUE, TRUE);

    {nbre salariés femmes de + de 65 ans}
    Eff65F := TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','+65ANS','PSA_PRISEFFECTIF'],['F','X','X'],TRUE, TRUE);

    {Nbre d'hommes en contrat de professionnalisation}
//PT62    EffCesH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['M','023'],TRUE, TRUE);
    EffProfH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PROFILSPECIAL','PSA_PRISEFFECTIF'],['M','296','X'],TRUE, TRUE);

    {Nbre de femmes titulaires d'un CES}
//PT62    EffCesF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PSA_PROFIL'],['F','023'],TRUE, TRUE);
    EffProfF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PROFILSPECIAL','PSA_PRISEFFECTIF'],['F','296','X'],TRUE, TRUE);

    {Nbre d'hommes en CDI}
//  PT62  EffCDIH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['M','CDI'],TRUE, TRUE);
    EffCDIH:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','X'],
                               TRUE, TRUE);
    EffCadH22:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','22','X'],
                               TRUE, TRUE);
    EffCadH23:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','23','X'],
                               TRUE, TRUE);
    EffCadH24:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','24','X'],
                               TRUE, TRUE);
    EffCadH25:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['M','CDI','-','25','X'],
                               TRUE, TRUE);
    EffCDIH := EffCDIH - EffCadH22 - EffCadH23 - EffCadH24 - EffCadH25;

    {Nbre de femmes en CDI}
//PT62   EffCDIF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT'],['F','CDI'],TRUE, TRUE);
    EffCDIF:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','X'],
                               TRUE, TRUE);
    EffCadF22:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','22','X'],
                               TRUE, TRUE);
    EffCadF23:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','23','X'],
                               TRUE, TRUE);
    EffCadF24:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','24','X'],
                               TRUE, TRUE);
    EffCadF25:= TEffectifs.Somme('PSA_SEXE',
                               ['PSA_SEXE','PCI_TYPECONTRAT','+65ANS','PSA_DADSPROF','PSA_PRISEFFECTIF'],
                               ['F','CDI','-','25','X'],
                               TRUE, TRUE);
    EffCDIF := EffCDIF - EffCadF22 - EffCadF23 - EffCadF24 - EffCadF25;

    {Nbre d'hommes en CDD}
    EffCDDH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['M','CCD','X'],TRUE, TRUE);

    {Nbre de femmes en CDD}
    EffCDDF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['F','CCD','X'],TRUE, TRUE);

    {Nbre d'hommes en CNE}
    EffCNEH:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['M','CNE','X'],TRUE, TRUE);

    {Nbre de femmes en CNE}
    EffCNEF:= TEffectifs.Somme('PSA_SEXE',['PSA_SEXE','PCI_TYPECONTRAT','PSA_PRISEFFECTIF'],['F','CNE','X'],TRUE, TRUE);
// f PT69
  end;
  FreeAndNIl(Teffectifs);
{$ENDIF}
// f ducs edi v4.2
  { FIN - CALCUL DES EFFECTIFS PRESENTS EN FIN DE PERIODE}
end; { fin Calculer effectifs}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Edition des états DUCS (Néant ou non néant)
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure TOF_PGMULDUCSINIT.EditerDucs(Pages: TPageControl);
var
  StWhereBis, StSQL: string;
begin
  StSQL := '';
  if (IndNeant = TRUE) then
  begin
  {DUCS NEANT}
    StWhereBis := 'AND ((PDU_PAIEMENT = "' + UsDateTime(IDate1900) + '") AND ' +
      '(PDU_NBSALFPE = 0) ' +
      'AND (PDU_NBSALHAP = 0))';

    StSQL := 'SELECT DUCSENTETE.*,' +
      'ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ' +
      'ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_FAX, ' +
      'POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_NATUREORG, ' +
      'POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, ' +
      'POG_VILLE , POG_LONGEDITABLE, ' +
      'POG_PERIODICITDUCS, POG_AUTREPERIODUCS, POG_PERIODCALCUL, ' +
      'POG_AUTPERCALCUL, ' +
      'POG_POSTOTAL,POG_LONGTOTAL ' +
      'FROM DUCSENTETE ' +
      'LEFT JOIN ETABLISS ON PDU_ETABLISSEMENT = ET_ETABLISSEMENT ' +
      'LEFT JOIN ORGANISMEPAIE ON PDU_ETABLISSEMENT = POG_ETABLISSEMENT ' +
      'AND PDU_ORGANISME = POG_ORGANISME ' +
      StWhereEd + ') ' + StWhereBis;

    FunctPGDebEditDucs(); {raz n° page}

    LanceEtat('E', 'PDU', 'DUN', True, False, False, Pages, StSQL, '', False);

    FunctPGFinEditDucs(); {raz AncNumPagSys}
  end;

  if (IndDucs = TRUE) then
  begin
  {DUCS NON NEANT}
    StWhereBis := 'AND ((PDU_PAIEMENT <> "' + UsDateTime(IDate1900) + '") OR ' +
      '(PDU_NBSALFPE <> 0) ' +
      'OR (PDU_NBSALHAP <> 0))';

    StSQL := 'SELECT DUCSENTETE.*,PDD_CODIFICATION,PDD_CODIFEDITEE, ' +
      'PDD_LIBELLE, PDD_BASECOTISATION,PDD_TAUXCOTISATION, PDD_MTCOTISAT, ' +
      'PDD_EFFECTIF, PDD_TYPECOTISATION, PDD_INSTITUTION, ' +
      'PDD_LIBELLESUITE, ' +
      'PDD_REGIMEALSACE, '+ //PT63
      'ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ' +
      'ET_CODEPOSTAL, ET_VILLE, ET_TELEPHONE, ET_FAX, ' +
      'POG_ETABLISSEMENT, POG_ORGANISME, POG_LIBELLE, POG_NATUREORG, ' +
      'POG_ADRESSE1, POG_ADRESSE2, POG_ADRESSE3, POG_CODEPOSTAL, ' +
      'POG_PERIODICITDUCS, POG_AUTREPERIODUCS, POG_PERIODCALCUL, ' +
      'POG_AUTPERCALCUL, ' +
      'POG_POSTOTAL,POG_LONGTOTAL, ' +
      'POG_VILLE , POG_LONGEDITABLE,POG_BASETYPARR,' +
      'POG_MTTYPARR ' +
      'FROM DUCSENTETE ' +
      'LEFT JOIN ETABLISS ON PDU_ETABLISSEMENT = ET_ETABLISSEMENT ' +
      'LEFT JOIN ORGANISMEPAIE ON PDU_ETABLISSEMENT = POG_ETABLISSEMENT ' +
      'AND PDU_ORGANISME = POG_ORGANISME ' +
      'LEFT JOIN DUCSDETAIL ON  PDU_ETABLISSEMENT =  PDD_ETABLISSEMENT ' +
      'AND PDU_ORGANISME = PDD_ORGANISME AND PDU_DATEDEBUT = PDD_DATEDEBUT ' +
      'AND PDU_DATEFIN = PDD_DATEFIN AND PDU_NUM = PDD_NUM ' +
      StWhereEd + ') ' + StWhereBis + ' ' +
      'ORDER BY PDU_ETABLISSEMENT,PDU_ORGANISME,PDU_DATEDEBUT,PDU_DATEFIN,' +
      'PDU_NUM,PDD_INSTITUTION,PDD_CODIFICATION';
    FunctPGDebEditDucs(); {raz n° page}

    LanceEtat('E', 'PDU', 'DUC', True, False, False, Pages, StSQL, '', False);

    FunctPGFinEditDucs(); {raz AncNumPagSys}
  end;
end; {fin EditerDucs}

// d PT65
{$IFNDEF EAGLSERVER}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE MF
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : Procédure de programmation du process server
Mots clefs ... : PAIE ; PROCESS SERVER
*****************************************************************}
procedure TOF_PGMULDUCSINIT.ParamInitDucs(Sender: TObject);
var
  TobParam, T1: TOB;
begin
  TobParam := TOB.create('Ma Tob de Param', nil, -1);
  T1 := TOB.Create('XXX', TobParam, -1);
  T1.AddChampSupValeur('PERIODETAT', PeriodEtat);
  T1.AddChampSupValeur('NATDEB', NatDeb);
  T1.AddChampSupValeur('NATFIN', NatFin);
  T1.AddChampSupValeur('ETABDEB', EtabDeb);
  T1.AddChampSupValeur('ETABFIN', EtabFin);
//  LancePRocess ....
  AGLFicheJob(0, taCreat, 'cgiPaieS5', 'INITDUCS', TobParam);
  FreeAndNil(TOBParam);
end;
{$ENDIF}
procedure TOF_PGMULDUCSINIT.OnChangeNat(Sender: TObject);
begin
  NatDeb := GetControlText('POG_NATUREORG');
  NatFin := GetControlText('POG_NATUREORG_');

end;
procedure TOF_PGMULDUCSINIT.OnChangeEtab(Sender: TObject);
begin
  EtabDeb := GetControlText('POG_ETABLISSEMENT');
  EtabFin := GetControlText('POG_ETABLISSEMENT_');
end;

// f PT65

// d PT70
{$IFNDEF DUCS41}
Function  TOF_PGMULDUCSINIT.AffectTypeBordereau( NatureOrg, Periodicite : string; DucsDossier : boolean):string;
var TypePeriod : string;
begin
  TypePeriod := '';
  result := '';
    if (NatureOrg = '100') then
    // ACOSS
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '913';    // BRC d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '914';    // BRC de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '915';    // TR d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '916'     // TR de plusieurs établissements
    end
    else
    if (NatureOrg = '200') then
    // UNEDIC
    begin
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (not DucsDossier) then
        TypePeriod := '920';    // ADV d'un établissement
      if ((Periodicite = 'M') or (Periodicite = 'T')) and (DucsDossier) then
        TypePeriod := '921';    // ADV de plusieurs établissements
      if (Periodicite = 'A') and (not DucsDossier) then
        TypePeriod := '922';     // DRA d'un établissement
      if (Periodicite = 'A') and (DucsDossier) then
        TypePeriod := '923'     // DRA de plusieurs établissements
    end
    else
    if (NatureOrg = '300') then
    // IRC
    begin
      if (Periodicite = 'M') then
        TypePeriod := '931';    // Bordereau mensuel
      if (Periodicite = 'T') then
        TypePeriod := '930';    // Bordereau trimestriel
      if (Periodicite = 'A') then
        TypePeriod := '932';    // Bordereau annuel
    end;
  Result := TypePeriod;
end;
{$ENDIF}
// f PT70
initialization
  registerclasses([TOF_PGMULDUCSINIT]);
end.


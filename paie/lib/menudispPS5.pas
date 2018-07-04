{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 06/07/2001
Modifié le ... :   /  /
Description .. : Menu de la paie
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 06/08/2001 V547 PH rajout fonctions menu 43 eAgl car il n'a plus besoins
                           de script dans cws tout est dans la procedure
                           dispatch
PT2   : 17/09/2001 V547 SB Copy des .dot dans le répertoire modèles
                           Ajout de la procédure CopyDotInModele
                           15/11/2002 suppression de la procédure
PT3   : 26/09/2001 V562 SB ReConception module absence
                          Modif. Multicritère
                          Suppression ligne de menu 42254 Gestion absence par
                          salarie
PT4   : 08/10/2001 V562 SB Mise en commentaire de l'appel de la fonction
                          CopyDotInModele..
                          Le copie se fera au niveau du multicritère de
                          selection des salaries voir source UTOFPGSaisieSalRtf
PT5   : 25/10/2001 V562 SB Maj de l'ordreetat des rubriques de cotisations
PT6   : 15/11/2001 V562 SB Ajout d'une ligne de menu : Paramétrage des CP
PT7   : 23/01/2002 V571 SB Ajout point d'entrée pour paramétrage des pays..
PT8   : 14/02/2002 V571 MF Ajout d'une ligne de menu : Groupes Internes
PT9   : 22/02/2002 V571 VG Suppression du calcul DADS-U BTP si SO_PGBTP = False
                           + Gestion de la saisie BTP
PT10  : 27/03/2002 V571 VG Purge des envois DADS-U et prud'hommes
PT11  : 30/04/2002 V582 PH Activation désactivation des lignes de menus
                           correspondant à des options gérées
PT12  : 13/05/2002 V582 MF Ajout menu 41650 : Qualifiant de cotisation
PT13  : 15/05/2002 V582 MF Désactivation menus 42358 et 41840 (DucsEdi)
PT14-1: 21/05/2002 V582 VG Lancement de l'assistant après connexion s'il
                           n'existe pas d'établissement complémentaire
PT14-2: 21/05/2002 V582 VG Version S3
PT15  : 04/06/2002 V582 JL Fiche de bug n°10019 : Si pas d'éxercice social alors ouverture en mode création
PT16  : 11/06/2002 V582 PH Gestion des controles avec saisie des primes
PT17  : 13/06/2002 V582 MF Ajout menu 41841,41843,41845 Envois Ducs EDI
PT18  : 13/08/2002 V582 JL Ajout Menu 47 pour formation
PT19  : 15/11/2002 V591 SB Ajout Traitement par lot pour état chainés
PT20  : 11/12/2002 V591 SB FQ 10326 Ajout Menu Calendrier & jours feriés
PT21  : 06/01/2003 V591 PH Suppression des accès à la compta ==> pgimajver
PT22  : 30/01/2003 V912 JL FQ 10444 Suppression tablette activité MS du menu si paramsoc MSA non coché
PT23  : 10/04/2003 V_42 MF Ajout Menu 42359 : Ducs récapitulative
PT24  : 05/05/2003 V_42 MF Ajout Menu 41487 : Tablette Codes distribution tickets restau./
PT25  : 05/06/2003 V_42 MF Ajout Menus 46961 à 46966 : Tickets restaurant
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT26  : 17/06/2003 V_421 PH initialisation Préférence et gestion des favoris
PT27  : 21/07/2003 V_421 MF Modification appel menu 41260 (ORGANISMES)
PT28  : 04/08/2003 V_421 MF Menus "titres restaurant" déplacés
PT29  : 29/04/2004 V_50  Mf Ajout Menus "Evolution masse salariale et effectifs"
PT30  : 09/06/2004 V_50  MF Ajout Menu "Traitement des IJSS"
PT31  : 20/07/2004 V_50  MF Correction : affichages des modules "bilan social" &
                            "Gestion des compétence" qd sérialisés.
PT32  : 29/07/2004 V_50  MF Complément Menu "Traitement des IJSS & maintien"
PT33  : 07/09/2004 V_50  MF Complément Menu "Traitement des IJSS & maintien"
PT34  : 16/09/2004 V_50  PH FQ 10913 Menu code PCS
PT35  : 29/09/2004 V_50  MF Complément Menu "Traitement des IJSS & maintien"
PT36  : 29/09/2004 V_50  PH Gestion de la sérialisation de la GED PCL
PT37  : 19/10/2004 V_60  JL Gestion des contrôles pour eFormation
PT38  : 08/02/2005 V_60  MF Menu "Traitement des IJSS & maintien", ajout Etats Libres
PT39  : 08/09/2005 V_60  PH FQ 12208 Suppression gestion de la confidentialité en PCL
PT40  : 28/09/2005 V_60  PH FQ 11970 MAJ DPSOCIAL en PCL à chaque sortie de prog
PT41  : 29/12/2005 V_650 PH FQ 12706 Suppression des menus dans exe RH
PT42  : 10/01/2006 V_650 PH FQ 12814 Suppression du positionnement de la barre outil en PCL
PT43  : 27/01/2006 V_650 PH FQ 12849 Affichage des modules cwas PCL uniquement
PT42B : 06/02/2006 V_650 NA MODULE IDR
PT43  : 13/02/2006 V_650 GGS FQ 12854 Tobviewer Multi dossiers Analyse des Salariés
PT44  : 15/02/2006 V_650 RMA FQ 12710 Ajout appel fonction RECAPSAL à partir du menu
PT45  : 24/02/2006 V_650 GGS GRH Etat Analyse de la classification
PT46  : 28/02/2006 V_650 EPI Modification appel exercices sociaux FQ 12380 et 12738
PT47  : 20/03/2006 V_650 RMA Ajout menu 46630 Masse salariale et 46660 Effectifs
PT48  : 30/03/2006 V_650 EPI Ajout menu 42652 et 42653 gestion des processus
PT49  : 09/05/2006 V_650 PH FQ 12377 Ergonomie suppression ligne Analytique si non gérée
PT50  : 29/05/2006 V_650 PH FQ 13156 Séria Diode
PT51  : 31/05/2006 V_650 PH FQ 13212 Import activé en CWAS
PT52  : 14/06/2006 V_700 RMA 46667 Analyse de la classification
PT53  : 15/06/2006 V_650 PH FQ 13293 Menu specifique en paramétrage confidentialité
PT54  : 19/06/2006 V_650 SB Refonte appel ligne de menu du module Absence
PT55  : 06/07/2006 V_70 JL Suppression des menus compétences en trop dans la SOCREF 750
PT56  : 16/10/2006 V_700 SB Maj etats chaines
PT57  : 24/10/2006 V7_00 GGS Tobwiewer analyse des contrats
PT58  : 10/11/2006 V_70 GGR Modif option gestion des contrats(Analyses devient Liste) et Ajout appel Tobwiew_contrat sous nouvelle option analyse
PT58-2: 15/11/2006 V_700 Gestion de la paie 50 & 15
PT59  : 13/12/2006 V_70 Mise en place GED Salarié
PT60  : 02/01/2007 V_80 GGU FQ 13072 Ajout du menu pour la purge des DUCS
PT61  : 01/01/2007 V_80 GGU FQ 13404 Ajout du menu pour la saisie des Bulletins
                            Complémentaires par rubrique
PT62  : 01/01/2007 V_80 GGU FQ 13359 Ajouter paramsoc TD bilatéral et prud'hommes
PT63  : 06/02/2007 V_80 GGU 41460 Ajout du menu pour la liste des tables dynamiques.
PT64  : 09/02/2007 V_80 FC 41215 Ajout du menu pour les éléments nationaux au niveau dossier
PT65  : 20/02/2007 V_80 RMA 42176-42178 Ajout des menus pour la gestion de la DUE
PT66  : 02/03/2007 V_70 FC suppression du menu 41103 si le paramsoc PGGESTELTDYNDOS est à False
PT67  : 07/03/2007 V_70 FC Ajout des menus pour les différents modes d'édition des bulletins
PT68  : 16/03/2007 V_702 RMA Ajout menu 42171 Déclaration accident du travail MSA
        26/03/2007 V_702 RMA + Menu 42172 & 42173 Attestation pour le MSA
PT69  : 27/03/2007 V_702 GGU Menu 46372 Association Paramètres - population.
PT70  : 02/04/2007 V_702 FC Ajout des menus Niveau préconisé 41104 et Synthèse des éléments 41106
PT71  : 03/04/2007 V_702 NA  Menu 46371 Initialisation des affectations des salariés à une population
PT72  : 04/04/2007 V_702 MF ajout appel Initialisation DUCS 42301, Impression DUCS 42302, Impression VLU 42303
PT73  : 04/04/2007 V_702 GGU Modification des menus 41410 et 41460 qui deviennent 41411 et 41412
PT74  : 11/04/2007 V_702 FL Correction du menu "Validation par responsable" en prévisionnel
PT75  : 27/04/2007 V_702 RMA FQ12710 Enlever menu 49814 Export Solde CP
                             + Cacher les menus de la DUE pour Cette version .
PT76  : 02/05/2007 V_72  FC Gestion des menus historiques salarié
PT77  : 30/04/2007 V_720 FL Utilisation de la même fiche pour l'édition des frais de stagiaires depuis le menu
                            Editions et le menu Plan de formation
PT78  : 04/05/2007 V_72 Mf  Séria des IDR;
PT79  : 07/05/2007 V_72 NA  Ajout ligne au menu paramètre
PT80  : 14/05/2007 V_700 Paie 15 et 50 tests sur salariés sortis avec + 2 mois par rapport à la date du jour
PT81  : 22/05/2007 V_720 FL FQ 14051 Ajout du menu 47160 pour l'association d'emplois aux formations
PT82  : 05/06/2007 V_720 FL Ajout du menu 47526 pour le suivi des heures de formation
PT83  : 15/06/2007 V_720 GGU Ajout du menu 41456 pour la gestion de liaisons hierarchiques de la Ged
PT84  : 02/07/2007 V_720 PH Pour pouvoir se connecter à eCabs dans le cas où on n'est pas responsable de formation
PT85  : 13/09/2007 V_80 GGU Contrôle de l'existance des lignes de la tablette gérant les tables libres des compétences
PT86  : 18/09/2007 V_80 MF correction d'un access vio suite à mise enplace nvelle SOC de dev
PT87  : 10/10/2007 V_80 GGU Ajout du menu de lancement de l'outils de diagnostic des rubriques ou des variables
PT88  : 06/11/2007 V_80 PH Implementation en case 12 de l'initialisation du comptage des éditions pour facturation informatique PCL
---- PASSAGE en V810 --------
PT95  : 12/11/2007 V_80 GGU Gestion des menus de la prévoyance DADSU
PT96  : 22/11/2007 V_81 FLO Ajout du traitement multi dossier pour la clôture multi bases
PT97  : 29/11/2007 V_81 GGU Ajout du traitement multi dossier pour la préparation automatique multi bases
PT98  : 04/12/2007 V_81 FLO Gestion du lancement de l'application en mode FICHE
PT99  : 11/12/2007 V_81 RMA Ajout du menu 46668 pour l'edition des salariés présents sur une période FQ13957
PT100 : 21/12/2007 V_81 FLO Additif pour le lancement en mode fiche + Ajout menu 49864
PT101 : 02/01/2008 V_81 GGU Le filtre reste actif sur les tablettes libres après la saisie
PT102 : 08/01/2008 V_81 GGU Ajout du menu 48;1;3;8 (48138) : Compétences manquantes
PT103 : 18/01/2008 V_81 NA  FQ 15126 Ajout au menu TD bilatéral la saisie établissement (pour ajout
PT104 : 13/02/2008 V_81 GGU FQ 15218 Changement du lien
PT105 : 13/02/2008 V_81 FLO Déplacement de l'appel à MultiDosFormation
PT107 : 15/02/2008 V_81 NA  Code produit presence = code produit PAIE si PRESENCE
PT108 : 04/03/2008 V_81 PH  FQ 15266  Si non séria IDR alors module non visible
PT114 : 24/04/2008 V_81 FLO Autorisation du paramétrage regroupement en PCL si param soc PGTRTMULTIDOS
PT115 : 24/04/2008 V_81 FLO Pour CABS5, lors de la connexion, rechercher l'utilisateur salarié aussi dans la table Intérimaires  
PT116 : 07/04/2008 V_81 VG  Correction de la mise à jour de la table intérimaires par rapport à la table salariés
PT117 : 06/06/2008 V_82 MF  FQ 15512
PT119 : 07/08/2008 v_82 NA  FQ 15493 Pb si matricule alpha : ne pas faire le remplissage à 0
}

{======= PENSER A
Rajouter le menu 43 conditionné au code seria

}
unit MenuDispPS5;

interface
uses HEnt1,
  Forms,
  sysutils,
  HMsgBox,
  Classes,
  Controls,
  HPanel,
  UIUtil,
  ImgList,
  Windows,
  Hdebug,
  UTreeLinks,
  UTOB,
  ParamSoc,
  Ent1,
  UtilSoc,
  PGOutilsFormation,
  PGMenusPaie,
{$IFNDEF EABSENCES}
  uObjEtatsChaines,
{$ENDIF EABSENCES}
{$IFNDEF CCS3}
{$IFNDEF EAGLCLIENT}
{$IFNDEF EABSENCES}
  UNetListe,
{$ENDIF EABSENCES}
{$ENDIF EAGLCLIENT}
{$ENDIF CCS3}
  AGLUtilOLE, Hctrls, EntPGI
{$IFNDEF EABSENCES}
  , ImportISIS, AssistSOPaie
{$ENDIF EABSENCES}
{$IFDEF EAGLCLIENT}
  , MenuOLX, M3FP, UtileAGL, MaineAgl, eTablette
{$ELSE}
  , {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, EdtEtat, AGLInit, Tablette, hplanning
  , UTomPaie, MenuOLG, PgiExec, FichComm,
  Doc_parser, UEdtComp
{$IFNDEF SEUL}
  , EtbMce, CODEAFB
{$ENDIF SEUL}
{$ENDIF EAGLCLIENT}

{$IFDEF ETEMPS}
  , eTempsObject, eTempsMenu { Source de saisie des temps }
{$ENDIF ETEMPS}

{$IFDEF MEMCHECK}
  , memcheck
{$ENDIF MEMCHECK}
  , YNewMessage // pour ShowNewMessage
  , AglMail // pour AglMailForm
  , AnnOutils
  , AssistPgExcelOle // Assistant creation des verbes OLE
  , galoutil, galenv // Pour Import automatique des bobs en CWAS
  , PGRepertoire
  , cbpPath
  , BobGestion
{$IFNDEF EAGLCLIENT}
  , Web
{$ENDIF EAGLCLIENT}
  , Messages //PT98
{*  , AfficheOpenOlap*};

  procedure InitApplication;
  procedure RechargeMenuPopPaie;
  function PgAnalConnect: boolean;
  // PT16  : 11/06/2002 V582 PH Gestion des controles avec saisie des primes
  function PgAnalConnectVar: boolean;
  //PT37
  function PgAnalConnectFor: boolean;

  procedure DispatchTraitementLot; //PT19
  Function  ChargeMagHalleyPaie : Boolean; //PT98

{$IFDEF EAGLCLIENT}
// Nothing !
{$ELSE}
  function MyAglMailForm ( var Sujet : HString; Var A, CC: String ; Body: HTStringList ; var Files: String ;
    ZoneTechDataVisible: boolean=True) : TResultMailForm ;
{$ENDIF EAGLCLIENT}

type
  TFMenuDisp = class(TDatamodule)
    procedure FMenuDispCreate(Sender: TObject);
  public
  end;

var
  FMenuDisp: TFMenuDisp;
  SeriaGedPcl: Boolean; // PT36
  SeriaMessOk: Boolean;
  PositionMul : TPoint; //PT100

implementation
uses
  uMultiDossier,
  P5Def,
  P5Util,
  EntPaie,
  PgOutils,
  PgOutils2,
  PGPlanning,
  PgPlanningOutils,
  PGEdtEtat,
  UTofPG_CLOTURE,
{$IFNDEF EAGLCLIENT}
  YYBUNDLE_TOF,
{$ENDIF EAGLCLIENT}
  Variants, //PT62
  GrpFavoris
{$IFNDEF EPRIMES}
{$IFNDEF EABSENCES}
{$IFNDEF EMANAGER}
  , LGCOMPTE_TOF
  , UTOFMULPARAMGEN
{$ENDIF EMANAGER}
{$ENDIF EABSENCES}
{$ENDIF EPRIMES}
  , UtilPgi, UtomUtilisat, Reseau, Mullog
  , Utilged
  , UtilDossier
  , AfterImportGEDPAIE  //MyAfterImportPaie PT59
  , YNewDocument
  , YDocuments_Tom
  , UGedFiles
  , HDTLINKS
{$IFNDEF EAGLCLIENT}
  , CreerSoc, CopySoc, InitSoc
  //  , MajHalley
  , uBob
{$ENDIF EAGLCLIENT}
{$IFDEF PAIEOLE}
  , OlePaiePgi
{$ENDIF PAIEOLE}
{$IFDEF WEBSERVICES}
  , uMajServices
{$ENDIF WEBSERVICES}
{$IFDEF TESTSIC}
{$IFDEF EAGLCLIENT}
  , CegidIEU
{$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}
  , DateUtils
  ;
//PT26  : 17/06/2003 V_421 PH gestion des favoris
var
MemS: Integer;
CodePlace: array[0..39] of string=('05050080',   //0
                                   '00352080',   //1
                                   '00353080',   //2
                                   '00354080',   //3
                                   '00355080',   //4
                                   '00356080',   //5
                                   '00357080',   //6
                                   '00369080',   //7
                                   '00371080',   //8
                                   '00372080',   //9
                                   '00373080',   //10
                                   '00374080',   //11
                                   '05060080',   //12
                                   '00358080',   //13
                                   '00359080',   //14
                                   '00361080',   //15
                                   '00362080',   //16
                                   '00363080',   //17
                                   '00092080',   //18
                                   '00364080',   //19
                                   '00365080',   //20
                                   '00366080',   //21
                                   '00367080',   //22
                                   '00368080',   //23
                                   '00248080',   //24
                                   '00381080',   //25
                                   '00382080',   //26
                                   '00383080',   //27
                                   '00384080',   //28
                                   '00385080',   //29
                                   '05061080',   //30
                                   '00375080',   //31
                                   '00376080',   //32
                                   '00377080',   //33
                                   '00378080',   //34
                                   '00379080',   //35
                                   '00247080',   //36
                                   '00583080',   //37
                                   '01234080',   //38
                                  {$IFDEF PRESENCE}     // pt107
                                   '05050080');   //39  // pt107
                                  {$ELSE}               // pt107
                                   '01234080');  //39
                                  {$ENDIF PRESENCE}     // pt107
LibelPlace: array[0..39] of widestring=('Cegid Business Place Paie',                                     //0
                                    'Cegid Business Place Paie < 300 salariés réseau',               //1
                                    'Cegid Business Place Paie < 500 salariés réseau',               //2
                                    'Cegid Business Place Paie < 750 salariés réseau',               //3
                                    'Cegid Business Place Paie < 1000 salariés réseau',              //4
                                    'Cegid Business Place Paie < 1500 salariés réseau',              //5
                                    'Cegid Business Place Paie > 1500 salariés réseau',              //6
                                    'Option Analyse < 300 salariés Business Place ',                 //7
                                    'Option Analyse < 500 salariés Business Place',                  //8
                                    'Option Analyse < 750 salariés Business Place',                  //9
                                    'Option Analyse < 1000 salariés Business Place',                 //10
                                    'Option Analyse < 1500 salariés Business Place',                 //11
                                    'Option Analyse > 1500 salariés Business Place ',                //12
                                    'Option Formation < 300 salariés Business Place ',               //13
                                    'Option Formation < 500 salariés Business Place',                //14
                                    'Option Formation < 750 salariés Business Place',                //15
                                    'Option Formation < 1000 salariés Business Place',               //16
                                    'Option Formation < 1500 salariés Business Place',               //17
                                    'Option Formation > 1500 salariés Business Place ',              //18
                                    'Option Compétences et Carrières < 300 salariés Business Place ',//19
                                    'Option Compétences et Carrières < 500 salariés Business Place', //20
                                    'Option Compétences et Carrières < 750 salariés Business Place', //21
                                    'Option Compétences et Carrières < 1000 salariés Business Place',//22
                                    'Option Compétences et Carrières < 1500 salariés Business Place',//23
                                    'Option Compétences et Carrières > 1500 salariés Business Place',//24
                                    'Option Econgé < 300 salariés Business Place ',                  //25
                                    'Option Econgé < 500 salariés Business Place',                   //26
                                    'Option Econgé < 750 salariés Business Place',                   //27
                                    'Option Econgé < 1000 salariés Business Place',                  //28
                                    'Option Econgé < 1500 salariés Business Place',                  //29
                                    'Option Econgé > 1500 salariés Business Place',                  //30
                                    'Option Bilan social < 300 salariés Business Place ',            //31
                                    'Option Bilan social < 500 salariés Business Place',             //32
                                    'Option Bilan social < 750 salariés Business Place',             //33
                                    'Option Bilan social < 1000 salariés Business Place',            //34
                                    'Option Bilan social < 1500 salariés Business Place',            //35
                                    'Option Bilan social > 1500 salariés Business Place',            //36
                                    'Option IDR',                                                    //37
                                    'Option Masse salariale',                                        //38
                                    'Option Gestion de la présence');                                //39

CodeSuite: array[0..8] of string=('00387080',     //0
                                  '00071080',     //1
                                  '00389080',     //2
                                  '00391080',     //3
                                  '00392080',     //4
                                  '00393080',     //5
                                  '00394080',     //6
                                  '00395080',     //7
                                  '00582080');    //8
LibelSuite: array[0..8] of widestring=('Cegid Business Suite Paie < 100 salariés',      //0
                                   'Cegid Business Suite Paie < 200 salariés',      //1
                                   'Option Formation < 100 salariés Business Suite',//2
                                   'Option Formation < 200 salariés Business Suite',//3
                                   'Option Analyse < 100 salariés Business Suite',  //4
                                   'Option Analyse < 200 salariés Business Suite',  //5
                                   'Option Econgés < 100 salariés Business Suite',  //6
                                   'Option Econgés < 200 salariés Business Suite', //7
                                   'Option IDR');                                  //8

CodeExpert: array[0..8] of string=('00250080',          //0
                                   '00033080',          //1
                                   '00614080',          //2   pt118
                                   '00042080',          //3
                                   '00616080',          //4   pt118
                                   '00615080',          //5   pt118
                                   '00127080',          //6
                                   '00119080',          //7
                                   '00584080');         //8
LibelExpert: array[0..8] of widestring= ('Paie et GRH',                             //0
                                     'Option Analyse',                          //1   pt118
                                     'Option Formation',                        //2   pt118
                                     'Option Saisie déportée',                  //3   pt118
                                     'Option Bilan social',                     //4    pt118
                                     'Option Compétences',                      //5   pt118
                                     'Messagerie/Agenda',                       //6
                                     'GED',                                     //7
                                     'Option IDR');                             //8

// PT58-2
CodeBL: array[0..1] of string=('00489080','00452080');
LibelBL: array[0..1] of widestring=('Cegid Business Paie 15 pour Cegid Business Line',
                                'Cegid Business Paie 50 pour Cegid Business Line');
// PT58-2
LeMode,LeModeTablette : TActionFiche;
ChargementComplet : Boolean; //PT100

// DEB PT40
procedure PGMAJDPSOCIAL();
var LaDate, DD, DF: TDateTime;
  AA, MM, JJ: WORD;
begin // Uniquement en mode PCL
  LaDate := Now;
  Decodedate(LaDate, AA, MM, JJ);
  DD := DebutDemois(LaDate);
  DF := FindeMois(Ladate);
// On alimente sur le mois en cours que si on est au moins le 20 du mois
  if jj > 20 then PgAlimDPsocial(DD, DF);
  DD := PLUSMOIS(DD, -1);
  DF := PLUSMOIS(DF, -1);
  DF := FindeMois(DF);
  PgAlimDPsocial(DD, DF);
end;
// FIN PT40

function PgChargeFavoris: boolean;
begin
  AddGroupFavoris(FMenuG, ''); // 2eme parama lesTagsToRemove
  result := true;
end;

procedure MyAfterImport(Sender: TObject; FileGuid: string; var Cancel: Boolean);
// Evenement après scannérisation : retourne le FileID du fichier scanné
var
  LastError: string;
  ParGed: TParamGedDoc;
  TobDoc, TobDocGed: TOB;
  DocId: Integer;
begin
    if FileGuid = '' then exit;
  // Propose le rangement dans la GED
        // Propose le rangement dans la GED
      ParGed.SDocGUID := '';
      ParGed.NoDossier := V_PGI.NoDossier;
      ParGed.CodeGed := '';
      // FileId est le n° de fichier obtenu par la GED suite à l'insertion
      ParGed.SFileGUID := FileGuid;

      // Description par défaut du document à archiver...
      if Sender is TForm then ParGed.DefName := TForm(Sender).Caption;

      Application.BringToFront;

      if ((not (ctxPCL in V_PGI.PGIContexte )) or ((ctxPCL in V_PGI.PGIContexte) and (JaileDroitConceptBureau (187315)))) then
      begin
        if ShowNewDocument(ParGed)='##NONE##' then
          // Fichier refusé, suppression dans la GED
          V_GedFiles.Erase(FileGuid);
      end
      else
      begin
        V_GedFiles.Erase(FileGuid);
        PGIInfo ('Vous n''avez pas les droits d''insertion dans la GED.',TitreHalley);
      end;
end;

{$R *.DFM}
// Procedure qui indique que le dossier en mode nomade . Uniquement PCL !!!

procedure AfficheMessPaie;
begin
  PGIBox('Attention, le dossier est parti en nomade.#13#10 Vous ne pouvez pas accèder à la comptabilité.',
    'Gestion du Dossier');
end;

procedure Salarie_ressource;
var
  Tob_Sal: TOB;
  St: string;
  Rep: Integer;
  Q: TQuery;
begin
  Rep := PGIAsk('Vous allez exporter vos salariés pour gérer vos ressources', 'Exportation salariés');
  if Rep <> mrYes then exit;
  st := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_LIBELLEEMPLOI FROM SALARIES';
  Tob_Sal := tob.create('Mon export', nil, -1);
  if St <> '' then
  begin
    Q := OpensQl(st, True);
    Tob_Sal.loaddetaildb('SALARIE', '', '', Q, false);
  end;
  if TOB_Sal.detail.count >= 1 then TOB_Sal.SaveToFile('C:\PGI00\STD\EXPORTSR.TXT', False, True, True, '');
  if st <> '' then Ferme(Q);
  FreeAndNil(Tob_Sal);
  PGIBOX('Traitement terminé', 'Exportation salariés');
end;

procedure TESTSQL(); // Fionction de tests SQL avant modif pour le pgimajver
var St: string;
begin
  ExecuteSql(St);
end;

procedure Dispatch (Num: Integer; PRien: THPanel; var retourforce, sortiehalley: boolean);
var
  Q : TQuery;
  QVerifTablesLibre : Tob; //PT85
  Nbre, NbreI, NbreS, Rest, LeNbre, IndexTableLibre: Integer;
  st, LesModulPaie, LeMess: string;
  NiveauRupt: TNiveauRupture;
  LeCodeseria: array of string;
  CEG, STD, DOS: Boolean;
  MoisE, AnneeE, ComboExer: string;
  DebExer, FinExer: TDateTime;
{$IFDEF CPS1}
  LaDate : TDateTime; // PT80
{$ENDIF}
  Ok: Boolean;
begin
  case Num of
    // PT-1 : 06/08/2001 : V547 : PH Rajout clause identifictaion eAGL pour supprimer accès au script
    10:
      begin
      if (V_PGI.ModePcl <> '1') then
         begin
         PgAffecteNoDossier(); // Affectation Nodossier en multi-entreprise
         PGRendNoDossier();
         end;
        // En mode fiche : Envoi du message 3 : Chargement des paramètres
        If VH_Paie.ModeFiche Then PostMessage(HWND_BROADCAST, MessageInterApp, 3, 0); //PT98

//      if V_PGI.ModePcl <> '1' then NbreS := PGRendNbreSal();
//      TESTSQL();
{$IFNDEF EABSENCES}
{$IFNDEF EMANAGER}
{$IFNDEF EPRIMES}
        RAZTOBPaie();
        { Le chargement du contexte de la paie se fait uniquement
        dans le cas du calcul du bulletin
        if VH_Paie.PGContextePaie = 'OUV' then // Après connection
        begin
          InitLesTOBPaie;
          ChargeLesTOBPaie;
        end;
        }
{$ENDIF EPRIMES}
{$ENDIF EMANAGER}
{$ENDIF EABSENCES}
{$IFNDEF EABSENCES}
        PGRempliDefaut; // Initialisation des tablettes choixcod par defaut
{$ENDIF EABSENCES}
        PgModifNomChamp;

        //PT98 //PT100
        If ChargementComplet Then
        Begin
            MultiDosFormation; //PT105

            PGChangeTabletteResponsable;
            //         MajConvention;
            //            MajRubrique;
            //            MajCotisation; //PT-5 25/10/2001 SB V562 Pour maj de l'ordreetat des cotisations
            // Fonction analyse de la connection eAGL pour savoir si les descriptifs outlook des salaries sont corrects
{$IFDEF EABSENCES}
            if PgActivAbsences = FALSE then sortiehalley := TRUE; // Attention, non autorisation connection à la base eAGL
{$IFDEF EMANAGER}
            Ok := False;
            if not PgActivFormation then EFormation_menus(-1, PRien)
            else OK := True;
            if not PgActivPrimes then EPrimes_Menus(-1, PRien)
            else OK := True;
            if not PgActivAugmentation then EAugmentation_Menus(-1, PRien)
            else OK := True;
{$ENDIF EMANAGER}
{$ENDIF EABSENCES}

{$IFDEF ENCOURS}
//          if PgAnalConnect = FALSE then sortiehalley := TRUE;
{$ENDIF ENCOURS}

{$IFDEF EPRIMES}
            if PgAnalConnectVar = FALSE then sortiehalley := TRUE; // Attention, non autorisation connection à la base eAGL
{$ENDIF EPRIMES}
{$IFDEF EMANAGER}
            if PgAnalConnectFor = FALSE then sortiehalley := TRUE; // Attention, non autorisation connection à la base eAGL
{$ENDIF EMANAGER}

            if V_PGI.ModePcl = '1' then
            begin
                V_PGI.QRPdf := True;
                // (Le container de la GED est la base commune)
                {FCOM bureau N°22261
               if SeriaGedPcl then
                  begin
                }
               if V_PGI.RunFromLanceur then
                  InitializeGedFiles (V_PGI.DeFaultSectionDBName, MyAfterImport)
               else
                  InitializeGedFiles(V_PGI.CurrentAlias, MyAfterImport);
                {FCOM bureau N°22261
                              end;
                }
                V_PGI.GedActive:= SeriaGedPcl;
                //FIN FCOM bureau N°22261
            end
            else
            begin
//              if V_PGI.DebugSQL then PGIBox('GED non sérialisée', 'Sérialisation');
                InitializeGedFiles(V_PGI.DBName, MyAfterImportPaie); // PT59
            end;

            //MultiDosFormation; //PT105

            //PT14-1  //PT116
            Q := OpenSql('Select count (*) FROM SALARIES WHERE '+
                        'PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '"', True);
            if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
            ferme(Q);

            {$IFNDEF EAGLCLIENT}
            if (VH_PAIE.PGGESTIONCARRIERE) or (GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE)) or (PGBundleHierarchie) then
            begin
			//PT116
            Q:= OpenSql ('SELECT COUNT(*)'+
                         ' FROM INTERIMAIRES '+
                         ' LEFT JOIN SALARIES ON PSI_INTERIMAIRE=PSA_SALARIE '+
                         'WHERE PSI_NODOSSIER="'+V_PGI.NoDossier+'" AND'+
                         ' PSI_TYPEINTERIM="SAL" AND PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '"', TRUE);
                if not Q.EOF then NbreI := Q.Fields[0].AsInteger;
                ferme(Q);
                if Nbre <> NbreI then DupliqSalInterimaire; //différent, pas inférieur
            end;
            if not (V_PGI.ModePcl = '1') then
            begin
                if NBreS = 0 then
                begin
                    // DEB PT80
{$IFDEF CPS1}
                    LaDate := PlusMois (Now, 2);
{$ENDIF}
                    st := 'Select count (*) FROM SALARIES WHERE PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE > "' +
{$IFNDEF CPS1}
                    UsDateTime(Now) + '"';
{$ELSE}
                    UsDateTime(LaDate) + '"';
{$ENDIF}
                    // FIN PT80
                    Q := OpenSql(St, TRUE);
                    if not Q.EOF then NbreS := Q.Fields[0].AsInteger;
                    ferme(Q);
                end;
{$IFNDEF CPS3}
                if MemS = 0 then LeNbre := 5;
                if MemS = 1 then LeNbre := 300;
                if MemS = 2 then LeNbre := 500;
                if MemS = 3 then LeNbre := 750;
                if MemS = 4 then LeNbre := 1000;
                if MemS = 5 then LeNbre := 1500;
                if MemS = 6 then LeNbre := 9999999;
{$ELSE}
{$IFDEF CPS3}
                if MemS = 0 then LeNbre := 5;
                if MemS = 1 then LeNbre := 100;
                if MemS = 2 then LeNbre := 200;
{$ENDIF CPS3}
{$IFDEF CPS1} // Paie 15 & 50 Limite à 15 ou 50 salariés maxi
                if Mems = 2 then LeNbre := 50
                else LeNbre := 15;
{$ENDIF CPS1}
{$ENDIF CPS3}
                Rest := LeNbre - NbreS;
                LeMess := 'Votre licence autorise la création de : ' + IntToStr(LeNbre) + '  salariés alors que vous gérez : ' + IntToStr(NbreS) + ' salariés.';
{$IFDEF CPS1} // Paie 15 & 50 Limite à 15 ou 50 salariés maxi
                if Rest < 0 then
                    PgiBox('Attention, vous n''êtes plus autorisé à créer des salariés. #13#10Vous devez changer de licence.#13#10' + LeMess)
                else
                    if (Rest < 5)  then  PgiInfo(LeMess);
{$ELSE}
                if V_PGI.SAV then PgiInfo(LeMess, 'Contrôle sérialisation');
                if (Rest < 30) and (Rest > 0) then
                    PgiInfo ('Attention, votre licence vous permet de créer encore '+IntToStr(Rest)+' salarié(s).#13#10'+LeMess);
                if Rest < -30 then
                    PgiBox('Attention, vous n''êtes plus autorisé à créer des salariés. #13#10Vous devez changer de licence.#13#10' + LeMess);
{$ENDIF CPS1}
            end;

            FMEnuG.ChoixModule; // Pour afficher la liste des modules
{$ENDIF EAGLCLIENT}
        End;

        DispatchTraitementLot; //PT19

        If Not VH_PAIE.ModeFiche Then //PT98
        Begin
{$IFDEF EAGLCLIENT}
{$ELSE}
            if SeriaMessOK then OnAglMailForm := MyAglMailForm;
{$ENDIF EAGLCLIENT}

{$IFNDEF EABSENCES}
{$IFDEF EAGLCLIENT}
            if (V_PGI.ModePcl='1') then
{$ENDIF EAGLCLIENT}
               BOB_IMPORT_PCL_STD ('CPS5', 'Bob intégrée', False);
{$ENDIF EABSENCES}

{$IFNDEF CPS3}
            {$IFNDEF EAGLCLIENT}
{$IFNDEF EABSENCES}
            if (V_PGI.ModePcl = '1') then
            begin
              if InterrogeCorbeil('PAIE', 'NSV', FALSE) <> nil then
              begin
                V_PGI.ZoomOLE := TRUE;
                AGlLanceFiche('CP', 'NETLISTEFICHIER', '', '', 'PAIE;NSV');
                V_PGI.ZoomOLE := FALSE;
              end;
            end;
{$ENDIF EABSENCES}
            {$ENDIF EAGLCLIENT}
{$ENDIF CPS3}
            if (V_PGI.ModePcl <> '1') then
            begin
              PGRendNoDossier();
              PgAffecteNoDossier(); // Affectation Nodossier en multi-entreprise
            end;
{$IFNDEF EABSENCES}
            if (Nbre = 0) then
                LanceAssistantPaie; // Assistant de création
{$ENDIF EABSENCES}
            //FIN PT14-1

            //Mise à jour table ABSENCESALARIE
            try
                if (not (GetParamSocSecur ('SO_PGMAJABS', False))) then
                   begin
                   SetParamSoc ('SO_PGMAJABS', True);
                   ExecuteSQL ('UPDATE ABSENCESALARIE SET'+
                               ' PCN_ETABLISSEMENT = (SELECT PSA_ETABLISSEMENT FROM'+
                               ' SALARIES WHERE PSA_SALARIE=PCN_SALARIE) WHERE'+
                               ' PCN_TYPEMVT="CPA" AND'+
                               ' PCN_TYPECONGE="PRI" AND'+
                               ' PCN_ETABLISSEMENT NOT IN (SELECT ETB_ETABLISSEMENT FROM ETABCOMPL)');
                   end;
            except
            end;
            //FIN mise à jour table ABSENCESALARIE
            if JaiLeDroitTag(200121) then LeModeTablette := taCreat
            else LeModeTablette := taConsult;
        end;
        PositionMul := PRien.ClientOrigin; //PT100
      end;
    11:
      begin
        VideLesTOBPaie(FALSE); //Après déconnection
{$IFDEF PAIEOLE}
        PaieVideAnc ();
{$ENDIF PAIEOLE}
        PgVideHabilitation (); // Suppression de l'objet habilitations
      end;
    12:
      begin // avant connexion
        //PT98 - Début
        If VH_PAIE.ModeFiche Then
        Begin
            // En mode fiche, envoi du message 2 : Connexion à la base
            PostMessage(HWND_BROADCAST, MessageInterApp, 2, 0);
            // On cache la fenêtre principale et on la met transparente pour
            // limiter au maximum l'apparition à l'utilisateur de la nouvelle application lancée
            Application.MainForm.Visible := False;
            Application.MainForm.AlphaBlend := False;
        End;
        //PT98 - Fin
        V_PGI.CodeProduit := '013';
        V_PGI.QRPdf := True;
        V_PGI.LaSerie := S5;
        {
        if (Not DelphiRunning) and (V_PGI.ModePcl<>'1') then
            begin  // Pas de connection en PCL sur la base si on passe pas par le lanceur
              FMenuG.ForceClose := True;
              FMenuG.Close;
              exit;
            end;
         }
        if FileExists('CEGID.par') then DroitCegPaie := TRUE
        else DroitCegPaie := FALSE;
        //  Modif Pour tester la seria PCL , pour avoir les variables V_PGI... initilalisees on fait le test aprs la connexion et pas dans init de l'appli
{$IFDEF CPS3}
{$IFDEF CPS1}
        FMenuG.SetSeria('00453010', CodeBL, LibelBL); // Paie 50 BL
{$ELSE}
        FMenuG.SetSeria('00388010', CodeSuite, LibelSuite); // Paie suite
{$ENDIF CPS1}
{$ELSE}
//        PgiBox ('Séria = ' + V_PGI.ModePcl, 'TEST');
        if (V_PGI.ModePcl = '1') then
{$IFNDEF RHSEUL} // PT50
          FMenuG.SetSeria('00280011', CodeExpert,LibelExpert) //PCL  PT78
        else
          FMenuG.SetSeria('05970012', CodePlace, LibelPlace);
{         FMenuG.SetSeria('05970012', ['05050060', '05050060', '05050060', '05050060', '05050060', '05060060', '00092060', '05061060', '05051060', '00247060', '00248060'],
            ['Paie et GRH < 300 salariés', 'Paie et GRH < 500 salariés', 'Paie et GRH < 750 salariés', 'Paie et GRH < 1000 salariés', 'Paie et GRH + 1000 salariés', 'Analyses', 'Gestion de la formation', 'Saisie déportée', 'TD Bilatéral', 'Bilan social', 'Gestion des compétences']); // Paie entreprise
}
{$ELSE}
          FMenuG.SetSeria('00280011', ['00250080', '00092080', '00247080', '00248080', '00127080', '00119080'],
            ['Paie et GRH', 'Gestion de la formation', 'Bilan social', 'Gestion des compétences', 'Messagerie/Agenda', 'GED']) //PCL
        else
          FMenuG.SetSeria('05970012', CodePlace, LibelPlace);
{
          FMenuG.SetSeria('05970012', ['05050060', '05050060', '05050060', '05050060', '05050060', '00092060', '00247060', '00248060'],
            ['Paie et GRH < 300 salariés', 'Paie et GRH < 500 salariés', 'Paie et GRH < 750 salariés', 'Paie et GRH < 1000 salariés', 'Paie et GRH + 1000 salariés', 'Gestion de la formation', 'Bilan social', 'Gestion des compétences']); // Paie entreprise
            }
{$ENDIF RHSEUL}
{$ENDIF CPS3}
{$IFNDEF EAGLCLIENT}
        // Interdiction de lancer en direct
        if (not V_PGI.RunFromLanceur) and (V_PGI.ModePCL = '1') then
        begin
          FMenuG.FermeSoc;
          FMenuG.Quitter;
          exit;
        end;
{$ENDIF EAGLCLIENT}
        // A rajouter ici tous les tests sur modepcl ou bien voir si tjrs possible dans InitLaVariablePGI car modif AGL
{$IFNDEF CCS3}
{$IFNDEF RHSEUL}
        if (V_PGI.ModePcl = '1') then
        begin // PT88
          TitreHalley := 'PAIE-GRH';
          V_PGI.CompterLesimpressions:= true;
          V_PGI.PRocCompteImpression:= CompteImpression;
        end
        else TitreHalley := 'PAIE-GRH';

{$ELSE}
        if (V_PGI.ModePcl = '1') then TitreHalley := 'Ressources humaines'
        else TitreHalley := 'Ressources humaines';
{$ENDIF RHSEUL}
{$ELSE}
        TitreHalley := 'PAIE';
{$ENDIF CCS3}
{$IFDEF CPS1}
        TitreHalley := 'Paie 50';
{$ENDIF CPS1}
      end;
    13: begin
        if V_PGI.ModePcl = '1' then PGMAJDPSOCIAL(); // PT40
        FinalizeGedFiles; // le container de la ged est la base commune
        end;
    15 : ;
{$IFDEF etemps}
    16:
      begin
        NotSalAss := False;
        SalAdm := 'X';
        SalVal := 'X';
        et_menu(-1);
      end;
{$ELSE}
{$IFDEF EABSENCES}
    16:
      begin
        NotSalAss := False;
        SalAdm := 'X';
        SalVal := 'X';
      end;
    100: ;
{$ELSE}
    // suppression des fichiers temporaires PDF qui sont en bases
    16: begin
        //If Not VH_PAIE.ModeFiche Then //PT98
        If ChargementComplet Then //PT100
        begin
            If existeSQL('SELECT * FROM YFILESTD WHERE YFS_NODOSSIER="" AND YFS_CODEPRODUIT="PAIE"') then
            begin
                ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_NODOSSIER="" AND YFS_CODEPRODUIT="PAIE")');
                ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID IN (SELECT YFS_FILEGUID FROM YFILESTD WHERE YFS_NODOSSIER="" AND YFS_CODEPRODUIT="PAIE")');
                ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_NODOSSIER="" AND YFS_CODEPRODUIT="PAIE"');
            end;
            PGDeleteRepertTemporaire(PGGetTempPath + 'PGI\STD\PAIE');
            st := ChangeStdDatPath('$STD');
            if POS('STD', St) > 0 then PGDeleteRepertSTDPDF(st);
        end;
      end;
{$ENDIF EABSENCES}
{$ENDIF etemps}
    {------------------------------------------------------------
                                STANDARDS
    -------------------------------------------------------------}
{$IFNDEF EABSENCES} // EABSENCES -0-
    100: // action A Faire quand on passe par le lanceur
      begin
        if V_PGI.ModePCL = '1' then
        begin
          Q := OpenSql('Select count (*) FROM SALARIES', TRUE);
          if not Q.EOF then
          begin
            Nbre := Q.Fields[0].AsInteger;
          end;
          ferme(Q);
          if (not GetFlagAppli('CPS5.EXE')) and (Nbre >= 1) then
             SetFlagAppli('CPS5.EXE', TRUE);
        end;

        FMEnuG.ChoixModule; // Pour afficher la liste des modules PT43
      end;
    {    40015: AglLanceFiche('PAY', 'ELTNAT_MUL', '', '', '');
        40020: AglLanceFiche('PAY', 'CONVENTION_MUL', '', '', ''); //07/04/03 Modif. fiche
        40160: AglLanceFiche('PAY', 'CUMUL', '', '', '');
        40115: AglLanceFiche('PAY', 'COEFFICIENT_MUL', 'PMI_TYPENATURE=INT', '', '');
        40140: AglLanceFiche('PAY', 'VARIABLE_MUL', '', '', '');
        40125: AglLanceFiche('PAY', 'COTISATION_MUL', '', '', '');
        40130: AglLanceFiche('PAY', 'REMUNERATION_MUL', '', '', '');
        40150: AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
    }
    4006: AglLanceFiche ('YY','YYDEFPRT', '','','TYPE=E;FILTRENATURE= LIKE "P%"'); // Imprimantes par défaut
{$IFNDEF EAGLCLIENT}
    4042: Begin
               if (V_PGI.ModePcl = '1') then ExecuteSQL ('UPDATE DOSSIER SET DOS_NOMBASE="DB"||DOS_NODOSSIER');
               YYLanceFiche_Bundle('YY', 'YYBUNDLE', '', '', '');
          End;
{$ENDIF EAGLCLIENT}
    4041: ParamRegroupementMultidossier;
        {------------------------------------------------------------
                   PARAMETRES  : Gestion du paramètrage Dossier
        -------------------------------------------------------------}
    //   parametrage des éléments nationaux dossier
    41105: AglLanceFiche('PAY', 'ELTNAT_MUL', '', '', '');
    41104: AglLanceFiche('PAY', 'ELTNIVREQUIS_MUL', '', '', '');  //PT70
//    41103: AglLanceFiche('PAY', 'ELTNATDOS_MUL', '', '', '');  //PT64
//    41106: AglLanceFiche('PAY', 'SYNELTNAT_MUL', '', '', '');  //PT70
    41107: AglLanceFiche('PAY', 'CONVENTION_MUL', '', '', ''); //07/04/03 Modif. fiche
    41150: AglLanceFiche('PAY', 'CUMUL_MUL', '', '', '');
    41115: AglLanceFiche('PAY', 'COEFFICIENT_MUL', 'PMI_TYPENATURE=INT', '', '');
    41140: AglLanceFiche('PAY', 'VARIABLE_MUL', '', '', '');
    41130: AglLanceFiche('PAY', 'COTISATION_MUL', '', '', '');
    41120: AglLanceFiche('PAY', 'REMUNERATION_MUL', '', '', '');
    41110: AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
    41160: AglLanceFiche('PAY', 'MUL_MASQSAISRUB', '', '', '');
    41165: AglLanceFiche('PAY', 'ACTIVITE', '', '', '');
    41171: AglLanceFiche('PAY', 'PGREAFFCODEPCS', '', '', '');
    41172: AglLanceFiche('PAY', 'PGREAFFCODEPCS', '', '', 'N');
    41173:
      begin
        ChargePCS(); // Test et chargement de la table de correspondance à partir de la BOB
        AglLanceFiche('PAY', 'MUL_CORRESPCS', '', '', '');
      end;
    41174: AglLanceFiche('PAY', 'DOUBLON_PCS', '', '', '');

    41181: begin
        AccesPredefini('TOUS', CEG, STD, DOS);
        if (DOS or STD) then AglLanceFiche('PAY', 'MUL_PGEXECPT', '', '', '')
        else PgiInfo('Vous n''êtes pas autorisé(e) à personnaliser des rubriques.', 'Personnalisations des rubriques');
      end;

    41183: begin
           AccesPredefini('TOUS', CEG, STD, DOS);
           if STD then
              AglLanceFiche('PAY', 'MUL_PGCUMEXCEPT', '', '', '')
           else
              PgiInfo ('Vous n''êtes pas autorisé(e) à personnaliser des cumuls.',
                       'Personnalisations des cumuls');
      end;
    41191: AGLLanceFiche('PAY', 'PGARCHIREGLPAIE', '', '', 'ARCHITECTURE'); //PT87
    41192: AGLLanceFiche('PAY', 'PGARCHIREGLPAIE', '', '', ''); //PT87
    {------------------------------------------------------------
                                DOSSIER
    -------------------------------------------------------------}
    //   parametrage des paies
    41210: AglLanceFiche('PAY', 'PARAMETRESOC', '', '', '');
    //     41220 : FicheEtablissement_AGL(taModif) ;
    41220:
      begin
        AglLanceFiche('PAY', 'ETABLISSEMENT', '', '', '');
        MajEtabCompl; // Analyse et controle des champs à NULL
      end;
    41240: AglLanceFiche('YY', 'YYPAYS', '', '', '');
    {begin     //PT-7 Paramétrage Pays => fiche Bq appellé dans un autre module
      If NOT IsDossierPartiSauf(V_PGI_Env.NoDossier, 'CPS5.EXE') then
              FicheBanque_agl('',taModif,0)
         else AfficheMessPaie;
     end;  }
    41250: AglLanceFiche('PAY', 'DONNEURORDRE_MUL', '', '', '');
    41260: AglLanceFiche('PAY', 'MUL_ORGANISMES', '', '', ''); // PT27
    41270: AglLanceFiche('PAY', 'PAIEBQORG', '', '', ''); // PT18

    41230:
      begin //PT15 AglLanceFiche ('PAY','EXERCICESOCIAL', '', '' , '' );
        if not ExisteSQL('SELECT PEX_EXERCICE FROM EXERSOCIAL') then
         // PT46 AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', 'ACTION=CREATION')
          AglLanceFiche('PAY', 'EXERSOCIAL', '', '', 'ACTION=CREATION')
        else // PT46 AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', '');
          AglLanceFiche('PAY', 'MUL_EXERSOCIAL', 'GRILLE=', '', '');
      end;

    //PT20 Calendriers & Jours fériés
    41282: AGLLanceFiche('AFF', 'HORAIRESTD', '', '', 'TYPE:STD');
    41285: AglLanceFiche('AFF', 'JOURFERIE', '', '', '');

    41290: AglLanceFiche('PAY', 'DECLARANT_MUL', '', '', '');

    41296: AglLanceFiche('PAY', 'MUL_CONTRATPREV', '', '', ''); //PT95
    41297: AglLanceFiche('PAY', 'MUL_DADSUAFFPREV', '', '', ''); //PT95

    //Banques PH le 201202 suppression des accès aux télétrans ==> fait par CTTRANS de XP pour l'ensemble des applis
{     41510 : begin
             If NOT IsDossierPartiSauf(V_PGI_Env.NoDossier, 'CPS5.EXE') then
                    FicheBanque_agl('',taModif,0)
                else AfficheMessPaie;
             end;}
    //41520 : ParamTable('TTCFONB',taCreat,0,PRien) ;
   // 41530 : ParamCodeAFB ;
{     41540 : begin
             If NOT IsDossierPartiSauf(V_PGI_Env.NoDossier, 'CPS5.EXE') then
                    ParamTeletrans;
             end;
    41550 : AGLLanceFiche('CP', 'RLVEMISSION', '', '', '') ;
}
// edition des Paramètres
    41310: AglLanceFiche('PAY', 'VARIABLE_ETAT', '', '', '');
    41320: AglLanceFiche('PAY', 'PROFIL_ETAT', '', '', '');
    41330: AglLanceFiche('PAY', 'CUMUL_ETAT', '', '', '');
    41340: AglLanceFiche('PAY', 'GESTASSOC_ETAT', '', '', '');
    41350: AglLanceFiche('PAY', 'RUBRIQUE_ETAT', '', '', '');
    41360: AglLanceFiche('PAY', 'EDITLIENRUBRIQUE', '', '', '');

    //   gestion des tables du dossier
    41411: AglLanceFiche('PAY', 'COEFFICIENT_MUL', 'PMI_TYPENATURE=INT', '', '');  //PT73
    41415: AglLanceFiche('PAY', 'PGMULTITABLE_MUL', '', '', '');
    //PT63 Tables dynamiques
    41412: AglLanceFiche('PAY', 'PGMULTABLESDYNA', '', '', '');   //PT73
    // Annuaire sur les organismes de type social du DP
    41420: AglLanceFiche('PAY', 'PGANNUAIRE', '', '', '');
    41431:If not DroitCegPaie then  AglLanceFiche('PAY', 'PARAMHISTOSTDDOS', '', '', '')
    else AglLanceFiche('PAY', 'MUL_PARAMHISTO', '', '', 'SAL');
    41432: AglLanceFiche('PAY', 'MUL_INITHISTO', '', '', '');
    41441: AglLanceFiche('PAY', 'MUL_PARAMHISTO', '', '', 'ZLS');
    41442: AglLanceFiche('PAY', 'MULZLNIVEAUPOP', '', '', 'POP');
    41443: AglLanceFiche('PAY', 'MULZLNIVEAUETAB', '', '', 'ETB');
    41444: AglLanceFiche('PAY', 'SAISIEZLNIVEAU', '', '', 'DOS');
    // GED Salarié
    41451: AglLanceFiche('PAY', 'RECHGEDPAIE', '', '', '');
    41452: ParamTable('PGLIBGED', LeModeTablette, 0, PRien);
    41453: ParamTable('PGLIBGED1', LeModeTablette, 0, PRien);
    41454: ParamTable('PGLIBGED2', LeModeTablette, 0, PRien);
    41455: ParamTable('PGLIBGED3', LeModeTablette, 0, PRien);
    41456: begin  //PT83
      with ShowTreeLinks(PRien, TraduireMemoire('Définition des liaisons Hiérarchiques'),'PAIEGED%') do
      begin
        btnMainNormalize.Visible := False; // Normaliser les liaisons hiérarchiques
      end;
    end;
    //  gestion des tablettes utilisateurs
    41471: ParamTable('PGTRAVAILN1', LeModeTablette, 0, PRien);
    41472: ParamTable('PGTRAVAILN2', LeModeTablette, 0, PRien);
    41473: ParamTable('PGTRAVAILN3', LeModeTablette, 0, PRien);
    41474: ParamTable('PGTRAVAILN4', LeModeTablette, 0, PRien);
    41475: begin
             if JaiLeDroitTag(200102) then LeMode := taCreat
               else LeMode := taConsult;
             ParamTable('PGLIBEMPLOI', LeMode, 0, PRien);
           end;
    41476: AglLanceFiche('PAY', 'MINICONVENT_MUL', 'PMI_NATURE=COE', 'PMI_TYPENATURE=VAL', 'C');
    41477: AglLanceFiche('PAY', 'MINICONVENT_MUL', 'PMI_NATURE=IND', 'PMI_TYPENATURE=VAL', 'I');
    41478: AglLanceFiche('PAY', 'MINICONVENT_MUL', 'PMI_NATURE=NIV', 'PMI_TYPENATURE=VAL', 'N');
    41479: AglLanceFiche('PAY', 'MINICONVENT_MUL', 'PMI_NATURE=QUA', 'PMI_TYPENATURE=VAL', 'Q');
    41480: ParamTable('PGCODESTAT', LeModeTablette, 0, PRien);
    41481: ParamTable('PGTYPEORGANISME', taCreat, 0, PRien, 3, 'Type d''organisme');
    41485: ParamTable('PGMSAACTIVITE', taCreat, 0, PRien);
    //     41487 : ParamTable('PGCODISTRIBUTION',taCreat,0,PRien) ; //PT24 PT28
    //Motifs
        // 41491:AglLanceFiche('PAY','MOTIFENTREE','','','');
        // 41492 :AglLanceFiche('PAY','MOTIFSORTIE','','','');
    41495: AglLanceFiche('PAY', 'MOTIFABSENCE_MUL', '', '', '');
    41196: ParamTable('PGMOTIFSUSPAIE', taCreat, 0, PRien);
    41197: ParamTable('PGMOTIFINTERIM', taCreat, 0, PRien);

    //   gestion parametrage de la DUCS
    41610: AglLanceFiche('PAY', 'MUL_INSTITUTIONS', '', '', '');
    41620: AglLanceFiche('PAY', 'MUL_CODIFDUCS', '', '', '');
    41630: AglLanceFiche('PAY', 'MUL_DUCSAFFECT', '', '', 'D');
    41640: AglLanceFiche('PAY', 'GROUPINTERNE', '', '', ''); // PT8
    41650: AglLanceFiche('PAY', 'MUL_COTISQUAL', '', '', ''); // PT12

    //   gestion parametrage de la DADS-U
    41665: AglLanceFiche('PAY', 'MUL_DADSUAFFECT', '', '', '');


    //gestion des tablettes libres
{$IFDEF COMPTAPAIE}
    41701: ParamTable('PGLIBREPCMB1', LeModeTablette, 0, PRien);
    41702: ParamTable('PGLIBREPCMB2', LeModeTablette, 0, PRien);
    41703: ParamTable('PGLIBREPCMB3', LeModeTablette, 0, PRien);
    41704: ParamTable('PGLIBREPCMB4', LeModeTablette, 0, PRien);
    //gestion des ecritures comptables == acces compta + Modele ecriture paie
    41710:
      begin
        // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
{$IFNDEF EAGLCLIENT}
        if not IsDossierPartiSauf(PgRendNoDossier(), 'CPS5.EXE') then
          ParamSociete (False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX',
                        '', ChargeSocieteHalley, ChargePageSoc, SauvePageSoc,
                        InterfaceSoc, 1105000)
        else AfficheMessPaie;
{$ELSE}
        ParamSociete (False, BrancheParamSocAVirer, 'SCO_PARAMETRESGENERAUX',
                      '', ChargeSocieteHalley, ChargePageSoc, SauvePageSoc,
                      InterfaceSoc, 1105000);
{$ENDIF EAGLCLIENT}
      end;
    // PT21    06/01/2003 V591 PH Suppression des accès à la compta ==> pgimajver
    41715: CCLanceFiche_MULGeneraux('C;');
    41720: CCLanceFiche_LongueurCompte;
    {     41720 :  begin
                  If NOT IsDossierPartiSauf(V_PGI_Env.NoDossier, 'CPS5.EXE') then
                          MultiCritereTiers(taModif)  // comptes auxiliaires
                     else AfficheMessPaie;
                  end;
         41730 :  begin
                  If NOT IsDossierPartiSauf(V_PGI_Env.NoDossier, 'CPS5.EXE') then
                          MulticritereJournal(taModif)  // Journaux
                     else AfficheMessPaie;
                  end;
    }
{$ENDIF COMPTAPAIE}
    41791: AglLanceFiche('PAY', 'VENTILSTD_MUL', '', '', 'R');
    41792: AglLanceFiche('PAY', 'VENTILCOT_MUL', '', '', 'C');
    41793: AglLanceFiche('PAY', 'VENTILORG_MUL', '', '', 'O');
    41795: AglLanceFiche('PAY', 'JEUECRPAIE', '', '', '');
    {------------------------------------------------------------
        Gestion des envois des concernant le social DADS-U,DUCS,Prud'hom
    -------------------------------------------------------------}
    41802: AglLanceFiche('PAY', 'MUL_EMETTEURSOC', '', '', '');
    41811: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSU');
    41812: AglLanceFiche('PAY', 'EDITFICHIERDADS', '', '', 'FICHIER'); // Edition DADS à partir fichier
    41813: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'DADSU');
    41815: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSUP'); //PT10
    41814: AglLanceFiche('PAY', 'EDITDADSCONT', '', '', 'S'); // Edition du contrôle STanDard

    41817: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSB');
    41818: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'DADSB');
    41819: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DADSBP');

    41821: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'PRUDH');
    41822: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'PRUDH');
    41825: AglLanceFiche('PAY', 'CONTROL_PRUD', '', '', 'PRUDH');
    41827: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'PRUDHP'); //PT10

    41841: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DUCS'); // PT17
    41843: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'DUCS'); // PT17
    41845: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'DUCSP'); // PT17

    41861: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'CONGSPEC');
    41863: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'CONGSPEC');
    41865: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'CONGSPECP');

    41871: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'MSA');
    41873: AglLanceFiche('PAY', 'MUL_CONSULTENVSOC', '', '', 'MSA');
    41875: AglLanceFiche('PAY', 'MUL_ENVOISOCIAL', '', '', 'MSAP');

    //MISE EN BASE FICHIERS
    41910: AglLanceFiche('PAY', 'IMPORTRTF', '', '', '');
    41920: AglLanceFiche('PAY', 'EXPORTRTF', '', '', '');
    41930: AglLanceFiche('PAY', 'RTFDEFAUT', '', '', '');
    41940: AglLanceFiche('PAY', 'SUPPRTF', '', '', '');

    // Gestion des populations
    41951: AglLancefiche('PAY', 'CODEPOPUL_MUL', '', '', '');
    41952: AglLancefiche('PAY', 'POPULATION_MUL', '', '', '');
    41953: AglLancefiche('PAY', 'CONTROLEPOP', '', '', '');
    41962: AglLancefiche('PAY', 'PGMULPARAMETRES', '', '', '');
    41963: AglLancefiche('PAY', 'PARAMETRESASS_MUL', '', '', ''); //PT69
    41965: AglLancefiche('PAY', 'AFFECTSALPOPUL', '', '', ''); //PT71
    41954: AglLancefiche('PAY', 'AFFECTSALPOP_MUL', '', '', ''); //PT79
//    46368: AglLancefiche('PAY', 'POPTOUSMODULES', '', '', ''); ??  Voir Ph. Dumet

    //Module habilitations/populations par popupmenu
    41990: AglLanceFiche('PAY', 'POPTOUSMODULES', '', '', '');

    {------------------------------------------------------------
                                RETROACTIVITE
    -------------------------------------------------------------}
    42111: AGLLanceFiche('PAY', 'MUL_PGRETROACTIV', '', '', '');
    42112: AGLLanceFiche('PAY', 'PGHISTRETROCOT', '', '', '');
    42113: AGLLanceFiche('PAY', 'PGHISTRETROREM', '', '', '');

    {------------------------------------------------------------
                                SALARIES
    -------------------------------------------------------------}
    42120:
      begin
        if not PopUpSalarie then AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'S')
{
        if not PopUpSalarie then
        begin
          if JaiLeDroitTag(200001) then AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'S')
          else AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'ACTION=CONSULTATION');
        end
}
        else AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'ACTION=CONSULTATION');
      end;
    42124: AglLanceFiche('PAY', 'PGHISTOSAL_MUL', '', '', ''); //ancien historique
    42126: AglLanceFiche('PAY', 'PGMULHISTODETAIL', '', '', '');
    42146: AglLanceFiche('PAY', 'MULZONELIBRES', '', '', '');
    42148: AglLanceFiche('PAY', 'MULSALZONELIBRE', '', '', '');
    42147: AglLanceFiche('PAY', 'ANALYSEZONELIBRE', '', '', '');
    42149: AglLanceFiche('PAY', 'PARAMLISTEZL', '', '', '');
    42127: AglLanceFiche('PAY', 'ANALYSE_SAL', '', '', 'HS'); // Analyse historique salarié
    42128: AglLanceFiche('PAY', 'PGEDIT_HISTOSAL', '', '', ''); // Edition historique salarié
    42129: AglLanceFiche('PAY', 'ANALYSE_HISTOSAL', '', '', ''); // Analyse historique salarié (PHD)
    42122: AglLanceFiche('PAY', 'MUL_SALARIEBQ', '', '', '');
    42132: AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=CP', '', 'CP'); // PT-6
    42133: AglLanceFiche('PAY', 'PGREGLES_MUL', '', '', ''); //
    42135: AglLanceFiche('PAY', 'CONGESPAY_MUL', '', '', ';C;;'); // saisie cp salarie
    42134: AglLanceFiche('PAY', 'CALCULACF_MUL', '', '', ''); // CP : calcul du fractionnement
    42139: AglLanceFiche('PAY', 'CLOTURECP_MUL', '', '', ''); // Clôture des CP
    42131: AglLanceFiche('PAY', 'CONGESRAZ_MUL', '', '', ''); // suppression des CP
    //PT9
    42140: AglLanceFiche('PAY', 'MUL_DEPORTSAL', 'GRILLE=BTP', '', 'BTP'); // Gestion de la saisie déportée des compléments BTP
    //FIN PT9
    42151: AglLanceFiche('PAY', 'MUL_DEPORTSAL', 'GRILLE=IS', '', 'IS'); // Gestion saisie intermittents du spectacle
    42152: AglLanceFiche('PAY', 'MUL_ASSEDICSPECTA', 'GRILLE=AS', '', 'AS'); // attestation ASSEDIC spectacle
    42159: AglLanceFiche('PAY', 'MUL_ASSSPECTACLE', '', '', ''); // Edition attestation ASSEDIC spectacle
    42154: AglLanceFiche('PAY', 'MUL_CALCCONGESSPE', '', '', ''); // Calcul Congés spectacles
    42155: AglLanceFiche('PAY', 'CONGESSPEC_MUL', '', '', 'SAISIE'); // Saisie Congés spectacles
    42156: AglLanceFiche('PAY', 'EDIT_CONGESSPEC', '', '', ''); // Edition Congés spectacles
    42157: AglLanceFiche('PAY', 'CONGESSPEC_MUL', '', '', 'FICHIER'); // Génération fichier Congés spectacles
    42158: AglLanceFiche('PAY', 'EDIT_CONGESSPECRE', '', '', 'RECAP'); // Edition Congés spectacles
    42161: AglLanceFiche('PAY', 'MUL_DEPORTSAL', 'GRILLE=MSA', '', 'MSA'); // Gestion saisie MSA
    42162: AglLanceFiche('PAY', 'MSAMULCALCPER', '', '', '');
    42163: AglLanceFiche('PAY', 'MSAMULSAISIEPER', '', '', '');
    42164: AglLanceFiche('PAY', 'MSAMULEVOLUTION', '', '', '');
    42165: AglLanceFiche('PAY', 'MSAMULSAISIEPER', '', '', 'GENERATION');
    42167: AglLanceFiche('PAY', 'EDITIONSMSA', '', '', 'PERIODES'); // Edition MSA
    42168: AglLanceFiche('PAY', 'EDITIONSMSA', '', '', 'EVOLUTIONS'); // Edition MSA
    42169: AglLanceFiche('PAY', 'EDITMSAGLOBAL', '', '', ''); // Edition envoi MSA
    42166: AglLanceFiche('PAY', 'MSA_DECLSALARIES', '', '', '');
    42170: AglLanceFiche('PAY', 'SALARIE_RES', '', '', '');
    42176: AglLanceFiche('PAY', 'PGDUELISTESAI', '', '', ''); // PT65 Gestion de la DUE : Saisie
    //42177: AglLanceFiche('PAY', 'PGDUELISTESAI', '', '', ''); // PT65 Gestion de la DUE : Impression
    //42178: AglLanceFiche('PAY', 'PGDUELISTESAI', '', '', ''); // PT65 Gestion de la DUE : Suppression

    42181: AglLanceFiche('PAY', 'PGREAFFCODEPCS', '', '', '');
    42182: AglLanceFiche('PAY', 'PGREAFFCODEPCS', '', '', 'N');
    42183:
      begin
        ChargePCS(); // Test et chargement de la table de correspondance à partir de la BOB
        AglLanceFiche('PAY', 'MUL_CORRESPCS', '', '', '');
      end;
    42184: AglLanceFiche('PAY', 'DOUBLON_PCS', '', '', '');

    42192: AglLanceFiche('PAY', 'MULHISTOGROUPEE', '', '', '');
    42194: AglLanceFiche('PAY', 'MULELTDYNGROUPEE', '', '', '');
    {------------------------------------------------------------
                          TRAITEMENT DE LA PAIE
    -------------------------------------------------------------}
    //   Gestion des bulletins
    //     40410 : LanceTVsalarie (TRUE);
    42211: AglLanceFiche('PAY', 'MUL_BULLETIN', '', '', ''); // saisie des bullletins
//    42212: AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=A', '', 'SA'); // Saisie des absences
    42212: AglLanceFiche('PAY', 'ABSENCE_MUL', '', '', ';MENU;;PAI'); // consultation des absences par type Absence //DEB PT-3
    //PT61
    42213: AglLanceFiche('PAY', 'MUL_SAISRUB', '', '', 'SRB'); // Saisie par rubrique
    //42213: AglLanceFiche('PAY', 'MUL_SAISRUB', '', '', ''); // Saisie par rubrique
    //Fin PT61
    42214: AglLanceFiche('PAY', 'CONGESGRP_MUL', '', '', 'CP'); // saisie groupee cp
    42215: AglLanceFiche('PAY', 'MULREPRISECP', '', '', ''); // saisie cp salarie
    42217: AglLanceFiche('PAY', 'BULCOMPL_MUL', '', '', ''); // Saisie par rubrique
    //PT61
    42218: AglLanceFiche('PAY', 'MUL_SAISRUB', '', '', 'BCP'); // Saisie par rubrique des bulletins complémentaires
    //Fin PT61
    //     42218 : AglLanceFiche ('PAY','TESTPAY', '', '' , '' );  // saisie cp salarie
    42222: AglLanceFiche('PAY', 'BULLETIN_MUL', '', '', ''); // preparation automantique
    42225: AglLanceFiche('PAY', 'BULLETIN_MUL', '', '', 'X'); // preparation automantique paies aux contrats

    42241: AglLanceFiche('PAY', 'VIREMENT_MUL', '', '', ''); // generation des virements
    42242: AglLanceFiche('PAY', 'EDITCHEQUESSA', '', '', ''); // Edition des chèques des salaires
    42252: AglLanceFiche('PAY', 'ABSENCE_MUL', '', '', ';MENU;;ABS'); // consultation des absences par type Absence //DEB PT-3
    //     42252 : AglLanceFiche ('PAY', 'ABSENCE_MUL', 'TYPEC=T;ACTION=CONSULTATION','',';MENU');  // consultation des absences par type Absence //PT-3
    //     42254 : AglLanceFiche ('PAY','SALARIE_MUL', 'GRILLE=A;ACTION=CONSULTATION', '' , 'CA' ); // Consultation des absences {FIN PT- 3}
    42253: AglLanceFiche('PAY', 'CONGESGRP_MUL', '', '', 'ABS'); // saisie groupee absences
    42256: AglLanceFiche('PAY', 'ABSPAYEES_MUL', '', '', 'ABS'); // Annulation des absences

    42255: AglLanceFiche('PAY', 'PLANNINGABSPAIE', '', '', ''); // planning des absences de la paie
    42257: AglLanceFiche('PAY', 'CALENDRIER_EXCEL', '', '', ''); // Calendrier généré sous excel

    //   IJSS et maintien du salaire   - PT30   - PT32   PT33 PT35
    42811: AglLanceFiche('PAY', 'SUIVIIJSS_MUL', '', '', ''); // Suivi IJSS
    42812: AglLanceFiche('PAY', 'MUL_REGLTIJSS', '', '', ''); // Règlements  IJSS
    42813: AglLanceFiche('PAY', 'STAREGLTIJSS', '', '', ''); // TobViewer sur les Rglts.
    42831: AglLanceFiche('PAY', 'CRITMAINTIEN_MUL', '', '', ''); // Paramétrage des divers règles de maintien
    //    42230: AglLanceFiche('PAY', 'GENEREMAINT_MUL', '', '', ''); // Génération des lignes de maintien
    42822: AglLanceFiche('PAY', 'CONSULTMAINT_MUL', '', '', ''); // Consultation des lignes de maintien
    42823: AglLanceFiche('PAY', 'STAMAINTIEN', '', '', ''); // TobViewer sur les Maintiens.
    42840: AglLanceFiche('PAY', 'LIBRE_ETAT', '', '', ''); // Etats libres PT38

    //   Outils

    42261: AglLanceFiche('PAY', 'CLOTURE', '', '', '');
    42262: AglLanceFiche('PAY', 'DECLOTURE', '', '', '');
    42263: AglLanceFiche('PAY', 'MUL_RAZSALARIE', '', '', '');
    42264: AglLanceFiche('PAY', 'RAZENRIMPORT', '', '', '');
    // PT46 42265: AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', 'O');
    42265: AglLanceFiche('PAY', 'MUL_EXERSOCIAL', 'GRILLE=O', '', '');
    42266: AglLanceFiche('PAY', 'MUL_SUPPBULL', '', '', 'S');
    42267: AglLanceFiche('PAY', 'MUL_SUPPBULL', '', '', 'C');
    {------------------------------------------------------------
                      GESTION DES ACOMPTES
    -------------------------------------------------------------}
    42272:
      begin
        if VH_PAIE.PGRubAcompte <> '' then AglLanceFiche('PAY', 'MUL_ACOMPTE', '', '', '')
        else PGIBox('Vous devez renseigner la rubrique acompte bulletin #13#10 dans les paramètres société #13#10 pour saisir vos acomptes',
            'Gestion des acomptes');
      end;
    42273:
      begin
        if VH_PAIE.PGRubAcompte <> '' then AglLanceFiche('PAY', 'MUL_ACOMPTEGRP', '', '', '')
        else PGIBox('Vous devez renseigner la rubrique acompte bulletin #13#10 dans les paramètres société #13#10 pour saisir vos acomptes',
            'Gestion des acomptes');
      end;
    42274:
      begin
        if VH_PAIE.PGRubAcompte <> '' then AglLanceFiche('PAY', 'GENERACOMPTE', '', '', '')
        else PGIBox('Vous devez renseigner la rubrique acompte bulletin #13#10 dans les paramètres société #13#10 pour générer vos acomptes',
            'Gestion des acomptes');
      end;
    42276: AglLanceFiche('PAY', 'VIREMENTACP_MUL', '', '', '');
    42277: AglLanceFiche('PAY', 'ACOMPTE_ETAT', '', '', '');
    42278: AglLanceFiche('PAY', 'EDITCHEQUESAC', '', '', ''); // Edition des chèques des acomptes
    //PT61
//    42279:
//      begin
//        if VH_PAIE.PGRubAcompte <> '' then AglLanceFiche('PAY', 'MUL_ACOMPTERUB', '', '', '')
//        else PGIBox('Vous devez renseigner la rubrique acompte bulletin #13#10 dans les paramètres société #13#10 pour générer vos acomptes',
//            'Gestion des acomptes');
//      end;
    //Fin PT61
    42289: AglLanceFiche('PAY', 'VENTILCOMPTA_MUL', '', '', '');
    42291: AglLanceFiche('PAY', 'SALDOTCP_MUL', '', '', 'CALC');
    42292: AglLanceFiche('PAY', 'DOTPROVCPRTT_ETAT', '', '', 'PDP');
    42294: AglLanceFiche('PAY', 'DOTPROVCPRTT_ETAT', '', '', 'PRD');
    42295: AglLanceFiche('PAY', 'SALDOTCP_MUL', '', '', '');
    {------------------------------------------------------------
                      GESTION DES EDITIONS
    -------------------------------------------------------------}
    42312: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENU'); // bulletin de paie
    //DEB PT67
    42313: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;ORIGINAL'); // bulletin de paie original
    42314: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;DUPLICATA'); // bulletin de paie duplicata
    42316: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;SPECIMEN'); // bulletin de paie specimen
    //FIN PT67
    42315: AglLanceFiche('PAY', 'BULLPREPA_ETAT', '', '', ''); // Bulletin préparatoire

    //Salarié
    42321: AglLanceFiche('PAY', 'PERSONNEL_ETAT', '', '', '');
    42323: AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=SOLDE', '', 'SOLDE');
    42324: AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CERTIFICAT', '', 'CERTIFICAT');
    42325: AglLanceFiche('PAY', 'MUL_ATTESTASSED', '', '', ''); // attestation assedic
    42326: AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=ACCTRAVAIL', '', 'ACCTRAVAIL');
    42172: AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=ACCTRAVAIL_MSA', '', 'ACCTRAVAIL_MSA'); //PT68
    42327: AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=MALADIE', '', 'MALADIE');
    42173: AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=MALADIE_MSA', '', 'MALADIE_MSA'); //PT68
    42328: AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=CONTRAT', '', 'CONTRAT');
    42329: AglLanceFiche('PAY', 'SALARIE_ATTES', 'OBJET=DUE', '', 'DUE');
    42333: AglLanceFiche('PAY', 'EDIT_ENFANTS', '', '', ''); // Edition des enfants par salariés
    42338: AglLanceFiche('PAY', 'EDIT_FICHESALARIE', '', '', ''); // Edition des fiches salariés
    42339: AglLanceFiche('PAY', 'EDIT_FICHESALARIE', '', '', 'CONTROLE'); // Edition contrôle des fiches salariés
    //42343: AglLanceFiche('PAY', 'MUL_DECLACCTRAV', '', '', ''); // Déclaration accident du travail  PT68
    42343: AglLanceFiche('PAY', 'MUL_DECLACCTRAV', '', '', 'ACT'); // Déclaration accident du travail PT68
    42171: AglLanceFiche('PAY', 'MUL_DECLACCTRAV', '', '', 'MSA'); // Déclaration accident du travail PT68
    42349: AglLanceFiche('PAY', 'LISTE_PERSONNEL_D', '', '', ''); // Liste du personnel présents entre 2 dates

    //ETATS
    42322: AglLanceFiche('PAY', 'FICHEIND', '', '', 'MENU'); // fiche individuelle
    42341: AglLanceFiche('PAY', 'FICHECUMSAL', '', '', '');
    42331: AglLanceFiche('PAY', 'BASECOT_ETAT', '', '', '');
    42332: AglLanceFiche('PAY', 'RECAPPAIE_ETAT', '', '', '');
    42334: AglLanceFiche('PAY', 'PROVISIONCP_ETAT', '', '', '');
    42335: AglLanceFiche('PAY', 'SOLDECP_ETAT', '', '', '');
    42337: AglLanceFiche('PAY', 'SOLDERTT_ETAT', '', '', '');
    42336: AglLanceFiche('PAY', 'CALENDRIER_ETAT', '', '', '');

    //Règlements
    42342: AglLanceFiche('PAY', 'MODEREGLE_ETAT', '', '', '');
    42345: AglLanceFiche('PAY', 'RIB_ETAT', '', '', '');
    42346: AglLanceFiche('PAY', 'ANOAFFECTAUX_ETAT', '', '', '');

    //Declarations periodiques
    42351: AglLanceFiche('PAY', 'EDITMMO', '', '', ''); // Edition des mouvements de main d''oeuvre

    // Initialisation - Consultation/Modification DUCS
// d PT72
    42301: AglLanceFiche('PAY', 'MUL_INITDUCS', '', '', 'INIT');
    42302: AglLanceFiche('PAY', 'IMPRDUCS', '', '', '');
    42303: AglLanceFiche('PAY', 'IMPRDUCS', '', '', 'PVU');
// f PT72
    42353: AglLanceFiche('PAY', 'MUL_DUCSPURGE', '', '', ''); // PT60
    42354: AglLanceFiche('PAY', 'MUL_INITDUCS', '', '', '');
    42355: AglLanceFiche('PAY', 'MUL_DUCSENTETE', '', '', '');

    42356: AglLanceFiche('PAY', 'CHARGESOCIAL_ETAT', '', '', '');
    42357: AglLanceFiche('PAY', 'REDUCTION_ETAT', '', '', 'SALAIRE');
    42358: AglLanceFiche('PAY', 'MUL_DUCSEDI', '', '', '');
    42359: AglLanceFiche('PAY', 'MUL_DUCSRECAP', '', '', ''); // PT23

    //Etat Ventilations analytiques
    42362: AglLanceFiche('PAY', 'VENTILSAL_ETAT', '', '', '');
    42363: AglLanceFiche('PAY', 'VENTILRUB_ETAT', '', '', '');
    42365: AglLanceFiche('PAY', 'ANALYSE_ETAT', '', '', '');
    42366: AglLanceFiche('PAY', 'ANALRUBRIQUE_ETAT', '', '', '');
    42368: AglLanceFiche('PAY', 'TOBWIEW_ANAL', '', '', 'A'); // Analyse detaillee => TobViewer histoanalpaie
    42369: AglLanceFiche('PAY', 'ANALYSE_ACHARGE', '', '', ''); // TobViewer personnes à charge par salarié

    42370: AglLanceFiche('PAY', 'JOURNALPAIE_ETAT', '', '', '');
    42380: AglLanceFiche('PAY', 'LIBRE_ETAT', '', '', '');
    //PT14-2
    42386: AglLanceFiche('PAY', 'ANALYSE_DET', '', '', 'D'); // Analyse detaillee => TobViewer histoRub
    42387: AglLanceFiche('PAY', 'ANALYSE_CUM', '', '', 'C'); // Analyse Cumulee => TobViewer histoCum
    42388: AglLanceFiche('PAY', 'CUBE_DET', '', '', 'D'); // SYNTHESE detaillee => CUBE histoRub
    42389: AglLanceFiche('PAY', 'CUBE_CUM', '', '', 'C'); // SYNTHESE Cumulee => CUBE histoCum
    //FIN PT14-2
    //     45110 : PGCalendAnnuel;//AglLanceFiche('PAY','CALENDANNUEL_ETAT', '', '' , '' );
    42394: AglLanceFiche('CP', 'ETATSCHAINES', '', '', 'PG');  { PT54 }
    42395: AglLanceFiche('PAY', 'SALARIE_RTF', 'OBJET=DOCLIBRE', '', 'DOCLIBRE');

    {------------------------------------------------------------
                          ANALYTIQUE
    ------------------------------------------------------------}
    42410: AglLanceFiche('PAY', 'SALARIE_ANAL', '', '', 'S'); // Liste des salaries sans vantilation analytique
    42415:
      begin // Liste des salaries pour lesquels on va faire une affectation des ventilations analytiques par defaut
        if VH_Paie.PGSectionAnal then AglLanceFiche('PAY', 'SALARIE_ANAL', '', '', 'Y')
        else
          PGIBOX('Pour effectuer l''affectation par défaut,vous devez #13#10 activer la création automatique des sections analytiques dans vos paramètres société',
            'Affectation des ventilations analytiques');
      end;
    42420: AglLanceFiche('PAY', 'REAFFANAL_MUL', '', '', ''); // Réaffectation des ventilations
    42430: AglLanceFiche('PAY', 'SUPPRANAL_MUL', '', '', ''); // Suppression des ventilations analytiques salarié/rubrique

    // d PT28
{------------------------------------------------------------
                     TICKETS RESTAURANT
------------------------------------------------------------}
    42450: AglLanceFiche('PAY', 'TICKETRESTAURANT', '', '', '');
    42452: AglLanceFiche('PAY', 'MUL_SAITICKRESTAU', '', '', '');
    42454: AglLanceFiche('PAY', 'PREPCDETICKET', '', '', '');
    42456: AglLanceFiche('PAY', 'CDETICKETRESTAU', '', '', '');
    42458: AglLanceFiche('PAY', 'TICKETINTEGR', '', '', '');
    42460: AglLanceFiche('PAY', 'ANALYSE_CDETCK', '', '', '');
    42462: AglLanceFiche('PAY', 'ANALYSE_RECAPTCK', '', '', '');
    42464: AglLanceFiche('PAY', 'TICKETPURGE', '', '', '');
    42467: AglLanceFiche('PAY', 'TICKETRESTAURANT', '', '', '');
    42468: ParamTable('PGCODISTRIBUTION', taCreat, 0, PRien);
    42469: AglLanceFiche('PAY', 'MUL_DEPORTSAL', 'GRILLE=TCK', '', 'TCK');
    // f PT28

    {------------------------------------------------------------
                          GESTION DES DADS-U
    ------------------------------------------------------------}
    42511: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'C'); // Multi-critère Calcul standard
    42512: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'B'); // Multi-critère Calcul BTP
    42513: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'E'); // Multi-critère Calcul Elements du salarié
    42515: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'J'); // Multi-critère Calcul inactivité
    42521: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'S'); // Multi-critère Saisie activité
    42522: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'I'); // Multi-critère Saisie inactivité
    42523: AglLanceFiche('PAY', 'MUL_DADS', '', '', 'U'); // Multi-critère Suppression périodes
    42524: AglLanceFiche('PAY', 'DADSLEXIQUE_MUL', '', '', ''); // Multi-critère Lexique DADS-U
    42526: AglLanceFiche('PAY', 'MUL_DADS_ETAB', '', '', ''); // Multi-critère Saisie établissements
    42531: AglLanceFiche('PAY', 'EDITDADSPER', '', '', 'N'); // Edition des périodes DADS
    42535: AglLanceFiche('PAY', 'EDIT_COMPDADS', '', '', ''); // Edition comparatif DADS / fiche individuelle
    42537: AglLanceFiche('PAY', 'EDIT_EXODADS', '', '', ''); // Edition comparatif DADS / fiche individuelle
    42532: AglLanceFiche('PAY', 'EDITDADSPER', '', '', 'B'); // Edition des périodes DADS congés BTP
    42533: AglLanceFiche('PAY', 'EDITDADSPER', '', '', 'S'); // Traitements et salaires payés au salarié
    42534: AglLanceFiche('PAY', 'EDITDADSPER', '', '', 'I'); // Périodes d'inactivité
    42539: AglLanceFiche('PAY', 'EDITDADSCONT', '', '', 'D'); // Edition du contrôle dossier
    42540: AglLanceFiche('PAY', 'MUL_DADS_FICHIER', '', '', 'N'); // Multi-critère préparation normale
    42541: AglLanceFiche('PAY', 'MUL_DADS_FICHIER', '', '', 'C'); // Multi-critère préparation complémentaire
    42551: AglLanceFiche('PAY', 'MUL_DADS_HONOR', '', '', 'S'); // Multi-critère Honoraires
    42552: AglLanceFiche('PAY', 'EDITDADSHON', '', '', ''); // Edition des honoraires
    42553: AglLanceFiche('PAY', 'MUL_DADS_HONOR', '', '', 'D'); // Multi-critère Duplication des honoraires
    42554: AglLanceFiche('PAY', 'DADSHON_RECUPXLS', '', '', ''); // Récupération de données Excel pour création d'honoraires
    42555: AglLanceFiche('PAY', 'MUL_DADS_HONOR', '', '', 'U'); // Multi-critère suppression des honoraires

    {------------------------------------------------------------
                          GESTION DES DADS-B
    ------------------------------------------------------------}
    42562: AglLanceFiche('PAY', 'MUL_DADS2', '', '', 'C');
    42563: AglLanceFiche('PAY', 'MUL_DADS2TIERS', '', '', '');
    42567: AglLanceFiche('PAY', 'MUL_DADS2', '', '', 'S');
    42565: AglLanceFiche('PAY', 'MUL_DADS_ETAB', '', '', ''); // PT103 Saisie établissements
    42568: AglLanceFiche('PAY', 'MUL_DADS2_HONOR', '', '', 'S');
    42569: AglLanceFiche('PAY', 'MUL_DADS2_HONOR', '', '', 'D');
    42571: AglLanceFiche('PAY', 'DADSB_EDITION', '', '', '');
    42576: AglLanceFiche('PAY', 'DADS2_PREP', '', '', '');
    {------------------------------------------------------------
                          GESTION DES PRUD'HOM
    ------------------------------------------------------------}
    42610: AglLanceFiche('PAY', 'PRUDH_DEF', '', '', 'DEF'); // Affectation par defaut des donnes salaries
    42615: AglLanceFiche('PAY', 'PREP_PRUDH', '', '', ''); // Confection du fichier

    {------------------------------------------------------------
                          GESTION DES PROCESSUS
    ------------------------------------------------------------}
    // PT48 début
    42652: AglLanceFiche('PAY', 'PROCESSUSTRT_MUL', '', '', '');
    42653: AglLanceFiche('PAY', 'PROCESSUS_MUL', '', '', '');
    // PT48 fin

    {------------------------------------------------------------
                          RETENUES SALAIRES
    ------------------------------------------------------------}
    42710: AglLanceFiche('PAY', 'RETENUESAL_MUL', '', '', '');
    42761: AglLanceFiche('PAY', 'TYPERETENUE_MUL', '', '', '');
    42730: AglLanceFiche('PAY', 'HISTORETENUE_MUL', '', '', '');
    42741: AglLanceFiche('PAY', 'EDITRETENUESAL', '', '', '');
    42742: AglLanceFiche('PAY', 'EDITECHEANCIER', '', '', '');
    42750: AglLanceFiche('PAY', 'TRANCHEREMRET', '', '', '');
    42762: AglLanceFiche('PAY', 'CALCFRACTION_MUL', '', '', 'RETENUESAL');


    {------------------------------------------------------------
                          GESTION DES CONTRATS
    ------------------------------------------------------------}
    46115: AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=GC', '', 'GC'); // saisie cp salarie
    46120: AglLanceFiche('PAY', 'MUL_ANALCONTRAT', '', '', '');
    46117: AglLanceFiche('PAY', 'EDIT_CONTRATS', '', '', '');
    46119: AglLanceFiche('PAY', 'TOBWIEW_CONTRAT', '', '', '');         //pt57

    {------------------------------------------------------------
                          GESTION DES INTERIMAIRES
    ------------------------------------------------------------}
    46151: AglLanceFiche('PAY', 'INTERIMAIRES_MUL', 'GRILLE=INT', '', 'INT');
    46155: AglLanceFiche('PAY', 'INTERIMAIRES_MUL', 'GRILLE=EMP', '', 'EMP');
    46157: AglLanceFiche('PAY', 'INTERIMAIRES_ETAT', '', '', '');

    {------------------------------------------------------------
                          TRAVAILLEURS HANDICAPES
    ------------------------------------------------------------}
    46211: AglLanceFiche('PAY', 'MUL_HANDICAPE', '', '', '');
    46213: AglLanceFiche('PAY', 'HANDICAPE_EDIT', '', '', '');
    {------------------------------------------------------------
                          MEDECINE DU TRAVAIL
    ------------------------------------------------------------}
    46220: AglLanceFiche('PAY', 'PLANRDVMDT', '', '', '');
    46231: AglLanceFiche('PAY', 'EDITPLANMT', '', '', '');
    46232: AglLanceFiche('PAY', 'EDITMEDTRAV', '', '', '');
    46233: AglLanceFiche('PAY', 'EDITSANSVISITES', '', '', '');
    46234: AglLanceFiche('PAY', 'SALARIE_RTFMED', 'OBJET=MEDECINE', '', 'MEDECINE');
    {------------------------------------------------------------
                          GESTION DE LA PARTICIPATION
    ------------------------------------------------------------}
    46411: AglLanceFiche('PAY', 'MUL_PARPARTICIP', '', '', '');
    46412: ParamTable('PGTYPPARTICIP', taCreat, 0, PRien);
    46421: AglLanceFiche('PAY', 'MUL_SALPARTICIP', '', '', 'PRE');
    46422: AglLanceFiche('PAY', 'MUL_SALPARTICIP', '', '', 'CAL');
    46445: AglLanceFiche('PAY', 'BULPARTICIP_MUL', '', '', ''); // preparation automantique de bulletins complémentaires
    46447: AglLanceFiche('PAY', 'MUL_BULLETIN', '', '', ''); // saisie des bullletins
    46448: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENU'); // bulletin de paie
    46449: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;ORIGINAL'); // bulletin de paie originaux
    46450: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;DUPLICATA'); // bulletin de paie duplicatas
    46451: AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', 'MENUSPEC;SPECIMEN'); // bulletin de paie specimens
    {------------------------------------------------------------
                          GESTION DES SERVICES
    ------------------------------------------------------------}
    46510: AglLanceFiche('PAY', 'ORGANIGRAMME_TV', '', '', '');
    46520: AglLanceFiche('PAY', 'SERVICES_MUL', '', '', '');
    46530: AglLanceFiche('PAY', 'HIERARCHIE_MUL', '', '', '');
    46540: AglLanceFiche('PAY', 'RECHERCHE_RESP', '', '', '');
    46550: AglLanceFiche('PAY', 'EDIT_ORGANIGRAMME', '', '', ''); // Edition organigramme
    {------------------------------------------------------------
                         GESTION DE LA MASSE SALARIALE
    ------------------------------------------------------------}
    46611: AglLanceFiche('PAY', 'EVOLUTIONSAL', '', '', '');
    46612: AglLanceFiche('PAY', 'MODIFCOEFF_MUL', '', '', '');
    46618: AglLanceFiche('PAY', 'MODIFCOEFFSAL_MUL', '', '', '');
    46631: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_REMUNERATION');   //PT47
    46632: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_MASSESALARIALE'); //PT47
    46633: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_PREPARIDR');      //PT47
    46634: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_ANALBRUT');       //PT47
    46635: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_PCHARGES');       //PT47
    46661: AglLanceFiche('PAY', 'EDIT_EFFECTIF', '', '', 'EDIT_SUIVIEFFECTIF');      //PT47
    46662: AglLanceFiche('PAY', 'EDIT_EFFECTIF', '', '', 'EDIT_ENTREESORTIE');       //PT47
    46663: AglLanceFiche('PAY', 'EDIT_EFFECTIF', '', '', 'EDIT_NATIONALITE');        //PT47
    46664: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_SUIAPPRENTI');    //PT47
    46665: AglLanceFiche('PAY', 'EDIT_EFFECTIF', '', '', 'EDIT_ABSENCE');            //PT47
    46666: AglLanceFiche('PAY', 'EDIT_EFFECTIF_2', '', '', '');                      //PT47
    46667: AglLanceFiche('PAY', 'EDIT_REMUNERATION', '', '', 'EDIT_CLASSIF');        //PT52
    46668: AglLanceFiche('PAY', 'EDIT_EFFECTIF', '', '', 'EDIT_PRESENT');            //PT99
    46767: AglLanceFiche('PAY', 'ANALCLASSIF', '', '', ''); // PT45 Analyse classification
    {------------------------------------------------------------
                          ANALYSES stat basees sur TOBVIEWER
    ------------------------------------------------------------}
    46711: AglLanceFiche('PAY', 'ANALYSE_DET', '', '', 'D'); // Analyse detaillee => TobViewer histoRub
    46715: AglLanceFiche('PAY', 'ANALYSE_CUM', '', '', 'C'); // Analyse Cumulee => TobViewer histoCum
    46730: AglLanceFiche('PAY', '', '', '', ''); //
    {------------------------------------------------------------
                          SYNTHESE basees sur CUBE DECISIONNEL
    ------------------------------------------------------------}
    46721: AglLanceFiche('PAY', 'CUBE_DET', '', '', 'D'); // SYNTHESE detaillee => CUBE histoRub
    46725: AglLanceFiche('PAY', 'CUBE_CUM', '', '', 'C'); // SYNTHESE Cumulee => CUBE histoCum
    46751: AglLanceFiche('PAY', 'EFFECTIF_CUB', '', '', 'C'); // SYNTHESE Cumulee => CUBE histoCum
    46752: AglLanceFiche('PAY', 'RUB_EFFECCUB', '', '', 'D'); // SYNTHESE  => CUBE histoBulletin
    46761: AglLanceFiche('PAY', 'EVOLUTIONEFFECTIF', '', '', ''); // PT29
    46763: AglLanceFiche('PAY', 'EVOMASSESALARIALE', '', '', ''); // PT29
    46765: AglLanceFiche('PAY', 'EDITFRAISGENERAUX', '', '', '');
    46781: AglLanceFiche('PAY', 'MUL_EXTRACT', '', '', 'EX');
    46782: AglLanceFiche('PAY', 'PARAMEXTPAIE', '', '', '');
    {------------------------------------------------------------
          ANALYSES et SYNTHESES MULTI DOSSIER Menus spécifiques
    ------------------------------------------------------------}
    46810: AglLanceFiche('PAY', 'ANALMULTI_PAIE', '', '', 'PU'); // Analyse => TobViewer Paieencours jointure salaries
    46812: AglLanceFiche('PAY', 'EFFECTIF_CUBM', '', '', 'C'); // SYNTHESE Cumulee => CUBE histoCum
    46813: AglLanceFiche('PAY', 'RUB_EFFECCUBM', '', '', 'D'); // SYNTHESE  => CUBE histoBulletin
    46814: AglLanceFiche('PAY', 'ANALMULTI_SAL', '', '', ''); // PT43 Tobviewer Salarié Multi_dossiers
    46815: AglLanceFiche('PAY', 'ANALYSE_DETM', '', '', 'D'); // Analyse detaillee => TobViewer histoRub
    46816: AglLanceFiche('PAY', 'ANALYSE_CUMM', '', '', 'C'); // Analyse Cumulee => TobViewer histoCum
    46817: AglLanceFiche('PAY', 'CUBE_DETM', '', '', 'D'); // SYNTHESE detaillee => CUBE histoRub
    46818: AglLanceFiche('PAY', 'CUBE_CUMM', '', '', 'C'); // SYNTHESE Cumulee => CUBE histoCum
    46819: AglLanceFiche('PAY', 'SALARIE_MULDOS', 'GRILLE=S', '', 'S');

    {------------------------------------------------------------
                   BILAN SOCIAL
    ------------------------------------------------------------}
    46911: AglLanceFiche('PAY', 'BILANSIMPLE', '', '', '');
    {  lignes pour tests
         46910 : AglLanceFiche ('PAY','MULEXPORT', '', '' , '' );  // Export GRH specif CEGID
         46920 : AglLanceFiche ('PAY','ANALYSE_REPORT','GRILLE=RESP','', 'M'); // Tobviewer report grh
         46930 : AglLanceFiche ('PAY','CUBE_REPORT','GRILLE=RESP','', 'M'); // Tobviewer report grh
    }

    {------------------------------------------------------------
                          GESTION DES IDR  //PT42B
    ------------------------------------------------------------}
    46313: AglLanceFiche('PAY', 'CALCULSALMOY_MUL', '', '', '');
    46314: AglLanceFiche('PAY', 'SALAIREMOYEN_MUL', '', '', '');
    46316: AglLanceFiche('PAY', 'SALMOYNONCAL_MUL', '', '', '');
    46317: AglLanceFiche('PAY', 'ANALYSE_SALMOYEN', '', '', '');
    46318: AglLanceFiche('PAY', 'CUBE_SALMOYEN', '', '', '');
    46322: AglLancefiche('PAY', 'CALCULSIMUL', '', '', '');
    46324: AglLancefiche('PAY', 'SIMULATION_MUL', '', '', '');
    46326: AglLancefiche('PAY', 'ANALYSE_SIMUL', '', '', '');
    46328: AglLancefiche('PAY', 'CUBE_SIMULATIONID', '', '', '');
    46336: AglLancefiche('PAY', 'EDIT_SIMULATIONID', '', '', '');
    46344: AglLanceFiche('PAY', 'REGLESPOPIDR_MUL', '', '', '');
    46346: AglLanceFiche('PAY', 'CODTAB_MUL', '', '', '');
    46348: AglLanceFiche('PAY', 'CODECALCUL_MUL', '', '', '');

//    46900: AglMontreCubeOlap ('PA1');
    {------------------------------------------------------------
                         MODULE FORMATION MENU 47                     //PT18
    ------------------------------------------------------------}
    //dif
    47311: AglLanceFiche('PAY', 'MUL_COMPTEURSDIF', '', '', '');
    47312: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'DEMANDEDIF');
    47314: AglLanceFiche('PAY', 'EDITRECAPDIF', '', '', '');
    47315: AglLanceFiche('PAY', 'EDITCOMPTEURSDIF', '', '', '');
    47180: AglLanceFiche('PAY', 'MULPARAMFORMABS', '', '', '');
    47317: AglLanceFiche('PAY', 'EDITDEMANDEDIF', '', '', '');
    47318: AglLanceFiche('PAY', 'EDITMVTSDIF', '', '', '');
    // PARAMETRE
    47111: ParamTable('PGFORMATION1', taCreat, 0, PRien);
    47112: ParamTable('PGFORMATION2', taCreat, 0, PRien);
    47113: ParamTable('PGFORMATION3', taCreat, 0, PRien);
    47114: ParamTable('PGFORMATION4', taCreat, 0, PRien);
    47115: ParamTable('PGFORMATION5', taCreat, 0, PRien);
    47116: ParamTable('PGFORMATION6', taCreat, 0, PRien);
    47117: ParamTable('PGFORMATION7', taCreat, 0, PRien);
    47118: ParamTable('PGFORMATION8', taCreat, 0, PRien);
    47119: ParamTable('PGFRAISSALFORM', taCreat, 0, PRien);
    47120: ParamTable('PGMOTIFINSCFORMATION', taCreat, 0, PRien);
    47121: ParamTable('PGMOTIFETATINSC1', taCreat, 0, PRien);
    47122: ParamTable('PGMOTIFETATINSC2', taCreat, 0, PRien);
    47123: ParamTable('PGMOTIFETATINSC3', taCreat, 0, PRien);
    47124: AglLanceFiche('PAY', 'LIEUFORMATION', '', '', '');
    47125: ParamTable('PGPUBLICCONCERNE', taCreat, 0, PRien);

    47128: AglLanceFiche('PAY', 'PGANNUAIRE', '', '', '');
    47131: AglLanceFiche('PAY', 'MUL_FORFAITFORM', '', '', '');
    47132: AglLanceFiche('PAY', 'MUL_PLAFONDFORM', '', '', '');
    47140: AglLanceFiche('PAY', 'EXERFORMATION', '', '', 'EXERCICE');
    47150: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=SAISIECAT', '', 'SAISIECAT');
    47156: AglLanceFiche('PAY', 'CURSUSLISTE', '', '', '');
    47160: AglLanceFiche('PAY', 'LISTESTAGEEMPLOI', '', '', ''); //PT81
    47157: AglLanceFiche('PAY', 'EDIT_CURSUS', '', '', 'CATALOGUE');
    47161: AglLanceFiche('PAY', 'MULSALAIREFORM', '', '', '');
    47162: AglLanceFiche('PAY', 'SALAIREANIM', '', '', 'SALAIREANIM');
    47170: AglLanceFiche('AFF', 'COMPETENCE', '', '', '');


    //MENU WEB ACCESS
    47191: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=CONSULTCAT', '', 'CONSULTCAT');
    // BUDGET
    47210: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=SAISIEBUD', '', 'SAISIEBUD'); // Concultation des formations incluses dans le budget
    //     47220 : AglLanceFiche ('PAY','CONSULTINSCBUD', '', '' , 'SAISIEVALID;BUDGET;;;;ACTION=MODIFICATION' ); // Visualisation des inscriptions
    47220: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'VALIDATION');
    47230: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'CONSULTATION');
    47240: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=INSCBUDGET', '', 'INSCBUDGET'); // inscriptions a partir des formations
    47245: AglLanceFiche('PAY', 'MUL_CURSUS', '', '', '');
    47241: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=BUDGETR', '', 'BUDGETR'); // Visualisation des stages par responsables
    47242: AglLanceFiche('PAY', 'SAISIEINSCBUDGET', '', '', 'SAISIEWEB;BUDGET'); // Saisie responables
    47281: AglLanceFiche('PAY', 'VALID_RESPONSFOR', '', '', 'VALIDRESPONSFOR');
    47282: AglLanceFiche('PAY', 'VALID_RESPCURSUS', '', '', 'VALIDRESPCURSUS');
    47243: AglLanceFiche('PAY', 'CONSULTINSCBUD', '', '', 'CONSULTATION;BUDGET'); // Consultation Rsponsable
{$IFNDEF EMANAGER}
    47251: AglLanceFiche('PAY', 'EDITSUIVIBUDFORM', '', '', '');
    47253: AglLanceFiche('PAY', 'EDITBUDGETFOR', '', '', '');
    47255: AglLanceFiche('PAY', 'EDITBUDGETFOR', '', '', 'INDIVIDUEL');
{$ELSE}
    47251: AglLanceFiche('PAY', 'EDITSUIVIBUDFORM', '', '', 'CWASBUDGET');
    47253: AglLanceFiche('PAY', 'EDITBUDGETFOR', '', '', 'CWASBUDGET');
    47255: AglLanceFiche('PAY', 'EDITBUDGETFOR', '', '', 'INDIVIDUEL');
{$ENDIF EMANAGER}
    //MENU WEB ACCESS
    47261: AglLanceFiche('PAY', 'MUL_STAGE', 'GRILLE=INSCBUDGET', '', 'CWASINSCBUDGET'); // inscriptions a partir des formations
    47263: AglLanceFiche('PAY', 'CONSULTINSCBUD', '', '', 'CWASSAISIEVALID;BUDGET;;;;ACTION=MODIFICATION'); // Visualisation des inscriptions
    47265: AglLanceFiche('PAY', 'VALID_RESPONSFOR', '', '', 'CWASVALIDRESPONSFOR');
    47267: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'CONSULTATION');

    //FORMATION
    47320: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'SAISIE');
    47321: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'SIMPLIFIE');

    47325: AglLanceFiche('PAY', 'PLANNINGFORMATION', '', '', '');

{$IFNDEF EMANAGER}
    47331: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'SAISIEFORMATION');
    47332: AglLanceFiche('PAY', 'VALIDINSSESSION', '', '', 'VALIDINSC');
    47333: AglLanceFiche('PAY', 'MUL_FORMATIONS', '', '', '');
{$ELSE}
    47331: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'CWASSAISIEFORMATION');
    47332: AglLanceFiche('PAY', 'VALIDINSSESSION', '', '', 'CWASVALIDINSC');
    47333: AglLanceFiche('PAY', 'MUL_FORMATIONS', '', '', 'CWASCONSULTATION');
{$ENDIF EMANAGER}
    47341: AglLanceFiche('PAY', 'FEUILLEPRESENCE', '', '', 'PRESENCE');
    47342: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'PRESENCE');
    47351:
      begin
        if not GetParamSocSecur('SO_IFDEFCEGID', FALSE) then AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'FRAIS')
        else AglLanceFiche('PAY', 'MULFRAISSALFORM', '', '', 'SAISIESALARIE');
      end;
    47352: AglLanceFiche('PAY', 'EDITFRAISSAL', '', '', 'CONSULTFRAIS');
    47353: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'CONSULTFRAIS');
    47390: AglLanceFiche('PAY', 'MUL_EXERFORM', '', '', '');
    //MENU WEB ACCESS
    47361: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'CONSULTATION');
    47363: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'CWASSAISIEFORMATION');
    47365: AglLanceFiche('PAY', 'VALIDINSSESSION', '', '', 'CWASVALIDINSC');
    47367: AglLanceFiche('PAY', 'MUL_FORMATIONS', '', '', 'CONSULTATION');
    47370: AglLanceFiche('PAY', 'MUL_SESSIONANIMAT', '', '', 'CONSULTATION');

    // SCORING
    47711: ParamTable('PGFTHEMESCORING', taCreat, 0, PRien);
    47712:
      begin
        St := 'PGZONELIBRE';
        V_PGI.DECombos[TTToNum(St)].Where := 'CC_TYPE="PGZ" AND CC_CODE > "SC" AND CC_CODE < "SCZ"';
        ParamTable('PGZONELIBRE', taModif, 0, PRien);
        V_PGI.DECombos[TTToNum(St)].Where := '';   //PT101
      end;
    47713: ParamTable('PGFSCORING1', taCreat, 0, PRien);
    47714: ParamTable('PGFSCORING2', taCreat, 0, PRien);
    47715: ParamTable('PGFSCORING3', taCreat, 0, PRien);
    47716: ParamTable('PGFSCORING4', taCreat, 0, PRien);
    47717: ParamTable('PGFSCORING5', taCreat, 0, PRien);
    47720: AglLanceFiche('PAY', 'MUL_SCORINGDEF', '', '', '');
    47730: AglLanceFiche('PAY', 'MUL_SCORING', '', '', '');
    47741: AglLanceFiche('PAY', 'EDIT_SCORING', '', '', '');
    47750: AglLanceFiche('PAY', 'CUBE_SCORING', '', '', '');

    // EDITION
    47510: AglLanceFiche('PAY', 'MUL_DECL2483', '', '', '');
    47521: AglLanceFiche('PAY', 'EDITCOLFORMATION', '', '', '47521');
    47522: AglLanceFiche('PAY', 'EDITSTAFOR', '', '', '47522');
    47523: AglLanceFiche('PAY', 'EDITSTAFOR', '', '', '47523');
    47524: AglLanceFiche('PAY', 'ETATLIBREFOR', '', '', '47524');
    47525: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47525');
    47526: AglLanceFiche('PAY', 'EDITHEURESFORM', '', '', ''); //PT82
    47527: AglLanceFiche('PAY', 'EDITINTEGRFOR', '', '', ''); //FLO
    47531: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47531');
    47532: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47532');
    47541: AglLanceFiche('PAY', 'EDITFORANIM', '', '', '47541');
    47542: AglLanceFiche('PAY', 'EDITFORANIM', '', '', '47542');
    47551: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47551');
    47552: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47552');
    47561: AglLanceFiche('PAY', 'EDITFRAISSAL', '', '', 'CONSULTFRAIS');  //PT77
    47562: AglLanceFiche('PAY', 'EDITFRAISSAL', '', '', 'FRAISANIM');     //PT77
    47563: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47563');
    47564: AglLanceFiche('PAY', 'EDITFORSUIVIINV', '', '', 'FISCALETFINANCIER');
    47571: AglLanceFiche('PAY', 'EDITFORREABUD', '', '', '47571');
    47572: AglLanceFiche('PAY', 'EDITFORREABUD', '', '', '47572');
    47573: AglLanceFiche('PAY', 'EDITFORREABUD', '', '', '47573');
    47574: AglLanceFiche('PAY', 'EDITBILANFOR', '', '', '');
    47575: AglLanceFiche('PAY', 'EDITFORREABUD', '', '', '47575');
    47581: AglLanceFiche('PAY', 'EDITFORSUIVIINV', '', '', '');
    47590: AglLanceFiche('PAY', 'EDITCURSUS', '', '', '');

    // ANALYSE
    47610: AglLanceFiche('PAY', 'CUBE_FORMATION', '', '', '');
    47620: AglLanceFiche('PAY', 'CUBE_PREVFORM', '', '', '');
    47630: AglLanceFiche('PAY', 'CUBE_ANIMAT', '', '', '');
    // PASSERELLE
    47801: AglLanceFiche('PAY', 'ENVOISESSION_MUL', '', '', '');
    47802: AglLanceFiche('PAY', 'ENVOIFORM_MUL', '', '', 'MODIF');
    47803: AglLanceFiche('PAY', 'ENVOIFORM_MUL', '', '', 'ENVOI');
    47805: AglLanceFiche('PAY', 'ENVOIFORMFACT_MUL', '', '', 'FACTURE');
    47807: AglLanceFiche('PAY', 'ENVOIFORM_MUL', '', '', 'PURGE');
    // OUTILS
    47911: AglLanceFiche('PAY', 'MUL_SESSIONSTAGE', '', '', 'CALCSALAIRE');
    47913: AglLanceFiche('PAY', 'MUL_STAGE', '', '', 'CALCSALAIRE');
    47921: AglLanceFiche('PAY', 'REAFFFORMATION', '', '', '');
    47923: AglLanceFiche('PAY', 'REAFFPREVFOR', '', '', '');
    47930: AglLanceFiche('PAY', 'RECUPBUDGETFOR', '', '', '');
    47940: AglLanceFiche('PAY', 'REMISEAZEROFORM', '', '', '');
    47950: AglLanceFiche('PAY', 'IMPORTFORMATIONS', '', '', '');

    //EFORMATION
    47401: AglLanceFiche('PAY', 'MUL_STAGECEGID', 'GRILLE=INSCBUDGET', '', 'CWASINSCBUDGET'); // inscriptions a partir des formations
    47402: AglLanceFiche('PAY', 'MUL_CURSUS', '', '', '');
    47403: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'CONSULTATION');
    47405: AglLanceFiche('PAY', 'MUL_INSCFOR', '', '', 'VALIDATION');
//    47405 : AglLanceFiche('PAY', 'VALID_RESPONSFOR', '', '', 'VALIDRESPONSFOR;;;;ACTION=MODIFICATION');
    47406: AglLanceFiche('PAY', 'VALID_RESPCURSUS', '', '', 'VALIDRESPCURSUS');
//    47408 : AglLanceFiche('PAY', 'EDITBUDGETFOR', '', '', 'INDIVIDUEL');
    47408: AglLanceFiche('PAY', 'EDITSUIVIBUDFORM', '', '', '');
    47409: AglLanceFiche('PAY', 'EDITFORMATIONS', '', '', '47531');

    {----------------------------------------------------------
              Gestion des augmentations
    -------------------------------------------------------------}
    303110: AglLanceFiche('PAY', 'MUL_AUGMENTATION', '', '', 'SAISIE');
    303112: AglLanceFiche('PAY', 'MUL_AUGMENTATION', '', '', 'VALIDRESP');
    303114: AglLanceFiche('PAY', 'MUL_AUGMVALIDDRH', '', '', 'VALIDDRH');
    303120: AglLanceFiche('PAY', 'MUL_AUGMINTEGRE', '', '', '');
    303130: AglLanceFiche('PAY', 'EDIT_AUGM', '', '', '');
    303140: AglLanceFiche('PAY', 'ANALYSE_AUGM', '', '', '');
    303150: AglLanceFiche('PAY', 'MUL_AUGMGENERE', '', '', '');
    303160: AglLanceFiche('PAY', 'MUL_AUGMPARAM', '', '', '');
    303170: AglLanceFiche('PAY', 'MUL_AUGMVALIDDRH', '', '', 'SUPPRESSION');
    303180: ParamTable('PGMOTIFAUGM', taCreat, 0, PRien);
    303201: ParamTable('PGTYPEEMBAUCHE', taCreat, 0, PRien);

    {------------------------------------------------------------
                          TRAITEMENT DE FIN D'ANNEE
    -------------------------------------------------------------}

    {------------------------------------------------------------
                          IMPORT/EXPORT
    -------------------------------------------------------------}
    //   Import
{$IFDEF COMPTAPAIE}
    42910: LanceImportISIS; // traitement import ISIS2 ou TDS //PT51
{$ENDIF COMPTAPAIE}
    42915: AglLanceFiche ('PAY', 'IMPORTASC', '', '', '');
    42925: AglLanceFiche('PAY', 'IMPORTFIC', '', '', ''); // traitement d'un fichier import ==> produit externe
    42927: AglLanceFiche('PAY', 'PGMULHISTOIMPORT', '', '', '');
    42928: AglLanceFiche('PAY', 'PGMULHISTOIMPORT', '', '', 'SUPP');

    //   42940 : ExportMvt('FEC') ; // Exportation des écritures de la compta
    42960: AglLanceFiche('PAY', 'MUL_REPRISE', '', '', ''); // saisie des cumuls de reprise
    42961: AglLanceFiche('PAY', 'MUL_REPRISE', '', '', ''); // saisie des cumuls de reprise
    42962: AglLanceFiche('PAY', 'MUL_REPRISECUMUL', '', '', ''); // saisie des cumuls de reprise

    // PT-1 : 06/08/2001 : V547 : PH Menu specifique eAgl  Saisie des absences déportées
    // 27/11/01 Mis en Menu classique car eAGL ne supporte pas les traitements lourds
    43000..43999: Paie_Menus(Num, PRien); { PT54  }

    // Fin PT-1 : 06/08/2001 : V547 : PH
    {----------------------------------------------------------
               MODULE RH AVANCEE BILAN SOCIAL
    -------------------------------------------------------------}
    44110: AglLanceFiche('PAY', 'BSINDCALCUL', '', '', ''); // Calcul du bilan social
    44115: AglLanceFiche('PAY', 'MUL_BSINDCALCULES', '', '', '');

    44210: ParamTable('PGCATBILAN', taCreat, 0, PRien);
    44220: AglLanceFiche('PAY', 'MUL_BSINDICATEUR', '', '', ''); // INDICATEURS Bilan Social
    44230: AglLanceFiche('PAY', 'MUL_BSINDSELECT', '', '', ''); // Sélection des INDICATEURS Bilan Social

    44310: AglLanceFiche('PAY', 'BSINDICATEUR_ETAT', '', '', '');
    44320: AglLanceFiche('PAY', 'BSINDSELECT_ETAT', '', '', '');
    44330: AglLanceFiche('PAY', 'BSINDBILAN_ETAT', '', '', '');
    44340: LancePgAssistExcelOle;
    44350: AglLanceFiche('PAY', 'BSINDBILAN_STA', '', '', '');
    44360: AglLanceFiche('PAY', 'BSINDBILAN_CUB', '', '', '');


    {----------------------------------------------------------
               Gestion des compténces et suivi de carrière
     -------------------------------------------------------------}
    48951: AglLanceFiche('PAY', 'MUL_REFMETIERS', '', '', '');
    48953: ParamTable('PGCLASSIFICATION', taCreat, 0, PRien);
    48955: ParamTable('PGSOUSCLASSIFICATION', taCreat, 0, PRien);

    48961: AglLanceFiche('PAY', 'MUL_STAGEINIT', '', '', '');
    48962: ParamTable('PGFINIVEAU', taCreat, 0, PRien);
    48963: ParamTable('PGFITYPE', taCreat, 0, PRien);
    48964: ParamTable('PGFIDOMAINE', taCreat, 0, PRien);
    {--- ancienne version gestion des compétences
          48101 : AglLanceFiche ('PAY','SALARIE_MUL', 'GRILLE=CT', '' , 'CT' );  // Saisie des competences salaries
          48103 : AglLanceFiche ('AFF','COMPETENCE', '', '' , '' ); // gestion des compétences
          48104 : ParamTable ( 'AFTNIVCOMPETENCE', taCreat, 0, PRien) ; // tablette niveau de compétence
          48105 : ParamTable ( 'AFTNIVEAUDIPLOME', taCreat, 0, PRien) ; // tablette niveau du diplome
          48106 : AglLanceFiche ('PAY','TOBV_COMPETEN', '', '' , 'HS' ); // TobViewer des competences salaries
          48107 : AglLanceFiche ('PAY','CUB_COMPETRESS', '', '' , 'HS' ); // Cube des competences salaries
          48108 : AglLanceFiche ('PAY','GESTCARRIERE', '', '' , '' ); // Edition suivi gestion de carriere
    ---}

    48111: AglLanceFiche('PAY', 'SAL_COMPETENCES', '', '', 'SALARIES'); //Saisie des compétences par salariés
    48112: AglLanceFiche('PAY', 'SAISIEFORMINITIAL', '', '', 'INITIALE'); //Saisie des compétences par salariés
    48120: AglLanceFiche('PAY', 'INT_COMPETENCES', '', '', 'INTERIMAIRES'); //Saisie des compétences par intérimaires/Stagiaires
    48131: AglLanceFiche('PAY', 'EDIT_COMPETINTER', '', '', 'SALARIES'); //Edition des compétences par salariés
    48132: AglLanceFiche('PAY', 'EDIT_INTERCOMPET', '', '', ''); //Edition des salariés par compétences
    48133: AglLanceFiche('PAY', 'EDITEMPLOICOMP', '', '', ''); //Edition des compétences par emploi
    48134: AglLanceFiche('PAY', 'EDITFORMCOMP', '', '', ''); //Edition des compétences par formation
    48135: AglLanceFiche('PAY', 'EDITCOMPSALEMPLOI', '', '', '');
    48136: AglLanceFiche('PAY', 'EDITCOMPEMPLOI', '', '', 'COMPEMPLOI'); //Edition des emplois par compétences
    48137: AglLanceFiche('PAY', 'EDITCOMPETFORM', '', '', 'COMPFORM'); //Edition des formations par compétences
    48138: AglLanceFiche('PAY', 'EDITCPTCESACQUER', '', '', ''); //PT102 Edition des compétences manquantes aux salariés
    48141: AglLanceFiche('PAY', 'PGMUL_RHCOMPET', '', '', ''); //Saisie des compétences
    48142: AglLanceFiche('PAY', 'EMPLOICOMPETENCE', '', '', 'EMPLOI'); //Saisie des compétences par libellé emploi
    48143: AglLanceFiche('PAY', 'OBJECTIFSTAGE', '', '', 'STAGE'); //Saisie des compétences par stage
    48189:
      begin //PT85
        { Contrôle de l'existance des lignes des tables libres }
        { Si les lignes n'existent pas, les créer }
        QVerifTablesLibre := Tob.Create('lignes pour les tables libres', nil, -1);
        QVerifTablesLibre.LoadDetailFromSQL('select * from choixcod where cc_type = "ZLI" and cc_code in ("RH1","RH2","RH3","RH4","RH5") order by cc_code');
        for IndexTableLibre := 1 to 5 do
        begin
          if not Assigned(QVerifTablesLibre.FindFirst(['CC_CODE'],  ['RH'+IntToStr(IndexTableLibre)], False)) then
            ExecuteSQL('INSERT into choixcod(cc_type, cc_code, cc_libelle, cc_abrege, cc_libre) values ("ZLI", "RH'+IntToStr(IndexTableLibre)+'", "Table libre '+IntToStr(IndexTableLibre)+'", "Table Libre '+IntToStr(IndexTableLibre)+'", "")')
        end;
        FreeAndNil(QVerifTablesLibre);

        St := 'GCZONELIBRE';
        V_PGI.DECombos[TTToNum(St)].Where := 'CC_TYPE="ZLI" AND CC_CODE > "RH" AND CC_CODE < "RHZ"';
        ParamTable('GCZONELIBRE', taModif, 0, PRien);
        V_PGI.DECombos[TTToNum(St)].Where := '';   //PT101
      end;
    48199:
      begin //PT85
        { Contrôle de l'existance des lignes des tables libres }
        { Si les lignes n'existent pas, les créer }
        QVerifTablesLibre := Tob.Create('lignes pour les tables libres', nil, -1);
        QVerifTablesLibre.LoadDetailFromSQL('select * from choixcod where cc_type = "ZLI" and cc_code in ("CP1","CP2","CP3") order by cc_code');
        for IndexTableLibre := 1 to 3 do
        begin
          if not Assigned(QVerifTablesLibre.FindFirst(['CC_CODE'],  ['CP'+IntToStr(IndexTableLibre)], False)) then
            ExecuteSQL('INSERT into choixcod(cc_type, cc_code, cc_libelle, cc_abrege, cc_libre) values ("ZLI", "CP'+IntToStr(IndexTableLibre)+'", "Table libre '+IntToStr(IndexTableLibre)+'", "Table Libre '+IntToStr(IndexTableLibre)+'", "")')
        end;
        FreeAndNil(QVerifTablesLibre);

        St := 'GCZONELIBRE';
        V_PGI.DECombos[TTToNum(St)].Where := 'CC_TYPE="ZLI" AND CC_CODE > "CP" AND CC_CODE < "CPZ"';
        ParamTable('GCZONELIBRE', taModif, 0, PRien);
        V_PGI.DECombos[TTToNum(St)].Where := '';  //PT101
      end;
    48191: ParamTable('PGTABLELIBRERH1', taCreat, 0, PRien);
    48192: ParamTable('PGTABLELIBRERH2', taCreat, 0, PRien);
    48193: ParamTable('PGTABLELIBRERH3', taCreat, 0, PRien);
    48194: ParamTable('PGTABLELIBRERH4', taCreat, 0, PRien);
    48195: ParamTable('PGTABLELIBRERH5', taCreat, 0, PRien);
    48196: ParamTable('PGTABLELIBRECR1', taCreat, 0, PRien);
    48197: ParamTable('PGTABLELIBRECR2', taCreat, 0, PRien);
    48198: ParamTable('PGTABLELIBRECR3', taCreat, 0, PRien);
    48211: AglLanceFiche('PAY', 'MUL_GESTCARRIERE', '', '', 'SALARIES');
    48212: AglLanceFiche('PAY', 'EDIT_GESTCARRIERE', '', '', 'SA');
    48220: AglLanceFiche('PAY', 'MUL_CHANGEPOSTE', '', '', 'SA');
    48241: AglLanceFiche('PAY', 'PG_TVSALCOMPET', '', '', 'SA');
    48242: AglLanceFiche('PAY', 'PG_TVINTCOMPET', '', '', 'SI');
    48245: AglLanceFiche('PAY', 'PG_TVGLOBCOMPET', '', '', 'SI');
    48246: AglLanceFiche('PAY', 'PG_TVSALDIP', '', '', 'SI');
    48247: AglLanceFiche('PAY', 'PG_TVCOMPDIP', '', '', 'SI');
    48260: AglLanceFiche('PAY', 'CUBE_COMPET', '', '', 'SI');

    48271: AglLanceFiche('PAY', 'MUL_SUIVIPERSON', '', '', '');
    48272: AglLanceFiche('PAY', 'MUL_ANANENTRET', '', '', 'EAN');
    48273: ParamTable('PGREPONSSUIVI', taCreat, 0, PRien);


    48511: AglLanceFiche('PAY', 'INTERIMAIRES_MUL', '', '', 'CANDIDATS'); // Saisie des candidats
    48515: AglLanceFiche('PAY', 'CAN_COMPETENCES', '', '', 'CANDIDATS'); // Saisie des compétences par candidats
    48516: AglLanceFiche('PAY', 'EDIT_COMPETINTER', '', '', 'CANDIDATS'); //Edition des compétences par candidats
    {----------------------------------------------------------
              Gestion des récrutements et CV 48200
    -------------------------------------------------------------}
    48542: AglLanceFiche('PAY', 'MUL_OFFREEMPLOI', '', '', '');
    48547: ParamTable('PGSUPPORTOFFRE', taCreat, 0, PRien);
    48551: AglLanceFiche('PAY', 'MUL_SUIVIPERSON', '', '', 'SUIV');
    48582: AglLanceFiche('PAY', 'PG_TVBUDGET', '', '', 'BG');
    48583: AglLanceFiche('PAY', 'CUBE_OFFRE', '', '', 'BG');

{$IFNDEF EAGLCLIENT}
    ///// Administration et utilitaires
       ////// Société
    49110: GestionSociete(PRien, @InitSociete, nil); // Gestionnaire société
    49130: RecopieSoc(PRien, True); // recopie société à société
{$ENDIF EAGLCLIENT}
    49140: LanceAssistantPaie; // Assistant de création
    49150: Begin
    		if (V_PGI.ModePcl = '1') then ExecuteSQL ('UPDATE DOSSIER SET DOS_NOMBASE="DB"||DOS_NODOSSIER'); //PT114
    		ParamRegroupementMultiDossier; // Gestion du multi dossier
    	   End;
    //////// Utilisateurs et accès
    49210: AglLanceFiche('PAY', 'LIBRE', '', '', 'REAFFNODOSSIER'); // Lancement utilitaire de controle
    49220: AGLLanceFiche('YY', 'YYUSERGROUP', '', '', ''); {FicheUserGrp } // groupe utilisateurs
    49230:
      begin
        FicheUSer(V_PGI.User);
        ControleUsers;
      end; // Utilisateurs
    49235: begin
           LesModulPaie := '41;42;49';
{$IFNDEF CPS1}
           if (VH_PAIE.PgSeriaAnalyses) then
              LesModulPaie:= LesModulPaie+';46';
           if (VH_PAIE.PgSeriaFormation) then
              LesModulPaie:= LesModulPaie+';47';
           if (VH_PAIE.PgSeriaCompetence) then
              LesModulPaie:= LesModulPaie+';48';
           if (VH_PAIE.PgSeriaBilan) then
              LesModulPaie:= LesModulPaie+';44';
{Demande PH - Suppression présence dans gestion des droits d'accès - VG - 18/09/2007}
       {$IFDEF PRESENCE}
        LesModulPaie := LesModulPaie + ';200;303;310;43;347';
       {$ELSE}
        LesModulPaie := LesModulPaie + ';200;310;43';
       {$ENDIF PRESENCE}
{$ENDIF CPS1}
        AGLLanceFiche('YY', 'YYCONFIDENTIALITE', '', '', LesModulPaie); {FicheUserGrp } // groupe utilisateurs
      end;

    49240: ReseauUtilisateurs(False); // utilisateurs connectés
    49250: VisuLog; // Suivi d'activité
    49260: ReseauUtilisateurs(True); // RAZ connexions
    49270: AglLanceFiche('PAY', 'MUL_PGDROITACCES', '', '', ''); // gestion des acces
    49272: AglLancefiche('PAY', 'MUL_HABILITATIONS', '', '', '');
    49410: ParamFavoris('', '', False, False); // Gestion des favoris
    49413: ShowAdminHdtLinks(Prien); // saisie des tablettes hierarchiques
    49414: ShowTreeLinks(Prien); // saisie des hierarchie dans les tablettes
    49861: AglLAnceFiche ('PAY', 'REMUNER_MULTI',  '', '', '');
    49710: AglLanceFiche('YY', 'YYBUNDLE', '', '', '');
    // fonctionnalités traitements
    49810: AglLanceFiche('PAY', 'NETSERVICE', '', '', '');
    49812: AglLanceFiche('PAY', 'MUL_EVT', '', '', '');
    49813: ParamTable('GCEMPLOIBLOB', taCreat, 0, PRien); // Emploi blob
    49814: AglLanceFiche('PAY', 'EXPORT_SOLDE_CP', '', '', ''); // PT44
    49822: AglLanceFiche('PAY', 'REAFFECTSALARIE', '', '', '');
    49823: AglLanceFiche('PAY', 'REAFFECTTABLE', '', '', '');
    49824: AglLanceFiche('PAY', 'REAFFECTORGANISME', '', '', '');
    49825: AglLanceFiche('PAY', 'RECUP_XLS', '', '', '');
    49881: AglLanceFiche('PAY', 'PAIEBASCULEEURO', '', '', 'E');
    49882: AglLanceFiche('PAY', 'PAIEBASCULEEURO', '', '', 'F');
    49883: AglLanceFiche('PAY', 'BASCULEURO_ETAT', '', '', '');
    // FQ11379    49852: AglLanceFiche('PAY', 'UTILITAIRECP', '', '', 'DECLOTURE');
    49890: AglLanceFiche('YY', 'YYMULJOBS', '', '', '');
    49853: AglLanceFiche('PAY', 'UTILITAIRECP', '', '', 'RELIQUAT');
    49854: AglLanceFiche('PAY', 'UTILITAIRECP', '', '', 'COMPTEUR');
    49855: AglLanceFiche('PAY', 'UTILITAIRECP', '', '', 'VERIFICATION');
    49856: AglLanceFiche('PAY', 'UTILITAIRECP', '', '', 'BASECP');

    49862: AglLanceFiche('PAY', 'CLOTURE', '', '', 'MULTI'); //PT96
    49863: AglLanceFiche('PAY', 'BULLETIN_MUL', '', '', 'MULTI'); //PT97
    49864: AglLanceFiche('PAY', 'VIREMENT_DOS', '', '', ''); //PT100

    49610: AglLanceFiche('PAY', 'GESTION_STD', '', '', 'P');
    49620:
      begin
        AglLanceFiche('PAY', 'GESTION_DOS', '', '', '');
        //           sortiehalley := TRUE;
      end;
{$IFNDEF EAGLCLIENT}
{Suppression des C:\ en dur - VG - 18/09/2007
    49871: AglExportBob('C:\Pgi00\DAT\ELTNAT.BOB', FALSE, False, siFichier, 'C:\Pgi00\DAT\ELTNAT.MDL', TRUE);
    49872: AglExportBob('C:\Pgi00\DAT\PLANPAIE.BOB', FALSE, False, siFichier, 'C:\Pgi00\DAT\PLANPAIE.MDL', TRUE);
}
{$ENDIF EAGLCLIENT}


{$IFDEF WEBSERVICES}
    49865: LanceMajServices(PRien); // WebServices
{$ENDIF WEBSERVICES}

    {--------------------------------------------------------------------------------
       Fonctionnalités de la gestion eAgl des éléments variables
     -----------------------------------------------------------------------------------}

    40110: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P;X'); // Saisie des primes
    40111: AglLanceFiche('PAY', 'A', '', '', ''); // Saisie des primes
    40114: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    40120: AglLanceFiche('PAY', 'PGMULHISTOSAL', '', '', 'P'); // ANAL_HISTPRIMconsultation historique des primes
    49920: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    49921: AglLanceFiche('PAY', 'INTEGPRIM', '', '', ''); // Intégration des primes
    // PT46 49925: AglLanceFiche('PAY', 'EXERCICESOCIAL', '', '', 'V');
    49925: AglLanceFiche('PAY', 'MUL_EXERSOCIAL', 'GRILLE=V', '', '');
    {--------------------------------------------------------------------------------------
       Fonctionnalités de gestion eAGL notament la saisie déportée des absences
       Administration faite sur la base réelle ou base de production
     -------------------------------------------------------------------------------------}
    49910: AglLanceFiche('PAY', 'MUL_DEPORTSALINT', 'GRILLE=RESP', '', 'RESP'); // Gestion des droits salariés pour la saisie déportée
    49990: AglLanceFiche('PAY', 'INTERIMAIRES_MUL', '', '', 'EXTERIEUR');
    49950: AglLanceFiche('PAY', 'ELISTETABLES', '', '', 'E'); // Export des donnees pour la base eAGL

    49952: Salarie_Ressource;
    // traitement d'un fichier import ==> produit externe ou integration des absences validées en saisie déportéé
    49945: AglLanceFiche('PAY', 'IMPORTFIC', '', '', '');
{$IFNDEF EAGLCLIENT}
    // 49980 : ProcPlanningSalarie ('ADM');
{$ENDIF EAGLCLIENT}

// Gestion du Budget (VIENOT Etienne)

    216110 : AglLanceFiche('Q', 'QUFMBPSESSIONENT', '', '', ''); { Session }
    216120 : AglLanceFiche('Q', 'QUFCBPBUDGETPREV', '', '', '');  { Cube }
    216130 : AglLanceFiche('Q', 'QUFSBPBUDGETR3', '', '', '');  { Tableau de bord }
    216141 : AglLanceFiche('Q', 'QUFQBPBUDGET', '', '', 'CONTEXTE=PERIODE'); { Editions }

{$ELSE EABSENCES -0-}
{$IFDEF EABSENCES}
    // PT-1 : 06/08/2001 : V547 : PH Menu specifique eAgl  Saisie des absences déportées
//    43165: AglLanceFiche('PAY', 'SUPERVABSENCES', '', '', 'SAL;' + LeSalarie + ';ECONGES'); // planning des absences

    // Fin PT-1 : 06/08/2001 : V547 : PH

    43000..43999: Paie_Menus(Num, PRien);

{$IFDEF EMANAGER}
    47001..47999: Eformation_Menus(Num, PRien);
    303100..303999: EAugmentation_Menus(Num, PRien);
{$ENDIF EMANAGER}

{$IFDEF etemps}
    263000..263999: Et_Menu(Num);
{$ENDIF etemps}

{$ELSE EABSENCES} //On n'y passe jamais !!!
    40110: AglLanceFiche('PAY', 'MUL_SAISPRIM', '', '', 'P'); // Saisie des primes
    40120: AglLanceFiche('PAY', 'TESTPRIME', '', '', ''); // consultation historique des primes
    40111: AglLanceFiche('PAY', 'INTEGPRIM', '', '', ''); // Intégration des primes
{$ENDIF EABSENCES}

{$ENDIF EABSENCES -0-}
    347000..347999: Presence_Menus(Num, PRien);
  else
    HShowMessage('2;?caption?;' + TraduireMemoire('Fonction non disponible : ') + ';W;O;O;O;', TitreHalley, IntToStr(Num));
  end;
end;

procedure DispatchTT(Num: Integer; Action: TActionFiche; Lequel, TT: string; Range: string);
var LeMode : TActionFiche;
begin
  case Num of
    1: ;
    900: ParamRegroupementMultiDossier(Lequel);
    999: ParamTable(TT, taCreat, 0, nil); {06/08/2003 SB FQ 10058}
    40135: AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
    40190: ParamTable('PGTYPEORGANISME', taCreat, 0, nil);
    40476: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'COE');
    40477: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'IND');
    40478: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'NIV');
    40479: AglLanceFiche('PAY', 'MINICONVENT_VAL', '', '', 'ACTION=CREATION;' + 'QUA');
    40480: begin
             if JaiLeDroitTag(200102) then LeMode := taCreat
               else LeMode := taConsult;
             ParamTable('PGLIBEMPLOI', LeMode, 0, nil);
            end;
    42300: ParamTable('PGQUALDECLARANT', taCreat, 0, nil);
    // Types de voies
    85951: ParamTable('YYVOIENOCOMPL', taCreat, 0, nil, 3);
    85952: ParamTable('YYVOIETYPE', taCreat, 0, nil, 3);

  end;
end;

procedure ZoomEdt(Qui: integer; Quoi: string);

begin
  PgGetZoomPaieEdt(Qui, Quoi);

  Screen.Cursor := crDefault;
end;

procedure ZoomEdtEtat(Quoi: string);
var
  i, Qui: integer;
begin
  i := Pos(';', Quoi);
  Qui := StrToInt(Copy(Quoi, 1, i - 1));
  Quoi := Copy(Quoi, i + 1, Length(Quoi) - i);
  ZoomEdt(Qui, Quoi);
end;

procedure ChangeMenu41;
var
  i: Integer;
  libelle: string;
begin
  if (V_PGI.ModePcl <> '1') then
  begin
    PgAffecteNoDossier(); // Affectation Nodossier en multi-entreprise
    PGRendNoDossier();
  end;
  FMenuG.RemoveItem(41415); // Gestion multitables
  FMenuG.RenameItem(41701, VH_Paie.PgLibCombo1); //Tablette Libre
  FMenuG.RenameItem(41702, VH_Paie.PgLibCombo2);
  FMenuG.RenameItem(41703, VH_Paie.PgLibCombo3);
  FMenuG.RenameItem(41704, VH_Paie.PgLibCombo4);
  FMenuG.RenameItem(41471, VH_Paie.PGLibelleOrgStat1); //Organisation
  FMenuG.RenameItem(41472, VH_Paie.PGLibelleOrgStat2);
  FMenuG.RenameItem(41473, VH_Paie.PGLibelleOrgStat3);
  FMenuG.RenameItem(41474, VH_Paie.PGLibelleOrgStat4);
  FMenuG.RenameItem(41480, VH_Paie.PGLibCodeStat);
  if VH_Paie.PGLibCodeStat = '' then FMenuG.RemoveItem(41480); //Code Stat
  // PT28 if VH_Paie.PGTicketRestau = False then FMenuG.RemoveItem (41487); //PT24
//FQ N°14672
  FMenuG.RemoveItem (-41211); //Paramètres
  FMenuG.RemoveItem(41212); //Paramètrage profils salarié
  FMenuG.RemoveItem(41213); //Paramètrage banque
  FMenuG.RemoveItem(41214); //Paramètrage congés payés
  FMenuG.RemoveItem(41215); //Paramètrage fractionnement
  FMenuG.RemoveItem(41216); //Paramètrage édition bulletin
//FIN FQ N°14672

  for i := 1 to 4 do
  begin
    Libelle := '';
    if i = 1 then libelle := VH_Paie.PGLibelleOrgStat1;
    if i = 2 then libelle := VH_Paie.PGLibelleOrgStat2;
    if i = 3 then libelle := VH_Paie.PGLibelleOrgStat3;
    if i = 4 then libelle := VH_Paie.PGLibelleOrgStat4;
    if (Libelle = '') then FMenuG.RemoveItem(i + 41470);
  end;
  if (VH_Paie.PgNbCombo<5) then
     for i:= VH_Paie.PgNbCombo+1 to 4 do
         begin
         FMenuG.RemoveItem (i+41700);
         end;
     if (VH_Paie.PgNbCombo=0) then
        FMenuG.RemoveItem (-41700);

  //ChargeMenuPop(hm6,FMenuG.Dispatch) ;
  If GetParamSocSecur('SO_PGHISTOAVANCE',False) <> True then FMenuG.RemoveItem(-41430) ;
  if (V_PGI.ModePcl = '1') then FMenuG.RemoveGroup(-41450, TRUE);
end;


// d PT25  Tickets restaurant    PT28

procedure ChangeMenu42;
begin
  if (VH_Paie.PGCodeClient <> VH_Paie.PGCodeRattach) then
  begin
    {confection du fichier final uniquement si client principal}
    FMenuG.RemoveItem(42456);
    {Analyse du récapitulatif uniquement si client principal}
    FMenuG.RemoveItem(42462);
  end;
  if (VH_Paie.PGTicketRestau = False) then
    FMenuG.RemoveGroup(-42450, TRUE);
  // d PT33
  if VH_Paie.PGGestIJSS = False then
  begin
    FMenuG.RemoveGroup(-42800, TRUE);
  end;
  if VH_Paie.PGMaintien = False then
  begin
    FMenuG.RemoveItem(-42820);
  end;
  // f PT33
//SUPPRESSION LIGNE HISTO DATE A DATE POUR V7
  FMenuG.RemoveItem(42191);
  FMenuG.RemoveItem(42193);
  FMenuG.RemoveItem(42133);

  if VH_Paie.PGGESTACPT = False then
  begin
    FMenuG.RemoveGroup(-42270, TRUE);
    FMenuG.RemoveItem(-42270);
  end;
  if VH_Paie.PGAnalytique = FALSE then
  begin // PT49
    FMenuG.RemoveGroup(-42400, TRUE);
    FMenuG.RemoveGroup(-42360, TRUE);
    FMenuG.RemoveItem(-42360);
  end;
  //DEB PT67
  if not GetParamSocSecur('SO_PGGESTORIDUPSPE',False) then
  begin
    FMenuG.RemoveItem(42313);
    FMenuG.RemoveItem(42314);
    FMenuG.RemoveItem(42316);
  end
  else
    FMenuG.RemoveItem(42312);
  //FIN PT67
  //DEB PT76
  if (not GetParamSocSecur('SO_PGHISTORISATION',False)) and (not GetParamSocSecur('SO_PGHISTOAVANCE',False)) then
  begin
    FMenuG.RemoveGroup(-42125, True);
    FMenuG.RemoveItem(-42125);
  end
  else
  begin
    if not GetParamSocSecur('SO_PGHISTORISATION',False) then
    begin
      FMenuG.RemoveItem(42124);
      FMenuG.RemoveItem(42127);
    end;
    if not GetParamSocSecur('SO_PGHISTOAVANCE',False) then
    begin
      FMenuG.RemoveItem(42126);
      FMenuG.RemoveItem(42129);
      FMenuG.RemoveItem(42128);
    end;
  end;
  //FIN PT76
  FMenuG.RemoveItem(42832);  // PT117 Critères de subrogation 
end;
// f PT25

procedure ChangeMenu44;
begin
  FMenuG.RemoveItem(44310);
  FMenuG.RemoveItem(44320);
{$IFNDEF PGTOUTMODULE}
  FMenuG.RemoveItem(44340); // Export excel bilansocial
{$ENDIF}
  FMenuG.RemoveItem(-44400);
  FMenuG.RemoveGroup(-44400, True); // Multidossier bilansocial
  FMenuG.RemoveItem(44410);
end;

//DEB PT67
procedure ChangeMenu46;
begin
  if not GetParamSocSecur('SO_PGGESTORIDUPSPE',False) then
  begin
    FMenuG.RemoveItem(46449);
    FMenuG.RemoveItem(46450);
    FMenuG.RemoveItem(46451);
  end
  else
    FMenuG.RemoveItem(46448);
  if NOT VH_PAIE.PgSeriaIDR then  FMenuG.RemoveGroup(-46300, TRUE) ; // Menu IDR PT108
end;
//FIN PT67

procedure ChangeMenu47; //PT18 Affichage des libellés des tablettes
var
  i: Integer;
  libelle: string;
begin
{$IFDEF EMANAGER}
  FMenuG.RemoveGroup(-47300, TRUE); //Plan de formation
  FMenuG.RemoveGroup(-47700, TRUE); //Scoring
  FMenuG.RemoveGroup(-47600, TRUE); //Analyses
  FMenuG.RemoveGroup(-47100, TRUE); //Paramètres
  FMenuG.RemoveGroup(-47900, TRUE); //Outils
  FMenuG.RemoveGroup(-47800, TRUE); //Passerelle OPCA
  FMenuG.RemoveGroup(-47200, TRUE); //Prévisionnel
  FMenuG.RemoveGroup(-47500, TRUE); //Editions
  LeSalarie := V_PGI.UserSalarie;
  if not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSFOR="' + LeSalarie + '"' +
    'AND PGS_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE)') then
//  If Not ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES LEFT JOIN HIERARCHIE ON PGS_HIERARCHIE=PHO_HIERARCHIE '+
//        'WHERE PHO_NIVEAUH<=2 AND PGS_RESPONSFOR="'+LeSalarie+'"') then
  begin
    FMenuG.RemoveItem(47405);
    FMenuG.RemoveItem(47406);
  end;
{$ELSE}
  FMenuG.RenameItem(47111, VH_Paie.FormationLibre1);
  FMenuG.RenameItem(47112, VH_Paie.FormationLibre2);
  FMenuG.RenameItem(47113, VH_Paie.FormationLibre3);
  FMenuG.RenameItem(47114, VH_Paie.FormationLibre4);
  FMenuG.RenameItem(47115, VH_Paie.FormationLibre5);
  FMenuG.RenameItem(47116, VH_Paie.FormationLibre6);
  FMenuG.RenameItem(47117, VH_Paie.FormationLibre7);
  FMenuG.RenameItem(47118, VH_Paie.FormationLibre8);
  for i := 1 to 8 do
  begin
    if i = 1 then Libelle := VH_Paie.FormationLibre1;
    if i = 2 then Libelle := VH_Paie.FormationLibre2;
    if i = 3 then Libelle := VH_Paie.FormationLibre3;
    if i = 4 then Libelle := VH_Paie.FormationLibre4;
    if i = 5 then Libelle := VH_Paie.FormationLibre5;
    if i = 6 then Libelle := VH_Paie.FormationLibre6;
    if i = 7 then Libelle := VH_Paie.FormationLibre7;
    if i = 8 then Libelle := VH_Paie.FormationLibre8;
    if Libelle = '' then
    begin
      FMenuG.RemoveItem(i + 47110);
    end;
  end;
  if VH_Paie.PGForPrevisionnel = False then
  begin
    FMenuG.RemoveGroup(-47200, TRUE); //Prévisionnel
    FMenuG.RemoveItem(-47570); //Editions comparatif
    FMenuG.RemoveItem(47620); //Analyse
    FMenuG.RemoveGroup(-47155, TRUE); //Gestion des cursus
    FMenuG.RemoveItem(47162); //Salaires animateurs
    FMenuG.RemoveItem(47131); //Forfait frais
    FMenuG.RemoveItem(47913); //Maj coûts
    FMenuG.RemoveItem(47923); //Reaffectation
    FMenuG.RemoveItem(47930); //Récupération
    FMenuG.RemoveItem(47571); //Edit comparatif
    FMenuG.RemoveItem(47572); //Edit comparatif
  end;
  if VH_Paie.PGSCORINGFORM = False then FMenuG.RemoveGroup(-47700, TRUE);
  if VH_Paie.PGForValidSession = False then FMenuG.RemoveItem(47332);
  if VH_Paie.PGForGestionOPCA = False then FMenuG.RemoveGroup(-47800, TRUE);
  if VH_Paie.PGForValidPrev = False then
  begin
    FMenuG.RemoveItem(47220);
    FMenuG.RemoveItem(47250);
    FMenuG.RemoveItem(-47280);   //PT74
  end;
  if (VH_Paie.PGForValidPrev = False) and (VH_Paie.PGForValidSession = False) then
  begin
    FMenuG.RemoveItem(47121);
    FMenuG.RemoveItem(47122);
    FMenuG.RemoveItem(47123);
  end;

  //SCORING
  FMenuG.RenameItem(47713, RechDom('PGZONELIBRE', 'SC1', False));
  FMenuG.RenameItem(47714, RechDom('PGZONELIBRE', 'SC2', False));
  FMenuG.RenameItem(47715, RechDom('PGZONELIBRE', 'SC3', False));
  FMenuG.RenameItem(47716, RechDom('PGZONELIBRE', 'SC4', False));
  FMenuG.RenameItem(47717, RechDom('PGZONELIBRE', 'SC5', False));
  FMenuG.RemoveItem(47553);
  FMenuG.RemoveItem(47582);
  FMenuG.RemoveItem(47742);
  FMenuG.RemoveItem(47743);
  FMenuG.RemoveItem(47573);
{$ENDIF EMANAGER}
  FMenuG.RemoveItem(47950);
end;

procedure ChangeMenu48;
begin
  FMenuG.RenameItem(48191, RechDom('GCZONELIBRE', 'RH1', False));
  FMenuG.RenameItem(48192, RechDom('GCZONELIBRE', 'RH2', False));
  FMenuG.RenameItem(48193, RechDom('GCZONELIBRE', 'RH3', False));
  FMenuG.RenameItem(48194, RechDom('GCZONELIBRE', 'RH4', False));
  FMenuG.RenameItem(48195, RechDom('GCZONELIBRE', 'RH5', False));
  FMenuG.RenameItem(48196, RechDom('GCZONELIBRE', 'CP1', False));
  FMenuG.RenameItem(48197, RechDom('GCZONELIBRE', 'CP2', False));
  FMenuG.RenameItem(48198, RechDom('GCZONELIBRE', 'CP3', False));
{$IFNDEF RHSEUL}
  FMenuG.RemoveGroup(-48400, TRUE); // Gestion des augmentations dans Menu 303 en prevision gestion des budgets
{$ENDIF RHSEUL}
  //DEBUT PT55
  if existeSQL('SELECT MN_TAG FROM MENU WHERE MN_1=48 AND MN_2=1 AND MN_3=4') then
     ExecuteSQL('DELETE FROM MENU WHERE MN_1=48 AND MN_2=1 AND MN_3=4');
  FMenuG.RenameItem(48136, 'Emplois par compétence');
  FMenuG.RenameItem(48137, 'Formations par compétence');
  //FIN PT55
end;

procedure ChangeMenu49;
begin
{$IFDEF RHSEUL}
  FMenuG.RemoveGroup(-49800, TRUE); // Menu traitements
  FMenuG.RemoveGroup(49600, TRUE); // Menu transfert
{$ENDIF}
if not GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) then FMenuG.RemoveItem(49990);
end;

procedure RechargeMenuPopPaie;
begin
  //
end;
//

procedure AfterChangeModule(NumModule: Integer);
var
  CEG, STD, DOS: Boolean;
begin
{$IFDEF ENCOURS}
  VH_Paie.PgeAbsences := TRUE;
{$ENDIF}

  // Nécessaire car les menus sont chargés à chaque fois... !!!! ????
{$IFDEF eTemps}
  ET_Menu(-1);
{$ENDIF}
{$IFDEF EABSENCES}
  if SalAdm = '-' then FMenuG.RemoveGroup(-43800, TRUE); //  désactiver menu administrateur
  if SalVal = '-' then FMenuG.RemoveGroup(43700, TRUE); //  désactiver menu validation absences salarie
  if NotSalAss then FMenuG.RemoveItem(43165);
{$ENDIF}

  if V_PGI.ModePcl = '1' then FMenuG.RemoveItem(49235); // PT39
  FMenuG.RemoveItem(49413);
  FMenuG.RemoveItem(49414);
  // A VOIR pour rajouter les suppressions de lignes de menus spéciques eAgl
  if VH_Paie.PGEcabMonoBase then
  begin //si applicatif monobase, desactive menu import export
    FMenuG.RemoveItem(43810); //  Import des données via ECAB
    FMenuG.RemoveItem(43820); //  Export des absences via ECAB
    FMenuG.RemoveItem(49950); //  Export des données via la paie
  end;
  if not VH_Paie.PgLienRessource then FMenuG.RemoveItem(42170);
{$IFNDEF CPS3}
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
  case NumModule of // Au cas où on change le nom du fichier d'aide
    44: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
    47: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
    48: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
    303: Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';
  end;
{$ENDIF}

  case NumModule of
    41: ChangeMenu41;
    42: ChangeMenu42; // PT25 Tickets restaurant PT28
    44: ChangeMenu44; // PT25 Tickets restaurant PT28
    46: ChangeMenu46; // PT67
    47: ChangeMenu47; // PT18 Affichage des libellés des tablettes
    48: ChangeMenu48;
    49: ChangeMenu49; // PT41 Nettoyage du menu 49
  end;
  /// Menu Pop général
  case NumModule of
    41, 42: ChargeMenuPop(6, FMenuG.DispatchX);
  else ChargeMenuPop(NumModule, FMenuG.DispatchX);
  end;
{$IFDEF ENCOURS}
  VH_Paie.PgSeriaAnalyses := True;
  VH_Paie.PgSeriaDADSB := True;
{$ENDIF}
  if VH_Paie.PGPCS2003 then FMenuG.RemoveItem(-42180);

  if not VH_Paie.PgSeriaAnalyses then // Module Analyse non sérialisé
  begin
    FMenuG.RemoveItem(46710);
    FMenuG.RemoveItem(46720);
    FMenuG.RemoveGroup(46700, TRUE);
    FMenuG.RemoveItem(46810);
    FMenuG.RemoveItem(46820);
    FMenuG.RemoveGroup(46800, TRUE);
  end;

  if not VH_Paie.PgSeriaDADSB then // Module TD Bilatéral non sérialisés
  begin
    FMenuG.RemoveItem(-41816);
    FMenuG.RemoveGroup(-42560, TRUE);
  end;

  AccesPredefini('CEG', CEG, STD, DOS);
  if (not FileExists('UDC.CEG')) then FMenuG.RemoveItem(49210) // désactiver serialisation car dans barre outils
  else FMenuG.RenameItem(49210, 'Utilitaire de contrôle');
{$IFDEF EAGLCLIENT}
    FMenuG.RemoveItem(49110);
    //FMenuG.RemoveItem(49150); //PT114
//  FMenuG.RemoveGroup(49100, TRUE);
  //  FMenuG.RemoveItem(49220);
  //  FMenuG.RemoveItem(49230);
  //  FMenuG.RemoveItem(49240);
  //  FMenuG.RemoveItem(49250);
  //  FMenuG.RemoveItem(49260);
  //  FMenuG.RemoveItem(42910);
  FMenuG.RemoveItem(42336); //Calendrier bimensuel inopérant CWAS
{$ENDIF}
  //  FMenuG.RenameItem(49810, 'S1 et NetService');
    //FMenuG.RemoveItem(43730);   //  Planning
  FMenuG.RemoveItem(49980); //  Planning 41835
  FMenuG.RemoveItem(41835); //  Planning 41835
  //ANCIEN MENU ETEMPS
  FMenuG.RemoveItem(43210); // Gestion Etemps collaborateur
  FMenuG.RemoveGroup(-43240); // Gestion Etemps responsable
  FMenuG.RemoveItem(-43240); // Gestion Etemps responsable
  FMenuG.RemoveItem(41651);   //  Taux Transport : fct à développer (MF)



{$IFNDEF ENCOURS}
  //Spécificité CEGID SA
  if not GetParamSocSecur('SO_IFDEFCEGID', FALSE) then
  begin
    FMenuG.RemoveItem(42257); //  Calendrier excel
  end;
  FMenuG.RemoveItem(42390); //  Calendrier annuel
  FMenuG.RemoveItem(42395); //  Point d'entrée
  //FMenuG.RemoveItem(42255); //  Planning des absences
  FMenuG.RemoveItem(49852); //  Annulation clôture CP
  FMenuG.RemoveItem(49853); //  Génération reliquat
  //PT47 if VH_Paie.PGMODIFCOEFF = False then FMenuG.RemoveGroup(-46600, TRUE); // Evolution de modif
  if VH_Paie.PGMODIFCOEFF = False then FMenuG.RemoveGroup(-46610, TRUE); // Evolution de modif  PT47
  if (not VH_Paie.PgeAbsences) and (not VH_Paie.PgSeriaFormation) then FMenuG.RemoveGroup(-46500, TRUE); // organisation
  if not VH_Paie.PGSaisieArret then FMenuG.RemoveGroup(-42700, TRUE); // saisie arret
  FMenuG.RemoveItem(42153); //  Congés spectacles
  // Accès compatibilité supprimé
  //FMenuG.RemoveItem(41710);
//  FMenuG.RenameItem(41715, 'Longueur des comptes');
//  FMenuG.RemoveItem(41720);
  FMenuG.RemoveItem(41730);

  FMenuG.RemoveItem(41491); //  Motif Entree
  FMenuG.RemoveItem(41492); //  Motif sortie
  //FMenuG.RemoveItem(-41840);   //  PT14  (ducsEdi)
  //FMenuG.RemoveItem(42358);   //  PT14 (ducsEdi)

  if VH_Paie.PGBTP = False then
  begin
    FMenuG.RemoveItem(42140); //  Compléments BTP salarié
    FMenuG.RemoveItem(42512); //  Calcul dads-u CP BTP
    FMenuG.RemoveItem(42532); //  Edition DADS-U CP BTP
  end;

{$IFNDEF VG}
  FMenuG.RemoveItem(42514); //  DADS-U Assurances
  FMenuG.RemoveItem(42524); //  Multi-critère Lexique DADS-U
  FMenuG.RemoveItem(42525); //  Eléments de cotisation
{$ELSE}
  FMenuG.RenameItem(42524, 'Lexique DADS-U');
{$ENDIF VG}

 if VH_Paie.PGMODIFCOEFF = False then
    FMenuG.RemoveItem(-46610); // Evolution du coefficient PT47

  //PT14-2
{$IFDEF CCS3}
  //Seuls les menus 41, 42 et 49 sont concernés
  FMenuG.RemoveItem(41270); //Banques des organismes                  OK
  FMenuG.RemoveItem(-41816); //TD Bilatéral                           OK
  FMenuG.RemoveItem(-41840); //Ducs-EDI                               OK
  FMenuG.RemoveItem(-41860); //Congés spectacles                      OK
  FMenuG.RemoveItem(-41870); //MSA                                    OK
  FMenuG.RemoveItem(-42150); //Intermittents                          OK
  FMenuG.RemoveItem(42162); //EDI-Calcul des périodes
  FMenuG.RemoveItem(42163); //EDI-Saisie des périodes
  FMenuG.RemoveItem(42164); //EDI-Saisie des évolutions
  FMenuG.RemoveItem(42165); //EDI-Préparation du fichier
  FMenuG.RemoveItem(42167); //EDI-Edition des périodes
  FMenuG.RemoveItem(42168); //EDI-Edition des évolutions
  FMenuG.RemoveItem(42217); //Création bulletin complémentaire        OK
  FMenuG.RemoveItem(-42290); //Dotations provisions CP et RTT         OK
  FMenuG.RemoveItem(42343); //Déclaration accident du travail         OK
  FMenuG.RemoveItem(42171); //Déclaration accident du travail MSA     OK PT68
  FMenuG.RemoveItem(42172); //Attestation accident du travail MSA     OK PT68
  FMenuG.RemoveItem(42173); //Attestation maladie MSA                 OK PT68
  FMenuG.RemoveItem(42351); //Déclaration MO                          OK
  FMenuG.RemoveItem(42358); //Ducs-EDI                                OK
  FMenuG.RemoveItem(42380); //Etat libre                              OK
  FMenuG.RemoveItem(42394); //Etats chaînés                           OK
  FMenuG.RemoveGroup(-42450, TRUE); //Titres restaurant               OK
  FMenuG.RemoveItem(42554); //Import fichier Excel                    OK
  FMenuG.RemoveGroup(-42560, TRUE); //TD bilatéral                    OK
  FMenuG.RemoveGroup(-42700, TRUE); //saisie arret                    OK
  FMenuG.RemoveGroup(-42800, TRUE); //IJSS & maintien du salaire      OK
  FMenuG.RemoveItem(49810); //NetExpert                               OK
  FMenuG.RemoveItem(-49880); //Basculement en Euro                    OK
  FMenuG.RemoveGroup(49900, TRUE); //ePaie                            OK
{$ELSE}
//  FMenuG.RemoveItem(-42385); //Analyse et Synthèse PH pour laisser un bout des analyses
{$ENDIF CCS3}
  //FIN PT14-2
  // Suppression des lignes de menu concernant le module recrutement
  FMenuG.RemoveItem(-48270); // Gestion des entretiens
  FMenuG.RemoveGroup(-48500, TRUE); // Recrutement
  FMenuG.RemoveGroup(-46400, TRUE); // Participation
  FMenuG.RemoveItem(-42280); // Frais professionnel
  FMenuG.RemoveItem(42840); // Etat libre IJSS
  // Suppression des lignes de menu concernant la TOX
  FMenuG.RemoveGroup(49300, TRUE);
  FMenuG.RemoveItem(49865); //  Web Services en devenir en case 30ou 31
  // if NOT V_PGI.DebugSQL then FMenuG.RemoveItem  (49861);  // Transfert dossier
  if not V_PGI.DebugSQL then FMenuG.RemoveGroup(49400); // menu hiérarchie
{$IFDEF EAGLCLIENT}
  FMenuG.RemoveItem(42394); // En CWAS, non applicable
{$ENDIF EAGLCLIENT}
  if Not VH_Paie.PgeAbsences then
    FMenuG.RemoveGroup(49900, TRUE) // EPaie si non séria
  else
    if VH_Paie.PGEcabMonoBase then
    begin //si applicatif monobase, desactive menu import export
      FMenuG.RemoveItem(43810); //  Import des données via ECAB
      FMenuG.RemoveItem(43820); //  Export des absences via ECAB
      FMenuG.RemoveItem(49950); //  Export des données via la paie
    end;
  //   PT11 30/04/2002 V582 PH Activation désactivation des lignes de menus aux différents cas gérés
  // Gestion des intérimaires
  if not VH_Paie.PGInterimaires then
    FMenuG.RemoveGroup(-46150, TRUE);

  // Gestion MSA
  if not VH_Paie.PGMsa then
  begin
    FMenuG.RemoveItem(-42160);
    FMenuG.RemoveItem(41485); // PT22
    FMenuG.RemoveItem(42171); // PT68
    FMenuG.RemoveItem(42172); // PT68
    FMenuG.RemoveItem(42173); // PT68
  end;

  // Gestion des intermittants du spectacle
  if not VH_Paie.PGIntermittents then
    FMenuG.RemoveItem(-42150);

  // Gestion des Services
{$IFDEF SERVICE}
  VH_Paie.PGRESPONSABLES := True;
{$ENDIF SERVICE}
  if not VH_Paie.PGRESPONSABLES then
    FMenuG.RemoveGroup(-46500, TRUE);
  // FIN PT11


{$ENDIF ENCOURS}

    FMenuG.RemoveItem(49814); // PT75 FQ12710
{$IFNDEF RMA}                 // PT75 a supprimer quand DUE à livrer
    FMenuG.RemoveItem(-42175);
    FMenuG.RemoveItem(-41880);
{$ENDIF RMA}

  //FORMATION
  FMenuG.RemoveItem(47170);
  FMenuG.RemoveItem(-47190);
  {futur menu V6}

  //SB 15/09/03 FQ 10781 Affichage module CP si gestion CP Paramètre Soc
  if not VH_PAie.PGCongesPayes then
  begin
    FMenuG.RemoveItem(-42130); //Module Gestion par salarié
    FMenuG.RemoveItem(42334); //Provision CP
    FMenuG.RemoveItem(42335); //Solde CP
  end;

  if not VH_PAie.PGAbsence then //Module Absence
  begin
    FMenuG.RemoveItem(-42250); //Module absence
    FMenuG.RemoveItem(42252); //Gestion par salarié
    FMenuG.RemoveItem(42253); //Saisie groupée absence
    FMenuG.RemoveItem(42253); //Saisie groupée absence
    FMenuG.RemoveItem(42212); //Saisie des absences
  end;

  if (not VH_PAie.PGCongesPayes) and (not VH_PAie.PGAbsence) then
  begin
    FMenuG.RemoveItem(41495); //Gestion Motif absence
  end;

  //FMenuG.RemoveItem (49413); // tablettes hierarchiques
  //FMenuG.RemoveItem (49414); // tablettes hierarchiques

  if (V_PGI.Driver = dbORACLE7) then FMenuG.RemoveItem(42337);
  if VH_Paie.PGPCS2003 then
  begin // PT34
    FMenuG.RemoveItem(41171);
    FMenuG.RemoveItem(41172);
    FMenuG.RemoveItem(41174);
  end;
  // d PT30 - IJSS
  if VH_Paie.PGGestIJSS = False then
  begin
    FMenuG.RemoveGroup(-42800, TRUE); // PT33
  end;
  // f PT30 - IJSS
  // d PT32 - Maintien
  if VH_Paie.PGMaintien = False then
  begin
    FMenuG.RemoveItem(-42820); // PT33
  end;
  // f PT32 - Maintien

  // Menu multi dossier n'existe pas en PCL du moins dans la V6
// Activation pour test des fonctionnalités multi-société en PCL
//  if (not FileExists('UDC.CEG')) then
//  begin
  if (V_PGI.ModePCL = '1') And Not GetParamSocSecur('SO_PGTRTMULTIDOS', False) then //PT114
  begin
    FMenuG.RemoveItem(49150);
    FMenuG.RemoveItem(-46800);
  end;
//  end;

//PT62
// d PT86
  if not GetParamSocSecur('SO_PGTDBILATERAL',False) then
  begin
    FMenuG.RemoveGroup(-42560, TRUE);
  end;
  if not GetParamSocSecur('SO_PGPRUDHOMME',False) then
  begin
    FMenuG.RemoveGroup(-42600, TRUE);
  end;
// f PT86
//Fin PT62
  //DEB PT66
  if not GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
//    FMenuG.RemoveItem(41103);
//    FMenuG.RemoveItem(41106);   //PT70
    FMenuG.RemoveItem(41104);
  end;
  //FIN PT66

  //PT98
  If Not GetParamSocSecur('SO_PGTRTMULTIDOS', False)  Then
  Begin
      FMenuG.RemoveGroup(-49860,TRUE);
      FMenuG.RemoveItem(-49860);
  End;
end;

procedure AfterProtec(sAcces: string); //Sérialisation des modules
var i: Integer;
begin
    //If VH_PAIE.ModeFiche Then Exit; //PT98
    If Not ChargementComplet Then Exit; //PT100

// cas où on coche "version non sérialisée' dans
// l'écran de séria. On force la non sérialisation   47, 48, 44,
if (V_PGI.NoProtec) then
   V_PGI.VersionDemo:= TRUE
else
   V_PGI.VersionDemo:= False;
if V_PGI.ModePCL <> '1' then
   begin
{$IFNDEF CPS3}
   if (sAcces[1]='-') and (sAcces[2]='-') and (sAcces[3]='-') and
      (sAcces[4]='-') and (sAcces[5]='-') then
      V_PGI.VersionDemo:= TRUE;

   if V_PGI.VersionDemo then
      begin
      MemS:= 6;
      VH_Paie.PgSeriaPaie:= TRUE;
      VH_Paie.PgSeriaAnalyses:= TRUE;
      VH_Paie.PgSeriaFormation:= TRUE;
      VH_Paie.PgeAbsences:= FALSE;
      VH_Paie.PgSeriaDADSB:= TRUE;
      SeriaGedPcl:= TRUE; // PT36
//      VH_Paie.SeriaDiodePcl:= FALSE; // PT50
      VH_Paie.PgSeriaBilan:= True; // PT31
      VH_Paie.PgSeriaCompetence:= TRUE; // PT31
      VH_Paie.PgSeriaMasseSalariale:= True;
      VH_Paie.PgSeriaPresence:= True;
      VH_PAIE.PgSeriaIDR:= TRUE; //PT78
      end
   else
      begin
      MemS:= 0;
      for i := 2 to 7 do // Si 6 alors licence unlimited
          begin
          if (sAcces[i] = 'X') then
             MemS:= i-1;
          end;
      VH_Paie.PgSeriaPaie:= (sAcces[1]='X') or (sAcces[2]='X') or
                            (sAcces[3]='X') or (sAcces[4]='X') or
                            (sAcces[5]='X') or (sAcces[6]='X') or
                            (sAcces[7]='X');
      VH_Paie.PgSeriaAnalyses:= (sAcces[8]='X') or (sAcces[9]='X') or
                                (sAcces[10]='X') or (sAcces[11]='X') or
                                (sAcces[12]='X') or (sAcces[13]='X');
      VH_Paie.PgSeriaFormation:= (sAcces[14]='X') or (sAcces[15]='X') or
                                 (sAcces[16]='X') or (sAcces[17]='X') or
                                 (sAcces[18]='X') or (sAcces[19]='X');
      VH_Paie.PgSeriaCompetence:= (sAcces[20]='X') or (sAcces[21]='X') or
                                  (sAcces[22]='X') or (sAcces[23]='X') or
                                  (sAcces[24]='X') or (sAcces[25]='X');
      VH_Paie.PgeAbsences:= (sAcces[26]='X') or (sAcces[27]='X') or
                            (sAcces[28]='X') or (sAcces[29]='X') or
                            (sAcces[30]='X') or (sAcces[31]='X');
      VH_Paie.PgSeriaDADSB:= TRUE;
      VH_Paie.PgSeriaBilan:= (sAcces[32]='X') or (sAcces[33]='X') or
                             (sAcces[34]='X') or (sAcces[35]='X') or
                             (sAcces[36]='X') or (sAcces[37]='X');
      VH_Paie.PgSeriaIDR := (sAcces[38]='X');
      VH_Paie.PgSeriaMasseSalariale:= (sAcces[39]='X');
      VH_Paie.PgSeriaPresence:= (sAcces[40]='X');
      end;
{$ELSE}
   if (sAcces[1]='-') and (sAcces[2]='-') then
      V_PGI.VersionDemo:= TRUE;

   if V_PGI.VersionDemo then
      begin
      MemS:= 2;
      VH_Paie.PgSeriaPaie:= TRUE;
      VH_Paie.PgSeriaAnalyses:= TRUE;
      VH_Paie.PgSeriaFormation:= TRUE;
      VH_Paie.PgeAbsences:= FALSE;
      VH_Paie.PgSeriaDADSB:= TRUE;
      VH_PAIE.PgSeriaIDR := TRUE; //PT78
      end
   else
      begin
      MemS:= 0;
      for i := 0 to 1 do // Détection du nombre de licence salariés
          begin
          if (sAcces[i+1]='X') then
             MemS:= i+1;
          end;
      VH_Paie.PgSeriaPaie:= (sAcces[1]='X') or (sAcces[2]='X');
      VH_Paie.PgSeriaFormation:= (sAcces[3]='X') or (sAcces[4]='X');
      VH_Paie.PgSeriaAnalyses:= (sAcces[5]='X') or (sAcces[6]='X');
      VH_Paie.PgeAbsences:= (sAcces[7]='X') or (sAcces[8]='X');
      VH_Paie.PgSeriaDADSB:= TRUE;
      VH_Paie.PgSeriaIDR := (sAcces[9]='X');
      end;
{$ENDIF}
  end
  else
  begin // Agenda PCL
    if (sAcces[1] = '-') then V_PGI.VersionDemo := TRUE;
    SeriaMessOK := False;
{$IFNDEF RHSEUL}
    VH_Paie.PgSeriaPaie := (sAcces[1] = 'X');
    VH_Paie.PgSeriaAnalyses := (sAcces[2] = 'X');
    VH_Paie.PgSeriaFormation := (sAcces[3] = 'X');
    VH_Paie.PgeAbsences := (sAcces[4] = 'X');
    VH_Paie.PgSeriaDADSB := TRUE;
//    SeriaGedPcl := (sAcces[6] = 'X'); // PT36
    VH_Paie.PgSeriaBilan := (sAcces[5] = 'X'); // PT31
    VH_Paie.PgSeriaCompetence := (sAcces[6] = 'X'); // PT31
    if (sAcces[7] = 'X') then SeriaMessOK := TRUE;
    if (sAcces[8] = 'X') then SeriaGedPcl := TRUE;
//    VH_Paie.SeriaDiodePcl := (sAcces[9] = 'X'); // PT50
    VH_Paie.PgSeriaIDR := (sAcces[9] = 'X')
{$ELSE}
    VH_Paie.PgSeriaPaie := (sAcces[1] = 'X');
    VH_Paie.PgSeriaFormation := (sAcces[2] = 'X');
    VH_Paie.PgSeriaCompetence := (sAcces[4] = 'X');
    VH_Paie.PgSeriaBilan := (sAcces[3] = 'X');
    if (sAcces[5] = 'X') then SeriaMessOK := TRUE;
    if (sAcces[6] = 'X') then SeriaGedPcl := TRUE;
{$ENDIF}
  end;
{$IFDEF CPS1}
    if (sAcces[1] = '-') AND (sAcces[2] = '-')then V_PGI.VersionDemo := TRUE;
    Paie1550 := 15;
    if V_PGI.VersionDemo then
    begin
      MemS := 3;
      VH_Paie.PgSeriaPaie := FALSE;
    end
    else
    begin
      if sAcces[2] = 'X' then
      begin
        MemS := 2;
        Paie1550 := 50;
      end
      else if sAcces[1] = 'X' then MemS := 1
      else Mems := 3;
      if (Mems = 1) OR (Mems = 2) then VH_Paie.PgSeriaPaie := TRUE;
    end;
{$ENDIF}

{$IFDEF PGFORCESERIA}
  VH_Paie.PgeAbsences := TRUE;
  VH_Paie.PgSeriaFormation := TRUE;
  VH_Paie.PgSeriaAnalyses := TRUE;
  VH_Paie.PgSeriaDADSB := TRUE;
  VH_Paie.PgSeriaCompetence := TRUE;
  VH_Paie.PgSeriaBilan := TRUE;
  VH_Paie.PgSeriaMasseSalariale:= True;
  VH_Paie.PgSeriaPresence:= True;
  VH_Paie.PgSeriaIDR := TRUE;
{$ENDIF}
  //PT14-2
{$IFNDEF CPS3}
{$IFNDEF EABSENCES}
{$IFNDEF EPRIMES}
{$IFNDEF EMANAGER}
  ChargeModules_PAIE;
{$ENDIF EMANAGER}
{$ENDIF}
{$ENDIF}
{$ELSE}
{$IFDEF CPS1}
  FMenuG.SetModules([42, 41, 46, 49], [42, 41, 46, 49]);
{$ELSE}
  FMenuG.SetModules([42, 41, 46, 47, 43, 49,310], [42, 41, 46, 47, 43, 49,0]);
{$ENDIF}
{$ENDIF}
  //FIN PT14-2

{$IFDEF ETEMPS}
  V_PGI.VersionDemo := False;
{$ENDIF}
{$IFDEF EABSENCES}
{$IFDEF EMANAGER}
  FMenuG.SetModules([43,914, 915], [43,-1,-1]);
{$ELSE}
  FMenuG.SetModules([43], [43]);
{$ENDIF}
{$ENDIF}

{$IFDEF PGTOUTMODULE}
  FMenuG.SetModules([41, 42, 43, 44, 46, 47, 48, 49, 303, 347], [41, 42, 43, 44, 46, 47, 48, 49, 3, 80]);
{$ENDIF}

end;

procedure AGLAfterChangeModule(Parms: array of variant; nb: integer);
begin
  AfterChangeModule(Integer(Parms[0]));
  ProcCalcEdt := CalcOLEEtatPG;
  ProcCalcMul := PGCalcMul;
end;
{ Fonction qui recupere le numero de Salarie dans le cas d'une connection NT avec
  de façon à savoir, s'il est adminstrateur et/ou responsable validation absences
}

function PgAnalConnect: boolean;
var
  Q: TQuery;
  LaChaine, LeResp: string;
begin
  LaChaine := VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl:= VerifieCheminPG( Q.FindField('SOC_DATA').AsString);
    Ferme(Q);
  end;
  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt119  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') And  (Length(LeSalarie) < 10) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie; // pt119
  if LeSalarie = '' then
  begin
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', 'Saisie des absences');
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString;
      SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
      result := true;
    end
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      Ferme(Q);
      result := false;
      exit;
    end;
    LeResp := Q.FindField('PSE_RESPONSABS').AsString;
    Ferme(Q);
    if LeResp = '' then
    begin
      PgiBox('Vous n''avez pas de responsable affecté !#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
      result := false;
      exit;
    end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    //PT115 - Début
    else
      begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         Begin
           LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
         End
    //PT115 - Fin
      else
      begin
        PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', 'Saisie des absences');
        Ferme(Q);
        result := false;
        exit;
      end;
      Ferme(Q);
    end;
  end; //PT115
  // Desactivation des fonctionnalités administrateur ou responsable validation selon le salarié
  if SalAdm <> 'X' then FMenuG.RemoveGroup(-43800, TRUE); //  désactiver menu administrateur
  if SalVal <> 'X' then FMenuG.RemoveGroup(43700, TRUE); //  désactiver menu validation absences salarie
  //Salarié secrétaire
  NotSalAss := not (ExisteSql('SELECT PSE_ASSISTABS FROM DEPORTSAL WHERE PSE_ASSISTABS="' + LeSalarie + '"'));
  if NotSalAss then
  begin
    FMenuG.RemoveGroup(43200, TRUE); { PT1 FMenuG.RemoveItem(43165); }
    FMenuG.RemoveItem(43165);
  end;

  result := true;
  //BaseEPaie := (GetParamSoc ('SO_PGBASEEPAIE'));

{$IFDEF etemps}
  // XP le 24-07-2003
  ET_Menu(-1);
  if ldSaisieTemps in LeTemps.MesDroits then
  begin
    FMenuG.ChangeModule(263);
    ET_Menu(263001);
  end
  else
  begin
    if result and (SalAdm <> 'X') and (SalVal <> 'X') then
      AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
  end;
{$ELSE}
  if result and (SalAdm <> 'X') and (SalVal <> 'X') then
    AglLanceFiche('PAY', 'EABSENCE_MUL', '', '', 'SAL');
{$ENDIF}
end;
// PT16  : 11/06/2002 VDev PH Gestion des controles avec saisie des primes
{ Fonction qui recupere le numero de Salarie dans le cas d'une connection NT avec
  de façon à savoir, s'il est adminstrateur et/ou responsable validation des primes
}
{***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/11/2002
Modifié le ... : 15/11/2002
Description .. : Traitement par lots
Mots clefs ... : PAIE;LOT
*****************************************************************}

function PgAnalConnectVar: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
  BaseEPaie: Boolean;
begin
  Lib := 'Saisie ' + VH_Paie.PGLibSaisPrim;
  LaChaine := VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
    Ferme(Q);
  end;
  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt119  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
   if (VH_Paie.PgTypeNumSal='NUM') AND (Length(LeSalarie) < 10 ) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;    // pt119
  if LeSalarie = '' then
  begin
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', Lib);
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    Q := OpenSql('SELECT * FROM DEPORTSAL WHERE PSE_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
    begin
      SalAdm := Q.FindField('PSE_ADMINISTABS').AsString; // Administrateur eAgl des absences et des primes
      SalVal := Q.FindField('PSE_OKVALIDVAR').AsString; // Autorisation de valider les primes = saisie des primes
      ConsultP := FALSE;
      // On va regarder si le salarie valide les absences pour autoriser les accès en consultation uniquement
      if SalVal <> 'X' then
      begin
        SalVal := Q.FindField('PSE_OKVALIDABS').AsString;
        ConsultP := TRUE;
      end;
      result := true;
    end
    else
    begin
      PgiBox('Vous êtes inconnu comme salarié#13#10 Renseignez vous au service du personnel', Lib);
      Ferme(Q);
      result := false;
      exit;
    end;
    if (SalAdm <> 'X') and (SalVal <> 'X') then
    begin
      PgiBox('Vous n''êtes pas autorisé à vous connecter, !#13#10Renseignez vous au service du personnel', Lib);
      result := false;
      exit;
    end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    //PT115 - Début
    else
      begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         Begin
           LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
         End
    //PT115 - Fin
      else
      begin
        PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', Lib);
        Ferme(Q);
        result := false;
        exit;
      end;
      Ferme(Q);
    end;
  end; //PT115
  result := true;
  BaseEPaie := GetParamSoc('SO_PGBASEEPAIE');
  if (not BaseEPaie) then
    if (SalAdm <> 'X') then
    begin
      PgiBox('L''accès à la saisie ' + VH_Paie.PGLibSaisPrim + ' est interdit#13#10 Reconnectez vous ultérieurement SVP', Lib);
      Result := FALSE; // Base non accessible en tant que base de saisie déportée
    end;
end;
// FIN PT16

//DEBUT PT37

function PgAnalConnectFor: boolean;
var
  Q: TQuery;
  LaChaine, Lib: string;
  BaseEPaie: Boolean;
begin
  Lib := 'Saisie formations';
  LaChaine := VerifieCheminPG (string(GetParamSoc('SO_PGCHEMINEAGL')));
  if LaChaine = '' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_PGCHEMINEAGL"', True);
    if not Q.eof then
      VH_Paie.PGCheminEagl:= VerifieCheminPG (Q.FindField('SOC_DATA').AsString);
    Ferme(Q);
  end;
  result := FALSE;
  LaChaine := '0000000000'; // Masque du nunéro de salarié
  LeSalarie := V_PGI.UserSalarie;
 // pt119  if Length(LeSalarie) < 10 then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;
  if (VH_Paie.PgTypeNumSal='NUM') AND (Length(LeSalarie) < 10 ) then LeSalarie := Copy(LaChaine, 1, 10 - Length(LeSalarie)) + LeSalarie;    // pt119
  if LeSalarie = '' then
  begin
    PgiBox('Vous n''avez pas de numéro salarié en tant qu''utilisateur PGI#13#10 Renseignez vous auprès de l''informatique interne !', Lib);
    result := FALSE;
    exit;
  end;
  if LeSalarie <> '' then
  begin
    if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_RESPONSFOR="' + LeSalarie + '"') then Result := True
    else
    begin
      if ExisteSQL('SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_SECRETAIREFOR="' + LeSalarie + '"') then Result := True
      else
      begin
//        PgiBox('Vous n''êtes pas autorisé à vous connecter, !#13#10Renseignez vous au service du personnel', Lib);//PT84
        result := false;
//        exit; PT84
      end;
    end;
    Q := OpenSql('SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE = "' + LeSalarie + '"', True);
    if not Q.eof then
      LeNomSal := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
    //PT115 - Début
    else
      begin
      Ferme(Q);
      Q := OpenSql('SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE = "' + LeSalarie + '"', True);
      if not Q.eof then
         Begin
         LeNomSal := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
         End
    //PT115 - Fin
      else
      begin
        PgiBox('Vous êtes inconnu comme salarié de la paie#13#10 Renseignez vous au service du personnel', Lib);
        Ferme(Q);
        result := false;
        exit;
      end;
      Ferme(Q);
    end; //PT115
  end;
  result := true;
end;
//FIN PT37

{DEB PT19 : Traitement par lots}

procedure DispatchTraitementLot;
var
  i: Integer;
  St, Nom, Value, Fiche, Cle: string;
begin
  for i := 1 to ParamCount do
  begin
    St := ParamStr(i);
    Nom := UpperCase(Trim(ReadTokenPipe(St, '=')));
    Value := UpperCase(Trim(St));
    if Nom = '/PGETATSCHAINES' then
    begin
        // Lancement d' une fiche pour mettre en inside, sinon le premier état chainé de s' imprime pas
        Value := Value + ';PG';
{$IFNDEF EABSENCES}
        //AGLLanceFiche('CP', 'ETATSCHAINES', '', '', Value); PT54
        CPEtatsChainesAuto('PG');
{$ENDIF}
        // Fermeture de l'application
        Application.ProcessMessages;
{$IFNDEF EAGLCLIENT}
        FMenuG.ForceClose := True;
{$ENDIF}
        FMenuG.Close;
        Exit;
    end
    //PT98 - Début //PT100
    // Mode fiche : Ouverture d'une fiche quelconque en modal
    // Paramètres à passer (sans les espaces) : /FICHE = NomFiche & CléElement & ParamètresOptionnels
    // Exemple1 : /FICHE=SALARIE&0000000001&ACTION=CONSULTATION
    // Exemple2 : /FICHE=SALARIE_MUL&&
    Else If Nom = '/FICHE' Then
    Begin
        // Libération des traitements parallèles
        Application.ProcessMessages;

        // Récupération des paramètres de lancement
        Fiche := ReadTokenPipe(Value, '&');
        Cle   := ReadTokenPipe(Value, '&');

{$IFNDEF EAGLCLIENT}
        // On lance une fiche vide pour que l'AGL la mette en fond de la fenêtre principale
        // Sans ça, si on lance la fiche spécifiée directement, elle sera intégrée dans la MainForm
        // de l'application, invisible (car la mainform est cachée), et le reste du traitement (c'est-à-dire
        // l'arrêt de l'appli) sera exécuté (car une fiche en fond n'est pas modale).
        AglLanceFiche('PAY', 'VIERGE', '', '', '');
{$ENDIF}

        // Envoi d'un message interapp pour indiquer le chargement de la fiche désirée
        If Fiche = 'SALARIE' Then
            PostMessage(HWND_BROADCAST, MessageInterApp, 11, 0)
        Else If Fiche = 'VIREMENT_MUL' Then
            PostMessage(HWND_BROADCAST, MessageInterApp, 21, 0)
        Else
            PostMessage(HWND_BROADCAST, MessageInterApp, 10, 0);

        // Lancement de la fiche en mode modal
        // Verrue en cas de problème avec l'AGL qui ne lance pas forcément la fiche, on ne sait pourquoi :
        // On relance la fiche tant qu'on n'est pas entré dedans (dans le onArgument, FicheOuverte est passé à True
        VH_PAIE.FicheOuverte := False;
        While Not VH_PAIE.FicheOuverte Do
        Begin
            Debug('Lancement de la fiche '+Fiche+'...');
            AglLanceFiche('PAY', Fiche, '', Cle, Value);
            Debug('Fiche '+Fiche+' fermée.');
        End;

        // Fermeture de l'application
        FMenuG.ForceClose := True;
        FMenuG.Close;
        Exit;
    End;
    //PT98 - Fin
  end;
end;
{FIN}

{$IFDEF EAGLCLIENT}
// Nothing !
{$ELSE}
function MyAglMailForm (var Sujet : HString; Var A, CC: String; Body: HTStringList; var Files: String;
ZoneTechDataVisible: boolean=True) : TResultMailForm;
begin;
    Result := rmfBad;
    if  ShowNewMessage('', '', '', A,V_PGI.NoDossier,Files) then
    Result := rmfOk;
end;
{$ENDIF}

procedure InitApplication;
var i  : Integer;
    St : String;
    hdl : Cardinal;
    handles : TList;
    j : integer;
begin
{$IFDEF MEMCHECK}
    MemCheckLogFileName := ChangeFileExt(Application.exename, '.log');
    MemChk;
{$ENDIF}

    //PT98 - Début
    // Enregistrement d'un type de message interapp pour pouvoir
    // dialoguer avec une autre instance de l'application
    MessageInterApp := RegisterWindowMessage('MSG_INTERAPP_CEGID');

    // On vérifie les paramètres de lancement de l'Application.
    // Si on trouve un /FICHE=, c'est qu'on est en mode de lancement particulier
    VH_PAIE.ModeFiche := False;
    ChargementComplet := True; //PT100

    For i:=1 To ParamCount Do
    Begin
        St := ParamStr(i);
        If pos('/FICHE=', St) <> 0 Then
        Begin
            Debug ('Lancement de l''application en mode FICHE');
            VH_PAIE.ModeFiche := True;

            // On rend transparente la fenêtre principale
            Application.MainForm.AlphaBlend := True;
            Application.MainForm.AlphaBlendValue := 0;

            // Cas particulier :
            // On change le titre de l'application au besoin
            If Pos('/FICHE=SALARIE&', St) <> 0 Then
            Begin
                Application.Title := 'Salarié'; //!!!!
                ChargementComplet := False;
            End;

            // Envoi du message 1 : Initialisation
            PostMessage(HWND_BROADCAST, MessageInterApp, 1, 0);
            Break;
        End;
    End;
    //PT98 - Fin

    ProcCalcEdt := CalcOLEEtatPG;
    ProcZoomEdt := ZoomEdtEtat;
    ProcCalcMul := PGCalcMul;


    //ProcGetVH:=GetVS1 ;
    //ProcGetDate:=GetDate ;
    FMenuG.OnDispatch := Dispatch;
    //PT98
    //  FMenuG.OnChargeMag := ChargeMagHalley;
    FMenuG.OnChargeMag := ChargeMagHalleyPaie;

{$IFNDEF EAGLCLIENT}
    FMenuG.OnMajAvant := nil;
    FMenuG.OnMajApres := nil;
{$ENDIF}
    // Gestion de la seria Modif faite pour V1.1 du 12-01-2001 pour reconnaitre Seria PCL
    // Deplaceé dans apres connexion pour avoir les variables V_PGI... initialisées
    // Mais initialisation faite quand meme pour avoir la variable sAcces definie pour la fonction AfterProtect
    // PH suppression test seria

    FMenuG.OnAfterProtec := AfterProtec;
    FMenuG.SetModules([45], [45]);
{$IFDEF EPRIMES}
    FMenuG.SetModules([40], [40]);
{$ENDIF}

    FMenuG.OnChangeModule := AfterChangeModule;

    V_PGI.DispatchTT := DispatchTT;

    V_PGI.OnAppliquerRole := PGAffecteEtabByUser;
    //PT26  : 17/06/2003 V_421 PH initialisation Préférence et gestion des favoris
{$IFNDEF EAGLCLIENT}
    FMenuG.SetPreferences(['Divers'], False);
{$ENDIF}
    FMenuG.OnChargeFavoris := PgChargeFavoris;
    V_PGI.DispatchTT := DispatchTT;
{$IFNDEF CPS3}
{$IFNDEF RHSEUL}
    if (V_PGI.ModePcl = '1') then TitreHalley := 'PAIE-GRH'
    else TitreHalley := 'PAIE-GRH';
{$ELSE}
    if (V_PGI.ModePcl = '1') then TitreHalley := 'Ressources humaines PGI Expert'
    else TitreHalley := 'Ressources humaines';
{$ENDIF}
{$ELSE}
{$IFDEF CPS1}
    TitreHalley := 'Paie 50';
{$ELSE}
    TitreHalley := 'PAIE-GRH';
{$ENDIF}
{$ENDIF}
{$IFDEF TESTSIC}
{$IFDEF EAGLCLIENT}
    SaveSynRegKey('eAGLHost', 'SIC-DEV2', true);
    FCegidIE.HostN.Enabled := false;
    TitreHalley:='TEST Paie&RH S5' ;
    NomHalley:='eCPS5' ;
    //  SaveSynRegKey('eAGLHost', 'CWASDRH:80', true);
    //  SaveSynRegKey('eAGLHost', 'TSESIC:8096', true);
    //  SaveSynRegKey('eAGLHost', 'SRV-LYO-TECH9:8096', true);
    //  FCegidIE.HostN.Enabled := false;
{$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}
end;

{ TFMenuDisp }

procedure TFMenuDisp.FMenuDispCreate(Sender: TObject);
begin
  // PgiAppAlone:= TRUE;
  // CreatePGIApp;          supprimer à partir de la version 422 AGL
  SeriaMessOk := False;
end;

procedure InitLaVariablePGI;
begin
  Apalatys := 'CEGID';
  HalSocIni := 'cegidpgi.ini';
  NomHalley := 'CPS5';

{$IFDEF EABSENCES}
  HalSocIni := 'cegidpgi.ini';
  NomHalley := 'eCABS5';
{$ENDIF}
{$IFDEF ETEMPS}
  NomHalley := 'eCABS5';
  HalSocIni := 'IlEstTemps.ini';
{$ENDIF}
  Copyright := '© Copyright ' + Apalatys;
  // Serie S5
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S5;
  if (V_PGI.ModePcl = '1') then
     begin
     TitreHalley := 'PAIE-GRH PGI Expert';
     V_PGI.CompterLesimpressions:= true;
     V_PGI.PRocCompteImpression:= CompteImpression;
     end
  else
     TitreHalley := 'PAIE-GRH ';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS5.HLP';

{$IFDEF EPRIMES}
  TitreHalley := 'SAISIE DES PRIMES';
  NomHalley := 'eCPRIS5';
  HalSocIni := 'reganame.ini';
{$ENDIF}
{$IFDEF CPS3}
{$IFDEF CPS1}
  // Serie S3
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S1;
  TitreHalley := 'PAIE';
//  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS1.HLP';
  NomHalley := 'CPS1';
{$ELSE}
  // Serie S3
  V_PGI.OutLook := TRUE;
  V_PGI.LaSerie := S3;
  TitreHalley := 'PAIE';
//  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS3.HLP';
  NomHalley := 'CPS3';
{$ENDIF}
{$ELSE}
{$IFDEF SERIES7}
  // Serie S7
  V_PGI.OutLook := FALSE;
  V_PGI.LaSerie := S7;
  TitreHalley := 'PAIE-GRH S7';
  NomHalley := 'CPS7';
  Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CPS7.HLP';
{$ENDIF}
{$ENDIF}
  V_PGI.ToolsBarRight := TRUE;

{$IFDEF RHSEUL}
  TitreHalley := 'Ressources humaines';
  NomHalley := 'CPRHS5';
  HalSocIni := 'cegidpgi.ini';
{$ENDIF}

  if (V_PGI.ModePcl = '1') then ForceCascadeForms;
{$IFDEF EAGLCLIENT}
  V_PGI.AutoSearch := False;
{$ENDIF}
  V_PGI.MultiUserLogin := true;
  V_PGI.CegidUpdateServerParams := 'www.update.cegid.fr';
  ChargeXuelib;
  V_PGI.OfficeMsg := TRUE;
  V_PGI.NoModuleButtons := FALSE;
{$IFNDEF CPS3}
  V_PGI.NbColModuleButtons := 1;
  V_PGI.NbRowModuleButtons:=6;
{$ELSE}
  V_PGI.NbColModuleButtons := 1;
  V_PGI.NbRowModuleButtons:=7;
{$ENDIF}
{$IFDEF PGTOUTMODULE}
 V_PGI.NbColModuleButtons := 3;
  V_PGI.NbRowModuleButtons:=5;
{$ENDIF}
  V_PGI.VersionDemo := True;
  V_PGI.MenuCourant := 0;
  V_PGI.VersionReseau := True;
  V_PGI.NumVersion := '8.06';
  V_PGI.NumBuild := '000.036';
{$IFDEF eTemps}
  V_PGI.NumVersionBase := 620;
{$ELSE}
  V_PGI.NumVersionBase := 850;
{$ENDIF}

  V_PGI.DateVersion := EncodeDate(2008, 09, 29);
  //V_PGI.SAV:=True ;
  V_PGI.ImpMatrix := True;
  V_PGI.OKOuvert := FALSE;
  V_PGI.Halley := TRUE;
  //V_PGI.NiveauAcces:=1 ;
  V_PGI.MenuCourant := 0;
  //V_PGI.Decla100:=TRUE ;
  { Pour faire du DEBUG  }
{$IFDEF ENCOURS}
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.SAV := False;
{$ELSE}
  V_PGI.DebugSQL := False;
  V_PGI.Debug := False;
  V_PGI.SAV := False;
{$ENDIF}
  V_PGI.QRPdf := True;
  V_PGI.NumMenuPop := 27;
  //  version 422 AGL
  V_PGI.CegidBureau := True;
  V_PGI.CegidApalatys := FALSE;
  // pour autoriser l'acces au dp comme dossier de travail,au multi dossier
  V_PGI.StandardSurDP := true;
  // pour autoriser l'acces au dp comme base de reference pour tous les dossiers ==> Table DP+STD
  V_PGI.MajPredefini := True;
  // Confidentialité
  V_PGI.PGIContexte := [CtxPaie];
  V_PGI.RAZForme := TRUE;
  V_PGI.DispatchTT := DispatchTT;
  V_PGI.TabletteHierarchiques := TRUE;
  // Blocage des mises à jour de structure par l'appli. Elle doit etre faite par PgiMajVEr fourni dans le kit socref
  V_PGI.BlockMAJStruct := TRUE;
{$IFDEF PAIEOLE}
  RegisterPaiePGI;
{$ENDIF}
  if V_PGI.ModePcl = '1' then
  begin
{$IFNDEF EAGLCLIENT}
//PT42    V_PGI.ToolBarreUp := False;
    V_PGI.PortailWeb := 'http://clients.cegid.fr';   //PT104
{$ENDIF}
  end
  else V_PGI.PortailWeb := 'http://clients.cegid.fr';  //PT104
{$IFNDEF EAGLCLIENT}
  if V_PGI.ModePcl = '1' then
    V_PGI.SansMessagerie := true
  else V_PGI.SansMessagerie := FALSE;
{$ENDIF}
  V_PGI.CodeProduit := '013';
  V_PGI.ScriptApp:= 'SCRIPTAPP_PAIE';
  V_PGI.DrawAero:= False;
{$IFDEF TESTSIC}
  TitreHalley := 'TEST ' + TitreHalley;
{$IFNDEF EAGLCLIENT}
  NomHalley := 'CPS5SIC';
{$ELSE}
  NomHalley := 'CPS5SIC';
{$ENDIF}
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  SetDidacticielPath (TCbpPath.getCegidDistriDoc);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 04/12/2007 / PT98
Modifié le ... :   /  /
Description .. : Fonction permettant de fermer l'application si le chargement
Suite ........ : du magHalley ne s'est pas correctement effectué alors qu'on
Suite ........ : est en mode FICHE (donc avec fenêtre principale cachée)
Mots clefs ... :
*****************************************************************}
Function ChargeMagHalleyPaie : Boolean;
Begin
    Result :=  ChargeMagHalley;
    If Not Result Then
    Begin
        If VH_PAIE.ModeFiche Then
        Begin
            FMenuG.ForceClose := True;
            FMenuG.Close;
        End;
    End;
End;

initialization
  ProcChargeV_PGI := InitLaVariablePGI;
end.


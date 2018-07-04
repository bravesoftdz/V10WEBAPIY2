{********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 30/05/2001
Modifié le ... :   /  /
Description .. : Saisie des bulletins
Suite ........ : Forme DELPHI compatible eAGL
Suite ........ :
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
unit SaisBul;

{ Unité Saisie du bulletin

Gestion de la FORME DELPHI, des Grids ...
Appel aux fonctions communes du moteur de calcul de la paie de P5UTIL
La fonction recupInfoModules recuperer les informations provenant des modules annexes tq
les CP,les Absences, les Frais Professionnels, les Acomptes, la saisie par rubrique.

------------------------------------------------------------------------------
PT1 : 04/07/2001 V547 SB Visualisation des CP en mode consultation
    On ne permet pas la consultation des CP,Les fiches dindemnité et de
    valorisation passe en consultation aussi..
PT2 : 29/08/2001 V547 PH Affichage des cumuls en bas de bulletin quand la paie
    est cloturée
PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
PT4 : 03/09/2001 V547 PH Correction zones et boutons non accessibles en consultation
PT5 : 04/09/2001 V547 SB Gestion consultation ou saisie des CP sur dernier bulletin
PT6 : 07/09/2001 V547 PH Controle de la date de la paie en création de bulletin
PT7 : 10/09/2001 V547 SB Fiche de bug n°266
                         Gestion des msg d'erreur CP en prep auto
                         fn RecupInfoModules modifiée
PT8 : 26/09/2001 V562 SB Saisie des absences via la saisie du bulletin
PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
PT10 : 25/10/2001 V562 PH controle en modification des dates du bulletin qu'il n'existe pas de
                          bulletin sur les mêmes critères
PT11 : 07/11/2001 V562 PH En création bulletin, on se remet en debut et en fin de mois pour les
                          dates de paie
PT12 : 13/11/2001 V562 SB Modification générée par la mise en place des prises de CP sur acquis en cours
PT13 : 21/11/2001 V562 PH test de suppression d'une ligne existante
PT14 : 27/11/2001 V562 PH Propagation d'une ligne de commentaire
PT15 : 19/12/2001 V571 PH Mémorisation faction du trentième si forcé
PT16 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
PT17 : 26/12/2001 V571 PH Insertion des lignes de commentaire multiples
PT18 : 27/12/2001 V571 PH Insertion d'une cotisation plus de message rubrique déjà utilisée si rem avec le meme numéro
PT19 : 02/01/2002 V571 SB Recalcul cp non fait quand on modifie les trentièmes
PT20 : 02/01/2002 V571 PH Boutons non accessible ou visible dans la saisie des bases de cotisation
PT21 : 07/01/2002 V571 SB Recalcul des CP suite modification du trentième
PT22 : 08/01/2002 V571 SB Edition du bulletin complémentaire
PT23 : 21/01/2002 V571 PH Réaffectation des dates début et fin pour identifiant analytique = param à la fonction
PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
PT25 : 19/03/2002 V571 SB Fiche de bug n°452 Ajout variable de paie prédefini Cegid
                          Heures et jours absences pris
PT26 : 20/03/2002 V571 SB Rechargement des variables de paie CP
PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
PT28 : 26/03/2002 V571 PH Appel fonction raz structure entete ChpEntete
PT29 : 04/04/2002 V571 PH Acces saisie des colonnes des cotisations si element variable
PT30-1 23/04/2002 V571 SB Fiche de bug 10021 : Variable de paie 0039 non calculé en préparation auto
                          Calculé ds ChargeBulletin non appellé par préparation auto
                          Déplacement du code ds RecupInfoModules
PT30-2 24/04/2002 V571 SB Fiche de bug 10021 : Gestion des éléments onglet Absences (heures,jours ouvrés, ouvrables)
PT31 : 24/04/2002 V582 PH seule la zone libelle est accessible pour une rubrique de commentaire
PT32 : 29/04/2002 V582 PH Optimisation requete SQL sur Salarié
PT33 : 30/04/2002 V571 SB Test Date de paiement renseignée suite plantage..
PT34 : 02/05/2002 V571 SB Erreur : Pour un SLD, On consomme systématiquement l'acquis en cours dans sa totalité
                          même si on ne solde pas la totalité sur ce mvt (cas AJU Ou AJP) du fait de
                          l'utilisation de pls tob pour l'integration des acquis en cours
PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
PT36 : 14/05/2002 V582 PH FQ 10112 Controle si rubrique de base de cotisation
PT37 : 15/05/2002 V582 PH FQ 331 Controle si ventilation analytique existente en sauvegarde
PT38-1 : 22/05/2002 V582 SB Fiche de bug n°10129 point 1
                          Validation d'un bulletion pré-solde ou cloturé (ajout var SoldeOuClot)
PT38-2 25/05/2002 V582 SB Fiche de bug n°10129 point 2 Gestion de pls SLD dans le même mois,
                         affectation des cons. et codeEtape erronné
                         Lors de la création d'un second bulletin, on réinitialise les CP
PT39 : 03/06/2002 V582 PH Gestion historique des évènements
PT40 : 19/06/2002 V582 PH Integration des lignes de la saisie des primes dans le bulletin complémentaire
PT41 : 08/08/2002 V582 PH Rechargement des infos salariés après accès fiche
PT42 : 08/08/2002 V582 PH Pb suppression apres modif bulletin ==> provoquait une validation
PT43 : 08/08/2002 V582 PH Calcul du bulletin lancé en fonction du paramètre (automatique ou pas)
PT44 : 08/08/2002 V582 PH controle en modification des dates du bulletin
PT45 : 06/09/2002 V585 PH Historisation 5 eme elt de salaire + Date retroactivité par defaut = date debut paie
PT46   03/10/2002 V585 SB Optimisation requête saisie de la paie
PT47-1 03/10/2002 V585 SB Re-calcul absence suite zoom absence en saisie bulletin
PT47-2 16/10/2002 V585 SB Réintégration CP suite zoom Congés payés erroné si suppr CP Pri
PT47-3 16/10/2002 V585 SB On recalcul les CP lorsqu'on décoche le trentième forcé !!Prise de congé sur acquis en cours
PT47-4 17/10/2002 V585 SB FQ n°10235 Intégration du calcul des variables 28,29 heures ouvrées ouvrables
PT48   17/10/2002 V591 PH controle des dates du bulletin lors de la creation plusieurs paies sur la même période
PT49   05/11/2002 V591 PH Suppression majuscule du libelle saisie dans les cotisations
PT50   05/11/2002 V591 PH Affichage du libelle de la bse de cotisation en insertion de ligne
PT51   14/11/2002 V591 PH Calcul du bulletin en changeant d'onglet si une modif faite
PT52   14/11/2002 V591 PH Possibilité de changer (et conserver) le libellé de la ligne en creation de ligne
PT53   26/11/2002 V591 PH le champ non imprimable d'une rémunération est affiché en gras+italiq
PT54   27/11/2002 V591 PH Recupération date de sortie et recalcul du trentième si modif salarié dans le bulletin
PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
PT56   06/12/2002 V591 PH Accès aux grilles en consultation dans le cas où la paie est close
PT57   13/12/2002 V591 SB FQ 10332 Suite PT5 en calcul congés on vérifie que l'on est bien sur le dernier bulletin en cours
PT58   16/12/2002 V591 PH FQ 10317 gestion du navigateur
PT59   18/12/2002 V591 PH Suppression ligne Dossier du PopUpMenu - non utilisé
PT60   19/12/2002 V591 SB FQ 10391 Modification message d'erreur
PT61   19/12/2002 V591 PH Affectation de la date de paiement à la date de fin de paie si non renseignée
PT62   27/12/2002 V591 SB FQ 10422 Création de pls SLD dans session erronée
PT63   07/01/2003 V591 PH Ne calcule plus le bulletin systématiquement si on modifie les cotisations ie regarde paramsoc
PT64   15/01/2003 V591 PH FQ 10449 Correction duplication ventilations analytiques en changeant de date de fin de paie
PT65-1 16/01/2003 V591 PH FQ 10449 Désactivation insertion et suppression de ligne pdt la modif des dates de paie
PT65-2 20/01/2003 V591 SB Recalcul du bulletin généré que si modif dates de paie
PT66   30/01/2003 V591 SB Vidage Tob_Abs en sortie et ré entrée de bulletin
PT67   05/06/2003 V_421 PH FQ 10510 Création illimitée de bulletins complémentaires
PT68   05/06/2003 V_421 PH FQ 10700 Gestion origine de la rubrique dans les bases de cotisations
PT69-1 13/06/2003 V_42 SB FQ 10731 CP : non prise en compte des bulletins compl. postérieur session
PT69-2 18/06/2003 V_42 SB FQ 10640 & FQ 10642 Mise en place des raccourcis clavier
PT70   19/06/2003 V_421 PH FQ 10590 Suppression automatique ligne si rubrique inexistante
PT71-1 24/06/2003 V_42 SB FQ 10454 Calcul au maintien sur premier bulletin implique un recalcul des CP
PT71-2 25/06/2003 V_42 SB FQ 10556 Verrouillage du recalculCp en modification trentième car génère des anomalies
PT71-3 25/06/2003 V_42 SB FQ 10097 Verrouillage paie à cheval date de cloture CP
PT71-4 27/06/2003 V_42 SB FQ 10628 En chargement variable 36 non alimenté
                                   En sortie date paie variable 27 & 36 non recalculé
PT72   02/07/2003 V_421 PH FQ 10737 Message indiquant le rechargement du contexte pour inserer une rubrique
PT73   06/08/2003 V_421 PH FQ 10665 Activation ascenseur onglet base de cotisation en consultation de bulletin
PT74   08/08/2003 V_421 PH FQ 10740 Rajout Champ etablissement dans table HISTOSALARIE
PT75   12/08/2003 V_421 SB FQ 10755 Division par zéro impossible
PT76   26/08/2003 V_421 PH Mise en place de la saisie arret
PT77   02/09/2003 V_421 PH Message indiquant la provenance extérieur de la ligne du bulletin
PT78   15/09/2003 V_421 PH Simulation saisie arret dans la saisie du bulletin
PT79-1 19/09/2003 V_42  SB FQ 10422 Gestion multi solde dans session bulletin
PT79-2 19/09/2003 V_42  SB FQ 10780 Recalcul SLD si modification trentième en création
PT80   19/09/2003 V_421 PH Boucle analyse contenu de la TOB
PT81   23/09/2003 V_421 PH Suppression ou maj saisie arret en suppression ou modification de bulletin
PT82   06/10/2003 V_421 PH FQ 10875 Activation acces analytique en CWAS
PT83   21/10/2003 V_421 PH FQ 10642 Raccorci saisie arret
PT84   18/11/2003 V_50  SB FQ 10794 Controle profil Congés payés existant
                                    Affichage Zoom CP si gestion CP
PT85   15/01/2004 V_50 SB FQ 11004  Modification du lancement de la fiche individuelle
PT86   12/03/2004 V_50 SB FQ 11162 Encodage de la date de cloture erroné si fin fevrier
PT87   18/03/2004 V_50 SB FQ 11020 Modification utilisateur des acquis sur la saisie de la paie
PT88   22/03/2004 V_50 PH Intégration de histosaisprim dans un bulletin complémentaire
PT89   09/04/2004 V_50 SB FQ 11136 Ajout Gestion des congés payés niveau salarié
PT90   16/04/2004 V_50 SB FQ 11205 Suppression du contrôle de chargement des Cp si pas dernier bulletin
PT91   29/04/2004 V_50 SB FQ 10794 ZommCp : gestion message si pas de profil CP lors d'un SLD
PT92-1 07/05/2004 V_50 SB FQ 11199 Dysfonctionnement Calcul Indemnité CP sur 1er bulletin
PT93   12/05/2004 V_50 PH FQ 10973 Restriction droit accès au salarié
PT94   17/06/2004 V_50 MF Ajout traitement d'intégration des IJSS (brutes et nettes)
PT95   18/06/2004 V_50 PH FQ 11164 Alimentation de la Base SS pratiquée dans la table PAIIENCOURS
PT96   21/06/2004 V_50 PH FQ 10793 Indicateur de contrôle de modification des salariés pdt la saisie du bulletin
PT97   23/06/2004 PH V_50 FQ 11380 Principe en modification de bulletin pas d'alignement des rubriques par rapport aux profils
                                     en création alignement, bouton Défaire ou revenir à l'état initial provoque l'alignement
PT98   28/06/2004 V_50 PH Modification d'une ligne d'import avec controle du paramsoc PGModifLigneIMP
PT99   08/07/2004 V_50 PH Accès à la consultation des cumuls en consultation de bulletin
PT100  12/07/2004 V_50 PH FQ 11049 Présentation détaillée du bulletin affectation des rubriques dans les
                          bonnes grilles lors du 1er lancement
PT101  12/08/2004 VG V_50 Suppression "Voir Simulation saisie arrêt" en S3
                           (bouton [Zoom]) - FQ N°11128
PT102  13/08/2004 V_50 PH FQ 11497 Ergonomie
PT103  13/08/2004 V_50 PH FQ 11491 En bulletin complémentaire pas d'accès CP modifiés
PT104  13/08/2004 V_50 PH FQ 11500 Affichage CP modifiés sur bulletin clos et saisie bloquée
PT105  16/08/2004 V_50 PH Suppression de ligne
PT106  18/08/2004 V_50 PH FQ 11453 Message de contrôle date de sortie du salarié après zoom salarié en creation bulletin
PT107  23/08/2004 PH V_50 FQ 11459 Limitation des trentièmes si 2 bulletins dans le même mois
PT108  23/08/2004 PH V_50 FQ 11524 Message accès CP en création de bulletin
PT109  23/08/2004 PH V_50 FQ 11034 modif date de sortie du salarié, et annulation de la modif
PT110  25/08/2004 MF V_50 Ajout traitement d'intégration des lignes de maintien
PT111  10/09/2004 MF V_50 complément traitement intégration IJSS et maintien
PT112  13/09/2004 PH V_50 FQ 11581 Controle origine de la rubrique en cas de commentaire
PT113  14/09/2004 PH V_50 FQ 10949 Message d'alerte de suppression de bulletin comptabilisé
PT114  16/09/2004 PH V_50 FQ 11607 Trentième affecté si modifié sur bulletin complémentaire
PT115  23/09/2004 MF V_50 Intégration IJSS & maintien
PT116  04/10/2004 MF V_50 Intégration IJSS & maintien ( la Tob_Abs n'était pas libérée en suppression de bulletin)
PT117  06/10/2004 JL V_50 Modifs LookUpList avec StSQL <> ''
PT118  08/10/2004 MF V_50 Intégration IJSS & maintien : Suppression des lignes
                          de maintien générées en cas de non validation du bulletin
PT119  12/10/2004 MF V_50 Intégration IJSS & maintien : FQ 11673 traitement des IJ uniquement
PT120  13/10/2004 PH V_50 Nouveau thème des rémunérations Net à Payer idem Salaires
PT121  26/10/2004 PH V_60 RAZ Tob des absences du salarie
PT122  05/11/2004 PH V_60 FQ 11766 Bulletin complémentaire sur une date de début <> debut du mois
PT123  30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT124  30/11/2004 MF V_60 Correction Après Voir absence la tob doit être vidée
PT125  06/12/2004 PH V_60 FQ 11831 Erreur SQL spécif DB2
PT126  10/01/2005 PH V_60 Touche Return dans la saisie de la grille provoque TABULATION
PT127-1 21/01/2005 SB V_60 Ajout raccourci pour recalcul des congés payés
PT127-2 21/01/2005 SB V_60 FQ 11876 Correction anomalies Bulletin SLD
PT127-3 21/01/2005 SB V_60 Ajout traitement Ergonomie sur zoom CP
PT127-4 21/01/2005 SB V_60 Modif. var globale pour non confusion avec var locale
PT128   28/01/2005 MF V_60 IJSS & maintien: Ajout de 2 paramètres à l'appel de RecupereRegltIJSS
PT129   02/02/2005 PH V_602 Positionnement du focus sur la grille de saisie au lieu des dates
PT130   04/02/2005 PH V_602 Prise en compte cumul alpha dans zoom voir cumuls
PT131   18/03/2005 SB V_60 FQ 11406 Chargement de la tob des calendriers, et des jours feriès
PT132   29/03/2005 SB V_60 FQ 11990 Optimisation traitement Calendrier
PT133   04/01/2005 SB V_60 FQ 11806 Voir Absence : Refonte traitement pour recalcul bulletin si modif absence
PT134   12/04/2005 PH V603 FQ 11840 Accès aux elements nationaux dans le menu click droit FQ 11840
PT135   13/04/2005 PH V603 FQ 11975 Sortie de la forme par la croix rouge
PT136   02/03/2005 Mf V_60 IJSS & maintien : correction en prépa auto il faut intégrer les IJ
                           si gestion des IJSS et coche maintien décochée  (PT132 V600 correction)
PT137   31/05/2005 PH V603 FQ 12337 Ne pas bloquer la préparation automatique si erreur de validation
PT138   02/06/2005 MF V603 Qd "NON" en validation on ne supprime les lignes de maintien
                           générées qu'en création de bulletin
PT139   11/07/2005 PH V603 FQ 11337 Gestion des cumuls en alphanumérique
PT140   11/07/2005 PH V603 FQ 11380 Bouton retour Etat initial on se positionne sur saisie rémunérations comme
                      si on rentre en saisie du bulletin pour la premiere fois
PT141   25/07/2005 PH V603 FQ 12457 Rechargement des infos salariés suite mode création à partir du bulletin existant
PT142   02/08/2005 PH V603 FQ 12474 Ergonomie message
PT142   08/08/2005 PH V603 FQ 11746 Rechargement des taux AT en cas de modif etab
PT143   08/08/2005 PH V603 FQ 10929 Controle de la date de fin de bulletin en fonction des contrats
PT144   31/08/2005 PH V603 FQ 12524 Conservation contenue de la zone si click droit rubrique/profil
PT145   31/08/2005 PH V603 FQ 12523 Liste des profils/rubrique non OK CWAS-->Bug AGL OK en remettant la zone saisie à blanc
PT146   12/09/2005 PH V603 Rechargement des caches rem,cot et personnalisations dans les focntions voir rem et cot
PT147   15/09/2005 PH V603 FQ 12494 Bulletin complémentaire pour un salarié entré en cours de mois
PT148-1 27/09/2005 SB V65  Ajout sous procédures CP dans recupinfomodules
PT148-2 27/09/2005 SB V65  FQ 12602 Ajout appel vérification CP en création SLDC
PT149   27/09/2005 SB V65  FQ 12507 enabled raccourci
PT150   29/09/2005 SB V65  FQ 12287 Gestion coche param soc. gestion des acomptes
PT151   08/12/2005 SB V_65 FQ 12700 Congé : Codetape = "S" pour ajustement négatif
PT152   18/01/2006 PH V_65 FQ 12826 Impossible de focaliser la fenetre avec le navigateur
                           si l'onglet remunérations n'est pas activé
PT153   18/01/2006 PH V_65 FQ 12672 navigateur et trentième forcé
PT154   20/02/2006 MF V_65 correction memcheck
PT155   28/02/2006 MF V_65 Traitement maintien qd IJSS : ajout rubrique de garantie
PT156   02/01/2006 PH V_65 FQ 12875 Répartition analytique ==> PT152 V6.50
PT157   14/04/2006 PH V_65 FQ 11914 Ajout paramètre qui indique si bulletin de solde pour forcer la récup de la
                           saisie par rubrique par rapport au debut et fin de mois
PT158   21/04/2006 MF V_65 FQ 13077 : modification du message d'anomalie qd rub. de garantie absente
PT159   21/04/2006 SB V_65 FQ 12556 : Violation d'acces en suppr. absence
PT160	  24/04/2006 SB V_65 FQ 12786 Anomalie calcul variable 0027 en prep. auto
PT161	  05/05/2006 PH V_65 FQ 13136 Cas spécifique sortie sans saisie avec focus sur ligne vide
PT162   09/05/2006 SB V_65 FQ 12995 Insertion des absences en fonction des dates de paie
PT163	  15/05/2006 PH V_65 Recup dans TOB_RUB PPU_DATEDEBUT et PPU_DATEFIN des valeurs modifiées
PT164   16/05/2006 MF V_65 FQ 12995 (pls bull/mois) maintien et ijss intégrés en fonction des dates du bulletin
PT165   29/05/2006 PH V_65 FQ 13156 Séria Diode
PT166   29/05/2006 PH V_65 FQ 13210 Ergonomie
PT167	  02/06/2006 SB V_65 Optimisation mémoire
PT168	  19/06/2006 PH V_65 FQ 12289 Tjrs le calcul dans le cas d'une régul de cotisation
PT169   04/08/2006 PH Ecriture de la fonctionnalité Outil de diagnostic
PT170   05/10/2006 PH Chargement des bases de cotisations en fonction des dates du bulletins modifiées
                      cas du bulletin complémentaire de régul
PT171   02/11/2006 SB V_70 FQ 13399 Création variable pour recup jour & montant CP N-1
PT172   24/11/2006 PH V_70 FQ 13419 salarié rentre en cours de mois la saisie par rubrique ESt reprise
PT173   30/11/2006 SB V_70 FQ 13576 Anomalie Cp sur navigation
PT174   01/12/2006 MF  V_70 FQ 13718 Correction absence doublée qd maintien
PT175   06/12/2006 SB V_70 FQ 13734 refonte gestion origine rubrique
PT176   11/12/2006 PH V_70 Régularisation des rémunérations
PT177   18/12/2006 PH V_70 Insertion automatiques des rubriques manquantes
PT178   11/01/2007 PH V_70 FQ 13808 Orgine de la rubrique de commentaire
PT179   22/01/2007 GGU V_70 Lignes de commentaire multiple limité à 9 lignes par rubrique
PT180   22/01/2007 PH V_70 Gestion de la rétroactivité
PT181   02/03/2007 FC V_70 Rajout de la gestion des éléments dynamiques dossier
PT182   07/03/2007 FC V_70 Edition du bulletin, choisir le mode d'édition duplicata ou specimen
PT183   23/03/2007 FC V_70 FQ 13350 Rajout du menu Voir historique des cumuls
PT184   27/03/2007 PH Rajout paramètre fonction RecupTobSalarie pour prise en compte Historique date
PT186   02/04/2007 MF V_702 ajout condition de compil IFNDEF EAGLSERVER  à voir qd mise au point Process Server
PT187   16/04/2007 FC V_72 FQ 14089 Rajout d'un menu pour visualiser tous les bulletins d'un salarié
PT188   25/04/2007 PH V_72 FQ 13414 Navigateur et trentième forcé
PT189   03/05/2007 MF V_72 FQ 14088 Appel du calcul des IJSS sur clique droit
PT190   15/05/2007 FC V_72 Rajout du menu contextuel Eléments dossier
PT191   15/06/2007 FC V_72 FQ 14308
PT192   19/06/2007 FC V_72 FQ 14385 Insertion d'un ligne en saisie de bulletin : La liste doit contenir
                                    les lignes associées à l'activité + les lignes "toutes activités".
PT193   02/07/2007 MF V_72 FQ 14468 Affichage de l'historique des cumuls dans ordre chronologique
PT194   02/07/2007 PH V_72 FQ 13361 Analytique salarié dans la saisie du bulletin
PT195   09/07/2007 FC V_72 FQ 14535 Historisation population
PT196   10/07/2007 MF V_72 FQ 14338 calcul IJSS sur clique droit : ajout contrôle
                           avec paramètre société  "Gestion des IJSS"
PT197   11/07/2007 FC V_72 FQ 14550 concepts Ajout / Suppression de lignes dans le bulletin
PT198   24/07/2007 FC V_72 FQ 14610 Rechargement du contexte pour les éléments dossier
PT199   07/08/2007 PH V_72 FQ 13723 Saisie de z provoquait l'appel au click droit = raccourci
PT200   07/08/2007 PH V_72 FQ 14642 Erreur gestion des formats de colonnes des grilles
PT201   07/08/2007 PH V_80 FQ 14086 Plus d'erreur si absndon
PT202   08/08/2007 PH V_80 FQ 14087 Nouvel onglet des rémunérations saisissables tous thèmes confondus
PT203   10/08/2007 FC V_72 FQ 14604 Edition diagnostic
PT204   20/08/2007 PH V_72 FQ 14623 Paie envers sur une rubrique et non sur un cumul = paiement en net de prime
PT205   24/08/2007 PH V_72 FQ 14687 Date intégration rétroactivité=date de fin de paie et non date du jour
PT206   24/08/2007 PH V_72 FQ 14669 Analytique salarié en consultation si bulletin cloturé
PT207   28/08/2007 PH V_72 FQ 14691 Ergonomie nouvel onglet élément à saisir
PT208   17/09/2007 PH V_72 FQ 14773 Pb Raz compteur absence V0054 à tord dans le cas CP forcés ds bulletin
PT209   01/10/2007 GGU V_80 FQ 14492 Intégration automatique des saisies arrêt
PT210   05/10/2007 MF V_80 Traitement des jours de fractionnement
PT211   08/10/2007 GGU V_80 Fenêtre de diagnostique en TreeView au lieu de la ListeBox
PT212   16/10/2007 PH  V_80 FQ 14863 Saisie des reguls de rem sauvegardée sans avoir à calculer le bulletin auparavant
PT213   26/10/2007 NA V-80 Si Gestion de la Presence : mise à jour table PRESENCESALARIE lors de la validation
PT214   15/11/2007 FC V_80 Pas d'édition bulletin si paramsoc état bul par défaut non renseigné et dans salarie idem
PT215   15/11/2007 PH V_80 FQ 14933 Contrôle des droits dans le menupop de la paie
PT216   20/11/2007 PH V_80 FQ 13414 Raz du trentième uniquement en phase de creation
PT220   05/06/2008 MF V_82  FQ 459 et FQ 15466
PT223   13/08/2008 FC Gestion des confidentiels pour l'affichage de la fiche salarié
PT224   08/09/2008 PH Optimisation gestion du fractionnement
PT225   02/10/2008 MF FQ 15781 : pas de calcul du maintien qd ancienneté <= 01/01/1900
}

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Menus,
  ImgList,
  hmsgbox,
  HSysMenu,
  Grids,
  Hctrls,
  ComCtrls,
  HTB97,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  DBCtrls,
{$IFNDEF EAGLSERVER}
  FE_Main,
  EdtEtat,
  EdtRetat,
  QRE,
{$ENDIF}
  Db,
{$ELSE}
  MaineAGL,
  UtileAgl,
{$ENDIF}
{$IFNDEF EAGLSERVER}
  LookUp,
  AGLInit,
  Ventil,
{$ENDIF}
  uPaieRemunerations,
  uPaieCotisations,
  uPaieBases,
  uPaieEtabCompl,
  uPaieCumuls,
  uPaieVentiCot,
  uPaieVentiRem,
  uPaieEltDynSal, //PT181
  uPaieEltDynEtab, //PT181
  uPaieEltDynPop, //PT181
  uPaieEltNiveauRequis, //PT181
{$IFNDEF CPS1}
  PGPOPULOUTILS, //PT181
{$ENDIF}
  StdCtrls,
  Mask,
  ExtCtrls,
  P5Util,
{$IFNDEF CCS3}
  ED_Tools, // PT115
  PGIJSSMaintien, // PT115
{$ENDIF}
  UTOB,
  HPanel,
  UIUtil,
  HEnt1,
  HDimension,
  P5Def,
  HStatus,
//  HCalc, // PGVisuObjet,
  Hqry,
  PGoutils,
  PgCongesPayes,
  EntPaie,
  P5RecupInfos,
  ParamSoc,
  uPaieEltNatDOS,
  uPaieEltNatSTD,
  uPaieEltNatCEG,
  uPaieExecpt,
  PGCommun, HImgList, ULibEditionPaie;
// PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire
procedure SaisieBulletin(CodeSal, Etabl, BulCompl, ProfilPart: string; DateDebut, DateFin: TDateTime; Action: TActionBulletin; Q_Mul: THquery; TypeListe, CPG: Boolean);

function SauvegardeBul(CodeSalarie, Etab: string; DateD, DateF, DateS: TDateTime; Tob_Rub: TOB; NTrent, DTrent: Integer; ActionFiche: Boolean; var TOB_IJSS, TOB_Maintien: TOB): Boolean; //PT115

procedure RecupInfoModules(TOB_Rub, Tob_Sal: TOB; ActionB: TActionBulletin; Auto, GCPAuto: boolean; var StMsgErr: string; GenereMaintien, IJSS: boolean; var TOB_IJ: TOB); { PT111 PT115 PT119 }
{ DEB PT148-1 }
procedure PrefaceExecuteCP(TETABSAL: Tob; Auto: Boolean; var StMsgErr: string; TheSal: string);
procedure IntegreCPCreate(TETABSAL, TOB_RUB, TOB_Sal: Tob; Auto: Boolean; var StMsgErr: string);
procedure IntegreCPModif(TETABSAL, TOB_RUB, TOB_Sal: Tob; Auto: Boolean; var StMsgErr: string);
procedure IntegreCPConsult(TOB_RUB: Tob);
{ FIN PT148-1 }
procedure RechargePrisEtSoldePayes(Tob_rub, Tob_sal: tob);
type
  TFSaisBul = class(TForm)
    PCrit: THPanel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    Binsert: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    bDefaire: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    BFerme: TToolbarButton97;
    BDelete: TToolbarButton97;
    PTotaux: THPanel;
    Pages: TPageControl;
    PRemSal: TTabSheet;
    PBas: TTabSheet;
    PCot: TTabSheet;
    GRemSal: THGrid;
    GBas: THGrid;
    GCot: THGrid;
    HMTrad: THSystemMenu;
    Msg: THMsgBox;
    FSortieDef: TCheckBox;
    LblCumBrut: THLabel;
    LblCumBrutFiscal: THLabel;
    LblCumNetApayer: THLabel;
    LblCumNetImp: THLabel;
    LblCumHeures: THLabel;
    LblBrut: THLabel;
    LblBrutFiscal: THLabel;
    LblNetImp: THLabel;
    LblNetAPayer: THLabel;
    LblHeures: THLabel;
    bNewligne: TToolbarButton97;
    bDelLigne: TToolbarButton97;
    BChercher: TToolbarButton97;
    PRemHeures: TTabSheet;
    PRemPrimes: TTabSheet;
    PRemAbs: TTabSheet;
    PRemCplt: TTabSheet;
    PRemAvt: TTabSheet;
    PRemAbt: TTabSheet;
    PRemRet: TTabSheet;
    PRemNonImp: TTabSheet;
    PRemPrimeNonImpSoumis: TTabSheet;
    GRemHeures: THGrid;
    GRemPrimes: THGrid;
    GRemAbs: THGrid;
    GRemCplt: THGrid;
    GRemAvt: THGrid;
    GRemAbt: THGrid;
    GRemRet: THGrid;
    GRemNonImpSoumis: THGrid;
    GRemNonImp: THGrid;
    LblAbsHeures: TLabel;
    LblOuvres: TLabel;
    ChbxControlSmic: TCheckBox;
    ZChbxBaseForcee: TCheckBox;
    BCalculBull: TToolbarButton97;
    ImageList1: THImageList;
    BtnDetail: TToolbarButton97;
    BtnSimple: TToolbarButton97;
    MulProfilTempo: THMultiValComboBox;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    PopupMenuGlobal: TPopupMenu;
    Sal: TMenuItem;
    Rem: TMenuItem;
    EtablissPaie: TMenuItem;
    Variab: TMenuItem;
    CumSess: TMenuItem;
    BCommentaire: TToolbarButton97;
    ProfilPgi: TMenuItem;
    ProfilRub: TMenuItem;
    ChbxHoraireSalarie: TCheckBox;
    EdtHorReel: TEdit;
    EdtJoursOuvres: TEdit;
    LblOuvrable: TLabel;
    EdtJoursOuvrables: TEdit;
    LblProfilMulti: THLabel;
    BtnAligProfPer: TToolbarButton97;
    ZChbxTranchesForcees: TCheckBox;
    BlvlCumul: TBevel;
    Bevel2: TBevel;
    LblPaye: THLabel;
    LblPar: THLabel;
    Bevel1: TBevel;
    Bevel3: TBevel;
    Zoom: TMenuItem;
    PaieEnvers: TMenuItem;
    PopupMenuZoom: TPopupMenu;
    Salarie1: TMenuItem;
    Etablissement1: TMenuItem;
    CumSessSal1: TMenuItem;
    FicheInd1: TMenuItem;
    BasesCotis1: TMenuItem;
    Calendrier: TMenuItem;
    VoirCP: TMenuItem;
    BZoom: TToolbarButton97;
    Tablesdossier: TMenuItem;
    VoirhistoriquecoTisations1: TMenuItem;
    VoirhistoriqueRemunrations1: TMenuItem;
    Voir1: TMenuItem;
    VoirvalorisationAbsence1: TMenuItem;
    BRechargContext: TToolbarButton97;
    LblEtab: THLabel;
    TPSA_ETABLISSEMENT: THLabel;
    TSalarie: THLabel;
    FSalarie: THLabel;
    TPSA_LIBELLE: THLabel;
    TDateEntree: THLabel;
    LblDateSortie: THLabel;
    TPSA_DATEENTREE: THLabel;
    LDateNaissance: THLabel;
    TPSA_DATENAISS: THLabel;
    BtnOrigine: TToolbarButton97;
    Voirindemnitcompensatricedecongs1: TMenuItem;
    TPSA_DATESORTIE: THLabel;
    BVentil: TToolbarButton97;
    LTitre: TLabel;
    FTypeActionBul: TLabel;
    LblPeriodePaie: THLabel;
    FDate1: THCritMaskEdit;
    FDate2: THCritMaskEdit;
    HLabel3: THLabel;
    FTrentieme: THCritMaskEdit;
    LblSlash: TLabel;
    FTrentDenominateur: THCritMaskEdit;
    ChbxTrentiem: TCheckBox;
    CPMOIS: THCritMaskEdit;
    CPSUPP: THCritMaskEdit;
    CPANC: THCritMaskEdit;
    HLabel1: THLabel;
    TCPSUPP: THLabel;
    TCPANC: THLabel;
    FrancImage: TImage;
    EuroImage: TImage;
    Absence: TMenuItem;
    LBLEDTDU: THLabel;
    FEDTDU: THCritMaskEdit;
    LBLEDTAU: THLabel;
    FEDTAU: THCritMaskEdit;
    Image1: TImage;
    ImageBulCompl: TImage;
    TimerBul: TTimer;
    ModeRegle: THValComboBox;
    FDatePaie: THCritMaskEdit;
    ChbxRegltMod: TCheckBox;
    HLabel6: THLabel;
    // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
    ImageCalcul: TImage;
    VSaisieArret: TMenuItem;
    Cotisations1: TMenuItem;
    Cotzz: TMenuItem;
    Bevel4: TBevel;
    CBModifAcquis: TCheckBox;
    LbllTrentieme: THLabel;
    HFSortieDef: THLabel;
    HCBModifAcquis: THLabel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    CongesPayes1: TMenuItem;
    CalculCpPaie: TMenuItem;
    EltNationaux: TMenuItem;
    BRegul: TToolbarButton97;
    ChbxBaseForcee: TCheckBox;
    ChbxTranchesForcees: TCheckBox;
    TBSHTDIAG: TTabSheet;
    BTNDIAGNOSTIC: TToolbarButton97; //PT211 Ancienne fenêtre de diagnostique renommée
    PopupMenuDiag: TPopupMenu;
    DiagRub: TMenuItem;
    DiagBull: TMenuItem;
    Voirhistoriquecumuls1: TMenuItem;
    OpenDialog1: TOpenDialog;
    SauvDiag: TMenuItem;
    EnvMailDiag: TMenuItem;
    VisualiserBulletins: TMenuItem;
    CalculdesIJSS1: TMenuItem;
    Tablesdynamiques1: TMenuItem;
    PopupMenuAnal: TPopupMenu;
    AnalSal: TMenuItem;
    Analrubbul: TMenuItem;
    PRemSais: TTabSheet; // PT202
    GRemSais: THGrid; // PT202
    BtnSaisie: TToolbarButton97; // PT202
    PaieNetRub: TMenuItem;
    Editiondudiagnostic1: TMenuItem;
    Diag: TTreeView; //PT211
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BChercheClick(Sender: TObject);
    procedure bNewligneClick(Sender: TObject);
    procedure PagesChange(Sender: TObject);
    procedure bDelLigneClick(Sender: TObject);
    procedure bDetruitLMaLigne(ARow: Integer);
    procedure GRemSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GRemSalDblClick(Sender: TObject);
    procedure FDate2Exit(Sender: TObject);
    procedure FDateAuExit(Sender: TObject);
    procedure BSalarieClick(Sender: TObject);
    procedure GBasRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure LigneCommentaire(Sender: TObject);
    procedure GRemSalRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure BtnDetailClick(Sender: TObject);
    procedure BtnSimpleClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
    procedure FDate1Exit(Sender: TObject);
    procedure FDateDuExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure ChargeSalarie(Force: Boolean = FALSE);
    procedure EtatInitial;
    procedure NewEnreg; virtual;
    function OnSauve: boolean; virtual;
    function OnDelete: boolean; virtual;
    function EnregOK: boolean;
    procedure OnGrilleExit(Sender: TObject);
    procedure AfficheActionB;
    procedure GCotRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GRemSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCotSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCotSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GCotDblClick(Sender: TObject);
    procedure ZChbxBaseForceeClick(Sender: TObject);
    procedure GBaseSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBaseSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GBaseDblClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EtablissPaieClick(Sender: TObject);
    procedure RemClick(Sender: TObject);
    procedure CotzzClick(Sender: TObject);
    procedure SalClick(Sender: TObject);
    procedure VariabClick(Sender: TObject);
    procedure GRemSalElipsisClick(Sender: TObject);
    procedure GRemSalRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure OnGrilleEnter(Sender: TObject);
    procedure TraiteCotCellTxSal(Rub: string; var Cancel: Boolean; ACol, ARow: Integer; RegulCot: Boolean = FALSE); // PT168
    procedure TraiteCotCellMtSal(Rub: string; var Cancel: Boolean; ACol, ARow: Integer);
    procedure TraiteCotCellTxPat(Rub: string; var Cancel: Boolean; ACol, ARow: Integer; RegulCot: Boolean = FALSE); // PT168
    procedure TraiteCotCellMtPat(Rub: string; var Cancel: Boolean; ACol, ARow: Integer);
    procedure RemplissageLigneCot(Grille: THGrid; Lib: string; Base, TxSal, TxPat, MtSal, MtPat: Double; ARow: Integer; NbB, NbTxSal, NbTxPat, NbMtSal, NbMtPat: Integer);
    procedure GCotElipsisClick(Sender: TObject);
    procedure ProfilPgiClick(Sender: TObject);
    procedure ProfilRubClick(Sender: TObject);
    //    procedure BtnCalendrierClick(Sender: TObject);
    procedure ChbxHoraireSalarieClick(Sender: TObject);
    procedure EdtHorReelExit(Sender: TObject);
    procedure EdtJoursOuvresExit(Sender: TObject);
    procedure EdtJoursOuvrablesExit(Sender: TObject);
    procedure BtnAligProfPerClick(Sender: TObject);
    procedure FTrentiemeExit(Sender: TObject);
    procedure FTrentDenominateurExit(Sender: TObject);
    procedure ChbxTrentiemeClick(Sender: TObject);
    procedure GRemSalSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure BImprimerClick(Sender: TObject);
    procedure CumSessSalClick(Sender: TObject);
    procedure BCalculBullClick(Sender: TObject);
    procedure CumSessClick(Sender: TObject);
    procedure SalarieClick(Sender: TObject);
    procedure EtablissementClick(Sender: TObject);
    procedure FicheIndClick(Sender: TObject);
    procedure BZoomClick(Sender: TObject);
    procedure BasesCotis1Click(Sender: TObject);
    procedure TablesdossierClick(Sender: TObject);
    procedure VoirhistoriquecoTisations1Click(Sender: TObject);
    procedure VoirhistoriqueRemunrations1Click(Sender: TObject);
    procedure GBasRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure IndemcpClick(Sender: TObject);
    procedure ValoAbsenceClick(Sender: TObject);
    procedure CongesPayesClick(Sender: TObject);
    procedure BRechargContextClick(Sender: TObject);
    procedure PaieEnversClick(Sender: TObject);
    procedure PaieEnversRubClick(Sender: TObject);
    procedure BtnOrigineClick(Sender: TObject);
    procedure ZChbxTranchesForceesClick(Sender: TObject);
    procedure IndemSoldeClick(Sender: TObject);
    function AccesLigneBulletin(AccessSaisie, OrigineInfo: string): Boolean;
    function AccesSuppLigneBulletin(AccessSaisie: string): Integer;
    procedure FTrentiemeEnter(Sender: TObject);
    procedure BVentilClick(Sender: TObject);
    procedure GCotRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure HelpBtnClick(Sender: TObject);
    procedure AbsenceClick(Sender: TObject); //PT8
    procedure TimerBulTimer(Sender: TObject);
    procedure ChbxRegltModClick(Sender: TObject);
    procedure FDatePaieOnExit(Sender: TObject);
    procedure ModeRegleOnExit(Sender: TObject);
    procedure CalendrierClick(Sender: TObject);
    procedure LbllTrentiemeClick(Sender: TObject);
    procedure GRemKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FDatePaieEnter(Sender: TObject);
    procedure FTrentDenominateurEnter(Sender: TObject);
    procedure VSaisieArretClick(Sender: TObject);
    procedure CPAcquisExit(Sender: TObject); { PT87 }
    procedure CBModifAcquisClick(Sender: TObject);
    procedure CalculCpPaieClick(Sender: TObject);
    procedure EltNationauxClick(Sender: TObject);
    procedure LigneRegul(Sender: TObject);
    procedure AccderDiodefr1Click(Sender: TObject);
    procedure OnDiagClick(Sender: TObject);
    procedure DiagRubClick(Sender: TObject);
    procedure DiagBullClick(Sender: TObject);
    procedure Voirhistoriquecumuls1Click(Sender: TObject);
    procedure SauvDiagClick(Sender: TObject);
    procedure EnvMailDiagClick(Sender: TObject);
    procedure VisualiserBulletinsClick(Sender: TObject);
    procedure CalculIJSSClick(Sender: TObject);
    procedure TablesdynamiquesClick(Sender: TObject);
    procedure AnalSalClick(Sender: TObject); // PT194
    procedure AnalrubbulClick(Sender: TObject);
    procedure Editiondudiagnostic1Click(Sender: TObject); // PT189
  private { Déclarations privées }
    TOB_Sal, Tob_Rub: TOB;
    CurG: THGrid;
    CalculAFaire: Boolean; // Indique si Calcul du bulletin doit être fait
    StCellCur: string; // string contenant le contenu de la cellule en cours de traitement
    Modifier: Boolean;
    PageChang: Boolean;
    LaDateFin, LaDateDeb, DDOnEnter, DFOnEnter: TDateTime; // Stockage date de fin en creation de bulletin
    // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
    ActionSal: Boolean;
    MyTob: TOB;
    ActiveCalc: Boolean;
    ActivClick: Boolean;
    ETB_Activite: string;
    ForceAna: Boolean; // PT194
    StDiagRub, StDiagBul: string; //PT203
    procedure AccesBasesClick;
    procedure AccesCotClick;
    procedure AccesRemClick;
    procedure AccesCumulsClick; //PT183
    procedure AlimChampFiche;
    procedure LanceCalculTrentieme;
    procedure GrilleDansRub;
    procedure ChargeBulletin;
//    procedure bInsererligneClick(Sender: TObject);
    function Bouge(Button: TNavigateBtn): boolean;
    function RemFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
    procedure RemZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    procedure TraiteRemCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin);
    procedure TraiteRemCellBase(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean; TypT: string; TauxR: Double; NbT: Integer);
    procedure TraiteRemCellTaux(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean);
    procedure TraiteRemCellCoeff(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean);
    function TypeFonctRem(T1: TOB): WORD;
    function RemZoneAccessible(var ACol, ARow: Longint): Boolean;
    function CotZoneAccessible(var ACol, ARow: Longint): Boolean;
    procedure CotZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function CotFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
    procedure TraiteCotCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin);
    function BaseZoneAccessible(var ACol, ARow: Longint): Boolean;
    procedure BaseZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function BaseFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
    procedure TraiteBaseCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin; Arow: Integer);
    procedure ActiveElipsis; // Activation du bouton Elipsis
    function RendThemOnglet: string;
    procedure RemplissageLigneRem(Grille: THGrid; Lib: string; Base, Taux, Coeff, Montant: Double; ARow: Integer; NbB, NbT, NbC, NbM: Integer);
    procedure ReajustGrille; // Reajustement du nombre de lignes des grilles en focntion de leur contenu
//    function RemExist(Rub: string): Boolean; // Est que la rubrique saisie existe
    procedure CalculBullSaisie; // fonction de calcul de bulletin pour la saisie du bulletin
    procedure RazGrilles; // fonction de RAZ des grilles
    procedure CalAffectTrent; //Fonction qui calcule et affecte le trentieme en fonction de la saisie
    procedure AlimChampEntete; // Fonction qui passe en parametre au moteur de calcul la liste des éléments saisis et nécessaires aux calculs
    procedure GrilleEnable; // Fonction qui rend inactive les actions sur les grilles
//    procedure GrilleDisable; // Fonction qui active les actions sur les grilles
    procedure PresentationBul; // Fonction de presentation , Mise en page du bulletin et Type de Calcul (Automatique ou sur Click) en fonction des paramsoc
    procedure AffecteDetail; // Presentation detaillee du bulletin
    procedure AffecteSimple; // Presentation simplifiee du bulletin
    procedure GereQuery; // gestion de la qery
    procedure RempliTobHistoRem(Grille: THGrid; T_RemMere, T_RemFille: TOB);
    procedure LongueurChamp; // definition des longueurs des champs à saisir dans les différentes grilles
    procedure RemPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure BasPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure CotPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    function RemZoneGrise(ACol, ARow: Integer): Boolean;
    function BaseZoneGrisee(ACol, ARow: Longint): Boolean;
    function CotZoneGrisee(ACol, ARow: Longint): Boolean;
    procedure GridColorCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure AffectePreVentilAna;
    procedure AffecteEtatFiche(Comment: Boolean);
    procedure ReactiveBtn;
    function RemSaisissable(Trech: TOB): Boolean; // PT202
  public { Déclarations publiques }
    procedure RubDansGrille;
    procedure AfficheBasBulletin;
    procedure RecalculCp(ZoomCp: Boolean = False; BCalcul: Boolean = False); { PT92-1 }
    procedure PGEnabledRaccourciCp; //PT84
    procedure LanceCalculPaieCp; { PT127-1 }
  end;

var
  CodeSalarie, ClePGSynEltNAt, Etab, Trentieme_entree, Denominateur_Entree: string;
  ValidationOK: Boolean;
  DateD, DateF: Tdatetime; // Dates Passées en paramètre provenant du multi critères salariés à la paie
  ActionBul, ActionLigne, ActionCellule: TActionBulletin; // Type de sasie Creation, Modif,consult
  Date1, Date2: TDateTime; // Dates saisies en cas de creation de bulletin
  RubrComment: string; // Memorisation de la rubrique pour laquelle on va associer des commentaires
  ACal, TypeList, GblCP, FicheLoad, PGTopRecalculCp: Boolean; { PT92-1 17/05/2004 PT127-4 }
  QMul: THQuery;
  TobSolde: Tob;
  X: TFSaisBul;
  SoldeOuClot: boolean; //PT38-1
  TOB_IJSS, TOB_Maintien: TOB; // PT115
implementation
{$IFNDEF EAGLSERVER}
{$IFDEF EAGLCLIENT}
uses
  UTofPG_GeneCompta,
  Web, MailOL,
  Ventana, PGCalendrier, UTOFPGUtilitaireCP;
{$ELSE}
uses MenuDispPS5,
  UTofPG_GeneCompta,
  Web, MailOL,
  UtilTom,
  ventana, utobdebug,
  PGCalendrier, UTOFPGUtilitaireCP;
{$ENDIF}
{$ELSE}
uses UTOFPGUtilitaireCP, PGCalendrier;
{$ENDIF}


{$R *.DFM}

//PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire

procedure SaisieBulletin(CodeSal, Etabl, BulCompl, ProfilPart: string; DateDebut, DateFin: TDateTime; Action: TActionBulletin; Q_Mul: THquery; TypeListe, CPG: Boolean);
{$IFNDEF EAGLSERVER}
var
  PP: THPanel;
{$ENDIF}
begin
  X := TFSaisBul.Create(Application);
  GblCP := CPG; { PT127-4 }
{$IFNDEF EAGLSERVER}
  PP := FindInsidePanel;
{$ENDIF}
  // Recupération des paramètres
  CodeSalarie := CodeSal;
  ClePGSynEltNAt := CodeSalarie;
  DateD := DateDebut;
  DateF := DateFin;
  ActionBul := Action;
  ActionBulCP := Action;
  ActionSld := Action;
  Etab := Etabl;
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire
  BullCompl := BulCompl;
  if BullCompl = '' then Bullcompl := '-'; // au cas où initialisation forcée
  ProfilSpec := ProfilPart;
  TypeList := TypeListe;
  TypeTraitement := 'SAISIE';
  QMul := Q_Mul;

  ActionCellule := taModification;
  FicheLoad := false;
{$IFNDEF EAGLSERVER}
  if PP = nil then
  begin
    try
      X.ShowModal;
    finally
      X.Free;
      X := nil;
    end;
  end else
  begin
    InitInside(X, PP);
    X.Show;
  end;
{$ENDIF}
end;
// ajout mv 20/03 pour compiler

procedure TFSaisBul.FormCreate(Sender: TObject);
begin
  Tob_sal := nil;
  Tob_Rub := nil;
  ValoXP0N := 0;
  ValoXP0D := 0;
end;

procedure TFSaisBul.FormDestroy(Sender: TObject);
begin
  // Rien A VOIR
end;


procedure TFSaisBul.FormShow(Sender: TObject);
begin
  PopUpSalarie := TRUE; // PT96 Restriction en consultation accès salariés
  StCellCur := '';
  ChargeMenuPop(-1, nil);
  ACal := FALSE;
  if VH_paie.PGTenueEuro then EuroImage.Visible := TRUE
  else FrancImage.Visible := TRUE;
  HMTrad.ResizeGridColumns(GRemSal);
  HMTrad.ResizeGridColumns(GRemSais); // PT202
  HMTrad.ResizeGridColumns(GBas);
  HMTrad.ResizeGridColumns(GCot);
  HMTrad.ResizeGridColumns(GRemHeures);
  HMTrad.ResizeGridColumns(GRemPrimes);
  HMTrad.ResizeGridColumns(GRemAbs);
  HMTrad.ResizeGridColumns(GRemCplt);
  HMTrad.ResizeGridColumns(GRemAvt);
  HMTrad.ResizeGridColumns(GRemAbt);
  HMTrad.ResizeGridColumns(GRemRet);
  HMTrad.ResizeGridColumns(GRemNonImp);
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  HMTrad.ResizeGridColumns(GRemNonImpSoumis);
  GRemNonImpSoumis.GetCellCanvas := GridColorCell;
  GRemSal.GetCellCanvas := GridColorCell;
  GRemSais.GetCellCanvas := GridColorCell; // PT202
  GRemHeures.GetCellCanvas := GridColorCell;
  GRemPrimes.GetCellCanvas := GridColorCell;
  GRemAbs.GetCellCanvas := GridColorCell;
  GRemCplt.GetCellCanvas := GridColorCell;
  GRemAvt.GetCellCanvas := GridColorCell;
  GRemAbt.GetCellCanvas := GridColorCell;
  GRemRet.GetCellCanvas := GridColorCell;
  GRemNonImp.GetCellCanvas := GridColorCell;
  GBas.GetCellCanvas := nil;
  GCot.GetCellCanvas := nil;

  if Qmul <> nil then
  begin
    BPrev.Visible := TRUE;
    BNext.Visible := TRUE;
    BFirst.Visible := TRUE;
    BLast.Visible := TRUE;
    BPrev.Enabled := TRUE;
    BNext.Enabled := TRUE;
    BFirst.Enabled := TRUE;
    BLast.Enabled := TRUE;
  end
  else
  begin
    BPrev.Visible := FALSE;
    BNext.Visible := FALSE;
    BFirst.Visible := FALSE;
    BLast.Visible := FALSE;
    BPrev.Enabled := FALSE;
    BNext.Enabled := FALSE;
    BFirst.Enabled := FALSE;
    BLast.Enabled := FALSE;
  end;
{$IFDEF EAGLCLIENT}
  BPrev.Visible := FALSE;
  BNext.Visible := FALSE;
  BFirst.Visible := FALSE;
  BLast.Visible := FALSE;
  BPrev.Enabled := FALSE;
  BNext.Enabled := FALSE;
  BFirst.Enabled := FALSE;
  BLast.Enabled := FALSE;
{$ENDIF}
  // @@@@@
  if not TimerBul.Enabled then TimerBul.Enabled := True;
  if VH_Paie.PGSaisieBulletin = 'DET' then BtnDetail.Down := TRUE; // PT100 FQ11049
  if VH_Paie.PGSaisieBulletin = 'SSS' then BtnSaisie.Down := TRUE; // PT202
  PresentationBul; // Mise en forme de la presentation du bulletin
  MemoriseTrentieme(-1); // RAZ du trentieme
  GBas.Enabled := TRUE;
  ActionLigne := taConsultation; // MOde saisie de ligne <> Creation
  ChargeLesExercPaie(DateD, DateF);
  EtatInitial;
  if VH_Paie.PGSaisieBulletin = 'SSS' then
  begin
    PRemSais.TabVisible := TRUE;
    Pages.ActivePage := PRemSais; // PT202
  end
  else Pages.ActivePage := PRemSal;

  AfficheActionB;
  BCalculBullClick(nil); // PT202 pour forcer le 1er Affichage correct
  //PT101
{$IFDEF CCS3}
  VSaisieArret.Enabled := False;
  VSaisieArret.Visible := False;
{$ENDIF}
  //FIN PT101

end;

procedure TFSaisBul.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//PT135  On valide la saisie avant de vider les TOB
  // PT42 : 08/08/2002 V582 PH Pb suppression apres modif bulletin ==> provoquait une validation
  if (Modifier = TRUE) and (ActionBul <> taSuppression) then BValiderClick(nil);
// FIN PT135
{$IFNDEF EAGLSERVER}
{$IFDEF EAGLCLIENT}
{$ELSE}
  RechargeMenuPopPaie;
{$ENDIF}
{$ENDIF}
  PopUpSalarie := FALSE; // PT96 Restriction en consultation accès salariés
  VideLaTobExer;
  if TOB_Sal <> nil then Tob_Sal.Free;
  TOB_Sal := nil;
  //if TOB_Salarie <>Nil then Tob_Salarie.Free ;
  TOB_Salarie := nil;
  if TOB_DUSALARIE <> nil then TOB_DUSALARIE.Free;
  TOB_DUSALARIE := nil;
  if TOB_HistoBasesCot <> nil then
  begin
    TOB_HistoBasesCot.Free;
    TOB_HistoBasesCot := nil;
  end;
  if Tob_Rub <> nil then
  begin
    Tob_Rub.Free;
    Tob_Rub := nil;
  end;
  if TOBAna <> nil then
  begin
    TOBAna.Free;
    TOBAna := nil;
  end;
  if TobAcSaisieArret <> NIL then FreeAndNil(TobAcSaisieArret);
  if IsInside(Self) then Action := caFree;
end;

procedure TFSaisBul.BChercheClick(Sender: TObject);
var
  rep: Integer;
begin
  Rep := MSG.execute(4, Caption, '');
  if rep = mrNo then exit;
  ChargeSalarie;
  RubDansGrille;
  ChargeBulletin;
end;

procedure TFSaisBul.PagesChange(Sender: TObject);
begin
  if VH_Paie.PGAnalytique = FALSE then BVentil.Visible := FALSE
  else BVentil.Visible := TRUE;
  BCommentaire.Enabled := TRUE;
  BtnOrigine.Visible := TRUE;
  if (Pages.ActivePage = PBas) then BCommentaire.Enabled := FALSE;
  if Pages.ActivePage = PRemSal then CurG := GRemSal;
  if Pages.ActivePage = PRemSais then CurG := GRemSais; // PT202
  if Pages.ActivePage = PRemHeures then CurG := GRemHeures;
  if Pages.ActivePage = PRemPrimes then CurG := GRemPrimes;
  if Pages.ActivePage = PRemAbs then CurG := GRemAbs;
  if Pages.ActivePage = PRemCplt then CurG := GRemCplt;
  if Pages.ActivePage = PRemAvt then CurG := GRemAvt;
  if Pages.ActivePage = PRemAbt then CurG := GRemAbt;
  if Pages.ActivePage = PRemRet then CurG := GRemRet;
  if Pages.ActivePage = PRemNonImp then CurG := GRemNonImp;
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  if Pages.ActivePage = PRemPrimeNonImpSoumis then CurG := GRemNonImpSoumis;

  // Gestion accessibilite aux saisies de champs ou de la grille
  if Pages.ActivePage = PBas then
  begin
    //PT68   05/06/2003 V_421 PH FQ 10700 Gestion origine de la rubrique dans les bases de cotisations BtnOrigine.Visible := FALSE;
    // PT20 : 02/01/2001 V571 PH Boutons non accessible ou visible dans la saisie des bases de cotisation
    BCommentaire.Enabled := FALSE;
    BVentil.Visible := FALSE;
    // PT73   06/08/2003 V_421 PH FQ 10665 Activation ascenseur onglet base de cotisation en consultation de bulletin
    // if ActionBul = taConsultation then GBas.Enabled := FALSE;
    // FIN PT73
    CurG := GBas;
  end;

  if Pages.ActivePage = PRemAbs then
  begin
    if ChbxHoraireSalarie.Checked = TRUE then
    begin
      if ActionBul = taConsultation then
      begin
        EdtHorReel.Enabled := FALSE;
        EdtJoursOuvres.Enabled := FALSE;
        EdtJoursOuvrables.Enabled := FALSE;
      end
      else
      begin
        EdtHorReel.Enabled := TRUE;
        EdtJoursOuvres.Enabled := TRUE;
        EdtJoursOuvrables.Enabled := TRUE;
      end;
    end
    else
    begin
      EdtHorReel.Enabled := FALSE;
      EdtJoursOuvres.Enabled := FALSE;
      EdtJoursOuvrables.Enabled := FALSE;
    end;
  end;
  PageChang := FALSE;
  if Pages.ActivePage <> PBas then // @@@@
  begin
    if Pages.ActivePage = PCot then CurG := GCot;
    BRegul.Visible := TRUE;
    PageChang := TRUE;
  end
  else BRegul.Visible := FALSE;
  // PT4 : 03/09/2001 V547 PH Correction zones et boutons non accessibles en consultation
  if ActionBul = taConsultation then
  begin
    // PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
    ChbxRegltMod.Enabled := FALSE;
    FDatePaie.Enabled := FALSE;
    ModeRegle.Enabled := FALSE;
    ChbxControlSmic.Enabled := FALSE;
    ChbxTranchesForcees.Enabled := FALSE;
    ChbxTrentiem.Enabled := FALSE;
    ChbxBaseForcee.Enabled := FALSE;
    ChbxHoraireSalarie.Enabled := FALSE;
    EdtHorReel.Enabled := FALSE;
    EdtJoursOuvres.Enabled := FALSE;
    EdtJoursOuvrables.Enabled := FALSE;
    BRechargContext.Enabled := FALSE;
    BCommentaire.Enabled := FALSE;
    BtnOrigine.Enabled := FALSE;
    FEDTDU.Enabled := FALSE;
    FEDTAU.Enabled := FALSE;
  end
  else
  begin
    ChbxRegltMod.Enabled := TRUE;
    //   FDatePaie.Enabled := TRUE;
    //   ModeRegle.Enabled := TRUE;
    ChbxControlSmic.Enabled := TRUE;
    ChbxTranchesForcees.Enabled := TRUE;
    ChbxTrentiem.Enabled := TRUE;
    ChbxBaseForcee.Enabled := TRUE;
    ChbxHoraireSalarie.Enabled := TRUE;
    {  EdtHorReel.Enabled := TRUE;
      EdtJoursOuvres.Enabled := TRUE;
      EdtJoursOuvrables.Enabled := TRUE;  }
    BRechargContext.Enabled := TRUE;
    BtnOrigine.Enabled := TRUE;
    FEDTDU.Enabled := TRUE;
    FEDTAU.Enabled := TRUE;
  end;
  if (not JaileDroitTag(200151)) then ChbxTrentiem.Enabled := FALSE;
  if (not JaileDroitTag(200152)) then
  begin
    ChbxBaseForcee.Enabled := FALSE;
    ChbxTranchesForcees.Enabled := FALSE;
  end;
  if (not JaileDroitTag(200153)) then
  begin
    ChbxRegltMod.Enabled := FALSE;
    FDatePaie.Enabled := FALSE;
    ModeRegle.Enabled := FALSE;
  end;
  if (not JaileDroitTag(200154)) then
  begin
    BNewLigne.Visible := FALSE;
    bDelLigne.Visible := FALSE;
  end;
  if (not JaileDroitTag(41220)) then PopUpMenuGlobal.Items[2].Visible := FALSE;
  if (not JaileDroitTag(41120)) then PopUpMenuGlobal.Items[4].Visible := FALSE;
  if (not JaileDroitTag(41150)) then PopUpMenuGlobal.Items[5].Visible := FALSE;
  if (not JaileDroitTag(41130)) then PopUpMenuGlobal.Items[7].Visible := FALSE;
  if (not JaileDroitTag(41140)) then PopUpMenuGlobal.Items[8].Visible := FALSE;
  if (not JaileDroitTag(41110)) then PopUpMenuGlobal.Items[9].Visible := FALSE;
  if (not JaileDroitTag(41411)) then PopUpMenuGlobal.Items[11].Visible := FALSE;
  if (not JaileDroitTag(41412)) then PopUpMenuGlobal.Items[12].Visible := FALSE;
  if (not JaileDroitTag(41105)) then PopUpMenuGlobal.Items[13].Visible := FALSE;

  CurG.Col := 1; // On se place tjs dans la 2eme Colonne de la grille
end;

procedure TFSaisBul.bNewligneClick(Sender: TObject);
var
  T: TOB;
  ARow, i: Integer;
  TTab: TTabSheet;
  Origin: string;
begin
  if ActionBul = taConsultation then exit;
  if (CurG.Row < 1) then exit;
// DEB PT178
  if TToolbarButton97(Sender).Name = 'BCommentaire' then
  begin
    if Curg = GCot then Origin := CurG.Cells[7, CurG.Row]
    else if Curg = GBAS then Origin := CurG.Cells[10, CurG.Row]
    else Origin := CurG.Cells[6, CurG.Row];
  end
  else Origin := RechDom('PGACTIONBULLETIN', 'BUL', FALSE);
// FIN PT178
  RubrComment := CurG.Cells[0, CurG.Row];
  if RubrComment = '' then RubrComment := CurG.Cells[0, 1];
  RubrComment := Copy(RubrComment, 1, 4); // pour extraire les 4 premiers de la rubrique de commentaire !
  ActiveElipsis;
  ARow := CurG.Row;
  CurG.CacheEdit;
  CurG.SynEnabled := False;
  CurG.InsertRow(CurG.Row);
  CurG.Row := ARow;
  CurG.MontreEdit;
  CurG.SynEnabled := True;
  CurG.Col := 0;
// DEB PT178
{  if CurG = GCot then CurG.Cells[7, CurG.Row] := RechDom('PGACTIONBULLETIN', Origin, FALSE)
    //PT68   05/06/2003 V_421 PH FQ 10700 Gestion origine de la rubrique dans les bases de cotisations
  else if CurG = GBas then CurG.Cells[10, CurG.Row] := RechDom('PGACTIONBULLETIN', Origin, FALSE)
  else CurG.Cells[6, CurG.Row] := RechDom('PGACTIONBULLETIN', Origin, FALSE);}
  if Curg = GCot then CurG.Cells[7, CurG.Row] := Origin
  else if Curg = GBAS then CurG.Cells[10, CurG.Row] := Origin
  else CurG.Cells[6, CurG.Row] := Origin;
// FIN PT178
  T := TOB.Create('HISTOBULLETIN', Tob_Rub, -1);
  CurG.Objects[0, CurG.Row] := T;
  ActionLigne := taCreation;
  ActionCellule := taCreation;
  PageChang := FALSE; // @@KK

  CurG.Invalidate;
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is TTabSheet then
    begin
      TTab := TTabSheet(Components[i]);
      if (TTab <> nil) and (Curg.Parent = TTab) then
      begin
        TTab.Invalidate;
        break;
      end;
    end;
  end;

  // BCalculBullClick (NIL);   // pour forcer le calcul de la rubrique afin d'avoir accès en saisie
end;
{
procedure TFSaisBul.bInsererligneClick(Sender: TObject);
var
  T: TOB;
  ARow: Integer;
begin
  if ActionBul = taConsultation then exit;
  if (CurG.Row < 1) then exit;
  RubrComment := CurG.Cells[0, CurG.Row];
  if RubrComment = '' then RubrComment := CurG.Cells[0, 1];
  RubrComment := Copy(RubrComment, 1, 4); // pour extraire les 4 premiers de la rubrique de commentaire !
  ActiveElipsis;
  ARow := CurG.Row + 1;

  CurG.CacheEdit;
  CurG.SynEnabled := False;
  CurG.InsertRow(CurG.Row + 1);
  CurG.Row := ARow;
  CurG.MontreEdit;
  CurG.SynEnabled := True;
  CurG.Col := 0;
  T := TOB.Create('HISTOBULLETIN', Tob_Rub, -1);
  CurG.Objects[0, CurG.Row + 1] := T;
  ActionLigne := taCreation;
  ActionCellule := taCreation;
end;
}

procedure TFSaisBul.bDelLigneClick(Sender: TObject);
var
  T, Trech: TOB;
  ARow, Nbre, i: Integer;
  St, RefA: string;
  ASupprimer: Integer;
begin
  Nbre := Curg.RowCount;
  St := CurG.Cells[0, 1];
  if ActionBul = taConsultation then exit;
  if (CurG.Row < 1) then exit;
  if (CurG.Row = 1) and (Nbre = 2) then exit;
  CurG.Col := 1;
  ActionLigne := taModification;
  ARow := CurG.Row;
  T := TOB(CurG.Objects[0, CurG.Row]);
  // PT13 : 21/11/2001 V562 PH test de suppression d'une ligne existante
  if T = nil then exit;
  // DEB PT97 Message alerte suppresion rubrique provenant d'un profil
  if Curg.Name = 'GBas' then i := 10
  else if Curg.Name = 'GCot' then i := 7
  else i := 6;
  if CurG.Cells[i, CurG.Row] = 'Profil' then
  begin
    ASupprimer := PgiAsk('Attention ! Cette rubrique fait partie du profil du salarié. Confirmez-vous sa suppression du bulletin en cours ?', 'Saisie des bulletins');
    if Asupprimer <> mrYes then exit;
  end;
  // FIN PT97
  RefA := EncodeRefPaie(CodeSalarie, DateD, DateF, T);
  ASupprimer := AccesSuppLigneBulletin(T.GetValue('PHB_ORIGINELIGNE'));
  if ASupprimer = mrNo then exit;
  T.Free;
  CurG.CacheEdit;
  CurG.SynEnabled := False;
  // CurG.RowCount:=CurG.RowCount-1;
  CurG.DeleteRow(CurG.Row);
  CurG.Row := ARow;
  CurG.MontreEdit;
  CurG.SynEnabled := True;
  ReajustGrille;
  AfficheBasBulletin;
  CurG.ElipsisButton := FALSE;
  if (CurG <> GBas) and (VH_Paie.PGAnalytique = TRUE) then
  begin // Suppression des preventilations par defaut concernant la ligne à supprimer
    TRech := TobAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
    while TRech <> nil do
    begin
      TRech.free; // suppression des TOB lignes preventilations analytiques concernant la ligne supprimee
      TRech := TobAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
    end;
  end;
end;

procedure TFSaisBul.bDetruitLMaLigne(ARow: Integer);
var
  T, Trech: TOB;
  ZRow, Nbre: Integer;
  St, RefA: string;
  ASupprimer: Integer;
begin
  Nbre := Curg.RowCount;
  St := CurG.Cells[0, 1];
  if ActionBul = taConsultation then exit;
  if (ARow < 1) then exit;
  if (ARow = 1) and (Nbre = 2) then exit;
  CurG.Col := 1;
  ActionLigne := taModification;
  ZRow := CurG.Row;

  T := TOB(CurG.Objects[0, ARow]);
  if T = nil then exit;
  RefA := EncodeRefPaie(CodeSalarie, DateD, DateF, T);
  ASupprimer := AccesSuppLigneBulletin(T.GetValue('PHB_ORIGINELIGNE'));
  if ASupprimer = mrNo then exit;
  T.Free;
  CurG.CacheEdit;
  CurG.SynEnabled := False;
  CurG.DeleteRow(ARow);
  if (CurG.Cells[0, ZRow] = '') then ZRow := ZRow - 1; // PT161
  CurG.Row := ZRow;
  CurG.MontreEdit;
  CurG.SynEnabled := True;
  ReajustGrille;
  AfficheBasBulletin;
  CurG.ElipsisButton := FALSE;
  if (CurG <> GBas) and (VH_Paie.PGAnalytique = TRUE) then
  begin // Suppression des preventilations par defaut concernant la ligne à supprimer
    TRech := TobAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
    while TRech <> nil do
    begin
      TRech.free; // suppression des TOB lignes preventilations analytiques concernant la ligne supprimee
      TRech := TobAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
    end;
  end;
end;

procedure TFSaisBul.GRemSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1: TOB;
begin
  CurG.ElipsisButton := FALSE; // pas de bouton ellipsis par defaut
  if ActionBul = taConsultation then exit;
  if (ACol = 0) and (ActionCellule = taCreation) then
  begin
    ActionCellule := taModification;
  end;
  if ActionLigne = taCreation then
  begin
    if ACol = 0 then CurG.ElipsisButton := TRUE
    else CurG.ElipsisButton := FALSE;
  end
  else
  begin
    T1 := TOB(CurG.Objects[0, CurG.Row]);
    if T1 = nil then
    begin
      cancel := TRUE;
      CurG.Col := ACol;
      CurG.Row := ARow;
      exit;
    end;
  end;
  RemZoneSuivanteouOk(ACol, ARow, Cancel);

  if not Cancel then StCellCur := CurG.Cells[CurG.Col, CurG.Row];
end;

procedure TFSaisBul.GRemSalDblClick(Sender: TObject);
var
  T, TS, TR: TOB;
  Rubriq, Ret: string;
  i: Integer;
  st: string;
  Q: TQuery;
begin
  Rubriq := Copy(CurG.Cells[0, CurG.Row], 1, 4);
  T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rubriq], TRUE);
  if T <> nil then
  begin
{$IFNDEF EAGLSERVER}
    ret := AglLanceFiche('PAY', 'REMUNERATION', '', T.GetValue('PRM_PREDEFINI') + ';' + T.GetValue('PRM_NODOSSIER') + ';' + T.GetValue('PRM_RUBRIQUE'), '');
{$ENDIF}
  end;
  if Ret <> '' then
  begin
    TS := TOB.create('La remuneration', nil, -1);
    St := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="' + Rubriq + '"';
    Q := OpenSql(st, TRUE);
    TS.LoadDetailDB('REMUNERATION', '', '', Q, FALSE, FALSE);
    TR := TS.detail[0];
    if TR <> nil then for i := 1 to T.NbChamps do T.Valeurs[i] := TR.Valeurs[i];
    TS.Free;
    ferme(Q);
  end;
  // PT146 rechargement des personnalisations
  initTOB_Execpt(); // Chargement de la tob des execptions
  Align_Execpt('', ''); // Récuperation des personnalisations
end;

procedure TFSaisBul.FDate2Exit(Sender: TObject);
var
  Q: TQuery;
  stq: string;
  T_Etab: Tob;
  DtCloture, Dfc, AncDateD, AncDateF: TDateTime;
  TDD, TDF: TDateTime;
  YY, MM, JJ, YYCl, MMCl: Word;
  NbreC: Integer;
// d PT164
  anomalie, reponse: integer;
  ChampCateg: string;
// f PT164
begin
  // PT65-1   16/01/2003 V591 PH FQ 10449 Désactivation insertion et suppression de ligne pdt la modif des dates de paie
  //        Fonction ReactiveBtn
  // En sortie du champ date de fin de paie, on active le calcul du trentieme
  AncDateD := DateD;
  AncDateF := DateF;
  if not IsValidDate(FDate2.Text) then
  begin
    Showmessage('La date de fin n''est pas valide');
    FDate2.SetFocus;
    Date2 := 0;
    ReactiveBtn;
    exit;
  end;
  Date2 := StrToDate(FDate2.Text);
  Date1 := StrToDate(FDate1.Text);
  //PT65-2 Test si date debut et fin bulletin non modifiées
  if (Date1 = DDOnEnter) and (Date2 = DFOnEnter) then
  begin
    ReactiveBtn;
    exit;
  end;

  if Date1 > Date2 then
  begin
    Showmessage('La date de début est supérieure à la date de fin');
    FDate1.SetFocus;
    FDate2.Text := '';
    Date2 := 0;
    Date1 := 0;
    ReactiveBtn;
    exit;
  end;
  if (Date2 > DateF) and (ActionBul <> taCreation) then
  begin
    Showmessage('La date de fin ne peut pas être ultérieure au mois sélectionné');
    FDate2.SetFocus;
    Date2 := 0;
    ReactiveBtn;
    exit;
  end;
  if BullCompl <> 'X' then // PT67 Test mis en place si bulletin complémentaire
  begin
    // PT10 : 25/10/2001 V562 PH controle en modification des dates du bulletin
    Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
      'PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT>="' + UsDateTime(Date1) +
      '" AND PPU_DATEFIN<="' + UsDateTime(Date2) + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      Stq := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(Date1) + ' au ' + DateToStr(Date2) +
        '#13#10 Vous devrez saisir des dates de période de bulletin qui ne correspondent ' +
        '#13#10 pas à une paie existante et indiquer les dates d''édition du bulletin';
      PGIBox(Stq, Caption);
      Ferme(Q);
      FDate2.SetFocus;
      Date2 := 0;
      ReactiveBtn;
      exit;
    end;
    Ferme(Q);

    // PT44 : 25/10/2001 V562 PH controle en modification de la date de début du bulletin
    Stq := 'SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
      'PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT>="' + UsDateTime(Date1) +
      '" AND PPU_DATEFIN<="' + UsDateTime(FINDEMOIS(Date2)) + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      TDD := Q.FindField('PPU_DATEDEBUT').AsDateTime;
      TDF := Q.FindField('PPU_DATEFIN').AsDateTime;
      if (Date2 >= TDD) then
      begin
        STQ := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(TDD) + ' au ' + DateToStr(TDF) +
          '#13#10 La date de fin de bulletin chevauche la date du début d''un autre bulletin.' +
          '#13#10 Vous devez changer votre date de fin de bulletin.';
        PGIBox(Stq, Caption);
        reponse := 7;
      end
      else
      begin
        Stq := 'Contrôle de la date de fin. #13#10Il existe déjà un bulletin sur la période du ' + DateToStr(TDD) + ' au ' + DateToStr(TDF) +
          '#13#10Voulez-vous continuer la saisie de ce bulletin ?';
        reponse := HShowMessage('5;;' + STQ + ';Q;YN;Y;N', '', '');
      end;
      if reponse <> 6 then
      begin
        Ferme(Q);
        FDate2.SetFocus;
        Date2 := 0;
        ReactiveBtn;
        exit;
      end;
    end;
    Ferme(Q);
    // PTXX : 17/10/2001 V591 PH controle en modification de la date de fin du bulletin
    Stq := 'SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
      'PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT >="' + UsDateTime(Date1) +
      '" AND PPU_DATEFIN >="' + UsDateTime(FINDEMOIS(Date2)) + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      TDD := Q.FindField('PPU_DATEDEBUT').AsDateTime;
      TDF := Q.FindField('PPU_DATEFIN').AsDateTime;
      if (Date1 <= TDD) then
      begin
        STQ := ' Attention, il existe déjà un bulletin sur la période du ' + DateToStr(TDD) + ' au ' + DateToStr(TDF) +
          '#13#10 La date de début de bulletin chevauche la date du fin d''un autre bulletin.' +
          '#13#10 Vous devez changer votre date de début du bulletin.';
        PGIBox(Stq, Caption);
        reponse := 7;
      end
      else
      begin
        Stq := 'Contrôle de la date de début. #13#10Il existe déjà un bulletin sur la période du ' + DateToStr(TDD) + ' au ' + DateToStr(TDF) +
          '#13#10 Voulez-vous continuer la saisie de ce bulletin ?';
        reponse := HShowMessage('5;;' + STQ + ';Q;YN;Y;N', '', '');
      end;
      if reponse <> 6 then
      begin
        Ferme(Q);
        FDate2.SetFocus;
        Date2 := 0;
        ReactiveBtn;
        exit;
      end;
    end;
    Ferme(Q);
  end
  else
    //PT67   05/06/2003 V_421 PH FQ 10510 Création illimitée de bulletins complémentaires
  begin // Bulletin complémentaire, on regarde si une paie identique existe
    Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
      'PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT="' + UsDateTime(Date1) +
      '" AND PPU_DATEFIN="' + UsDateTime(Date2) + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      Stq := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(Date1) + ' au ' + DateToStr(Date2) +
        '#13#10 Vous devrez saisir des dates de période de bulletin qui ne correspondent ' +
        '#13#10 pas à une paie existante et indiquer les dates d''édition du bulletin';
      PGIBox(Stq, Caption);
      Ferme(Q);
      FDate2.SetFocus;
      Date2 := 0;
      ReactiveBtn;
      exit;
    end;
    Ferme(Q);
    Stq := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND ' +
      'PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT>="' + UsDateTime(DEBUTDEMOIS(Date1)) +
      '" AND PPU_DATEFIN="' + UsDateTime(FINDEMOIS(Date2)) + '"';
    Q := Opensql(stq, true);
    if not Q.eof then
    begin
      Stq := 'Attention, il existe déjà un bulletin sur la période du ' + DateToStr(Date1) + ' au ' + DateToStr(Date2) +
        '#13#10 Vous devrez saisir des dates de période de bulletin qui ne correspondent ' +
        '#13#10 pas à une paie existante et indiquer les dates d''édition du bulletin';
      PGIBox(Stq, Caption);
    end;
    Ferme(Q);
  end;
  // FIN PT67
  // DEB PT143
  Stq := 'SELECT COUNT (*) NBREC FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + CodeSalarie + '"';
  Q := Opensql(stq, true);
  if not Q.eof then
  begin
    NbreC := Q.FindField('NBREC').AsInteger;
    if NbreC > 0 then
    begin
      Ferme(Q);
      Stq := 'SELECT PCI_FINCONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + CodeSalarie + '" ORDER BY PCI_FINCONTRAT DESC';
      Q := Opensql(stq, true);
      if not Q.eof then
      begin
        Dfc := Q.FindField('PCI_FINCONTRAT').AsDateTime;
        if Dfc < Date2 then
        begin
          Stq := 'Attention, le contrat de travail se termine le ' + DateToStr(Dfc) +
            '#13#10 alors que la date de fin de paie est au ' + DateToStr(Date2);
          PGIBox(Stq, Caption);
        end;
      end;
      Ferme(Q);
    end;
  end;
  Ferme(Q);
  // FIN PT143
  // PT6 : 07/09/2001 V547 PH Controle de la date de la paie en création de bulletin
  if not ControlPaieCloture(Date1, Date2) then
  begin
    Showmessage('Vous ne pouvez pas saisir de paie sur une période close');
    FDate1.SetFocus;
    ReactiveBtn;
    exit;
  end;
  // FIN PT6

  { DEB PT71-3}
  if (TOB_Etablissement <> nil) and (Etab <> '') and (VH_paie.PGCongesPayes = True) then
  begin
    T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
    if Assigned(T_Etab) then
      if T_Etab.GetValue('ETB_CONGESPAYES') = 'X' then
      begin
        DtCloture := T_Etab.GetValue('ETB_DATECLOTURECPN');
        DeCodeDate(Date2, YY, MM, JJ);
        DeCodeDate(DtCloture, YYCl, MMCl, JJ);
        DtCloture := PGEncodeDateBissextile(YY, MMCl, JJ); { PT86 }
        if (DtCloture > Date1) and (DtCloture < Date2) then
        begin
          PGIBox('Vous ne pouvez pas saisir de paie à cheval sur la date de clotûre des congés payés', Caption);
          FDate2.Text := DateToStr(DtCloture);
          ReactiveBtn;
          exit;
        end;
      end;
  end;
  { FIN PT71-3 }
  DateD := Date1;
  DateF := Date2;
  // PT64   15/01/2003 V591 PH FQ 10449 Correction duplication ventilations analytiques en changeant de date de fin de paie
  if (ActionBul = taCreation) and (VH_Paie.PGAnalytique) then
  begin // RAZ des ventilations analytiques car elles ne sont pas à la bonne date
    if TOBAna <> nil then TOBAna.Free;
    TOBAna := nil;
    TOBAna := TOB.Create('Analytique bulletin Salarié', nil, -1);
  end;
  // FIN PT64
  { DEB PT71-4 Recalcul des Jours et heures travaillées théoriques }
  InitVarTHTRAVAILLES;
  CalculJoursTheorique(Tob_sal, DateD, DateF, JTHTRAVAILLES, HTHTRAVAILLES);
  AlimChampFiche;
  AlimChampEntete;
  // DEB PT163
  if (DateD <> tob_rub.GetValue('PPU_DATEDEBUT')) then
    tob_rub.PutValue('PPU_DATEDEBUT', dateD);
  if (dateF <> tob_rub.GetValue('PPU_DATEFIN')) then
    tob_rub.PutValue('PPU_DATEFIN', dateF);
   // FIN PT163
  { FIN PT71-4 }
  if (AncDateD <> DateD) or (AncDateF <> DateF) then // PT170
    ChargeBasesSal(CodeSalarie, Etab, DateD, DateF);
  // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if BullCompl <> 'X' then
  begin
    { Gestion des congés payés }
    if not ChpEntete.CpAcquisMod then InitialiseVariableStatCP; { PT87 } //PT62 Réinitialisation des variables Cp pour création nouveau bulletin
    ReaffecteDateMvtcp(Datef);
    if not ChpEntete.CpAcquisMod then Recalculcp; { PT87 }
    { DEB PT162 Gestion des absences }
    if (VH_PAIE.PGAbsence) then
    begin
      EnleveRubAbsence(Tob_Sal, Tob_Rub, AncDateD, AncDateF);
      EnleveRubAbsence(Tob_Sal, Tob_Rub, DateD, DateF);
      if Assigned(Tob_Abs) then LibereTobAbs;
// PT164     if (RecupereAbsences(Tob_ABS, Datef, CodeSalarie, 'C', False, False)) then
{PT220
      if (RecupereAbsences(Tob_ABS, Datef, CodeSalarie, 'C', VH_Paie.PGMaintien, False)) then}
      if (RecupereAbsences(Tob_ABS, Datef, CodeSalarie, 'C', VH_Paie.PGMaintien, False,'')) then
        IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, 'C');
      RubDansGrille;
      CalculBullSaisie;
// d PT164
      if VH_Paie.PGMaintien then
      begin
        EnleveRubIJS(Tob_Sal, Tob_Rub, AncDateD, AncDateF);
        AnnuleIJSSBulletin(CodeSalarie, AncDateF);
        EnleveRubIJS(Tob_Sal, Tob_Rub, DateD, DateF);
        AnnuleIJSSBulletin(CodeSalarie, DateF);
        EnleveRubMAI(Tob_Sal, Tob_Rub, AncDateD, AncDateF);
        AnnuleMaintienBulletin(CodeSalarie, AncDateD, AncDateF);
        EnleveRubMAI(Tob_Sal, Tob_Rub, DateD, DateF);
        AnnuleMaintienBulletin(CodeSalarie, DateD, DateF);

        ChampCateg := '';
       // récupération du champ critère suplémentaire IJSS et Maintien
        if (VH_Paie.PGCritMaintien <> '') then
        begin
          Q := OpenSQL('SELECT PAI_PREFIX, PAI_SUFFIX FROM PAIEPARIM WHERE PAI_IDENT = ' +
            VH_Paie.PGCritMaintien, TRUE);
          if not Q.EOF then
          begin
            ChampCateg := Q.Fields[0].AsString + '_' + Q.Fields[1].AsString;
          end;
          ferme(Q);
        end;

       // calcul et maj Maintien   (paie envers...)
        MaintienDuSalarie(Tob_salarie, TOB_Rub, Tob_etablissement, Tob_Abs, DateToStr(Datef), DateToStr(dated), 'C', Anomalie, ChampCateg, TOB_IJSS, TOB_Maintien);

        if (anomalie <> 0) then
        begin
          if (Anomalie = 1) then
            PGIBox('Pas de règle de Maintien pour ce salarié', Caption)
          else
            if (Anomalie = 2) then
              PGIBox('Calcul du montant du maintien,lecalcul de la paie à l''envers a échoué', Caption)
            else
              if (Anomalie = 3) then
                PGIBox('Vérifiez le maintien de salaire et les ijss, la rubrique de garantie de salarie n''a pu être calculée ou est absente', Caption)
//d PT225
              else
                if (Anomalie = 4) then
                  PGIBox('La date d''ancienneté du salarié n''est pas renseignée : le maintien ne peut pas être calculé', Caption);
// f PT225
        end;

        RubDansGrille;

       // intégration abs IJSS, IJSS, Maintien
        CalculBullSaisie;
      end;
// f PT164

    end;
    { FIN PT162 }
  end;
  AfficheActionB;
  // ChargeBulletin;
  ReactiveBtn;
  // FIN PT65-1
end;

procedure TFSaisBul.FDateAuExit(Sender: TObject);
var
  Da1, Da2: TDateTime;
begin
  // En sortie du champ date de fin de paie, on active le calcul du trentieme
  if not IsValidDate(FEDTAU.Text) then
  begin
    Showmessage('La date de fin d''édition n''est pas valide');
    FEDTAU.SetFocus;
    exit;
  end;
  Da2 := StrToDate(FEDTAU.Text);
  Da1 := StrToDate(FEDTDU.Text);
  if Da1 > Da2 then
  begin
    Showmessage('La date de début d''édition est supérieure à la date de fin d''édition ');
    FEDTDU.SetFocus;
    FEDTAU.Text := '';
    exit;
  end;
  ChpEntete.Edtau := StrToDate(FEDTAU.Text);
  AlimChampFiche;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.BSalarieClick(Sender: TObject);
begin
end;

// Procedure d'alimentation des champs de la fiche

procedure TFSaisBul.AlimChampFiche;
var
  ZDate, LD, LF: TDateTime;
  T2, TPR, T_Etab: TOB;
  st: string;
  Q: TQuery;
  Val1, Val2: Double;
  DateE, DateS: TDateTime;
  rep: Integer;
begin
  FSalarie.Caption := CodeSalarie;
  if (TOB_Sal <> nil) and (CodeSalarie <> TOB_Sal.GetValeur(iPSA_SALARIE)) then
  begin
    if TOB_Sal <> nil then
    begin
      TOB_Sal.Free;
      TOB_Sal := nil;
    end;
  end;
  if TOB_Sal = nil then ChargeSalarie;
  if TOB_Sal <> nil then
  begin
    FSalarie.Caption := CodeSalarie;
    // PT54   27/11/2002 V591 PH Récupération date de sortie et recalcul du trentième si modif salarié dans le bulletin
    if (ActionBul = taCreation) and (ActionSal) then
    begin
      ActionSal := FALSE;
      DateE := TOB_Sal.GetValeur(iPSA_DATEENTREE);
      DateS := TOB_Sal.GetValeur(iPSA_DATESORTIE);
      if (DateE > DateD) and (DateE > LaDateDeb) then
      begin
        if (DateE <= LaDateFin) then DateD := DateE; // pour avoir les bonnes dates de la paie en fonction entree/sortie
      end;
      if DateS > 100 then
      begin
        if (DateS <> DateF) and (DateS <= LaDateFin) and (DateS >= LaDateDeb) and (BullCompl <> 'X') then DateF := Dates; // afin d'avoir un calcul de trentieme exact des le depart
        if (DateS < DateF) and (DateS < LaDateFin) and (DateS >= LaDateDeb) and (BullCompl <> 'X') then // PT106
        begin
          rep := PgiAsk('Attention, la date de sortie ' + DateToStr(DateS) + ' est supérieure à la date de fin de période ' + DateToStr(LaDateFin) +
            ',#13#10Pour reprendre la date de sortie et faire une paie sur une période différente#13#10de celle initialement prévue, répondez OUI à la question ci-dessous',
            'Période du bulletin');
          if rep = mrYes then DateF := DateS;
        end;
      end;
      if DateS <= IDate1900 then
      begin // PT109  // PT114 Fin de mois si non bulletin complémentaire
        if (not EstFinMois(Datef)) and (BullCompl <> 'X') then DateF := FINDEMOIS(DateF);
      end;
    end; // FIN PT54
    FDate1.Text := DateToStr(DateD);
    FDate2.Text := DateToStr(DateF);
    //    FDate1.Text := DateToStr(DateD);
    //    FDate2.Text := DateToStr(DateF);
    TPSA_LIBELLE.Caption := TOB_Sal.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Sal.GetValeur(iPSA_PRENOM); { PT87 }
    ZDate := TOB_Sal.GetValeur(iPSA_DATEENTREE);
    if ZDate <= 10 then TPSA_DATEENTREE.Caption := ' '
    else TPSA_DATEENTREE.Caption := DateToStr(Zdate);
    ZDate := TOB_Sal.GetValeur(iPSA_DATENAISSANCE);
    if ZDate <= 10 then TPSA_DATENAISS.Caption := ' '
    else TPSA_DATENAISS.Caption := DateToStr(Zdate);
    ZDate := TOB_Sal.GetValeur(iPSA_DATESORTIE);
    if ZDate <= 10 then TPSA_DATESORTIE.Caption := ' '
    else TPSA_DATESORTIE.Caption := DateToStr(Zdate);
    if TOB_Sal.GetValeur(iPSA_SORTIEDEFINIT) = 'X' then FSortieDef.Checked := TRUE
    else FSortieDef.Checked := FALSE;
    if ActionBul = taCreation then Etab := TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)
    else
    begin
      if QMul.FindField('PPU_ETABLISSEMENT') <> nil then Etab := QMul.FindField('PPU_ETABLISSEMENT').AsString
      else Etab := TOB_Sal.GetValeur(iPSA_ETABLISSEMENT);
    end;
    TPSA_ETABLISSEMENT.Caption := RechDom('TTETABLISSEMENT', Etab, FALSE);
    // PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
    // LblModeRegle.Caption :=RechDom ('PGMODEREGLE',RendModeRegle (TOB_Sal),FALSE) ;
    T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
    ZDate := CalculDatePaie(TOB_Sal, DateF);
    if ZDate > 10 then FDatePaie.Text := DateToStr(ZDate)
    else FDatePaie.Text := '';
    with ChpEntete do
    begin
      Reglt := RendModeRegle(TOB_Sal);
      if ActionBul = taCreation then ModeRegle.value := Reglt;
      Ouvres := 0;
      Ouvrables := 0;
      HOuvres := 0;
      HOuvrables := 0;
      if T_Etab <> nil then
      begin
        {DEB PT47-4 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables
        Ouvres:=NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,TRUE);
        Ouvrables:=NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,FALSE);}
        {PT132 Mise en commentaire , utilisation du calendier chargé en globale
          Tob_CalendrierSalarie := ChargeCalendrierSalarie(TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT),
          TOB_Salarie.GetValeur(iPSA_SALARIE), TOB_Salarie.GetValeur(iPSA_CALENDRIER),
          TOB_Salarie.GetValeur(iPSA_STANDCALEND), T_Etab.getvalue('ETB_STANDCALEND')); }
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, GblTob_Semaine, DateD, DateF, TRUE, Ouvres, HOuvres); { PT132 }
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, GblTob_Semaine, DateD, DateF, FALSE, Ouvrables, HOuvrables); { PT132 }
      end; //FIN PT47-4
      EdtJoursOuvres.Text := DoubleToCell(Ouvres, 2);
      EdtJoursOuvrables.Text := DoubleToCell(Ouvrables, 2);
    end;
    if TOB_Rub = nil then
    begin
      TPR := TOB.Create('Le Bulletin de la Paie En Cours', TOB_Paie, -1);
      // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
      st := 'SELECT * FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '"' + ' AND PPU_SALARIE="' + TOB_Sal.GetValeur(iPSA_SALARIE) + '" AND PPU_DATEDEBUT="' +
        USDateTime(DateD) + '" AND PPU_DATEFIN="' + USDateTime(DateF) + '" AND PPU_BULCOMPL = "' + BullCompl + '"';
{Flux optimisé
      Q := OpenSql(st, TRUE);
      TPR.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
}
      TPR.LoadDetailDBFROMSQL('PAIEENCOURS', st);
      T2 := TPR.FindFirst(['PPU_ETABLISSEMENT', 'PPU_SALARIE'], [Etab, TOB_Sal.GetValeur(iPSA_SALARIE)], TRUE);
    end
    else T2 := TOB_Rub;
    if (T2 <> nil) then
    begin
      ChpEntete.DatePai := T2.GetValue('PPU_PAYELE');
      ChpEntete.DateVal := T2.GetValue('PPU_VALIDELE');
      // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
      if ChpEntete.Edtdu < 10 then ChpEntete.Edtdu := T2.GetValue('PPU_EDTDEBUT');
      if ChpEntete.Edtau < 10 then ChpEntete.Edtau := T2.GetValue('PPU_EDTFIN');
      // initialisation avec des valeurs par defaut au cas où
      if ChpEntete.Edtau < 10 then ChpEntete.Edtau := DateF;
      if ChpEntete.Edtdu < 10 then ChpEntete.Edtdu := DateD;
      FEdtDu.Text := DateToStr(ChpEntete.Edtdu);
      FEdtAu.Text := DateToStr(ChpEntete.Edtau);
      //PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
      if T2.GetValue('PPU_REGLTMOD') = 'X' then
      begin
        ChpEntete.RegltMod := TRUE;
        ChpEntete.Reglt := T2.getvalue('PPU_PGMODEREGLE');
        ModeRegle.value := T2.getvalue('PPU_PGMODEREGLE');
        FDatePaie.Text := T2.getvalue('PPU_PAYELE');
        ChbxRegltMod.Checked := TRUE;
        FDatePaie.Enabled := TRUE;
        ModeRegle.Enabled := TRUE;
      end
      else
      begin
        ChbxRegltMod.Checked := FALSE;
        FDatePaie.Enabled := FALSE;
        ModeRegle.Enabled := FALSE;
      end;
      { DEB PT87 }
      ChpEntete.CpAcquisMod := (T2.GetValue('PPU_CPACQUISMOD') = 'X');
      CBModifAcquis.Checked := ChpEntete.CpAcquisMod;
      { FIN PT87 }
      if T2.GetValue('PPU_BASESMOD') = 'X' then
      begin
        ChbxBaseForcee.Checked := TRUE;
        ChpEntete.BasesMod := TRUE;
      end
      else
      begin
        ChbxBaseForcee.Checked := FALSE;
        ChpEntete.BasesMod := FALSE;
      end;
      if T2.GetValue('PPU_TRANCHESMOD') = 'X' then
      begin
        ChbxTranchesForcees.Checked := TRUE;
        ChpEntete.TranchesMod := TRUE;
      end
      else
      begin
        ChbxTranchesForcees.Checked := FALSE;
        ChpEntete.TranchesMod := FALSE;
      end;
      if T2.GetValue('PPU_HORAIREMOD') = 'X' then
      begin
        ChbxHoraireSalarie.Checked := TRUE;
        ChpEntete.HorMod := TRUE;
      end
      else
      begin
        ChbxHoraireSalarie.Checked := FALSE;
        ChpEntete.HorMod := FALSE;
      end;
      // DEB PT114  Prise en compte du bulletin complémentaire en mode création affectation cohérente des champs de la TOB avec la fiche
      if (T2.GetValue('PPU_TRENTIEMEMOD') = 'X') or (ChbxTrentiem.Checked and (ActionBul = taCreation)) then
      begin
        ChbxTrentiem.Checked := TRUE;
        FTrentieme.Enabled := TRUE;
        ChpEntete.TrentMod := TRUE;
        FTrentDenominateur.Enabled := TRUE;
        if (ChbxTrentiem.Checked and (ActionBul = taCreation)) then
        begin
          if (T2.GetValue('PPU_NUMERATTRENT') <> ChpEntete.NTrent) then T2.PutValue('PPU_NUMERATTRENT', ChpEntete.NTrent);
          if (T2.GetValue('PPU_DENOMINTRENT') <> ChpEntete.DTrent) then T2.PutValue('PPU_DENOMINTRENT', ChpEntete.DTrent);
        end;
        // FIN PT114
        FTrentieme.Text := T2.GetValue('PPU_NUMERATTRENT');
        ChpEntete.NTrent := T2.GetValue('PPU_NUMERATTRENT');
        FTrentDenominateur.Text := T2.GetValue('PPU_DENOMINTRENT');
        ChpEntete.DTrent := T2.GetValue('PPU_DENOMINTRENT');
      end
      else
      begin
        ChpEntete.TrentMod := FALSE;
        ChbxTrentiem.Checked := FALSE;
        FTrentieme.Enabled := FALSE;
        FTrentDenominateur.Enabled := FALSE;
      end;
      if T2.GetValue('PPU_HORAIREMOD') = 'X' then
      begin
        ChbxHoraireSalarie.checked := TRUE;
        EdtHorReel.Enabled := TRUE;
        EdtJoursOuvres.Enabled := TRUE;
        EdtJoursOuvrables.Enabled := TRUE;
        EdtHorReel.Text := DoubleToCell(T2.GetValue('PPU_HEURESREELLES'), 2);
        ChpEntete.HeuresTrav := T2.GetValue('PPU_HEURESREELLES'); //PT71-4
        EdtJoursOuvres.Text := DoubleToCell(T2.GetValue('PPU_JOURSOUVRES'), 2);
        EdtJoursOuvrables.Text := DoubleToCell(T2.GetValue('PPU_JOURSOUVRABLE'), 2);
      end
      else
      begin
        ChbxHoraireSalarie.checked := FALSE;
        EdtHorReel.Enabled := FALSE;
        EdtJoursOuvres.Enabled := FALSE;
        EdtJoursOuvrables.Enabled := FALSE;
        //DEB PT30-2 en modif bulletin affectation des éléments sur calcul théoriques si non coché horaire personalisé
        EdtHorReel.Text := DoubleToCell(HTHTRAVAILLES, 2);
        {  DEB PT47-4 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables
           Contrairement aux jours les heures ne sont pas sauvegardés dans la fiche ni dans la base..
           EdtJoursOuvres.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,TRUE),2);
           EdtJoursOuvrables.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,FALSE),2);}
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, TRUE, Val1, Val2);
        EdtJoursOuvres.Text := DoubleToCell(Val1, 2);
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, FALSE, Val1, Val2);
        EdtJoursOuvrables.Text := DoubleToCell(Val1, 2);
        //FIN PT30-2 & PT47-4
      end;

    end;
    if TOB_Rub = nil then
    begin
// Flux optimisé      Ferme(Q);
      if Assigned(TPR) then TPR.Free;
    end;
    // end;
    if ActionBul = taModification then
    begin
      if (T_ETAB.GetValue('ETB_JEDTDU') <> 0) and (T_ETAB.GetValue('ETB_JEDTAU') <> 0) then
      begin // les dates sont calculées et accessibles
        FEDTDU.Enabled := TRUE;
        FEDTAU.Enabled := TRUE;
      end;
    end;

    if ActionBul = taCreation then
    begin
      // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
      FEDTDU.Text := FDate1.Text; // On met des valeurs par défaut avant de voir ce que l'on fait
      FEDTAU.Text := FDate2.Text;
      if BullCompl <> 'X' then
      begin
        FDate1.Enabled := FALSE;
        FDate2.Enabled := FALSE;
      end
      else
      begin
        FDate1.Enabled := TRUE;
        FDate2.Enabled := TRUE;
      end;

      if (T_ETAB <> nil) and IsValidDate(FEDTDU.Text) and IsValidDate(FEDTAU.Text) then
      begin
        if (T_ETAB.GetValue('ETB_JEDTDU') <> 0) and (T_ETAB.GetValue('ETB_JEDTAU') <> 0) then
        begin // les dates sont calculées et accessibles
          FEDTDU.Enabled := TRUE;
          FEDTAU.Enabled := TRUE;
          LD := StrToDate(FEDTDU.Text);
          LF := StrToDate(FEDTAU.Text);
          CalculeDateEdtPaie(LD, LF, T_Etab);
          FEDTDU.Text := DateToStr(LD);
          FEDTAU.Text := DateToStr(LF);
        end;
      end;
      ChbxHoraireSalarie.checked := FALSE;
      EdtHorReel.Enabled := FALSE;
      EdtJoursOuvres.Enabled := FALSE;
      EdtJoursOuvrables.Enabled := FALSE;
      //DEB PT30-2 En création calcul des éléments selon le calendrier salarié et non
      //par la gestion des jours (TABLE JOURTRAV) jamais utilisé
      EdtHorReel.Text := DoubleToCell(HTHTRAVAILLES, 2);
      { DEB PT47-4 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables
        Contrairement aux jours les heures ne sont pas sauvegardés dans la fiche ni dans la base..
        EdtJoursOuvres.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,TRUE),2);
        EdtJoursOuvrables.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,FALSE),2);}
      CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, TRUE, Val1, Val2);
      EdtJoursOuvres.Text := DoubleToCell(Val1, 2);
      CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, FALSE, Val1, Val2);
      EdtJoursOuvrables.Text := DoubleToCell(Val1, 2);
      //FIN PT47-4
       {  RendMoisAnnee (DateF, Mois, Annee); Mise en commentaire Table JOURTRAV non utilisé
      T1:=RendCalendrierMois (Etab, Annee, Mois, TOB_JourTrav);
      if T1 <> NIL then
       begin
       EdtHorReel.Text:=DoubleToCell (T1.GetValue('PJT_HEURESREELLES'),2);
       EdtJoursOuvres.Text:=DoubleToCell (T1.GetValue('PJT_JOURSOUVRES'),2);
       EdtJoursOuvrables.Text:=DoubleToCell (T1.GetValue('PJT_JOURSOUVRABLE'),2);
       end;
      FIN PT30-2    }
    end;

    if ChbxTrentiem.Checked = FALSE then LanceCalculTrentieme;
  end;

  if ActionBul = taConsultation then
  begin
    BDelete.visible := FALSE;
    BDelete.Enabled := FALSE;
  end
  else
  begin
    BDelete.visible := TRUE;
    BDelete.Enabled := TRUE;
  end;

  FDate1.Enabled := FALSE;
  FDate2.Enabled := FALSE;
  if ActionBul = taCreation then
  begin
    FDate1.Enabled := TRUE;
    FDate2.Enabled := TRUE;
  end;
  if TOB_HistoBasesCot <> nil then
  begin
    TOB_HistoBasesCot.Free;
    TOB_HistoBasesCot := nil;
  end;
  ChargeBasesSal(CodeSalarie, Etab, DateD, DateF);
  FicheLoad := true;

  if BullCompl = 'X' then CBModifAcquis.Enabled := FALSE; // PT103
  if ActionBul = taConsultation then CBModifAcquis.Enabled := FALSE; // PT104
end;
// Procedure de lancement du calcul du trentieme en fonction des dates de la paie

procedure TFSaisBul.LanceCalculTrentieme;
var
  DateDebut, DateFin: TDateTime;
  Trentieme: Integer;
begin
  DateDebut := StrToDate(FDate1.text);
  DateFin := StrToDate(FDate2.Text);
  Trentieme_entree := FTrentieme.Text; //PT79-2
  Denominateur_Entree := FTrentDenominateur.text; //PT79-2
  Trentieme := CalculTrentieme(DateDebut, DateFin, TRUE); // PT107
  FTrentieme.Text := IntToStr(Trentieme);
  // Affichage jours cp acquis ds bulletin
  {DEB PT47-3 Mise en commentaire
  On recalcul les CP lorsqu'on décoche le trentième forcé
  Redescendu dans la procédure évenementielle click de la coche trentième : ChbxTrentiemeClick
  Appel de la procedure RecalculCP..!!Prise de congé sur acquis en cours
  if Tob_salarie <> nil then
     begin
     T := Tob_etablissement.findfirst(['ETB_ETABLISSEMENT'],[Tob_salarie.GetValeur (iPSA_ETABLISSEMENT)],false);
     if T <> nil then
        begin
        Mois:=0;Supp:=0;Anc:=0;
  // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
        if BullCompl <> 'X' then
         val:=AlimCongesAcquis(TOB_Salarie,T,DateDebut,DateFin,Trentieme,30,TOB_RUB,false,base,nbmois,suivant,Mois,Supp,ANc); // Génération cp acquis
        CpMois.text := floattostr(Mois);
        CpSupp.text := floattostr(Supp);
        CpAnc.text  := floattostr(Anc);
        end;
     end;
  FIN PT47-3}
  CalAffectTrent;
end;


procedure TFSaisBul.GBasRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if ACal = TRUE then CalculBullSaisie;
  ACal := FALSE;
  Modifier := TRUE;
end;
{ Gestion automatique des lignes de commentaire
Activation par le bouton ligne de commentaire avec recherche de la derniere
ligne de commenatire utilisée, Puis reprise des fonctionnalités du bouton création de
ligne pour avoir les actions de création d'objet ...
}

procedure TFSaisBul.LigneCommentaire(Sender: TObject);
var
  Num: Integer;
  Rub, St1: string;
begin
  St1 := '';
  if ActionBul = taConsultation then exit;
  if (CurG.Row < 1) then exit;
  St1 := CurG.Cells[0, CurG.Row];
  // PT17 : 26/12/2001 V571 PH Insertion des lignes de commentaire multiples
  if St1 = '' then
  begin
    ShowMessage('Vous ne pouvez insérer de ligne de commentaire sans rubrique !');
    exit;
  end;
  Num := RendRubrCommentaire(St1, CurG);
  if (Num = 0) or (Num > 9) then //PT179
  begin
    ShowMessage('Vous ne pouvez avoir plus de 9 lignes de commentaire');
  end;
  if (Num <= 9) and (Num > 0) then //PT179
  begin
    // bInsererligneClick (Sender);
    bNewLigneClick(Sender);
    CurG.ElipsisButton := FALSE;
    Rub := Copy(St1, 1, 4) + '.' + IntToStr(Num);
    CurG.Cells[0, CurG.Row] := Rub;
    CurG.Col := 1;
  end;
end;

procedure TFSaisBul.GRemSalRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if (ActionLigne = taCreation) and (VH_Paie.PGAnalytique = TRUE) then
    AffectePreVentilAna; // en insertion rajout des pré ventilations analytiques
  ActionLigne := taConsultation;
  if ACal = TRUE then
  begin
    PGTopRecalculCp := True;
    CalculBullSaisie;
  end; { PT92-1 17/05/2004 }
  ACal := False;
  Modifier := TRUE;
end;

{ Fonction qui remplit les onglets en fonction de la liste des rubriques du bulletin
}

procedure TFSaisBul.RubDansGrille;
var
  i, NbD, LigCour: Integer;
  T, TRech: TOB;
  Nat, Them, St, Rubriq, ChampCsv, AccessSaisie: string;
  Grille: THGrid;
begin
  Nat := '';
  Them := '';
  St := '';
  Rubriq := '';
  Grille := CurG;
  if Grille <> nil then LigCour := Grille.Row
  else LigCour := 1;
  RazGrilles;
  // alimentation des grilles de saisie en fonction de la TOB de la liste des rubriques du bulletins
  for i := 0 to Tob_Rub.Detail.Count - 1 do
  begin
//    CurG := nil; // PT176
    T := Tob_Rub.Detail[i];
    if Assigned(T) and (iPHB_NATURERUB = 0) then MemorisePhb(T);
    with T do
    begin
      Nat := GetValeur(iPHB_NATURERUB);
      Rubriq := RendCodeRub(GetValeur(iPHB_RUBRIQUE));
      if (Nat = '') or (Rubriq = '') then continue;
      if Nat = 'AAA' then
      begin
        if GetValeur(iPHB_COTREGUL) = 'REG' then
          if Pos('.R', CurG.Cells[0, CurG.Row]) <= 0 then
            CurG.Cells[0, CurG.Row] := GetValeur(iPHB_RUBRIQUE) + '.R';
        if (POS('.R', Rubriq) > 0) then
          rubriq := copy(rubriq, 1, 4);
        TRech := Paie_RechercheOptimise(tob_rem, 'PRM_RUBRIQUE', rubriq);
        //TRech:=TOB_Rem.FindFirst (['PRM_RUBRIQUE'], [Rubriq], TRUE) ;
        if assigned(trech) and (iPRM_THEMEREM = 0) then MemorisePrm(Trech);
        if not assigned(trech) then
        begin
          Them := '';
          St := 'La rémunération ' + GetValeur(iPHB_RUBRIQUE) + ' a été utilisée dans le bulletin mais n''existe plus ! ';
          ShowMessage(St);
          continue;
        end
        else Them := TRech.GetValeur(iPRM_THEMEREM);

        if BtnDetail.Down = FALSE then
        begin
          CurG := GRemSal;
          if Them = 'RNI' then CurG := GRemNonImp
          else if Them = 'RSS' then CurG := GRemRet;
          // DEB PT202
          if BtnSaisie.down then
          begin
            if RemSaisissable(Trech) then
              CurG := GRemSais; // Pour affecter systématiquement les rubriques saisissables dans la bonne grille
          end;
        end
        else
        begin // PT120 Au moins pour ranger la rubrique thème Net à Payer dans un onglet
          if BtnSaisie.down then
          begin
            if RemSaisissable(Trech) then
              CurG := GRemSais; // Pour affecter systématiquement les rubriques saisissables dans la bonne grille
          end
          else // FIN PT202
          begin
            if (Them = 'SAL') or (Them = 'ZZZ') then CurG := GRemSal
            else if Them = 'ABS' then CurG := GRemAbs
            else if Them = 'ABT' then CurG := GRemAbt
            else if Them = 'AVT' then CurG := GRemAvt
            else if Them = 'RNI' then CurG := GRemNonImp
            else if Them = 'INI' then CurG := GRemPrimes
            else if Them = 'RSS' then CurG := GRemRet
            // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
            else if Them = 'RSC' then CurG := GRemNonImpSoumis
            // FIN PT3
            else if Them = 'COM' then CurG := GRemCplt
            else if Them = 'HEU' then CurG := GRemHeures;
          end;
        end;
      end;

      if Nat = 'BAS' then CurG := GBas
      else if Nat = 'COT' then CurG := GCot;

      if CurG <> nil then
      begin
        if CurG.Row >= CurG.RowCount - 1 then CurG.RowCount := CurG.RowCount + 10;
        CurG.Cells[0, CurG.Row] := GetValeur(iPHB_RUBRIQUE);
        AccessSaisie := GetValeur(iPHB_ORIGINELIGNE);
        CurG.Objects[0, CurG.Row] := T;
        CurG.Cells[1, CurG.Row] := GetValeur(iPHB_LIBELLE);
        if (CurG <> GCot) and (CurG <> GBas) then // Les differentes Remunerations
        begin
          if Assigned(TRech) then NbD := TRech.GetValeur(iPRM_DECBASE)
          else NbD := 2;
          CurG.Cells[2, CurG.Row] := DoubleToCell(GetValeur(iPHB_BASEREM), NbD);
          if Assigned(TRech) then NbD := TRech.GetValeur(iPRM_DECTAUX)
          else NbD := 2;
          CurG.Cells[3, CurG.Row] := DoubleToCell(GetValeur(iPHB_TAUXREM), NbD);
          if Assigned(TRech) then NbD := TRech.GetValeur(iPRM_DECCOEFF)
          else NbD := 2;
          CurG.Cells[4, CurG.Row] := DoubleToCell(GetValeur(iPHB_COEFFREM), NbD);
          if Assigned(TRech) then NbD := TRech.GetValeur(iPRM_DECMONTANT)
          else NbD := 2;
          CurG.Cells[5, CurG.Row] := DoubleToCell(GetValeur(iPHB_MTREM), NbD);
          ChampCsv := GetValeur(iPHB_CONSERVATION);
          CurG.Cells[6, CurG.Row] := RechDom('PGACTIONBULLETIN', ChampCsv, FALSE);
        end
        else if CurG = GCot then // Les Cotisations
        begin
          if GetValeur(iPHB_COTREGUL) = 'REG' then
            if Pos('.R', CurG.Cells[0, CurG.Row]) <= 0 then
              CurG.Cells[0, CurG.Row] := GetValeur(iPHB_RUBRIQUE) + '.R';
          if (POS('.R', Rubriq) > 0) then
            rubriq := copy(rubriq, 1, 4);
          TRech := Paie_RechercheOptimise(TOB_cotisations, 'PCT_RUBRIQUE', rubriq);
          //  optimisation         TRech:=TOB_Cotisations.FindFirst (['PCT_RUBRIQUE','PCT_NATURERUB'], [Rubriq,'COT'], TRUE) ;
          if assigned(trech) and (iPCT_DECBASE = 0) then MemorisePct(Trech);
          NbD := TRech.GetValeur(iPCT_DECBASE);
          CurG.Cells[2, CurG.Row] := DoubleToCell(GetValeur(iPHB_BASECOT), NbD);
          NbD := TRech.GetValeur(iPCT_DECTXSAL);
          CurG.Cells[3, CurG.Row] := DoubleToCell(GetValeur(iPHB_TAUXSALARIAL), NbD);
          NbD := TRech.GetValeur(iPCT_DECMTSAL);
          CurG.Cells[4, CurG.Row] := DoubleToCell(GetValeur(iPHB_MTSALARIAL), NbD);
          NbD := TRech.GetValeur(iPCT_DECTXPAT);
          CurG.Cells[5, CurG.Row] := DoubleToCell(GetValeur(iPHB_TAUXPATRONAL), NbD);
          NbD := TRech.GetValeur(iPCT_DECMTPAT);
          CurG.Cells[6, CurG.Row] := DoubleToCell(GetValeur(iPHB_MTPATRONAL), NbD);
          CurG.Cells[7, CurG.Row] := RechDom('PGACTIONBULLETIN', GetValeur(iPHB_CONSERVATION), FALSE);
        end
        else if CurG = GBas then // Les bases de cotisations
        begin
          TRech := Paie_RechercheOptimise(TOB_Bases, 'PCT_RUBRIQUE', rubriq);
          // Optimisation TRech:=TOB_Bases.FindFirst (['PCT_RUBRIQUE','PCT_NATURERUB'], [Rubriq,'BAS'], TRUE) ;
          if assigned(trech) and (iPCT_DECBASECOT = 0) then MemorisePct(Trech);
          NbD := TRech.GetValeur(iPCT_DECBASECOT);
          CurG.Cells[2, CurG.Row] := DoubleToCell(GetValeur(iPHB_BASECOT), NbD);
          CurG.Cells[3, CurG.Row] := DoubleToCell(GetValeur(iPHB_PLAFOND), NbD);
          CurG.Cells[4, CurG.Row] := DoubleToCell(GetValeur(iPHB_TRANCHE1), NbD);
          CurG.Cells[5, CurG.Row] := DoubleToCell(GetValeur(iPHB_TRANCHE2), NbD);
          CurG.Cells[6, CurG.Row] := DoubleToCell(GetValeur(iPHB_TRANCHE3), NbD);
          CurG.Cells[7, CurG.Row] := DoubleToCell(GetValeur(iPHB_PLAFOND1), NbD);
          CurG.Cells[8, CurG.Row] := DoubleToCell(GetValeur(iPHB_PLAFOND2), NbD);
          CurG.Cells[9, CurG.Row] := DoubleToCell(GetValeur(iPHB_PLAFOND3), NbD);
          //PT68   05/06/2003 V_421 PH FQ 10700 Gestion origine de la rubrique dans les bases de cotisations
          CurG.Cells[10, CurG.Row] := RechDom('PGACTIONBULLETIN', GetValeur(iPHB_CONSERVATION), FALSE);
        end;
        CurG.Row := CurG.Row + 1;
      end;
    end;
  end;
  ReajustGrille;
  CurG := Grille;
  if CurG <> nil then CurG.Row := LigCour
  else Pages.ActivePage := PRemSal;
  PagesChange(nil);
end;

procedure TFSaisBul.BtnDetailClick(Sender: TObject);
var
  CurG: THGrid;
begin
  GrilleDansRub;
  if BtnDetail.Down = TRUE then
  begin
    AffecteDetail;
  end
  else
  begin
    AffecteSimple;
  end;
  RubDansGrille;
  Pages.ActivePage := PRemSal;
  Curg := GRemSal;
  // DEB PT202
  if BtnSaisie.down then
  begin
    PRemSais.TabVisible := TRUE;
    Pages.ActivePage := PRemSais;
    CurG := GRemSais;
  end
  else
  begin
    Pages.ActivePage := PRemSal;
    CurG := GRemSal;
    PRemSais.TabVisible := FALSE;
  end;
  // FIN PT202
  BCalculBullClick(nil); // PT207
  if Curg <> nil then Curg.Row := 1;
  Pages.Refresh;
  PagesChange(nil);
end;

procedure TFSaisBul.BtnSimpleClick(Sender: TObject);
var
  CurG: THGrid;
begin
  GrilleDansRub;
  AffecteSimple;
  RubDansGrille;
  // DEB PT202
  if BtnSaisie.down then
  begin
    Pages.ActivePage := PRemSais;
    CurG := GRemSais;
  end
  else
  begin
    Pages.ActivePage := PRemSal;
    CurG := GRemSal;
  end;
  // FIN PT202
  if Curg <> nil then Curg.Row := 1;
  Pages.Refresh;
  PagesChange(nil);
end;
{Fonction qui remplit la Tob de la liste des rubriques du bulletins à partir des
grilles de données.
}

procedure TFSaisBul.GrilleDansRub;
var
  i: Integer;
  Nature: string;
  Grille: THGrid;
begin
  // RAZ des lignes bulletins de la TOB TPE
  //RazTPELignes(TPE);
  CreationTOBCumSal;
  // Pour chq grille, on alimente la liste des rubriques
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is THGrid then
    begin
      Grille := THGrid(Components[i]);
      if (Grille = GRemSal) or (Grille = GRemSais) or (Grille = GRemAbs) or (Grille = GRemAbt) or (Grille = GRemAvt)
        or (Grille = GRemNonImp) or (Grille = GRemPrimes) or (Grille = GRemRet) or (Grille = GRemCplt)
        or (Grille = GRemHeures) or (Grille = GRemNonImpSoumis) then Nature := 'AAA'; // PT202
      if Grille = GBas then Nature := 'BAS';
      if Grille = GCot then Nature := 'COT';
      GrilleAlimLignes(TOB_RUB, Etab, CodeSalarie, Nature, DateD, DateF, Grille, ActionBul);
      Nature := 'COT';
    end;
  end;
  // Pour Chaque ligne de grille recupérer les cellules et créer un objet TOB
end;
{ Fonction Initialisation du bulletin salarié
}

procedure TFSaisBul.ChargeBulletin;
var
  DateDebut, DateFin: TDateTime;
  StTemp: string;
// d PT115
  GenereMaintien, IJSS: Boolean;
{$IFNDEF CCS3}
  mode, ChampCateg: string;
  anomalie: integer;
  TR: TOB;
  Q: TQuery;
{$ENDIF}
// f PT115
  Tob_IJ: TOB; // PT119
begin
// d PT115
  GenereMaintien := False;
  IJSS := False;
// f PT115

  InitialiseVariableStatCP;
  InitVarTHTRAVAILLES; //PT71-4
  LibereTobCp(TR); { PT173 }
  if Tob_Rub <> nil then
  begin
    Tob_Rub.Free;
    Tob_Rub := nil;
  end; // ph memcheck
  if TOBAna <> nil then
  begin
    TOBAna.Free;
    TOBAna := nil;
  end;

  //DEB PT181 Chargement ici et non plus dans SaisieBulletin car sinon ne met pas à jour avec les boutons de navigation
  // chargement des la liste des rubriques en foction des profils
  DateDebut := StrToDate(FDate1.text);
  DateFin := StrToDate(FDate2.Text);
  if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
  begin
    Nettoyage_EltDynSal();
    Nettoyage_EltDynPop();
    Nettoyage_EltDynEtab();
    initTOB_EltDynSal(CodeSalarie, DateFin);
    initTOB_EltDynEtab(Etab, DateFin);
    initTOB_EltDynPop(CodeSalarie, DateFin);
  end;
  //FIN PT181

// d PT115
{$IFNDEF CCS3}
  // Gère-t'on le maintien ?
  if VH_Paie.PgMaintien then
  begin
    GenereMaintien := True;
    IJSS := False;
  end;
// récupération du champ critère suplémentaire IJSS et Maintien
  if (VH_Paie.PGCritMaintien <> '') then
  begin // PT125 Suppression des "" pour un integer en DB2
    Q := OpenSQL('SELECT PAI_PREFIX, PAI_SUFFIX FROM PAIEPARIM WHERE PAI_IDENT = ' +
      VH_Paie.PGCritMaintien, TRUE);
    if not Q.EOF then
    begin
      ChampCateg := Q.Fields[0].AsString + '_' + Q.Fields[1].AsString;
    end
    else
      ChampCateg := '';
    ferme(Q);
  end;
{$ENDIF}
// f PT115

{  // chargement des la liste des rubriques en foction des profils
  DateDebut := StrToDate(FDate1.text);
  DateFin := StrToDate(FDate2.Text);}//PT181 déplacement plus haut
  //JTHTRAVAILLES := CalculJoursTheorique(Tob_sal,datedebut,datefin);  //PT30-1 Déplacement du code dans RecupInfoModules pour passage prep. auto
  CreationTOBCumSal; // Creation de la TOB des Cumuls Salariés
  // PT76   26/08/2003 V_421 PH Mise en place de la saisie arret
  FreeAndNil(TOB_SaisieArret);
  // Gestion des Modules Complémentaires : CP, Gestion des Absences, Saisie par rubrique
  // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  Tob_Rub := ChargeRubriqueSalarie(Tob_Sal, DateDebut, DateFin, ActionBul, 0, ProfilSpec);
  TopRecalculCPBull := False; //PT71-1
// d PT115
{$IFNDEF CCS3}
  // vérif. si maintien à intégrer
  if (GenereMaintien) then
  // On gère le maintien
  begin
    if ActionBul = TaCreation then Mode := 'C' else Mode := 'M';
//PT123  TR :=  (RecupereRegltIJSS(Datef, CodeSalarie, Mode));
    TR := (RecupereRegltIJSS(Datef, CodeSalarie, Mode, ChampCateg, nil, nil)); // PT128
    if not RecupereAbsencesIJSS(Datef, CodeSalarie, Mode) and
      (TR.Detail.Count <= 0) then
    // Pas d'absence ni de règlement d'IJSS pour ce salarié donc pas
    // de maintien à intégrer
    begin
      GenereMaintien := False;
    end;
    FreeAndNil(TR); //PT154 memcheck
  end;
{$ENDIF}
// f PT115

// d PT111 PT115 PT119
  // if BullCompl <> 'X' then RecupInfoModules(TOB_Rub, Tob_Sal, ActionBul, False, False, StTemp); // PT7
  // if BullCompl <> 'X' then RecupInfoModules(TOB_Rub, Tob_Sal, ActionBul, False, False, StTemp, False, False);
  // if BullCompl <> 'X' then RecupInfoModules(TOB_Rub, Tob_Sal, ActionBul, False, False, StTemp, GenereMaintien, IJSS);
//  if BullCompl <> 'X' then
  RecupInfoModules(TOB_Rub, Tob_Sal, ActionBul, False, False, StTemp, GenereMaintien, IJSS, Tob_IJ);
//f  PT111 PT115 PT119

    { PT71-4 En chargement si horaire non personnalisé on affecte les Heures trav à EdtHorReel.text
    EdtHorReel.text => ChpEntete.heurestrav => Variable 36 par la suite du traitement }
  if not ChbxHoraireSalarie.checked then EdtHorReel.text := DoubleToCell(HTHTRAVAILLES, 2);
  if ActionBul = taCreation then RubDansGrille;
  CalculBullSaisie;
// d PT115
{$IFNDEF CCS3}
// PT174 30/11 if GenereMaintien then
  if (GenereMaintien) and (ActionBul <> taModification) then
  // maintien à calculer et intégrer
  begin
    InitMoveProgressForm(nil, 'Calcul du maintien en cours ...', '', 50, FALSE, TRUE);
    MoveCurProgressForm('Veuillez patienter SVP ...');

    // calcul et maj Maintien   (paie envers...)
    MaintienDuSalarie(Tob_salarie, TOB_Rub, Tob_etablissement, Tob_Abs, DateToStr(Datef), DateToStr(dated), Mode, Anomalie, ChampCateg, TOB_IJSS, TOB_Maintien);

    MoveCurProgressForm('Calcul terminé');

    if (anomalie <> 0) then
    begin
      if (Anomalie = 1) then
        PGIBox('Pas de règle de Maintien pour ce salarié', Caption)
      else
        if (Anomalie = 2) then
          PGIBox('Calcul du montant du maintien,lecalcul de la paie à l''envers a échoué', Caption)
//d PT155
        else
          if (Anomalie = 3) then
// PT158          PGIBox('La rubrique de garantie du salaire n''a pas été paramétrée', Caption);
            PGIBox('Vérifiez le maintien de salaire et les ijss, la rubrique de garantie de salarie n''a pu être calculée ou est absente', Caption)
//f PT155
//d PT225
            else
              if (Anomalie = 4) then
                PGIBox('La date d''ancienneté du salarié n''est pas renseignée : le maintien ne peut pas être calculé', Caption);
// f PT225

    end;
    RubDansGrille;
    // intégration abs IJSS, IJSS, Maintien
    CalculBullSaisie;

    FiniMoveProgressForm;
  end;
{$ENDIF}
// f PT115

//  if (TopRecalculCPBull) and (RendCumulSalSess('12') <> 0)  then RecalculCp; //PT71-1 { PT127-2 Mise en commentaire provoque des anomalies sur mvt SLD )

end;
// ¨procedure qui recupere et affiche les cumuls du bas de bulletin

procedure TFSaisBul.AfficheBasBulletin;
var
  St, Mask: string;
  mt: double;
begin
  Mask := '#,##0.00';
  mt := RendCumulSalSess('01');
  St := FormatFloat(Mask, mt);
  LblBrut.Caption := St;
  mt := RendCumulSalSess('02');
  St := FormatFloat(Mask, mt);
  LblBrutFiscal.Caption := St;
  mt := RendCumulSalSess('09');
  St := FormatFloat(Mask, mt);
  LblNetImp.Caption := St;
  mt := RendCumulSalSess('10');
  St := FormatFloat(Mask, mt);
  LblNetAPayer.Caption := St;
  mt := RendCumulSalSess('20');
  St := FormatFloat(Mask, mt);
  LblHeures.Caption := St;
end;

procedure TFSaisBul.BinsertClick(Sender: TObject);
begin
  NewEnreg;
  ChargeBulletin; // PT141 Rechargement des infos salariés
end;

procedure TFSaisBul.FDate1Exit(Sender: TObject);
begin
  if not IsValidDate(FDate1.Text) then
  begin
    Showmessage('La date de début n''est pas valide.');
    FDate1.SetFocus;
    Date1 := 0;
    ReactiveBtn;
    exit;
  end;
  //PT65-2 Le calcul bulletin s'effectue sortie de date fin bulletin
  if (StrToDate(FDate1.Text) <> DDOnEnter) then FDate2Exit(FDate2);
  //PT65-1   16/01/2003 V591 PH FQ 10449 Désactivation insertion et suppression de ligne pdt la modif des dates de paie
  ReactiveBtn;
end;

procedure TFSaisBul.FDateDuExit(Sender: TObject);
begin
  if not IsValidDate(FEDTDU.Text) then
  begin
    Showmessage('La date de début d''édition n''est pas valide.');
    FEDTDU.SetFocus;
    exit;
  end;
  ChpEntete.Edtdu := StrToDate(FEDTDU.Text);
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.BValiderClick(Sender: TObject);
var
  rep: Integer;
  LDD, LDF: TDateTime;
begin
  if ActionBul = taConsultation then exit;
  //Old:=Modifier ;
  Modifier := FALSE;
  // if Not Bouge(nbPost) then Modifier:=Old;

  Rep := MSG.execute(0, Caption, '');
  if rep = mrNo then
  begin
// d PT118
{$IFNDEF CCS3}
    // il faut annuler les lignes de maintien générées au fur et à mesure
    // du traitment du maintien.
    if (ActionBul = taCreation) then // PT138
      AnnuleMaintienBulletin(CodeSalarie, DateD, DateF);
{$ENDIF}
// f PT118
    Close;
  end;
  if rep = mrCancel then exit; // retour etat initial   EtatInitial
  if (rep = mrYes) and (ActionBul = taModification) then
  begin
    // OnDelete; // destruction des enregistrements à l'origine du bulletin car on est en modification
    // GrilleDansRub; // Recopie des grilles dans la tob des rubriques.
    CalculAFaire := TRUE;
    CalculBullSaisie;
    OnSauve;
    Modifier := FALSE; // PT135-1
  end;
  if (rep = mrYes) and (ActionBul = taCreation) then
  begin
    LDD := StrToDate(FDate1.Text);
    LDF := StrToDate(FDate2.Text);
    // PT6 : 10/09/2001 V547 PH Controle de la date de la paie en création de bulletin
    if (not ControlPaieCloture(LDD, LDF)) and (ActionBul = taCreation) then
    begin
      Showmessage('Vous ne pouvez pas saisir de paie sur une période close');
      FDate1.SetFocus;
    end
    else
      // GrilleDansRub; // Recopie des grilles dans la tob des rubriques.
    begin
      CalculAFaire := TRUE;
      CalculBullSaisie;
      OnSauve;
      ActionBul := taModification;
      ActionBulCP := taModification; { PT127-2 }
      AfficheActionB;
      FDate1.Enabled := FALSE;
      FDate2.Enabled := FALSE;
    end;
    // FIN PT6
  end;

end;

procedure TFSaisBul.BDeleteClick(Sender: TObject);
var
  rep: Integer;
begin
  if (ActionBul = taCreation) or (ActionBul = taConsultation) then exit;
  Rep := MSG.execute(1, Caption, '');
  if rep = mrNo then exit;
  if rep = mrCancel then EtatInitial; // retour etat initialif
  if (ActionBul = taModification) and (rep = mrYes) then
  begin
    ActionBul := taSuppression;
    Bouge(nbDelete); // Suppression physique du bulletin
    LibereTobAbs; // PT116
    LibereTobCalendrier; { PT160 }
// d PT119
    if Assigned(Tob_IJSS) then
      FreeAndNil(Tob_IJSS);
    if Assigned(Tob_Maintien) then
      FreeAndNil(Tob_Maintien);
// f PT119
    FreeAndNil(Tob_PSD); { PT167 }
    FreeAndNil(Tob_PSP); { PT167 }

    Close; // Sortie de la fenetre pour retourner au multi critere de selection des paies
  end;
end;

procedure TFSaisBul.bDefaireClick(Sender: TObject);
var
  rep: Integer;
begin
  if ActionBul = taConsultation then exit;
  Rep := PGIAsk('Cette action provoque l''abandon de vos modifications et le rechargement du bulletin', 'Retour à l''état initial de la paie');
  //Rep:=MSG.execute(2,Caption,'');
  if rep = mrNo then exit;
  if rep = mrYes then
  begin // PT97
    ForceAlignProfil := TRUE;
    Pages.ActivePage := PRemSal;
    Etatinitial;
  end;
end;

procedure TFSaisBul.BFermeClick(Sender: TObject);
begin
  if ActionBul = taCreation then
  begin
    BValiderClick(Sender);
  end;
  if ActionBul = taModification then
  begin
    if Modifier = TRUE then
    begin // Modification detectee alors Action Valider
      BValiderClick(Sender);
      Modifier := FALSE;
    end
    else Close; // Pas de Modif Alors Sortie
  end;
  LibereTobCp(Tob_Pris);
  LibereTobAbs; //PT66 Ajout proc
  LibereTobCalendrier; { PT160 }

{$IFNDEF CCS3}
  // d PT94 IJSS   libère TOB
  if Assigned(Tob_IJSS) then
    FreeAndNil(Tob_IJSS);
  // f PT94 IJSS
{$ENDIF}

{$IFNDEF CCS3}
  // d PT110 Maintien   libère TOB
  if Assigned(Tob_Maintien) then
    FreeAndNil(Tob_Maintien);
  // f PT110 Maintien
{$ENDIF}

  FreeAndNil(Tob_PSD); { PT167 }
  FreeAndNil(Tob_PSP); { PT167 }

  InitialiseVariableStatCP; { PT87 }
  // PT4 : 03/09/2001 V547 PH Correction zones et boutons non accessibles en consultation
  if ActionBul = taConsultation then Close;
end;

// Attention, le navigateur est supprimé en eAGL

procedure TFSaisBul.BFirstClick(Sender: TObject);
begin
  Bouge(nbFirst);
  BLast.Enabled := TRUE;
  BNext.Enabled := TRUE;
  BFirst.Enabled := FALSE;
  BPrev.Enabled := FALSE;
end;

procedure TFSaisBul.BPrevClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  Bouge(nbPrior);
  if QMul.BOF then
  begin
    BFirst.Enabled := FALSE;
    BPrev.Enabled := FALSE;
  end;
  BLast.Enabled := TRUE;
  BNext.Enabled := TRUE;
{$ENDIF}
end;

procedure TFSaisBul.BNextClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  Bouge(nbNext);
  if QMul.EOF then
  begin
    BLast.Enabled := FALSE;
    BNext.Enabled := FALSE;
  end;
  BFirst.Enabled := TRUE;
  BPrev.Enabled := TRUE;
{$ENDIF}
end;

procedure TFSaisBul.BLastClick(Sender: TObject);
begin
  Bouge(nbLast);
  BLast.Enabled := FALSE;
  BNext.Enabled := FALSE;
  BFirst.Enabled := TRUE;
  BPrev.Enabled := TRUE;
end;
// Fonction de gestion d'un navigateur car on navique soit sur des paies déjà faites
// soit sur des paies à faire donc des infos inexistantes !!!!!!

function TFSaisBul.Bouge(Button: TNavigateBtn): boolean;
begin
  result := FALSE;
  case Button of
    nblast, nbprior, nbnext,
      nbfirst, nbinsert:
      begin
        if CalculAFaire then CalculBullSaisie;
        if (not OnSauve) and (ActionBul <> taConsultation) then Exit;
      end;
    nbPost: if not EnregOK then Exit;
    nbDelete: if not OnDelete then Exit;
  end;
  // if not TransacNav(DBNav.BtnClick,Button,25) then MessageAlerte(HM.Mess[5]);
  result := TRUE;
  if Button = NbInsert then
  begin
    NewEnreg;
    exit;
  end;
  if Button = NbPost then
  begin
    Modifier := False;
    exit;
  end;
  if Button = nbDelete then
  begin
{$IFNDEF EAGLCLIENT}
    if QMul.EOF = FALSE then
    begin
      QMul.Next;
      if QMul.EOF = TRUE then
      begin
        QMul.prior;
        if QMul.BOF then Close;
      end
    end
    else
    begin
      if QMul.BOF = FALSE then
      begin
        QMul.prior;
        if QMul.BOF = TRUE then Close;
      end;
    end;
{$ENDIF}
  end;
  if QMul <> nil then
  begin
    // PT58   16/12/2002 V591 PH FQ 10317 gestion du navigateur

{$IFNDEF EAGLCLIENT}
    case Button of
      nblast: if not QMul.EOF then QMul.Last else Close;

      nbfirst: if not QMul.BOF then QMul.First else Close;
      nbnext:
        begin
          if not QMul.EOF then QMul.Next else Close;
        end;

      nbprior: if not QMul.BOF then QMul.prior else Close;
    end;
{$ENDIF}
  end;
  if ActionBul <> taSuppression then
    // PT58   16/12/2002 V591 PH FQ 10317 gestion du navigateur
{$IFNDEF EAGLCLIENT}
    // if (Not QMul.Eof AND Not Qmul.Bof then
{$ENDIF}
    GereQuery;
  // FIN PT58
end;

procedure TFSaisBul.NewEnreg;
begin
  FDate1.Enabled := TRUE;
  FDate2.Enabled := TRUE;
  FDate1.SetFocus;
  ActionBul := taCreation;
  { DEB PT79-1 Nouveau Bulletin, donc nouvelle génération CP }
  ActionSld := taCreation;
  ActionBulCP := taCreation;
  InitialiseVariableStatCP;
  { FIN PT79-1 }
  AfficheActionB;
  BInsert.Enabled := False;
  BDelete.Enabled := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 17/06/2004
Modifié le ... :   /  /
Description .. : ATTENTION !!
Suite ........ : A Faire
Suite ........ : Vérif des diff entre OnDelete(SaisBul) et SuppressionBull
Suite ........ : (P5DEF)
Mots clefs ... :
*****************************************************************}

function TFSaisBul.OnDelete: boolean;
var
  Zdate1, Zdate2: TDateTime;
  St, RefA: string;
  Nbre: longint;
  Trace: TStringList;
  Q: TQuery;
  rep: Integer;
begin
  result := False;
  // DEB PT113 Confirmation de suppression de bulletin comptabilisé
  ZDate1 := StrToDate(FDate1.Text);
  ZDate2 := StrToDate(FDate2.Text);
  Q := OpenSQL('SELECT PPU_TOPGENECR FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT="' + UsDateTime(ZDate1) +
    '" AND PPU_DATEFIN="' + UsDateTime(ZDate2) + '"', TRUE);
  if not Q.EOF then
  begin
    st := Q.FindField('PPU_TOPGENECR').AsString;
    if st = 'X' then
    begin
      Ferme(Q);
      rep := PgiAsk('Attention, le bulletin a déjà été comptabilisé. #13#10 Voulez-vous continuer ?', 'Suppression de bulletin');
      if rep <> mrYes then exit;
    end;
  end;
  Ferme(Q);
  // FIN PT113
  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  Trace := TStringList.Create;
  try
    BeginTrans;
    st := '';
    Nbre := 0;
    if ActionBul = taSuppression then
    begin
      st := '';
      Nbre := ExecuteSQL('DELETE FROM HISTOBULLETIN WHERE PHB_ETABLISSEMENT="' + Etab + '" AND PHB_SALARIE="' + CodeSalarie + '" AND PHB_DATEDEBUT="' + UsDateTime(ZDate1) +
        '" AND PHB_DATEFIN="' + UsDateTime(ZDate2) + '"');
      if Nbre >= 0 then
      begin
        // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
        Nbre := ExecuteSQL('DELETE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '" AND PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT="' + UsDateTime(ZDate1) +
          '" AND PPU_DATEFIN="' + UsDateTime(ZDate2) + '"');
        if Nbre >= 0 then
          Nbre := ExecuteSQL('DELETE FROM HISTOCUMSAL WHERE PHC_ETABLISSEMENT="' + Etab + '" AND PHC_SALARIE="' + CodeSalarie + '" AND PHC_DATEDEBUT="' + UsDateTime(ZDate1) +
            '" AND PHC_DATEFIN="' + UsDateTime(ZDate2) + '" AND PHC_REPRISE="-"');
      end;
    end;
    ExecuteSQL('UPDATE HISTOSAISRUB SET PSD_DATEINTEGRAT ="' + usdatetime(0) +
      ' " WHERE PSD_SALARIE = "' + CodeSalarie + '"' +
      ' AND (PSD_DATEINTEGRAT >="' + usdatetime(ZDate1) +
      '" AND PSD_DATEINTEGRAT <="' + usdatetime(ZDate2) + '")');

    if VH_Paie.PGMODULEPRESENCE then
    MiseAjourPresence(Zdate1,Zdate2,Codesalarie,'SUP'); // PT213
    
    //PT81   23/09/2003 V_421 PH Suppression saisie arret en suppression de bulletin
    if Assigned(TOB_SaisieArret) then PGSuppHistoRetenue(CodeSalarie, ZDate1, ZDate2);
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    if BullCompl <> 'X' then AnnuleAbsenceBulletin(CodeSalarie, DateF);

{$IFNDEF CCS3}
    // d PT94 IJSS - Raz date intégration dans REGLTIJSS
    if (BullCompl <> 'X') then
      AnnuleIJSSBulletin(CodeSalarie, DateF);
    // f PT94 IJSS
{$ENDIF}

{$IFNDEF CCS3}
    // d PT110 Maintien- Raz date intégration dans Maintien
    if (BullCompl <> 'X') then
      AnnuleMaintienBulletin(CodeSalarie, DateD, DateF);
    // f PT110 Maintien
{$ENDIF}

    if Nbre < 0 then Showmessage('Une erreur s''est produite pendant l''annulation');
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    if (BullCompl <> 'X') and ((VH_paie.PGCongesPayes) and (GblCP)) then { PT127-4 }
    begin
      AnnuleCongesPris(CodeSalarie, Etab, DateD, DateF); // reaffectation des cp pris et acquis
      AnnuleCongesAcquis(CodeSalarie, DateD, DateF);
    end;
    // Suppression historique analytique
    if VH_Paie.PGAnalytique = TRUE then
    begin
      st := 'DELETE FROM HISTOANALPAIE WHERE PHA_SALARIE="' + CodeSalarie + '" AND PHA_DATEDEBUT="' + USDateTime(DateD) + '" AND PHA_DATEFIN="' + USDateTime(DateF) + '"';
      ExecuteSQL(St);
      RefA := CodeSalarie + ';' // Code Salarie
        + FormatDateTime('ddmmyyyy', DateD) + ';' // Date Debut bulletin
        + FormatDateTime('ddmmyyyy', DateF) + ';'; // Date Fin bulletin
      st := 'DELETE FROM VENTANA WHERE YVA_TABLEANA="PPU" AND YVA_IDENTIFIANT like "' + RefA + '%"'; // suppression des ventilations analytiques du bulletin
      ExecuteSQL(St);
    end;
    // Destruction enregistrement historique salarié
    // PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
    if VH_Paie.PgHistorisation then
      ExecuteSQL('DELETE FROM HISTOSALARIE WHERE PHS_SALARIE="' + CodeSalarie + '" AND PHS_DATEEVENEMENT="' + USDateTime(DateD) + '" AND PHS_DATEAPPLIC="' + USDateTime(DateF) +
        '"');
    // FIN PT24
    // Annulation de l'intégration des rubriques de réguls des cotisations
    AnnulRegulCot(CodeSalarie, DateD, DateF); // PT180

    CommitTrans;
    // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
    Trace.add('Suppression bulletin du ' + FDate1.text + ' au ' + FDate2.text + ' pour le salarié ' + CodeSalarie);
    CreeJnalEvt('001', '002', 'OK', nil, nil, Trace);
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la suppression du bulletin', 'Saisie du bulletin');
    // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
    Trace.add('Erreur de suppression bulletin du ' + FDate1.text + ' au ' + FDate2.text + ' pour le salarié ' + CodeSalarie);
    CreeJnalEvt('001', '002', 'ERR', nil, nil, Trace);
  end;
  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  Trace.Free;
end;

function TFSaisBul.OnSauve: boolean;
var
  DateS: TdateTime;
begin
  result := False;
  if ActionBul = taConsultation then exit;
  // PT6 : 10/09/2001 V547 PH Controle de la date de la paie en création de bulletin
  if (not ControlPaieCloture(DateD, DateF)) and (ActionBul = taCreation) then
  begin
    Showmessage('Vous ne pouvez pas saisir de paie sur une période close');
    exit;
  end;
  // FIN PT6
  AlimChampEntet(CodeSalarie, Etab, DateD, DateF, TOB_Rub);
  DateS := Tob_sal.GetValeur(iPSA_DATESORTIE);
// d PT115
// result := SauvegardeBul(CodeSalarie, Etab, DateD, DateF, DateS, Tob_Rub, StrToInt(FTrentieme.Text), StrToInt(FTrentDenominateur.Text), TRUE);
  result := SauvegardeBul(CodeSalarie, Etab, DateD, DateF, DateS, Tob_Rub, StrToInt(FTrentieme.Text), StrToInt(FTrentDenominateur.Text), TRUE, TOB_IJSS, TOB_Maintien);
// f PT115

  if Tob_PSD <> nil then
  begin
    // PT213 Mise à jour de la table PRESENCESALARIE
    if VH_Paie.PGMODULEPRESENCE then
    MiseAjourPresence(DateD, DateF, Codesalarie, 'MAJ');

    TOB_PSD.UpdateDB(false, false);
    FreeAndNil(Tob_PSD); // if notVide(Tob_PSD, true) then Tob_PSD.free; { PT167 }
    //Tob_PSD := nil;
  end;
  // PT40 : 19/06/2002 V582 PH gestion de la TOB des lignes saisie des primes
  if Tob_PSP <> nil then
  begin
    TOB_PSP.UpdateDB(false, false);
    FreeAndNil(Tob_PSP); //if notVide(Tob_PSP, true) then Tob_PSP.free; { PT167 }
    //Tob_PSP := nil;
  end;

  //LibereTobCP(Tob_pris) ;// juste pour libérer les tob
  //LibereTobAbs(Tob_ABs) ;// juste pour libérer les tob Modif ph
end;

function TFSaisBul.EnregOK: boolean;
begin
  result := FALSE;
  Modifier := True;
end;

procedure TFSaisBul.EtatInitial;
var
  DD, DF, DateE, DateS: TDateTime;
  St: string;
  Q: TQuery;
  MaTob, T: TOB;
begin
  DD := DateD;
  DF := DateF;
  InitVarTHTRAVAILLES; //PT71-4
  // PT54   27/11/2002 V591 PH Récupération date de sortie et recalcul du trentième si modif salarié dans le bulletin
  ActionSal := FALSE;
  // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
  ActiveCalc := FALSE;
  ImageCalcul.Visible := FALSE;
  // PT28 : 26/03/2002 V571 PH Appel focntion raz structure entete ChpEntete
  RazChptEntete(DateF); // RAZ structure entete
  if Tob_Rub <> nil then
  begin
    Tob_Rub.Free;
    Tob_Rub := nil;
  end; // ph memcheck
  ChbxHoraireSalarie.checked := FALSE;
  ChbxBaseForcee.checked := FALSE;
  ChbxTranchesForcees.checked := FALSE;
   if ActionBul = taCreation then ChbxTrentiem.checked := FALSE; // PT153 et PT188 PT216
  // PT9 : 22/10/2001 V562 PH Gestion des dates edition particulières
  ChpEntete.EdtDu := 0;
  ChpEntete.EdtAu := 0;
  { DEB PT173 }
  CpMois.text := '0';
  CpSupp.text := '0';
  CpAnc.text := '0';
  CBModifAcquis.checked := FALSE;
  { FIN PT173 }
  if TOBAna <> nil then
  begin
    TOBAna.free;
    TOBAna := nil;
  end;
   // DEB PT121
  if TOB_Abs <> nil then
  begin
    TOB_Abs.Free;
    TOB_Abs := nil;
  end;
    // FIN PT121
  if (QMul <> nil) and (ActionBul <> taCreation) then
  begin
    if TypeList = TRUE then // Query complète avec Paieencours
    begin
      DD := QMul.FindField('PPU_DATEDEBUT').AsDateTime;
      DateD := DD;
      DF := QMul.FindField('PPU_DATEFIN').AsDateTime;
      DateF := DF;
      FDate1.Text := DateToStr(DD);
      FDate2.Text := DateToStr(DF);
      BullCompl := QMul.FindField('PPU_BULCOMPL').AsString;
    end
    else // Paie faite mais on navigue sur la liste des paies A Faire donc il faut reprendre
    begin // Pb il faut trouver la paie faite avec les bonnes dates
      st := ' SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_DATEDEBUT >="' + USDateTime(DebutDeMois(DateD)) +
        '" AND PPU_DATEFIN<="' + UsDateTime(FinDeMois(DateF)) + '" AND PPU_SALARIE="' + CodeSalarie +
        '" AND PPU_BULCOMPL <> "X"';
{Flux optimisé
      Q := OpenSql(st, TRUE);
      MaTob := TOB.Create('Ma TOB', nil, -1);
      MaTOB.LoadDetailDB('PAIEENCOURS', '', '', Q, FALSE);
      Ferme(Q);
}
      MaTob := TOB.Create('Ma TOB', nil, -1);
      MaTOB.LoadDetailDBFROMSQL('PAIEENCOURS', st);

      if MaTob.Detail.Count - 1 >= 1 then
      begin
        T := MaTob.Detail[1];
        DD := T.GetValue('PPU_DATEDEBUT').AsDateTime;
        DateD := DD;
        DF := T.GetValue('PPU_DATEFIN').AsDateTime;
        DateF := DF;
        FDate1.Text := DateToStr(DD);
        FDate2.Text := DateToStr(DF);
        BullCompl := '-';
      end;
      MaTob.free;
    end;
  end;
  if ActionBul = taCreation then // Cas de la creation, on regarde les dates entree et sortie
  begin
    // PT11 : 07/11/2001 V562 PH En création bulletin, on se remet en debut et en fin de mois
    if (BullCompl <> 'X') then
    begin // PT122 Si bulletin complémentaire on laisse les dates pasées par defaut
      DateD := DebutDemois(DateD);
      DateF := FinDeMois(DateF);
    end;
    DD := DateD;
    DF := DateF;
    LaDateFin := DateF;
    LaDateDeb := DateD;
    DateE := QMul.FindField('PSA_DATEENTREE').AsDateTime;
    DateS := QMul.FindField('PSA_DATESORTIE').AsDateTime;
    if DateE > DateD then DateD := DateE; // pour avoir les bonnes dates de la paie en fonction entree/sortie
    if (DateS < DateF) and (DateS > 100) and (BullCompl <> 'X') then DateF := Dates; // afin d'avoir un calcul de trentieme exact des le depart
    if (BullCompl = 'X') and (DateF < DateD) then DateF := DateD; // PT147
    FDate1.Text := DateToStr(DateD);
    FDate2.Text := DateToStr(DateF);
    // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    FEDTDU.Text := DateToStr(DateD);
    FEDTAU.Text := DateToStr(DateF);
  end;
  if TobAcSaisieArret <> NIL then FreeAndNil(TobAcSaisieArret); //PT224
  AlimChampFiche(); // Alimentation de la FORME avec les champs
  RazGrilles; // Preparation  des Grilles
  if TOB_Sal = nil then ChargeSalarie; // Recup des infos Salaries PT132 Déjà chargé dans AlimChampFiche
  CalculAFaire := TRUE; // pour forcer le calcul du bulletin la premiere foi et pour remplir les grilles
  //PT15 : 19/12/2001 V571 PH Mémorisation faction du trentième si forcé
  CalAffectTrent;
  IntegreAuto := FALSE; // PT177
  ChargeBulletin; // Chargement du bulletin  et calcul de la paie
  // DEB PT177
  if IntegreAuto then //
  begin
    TOB_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
    CalculBulletin(TOB_Rub);
    IntegreAuto := FALSE;
    RubDansGrille;
  end;
  // FIN PT177
  AlimChampFiche(); // utile car on a des infos de la base
  if ActionBul <> taConsultation then Modifier := TRUE
  else Modifier := FALSE; // consultation pas de calcul donc pas à reecrire
  CalculAFaire := VH_Paie.PGCalculBulletin; // Pour Initialiser le calcul auto en fonction de ParamSoc
  if ActionBul = taCreation then
  begin
    if DateD < DD then DateD := DD; // Recuperation des dates de debut et de fin de mois pour les autres salariés qui ne
    if DateF > DF then DateF := DF; // sont pas sortis ni entrés en cours de période
  end;
  if (VH_Paie.PGAnalytique = TRUE) and (ActionBul <> taConsultation) then
  begin
    if TOBAna <> nil then
    begin
      TOBAna.Free;
      TOBAna := nil;
    end;
    if (ActionBul <> taCreation) then
    begin
      TOBAna := PreChargeVentileLignePaie(CodeSalarie, DateD, DateF);
      ControlAffecteAnal(CodeSalarie, DateD, DateF, TOB_VenRem, TOB_VenCot, TOB_Rub, TOBAna);
    end
    else TOBAna := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, TOB_Rub, CodeSalarie, 'PRE', DateD, DateF);
  end;
  // PT2 : 29/08/2001 V547 PH affichage des cumuls de bas de bulletin
  if ActionBul = taConsultation then // Les cumuls ne sont pas stockés car ils ne sont pas calculées
  begin
    if TOB_CumulSal <> nil then
    begin
      TOB_CumulSal.free;
      TOB_CumulSal := nil;
    end;
    TOB_CumulSal := TOB.Create('Les Cumuls Salariés', nil, -1);
    st := 'SELECT * FROM HISTOCUMSAL WHERE PHC_SALARIE="' + CodeSalarie + '" AND PHC_DATEDEBUT="' + USDateTime(DateD) + '" AND PHC_DATEFIN="' + USDateTime(DateF) +
      '" ORDER BY PHC_CUMULPAIE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    TOB_CumulSal.LoadDetailDB('HISTOCUMSAL', '', '', Q, False); // PT99
    Ferme(Q);
}
    TOB_CumulSal.LoadDetailDBFROMSQL('HISTOCUMSAL', st);

    AfficheBasBulletin;
    // PT99   TOB_CumulSal.free;
    // PT99   TOB_CumulSal := nil;
  end;
  // FIN PT2
  if Pages.ActivePage = PRemSal then // PT152
    GREMSAL.SetFocus; // PT129
  if Pages.ActivePage = PRemSais then
    GREMSAis.SetFocus; // PT202
end;

procedure TFSaisBul.ChargeSalarie(Force: Boolean = FALSE);
var S1, E1: string;
  D1, D2: TDateTime;
  T1, T_Etab: Tob;
  Calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure: string;
begin
  //PT41 : 08/08/2002 V582 PH Rechargement des infos salariés après accès fiche : rajout parametre
  if ((TOB_Sal <> nil) and (CodeSalarie <> TOB_Sal.GetValeur(iPSA_SALARIE))) or (Force = TRUE) then
  begin
    if TOB_Sal <> nil then
    begin
      TOB_Sal.Free;
      TOB_Sal := nil;
    end;
  end;
  // PT32 : 29/04/2002 V582 PH Optimisation requete SQL sur Salarié
  //if TOb_Sal <> NIL Then TOb_sal.Free ;
  //Tob_Sal := NIL;
  if Tob_Sal = nil then
  begin
    //   TOB_Salarie:=TOB.Create('SALARIES',Nil,-1) ;
    //   Tob_Sal.SelectDB('"'+FSalarie.Caption+'"',Nil,TRUE) ;
    RecupTobSalarie(FSalarie.Caption, DateD, DateF); // PT184
    Tob_Sal := TOB_Salarie;
  end;
  ///mise à jour des champs de l'entete
  //Tob_Sal.PutEcran(Self) ;
  //TOB_Salarie := Tob_Sal;
  { DEB PT131 }
  if Assigned(GblTob_Semaine) then FreeAndNil(GblTob_Semaine);
  if Assigned(GblTob_Standard) then FreeAndNil(GblTob_Standard);
  if Assigned(GblTob_JourFerie) then FreeAndNil(GblTob_JourFerie);
  S1 := TOB_Sal.GetValeur(iPSA_SALARIE);
  E1 := TOB_Sal.GetValeur(iPSA_ETABLISSEMENT);
  D1 := TOB_Sal.GetValeur(iPSA_DATEENTREE);
  D2 := TOB_Sal.GetValeur(iPSA_DATESORTIE);

  //DEB PT192
  ETB_Activite := TOB_Sal.GetValeur(iPSA_ACTIVITE);
  //FIN PT192

{$IFNDEF EAGLSERVER}
// PROCESS-SERVEUR
  ChargeTobCalendrier(DateD, DateF, S1, FALSE, True, False, GblTob_Semaine, { PT132 Ajout var rech etab }
    GblTob_Standard, GblTob_JourFerie, T1, E1, Calend, StandCalend,
    EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure, D1, D2);
  { FIN PT131 }
{$ENDIF}
end;

procedure TFSaisBul.OnGrilleExit(Sender: TObject);
begin
  if ACal = TRUE then CalculBullSaisie;
  ACal := FALSE;
end;

procedure TFSaisBul.AfficheActionB;
begin
  if BullCompl <> 'X' then
  begin
    ImageBulCompl.Visible := FALSE;
    Caption := 'Saisie du bulletin';
  end
  else
  begin
    Caption := 'Saisie du bulletin complémentaire';
    ImageBulCompl.Visible := TRUE;
  end;

  Modifier := FALSE;
  case ActionBul of
    PremCreation: FTypeActionBul.Caption := 'CREATION';
    taCreation: FTypeActionBul.Caption := 'CREATION';
    taModification: FTypeActionBul.Caption := 'MODIFICATION';
    taConsultation: FTypeActionBul.Caption := 'CONSULTATION';
  end;
  if ActionBul = PremCreation then ActionBul := taCreation;
  if ActionBul = taConsultation then
  begin
    BNewligne.Enabled := FALSE;
    BDelLigne.Enabled := FALSE;
    BCalculBull.Enabled := FALSE;
    BValider.Enabled := FALSE;
    BDefaire.Enabled := FALSE;
    BDelete.Enabled := FALSE;
    BInsert.Enabled := FALSE;
    BCommentaire.Enabled := FALSE;
    //PT56   06/12/2002 V591 PH Accès aux grilles en consultation dans le cas où la paie est close
    { GRemSal.Enabled := FALSE;
     GRemAbs.Enabled := FALSE;
     GRemAbt.Enabled := FALSE;
     GRemAvt.Enabled := FALSE;
     GRemNonImp.Enabled := FALSE;
     GRemPrimes.Enabled := FALSE;
     GRemRet.Enabled := FALSE;
     GRemCplt.Enabled := FALSE;
     GRemHeures.Enabled := FALSE;
    //PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
     GRemNonImpSoumis.Enabled := FALSE;
     GRemRet.Enabled := FALSE;
     GBas.Enabled := FALSE;
     GCot.Enabled := FALSE;}
    GRemSal.Options := GRemSal.Options - [goEditing];
    GRemSais.Options := GRemSais.Options - [goEditing]; // PT202
    GRemAbs.Options := GRemAbs.Options - [goEditing];
    GRemAbt.Options := GRemAbt.Options - [goEditing];
    GRemAvt.Options := GRemAvt.Options - [goEditing];
    GRemNonImp.Options := GRemNonImp.Options - [goEditing];
    GRemPrimes.Options := GRemPrimes.Options - [goEditing];
    GRemRet.Options := GRemRet.Options - [goEditing];
    GRemCplt.Options := GRemCplt.Options - [goEditing];
    GRemHeures.Options := GRemHeures.Options - [goEditing];
    GRemNonImpSoumis.Options := GRemNonImpSoumis.Options - [goEditing];
    GRemRet.Options := GRemRet.Options - [goEditing];
    GBas.Options := GBas.Options - [goEditing];
    GCot.Options := GCot.Options - [goEditing]; // FIN PT56
  end
  else
  begin
    BNewligne.Enabled := TRUE;
    BDelLigne.Enabled := TRUE;
    BCalculBull.Enabled := TRUE;
    BValider.Enabled := TRUE;
    BDefaire.Enabled := TRUE;
    BDelete.Enabled := TRUE;
    BInsert.Enabled := TRUE;
    BCommentaire.Enabled := TRUE;
    //PT56   06/12/2002 V591 PH Accès aux grilles en consultation dans le cas où la paie est close
    { GRemSal.Enabled := TRUE;
     GRemAbs.Enabled := TRUE;
     GRemAbt.Enabled := TRUE;
     GRemAvt.Enabled := TRUE;
     GRemNonImp.Enabled := TRUE;
     GRemPrimes.Enabled := TRUE;
     GRemRet.Enabled := TRUE;
     GRemCplt.Enabled := TRUE;
     GRemHeures.Enabled := TRUE;
    // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
     GRemNonImpSoumis.Enabled := TRUE;
     GRemRet.Enabled := TRUE;
     GBas.Enabled := TRUE;
     GCot.Enabled := TRUE;}
    GRemSal.Options := GRemSal.Options + [goEditing];
    GRemSais.Options := GRemSais.Options + [goEditing]; // PT202
    GRemAbs.Options := GRemAbs.Options + [goEditing];
    GRemAbt.Options := GRemAbt.Options + [goEditing];
    GRemAvt.Options := GRemAvt.Options + [goEditing];
    GRemNonImp.Options := GRemNonImp.Options + [goEditing];
    GRemPrimes.Options := GRemPrimes.Options + [goEditing];
    GRemRet.Options := GRemRet.Options + [goEditing];
    GRemCplt.Options := GRemCplt.Options + [goEditing];
    GRemHeures.Options := GRemHeures.Options + [goEditing];
    GRemNonImpSoumis.Options := GRemNonImpSoumis.Options + [goEditing];
    GRemRet.Options := GRemRet.Options + [goEditing];
    GBas.Options := GBas.Options + [goEditing];
    GCot.Options := GCot.Options + [goEditing]; // FIN PT56
  end;
  BDelete.Enabled := TRUE;
  if (ActionBul = taCreation) or (ActionBul = taConsultation) then BDelete.Enabled := FALSE;

end;

procedure TFSaisBul.GCotRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  if (ActionLigne = taCreation) and (VH_Paie.PGAnalytique = TRUE) then
    AffectePreVentilAna; // en insertion rajout des pré ventilations analytiques
  if ACal = TRUE then CalculBullSaisie;
  ACal := FALSE;
  Modifier := TRUE;
end;

// Fonction qui indique si la zone suivante doit être saisie
// derniere colonne indiquant la provenance de la rubrique n'est pas accessible

function TFSaisBul.RemZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1, LaLigne: TOB;
  cas: WORD;
  Rubriq, AccesSaisie: string;
begin
  result := FALSE;
  CurG.ElipsisButton := FALSE;
  // Colonne 0 interdite sauf en creation
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := FALSE;
    exit;
  end;
  // colonne 1 libelle tjrs accessible
  if Acol = 1 then
  begin
    result := TRUE;
    exit;
  end;
  if Acol > SG_Mt then
  begin
    result := FALSE;
    exit;
  end; // Derniere Colonne jamais Saisie, elle est activée par une fonction speciale qui precise les conséquences des modifications
  if ActionLigne = taCreation then
  begin
    if ACol = 0 then CurG.ElipsisButton := TRUE
    else CurG.ElipsisButton := FALSE;
  end;
  if 1 = 1 then
  begin
    Rubriq := CurG.Cells[0, CurG.Row];
    if (RechCommentaire(Rubriq) = TRUE) then
    begin
      // PT31 : 24/04/2002 V582 PH seule la zone libelle est accessible pour une rubrique de commentaire
      if ACol > SG_Lib then
      begin
        Result := FALSE;
        exit;
      end;
    end;
    if (Copy(Rubriq, 5, 2) = '.R') then
    begin
      result := TRUE;
      exit;
    end;

    T1 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rubriq], TRUE);
    if T1 <> nil then
    begin
      if not ExamCasValeurRem(ACol, T1) then // Test sur le type du champ dans le cas valeur alors on ne saisit jamais
      begin
        cas := TypeFonctRem(T1); // recup du type de saisie donc renseigne en partie le fonctionnement de la rémunération
        LaLigne := TOB(CurG.Objects[0, ARow]); // Recup de la TOB TOB_Rub associee a chaque ligne
        AccesSaisie := LaLigne.GetValue('PHB_ORIGINELIGNE'); // recup du champ origine de la ligne pour autoriser la saisie
        if cas = 1 then // saisie du montant
        begin
          if (ACol = SG_Mt) or (ACol = SG_Base) then result := TRUE;
        end;
        if cas = 2 then // saisie de la Base et du Taux
        begin
          if (ACol = SG_Base) or (ACol = SG_Taux) then result := TRUE;
        end;
        if cas = 3 then // saisie de la Base du taux et du Coeff
        begin
          if (ACol = SG_Base) or (ACol = SG_Taux) or (ACol = SG_Coeff) then result := TRUE;
        end;
        if cas = 4 then // saisie de la base et du coeff
        begin
          if (ACol = SG_Base) or (ACol = SG_Coeff) then result := TRUE;
        end;
        if result = TRUE then result := AccesLigneBulletin(AccesSaisie, LaLigne.GetValue('PHB_ORIGINEINFO')); // Traitement du cas ou la ligne a été calculée par un module externe
      end; // Cas ExamineCasValeur
    end; // TOb non nulle

  end;
end;
// Fonction qui recherche si la zone est accessible et qui passe à la zone suivante le cas échéant

procedure TFSaisBul.RemZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
begin
  OldEna := CurG.SynEnabled;
  CurG.SynEnabled := False;
  Sens := -1;
  if CurG.Row > ARow then Sens := 1 else if ((CurG.Row = ARow) and (ACol < CurG.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = 0) or (CurG.Col = 0)) and (ARow = 1) and (ActionLigne <> taCreation) then
  begin
    CurG.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := CurG.Col;
  ARow := CurG.Row;
  ii := 0;
  while not RemZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = 0) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = CurG.ColCount - 1) and (ARow = CurG.RowCount - 1)) then
      begin
        ACol := CurG.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < CurG.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := CurG.FixedCols;
      end;
    end else
    begin
      if ((ACol = CurG.FixedCols) and (ARow = 1)) then Break;
      if ACol > CurG.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := CurG.ColCount - 1;
      end;
    end;
  end;
  CurG.SynEnabled := OldEna;
end;
// Fonction de formattage de saisie de la cellule

function TFSaisBul.RemFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
var
  St, StC, Rub, Rubriq: string;
  T1, T2: TOB;
  NbB, NbT, NbC, NbM: Integer;
begin
  T2 := nil;
  St := CurG.Cells[ACol, ARow];
  StC := St;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    cancel := TRUE;
    CurG.Col := ACol;
    CurG.Row := ARow;
    result := T2;
    exit;
  end;
  Rub := CurG.Cells[0, ARow];
  // on lit la TOB des rubrique de rémunération
  Rubriq := RendCodeRub(Rub);
  if (Copy(Rub, 5, 2) = '.R') then Rub := Copy(Rub, 1, 4); // PT176
  T2 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rubriq], TRUE);
  if T2 <> nil then
  begin
    NbB := T2.GetValue('PRM_DECBASE');
    NbT := T2.GetValue('PRM_DECTAUX');
    NbC := T2.GetValue('PRM_DECCOEFF');
    NbM := T2.GetValue('PRM_DECMONTANT');
    if ACol = SG_Lib then {} else
      if ACol = SG_Base then StC := DoubleToCell(Valeur(St), NbB) else
        if ACol = SG_Taux then StC := DoubleToCell(Valeur(St), NbT) else
          if ACol = SG_Coeff then StC := DoubleToCell(Valeur(St), NbC) else
            if ACol = SG_Mt then StC := DoubleToCell(Valeur(St), NbM);
  end;
  CurG.Cells[ACol, ARow] := StC;
  result := T2;
end;
// Fonction de tests en sortie de cellule la grille des Rem

procedure TFSaisBul.GRemSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1, TOB_Rech: TOB;
  Rub, Rubriq, Cc, Lib, TypT: string;
  cas: WORD; // type de fonctionnement de la rubrique de rémunération
  NbM, NbB, NbT, NbC: Integer;
  Comment, RegulRem: Boolean;
  Base, Taux, Coeff, Montant: Double;
begin
  T1 := nil;
  if ActionBul = taConsultation then exit;
  if CurG.Cells[ACol, ARow] = StCellCur then Exit;
  if ACol > 0 then CurG.ElipsisButton := FALSE;
  //cancel:= FALSE;

  Rub := CurG.Cells[0, ARow];
  Comment := RechCommentaire(Rub);
  if (Copy(Rub, 5, 2) = '.R') then RegulRem := TRUE
  else RegulRem := FALSE;

  Rubriq := RendCodeRub(Rub);
  if (Comment = TRUE) and (ACol > 0) then
  begin
    exit;
  end;
  if (ActionLigne <> taCreation) then T1 := RemFormateZoneSaisie(ACol, ARow, Cancel)
  else
  begin
    if ACol = 0 then
    begin
      if Comment = FALSE then CurG.ElipsisButton := TRUE;
      Tob_Rech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Rub], TRUE);
      if (Tob_Rech <> nil) and (ActionLigne = taCreation) and not RegulRem then
      begin
        ShowMessage('La rubrique est déjà utilisée pour ce bulletin');
        bDetruitLMaLigne(ARow);
        Cancel := TRUE;
        exit;
      end;
    end;

    if Comment = FALSE then
    begin
      if RegulRem then Rub := Copy(Rub, 1, 4);
      T1 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], TRUE);
      if T1 = nil then
      begin
        // PT72   02/07/2003 V_421 PH FQ 10737 Message indiquant le rechargement du contexte pour inserer une rubrique
        if ActivClick then ShowMessage('Cliquez sur le bouton [Recharger contexte de la paie] pour pouvoir intégrer la nouvelle rubrique')
        else ShowMessage('La rubrique n''existe pas, la ligne va être supprimée');
        ActivClick := FALSE;
        // FIN PT72
        Cancel := TRUE;
        // PT70   19/06/2003 V_421 PH FQ 10590 Suppression automatique ligne si rubrique inexistante
        //    bDelLigneClick (Sender) ;
        bDetruitLMaLigne(ARow);
        exit;
      end;
      NbM := T1.GetValue('PRM_DECMONTANT');
      NbB := T1.GetValue('PRM_DECBASE');
      NbC := T1.GetValue('PRM_DECCOEFF');
      NBT := T1.GetValue('PRM_DECTAUX');
      GrilleBull := CurG;
      IntegreAuto := FALSE; // PT177
      if not RegulRem then
        EvalueRem(TOB_Rub, Rub, Base, Taux, Coeff, Montant, Lib, DateD, DateF, taCreation, ARow)
      else
      begin // DEB PT212
        if ACol = SG_Base then
        begin
          base := Valeur(CurG.Cells[ACol, ARow]);
          taux := Valeur(CurG.Cells[ACol + 1, ARow]);
          coeff := Valeur(CurG.Cells[ACol + 2, ARow]);
          Montant := Valeur(CurG.Cells[ACol + 3, ARow]);
        end
        else
        begin
          if ACol = SG_Taux then
          begin
            base := Valeur(CurG.Cells[ACol - 1, ARow]);
            taux := Valeur(CurG.Cells[ACol, ARow]);
            coeff := Valeur(CurG.Cells[ACol + 1, ARow]);
            Montant := Valeur(CurG.Cells[ACol + 2, ARow]);
          end
          else
          begin
            if ACol = SG_Coeff then
            begin
              base := Valeur(CurG.Cells[ACol - 2, ARow]);
              taux := Valeur(CurG.Cells[ACol - 1, ARow]);
              coeff := Valeur(CurG.Cells[ACol, ARow]);
              Montant := Valeur(CurG.Cells[ACol + 1, ARow]);
            end
            else
              if ACol = SG_Mt then
              begin
                base := Valeur(CurG.Cells[ACol - 3, ARow]);
                taux := Valeur(CurG.Cells[ACol - 2, ARow]);
                coeff := Valeur(CurG.Cells[ACol - 1, ARow]);
                Montant := Valeur(CurG.Cells[ACol, ARow]);
              end;
          end;
        end; // FIN PT212
      end;
      // PT52   14/11/2002 V591 PH Possibilité de changer (et conserver) le libellé de la ligne en creation de ligne
      if (ActionLigne = taCreation) and (CurG.Cells[SG_LIB, ARow] <> lib) and (CurG.Cells[SG_LIB, ARow] <> '') then lib := CurG.Cells[SG_LIB, ARow];
      RemplissageLigneRem(CurG, Lib, Base, Taux, Coeff, Montant, ARow, NbB, NbT, NbC, NbM);
//      RubDansGrille;
// DEB PT177
      if IntegreAuto then
      begin
        TOB_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
        CalculBulletin(TOB_Rub);
        BCalculBullClick(Sender);
        IntegreAuto := FALSE;
      end;
// FIN PT177
//      AfficheBasBulletin;
      exit;
    end;
  end;
  if RegulRem then Cancel := FALSE;
  rub := '';
  //  if ARow=CurG.RowCount-1 then begin Cancel:=FALSE; exit; end;
  if T1 <> nil then
  begin
    rub := T1.GetValue('PRM_RUBRIQUE');
    Cc := T1.GetValue('PRM_CODECALCUL'); // Recup du code calcul
    NbM := T1.GetValue('PRM_DECMONTANT');
    TypT := T1.GetValue('PRM_TYPETAUX');
    NbT := T1.GetValue('PRM_DECTAUX');
    if TypT = '08' then Taux := Valeur(T1.GetValue('PRM_TAUXREM'));
    cas := TypeFonctRem(T1); // recup du type de saisie donc renseigne en partie le fonctionnement de la rémunération
    if ACol = SG_Rub then TraiteRemCellRub(Rub, Cancel, ActionLigne) else
      if ACol = SG_Base then TraiteRemCellBase(cas, Cc, ACol, ARow, NbM, Cancel, TypT, Taux, NbT) else
        if ACol = SG_Taux then TraiteRemCellTaux(cas, Cc, ACol, ARow, NbM, Cancel) else
          if ACol = SG_Coeff then TraiteRemCellCoeff(cas, Cc, ACol, ARow, NbM, Cancel);
  end;
  if (ACol > SG_Mt) then Cancel := TRUE; // derniere colonne jamais saisie
end;
// Identification de la méthode de fonctiopnement de la rémunération

function TFSaisBul.TypeFonctRem(T1: TOB): WORD;
var
  Cc: string; // Code Calcul de la rémunération
begin
  result := 0;
  Cc := T1.GetValue('PRM_CODECALCUL');
  if Cc <> '' then result := RechCasCodeCalcul(Cc);
end;

{ fonction de controle de saisie du numéro de rubrique - Controle si la rubrique est déjà
utilisée}

procedure TFSaisBul.TraiteRemCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin);
var
  Tob_Rech: TOB;
begin
  if ActionLigne <> taCreation then
  begin
    exit;
  end;
  Tob_Rech := TOB_Rem.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Rub], TRUE);
  if Tob_Rech <> nil then ShowMessage('La rubrique est déjà utilisée pour ce bulletin')
  else ; // cancel:=FALSE;
end;
// Fonction qui calcule le montant en fonction de la base et du taux et du coeff

procedure TFSaisBul.TraiteRemCellBase(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean; TypT: string; TauxR: Double; NbT: Integer);
var
  Base, Taux, Coeff, Mt: Double;
  St: string;
begin
  St := CurG.Cells[ACol, ARow];
  Base := Valeur(St);
  if (Cas = 2) and (TypT = '08') and (TauxR <> 0) then
  begin
    Taux := TauxR;
    CurG.Cells[ACol + 1, ARow] := DoubleToCell(Taux, NbT);
  end
  else
  begin
    St := CurG.Cells[ACol + 1, ARow];
    Taux := Valeur(St);
  end;
  St := CurG.Cells[ACol + 2, ARow];
  Coeff := Valeur(St);
  Mt := ValoriseMt(cas, Cc, NbreDec, Base, Taux, Coeff); // Calcul et arrondi sur le champ montant
  if (Cas = 1) and (Cc = '01') then Mt := Valeur(CurG.Cells[ACol + 3, ARow]); // Base et Montant saisissable mais sans aucun calcul
  CurG.Cells[ACol + 3, ARow] := DoubleToCell(Mt, NbreDec);

end;
// Fonction qui calcule le montant en fonction de la base et du taux et du coeff

procedure TFSaisBul.TraiteRemCellTaux(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean);
var
  Base, Taux, Coeff, Mt: Double;
  St: string;
begin
  Mt := 0;
  St := CurG.Cells[ACol - 1, ARow];
  Base := Valeur(St);
  St := CurG.Cells[ACol, ARow];
  Taux := Valeur(St);
  St := CurG.Cells[ACol + 1, ARow];
  Coeff := Valeur(St);
  if (cas = 2) and (ACol = 3) then
  begin
    if Cc = '04' then Mt := Base * Taux
    else Mt := Base * (Taux / 100);
    CurG.Cells[ACol + 2, ARow] := DoubleToCell(Mt, NbreDec);
  end;
  if (cas = 3) and (ACol = 3) then // le taux est saisi mais il faut aussi saisir le coeff donc on continue
  begin
    Mt := ValoriseMt(cas, Cc, NbreDec, Base, Taux, Coeff);
    CurG.Cells[ACol + 2, ARow] := DoubleToCell(Mt, NbreDec);
  end;
  if (cas = 4) then // Calcul du montant
  begin
    Mt := ValoriseMt(cas, Cc, NbreDec, Base, Taux, Coeff);
    CurG.Cells[ACol + 2, ARow] := DoubleToCell(Mt, NbreDec);
  end;
  if Mt <> 0 then
  begin
    Mt := ARRONDI(Mt, NbreDec); // Calcul et arrondi sur le champ montant
    CurG.Cells[ACol + 2, ARow] := DoubleToCell(Mt, NbreDec);
  end;
end;
// Fonction qui calcule le montant en fonction de la base,du taux, du ceoff

procedure TFSaisBul.TraiteRemCellCoeff(cas: WORD; Cc: string; Acol, ARow, NbreDec: Integer; var Cancel: Boolean);
var
  Base, Taux, Mt, Coeff: Double;
  St: string;
begin
  mt := 0;
  St := CurG.Cells[ACol - 2, ARow];
  Base := Valeur(St);
  St := CurG.Cells[ACol - 1, ARow];
  Taux := Valeur(St);
  St := CurG.Cells[ACol, ARow];
  Coeff := Valeur(St);
  if (cas = 3) and (ACol = 4) then // cas 3 et saisie du coeff
  begin
    if Cc = '02' then Mt := Base * Taux * Coeff;
    if Cc = '03' then Mt := Base * (Taux / 100) * Coeff;
    if Coeff <> 0 then { PT75 }
    begin
      if Cc = '06' then Mt := Base * (Taux / Coeff);
      if Cc = '07' then Mt := Base * (Taux / 100 / Coeff);
    end;
  end;
  // cas 4 et saisie du coeff Cc vaut tjrs 08
  if (cas = 4) and (ACol = 4) then Mt := Base * Coeff;
  if Mt <> 0 then
  begin
    Mt := ARRONDI(Mt, NbreDec); // Calcul et arrondi sur le champ montant
    CurG.Cells[ACol + 1, ARow] := DoubleToCell(Mt, NbreDec);
  end;
  //Cancel:=FALSE;
end;

// Ensemble de fonctions pour la gestion de la grille des cotisations
// Fonction qui indique si la zone suivante doit être saisie

function TFSaisBul.CotZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1: TOB;
  Rubriq: string;
begin
  result := FALSE;
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := FALSE;
    exit;
  end;
  if Acol = 1 then
  begin
    result := TRUE;
    exit;
  end;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    result := FALSE;
    exit;
  end;
  // PT16 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
  Rubriq := CurG.Cells[0, CurG.Row];
  if (RechCommentaire(Rubriq)) then
  begin
    // PT31 : 24/04/2002 V582 PH seule la zone libelle est accessible pour une rubrique de commentaire
    if ACol > SG_Lib then
    begin
      Result := FALSE;
      exit;
    end;
  end;
  if ACol > SG_MtPat then
  begin
    Result := FALSE;
    exit;
  end;
  if (Copy(Rubriq, 5, 2) = '.R') then
  begin
    result := TRUE;
    exit;
  end;
  // on lit la TOB concernant la  ligne concernée
  T1 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [T1.GetValue('PHB_RUBRIQUE'), 'COT'], TRUE);
  if (T1 <> nil) then
  begin
    // PT29 : 04/04/2002 V571 PH Acces saisie des colonnes des cotisations si element variable
    if ACol = SG_LIB then result := TRUE; // Libelle tjrs accessible
    if (ACol = SG_TxSal) and ((T1.GetValue('PCT_TYPETAUXSAL') = 'ELP') or (T1.GetValue('PCT_TYPETAUXSAL') = 'ELV'))
      then result := TRUE;
    if (ACol = SG_TxPat) and ((T1.GetValue('PCT_TYPETAUXPAT') = 'ELP') or (T1.GetValue('PCT_TYPETAUXPAT') = 'ELV'))
      then result := TRUE;
    if (ACol = SG_MtSal) and ((T1.GetValue('PCT_TYPEFFSAL') = 'ELP') or (T1.GetValue('PCT_TYPEFFSAL') = 'ELV'))
      then result := TRUE;
    if (ACol = SG_MtPat) and ((T1.GetValue('PCT_TYPEFFPAT') = 'ELP') or (T1.GetValue('PCT_TYPEFFPAT') = 'ELV'))
      then result := TRUE;
    if (ACol = SG_Base) and ((T1.GetValue('PCT_TYPEBASE') = 'ELP') or (T1.GetValue('PCT_TYPEBASE') = 'ELV'))
      then result := TRUE;
    // FIN PT29
  end;
end;
// Fonction qui recherche si la zone est accessible et qui passe à la zone suivante le cas échéant

procedure TFSaisBul.CotZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
begin
  OldEna := CurG.SynEnabled;
  CurG.SynEnabled := False;
  Sens := -1;
  if CurG.Row > ARow then Sens := 1 else if ((CurG.Row = ARow) and (ACol < CurG.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = 0) or (CurG.Col = 0)) and (ARow = 1) and (ActionLigne <> taCreation) then
  begin
    CurG.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := CurG.Col;
  ARow := CurG.Row;
  ii := 0;
  while not CotZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = 0) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = CurG.ColCount - 1) and (ARow = CurG.RowCount - 1)) then
      begin
        ACol := CurG.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < CurG.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := CurG.FixedCols;
      end;
    end else
    begin
      if ((ACol = CurG.FixedCols) and (ARow = 1)) then Break;
      if ACol > CurG.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := CurG.ColCount - 1;
      end;
    end;
  end;
  CurG.SynEnabled := OldEna;
end;
// Fonction de formattage de saisie de la cellule

function TFSaisBul.CotFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
var
  St, StC, Rub: string;
  T1, T2: TOB;
  NbB, NbTPat, NbTSal, NbMtPat, NbMtSal: Integer;
begin
  T2 := nil;
  St := CurG.Cells[ACol, ARow];
  StC := St;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    cancel := TRUE;
    CurG.Col := ACol;
    CurG.Row := ARow;
    result := T2;
    exit;
  end;
  Rub := CurG.Cells[0, ARow];
  // on lit la TOB des rubrique de cotisation
  if (Copy(Rub, 5, 2) = '.R') then Rub := Copy(Rub, 1, 4);
  T2 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'COT'], TRUE);
  if T2 <> nil then
  begin
    NbB := T2.GetValue('PCT_DECBASE');
    NbTSal := T2.GetValue('PCT_DECTXSAL');
    NbTPat := T2.GetValue('PCT_DECTXPAT');
    NbMtSal := T2.GetValue('PCT_DECMTSAL');
    NbMtPat := T2.GetValue('PCT_DECMTPAT');
    // PT49   05/11/2002 V591 PH Suppression majuscule du libelle saisie dans les cotisations
    if ACol = SG_Lib then StC := Trim(St) else
      if ACol = SG_Base then StC := DoubleToCell(Valeur(St), NbB) else
        if ACol = SG_TxSal then StC := DoubleToCell(Valeur(St), NbTSal) else
          if ACol = SG_MtSal then StC := DoubleToCell(Valeur(St), NbMtSal) else
            if ACol = SG_TxPat then StC := DoubleToCell(Valeur(St), NbTPat) else
              if ACol = SG_MtPat then StC := DoubleToCell(Valeur(St), NbMtPat);
  end;
  CurG.Cells[ACol, ARow] := StC;
  result := T2;
end;
// Fonction de tests en sortie de cellule la grille des Cotisation

procedure TFSaisBul.GCotSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1, Tob_Rech: TOB;
  rub, lib: string;
  NbB, NbTSal, NbTPat, NbMSal, NbMPat: INTEGER;
  Base, TxSal, TxPat, MtSal, MtPat: Double;
  AT, Comment, RegulCot: Boolean;
begin
  Base := 0;
  TxSal := 0;
  TxPat := 0;
  MtSal := 0;
  MtPat := 0;
  rub := '';
  lib := '';
  CurG.ElipsisButton := FALSE;
  if ActionBul = taConsultation then exit;
  if CurG.Cells[ACol, ARow] = StCellCur then exit;
  Cancel := TRUE;
  rub := CurG.Cells[0, ARow];

  // PT16 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
  Comment := RechCommentaire(Rub);
  if (Copy(Rub, 5, 2) = '.R') then RegulCot := TRUE
  else RegulCot := FALSE;
  if (Comment = TRUE) and (ACol > 0) then
  begin
    exit;
  end;
  if (ActionLigne <> taCreation) and (ACol > 0) then T1 := CotFormateZoneSaisie(ACol, ARow, Cancel)
  else // Cas de la Creation d'une ligne de cotisation : Insertion
  begin
    if ACol = 0 then CurG.ElipsisButton := TRUE;
    // PT18 : 27/12/2001 V571 PH  message lors insertion cotisation si rem avec le meme numéro ==> Pb nom du champ nature(rub)
    if RegulCot then Rub := Copy(Rub, 1, 4);
    Tob_Rech := Tob_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [rub, 'COT'], TRUE);
    // @@@@@@@@@@@@@@ Rajouter un indicateur de chgt de pages ==> pageschange et le tester ici
    // @@KK
    if (Tob_Rech <> nil) and not PageChang and (ActionLigne = taCreation) and not RegulCot then
    begin
      ShowMessage('La rubrique est déjà utilisée dans le bulletin');
      PageChang := FALSE;
      bDetruitLMaLigne(ARow);
      exit;
    end;
    T1 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE'], [rub], TRUE);
    if T1 = nil then
    begin
      Cancel := TRUE;
      // PT70   19/06/2003 V_421 PH FQ 10590 Suppression automatique ligne si rubrique inexistante
      //  bDelLigneClick (NIL) ;

      // PT72   02/07/2003 V_421 PH FQ 10737 Message indiquant le rechargement du contexte pour inserer une rubrique
      if ActivClick then ShowMessage('Cliquez sur le bouton [Recharger contexte de la paie] pour pouvoir intégrer la nouvelle rubrique')
      else ShowMessage('La rubrique n''existe pas, la ligne va être supprimée');
      bDetruitLMaLigne(ARow);
      ActivClick := FALSE;
      // FIN PT72
      exit;
    end;
    cancel := FALSE;
    NbB := T1.GetValue('PCT_DECBASE');
    NbTSal := T1.GetValue('PCT_DECTXSAL');
    NbTPAT := T1.GetValue('PCT_DECTXPAT');
    NbMSal := T1.GetValue('PCT_DECMTSAL');
    NbMPat := T1.GetValue('PCT_DECMTPAT');
    IntegreAuto := FALSE; // PT177
    if not RegulCot then EvalueCot(Tob_Rub, rub, Base, TxSal, TxPat, MtSal, MtPat, Lib, DateD, DateF, taCreation, AT);
    // PT52   14/11/2002 V591 PH Possibilité de changer (et conserver) le libellé de la ligne en creation de ligne
    if (ActionLigne = taCreation) and (CurG.Cells[SG_LIB, ARow] <> lib) and (CurG.Cells[SG_LIB, ARow] <> '') then lib := CurG.Cells[SG_LIB, ARow];
    RemplissageLigneCot(CurG, lib, Base, TxSal, TxPat, MtSal, MtPat, ARow, NbB, NbTSal, NbTPat, NbMSal, NbMPat);
// DEB PT177
    if IntegreAuto then
    begin
      TOB_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
      CalculBulletin(TOB_Rub);
      BCalculBullClick(Sender);
      IntegreAuto := FALSE;
    end;
// FIN PT177
    // GrilleDansRub;
//    AfficheBasBulletin;
    ActionLigne := taModification;
    exit;
  end;

  if RegulCot then
  begin
    Cancel := FALSE;
//    exit;
  end;
  rub := '';
  if T1 <> nil then
  begin
    rub := T1.GetValue('PCT_RUBRIQUE');
    if ACol = SG_Lib then Cancel := FALSE; // @TRUE
    if ACol = SG_Rub then TraiteCotCellRub(Rub, Cancel, ActionLigne) else
      if ACol = SG_TxSal then TraiteCotCellTxSal(Rub, Cancel, ACol, ARow, RegulCot) else // PT168
        if ACol = SG_MtSal then TraiteCotCellMtSal(Rub, Cancel, ACol, ARow) else
          if ACol = SG_TxPat then TraiteCotCellTxPat(Rub, Cancel, ACol, ARow, RegulCot) else // PT168
            if ACol = SG_MtPat then TraiteCotCellMtPat(Rub, Cancel, ACol, ARow);
  end;
end;

{ fonction de controle de saisie du numéro de rubrique - Controle si la rubrique est déjà
utilisée}

procedure TFSaisBul.TraiteCotCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin);
var
  Tob_Rech: TOB;
begin
  //if ActionLigne <> taCreation then begin Cancel:=FALSE; exit; end;
  Tob_Rech := TOB_Cotisations.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['COT', Rub], TRUE);
  if Tob_Rech <> nil then ShowMessage('La rubrique de cotisation est déjà utilisée pour ce bulletin')
  else cancel := FALSE;
end;

procedure TFSaisBul.GCotSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1: TOB;
begin
  CurG.ElipsisButton := FALSE; // pas de bouton ellipsis par defaut
  if ActionBul = taConsultation then exit;
  if (ACol = 0) and (ActionCellule = taCreation) then
  begin
    ActionCellule := taModification;
  end;
  if ActionLigne = taCreation then
  begin
    if ACol = 0 then CurG.ElipsisButton := TRUE
    else CurG.ElipsisButton := FALSE;
  end
  else
  begin
    T1 := TOB(CurG.Objects[0, CurG.Row]);
    if T1 = nil then
    begin
      cancel := TRUE;
      CurG.Col := ACol;
      CurG.Row := ARow;
      exit;
    end;
  end;
  CotZoneSuivanteouOk(ACol, ARow, Cancel); // @@@@
  if not Cancel then StCellCur := CurG.Cells[CurG.Col, CurG.Row];
end;

procedure TFSaisBul.GCotDblClick(Sender: TObject);
var
  T, TS, TR: TOB;
  i: Integer;
  Rub, Ret, St: string;
  Q: TQuery;
begin
  Rub := Copy(CurG.Cells[0, Curg.Row], 1, 4);
  T := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE'], [Rub], TRUE);
  if T <> nil then
  begin
{$IFNDEF EAGLSERVER}
    ret := AglLanceFiche('PAY', 'COTISATION', '', T.GetValue('PCT_PREDEFINI') + ';' + T.GetValue('PCT_NODOSSIER') + ';' + T.GetValue('PCT_RUBRIQUE'), '');
{$ENDIF}
  end;
  if Ret <> '' then
  begin
    TS := TOB.create('La Cotisation', nil, -1);
    st := 'SELECT * FROM COTISATION WHERE PCT_NATURERUB="COT" AND ##PCT_PREDEFINI## PCT_RUBRIQUE="' + Rub + '"'; //**//
    q := OpenSql(st, TRUE);
    TS.LoadDetailDB('COTISATION', '', '', q, FALSE, FALSE);
    ferme(q);
    TR := TS.detail[0];
    if TR <> nil then for i := 1 to T.NbChamps do T.Valeurs[i] := TR.Valeurs[i];
    TS.Free;
  end;
  // PT146 rechargement des personnalisations
  initTOB_Execpt(); // Chargement de la tob des execptions
  Align_Execpt('', ''); // Récuperation des personnalisations
end;

// Ensemble de fonctions pour la gestion de la grille des bases de cotisations

procedure TFSaisBul.ZChbxBaseForceeClick(Sender: TObject);
var
  i, j: Integer;
begin
  if CurG = nil then exit; // %%%
  for i := 1 to CurG.RowCount - 1 do
  begin
    for j := 2 to 3 do
    begin
      if BaseZoneGrisee(j, i) then
        GridGriseCell(CurG, j, i, Canvas)
      else
        GridDeGriseCell(CurG, j, i, Canvas);
    end;
  end;
  CurG.Refresh;
end;
// Fonction qui indique si la zone suivante doit être saisie

function TFSaisBul.BaseZoneAccessible(var ACol, ARow: Longint): Boolean;
var
  T1: TOB;
begin
  result := FALSE;
  if Acol <> 0 then CurG.ElipsisButton := FALSE;
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := FALSE;
    exit;
  end;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    result := FALSE;
    exit;
  end;
  // on lit la TOB concernant la  ligne concernée
  // T1 := TOB_Bases.FindFirst (['PCT_RUBRIQUE','PCT_NATURERUB'], [T1.GetValue ('PHB_RUBRIQUE'),'BAS'], TRUE);
  if (T1 <> nil) then
  begin
    if ACol = SG_LIB then result := TRUE; // Libelle tjrs accessible
    if (ACol = SG_Base) and (ChbxBaseForcee.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Plfd) and (ChbxBaseForcee.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Tr1) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Tr2) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Tr3) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Plf1) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Plf2) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
    if (ACol = SG_Plf3) and (ChbxTranchesForcees.Checked = TRUE) then result := TRUE;
  end;
end;
// Fonction qui recherche si la zone est accessible et qui passe à la zone suivante le cas échéant

procedure TFSaisBul.BaseZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var
  Sens, ii: integer;
  OldEna: boolean;
begin
  OldEna := CurG.SynEnabled;
  CurG.SynEnabled := False;
  Sens := -1;
  if CurG.Row > ARow then Sens := 1 else if ((CurG.Row = ARow) and (ACol < CurG.Col)) then Sens := 1;
  if (sens = -1) and ((ACol = 0) or (CurG.Col = 0)) and (ARow = 1) and (ActionLigne <> taCreation) then
  begin
    CurG.SynEnabled := OldEna;
    Cancel := TRUE;
    exit;
  end;
  ACol := CurG.Col;
  ARow := CurG.Row;
  ii := 0;
  while not BaseZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 1000 then Break;
    if (ACol = 0) and (ARow = 1) then
    begin
      ACol := 1;
      break;
    end;
    if Sens = 1 then
    begin
      if ((ACol = CurG.ColCount - 1) and (ARow = CurG.RowCount - 1)) then
      begin
        ACol := CurG.FixedCols;
        ARow := 1;
        Break;
      end;
      if ACol < CurG.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := CurG.FixedCols;
      end;
    end else
    begin
      if ((ACol = CurG.FixedCols) and (ARow = 1)) then Break;
      if ACol > CurG.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := CurG.ColCount - 1;
      end;
    end;
  end;
  CurG.SynEnabled := OldEna;
end;
// Fonction de formattage de saisie de la cellule

function TFSaisBul.BaseFormateZoneSaisie(ACol, ARow: Longint; var Cancel: Boolean): TOB;
var
  St, StC, Rub: string;
  T1, T2: TOB;
  NbB: Integer;
begin
  T2 := nil;
  St := CurG.Cells[ACol, ARow];
  StC := St;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    cancel := TRUE;
    CurG.Col := ACol;
    CurG.Row := ARow;
    result := T2;
    exit;
  end;
  Rub := CurG.Cells[0, ARow];
  // on lit la TOB des rubrique de cotisation
  T2 := TOB_Bases.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'BAS'], TRUE);
  if T2 <> nil then
  begin
    NbB := T2.GetValue('PCT_DECBASECOT');
    if ACol = SG_Lib then {StC:=uppercase(Trim(St))} else
      if ACol <> SG_Rub then StC := DoubleToCell(Valeur(St), NbB);
  end;
  CurG.Cells[ACol, ARow] := StC;
  result := T2;
end;
// Fonction de tests en sortie de cellule la grille des Cotisation

procedure TFSaisBul.GBaseSalCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1: TOB;
  Rub: string;
begin
  if ActionBul = taConsultation then exit;
  if CurG.Cells[ACol, ARow] = StCellCur then exit;
  T1 := BaseFormateZoneSaisie(ACol, ARow, Cancel);
  rub := '';
  cancel := TRUE;
  if T1 <> nil then
  begin
    // PT50   05/11/2002 V591 PH Affichage du libelle de la bse de cotisation en insertion de ligne
    if (ActionLigne = taCreation) and (ACol = SG_Rub) then CurG.Cells[ACol + 1, ARow] := T1.GetValue('PCT_LIBELLE');
    rub := T1.GetValue('PCT_RUBRIQUE');
    IntegreAuto := FALSE; // PT177
    if ACol = SG_Lib then Cancel := TRUE;
    if ACol = SG_Rub then TraiteBaseCellRub(Rub, Cancel, ActionLigne, ARow);
 // DEB PT177
    if IntegreAuto then
    begin
      TOB_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
      CalculBulletin(TOB_Rub);
      BCalculBullClick(Sender);
      IntegreAuto := FALSE;
    end;
// FIN PT177
    if Acol > SG_Rub then Cancel := FALSE;
  end;
end;

{ fonction de controle de saisie du numéro de rubrique - Controle si la rubrique est déjà
utilisée}

procedure TFSaisBul.TraiteBaseCellRub(Rub: string; var Cancel: Boolean; ActionLigne: TActionBulletin; Arow: Integer);
var
  Tob_Rech: TOB;
begin
  if ActionLigne <> taCreation then
  begin
    Cancel := FALSE;
    exit;
  end;
  // PT36 : 14/05/2002 V582 PH FQ 10112 Controle si rubrique de base de cotisation déjà présente
  Tob_Rech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['BAS', Rub], TRUE);
  if Tob_Rech <> nil then
  begin
    ShowMessage('La rubrique de Base de cotisations est déjà utilisée pour ce bulletin');
    bDetruitLMaLigne(ARow);
  end
  else cancel := FALSE;
end;

procedure TFSaisBul.GBaseSalCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  T1: TOB;
begin
  if ActionBul = taConsultation then exit;
  T1 := TOB(CurG.Objects[0, CurG.Row]);
  if T1 = nil then
  begin
    cancel := TRUE;
    CurG.Col := ACol;
    CurG.Row := ARow;
    exit;
  end;
  BaseZoneSuivanteouOk(ACol, ARow, Cancel); // @@@@
  if not Cancel then StCellCur := CurG.Cells[CurG.Col, CurG.Row];
end;

procedure TFSaisBul.GBaseDblClick(Sender: TObject);
var
  T, TS, TR: TOB;
  i: integer;
  Rub, Ret, St: string;
  Q: TQuery;
begin
  Rub := Copy(CurG.Cells[0, CurG.Row], 1, 4);
  T := TOB_Bases.FindFirst(['PCT_RUBRIQUE'], [Rub], TRUE);
  if T <> nil then
  begin
{$IFNDEF EAGLSERVER}
    Ret := AglLanceFiche('PAY', 'COTISATION', '', T.GetValue('PCT_PREDEFINI') + ';' + T.GetValue('PCT_NODOSSIER') + ';' + Rub, '');
{$ENDIF}
  end;
  if Ret <> '' then
  begin
    TS := TOB.create('La BASE', nil, -1);
    st := 'SELECT * FROM COTISATION WHERE PCT_NATURERUB="BAS" AND ##PCT_PREDEFINI## PCT_RUBRIQUE="' + Rub + '"'; //**//
    q := OpenSql(st, TRUE);
    TS.LoadDetailDB('COTISATION', '', '', Q, FALSE, FALSE);
    TR := TS.detail[0];
    if TR <> nil then for i := 1 to T.NbChamps do T.Valeurs[i] := TR.Valeurs[i];
    TS.Free;
    Ferme(Q);
  end;
end;

// Gestion du PopUpMenu

procedure TFSaisBul.FormKeyPress(Sender: TObject; var Key: Char);
begin
  popupmenuglobal.items[7].Enabled := TRUE;
  if (ORD(Key) = VK_F11) and (Key <> 'z') then // PT199
    if popupmenuglobal.items.Count > 0 then
      popupmenuglobal.Popup(Mouse.CursorPos.x - 100, Mouse.CursorPos.x - 100);
end;

procedure TFSaisBul.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  { DEB PT69-2 Raccouris clavier bulletin }
  case Key of
    VK_RETURN: Key := VK_TAB; // PT126
    VK_F1: HelpBtn.Click; //Bouton aide
    VK_F2: if (Shift = [ssShift]) and (PRemHeures.TabVisible) then
        Pages.ActivePage := PRemHeures //Active Onglet Heure
      else if (Shift = []) then Pages.ActivePage := PRemNonImp; //Active Onglet Primes non imposable
    VK_F3: if (Shift = [ssShift]) and (PRemPrimes.TabVisible) then
        Pages.ActivePage := PRemPrimes //Active Onglet Prime
      else if (Shift = []) then Pages.ActivePage := PRemSal; //Active Onglet Salaire
    VK_F4: if (Shift = [ssShift]) and (PRemAbs.TabVisible) then
        Pages.ActivePage := PRemAbs //Active Onglet Absences
      else if (Shift = []) then Pages.ActivePage := PBas; //Active Onglet Base de cotisation
    VK_F5: if (Shift = [ssShift]) and (PRemCplt.TabVisible) then
        Pages.ActivePage := PRemCplt //Active Onglet Complement
      else if (Bchercher.enabled) and (Shift = []) then Bchercher.click; //Bouton rechercher
    VK_F6: if (Shift = [ssShift]) and (PRemAvt.TabVisible) then
        Pages.ActivePage := PRemAvt //Active Onglet Avantage
      else if (Shift = []) then BFerme.click; //Bouton Ferme
    VK_F7: if (Shift = [ssShift]) and (PRemAbt.TabVisible) then
        Pages.ActivePage := PRemAbt //Active Onglet Abbatement
      else if (Shift = []) then Pages.ActivePage := PCot; //Active Onglet cotisation
    VK_F8: if (Shift = [ssShift]) and (PRemPrimeNonImpSoumis.TabVisible) then
        Pages.ActivePage := PRemPrimeNonImpSoumis //Active Onglet Primes non soumis cot
      else if (Shift = []) then Pages.ActivePage := PRemRet; //Active Onglet Retenues
    VK_F9: if (Shift = [ssShift]) and (FDatePaie.enabled) then FDatePaie.SetFocus
      else if (Shift = []) and (BCalculBull.enabled) then BCalculBull.click; //Bouton calculer
    VK_F10: if (Shift = [ssShift]) and (ModeRegle.enabled) then ModeRegle.SetFocus
      else if (Shift = []) and (BValider.enabled) then BValider.click; //Bouton valider
    VK_F11:
      begin //Zoom menu
        popupmenuglobal.items[7].Enabled := TRUE;
        if popupmenuglobal.items.Count > 0 then
          popupmenuglobal.Popup(Mouse.CursorPos.x - 100, Mouse.CursorPos.x - 100);
      end;
    VK_F12: if BtnDetail.enabled then BtnDetail.Click; //Présentation simplifiée/Detail
    VK_INSERT:
      begin
        if (JaileDroitTag(200154)) then //PT197
          if (shift = []) and (BNewLigne.Enabled) then BNewLigne.Click; //Insertion d'une ligne dans la grille
      end;
    VK_DELETE:
      begin
        if (JaileDroitTag(200154)) then //PT197
          if (Shift = [ssCtrl]) and (BDelLigne.Enabled) then BDelLigne.Click; //Suppression d'une ligne
      end;
    VK_END: if (Shift = [ssCtrl, ssshift]) then LanceCalculPaieCp; { PT127-1 }
    Ord('A'): if (Shift = [ssAlt]) and (BVentil.enabled) then BVentil.click //Répartition analytique
      else if (Shift = [ssCtrl]) and (VoirvalorisationAbsence1.enabled) then
        VoirvalorisationAbsence1.click; //Zoom sal valorisation absence congés payés
    Ord('B'): if (Shift = [ssAlt]) and (BRechargContext.enabled) then BRechargContext.Click //Paramétrage de l'origine rubrique
      else if (Shift = [ssCtrl]) and (BasesCotis1.enabled) then
        BasesCotis1.click; //Zoom sal Histo base cotisation
    Ord('C'): if (Shift = [ssAlt]) and (Bcommentaire.enabled) then Bcommentaire.Click; //Ajout Ligne de commentaire
    Ord('D'): if (Shift = [ssAlt]) and (FDate1.enabled) then FDate1.SetFocus //Focus date debut
      else if (Shift = [ssCtrl]) and (ChbxRegltMod.enabled) then
        ChbxRegltMod.Checked := not (ChbxRegltMod.Checked); //Bouton reglement
    Ord('E'): if (Shift = [ssAlt]) then EtablissPaie.Click //Zoom Etablissement
      else if (Shift = [ssCtrl]) and (Etablissement1.enabled) then
        Etablissement1.click; //Zoom sal Etablissement
    Ord('F'): if (Shift = [ssAlt]) then ProfilPgi.Click //Zoom Profil
      else if (Shift = [ssCtrl]) and (FicheInd1.enabled) then
        FicheInd1.click; //Zoom sal Fiche individuelle
    Ord('G'): if (Shift = [ssAlt]) and (FDate2.enabled) then FDate2.SetFocus //Focus date Fin
      else if (Shift = [ssCtrl]) and (CongesPayes1.enabled) then
        CongesPayes1.click; //Zoom sal congés payés
    Ord('H'): if (Shift = [ssAlt]) and (BDelete.enabled) then BDelete.click //Bouton supprimer
      else if (Shift = [ssCtrl]) and (Pages.ActivePage = PRemAbs) and (ChbxHoraireSalarie.enabled) then
      begin
        ChbxHoraireSalarie.checked := not (ChbxHoraireSalarie.checked); //Horaire Personalisé
        ChbxHoraireSalarie.SetFocus;
      end;
    Ord('I'): if (Shift = [ssAlt]) and (BImprimer.enabled) then BImprimer.click //Bouton imprimer
      else if (Shift = [ssCtrl]) and (Voir1.enabled) then Voir1.click; //Zoom sal indemnité congés payés
    Ord('J'): if (Shift = [ssAlt]) and (FEDTDU.enabled) then FEDTDU.SetFocus //Focus date debut edition
      else if (Shift = [ssCtrl]) and (Pages.ActivePage = PBas) and (ChbxBaseForcee.enabled) then
        ChbxBaseForcee.checked := not (ChbxBaseForcee.checked); //Bases Forcées
    Ord('K'): if (Shift = [ssAlt]) then BZoom.Click //Zoom Salariés
      else if (Shift = [ssCtrl]) and (Pages.ActivePage = PBas) and (ChbxTranchesForcees.enabled) then
        ChbxTranchesForcees.checked := not (ChbxTranchesForcees.checked); //Tranches Forcées
    Ord('L'): if (Shift = [ssAlt]) and (bDefaire.enabled) then bDefaire.click //Bouton défaire
      else if (Shift = [ssCtrl]) and (Calendrier.enabled) then
        Calendrier.click; //Zoom sal calendrier
    Ord('M'): if (Shift = [ssAlt]) then Rem.Click //Zoom Remunérations
      else if (Shift = [ssCtrl]) and (ChbxTrentiem.enabled) then //Active Tremtième
      begin
        ChbxTrentiem.checked := not (ChbxTrentiem.checked);
        ChbxTrentiem.SetFocus;
      end;
    Ord('N'): if (Shift = [ssAlt]) then CumSess.Click //Zoom Cumuls
      else if (Shift = [ssCtrl]) and (Absence.enabled) then Absence.click; //Zoom sal Absence
    Ord('P'): if (Shift = [ssAlt]) then PaieEnvers.Click //Zoom Calcul Paie envers
      else if (Shift = [ssCtrl]) and (Voirindemnitcompensatricedecongs1.enabled) then
        Voirindemnitcompensatricedecongs1.click; //Zoom sal indem. compens. CP
    Ord('Q'): if (Shift = [ssAlt]) and (Binsert.enabled) then Binsert.Click //Bouton Nouveau
      else if (Shift = [ssCtrl]) and (Pages.ActivePage = PRemAbt) and (ChbxControlSmic.enabled) then
        ChbxControlSmic.checked := not (ChbxControlSmic.checked); //Contrôle limite SMIC
    Ord('R'): if (Shift = [ssAlt]) and (BtnOrigine.enabled) then BtnOrigine.Click //Paramétrage de l'origine rubrique
      else if (Shift = [ssCtrl]) and (VoirhistoriqueRemunrations1.enabled) then
        VoirhistoriqueRemunrations1.click; //Zoom sal Histo remuneration
    Ord('S'): if (Shift = [ssAlt]) then Sal.Click //Zoom mul salarié
      else if (Shift = [ssCtrl]) and (Salarie1.enabled) then Salarie1.click; //Zoom salarié
    Ord('T'): if (Shift = [ssAlt]) then Tablesdossier.Click //Zoom table dossier
      else if (Shift = [ssCtrl]) and (VoirhistoriquecoTisations1.enabled) then
        VoirhistoriquecoTisations1.click; //Zoom sal histo cotisation
    Ord('U'): if (Shift = [ssAlt]) then ProfilRub.Click //Zoom Profil Rubrique
      else if (Shift = [ssCtrl]) and (CumSessSal1.enabled) then CumSessSal1.click; //Zoom sal cumul
    Ord('V'): if (Shift = [ssAlt]) then Variab.Click; //Zoom Variables
    Ord('Y'): if (Shift = [ssAlt]) and (FEDTAU.enabled) then FEDTAU.SetFocus; //Focus date fin edition
    Ord('Z'): if (Shift = [ssAlt]) and (BZoom.enabled) then BZoom.Click; //Zoom
    // PT83   21/10/2003 V_421 FQ 10642 PH Raccorci saisie arret
    Ord('W'): if (Shift = [ssCtrl]) and (VSaisieArret.enabled) then VSaisieArret.Click; //Zoom
  end;
  { FIN PT69-2 }
end;

procedure TFSaisBul.EtablissPaieClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (JaileDroitTag(41220)) then // PT215
  begin
    AglLanceFiche('PAY', 'ETABLISSEMENT', '', '', 'ACTION=CONSULTATION');
    Charg_TauxAt(); // PT142
  end
  else
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
{$ENDIF}
end;

procedure TFSaisBul.RemClick(Sender: TObject);
var
  TS, TR, T: TOB;
  i: Integer;
  LaDate: TDateTime;
  Q: TQUERY;
  St, Rub: string;
begin
  if (not JaileDroitTag(41120)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
    exit;
  end;

  LaDate := NOW;
  TS := TOB.Create('Les Remunerations modifiees', nil, -1);
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'REMUNERATION_MUL', '', '', '');
{$ENDIF}
  St := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_DATEMODIF >= "' + USDateTime(LaDate) + '" ORDER BY PRM_RUBRIQUE';
  Q := OpenSql(st, TRUE);
  TS.LoadDetailDB('REMUNERATION', '', '', Q, FALSE, False);
  Ferme(Q);
  TR := TS.FindFirst([''], [''], FALSE);
  while TR <> nil do
  begin
    rub := TR.GetValue('PRM_RUBRIQUE');
    T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], TRUE);
    if T <> nil then
    begin
      for i := 1 to T.NbChamps do T.Valeurs[i] := TR.Valeurs[i];
    end;

    TR := TS.FindNext([''], [''], FALSE);
  end;
  FreeAndNil(TS);
  // PT146 rechargement des personnalisations
  initTOB_Execpt(); // Chargement de la tob des execptions
  Align_Execpt('', ''); // Récuperation des personnalisations
end;

procedure TFSaisBul.CotzzClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41130)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
    exit;
  end;

  AglLanceFiche('PAY', 'COTISATION_MUL', '', '', 'ACTION=CONSULTATION');
  // PT146 rechargement des personnalisations + cotisations modifiées
  ChargeCotisations;
  initTOB_Execpt(); // Chargement de la tob des execptions
  Align_Execpt('', ''); // Récuperation des personnalisations
{$ENDIF}
end;

procedure TFSaisBul.SalClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  If JaiLeDroitTag(200001) then  //PT223
    AglLanceFiche('PAY', 'SALARIE_MUL', 'GRILLE=S', '', 'ACTION=CONSULTATION')
  else
  begin
    PGIBox('Vous n''avez pas les droits pour accéder à la liste des salariés');  //PT223
    exit;                                                                        //PT223
  end;
{$ENDIF}
end;

procedure TFSaisBul.VariabClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41140)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
  end
  else AglLanceFiche('PAY', 'VARIABLE_MUL', '', '', 'ACTION=CONSULTATION');
{$ENDIF}
end;

procedure TFSaisBul.GRemSalElipsisClick(Sender: TObject);
var
  sWhere, Them, Titre: string;
begin
  sWhere := '';
  Them := RendThemOnglet;
  if Them <> '' then sWhere := '##PRM_PREDEFINI## PRM_THEMEREM="' + Them + '"' else swhere := '##PRM_PREDEFINI##'; //**//
  Titre := 'Rémunérations de type ' + RechDom('PGTHEMEREM', Them, FALSE);
  if VH_Paie.PgRubAcompte <> '' then
    sWhere := sWhere + ' AND PRM_RUBRIQUE <> "' + VH_Paie.PgRubAcompte + '"';
  //DEB PT192 Ne tenir compte que des rémunérations donc l'activité correspond à celle de l'établissement
  if (ETB_Activite <> '') then
    sWhere := sWhere + ' AND (PRM_ACTIVITE = "' + ETB_Activite + '" OR PRM_ACTIVITE = "")';
  //FIN PT192
{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
  if Them <> '' then LookupList(CurG, Titre, 'REMUNERATION', 'PRM_RUBRIQUE', 'PRM_LIBELLE,PRM_PREDEFINI', sWhere, 'PRM_RUBRIQUE', TRUE, -1)
  else
  begin
    if (ETB_Activite <> '') then //PT192
      sWhere := 'SELECT PEN_RUBRIQUE PRM_RUBRIQUE,PEN_LIBELLE PRM_LIBELLE,PEN_PREDEFINI PRM_PREDEFINI FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## PEN_NATURERUB="AAA"' +
        ' UNION SELECT PRM_RUBRIQUE ,PRM_LIBELLE,PRM_PREDEFINI  FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE <> "7120"' +
        ' AND (PRM_ACTIVITE = "' + ETB_Activite + '" OR PRM_ACTIVITE = "") AND NOT EXISTS (SELECT PEN_RUBRIQUE FROM PGEXCEPTIONS ' +
        ' WHERE ##PEN_PREDEFINI## PEN_RUBRIQUE=PRM_RUBRIQUE AND PEN_NATURERUB= "AAA") ORDER BY PRM_RUBRIQUE,PRM_PREDEFINI'
    else
      sWhere := 'SELECT PEN_RUBRIQUE PRM_RUBRIQUE,PEN_LIBELLE PRM_LIBELLE,PEN_PREDEFINI PRM_PREDEFINI FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## PEN_NATURERUB="AAA"' +
        ' UNION SELECT PRM_RUBRIQUE ,PRM_LIBELLE,PRM_PREDEFINI  FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE <> "7120" AND NOT EXISTS (SELECT PEN_RUBRIQUE FROM PGEXCEPTIONS ' +
        ' WHERE ##PEN_PREDEFINI## PEN_RUBRIQUE=PRM_RUBRIQUE AND PEN_NATURERUB= "AAA") ORDER BY PRM_RUBRIQUE,PRM_PREDEFINI';
{
  sWhere := 'SELECT PRM_RUBRIQUE ,PRM_LIBELLE,PRM_PREDEFINI  FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE <> "7120" AND NOT EXISTS (SELECT PEN_RUBRIQUE FROM PGEXCEPTIONS ' +
      ' WHERE ##PEN_PREDEFINI## PEN_RUBRIQUE=PRM_RUBRIQUE AND PEN_NATURERUB= "AAA") ORDER BY PRM_RUBRIQUE,PRM_PREDEFINI';
}
    LookupList(CurG, Titre, 'REMUNERATION', 'PRM_RUBRIQUE', 'PRM_LIBELLE,PRM_PREDEFINI', '', '', TRUE, -1, swhere);
  end;
{$ELSE}
  LookupList(CurG, Titre, 'REMUNERATION', 'PRM_RUBRIQUE', 'PRM_LIBELLE,PRM_PREDEFINI', sWhere, 'PRM_RUBRIQUE', TRUE, -1);
{$ENDIF}
{$ENDIF}
  // PT72   02/07/2003 V_421 PH FQ 10737 Message indiquant le rechargement du contexte pour inserer une rubrique
  ActivClick := TRUE;
end;

procedure TFSaisBul.GRemSalRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if VH_Paie.PGCalculBulletin then CalculAFaire := TRUE;
end;

procedure TFSaisBul.ActiveElipsis;
begin
  CurG.ElipsisButton := TRUE;
end;

function TFSaisBul.RendThemOnglet: string;
var
  Them: string;
begin
  Them := '';
  // Si Detail alors on filtre sur le theme  Salaire
  if (CurG = GRemSal) and (BtnDetail.Down = TRUE) then Them := 'SAL';
  if CurG = GRemAbs then Them := 'ABS';
  if CurG = GRemAbt then Them := 'ABT';
  if CurG = GRemAvt then Them := 'AVT';
  if CurG = GRemNonImp then Them := 'RNI';
  if CurG = GRemPrimes then Them := 'INI';
  if CurG = GRemRet then Them := 'RSS';
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  if CurG = GRemNonImpSoumis then Them := 'RSC';
  // FIN PT3
  if CurG = GRemCplt then Them := 'COM';
  if CurG = GRemHeures then Them := 'HEU';
  if CurG = GRemRet then Them := 'RSS';
  if CurG = GBas then Them := 'BAS';
  if CurG = GCot then Them := 'COT';
  result := Them;
end;

procedure TFSaisBul.OnGrilleEnter(Sender: TObject);
begin
  // PT43 : 08/08/2002 V582 PH Calcul du bulletin lancé en fonction du paramètre (automatique ou pas)
  // PT51   14/11/2002 V591 PH Calcul du bulletin en changeant d'onglet si une modif faite
  if ((Pages.ActivePage = PCot) or (Pages.ActivePage = PBas)) and (Modifier = TRUE) and (ActionLigne <> taCreation) then
  begin
    CalculAFaire := TRUE;
    CalculBullSaisie;
    CalCulAFaire := VH_Paie.PGCalculBulletin;
    Modifier := FALSE;
  end;
end;

procedure TFSaisBul.RemplissageLigneRem(Grille: THGrid; Lib: string; Base, Taux, Coeff, Montant: Double; ARow: Integer; NbB, NbT, NbC, NbM: Integer);
begin
  Grille.Cells[1, ARow] := Lib;
  Grille.Cells[2, ARow] := DoubleToCell(Base, NbB);
  Grille.Cells[3, ARow] := DoubleToCell(Taux, NbT);
  Grille.Cells[4, ARow] := DoubleToCell(Coeff, NbC);
  Grille.Cells[5, ARow] := DoubleToCell(Montant, NbM);
end;
{ Reajustement du nombre de lignes des gilles en fonction du nombre de ligne
afin d'eviter d'avoir des lignes blanches qui soient inaccessibles }

procedure TFSaisBul.ReajustGrille;
var
  Grille: THGrid;
  i: Integer;
begin
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is THGrid then
    begin
      Grille := THGrid(Components[i]);
      //    if Grille.RowCount > 2 then Grille.RowCount:=Grille.RowCount+10;  // @@@@@@@@
      Grille.SynEnabled := TRUE;
      Grille.Row := 1;
    end;
  end;
end;

{
function TFSaisBul.RemExist(Rub: string): Boolean;
var
  T: TOB;
begin
  T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], TRUE);
  if T = nil then result := FALSE
  else result := TRUE;
end;
}
// PT168

procedure TFSaisBul.TraiteCotCellTxSal(Rub: string; var Cancel: Boolean; ACol, ARow: Integer; RegulCot: Boolean = FALSE);
var
  T1: TOB;
  St: string;
  Base, Mt: Double;
begin
  Cancel := FALSE;
  T1 := Tob_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [rub, 'COT'], TRUE);
  if T1 <> nil then
  begin
    if (T1.GetValue('PCT_TYPETAUXSAL') <> '') or (RegulCot) then // PT168
    begin
      St := CurG.Cells[ACol - 1, ARow];
      Base := Valeur(St);
      Mt := (Base * Valeur(CurG.Cells[ACol, ARow])) / 100;
      CurG.Cells[ACol + 1, ARow] := DoubleToCell(Mt, T1.GetValue('PCT_DECMTSAL'));
    end;
  end;
end;
// PT168

procedure TFSaisBul.TraiteCotCellTxPat(Rub: string; var Cancel: Boolean; ACol, ARow: Integer; RegulCot: Boolean = FALSE);
var
  T1: TOB;
  St: string;
  Base, Mt: Double;
begin
  Cancel := FALSE;
  T1 := Tob_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [rub, 'COT'], TRUE);
  if T1 <> nil then
  begin
    if (T1.GetValue('PCT_TYPETAUXPAT') <> '') or (RegulCot) then // PT168
      St := CurG.Cells[ACol - 3, ARow];
    Base := Valeur(St);
    Mt := (Base * Valeur(CurG.Cells[ACol, ARow])) / 100;
    CurG.Cells[ACol + 1, ARow] := DoubleToCell(Mt, T1.GetValue('PCT_DECMTPAT'));
  end;
end;

procedure TFSaisBul.TraiteCotCellMtSal(Rub: string; var Cancel: Boolean; ACol, ARow: Integer);
begin
  Cancel := FALSE;
end;

procedure TFSaisBul.TraiteCotCellMtPat(Rub: string; var Cancel: Boolean; ACol, ARow: Integer);
begin
  Cancel := FALSE;
end;

procedure TFSaisBul.RemplissageLigneCot(Grille: THGrid; Lib: string; Base, TxSal, TxPat, MtSal, MtPat: Double; ARow: Integer; NbB, NbTxSal, NbTxPat, NbMtSal, NbMtPat: Integer);
begin
  Grille.Cells[1, ARow] := Lib;
  Grille.Cells[2, ARow] := DoubleToCell(Base, NbB);
  Grille.Cells[3, ARow] := DoubleToCell(TxSal, NbTxSal);
  Grille.Cells[5, ARow] := DoubleToCell(TxPat, NbTxPat);
  Grille.Cells[4, ARow] := DoubleToCell(MtSal, NbMtSal);
  Grille.Cells[6, ARow] := DoubleToCell(MtPat, NbMtPat);
end;

procedure TFSaisBul.GCotElipsisClick(Sender: TObject);
var
  sWhere, Them, Titre: string;
begin
  sWhere := '';
  if CurG = GCot then
  begin
    Them := 'COT';
    Titre := 'Cotisations';
  end
  else
  begin
    Them := 'BAS';
    Titre := 'Bases de Cotisations';
  end;
  sWhere := '##PCT_PREDEFINI## PCT_NATURERUB="' + Them + '"';
  //DEB PT192 Ne tenir compte que des rémunérations donc l'activité correspond à celle de l'établissement
  if (ETB_Activite <> '') then
    sWhere := sWhere + ' AND (PCT_ACTIVITE = "' + ETB_Activite + '" OR PCT_ACTIVITE = "")';
  //FIN PT192
{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
  if Them = 'BAS' then LookupList(CurG, Titre, 'COTISATION', 'PCT_RUBRIQUE', 'PCT_LIBELLE,PCT_PREDEFINI', sWhere, 'PCT_RUBRIQUE', TRUE, -1)
  else
  begin
    if (ETB_Activite <> '') then //PT192
      sWhere := 'SELECT PEN_RUBRIQUE PCT_RUBRIQUE,PEN_LIBELLE PCT_LIBELLE,PEN_PREDEFINI PCT_PREDEFINI FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## PEN_NATURERUB="' + Them + '"' +
        ' UNION SELECT PCT_RUBRIQUE ,PCT_LIBELLE,PCT_PREDEFINI  FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="COT" ' +
        ' AND (PCT_ACTIVITE = "' + ETB_Activite + '" OR PCT_ACTIVITE = "") AND NOT EXISTS (SELECT PEN_RUBRIQUE FROM PGEXCEPTIONS ' +
        ' WHERE ##PEN_PREDEFINI## PEN_RUBRIQUE=PCT_RUBRIQUE AND PEN_NATURERUB= "' + Them + '") ORDER BY PCT_RUBRIQUE,PCT_PREDEFINI'
    else
      sWhere := 'SELECT PEN_RUBRIQUE PCT_RUBRIQUE,PEN_LIBELLE PCT_LIBELLE,PEN_PREDEFINI PCT_PREDEFINI FROM PGEXCEPTIONS WHERE ##PEN_PREDEFINI## PEN_NATURERUB="' + Them + '"' +
        ' UNION SELECT PCT_RUBRIQUE ,PCT_LIBELLE,PCT_PREDEFINI  FROM COTISATION WHERE ##PCT_PREDEFINI## PCT_NATURERUB="COT" AND NOT EXISTS (SELECT PEN_RUBRIQUE FROM PGEXCEPTIONS ' +
        ' WHERE ##PEN_PREDEFINI## PEN_RUBRIQUE=PCT_RUBRIQUE AND PEN_NATURERUB= "' + Them + '") ORDER BY PCT_RUBRIQUE,PCT_PREDEFINI';
    LookupList(CurG, Titre, 'COTISATION', 'PCT_RUBRIQUE', 'PCT_LIBELLE,PCT_PREDEFINI', '', '', TRUE, -1, swhere);
  end;
{$ELSE}
  LookupList(CurG, Titre, 'COTISATION', 'PCT_RUBRIQUE', 'PCT_LIBELLE,PCT_PREDEFINI', sWhere, 'PCT_RUBRIQUE', TRUE, -1);
{$ENDIF}
{$ENDIF}
  // PT72   02/07/2003 V_421 PH FQ 10737 Message indiquant le rechargement du contexte pour inserer une rubrique
  ActivClick := TRUE;
end;
{Fonction qui se charge de copier les grilles dans la liste des rubriques,
de calculer les lignes de bulletin, de remplir les grilles en fonction  des
calculs fait par le moteur, d'afficher le bas de bulletin.
Le Calcul est soit automatique, soit déclencher par un click utilisateur
}

procedure TFSaisBul.CalculBullSaisie;
//var SystemTime0,SystemTime1,SystemTime2,SystemTime3,SystemTime4,SystemTime5,SystemTime6,SystemTime7 : TSystemTime;
var
  Q: TQuery;
  NbEltManquant: Integer;
begin
  if ActionBul = taConsultation then
  begin
    RubDansGrille;
    exit;
  end; // Pas de calcul en consultation
  if CalculAFaire = FALSE then exit; // Calcul non Auto donc rejet

  //GetLocalTime(SystemTime0);
  SourisSablier;
  //GrilleDisable;
  if ActionBul = PremCreation then RubDansGrille;
  GrilleDansRub;
  //GetLocalTime(SystemTime1);
  AlimChampEntete;
  //GetLocalTime(SystemTime2);
  CalculBulletin(TOB_Rub);
  //GetLocalTime(SystemTime3);
  if (ActionBul = taCreation) or (ForceAna) then // en modif , on relit les ventilations analytiques du bulletin
  begin
    if (VH_Paie.PGAnalytique = TRUE) and (TobAna = nil) then // on ne cree les pre ventilations qu'à la 1ere creation
    begin
      if TobAna <> nil then
      begin
        TobAna.free;
        TObAna := nil;
      end;
      TobAna := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, TOB_Rub, CodeSalarie, 'PRE', DateD, DateF);
      //      GetLocalTime(SystemTime4);
    end;
  end;
  if (ActionBul <> taConsultation) and (VH_Paie.PGAnalytique) then PGReaffSectionAnal(TOB_Rub, TOBAna);
  // Reaffectation des ventilations analytiques des cotisations en focntion de la valorisation des remunerations
  // @@@@ Analytique
  if (VH_Paie.PGAnalytique = TRUE) then
  begin
    if TOBAna = nil then TOBAna := TOB.Create('Analytique bulletin Salarié', nil, -1);
    ControlAffecteAnal(CodeSalarie, DateD, DateF, TOB_VenRem, TOB_VenCot, TOB_Rub, TOBAna);
    //  GetLocalTime(SystemTime5);
  end;
  RubDansGrille;
  //GetLocalTime(SystemTime6);
  AfficheBasBulletin;
  //GetLocalTime(SystemTime7);
  GrilleEnable;
  SourisNormale;
  // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
  ActiveCalc := FALSE;
  ImageCalcul.Visible := FALSE;

  //DEB PT181
{$IFNDEF EAGLSERVER}
  Q := OpenSql('SELECT COUNT(*) AS COMPTEUR FROM PGSYNELTNAT WHERE (PEY_TYPUTI = "1" OR PEY_TYPUTI = "3")AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"', True);
  if not Q.eof then
    NbEltManquant := Q.FindField('COMPTEUR').AsInteger;
  Ferme(Q);
  if NbEltManquant <> 0 then
  begin
    if PgiAsk(TraduireMemoire('Des éléments ont été détectés manquants. Voulez-vous les saisir maintenant ?'), TraduireMemoire('Eléments nationaux manquants')) = mrYes then
    begin
      AglLanceFiche('PAY', 'SAISIESYNELTNAT', '', '', ClePGSynEltNAt);
      if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
      begin
        Nettoyage_EltDynSal();
        Nettoyage_EltDynPop();
        Nettoyage_EltDynEtab();
        Nettoyage_EltNationauxSTD();
        Nettoyage_EltNationauxCEG();
        Nettoyage_EltNationauxDOS();
        initTOB_EltDynSal(CodeSalarie, DateF);
        initTOB_EltDynEtab(Etab, DateF);
        initTOB_EltDynPop(CodeSalarie, DateF);
        initTOB_EltNatDOS();
        initTOB_EltNatSTD();
        initTOB_EltNatCEG();
      end;
      CalculBullSaisie;
      exit;
    end;
    ExecuteSQL('DELETE FROM PGSYNELTNAT WHERE (PEY_TYPUTI = "1" OR PEY_TYPUTI = "3") AND PEY_NODOSSIER = "' + ClePGSynEltNAt + '"');
  end;
{$ENDIF}
  //FIN PT181
end;

procedure TFSaisBul.ProfilPgiClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41110)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
  end
  else
    AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
{$ENDIF}
end;

procedure TFSaisBul.ProfilRubClick(Sender: TObject);
var
  stq, Nature, Titre, Rub: string;
  VV: Variant; // PT144
begin
  Rub := CurG.Cells[0, CurG.Row];
  VV := CurG.Cells[Curg.Col, CurG.Row]; // PT144
  if Rub = '' then
  begin
    ShowMessage('Vous devez sélectionner une ligne du bulletin');
    exit;
  end;
  Titre := 'Liste des Profils utilisant la Rubrique ' + Rub;
  if CurG = GCot then Nature := 'COT'
  else
  begin
    if CurG = GBas then Nature := 'BAS'
    else Nature := 'AAA';
  end;
  //DEBUT PT117
  stq := '##PPI_PREDEFINI## PPM_NATURERUB="';
  stq := stq + Nature + '" AND ##PPM_PREDEFINI## PPM_RUBRIQUE="' + Rub + '"';
{$IFNDEF EAGLSERVER}
{$IFDEF EAGLCLIENT}
  CurG.Cells[Curg.Col, CurG.Row] := ''; // PT145
{$ENDIF}
  LookupList(CurG, Titre, 'PROFILRUB LEFT JOIN PROFILPAIE ON PPI_PROFIL=PPM_PROFIL ', 'PPI_PROFIL', 'PPI_LIBELLE', Stq, '', TRUE, 0);
{$ENDIF}
  CurG.Cells[Curg.Col, CurG.Row] := VV; //PT144
  //FIN PT117
end;

{PT30-2
procedure TFSaisBul.BtnCalendrierClick(Sender: TObject);
var Annee,Mois,St : String;
    T1 : TOB;
begin
st:='';
RendMoisAnnee (DateF, Mois, Annee);
T1:=RendCalendrierMois (Etab, Annee, Mois, TOB_JourTrav );
if T1<>NIL then
 begin
 if etab <> '' then st:=Annee+';'+Mois+';'+Etab else st:=Annee+';'+Mois;
 end;
AglLanceFiche ('PAY','JOURTRAV', '',st , '' );
end;}

procedure TFSaisBul.ChbxHoraireSalarieClick(Sender: TObject);
var
  T_Etab: tob;
  Val1, Val2: Double;
begin
  if ChbxHoraireSalarie.Checked = TRUE then
  begin
    EdtHorReel.Enabled := TRUE;
    EdtJoursOuvres.Enabled := TRUE;
    EdtJoursOuvrables.Enabled := TRUE;
  end
  else
  begin
    EdtHorReel.Enabled := FALSE;
    EdtJoursOuvres.Enabled := FALSE;
    EdtJoursOuvrables.Enabled := FALSE;
    //DEB PT30-2 En décoche horaire personnalisé réaffectation des valeurs théoriques
    EdtHorReel.Text := DoubleToCell(HTHTRAVAILLES, 2);
    T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
    if T_Etab = nil then exit;
    { DEB PT47-4 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables
      Contrairement aux jours les heures ne sont pas sauvegardés dans la fiche ni dans la base..
      EdtJoursOuvres.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,TRUE),2);
      EdtJoursOuvrables.Text:=DoubleToCell (NombreJoursOuvresOuvrablesMois(T_Etab,DateD,DateF,FALSE),2);}
    CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, TRUE, Val1, Val2);
    EdtJoursOuvres.Text := DoubleToCell(Val1, 2);
    CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DateD, DateF, FALSE, Val1, Val2);
    EdtJoursOuvrables.Text := DoubleToCell(Val1, 2);
    //FIN PT30-2 & PT47-4
  end;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.EdtHorReelExit(Sender: TObject);
begin
  EdtHorReel.Text := DoubleToCell(Valeur(EdtHorReel.Text), 2);
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.EdtJoursOuvresExit(Sender: TObject);
var
  St: string;
begin
  EdtJoursOuvres.Text := DoubleToCell(Valeur(EdtJoursOuvres.Text), 2);
  GrilleDansRub;
  AlimChampEntete;
  if ((VH_paie.PGCongesPayes) and (GblCP)) then { PT127-4 }
  begin
    if ActionBul <> taCreation then
      AnnuleCongesPris(Tob_sal.GetValeur(iPSA_SALARIE), Tob_sal.GetValeur(iPSA_ETABLISSEMENT), DateD, DateF);
    Tob_Pris := SalIntegreCP(TOB_Sal, TOB_Rub, T_MvtAcquis, DateD, DateF, False, St); //PT12 passage en parametre de la tob des acqs en cours
    RubDansGrille;
    if VH_Paie.PGCalculBulletin then
      BCalculBullClick(nil);
  end;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.EdtJoursOuvrablesExit(Sender: TObject);
var
  St: string;
begin
  EdtJoursOuvrables.Text := DoubleToCell(Valeur(EdtJoursOuvrables.Text), 2);
  GrilleDansRub;
  AlimChampEntete;
  if ((VH_paie.PGCongesPayes) and (GblCP)) then { PT127-4 }
  begin
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    if BullCompl <> 'X' then
    begin
      if ActionBul <> taCreation then
        AnnuleCongesPris(Tob_sal.GetValeur(iPSA_SALARIE), Tob_sal.GetValeur(iPSA_ETABLISSEMENT), DateD, DateF);
      Tob_Pris := SalIntegreCP(TOB_Sal, TOB_Rub, T_MvtAcquis, DateD, DateF, False, St); //PT12 passage en parametre de la tob des acqs en cours
    end;
    RubDansGrille;
    if VH_Paie.PGCalculBulletin then
      BCalculBullClick(nil);
  end;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

// fonction de RAZ des grilles

procedure TFSaisBul.RazGrilles;
begin
  GRemSal.SynEnabled := FALSE;
  GRemSais.SynEnabled := FALSE; // PT202
  GBas.SynEnabled := FALSE;
  GCot.SynEnabled := FALSE;
  GRemHeures.SynEnabled := FALSE;
  GRemNonImp.SynEnabled := FALSE;
  GRemPrimes.SynEnabled := FALSE;
  GRemAbs.SynEnabled := FALSE;
  GRemCplt.SynEnabled := FALSE;
  GRemAbt.SynEnabled := FALSE;
  GRemRet.SynEnabled := FALSE;
  GRemSal.SynEnabled := FALSE;
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  GRemNonImpSoumis.SynEnabled := FALSE;
  PaieRazPile(GRemNonImpSoumis);
  PaieRazPile(GRemSal);
  PaieRazPile(GRemSais); // PT202
  PaieRazPile(GBas);
  PaieRazPile(GCot);
  PaieRazPile(GRemHeures);
  PaieRazPile(GRemNonImp);
  PaieRazPile(GRemPrimes);
  PaieRazPile(GRemRet);
  PaieRazPile(GRemAbs);
  PaieRazPile(GRemAbt);
  PaieRazPile(GRemCplt);
  PaieRazPile(GRemAvt);
  GCot.Row := 1;
  GBas.Row := 1;
  GRemSal.Row := 1;
  GRemSais.Row := 1; // PT202
  GRemHeures.Row := 1;
  GRemPrimes.Row := 1;
  GRemAbs.Row := 1;
  GRemCplt.Row := 1;
  GRemAvt.Row := 1;
  GRemAbt.Row := 1;
  GRemRet.Row := 1;
  GRemNonImp.Row := 1;
  GRemNonImpSoumis.Row := 1;
end;

procedure TFSaisBul.BtnAligProfPerClick(Sender: TObject);
var
  ListeProfil: string;
  rep: Integer;
begin
  if MulProfilTempo.Text = '' then exit;
  Rep := MSG.execute(3, Caption, '');
  if rep = mrNo then exit;
  ChargeSalarie;
  ListePRofil := MulProfilTempo.Text;
  AlimProfilMulti(ListePRofil, TOB_Sal, TOB_Rub);
  RubDansGrille;
  CalculBullSaisie;
end;

procedure TFSaisBul.FTrentiemeExit(Sender: TObject);
var
  ZDate: tdatetime;
begin
  // DEB PT114
  if (BullCompl = 'X') then
  begin
    CalAffectTrent;
    Modifier := TRUE;
  end;
  // FIN PT114
    { DEB PT71-2 }
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl <> 'X') and (Trentieme_entree <> FTrentieme.Text) then
  begin
    CalAffectTrent;
    // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
    Modifier := TRUE;
    ZDate := TOB_Sal.GetValeur(iPSA_DATESORTIE);
    if (((ZDate > Datef) or (Zdate < DateD)) and (ActionBul = tamodification)) or (ActionBul = TaCreation) then
    begin //PT79-2 Réinitialisation des variables car acquis modifié
      if not ChpEntete.CpAcquisMod then { PT87 }
      begin
        InitialiseVariableStatCP;
        RecalculCp;
      end;
    end
    else
      if (FicheLoad) and (ZDate <= Datef) and (Zdate >= DateD) and (ActionBul = tamodification) then
        PGiBox('Le mouvement de solde a déjà été généré.#13#10' +
          'Pour le modifier, veuillez supprimer et recréer votre bulletin.', Caption);
  end;
  { FIN PT71-2 }
end;

// Fonction qui calcule et affecte le trentieme en fonction de la saisie

procedure TFSaisBul.CalAffectTrent;
var
  Numer, Denomin: Integer;
begin
  Numer := StrToInt(FTrentieme.Text);
  Denomin := StrToInt(FTrentDenominateur.Text);
  if Denomin = 0 then Denomin := 30;
  MemoriseTrentieme(Numer / Denomin); // Memorisation fraction trentieme
  ChpEntete.NTrent := Numer; // Memorisation Numerateur et denominateur trentieme dans la structure ChpEntente
  ChpEntete.DTrent := Denomin;
end;

procedure TFSaisBul.RecalculCp(ZoomCp: Boolean = False; BCalcul: Boolean = False);
var
//PT210  nbmois, base, Mois, Supp, Anc: double;
  nbmois, base, Mois, Supp, Anc, Fract: double;
  Tob_TestPris: TOB;
  T: tob;
  StMsgErr: string;
  New, ModifAcq, RecupCongePri, SupprConge: boolean;
begin
  // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  // PT19 : 02/01/2001 V571 SB Recalcul cp non fait quand on modifie les trentièmes - correction PT9
  if BullCompl = 'X' then exit;
  T := Tob_etablissement.findfirst(['ETB_ETABLISSEMENT'], [Tob_salarie.GetValeur(iPSA_ETABLISSEMENT)], false);
  if T = nil then exit;
  GblCP := (VH_Paie.PGCongesPayes) and (T.GetValue('ETB_CONGESPAYES') = 'X') and (TOB_SALARIE.GetValeur(IPSA_CONGESPAYES) = 'X'); { PT89 } { PT127-4 }
  if not GblCP then exit; { PT127-4 }
  Mois := 0;
  Supp := 0;
  Anc := 0;
  { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
  { JCPACQUIS:=AlimCongesAcquis(TOB_Salarie,T,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,false,base,nbmois,suivant,Mois,Supp,ANc); // Génération cp acquis }
//PT210    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, True, base, Nbmois, Mois, Supp, Anc);
    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, True, base, Nbmois, Mois, Supp, Anc, Fract);
//PT210  ChargeT_MvtAcquis(Tob_salarie.GetValeur(iPSA_SALARIE), DateD, DateF, Mois, Supp, Anc, ModifAcq, New); //PT21
  ChargeT_MvtAcquis(Tob_salarie.GetValeur(iPSA_SALARIE), DateD, DateF, Mois, Supp, Anc, Fract, ModifAcq, New); //PT21
  //DEB PT38-2 Lors de la création d'un second bulletin, on réinitialise les CP
  if (New) then {and ((ActionBul = TaCreation) or (ChpEntete.CpAcquisMod = True))) then PT87 13/05/2004 }
  begin
    ModifAcq := True;
    LibereTobCP(Tob_Pris);
    { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
    { JCPACQUIS:=AlimCongesAcquis(TOB_Salarie,T,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,True,base,nbmois,suivant,Mois,Supp,Anc); }
//PT210    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, True, base, Nbmois, Mois, Supp, Anc);
    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, True, base, Nbmois, Mois, Supp, Anc, Fract);
    InitTobAcquisCp(TOB_Salarie, T, DateD, DateF, base, Nbmois, Mois, Supp, Anc, suivant);
  end;
  //FIN PT38-2
  if X <> nil then
  begin
    X.CpMois.text := floattostr(Mois);
    X.CpSupp.text := floattostr(Supp);
    X.CpAnc.text := floattostr(Anc);
  end;
  MCPACQUIS := nbmois;
  BCPACQUIS := base;
  //DEB PT21
  Tob_TestPris := Tob.create('CONGES PRIS', nil, -1);
  {DEB PT47-2 Pour une suppr de CP Pri on réinitialise les variables globales
  et réintegre les CP pour recalcul}
  RecupCongePri := RecupereCongespris(Tob_TestPris, Tob_salarie, DateF);
  if ((RecupCongePri = False) and (JCPPris > 0) and (ZoomCp)) or (BCalcul) then { PT92-1 }
  begin
    HCPPRIS := 0;
    HCPPAYES := 0;
    JCPPRIS := 0;
    JCPPAYES := 0;
    //JCPSOLDE := 0;
    CPMTABS := 0;
    CPMTINDABS := 0;
    SupprConge := TRUE;
  end
  else SupprConge := False;

  //FIN PT47-2
  if (RecupCongePri = True) or (ModifAcq = True) or (SupprConge) or (BCalcul) then //PT47-2 ajout clause { PT92-1 }
  begin
    AnnuleCongesPris(Tob_salarie.GetValeur(iPSA_SALARIE), Tob_salarie.GetValeur(iPSA_ETABLISSEMENT), DateD, DateF);
    Tob_Pris := SalIntegreCP(Tob_salarie, TOB_Rub, T_MvtAcquis, DateD, DateF, False, StMsgErr);
    RubDansGrille;
    //  CalculBullSaisie;
    if not BCalcul then BCalculBull.Click; //PT47 Un seul calcul bulletin ne suffit pas pour réactualiser la grille { PT92-1 }
    Modifier := TRUE;
  end;
  if Tob_TestPris <> nil then
    FreeAndNil(Tob_TestPris);
  //FIN PT21
end;

procedure TFSaisBul.FTrentDenominateurExit(Sender: TObject);
var
  ZDate: tdatetime;
begin
  { DEB PT71-2 }
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl <> 'X') and (Denominateur_entree <> FTrentDenominateur.Text) then
  begin
    CalAffectTrent;
    // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
    Modifier := TRUE;
    ZDate := TOB_Sal.GetValeur(iPSA_DATESORTIE);
    if (((ZDate > Datef) or (Zdate < DateD)) and (ActionBul = tamodification)) or (ActionBul = TaCreation) then
    begin //PT79-2 Réiniitialisation des variables car acquis modifié
      if not ChpEntete.CpAcquisMod then { PT87 }
      begin
        InitialiseVariableStatCP;
        RecalculCp;
      end;
    end
    else
      if (FicheLoad) and (ZDate <= Datef) and (Zdate >= DateD) and (ActionBul = tamodification) then
        PGiBox('Le mouvement de solde a déjà été généré.#13#10' +
          'Pour le modifier, veuillez supprimer et recréer votre bulletin.', Caption);
  end;
  { FIN PT71-2 }
end;

procedure TFSaisBul.ChbxTrentiemeClick(Sender: TObject);
var
  T_Etab: Tob;
begin
  { DEB PT84 Verrouillage calcul trentième si profil CP non affecté et calcul solde }
  if (VH_paie.PGCongesPayes) then
    if (TOB_Salarie.GetValeur(iPSA_DATESORTIE) <= Datef) and (TOB_Salarie.GetValeur(iPSA_DATESORTIE) >= DateD) then
      if TOB_Etablissement <> nil then
      begin
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
        if T_Etab <> nil then
          if (T_Etab.GetValue('ETB_CONGESPAYES') = 'X') and (TOB_Salarie.GetValeur(iPSA_CONGESPAYES) = 'X') { PT89 }
            and (PGRecupereProfilCge(TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)) = '') then
          begin
            PgiBox('Vous devez gérer un profil congés payés pour forcer le trentième.', 'Mouvement de solde non généré.');
            Exit;
          end;
      end;
  { FIN PT84 }
  if ChbxTrentiem.Checked = TRUE then
  begin
    FTrentieme.Enabled := TRUE;
    FTrentDenominateur.Enabled := TRUE;
  end
  else
  begin
    FTrentieme.Enabled := FALSE;
    FTrentDenominateur.Enabled := FALSE;
    LanceCalculTrentieme;
    //PT79-2 Si décoche test trentième modifié
    if (Denominateur_Entree <> FTrentDenominateur.text) or (Trentieme_entree <> FTrentieme.text) then
      if not ChpEntete.CpAcquisMod then { PT87 }
      begin
        InitialiseVariableStatCP; //PT79-2 Réiniitialisation des variables car acquis modifié
        RecalculCp; //PT47-3 On recalcul les CP lorsqu'on décoche le trentième forcé
      end;
  end;
  if VH_Paie.PGCalculBulletin then CalculAFaire := TRUE; // Pour forcer les calculs suite à modif avent la validation du bulletin
end;

procedure TFSaisBul.AlimChampEntete;
begin
  with ChpEntete do
  begin
    // PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
    if TOB_Sal <> nil then
    begin
      if not ChbxRegltMod.checked then reglt := RendModeRegle(TOB_Sal)
      else reglt := ModeRegle.Value;
    end;
    // PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
    if (Reglt <> '') and (ModeRegle.Value = '') then ModeRegle.Value := Reglt;
    //DEB PT33 Modification de la fonctionnalité
    if IsValidDate(FDatePaie.Text) then Datepai := StrToDate(FDatePaie.Text)
    else
    begin
      PGIBox('La date de paiement du bulletin n''est pas renseignée.#13#10 Donc la date de fin de bulletin sera prise par défaut', Caption);
      //    DatePai := idate1900;
      FDatePaie.Text := DateToStr(DatePai);
    end; //FIN PT33 Modification de la fonctionnalité
    DateVal := NOW;
    if ChbxHoraireSalarie.checked = TRUE then HorMod := TRUE else HorMod := FALSE;
    if ChbxBaseForcee.checked = TRUE then BasesMod := TRUE else BasesMod := FALSE;
    if ChbxTranchesForcees.checked = TRUE then TranchesMod := TRUE else TranchesMod := FALSE;
    if ChbxTrentiem.checked = TRUE then TrentMod := TRUE else TrentMod := FALSE;
    // PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement
    if ChbxRegltMod.checked = TRUE then RegltMod := TRUE else RegltMod := FALSE;

    NTrent := StrToInt(FTrentieme.Text);
    DTrent := StrToInt(FTrentDenominateur.Text);
    Ouvres := VALEUR(EdtJoursOuvres.Text);
    Ouvrables := VALEUR(EdtJoursOuvrables.Text);
    HeuresTrav := VALEUR(EdtHorReel.Text);
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    Edtdu := StrToDate(FEdtDu.Text);
    Edtau := StrToDate(FEdtAu.Text);
  end;
end;

procedure TFSaisBul.GRemSalSetEditText(Sender: TObject; ACol, ARow: Integer; const Value: string);
var
  rubrique: string;
begin
  Rubrique := CurG.Cells[0, Curg.Row];
  if RechCommentaire(Rubrique) = FALSE then ACal := TRUE; // on ne declenche pas le calcul du bulletin dans le cas d'insertion d'une rubrique de commentaire
end;
{ Impression du bulletin
Calcul et validation du bulletion en creation et en modification, car le bulletin lit
la Base.
En consultation, ce n'est qu'une impression.
}

procedure TFSaisBul.BImprimerClick(Sender: TObject);
var
  St: string;
begin
  if ActionBul <> taConsultation then BValiderClick(nil); // Calcul et validation du bulletin avant impression
  St := Etab + ';' + CodeSalarie + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
{$IFNDEF EAGLSERVER}
  if (Etab <> '') and (CodeSalarie <> '') and (DateD > 0) and (DateF > 0) then
    //DEB PT214
    if GetParamSocSecur('SO_PGBULDEFAUT', '') = '' then
    begin
      ExecuteSQL('UPDATE PARAMSOC SET SOC_DATA="PBP" WHERE SOC_NOM="SO_PGBULDEFAUT"');
      SetParamSoc('SO_PGBULDEFAUT', 'PBP');
    end;
  ExecuteSQL('UPDATE SALARIES SET PSA_ETATBULLETIN="' + GetParamSocSecur('SO_PGBULDEFAUT', '') + '"' +
    ' WHERE PSA_ETATBULLETIN=""');
    //FIN PT214

    //DEB PT182
    // Si on ne gère pas le mode d'édition
  if not GetParamSocSecur('SO_PGGESTORIDUPSPE', False) then
    AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';BULL;' + BullCompl) //PT22
  else
  begin
      // Vérifier que l'utilisateur a au moins le droit d'éditer un duplicata ou un specimen
      // A ce niveau, on ne propose pas l'édition d'un original
    if JaiLeDroitTag(42314) or JaiLeDroitTag(42316) then
    begin
      AglLanceFiche('PAY', 'MODEEDT_BUL', '', '', 'SAISBUL');
      if PGModeEdition <> '' then
        AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';BULL;' + BullCompl + ';' + PGModeEdition);
    end
    else
      PGIBox(TraduireMemoire('Vous n''êtes pas autorisé à lancer une édition de bulletins'));
  end;
    //FIN PT182
{$ENDIF}
end;
{ Fonction de recherche et des calculs des cumuls dans le bulletin de paie.
Constitution TOb à l'image de la liste à afficher et affection de cette TOb
à la variable globale TheTOB.

Cette TOB sera exploitée par le TOF de la fiche qui va afficher cette TOB.
}

procedure TFSaisBul.CumSessSalClick(Sender: TObject);
var
  TOB_CumAff, T1, T2, LeCumul, TC: TOB;
  Q: TQuery;
  St, CodeC: string;
  DD, DF: TDateTime;
  i: Integer;
  DebEx: Boolean;
begin
  DebEx := FALSE;
  TOB_CumAff := TOB.create('Les Cumuls Bulletin', nil, -1);
  RendDateExerSocial(DateD, DateF, DD, DF);
  if DF < DD then DF := DD;
  St := 'select PCL_CUMULPAIE,PCL_LIBELLE,SUM(PHC_MONTANT) MONTANT FROM CUMULPAIE ' +
    'left join HISTOCUMSAL ON PHC_CUMULPAIE = PCL_CUMULPAIE ' +
    'where ##PCL_PREDEFINI## PHC_SALARIE="' + CodeSalarie + //**//
    '" AND PHC_DATEDEBUT >="' + UsDateTime(DD) + '" AND PHC_DATEFIN <="' + UsDateTime(DF) + '" ' +
    'group by PCL_CUMULPAIE,PCL_LIBELLE';
  Q := OpenSql(st, TRUE);
  // Chargement de l'historique des cumuls salariés en fonction de la requete qui crée les bonnes colonnes
  TOB_CumAff.LoadDetailDB('CUMULPAIE', '', '', Q, FALSE, TRUE);
  Ferme(Q);
  if TOB_CumAff.Detail.Count = 0 then // Cas on est au debut de l'exercice
  begin
    FreeAndNil(TOB_CumAff);
    TOB_CumAff := TOB.create('Les Cumuls Bulletin', nil, -1);
    DebEx := TRUE;
    St := 'select PCL_CUMULPAIE,PCL_LIBELLE FROM CUMULPAIE WHERE ##PCL_PREDEFINI## PCL_CUMULPAIE<>"AAA"'; //**//
    Q := OpenSql(st, TRUE);
    TOB_CumAff.LoadDetailDB('CUMULPAIE', '', '', Q, FALSE, TRUE);
    Ferme(Q);
  end
    // On va comparer la liste des cumuls trouvés avec la liste des cumuls
  else
  begin
    for i := 0 to TOB_Cumuls.detail.count - 1 do
    begin
      LeCumul := TOB_Cumuls.detail[i];
      if LeCumul <> nil then
      begin
        TC := TOB_CumAff.FindFirst(['PCL_CUMULPAIE'], [LeCumul.GetValue('PCL_CUMULPAIE')], FALSE);
        if TC = nil then // Cumul non trouvé, il faut donc le rajouter
        begin
          T1 := TOB.Create('CUMULPAIE', TOB_CumAff, -1);
          if T1 <> nil then
          begin
            T1.PutValue('PCL_CUMULPAIE', LeCumul.GetValue('PCL_CUMULPAIE'));
            T1.PutValue('PCL_LIBELLE', LeCumul.GetValue('PCL_LIBELLE'));
            T1.AddChampSup('MONTANT', FALSE);
            T1.PutValue('MONTANT', 0);
          end;
        end;
      end;
    end;
  end;

  TOB_CumAff.Detail.Sort('PCL_CUMULPAIE'); // tri par numero de cumul

  if TOB_CumulSal = nil then
  begin
    FreeAndNil(TOB_CumAff);
    PGIBox('Vous n''avez pas de cumuls antérieurs', 'Gestion des cumuls');
    exit;
    {
    TOB_CumulSal :=TOB.create ('Les Cumuls de mon bulletin',NIL, -1);
    St:='select PCL_CUMULPAIE,PCL_LIBELLE,SUM(PHC_MONTANT) MONTANT FROM CUMULPAIE '+
    'left join HISTOCUMSAL ON PHC_CUMULPAIE = PCL_CUMULPAIE '+
    'where ##PCL_PREDEFINI## PHC_SALARIE="'+CodeSalarie+  //**//
    '" AND PHC_DATEDEBUT >="'+UsDateTime (DateD)+'" AND PHC_DATEFIN <="'+UsDateTime (DateF)+'" '+
    'group by PCL_CUMULPAIE,PCL_LIBELLE';
    Q:=OpenSql(st, TRUE);
    // Chargement de l'historique des cumuls salariés en fonction de la requete qui crée les bonnes colonnes
    TOB_CumulSal.LoadDetailDB('CUMULPAIE','','',Q,FALSE,TRUE) ;
    Ferme (Q);
    }
  end;
  // Boucle de parcours des cumuls calculés dans la paie pour remplir la colonne En cours de la TOB
  for i := TOB_CumulSal.Detail.Count - 1 downto 0 do
  begin
    T1 := TOB_CumulSal.Detail[I];
    CodeC := T1.GetValue('PHC_CUMULPAIE'); // Recup du code du cumul sur les cumuls en cours
    if (CodeC >= '01') then // and (CodeC <= '99') PT130
    begin
      T2 := TOB_CumAff.FindFirst(['PCL_CUMULPAIE'], [CodeC], FALSE); // Recup de la tob à afficher
      if T2 <> nil then
      begin
        if DebEx = TRUE then
        begin
          T2.AddChampSup('MONTANT', FALSE);
          T2.PutValue('MONTANT', 0);
        end;
        T2.AddChampSup('ENCOURS', FALSE); // Creation d'une colonne supplémentaire pour afficher le cumul en cours de calcul sur ttes les TOB Filles
        T2.PutValue('ENCOURS', T1.GetValue('PHC_MONTANT'));
      end;
    end;
  end;


  for i := TOB_CumAff.Detail.Count - 1 downto 0 do
  begin // Recherche des incoherence entre la liste des cumuls et la liste des cumuls alimentes
    T1 := TOB_CumAff.Detail[I];
    if T1 <> nil then
    begin
      if T1.FieldExists('MONTANT') = FALSE then
      begin
        T1.AddChampSup('MONTANT', FALSE);
        T1.PutValue('MONTANT', 0);
      end;
      if T1.FieldExists('ENCOURS') = FALSE then
      begin
        T1.AddChampSup('ENCOURS', FALSE);
        T1.PutValue('ENCOURS', 0);
      end;
    end;
  end;
{$IFNDEF EAGLSERVER}
  TheTob := TOB_CumAFF;
  AglLanceFiche('PAY', 'BUL_CUMULS', '', '', '');
  TheTOB := nil;
{$ENDIF}
  FreeAndNil(TOB_CumAff);
end;
{ Fonction qui va lancer la fiche de consultation de l'historique des bases de cotisation
- Création d'une TOB qui contient le code et le libellé des rubriques de bases
  contenu dans l'onglet
- LAncement de la fiche
}

procedure TFSaisBul.AccesBasesClick;
var
  T_BaseMere, T_BaseFille: TOB; // TOB des bases de cotisations presentes dans le bulletin
  i: INTEGER;
  st: string;
begin
  T_BaseMere := TOB.create('Les Bases du bulletin', nil, -1);
  for i := 1 to GBas.RowCount - 1 do
  begin // Constitution de la TOB Virtuelle
    if GBas.Cells[0, i] <> '' then
    begin
      T_BaseFille := TOB.create('Une Base', T_BaseMere, -1);
      T_BaseFille.AddChampSup('CODEBASE', False);
      T_BaseFille.AddChampSup('LIBELBASE', False);
      T_BaseFille.PutValue('CODEBASE', GBas.Cells[0, i]);
      T_BaseFille.PutValue('LIBELBASE', GBas.Cells[1, i]);
    end;
  end;
{$IFNDEF EAGLSERVER}
  TheTob := T_BaseMere;
  St := CodeSalarie + ';' + Etab + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
  AglLanceFiche('PAY', 'BUL_BASES', '', '', St);
  TheTOB := nil;
{$ENDIF}
  T_BaseMere.free;
end;

{ Fonction qui va lancer la fiche de consultation de l'historique des cotisations
- Création d'une TOB qui contient le code et le libellé des rubriques de cotisations
  contenu dans l'onglet
- LAncement de la fiche
}

procedure TFSaisBul.AccesCotClick;
var
  T_CotMere, T_CotFille: TOB; // TOB des cotisations presentes dans le bulletin
  i: INTEGER;
  st: string;
begin
  T_CotMere := TOB.create('Les Cotisations du bulletin', nil, -1);
  for i := 1 to GCot.RowCount - 1 do
  begin // Constitution de la TOB Virtuelle
    if GCot.Cells[0, i] <> '' then
    begin
      T_CotFille := TOB.create('Une Base', T_cotMere, -1);
      T_CotFille.AddChampSup('CODEBASE', False);
      T_CotFille.AddChampSup('LIBELBASE', False);
      T_CotFille.PutValue('CODEBASE', GCot.Cells[0, i]);
      T_CotFille.PutValue('LIBELBASE', GCot.Cells[1, i]);
    end;
  end;
{$IFNDEF EAGLSERVER}
  TheTob := T_CotMere;
  St := CodeSalarie + ';' + Etab + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
  AglLanceFiche('PAY', 'BUL_COT', '', '', St);
  TheTOB := nil;
{$ENDIF}
  T_CotMere.free;
end;


{ Fonction qui va lancer la fiche de consultation de l'historique des rémunérations
- Création d'une TOB qui contient le code et le libellé des rubriques de rémunérations
  contenu dans l'onglet
- LAncement de la fiche
}

procedure TFSaisBul.AccesRemClick;
var
  T_RemMere, T_RemFille: TOB; // TOB des rémunérations presentes dans le bulletin
  j: INTEGER;
  st: string;
  Grille: THGrid;
begin
  T_RemMere := TOB.create('Les Rémunérations du bulletin', nil, -1);
  for j := 0 to ComponentCount - 1 do
  begin
    if Components[j] is THGrid then
    begin
      Grille := THGrid(Components[j]);
      if (Grille = GRemSal) or (Grille = GRemSais) or (Grille = GRemAbs) or (Grille = GRemAbt) or (Grille = GRemAvt) // PT202
        or (Grille = GRemNonImp) or (Grille = GRemPrimes) or (Grille = GRemRet) or (Grille = GRemCplt)
        or (Grille = GRemHeures) or (Grille = GRemNonImpSoumis) then RempliTobHistoRem(Grille, T_RemMere, T_RemFille);
    end;
  end;
{$IFNDEF EAGLSERVER}
  TheTob := T_RemMere;
  St := CodeSalarie + ';' + Etab + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
  AglLanceFiche('PAY', 'BUL_REM', '', '', St);
  TheTOB := nil;
{$ENDIF}
  T_RemMere.free;
end;

// Fonction de remplissage de la tob des remunerations en fonction des differentes grilles de remunerations

procedure TFSaisBul.RempliTobHistoRem(Grille: THGrid; T_RemMere, T_RemFille: TOB);
var
  i: Integer;
begin
  for i := 1 to Grille.RowCount - 1 do
  begin // Constitution de la TOB Virtuelle
    if Grille.Cells[0, i] <> '' then
    begin
      T_RemFille := TOB.create('Une Rem', T_RemMere, -1);
      T_RemFille.AddChampSup('CODEBASE', False);
      T_RemFille.AddChampSup('LIBELBASE', False);
      T_RemFille.PutValue('CODEBASE', Grille.Cells[0, i]);
      T_RemFille.PutValue('LIBELBASE', Grille.Cells[1, i]);
    end;
  end;
end;

{
procedure TFSaisBul.GrilleDisable;
begin
  GRemSal.SynEnabled := FALSE;
  GBas.SynEnabled := FALSE;
  GCot.SynEnabled := FALSE;
  GRemHeures.SynEnabled := FALSE;
  GRemPrimes.SynEnabled := FALSE;
  GRemAbs.SynEnabled := FALSE;
  GRemCplt.SynEnabled := FALSE;
  GRemAvt.SynEnabled := FALSE;
  GRemAbt.SynEnabled := FALSE;
  GRemRet.SynEnabled := FALSE;
  GRemNonImp.SynEnabled := FALSE;
end;
}

procedure TFSaisBul.GrilleEnable;
begin
  GRemSal.SynEnabled := TRUE;
  GRemSais.SynEnabled := TRUE; // PT202
  GBas.SynEnabled := TRUE;
  GCot.SynEnabled := TRUE;
  GRemHeures.SynEnabled := TRUE;
  GRemPrimes.SynEnabled := TRUE;
  GRemAbs.SynEnabled := TRUE;
  GRemCplt.SynEnabled := TRUE;
  GRemAvt.SynEnabled := TRUE;
  GRemAbt.SynEnabled := TRUE;
  GRemRet.SynEnabled := TRUE;
  GRemNonImp.SynEnabled := TRUE;
end;
{ Fonction qui renseigne le type de présentation du bulletin
et qui indique la périodicité de calucl du bulletin
Ces infos proviennent de ParamSoc.
}

procedure TFSaisBul.PresentationBul;
begin
  if VH_Paie.PGSaisieBulletin = 'DET' then AffecteDetail
  else AffecteSimple;
  CalculAFaire := VH_Paie.PGCalculBulletin;
  LongueurChamp;
  if VH_Paie.PGAnalytique = FALSE then BVentil.Visible := FALSE
  else BVentil.Visible := TRUE;
  //if ActionBul <> taCreation then BVentil.Enabled := TRUE
  //                           else BVentil.Enabled := FALSE;
end;
{ Fonction de mise en conformite de longueur des champs à saisir dans chaque
grille
}

procedure TFSaisBul.LongueurChamp;
var
  i: Integer;
  Grille: THGrid;
begin
  for i := 0 to ComponentCount - 1 do
  begin
    if Components[i] is THGrid then
    begin
      Grille := THGrid(Components[i]);
      Grille.ColLengths[0] := 4;
      Grille.ColTypes[0] := 'C';
      Grille.ColAligns[0] := taCenter;
      Grille.ColLengths[1] := 35;
      Grille.ColTypes[1] := 'C';
      Grille.ColAligns[1] := taLeftJustify;
// Deb PT200
      if (Grille = GRemSal) or (Grille = GRemSais) or (Grille = GRemAbs) or (Grille = GRemAbt) or (Grille = GRemAvt) // PT202
        or (Grille = GRemNonImp) or (Grille = GRemPrimes) or (Grille = GRemRet) or (Grille = GRemCplt)
        or (Grille = GRemHeures) then
      begin
        Grille.PostDrawCell := RemPostDrawCell;
        Grille.ColLengths[2] := 12;
        Grille.ColTypes[2] := 'R';
        Grille.ColAligns[2] := taRightJustify;
        Grille.ColLengths[3] := 10;
        Grille.ColTypes[3] := 'R';
        Grille.ColAligns[3] := taRightJustify;
        Grille.ColLengths[4] := 10;
        Grille.ColTypes[4] := 'R';
        Grille.ColAligns[4] := taRightJustify;
        Grille.ColLengths[5] := 12;
        Grille.ColTypes[5] := 'R';
        Grille.ColAligns[5] := taRightJustify;
      end
      else
      begin
        if Grille = GCot then
        begin
          Grille.PostDrawCell := CotPostDrawCell;
          Grille.ColLengths[2] := 12;
          Grille.ColTypes[2] := 'R';
          Grille.ColAligns[2] := taRightJustify;
          Grille.ColLengths[3] := 12;
          Grille.ColTypes[3] := 'R';
          Grille.ColAligns[3] := taRightJustify;
          Grille.ColLengths[4] := 12;
          Grille.ColTypes[4] := 'R';
          Grille.ColAligns[4] := taRightJustify;
          Grille.ColLengths[5] := 12;
          Grille.ColTypes[5] := 'R';
          Grille.ColAligns[5] := taRightJustify;
          Grille.ColLengths[6] := 12;
          Grille.ColTypes[6] := 'R';
          Grille.ColAligns[6] := taRightJustify;
        end;
        if Grille = GBas then
        begin
          Grille.PostDrawCell := BasPostDrawCell;
          {            Grille.ColLengths [7]:=12;
                      Grille.ColLengths [8]:=12;
                      Grille.ColLengths [8]:=12;}
          Grille.ColTypes[2] := 'R';
          Grille.ColAligns[2] := taRightJustify;
          Grille.ColTypes[3] := 'R';
          Grille.ColAligns[3] := taRightJustify;
          Grille.ColTypes[4] := 'R';
          Grille.ColAligns[4] := taRightJustify;
          Grille.ColTypes[5] := 'R';
          Grille.ColAligns[5] := taRightJustify;
          Grille.ColTypes[6] := 'R';
          Grille.ColAligns[6] := taRightJustify;
          Grille.ColTypes[7] := 'R';
          Grille.ColAligns[7] := taRightJustify;
          Grille.ColTypes[8] := 'R';
          Grille.ColAligns[8] := taRightJustify;
          Grille.ColTypes[9] := 'R';
          Grille.ColAligns[9] := taRightJustify;
          Grille.ColLengths[2] := 12;
          Grille.ColLengths[3] := 12;
          Grille.ColLengths[4] := 12;
          Grille.ColLengths[5] := 12;
          Grille.ColLengths[6] := 12;
          Grille.ColLengths[7] := 12;
          Grille.ColLengths[8] := 12;
          Grille.ColLengths[9] := 12;
          Grille.ColTypes[10] := 'C';
          Grille.ColAligns[10] := taCenter; // FIN PT200
        end;
      end;
    end;
  end;
end;

procedure TFSaisBul.AffecteSimple;
begin
  PRemAbs.Enabled := FALSE;
  PRemAbs.Tabvisible := FALSE;
  PRemAbt.Enabled := FALSE;
  PRemAbt.Tabvisible := FALSE;
  PRemAvt.Enabled := FALSE;
  PRemAvt.Tabvisible := FALSE;
  PRemHeures.Enabled := FALSE;
  PRemHeures.Tabvisible := FALSE;
  PRemCplt.Enabled := FALSE;
  PRemCplt.Tabvisible := FALSE;
  PRemPrimes.Enabled := FALSE;
  PRemPrimes.Tabvisible := FALSE;
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  PRemPrimeNonImpSoumis.Enabled := FALSE;
  PRemPrimeNonImpSoumis.Tabvisible := FALSE;
end;

procedure TFSaisBul.AffecteDetail;
begin
  PRemAbs.Enabled := TRUE;
  PRemAbs.Tabvisible := TRUE;
  PRemAvt.Enabled := TRUE;
  PRemAvt.Tabvisible := TRUE;
  PRemAbt.Enabled := TRUE;
  PRemAbt.Tabvisible := TRUE;
  PRemHeures.Enabled := TRUE;
  PRemHeures.Tabvisible := TRUE;
  PRemCplt.Enabled := TRUE;
  PRemCplt.Tabvisible := TRUE;
  PRemPrimes.Enabled := TRUE;
  PRemPrimes.Tabvisible := TRUE;
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  PRemPrimeNonImpSoumis.Enabled := TRUE;
  PRemPrimeNonImpSoumis.Tabvisible := TRUE;
end;

procedure TFSaisBul.BCalculBullClick(Sender: TObject);
//var SystemTime1,SystemTime2 : TSystemTime;
begin
  Enabled := FALSE;
  CalculAFaire := TRUE;
//  Cumul12 := RendCumulSalSess('12'); { PT92-1 }  PT127-2 Mise en commentaire
  //GetLocalTime(SystemTime0);
  CalculBullSaisie;
(* PT127-2 Mise en commentaire
  { DEB PT92-1 }
  if ((Cumul12 <> RendCumulSalSess('12')) or (PGTopRecalculCp)) and (FirstBull) then { PT92-1 17/05/2004 }
  begin
    RecalculCp(False, True);
    CalculBullSaisie;
    PGTopRecalculCp := False; { PT92-1 17/05/2004 }
  end;                   *)
  { FIN PT92-1 }
  CalculAFaire := VH_Paie.PGCalculBulletin;
  Enabled := TRUE;
end;

{ Fonction de sauvegarde d'un bulletin de paie}

// d PT115
//function SauvegardeBul(CodeSalarie, Etab: string; DateD, DateF, DateS: TDateTime; Tob_Rub: TOB; NTrent, DTrent: Integer; ActionFiche: Boolean): Boolean;

function SauvegardeBul(CodeSalarie, Etab: string; DateD, DateF, DateS: TDateTime; Tob_Rub: TOB; NTrent, DTrent: Integer; ActionFiche: Boolean; var TOB_IJSS, TOB_Maintien: TOB): Boolean;
// f PT115
var
  Cumul, St, S1, S2, rub: string;
  T_ECP, TOB_LesCumulSal, TT: TOB;
  TH, THisto, TMonHisto, TM: TOB;
  Mt: Double;
  i, MaxChamp, Rep: Integer;
  Q: TQuery;
  ModifHist: Boolean;
  Trace: TStringList;
  St1: string;
  // added by XP le 17-03-2003
  iChampCumul, iChampMt: integer;
  DatePaiement: TDateTime;
  CodePopul: string; //PT195
  // PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
  // PT45 : 06/09/2002 V585 PH Historisation 5 eme elt de salaire
  // PT74   08/08/2003 V_421 PH FQ 10740 Rajout Champ etablissement dans table HISTOSALARIE
const
  HistS: array[1..50] of string = (
    'CODEEMPLOI', 'LIBELLEEMPLOI', 'QUALIFICATION', 'COEFFICIENT', 'INDICE', 'NIVEAU', 'CODESTAT', 'TRAVAILN1',
    'TRAVAILN2', 'TRAVAILN3', 'TRAVAILN4', 'SALAIREMOIS1', 'SALAIREMOIS2', 'SALAIREMOIS3', 'SALAIREMOIS4', 'SALAIREMOIS5',
    'SALAIRANN1', 'SALAIRANN2', 'SALAIRANN3', 'SALAIRANN4', 'SALAIRANN5',
    'DATELIBRE1', 'DATELIBRE2', 'DATELIBRE3', 'DATELIBRE4', 'BOOLLIBRE1',
    'BOOLLIBRE2', 'BOOLLIBRE3', 'BOOLLIBRE4', 'LIBREPCMB1', 'LIBREPCMB2', 'LIBREPCMB3', 'LIBREPCMB4', 'PROFIL',
    'PERIODBUL', 'DADSPROF', 'DADSCAT', 'TAUXPARTIEL', 'PROFILREM', 'CONDEMPLOI', 'HORAIREMOIS', 'TAUXHORAIRE',
    'HORHEBDO', 'HORANNUEL', 'SALAIRETHEO', 'LIBRE1', 'LIBRE2', 'LIBRE3', 'LIBRE4', 'ETABLISSEMENT');
const
  HistH: array[1..50] of string = (
    'CODEEMPLOI', 'LIBELLEEMPLOI', 'QUALIFICATION', 'COEFFICIENT', 'INDICE', 'NIVEAU', 'CODESTAT', 'TRAVAILN1',
    'TRAVAILN2', 'TRAVAILN3', 'TRAVAILN4', 'SALAIREMOIS1', 'SALAIREMOIS2', 'SALAIREMOIS3', 'SALAIREMOIS4', 'SALAIREMOIS5',
    'SALAIREANN1', 'SALAIREANN2', 'SALAIREANN3', 'SALAIREANN4', 'SALAIREANN5',
    'DTLIBRE1', 'DTLIBRE2', 'DTLIBRE3', 'DTLIBRE4', 'BOOLLIBRE1',
    'BOOLLIBRE2', 'BOOLLIBRE3', 'BOOLLIBRE4', 'CBLIBRE1', 'CBLIBRE2', 'CBLIBRE3', 'CBLIBRE4', 'PROFIL',
    'PERIODBUL', 'DADSPROF', 'DADSCAT', 'TTAUXPARTIEL', 'PROFILREM', 'CONDEMPLOI', 'HORAIREMOIS', 'TAUXHORAIRE',
    'PGHHORHEBDO', 'PGHHORANNUEL', 'SALAIRETHEO', 'CHARLIBRE1', 'CHARLIBRE2', 'CHARLIBRE3', 'CHARLIBRE4', 'ETABLISSEMENT');
  // FIN PT24
begin
  MaxChamp := 50;
  // FIN PT74
       // FIN PT45
  result := TRUE;
  if ActionBul = taConsultation then exit;
  // PT37 : 15/05/2002 V582 PH FQ 331 Controle si ventilation analytique existente en sauvegarde
  if (TypeTraitement <> 'PREPA') and (VH_PAIE.PGAnalytique) then
  begin
    if (TOBAna = nil) or (TOBAna.detail.count = 0) then
    begin
      st := 'Attention, il n''y a pas de ventilation analytique pour ce salarié';
      st := st + '#13#10 Voulez vous enregistrer le bulletin ?';
      Rep := PGIAsk(st, 'Gestion analytique');
      if rep = mrNo then exit;
    end;
  end;

  // FIN PT37
  st := '';

  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  if (TypeTraitement <> 'PREPA') then Trace := TStringList.Create;
  SourisSablier;
  if ActionFiche then X.AffecteEtatFiche(FALSE);
  try

    BeginTrans;
    st := 'DELETE FROM PAIEENCOURS WHERE PPU_SALARIE="' + CodeSalarie + '" AND PPU_ETABLISSEMENT="' + Etab + '" AND PPU_DATEDEBUT="' + USDateTime(DateD) + '" AND PPU_DATEFIN="' +
      USDateTime(DateF) + '"';
    ExecuteSQL(st);
    st := 'DELETE FROM HISTOBULLETIN WHERE PHB_SALARIE="' + CodeSalarie + '" AND PHB_ETABLISSEMENT="' + Etab + '" AND PHB_DATEDEBUT="' + USDateTime(DateD) + '" AND PHB_DATEFIN="'
      + USDateTime(DateF) + '"';
    ExecuteSQL(St);
    if VH_PAIE.PGAnalytique = TRUE then InitMove(5, ' ') // gestion des lignes analytiques
    else InitMove(3, ' ');
    // Boucle analyse de la tob suite mauvaise saisie ==> suppression des objets non saisis en entier
    TT := TOB_Rub.findFirst([''], [''], TRUE);
    if (TT.FieldExists('PHB_SALARIE')) and (TT.getvaleur(iPHB_SALARIE) <> CodeSalarie)
      then TT.free;

    TT := TOB_Rub.findNext([''], [''], TRUE);
    while TT <> nil do
    begin
      Rub := TT.getvaleur(iPHB_RUBRIQUE);
      if (POS('.R', Rub) > 0) then
        TT.PutValeur(iPHB_RUBRIQUE, Copy(Rub, 1, 4));
      if (TT.FieldExists('PHB_SALARIE')) and (TT.getvaleur(iPHB_SALARIE) <> CodeSalarie)
        then TT.free;
      TT := TOB_Rub.findNext([''], [''], TRUE);
    end;
    with TOB_Rub do
    begin
      PutValue('PPU_CBRUT', RendCumulSalSess('01'));
      PutValue('PPU_CBRUTFISCAL', RendCumulSalSess('02'));
      PutValue('PPU_CNETIMPOSAB', RendCumulSalSess('09'));
      PutValue('PPU_CNETAPAYER', RendCumulSalSess('10'));
      PutValue('PPU_CCOUTSALARIE', RendCumulSalSess('08'));
      PutValue('PPU_CCOUTPATRON', RendCumulSalSess('07'));
      PutValue('PPU_CPLAFONDSS', RendCumulSalSess('30'));
      PutValue('PPU_CBASESS', RendCumulSalSess('31'));
      // PT95   18/06/2004 V_50 PH FQ 11164 Alimentation de la Base SS pratiquée dans la table PAIIENCOURS
      PutValue('PPU_CBASESSPRAT', RendCumulSalSess('40'));
      PutValue('PPU_CHEURESTRAV', RendCumulSalSess('20'));
      PutValue('PPU_HEURESREELLES', ChpEntete.HeuresTrav); //PT71-4 pour maj en prep auto

{$IFNDEF CPS1}
      //DEB PT195
      CodePopul := '';
      Q := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
        + ' WHERE PNA_SALARIE = "' + CodeSalarie + '"'
        + ' AND PNA_TYPEPOP = "PAI"', True);
      if not Q.Eof then
        CodePopul := Q.FindField('PNA_POPULATION').AsString;
      Ferme(Q);
      PutValue('PPU_USEREDIT', CodePopul);
      //FIN PT195
{$ENDIF}

      DatePaiement := GetValue('PPU_PAYELE');
      SetAllModifie(TRUE);
      result := InsertDB(nil, FALSE);
    end;

    MoveCur(FALSE);
    if not result then Showmessage('Une erreur s''est produite pendant l''écriture des lignes')
    else
    begin
      ExecuteSQL('DELETE FROM HISTOCUMSAL WHERE PHC_SALARIE="' + CodeSalarie + '" AND PHC_ETABLISSEMENT="' + Etab + '" AND PHC_DATEDEBUT="' + USDateTime(DateD) +
        '" AND PHC_DATEFIN="' + USDateTime(DateF) + '" AND PHC_REPRISE <> "X"');
      iCHampCumul := TOB_CumulSal.detail[0].GetNumChamp('PHC_CUMULPAIE');
      iCHampMt := TOB_CumulSal.detail[0].GetNumChamp('PHC_MONTANT');
      for i := tob_cumulsal.detail.count - 1 downto 0 do
      begin
        with tob_cumulsal.detail[i] do
        begin
          Cumul := GetValeur(iChampCumul);
          Mt := GetValeur(iChampMt);
          if ((Cumul < '00')) or ((Cumul > '99') and IsNumeric(Cumul[1])) or (Cumul = '') then // PT139 On adment un cumul alpha en 1ere position
            HShowMessage('2;Cumul Erroné: ' + Cumul + ' ;Vous devez saisir un code numérique supérieur à 0!;W;O;O;;;', '', '');
          if Mt = 0 then Free; // on elimnine les cumuls non alimentés par le calcul du bulletin
        end;
      end; // fin du for

      // bug du BDE des 10000 !!
      TOB_LesCumulSal := TOB.Create('Les_Cumuls_du_Salarie', nil, -1);
      for i := 0 to tob_cumulsal.detail.count - 1 do
        with TOB.create('HISTOCUMSAL', TOB_LesCumulSal, -1) do
          Dupliquer(TOB_CumulSal.detail[i], FALSE, TRUE, TRUE);

      result := TOB_LesCumulSal.InsertDB(nil, FALSE);

      FreeAndNil(TOB_LesCumulSal);

      //PT81   23/09/2003 V_421 PH maj saisie arret en modification de bulletin
      if Assigned(TOB_SaisieArret) then PGMajHistoRetenue(TOB_SaisieArret, DatePaiement);

      MoveCur(FALSE);
      if not result then Showmessage('Une erreur s''est produite pendant l''écriture des cumuls salariés');

      T_ECP := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', Etab);

      //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
      //PT38-1 on ne valide pas les CP si acq bulletin pré solde ou cloturé
      if ActionBul = TaCreation then SoldeOuClot := False;

      { PT89 }
      GblCP := (VH_Paie.PGCongesPayes) and (T_ECP.GetValue('ETB_CONGESPAYES') = 'X') and (TOB_SALARIE.GetValeur(IPSA_CONGESPAYES) = 'X');
      if (BullCompl <> 'X') and (SoldeOuClot = False) and (GblCP) then { PT127-4 }
      begin
        RetireAcquiscourant(Tob_Acquis);
        AffectBaseAcquisEncours;
        {  PT12 mise en commentaire puis supprimé
        **On n'alimente plus les acqs sur la validation mais en chargement de fiche **
        **afin que la prise de CP sur l'acquis s'effectue** }
        if ((DateS >= DateD) and (DateS <= DateF)) and (PGRecupereProfilCge(Etab) <> '') then //PT84
        begin
          SoldeAcquis(T_MvtAcquis, True, DateF); //PT34 Solde des consommés des acquis en cours PT151
          SoldeAcquis(Tob_Acquis, True, DateF); //PT12 Ajout False   //PT38-2 Affecte True, deplacement(Après avoir soldé T_MvtAcquis) PT151
          SoldeAcquis(Tob_Delta, True, DateF); { PT151 }
        end;
        SalEcritCP(Tob_Pris); // Ecriture des mvts Cp generés dans le bulletin

      end;

      //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
      if (BullCompl <> 'X') then SalEcritAbs(tob_abs);
{$IFNDEF CCS3}
      // d PT94 IJSS  Maj table REGLTIJSS
      if (VH_Paie.PGGestIJSS) then
        if (BullCompl <> 'X') then SalEcritIJSS(tob_IJSS);
      // f PT94 IJSS
{$ENDIF}

{$IFNDEF CCS3}
      // d PT110 Maintien  Maj table MAINTIEN
      if (VH_Paie.PGMaintien) then
        if (BullCompl <> 'X') then SalEcritMaintien(tob_Maintien);
      // f PT110 Maintien
{$ENDIF}

      MoveCur(FALSE);

      // Ecriture des ventilations analytiques du bulletin cad la tob des ventana
      if VH_PAIE.PGAnalytique = TRUE then // gestion des lignes analytiques
      begin
        // Reaffectation des ventilations analytiques des cotisations en fonction de la valorisation des remunerations
        if ActionBul <> taCreation then PGReaffSectionAnal(TOB_Rub, TOBAna)
          // PT23 : 21/01/2002 V571 PH Réaffectation des dates début et fin pour identifiant analytique = param à la fonction
        else PGReaffSectionAnal(TOB_Rub, TOBAna, DateD, DateF);
        SauvegardeAnal(TOBAna, TOB_Rub, TOB_VentilRem, TOB_VentilCot, TRUE, CodeSalarie, DateD, DateF);
        MoveCur(FALSE);
      end;

      FiniMove;
    end;

    // PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
    // Ecriture enregistrement historique salarié
    if VH_Paie.PgHistorisation then
    begin
      TH := TOB.create('Mes lignes histosalarie', nil, -1);
      TM := TOB.create('Ma ligne histosalarie', nil, -1);
      st := 'SELECT * FROM HISTOSALARIE WHERE PHS_SALARIE="' + CodeSalarie + '" AND ' +
        'PHS_DATEEVENEMENT <="' + USDateTime(DateD) + '" ORDER BY PHS_DATEEVENEMENT DESC';
      Q := OpenSql(st, TRUE, 1);
      TH.LoadDetailDB('HISTOSALARIE', '', '', Q, FALSE);
      THisto := TH.FindFirst(['PHS_SALARIE'], [CodeSalarie], TRUE);
      ferme(Q);
      TMonHisto := TOB.Create('HISTOSALARIE', TM, -1);
      TMonHisto.putvalue('PHS_SALARIE', CodeSalarie);
      ModifHist := FALSE;
      for i := 1 to MaxChamp do
      begin
        if (THisto = nil) or (THisto.getvalue('PHS_' + HistH[i]) <> TOB_Salarie.getvalue('PSA_' + HistS[i])) then
        begin
          TMonHisto.putvalue('PHS_B' + HistH[i], 'X');
          ModifHist := True;
        end;
        TMonHisto.putvalue('PHS_' + HistH[i], TOB_Salarie.getvalue('PSA_' + HistS[i]));
      end;

      if ModifHist then
      begin
        if THisto = nil then TMonHisto.putValue('PHS_COMMENTAIRE', 'Init du ' + DateToStr(DateD) + ' au ' + DateToStr(DateF))
        else TMonHisto.putValue('PHS_COMMENTAIRE', 'Paie du ' + DateToStr(DateD) + ' au ' + DateToStr(DateF));
        TMonHisto.putvalue('PHS_DATEEVENEMENT', DateD);
        // PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
        TMonHisto.putvalue('PHS_DATEAPPLIC', DateF);
        // PT45 : 06/09/2002 V585 PH Historisation 5 eme elt de salaire + Date retroactivité par defaut = date debut paie
        TMonHisto.putvalue('PHS_DATERETRO', DateD);

        TMonHisto.putvalue('PHS_ORDRE', 1);
        // Suppression enrg déjà existant historique salarié
        ExecuteSQL('DELETE FROM HISTOSALARIE WHERE PHS_SALARIE="' + CodeSalarie + '" AND PHS_DATEEVENEMENT="' + USDateTime(DateD) + '" AND PHS_DATEAPPLIC="' +
          USDateTime(DateF) + '"');
        TM.InsertDB(nil, FALSE);
      end;
      FreeAndNil(TH);
      FreeAndNil(TM);
    end;
    // FIN PT24
    //  Fin Ecriture enregistrement historique salarié
    CommitTrans;
    if ActionFiche then X.AffecteEtatFiche(TRUE);
    { DEB PT127-2 }
    if ((DateS >= DateD) and (DateS <= DateF)) and (X <> nil) then
    begin
      X.CBModifAcquis.enabled := False;
      X.CPMOIS.enabled := False;
      X.CPSUPP.enabled := False;
      X.CPANC.enabled := False;
    end;
    { FIN PT127-2 }
    ExecuteSql('UPDATE PGHISTRETRO SET PGT_DATEINTEG="' + UsdateTime(DateF) + '" WHERE PGT_SALARIE="' + CodeSalarie + '" AND ' + //PT205
      'PGT_DATEVALIDITE <= "' + UsDateTime(DateF) + '" AND PGT_DATEVALIDITE >="' + UsDateTime(DateD) + '"');
    SourisNormale;
  except
    Rollback;
    if ActionFiche then X.AffecteEtatFiche(TRUE);
    SourisNormale;
    if (TypeTraitement <> 'PREPA') then PGIBox('Une erreur est survenue lors de la validation du bulletin', 'Saisie du bulletin'); // PT137
    result := FALSE;
  end;

  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  if (TypeTraitement <> 'PREPA') then
  begin
    S1 := DateToStr(DateD);
    S2 := DateToStr(DateF);
    St1 := '';
    if not result then st1 := 'Erreur de ';
    if ActionBul = taModification then st := 'Modification du' else st := 'Création du';
    Trace.add(St1 + St + ' bulletin du ' + S1 + ' au ' + S2 + ' pour le salarié ' + CodeSalarie);
    CreeJnalEvt('001', '003', 'OK', nil, nil, Trace);
    Trace.free;
  end;
  // FIN PT39
end;

{ Accès au paramètrage des cumuls de la paie }

procedure TFSaisBul.CumSessClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41150)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
  end
  else
    AglLanceFiche('PAY', 'CUMUL', '', '', '');
{$ENDIF}
end;

procedure TFSaisBul.SalarieClick(Sender: TObject);
var
  DateFin: TDateTime; //PT198
begin
  // PT93   12/05/2004 V_50 PH FQ 10973 Restriction droit accès au salarié
{$IFNDEF EAGLSERVER}
  If not JaiLeDroitTag(200001) then  //PT223
    AglLanceFiche('PAY', 'SALARIE', '', CodeSalarie, 'ETBBLQ;ACTION=CONSULTATION') //PT223
  else
    AglLanceFiche('PAY', 'SALARIE', '', CodeSalarie, 'ETBBLQ');
{$ENDIF}
  // FIN PT93
    // PT54   27/11/2002 V591 PH Récupération date de sortie et recalcul du trentième si modif salarié dans le bulletin
  ActionSal := TRUE;
  //PT41 : 08/08/2002 V582 PH Rechargement des infos salariés après accès fiche
  ChargeSalarie(TRUE);
  AlimChampFiche;

  //DEB PT198
  DateFin := StrToDate(FDate2.Text);
  if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
  begin
    Nettoyage_EltDynSal();
    Nettoyage_EltDynPop();
    Nettoyage_EltDynEtab();
    initTOB_EltDynSal(CodeSalarie, DateFin);
    initTOB_EltDynEtab(Etab, DateFin);
    initTOB_EltDynPop(CodeSalarie, DateFin);
  end;
  //FIN PT198
end;

procedure TFSaisBul.EtablissementClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'ETABSOCIAL', '', Etab, Etab);
{$ENDIF}
end;

procedure TFSaisBul.FicheIndClick(Sender: TObject);
var
  SQL, SRange, Exer: string;
  DateDPer, DateFPer: TDateTime;
  QExer: TQuery;
  //  StLE: Integer;
begin
  { DEB PT85 }
  //AglLanceFiche ('PAY','FICHEIND', '', '', 'BULL;'+Etab+';'+CodeSalarie+';'+DateToStr(DateD)+';'+DateToStr(DateF));
  Exer := '';
  DateDPer := Idate1900;
  DateFPer := Idate1900;
  QExer := OpenSql('SELECT MAX(PEX_EXERCICE),MAX(PEX_DATEDEBUT),MAX(PEX_DATEFIN) FROM EXERSOCIAL ' +
    'WHERE PEX_DATEDEBUT<="' + UsDateTime(DateD) + '" ' +
    'AND PEX_DATEFIN>="' + UsDateTime(DateF) + '"', true);
  if not QExer.eof then
  begin
    Exer := QExer.Fields[0].AsString;
    DateDPer := QExer.Fields[1].AsDateTime;
    DateFPer := QExer.Fields[2].AsDateTime;
  end;
  Ferme(QExer);
  if Exer <> '' then
  begin
    QExer := OpenSql('SELECT MIN(PPU_DATEDEBUT),MAX(PPU_DATEFIN) FROM PAIEENCOURS ' +
      'LEFT JOIN EXERSOCIAL ON PPU_DATEDEBUT>=PEX_DATEDEBUT AND PPU_DATEFIN<=PEX_DATEFIN ' +
      'WHERE PEX_EXERCICE="' + Exer + '"', TRUE);
    if not QExer.eof then
    begin
      DateDPer := QExer.Fields[0].AsDateTime;
      DateFPer := QExer.Fields[1].AsDateTime;
    end;
    Ferme(QExer);
  end;

  SQL := 'SELECT DISTINCT PHB_RUBRIQUE,PHB_ORDREETAT,PHB_NATURERUB,PSA_SALARIE,PSA_LIBELLE,PHB_SALARIE,' +
    'PSA_NUMEROSS,PSA_PRENOM,PSA_ADRESSE1, PSA_ADRESSE3,PSA_ADRESSE2,PSA_CODEPOSTAL,PSA_VILLE,' +
    'PSA_DATEENTREE,PSA_DATESORTIE, PSA_QUALIFICATION,PSA_COEFFICIENT,PSA_CODEEMPLOI,PSA_PGMODEREGLE,' +
    'PSA_LIBELLEEMPLOI, PCT_LIBELLE,PRM_LIBELLE ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN SALARIES ON PSA_SALARIE=PHB_SALARIE ' +
    'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PHB_RUBRIQUE=PCT_RUBRIQUE ' +
    'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
    'WHERE (((PHB_MTREM<>0 OR PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0) AND PHB_IMPRIMABLE="X") AND PHB_NATURERUB<>"BAS") ' +
    'AND PHB_DATEFIN>="' + USDateTime(DateDPer) + '"  AND PHB_DATEFIN<="' + USDateTime(DateFPer) + '" ' +
    'AND (PHB_ETABLISSEMENT = "' + Etab + '" AND PHB_SALARIE = "' + CodeSalarie + '")  ' +
    'ORDER BY PHB_SALARIE,PHB_ORDREETAT,PHB_NATURERUB,PHB_RUBRIQUE';

  SRange := 'XX_VARIABLEDEB=' + DateToStr(DateDPer)
    + '`XX_VARIABLEFIN=' + DateToStr(DateFPer)
    + '`CSAL=X'
    + '`CALPHA=-'
    + '`STWHERE=AND (PHB_ETABLISSEMENT = "' + Etab + '" AND PHB_SALARIE = "' + CodeSalarie + '")'
    + '`DOSSIER=' + GetParamSoc('SO_LIBELLE')
    + '`PHB_SALARIE=' + CodeSalarie
    + '`PHB_SALARIE_=' + CodeSalarie
    + '`PHB_ETABLISSEMENT=' + Etab
    + '`PHB_ETABLISSEMENT_=' + Etab;
  if GetParamSoc('SO_PGTENUEEURO') = True then SRange := SRange + '`CKEURO=X'
  else SRange := SRange + '`CKEURO=-';

  //StLE:=PGiAsk('Voulez-vous éditer la fiche individuelle sous forme de liste d''exportation?',Caption);
{$IFNDEF EAGLSERVER}
  LanceEtat('E', 'PFI', 'PFI', True, False, False, nil, SQL, '', False, 0, SRange);
{$ENDIF}
  { FIN PT85 }
end;
{ Accès direct au menu Zoom }

procedure TFSaisBul.BZoomClick(Sender: TObject);
begin
  if popupmenuzoom.items.Count > 0 then
    popupmenuzoom.Popup(Mouse.CursorPos.x, Mouse.CursorPos.x);
end;

procedure TFSaisBul.BasesCotis1Click(Sender: TObject);
begin
  AccesBasesClick;
end;

procedure TFSaisBul.GereQuery;
var
  st, Etab: string;
  Q: TQuery;
  //ActionSpec : TCloseAction;
begin
  if QMul = nil then exit;
  if QMul.EOF then
  begin
    //   PGIBox ('Attention, Il n''y a plus de paie à traiter par rapport à la liste !', 'Saisie des bulletins');
    //   ActionSpec := caFree;
    //   exit;
  end;
  CodeSalarie := QMul.FindField('PSA_SALARIE').AsString;
  if TypeList = TRUE then // Query complète avec Paieencours
  begin
    DateD := QMul.FindField('PPU_DATEDEBUT').AsDateTime;
    DateF := QMul.FindField('PPU_DATEFIN').AsDateTime;
    if (QMul.FindField('PPU_TOPGENECR').AsString = 'X') or (QMul.FindField('PPU_TOPCLOTURE').AsString = 'X')
      then ActionBul := taConsultation else ActionBul := taModification;
  end
  else
  begin
    Etab := QMul.FindField('PSA_ETABLISSEMENT').AsString;
    st := 'SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + Etab + '"' + ' AND PPU_SALARIE="' + CodeSalarie + '" AND PPU_DATEDEBUT >="' + USDateTime(DateD) +
      '" AND PPU_DATEFIN <="' + USDateTime(DateF) + '"';
    Q := OpenSQL(St, FALSE);
    if Q.EOF then ActionBul := taCreation else ActionBul := taModification;
    ActionBulCp := ActionBul; { PT127-2 }
    ferme(Q);
  end;
  EtatInitial;
  AfficheActionB;
end;

procedure TFSaisBul.TablesdossierClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41411)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
  end
  else
    AglLanceFiche('PAY', 'COEFFICIENT_MUL', 'PMI_TYPENATURE=INT', '', '');
{$ENDIF}
end;

procedure TFSaisBul.VoirhistoriquecoTisations1Click(Sender: TObject);
begin
  AccesCotClick;
end;

procedure TFSaisBul.VoirhistoriqueRemunrations1Click(Sender: TObject);
begin
  AccesRemClick;
end;

procedure TFSaisBul.GBasRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  if ChbxBaseForcee.Checked = FALSE then exit;
  CalculAFaire := TRUE;
end;

// Gestion des Modules Complémentaires : CP, Gestion des Absences, Saisie par rubrique
// L'ecriture des donnees se fait uniquement dans la phase de validation/ecriture du bulletin
{PT7 : Ajout param var StMsgErr : string}

// d PT111  PT119
//procedure RecupInfoModules(TOB_Rub, Tob_Sal: TOB; ActionB: TActionBulletin; Auto, GCPAuto: boolean; var StMsgErr: string);
//procedure RecupInfoModules(TOB_Rub, Tob_Sal: TOB; ActionB: TActionBulletin; Auto, GCPAuto: boolean; var StMsgErr: string; GenereMaintien, IJSS: boolean);

procedure RecupInfoModules(TOB_Rub, Tob_Sal: TOB; ActionB: TActionBulletin; Auto, GCPAuto: boolean; var StMsgErr: string; GenereMaintien, IJSS: boolean; var TOB_IJ: TOB);
// f PT111 PT119
var
  TETABSAL: Tob;
  TheSal, TheEtab: string;
  Subrogation: boolean; //PT119
  Champcateg, RubNettes: string; //PT119
  SortieSal, EntreeSal: Boolean; // PT157 //PT172
  DateESalarie: TDateTime; // PT172

  //Debut PT209 : Intégration automatique des saisies arrêt du salarié
  procedure IntegreRubSaisiesArret(TobRubriquesDuSalarie, TobSalarie: TOB; DateDPaie, DateFPaie: TDateTime);
  var
    TOB_SaisieArret, TempTob_RemSaisieArret, TobRetenueSal: Tob;
    indexDetail: Integer;
    Salarie: string;
  begin
    if (VH_Paie.PGSaisieArret and VH_Paie.PGIntegAutoRub) then
    begin
      { Recherche des saisies du salarié }
      Salarie := TobSalarie.GetValeur(iPSA_SALARIE);
      { On force un gros montant de salaire pour ramener toutes les saisies. Le montant sera calculé lors de l'intégration, Pour l'instant, seul la liste des saisies nous intéresse. }
      TOB_SaisieArret := GetTobSaisieArret(Salarie, DateFPaie);
      if Assigned(TOB_SaisieArret) then
      begin
        for indexDetail := 0 to TOB_SaisieArret.Detail.count - 1 do
        begin
          { On récupère la variable de paie, il faut trouver la rubrique de rémunération correspondante }
          TempTob_RemSaisieArret := TOB_Rem.FindFirst(['PRM_MONTANT', 'PRM_TYPEMONTANT', 'PRM_PREDEFINI'], [TOB_SaisieArret.Detail[indexDetail].getString('PTR_RUBRIQUE'), '03', 'CEG'], False);
          if Assigned(TempTob_RemSaisieArret) then
          begin
            { On vérifie que la rubrique est déjà intégrée au bulletin. }
            if not assigned(TobRubriquesDuSalarie.FindFirst(['PHB_RUBRIQUE'], [TempTob_RemSaisieArret.getString('PRM_RUBRIQUE')], False)) then
            begin
              { Sinon, on l'intègre }
              IntegreRubBul(TobSalarie, TobRubriquesDuSalarie, TempTob_RemSaisieArret.getString('PRM_RUBRIQUE'), 'AAA');
              IntegreAuto := TRUE;
            end;
          end;
        end;
        FreeAndNil(TOB_SaisieArret);
      end;
    end
  end;
  //Fin PT209

begin // B1
  CalcFractOk := FALSE; // PT224
  // DEB PT104 Affichage des CPs modifiés si bulletin cloturé
  if ActionB = TaConsultation then
  begin
    IntegreCPConsult(TOB_RUB);
    exit;
  end;
  // FIN PT104
  X.PGEnabledRaccourciCp; //PT84
  SoldeOuClot := False; //PT38-1
  DateSSalarie := Tob_sal.GetValeur(iPSA_DATESORTIE);
  DateESalarie := Tob_sal.GetValeur(iPSA_DATEENTREE); // PT152
  TheSal := TOB_SALARIE.GetValeur(iPSA_SALARIE);
  TheEtab := TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT);
  // DEB PT157
  if (DateSSalarie >= DateD) and (DateSSalarie <= DateF) then SortieSal := TRUE
  else SortieSal := FALSE;
  // FIN PT157
  // DEB PT172
  if (DateESalarie > DateD) then EntreeSal := TRUE
  else EntreeSal := FALSE;
  // FIN PT172
  // PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  // PT88   22/03/2004 V_50 PH Intégration de histosaisprim dans un bulletin complémentaire
  // PT40 : 03/06/2002 V582 PH Integration des lignes de la saisie des primes dans le bulletin complémentaire
  if BullCompl = 'X' then
  begin
    if (ActionB = taCreation) or (ACtionB = PremCreation) then
    begin
      Tob_PSP := Tob.create('Historique saisie primes', nil, -1);
      RecupereHistoSaisPrim(Tob_PSP, DateD, Datef, TheSal, ActionB);
      IntegreHistoSaisPrim(Tob_Rub, Tob_PSP, DateD, DateF, TheSal, TheEtab, 'C');
      Tob_PSD := Tob.create('Historique saisie Rub', nil, -1);
      RecupereHistoSaisRub(Tob_PSD, DateD, Datef, TheSal, SortieSal, 'BCP'); // PT157
      IntegreHistoSaisrub(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'C');
    end
    else
    begin
      Tob_PSP := Tob.create('Historique saisie primes', nil, -1);
      RecupereHistoSaisPrim(Tob_PSP, DateD, Datef, TheSal, ActionB);
      IntegreHistoSaisPrim(Tob_Rub, Tob_PSP, DateD, DateF, TheSal, TheEtab, 'M');
      Tob_PSD := Tob.create('Historique saisie Rub', nil, -1);
      if RecupereNouvellesHistoSaisRub(Tob_PSD, DateD, Datef, TheSal, 'BCP') then
        IntegreHistoSaisrub(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'M');
    end;
    exit;
  end;
  // FIN PT40
  // FIN PT88
  CalculJoursTheorique(Tob_sal, DateD, DateF, JTHTRAVAILLES, HTHTRAVAILLES); //PT30-1 Calcul de la variable 0039
  ChpEntete.HeuresTrav := HTHTRAVAILLES; //PT71-4 en chargement on initialise
  ChpEntete.CpAcq := 0; { PT87 }
  ChpEntete.CpSupl := 0; { PT87 }
  ChpEntete.CpAnc := 0; { PT87 }
  if Auto then GblCP := GCPAuto; { PT127-4 }

  ActionSld := taCreation;
  suivant := IncrementeSeqNoOrdre('CPA', TheSal);
  TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TheEtab], True);
  if TETABSAL = nil then HShowMessage('2;Erreur sur la tob Etablissement;W;O;O;;;', '', ''); //PT12  Redondence de code


  PrefaceExecuteCP(TETABSAL, Auto, StMsgErr, TheSal);


  Subrogation := (TETABSAL.GetValue('ETB_SUBROGATION') = 'X'); // PT119

  if (ActionB = taCreation) or (ACtionB = PremCreation) then
  begin //B3
    {*****  Intégration des absences *****}
    { Tob_ABS := Tob.create('ABSENCES SALARIE',nil,-1);  PT46 Mise en commentaire : integrer dans RecupereAbsences
    PT46   HABSPRIS:=Tob_ABS.Somme('PCN_HEURES',[''],[''],False); //PT25 Affectation Var Heures Absence pris
    PT46   JABSPRIS:=Tob_ABS.Somme('PCN_JOURS',[''],[''],False);  //PT25 Affectation Var Jours Absence pris
   if RecupereAbsences(Tob_ABS, Datef, TheSal, 'C') then}
{PT220
    if (VH_PAIE.PGAbsence) and (RecupereAbsences(Tob_ABS, Datef, TheSal, 'C', GenereMaintien, IJSS)) then} { PT132 Ajout Paramsoc PT111}
    if (VH_PAIE.PGAbsence) and (RecupereAbsences(Tob_ABS, Datef, TheSal, 'C', GenereMaintien, IJSS,'')) then { PT132 Ajout Paramsoc PT111}
      IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, 'C');

{$IFNDEF CCS3} // d PT119
    // On gère uniquement les IJSS
    if ((not VH_Paie.Pgmaintien) or (not GenereMaintien)) and (VH_Paie.PgGestIJSS) then //PT136
    begin
      RecupIJSS(Dated, Datef, 'C', Tob_Rub, TOB_IJ, Tob_Etablissement, Tob_Salarie, ChampCateg, RubNettes, Subrogation);
      TOB_IJSS := TOB_IJ;
    end;
{$ENDIF} // f PT119

    {***** Gestion des formations HTT *****}
    {si formation ... }
    { si Formation HTT et Si Rub existe dans Table Histosaisrub pour la période }
    {  alors Récup profil associé Parametrage Formation HTT.}

    {if Profil <> '' then ChargeProfil (SAlarie, Tob_RUb, Profil);}

    {***** Intégration des acomptes *****}
    Tob_PSD := Tob.create('Historique saisie Rub', nil, -1);
    RecupereHistoSaisRub(Tob_PSD, DateD, Datef, TheSal, (SortieSal) or (EntreeSal)); // PT157 // PT172
    IntegreHistoSaisrub(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'C');
    if VH_Paie.PGGESTACPT then { PT150 }
      IntegreHistoSaisrubACP(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'C');

    {*****  Intégration des CP *****}
    IntegreCPCreate(TETABSAL, TOB_RUB, TOB_Sal, Auto, StMsgErr);
  end //E3
  else
  begin // modification de bulletin B6
     {*****  Intégration des CP *****}
    IntegreCPModif(TETABSAL, TOB_RUB, TOB_Sal, Auto, StMsgErr);

    {***** Intégration des absences en modif de bulletin *****}
    {DEB PT25 extraction du code de la condition afin de recharger les abs. en modif.}
{PT220
    if (VH_PAIE.PGAbsence) and (RecupereAbsences(Tob_ABS, Datef, TheSal, 'M', GenereMaintien, IJSS)) then} { PT132 Ajout Paramsoc PT111}
    if (VH_PAIE.PGAbsence) and (RecupereAbsences(Tob_ABS, Datef, TheSal, 'M', GenereMaintien, IJSS, '')) then { PT132 Ajout Paramsoc PT111}
      IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, 'M');
    {PT46 Mise en commentaire puis supprimé }

{$IFNDEF CCS3} // d PT119
    // On gère uniquement les IJSS
    if ((not VH_Paie.Pgmaintien) or (not GenereMaintien)) and (VH_Paie.PgGestIJSS) then { PT136 }
    begin
      RecupIJSS(Dated, Datef, 'M', Tob_Rub, TOB_IJ, Tob_Etablissement, Tob_Salarie, ChampCateg, RubNettes, Subrogation);
      TOB_IJSS := TOB_IJ;
    end;
{$ENDIF} // f PT119

    {***** Intégration des acomptes *****}
    Tob_PSD := Tob.create('Historique saisie Rub', nil, -1);
    if RecupereNouvellesHistoSaisRub(Tob_PSD, DateD, Datef, TheSal) then
    begin
      IntegreHistoSaisrub(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'M');
      if VH_Paie.PGGESTACPT then { PT150 }
        IntegreHistoSaisrubACP(Tob_Rub, Tob_PSD, DateD, DateF, TheSal, TheEtab, 'M');
    end;

  end; //E6

  {*****  Intégration automatique des saisies arrêt *****}
  IntegreRubSaisiesArret(TOB_Rub, Tob_Sal, DateD, DateF); //PT209

  ChargeRegulCot(TOB_Sal, Tob_Rub, DateD, DateF); // PT180
end; //E1

procedure PrefaceExecuteCP(TETABSAL: Tob; Auto: Boolean; var StMsgErr: string; TheSal: string);
var
  LastBull: boolean; //PT57 ajout LastBull
  Q: TQuery;
begin

  GblCP := (VH_Paie.PGCongesPayes) and (TETABSAL.GetValue('ETB_CONGESPAYES') = 'X') and (TOB_SALARIE.GetValeur(IPSA_CONGESPAYES) = 'X'); { PT89 } { PT127-4 }

  {DEB PT57 ajout requête et message}
  if not GblCP then { DEB PT89 09/06/2004 } { PT127-4 }
  begin
    if X <> nil then
    begin
      X.CBModifAcquis.Enabled := False;
      X.CPMOIS.Enabled := False;
      X.CpSupp.Enabled := False;
      X.CpAnc.Enabled := False;
    end;
  end { FIN PT89 }
  else
    if (GblCP) then { PT89 } { PT92-1 17/05/2004 modif clause } { PT127-4 }
    begin
    { DEB PT92-1 }
      LastBull := True;
      FirstBull := True;
      Q := OpenSql('SELECT MIN(PPU_DATEDEBUT) MINDATE ,MAX(PPU_DATEDEBUT) MAXDATE FROM PAIEENCOURS ' +
        'WHERE PPU_SALARIE = "' + TheSal + '" AND PPU_BULCOMPL="-"', True); //PT69-1 Ajout clause AND PPU_DATEDEBUT>"' + UsDateTime(DateD) + '"
      if not Q.eof then
      begin
        if (Q.FindField('MINDATE').AsDateTime > idate1900) and (Q.FindField('MINDATE').AsDateTime <> DateD) then FirstBull := False;
        if (Q.FindField('MAXDATE').AsDateTime > idate1900) and (Q.FindField('MAXDATE').AsDateTime > DateD) then LastBull := False;
      end;
      Ferme(Q);
    { FIN PT92-1 }
      if not LastBull then
      begin
      { DEB PT90 }
        if auto then
        // PT102
          StMsgErr := 'Bulletin postérieur existant.la validation du bulletin a pu entraîner des décalages congés payés pour le salarié ' + Tob_Salarie.GetValeur(iPSA_SALARIE) + '.'
        else
          PgiInfo('Vous avez déjà généré des paies postérieures, #13#10' +
            'la validation du bulletin peut entrainer des décalages de compteurs congés payés.', 'Bulletin posterieur existant');
        if X <> nil then
        begin
          X.CBModifAcquis.Enabled := False;
          X.CPMOIS.Enabled := False;
          X.CpSupp.Enabled := False;
          X.CpAnc.Enabled := False;
        end;
      end;
    { FIN PT90 }
    end
    else LastBull := True; //Tjrs dernier Bull pour un solde
  {FIN PT57}
end;


procedure IntegreCPCreate(TETABSAL, TOB_RUB, TOB_Sal: Tob; Auto: Boolean; var StMsgErr: string);
var
//PT210  base, nbmois, Mois, Supp, Anc: double;
  base, nbmois, Mois, Supp, Anc, Fract: double;
begin
  if ((VH_paie.PGCongesPayes) and (GblCP) { and (LastBull) PT90 }) then {PT57 ajout LastBull} { PT127-4 }
  begin //B5
    Date1 := TOB_Sal.GetValeur(iPSA_DATESORTIE);
    if ((Date1 >= DateD) and (Date1 <= DateF)) then
      VerificationCP(Tob_Sal.GetValeur(iPSA_SALARIE), Tob_Sal.GetValeur(iPSA_ETABLISSEMENT), True); { PT148-2 }
    Mois := 0;
    Supp := 0;
    Anc := 0; //  PT12 Boolean à true pour générer les acqs au chargement de la fiche
      { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
      { JCPACQUIS:=AlimCongesAcquis(TOB_Salarie,TETABSAL,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,True,base,nbmois,suivant,Mois,Supp,Anc); // Génération cp acquis }
//PT210    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc);
    JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc, Fract);
    InitTobAcquisCp(TOB_Salarie, TETABSAL, DateD, DateF, base, Nbmois, Mois, Supp, Anc, suivant);
    if X <> nil then
    begin
      X.CPMOIS.text := floattostr(Mois);
      X.CpSupp.text := floattostr(Supp);
      X.CpAnc.text := floattostr(Anc);
    end;
    MCPACQUIS := nbmois;
    BCPACQUIS := base;
    Tob_Pris := SalIntegreCP(TOB_Sal, TOB_Rub, T_MvtAcquis, DateD, DateF, Auto, StMsgErr); //PT12 passage en parametre de la tob des acqs en cours
  end; //E5
end;

procedure IntegreCpModif(TETABSAL, TOB_RUB, TOB_Sal: Tob; Auto: Boolean; var StMsgErr: string);
var
  Tob_TestPris, Tob_testSolde: TOB;
  ZDate: TDateTime;
  TheSal, TheEtab: string;
//PT210  base, nbmois, Mois, Supp, Anc : double;
  base, nbmois, Mois, Supp, Anc, Fract: double;
  ModifAcq, New: Boolean;
begin
  TheSal := TOB_SALARIE.GetValeur(iPSA_SALARIE);
  TheEtab := TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT);
  if ((VH_paie.PGCongesPayes) and (GblCP) {and (LastBull) PT90 }) then {PT57 ajout LastBull} { PT127-4 }
  begin //B7
      // on recalcule les congés seulement si il y a eu de nvx mvts créés
    Tob_TestPris := Tob.create('CONGES PRIS', nil, -1);
    Tob_testSolde := Tob.create('CONGES SOLDE', nil, -1);
    ZDate := TOB_Sal.GetValeur(iPSA_DATESORTIE);
    RecupereCongesPrisPayes(Tob_TestPris, Tob_testSolde, Datef, TheSal, 'SLD');
    if not ((Assigned(tob_testsolde) and (tob_testsolde.detail.count > 0)) {(notvide(tob_testsolde, true) } { PT167 }
      and (ZDate <= Datef) and (Zdate >= DateD)) then
    begin // B71
      Tob_testPris.free;
      Tob_testSolde.free;
      Tob_TestPris := Tob.create('CONGES PRIS', nil, -1);
      if MCPACQUIS = 0 then
        if ChpEntete.DTrent > 0 then
          MCPACQUIS := ChpEntete.NTrent / ChpEntete.DTrent else NbMois := 0;
        //DEB PT12 On alimente la tob des acqs en cours
        { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
        { AlimCongesAcquis(TOB_Salarie,TETABSAL,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,False,base,nbmois,suivant,Mois,Supp,Anc); }
//PT210      ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc);
      ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc, Fract);
        { PT38-1 Si mvt acq est topé C ou S, on ne recharge pas les CP pour ne pas fausser le décompte }
        { Cas bulletin précedent un bulletin SLD ou bulletin en N-2 }
//PT210      SoldeOuClot := ChargeT_MvtAcquis(TheSal, DateD, DateF, Mois, Supp, Anc, ModifAcq, New);
      SoldeOuClot := ChargeT_MvtAcquis(TheSal, DateD, DateF, Mois, Supp, Anc, Fract, ModifAcq, New);
      if (New) then
      begin
          { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
          { JCPACQUIS:=AlimCongesAcquis(TOB_Salarie,TETABSAL,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,True,base,nbmois,suivant,Mois,Supp,Anc); }
//PT210        JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc);
        JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc, Fract);
        InitTobAcquisCp(TOB_Salarie, TETABSAL, DateD, DateF, base, Nbmois, Mois, Supp, Anc, suivant);
      end;
      if X <> nil then
      begin
        X.CpMois.text := floattostr(Mois);
        X.CpSupp.text := floattostr(Supp);
        X.CpAnc.text := floattostr(Anc);
      end;
      MCPACQUIS := nbmois;
      BCPACQUIS := base;
//PT210      JCPACQUIS := Mois + Supp + Anc; //PT26
      JCPACQUIS := Mois + Supp + Anc + Fract;
        //FIN PT12
      if (RecupereCongespris(Tob_TestPris, Tob_Sal, DateF)) or (ModifAcq = True) then
      begin //B8
        AnnuleCongesPris(TheSal, TheEtab, DateD, DateF);
        Tob_Pris := SalIntegreCP(TOB_Sal, TOB_Rub, T_MvtAcquis, DateD, DateF, Auto, StMsgErr); //PT12 passage en parametre de la tob des acqs en cours
          //          ActionSld:=taCreation;
      end //E8
      else
        RechargePrisEtSoldePayes(Tob_rub, tob_sal);
      Tob_testPris.free;
    end
    else //E71
        //on a déjà payé le STC, on se contente de recharger les tob Tob_pris et tob_solde
    begin //DEB PT26
        { PT87 Modification de la fonction d'appel des acquis, séparation du chargement de la mise en tob }
        { JCPACQUIS:=AlimCongesAcquis(TOB_Salarie,TETABSAL,DateD,DateF,ChpEntete.NTrent,ChpEntete.DTrent,TOB_RUB,False,base,nbmois,suivant,Mois,Supp,Anc); }
//PT210      JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc);
      JCPACQUIS := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc, Fract);
      MCPACQUIS := nbmois;
      BCPACQUIS := base;
        //FIN PT26
      RechargePrisEtSoldePayes(Tob_rub, tob_sal);
      if (X <> nil) then
      begin
        if (Zdate >= DateD) and (ZDate <= Datef) then
        begin
          X.CBModifAcquis.enabled := False;
          X.CPMOIS.Enabled := False; { PT87 13/05/2004 }
          X.CpSupp.Enabled := False; { PT87 13/05/2004 }
          X.CpAnc.Enabled := False; { PT87 13/05/2004 }
        end;
        X.CpMois.text := floattostr(Mois); { PT87 }
        X.CpSupp.text := floattostr(Supp);
        X.CpAnc.text := floattostr(Anc);
      end;
    end;
      //DEB PT38-1 Gestion message d'erreur pour des mvts acquis soldés ou cloturés
    if SoldeOuClot then
    begin
      if Auto then //PT60 Modification message
        StMsgErr := 'Les mouvements de congés payés du salarié ' + TheSal + ' ' + TOB_Sal.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Sal.GetValeur(iPSA_PRENOM) +
          ' de cette période sont soldés ou cloturés.Aucune modification ne sera sauvegardée.'
      else
        HShowMessage('5;Bulletin de paie du salarié : ' + TheSal + ' ' + TOB_Sal.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Sal.GetValeur(iPSA_PRENOM) +
          ' :;Attention les mouvements de congés payés de cette période sont soldés ou cloturés.#13#10Aucune modification ne sera sauvegardée;E;O;O;O;;;', '', '');
    end;
      //FIN PT38-1 Gestion message d'erreur
  end; //E7
end;


procedure IntegreCPConsult(TOB_RUB: Tob);
var
  TETABSAL: Tob;
  TheSal, TheEtab: string;
//PT210  base, Nbmois, Mois, Supp, Anc: double;
  base, Nbmois, Mois, Supp, Anc, Fract : double;
begin
  TheEtab := TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT);
  TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TheEtab], True);
//PT210  ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc);
  ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, DateD, DateF, True, True, base, Nbmois, Mois, Supp, Anc, Fract);
  if X <> nil then
  begin
    X.CBModifAcquis.Enabled := False;
    X.CPMOIS.Enabled := False;
    X.CpSupp.Enabled := False;
    X.CpAnc.Enabled := False;
    X.CPMOIS.text := floattostr(Mois);
    X.CpSupp.text := floattostr(Supp);
    X.CpAnc.text := floattostr(Anc);
  end;
end;


procedure RechargePrisEtSoldePayes(Tob_rub, Tob_sal: tob);
begin
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl = 'X') then exit;

  Tob_Pris := Tob.create('CONGES PRIS', nil, -1);
  RecupereCongesPrisPayes(Tob_Pris, Tob_solde, Datef, Tob_sal.GetValeur(iPSA_SALARIE), 'PRI');
  GestionCalculAbsence(tob_rub, Dated, datef, 'PRI'); //ajout MVGR
  if Assigned(Tob_Pris) and (Tob_Pris.detail.count > 0) then // if notVide(Tob_Pris, true) then { PT167 }
    IntegrePaye(Tob_Rub, Tob_Pris, Etab, CodeSalarie, DateD, DateF, 'PRI');
  Tob_Solde := Tob.create('SOLDE CONGES', nil, -1);
  RecupereCongesPrisPayes(Tob_Pris, Tob_solde, Datef, Tob_sal.GetValeur(iPSA_SALARIE), 'SLD');
  GestionCalculAbsence(tob_rub, Dated, datef, 'SLD'); //ajout MVGR
  if Assigned(Tob_solde) and (Tob_solde.detail.count > 0) then //if notVide(Tob_solde, true) then { PT167 }
    IntegrePaye(Tob_Rub, Tob_Solde, Etab, CodeSalarie, DateD, DateF, 'SLD');
end;

procedure TFSaisBul.IndemcpClick(Sender: TObject);
var
  rep, Act: string; {PT1 Ajout Act}
  t: tob;
  // mv le 14/11 replace Tob_pris par Tob_pris
begin
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl = 'X') then
  begin
    PGIBox('Vous ne pouvez pas visualiser les indemnités congés payés '
      + '#13#10 dans un bulletin complémentaire ', 'Saisie des bulletins');
    exit;
  end;
  {DEB PT1}
  if ActionBul = TaConsultation then
    Act := 'Consultation;PRI' { PT127-3 Ajout PRI }
  else
    Act := ';PRI'; { PT127-3 Ajout PRI }
  {FIN PT1}
  if (Tob_Pris <> nil) and (Tob_Pris.detail.count > 0) then { PT127-3 }
  begin
    t := Tob_pris.findfirst([''], [''], false);
    if t <> nil then
    begin
{$IFNDEF EAGLSERVER}
      TheTob := Tob_Pris;
      GrilleDansRub;
      Rep := AglLanceFiche('PAY', 'INDEMNITECP', '', '', Act); {PT1 Ajout Act}
      thetob := nil;
{$ENDIF}
      if Rep = '0' then
      begin
        AgregeCumulPris(Tob_Pris);
        IntegrePaye(Tob_Rub, Tob_Pris, Etab, CodeSalarie, DateD, DateF, 'PRI');
        RubDansGrille;
        //   if VH_Paie.PGCalculBulletin then
        BCalculBullClick(nil);
      end;
    end;
  end
  else { PT127-3 Ajout message }
    PgiBox('Vous n''avez pas d''indemnités congés payés à visualiser!', 'Congés payés');

end;

procedure TFSaisBul.IndemSoldeClick(Sender: TObject);
var
  rep, Act: string; {PT1 Ajout Act}
  t: tob;
begin
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl = 'X') then
  begin
    PGIBox('Vous ne pouvez pas visualiser le solde congés payés '
      + '#13#10 dans un bulletin complémentaire ', 'Saisie des bulletins');
    exit;
  end;
  {DEB PT1}
  if ActionBul = TaConsultation then
    Act := 'Consultation;SLD' { PT127-3 Ajout SLD }
  else
    Act := ';SLD'; { PT127-3 Ajout SLD }
  {FIN PT1}
  if (Tob_Solde <> nil) and (Tob_Solde.detail.count > 0) then { PT127-3 Ajout detail.count }
  begin
    t := Tob_solde.findfirst([''], [''], false);
    if t <> nil then
    begin
{$IFNDEF EAGLSERVER}
      TheTob := Tob_Solde;
      GrilleDansRub;
      Rep := AglLanceFiche('PAY', 'INDEMNITECP', '', '', Act); {PT1 Ajout Act}
      thetob := nil;
{$ENDIF}
      if Rep = '0' then
      begin
        AgregeCumulPris(Tob_Solde);
        IntegrePaye(Tob_Rub, Tob_Solde, Etab, CodeSalarie, DateD, DateF, 'SLD');
        RubDansGrille;
        //   if VH_Paie.PGCalculBulletin then
        BCalculBullClick(nil);
      end;
    end;
  end
  else { PT127-3 Ajout message }
    PgiBox('Vous n''avez pas d''indemnités compensatrice congés payés à visualiser!', 'Congés payés');

end;


procedure TFSaisBul.ValoAbsenceClick(Sender: TObject);
var
  rep, Act: string; {PT1 Ajout Act}
  t: tob;
  // mv le 14/11 replace Tob_pris par Tob_pris
begin
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl = 'X') then
  begin
    PGIBox('Vous ne pouvez pas visualiser les absences '
      + '#13#10 dans un bulletin complémentaire ', 'Saisie des bulletins');
    exit;
  end;
  {DEB PT1}
  if ActionBul = TaConsultation then
    Act := 'Consultation'
  else
    Act := '';
  {FIN PT1}
  if (Tob_Pris <> nil) and (Tob_Pris.detail.count > 0) then { PT127-3 Ajout detail.count }
  begin
    t := Tob_pris.findfirst([''], [''], false);
    if t <> nil then
    begin
{$IFNDEF EAGLSERVER}
      TheTob := Tob_Pris;
      GrilleDansRub;
      Rep := AglLanceFiche('PAY', 'VALOABSENCE', '', '', Act); {PT1 Ajout Act}
      thetob := nil;
{$ENDIF}
      if Rep = '0' then
      begin
        AgregeCumulPris(Tob_Pris);
        IntegrePaye(Tob_Rub, Tob_Pris, Etab, CodeSalarie, DateD, DateF, 'PRI');
        RubDansGrille;
        //if VH_Paie.PGCalculBulletin then
        BCalculBullClick(nil) {else
        // sinon on trie qd même
           begin
           Tob_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE') ;
           RubDansGrille;
           end;                   }
      end;
    end;
  end
  else { PT127-3 Ajout message }
    PgiBox('Vous n''avez pas de valorisation absence congés payés à visualiser!', 'Congés payés');

end;

procedure TFSaisBul.CongesPayesClick(Sender: TObject);
var
  DateEntree, Zdate: tdatetime;
  Init: word;
  St: string;
  Q: TQuery;
begin
  //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
  if (BullCompl = 'X') then
  begin
    PGIBox('Vous ne pouvez pas visualiser les congés payés '
      + '#13#10 dans un bulletin complémentaire ', 'Saisie des bulletins');
    exit;
  end;

  if not ((VH_paie.PGCongesPayes) and (GblCP)) then {Deb1 PT5} { PT127-4 }
  begin
    PgiBox('Vous ne gérez pas les congés payés : Aucune consultation n''est possible.', 'Visualisation des congés payés');
    exit;
  end; {Fin1 PT5}
  {DEB PT1}
  if ActionBul = TaConsultation then
  begin
    PgiBox('Vous ne pouvez consulter les congés payés : la paie est cloturée.', 'Visualisation des congés payés');
    exit;
  end;
  {FIN PT1}
  ZDate := TOB_Sal.GetValeur(iPSA_DATESORTIE);
  if (ZDate <= Datef) and (Zdate >= DateD) and (ActionBul = tamodification) then
  begin
    { DEB PT91 }
    if (PGRecupereProfilCge(TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)) = '') then
      HShowMessage('1;Attention;Le mouvement de solde n''a pas été généré.#13#10' +
        ' Vous devez gérer un profil congés payés pour le salarié ' + CodeSalarie + '.#13#10' +
        ' Si vous souhaitez recalculer le mouvement de solde, supprimez et recréez votre bulletin;W;Y;Y;N;', '', '')
    else
      { FIN PT91 }
      HShowMessage('1;Attention;Le mouvement de solde a déjà été généré.#13#10' +
        ' Si vous réalisez des modifications de mouvements CP, ils ne seront pas pris en compte dans le bulletin.#13#10' +
        ' Si vous souhaitez recalculer le mouvement de solde, supprimez et recréez votre bulletin;W;Y;Y;N;', '', '');
{$IFNDEF EAGLSERVER}
    AglLanceFiche('PAY', 'CONGESPAY_MUL', '', CodeSalarie, CodeSalarie + ';C;'';ORDREG;BULL');
{$ENDIF}
    //PCN_PERIODECP=0;PCN_TYPECONGE=PRI +Tob_sal.GetValeur (iPSA_ETABLISSEMENT)+
    exit;
  end;
  DateEntree := Now;
  GrilleDansRub;
  if ActionBul <> taCreation then //TaConsultation
  begin
    St := 'SELECT PPU_DATEDEBUT FROM PAIEENCOURS ' + {Deb2 PT5}
      'WHERE PPU_SALARIE = "' + CodeSalarie + '" AND PPU_DATEDEBUT>"' + UsDateTime(DateD) + '"';
    Q := OpenSql(st, True);
    if Q.eof then
    begin {Fin2 PT5}
      Init := HShowMessage('1;Congés payés;Attention, la modification d''un enregistrement congés payés concernant ce bulletin#13#10' +
        'entrainera le recalcul du bulletin et la perte de votre éventuelle personalisation sur vos calcul congés payés.#13#10Etes vous sûr de vouloir continuer ?;Q;YN;N;N;', '',
        '');
      if (Init = mrNo) or (Init = MrCancel) then exit;
      AnnuleCongesPris(Tob_sal.GetValeur(iPSA_SALARIE), Tob_sal.GetValeur(iPSA_ETABLISSEMENT), DateD, DateF);
      {AglLanceFiche ('PAY','CONGESPAY_MUL', 'PCN_PERIODECP=0;PCN_TYPECONGE=PRI',
       Tob_sal.GetValeur (iPSA_SALARIE) , Tob_sal.GetValeur (iPSA_SALARIE)+';SAISBUL' );}
{$IFNDEF EAGLSERVER}
      AglLanceFiche('PAY', 'CONGESPAY_MUL', 'PCN_PERIODECP=0;PCN_TYPECONGE=PRI',
        CodeSalarie, CodeSalarie + ';C;'';ORDREG;BULL');
{$ENDIF}
      if not Assigned(GblTob_VarCp) then FreeAndNil(GblTob_VarCp); { PT171 }

      if not ModificationCongesPris(Tob_pris, Tob_Sal, DateF, Dateentree) then exit;
      RecalculCp(True); //PT12  RecalculCp permet de réalimenter les variables globales des acquis en cours : JCPACQUIS
      {DEB PT21 Mise en commentaire Integré ds RecalculCp
       Tob_Pris := SalIntegreCP(TOB_Sal, TOB_Rub,T_MvtAcquis, DateD,DateF,False,StTemp);//PT12 passage en parametre de la tob des acqs en cours
       RubDansGrille;
       CalculBullSaisie; FIN PT21 }
      OnSauve;
      Modifier := TRUE;
    end
    else
    begin {Deb3 PT5}
      Init := HShowMessage('1;Congés payés;Vous ne pouvez pas modifier les enregistrements congés payés concernant ce bulletin : #13#10' +
        'Aucune(s) saisie(s) ne sera pris en compte.#13#10Etes vous sûr de vouloir continuer ?;Q;YN;N;N;', '', '');
      if (Init = mrNo) or (Init = MrCancel) then exit;
{$IFNDEF EAGLSERVER}
      AglLanceFiche('PAY', 'CONGESPAY_MUL', '', CodeSalarie, CodeSalarie + ';C;'';ORDREG;BULL');
{$ENDIF}
      //PCN_PERIODECP=0;PCN_TYPECONGE=PRI                {Fin3 PT5}
    end;
    Ferme(Q);
  end
  else PgiBox('Accès impossible en création de bulletin, veuillez valider votre bulletin', 'Saisie des bulletins'); // PT108
end;

procedure TFSaisBul.BRechargContextClick(Sender: TObject);
begin
  VideLesTOBPaie(TRUE);
  ChargeLesTOBPaie;
end;


procedure TFSaisBul.PaieEnversClick(Sender: TObject);
var
  rep: string;
begin
{$IFNDEF EAGLSERVER}
  BCalculBullClick(Sender);
  TheTob := TOB_Rub;
  rep := AglLanceFiche('PAY', 'PAIEENVERS', '', '', '');
  if rep = 'O' then
  begin
    RubDansGrille;
    AfficheBasBulletin;
  end;
  TheTob := nil;
{$ENDIF}
end;
// DEB PT204

procedure TFSaisBul.PaieEnversRubClick(Sender: TObject);
var
  rep: string;
begin
{$IFNDEF EAGLSERVER}
  BCalculBullClick(Sender);
  TheTob := TOB_Rub;
  rep := AglLanceFiche('PAY', 'PAIEENVERS', '', '', 'RUB');
  if rep = 'O' then
  begin
    RubDansGrille;
    AfficheBasBulletin;
  end;
  TheTob := nil;
{$ENDIF}
end;
// FIN PT204

procedure TFSaisBul.BtnOrigineClick(Sender: TObject);
var
  st, Ret, Rubriq, Nat: string;
begin
  Nat := ''; // PT112
  if CurG = GCot then
    st := CurG.Cells[7, CurG.Row] // Grille Cotisations
  else
    if CurG = GBas then
    //PT68   05/06/2003 V_421 PH FQ 10700 Gestion origine de la rubrique dans les bases de cotisations
      st := CurG.Cells[10, CurG.Row] // Grille Bases de Cotisations
    else
    begin
      st := CurG.Cells[6, CurG.Row]; // Grille Remunerations
      Nat := 'REM' // PT112
    end;
  St := AnsiUpperCase(Copy(St, 1, 3));
  Rubriq := CurG.Cells[0, CurG.Row];
  // PT14 : 27/11/2001 V562 PH Propagation d'une ligne de commentaire
//  if (RechCommentaire(Rubriq) = TRUE) then st := 'BUL';   // PT178

  if St = '' then exit;
  {if Nat = 'REM' then PT175 }st := St + ';' + Rubriq + ';' + Nat; // PT 112
{$IFNDEF EAGLSERVER}
  ret := AglLanceFiche('PAY', 'ORIGINE_RUBRIQ', '', '', st);
{$ENDIF}
  if ret <> '' then
  begin
    if CurG = GCot then CurG.Cells[7, CurG.Row] := RechDom('PGACTIONBULLETIN', ret, FALSE)
    else
      if CurG = GBas then CurG.Cells[10, CurG.Row] := RechDom('PGACTIONBULLETIN', ret, FALSE)
      else
        CurG.Cells[6, CurG.Row] := RechDom('PGACTIONBULLETIN', ret, FALSE);
  end;
end;

procedure TFSaisBul.RemPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then exit;
  if ActionLigne = taCreation then exit;
  if CurG.Cells[0, ARow] = '' then exit;
  if RemZoneGrise(ACol, ARow) then
  begin
    GridGriseCell(CurG, Acol, Arow, Canvas);
//    Canvas.Font.Color := clyellow;
  end;
end;

function TFSaisBul.RemZoneGrise(ACol, ARow: Integer): Boolean;
var
  T1, LaLigne: TOB;
  cas: WORD;
  Rubriq, AccesSaisie: string;
begin
//  result := FALSE;
  // Colonne 0 interdite sauf en creation
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := TRUE;
  end;
  // colonne 1 libelle tjrs accessible
  if Acol = 1 then
  begin
    result := FALSE;
    exit;
  end;
  if Acol > SG_Mt then
  begin
    result := TRUE;
    exit;
  end; // Derniere Colonne jamais Saisie
  Rubriq := CurG.Cells[0, ARow];
  if (RechCommentaire(Rubriq) = TRUE) then
  begin
    if ACol > SG_Lib then
    begin
      Result := TRUE;
      exit;
    end;
  end;
  result := TRUE;
  // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
  if (MyTob = nil) or (MyTob.NomTable <> 'REMUNERATION') or ((MyTob.getvalue('PRM_RUBRIQUE') <> Rubriq)) then
  begin
    // Optimisation T1 := TOB_Rem.FindFirst (['PRM_RUBRIQUE'], [Rubriq], TRUE);
    T1 := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', rubriq);
    MyTob := T1;
  end
  else T1 := MyTob;

  if T1 <> nil then
  begin
    if not ExamCasValeurRem(ACol, T1) then // Test sur le type du champ dans le cas valeur alors on ne saisit jamais
    begin
      cas := TypeFonctRem(T1); // recup du type de saisie donc renseigne en partie le fonctionnement de la rémunération
      if CurG.Objects[0, ARow] = nil then
        exit;

      if CurG.Objects[0, ARow] = nil then
        LaLigne := nil
      else
        LaLigne := TOB(CurG.Objects[0, ARow]); // Recup de la TOB TOB_Rub associee a chaque ligne
      if (Laligne <> nil) then // beurk@@@@@@@@
        AccesSaisie := CurG.Cells[SG_Origine, ARow] // LaLigne.GetValue ('PHB_ORIGINELIGNE') // recup du champ origine de la ligne pour autoriser la saisie
      else AccesSaisie := '';
      if cas = 1 then // saisie du montant
      begin
        if (ACol = SG_Mt) or (ACol = SG_Base) then result := FALSE;
//        if (AccesSaisie <> '') then result := TRUE;
      end;
      if cas = 2 then // saisie de la Base et du Taux
      begin
        if (ACol = SG_Base) or (ACol = SG_Taux) then result := FALSE;
//        if (AccesSaisie <> '') then result := TRUE;
      end;
      if cas = 3 then // saisie de la Base du taux et du Coeff
      begin
        if (ACol = SG_Base) or (ACol = SG_Taux) or (ACol = SG_Coeff) then result := FALSE;
//        if (AccesSaisie <> '') then result := TRUE;
      end;
      if cas = 4 then // saisie de la base et du coeff
      begin
        if (ACol = SG_Base) or (ACol = SG_Coeff) then result := FALSE;
//        if (AccesSaisie <> '') then result := TRUE;
      end;
    end; // Cas ExamineCasValeur
  end; // TOb non nulle

end;

procedure TFSaisBul.BasPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then exit;
  if ActionLigne = taCreation then exit;
  if CurG.Cells[0, ARow] = '' then exit;
  if BaseZoneGrisee(ACol, ARow) then GridGriseCell(CurG, Acol, Arow, Canvas);
end;

function TFSaisBul.BaseZoneGrisee(ACol, ARow: Longint): Boolean;
var
  T1: TOB;
begin
  result := TRUE;
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := FALSE;
    exit;
  end;
  T1 := TOB(CurG.Objects[0, ARow]);
  if T1 = nil then
  begin
    result := TRUE;
    exit;
  end;
  if (T1 <> nil) then
  begin
    if ACol = SG_LIB then result := FALSE; // Libelle tjrs accessible
    if (ACol = SG_Base) and (ChbxBaseForcee.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Plfd) and (ChbxBaseForcee.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Tr1) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Tr2) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Tr3) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Plf1) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Plf2) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
    if (ACol = SG_Plf3) and (ChbxTranchesForcees.Checked = TRUE) then result := FALSE;
  end;
end;

procedure TFSaisBul.CotPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow = 0 then exit;
  if ActionLigne = taCreation then exit;
  if CurG.Cells[0, ARow] = '' then exit;
  if CotZoneGrisee(ACol, ARow) then GridGriseCell(CurG, Acol, Arow, Canvas);
end;

function TFSaisBul.CotZoneGrisee(ACol, ARow: Longint): Boolean;
var
  T1: TOB;
  Rubriq: string;
begin
  result := TRUE;
  if (Acol = 0) and (ActionLigne <> taCreation) then
  begin
    result := TRUE;
    exit;
  end;
  if Acol = 1 then
  begin
    result := FALSE;
    exit;
  end;
  T1 := TOB(CurG.Objects[0, ARow]);
  if T1 = nil then
  begin
    result := TRUE;
    exit;
  end;

  // PT16 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
  Rubriq := CurG.Cells[0, ARow];
  if (RechCommentaire(Rubriq) = TRUE) then
  begin
    if ACol > SG_Lib then
    begin
      Result := TRUE;
      exit;
    end;
  end;
  // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
  if (MyTob = nil) or (MyTob.NomTable <> 'COTISATION') or ((MyTob.getvalue('PCT_RUBRIQUE') <> Rubriq) and (MyTob.getvalue('PCT_NATURERUB') <> 'COT')) then
  begin
    // on lit la TOB concernant la  ligne concernée
    // T1 := TOB_Cotisations.FindFirst (['PCT_RUBRIQUE','PCT_NATURERUB'], [T1.GetValue ('PHB_RUBRIQUE'),'COT'], TRUE);
    T1 := Paie_RechercheOptimise(tob_cotisations, 'PCT_RUBRIQUE', rubriq);
    MyTob := T1;
  end
  else T1 := MyTob;

  if (T1 <> nil) then
  begin
    if ACol = SG_LIB then result := TRUE; // Libelle tjrs accessible
    // PT29 : 04/04/2002 V571 PH Acces saisie des colonnes des cotisations si element variable
    if (ACol = SG_TxSal) and ((T1.GetValue('PCT_TYPETAUXSAL') = 'ELP') or (T1.GetValue('PCT_TYPETAUXSAL') = 'ELV'))
      then result := FALSE;
    if (ACol = SG_TxPat) and ((T1.GetValue('PCT_TYPETAUXPAT') = 'ELP') or (T1.GetValue('PCT_TYPETAUXPAT') = 'ELV'))
      then result := FALSE;
    if (ACol = SG_MtSal) and ((T1.GetValue('PCT_TYPEFFSAL') = 'ELP') or (T1.GetValue('PCT_TYPEFFSAL') = 'ELV'))
      then result := FALSE;
    if (ACol = SG_MtPat) and ((T1.GetValue('PCT_TYPEFFPAT') = 'ELP') or (T1.GetValue('PCT_TYPEFFPAT') = 'ELV'))
      then result := FALSE;
    if (ACol = SG_Base) and ((T1.GetValue('PCT_TYPEBASE') = 'ELP') or (T1.GetValue('PCT_TYPEBASE') = 'ELV'))
      then result := FALSE;
  end;
end;

procedure TFSaisBul.ZChbxTranchesForceesClick(Sender: TObject);
var
  i, j: Integer;
begin
  if CurG = nil then exit;
  for i := 1 to CurG.RowCount - 1 do
  begin
    for j := 4 to CurG.ColCount - 1 do
    begin
      if BaseZoneGrisee(j, i) then GridGriseCell(CurG, j, i, Canvas)
      else GridDeGriseCell(CurG, j, i, Canvas);
    end;
  end;
  CurG.Invalidate;
end;

procedure TFSaisBul.GridColorCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
var
  T: TOB;
  Sens, Imprim: string;
  GS: THGrid;
begin
  // added by XP le 14-03-2003
  if not Assigned(CurG) then exit;

  GS := CurG;

  Sens := 'P';
  // PT53   26/11/2002 V591 PH le champ non imprimable d'une rémunération est affiché en gras+italiq
  if (ACol > 5) then exit;
  if (GS.Name <> 'GCot') and (GS.Name <> 'GBas') then
  begin
    T := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [GS.Cells[0, ARow]], FALSE);
    if T <> nil then Sens := T.GetValue('PRM_SENSBUL');
  end
  else exit;
  if (Sens <> 'M') then
    // Canvas.Font.Color   := Canvas.Brush.color
  else
    if ((ACol = 0) or (ACol = 5)) then Canvas.Font.Color := clRed;
  Imprim := 'X';
  if T = nil then exit; // cas 1ere ligne de la grille = intitulé de colonne
  case ACol of
    2: Imprim := T.GetValue('PRM_BASEIMPRIMABLE');
    3: Imprim := T.GetValue('PRM_TAUXIMPRIMABLE');
    4: Imprim := T.GetValue('PRM_COEFFIMPRIM');
  end;
  if Imprim = '-' then Canvas.font.Style := Canvas.Font.Style + [fsBold, fsItalic];
  if (T.Getvalue('PRM_IMPRIMABLE') = '-') and (Acol = 0) then Canvas.font.Style := Canvas.Font.Style + [fsUnderline];
end;

function TFSaisBul.AccesLigneBulletin(AccessSaisie, OrigineInfo: string): Boolean;
var
  st: string;
  rep: integer; // PT111
begin
  result := TRUE;
  if (AccessSaisie <> 'ABS') and (VH_Paie.PGModifLigneIMP) then exit; // PT98
  if (AccessSaisie <> '') or (OrigineInfo <> '') then
  begin
    st := 'La ligne n''est pas modifiable car elle provient ';
    if Origineinfo <> '' then
      st := st + 'd''un fichier d''import'
    else
      if AccessSaisie = 'ABS' then st := st + 'de la gestion des absences';
    //PT77   02/09/2003 V_421 PH Message indiquant la provenance extérieur de la ligne du bulletin
// d PT111
    if (AccessSaisie = 'MAI') then
    begin
//      rep := mrNo;
      st := 'La ligne ne devrait pas être modifiée car elle provient ';
      st := st + 'du traitement du maintien';
      st := st + '#13#10';
      st := st + 'Voulez vous continuer ?';
      rep := PGIAsk(st, 'Saisie du bulletin');
      if (rep = mrNo) then result := False;
    end
    else
    begin
      st := st + ' d''un module extérieur à la saisie du bulletin.';
      PGIBox(st, 'Saisie du bulletin');
      result := FALSE;
    end;
    // f PT111
  end;
end;

function TFSaisBul.AccesSuppLigneBulletin(AccessSaisie: string): Integer;
var
  st: string;
begin
  result := mrNo;
  if (AccessSaisie <> '') then
  begin
    st := 'La ligne ne devrait pas être supprimée car elle provient ';
    if AccessSaisie = 'ABS' then st := st + 'de la gestion des absences'
      //PT77   02/09/2003 V_421 PH Message indiquant la provenance extérieur de la ligne du bulletin
    else st := st + 'd''un module extérieur à la saisie du bulletin';
    st := st + '#13#10';
    st := st + 'Voulez vous continuer ?';
    result := PGIAsk(st, 'Saisie du bulletin');
  end
  else result := mrYes; // PT105
end;

procedure TFSaisBul.FTrentiemeEnter(Sender: TObject);
begin
  Trentieme_entree := FTrentieme.Text;
end;

procedure TFSaisBul.FTrentDenominateurEnter(Sender: TObject);
begin
  Denominateur_Entree := FTrentDenominateur.Text; //PT71-2
end;

procedure TFSaisBul.BVentilClick(Sender: TObject);
var
  ActionSurLaFiche: TActionFiche;
  TobAnal, TOBA, TPR, TPZ, UneTOB: TOB;
  i, ll: Integer;
  Nature, Rubrique, RefA, LeSal: string;
  okok, LeMode: Boolean;
begin
  if VH_Paie.PGAnalytique = FALSE then exit;
  if (CurG.Row < 1) then exit;
  Rubrique := CurG.Cells[0, CurG.Row];
  if (Copy(Rubrique, 5, 1) = '.') then
  begin
    if (POS('.R', Rubrique) <= 0) then
    begin
      PGIBOX('Une rubrique de commentaire n''a pas de ventilation', 'Saisie des bulletins');
      exit; // pas de ventilation sur une rubrique de commentaire
    end;
  end;
  if ActionBul = taConsultation then ActionSurLaFiche := taConsult;
  if ActionBul = taModification then ActionSurLaFiche := taModif;
  if ActionBul = taCreation then ActionSurLaFiche := taCreat;
  if (ActionSurLaFiche <> taConsult) and (ActionSurLaFiche <> taModif) and (ActionSurLaFiche <> taCreat) then exit;
  if Pages.ActivePage = PRemSal then CurG := GRemSal;
  if Pages.ActivePage = PRemSais then CurG := GRemSais; // PT202
  if Pages.ActivePage = PRemHeures then CurG := GRemHeures;
  if Pages.ActivePage = PRemPrimes then CurG := GRemPrimes;
  if Pages.ActivePage = PRemAbs then CurG := GRemAbs;
  if Pages.ActivePage = PRemCplt then CurG := GRemCplt;
  if Pages.ActivePage = PRemAvt then CurG := GRemAvt;
  if Pages.ActivePage = PRemAbt then CurG := GRemAbt;
  if Pages.ActivePage = PRemRet then CurG := GRemRet;
  if Pages.ActivePage = PRemNonImp then CurG := GRemNonImp;
  // PT3 : 03/09/2001 V547 PH Rajout un onglet Primes non imposables soumises à cotisation
  if Pages.ActivePage = PRemPrimeNonImpSoumis then CurG := GRemNonImpSoumis;
  Nature := 'AAA';
  if Pages.ActivePage = PBas then
  begin
    PGIBox('Il n''a pas de ventilations analytiques pour les bases de cotisation', 'Ventilations analytiques');
    exit;
  end;
  if Pages.ActivePage = PCot then
  begin
    CurG := GCot;
    Nature := 'COT';
    if ActionLigne <> taCreation then
    begin // pour forcer le calcul du bulletin car les ventilations analytiques des cotisations dependent des remunerations
      if CalculAFaire then CalculBullSaisie;
      CalculAFaire := FALSE;
    end;
  end;
  if ActionLigne = taCreation then
  begin
    CalculAFaire := TRUE; // pour forcer le calcul du bulletin car on a inseré une ligne
    CalculBullSaisie;
    CalculAFaire := FALSE;
    UneTob := nil;
    for i := 1 to CurG.RowCount - 1 do
    begin
      if (CurG.Cells[0, i] = Copy(Rubrique, 1, 4)) and ((POS('.R', CurG.Cells[0, i]) <= 0)) then
      begin
        UneTob := TOB(CurG.Objects[0, i]); // recupère la tob contenue dans la ligne
        Nature := UneTOB.GetVAlue('PHB_NATURERUB');
        Rubrique := UneTOB.GetVAlue('PHB_RUBRIQUE');
        LeSal := UneTOB.GetVAlue('PHB_SALARIE');
        break;
      end;
    end;
    if UneTob <> nil then
    begin
      if TOBAnal <> nil then
        FreeAndNil(TOBAnal);
      TOBAnal := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, UneTob, CodeSalarie, 'INS', DateD, DateF);
    end;
  end
  else
  begin
    ll := CurG.Row;
//    if ll > 1 then UneTob := TOB(CurG.Objects[0, ll - 1]) PT156
    UneTob := TOB(CurG.Objects[0, ll]); // recupère la tob contenue dans la ligne
  end;
  LeMode := FALSE; // Mode pour savoir si ventilation cotisation personnalisee ou remunération
  if (TobAna <> nil) then
  begin
    if (ActionLigne <> taCreation) then
    begin
      TobAnal := TOb.Create('Analytiques Lignes Bulletin', nil, -1);
      TPZ := TOBAna.findfirst([''], [''], TRUE);
      while TPZ <> nil do
      begin
        RefA := TPZ.GetValue('YVA_IDENTIFIANT');
        okok := DecodeRefPaieAnal(Rubrique, Nature, RefA);
        if (okok) then
          TPZ.ChangeParent(TobAnal, -1);
        TPZ := TOBAna.findnext([''], [''], TRUE);
      end;
    end;
    if (Nature = 'COT') and (TobAnal.Detail.Count <= 0) then
    begin // cas cotisation non personalisée alors on prend les ventilations par defaut
      LeMode := TRUE;
      TPZ := TOBAna.findfirst([''], [''], TRUE);
      while TPZ <> nil do
      begin
        RefA := TPZ.GetValue('YVA_IDENTIFIANT');
        OkOk := FALSE;
        //        i := ;    // @@@@@@@@@@@@@@
        if Pos('COT', RefA) > 0 then OkOk := TRUE;
        if (okok) then
        begin
          TOBA := TOB.Create('VENTANA', TobAnal, -1);
          TOBA.Dupliquer(TPZ, TRUE, TRUE, TRUE);
        end;
        TPZ := TOBAna.findnext([''], [''], TRUE);
      end;
    end;
    RefA := EncodeRefPaie(CodeSalarie, DateD, DateF, UneTob);
    if UneTob <> nil then
    begin
      // PT82   06/10/2003 V_421 PH Activation acces analytique en CWAS
{$IFNDEF EAGLSERVER}
      if not LeMode then YYVentilAna(TOB_Rub, TobAnal, ActionSurLaFiche, 0, RefA)
      else YYVentilAna(TOB_Rub, TobAnal, taConsult, 0, RefA); // ventilation cotisation non personnalisees donc non modifiables
{$ENDIF}
      // FIN PT82
    end
    else PGIBox('Impossible d''identifier la ligne du bulletin', 'Ventilation analytique');

    if not LeMode then
    begin // Recopie des ventilations modifiees
      TPR := TOBAnal.findfirst([''], [''], TRUE);
      while TPR <> nil do
      begin
        TPR.ChangeParent(TOBAna, -1);
        TPR := TOBAnal.findnext([''], [''], TRUE);
      end;
    end;

    if TobAnal <> nil then
      FreeAndNil(TobAnal);
    TOBAna.Detail.Sort('YVA_TABLEANA;YVA_IDENTIFIANT');
  end;

end;

// fonction de preaffectation des preventilations analytiques par defaut lors de la creation d'une ligne du bulletin

procedure TFSaisBul.AffectePreVentilAna;
var
  UneTob, TOBAnal, TPR: TOB;
  Nature, Rubrique: string;
begin
  TobAnal := nil;
  GrilleDansRub;
  UneTob := TOB(CurG.Objects[0, CurG.Row - 1]); // recupère la tob contenue dans la ligne
  Nature := UneTOB.GetVAlue('PHB_NATURERUB');
  Rubrique := UneTOB.GetVAlue('PHB_RUBRIQUE');
  if UneTob <> nil then
  begin
    if TOBAnal <> nil then
      FreeAndNil(TOBAnal);
    TOBAnal := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, UneTob, CodeSalarie, 'INS', DateD, DateF);
  end;
  if (TobAna <> nil) then
  begin
    TPR := TOBAnal.findfirst([''], [''], TRUE);
    while TPR <> nil do
    begin
      TPR.ChangeParent(TOBAna, -1);
      TPR := TOBAnal.findnext([''], [''], TRUE);
    end;
    if TobAnal <> nil then
      FreeAndNil(TobAnal);
    TOBAna.Detail.Sort('YVA_TABLEANA;YVA_IDENTIFIANT');
  end;
end;


procedure TFSaisBul.GCotRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  // CalculAFaire:=TRUE;
  // PT63   07/01/2003 V591 PH Ne calcule plus le bulletin systématiquement si on modifie les cotisations ie regarde paramsoc
  if VH_Paie.PGCalculBulletin then CalculAFaire := TRUE;
end;

procedure TFSaisBul.AffecteEtatFiche(Comment: Boolean);
begin
  //Enabled := Comment;
  EnableControls(X, Comment, TRUE);
end;

procedure TFSaisBul.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

{DEB PT8}

procedure TFSaisBul.AbsenceClick(Sender: TObject);
var
  Mode: string;
  Rep: Integer;
  NbJAbs: Double; { PT159 }
begin
  if ActionBul = TaConsultation then
  begin
    PgiBox('Vous ne pouvez consulter les absences : la paie est cloturée.', 'Visualisation des absences');
    exit;
  end;
{ DEB PT133 }
  Rep := MrNo;
  if ActionBul = TaCreation then Mode := 'C' else Mode := 'M';
  NbJAbs := JABSPris; { PT159 }
  if Mode = 'M' then
  begin
    Rep := PgiAsk('Attention! La visualisation des absences entraînera la modification du bulletin.#13#10' +
      'Voulez-vous continuer?', 'Visualisation des absences');
    if Rep = MrYes then AnnuleAbsenceBulletin(CodeSalarie, DateF);
  end;
  EnleveRubAbsence(Tob_Sal, Tob_Rub, DateD, DateF);
  if Assigned(Tob_Abs) then LibereTobAbs;
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'ABSENCE_MUL', '', 'PCN_SALARIE=' + CodeSalarie, CodeSalarie + ';BULLETIN;;' + DateToStr(DateF));
{$ENDIF}
  if (Mode = 'M') and (Rep = MrYes) then OnSauve;


  if (VH_PAIE.PGAbsence) then
  begin
{PT220
    if (RecupereAbsences(Tob_ABS, Datef, Tob_sal.GetValeur(iPSA_SALARIE), Mode, false, false)) then}
    if (RecupereAbsences(Tob_ABS, Datef, Tob_sal.GetValeur(iPSA_SALARIE), Mode, false, false,'')) then
    begin
      IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, Mode);
      RubDansGrille;
      CalculBullSaisie;
    end
    else { DEB PT159 }
      if (NbJAbs) <> (JABSPris) then
      begin
        RubDansGrille;
        CalculBullSaisie;
      end;
  end; { FIN PT159 }

{ FIN PT133 }
(*
//  DecodeDate(DateF, AA, MM, JJ);
///  Mois := IntToStr(MM);
//  if Length(Mois) = 1 then Mois := '0' + Mois;
// d PT124
// On vide la Tob_abs des absence non IJSS
  if Assigned(Tob_Abs) then    //Rep PT124 07/02/2005
  Begin
    t := Tob_Abs.findfirst(['PCN_GESTIONIJSS'], ['-'], false);
    while (t <> nil) do
    begin
      t.free;
      t := Tob_Abs.findnext(['PCN_GESTIONIJSS'], ['-'], false);
    end;
  End;
  // on charge la Tob_abs avec les absences non ijss
  if (VH_PAIE.PGAbsence) AND (RecupereAbsences(Tob_ABS, Datef, Tob_sal.GetValeur(iPSA_SALARIE), Mode, false, false)) then { PT132 Ajout Paramsoc }
  begin
    IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, Mode);
    RubDansGrille;
    CalculBullSaisie;
  end;
  // On vide la Tob_abs des absence  IJSS
  if Assigned(Tob_Abs) then    //Rep PT124 07/02/2005
  Begin
    t := Tob_Abs.findfirst(['PCN_GESTIONIJSS'], ['X'], false);
    while (t <> nil) do
    begin
      t.free;
      t := Tob_Abs.findnext(['PCN_GESTIONIJSS'], ['X'], false);
    end;
  End;
// f PT124
  // d PT111
  //  if RecupereAbsences(Tob_ABS, Datef, Tob_sal.GetValeur(iPSA_SALARIE), Mode) then
  // on n charge la Tob_abs avec les absences  ijss
  if (VH_PAIE.PGAbsence) AND (RecupereAbsences(Tob_ABS, Datef, Tob_sal.GetValeur(iPSA_SALARIE), Mode, false, true)) then { PT132 Ajout Paramsoc }
    // f PT111
  begin
    IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Tob_Sal, DateD, DateF, Mode);
    RubDansGrille;
    CalculBullSaisie;
  end;
  //FIN PT47-1  *)
end;
{FIN PT8}

procedure TFSaisBul.TimerBulTimer(Sender: TObject);
begin
  // PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin
  if ActiveCalc then
  begin
    if ImageCalcul.Visible then ImageCalcul.Visible := FALSE
    else ImageCalcul.Visible := TRUE;
  end; // FIN PT5
  if BullCompl <> 'X' then exit;
  if not TimerBul.Enabled then exit;
  if ImageBulCompl.Visible then ImageBulCompl.Visible := FALSE
  else ImageBulCompl.Visible := TRUE;
end;
// PT27 : 25/03/2002 V571 PH Modif Paieencours pour stockage Civilité et mode de réglement

procedure TFSaisBul.ChbxRegltModClick(Sender: TObject);
var
  ZDate: TDateTime;
begin
  if ChbxRegltMod.Checked = TRUE then
  begin
    FDatePaie.Enabled := TRUE;
    ModeRegle.Enabled := TRUE;
  end
  else
  begin // Reaffectation des valeurs par defaut en fonction des caractéristiques salariés
    ChpEntete.Reglt := RendModeRegle(TOB_Sal);
    ZDate := CalculDatePaie(TOB_Sal, DateF);
    if ZDate > 10 then FDatePaie.Text := DateToStr(ZDate)
    else FDatePaie.Text := '';
    ModeRegle.value := ChpEntete.Reglt;
    ChpEntete.DatePai := Zdate;
    FDatePaie.Enabled := FALSE;
    ModeRegle.Enabled := FALSE;
  end;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.FDatePaieOnExit(Sender: TObject);
begin
  ChpEntete.DatePai := StrToDate(FDatePaie.Text);
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.ModeRegleOnExit(Sender: TObject);
begin
  ChpEntete.Reglt := ModeRegle.Value;
  // PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation
  Modifier := TRUE;
end;

procedure TFSaisBul.CalendrierClick(Sender: TObject);
var
  Standard, TypeHoraire: string;
  T1: Tob;
begin
  Standard := '';
  if Tob_sal.GetValeur(iPSA_STANDCALEND) = 'PER' then
  begin
    Standard := Tob_sal.GetValeur(iPSA_CALENDRIER);
    TypeHoraire := 'SAL';
  end
  else
    if Tob_sal.GetValeur(iPSA_STANDCALEND) = 'ETB' then
    begin
      T1 := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [Tob_sal.GetValeur(iPSA_ETABLISSEMENT)], True);
      Standard := T1.getvalue('ETB_STANDCALEND');
      TypeHoraire := 'STD';
    end
    else
      if Tob_sal.GetValeur(iPSA_STANDCALEND) = 'ETS' then
      begin
        Standard := Tob_sal.GetValeur(iPSA_CALENDRIER);
        TypeHoraire := 'STD';
      end;

  if Standard = '' then
  begin
    PgiBox('Aucun calendrier n''est affecté au salarié!', caption);
    exit;
  end;
  //TypeHoraire:=Tob_sal.GetValeur (iPSA_STANDCALEND);
{$IFNDEF EAGLSERVER}
  AgllanceFiche('AFF', 'HORAIRESTD', '', '', 'TYPE:' + TypeHoraire + ';CODE:' + CodeSalarie + ';LIBELLE:' + Tob_sal.GetValeur(iPSA_LIBELLE) + ' ' + Tob_sal.GetValeur(iPSA_PRENOM)
    +
    ';STANDARD:' + Standard);
{$ENDIF}
  {DEB PT47-4 Suite au Zoom Calendrier on recalcul les variables en fn du nouveau calendrier
    recup en parametre heures et jours ouvrés ouvrables}
  with ChpEntete do
  begin
      { PT132 Mise en commentaire , utilisation du calendier chargé en globale
      Tob_CalendrierSalarie := ChargeCalendrierSalarie(Tob_sal.GetValeur(iPSA_ETABLISSEMENT),
      Tob_sal.GetValeur(iPSA_SALARIE), Tob_sal.GetValeur(iPSA_CALENDRIER),
      Tob_sal.GetValeur(iPSA_STANDCALEND), Standard); }
    CalculVarOuvresOuvrablesMois(TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [Tob_sal.GetValeur(iPSA_ETABLISSEMENT)], True),
      Tob_sal, GblTob_Semaine, DateD, DateF, TRUE, Ouvres, HOuvres); { PT132 }
    CalculVarOuvresOuvrablesMois(TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [Tob_sal.GetValeur(iPSA_ETABLISSEMENT)], True),
      Tob_sal, GblTob_Semaine, DateD, DateF, FALSE, Ouvrables, HOuvrables); { PT132 }
  end; //FIN PT47-4
end;

// PT35 : 14/05/2002 V582 PH FQ 10113 Controles si zones modifiées pour validation

procedure TFSaisBul.LbllTrentiemeClick(Sender: TObject);
begin
  Modifier := TRUE;
end;
// PT55   06/12/2002 V591 PH Optimisation fonction dessin cellules et rajout indicateur calcul bulletin

procedure TFSaisBul.GRemKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (ActionBul <> taConsultation) and (CurG <> nil) and (CurG.Col <> SG_LIB) then ActiveCalc := TRUE;
end;
// PT65-1   16/01/2003 V591 PH FQ 10449 Désactivation insertion et suppression de ligne pdt la modif des dates de paie

procedure TFSaisBul.FDatePaieEnter(Sender: TObject);
begin
  BNewLigne.Enabled := FALSE;
  BDelLigne.Enabled := FALSE;
  //DEB PT65-2
  if IsValidDate(FDate1.Text) then DDOnEnter := StrToDate(FDate1.Text);
  if IsValidDate(FDate2.Text) then DFOnEnter := StrToDate(FDate2.Text);
  //FIN PT65-2
end;

procedure TFSaisBul.ReactiveBtn;
begin
  BNewLigne.Enabled := TRUE;
  BDelLigne.Enabled := TRUE;
end;
// FIN PT65-1
// PT78   15/09/2003 V_421 PH Simulation saisie arret dans la saisie du bulletin

procedure TFSaisBul.VSaisieArretClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  AGLLANCEFICHE('PAY', 'TRANCHEREMRET', '', '', FSalarie.Caption + ';' + FDate1.text + ';' + FDate2.text);
{$ENDIF}
end;
// PT78

{ DEB PT84 Gestion affichage des zooms liés au CP }

procedure TFSaisBul.PGEnabledRaccourciCp;
var
  T_Etab: Tob;
begin
  if not assigned(X) then exit;
  if (VH_paie.PGCongesPayes) then
  begin
    X.CongesPayes1.Visible := True;
    X.Voir1.Visible := True;
    X.VoirCP.Visible := True;
    X.VoirvalorisationAbsence1.Visible := True;
    X.Voirindemnitcompensatricedecongs1.Visible := True;
    X.CalculCpPaie.Visible := True;
    X.Salarie1.Enabled := True;
    X.Etablissement1.Enabled := True;
    X.Sal.Enabled := True; { PT149 }
    X.EtablissPaie.Enabled := True; { PT149 }
    if (TOB_Salarie.GetValeur(iPSA_DATESORTIE) <= Datef) and (TOB_Salarie.GetValeur(iPSA_DATESORTIE) >= DateD) then
      if TOB_Etablissement <> nil then
      begin
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
        if (T_Etab <> nil) then
          if (T_Etab.GetValue('ETB_CONGESPAYES') = 'X') and (TOB_Salarie.GetValeur(iPSA_CONGESPAYES) = 'X') then { PT89 }
            if (PGRecupereProfilCge(TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)) = '') and (ActionBul = TaCreation) then
            begin
              X.Salarie1.Enabled := False;
              X.Etablissement1.Enabled := False;
              X.Sal.Enabled := False; { PT149 }
              X.EtablissPaie.Enabled := False; { PT149 }
            end;
      end;
  end
  else
  begin
    X.CongesPayes1.Visible := False;
    X.Voir1.Visible := False;
    X.VoirCP.Visible := False;
    X.VoirvalorisationAbsence1.Visible := False;
    X.Voirindemnitcompensatricedecongs1.Visible := False;
    X.CalculCpPaie.Visible := False;
  end;
end;
{ FIN PT84 }
{ DEB PT87 }

procedure TFSaisBul.CPAcquisExit(Sender: TObject);
var
  Modif: Boolean;
begin
  Modif := False;
  if Sender <> nil then { DEB PT87 13/05/2004 }
  begin
    if IsNumeric(THCritMaskEdit(sender).Text) then
      if Valeur(THCritMaskEdit(sender).Text) < 0 then
      begin
        PgiBox('Vous ne pouvez saisir de valeur négative.', Caption);
        if THCritMaskEdit(sender).Name = 'CPMOIS' then THCritMaskEdit(sender).Text := FloatToStr(Arrondi(ChpEntete.CpAcq, DCP))
        else if THCritMaskEdit(sender).Name = 'CPSUPP' then THCritMaskEdit(sender).Text := FloatToStr(Arrondi(ChpEntete.CpSupl, DCP))
        else if THCritMaskEdit(sender).Name = 'CPANC' then THCritMaskEdit(sender).Text := FloatToStr(Arrondi(ChpEntete.CpAnc, DCP));
      end;
  end; { FIN PT87 13/05/2004 }
  if IsNumeric(CPMois.Text) then
    if Arrondi(Valeur(CPMois.Text), DCP) <> Arrondi(ChpEntete.CpAcq, DCP) then
    begin
      Modif := True;
      ChpEntete.CpAcq := Arrondi(Valeur(CPMois.Text), DCP);
    end;

  if IsNumeric(CpSupp.Text) then
    if Arrondi(Valeur(CpSupp.Text), DCP) <> Arrondi(ChpEntete.CpSupl, DCP) then
    begin
      Modif := True;
      ChpEntete.CpSupl := Arrondi(Valeur(CpSupp.Text), DCP);
    end;

  if IsNumeric(CpAnc.Text) then
    if Arrondi(Valeur(CpAnc.Text), DCP) <> Arrondi(ChpEntete.CpAnc, DCP) then
    begin
      Modif := True;
      ChpEntete.CpAnc := Arrondi(Valeur(CpAnc.Text), DCP);
    end;

  if Modif then
  begin
    InitialiseVariableStatCP(FALSE); // PT208
    RecalculCp(False);
  end;

end;
{ FIN PT87 }

{ DEB PT87 13/05/2004 }

procedure TFSaisBul.CBModifAcquisClick(Sender: TObject);
var
  T: Tob;
//PT210  nbmois, base, Mois, Supp, Anc: double;
  nbmois, base, Mois, Supp, Anc, Fract: double;
begin
  ChpEntete.CpAcquisMod := CBModifAcquis.Checked;
  CPMois.Enabled := CBModifAcquis.Checked;
  CPSUPP.Enabled := CBModifAcquis.Checked;
  CPANC.Enabled := CBModifAcquis.Checked;
  if CBModifAcquis.Checked = False then
  begin
    T := Tob_etablissement.findfirst(['ETB_ETABLISSEMENT'], [Tob_salarie.GetValeur(iPSA_ETABLISSEMENT)], false);
    if T = nil then exit;
    FreeAndNil(T_MvtAcquisAVirer);
//PT210    ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, False, base, Nbmois, Mois, Supp, Anc);
    ChargeAcquisParametre(TOB_Salarie, T, TOB_RUB, DateD, DateF, False, False, base, Nbmois, Mois, Supp, Anc, Fract);
    if (Arrondi(ChpEntete.CpAcq, DCP) <> Arrondi(Mois, DCP))
      or (Arrondi(ChpEntete.CpSupl, DCP) <> Arrondi(Supp, DCP))
      or (Arrondi(ChpEntete.CpAnc, DCP) <> Arrondi(Anc, DCP)) then
    begin
      InitialiseVariableStatCP;
      ChpEntete.CpAcq := Arrondi(Mois, DCP);
      ChpEntete.CpSupl := Arrondi(Supp, DCP);
      ChpEntete.CpAnc := Arrondi(Anc, DCP);
      RecalculCp(False);
    end;
  end;
end;
{ FIN PT87 }

{DEB  PT127-1 }

procedure TFSaisBul.LanceCalculPaieCp;
begin
  if Assigned(TOB_Sal) then
    if not ((ActionBul = Tamodification) and (TOB_Sal.GetValeur(iPSA_DATESORTIE) >= DateD) and (TOB_Sal.GetValeur(iPSA_DATESORTIE) <= DateF)) then { PT159 }
    begin
      InitialiseVariableStatCP;
      RecalculCp(False);
    end
    else
      PGIBox('Le recalcul des congés payés ne s''effectue pas en modification du solde de tout compte.', 'Congés payés');
end;
{FIN  PT127-1 }

{DEB  PT127-1 }

procedure TFSaisBul.CalculCpPaieClick(Sender: TObject);
begin
  LanceCalculPaieCp;
end;
{FIN PT127-1 }
// DEB PT134  Acces aux elements nationaux pdt la saisie du bulletin

procedure TFSaisBul.EltNationauxClick(Sender: TObject);
var
  DateFin: TDateTime; //PT198
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41105)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
    exit;
  end;
  AglLanceFiche('PAY', 'ELTNAT_MUL', '', '', '');
// Faire le nettoyage et le rechargement des TOB DOS,STD,CEG des eltnat FQ11840
  Nettoyage_EltNationauxSTD();
  Nettoyage_EltNationauxCEG();
  Nettoyage_EltNationauxDOS();
  Nettoyage_EltNiveauRequis();
  initTOB_EltNatDOS();
  initTOB_EltNatSTD();
  initTOB_EltNatCEG();
  initTOB_EltNiveauRequis();

  //DEB PT198
  DateFin := StrToDate(FDate2.Text);
  if GetParamSocSecur('SO_PGGESTELTDYNDOS', False) then
  begin
    Nettoyage_EltDynSal();
    Nettoyage_EltDynPop();
    Nettoyage_EltDynEtab();
    initTOB_EltDynSal(CodeSalarie, DateFin);
    initTOB_EltDynEtab(Etab, DateFin);
    initTOB_EltDynPop(CodeSalarie, DateFin);
  end;
  //FIN PT198
{$ENDIF}
end;
// FIN PT134

procedure TFSaisBul.LigneRegul(Sender: TObject);
var
  i: Integer;
  Rub, St1: string;
begin
  if ActionBul = taConsultation then exit;
  if Curg.Name = 'PBas' then
  begin
    PGIBOX('Vous ne pouvez pas régulariser des rubriques de base de cotisation', 'Saisie des bulletins'); // PT142
    exit;
  end;
  if (CurG.Row < 1) then exit;
  Rub := CurG.Cells[0, CurG.Row];
  if Rub = '' then
  begin
    PGIBOX('Vous ne pouvez pas insérer de ligne de régularisation sans rubrique !', 'Saisie des bulletins');
    exit;
  end;
  if (Pos('.R', Rub)) > 0 then
  begin
    PGIBOX('Vous ne pouvez pas créer de régularisation sur une rubrique de régularisation !', 'Saisie des bulletins');
    exit;
  end;

  for i := 1 to CurG.RowCount - 1 do // PT176
  begin
    St1 := CurG.Cells[0, i]; // PT176
    if (Copy(St1, 1, 4) = rub) and (POS('.R', St1) > 0) then
    begin
      PGIBOX('Vous ne pouvez pas insérer 2 lignes de régularisation pour la même rubrique !', 'Saisie des bulletins');
      exit;
    end;
  end;
  bNewLigneClick(Sender);
  CurG.ElipsisButton := FALSE;
  Rub := Rub + '.R';
  CurG.Cells[0, CurG.Row] := Rub;
  CurG.Cells[1, CurG.Row] := 'Régularisation ' + CurG.Cells[1, CurG.Row + 1];
  CurG.Cells[2, CurG.Row] := CurG.Cells[2, CurG.Row + 1];
  CurG.Col := 1;
end;

procedure TFSaisBul.AccderDiodefr1Click(Sender: TObject);
var strWeb, strCodeDiode: string;
  Q: TQuery;
  CodePer: Integer;
begin
{$IFNDEF EAGLSERVER}
  if (V_PGI.ModePCL <> '1') then LanceWeb('http://www.diode.fr', false)
  else
  begin
  // DEB PT161
    if not VH_Paie.SeriaDiodePcl then
    begin
      PgiInfo('Le Service DIODE n''est pas sérialisé, vous ne pouvez pas y accéder.', 'Service DIODE'); // PT166
      exit;
    end;
  // FIN PT161
    strWeb := GetParamSoc('SO_ENVDISK8');
    if strWeb = '' then
    begin
      try
        Q := OpenSql('@@select SOC_DATA from DB000000.DBO.PARAMSOC WHERE SOC_NOM = "SO_ENVDISK8"', TRUE);
        if not Q.EOF then strWeb := Q.FindField('SOC_DATA').AsString;
        FERME(Q);
      except
      end;
    end;
    if strWeb <> '' then
    begin
      Q := OpenSql('SELECT DOS_CODEPER FROM DOSSIER WHERE DOS_NODOSSIER="' + V_PGI.NODOSSIER + '"', TRUE);
      if not Q.EOF then CodePer := Q.FindField('DOS_CODEPER').AsInteger;
      FERME(Q);
      Q := OpenSQL('SELECT DSO_ORIGDADS FROM DPSOCIAL WHERE DSO_NODP=' + IntToStr(CodePer), TRUE);
      if not Q.EOF then
      begin
        strCodeDiode := Q.FindField('DSO_ORIGDADS').AsString;
        if strCodeDiode <> '' then LanceWeb(strWeb + '\cc_index_obligation.html', FALSE)
        else LanceWeb(strWeb + '\index.html', FALSE);
      end
      else LanceWeb('http://www.diode.fr', FALSE);
      Ferme(Q);
    end;
  end;
{$ENDIF}
end;
// DEB PT169

procedure TFSaisBul.OnDiagClick(Sender: TObject);
var Rub, lib: string;
  Base, Taux, Coeff, Montant: Double;
  TxSal, TxPat, MtSal, MtPat: Double;
  Plafond, Plf1, Plf2, Plf3, Tr1, Tr2, Tr3: Double;
  AT: Boolean;
begin
  Rub := CurG.Cells[0, Curg.Row];
  LogClear(Diag); //PT211   Diag.Clear;
  CalculBulletin(Tob_Rub);
  if Curg.Name = 'GCot' then
  begin
    EvalueCot(Tob_Rub, rub, Base, TxSal, TxPat, MtSal, MtPat, Lib, DateD, DateF, taModification, AT
      , LogGetChildLevel(Diag.Items.Add(nil, TraduireMemoire('Calcul de la cotisation'))))
  end else begin
    if Curg.Name = 'GBas' then
    begin
      EvalueBas(Tob_Rub, rub, Base, Plafond, Plf1, Plf2, Plf3, Tr1, Tr2, Tr3, lib, DateD, DateF, taModification, ChpEntete.BasesMod, ChpEntete.TranchesMod
        , LogGetChildLevel(Diag.Items.Add(nil, TraduireMemoire('Calcul de la base'))))
    end else begin
      EvalueRem(TOB_Rub, Rub, Base, Taux, Coeff, Montant, Lib, DateD, DateF, taModification, Curg.Row
        , LogGetChildLevel(Diag.Items.Add(nil, TraduireMemoire('Calcul de la rémunération'))));
    end;
  end;
  TBSHTDIAG.TabVisible := TRUE;
  Pages.ActivePage := TBSHTDiag;
end;

procedure TFSaisBul.DiagRubClick(Sender: TObject);
begin
  OnDiagClick(Sender);
  //DEB PT203
  StDiagRub := CurG.Cells[0, Curg.Row] + ' ' + CurG.Cells[1, Curg.Row];
  StDiagBul := '';
  //FIN PT203
end;

procedure TFSaisBul.DiagBullClick(Sender: TObject);
begin
  LogClear(Diag); //PT211    Diag.Clear;
  { Création du noeud racine }
  CalculBulletin(Tob_Rub, FALSE, Diag.Items.Add(nil, TraduireMemoire('Calcul du bulletin')));
  TBSHTDIAG.TabVisible := TRUE;
  Pages.ActivePage := TBSHTDiag;
  //DEB PT203
  StDiagRub := '';
  StDiagBul := 'Bull';
  //FIN PT203
end;
// FIN PT169

//DEB PT183

procedure TFSaisBul.Voirhistoriquecumuls1Click(Sender: TObject);
begin
  AccesCumulsClick;
end;

procedure TFSaisBul.AccesCumulsClick;
var
  TOB_CumAff, TOB_CumAffFille, T1, T2, LeCumul: TOB;
  Q: TQuery;
  St: string;
  DD, DF: TDateTime;
  i: Integer;
begin
  TOB_CumAff := TOB.create('Les histo Cumuls Bulletin', nil, -1);
  RendDateExerSocial(DateD, DateF, DD, DF);
  if DF < DD then DF := DD;
  St := 'select PCL_CUMULPAIE,PCL_LIBELLE,PHC_DATEDEBUT,PHC_DATEFIN,SUM(PHC_MONTANT) MONTANT FROM CUMULPAIE ' +
    'left join HISTOCUMSAL ON PHC_CUMULPAIE = PCL_CUMULPAIE ' +
    'where ##PCL_PREDEFINI## PHC_SALARIE="' + CodeSalarie +
    '" AND PHC_DATEDEBUT >="' + UsDateTime(DD) + '" AND PHC_DATEFIN <="' + UsDateTime(DF) + '" ' +
    'group by PCL_CUMULPAIE,PCL_LIBELLE,PHC_DATEDEBUT,PHC_DATEFIN';
  Q := OpenSql(st, TRUE);
  while not Q.Eof do
  begin
    TOB_CumAffFille := TOB.create('Un cumul', TOB_CumAff, -1);
    TOB_CumAffFille.AddChampSup('PCL_CUMULPAIE', False);
    TOB_CumAffFille.AddChampSup('PCL_LIBELLE', False);
    TOB_CumAffFille.AddChampSup('PHC_DATEDEBUT', False);
    TOB_CumAffFille.AddChampSup('PHC_DATEFIN', False);
    TOB_CumAffFille.AddChampSup('MONTANT', False);
    TOB_CumAffFille.PutValue('PCL_CUMULPAIE', Q.FindField('PCL_CUMULPAIE').AsString);
    TOB_CumAffFille.PutValue('PCL_LIBELLE', Q.FindField('PCL_LIBELLE').AsString);
{    TOB_CumAffFille.PutValue('PHC_DATEDEBUT', DateToStr(Q.FindField('PHC_DATEDEBUT').AsDateTime));
    TOB_CumAffFille.PutValue('PHC_DATEFIN', DateToStr(Q.FindField('PHC_DATEFIN').AsDateTime));
    }
// DEB  PT193
    TOB_CumAffFille.PutValue('PHC_DATEDEBUT', Q.FindField('PHC_DATEDEBUT').AsDateTime);
    TOB_CumAffFille.PutValue('PHC_DATEFIN', Q.FindField('PHC_DATEFIN').AsDateTime);
// FIN PT193
    TOB_CumAffFille.PutValue('MONTANT', Q.FindField('MONTANT').AsFloat);
    Q.Next;
  end;
  Ferme(Q);

  for i := TOB_CumulSal.Detail.Count - 1 downto 0 do
  begin
    T1 := TOB_CumulSal.Detail[I];
    T2 := TOB_CumAff.FindFirst(['PCL_CUMULPAIE'], [T1.GetValue('PHC_CUMULPAIE')], FALSE);

    if (T1 <> nil) then
    begin
      if (T1.GetValue('PHC_CUMULPAIE') >= '01') then
      begin
        TOB_CumAffFille := TOB.create('Un cumul', TOB_CumAff, -1);
        TOB_CumAffFille.AddChampSup('PCL_CUMULPAIE', False);
        TOB_CumAffFille.AddChampSup('PCL_LIBELLE', False);
        TOB_CumAffFille.AddChampSup('PHC_DATEDEBUT', False);
        TOB_CumAffFille.AddChampSup('PHC_DATEFIN', False);
        TOB_CumAffFille.AddChampSup('MONTANT', False);
        TOB_CumAffFille.PutValue('PCL_CUMULPAIE', T1.GetValue('PHC_CUMULPAIE'));
        if (T2 <> nil) then
          TOB_CumAffFille.PutValue('PCL_LIBELLE', T2.GetValue('PCL_LIBELLE'))
        else
        begin
          Q := OpenSql('SELECT PCL_LIBELLE FROM CUMULPAIE WHERE PCL_CUMULPAIE = "' + T1.GetValue('PHC_CUMULPAIE') + '"', TRUE);
          TOB_CumAffFille.PutValue('PCL_LIBELLE', Q.FindField('PCL_LIBELLE').AsString);
          Ferme(Q);
        end;
{        TOB_CumAffFille.PutValue('PHC_DATEDEBUT', StrToDateTime(T1.GetValue('PHC_DATEDEBUT')));
        TOB_CumAffFille.PutValue('PHC_DATEFIN', StrToDateTime(T1.GetValue('PHC_DATEFIN')));
}
// DEB PT193
        TOB_CumAffFille.PutValue('PHC_DATEDEBUT', T1.GetValue('PHC_DATEDEBUT'));
        TOB_CumAffFille.PutValue('PHC_DATEFIN', T1.GetValue('PHC_DATEFIN'));
// FIN PT193
        TOB_CumAffFille.PutValue('MONTANT', StrToFloat(T1.GetValue('PHC_MONTANT')));
      end;
    end;
  end;

  TOB_CumAFF.Detail.Sort('PCL_CUMULPAIE;PHC_DATEDEBUT;PHC_DATEFIN');

{$IFNDEF EAGLSERVER}
  TheTob := TOB_CumAFF;
  St := CodeSalarie + ';' + Etab + ';' + DateToStr(DateD) + ';' + DateToStr(DateF);
  AglLanceFiche('PAY', 'BUL_HISTOCUMULS', '', '', St);
  TheTOB := nil;
{$ENDIF}
  FreeAndNil(TOB_CumAff);
end;
//FIN PT183

procedure TFSaisBul.SauvDiagClick(Sender: TObject);
var
  Entete: TStringList;
begin
  Entete := TStringList.Create;
  Entete.Add(TraduireMemoire('                   DIAGNOSTIC BULLETIN '));
  Entete.Add(TraduireMemoire(' Salarié : ') + FSalarie.Caption + TraduireMemoire(' Paie du ') + FDate1.text + TraduireMemoire(' au ') + FDate2.text);
  if OpenDialog1.Execute then
    LogSaveToFile(OpenDialog1.Filename, Diag, Entete); //PT211
  FreeAndNil(Entete);
//  if OpenDialog1.Execute then
//  begin
//    LeFic := OpenDialog1.FileName;
//    if LeFic = '' then exit;
//    AssignFile(F, OpenDialog1.FileName);
//    if not FileExists(LeFic) then
//    begin
//      AssignFile(F, LeFic);
//{$I-}ReWrite(F);
//{$I+}if IoResult <> 0 then
//      begin
//        PGIBox('Fichier inaccessible : ' + LeFic, 'Abandon du traitement');
//        Exit;
//      end;
//      closeFile(F);
//    end;
//
//    if OpenDialog1.FileName <> '' then
//    begin
//      Nb := Diag.Items.Count;
//      AssignFile(F, LeFic);
//      reset(F);
//      ReWrite(F);
//      Writeln(F, TraduireMemoire('                   DIAGNOSTIC BULLETIN '));
//      Writeln(F, TraduireMemoire(' Salarié : ') + FSalarie.Caption + TraduireMemoire(' Paie du ') + FDate1.text + TraduireMemoire(' au ') + FDate2.text);
//      Writeln(F, ' ');
//      for i := 0 to Nb - 1 do Writeln(F, Diag.items[i]);
//      CloseFile(F);
//    end;
//  end; // PT201
end;

procedure TFSaisBul.EnvMailDiagClick(Sender: TObject);
var ListeJ: HTStrings;
begin
{$IFNDEF EAGLSERVER} //PT186
  SauvDiagClick(Sender);
  if OpenDialog1.FileName <> '' then
  begin
    ListeJ := HTStringList.Create;
    ListeJ.Add('Veuillez trouver ci-joint notre diagnostic bulletin.');
    ListeJ.Add('Cordialement');

    SendMail(TraduireMemoire(' Cegid Business Paie : diagnostic bulletin '), '', '', ListeJ, OpenDialog1.FileName, FALSE);
    ListeJ.Clear;
    ListeJ.Free;
  end;
{$ENDIF} //PT186
end;

//DEB PT187

procedure TFSaisBul.VisualiserBulletinsClick(Sender: TObject);
var
  St: string;
begin
  St := Etab + ';' + CodeSalarie + ';' + DateToStr(DateF);
{$IFNDEF EAGLSERVER}
  if (Etab <> '') and (CodeSalarie <> '') and (DateF > 0) then
    //DEB PT182
    // Si on ne gère pas le mode d'édition
    if not GetParamSocSecur('SO_PGGESTORIDUPSPE', False) then
      AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';VISUBUL;' + BullCompl) //PT22
    else
    begin
      // Vérifier que l'utilisateur a au moins le droit d'éditer un duplicata ou un specimen
      // A ce niveau, on ne propose pas l'édition d'un original
      if JaiLeDroitTag(42314) or JaiLeDroitTag(42316) then
      begin
//PT191        AglLanceFiche('PAY', 'MODEEDT_BUL', '', '', 'SAISBUL');
//PT191        if PGModeEdition <> '' then
        AglLanceFiche('PAY', 'EDITBUL_ETAT', '', '', St + ';VISUBUL;' + BullCompl + ';DUPLICATA');
      end
      else
        PGIBox(TraduireMemoire('Vous n''êtes pas autorisé à lancer une édition de bulletins'));
    end;
    //FIN PT182
{$ENDIF}
end;
//FIN PT187

// d PT189
{**************************************************************
Auteur  ...... : Monique FAUDEL
Créé le ...... : 03/05/2007
Modifié le ... :   /  /
Description .. : Click sur bouton CALC lance le Calcul des
Suite ........ : IJSS brutes (utlisation elt nat 0098)
Suite ........ : et de la garantie du net (calcul de la paie
Suite ........ : à l'envers)
Mots clefs ... : PAIE;PGBULLETIN ; IJSS
*****************************************************************}

procedure TFSaisBul.CalculIJSSClick(Sender: TObject);
var
  rep: string;
begin
{$IFNDEF EAGLSERVER}
// d PT196
  if (VH_Paie.PGGestIJSS) then
    PgiBox('L''utilisation de cette commande peut être incompatible #13#10' +
      ' avec la gestion automatique des IJSS.', '');
// f PT196

  TheTob := TOB_Rub;
  rep := AglLanceFiche('PAY', 'PAIEENVERS', '', '', 'IJSS');
  if rep = 'O' then
  begin
    RubDansGrille;
    AfficheBasBulletin;
  end;
  TheTob := nil;
{$ENDIF}
end;
// f PT189

procedure TFSaisBul.TablesdynamiquesClick(Sender: TObject);
begin
{$IFNDEF EAGLSERVER}
  if (not JaileDroitTag(41412)) then // PT215
  begin
    PgiInfo('Vous n''êtes pas autorisé(e) à accèder à cette fonction.', 'Saisie du bulletin');
  end
  else
    AglLanceFiche('PAY', 'PGMULTABLESDYNA', '', '', '');
{$ENDIF}
end;

// DEB PT194

procedure TFSaisBul.AnalSalClick(Sender: TObject);
var ActionAfaire : TActionFiche;
begin
// DEB PT206
  if ActionBul = taConsultation then
      ActionAfaire := taConsult
    else
      ActionAfaire := taModif;
// FIN PT206      
{$IFNDEF EAGLSERVER}
  ParamVentil('SA', CodeSalarie, '12345', ActionAfaire, FALSE);
{$ENDIF}
  CalculAFaire := TRUE; // Calcul forcé
  ForceAna := TRUE;
  if TobAna <> nil then
  begin
    TobAna.free;
    TObAna := nil;
  end;
  CalculBullSaisie;
  ForceAna := FALSE;
end;
// FIN PT194

procedure TFSaisBul.AnalrubbulClick(Sender: TObject);
begin
  BVentilClick(Sender)
end;
// DEB PT202

function TFSaisBul.RemSaisissable(Trech: TOB): Boolean;
var st: string;
begin
  result := FALSE;
  if TRech = nil then exit;
  st := TRech.GetValue('PRM_TYPEBASE');
  if (st = '00') or (st = '01') then
  begin // base saisissable
    result := True;
    exit;
  end;
  st := TRech.GetValue('PRM_TYPETAUX');
  if (st = '00') or (st = '01') then
  begin // TAUX saisissable
    result := True;
    exit;
  end;
  st := TRech.GetValue('PRM_TYPECOEFF');
  if (st = '00') or (st = '01') then
  begin // Coefficient saisissable
    result := True;
    exit;
  end;
  st := TRech.GetValue('PRM_TYPEMONTANT');
  if (st = '00') or (st = '01') then
  begin // Montant saisissable
    result := True;
    exit;
  end;
end;
// FIN PT202

//DEB PT203

procedure TFSaisBul.Editiondudiagnostic1Click(Sender: TObject);
var
  TobEdit, TOBE: Tob;
  i: integer;
  Titre: string;
  StLLog: TStringList;
begin
{$IFNDEF EAGLSERVER}
  StLLog := LogToStrings(Diag); //PT211
  if StLLog.Count <> 0 then //PT211
//  if (Diag.Count <> 0) then
  begin
    TobEdit := TOB.Create('Leserreurs', nil, -1);

    TOBE := TOB.Create('Leserreursdetail', TobEdit, -1);
    TOBE.AddChampSupValeur('TYPE', 'E', false);
    TOBE.AddChampSupValeur('LIGNE', 'Etablissement ' + TPSA_ETABLISSEMENT.Caption);
    TOBE := TOB.Create('Leserreursdetail', TobEdit, -1);
    TOBE.AddChampSupValeur('TYPE', 'E', false);
    TOBE.AddChampSupValeur('LIGNE', 'Salarié ' + FSalarie.Caption + ' ' + TPSA_LIBELLE.Caption);
    TOBE := TOB.Create('Leserreursdetail', TobEdit, -1);
    TOBE.AddChampSupValeur('TYPE', 'E', false);
    TOBE.AddChampSupValeur('LIGNE', 'Bulletin du ' + FDate1.Text + ' au ' + FDate2.Text);
    for i := 0 to StLLog.Count - 1 do // PT211
//    For i := 0 to Diag.Count -1 do
    begin
      TOBE := TOB.Create('Leserreursdetail', TobEdit, -1);
      TOBE.AddChampSupValeur('TYPE', 'D', false);
      TOBE.AddChampSupValeur('LIGNE', StLLog.Strings[i], false);
//      TOBE.AddChampSupValeur('LIGNE',Diag.Items[i],false);
    end;
    if (StDiagRub <> '') then
      Titre := 'Diagnostic de la rubrique ' + StDiagRub
    else
      Titre := 'Diagnostic du bulletin';
    LanceEtatTob('E', 'PAY', 'GLS', TobEdit, True, False, False, Pages, '', Titre, False);
    FreeAndNil(TobEdit);
  end
  else
    PGIBox('Il n''y a aucun diagnostic à imprimer');
  FreeAndNil(StLLog); //PT211
{$ENDIF}
end;
//FIN PT203

end.


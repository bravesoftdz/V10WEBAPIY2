{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/06/2001
Modifié le ... :   /  /
Description .. : Gestion de la fiche salaries
Mots clefs ... :
*****************************************************************}
{
PT1   : 15/06/2001 PH V536 Affichage des congés payés du salariés
                           Cas : Les salaries ayant des mvts Cp sur pls
                           etablissements
                           Correction : On ne tient pas compte de la notion
                           d'etab pour l'affichage des CP.
PT2   : 25/06/2001 PH V536 Boucle affectation automatique du numéro de salarié
                           et suppression order by dans clause max
PT3   : 25/07/2001 VG V540 Ajout champs dans Historique salarié
PT4   : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT5   : 07/09/2001 SB V547 Controle Date de sortie / Bulletin de paie
PT6   : 10/09/2001 SB V547 Champ PSA_PGMODEREGLE DataType non renseigné
PT7   : 12/09/2001 SB V547 Champ PPS_SALARIE inexistant modification requête
PT8   : 12/09/2001 PH V547 Convention collective mise par défaut en creation
PT9   : 18/09/2001 PH V547 Rajout mise à jour ventilation analytique salarie par
                           défaut si on gère une codification atomatique
PT10  : 02/10/2001 PH V642 Rajout historisation zone profil rémunérations
PT11  : 02/10/2001 PH V642 traitement Idem etab pour Profilrem
PT12  : 02/10/2001 PH V642 Champ analytique personnalisé visible si gestion
                           automatique
PT13  : 08/10/2001 SB V642 Test periode bulletin posterieur à la date de sortie
                           fiche bug n°325
PT14  : 11/10/2001 V562    correction du PT5 qui est un controle bloquant et qui
                           empeche de faire un bulletin de régularisation après
                           le solde de tout compte
PT15  : 15/10/2001 VG V562 Contrôle sur la fraction DADS-U par rapport au
                           paramètre société
PT16  : 15/10/2001 VG V562 Ajout champs dans Historique salarié
PT17  : 25/10/2001 SB V563 Acces Zone CP Acq en cours et Date Acq Anc Mise en
                           commentaire
PT18  : 26/10/2001 SB V563 Controle SetField systématique et champ idem etab
                           erronne
PT19  : 31/10/2001 VG V563 Sur création d'un enregistrement, initialisation de
                           PSA_DADSFRACTION à 1 et PSA_REGIMESS à '200'
PT20  : 12/11/2001 SB V563 Affichage du contrat fiche salarie
                           Le contrat affiché ne correspond pas au dernier
                           contrat en cours
PT21  : 12/11/2001 VG V563 Gestion des champs Fraction et prud'hommes : idem
                           etablissement
PT21-2: 29/11/2001 VG V563 Initialisation de PSA_TAUXPARTIEL, PSA_TAUXPARTSS et
                           PSA_DADSDATE sur la création + dans l'historisation,
                           contrôle que PSA_TAUXPARTIEL <> Null
PT22  : 06/12/2001 JL V563 Ajout de contrôles :
                           - Numéro de SS : année, mois, département
                           - Pays de naissance<>France : département=99
                           - Bouton MultiEmployeur désactivée si case non cochée
                           - Dates Entrée/Sortie par rapport dates debut/fin de
                             contrat
PT23  : 14/12/2001 SB V563 Suppression des blancs sur code salarie en
                           alphanumerique
PT24  : 14/01/2002 VG V571 Correction historisation des champs de la DADSU
PT25  : 07/02/2002 PH V571 suppression historisation salarié pendant la saisie
                           salarié ==> faite au niveau de chaque paie
PT26  : 22/03/2002 JL V571 Ajout procedure pour lancement de l'impression de
                           l'état de la fiche salarié
PT27-1: 22/03/2002 JL V571 Fiche Bug N° 446 : Zone Pris effectif non coché après
                           validation
PT27-2: 22/03/2002 JL V571 Fiche Bug N° 425 : postérieure au lieu de posterieur
PT28  : 22/03/2002 SB V571 Fiche de bug n°426 Controle date sortie que sur 1ère
                           sortie
PT29  : 27/03/2002 JL V571 Fiche de bug n° 333 Si code numérique, valeur
                           0000000001 par défaut si création nouveau salarié
PT30  : 02/04/2002 Ph V571 traitement champ profil retraite et fonction
                           traitement des idem etab
PT31  : 10/04/2002 JL V571 Ajout gestion incrément automatique en fonction des
                           intérimaires
PT32  : 11/04/2002 VG V571 En création de fiche salarié, la date d'ancienneté
                           doit reprendre automatiquement la date d'entrée
PT33  : 18/04/2002 SB V571 Fiche de bug n°10087 : V_PGI_env.LibDossier non
                           renseigné en Mono
PT34  : 30/04/2002 JL V571 Ajout gestion récupération intérimaire
PT35  : 10/05/2002 JL V582 Ajout vérification libellé du pays dans contrôle
                           Rechargement des tablettes TTPAYS et YYNATIONALITE
                           Déplacé le 15/09/2003 : test à la sortie du
                           département et blocage
PT36  : 13/05/2002 JL V582 Fiche de bug n° 10121 : Pas de motif de sortie si
                           date de sortie nulle ou 01/01/1900
PT37  : 13/05/2002 JL V582 Fiche de bug n° 10119 : Affichage de toutes les CCN
                           si non renseigné ds établisement
PT38  : 13/05/2002 JL V582 Fiche de bug n° 10125 : Contrôle des horaires
                           (hebdo<mensuel<annuel)
PT39  : 22/05/2002 JL V582 Contrôle du nombre d'enfants a charge a la sortie de
                           la fiche enfants (lancement de la fiche ds delphi au
                           lieu du script) et fermeture de la requête QNBCharge
PT40  : 28/05/2002 PH V582 Controle département de naissance 75 du numss <>
                           departement de naissance
                           si Annee de naissance >= 1964 et <= 1968 alors Ok
PT41  : 03/06/2002 VG V582 Version S3 : Rend invisible l'onglet règlement des
                           acomptes
PT43  : 28/06/2002 PH V582 Controle contenu des champs Horaires avnt controles
                           de cohérence
PT44  : 19/06/2002 SB V582 Controle date de sortie en fonction de la date sortie
                           paiencours
PT45  : 08/08/2002 PH V582 Controle 4 combo libres saisie obligatoire idem codes
                           org
PT46  : 05/09/2002 SB V585 Intégration de la gestion salarié de la méthode de
                           valorisation au maintien des CP
PT47  : 06/09/2002 PH V585 Prise en compte 5 elt de salaire
PT48  : 09/09/2002 VG V585 Ajout champ pour ancienneté dans le poste
PT49  : 18/09/2002 SB V585 FQ n°10022 Ajout champ Organisme à editer sur
                           bulletin
PT50  : 20/09/2002 PH V585 FQ n°10218 Imputation par defaut convention
                           collective de l'etablissement en creation
PT51  : 10/10/2002 VG V585 Fonction TestNumeroSS déplacée dans PgOutils
PT52  : 18/10/2002 PH V585 blocage de la longueur à 10c de la saisie de la date
                           de sortie
PT53  : 09/12/2002 PH V591 Controle date entree et date de sortie et paie déjà
                           faite pour autoriser la modif date de sortie FQ 10312
PT54  : 12/12/2002 SB V591 FQ 10390 Affectation ETB à PSA_TYPEDITORG en création
PT55  : 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 au
                           lieu de null
PT56  : 17/12/2002 PH V591 Passage en paramètre du nom+prenom du salarie pour
                           acces enfants salariés
PT57  : 18/12/2002 PH V591 Controle des types de données sur les setfield
PT58  : 26/12/2002 JL V591 FQ 10262 Test du numéro de sécu éffectué à la
                           validation pour ne pas avoir le message plusieurs
                           fois
PT59  : 03/01/2003 PH V591 Controles et affichage des dates dans onglet emploi
                           pour ne pas afficherv idate1900
PT60  : 03/01/2003 PH V591 FQ10417 initialisation nationalités
PT61  : 11/02/2003 JL V_42 Correction calcul personnes à charge (variable Nb non
                           initialisée)
PT62  : 19/02/2003 SB V595 FQ 10489 Faute Orthographe
PT63  : 21/02/2003 VG V_42 Correction du contrôle du code salarié
PT64  : 28/02/2003 SB V_42 FQ 10279 Gestion des contrats
PT65-1: 12/03/2003 VG V_42 Ajout de la gestion des acomptes en version S3
PT65-2: 12/03/2003 VG V_42 fiche SALARIE_BANQUE renommée en SAL_BANQUECWAS
PT66-1: 19/03/2003 VG V_42 Initialisation du champ PSA_ANCIENPOSTE en création
PT66-2: 19/03/2003 VG V_42 Chargement des valeurs de l'établissement si le champ
                           PSA_ETABLISSEMENT est renseigné
PT67  : 24/04/2003 SB V_42 Integration de la gestion des ressources
PT68  : 24/04/2003 SB V_42 Intégration de la gestion des activités
PT69  : 19/05/2003 SB V_42 FQ 10693 Ajout champ PSA_SSDECOMPTE pour calcul des
                           temps CP
PT70  : 01/07/2003 JL V_42 FQ 10652 modif test sur numéro SS
PT71  : 04/07/2003 JL V_42 suppression de PSA_DATENTREECRC,PSA_DATESORTIECRC
PT72  : 08/07/2003 SB V_42 Ajout controle date de sortie message doublon
PT73  : 04/09/2003 PH V_42 Chgt tablettes CODEEMPLOI dans le cas passage 2003
PT74-1: 15/09/2003 SB V_42 FQ 10656 Suppr. de l'affichage du 1er onglet aprés
                           validation
PT74-2: 15/09/2003 SB V_42 FQ 10781 Acces parametre CP verrouillé si non gestion
                           CP au niveau Etablissement ou parametre société
PT74-3: 15/09/2003 SB V_42 FQ 10804 Paramétrage calendrier : Anomalie ergonomie
                           dû à l'utilisation de la propriété Vide pour le champ
PT74-4: 16/09/2003 SB V_42 FQ 10758 Date d'entrée : blocage modification date
                           entrée
PT74-5: 23/09/2003 SB V_42 Portage Cwas : Suppression de la double zone
                           Ancienneté
PT75  : 23/09/2003 PH V_42 Suppression bouton compétence FQ 10075
PT76  : 25/09/2003 SB V_42 Refonte dû anomalie initialisation champ idemetab
                           suite importisis
PT77  : 25/09/2003 SB V_42 Raffraichissement des grilles RIB sur navigateur et
                           aprés saisie
PT78  : 21/10/2003 PH V_42 Inverse de FQ 10075
PT79  : 04/11/2003 PH V_50 Correction erreur QFiche FQ  10822
PT80  : 09/12/2003 SB V_50 FQ 10370 Edition compteurs Cp selon paramétrage
PT82  : 18/12/2003 VG V_50 FQ 10542 Tolérance N° SS provisoire
PT83  : 26/01/2004 JL V_50 FQ 10697 Ajout nom jeune fille dans récupération
                           intérimaires
PT84  : 09/02/2004 JL V_50 Duplication des salariés dans table INTERIMAIRES pour
                           gestion RH Compétences
PT85  : 09/02/2004 PH V_50 Recopie des date entrée et sortie uniquement en modif
                           dans les zones précédentes
PT86-1: 08/04/2004 SB V_50 FQ 11136 Ajout coche Gestion des congés payés niveau
                           salarié
PT86-2: 08/04/2004 SB V_50 FQ 11237 Ajout Gestion idem etab jours Cp Suppl
PT86-3: 16/04/2004 SB V_50 Si modif date entrée réaffectation date acquisition
                           anciénneté si date entrée
PT87  : 12/05/2004 PH V_50 FQ 10973 Restriction droit accès au salarié
PT88  : 25/05/2004 PH V_50 Suppression interimaire et compétences liés au
                           salarie
PT89  : 21/06/2004 PH V_50 FQ 11366 Affichage titre écran NOM+PRENOM
PT90  : 28/06/2004 JL V_50 FQ 10543 si condition emploi = Temps plein, alors
                           griser la zone Taux temps partiel
PT91  : 03/08/2004 JL V_50 FQ 11449 Bouton Multi-employeur inaccesible après
                           validation si multi-employeur non coché
PT92  : 09/08/2004 VG V_50 Affichage de l'ancienneté en nombre d'années et
                           nombre de mois - FQ N°11217
PT93  : 12/08/2004 PH V_50 FQ 11494 Zone catégorie bilan social afichée en
                           fonction de la séria
PT94  : 19/08/2004 PH V_50 FQ 11432 Controle contenu de la zone jour acquis
                           bulletin obligatoire
PT95  : 25/08/2004 PH Mise en place du lien avec ISOFLEX SGED en mode non PCL
                           Activation par Bouton Fiche Salarié
PT96  : 14/09/2004 PH V_50 FQ 11604 Problème profils spéciaux
PT96-1: 15/10/2004 PH V_50 FQ 11604 Problème profils spéciaux
PT97  : 30/09/2004 JL V_50 FQ 11613 Gestion salariés confidentiel : maj de
                           toutes de la confidentialité dans toutes les tables
                           du salarié
PT98  : 15/11/2004 PH V_60 FQ 11640 Prise en compte de la date de sortie en mode
                           création
PT99  : 15/11/2004 PH V_60 FQ 11640 Contrôle effacement du champ régul
                           ancienneté
PT100 : 13/12/2004 JL V_60 FQ 11735 Département corse forcé en majuscule (2A,2B)
PT101 : 15/12/2004 JL V_60 FQ 11815 Contrôle format département (2 chiffres, et
                           pas d'espace)
PT101-1  : 17/02/2005 JL V_60 Test espace dans  numéro SS
PT102 : 14/01/2004 SB V_60 FQ 11353 Suppression Cp si décoche gestion CP
PT103 : 25/01/2004 PH V_601 FQ 11950 Gestion de la confidentialité en CWAS
PT104 : 09/02/2005 JL V_60 FQ 11841 Dupliquer fiche salarié
PT105 : 10/05/2005 PH V_60 FQ 12120 Ergonomie
PT106 : 19/05/2005 SB V_60 Remplacement des messages LastErrorMsg en PgiBox
PT107 : 10/06/2005 PH V_60 FQ 12374 Message est Num sécu incomplet.
PT108 : 10/06/2005 SB V_60 CWAS : Chargement tablette compatibilite CWAS
PT109 : 12/07/2005 PH V_60 Valeur par defaut établissement en cas de
                           confidentialité par étab
PT110 : 02/08/2005 PH V_60 FQ 12443 Ne pas renseigner des valeurs de champs CP
                           si le champ n'est pas dans la fiche
PT111 : 08/08/2005 PH V_60 FQ 11432 Controle de la variable saisie nbre jours CP
                           acquis Suite PT94
PT112 : 20/09/2005 SB V_65 FQ 12531 Correction suite modif PT110
PT113 : 10/10/2005 VG V_60 Initialisation PSA_PERSACHARGE à 99 - FQ N°11917
PT114 : 13/10/2005 SB V_65 initialisation PSA_CATBILAN à 000
PT115 : 13/10/2005 PH V_65 FQ 11908 Creation et modification contrat automatique
PT116 : 14/10/2005 PH V_65 En PCL et base de prod optimisée pas d'accès aux
                           ressources
PT117 : 20/10/2005 SB V_65 FQ 12068
PT118 : 09/12/2005 PH V_65 FQ 12676 Controle date d'entrée
PT119 : 12/12/2005 VG V_65 Correction de la modification du nombre de personnes
                           à charge - FQ N°12749
PT120 : 18/01/2006 JL V_65 FQ 12828 La duplication d'une fiche salarié doit
                           aussi dupliquer la gestion des CP.
PT121 : 18/01/2006 JL V_65 FQ 11841 récupéreration profil des Cas particuliers
                           (grille non rechargé si duplication)
PT122 : 06/02/2006 EP V_65 FQ 12858 pas de maj automatique de ventilation
                           analytique si personnalisée
PT123 : 09/02/2006 EP V_65 FQ 12855 en création de salarié, proposer
                           l'établissement par défaut
PT124 : 13/02/2006 SB V_65 FQ 12451 Ajout option edition bulletin salarié par
                           défaut
PT125 : 06/03/2006 PH V_65 FQ 12668 Gestion du matricule origine BL(EWS)
PT126 : 03/04/2006 SB V_65 FQ 11426 Mis en place paramètre pour calcul au
                           maintien CP
PT127 : 06/04/2006 PH V_65 FQ 13054 Erreur type du champ
PT128 : 07/04/2006 SB V_65 FQ 12451 Ajout btn bulletin par défaut
PT129 : 26/04/2006 JL V_65 FQ 13109 Ajout alimentation PSA_TYPCONVENTION  en
                           dupli fiche salarié
PT130 : 09/05/2006 PH V_65 FQ 13124 Accès ttes conventions si besoin de + 3 CCN
PT131 : 15/05/2006 PH V_65 FQ 12806 gestion des zones obligatoires
PT132 : 19/05/2006 SB V_65 FQ 13167 Ajout param. salaire maintien rep PT126
PT133 : 19/05/2006 SB V_65 FQ 13130 Ajout contrôle pour acces btn en création
PT134 : 08/06/2006 PH V_65 FQ 12377 Bouton analytique non visible si analytique
                           non gérée
PT135 : 07/07/2006 SB V_65 FQ 12068 Reprise calcul ancienneté onglet emploi
PT136 : 21/07/2006 PH V_65 traitement des dates pour avoir un masque de saisie
PT137 : 28/07/2006 PH V_65 FQ 13410 Creation salarie avec récup gestion cp
                           salarie
PT138 : 25/08/2006 VG V_70 Adaptation cahier des charges DADS-U V08R04
PT139 : 28/09/2006 VG V_70 Adaptation cahier des charges DADS-U V08R04
PT140 : 19/10/2006 SB V_70 FQ 13324 Anomalie duplication champ CP
PT141 : 23/10/2006 VG V_70 Contrôle Temps partiel et taux temps partiel
                           FQ 13502
PT141B: 25/10/2006 GGS V_70 Désactivation des champs afferents à la
                           carte de séjour si nationalité d'un pays européen
PT142 : 13/12/2006 PH V_70 Mise en place GED Dossier salarié
PT143 : 13/12/2006 JL V_80 Ajout gestion historique daté
PT144 : 09/01/2007 FCO V_80 Possibilité de sélectionner 00:00 dans heure
                           d'embauche
PT145 : 10/01/2007 FCO V_80 Rajout de la date prévisionnelle de prochaine visite
                           médicale
PT146 : 25/01/2007 SB V_70 FQ 11992 Ajout valeur methode date d'acquisition CP
PT147 : 25/01/2007 SB V_70 Contrôle cohérence saisie date si modif. paramètres
                           régionaux PC en Anglais
PT148 : 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT149 : 21/03/2007 FC V_70 FQ 13286 Ne pas pouvoir sélectionner un établissement
                           sans données complémentaire
PT150 : 31/03/2007 GGS V_80 FQ 14010 Affichage onglet emploi si on vient de
                           processus type sortie de salarie
PT151 : 18/04/2007 GGU V_72 Control de l'influance des champs modifiés sur les
                           populations Modifié le 03/05/2007
PT152 : 04/05/2007 MF V_72 FQ 14108 : contrôle département de naissance
                           renseigné
PT153 : 04/05/2007 MF V_72 FQ 14109 : correction heure d'embauche non renseignée
PT154 : 16/05/2007 GGU V_72 Rends obligatoire les champs utilisés dans le
                           paramétrage des population (le forçage reste
                           possible)
PT155-1: 22/05/2007 VG V_72 Intégration du planning unifié
PT155-2: 22/05/2007 VG V_72 Déplacement de la fonction TestNumeroSSNaissance
PT156  : 29/05/2007 JL V_72 Modification saisie des cas particuliers
PT157  : 31/05/2007 MF V_72 FQ 14383
PT158  : 01/06/2007 FC V_72 FQ 12925
PT159  : 05/06/2007 JL V_72 FQ 14125 Correction affichage libellés zones libres
                            correction le 20/09/2007
PT160  : 05/06/2007 JL V_72 FQ 14084 gestion des zones obligatoires, PT131
                            déplacé en fin du OnUpdate
PT161  : 26/06/2007 JL V_72 FQ 14312 Suppression article dans message zone
                            obligatoires
PT162-1: 27/06/2007 VG V_72 Initialisation en création - FQ N°13324
PT162-2: 27/06/2007 VG V_72 Correction PT129
PT162-3: 27/06/2007 VG V_72 Procedure pour mettre à jour les tiers sur
                            modification du salarié - FQ N°13476
PT162-4: 27/06/2007 VG V_72 Optimisation
PT162-5: 27/06/2007 VG V_72 Suppression des données de la DADS-U lors de la
                            suppression du salarié - FQ N°13868
PT162-6: 27/06/2007 VG V_72 Création du tiers même si anomalie lors de la
                            création - FQ N°14285
PT163  : 10/07/2007 FC V_72 FQ 14534
PT164  : 19/07/2007 FC V_72 FQ 13022 DADS-U Gestion des différents régimes
PT165  : 23/07/2007 FC V_72 FQ 14589 Erreur à la consultation d'un élément
                            national de niveau salarié
PT166  : 31/07/2007 JL V_80 FQ 13683 Ajout duplication champs libres
PT167  : 02/08/2007 JL V_80 FQ 14193 Correction affichage titre grille RIB
PT168  : 06/08/2007 JL V_80 FQ 14167 Gestion concept modif salarié/historique
PT169  : 10/08/2007 GGU V_80 mise de PGPopulOutils dans un ifndef
PT170  : 24/08/2007 JL V_80 FQ 14271 Déplacement gestion boputon contrat dans
                            éxé pour gérer rafraichissement de l'affichage
PT171  : 07/09/2007 VG V_80 Avec la ribbon-bar l'accès à l'attestation ASSEDIC
                            créait une attestation à chaque accès - FQ N°14377
PT172  : 07/09/2007 JL V_80 Correction erreur lors de l'accès aux rib, CP en PCL
PT173  : 10/09/2007 JL V_80 FQ 14706 Prise en compte valeur <<Aucun>> dans
                            historique
PT174  : 25/09/2007 FC V_80 FQ 14806 Afficher la première valeur postérieure à
                            la date du jour d'un élément dynamique
PT175  : 28/09/2007 VG V_80 MemCheck
PT176  : 03/10/2007 FC V_80 Concept salarié
PT177  : 06/11/2007 FC V_80 FQ 14505 En CWAS, cacher les boutons pagination dans
                            l'accès rapide
PT178  : 16/11/2007 FC V_80 FQ 14947 Si création salarié, impossible de saisir
                            des éléments dossier
PT179  : 05/12/2007 VG V_80 Affichage du caractère interdit - FQ N°14961
PT180  : 05/12/2007 FL V_81 Optimisation + correction warnings
PT181  : 11/12/2007 FC V_81 Duplication salarié : la catégorie DUCS et la
                            catégorie Bilan social n'était pas reprises
PT182  : 14/12/2007 FC V_81 Correction bug : Quand on ne crée pas le salarié car
                            erreur, on crée quand même le tiers
PT183  : 19/12/2007 FC V_81 FQ 15064 libellés coeff, qualif, niveau et indice
PT184  : 14/12/2007 FC V_81 FQ 14996 Nouveaux concepts de visibilité des onglets
                            salarié
PT185  : 03/01/2008 FC V_81 FQ 14762 Journal des évènements
PT186  : 07/01/2008 FC V_81 FQ 14872 Rajouter la mise à jour motif fin de
                            contrat du contrat
PT187  : 07/01/2008 FC V_81 FQ 13938 Modifier automatiquement la date de fin
                            travail effectif dans le contrat de travail
PT188  : 09/01/2008 GGU V_81 FQ 14268 Profil / détail des rubriques associées
PT189  : 23/01/2008 GGU V_81 Lors de l'affichage des banques du salarié en vision
                            SAV, on a des messages
PT190  : 29/01/2008 GGU V_81 FQ 15183 Paramètres cp, validation de fiche : " La
                            ligne n'a pu étre trouvée pour la mise à jour. "
PT191  : 06/02/2008 FC V_81 Les appels Banque et CP ont été remis dans le script
PT192  : 04/03/2008 FC V_90 FQ 15270 historiser la convention collective comme
                            les autres champs historisés par avance
PT194  : 07/04/2008 VG V_81 Ajout du cas où la table intérimaires est partagée
                            entre plusieurs dossiers, auquel cas on vérifie les
                            doublons
PT197  : 18/04/2008 GGU V_81 FQ 15361 Gestion uniformisée des zones libres
                            tables dynamiques
PT203  : 15/05/2008 FLO V_81 FQ 15371 Lorsqu'une date de sortie est positionnée,
                            annulation des formations en cours
PT204  : 23/05/2008 GGU V_81 FQ 15465 elements dynamiques : violation d'accès
                            dans la liste des non saisies
PT214  : 03/07/2008 PH V_85 FQ 15599 Rajout initialisation du champ
                            traitementok = '-'
PT215  : 04/07/2008 VG V_82 Retour en arrière d'urgence du PT134. Même si
                            l'analytique n'est pas géré, on autorise le
                            paramétrage de l'analytique dans la fiche du salarié
                            FQ N°15600
PT216  : 30/09/2008 JPP FQ13328 Test que le taux d'abattement soit positif si le
                        profil d'abattement est renseigné
PT217  : 02/10/2008 JS FQ n°15431 Uniformisation de la saisie code postal et de
                       l'affichage de la ville
PT218  : 06/10/2008 VG Conditionner la question sur la mutation du salarié par
                       le concept "Modification des éléments standards"
PT219 : 13/10/2008 JS      FQ n°15288 Création automatique du compte tiers lorsque le matricule est alpha-numérique
}
unit UTOMSalaries;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  StdCtrls, Controls, Classes, sysutils, ComCtrls, lookup,
// CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  forms,
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
   CPCODEPOSTAL_TOF,//PT217
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, DBCtrls, Fe_Main, Fiche,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
{$ENDIF}
{$ELSE}
  MaineAgl, eFiche, UtileAgl,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97, FichComm, Spin, HSysMenu,
  extctrls, ParamSoc, UtilPGI, Pgoutils, PGOutils2, ParamDat, EntPaie, graphics,
// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}
  PgDADSControles,
  PgCongesPayes,
  URibbonBar,
  PgoutilsHistorique,
  utofrechgedpaie, // PT142
// Gestion des profils
  utofPgsalarieRub,
  PGCalendrier, // PT145
  PgOutilsTreso,
  RubriqueProfil,//PT188
{$ENDIF}
  UtilPaieAffaire,
// FIN CCMX-CEGID ORGANIGRAMME DA
  ed_tools, Windows, Grids, Commun, Menus,
  YRessource, ULibEditionPaie
{$IFDEF PAIEGRH}
  ,PAIETOM   //PT185
{$ENDIF PAIEGRH}
 ;
// HQuickRP,fichlist,

type
{$IFDEF PAIEGRH}
  TOM_SALARIES = class(PGTOM)  //PT185
{$ELSE}
  TOM_SALARIES = class(TOM)
{$ENDIF PAIEGRH}
  public
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnClose; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnBeforeUpdateRecord; override;
    procedure ExitEdit(Sender: TObject);  //PT217

  // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
    procedure OnCancelRecord; override;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  private
    //  PT9 : 18/09/2001 V547 PH Rajout TOB axes et sections analytiques
    TOB_Axes, TOB_Section: TOB;
    InitEtab, numss, codesexe, MesError, TypeCalend, GblCalend: string; //PT64
    DepartNaissanceSS, MoisSS, AnneeSS: string;
    resultat, ResultAnnee, ResultMois, ResultDepart: Integer;
    reponse: Word;
    Tob_Etab, T, Tob_DonneurOrdre, TDO: TOB;
    DTClot, GblDtEntree, GblDtSortie: tdatetime;
    GblMotifSortie:String;//PT186
    EtbCP, FirstCp, Majok, ChangeEtab, Creation: boolean; //PT74-2
    Argument, Interimaire: string; //PT34
    GblCpCoche: Boolean; { PT102 }
    SalarieDupliquer: string; //PT104
    CreatEnCours: Boolean; // PT115
    NumOrdre: Integer;
    HeureEmbauche: string; //PT144
    PeriodVisitDefaut: string; // PT145
    argappel: string; //PT150
    CreationTier:Boolean; //PT182
    {$IFNDEF GCGC}
    NonGed:boolean;//PT184
    {$ENDIF !GCGC}
{$IFDEF PAIEGRH}
    Trace: TStringList; //PT185
    DerniereCreate: string; //PT185
    LeStatut:TDataSetState; //PT185
{$ENDIF PAIEGRH}
{$IFNDEF GCGC}
    OngletsRibbonBar: array[0..6] of string;
    CurProfil : String; //PT188
{$ENDIF !GCGC}
    CouleurHisto: Integer;
    GereHistoConvention:Boolean;  //PT192
    EstSalarieMSA, EstSalarieSpectacle: boolean;
//    SavTableFields, SavTableValues : array of string;//PT151
    // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
    TobZonesCompl, TobZonesComplSauv: TOB;
    procedure ZonesComplIsModified;
    procedure Controle_SalarieDA;
{$ENDIF}
    // FINCCMX-CEGID ORGANIGRAMME DA - 13.06.2006
//    procedure PCasRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: boolean);
//    procedure PCasOnClick(Sender: TObject);
    procedure MesErrorSecu(resultat: Integer);
//    function TestNumeroSSNaissance(numss: string; SSNaiss: string): integer; // PT22
    procedure MesErrorSecuNaissance(resultAnnee, ResultMois, ResultDepart: Integer); // PT22
    procedure AffectCodeNewEnr;
    procedure MajChampIdemEtab;
    procedure AffectChampIdemEtab(TFils: TOB; StType, ChampType, ChampVal, ChampEtab: string; ValSoc: Variant);
    procedure EnabledZonesAnciennete;
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
    procedure AlimentecumulConge;
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    procedure AffichePhoto;
    procedure RendVisibleOrg;
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
    procedure BTCongesPayesClick(Sender: TObject);
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    procedure FRibDblClick(Sender: TObject);
    //    procedure BasanccpChange(Sender: TObject);
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
    procedure TCalendrierClick(Sender: TObject); //button calendrier
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    procedure TauxClick(Sender: TObject);
    procedure ConvClick(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure AffectDateAcqCpAnc(ValTypDatAnc: string);
    procedure OnChangeStandCalend(Sender: TObject);
    procedure ClickMultiEmpl(Sender: TObject); // PT22
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
    procedure ImpressionSal(Sender: TObject); //PT26
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    // PT30 02/04/2002 V571 Ph traitement champ profil retraite
    procedure IdemEtab(ChampTyp, LeChamp, ChampEtab: string; F: TField);
    procedure ControleHoraires(Hebdo, Mensuel, Annuel: Double); //PT38
    function TestHoraire(HoraireInf, HoraireSup: Double): Boolean;
    procedure BEnfantsClick(Sender: TObject);
    procedure ControleDate(LeChamp, LeBouton: string);
    procedure BtnDate(LeBouton: string);
    procedure DateActivee(Sender: TObject);
    procedure ChargeProfilActivite(LeChamp, Nat: string); //PT68
    procedure OnEnterProfil(Sender: Tobject); //PT68
    procedure ChargeGrilleRib(Grid: THGrid; StType: string); //PT77
    {$IFDEF PAIEGRH}
    procedure DupliquerSalInterimaire; //PT84
    {$ENDIF PAIEGRH}
    function GereIsoflex: Boolean;
    procedure AppelIsoflex(Sender: TObject);
    procedure MajConfidentialite; //PT97
    procedure OnExitRegul(Sender: TObject); //PT99
    // Gestion coche confidentiel -> PSA_CONFIDENTIEL
    procedure OnClickConfidentiel(Sender: TObject);
    procedure DuplicationSalarie(leSalarie: string);
    procedure BDupliqueClick(Sender: Tobject);
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
    procedure BTNClick(Sender: TObject);
    procedure PBTNClick(Sender: TObject);
    procedure RempliOnglet(Nom: string);
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    procedure PersAChargeExit(Sender: TObject);
    procedure ClickBtnBulDefaut(Sender: TObject); { PT128 }
    procedure PCLConvClick(Sender: TObject); // PT130
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState); // PT136
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 11.12.2006
{$IFNDEF GCGC}
    //gestion des profils
    procedure FicheKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 11.12.2006
{$IFNDEF GCGC}
    procedure AccesHistoEdit(Sender: TObject);
    procedure AccesHistoCombo(Sender: TObject);
    procedure AccesHistoCheck(Sender: TObject);
    //procedure HistorisationDonnees;
    function CreerBoutonHisto(THBST: TTabsheet; Nom: string; L, W, T: Integer): TToolBarButton97;
    procedure RendAccesChampHisto(Acces, Maj: Boolean);
    procedure AfficheEltCompl(ForcerModif: Boolean = False);
    procedure AfficheEltNat(AjoutElt: Boolean = False);
    procedure GrilleEltComplClick(Sender: TObject);
    procedure GrilleEltNatDos(Sender: TObject);
    procedure GrilleGetCellCanvas(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure ChangeComboEltComp(Sender: Tobject);
    procedure ClickBtEltCompl(Sender: Tobject);
    procedure AccesMulsZonesLibres(Sender: Tobject);
    procedure BeforeExec(const ControlName: string; var Cancel: boolean);
    procedure AccesRibbonBar(Sender: Tobject);
{$ENDIF}
    //gestion des profils
{$IFNDEF GCGC}
    function GetInfoSal: string; // PT142
    procedure bGedOnClick(Sender: Tobject);
    procedure GerebGed;
{$ENDIF}
    procedure ChangementHeure(Sender: TObject); // PT144
    procedure CasParticulierDblClick(Sender: TObject);
    procedure BtnContratClick(Sender: Tobject); //PT170
    procedure AfficheContrat; //PT170
//    Procedure SavFieldsPopul; //PT151
//    function  ValidUpdateFields : Boolean; //PT151
    //DEB PT183
    procedure CoeffElipsisclick(Sender: Tobject);
    procedure QualifElipsisclick(Sender: Tobject);
    procedure IndiceElipsisclick(Sender: Tobject);
    procedure NiveauElipsisclick(Sender: Tobject);
    function ClauseSQL(Nature:String):String;
    //FIN PT183
    //DEB PT184
    {$IFNDEF GCGC}
//PT191    procedure BBANQUE_OnClick(Sender: TObject);
    procedure BMEMO_OnClick(Sender: TObject);
//PT191    procedure BCP_OnClick(Sender: TObject);
    procedure BtnSALRUBClick(Sender: Tobject);
    procedure AccesVentil(Sender: Tobject);
    {$ENDIF !GCGC}
    //FIN PT184
//Début PT188
{$IFNDEF GCGC}
    procedure ShowRubAssocProfil(Sender: TObject);
    procedure OnContextMenuPROFILTPS   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFIL      (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILREM   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPERIODEBUL  (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILRBS   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuREDRTT2     (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuREDRTT1     (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuREDREPAS    (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILAFP   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILAPP   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILRET   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILMUT   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILPRE   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILTSS   (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILFNAL  (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPROFILTRANS (Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure OnContextMenuPPROFILSUITE(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
{$ENDIF}
//Fin PT188
  end;

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 16.06.2006
{$IFDEF GCGC}
function RechSalarieModuleDA(Lesalarie: string): Boolean;
function SalarieDansCirCuitDA_HierarchieInDirecte(LeSalarie: string): Boolean;
function SalarieDansCirCuitDA_HierarchieDirecte(LeSalarie: string): Boolean;
function SalarieDansCirCuitDA_Salarie(LeSalarie: string): Boolean;
function DA_aValiderParSalarie(Lesalarie: string): Boolean;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 16.06.2006
// GED Salarie PT142
{$IFNDEF GCGC}
function Salarie_MyAfterImport(Sender: TObject): string;
procedure Salarie_GestionBoutonGED(Sender: TObject);
{$ENDIF}

implementation
uses AglIsoflex,
  TiersUtil,
  HPanel,
  P5DEF,
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  PGPOPULOUTILS, // PT151   PT169
  PgoutilsFormation,
  P5Util,
{$ELSE}
  M3FP,
{$ENDIF}
{$IFDEF EMANAGER}
  uGedFiles,
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  PGCommun, StrUtils,
  Ventil//PT184
  , PGTablesDyna //PT197
  , MailOL //PT203
{$IFDEF STATDIR}
, DPMajPaieOutils    // BTY FQ 15526
{$ENDIF}
  ;

{ TOM_SALARIES }

//Debut PT188
{$IFNDEF GCGC}
procedure TOM_SALARIES.OnContextMenuPROFILTPS(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILTPS';
end;

procedure TOM_SALARIES.OnContextMenuPROFILREM(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILREM';
end;

procedure TOM_SALARIES.OnContextMenuPROFIL(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFIL';
end;

procedure TOM_SALARIES.OnContextMenuPERIODEBUL(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PERIODBUL';
end;

procedure TOM_SALARIES.OnContextMenuPROFILAFP(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILAFP';
end;

procedure TOM_SALARIES.OnContextMenuPROFILAPP(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILAPP';
end;

procedure TOM_SALARIES.OnContextMenuPROFILFNAL(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILFNAL';
end;

procedure TOM_SALARIES.OnContextMenuPROFILMUT(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILMUT';
end;

procedure TOM_SALARIES.OnContextMenuPROFILPRE(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILPRE';
end;

procedure TOM_SALARIES.OnContextMenuPROFILRBS(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILRBS';
end;

procedure TOM_SALARIES.OnContextMenuPROFILRET(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILRET';
end;

procedure TOM_SALARIES.OnContextMenuPROFILTRANS(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILTRANS';
end;

procedure TOM_SALARIES.OnContextMenuPROFILTSS(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'PROFILTSS';
end;

procedure TOM_SALARIES.OnContextMenuREDREPAS(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'REDREPAS';
end;

procedure TOM_SALARIES.OnContextMenuREDRTT1(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'REDRTT1';
end;

procedure TOM_SALARIES.OnContextMenuREDRTT2(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
begin
  CurProfil := 'REDRTT2';
end;

procedure TOM_SALARIES.OnContextMenuPPROFILSUITE(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
var
  TmpControl, TmpControlSansHisto : TControl;
  tmpname : String;
  ControlSansHisto : {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF};
  CallHandled : Boolean;
begin
  if not(Sender is TTabSheet) and not(Sender is THPanel) then exit;
  if Sender is TWinControl then
    TmpControl := (Sender as TWinControl).ControlAtPos(MousePos, True)
  else
    TmpControl := nil;
  if Assigned(TmpControl) then
  begin
    tmpname := TmpControl.Name;
    if LeftStr(tmpname, 5) = 'BPSA_' then
    begin
      TmpControlSansHisto := GetControl(RightStr(tmpname, Length(tmpname)-1));
      if not Assigned(TmpControlSansHisto) then exit;
      if TmpControlSansHisto is {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF} then
      begin
        ControlSansHisto := (TmpControlSansHisto as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF});
        if not Assigned(ControlSansHisto.PopupMenu) then exit;
        if not Assigned(ControlSansHisto.OnContextPopup) then exit;
        if (Sender is TTabSheet) then
        begin
          ControlSansHisto.OnContextPopup(Sender, MousePos, CallHandled);
          (Sender as TTabSheet).PopupMenu := ControlSansHisto.PopupMenu;
          (Sender as TTabSheet).PopupMenu.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
          (Sender as TTabSheet).PopupMenu := nil;
        end else if (Sender is THPanel) then begin
          ControlSansHisto.OnContextPopup(Sender, MousePos, CallHandled);
          (Sender as THPanel).PopupMenu := ControlSansHisto.PopupMenu;
          (Sender as THPanel).PopupMenu.Popup(Mouse.CursorPos.X,Mouse.CursorPos.Y);
          (Sender as THPanel).PopupMenu := nil;
        end;
        Handled := True;
      end;
    end;
  end;
end;

procedure TOM_SALARIES.ShowRubAssocProfil(Sender : TObject);
var
  PROFIL_CODE, PROFIL_LIBELLE, PROFIL_PREDEFINI : String;
  Q : TQuery;
begin
  //On récupère le profil sélectionné pour le salarié
  PROFIL_CODE := GetField('PSA_'+CurProfil);
  if Assigned(GetControl('PSA_'+CurProfil)) then
    PROFIL_LIBELLE := RechDom((GetControl('PSA_'+CurProfil) as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).DataType
                             ,PROFIL_CODE
                             ,False);
  Q := OpenSQL('Select PPI_PREDEFINI from PROFILPAIE WHERE PPI_PROFIL = "'+PROFIL_CODE+'"', TRUE);
  While not Q.Eof do
  begin
    if Q.Fields[0].AsString = 'DOS' then
    begin
      PROFIL_PREDEFINI := 'DOS';
      Break;
    end;
    if Q.Fields[0].AsString = 'STD' then
      PROFIL_PREDEFINI := 'STD';
    if (Q.Fields[0].AsString = 'CEG') AND (PROFIL_PREDEFINI <> 'STD') then
      PROFIL_PREDEFINI := 'CEG';
    Q.Next;
  end;
  Ferme(Q);
  //On affiche les rubriques associées
  ProfilPaieDetail(PROFIL_CODE, PROFIL_LIBELLE, PROFIL_PREDEFINI);
end;
{$ENDIF} 
//Fin PT188

procedure TOM_SALARIES.EnabledZonesAnciennete;
var
  StCpAcqAnc, StBasAncCp, SpDatAnc, SpTypDatAnc: string;
begin
  StCpAcqAnc := GetField('PSA_CPACQUISANC');
  StBasAncCp := GetField('PSA_BASANCCP');
  { DEB PT74-5 }
  SetControlProperty('PSA_VALANCCP', 'ElipsisButton', False);
  SetControlProperty('PSA_VALANCCP', 'Datatype', '');
  if GetField('PSA_BASANCCP') = 'VAR' then
  begin
    SetControlProperty('PSA_VALANCCP', 'Datatype', 'PGVARIABLE');
    SetControlProperty('PSA_VALANCCP', 'ElipsisButton', True);
  end
  else
    if GetField('PSA_BASANCCP') = 'TAB' then
    begin
      SetControlProperty('PSA_VALANCCP', 'Datatype', 'PGTABINTANC');
      SetControlProperty('PSA_VALANCCP', 'ElipsisButton', True);
    end
    else
      if GetField('PSA_BASANCCP') = 'VAL' then
        SetControlProperty('PSA_VALANCCP', 'Datatype', 'FTFLOAT');
  //  SetControlVisible ('EPSA_VALANCCP', (StBasAncCp = 'VAL'));
  //  SetControlVisible ('CPSA_VALANCCP', (StBasAncCp <> 'VAL'));
    { FIN PT74-5 }

    // SetControlEnabled('PSA_CPACQUISMOIS',profil.value<>''); // PT17
    // SetControlEnabled('PSA_DATANC',profil.value<>'');       // PT17
  {      SetControlEnabled('PSA_NBREACQUISCP',profil.value<>'');}
  //      SetControlEnabled('PSA_CPACQUISANC',profil.value<>'');
  SetControlEnabled('PSA_VALANCCP', ((StBasAncCp <> '') and (StCpAcqAnc <> 'ETB')) and (GetField('PSA_PROFILCGE') <> '')); { PT74-5 }
  SetControlEnabled('PSA_BASANCCP', (StCpAcqAnc <> 'ETB') and (GetField('PSA_PROFILCGE') <> ''));
  SpDatAnc := GetField('PSA_DATANC');
  SpTypDatAnc := GetField('PSA_TYPDATANC');
  SetControlEnabled('PSA_TYPDATANC', (SpDatAnc <> 'ETB')); // PT17 Retrait condition : and(profil.value<>'')
  SetControlEnabled('PSA_DATEACQCPANC', (SpDatAnc <> 'ETB') and (SpTypDatAnc <> '') and (SpTypDatAnc = '2')); { PT146 }
  //PT17 Retrait condition : and(profil.value<>'')
end;

//*******************PROCEDURE On Argument***********************//

procedure TOM_SALARIES.OnArgument(stArgument: string);
var
  Themes: THValCombobox;
  PCas: THGrid;
//  PPart: THValCombobox;
  i: Integer;
  st: string;
  Q: TQuery;
  THHeure: THEdit;
{$IFNDEF EAGLCLIENT}
  FListeRib: THMulDbGrid;
  MCombo, StandCalend, PRet: THDBValCombobox;
  TauxAt, Edit, ConvCol: THDBEdit;
  MultiEmpl: TDBCheckBox; // PT22
{$IFNDEF GCGC}
  Check: TDBCheckBox;
{$ENDIF !GCGC}
{$ELSE}
  FListeRib: THGrid;
  MCombo, (*Basanccp,*) StandCalend, PRet: THValCombobox;
  TauxAt, Edit, ConvCol: THEdit;
  MultiEmpl: TCheckBox; // PT22
{$IFNDEF GCGC}
  Check: TCheckBox;
{$ENDIF !GCGC}
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  MSpin: THDBSpinEdit;
{$ELSE}
  MSpin: THSpinEdit;
{$ENDIF}

  Btn: TToolBarButton97;
  Zone: Tcontrol;
  Action, TAction: string; //PT34
{$IFNDEF GCGC}
  LePanel: THPanel;
  Cpt: Integer;
{$ENDIF}
  PersACharge: THEdit;
{$IFNDEF GCGC}
  MenuPopProfil : TPopUpMenu;//PT188
  Ribbon: THRibbonBarBuilder;
  Grille: THGrid;
  TobParamSal: Tob;
  Combo: THValCombobox;
  LeControl: TToolBarButton97;
  UnControl: TComponent;
  LeChamp : string;
  MenuPop: TPopUpMenu;
{$ENDIF !GCGC}
begin
  inherited;
{$IFNDEF GCGC}
//Début PT188
  if Assigned(GetControl('PSA_PROFILTPS')     ) then
    (GetControl('PSA_PROFILTPS')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILTPS;
  if Assigned(GetControl('PSA_TYPPROFILREM')  ) then
    (GetControl('PSA_TYPPROFILREM')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILREM;
  if Assigned(GetControl('PSA_PROFILREM')     ) then
    (GetControl('PSA_PROFILREM')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILREM;
  if Assigned(GetControl('PSA_TYPPROFIL')     ) then
    (GetControl('PSA_TYPPROFIL')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFIL;
  if Assigned(GetControl('PSA_PROFIL')        ) then
    (GetControl('PSA_PROFIL')         as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFIL;
  if Assigned(GetControl('PSA_TYPPERIODEBUL') ) then
    (GetControl('PSA_TYPPERIODEBUL')  as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPERIODEBUL;
  if Assigned(GetControl('PSA_PERIODBUL')     ) then
    (GetControl('PSA_PERIODBUL')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPERIODEBUL;
  if Assigned(GetControl('PSA_TYPPROFILRBS')  ) then
    (GetControl('PSA_TYPPROFILRBS')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILRBS;
  if Assigned(GetControl('PSA_PROFILRBS')     ) then
    (GetControl('PSA_PROFILRBS')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILRBS;
  if Assigned(GetControl('PSA_TYPREDRTT2')    ) then
    (GetControl('PSA_TYPREDRTT2')     as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDRTT2;
  if Assigned(GetControl('PSA_REDRTT2')       ) then
    (GetControl('PSA_REDRTT2')        as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDRTT2;
  if Assigned(GetControl('PSA_TYPREDRTT1')    ) then
    (GetControl('PSA_TYPREDRTT1')     as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDRTT1;
  if Assigned(GetControl('PSA_REDRTT1')       ) then
    (GetControl('PSA_REDRTT1')        as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDRTT1;
  if Assigned(GetControl('PSA_TYPREDREPAS')   ) then
    (GetControl('PSA_TYPREDREPAS')    as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDREPAS;
  if Assigned(GetControl('PSA_REDREPAS')      ) then
    (GetControl('PSA_REDREPAS')       as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuREDREPAS;
  if Assigned(GetControl('PSA_TYPPROFILAFP')  ) then
    (GetControl('PSA_TYPPROFILAFP')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILAFP;
  if Assigned(GetControl('PSA_PROFILAFP')     ) then
    (GetControl('PSA_PROFILAFP')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILAFP;
  if Assigned(GetControl('PSA_TYPPROFILAPP')  ) then
    (GetControl('PSA_TYPPROFILAPP')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILAPP;
  if Assigned(GetControl('PSA_PROFILAPP')     ) then
    (GetControl('PSA_PROFILAPP')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILAPP;
  if Assigned(GetControl('PSA_TYPPROFILRET')  ) then
    (GetControl('PSA_TYPPROFILRET')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILRET;
  if Assigned(GetControl('PSA_PROFILRET')     ) then
    (GetControl('PSA_PROFILRET')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILRET;
  if Assigned(GetControl('PSA_TYPPROFILMUT')  ) then
    (GetControl('PSA_TYPPROFILMUT')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILMUT;
  if Assigned(GetControl('PSA_PROFILMUT')     ) then
    (GetControl('PSA_PROFILMUT')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILMUT;
  if Assigned(GetControl('PSA_TYPPROFILPRE')  ) then
    (GetControl('PSA_TYPPROFILPRE')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILPRE;
  if Assigned(GetControl('PSA_PROFILPRE')     ) then
    (GetControl('PSA_PROFILPRE')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILPRE;
  if Assigned(GetControl('PSA_TYPPROFILTSS')  ) then
    (GetControl('PSA_TYPPROFILTSS')   as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILTSS;
  if Assigned(GetControl('PSA_PROFILTSS')     ) then
    (GetControl('PSA_PROFILTSS')      as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILTSS;
  if Assigned(GetControl('PSA_TYPPROFILFNAL') ) then
    (GetControl('PSA_TYPPROFILFNAL')  as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILFNAL;
  if Assigned(GetControl('PSA_PROFILFNAL')    ) then
    (GetControl('PSA_PROFILFNAL')     as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILFNAL;
  if Assigned(GetControl('PSA_TYPPROFILTRANS')) then
    (GetControl('PSA_TYPPROFILTRANS') as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILTRANS;
  if Assigned(GetControl('PSA_PROFILTRANS')   ) then
    (GetControl('PSA_PROFILTRANS')    as {$IFNDEF EAGLCLIENT}THDBValComboBox{$ELSE}THValComboBox{$ENDIF}).OnContextPopup := OnContextMenuPROFILTRANS;
  if Assigned(GetControl('PPROFILSUITE')      ) then
    (GetControl('PPROFILSUITE')       as TTabSheet).OnContextPopup := OnContextMenuPPROFILSUITE;
  if Assigned(GetControl('PANELPROFILSAL')    ) then
    (GetControl('PANELPROFILSAL')     as THPanel).OnContextPopup := OnContextMenuPPROFILSUITE;
  MenuPopProfil := TPopUpMenu(GetControl('POPUPPROFIL'));
  if MenuPopProfil <> nil then
  begin
    MenuPopProfil.Items[0].OnClick := ShowRubAssocProfil;
  end;
//Fin PT188
  CouleurHisto := AltColors[V_PGI.NumAltCol];
  if V_PGI.RibbonBar then
  begin
    SetControlVisible('BGED', False);
    SetControlVisible('RIB__BELTDOSPLUS', False);
    Ribbon := THRibbonBarBuilder(GetControl('FE__HRibbonBarBuilder1'));
    for i := 0 to 6 do
    begin
      OngletsRibbonBar[i] := '';
    end;
//  Ribbon.Activate := True;
    if Ribbon <> nil then
    begin
      Ribbon.Activate := True;
      Ribbon.OnRibbonBeforeExecute := BeforeExec;
    end;
    SetControlVIsible('RIB__BMSA', GetParamSocSecur('SO_PGMSA', False));
    SetControlVIsible('RIB__BSAISEIARRET', GetParamSocSecur('SO_PGSAISIEARRET', False));
    SetControlVIsible('RIB__BELTNATDOS', GetParamSocSecur('SO_PGGESTELTDYNDOS', False));
    SetControlVIsible('RIB__BBTP', GetParamSocSecur('SO_PGBTP', False));
    SetControlVIsible('RIB__BINTERMITTENTS', GetParamSocSecur('SO_PGINTERMITTENTS', False));
    SetControlVIsible('RIB__BHIERARCHIE', GetParamSocSecur('SO_PGRESPONSABLES', False));
    SetControlVIsible('RIB__BTICKET', GetParamSocSecur('SO_PGTICKETRESTAU', False));
    if not ExisteSQL('SELECT PNA_SALARIE FROM SALARIEPOPUL') then SetControlVIsible('RIB__HPOPULATION', False);
  end
  else
  begin
//    Btn := TtoolBarButton97(GetControl('BACCES'));
    MenuPop := TPopUpMenu(GetControl('POPUPRIBBON'));
    if menuPop <> nil then
    begin
      for i := 0 to MenuPop.Items.Count - 1 do
      begin
        if (MenuPop.Items[i].Name = 'RIB__BMSA') and not (GetParamSocSecur('SO_PGMSA', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BBTP') and not (GetParamSocSecur('SO_PGBTP', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BINTERMITTENTS') and not (GetParamSocSecur('SO_PGINTERMITTENTS', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BHIERARCHIE') and not (GetParamSocSecur('SO_PGRESPONSABLES', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BTICKET') and not (GetParamSocSecur('SO_PGTICKETRESTAU', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BSAISEIARRET') and not (GetParamSocSecur('SO_PGSAISIEARRET', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BELTNATDOS') and not (GetParamSocSecur('SO_PGGESTELTDYNDOS', False)) then MenuPop.Items[i].Visible := False;
        if (MenuPop.Items[i].Name = 'RIB__BELTDOSPLUS') and not (GetParamSocSecur('SO_PGGESTELTDYNDOS', False)) then MenuPop.Items[i].Visible := False;
        MenuPop.Items[i].OnClick := AccesRibbonBar;
      end;
    end;
  end;
{$ENDIF GCGC}
// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
  if Ecran.name = 'SALARIE_DA' then
  begin
    TobZonesCompl := TOB.Create('DEPORTSAL', nil, -1);
    TobZonesComplSauv := TOB.Create('DEPORTSAL', nil, -1);
  end;
  SetControlVisible('BCONTRAT', FALSE); //mcd 04/12/2006 sinon obligation mettre tof associé, et on ne doti pas depuis GI/GA/GC avori acces à ce bouton
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  Avertirtable('TTPAYS'); // PT35
  Avertirtable('YYNATIONALITE');
  // PT87  : 12/05/2004 PH V_50 FQ 10973 Restriction droit accès au salarié
  if StArgument = 'ETBBLQ' then SetControlEnabled('PSA_ETABLISSEMENT', FALSE)
  else
  begin
    argappel := StArgument; //PT150 contient le code processus
    Action := Trim(ReadTokenPipe(StArgument, ';')); //DEBUT PT34
    TAction := Trim(ReadTokenPipe(Action, '='));
    Argument := Trim(ReadTokenPipe(StArgument, ';'));
    Interimaire := StArgument; //FIN PT34
  end;
  // FIN PT87
    // PT75   : 23/09/2003 PH V_421 Suppression bouton compétence FQ 10075
    // SetControlVisible ('BCOMPETENCE' ,FALSE ) ; FQ 10919
  Btn := TToolBarButton97(GetControl('BENFANT'));
  if Btn <> nil then Btn.OnClick := BEnfantsClick;
  setcontrolvisible('LBLNBREMOISANC', FALSE);
  Firstcp := true;
  Majok := false;
  if not VH_Paie.PgSeriaBilan then // PT93
  begin
    SetControlVisible('PSA_CATBILAN', FALSE);
    SetControlVisible('TPSA_CATBILAN', FALSE);
  end;
  RendVisibleOrg; //Rend Visible les lieux de travail et le code stat en fonction des PARAMSOC
  SetControlProperty('PSA_PGMODEREGLE', 'Datatype', 'PGMODEREGLE'); //PT6
// DEB PT136
{$IFNDEF EAGLCLIENT}
//  SetControlProperty('PSA_DATESORTIE', 'EditMask', '');
//  SetControlProperty('PSA_DATESORTIEPREC', 'EditMask', '');
//  SetControlProperty('PSA_DATEENTREEPREC', 'EditMask', '');
{$ENDIF}
// FIN PT136
  // PT52   : 18/10/2002 PH V585 blocage de la longueur à 10c de la saisie de la date de sortie
  SetControlProperty('PSA_DATESORTIE', 'MaxLength', 10);

  Themes := THValCombobox(GetControl('THEME'));
  PCas := THGrid(GetControl('PCAS'));
//DEBUT PT156
  if PCas <> nil then
  begin
    If JaiLeDroitTag(200001) then //PT176
      PCas.OnDblClick := CasParticulierDblClick;
    PCas.ColCount := 3;
    PCas.ColWidths[2] := -1;
  end;
  if (Themes <> nil) and (PCas <> nil) then
  begin
    Themes.Visible := False;
    PCas.RowCount := Themes.Items.Count + 1;
    for i := 0 to Themes.Items.Count - 1 do PCas.Cells[0, i + 1] := Themes.Items[i];
    PCas.ColAligns[0] := TaCenter;
    // DEB P96
//    PPart := THValCombobox(GetControl('PPART'));
//    if PPart <> nil then PPart.Plus := 'ART';
//    PCas.OnRowEnter := PCasRowEnter;
//    PCas.OnClick := PCasOnClick;
    // FIN PT96
//    if (TFFiche(Ecran).FtypeAction <> taConsult) then PCas.Options := PCas.Options + [goEditing];
//FIN PT156
  end;
{$IFNDEF EAGLCLIENT}
  TauxAt := THDBEdit(GetControl('PSA_ORDREAT'));
  ConvCol := THDBEdit(GetControl('PSA_CONVENTION'));
  StandCalend := THDBValCombobox(GetControl('PSA_STANDCALEND'));
  FlisteRib := THMulDbGrid(GetControl('FRIB'));
  PRet := THDBValCombobox(GetControl('PSA_PROFILRET'));
{$ELSE}
  TauxAt := THEdit(GetControl('PSA_ORDREAT'));
  ConvCol := THEdit(GetControl('PSA_CONVENTION'));
  StandCalend := THValCombobox(GetControl('PSA_STANDCALEND'));
  FlisteRib := THGrid(GetControl('FRIB'));
  PRet := THValCombobox(GetControl('PSA_PROFILRET'));
{$ENDIF}
  if (PRet <> nil) and (Pret.VideString = '') then
  begin
    if Pret.VideString = '' then Pret.Items.Add(Traduirememoire('Aucun'));
    Pret.Values.Add('');
  end;
  if TauxAT <> nil then TauxAt.OnElipsisClick := TauxClick;
  if ConvCol <> nil then ConvCol.OnElipsisClick := ConvClick;
  if StandCalend <> nil then StandCalend.OnChange := OnChangeStandCalend;
  if FlisteRib <> nil then FListeRib.OnDblClick := FRibDblClick;
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  Btn := TToolBarButton97(GetControl('TCONGESPAYES'));
  if Btn <> nil then Btn.OnClick := BTCongesPayesClick;

  Btn := TToolBarButton97(GetControl('BIMPRIMER1')); //PT26
  if Btn <> nil then Btn.OnClick := ImpressionSal;

  //BTVentil:=TToolBarButton97(GetControl('BVENTIL'));
  //if BTVentil <> NIL Then BTVentil.OnClick := BTVentilClick;
  Btn := TToolBarButton97(GetControl('BMEMO'));
  if Btn <> nil then
  begin
    Btn.Visible := TRUE;
    Btn.Enabled := TRUE;
  end;
  Btn := TToolBarButton97(GetControl('TCALENDRIER'));
  if Btn <> nil then Btn.OnClick := TcalendrierClick;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

  TypeCalend := '';

  (*{$IFNDEF EAGLCLIENT} PT74-5 Mise en commentaire
  Basanccp := THDBValCombobox(getcontrol('PSA_BASANCCP'));
  {$ELSE}
  Basanccp := THValCombobox(getcontrol('PSA_BASANCCP'));
  {$ENDIF}
  if Basanccp <> nil then
     Basanccp.OnChange := BasanccpChange; *)

{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATENAISSANCE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATENAISSANCE'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATESORTIE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATESORTIE'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;

{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATEENTREEPREC'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATEENTREEPREC'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATESORTIEPREC'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATESORTIEPREC'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  //PT48
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_ANCIENPOSTE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_ANCIENPOSTE'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
  //FIN PT48
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATEANCIENNETE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATEANCIENNETE'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATELIBRE1'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATELIBRE1'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATELIBRE2'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATELIBRE2'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATELIBRE3'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATELIBRE3'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_DATELIBRE4'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_DATELIBRE4'));
{$ENDIF}
  if Edit <> nil then Edit.OnDblClick := DateElipsisclick;
//debut PT217
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_CODEPOSTAL'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_CODEPOSTAL'));
{$ENDIF}
  if Edit <> nil then Edit.OnExit := ExitEdit;
//fin PT217
  Zone := getcontrol('PSA_BASANCCP');
  if Zone <> nil then InitialiseCombo(Zone);
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_SALARIE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_SALARIE'));
{$ENDIF}
  if (VH_PAIE.PgTypeNumSal = 'ALP') then
    if Edit <> nil then Edit.EditMask := '';
  if (VH_PAIE.PgTypeNumSal = 'NUM') then
    if Edit <> nil then Edit.EditMask := '9999999999';
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_QUALIFICATION'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_QUALIFICATION'));
{$ENDIF}
  if Edit <> nil then Edit.Libelle := nil;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_COEFFICIENT'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_COEFFICIENT'));
{$ENDIF}
  if Edit <> nil then Edit.Libelle := nil;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_INDICE'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_INDICE'));
{$ENDIF}
  if Edit <> nil then Edit.Libelle := nil;
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl('PSA_NIVEAU'));
{$ELSE}
  Edit := THEdit(GetControl('PSA_NIVEAU'));
{$ENDIF}
  if Edit <> nil then Edit.Libelle := nil;
  Tob_Etab := TOB.Create('ETABCOMPL', nil, -1);
  Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, FALSE);

  Tob_DonneurOrdre := TOB.Create('DONNEURORDRE', nil, -1);
  Tob_DonneurOrdre.LoadDetailDB('DONNEURORDRE', '', '', nil, FALSE);
  // PT25 07/02/2002 V571 PH suppression historisation salarié pendant la saisie salarié puis supprimé
  // PT9 : 18/09/2001 V547 PH Rajout mise à jour ventilation analytique salarie par défaut
  if VH_Paie.PGAnalytique and VH_Paie.PGSectionAnal then
  begin // Chargement des axes et sections analytiques
    TOB_Axes := TOB.Create('Les axes', nil, -1);
    st := 'SELECT * FROM AXE';
    Q := OpenSql(st, TRUE);
    TOB_Axes.LoadDetailDB('AXE', '', '', Q, FALSE, False);
    Ferme(Q);
    st := 'SELECT S_SECTION,S_AXE FROM SECTION ORDER BY S_SECTION,S_AXE';
    Q := OpenSql(st, TRUE);
    TOB_Section := TOB.Create('Les sections', nil, -1);
    TOB_Section.LoadDetailDB('SECTION', '', '', Q, FALSE, False);
    Ferme(Q);
  end
  else SetControlProperty('Visible', 'PSA_ANAPERSO', FALSE);
  // PT22   Prise de controle sur MultiEmpl pour rendre le bouton accessible ou non
{$IFNDEF EAGLCLIENT}
  MultiEmpl := TDBCheckBox(GetControl('PSA_MULTIEMPLOY'));
{$ELSE}
  MultiEmpl := TCheckBox(GetControl('PSA_MULTIEMPLOY'));
{$ENDIF}
  if MultiEmpl <> nil then MultiEmpl.OnClick := ClickMultiEmpl;
  BtnDate('ZPSA_DATESORTIE');
  BtnDate('ZPSA_DATEENTREEPREC');
  BtnDate('ZPSA_DATESORTIEPREC');

{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILAFP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILAFP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILANCIEN'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILANCIEN'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILAPP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILAPP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILAFP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILAFP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILCGE'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILCGE'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILFNAL'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILFNAL'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILMUT'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILMUT'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PERIODBUL'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PERIODBUL'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILPRE'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILPRE'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFIL'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFIL'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILRBS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILRBS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILREM'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILREM'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILRET'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILRET'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_REDREPAS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_REDREPAS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_REDRTT1'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_REDRTT1'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_REDRTT2'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_REDRTT2'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILTPS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILTPS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILTRANS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILTRANS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('PSA_PROFILTSS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('PSA_PROFILTSS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
  // PT73  : 04/09/2003 PH V_42 Chgt tablettes CODEEMPLOI dans le cas passage 2003
  if VH_Paie.PGPCS2003 then SetControlProperty('PSA_CODEEMPLOI', 'datatype', 'PGCODEPCSESE');
  // PT88 Suppression accès compétences salarié ==> module gestion de compétences
  SetControlVisible('BCOMPETENCE', FALSE);
  // DEB PT95
  if V_PGI.ModePcl <> '1' then
  begin
    if GereIsoflex then
    begin
      Btn := TToolBarButton97(GetControl('BISOFLEx'));
      if Btn <> nil then Btn.onClick := AppelIsoflex;
    end;
  end;
  // FIN PT95
  //DEB PT99 traitement du champ PSA_REGULANCIEN
{$IFNDEF EAGLCLIENT}
  MSpin := THDBSpinEdit(getcontrol('PSA_REGULANCIEN'));
{$ELSE}
  MSpin := THSpinEdit(getcontrol('PSA_REGULANCIEN'));
{$ENDIF}
  if MSpin <> nil then MSpin.OnExit := OnExitRegul;
  // FIN PT99
  Btn := TToolBarButton97(GetControl('BDUPLIQUER'));
  if Btn <> nil then Btn.OnClick := BDupliqueClick;
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  LePanel := THPanel(GetControl('PANELCHOIX'));
  if LePanel <> nil then
  begin
    for Cpt := 0 to Ecran.ComponentCount - 1 do
    begin
      Btn := TToolBarButton97(Ecran.Components[Cpt]);
      if (Btn <> nil) and (Btn.Parent = LePanel) then
        Btn.OnClick := PBTNClick;
      if (Btn <> nil) and (copy(Btn.Name, 1, 3) = 'BZL') then Btn.OnClick := AccesMulsZonesLibres;
    end;
  end;

  CarriereHistoInit(THMultiValComboBox(GetControl('CRITERESHISTO')));
  Btn := TToolBarButton97(GetControl('BHISTO'));
  if Btn <> nil then Btn.OnClick := BTNCLICK;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
//PT113
  PersACharge := THEdit(GetControl('EPERSACHARGE'));
  if PersACharge <> nil then
    PersACharge.OnExit := PersAChargeExit;
//FIN PT113
  SalarieDupliquer := ''; //PT121 Initialisation car utilisé dans OnLoad

  Btn := TToolBarButton97(GetControl('BTNBULDEFAUT')); { PT128 }
  if Assigned(Btn) then Btn.OnClick := ClickBtnBulDefaut; { PT128 }

  Btn := TToolBarButton97(GetControl('BTNCONVPCL')); { PT130}
  if Assigned(Btn) then Btn.OnClick := PCLConvClick; { PT130 }
{PT215
  // PT134
  if VH_Paie.PGAnalytique then
//    SetControlVisible('BTNVENTIL', TRUE)
    SetControlPropertyRibOrNot('Visible', 'BTNVENTIL', TRUE)
  else
//    SetControlVisible('BTNVENTIL', FALSE);
    SetControlPropertyRibOrNot('Visible', 'BTNVENTIL', FALSE);
}
// DEB PT136
  if (GetControl('PSA_DATESORTIE') <> nil) then THEdit(GetControl('PSA_DATESORTIE')).OnKeyDown := FormKeyDown;
  if (GetControl('PSA_DATEENTREEPREC') <> nil) then THEdit(GetControl('PSA_DATEENTREEPREC')).OnKeyDown := FormKeyDown;
  if (GetControl('PSA_DATESORTIEPREC') <> nil) then THEdit(GetControl('PSA_DATESORTIEPREC')).OnKeyDown := FormKeyDown;
// FIN PT136
// DEBUT CCMX-CEGID ORGANIGRAMME DA - 11.12.2006
{$IFNDEF GCGC}
  Ecran.OnKeyDown := FicheKeyDown;
//DEBUT PT43
//PT175
  if (Ecran.name<>'SALARIE_CP') AND (Ecran.name<>'SAL_BANQUECWAS') AND (Ecran.name<>'SALARIE_BANQUE')  then //PT189
     begin
     SetControlText('CTHEMEELTCOMPL', '');
     SetControlText('CZONELIBRE', 'OUI');
     end;
//FIN PT175
  Grille := THGrid(GetControl('GELTCOMPL'));
  if Grille <> nil then
  begin
    Grille.OnDblClick := GrilleEltComplClick;
    Grille.GetCellCanvas := GrilleGetCellCanvas;
  end;
  Grille := THGrid(GetControl('GELTNATDOS'));
  if Grille <> nil then
  begin
    Grille.OnDblClick := GrilleEltNatDos;
  end;
  //DEB PT184
//PT191  Btn := TToolBarbutton97(GetControl('BBANQUE'));
//PT191  if Btn <> nil then Btn.OnClick := BBANQUE_OnClick;
  Btn := TToolBarbutton97(GetControl('BMEMO'));
  if Btn <> nil then Btn.OnClick := BMEMO_OnClick;
//PT191  Btn := TToolBarbutton97(GetControl('BCP'));
//PT191  if Btn <> nil then Btn.OnClick := BCP_OnClick;
  //FIN PT184
  Btn := TToolBarbutton97(GetControl('BRECHELTCOMP'));
  if Btn <> nil then Btn.OnClick := ClickBtEltCompl;
  combo := THValComboBox(GetControl('CTHEMEELTCOMPL'));
  if combo <> nil then combo.OnClick := ChangeComboEltComp;
  combo := THValComboBox(GetControl('CZONELIBRE'));
  if combo <> nil then combo.OnClick := ChangeComboEltComp;
  GereHistoConvention := False; //PT192
  if GetParamSocSecur('SO_PGHISTOAVANCE', False) = True then
  begin
    Q := OpenSQL('SELECT DISTINCT(PPP_PGINFOSMODIF) FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_PGTYPEINFOLS="SAL" AND PPP_HISTORIQUE="X"', True);
    TobParamSal := Tob.Create('parametrage salarie', nil, -1);
    TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
    ferme(Q);
    for i := 0 to TobParamSal.Detail.Count - 1 do
    begin
      LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
      //DEB PT192
      if LeChamp = 'PSA_CONVENTION' then
        GereHistoConvention := True;
      //FIN PT192
      if LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
        //UnControl := TControl(LeChamp);
      UnControl := Ecran.FindComponent(LeChamp);
{$IFNDEF EAGLCLIENT}  // @@@@@@@@@@@@    
      if UnControl is THDBEdit then
      begin
        Edit := THDBEdit(getControl(LeChamp));
        if (Edit <> nil) then
        begin
          if Edit.Visible = TRUE then
          begin
            Edit.Color := CouleurHisto;
            LeControl := CreerBoutonHisto(TTabSheet(Edit.parent), LeChamp, Edit.Left, Edit.Width, Edit.Top);
            if LeControl <> nil then LeControl.OnClick := AccesHistoEdit;
            Edit.Enabled := False;
          end;
        end;
      end;
      if UnControl is THDBValComboBox then
      begin
        MCombo := THDBValComboBox(getControl(LeChamp));
        if MCombo <> nil then
        begin
          if MCombo.visible = true then
          begin
            MCombo.Color := CouleurHisto;
            LeControl := CreerBoutonHisto(TTabSheet(MCombo.parent), LeChamp, MCombo.Left, MCombo.Width, MCombo.Top);
            if LeControl <> nil then LeControl.OnClick := AccesHistoCombo;
            MCombo.Enabled := False;
          end;
        end;
      end;
      if UnControl is THDBcheckBox then
      begin
        Check := THDBcheckBox(getControl(LeChamp));
        if Check <> nil then
        begin
          if Check.visible = true then
          begin
            Check.Color := CouleurHisto;
            LeControl := CreerBoutonHisto(TTabSheet(Check.parent), LeChamp, Check.Left, Check.Width, Check.Top);
            if LeControl <> nil then LeControl.OnClick := AccesHistocheck;
            Check.Enabled := False;
          end;
        end;
      end;
{$ELSE}
      if UnControl is THEdit then
      begin
        Edit := THEdit(getControl(LeChamp));
        Edit.Color := CouleurHisto;
        LeControl := CreerBoutonHisto(TTabSheet(Edit.parent), LeChamp, Edit.Left, Edit.Width, Edit.Top);
        if LeControl <> nil then LeControl.OnClick := AccesHistoEdit;
        Edit.Enabled := False;
      end;
      if UnControl is THValComboBox then
      begin
        MCombo := THValComboBox(getControl(LeChamp));
        MCombo.Color := CouleurHisto;
        LeControl := CreerBoutonHisto(TTabSheet(MCombo.parent), LeChamp, MCombo.Left, MCombo.Width, MCombo.Top);
        if LeControl <> nil then LeControl.OnClick := AccesHistoCombo;
        MCombo.Enabled := False;
      end;
      if UnControl is THcheckBox then
      begin
        Check := THcheckBox(getControl(LeChamp));
        Check.Color := CouleurHisto;
        LeControl := CreerBoutonHisto(TTabSheet(Check.parent), LeChamp, Check.Left, Check.Width, Check.Top);
        if LeControl <> nil then LeControl.OnClick := AccesHistocheck;
        Check.Enabled := False;
      end;
{$ENDIF}
    end;
    TobParamSal.Free;
  end;
//FIN PT143
{$ENDIF}
//DEB PT144
  // DEBUT CCMX-CEGID FQ N° GC14326
  if Assigned(GetControl('HEUREMBAUCHE')) then
  begin
  // FIN CCMX-CEGID FQ N° GC14326
    THHeure := THEdit(GetControl('HEUREMBAUCHE'));
    if THHeure <> nil then THHeure.OnExit := ChangementHeure;
  end; // CCMX-CEGID FQ N° GC14326
//FIN PT144
  PeriodVisitDefaut := GetParamSocSecur('SO_PGPERIODVISITMED', ''); // PT145

//PT148
  SetPlusBanqueCP(GetControl('PSA_RIBVIRSOC'));
  SetPlusBanqueCP(GetControl('PSA_RIBACPSOC'));
  SetPlusBanqueCP(GetControl('PSA_RIBFRAISSOC'));
//FIN PT148
  Btn := TToolBarButton97(GetControl('BCONTRAT'));
  if Btn <> nil then Btn.OnClick := BtnContratClick; //PT170

  //DEB PT184
{$IFNDEF GCGC}
  MenuPop := TPopUpMenu(GetControl('POPUPVENTIL'));
  if MenuPop <> nil then
  begin
    MenuPop.Items[0].OnClick := AccesVentil;
    MenuPop.Items[1].OnClick := AccesVentil;
  end;
  //FIN PT184
{$ENDIF GCGC}
  If VH_PAIE.ModeFiche Then VH_PAIE.FicheOuverte := True; //FLO
end;

//debut PT217
procedure TOM_SALARIES.ExitEdit(Sender: Tobject);
begin
VerifCodePostal(DS, THEdit(GetControl('PSA_CODEPOSTAL')), THEdit(GetControl('PSA_VILLE')), True);
SetField('PSA_VILLE', getcontroltext('PSA_VILLE'));
SetfocusControl ('PSA_VILLE');
end;
//fin PT217

//DEB PT184
{$IFNDEF GCGC}
procedure TOM_SALARIES.AccesVentil(Sender: Tobject);
begin
  if TMenuItem(Sender).Name = 'SALARIE' then
  begin
    if not JaiLeDroitTag(200069) then
      ParamVentil ('SA',GetField('PSA_SALARIE'),'12345',taConsult,FALSE)
    else
      ParamVentil ('SA',GetField('PSA_SALARIE'),'12345',taModif,FALSE);
  end;
  if TMenuItem(Sender).Name = 'SALRUB' then
  begin
    if not JaiLeDroitTag(200068) then
      AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM')+';CONSULTATION')
    else
      AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM'));
  end;
end;
//FIN PT184
{$ENDIF !GCGC}

//DEB PT144
procedure TOM_SALARIES.ChangementHeure(Sender: TObject); //PT- 6
var
  NouvelHeure: string;
begin
  NouvelHeure := FormatDateTime('hh:mm', StrToTime(GetControlText('HEUREMBAUCHE')));
  if NouvelHeure <> HeureEmbauche then ForceUpdate;
end;
//FIN PT144

//*******************PROCEDURE On Load Record***********************//

procedure TOM_SALARIES.OnLoadRecord;
var
  TPS: TOB;
  PCas: THGrid;
  Themes: THValCombobox;
//  PPart: THValCombobox;
  Theme, Profil: string;
  i, j: integer;
  Q: TQuery;
  prof, reglement, Etab: string;
  LblConv: THLabel;
{$IFNDEF GCGC}
  DateVisite: string;
  LePanel: THPanel;
  LeControl: TToolBarButton97;
  Cpt: Integer;
  LblPeriodVisit: THLabel; // PT145
  DatePrevVisitMed: TDateTime; // PT145
  NbrAjout: integer; // PT145
  TypAjout: string; // PT145
  Tob_Ferie: Tob; // PT145
  k: integer; // PT145
  //DEB PT184
  Non200003,Non200004,Non200005,Non200006,Non200007,Non200008,Non200009,Non200010,Non200011 : boolean;
  Non200012,Non200013,Non200014,Non200015,Non200016,Non200017,Non200018,Non200019,Non200020 : boolean;
  Non200021,Non200022,Non200023,Non200024,Non200025,Non200026,Non200027,Non200028,Non200029 : boolean;
  Non200030,Non200031 : boolean;
  //FIN PT184
{$ENDIF}
  Nb: integer; //PT158
begin
  inherited;
{$IFDEF PAIEGRH}
  if not (DS.State in [dsInsert]) then DerniereCreate := '';
{$ENDIF PAIEGRH}
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
  if TobZonesCompl <> nil then //mcd 04/12/2006
  begin
    if GetField('PSA_SALARIE') <> '' then
    begin
      if (not TobZonesCompl.SelectDB('"' + GetField('PSA_SALARIE') + '"', nil)) then
        TobZonesCompl.InitValeurs;
    end
    else TobZonesCompl.InitValeurs;
    if TFfiche(Ecran) <> nil then
      TobZonesCompl.PutEcran(TFfiche(Ecran));
    TobZonesComplSauv.Dupliquer(TobZonesCompl, True, True, False);
  end;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  ChargementTablette('ETB', ''); //PT108
  ChargementTablette('ET', ''); //PT108
// DEB PT103 Gestion de la confidentialité en CWAS
  if GetControl('PCONFIDENTIEL') <> nil then
  begin
    TCheckBox(GetControl('PCONFIDENTIEL', true)).OnClick := nil;
    TCheckBox(GetControl('PCONFIDENTIEL', true)).Checked := (GetField('PSA_CONFIDENTIEL') = '1');
    TCheckBox(GetControl('PCONFIDENTIEL', true)).OnClick := OnClickConfidentiel;
  end;
// FIN PT103
  PGEnabledEtabUser(THValComboBox(GetControl('PSA_ETABLISSEMENT')));

  MajOk := False;
  GblCpCoche := False; { PT102 }
  if (GetField('PSA_DATESORTIE') <> null) and
    (GetField('PSA_DATESORTIE') > Idate1900) then
    GblDtSortie := StrToDate(GetField('PSA_DATESORTIE'))
  else
    GblDtSortie := Idate1900;
  GblMotifSortie := GetField('PSA_MOTIFSORTIE');//PT186
  if (GetField('PSA_DATEENTREE') <> null) then
    GblDtEntree := StrToDate(GetField('PSA_DATEENTREE'))
  else
    GblDtEntree := Idate1900;

  if (TFFiche(Ecran).FTypeAction = taCreat) then
    ;
  Themes := THValCombobox(GetControl('THEME'));
//  PPart := THValCombobox(GetControl('PPART'));
  PCas := THGrid(GetControl('PCAS'));
  if SalarieDupliquer = '' then //PT121
  begin
    if (PCas <> nil) and (Themes <> nil) then
    begin
      for j := 1 to PCas.RowCount do
      begin
        PCas.CellValues[1, j] := '';
        PCas.CellValues[2, j] := '';
      end;
      for j := 0 to Themes.Items.Count - 1 do
        if Theme = Themes.Values[j] then
        begin
          PCas.CellValues[1, j + 1] := '';
          PCas.CellValues[2, j + 1] := '';
          break;
        end;
      TPS := TOB.Create('profils speciaux', nil, -1);
      TPS.LoadDetailDB('PROFILSPECIAUX', '"-";"' + GetField('PSA_SALARIE') + '"', '',
        nil, FALSE, FALSE);
      for i := 0 to TPS.Detail.Count - 1 do
      begin
        Theme := TPS.Detail[i].GetValue('PPS_THEMEPROFIL');
        Profil := TPS.Detail[i].GetValue('PPS_PROFIL'); //PT156
  //       PPart.Text; PT96-1 rechargement des profils
        for j := 0 to Themes.Items.Count - 1 do
          if Theme = Themes.Values[j] then
          begin
            PCas.CellValues[2, j + 1] := Profil;
            PCas.CellValues[1, j + 1] := RechDom('PGPROFILPARTICULIER', Profil, FALSE); // PT156
            break;
          end;
      end;
      TPS.Free;
    end;
  end;

  TypeCalend := getfield('PSA_STANDCALEND');
  GblCalend := getfield('PSA_CALENDRIER');

  if TFFiche(Ecran).name = 'SALARIE' then
  begin
    InitEtab := GetField('PSA_ETABLISSEMENT');
{$IFDEF CCS3}
    SetControlVisible('BCOMPETENCE', FALSE);
{$ENDIF}
  end;

  Etab := GetField('PSA_ETABLISSEMENT');
//On init une tob Réelle sur la table EtablCompl pour la gestion des Champs IdemEtablissement
  if ((Tob_Etab <> nil) and (Etab <> '')) then
    T := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etab], TRUE);

//On init une tob Réelle sur la table Donneur Ordre pour la gestion des Champs Idem Profil
  if TFFiche(Ecran).Name = 'SAL_BANQUECWAS' then
  begin
    Prof := GetField('PSA_PROFIL');
    Reglement := GetField('PSA_PGMODEREGLE');
    TDO := Tob_DonneurOrdre.FindFirst(['PDO_ETABLISSEMENT', 'PDO_PROFIL', 'PDO_PGMODEREGLE'],
      [Etab, prof, reglement], TRUE);
{$IFDEF CCS3}
    SetControlVisible('TBSHTFP', FALSE);
{$ENDIF}
  end;

// Le bouton de routage vers les Congés payés est visible seulement si l'option a été cochée dans
// l'établissement correspondant.
  if Ds.State in [DsEdit, DsBrowse] then SetControlEnabled('BCP', VH_paie.PGCongesPayes); { PT133 }
  EtbCp := False;
  if VH_paie.PGCongesPayes then
  begin
    Q := OpenSQL('SELECT ETB_CONGESPAYES,ETB_DATECLOTURECPN' +
      ' FROM ETABCOMPL WHERE' +
      ' ETB_ETABLISSEMENT="' + GetField('PSA_ETABLISSEMENT') + '"', TRUE);
    if not Q.Eof then
    begin
      EtbCP := (Q.FindField('ETB_CONGESPAYES').Asstring = 'X');
      Dtclot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
    end;
    if Ds.State in [DsEdit, DsBrowse] then SetControlEnabled('BCP', EtbCP); { PT133 }
    Ferme(Q);
  end;

// Gestion de la photo du salarié contenue dans la table LIENSOLE
  if (VH_Paie.PGPhotoSal <> '') and (VH_Paie.PGPhotoSal <> 'NON') then
  begin
    AffichePhoto;
  end;
//Affichage contrat
  AfficheContrat; //PT170

  LblConv := THLabel(GetControl('LBLPSA_CONVENTION'));
  if LblConv <> nil then
  begin
    SetControlProperty('LBLPSA_CONVENTION', 'Caption', RechDom('PGCONVENTION', GetField('PSA_CONVENTION'), FALSE))
  end;

  MajChampIdemEtab;
  if (GetField('PSA_MULTIEMPLOY') = 'X') and JaiLeDroitTag(200052) then   //PT184
  begin
    SetControlEnabled('BMULTIEMPLOY', True);
    SetControlEnabled('PSA_SALAIREMULTI', True);
    SetControlVisible('RIB__BMULTIEMP', True);
  end
  else
  begin
    SetControlEnabled('BMULTIEMPLOY', False);
    SetControlEnabled('PSA_SALAIREMULTI', False);
    SetControlVisible('RIB__BMULTIEMP', False);
  end;

  SetControlProperty('PSA_EDITORG', 'Plus', GetField('PSA_ETABLISSEMENT'));
  ChargeGrilleRib(THGrid(GetControl('FRIB')), '');
  ChargeGrilleRib(THGrid(GetControl('FRIBACPTE')), 'ACP');
  ChargeGrilleRib(THGrid(GetControl('FRIBFP')), 'FPR');
  ControleDate('PSA_DATESORTIE', 'BTNDATESORTIE');
  ControleDate('PSA_DATEENTREEPREC', 'BTNENTREEPREC');
  ControleDate('PSA_DATESORTIEPREC', 'BTNSORTIEPREC');
  TFFiche(Ecran).caption := 'Salarié : ' + GetField('PSA_SALARIE') + ' ' +
    GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
  UpdateCaption(TFFiche(Ecran));

  SalarieDupliquer := ''; // PT104
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  LePanel := THPanel(GetControl('PANELCHOIX'));
  if LePanel <> nil then
  begin
    for Cpt := 0 to Ecran.ComponentCount - 1 do
    begin
      LeControl := TToolBarButton97(Ecran.Components[Cpt]);
      if (LeControl <> nil) and (LeControl.Parent = LePanel) then
        BTNClick(LeControl);
    end;
  end;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

  if Assigned(GetControl('PSA_PERSACHARGE')) then
  begin
  //PT113
    if (GetControlText('PSA_PERSACHARGE') = '99') then
      SetControlText('EPERSACHARGE', '')
    else
      SetControlText('EPERSACHARGE', GetControlText('PSA_PERSACHARGE'));
//FIN PT113
  end;
{$IFNDEF GCGC}
  GerebGed; // PT142
{$ENDIF}
  //DEB PT144
  // DEBUT CCMX-CEGID FQ N° GC14326
  if Assigned(GetControl('HEUREMBAUCHE')) then
  begin
  // FIN CCMX-CEGID FQ N° GC14326
    if TFFiche(Ecran).FTypeAction <> taCreat then
    begin
      HeureEmbauche := FormatDateTime('hh:mm', GetField('PSA_HEUREMBAUCHE'));
      SetControlText('HEUREMBAUCHE', HeureEmbauche);
    end;
    //FIN PT144
  end; // CCMX-CEGID FQ N° GC14326
{$IFNDEF GCGC}
  //DEB PT145
  LblPeriodVisit := THLabel(GetControl('TPSA_DATEVISITMED'));
  Q := OpenSQL('SELECT PVM_DATEVISITMED,PVM_DATEVISITE' +
    ' FROM VISITEMEDTRAV WHERE' +
    ' PVM_SALARIE="' + GetField('PSA_SALARIE') + '" order by PVM_DATEVISITE DESC', TRUE);
  if not Q.Eof then
  begin
    DateVisite := Q.findfield('PVM_DATEVISITMED').AsString;
    if (DateVisite <> '01/01/1900') and (DateVisite <> '') then
      LblPeriodVisit.Caption := 'Prévision visite médicale le ' + DateVisite
    else
    begin
      // Calculer la date prévisionnelle de la prochaine visite en fonction de la périodicité
      if (PeriodVisitDefaut = '3') then
      begin
        NbrAjout := 1;
        TypAjout := 'A';
      end;
      if (PeriodVisitDefaut = '4') then
      begin
        NbrAjout := 2;
        TypAjout := 'A';
      end;
      if (PeriodVisitDefaut = '5') then
      begin
        NbrAjout := 6;
        TypAjout := 'M';
      end;
      if (PeriodVisitDefaut = '6') then
      begin
        NbrAjout := 3;
        TypAjout := 'M';
      end;
      DateVisite := Q.findfield('PVM_DATEVISITE').AsString;
      DatePrevVisitMed := StrToDateTime(DateVisite);
      DatePrevVisitMed := PlusDate(DatePrevVisitMed, NbrAjout, TypAjout);
      // Vérifier qu'il s'agit d'un jour ouvré.
      if not Assigned(Tob_Ferie) then Tob_Ferie := ChargeTobFerie(DatePrevVisitMed, PlusDate(DatePrevVisitMed, 10, 'J'));
      k := 10;
      while k > 0 do
      begin
        if (IfJourFerie(Tob_Ferie, DatePrevVisitMed)) or (DayOfWeek(DatePrevVisitMed) in [1, 7]) then
        begin
          DatePrevVisitMed := PlusDate(DatePrevVisitMed, 1, 'J');
          Continue;
        end
        else
          Break;
      end;
      Tob_Ferie.Free;
      LblPeriodVisit.Caption := 'Prévision visite médicale le ' + DateVisite;
    end;
  end
  else
    LblPeriodVisit.Caption := 'Aucune prévision de visite médicale';
  Ferme(Q);
  //FIN PT145
{$ENDIF}
  if Argappel = 'S04' then SetActiveTabSheet('PEMPLOI'); //PT150

//  SavFieldsPopul; PT151
{$IFNDEF GCGC}
  if GetActiveTabSheet('Pages').name = 'PELTCOMPL' then AfficheEltCompl
  else if GetActiveTabSheet('Pages').name = 'PELTNATDOS' then
  begin
    AfficheEltNat;
    If JaiLeDroitTag(200001) and JaiLeDroitTag(200060) then    //PT176  //PT184
      SetControlVisible('RIB__BELTDOSPLUS', True);
  end
  else if GetActiveTabSheet('Pages').name = 'PFORMINT' then RempliOnglet('FORMATIONSINT')
  else if GetActiveTabSheet('Pages').name = 'PFORMEXT' then RempliOnglet('FORMATIONSEXT')
  else if GetActiveTabSheet('Pages').name = 'PDIPLOMES' then RempliOnglet('DIPLOMES')
  else if GetActiveTabSheet('Pages').name = 'PELTNATDOS' then RempliOnglet('BTNCOMPETENCE');

  //DEB PT176
  If Not JaiLeDroitTag(200001) then
  begin
    SetControlEnabled('RIB__HRibbonControl1',False);
    SetControlEnabled('RIB__HRibbonControl4',False);
    SetControlEnabled('RIB__BValider',False);

    SetControlEnabled('RIB__HRibbonControl54',False);
    SetControlEnabled('RIB__HRibbonControl5',False);
    SetControlEnabled('RIB__HRibbonControl3',False);

    SetControlEnabled('RIB__HRibbonControl56',False);
    SetControlEnabled('RIB__HRibbonControl6',False);
    SetControlEnabled('RIB__Binsert',False);

    SetControlEnabled('RIB__HRibbonControl61',False);
    SetControlEnabled('RIB__BDUPLIQUER',False);

    //DEB PT176
    SetControlEnabled('BENFANT',False);
    //FIN PT176

    //DEB PT184
    SetControlEnabled('BMULTIEMPLOY',False);
    SetControlEnabled('BTNCONVPCL',False);
    //FIN PT184
  end;
  //FIN PT176

  //DEB PT177
  if V_PGI.CWAS then
  begin
    SetControlVisible('RIB__HRibbonControl9',False);
    SetControlVisible('RIB__HRibbonControl10',False);
    SetControlVisible('RIB__HRibbonControl11',False);
    SetControlVisible('RIB__HRibbonControl12',False);
  end;
  //FIN PT177
{$ENDIF}

//DEB PT158
//PT175
  if (Ecran.name<>'SALARIE_CP') AND (Ecran.name<>'SAL_BANQUECWAS') AND (Ecran.name<>'SALARIE_BANQUE') then  //PT189
     begin
     Q:= OpenSQL ('SELECT COUNT (PEF_ACHARGE)'+
                  ' FROM ENFANTSALARIE WHERE'+
                  ' (PEF_SALARIE="'+GetField ('PSA_SALARIE')+'") AND'+
                  ' (PEF_ACHARGE="X") AND (PEF_TYPEPARENTAL="001")', TRUE);
     if not Q.EOF then
        Nb:= Q.Fields[0].Asinteger
     else
        Nb:= 0;
     if (Nb>1) then
        SetControlText ('LBLNBENFACHARGE', 'Dont '+IntToStr (Nb)+' enfants')
     else
        SetControlText ('LBLNBENFACHARGE', 'Dont '+IntToStr (Nb)+' enfant');
     Ferme (Q);
     end;
//FIN PT175
//FIN PT158
  EstSalarieMSA := False;
  if GetParamSocSecur('SO_PGMSA', False) then
  begin
    if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PSA_SALARIE') + '" AND PSE_MSA="X"') then EstSalarieMSA := True;
  end;
  EstSalarieSpectacle := False;
  if GetParamSocSecur('SO_PGINTERMITTENTS', False) then
  begin
    if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PSA_SALARIE') + '" AND PSE_INTERMITTENT="X"') then EstSalarieSpectacle := True;
  end;

{$IFNDEF GCGC}
  //DEB PT184
  //Onglet Salarié / état civil
  If Not JaiLeDroitTag(200002) then
  begin
    SetControlVisible('RIB__BETATCIVIL',False);
    SetControlVisible('PETATCIVIL',False);
  end
  else
    If Not JaiLeDroitTag(200051) then
    begin
      SetControlEnabled('PETATCIVIL',False);
      SetControlEnabled('PSA_NATIONALITE',False);
      SetControlEnabled('PSA_DATENAISSANCE',False);
      SetControlEnabled('PSA_DEPTNAISSANCE',False);
      SetControlEnabled('PSA_COMMUNENAISS',False);
      SetControlEnabled('PSA_PAYSNAISSANCE',False);
      SetControlEnabled('PSA_NUMEROSS',False);
      SetControlEnabled('PSA_AUTRENUM',False);
      SetControlEnabled('PSA_SITUATIONFAMIL',False);
      SetControlEnabled('EPERSACHARGE',False);
      SetControlEnabled('PSA_PERSACHARGE',False);
      SetControlEnabled('BENFANT',False);
      SetControlEnabled('PSA_CARTESEJOUR',False);
      SetControlEnabled('PSA_DELIVPAR',False);
      SetControlEnabled('PSA_DATEXPIRSEJOUR',False);
    end;

  //Onglet Salarié / emploi
  Non200003 := False;
  If Not JaiLeDroitTag(200003) then
  begin
    SetControlVisible('RIB__BEMPLOI',False);
    SetControlVisible('PEMPLOI',False);
    Non200003 := True;
  end
  else
    If Not JaiLeDroitTag(200052) then
    begin
      SetControlEnabled('PEMPLOI',False);
      SetControlEnabled('PSA_DATEENTREE',False);
      SetControlEnabled('PSA_MOTIFENTREE',False);
      SetControlEnabled('HEUREMBAUCHE',False);
      SetControlEnabled('PSA_MOTIFSUSPPAIE',False);
      SetControlEnabled('ZPSA_DATEENTREEPREC',False);
      SetControlEnabled('PSA_DATEENTREEPREC',False);
      SetControlEnabled('PSA_CATDADS',False);
      SetControlEnabled('PSA_MULTIEMPLOY',False);
      SetControlEnabled('BMULTIEMPLOY',False);
      SetControlEnabled('PSA_DATESORTIE',False);
      SetControlEnabled('ZPSA_DATESORTIE',False);
      SetControlEnabled('PSA_MOTIFSORTIE',False);
      SetControlEnabled('PSA_SORTIEDEFINIT',False);
      SetControlEnabled('PSA_SUSPENSIONPAIE',False);
      SetControlEnabled('ZPSA_DATESORTIEPREC',False);
      SetControlEnabled('PSA_DATESORTIEPREC',False);
      SetControlEnabled('PSA_CATBILAN',False);
      SetControlEnabled('PSA_SALAIREMULTI',False);
      SetControlEnabled('PSA_DATEANCIENNETE',False);
      SetControlEnabled('PSA_ANCIENNETE',False);
      SetControlEnabled('PSA_TYPPROFILANC',False);
      SetControlEnabled('PSA_PROFILANCIEN',False);
      SetControlEnabled('PSA_ANCIENPOSTE',False);
      SetControlEnabled('PSA_REGULANCIEN',False);
    end;

  //Onglet Salarié / affectation
  Non200004 := False;
  If Not JaiLeDroitTag(200004) then
  begin
    SetControlVisible('RIB__BAFFECTATION',False);
    SetControlVisible('PAFFECTATION',False);
    Non200004 := True;
  end
  else
    If Not JaiLeDroitTag(200053) then
    begin
      SetControlEnabled('PAFFECTATION',False);
      SetControlEnabled('PSA_TYPCONVENTION',False);
      SetControlEnabled('PSA_CONVENTION',False);
      SetControlEnabled('BTNCONVPCL',False);
      SetControlEnabled('PSA_CODEEMPLOI',False);
      SetControlEnabled('PSA_LIBELLEEMPLOI',False);
      SetControlEnabled('PSA_QUALIFICATION',False);
      SetControlEnabled('PSA_COEFFICIENT',False);
      SetControlEnabled('PSA_INDICE',False);
      SetControlEnabled('PSA_NIVEAU',False);
      SetControlEnabled('PSA_PRISEFFECTIF',False);
      SetControlEnabled('PSA_UNITEPRISEFF',False);
      SetControlEnabled('PSA_CODESTAT',False);
      SetControlEnabled('PSA_TRAVAILN1',False);
      SetControlEnabled('PSA_TRAVAILN2',False);
      SetControlEnabled('PSA_TRAVAILN3',False);
      SetControlEnabled('PSA_TRAVAILN4',False);
      SetControlEnabled('PSA_TYPEDITORG',False);
      SetControlEnabled('PSA_EDITORG',False);
      SetControlEnabled('PSA_ETATBULLETIN',False);
      SetControlEnabled('BTNBULDEFAUT',False);
      SetControlEnabled('PSA_TYPPERIODCT',False);
      SetControlEnabled('PSA_PERIODCT',False);
    end;

  if Non200003 and Non200004 then
    SetControlVisible('RIB__HRibbonGroup10',False);

  //Onglet Salarié / profils
  Non200005 := False;
  If Not JaiLeDroitTag(200005) then
  begin
    SetControlVisible('RIB__BPROFILS',False);
    SetControlVisible('PPROFILS',False);
    Non200005 := True;
  end
  else
    If Not JaiLeDroitTag(200054) then
    begin
      SetControlEnabled('PPROFILS',False);
      SetControlEnabled('PSA_TYPACTIVITE',False);
      SetControlEnabled('PSA_ACTIVITE',False);
      SetControlEnabled('PSA_TYPPROFILREM',False);
      SetControlEnabled('PSA_PROFILREM',False);
      SetControlEnabled('PSA_TYPPROFIL',False);
      SetControlEnabled('PSA_PROFIL',False);
      SetControlEnabled('PSA_TYPPERIODEBUL',False);
      SetControlEnabled('PSA_PERIODBUL',False);
      SetControlEnabled('PSA_PROFILTPS',False);
      SetControlEnabled('PSA_SALAIRETHEO',False);
      SetControlEnabled('PCAS',False);
    end;

  //Onglet Salarié / autres profils
  Non200006 := False;
  If Not JaiLeDroitTag(200006) then
  begin
    SetControlVisible('RIB__BAUTREPROFILS',False);
    SetControlVisible('PPROFILSUITE',False);
    Non200006 := True;
  end
  else
    If Not JaiLeDroitTag(200055) then
    begin
      SetControlEnabled('PPROFILSUITE',False);
      SetControlEnabled('PSA_TYPPROFILRBS',False);
      SetControlEnabled('PSA_PROFILRBS',False);
      SetControlEnabled('PSA_TYPREDRTT2',False);
      SetControlEnabled('PSA_REDRTT2',False);
      SetControlEnabled('PSA_TYPREDRTT1',False);
      SetControlEnabled('PSA_REDRTT1',False);
      SetControlEnabled('PSA_TYPREDREPAS',False);
      SetControlEnabled('PSA_REDREPAS',False);
      SetControlEnabled('PSA_TYPPROFILAFP',False);
      SetControlEnabled('PSA_PROFILAFP',False);
      SetControlEnabled('PSA_PCTFRAISPROF',False);
      SetControlEnabled('PSA_TYPPROFILAPP',False);
      SetControlEnabled('PSA_PROFILAPP',False);
      SetControlEnabled('PSA_TYPPROFILRET',False);
      SetControlEnabled('PSA_PROFILRET',False);
      SetControlEnabled('PSA_TYPPROFILMUT',False);
      SetControlEnabled('PSA_PROFILMUT',False);
      SetControlEnabled('PSA_TYPPROFILPRE',False);
      SetControlEnabled('PSA_PROFILPRE',False);
      SetControlEnabled('PSA_TYPPROFILTSS',False);
      SetControlEnabled('PSA_PROFILTSS',False);
      SetControlEnabled('PSA_TYPPROFILFNAL',False);
      SetControlEnabled('PSA_PROFILFNAL',False);
      SetControlEnabled('PSA_TYPPROFILTRANS',False);
      SetControlEnabled('PSA_PROFILTRANS',False);
    end;

  if Non200005 and Non200006 then
    SetControlVisible('RIB__HRibbonGroup11',False);

  //Onglet Salarié / contrat
  Non200007 := False;
  If Not JaiLeDroitTag(200007) then
  begin
    SetControlVisible('RIB__BCONTRATSALAIRES',False);
    SetControlVisible('PCONTRAT',False);
    Non200007 := True;
  end
  else
    If Not JaiLeDroitTag(200056) then
    begin
      SetControlEnabled('PCONTRAT',False);
      SetControlEnabled('PSA_HORAIREMOIS',False);
      SetControlEnabled('PSA_HORHEBDO',False);
      SetControlEnabled('PSA_SALAIREMOIS1',False);
      SetControlEnabled('PSA_SALAIREMOIS2',False);
      SetControlEnabled('PSA_SALAIREMOIS3',False);
      SetControlEnabled('PSA_SALAIREMOIS4',False);
      SetControlEnabled('PSA_SALAIREMOIS5',False);
      SetControlEnabled('PSA_HORANNUEL',False);
      SetControlEnabled('PSA_TAUXHORAIRE',False);
      SetControlEnabled('PSA_SALAIRANN1',False);
      SetControlEnabled('PSA_SALAIRANN2',False);
      SetControlEnabled('PSA_SALAIRANN3',False);
      SetControlEnabled('PSA_SALAIRANN4',False);
      SetControlEnabled('PSA_SALAIRANN5',False);
      SetControlEnabled('PSA_STANDCALEND',False);
      SetControlEnabled('PSA_CALENDRIER',False);
      SetControlEnabled('PSA_TYPJOURHEURE',False);
      SetControlEnabled('PSA_JOURHEURE',False);
      SetControlEnabled('PSA_FONCTIONSAL',False);
      SetControlEnabled('PSA_PROFILCDD',False);
    end;

  //Onglet Salarié / DADS
  Non200008 := False;
  If Not JaiLeDroitTag(200008) then
  begin
    SetControlVisible('RIB__BDADS',False);
    SetControlVisible('PDADSCRC',False);
    Non200008 := True;
  end
  else
    If Not JaiLeDroitTag(200057) then
    begin
      SetControlEnabled('PDADSCRC',False);
      SetControlEnabled('PSA_REMPOURBOIRE',False);
      SetControlEnabled('PSA_DADSCAT',False);
      SetControlEnabled('PSA_DADSPROF',False);
      SetControlEnabled('PSA_UNITETRAVAIL',False);
      SetControlEnabled('PSA_CONDEMPLOI',False);
      SetControlEnabled('PSA_TAUXPARTIEL',False);
      SetControlEnabled('PSA_TYPDADSFRAC',False);
      SetControlEnabled('PSA_DADSFRACTION',False);
      SetControlEnabled('PSA_TRAVETRANG',False);
      SetControlEnabled('PSA_ORDREAT',False);
      SetControlEnabled('PSA_TYPEREGIME',False);
      SetControlEnabled('PSA_REGIMESS',False);
      SetControlEnabled('PSA_REGIMEMAL',False);
      SetControlEnabled('PSA_REGIMEAT',False);
      SetControlEnabled('PSA_REGIMEVIP',False);
      SetControlEnabled('PSA_REGIMEVIS',False);
      SetControlEnabled('PSA_ALLOCFORFAIT',False);
      SetControlEnabled('PSA_REMBJUSTIF',False);
      SetControlEnabled('PSA_PRISECHARGE',False);
      SetControlEnabled('PSA_AUTREDEPENSE',False);
      SetControlEnabled('PSA_TYPPRUDH',False);
      SetControlEnabled('PSA_PRUDHCOLL',False);
      SetControlEnabled('PSA_PRUDHSECT',False);
      SetControlEnabled('PSA_PRUDHVOTE',False);
    end;

  if Non200007 and Non200008 then
    SetControlVisible('RIB__HRibbonGroup12',False);

  //Onglet Salarié / Zones libres
  Non200009 := False;
  If Not JaiLeDroitTag(200009) then
  begin
    SetControlVisible('RIB__BZONESLI',False);
    SetControlVisible('PZONELIBRE',False);
    Non200009 := True;
  end
  else
    If Not JaiLeDroitTag(200058) then
    begin
      SetControlEnabled('PZONELIBRE',False);
      SetControlEnabled('PSA_DATELIBRE1',False);
      SetControlEnabled('PSA_DATELIBRE2',False);
      SetControlEnabled('PSA_DATELIBRE3',False);
      SetControlEnabled('PSA_DATELIBRE4',False);
      SetControlEnabled('PSA_BOOLLIBRE1',False);
      SetControlEnabled('PSA_BOOLLIBRE2',False);
      SetControlEnabled('PSA_BOOLLIBRE3',False);
      SetControlEnabled('PSA_BOOLLIBRE4',False);
      SetControlEnabled('PSA_LIBREPCMB1',False);
      SetControlEnabled('PSA_LIBREPCMB2',False);
      SetControlEnabled('PSA_LIBREPCMB3',False);
      SetControlEnabled('PSA_LIBREPCMB4',False);
    end;

  //Onglet Salarié / Eléments dynamiques
  Non200010 := False;
  If Not JaiLeDroitTag(200010) then
  begin
    SetControlVisible('RIB__BZONEDYNA',False);
    SetControlVisible('PELTCOMPL',False);
    Non200010 := True;
  end
  else
    If Not JaiLeDroitTag(200059) then
      SetControlEnabled('PELTCOMPL',False);

  //Onglet Salarié / Eléments nationaux
  Non200011 := False;
  If (Not JaiLeDroitTag(200011)) or (not GetParamSocSecur('SO_PGGESTELTDYNDOS', False)) then
  begin
    SetControlVisible('RIB__BELTNATDOS',False);
    SetControlVisible('RIB__BELTDOSPLUS',False);
    SetControlVisible('PELTNATDOS',False);
    Non200011 := True;
  end
  else
    If Not JaiLeDroitTag(200060) then
    begin
      SetControlEnabled('PELTNATDOS',False);
      SetControlVisible('RIB__BELTDOSPLUS',False);
    end;

  if Non200009 and Non200010 and Non200011 then
    SetControlVisible('RIB__HRibbonGroup13',False);

  //Onglet Compléments / Calendrier
  Non200012 := False;
  If Not JaiLeDroitTag(200012) then
  begin
    SetControlVisible('RIB__HRibbonControl39',False);
    SetControlVisible('TCALENDRIER',False);
    Non200012 := True;
  end;

  //Onglet Compléments / Contrats
  Non200013 := False;
  If Not JaiLeDroitTag(200013) then
  begin
    SetControlVisible('RIB__HRibbonControl43',False);
    SetControlVisible('BCONTRAT',False);
    Non200013 := True;
  end;
  //Onglet Compléments / Banque
  Non200014 := False;
  If Not JaiLeDroitTag(200014) then
  begin
    SetControlVisible('RIB__HRibbonControl62',False);
    SetControlVisible('BBANQUE',False);
    Non200014 := True;
  end;
  //Onglet Compléments / Mémos
  Non200015 := False;
  If Not JaiLeDroitTag(200015) then
  begin
    SetControlVisible('RIB__HRibbonControl63',False);
    SetControlVisible('BMEMO',False);
    Non200015 := True;
  end;
  //Onglet Compléments / Documents GED
  Non200016 := False;
  If NonGed or (Not JaiLeDroitTag(200016)) then
  begin
    SetControlVisible('RIB__BGED',False);
    Non200016 := True;
  end;
  if Non200012 and Non200013 and Non200014 and Non200015 and Non200016 then
    SetControlVisible('RIB__HRibbonGroup7',False);

  //Onglet Compléments / Paramètres CP
  Non200017 := False;
  If Not JaiLeDroitTag(200017) then
  begin
    SetControlVisible('RIB__HRibbonControl40',False);
    Non200017 := True;
  end;
  //Onglet Compléments / Détail CP
  Non200018 := False;
  If Not JaiLeDroitTag(200018) then
  begin
    SetControlVisible('RIB__BDETAILCP',False);
    SetControlVisible('BCP',False);
    Non200018 := True;
  end;
  if Non200017 and Non200018 then
    SetControlVisible('RIB__HRibbonGroup14',False);

  //Onglet Compléments / Salarié/Rubrique (pré ventilation)
  Non200019 := False;
  If Not JaiLeDroitTag(200019) then
  begin
    SetControlVisible('RIB__BVENTILRUB',False);
    SetControlVisible('SALRUB',False);
    Non200019 := True;
  end;
  //Onglet Compléments / Salarié (ventilation)
  Non200020 := False;
  If Not JaiLeDroitTag(200020) then
  begin
    SetControlVisible('RIB__BVENTILSALARIE',False);
    SetControlVisible('SALARIE',False);
    Non200020 := True;
  end;
  if Non200019 and Non200020 then
  begin
    SetControlVisible('RIB__HRibbonGroup1',False);
    SetControlVisible('BTNVENTIL',False);
  end;

  //Onglet Compléments / MSA
  Non200021 := False;
  If (Not JaiLeDroitTag(200021)) or (not GetParamSocSecur('SO_PGMSA', False)) then
  begin
    SetControlVisible('RIB__BMSA',False);
    Non200021 := True;
  end;
  //Onglet Compléments / BTP
  Non200022 := False;
  If (Not JaiLeDroitTag(200022)) or (not GetParamSocSecur('SO_PGBTP', False)) then
  begin
    SetControlVisible('RIB__BBTP',False);
    Non200022 := True;
  end;
  //Onglet Compléments / Intermittents
  Non200023 := False;
  If (Not JaiLeDroitTag(200023)) or (not GetParamSocSecur('SO_PGINTERMITTENTS', False)) then
  begin
    SetControlVisible('RIB__BINTERMITTENTS',False);
    Non200023 := True;
  end;
  //Onglet Compléments / Ticket restaurant
  Non200024 := False;
  If (Not JaiLeDroitTag(200024)) or (not GetParamSocSecur('SO_PGTICKETRESTAU', False)) then
  begin
    SetControlVisible('RIB__BTICKET',False);
    Non200024 := True;
  end;
  //Onglet Compléments / Hiérarchie
  Non200025 := False;
  If (Not JaiLeDroitTag(200025)) or (not GetParamSocSecur('SO_PGRESPONSABLES', False)) then
  begin
    SetControlVisible('RIB__BHIERARCHIE',False);
    Non200025 := True;
  end;
  //Onglet Compléments / Saisie arrêt
  Non200026 := False;
  If (Not JaiLeDroitTag(200026)) or (not GetParamSocSecur('SO_PGSAISIEARRET', False)) then
  begin
    SetControlVisible('RIB__BSAISEIARRET',False);
    Non200026 := True;
  end;
  //Onglet Compléments / Visites médicales
  Non200027 := False;
  If Not JaiLeDroitTag(200027) then
  begin
    SetControlVisible('RIB__BMEDECINE',False);
    Non200027 := True;
  end;
  //Onglet Compléments / Travailleurs handicapés
  Non200028 := False;
  If Not JaiLeDroitTag(200028) then
  begin
    SetControlVisible('RIB__BHANDICAPE',False);
    Non200028 := True;
  end;
  if Non200021 and Non200022 and Non200023 and Non200024 and Non200025 and Non200026 and Non200027 and Non200028 then
    SetControlVisible('RIB__HRibbonGroup15',False);
  if Non200012 and Non200013 and Non200014 and Non200015 and Non200016 and Non200017 and Non200018 and
     Non200019 and Non200020 and Non200021 and Non200022 and Non200023 and Non200024 and Non200025 and
     Non200026 and Non200027 and Non200028 then
    SetControlVisible('RIB__HCOMPLEMENTS',False);

  //Onglet Documents
  Non200029 := False;
  If Not JaiLeDroitTag(200029) then
  begin
    SetControlVisible('RIB__BDUE',False);
    SetControlVisible('RIB__BMALADIE',False);
    SetControlVisible('RIB__BASSEDIC',False);
    SetControlVisible('RIB__BACCTRAV',False);
    SetControlVisible('RIB__BDECLACCTRAV',False);
    SetControlVisible('RIB__HRibbonGroup28',False);
    SetControlVisible('RIB__HRibbonTabSheet5',False);
    Non200029 := True;
  end;

  //Onglet Historique
  Non200030 := False;
  If Not JaiLeDroitTag(200030) then
  begin
    SetControlVisible('RIB__BBULLETIN',False);
    SetControlVisible('RIB__BFICHIND',False);
    SetControlVisible('RIB__BABSENCE',False);
    SetControlVisible('RIB__HRibbonGroup25',False);
    SetControlVisible('RIB__BHISTOSAL',False);
    SetControlVisible('RIB__HRibbonGroup26',False);
    SetControlVisible('RIB__BIJSS',False);
    SetControlVisible('RIB__HRibbonGroup27',False);
    SetControlVisible('RIB__BPOPULATION',False);
    SetControlVisible('RIB__HRibbonGroup2',False);
    SetControlVisible('RIB__HHISTORIQUE',False);
    Non200030 := True;
  end;

  //Onglet Formation
  Non200031 := False;
  If Not JaiLeDroitTag(200031) then
  begin
    SetControlVisible('RIB__HFORMCOMP',False);
    SetControlVisible('RIB__BFORMATIONSINT',False);
    SetControlVisible('RIB__BFORMATIONSEXT',False);
    SetControlVisible('RIB__BDIPLOMES',False);
    SetControlVisible('RIB__BDIF',False);
    SetControlVisible('RIB__BOMPCETENCES',False);
    Non200031 := True;
  end;

  if Non200010 and Non200011 and Non200021 and Non200022 and Non200023 and Non200024 and Non200025 and
     Non200026 and Non200027 and Non200028 and Non200029 and Non200030 and Non200031  then
    SetControlVisible('BACCES',False);
  //FIN PT184
{$ENDIF}

end;


//*******************PROCEDURE On Update Record***********************//
procedure TOM_SALARIES.OnUpdateRecord;
var
HORANNUEL, HORHEBDO, HORAIREMOIS: Double;
XX: Variant;
PCas: THGrid;
Themes: THValCombobox;
//PPart: THValCombobox;
{$IFNDEF EAGLCLIENT}
ComboOrg: THDbValComboBox;
StDate: THDBEdit;
{$ELSE}
ComboOrg: THValComboBox;
StDate: THEdit;
{$ENDIF}
Profil: string;
EtabSal, NomSal, PrenSal, civilite, Adr1Sal, VilleSal, CpSal, mes, DateEntree: string;
DeptNaiss: string; // PT152
mes1, mes2, mes3, mes4, mes5, mes6, mes7, mes8, Mes9, numero, Champs: string;
longsal, i, num : integer;
//LongLib, LongPre: integer;
Onglet: TPageControl;
SauvError : integer;
{$IFDEF GCGC}
  Q: Tquery; // CCMX-CEGID ORGANIGRAMME DA 16.06.2006
{$ENDIF}
  Qt: TQuery; //PT149
//UpdateIdemPop : TUpdateIdemPop;
{$IFNDEF GCGC}
  FieldList: TStringList; //, TmpStringList
  indexfield: Integer;
  stQFieldPopul, champ, suffixe, prefixe: string;
  QFieldPopul: TQuery;
  mr: TModalResult;
{$ENDIF !GCGC}
  // d PT157
  DateNaissanceSS: TDateTime;
  aa, mm, jj: word;
// f PT157
LeTypCp,LaVar,LeLib,Deno: string;
PctFraisProf: string;   // PT216
flPctFraisProf: Double; // PT216
BufBool : boolean;
const
Voyelle = ['a', 'e', 'i', 'u', 'y', 'o', 'A', 'E', 'I', 'U', 'Y', 'O'];

begin
inherited;
{$IFDEF PAIEGRH}
  //DEB PT185
  LeStatut := DS.State;
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PSA_SALARIE');
  //FIN PT185
{$ENDIF PAIEGRH}

Themes:= THValCombobox (GetControl ('THEME'));
//PPart:= THValCombobox (GetControl ('PPART'));
PCas:= THGrid (GetControl ('PCAS'));
if (Ecran.name = 'SALARIE_CP') then
   begin
   LeTypCP:= GetField ('PSA_TYPNBACQUISCP');
   if (LeTypCp = 'PER') then
      begin
      LaVar:= GetField ('PSA_NBACQUISCP');
      LeLib:= RechDom ('PGVARIABLE', LaVar, FALSE);
      if (LeLib = '') or (LeLib = 'Error') then
         begin
         LastError:= 1;
         PgiBox ('Vous devez renseigner le champ jour acquis bulletin avec une'+
                 ' variable existante', TFFiche (Ecran).Caption);
         SetFocusControl ('PSA_NBACQUISCP');
         end;
      end;
   end;
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
  if Ecran.name = 'SALARIE_DA' then
  begin
    Controle_SalarieDA;
    if LastError = 0 then
    begin
      if (ds <> nil) then
      begin
        TobZonesCompl.GetEcran(TFfiche(Ecran), nil);
        TobZonesCompl.PutValue('PSE_SALARIE', GetField('PSA_SALARIE'));
        TobZonesCompl.Putvalue('PSE_ORGANIGRAMME', 'X');
        if TobZonesCompl.GetValue('PSE_CODESERVICE') <> '' then
        begin
          Q := OpenSQL('SELECT PGS_SECRETAIREABS,PGS_SECRETAIRENDF,PGS_SECRETAIREVAR,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_RESPSERVICE' +
            ' FROM SERVICES WHERE PGS_CODESERVICE="' + TobZonesCompl.GetValue('PSE_CODESERVICE') + '"', True);
          if not Q.eof then
          begin
            TobZonesCompl.PutValue('PSE_ASSISTABS', Q.FindField('PGS_SECRETAIREABS').AsString);
            TobZonesCompl.PutValue('PSE_ASSISTVAR', Q.FindField('PGS_SECRETAIREVAR').AsString);
            TobZonesCompl.PutValue('PSE_ASSISTNDF', Q.FindField('PGS_SECRETAIRENDF').AsString);
            TobZonesCompl.PutValue('PSE_RESPONSABS', Q.FindField('PGS_RESPONSABS').AsString);
            TobZonesCompl.PutValue('PSE_RESPSERVICE', Q.FindField('PGS_RESPSERVICE').AsString);
            TobZonesCompl.PutValue('PSE_RESPONSNDF', Q.FindField('PGS_RESPONSNDF').AsString);
            TobZonesCompl.PutValue('PSE_RESPONSVAR', Q.FindField('PGS_RESPONSVAR').AsString);
            TobZonesCompl.PutValue('PSE_RESPONSFOR', Q.FindField('PGS_RESPONSFOR').AsString);
          end;
        end
        else
        begin
          TobZonesCompl.PutValue('PSE_ASSISTABS', '');
          TobZonesCompl.PutValue('PSE_ASSISTVAR', '');
          TobZonesCompl.PutValue('PSE_ASSISTNDF', '');
          TobZonesCompl.PutValue('PSE_RESPONSABS', '');
          TobZonesCompl.PutValue('PSE_RESPSERVICE', '');
          TobZonesCompl.PutValue('PSE_RESPONSNDF', '');
          TobZonesCompl.PutValue('PSE_RESPONSVAR', '');
          TobZonesCompl.PutValue('PSE_RESPONSFOR', '');
        end;
        TobZonesCompl.InsertOrUpdateDB(FALSE);
        TobZonesCompl.SetAllModifie(False);
      end;
    end;
    Exit;
  end;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

  if (PCas = nil) or (Themes = nil) then
    exit;
ExecuteSQL ('DELETE FROM PROFILSPECIAUX WHERE'+
            ' PPS_ETABSALARIE="-" AND'+
            ' PPS_CODE="'+GetField ('PSA_SALARIE')+'"');
for i := 1 to PCas.RowCount - 1 do
    begin
    Profil := PCas.CellValues[2, i];
    if (Profil <> '') then
       ExecuteSQL ('INSERT INTO PROFILSPECIAUX'+
                   ' (PPS_ETABSALARIE, PPS_CODE, PPS_PROFIL, PPS_THEMEPROFIL)'+
                   ' VALUES'+
                   ' ("-", "'+GetField ('PSA_SALARIE')+'", "'+Profil+'",'+
                   ' "'+Themes.Values[i - 1]+'")');
    end;


//Contrôle sur la Page Identité
Onglet:= TPageControl (GetControl ('Pages'));
EtabSal:= GetField ('PSA_ETABLISSEMENT');
NomSal:= GetField ('PSA_LIBELLE');
PrenSal:= GetField ('PSA_PRENOM');
civilite:= GetField ('PSA_CIVILITE');
Adr1Sal:= GetField ('PSA_ADRESSE1');
VilleSal:= GetField ('PSA_VILLE');
CpSal:= GetField ('PSA_CODEPOSTAL');
codesexe:= GetField ('PSA_SEXE');
numss:= GetField ('PSA_NUMEROSS');
DeptNaiss := GetField('PSA_DEPTNAISSANCE'); // PT152
{$IFNDEF EAGLCLIENT}
StDate:= THDBEdit (GetControl ('PSA_DATEENTREE'));
{$ELSE}
StDate:= THEdit (GetControl ('PSA_DATEENTREE'));
{$ENDIF}
if (StDate <> nil) then
   DateEntree:= StDate.Text;
resultat:= 0;
Mes:= '';
Mes1:= '';
Mes2:= '';
Mes3:= '';
Mes4:= '';
Mes5:= '';
Mes6:= '';
Mes7:= '';
Mes8:= '';
Mes9:= '';
longsal:= length (GetField ('PSA_SALARIE'));
if ((GetField ('PSA_SALARIE') <> '') and (LongSal < 11)) then
   begin
   if (EtabSal = '') then
      begin
      Mes1 := '#13#10    - établissement ';
      Mes:= Mes1;
      end;
   if (NomSal = '') then
      begin
      Mes2 := '#13#10    - nom ';
      Mes:= Mes+Mes2;
      end;
   if (PrenSal = '') then
      begin
      Mes3 := '#13#10    - prénom ';
      Mes:= Mes+Mes3;
      end;
   if (civilite = '') then
      begin
      Mes4 := '#13#10    - civilité ';
      Mes:= Mes+Mes4;
      end;
   if (codesexe = '') then
      begin
      Mes5 := '#13#10    - sexe ';
      Mes:= Mes+Mes5;
      end;
   if (Adr1Sal = '') then
      begin
      Mes6 := '#13#10    - adresse ';
      Mes:= Mes+Mes6;
      end;
   if (CpSal = '') then
      begin
      Mes7 := '#13#10    - code postal ';
      Mes:= Mes+Mes7;
      end;
   if (VilleSal = '') then
      begin
      Mes8 := '#13#10    - ville';
      Mes:= Mes+Mes8;
      end;
   if (DateEntree = '') then
      begin
      Mes9 := '#13#10    - date d''entrée';
      Mes:= Mes+Mes9;
      end;
{$IFNDEF EAGLCLIENT}
   ComboOrg:= THDbValComboBox (GetControl ('PSA_CODESTAT'));
{$ELSE}
   ComboOrg:= THValComboBox (GetControl ('PSA_CODESTAT'));
{$ENDIF}
   if (ComboOrg <> nil) then
      if (ComboOrg.visible = True) then
         if ((ComboOrg.value='') and (VH_Paie.PGLibCodeStat<>'')) then
            begin
            {if (VH_Paie.PGLibCodeStat[1] in Voyelle) then
               Deno:= ' l'' '
            else
               Deno:= ' le ';}//PT161
          Deno := ' ';
            Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibCodeStat;
            Champs:= 'PSA_CODESTAT';
            end;

   for num := 1 to VH_Paie.PGNbreStatOrg do
       begin
       Numero:= InttoStr (num);
       if (Num > VH_Paie.PGNbreStatOrg) then
          exit;
{$IFNDEF EAGLCLIENT}
       ComboOrg:= THDbValComboBox (GetControl ('PSA_TRAVAILN'+Numero));
{$ELSE}
       ComboOrg:= THValComboBox (GetControl ('PSA_TRAVAILN'+Numero));
{$ENDIF}
       if (ComboOrg <> nil) then
          if (ComboOrg.visible = True) then
             if (ComboOrg.value = '') then
                begin
                if ((Num=1) and (VH_Paie.PGLibelleOrgStat1<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat1[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}//PT161
              Deno := ' ';
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibelleOrgStat1;
                   Champs:= 'PSA_TRAVAILN1';
                   end;
                if ((Num=2) and (VH_Paie.PGLibelleOrgStat2<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibelleOrgStat2;
                   Champs:= 'PSA_TRAVAILN2';
                   end;
                if ((Num=3) and (VH_Paie.PGLibelleOrgStat3<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibelleOrgStat3;
                   Champs:= 'PSA_TRAVAILN3';
                   end;
                if ((Num=4) and (VH_Paie.PGLibelleOrgStat4<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibelleOrgStat4;
                   Champs:= 'PSA_TRAVAILN4';
                   end;
                end;
       end;

   for num := 1 to VH_Paie.PGNbCombo do
       begin
       Numero:= InttoStr (num);
       if (Num > VH_Paie.PGNbCombo) then
          exit;
{$IFNDEF EAGLCLIENT}
       ComboOrg:= THDbValComboBox (GetControl ('PSA_LIBREPCMB'+Numero));
{$ELSE}
       ComboOrg:= THValComboBox (GetControl ('PSA_LIBREPCMB'+Numero));
{$ENDIF}
       if (ComboOrg <> nil) then
          if (ComboOrg.visible = True) then
             if (ComboOrg.value = '') then
                begin
                if ((Num=1) and (VH_Paie.PGLibCombo1<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibCombo1;
                   Champs:= 'PSA_LIBREPCMB1';
                   end;
                if ((Num=2) and (VH_Paie.PGLibCombo2<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibCombo2;
                   Champs:= 'PSA_LIBREPCMB2';
                   end;
                if ((Num=3) and (VH_Paie.PGLibCombo3<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibCombo3;
                   Champs:= 'PSA_TLIBREPCMB3';
                   end;
                if ((Num=4) and (VH_Paie.PGLibCombo4<>'')) then
                   begin
                   {if (VH_Paie.PGLibelleOrgStat2[1] in Voyelle) then
                      Deno:= ' l'' '
                   else
                      Deno:= ' le ';}
              Deno := ' '; //PT161
                   Mes:= Mes+'#13#10    -'+Deno+VH_Paie.PGLibCombo4;
                   Champs:= 'PSA_LIBREPCMB4';
                   end;
                end;
       end;

   if (Champs <> '') then
      if (Onglet <> nil) then
         SetfocusControl (Champs);

   if (DateEntree = '') then
      begin
      Onglet.ActivePage:= Onglet.Pages[2];
      SetfocusControl ('PSA_DATEENTREE');
      end;
   if (VilleSal = '') then
      SetfocusControl ('PSA_VILLE');
   if (CpSal = '') then
      SetfocusControl ('PSA_CODEPOSTAL');
   if (Adr1Sal = '') then
      SetfocusControl ('PSA_ADR1');
   if (codesexe = '') then
      SetfocusControl ('PSA_SEXE');
   if (civilite = '') then
      SetfocusControl ('PSA_CIVILITE');
   if (PrenSal = '') then
      SetfocusControl ('PSA_PRENOM');
   if (NomSal = '') then
      SetfocusControl ('PSA_LIBELLE');
   if (EtabSal = '') then
      SetfocusControl ('PSA_ETABLISSEMENT');
   end;

//PT141
if ((GetField ('PSA_CONDEMPLOI')='P') and
   (GetField ('PSA_TAUXPARTIEL')=0)) then
   begin
    Mes1 := '#13#10    - taux temps partiel';
   Mes:= Mes1;
   end;
//FIN PT141   

  if mes <> '' then
  begin
    LastError := 1;
    PgiBox('Saisie obligatoire. Veuillez renseigner :' + mes + ' du salarié.',
      Ecran.caption);
  end;
// d PT152
if (DeptNaiss = '') then
   HShowMessage('1;Salarié:;Vous n''avez pas renseigné le département de naissance.' +
      '#13#10#13#10Pensez-y ulterieurement;E;O;O;O;;;', '', '');
// f PT152
  SauvError := LastError; // PT131
  codesexe := GetField('PSA_SEXE');
  numss := GetField('PSA_NUMEROSS');
  if (codesexe <> '') then
  begin
    if (numss <> '') then
    begin
//Les test sur l'année, le mois et le départements de naissance
//sont effectués que si le numéro SS fait 15 caractères, et avant le test de la clé
      if (Length(numss) = 15) or (Length(numss) = 13) then
      begin
{PT155-2
         resultAnnee:= TestNumeroSSNaissance (numss, 'Annee');
         resultMois:= TestNumeroSSNaissance (numss, 'Mois');
         resultDepart:= TestNumeroSSNaissance (numss, 'Depart');
}
        TestNumeroSSNaissance(numss, 'Annee', GetField('PSA_DATENAISSANCE'),
          GetField('PSA_DEPTNAISSANCE'), resultAnnee,
          resultMois, resultDepart);
        TestNumeroSSNaissance(numss, 'Mois', GetField('PSA_DATENAISSANCE'),
          GetField('PSA_DEPTNAISSANCE'), resultAnnee,
          resultMois, resultDepart);
        TestNumeroSSNaissance(numss, 'Depart', GetField('PSA_DATENAISSANCE'),
          GetField('PSA_DEPTNAISSANCE'), resultAnnee,
          resultMois, resultDepart);
//FIN PT155-2
// d PT157
        if (GetField('PSA_DATENAISSANCE') <> NULL) then
        begin
          DateNaissanceSS := GetField('PSA_DATENAISSANCE');
          if DateNaissanceSS <> idate1900 then
          begin
            DecodeDate(DateNaissanceSS, aa, mm, jj);
            if mm < 10 then
              MoisSS := '0' + IntToStr(mm)
            else
              MoisSS := IntToStr(mm);

            AnneeSS := IntToStr(aa);
          end;
        end;
// PT157

        MesErrorSecuNaissance(resultAnnee, ResultMois, ResultDepart);
        if (ResultAnnee <> 0) or (ResultMois <> 0) or (Resultdepart <> 0) then
          numss := GetField('PSA_NUMEROSS');
      end;
      numss := GetField('PSA_NUMEROSS');
    end;
    resultat := TestNumeroSS(numss, codesexe);
    MesErrorSecu(resultat);
    if (resultat = -6) or (Resultat = -5) or (Resultat = -4) then
      LastError := 1;
    if (SauvError <> 0) and (LastError = 0) then
      LastError := SauvError; // PT131
  end;

   //DEB PT149
  Qt := OpenSql('SELECT ETB_ETABLISSEMENT FROM ETABCOMPL WHERE ETB_ETABLISSEMENT = "' + EtabSal + '"', TRUE);
  if Qt.Eof then
  begin
    LastError := 1;
    PgiBox('Vous ne pouvez pas sélectionner un établissement qui ne comporte pas d''informations complémentaires sur le social .',
      Ecran.caption);
    exit;
  end;
  Ferme(Qt);
   //FIN PT149

{$IFNDEF GCGC}
//debut PT154
//Rend obligatoire les champs utilisés dans le parametrage des populations
  FieldList := GetFieldsListFromPopulParam();
  for indexfield := 0 to FieldList.count - 1 do
  begin
    if (GetControlText(FieldList.Strings[indexfield]) = '')
      and (assigned(getcontrol(FieldList.Strings[indexfield]))) then
    begin
      //On récupère le nom du champ
      champ := '';
      prefixe := LeftStr(FieldList.Strings[indexfield], 3);
      suffixe := RightStr(FieldList.Strings[indexfield], length(FieldList.Strings[indexfield]) - 4);
      stQFieldPopul := 'select PAI_LIBELLE from paieparim where pai_prefix = "'
        + prefixe + '" and pai_suffix = "' + suffixe + '"';
      QFieldPopul := OpenSQL(stQFieldPopul, True, 1);
      if not QFieldPopul.Eof then
        champ := QFieldPopul.Fields[0].asstring;
      Ferme(QFieldPopul);
      if champ = '' then
        champ := suffixe;
      mr := PGIAskCancel('Le champ ' + champ + ' est utilisé dans le paramétrage des populations.#10#13Si vous ne renseignez pas ce champ, le salarié risque d''être exclu des populations.#10#13#10#13Voulez-vous renseigner ce champ ?');
      if mr = mrYes then
      begin
        SetFocusControl(FieldList.Strings[indexfield]);
        LastError := 1;
        FreeAndNil(FieldList);
        exit;
      end else if mr = mrNo then
      begin
      end else if mr = mrCancel then
      begin
        LastError := 1;
        FreeAndNil(FieldList);
        Ecran.close;
        exit;
      end;
    end;
  end;
  FreeAndNil(FieldList);
//fin PT154
{$ENDIF}


//Controle contenu des champs Horaires avnt controles de cohérence
  XX := GetField('PSA_HORANNUEL');
  if XX = NULL then
    HORANNUEL := 0
  else
    HORANNUEL := GetField('PSA_HORANNUEL');
  XX := GetField('PSA_HORAIREMOIS');
  if XX = NULL then
    HORAIREMOIS := 0
  else
    HORAIREMOIS := GetField('PSA_HORAIREMOIS');
  XX := GetField('PSA_HORHEBDO');
  if XX = NULL then
    HORHEBDO := 0
  else
    HORHEBDO := GetField('PSA_HORHEBDO');

  ControleHoraires(HORHEBDO, HORAIREMOIS, HORANNUEL);

{//PT151 : Contrôle de l'influance des champs modifiés sur les populations
  if not ValidUpdateFields then
  begin
    LastError:= 1;
    PgiBox (TraduireMemoire('Les modifications n''ont pas été enregistrées.'), Ecran.caption);
    exit;
  end;}

//Rechargement des tablettes
{PT162-6
if (LastError=0) and (Getfield ('PSA_SALARIE')<>'') and
   (Getfield ('PSA_LIBELLE')<>'') then
   begin
}

  //DEB PT164
  if (GetField('PSA_TYPEREGIME') = '-') then
  begin
    if (GetField('PSA_REGIMESS') = '147') or (GetField('PSA_REGIMESS') = '148') or (GetField('PSA_REGIMESS') = '149') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime SS : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMESS'), False) + ' n''est pas autorisée.'), Ecran.caption);
      exit;
    end;
  end
  else
  begin
    if (GetField('PSA_REGIMEMAL') = '147') or (GetField('PSA_REGIMEMAL') = '148') or (GetField('PSA_REGIMEMAL') = '150') or (GetField('PSA_REGIMEMAL') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire risque maladie : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEMAL'), False) + ' n''est pas autorisée.'), Ecran.caption);
      exit;
    end;
    if (GetField('PSA_REGIMEVIP') = '150') or (GetField('PSA_REGIMEVIP') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire vieillesse (PP) : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEVIP'), False) + ' n''est pas autorisée.'), Ecran.caption);
      exit;
    end;
    if (GetField('PSA_REGIMEVIS') = '150') or (GetField('PSA_REGIMEVIS') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire vieillesse (PS) : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEVIS'), False) + ' n''est pas autorisée.'), Ecran.caption);
      exit;
    end;
  end;
  //FIN PT164

  CreationTier := False;  //PT182
  if (DS.State in [dsInsert]) then
  begin // Uniquement en creation ==> creation du tiers associé
  // Chargement des infos nescessaires à la creation du compte de tiers associé
      {if (VH_Paie.PgTypeNumSal = 'NUM') and (VH_Paie.PgTiersAuxiAuto = TRUE) then}//PT219
      if (VH_Paie.PgTiersAuxiAuto = TRUE) then//PT219
        //DEB PT182
        CreationTier := True;
{    Libell := GetField('PSA_LIBELLE');
    LongLib := Length(Libell);
    Preno := GetField('PSA_PRENOM');
    LongPre := Length(Preno);
    Libell := Copy(Libell, 1, LongLib) + ' ' + Copy(Preno, 1, LongPre);
    with InfTiers do
    begin
      Libelle := Copy(Libell, 1, 35);
      Adresse1 := GetField('PSA_ADRESSE1');
      Adresse2 := GetField('PSA_ADRESSE2');
      Adresse3 := GetField('PSA_ADRESSE3');
      Ville := GetField('PSA_VILLE');
      Telephone := GetField('PSA_TELEPHONE');
      CodePostal := GetField('PSA_CODEPOSTAL');
      Pays := '';
    end;
// Creation du compte de tiers en automatique et recup du numéro
    CodeAuxi := CreationTiers(InfTiers, LeRapport, 'SAL',
      GetField('PSA_SALARIE'));
    if (CodeAuxi <> '') then
      SetField('PSA_AUXILIAIRE', CodeAuxi);

    CreateYRS (GetField ('PSA_SALARIE'), '', ''); //PT155-1  }
    //FIN PT182
  end;
{PT162-6
   end;
}

//Suppression historisation salarié pendant la saisie salarié puis supprimé
  if (GetField('PSA_DATESORTIE') = NULL) or (GetField('PSA_DATESORTIE') < 0) then
    SetField('PSA_DATESORTIE', IDate1900);
  if IsFieldModified('PSA_CONFIDENTIEL') then
    MajConfidentialite;

  if (LastError = 0) and (DS.State in [dsInsert]) then
    CreatEnCours := TRUE
  else
    CreatEnCours := FALSE;
  // DEBUT CCMX-CEGID FQ N° GC14326
  if Assigned(GetControl('HEUREMBAUCHE')) then
  begin
  // FIN CCMX-CEGID FQ N° GC14326
// d PT153
    if (GetControlText('HEUREMBAUCHE') <> '') and (GetControlText('HEUREMBAUCHE') <> '  :  ') then
      SetField('PSA_HEUREMBAUCHE', StrToTime(GetControlText('HEUREMBAUCHE'))) //PT144
    else
      SetField('PSA_HEUREMBAUCHE', StrToTime('00:00'));
// f PT153
  end; // CCMX-CEGID FQ N° GC14326
{$IFNDEF GCGC}

  if TFFiche(Ecran).FTypeAction = TaCreat then
  begin
    if (SauvError = 0) and (lastError = 0) then
      if GetParamSocSecur('SO_PGHISTOAVANCE', False) = True then RendAccesChampHisto(False, true);
  end;
{$ENDIF GCGC}

//PT193
{PT218
if ((DS.State in [dsInsert]) and
   (ExisteSQL ('SELECT PSI_INTERIMAIRE'+
               ' FROM INTERIMAIRES WHERE'+
               ' PSI_INTERIMAIRE="'+GetField ('PSA_SALARIE')+'"'))) then
   begin
   i:= PGIAsk ('Ce matricule salarié existe déjà. S''agit-il d''une mutation ?',
               Ecran.caption);
   if (i=mrNo) then
      begin
      LastError:= 1;
      PgiBox (TraduireMemoire ('Veuillez modifier le matricule salarié'),
              Ecran.caption);
      exit;
      end;
   end;
}
if (DS.State in [dsInsert]) then
   begin
   i:= 0;
   BufBool:= ExisteSQL ('SELECT PSI_INTERIMAIRE'+
                        ' FROM INTERIMAIRES WHERE'+
                        ' PSI_INTERIMAIRE="'+GetField ('PSA_SALARIE')+'"');
   if (JaiLeDroitTag (200101)) then
      begin
      if (BufBool) then
         i:= PGIAsk ('Ce matricule salarié existe déjà. S''agit-il d''une'+
                     ' mutation ?', Ecran.caption);
      end
   else
      if (BufBool) then
         i:= mrNo;
   if (i=mrNo) then
      begin
      LastError:= 1;
      PgiBox (TraduireMemoire ('Veuillez modifier le matricule salarié'),
              Ecran.caption);
      SetFocusControl ('PSA_SALARIE');
      exit;
      end;
   end;
//FIN PT218
//FIN 193

//DEB PT216
if GetField('PSA_PROFILAFP') <> '' then
begin
  PctFraisProf := GetField('PSA_PCTFRAISPROF');
  try
    flPctFraisProf := Valeur(PctFraisProf);
    if flPctFraisProf <= 0.0 then
    begin
      LastError:= 1;
      PgiBox(TraduireMemoire('Le pourcentage d''abattement des frais professionels est nul ou négatif alors qu''un profil d''abattement est définit'), Ecran.caption);
      Exit;
    end;
  except
    LastError:= 1;
    PgiBox(TraduireMemoire('Le pourcentage d''abattement des frais professionels n''est pas pas un nombre alors qu''un profil d''abattement est définit'), Ecran.caption);
    Exit;
  end;
end;
//FIN PT216

  if (SauvError <> 0) and (LastError = 0) then
    LastError := SauvError; // PT131 -> PT160 Laisser en fin du OnUpate pour contrôle champs obligatoires
end;

//*******************PROCEDURE PCasRowEnter***********************//
(*
procedure TOM_SALARIES.PCasRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: boolean);
var
  PPart: THValCombobox;
  Theme: THValCombobox;
begin
  PPart := THValCombobox(GetControl('PPART'));
  Theme := THValCombobox(GetControl('THEME'));
  if (PPArt <> nil) and (Theme <> nil) then PPart.Plus := Theme.Values[THGrid(Sender).Row - 1];
end;
*)


//*******************PROCEDURE On Change Field**************************//

procedure TOM_Salaries.OnChangeField(F: TField);
var
  Pa: TPageControl;
  Q, QNbCharge, QRech, QRibSoc: TQuery;
{$IFNDEF EAGLCLIENT}
  CbBasanccp, civilite, ComboEtab, ComboMaj, ComboMotif: THDBValCombobox;
  ComboSAl, MCombo, sexe: THDBValCombobox;
  CarteSejour, ChampEdit, ChampLib, DateExpirSejour, DelivPar: THDBEdit;
  EditMaj, MEdit, Valanc, Edit: THDBEdit;
  SpinEditMaj, NbPerACharge: THDBSpinEdit;
{$ELSE}
  CbBasanccp, civilite, ComboEtab, ComboMaj, ComboMotif: THValCombobox;
  ComboSAl, MCombo, sexe: THValCombobox;
  CarteSejour, ChampEdit, ChampLib, DateExpirSejour, DelivPar: THEdit;
  EditMaj, MEdit, Valanc, Edit: THEdit;
  SpinEditMaj, NbPerACharge: TSpinEdit;
{$ENDIF}
  LblConv, LblRibSoc, LDateclotcp, LibDebContrat, LibFinContrat: THLabel;
  Libelle, MLabel: THLabel;
  Btn: TToolBarButton97;
  i, IntSal, j, k, long, Nb, NbA, NbM, rep, zz: integer;
  Basanccp, CInterdit, CodeSal, DD, DF, DeptNaiss, jour, mois, salarie : string;
  St, StPlus, TempSal, TestPays, typdatanc, ValCoEtabChamp, Valeur : string;
  vide, Vrai : string;
  Init, aa, jj, mm, PremAnnee, PremMois: word;
  NouvelEntree, okok: Boolean;
  DDebContrat, DFinContrat, DEntree, DSortie, DNaissance, DDPaie, DTemp: TDateTime;
  DFPaie, DDPaieEntree, DFPaieSortie, Ladate: TDateTime;
  TEtabl: TOB;
  NbAnneeMois: Double;
  cee: Boolean; //PT141B
  MotifFin:String;//PT186
begin
inherited;
vide:= '';
if ds.state in [dsedit, dsinsert] then
   if not majok then
      MajChampIdemEtab;

if (F.FieldName = 'PSA_CONGESPAYES') then
   begin
   if GetField ('PSA_CONGESPAYES') = 'X' then
      begin
      GblCpCoche:= True;
      St:= RechDom ('PGETABCONGES', GetField ('PSA_ETABLISSEMENT'), False);
      if (St = '') or (St = 'Error') then
         begin
        // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
         PgiBox ('Gestion des congés payés impossible. Vous ne gérez pas les'+
                 ' congés payés au niveau établissement.', Ecran.caption);
{$ENDIF}
        // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
         if GetField ('PSA_CONGESPAYES') <> '-' then
            SetField ('PSA_CONGESPAYES', '-');
         end;
      end
   else
      begin
      if GblCpCoche then
         begin
        if ExisteSql('SELECT ##TOP 1## 1 FROM ABSENCESALARIE WHERE' + //PT180
                       ' PCN_SALARIE="'+GetField('PSA_SALARIE')+'" AND'+
                       ' PCN_TYPEMVT="CPA"') then
            begin
            if PgiAsk ('Souhaitez-vous supprimer les mouvements de congés'+
                       ' payés du salarié ?',
                       'Congés payés du salarié : '+GetField('PSA_SALARIE'))=MrYes then
               ExecuteSQL ('DELETE FROM ABSENCESALARIE WHERE'+
                           ' PCN_SALARIE="'+GetField('PSA_SALARIE')+'" AND'+
                           ' PCN_TYPEMVT="CPA"');
            end;
         GblCpCoche:= False;
         end;
      end;
   SetControlEnabled ('PSA_TYPEDITBULCP', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_EDITBULCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                       (Getfield ('PSA_TYPEDITBULCP')='PER'));
   SetControlEnabled ('PSA_CPACQUISMOIS', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_NBREACQUISCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                          (Getfield ('PSA_CPACQUISMOIS')='PER'));
   SetControlEnabled ('PSA_SSDECOMPTE', (Getfield ('PSA_CONGESPAYES')='X') and
                                        (Getfield ('PSA_CPACQUISMOIS')='PER'));
   SetControlEnabled ('PSA_TYPNBACQUISCP', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_NBACQUISCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                        (Getfield ('PSA_TYPNBACQUISCP')='PER'));
   SetControlEnabled ('PSA_CPACQUISSUPP', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_NBRECPSUPP', (Getfield ('PSA_CONGESPAYES')='X') and
                                        (Getfield ('PSA_CPACQUISSUPP')='PER'));
   SetControlEnabled ('PSA_CPACQUISANC', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_BASANCCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                      (Getfield ('PSA_CPACQUISANC')='PER'));
   SetControlEnabled ('PSA_VALANCCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                      (Getfield ('PSA_CPACQUISANC')='PER'));
   SetControlEnabled ('PSA_DATANC', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_TYPDATANC', (Getfield ('PSA_CONGESPAYES')='X') and
                                       (Getfield ('PSA_DATANC')='PER'));
   SetControlEnabled ('PSA_DATEACQCPANC', (Getfield ('PSA_CONGESPAYES')='X') and
                                          (Getfield ('PSA_DATANC')='PER') and
                                          (Getfield ('PSA_TYPDATANC')='2'));
   SetControlEnabled ('PSA_CPTYPEMETHOD', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_VALORINDEMCP', (Getfield ('PSA_CONGESPAYES')='X') and
                                          (Getfield ('PSA_CPTYPEMETHOD')='PER'));
   SetControlEnabled ('PSA_CPTYPEVALO', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_MVALOMS', (Getfield ('PSA_CONGESPAYES')='X') and
                                     (Getfield ('PSA_CPTYPEVALO') = 'PER'));
   SetControlEnabled ('PSA_VALODXMN', (Getfield ('PSA_CONGESPAYES')='X') and
                                      (Getfield ('PSA_CPTYPEVALO')='PER') and
                                      (Getfield ('PSA_MVALOMS')='M'));
   SetControlEnabled ('PSA_CPTYPERELIQ', (Getfield ('PSA_CONGESPAYES')='X'));
   SetControlEnabled ('PSA_RELIQUAT', (Getfield ('PSA_CONGESPAYES')='X') and
                                      (Getfield ('PSA_CPTYPERELIQ')='PER'));
   SetControlEnabled ('PSA_TYPPAIEVALOMS', (Getfield ('PSA_CONGESPAYES')='X')); { PT132 }
   SetControlEnabled ('PSA_PAIEVALOMS', (Getfield ('PSA_CONGESPAYES')='X') and
                                        (Getfield('PSA_TYPPAIEVALOMS')='PER')); { PT126 }
   Exit;
   end;

//ONGLET IDENTITE
if (F.FieldName = 'PSA_SALARIE') then
   if (GetField ('PSA_SALARIE')) <> '' then
      begin
      Salarie:= Trim (GetField ('PSA_SALARIE'));
      if (isnumeric (Salarie)) and (Salarie <> '') then
         if ((Length (salarie) = 10) and (salarie > '2147483647')) then
            begin
            PgiBox ('Vous ne pouvez pas saisir un matricule superieur à'+
                    ' 2147483647!', 'Matricule du salarié');
            SetField ('PSA_SALARIE', '2147483647');
            exit;
            end;
      if (VH_PAIE.PgTypeNumSal = 'NUM') and (isnumeric(Salarie)) and
         (Salarie <> '') then
         begin
         Salarie:= ColleZeroDevant (StrToInt (trim (Salarie)), 10);
         if Salarie <> (GetField ('PSA_SALARIE')) then
            SetField ('PSA_SALARIE', Salarie);
         end
      else
         if (VH_PAIE.PgTypeNumSal = 'ALP') and
            (Salarie <> GetField ('PSA_SALARIE')) then
            SetField ('PSA_SALARIE', Salarie);
    // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
      AlimentecumulConge;
{$ENDIF}
    // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
      LDateclotcp:= THLabel(getcontrol('LDATECLOTCP'));
      if LDateclotcp <> nil then
         begin
         decodedate (DtClot, aa, mm, jj);
         LDateclotcp.caption:= inttostr (jj)+'/'+inttostr (mm)+'/'+inttostr (aa);
         end;
      exit;
   end;

if (F.FieldName = 'PSA_ETABLISSEMENT') then
   begin
   Libelle:= THLabel (GetControl ('TPSA_ETABLISSEMENT'));
   if Libelle <> nil then
      Libelle.Caption:= RechDom ('TTETABLISSEMENT',
                                 GetField ('PSA_ETABLISSEMENT'), FALSE);
//Reaffectation tablette organisme à editer
   SetControlProperty ('PSA_EDITORG', 'Plus', GetField('PSA_ETABLISSEMENT'));
//Afin d'obliger la validation sur le Changement d'Etablissement;
   if DS.State in [dsEdit, dsInsert] then
      if GetControlText('PSA_ETABLISSEMENT') <> InitEtab then
         begin
         ChangeEtab:= True;
         St:= RechDom ('PGETABCONGES', GetField('PSA_ETABLISSEMENT'), False);
         EtbCP:= ((St <> '') and (St <> 'Error'));
{$IFNDEF EAGLCLIENT}
         okok:= (TFFiche(Ecran).Modifier = False);
{$ELSE !EAGLCLIENT}
         okok := false;
{$ENDIF}
         SetControlEnabled ('TCALENDRIER', OkOk);
         if (VH_Paie.PgCongesPayes) and (EtbCP) then
            SetControlEnabled ('BCP', OkOk);
         if DS.State in [dsInsert] then
            begin
            if (VH_Paie.PgCongesPayes) and (EtbCP) then
               SetField ('PSA_CONGESPAYES', 'X')
            else
               SetField ('PSA_CONGESPAYES', '-');
            end;
         SetControlEnabled ('BBANQUE', OkOk);
         MajChampIdemEtab;
         end
      else
         if DS.State in [dsEdit] then
            begin
            okok:= True;
            SetControlEnabled ('TCALENDRIER', OkOk);
            if (VH_Paie.PgCongesPayes) and (EtbCP) then
               SetControlEnabled ('BCP', OkOk);
            SetControlEnabled('BBANQUE', OkOk);
            if ChangeEtab = True then
               begin
               MajChampIdemEtab;
               ChangeEtab:= False;
               end;
            end;

//Imputation par defaut convention collective de l'etablissement en creation
   if DS.State in [dsInsert] then
      begin // Affectation par defaut convention collective etab sur le salarie
      TEtabl:= Tob_Etab.FindFirst (['ETB_ETABLISSEMENT'],
                                   [GetField('PSA_ETABLISSEMENT')], TRUE);
      if TEtabl <> nil then
         SetField ('PSA_CONVENTION', TEtabl.GetValue ('ETB_CONVENTION'));
      end;
   exit;
   end;

if (F.FieldName = 'PSA_SEXE') then
   begin
{$IFNDEF EAGLCLIENT}
   sexe:= THDBValComboBox (GetControl ('PSA_SEXE'));
   civilite:= THDBValComboBox (GetControl ('PSA_CIVILITE'));
{$ELSE}
   sexe:= THValComboBox (GetControl ('PSA_SEXE'));
   civilite:= THValComboBox (GetControl ('PSA_CIVILITE'));
{$ENDIF}
   if (sexe <> nil) and (civilite <> nil) then
      begin
      codesexe:= sexe.value;
      numss:= (GetField ('PSA_NUMEROSS'));
      if (sexe.value = 'M') and (Getfield ('PSA_NOMJF') <> '') then
         begin
         PGIBox ('Incohérence Nom de jeune fille et sexe !',
                 'Incohérence '+Sexe.text+' :');
         Setfield ('PSA_NOMJF', vide);
         end;
      if (civilite.value <> '') and (sexe.value <> '') then
         begin
         if sexe.value = 'M' then
            begin
            if (civilite.value = 'MLE') or (civilite.value = 'MME') then
               begin
               PGIBox ('Veuiller re-saisir la civilité du salarié !',
                       'Incohérence '+Civilite.text+' .. '+Sexe.text+' :');
               SetField ('PSA_CIVILITE', vide);
               end;
            end
         else
         if sexe.value = 'F' then
            if civilite.value = 'MR' then
               begin
               PGIBox ('Veuiller re-saisir la civilité du salarié !',
                       'Incohérence '+Civilite.text+' .. '+Sexe.text+' :');
               SetField('PSA_CIVILITE', vide);
               end;
         end;
      end;
   exit;
   end;

if (F.FieldName = 'PSA_CIVILITE') then
   begin
{$IFNDEF EAGLCLIENT}
   civilite:= THDBValComboBox (GetControl ('PSA_CIVILITE'));
   sexe:= THDBValComboBox (GetControl ('PSA_SEXE'));
{$ELSE}
   civilite:= THValComboBox (GetControl ('PSA_CIVILITE'));
   sexe:= THValComboBox (GetControl ('PSA_SEXE'));
{$ENDIF}
   if (sexe <> nil) and (civilite <> nil) then
      begin
      if (sexe.value = '') then
         begin
         if (civilite.value = 'MLE') or (civilite.value = 'MME') then
            Setfield ('PSA_SEXE', 'F')
         else
         if civilite.value = 'MR' then
            Setfield ('PSA_SEXE', 'M');
         end
      else
      if sexe.value = 'M' then
         begin
         if (civilite.value = 'MLE') or (civilite.value = 'MME') then
            begin
            PGIBox ('Veuiller re-saisir la civilité du salarié !',
                    'Incohérence '+Civilite.text+' .. '+Sexe.text+' :');
            civilite.Value:= vide;
            end;
         end
      else
      if (sexe.value = 'F') then
         begin
         if civilite.value = 'MR' then
            begin
            PGIBox ('Veuiller re-saisir la civilité du salarié !',
                    'Incohérence '+Civilite.text+' .. '+Sexe.text+' :');
            civilite.Value:= vide;
            end;
         end;
      end;
   exit;
   end;

if (F.FieldName = 'PSA_NOMJF') then
   begin
{$IFNDEF EAGLCLIENT}
   sexe:= THDBValComboBox (GetControl ('PSA_SEXE'));
   civilite:= THDBValComboBox (GetControl ('PSA_CIVILITE'));
{$ELSE}
   sexe:= THValComboBox (GetControl ('PSA_SEXE'));
   civilite:= THValComboBox (GetControl ('PSA_CIVILITE'));
{$ENDIF}
   if (sexe <> nil) and (civilite <> nil) then
      begin
      if ((sexe.value = 'M') and (civilite.value = 'MLE')) and
         (Getfield ('PSA_NOMJF') <> '') then
         PGIBox ('Incohérence Nom de jeune fille, sexe et civilité !',
                 'Incohérence '+Sexe.text+', '+civilite.text+' :');
      if ((sexe.value = 'M') and (civilite.value <> 'MLE')) and
         (Getfield ('PSA_NOMJF') <> '') then
         PGIBox ('Incohérence Nom de jeune fille et sexe !',
                 'Incohérence '+Sexe.text+' :');
      if ((civilite.value = 'MLE') and (sexe.value <> 'M')) and
         (Getfield ('PSA_NOMJF') <> '') then
         PGIBox ('Incohérence Nom de jeune fille et civilité !',
                 'Incohérence '+civilite.text+' :');
      if ((sexe.value = 'M') or (civilite.value = 'MLE')) and
         (Getfield ('PSA_NOMJF') <> '') then
         Setfield ('PSA_NOMJF', vide);
      end;
   exit;
   end;

//ONGLET ETAT CIVIL
if (F.FieldName = 'PSA_NATIONALITE') then
   begin
//PT141B   if (GetField ('PSA_NATIONALITE') <> 'FRA') then
   Q:= OpenSql ('SELECT PY_MEMBRECEE FROM PAYS WHERE PY_PAYS ="'+
                GetField('PSA_NATIONALITE') + '"',True);
   cee := false;
   if Not Q.EoF then
      begin
      If Q.FindField('PY_MEMBRECEE').asString = 'X' then
         cee:= True;
      end;
   Ferme(Q);
   if (cee = False) then
      begin
      SetControlEnabled ('PSA_CARTESEJOUR', True);
      SetControlEnabled ('PSA_DELIVPAR', True);
      SetControlEnabled ('PSA_DATEXPIRSEJOUR', True);
      SetfocusControl ('PSA_DATENAISSANCE');
      end
   else
      begin
      if Getfield ('PSA_CARTESEJOUR') <> '' then
         SetField ('PSA_CARTESEJOUR', '');
      if Getfield ('PSA_DELIVPAR') <> '' then
         SetField ('PSA_DELIVPAR', '');
      SetControlEnabled ('PSA_CARTESEJOUR', FALSE);
      SetControlEnabled ('PSA_DELIVPAR', FALSE);
      SetControlEnabled ('PSA_DATEXPIRSEJOUR', FALSE);
      end;
   exit;
   end;

if (F.FieldName = ('PSA_DATENAISSANCE')) then
   begin
//Test Date entrée par rapport à date naissance
   if (GetField ('PSA_DATENAISSANCE') <> NULL) and
      (GetField ('PSA_DATEENTREE') <> NULL) then
      begin
      DEntree:= GetField ('PSA_DATEENTREE');
      DNaissance:= GetField ('PSA_DATENAISSANCE');
      if (DEntree <> idate1900) and (DNaissance <> idate1900) and
         (DEntree < DNaissance) then
         begin
         LastError:= 1;
         PgiBox ('La date de naissance ne peut être supérieur à la date'+
                 ' d''entrée.', Ecran.caption);
         end;
      end;
   exit;
   end;
// CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}
if (F.FieldName = ('PSA_DEPTNAISSANCE')) then
   begin
   if GetField ('PSA_DEPTNAISSANCE') = '2a' then
      SetField ('PSA_DEPTNAISSANCE', '2A');
   if GetField ('PSA_DEPTNAISSANCE') = '2b' then
      SetField ('PSA_DEPTNAISSANCE', '2B');
   DeptNaiss:= GetField ('PSA_DEPTNAISSANCE');
{PT138
   if Length (DeptNaiss) = 1 then
      begin
      PGIBox ('Attention, le département doit être saisi sur 2 caractères',
              'Département de naissance');
      SetFocusControl ('PSA_DEPTNAISSANCE');
      SetField ('PSA_DEPTNAISSANCE', '');
      end
   else
   if Length(deptNaiss) = 2 then
      begin
      if (Copy(DeptNaiss, 1, 1) = ' ') or (Copy(DeptNaiss, 2, 1) = ' ') then
         begin
         PGIBox ('Attention, vous ne pouvez pas saisir d''espaces dans le'+
                 ' département de naissance', 'Département de naissance');
         SetFocusControl ('PSA_DEPTNAISSANCE');
         SetField ('PSA_DEPTNAISSANCE', '');
         end;
      end;
}
// d PT152
    if (DS.State in [dsEdit]) or (DeptNaiss <> '') then
    begin
// f PT152
{PT179
      rep:= ControleCar (DeptNaiss, '4', True);
}
      rep:= ControleCar (DeptNaiss, '4', True, CInterdit);
//FIN PT179
      if (not ((Rep = 0) and (Length(DeptNaiss) = 2))) then
      begin
        if (DeptNaiss = '') then
          PGIBox('Le département de naissance est obligatoire', Ecran.caption)
        else
        begin
          PGIBox('Le département de naissance est incorrect', Ecran.caption);
          SetField('PSA_DEPTNAISSANCE', '');
        end;
        SetFocusControl('PSA_DEPTNAISSANCE');
      end
      else
      begin
        if (not (IsNumeric(DeptNaiss))) then
        begin
          if ((DeptNaiss <> '2A') and (DeptNaiss <> '2B')) then
          begin
            PGIBox('Le département de naissance est incorrect', Ecran.caption);
            SetField('PSA_DEPTNAISSANCE', '');
            SetFocusControl('PSA_DEPTNAISSANCE');
          end
          else
          begin
            Decodedate(GetField('PSA_DATENAISSANCE'), aa, mm, jj);
            if (aa < 1976) then
            begin
              PGIBox('Incohérence entre l''année de naissance et le' +
                ' département de naissance', Ecran.caption);
              SetField('PSA_DEPTNAISSANCE', '');
              SetFocusControl('PSA_DEPTNAISSANCE');
            end;
          end;
        end
        else
        begin
          if (DeptNaiss = '20') then
          begin
            Decodedate(GetField('PSA_DATENAISSANCE'), aa, mm, jj);
            if (aa >= 1976) then
            begin
              PGIBox('Incohérence entre l''année de naissance et le' +
                ' département de naissance', Ecran.caption);
              SetField('PSA_DEPTNAISSANCE', '');
              SetFocusControl('PSA_DEPTNAISSANCE');
            end;
          end;

          if (DeptNaiss = '96') then
          begin
            Decodedate(GetField('PSA_DATENAISSANCE'), aa, mm, jj);
            if (aa >= 1968) then
            begin
              PGIBox('Incohérence entre l''année de naissance et le' +
                ' département de naissance', Ecran.caption);
              SetField('PSA_DEPTNAISSANCE', '');
              SetFocusControl('PSA_DEPTNAISSANCE');
            end;
          end;
        end;
      end;
    end; // PT152
//FIN PT138

    TestPays := RechDom('TTPAYS', GetField('PSA_PAYSNAISSANCE'), False);
    if (GetField('PSA_PAYSNAISSANCE') = 'FRA') or (TestPays = 'FRANCE') then
    begin
      if GetField('PSA_DEPTNAISSANCE') = '99' then
      begin
        PGIBox('Le département de naissance doit être différent de 99 pour' +
          ' les salariés nés en France', Ecran.Caption);
        SetField('PSA_DEPTNAISSANCE', '');
        SetFocusControl('PSA_DEPTNAISSANCE');
      end;
    end;
    exit;
  end;
{$ENDIF}
// CCMX-CEGID ORGANIGRAMME DA
//PT138
// CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}
if (F.FieldName=('PSA_COMMUNENAISS')) then
   begin
   Valeur:= GetField ('PSA_COMMUNENAISS');
   if ((Valeur='') and ((GetField ('PSA_DEPTNAISSANCE')>='01') and
      (GetField ('PSA_DEPTNAISSANCE')<='98') or
      (GetField ('PSA_DEPTNAISSANCE')='2A') or
      (GetField ('PSA_DEPTNAISSANCE')='2B'))) then
      begin
      PGIBox ('La commune de naissance est obligatoire pour les salariés nés'+
              ' en France', Ecran.caption);
      SetFocusControl ('PSA_COMMUNENAISS');
      end;
{PT179
   rep:= ControleCar (Valeur, '9', True);
}
   rep:= ControleCar (Valeur, '9', True, CInterdit);
//FIN PT179
   if ((rep<>0) and (Valeur<>'')) then
      begin
{PT179
      PGIBox ('La commune de naissance est composée de caractères interdits',
              Ecran.caption);
}
      PGIBox ('Caractère '+CInterdit+' interdit dans la commune de naissance',
              Ecran.caption);
//FIN PT179
      SetFocusControl ('PSA_COMMUNENAISS');
      end
   else
      if (Valeur<>GetField ('PSA_COMMUNENAISS')) then
         SetField ('PSA_COMMUNENAISS', Valeur);
   exit;
   end;
//FIN PT138
{$ENDIF}
// CCMX-CEGID ORGANIGRAMME DA
  if (F.FieldName = ('PSA_PAYSNAISSANCE')) then
  begin
//Si le salarié est né à l'étranger, message si le département est différend de 99
    TestPays := RechDom('TTPAYS', GetField('PSA_PAYSNAISSANCE'), False);
    if (GetField('PSA_PAYSNAISSANCE') <> 'FRA') and (TestPays <> 'FRANCE') and
      (GetField('PSA_PAYSNAISSANCE') <> '') then
    begin
      if GetField('PSA_DEPTNAISSANCE') = '' then
        SetField('PSA_DEPTNAISSANCE', '99')
      else
      begin
        if GetField('PSA_DEPTNAISSANCE') <> '99' then
        begin
          reponse := PGIAsk('Le département vaut 99,#13#10#13#10' +
            'Voulez-vous le modifier?',
            'Le département de naissance est erronné : ' +
            GetField('PSA_DEPTNAISSANCE'));
          if reponse = MrYes then
            SetField('PSA_DEPTNAISSANCE', '99');
        end;
      end;
    end;
    exit;
  end;

  if (F.FieldName = 'PSA_PERSACHARGE') then
  begin
    Pa := TPageControl(GetControl('Pages'));
{$IFNDEF EAGLCLIENT}
    NbPerACharge := THDBSpinEdit(GetControl('PSA_PERSACHARGE'));
{$ELSE}
    NbPerACharge := TSpinEdit(GetControl('PSA_PERSACHARGE'));
{$ENDIF}
    if (NbPerACharge <> nil) and (PA <> nil) then
      if (Pa.ActivePage = Pa.Pages[1]) then
      begin
        QNbCharge := OpenSQL('SELECT COUNT (PEF_ACHARGE)' +
          ' FROM ENFANTSALARIE WHERE' +
          ' (PEF_SALARIE="' + GetField('PSA_SALARIE') + '") AND' +
          ' (PEF_ACHARGE="X") AND (PEF_TYPEPARENTAL="001")', TRUE); //PT158
        if not QNbCharge.EOF then
          Nb := QNbCharge.Fields[0].Asinteger
        else
          Nb := 0;
        if (Nb > 1) then
          SetControlText('LBLNBENFACHARGE', 'Dont ' + IntToStr(Nb) + ' enfants') //PT158
        else
          SetControlText('LBLNBENFACHARGE', 'Dont ' + IntToStr(Nb) + ' enfant'); //PT158
        Ferme(QNBCharge);

        if (Getfield('PSA_PERSACHARGE') < Nb) then
        begin
          PGIBox('Le nombre de personnes à charges ne peut être#13#10' +
            'inferieur au nombre d''enfant à charges !', 'Incohérence');
          NbPerACharge.Value := Nb;
        end;
      end;
    Exit;
  end;

//Onglet Emploi
  if (F.FieldName = ('PSA_DATEENTREE')) then
  begin
//Controle date entree
    if getfield('PSA_TYPDATANC') = '1' then
      AffectDateAcqCpAnc(getfield('PSA_TYPDATANC'));
    DEntree := GetField('PSA_DATEENTREE');
    DDPaie := Idate1900;
    DFPaie := Idate1900;
    DDPaieEntree := Idate1900;
    St := 'SELECT ##TOP 1## PPU_DATEDEBUT, PPU_DATEFIN, PPU_DATEENTREE, PPU_DATESORTIE' + //PT180
      ' FROM PAIEENCOURS' +
      ' WHERE PPU_SALARIE="' + GetField('PSA_SALARIE') + '"' +
      ' ORDER BY PPU_DATEFIN DESC';
    Q := OpenSql(St, True);
    if not Q.Eof then
    begin
      DDPaie := Q.FindField('PPU_DATEDEBUT').AsDateTime;
      DFPaie := Q.FindField('PPU_DATEFIN').AsDateTime;
      DDPaieEntree := Q.FindField('PPU_DATEENTREE').AsDateTime;
    end;
    Ferme(Q);
    if ((DEntree <= DFPaie) and (DEntree <> DDPaieEntree)) or
      (DEntree < DDPaieEntree) then
    begin
// PT118
      PgiBox('Vous avez généré un bulletin de paie pour la période du ' +
        DateToStr(DDPaie) + ' au ' + DateToStr(DFPaie) + '#13#10' +
        'avec une date d''entrée au ' + DateToStr(DDPaieEntree) + '.#13#10' +
        'Pour rester cohérent, vous devez saisir une date d''entrée' +
        ' postérieure#13#10à cette période ou supprimer vos bulletins' +
        ' correspondants.', 'Incohérence de la date d''entrée');
// PT118
      SetField('PSA_DATEENTREE', GblDtEntree);
    end;
    if (DS.State in [dsInsert]) then
    begin
      SetField('PSA_DATEANCIENNETE', Getfield('PSA_DATEENTREE'));
      SetField('PSA_ANCIENPOSTE', Getfield('PSA_DATEENTREE'));
    end;
    if (GetField('PSA_DATENAISSANCE') <> NULL) and
      (GetField('PSA_DATEENTREE') <> NULL) then
    begin
//Test Date entrée par rapport à date naissance
      DNaissance := GetField('PSA_DATENAISSANCE');
      if (DEntree <> idate1900) and (DNaissance <> idate1900) and
        (DEntree < DNaissance) then
      begin
        LastError := 1;
        PgiBox('La date d''entrée ne peut être antérieure à la date de' +
          ' naissance.', Ecran.Caption);
      end;
    end;
    exit;
  end;

  if (F.FieldName = ('PSA_DATESORTIE')) then
  begin
    DEntree := GetField('PSA_DATEENTREE');
    if GetField('PSA_DATESORTIE') = null then
      DSortie := IDate1900
    else
      DSortie := GetField('PSA_DATESORTIE');
    if (DSortie < Dentree) and (DSortie > IDate1900) then
    begin
      PGIBox('La date de sortie du salarié doit être postérieure à la date' +
        ' d''entrée!', Ecran.Caption);
      SetField('PSA_DATESORTIE', IDate1900);
    end;
    NouvelEntree := False;
    DDPaie := Idate1900;
    DFPaie := Idate1900;
    DDPaieEntree := Idate1900;
    DFPaieSortie := Idate1900;
    St := 'SELECT ##TOP 1## PPU_DATEDEBUT,PPU_DATEFIN,PPU_DATEENTREE,PPU_DATESORTIE' + //PT180
      ' FROM PAIEENCOURS' +
      ' WHERE PPU_SALARIE="' + GetField('PSA_SALARIE') + '"' +
      ' ORDER BY PPU_DATEFIN DESC';
    Q := OpenSql(St, True);
    if not Q.Eof then
    begin
      DDPaie := Q.FindField('PPU_DATEDEBUT').AsDateTime;
      DFPaie := Q.FindField('PPU_DATEFIN').AsDateTime;
      DDPaieEntree := Q.FindField('PPU_DATEENTREE').AsDateTime;
      DFPaieSortie := Q.FindField('PPU_DATESORTIE').AsDateTime;
    end
    else
      if GblDtSortie = Idate1900 then
        NouvelEntree := True;
    Ferme(Q);
    if (DDPaieEntree <> idate1900) and (DDPaieEntree <> DEntree) then
      NouvelEntree := True; //On test si il ne s'agit pas d'une nouvelle entree

//Controle date entree et date de sortie et paie déjà faite
    if (NouvelEntree = TRUE) and (GblDtSortie <> DSortie) then
    begin
      if (GetField('PSA_DATEENTREEPREC') <= Idate1900) or
        (GetField('PSA_DATESORTIEPREC') <= Idate1900) or
        (GetField('PSA_DATEENTREEPREC') = null) or
        (GetField('PSA_DATESORTIEPREC') = NULL) and
        (DS.State in [dsEdit]) then // PT136
//Recopie des date entrée et sortie uniquement en modif dans les zones précédentes
      begin
        if PgiAsk('Attention, il s''agit d''une nouvelle entrée :#13#10' +
          'Vous devez reporter vos dates d''entrée et sortie sur les' +
          ' zones date d''entrée et de sortie précédente.#13#10' +
          'Voulez-vous les copier?', Ecran.caption) = MrYes then
        begin
          DateActivee(THEDit(GetControl('ZPSA_DATEENTREEPREC')));
          DateActivee(THEDit(GetControl('ZPSA_DATESORTIEPREC')));
          SetField('PSA_DATEENTREEPREC', GblDtEntree);
          SetField('PSA_DATESORTIEPREC', GblDtSortie);
          Exit;
        end;
      end;
    end
    else
      if (NouvelEntree = False) and (VH_paie.PGCongesPayes) then
      begin
        if (GblDtSortie <> DSortie) and (DSortie > IDate1900) then
          if (DFPaie >= DSortie) and (DDPaie <> idate1900) and
            (DSortie <> DFPaieSortie) then
          begin
            if (DDPaie <= DSortie) then
              PgiBox('Vous avez généré un bulletin de paie pour la' +
                ' période du ' + DateToStr(DDPaie) + ' au ' +
                DateToStr(DFPaie) + '.#13#10' +
                'Vous devez supprimer et recréer le bulletin pour' +
                ' que le solde de tout compte soit intégré.',
                'Date de sortie');
            if (DDPaie > DSortie) and (DSortie > IDate1900) then
              PgiBox('Vous avez généré un bulletin de paie pour la' +
                ' période du ' + DateToStr(DDPaie) + ' au ' +
                DateToStr(DFPaie) + '.#13#10' +
                'Vous devez saisir une date de sortie postérieure à' +
                ' cette période.', 'Date de sortie');
            SetField('PSA_DATESORTIE', GblDtSortie);
          end;
        if (GblDtSortie <> idate1900) and (GblDtSortie <> DSortie) then
        begin
//Modification de la date de sortie avec solde de tt compte déja fait
          St := 'SELECT PCN_DATEDEBUT, PCN_DATEFIN' +
            ' FROM ABSENCESALARIE WHERE' +
            ' PCN_SALARIE="' + GetField('PSA_SALARIE') + '" AND' +
            ' PCN_DATEFIN="' + UsDateTime(GblDtSortie) + '"';
          Q := OpenSql(St, True);
          if not Q.eof then
          begin
            DD := DateToStr(Q.FindField('PCN_DATEDEBUT').AsDateTime);
            DF := DateToStr(Q.FindField('PCN_DATEFIN').AsDateTime);
            rep := PgiAsk('Attention, vous avez généré le solde de tout' +
              ' compte de ce salarié pour la période du ' + DD +
              ' au ' + DF + '.#13#10' +
              'Vous devez supprimer le bulletin de paie pour' +
              ' modifier la date de sortie.#13#10' +
              'S''il s''agit d''une nouvelle entrée, modifiez' +
              ' la date d''entrée au préalable.#13#10' +
              'Voulez-vous abandonner la modification de la' +
              ' date de sortie ?', 'Date de Sortie');
            if rep <> mrNo then
            begin
              SetField('PSA_DATESORTIE', GblDtSortie);
              Ferme(Q);
              exit;
            end;
          end;
          Ferme(Q);
        end;
      end;
    LibFinContrat := THLabel(GetControl('LBLCONTRATDF'));
//Contrôle date sortie avec date fin contrat
    if (LibFinContrat <> nil) and (LibFinContrat.Caption <> '') and
      (IsValidDate(LibFinContrat.Caption)) and (DSortie <> iDate1900) then
    begin
      DFinContrat := StrToDate(LibFinContrat.Caption);
      if (DFinContrat <> DSortie) and (DSortie > Idate1900) then
      begin
        PGIInfo('La date de sortie ne correspond pas à la date de fin du' +
          ' contrat en cours.', Ecran.caption);
// DEB PT115
        rep := PgiAsk('Voulez- vous modifier le contrat de travail avec la' +
          ' nouvelle date de sortie ?', Ecran.Caption);
        if rep = mrYes then
        begin
          st := 'UPDATE CONTRATTRAVAIL SET' +
            ' PCI_FINCONTRAT="' + UsDateTime(DSortie) + '",' +
            ' PCI_FINTRAVAIL="' + UsDateTime(DSortie) + '" WHERE' +   //PT187
            ' PCI_SALARIE="' + GetField('PSA_SALARIE') + '" AND' +
            ' PCI_ORDRE=' + IntToStr(NumOrdre);
          ExecuteSQL(st);
        end;
        if rep = mrNo then
        begin
          rep := PgiAsk('Voulez-vous accéder au contrat de travail du salarié ?',
            Ecran.Caption);
          if rep = mrYes then
            AglLanceFiche('PAY', 'MUL_CONTRAT', '', GetField('PSA_SALARIE'),
              GetField('PSA_SALARIE'));
        end;
// FIN PT115
      end;
    end;
    exit;
  end;

if (F.FieldName='PSA_MOTIFSORTIE') then
   begin
   if (GetField ('PSA_DATESORTIE')<Idate1900) and
      (GetField ('PSA_MOTIFSORTIE')<>'') then
      begin
      PgiBox ('Aucun motif de sortie ne peut être saisi car la date de sortie'+
              ' n''est pas une date valide', 'Motif de sortie');
{$IFNDEF EAGLCLIENT}
      ComboMotif:= THDBValCombobox (GetControl ('PSA_MOTIFSORTIE'));
{$ELSE}
      ComboMotif:= THValCombobox (GetControl ('PSA_MOTIFSORTIE'));
{$ENDIF}
      SetField ('PSA_MOTIFSORTIE', '');
      ComboMotif.Value:= '';
      end
//DEB PT186
   else
   if (GblMotifSortie<>GetField ('PSA_MOTIFSORTIE')) then
      begin
      st:= 'SELECT PCI_MOTIFSORTIE'+
           ' FROM CONTRATTRAVAIL WHERE'+
           ' PCI_SALARIE="'+GetField ('PSA_SALARIE')+'" AND'+
           ' PCI_ORDRE='+IntToStr (NumOrdre);
      Q:= OpenSQL (St, True);
      MotifFin:= '';
      if (not Q.Eof) then
        MotifFin := Q.FindField('PCI_MOTIFSORTIE').AsString;
      Ferme(Q);
      if (MotifFin <> GetField('PSA_MOTIFSORTIE')) then
      begin
        rep := PgiAsk('Voulez- vous modifier le contrat de travail avec le' +
          ' nouveau motif de sortie ?', Ecran.Caption);
        if rep = mrYes then
        begin
          st := 'UPDATE CONTRATTRAVAIL SET' +
            ' PCI_MOTIFSORTIE="' + GetField('PSA_MOTIFSORTIE') + '" WHERE' +
            ' PCI_SALARIE="' + GetField('PSA_SALARIE') + '" AND' +
            ' PCI_ORDRE=' + IntToStr(NumOrdre);
          ExecuteSQL(st);
          GblMotifSortie := GetField('PSA_MOTIFSORTIE');
        end;
        if rep = mrNo then
        begin
          rep := PgiAsk('Voulez-vous accéder au contrat de travail du salarié ?',
            Ecran.Caption);
          if rep = mrYes then
            AglLanceFiche('PAY', 'MUL_CONTRAT', '', GetField('PSA_SALARIE'),
              GetField('PSA_SALARIE'));
        end;
      end;
    end;
//FIN PT186
   end;

  if (F.FieldName = ('PSA_TYPPROFILANC')) then
  begin
    IdemEtab('PSA_TYPPROFILANC', 'PSA_PROFILANCIEN', 'ETB_PROFILANCIEN', F);
    exit;
  end;
// CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  if (((F.FieldName = ('PSA_DATEANCIENNETE')) or
    (F.FieldName = ('PSA_REGULANCIEN'))) and (Ecran.Name = 'SALARIE')) then
  begin
    Ladate := DebutDeMois(PlusMois(GetField('PSA_DATEANCIENNETE'),
      GetField('PSA_REGULANCIEN')));
{DEB PT135}
    DTemp := Idate1900;
{PT117 Modif requête, recherche de la période en cours et non la dernière paie}
    St := 'SELECT MAX(PEX_DEBUTPERIODE) AS DEBPAIECOURS' +
      ' FROM EXERSOCIAL WHERE' +
      ' PEX_ACTIF="X"';
    QRech := OpenSql(St, True);
    if (not Qrech.EOF) then
      DTemp := DebutDeMois(QRech.FindField('DEBPAIECOURS').AsDateTime);
    Ferme(QRech);
{FIN PT135}
    if GetField('PSA_DATESORTIE') <= idate1900 then {PT117 au 22/11/2005}
    begin
      if DTemp <> idate1900 then {PT135}
      begin
        DDPaie := Dtemp; {PT135}
        SetControlVisible('LBLNBREMOISANC', True);
      end
      else
      begin
        SetControlVisible('LBLNBREMOISANC', False);
        exit;
      end;
      Ferme(QRech);
    end
    else {PT117 au 22/11/2005 Sauf si salarié sorti, reprise date de sortie}
    begin
      DDPaie := GetField('PSA_DATESORTIE');
      if (DTemp <> idate1900) and (DTemp < DDPaie) then
        DDPaie := DTemp; {PT135}
      SetControlVisible('LBLNBREMOISANC', True);
    end;

    DiffMoisJour(LaDate, DDPaie, Nb, i);
    NbAnneeMois := Nb / 12;
    NbA := StrToInt(FloatToStr(Int(NbAnneeMois)));
    NbM := Round(Frac(NbAnneeMois) * 12);
    if (NbA = 0) then
    begin
      if (Nb = 0) then
        SetControlVisible('LBLNBREMOISANC', False)
      else
        St := 'Soit ' + IntToStr(Nb) + ' mois';
    end
    else
    begin
      St := 'Soit ' + IntToStr(NbA) + ' ans';
      if (NbM <> 0) then
        St := St + ' et ' + IntToStr(NbM) + ' mois';
      St := St + ' ou ' + IntToStr(Nb) + ' mois';
    end;
    SetControlText('LBLNBREMOISANC', St);
    exit;
  end;
// CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$ENDIF}

//Onglet Affectation
// DEB PT129
  if (F.FieldName = ('PSA_TYPCONVENTION')) then
  begin
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPCONVENTION', 'PSA_CONVENTION',
      'ETB_CONVENTION', 0);
{$IFNDEF EAGLCLIENT}
    Edit := THDBEdit(GetControl('PSA_TYPCONVENTION'));
{$ELSE}
    Edit := THEdit(GetControl('PSA_TYPCONVENTION'));
{$ENDIF}
    if Edit = nil then SetControlEnabled('PSA_CONVENTION', (GetField('PSA_TYPCONVENTION') <> 'ETB'))
    else if Edit.Color <> CouleurHisto then
    begin
      if GereHistoConvention then //PT192
        SetControlEnabled('PSA_CONVENTION', False)  //PT192
      else
        SetControlEnabled('PSA_CONVENTION', (GetField('PSA_TYPCONVENTION') <> 'ETB'));
    end;
    if GereHistoConvention then   //PT192
      SetControlVisible('BTNCONVPCL', False)   //PT192
    else
      SetControlVisible('BTNCONVPCL', (GetField('PSA_TYPCONVENTION') <> 'ETB'));
    exit;
  end;
// FIN PT129

  if (F.FieldName = 'PSA_CONVENTION') then
  begin
    LblConv := THLabel(GetControl('LBLPSA_CONVENTION'));
    if LblConv <> nil then
      SetControlProperty('LBLPSA_CONVENTION', 'Caption',
        RechDom('PGCONVENTION',
        GetControlText('PSA_CONVENTION'), FALSE));
    if Getfield('PSA_CONVENTION') <> '' then
    begin
{$IFNDEF EAGLCLIENT}
      Edit := THDBEdit(GetControl('PSA_QUALIFICATION'));
{$ELSE}
      Edit := THEdit(GetControl('PSA_QUALIFICATION'));
{$ENDIF}
      if Edit = nil then
      begin
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_QUALIFICATION', True);  //PT184
      end
      else if Edit.Color <> CouleurHisto then
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_QUALIFICATION', True); //PT184
{$IFNDEF EAGLCLIENT}
      Edit := THDBEdit(GetControl('PSA_COEFFICIENT'));
{$ELSE}
      Edit := THEdit(GetControl('PSA_COEFFICIENT'));
{$ENDIF}
      if Edit = nil then
      begin
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_COEFFICIENT', True);    //PT184
      end
      else if Edit.Color <> CouleurHisto then
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_COEFFICIENT', True);   //PT184
{$IFNDEF EAGLCLIENT}
      Edit := THDBEdit(GetControl('PSA_INDICE'));
{$ELSE}
      Edit := THEdit(GetControl('PSA_INDICE'));
{$ENDIF}
      if Edit = nil then
      begin
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_INDICE', True); //PT184
      end
      else if Edit.Color <> CouleurHisto then
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_INDICE', True); //PT184
{$IFNDEF EAGLCLIENT}
      Edit := THDBEdit(GetControl('PSA_NIVEAU'));
{$ELSE}
      Edit := THEdit(GetControl('PSA_NIVEAU'));
{$ENDIF}
      if Edit = nil then
      begin
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_NIVEAU', True);  //PT184
      end
      else if Edit.Color <> CouleurHisto then
        If JaiLeDroitTag(200053) then SetControlEnabled('PSA_NIVEAU', True); //PT184
{$IFNDEF EAGLCLIENT}
      ChampLib := THDBEdit(GetControl('PSA_QUALIFICATION'));
{$ELSE}
      ChampLib := THEdit(GetControl('PSA_QUALIFICATION'));
{$ENDIF}
      if ChampLib <> nil then
        //DEB PT183
        ChampLib.OnElipsisClick := QualifElipsisclick;
//        ChampLib.Plus := ' AND (PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '"' +
//          ' OR PMI_CONVENTION="000")';
        //FIN PT183
{$IFNDEF EAGLCLIENT}
      ChampLib := THDBEdit(GetControl('PSA_COEFFICIENT'));
{$ELSE}
      ChampLib := THEdit(GetControl('PSA_COEFFICIENT'));
{$ENDIF}
      if ChampLib <> nil then
        //DEB PT183
        ChampLib.OnElipsisClick := CoeffElipsisclick;
//        ChampLib.Plus := ' AND PMI_NATURE="COE" AND' + //PT163
//          ' (PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '" OR' +
//          ' PMI_CONVENTION="000")';
        //FIN PT183
{$IFNDEF EAGLCLIENT}
      ChampLib := THDBEdit(GetControl('PSA_INDICE'));
{$ELSE}
      ChampLib := THEdit(GetControl('PSA_INDICE'));
{$ENDIF}
      if ChampLib <> nil then
        //DEB PT183
        ChampLib.OnElipsisClick := IndiceElipsisclick;
//        ChampLib.Plus := ' AND (PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '" OR' +
//          ' PMI_CONVENTION="000")';
        //FIN PT183
{$IFNDEF EAGLCLIENT}
      ChampLib := THDBEdit(GetControl('PSA_NIVEAU'));
{$ELSE}
      ChampLib := THEdit(GetControl('PSA_NIVEAU'));
{$ENDIF}
      if ChampLib <> nil then
        //DEB PT183
        ChampLib.OnElipsisClick := NiveauElipsisclick;
//        ChampLib.Plus := ' AND (PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '" OR' +
//          ' PMI_CONVENTION="000")';
        //FIN PT183
    end;
    if ((Getfield('PSA_CONVENTION') = '') and (Ecran.Name = 'SALARIE')) then
    begin
      if (GetControlText('PSA_QUALIFICATION') <> '') then
        SetField('PSA_QUALIFICATION', vide);
      if (GetControlText('PSA_COEFFICIENT') <> '') then
        SetField('PSA_COEFFICIENT', vide);
      if (GetControlText('PSA_INDICE') <> '') then
        SetField('PSA_INDICE', vide);
      if (GetControlText('PSA_NIVEAU') <> '') then
        SetField('PSA_NIVEAU', vide);
      SetControlProperty('PSA_QUALIFICATION', 'Plus', vide);
      SetControlProperty('PSA_COEFFICIENT', 'Plus', vide);
      SetControlProperty('PSA_INDICE', 'Plus', vide);
      SetControlProperty('PSA_NIVEAU', 'Plus', vide);
      SetControlEnabled('PSA_QUALIFICATION', False);
      SetControlEnabled('PSA_COEFFICIENT', False);
      SetControlEnabled('PSA_INDICE', False);
      SetControlEnabled('PSA_NIVEAU', False);
    end;          
    exit;
  end;

// Affichage des libelles des tables
{Inutile de créer des tablettes pour récupérer les minimun conventionnels de
toutes les conventions et d'une convention précise. il s'agit de minimun cegid
ou standard, on insere qd même le prédefini au cas où!!}
  if (F.FieldName = 'PSA_QUALIFICATION') then
  begin
    Libelle := THLabel(GetControl('TPSA_QUALIFICATION'));
    if (Libelle <> nil) then
      if (Getfield('PSA_QUALIFICATION') <> '') then
      begin
        //PT183
        QRech := OpenSQL('Select PMI_LIBELLE FROM MINIMUMCONVENT M1' +
          ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="QUA"' +
          ' AND PMI_CODE="' + GetField('PSA_QUALIFICATION') + '"' +
          ' AND (PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" OR PMI_CONVENTION="000")' +
          ' AND (PMI_PREDEFINI="DOS"' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
          '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'")))))',True);
{PT183        QRech := OpenSql('SELECT PMI_LIBELLE' +
          ' FROM MINIMUMCONVENT WHERE' +
          ' PMI_NATURE="QUA" AND' +
          ' ##PMI_PREDEFINI## PMI_CODE="' + GetField('PSA_QUALIFICATION') + '" AND' +
          ' (PMI_CONVENTION="000" OR' +
          ' PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '")',
          True);}
        if (not QRech.EOF) and (QRech.Fields[0].AsString <> '') then
          Libelle.Caption := QRech.Fields[0].AsString
        else
          Libelle.Caption := 'Qualification inexistante!';
        Ferme(QRech);
      end
      else
        Libelle.Caption := '';
    exit;
  end;

  if (F.FieldName = 'PSA_COEFFICIENT') then
  begin
    Libelle := THLabel(GetControl('TPSA_COEFFICIENT'));
    if (Libelle <> nil) then
      if (Getfield('PSA_COEFFICIENT') <> '') then
      begin
        Libelle.Caption := '';
        //PT183
        QRech := OpenSQL('Select PMI_LIBELLE FROM MINIMUMCONVENT M1' +
          ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="COE" AND PMI_TYPENATURE="VAL"' +
          ' AND PMI_CODE="' + GetField('PSA_COEFFICIENT') + '"' +
          ' AND (PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" OR PMI_CONVENTION="000")' +
          ' AND (PMI_PREDEFINI="DOS"' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
          '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'")))))',True);
{PT183        QRech := OpenSql('SELECT PMI_LIBELLE' +
          ' FROM MINIMUMCONVENT WHERE' +
          ' PMI_NATURE="COE" AND' +
          ' PMI_TYPENATURE="VAL" AND' +
          ' ##PMI_PREDEFINI## PMI_CODE="' + GetField('PSA_COEFFICIENT') + '" AND' +
          ' (PMI_CONVENTION="000" OR' +
          ' PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '")',
          True);}
        if (not QRech.EOF) and
          (QRech.FindField('PMI_LIBELLE').AsString <> '') then
          Libelle.Caption := QRech.FindField('PMI_LIBELLE').AsString
        else
          Libelle.Caption := 'Coefficient inexistant!';
        Ferme(QRech);
      end
      else
        Libelle.Caption := '';
    exit;
  end;

  if (F.FieldName = 'PSA_INDICE') then
  begin
    Libelle := THLabel(GetControl('TPSA_INDICE'));
    if Libelle <> nil then
      if Getfield('PSA_INDICE') <> '' then
      begin
        //PT183
        QRech := OpenSQL('Select PMI_LIBELLE FROM MINIMUMCONVENT M1' +
          ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="IND"' +
          ' AND PMI_CODE="' + GetField('PSA_INDICE') + '"' +
          ' AND (PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" OR PMI_CONVENTION="000")' +
          ' AND (PMI_PREDEFINI="DOS"' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
          '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'")))))',True);
{PT183        QRech := OpenSql('SELECT PMI_LIBELLE' +
          ' FROM MINIMUMCONVENT WHERE' +
          ' PMI_NATURE="IND" AND' +
          ' ##PMI_PREDEFINI## PMI_CODE="' + GetField('PSA_INDICE') + '" AND' +
          ' (PMI_CONVENTION="000" OR' +
          ' PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '")',
          True); }
        if (not QRech.EOF) and (QRech.Fields[0].AsString <> '') then
          Libelle.Caption := QRech.Fields[0].AsString
        else
          Libelle.Caption := 'Indice inexistant!';
        Ferme(QRech);
      end
      else
        Libelle.Caption := '';
    exit;
  end;

  if (F.FieldName = 'PSA_NIVEAU') then
  begin
    Libelle := THLabel(GetControl('TPSA_NIVEAU'));
    if Libelle <> nil then
      if Getfield('PSA_NIVEAU') <> '' then
      begin
        //PT183
        QRech := OpenSQL('Select PMI_LIBELLE FROM MINIMUMCONVENT M1' +
          ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="NIV"' +
          ' AND PMI_CODE="' + GetField('PSA_NIVEAU') + '"' +
          ' AND (PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" OR PMI_CONVENTION="000")' +
          ' AND (PMI_PREDEFINI="DOS"' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
          ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
          '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
          '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
          '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'")))))',True);
{PT183        QRech := OpenSql('SELECT PMI_LIBELLE' +
          ' FROM MINIMUMCONVENT WHERE' +
          ' PMI_NATURE="NIV" AND' +
          ' ##PMI_PREDEFINI## PMI_CODE="' + GetField('PSA_NIVEAU') + '" AND' +
          ' (PMI_CONVENTION="000" OR' +
          ' PMI_CONVENTION="' + GetField('PSA_CONVENTION') + '")',
          True);}
        if (not Qrech.EOF) and (QRech.Fields[0].AsString <> '') then
          Libelle.Caption := QRech.Fields[0].AsString
        else
          Libelle.Caption := 'Niveau inexistant!';
        Ferme(QRech);
      end
      else
        Libelle.Caption := '';
    exit;
  end;

  IdemEtab('PSA_TYPEDITORG', 'PSA_EDITORG', 'ETB_EDITORG', F);

{ ******************** Gestion des Champs Idem Etablissement ******************}

// Onglets Autre profil
  if (F.FieldName = ('PSA_TYPPROFILAFP')) then
  begin
    IdemEtab('PSA_TYPPROFILAFP', 'PSA_PROFILAFP', 'ETB_PROFILAFP', F);
    IdemEtab('PSA_TYPPROFILAFP', 'PSA_PCTFRAISPROF', 'ETB_PCTFRAISPROF', F);
    exit;
  end;

  if (F.FieldName = ('PSA_TYPPROFILFNAL')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPPROFILFNAL'));
    ComboMaj := THDBValCombobox(GetControl('PSA_PROFILFNAL'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPPROFILFNAL'));
    ComboMaj := THValCombobox(GetControl('PSA_PROFILFNAL'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'DOS' then
      begin
        if (ComboMaj.value <> VH_Paie.PGProfilFnal) then
          ComboMaj.value := VH_Paie.PGProfilFnal;
        ComboMaj.enabled := False;
      end
      else
        if ComboMaj.Color <> CouleurHisto then ComboMaj.enabled := True;
    exit;
  end;

//Traitement Idem etab pour Profilrem
//Traitement champ retraite et idem etab
  IdemEtab('PSA_TYPPROFIL', 'PSA_PROFIL', 'ETB_PROFIL', F);
  IdemEtab('PSA_TYPPERIODEBUL', 'PSA_PERIODBUL', 'ETB_PERIODBUL', F);
  IdemEtab('PSA_TYPPROFILREM', 'PSA_PROFILREM', 'ETB_PROFILREM', F);
  IdemEtab('PSA_TYPPROFILRBS', 'PSA_PROFILRBS', 'ETB_PROFILRBS', F);
  IdemEtab('PSA_TYPREDREPAS', 'PSA_REDREPAS', 'ETB_REDREPAS', F);
  IdemEtab('PSA_TYPREDRTT1', 'PSA_REDRTT1', 'ETB_REDRTT1', F);
  IdemEtab('PSA_TYPREDRTT2', 'PSA_REDRTT2', 'ETB_REDRTT2', F);
  IdemEtab('PSA_TYPPROFILRET', 'PSA_PROFILRET', 'ETB_PROFILRET', F);
  IdemEtab('PSA_TYPPROFILMUT', 'PSA_PROFILMUT', 'ETB_PROFILMUT', F);
  IdemEtab('PSA_TYPPROFILPRE', 'PSA_PROFILPRE', 'ETB_PROFILPRE', F);
  IdemEtab('PSA_TYPPROFILTSS', 'PSA_PROFILTSS', 'ETB_PROFILTSS', F);
  IdemEtab('PSA_TYPPROFILTRANS', 'PSA_PROFILTRANS', 'ETB_PROFILTRANS', F);
  IdemEtab('PSA_TYPPROFILAPP', 'PSA_PROFILAPP', 'ETB_PROFILAPP', F);
  IdemEtab('PSA_TYPACTIVITE', 'PSA_ACTIVITE', 'ETB_ACTIVITE', F); //PT68

// ONGLET CONTRAT
  if (F.FieldName = ('PSA_TYPJOURHEURE')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPJOURHEURE'));
    ComboMaj := THDBValCombobox(GetControl('PSA_JOURHEURE'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPJOURHEURE'));
    ComboMaj := THValCombobox(GetControl('PSA_JOURHEURE'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
    begin
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_JOURHEURE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_JOURHEURE');
        ComboMaj.enabled := False;
      end
      else
        if ComboEtab.value = 'DOS' then
        begin
          ComboMaj.value := VH_Paie.PGJourHeure;
          ComboMaj.enabled := False;
        end
        else
          if ComboEtab.value = 'PER' then
            ComboMaj.enabled := True;
    end;
    exit;
  end;

  if (F.FieldName = ('PSA_STANDCALEND')) then
  begin
{ Message d'alerte pour la personnalisation du calendrier :
une partie du code a été dérivé dans une procédure evenementiel
sur le OnChange de l'objet PSA_STANDCALEND..}
    IdemEtab('PSA_STANDCALEND', 'PSA_CALENDRIER', 'ETB_STANDCALEND', F);
{$IFDEF EPRIMES}
    SetcontrolEnabled('PSA_CALENDRIER', False);
{$ENDIF}
    MLabel := ThLabel(getcontrol('TPSA_CALENDRIER'));
    if MLabel <> nil then
      if GetField('PSA_STANDCALEND') <> 'ETS' then
        MLabel.Caption := 'Basé sur le calendrier standard'
      else
        MLabel.caption := 'Sélection du calendrier';
    exit;
  end;

{Suppression du calendrier si réinitialiser à Aucun}
  if (F.FieldName = ('PSA_CALENDRIER')) then
  begin
    if (GblCalend <> GetField('PSA_CALENDRIER')) and
      (GetField('PSA_CALENDRIER') = '') then
      SupTablesLiees('CALENDRIER', 'ACA_SALARIE', getfield('PSA_SALARIE'), '',
        True);
    if (DS.State in [dsEdit, DsBrowse]) then { PT133 }
      SetControlEnabled('TCALENDRIER', (GetField('PSA_CALENDRIER') <> ''));
  end;

// Onglet Congés payés
  if F.Fieldname = 'PSA_PROFILCGE' then
  begin
    EnabledZonesAnciennete;
    exit;
  end;

  IdemEtab('PSA_TYPPROFILCGE', 'PSA_PROFILCGE', 'ETB_PROFILCGE', F);

// Base congés supplémentaire ancienneté
  if F.FieldName = 'PSA_BASANCCP' then
  begin
    EnabledZonesAnciennete;
    exit;
  end;

  if (F.FieldName = ('PSA_TYPEDITBULCP')) then
  begin
    IdemEtab('PSA_TYPEDITBULCP', 'PSA_EDITBULCP', 'ETB_EDITBULCP', F);
    exit;
  end;

  if (F.FieldName = ('PSA_CPACQUISMOIS')) then
  begin
    IdemEtab('PSA_CPACQUISMOIS', 'PSA_NBREACQUISCP', 'ETB_NBREACQUISCP', F);
    SetControlEnabled('PSA_SSDECOMPTE', GetField('PSA_CPACQUISMOIS') <> 'ETB');
    if (GetField('PSA_CPACQUISMOIS') = 'ETB') and
      (GetField('PSA_SSDECOMPTE') <> '-') then
      SetField('PSA_SSDECOMPTE', '-');
    exit;
  end;

  if (F.FieldName = ('PSA_CPACQUISSUPP')) then
  begin
    IdemEtab('PSA_CPACQUISSUPP', 'PSA_NBRECPSUPP', 'ETB_NBRECPSUPP', F);
    exit;
  end;

  if (F.FieldName = 'PSA_NBRECPSUPP') then
  begin
    if (HEnt1.Valeur(GetField('PSA_NBRECPSUPP')) < 0) and
      (GetField('PSA_CPACQUISSUPP') = 'PER') then
    begin
      PgiBox('Le nombre de jours acquis supplémentaire doit être positif.',
        Ecran.caption);
      if Hent1.Valeur(GetField('PSA_NBRECPSUPP')) <> 0 then
        SetField('PSA_NBRECPSUPP', '0');
    end;
    exit;
  end;

  if (F.FieldName = 'PSA_TYPNBACQUISCP') then
  begin
    if GetField('PSA_TYPNBACQUISCP') = 'ETB' then
    begin
      if (T <> nil) then
        if (T.GetValue('ETB_NBACQUISCP') <> GetField('PSA_NBACQUISCP')) then
          Setfield('PSA_NBACQUISCP', T.GetValue('ETB_NBACQUISCP'));
      SetControlEnabled('PSA_NBACQUISCP', False);
    end
    else
      SetControlEnabled('PSA_NBACQUISCP', True);
    exit;
  end;

  if (F.FieldName = ('PSA_DATANC')) then
  begin
    IdemEtab('PSA_DATANC', 'PSA_TYPDATANC', 'ETB_TYPDATANC', F);
    if (GetField('PSA_DATANC') = 'ETB') then
    begin
      if (T <> nil) then
      begin
{On recupère la date d'acq Anc que si ce n'est pas la date entrée salarié }
        if (GetField('PSA_TYPDATANC') <> '1') and
          (T.GetValue('ETB_DATEACQCPANC') <> GetField('PSA_DATEACQCPANC')) then
          if (GetControl('PSA_DATEACQCPANC') <> nil) then
            setfield('PSA_DATEACQCPANC', T.GetValue('ETB_DATEACQCPANC'));
        SetControlEnabled('PSA_DATEACQCPANC', False);
      end;
      Firstcp := true;
    end
    else
      SetControlEnabled('PSA_DATEACQCPANC', True);
    EnabledZonesAnciennete;
    exit;
  end;

  if (F.FieldName = ('PSA_TYPDATANC')) then
  begin
    AffectDateAcqCpAnc(getfield('PSA_TYPDATANC'));
    EnabledZonesAnciennete;
    exit;
  end;

// CCMX-CEGID Le 29.01.2006
{$IFNDEF GCGC}
  if ((F.FieldName = 'PSA_DATEACQCPANC') and (Getfield('PSA_BASANCCP') <> '')) then
  begin { DEB PT146 }
    if (GetField('PSA_DATEACQCPANC') <> '') then
      if not PgOkFormatDateJJMM(Copy(Getfield('PSA_DATEACQCPANC'), 1, 2) + Copy(Getfield('PSA_DATEACQCPANC'), 3, 2)) then
      begin
        PgiBox('Format date incorrect. Vous devez saisir une date au format jj/mm.', Ecran.caption);
        if GetField('PSA_DATEACQCPANC') <> '' then setfield('PSA_DATEACQCPANC', '');
      end; { FIN PT146 }

    exit;
  end;
{$ENDIF}
// FIN CCMX-CEGID Le 29.01.2006
  if (F.FieldName = ('PSA_CPTYPEMETHOD')) then
  begin
    IdemEtab('PSA_CPTYPEMETHOD', 'PSA_VALORINDEMCP', 'ETB_VALORINDEMCP', F);
    exit;
  end;

  if (F.FieldName = ('PSA_CPTYPERELIQ')) then
  begin
    IdemEtab('PSA_CPTYPERELIQ', 'PSA_RELIQUAT', 'ETB_RELIQUAT', F);
    exit;
  end;

{DEB PT126}
  if (F.FieldName = ('PSA_TYPPAIEVALOMS')) then
  begin
    IdemEtab('PSA_TYPPAIEVALOMS', 'PSA_PAIEVALOMS', 'ETB_PAIEVALOMS', F);
    exit;
  end;
{FIN PT126}

  if (F.FieldName = ('PSA_CPACQUISANC')) then
  begin
    if GetField('PSA_CPACQUISANC') = 'ETB' then
    begin
      if (T <> nil) then
        if (T.GetValue('ETB_BASANCCP') <> GetField('PSA_BASANCCP')) then
          if (GetControl('PSA_BASANCCP') <> nil) then
            SetField('PSA_BASANCCP', T.GetValue('ETB_BASANCCP'));
      SetControlEnabled('PSA_BASANCCP', False);
      Firstcp := true;
      if T <> nil then
        if (T.GetValue('ETB_VALANCCP') <> GetField('PSA_VALANCCP')) then
          if (GetControl('PSA_VALANCCP') <> nil) then
            Setfield('PSA_VALANCCP', T.GetValue('ETB_VALANCCP'));
      SetControlEnabled('PSA_VALANCCP', False);
    end
    else
      SetControlEnabled('PSA_BASANCCP', True);
    EnabledZonesAnciennete;
    exit;
  end;
// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}
  if F.fieldName = 'PSA_VALANCCP' then
  begin
    if (GetField('PSA_BASANCCP') = 'VAL') and
      (GetField('PSA_VALANCCP') <> '') then
      if not IsOkFloat(GetField('PSA_VALANCCP')) then
        SetFocuscontrol('PSA_VALANCCP');
    exit;
  end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

//Gestion salarié de la méthode de valorisation des CP
  if (F.FieldName = ('PSA_CPTYPEVALO')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_CPTYPEVALO'));
    ComboMaj := THDBValCombobox(GetControl('PSA_MVALOMS'));
    EditMaj := THDBEdit(GetControl('PSA_VALODXMN'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_CPTYPEVALO'));
    ComboMaj := THValCombobox(GetControl('PSA_MVALOMS'));
    EditMaj := THEdit(GetControl('PSA_VALODXMN'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) and (EditMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
        begin
          if (T.GetValue('ETB_MVALOMS') <> ComboMaj.Value) then
            setfield('PSA_MVALOMS', T.GetValue('ETB_MVALOMS'));
          ComboMaj.enabled := False;
          if (FloatToStr(T.GetValue('ETB_VALODXMN')) <> EditMaj.Text) then
            setfield('PSA_VALODXMN', T.GetValue('ETB_VALODXMN'));
          EditMaj.enabled := False;
        end;
      end
      else
      begin
        EditMaj.enabled := (getfield('PSA_MVALOMS') = 'M');
        ComboMaj.enabled := True;
      end;
    exit;
  end;

  if (F.fieldName = 'PSA_MVALOMS') and (GetControl('PSA_MVALOMS') <> nil) then
  begin
    SetcontrolEnabled('PSA_VALODXMN', (getfield('PSA_MVALOMS') = 'M'));
    if (getfield('PSA_MVALOMS') = 'R') then
      Setfield('PSA_VALODXMN', 0);
    if (getfield('PSA_MVALOMS') = 'T') and (T <> nil) then
    begin
      if T.GetValue('ETB_NBJOUTRAV') = 5 then
        Setfield('PSA_VALODXMN', 21.67)
      else
        Setfield('PSA_VALODXMN', 26);
    end;
  end;
// Fin onglet congés payés

//Fiche Banque : Gestion des champs idem Etablissement,idem Profil
  if (F.FieldName = ('PSA_TYPREGLT')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPREGLT'));
    ComboMaj := THDBValCombobox(GetControl('PSA_PGMODEREGLE'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPREGLT'));
    ComboMaj := THValCombobox(GetControl('PSA_PGMODEREGLE'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PGMODEREGLE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PGMODEREGLE');
        ComboMaj.enabled := False;
      end
      else
        ComboMaj.enabled := True;
    exit;
  end;

  if (F.FieldName = ('PSA_TYPDATPAIEMENT')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPDATPAIEMENT'));
    ComboMaj := THDBValCombobox(GetControl('PSA_MOISPAIEMENT'));
    SpinEditMaj := THDBSpinEdit(GetControl('PSA_JOURPAIEMENT'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPDATPAIEMENT'));
    ComboMaj := THValCombobox(GetControl('PSA_MOISPAIEMENT'));
    SpinEditMaj := TSpinEdit(GetControl('PSA_JOURPAIEMENT'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) and (SpinEditMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_MOISPAIEMENT') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_MOISPAIEMENT');
        ComboMaj.enabled := False;
        if (T <> nil) then
          if (T.GetValue('ETB_JOURPAIEMENT') <> SpinEditMaj.value) then
          begin
            ValCoEtabChamp := T.GetValue('ETB_JOURPAIEMENT');
            if (ValCoEtabChamp <> '') then
              SpinEditMaj.value := StrToInt(ValCoEtabChamp);
          end;
        SpinEditMaj.enabled := False;
      end
      else
      begin
        ComboMaj.enabled := True;
        SpinEditMaj.enabled := True;
      end;
    exit;
  end;
{$IFNDEF GCGC}
  if (F.FieldName = ('PSA_TYPVIRSOC')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPVIRSOC'));
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBVIRSOC'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPVIRSOC'));
    ComboMaj := THValCombobox(GetControl('PSA_RIBVIRSOC'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
    begin
      if (ComboEtab.value = 'ETB') then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_RIBSALAIRE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_RIBSALAIRE');
        ComboMaj.enabled := False;
      end;
      if (ComboEtab.value = 'PRO') then
      begin
        if (TDO <> nil) then
          if (TDO.GetValue('PDO_RIBASSOCIE') <> ComboMaj.value) then
            ComboMaj.value := TDO.GetValue('PDO_RIBASSOCIE');
        ComboMaj.enabled := False;
      end;
      if ComboEtab.value = 'PER' then
        ComboMaj.enabled := True;
    end;
    LblRibSoc := THLabel(GetControl('LBLRIBORDRESAL'));
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QRibSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;

  if (F.FieldName = ('PSA_TYPPAIACOMPT')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPPAIACOMPT'));
    ComboMaj := THDBValCombobox(GetControl('PSA_PAIACOMPTE'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPPAIACOMPT'));
    ComboMaj := THValCombobox(GetControl('PSA_PAIACOMPTE'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PAIACOMPTE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PAIACOMPTE');
        ComboMaj.enabled := False;
      end
      else
        ComboMaj.enabled := True;
    exit;
  end;

  if (F.FieldName = ('PSA_TYPACPSOC')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPACPSOC'));
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBACPSOC'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPACPSOC'));
    ComboMaj := THValCombobox(GetControl('PSA_RIBACPSOC'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
    begin
      if (ComboEtab.value = 'ETB') then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_RIBACOMPTE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_RIBACOMPTE');
        ComboMaj.enabled := False;
      end;

      if (ComboEtab.value = 'PRO') then
      begin
        if (TDO <> nil) then
          if (TDO.GetValue('PDO_RIBASSOCIE') <> ComboMaj.value) then
            ComboMaj.value := TDO.GetValue('PDO_RIBASSOCIE');
        ComboMaj.enabled := False;
      end;
      if ComboEtab.value = 'PER' then
        ComboMaj.enabled := True;
    end;

    LblRibSoc := THLabel(GetControl('LBLRIBORDREACPT'));
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QribSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;

  if (F.FieldName = ('PSA_TYPPAIFRAIS')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPPAIFRAIS'));
    ComboMaj := THDBValCombobox(GetControl('PSA_PAIFRAIS'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPPAIFRAIS'));
    ComboMaj := THValCombobox(GetControl('PSA_PAIFRAIS'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PAIFRAIS') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PAIFRAIS');
        ComboMaj.enabled := False;
      end
      else
        ComboMaj.enabled := True;
    exit;
  end;

  if (F.FieldName = ('PSA_TYPFRAISSOC')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPFRAISSOC'));
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBFRAISSOC'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPFRAISSOC'));
    ComboMaj := THValCombobox(GetControl('PSA_RIBFRAISSOC'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
    begin
      if (ComboEtab.value = 'ETB') then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_RIBFRAIS') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_RIBFRAIS');
        ComboMaj.enabled := False;
      end;

      if (ComboEtab.value = 'PRO') and (TDO <> nil) then
      begin
        if (TDO.GetValue('PDO_RIBASSOCIE') <> ComboMaj.value) then
          ComboMaj.value := TDO.GetValue('PDO_RIBASSOCIE');
        ComboMaj.enabled := False;
      end;
      if ComboEtab.value = 'PER' then
        ComboMaj.enabled := True;
    end;
    LblRibSoc := THLabel(GetControl('LBLRIBORDREFP'));
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QRibSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;

  if (F.FieldName = ('PSA_RIBVIRSOC')) then
  begin
    LblRibSoc := THLabel(GetControl('LBLRIBORDRESAL'));
{$IFNDEF EAGLCLIENT}
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBVIRSOC'));
{$ELSE}
    ComboMaj := THValCombobox(GetControl('PSA_RIBVIRSOC'));
{$ENDIF}
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QRibSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;

  if (F.FieldName = ('PSA_RIBACPSOC')) then
  begin
    LblRibSoc := THLabel(GetControl('LBLRIBORDREACPT'));
{$IFNDEF EAGLCLIENT}
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBACPSOC'));
{$ELSE}
    ComboMaj := THValCombobox(GetControl('PSA_RIBACPSOC'));
{$ENDIF}
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QRibSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;

  if (F.FieldName = ('PSA_RIBFRAISSOC')) then
  begin
    LblRibSoc := THLabel(GetControl('LBLRIBORDREFP'));
{$IFNDEF EAGLCLIENT}
    ComboMaj := THDBValCombobox(GetControl('PSA_RIBFRAISSOC'));
{$ELSE}
    ComboMaj := THValCombobox(GetControl('PSA_RIBFRAISSOC'));
{$ENDIF}
    if (LblRibSoc <> nil) and (ComboMaj <> nil) then
      if ComboMaj.Value <> '' then
      begin
        StPlus := PGBanqueCP(True); //PT148
        QRibSoc := OpenSql('SELECT BQ_ETABBQ, BQ_NUMEROCOMPTE, BQ_GUICHET,' +
          ' BQ_CLERIB' +
          ' FROM BANQUECP WHERE' +
          ' BQ_GENERAL="' + ComboMaj.Value + '"' + StPlus, True);
        if not QRibSoc.EOF then
          LblRibSoc.Caption := 'Banque : ' + QRibSoc.FindField('BQ_ETABBQ').asstring +
            '  Guichet : ' + QRibSoc.FindField('BQ_GUICHET').asstring +
            '  Compte : ' + QRibSoc.FindField('BQ_NUMEROCOMPTE').asstring +
            '  Clé : ' + QRibSoc.FindField('BQ_CLERIB').asstring
        else
          LblRibSoc.Caption := 'Banque inexistante';
        Ferme(QRibSoc);
      end
      else
        LblRibSoc.Caption := '';
    exit;
  end;
{$ENDIF}
// Onglet DADS-U
  if (F.FieldName = ('PSA_DADSFRACTION')) then
  begin
    if ((GetField('PSA_DADSFRACTION') <> '') and
      (GetField('PSA_DADSFRACTION') > Getparamsoc('SO_PGFRACTION'))) then
    begin
      PGIBox('La valeur saisie est supérieure à ' + Getparamsoc('SO_PGFRACTION') +
        '#13#10qui est la valeur maximale des paramètres société',
        'Fraction DADS');
      SetField('PSA_DADSFRACTION', '1');
    end;
  end;
  //DEB PT164
  if (F.FieldName = ('PSA_TYPEREGIME')) then
  begin
    if (GetField('PSA_TYPEREGIME') = 'X') then
    begin
      SetControlEnabled('PSA_REGIMESS', False);
      SetField('PSA_REGIMESS', '');
      SetControlEnabled('PSA_REGIMEMAL', True);
      SetControlEnabled('PSA_REGIMEAT', True);
      SetControlEnabled('PSA_REGIMEVIP', True);
      SetControlEnabled('PSA_REGIMEVIS', True);
      if (GetField('PSA_REGIMEMAL') = '') then SetField('PSA_REGIMEMAL', '200');
      if (GetField('PSA_REGIMEAT') = '') then SetField('PSA_REGIMEAT', '200');
      if (GetField('PSA_REGIMEVIP') = '') then SetField('PSA_REGIMEVIP', '200');
      if (GetField('PSA_REGIMEVIS') = '') then SetField('PSA_REGIMEVIS', '200');
    end
    else
    begin
      If JaiLeDroitTag(200057) then SetControlEnabled('PSA_REGIMESS', True);  //PT184
      if (GetField('PSA_REGIMESS') = '') then SetField('PSA_REGIMESS', '200');
      SetControlEnabled('PSA_REGIMEMAL', False);
      SetControlEnabled('PSA_REGIMEAT', False);
      SetControlEnabled('PSA_REGIMEVIP', False);
      SetControlEnabled('PSA_REGIMEVIS', False);
      if (GetField('PSA_REGIMEMAL') <> '') then SetField('PSA_REGIMEMAL', '');
      if (GetField('PSA_REGIMEAT') <> '') then SetField('PSA_REGIMEAT', '');
      if (GetField('PSA_REGIMEVIP') <> '') then SetField('PSA_REGIMEVIP', '');
      if (GetField('PSA_REGIMEVIS') <> '') then SetField('PSA_REGIMEVIS', '');
    end;
  end;

  if (F.FieldName = ('PSA_REGIMESS')) then
  begin
    if (GetField('PSA_REGIMESS') = '147') or (GetField('PSA_REGIMESS') = '148') or (GetField('PSA_REGIMESS') = '149') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime SS : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMESS'), False) + ' n''est pas autorisée.'), Ecran.caption);
      SetFocusControl('PSA_REGIMESS');
      exit;
    end;
  end;

  if (F.FieldName = ('PSA_REGIMEMAL')) then
  begin
    if (GetField('PSA_REGIMEMAL') = '147') or (GetField('PSA_REGIMEMAL') = '148') or (GetField('PSA_REGIMEMAL') = '150') or (GetField('PSA_REGIMEMAL') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire risque maladie : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEMAL'), False) + ' n''est pas autorisée.'), Ecran.caption);
      SetFocusControl('PSA_REGIMEMAL');
      exit;
    end;
  end;

  if (F.FieldName = ('PSA_REGIMEVIP')) then
  begin
    if (GetField('PSA_REGIMEVIP') = '150') or (GetField('PSA_REGIMEVIP') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire vieillesse (PP) : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEVIP'), False) + ' n''est pas autorisée.'), Ecran.caption);
      SetFocusControl('PSA_REGIMEVIP');
      exit;
    end;
  end;

  if (F.FieldName = ('PSA_REGIMEVIS')) then
  begin
    if (GetField('PSA_REGIMEVIS') = '150') or (GetField('PSA_REGIMEVIS') = '160') then
    begin
      LastError := 1;
      PgiBox(TraduireMemoire('La valeur du Régime obligatoire vieillesse (PS) : ' + RechDom('PGREGIMESS', GetControlText('PSA_REGIMEVIS'), False) + ' n''est pas autorisée.'), Ecran.caption);
      SetFocusControl('PSA_REGIMEVIS');
      exit;
    end;
  end;
  //FIN PT164

//Gestion des Champs Idem Etablissement PRUD'HOMMAUX
  if (F.FieldName = ('PSA_TYPDADSFRAC')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPDADSFRAC'));
    ComboMaj := THDBValCombobox(GetControl('PSA_DADSFRACTION'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPDADSFRAC'));
    ComboMaj := THValCombobox(GetControl('PSA_DADSFRACTION'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_CODESECTION') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_CODESECTION');
        ComboMaj.enabled := False;
      end
      else
        if JaiLeDroitTag(200001) then ComboMaj.enabled := True;  //PT184
    exit;
  end;

  if (F.FieldName = ('PSA_TYPPRUDH')) then
  begin
{$IFNDEF EAGLCLIENT}
    ComboEtab := THDBValComboBox(GetControl('PSA_TYPPRUDH'));
    ComboMaj := THDBValCombobox(GetControl('PSA_PRUDHCOLL'));
{$ELSE}
    ComboEtab := THValComboBox(GetControl('PSA_TYPPRUDH'));
    ComboMaj := THValCombobox(GetControl('PSA_PRUDHCOLL'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PRUDHCOLL') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PRUDHCOLL');
        ComboMaj.enabled := False;
      end
      else
        if JaiLeDroitTag(200001) then ComboMaj.enabled := True;  //PT184

{$IFNDEF EAGLCLIENT}
    ComboMaj := THDBValCombobox(GetControl('PSA_PRUDHSECT'));
{$ELSE}
    ComboMaj := THValCombobox(GetControl('PSA_PRUDHSECT'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PRUDHSECT') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PRUDHSECT');
        ComboMaj.enabled := False;
      end
      else
        if JaiLeDroitTag(200001) then ComboMaj.enabled := True;  //PT184

{$IFNDEF EAGLCLIENT}
    ComboMaj := THDBValCombobox(GetControl('PSA_PRUDHVOTE'));
{$ELSE}
    ComboMaj := THValCombobox(GetControl('PSA_PRUDHVOTE'));
{$ENDIF}
    if (ComboEtab <> nil) and (ComboMaj <> nil) then
      if ComboEtab.value = 'ETB' then
      begin
        if (T <> nil) then
          if (T.GetValue('ETB_PRUDHVOTE') <> ComboMaj.value) then
            ComboMaj.value := T.GetValue('ETB_PRUDHVOTE');
        ComboMaj.enabled := False;
      end
      else
        if JaiLeDroitTag(200001) then ComboMaj.enabled := True;  //PT184
    exit;
  end;

  if (F.FieldName = ('PSA_CONDEMPLOI')) then
  begin
    if GetField('PSA_CONDEMPLOI') = 'C' then
      SetControlEnabled('PSA_TAUXPARTIEL', False)
    else
    begin
{$IFNDEF EAGLCLIENT}
      Edit := THDBEdit(GetControl('PSA_TAUXPARTIEL'));
{$ELSE}
      Edit := THEdit(GetControl('PSA_TAUXPARTIEL'));
{$ENDIF}
      if Edit <> nil then
        if Edit.Color <> CouleurHisto then
        if JaiLeDroitTag(200057) then SetControlEnabled('PSA_TAUXPARTIEL', True);  //PT184
    end;
  end;
end;


(*
procedure TOM_Salaries.BasanccpChange(Sender: TObject);
var
  {$IFNDEF EAGLCLIENT}
  Valanccp: THDbEdit;
  CbValanccp: THdbValComboBox;
  {$ELSE}
  Valanccp: THEdit;
  CbValanccp: THValComboBox;
  {$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  Valanccp := THDbEdit(getcontrol('EPSA_VALANCCP'));
  CbValanccp := THdbValComboBox(getcontrol('CPSA_VALANCCP'));
  {$ELSE}
  Valanccp := THEdit(getcontrol('EPSA_VALANCCP'));
  CbValanccp := THValComboBox(getcontrol('CPSA_VALANCCP'));
  {$ENDIF}
  if Valanccp <> nil then
    Valanccp.Text := '';
  if CbValanccp <> nil then
    CbValanccp.value := '';
end;
*)

//*******************PROCEDURE MesErrorSecu***********************//
{Jusqu'à version 536 on effaçait le numero de secu si code clé sexe incorrect
Suite demande CAFSA on laisse le numéro érroné..Suppression du code source
et modification des messages d'erreur SB modif 11/06/01}

procedure TOM_SALARIES.MesErrorSecu(resultat: integer); // PT 58 Message bloquant pour sexe erronné au lieu question
var
  cle2, vide, NumSecuSaisie, CleSexe: string;
begin
  NumSecuSaisie := GetField('PSA_NUMEROSS');
  if (Length(NumSecuSaisie) <> 0) then
    CleSexe := NumSecuSaisie[1];
  vide := '';
  if (resultat <> 0) then // resultat intialisé dans la fonction  TestNumeroSS
  begin
    case resultat of
      1: HShowMessage('1;Salarié:;Vous n''avez pas renseigné le numéro de sécurité sociale.' +
          '#13#10#13#10Pensez-y ulterieurement;E;O;O;O;;;', '', '');

      //PT81 //PT105
      3: HShowMessage('1;Salarié:;Le salarié possède un numéro de sécurité sociale provisoire.' +
          '#13#10#13#10Vous devrez le remplacer par un numéro définitif;E;O;O;O;;;', '', '');
      //FIN PT81

      -2: MesError := 'ne comporte pas de clé.#13#10#13#10Veuillez la saisir!';

      -3: MesError := 'est incomplet.#13#10NB : 15 positions obligatoires'; // PT107

      -7: MesError := 'est incorrect.#13#10#13#10Vous devez saisir une valeur numérique';

      -5:
        begin //incohérence Psa_Sexe=Femme cle sexe=1
          //reponse:=HShowMessage('1;Clé Sexe Erronée :"'+CleSexe+'";Votre clé sexe du numéro de sécurité sociale vaut "2",'+
          //                      '#13#10#13#10Voulez-vous la modifier?;Q;YN;Y;N','','');
          //if reponse = 6 then  // 6 : Oui ; 7 : Non
          //  begin  numss[1]:='2';   SetField('PSA_NUMEROSS',numss); end;
          PGIBox('Clé Sexe Erronée :"' + CleSexe + '";Votre clé sexe du numéro de sécurité sociale vaut "2"', Ecran.Caption);
          numss[1] := '2';
          SetField('PSA_NUMEROSS', numss);
          SetFocusControl('PSA_NUMEROSS');
        end;

      -6:
        begin // incohérence Psa_Sexe=Homme  cle sexe=2
          // reponse:=HShowMessage('1;Clé Sexe Erronée :"'+CleSexe+'";Votre clé sexe du numéro de sécurité sociale vaut "1",#13#10#13#10'+
          //                          'Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
          //  if reponse = 6 then
          //     begin numss[1]:='1'; SetField('PSA_NUMEROSS',numss);  end;
          PGIBox('Clé Sexe Erronée :"' + CleSexe + '";Votre clé sexe du numéro de sécurité sociale vaut "1"', Ecran.Caption);
          numss[1] := '1';
          SetField('PSA_NUMEROSS', numss);
          SetFocusControl('PSA_NUMEROSS');
        end;

      -4:
        begin // clé sexe different de 1 et de 2
          if (GetField('PSA_SEXE') = 'M') then
          begin
            //reponse:=HShowMessage('1;Clé Sexe Erronée :'+CleSexe+';Votre clé sexe vaut 1,#13#10#13#10'+
          //                               'Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
           // if reponse = 6 then
            //  begin numss[1]:='1';  SetField('PSA_NUMEROSS',numss); end;
            PGIBox('Clé Sexe Erronée :' + CleSexe + ';Votre clé sexe vaut 1', Ecran.Caption);
            numss[1] := '1';
            SetField('PSA_NUMEROSS', numss);
            SetFocusControl('PSA_NUMEROSS');
          end;
          if (GetField('PSA_SEXE') = 'F') then
          begin
            // reponse:=HShowMessage('1;Clé Sexe Erronée :'+CleSexe+';Votre clé sexe vaut 2,#13#10Voulez-vous la modifier?;Q;YN;Y;N;;;','','');
            // if reponse = 6 then
            //     begin numss[1]:='2';   SetField('PSA_NUMEROSS',numss); end;
            PGIBox('Clé Sexe Erronée :' + CleSexe + ';Votre clé sexe vaut 2', Ecran.Caption);
            numss[1] := '2';
            SetField('PSA_NUMEROSS', numss);
            SetFocusControl('PSA_NUMEROSS');
          end;
        end;

      -1:
        begin //clé érronnée
          if cle < 10 then
          begin
            cle2 := InttoStr(cle);
            numss[14] := '0';
            Numss[15] := cle2[1];
          end;
          if cle >= 10 then
          begin
            cle2 := InttoStr(cle);
            numss[14] := cle2[1];
            numss[15] := cle2[2];
          end;
          reponse := HShowMessage('1;clé erronée :;Votre clé vaut ' + cle2 + '!! #13#10Voulez-vous la modifier?;Q;YN;Y;N;;;', '', '');
          if reponse = 6 then SetField('PSA_NUMEROSS', numss);
        end;
    end; // end du case
    if (resultat = -2) or (resultat = -3) or (resultat = -7) then
      HShowMessage('3;N° :' + NumSecuSaisie + ' ;Le numéro de sécurité sociale ' + MesError + '' +
        ';E;O;O;O;;;', '', '');


    if (resultat = -7) or (resultat = -3) then
    begin
      numss := '';
      SetField('PSA_NUMEROSS', numss);
      SetfocusControl('PSA_NUMEROSS');
    end;
  end;
end;

// PT22 procédure d'affichage si erreur ds numéro SS pour : année,mois,département de naissance

procedure TOM_SALARIES.MesErrorSecuNaissance(resultAnnee, ResultMois, ResultDepart: integer);
var
  NumSecuSaisie: string;
  Mess: string;
begin
  if (ResultAnnee <> -8) and (ResultMois <> -9) and (ResultDepart <> -10) and (ResultDepart <> -12) then Exit;
  Mess := '';
  NumSecuSaisie := GetField('PSA_NUMEROSS');
  if resultAnnee = -8 then // resultat intialisé dans la fonction  TestNumeroSSNaissance
  begin // Année de naissance non valide
    Mess := Mess + '- Année de naissance erronée : ' + NumSecuSaisie[2] + NumSecuSaisie[3] + '#13#10l''année vaut ' + AnneeSS[3] + AnneeSS[4];
  end;

  if resultMois = -9 then // resultat intialisé dans la fonction  TestNumeroSSNaissance
  begin // Mois de naissance non valide
    Mess := Mess + '#13#10- Mois de naissance erronée : ' + NumSecuSaisie[4] + NumSecuSaisie[5] + '#13#10le mois vaut ' + MoisSS[1] + MoisSS[2];
  end;

  if resultDepart = -10 then // resultat intialisé dans la fonction  TestNumeroSSNaissance
  begin // Département de naissance non valide
// PT157    Mess := Mess + '#13#10- Département de naissance erroné : ' + NumSecuSaisie[6] + NumSecuSaisie[7] + '#13#10le département vaut ' + DepartNaissanceSS[1] + DepartNaissanceSS[2];
    Mess := Mess + '#13#10- Département de naissance erroné : ' + NumSecuSaisie[6] + NumSecuSaisie[7] + '#13#10le département vaut ' + GetField('PSA_DEPTNAISSANCE');

  end;
  //DEBUT PT70
  if resultDepart = -12 then // resultat intialisé dans la fonction  TestNumeroSSNaissance
  begin // Cas salariés corse
    Mess := Mess + '#13#10- La date de naissance du salarié étant inférieure au 01/01/1976, le code du département ne peut être 2A ou 2B';
  end;
  //FIN PT70
  reponse := HShowMessage('1;N° de Sécurité Sociale; ' + Mess + '#13#10#13#10' +
    'Voulez-vous modifier le N° de Sécurité Sociale ?;Q;YN;Y;N;;;', '', '');
  if reponse = 6 then
  begin
    if resultAnnee = -8 then
    begin
      numss[2] := AnneeSS[3];
      numss[3] := AnneeSS[4];
    end;
    if resultMois = -9 then
    begin
      numss[4] := MoisSS[1];
      numss[5] := MoisSS[2];
    end;
    if resultDepart = -10 then
    begin
      DepartNaissanceSS := GetField('PSA_DEPTNAISSANCE'); // PT157
      numss[6] := DepartNaissanceSS[1];
      numss[7] := DepartNaissanceSS[2];
    end;
    if resultDepart = -12 then
    begin
      ;
      numss[6] := '2';
      numss[7] := '0';
      SetField('PSA_DEPTNAISSANCE', '20');
    end; //PT70
    SetField('PSA_NUMEROSS', numss);
  end;
end;




//------------------------------------------------------------------------------
//- détermine, pour l'onglet Congés Payés, la visibilité et les options de saisie pour
//- champs de cet onglet
//------------------------------------------------------------------------------
//procedure TOM_Salaries.OnChangeFieldCongesPayes(F :TField) ;
//begin
//end;

{procedure TOM_Salaries.EnableChampsOngletCongesPayes;
var
 MCombo, Profil: THDBVALComboBox;
begin
   Profil := THDBVALComboBox(Getcontrol('PSA_PROFILCGE'));
// les champs sont non saisissables si idem établissement
   MCombo := THDBVALComboBox(Getcontrol('PSA_CPACQUISMOIS'));
   if ((MCombo <> nil)  and (profil <> nil ))then
     SetControlEnabled('PSA_NBREACQUISCP', ((MCombo.Value <> 'ETB')and (profil.value<>'')));

   MCombo := THDBVALComboBox(Getcontrol('PSA_CPTYPEMETHOD'));
   if ((MCombo <> nil)  and (profil <> nil )) then
     SetControlEnabled('PSA_VALORINDEMCP', ((MCombo.Value <> 'ETB')and (profil.value<>'')));

   MCombo := THDBVALComboBox(Getcontrol('PSA_CPTYPERELIQ'));
   if ((MCombo <> nil)  and (profil <> nil )) then
     SetControlEnabled('PSA_RELIQUAT', ((MCombo.Value <> 'ETB')and (profil.value<>'')));

   MCombo := THDBVALComboBox(Getcontrol('PSA_TYPEPROFILCGE'));
   if ((MCombo <> nil)  and (profil <> nil )) then
     SetControlEnabled('PSA_PROFILCGE', ((MCombo.Value <> 'ETB')and (profil.value<>'')));

   EnabledZonesAnciennete;

end;   }
{
procedure TOM_SALARIES.PagesChange(Sender: TObject);
var
MaPageControl:TPageControl;
begin
MaPageControl := TPageControl(getControl('Pages'));
if MaPageControl <> nil then
 if MaPageControl.ActivePage.Name = 'PCONGES' then
   AlimentecumulConge;
end;                        }

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.AlimentecumulConge;
var
  Acquis, Pris, Restants, Bases, MBase: double;
  MAcquis, MPris, MRestants, MBases: THLabel;
begin
  //SB 28/05/2002 Ajout Variable Idate1900 suite modif function
  AffichelibelleAcqPri('1', Getfield('PSA_SALARIE'), Idate1900, 0, Pris, Acquis, Restants, Bases, MBase, false, False);
  MAcquis := THLabel(Getcontrol('ACQUISN1'));
  if MAcquis <> nil then
    MAcquis.caption := floattostr(arrondi(Acquis, 2));
  MPris := THLabel(Getcontrol('PRISN1'));
  if MPris <> nil then
    MPris.caption := floattostr(arrondi(Pris, 2));
  MRestants := THLabel(Getcontrol('RESTANTSN1'));
  if MRestants <> nil then
    MRestants.caption := floattostr(arrondi(Restants, 2));
  MBases := THLabel(Getcontrol('BASESN1'));
  if MBases <> nil then
    MBases.caption := floattostr(arrondi(Bases, 2));
  //SB 28/05/2002 Ajout Variable Idate1900 suite modif function
  AffichelibelleAcqPri('0', Getfield('PSA_SALARIE'), Idate1900, 0, Pris, Acquis, Restants, Bases, MBase, false, False);
  MAcquis := THLabel(Getcontrol('ACQUISN'));
  if MAcquis <> nil then
    MAcquis.caption := floattostr(arrondi(Acquis, 2));
  MPris := THLabel(Getcontrol('PRISN'));
  if MPris <> nil then
    MPris.caption := floattostr(arrondi(Pris, 2));
  MRestants := THLabel(Getcontrol('RESTANTSN'));
  if MRestants <> nil then
    MRestants.caption := floattostr(arrondi(Restants, 2));
  MBases := THLabel(Getcontrol('BASESN'));
  if MBases <> nil then
    MBases.caption := floattostr(arrondi(Bases, 2));

end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
//*******************PROCEDURE On New Record***************************//

procedure TOM_SALARIES.OnNewRecord;
var
  Q: TQuery;
{$IFNDEF EAGLCLIENT}
  Edit: ThDBEdit;
{$ELSE}
  Edit: ThEdit;
{$ENDIF}
  okok: boolean;
  St: string;
  // PT123
  Etab: string;

begin
  inherited;
{$IFNDEF GCGC}
  if GetParamSocSecur('SO_PGHISTOAVANCE', False) = True then RendAccesChampHisto(True, False);
{$ENDIF GCGC}
  Creation := (DS.State in [dsInsert]);
// DEB 137 Positin de l'établissement en debut de fonction pour initialisation correcte des champs
  // DEB PT109
  st := PGRendEtabUser();
  PGPositionneEtabUser(THValComboBox(PGGetControl('PSA_ETABLISSEMENT', Ecran)));
  if St <> '' then
    SetField('PSA_ETABLISSEMENT', St);
// FIN PT109

// DEB PT123
  if St = '' then
  begin
    Etab := Getparamsoc('SO_ETABLISDEFAUT');
    if Etab <> '' then
      SetField('PSA_ETABLISSEMENT', Etab);
  end;
// FIN PT123
// FIN PT137
  if Argument = 'INT' then
  begin
    SetField('PSA_SALARIE', Interimaire);
    setControlEnabled('PSA_SALARIE', False);
    Q := OpenSQL('SELECT PSI_INTERIMAIRE,PSI_LIBELLE,PSI_PRENOM,PSI_ADRESSE1,' +
      ' PSI_ADRESSE2,PSI_ADRESSE3,PSI_DATENAISSANCE,PSI_CODEPOSTAL,' +
      ' PSI_VILLE,PSI_PAYS,PSI_NATIONALITE,PSI_SEXE,' +
      ' PSI_SITUATIONFAMIL,PSI_CARTESEJOUR,PSI_DATEXPIRSEJOUR,' +
      ' PSI_DELIVPAR,PSI_CONFIDENTIEL,PSI_NUMEROSS,PSI_TELEPHONE,' +
      ' PSI_PORTABLE,PEI_ETABLISSEMENT,PEI_TRAVAILN1,PEI_TRAVAILN2,' +
      ' PEI_TRAVAILN3,PEI_TRAVAILN4,PEI_CODESTAT,PEI_COEFFICIENT,' +
      ' PEI_QUALIFICATION,PEI_LIBELLEQUALIF,PEI_CODEEMPLOI,' +
      ' PEI_LIBELLEEMPLOI,PEI_LIBREPCMB1,PEI_LIBREPCMB2,' +
      ' PEI_LIBREPCMB3,PEI_LIBREPCMB4,PSI_NOMJF' +
      ' FROM INTERIMAIRES' +
      ' LEFT JOIN EMPLOIINTERIM ON' +
      ' PEI_INTERIMAIRE=PSI_INTERIMAIRE WHERE' +
      ' PSI_INTERIMAIRE="' + Interimaire + '"' +
      ' ORDER BY PEI_ORDRE DESC', true);
    // portageCWAS
    if not Q.eof then
    begin
      SetField('PSA_LIBELLE', Q.FindField('PSI_LIBELLE').AsString);
      SetField('PSA_PRENOM', Q.FindField('PSI_PRENOM').AsString);
      SetField('PSA_NOMJF', Q.FindField('PSI_NOMJF').AsString);
      SetField('PSA_ADRESSE1', Q.FindField('PSI_ADRESSE1').AsString);
      SetField('PSA_ADRESSE2', Q.FindField('PSI_ADRESSE2').AsString);
      SetField('PSA_ADRESSE3', Q.FindField('PSI_ADRESSE3').AsString);
      SetField('PSA_DATENAISSANCE', Q.FindField('PSI_DATENAISSANCE').AsDateTime);
      SetField('PSA_CODEPOSTAL', Q.FindField('PSI_CODEPOSTAL').AsString);
      SetField('PSA_VILLE', Q.FindField('PSI_VILLE').AsString);
      SetField('PSA_PAYS', Q.FindField('PSI_PAYS').AsString);
      SetField('PSA_NATIONALITE', Q.FindField('PSI_NATIONALITE').AsString);
      SetField('PSA_SEXE', Q.FindField('PSI_SEXE').AsString);
      if Q.FindField('PSI_SEXE').AsString = 'M' then
        SetField('PSA_CIVILITE', 'MR');
      SetField('PSA_SITUATIONFAMIL', Q.FindField('PSI_SITUATIONFAMIL').AsString);
      SetField('PSA_CARTESEJOUR', Q.FindField('PSI_CARTESEJOUR').AsString);
      SetField('PSA_DATEXPIRSEJOUR', Q.FindField('PSI_DATEXPIRSEJOUR').AsDateTime);
      SetField('PSA_DELIVPAR', Q.FindField('PSI_DELIVPAR').AsString);
      if Q.FindField('PSI_CONFIDENTIEL').AsString = 'X' then
        SetField('PSA_CONFIDENTIEL', '1')
      else
        SetField('PSA_CONFIDENTIEL', '0');
      SetField('PSA_NUMEROSS', Q.FindField('PSI_NUMEROSS').AsString);
      SetField('PSA_TELEPHONE', Q.FindField('PSI_TELEPHONE').AsString);
      SetField('PSA_PORTABLE', Q.FindField('PSI_PORTABLE').AsString);
      SetField('PSA_ETABLISSEMENT', Q.FindField('PEI_ETABLISSEMENT').AsString);
      SetField('PSA_TRAVAILN1', Q.FindField('PEI_TRAVAILN1').AsString);
      SetField('PSA_TRAVAILN2', Q.FindField('PEI_TRAVAILN2').AsString);
      SetField('PSA_TRAVAILN3', Q.FindField('PEI_TRAVAILN3').AsString);
      SetField('PSA_TRAVAILN4', Q.FindField('PEI_TRAVAILN4').AsString);
      SetField('PSA_CODESTAT', Q.FindField('PEI_CODESTAT').AsString);
      SetField('PSA_COEFFICIENT', Q.FindField('PEI_COEFFICIENT').AsString);
      SetField('PSA_QUALIFICATION', Q.FindField('PEI_QUALIFICATION').AsString);
      SetField('PSA_LIBELLEQUALIF', Q.FindField('PEI_LIBELLEQUALIF').AsString);
      SetField('PSA_CODEEMPLOI', Q.FindField('PEI_CODEEMPLOI').AsString);
      SetField('PSA_LIBELLEEMPLOI', Q.FindField('PEI_LIBELLEEMPLOI').AsString);
      SetField('PSA_LIBREPCMB1', Q.FindField('PEI_LIBREPCMB1').AsString);
      SetField('PSA_LIBREPCMB2', Q.FindField('PEI_LIBREPCMB2').AsString);
      SetField('PSA_LIBREPCMB3', Q.FindField('PEI_LIBREPCMB3').AsString);
      SetField('PSA_LIBREPCMB4', Q.FindField('PEI_LIBREPCMB4').AsString);
    end
    else
      PGIBox('Les données concernant l''intérimaire n''existent pas', Ecran.Caption);
    Ferme(Q);
  end;
// DEB PT137
  St := RechDom('PGETABCONGES', GetField('PSA_ETABLISSEMENT'), False);
  EtbCP := ((St <> '') and (St <> 'Error'));
  if (VH_Paie.PgCongesPayes) and (EtbCP) then SetField('PSA_CONGESPAYES', 'X');
// FIN PT137
  okok := False;
  SetControlEnabled('TCALENDRIER', OkOk);
  SetControlEnabled('BCP', OkOk);
  SetControlEnabled('BCOMPETENCE', OkOk);
  SetControlEnabled('BBANQUE', OkOk);
  SetControlEnabled('BMEMO', OkOk);
  SetControlEnabled('BENFANT', OkOk);
  SetControlEnabled('BMULTIEMPLOY', OkOk);
  SetControlVisible('RIB__BMULTIEMP', OkOk);
  SetControlEnabled('BCONTRAT', OkOk);
  SetControlEnabled('BTNVENTIL', OkOk);
  SetControlEnabled('BDUPLIQUER', OkOk); { PT133 }

// Champs idem etab
  SetField('PSA_TYPPROFIL', 'ETB');
  SetField('PSA_TYPPROFILREM', 'ETB');
  SetField('PSA_TYPPERIODEBUL', 'ETB');
  SetField('PSA_TYPPROFILRBS', 'ETB');
  SetField('PSA_TYPPROFILAFP', 'ETB');
  SetField('PSA_TYPPROFILAPP', 'ETB');
  SetField('PSA_TYPPROFILRET', 'ETB');
  SetField('PSA_TYPPROFILMUT', 'ETB');
  SetField('PSA_TYPPROFILPRE', 'ETB');
  SetField('PSA_TYPPROFILTSS', 'ETB');
  SetField('PSA_TYPPROFILCGE', 'ETB');
  SetField('PSA_TYPPROFILANC', 'ETB');
  SetField('PSA_TYPEDITBULCP', 'ETB');
  SetField('PSA_CPACQUISMOIS', 'ETB');
  SetField('PSA_CPACQUISSUPP', 'ETB');
  SetField('PSA_TYPNBACQUISCP', 'ETB');
  SetField('PSA_CPTYPEMETHOD', 'ETB');
  SetField('PSA_CPTYPERELIQ', 'ETB');
  SetField('PSA_CPACQUISANC', 'ETB');
  SetField('PSA_DATANC', 'ETB');
  SetField('PSA_TYPPROFILTRANS', 'ETB');
  SetField('PSA_TYPPROFILFNAL', 'DOS');
  SetField('PSA_STANDCALEND', 'ETB');
  SetField('PSA_TYPREDREPAS', 'ETB');
  SetField('PSA_TYPREDRTT1', 'ETB');
  SetField('PSA_TYPREDRTT2', 'ETB');
  SetField('PSA_TYPDADSFRAC', 'ETB');
  SetField('PSA_TYPPRUDH', 'ETB');
  SetField('PSA_CPTYPEVALO', 'ETB');
  SetField('PSA_TYPPAIEVALOMS', 'ETB'); { PT132 }
//Fiche Banque
  SetField('PSA_TYPREGLT', 'ETB');
  SetField('PSA_TYPVIRSOC', 'ETB');
  SetField('PSA_TYPDATPAIEMENT', 'ETB');
  SetField('PSA_TYPPAIACOMPT', 'ETB');
  SetField('PSA_TYPACPSOC', 'ETB');
  SetField('PSA_TYPPAIFRAIS', 'ETB');
  SetField('PSA_TYPFRAISSOC', 'ETB');
  SetField('PSA_TYPJOURHEURE', 'ETB');
  SetField('PSA_TYPEDITORG', 'ETB');
  SetField('PSA_TYPACTIVITE', 'ETB');

//Date d'entree
  SetField('PSA_DATEENTREE', Date);
  SetField('PSA_DATEANCIENNETE', Getfield('PSA_DATEENTREE'));
  SetField('PSA_ANCIENPOSTE', GetField('PSA_DATEENTREE'));
  SetField('PSA_DATENAISSANCE', Idate1900);
  SetField('PSA_DATESORTIE', Idate1900);
  SetField('PSA_DATEENTREEPREC', Idate1900);
  SetField('PSA_DATESORTIEPREC', Idate1900);
  SetField('PSA_DATELIBRE1', Idate1900);
  SetField('PSA_DATELIBRE2', Idate1900);
  SetField('PSA_DATELIBRE3', Idate1900);
  SetField('PSA_DATELIBRE4', Idate1900);

// Ordre AT
  SetField('PSA_ORDREAT', '1');

//DADS-U
  SetField('PSA_DADSFRACTION', '1');
  SetField('PSA_REGIMESS', '200');
  //DEB PT164
  SetControlEnabled('PSA_REGIMESS', True);
  setfield('PSA_TYPEREGIME', '-');
  SetControlEnabled('PSA_REGIMEMAL', False);
  SetControlEnabled('PSA_REGIMEAT', False);
  SetControlEnabled('PSA_REGIMEVIP', False);
  SetControlEnabled('PSA_REGIMEVIS', False);
  SetField('PSA_REGIMEMAL', '');
  SetField('PSA_REGIMEAT', '');
  SetField('PSA_REGIMEVIP', '');
  SetField('PSA_REGIMEVIS', '');
  //FIN PT164
  SetField('PSA_TAUXPARTSS', 0);
  SetField('PSA_TAUXPARTIEL', 0);
  SetField('PSA_DADSDATE', IDate1900);

  SetField('PSA_PERSACHARGE', 99); //PT113

  if Argument <> 'INT' then
  begin
{$IFNDEF EAGLCLIENT}
    Edit := THDBEdit(GetControl('PSA_SALARIE'));
{$ELSE}
    Edit := THEdit(GetControl('PSA_SALARIE'));
{$ENDIF}
    if (VH_PAIE.PgTypeNumSal = 'ALP') then
      if Edit <> nil then
        Edit.EditMask := '';
    if (VH_PAIE.PgTypeNumSal = 'NUM') then
    begin
      if Edit <> nil then
        Edit.EditMask := '9999999999';
      if (VH_Paie.PGIncSalarie = TRUE) then
        AffectCodeNewEnr;
    end;
  end;

//Effectif
  SetField('PSA_PRISEFFECTIF', 'X');
  SetField('PSA_UNITEPRISEFF', 1);

  Q := OpenSQL('SELECT ETB_CONVENTION' +
    ' FROM ETABCOMPL WHERE' +
    ' ETB_ETABLISSEMENT = "' + GetField('PSA_ETABLISSEMENT') + '"', TRUE);
  if not Q.Eof then
  begin
    SetField('PSA_TYPCONVENTION', 'ETB');
    SetField('PSA_CONVENTION', Q.FindField('ETB_CONVENTION').AsString);
  end
  else
  begin
    SetField('PSA_TYPCONVENTION', 'PER');
    SetField('PSA_CONVENTION', '');
  end;
  Ferme(Q);
  SetField('PSA_NATIONALITE', 'FRA');
  SetField('PSA_PAYSNAISSANCE', 'FRA');
  SetField('PSA_CONFIDENTIEL', '0'); //PT103 Gestion de la confidentialité en CWAS
  if SalarieDupliquer <> '' then
  begin
    DuplicationSalarie(SalarieDupliquer); // PT104
    TPageControl(GetControl('Pages')).ActivePage := TPageControl(GetControl('Pages')).Pages[0];
    TFFiche(Ecran).FTypeAction := TaCreat;
  end;
//  SalarieDupliquer := ''; PT121 Fais dans le OnLoadRecord
{ PT137 mis en commentaire car remonté au début de la fonction
// DEB PT109
  st := PGRendEtabUser();
  PGPositionneEtabUser(THValComboBox(PGGetControl('PSA_ETABLISSEMENT', Ecran)));
  if St <> '' then
    SetField('PSA_ETABLISSEMENT', St);
// FIN PT109

// DEB PT123
  if St = '' then
  begin
    Etab := Getparamsoc('SO_ETABLISDEFAUT');
    if Etab <> '' then
      SetField('PSA_ETABLISSEMENT', Etab);
  end;
// FIN PT123
}
if (SalarieDupliquer='') then   //PT181
   SetField ('PSA_CATBILAN', '000');
   ClickBtnBulDefaut (nil); { PT128 }
end;

//*******************PROCEDURE On Delete Record***********************//

procedure TOM_SALARIES.OnDeleteRecord;
var
  ExisteCod: Boolean;
  NomChamp: array[1..1] of Hstring;
  ValChamp: array[1..1] of variant;
  i, rep: integer;
  Q: TQuery;
begin
  inherited;
  if GetField('PSA_NUMEROBL') <> '' then
  begin
    rep := PgiAsk('Attention, ce salarié provient de l''importation d''un' +
      ' fichier EWS, voulez-vous le supprimer ?', Ecran.caption);
    if rep <> mrYes then
      LastError := 2
    else
      LastError := 0;
  end
  else
    LastError := 0;

  if LastError = 0 then
  begin
    i := 1;
    NomChamp[1] := 'PPU_SALARIE';
    ValChamp[1] := GetField('PSA_SALARIE');
    ExisteCod := RechEnrAssocier('PAIEENCOURS', NomChamp, ValChamp);
    if ExisteCod = TRUE then
    begin
      LastError := 1;
      PgiBox('Suppression impossible. Des bulletins de paie ont été générés' +
        ' pour ce salarié.', Ecran.caption);
    end
    else
// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
    begin
{$IFDEF GCGC}
{Vérification des Circuits de validation et message si salarié
est présent en tant que valideur dans un circuit de validation.}
      ExisteCod := RechSalarieModuleDA(GetField('PSA_SALARIE'));
      if ExisteCod = TRUE then
        LastError := 1;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
      if ExisteCod = False then
      begin
        InitMoveProgressForm(nil, 'Suppression des données',
          'Veuillez patienter SVP ...', i, FALSE, TRUE);
{ Inutile de recherche si l'enregistrement existe pour le supprimer
 Suppression ligne de code }
        try
          BeginTrans;
          MoveCurProgressForm('Suppression des données : profils spéciaux..');
          ExecuteSQL('DELETE FROM PROFILSPECIAUX WHERE' +
            ' PPS_CODE="' + ValChamp[1] + '" AND' +
            ' PPS_ETABSALARIE="-"');
          MoveCurProgressForm('Suppression des données : compétences..');
//Suppression interimaire et compétences liés au salarie
          ExecuteSQL('DELETE FROM RHCOMPETRESSOURCE WHERE' +
            ' PCH_SALARIE="' + ValChamp[1] + '" AND' +
            ' PCH_TYPERESSOURCE="SAL"');
          ExecuteSQL('DELETE FROM INTERIMAIRES WHERE' +
            ' PSI_INTERIMAIRE="' + ValChamp[1] + '" AND' +
            ' PSI_TYPEINTERIM="SAL"');
          MoveCurProgressForm('Suppression des données : enfants..');
          ExecuteSQL('DELETE FROM ENFANTSALARIE WHERE' +
            ' PEF_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : multiemployeur..');
          ExecuteSQL('DELETE FROM MULTIEMPLOYSAL WHERE' +
            ' PML_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : absences & congés payés..');
          ExecuteSQL('DELETE FROM ABSENCESALARIE WHERE' +
            ' PCN_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : attestations..');
          ExecuteSQL('DELETE FROM ATTESTATIONS WHERE' +
            ' PAS_SALARIE="' + ValChamp[1] + '"');
          ExecuteSQL('DELETE FROM ATTSALAIRES WHERE' +
            ' PAL_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : contrat de travail..');
          ExecuteSQL('DELETE FROM CONTRATTRAVAIL WHERE' +
            ' PCI_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : calendrier..');
          ExecuteSQL('DELETE FROM CALENDRIER WHERE' +
            ' ACA_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : appoints..');
          ExecuteSQL('DELETE FROM APPOINTSALARIE WHERE' +
            ' PAO_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : historique des bulletins..');
          ExecuteSQL('DELETE FROM HISTOBULLETIN WHERE' +
            ' PHB_SALARIE="' + ValChamp[1] + '"');
          ExecuteSQL('DELETE FROM HISTOCUMSAL WHERE' +
            ' PHC_SALARIE="' + ValChamp[1] + '"');
          ExecuteSQL('DELETE FROM HISTOSAISRUB WHERE' +
            ' PSD_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Suppression des données : historique du salarié..');
          ExecuteSQL('DELETE FROM HISTOSALARIE WHERE' +
            ' PHS_SALARIE="' + ValChamp[1] + '"');
          ExecuteSQL('DELETE FROM PGHISTODETAIL WHERE' +
            ' PHD_SALARIE="' + ValChamp[1] + '"');
          MoveCurProgressForm('Effacement des données : ressource..');
          Q := OpenSql('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE' +
            ' ARS_SALARIE="' + ValChamp[1] + '"', True);
          if not Q.eof then
            PGSupprRessource(Q.FindField('ARS_RESSOURCE').AsString);
          Ferme(Q);
          DeleteYRS(ValChamp[1], '', ''); //PT155-1
//PT162-5
          MoveCurProgressForm('Suppression des données : DADS-U...');
          ExecuteSQL('DELETE FROM DADSPERIODES WHERE' +
            ' PDE_SALARIE="' + ValChamp[1] + '"');
          ExecuteSQL('DELETE FROM DADSDETAIL WHERE' +
            ' PDS_SALARIE="' + ValChamp[1] + '"');
//FIN PT162-5
          CommitTrans;
        except
          Rollback;
          LastError := 1;
          PgiBox('Une erreur est survenue lors de la suppression des' +
            ' données associées au salarié.', Ecran.caption);
        end;
        FiniMoveProgressForm;
{$IFNDEF EAGLCLIENT}
        ChargementTablette(TFFiche(Ecran).TableName, '');
{$ELSE}
        ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');
{$ENDIF}
      end;
{$IFDEF PAIEGRH}
      //DEB PT185
      Trace := TStringList.Create ;
      Trace.Add('SUPPRESSION SALARIE '+GetField('PSA_SALARIE')+' '+ GetField('PSA_LIBELLE'));
      CreeJnalEvt('003','097','OK',nil,nil,Trace);
      FreeAndNil (Trace);
      //FIN PT185
{$ENDIF PAIEGRH}
    end; // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
  end;
end;

//*******************FONCTIONS TestNumSecu****************************//

//fonction testant le num de sécu
{PT51
Function TOM_SALARIES.TestNumeroSS( numss: String ; codesexe : String ) : integer;
var
  i,resultat,k,temp,code:integer;
  numss13,numsscle:string;
  entier : Extended;
  DblCle : Double;
  numcle :integer;
  num13 : Extended;
  // num13 : int64;
begin
 k:= length(numss);
 resultat:=0;
 if ((numss<>'') or (numss='000000000000000'))  then
    begin
        if (resultat=0) and (k<15)  then resultat:=-3;
        if (resultat=-3) and (k=13) then  resultat:=-2;
        if (k=15) and (resultat=0) then          //contrôle du code sexe
          begin
            if ((numss[1]<'1') or (numss[1]>'2'))then resultat:=-4; //incohérence entre code sexe saisie et code sexe conforme
            if numss[1]='1' then     //psa_sexe = homme
                begin
                    if (codesexe='F') then resultat:=-5; //Incohérence entre psa_sexe et code sexe
                end
            else
                begin
                   if numss[1]='2' then         //psa_sexe = femme
                      begin
                          if (codesexe='M') then resultat:=-6; //Incohérence entre psa_sexe et code sexe
                      end;
                end;
          end;
         if (k=15) and (resultat=0) and (numss[6]='2') then  //contrôle du code département Corse
           begin
            if (numss[7]='A') or (numss[7]='B') then resultat:=2;
           end;
        for i:=1 to k do      //contrôle de numéricité du n° de sécu
          begin
            if resultat<>2 then
            begin
                 if numss[i] in ['0'..'9']  then else resultat:=-7;  //non numérique
            end;
          end;
       if (resultat=0) and (k=15) then      //calcul de la clé
          begin
            //Récupération 13 premier char et cle saisie
            numss13:=Copy(numss,1,13);
            if numss13<>'' then
              Begin
              Num13:=valeur(Numss13);
              numsscle:=numss[14]+numss[15];
              val(numsscle,temp,code);
              numcle:=temp;
              entier:=Int(Num13/97);
              DblCle:= 97-(Num13-(entier*97));
              Cle:=StrToInt(FloatToStr(DblCle));
              if (cle<>numcle)  and (numcle>0) then resultat:=-1
              End;
          end;
    end
  else
   if numss='' then resultat:=1;

result:=resultat;
end;
}

// PT22 fonction de contrôle ds numéro SS pour : année,mois et département de naissance
{PT155-2
function TOM_SALARIES.TestNumeroSSNaissance(numss: string; SSNaiss: string): integer;
var
  aa, mm, jj: word;
  DateNaissanceSS: TDateTime;
  TestMoisSS, TestNumDepartSS: string;
begin
  result := 0;
  //DEBUT PT70
  if (SSNaiss = 'Mois') or (SSNaiss = 'Annee') then
  begin
    TestMoisSS := Copy(numss, 4, 2);
    if IsNumeric(TestMoisSS) then
    begin
      if (StrToInt(TestMoisSS) >= 20) and (StrToInt(TestMoisSS) <= 42) then
      begin
        resultMois := 0;
        ResultAnnee := 0;
        exit;
      end;
      if StrToInt(TestMoisSS) >= 50 then
      begin
        resultMois := 0;
        ResultAnnee := 0;
        exit;
      end;
    end;
  end;
  //FIN PT70
  if SSNaiss = 'Mois' then
  begin
    resultMois := 0;
    if (numss <> '') or (numss = '000000000000000') then
    begin
      if (GetField('PSA_DATENAISSANCE') <> NULL) then
      begin
        DateNaissanceSS := GetField('PSA_DATENAISSANCE');
        if DateNaissanceSS <> idate1900 then
        begin
          DecodeDate(DateNaissanceSS, aa, mm, jj);
          if mm < 10 then MoisSS := '0' + IntToStr(mm)
          else MoisSS := IntToStr(mm);
          if (NumSS[4] <> MoisSS[1]) or (NumSS[5] <> MoisSS[2]) then ResultMois := -9;
        end;
      end;
    end;
    result := ResultMois;
  end;
  if SSNaiss = 'Annee' then
  begin
    ResultAnnee := 0;
    if ((numss <> '') or (numss = '000000000000000')) then
    begin
      if (GetField('PSA_DATENAISSANCE') <> NULL) then
      begin
        DateNaissanceSS := GetField('PSA_DATENAISSANCE');
        if DateNaissanceSS <> idate1900 then
        begin
          DecodeDate(DateNaissanceSS, aa, mm, jj);
          AnneeSS := IntToStr(aa);
          if (NumSS[2] <> AnneeSS[3]) or (NumSS[3] <> AnneeSS[4]) then ResultAnnee := -8;
        end;
      end;
    end;
    result := resultAnnee;
  end;
  if SSNaiss = 'Depart' then
  begin
    resultDepart := 0;
    //PT40 28/05/2002 V582 PH Controle département de naissance 75 du numss <> departement de naissance
    //                        si Annee de naissance >= 1964 et <= 1968 alors Ok
    if ((numss <> '') or (numss = '000000000000000')) and (GetField('PSA_DATENAISSANCE') <> NULL)
      and (GetField('PSA_DEPTNAISSANCE') <> NULL) then
    begin
      //DEBUT PT70
      TestNumDepartSS := Copy(numss, 6, 2);
      if (TestNumDepartSS = '  ') or (Copy(TestNumDepartSS, 1, 1) = ' ') or (Copy(TestNumDepartSS, 2, 1) = ' ') then Exit; //PT101-1
      if IsNumeric(TestNumDepartSS) then
      begin
        if (StrToInt(TestNumDepartSS) >= 91) and (StrToInt(TestNumDepartSS) <= 96) and (GetField('PSA_DEPTNAISSANCE') = '99') then exit;
      end;
      if IsNumeric(GetField('PSA_DEPTNAISSANCE')) then
        if (StrToInt(GetField('PSA_DEPTNAISSANCE')) >= 91) and (StrToInt(GetField('PSA_DEPTNAISSANCE')) >= 95) then exit;
      //FIN PT70
      DepartNaissanceSS := GetField('PSA_DEPTNAISSANCE');
      DateNaissanceSS := GetField('PSA_DATENAISSANCE');
      aa := 0;
      if DateNaissanceSS <> idate1900 then
        DecodeDate(DateNaissanceSS, aa, mm, jj);
      if Length(DepartNaissanceSS) = 2 then
      begin
        if (NumSS[6] <> DepartNaissanceSS[1]) or (NumSS[7] <> DepartNaissanceSS[2]) then ResultDepart := -10;
        if (Copy(NumSS, 6, 2) = '75') and (Copy(DepartNaissanceSS, 1, 2) <> Copy(NumSS, 6, 2))
          and (aa <= 1968) then ResultDepart := 0;
        //DEBUT PT70
        if (DateNaissanceSS >= StrToDate('01/01/1976')) then //salariés nés en corse
        begin
          if ((Copy(NumSS, 6, 2) = '2A') or (Copy(NumSS, 6, 2) = '2B')) and (DepartNaissanceSS = '20') then ResultDepart := 0;
        end
        else
        begin
          if (DateNaissanceSS <> IDate1900) and ((Copy(NumSS, 6, 2) = '2A') or (Copy(NumSS, 6, 2) = '2B')) then ResultDepart := -12;
          if (Copy(NumSS, 6, 2) = '20') and ((DepartNaissanceSS = '2A') or (DepartNaissanceSS = '2B')) then ResultDepart := 0;
        end;
        //FIN PT70
// FIN PT40
      end;
    end;
    result := resultDepart;
  end;
end;
}

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.TCalendrierClick(Sender: TObject);
var
  Init: word;
begin
  if JaiLeDroitTag(200061) then //PT184
  begin
    Init := HShowMessage('1;Calendrier salarié;Attention, avant d''aller sur le calendrier, vous devez enregistrer vos modifications.#13#10Voulez-vous continuer ?;Q;YN;N;N;', '','');
    if Init = mrNo then
      exit
    else
      //     OnUpdaterecord;
      TFFiche(Ecran).BValiderClick(Sender);
  end;
  if ((Getfield('PSA_STANDCALEND') = 'ETB') or (Getfield('PSA_STANDCALEND') = 'ETS')) then
    if (not JaiLeDroitTag(200061))  then   //PT184
      AgllanceFiche('AFF', 'HORAIRESTD', '', '',
        'TYPE:STD;CODE:' + Getfield('PSA_SALARIE') + ';LIBELLE:' + Getfield('PSA_LIBELLE') + ';STANDARD:' + Getfield('PSA_CALENDRIER') + ';ACTION:CONSULTATION;APPEL:SALARIES')
    else
      AgllanceFiche('AFF', 'HORAIRESTD', '', '',
        'TYPE:STD;CODE:' + Getfield('PSA_SALARIE') + ';LIBELLE:' + Getfield('PSA_LIBELLE') + ';STANDARD:' + Getfield('PSA_CALENDRIER') + ';ACTION:CONSULTATION');
  if (Getfield('PSA_STANDCALEND') = 'PER') then
    if (not JaiLeDroitTag(200061))  then  //PT184
      AgllanceFiche('AFF', 'HORAIRESTD', '', '',
        'TYPE:SAL;CODE:' + Getfield('PSA_SALARIE') + ';LIBELLE:' + Getfield('PSA_LIBELLE') + ';STANDARD:' + Getfield('PSA_CALENDRIER') + ';ACTION:CONSULTATION;APPEL:SALARIES')
    else
      AgllanceFiche('AFF', 'HORAIRESTD', '', '',
        'TYPE:SAL;CODE:' + Getfield('PSA_SALARIE') + ';LIBELLE:' + Getfield('PSA_LIBELLE') + ';STANDARD:' + Getfield('PSA_CALENDRIER') + ';ACTION:MODIFICATION');
end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.BTCongesPayesClick(Sender: TObject);
var
  Pris, Acquis, Base, Restants, MBase: double;
begin
  //  AglLanceFiche ('PAY','CONGESPAY_MUL', 'PCN_SALARIE='+getfield('PSA_SALARIE'), 'PMI_TYPENATURE=VAL' , '' );
   // AglLanceFiche ('PAY','CONGESPAYES', 'PCN_SALARIE='+getfield('PSA_SALARIE'), '' , getfield('PSA_SALARIE') );
  FirstCp := false;
  {DEB PT- 1}
    //  AglLanceFiche ('PAY','CONGESPAY_MUL', 'PCN_PERIODECP=0;PCN_TYPECONGE= ', getfield('PSA_SALARIE') , getfield('PSA_SALARIE') +';C;'+getfield('PSA_ETABLISSEMENT')+';;SAL');
  if not JaiLeDroitTag(200067) then //PT184
    AglLanceFiche('PAY', 'CONGESPAY_MUL', 'PCN_PERIODECP=0;PCN_TYPECONGE= ', getfield('PSA_SALARIE'), getfield('PSA_SALARIE') + ';C;;;SAL;ACTION=CONSULTATION')
  else
    AglLanceFiche('PAY', 'CONGESPAY_MUL', 'PCN_PERIODECP=0;PCN_TYPECONGE= ', getfield('PSA_SALARIE'), getfield('PSA_SALARIE') + ';C;;;SAL');
  {FIN PT- 1}
   //SB 28/05/2002 Ajout Variable Idate1900 suite modif function
  AffichelibelleAcqPri('0', getfield('PSA_SALARIE'), Idate1900, 0, Pris, Acquis, Restants, Base, MBase, false, False);
end;
{$ENDIF}
// fin CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

procedure TOM_SALARIES.FRibDblClick(Sender: TObject);
begin
{$IFNDEF GCGC}
  if JaiLeDroitTag(200063) then //PT184
  begin
{$ENDIF}
    FicheRIB_AGL(GetField('PSA_AUXILIAIRE'), taModif, FALSE, '', TRUE);
    ChargeGrilleRib(THGrid(GetControl('FRIB')), '');
    ChargeGrilleRib(THGrid(GetControl('FRIBACPTE')), 'ACP');
    ChargeGrilleRib(THGrid(GetControl('FRIBFP')), 'FPR');
{$IFNDEF GCGC}
  end;
{$ENDIF}
end;

procedure TOM_SALARIES.AffichePhoto;
var
  QQ: TQuery;
  SQL: string;
  CC: TImage;
begin
  CC := TImage(GetControl('LAPHOTO'));
  if CC = nil then exit;
  SQL := 'SELECT * from LIENSOLE where LO_TABLEBLOB="PG" AND LO_IDENTIFIANT="' + getField('PSA_SALARIE') + '"';
  SQL := SQL + ' AND LO_QUALIFIANTBLOB="PHO" AND LO_EMPLOIBLOB="' + VH_Paie.PGPhotoSal + '"';
  QQ := OpenSQL(SQL, true);
  if not QQ.EOF then
  begin
    LoadBitMapFromChamp(QQ, 'LO_OBJET', CC);
  end;
  ferme(QQ);
end;


procedure TOM_SALARIES.RendVisibleOrg;
var
  TLieu, TCodeStat: THLabel;
{$IFNDEF EAGLCLIENT}
  Combo, ComboLieu: THDBValComboBox;
  Coche: TDBCheckBox;
  Date, Mensuel, Annuel: THDBEdit;
{$ELSE}
  Combo, ComboLieu: THValComboBox;
  Coche: TCheckBox;
  Date, Mensuel, Annuel: THEdit;
{$ENDIF}
  Cadre: TBevel;
  num: integer;
  numero, libelle: string;
begin
  //Visiblilité Champ Organisation
  for num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PGNbreStatOrg then break;
    libelle := '';
    if Num = 1 then libelle := VH_Paie.PGLibelleOrgStat1;
    if Num = 2 then libelle := VH_Paie.PGLibelleOrgStat2;
    if Num = 3 then libelle := VH_Paie.PGLibelleOrgStat3;
    if Num = 4 then libelle := VH_Paie.PGLibelleOrgStat4;
    if (Libelle <> '') then
    begin
{$IFNDEF EAGLCLIENT}
      ComboLieu := THDBValComboBox(GetControl('PSA_TRAVAILN' + Numero));
{$ELSE}
      ComboLieu := THValComboBox(GetControl('PSA_TRAVAILN' + Numero));
{$ENDIF}
      if (ComboLieu <> nil) then ComboLieu.Visible := TRUE;
      TLieu := THLabel(GetControl('TPSA_TRAVAILN' + Numero));
      if (TLieu <> nil) then
      begin
        TLieu.Visible := TRUE;
        if Num = 1 then
        begin
          TLieu.Caption := VH_Paie.PGLibelleOrgStat1;
        end;
        if Num = 2 then
        begin
          TLieu.Caption := VH_Paie.PGLibelleOrgStat2;
        end;
        if Num = 3 then
        begin
          TLieu.Caption := VH_Paie.PGLibelleOrgStat3;
        end;
        if Num = 4 then
        begin
          TLieu.Caption := VH_Paie.PGLibelleOrgStat4;
        end;
      end;
    end; // begin libelle
  end;
  //Visibilité du code statistique
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    TCodeStat := THLabel(GetControl('TPSA_CODESTAT'));
{$IFNDEF EAGLCLIENT}
    ComboLieu := THDBValComboBox(GetControl('PSA_CODESTAT'));
{$ELSE}
    ComboLieu := THValComboBox(GetControl('PSA_CODESTAT'));
{$ENDIF}
    if TCodeStat <> nil then
    begin
      TCodeStat.Caption := VH_Paie.PGLibCodeStat;
      TCodeStat.visible := True;
    end;
    if ComboLieu <> nil then ComboLieu.Visible := TRUE;
  end;
  //Visiblilité Champ Libre Date,Tablette,checkBox
  for num := 1 to VH_Paie.PgNbDate do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbDate then break;
    Cadre := TBevel(GetControl('BVDATE'));
    if cadre <> nil then Cadre.visible := True;
{$IFNDEF EAGLCLIENT}
    Date := THDBEdit(GetControl('PSA_DATELIBRE' + Numero));
{$ELSE}
    Date := THEdit(GetControl('PSA_DATELIBRE' + Numero));
{$ENDIF}
    if Date <> nil then Date.Visible := TRUE;
    TLieu := THLabel(GetControl('TPSA_DATELIBRE' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgLibDate1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgLibDate2;
      end;
      if Num = 3 then
      begin
        TLieu.Caption := VH_Paie.PgLibDate3;
      end;
      if Num = 4 then
      begin
        TLieu.Caption := VH_Paie.PgLibDate4;
      end;
    end;
  end; //Fin For

  for num := 1 to VH_Paie.PgNbCoche do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbCoche then break;
    Cadre := TBevel(GetControl('BVCOCHE'));
    if cadre <> nil then Cadre.visible := True;
{$IFNDEF EAGLCLIENT}
    Coche := TDBCheckBox(GetControl('PSA_BOOLLIBRE' + Numero));
{$ELSE}
    Coche := TCheckBox(GetControl('PSA_BOOLLIBRE' + Numero));
{$ENDIF}
    if Coche <> nil then
    begin
      Coche.Visible := TRUE;
      if Num = 1 then
      begin
        SetControlCaption('PSA_BOOLLIBRE1', VH_Paie.PgLibCoche1); //PT159
      end;
      if Num = 2 then
      begin
        SetControlCaption('PSA_BOOLLIBRE2', VH_Paie.PgLibCoche2); //PT159
      end;
      if Num = 3 then
      begin
        SetControlCaption('PSA_BOOLLIBRE3', VH_Paie.PgLibCoche3); //PT159
      end;
      if Num = 4 then
      begin
        SetControlCaption('PSA_BOOLLIBRE4', VH_Paie.PgLibCoche4); //PT159
      end;
    end;
  end;

  for num := 1 to VH_Paie.PgNbCombo do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbCombo then break;
    Cadre := TBevel(GetControl('BVCOMBO'));
    if cadre <> nil then Cadre.visible := True;
{$IFNDEF EAGLCLIENT}
    Combo := THDbValComboBox(GetControl('PSA_LIBREPCMB' + Numero));
{$ELSE}
    Combo := THValComboBox(GetControl('PSA_LIBREPCMB' + Numero));
{$ENDIF}
    if Combo <> nil then Combo.Visible := TRUE;
    TLieu := THLabel(GetControl('TPSA_LIBREPCMB' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo2;
      end;
      if Num = 3 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo3;
      end;
      if Num = 4 then
      begin
        TLieu.Caption := VH_Paie.PgLibCombo4;
      end;
    end;
  end;
  //Affichage des champs salaires
  for num := 1 to VH_Paie.PgNbSalLib do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbSalLib then break;
{$IFNDEF EAGLCLIENT}
    Mensuel := THDbEdit(GetControl('PSA_SALAIREMOIS' + Numero));
    Annuel := THDbEdit(GetControl('PSA_SALAIRANN' + Numero));
{$ELSE}
    Mensuel := THEdit(GetControl('PSA_SALAIREMOIS' + Numero));
    Annuel := THEdit(GetControl('PSA_SALAIRANN' + Numero));
{$ENDIF}
    if Mensuel <> nil then Mensuel.Visible := TRUE;
    if Annuel <> nil then Annuel.Visible := TRUE;
    TLieu := THLabel(GetControl('SALLIB' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib2;
      end;
      if Num = 3 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib3;
      end;
      if Num = 4 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib4;
      end;
      // PT47     06/09/2002 V585 PH Prise en compte 5 elt de salaire
      if Num = 5 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib5;
      end;
    end;
    if num = 1 then
    begin
      Tlieu := THLabel(GetControl('LBLMENSUEL'));
      if (TLieu <> nil) then TLieu.Visible := TRUE;
      Tlieu := THLabel(GetControl('LBLANNUEL'));
      if (TLieu <> nil) then TLieu.Visible := TRUE;
    end;
  end;
end;

procedure TOM_SALARIES.TauxClick(Sender: TObject);
var
  StSql: string;
{$IFNDEF EAGLCLIENT}
  LeChamp: THDBEdit;
{$ELSE}
  LeChamp: THEdit;
{$ENDIF}
begin
  if GetField('PSA_ETABLISSEMENT') = '' then exit;
  StSql := ' PAT_ETABLISSEMENT ="' + GetField('PSA_ETABLISSEMENT') + '"';
{$IFNDEF EAGLCLIENT}
  LeChamp := THDBEdit(GetControl('PSA_ORDREAT'));
{$ELSE}
  LeChamp := THEdit(GetControl('PSA_ORDREAT'));
{$ENDIF}
  if LeChamp <> nil then LookUpList(LeChamp, 'Taux AT', 'TAUXAT', 'PAT_ORDREAT', 'PAT_LIBELLE', StSql, 'PAT_ORDREAT', TRUE, -1);
end;

procedure TOM_SALARIES.OnClose;
begin
  inherited;
  //DEB PT178
  //Si on est toujours en dsinsert alors il faut supprimer les éventuels éléments dynamiques et dossier
  if ds.State in [dsinsert] then
  begin
    if not ExisteSQL('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="' + GetField('PSA_SALARIE') + '"') then
    begin
      ExecuteSQL('DELETE FROM ELTNATIONDOS WHERE PED_TYPENIVEAU="SAL"' +
        ' AND PED_VALEURNIVEAU = "' + GetField('PSA_SALARIE') + '"');
      ExecuteSQL('DELETE FROM PGHISTODETAIL WHERE PHD_SALARIE = "' + GetField('PSA_SALARIE') + '"' +
        ' AND PHD_PGTYPEINFOLS="ZLS"');
    end;
  end;
  //FIN PT178

  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
  TobZonesCompl.free;
  TobZonesComplSauv.free;
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{PT175
  if Tob_Etab <> nil then Tob_Etab.Free;
  if Tob_DonneurOrdre <> nil then Tob_DonneurOrdre.Free;
}
  if (Tob_Etab<>nil) then
     FreeAndNil (Tob_Etab);
  if (Tob_DonneurOrdre<>nil) then
     FreeAndNil (Tob_DonneurOrdre);
//FIN PT175
  //PT25 07/02/2002 V571 PH suppression historisation salarié pendant la saisie salarié puis supprimé
  // PT9 : 18/09/2001 V547 PH Rajout mise à jour ventilation analytique salarie par défaut
  if VH_Paie.PGAnalytique and VH_Paie.PGSectionAnal then
  begin // Chargement des axes et sections analytiques
    if TOB_Axes <> nil then
    begin
      TOB_Axes.Free;
      TOB_Axes := nil;
    end;
    if TOB_Section <> nil then
    begin
      TOB_Section.Free;
      TOB_Section := nil;
    end;
  end;
end;


procedure TOM_SALARIES.MajChampIdemEtab;
var
  Etab: string;
begin
  Majok := true;
  Etab := GetField('PSA_ETABLISSEMENT');
  if (Etab = '') then
    exit;
//On init une tob Réelle sur la table EtablCompl pour la gestion des Champs IdemEtablissement
  T := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etab], TRUE);
  if (T <> nil) then
  begin
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPACTIVITE', 'PSA_ACTIVITE',
      'ETB_ACTIVITE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPEDITORG', 'PSA_EDITORG',
      'ETB_EDITORG', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILANC', 'PSA_PROFILANCIEN',
      'ETB_PROFILANCIEN', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFIL', 'PSA_PROFIL', 'ETB_PROFIL',
      0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILREM', 'PSA_PROFILREM',
      'ETB_PROFILREM', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPERIODEBUL', 'PSA_PERIODBUL',
      'ETB_PERIODBUL', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILRBS', 'PSA_PROFILRBS',
      'ETB_PROFILRBS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPREDREPAS', 'PSA_REDREPAS',
      'ETB_REDREPAS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPREDRTT1', 'PSA_REDRTT1',
      'ETB_REDRTT1', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPREDRTT2', 'PSA_REDRTT2',
      'ETB_REDRTT2', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILAFP', 'PSA_PROFILAFP',
      'ETB_PROFILAFP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILAFP', 'PSA_PCTFRAISPROF',
      'ETB_PCTFRAISPROF', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILAPP', 'PSA_PROFILAPP',
      'ETB_PROFILAPP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILRET', 'PSA_PROFILRET',
      'ETB_PROFILRET', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILMUT', 'PSA_PROFILMUT',
      'ETB_PROFILMUT', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILPRE', 'PSA_PROFILPRE',
      'ETB_PROFILPRE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILTSS', 'PSA_PROFILTSS',
      'ETB_PROFILTSS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILTRANS', 'PSA_PROFILTRANS',
      'ETB_PROFILTRANS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_STANDCALEND', 'PSA_CALENDRIER',
      'ETB_STANDCALEND', 0);
    AffectChampIdemEtab(T, 'DOS', 'PSA_TYPPROFILFNAL', 'PSA_PROFILFNAL', '',
      VH_Paie.PGProfilFnal);

//Fiche Congés payés
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPROFILCGE', 'PSA_PROFILCGE',
      'ETB_PROFILCGE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPACQUISMOIS', 'PSA_NBREACQUISCP',
      'ETB_NBREACQUISCP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPNBACQUISCP', 'PSA_NBACQUISCP',
      'ETB_NBACQUISCP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPACQUISSUPP', 'PSA_NBRECPSUPP',
      'ETB_NBRECPSUPP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPTYPEMETHOD', 'PSA_VALORINDEMCP',
      'ETB_VALORINDEMCP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPTYPERELIQ', 'PSA_RELIQUAT',
      'ETB_RELIQUAT', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPAIEVALOMS', 'PSA_PAIEVALOMS',
      'ETB_PAIEVALOMS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPACQUISANC', 'PSA_BASANCCP',
      'ETB_BASANCCP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_DATANC', 'PSA_TYPDATANC',
      'ETB_TYPDATANC', 0);
//PT162-1
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPACQUISANC', 'PSA_VALANCCP',
      'ETB_VALANCCP', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_DATANC', 'PSA_DATEACQCPANC',
      'ETB_DATEACQCPANC', 0);
//FIN PT162-1
//Gestion salarié de la méthode de valorisation des CP
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPTYPEVALO', 'PSA_MVALOMS',
      'ETB_MVALOMS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_CPTYPEVALO', 'PSA_VALODXMN',
      'ETB_VALODXMN', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPEDITBULCP', 'PSA_EDITBULCP',
      'ETB_EDITBULCP', 0);
//onglet contrat
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPJOURHEURE', 'PSA_JOURHEURE',
      'ETB_JOURHEURE', 0);
    AffectChampIdemEtab(T, 'DOS', 'PSA_TYPJOURHEURE', 'PSA_JOURHEURE', '',
      VH_Paie.PGJourHeure);
    AffectChampIdemEtab(T, '', 'PSA_TYPJOURHEURE', 'PSA_TYPJOURHEURE', '',
      'DOS');

//onglet DADS-U
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPDADSFRAC', 'PSA_DADSFRACTION',
      'ETB_CODESECTION', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPRUDH', 'PSA_PRUDHCOLL',
      'ETB_PRUDHCOLL', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPRUDH', 'PSA_PRUDHSECT',
      'ETB_PRUDHSECT', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPRUDH', 'PSA_PRUDHVOTE',
      'ETB_PRUDHVOTE', 0);

//Fiche Banque
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPREGLT', 'PSA_PGMODEREGLE',
      'ETB_PGMODEREGLE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPVIRSOC', 'PSA_RIBVIRSOC',
      'ETB_RIBSALAIRE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPAIACOMPT', 'PSA_PAIACOMPTE',
      'ETB_PAIACOMPTE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPACPSOC', 'PSA_RIBACPSOC',
      'ETB_RIBACOMPTE', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPPAIFRAIS', 'PSA_PAIFRAIS',
      'ETB_PAIFRAIS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPFRAISSOC', 'PSA_RIBFRAISSOC',
      'ETB_RIBFRAIS', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPDATPAIEMENT', 'PSA_MOISPAIEMENT',
      'ETB_MOISPAIEMENT', 0);
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPDATPAIEMENT', 'PSA_JOURPAIEMENT',
      'ETB_JOURPAIEMENT', 0);
//PT162-2
    AffectChampIdemEtab(T, 'ETB', 'PSA_TYPCONVENTION', 'PSA_CONVENTION',
      'ETB_CONVENTION', 0);
//FIN PT162-2
  end;

// Champs idem profil
{On init une tob Réelle sur la table Donneur Ordre pour la gestion des Champs
Idem Profil}
  TDO := Tob_DonneurOrdre.FindFirst(['PDO_ETABLISSEMENT', 'PDO_PROFIL', 'PDO_PGMODEREGLE'],
    [Etab, GetField('PSA_PROFIL'), GetField('PSA_PGMODEREGLE')],
    TRUE);
  if (TDO <> nil) then
  begin
    AffectChampIdemEtab(TDO, 'PRO', 'PSA_TYPVIRSOC', 'PSA_RIBVIRSOC',
      'PDO_RIBASSOCIE', 0);
    AffectChampIdemEtab(TDO, 'PRO', 'PSA_TYPACPSOC', 'PSA_RIBACPSOC',
      'PDO_RIBASSOCIE', 0);
    AffectChampIdemEtab(TDO, 'PRO', 'PSA_TYPFRAISSOC', 'PSA_RIBFRAISSOC',
      'PDO_RIBASSOCIE', 0);
  end;

{PT162-2
AffectChampIdemEtab (T, 'ETB', 'PSA_TYPCONVENTION', 'PSA_CONVENTION',
                     'ETB_CONVENTION', 0);
}

end;

procedure TOM_SALARIES.AffectChampIdemEtab(TFils: TOB; StType, ChampType, ChampVal, ChampEtab: string; ValSoc: Variant);
begin
  if T <> nil then //PT149
  begin
    { DEB PT76 Reprise de la valeur au niveau etablissement ou profil ou société }
    if (ChampType <> '') and (ChampVal <> '') and (ChampEtab <> '') then
      if (GetField(ChampType) = StType) and (GetField(ChampVal) <> T.GetValue(ChampEtab)) then
      begin
        { Force mode modification pour executer le setfield }
        if not (ds.state in [dsinsert, dsedit]) then ds.edit;
        SetField(ChampVal, T.GetValue(ChampEtab));
        if ChampType = 'PSA_DATANC' then AffectDateAcqCpAnc(T.GetValue(ChampEtab));
      end;
    //Champ idem société
    if (ChampEtab = '') and (ValSoc <> '') then
      if (GetField(ChampType) = StType) and (GetField(ChampVal) <> ValSoc) then
      begin
        if not (ds.state in [dsinsert, dsedit]) then ds.edit;
        SetField(ChampVal, ValSoc);
      end;
    { FIN PT76 }
  end;
end;



procedure TOM_SALARIES.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  // PT79   : 04/11/2003 PH V_500 Correction erreur QFiche FQ  10822
  if ds.state in [dsedit, dsinsert] then
  else ds.edit;
  ParamDate(Ecran, Sender, Key);
end;
// Affectation par defaut du numero de salarié en création dans le cas d'un increment auto
// Il n'y a plus de recherche des trous car cela fait boucler sur la totalité de la table salarié

procedure TOM_SALARIES.AffectCodeNewEnr;
var
  QQuery: TQuery;
  IMax, MaxInterim: integer;
begin
  // PT2 25/06/2001 V536
  MaxInterim := 0; // Debut PT31
  if (VH_Paie.PGCodeInterim = True) and (VH_Paie.PGInterimaires = True) then
  begin
    QQuery := OpenSQL('SELECT MAX(PSI_INTERIMAIRE) FROM INTERIMAIRES', True);
    if (not QQuery.EOF) and (QQuery.Fields[0].AsString <> '') then MaxInterim := StrToInt(QQuery.Fields[0].AsString)
    else MaxInterim := 0; // PORTAGE CWAS
    Ferme(QQuery);
  end;

  // Fin PT31
  QQuery := OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES', TRUE);
  if not QQuery.Eof then
  begin // Attention, respecter le type du numéro salarié soit alpha soit numérique avec increment auto
    if QQuery.Fields[0].AsString <> '' then Imax := StrToInt(QQuery.Fields[0].AsString) // On récupère le max du numéro de salarié
    else Imax := 0; //PT29
    if MaxInterim > IMax then IMax := MaxInterim; // PT31
    if IMax < 2147483647 then //PT63
      SetField('PSA_SALARIE', IntToStr(Imax + 1));
  end
  else SetField('PSA_SALARIE', IntToStr(1));
  Ferme(QQuery);

  {
    QQuery:=OpenSQL('SELECT PSA_SALARIE FROM SALARIES ORDER BY PSA_SALARIE',TRUE) ;
    if QQuery.Fields[0].Asstring = '' then
      begin
      imax:=1;  Ferme(QQuery) ;
      SetField('PSA_SALARIE',IMax); exit;
      end
    else
      begin
      i:=1;
      While not i <= 9999999999 do
        begin
        if (QQuery.Fields[0].AsInteger = i) then
          begin
          QQuery.next; i:=i+1;
          end
        else
          begin
          imax:=i;  Ferme(QQuery) ;
          SetField('PSA_SALARIE',IMax);  exit;
          end;
        end;

    end; //End du If VH_Paie.PGIncSalarie
  }
end;

procedure TOM_SALARIES.ConvClick(Sender: TObject);
var
  StSql, Conv1, Conv2, Conv3: string;
  Q: TQuery;
{$IFNDEF EAGLCLIENT}
  ConvCol: THDBEdit;
{$ELSE}
  ConvCol: THEdit;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  ConvCol := THDBEdit(GetControl('PSA_CONVENTION'));
{$ELSE}
  ConvCol := THEdit(GetControl('PSA_CONVENTION'));
{$ENDIF}
  if ConvCol = nil then exit;
  StSql := 'Select ETB_CONVENTION,ETB_CONVENTION1,ETB_CONVENTION2 from ETABCOMPL WHERE ETB_ETABLISSEMENT="' + GetField('PSA_ETABLISSEMENT') + '"';
  Q := OpenSql(StSql, TRUE);
  if not Q.EOF then
  begin
    StSql := '';
    Conv1 := Q.FindField('ETB_CONVENTION').AsString;
    Conv2 := Q.FindField('ETB_CONVENTION1').AsString;
    Conv3 := Q.FindField('ETB_CONVENTION2').AsString;
    Ferme(Q); // Fermeture de la query car il y a une question de posée
    if Conv1 <> '' then StSql := ' PCV_CONVENTION = "' + Conv1 + '"';
    if Conv2 <> '' then StSql := StSql + ' OR PCV_CONVENTION = "' + Conv2 + '"';
    if Conv3 <> '' then StSql := StSql + ' OR PCV_CONVENTION = "' + Conv3 + '"';
    // if StSql = '' then  StSql := ' PCV_CONVENTION = "000"'         // PT37 Mise en commentaire
     //Else
    if StSql <> '' then StSql := '##PCV_PREDEFINI## ' + '(' + StSql + ')';
    LookUpList(ConvCol, 'Conventions collectives ', 'CONVENTIONCOLL', 'PCV_CONVENTION', 'PCV_LIBELLE', StSql, '', TRUE, -1);
  end
  else Ferme(Q);
end;
// DEB PT130

procedure TOM_SALARIES.PCLConvClick(Sender: TObject);
var
  StSql: string;
{$IFNDEF EAGLCLIENT}
  ConvCol: THDBEdit;
{$ELSE}
  ConvCol: THEdit;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  ConvCol := THDBEdit(GetControl('PSA_CONVENTION'));
{$ELSE}
  ConvCol := THEdit(GetControl('PSA_CONVENTION'));
{$ENDIF}
  if ConvCol = nil then exit;
  StSql := '##PCV_PREDEFINI## ';
  LookUpList(ConvCol, 'Conventions collectives ', 'CONVENTIONCOLL', 'PCV_CONVENTION', 'PCV_LIBELLE', '', '', TRUE, -1);
end;
// FIN PT130

procedure TOM_SALARIES.OnAfterUpdateRecord;
var
  OkOk: Boolean;
  rep: integer;
  st: string;
  TS, TSal: TOB;
{$IFNDEF GCGC}
//@@  {$IFDEF STATDIR}
  Q: TQuery;    // BTY 24/06/08 FQ 15526 pour ne pas redéclarer
//@@  {$ENDIF STATDIR}
  UpdateIdemPop: TUpdateIdemPop;
  TmpStringList: TStringList;
{$ENDIF}
InfTiers: Info_Tiers;
CodeAuxi, Libell, Preno, LeRapport: string;
LongLib, LongPre: integer;
{$IFDEF PAIEGRH}
  even: boolean;      //PT185
{$ENDIF PAIEGRH}
{$IFDEF STATDIR}
  GuidPerDos, GuidPer, FormeJ, Salarie, MsgErr : string;   // BTY 24/06/08 FQ 15526
//@@  Q: TQuery;                                                     //  "
  Fonction : string; // FQ 12293
{$ENDIF STATDIR}
  {$IFDEF PAIEGRH}
  //PT203
  TobInfos,TobInfos2 : TOB;
  ExisteForm,SupprForm,ExisteInsc,SupprInsc : Boolean;
  Titre, AdrResp, Mes : String;
  Texte : HTStrings;

  // Sous-fonction retournant l'adresse mail du responsable formation
  Function GetMail : String;
  var Q : TQuery;
  Begin
    Result := '';
        // Récupération de l'adresse email du responsable formation
        // Attention, celui-ci n'étant pas forcément affecté à un service, on est obligé de faire une jointure sur deportsal
        Q := OpenSQL('SELECT PSE2.PSE_EMAILPROF, US_EMAIL FROM DEPORTSAL PSE1 '+
                      'LEFT JOIN DEPORTSAL PSE2 ON PSE2.PSE_SALARIE=PSE1.PSE_RESPONSFOR '+
                      'LEFT JOIN UTILISAT ON US_AUXILIAIRE=PSE1.PSE_RESPONSFOR '+
                      'WHERE PSE1.PSE_SALARIE="'+GetField('PSA_SALARIE')+'"', True);
        If Not Q.EOF Then
        Begin
            Result := Q.FindField('US_EMAIL').AsString;
            If Result = '' Then Result := Q.FindField('PSE_EMAILPROF').AsString;
        End;
        Ferme(Q);
  End;
  {$ENDIF}
begin
  inherited;
{$IFNDEF GCGC}
  // PT151 On recharge les nouvelles valeurs des champs à surveiller.
  //  SavFieldsPopul;
  //PT151
  UpdateIdemPop := TUpdateIdemPop.create;
  TmpStringList := UpdateIdemPop.MajSALARIEPOPULSalarie(GetField('PSA_SALARIE'), Date);
  FreeAndNil(TmpStringList);
  FreeAndNil(UpdateIdemPop);
{$ENDIF}

{$IFDEF PAIEGRH}
  if (VH_PAIE.PGGESTIONCARRIERE) or (PGBundleHierarchie) then //PT84-2
    DupliquerSalInterimaire;

  //DEB PT185
  Trace := TStringList.Create ;
  even := IsDifferent(dernierecreate,PrefixeToTable(TFFiche(Ecran).TableName),'PSA_SALARIE',TFFiche(Ecran).LibelleName,Trace,TFFiche(Ecran),LeStatut);
  FreeAndNil (Trace);
  //FIN PT185
{$ENDIF PAIEGRH}

{$IFNDEF EAGLCLIENT}
  ChargementTablette(TFFiche(Ecran).TableName, '');
{$ELSE}
  ChargementTablette(TableToPrefixe(TFFiche(Ecran).TableName), '');
{$ENDIF}
  SetControlEnabled('PSA_SALARIE', False);
  InitEtab := GetField('PSA_ETABLISSEMENT');
  okok := True;

  SetControlEnabled('TCALENDRIER', OkOk);
  if (VH_Paie.PgCongesPayes) and (EtbCP) then
    SetControlEnabled('BCP', OkOk);
  SetControlEnabled('BCOMPETENCE', OkOk);
  SetControlEnabled('BBANQUE', OkOk);
  SetControlEnabled('BMEMO', OkOk);
  SetControlEnabled('BENFANT', OkOk);
  SetControlEnabled('BDUPLIQUER', OkOk);
  if GetField('PSA_MULTIEMPLOY') = 'X' then
  begin
    SetControlEnabled('BMULTIEMPLOY', True);
    SetControlVisible('RIB__BMULTIEMP', True);
  end
  else
  begin
    SetControlEnabled('BMULTIEMPLOY', False);
    SetControlVisible('RIB__BMULTIEMP', False);
  end;
  SetControlEnabled('BCONTRAT', OkOk);
  SetControlEnabled('BTNVENTIL', OkOk);

//Rajout mise à jour ventilation analytique salarie par défaut
  if (VH_Paie.PGAnalytique and VH_Paie.PGSectionAnal) and
    (Ecran.name = 'SALARIE') and (GetControlText('PSA_ANAPERSO') <> 'X') then
  begin // des Reaffectation automatique des ventilations salarie par défaut
    St := GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    rep := PGIAsk('Voulez vous réaffecter les ventilations analytiques de ' + st,
      Ecran.Caption);
    if rep = mrYes then
    begin
      st := 'SELECT PSA_SALARIE, PSA_ETABLISSEMENT, PSA_TRAVAILN1,' +
        ' PSA_TRAVAILN2, PSA_TRAVAILN3, PSA_TRAVAILN4, PSA_CODESTAT,' +
        ' PSA_LIBREPCMB1, PSA_LIBREPCMB2, PSA_LIBREPCMB3, PSA_LIBREPCMB4' +
        ' FROM SALARIES WHERE' +
        ' PSA_SALARIE="' + GetField('PSA_SALARIE') + '"';
      TSal := TOB.create('Le Salarie', nil, -1);
// chargement des infos salaries necessaires à la décomposition des sections
{PT162-4 - Optimisation
      Q:= OpenSql (st, TRUE);
      TSal.LoadDetailDB ('SALARIES', '', '', Q, FALSE, False);
      Ferme (Q);
}
      TSal.LoadDetailDBFromSQL('SALARIES', St);
//FIN PT162-4
// traitement du salarié  car la tob ne contient qu'un item
      if (TSal <> nil) and (TSal.detail.count >= 1) then
      begin
        TS := TSal.Detail[0];
        Okok := AffectVentilSalarie(TS, TOB_Axes, TOB_Section,
          VH_Paie.PGCreationSection);
      end;
      if Okok then
        PgiBox('Ventilations analytiques générées pour le salarié ' +
          GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' +
          GetField('PSA_PRENOM'), Ecran.caption)
      else
        PgiBox('Pas de ventilation analytique générée pour le salarié ' +
          GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' +
          GetField('PSA_PRENOM'), Ecran.caption);
{PT162-4
      if TSal <> nil then
         TSal.Free;
}
      FreeAndNil(TSal);
//FIN PT162-4
    end;
  end;

{$IFNDEF GCGC} // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
//PCL base allégée pas d'accès aux ressources
  if (VH_Paie.PgLienRessource) and (not EstBasePclAllegee) then
  begin
    OkOk := False;
    if Creation then
      OkOk := PGCreeRessource(GetField('PSA_SALARIE'), '', True)
    else
    begin
      Q := OpenSql('SELECT ARS_RESSOURCE' +
        ' FROM RESSOURCE WHERE' +
        ' ARS_SALARIE="' + GetField('PSA_SALARIE') + '"', True);
      if not Q.eof then //Si ressource existe maj des données complémentaires
        PGCreeRessource(GetField('PSA_SALARIE'),
          Q.FindField('ARS_RESSOURCE').AsString, True)
      else
        OkOk := PGCreeRessource(GetField('PSA_SALARIE'), '', True);
      Ferme(Q);
    end;
    if Creation then
    begin
      if OkOk then
        PGIInfo('L''affectation du salarié à une ressource s''est bien' +
          ' effectuée.', ECran.caption)
      else
        PGIInfo('Vous devez utiliser la gestion des ressources pour affecter' +
          ' le salarié.', ECran.caption);
    end;
  end;

{$ENDIF} // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

  //DEB PT182
  if CreationTier then
  begin
    Libell := GetField('PSA_LIBELLE');
    LongLib := Length(Libell);
    Preno := GetField('PSA_PRENOM');
    LongPre := Length(Preno);
    Libell := Copy(Libell, 1, LongLib) + ' ' + Copy(Preno, 1, LongPre);
    with InfTiers do
    begin
      Libelle := Copy(Libell, 1, 35);
      Adresse1 := GetField('PSA_ADRESSE1');
      Adresse2 := GetField('PSA_ADRESSE2');
      Adresse3 := GetField('PSA_ADRESSE3');
      Ville := GetField('PSA_VILLE');
      Telephone := GetField('PSA_TELEPHONE');
      CodePostal := GetField('PSA_CODEPOSTAL');
      Pays := '';
    end;
// Creation du compte de tiers en automatique et recup du numéro
    CodeAuxi := CreationTiers(InfTiers, LeRapport, 'SAL',
      GetField('PSA_SALARIE'));
    if (CodeAuxi <> '') then
      SetField('PSA_AUXILIAIRE', CodeAuxi);

    CreateYRS (GetField ('PSA_SALARIE'), '', ''); //PT155-1

    CreationTier := False;
  end;
  //FIN PT182

  if CreatEnCours then
  begin
    rep := PgiAsk('Voulez-vous créer un contrat de travail pour ce salarié ?',
      Ecran.Caption);
    if rep = mrYes then
      AglLanceFiche('PAY', 'CONTRATTRAVAIL', '', '',
        GetField('PSA_SALARIE') + ';ACTION=CREATION');
    CreatEnCours := FALSE;
  end;

  SynchroniseTiers(Getfield('PSA_AUXILIAIRE')); //PT162-3

  {$IFDEF PAIEGRH}
  //PT203
  // Si une date de sortie est positionnée, on vérifie si le salarié a des demandes de formation en cours
  If IsFieldModified('PSA_DATESORTIE') And GetField('PSA_DATESORTIE') > iDate1900 Then
  Begin
  	ExisteForm := False; SupprForm := False;
  	ExisteInsc := False; SupprInsc := False;

  	TobInfos  := TOB.Create('~LesFormationsRealisées',Nil,-1);
  	TobInfos2 := TOB.Create('~LesFormationsPrévisionnelles',Nil,-1);

  	TobInfos.AddChampSupValeur('PSA_DATESORTIE', GetField('PSA_DATESORTIE'));
  	TobInfos2.AddChampSupValeur('PSA_DATESORTIE', GetField('PSA_DATESORTIE'));

	// Recherche des formations au réalisé
  	TobInfos.LoadDetailFromSQL('SELECT "REAL" AS TYPEFOR, PFO_CODESTAGE,PST_LIBELLE||" "||PST_LIBELLE1 AS LIBELLE,PFO_DATEDEBUT,PFO_DATEFIN,PFO_ETATINSCFOR,PFO_TYPEPLANPREV '+
  								'FROM FORMATIONS LEFT JOIN STAGE ON PST_CODESTAGE=PFO_CODESTAGE AND PST_MILLESIME="0000" '+
  								'WHERE PFO_SALARIE="'+GetField('PSA_SALARIE')+'" AND PFO_EFFECTUE="-" AND (PFO_ETATINSCFOR="ATT" OR PFO_ETATINSCFOR="VAL") AND '+
  								'(PFO_DATEDEBUT>="'+UsDateTime(GetField('PSA_DATESORTIE'))+'" OR (PFO_DATEDEBUT<"'+
                                UsDateTime(GetField('PSA_DATESORTIE'))+'" AND PFO_DATEFIN>="'+UsDateTime(GetField('PSA_DATESORTIE'))+'"))');
  	If TobInfos.Detail.Count > 0 Then
  		ExisteForm := True;

    // Recherche des formations au prévisionnel
  	TobInfos2.LoadDetailFromSQL('SELECT "PREV" AS TYPEFOR, PFI_CODESTAGE,PST_LIBELLE||" "||PST_LIBELLE1 AS LIBELLE,PFI_MILLESIME,PFI_TYPEPLANPREV '+
  								'FROM INSCFORMATION LEFT JOIN STAGE ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME="0000" '+
                                'WHERE PFI_SALARIE="'+GetField('PSA_SALARIE')+'" AND PFI_ETATINSCFOR="ATT"');
  	If TobInfos2.Detail.Count > 0 Then
  		ExisteInsc := True;

	// Suppression des mouvements de formation au réalisé
  	If ExisteForm Then
  	Begin
  		If PGIAsk(TraduireMemoire('Il y a des mouvements de formation pour ce salarié après la date de sortie.#13#10'+
  			      'Voulez-vous annuler les demandes en cours?')) = mrYes Then
  		Begin
  			ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="NAN" WHERE PFO_SALARIE="'+GetField('PSA_SALARIE')+
                       '" AND PFO_EFFECTUE="-" AND (PFO_ETATINSCFOR="ATT" OR PFO_ETATINSCFOR="VAL") AND (PFO_DATEDEBUT>="'+UsDateTime(GetField('PSA_DATESORTIE'))+
                       '" OR (PFO_DATEDEBUT<"'+UsDateTime(GetField('PSA_DATESORTIE'))+'" AND PFO_DATEFIN>="'+
                       UsDateTime(GetField('PSA_DATESORTIE'))+'"))');
  			SupprForm := True;
  		End;
  	End;

	// Suppression des mouvements de formation au prévisionnel
  	If ExisteInsc Then
  	Begin
  		If ExisteForm Then
  			Mes := 'Il y a également des '
  		Else
  			Mes := 'Attention, il y a ';
  		Mes := Mes + 'formations prévues au budget prévisionnel pour ce salarié.#13#10Voulez-vous annuler les demandes en cours?';
  		If PGIAsk(TraduireMemoire(Mes)) = mrYes Then
  		Begin
  			ExecuteSQL('UPDATE INSCFORMATION SET PFI_ETATINSCFOR="NAN" WHERE PFI_SALARIE="'+GetField('PSA_SALARIE')+
                       '" AND PFI_ETATINSCFOR="ATT"');
  			SupprInsc := True;
  		End;
  	End;

  	// Envoi d'un mail au service formation dans le cas d'annulation(s) de formation(s)
  	If (SupprForm Or SupprInsc) Then
  	Begin
        If Not SupprForm Then TobInfos.Detail.Clear;
        If Not SupprInsc Then TobInfos2.Detail.Clear;
        TobInfos2.ChangeParent(TobInfos,-1,True);

        // Création du corps du mail
		PrepareMailFormation(GetField('PSA_SALARIE'),'','SORTIESALARIE',TobInfos,Titre,Texte);

        // Récupération de l'adresse email du responsable formation
        AdrResp := GetMail;

        // Envoi du mail
  		SendMail(Titre, AdrResp, VH_PAIE.PGForMailAdr, Texte, '', V_PGI.MailMethod<>mmOutLook, 1,'','',False,False) ;
  	End
    // Si aucune annulation du service du personnel, on envoie tout de même un mail pour avertir les responsables formation
    Else If (ExisteInsc) Or (ExisteForm) Then
    Begin
        TobInfos2.ChangeParent(TobInfos,-1,True);

        // Création du corps du mail
		PrepareMailFormation(GetField('PSA_SALARIE'),'','SORTIESALARIENOP',TobInfos,Titre,Texte);

        // Récupération de l'adresse email du responsable formation
        AdrResp := GetMail;

        // Envoi du mail
  		SendMail(Titre, AdrResp, VH_PAIE.PGForMailAdr, Texte, '', V_PGI.MailMethod<>mmOutLook, 1,'','',False,False) ;
    End;
  End;
  {$ENDIF}
  {$IFDEF STATDIR}
  // BTY 24/06/08 FQ 15526
  if (V_PGI.ModePCL = '1') then
  begin
     GuidPerDos := GetColonneSQL('DOSSIER','DOS_GUIDPER','DOS_NODOSSIER="' + PGRendNoDossier + '"');
     FormeJ     := GetColonneSQL('JURIDIQUE','JUR_FORME','JUR_GUIDPERDOS="' + GuidPerDos + '"');
     if  ((FormeJ = 'SARL') or (FormeJ = 'SA') or (FormeJ = 'SAS') or (FormeJ = 'SASU')) {FQ 12293}
     and (GetField('PSA_DADSPROF') = '13') and (GetField('PSA_DADSCAT') = '01')
     and (GetField('PSA_NUMEROSS') <> '') then
     begin
       // Vérif no SS attribué à un seul dirigeant dans la Paie
       if DPSSUniqueSalarie (PGRendNoDossier, GetField ('PSA_NUMEROSS'), False, MsgErr, Salarie) then
       begin
           // Vérif no SS attribué à un seul gérant dans le DP
           // FQ 12293
           if DPGerantUniqueDP (PGRendNoDossier, GetField ('PSA_NUMEROSS'), FormeJ, GuidPer, Fonction) then
           begin
             // MAJ fiche de la personne dans le DP
             DPMajPersonne_Paie (PGRendNoDossier, Salarie, GuidPer, False, Fonction);
           end;
       end;
     end;
  end;
 {$ENDIF}
end;

procedure TOM_SALARIES.AffectDateAcqCpAnc(ValTypDatAnc: string);
var
  LaDate: TdateTime;
  jour, mois: string;
  aa, mm, jj: Word;
begin
  if (ValTypDatAnc = '1') or (ValTypDatAnc = '3') then //date entrée, ancienneté du salarié
  begin
    if (ValTypDatAnc = '1') then
      Ladate := getfield('PSA_DATEENTREE')
    else
      Ladate := getfield('PSA_DATEANCIENNETE');
    if Ladate > 0 then
    begin
      decodedate(Ladate, aa, mm, jj);
      if jj > 9 then jour := inttostr(jj) else jour := '0' + inttostr(jj);
      if mm > 9 then mois := inttostr(mm) else mois := '0' + inttostr(mm);
      if GetField('PSA_DATEACQCPANC') <> (jour + mois) then //PT18 SetField systématique
        setfield('PSA_DATEACQCPANC', jour + mois);
    end;
  end
  else
    if (ValTypDatAnc = '0') or (ValTypDatAnc = '4') then // date debut ou fin d'exercice cp
    begin
      if (T <> nil) then Ladate := T.getValue('ETB_DATECLOTURECPN') else Ladate := 0;
      if Ladate > 0 then
      begin
        if (ValTypDatAnc = '0') then
          decodedate(Ladate + 1, aa, mm, jj)
        else
          decodedate(Ladate, aa, mm, jj);
        if jj > 9 then jour := inttostr(jj) else jour := '0' + inttostr(jj);
        if mm > 9 then mois := inttostr(mm) else mois := '0' + inttostr(mm);
        if GetField('PSA_DATEACQCPANC') <> (jour + mois) then //PT18 SetField systématique
          setfield('PSA_DATEACQCPANC', jour + mois);
      end;
    end;
  // else setfield('PSA_DATEACQCPANC','')   modif SB 11/06/01
end;

procedure TOM_SALARIES.OnChangeStandCalend(Sender: TObject);
var
  Init: word;
  Calend: string;
begin
  Calend := getfield('PSA_STANDCALEND');
  if (Calend <> 'PER') and (typeCalend = 'PER') then
  begin
    Init := HShowMessage('1;Calendrier salarié;Attention, vous allez supprimer la personnalisation du calendrier pour ce salarié.#13#10 Voulez-vous continuer ?;Q;YN;N;N;', '',
      '');
    if Init = mrNo then
    begin
      Calend := '';
      typeCalend := 'PER';
      SetField('PSA_STANDCALEND', 'PER');
      exit;
    end
    else
    begin
      typeCalend := getfield('PSA_STANDCALEND');
      SupTablesLiees('CALENDRIER', 'ACA_SALARIE', getfield('PSA_SALARIE'), '', True);
    end;
  end;
  typeCalend := getfield('PSA_STANDCALEND');
end;
// PT22 procédure pour rendre le bouton Multi Employeur accesible ou non

procedure TOM_SALARIES.ClickMultiEmpl(Sender: TObject);
begin
  if GetControlText('PSA_MULTIEMPLOY') = 'X' then
  begin
    SetControlEnabled('BMULTIEMPLOY', True);
    SetControlEnabled('PSA_SALAIREMULTI', True);
    SetControlVisible('RIB__BMULTIEMP', True);
  end
  else
  begin
    SetControlEnabled('BMULTIEMPLOY', False);
    SetControlEnabled('PSA_SALAIREMULTI', False);
    SetField('PSA_SALAIREMULTI', 0);
    SetControlVisible('RIB__BMULTIEMP', False);
  end;
end;

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.ImpressionSal(Sender: TObject); //PT26
var
  Pages: TPageControl;
  Defaut: THEdit;
  St: string;
begin
  Defaut := ThEdit(getcontrol('DOSSIER'));
  if Defaut <> nil then
    //  Defaut.text:=V_PGI_env.LibDossier; //PT33 Mise en commentaire
    Defaut.text := GetParamSoc('SO_LIBELLE');
  Pages := TPageControl(GetControl('PAGES'));
  if Pages <> nil then
  begin
    St := 'SELECT * FROM SALARIES WHERE PSA_SALARIE="' + GetField('PSA_SALARIE') + '"';
    LanceEtat('E', 'PSA', 'PES', True, False, False, Pages, St, '', False);
  end;
end;
{$ENDIF} // CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

// PT30 Fonction de traitement des champs idem etab

procedure TOM_SALARIES.IdemEtab(ChampTyp, LeChamp, ChampEtab: string; F: TField);
var{$IFNDEF EAGLCLIENT}
  Combo: THDBValComboBox;
{$ELSE}
  Combo: THValComboBox;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  Combo := THDBValComboBox(getcontrol(LeChamp));
{$ELSE}
  Combo := THValComboBox(getcontrol(LeChamp));
{$ENDIF}
  if (F.FieldName = (ChampTyp)) then
  begin
    if GetField(ChampTyp) = 'ETB' then
    begin
      if (T <> nil) then
        if (T.GetValue(ChampEtab) <> GetField(LeChamp)) and (GetControl(Lechamp) <> nil) then // PT110 PT112
          SetField(Lechamp, T.GetValue(ChampEtab));
      SetControlEnabled(Lechamp, False);
    end
    else
      if Combo = nil then SetControlEnabled(Lechamp, True)
      else if Combo.Color <> CouleurHisto then SetControlEnabled(Lechamp, True);
  end;
end;

procedure TOM_SALARIES.ControleHoraires(Hebdo, Mensuel, Annuel: Double); //PT38
var
  mess: string;
begin
  mess := '';
  if TestHoraire(Hebdo, Mensuel) = True then
  begin
    Mess := Mess + '- L''horaire hebdomadaire ne peut être supérieur à l''horaire mensuel';
    SetFocuscontrol('PSA_HORAIREMOIS');
  end;
  if TestHoraire(Hebdo, Annuel) = True then
  begin
    Mess := Mess + '#13#10 - L''horaire hebdomadaire ne peut être supérieur à l''horaire annuel';
    SetFocusControl('PSA_HORANNUEL');
  end;
  if TestHoraire(Mensuel, Annuel) = True then
  begin
    Mess := Mess + '#13#10 - L''horaire mensuel ne peut être supérieur à l''horaire annuel';
    SetFocusControl('PSA_HORANNUEL');
  end;
  if Mess <> '' then
  begin
    LastError := 1;
    PgiBox(Mess, 'Contrat : saisie des horaires');
    SetActiveTabSheet('PCONTRAT');
  end;
end;

function TOM_SALARIES.TestHoraire(HoraireInf, HoraireSup: Double): Boolean;
begin
  Result := False;
  if HoraireSup <> 0 then
  begin
    if HoraireInf > HoraireSup then Result := True;
  end;
end;

procedure TOM_SALARIES.BEnfantsClick(Sender: TObject);
var
  QNb: TQuery;
  Nb: Integer;
begin

  // PT56   : 17/12/2002 PH V591 Passage en paramètre du nom+prenom du salarie pour acces enfants salariés
  AGLLanceFiche('PAY', 'ENFSALARIE', GetField('PSA_SALARIE'), '', GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM'));
  QNb := OpenSQL('SELECT COUNT (PEF_ACHARGE) AS NBACHARGE FROM ENFANTSALARIE' +
    ' WHERE (PEF_SALARIE="' + GetField('PSA_SALARIE') + '") AND (PEF_ACHARGE="X")', TRUE);
  if QNb.EOF then Nb := 0 // PORTAGECWAS
  else Nb := QNb.FindField('NBACHARGE').AsInteger; //PT61
  if (GetField('PSA_PERSACHARGE') <> nb) then
  begin
    PGIBox('Le nombre de personnes à charge doit être égal à ' + IntToSTr(nb) + '.', 'Personnes à charge'); //PT158
    SetField('PSA_PERSACHARGE', Nb);
    SetControlText('EPERSACHARGE', IntToStr(Nb));
    SetFocusControl('EPERSACHARGE');
  end;
  Ferme(QNb);
end;
// PT59   : 03/01/2003 PH V591 Controles et affichage des dates dans onglet emploi pour ne pas afficherv idate1900

procedure TOM_SALARIES.ControleDate(LeChamp, LeBouton: string);
begin
  if GetField(LeChamp) > IDate1900 then
  begin
    SetControlVisible(LeChamp, TRUE);
    SetControlVisible('Z' + LeChamp, FALSE);
    SetControlVisible(LeBouton, FALSE);
  end
  else
  begin
    SetControlVisible(LeChamp, FALSE);
    SetControlVisible('Z' + LeChamp, TRUE);
    SetControlVisible(LeBouton, TRUE);
  end;

end;

procedure TOM_SALARIES.BtnDate(LeBouton: string);
var
  Edit: THEdit;
begin
  Edit := THedit(GetControl(LeBouton));
  if Edit <> nil then
  begin //PT98 Activation aussi du controle de la date en sortie de zone
    Edit.OnClick := DateActivee;
    Edit.OnExit := DateActivee;
    Edit.OnEnter := DateActivee; // PT136
  end;
end;

procedure TOM_SALARIES.DateActivee(Sender: TObject);
var
  LeNom: string;
begin
  if sender = nil then exit;
  LeNom := THEdit(Sender).Name;
  SetControlVisible(LeNom, FALSE);
  LeNom := Copy(LeNom, 2, Strlen(PChar(LeNom)) - 1);
  SetControlVisible(LeNom, TRUE);
  // DEB PT98
  if (THEdit(Sender).Text = '') then
  begin
    SetField(LeNom, Idate1900);
    SetControlEnabled(LeNom, True); // PT136
    SetFocusControl(LeNom); // PT136
    exit;
  end;
  //  SetFocusControl(LeNom);
  //  if (GetField(LeNom) < Idate1900) then SetField(LeNom, Idate1900);
  if (StrToDate(THEdit(Sender).Text) < Idate1900) then SetField(LeNom, Idate1900)
  else SetField(LeNom, StrToDate(THEdit(Sender).Text));
  // FIN PT98
end;
// FIN PT59

{ DEB PT68 }

procedure TOM_SALARIES.ChargeProfilActivite(LeChamp, Nat: string);
begin
  if RechDom('PGTYPEPROFIL', Nat, True) = 'ACT' then
  begin
    if GetField('PSA_ACTIVITE') <> '' then
      SetControlProperty(LeChamp, 'Plus', 'AND (PPI_ACTIVITE="' + GetField('PSA_ACTIVITE') + '" OR PPI_ACTIVITE="" OR PPI_ACTIVITE IS NULL)')
    else
      SetControlProperty(LeChamp, 'Plus', '');
  end;
end;
{ FIN PT68 }
{ DEB PT68 }

procedure TOM_SALARIES.OnEnterProfil(Sender: Tobject);
var
  Name: string;
begin
{$IFNDEF GCGC}
  CurProfil := 'PROFIL'; //pt188
{$ENDIF}

  if sender = nil then exit;
  { Affectation de la propriété plus pour chaque profil gérant l'activité }
{$IFNDEF EAGLCLIENT}
  Name := THDBValComboBox(sender).Name;
{$ELSE}
  Name := THValComboBox(sender).Name;
{$ENDIF}
  if Name = 'PSA_PROFILAFP' then ChargeProfilActivite('PSA_PROFILAFP', 'AFP')
  else if Name = 'PSA_PROFILANCIEN' then ChargeProfilActivite('PSA_PROFILANCIEN', 'ANC')
  else if Name = 'PSA_PROFILAPP' then ChargeProfilActivite('PSA_PROFILAPP', 'APP')
  else if Name = 'PSA_PROFILCDD' then ChargeProfilActivite('PSA_PROFILCDD', 'CDD')
  else if Name = 'PSA_PROFILCGE' then ChargeProfilActivite('PSA_PROFILCGE', 'CGE')
  else if Name = 'PSA_PROFILFNAL' then ChargeProfilActivite('PSA_PROFILFNAL', 'FNL')
  else if Name = 'PSA_PROFILMUT' then ChargeProfilActivite('PSA_PROFILMUT', 'MUT')
  else if Name = 'PSA_PERIODBUL' then ChargeProfilActivite('PSA_PERIODBUL', 'PER')
  else if Name = 'PSA_PROFILPRE' then ChargeProfilActivite('PSA_PROFILPRE', 'PRE')
  else if Name = 'PSA_PROFIL' then ChargeProfilActivite('PSA_PROFIL', 'PRO')
  else if Name = 'PSA_PROFILRBS' then ChargeProfilActivite('PSA_PROFILRBS', 'RBS')
  else if Name = 'PSA_PROFILREM' then ChargeProfilActivite('PSA_PROFILREM', 'REM')
  else if Name = 'PSA_PROFILRET' then ChargeProfilActivite('PSA_PROFILRET', 'RET')
  else if Name = 'PSA_REDREPAS' then ChargeProfilActivite('PSA_REDREPAS', 'RRE')
  else if Name = 'PSA_REDRTT1' then ChargeProfilActivite('PSA_REDRTT1', 'RT1')
  else if Name = 'PSA_REDRTT2' then ChargeProfilActivite('PSA_REDRTT2', 'RT2')
  else if Name = 'PSA_PROFILTPS' then ChargeProfilActivite('PSA_PROFILTPS', 'TPS')
  else if Name = 'PSA_PROFILTRANS' then ChargeProfilActivite('PSA_PROFILTRANS', 'TRA')
  else if Name = 'PSA_PROFILTSS' then ChargeProfilActivite('PSA_PROFILTSS', 'TSS');
end;
{ FIN PT68 }

procedure TOM_SALARIES.ChargeGrilleRib(Grid: THGrid; StType: string);
var
  StWhere: string;
  Q: Tquery;
  T1: Tob;
begin
  { DEB PT77 Chargement des données, mise en forme des grilles RIB }
  if Grid = nil then exit;
  Grid.RowCount := 5;
  Grid.FixedRows := 1;
  if StType = 'ACP' then StWhere := 'AND R_ACOMPTE="X"'
  else if StType = 'FPR' then StWhere := 'AND R_FRAISPROF="X"'
  else StWhere := '';
  Q := OpenSql('SELECT R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB ' +
    'FROM RIB WHERE R_AUXILIAIRE= "' + GetField('PSA_AUXILIAIRE') + '" ' + StWhere, TRUE);
  if not Q.EOF then
  begin
    T1 := TOB.Create('Les rib', nil, -1);
    T1.LoadDetailDb('RIB', '', '', Q, FALSE, False); //PT167
    if T1.detail.count > 0 then
      T1.PutGridDetail(Grid, TRUE, TRUE, 'R_DOMICILIATION;R_ETABBQ;R_GUICHET;R_NUMEROCOMPTE;R_CLERIB', True);
    T1.Free;
  end
  else
    Grid.RowCount := 0;
  Ferme(Q);
  if Grid.RowCount > 0 then
  begin
    Grid.Titres.Clear;
    Grid.Titres.Add('Domiciliation');
    Grid.Titres.Add('Banque');
    Grid.Titres.Add('Guichet');
    Grid.Titres.Add('Numéro de compte');
    Grid.Titres.Add('Clé');
    Grid.UpdateTitres;
  end;
  { FIN PT77 }
end;

{$IFDEF PAIEGRH}
procedure TOM_SALARIES.DupliquerSalInterimaire; //PT84
var
  TobInterim, T, TobChamp: Tob;
  Q: TQuery;
  Salarie, TypeChamp, Champ, Prefixe, Suffixe: string;
  UneDate: TDateTime;
  i: Integer;
begin
  Q := OpenSQL('SELECT * FROM DECHAMPS WHERE DH_PREFIXE="PSI"', True);
  TobChamp := Tob.Create('MesChamps', nil, -1);
  TobChamp.LoadDetailDB('MesChamps', '', '', Q, False);
  Ferme(Q);
  Salarie := GetField('PSA_SALARIE');
  Q := OpenSQL('SELECT *' +
    ' FROM INTERIMAIRES WHERE' +
    ' PSI_INTERIMAIRE="' + Salarie + '" AND' +
    ' PSI_TYPEINTERIM="SAL"', True);
  TobInterim := Tob.Create('INTERIMAIRES', nil, -1);
  TobInterim.LoadDetailDB('INTERIMAIRES', '', '', Q, False);
  Ferme(Q);
  T := TobInterim.FindFirst(['PSI_INTERIMAIRE'], [Salarie], False);
  if T = nil then
    T := Tob.Create('INTERIMAIRES', nil, -1);
  for i := 0 to TobChamp.Detail.Count - 1 do
  begin
    Champ := TobChamp.Detail[i].GetValue('DH_NOMCHAMP');
    TypeChamp := TobChamp.Detail[i].GetValue('DH_TYPECHAMP');
    Prefixe := ReadTokenPipe(Champ, '_');
    Suffixe := Champ;
    if (Suffixe = 'DISPOMOIS') or (Suffixe = 'SALAIREDEM') then
      continue
// Attention aux champs existants dans la table interimaires et qui n'existent
// pas dans la table salaries
    else
      if Suffixe = 'TYPEINTERIM' then
        T.PutValue('PSI_TYPEINTERIM', 'SAL')
      else
        if Suffixe = 'INTERIMAIRE' then
          T.PutValue('PSI_INTERIMAIRE', Salarie)
        else
          if Suffixe = 'NODOSSIER' then
            T.PutValue('PSI_NODOSSIER', V_PGI.NoDossier)
          else if Suffixe = 'PREDEFINI' then
            T.PutValue('PSI_PREDEFINI', 'DOS')
          else if ((suffixe = 'EMAIL') or (Suffixe = 'PORTABLE')) then
            T.PutValue('PSI_' + Suffixe, '')
          else
          begin
            if TypeChamp = 'DATE' then
            begin
              if GetField('PSA_' + Suffixe) <> Null then
                UneDate := GetField('PSA_' + Suffixe)
              else
                UneDate := IDate1900;
              T.PutValue('PSI_' + Suffixe, UneDate);
            end
            else
              T.PutValue('PSI_' + Suffixe, GetField('PSA_' + Suffixe));
          end;
  end;
  FreeAndNil(TobChamp);
  T.InsertOrUpdateDB(False);
  FreeAndNil(T);
  FreeAndNil(TobInterim);
end;
// DEB PT95
{$ENDIF PAIEGRH}

procedure TOM_SALARIES.AppelIsoflex(Sender: TObject);
begin
  AglIsoflexViewDoc(NomHalley, Ecran.Name, 'SALARIES', 'PSA_SALARIE', '', GetField('PSA_SALARIE'), '');
end;

function TOM_SALARIES.GereIsoflex: Boolean;
begin
  if AglIsoflexPresent then
  begin
    SetControlVisible('BISOFLEX', TRUE);
    result := TRUE;
  end
  else
    result := FALSE;
end;
// FIN PT95
// DEB PT96

(*
procedure TOM_SALARIES.PCasOnClick(Sender: TObject);
begin
  ForceUpdate;
end;
// FIN PT96
*)

procedure TOM_SALARIES.MajConfidentialite; //PT97
begin
  ExecuteSQL('UPDATE ABSENCESALARIE SET PCN_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PCN_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE CONTRATTRAVAIL SET PCI_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PCI_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE FORMATIONS SET PFO_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PFO_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOANALPAIE SET PHA_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PHA_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOBULLETIN SET PHB_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PHB_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOCUMSAL SET PHC_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PHC_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOPARTICIP SET PHP_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PHP_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOSAISPRIM SET PSP_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PSP_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE HISTOSAISRUB SET PSD_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PSD_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE INSCFORMATION SET PFI_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PFI_SALARIE="' + GetField('PSA_SALARIE') + '"');
  ExecuteSQL('UPDATE PAIEENCOURS SET PPU_CONFIDENTIEL="' + getField('PSA_CONFIDENTIEL') + '" WHERE PPU_SALARIE="' + GetField('PSA_SALARIE') + '"');
end;
//DEB PT99

procedure TOM_SALARIES.OnExitRegul(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  MSpin: THDBSpinEdit;
{$ELSE}
  MSpin: THSpinEdit;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  MSpin := THDBSpinEdit(getcontrol('PSA_REGULANCIEN'));
{$ELSE}
  MSpin := THSpinEdit(getcontrol('PSA_REGULANCIEN'));
{$ENDIF}
  if MSpin <> nil then
  begin
    if MSpin.Text = '' then SetField('PSA_REGULANCIEN', 0);
  end;
end;
// FIN PT99
// DEB PT103 Gestion de la confidentialité en CWAS

procedure TOM_SALARIES.OnClickConfidentiel(Sender: TObject);
begin
  if (Sender = nil) then exit;
  if (not (DS.State in [dsEdit, dsInsert])) then DS.Edit;
  if (GetCheckBoxState('PCONFIDENTIEL') = cbChecked) then SetField('PSA_CONFIDENTIEL', '1')
  else SetField('PSA_CONFIDENTIEL', '0');
end;
// FIN PT103

// DEBUT PT104

procedure TOM_SALARIES.DuplicationSalarie(leSalarie: string);
var
Q : TQuery;
begin
Q:= OpenSQL ('SELECT *'+
             ' FROM SALARIES WHERE'+
             ' PSA_SALARIE="'+LeSalarie+'"', True);
if (not Q.Eof) then
   begin
   SetField ('PSA_ETABLISSEMENT', Q.FindField ('PSA_ETABLISSEMENT').AsString);
//DEB PT181
//Onglet emploi
   SetField ('PSA_CATDADS', Q.FindField ('PSA_CATDADS').AsString);
   SetField ('PSA_CATBILAN', Q.FindField ('PSA_CATBILAN').AsString);
//FIN PT181
//Onglet affectation
   SetField ('PSA_CONVENTION', Q.FindField ('PSA_CONVENTION').AsString);
   SetField ('PSA_TYPCONVENTION', Q.FindField ('PSA_TYPCONVENTION').AsString); //PT129
   SetField ('PSA_CODEEMPLOI', Q.FindField ('PSA_CODEEMPLOI').AsString);
   SetField ('PSA_LIBELLEEMPLOI', Q.FindField ('PSA_LIBELLEEMPLOI').AsString);
   SetField ('PSA_QUALIFICATION', Q.FindField ('PSA_QUALIFICATION').AsString);
   SetField ('PSA_COEFFICIENT', Q.FindField ('PSA_COEFFICIENT').AsString);
   SetField ('PSA_INDICE', Q.FindField ('PSA_INDICE').AsString);
   SetField ('PSA_NIVEAU', Q.FindField ('PSA_NIVEAU').AsString);
   SetField ('PSA_CODESTAT', Q.FindField ('PSA_CODESTAT').AsString);
   SetField ('PSA_TRAVAILN1', Q.FindField ('PSA_TRAVAILN1').AsString);
   SetField ('PSA_TRAVAILN2', Q.FindField ('PSA_TRAVAILN2').AsString);
   SetField ('PSA_TRAVAILN3', Q.FindField ('PSA_TRAVAILN3').AsString);
   SetField ('PSA_TRAVAILN4', Q.FindField ('PSA_TRAVAILN4').AsString);
   SetField ('PSA_PRISEFFECTIF', Q.FindField ('PSA_PRISEFFECTIF').AsString);
   SetField ('PSA_UNITEPRISEFF', Q.FindField ('PSA_UNITEPRISEFF').AsFloat);
   SetField ('PSA_TYPEDITORG', Q.FindField ('PSA_TYPEDITORG').AsString);
   SetField ('PSA_EDITORG', Q.FindField ('PSA_EDITORG').AsString);
//profils
   SetField ('PSA_TYPACTIVITE', Q.FindField ('PSA_TYPACTIVITE').AsString);
   SetField ('PSA_ACTIVITE', Q.FindField ('PSA_ACTIVITE').AsString);
   SetField ('PSA_TYPPROFILREM', Q.FindField ('PSA_TYPPROFILREM').AsString);
   SetField ('PSA_PROFILREM', Q.FindField ('PSA_PROFILREM').AsString);
   SetField ('PSA_TYPPROFIL', Q.FindField ('PSA_TYPPROFIL').AsString);
   SetField ('PSA_PROFIL', Q.FindField ('PSA_PROFIL').AsString);
   SetField ('PSA_TYPPERIODEBUL', Q.FindField ('PSA_TYPPERIODEBUL').AsString);
   SetField ('PSA_PERIODBUL', Q.FindField ('PSA_PERIODBUL').AsString);
   SetField ('PSA_PROFILTPS', Q.FindField ('PSA_PROFILTPS').AsString);
//     SetField('PSA_SALAIRETHEO',Q.FindField ('PSA_SALAIRETHEO').AsFloat);
//Autre profils
   SetField ('PSA_TYPPROFILRBS', Q.FindField ('PSA_TYPPROFILRBS').AsString);
   SetField ('PSA_PROFILRBS', Q.FindField ('PSA_PROFILRBS').AsString);
   SetField ('PSA_TYPREDRTT2', Q.FindField ('PSA_TYPREDRTT2').AsString);
   SetField ('PSA_REDRTT2', Q.FindField ('PSA_REDRTT2').AsString);
   SetField ('PSA_TYPREDRTT1', Q.FindField ('PSA_TYPREDRTT1').AsString);
   SetField ('PSA_REDRTT1', Q.FindField ('PSA_REDRTT1').AsString);
   SetField ('PSA_TYPREDREPAS', Q.FindField ('PSA_TYPREDREPAS').AsString);
   SetField ('PSA_REDREPAS', Q.FindField ('PSA_REDREPAS').AsString);
   SetField ('PSA_TYPPROFILAFP', Q.FindField ('PSA_TYPPROFILAFP').AsString);
   SetField ('PSA_PROFILAFP', Q.FindField ('PSA_PROFILAFP').AsString);
   SetField ('PSA_PCTFRAISPROF', Q.FindField ('PSA_PCTFRAISPROF').AsFloat);
   SetField ('PSA_TYPPROFILAPP', Q.FindField ('PSA_TYPPROFILAPP').AsString);
   SetField ('PSA_PROFILAPP', Q.FindField ('PSA_PROFILAPP').AsString);
   SetField ('PSA_TYPPROFILRET', Q.FindField ('PSA_TYPPROFILRET').AsString);
   SetField ('PSA_PROFILRET', Q.FindField ('PSA_PROFILRET').AsString);
   SetField ('PSA_TYPPROFILMUT', Q.FindField ('PSA_TYPPROFILMUT').AsString);
   SetField ('PSA_PROFILMUT', Q.FindField ('PSA_PROFILMUT').AsString);
   SetField ('PSA_TYPPROFILPRE', Q.FindField ('PSA_TYPPROFILPRE').AsString);
   SetField ('PSA_PROFILPRE', Q.FindField ('PSA_PROFILPRE').AsString);
   SetField ('PSA_TYPPROFILTSS', Q.FindField ('PSA_TYPPROFILTSS').AsString);
   SetField ('PSA_PROFILTSS', Q.FindField ('PSA_PROFILTSS').AsString);
   SetField ('PSA_TYPPROFILFNAL', Q.FindField ('PSA_TYPPROFILFNAL').AsString);
   SetField ('PSA_PROFILFNAL', Q.FindField ('PSA_PROFILFNAL').AsString);
   SetField ('PSA_TYPPROFILTRANS', Q.FindField ('PSA_TYPPROFILTRANS').AsString);
   SetField ('PSA_PROFILTRANS', Q.FindField ('PSA_PROFILTRANS').AsString);
//Contrat
   SetField ('PSA_HORAIREMOIS', Q.FindField ('PSA_HORAIREMOIS').AsFloat);
   SetField ('PSA_HORHEBDO', Q.FindField ('PSA_HORHEBDO').AsFloat);
   SetField ('PSA_HORANNUEL', Q.FindField ('PSA_HORANNUEL').AsFloat);
{     SetField('PSA_TAUXHORAIRE',Q.FindField('PSA_TAUXHORAIRE').AsFloat);
      SetField('PSA_SALAIREMOIS1',Q.FindField('PSA_SALAIREMOIS1').AsFloat);
      SetField('PSA_SALAIREMOIS2',Q.FindField('PSA_SALAIREMOIS2').AsFloat);
      SetField('PSA_SALAIREMOIS3',Q.FindField('PSA_SALAIREMOIS3').AsFloat);
      SetField('PSA_SALAIREMOIS4',Q.FindField('PSA_SALAIREMOIS4').AsFloat);
      SetField('PSA_SALAIREMOIS5',Q.FindField('PSA_SALAIREMOIS5').AsFloat);
      SetField('PSA_SALAIRANN1',Q.FindField('PSA_SALAIRANN1').AsFloat);
      SetField('PSA_SALAIRANN2',Q.FindField('PSA_SALAIRANN2').AsFloat);
      SetField('PSA_SALAIRANN3',Q.FindField('PSA_SALAIRANN3').AsFloat);
      SetField('PSA_SALAIRANN4',Q.FindField('PSA_SALAIRANN4').AsFloat);
      SetField('PSA_SALAIRANN5',Q.FindField('PSA_SALAIRANN5').AsFloat);  }
   SetField ('PSA_STANDCALEND', Q.FindField ('PSA_STANDCALEND').AsString);
   SetField ('PSA_CALENDRIER', Q.FindField ('PSA_CALENDRIER').AsString);
   SetField ('PSA_TYPJOURHEURE', Q.FindField ('PSA_TYPJOURHEURE').AsString);
   SetField ('PSA_JOURHEURE', Q.FindField ('PSA_JOURHEURE').AsString);
   SetField ('PSA_PROFILCDD', Q.FindField ('PSA_PROFILCDD').AsString);  //PT181
//DADS
   SetField ('PSA_REMPOURBOIRE', Q.FindField ('PSA_REMPOURBOIRE').AsString);
   SetField ('PSA_TRAVETRANG', Q.FindField ('PSA_TRAVETRANG').AsString);
   SetField ('PSA_REGIMESS', Q.FindField ('PSA_REGIMESS').AsString);
//DEB PT164
   SetField ('PSA_TYPEREGIME', Q.FindField ('PSA_TYPEREGIME').AsString);
   SetField ('PSA_REGIMEMAL', Q.FindField ('PSA_REGIMEMAL').AsString);
   SetField ('PSA_REGIMEAT', Q.FindField ('PSA_REGIMEAT').AsString);
   SetField ('PSA_REGIMEVIP', Q.FindField ('PSA_REGIMEVIP').AsString);
   SetField ('PSA_REGIMEVIS', Q.FindField ('PSA_REGIMEVIS').AsString);
//FIN PT164
   SetField ('PSA_ORDREAT', Q.FindField ('PSA_ORDREAT').AsString);
   SetField ('PSA_DADSCAT', Q.FindField ('PSA_DADSCAT').AsString);
   SetField ('PSA_CONDEMPLOI', Q.FindField ('PSA_CONDEMPLOI').AsString);
   SetField ('PSA_DADSPROF', Q.FindField ('PSA_DADSPROF').AsString);
   SetField ('PSA_TAUXPARTIEL', Q.FindField ('PSA_TAUXPARTIEL').AsFloat);
   SetField ('PSA_UNITETRAVAIL', Q.FindField ('PSA_UNITETRAVAIL').AsString);
   SetField ('PSA_TYPDADSFRAC', Q.FindField ('PSA_TYPDADSFRAC').AsString);
   SetField ('PSA_DADSFRACTION', Q.FindField ('PSA_DADSFRACTION').AsString);
   SetField ('PSA_ALLOCFORFAIT', Q.FindField ('PSA_ALLOCFORFAIT').AsString);
   SetField ('PSA_REMBJUSTIF', Q.FindField ('PSA_REMBJUSTIF').AsString);
   SetField ('PSA_PRISECHARGE', Q.FindField ('PSA_PRISECHARGE').AsString);
   SetField ('PSA_AUTREDEPENSE', Q.FindField ('PSA_AUTREDEPENSE').AsString);
   SetField ('PSA_TYPPRUDH', Q.FindField ('PSA_TYPPRUDH').AsString);
   SetField ('PSA_PRUDHCOLL', Q.FindField ('PSA_PRUDHCOLL').AsString);
   SetField ('PSA_PRUDHSECT', Q.FindField ('PSA_PRUDHSECT').AsString);
   SetField ('PSA_PRUDHVOTE', Q.FindField ('PSA_PRUDHVOTE').AsString);
//DEBUT PT120
   SetField ('PSA_CONGESPAYES', Q.FindField ('PSA_CONGESPAYES').AsString);
   SetField ('PSA_TYPPROFILCGE', Q.FindField ('PSA_TYPPROFILCGE').AsString);
   SetField ('PSA_PROFILCGE', Q.FindField ('PSA_PROFILCGE').AsString);
   SetField ('PSA_STANDCALEND', Q.FindField ('PSA_STANDCALEND').AsString);
   SetField ('PSA_CALENDRIER', Q.FindField ('PSA_CALENDRIER').AsString);
   SetField ('PSA_TYPEDITBULCP', Q.FindField ('PSA_TYPEDITBULCP').AsString);
   SetField ('PSA_EDITBULCP', Q.FindField ('PSA_EDITBULCP').AsString);
   SetField ('PSA_CPACQUISMOIS', Q.FindField ('PSA_CPACQUISMOIS').AsString);
   SetField ('PSA_NBREACQUISCP', Q.FindField ('PSA_NBREACQUISCP').AsFloat);
   SetField ('PSA_SSDECOMPTE', Q.FindField ('PSA_SSDECOMPTE').AsString);
   SetField ('PSA_TYPNBACQUISCP', Q.FindField ('PSA_TYPNBACQUISCP').AsString);
   SetField ('PSA_NBACQUISCP', Q.FindField ('PSA_NBACQUISCP').AsString);
   SetField ('PSA_CPACQUISSUPP', Q.FindField ('PSA_CPACQUISSUPP').AsString);
   SetField ('PSA_NBRECPSUPP', Q.FindField ('PSA_NBRECPSUPP').AsFloat);
   SetField ('PSA_CPACQUISANC', Q.FindField ('PSA_CPACQUISANC').AsString);
   SetField ('PSA_BASANCCP', Q.FindField ('PSA_BASANCCP').AsString); // PT127
   SetField ('PSA_VALANCCP', Q.FindField ('PSA_VALANCCP').AsString); { PT140 }
   SetField ('PSA_DATANC', Q.FindField ('PSA_DATANC').AsString);
   SetField ('PSA_TYPDATANC', Q.FindField ('PSA_TYPDATANC').AsString);
   SetField ('PSA_DATEACQCPANC', Q.FindField ('PSA_DATEACQCPANC').AsString);
   SetField ('PSA_CPTYPEMETHOD', Q.FindField ('PSA_CPTYPEMETHOD').AsString);
   SetField ('PSA_VALORINDEMCP', Q.FindField ('PSA_VALORINDEMCP').AsString);
   SetField ('PSA_CPTYPEVALO', Q.FindField ('PSA_CPTYPEVALO').AsString);
   SetField ('PSA_MVALOMS', Q.FindField ('PSA_MVALOMS').AsString);
   SetField ('PSA_VALODXMN', Q.FindField ('PSA_VALODXMN').AsFloat);
   SetField ('PSA_CPTYPERELIQ', Q.FindField ('PSA_CPTYPERELIQ').AsString);
   SetField ('PSA_RELIQUAT', Q.FindField ('PSA_RELIQUAT').AsString);
   SetField ('PSA_TYPPAIEVALOMS', Q.FindField ('PSA_TYPPAIEVALOMS').AsString); { PT132 }
   SetField ('PSA_PAIEVALOMS', Q.FindField ('PSA_PAIEVALOMS').AsString); { PT126 }
//FIN PT120
//DEBUT PT166
   SetField ('PSA_DATELIBRE1', Q.FindField ('PSA_DATELIBRE1').AsDateTime);
   SetField ('PSA_DATELIBRE2', Q.FindField ('PSA_DATELIBRE2').AsDateTime);
   SetField ('PSA_DATELIBRE3', Q.FindField ('PSA_DATELIBRE3').AsDateTime);
   SetField ('PSA_DATELIBRE4', Q.FindField ('PSA_DATELIBRE4').AsDateTime);
   SetField ('PSA_BOOLLIBRE1', Q.FindField ('PSA_BOOLLIBRE1').AsString);
   SetField ('PSA_BOOLLIBRE2', Q.FindField ('PSA_BOOLLIBRE2').AsString);
   SetField ('PSA_BOOLLIBRE3', Q.FindField ('PSA_BOOLLIBRE3').AsString);
   SetField ('PSA_BOOLLIBRE4', Q.FindField ('PSA_BOOLLIBRE4').AsString);
   SetField ('PSA_LIBREPCMB1', Q.FindField ('PSA_LIBREPCMB1').AsString);
   SetField ('PSA_LIBREPCMB2', Q.FindField ('PSA_LIBREPCMB2').AsString);
   SetField ('PSA_LIBREPCMB3', Q.FindField ('PSA_LIBREPCMB3').AsString);
   SetField ('PSA_LIBREPCMB4', Q.FindField ('PSA_LIBREPCMB4').AsString);
//FIN PT166
   end;
Ferme (Q);
end;

procedure TOM_SALARIES.BDupliqueClick(Sender: Tobject);
begin
SalarieDupliquer:= GetField ('PSA_SALARIE');
TFFiche (Ecran).BInsert.Click;
end;
// FIN PT104
// Fonction click des boutons de sélection des onglets

{$IFNDEF GCGC}

procedure TOM_SALARIES.BTNClick(Sender: TObject);
var
  BTN: TToolBarButton97;
  Nom: string;
begin
  if not (Sender is TToolBarButton97) then exit;
  BTN := TToolBarButton97(Sender);
  Nom := 'P' + Copy(BTN.name, 4, Length(BTN.name) - 3);
  if BTN.Name = 'BHISTO' then
  begin
    RempliOnglet(BTN.Name);
    exit;
  end;
  if BTN.Down then
  begin
    SetControlProperty(Nom, 'TabVisible', TRUE);
    SetControlProperty(Nom + '_', 'TabVisible', TRUE);
    SetControlProperty(Nom, 'TabVisible', TRUE);
    RempliOnglet(BTN.Name);
    SetControlVisible('IMAGE' + BTN.name, True);
  end
  else
  begin
    SetControlProperty(Nom, 'TabVisible', FALSE);
    SetControlProperty(Nom + '_', 'TabVisible', FALSE);
    SetControlVisible('IMAGE' + BTN.name, False);
  end;
end;
{$ENDIF}

{$IFNDEF GCGC}

procedure TOM_SALARIES.RempliOnglet(Nom: string);
var
  DateDeb, DateFin: TDateTime;
  Titre: string;
begin
  if Nom = 'BTNCOMPETENCE' then
  begin
    if OngletsRibbonBar[6] = GetField('PSA_SALARIE') then Exit;
    OngletsRibbonBar[6] := GetField('PSA_SALARIE');
    CarriereAfficheCompetences(GetField('PSA_SALARIE'), GetField('PSA_LIBELLEEMPLOI'), THGrid(GetControl('GCOMPETEMPLOI')), THGrid(GetControl('GCOMPETSAL')));
    If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GCOMPETEMPLOI')));
    If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GCOMPETSAL')));
    Titre := 'Compétences - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
  end
  else
    if Nom = 'DIPLOMES' then
    begin
      if OngletsRibbonBar[4] = GetField('PSA_SALARIE') then Exit;
      OngletsRibbonBar[4] := GetField('PSA_SALARIE');
      CarriereAfficheFormations(GetField('PSA_SALARIE'), THGrid(GetControl('GINIT')), nil, nil);
      If (GetControl('HMTrad') <> Nil) And (GetControl('GINIT') <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GINIT')));
      Titre := 'Diplômes - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    end
    else
      if Nom = 'FORMATIONSINT' then
      begin
        if OngletsRibbonBar[2] = GetField('PSA_SALARIE') then Exit;
        OngletsRibbonBar[2] := GetField('PSA_SALARIE');
        CarriereAfficheFormations(GetField('PSA_SALARIE'), nil, THGrid(GetControl('GINTERNE')), nil);
        If (GetControl('HMTrad') <> Nil) And (GetControl('GINTERNE') <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GINTERNE')));
        Titre := 'Formations internes - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
      end
      else
        if Nom = 'FORMATIONSEXT' then
        begin
          if OngletsRibbonBar[3] = GetField('PSA_SALARIE') then Exit;
          OngletsRibbonBar[3] := GetField('PSA_SALARIE');
          CarriereAfficheFormations(GetField('PSA_SALARIE'), nil, nil, THGrid(GetControl('GEXTERNE')));
          If (GetControl('HMTrad') <> Nil) And (GetControl('GEXTERNE') <> Nil) Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GEXTERNE')));
          Titre := 'Formations externes - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
        end
        else
          if (Nom = 'BTNHISTO') or (Nom = 'BHISTO') then
          begin
            if IsValidDate(GetControlText('DATEHISTO')) then DateDeb := StrToDate(GetControlText('DATEHISTO'))
            else
            begin
              PGIBox(GetControlText('DATEHISTO') + ' n''est pas une date correcte', Ecran.Caption);
              Exit;
            end;
            if IsValidDate(GetControlText('DATEHISTO1')) then DateFin := StrToDate(GetControlText('DATEHISTO1'))
            else
            begin
              PGIBox(GetControlText('DATEHISTO1') + ' n''est pas une date correcte', Ecran.Caption);
              Exit;
            end;
            CarriereAfficheHisto(GetField('PSA_SALARIE'), DateDeb, DateFin, THGrid(GetControl('GHISTO')), THMultiValComboBox(GetControl('CRITERESHISTO')));
            If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GHISTO')));
          end
          else
            if Nom = 'BTNCONTRAT' then
            Begin
              CarriereAfficheContrats(GetField('PSA_SALARIE'), THGrid(GetControl('GCONTRATS')));
              If GetControl('HMTrad') <> Nil Then THSystemMenu(GetControl('HMTrad')).ResizeGridColumns(THGrid(GetControl('GCONTRATS')));
            End
//            else If Nom = 'BTNELTCOMPL' then AfficheEltCompl;
            else if Nom = 'BTNELTCOMPL' then AfficheEltNat;
  TFFiche(Ecran).Caption := Titre;
  UpdateCaption(TFFiche(Ecran));
end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.PBTNClick(Sender: TObject);
var
  LePanel: TPageControl;
  BTN: TToolBarButton97;
  Nom: string;
begin
  BTNClick(Sender); // Activation
  if not (Sender is TToolBarButton97) then exit;
  BTN := TToolBarButton97(Sender);
  Nom := 'P' + Copy(BTN.name, 4, Length(BTN.name) - 3);
  LePanel := TPageControl(GetControl('PAGES'));
  if Assigned(LePanel) and BTN.down then LePanel.Activepage := TTabSheet(GetControl(Nom))
  else LePanel.Activepage := TTabSheet(GetControl('PGeneral'));
end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

//PT113
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/10/2005
Modifié le ... :   /  /
Description .. : PersAChargeExit
Mots clefs ... : PAIE;PGSALARIE
*****************************************************************}

procedure TOM_SALARIES.PersAChargeExit(Sender: TObject);
begin
//PT119
  if Ds.State in ([DsBrowse]) then
    ds.Edit;
//FIN PT119
  if (GetControlText('EPERSACHARGE') = '') then
   SetField ('PSA_PERSACHARGE', 99)
  else
    SetField('PSA_PERSACHARGE', StrToInt(GetControlText('EPERSACHARGE')));
end;
//FIN PT113

{ DEB PT128 }

procedure TOM_SALARIES.ClickBtnBulDefaut(Sender: TObject);
begin
  if VH_Paie.PgBulDefaut <> '' then
    SetField('PSA_ETATBULLETIN', VH_Paie.PgBulDefaut)
  else
    SetField('PSA_ETATBULLETIN', 'PBP');
end;
{ FIN PT128 }

// DEB PT136

procedure TOM_SALARIES.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var LeNom: string;
begin
  LeNom := THEdit(Sender).Name;
  case Key of
    VK_DELETE:
      begin
        key := 0;
        if LeNom = 'PSA_DATESORTIE' then
        begin
          SetField('PSA_DATESORTIE', IDate1900);
          SetControlText('ZPSA_DATESORTIE', '');
          SetControlVisible('PSA_DATESORTIE', FALSE);
          SetControlVisible('ZPSA_DATESORTIE', true);
          SetFocusControl('PSA_MOTIFENTREE');
        end
        else
        begin
          if LeNom = 'PSA_DATEENTREEPREC' then
          begin
            SetField('PSA_DATEENTREEPREC', IDate1900);
            SetControlText('ZPSA_DATEENTREEPREC', '');
            SetControlVisible('PSA_DATEENTREEPREC', FALSE);
            SetControlVisible('ZPSA_DATEENTREEPREC', true);
            SetFocusControl('PSA_CATDADS');
          end
          else
          begin
            SetField('PSA_DATESORTIEPREC', IDate1900);
            SetControlText('ZPSA_DATESORTIEPREC', '');
            SetControlVisible('PSA_DATESORTIEPREC', FALSE);
            SetControlVisible('ZPSA_DATESORTIEPREC', true);
            SetFocusControl('PSA_CATDADS');
          end;
        end;
      end;
  end;
end;
// FIN PT136
// DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}

procedure TOM_SALARIES.Controle_SalarieDA;
var
  EtabSal, NomSal, PrenSal, mes: string;
  mes1, mes2, mes3: string;
  longsal: integer;
begin
  EtabSal := (GetField('PSA_ETABLISSEMENT'));
  NomSal := (GetField('PSA_LIBELLE'));
  PrenSal := (GetField('PSA_PRENOM'));

  Mes := '';
  Mes1 := '';
  Mes2 := '';
  Mes3 := '';
  longsal := length(GetField('PSA_SALARIE'));
  if (GetField('PSA_SALARIE') <> '') and (LongSal < 11) then
  begin
    if (EtabSal = '') then
    begin
      Mes1 := '#13#10    - l''établissement ';
      Mes := Mes1;
    end;
    if (NomSal = '') then
    begin
      Mes2 := '#13#10    - le nom ';
      Mes := Mes + Mes2;
    end;
    if (PrenSal = '') then
    begin
      Mes3 := '#13#10    - le prénom ';
      Mes := Mes + Mes3;
    end;
    if (PrenSal = '') then SetfocusControl('PSA_PRENOM');
    if (NomSal = '') then SetfocusControl('PSA_LIBELLE');
    if (EtabSal = '') then SetfocusControl('PSA_ETABLISSEMENT');
  end;
  if mes <> '' then
  begin
    PgiBox('Saisie obligatoire. Veuillez renseigner :' + mes + ' du salarié.', Ecran.caption);
    LastError := 1;
  end;
end;
{$ENDIF}
{$IFDEF GCGC}

procedure TOM_SALARIES.OnCancelRecord;
begin
  inherited;
  if TFfiche(Ecran) <> nil then
  begin
    TobZonesCompl.PutEcran(TFfiche(Ecran));
  end;
end;
{$ENDIF}
{$IFDEF GCGC}

procedure TOM_SALARIES.ZonesComplIsModified;
begin
  TobZonesCompl.GetEcran(TFfiche(Ecran), nil);
  if TobZonesCompl.IsOneModifie and not (DS.State in [dsInsert, dsEdit]) then
  begin
    DS.edit; // pour passer DS.state en mode dsEdit
{$IFDEF EAGLCLIENT}
    TFFiche(Ecran).QFiche.CurrentFille.Modifie := true;
{$ELSE}
    SetField('PSA_SALARIE', GetControlText('PSA_SALARIE'));
{$ENDIF}
  end;
end;
{$ENDIF}

{$IFDEF GCGC}

procedure AGLZonesComplIsModified(parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else exit;
  if (OM is TOM_SALARIES)
    then TOM_SALARIES(OM).ZonesComplIsModified
  else exit;
end;
{$ENDIF}

{$IFDEF GCGC}

function RechSalarieModuleDA(Lesalarie: string): Boolean;
begin
  Result := False;

  if (SalarieDansCirCuitDA_Salarie(Lesalarie))
    or (SalarieDansCirCuitDA_HierarchieDirecte(Lesalarie))
    or (SalarieDansCirCuitDA_HierarchieInDirecte(Lesalarie)) then
  begin
    PgiBox('Suppression impossible. Ce salarié est présent sur des circuits de validation ', ' Salariés ');
    Result := True;
  end;
  if (DA_aValiderParSalarie(Lesalarie)) then
    Result := True;
end;

function SalarieDansCirCuitDA_Salarie(LeSalarie: string): Boolean;
var
  sSql, sMessageDA: string;
  TobTypeDA: Tob;
begin
  Result := False;

  sSQL := 'SELECT DISTINCT DAV_TYPEDA,DAV_CODECIRCUIT FROM CIRCUITDA '
    + ' WHERE  (DAV_TYPECIRCUIT = "Z" AND DAV_SALARIE ="' + LeSalarie + '") ';

  sMessageDA := '';

  TobTypeDA := Tob.Create('_TYPEDA', nil, -1);
  try
    if TobTypeDA.LoadDetailDBFromSql('_TYPEDA', sSQL) then
    begin
      if TobTypeDA.Detail.Count <> 0 then
        Result := True;
    end;
  finally
    TobTypeDA.Free;
  end;
end;


function SalarieDansCirCuitDA_HierarchieDirecte(LeSalarie: string): Boolean;
var
  sSql, sMessageDA: string;
  TobTypeDA: Tob;
begin
  Result := False;

  sSQL := 'SELECT DISTINCT DAV_TYPEDA,DAV_CODECIRCUIT FROM CIRCUITDA,SERVICES '
    + ' WHERE (DAV_TYPECIRCUIT = "H" or DAV_TYPECIRCUIT = "W" or DAV_TYPECIRCUIT = "V" ) '
    + ' AND ((PGS_RESPSERVICE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSNDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_SECRETAIRE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIRENDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "S")'
    + ' Or (PGS_ADJOINTSERVICE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "A")'
    + ' Or (PGS_ADJOINTABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTNDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "A"))';

  sMessageDA := '';

  TobTypeDA := Tob.Create('_TYPEDA', nil, -1);
  try
    if TobTypeDA.LoadDetailDBFromSql('_TYPEDA', sSQL) then
    begin
      if TobTypeDA.Detail.Count <> 0 then
        Result := True;
    end;
  finally
    TobTypeDA.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CCMX-CEGID
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : Chargement des couples types de DA, services autorisés
Suite ........ : pour le User.
Mots clefs ... :
*****************************************************************}

function SalarieDansCirCuitDA_HierarchieInDirecte(LeSalarie: string): Boolean;
var
  sSql, sMessageDA: string;
  TobTypeDA: Tob;
begin
  Result := False;

  sSQL := 'SELECT DISTINCT DAV_TYPEDA,DAV_CODECIRCUIT FROM CIRCUITDA,SERVICES '
    + ' LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE '
    + ' WHERE  ((DAV_TYPECIRCUIT = "H" or DAV_TYPECIRCUIT = "W") '
    + ' AND ( PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES  '
    + ' WHERE (PGS_RESPSERVICE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_RESPONSNDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "R") '
    + ' Or (PGS_SECRETAIRE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIREFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_SECRETAIRENDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "S") '
    + ' Or (PGS_ADJOINTSERVICE="' + LeSalarie + '" AND DAV_TYPERESPONS = "S" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTABS="' + LeSalarie + '" AND DAV_TYPERESPONS = "A" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTVAR="' + LeSalarie + '" AND DAV_TYPERESPONS = "V" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTFOR="' + LeSalarie + '" AND DAV_TYPERESPONS = "F" AND DAV_FONCTIONVALDA = "A") '
    + ' Or (PGS_ADJOINTNDF="' + LeSalarie + '" AND DAV_TYPERESPONS = "N" AND DAV_FONCTIONVALDA = "A"))))';

  sMessageDA := '';

  TobTypeDA := Tob.Create('_TYPEDA', nil, -1);
  try
    if TobTypeDA.LoadDetailDBFromSql('_TYPEDA', sSQL) then
    begin
      if TobTypeDA.Detail.Count <> 0 then
        Result := True;
    end;

  finally
    TobTypeDA.Free;
  end;
end;

function DA_aValiderParSalarie(Lesalarie: string): Boolean;
var
  sSql, sMessageDA: string;
  TobHistoDA: Tob;
begin
  Result := False;

  sSQL := 'SELECT DAH_SOUCHE,DAH_NUMERO FROM HISTODA '
    + ' WHERE  (DAH_STATUTDA="" AND DAH_SALARIE ="' + LeSalarie + '")';

  sMessageDA := '';

  TobHistoDA := Tob.Create('_HISTODA', nil, -1);
  try
    if TobHistoDA.LoadDetailDBFromSql('_HISTODA', sSQL) then
    begin
      if TobHistoDA.Detail.Count <> 0 then
      begin
        PgiBox('Suppression impossible. Il existe des DA en attente de validation par ce salarié. ', ' Salariés ');
        Result := True;
      end;
    end;
  finally
    TobHistoDA.Free;
  end;
end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFNDEF GCGC}

procedure TOM_SALARIES.FicheKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case key of
    VK_F7:
      begin
        PgSalarieRub_LanceFiche('SALARIE=' + GetControlText('PSA_SALARIE'));
      end;
    VK_F10: //appel du mul recherche des documents Ged
      begin
        key := 0;
{$IFDEF EMANAGER}
		// Initialisation de la GED avant l'appel de la fenêtre
		InitializeGedFiles('', nil, gfmData, '', True);
{$ENDIF}
        if not JaiLeDroitTag(200065) then  //PT184
          LanceFiche_RechGedPaie('PAY', 'RECHGEDPAIE', '', '', 'SALARIE=' + Getfield('PSA_AUXILIAIRE')+';ACTION=CONSULTATION')
        else
          LanceFiche_RechGedPaie('PAY', 'RECHGEDPAIE', '', '', 'SALARIE=' + Getfield('PSA_AUXILIAIRE')+';ACTION=MODIFICATION');
{$IFDEF EMANAGER}
		// Terminaison de la GED
		FinalizeGedFiles();
{$ENDIF}
      end;
  end;
end;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/12/2006
Modifié le ... :   /  /
Description .. : passage des infos salariés
Mots clefs ... :
*****************************************************************}
{$IFNDEF GCGC}

function salarie_MyAfterImport(Sender: TObject): string;
var
  OM: TOM;
begin
  Result := '';
  if sender is TFFICHE then
    OM := TFFICHE(Sender).OM
  else
    exit;

  if (OM is TOM_SALARIES) then
    Result := TOM_SALARIES(OM).GetInfoSal
  else
    exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/12/2006
Modifié le ... :   /  /
Description .. : permet l'affichage de l'indicateur de document
Mots clefs ... :
*****************************************************************}

procedure Salarie_GestionBoutonGED(Sender: TObject);
var
  OM: TOM;
begin
  if sender is TFFICHE then
    OM := TFFICHE(Sender).OM
  else
    exit;

  if (OM is TOM_SALARIES) then
    TOM_SALARIES(OM).GerebGed
  else
    exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/12/2006
Modifié le ... :   /  /
Description .. : Retourne le code salarie
Mots clefs ... :
*****************************************************************}

function TOM_SALARIES.GetInfoSal: string;
var
  Salinfo: string;

begin
  Salinfo := 'SALARIE=' + Getfield('PSA_AUXILIAIRE');
  Result := Salinfo;
end;

procedure TOM_SALARIES.bGedOnClick(Sender: Tobject);
begin
{$IFDEF EMANAGER}
		// Initialisation de la GED avant l'appel de la fenêtre
		InitializeGedFiles('', nil, gfmData, '', True);
{$ENDIF}
  if not JaiLeDroitTag(200065) then  //PT184
    LanceFiche_RechGedPaie('PAY', 'RECHGEDPAIE', '', '', 'SALARIE=' + Getfield('PSA_AUXILIAIRE')+';ACTION=CONSULTATION')
  else
    LanceFiche_RechGedPaie('PAY', 'RECHGEDPAIE', '', '', 'SALARIE=' + Getfield('PSA_AUXILIAIRE')+';ACTION=MODIFICATION');
{$IFDEF EMANAGER}
		// Terminaison de la GED
		FinalizeGedFiles();
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/12/2006
Modifié le ... :   /  /
Description .. : Test présence de documents Ged pour un salarié. Dans le
Suite ........ : cas positif, on affiche le boutton pour accéder au doc.
Mots clefs ... :
*****************************************************************}

procedure TOM_SALARIES.GerebGed;
var
  bGed: TToolbarButton97;

begin
  bGed := TToolbarButton97(GetControl('BGED'));
  If BGed = Nil then exit;    //PT172
  if V_PGI.ModePCL = '1' then
  begin
    if V_PGI.RibbonBar then
    //DEB PT184
    begin
      SetControlVisible('RIB__BGED', False);
      NonGed := True;
    end
    //FIN PT184
    else bGed.Visible := False;
    exit;
  end;
  if bGed <> nil then
  begin
    // s'il y a au moins 1 document
    if ExisteSQL('SELECT 1 FROM RTDOCUMENT WHERE RTD_TIERS="' + GetField('PSA_AUXILIAIRE') + '"') then
    begin
      if V_PGI.RibbonBar then SetControlVisible('RIB__BGED', True)
      else
      begin
        bGed.Visible := True;
        bGed.OnClick := bGedOnClick;
      end;
    end;
  end;
end;
{$ENDIF}

{DEBUT PT143}
{$IFNDEF GCGC}

procedure TOM_SALARIES.AccesHistoEdit(Sender: TObject);
var Longueur: Integer;
  ChampIdemEtb, Champ, Salarie, Tablette: string;
  Ret: string;
  Modif: Boolean;
{$IFNDEF EAGLCLIENT}
  Edit: THDBEdit;
{$ELSE}
  Edit: THEdit;
{$ENDIF}
begin
  Longueur := Length(TToolBarButton97(Sender).Name);
  Champ := Copy(TToolBarButton97(Sender).Name, 2, Longueur - 1);
  ChampIdemEtb := 'PSA_TYP' + Copy(TToolBarButton97(Sender).Name, 6, Longueur - 1);
{$IFNDEF EAGLCLIENT}
  Edit := THDBEdit(GetControl(Champ));
{$ELSE}
  Edit := THEdit(GetControl(Champ));
{$ENDIF}
  Tablette := Edit.DataType;
  Modif := True;
{$IFNDEF EAGLCLIENT}
  if THDBValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ELSE}
  if THValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ENDIF EAGLCLIENT}
  begin
    if GetField(ChampIdemEtb) = 'ETB' then Modif := False;
  end;
  if TFFiche(Ecran).FTypeAction = TaConsult then Modif := False; //PT168
  Salarie := GetField('PSA_SALARIE');
  if Modif then Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODE', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + Tablette)
  else Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODE', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + Tablette + ';ACTION=CONSULTATION');
  If JaiLeDroitTag(200001) then //PT176
    SetField(Champ, Ret); //PT173
end;

procedure TOM_SALARIES.AccesHistoCombo(Sender: TObject);
var Longueur: Integer;
  Champ, ChampTyp, Salarie, Tablette, ChampIdemEtb: string;
  Ret: string;
  Modif: Boolean;
{$IFNDEF EAGLCLIENT}
  Combo: THDBValComboBox;
{$ELSE}
  Combo: THValComboBox;
{$ENDIF}
begin
  //Modif := True;
  Longueur := Length(TToolBarButton97(Sender).Name);
  Champ := Copy(TToolBarButton97(Sender).Name, 2, Longueur - 1);
  ChampIdemEtb := 'PSA_TYP' + Copy(TToolBarButton97(Sender).Name, 6, Longueur - 1);
  Modif := True;
{$IFNDEF EAGLCLIENT}
  if THDBValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ELSE}
  if THValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ENDIF EAGLCLIENT}
  begin
    if GetField(ChampIdemEtb) = 'ETB' then Modif := False;
  end;
  if Champ = 'PSA_PERIODBUL' then ChampTyp := 'PSA_TYPPERIODEBUL'
  else ChampTyp := Copy(Champ, 1, 4) + 'TYP' + Copy(Champ, 5, Longueur - 5);
{$IFNDEF EAGLCLIENT}
  Combo := THDBValComboBox(GetControl(Champ));
{$ELSE}
  Combo := THValComboBox(GetControl(Champ));
{$ENDIF}
  if TFFiche(Ecran).FTypeAction = TaConsult then Modif := False; //PT168
  Tablette := Combo.DataType;
  Salarie := GetField('PSA_SALARIE');
  if modif then Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODC', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + Tablette)
  else Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODC', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + Tablette + ';ACTION=CONSULTATION');
  If JaiLeDroitTag(200001) then //PT176
    SetField(Champ, Ret); //PT173
end;

procedure TOM_SALARIES.AccesHistoCheck(Sender: TObject);
var Longueur: Integer;
  Champ, Salarie, ChampIdemEtb: string;
  Ret: string;
  Modif: Boolean;
begin
  Longueur := Length(TToolBarButton97(Sender).Name);
  ChampIdemEtb := 'PSA_TYP' + Copy(TToolBarButton97(Sender).Name, 6, Longueur - 1);
{$IFNDEF EAGLCLIENT}
  if THDBValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ELSE}
  if THValComboBox(GetControl(ChampIdemEtb)) <> nil then
{$ENDIF EAGLCLIENT}
  begin
    if GetField(ChampIdemEtb) = 'ETB' then Exit;
  end;
  Modif := True;
  if TFFiche(Ecran).FTypeAction = TaConsult then Modif := False; //PT168
  Champ := Copy(TToolBarButton97(Sender).Name, 2, Longueur - 1);
  Salarie := GetField('PSA_SALARIE');
  if Modif then Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODH', '', '', 'SAL;' + Salarie + ';' + Champ)
  else Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODH', '', '', 'SAL;' + Salarie + ';' + Champ + ';;ACTION=CONSULTATION');
  If JaiLeDroitTag(200001) then //PT176
    SetField(Champ, Ret); //PT173
end;

(*procedure TOM_SALARIES.HistorisationDonnees;
var TChamps, TSal, TS: Tob;
  Q, QSal: TQuery;
  i: Integer;
  Salarie, Champ, Valeur, Tablette, TypeValeur, AncValeur, Ident: string;
begin
  Salarie := GetField('PSA_SALARIE');
  Q := OpenSQL('SELECT * FROM PAIEPARIM  WHERE PAI_LIBELLE<>"" AND PAI_PREFIX="PSA"', True);
  TChamps := Tob.Create('LesChamps', nil, -1);
  TChamps.LoadDetailDB('LesChamps', '', '', Q, False);
  Ferme(Q);
  QSal := OpenSQL('SELECT * FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"', True);
  TSal := Tob.Create('LeSalarie', nil, -1);
  TSal.LoadDetailDB('LeSalarie', '', '', QSal, False);
  Ferme(QSal);
  TS := TSal.FindFirst([''], [''], False);
  for i := 0 to TChamps.Detail.Count - 1 do
  begin
    Tablette := TChamps.Detail[i].getValue('PAI_LIENASSOC');
    TypeValeur := TChamps.Detail[i].getValue('PAI_LETYPE');
    Ident := TChamps.Detail[i].getString('PAI_IDENT');
    Champ := TChamps.Detail[i].getValue('PAI_PREFIX') + '_' + TChamps.Detail[i].getValue('PAI_SUFFIX');
    if TypeValeur = 'D' then Valeur := DateToStr(GetField(Champ))
    else if TypeValeur = 'F' then Valeur := FloatToStr(GetField(Champ))
    else if TypeValeur = 'I' then Valeur := IntToStr(GetField(Champ))
    else Valeur := GetField(Champ);
    if TS = nil then CreerHisto(Salarie, V_PGI.DateEntree, '001', Ident, Valeur, '', Tablette, TypeValeur)
    else
    begin
      AncValeur := TS.getString(Champ);
      if AncValeur <> Valeur then CreerHisto(Salarie, V_PGI.DateEntree, '002', Ident, Valeur, AncValeur, Tablette, TypeValeur);
    end;
  end;
  TSal.Free;
  TChamps.Free;
end;*)

function TOM_SALARIES.CreerBoutonHisto(THBST: TTabsheet; Nom: string; L, W, T: Integer): TToolBarButton97;
var
  Bt: TToolBarButton97;
  Nb: string;
  Numero: Integer;
begin
  if (Copy(Nom, 1, 8) = 'PSA_TRAV') then
  begin
    Nb := copy(Nom, 13, 1);
    Numero := StrtoInt(Nb);
    if (Numero > VH_Paie.PGNbreStatOrg) then
    begin
      result := nil;
      exit;
    end;
  end;
  Bt := TToolBarButton97.Create(THBST);
  Bt.Parent := THBST;
  Bt.Left := L;
  Bt.Width := W;
  Bt.Top := T;
  Bt.Name := 'B' + Nom;
  Bt.Caption := '';
  Bt.Visible := True;
  Result := Bt;
end;

procedure TOM_SALARIES.RendAccesChampHisto(Acces, Maj: Boolean);
var Q: TQuery;
  i: Integer;
  TobParamSal: Tob;
  LeChamp, TypeDonnee, Tablette: string;
  UnControl: TComponent;
{$IFNDEF EAGLCLIENT}
  Edit: THDBEdit;
  Check: THDBcheckBox;
  MCombo: THDBValComboBox;
{$ELSE}
  Edit: THEdit;
  Check: THcheckBox;
  MCombo: THValComboBox;
{$ENDIF}
begin
  Q := OpenSQL('SELECT DISTINCT PPP_PGINFOSMODIF,PPP_LIENASSOC,PPP_PGTYPEDONNE FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_PGTYPEINFOLS="SAL" AND PPP_HISTORIQUE="X"', True);
  TobParamSal := Tob.Create('parametrage salarie', nil, -1);
  TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
  ferme(Q);
  for i := 0 to TobParamSal.Detail.Count - 1 do
  begin
    LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
    if LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
    Tablette := TobParamSal.Detail[i].GetValue('PPP_LIENASSOC');
    TypeDonnee := TobParamSal.Detail[i].GetValue('PPP_PGTYPEDONNE');
    UnControl := Ecran.FindComponent(LeChamp);
{$IFNDEF EAGLCLIENT}
    if UnControl is THDBEdit then
    begin
      Edit := THDBEdit(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        Edit.Color := ClWindow;
        Edit.Enabled := True;
      end
      else
      begin
        Edit.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        Edit.Enabled := False;
      end;
    end;
    if UnControl is THDBValComboBox then
    begin
      MCombo := THDBValComboBox(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        MCombo.Color := ClWindow;
        MCombo.Enabled := True;
      end
      else
      begin
        MCombo.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        MCombo.Enabled := False;
      end;
    end;
    if UnControl is THDBcheckBox then
    begin
      Check := THDBcheckBox(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        Check.Color := ClWindow;
        Check.Enabled := True;
      end
      else
      begin
        Check.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        Check.Enabled := False;
      end;
    end;
{$ELSE}
    if UnControl is THEdit then
    begin
      Edit := THEdit(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        Edit.Color := ClWindow;
        Edit.Enabled := True;
      end
      else
      begin
        Edit.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        Edit.Enabled := False;
      end;
    end;
    if UnControl is THValComboBox then
    begin
      MCombo := THValComboBox(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        MCombo.Color := ClWindow;
        MCombo.Enabled := True;
      end
      else
      begin
        MCombo.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        MCombo.Enabled := False;
      end;
    end;
    if UnControl is THcheckBox then
    begin
      Check := THcheckBox(getControl(LeChamp));
      if Acces then
      begin
        SetControlVisible('B' + LeChamp, False);
        Check.Color := ClWindow;
        Check.Enabled := True;
      end
      else
      begin
        Check.Color := CouleurHisto;
        SetControlVisible('B' + LeChamp, True);
        Check.Enabled := False;
      end;
    end;
{$ENDIF}
//    SetControlEnabled(LeChamp,Acces);
    if Maj then
    begin
//PT214
      if TypeDonnee = 'I' then CreerHisto(GetField('PSA_SALARIE'), GetField('PSA_DATEENTREE'), '001', LeChamp, IntToStr(GetField(LeChamp)), '', Tablette, TypeDonnee, '-')
      else if TypeDonnee = 'F' then CreerHisto(GetField('PSA_SALARIE'), GetField('PSA_DATEENTREE'), '001', LeChamp, FloatToStr(GetField(LeChamp)), '', Tablette, TypeDonnee, '-')
      else if TypeDonnee = 'D' then CreerHisto(GetField('PSA_SALARIE'), GetField('PSA_DATEENTREE'), '001', LeChamp, DateToStr(GetField(LeChamp)), '', Tablette, TypeDonnee, '-')
      else CreerHisto(GetField('PSA_SALARIE'), GetField('PSA_DATEENTREE'), '001', LeChamp, GetField(LeChamp), '', Tablette, TypeDonnee, '-');
//FIN PT214
    end;
  end;
  TobParamSal.Free;
end;
{FIN PT143}

procedure TOM_SALARIES.AfficheEltNat(AjoutElt: Boolean = False);
var Q: TQuery;
  TobParamSal, TobGrille, TG : Tob;
  i: Integer;
  Libelle : string;
  GrilleEC: THGrid;
  Theme, ThemePrec, Niveau, Img, Element, NiveauMax : string;
  DateApplic: TDateTime;
  Montant: Double;
begin
  if not AjoutElt then
  begin
    if OngletsRibbonBar[1] = GetField('PSA_SALARIE') then Exit;
  end;
  OngletsRibbonBar[1] := GetField('PSA_SALARIE');
  GrilleEC := THGrid(GetControl('GELTNATDOS'));
  for i := 1 to GrilleEC.RowCount do
  begin
    GrilleEC.Rows[i].Clear;
  end;
  GrilleEC.RowCount := 2;
  GrilleEC.FixedRows := 1;
  Theme := '';
  ThemePrec := '';
  TobGrille := Tob.Create('remplir grille', nil, -1);
  Q := OpenSQL('SELECT PED_LIBELLE,PED_CODEELT FROM ELTNATIONDOS' +
    ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '"' + // AND PED_DATEVALIDITE<="' + UsdateTime(V_PGI.DateEntree) + '"' + //PT174
    ' GROUP BY PED_LIBELLE,PED_CODEELT ORDER BY PED_LIBELLE', True);
  TobParamSal := Tob.Create('parametrage element', nil, -1);
  TobParamSal.LoadDetailDB('Parametrageelt', '', '', Q, FALSE, False);
  Ferme(Q);
  for i := 0 to TobParamSal.Detail.Count - 1 do
  begin
    Libelle := TobParamSal.Detail[i].Getvalue('PED_LIBELLE');
    Element := TobParamSal.Detail[i].GeTvalue('PED_CODEELT');
    //DEB PT174
    if ExisteSQL('SELECT PED_MONTANTEURO FROM ELTNATIONDOS' +
      ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE<="' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Element + '"') then
    begin
      Q := OpenSQL('SELECT ##TOP 1## PED_MONTANTEURO,PED_DATEVALIDITE FROM ELTNATIONDOS' +
        ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE<="' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Element + '"' +
        ' ORDER BY PED_DATEVALIDITE DESC', True);
    end
    else
    begin
      Q := OpenSQL('SELECT ##TOP 1## PED_MONTANTEURO,PED_DATEVALIDITE FROM ELTNATIONDOS' +
        ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE>"' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Element + '"' +
        ' ORDER BY PED_DATEVALIDITE ASC', True);
    end;
    //FIN PT174
    if not Q.Eof then
    begin
      Montant := Q.FindField('PED_MONTANTEURO').AsFloat;
      DateApplic := Q.FindField('PED_DATEVALIDITE').AsFloat;
    end
    else
    begin
      Montant := 0;
      DateApplic := IDate1900;
    end;
    Ferme(Q);
    Q := OpenSQL('SELECT PNR_TYPENIVEAU,PNR_NIVMAXPERSO FROM ELTNIVEAUREQUIS WHERE ##PNR_PREDEFINI## PNR_CODEELT="' + Element + '" ORDER BY PNR_PREDEFINI DESC', True); //PT165
    if not Q.Eof then
    begin
      Niveau := Q.FindField('PNR_TYPENIVEAU').AsString;
      NiveauMax := Q.FindField('PNR_NIVMAXPERSO').AsString;
    end;
    Ferme(Q);
    if Niveau = 'SAL' then Img := '#ICO#44'
    else Img := '#ICO#91';
    TG := Tob.Create('ligne grille', TobGrille, -1);
    TG.AddChampSupValeur('NIVEAU', 'Salarié');
    TG.AddChampSupValeur('ZONE', Libelle);
    TG.AddChampSupValeur('VALEURAFFICHE', FloatToStr(Montant));
    TG.AddChampSupValeur('DD', DateApplic);
    TG.AddChampSupValeur('ETAT', Img);
    TG.AddChampSupValeur('LETYPE', 'F');
    TG.AddChampSupValeur('CHAMP', Element);
    TG.AddChampSupValeur('THEME', '');
  end;
  TobGrille.Detail.Sort('THEME');
  if TobGrille.Detail.Count > 0 then GrilleEC.RowCount := TobGrille.Detail.Count + 1
  else GrilleEC.RowCount := 2;
  GrilleEC.ColCount := 7;
  GrilleEC.ColWidths[0] := -1;
  GrilleEC.ColWidths[1] := 200;
  GrilleEC.ColWidths[2] := 200;
  GrilleEC.ColWidths[3] := 150;
  GrilleEC.ColWidths[4] := -1;
  GrilleEC.ColWidths[5] := -1;
  GrilleEC.ColWidths[6] := -1;
  GrilleEC.FixedCols := 1;
  GrilleEC.ColTypes[2] := 'R';
  GrilleEC.ColFormats[2] := '# ##0.00';
  GrilleEC.ColAligns[2] := taRightJustify;
  GrilleEC.ColAligns[3] := taCenter;
  GrilleEC.ColFormats[3] := ShortDateFormat;
  if TobGrille.Detail.Count > 0 then TobGrille.PutGridDetail(GrilleEC, False, False, 'NIVEAU;ZONE;VALEURAFFICHE;DD;ETAT;LETYPE;CHAMP', False)
  else
  begin
    GrilleEC.cellValues[1, 1] := '<<Aucun élément>>';
  end;
  TFFiche(Ecran).HMTrad.ResizeGridColumns(GrilleEC);
  TobGrille.Free;
  TobParamSal.free;
  TFFiche(Ecran).Caption := 'Eléments nationaux - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
  UpdateCaption(TFFiche(Ecran));
end;

procedure TOM_SALARIES.AfficheEltCompl(ForcerModif: Boolean = False);
var Q: TQuery;
  TobParamSal, TobGrille, TG: Tob;
  i: Integer;
  Libelle, LeChamp, Valeur, Letype: string;
  GrilleEC: THGrid;
  WhereLib, WhereTheme: string;
  Theme, ThemePrec, Niveau, Img, Tablette: string;
  DateApplic: TDateTime;
begin
  DateApplic:= iDate1900;
  if not ForcerModif then
  begin
    if (OngletsRibbonBar[0] = GetField('PSA_SALARIE')) and (GetField('PSA_SALARIE') <> '') then Exit;
    OngletsRibbonBar[0] := GetField('PSA_SALARIE');
  end;
  SetControlVisible('GRBXCIRTELTCOMPL', True);
  if GetControlText('CTHEMEELTCOMPL') <> '' then WhereTheme := ' AND PPP_PGTHEMESALARIE="' + GetControlText('CTHEMEELTCOMPL') + '"'
  else Wheretheme := '';
  if GetControlText('ELIBELTCOMPL') <> '' then WhereLib := ' AND PPP_LIBELLE LIKE "%' + GetControlText('ELIBELTCOMPL') + '%"'
  else WhereLib := '';
  GrilleEC := THGrid(GetControl('GELTCOMPL'));
  for i := 1 to GrilleEC.RowCount do
  begin
    GrilleEC.Rows[i].Clear;
  end;
  GrilleEC.RowCount := 2;
  GrilleEC.FixedRows := 1;
  if GetControlText('CZONELIBRE') = 'OUI' then
  begin
    Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_TYPENIVEAU="SAL" AND PPP_PGTYPEINFOLS="ZLS"' + WhereTheme + WhereLib +
      ' AND PPP_PGINFOSMODIF IN (SELECT PHD_PGINFOSMODIF FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
      ') ORDER BY PPP_PGTHEMESALARIE', True); //PT204
//PT174      'AND PHD_DATEAPPLIC<="' + UsDateTime(V_PGI.DateEntree) + '") ORDER BY PPP_PGTHEMESALARIE', True);
    TobParamSal := Tob.Create('parametrage salarie', nil, -1);
    TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
    ferme(Q);
    Theme := '';
    ThemePrec := '';
    TobGrille := Tob.Create('remplir grille', nil, -1);
    for i := 0 to TobParamSal.Detail.Count - 1 do
    begin
      Theme := TobParamSal.Detail[i].GetValue('PPP_PGTHEMESALARIE');
      Niveau := TobParamSal.Detail[i].GetValue('PPP_TYPENIVEAU');
      Tablette := TobParamSal.Detail[i].GetValue('PPP_CODTABL');
      if (Theme = '') and (i = 0) then
      begin
        TG := Tob.Create('ligne grille', TobGrille, -1);
        TG.AddChampSupValeur('NIVEAU', 'Aucun thème');
        TG.AddChampSupValeur('ZONE', 'Aucun thème');
        TG.AddChampSupValeur('VALEURAFFICHE', '');
        TG.AddChampSupValeur('DD', '');
        TG.AddChampSupValeur('LETYPE', 'THEME');
        TG.AddChampSupValeur('CHAMP', '');
      end
      else if (Theme <> '') and (Theme <> ThemePrec) then
      begin
        TG := Tob.Create('ligne grille', TobGrille, -1);
        TG.AddChampSupValeur('NIVEAU', RechDom('PGTHEMESAL', Theme, False));
        TG.AddChampSupValeur('ZONE', RechDom('PGTHEMESAL', Theme, False));
        TG.AddChampSupValeur('VALEURAFFICHE', '');
        TG.AddChampSupValeur('DD', '');
        TG.AddChampSupValeur('ETAT', '');
        TG.AddChampSupValeur('LETYPE', 'THEME');
        TG.AddChampSupValeur('CHAMP', '');
      end;
      LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
      if LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
      Libelle := TobParamSal.Detail[i].GetValue('PPP_LIBELLE');
      Letype := TobParamSal.Detail[i].GetValue('PPP_PGTYPEDONNE');
      Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
        'AND PHD_PGINFOSMODIF="' + LeChamp + '" AND PHD_DATEAPPLIC<="' + UsDateTime(V_PGI.DateEntree) + '" ORDER BY PHD_DATEAPPLIC DESC', True);
      if not Q.Eof then
      begin
        Q.First;
        Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
        DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
        if Niveau <> 'SAL' then Img := '#ICO#91'
        else Img := '#ICO#44';
      end
      else
      begin
        //DEB PT174
        Ferme(Q);
        Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
          'AND PHD_PGINFOSMODIF="' + LeChamp + '" AND PHD_DATEAPPLIC>="' + UsDateTime(V_PGI.DateEntree) + '" ORDER BY PHD_DATEAPPLIC ASC', True);
//        Valeur := 'Aucune valeur';
//        Img := '#ICO#43';
        if not Q.Eof then
        begin
          Q.First;
          Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
          DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
          if Niveau <> 'SAL' then Img := '#ICO#91'
          else Img := '#ICO#44';
        end;
        //FIN PT174
      end;
      Ferme(Q);
      TG := Tob.Create('ligne grille', TobGrille, -1);
      TG.AddChampSupValeur('NIVEAU', RechDom('PGNIVEAUAVDOS', Niveau, False));
      TG.AddChampSupValeur('ZONE', Libelle);
      if LeType = 'T' then
      begin
        //PT197
        Valeur := GetLibelleZLTableDyna(DateApplic, Tablette, GetField('PSA_SALARIE'), Valeur);
{//PT197
        Q := OpenSQL('SELECT ##TOP 1## PTD_LIBELLECODE FROM TABLEDIMDET WHERE PTD_DTVALID<="' + UsDateTime(DateApplic) + '"' +
          ' AND PTD_CODTABL="' + Tablette + '" AND PTD_VALCRIT1="' + Valeur + '" ORDER BY PTD_DTVALID DESC', True);
        if not Q.Eof then Valeur := Q.FindField('PTD_LIBELLECODE').AsString;
        Ferme(Q);
}
      end;
      TG.AddChampSupValeur('VALEURAFFICHE', Valeur);
      TG.AddChampSupValeur('DD', DateApplic);
      TG.AddChampSupValeur('ETAT', Img);
      TG.AddChampSupValeur('LETYPE', Letype);
      TG.AddChampSupValeur('CHAMP', LeChamp);
      ThemePrec := Theme;
    end;
  end
  else
  begin
    Q := OpenSQL('SELECT * FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_TYPENIVEAU="SAL" AND PPP_PGTYPEINFOLS="ZLS"' + WhereTheme + WhereLib + ' ' +
      'AND PPP_PGINFOSMODIF NOT IN (SELECT PHD_PGINFOSMODIF FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
      ') ORDER BY PPP_PGTHEMESALARIE', True); //PT204
//PT174      'AND PHD_DATEAPPLIC<="' + UsDateTime(V_PGI.DateEntree) + '") ORDER BY PPP_PGTHEMESALARIE', True);
    TobParamSal := Tob.Create('parametrage salarie', nil, -1);
    TobParamSal.LoadDetailDB('Parametrage', '', '', Q, FALSE, False);
    ferme(Q);
    Theme := '';
    ThemePrec := '';
    TobGrille := Tob.Create('remplir grille', nil, -1);
    for i := 0 to TobParamSal.Detail.Count - 1 do
    begin
      Theme := TobParamSal.Detail[i].GetValue('PPP_PGTHEMESALARIE');
      Niveau := TobParamSal.Detail[i].GetValue('PPP_TYPENIVEAU');
      Tablette := TobParamSal.Detail[i].GetValue('PPP_CODTABL');
      if (Theme = '') and (i = 0) then
      begin
        TG := Tob.Create('ligne grille', TobGrille, -1);
        TG.AddChampSupValeur('NIVEAU', 'Aucun thème');
        TG.AddChampSupValeur('ZONE', 'Aucun thème');
        TG.AddChampSupValeur('VALEURAFFICHE', '');
        TG.AddChampSupValeur('DD', '');
        TG.AddChampSupValeur('LETYPE', 'THEME');
        TG.AddChampSupValeur('CHAMP', '');
      end
      else if (Theme <> '') and (Theme <> ThemePrec) then
      begin
        TG := Tob.Create('ligne grille', TobGrille, -1);
        TG.AddChampSupValeur('NIVEAU', '');
        TG.AddChampSupValeur('ZONE', RechDom('PGTHEMESAL', Theme, False));
        TG.AddChampSupValeur('VALEURAFFICHE', '');
        TG.AddChampSupValeur('DD', '');
        TG.AddChampSupValeur('ETAT', '');
        TG.AddChampSupValeur('LETYPE', 'THEME');
        TG.AddChampSupValeur('CHAMP', '');
      end;
      LeChamp := TobParamSal.Detail[i].GetValue('PPP_PGINFOSMODIF');
      if LeChamp = 'PSA_SITUATIONFAMI' then LeChamp := 'PSA_SITUATIONFAMIL';
      Libelle := TobParamSal.Detail[i].GetValue('PPP_LIBELLE');
      Letype := TobParamSal.Detail[i].GetValue('PPP_PGTYPEDONNE');
      Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
        'AND PHD_PGINFOSMODIF="' + LeChamp + '" AND PHD_DATEAPPLIC<="' + UsDateTime(V_PGI.DateEntree) + '" ORDER BY PHD_DATEAPPLIC DESC', True);
      if not Q.Eof then
      begin
        Q.First;
        Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
        DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
        if Niveau <> 'SAL' then Img := '#ICO#91'
        else Img := '#ICO#44';
      end
      else
      begin
        Valeur := 'Aucune valeur';
        Img := '#ICO#43';
      end;
      Ferme(Q);
      TG := Tob.Create('ligne grille', TobGrille, -1);
      TG.AddChampSupValeur('NIVEAU', RechDom('PGNIVEAUAVDOS', Niveau, False));
      TG.AddChampSupValeur('ZONE', Libelle);
      if LeType = 'T' then
      begin
        //PT197
        Valeur := GetLibelleZLTableDyna(DateApplic, Tablette, GetField('PSA_SALARIE'), Valeur);
{//PT197
        Q := OpenSQL('SELECT ##TOP 1## PTD_LIBELLECODE FROM TABLEDIMDET WHERE PTD_DTVALID<="' + UsDateTime(DateApplic) + '"' +
          ' AND PTD_CODTABL="' + Tablette + '" AND PTD_VALCRIT1="' + Valeur + '" ORDER BY PTD_DTVALID DESC', True);
        if not Q.Eof then Valeur := Q.FindField('PTD_LIBELLECODE').AsString;
        Ferme(Q);
}
      end;
      TG.AddChampSupValeur('VALEURAFFICHE', Valeur);
      TG.AddChampSupValeur('DD', DateApplic);
      TG.AddChampSupValeur('ETAT', Img);
      TG.AddChampSupValeur('LETYPE', Letype);
      TG.AddChampSupValeur('CHAMP', LeChamp);
      ThemePrec := Theme;
    end;

  end;
  if TobGrille.Detail.Count > 0 then GrilleEC.RowCount := TobGrille.Detail.Count + 1
  else GrilleEC.RowCount := 2;
  GrilleEC.ColCount := 7;
  GrilleEC.ColWidths[0] := -1;
  GrilleEC.ColWidths[1] := 200;
  GrilleEC.ColWidths[2] := 200;
  GrilleEC.ColWidths[3] := 150;
  GrilleEC.ColWidths[4] := -1;
  GrilleEC.ColWidths[5] := -1;
  GrilleEC.ColWidths[6] := -1;
  GrilleEC.FixedCols := 1;
  GrilleEC.ColFormats[2] := '';
  GrilleEC.ColAligns[2] := taCenter;
  GrilleEC.ColAligns[3] := taCenter;
  GrilleEC.ColFormats[3] := ShortDateFormat;
  if TobGrille.Detail.Count > 0 then TobGrille.PutGridDetail(GrilleEC, False, False, 'NIVEAU;ZONE;VALEURAFFICHE;DD;ETAT;LETYPE;CHAMP', False)
  else
  begin
    GrilleEC.cellValues[1, 1] := '<<Aucun élément>>';
  end;
  TFFiche(Ecran).HMTrad.ResizeGridColumns(GrilleEC);
  TobGrille.Free;
  TobParamSal.free;
  TFFiche(Ecran).Caption := 'Eléments dynamiques - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
  //DEB PT178
  if GetField('PSA_SALARIE') = '' then
  begin
    SetControlEnabled('CZONELIBRE',false);
    PGIBox('Saisie obligatoire. Veuillez renseigner le matricule du salarié pour accéder à la saisie des éléments dynamiques.',Ecran.caption)
  end
  else
    SetControlEnabled('CZONELIBRE',true);
  //FIN PT178
  UpdateCaption(TFFiche(Ecran));
end;


procedure TOM_SALARIES.GrilleEltComplClick(Sender: TObject);
var GrilleEC: THGrid;
  Libelle, Tablette, Salarie, Ret, LeType, Champ: string;
  Ligne: Integer;
  Img, Valeur, Niveau: string;
  DateApplic: TDateTime;
  Q: TQuery;
begin
  if GetField('PSA_SALARIE') = '' then exit; //PT178
  GrilleEC := THGrid(GetControl('GELTCOMPL'));
  Ligne := GrilleEC.Row;
  Tablette := '';
  LeType := GrilleEC.CellValues[5, Ligne];
  if LeType = 'THEME' then exit;
  Champ := GrilleEC.CellValues[6, Ligne];
  Libelle := GrilleEC.CellValues[1, Ligne];
  Salarie := GetField('PSA_SALARIE');
  if LeType = 'F' then
    //PT176
    If JaiLeDroitTag(200001) then
      Ret := AGLLanceFiche('PAY', 'ELTNATDOSLISTE', '', '', ';SAL;' + GetField('PSA_SALARIE') + ';' + Champ + ';01/01/1900;' + Libelle )
    else
      Ret := AGLLanceFiche('PAY', 'ELTNATDOSLISTE', '', '', ';SAL;' + GetField('PSA_SALARIE') + ';' + Champ + ';01/01/1900;' + Libelle + ';ACTION=CONSULTATION')
  else if LeType = 'B' then Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODH', '', '', 'SAL;' + Salarie + ';' + Champ)
  else if (LeType = 'C') or (LeType = 'T') then
    //PT176
    If JaiLeDroitTag(200001) then
      Ret := AGLLanceFiche('PAY', 'PGLISTEHITODCL', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + 'PGCOMBOZONELIBRE')
    else
      Ret := AGLLanceFiche('PAY', 'PGLISTEHITODCL', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + 'PGCOMBOZONELIBRE;ACTION=CONSULTATION')
  else Ret := AGLLanceFiche('PAY', 'PGLISTEHISTODE', '', '', 'SAL;' + Salarie + ';' + Champ + ';' + Tablette);
  if Ret <> '' then
  begin
    Q := OpenSQL('SELECT * FROM PGHISTODETAIL WHERE PHD_SALARIE="' + GetField('PSA_SALARIE') + '" ' +
      'AND PHD_PGINFOSMODIF="' + GrilleEC.CellValues[6, Ligne] + '" AND PHD_DATEAPPLIC<="' + UsDateTime(V_PGI.DateEntree) + '" ORDER BY PHD_DATEAPPLIC DESC', True);
    Valeur := Q.FindField('PHD_NEWVALEUR').AsString;
    DateApplic := Q.FindField('PHD_DATEAPPLIC').AsDateTime;
    Ferme(Q);
    Niveau := GrilleEC.CellValues[0, Ligne];
    if Niveau <> 'Salarié' then Img := '#ICO#91'
    else Img := '#ICO#44';
    GrilleEC.CellValues[2, Ligne] := Valeur;
    GrilleEC.CellValues[3, Ligne] := DateToStr(DateApplic);
    GrilleEC.CellValues[4, Ligne] := Img;
  end;
end;

procedure TOM_SALARIES.GrilleEltNatDos(Sender: TObject);
var GrilleEC: THGrid;
  Libelle, Tablette, Salarie, Ret, LeType, Champ: string;
  Ligne: Integer;
  Valeur : string;
  DateApplic: TDateTime;
  Q: TQuery;
begin
  GrilleEC := THGrid(GetControl('GELTNATDOS'));
  Ligne := GrilleEC.Row;
  Tablette := '';
  LeType := GrilleEC.CellValues[5, Ligne];
  if LeType = 'THEME' then exit;
  Champ := GrilleEC.CellValues[6, Ligne];
  Libelle := GrilleEC.CellValues[1, Ligne];
  Salarie := GetField('PSA_SALARIE');
  if Champ = '' then exit;
  //DEB PT174
  //PT176
  If JaiLeDroitTag(200001) then
    Ret := AGLLanceFiche('PAY', 'ELTNATDOSLISTE', '', '', ';SAL;' + Salarie + ';' + Champ + ';01/01/1900;' + Libelle)
  else
    Ret := AGLLanceFiche('PAY', 'ELTNATDOSLISTE', '', '', ';SAL;' + Salarie + ';' + Champ + ';01/01/1900;' + Libelle+ ';ACTION=CONSULTATION');
  if ExisteSQL('SELECT PED_MONTANTEURO FROM ELTNATIONDOS' +
    ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE<="' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Champ + '"') then
  begin
    Q := OpenSQL('SELECT ##TOP 1## PED_MONTANTEURO,PED_DATEVALIDITE FROM ELTNATIONDOS' +
      ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE<="' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Champ + '"' +
      ' ORDER BY PED_DATEVALIDITE DESC', True);
  end
  else
  begin
    Q := OpenSQL('SELECT ##TOP 1## PED_MONTANTEURO,PED_DATEVALIDITE FROM ELTNATIONDOS' +
      ' WHERE PED_TYPENIVEAU="SAL" AND PED_VALEURNIVEAU="' + GetField('PSA_SALARIE') + '" AND PED_DATEVALIDITE>"' + UsdateTime(V_PGI.DateEntree) + '" AND PED_CODEELT="' + Champ + '"' +
      ' ORDER BY PED_DATEVALIDITE ASC', True);
  end;
  //FIN PT174
  if not Q.Eof then
  begin
    Valeur := FloatToStr(Q.FindField('PED_MONTANTEURO').AsFloat);
    DateApplic := Q.FindField('PED_DATEVALIDITE').AsDateTime;
    GrilleEC.CellValues[2, Ligne] := Valeur;
    GrilleEC.CellValues[3, Ligne] := DateToStr(DateApplic);
  end
  else GrilleEC.DeleteRow(Ligne);
  Ferme(Q);
end;

procedure TOM_SALARIES.GrilleGetCellCanvas(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
var GrilleEC: THGrid;
begin
  GrilleEC := THGrid(GetControl('GELTCOMPL'));
  if (GrilleEC.CellValues[5, ARow] = 'THEME') and not (GrilleEC.IsSelected(ARow)) then
  begin
    Canvas.Font.Style := [fsBold];
    Canvas.Brush.Color := COuleurHisto;
  end;
end;

procedure TOM_SALARIES.ChangeComboEltComp(Sender: Tobject);
begin
  AfficheEltCompl(True);
end;

procedure TOM_SALARIES.ClickBtEltCompl(Sender: Tobject);
begin
  AfficheEltCompl(True);
end;

procedure TOM_SALARIES.AccesMulsZonesLibres(Sender: Tobject);
var Theme: string;
begin
  Theme := Copy(TToolBarButton97(Sender).Name, 4, length(TToolBarButton97(Sender).Name));
  AglLanceFiche('PAY', 'SAISIEZLNIVEAU', '', '', 'SAL;' + Getfield('PSA_SALARIE') + ';' + Theme);
end;

procedure TOM_SALARIES.AccesRibbonBar(Sender: Tobject);
var Cancel: Boolean;
begin
  BeforeExec(TMenuItem(Sender).Name, Cancel);
end;

procedure TOM_SALARIES.BeforeExec(const ControlName: string; var Cancel: boolean);
var Q: TQuery;
Salarie: string;
Existe : Boolean;
begin
  if V_PGI.RibbonBar then
  begin
    if (ControlName = 'RIB__BELTNATDOS') or (ControlName = 'RIB__BELTDOSPLUS') then
    begin
      //PT176
      If JaiLeDroitTag(200001) and JaiLeDroitTag(200060) then //PT184
        SetControlVisible('RIB__BELTDOSPLUS', True)
    end
    else SetControlVisible('RIB__BELTDOSPLUS', False);
  end;
  Salarie := GetField('PSA_SALARIE');
  if ControlName = 'RIB__BCOMPTEURSCP' then
  begin
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BIDENTITE' then
  begin
    TFFiche(Ecran).Caption := 'Identité - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BETATCIVIL' then
  begin
    TFFiche(Ecran).Caption := 'Etat civil - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BEMPLOI' then
  begin
    TFFiche(Ecran).Caption := 'Emploi - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BAFFECTATION' then
  begin
    TFFiche(Ecran).Caption := 'Affectation - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BPROFILS' then
  begin
    TFFiche(Ecran).Caption := 'Profils - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BAUTREPROFILS' then
  begin
    TFFiche(Ecran).Caption := 'Autres profils - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BCONTRATSALAIRES' then
  begin
    TFFiche(Ecran).Caption := 'Contrat - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BDADS' then
  begin
    TFFiche(Ecran).Caption := 'DADS - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BZONESLI' then
  begin
    TFFiche(Ecran).Caption := 'Zones libres - Salarié : ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
    UpdateCaption(TFFiche(Ecran));
    cancel := False;
    exit;
  end
  else if ControlName = 'RIB__BPERSACHARGE' then
  begin
    BEnfantsClick(TToolBarButton97(GetControl('BENFANT')));
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BMULTIEMP' then
  begin
    AGLLanceFiche('PAY', 'MULTIEMPLOYEUR', GetField('PSA_SALARIE'), '', GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM'));
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BBDETAILCP' then
  begin
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BMSA' then
  begin
    Q := OpenSQL('SELECT PSE_MSA FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('PSE_MSA').AsString = 'X' then
        if not JaiLeDroitTag(200070) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';MSA;;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';MSA;;ACTION=MODIFICATION')
      else
        if not JaiLeDroitTag(200070) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';MSA;NOUVEAU;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';MSA;NOUVEAU;ACTION=MODIFICATION');
    end
    else
      if JaiLeDroitTag(200070) then   //PT184
        AGLLanceFiche('PAY', 'COMPLSALARIE', '', '', Salarie + ';MSA;;ACTION=CREATION');
    Ferme(Q);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BBTP' then
  begin
    Q := OpenSQL('SELECT PSE_BTP FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('PSE_BTP').AsString = 'X' then
        if not JaiLeDroitTag(200071) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';BTP;;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';BTP;;ACTION=MODIFICATION')
      else
        if not JaiLeDroitTag(200071) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';BTP;NOUVEAU;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';BTP;NOUVEAU;ACTION=MODIFICATION');
    end
    else
      if JaiLeDroitTag(200071) then //PT184
        AGLLanceFiche('PAY', 'COMPLSALARIE', '', '', Salarie + ';BTP;;ACTION=CREATION');
    Ferme(Q);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BINTERMITTENTS' then
  begin
    Q := OpenSQL('SELECT PSE_INTERMITTENT FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('PSE_INTERMITTENT').AsString = 'X' then
        if not JaiLeDroitTag(200072) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';IS;;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';IS;;ACTION=MODIFICATION')
      else
        if not JaiLeDroitTag(200072) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';IS;NOUVEAU;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';IS;NOUVEAU;ACTION=MODIFICATION');
    end
    else
      if JaiLeDroitTag(200072) then //PT184
        AGLLanceFiche('PAY', 'COMPLSALARIE', '', '', Salarie + ';IS;;ACTION=CREATION');
    Ferme(Q);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BTICKET' then
  begin
    Q := OpenSQL('SELECT PSE_TICKETREST FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('PSE_TICKETREST').AsString = 'X' then
        if not JaiLeDroitTag(200073) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';TCK;;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';TCK;;ACTION=MODIFICATION')
      else
        if not JaiLeDroitTag(200073) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';TCK;NOUVEAU;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';TCK;NOUVEAU;ACTION=MODIFICATION');
    end
    else
      if JaiLeDroitTag(200073) then //PT184
        AGLLanceFiche('PAY', 'COMPLSALARIE', '', '', Salarie + ';TCK;;ACTION=CREATION');
    Ferme(Q);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BHIERARCHIE' then
  begin
    Q := OpenSQL('SELECT PSE_ORGANIGRAMME FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', True);
    if not Q.Eof then
    begin
      if Q.FindField('PSE_ORGANIGRAMME').AsString = 'X' then
        if not JaiLeDroitTag(200074) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';RESP;;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';RESP;;ACTION=MODIFICATION')
      else
        if not JaiLeDroitTag(200074) then //PT184
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';RESP;NOUVEAU;ACTION=CONSULTATION')
        else
          AGLLanceFiche('PAY', 'COMPLSALARIE', '', Salarie, Salarie + ';RESP;NOUVEAU;ACTION=MODIFICATION');
    end
    else
      if JaiLeDroitTag(200074) then //PT184
        AGLLanceFiche('PAY', 'COMPLSALARIE', '', '', Salarie + ';RESP;;ACTION=CREATION');
    Ferme(Q);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BSAISIEARRET' then
  begin
    AGLLanceFiche('PAY', 'RETENUESAL_MUL', '', '', Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BMEDECINE' then
  begin
    if not JaiLeDroitTag(200076) then   //PT184
      AGLLanceFiche('PAY', 'PLANRDVMDT', '', '', Salarie+';ACTION=CONSULTATION')  //PT184
    else
      AGLLanceFiche('PAY', 'PLANRDVMDT', '', '', Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BHANDICAPE' then
  begin
    if ExisteSQL('SELECT PGH_SALARIE FROM HANDICAPE WHERE PGH_SALARIE="' + Salarie + '"') then
      if not JaiLeDroitTag(200077) then   //PT184
        AGLLanceFiche('PAY', 'PGHANDICAPE', '', Salarie, 'ACTION=CONSULTATION') //PT184
      else
        AGLLanceFiche('PAY', 'PGHANDICAPE', '', Salarie, '')
    else
      if JaiLeDroitTag(200077) then  //PT184
        AGLLanceFiche('PAY', 'PGHANDICAPE', '', '', 'ACTION=CREATION;' + Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BZONEDYNA' then
  begin
    SetActiveTabSheet('PELTCOMPL');
    AfficheEltCompl;
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BELTNATDOS' then
  begin
    SetActiveTabSheet('PELTNATDOS');
    AfficheEltNat;
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BELTDOSPLUS' then
  begin
//    AGLLanceFiche ('PAY', 'ELTNATDOS_MUL', '','','SAL;'+Salarie);
    //PT178
    if Salarie = '' then
      PgiBox('Saisie obligatoire. Veuillez renseigner le matricule du salarié pour accéder à la saisie des éléments dossier.',Ecran.caption)
    else
    begin
    //FIN PT178
      AglLanceFiche('PAY', 'ELEMENTNATDOS', '', '', 'ACTION=CREATION;' + 'SAL' + ';' + Salarie);
      AfficheEltNat(True);
    end;
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BIJSS' then
  begin
    if not jailedroittag(200078) then //PT184
      AGLLanceFiche('PAY', 'REGLTIJSS', '', '', 'FICHESAL;' + Salarie + ';' + getfield('PSA_ETABLISSEMENT') + ';' + getfield('PSA_LIBELLE') +
        ';' + getfield('PSA_PRENOM') + ';' + DateToStr(V_PGI.DateEntree) + ';' + DateToStr(PlusMois(V_PGI.DateEntree, -12)) + ';' + getfield('PSA_NUMEROSS') + ';' + DateToStr(getfield('PSA_DATENAISSANCE')) + ';ACTION=CONSULTATION')
    else
      AGLLanceFiche('PAY', 'REGLTIJSS', '', '', 'FICHESAL;' + Salarie + ';' + getfield('PSA_ETABLISSEMENT') + ';' + getfield('PSA_LIBELLE') +
        ';' + getfield('PSA_PRENOM') + ';' + DateToStr(V_PGI.DateEntree) + ';' + DateToStr(PlusMois(V_PGI.DateEntree, -12)) + ';' + getfield('PSA_NUMEROSS') + ';' + DateToStr(getfield('PSA_DATENAISSANCE')));

    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BPOPULATION' then
  begin
    AGLLanceFiche('PAY', 'AFFECTSALPOP_MUL', '', '', Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BFORMATIONSINT' then
  begin
    RempliOnglet('FORMATIONSINT');
    SetActiveTabSheet('PFORMINT');
    if not jailedroittag(200079) then //PT184
      SetControlEnabled('PFORMINT',false);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BFORMATIONSEXT' then
  begin
    RempliOnglet('FORMATIONSEXT');
    SetActiveTabSheet('PFORMEXT');
    if not jailedroittag(200079) then //PT184
      SetControlEnabled('PFORMEXT',false);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BDIPLOMES' then
  begin
    RempliOnglet('DIPLOMES');
    SetActiveTabSheet('PDIPLOMES');
    if not jailedroittag(200079) then //PT184
      SetControlEnabled('PDIPLOMES',false);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BDIF' then
  begin
    cancel := true;
    if not jailedroittag(200079) then //PT184
      AGLLanceFiche('PAY', 'COMPTEURSDIF', '', '', GetField('PSA_SALARIE') + ';' + DateToStr(V_PGI.DateEntree)+';ACTION=CONSULTATION')
    else
      AGLLanceFiche('PAY', 'COMPTEURSDIF', '', '', GetField('PSA_SALARIE') + ';' + DateToStr(V_PGI.DateEntree)+';ACTION=MODIFICATION');
    exit;
  end
  else if ControlName = 'RIB__BOMPCETENCES' then
  begin
    RempliOnglet('BTNCOMPETENCE');
    SetActiveTabSheet('PCOMPETENCES');
    if not jailedroittag(200079) then //PT184
      SetControlEnabled('PCOMPETENCES',false);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BDUE' then
  begin
    AglLanceFiche('PAY', 'DUE', '', '', Salarie + ';CREATION');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BMALADIE' then
  begin
    if EstSalarieMSA then AglLanceFiche('PAY', 'MALADIE_MSA', '', '', Salarie + ';CREATION')
    else AglLanceFiche('PAY', 'MALADIE', '', '', Salarie + ';CREATION');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BACCTRAV' then
  begin
    if EstSalarieMSA then AglLanceFiche('PAY', 'ACCIDENTTRAV_MSA', '', '', Salarie + ';CREATION')
    else AglLanceFiche('PAY', 'ACCIDENTTRAVAIL', '', '', Salarie + ';CREATION');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BDECLACCTRAV' then
  begin
    if EstSalarieMSA then AGLLanceFiche('PAY', 'DECLACCTRAV_MSA', '', '', 'ACTION=CREATION;MSA;' + Salarie)
    else AGLLanceFiche('PAY', 'DECLACCTRAV', '', '', 'ACTION=CREATION;ACT;' + Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BDETAILCP' then
  begin
    if not JaiLeDroitTag(200067) then //PT184
      AGLLanceFiche('PAY', 'CONGESPAY_MUL', '', '', Salarie + ';C;;;SAL;ACTION=CONSULTATION')
    else
      AGLLanceFiche('PAY', 'CONGESPAY_MUL', '', '', Salarie + ';C;;;SAL');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BFICHIND' then
  begin
    AglLanceFiche('PAY', 'FICHEIND', '', '', 'SAL;' + Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BABSENCE' then
  begin
    if not jailedroittag(200078) then  //PT184
      AglLanceFiche('PAY', 'ABSENCE_MUL', '', '', Salarie+';SALARIES;;ACTION=CONSULTATION')
    else
      AglLanceFiche('PAY', 'ABSENCE_MUL', '', '', Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BHISTOSAL' then
  begin
    if not jailedroittag(200078) then  //PT184
      AGLLanceFiche('PAY', 'PGHISTOSAL_MUL', '', '', 'FICHESAL' + Salarie + ';S;ACTION=CONSULTATION')
    else
      AGLLanceFiche('PAY', 'PGHISTOSAL_MUL', '', '', 'FICHESAL' + Salarie + ';S');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BBULLETIN' then
  begin
    AGLLanceFiche('PAY', 'BUL_CONSULT', '', '', Salarie);
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BGED' then
  begin
    bGedOnClick(TToolBarButton97(Getcontrol('BGED')));
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BSAISEIARRET' then
  begin
    if not JaiLeDroitTag(200075) then //PT184
      AGLLanceFiche('PAY', 'RETENUESAL_MUL', '', '', Salarie+';ACTION=CONSULTATION')
    else
      AGLLanceFiche('PAY', 'RETENUESAL_MUL', '', '', Salarie+';ACTION=MODIFICATION');
    cancel := true;
    exit;
  end
  else if ControlName = 'RIB__BVENTILRUB' then
  begin
    if not JaiLeDroitTag(200068) then //PT184
      AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM')+';CONSULTATION')
    else
      AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM'));
    cancel := true;
    exit;
  end
  //DEB PT184
  else if ControlName = 'RIB__BVENTILSALARIE' then
  begin
    if not JaiLeDroitTag(200069) then  //PT184
      ParamVentil ('SA',GetField('PSA_SALARIE'),'12345',taConsult,FALSE)
    else
      ParamVentil ('SA',GetField('PSA_SALARIE'),'12345',taModif,FALSE);
    cancel := true;
    exit;
  end
  //FIN PT184
  else if ControlName = 'RIB__BASSEDIC' then
   begin
   if EstSalarieSPectacle then
      AGLLanceFiche ('PAY', 'MUL_SALSPECTACLE', '', '',
                     Salarie+';'+DateToStr (V_PGI.DateEntree)+';'+DateToStr (V_PGI.DateEntree))
   else
{PT171
      AGLLanceFiche ('PAY', 'ASSEDIC', '', '',
                     Salarie+';ASS;-1;ACTION=CREATION');
}
      begin
      Existe:= ExisteSQL ('SELECT PAS_SALARIE'+
		          ' FROM ATTESTATIONS WHERE'+
                	  ' PAS_SALARIE = "'+Salarie+'" AND'+
                	  ' PAS_TYPEATTEST = "ASS"');
      if (Existe=True) then
         AGLLanceFiche ('PAY', 'MUL_ATTESTASSED', '', '', 'S')
      else
         AGLLanceFiche ('PAY', 'ASSEDIC', '', '',
                        Salarie+';ASS;-1;ACTION=CREATION');
      end;
//FIN PT171
   cancel:= true;
   exit;
   end;
end;


{$ENDIF}
//DEBUT PT156

procedure TOM_SALARIES.CasParticulierDblClick(Sender: TObject);
var Theme, Profil, Ret: string;
  PCas: THGrid;
  Themes: THValCombobox;
begin
  PCas := THGrid(GetControl('PCAS'));
  Themes := THValCombobox(GetControl('THEME'));
  Theme := Themes.Values[PCas.Row - 1];
  Profil := PCas.CellValues[2, PCas.Row];
  Ret := AGLLanceFiche('PAY', 'CASPARTICULIERS', '', '', Theme + ';' + Profil);
  if Ret <> 'NONMODIFIE' then
  begin
    PCas.CellValues[2, PCas.Row] := Ret;
    PCas.CellValues[1, PCas.Row] := RechDom('PGPROFILPARTICULIER', Ret, FALSE);
    DS.Edit; //FC 20071004 Pour forcer le déclenchement de la mise à jour
  end;
end;

procedure TOM_SALARIES.BtnContratClick(Sender: Tobject); //PT170
begin
  if not JaiLeDroitTag(200062) then  //PT184
    AGLLanceFiche('PAY', 'MUL_CONTRAT', '', GetField('PSA_SALARIE'), GetField('PSA_SALARIE')+';ACTION=CONSULTATION', '')
  else
    AGLLanceFiche('PAY', 'MUL_CONTRAT', '', GetField('PSA_SALARIE'), GetField('PSA_SALARIE'), '');
  AfficheContrat;
end;

procedure TOM_SALARIES.AfficheContrat; //PT170
var DD, DF, DE, DR: TDateTime;
  TypeC: string;
  Q: TQuery;
begin
  // Gestion des contrats de travail ==> Affichage du dernier contrat
  DD := 0;
  DF := 0;
  DE := 0;
  TypeC := '';
  Q := OpenSQL('SELECT PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_TYPECONTRAT,' +
    'PCI_ESSAIFIN,PCI_RNVESSAIFIN,PCI_ORDRE' +
    ' FROM CONTRATTRAVAIL WHERE' +
    ' PCI_SALARIE = "' + GetField('PSA_SALARIE') + '"' +
    ' ORDER BY PCI_DEBUTCONTRAT DESC', TRUE); // PT115 Ordre des contrats par date debut decroissante
  if not Q.Eof then
  begin
    DD := Q.findfield('PCI_DEBUTCONTRAT').AsDateTime;
    DF := Q.findfield('PCI_FINCONTRAT').AsDateTime;
    DE := Q.findfield('PCI_ESSAIFIN').AsDateTime;
    DR := Q.findfield('PCI_RNVESSAIFIN').AsDateTime;
    if DR > DE then
      DE := DR;
    TypeC := Q.findfield('PCI_TYPECONTRAT').AsString;
    NumOrdre := Q.findfield('PCI_ORDRE').AsInteger;
  end
  else NumOrdre := 0; // PT115
  Ferme(Q);
  if DD <> 0 then
    SetControlProperty('LBLCONTRATDD', 'Caption', DateToStr(DD))
  else
    SetControlProperty('LBLCONTRATDD', 'Caption', '');
  if DF <> 0 then
    SetControlProperty('LBLCONTRATDF', 'Caption', DateToStr(DF))
  else
    SetControlProperty('LBLCONTRATDF', 'Caption', '');
  if DE <> 0 then
    SetControlProperty('LBLFINESSAI', 'Caption', DateToStr(DE))
  else
    SetControlProperty('LBLFINESSAI', 'Caption', '');
  if TypeC <> '' then
    SetControlProperty('TPSA_TYPEDECONTRAT', 'Caption', RechDOM('PGTYPECONTRAT', TypeC, FALSE))
  else
    SetControlProperty('TPSA_TYPEDECONTRAT', 'Caption', '');
end;

//FIN PT156
{//Debut PT151
procedure TOM_SALARIES.SavFieldsPopul;
var
  QPaieParim : TQuery;
  indexField, RowCount : Integer;
begin
  QPaieParim := OpenSQL('Select PAI_suffix from PAIEPARIM where pai_prefix = "PSA" and PAI_CRITERETABLE = "X"',True);
  RowCount := QPaieParim.RecordCount;
  SetLength(SavTableFields, RowCount);
  SetLength(SavTableValues, RowCount);
  indexField := 0;
  while not QPaieParim.Eof do
  begin
    SavTableFields[indexField] := QPaieParim.Fields[0].AsString;
    SavTableValues[indexField] := GetField('PSA_'+SavTableFields[indexField]);
    Inc(indexField);
    QPaieParim.Next;
  end;
end;

function TOM_SALARIES.ValidUpdateFields : Boolean;
var
  UpdateIdemPop : TUpdateIdemPop;
  TobSalaries, TobUnSalarie, TobChampsUnSalarie : Tob;
  indexChamp : integer;
begin
  TobSalaries := TOB.Create('Les salariés et leurs modifications',nil,-1);
  TobUnSalarie := TOB.Create('Un salariés et ses modifications',TobSalaries,-1);
  TobUnSalarie.AddChampSupValeur('SALARIE',GetField('PSA_SALARIE'));
  //On parcours tous les champs sauvés, et en cas de modification on met le champs
  //dans la tob.
  for indexChamp := 0 to Length(SavTableFields)-1 do
  begin
    if GetField('PSA_'+SavTableFields[indexChamp]) <> SavTableValues[indexChamp] then
    begin
      TobChampsUnSalarie := TOB.Create('Les champs modifiés du salarié',TobUnSalarie,-1);
      TobChampsUnSalarie.AddChampSupValeur('CHAMPS',SavTableFields[indexChamp]);
      TobChampsUnSalarie.AddChampSupValeur('VALEUR',GetField('PSA_'+SavTableFields[indexChamp]));
    end;
  end;
  UpdateIdemPop := TUpdateIdemPop.Create;
  result := UpdateIdemPop.ValideModifSalarie(TobSalaries);
  TobSalaries.Free;
  UpdateIdemPop.Free;
end;
//Fin PT151   }

//DEB PT183
function TOM_SALARIES.ClauseSQL(Nature:String):String;
begin
  Result := 'Select PMI_CODE,PMI_LIBELLE FROM MINIMUMCONVENT M1' +
    ' WHERE  ##PMI_PREDEFINI## PMI_NATURE="' + Nature + '"' +
    ' AND (PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" OR PMI_CONVENTION="000")' +
    ' AND (PMI_PREDEFINI="DOS"' +
    ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'" ' +
    '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT ' +
    '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE AND PMI_PREDEFINI="DOS"))' +
    ' OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="000" ' +
    '   AND NOT EXISTS(SELECT PMI_CODE FROM MINIMUMCONVENT '+
    '   WHERE ##PMI_PREDEFINI## PMI_CODE=M1.PMI_CODE ' +
    '   AND (PMI_PREDEFINI="DOS" OR (PMI_PREDEFINI="STD" AND PMI_CONVENTION="'+getfield('PSA_CONVENTION')+'")))))';
//    ' ORDER BY PMI_CODE';
end;

procedure TOM_SALARIES.CoeffElipsisclick(Sender: Tobject);
var
{$IFNDEF EAGLCLIENT}
  ChampLib : THDBEdit;
{$ELSE}
  ChampLib : THEdit;
{$ENDIF}
  St:String;
begin
{$IFNDEF EAGLCLIENT}
  ChampLib := THDBEdit(GetControl('PSA_COEFFICIENT'));
{$ELSE}
  ChampLib := THEdit(GetControl('PSA_COEFFICIENT'));
{$ENDIF}
  If ChampLib <> Nil then
  begin
    St := ClauseSQL('COE');
    LookUpList(ChampLib, 'Coefficient', 'PGLIBCOEFFICIENT', 'PMI_CODE', 'PMI_LIBELLE', '', '', TRUE, -1,St);
  end;
end;

procedure TOM_SALARIES.IndiceElipsisclick(Sender: Tobject);
var
{$IFNDEF EAGLCLIENT}
  ChampLib : THDBEdit;
{$ELSE}
  ChampLib : THEdit;
{$ENDIF}
  St:String;
begin
{$IFNDEF EAGLCLIENT}
  ChampLib := THDBEdit(GetControl('PSA_INDICE'));
{$ELSE}
  ChampLib := THEdit(GetControl('PSA_INDICE'));
{$ENDIF}
  If ChampLib <> Nil then
  begin
    St := ClauseSQL('IND');
    LookUpList(ChampLib, 'Indice', 'PGLIBINDICE', 'PMI_CODE', 'PMI_LIBELLE', '', '', TRUE, -1,St);
  end;
end;

procedure TOM_SALARIES.NiveauElipsisclick(Sender: Tobject);
var
{$IFNDEF EAGLCLIENT}
  ChampLib : THDBEdit;
{$ELSE}
  ChampLib : THEdit;
{$ENDIF}
  St:String;
begin
{$IFNDEF EAGLCLIENT}
  ChampLib := THDBEdit(GetControl('PSA_NIVEAU'));
{$ELSE}
  ChampLib := THEdit(GetControl('PSA_NIVEAU'));
{$ENDIF}
  If ChampLib <> Nil then
  begin
    St := ClauseSQL('NIV');
    LookUpList(ChampLib, 'Niveau', 'PGLIBNIVEAU', 'PMI_CODE', 'PMI_LIBELLE', '', '', TRUE, -1,St);
  end;
end;

procedure TOM_SALARIES.QualifElipsisclick(Sender: Tobject);
var
{$IFNDEF EAGLCLIENT}
  ChampLib : THDBEdit;
{$ELSE}
  ChampLib : THEdit;
{$ENDIF}
  St:String;
begin
{$IFNDEF EAGLCLIENT}
  ChampLib := THDBEdit(GetControl('PSA_QUALIFICATION'));
{$ELSE}
  ChampLib := THEdit(GetControl('PSA_QUALIFICATION'));
{$ENDIF}
  If ChampLib <> Nil then
  begin
    St := ClauseSQL('QUA');
    LookUpList(ChampLib, 'Qualification', 'PGLIBQUALIFICATION', 'PMI_CODE', 'PMI_LIBELLE', '', '', TRUE, -1,St);
  end;
end;
//FIN PT183

//DEB PT184
{$IFNDEF GCGC}
{PT191 procedure TOM_SALARIES.BBANQUE_OnClick(Sender: TObject);
var
  Init: word;  //PT190
begin
  if DS.State in [DsEdit, DsInsert] then  //PT190
  begin
    Init := HShowMessage('0;'+Ecran.Caption+';Attention, vous devez enregistrer vos modifications.#13#10Voulez-vous continuer ?;Q;YN;N;N;', '','');
    if Init = mrNo then
      exit
    else
      TFFiche(Ecran).BValiderClick(Sender);
//    PGIError('Vous devez d''abord valider vos modifications.', Ecran.Caption);
//    exit;
  end;
  if (not JaiLeDroitTag(200063))  then
    AgllanceFiche('PAY', 'SAL_BANQUECWAS', '', GetField('PSA_SALARIE'),'ACTION=CONSULTATION')
  else
    AgllanceFiche('PAY', 'SAL_BANQUECWAS', '', GetField('PSA_SALARIE'),'');
  RefreshDB;                              //PT190
end; }

procedure TOM_SALARIES.BMEMO_OnClick(Sender: TObject);
begin
  if (not JaiLeDroitTag(200064))  then
    AgllanceFiche('YY','YYLIENSOLE','PG;' + GetField('PSA_SALARIE'), '', 'ACTION=CONSULTATION')
  else
    AgllanceFiche('YY','YYLIENSOLE','PG;' + GetField('PSA_SALARIE'), '', '');
end;

{PT191 procedure TOM_SALARIES.BCP_OnClick(Sender: TObject);
var
  Init: word;  //PT190
begin
  if DS.State in [DsEdit, DsInsert] then  //PT190
  begin
    Init := HShowMessage('0;'+Ecran.Caption+';Attention, vous devez enregistrer vos modifications.#13#10Voulez-vous continuer ?;Q;YN;N;N;', '','');
    if Init = mrNo then
      exit
    else
      TFFiche(Ecran).BValiderClick(Sender);
//    PGIError('Vous devez d''abord valider vos modifications.', Ecran.Caption);
//    exit;
  end;
  if (not JaiLeDroitTag(200066))  then
    AgllanceFiche('PAY','SALARIE_CP','',GetField('PSA_SALARIE'),GetField('PSA_SALARIE') + ';ACTION=CONSULTATION')
  else
    AgllanceFiche('PAY','SALARIE_CP','',GetField('PSA_SALARIE'),GetField('PSA_SALARIE'));
  RefreshDB;                              //PT190
end;}

procedure TOM_SALARIES.BtnSALRUBClick(Sender: Tobject);
begin
  if not JaiLeDroitTag(200068) then
    AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM')+';CONSULTATION')
  else
    AGLLanceFiche('PAY', 'ANA_SALRUB', '', ' ', GetField('PSA_SALARIE') + ';' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM'));
end;
{$ENDIF !GCGC}
//FIN PT184


procedure TOM_SALARIES.OnBeforeUpdateRecord;
begin
  inherited;
{
  if (Argument <> 'INT') AND (DS.State in [dsInsert]) then
  begin
    if (VH_PAIE.PgTypeNumSal = 'NUM') then
    begin
      if (VH_Paie.PGIncSalarie = TRUE) then
      begin
        if (existesql ('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+GetField('PSA_SALARIE')+'"')) then
        begin
          AffectCodeNewEnr;
          TFFiche(Ecran).caption := 'Salarié : ' + GetField('PSA_SALARIE') + ' ' +
            GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM');
          UpdateCaption(TFFiche(Ecran));
        end;
      end;
    end;
  end;
}
end;

initialization
  registerclasses([TOM_SALARIES]);
  // DEBUT CCMX-CEGID ORGANIGRAMME DA - 13.06.2006
{$IFDEF GCGC}
  RegisterAglProc('ZonesComplIsModified', TRUE, 0, AGLZonesComplIsModified);
{$ENDIF}
  // FIN CCMX-CEGID ORGANIGRAMME DA - 13.06.2006

end.



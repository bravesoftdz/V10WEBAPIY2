{***********UNITE*************************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 30/05/2001
Modifié le ... : 30/05/2001
Description .. : Unit de définition de toutes les variables nécessaires
Suite ........ : aux calculs des bulletins
Suite ........ : Fonctions de calculs
Mots clefs ... : PAIE;PGBULLETIN;PGREGUL
*****************************************************************}
{
PT1    : 25/06/2001 PH V536 Suppression des Q.next
PT2    : 25/07/2001 VG V540 Ajout champs dans Historique salarié
PT3    : 31/07/2001 SB V547 Modification champ suffixe MODEREGLE
PT4    : 29/08/2001 PH V547 correction bornes de validité remuneration sur 4
                            mois
PT5    : 29/08/2001 PH V547 correction bornes de validité cotisaation sur 4 mois
PT6    : 30/08/2001 SB V547 Calcul du HCPPayés fiche bug n° 294
PT7    : 03/09/2001 PH V547 correction validité du au pour les rémunérations
PT8    : 03/09/2001 PH V547 on force dans tous les cas le calcul de l'arrondi
                            sur le montant calculé
PT9    : 05/09/2001 SB V547 L'acquisition annuelle des CP ne correspond plus à
                            25 ou 30 mais à l'acquis mensuelle*12
                            On passe en parametre ETB_NBREACQUISCP et non
                            ETB_NBJOUTRAV
PT10   : 07/09/2001 PH V547 Test du profil FNAL on recherche le profil FNAL que
                            dans la table Salarie ou dans les paramsoc. test
                            plus sévère
PT11   : 10/09/2001 SB V547 Fiche de bug n°266 : Gestion des msg d'erreur CP en
                            prep auto . fn SalIntegreCP modifiée (Aussi ds
                            source SaisBul)
PT12   : 02/10/2001 PH V562 Rajout historisation salarie profil rémunération
PT13   : 15/10/2001 VG V562 Ajout champs dans Historique salarié
PT14   : 23/10/2001 PH V562 Rajout traitement bulletion complémentaire et dates
                            édition
PT15   : 24/10/2001 PH V562 Calcul periode glissante pour variables de type
                            cumul ou rubrique
PT16   : 24/10/2001 PH V562 Date à date glissante dans le cas ou mois fin < mois
                            début
PT17   : 13/11/2001 SB V562 Prise de CP sur Acquis en cours
PT18   : 14/11/2001 PH V562 Identification du taux AT en fonction des dates de
                            validité
PT19   : 16/11/2001 PH V562 Modif du calcul de la date de paiement dans le cas
                            bulletin complémentaire
PT20   : 20/11/2001 PH V562 Modif du calcul des dates d'édition automatique en
                            fonction de l'établissement
PT21   : 20/11/2001 PH V562 Modif du calcul des variables recherchant un cumul
                            annuel + paie en cours
PT22   : 22/11/2001 PH V562 ORACLE, Le champ PHB_ORGANISME doit être renseigné
PT23   : 22/11/2001 SB V563 Pour ORACLE, Le champ PCN_CODETAPE doit être
                            renseigné
PT24   : 27/11/2001 PH V563 Propagation d'une ligne de commentaire
PT25   : 04/12/2001 PH V563 Mode de réglement pour les bulletins complémentaires
PT26   : 19/12/2001 PH V571 Traitement ligne de commentaire pour les cotisations
PT27   : 26/12/2001 PH V571 Dédoublement des lignes multiples de commentaire
PT28   : 27/12/2001 PH V571 Calcul des variables de type cumul en tenant compte
                            du mois de RAZ
PT29   : 27/12/2001 PH V571 Calcul des variables de type rémunération tq sur
                            Mois - 1
PT30   : 27/12/2001 PH V571 Calcul des dates Trimestres,années - X et en tenant
                            compte du décalage
PT31   : 08/01/2002 MF V571 correction fct DiffMoisJour : initialisation de la
                            variable Calcul
PT32   : 07/02/2002 PH V571 suppression historisation salarié
PT33   : 19/03/2002 SB V571 Fiche de bug n°452 Ajout variable de paie prédefini
                            Cegid : Heures et jours absences pris
PT34   : 22/03/2002 PH V571 Traitement de la civilité  et de Mode réglement
                            modifie dans la table Paieencours
PT35   : 26/03/2002 PH V571 Fonction de raz de ChptEntete
PT36   : 26/03/2002 PH V571 10eme ligne de calcul des variables de type calcul
PT37   : 02/04/2002 PH V571 traitement du profil retraite
PT38   : 03/04/2002 PH V571 Modif Calcul base de précarité
PT39   : 04/04/2002 PH V571 Modif insertion rub contenant une variable de nature
                            cumul sur période annuelle
PT40   : 04/04/2002 PH V571 Test si colonne cotisation saisissable
PT41   : 05/04/2002 PH V571 variable glissante sur X mois prennait 1 mois de
                            trop
PT42   : 23/04/2002 SB V571 Fiche de bug 10003 : Calcul du montant de l'absence
                            non calculé sur 1er bulletin en création
PT43   : 23/04/2002 SB V571 Fiche de bug 10021 : Ajout de la variable
                            HTHTRAVAILLES pour calcul heures théoriques
                            travaillés
PT44   : 29/04/2002 Ph V582 Rajout clause group by pour ????? forcer index
PT45   : 02/05/2002 Ph V582 initialisation du taux at au debut du traitement
PT46   : 02/05/2002 SB V571 Erreur : Pour un SLD, On consomme systématiquement
                            l'acquis en cours dans sa totalité même si on ne
                            solde pas la totalité sur ce mvt (cas AJU Ou AJP) du
                            fait de l'utilisation de pls tob pour l'integration
                            des acquis en cours
PT47   : 28/05/2002 PH V571 Modif Calcul base de précarité Seuil ramené à 0 jour
PT48   : 11/06/2002 PH V582 Modif des variables des noms des responsables pour
                            traiter tous les cas
PT49   : 18/06/2002 PH V582 Modif Refonte complète du calcul de la précarité
PT50   : 01/07/2002 PH V582 modif sur calcul fourchette de dates pour les
                            variables de type cumul
PT51   : 19/06/2002 PH V582 Integration des lignes de la saisie des primes dans
                            le bulletin complémentaire Voir PT40 dans SaisBul
PT52-1 : 18/07/2002 VG V585 Correction précarité si un seul contrat et que la
                            date de début de contrat est la même que la date de
                            début de bulletin, on ne fait pas le calcul du
                            trentième, on prend la totalité de la base
PT52-2 : 19/07/2002 VG V585 Correction précarité si plusieurs contrats dans le
                            cas de début et/ou fin de contrat en cours de mois
PT53   : 22/07/2002 SB V582 Modification du calcul des variables 65, 66
PT54   : 25/07/2002 SB V585 Dû à l'intégration des mvts d'annulation d'absences
                            Contrôle des requêtes si typemvt et sensabs en
                            critere
PT55   : 09/08/2002 PH V582 Evaluation des variables de tests avec des strings
PT56   : 09/08/2002 PH V582 Creation des variables CEGID pour recup champs
                            salaries
PT57   : 05/09/2002 SB V585 Intégration de la gestion salarié de la méthode de
                            valorisation au maintien des CP
PT58   : 06/09/2002 PH V585 Variables pour restituer les elements de salaires 5
PT59   : 03/10/2002 SB V585 Optimisation requête saisie de la paie
PT60   : 07/10/2002 PH V585 Recherche d'un element national par préférence
                            DOS,STD,CEG
PT61   : 17/10/2002 SB V585 FQ n°10235 Intégration du calcul des variables 28,29
                            heures ouvrées ouvrables
PT62   : 07/11/2002 PH V591 Prise en compte du decalage de paie sur les éléments
                            nationaux en fonction du champ decalage
PT63   : 07/11/2002 PH V591 Creation variable 80 rend le mois de la paie
                            (datefin) FQ 10302
PT64   : 07/11/2002 PH V591 Variable de paie rend le taux patronal ou salarial
                            FQ 10276
PT65   : 02/12/2002 PH V591 Recup des infos du bulletin précédent en excluant
                            les bulletins complémentaires
PT66   : 06/12/2002 PH V591 Variable 2 rend horaire Hebdo de la fiche salarié
PT67   : 06/12/2002 PH V591 Nouvelles Variables 72 à 75 + modif 70 et 71
PT68   : 19/12/2002 PH V591 Affectation de la date de paiement à la date de fin
                            de paie si non renseignée
PT69-1 : 14/01/2003 SB V591 Suppression des lignes de commentaire CP inopérandes
                            si modification date de fin de bulletin
PT69-2 : 16/01/2003 SB V591 Calcul variable 50 : Integration du solde que si
                            salarié sorti
PT70   : 23/01/2003 PH V591 Correction du calcul des bornes de dates pour la
                            valorisation d'un cumul antérieur avec une période
                            de raz
PT71   : 30/01/2003 V591 SB Vidage Tob_Abs en sortie et ré entrée de bulletin
PT72   : 05/02/2003 PH V595 Traitement des tables dossiers en string au lieu
                            d'integer
PT73   : 05/02/2003 PH V595 Tables dossier gèrent aussi une élément national
                            comme retour
PT74   : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
PT75-1 : 02/06/2003 SB V595 FQ 10705 Mvt SLD de Cloture non valorisé..
PT75-2 : 04/06/2003 SB V595 FQ 10523 Intégration de l'acquis bulletin calculé
                            sur une variable de paie
PT76   : 04/06/2003 PH V421 FQ 10425,10620,10655 nvelles variables
PT77   : 04/06/2003 PH V421 FQ 10689 Salarié sort dans la période mais la
                            cotisation ne se calcule que si le salarié est
                            present le dernier jour du mois
PT78   : 05/06/2003 PH V421 FQ 10700 Gestion origine de la rubrique dans les
                            bases de cotisations
PT79-1 : 06/06/2003 SB V595 Test minimum conventionnel en alpha même pour du
                            numerique
PT80   : 11/06/2003 PH V421 Modification fct de contrôle division par 0 Pour
                            tester si 0.xxx au lieu de O
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT81   : 23/06/2003 PH V421 Modif variables 72,73,74 FQ 10565  appel variable 35
                            au lieu de 65 Horaire etablissment
PT82-1 : 24/06/2003 SB V_42 FQ 10454 Calcul au maintien sur premier bulletin
                            implique un recalcul des CP
PT82-2 : 25/06/2003 SB V_42 FQ 10717 Calcul nombre de mois anciennete erroné :
                            Nombre de mois entier
PT83-3 : 27/06/2003 SB V_42 FQ 10628 Intégration procedure à part de l'init des
                            variables heures et jours travaillés
PT83-4 : 11/08/2003 SB V_42 FQ 10675 Gestion des lignes de commentaire CP par
                            type
PT84   : 26/08/2003 PH V_42 Mise en place saisie arret
PT85   : 09/09/2003 PH V_42 Correction variable 75,74
PT86   : 17/09/2003 PH V421 Traitement des tables DIW pour une variable de type
                            valeur (idem DIV)
PT87   : 06/10/2003 PH V421 FQ 10595 Gstion des apppoints avec date de sortie
PT88   : 10/10/2003 SB V_42 Intégration d'un indicateur des CP pris non payés
                            par manque d'acquis
PT89   : 21/10/2003 PH V_42 FQ 10928 Recup de l'appoint precedant en négatif
                            sinon mauvais sens
PT90   : 18/11/2003 SB V_50 FQ 10794 Controle profil Congés payés existant
PT90-1 : 18/11/2003 SB V_50 FQ 11121 Contrôle rubrique alimenté pour génération
                            libellé CP
PT91   : 05/12/2003 PH V_50 Création de la variable 103 FQ 11002
PT92   : 10/12/2003 PH V_50 Recalcul systématique de la saisie arrêt
PT93   : 11/12/2003 PH V_50 initialisation valeur par defaut FQ 11003
PT94   : 18/12/2003 PH V_50 FQ 11028 Calcul des cotisations que si présent en
                            fin de mois-traitement de la date à idate1900
PT95   : 12/01/2004 PH V_50 FQ 11024 Rubrique sur le bulletin en origine salarié
                            reprend les valeurs même si elt variable
PT96   : 24/02/2004 SB V_50 FQ 11130 Calcul Ancienneté érronné si salarié entrée
                            sortie le même jour
PT97   : 18/03/2004 SB V_50 FQ 11020 Modification utilisateur des acquis sur la
                            saisie de la paie
PT98   : 19/03/2004 PH V_50 FQ 11200 Prise en compte de nbre de décimales dans
                            le calcul des variables
PT99   : 19/03/2004 PH V_50 FQ 11106 Variable 0070 et 0071
PT100  : 05/04/2004 PH V_50 FQ 11231 Prise en compte du cas spécifique
                            Alsace Lorraine
PT100-1: 09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                            salarié
PT100-2: 09/04/2004 SB V_50 FQ 11237 Ajout Gestion idem etab jours Cp Suppl
PT101  : 07/05/2004 SB V_50 FQ 11199 Dysfonctionnement Calcul Indemnité CP sur
                            1er bulletin
PT102  : 10/05/2004 PH V_50 restriction droit accès au paramétrage CEGID
PT103  : 17/06/2004 MF V_50 Ajout traitement de suppression des IJSS
PT104  : 21/06/2004 PH V_50 FQ 10793 Indicateur de contrôle de modification des
                            salariés pdt la saisie du bulletin
PT105  : 22/06/2004 PH V_50 Correction fonction RendDateVar ==> calcul de
                            variable de type cumul sur trimestre
PT106  : 23/06/2004 PH V_50 FQ 11380 Principe en modification de bulletin pas
                            d'alignement des rubriques par rapport aux profils
                            en création alignement, bouton Défaire ou revenir à
                            l'état initial provoque l'alignement
PT107  : 23/06/2003 MF V_50 IJSS- Modif message qd ligne non modifiable-->modif
                            PHB_ORIGINELIGNE
PT108  : 05/07/2004 PH V_50 FQ 11052 Modif Fiche ETABSOCIAL et contrôle sur les
                            dates d'édition
PT109  : 05/07/2004 PH V_50 FQ 10959 Nouvelles variables prédéfinies CEGID
PT110  : 06/07/2004 PH V_50 FQ 10991 Nouvelles variables prédéfinies CEGID
PT111  : 12/07/2004 PH V_50 FQ 11410 Nouvelles variables prédéfinies CEGID
PT112  : 10/08/2004 PH V_50 FQ 11464 Prise en compte de la période en cours pour
                            les variables de type rem,cot
PT113  : 12/08/2004 PH V_50 FQ 11148 Prise en compte du montant de l'acompte
                            dans le calcul de la saisie arret
PT114  : 19/08/2004 PH V_50 FQ 11432 Message et non calcul d'une variable si
                            celle-ci n'existe pas sionon provoque un ACCESS VIO
PT115  : 19/08/2004 PH V_50 FQ 10991 Suite : inversion du calcul en fonction du
                            code de la rubrique Mois et Année
PT116  : 20/08/2004 PH V_50 FQ 10959 Récupération du champ sortie définitive
PT117  : 23/08/2004 PH V_50 FQ 11459 Limitation des trentièmes si 2 bulletins
                            dans le même mois
PT118  : 23/08/2004 PH V_50 FQ 11425 Variable 140
PT119  : 23/08/2004 PH V_50 FQ 11429 et 11306 Variable recherchant une table
                            Divers
PT120  : 24/08/2004 PH V_50 FQ 11454 Suite Si pas élément Nat trouvé alors on
                            rend 0
PT121  : 25/08/2004 MF V_50 Ajout traitement de suppression des lignes de
                            maintien
PT122  : 13/09/2004 PH V_50 FQ 10992 Variable 0092 Nombre de jours calendaires
                            du mois
PT123  : 23/09/2004 PH V_50 FQ 11638 Calcul apprenti sur mauvaise variable
PT124  : 30/11/2004 MF V_60 ajout de commentaires pour le maintien
PT125  : 02/12/2004 PH V_60 FQ 11818 Calcul automatique de la variable 0033
PT126  : 23/12/2004 PH V_60 Prise en compte variable de type cumul sur période
                            annuelle pour le 1er mois dans le cas du décalage de
                            paie
PT127-1: 21/01/2005 SB V_60 FQ 11889 Calcul Cp Acquis sur variable de paie
PT127-2: 21/01/2005 SB V_60 FQ 11882 Calcul Indemnité & absence indemnité CP
PT128  : 03/02/2005 SB V_60 FQ 11089 Calcul erronné ancienneté en jours variable
                            0010
PT129  : 08/02/2005 MF V_60 Traitement des commentaires des règlements d'IJSS
                            (2 nvelles fcts)
PT130  : 18/03/2005 SB V_60 FQ 11406 Ajout var déduction entrée et déduction
                            sortie
PT131-1: 04/01/2005 SB V_60 FQ 11781 Absence : Intégration de pls motifs sur
                            idem rubrique
PT131-2: 04/01/2005 SB V_60 FQ 11806 Voir Absence : Refonte traitement pour
                            recalcul bulletin si modif absence
PT133  : 10/05/2005 PH V_60 FQ 11773 Variable 104 qui rend le code etab si
                            numérique
PT134  : 11/05/2005 PH V_60 FQ 11687 Variable 105 qui rend le code DADS-U du
                            contrat de travail, -1 si non trouvé
PT134-1: 02/06/2005 SB V_60 FQ 12327 Econges : En monobase mise à jour du
                            exportok via la paie
PT135  : 29/06/2005 PH V_60 FQ 12410 calcul SMIC par rapport aux cumuls des
                            heures
PT136  : 12/07/2005 SB V_60 FQ 12308 Modification de l'ordre de présentation des
                            bulletins
PT137  : 08/08/2005 PH V_60 FQ 11746 Rechargement des taux AT en cas de modif
                            etablissement
PT138  : 01/09/2005 PH V_60 FQ 12510 Identification du bon taux AT
PT139  : 02/09/2005 PH V_60 FQ 12553 Gestion des commentaires sur les rubriques
                            de réguls
PT140  : 06/09/2005 SB V_65 FQ 12384 Bulletin ICCP : Rédéfinition des acquis
                            pris en compte pour le diviseur
PT141  : 16/09/2005 SB V_65 FQ 12535 Absences : numéro de rubrique de
                            commentaire mal incrémenté
PT142  : 07/10/2005 PH V_60 FQ 12629 Cas exclusion profil CSG/CRDS deductible et
                            non deductible ensembles
PT143  : 24/10/2005 SB V_65 FQ 12659 Absences : dysfonctionnement pls absences
                            identiques sur même bulletin
PT144  : 17/11/2005 SB V_65 FQ 12700 Conges : Ajout clause
                            date validité <= datefin bulltin
PT145  : 07/12/2005 PH V_65 FQ 12744 Ne pas récupérer le libellé d'une
                            cotisation de régularisation
PT146  : 08/12/2005 SB V_65 FQ 12700 Traitement de la tob des ajustement - en
                            tob globale
PT147  : 09/02/2006 SB V_65 FQ 12901 Vidage de la tob des rubriques CP pris
PT148  : 28/02/2006 MF V_65 FQ 12827 Exclusion de rub. de rem. pour calcul
                            maintien qd absence ou IJSS + Calcul Paie à l'envers
PT149  : 27/03/2006 PH V_65 Prise en compte des champs motifs absences dédoublés
                            (Heures,Jours)
PT150  : 03/04/2006 SB V_65 FQ 11426 Mis en place paramètre pour calcul au
                            maintien CP
PT151  : 10/04/2006 PH V_65 FQ 13037 Contrôle bon établissement si chgt
                            d'etablissement alors que bulletin déjà fait
PT152  : 10/04/2006 PH V_65 FQ 13057 initialisation du champ PHB_COTREGUL à ...
                            dans tous les cas
PT153  : 11/04/2006 SB V_65 Refonte intégration rubrique d'absence , gestion en
                            jours et en heures
PT154  : 18/04/2006 SB V_65 FQ 12885 Calcul des jours fériés
PT155  : 21/04/2006 SB V_65 FQ 12655 Génération des lignes de commentaires CP
                            que si rubrique liée intégré ds le bulletin
PT156  : 02/05/2006 PH V_65 FQ 13095 Rémunérations ne se calcule plus si en
                            dehors de la validité au lieu de calcul et raz des
                            montants
PT157  : 02/05/2006 PH V_65 FQ 13119 Calcul de variable de type cumul sur
                            timestre courant = rajout du mois en cours.
PT158  : 04/05/2006 SB V_65 FQ 12786 Anomalie calcul variable 0027 en prep. auto
PT159  : 16/05/2006 MF V_65 FQ 12995 (pls bull/mois) maintien et ijss intégrés
                            en fonction des dates du bulletin
PT160  : 02/06/2006 SB V_65 Optimisation mémoire
PT160-1: 31/08/2006 SB V_65 Optimisation mémoire
PT161  : 04/08/2006 PH V_65 Outil de dignostic refonte complète de toutes les
                            fonctions de calcul en rajoutant des paramètres.
PT162  : 27/09/2006 MF V7_0 FQ 13528 Correction Exclusion de rub. de rem. pour
                            calcul paie à l'envers
PT163  : 28/09/2006 MF V7_0 FQ 13490 Création variable 184 pour calcul garantie
                            mini au SMIC, comparaison avec cumul 05(brut
                            habituel)
PT164  : 02/11/2006 SB V_70 FQ 13399 Création variable pour recup jour &
                            montant CP N-1
PT165  : 11/12/2006 PH V_70 Gestion des réguls sur les rémunérations
PT166  : 29/12/2006 PH V_70 Gestion insertion automatique des rubriques Avec
                            PT177 Saisbul et PT44 Entpaie et
                            PT38 UTofPG_PrepAuto
PT167  : 02/01/2007 PH V_70 Alimentation champ rémunération à partir de la
                            valeur d'un cumul de la paie en cours
PT168  : 03/01/2007 PH V_70 FQ 13595 Alimentation du montant d'un cumul par le
                            taux d'une rémunération
PT169  : 19/01/2007 GGU V_80 Préparation automatique des payes au contrat
PT170  : 19/01/2007 GGU V_80 Possibilité d'augentation de la limite de lignes de
                            commentaire par rubrique dans le bulletin. Mais on
                            laisse la limite à 9 pour l'instant
PT171  : 02/02/2007 PH V_70 Paies au contrat
PT172  : 28/02/2007 FC V_70 Rajout de la gestion de la convention collective
                            pour STD
                            Intégration des éléments dynamiques dossier dans le
                            calcul du bulletin
PT173  : 27/03/2007 PH      Rajout paramètre fonction RecupTobSalarie pour prise
                            en compte Historique date
PT174  : 29/03/2007 PH      FQ 13628 Rajout champs date de validité
PT175  : 06/04/2007 PH      Prise en compte dans les variables des tables
                            dynamiques
PT176  : 17/04/2007 FC V_72 FQ 13918 Rajout du matricule/nom/prénom du salarié
                            dans les messages d'alerte dans la trace
PT177  : 02/05/2007 VG V_72 Optimisation mémoire
PT178  : 02/05/2007 GGU V_72 Gestion des tables dynamiques avec variable de paie
                            en entrée.
PT179  : 04/05/2007 PH V_72 FQ 13557 Nouvelle variable 0087
PT180  : 25/05/2007 FC V_72 Quand on saisissait une rubrique qui n'existait pas
                            et qu'on lançait le calcul, message d'erreur qui
                            restait même quand la ligne était automatiquement
                            supprimée
PT181  : 14/06/2007 FC V_72 Mettre la date de début du mois dans la date de
                            validité des éléments manquants, passer la clé de
                            PGSYNELTNAT à la fonction getvalueFromDynaTable
PT182  : 21/06/2007 FC V_72 FQ 14435
PT183  : 05/07/2007 PH V_70 FQ 14517 et 12188 Rajout Taux horaire et Heures
                            Mensuelles du contrat
PT184  : 18/07/2007 GGU V_8 Gestion du calcul des variables de présence
PT185  : 08/08/2007 FC V_72 FQ 14615 Ajout variable 0193 Entrée du salarié en
                            cours de mois
PT186  : 21/08/2007 PH V_72 FQ 14661 Ajout variable 0194 Sortie du salarié en
                            cours de mois
PT187  : 28/08/2007 FC V_8  Extraction de fonction dans Ulibeditionpaie
PT188  : 03/09/2007 FC V_80 FQ 14721 Prise en compte de la convention collective
                            pour prédéfini CEG
PT189  : 12/09/2007 PH V_80 FQ 14742 Non initialisation des variables 141 et 142
                            en cas de chgt de salarié
PT190  : 13/09/2007 FC V_80 FQ 14613 Ajout variable 0195 Nombre de salariés
                            handicapés du mois
PT191  : 17/09/2007 PH V_80 FQ 14773 Pb Raz compteur absence V0054 à tord dans
                            le cas CP forcés ds bulletin
PT192  : 02/10/2007 PH V_80 FQ 14818 Erreur SQl suppression bulletin dans le cas
                            de réglement IJSS
PT193  : 04/10/2007 GGU V_80 FQ 14313 Variables avec des éléments dynamiques
PT194  : 05/01/2007 GGU V_80 Gestion des logs et des diagnostiques
PT195  : 16/10/2007 GGU V_80 FQ 14498 Solde du reliquat des prêts saisie arrêt
PT196  : 16/10/2007 PH V_80 FQ 14837 Alimentation Acquis CP dans le bulletin à
                            partir d'une variable qui recherche des rem du
                            bulletin
PT197  : 17/10/2007 MF V_80 mie en place du traitement des jours de
                            fractionnment
PT198  : 26/10/2007 NA V_80 Si Module Presence : Mise à jour de la table
                            PRESENCESALARIE lors du traitement de la paie
PT199  : 30/10/2007 GGU V_80 Gestion des absences en horaire
PT200  : 26/12/2007 PH V_80 FQ 15076 Alimentation des cumuls si rémunérations de
                            régul de bas de bulletin.
PT201  : 27/12/2007 PH V_80 FQ 14705 variable de type cumul avec période de RAZ
                            exemple cumul 11 base CP.
PT202  : 27/12/2007 PH V_80 FQ 15077 Variable 0093 Taux partiel DADS du salarié
PT203  : 27/12/2007 PH V_80 FQ 15077 Variable 0094 Catégorie DUCS du salarié
PT204  : 21/01/2008 FC V_81 FQ 15144 Pb valorisation élément dynamique boite à
                            cocher
PT206  : 03/03/2008 MF V_81 FQ 15276 Modif commentaire maintien
PT218  : 08/09/2008 PH V_81 Optimisation gestion du fractionnement
PT219  : 08/09/2008 JS Optimisation GBL
PT221  : 09/10/2008 FC FQ 15765 Division par 0 ne donnait pas 0 quand il y avait des décimales        
}
unit P5Util;

interface

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Windows, SysUtils, Classes, Graphics, Dialogs,
  StdCtrls, Hent1, HCtrls, HMsgBox, HStatus,
  Grids, Mask, UTOB, P5Def,

  Controls, //PGIAsk
  strutils, //PT181

  M3VM, // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
  //  M3Code, PGVisuObjet,
  ParamSoc, EntPaie,
  // added BY XP le 13-03-2003
  uPaieVariables,
  uPaieRemunerations,
  uPaieCotisations,
  uPaieBases,
  uPaieExecpt,
  uPaieProfilsPaie,
  uPaieCumuls,
  uPaieEltNatDOS,
  uPaieEltNatSTD,
  uPaieEltNatCEG,
  uPaieEltDynSal,//PT172
  uPaieEltDynEtab,//PT172
  uPaieEltDynPop,//PT172
  uPaieEltNiveauRequis,//PT172
  PGTablesDyna,
  PgPresence, //PT184
  UlibEditionPaie, // pt187
{$IFNDEF CPS1}
  PGPOPULOUTILS,//PT172
{$ENDIF}
  uPaieEtabCompl,
  uPaieMinConvPaie,
  uPaieMinimumConv,
  uPaieVentiRem,
  uPaieVentiCot,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbTables{$ELSE}uDbxDataSet{$ENDIF}
{$IFNDEF EAGLSERVER}
  , Fe_Main;
{$ELSE}
  ;
{$ENDIF}
{$ELSE}
  MaineAGL;
{$ENDIF}
const
  SG_Rub: integer = 0;
  SG_Lib: integer = 1;
  SG_Base: integer = 2;
  SG_Taux: integer = 3;
  SG_Coeff: integer = 4;
  SG_Mt: integer = 5;
  SG_TxSal: integer = 3;
  SG_MtSal: integer = 4;
  SG_TxPat: integer = 5;
  SG_MtPat: integer = 6;
  SG_Origine: integer = 6;
  SG_Plfd: integer = 3;
  SG_Tr1: integer = 4;
  SG_Tr2: integer = 5;
  SG_Tr3: integer = 6;
  SG_Plf1: integer = 7;
  SG_Plf2: integer = 8;
  SG_Plf3: integer = 9;

  // Definition de champs mis à jour pendant la saisie du bulletin
type
  TChampBul = record { PT97 Ajout champ }
    Reglt: string;
    DatePai, DateVal: TDateTime;
    HorMod, BasesMod, TranchesMod, TrentMod: Boolean;
    DTrent, NTrent: Integer;
    Ouvres, Houvres, Ouvrables, HOuvrables, HeuresTrav: Double; {PT61 ajout Heures ouvrées & ouvrables}
    // PT34 : 22/03/2002 PH V571  Civilté et mode reglement modifié
    RegltMod: Boolean;
    // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
    Edtdu, Edtau: TDateTime;
    CpAcquisMod: Boolean;
// PT197    CpAcq, CpSupl, CpAnc: Double;
    CpAcq, CpSupl, CpAnc, CpFract: Double;
  end;
  // Definition d'un type record pour calculer toutes les infos à la fois
type
  TRendRub = record
    MontRem,
      BasRem,
      TauxRem,
      CoeffRem,
      BasCot,
      TSal,
      MSal,
      TPat,
      MPat,
      Plfd1,
      Tr1,
      Tr2,
      Tr3,
      Base,
      Plfd2,
      Plfd3: double;
  end;
  // PT32 : 07/02/2002 PH V571 suppression historisation salarié
  {
  Type HistoSal = class //pour l'historisation du salarié
                  salarie       : String;
                  CodeEmploi    : string;
                  LibelleEmploi : string;
                  Qualification : string;
                  Coefficient   : string;
                  Indice        : string;
                  Niveau        : string;
                  CodeStat      : string;
                  TravailN1     : string;
                  TravailN2     : string;
                  TravailN3     : string;
                  TravailN4     : string;
                  Groupepaie    : string;
                  SalaireMois1  : double;
                  SalaireAnn1   : double;
                  SalaireMois2  : double;
                  SalaireAnn2   : double;
                  SalaireMois3  : double;
                  SalaireAnn3   : double;
                  SalaireMois4  : double;
                  SalaireAnn4   : double;
                  Horairemois   : double;
                  Horairehebdo  : double;
                  HoraireAnnuel : double;
                  TauxHoraire   : double;
                  Dtlibre1      : tdatetime;
                  Dtlibre2      : tdatetime;
                  Dtlibre3      : tdatetime;
                  Dtlibre4      : tdatetime;
                  Boollibre1    : string;
                  Boollibre2    : string;
                  Boollibre3    : string;
                  Boollibre4    : string;
                  CSlibre1      : string;
                  CSlibre2      : string;
                  CSlibre3      : string;
                  CSlibre4      : string;
                  Profil        : string;
                  PeriodBul     : string;
  //PT-12 02/10/01 V562 PH Rajout historisation salarie profil rémunération
                  ProfilRem     : string;
  //PT2
                  DADSProfessio : string;
                  DADSCategorie : string;
                  TauxPartiel   : string;
                  Condemploi    : string;//PT13

                  Bsalarie       : Char;
                  BCodeEmploi    : Char;
                  BLibelleEmploi : Char;
                  BQualification : Char;
                  BCoefficient   : Char;
                  BIndice        : Char;
                  BNiveau        : Char;
                  BCodeStat      : Char;
                  BTravailN1     : Char;
                  BTravailN2     : Char;
                  BTravailN3     : Char;
                  BTravailN4     : Char;
                  BGroupepaie    : Char;
                  BSalaireMois1  : Char;
                  BSalaireAnn1   : Char;
                  BSalaireMois2  : Char;
                  BSalaireAnn2   : Char;
                  BSalaireMois3  : Char;
                  BSalaireAnn3   : Char;
                  BSalaireMois4  : Char;
                  BSalaireAnn4   : Char;
                  BDtlibre1      : Char;
                  BDtlibre2      : Char;
                  BDtlibre3      : Char;
                  BDtlibre4      : Char;
                  BBoollibre1    : Char;
                  BBoollibre2    : Char;
                  BBoollibre3    : Char;
                  BBoollibre4    : Char;
                  BCSlibre1      : Char;
                  BCSlibre2      : Char;
                  BCSlibre3      : Char;
                  BCSlibre4      : Char;
                  BProfil        : Char;
                  BPeriodBul     : Char;
  //PT-12 02/10/01 V562 PH Rajout historisation salarie profil rémunération
                  BProfilRem     : Char;
                  BHorairemois   : Char;
                  BHorairehebdo  : Char;
                  BHoraireAnnuel : Char;
                  BTauxHoraire   : Char;
  //PT2
                  BDADSProfessio : Char;
                  BDADSCategorie : Char;
                  BTauxPartiel   : Char;
                  BCondemploi    : Char;//PT13
                  END ;
  }
  // FIN PT32
var
  TOB_Paie, // Tob liste des paies entete
    TOB_RemSal, // Tob liste des rémunérations type Salaires
    TOB_RemAbs,
    TOB_RemAbt,
    TOB_RemAvt,
    TOB_RemCplt,
    TOB_RemPrimes,
    TOB_RemHeures,
    TOB_RemNonImp,
    TOB_RemRet,
{$IFDEF aucasou}
  TOB_ProfilPaies, // Tob liste des Profils
    TOB_Cotisations, // Tob Liste des Cotisations
    TOB_Bases, // Tob Liste des rubriques Base de cotisations
    TOB_ProfilRubs, // Tob liste des rubriques composant les profils
    TOB_Rem, // Tob Liste de TOUTES les rémunérations
    TOB_Variables, // Tob liste des Varaibles
    TOB_Cumuls, // TOb liste des cumuls gérés par la paie
    TOB_Etablissement, // Tob liste des établisssements de la paie pour récupèrer le paramètrage par défaut des profils
    TOB_EltNationauxCEG: tob; // Tob des éléments nationaux CEGID
  TOB_EltNationauxSTD: tob; // Tob des éléments nationaux STANDARD
  TOB_EltNationauxDOS: tob; // Tob des éléments nationaux Dossier
  TOB_EltNiveauRequis: tob; // PT172 Tob des niveaux préconisés des éléments nationaux
{$ENDIF}
  TOB_HistoBasesCot, // TOB historiques des Bases de Cotisation
    TOB_Salarie, // Tob contenant le salarié sur lequel on positionné
    TOB_CumulSal, // Tob des  Historiques cumuls salaries - Cumuls alimentes par les rubriques du bulletin
    TOB_JourTrav, // TOB calendrier des jours travaillés
    TOB_ExerSocial, // TOB des exercices sociaux gérés
  //    TOB_Minimum,             // TOB des tables Dossiers Minimum Conventionnels, Age, Anciennete, Coeff, Qualif, Indices, niveaux
  //    TOB_DetailMin,           // TOB Détails des intervalles des tables dossiers
  TOB_Pris, // TOB Congés payés pris pour le salarié
    TOB_Acquis, // TOB Congés payés Acquis pour le salarié
    TOB_Solde, // TOB Congés payés solde pour le salarié
    TOB_delta, // TOB Congés payés des ajustements négatif { PT146 }
    TOB_ABS, // Tob des absences du salarié
    Tob_VenRem, // Tob des ventilations analytiques des remunerations par defaut
    Tob_VenCot, // Tob des ventilations analytiques des Cotisations par defaut
  //    Tob_VentilRem,           // Tob des preventilations analytiques des remunerations pour recuperer le numero de compte
  //    Tob_VentilCot,           // Tob des preventilations analytiques des Cotisations pour recuperer le numero de compte
  TOBAna, // TOB des ventilations analytiques du bulletin
    TOB_AcquisAVirer, // Tob des acquis à virer
    T_MvtAcquisAVirer, // Tob des acquis en cours à virer
    TOB_PSD, // Tob de HistoSaisRub
  // PT51 : 20/06/2002 PH V582 Traitement des lignes de la saisie des primes
  TOB_PSP, // Tob de HistoSaisPrim
    T_MvtAcquis, // Tob des mvts acquis du bulletin
    Tob_CalendrierSalarie, // Tob calendrier Salarie PT61 nouvel tob
  // PT84   : 26/08/2003 PH V_42 Mise en place saisie arret
  TOB_SaisieArret, // Tob contenant les saisies arrêts
    GblTob_Semaine, { PT130 }
    GblTob_Standard,
    GblTob_VarCp, { PT164 }
    GblTob_JourFerie: TOB; // Tob de la paie en cours de traitement contient Entete + ligne

var
  ForceAlignProfil: Boolean; // PT106 Pour force la la mise à jour des rubriques des profils
  DroitCegPaie: Boolean; // PT102 10/05/2004 PH V_50 restriction droit accès au paramétrage CEGID
  Codsal: string;
  // PT14 : 23/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire
  BullCompl, ProfilSpec, BullContrat: string; // Bulletin compléméntaire X ou -  correspondra à un Boolean dans la table et son profil
  RegimeAlsace: string; // PT100  : 05/04/2004 PH V_50 FQ 11231 Prise en compte du cas spécifique Alsace Lorraine
  TypeTraitement: string; // indique si SAISIE ou PREPA (preparatiàon automatique)
  Trentieme: Double = -1;
  Suivant, OrdreArrBull: integer; // no d'ordre pour insertion mvt ds table absencesalarie
  //OrdreArrBull : numéro ordre de l'arrondi créé ds le bulletin
  TOB_PaieTraite: TOB;
  DateSSalarie: Tdatetime; // date de sortie du salarié
  GrilleBull: THGrid; // Grille active en saisie de bulletin pour recuperer les infos saisies
  TraceERR: TListBox; // Composant Trace gestion des erreurs pour conserver les anamalies de la paie sans arreter le pgm
  ChpEntete: TChampBul; // Definition d'une structure des champs entete mis à jour par la saisie de bulletin
  ActionSLD, ActionBulCP: TActionBulletin;
  Chargement, PP, TopJPAVANTBULL, ACquisMaj: boolean; // AcquisMaj : écriture depuis le load par calculvalocp
  ValoXP0N, ValoXP0D, JPAVANTBULL: double;
  TauxAT: double; // Taux AT constaté sur le bulletin
  HCPPRIS, // Heures CP pris dans la session
    HCPPAYES, // Heures CP payees dans la session
    HCPASOLDER, // Heures CP à solder dans la session
    JCPACQUIS, // Jours CP acquis dans la session
    JCPACQ, // Jours CP acquis dans la session               { PT140 }
    JCPACS, // Jours CP acquis suppl. dans la session        { PT140 }
    JCPACA, // Jours CP acquis Ancienneté dans la session    { PT140 }
    JCPACQUISACONSOMME, // Valeur a consommé sur Jours CP acquis dans la session  // PT17
    JCPACQUISTHEORIQUE, // Jours CP acquis théorique PT75-2
    MCPACQUIS, // Mois acquis ds la session (trentieme)
    BCPACQUIS, // Base acquis ds la session
    JCPPRIS, // Jours CP pris dans la session
    JCPSOLDE, // Jours CP de type solde payé ds la session
    JCPPAYES, // Jours CP payés dans la session
    JCPPAYESPOSES, // Jours CP pris à intégrer dans la session PT88
    JCPASOLDER, // Jours CP A Solder dans la session
    CPMTABS, // CP Montant Absence
    CPMTINDABS, // CP Montant Indemnite Absence
    CPMTABSSOL, // CP Montant Absence Solde
    CPMTINDABSSOL, // CP Montant Indemnite Absence Solde
    CPMTINDCOMP, // CP Montant Indemnite compensatrice
    ASOLDERAVBULL, // CP à solder hormis ceux du bulletin courant
    JRSAARRONDIR, // CP de la période à solder hors acquis bulletin
    JTHTRAVAILLES, // Jours théorique travaillés par le salarié sur la période
    HTHTRAVAILLES, // Heures théoriques travaillées par le salarié sur la période //PT43
    PAYECOUR, // payé sur ce qu'on est en train d'acquérir-permet de positionner
  // APAYE3 en alimcongeacquis
  HABSPRIS, //Heure Absence pris  PT33 Ajout nouvels variable
    JABSPRIS: Double; //Jour Absence pris   PT33
  CalcFractOk : Boolean; //PT218
  SLDCLOTURE: Boolean; //PT75-1 Indique s'il existe un mvt de solde provenant de la clôture CP
  FirstBull: Boolean; { PT101 }
  TopRecalculCPBull: Boolean = False; //PT82-1 Si cumul 12 (maintien) = 0 alors TopRecalculCPBull := True;
  ObjTM3VM: TM3VM; // Objet pour evaluation des variables, utilisation VM du script {$IFNDEF EAGLSERVER} {$ENDIF}
  // PT32 : 07/02/2002 PH V571 suppression historisation salarié
  //    TheHistoSal : HistoSal;
  SoldeDeToutCompte : Boolean; //PT195
  ObjTableDyn : TTablesDynamiques; // Objet gestion des tables dynamiques PT175
  ObjCalcuVarPre : TCalcuVarPre;  //PT184 Objet de calcul des variables de présence
  // Gestion des absences déportées Salarie,Administrateur ,Responsable validation
  // PT48 : 11/06/2002 PH V582 Modif des variables des noms des responsables pour traiter tous les cas
  LeSalarie, LeNomSal, SalAdm, SalVal, GblTypeSal: string;
  NotSalAss, ConsultP: Boolean; // Acces en consultation uniquement
  PopUpSalarie: Boolean; // PT104 Restriction à la consultion des salariés pdt la saisie des bulletin
  // Optimisations
  TOB_DUSALARIE: TOB;
  iPCR_CUMULPAIE, iPCR_SENS, iPCL_ALIMCUMUL, iPCL_ALIMCUMULCOT, iPCL_COEFFAFFECT: Integer;
  iPHC_MONTANT, iPHC_ETABLISSEMENT, iPHC_SALARIE, iPHC_DATEDEBUT, iPHC_DATEFIN, iPHC_REPRISE: Integer;
  iPHC_CUMULPAIE, iPHC_TRAVAILN1, iPHC_TRAVAILN2, iPHC_TRAVAILN3, iPHC_TRAVAILN4, iPHC_CODESTAT: Integer;
  iPHC_CONFIDENTIEL, iPHC_LIBREPCMB1, iPHC_LIBREPCMB2, iPHC_LIBREPCMB3, iPHC_LIBREPCMB4: Integer;
  iPHB_BASEREM, iPHB_MTREM, iPHB_BASECOT, iPHB_MTSALARIAL, iPHB_MTPATRONAL, iPHB_ORIGINEINFO: Integer;
  iPHB_ETABLISSEMENT, iPHB_SALARIE, iPHB_DATEDEBUT, iPHB_DATEFIN, iPHB_NATURERUB, iPHB_RUBRIQUE: Integer;
  iPHB_LIBELLE, iPHB_IMPRIMABLE, iPHB_TAUXREM, iPHB_COEFFREM, iPHB_TAUXSALARIAL, iPHB_TAUXPATRONAL, iPHB_BASEREMIMPRIM: Integer;
  iPHB_TAUXREMIMPRIM, iPHB_COEFFREMIMPRIM, iPHB_BASECOTIMPRIM, iPHB_TAUXSALIMPRIM, iPHB_TAUXPATIMPRIM, iPHB_ORGANISME: Integer;
  iPHB_TAUXAT, iPHB_PLAFOND, iPHB_PLAFOND1, iPHB_PLAFOND2, iPHB_PLAFOND3, iPHB_CONSERVATION, iPHB_ORDREETAT, iPHB_SENSBUL: Integer;
  iPHB_TRAVAILN1, iPHB_TRAVAILN2, iPHB_TRAVAILN3, iPHB_TRAVAILN4, iPHB_CODESTAT, iPHB_CONFIDENTIEL, iPHB_COTREGUL: Integer;
  iPHB_TRANCHE1, iPHB_TRANCHE2, iPHB_TRANCHE3, iPHB_LIBREPCMB1, iPHB_LIBREPCMB2, iPHB_LIBREPCMB3, iPHB_LIBREPCMB4, iPHB_OMTSALARIAL: Integer; { PT136 }
  iPRM_IMPRIMABLE, iPRM_BASEIMPRIMABLE, iPRM_TAUXIMPRIMABLE, iPRM_COEFFIMPRIM, iPRM_ORDREETAT, iPRM_SENSBUL: Integer;
  iPRM_THEMEREM, iPRM_DECBASE, iPRM_DECTAUX, iPRM_DECCOEFF, iPRM_DECMONTANT: Integer;
  iPRM_LIBELLE, iPHB_ORIGINELIGNE, iPRM_LIBCONTRAT: Integer;
  iPCT_IMPRIMABLE, iPCT_BASEIMP, iPCT_TXSALIMP, iPCT_ORGANISME, iPCT_ORDREETAT, iPCT_LIBELLE, iPCT_DU, iPCT_AU: Integer;
  iPCT_MOIS1, iPCT_MOIS2, iPCT_MOIS3, iPCT_MOIS4, iPCT_DECBASE, iPCT_TYPEBASE, iPCT_TYPETAUXSAL, iPCT_TAUXSAL, iPCT_DECTXSAL: Integer;
  iPCT_TYPETAUXPAT, iPCT_TAUXPAT, iPCT_DECTXPAT, iPCT_TYPEFFSAL, iPCT_FFSAL, iPCT_DECMTSAL, iPCT_TYPEFFPAT, iPCT_FFPAT, iPCT_DECMTPAT: Integer;
  iPCT_TYPEMINISAL, iPCT_VALEURMINISAL, iPCT_TYPEMAXISAL, iPCT_VALEURMAXISAL, iPCT_TYPEMINIPAT, iPCT_VALEURMINIPAT: Integer;
  iPCT_TYPEMAXIPAT, iPCT_VALEURMAXIPAT, iPCT_DECBASECOT, iPCT_SOUMISREGUL, iPCT_BASECOTISATION, iPCT_TXPATIMP: Integer;
  iPCT_TYPEPLAFOND, iPCT_PLAFOND, iPCT_TYPETRANCHE1, iPCT_TYPETRANCHE2, iPCT_TYPETRANCHE3, iPCT_TYPEREGUL: Integer;
  iPCT_TRANCHE1, iPCT_TRANCHE2, iPCT_TRANCHE3, iPCT_ORDREAT, iPCT_CODETRANCHE, iPCT_PRESFINMOIS: Integer;
  iCHampPredefini, iCHampDossier, iCHampCodeElt, iCHampDateValidite, iCHampDecaleMois, iChampMontant, iChampMontantEuro, iChampRegimeAlsace: Integer; // PT100
  iCHampConvention:integer; //PT172
  iPSA_SALARIE, iPSA_ETABLISSEMENT, iPSA_TRAVAILN1, iPSA_TRAVAILN2, iPSA_TRAVAILN3, iPSA_TRAVAILN4, iPSA_CODESTAT, iPSA_LIBREPCMB1: Integer;
  iPSA_LIBREPCMB2, iPSA_LIBREPCMB3, iPSA_LIBREPCMB4, iPSA_CONFIDENTIEL, iPSA_TAUXHORAIRE, iPSA_SALAIREMOIS1, iPSA_SALAIREMOIS2, iPSA_SALAIREMOIS3: Integer;
  iPSA_SALAIREMOIS4, iPSA_SALAIREMOIS5, iPSA_SALAIRANN1, iPSA_SALAIRANN2, iPSA_SALAIRANN3, iPSA_SALAIRANN4, iPSA_SALAIRANN5, iPSA_DATENAISSANCE, iPSA_DATEENTREE: Integer;
  iPSA_DATESORTIE, iPSA_TYPPROFILREM, iPSA_PROFILREM, iPSA_TYPPROFIL, iPSA_PROFIL, iPSA_TYPPERIODEBUL, iPSA_PERIODBUL: Integer;
  iPSA_TYPPROFILRBS, iPSA_PROFILRBS, iPSA_TYPREDREPAS, iPSA_REDREPAS, iPSA_TYPREDRTT1, iPSA_REDRTT1, iPSA_TYPREDRTT2, iPSA_REDRTT2: Integer;
  iPSA_PROFILTPS, iPSA_TYPPROFILAFP, iPSA_PROFILAFP, iPSA_TYPPROFILAPP, iPSA_PROFILAPP, iPSA_TYPPROFILRET, iPSA_PROFILRET: Integer;
  iPSA_TYPPROFILMUT, iPSA_PROFILMUT, iPSA_TYPPROFILPRE, iPSA_PROFILPRE, iPSA_TYPPROFILTSS, iPSA_PROFILTSS, iPSA_TYPPROFILCGE, iPSA_PROFILCGE: Integer;
  iPSA_PROFILCDD, iPSA_PROFILMUL, iPSA_TYPPROFILFNAL, iPSA_PROFILFNAL, iPSA_TYPPROFILTRANS, iPSA_PROFILTRANS, iPSA_TYPPROFILANC, iPSA_PROFILANCIEN: Integer;
  iPSA_LIBELLE, iPSA_PRENOM, iPSA_NUMEROSS, iPSA_ADRESSE1, iPSA_ADRESSE2, iPSA_ADRESSE3, iPSA_CODEPOSTAL, iPSA_VILLE, iPSA_INDICE, iPSA_NIVEAU, iPSA_CONVENTION: Integer;
  iPSA_CODEEMPLOI, iPSA_AUXILIAIRE, iPSA_DATEANCIENNETE, iPSA_QUALIFICATION, iPSA_COEFFICIENT, iPSA_LIBELLEEMPLOI, iPSA_CIVILITE, iPSA_CPACQUISMOIS, iPSA_NBREACQUISCP: Integer;
  iPSA_TYPDATPAIEMENT, iPSA_MOISPAIEMENT, iPSA_JOURPAIEMENT, iPSA_TYPREGLT, iPSA_PGMODEREGLE, iPSA_REGULANCIEN, iPSA_HORHEBDO, iPSA_HORAIREMOIS, iPSA_HORANNUEL: Integer;
  iPSA_PERSACHARGE, iPSA_PCTFRAISPROF, iPSA_MULTIEMPLOY, iPSA_SALAIREMULTI, iPSA_ORDREAT, iPSA_SALAIRETHEO, iPSA_DATELIBRE1, iPSA_DATELIBRE2, iPSA_DATELIBRE3: Integer;
  iPSA_DATELIBRE4, iPSA_VALANCCP, iPSA_ANCIENNETE, iPSA_CALENDRIER, iPSA_STANDCALEND, iPSA_BOOLLIBRE1, iPSA_BOOLLIBRE2, iPSA_BOOLLIBRE3, iPSA_BOOLLIBRE4: Integer;
  iPSA_DADSPROF, iPSA_DADSCAT, iPSA_TAUXPARTIEL, iPSA_CPTYPEMETHOD, iPSA_VALORINDEMCP, iPSA_CPTYPEVALO, iPSA_MVALOMS, iPSA_VALODXMN, iPSA_CPACQUISANC, iPSA_BASANCCP: Integer;
  iPSA_DATANC, iPSA_TYPDATANC, iPSA_DATEACQCPANC, iPSA_NBRECPSUPP, iPSA_CPTYPERELIQ, iPSA_RELIQUAT, iPSA_SORTIEDEFINIT, iPSA_SEXE, IPSA_CONDEMPLOI: Integer;
  iPSA_CONGESPAYES, iPSA_CPACQUISSUPP, iPSA_ANCIENPOSTE, iPSA_TYPPAIEVALOMS, iPSA_PAIEVALOMS, iPSA_UNITEPRISEFF, iPSA_BULLCONTRAT, iPSA_ACTIVITE, iPSA_CATDADS : integer; { PT100-1 & PT100-2, PT150, PT203}
  IntegreAuto: Boolean; // PT166
  iCHampPEDCodeElt, iCHampPEDDateValidite, iCHampPEDDecaleMois, iChampPEDMontant, iChampPEDMontantEuro, iChampPEDRegimeAlsace, iChampPEDValeurNiveau: Integer; //PT172
  iCHampPNRCodeElt, iCHampPNRNiveauRequis, iCHampPNRNiveauMaxi: Integer; //PT172
  stNiveauRequis,stNiveauMaxi : String;//PT172
  TobAcSaisieArret: Tob; //PT219
// optimisation
procedure MemorisePsa(Unsal: TOB);
procedure MemorisePrm(UneRem: TOB);
procedure MemorisePel(UnElt: TOB);
procedure MemorisePed(UnElt: TOB);  //PT172
procedure MemorisePnr(UnElt: TOB);  //PT172
procedure MemorisePct(UneCot: TOB);
procedure MemorisePcl(TRechCum: TOB);
procedure MemorisePcr(TRechRub: TOB);
procedure MemorisePhb(THB: TOB);
procedure MemorisePhc(UneTob: TOB);
Procedure InitMemorise;
function NumChampTobS(const zz: string; const Ind: Integer): Integer;
function NumChampProfS(const Ch: string): Integer;
// Fin optimisation

// Charge l'ensemble des rubriques composant le bulletin en fonction des profils du salarié
// PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
function ChargeRubriqueSalarie(Salarie: TOB; DateDeb, DateFin: TDateTime; ActionB: TActionBulletin; MtEnvers: Double; TheProfilPart: string): TOB;
procedure ChargeLesTOBPaie; // charge toutes les TOB pour le calcul du bulletin
procedure ChargePresence(FiltreSalarie : String = ''; DateDebut : TDateTime = 0; DateFin : TDateTime = 0; DateFinEstFinDeMois : Boolean = False); //PT184
procedure ChargeLesExercPaie(const DateDeb, DateFin: TDateTime); // Idnetification de l'exercice social sur lequel on travaille
procedure VideLesTOBPaie(AvecInit: Boolean); // detruit les TOB pour le calcul du bulletin
procedure VideLaPresence; // PT184
procedure VideLaTobExer; // detruit la tob contenant l'exercice social
procedure InitLesTOBPaie; // Initialisation des TOB pour le calcul du bulletin
function ValEltNat(const Elt: string; const DatVal: TDateTime; Diag: TObject = nil): double; // Rend Valeur Element National
function ValEltDyna(TOBSalarie : Tob; Date : TDateTime; Elt : String; var DateValidite : TDateTime) : Double;
function ValCumulDate(const Cumul: string; const DateDebut, DateFin: TDateTime): double; // Rend la valeur d'un cumul de date à date
function ValRubDate(const Rubrique: string; Nature: string; const DateDebut, DateFin: TDateTime): TRendRub; // idem sur une rubrique
function TauxHoraireSal(): double; // Rend Taux Horaire fiche Salarié
function SalaireMensuelSal(): double; // Rend Salaire Mensuel Fiche Salarie
function SalaireAnnuelSal(): double; // Rend Salaire Annuel Fiche Salarié
function CalculTrentieme(const DateDebut, DateFin: TDateTime; TestPaie: Boolean = FALSE): Integer; // Calcul le numerateur du trentieme
// pt187 function EstDebutMois(const LaDate: TDateTime): boolean; // Indique si la date est le début d'un mois
// pt187 function EstFinMois(const LaDate: TDateTime): boolean; //  Indique si la date est la fin d'un mois
// pt187 function DiffMoisJour(const DateDebut, DateFin: TDateTime; var NbreMois, NbreJour: Integer): Boolean; // Calcule difference de date à date en mois et jours restants
function AncienneteAnnee(const DateEntree, DateFin: TDateTime): WORD; // Rend le nbre Annee Anciennete entre Date Entree et date du jout
function AncienneteMois(const DateEntree, DateFin: TDateTime): WORD; // Rend le nbre Mois Anciennete entre Date Entree et date du jout
procedure AgeSalarie(const DateFin: TDateTime; var Annee, Mois: WORD); // Rend Age du Salarie en Mois et en Annee
function RendCumulSalSess(const Cumul: string): double; // rend la valeur du cumul salarié calculé dans la session de paie
procedure RechercheProfil(const Champ1, Champ2: string; Salarie, T_Etab, TPE: TOB); // Recherche en cascade d'un profil à partir du salarié à l'établissement
procedure MajRubriqueBulletin(TPE, TPR: TOB; const CodSal, Etab: string; const DateDeb, DateFin: TDateTime); // recup des rubriques du bulletin précédent et des éléments permanents
procedure RecupEltPermanent(TPE, TPR: TOB; const Rubrique: string; Conservation: string = '');
// PT95 // Recupération des éléments permanents du bulletin précédent sur les rémunérations
procedure AligneProfil(Salarie, T_Etab, TPE: TOB; ActionB: TActionBulletin); // Aligne les lignes du bulletin sur les profils du salarié
procedure AlimCumulSalarie(THB: TOB; const Salarie, NatureRub, Rubrique, Etab: string; const DateDebut, DateFin: TDateTime);
// Alimentation des cumuls Salariés pour unn salarié,une rubrique
procedure RazTPELignes(TPE: TOB); // Fonction qui détruit les TOB filles de TPE (destructioon de la TOB des lignes Bulletins)
procedure GrilleAlimLignes(TPE: TOB; Etab, Salarie, Nature: string; DateDebut, DateFin: TDateTime; GrilleBulletin: THGrid; ActionBul: TActionBulletin);
// Fonction qui alimente les lignes de bulletins en fonction de la grille de Saisie
procedure CreationTOBCumSal; // Creation de la TOB des Cumuls Salaries gérés dans la paie en cours
procedure DestTOBCumSal; // Fonction de destruction de la  TOB des Cumuls Salariés
function OkCumSal: Boolean; // Fonction qui indique si les cumuls salariés ont été calculés
function DoubleToCell(const X: Double; const DD: integer): string; // Fonction d'affchage d'un montant en fonction du nombre de décimales gérées pour le champ
function ValoriseMt(const cas: WORD; const Cc: string; const NbreDec: Integer; Base, Taux, Coeff: Double): Double; // Fonction de valorisation d'une ligne de rémunération
function EvalueRem(Tob_Rub: TOB; const Rub: string; var Base, Taux, Coeff, Montant: Double; var lib: string; DateDeb, DateFin: TDateTime; ActionCalcul: TActionBulletin; Ligne:
  Integer; Diag: TObject = nil): Boolean; // Calcul de la rémunération
function EvalueChampRem(Tob_Rub: TOB; const TypeChamp, Champ, Rub: string; const DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double; // Calcul d'un champ d'une Remunération
function RechCasCodeCalcul(const Cc: string): WORD; // Recherche du cas code calcul
function ValVariable(const VariablePaie: string; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil): double; // Fonction d'évaluation d'une variable
function RechCommentaire(const Rub: string): Boolean; // Fonction qui indique si on c'est une ligne de commentaire associée à la rubrique
function RendCodeRub(const Rub: string): string; // Fonction qui rend le code de la rubrique  en supprimant .x pour les lignes de commentaire
procedure AlimChampEntet(const Salarie, Etab: string; const DateDebut, DateFin: TDateTime; Tob_Rub: TOB);
// Fonction alimentation enreg de PAIEENCOURS avec les champs salarié et la saisie
function RendRubrCommentaire(const rub: string; const CurG: THGrid): Integer; // Fonction qui calcule automatiquement le numéro de la ligne de commentaire
function RecupRem(const ActionCalcul: TActionBulletin; const TypeChamp: string): Boolean; // Fonction qui indique si doit calculer ou récuperer le champ de la rémunération
function EvalueCot(Tob_Rub: TOB; const Rub: string; var Base, TxSal, TxPat, MtSal, MtPat: Double; var lib: string; DateDeb, DateFin: TDateTime; ActionCalcul: TActionBulletin; var
  At: Boolean; Diag: TObject = nil): Boolean; // Calcul de la Cotisation
function EvalueChampCot(Tob_Rub: TOB; const TypeChamp, Champ, Rub: string; const DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double;
//Fonction d'évaluation d'un champ d'une cotisation en fonction du type du cham
{$IFNDEF EMANAGER}
//PT148 ajout paramètre Calcul du maintien (booléen)
//PT162  ajout parmètre Calcul paie à l'envers (booléen)
procedure CalculBulletin(Tob_Rub: TOB; CalcMaint: Boolean = FALSE; Diag: TObject = nil; PaieEnvers: Boolean = FALSE); // procedure qui valorise les lignes de bulletins en fonction du paramètrage et A PARTIR de TOB_Rub

{$ENDIF EMANAGER}
function CalculDatePaie(TOB_Sal: TOB; DatePaie: TDateTime): TDateTime; // fonction de calcul de la date de la paie
procedure AlimProfilMulti(ListeProfil: string; Salarie, TPE: TOB); //  Fonction d'alimentation en rubriques en fonction des profils tq maladie, chomage
function RechVarDefinie(const VariablePaie: string; const DateDebut, DateFin: TDateTime; var MtCalcul: Variant; TOB_Rub: TOB; Diag: TObject = nil): Boolean; //PT172
function RechVarPresenceDefinie(const VariablePaie: string; const DateDebut, DateFin: TDateTime; var MtCalcul: Variant; Diag: TObject = nil): Boolean; //PT172
procedure MemoriseTrentieme(const MtTrent: Double); // Memorisation du trentieme pour les calculs et pour la variable qui rend le trentieme calculé
//PT 169 TEST GGU
procedure RecupTobSalarie(const Salarie: string; DateDebut, DateFin : TDateTime); // Rechargement de la TOB_Salarié ¨PT173
//procedure RecupTobSalarie(const Salarie: string; DtDebut : TDateTime = 0; DtFin: TDateTime = 0);
function ExamCasValeurRem(const ACol: Integer; T1: TOB): Boolean; //fonction qui examine si le champ est du type valeur alors jamais de saisie
function EvalueVarVal(T_Rech: TOB; const DateDebut, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double; // fonction evaluation variable de type valeur
//PT184 Gestion des variables de présence
function EvalueVarPres(TypeBase : Integer; Salarie, LaBase : String; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = '') : Double; //TobVar : Tob;

function EvalueVarCum(T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double; // fonction evaluation variable de type cumul
//PT184
function EvalueVarCumulPresence(T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double; // fonction evaluation variable de type cumul
procedure RendDateVar(T_Rech: TOB; const DateDeb, DateFin: TDateTime; var ZdatD, ZdatF: TDateTime);
// fonction qui calcule les dates de début et de fin pour les cumusls,cotisations, rem...
function EvalueVarRub(const Nature: string; T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; TypeB, ChampBase: string; Diag: TObject = nil): Double;
// fonction evaluation variable de type rubrique
function EvalueUnChampVar(const TypeB: string; ChampBase: string; TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; Diag: TObject = nil; VariablePaie: string = ''): Variant;
function IsOperateurInegalite(const Operateur: string): boolean; //
function IsOperateurLogique(const Operateur: string): boolean;
function EvalueExpVar(TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; const Ztype, ZBase, Zoperat: string; const maxO: Integer; Diag: TObject = nil; VariablePaie: string = ''): string;
// evalue une expression de type variable
function EvalueVarCalcul(TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; const DebI, FinI: integer; const Ztype, ZBase, Zoperat: string; Diag: TObject = nil; VariablePaie: string = ''): string;
// evalue une expression de type calcul
function EvalueBas(Tob_Rub: TOB; const Rub: string; var Base, Plafond, Plf1, Plf2, Plf3, TR1, TR2, TR3: Double; var lib: string; const DateDeb, DateFin: TDateTime; ActionCalcul:
  TActionBulletin; const BaseForcee, TrancheForcee: Boolean; Diag: TObject = nil): Boolean; // Calcul de la Base de cotisation
procedure CalculCumulsRegul(const Rub: string; var Base, Plf1, Plf2, Plf3, TR1, TR2, TR3: Double; DateDeb, DateFin: TDateTime; const ARegulariser: Boolean; const TPlf1, TPlf2,
  TPlf3: string); // fonction de calcul des reguls
function EvalueChampBas(Tob_Rub: TOB; TypeChamp, Champ, Rub: string; DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double; // evalue un champ d'une base de cotisations
function EvaluePlafond(TOB_Rub: TOB; TypeChamp, Champ, Rub: string; Plafond: Double; DateDeb, DateFin: TDateTime; NbreDec: Integer; TypeRegul: string; Diag: TObject = nil; Quoi: string = ''): Double;
// Fonction de calcul des différents plafond associés aux tranches
function ValVarTable(const TypeVar: string; var TDossier: string; const DateDebut, DateFin: TDateTime; TOB_Rub: TOB; Origine: string = ''): Double;
function RendValeurTable(const AgeAnc: string; const TypeVar, Champ, Conv, Sens, NatureRes: string; const DateDebut, DateFin: TDateTime; TOB_RUB: TOB; Period: string = ''):
  Double;
function TestSensTable(AgeAnc, NombreLu: Variant; Sens: string; Period: string = ''): Boolean;
function ValBase(const Rubrique: string): TRendRub;
function RendBaseRubBase(const Rubrique: string; TOB_Rub: TOB; Quoi: Integer = -1): Double;
procedure ChargeBasesSal(const Salarie, Etab: string; DateD, DateF: TDateTime); // Chargement des bases de cotisation pour le calcul des réguls
function RendModeRegle(TOB_Sal: TOB): string; // Rend le mode de réglements salariés
procedure RendDatePaiePrec(var Date1, Date2: TDateTime; const DD, DF: TDateTime); // Dates de la paie précédentes
procedure Appellesalaries; // Fonctions pour lancer les fiches AGL du zoom
procedure AppelleCotisations;
procedure AppelleVariables;
procedure AppelleEtablissements;
procedure AppelleRemunerations;
procedure AppelleDossier;
procedure AppelleTablesdossier;
procedure AppelleCumuls;
procedure AppelleProfils;
procedure SuppCotExclus(const ThemeExclus: string; TPE: TOB); // Suppression des rubriques ayant un theme de la liste des rubriques déjà identifiées ex : URSSAF
procedure EvalueChaineDiv0(var St: string); // Controle chaine dans laquelle il y a une division par 0
procedure IntegrePaye(Tob_Rub, Tob_prisCP: tob; const Etab, Salarie: string; const DateD, DateF: TDateTime; const Typecp: string; ManqueAcquis: Boolean = False);
function PGRecupereProfilCge(const Etab: string): string; //PT90
procedure SalEcritCP(var Tob_cppris: tob);
procedure LibereTobCp(var Tob_cppris: tob);
procedure SalEcritAbs(var Tob_abs: tob);
function SalIntegreCP(TOB_Sal, TOB_Rub, Tob_AcqEnCours: TOB; const DateD, DateF: TDateTime; const Auto: Boolean; var StMsgErr: string): tob; //PT7
procedure GridGriseCell(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas); //
procedure GridDeGriseCell(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas);
function IsOkType(TR: tob; const typecp: string): boolean;
function RendBasePrecarite(const DateDebut, DateFin: TDateTime; const DuSalarie: string): double;
procedure IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, salarie: tob; const DateD, DateF: tdatetime; const Action: string);
// PT107 IJSS procedure IntegreRub(tob_rub, Salarie: tob; const TypeAbs, aliment: string; const DateD, DateF: tdatetime; const compteur: double; natrub, arub: string);
procedure IntegreRub(tob_rub, Salarie: tob; const TypeAbs, aliment: string; const DateD, DateF: tdatetime; const compteur: double; natrub, arub: string; IJSS, Maintien: Boolean);
function IntegreAbsInRub(tob_rub, Tob_ParSal: TOB; const natrub, arub: string): Tob;
procedure IntegreRubBul(Salarie, TPE: TOB; Rubrique, Natrub: string);
procedure Ecritlignecomm(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const Rub, Natrub: string);
procedure RechercheCarMotifAbsence(const Abs: string; var ProfilH, RubriqueH, AlimentH, ProfilJ, RubriqueJ, AlimentJ: string; var GereComm, Heure, NetH, NetJ: boolean); { PT153 }
procedure PositionnePaye(t: tob; const DateF: TDateTime);
procedure AnnuleAbsenceBulletin(const CodeSalarie: string; const DateF: tdatetime);
function IsOkFloat(const val: string): boolean;
function IsNumerique(const S: string): boolean;
function ChercheNatRub(Salarie, TPE: TOB; const Profil, rub: string): string;
procedure EnleveCommAbsence(Salarie, Tob_rub: tob; const Arub, Natrub: string; const dated, datef: tdatetime);
procedure EnleveRubAbsence(Salarie, Tob_rub: tob; const dated, datef: tdatetime);

procedure IntegreMvtSolde(Tob_Rub, TOB_Salarie: tob; const Etab, Salarie: string; const DateD, DateF: TdateTime);

procedure recalculeValoSolde(const DateD, DateF: tdatetime; P5Etab, TSal, Tob_solde, Tob_rub: tob; const Salarie: string);
//procedure  ChargeACquis(periode:integer;DateD,DateF,DateDP:tdatetime;Salarie:string);
function notVide(T: tob; const MulNiveau: boolean): boolean;
procedure GestionCalculAbsence(tob_rub: tob; const DateD, DateF: tdatetime; const typem: string);
//procedure SupprimeAbsenceDePaye(Tob_Rub, Tob_ABSTEMP,Salarie:tob;DateD, DateF:tdatetime) ;
procedure RemplirHistoBulletin(HistoBul, Salarie, Rubrique, ProfilRub: TOB; const DateDeb, DateFin: TDateTime);
procedure InitialiseVariableStatCP(AvecAbs : Boolean = TRUE);
procedure InitVarTHTRAVAILLES; //PT83-3
procedure ReaffecteDateMvtcp(const DateFin: tdatetime);
// PT35 : 26/03/2002 PH V571 Fonction de raz de ChptEntete
//PT68    19/12/2002 PH V591 Affectation de la date de paiement à la date de fin de paie si non renseignée
procedure RazChptEntete(const DateF: TDateTime);

//PT32 : 07/02/2002 PH V571 suppression historisation salarié
{
function  IncrementeSeqNoOrdrePHS(cle_sal: string): Integer;
procedure InitialiseBoolean(var M : Histosal;ValB : boolean);
procedure CopyHistosal(var A,N:Histosal);
}
// FIN PT32
function RecupCot(const ActionCalcul: TActionBulletin; const TypeChamp: string): Boolean;
// PT20 : 20/11/01 V562 PH Modif du calcul des dates d'édition automatique en fonction de l'établissement
procedure CalculeDateEdtPaie(var LDatD, LDatF: TDateTime; T_Etab: TOB);
// PT28 : 27/12/2001 V571 PH Calcul des variables de tupe cumul en tenant compte du mois de RAZ
// PT70 : 23/01/2003 PH V591 Correction du calcul des bornes de dates pour la valorisation d'un cumul antérieur avec une période de raz
procedure BorneDateCumul(const DateDeb, DateFin: TDateTime; var ZDatD, ZDatF: TDateTime; const PerRaz: string);
procedure RendBorneDateCumul(const DateDeb, DateFin: TDateTime; var ZDatD, ZdatF: TDateTime; const PerRaz: string);
//PT60 07/10/2002 PH V585 Recherche d'un element national par préférence DOS,STD,CEG
function ValEltNatPref(const Elt, Pref: string; const DatTrait: TDatetime): double;
//PT172 Procédure voisine de ValEltNatPref mais spécifique pour les éléments dynamiques dossier
function ValEltNatDynDos(const Elt, Pref: string; const DatTrait: TDatetime): double;
//PT172 Fonction qui renvoie la valeur du paramètre dont on vient de demander la saisie
function DemandeSaisie(TypeSaisie:String;Elt:STring;NiveauRequis:String;DatVal:TDateTime) : Double;
//PT172 Procédure qui va mémoriser l'élément national à saisir dans la table temporaire PGSYNELTNAt
procedure MemorisationElt(Elt:STring;NiveauRequis:String;DatVal:TDateTime);
//PT172 Fonction qui renvoie le niveau requis pour un élément national
procedure RecupNiveauRequis(const Elt:STring);
procedure LibereTobAbs;
procedure LibereTobCalendrier; { PT158 }

function Paie_RechercheOptimise(const TobSrc: tob; const Field: string; const Valeur: Variant): TOB;
function Paie_RechercheRubrique(const typerubrique, rubrique: string): tob;
procedure RAZTOBPaie();

{$IFNDEF CCS3}
procedure AnnuleIJSSBulletin(const CodeSalarie: string; const DateF: tdatetime); //PT103 IJSS
{$ENDIF}
{$IFNDEF CCS3}
procedure AnnuleMaintienBulletin(const CodeSalarie: string; const DateD, DateF: tdatetime); // PT121 Maintien
{$ENDIF}
{$IFNDEF CCS3}
// d PT124
procedure EcritCommMaint(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const RubMaintien, NatRub: string);
// f PT124
// d PT129
procedure EcritCommIJ(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const RubIJ, NatRub: string);
procedure EnleveCommIJ(Salarie, Tob_rub: tob; const Arub, Natrub: string; const dated, datef: tdatetime);
// f PT129
{$ENDIF}

procedure Align_Execpt(NatureRub, Rub: string); // Gestion de la personnalisation
procedure Charg_TauxAt; // Gestion de la TOB des Taux AT
procedure ChargeCotisations;

// d PT159
procedure EnleveRubIJS(Salarie, Tob_rub: tob; const dated, datef: tdatetime);
procedure EnleveRubMAI(Salarie, Tob_rub: tob; const dated, datef: tdatetime);
// f PT159

//PT172
function RecupValStd(EltNat:Tob;Convention:String;Pref:String;Elt:String;DatTrait: TDatetime) : Double;

//PT194 : Gestion des logs et des diagnostiques
Procedure LogMessage(Diag : TObject; Msg : String; NatureElementCalcule : String = ''; Code : String = ''; NonValueMsg : String = ''; NonValue : Boolean = False; Niveau : Integer = -1; stType : String = ''; stOrigine : String = '');
Procedure LogSaveToFile(Filename : TFileName; Diag : TObject; Entete : TStringList); //PT194
Procedure LogClear(Diag : TObject); //PT194
Function LogToStrings(Diag : TObject) : TStringList; //PT194
Function LogGetChildLevel(Diag : TObject) : TObject; //PT194

// PT198
procedure MiseAjourPresence(DateD, DateF : TDateTime; Salarie, typetraitement : String);

implementation

uses
  PgCongesPayes,
  PgOutils2,
  P5RecupInfos,
{$IFNDEF EMANAGER}
  SaisBul,
{$ENDIF EMANAGER}
  PGOutilsHistorique,
  PGCalendrier, DB, Math, DateUtils
  ,ComCtrls//PT194
  ;

var
  TOB_TauxAT: tob; // TOB des Taux AT de la Société
  TOB_PaiePrecedente: tob; // Tob Liste des rubriques du bulletin précédent
{$IFDEF EMANAGER}
  ValidationOK:Boolean;
  ClePGSynEltNAt:String;
{$ENDIF EMANAGER}

//PT194
{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/10/2007
Modifié le ... :   /  /    
Description .. : Procedure qui permet d'ajouter une ligne de log à un 
Suite ........ : récepteur de message (Diag).
Suite ........ : Diag est le récepteur de log. Il peut etre de type
Suite ........ : TTreeView, TTreeNode, TListBox ou Tob.
Suite ........ : Msg contient le message à ajouter
Suite ........ : NatureElementCalcule contient le code de la nature de
Suite ........ : l'élément produisant le message
Suite ........ : Code contient le Code de l'élément produisant le message
Suite ........ : NonValueMsg contient le message a afficher si on fait un 
Suite ........ : simple diagnostique (sans calcul des valeurs)
Suite ........ : Si NonValue = True, on affiche le message de 
Suite ........ : NonValueMsg
Suite ........ : Niveau permet de gérer des tabulations dans les récepteurs 
Suite ........ : de log de type TTreeView ou TListBox.
Mots clefs ... : 
*****************************************************************}
Procedure LogMessage(Diag : TObject; Msg : String; NatureElementCalcule : String = ''; Code : String = ''; NonValueMsg : String = ''; NonValue : Boolean = False; Niveau : Integer = -1; stType : String = ''; stOrigine : String = '');
var
  TempStr : String;
  indexNiveau : Integer;
  TempNodes : TTreeNodes;
  TempNode  : TTreeNode;
  Temptob : Tob;
  StMessage : String;
begin
  if not Assigned(Diag) then exit;
  if (not NonValue) or (NonValueMsg = '')  then
    StMessage := Msg
  else
    StMessage := NonValueMsg;
  if Diag is TTreeView then
  begin
    StMessage := TrimLeft(StMessage);
    if StMessage = '' then exit;
    TempNodes := (Diag as TTreeView).Items;
    if TempNodes.count > 0 then
      TempNode := TempNodes.Item[TempNodes.count -1]
    else
      TempNode := TTreeNode.Create(TempNodes);
    if Niveau >= 0 then
    begin
      while TempNode.Level < Niveau do
      begin
        if not TempNode.HasChildren then
          TempNode := TempNodes.AddChild(TempNode, '');
        TempNode := TempNode.GetLastChild;
      end;
    end;
    TempNodes.Add(TempNode, StMessage);
  end else if Diag is TListBox then
  begin
    if Niveau >= 0 then
    begin
      StMessage := TrimLeft(StMessage);
      for indexNiveau := 0 to Niveau do
      begin
        //On indente le début du message
        TempStr := TempStr + ' ';
      end;
    end;
    TempStr := TempStr + StMessage;
    (Diag as TListBox).Items.Add(TempStr);
  end else if Diag is TTreeNode then
  begin
    StMessage := TrimLeft(StMessage);
    if StMessage = '' then exit;
    if (Diag as TTreeNode).Text  = '' then
      (Diag as TTreeNode).Text := StMessage
    else
      (Diag as TTreeNode).Owner.Add((Diag as TTreeNode), StMessage);
  end else if Diag is Tob then
  begin
    StMessage := TrimLeft(StMessage);
    if StMessage = '' then exit;
    if not (Diag as Tob).FieldExists('ARP_MESSAGE') then // (Diag as Tob).GetString('ARP_MESSAGE') = ''
    begin
      (Diag as Tob).AddChampSupValeur('ARP_CODE', Code);
      (Diag as Tob).AddChampSupValeur('ARP_NATUREARCHI', NatureElementCalcule);
      (Diag as Tob).AddChampSupValeur('ARP_MESSAGE', StMessage);
      (Diag as Tob).AddChampSupValeur('ARP_TYPENATUREARCHI', stType);
      (Diag as Tob).AddChampSupValeur('ARP_ORIGINE', stOrigine);
    end else begin
      Temptob := Tob.Create((Diag as Tob).NomTable, (Diag as Tob), -1);
      Temptob.AddChampSupValeur('ARP_CODE', Code);
      Temptob.AddChampSupValeur('ARP_NATUREARCHI', NatureElementCalcule);
      Temptob.AddChampSupValeur('ARP_MESSAGE', StMessage);
      Temptob.AddChampSupValeur('ARP_TYPENATUREARCHI', stType);
      Temptob.AddChampSupValeur('ARP_ORIGINE', stOrigine);
    end;
  end else
  begin
    raise Exception.Create('Classe non valide');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Procedure qui permet de sauver dans un fichier de type
Suite ........ : texte les messages contenus dans le récepteur de logs Diag
Mots clefs ... :
*****************************************************************}
Procedure LogSaveToFile(Filename : TFileName; Diag : TObject; Entete : TStringList);
var
  F: TextFile;
  i, IndexString: Integer;
  TempString : TStringList;
begin
  if Filename = '' then exit;
  if not Assigned(Diag) then exit;
  if not FileExists(FileName) then
  begin
    AssignFile(F, FileName);
{$I-}
    ReWrite(F);
{$I+}
    if IoResult <> 0 then
    begin
      PGIBox('Fichier inaccessible : ' + Filename, 'Abandon du traitement');
      Exit;
    end;
    closeFile(F);
  end;
  AssignFile(F, FileName);
  reset(F);
  ReWrite(F);
  if Assigned(Entete) then
  begin
    for IndexString := 0 to Entete.Count -1  do
    begin
      Writeln(F, Entete.Strings[IndexString]);
    end;
    Writeln(F, ' ');
  end;
  TempString := LogToStrings(Diag);
  for i := 0 to TempString.Count -1 do
  begin
    Writeln(F, TempString.Strings[i]);
  end;
  FreeAndNil(TempString);
  CloseFile(F);
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/10/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de vider le récepteur de log (Diag) de 
Suite ........ : tous ses messages
Mots clefs ... : 
*****************************************************************}
Procedure LogClear(Diag : TObject);
var
  Nb, i : Integer;
begin
  if not Assigned(Diag) then exit;
  if Diag is TTreeView then
  begin
    Nb := (Diag as TTreeView).Items.Count;
    for i := Nb -1 downto 0 do
    begin
      (Diag as TTreeView).Items.Item[i].DeleteChildren;
      (Diag as TTreeView).Items.Item[i].Delete;
    end;
  end else if Diag is TListBox then
  begin
    (Diag as TListBox).Clear;
  end else if Diag is TTreeNode then
  begin
    (Diag as TTreeNode).DeleteChildren;
    (Diag as TTreeNode).Delete;
  end else if Diag is Tob then
  begin
    (Diag as Tob).Detail.Clear;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Fonction qui rends la liste des messages contenus dans le 
Suite ........ : conteneur de log sous forme de TStringList
Mots clefs ... : 
*****************************************************************}
Function LogToStrings(Diag : TObject) : TStringList;
var
  Nb, i : Integer;
  TempString : TStringList;
  Function NodeToStrings(Noeud : TTreeNode) : TStringList;
  var
    i : Integer;
    tabSt : String;
    TempString : TStringList;
    function TabNode(Noeud : TTreeNode) : String;
    var
      i : Integer;
    begin
      for i := 0 to Noeud.Level do
        result := result + '  ';
    end;
  begin
    result := TStringList.Create;
    if Noeud.Text <> '' then
    begin
      tabSt := TabNode(Noeud);
      result.Add(tabSt + Noeud.Text);
    end;
    for i := 0 to Noeud.Count -1 do
    begin
      TempString := NodeToStrings(Noeud.Item[i]);
      result.AddStrings(TempString);
      FreeAndNil(TempString);
    end;
  end;
  Function TobToStrings(LaTob : Tob) : TStringList;
  var
    i : Integer;
    tabSt : String;
    TempString : TStringList;
    function TabTob(LaTob : Tob) : String;
    var
      i : Integer;
    begin
      for i := 0 to LaTob.Niveau do
        result := result + '  ';
    end;
  begin
    result := TStringList.Create;
    if (LaTob.FieldExists('ARP_MESSAGE')) and (LaTob.GetString('ARP_MESSAGE') <> '') then
    begin
    tabSt := TabTob(LaTob);
    result.Add(tabSt + LaTob.GetString('ARP_MESSAGE'));
    end;
    for i := 0 to LaTob.Detail.Count -1 do
    begin
      TempString := TobToStrings(LaTob.Detail.Items[i]);
      result.AddStrings(TempString);
      FreeAndNil(TempString);
    end;
  end;
begin
  result := TStringList.Create;
  if not Assigned(Diag) then exit;
  if Diag is TTreeView then
  begin
    Nb := (Diag as TTreeView).Items.Count;
    for i := 0 to Nb -1 do
    begin
      TempString := NodeToStrings((Diag as TTreeView).Items.Item[i]);
      Result.AddStrings( TempString );
      FreeAndNil(TempString);
    end;
  end else if Diag is TListBox then
  begin
    Nb := (Diag as TListBox).Items.Count;
    for i := 0 to Nb - 1 do
      Result.Add((Diag as TListBox).items[i])
  end else if Diag is TTreeNode then
  begin
    TempString := NodeToStrings((Diag as TTreeNode));
    Result.AddStrings( TempString );
    FreeAndNil(TempString);
  end else if Diag is Tob then
  begin
    TempString := TobToStrings((Diag as Tob));
    Result.AddStrings( TempString );
    FreeAndNil(TempString);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/10/2007
Modifié le ... :   /  /    
Description .. : Fonction qui renvoi un élément fils de l'élément passé en 
Suite ........ : paramètre. Si l'élément possède déjà des fils, on renvoi le 
Suite ........ : dernier fils. Si l'élément ne possède aucun fils, on en crée 
Suite ........ : un. Si l'élément ne gère pas les différents niveaux, c'est
Suite ........ : lélément lui-même qui est renvoyé.
Mots clefs ... : 
*****************************************************************}
Function LogGetChildLevel(Diag : TObject) : TObject;
begin
  result := Diag;
  if not Assigned(Diag) then exit;
  if Diag is TTreeNode then
  begin
    if (Diag as TTreeNode).HasChildren then
      result := (Diag as TTreeNode).GetLastChild
    else begin
      result := (Diag as TTreeNode).Owner.AddChild((Diag as TTreeNode), '')
    end;
  end else if Diag is Tob then
  begin
    if (Diag as Tob).Detail.Count > 0 then
      result := (Diag as Tob).Detail.Items[(Diag as Tob).Detail.Count -1]
    else begin
      result := Tob.Create('Log message', (Diag as Tob), -1);
    end;
  end;
end;

procedure IntegreRubBul(Salarie, TPE: TOB; Rubrique, Natrub: string);
var
  TRC, THB: TOB;
begin
  if Rubrique = '' then exit;
  if Natrub = 'AAA' then TRC := Paie_RechercheOptimise(tob_rem, 'PRM_RUBRIQUE', rubrique)
  else if Natrub = 'COT' then TRC := Paie_RechercheOptimise(tob_cotisations, 'PCT_RUBRIQUE', rubrique)
  else if Natrub = 'BAS' then TRC := Paie_RechercheOptimise(tob_bases, 'PCT_RUBRIQUE', rubrique);
  if TRC = nil then exit; // Integration rubrique de remunératoin inexistante
  if Salarie = nil then exit;
  if TPE.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [NatRub, Rubrique], FALSE) = nil then // $$$$
  begin
    THB := TOB.Create('HISTOBULLETIN', TPE, -1);
    RemplirHistoBulletin(THB, Salarie, TRC, nil, TPE.GetValue('PPU_DATEDEBUT'), TPE.GetValue('PPU_DATEFIN'));
  end;
end;

  {***********A.G.L.***********************************************
  Auteur  ...... : Ph
  Créé le ...... : 22/08/2003
  Modifié le ... :   /  /
  Description .. : Fonction qui calcule le montant de la saisie arret sur chaque ligne
  Suite ........ : PT84   : 26/08/2003 PH V_42 Mise en place saisie arret
  Mots clefs ... :
  *****************************************************************}

function RendMtLigneSaisieArret(VariablePaie: string; TOB_SaisieArretLocal: TOB): Double;
var
  i: Integer;
begin
  result := 0;
  for i := 0 to TOB_SaisieArretLocal.detail.count - 1 do
  begin
    if (TOB_SaisieArretLocal.detail[i].GetValue('PTR_RUBRIQUE') = VariablePaie) then
    begin
      result := result + TOB_SaisieArretLocal.detail[i].GetValue('MONTANTBUL');
      TOB_SaisieArretLocal.detail[i].PutValue('INTBULL', 'X');
    end;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Fonction qui retourne la bonne tob en fonction du type de
Suite ........ : la rurique
Mots clefs ... :
*****************************************************************}

function PaieExpression(const Expression, Vrai, Faux: variant): variant;
begin
  if Expression then result := Vrai else Result := Faux;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Fonction qui retourne la bonne tob en fonction du type de
Suite ........ : la rubrique
Mots clefs ... :
*****************************************************************}

function Paie_RechercheOptimise(const TobSrc: tob; const Field: string; const Valeur: Variant): TOB;
var
  iStart, iMax, iNumChamp, iCount, iPrevCount: integer;
begin
  result := nil;
  if (Pos('PHC_', Field) > 0) then TobSrc.Detail.Sort(Field);
  if Assigned(TobSrc) and (TobSrc.Detail.Count > 0) then
  begin
    // récupération du n° de champ
    iNumChamp := TobSrc.Detail[0].GetNumChamp(Field);
    if iNumChamp > 0 then
    begin
      iStart := 0;
      iPrevCount := -1;
      iCOunt := 0;
      iMax := TobSrc.Detail.Count - 1;
      if TobSrc.Detail[iMax].GetValeur(iNumChamp) = Valeur then result := TobSrc.Detail[iMax]
      else if TobSrc.Detail[0].GetValeur(iNumChamp) = Valeur then result := TobSrc.Detail[0]
      else while not assigned(result) do
        begin
          if (iPrevCount = iCount) then break;
          iPrevCount := iCount;
          iCount := ((iMax - iStart) shr 1 + iStart);
          if TobSrc.Detail[iCount].GetValeur(iNumChamp) = Valeur then result := TobSrc.Detail[iCount]
          else if TobSrc.Detail[iCount].GetValeur(iNumChamp) < Valeur then iStart := iCount
          else iMax := iCount;
        end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Fonction qui retourne la bonne tob en fonction du type de
Suite ........ : la rurique
Mots clefs ... :
*****************************************************************}

function Paie_RechercheRubrique(const typerubrique, rubrique: string): tob;
begin
  result := nil;
  if typeRubrique = '' then exit;

  if typerubrique = 'AAA' then result := Paie_RechercheOptimise(tob_rem, 'PRM_RUBRIQUE', rubrique)
  else if typerubrique = 'COT' then result := Paie_RechercheOptimise(tob_cotisations, 'PCT_RUBRIQUE', rubrique)
  else if typerubrique = 'BAS' then result := Paie_RechercheOptimise(tob_bases, 'PCT_RUBRIQUE', rubrique);
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. : Fonction qui retourne le préfixe en fonction type de
Suite ........ : la rurique
Mots clefs ... :
*****************************************************************}

function Paie_PrefixeRubrique(const typerubrique: string): string;
begin
  result := 'PRM';
  if (typerubrique = 'COT') or (typerubrique = 'BAS') then result := 'PCT';
end;

{***********A.G.L.***********************************************
Auteur  ...... : D.T.
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. : Fonction de chargement des tob pour les bases et les cotisations
Mots clefs ... :
*****************************************************************}

procedure ChargeCotisations;
begin
  initTOB_Cotisations();
  initTOB_Bases();
end;

// Procedure qui remplit l'historique de paies

procedure RemplirHistoBulletin(HistoBul, Salarie, Rubrique, ProfilRub: TOB; const DateDeb, DateFin: TDateTime);
var
  Prefixe, St: string;
  Nature, Rub, Etab, CodeSal: string;
begin
  if Rubrique = nil then exit;
  Prefixe := TableToPrefixe(Rubrique.NomTable);
  if Prefixe = '' then
  begin
    st := TCS(Rubrique.ChampsSup[0]).Nom;
    Prefixe := ReadTokenPipe(St, '_');
  end;

  Nature := Rubrique.GetValue(Prefixe + '_NATURERUB');
  Rub := Rubrique.GetValue(Prefixe + '_RUBRIQUE');
  Etab := Salarie.GetValeur(iPSA_ETABLISSEMENT);
  CodeSal := Salarie.GetValeur(iPSA_SALARIE);
  if (Etab = '') or (CodeSal = '') or (Nature = '') or (DateDeb = 0) or (DateFIN = 0) then
  begin
    St := 'Salarié : ' + CodeSal + ' Confection du Bulletin : Erreurs Champ non renseigné';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '') else TraceErr.Items.Add(St);
  end;
  if Assigned(HistoBul) and (iPHB_ETABLISSEMENT = 0) then MemorisePhb(HistoBul);
  with HistoBul do
  begin
    PutValeur(iPHB_ETABLISSEMENT, Salarie.GetValeur(iPSA_ETABLISSEMENT));
    PutValeur(iPHB_SALARIE, Salarie.GetValeur(iPSA_SALARIE));
    PutValeur(iPHB_DATEDEBUT, DateDeb);
    PutValeur(iPHB_DATEFIN, DateFin);
    PutValeur(iPHB_NATURERUB, Rubrique.GetValue(Prefixe + '_NATURERUB'));
    PutValeur(iPHB_RUBRIQUE, Rubrique.GetValue(Prefixe + '_RUBRIQUE'));
    PutValeur(iPHB_LIBELLE, Rubrique.GetValue(Prefixe + '_LIBELLE'));
    PutValeur(iPHB_IMPRIMABLE, Rubrique.GetValue(Prefixe + '_IMPRIMABLE'));
    if Prefixe = 'PRM' then
    begin
      //A revoir
      PutValeur(iPHB_BASEREM, 0); //Rubrique.GetValue('PRM_BASEREM')) ;
      PutValeur(iPHB_TAUXREM, 0); //Rubrique.GetValue('PRM_TAUXREM')) ;
      PutValeur(iPHB_COEFFREM, 0); //Rubrique.GetValue('PRM_COEFFREM')) ;
      PutValeur(iPHB_MTREM, 0); //Rubrique.GetValue('PRM_MONTANT')) ;
      //////
      PutValeur(iPHB_BASEREMIMPRIM, Rubrique.GetValue('PRM_BASEIMPRIMABLE'));
      PutValeur(iPHB_TAUXREMIMPRIM, Rubrique.GetValue('PRM_TAUXIMPRIMABLE'));
      PutValeur(iPHB_COEFFREMIMPRIM, Rubrique.GetValue('PRM_COEFFIMPRIM'));
      PutValeur(iPHB_ORDREETAT, Rubrique.GetValue('PRM_ORDREETAT'));
      PutValeur(iPHB_OMTSALARIAL, Rubrique.GetValue('PRM_ORDREETAT')); { PT136 }
      PutValeur(iPHB_SENSBUL, Rubrique.GetValue('PRM_SENSBUL'));
      PutValeur(iPHB_COTREGUL, '...'); // PT152
    end
    else
    begin
      PutValeur(iPHB_BASECOT, 0); //Rubrique.GetValue('PCT_BASECOTISATION')) ;
      PutValeur(iPHB_TAUXSALARIAL, 0); //Rubrique.GetValue('PCT_TAUXSAL')) ;
      PutValeur(iPHB_TAUXPATRONAL, 0); //Rubrique.GetValue('PCT_TAUXPAT')) ;
      PutValeur(iPHB_MTSALARIAL, 0);
      PutValeur(iPHB_MTPATRONAL, 0);
      PutValeur(iPHB_COTREGUL, '...');
      PutValeur(iPHB_ORGANISME, Rubrique.GetValue('PCT_ORGANISME'));
      PutValeur(iPHB_BASECOTIMPRIM, Rubrique.GetValue('PCT_BASEIMP'));
      PutValeur(iPHB_TAUXSALIMPRIM, Rubrique.GetValue('PCT_TXSALIMP'));
      PutValeur(iPHB_TAUXPATIMPRIM, Rubrique.GetValue('PCT_TXPATIMP'));
      PutValeur(iPHB_ORDREETAT, 3); { PT136 }
      PutValeur(iPHB_OMTSALARIAL, Rubrique.GetValue('PCT_ORDREETAT')); { PT136 }
      PutValeur(iPHB_SENSBUL, 'P');
    end;

    PutValeur(iPHB_TRAVAILN2, Salarie.GetValeur(iPSA_TRAVAILN2));
    PutValeur(iPHB_TRAVAILN3, Salarie.GetValeur(iPSA_TRAVAILN3));
    PutValeur(iPHB_TRAVAILN4, Salarie.GetValeur(iPSA_TRAVAILN4));
    PutValeur(iPHB_TRAVAILN1, Salarie.GetValeur(iPSA_TRAVAILN1));
    PutValeur(iPHB_CODESTAT, Salarie.GetValeur(iPSA_CODESTAT));
    PutValeur(iPHB_LIBREPCMB1, Salarie.GetValeur(iPSA_LIBREPCMB1));
    PutValeur(iPHB_LIBREPCMB2, Salarie.GetValeur(iPSA_LIBREPCMB2));
    PutValeur(iPHB_LIBREPCMB3, Salarie.GetValeur(iPSA_LIBREPCMB3));
    PutValeur(iPHB_LIBREPCMB4, Salarie.GetValeur(iPSA_LIBREPCMB4));
    //PT22 : 22/11/01 V562 PH ORACLE, Le champ PHB_ORGANISME doit être renseigné
    if GetValeur(iPHB_ORGANISME) = '' then PutValeur(iPHB_ORGANISME, '....');
    PutValeur(iPHB_CONFIDENTIEL, Salarie.GetValeur(iPSA_CONFIDENTIEL));
    PutValeur(iPHB_CONSERVATION, 'PRO'); // Pour indiquer que la rubrique vient d'un profil
  end;
  AlimCumulSalarie(HistoBul, CodeSal, Nature, Rub, Etab, DateDeb, DateFin);
end;
//PT32 : 07/02/2002 PH V571 suppression historisation salarié
{
function IncrementeSeqNoOrdrePHS(cle_sal: string): Integer;
var
   q: TQuery;
begin
    result := 1;
    q := OpenSQL('select max (PHS_ORDRE) as maxno from HISTOSALARIE ' +
                 ' where PHS_SALARIE = "' + cle_sal+'"', TRUE);

    if not q.eof then
// $IFDEF EAGLCLIENT
    if q.Fields[0].AsInteger = 0 then
// $ELSE
          if not q.Fields[0].IsNull then
//$ENDIF
             result := ((q.Fields[0]).AsInteger+1);

    Ferme(q);
end;
}
// FIN PT32
{Fonction qui charge la liste des rubriques en fonction du profil
Elle traite les cas particuliers qui consistent à exclure des rubriques de cotisation qui ont un
thème défini exemple EXO exclusion de l'URSSAF.
Le theme à exclure est renseigné au niveau du profil.
Cela provoque alors la suppression de toutes les rubriques de cotisation qui ont ce théme de la
liste des rubriques composant le bulletin de paie.
}

procedure ChargeProfil(Salarie, TPE: TOB; Profil: string);
var
  TPP, TPR, TRC, THB: TOB;
  ThemeRub, ThemeExclus: string;
  I: Integer;
  iChampNatureRub: integer;
  iChampRubrique: integer;
begin
  if Profil = '' then exit;
  ThemeRub := '';
  ThemeExclus := '';
  TPP := Paie_RechercheOptimise(TOB_ProfilPaies, 'PPI_PROFIL', Profil); // $$$$
  if TPP <> nil then
  begin
    ThemeExclus := TPP.GetValue('PPI_THEMECOT'); // theme des rubriques de cotisations à exclure
    if ThemeExclus <> '' then SuppCotExclus(ThemeExclus, TPE);
    if tpp.Detail.Count > 0 then
    begin
      iChampNatureRub := tpp.detail[0].GetNumChamp('PPM_NATURERUB');
      iChampRubrique := tpp.detail[0].GetNumChamp('PPM_RUBRIQUE');
      for I := 0 to TPP.Detail.count - 1 do
      begin
        TPR := TPP.Detail[I];
        //          SSS := TPR.GetValeur(iChampNatureRub) + TPR.GetValeur(iChampRubrique) ;
        Trc := Paie_RechercheRubrique(TPR.GetValeur(iChampNatureRub), TPR.GetValeur(iChampRubrique));
        if TPE.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [TPR.GetValeur(iChampNatureRub), TPR.GetValeur(iChampRubrique)], FALSE) = nil then // $$$$
        begin
          // on va ecrire une ligne bulletin THB comme fille de la TOB TPE qui est l'entete bulletin
          THB := TOB.Create('HISTOBULLETIN', TPE, -1);
          RemplirHistoBulletin(THB, Salarie, TRC, TPR, TPE.GetValue('PPU_DATEDEBUT'), TPE.GetValue('PPU_DATEFIN'));
        end;
      end;
    end;
  end;
end;
{ Fonction de recherche de la liste des rubriques composant chaque profil
Attention, Pour chaque profil, il faut regarder son type et s'il est identique à celui de l'établissement
il faut alors aller rechercher les profils définis par défaut au nivaeu de l'établissement
Attention, il manque le traitement de modification : @@@@@ ??????
CAD on ne reprend pas les rubriques du profil mais le bulletin tel qu'il était valorisé
lors de la paie.
Modif 4/4/01 pour reprendre la paie precedente du salarie et non pas Etablissement/Salarie
 donc modif des requetes SQL et des recherches dans les TOB
}
// PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
//        Rajout parametres TheBullCompl, TheProfilPart Pour indiquer si bull complémentaire et son profil associé

function ChargeRubriqueSalarie(Salarie: TOB; DateDeb, DateFin: TDateTime; ActionB: TActionBulletin; MtEnvers: Double; TheProfilPart: string): TOB;
var
  T_Etab, TPR, TPE, TT: TOB;
  Q: TQuery;
  TheEtab, St: string; // PT151
  Date1, Date2, ZDate: TDateTime;
  i: integer;
begin
  TOB_Salarie := Salarie; // Memorisation du salarie sur lequel on travaille
  TPE := TOB.Create('PAIEENCOURS', nil, -1);
  // DEB PT151
  if ActionB = taCreation then TheEtab := string(Salarie.GetValeur(iPSA_ETABLISSEMENT))
  else
  begin
    st := 'SELECT PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_SALARIE="' + Salarie.GetValeur(iPSA_SALARIE) +
      '" AND PPU_DATEDEBUT="' + UsDateTime(DateDeb) + '" AND PPU_DATEFIN="' + UsDateTime(DateFin) + '"';
    if BullCompl = 'X' then st := St + ' AND PPU_BULCOMPL = "X"'
    else st := St + ' AND PPU_BULCOMPL <> "X"';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then TheEtab := Q.FindField('PPU_ETABLISSEMENT').AsString
    else TheEtab := string(Salarie.GetValeur(iPSA_ETABLISSEMENT));
    FERME(Q);
  end;

  TPE.PutValue('PPU_ETABLISSEMENT', TheEtab);
  T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TheEtab);
  if T_Etab = nil then
  begin
    St := 'Salarié : ' + CodSal + ' Charge Rubrique Salarié : Etablissement non Référencé';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '') else TraceErr.Items.Add(St);
    result := nil;
    exit;
  end;
  // FIN PT151
  // PT100  : 05/04/2004 PH V_50 FQ 11231 Prise en compte du cas spécifique Alsace Lorraine
  RegimeAlsace := T_Etab.GetValue('ETB_REGIMEALSACE');
  // Alimentation des champs Entete du bulletin
  AlimChampEntet(Salarie.GetValeur(iPSA_SALARIE), TheEtab, DateDeb, DateFin, TPE); // PT151

  // Chargement des rubriques de chaque profil reférencé au niveau salarié en création
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl <> 'X' then
  begin
    if ActionB = taCreation then AligneProfil(Salarie, T_Etab, TPE, ActionB);
  end
  else
  begin
    if ActionB = taCreation then
    begin
      if TheProfilPart <> '' then ChargeProfil(Salarie, TPE, TheProfilPart)
      else AligneProfil(Salarie, T_Etab, TPE, ActionB);
    end;
  end;

  {CHargement des rubriques composant le bulletin précédant
  Lecture des paieencours @@@ si on est en mode creation uniquement , en modif on recupere le bulletin tq
  cad le bulletin sur lequel on travaille.

  Cas bulletin complémentaire, on ne recherche pas à aligner les rubriques du
  bulletin par<rapport au profil. C'est une image faite à la création du bulletin
  Si on veut modifier son contenu et le réaligner par rapport au profil, il faut
  supprimer le bulletin.
  Le bulletin complémentaire est un bulletin exceptionnel fait pour saisir des infos
  de paie apès que le salarié soit parti ou bien por faire un bulletin de type participation
  pour lequel il n'y a pas d'éléments permananents à récupérer.
  Une Rémunération suivie des cotisations CSG et CRDS.
  }
  if ActionB = taCreation then
  begin // Récupération de la dernière paie faite pour le salarié
    TOB_PaiePrecedente := TOB.Create('Le Bulletin de la Paie Précédante', nil, -1);
    // PT65    02/12/2002 PH V591 Recup des infos du bulletin précédent en excluant les bulletins complémentaires
    // PT151 Optimisation du code on ne récupère que les dates et non la totalité de la table
    st := 'SELECT PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_BULCOMPL <> "X" AND PPU_SALARIE="' + Salarie.GetValeur(iPSA_SALARIE) + '" ORDER BY PPU_DATEDEBUT,PPU_DATEFIN';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    TOB_PaiePrecedente.LoadDetailDB('PAIEENCOURS', '', '', Q, FALSE);
    ferme(Q);
}
    TOB_PaiePrecedente.LoadDetailDBFromSQL('PAIEENCOURS', st);
    Date1 := 0;
    Date2 := 0;
    for i := 0 to TOB_PaiePrecedente.Detail.Count - 1 do
      with TOB_PaiePrecedente.Detail[i] do
      begin
        ZDate := TDateTime(GetValue('PPU_DATEFIN'));
        if (Zdate < Date2) and (Date2 <> 0) then break; // cas où il y a un bulletin postérieur alors on s'arrete
        if ZDate > DateFin then break;
        Date1 := TDateTime(GetValue('PPU_DATEDEBUT'));
        Date2 := TDateTime(GetValue('PPU_DATEFIN'));
      end;

    // PT65    02/12/2002 PH V591 Free avt le NIL !!!
    TOB_PaiePrecedente.Free;
    TOB_PaiePrecedente := nil;
    if Date2 >= DateFin then // On calcule une paie qui a une paie postérieure donc on ne prend pas les lignes historiques
    begin
      Date2 := DateFin;
      PLUSMOIS(Date2, -1); // Mois -1 au cas ou mais peu probable car si on a des trous dans les paies
      Date1 := DEBUTDEMOIS(Date2); // Recupération du mois precedent
    end;
  end
  else
  begin // Les dates de la paie sont les dates passées en paramères
    Date1 := DateDeb;
    Date2 := DateFin;
  end;

  if (Date1 <> 0) and (Date2 <> 0) then
  begin // recup des lignes du dernier bulletin fait pour le salarié/etablissement
    TOB_PaiePrecedente := TOB.Create('Les lignes de la Paie Précédante', nil, -1);
    st := 'SELECT * FROM HISTOBULLETIN WHERE PHB_SALARIE="' + Salarie.GetValeur(iPSA_SALARIE) + '" AND PHB_DATEDEBUT="' + USDateTime(Date1) + '" AND PHB_DATEFIN="' +
      USDateTime(Date2) + '"';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    if not q.eof then TOB_PaiePrecedente.LoadDetailDB('HISTOBULLETIN', '', '', Q, False);
    Ferme(Q);
}
    TOB_PaiePrecedente.LoadDetailDBFROMSQL('HISTOBULLETIN', St);
    // Modified by XP : 09-04-2003
    // TPR:= TOB_PaiePrecedente.FindFirst(['PHB_SALARIE','PHB_DATEDEBUT','PHB_DATEFIN'],[Salarie.GetValeur (iPSA_SALARIE'),Date1, Date2],TRUE) ; // $$$$
    // while TPR <> NIL do
    for i := 0 to TOB_PaiePrecedente.Detail.Count - 1 do
    begin
      tpr := TOB_PaiePrecedente.Detail[i];
      if ActionB = taCreation then
      begin
        MajRubriqueBulletin(TPE, TPR, CodSal, TheEtab, DateDeb, DateFin); //PT151
        // controle si la rubrique est presénte dans le bulletin et récup élément permanent
      end
      else
      begin
        TT := TOB.create('HISTOBULLETIN', TPE, -1);
        if Assigned(TT) and (iPHB_ETABLISSEMENT = 0) then MemorisePhb(TT);
        with TT do
        begin
          Dupliquer(TPR, FALSE, TRUE, TRUE);
          PutValeur(iPHB_ETABLISSEMENT, TheEtab); // PT151
          PutValeur(iPHB_SALARIE, CodSal);
          PutValeur(iPHB_DATEDEBUT, DateDeb);
          PutValeur(iPHB_DATEFIN, DateFin);
          if GetValeur(iPHB_COTREGUL) = 'REG' then
            PutValeur(iPHB_RUBRIQUE, GetValeur(iPHB_RUBRIQUE) + '.R');
        end;
      end;
      // TPR:=TOB_PaiePrecedente.FindNext(['PHB_SALARIE','PHB_DATEDEBUT','PHB_DATEFIN'],[Salarie.GetValeur (iPSA_SALARIE),Date1, Date2],TRUE) ;
    end;

    if ActionB = taModification then
    begin // il oonvient de regarder si on a une nouvelle rubrique dans les profils qui ne soient pas dans le bulletin
      if BullCompl <> 'X' then AligneProfil(Salarie, T_Etab, TPE, ActionB)
      else
      begin
        if TheProfilPart <> '' then ChargeProfil(Salarie, TPE, TheProfilPart)
        else AligneProfil(Salarie, T_Etab, TPE, ActionB);
      end;
    end;

    FreeAndNil(TOB_PaiePrecedente);
  end;

  TPE.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
  // Tri des rubriques par nature et par code afin de pouvoir les retrouver dans un ordre logique et chrnonologique
  // appel foonction de calcul du bulletin avec la valorisation des cumuls salariés
  Result := TPE;
end;

{// PT184
 Fonction de chargement du contexte de la présence
 Charge Toutes les TOB concernées en memoire
 Nécéssaire au calcul des variables de présence
}
procedure ChargePresence(FiltreSalarie : String = ''; DateDebut : TDateTime = 0; DateFin : TDateTime = 0; DateFinEstFinDeMois : Boolean = False);
begin
    if (ObjCalcuVarPre<>nil) then
       FreeAndNil(ObjCalcuVarPre);
    ObjCalcuVarPre := TCalcuVarPre.Create(FiltreSalarie, DateDebut, DateFin, DateFinEstFinDeMois);
end;

{ Fonction de chargement du contexte de la paie
Charge Toutes les TOB en memoire
}
procedure ChargeLesTOBPaie;
var
  I: Integer;
  Them, St: string;
  T, T1: TOB;
//  Q: TQuery;
//  SystemTime0, SystemTime1: TSystemTime;

  // added by XP le 14-03-2003
//  TobTmp: tob; // Tob liste des rubriques composant les profils
  NumError: Integer;
begin
  NumError := 1;
  try
//  GetLocalTime(SystemTime0);

    // PROCESS-SERVEUR  {$IFNDEF EAGLSERVER} {$ENDIF}
//PT177
    if (ObjTM3VM<>nil) then
       FreeAndNil (ObjTM3VM);
//FIN PT177
    ObjTM3VM := TM3VM.Create; // creation objet VM Script
    // Gestion des tables dynamiques
    if (ObjTableDyn<>nil) then
       FreeAndNil(ObjTableDyn);
    ObjTableDyn := TTablesDynamiques.Create; // PT175
{$IFNDEF EAGLSERVER}
    if VH_PAIE.PGAnalytique = TRUE then InitMove(20, ' ') else InitMove(18, ' ');
{$ENDIF}
    initTOB_ProfilPaies();
{$IFNDEF EAGLSERVER}
    NumError := 2;
    MoveCur(FALSE);
    NumError := 3;
    MoveCur(FALSE);
{$ENDIF}
    initTOB_Rem();
    NumError := 4;
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    { TOB_Rem contient la totalité des Remunerations
    Il convient d'eclater ces rubriques en fonction de leur theme et ainsi
    de remplir les onglets de la saisie de bulletin dans le cas d'une saisie
    avec tous les onglets}
    for I := 0 to TOB_Rem.Detail.count - 1 do
    begin
      T := TOB_REM.Detail[I];
      Them := T.GetValue('PRM_THEMEREM');
      if Them = 'SAL' then T1 := TOB.create('RemSal', TOB_RemSal, -1)
      else if Them = 'ABS' then T1 := TOB.create('RemAbs', TOB_RemAbs, -1)
      else if Them = 'ABT' then T1 := TOB.create('RemAbt', TOB_RemAbt, -1)
      else if Them = 'AVT' then T1 := TOB.create('RemAvt', TOB_RemAvt, -1)
      else if Them = 'RNI' then T1 := TOB.create('RemNonImp', TOB_RemNonImp, -1)
      else if Them = 'INI' then T1 := TOB.create('RemPrimes', TOB_RemPrimes, -1)
      else if Them = 'RSS' then T1 := TOB.create('RemRet', TOB_RemRet, -1)
      else if Them = 'COM' then T1 := TOB.create('RemCplt', TOB_RemCplt, -1)
      else if Them = 'HEU' then T1 := TOB.create('RemHeures', TOB_RemHeures, -1);
      if (Them <> '') and (T1 <> nil) then T1.Dupliquer(TOB_Rem.Detail[i], TRUE, TRUE);
    end;

    NumError := 5;
    // il convient aussi de différencier les rubriques de Base des rubriques de Cotisations
    ChargeCotisations;
    NumError := 6;
    initTOB_Execpt(); // Chargement de la tob des execptions
    Align_Execpt('', ''); // Récuperation des personnalisations
    initTOB_Variables();
{$IFNDEF EAGLSERVER}
    NumError := 7;
    MoveCur(FALSE);
{$ENDIF}
    initTOB_EltNatDOS();
    initTOB_EltNatSTD();
    initTOB_EltNiveauRequis(); //PT172
    initTOB_EltNatCEG();
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}

    TOB_JourTrav.LoadDetailDB('JOURSTRAV', '', '', nil, FALSE, False);
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    // Les cumuls et la gestion associée = CUMULRUBRIQUE
    initTOB_Cumuls();
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    initTOB_MinimumConv();
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}

    // PT72  : 05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer
    initTOB_MinConvPaie();
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
    NumError := 8;
    NumError := 9;
{$ENDIF}

    initTOB_EtabCompl();
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    Charg_TauxAt(); // PT137
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}

    st := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "RR%" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    TOB_VenRem.LoadDetailDB('VENTIL', '', '', Q, FALSE, False);
    Ferme(Q);
}
    TOB_VenRem.LoadDetailDBFROMSQL('VENTIL', st);
{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    st := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "RC%" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    TOB_VenCot.LoadDetailDB('VENTIL', '', '', Q, FALSE, False);
    Ferme(Q);
}
    TOB_VenCot.LoadDetailDBFromSQL('VENTIL',st);

{$IFNDEF EAGLSERVER}
    MoveCur(FALSE);
{$ENDIF}
    if VH_PAIE.PGAnalytique = TRUE then
    begin
      initTOB_VentiRemPaie();
      initTOB_VentiCotPaie();
{$IFNDEF EAGLSERVER}
      MoveCur(FALSE);
{$ENDIF}
    end;
{$IFNDEF EAGLSERVER}
    FiniMove;
{$ENDIF}

    NumError := 10;

  except
    on E: Exception do
    begin
{$IFNDEF EAGLSERVER}
      ShowMessage('Exception N° ' + IntToStr(NumError));
{$ENDIF}
    end;
  end;
//  GetLocalTime(SystemTime1);
//  st := 'Debut : '+TimeToStr(SystemTimeToDateTime(SystemTime0))+ ' Fin : '+ TimeToStr(SystemTimeToDateTime(SystemTime1));

end;

{ Fonction de chargement de l'exercice de la paie en fonction des Date de Début
et de FIN
}

procedure ChargeLesExercPaie(const DateDeb, DateFin: TDateTime);
var
//  Q: TQuery;
  St: string;
begin
  if TOB_ExerSocial <> nil then TOB_ExerSocial.Free;
  TOB_ExerSocial := nil;
  TOB_ExerSocial := TOB.Create('Les exerices du dossier', nil, -1);
  st := 'SELECT * FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + UsDateTime(DateDeb) + '" AND PEX_DATEFIN>="' + UsDateTime(DateFin) + '"';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_ExerSocial.LoadDetailDB('EXERSOCIAL', '', '', Q, False);
  Ferme(Q);
}
 TOB_ExerSocial.LoadDetailDBFROMSQL('EXERSOCIAL',st);
end;


procedure InitLesTOBPaie;
begin
//PT177
   FreeandNil (TOB_Paie);
   FreeandNil (TOB_RemSal);
   FreeandNil (TOB_RemHeures);
   FreeandNil (TOB_RemPrimes);
   FreeandNil (TOB_RemAbs);
   FreeandNil (TOB_RemCplt);
   FreeandNil (TOB_RemAvt);
   FreeandNil (TOB_RemAbt);
   FreeandNil (TOB_RemRet);
   FreeandNil (TOB_RemNonImp);
   FreeandNil (TOB_JourTrav);
   FreeandNil (TOB_TauxAT);
   FreeandNil (TOB_VenRem);
   FreeandNil (TOB_VenCot);
//FIN PT177

  TOB_Paie := TOB.Create('Les TOB de la paie', nil, -1);
  //TOB_ProfilRubs:=TOB.Create('Les Rubriques des Profils',Nil,-1) ;
  // Creation des TOB Remunerations Idem Onglets saisie de bulletin
  //TOB_Rem:=TOB.Create('Les Remunerations',Nil,-1) ;
  TOB_RemSal := TOB.Create('Les Salaires', nil, -1);
  TOB_RemHeures := TOB.Create('Les Heures', nil, -1);
  TOB_RemPrimes := TOB.Create('Les Primes', nil, -1);
  TOB_RemAbs := TOB.Create('Les Absences', nil, -1);
  TOB_RemCplt := TOB.Create('Les Complements', nil, -1);
  TOB_RemAvt := TOB.Create('Les Avantages', nil, -1);
  TOB_RemAbt := TOB.Create('Les Abattements', nil, -1);
  TOB_RemRet := TOB.Create('Les Retenues', nil, -1);
  TOB_RemNonImp := TOB.Create('Les Non Imposables', nil, -1);
  TOB_JourTrav := TOB.Create('Le Calendrier', nil, -1);
  TOB_TauxAT := TOB.Create('Les Taux AT', nil, -1);
  TOB_VenRem := TOB.Create('Les VentilAna des Remunerations', nil, -1);
  TOB_VenCot := TOB.Create('Les VentilAna des Cotisations', nil, -1);
end;
{ Fonction de desallocation de la TOB des Exercices du dossier de paie}

procedure VideLaTobExer;
begin
  if TOB_ExerSocial <> nil then TOB_ExerSocial.Free;
  TOB_ExerSocial := nil;
end;

procedure VideLaPresence;
begin
  // PT184 Destruction de l'objet de gestion des variables de présence
  if Assigned(ObjCalcuVarPre) then
    FreeAndNil(ObjCalcuVarPre);
end;

{ Fonction de desallocation de toutes les tob du contexte de la Paie
}
procedure VideLesTOBPaie(AvecInit: Boolean);
begin
  if TOB_Paie <> nil then TOB_Paie.Free;
  //if TOB_ProfilRubs  <> NIL then TOB_ProfilRubs.Free;
  if TOB_RemSal <> nil then TOB_RemSal.Free;
  if TOB_RemHeures <> nil then TOB_RemHeures.Free;
  if TOB_RemPrimes <> nil then TOB_RemPrimes.Free;
  if TOB_RemAbs <> nil then TOB_RemAbs.Free;
  if TOB_RemCplt <> nil then TOB_RemCplt.Free;
  if TOB_RemAvt <> nil then TOB_RemAvt.Free;
  if TOB_RemAbt <> nil then TOB_RemAbt.Free;
  if TOB_RemRet <> nil then TOB_RemRet.Free;
  if TOB_RemNonImp <> nil then TOB_RemNonImp.Free;
{$IFDEF aucasou}
  if TOB_Rem <> nil then TOB_Rem.Free;
  if TOB_ProfilPaies <> nil then TOB_ProfilPaies.Free;
  if TOB_Cotisations <> nil then TOB_Cotisations.Free;
  if TOB_Bases <> nil then TOB_Bases.Free;
  if TOB_Variables <> nil then TOB_Variables.Free;
  if TOB_CumulRubrique <> nil then TOB_CumulRubrique.Free;
  if TOB_Cumuls <> nil then TOB_Cumuls.Free;
  if TOB_Etablissement <> nil then TOB_Etablissement.Free;
  if TOB_EltNationauxSTD <> nil then TOB_EltNationauxSTD.Free;
  if TOB_EltNationauxCEG <> nil then TOB_EltNationauxCEG.Free;
  if TOB_EltNationauxDOS <> nil then TOB_EltNationauxDOS.Free;
  if TOB_EltNiveauRequis <> nil then TOB_EltNiveauRequis.Free; //PT172
  if TOB_Minimum <> nil then TOB_Minimum.Free;
  if TOB_DetailMin <> nil then TOB_DetailMin.Free;
  if TOB_VentilRem <> nil then TOB_VentilRem.Free;
  if TOB_VentilCot <> nil then TOB_VentilCot.Free;
{$ENDIF}
  if TOB_JourTrav <> nil then TOB_JourTrav.Free;
  if TOB_TauxAT <> nil then TOB_TauxAT.Free;
  //TOB_Salarie.Free;
  if TOB_CumulSal <> nil then TOB_CumulSal.Free;
  if TOB_VenRem <> nil then TOB_VenRem.Free;
  if TOB_VenCot <> nil then TOB_VenCot.Free;

  if TOBAna <> nil then TOBAna.Free;

  TOB_Paie := nil;
  TOB_RemSal := nil;
  TOB_RemHeures := nil;
  TOB_RemPrimes := nil;
  TOB_RemAbs := nil;
  TOB_RemCplt := nil;
  TOB_RemAvt := nil;
  TOB_RemAbt := nil;
  TOB_RemRet := nil;
  TOB_RemNonImp := nil;
{$IFDEF aucasou}
  TOB_ProfilPaies := nil;
  TOB_Cotisations := nil;
  TOB_Bases := nil;
  TOB_ProfilRubs := nil;
  TOB_Rem := nil;
  TOB_Variables := nil;
  TOB_CumulRubrique := nil;
  TOB_Cumuls := nil;
  TOB_Etablissement := nil;
  TOB_EltNationauxCEG := nil;
  TOB_EltNationauxSTD := nil;
  TOB_EltNationauxDOS := nil;
  TOB_EltNiveauRequis := nil; //PT172
  TOB_Minimum := nil;
  TOB_DetailMin := nil;
  TOB_VentilRem := nil;
  TOB_VentilCot := nil;
{$ENDIF}

  TOB_JourTrav := nil;
  TOB_TauxAT := nil;
  TOBAna := nil;
  // TOB_Salarie:=NIL;
  TOB_CumulSal := nil;
  TOB_VenRem := nil;
  TOB_VenCot := nil;
  if AvecInit = TRUE then InitLesTOBPaie;

  if Assigned(ObjTM3VM) then
  begin // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
    ObjTM3VM.Free; // destruction objet VM script
    ObjTM3VM := nil;
  end;
  // PT175
  if Assigned(ObjTableDyn) then
  begin
    ObjTableDyn.Free; // destruction objet gestion des table dynamiques
    ObjTableDyn := nil;
  end;
  //PT61 Vidage de la tob calendrier
  if Tob_CalendrierSalarie <> nil then
  begin
    Tob_CalendrierSalarie.free;
    Tob_CalendrierSalarie := nil;
  end;
  // PT84   : 26/08/2003 PH V_42 Mise en place saisie arret
  if Assigned(TOB_SaisieArret) then
    FreeAndNil(TOB_SaisieArret);

  if Assigned(GblTob_VarCp) then { PT164 }
    FreeAndNil(GblTob_VarCp);

end;

procedure RAZTOBPaie();
begin
  if Assigned(TOB_Paie) then FreeAndNil(TOB_Paie);
  if Assigned(TOB_RemSal) then FreeAndNil(TOB_RemSal);
  if Assigned(TOB_RemHeures) then FreeAndNil(TOB_RemHeures);
  if Assigned(TOB_RemPrimes) then FreeAndNil(TOB_RemPrimes);
  if Assigned(TOB_RemAbs) then FreeAndNil(TOB_RemAbs);
  if Assigned(TOB_RemCplt) then FreeAndNil(TOB_RemCplt);
  if Assigned(TOB_RemAvt) then FreeAndNil(TOB_RemAvt);
  if Assigned(TOB_RemAbt) then FreeAndNil(TOB_RemAbt);
  if Assigned(TOB_RemRet) then FreeAndNil(TOB_RemRet);
  if Assigned(TOB_RemNonImp) then FreeAndNil(TOB_RemNonImp);
  Nettoyage_Rem();
  Nettoyage_ProfilPaies();
  Nettoyage_Cotisations();
  Nettoyage_Bases();
  Nettoyage_Variables();
  Nettoyage_Cumuls();
  Nettoyage_Etablissement();
  Nettoyage_EltNationauxSTD();
  Nettoyage_EltNationauxCEG();
  Nettoyage_EltNationauxDOS();
  //DEB PT172
  Nettoyage_EltNiveauRequis();
  if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
  begin
    Nettoyage_EltDynSal();
    Nettoyage_EltDynPop();
    Nettoyage_EltDynEtab();
  end;
  //FIN PT172
  Nettoyage_Minimum();
  Nettoyage_DetailMin();
  Nettoyage_VentilRem();
  Nettoyage_VentilCot();
  if Assigned(TOB_JourTrav) then FreeAndNil(TOB_JourTrav);
  if Assigned(TOB_TauxAT) then FreeAndNil(TOB_TauxAT);
  if Assigned(TOB_CumulSal) then FreeAndNil(TOB_CumulSal);
  if Assigned(TOB_VenRem) then FreeAndNil(TOB_VenRem);
  if Assigned(TOB_VenCot) then FreeAndNil(TOB_VenCot);
  if Assigned(TOBAna) then FreeAndNil(TOBAna);
  if Assigned(Tob_CalendrierSalarie) then FreeAndNil(Tob_CalendrierSalarie);
end;


{Rend la valeur d'un element national à une date donnée
Si pas de date Alors rend le dernier élément/date trouvé
Si élément non renseigné alors Message Erreur ==> Problème de paramètrage du
plan de paie
Attention, il manque la gestion de l'EURO pour retourner la valeur dans la bonne monnaie
}

function ValEltNat(const Elt: string; const DatVal: TDateTime; Diag: TObject = nil): double;
var
  ret: double;
  St: string;
  NiveauRequis:String;
  sNiveau:String;
begin
  ret := -123456; // initialisation dans le cas peu probable où on ne trouve pas
  // Il convient de differencier le petit du grand decalage de paie
  // Le petit est basé sur un exercice social correspond à l'exercice civil
  // alors que le grand est décalé - 1 mois
  // PT62    07/11/2002 PH V591 Prise en compte du decalage de paie donc le traitement depend de l'élément
  //         national et du decalage, il faut en tenir compte dans la focntion d'évaluation
  //if (VH_Paie.PGDecalage=True) then DatTrait := PLUSMOIS(DatTrait, 1);
  //if (VH_Paie.PGDecalagePetit=True) then DatTrait := PLUSMOIS(DatTrait, 1);

  if Elt = '' then
  begin
    St := 'Salarié : ' + CodSal + ' Elément National non renseigné';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
  end
  else
  begin
    //DEB PT172
    // Récupérer le niveau préconisé et le niveau maxi de personnalisation pour l'élément national en cours
    RecupNiveauRequis(Elt);
    NiveauRequis := stNiveauRequis;
    if NiveauRequis = '' then NiveauRequis := 'TOU'; // Cas où aucun paramétrage, on regarde tout
    if stNiveauMaxi = '' then stNiveauMaxi := 'TOU'; // Cas où aucun paramétrage, on regarde tout
    while NiveauRequis <> '' do
    begin
      // Si le niveau requis est renseigné, regarder en priorité jusqu'à ce niveau
      // Si la valeur n'est pas trouvé alors proposer la saisie et mettre à jour la TOB correspondante si besoin
      if GetParamSocSecur('SO_PGGESTELTDYNDOS',False) then
      begin
        if (stNiveauMaxi = 'SAL') or (stNiveauMaxi = 'TOU') then
        begin
          sNiveau := 'Salarié';
          ret := ValEltNatDynDos(Elt, 'SAL', DatVal);
          if (ret = -123456) and (NiveauRequis = 'SAL') then
            ret := DemandeSaisie('SAL',Elt,NiveauRequis,DatVal);
        end;
        if (stNiveauMaxi = 'POP') or (stNiveauMaxi = 'SAL') or (stNiveauMaxi = 'TOU') then
        begin
          if (ret = -123456) then
          begin
            sNiveau := 'Population';
            ret := ValEltNatDynDos(Elt, 'POP', DatVal);
          end;
        end;
        if (ret = -123456) and (NiveauRequis = 'POP') then
          ret := DemandeSaisie('POP',Elt,NiveauRequis,DatVal);
        if (stNiveauMaxi = 'ETB') or (stNiveauMaxi = 'POP') or (stNiveauMaxi = 'SAL') or (stNiveauMaxi = 'TOU') then
        begin
          if ret = -123456 then
          begin
            sNiveau := 'Etablissement';
            ret := ValEltNatDynDos(Elt, 'ETB', DatVal);
          end;
        end;
        if (ret = -123456) and (NiveauRequis = 'ETB') then
          ret := DemandeSaisie('ETB',Elt,NiveauRequis,DatVal);
      end;
      if (stNiveauMaxi = 'DOS') or (stNiveauMaxi = 'ETB') or (stNiveauMaxi = 'POP') or (stNiveauMaxi = 'SAL') or (stNiveauMaxi = 'TOU') then
      begin
        if (ret = -123456) then
        begin
          sNiveau := 'Dossier';
          ret := ValEltNatPref(Elt, 'DOS', DatVal);
        end;
      end;
      if (ret = -123456) and (NiveauRequis = 'DOS') then
        ret := DemandeSaisie('DOS',Elt,NiveauRequis,DatVal);
      if (stNiveauMaxi = 'STD') or (stNiveauMaxi = 'DOS') or (stNiveauMaxi = 'ETB') or (stNiveauMaxi = 'POP') or (stNiveauMaxi = 'SAL') or (stNiveauMaxi = 'TOU') then
      begin
        if (ret = -123456) then
        begin
          sNiveau := 'Standard';
          ret := ValEltNatPref(Elt, 'STD', DatVal);
        end;
      end;
      if (ret = -123456) and (NiveauRequis = 'STD') then
        ret := DemandeSaisie('STD',Elt,NiveauRequis,DatVal);
      if ret = -123456 then
      begin
        sNiveau := 'CEGID';
        ret := ValEltNatPref(Elt, 'CEG', DatVal);
      end;
      if (ret = -123456) and (NiveauRequis = 'CEG') then
        ret := DemandeSaisie('CEG',Elt,NiveauRequis,DatVal);

      if (ret <> -123456) then
//Debut PT194        if (Diag <> nil) then
//          Diag.Items.add('La valeur de l''élément national ' + Elt + ' ' + RechDom('PGELEMENTNAT',Elt,FALSE) + ' (' + FloatToStr(ret) + ') a été récupérée au niveau ' + sNiveau);
        LogMessage(Diag, 'La valeur de l''élément national ' + Elt + ' ' + RechDom('PGELEMENTNAT',Elt,FALSE) + ' (' + FloatToStr(ret) + ') a été récupérée au niveau ' + sNiveau, 'ELT', Elt, 'Recherche de la valeur de l''élément national '+Elt+' '+RechDom('PGELEMENTNAT',Elt,FALSE), (CodSal = ''))
      else
        LogMessage(Diag, 'La valeur de l''élément national ' + Elt + ' n''est pas renseignée.', 'ELT', Elt, 'Recherche de la valeur de l''élément national '+Elt+' '+RechDom('PGELEMENTNAT',Elt,FALSE), (CodSal = ''));

//Fin PT194
      NiveauRequis := '';
    end;

{    // PT60 07/10/2002 PH V585 Recherche d'un element national par préférence DOS,STD,CEG
    ret := ValEltNatPref(Elt, 'DOS', DatVal);
    if ret = -123456 then ret := ValEltNatPref(Elt, 'STD', DatVal);
    if ret = -123456 then ret := ValEltNatPref(Elt, 'CEG', DatVal); }

    //FIN PT172
  end;
  if ret = -123456 then ret := 0; // PT120
  result := ret;
end;


{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 03/10/2007
Modifié le ... :   /  /
Description .. : Evalue l'élément dynamique d'un salarié à une date donnée.
Suite ........ : Renvoie la valeur et la date de validité de l'élément trouvé
Mots clefs ... :
*****************************************************************}
function ValEltDyna(TOBSalarie : Tob; Date : TDateTime; Elt : String; var DateValidite : TDateTime) : Double;
var
  Q : TQuery;
  stQ, Salarie, Population, Etablissement : String;
  stWhereNiveau : String;
  TypeNiveau : String;
  //DEB PT204
  TobCodeLibre,TC:Tob;
  sDate : String;
  St,Libelle,Predefini:String;
  //FIN PT204
begin
//PT204  result := 0;
  result := -123456; //PT204
  Salarie := TOBSalarie.GetValeur(iPSA_SALARIE);
  { Recherche du TypeNiveau }
  stQ := 'SELECT PPP_TYPENIVEAU FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## AND PPP_PGTYPEINFOLS="ZLS" AND PPP_PGINFOSMODIF="'+Elt+'"';
  Q := OpenSQL(stQ, True, 1);
  if not Q.Eof then
  begin
    TypeNiveau := Q.Fields[0].AsString;
  end;
  Ferme(Q);
  DateValidite := Date;
  if TypeNiveau = 'SAL' then stWhereNiveau := ' AND PHD_SALARIE = "'+Salarie+'" ';
  if TypeNiveau = 'POP' then
  begin
    { Recherche de la population }
    if population = '' then
    begin
      stQ := 'select PNA_POPULATION from salariepopul where pna_salarie = "'+Salarie+'" and pna_typepop = "PAI"';
      Q := OpenSQL(stQ, True, 1);
      if not Q.Eof then
      begin
        population := Q.Fields[0].AsString;
      end;
      Ferme(Q);
    end;
    stWhereNiveau := ' AND PHD_CODEPOP = "'+population+'" ';
  end;
  if TypeNiveau = 'DOS' then stWhereNiveau := '';
  if TypeNiveau = 'ETB' then
  begin
    Etablissement := TOBSalarie.GetValeur(iPSA_ETABLISSEMENT);
    stWhereNiveau := ' AND PHD_ETABLISSEMENT = "'+Etablissement+'" ';
  end;

  stQ := ' SELECT PHD_NEWVALEUR, PHD_DATEAPPLIC FROM PGHISTODETAIL '
       + ' WHERE PHD_PGTYPEINFOLS = "ZLS" AND PHD_PGINFOSMODIF = "'+Elt+'" '
       + stWhereNiveau
       + ' AND PHD_DATEAPPLIC = (SELECT MAX(PHD_DATEAPPLIC) '
       +                       ' FROM PGHISTODETAIL '
       +                       ' WHERE PHD_PGTYPEINFOLS = "ZLS" '
       +                       ' AND PHD_PGINFOSMODIF = "'+Elt+'" '
       + stWhereNiveau
       +                       ' AND PHD_DATEAPPLIC <= "'+ USDATETIME(Date) +'") '
       + ' ORDER BY PHD_DATECREATION DESC ';
  Q := OpenSQL(stQ, True, 1);
  if not Q.Eof then
  begin
    if (Q.Fields[0].AsString = 'X') then Result := 1 else result := 0;
    DateValidite := Q.Fields[1].AsDateTime;
  end;
  Ferme(Q);

  //DEB PT204
  if result = -123456 then
  begin
    ValidationOK := False;
    if TypeTraitement = 'PREPA' then
    begin
      { Si on est en préparation automatique, ajouter un message d'erreur dans la liste }
      St := 'L''élément dynamique ' + Elt + ' n''a pas de valeur au niveau salarié ' + Salarie + '. Le bulletin n''a pas pu être calculé juste pour le salarié ' + Salarie;
      TraceErr.Items.Add(St);
    end;

    { Récupérer les valeurs par défaut dans le cas où un niveau plus générique existe (DOS ou STD) }
    TobCodeLibre := TOB.Create('Lescodeslibres', Nil, -1);
    TC := TOB.Create('PGSYNELTNAT', TobCodeLibre, -1);

    St := 'SELECT PPP_LIBELLE,PPP_PREDEFINI FROM PARAMSALARIE WHERE ##PPP_PREDEFINI## PPP_PGINFOSMODIF="' + Elt + '"';
    Q := OpenSQL(St, True, 1);
    if not Q.Eof then
    begin
      Libelle := Q.Fields[0].AsString;
      Predefini := Q.Fields[1].AsString;
    end;
    Ferme(Q);
    TC.PutValue('PEY_LIBELLE', Libelle);
    TC.PutValue('PEY_THEMEELT', ' ');
    TC.PutValue('PEY_MONETAIRE', 'X');
    TC.PutValue('PEY_ABREGE', Elt);
    TC.PutValue('PEY_BLOCNOTE', ' ');
    TC.PutValue('PEY_DECALMOIS', 'X');
    TC.PutValue('PEY_REGIMEALSACE', '-');
    TC.PutValue('PEY_CONVENTION', ' ');
    TC.PutValue('PEY_TYPUTI','3');
    TC.PutValue('PEY_TYPNIV', 'SAL');
    TC.PutValue('PEY_VALNIV', Salarie);
    TC.PutValue('PEY_LIBVALNIV', Salarie);
    TC.PutValue('PEY_ETABLISSEMENT', ' ');
    TC.PutValue('PEY_CODEPOP', ' ');
    TC.PutValue('PEY_SALARIE', Salarie);
    TC.PutValue('PEY_SALARIE', Salarie);
    TC.PutValue('PEY_CODEELT', '');

    sDate := DateTimeToStr(Date);
    if MidStr(sDate,1,2) <> '01' then
      sDate := '01' + MidStr(sDate,3,8);
    TC.PutValue('PEY_DATVAL', StrToDateTime(sDate));

    TC.PutValue('PEY_MONTANT', 0);
    TC.PutValue('PEY_MONTANTEURO', 0);
    TC.PutValue('PEY_PREDEFINI', Predefini);

{$IFNDEF EABSENCES}
    TC.PutValue('PEY_NODOSSIER', ClePGSynEltNAt); // Sert d'identifiant, contient un n° de salarié :
                                                  // - soit celui en cours pour le traitement individuel
                                                  // - soit le 1er de la liste si préparation automatique
{$ENDIF EABSENCES}

    TC.InsertOrUpdateDB;
    TobCodeLibre.Free;

    result := 0;
  end;
  //FIN PT204
end;


//DEB PT172
procedure RecupNiveauRequis(const Elt : String);
var
  EltNiveau: TOB;
  T: TOB;
  FindNum: integer;
  i: integer;
begin
  stNiveauRequis := '';
  stNiveauMaxi := '';
  EltNiveau := TOB_EltNiveauRequis;

  if assigned(EltNiveau) and (EltNiveau.detail.Count > 0) then
  begin
    if (iCHampPNRCodeElt = 0) then MemorisePnr(EltNiveau.Detail[0]);
    i := 0;
    FindNum := -1;
    while i < EltNiveau.Detail.Count do
    begin
      with EltNiveau.Detail[i] do
      begin
        if (GetValeur(iCHampPNRCodeElt) = Elt) then
        begin
          FindNum := i;
          break;
        end;
      end;
      Inc(i);
    end;

    if FindNum > -1 then
    begin
      t := EltNiveau.Detail[FindNum];
      stNiveauRequis := T.GetValeur(iCHampPNRNiveauRequis);
      stNiveauMaxi := T.GetValeur(iCHampPNRNiveauMaxi);
    end;
  end;
end;

//Fonction spécifique pour les éléments dynamiques dossier
function DemandeSaisie(TypeSaisie:String;Elt:STring;NiveauRequis:String;DatVal:TDateTime) : double;
var
  ret : double;
  St : String; // RetourSaisie,
  sNiveau:String;
begin
  ret := 0;
  if (NiveauRequis = 'CEG') then sNiveau := 'CEGID';
  if (NiveauRequis = 'STD') then sNiveau := 'Standard';
  if (NiveauRequis = 'DOS') then sNiveau := 'Dossier';
  if (NiveauRequis = 'ETB') then sNiveau := 'Etablissement';
  if (NiveauRequis = 'POP') then sNiveau := 'Population';
  if (NiveauRequis = 'SAL') then sNiveau := 'Salarié ' + TOB_Salarie.GetValeur(iPSA_SALARIE) + ' ' + TOB_Salarie.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Salarie.GetValeur(iPSA_PRENOM);
  if TypeTraitement = 'PREPA' then
  begin
    // Si on est en préparation automatique, ajouter un message d'erreur dans la liste
    St := 'L''élément national ' + Elt + ' n''a pas de valeur pour le niveau requis ' + sNiveau + '. Le bulletin n''a pas pu être calculé juste pour le salarié ' + TOB_Salarie.GetValeur(iPSA_SALARIE);
    TraceErr.Items.Add(St);
  end;
  // Stocker l'élément à saisir dans la table temporaire PGSYNELTNAT
  MemorisationElt(Elt,NiveauRequis,DatVal);
  result := ret;
end;

procedure MemorisationElt(Elt:STring;NiveauRequis:String;DatVal:TDateTime);
var
  Req:String;
  Q,Qt,Qp:TQuery;
  TobElt,TE:Tob;
  Population:String;
  sDate:String;
begin
  ValidationOK := False;
  // Récupérer les valeurs par défaut dans le cas où un niveau plus générique existe (DOS ou STD)
  TobElt := TOB.Create('Leselementsnationaux', Nil, -1);
  TE := TOB.Create('PGSYNELTNAT', TobElt, -1);

  Req := 'SELECT PEL_LIBELLE,PEL_THEMEELT,PEL_MONETAIRE,PEL_ABREGE,PEL_BLOCNOTE,PEL_DECALMOIS,'
    + ' PEL_REGIMEALSACE,PEL_PREDEFINI,PEL_NODOSSIER,PEL_CONVENTION,PEL_MONTANT,PEL_MONTANTEURO FROM ELTNATIONAUX'
    + ' WHERE PEL_DATEVALIDITE <= "' + UsDateTime(DatVal) + '"'
    + ' AND PEL_CODEELT = "' + Elt + '"'
    + ' ORDER BY PEL_DATEVALIDITE DESC';
  Q := OpenSQL(Req,True);
  if not Q.eof then
  begin
    TE.PutValue('PEY_LIBELLE', Q.FindField('PEL_LIBELLE').AsString);
    TE.PutValue('PEY_THEMEELT', Q.FindField('PEL_THEMEELT').AsString);
    TE.PutValue('PEY_MONETAIRE', Q.FindField('PEL_MONETAIRE').AsString);
    TE.PutValue('PEY_ABREGE', Q.FindField('PEL_ABREGE').AsString);
    TE.PutValue('PEY_BLOCNOTE', Q.FindField('PEL_BLOCNOTE').AsString);
    TE.PutValue('PEY_DECALMOIS', Q.FindField('PEL_DECALMOIS').AsString);
    TE.PutValue('PEY_REGIMEALSACE', Q.FindField('PEL_REGIMEALSACE').AsString);
    if (NiveauRequis = 'CEG') or (NiveauRequis = 'STD') or (NiveauRequis = 'DOS') then
      TE.PutValue('PEY_CONVENTION', Q.FindField('PEL_CONVENTION').AsString)
    else
      TE.PutValue('PEY_CONVENTION', ' ');
  end
  else
  begin
    //DEB PT182
    Ferme(Q);
    Req := 'SELECT PEL_LIBELLE,PEL_THEMEELT,PEL_MONETAIRE,PEL_ABREGE,PEL_BLOCNOTE,PEL_DECALMOIS,'
      + ' PEL_REGIMEALSACE,PEL_PREDEFINI,PEL_NODOSSIER,PEL_CONVENTION,PEL_MONTANT,PEL_MONTANTEURO FROM ELTNATIONAUX'
      + ' WHERE PEL_CODEELT = "' + Elt + '"'
      + ' ORDER BY PEL_DATEVALIDITE DESC';
    Q := OpenSQL(Req,True);
    if not Q.eof then
    begin
      TE.PutValue('PEY_LIBELLE', Q.FindField('PEL_LIBELLE').AsString);
      TE.PutValue('PEY_THEMEELT', Q.FindField('PEL_THEMEELT').AsString);
      TE.PutValue('PEY_MONETAIRE', Q.FindField('PEL_MONETAIRE').AsString);
      TE.PutValue('PEY_ABREGE', Q.FindField('PEL_ABREGE').AsString);
      TE.PutValue('PEY_BLOCNOTE', Q.FindField('PEL_BLOCNOTE').AsString);
      TE.PutValue('PEY_DECALMOIS', Q.FindField('PEL_DECALMOIS').AsString);
      TE.PutValue('PEY_REGIMEALSACE', Q.FindField('PEL_REGIMEALSACE').AsString);
      if (NiveauRequis = 'CEG') or (NiveauRequis = 'STD') or (NiveauRequis = 'DOS') then
        TE.PutValue('PEY_CONVENTION', Q.FindField('PEL_CONVENTION').AsString)
      else
        TE.PutValue('PEY_CONVENTION', ' ');
    end
    //FIN PT182
    else
    begin
      TE.PutValue('PEY_LIBELLE', ' ');
      TE.PutValue('PEY_THEMEELT', ' ');
      TE.PutValue('PEY_MONETAIRE', 'X');
      TE.PutValue('PEY_ABREGE', ' ');
      TE.PutValue('PEY_BLOCNOTE', ' ');
      TE.PutValue('PEY_DECALMOIS', 'X');
      TE.PutValue('PEY_REGIMEALSACE', '-');
      if (NiveauRequis = 'CEG') or (NiveauRequis = 'STD') or (NiveauRequis = 'DOS') then
        TE.PutValue('PEY_CONVENTION', '000')
      else
        TE.PutValue('PEY_CONVENTION', ' ');
    end;
  end;
  TE.PutValue('PEY_TYPUTI','1');
  TE.PutValue('PEY_TYPNIV', ' ');
  TE.PutValue('PEY_VALNIV', ' ');
  TE.PutValue('PEY_LIBVALNIV', ' ');
  TE.PutValue('PEY_ETABLISSEMENT', ' ');
  TE.PutValue('PEY_CODEPOP', ' ');
  TE.PutValue('PEY_SALARIE', ' ');
  if (NiveauRequis = 'ETB') then
  begin
    TE.PutValue('PEY_TYPNIV', 'ETB');
    Qt := OpenSQL('SELECT ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT = "' + TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT) + '"'  ,True);
    if not Qt.Eof then
      TE.PutValue('PEY_LIBVALNIV', Qt.FindField('ET_LIBELLE').AsString);
    ferme(Qt);
    TE.PutValue('PEY_VALNIV', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT));
    TE.PutValue('PEY_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT));
  end;
{$IFNDEF CPS1}
  if (NiveauRequis = 'POP') then
  begin
    TE.PutValue('PEY_TYPNIV', 'POP');
    Population := '';
    Qp := OpenSQL('SELECT PNA_POPULATION FROM SALARIEPOPUL '
      + ' WHERE PNA_SALARIE = "' + TOB_Salarie.GetValeur(iPSA_SALARIE) + '"'
      + ' AND PNA_TYPEPOP = "PAI"', True);  // Anciennement SAL
    if not Qp.Eof then
      Population := Qp.FindField('PNA_POPULATION').AsString;
    Ferme(Qp);

    Qt := OpenSQL('SELECT PPC_LIBELLE FROM ORDREPOPULATION WHERE PPC_POPULATION = "' + Population + '"'  ,True);
    if not Qt.Eof then
      TE.PutValue('PEY_LIBVALNIV', Qt.FindField('PPC_LIBELLE').AsString);
    ferme(Qt);
    TE.PutValue('PEY_VALNIV', Population);
    TE.PutValue('PEY_CODEPOP', Population);
  end;
{$ENDIF}
  if (NiveauRequis = 'SAL') then
  begin
    TE.PutValue('PEY_TYPNIV', 'SAL');
    TE.PutValue('PEY_VALNIV', TOB_Salarie.GetValeur(iPSA_SALARIE));
    TE.PutValue('PEY_LIBVALNIV', TOB_Salarie.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Salarie.GetValeur(iPSA_PRENOM));
    TE.PutValue('PEY_SALARIE', TOB_Salarie.GetValeur(iPSA_SALARIE));
  end;
  TE.PutValue('PEY_CODEELT', Elt);
  //DEB PT181
  sDate := DateTimeToStr(DatVal);
  if MidStr(sDate,1,2) <> '01' then
    sDate := '01' + MidStr(sDate,3,8);
  TE.PutValue('PEY_DATVAL', StrToDateTime(sDate));
//  TE.PutValue('PEY_DATVAL', DatVal);
  //FIN PT181
  TE.PutValue('PEY_MONTANT', Q.FindField('PEL_MONTANT').AsFloat);
  TE.PutValue('PEY_MONTANTEURO', Q.FindField('PEL_MONTANTEURO').AsFloat);
  TE.PutValue('PEY_PREDEFINI', ' ');

  if (NiveauRequis = 'CEG') or (NiveauRequis = 'STD') or (NiveauRequis = 'DOS') then
  begin
    TE.PutValue('PEY_PREDEFINI', NiveauRequis);
    TE.PutValue('PEY_TYPNIV', NiveauRequis);
  end;
{$IFNDEF EABSENCES}
  TE.PutValue('PEY_NODOSSIER', ClePGSynEltNAt); // Sert d'identifiant, contient un n° de salarié :
                                                // - soit celui en cours pour le traitement individuel
                                                // - soit le 1er de la liste si préparation automatique
{$ENDIF EABSENCES}

  TE.InsertOrUpdateDB;
  TobElt.Free;
  ferme(Q);
end;
//FIN PT172

{ Fonction qui rend le taux horaire stocké dans la fiche Salarié
Dans le cas ou celui ci est nul, il faut essayer de la caculer}

function TauxHoraireSal(): double;
begin
  result := TOB_Salarie.GetValeur(iPSA_TAUXHORAIRE);
end;
{ Fonction qui rend le salaire mensuel de la fiche salarié}

function SalaireMensuelSal(): double;
begin
  result := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS1);
end;
{ Fonction qui rend le salaire annuel de la fiche salarié}

function SalaireAnnuelSal(): double;
begin
  result := TOB_Salarie.GetValeur(iPSA_SALAIRANN1);
end;
{ Rend la valeur d'un cumul de date à date.   Les dates sont renseignées par la function appelante
Attention, pour le moment la fonction fait le calcul quelque soit l'etablissement
Voir s'il faut filtrer l'etablissement ou rajouter un paramètre pour traiter ou non l'etablissement
}

function ValCumulDate(const Cumul: string; const DateDebut, DateFin: TDateTime): double;
var
  Q: TQuery;
  st: string;
begin
  // PT44  group by pour forcer index ????
  // PT50 : 01/07/2002 PH V582 modif sur clacul fourchette de dates pour les variables de type cumul
  St := 'SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_SALARIE="' + CodSal + '" AND PHC_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PHC_DATEFIN<="' + UsDateTime(DateFin) +
    '" AND PHC_CUMULPAIE="' + Cumul + '"'; // GROUP BY PHC_SALARIE,PHC_DATEDEBUT,PHC_DATEFIN,PHC_REPRISE,PHC_CUMULPAIE';
  Q := OpenSQL(St, TRUE);
  if not Q.EOF then result := Q.Fields[0].AsFloat
  else result := 0;
  // PT1
  Ferme(Q);
end;
{ Rend la valeur d'une rubrique de date à date.   Les dates sont renseignées par la function appelante
Attention, pour le moment la fonction fait le calcul quelque soit l'etablissement
Voir s'il faut filtrer l'etablissement ou rajouter un paramètre pour traiter ou non l'etablissement
Optimisation de la requete SQL faite sur le serveur car c'est la requete qui fait les calculs
A la charge de la fonction de renseigner le record en fonction du type de rubrique
}

function ValRubDate(const Rubrique: string; Nature: string; const DateDebut, DateFin: TDateTime): TRendRub;
var
  Q: TQuery;
  ch: string;
  RendRub: TRendRub; // variable contenant le record retourné par la fonction
begin
  with RendRub do
  begin
    MontRem := 0;
    BasRem := 0;
    TauxRem := 0;
    CoeffRem := 0;
    BasCot := 0;
    TSal := 0;
    MSal := 0;
    TPat := 0;
    MPat := 0;
    Plfd1 := 0;
    Plfd2 := 0;
    Plfd3 := 0;
    Tr1 := 0;
    Tr2 := 0;
    Tr3 := 0;
    Base := 0;
  end;
  // PT29 : 27/12/2001 V571 PH Calcul des variables de type rémunération tq sur Mois - 1
  if Nature = 'REM' then Nature := 'AAA';

  if Nature = 'AAA' then Ch := 'SUM(PHB_MTREM),SUM(PHB_BASEREM),MAX(PHB_TAUXREM),MAX(PHB_COEFFREM)';
  if Nature = 'BAS' then Ch := 'SUM(PHB_PLAFOND1),SUM(PHB_TRANCHE1),SUM(PHB_TRANCHE2),SUM(PHB_TRANCHE3),SUM(PHB_BASECOT),SUM(PHB_PLAFOND2),SUM(PHB_PLAFOND3) ';
  if Nature = 'COT' then Ch := 'SUM(PHB_BASECOT),MAX(PHB_TAUXSALARIAL),MAX(PHB_TAUXPATRONAL),SUM(PHB_MTSALARIAL),SUM(PHB_MTPATRONAL)';
  Q := OpenSQL('SELECT ' + Ch + ' FROM HISTOBULLETIN WHERE PHB_SALARIE="' + CodSal + '" AND PHB_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin)
    + '" AND PHB_NATURERUB="' + Nature + '" AND PHB_RUBRIQUE="' + Rubrique + '"', TRUE);
  if not Q.EOF then
  begin
    with RendRub do
    begin
      if Nature = 'AAA' then
      begin
        MontRem := Q.Fields[0].AsFloat;
        BasRem := Q.Fields[1].AsFloat;
        TauxRem := Q.Fields[2].AsFloat;
        CoeffRem := Q.Fields[3].AsFloat;
      end;
      if Nature = 'BAS' then
      begin
        Plfd1 := Q.Fields[0].AsFloat;
        Tr1 := Q.Fields[1].AsFloat;
        Tr2 := Q.Fields[2].AsFloat;
        Tr3 := Q.Fields[3].AsFloat;
        Base := Q.Fields[4].AsFloat;
        Plfd2 := Q.Fields[5].AsFloat;
        Plfd3 := Q.Fields[6].AsFloat;
      end;
      if Nature = 'COT' then
      begin
        BasCot := Q.Fields[0].AsFloat;
        TSal := Q.Fields[1].AsFloat;
        TPat := Q.Fields[2].AsFloat;
        MSal := Q.Fields[3].AsFloat;
        MPat := Q.Fields[4].AsFloat;
      end;
    end;
  end;
  Ferme(Q);
  result := RendRub;
end;
{
Fonction qui calcule le numérateur du trentieme. Par defaut, le denominateur est tjrs 30.
Quand les dates correspondent à un début et une fin de mois, on a tjrs 30/30.
Quand le mois n'est pas complet alors on prend le nombre de jours travaillés comme numérateur
quelque soit le nbre de jours dans le mois (28,29,30 ou 31).
Exceptions : Si la période est supérieure à 1 mois.
Alors, il convient de calculer le nombre de mois entier calcule en trentieme et de rajouter le nombre
de jours travaillés sur le dernier mois
exemple : 010199 au 150299 Janvier =30/30 Février =15/30 soit pour 1,5 mois 45/30 et
non 31 + 15 = 46 jours.
Pour des raisons de convenance, le numérateur et le dénominateur sont modifiables dans la saisie
de bulletin pour saisir 31/31 sur Janvier.
}

function CalculTrentieme(const DateDebut, DateFin: TDateTime; TestPaie: Boolean = FALSE): Integer;
var
  calcul, NbreJour, NbreMois, Nbt: Integer;
  NMois, PremMois, PremAnnee: WORD;
  St: string;
  OkOk: Boolean;
  Q: TQuery;
begin
  calcul := 0;
  NbreMois := 0;
  NbreJour := 0;
  NMois := 0;
  PremMois := 0;
  PremAnnee := 0;
  if (DateDebut = 0) or (DateFin = 0) then
  begin
    St := 'Salarié : ' + CodSal + ' Trentième : La date de début ou la date de fin ne sont pas renseignées';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
  end
  else
  begin
    if (EstDebutMois(DateDebut)) and (EstFinMois(DateFin)) then
    begin
      NOMBREMOIS(DateDebut, DateFin, PremMois, PremAnnee, NMois);
      if NMois = 1 then calcul := 30
      else calcul := 30 * NMois;
    end
    else
    begin
      okok := DiffMoisJour(DateDebut, DateFin, NbreMois, NbreJour);
      if NbreMois <> 0 then calcul := 30 * NbreMois; // conversion du nombre de mois en trentieme de mois
      if NbreJour <> 0 then // Rajout du nombre de jour
      begin
        calcul := calcul + NbreJour;
        if not OkOk then calcul := calcul + 1;
      end;
    end;
  end;
  // DEB PT117
  if (not EstDebutMois(DateDebut)) and TestPaie then
  begin // Cas ou il y a une paie antérieure donc on va limiter à 30 trentième
    st := 'SELECT SUM (PPU_NUMERATTRENT) NBTRENT FROM PAIEENCOURS WHERE PPU_SALARIE = "' + CodSal + '" AND PPU_DATEDEBUT >="' +
      UsDatetime(DEBUTDEMOIS(DateDebut)) + '" AND PPU_DATEFIN < "' + UsDatetime(DateDebut) + '"';
    Q := OpenSQl(St, TRUE);
    if not Q.EOF then Nbt := Q.FindField('NBTRENT').AsInteger
    else Nbt := 0;
    FERME(Q);
    if Calcul + Nbt > 30 then Calcul := 30 - Nbt;
  end; // FIN PT117
  result := calcul;
end;
// Fonction qui indique si la date est un début de mois

{ pt187 function EstDebutMois(const LaDate: TDateTime): boolean;
var
  UneDate: TDateTime;
begin
  result := FALSE;
  UneDate := LaDate;
  if UneDate = DebutdeMois(LaDate) then result := TRUE;
end;    }
// Fonction qui indique si la date est une fin de mois

{ pt187 function EstFinMois(const LaDate: TDateTime): boolean;
var
  UneDate: TDateTime;
begin
  result := FALSE;
  UneDate := LaDate;
  if UneDate = FindeMois(LaDate) then result := TRUE;
end; }
{ Fonction qui calcule le nombre de mois entre 2 dates et le nombre de jours restants
NbreMois contient le nombre de mois entier entre les 2 dates et NbreJour contient le nombre de
jours entre la date de fin et le debut du mois concernant la date de fin
}

{pt187 function DiffMoisJour(const DateDebut, DateFin: TDateTime; var NbreMois, NbreJour: Integer): Boolean;
var
  PremMois, PremAnnee, NMois: WORD;
  DateCal: TDateTime;
  Calcul: double;
begin
  result := FALSE;
  PremMois := 0;
  PremAnnee := 0;
  NMois := 0;
  Calcul := 0; // PT31
  if DateDebut = DateFin then
  begin
    NbreJour := 1;
    NbreMois := 0;
    result := TRUE;
    exit;
  end;
  NOMBREMOIS(DateDebut, DateFin, PremMois, PremAnnee, NMois);
  //if (EstFinMois(DateFin)) then NMois:=NMois-1; // Car prend en compte le mois non complet
  if NMois > 0 then NMois := NMois - 1;
  NbreMois := NMois;
  if NMois = 0 then
  begin
    Calcul := DateFin - DateDebut;
    NbreJour := StrToInt(FloatToStr(Calcul));
    exit;
  end;
  {DateCal:=FindeMois (DateDebut);
  Calcul:=DateCal-DateDebut;}
  // 3 cas à gérer
 { if (EstDebutMois(DateDebut)) and (not EstFinMois(DateFin)) then
  begin
    DateCal := DebutdeMois(DateFin);
    Calcul := DateFin - DateCal;
  end;
  if (not EstDebutMois(DateDebut)) and (EstFinMois(DateFin)) then
    Calcul := (FinDeMois(DateDebut)) - DateDebut;
  if Calcul >= 31 then Calcul := 30;
  if (not EstDebutMois(DateDebut)) and (not EstFinMois(DateFin)) then
  begin // On donne simplement le nombre de jours entre les 2 dates
    Calcul := DateFin - Datedebut;
    NbreMois := 0;
    result := TRUE;
  end;
  if Calcul = 0 then
  begin
    result := TRUE;
    Calcul := 1;
  end; // La date de fin correspond à un début de mois donc 1 trentieme
  // calcul ne peut contenir qu'une valeur entiere <= 31 jours ou > selon les 3 cas si on est sur plusieurs mois
  NbreJour := StrToInt(FloatToStr(Calcul));
end;   fin pt187}

// Fonction qui rend le Nombre de mois D'anciennete

function AncienneteMois(const DateEntree, DateFin: TDateTime): WORD;
var
  PremMois, PremAnnee: WORD;
begin
  PremMois := 0;
  PremAnnee := 0;
  result := 0;
  if DateEntree = 0 then result := 0
    //PT82-2 On tient compte du nombre de mois entier
  else
  begin
    AglNombreDeMoisComplet(DateEntree, DateFin, PremMois, PremAnnee, result);
    if DateEntree = DateFin then result := 0; //PT96
  end;
end;
// Fonction qui rend le Nombre d'année D'anciennete, elle calcule l'anciennete en mois et divise par 12

function AncienneteAnnee(const DateEntree, DateFin: TDateTime): WORD;
begin
  result := (AncienneteMois(DateEntree, DateFin)) div 12;
end;
// Fonction qui rend le nombre d'année et de mois pour le calcul de l'age du salarié

procedure AgeSalarie(const DateFin: TDateTime; var Annee, Mois: WORD);
var
  PremMois, PremAnnee: WORD;
  DateNaissance: TDateTime;
begin
  PremMois := 0;
  PremAnnee := 0;
  DateNaissance := TOB_Salarie.GetValeur(iPSA_DATENAISSANCE);
  Annee := AncienneteAnnee(DateNaissance, DateFin);
  NOMBREMOIS(DateNaissance, DateFin, PremMois, PremAnnee, Mois);
end;
// Rend la valeur du cumul salarié calculé dans la session de paie

function RendCumulSalSess(const Cumul: string): double;
var
  TRech: TOB;
begin
  result := 0;
  if TOB_CumulSal = nil then exit;
  TRech := Paie_RechercheOptimise(TOB_CumulSal, 'PHC_CUMULPAIE', Cumul); // $$$$
  if Assigned(TRech) and (iPHC_MONTANT = 0) then MemorisePhc(TRech);
  if TRech <> nil then Result := TRech.GetValeur(iPHC_MONTANT);
end;
{ Recherche en cascade d'un profil à partir du salarié à l'établissement
La fonction a pour but de rechercher le profil mis par défaut au niveau de l'établissement
dans le cas où le salarié n'est pas personnalisé
Profil FNAL : Idem Dossier donc lecture des paramètres de la paie
}

procedure RechercheProfil(const Champ1, Champ2: string; Salarie, T_Etab, TPE: TOB);
var
  suffixe, Profil: string;
begin
  suffixe := ExtractSuffixe(Champ2); // recup du suffixe du champ profil à traiter
  Profil := string(Salarie.GetValue(Champ1)); // Type Profil : Idem Etab ou Personnalisé
  {  PT-10 07/09/01 V547 PH Test du profil FNAL si Profil '' et non FNAL alors on recherche idem etab
   ceci pour compenser le fait que la zone n'est pas saisie ou mal initialisée ne provoque pas d'erreur mais
   une exception en vision SAV
  }
  if ((Profil = 'ETB') or (Profil = '')) and (suffixe <> 'PROFILFNAL') then
  begin
    if T_Etab.GetValue('ETB_' + suffixe) <> NULL then Profil := T_Etab.GetValue('ETB_' + suffixe)
    else Profil := '';
  end
  else
  begin // Il n'y a que le profil FNAL qui soit paramètrable au niveau du dossier
    if (suffixe = 'PROFILFNAL') and (Profil = 'DOS') then Profil := string(GetParamSoc('SO_PG' + suffixe))
    else Profil := string(Salarie.GetValeur(NumChampProfS(Champ2)));
  end;
  if Profil <> '' then ChargeProfil(Salarie, TPE, Profil);
end;


{ Fonction qui compare la liste TPE des rubriques provenant des profils du salariés
et une rubrique du bulletin précédent TPR. Si une rubrique n'existe pas dans TPE
mais existe dans TPR alors on l'a prend si PHB_CONVERSATION=TRUE.
On en profite pour traiter chaque rubrique de TPR pour voir les éléménts permanents
et si on en trouve on récupère les valeurs.
Au départ les rubriques en provenance des profils ne sont pas valorisées, c'est la fonction
de calcul du bulletin qui valorise les lignes en fonction des paramètrages. Il reste à la
fonction de calcul du bulletin de valoriser toutes les rubriques en fonction des éléments
nationaux, remunération, cotisations et variable.
On récupère aussi les libellés modifiés des rubriques composant le bulletin precédant
}
procedure MajRubriqueBulletin(TPE, TPR: TOB; const CodSal, Etab: string; const DateDeb, DateFin: TDateTime);
var
THB, TRech: TOB;
Rubrique, NatureRub, Libel: string;
begin
// PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
// dans ce cas on ne recherche aucun élément sur une paie antérieure
if (BullCompl='X') then
   exit;

Rubrique:= TPR.GetValue ('PHB_RUBRIQUE');
NatureRub:= TPR.GetValue ('PHB_NATURERUB');
if (TPR.GetValue ('PHB_COTREGUL')='REG') and (NatureRub='COT') then
Libel:= TPR.GetValue ('PHB_LIBELLE');
TRech:= TPE.FindFirst (['PHB_NATURERUB', 'PHB_RUBRIQUE'], [NatureRub, Rubrique], TRUE);
if (TRech<>nil) then
   begin // Rubrique existante = recup libelle
   TRech.PutValue ('PHB_LIBELLE', Libel);
   if (NatureRub='AAA') then
      RecupEltPermanent (TRech, TPR, Rubrique); // les éléments permanents ne sont que sur les Rémunérations
   end
else
   begin // Voir pour le cas de suppression de ligne de profils
   if (TPR.GetValue ('PHB_CONSERVATION')='SAL') then // Cas ou la rubrique est à conserver
      begin
      THB:= TOB.Create ('HISTOBULLETIN', TPE, -1);
      THB.Dupliquer (TPR, FALSE, TRUE, TRUE);
//  ok:=TPR.ChangeParent (THB, -1); // recup de la ligne bulletin dans la totalité
      THB.PutValue ('PHB_ETABLISSEMENT', Etab);
      THB.PutValue ('PHB_SALARIE', CodSal);
      THB.PutValue ('PHB_DATEDEBUT', DateDeb);
      THB.PutValue ('PHB_DATEFIN', DateFin);
// DEB PT95 rajout appel fonction recup des élements qui remet aussi à 0 les champs concernés si ils concernent des éléments variables
      if (NatureRub='AAA') then
         RecupEltPermanent (THB, TPR, Rubrique, 'SAL');
// FIN PT95
      end;
   end;
end;

{ Fonction qui récupère les éléments permanents de la paie précédente
Elle permet aussi de récupèrer en fonction du paramètrage société de la récupération des
éléments permanents
PT95   : 12/01/2004 PH V_50 FQ 11024 Rubrique sur le bulletin en origine salarié reprend les valeurs même si elt variable
Donc si on a une rubrique d'origine salarié alors on regarde si c'est un élement variable alors
on remet à 0 les zones concernées.
}
// DEB PT95

procedure RecupEltPermanent(TPE, TPR: TOB; const Rubrique: string; Conservation: string = '');
var
  Rub: TOB;
  TypB, TypC, TypT, TypM: string; // Type des differents champs composant une rémunération
begin
  Rub := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', Rubrique); // $$$$
  if Rub <> nil then
  begin
    TypB := Rub.GetValue('PRM_TYPEBASE');
    TypT := Rub.GetValue('PRM_TYPETAUX');
    TypC := Rub.GetValue('PRM_TYPECOEFF');
    TypM := Rub.GetValue('PRM_TYPEMONTANT');
    if TypB = '01' then TPE.PutValue('PHB_BASEREM', TPR.GetValue('PHB_BASEREM'))
    else if (Conservation = 'SAL') and (TypB = '00') then TPE.PutValue('PHB_BASEREM', 0);
    if TypT = '01' then TPE.PutValue('PHB_TAUXREM', TPR.GetValue('PHB_TAUXREM'))
    else if (Conservation = 'SAL') and (TypT = '00') then TPE.PutValue('PHB_TAUXREM', 0);
    if TypC = '01' then TPE.PutValue('PHB_COEFFREM', TPR.GetValue('PHB_COEFFREM'))
    else if (Conservation = 'SAL') and (TypC = '00') then TPE.PutValue('PHB_COEFFREM', 0);
    if TypM = '01' then TPE.PutValue('PHB_MTREM', TPR.GetValue('PHB_MTREM'))
    else if (Conservation = 'SAL') and (TypM = '00') then TPE.PutValue('PHB_MTREM', 0);
  end;
end;
// FIN PT95

procedure AligneProfil(Salarie, T_Etab, TPE: TOB; ActionB: TActionBulletin);
var
  Q: TQuery;
  ThemeProfil: string;
begin
  ThemeProfil := '';
  //DEB PT106
  if not ForceAlignProfil and (ActionB = taModification) then exit;
  ForceAlignProfil := FALSE;
  //FIN PT106
    // Chargement des rubriques de chaque profil reférencé au niveau salarié
    //PT-12 02/10/01 V562 PH Rajout Profil Rémunération du Salarié
  RechercheProfil('PSA_TYPPROFILREM', 'PSA_PROFILREM', Salarie, T_Etab, TPE);
  // Profil Salarié modèle de bulletin
  RechercheProfil('PSA_TYPPROFIL', 'PSA_PROFIL', Salarie, T_Etab, TPE);
  // Profil Periodicite bulletin
  RechercheProfil('PSA_TYPPERIODEBUL', 'PSA_PERIODBUL', Salarie, T_Etab, TPE);
  // Profil Réduction Bas Salaire
  RechercheProfil('PSA_TYPPROFILRBS', 'PSA_PROFILRBS', Salarie, T_Etab, TPE);
  // Profil Réduction REPAS
  RechercheProfil('PSA_TYPREDREPAS', 'PSA_REDREPAS', Salarie, T_Etab, TPE);
  // Profil Réduction RTT 1 Loi Aubry 2
  RechercheProfil('PSA_TYPREDRTT1', 'PSA_REDRTT1', Salarie, T_Etab, TPE);
  // Profil Réduction RTT 2 Loi Robien
  RechercheProfil('PSA_TYPREDRTT2', 'PSA_REDRTT2', Salarie, T_Etab, TPE);
  // Profil Temps Partiel
  ChargeProfil(Salarie, TPE, Salarie.GetValeur(iPSA_PROFILTPS));
  // Profil Abattement pour frais Professionnels
  RechercheProfil('PSA_TYPPROFILAFP', 'PSA_PROFILAFP', Salarie, T_Etab, TPE);
  // Profil Gestion des Appoints
  RechercheProfil('PSA_TYPPROFILAPP', 'PSA_PROFILAPP', Salarie, T_Etab, TPE);
  // Profil Maladie Maintien Salaire supprimé car inclus dans les profils temporaires du bulletin
  //RechercheProfil ('PSA_TYPPROFILMMS', 'PSA_PROFILMMS', Salarie,T_Etab,TPE, ThemRubExclus);
  // PT37 : 02/04/2002 PH V571 traitement du profil retraite
  // Profil retraite
  RechercheProfil('PSA_TYPPROFILRET', 'PSA_PROFILRET', Salarie, T_Etab, TPE);
  // PRofil Mutuelle
  RechercheProfil('PSA_TYPPROFILMUT', 'PSA_PROFILMUT', Salarie, T_Etab, TPE);
  // Profil Prévoyance
  RechercheProfil('PSA_TYPPROFILPRE', 'PSA_PROFILPRE', Salarie, T_Etab, TPE);
  // Profil Taxe sur les Salaires
  RechercheProfil('PSA_TYPPROFILTSS', 'PSA_PROFILTSS', Salarie, T_Etab, TPE);
  // Profil Repos Compensateur
  // ReChercheProfil ('PSA_TYPPROFILRCO', 'PSA_PROFILRCO', Salarie,T_Etab,TPE);
  // Profil Congés Payés
  RechercheProfil('PSA_TYPPROFILCGE', 'PSA_PROFILCGE', Salarie, T_Etab, TPE);
  // Profil fin de CDD
  ChargeProfil(Salarie, TPE, Salarie.GetValeur(iPSA_PROFILCDD));
  // Profil Multi Employeur
  ChargeProfil(Salarie, TPE, Salarie.GetValeur(iPSA_PROFILMUL));
  // Profil Multi FNAL
  RechercheProfil('PSA_TYPPROFILFNAL', 'PSA_PROFILFNAL', Salarie, T_Etab, TPE);
  // Profil Multi Transport
  RechercheProfil('PSA_TYPPROFILTRANS', 'PSA_PROFILTRANS', Salarie, T_Etab, TPE);
  // Profil Ancienneté
  RechercheProfil('PSA_TYPPROFILANC', 'PSA_PROFILANCIEN', Salarie, T_Etab, TPE);

  // Cas des profils temporaires cad maladie, chomage qui sont uniquement sur le bulletin
  // Chargement des rubriques pour les profils CAS Particuliers dans ce cas
  // Boucle de chargement sur 2 ou 3 lignes maxi
  Q := OPENSQL('SELECT * FROM PROFILSPECIAUX WHERE PPS_ETABSALARIE="-" AND PPS_CODE="' + Salarie.GetValeur(iPSA_SALARIE) + '"', TRUE);
  while not Q.EOF do
  begin
    ThemeProfil := Q.FindField('PPS_THEMEPROFIL').AsString;
    ChargeProfil(Salarie, TPE, Q.FindField('PPS_PROFIL').AsString);
    Q.Next;
  end;
  Ferme(Q);
end;

{ Fonction d'alimentation des cumuls du salarié pour une rubrique
Prépare la structure HISTOCUMSAL qui donnera la possibilité d'écrire les enregistrements
dans la table.
}

procedure AlimCumulSalarie(THB: TOB; const Salarie, NatureRub, Rubrique, Etab: string; const DateDebut, DateFin: TDateTime);
var
  TRechRub, TRechCum, TRechCumRub, TRechCumSal: TOB;
  Cumul, TypeAlim, Sens, St, LeCoeff, Rub: string;
  I: Integer;
  Montant, MtSupp, Coeff: Double;
begin
  MtSupp := 0; // Montant à ajouter/supprimer du cumul salarié
  Cumul := '';
  TypeAlim := '';
  Sens := '';
  Rub := Rubrique;
  if (Copy(Rub, 5, 2) = '.R') then
    Rub := Copy(Rub, 1, 4);
  TRechRub := Paie_RechercheRubrique(NatureRub, Rub);
  if not Assigned(TRechRub) then exit; // Rubrique non trouvée ne devrait jamais se produire
  if (TRechRub.detail.count > 0) and (iPCR_SENS = 0) then MemorisePcr(TRechRub.detail[0]);
  if Assigned(THB) and (iPHB_BASEREM = 0) then MemorisePhb(THB);

  for I := 0 to TRechRub.Detail.Count - 1 do // boucle sur la recherche des cumuls alimentés par la rubrique
  begin
    TRechCumRub := TRechRub.Detail[I];
    Cumul := TRechCumRub.GetValeur(iPCR_CUMULPAIE);
    if Cumul = '' then
    begin
      St := 'Salarié : ' + CodSal + ' Cumul Erroné';
      if TypeTraitement <> 'PREPA' then PGIBox(St, '') else TraceErr.Items.Add(St);
    end;

    Sens := TRechCumRub.GetValeur(iPCR_SENS); // Sens alimentation du cumul
    TRechCum := Paie_RechercheOptimise(TOB_Cumuls, 'PCL_CUMULPAIE', Cumul); // $$$$
    if Assigned(TRechCum) and (iPCL_ALIMCUMUL = 0) then MemorisePcl(TRechCum);

    if TrechCum <> nil then
    begin // Cumul trouvé et récupération du type d'alimentation en fonction du type de rubrique
      if NatureRub = 'AAA' then TypeAlim := TRechCum.GetValeur(iPCL_ALIMCUMUL)
      else TypeAlim := TRechCum.GetValeur(iPCL_ALIMCUMULCOT);

      if (TypeAlim <> '') and (TypeAlim <> 'ZZZ') then
      begin
        TRechCumSal := TOB_CumulSal.FindFirst(['PHC_ETABLISSEMENT', 'PHC_SALARIE', 'PHC_DATEDEBUT', 'PHC_DATEFIN', 'PHC_REPRISE', 'PHC_CUMULPAIE'],
          [Etab, Salarie, DateDebut, DateFin, '-', Cumul], FALSE); // $$$$
        if TRechCumSal = nil then
        begin // Création du cumul salarié avec initialisation de tous les champs
          Montant := 0;
          TRechCumSal := TOB.Create('HISTOCUMSAL', TOB_CumulSal, -1);
          with TRechCumSal do
          begin // Optimisation
            if iPHC_ETABLISSEMENT = 0 then MemorisePhc(TRechCumSal);

            PutValeur(iPHC_ETABLISSEMENT, Etab);
            PutValeur(iPHC_SALARIE, Salarie);
            PutValeur(iPHC_DATEDEBUT, DateDebut);
            PutValeur(iPHC_DATEFIN, DateFin);
            PutValeur(iPHC_REPRISE, '-');
            PutValeur(iPHC_CUMULPAIE, Cumul);
            PutValeur(iPHC_MONTANT, Montant);
            PutValeur(iPHC_TRAVAILN1, TOB_Salarie.GetValeur(iPSA_TRAVAILN1));
            PutValeur(iPHC_TRAVAILN2, TOB_Salarie.GetValeur(iPSA_TRAVAILN2));
            PutValeur(iPHC_TRAVAILN3, TOB_Salarie.GetValeur(iPSA_TRAVAILN3));
            PutValeur(iPHC_TRAVAILN4, TOB_Salarie.GetValeur(iPSA_TRAVAILN4));
            PutValeur(iPHC_CODESTAT, TOB_Salarie.GetValeur(iPSA_CODESTAT));
            PutValeur(iPHC_CONFIDENTIEL, TOB_Salarie.GetValeur(iPSA_CONFIDENTIEL));
            PutValeur(iPHC_LIBREPCMB1, TOB_Salarie.GetValeur(iPSA_LIBREPCMB1));
            PutValeur(iPHC_LIBREPCMB2, TOB_Salarie.GetValeur(iPSA_LIBREPCMB2));
            PutValeur(iPHC_LIBREPCMB3, TOB_Salarie.GetValeur(iPSA_LIBREPCMB3));
            PutValeur(iPHC_LIBREPCMB4, TOB_Salarie.GetValeur(iPSA_LIBREPCMB4));
          end;
        end
        else Montant := TRechCumSal.GetValeur(iPHC_MONTANT);

        if NatureRub = 'AAA' then // Rémunération
        begin
          if TypeAlim = 'BAS' then MtSupp := THB.GetValeur(iPHB_BASEREM)
          else if TypeAlim = 'MSA' then MtSupp := THB.GetValeur(iPHB_MTREM)
          else if TypeAlim = 'TAU' then MtSupp := THB.GetValeur(iPHB_TAUXREM); // PT178
        end
        else if NatureRub = 'BAS' then // Base de Cotisation
        begin
          if TypeAlim = 'BAS' then MtSupp := THB.GetValeur(iPHB_BASECOT); // @@@@ Base + Tranches + Palfond
        end
        else if NatureRub = 'COT' then // Cotisation
        begin
          if TypeAlim = 'BAS' then MtSupp := THB.GetValeur(iPHB_BASECOT)
          else if TypeAlim = 'MSA' then MtSupp := THB.GetValeur(iPHB_MTSALARIAL)
          else if TypeAlim = 'MPT' then MtSupp := THB.GetValeur(iPHB_MTPATRONAL)
          else if TypeAlim = 'TRB' then MtSupp := THB.GetValeur(iPHB_TRANCHE2)
          else if TypeAlim = 'TRC' then MtSupp := THB.GetValeur(iPHB_TRANCHE3);
        end;
        // Recherche du coefficient à appliquer ex CSG 95%
        LeCoeff := TRechCum.GetValeur(iPCL_COEFFAFFECT);
        if LeCoeff <> '' then
        begin
          Coeff := ValEltNat(LeCoeff, DateFin, nil); //PT172
          if Coeff <> 0 then MtSupp := MtSupp * Coeff;
        end;

        if Sens = '+' then MtSupp := Montant + MtSupp else MtSupp := Montant - MtSupp;

        TRechCumSal.PutValeur(iPHC_MONTANT, ARRONDI(MtSupp, 2));
      end; // fin si TypeAlim reconnu
    end; // fin si cumul existe
  end; // fin boucle sur la liste des cumuls alimentés par une rubrique
end;
{ Procedure RAZ de la TOB TPE de la paie en cours
Elle détruit les filles (lignes du bulletin) de la TOB TPE
Cette procédure est utilisée pour la saisie du bulletin dans les cas du déchargement des lignes afin
de pouvoir historiser les lignes du bulletins
}

procedure RazTPELignes(TPE: TOB);
begin
  TPE.ClearDetail;
end;
{ Fonction qui alimente la liste des rubriques du bulletin en fonction d'une grille de données
}

procedure GrilleAlimLignes(TPE: TOB; Etab, Salarie, Nature: string; DateDebut, DateFin: TDateTime; GrilleBulletin: THGrid; ActionBul: TActionBulletin);
var
  i: Integer;
  Rubrique, Libelle, St, LaRubrique : string;
  THB, T_Rub, T_Com, TSal: TOB;
  ObjetACreer: Boolean;
  TRech:TOB;
begin
  if (Etab = '') or (Salarie = '') or (Nature = '') or (DateDebut = 0) or (DateFIN = 0) then
  begin
    St := 'Salarié : ' + CodSal + ' Erreur Champs non renseignés dans la Grille';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
  end;

  TSal := TOB_Salarie;
  if TSal = nil then
  begin
    // TOB_Salarie:=TOB.Create('SALARIES',Nil,-1) ;
    // TOB_Salarie.SelectDB('"'+CodSal+'"',Nil,TRUE) ;
    RecupTobSalarie(CodSal, DateDebut, DateFin); // PT173
    TSal := TOB_Salarie;
  end;
  Libelle := '';
  for i := 1 to GrilleBulletin.RowCount - 1 do
  begin
    Rubrique := GrilleBulletin.Cells[0, i];
    Libelle := GrilleBulletin.Cells[1, i];
    if Rubrique <> '' then
    begin
      //DEB PT180
      if (POS('.R', Rubrique) > 0) then
        LaRubrique := copy(Rubrique, 1, 4);
      TRech := Paie_RechercheOptimise(TOB_REM, 'PRM_RUBRIQUE', LaRubrique);
      if (not assigned(trech)) and (Libelle = '') then
      begin
        St := 'La rubrique ' + LaRubrique + ' n''existe pas. La ligne a été supprimée ';
        ShowMessage(St);
        continue;
      end;
      //FIN PT180

      ///   THB:=TOB.Create('HISTOBULLETIN',TPE,-1) ;
      //  if (ActionBul <> taCreation) AND (ActionBul<> PremCreation) then THB:=TOB(GrilleBulletin.Objects[0,i]) ;
      THB := TOB(GrilleBulletin.Objects[0, i]);
      ObjetACreer := FALSE;
      if THB = nil then
      begin
        THB := TOB.Create('HISTOBULLETIN', TPE, -1);
        ObjetACreer := TRUE;
      end;
      if Assigned(THB) and (iPHB_ETABLISSEMENT = 0) then MemorisePhb(THB);
      with THB do
      begin
        PutValeur(iPHB_ETABLISSEMENT, Etab);
        PutValeur(iPHB_SALARIE, Salarie);
        PutValeur(iPHB_DATEDEBUT, DateDebut);
        PutValeur(iPHB_DATEFIN, DateFin);
        PutValeur(iPHB_NATURERUB, Nature);
        PutValeur(iPHB_RUBRIQUE, Rubrique);
        PutValeur(iPHB_LIBELLE, Libelle);
        if TSal <> nil then
        begin
          PutValeur(iPHB_TRAVAILN2, TSal.GetValeur(iPSA_TRAVAILN2));
          PutValeur(iPHB_TRAVAILN3, TSal.GetValeur(iPSA_TRAVAILN3));
          PutValeur(iPHB_TRAVAILN4, TSal.GetValeur(iPSA_TRAVAILN4));
          PutValeur(iPHB_TRAVAILN1, TSal.GetValeur(iPSA_TRAVAILN1));
          PutValeur(iPHB_CODESTAT, TSal.GetValeur(iPSA_CODESTAT));
          PutValeur(iPHB_CONFIDENTIEL, TSal.GetValeur(iPSA_CONFIDENTIEL));
          PutValeur(iPHB_LIBREPCMB1, TSal.GetValeur(iPSA_LIBREPCMB1));
          PutValeur(iPHB_LIBREPCMB2, TSal.GetValeur(iPSA_LIBREPCMB2));
          PutValeur(iPHB_LIBREPCMB3, TSal.GetValeur(iPSA_LIBREPCMB3));
          PutValeur(iPHB_LIBREPCMB4, TSal.GetValeur(iPSA_LIBREPCMB4));
        end
        else
        begin
          PutValeur(iPHB_TRAVAILN2, '');
          PutValeur(iPHB_TRAVAILN3, '');
          PutValeur(iPHB_TRAVAILN4, '');
          PutValeur(iPHB_TRAVAILN1, '');
          PutValeur(iPHB_CODESTAT, '');
          PutValeur(iPHB_CONFIDENTIEL, '0');
          PutValeur(iPHB_LIBREPCMB1, '');
          PutValeur(iPHB_LIBREPCMB2, '');
          PutValeur(iPHB_LIBREPCMB3, '');
          PutValeur(iPHB_LIBREPCMB4, '');
        end;
        if Nature = 'AAA' then
        begin // Rémunération
          if (Copy(Rubrique, 5, 2) = '.R') then PutValeur(iPHB_COTREGUL, 'REG') // PT165
          else PutValeur(iPHB_COTREGUL, '...'); // PT165

          T_RUB := Paie_RechercheOptimise(TOB_REM, 'PRM_RUBRIQUE', Copy(Rubrique, 1, 4)); // PT165
          if Assigned(T_RUB) and (iPRM_IMPRIMABLE = 0) then MemorisePrm(T_RUB); // Optimisation // BEURK
          if T_RUB <> nil then
          begin
            PutValeur(iPHB_IMPRIMABLE, T_RUB.GetValeur(iPRM_IMPRIMABLE));
            PutValeur(iPHB_BASEREM, Valeur(GrilleBulletin.Cells[2, i]));
            PutValeur(iPHB_TAUXREM, Valeur(GrilleBulletin.Cells[3, i]));
            PutValeur(iPHB_COEFFREM, Valeur(GrilleBulletin.Cells[4, i]));
            PutValeur(iPHB_MTREM, Valeur(GrilleBulletin.Cells[5, i]));
            //    THB.PutValue('PHB_PREDEFINI', T_RUB.GetValue('PRM_PREDEFINI'));
            PutValeur(iPHB_IMPRIMABLE, T_RUB.GetValeur(iPRM_IMPRIMABLE));
            PutValeur(iPHB_BASEREMIMPRIM, T_RUB.GetValeur(iPRM_BASEIMPRIMABLE));
            PutValeur(iPHB_TAUXREMIMPRIM, T_RUB.GetValeur(iPRM_TAUXIMPRIMABLE));
            PutValeur(iPHB_COEFFREMIMPRIM, T_RUB.GetValeur(iPRM_COEFFIMPRIM));
            PutValeur(iPHB_ORDREETAT, T_RUB.GetValeur(iPRM_ORDREETAT));
            PutValeur(iPHB_OMTSALARIAL, T_RUB.GetValeur(iPRM_ORDREETAT)); { PT136 }
            //PT22 : 22/11/01 V562 PH ORACLE, Le champ PHB_ORGANISME doit être renseigné
            PutValeur(iPHB_ORGANISME, '....');
// PT165    PutValeur(iPHB_COTREGUL, '...'); // PT152
            PutValeur(iPHB_SENSBUL, T_RUB.GetValeur(iPRM_SENSBUL));
            PutValeur(iPHB_CONSERVATION, AnsiUpperCase(Copy(GrilleBulletin.CellValues[6, i], 1, 3)));
          end
          else // Rubrique de Remunération non trouvée, on regarde si ce n'est pas un commentaire
          begin
            if RechCommentaire(Rubrique) = TRUE then
            begin // Ordre de presentation de la rubrique lors de l'edition du bulletin est identique à la Rem dont elle dépend
              PutValeur(iPHB_IMPRIMABLE, 'X'); // Rubrique de commentaire tjrs imprimable
              PutValeur(iPHB_COTREGUL, '...'); //PT152
              // PT24   27/11/2001 V563 PH Propagation d'une ligne de commentaire
              // Memorisation pour savoir si on doit reprendre la rubrique de commentaire sur le bulletin suivant
              if Copy(GrilleBulletin.CellValues[6, i], 1, 3) <> '' then
                PutValeur(iPHB_CONSERVATION, AnsiUpperCase(Copy(GrilleBulletin.CellValues[6, i], 1, 3)));
              T_Com := Paie_RechercheOptimise(TOB_REM, 'PRM_RUBRIQUE', RendCodeRub(Rubrique)); // $$$$
              if T_Com <> nil then
              begin
                PutValeur(iPHB_ORDREETAT, T_Com.GetValeur(iPRM_ORDREETAT));
                PutValeur(iPHB_OMTSALARIAL, T_Com.GetValeur(iPRM_ORDREETAT)); { PT136 }
              end;
            end;
          end;
        end
        else
        begin
          if (Copy(Rubrique, 5, 2) = '.R') then PutValeur(iPHB_COTREGUL, 'REG')
          else PutValeur(iPHB_COTREGUL, '...'); // PT152

          T_Rub := Paie_RechercheRubrique(Nature, Copy(Rubrique, 1, 4));
          if Assigned(T_Rub) and (iPCT_IMPRIMABLE = 0) then MemorisePct(T_Rub);
{$IFDEF aucasou}
          if Nature = 'COT' then T_RUB := TOB_Cotisations.FindFirst(['PCT_NATURERUB', 'PCT_RUBRIQUE'], [Nature, Rubrique], FALSE) // $$$$
          else T_RUB := TOB_Bases.FindFirst(['PCT_NATURERUB', 'PCT_RUBRIQUE'], [Nature, Rubrique], FALSE); // $$$$
{$ENDIF}
          if assigned(t_rub) then
          begin
            PutValeur(iPHB_IMPRIMABLE, T_RUB.GetValeur(iPCT_IMPRIMABLE));
            PutValeur(iPHB_BASECOTIMPRIM, T_RUB.GetValeur(iPCT_BASEIMP));
            PutValeur(iPHB_TAUXSALIMPRIM, T_RUB.GetValeur(iPCT_TXSALIMP));
            //    THB.PutValue('PHB_MTSALIMPRIM',T_RUB.GetValue('PCT_FFSALIMP')) ;
            PutValeur(iPHB_TAUXPATIMPRIM, T_RUB.GetValeur(iPCT_TXPATIMP));
            //    THB.PutValue('PHB_MTPATIMPRIM',T_RUB.GetValue('PCT_FFPATIMP')) ;
            PutValeur(iPHB_ORGANISME, T_RUB.GetValeur(iPCT_ORGANISME));
            PutValeur(iPHB_ORDREETAT, 3); { PT136 }
            PutValeur(iPHB_OMTSALARIAL, T_RUB.GetValeur(iPCT_ORDREETAT)); { PT136 }
            PutValeur(iPHB_SENSBUL, 'P');
          end
          else
          begin
            if RechCommentaire(Rubrique) = TRUE then
            begin // Ordre de presentation de la rubrique lors de l'edition du bulletin est identique à la Rem dont elle dépend
              PutValeur(iPHB_IMPRIMABLE, 'X'); // Rubrique de commentaire tjrs imprimable
              // PT24   27/11/2001 V563 PH Propagation d'une ligne de commentaire
              // Memorisation pour savoir si on doit reprendre la rubrique de commentaire sur le bulletin suivant
              if Copy(GrilleBulletin.CellValues[6, i], 1, 3) <> '' then
                PutValeur(iPHB_CONSERVATION, AnsiUpperCase(Copy(GrilleBulletin.CellValues[7, i], 1, 3)));
              T_Com := Paie_RechercheOptimise(TOB_Cotisations, 'PCT_RUBRIQUE', RendCodeRub(Rubrique)); // $$$$
{$IFDEF aucasou}
              T_Com := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE'], [RendCodeRub(Rubrique)], FALSE); // $$$$
{$ENDIF}
              if T_Com <> nil then
              begin
                PutValeur(iPHB_ORDREETAT, 3); { PT136 }
                PutValeur(iPHB_OMTSALARIAL, T_Com.GetValeur(iPCT_ORDREETAT)); { PT136 }
              end;
            end;
          end;

          if Nature = 'COT' then
          begin
            PutValeur(iPHB_BASECOT, Valeur(GrilleBulletin.Cells[2, i]));
            PutValeur(iPHB_TAUXSALARIAL, Valeur(GrilleBulletin.Cells[3, i]));
            PutValeur(iPHB_MTSALARIAL, Valeur(GrilleBulletin.Cells[4, i]));
            PutValeur(iPHB_TAUXPATRONAL, Valeur(GrilleBulletin.Cells[5, i]));
            PutValeur(iPHB_MTPATRONAL, Valeur(GrilleBulletin.Cells[6, i]));
            PutValeur(iPHB_CONSERVATION, AnsiUpperCase(Copy(GrilleBulletin.CellValues[7, i], 1, 3)));
          end;
          if Nature = 'BAS' then
          begin
            PutValeur(iPHB_BASECOT, Valeur(GrilleBulletin.Cells[2, i]));
            PutValeur(iPHB_PLAFOND, Valeur(GrilleBulletin.Cells[3, i]));
            PutValeur(iPHB_TRANCHE1, Valeur(GrilleBulletin.Cells[4, i]));
            PutValeur(iPHB_TRANCHE2, Valeur(GrilleBulletin.Cells[5, i]));
            PutValeur(iPHB_TRANCHE3, Valeur(GrilleBulletin.Cells[6, i]));
            PutValeur(iPHB_PLAFOND1, Valeur(GrilleBulletin.Cells[7, i]));
            PutValeur(iPHB_PLAFOND2, Valeur(GrilleBulletin.Cells[8, i]));
            PutValeur(iPHB_PLAFOND3, Valeur(GrilleBulletin.Cells[9, i]));
            // PT78  : 05/06/03 PH V_421 FQ 10700 Gestion origine de la rubrique dans les bases de cotisations
            PutValeur(iPHB_CONSERVATION, AnsiUpperCase(Copy(GrilleBulletin.CellValues[10, i], 1, 3)));
          end;
        end;
        if (ObjetACreer = TRUE) or (Copy(Rubrique, 5, 2) = '.R') then
          GrilleBulletin.Objects[0, i] := THB;
        //PT22 : 22/11/01 V562 PH ORACLE, Le champ PHB_ORGANISME doit être renseigné
        if GetValeur(iPHB_ORGANISME) = '' then PutValeur(iPHB_ORGANISME, '....');

        if RechCommentaire(Rubrique) = FALSE then AlimCumulSalarie(THB, Salarie, Nature, Rubrique, Etab, DateDebut, DateFin);
      end;
    end;
  end;
end;
// Fonction de creation de la TOB des Cumuls Salariés

procedure CreationTOBCumSal;
begin
  if TOB_Cumulsal <> nil then
  begin
    TOB_CumulSal.Free;
    TOB_CumulSal := nil;
  end;
  TOB_CumulSal := TOB.Create('Les Cumuls Salariés', nil, -1);
end;
// Fonction de destruction de la  TOB des Cumuls Salariés

procedure DestTOBCumSal;
begin
  if TOB_CumulSal <> nil then TOB_CumulSal.Free;
  TOB_CumulSal := nil;
end;
// Fonction qui indique si les cumuls salariés ont été calculés

function OkCumSal: Boolean;
begin
  result := FALSE;
  if TOB_CumulSal <> nil then result := TRUE;
end;
// Fonction d'affchage d'un montant en fonction du nombre de décimales gérées pour le champ

function DoubleToCell(const X: Double; const DD: integer): string;
begin
  if X = 0 then Result := '' else Result := StrfMontant(X, 15, DD, '', TRUE);
end;

// Fonction de valorisation d'une ligne de rémunération

function ValoriseMt(const cas: WORD; const Cc: string; const NbreDec: Integer; Base, Taux, Coeff: Double): Double;
var
  Mt: Double;
begin
  Mt := 0;
  if cas = 2 then // cas 2
  begin
    if Cc = '04' then Mt := Base * Taux
    else Mt := Base * (Taux / 100);
  end;
  if cas = 3 then // cas 3
  begin
    if Coeff = 0 then
    begin
      result := 0;
      exit;
    end;
    if Cc = '02' then Mt := Base * Taux * Coeff;
    if Cc = '03' then Mt := Base * (Taux / 100) * Coeff;
    if Cc = '06' then Mt := Base * (Taux / Coeff);
    if Cc = '07' then Mt := Base * (Taux / 100 / Coeff);
  end;
  // cas 4 et saisie du coeff Cc vaut tjrs 08
  if cas = 4 then Mt := Base * Coeff;
  Mt := ARRONDI(Mt, NbreDec);
  result := Mt;
end;
// Calcul de la rémunération

function EvalueRem(Tob_Rub: TOB; const Rub: string; var Base, Taux, Coeff, Montant: Double; var lib: string; DateDeb, DateFin: TDateTime; ActionCalcul: TActionBulletin; Ligne:
  Integer; Diag: TObject = nil): Boolean; // Calcul de la rémunération
var
  T1, TRecup: TOB;
  St, Cc, TypeChamp, Champ: string;
  NbreDec, M1, M2, M3, M4, M5, M6 : Integer; // PT174
  cas: WORD;
  ARecupererRem: Boolean;
  TobADetruire: Boolean;
  BaseRaz, Decalage: Boolean;
  MD, MF, MS, Mois, Annee: string;
  Mini, Maxi: Double; // Minimum et Maximum de la rubrique de rémunération
{$IFNDEF  EABSENCES}
  Q: TQuery;
{$ENDIF}
  lelibelle: string;
begin
  result := FALSE;
  TobADetruire := FALSE;
  Base := 0;
  Taux := 0;
  Coeff := 0;
  Montant := 0;
  lib := '';
// DEB PT165
  if (Copy(Rub, 5, 2) = '.R') then
  begin
    Champ := Copy(Rub, 1, 4);
    TRecup := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Rub, 'AAA'], TRUE);
    if Assigned(TRecup) then
    begin
      lib := TRecup.GetValeur(iPHB_LIBELLE);
      Base := TRecup.GetValeur(iPHB_BASEREM);
      Taux := TRecup.GetValeur(iPHB_TAUXREM);
      Coeff := TRecup.GetValeur(iPHB_COEFFREM);
      Montant := TRecup.GetValeur(iPHB_MTREM);
      T1 := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', Rub);
      if Assigned(T1) and (iPRM_LIBELLE = 0) then MemorisePrm(T1);
    end;
    exit;
  end;
// FIN PT165

  T1 := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
  T1 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE); // $$$$
{$ENDIF}
  if Assigned(T1) and (iPRM_LIBELLE = 0) then MemorisePrm(T1);

  if T1 = nil then
  begin
    St := 'Le salarié ' + CodSal + ' La Rémunération ' + Rub + ' n''existe pas, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  TRecup := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Rub, 'AAA'], TRUE);
  if (TRecup = nil) and (ActionCalcul <> taCreation) then
  begin
    St := 'Le salarié ' + CodSal + ' La Rémunération ' + Rub + ' n''est pas dans le bulletin, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  result := TRUE;
  // doit correspondre au cas creation d'une ligne, si on doit recupèrer les valeurs déjà saisies ????
  if TRecup = nil then // alors creation d'une tob vide qui existe
  begin
    TobADetruire := TRUE;
    TRecup := TOB.Create('HISTOBULLETIN', nil, -1);
    TRecup.PutValue('PHB_RUBRIQUE', Rub);
    TRecup.PutValue('PHB_NATURERUB', 'AAA');
    if (TypeTraitement <> 'SAISIE') or (GrilleBull = nil) then
    begin
      TRecup.PutValue('PHB_BASEREM', 0);
      TRecup.PutValue('PHB_TAUXREM', 0);
      TRecup.PutValue('PHB_COEFFREM', 0);
      TRecup.PutValue('PHB_MTREM', 0);
      lib := T1.GetValue('PRM_LIBELLE');
    end
    else // cas creation recupere les cellules saisies pour avoir la saisie des elements variables
    begin
      if Ligne <> 0 then
      begin
        TRecup.PutValue('PHB_BASEREM', Valeur(GrilleBull.Cells[2, Ligne]));
        TRecup.PutValue('PHB_TAUXREM', Valeur(GrilleBull.Cells[3, Ligne]));
        TRecup.PutValue('PHB_COEFFREM', Valeur(GrilleBull.Cells[4, Ligne]));
        TRecup.PutValue('PHB_MTREM', Valeur(GrilleBull.Cells[5, Ligne]));
      end
      else
      begin
        TRecup.PutValue('PHB_BASEREM', 0);
        TRecup.PutValue('PHB_TAUXREM', 0);
        TRecup.PutValue('PHB_COEFFREM', 0);
        TRecup.PutValue('PHB_MTREM', 0);
      end;
      lib := T1.GetValue('PRM_LIBELLE');
    end
  end
  else
  begin
    lib := TRecup.GetValue('PHB_LIBELLE');
    if lib = '' then lib := T1.GetValue('PRM_LIBELLE');
  end;
  RendMoisAnnee(DateFin, Mois, Annee);
  Decalage := VH_Paie.PGDecalage; // récupération des paramètres généraux de la paie
  MD := T1.GetValue('PRM_DU'); // recherche des bornes de validité de la rémunération
  MF := T1.GetValue('PRM_AU');
  MS := '';
  BaseRaz := FALSE;
  // Cas où validité Saisie dans une fourchette
  if (MD <> '') and (MD <> '00') then
  begin
    if Decalage = TRUE then
    begin
      MS := '12'; // 1er mois en decalage concerne le mois de décembre
      // PT7 : 03/09/2001 V547 PH correction validité du au
      if ((StrToInt(Mois) < StrToInt(MD)) or (StrToInt(Mois) > StrToInt(MF))) and (StrToInt(Mois) <> StrToInt(MS)) then BaseRaz := TRUE;
    end;
    if (StrToInt(Mois) < StrToInt(MD)) or (StrToInt(Mois) > StrToInt(MF)) then BaseRaz := TRUE;
    // FIN PT7
  end
  else
  begin
    // Cas où la validité est définie par des mois de validite
    // PT4 : 29/08/2001 V547 PH correction bornes de validité remuneration sur 4 mois
    BaseRaz := TRUE;
    M1 := 0;
    M2 := 0;
    M3 := 0;
    M4 := 0;
    M5 := 0; // PT174
    M6 := 0; // PT174
    if T1.GetValue('PRM_MOIS1') <> '' then M1 := StrToInt(T1.GetValue('PRM_MOIS1'));
    if T1.GetValue('PRM_MOIS2') <> '' then M2 := StrToInt(T1.GetValue('PRM_MOIS2'));
    if T1.GetValue('PRM_MOIS3') <> '' then M3 := StrToInt(T1.GetValue('PRM_MOIS3'));
    if T1.GetValue('PRM_MOIS4') <> '' then M4 := StrToInt(T1.GetValue('PRM_MOIS4'));
    if T1.GetValue('PRM_MOIS5') <> '' then M5 := StrToInt(T1.GetValue('PRM_MOIS5')); // PT174
    if T1.GetValue('PRM_MOIS6') <> '' then M6 := StrToInt(T1.GetValue('PRM_MOIS6')); // PT174

    if (StrToInt(Mois) = M1) or
      (StrToInt(Mois) = M2) or
      (StrToInt(Mois) = M3) or
      (StrToInt(Mois) = M4) or // PT174
      (StrToInt(Mois) = M5) or // PT174
      (StrToInt(Mois) = M6)
      then BaseRaz := FALSE;
    if (M1 = 0) and (M2 = 0) and (M3 = 0) and (M4 = 0) and (M5 = 0) and (M6 = 0)  then BaseRaz := FALSE; // PT174
    // FIN PT4
  end;
  if BaseRaz = TRUE then exit; // PT156
  Cc := T1.GetValue('PRM_CODECALCUL');
//Debut PT194
//  if Diag <> nil then
//  begin
//    Diag.Items.add('');
//    Diag.Items.add('Début calcul de la rémunération ' + Rub + ' ' + T1.getvalue('PRM_LIBELLE') + ' de type de calcul ' + RechDom('PGCODECALCUL', cc, FALSE));
//  end;
  LogMessage(Diag, '');
  LogMessage(Diag, 'Début calcul de la rémunération ' + Rub + ' ' + T1.getvalue('PRM_LIBELLE') + ' de type de calcul ' + RechDom('PGCODECALCUL', cc, FALSE), 'REM', Rub);
// Fin PT194

  NbreDec := T1.GetValue('PRM_DECBASE');
  TypeChamp := '';
  if BaseRaz = FALSE then // Rubrique A ne pas Calculer car Mois ne correspond pas aux dates de validité
  begin
    TypeChamp := T1.GetValue('PRM_TYPEBASE');
    ARecupererRem := RecupRem(ActionCalcul, TypeChamp);
    if ARecupererRem = TRUE then Base := TRecup.GetValue('PHB_BASEREM')
    else Champ := T1.GetValue('PRM_BASEREM');
    if (TypeChamp <> '') and (ARecupererRem = FALSE) then Base := EvalueChampRem(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la base');
    Base := ARRONDI(Base, NbreDec);
  end;
  TypeChamp := '';
  TypeChamp := T1.GetValue('PRM_TYPETAUX');
  ARecupererRem := RecupRem(ActionCalcul, TypeChamp);
  if ARecupererRem = TRUE then Taux := TRecup.GetValue('PHB_TAUXREM')
  else Champ := T1.GetValue('PRM_TAUXREM');
  if (TypeChamp <> '') and (ARecupererRem = FALSE) then Taux := EvalueChampRem(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le taux');
  NbreDec := T1.GetValue('PRM_DECTAUX');
  Taux := ARRONDI(Taux, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValue('PRM_TYPECOEFF');
  ARecupererRem := RecupRem(ActionCalcul, TypeChamp);
  if ARecupererRem = TRUE then Coeff := TRecup.GetValue('PHB_COEFFREM')
  else Champ := T1.GetValue('PRM_COEFFREM');
  if (TypeChamp <> '') and (ARecupererRem = FALSE) then Coeff := EvalueChampRem(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le coefficient');
  NbreDec := T1.GetValue('PRM_DECCOEFF');
  Coeff := ARRONDI(Coeff, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValue('PRM_TYPEMONTANT');
  ARecupererRem := RecupRem(ActionCalcul, TypeChamp);
  if ARecupererRem = TRUE then Montant := TRecup.GetValue('PHB_MTREM')
  else Champ := T1.GetValue('PRM_MONTANT');
  if (TypeChamp <> '') and (ARecupererRem = FALSE) then Montant := EvalueChampRem(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le montant');
  cas := RechCasCodeCalcul(Cc);
  NbreDec := T1.GetValue('PRM_DECMONTANT');
  if Montant = 0 then Montant := ValoriseMt(cas, Cc, NbreDec, Base, Taux, Coeff);
  if TobADetruire = TRUE then TRecup.Free; // uniquement dans le cas où la tob a été crée artificiellement
//Debut PT194
  if Diag <> nil then
  begin
    If Codsal <> '' then
    begin
      st := 'La rémunération ' + T1.GetValue('PRM_LIBELLE') + ' de mode fonctionnement ' + RechDom('PGCODECALCUL', cc, FALSE);
      LogMessage(Diag, st, 'REM', Rub);
  //    Diag.Items.Add(st);
      st := '   a une base de ' + FloatToStr(base) + ' un taux de ' + floatToStr(taux) + ' un coefficient de ' + FloatToStr(coeff) + ' un montant de ' + FloatToStr(Montant);
      LogMessage(Diag, st, 'REM', Rub);
  //    Diag.Items.Add(st);
    end else begin
      st := 'Calcul de la rémunération ' + T1.GetValue('PRM_LIBELLE') + ' de mode fonctionnement ' + RechDom('PGCODECALCUL', cc, FALSE);
      LogMessage(Diag, st, 'REM', Rub);
    end;
  end;
//Fin PT194
  // Gestion du minimum et du maximum
  if T1.GetValue('PRM_TYPEMINI') <> '' then
  begin
    Mini := EvalueChampRem(Tob_Rub, T1.GetValue('PRM_TYPEMINI'), T1.GetValue('PRM_VALEURMINI'), Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'valeur minimale');
    if Montant < Mini then
    begin
      Montant := Mini; // Seuil Minimum
//Debut PT194
//      if Diag <> nil then Diag.Items.add('   le montant a été minimisé par la valeur ' + FloatToStr(mini));
      LogMessage(Diag, '   le montant a été minimisé par la valeur ' + FloatToStr(mini), 'REM', Rub, 'Le montant est minimisé.', (Codsal = ''));
//Fin PT194
    end;
  end;
  if T1.GetValue('PRM_TYPEMAXI') <> '' then
  begin
    Maxi := EvalueChampRem(Tob_Rub, T1.GetValue('PRM_TYPEMAXI'), T1.GetValue('PRM_VALEURMAXI'), Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'valeur maximale');
    if Montant > Maxi then
    begin
      Montant := Maxi; // Seuil Maximum
//Debut PT194
//      if Diag <> nil then Diag.Items.add('   le montant a été maximisé par la valeur ' + FloatToStr(maxi));
      LogMessage(Diag, '   le montant a été maximisé par la valeur ' + FloatToStr(maxi), 'REM', Rub, 'Le montant est maximisé.', (Codsal = ''));
//Fin PT194
    end;
  end;
  // PT8 : 03/09/2001 V547 PH on force dans tous les cas le calcul de l'arrondi sur le montant calculé
  Montant := ARRONDI(Montant, NbreDec);
  // PT4 : 29/08/2001 V547 PH correction bornes de validité remuneration sur 4 mois
  if BaseRaz = TRUE then Montant := 0; // Rubrique A ne pas Calculer car Mois ne correspond pas aux dates de validité
  // FIN PT4
  // DED PT171
{$IFNDEF  EABSENCES}
  if ((ActionBul = taCreation) or (ActionBul = premCreation)) and (T1 <> nil) then
  begin
    if (T1.GetValeur(iPRM_LIBCONTRAT) = 'X') then
    begin
      st := 'SELECT PCI_DEBUTCONTRAT,PCI_FINCONTRAT FROM CONTRATTRAVAIl WHERE PCI_SALARIE="' + Codsal +
        '" AND ((PCI_FINCONTRAT >="' + UsDateTime(DateD) +
        '") AND (PCI_FINCONTRAT <="' + UsDateTime(DateF) + '") AND (PCI_FINCONTRAT IS NOT NULL) AND (PCI_FINCONTRAT > "' + UsDateTime(iDate1900) + '"))' +
        ' AND (PCI_DEBUTCONTRAT <="' + UsDateTime(DateF) + '") ORDER BY PCI_FINCONTRAT ASC';
      Q := OpenSql(St, FALSE);
      while not Q.EOF do
      begin
        lelibelle := 'Contrat du ' + DateToStr(Q.FindField('PCI_DEBUTCONTRAT').AsDateTime) +
          ' au ' + DateToStr(Q.FindField('PCI_FINCONTRAT').AsDateTime);
        IntegreRubCommentaire(tob_rub, nil, Tob_Salarie, 'SRB', Rub, CodSal, Etab, DateD, DateF, lelibelle);
        Q.Next;
      end;
      Ferme(Q);
    end;
  end;
{$ENDIF}
  // FIN PT171
end;
//Fonction d'évaluation d'un champ d'une rémunération en fonction du type du champ

function EvalueChampRem(Tob_Rub: TOB; const TypeChamp, Champ, Rub: string; const DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double;
var
  ii: Integer;
  Trech, Trub: TOB;
  St: string;
begin
  result := 0;
  ii := StrToInt(TypeChamp);
  if (ii < 0) or (ii > 10) then exit;
  Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Champ], TRUE);
  if ((ii > 3) and (ii < 8)) and (Trech = nil) then
  begin
  // DEB PT166
    if VH_Paie.PGIntegAutoRub then
    begin
      IntegreRubBul(TOB_Salarie, Tob_Rub, Champ, 'AAA');
      IntegreAuto := TRUE;
      Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Champ], TRUE);
    end
    else
    begin
      St := St + ' La Rubrique ' + Rub + ' recherche : ' + RechDom('PGTYPECHAMPREM', TypeChamp, FALSE) + ' de la rémunération ' + Champ + ' qui n''existe pas dans le bulletin';
      if TypeTraitement <> 'PREPA' then PGIBox(St, '')
      else TraceErr.Items.Add(St);
      exit;
    end;
  // FIN PT166
  end;
  Trub := Paie_RechercheOptimise(Tob_Rem, 'PRM_RUBRIQUE', rub); // $$$$
{$IFDEF aucasou}
  Trub := Tob_Rem.FindFirst(['PRM_RUBRIQUE'], [rub], FALSE); // $$$$
{$ENDIF}
  case ii of
    0..1: result := 0;
    2: begin
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' a été calculé(e) par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date du ' + DateToStr(DateFin));
        LogMessage(Diag, Quoi + ' est calculé(e) par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date du ' + DateToStr(DateFin), 'ELT', Champ);
//Fin PT194
        result := ValEltNat(Champ, DateFin, LogGetChildLevel(Diag)); //PT172
      end;
    3: begin
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' a été calculé(e) par la variable ' + Champ);
        LogMessage(Diag, Quoi + ' est calculé(e) par la variable ' + Champ, 'VAR', Champ);
//Fin PT194
        result := ValVariable(Champ, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
      end;
    4: begin
        result := Trech.GetValue('PHB_BASEREM');
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' correspond à la base calculée de la rémunération ' + Champ);
        LogMessage(Diag, Quoi + ' correspond à la base calculée de la rémunération ' + Champ, 'REM', Champ);
//Fin PT194
      end;
    5: begin
        result := Trech.GetValue('PHB_TAUXREM');
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' correspond au taux calculé de la rémunération ' + Champ);
        LogMessage(Diag, Quoi + ' correspond au taux calculé de la rémunération ' + Champ, 'REM', Champ);
//Fin PT194
      end;
    6: begin
        result := Trech.GetValue('PHB_COEFFREM');
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' correspond au coefficient calculé de la rémunération ' + Champ);
        LogMessage(Diag, Quoi + ' correspond au coefficient calculé de la rémunération ' + Champ, 'REM', Champ);
//Fin PT194
      end;
    7: begin
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' correspond au montant calculé de la rémunération ' + Champ);
        LogMessage(Diag, Quoi + ' correspond au montant calculé de la rémunération ' + Champ, 'REM', Champ);
//Fin PT194
        result := Trech.GetValue('PHB_MTREM');
      end;
    8:
      begin
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' a comme valeur fixe ' + Champ);
        LogMessage(Diag, Quoi + ' a comme valeur fixe ' + Champ, 'VAL', Champ);
//Fin PT194
        if Trub <> nil then result := Valeur(champ);
      end;
// DEB PT167
    9:
      begin
        result := RendCumulSalSess(Champ);
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' rend le montant de la paie en cours du cumul ' + Champ);
        LogMessage(Diag, Quoi + ' rend le montant de la paie en cours du cumul ' + Champ, 'CUM', Champ);
//Fin PT194
      end;
// FIN PT167
  end;

end;

{ Fonction pour savoir si on recupere le montant déjà calculé ou saisi de la rubrique
cela correspond à un elt permanent ou un elt variable
Dans le cas d'un champ de type valeur ou prend tjrs la valeur saisie dans  la rubrique
}

function RecupRem(const ActionCalcul: TActionBulletin; const TypeChamp: string): Boolean;
var
  ii: Integer;
begin
  result := FALSE;
  if TypeChamp = '' then exit;
  ii := StrToInt(TypeChamp);
  if (ii = 0) or (ii = 1) then result := TRUE;
end;
// Fonction d'évaluation d'une variable

function ValVariable(const VariablePaie: string; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil): double;
var
  T_Rech, T_Var: TOB;
  St, St3, St2, TypeVar: string;
  val1, val2, val3: Double;
  OkDef: Boolean;
  jj: Integer;
  ThenAlert, ElseAlert: Boolean; // Indicateur si une variable de test renvoie une alerte
  RVar: Variant;
begin
  result := 0;
  if VariablePaie = '' then
  begin
    St := 'Salarié : ' + CodSal + ' On ne peut pas calculer une variable qui n''a pas de code';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  ThenAlert := FALSE;
  ElseAlert := FALSE;
  t_rech := Paie_RechercheOptimise(Tob_Variables, 'PVA_VARIABLE', VariablePaie);
{$IFDEF aucasou}
  T_Rech := TOB_Variables.FindFirst(['PVA_VARIABLE'], [VariablePaie], FALSE); // $$$$
{$ENDIF}
  if (T_Rech = nil) then // PT114 Avant, on ne generait pas d'erreur si variable predefinie CEGID  ==> and (StrToInt(VariablePaie) > 200)
  begin
    St := 'Salarié : ' + CodSal + ' La Variable ' + VariablePaie + ' n''existe pas, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  OkDef := RechVarDefinie(VariablePaie, DateDeb, DateFin, RVar, TOB_Rub, Diag);
  if OkDef = TRUE then
  begin
    try
      result := RVar;
  //Debut PT194
  //    if Diag <> nil then
  //      Diag.Items.add('La variable  ' + VariablePaie + ' ' + T_Rech.getvalue('PVA_LIBELLE') + ' pré-calculée CEGID a pour résultat ' + FloatTostr(result));
      LogMessage(Diag, 'La variable ' + VariablePaie + ' ' + T_Rech.getvalue('PVA_LIBELLE') + ' pré-calculée CEGID a pour résultat ' + FloatTostr(result), 'VAR', VariablePaie, 'Calcul de la variable ' + VariablePaie + ' ' + T_Rech.getvalue('PVA_LIBELLE') + ' pré-calculée CEGID', (CodSal = ''));
  //Fin PT194
    except
      result := 0;
      LogMessage(Diag, 'La variable ' + VariablePaie + ' ' + T_Rech.getvalue('PVA_LIBELLE') + ' pré-calculée CEGID est indéfinie.', 'VAR', VariablePaie, 'Calcul de la variable ' + VariablePaie + ' ' + T_Rech.getvalue('PVA_LIBELLE') + ' pré-calculée CEGID', (CodSal = ''));
    end;
    exit;
  end; // Fin Variable CEGID Sortie Fonction
  TypeVar := T_Rech.GetValue('PVA_NATUREVAR');
  if TypeVar = 'VAL' then
  begin
    LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type valeur', 'VAR', VariablePaie, '', False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
    result := EvalueVarVal(T_Rech, DateDeb, DateFin, TOB_Rub);
  end;
  if TypeVar = 'ALE' then
  begin
    if TypeTraitement <> 'PREPA' then PGIBox(T_Rech.GetValue('PVA_MESSAGE1'), T_Rech.GetValue('PVA_MESSAGE2'))
    else
    begin
      // PT176
      TraceErr.Items.Add('Salarié ' + TOB_Salarie.GetValeur(iPSA_SALARIE) + ' ' + TOB_Salarie.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Salarie.GetValeur(iPSA_PRENOM) + ' : ' + T_Rech.GetValue('PVA_MESSAGE1'));
      TraceErr.Items.Add('Salarié ' + TOB_Salarie.GetValeur(iPSA_SALARIE) + ' ' + TOB_Salarie.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Salarie.GetValeur(iPSA_PRENOM) + ' : ' + T_Rech.GetValue('PVA_MESSAGE2'));
    end;
  end;
  if TypeVar = 'CUM' then
  begin
    LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type cumul', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
    result := EvalueVarCum(T_Rech, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag), VariablePaie);
  end;
  if TypeVar = 'CUP' then //PT184
  begin
    LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type cumul de présence', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
    result := EvalueVarCumulPresence(T_Rech, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag), VariablePaie);
  end;
  if (TypeVar = 'COT') or (TypeVar = 'REM') then
  begin
    if TypeVar = 'COT' then
      LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type cotisation', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False))
    else if TypeVar = 'REM' then
      LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type remunération', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
    result := EvalueVarRub(TypeVar, T_Rech, DateDeb, DateFin, TOB_Rub, '', '',Diag);
  end;
  if (TypeVar = 'CAL') then
  begin
//Debut PT194
//    if Diag <> nil then Diag.Items.add('Calcul de la variable ' + VariablePaie + ' de type calcul');
    LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type calcul', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194
    st := EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPEBASE', 'PVA_BASE', 'PVA_OPERATMATH', 9, Diag, VariablePaie);
    EvalueChaineDiv0(St);

    // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
    jj := ObjTM3VM.AddExpr('ExpPaie', St);
    result := ObjTM3VM.GetExprValue(jj);
    ObjTM3VM.DeleteExpr(jj);
//Debut PT194
//    if Diag <> nil then
//      Diag.Items.add('La variable  ' + VariablePaie + '  de type calcul a comme calcul ' + st + ' et pour résultat ' + FloatTostr(result));
    LogMessage(Diag, 'La variable  ' + VariablePaie + '  de type calcul a comme calcul ' + st + ' et pour résultat ' + FloatTostr(result), 'VAR', VariablePaie, 'La variable  ' + VariablePaie + '  de type calcul a comme calcul ' + st, (CodSal =  ''), -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194
  end;
  // PT98   : 19/03/2004 PH V_50 FQ 11200 Prise en compte de nbre de décimales dans le calcul des variables
  if (TypeVar = 'COT') or (TypeVar = 'REM') or (TypeVar = 'CUM') or (TypeVar = 'CAL') or (TypeVar = 'CUP') then //PT184
  begin
    if (T_Rech.GetValue('PVA_MTARRONDI') >= 0) and (T_Rech.GetValue('PVA_MTARRONDI') < 7) then
      result := ARRONDI(result, T_Rech.GetValue('PVA_MTARRONDI'));
  end;
  // FIN PT98
  if (TypeVar = 'TES') then
  begin
    // ObjTM3VM:=TM3VM.Create ; // creation objet VM Script
//Debut PT194
//    if Diag <> nil then Diag.Items.add('Calcul de la variable ' + VariablePaie + ' de type test');
    LogMessage(Diag, 'Calcul de la variable ' + VariablePaie + ' de type test', 'VAR', VariablePaie, '',False, -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194
    st := EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPEBASE', 'PVA_BASE', 'PVA_OPERATMATH', 6, Diag, VariablePaie);
    EvalueChaineDiv0(St);

    // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
    jj := ObjTM3VM.AddExpr('ExpPaie', St);
    val1 := ObjTM3VM.GetExprValue(jj);
    ObjTM3VM.DeleteExpr(jj);
//Debut PT194
//    if Diag <> nil then
//      Diag.Items.add('La variable  ' + VariablePaie + '  de type test a comme condition SI ' + st + '  et pour résultat ' + FloatTostr(val1));
    LogMessage(Diag, 'La variable  ' + VariablePaie + '  de type test a comme condition SI ' + st + '  et pour résultat ' + FloatTostr(val1), 'VAR', VariablePaie, 'La variable  ' + VariablePaie + '  de type test a comme condition SI ' + st, (CodSal =  ''), -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194

    if (T_Rech.GetValue('PVA_TYPERESTHEN0') = '03') and (T_Rech.GetValue('PVA_RESTHEN0') > 200) then
    begin // cas d'une variable il ne faut pas chercher à executer le calcul de la variable si c'est une variable d'alerte
      t_var := Paie_RechercheOptimise(Tob_Variables, 'PVA_VARIABLE', T_Rech.GetValue('PVA_RESTHEN0'));
{$IFDEF aucasou}
      T_Var := TOB_Variables.FindFirst(['PVA_VARIABLE'], [T_Rech.GetValue('PVA_RESTHEN0')], FALSE); // $$$$
{$ENDIF}
      if (T_Var <> nil) then
        if T_Var.GetValue('PVA_NATUREVAR') = 'ALE' then ThenAlert := TRUE;
    end;
    val2 := 0;
    if not ThenAlert then
    begin
      if (T_Rech.GetValue('PVA_OPERATRESTHEN0') <> '') and (T_Rech.GetValue('PVA_OPERATRESTHEN0') <> '') then
        st2 := EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPERESTHEN', 'PVA_RESTHEN', 'PVA_OPERATRESTHEN', 3, Diag, VariablePaie);
      // @@@@@ 2 <-- 3  else st2:=STRFPOINT (EvalueUnChampVar (T_Rech.GetValue ('PVA_TYPERESTHEN0'),T_Rech.GetValue ('PVA_RESTHEN0'), TOB_Rub, T_Rech, 0, DateFin));
      EvalueChaineDiv0(St2);

      // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
      jj := ObjTM3VM.AddExpr('ExpPaie', St2);
      val2 := ObjTM3VM.GetExprValue(jj);
      ObjTM3VM.DeleteExpr(jj);
//Debut PT194
//      if Diag <> nil then
//        Diag.Items.add('La variable  ' + VariablePaie + '  de type test a comme condition ALORS ' + st2 + '  et pour résultat ' + FloatTostr(val2));
    LogMessage(Diag, 'La variable  ' + VariablePaie + '  de type test a comme condition ALORS ' + st2 + '  et pour résultat ' + FloatTostr(val2), 'VAR', VariablePaie, 'La variable  ' + VariablePaie + '  de type test a comme condition ALORS ' + st2, (Codsal = ''), -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194

    end;
    if (T_Rech.GetValue('PVA_TYPERESELSE0') = '03') and (T_Rech.GetValue('PVA_RESELSE0') > 200) then
    begin // cas d'une variable, il ne faut pas chercher à executer le calcul de la variable si c'est une variable d'alerte
      t_var := Paie_RechercheOptimise(Tob_Variables, 'PVA_VARIABLE', T_Rech.GetValue('PVA_RESELSE0'));
{$IFDEF aucasou}
      T_Var := TOB_Variables.FindFirst(['PVA_VARIABLE'], [T_Rech.GetValue('PVA_RESELSE0')], FALSE); // $$$$
{$ENDIF}
      if (T_Var <> nil) then
        if T_Var.GetValue('PVA_NATUREVAR') = 'ALE' then ElseAlert := TRUE;
    end;
    Val3 := 0;
    if not ElseAlert then
    begin
      if (T_Rech.GetValue('PVA_OPERATRESELSE0') <> '') and (T_Rech.GetValue('PVA_OPERATRESELSE0') <> '') then
        st3 := EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPERESELSE', 'PVA_RESELSE', 'PVA_OPERATRESELSE', 3, Diag, VariablePaie) else
        st := STRFPOINT(EvalueUnChampVar(T_Rech.GetValue('PVA_TYPERESELSE0'), T_Rech.GetValue('PVA_RESELSE0'), TOB_Rub, T_Rech, 0, DateFin));
      EvalueChaineDiv0(St3);

      // PROCESS-SERVEUR {$IFNDEF EAGLSERVER} {$ENDIF}
      jj := ObjTM3VM.AddExpr('ExpPaie', St3);
      val3 := ObjTM3VM.GetExprValue(jj);
      ObjTM3VM.DeleteExpr(jj);
//Debut PT194
//      if Diag <> nil then
//        Diag.Items.add('La variable  ' + VariablePaie + '  de type test a comme condition SINON ' + st3 + ' et pour résultat ' + FloatTostr(val3));
    LogMessage(Diag, 'La variable  ' + VariablePaie + '  de type test a comme condition SINON ' + st3 + ' et pour résultat ' + FloatTostr(val3), 'VAR', VariablePaie, 'La variable  ' + VariablePaie + '  de type test a comme condition SINON ' + st3, (Codsal = ''), -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194
    end;
    if Val1 = 1 then
    begin // dans tous les cas si on a une variable d'alerte on excute la variable d'alerte
      result := val2;
      if ThenAlert then EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPERESTHEN', 'PVA_RESTHEN', 'PVA_OPERATRESTHEN', 3);
    end
    else
    begin
      result := val3;
      if ElseAlert then EvalueExpVar(TOB_Rub, T_Rech, DateDeb, DateFin, 'PVA_TYPERESELSE', 'PVA_RESELSE', 'PVA_OPERATRESELSE', 3);
    end;

    if VariablePaie = '0304' then
    begin
      st := VariablePaie;
    end;
    // ObjTM3VM.free;
//Debut PT194
//    if Diag <> nil then
//      Diag.Items.add('La variable  ' + VariablePaie + '  a pour résultat final ' + FloatTostr(result));
    LogMessage(Diag, 'La variable  ' + VariablePaie + '  a pour résultat final ' + FloatTostr(result), 'VAR', VariablePaie, '', (Codsal = ''), -1, RechDom('PGNATUREVARGBL', TypeVar, False));
//Fin PT194
  end;

end;
// Recherche du cas code calcul

function RechCasCodeCalcul(const Cc: string): WORD;
var
  i: Integer;
begin
  i := StrToInt(Cc);
  case i of
    1: result := 1;
    4..5: result := 2;
    8: result := 4;
    2..3: result := 3;
    6..7: result := 3;
  else
    result := 0;
  end;
end;

// Fonction qui indique si on c'est une ligne de commentaire associée à la rubrique

function RechCommentaire(const Rub: string): Boolean;
begin
  result := ((Copy(Rub, 1, 1) = '.') or (Copy(Rub, 5, 1) = '.')) and (not ((Copy(Rub, 6, 1) = 'R')));
end;

// Fonction qui rend le code de la rubrique  en supprimant .x pour les lignes de commentaire

function RendCodeRub(const Rub: string): string;
begin
  if not RechCommentaire(Rub) then result := Rub else Result := Copy(Rub, 1, 4);
end;

// Fonction alimentation enreg de PAIEENCOURS avec les champs salarié et la saisie

procedure AlimChampEntet(const Salarie, Etab: string; const DateDebut, DateFin: TDateTime; Tob_Rub: TOB);
begin
  if (TOB_Salarie = nil) then RecupTobSalarie(Salarie, DateDebut, DateFin) // PT173
  else
  begin
    if TOB_Salarie.GetValeur(iPSA_SALARIE) <> Salarie then RecupTobSalarie(Salarie, DateDebut, DateFin); // PT173 // Reprend le contenu de la table Salarié au moment où l'on écrit Entete Bulletin
  end;

  with TOB_Rub do
  begin
    PutValue('PPU_SALARIE', Salarie);
    Codsal := Salarie;
    PutValue('PPU_ETABLISSEMENT', Etab); // PT151 TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT
    PutValue('PPU_DATEDEBUT', DateDebut);
    PutValue('PPU_DATEFIN', DateFin);
    PutValue('PPU_TRAVAILN1', TOB_Salarie.GetValeur(iPSA_TRAVAILN1));
    PutValue('PPU_TRAVAILN2', TOB_Salarie.GetValeur(iPSA_TRAVAILN2));
    PutValue('PPU_TRAVAILN3', TOB_Salarie.GetValeur(iPSA_TRAVAILN3));
    PutValue('PPU_TRAVAILN4', TOB_Salarie.GetValeur(iPSA_TRAVAILN4));
    PutValue('PPU_CODESTAT', TOB_Salarie.GetValeur(iPSA_CODESTAT));
    PutValue('PPU_CONFIDENTIEL', TOB_Salarie.GetValeur(iPSA_CONFIDENTIEL));
    PutValue('PPU_LIBELLE', TOB_Salarie.GetValeur(iPSA_LIBELLE));
    PutValue('PPU_PRENOM', TOB_Salarie.GetValeur(iPSA_PRENOM));
    PutValue('PPU_NUMEROSS', TOB_Salarie.GetValeur(iPSA_NUMEROSS));
    PutValue('PPU_ADRESSE1', TOB_Salarie.GetValeur(iPSA_ADRESSE1));
    PutValue('PPU_ADRESSE2', TOB_Salarie.GetValeur(iPSA_ADRESSE2));
    PutValue('PPU_ADRESSE3', TOB_Salarie.GetValeur(iPSA_ADRESSE3));
    PutValue('PPU_CODEPOSTAL', TOB_Salarie.GetValeur(iPSA_CODEPOSTAL));
    PutValue('PPU_VILLE', TOB_Salarie.GetValeur(iPSA_VILLE));
    PutValue('PPU_INDICE', TOB_Salarie.GetValeur(iPSA_INDICE));
    PutValue('PPU_NIVEAU', TOB_Salarie.GetValeur(iPSA_NIVEAU));
    PutValue('PPU_CONVENTION', TOB_Salarie.GetValeur(iPSA_CONVENTION));
    PutValue('PPU_CODEEMPLOI', TOB_Salarie.GetValeur(iPSA_CODEEMPLOI));
    PutValue('PPU_AUXILIAIRE', TOB_Salarie.GetValeur(iPSA_AUXILIAIRE));
    PutValue('PPU_LIBREPCMB1', TOB_Salarie.GetValeur(iPSA_LIBREPCMB1));
    PutValue('PPU_LIBREPCMB2', TOB_Salarie.GetValeur(iPSA_LIBREPCMB2));
    PutValue('PPU_LIBREPCMB3', TOB_Salarie.GetValeur(iPSA_LIBREPCMB3));
    PutValue('PPU_LIBREPCMB4', TOB_Salarie.GetValeur(iPSA_LIBREPCMB4));
    PutValue('PPU_DATEANCIENNETE', TOB_Salarie.GetValeur(iPSA_DATEANCIENNETE));
    // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
    PutValue('PPU_BULCOMPL', BullCompl);
    PutValue('PPU_PROFILPART', ProfilSpec);
    PutValue('PPU_BULLCONTRAT', BullContrat);
    PutValue('PPU_UNITEPRISEFF', TOB_Salarie.GetValeur(iPSA_UNITEPRISEFF));
    if    (TOB_Salarie.GetValeur(iPSA_DATESORTIE) > TOB_Salarie.GetValeur(iPSA_DATEENTREE) )
      and (DateFin >= TOB_Salarie.GetValeur(iPSA_DATESORTIE))
      and (BullCompl <> 'X')
      then begin
        PutValue('PPU_BULLSOLDE', 'X');
        SoldeDeToutCompte := True;  //PT195
      end else begin
        PutValue('PPU_BULLSOLDE', '-'); // @@@@ Initialisation si bulletin de solde
        SoldeDeToutCompte := False; //PT195
      end;
    with ChpEntete do
    begin
      PutValue('PPU_PGMODEREGLE', Reglt); {PT3}
      PutValue('PPU_PAYELE', DatePai);
      PutValue('PPU_VALIDELE', DateVal);
      PutValue('PPU_QUALIFICATION', TOB_Salarie.GetValeur(iPSA_QUALIFICATION));
      PutValue('PPU_COEFFICIENT', TOB_Salarie.GetValeur(iPSA_COEFFICIENT));
      PutValue('PPU_LIBELLEEMPLOI', TOB_Salarie.GetValeur(iPSA_LIBELLEEMPLOI));
      PutValue('PPU_DATEENTREE', TOB_Salarie.GetValeur(iPSA_DATEENTREE));
      PutValue('PPU_DATESORTIE', TOB_Salarie.GetValeur(iPSA_DATESORTIE));
      // PT34 : 22/03/2002 PH V571 Traitement de la civilité dans la table Paieencours
      PutValue('PPU_CIVILITE', TOB_Salarie.GetValeur(iPSA_CIVILITE));
      // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
      PutValue('PPU_EDTDEBUT', Edtdu);
      PutValue('PPU_EDTFIN', Edtau);
      PutValue('PPU_BULCOMPL', BullCompl);

      if HorMod then PutValue('PPU_HORAIREMOD', 'X')
      else PutValue('PPU_HORAIREMOD', '-');


      PutValue('PPU_BASESMOD', PaieExpression(BasesMod, 'X', '-'));

      (*
                     if BasesMod then PutValue('PPU_BASESMOD', 'X')
                     else PutValue('PPU_BASESMOD', '-');
      *)

      if TranchesMod then PutValue('PPU_TRANCHESMOD', 'X')
      else PutValue('PPU_TRANCHESMOD', '-');

      if TrentMod then PutValue('PPU_TRENTIEMEMOD', 'X')
      else PutValue('PPU_TRENTIEMEMOD', '-');

      // PT34 : 22/03/2002 PH V571  Civilté et mode reglement modifié
      if RegltMod then PutValue('PPU_REGLTMOD', 'X')
      else PutValue('PPU_REGLTMOD', '-');
      { DEB PT97 }
      if CpAcquisMod then PutValue('PPU_CPACQUISMOD', 'X')
      else PutValue('PPU_CPACQUISMOD', '-');
      { FIN PT97 }
      PutValue('PPU_DENOMINTRENT', DTrent);
      PutValue('PPU_NUMERATTRENT', NTrent);
      PutValue('PPU_JOURSOUVRES', Ouvres);
      PutValue('PPU_JOURSOUVRABLE', Ouvrables);
      PutValue('PPU_HEURESREELLES', HeuresTrav);
    end;
  end;
end;
// Fonction qui calcule automatiquement le numéro de la ligne de commentaire

function RendRubrCommentaire(const rub: string; const CurG: THGrid): Integer;
var
  Num, i: Integer;
  St1: string;
begin
  Num := 1;
  // PT27 : 26/12/2001 V571 PH Dédoublement des lignes multiples de commentaire
  for i := 1 to CurG.RowCount - 1 do
  begin
    St1 := CurG.Cells[0, i];
    if Copy(rub, 1, 4) = Copy(St1, 1, 4) then
    begin
      if (Copy(st1, 5, 1) = '.') then
      begin
        //PT170      Copy(st1, 6, 1)  ->  Copy(st1, 6, length(st1)-6)  pour la gestion des commentaires avec num sur 2 chiffres
        if IsNumeric(Copy(st1, 6, length(st1) - 5)) then
        begin
          //PT170 si Num = 0, c'est que la limite a été dépassée, donc on arrete la recherche
          if (StrToInt(Copy(st1, 6, length(st1) - 5)) >= Num) and (Num <> 0) then
          begin
            Num := StrToInt(Copy(st1, 6, length(st1) - 5));
            Num := Num + 1;
          end;
          //PT170     if Num > 9 then Num := 0;
          if Num > 9 then Num := 0;
        end
        else Continue;
      end;
    end;
  end;
  result := Num;
end;

{ Ensemble de fonctions pour le calcul des rubriques de cotisation
 Calcul de la cotisation}

function EvalueCot(Tob_Rub: TOB; const Rub: string; var Base, TxSal, TxPat, MtSal, MtPat: Double; var lib: string; DateDeb, DateFin: TDateTime; ActionCalcul: TActionBulletin; var
  At: Boolean; Diag: TObject = nil): Boolean; // Calcul de la rémunération
var
  T1, TRecup: TOB;
  St, TypeChamp, Champ: string;
  NbreDec, Mois1, Mois2, Mois3, Mois4: Integer;
  BaseRaz, ARecupererCot, Decalage: Boolean;
  Mini, Maxi: Double;
  MD, MF, MS, Mois, Annee: string;
  UneDate: TDateTime;
begin
  if (Copy(Rub, 5, 2) = '.R') then
  begin
    Champ := Copy(Rub, 1, 4);
    TRecup := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Rub, 'COT'], TRUE);
    if Assigned(TRecup) then
    begin
      lib := TRecup.GetValeur(iPHB_LIBELLE);
      Base := TRecup.GetValeur(iPHB_BASECOT);
      TxSal := TRecup.GetValeur(iPHB_TAUXSALARIAL);
      TxPat := TRecup.GetValeur(iPHB_TAUXPATRONAL);
      MtSal := TRecup.GetValeur(iPHB_MTSALARIAL);
      MtPat := TRecup.GetValeur(iPHB_MTPATRONAL);
      T1 := Paie_RechercheOptimise(TOB_Cotisations, 'PCT_RUBRIQUE', Champ);
      if Assigned(T1) then
      begin
        if (iPCT_ORDREAT = 0) then MemorisePct(T1);
        if T1.GetValeur(iPCT_ORDREAT) = 'X' then At := TRUE else At := FALSE;
      end;
    end;
    exit;
  end;
  Base := 0;
  TxSal := 0;
  TxPat := 0;
  MtPat := 0;
  MtSal := 0;
  lib := '';
  T1 := Paie_RechercheOptimise(TOB_Cotisations, 'PCT_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
  T1 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'COT'], FALSE); // $$$$
{$ENDIF}
  if T1 = nil then
  begin
    St := 'Salarié : ' + CodSal + ' La Cotisation ' + Rub + ' n''existe pas, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  if Assigned(T1) and (iPCT_ORDREAT = 0) then MemorisePct(T1);
  if T1.GetValeur(iPCT_ORDREAT) = 'X' then At := TRUE else At := FALSE;
  TRecup := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Rub, 'COT'], TRUE);
  if (TRecup = nil) and (ActionCalcul <> taCreation) then
  begin
    St := 'Salarié : ' + CodSal + ' La Cotisation ' + Rub + ' n''est pas dans le bulletin, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  if Assigned(TRecup) and (iPHB_LIBELLE = 0) then MemorisePhb(TRecup);
  if TRecup = nil then lib := T1.GetValeur(iPCT_LIBELLE)
  else
  begin
    lib := TRecup.GetValeur(iPHB_LIBELLE);
    if lib = '' then lib := T1.GetValeur(iPCT_LIBELLE);
  end;
  // PT77  : 04/06/03 PH V_421 FQ 10689 Salarié sort dans la période mais la cotisation ne se calcule que si le salarié est present le dernier jour du mois
  if T1.GetValeur(iPCT_PRESFINMOIS) = 'X' then
  begin
    if (TOB_Salarie.GetValeur(iPSA_DATESORTIE) <> NULL) then
      UneDate := TOB_Salarie.GetValeur(iPSA_DATESORTIE)
    else UneDate := iDate1900; // PT94   : 18/12/2003 PH V_50 FQ 11028
    if (UneDate < FINDEMOIS(DateFin)) and (UneDate > IDate1900) then exit;
    // FIN PT94
  end;
  // FIN PT77
  Decalage := VH_Paie.PGDecalage; // récupération des paramètres généraux de la paie
  MD := T1.GetValeur(iPCT_DU); // recherche des bornes de validité de la cotisation
  MF := T1.GetValeur(iPCT_AU);
  MS := '';
  RendMoisAnnee(DateFin, Mois, Annee);
  BaseRaz := FALSE;
  // Cas où validité Saisie dans une fourchette
  if (MD <> '') and (MD <> '00') then
  begin
    if MF = '' then MF := MD; // PT93 initialisation valeur par defaut FQ 11003
    if Decalage = TRUE then
    begin
      MS := '12'; // 1er mois en decalage concerne le mois de décembre
      if ((StrToInt(Mois) < StrToInt(MD)) or (StrToInt(Mois) > StrToInt(MF))) and (StrToInt(Mois) <> StrToInt(MS)) then BaseRaz := TRUE;
    end;
    if (StrToInt(Mois) < StrToInt(MD)) or (StrToInt(Mois) > StrToInt(MF)) then BaseRaz := TRUE;
  end
  else
  begin
    // Cas où la validité est définie par des mois de validite
    //PT5 : 29/08/2001 V547 PH correction bornes de validité cotisaation sur 4 mois
    Mois1 := 0;
    Mois2 := 0;
    Mois3 := 0;
    Mois4 := 0;

    if T1.GetValeur(iPCT_MOIS1) <> '' then Mois1 := StrToInt(T1.GetValeur(iPCT_MOIS1));
    if T1.GetValeur(iPCT_MOIS2) <> '' then Mois2 := StrToInt(T1.GetValeur(iPCT_MOIS2));
    if T1.GetValeur(iPCT_MOIS3) <> '' then Mois3 := StrToInt(T1.GetValeur(iPCT_MOIS3));
    if T1.GetValeur(iPCT_MOIS4) <> '' then Mois4 := StrToInt(T1.GetValeur(iPCT_MOIS4));
    BaseRaz := TRUE;
    if (StrToInt(Mois) = Mois1) or
      (StrToInt(Mois) = Mois2) or
      (StrToInt(Mois) = Mois3) or
      (StrToInt(Mois) = Mois4)
      then BaseRaz := FALSE;
    if (Mois1 = 0) and (Mois2 = 0) and (Mois3 = 0) and (Mois4 = 0) then BaseRaz := FALSE;

    // FIN PT5
  end;
  // @@@@@ if (ActionCalcul = PremCreation) then lib:=T1.GetValue ('PCT_LIBELLE');
  NbreDec := T1.GetValeur(iPCT_DECBASE);
  TypeChamp := '';
//  ARecupererCot := FALSE;

//Debut PT194
//  if Diag <> nil then
//  begin
//    Diag.Items.add('');
//    Diag.Items.add('Début calcul de la cotisation ' + Rub + ' ' + T1.getvalue('PCT_LIBELLE'));
//  end;
    LogMessage(Diag, '');
    LogMessage(Diag, 'Début calcul de la cotisation ' + Rub + ' ' + T1.getvalue('PCT_LIBELLE'), 'COT', Rub);
//Fin PT194

  if BaseRaz = FALSE then // Rubrique A ne pas Calculer car Mois ne correspond pas aux dates de validité
  begin
    TypeChamp := T1.GetValeur(iPCT_TYPEBASE);
    ARecupererCot := RecupCot(ActionCalcul, TypeChamp);
    if ARecupererCot = TRUE then Base := TRecup.GetValeur(iPHB_BASECOT)
    else Champ := T1.GetValeur(iPCT_BASECOTISATION);
    if (TypeChamp <> '') and (ARecupererCot = FALSE) then Base := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la base ');
    Base := ARRONDI(Base, NbreDec);
  end;
  TypeChamp := '';
  // PT40 : 04/04/2002 PH V571 Test si colonne cotisation saisissable
  TypeChamp := T1.GetValeur(iPCT_TYPETAUXSAL);
  ARecupererCot := RecupCot(ActionCalcul, TypeChamp);
  if ARecupererCot = TRUE then TxSal := TRecup.GetValeur(iPHB_TAUXSALARIAL)
  else Champ := T1.GetValeur(iPCT_TAUXSAL);
  if (TypeChamp <> '') and (ARecupererCot = FALSE) then TxSal := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le taux salarial ');
  NbreDec := T1.GetValeur(iPCT_DECTXSAL);
  TxSal := ARRONDI(TxSal, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPETAUXPAT);
  // PT40 : 04/04/2002 PH V571 Test si colonne cotisation saisissable
  ARecupererCot := RecupCot(ActionCalcul, TypeChamp);
  if ARecupererCot = TRUE then TxPat := TRecup.GetValeur(iPHB_TAUXPATRONAL)
  else Champ := T1.GetValeur(iPCT_TAUXPAT);
  if (TypeChamp <> '') and (ARecupererCot = FALSE) then TxPat := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le taux patronal ');
  NbreDec := T1.GetValeur(iPCT_DECTXPAT);
  TxPat := ARRONDI(TxPat, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEFFSAL);
  // PT40 : 04/04/2002 PH V571 Test si colonne cotisation saisissable
  ARecupererCot := RecupCot(ActionCalcul, TypeChamp);
  if ARecupererCot = TRUE then MtSal := TRecup.GetValeur(iPHB_MTSALARIAL)
  else Champ := T1.GetValeur(iPCT_FFSAL);
  if (TypeChamp <> '') and (ARecupererCot = FALSE) then MtSal := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le montant salarial')
  else if TypeChamp <> 'ELV' then MtSal := Base * (TxSal / 100);
  NbreDec := T1.GetValeur(iPCT_DECMTSAL);
  MtSal := ARRONDI(MtSal, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEFFPAT);
  // PT40 : 04/04/2002 PH V571 Test si colonne cotisation saisissable
  ARecupererCot := RecupCot(ActionCalcul, TypeChamp);
  if ARecupererCot = TRUE then MtPat := TRecup.GetValeur(iPHB_MTPATRONAL)
  else Champ := T1.GetValeur(iPCT_FFPAT);
  if (TypeChamp <> '') and (ARecupererCot = FALSE) then MtPat := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'Le montant patronal')
  else if TypeChamp <> 'ELV' then MtPat := Base * (TxPat / 100);
  NbreDec := T1.GetValeur(iPCT_DECMTPAT);
  MtPat := ARRONDI(MtPat, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEMINISAL); // Recup du mimimum sur Mt salarial
  if (TypeChamp <> '') then
  begin
    Champ := T1.GetValeur(iPCT_VALEURMINISAL);
    Mini := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la valeur minimale montant salarial');
    Mini := ARRONDI(Mini, T1.GetValeur(iPCT_DECMTSAL));
    if MtSal < Mini then
    begin
      MtSal := Mini;
//Debut PT194
//      if Diag <> nil then Diag.Items.add('  le montant salarial a été minimisé par la valeur ' + FloatToStr(Mini));
      LogMessage(Diag, '  le montant salarial a été minimisé par la valeur ' + FloatToStr(Mini), 'COT', Rub, 'le montant salarial a été minimisé.', (Codsal = ''));
//Fin PT194 
    end;
  end;
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEMAXISAL); // Recup du maximum sur Mt salarial
  if (TypeChamp <> '') then
  begin
    Champ := T1.GetValeur(iPCT_VALEURMAXISAL);
    Maxi := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la valeur maximale montant salarial');
    Maxi := ARRONDI(Maxi, T1.GetValeur(iPCT_DECMTSAL));
    if MtSal > Maxi then
    begin
      MtSal := Maxi;
//Debut PT194
//      if Diag <> nil then Diag.Items.add('  le montant salarial a été maximisé par la valeur ' + FloatToStr(Maxi));
      LogMessage(Diag, '  le montant salarial a été maximisé par la valeur ' + FloatToStr(Maxi), 'COT', Rub, 'le montant salarial a été maximisé.', (Codsal = ''));
//Fin PT194
    end;
  end;
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEMINIPAT); // Recup du mimimum sur Mt patronal
  if (TypeChamp <> '') then
  begin
    Champ := T1.GetValeur(iPCT_VALEURMINIPAT);
    Mini := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la valeur minimale montant patronal');
    Mini := ARRONDI(Mini, T1.GetValeur(iPCT_DECMTPAT));
    if MtPat < Mini then
    begin
//Debut PT194
//      if Diag <> nil then Diag.Items.add('  le montant patronal a été minimisé par la valeur ' + FloatToStr(Mini));
      LogMessage(Diag, '  le montant patronal a été minimisé par la valeur ' + FloatToStr(Mini), 'COT', Rub, 'le montant salarial a été minimisé.', (Codsal = ''));
//Fin PT194
      MtPat := Mini;
    end;
  end;
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEMAXIPAT); // Recup du maximum sur Mt patronal
  if (TypeChamp <> '') then
  begin
    Champ := T1.GetValeur(iPCT_VALEURMAXIPAT);
    Maxi := EvalueChampCot(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la valeur maximale montant patronal');
    Maxi := ARRONDI(Maxi, T1.GetValeur(iPCT_DECMTPAT));
    if MtPat > Maxi then
    begin
      MtPat := Maxi;
//Debut PT194
//      if Diag <> nil then Diag.Items.add('  le montant salarial a été maxmisé par la valeur ' + FloatToStr(Maxi));
      LogMessage(Diag, '  le montant salarial a été maximisé par la valeur ' + FloatToStr(Maxi), 'COT', Rub, 'le montant salarial a été maximisé.', (Codsal = ''));
//Fin PT194
    end;
  end;
  if BaseRaz = TRUE then // Rubrique A ne pas Calculer car Mois ne correspond pas aux dates de validité
  begin
    MtPat := 0;
    MtSal := 0;
  end;
end;
{ Fonction d'évaluation d'un champ d'une rémunération en fonction du type du champ
Rajout du type de champ VAL pour indiquer une valeur à prendre
}

function EvalueChampCot(Tob_Rub: TOB; const TypeChamp, Champ, Rub: string; const DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double;
var
  T_Rech, T_Rub, T1: TOB;
  St, LaNature: string;
begin
  result := 0;
  if Champ = '' then exit; // tous les champs ne sont pas obligatoires
  if typeChamp = 'ZCU' then
  begin
    result := RendCumulSalSess(Champ);
//Debut PT194
//    if Diag <> nil then
//      Diag.Items.add(Quoi + ' rend le montant de la paie en cours du cumul ' + Champ);
      LogMessage(Diag, Quoi + ' rend le montant de la paie en cours du cumul ' + Champ, 'COT', Rub);
//Fin PT194
    exit;
  end;
  if TypeChamp = 'BDC' then T_Rech := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['BAS', Champ], TRUE)
  else T_Rech := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Champ], TRUE);
  if (T_Rech = nil) and ((TypeChamp = 'REM') or (TypeChamp = 'BDC')) then
  begin
  // DEB PT166
    if VH_Paie.PGIntegAutoRub then
    begin
      if TypeChamp = 'REM' then LaNature := 'AAA'
      else LaNature := 'BAS';
      IntegreRubBul(TOB_Salarie, Tob_Rub, Champ, LaNature);
      IntegreAuto := TRUE;
      T_rech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [Lanature, Champ], TRUE);
    end
    else
    begin
      St := 'Salarié : ' + CodSal + ' La Rubrique de cotisation ' + Rub + ' est mal paramètrée ';
      if TypeTraitement <> 'PREPA' then PGIBox(St, '')
      else TraceErr.Items.Add(St);
      exit;
    end;
  // FIN PT166
  end;
  if TypeChamp = 'ELN' then
  begin
    result := ValEltNat(Champ, DateFin, LogGetChildLevel(Diag)); //PT172
//Debut PT194
//    if Diag <> nil then
//      Diag.Items.add(Quoi + ' a été calculé(e) par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date du ' + DateToStr(Datefin));
      LogMessage(Diag, Quoi + ' a été calculé(e) par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date du ' + DateToStr(Datefin), 'COT', Rub,  Quoi + ' a été calculé(e) par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE), (Codsal = ''));
//Fin PT194
  end
  else
    if TypeChamp = 'VAR' then
    begin
      result := ValVariable(Champ, DateDeb, DateFin, Tob_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then
//        Diag.Items.add(Quoi + ' a été calculé(e) par la variable ' + Champ);
      LogMessage(Diag, Quoi + ' a été calculé(e) par la variable ' + Champ, 'COT', Rub);
//Fin PT194
    end
    else
      if TypeChamp = 'VAL' then
      begin
        t_rub := Paie_RechercheOptimise(TOB_Cotisations, 'PCT_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
        T_rub := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE'], [Rub], FALSE); // $$$$
{$ENDIF}
        if T_Rub <> nil then result := Valeur(Champ);
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add(Quoi + ' correspond à la valeur de ' + Champ + ' de la cotisation ' + Rub);
          LogMessage(Diag, Quoi + ' correspond à la valeur de ' + Champ + ' de la cotisation ' + Rub, 'COT', Rub);
//Fin PT194
      end
      else
      begin
        if Assigned(T_Rech) and (iPHB_TRANCHE1 = 0) then MemorisePhb(T_Rech);
        if TypeChamp = 'REM' then
        begin
          result := T_Rech.GetValeur(iPHB_MTREM); // Montant de Rémunération
//Debut PT194
//          if Diag <> nil then
//            Diag.Items.add(Quoi + ' correspond au montant de la rémunération ' + Champ);
          LogMessage(Diag, Quoi + ' correspond au montant de la rémunération ' + Champ, 'COT', Rub);
//Fin PT194
        end
        else
          if TypeChamp = 'BDC' then
          begin
            T1 := Paie_RechercheOptimise(TOB_Cotisations, 'PCT_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
            T1 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'COT'], FALSE); // $$$$
{$ENDIF}
            if T1 <> nil then
            begin
              if Assigned(T1) and (iPCT_CODETRANCHE = 0) then MemorisePct(T1);
              st := T1.GetValeur(iPCT_CODETRANCHE);
              if st = 'TOT' then
              begin
                result := T_Rech.GetValeur(iPHB_BASECOT); // Base rubrique de Base de  Cotisation
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la base de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la base de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end
              else if st = 'TR1' then
              begin
                result := T_Rech.GetValeur(iPHB_TRANCHE1); // TR1 de Base de  Cotisation
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la tranche1 de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la tranche1 de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end
              else if st = 'TR2' then
              begin
                result := T_Rech.GetValeur(iPHB_TRANCHE2); // TR2 de Base de  Cotisation
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la tranche2 de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la tranche2 de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end
              else if st = 'TR3' then
              begin
                result := T_Rech.GetValeur(iPHB_TRANCHE3); // TR3 de Base de  Cotisation
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la tranche3 de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la tranche3 de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end
              else if st = 'U12' then
              begin
                result := T_Rech.GetValeur(iPHB_TRANCHE1) + T_Rech.GetValeur(iPHB_TRANCHE2); // TR1+T2
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la tranche1+tranche2 de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la tranche1+tranche2 de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end
              else if st = 'U13' then
              begin
                result := T_Rech.GetValeur(iPHB_TRANCHE1) + T_Rech.GetValeur(iPHB_TRANCHE2) + T_Rech.GetValeur(iPHB_TRANCHE3); // TR1+T2+T3
//Debut PT194
//                if Diag <> nil then
//                  Diag.Items.add(Quoi + ' correspond à la tranche1+tranche2+tranche3 de la base de cotisation ' + Champ);
                LogMessage(Diag, Quoi + ' correspond à la tranche1+tranche2+tranche3 de la base de cotisation ' + Champ, 'COT', Rub);
//Fin PT194
              end;
            end;
          end;
      end;
end;

{ Fonction pour savoir si on recupere le montant déjà calculé ou saisi de la rubrique de cotisation
cela correspond à un elt permanent
Le 14/03/2000 suppression notion élément permanent remplacée par valeur
}

function RecupCot(const ActionCalcul: TActionBulletin; const TypeChamp: string): Boolean;
begin
  result := FALSE;
  // PT40 : 04/04/2002 PH V571 Test si colonne cotisation saisissable
  if (ActionCalcul <> taCreation) and ((TypeChamp = 'ELP') or (TypeChamp = 'ELV'))
    then result := TRUE; // ligne sans effet suite à modif
end;
{ Fonction de calcul du bulletin dans son intégralité
Parcours de la TOB contenant la liste des rubriques contenues dans le bulletin
et en fonction de la nature, les lignes sont valoriées
Il convient après suivant le cas d'écrire dans la base, de remplir les grilles de
saisie.
Cette fonction est utilisable pour la préparation automatique car elle prend en compte
les valeurs des informations (exemple : Elt nationaux)) mais aussi le paramètrage des
rubriques au moment où on fait le calcul==> Donc cette fonction peut être appellée
pour le recalcul d'une paie pour tenir compte des mofications apportées au plan de paie.
}

{ Ensemble de fonctions pour le calcul des rubriques de Bases de cotisation
 Calcul de la base, Plafond, T1,T2,T3 avec regul si rubrique soumise à régularisation}

function EvalueBas(Tob_Rub: TOB; const Rub: string; var Base, Plafond, Plf1, Plf2, Plf3, TR1, TR2, TR3: Double; var lib: string; const DateDeb, DateFin: TDateTime; ActionCalcul:
  TActionBulletin; const BaseForcee, TrancheForcee: Boolean; Diag: TObject = nil): Boolean; // Calcul de la rémunération
var
  T1, TRecup: TOB;
  St, TypeChamp, Champ, TPlf1, TPlf2, TPlf3: string;
  NbreDec: Integer;
  ARegulariser, ACalculer: Boolean;
begin
  Base := 0;
  Plf1 := 0;
  Plf2 := 0;
  Plf3 := 0;
  TR1 := 0;
  TR2 := 0;
  TR3 := 0;
  lib := '';
  T1 := Paie_RechercheOptimise(TOB_Bases, 'PCT_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
  T1 := TOB_Bases.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'BAS'], FALSE); // $$$$
{$ENDIF}
  if T1 = nil then
  begin
    St := 'Salarié : ' + CodSal + ' La Base de Cotisation ' + Rub + ' n''existe pas, On ne peut pas la calculer';
    if TypeTraitement <> 'PREPA' then PGIBox(St, '')
    else TraceErr.Items.Add(St);
    exit;
  end;
  TRecup := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Rub, 'BAS'], TRUE);
  if Assigned(T1) and (iPCT_LIBELLE = 0) then MemorisePct(T1);

  if (TRecup = nil) and (ActionCalcul <> taCreation) then
  begin
    St := 'La Base de Cotisation ' + Rub + ' n''est pas dans le bulletin, On ne peut pas la calculer';
    //  ShowMessage (St); On n'a pas de message car une base peut être la somme de 3 bases pouvant etre valorisées ou pas
    exit;
  end;
  if TRecup = nil then lib := T1.GetValeur(iPCT_LIBELLE)
  else
  begin
    lib := TRecup.GetValeur(iPHB_LIBELLE);
    if lib = '' then lib := T1.GetValeur(iPCT_LIBELLE);
  end;
  // @@@@ if (ActionCalcul = PremCreation) then lib:=T1.GetValue ('PCT_LIBELLE');
//Debut PT194
//  if Diag <> nil then
//  begin
//    Diag.Items.add('');
//    Diag.Items.add('Début de calcul de la base de cotisation ' + Rub +' '+lib);
//  end;
  LogMessage(Diag, '');
  LogMessage(Diag, 'Début de calcul de la base de cotisation ' + Rub +' '+lib, 'BAS', Rub);
//Fin PT194
  NbreDec := T1.GetValeur(iPCT_DECBASECOT);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEBASE);
  if T1.GetValeur(iPCT_SOUMISREGUL) = 'X' then ARegulariser := TRUE
  else ARegulariser := FALSE;
  if TrancheForcee = TRUE then ARegulariser := FALSE;
  if BaseForcee = TRUE then ACalculer := FALSE else ACalculer := TRUE;
  if ACalculer = FALSE then
  begin
    Base := TRecup.GetValeur(iPHB_BASECOT);
//Debut PT194
//    if Diag <> nil then DIag.Items.Add('  La base a été forcée donc elle n''est plus calculée');
    LogMessage(Diag, '  La base a été forcée donc elle n''est plus calculée', 'BAS', Rub);
//Fin PT194
  end
  else Champ := T1.GetValeur(iPCT_BASECOTISATION);
  if (TypeChamp <> '') and (ACalculer = TRUE) then Base := EvalueChampBas(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'la base ');
  Base := ARRONDI(Base, NbreDec);
  TypeChamp := '';
  TypeChamp := T1.GetValeur(iPCT_TYPEPLAFOND);
  if ACalculer = FALSE then
  begin
    Plafond := TRecup.GetValeur(iPHB_PLAFOND);
    Plf1 := TRecup.GetValeur(iPHB_PLAFOND1);
    Plf2 := TRecup.GetValeur(iPHB_PLAFOND2);
    Plf3 := TRecup.GetValeur(iPHB_PLAFOND3);
  end
  else Champ := T1.GetValeur(iPCT_PLAFOND);
  // Recup du plafond theorique
  if (TypeChamp <> '') and (ACalculer = TRUE) then Plafond := EvalueChampBas(Tob_Rub, TypeChamp, Champ, Rub, DateDeb, DateFin, LogGetChildLevel(Diag), 'le plafond ');
  Plafond := ARRONDI(Plafond, NbreDec);
  TPlf1 := T1.GetValeur(iPCT_TYPETRANCHE1);
  TPlf2 := T1.GetValeur(iPCT_TYPETRANCHE2);
  TPlf3 := T1.GetValeur(iPCT_TYPETRANCHE3);
  // il convient de calculer le plafond reel sur lequel on va calculer les tranches
  if TrancheForcee = TRUE then // si on saisit les tranches alors le calcul des plafonds est déactivé car on peut aussi les saisir
  begin
    TR1 := TRecup.GetValeur(iPHB_TRANCHE1);
    TR2 := TRecup.GetValeur(iPHB_TRANCHE2);
    TR3 := TRecup.GetValeur(iPHB_TRANCHE3);
    Plf1 := TRecup.GetValeur(iPHB_PLAFOND1);
    Plf2 := TRecup.GetValeur(iPHB_PLAFOND2);
    Plf3 := TRecup.GetValeur(iPHB_PLAFOND3);
//Debut PT194
//    if Diag <> nil then Diag.Items.add('   Les tranches ont été forcées donc elles ne sont plus calculées');
    LogMessage(Diag, '   Les tranches ont été forcées donc elles ne sont plus calculées', 'BAS', Rub);
//Fin PT194
  end
  else // Calcul des regularisations sur les 3 tranches
  begin
    Plf1 := EvaluePlafond(TOB_Rub, T1.GetValeur(iPCT_TYPETRANCHE1), T1.GetValeur(iPCT_TRANCHE1), Rub, Plafond, DateDeb, DateFin, NbreDec, T1.GetValeur(iPCT_TYPEREGUL), LogGetChildLevel(Diag), ' le planfond 1 ');
    Plf2 := EvaluePlafond(TOB_Rub, T1.GetValeur(iPCT_TYPETRANCHE2), T1.GetValeur(iPCT_TRANCHE2), Rub, Plafond, DateDeb, DateFin, NbreDec, T1.GetValeur(iPCT_TYPEREGUL), LogGetChildLevel(Diag), ' le planfond 2 ') - Plf1;
    if Plf2 < 0 then Plf2 := 0;
    Plf3 := EvaluePlafond(TOB_Rub, T1.GetValeur(iPCT_TYPETRANCHE3), T1.GetValeur(iPCT_TRANCHE3), Rub, Plafond, DateDeb, DateFin, NbreDec, T1.GetValeur(iPCT_TYPEREGUL), LogGetChildLevel(Diag), ' le planfond 3 ') - Plf1 -
      Plf2;
    if Plf3 < 0 then Plf3 := 0;
    St := FloatToStr(Base) + '  ' + FloatToStr(Plf1) + '  ' + FloatToStr(Plf2) + '  ' + FloatToStr(Plf3);
//Debut PT194
//    if Diag <> nil then Diag.Items.add('   la base et les plafonds calculés ont pour valeurs : ' + st);
    LogMessage(Diag, '   la base et les plafonds calculés ont pour valeurs : ' + st, 'BAS', Rub, 'Calcul de la base et des plafonds.', (Codsal = ''));
//Fin PT194
    CalculCumulsRegul(Rub, Base, Plf1, Plf2, Plf3, TR1, TR2, TR3, DateDeb, DateFin, ARegulariser, TPlf1, TPlf2, TPlf3);
    St := FloatToStr(TR1) + '  ' + FloatToStr(TR2) + '  ' + FloatToStr(TR3);
//Debut PT194
//    if Diag <> nil then Diag.Items.add('Fin de calcul de la base de cotisation ' + Rub + ' les tranches calculées ont pour valeurs : ' + st);
    LogMessage(Diag, 'Fin de calcul de la base de cotisation ' + Rub + ' les tranches calculées ont pour valeurs : ' + st, 'BAS', Rub, 'Fin de calcul de la base de cotisation ' + Rub, (Codsal = ''));
//Fin PT194
  end;
end;
// Fonction de calcul des regularisations des rubriques de base de cotisation

procedure CalculCumulsRegul(const Rub: string; var Base, Plf1, Plf2, Plf3, TR1, TR2, TR3: Double; DateDeb, DateFin: TDateTime; const ARegulariser: Boolean; const TPlf1, TPlf2,
  TPlf3: string);
var
  BaseCal: TRendRub;
  Cbas, CPlf1, CTr1, CTr2, CTr3, CPlf2, CPlf3: Double; // Cumuls des valeurs à calculer
  T1: TOB;
  //  TST2, TST3: Boolean;
  PlafondRub, TypeRegul: string; // champ plafond de la rubrique de base
begin
  Cbas := 0;
  CPlf1 := 0;
  CTr1 := 0;
  CTr2 := 0;
  CTr3 := 0;
  CPlf2 := 0;
  CPlf3 := 0;
  //  TST2 := FALSE;
  //  TST3 := FALSE;
    // Recherche exercice social en cours de traitement
  T1 := TOB_ExerSocial.Detail[0];
  DateDeb := T1.GetValue('PEX_DATEDEBUT');
  DateFin := DateDeb - 1; // Regul des montants jusqu'à la date (non incluse) de debut de la session
  // BaseCal:=ValRubDate (Rub, 'BAS', DatD, DatF);  // Calcul&Cumul des champs pour reguls
  if ARegulariser = TRUE then
  begin
    BaseCal := ValBase(Rub); // @@@@@
    with BaseCal do
    begin
      Cbas := Base;
      CPlf1 := Plfd1;
      CPlf2 := Plfd2;
      CPlf3 := Plfd3;
      CTr1 := Tr1;
      CTr2 := Tr2;
      CTr3 := Tr3;
    end;
  end;
  //  calcul Plf1 et Plf2 et Plf3 de la session en cours
  T1 := Paie_RechercheOptimise(TOB_Bases, 'PCT_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
  T1 := TOB_Bases.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'BAS'], FALSE); // $$$$
{$ENDIF}
  if T1 <> nil then
  begin
    PlafondRub := T1.GetValue('PCT_PLAFOND');
    TypeRegul := T1.GetValue('PCT_TYPEREGUL '); // True = Annuelle cas de la Taxe sur les salaires
  end;
  if ((Base + Cbas) >= (Plf1 + CPlf1)) then
  begin
    Tr1 := 0;
    //   if (TST2 = FALSE) AND (CTr2 <>0) then Tr1 := -CTR2;
    Tr1 := Tr1 + Plf1 + CPlf1 - CTr1; // Regul tranche 1
    if (Base + CBas) < (Plf1 + CTr1) then Tr1 := Base + CBas - Ctr1;
  end
  else Tr1 := Base + Cbas - CTr1;

  if (TPlf1 = 'DEP') then Tr1 := Base + Cbas - CTr1; // Deplafonnée

  if (Plf2 >= 0) or ((Plf2 = 0) and (TypeRegul = 'X')) then // Regul tranche 2 Cas plafond ou regul annuelle
  begin
    Tr2 := 0;
    //  if (TST3 = FALSE) AND (CTR3 <> 0) then TR2:=-CTr3;
    if (Base + CBas - CTr1 - Tr1) < (Plf2 + CTr2) then // plfd1
    begin
      if (Base + CBas - CTr1 - Tr1) > 0 then Tr2 := Base + CBas - Ctr2 - CTr1 - TR1
      else Tr2 := -CTr2;
      //   if (Base+CBas-CTr1-Tr1) < (CPlf1-CTr1) then Tr2 := -CTr2; // Déficit de tranche 2 > Tranche 2 sans régul donc pas de tranche 2
    end
    else
    begin
      Tr2 := Plf2 + CPlf2 - Ctr2;
      if (Base + CBas - CTr1 - Tr1 - CPlf2 - Plf2) < 0 then Tr2 := Tr2 + (Base + CBas - CTr1 - Tr1 - CPlf2 - Plf2);
    end;
    if (TypeRegul = 'X') and (Tr2 > Base) then // Cas regul annuelle et tranche 2 calculee > Base alors limitation à la base
      if TPlf2 <> 'DEP' then Tr2 := Base - Tr1;
    Tr3 := Base - Tr1 - Tr2;
  end;

  if (Plf2 = 0) and (TypeRegul <> 'X') then // Pas de plafond ou regul mensuelle
  begin
    if TPlf3 = '' then Tr3 := 0; // beurk
    //  if Plf2=0 then Tr2:=0; // regul sur tranche2 pratiquée
    if (TPlf2 = 'DEP') then Tr2 := Base - Tr1; // Deplafonnée
  end;
  if (Plf3 <> 0) or ((Plf3 = 0) and (TypeRegul = 'X')) then // Regul tranche 3    AND ((Plf2+CPlf2) >= (Plf3+CPlf3))
  begin
    Tr3 := CBas + Base - CTr1 - Tr1 - Ctr2 - Tr2 - Ctr3;
    if Tr3 > Plf3 + CPlf3 - CTr3 then Tr3 := Plf3 + CPlf3 - CTr3; // limitation sur la tranche 3
    if (TPlf3 = 'DEP') then Tr3 := Base - Tr1 - Tr2; // Deplafonnée

    // TR4:=Base-Tr1-Tr2-TR3;
  end;
  if ((Plf3 = 0) or ((Plf2 + CPlf2) < (CPlf3 + PLf3))) and (TypeRegul <> 'X') then //
  begin
    //  Tr4:=0;
    if (Plf2 + CPlf2) <= 0 then Tr3 := Base - Tr1 - Tr2;
    //  if (Plf3=0) then Tr3:=0; //  Regul Tranche 3
    if TPlf3 = 'DEP' then Tr3 := Base - Tr1 - Tr2; // Deplafonnée
    if TPlf3 = '' then Tr3 := 0;
  end;
  if PlafondRub = '' then
  begin
    Tr1 := 0;
    Tr2 := 0;
    Tr3 := 0;
  end;

end;
//Fonction d'évaluation d'un champ d'une Base de cotisation en fonction du type du champ

function EvalueChampBas(Tob_Rub: TOB; TypeChamp, Champ, Rub: string; DateDeb, DateFin: TDateTime; Diag: TObject = nil; Quoi: string = ''): Double;
var
  T_Rech: TOB;
  St: string;
begin
  result := 0;
  if Champ = '' then exit; // tous les champs ne sont pas obligatoires
  T_Rech := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Champ], TRUE);
  if (T_Rech = nil) and (TypeChamp = 'REM') then
  begin
  // DEB PT166
    if VH_Paie.PGIntegAutoRub then
    begin
      IntegreRubBul(TOB_Salarie, Tob_Rub, Champ, 'AAA');
      IntegreAuto := TRUE;
      T_rech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Champ], TRUE);
    end
    else
    begin
      St := 'Salarié : ' + CodSal + ' La Rubrique de base de cotisation ' + Rub + ' est mal paramètrée ';
      if TypeTraitement <> 'PREPA' then PGIBox(St, '')
      else TraceErr.Items.Add(St);
      exit;
    end;
  // FIN PT166
  end;
  if TypeChamp = 'ELN' then
  begin
    result := ValEltNat(Champ, DateFin, LogGetChildLevel(Diag)); // Recup elelment national //PT172
//Debut PT194
//    if Diag <> nil then Diag.Items.add(Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date ' + DateToStr(Datefin));
    LogMessage(Diag, Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date ' + DateToStr(Datefin), 'BAS', Rub
                   , Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE), (Codsal = ''));
//Fin PT194
  end
  else
    if TypeChamp = 'VAR' then
    begin
      result := ValVariable(Champ, DateDeb, DateFin, Tob_Rub, LogGetChildLevel(Diag)); // recup calcul d'une variable
//Debut PT194
//      if Diag <> nil then Diag.Items.add(Quoi + ' a été calculé par la variable ' + Champ);
      LogMessage(Diag, Quoi + ' a été calculé par la variable ' + Champ, 'BAS', Rub);
//Fin PT194
    end
    else
      if TypeChamp = 'REM' then
      begin
        result := T_Rech.GetValue('PHB_MTREM'); // Montant de Rémunération
//Debut PT194
//        if Diag <> nil then Diag.Items.add(Quoi + ' a repris le montant de la rémunération ' + Champ);
        LogMessage(Diag, Quoi + ' a repris le montant de la rémunération ' + Champ, 'BAS', Rub);
//Fin PT194
      end;
end;

// Fonction  de calcul des Plafond1 2 3 en fonction du plafond théorique calculé

function EvaluePlafond(TOB_Rub: TOB; TypeChamp, Champ, Rub: string; Plafond: Double; DateDeb, DateFin: TDateTime; NbreDec: Integer; TypeRegul: string; Diag: TObject = nil; Quoi: string = ''): Double;
var
  T_regul: TOB;
begin
  result := 0;
  if TypeRegul = 'X' then // regul de type annuelle utilisée pour la taxe sur les salaires
  begin
    t_regul := Paie_RechercheOptimise(TOB_HistoBasesCot, 'PHB_RUBRIQUE', Rub); // $$$$
{$IFDEF aucasou}
    T_regul := TOB_HistoBasesCot.FindFirst(['PHB_RUBRIQUE'], [Rub], TRUE);
{$ENDIF}
    if T_regul <> nil then exit; // on a trouvé une ligne historique sur la rubrique de base donc les plafonds annuels sont déjà connus
  end;
  if Champ = '' then exit; // tous les champs ne sont pas obligatoires donc par exemple pas de plafond en tranche 3
  if TypeChamp = 'ELN' then
  begin
    result := ValEltNat(Champ, DateFin, LogGetChildLevel(Diag)); // Recup elelment national //PT172
//Debut PT194
//    if Diag <> nil then Diag.Items.add(Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date ' + DateToStr(Datefin));
    LogMessage(Diag, Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE) + ' à la date ' + DateToStr(Datefin), 'PLF', Rub
                   , Quoi + ' a été calculé par l''élément national ' + Champ + ' ' + RechDom('PGELEMENTNAT',Champ,FALSE), (Codsal = ''));
//Fin PT194
  end
  else
    if TypeChamp = 'VAR' then
    begin
      result := ValVariable(Champ, DateDeb, DateFin, Tob_Rub); // recup calcul d'une variable
//Debut PT194
//      if Diag <> nil then Diag.Items.add(Quoi + ' a été calculé par la variable ' + Champ);
      LogMessage(Diag, Quoi + ' a été calculé par la variable ' + Champ, 'PLF', Rub);
//Fin PT194
    end
    else
      if TypeChamp = 'NBR' then
      begin
        result := StrToInt(Champ) * Plafond; // Montant du plafondx est  y  fois le plafond
//Debut PT194
//        if Diag <> nil then Diag.Items.add(Quoi + ' a été calculé par ' + Champ + ' fois le plafond : ' + FloatToStr(Plafond));
        LogMessage(Diag, Quoi + ' a été calculé par ' + Champ + ' fois le plafond : ' + FloatToStr(Plafond), 'PLF', Rub, Quoi + ' a été calculé par ' + Champ + ' fois le plafond', (Codsal = ''));
//Fin PT194
      end;
end;

// Fonction qui calcule le bulletion cad qui parcoure la liste des rubriques du bulletins et qui valorise les lignes

{$IFNDEF EMANAGER}

procedure CalculBulletin(Tob_Rub: TOB; CalcMaint: Boolean = FALSE; Diag: TObject = nil; PaieEnvers: Boolean = FALSE); //PT148 PT161 PT162
var
  i: Integer;
  Base, Taux, Coeff, Montant, TxSal, TxPat, MtSal, MtPat: double;
  Plafond, Plf1, Plf2, Plf3, Tr1, Tr2, Tr3: double;
// PT197  CpAcquis, CpMois, CpSupp, CpAnc, CpNbmois, Cpbase: Double; { PT127-1 }
  CpAcquis, CpMois, CpSupp, CpAnc, CpFract, CpNbmois, Cpbase: Double; { PT127-1 }
  Ligne, TRech, TETABSAL: TOB;
  lib, Nat, Rub, Them, Etab, CodeSal, StMsgErr : string; // PT196
  ZDateDeb, ZDateFin: TDateTime;
  AT, OkCalcul: Boolean; { PT127-1 }
  iChampNatureRub, iChampRubrique, iChampDateDebut, iChampDateFin, iChampEtab, iChampSal: integer;
  iChampBaseRem, iChampTauxRem, iChampCoeffRem, iChampMtRem, iChampLibelle: integer;
  iChampBaseCot, iChampTauxSalarial, iChampTauxPatronal, iChampMtSalarial, iChampMtPatronal: integer;
  iChampPlafond, iChampPlafond1, iChampPlafond2, iChampPlafond3, iCHampTranche1, iCHampTranche2, iCHampTranche3: integer;
  RubExclue: Boolean; // PT148
begin
  DestTOBCumSal; // RAZ des cumuls salariés car il vont etre recalculés
  CreationTOBCumSal;
  // PT92   : 10/12/2003 PH V_50 Recalcul systématique de la saisie arrêt
  if Assigned(TOB_SaisieArret) then FreeAndNil(TOB_SaisieArret);
  OkCalcul := False; { PT127-1 }
  Base := 0;
  Taux := 0;
  Coeff := 0;
  Montant := 0;
  TxSal := 0;
  TxPat := 0;
  MtSal := 0;
  MtPat := 0;
  Plf1 := 0;
  Plf2 := 0;
  Plf3 := 0;
  Tr1 := 0;
  Tr2 := 0;
  Tr3 := 0;
  Plafond := 0;
  lib := '';
  nat := '';
  rub := '';
  TauxAt := 0;
  // tri des rubriques par nature et par numero avant le calcul du bulletin
  Tob_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
  // Initialisation des champs par défaut
  iChampNatureRub := -1;
  iChampRubrique := -1;
  iChampDateDebut := -1;
  iChampDateFin := -1;
  iChampEtab := -1;
  iChampSal := -1;
  iChampBaseRem := -1;
  iChampTauxRem := -1;
  iChampCoeffRem := -1;
  iChampMtRem := -1;
  iChampLibelle := -1;
  iChampBaseCot := -1;
  iChampTauxSalarial := -1;
  iChampTauxPatronal := -1;
  iChampMtSalarial := -1;
  iChampMtPatronal := -1;
  iChampPlafond := -1;
  iChampPlafond1 := -1;
  iChampPlafond2 := -1;
  iChampPlafond3 := -1;
  iCHampTranche1 := -1;
  iCHampTranche2 := -1;
  iCHampTranche3 := -1;
  // optimisation
  if tob_rub.detail.count > 0 then
    with tob_rub.detail[0] do
    begin
      iChampNatureRub := GetNumChamp('PHB_NATURERUB');
      iChampRubrique := GetNumChamp('PHB_RUBRIQUE');
      iChampDateDebut := GetNumChamp('PHB_DATEDEBUT');
      iChampDateFin := GetNumChamp('PHB_DATEFIN');
      iChampEtab := GetNumChamp('PHB_ETABLISSEMENT');
      iChampSal := GetNumChamp('PHB_SALARIE');

      iChampBaseRem := GetNumChamp('PHB_BASEREM');
      iChampTauxRem := GetNumChamp('PHB_TAUXREM');
      iChampCoeffRem := GetNumChamp('PHB_COEFFREM');
      iChampMtRem := GetNumChamp('PHB_MTREM');
      iChampLibelle := GetNumChamp('PHB_LIBELLE');

      iChampBaseCot := GetNumChamp('PHB_BASECOT');
      iChampTauxSalarial := GetNumChamp('PHB_TAUXSALARIAL');
      iChampTauxPatronal := GetNumChamp('PHB_TAUXPATRONAL');
      iChampMtSalarial := GetNumChamp('PHB_MTSALARIAL');
      iChampMtPatronal := GetNumChamp('PHB_MTPATRONAL');

      iChampPlafond := GetNumChamp('PHB_PLAFOND');
      iChampPlafond1 := GetNumChamp('PHB_PLAFOND1');
      iChampPlafond2 := GetNumChamp('PHB_PLAFOND2');
      iChampPlafond3 := GetNumChamp('PHB_PLAFOND3');
      iCHampTranche1 := GetNumChamp('PHB_TRANCHE1');
      iCHampTranche2 := GetNumChamp('PHB_TRANCHE2');
      iCHampTranche3 := GetNumChamp('PHB_TRANCHE3');
    end;

  // 1ere Passe pour traiter toutes les rubriques sauf les remunérations venant après les cotisations
  for i := 0 to Tob_Rub.Detail.Count - 1 do
  begin
    Base := 0;
    Taux := 0;
    Coeff := 0;
    Montant := 0;
    TxSal := 0;
    TxPat := 0;
    MtSal := 0;
    MtPat := 0;
    lib := '';
    nat := '';
    rub := '';
    Ligne := Tob_Rub.Detail[i];
    if assigned(Ligne) then
      with Ligne do
      begin
        Nat := GetValeur(iChampNatureRub);
        Rub := GetValeur(iChampRubrique);
        ZDateDeb := GetValeur(iChampDateDebut);
        ZDateFin := GetValeur(iChampDateFin);
        Etab := GetValeur(iChampEtab);
        CodeSal := GetValeur(iChampSal);
// d PT148
        RubExclue := FALSE;
        trech := Paie_RechercheOptimise(tob_rem, 'PRM_RUBRIQUE', Rub);
        if assigned(tRech) then
        begin
          if ((trech.getValue('PRM_EXCLENVERS') = 'X') and (PaieEnvers = TRUE)) or // PT162
            ((trech.getValue('PRM_EXCLMAINT') = 'X') and (CalcMaint = TRUE)) then
            RubExclue := TRUE;
        end;
// f PT148
        if (Nat = 'AAA') and not RechCommentaire(Rub) then
        begin
          if (not RubExclue) then // PT148
          begin
            // Remunérations mais on ne calcule pas une rubrique de commentaire associée à une rémunération
            trech := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', Copy(Rub, 1, 4)); // PT165
            if assigned(TRech) then
            begin // On recherche le theme de la rubrique
              Them := TRech.GetValue('PRM_THEMEREM');
              { DEB PT127-1 }
              if (Them = 'ABS') and (GblCP) and (OkCalcul = False) and (JCPACQUISTHEORIQUE <> 0) then
              begin
                TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT)], True);
// PT197                CpAcquis := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, ZDateDeb, ZDateFin, True, True, Cpbase, CpNbmois, CpMois, CpSupp, CpAnc);
                CpAcquis := ChargeAcquisParametre(TOB_Salarie, TETABSAL, TOB_RUB, ZDateDeb, ZDateFin, True, True, Cpbase, CpNbmois, CpMois, CpSupp, CpAnc, CpFract);
                if (CpAcquis <> JCPACQUIS) then
                begin
                  OkCalcul := True;
                  JCPACQUIS := CpAcquis;
                  InitTobAcquisCp(TOB_Salarie, TETABSAL, ZDateDeb, ZDateFin, Cpbase, CpNbmois, CpMois, CpSupp, CpAnc, suivant);
                  if X <> nil then
                  begin
                    X.CPMOIS.text := floattostr(CpMois);
                    X.CpSupp.text := floattostr(CpSupp);
                    X.CpAnc.text := floattostr(CpAnc);
                  end;
                  MCPACQUIS := Cpnbmois;
                  BCPACQUIS := Cpbase;
                  Tob_Pris := SalIntegreCP(TOB_Salarie, TOB_Rub, T_MvtAcquis, DateD, DateF, true, StMsgErr); //PT196 Réalimentation de la TOB_Pris car le calcul de bulletin modifie le nbre de jours acquis
                end;
              end;
              { FIN PT127-1 }
              if (Them <> 'RNI') and (Them <> 'RSS') then
              begin
                EvalueRem(Tob_Rub, rub, Base, Taux, Coeff, Montant, lib, ZDateDeb, ZDateFin, taModification, 0, LogGetChildLevel(Diag));
                PutValeur(iChampBaseRem, Base);
                PutValeur(iChampTauxRem, Taux);
                PutValeur(iChampCoeffRem, Coeff);
                PutValeur(iChampMtRem, Montant);
                PutValeur(iChampLibelle, lib);
                AlimCumulSalarie(Ligne, CodeSal, Nat, Rub, Etab, ZDateDeb, ZDateFin);
              end;
            end;
          end; // fin if (not RubExclue) PT148
        end
        // PT26 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
        else if (Nat = 'COT') and (RechCommentaire(Rub) = FALSE) then
        begin
          // Calcul des rubriques de cotisation
          EvalueCot(Tob_Rub, rub, Base, TxSal, TxPat, MtSal, MtPat, lib, ZDateDeb, ZDateFin, taModification, AT, LogGetChildLevel(Diag));

          PutValeur(iChampBaseCot, Base);
          PutValeur(iChampTauxSalarial, TxSal);
          PutValeur(iChampTauxPatronal, TxPat);
          PutValeur(iChampMtSalarial, MtSal);
          PutValeur(iChampMtPatronal, MtPat);
          PutValeur(iChampLibelle, lib);
          if AT and (MtPat <> 0) then TauxAt := TxPat; // recup du taux at sur une rubrique AT valorisée
          AlimCumulSalarie(Ligne, CodeSal, Nat, Rub, Etab, ZDateDeb, ZDateFin);
        end
        else if Nat = 'BAS' then
        begin
          // Calcul des rubriques de bases de cotisation
          EvalueBas(Tob_Rub, rub, Base, Plafond, Plf1, Plf2, Plf3, Tr1, Tr2, Tr3, lib, ZDateDeb, ZDateFin, taModification, ChpEntete.BasesMod, ChpEntete.TranchesMod, LogGetChildLevel(Diag));

          PutValeur(iChampBaseCot, Base);
          PutValeur(iChampPlafond, Plafond);
          PutValeur(iChampPlafond1, Plf1);
          PutValeur(iChampPlafond2, Plf2);
          PutValeur(iChampPlafond3, Plf3);
          PutValeur(iCHampTranche1, Tr1);
          PutValeur(iCHampTranche2, Tr2);
          PutValeur(iCHampTranche3, Tr3);
          PutValeur(iChampLibelle, lib);
          AlimCumulSalarie(Ligne, CodeSal, Nat, Rub, Etab, ZDateDeb, ZDateFin);
        end;
      end; // fin du With
  end; // Fin de la boucle for

  // Deuxieme passe pour traiter uniquement les rubriques de rémunérations venant après les cotisations
  for i := 0 to Tob_Rub.Detail.Count - 1 do
  begin
    Base := 0;
    Taux := 0;
    Coeff := 0;
    Montant := 0;
    TxSal := 0;
    TxPat := 0;
    MtSal := 0;
    MtPat := 0;
    lib := '';
    nat := '';
    rub := '';
    Ligne := Tob_Rub.Detail[i];
    if Ligne <> nil then
      with ligne do
      begin
        Nat := GetValeur(iChampNatureRub);
        Rub := GetValeur(iChampRubrique);
        ZDateDeb := GetValeur(iChampDateDebut);
        ZDateFin := GetValeur(iChampDateFin);
        Etab := GetValeur(iChampEtab);
        CodeSal := GetValeur(iChampSal);
        if (Nat = 'AAA') and (RechCommentaire(Rub) = FALSE) then
        begin
          //   Remunérations mais on ne calcule pas une rubrique de commentaire associée à une rémunération
          trech := Paie_RechercheOptimise(TOB_Rem, 'PRM_RUBRIQUE', Copy(Rub, 1, 4)); // PT200 cas des rémunérations de régul de bas de bulletin
          if assigned(TRech) then
          begin // On recherche le theme de la rubrique
            Them := TRech.GetValue('PRM_THEMEREM');
            if (Them = 'RNI') or (Them = 'RSS') then
            begin
              EvalueRem(Tob_Rub, rub, Base, Taux, Coeff, Montant, lib, ZDateDeb, ZDateFin, taModification, 0, LogGetChildLevel(Diag));
              PutValeur(iChampBaseRem, Base);
              PutValeur(iChampTauxRem, Taux);
              PutValeur(iChampCoeffRem, Coeff);
              PutValeur(iChampMtRem, Montant);
              PutValeur(iChampLibelle, lib);
              AlimCumulSalarie(Ligne, CodeSal, Nat, Rub, Etab, ZDateDeb, ZDateFin);
            end;
          end;
        end
        else
        begin // cas d'une cotisation pour la ducs on recupere le TAUX AT
          // PT26 : 19/12/2001 V571 PH Traitement ligne de commentaire pour les cotisations
          if (Nat = 'COT') and (TauxAT <> 0) and (RechCommentaire(Rub) = FALSE) then Ligne.PutValue('PHB_TAUXAT', TauxAT);
        end; // fin traitement rem ou cot
      end; // Ligne <> nil
  end; // Fin de la boucle for

  Tob_Rub.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE');
  // recalcul de la valorisation absence et maintient
end;
{$ENDIF EMANAGER}

procedure IntegreMvtSolde(Tob_Rub, TOB_Salarie: tob; const Etab, Salarie: string; const DateD, DateF: TdateTime);
var
  P5T_Etab: Tob;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  p5t_etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', Etab); // $$$$
{$IFDEF aucasou}
  P5T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
{$ENDIF}

  if JCPSOLDE = 0 // JCPSOLDE écrit dans Integrepaye
    then
  begin
    BCPACQUIS := Arrondi(RendCumulSalSess('11'), DCP);
    if not CalculValoCP(Tob_solde, Tob_Acquis, P5T_Etab, TOB_Salarie, ChpEntete.Ouvres, ChpEntete.Ouvrables, DateD, DateF, suivant, 'SLD', salarie) //PT-9 ajout Tob_salarie
      then HShowMessage('5;Calcul bulletin :;1 ou plusieurs mouvements de solde n''ont pu être payés par manque d''acquis;E;O;O;O;;;', '', '');
    if Assigned(Tob_Solde) then { PT160-1 }
      if (Tob_Solde.detail.count > 0) then //notVide(Tob_Solde, true) then PT160
        IntegrePaye(Tob_Rub, Tob_Solde, Etab, Salarie, DateD, DateF, 'SLD');
  end
  else
  begin
    //     if BCPACQUIS = Arrondi(RendCumulSalSess ('11'),DCP) then exit;
    if ActionBulCP = tamodification then exit; //mvi
    if AcquisMaj then
      recalculeValoSolde(DateD, DateF, P5T_Etab, TOB_Salarie, Tob_solde, Tob_rub, Salarie) //PT-9 ajout Tob_salarie
    else
      if not CalculValoCP(Tob_solde, Tob_Acquis, P5T_Etab, TOB_Salarie, ChpEntete.Ouvres, ChpEntete.Ouvrables, DateD, DateF, suivant, 'SLD', salarie) //PT-9 ajout Tob_salarie
        then HShowMessage('5;Calcul bulletin :;1 ou plusieurs mouvements de solde n''ont pu être payés par manque d''acquis;E;O;O;O;;;', '', '');
    ;

  end;
end;
// si modif cumul 11 , recalcul de la valeur des cp solde obligatoire

procedure recalculeValoSolde(const DateD, Datef: tdatetime; P5Etab, TSal, Tob_solde, Tob_rub: tob; const Salarie: string);
var
  DTclotureP, DTFinP, DTDebutP: tdatetime;
  periode: integer;
  NbJoursRef, Valox: double;
  tp: tob;
  ModeValo: string;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  ModeValo := TrouveModeValorisation(P5etab, TSal); //PT59
  if TSal.GetValeur(iPSA_CPACQUISMOIS) = 'ETB' then //DEB PT-9
    NbJoursRef := JoursReference(P5Etab.getvalue('ETB_NBREACQUISCP'))
  else
    NbJoursRef := JoursReference(TSal.GetValeur(iPSA_NBREACQUISCP)); //FIN PT-9
  DTclotureP := P5Etab.getvalue('ETB_DATECLOTURECPN');
  periode := RendPeriode(DTFinP, DTDebutP, DTclotureP, DateF);

  CalculValoX(DTclotureP, DateF, DateD, periode, NbJoursRef, 'SLD', Tob_Acquis, nil);
  if NbJoursRef * 10 * (VALOXP0D + MCPACQUIS) = 0 then
    Valox := 0 else
    //     Valox := arrondi((ValoXP0N+BCPACQUIS)*12/(NbJoursRef*10*(VALOXP0D+MCPACQUIS)),DCP);
    Valox := CalculValoX(DTclotureP, DateF, DateD, periode, NbJoursRef, 'SLD', Tob_Acquis, nil);
  // plutôt avant de générer un solde
  {re juste pour voir 08-02     RetireAcquisCourant(Tob_Acquis);
       AjouteAcquiscourant(Tob_Acquis,Tob_salarie,DateF,suivant);
        if Tob_pris <> nil then
        begin
        TP := Tob_pris.findfirst([''],[''],True);
        while tp <> nil do
          begin
          if TP.getvalue('PCN_MVTDUPLIQUE') <> 'X' then
             PrisBul    := PrisBul + TP.getvalue('PCN_JOURS');
          TP := Tob_pris.findnext([''],[''],True);
          end;
        end;

       GenereSolde(Tob_Salarie,P5Etab,Tob_Pris,Tob_Acquis, Tob_Solde,DateF,DateD,suivant,JCPACQUIS,PrisBul,'M');}
  {     RetireAcquisCourant(Tob_Acquis);
       AjouteAcquiscourant(Tob_Acquis,Tob_salarie,DateF,suivant);}
  tp := nil;
  if tob_solde <> nil then
    tp := Tob_solde.findfirst(['PCN_PERIODEPY'], ['0'], true);
  while Tp <> nil do
  begin
    if tp.getvalue('PCN_CODERGRPT') <> -1 then
    begin
      TP.putValue('PCN_VALOX', Valox * TP.getvalue('PCN_JOURS'));
      //         ModeValo     := P5Etab.getvalue('ETB_VALORINDEMCP'); // ?????
      if modeValo = 'D' then
        TP.putvalue('PCN_VALORETENUE', TP.getValue('PCN_VALOX'))
      else if ModeValo <> 'M' then
        if TP.getValue('PCN_VALOX') > TP.getValue('PCN_VALOMS') then
          TP.putvalue('PCN_VALORETENUE', TP.getValue('PCN_VALOX'))
        else TP.putvalue('PCN_VALORETENUE', TP.getValue('PCN_VALOMS'));

      AgregeCumulPris(Tob_Solde);
      IntegrePaye(Tob_Rub, Tob_Solde, P5Etab.getvalue('ETB_ETABLISSEMENT'), Salarie, DateD, DateF, 'SLD');
      exit;
    end;
    tp := Tob_solde.findnext(['PCN_PERIODEPY'], ['0'], true);
  end;
end;

{ Fonction qui calcule la date de Paiement de la paie en fonction
de la fiche Salarié. A savoir si la date n'est pas personnalisée, il faut reprendre
les informations au niveau de l'etablissement
}

function CalculDatePaie(TOB_Sal: TOB; DatePaie: TDateTime): TDateTime;
var
  Zdate, TypeDatPaie, DatPaie, Jour: string;
  T_etab: TOB; // tob etablissement
  MM, JJ, AA: WORD;
begin
  //PT68    19/12/2002 PH V591 Affectation de la date de paiement à la date de fin de paie si non renseignée
  result := DatePaie;
  Zdate := '';
  TypeDatPaie := '';
  DatPaie := '';
  Jour := '';
  TypeDatPaie := string(TOB_Sal.GetValeur(iPSA_TYPDATPAIEMENT));
  T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
  T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)], TRUE);
{$ENDIF}
  // Type Date de paiement: Idem Etab ou Personnalisé
  if BullCompl <> 'X' then
  begin
    if (TypeDatPaie = 'ETB') or (TypeDatPaie = '') then
    begin
      if T_Etab <> nil then DatPaie := T_Etab.GetValue('ETB_MOISPAIEMENT');
    end
    else DatPaie := TOB_Sal.GetValeur(iPSA_MOISPAIEMENT);
    if DatPaie = 'FDM' then
    begin
      result := FindeMois(DatePaie);
      exit;
    end; // Fin de mois
    if (DatPaie = 'MEC') or (DatPaie = 'MSU') then
    begin
      if DatPaie = 'MSU' then DatePaie := PLUSMOIS(DatePaie, 1);
      Zdate := DateToStr(DatePaie);
      if (TypeDatPaie = 'ETB') or (TypeDatPaie = '') then
      begin
        if T_Etab <> nil then Jour := IntToStr(T_Etab.GetValue('ETB_JOURPAIEMENT'));
      end
      else Jour := IntToStr(TOB_Sal.GetValeur(iPSA_JOURPAIEMENT));
      ZDate := Copy(Jour, 1, 2) + Copy(ZDate, 3, 8);
      if HEnt1.IsValidDate(ZDate) then result := StrToDate(ZDate) // Si date valide alors ok
      else result := FindeMois(DatePaie); // sinon on donne le dernier jour du mois
    end;
  end
    //PT19 : 16/11/01 V562 PH Modif du calcul de la date de paiement dans le cas bulletin complémentaire
  else
  begin
    if T_Etab <> nil then DatPaie := T_Etab.GetValue('ETB_BCMOISPAIEMENT')
    else DatPaie := 'FDM'; // si pas renseigné alors fin de mois
    if DatPaie = 'FDM' then
    begin
      result := FindeMois(DatePaie);
      exit;
    end; // Fin de mois
    if (DatPaie = 'MEC') or (DatPaie = 'MSU') then
    begin
      if DatPaie = 'MSU' then DatePaie := PLUSMOIS(DatePaie, 1);
      Zdate := DateToStr(DatePaie);
      if T_Etab <> nil then Jour := IntToStr(T_Etab.GetValue('ETB_BCJOURPAIEMENT'))
      else Jour := IntToStr(TOB_Sal.GetValeur(iPSA_JOURPAIEMENT)); // quoi prendre ?
      if (Jour = '0') and (BullCompl = 'X') then
      begin
        DecodeDate(DatePaie, AA, MM, JJ);
        Jour := IntToStr(JJ);
      end;
      ZDate := Copy(Jour, 1, 2) + Copy(ZDate, 3, 8);
      if HEnt1.IsValidDate(ZDate) then result := StrToDate(ZDate) // Si date valide alors ok
      else result := FindeMois(DatePaie); // sinon on donne le dernier jour du mois
    end;
  end;
end;

{ Fonction qui donne le mode de paiement de la paie en fonction
de la fiche Salarié. A savoir si le mode de paiement n'est pas personnalisée, il faut reprendre
les informations au niveau de l'etablissement
PT25 : 04/12/2001 V563 PH Mode de réglement pour les bulletins complémentaires
}

function RendModeRegle(TOB_Sal: TOB): string;
var
  TypMode, ModeRegle: string;
  T_etab: TOB; // tob etablissement
begin
  TypMode := '';
  ModeRegle := '';
  TypMode := string(TOB_Sal.GetValeur(iPSA_TYPREGLT));
  T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
  T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Sal.GetValeur(iPSA_ETABLISSEMENT)], TRUE);
{$ENDIF}
  if BullCompl = 'X' then
  begin
    if T_Etab <> nil then result := T_Etab.GetValue('ETB_BCMODEREGLE');
    exit;
  end;
  // Type Date de paiement: Idem Etab ou Personnalisé
  if (TypMode = 'ETB') or (TypMode = '') then
  begin
    if T_Etab <> nil then ModeRegle := T_Etab.GetValue('ETB_PGMODEREGLE'); {PT3}
  end
  else ModeRegle := TOB_Sal.GetValeur(iPSA_PGMODEREGLE); {PT3}
  result := ModeRegle;
end;
{ Fonction d'alimentation des rubriques en fonction des profils liés uniquement à la session de paie
cas pour traiter la maladie, le chomage ...
}

procedure AlimProfilMulti(ListeProfil: string; Salarie, TPE: TOB);
var
  Profil: string;
begin
  Profil := READTOKENST(ListeProfil);
  while Profil <> '' do
  begin
    ChargeProfil(Salarie, TPE, Profil);
    Profil := READTOKENST(ListeProfil);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 01/08/2007
Modifié le ... :   /  /
Description .. : Fonction de calcul des variables de présence prédéfinies
Suite ........ : Code réservé de P001 à P200
Mots clefs ... :
*****************************************************************}
function RechVarPresenceDefinie(const VariablePaie: string; const DateDebut, DateFin: TDateTime; var MtCalcul: Variant; Diag: TObject = nil): Boolean; //PT172
var
  Code : Integer;
  Profil, Cycle, Modele, CodeSal, JourneeType, TypeIdem  : string;
  isExceptionSalarie, IsExceptionCycle : Boolean;
  TempTob : Tob;
  DureePause, ValeurIdem : Extended;
  AnneeM, MoisM, jourM : WORD;
//  DebutHeuresNuit, FinHeuresNuit : TDateTime;
  DateATester : TDateTime;
//  sModele, sJournee, sCycleOuModele, sTypeDeCycle : String;
//  bIsExceptionSalarie, bIsExceptionCycle : Boolean;
//  dDateLancement, DateDebutCycle, DateFinCycle : TDateTime;
//  iResteJours : Integer;
begin
  result := False;
  Code := ValeurI(VariablePaie);
  CodeSal := TOB_Salarie.GetValeur(iPSA_SALARIE);
  if ((Code = 1) or (Code = 3) or (Code = 5) or (Code = 7) or (code = 13) or (Code = 15)) then
    { Recherche de la journée type SANS LES EXCEPTIONS }
    ObjCalcuVarPre.GetInfoPresenceSalarie(DateFin, CodeSal, Profil, Cycle, Modele, JourneeType, isExceptionSalarie, IsExceptionCycle, False)
  else
    { Recherche de la journée type AVEC LES EXCEPTIONS }
    ObjCalcuVarPre.GetInfoPresenceSalarie(DateFin, CodeSal, Profil, Cycle, Modele, JourneeType, isExceptionSalarie, IsExceptionCycle, True);
  case Code of
    1..16 : begin
      TempTob := ObjCalcuVarPre.TobJourneesTypes.FindFirst(['PJO_JOURNEETYPE'],[JourneeType], False);
      case Code of
        1..2 : begin  {durée théorique sur la journée}
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := 0;
            MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE1') + TempTob.GetDouble('PJO_DUREEPLAGE2');
        //    MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE1') + TempTob.GetDouble('PJO_DUREEPLAGE2') - TempTob.GetDouble('PJO_DUREENOTRAV');
        //    if TempTob.GetString('PJO_PAUSEEFFECTIF') = '-'  then
        //      MtCalcul := MtCalcul - TempTob.GetDouble('PJO_DUREEPAUSE');
          end;
        end;
        3..4 : begin  { Panier jour théorique }
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := ObjCalcuVarPre.EvalVarPreJType('', 'PAJ', DateDebut, DateFin, 506, LogGetChildLevel(Diag), JourneeType);
          end;;
        end;
        5..6 : begin  { Panier nuit théorique }
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := ObjCalcuVarPre.EvalVarPreJType('', 'PAN', DateDebut, DateFin, 506, LogGetChildLevel(Diag), JourneeType);
          end;;
        end;
        7..8 : begin  { Heures de nuit théorique }
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := ObjCalcuVarPre.EvalVarPreJType('', 'HNU', DateDebut, DateFin, 506, LogGetChildLevel(Diag), JourneeType);
//           //@@GGU : Récupération du paramétrage des heures de nuit (début et fin)
//            DebutHeuresNuit := EncodeTime(21,0,0,0);
//            FinHeuresNuit := EncodeTime(6,0,0,0);
//            result := True;
//            { Calcul du nombre d'heures contenues dans les heures de nuit : }
//            MtCalcul := 0;
//            if TempTob.GetString('PJO_JOURJ1PLAGE1') = '-' then
//            begin
//              { pas de changement de journée pendant la plage 1 }
//              if TempTob.GetDouble('PJO_HORFINPLAGE1') > DebutHeuresNuit then
//                { la plage 1 à des heures de nuit }
//                MtCalcul := TempTob.GetDouble('PJO_HORFINPLAGE1') - Max(DebutHeuresNuit, TempTob.GetDouble('PJO_HORDEBPLAGE1'));
//            end else
//              { Changement de journée pendant la plage 1 }
//              MtCalcul :=  Min(FinHeuresNuit, TempTob.GetDouble('PJO_HORFINPLAGE1'))
//                           + 1 - Max(DebutHeuresNuit, TempTob.GetDouble('PJO_HORDEBPLAGE1'));
//            if TempTob.GetString('PJO_JOURJ1PLAGE2') = '-' then
//            begin
//              if TempTob.GetDouble('PJO_HORFINPLAGE2') > DebutHeuresNuit then
//                MtCalcul := MtCalcul + TempTob.GetDouble('PJO_HORFINPLAGE2') - Max(DebutHeuresNuit, TempTob.GetDouble('PJO_HORDEBPLAGE2'));
//            end else begin
//              if TempTob.GetString('PJO_JOURJ1PLAGE1') = '-' then
//              begin
//                if TempTob.GetDouble('PJO_HORDEBPLAGE2') < TempTob.GetDouble('PJO_HORFINPLAGE2') then
//                begin
//                  { Changement de journée entre la plage 1 et 2 }
//                  if FinHeuresNuit > TempTob.GetDouble('PJO_HORDEBPLAGE2') then
//                    MtCalcul := MtCalcul + Min(TempTob.GetDouble('PJO_HORFINPLAGE2'), FinHeuresNuit) - TempTob.GetDouble('PJO_HORDEBPLAGE2');
//                end else begin
//                  { Changement de journée pendant la plage 2 }
//                  MtCalcul :=  MtCalcul + Min(FinHeuresNuit, TempTob.GetDouble('PJO_HORFINPLAGE2'))
//                             + 1 - Max(DebutHeuresNuit, TempTob.GetDouble('PJO_HORDEBPLAGE2'));
//                end;
//              end else begin
//                { Changement de journée pendant la plage 1 }
//                if FinHeuresNuit > TempTob.GetDouble('PJO_HORDEBPLAGE2') then
//                  MtCalcul := MtCalcul + Min(TempTob.GetDouble('PJO_HORFINPLAGE2'), FinHeuresNuit) - TempTob.GetDouble('PJO_HORDEBPLAGE2');
//              end;
//            end;
//            MtCalcul := MtCalcul * 24;
          end;
        end;
        9  : begin  { Nombre d'heures sur la journée }
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := 0;
            if TempTob.GetString('PJO_JOURJ1PLAGE1') = '-' then
              MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE1')
            else { Durée jusqu'à minuit }
              MtCalcul := (1 - TempTob.GetDouble('PJO_HORDEBPLAGE1')) * 24;
            if TempTob.GetString('PJO_JOURJ1PLAGE2') = '-' then
              MtCalcul := MtCalcul + TempTob.GetDouble('PJO_DUREEPLAGE2')
            else begin
              if TempTob.GetString('PJO_JOURJ1PLAGE1') = '-' then
              begin
                { Si toute la plage 2 est sur le second jour, seule la plage 1 est prise en compte, sinon : }
                if TempTob.GetDouble('PJO_HORDEBPLAGE2') > TempTob.GetDouble('PJO_HORFINPLAGE2') then
                  { on ajoute la Durée jusqu'à minuit }
                  MtCalcul := MtCalcul + ((1 - TempTob.GetDouble('PJO_HORDEBPLAGE2') ) * 24);
              end;
            end;
          end;
        end;
        10 : begin  { Nombre d'heures sur la journée suivante}
          if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := 0;
            if TempTob.GetString('PJO_JOURJ1PLAGE1') = '-' then
            begin
              if TempTob.GetString('PJO_JOURJ1PLAGE2') = 'X' then
                { Si toute la plage 2 est sur le second jour on prends la durée total de la plage 2 }
                if TempTob.GetDouble('PJO_HORDEBPLAGE2') < TempTob.GetDouble('PJO_HORFINPLAGE2') then
                  MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE2')
                else
                { Sinon, on prends seulement la partie depuis minuit }
                  MtCalcul := TempTob.GetDouble('PJO_HORFINPLAGE2') * 24;
            end else
              { Durée depuis minuit + durée plage 2}
              MtCalcul := TempTob.GetDouble('PJO_HORFINPLAGE1') * 24 + TempTob.GetDouble('PJO_DUREEPLAGE2');
          end;
        end;
        11..12 : begin  { Journée spéciale }
          result := True;
          MtCalcul := 0;
          DateATester := DateFin;

          if Code = 12 then DateATester := IncDay(DateATester);
          decodedate(DateAtester, AnneeM, MoisM, JourM);
          if DayOfTheWeek(DateATester) = DaySunday then   MtCalcul := 1; // dimanche

          if IsJourFerie(ObjCalcuVarPre.Tob_Ferie, DateATester) then  MtCalcul := 2; // jour férié
          if ((JourM = 1) and (MoisM = 1)) then MtCalcul := 3;   //  1er mai

        end;
        13..16 : begin // Heures théoriques plage 1 (avec / sans exception)
         if Assigned(TempTob) then
          begin
            result := True;
            MtCalcul := 0;
            if ((code = 13) or (code = 14)) then MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE1')
            else MtCalcul := TempTob.GetDouble('PJO_DUREEPLAGE2');
          end;
        end;
      end;
    end;
    30..33 : begin // Recherche affectation
      result := True;
      case Code of
        30 : MtCalcul := JourneeType;
        31 : MtCalcul := Cycle;
        32 : MtCalcul := Modele;
        33 : MtCalcul := Profil;
      //  34 : begin
      //    ObjCalcuVarPre.GetJourneeTypeSalarie(DateFin, CodeSal, sCycleOuModele, sTypeDeCycle, bIsExceptionSalarie, bIsExceptionCycle);
      //    ObjCalcuVarPre.GetModeleAndJourneeInCycle(sModele, sJournee, dDateLancement, DateDebutCycle, DateFinCycle, iResteJours, DateFin, sCycleOuModele, sTypeDeCycle);
      //    MtCalcul := DateFinCycle - DateDebutCycle;
      //  end;
      end;
    end;
    40..57 : begin // Info journée Type
      TempTob := ObjCalcuVarPre.TobJourneesTypes.FindFirst(['PJO_JOURNEETYPE'],[JourneeType], False);
      if Assigned(TempTob) then
      begin
        result := True;
        case Code of
          40 : MtCalcul := TempTob.GetDouble('PJO_HORDEBPLAGE1') * 24;
          41 : MtCalcul := TempTob.GetDouble('PJO_HORFINPLAGE1') * 24;
          42 : MtCalcul := TempTob.GetDouble('PJO_HORDEBPLAGE2') * 24;
          43 : MtCalcul := TempTob.GetDouble('PJO_HORFINPLAGE2') * 24;
          44 : MtCalcul := ValeurI(TempTob.GetString('PJO_TYPEJOUR'));
          45 : MtCalcul := ValeurI(TempTob.GetString('PJO_TYPEJOUR1'));
          46 : MtCalcul := ValeurI(TempTob.GetString('PJO_TYPEJOUR2'));
          47 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE2');
          48 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE3');
          49 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE4');
          50 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE5');
          51 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE6');
          52 : MtCalcul := TempTob.GetDouble('PJO_TEMPSLIBRE7');
          53 : if TempTob.GetBoolean('PJO_JOURTRAVFERIE') then MtCalcul := 1 else MtCalcul := 0;
          54..57 : Begin
            DureePause := TempTob.GetDouble('PJO_DUREEPAUSE');
            case Code of
              54 : if TempTob.GetBoolean('PJO_PAUSEEFFECTIF') then MtCalcul := DureePause else MtCalcul := 0;
              55 : if not TempTob.GetBoolean('PJO_PAUSEEFFECTIF') then MtCalcul := DureePause else MtCalcul := 0;
              56 : if TempTob.GetBoolean('PJO_PAIEMENTPAUSE') then MtCalcul := DureePause else MtCalcul := 0;
              57 : if not TempTob.GetBoolean('PJO_PAUSEEFFECTIF') then MtCalcul := DureePause else MtCalcul := 0;
            end;
           end;
          58 : MtCalcul := TempTob.GetDouble('PJO_DUREENOTRAV');
          59 : MtCalcul := TempTob.GetDouble('PJO_DUREENOTRAVEFF');
        end;
      end;
    end;
    80 : begin
        decodedate(Datedebut, AnneeM, MoisM, JourM);
        Mtcalcul := MoisM;
        end;
    81 : begin
        decodedate(Datefin, AnneeM, MoisM, JourM);
        Mtcalcul := MoisM;
        end;
    69..78 : begin // Profil de présence
      TempTob := GetLastValideFromTob(DateFin, 'PPQ_DATEVALIDITE', ObjCalcuVarPre.TobProfilPresence, ['PPQ_PROFILPRES'], [Profil]);
      if Assigned(TempTob) then
      begin
        result := True;
        case Code of
          69 : MtCalcul := TempTob.Getstring('PPQ_NBJOURSCYCLE');
          70 : MtCalcul := TempTob.GetString('PPQ_TYPEPROFILPRES');
          71 : MtCalcul := TempTob.GetString('PPQ_TPSTRAVAIL');
          72 : MtCalcul := TempTob.GetDouble('PPQ_LIMITEHTEMODUL');
          73 : MtCalcul := TempTob.GetDouble('PPQ_LIMITEBASMODUL');
          74..78 : begin
            ValeurIdem := 0;
            if Code = 74 then
            begin
              TypeIdem := TempTob.GetString('PPQ_TYPSEUILT1');
              ValeurIdem := TempTob.GetDouble('PPQ_SEUILT1');
            end else if Code = 75 then
            begin
              TypeIdem := TempTob.GetString('PPQ_TYPSEUILT2');
              ValeurIdem := TempTob.GetDouble('PPQ_SEUILT2');
            end else if Code = 76 then
            begin
              TypeIdem := TempTob.GetString('PPQ_TYPSEUILT3');
              ValeurIdem := TempTob.GetDouble('PPQ_SEUILT3');
            end else if Code = 77 then
            begin
              TypeIdem := TempTob.GetString('PPQ_TYPCONTINGHSUP');
              ValeurIdem := TempTob.GetDouble('PPQ_CONTINGHSUP');
            end else if Code = 78 then
            begin
              TypeIdem := TempTob.GetString('PPQ_TYPEPLAFHREAN');
              ValeurIdem := TempTob.GetDouble('PPQ_PLAFHREANNU');
            end;
            if TypeIdem = 'PER' then
              MtCalcul := ValeurIdem
            else if TypeIdem = 'ELT' then
              MtCalcul := ValEltNat(lpad(FloatToStr(ValeurIdem),4), DateFin, LogGetChildLevel(Diag));
          end;
        end;
      end;
    end;
  end;
end;


{ Fonction qui calcule les variables prédéfinies et qui ne correspondent pas à une
expression mais qui font appel à des champs de tables
Codes Réservés de 1 à 200
//PT55 : 09/08/2002 PH V582 Evaluation des variables de tests avec des strings
                         Mtcalcul variant au lieu de Double
}

function RechVarDefinie(const VariablePaie: string; const DateDebut, DateFin: TDateTime; var MtCalcul: Variant; TOB_Rub: TOB; Diag: TObject = nil): Boolean; //PT172
var
  Code: Integer;
  Mois, Annee: WORD;
  DateAnciennete, Date1, Date2, ZDate: TDateTime;
  RegulMois: Integer;
  TypVar, Val, Etab, CodeSal, LeTypSmic, LeSmic: string;
  T_Etab, TAT, T1, TAcpt: TOB;
  RendRub: TRendRub;
  Borne, C1, CPR1, C2, CPR2, Velt, Abt, TauxAbt, Nb_J, NB_h: Double;
  MM, JJ, AA: WORD;
  JourOuv, HH, MtAcpt: double;
  Acompte, NbJrTravEtab, Repos1, Repos2, st: string;
  // PT219
  // TobAcSaisieArret: Tob;
  i: Integer;
  Q: TQuery;
  AM, PM: Boolean;
  TETABSAL: TOB;
begin
  result := FALSE;
  Val := '';
  Mois := 0;
  Annee := 0;
  MtCalcul := 0;
  Code := ValeurI(VariablePaie);
  if (Code < 1) or (Code > 200) then exit;
  if LeftStr(VariablePaie,1) = 'P' then
  begin
    result := RechVarPresenceDefinie(VariablePaie, DateDebut, DateFin, MtCalcul, LogGetChildLevel(Diag));
    exit;
  end;
  CodeSal := TOB_Salarie.GetValeur(iPSA_SALARIE);
  Etab := TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT);
  if TOB_Salarie.GetValeur(iPSA_DATEANCIENNETE) <= 1 then DateAnciennete := TOB_Salarie.GetValeur(iPSA_DATEENTREE)
  else DateAnciennete := TOB_Salarie.GetValeur(iPSA_DATEANCIENNETE);
  RegulMois := TOB_Salarie.GetValeur(iPSA_REGULANCIEN);
  DateAnciennete := PLUSMOIS(DateAnciennete, RegulMois);
  result := TRUE;
  case Code of
    1: MtCalcul := TOB_Salarie.GetValeur(iPSA_HORAIREMOIS); // Horaire Mensuel Salarié
    // PT66    06/12/2002 PH V591 Variable 2 rend horaire Hebdo de la fiche salarié
    2: MtCalcul := TOB_Salarie.GetValeur(iPSA_HORHEBDO); //   "     Hebdo    "
    3: MtCalcul := TOB_Salarie.GetValeur(iPSA_HORANNUEL); //   "     Mesuel   "
    4: MtCalcul := TOB_Salarie.GetValeur(iPSA_TAUXHORAIRE); //  taux horaire salarié
    5: MtCalcul := RendCumulSalSess('04'); // Salaire de base paie en cours
    6: MtCalcul := RendCumulSalSess('03'); // Horaire de base paie en cours
    7: MtCalcul := TOB_Salarie.GetValeur(iPSA_PERSACHARGE); // Nbre de personnes A charge
    8:
      begin
        AgeSalarie(DateFin, Annee, Mois);
        MtCalcul := Mois;
      end; // Age salarié en mois
    9:
      begin
        AgeSalarie(DateFin, Annee, Mois);
        MtCalcul := Annee;
      end; //       "       Annee
    10: if DateAnciennete < 10 then
        MtCalcul := 0 else MtCalcul := DateFin - DateAnciennete + 1; // Anciennté en jours PT128
    11: if DateAnciennete < 10 then
        MtCalcul := 0 else MtCalcul := AncienneteAnnee(DateAnciennete, DateFin); // Ancienneté en Année
    12: if DateAnciennete < 10 then
        MtCalcul := 0 else MtCalcul := AncienneteMois(DateAnciennete, DateFin); //   " en mois
    13: MtCalcul := JCPACQUISTHEORIQUE; //PT75-2 Jours CP Acquis JCPACQUIS
    14: MtCalcul := DateAnciennete; // Date Ancienneté
    15: MtCalcul := TOB_Salarie.GetValeur(iPSA_DATEENTREE); //
    16: MtCalcul := TOB_Salarie.GetValeur(iPSA_DATESORTIE); //
    17: MtCalcul := TOB_Salarie.GetValeur(iPSA_DATENAISSANCE); //

    20: MtCalcul := Valeur(TOB_Salarie.GetValeur(iPSA_QUALIFICATION));
    21: MtCalcul := Valeur(TOB_Salarie.GetValeur(iPSA_COEFFICIENT));
    22: MtCalcul := Valeur(TOB_Salarie.GetValeur(iPSA_NIVEAU));
    23: MtCalcul := Valeur(TOB_Salarie.GetValeur(iPSA_INDICE));
    24: MtCalcul := TOB_Salarie.GetValeur(iPSA_PCTFRAISPROF); //
    25:
      begin
        if TOB_Salarie.GetValeur(iPSA_MULTIEMPLOY) = 'X' then
        begin
          MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMULTI) + RendCumulSalSess('01');
        end
        else MtCalcul := 0;
      end;
    26:
      begin
        // PT18 : 14/11/01 V562 PH Identification du taux AT en fonction des dates de validité
        // PT45 : 02/05/2002 Ph V582 initialisation du taux at au debut du traitement
        MtCalcul := 0; // Pour avoir au moins 1 taux même à zéro
        TAT := TOB_TauxAT.FindFirst(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur(iPSA_ORDREAT)], TRUE);
        while TAT <> nil do
        begin
          if DateDebut >= TAT.GetValue('PAT_DATEVALIDITE') then
          begin
            MtCalcul := TAT.GetValue('PAT_TAUXAT');
            break;
          end;
          TAT := TOB_TauxAT.FindNext(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur(iPSA_ORDREAT)], TRUE);
        end;
      end;
    27: MtCalcul := HTHTRAVAILLES; // Heures théoriques travaillées par le salarié sur la période
    28: MtCalcul := ChpEntete.HOuvres; // PT61 Heures ouvrés
    29: MtCalcul := ChpEntete.HOuvrables; //PT61 Heures ouvrables
    30:
      begin
        if Trentieme > 0 then MtCalcul := Trentieme else MtCalcul := 0;
      end;
    31:
      begin
        TypVar := 'ANC';
        MtCalcul := ValVarTable(TypVar, Val, DateDebut, DateFin, TOB_Rub);
      end; // Ancienneté
    32:
      begin
        TypVar := 'AGE';
        MtCalcul := ValVarTable(TypVar, Val, DateDebut, DateFin, TOB_Rub);
      end; // Age
    33:
      begin
        // DEB PT125
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
        if T_Etab <> nil then
        begin
          TypVar := T_Etab.GetValue('ETB_TYPSMIC'); // Rend soit Elt ou Variable
          Val := T_Etab.GetValue('ETB_SMIC');
        end;
        if TypVar = 'ELT' then MtCalcul := ValEltNat(Val, DateFin, LogGetChildLevel(Diag)) //PT172
        else if TypVar = 'VAR' then MtCalcul := ValVariable(Val, DateDebut, DateFin, TOB_Rub)
        else MtCalcul := 0;
        //        TypVar := 'DIV';
        //        MtCalcul := ValVarTable(TypVar, Val, DateDebut, DateFin, TOB_Rub);
              // FIN PT125
      end; // Minimum Conventionnel
    34:
      begin
        if RendCumulSalSess('04') <> 0 then MtCalcul := (RendCumulSalSess('03') / RendCumulSalSess('04')) // Taux horaire de base reel
        else MtCalcul := 0;
      end;
    35:
      begin // Horaire Etablissement
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
{$ENDIF}
        if T_Etab <> nil then MtCalcul := T_Etab.GetValue('ETB_HORAIREETABL');
      end;
    36: MtCalcul := ChpEntete.HeuresTrav; // Heures travaillées  theoriques
    37: MtCalcul := ChpEntete.Ouvres; // nbre de jours ouvrés
    38: MtCalcul := ChpEntete.Ouvrables; // nbre de jours ouvrables
    39: MtCalcul := JTHTRAVAILLES; // Jours théorique travaillés par le salarié sur la période
    40: MtCalcul := HCPPRIS; // Heures CP pris
    41: MtCalcul := HCPPAYES; // Heures CP payées
    42: MtCalcul := HCPASOLDER; // Heures CP A Solder
    43: MtCalcul := JCPACQUIS; // Jours CP Acquis

    44:
      begin
        // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
        if BullCompl <> 'X' then
        begin
          //   if ActionSld <> TaCreation then //mvi PT42 mise en commentaire pour calcul mt absence même en création
          GestionCalculAbsence(tob_rub, Datedebut, datefin, 'PRI');
          MtCalcul := JCPPRIS; // Jours CP Pris
        end;
      end;
    45: MtCalcul := JCPPAYES; // Jours CP Payer
    46: MtCalcul := JCPASOLDER; // Jours CP à solder
    47: MtCalcul := CPMTABS; // Montant absence CP
    48: MtCalcul := CPMTINDABS; // Calcul Indemnité CP
    49: MtCalcul := CPMTINDCOMP; // Calcul Indemnité Compensatrice CP
    50:
      begin
        // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
        if BullCompl <> 'X' then
        begin
          if (PP) and (ActionSld = TaModification) then
            PP := false else
          begin
            //PT69-2 Ajout Condition date sortie
            Date1 := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
            if ((Date1 >= DateDebut) and (Date1 <= DateFin)) or (SLDCLOTURE) then //PT75-1
            begin
              IntegreMvtSolde(Tob_rub, TOB_Salarie, Etab, CodeSal, DateDebut, DateFin); //PT-9 ajout Tob_Salarie
              GestionCalculAbsence(tob_rub, Datedebut, datefin, 'SLD'); //mvi
            end;
            ActionSld := TaModification;
          end;
          MtCalcul := JCPSOLDE; // Jours CP solde
          PP := false;
        end;
      end;
    51: MtCalcul := CPMTABSSOL; // Montant absence solde
    52: MtCalcul := CPMTINDABSSOL; // Calcul Indemnité solde
    53: MtCalcul := HABSPRIS; //PT33 Ajout nouvel variable : Heures Absences pris
    54: MtCalcul := JABSPRIS; //PT33 Ajout nouvel variable : Jours Absences pris

    55: MtCalcul := RendCumulSalSess('13') + RendCumulSalSess('14'); // BAse CSG CRDS en Cours
    56: MtCalcul := RendCumulSalSess('15') + RendCumulSalSess('17'); // Base CSG Deductible
    57: MtCalcul := RendCumulSalSess('16'); // Base CSG Deductible

    60: MtCalcul := RendCumulSalSess('01'); // Brut mois
    61: MtCalcul := RendCumulSalSess('02'); // Brut Fiscal mois
    62: MtCalcul := RendCumulSalSess('10'); // Net A Payer Mois
    63: MtCalcul := RendCumulSalSess('09'); // Net Imposable Mois
    64: MtCalcul := RendCumulSalSess('20'); // Heures réellement travaillées
    65:
      begin //
        Borne := RendCumulSalSess('20'); // Heures travaillees salaries
        {PT53 Mise en commentaire :
        C1    := ValVariable ('35',DateDebut,DateFin, TOB_Rub);    // Horaire etablissement
         if Borne > C1 then Borne := C1;}
        MtCalcul := (RendCumulSalSess('31')) - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
      end;
    66:
      begin
        Borne := RendCumulSalSess('20'); // Heures travaillees salaries
        {PT53 Mise en commentaire :
         C1    := ValVariable ('35',DateDebut,DateFin, TOB_Rub);    // Horaire etablissement
         if Borne > C1 then Borne := C1;}
        MtCalcul := RendCumulSalSess('31') - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    67:
      begin
        MtCalcul := (RendCumulSalSess('31')) - (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.5); //PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.5); //PT172
      end;
    68:
      begin
        MtCalcul := RendCumulSalSess('31') - (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.5); //PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    // PT67    06/12/2002 PH V591 Nouvelles Variables 72 à 75 + modif 70 et 71
    70:
      begin
        // PT99   : 19/03/2004 PH V_50 FQ 11106 Variable 0070 et 0071
        Borne := RendCumulSalSess('22');
        if Borne > 130 then Borne := 130;
        MtCalcul := (RendCumulSalSess('31')) - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.2); //PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.2); //PT172
      end;
    71:
      begin
        Borne := RendCumulSalSess('22');
        if Borne > 130 then Borne := 130;
        MtCalcul := (RendCumulSalSess('31')) - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.2); //PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    72:
      begin
        Borne := RendCumulSalSess('20'); // Heures travaillees salaries
        C1 := ValVariable('0035', DateDebut, DateFin, TOB_Rub); // horaire établissement PT81
        if Borne > C1 then Borne := C1; // limitation à l'horaire établissement
        MtCalcul := (RendCumulSalSess('31')) - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
      end;
    73:
      begin
        Borne := RendCumulSalSess('20'); // Heures travaillees salaries
        C1 := ValVariable('0035', DateDebut, DateFin, TOB_Rub); // horaire établissement  PT81
        if Borne > C1 then Borne := C1; // limitation à l'horaire établissement
        MtCalcul := RendCumulSalSess('31') - (Borne * ValEltNat('0011', DateFin, LogGetChildLevel(Diag))); //PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    74:
      begin
        Borne := RendCumulSalSess('20'); // Heures travaillees salaries
        C1 := ValVariable('0035', DateDebut, DateFin, TOB_Rub); // horaire établissement PT81
        if Borne > C1 then Borne := C1; // limitation à l'horaire établissement
        DecodeDate(DateFin, AA, MM, JJ);
        ZDate := EncodeDate(AA, 01, 01); // Calcul de la date au 01/01 en fonction de la date de fin
        MtCalcul := Borne * ValEltNat('0011', ZDate, LogGetChildLevel(Diag)) * 1.2; //PT172
        // PT85   : 09/09/2003 PH V_42 Correction variable 74
        if (RendCumulSalSess('31') - MtCalcul < 0) then MtCalcul := RendCumulSalSess('31');
      end;
    75:
      begin
        MtCalcul := RendCumulSalSess('02') - ValVariable('0074', DateDebut, DateFin, TOB_Rub);
        // PT85   : 09/09/2003 PH V_42 Correction variable 75
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    // FIN PT67
    // DEB PT111 Nouvelles variables CEGID
    76:
      begin
        MtCalcul := (RendCumulSalSess('31')) - (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.3); //PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.3); //PT172
      end;
    77:
      begin
        MtCalcul := RendCumulSalSess('31') - (RendCumulSalSess('20') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.3);//PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    78:
      begin
        MtCalcul := (RendCumulSalSess('31')) - (RendCumulSalSess('21') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.4);//PT172
        if MtCalcul < 0 then MtCalcul := RendCumulSalSess('31')
        else MtCalcul := (RendCumulSalSess('21') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.4);//PT172
      end;
    79:
      begin
        MtCalcul := RendCumulSalSess('31') - (RendCumulSalSess('21') * ValEltNat('0011', DateFin, LogGetChildLevel(Diag)) * 1.4);//PT172
        if MtCalcul < 0 then MtCalcul := 0;
      end;
    // PT111
   // PT63    07/11/2002 PH V591 Creation variable 80 rend le mois de la paie (datefin) FQ 10302
    080:
      begin
        DecodeDate(DateFin, AA, MM, JJ);
        MtCalcul := MM;
      end;
    // PT76  : 04/06/03 PH V_421 FQ 10425,10620,10655 nvelles variables
    081: MtCalcul := DateDebut; // Date de début du bulletin
    082: MtCalcul := DateFin; // Date de fin du bulletin
    083:
      begin
        if (TOB_Salarie.GetValeur(iPSA_DATESORTIE) = NULL) or
          (TOB_Salarie.GetValeur(iPSA_DATESORTIE) <= 10) then MtCalcul := EncodeDate(2099, 01, 01)
        else MtCalcul := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
      end;
    084:
      begin
        if TOB_Salarie.GetValeur(iPSA_SEXE) = 'M' then MtCalcul := 1
        else MtCalcul := 0;
      end;
    085: MtCalCul := TOB_Salarie.GetValeur(iPSA_PERSACHARGE);
    086:
      begin
        if (VH_PAIE.PGProfilFnal <> '') and (VH_PAIE.PGProfilFnal <> NULL) then MtCalcul := 9
        else MtCalcul := 0;
      end;
    // FIN PT76
    // DEB PT179
    087:
      begin
        if VH_PAIE.PGMethodHeures = '1' then MtCalcul := ValVariable('0034', DateDebut, DateFin, TOB_Rub)
        else
          begin
            MtCalcul := RendCumulSalSess('03') / ValVariable('0028', DateDebut, DateFin, TOB_Rub);
          end;
      end;
    // FIN PT179  
    090: MtCalcul := RendBasePrecarite(DateDebut, DateFin, TOB_Salarie.GetValeur(iPSA_SALARIE));
    091:
      begin
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
{$ENDIF}
        if T_Etab <> nil then MtCalcul := T_Etab.GetValue('ETB_TAUXVERSTRANS');
      end;
    092: MtCalcul := (FinDEMOIS(DateFin) - DEBUTDEMOIS(DateDebut)) + 1; // PT122 Nombre de jours calendaires du mois
    093: MtCalcul := TOB_Salarie.GetValeur(iPSA_TAUXPARTIEL); // PT202 Taux partiel DADS
    094: begin
         MtCalcul := TOB_Salarie.GetValeur(iPSA_CATDADS); // PT203 Catégorie DUCS
         if MtCalcul = '' then MtCalcul := '000';
         end;
    // PT84   : 26/08/2003 PH V_42 Mise en place saisie arret
    095..098:
      begin
        // DEB PT113 recup du montant de l'acompte avant le calcul de la saisie arret
        if VH_Paie.PGGESTACPT = true then
        begin
          MtAcpt := 0;
      // debut PT219
          If TobAcSaisieArret = Nil  then
           begin
            TobAcSaisieArret := Tob.Create('Les rubriques acomptes', nil, -1);
{Flux optimisé
          Q := OpenSQL('SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_SAISIEARRETAC="X"', True);
          TobAcSaisieArret.LoadDetailDB('Les acomptes', '', '', Q, False);
          Ferme(Q);
}
            St :='SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_SAISIEARRETAC="X"';
            TobAcSaisieArret.LoadDetailDBFromSQL('Les acomptes',st);
          end;
          for i := 0 to TobAcSaisieArret.Detail.Count - 1 do
          begin
            Acompte := TobAcSaisieArret.Detail[i].GetValue('PRM_RUBRIQUE');
            TAcpt := TOB_Rub.FindFirst(['PHB_RUBRIQUE', 'PHB_NATURERUB'], [Acompte, 'AAA'], TRUE);
            if TAcpt <> nil then
              MtAcpt := MtAcpt + TAcpt.getvalue('PHB_MTREM');
          end;
            (* TobAcSaisieArret.Free; *)
        //fin PT 219
        end
        else MtAcpt := 0;
        if not Assigned(TOB_SaisieArret) then TOB_SaisieArret := PGCalcSaisieArret(CodeSal, DateDebut, DateFin, TOB_Salarie.GetValeur(iPSA_PERSACHARGE), RendCumulSalSess('10'), MtAcpt);
        // FIN PT113
        if Assigned(TOB_SaisieArret) then MtCalcul := RendMtLigneSaisieArret(VariablePaie, TOB_SaisieArret)
        else MtCalcul := 0;
      end;

    100: MtCalcul := ValEltNat('0001', DateFin, LogGetChildLevel(Diag)) * Trentieme; // Plafond salarié proratisé  30eme //PT172
    101: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRETHEO); // Salaire Salarié temps complet
    102: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRETHEO) * Trentieme; // Salaire Temps complet 30eme
    103:
      begin // PT91   : 05/12/2003 PH V_50 Création de la variable 103 Taux horaire chomage partiel
        C1 := ValVariable('0034', DateDebut, DateFin, TOB_Rub) * 0.5;
        C2 := ValEltNat('0017', DateFin, LogGetChildLevel(Diag));//PT172
        if C1 > C2 then MtCalcul := C1
        else MtCalcul := C2;
      end;
    104: // DEB PT133 Variable qui le code etablissement si numérique si rend 0
      begin
        if IsNumeric(Etab) then MtCalcul := Valeur(Etab)
        else MtCalcul := 0;
      end;
    // FIN PT133
    105: // DEB PT134 Variable qui rend le code contrat de travail DADS-U
      begin
        st := 'SELECT PCI_TYPECONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + CodeSal + '" AND ' +
          ' PCI_DEBUTCONTRAT <="' + UsDateTime(DateFin) + '" ORDER BY PCI_DEBUTCONTRAT DESC '; // AND PCI_FINCONTRAT >="' + UsDateTime(DateFin) + '" ';
        Q := OpenSql(st, TRUE);
        if not Q.EOF then
        begin
          st := RechDom('PGTYPECONTRAT', Q.FindField('PCI_TYPECONTRAT').AsString, TRUE);
          if IsNumeric(St) then MtCalcul := StrToInt(St)
          else MtCalcul := 0;
        end
        else MtCalcul := -1;
        Ferme(Q);
      end;
    // FIN PT134
    106: // PT135
      begin // calcul SMIC par rapport aux cumuls des heures
        MtCalcul := (RendCumulSalSess('22') + (RendCumulSalSess('23') * 1.1)
          + (RendCumulSalSess('24') * 1.25) + (RendCumulSalSess('25') * 1.33)
          + (RendCumulSalSess('26') * 1.5) + (RendCumulSalSess('27') * 2)) * ValVariable('0182', DateDebut, DateFin, TOB_Rub);
      end;
  // PT56 : 09/08/2002 PH V582 Creation des variables CEGID pour recup champs salaries
    110: MtCalCul := TOB_Salarie.GetValeur(iPSA_DATELIBRE1);
    111: MtCalCul := TOB_Salarie.GetValeur(iPSA_DATELIBRE2);
    112: MtCalCul := TOB_Salarie.GetValeur(iPSA_DATELIBRE3);
    113: MtCalCul := TOB_Salarie.GetValeur(iPSA_DATELIBRE4);
    114: MtCalCul := TOB_Salarie.GetValeur(iPSA_LIBREPCMB1);
    115: MtCalCul := TOB_Salarie.GetValeur(iPSA_LIBREPCMB2);
    116: MtCalCul := TOB_Salarie.GetValeur(iPSA_LIBREPCMB3);
    117: MtCalCul := TOB_Salarie.GetValeur(iPSA_LIBREPCMB4);
    118: MtCalCul := TOB_Salarie.GetValeur(iPSA_TRAVAILN1);
    119: MtCalCul := TOB_Salarie.GetValeur(iPSA_TRAVAILN2);
    120: MtCalCul := TOB_Salarie.GetValeur(iPSA_TRAVAILN3);
    121: MtCalCul := TOB_Salarie.GetValeur(iPSA_TRAVAILN4);
    122: MtCalCul := TOB_Salarie.GetValeur(iPSA_CODESTAT);
    // FIN PT56
    123: // DEB PT109
      begin
        if TOB_Salarie.GetValeur(iPSA_BOOLLIBRE1) = 'X' then MtCalCul := 1 // case 1 à cocher libres de la fiche salarié
        else MtCalCul := 0;
      end;
    124:
      begin
        if TOB_Salarie.GetValeur(iPSA_BOOLLIBRE2) = 'X' then MtCalCul := 1 // case 1 à cocher libres de la fiche salarié
        else MtCalCul := 0;
      end;
    125:
      begin
        if TOB_Salarie.GetValeur(iPSA_BOOLLIBRE3) = 'X' then MtCalCul := 1 // case 1 à cocher libres de la fiche salarié
        else MtCalCul := 0;
      end;
    126:
      begin
        if TOB_Salarie.GetValeur(iPSA_BOOLLIBRE4) = 'X' then MtCalCul := 1 // case 1 à cocher libres de la fiche salarié
        else MtCalCul := 0;
      end;
    127:
      begin
        if TOB_Salarie.GetValeur(iPSA_SORTIEDEFINIT) = 'X' then MtCalCul := 1 // Sortie définitive
        else MtCalCul := 0;
      end;
    128: // Jours OUVRES
      begin
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT));
        if T_Etab <> nil then
        begin
          CalCulVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebutdeMois(Datedebut), FindeMois(Datedebut), TRUE, JOUROuv, HH);
          MtCalcul := JOUROuv;
        end
        else MtCalcul := 0;
      end;
    129: // Jours Ouvrables
      begin
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT));
        if T_Etab <> nil then
        begin
          CalCulVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebutdeMois(Datedebut), FindeMois(Datedebut), FALSE, JOUROuv, HH);
          MtCalcul := JOUROuv;
        end
        else MtCalcul := 0;
      end;
    130: // Date début exercice
      begin
        if (TOB_ExerSocial <> nil) and (TOB_ExerSocial.Detail.Count > 0) then T1 := TOB_ExerSocial.Detail[0]
        else T1 := nil;
        if T1 <> nil then MtCalcul := Double(T1.GetValue('PEX_DATEDEBUT'))
        else MtCalCul := Double(IDate1900);
      end;
    131: // Date fin Exercice
      begin
        if (TOB_ExerSocial <> nil) and (TOB_ExerSocial.Detail.Count > 0) then T1 := TOB_ExerSocial.Detail[0]
        else T1 := nil;
        if T1 <> nil then MtCalcul := Double(T1.GetValue('PEX_DATEFIN'))
        else MtCalCul := Double(IDate1900);
      end;
    // FIN PT109
 // DEB PT110 Ancienneté dans le poste
    132:
      begin
        DateAnciennete := TOB_Salarie.GetValeur(iPSA_ANCIENPOSTE);
        if DateAnciennete < 10 then
          MtCalcul := 0 else MtCalcul := DateFin - DateAnciennete; // Anciennté en jours
      end;
    133:
      begin
        DateAnciennete := TOB_Salarie.GetValeur(iPSA_ANCIENPOSTE);
        if DateAnciennete < 10 then
          MtCalcul := 0 else MtCalcul := AncienneteMois(DateAnciennete, DateFin); // Ancienneté en Mois PT115
      end;
    134:
      begin
        DateAnciennete := TOB_Salarie.GetValeur(iPSA_ANCIENPOSTE);
        if DateAnciennete < 10 then
          MtCalcul := 0 else MtCalcul := AncienneteAnnee(DateAnciennete, DateFin); //  Idem en Année PT115
      end;
    135: // test si Date de sortie du salarié dans la période de paie
      begin
        if (TOB_Salarie.GetValeur(iPSA_DATESORTIE) >= DateDebut) and (TOB_Salarie.GetValeur(iPSA_DATESORTIE) <= DateFin) then MtCalcul := 1
        else MtCalcul := 0;
      end;
    // FIN PT110
    { DEB PT130 }
    { Déduction entrée }
    136:
      begin
        MtCalcul := 0;
        Date1 := TOB_Salarie.GetValeur(iPSA_DATEENTREE);
        if not ((Date1 >= DateDebut) and (Date1 <= FinDeMois(DateDebut))) then exit;
        Date2 := DebutDeMois(DateDebut);
        Nb_j := 0;
        if (TOB_Salarie.GetValeur(iPSA_CALENDRIER) <> '') then
        begin
          TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT)], True);
          NbJrTravEtab := IntToStr(TETABSAL.GetValue('ETB_NBJOUTRAV'));
          Repos1 := TETABSAL.GetValue('ETB_1ERREPOSH');
          if Repos1 <> '' then repos1 := Copy(repos1, 1, 1);
          Repos2 := TETABSAL.GetValue('ETB_2EMEREPOSH');
          if Repos2 <> '' then repos2 := Copy(repos2, 1, 1);
          while (Date1 > Date2) do { PT130 16/08/2005 Suppression du = }
          begin
            if IfJourNonTravaille(GblTob_Semaine, GblTob_Standard, GblTob_JourFerie,
              Date2, idate1900, idate1900,
              NbJrTravEtab, Repos1, Repos2,
              False, True, Nb_j, Nb_h, AM, PM) = False then
              MtCalcul := MtCalcul + 1;
            Date2 := Date2 + 1;
          end;
        end
        else
          MtCalcul := Date1 - Date2;
      end;
    { Déduction sortie }
    137:
      begin
        MtCalcul := 0;
        Date1 := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
        if not ((Date1 <= DateFin) and (Date1 >= DateDebut)) then exit;
        Date2 := FinDeMois(DateFin);
        Nb_j := 0;
        if (TOB_Salarie.GetValeur(iPSA_CALENDRIER) <> '') then
        begin
          TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT)], True);
          NbJrTravEtab := IntToStr(TETABSAL.GetValue('ETB_NBJOUTRAV'));
          Repos1 := TETABSAL.GetValue('ETB_1ERREPOSH');
          if Repos1 <> '' then repos1 := Copy(repos1, 1, 1);
          Repos2 := TETABSAL.GetValue('ETB_2EMEREPOSH');
          if Repos2 <> '' then repos2 := Copy(repos2, 1, 1);
          while (Date1 < Date2) do { PT130 16/08/2005 Suppression du = }
          begin
            if IfJourNonTravaille(GblTob_Semaine, GblTob_Standard, GblTob_JourFerie,
              Date1, idate1900, idate1900,
              NbJrTravEtab, Repos1, Repos2,
              False, True, Nb_j, Nb_h, AM, PM) = False then
              MtCalcul := MtCalcul + 1;
            Date1 := Date1 + 1;
          end;
        end
        else
          MtCalcul := Date2 - Date1;
      end;

    138: { DEB PT154 Calcul des jours fériés dans la période }
      begin
        Date1 := DateDebut;
        if not Assigned(GblTob_JourFerie) then GblTob_JourFerie := ChargeTobFerie(Date1, DateFin);
        while (Date1 <= DateFin) do
        begin
          if IfJourFerie(GblTob_JourFerie, Date1) then MtCalcul := MtCalcul + 1;
          Date1 := Date1 + 1;
        end;
      end;
         { FIN PT154 }
    { FIN PT130 }
    140: // Méthode de valorisation CP    PT118
      begin
        if TOB_Salarie.GetValeur(iPSA_VALORINDEMCP) = 'D' then MtCalcul := 1
        else if TOB_Salarie.GetValeur(iPSA_VALORINDEMCP) = 'M' then MtCalcul := 2
        else if TOB_Salarie.GetValeur(iPSA_VALORINDEMCP) = 'X' then MtCalcul := 3
        else MtCalcul := 0;
      end;
    { DEB PT164 }
    141..149: { Réservé CP }
      begin
      { Calcul de la periode N-1 }
        TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT)], True);
        if Assigned(TETABSAL) then
          i := RendPeriode(Date1, Date2, TETABSAL.GetValue('ETB_DATECLOTURECPN'), DateFin) + 1;


        if not Assigned(GblTob_VarCp) then
        begin
          St := 'SELECT * FROM ABSENCESALARIE ' +
            'WHERE PCN_SALARIE = "' + TOB_SALARIE.GetValeur(iPSA_SALARIE) + '" ' +
            'AND PCN_TYPEMVT="CPA" AND PCN_PERIODECP <= ' + IntToStr(i) + ' ' +
            'AND ((PCN_TYPECONGE="PRI" AND (PCN_MVTDUPLIQUE="X" OR PCN_CODERGRPT<>-1)  )  OR PCN_TYPECONGE<>"PRI") ' +
            'AND ((PCN_TYPECONGE="SLD" AND (PCN_MVTDUPLIQUE="X" OR PCN_CODERGRPT<>-1)  )  OR PCN_TYPECONGE<>"SLD") ';
{Flux optimisé
          Q := OpenSql(St, True);
          if not Q.Eof then
          begin
            GblTob_VarCp := Tob.Create('Les CP', nil, -1);
            GblTob_VarCp.LoadDetailDB('', '', '', Q, False);
          end
          else
            FreeAndNil(GblTob_VarCp);
          Ferme(Q);
}
         GblTob_VarCp := Tob.Create('Les CP', nil, -1);
         GblTob_VarCp.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
         if (GblTob_VarCp.Detail.count<=0) then
            FreeAndNil(GblTob_VarCp);
        end;

         //if TOB_Acquis = nil then TOB_Acquis := ChargeTobAcquis(DateFin,CodeSal);

        if Code = 141 then
        begin { Base CP N-1 }
          if Assigned(GblTob_VarCp) and (GblTob_VarCp.detail.count > 0) then
          begin
           { Les Acquis }
            MtCalcul := GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['ACQ', i, '...'], False);
           { Les Reprises d'acquis }
            MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['REP', i, '...'], False);
           { Les ajustements d'acquis + }
            MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS', 'PCN_CODETAPE'], ['AJU', i, '+', '...'], False);
           { Les Ajustements d'acquis - }
            MtCalcul := MtCalcul - GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS', 'PCN_CODETAPE'], ['AJU', i, '-', '...'], False);
           { Les Reliquats + }
            MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['REL', i, '...'], False);
          end;
        end
        else
          if Code = 142 then
          begin { Somme des jours acquis N-1 }
            if Assigned(GblTob_VarCp) and (GblTob_VarCp.detail.count > 0) then
            begin
           { Les Acquis }
              MtCalcul := GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['ACQ', i, '...'], False);
           { Les Reprises d'acquis }
              MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['REP', i, '...'], False);
           { Les Ajustements d'acquis + }
              MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS', 'PCN_CODETAPE'], ['AJU', i, '+', '...'], False);
           { Les Ajustements d'acquis - }
              MtCalcul := MtCalcul - GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS', 'PCN_CODETAPE'], ['AJU', i, '-', '...'], False);
           { Les Reliquats + }
              MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['REL', i, '...'], False);
            end;
          end
          else
            if Code = 143 then
            begin { Somme des jours pris N-1 }
              if Assigned(GblTob_VarCp) and (GblTob_VarCp.detail.count > 0) then
              begin
            { Les Pris }
                MtCalcul := GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['PRI', i, 'P'], False);
            { Les Reprise de Pris }
                MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['CPA', i, 'P'], False);
            { Les Ajustements de pris +}
                MtCalcul := MtCalcul - GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS'], ['AJU', i, '+'], False);
            { Les Ajustements de pris -}
                MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_JOURS', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS'], ['AJU', i, '-'], False);
              end;
            end
            else
              if Code = 144 then
              begin { Montant CP payés N-1 }
                if Assigned(GblTob_VarCp) and (GblTob_VarCp.detail.count > 0) then
                begin
            { Les Pris }
                  MtCalcul := GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['PRI', i, 'P'], False);
            { Les Reprise de Pris }
                  MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_CODETAPE'], ['CPA', i, 'P'], False);
            { Les Ajustements de pris +}
                  MtCalcul := MtCalcul - GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS'], ['AJU', i, '+'], False);
            { Les Ajustements de pris -}
                  MtCalcul := MtCalcul + GblTob_VarCp.Somme('PCN_BASE', ['PCN_TYPECONGE', 'PCN_PERIODECP', 'PCN_SENSABS'], ['AJU', i, '-'], False);
                end;
              end;
      end;
      { FIN PT164 }
    150: MtCalcul := StrToint(VH_Paie.PGMethodHeures); // Méthode de calcul des heures Travaillées ou Réelles PGMETHODHEURES
    // DEB PT171 PT183
    151..158:
      begin
        st := 'SELECT SUM(PCI_MONTANTCT) PCI_MONTANTCT, SUM(PCI_SALAIREMOIS1) PCI_SALAIREMOIS1, SUM(PCI_SALAIREMOIS2) PCI_SALAIREMOIS2,' +
          'SUM(PCI_SALAIREMOIS3) PCI_SALAIREMOIS3, SUM(PCI_SALAIREMOIS4) PCI_SALAIREMOIS4, SUM(PCI_SALAIREMOIS5) PCI_SALAIREMOIS5, ' +
          'SUM(PCI_ISNBEFFECTUE) PCI_ISNBEFFECTUE , SUM(PCI_HORAIREMOIS) PCI_HORAIREMOIS, SUM(PCI_TAUXHORAIRE) PCI_TAUXHORAIRE' +
          ' FROM CONTRATTRAVAIl WHERE PCI_SALARIE="' + Codsal +
          '" AND ((PCI_FINCONTRAT >="' + UsDateTime(DateDebut) +
          '") AND (PCI_FINCONTRAT <="' + UsDateTime(DateFin) + '") AND (PCI_FINCONTRAT IS NOT NULL) AND (PCI_FINCONTRAT > "' + UsDateTime(iDate1900) + '"))' +
          ' AND (PCI_DEBUTCONTRAT <="' + UsDateTime(DateFin) + '") ';
        Q := OpenSql(St, FALSE);
        if not Q.EOF then
        begin
          if Code = 151 then MtCalcul := Q.findField('PCI_ISNBEFFECTUE').AsFloat
          else if Code = 152 then MtCalcul := Q.findField('PCI_MONTANTCT').AsFloat
          else if Code = 153 then MtCalcul := Q.findField('PCI_SALAIREMOIS1').AsFloat
          else if Code = 154 then MtCalcul := Q.findField('PCI_SALAIREMOIS2').AsFloat
          else if Code = 155 then MtCalcul := Q.findField('PCI_SALAIREMOIS3').AsFloat
          else if Code = 156 then MtCalcul := Q.findField('PCI_SALAIREMOIS4').AsFloat
          else if Code = 157 then MtCalcul := Q.findField('PCI_SALAIREMOIS5').AsFloat
          else if Code = 158 then MtCalcul := Q.findField('PCI_HORAIREMOIS').AsFloat ;
        end
        else MtCalcul := 0;
        Ferme(Q);
      end;
    // FIN PT171 PT183
    159:
    begin
        st := 'SELECT PCI_TAUXHORAIRE' +
          ' FROM CONTRATTRAVAIl WHERE PCI_SALARIE="' + Codsal +
          '" AND ((PCI_FINCONTRAT >="' + UsDateTime(DateDebut) +
          '") AND (PCI_FINCONTRAT <="' + UsDateTime(DateFin) + '") AND (PCI_FINCONTRAT IS NOT NULL) AND (PCI_FINCONTRAT > "' + UsDateTime(iDate1900) + '"))' +
          ' AND (PCI_DEBUTCONTRAT <="' + UsDateTime(DateFin) + '") ORDER BY PCI_DEBUTCONTRAT';
        Q := OpenSql(St, FALSE);
        if not Q.EOF then
        MtCalcul := Q.findField('PCI_TAUXHORAIRE').AsFloat
        else MtCalcul := 0;
        FERME (Q);
    end;
    // FIN PT183
    160:
      begin // Prorata TVA etablissement du salarié
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
{$ENDIF}
        if T_Etab <> nil then MtCalcul := T_Etab.GetValue('ETB_PRORATATVA');
      end;
    162: MtCalcul := ValVariable('0061', DateDebut, DateFin, TOB_Rub) * ValVariable('0160', DateDebut, DateFin, TOB_Rub) / 100;
    164: MtCalcul := ValEltNat('0070', DateFin, LogGetChildLevel(Diag)) * ValVariable('0160', DateDebut, DateFin, TOB_Rub) / 100; //PT172
    166: MtCalcul := ValEltNat('0071', DateFin, LogGetChildLevel(Diag)) * ValVariable('0160', DateDebut, DateFin, TOB_Rub) / 100; //PT172
    170: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS1); //  Salaire Mensuel salarié 1
    171: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS2); //  Salaire Mensuel salarié 2
    172: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS3); //  Salaire Mensuel salarié 3
    173: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS4); //  Salaire Mensuel salarié 4

    174: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRANN1); // Salaire Annuel salarié 1
    175: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRANN2); // Salaire Annuel salarié 2
    176: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRANN3); // Salaire Annuel salarié 3
    177: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRANN4); // Salaire Annuel salarié 4
    //PT58 06/09/2002 PH V585 Variables pour restituer les elements de salaires 5
    178: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIREMOIS5); //  Salaire Mensuel salarié 5
    179: MtCalcul := TOB_Salarie.GetValeur(iPSA_SALAIRANN5); // Salaire Annuel salarié 5

    180: MtCalcul := TOB_Salarie.GetValeur(iPSA_PCTFRAISPROF); // % Abattement frais professionnel
    181:
      begin // abattement pour frais professionnels
        if RendDateExerSocial(DateDebut, DateFin, Date1, Date2, FALSE) then
        begin
          CPR1 := ValCumulDate('01', Date1, Date2);
          CPR2 := ValCumulDate('02', Date1, Date2);
          C1 := RendCumulSalSess('01');
          C2 := RendCumulSalSess('02');
          Velt := ValEltNat('0015', DateFin, LogGetChildLevel(Diag)); // limite abattement //PT172
          if ((CPR1 - CPR2) > Velt) then MtCalcul := 0 // plafaond abattement depasse donc ZERO
          else
          begin
            TauxAbt := ValVariable('0180', DateDebut, DateFin, TOB_Rub);
            Abt := (C1 * TauxAbt) / 100;
            MtCalcul := C1 - C2 + CPR1 - CPR2 + Abt; // calcul du differentiel
            if (MtCalcul > Velt) then MtCalcul := (Velt - (CPR1 - CPR2)) / (TauxAbt / 100) // ponderation
            else MtCalcul := C1; // abattement = brut du mois en court sur lequel on fait le calcul de l'abattement
          end;
        end;
        // la variable rend la base de l'abattement
      end;
    182:
      begin // rend valeur du smic ou minimum conventionnel de l'etablissement
        T_Etab := Paie_RechercheOptimise(TOB_Etablissement, 'ETB_ETABLISSEMENT', TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)); // $$$$
{$IFDEF aucasou}
        T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
{$ENDIF}
        if T_Etab <> nil then
        begin
          LeTypSmic := T_Etab.GetValue('ETB_TYPSMIC');
          LeSmic := T_Etab.GetValue('ETB_SMIC');
          if LeTypSmic <> '' then
          begin // le type smic est soit une element national soit une valeur dans une table
            if LeTypSmic = 'ELN' then MtCalcul := ValEltNat(LeSmic, DateFin, LogGetChildLevel(Diag))   //PT172
            else
            begin
              LeTypSmic := 'VAR';
              MtCalcul := ValVariable(LeSmic, DateDebut, DateFin, Tob_Rub);
            end;
          end;
        end;
      end;
    183:
      begin // differentiel calcul du brut par rapport Smic
        C1 := RendCumulSalSess('22') + RendCumulSalSess('23') * 1.1 + RendCumulSalSess('24') * 1.25 +
          RendCumulSalSess('25') * 1.33 + RendCumulSalSess('26') * 1.5 + RendCumulSalSess('27') * 2; // cumul des heures coefficienté
        MtCalcul := C1 * ValVariable('0182', DateDebut, DateFin, TOB_Rub);
        if MtCalcul < RendCumulSalSess('02') then MtCalcul := 0
        else MtCalcul := MtCalcul - RendCumulSalSess('02');
      end;
    184: // PT163
      begin // differentiel calcul du brut habituel par rapport Smic
        C1 := RendCumulSalSess('22') + RendCumulSalSess('23') * 1.1 + RendCumulSalSess('24') * 1.25 +
          RendCumulSalSess('25') * 1.33 + RendCumulSalSess('26') * 1.5 + RendCumulSalSess('27') * 2; // cumul des heures coefficienté
        MtCalcul := C1 * ValVariable('0182', DateDebut, DateFin, TOB_Rub);
        if MtCalcul < RendCumulSalSess('05') then MtCalcul := 0
        else MtCalcul := MtCalcul - RendCumulSalSess('05');
      end;
    190:
      begin // Recup appoint de la paie precedente
        Date1 := 0;
        Date2 := 0;
        RendDatePaiePrec(Date1, Date2, DateDebut, DateFin);
        if (Date1 <> 0) and (Date2 <> 0) then
        begin
          RendRub := ValRubDate('9992', 'AAA', Date1, Date2);
          // PT89   : 21/10/2003 PH V_42 FQ 10928 Recup de l'appoint precedant en négatif sinon mauvais sens
          MtCalcul := RendRub.MontRem * -1;
        end;
      end;
    191:
      begin // Calcul appoint de la paie en cours si salarié sorti alors pas d'appoint
        Date1 := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
        if (Date1 >= DateDebut) and (Date1 <= DateFin) then MtCalcul := 0 // Date de sortie dans la paie traitee donc pas appoint
        else
        begin
          //         if (Date1 < 10) then  PT87   : 06/10/2003 PH V_421 FQ 10595 Gstion des apppoints avec date de sortie
          MtCalcul := RendCumulSalSess('10') - ARRONDI(RendCumulSalSess('10'), 0);
          MtCalcul := ARRONDI(MtCalcul, 2) * -1;
        end;
      end;
    192:
      begin // Trop perçu rend la valeur du net à payer négatif de la paie précédente
        Date1 := 0;
        Date2 := 0;
        RendDatePaiePrec(Date1, Date2, DateDebut, DateFin);
        if (Date1 <> 0) and (Date2 <> 0) then
        begin
          MtCalcul := ValCumulDate('10', Date1, Date2);
          if MtCalcul > 0 then MtCalcul := 0;
        end;
      end;
    193: //PT185 Salarié entrée en cours de mois
      begin
        Date1 := TOB_Salarie.GetValeur(iPSA_DATEENTREE);
        if (Date1 >= DEBUTDEMOIS(DateDebut)) and (Date1 <= FINDEMOIS(DateFin)) then MtCalcul := 1
        else MtCalcul := 0;
      end;
    194: //PT186 Salarié sorti en cours de mois
      begin
        Date1 := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
        if (Date1 >= DEBUTDEMOIS(DateDebut)) and (Date1 <= FINDEMOIS(DateFin)) then MtCalcul := 1
        else MtCalcul := 0;
      end;
    195: //PT190 Nombre de salariés handicapés du mois
      begin
        St := 'SELECT COUNT (DISTINCT (PGH_SALARIE)) MT FROM HANDICAPE ' +
          ' WHERE EXISTS (SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_SALARIE=PGH_SALARIE ' +
          ' AND PPU_ETABLISSEMENT = "' + TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT) + '"' +
          ' AND (PGH_DATEFINCOT = "' + USDATETIME(iDate1900) + '"' +
          ' OR PGH_DATEFINCOT >= "' + USDATETIME(DateFin) + '")' + 
          ' AND PPU_DATEDEBUT >="' + UsDateTime(DateDebut) +
          '" AND PPU_DATEFIN<="' + UsDateTime(DateFin) + '")';
        Q := OPENSQL(st, TRUE);
        if not Q.eof then
          MtCalcul := Q.findField('MT').AsInteger
        else
          MtCalcul := 0;
        Ferme(Q);
      end;
  end;
end;

{ Fonction qui sert simplement à memoriser le calcul du trentieme
Attention, cette fonction est à utiliser obligatoirement pour affecter le trentieme.
En saisie, puisque le numérateur et le dénominateur peuvent etre saisis, elle est obligatoire
En préparation automatique, idem mais le dénominateur sera tjrs 30
}

procedure MemoriseTrentieme(const MtTrent: Double);
begin
  Trentieme := MtTrent;
end;

{ Fonction qui recupère le TOB salarié pour avoir tous les champs de la fiche salarié
à jour en fonction de la base : Il faudrait faire une requete SQL contenant que les champs utiles
pour éviter de récupèrer ts les champs : Requete trop longue
}
// PT173 rajout des paramètres date début et date de fin
procedure RecupTobSalarie(const Salarie: string; DateDebut, DateFin : TDateTime);
//La DtDebut et DtFin ne sont a renseigner que dans le cas de la récupération d'informations liées
//à 1 contrat
//procedure RecupTobSalarie(const Salarie: string; DtDebut : TDateTime = 0; DtFin: TDateTime = 0);
var
//  Q: TQuery;
  st: string;
begin
  if Assigned(GblTob_VarCp) then // PT189
    FreeAndNil(GblTob_VarCp);
  //if TOB_Salarie<>NIL then begin TOB_Salarie.free; TOB_Salarie := NIL end;
  if TOB_DUSALARIE <> nil then
  begin
    TOB_DUSALARIE.free;
    TOB_DUSALARIE := nil
  end;
  PGMajHistorique(Salarie,DateDebut,DateFin); // PT173 maj des infos salariés depuis historique par avance
  TOB_DUSALARIE := TOB.Create('Mon SALARIE', nil, -1);
  st := 'SELECT PSA_SALARIE,PSA_ETABLISSEMENT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_LIBREPCMB1,' +
    'PSA_LIBREPCMB2,PSA_LIBREPCMB3,PSA_LIBREPCMB4,PSA_CONFIDENTIEL,PSA_TAUXHORAIRE,PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,' +
    'PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_SALAIRANN1,PSA_SALAIRANN2,PSA_SALAIRANN3,PSA_SALAIRANN4,PSA_SALAIRANN5,PSA_DATENAISSANCE,PSA_DATEENTREE,' +
    'PSA_DATESORTIE,PSA_TYPPROFILREM,PSA_PROFILREM,PSA_TYPPROFIL,PSA_PROFIL,PSA_TYPPERIODEBUL,PSA_PERIODBUL,' +
    'PSA_TYPPROFILRBS,PSA_PROFILRBS,PSA_TYPREDREPAS,PSA_REDREPAS,PSA_TYPREDRTT1,PSA_REDRTT1,PSA_TYPREDRTT2,PSA_REDRTT2,';
  st := st + 'PSA_PROFILTPS,PSA_TYPPROFILAFP,PSA_PROFILAFP,PSA_TYPPROFILAPP,PSA_PROFILAPP,PSA_TYPPROFILRET,PSA_PROFILRET,' +
    'PSA_TYPPROFILMUT,PSA_PROFILMUT,PSA_TYPPROFILPRE,PSA_PROFILPRE,PSA_TYPPROFILTSS,PSA_PROFILTSS,PSA_TYPPROFILCGE,PSA_PROFILCGE,' +
    'PSA_PROFILCDD,PSA_PROFILMUL,PSA_TYPPROFILFNAL,PSA_PROFILFNAL,PSA_TYPPROFILTRANS,PSA_PROFILTRANS,PSA_TYPPROFILANC,PSA_PROFILANCIEN,' +
    'PSA_LIBELLE,PSA_PRENOM,PSA_NUMEROSS,PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_CODEPOSTAL,PSA_VILLE,PSA_INDICE,PSA_NIVEAU,PSA_CONVENTION,' +
    'PSA_CODEEMPLOI,PSA_AUXILIAIRE,PSA_DATEANCIENNETE,PSA_QUALIFICATION,PSA_COEFFICIENT,PSA_LIBELLEEMPLOI,PSA_CIVILITE,PSA_CPACQUISMOIS,PSA_NBREACQUISCP,' +
    'PSA_TYPDATPAIEMENT,PSA_MOISPAIEMENT,PSA_JOURPAIEMENT,PSA_TYPREGLT,PSA_PGMODEREGLE,PSA_REGULANCIEN,PSA_HORHEBDO,PSA_HORAIREMOIS,PSA_HORANNUEL,';
  st := st + 'PSA_PERSACHARGE,PSA_PCTFRAISPROF,PSA_MULTIEMPLOY,PSA_SALAIREMULTI,PSA_ORDREAT,PSA_SALAIRETHEO,PSA_DATELIBRE1,PSA_DATELIBRE2,PSA_DATELIBRE3,' +
    'PSA_DATELIBRE4,PSA_VALANCCP,PSA_ANCIENNETE,PSA_CALENDRIER,PSA_STANDCALEND,PSA_BOOLLIBRE1,PSA_BOOLLIBRE2,PSA_BOOLLIBRE3,PSA_BOOLLIBRE4,' +
    'PSA_DADSPROF,PSA_DADSCAT,PSA_TAUXPARTIEL,PSA_CPTYPEMETHOD,PSA_VALORINDEMCP,PSA_CPTYPEVALO,PSA_MVALOMS,PSA_VALODXMN,PSA_CPACQUISANC,PSA_BASANCCP,' +
    'PSA_VALANCCP,PSA_DATANC,PSA_TYPDATANC,PSA_DATEACQCPANC,PSA_NBRECPSUPP,PSA_CPTYPERELIQ,PSA_RELIQUAT,PSA_CONDEMPLOI,';
  st := st + ' PSA_TYPNBACQUISCP,PSA_NBACQUISCP,PSA_SEXE,PSA_SORTIEDEFINIT, '; //PT75-2 ajout new param. // PT116
  st := st + ' PSA_CONGESPAYES,PSA_CPACQUISSUPP,PSA_ANCIENPOSTE, '; { PT100-1 & PT100-2 }
  st := st + ' PSA_TYPPAIEVALOMS,PSA_PAIEVALOMS,PSA_UNITEPRISEFF,PSA_ACTIVITE,PSA_CATDADS '; { PT150, PT203 }
  st := st + 'FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_DUSALARIE.LoadDetailDB('SALARIES', '', '', Q, False);
}
  TOB_DUSALARIE.LoadDetailDBFROMSQL('SALARIES',St);

  TOB_Salarie := TOB_DUSALARIE.detail[0];
  PGRecupValeurHisto(TOB_Salarie,DateFin); // PT173 maj des donnés salarié avec rétroactivité de l'historique si modifs postèrieure
  if iPSA_ETABLISSEMENT = 0 then MemorisePsa(TOB_Salarie);
// Flux optimise  Ferme(Q);
end;
//fonction qui examine si le champ est du type valeur alors jamais de saisie

function ExamCasValeurRem(const ACol: Integer; T1: TOB): Boolean;
var
  St: string;
begin
  result := FALSE;
  case ACol of
    2: st := T1.GetValue('PRM_TYPEBASE');
    3: st := T1.GetValue('PRM_TYPETAUX');
    4: st := T1.GetValue('PRM_TYPECOEFF');
    5: st := T1.GetValue('PRM_TYPEMONTANT');
  end;
  if st <> '' then
  begin
    if StrToInt(st) > 1 then result := TRUE;
  end // @@@@ =8 au lieu de >1
  else result := TRUE;
end;

    //PT184 Gestion des variables de présence
{ Fonction d'évaluation des variables de présence }
function EvalueVarPres(TypeBase : Integer; Salarie, LaBase : String; const DateDeb, DateFin : TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = '') : Double; //TobVar : Tob;
var
  NeedSave, HadRegul : Boolean;
  Etat : String;
  TempTob, ParentTob : tob;
  function CreateTob(Titre : String; ParentTob : Tob; Champs : String; Valeur : Variant) : Tob;
  begin
    result := TOB.create(Titre, ParentTob, -1);
    result.AddChampSupValeur(Champs,Valeur);
  end;
begin
  ParentTob := ObjCalcuVarPre.TobSaveCalcVar;
  TempTob := ParentTob.FindFirst(['SALARIE'], [Salarie], False);
  if not Assigned(TempTob) then
  begin
    TempTob := CreateTob('niveau salarié', ParentTob, 'SALARIE',Salarie);
    TempTob := CreateTob('niveau type', TempTob, 'TYPEVAR',TypeBase);
    TempTob := CreateTob('niveau variable', TempTob, 'VARIABLE',LaBase);
    TempTob := CreateTob('niveau valeur', TempTob, 'DEBUT',DateDeb);
    TempTob.AddChampSupValeur('FIN',DateFin);
  end else begin
    ParentTob := TempTob;
    TempTob := ParentTob.FindFirst(['TYPEVAR'], [TypeBase], False);
    if not Assigned(TempTob) then
    begin
      TempTob := CreateTob('niveau type', ParentTob, 'TYPEVAR',TypeBase);
      TempTob := CreateTob('niveau variable', TempTob, 'VARIABLE',LaBase);
      TempTob := CreateTob('niveau valeur', TempTob, 'DEBUT',DateDeb);
      TempTob.AddChampSupValeur('FIN',DateFin);
    end else begin
      ParentTob := TempTob;
      TempTob := ParentTob.FindFirst(['VARIABLE'], [LaBase], False);
      if not Assigned(TempTob) then
      begin
        TempTob := CreateTob('niveau variable', ParentTob, 'VARIABLE',LaBase);
        TempTob := CreateTob('niveau valeur', TempTob, 'DEBUT',DateDeb);
        TempTob.AddChampSupValeur('FIN',DateFin);
      end else begin
        ParentTob := TempTob;
        TempTob := ParentTob.FindFirst(['DEBUT', 'FIN'], [DateDeb, DateFin], False);
        if not Assigned(TempTob) then
        begin
          TempTob := CreateTob('niveau valeur', ParentTob, 'DEBUT',DateDeb);
          TempTob.AddChampSupValeur('FIN',DateFin);
        end else begin
          result := TempTob.GetDouble('VALEUR');
          exit;
        end;
      end;
    end;
  end;

  case TypeBase of
    500 : begin
      result := ObjCalcuVarPre.EvalVarPreCompPres(Salarie, LaBase, periodiciteJournaliere, DateDeb, DateFin, TOB_Rub, NeedSave, HadRegul, Etat, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le compteur de présence journalière ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le compteur de présence journalière ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le compteur de présence journalière ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    501 : begin
      result := ValVariable(LaBase, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la variable de présence journalière ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la variable de présence journalière ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la variable de présence journalière ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    502 : begin
      result := ObjCalcuVarPre.EvalVarPreCompPres(Salarie, LaBase, periodiciteHebdomadaire, DateDeb, DateFin, TOB_Rub, NeedSave, HadRegul, Etat, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le compteur de présence hebdomadaire ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le compteur de présence hebdomadaire ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le compteur de présence hebdomadaire ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    503 : begin
      result := ValVariable(LaBase, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la variable de présence hebdomadaire ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la variable de présence hebdomadaire ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la variable de présence hebdomadaire ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    504 : begin
      result := ObjCalcuVarPre.EvalVarPreCompPres(Salarie, LaBase, periodiciteMensuelle, DateDeb, DateFin, TOB_Rub, NeedSave, HadRegul, Etat, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le compteur de présence mensuelle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le compteur de présence mensuelle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le compteur de présence mensuelle ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    505 : begin
      result := ValVariable(LaBase, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la variable de présence mensuelle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la variable de présence mensuelle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la variable de présence mensuelle ' + LaBase , (Salarie = ''));
//Fin PT194
    end;
    506..508 : begin
      result := ObjCalcuVarPre.EvalVarPreJType(Salarie, LaBase, DateDeb, DateFin, TypeBase, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le droit paramétré à journée type ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le droit paramétré à journée type ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le droit paramétré à journée type ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    510 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'QTE', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la quantité de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la quantité de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la quantité de l''évènement ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    512 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PCN_HEURES', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le nombre d''heures de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le nombre d''heures de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le nombre d''heures de l''évènement ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    513 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PCN_NBHEURESNUIT', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le nombre d''heures de nuit de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le nombre d''heures de nuit de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le nombre d''heures de nuit de l''évènement ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    514 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PCN_HDEB', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé l''heure de début de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé l''heure de début de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé l''heure de début de l''évènement ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    515 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PCN_HFIN', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé l''heure de fin de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé l''heure de fin de l''évènement ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé l''heure de fin de l''évènement ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    516 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'JOURS', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé si l''évènement est sur une journée complète ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur une journée complète ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur une journée complète ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    517 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PLAGE1', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 1 ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 1 ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 1 ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    518 : begin
      result := ObjCalcuVarPre.EvalVarPreEvt(Salarie, LaBase, 'PLAGE2', DateDeb, DateFin, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 2 ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 2 ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé si l''évènement est sur la plage 2 ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    520 : begin
      result := ObjCalcuVarPre.EvalVarPreCompPres(Salarie, LaBase, periodiciteFinDeCycle, DateDeb, DateFin, TOB_Rub, NeedSave, HadRegul, Etat, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le compteur de présence fin de cycle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le compteur de présence fin de cycle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le compteur de présence fin de cycle ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    521 : begin
      result := ValVariable(LaBase, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la variable de présence fin de cycle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la variable de présence fin de cycle ' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la variable de présence fin de cycle ' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    530 : begin
      result := ValVariable(LaBase, DateDeb, DateFin, TOB_Rub, LogGetChildLevel(Diag));
//Debut PT194
//      if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la variable prédéfini CEGID' + LaBase + ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la variable prédéfini CEGID' + LaBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la variable prédéfini CEGID' + LaBase, (Salarie = ''));
//Fin PT194
    end;
    else begin
      result := 0;
    end;
  end;
  TempTob.AddChampSupValeur('VALEUR',FloatToStr(result));
end;

{ Fonction d'évaluation d'une variable de paie de type VALEUR.
Les variables sont du type soit une valeur à proprement dit,
soit un élément national à une date  fixe : Exemple : Elt national au 01/01 alors
que l'on fait la paie du mois de Juillet.
Les autres cas cad Cotisation, Cumuls, Rem ... sont traités dans les variables de
type cotisation,cumul,rémunération ce qui offre plus de possibilités
}

function EvalueVarVal(T_Rech: TOB; const DateDebut, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double;
var
  i: Integer;
  Zdate: TDateTime;
  Mois, Jour, Annee: WORD;
  St, lib, Labase, StBool: string;
  TempDate : TDateTime;
begin
  result := 0;
  try
    i := StrToInt(T_Rech.GetValue('PVA_TYPEBASE0'));
  except
    LogMessage(Diag, ' Aucun type de base n''est défini pour la variable ' + VariablePaie, 'VAR', VariablePaie);
    exit;
  end;
  St := Trim(T_Rech.GetValue('PVA_DATEDEBUT'));
  if St <> '' then
  begin // Recup de l'annee dans la date et affection du jj,mois saisis
    DecodeDate(DateFin, Annee, Mois, Jour);
    Jour := StrToInt(Copy(St, 1, 2));
    Mois := StrToInt(Copy(St, 3, 2));
    ZDate := EncodeDate(Annee, Mois, Jour);
  end
  else ZDate := DateFin; // si pas de date de validite alors on prend la date de la paie
  lib := '';
  case i of
    20: lib := 'AGE';
    21: lib := 'ANC';
    22: lib := 'DIV'; // Table Divers ou minimum conventionnel
    // PT86   : 17/09/2003 PH V_421 Traitement des tables DIW pour une variable de type valeur (idem DIV)
    220: lib := 'DIW';
  end;
  LaBase := T_Rech.GetValue('PVA_BASE0');
  case i of
    2: begin
        result := ValEltNat(T_Rech.GetValue('PVA_BASE0'), ZDate, LogGetChildLevel(Diag)); //PT172
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé l''élément national ' + LaBase + ' ' + RechDom('PGELEMENTNAT',LaBase,FALSE) + ' à la date du ' + DateToStr(ZDate));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé l''élément national ' + LaBase + ' ' + RechDom('PGELEMENTNAT',LaBase,FALSE) + ' à la date du ' + DateToStr(ZDate), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé l''élément national ' + LaBase + ' ' + RechDom('PGELEMENTNAT',LaBase,FALSE), (Zdate < 10));
//Fin PT194
      end;
    4: begin
        result := T_Rech.GetValue('PVA_BASE0');
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a rendu la valeur ' + FloatToStr(result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a rendu la valeur ' + FloatToStr(result), 'VAR', VariablePaie);
//Fin PT194
      end;
    20..22:
      begin
        result := ValVarTable(lib, LaBase, DateDebut, DateFin, TOB_Rub);
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la table ' + LaBase + 'de nature' + RechDom('PGNATURECOEFF', LaBase, FALSE));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la table ' + LaBase + 'de nature' + RechDom('PGNATURECOEFF', LaBase, FALSE), 'VAR', VariablePaie);
//Fin PT194
      end;
    30: begin  //PT193
        result := ValEltDyna(TOB_Salarie, ZDate, T_Rech.GetValue('PVA_BASE0'), TempDate);
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + LaBase + ' ' + RechDom('PGZONEELTDYN',LaBase,FALSE) + ' à la date du ' + DateToStr(ZDate)+ ' avec pour résultat : ' + FloatToStr(result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + LaBase + ' ' + RechDom('PGZONEELTDYN',LaBase,FALSE) + ' à la date du ' + DateToStr(ZDate)+ ' avec pour résultat : ' + FloatToStr(result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + LaBase + ' ' + RechDom('PGZONEELTDYN',LaBase,FALSE), (Zdate < 10));
//Fin PT194
      end;
    31: begin  //PT193
        if      LaBase = 'SO_PGLIBCOCHE1' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE1)
        else if LaBase = 'SO_PGLIBCOCHE2' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE2)
        else if LaBase = 'SO_PGLIBCOCHE3' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE3)
        else if LaBase = 'SO_PGLIBCOCHE4' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE4);
        if StBool = 'X' then result := 1 else result := 0;
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',LaBase,FALSE) + ' du salarié avec pour résultat : ' + FloatToStr(result));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',LaBase,FALSE) + ' du salarié avec pour résultat : ' + FloatToStr(result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',LaBase,FALSE) + ' du salarié', (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    // PT86   : 17/09/2003 PH V_421 Traitement des tables DIW pour une variable de type valeur (idem DIV)
    220: begin
        result := ValVarTable(lib, LaBase, DateDebut, DateFin, TOB_Rub);
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la table ' + LaBase + 'de nature' + RechDom('PGNATURECOEFF', LaBase, FALSE));
      LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la table ' + LaBase + 'de nature' + RechDom('PGNATURECOEFF', LaBase, FALSE), 'VAR', VariablePaie);
//Fin PT194
      end;
    225: begin // PT175
//Debut PT194
//        if Diag <> nil then Diag.Items.add('  '+TraduireMemoire('La variable ') + VariablePaie + TraduireMemoire(' a calculé la table dynamique ') + LaBase);
      LogMessage(Diag, '  '+TraduireMemoire('La variable ') + VariablePaie + TraduireMemoire(' a calculé la table dynamique ') + LaBase, 'VAR', VariablePaie);
//Fin PT194
        result := ObjTableDyn.GetValueFromDynaTable (TOB_Salarie.GetValeur (iPSA_SALARIE) , LaBase,ClePGSynEltNAt,ValidationOk, DateDebut, DateFin, TOB_Rub, TOB_Salarie.GetValeur (iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur (iPSA_CONVENTION),0,0,LogGetChildLevel(Diag)); //PT178 //PT181
      end;
    //PT184 Gestion des variables de présence
    500..530 : result := EvalueVarPres(i, TOB_Salarie.GetValeur (iPSA_SALARIE), LaBase, DateDebut, DateFin, TOB_Rub, LogGetChildLevel(Diag), VariablePaie);//T_Rech,
  end;
end;

{ Fonction d'evaluation d'une variable de type Cumul cad qui valorise un cumul
à une date donnée
}
function EvalueVarCum(T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double;
var
  ZdatD, ZdatF: TDateTime;
  TCum: TOB;
  PerRaz, LeCum, st: string; // Periode de raz du cumul
begin
  result := 0;
  PerRaz := '';
  TCum := Paie_RechercheOptimise(TOB_Cumuls, 'PCL_CUMULPAIE', T_Rech.GetValue('PVA_BASE0')); // $$$$
{$IFDEF aucasou}
  TCum := TOB_Cumuls.findFirst(['PCL_CUMULPAIE'], [T_Rech.GetValue('PVA_BASE0')], FALSE);
{$ENDIF}
  if TCum <> nil then PerRaz := Tcum.getvalue('PCL_RAZCUMUL');
  LeCum := T_Rech.GetValue('PVA_BASE0');
  RendDateVar(T_Rech, DateDeb, DateFin, ZdatD, ZdatF);
  // PT28 : 27/12/2001 V571 PH Calcul des variables de tupe cumul en tenant compte du mois de RAZ
  if T_Rech.GetValue('PVA_PERIODECALCUL') = '005' then // Variable recherchant un cumul antérieur
    // PT70 : 23/01/2003 PH V591 Correction du calcul des bornes de dates pour la valorisation d'un cumul antérieur avec une période de raz
    BorneDateCumul(DateDeb, DateFin, ZdatD, ZdatF, PerRaz);
  // FIN PT28
  if (ZdatD = -1) and (Zdatf = -1) then exit; // variable non calculable session en dehors de la periode de validite
  if (ZdatD = 0) or (ZdatF = 0) then result := RendCumulSalSess(T_Rech.GetValue('PVA_BASE0'))
  else result := ValCumulDate(T_Rech.GetValue('PVA_BASE0'), ZdatD, ZdatF);
  // PT21 : 20/11/01 V562 PH Modif du calcul des variables recherchant un cumul annuel + paie en cours
  if T_Rech.GetValue('PVA_PERIODECALCUL') = '007' then
    result := result + RendCumulSalSess(T_Rech.GetValue('PVA_BASE0'));
  if (T_Rech.GetValue('PVA_PERIODECALCUL') = '002') and (T_Rech.GetValue('PVA_TRIMESTRE') = 'T00') then
    result := result + RendCumulSalSess(T_Rech.GetValue('PVA_BASE0')); // PT157
  if Diag <> nil then
  begin
    if (ZdatD = 0) or (ZdatF = 0) then st := ' de la paie en cours'
    else st := ' sur la période du ' + DateToStr(ZdatD) + ' au ' + DateToStr(ZdatF);
    if (T_Rech.GetValue('PVA_PERIODECALCUL') = '007') or ((T_Rech.GetValue('PVA_PERIODECALCUL') = '002') and (T_Rech.GetValue('PVA_TRIMESTRE') = 'T00')) then
      st := st + ' + la valeur du cumul sur la paie en cours';
//Debut PT194
//    Diag.Items.add('   La variable ' + VariablePaie + ' a repris la valeur du cumul ' + T_Rech.GetValue('PVA_BASE0') + St);
    LogMessage(Diag, '   La variable ' + VariablePaie + ' a repris la valeur du cumul ' + T_Rech.GetValue('PVA_BASE0') + St, 'CUM', T_Rech.GetValue('PVA_BASE0'));
//Fin PT194
  end;
end;

//PT184
{ Fonction d'evaluation d'une variable de type Cumul de présence cad qui somme les compteurs de présence d'une période donnée
}
function EvalueVarCumulPresence(T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; Diag: TObject = nil; VariablePaie: string = ''): Double;
var
  CodeSal, ACumuler, TypeACumuler, Periode, deb, fin: string;
  NbrPeriodeGliss, index : Integer;
  DateDeDebut, DateDeFin, TpDate, DateEntreeSalarie, DateSortieSalarie : TDateTime;
  TobCompteurCalcul, TobCompteursCalcul : Tob;
  CycleOuModele, TypeDeCycle : String;
  isExceptionSalarie, isExceptionCycle : Boolean;
begin
  Result := 0;
  CodeSal := TOB_Salarie.GetValeur(iPSA_SALARIE);
  DateEntreeSalarie := TOB_Salarie.GetValeur(iPSA_DATEENTREE);
  DateSortieSalarie := TOB_Salarie.GetValeur(iPSA_DATESORTIE);
  ACumuler := T_Rech.GetValue('PVA_BASE0');
  TypeACumuler := T_Rech.GetValue('PVA_TYPEBASE0');
  Periode := T_Rech.GetValue('PVA_PERIODECALCUL'); // PT112
  NbrPeriodeGliss := T_Rech.GetInteger('PVA_NBREMOISGLISS');
  { Passage du format court jj/mm au format long jj/mm/yyyy }
  deb := ObjCalcuVarPre.Str5cDateToStr10cDate(Date() , T_Rech.GetString('PVA_DATEDEBUT'));
  fin := ObjCalcuVarPre.Str5cDateToStr10cDate(Date() , T_Rech.GetString('PVA_DATEFIN'));
  case VALEURI(Periode) of
    4 : begin  { Date à Date }
      DateDeDebut := StrToDate(deb);
      DateDeFin := StrToDate(fin);
    end;
    8 : begin  { Mois en cours }
    //  ObjCalcuVarPre.GetCurrentMonth(DateDeDebut, DateDeFin);
      ObjCalcuVarPre.GetMonth(Datedeb,DateDeDebut, DateDeFin);
    end;
    9 : begin  { Semaine en cours }
   //   ObjCalcuVarPre.GetCurrentWeek(DateDeDebut, DateDeFin);
      ObjCalcuVarPre.GetWeek(Datedeb, DateDeDebut, DateDeFin);
    end;
    10 : begin  { Mois glissants (hors mois en cours) }
      ObjCalcuVarPre.GetMonth(NbrPeriodeGliss, DateFin, DateDeDebut, TpDate);
      ObjCalcuVarPre.GetMonth(1, DateFin, TpDate, DateDeFin);
    end;
    11 : begin  { Semaine glissante (hors semaine en cours) }
      ObjCalcuVarPre.GetWeek(NbrPeriodeGliss, DateDeDebut, TpDate, DateFin);
      ObjCalcuVarPre.GetWeek(1, TpDate, DateDeFin, DateFin);
    end;
    12 : begin  { Depuis (hors période en cours) }
      DateDeDebut := StrToDate(deb);
      ObjCalcuVarPre.GetMonth(1, DateFin, TpDate, DateDeFin);
    end;
    13 : begin  { Semaines précédentes du mois en cours }
      ObjCalcuVarPre.GetMonth(Datedeb,DateDeDebut, TpDate);
      ObjCalcuVarPre.GetWeek(Datedeb, DateDeFin, TpDate);
      DateDeFin := DateDeFin - 1;
    end;
    14 : begin  { Cycle en cours }
      ObjCalcuVarPre.GetJourneeTypeSalarie(Datedeb, CodeSal, CycleOuModele, TypeDeCycle, isExceptionSalarie, isExceptionCycle);
      ObjCalcuVarPre.GetDatesModeleCycle(DateDeDebut, DateDeFin, Datedeb, CycleOuModele, TypeDeCycle);
    end;
  end;
  { Pour les variables et compteurs de présence, on découpe la période en fonction de la périodicité }
  TobCompteursCalcul := Tob.Create('Compteurs à calculer',nil,-1);
  TobCompteurCalcul := Tob.Create('Compteur à calculer',TobCompteursCalcul,-1);
  if    ((TypeACumuler >= '500') and (TypeACumuler <= '505'))
     or (TypeACumuler  = '520') or (TypeACumuler  = '521') then
  begin
    // Création d'une TOB contenant les informations de compteur pour la période découpée
    // Obligation de créer la tob complète car le traitement suivant se base sur les indices
    // de champs et non sur les noms
    TobCompteurCalcul.AddChampSup('PYR_PREDEFINI', False);
    TobCompteurCalcul.AddChampSup('PYR_NODOSSIER', False);
    TobCompteurCalcul.AddChampSupValeur('PYR_COMPTEURPRES', ACumuler);
    TobCompteurCalcul.AddChampSup('PYR_DATEVALIDITE', False);
    TobCompteurCalcul.AddChampSup('PYR_LIBELLE', False);
    if (TypeACumuler = '500') or (TypeACumuler = '501') then TobCompteurCalcul.AddChampSupValeur('PYR_PERIODICITEPRE', periodiciteJournaliere);
    if (TypeACumuler = '502') or (TypeACumuler = '503') then TobCompteurCalcul.AddChampSupValeur('PYR_PERIODICITEPRE', periodiciteHebdomadaire);
    if (TypeACumuler = '504') or (TypeACumuler = '505') then TobCompteurCalcul.AddChampSupValeur('PYR_PERIODICITEPRE', periodiciteMensuelle);
    if (TypeACumuler = '520') or (TypeACumuler = '521') then TobCompteurCalcul.AddChampSupValeur('PYR_PERIODICITEPRE', periodiciteFinDeCycle);
    TobCompteurCalcul.AddChampSupValeur('PYR_PERIODERAZ', '');
    TobCompteurCalcul.AddChampSup('PYR_INTEGREPAIE', False);
    TobCompteurCalcul.AddChampSup('PYR_RUBRIQUE', False);
    TobCompteurCalcul.AddChampSup('PYR_TYPECHAMPRUB', False);
    TobCompteurCalcul.AddChampSup('PYR_VARIABLEPRES', False);
    TobCompteurCalcul.AddChampSup('PYR_EDITPLANPRES', False);
    TobCompteurCalcul.AddChampSup('PYR_PGCOLORPRE', False);
    TobCompteurCalcul.AddChampSup('PYR_THEMEPRE', False);
    TobCompteurCalcul.AddChampSupValeur('DATEDEBUT', DateDeDebut);
    TobCompteurCalcul.AddChampSupValeur('DATEFIN', DateDeFin);
    ObjCalcuVarPre.Traiteperiodicite(TobCompteursCalcul, CodeSal, DateEntreeSalarie, DateSortieSalarie,
//      True, True, True, True, True,
      DateDeDebut, DateDeDebut, DateDeDebut, DateDeDebut, DateDeDebut,
      DateDeFin, DateDeFin, DateDeFin, DateDeFin, DateDeFin );
  end;
  if (TypeACumuler >= '510') and (TypeACumuler <= '513') then
  begin
    TobCompteurCalcul.AddChampSupValeur('DATEDEBUT', DateDeDebut);
    TobCompteurCalcul.AddChampSupValeur('DATEFIN', DateDeFin);
  end;
  { Cumul des valeurs sur la période parametrée }
  for index := 0 to TobCompteursCalcul.FillesCount(0)-1 do
  begin
    TobCompteurCalcul := TobCompteursCalcul.Detail[index];
    DateDeDebut := TobCompteurCalcul.GetDateTime('DATEDEBUT');
    DateDeFin :=   TobCompteurCalcul.GetDateTime('DATEFIN');
    result := Result + EvalueVarPres(VALEURI(TypeACumuler), CodeSal, ACumuler, DateDeDebut, DateDeFin, TOB_Rub, LogGetChildLevel(Diag), VariablePaie);
  end;
//Debut PT194
//  if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé le cumul avec pour résultat : ' + FloatToStr(Result));
    LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé le cumul avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, ' La variable ' + VariablePaie + ' a calculé le cumul.');
//Fin PT194
  FreeAndNil(TobCompteursCalcul);
end;

{ Fonction d'evaluation d'une variable de type Rubrique cad qui valorise une cotisation ou rémunération
à une date donnée
}
function EvalueVarRub(const Nature: string; T_Rech: TOB; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; TypeB, ChampBase: string; Diag: TObject = nil): Double;
var
  ZdatD, ZdatF: TDateTime;
  Calcul: TRendRub;
  lib, Rub, TypC, Periode: string;
  TRech: TOB;
  AT: Boolean;
begin
  result := 0;
  lib := '';
  ZdatD := 0;
  ZdatF := 0;
  if (TypeB = '') and (ChampBase = '') then
  begin // cas evaluation d'une variable de type cotisation ou rémunération
    Rub := T_Rech.GetValue('PVA_BASE0');
    TypC := T_Rech.GetValue('PVA_TYPEBASE0');
    Periode := T_Rech.GetValue('PVA_PERIODECALCUL'); // PT112
  end
  else
  begin // cas evaluation d'un opérande d'une variable qui recupère des infos des rubriques de rem ou de cot
    Rub := Champbase;
    TypC := TypeB;
    Periode := ''; // PT112
  end;
  RendDateVar(T_Rech, DateDeb, DateFin, ZdatD, ZdatF);
  if (ZdatD = -1) and (Zdatf = -1) then exit; // variable non calculable session en dehors de la periode de valisite
  if (ZdatD = 0) or (ZdatF = 0) then
  begin
    if Nature = 'REM' then
      EvalueRem(TOB_Rub, Rub, Calcul.BasRem, Calcul.TauxRem, Calcul.CoeffRem, Calcul.MontRem, lib, DateDeb, DateFin, taModification, 0)
    else
      if Nature = 'COT' then
      begin
        EvalueCot(TOB_Rub, Rub, Calcul.BasCot, Calcul.TSal, Calcul.TPat, Calcul.MSal, Calcul.MPat, lib, DateDeb, DateFin, taModification, AT);
      // PT64    07/11/2002 PH V591 Variable de paie rend le taux patronal ou salarial
        if (TypC = '25') or (TypC = '26') then
        begin
          Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['COT', Rub], TRUE);
          if Trech = nil then exit;
          if TypC = '25' then result := Trech.GetValue('PHB_TAUXSALARIAL')
          else if TypC = '26' then result := Trech.GetValue('PHB_TAUXPATRONAL');
        end;
      end
      else
      begin
        if (Nature >= '05') and (Nature <= '08') then
        begin
          Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Rub], TRUE);
          if Trech = nil then exit;
        end
        else if (Nature >= '09') and (Nature <= '10') then
        begin
          Trech := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE); // $$$$
          if Trech = nil then exit;
        end
        // PT64    07/11/2002 PH V591 Variable de paie rend le taux patronal ou salarial FQ 10276
        else if ((Nature >= '12') and (Nature <= '14')) or ((Nature >= '25') and (Nature <= '26')) then
        begin
          Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['COT', Rub], TRUE);
          if Trech = nil then exit;
        end
        else if (Nature >= '16') and (Nature <= '19') then
        begin
          Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['BAS', Rub], TRUE);
          if Trech = nil then exit;
        end;
        if Nature = '20' then lib := 'AGE'
        else if Nature = '21' then lib := 'ANC'
        else if Nature = '22' then lib := 'MIN'
        // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
        else if Nature = '220' then lib := 'DIW'
        // PT175
        else if Nature = '225' then result := ObjTableDyn.GetValueFromDynaTable (TOB_Salarie.GetValeur (iPSA_SALARIE) , Rub,ClePGSynEltNAt,ValidationOK, DateDeb, DateFin,TOB_Rub, TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur (iPSA_CONVENTION),0,0,Nil)   //PT178 //PT181
        else if Nature = '02' then result := ValEltNat(Rub, DateFin, LogGetChildLevel(Diag)) //PT172
        else if Nature = '03' then result := ValVariable(Rub, DateDeb, DateFin, TOB_Rub)
        else if Nature = '04' then result := Valeur(Rub)
        else if Nature = '05' then result := Trech.GetValue('PHB_BASEREM')
        else if Nature = '06' then result := Trech.GetValue('PHB_TAUXREM')
        else if Nature = '07' then result := Trech.GetValue('PHB_COEFFREM')
        else if Nature = '08' then result := Trech.GetValue('PHB_MTREM')
        else if Nature = '09' then result := EvalueChampRem(TOB_Rub, Trech.GetValue('PRM_TYPEMINI'), Trech.GetValue('PRM_VALEURMINI'), Rub, DateDeb, DateFin) // Mini Rem
        else if Nature = '10' then result := EvalueChampRem(TOB_Rub, Trech.GetValue('PRM_TYPEMINI'), Trech.GetValue('PRM_VALEURMAXI'), Rub, DateDeb, DateFin) // maxi rem
        else if Nature = '12' then result := Trech.GetValue('PHB_BASECOT')
        else if Nature = '13' then result := Trech.GetValue('PHB_MTPATRONAL')
        else if Nature = '14' then result := Trech.GetValue('PHB_MTSALARIAL')
        else if Nature = '16' then result := Trech.GetValue('PHB_BASECOT')
        else if Nature = '17' then result := Trech.GetValue('PHB_TRANCHE1')
        else if Nature = '18' then result := Trech.GetValue('PHB_TRANCHE2')
        else if Nature = '19' then result := Trech.GetValue('PHB_TRANCHE3')
        else if Nature = '23' then result := StrToDate(Rub)
        // PT64    07/11/2002 PH V591 Variable de paie rend le taux patronal ou salarial
        else if Nature = '25' then result := Trech.GetValue('PHB_TAUXSALARIAL')
        else if Nature = '26' then result := Trech.GetValue('PHB_TAUXPATRONAL');
      // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
        if ((Nature >= '20') and (Nature <= '22')) or (Nature = '220') then result := ValVarTable(lib, Rub, DateDeb, DateFin, TOB_Rub);
      end;
  end
  else Calcul := ValRubDate(Rub, Nature, ZdatD, ZdatF);
  if Nature = 'REM' then
  begin
    if TypC = '05' then result := Calcul.BasRem
    else if TypC = '06' then result := Calcul.TauxRem
    else if TypC = '07' then result := Calcul.CoeffRem
    else if TypC = '08' then result := Calcul.MontRem;
    if Periode = '007' then // DEB PT112
    begin
      Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['AAA', Rub], TRUE);
      if Trech <> nil then
      begin
        if TypC = '05' then result := result + Trech.GetValue('PHB_BASEREM')
        else if TypC = '06' then result := result + Trech.GetValue('PHB_TAUXREM')
        else if TypC = '07' then result := result + Trech.GetValue('PHB_COEFFREM')
        else if TypC = '08' then result := result + Trech.GetValue('PHB_MTREM');
      end;
    end; // FIN PT112
  end;
  if Nature = 'COT' then
  begin
    if TypC = '12' then result := Calcul.BasCot
    else if TypC = '14' then result := Calcul.MSal
    else if TypC = '13' then result := Calcul.MPat;
    if Periode = '007' then // DEB PT112
    begin
      Trech := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['COT', Rub], TRUE);
      if Trech <> nil then
      begin
        if TypC = '12' then result := result + Trech.GetValue('PHB_BASECOT')
        else if TypC = '14' then result := result + Trech.GetValue('PHB_MTSALARIAL')
        else if TypC = '13' then result := result + Trech.GetValue('PHB_MTPATRONAL');
      end;
    end; // FIN PT112
  end;

end;
{ fonction qui calcule les dates de début et de fin pour borner le calcul sur une période donnée
pour les cumuls, les cotisations et les rémunérations
}

procedure RendDateVar(T_Rech: TOB; const DateDeb, DateFin: TDateTime; var ZdatD, ZdatF: TDateTime);
var
  i, Trimestre, NbreMois: Integer;
  PeriodCal: string;
  Annee, Mois, Jour, Mois1, AA, MM, JJ: WORD;
  LaDate: TDateTime;
begin
  PeriodCal := T_Rech.GetValue('PVA_PERIODECALCUL');
  if (PeriodCal = '000') or (PeriodCal = '') then // calcul sur la paie en cours
  begin
    ZdatD := 0;
    ZdatF := 0; // sur la session en cours
    exit; // Sortie de la fonction pour indiquer que c'est la session de paie en cours à prendre en compte
  end;
  if PeriodCal = '004' then // Calcul par rapport à 2 fourchettes de date
  begin
    DecodeDate(DateFin, Annee, Mois, Jour); // Pour Récupèrer l'année
    Jour := StrToInt(Copy(T_Rech.GetValue('PVA_DATEDEBUT'), 1, 2));
    Mois := StrToInt(Copy(T_Rech.GetValue('PVA_DATEDEBUT'), 3, 2));
    Mois1 := StrToInt(Copy(T_Rech.GetValue('PVA_DATEFIN'), 3, 2));
    if Mois1 < Mois then ZDatD := EncodeDate(Annee - 1, Mois, Jour)
    else ZDatD := EncodeDate(Annee, Mois, Jour);
    Jour := StrToInt(Copy(T_Rech.GetValue('PVA_DATEFIN'), 1, 2));
    ZDatF := EncodeDate(Annee, Mois1, Jour);
  end;
  if PeriodCal = '001' then // Calcul par rapport au mois
  begin
    i := (StrToInt(Copy(T_Rech.GetValue('PVA_MOISCALCUL'), 2, 2)) * -1);
    ZDatF := PLUSMOIS(DateFin, i);
    ZdatF := FindeMois(ZDatF);
    ZdatD := DebutdeMois(ZDatF);
  end;
  // PT21 : 20/11/01 V562 PH Modif du calcul des variables recherchant un cumul annuel + paie en cours
  if (PeriodCal = '003') or (PeriodCal = '007') then // Calcul par rapport à l'année
  begin
    DecodeDate(DateFin, Annee, Mois, Jour);
    // PT39 : 04/04/2002 PH V571 Modif insertion rub contenant une variable de nature cumul periode annuelle
    if (PeriodCal = '007') then i := 0
    else i := StrToInt(Copy(T_Rech.GetValue('PVA_ANNEE'), 2, 2));
    // PT30 : 27/12/2001 V571 PH Calcul des dates Trimestres,années - X et en tenant compte du décalage
    Annee := Annee - i;
    if not VH_Paie.PGDecalage then
    begin
      ZDatD := EncodeDate(Annee, 01, 01);
      ZDatF := EncodeDate(Annee, 12, 31);
    end
    else
    begin
      // PT126 Recup de la bonne année si décalage de paie
      if Mois <> 12 then ZDatD := EncodeDate(Annee - 1, 12, 01)
      else ZDatD := EncodeDate(Annee, 12, 01);
      ZDatF := EncodeDate(Annee, 11, 30);
    end;
  end;
  if PeriodCal = '002' then // Calcul par rapport au trimestre
  begin
    DecodeDate(DateFin, Annee, Mois, Jour);
    Trimestre := 4;
    // PT30 : 27/12/2001 V571 PH Calcul des dates Trimestres,années - X et en tenant compte du décalage
    if not VH_Paie.PGDecalage then
    begin
      if Mois <= 3 then Trimestre := 1
      else if Mois <= 6 then Trimestre := 2
      else if Mois <= 9 then Trimestre := 3;
      if Trimestre = 4 then
      begin
        ZDatD := EncodeDate(Annee, 10, 01);
        ZDatF := EncodeDate(Annee, 12, 31);
      end;
      if Trimestre = 3 then
      begin
        ZDatD := EncodeDate(Annee, 07, 01);
        ZDatF := EncodeDate(Annee, 09, 30);
      end;
      if Trimestre = 2 then
      begin
        ZDatD := EncodeDate(Annee, 04, 01); // PT105  04 au lieu de 03
        ZDatF := EncodeDate(Annee, 06, 30);
      end;
      if Trimestre = 1 then
      begin
        ZDatD := EncodeDate(Annee, 01, 01);
        ZDatF := EncodeDate(Annee, 03, 31);
      end;
    end
    else
    begin
      if Mois <= 2 then Trimestre := 1
      else if Mois <= 5 then Trimestre := 2
      else if Mois <= 8 then Trimestre := 3;
      if Trimestre = 4 then
      begin
        ZDatD := EncodeDate(Annee, 09, 01);
        ZDatF := EncodeDate(Annee, 11, 30);
      end;
      if Trimestre = 3 then
      begin
        ZDatD := EncodeDate(Annee, 06, 01);
        ZDatF := EncodeDate(Annee, 08, 31);
      end;
      if Trimestre = 2 then
      begin
        ZDatD := EncodeDate(Annee, 03, 01);
        ZDatF := EncodeDate(Annee, 05, 31);
      end;
      if Trimestre = 1 then
      begin
        LaDate := EncodeDate(Annee, 2, 1);
        LaDate := FINDEMOIS(LaDate);
        DecodeDate(LaDate, AA, MM, JJ);
        ZDatD := EncodeDate(Annee - 1, 12, 01);
        ZDatF := EncodeDate(Annee, 02, JJ);
      end;
    end;
    i := StrToInt(Copy(T_Rech.GetValue('PVA_TRIMESTRE'), 2, 2));
    ZDatD := PLUSMOIS(ZDatD, i * 3 * -1);
    ZDatF := FINDEMOIS(PLUSMOIS(ZDatF, i * 3 * -1));
  end;
  if PeriodCal = '005' then // Calcul par rapport depuis le Depuis de l'exercice jusqu'à date debut paie -1
  begin // Reprend en fait le cumul antérieur
    RendDateExerSocial(DateDeb, DateFin, ZDatD, ZDatF);
    if ZDatF < ZDatD then ZDatF := ZDatD;
    exit;
  end;
  //PT15 : 24/10/01 : V562 : PH Calcul periode glissnate pour variables de type cumul ou rubrique
  if PeriodCal = '006' then
  begin
    NbreMois := T_Rech.GetValue('PVA_NBREMOISGLISS');
    if NbreMois = 0 then
    begin // Pas de mois = pas de calcul
      ZDatD := -1;
      ZDatF := -1;
      exit;
    end;
    NbreMois := NbreMois * -1; // cas c'est toujours par rapport à des mois précédents
    ZDatF := DebutdeMois(DateDeb) - 1;
    // PT41 : 05/04/2002 PH V571 variable glissante sur X mois prennait 1 mois de trop ajout de 1
    NbreMois := NbreMois + 1;
    // PT50 : 01/07/2002 PH V582 modif sur clacul fourchette de dates pour les variables de type cumul : -1
    ZDatD := PLUSMOIS(ZDatF, NbreMois); // Retour de x mois par rapport fin de mois
    ZDatD := DebutdeMois(ZDatD); // On se positionne au debut du mois pour le prendre
    exit; // on sort pour éviter les contrôles de cohérence de dates des autres méthodes
  end;
  if (ZDatD > DateFin) then
  begin // cas Date debut calcul de la variable > date de fin de la session de paie = pas de calcul
    ZDatD := -1;
    ZDatF := -1;
    exit;
  end;
  if (ZDatF > DateDeb) and (PeriodCal <> '004') then
  begin // cas Date fin calcul de la variable < date de debut de la session de paie = pas de calcul
    // PT39 : 04/04/2002 V571 PH Modif insertion rub contenant une variable de nature cumul periode annuelle
    // PT50 : 01/07/2002 V582 PH modif sur clacul fourchette de dates pour les variables de type cumul
    if (PeriodCal = '007') or (PeriodCal = '003') or (PeriodCal = '001') or (PeriodCal = '002') then
    begin // PT105 Prise en compte aussi des périodes 001 et 002 et bornage Date debut
      ZDatF := DateDeb - 1;
      if ZDatD > ZDatF then ZDatD := ZDatF;
      exit;
    end;

    ZDatD := -1; // sauf dans le cas d'une variable de  calcul de date à date qui va rechercher un calcul antérieur
    ZDatF := -1;
    exit;
  end;
  if ZDatF > DateFin then
  begin // cas Date fin calcul de la variable > Date de fin de la session
    ZDatF := DateDeb - 1; // alors on exclus la sesion en cours
    if ZDatD > ZDatF then
    begin
      ZDatD := -1;
      ZDatF := -1;
      exit;
    end;
  end;
end;


//------------------------------------------------------------------------------
// fonction qui détecte la présence d'opérateur logique et construit une chaine
// d'évaluation autour de ceux ci
//------------------------------------------------------------------------------

function EvalueExpVar(TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; const Ztype, ZBase, Zoperat: string; const maxO: Integer; Diag: TObject = nil; VariablePaie: string = ''): string;
var
  chaine: string; // la chaine que je renvoie
  i: integer;
  DebI, FinI: integer; // indices bornant le controle de l'évaluation d'inégalité.
  Operateur: string; // valeur de l'opérateur en cours de traitement
  PO, PF: char;
begin
  PO := '(';
  PF := ')';
  DebI := 0;
  chaine := PO;
  for i := 0 to maxO
    do
  begin // for 1
    if T_Rech.FieldExists(Zoperat + inttostr(i)) then
    begin
      Operateur := T_Rech.GetValue(Zoperat + inttostr(i));
    end
    else Operateur := '#0';

    // l'opérateur est de type et/ou : on traite le bout de chaine jusqu'à cet opérateur !
    if IsOperateurLogique(Operateur) then
    begin
      FinI := i - 1;
      chaine := chaine + '(' + EvalueVarCalcul(TOB_Rub, T_Rech, DateDebut, DateFin, DebI, FinI, Ztype, ZBase, Zoperat, Diag, VariablePaie) + ')' + Operateur;
      DebI := i + 1;
    end;
    //       end; // fin si champ existe dans la TOB
  end; // le do begin for 1
  chaine := chaine + PO + EvalueVarCalcul(TOB_Rub, T_Rech, DateDebut, DateFin, DebI, MaxO, Ztype, ZBase, Zoperat, Diag, VariablePaie) + PF + PF;
  result := chaine;
end;

{
//------------------------------------------------------------------------------
// Fonction qui convertit une séquence d'opérandes et d'opérateurs dépourvues
// d'opérateurs logiques
//------------------------------------------------------------------------------
// PT55 : 09/08/2002 V582 PH Evaluation des variables de tests avec des strings
   Test si typebase = 24 (= String) alors pas de STRFPOINT
}

function EvalueVarCalcul(TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; const DebI, FinI: integer; const Ztype, ZBase, Zoperat: string; Diag: TObject = nil; VariablePaie: string = ''): string;
var
  i, j: integer;
  chaineI: string;
  Operateur: string;
  PO: char;
  PF: char;
begin
  PO := '(';
  PF := ')';
  chaineI := PO;
  if (Zoperat = 'PVA_OPERATRESTHEN') and (FinI = 3)
    then
  begin
    // FinI := 3;
  end;
  for i := DebI to FinI
    do
  begin // do begin no 1
    if T_Rech.FieldExists(Zoperat + inttostr(i)) then
      Operateur := T_Rech.GetValue(Zoperat + inttostr(i))
    else Operateur := '#0';
    // PT36 : 26/03/2002 V571 PH 10eme ligne de calcul des variables de type calcul
    if (Operateur = '') and (i = 9) then Operateur := 'FIN'; // Bornage de la 10Eme ligne
    //1er cas l'opérateur n'est pas de type opérateur d'inégalité, on stocke jusqu'a fin
    if (Operateur = '') and ((T_Rech.GetValue('PVA_NATUREVAR') = 'CAL') or (T_Rech.GetValue('PVA_NATUREVAR') = 'TES')) then break; // Pour sortir ???? @@@@@

    if not IsOperateurInegalite(Operateur) then
    begin // A voir
      if T_Rech.GetValue(Ztype + inttostr(i)) <> '24' then
        chaineI := chaineI + STRFPOINT(EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(i)),
          T_Rech.GetValue(ZBase + inttostr(i)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie))
      else chaineI := chaineI + EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(i)),
          T_Rech.GetValue(ZBase + inttostr(i)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie);

      if (Operateur = 'FIN') or (i = FinI)
        then
      begin
        chaineI := chaineI + PF;
        result := chaineI;
        exit; 
      end
      else chaineI := chaineI + Operateur;
      continue;
    end;
    //2ème cas l'opérateur est de type opérateur d'inégalité, on stocke jusqu'a fin
    //  On condidère qu'il n'y a plus d'autre opérateur logique dans la chaine traitée
    // sinon, il serait séparé par un opérateur logique.
    // A voir
    if T_Rech.GetValue(Ztype + inttostr(i)) <> '24' then
      chaineI := chaineI + STRFPOINT(EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(i)),
        T_Rech.GetValue(ZBase + inttostr(i)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie)) + ')'
    else
      chaineI := chaineI + EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(i)),
        T_Rech.GetValue(ZBase + inttostr(i)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie) + ')';

    chaineI := PO + chaineI + Operateur + PO;
    for j := i + 1 to FinI
      do
    begin // do begin no 2
      // A voir
      if T_Rech.GetValue(Ztype + inttostr(i)) <> '24' then
        chaineI := chaineI + STRFPOINT(EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(j)),
          T_Rech.GetValue(ZBase + inttostr(j)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie))
      else
        chaineI := chaineI + EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(j)),
          T_Rech.GetValue(ZBase + inttostr(j)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie);

      Operateur := T_Rech.GetValue(Zoperat + inttostr(j));
      if (Operateur <> 'FIN') and (Operateur <> '') then
        chaineI := chaineI + Operateur
      else
      begin
        result := chaineI + PF + PF;
        exit;
      end;
    end; // for do begin no 2
    // A voir
    if T_Rech.GetValue(Ztype + inttostr(i)) <> '24' then
      chaineI := ChaineI + STRFPOINT(EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(FinI + 1)),
        T_Rech.GetValue(ZBase + inttostr(FinI + 1)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie)) + PF + PF
    else
      chaineI := ChaineI + EvalueUnChampVar(T_Rech.GetValue(Ztype + inttostr(FinI + 1)),
        T_Rech.GetValue(ZBase + inttostr(FinI + 1)), TOB_Rub, T_Rech, DateDebut, DateFin, LogGetChildLevel(Diag), VariablePaie) + PF + PF;

    result := chaineI;
    break; // pas la peine de boucler; il ne peut y avoir 2 opérateur de
    // d'inégalité dans le bout de chaine délimité !
  end; // do begin du for  do begin no 1

  result := chaineI;
end;

//------------------------------------------------------------------------------
// Fonction qui vérifie si l'opérateur passé en paramètre est de type logique
//------------------------------------------------------------------------------

function IsOperateurLogique(const Operateur: string): boolean;
begin
  result := ((Operateur = 'AND') or (Operateur = 'OR'));
end;

//------------------------------------------------------------------------------
// Fonction qui vérifie si l'opérateur passé en paramètre est de type Inégalité
//------------------------------------------------------------------------------

function IsOperateurInegalite(const Operateur: string): boolean;
begin
  result := ((Operateur = '<') or (Operateur = '<=') or
    (Operateur = '>') or (Operateur = '>=') or
    (Operateur = '=') or (Operateur = '<>'));
end;
{ Fonction qui evalue un champ ou operande d'une variable de type calcul ou test
cas 23 specifique pour traiter des comparaisons de dates  Uniquement défini pour des
variables de paie.
// PT55 : 09/08/2002 V582 PH Evaluation des variables de tests avec des strings
          La focntion rend un Variant au lieu d'un double + case 24
}

function EvalueUnChampVar(const TypeB: string; ChampBase: string; TOB_Rub, T_Rech: TOB; const DateDebut, DateFin: TDateTime; Diag: TObject = nil; VariablePaie: string = ''): Variant;
var
  i: Integer;
  T1: TOB;
  lib: string;
  TempDate : TDateTime;
  StBool : String;
//   Base, Plafond, Plf1, Plf2, Plf3: Double;  , TR1, TR2, TR3
begin
  result := 0;
{  Base := 0;
  Plafond := 0;
   Plf1 := 0;
  Plf2 := 0;
  Plf3 := 0;
 TR1 := 0;
  TR2 := 0;
  TR3 := 0;}
  lib := '';
  if (TypeB = '') or (ChampBase = '') then exit;
  i := StrToInt(TypeB);
  if (i = 9) or (i = 10) then
  begin
    T1 := Paie_RechercheRubrique('AAA', ChampBase);
//    T1 := Tob_Rem.FindFirst(['PRM_RUBRIQUE'], [ChampBase], FALSE); // $$$$
  end
  else T1 := nil;

  case i of
    20: lib := 'AGE';
    21: lib := 'ANC';
    22: lib := 'DIV'; // Avt MIN ??? Table Divers ou minimum conventionnel
    // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
    220: lib := 'DIW';
  end;

  case i of
    2: begin
        result := ValEltNat(ChampBase, DateFin, LogGetChildLevel(Diag)); // cas d'un élement national //PT172
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a calculé l''élément national ' + ChampBase + ' ' + RechDom('PGELEMENTNAT',ChampBase,FALSE) + ' à la date du ' + DateToStr(Datefin) + ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a calculé l''élément national ' + ChampBase + ' ' + RechDom('PGELEMENTNAT',ChampBase,FALSE) + ' à la date du ' + DateToStr(Datefin) + ' avec pour résultat : ' + FloatToStr(Result), 'ELT', ChampBase, '  La variable ' + VariablePaie + ' a calculé l''élément national ' + ChampBase + ' ' + RechDom('PGELEMENTNAT',ChampBase,FALSE), (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    3: begin
        result := ValVariable(ChampBase, DateDebut, DateFin, TOB_Rub, LogGetChildLevel(Diag)); // cas d'une variable
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a calculé la variable ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a calculé la variable ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result), 'VAR', VariablePaie, '  La variable ' + VariablePaie + ' a calculé la variable ' + ChampBase, (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    4: begin
        result := VALEUR(ChampBase); // cas d'une valeur
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a comme valeur : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a comme valeur : ' + FloatToStr(Result), 'VAR', VariablePaie, '  La variable ' + VariablePaie + ' a été calculée.', (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    5..8: begin
        result := EvalueVarRub(TypeB, T_Rech, DateDebut, DateFin, TOB_Rub, TypeB, ChampBase, Diag); // cas d'une rémunération dans la session de paie
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la rémunération ' + ChampBase +
//            ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la rémunération ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result), 'REM', ChampBase
                       , '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la rémunération ' + ChampBase, (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    9:
      begin
        if T1 <> nil then result := EvalueChampRem(TOB_Rub, T1.GetValue('PRM_TYPEMINI'), T1.GetValue('PRM_VALEURMINI'), ChampBase, DateDebut, DateFin);
      end;
    10:
      begin
        if T1 <> nil then result := EvalueChampRem(TOB_Rub, T1.GetValue('PRM_TYPEMINI'), T1.GetValue('PRM_VALEURMAXI'), ChampBase, DateDebut, DateFin);
      end;
    12..14: begin
        result := EvalueVarRub(TypeB, T_Rech, DateDebut, DateFin, TOB_Rub, TypeB, ChampBase, Diag); // cas d'une cotisation dans la session de paie
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase +
//            ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result), 'COT', ChampBase
                       , '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase, (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    16..19:
      begin
        {        EvalueBas(Tob_Rub, ChampBase, Base, Plafond,Plf1,Plf2,Plf3,TR1,TR2,TR3,lib,DateDebut, DateFin, taModification,ChpEntete.BasesMod,ChpEntete.TranchesMod);
                 result:=Base;}
        result := RendBaseRubBase(ChampBase, TOB_Rub, i - 16); // cas d'une Base de cotisation dans la session de paie ou T1,T2,T3
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la base de cotisation ' + ChampBase +
//            ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la base de cotisation ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result), 'COT', ChampBase
                       , '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la base de cotisation ' + ChampBase, (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    20..22: begin
        result := ValVarTable(lib, Champbase, DateDebut, DateFin, TOB_Rub);
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) +
//            ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' avec pour résultat : ' + FloatToStr(Result), 'TAB', ChampBase
                       , '  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE), (TOB_Salarie.GetValeur(iPSA_SALARIE) = ''));
//Fin PT194
      end;
    23: begin
        result := StrToDate(ChampBase);
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a calculée la date :' + ChampBase);
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a calculée la date :' + ChampBase, 'VAR', VariablePaie);
//Fin PT194
      end;
    24: begin
        result := '''+ChampBase+''';
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a restituée la l''information suivante ' + result);
      LogMessage(Diag, '  La variable ' + VariablePaie + ' a restituée la l''information suivante ' + result, 'VAR', VariablePaie);
//Fin PT194
      end;
    // PT64    07/11/2002 PH V591 Variable de paie rend le taux patronal ou salarial
    25..26: begin
        result := EvalueVarRub(TypeB, T_Rech, DateDebut, DateFin, TOB_Rub, TypeB, ChampBase, Diag); // cas d'une cotisation dans la session de paie
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('');
//        Diag.Items.add('  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase +
//          ' avec pour résultat : ' + FloatToStr(Result));
      LogMessage(Diag, '');
      LogMessage(Diag, '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase + ' avec pour résultat : ' + FloatToStr(Result), 'COT', ChampBase
                     , '  La variable ' + VariablePaie + ' a repris : ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' de la cotisation ' + ChampBase);
//Fin PT194
      end;
    // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
    30: begin  //PT193
        result := ValEltDyna(TOB_Salarie, DateFin, ChampBase, TempDate);
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + ChampBase + ' ' + RechDom('PGZONEELTDYN',ChampBase,FALSE) + ' à la date du ' + DateToStr(DateFin)+ ' avec pour résultat : ' + FloatToStr(result));
        LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + ChampBase + ' ' + RechDom('PGZONEELTDYN',ChampBase,FALSE) + ' à la date du ' + DateToStr(DateFin)+ ' avec pour résultat : ' + FloatToStr(result), 'EDY', ChampBase
                       , ' La variable ' + VariablePaie + ' a calculé l''élément dynamique ' + ChampBase + ' ' + RechDom('PGZONEELTDYN',ChampBase,FALSE));
//Fin PT194
      end;
    31: begin  //PT193
        if      ChampBase = 'SO_PGLIBCOCHE1' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE1)
        else if ChampBase = 'SO_PGLIBCOCHE2' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE2)
        else if ChampBase = 'SO_PGLIBCOCHE3' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE3)
        else if ChampBase = 'SO_PGLIBCOCHE4' then StBool := TOB_Salarie.GetValeur(iPSA_BOOLLIBRE4);
        if StBool = 'X' then result := 1 else result := 0;
//Debut PT194
//        if Diag <> nil then Diag.Items.add(' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',ChampBase,FALSE) + ' du salarié avec pour résultat : ' + FloatToStr(result));
        LogMessage(Diag, ' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',ChampBase,FALSE) + ' du salarié avec pour résultat : ' + FloatToStr(result), 'ZLS', ChampBase
                       , ' La variable ' + VariablePaie + ' a calculé la zone libre ' + RechDom('PGZONELIBRESAL',ChampBase,FALSE) + ' du salarié.');
//Fin PT194
      end;
    220: begin
        result := ValVarTable(lib, Champbase, DateDebut, DateFin, TOB_Rub);
//Debut PT194
//        if Diag <> nil then
//          Diag.Items.add('  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) +
//            ' avec pour résultat : ' + FloatToStr(Result));
        LogMessage(Diag, '  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) + ' avec pour résultat : ' + FloatToStr(Result), 'TAB', ChampBase
                       , '  La variable ' + VariablePaie + ' a calculé une valeur à partir de la table ' + ChampBase + '  de type ' + RechDom('PGTYPECHAMPVAR', TypeB, FALSE) );
//Fin PT194
      end;
    // PT175  
    225: begin
//Debut PT194
//        if Diag <> nil then Diag.Items.add('  ' + TraduireMemoire('La variable ') + VariablePaie +  TraduireMemoire(' a calculé la table dynamique ') + ChampBase);
      LogMessage(Diag, '  ' + TraduireMemoire('La variable ') + VariablePaie +  TraduireMemoire(' a calculé la table dynamique ') + ChampBase, 'TDY', ChampBase);
//Fin PT194
        result := ObjTableDyn.GetValueFromDynaTable (TOB_Salarie.GetValeur (iPSA_SALARIE) , ChampBase,ClePGSynEltNAt,ValidationOk, DateDebut, DateFin, TOB_Rub,  TOB_Salarie.GetValeur (iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur (iPSA_CONVENTION), 0,0,LogGetChildLevel(Diag)); //PT181
      end;
        //PT184 Gestion des variables de présence
    500..530 : result := EvalueVarPres(i, TOB_Salarie.GetValeur (iPSA_SALARIE), ChampBase, DateDebut, DateFin, TOB_Rub, LogGetChildLevel(Diag), VariablePaie); //T_Rech,
  end; // Fin du case
end; // Fin fonction

{ Fonction qui rend la base de la rubrique de base dans la TOB_RUB
Si la rubrique de base n'existe elle rend 0
}

function RendBaseRubBase(const Rubrique: string; TOB_Rub: TOB; Quoi: Integer = -1): Double;
var
  T1: TOB;
begin
  result := 0;
  T1 := TOB_Rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], ['BAS', Rubrique], TRUE);
  if T1 <> nil then
  begin
    if Quoi = 0 then result := T1.GetValue('PHB_BASECOT')
    else if Quoi = 1 then result := T1.GetValue('PHB_TRANCHE1')
    else if Quoi = 2 then result := T1.GetValue('PHB_TRANCHE2')
    else if Quoi = 3 then result := T1.GetValue('PHB_TRANCHE3');
  end;
end;
{ Calcul d'une variable de type Ancienneté, minimum conventionnel, Age
La fonction va rechercher la table dossier avec la convention collective renseignée
au niveau du salarié
Si elle ne trouve pas alors on recherche la table dossier quelque soit la convention collective
PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas ==> refonte fonction
}

function ValVarTable(const TypeVar: string; var TDossier: string; const DateDebut, DateFin: TDateTime; TOB_Rub: TOB; Origine: string = ''): Double;
var
  Period, TypeNature, Champ, Profil, st: string;
  T1, T_Etab: TOB;
  AgeAnc: Double;
  Inclus, Pris: Boolean;
  Ind, zz: string;
  ii: Integer;
begin
  result := 0;
  if (TypeVar = 'ANC') or (TypeVar = 'DIV') or (TypeVar = 'DIW') or (TypeVar = 'AGE') then TypeNature := 'INT'
  else TypeNature := 'VAL'; // Pour Coeff,Qualif,Niveau,Indice
  if (TypeVar = 'ANC') then
  begin
    if TDossier = '' then
    begin
      if origine = 'CP' then
        TDossier := TOB_Salarie.GetValeur(iPSA_VALANCCP)
      else
        TDossier := TOB_Salarie.GetValeur(iPSA_ANCIENNETE);
    end;
  end;
  if (TypeVar = 'AGE') then
  begin
    if TDossier = '' then
    begin
      if Origine = 'CP' then
        TDossier := TOB_Salarie.GetValeur(iPSA_VALANCCP)
      else
        TDossier := TOB_Salarie.GetValeur(iPSA_ANCIENNETE);
    end;
  end;
  if (TypeVar = 'DIV') then
  begin
    if TDossier = '' then
    begin // ?? recuperation du smic ou MC par rapport à l'établissement et non plus par rapport au salarie (n'était pas utilisé)
      T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)], True);
      if T_Etab <> nil then TDossier := T_Etab.GetValue('ETB_SMIC');
    end;
  end;

  T1 := TOB_Minimum.FindFirst(['PMI_NATURE', 'PMI_CONVENTION', 'PMI_TYPENATURE', 'PMI_CODE'],
    [TypeVar, TOB_Salarie.GetValeur(iPSA_CONVENTION), TypeNature, TDossier], FALSE); // $$$$
  if T1 = nil then
    T1 := TOB_Minimum.FindFirst(['PMI_NATURE', 'PMI_CONVENTION', 'PMI_TYPENATURE', 'PMI_CODE'],
      [TypeVar, '000', TypeNature, TDossier], FALSE); // Cas convention non renseignee dans table dossier // $$$$
  Pris := TRUE;
  if T1 <> nil then // Recherche si un profil est associé à la table
  begin
    Profil := T1.GetValue('PMI_PROFIL');
    if Profil <> '' then // Un Profil est sélectionné
    begin
      St := T1.GetValue('PMI_INCLUS');
      if st = 'X' then Inclus := TRUE else Inclus := FALSE;
      Pris := FALSE;
      if (Inclus = TRUE) and (Profil = TOB_Salarie.GetValeur(iPSA_PROFIL)) then Pris := TRUE; // Prend que le profil
      if (Inclus = FALSE) and (Profil <> TOB_Salarie.GetValeur(iPSA_PROFIL)) then Pris := TRUE; // on exclut le profil
    end;
  end;
  if (T1 <> nil) and (Pris = TRUE) then
  begin // Table dossier identifiee
    Period := T1.GetValue('PMI_NATUREINT');
    if TypeNature = 'VAL' then // Cas Simple recherche par rapport à Coeff,Indice,Niveau,Qualification
    begin
      if TypeVar = 'COE' then Champ := TOB_Salarie.GetValeur(iPSA_COEFFICIENT)
      else if TypeVar = 'IND' then Champ := TOB_Salarie.GetValeur(iPSA_INDICE)
      else if TypeVar = 'NIV' then Champ := TOB_Salarie.GetValeur(iPSA_NIVEAU)
      else if TypeVar = 'QUA' then Champ := TOB_Salarie.GetValeur(iPSA_QUALIFICATION);
      if Champ <> '' then ; //Result:=ValTableDossier (TypeVar,Champ,TOB_Salarie.GetValeur (iPSA_CONVENTION),T1);  // Recherche par bornage de la valeur
    end
    else
    begin // Autre cas : table ancienneté, age ou Miminum conventionnel
      if (TypeVar = 'DIV') then
      begin
        if Period = 'COE' then Champ := TOB_Salarie.GetValeur(iPSA_COEFFICIENT)
        else if Period = 'IND' then Champ := TOB_Salarie.GetValeur(iPSA_INDICE)
        else if Period = 'NIV' then Champ := TOB_Salarie.GetValeur(iPSA_NIVEAU)
        else if Period = 'QUA' then Champ := TOB_Salarie.GetValeur(iPSA_QUALIFICATION);
      end;
      if (TypeVar = 'DIW') then
      begin
        zz := Copy(Period, 1, 2);
        if (zz <> 'ET') and (zz <> 'ST') then Ind := Copy(Period, 3, 1)
        else Ind := '';
        if IsNumeric(ind) then ii := NumChampTobS(zz, StrToInt(Ind)); // PT125 rajout du test si numerique
        if zz = 'OR' then Champ := TOB_Salarie.GetValeur(ii)
        else if zz = 'TC' then Champ := TOB_Salarie.GetValeur(ii)
        else if zz = 'DT' then Champ := DateToStr(TOB_Salarie.GetValeur(ii))
        else if Period = 'ETB' then Champ := TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT)
        else if Period = 'STA' then Champ := TOB_Salarie.GetValeur(iPSA_CODESTAT);
      end;

      if (TypeVar = 'AGE') or (TypeVar = 'ANC') then
      begin
        if (TypeVar = 'AGE') then
        begin // PT123 Modif Numéro des 2 variables - rajout 0
          if Period = 'MOI' then AgeAnc := ValVariable('0008', DateDebut, DateFin, TOB_RUB)
          else if Period = 'ANN' then AgeAnc := ValVariable('0009', DateDebut, DateFin, TOB_RUB)
          else if Period = 'JOU' then AgeAnc := DateFin - TOB_Salarie.GetValeur(iPSA_DATENAISSANCE);
        end;
        if (TypeVar = 'ANC') then
        begin
          if Period = 'MOI' then AgeAnc := ValVariable('0012', DateDebut, DateFin, TOB_RUB)
          else if Period = 'ANN' then AgeAnc := ValVariable('0011', DateDebut, DateFin, TOB_RUB)
          else if Period = 'JOU' then AgeAnc := ValVariable('0010', DateDebut, DateFin, TOB_RUB);
        end;
        Result := RendValeurTable(StrfMontant(AgeAnc, 15, 0, '', TRUE), TypeVar, TDossier, T1.GetValue('PMI_CONVENTION'), T1.GetValue('PMI_SENSINT'),
          T1.GetValue('PMI_NATURERESULT'), DateDebut, DateFin, TOB_RUB);
      end // Fin cas Age ou Ancienneté
      else // Cas table Minimum conventionnel
      begin
        Result := RendValeurTable(TDossier, TypeVar, Champ, T1.GetValue('PMI_CONVENTION'), T1.GetValue('PMI_SENSINT'), T1.GetValue('PMI_NATURERESULT'), DateDebut, DateFin,
          TOB_RUB, Period);
      end; // Cas table Minimum conventionnel
    end; // fin si table de type INT(ervalle)
  end; // fin si table dossier identifiee
end;
{ Fonction qui va recherche la valeur dans la table (Coeff, Qualif, Indice, Niveau)
La TOB T1 contient les données de la table des  Coeff dans laquelle on veut trouver la valeur
dans la TOB TOB_DetailMin qui contient le détail des itérations des Coeffs dans l'exemple.
}

function ValTableDossier(TypeVar, Champ, Convent: string; T1: TOB): Double; // Valorisation d'une table
var
  TDetail: TOB;
begin
  result := 0;
  TDetail := TOB_DetailMin.FindFirst(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, Champ, Convent], FALSE); // $$$$
  if TDetail = nil then
    TDetail := TOB_DetailMin.FindFirst(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, Champ, '000'], FALSE); // $$$$
  if TDetail <> nil then ;
end;
{ Fonction qui va recherche dans la TOB detail le champ resultat
AgeAnc Contient la valeur du code à tester
Champ le code de la table sur laquelle on va faire des tests
TypeVar la nature de la table
Conv la conventionj collective
Sens le sens des tests à mettre en oeuvre
La fonction rend une chaine de caracteres car le resultat peut exprimer une autre table, un montant
ou un taux  ou une variable ...
PT72  : 05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer
        revu la fonction sur les tests, cast
PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas ==> Modif la fonction
}

function RendValeurTable(const AgeAnc: string; const TypeVar, Champ, Conv, Sens, NatureRes: string; const DateDebut, DateFin: TDateTime; TOB_RUB: TOB; Period: string = ''):
  Double;
var
  TTrouve: TOB;
  // PT72  : 05/02/2003 PH V595 Traitement des tables dossiers en string au lieu d'integer
  NombreLu, NbreRecu: string;
  ValeurLue, Convent: string;
  Lu, Recu: TDateTime;
begin
  ValeurLue := '';
  if Conv = '' then Convent := '000' // valeur par défaut quand il n'y a pas de convention renseignée
  else Convent := Conv;
  if (TypeVar <> 'DIV') and (TypeVar <> 'DIW') then NbreRecu := AgeAnc
  else NbreRecu := Champ;
  ValeurLue := '';
  if (TypeVar <> 'DIV') and (TypeVar <> 'DIW') then
    TTrouve := TOB_DetailMin.FindFirst(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, Champ, Convent], FALSE) // $$$$
  else TTrouve := TOB_DetailMin.FindFirst(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, AgeAnc, Convent], FALSE); // $$$$

  while TTrouve <> nil do
  begin
    // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
    if (TypeVar = 'DIV') or (TypeVar = 'DIW') then
    begin
      if (Sens = '=') and ((TTrouve.GetValue('PCP_NBRE')) = NbreRecu) and (Copy(Period, 1, 2) <> 'DT') then
      begin
        ValeurLue := TTrouve.GetValue('PCP_TAUX');
        break;
      end;
      if (Sens = '=') and (Copy(Period, 1, 2) = 'DT') then
      begin
        Recu := StrToDAte(TTrouve.GetValue('PCP_NBRE'));
        Lu := StrToDate(NbreRecu);
        if (Lu = Recu) then
        begin
          ValeurLue := TTrouve.GetValue('PCP_TAUX');
          break;
        end;
      end;

      if TestSensTable(TTrouve.GetValue('PCP_NBRE'), NbreRecu, Sens, Period) = TRUE then
      begin
        ValeurLue := TTrouve.GetValue('PCP_TAUX');
        if (Sens = '<') or (Sens = '<=') then break;
      end;
    end
    else
    begin
      NombreLu := TTrouve.GetValue('PCP_NBRE');
      if (Sens = '=') and (NombreLu = AgeAnc) then
      begin
        ValeurLue := TTrouve.GetValue('PCP_TAUX');
        break;
      end;
      if TestSensTable(NombreLu, AgeAnc, Sens) = TRUE then // inversion du sens pour les tables apprentis
      begin
        ValeurLue := TTrouve.GetValue('PCP_TAUX');
        if (Sens = '<') or (Sens = '<=') then break;
      end;
    end;
    // PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
    if (TypeVar <> 'DIV') and (TypeVar <> 'DIW') then
      TTrouve := TOB_DetailMin.FindNext(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, Champ, Convent], FALSE) // $$$$
    else TTrouve := TOB_DetailMin.FindNext(['PCP_NATURE', 'PCP_CODE', 'PCP_CONVENTION'], [TypeVar, AgeAnc, Convent], FALSE);
  end;
  if ValeurLue = '' then
  begin
    result := 0;
    exit;
  end; // Anomalie n'a pas trouvé rend 0
  if NatureRes = 'MON' then result := Valeur(ValeurLue)
  else if NatureRes = 'VAR' then result := ValVariable(ValeurLue, DateDebut, DateFin, TOB_RUB)
    // PT73  : 05/02/2003 PH V595 Tables dossier gèrent aussi une élément national comme retour
  else if NatureRes = 'ELT' then result := ValEltNat(ValeurLue, DateFin, nil)   //PT172
  else result := ValVarTable(NatureRes, ValeurLue, DateDebut, DateFin, TOB_Rub);

end;
{ Fonction de test et de bornage en recherchent dans la TOB detail
cette fonction a uniquement pour but de dire si le test est vrai ou faux en fonction du sens
indiqué dans la table
PT74  : 06/02/2003 PH V595 Nouvelle nature DIW pour traiter nouveaux cas
         Test sur des dates
}

function TestSensTable(AgeAnc, NombreLu: Variant; Sens: string; Period: string = ''): Boolean;
var
  Lu, Recu: TDateTime;
  PrefZero: string;
begin
  result := FALSE;
  if Copy(Period, 1, 2) = 'DT' then // Cas de test sur une date
  begin
    Lu := StrToDate(NombreLu);
    Recu := StrToDate(AgeAnc);
    if Sens = '<' then if lu < Recu then result := TRUE;
    if Sens = '<=' then if Lu <= Recu then result := TRUE;
    if Sens = '>' then if Lu > Recu then result := TRUE;
    if Sens = '>=' then if Lu >= Recu then result := TRUE;
  end
  else
  begin
    { DEB PT79-1 Dans le cas de valeur numerique, on transtype en alphanumerique
    sur la longueur du code le plus long }
    if (IsNumeric(AgeAnc)) and (Isnumeric(NombreLu)) then
    begin
      if Length(AgeAnc) > Length(NombreLu) then
      begin
        PrefZero := StringOfChar('0', Length(AgeAnc) - Length(NombreLu));
        NombreLu := PrefZero + NombreLu;
      end
      else
        if Length(AgeAnc) < Length(NombreLu) then
        begin
          PrefZero := StringOfChar('0', Length(NombreLu) - Length(AgeAnc));
          AgeAnc := PrefZero + AgeAnc;
        end;
    end; { FIN PT79-1 }
    if Sens = '<' then if NombreLu < AgeAnc then result := TRUE;
    if Sens = '<=' then if NombreLu <= AgeAnc then result := TRUE;
    if Sens = '>' then if NombreLu > AgeAnc then result := TRUE;
    if Sens = '>=' then if NombreLu >= AgeAnc then result := TRUE;
  end;
end;

function ValBase(const Rubrique: string): TRendRub;
var
  T_Base: TOB;
  RendRub: TRendRub; // variable contenant le record retourné par la fonction
begin
  with RendRub do
  begin
    MontRem := 0;
    BasRem := 0;
    TauxRem := 0;
    CoeffRem := 0;
    BasCot := 0;
    TSal := 0;
    MSal := 0;
    TPat := 0;
    MPat := 0;
    Plfd1 := 0;
    Plfd2 := 0;
    Plfd3 := 0;
    Tr1 := 0;
    Tr2 := 0;
    Tr3 := 0;
    Base := 0;
    if TOB_HistoBasesCot <> nil then
    begin
      T_Base := TOB_HistoBasesCot.FindFirst(['PHB_RUBRIQUE'], [Rubrique], TRUE); // $$$$
      if T_Base <> nil then
      begin
        Plfd1 := T_Base.GetValue('P1');
        Tr1 := T_Base.GetValue('T1');
        Tr2 := T_Base.GetValue('T2');
        Tr3 := T_Base.GetValue('T3');
        Base := T_Base.GetValue('BC');
        Plfd2 := T_Base.GetValue('P2');
        Plfd3 := T_Base.GetValue('P3');
      end;
    end;
  end;
  result := RendRub;
end;

procedure ChargeBasesSal(const Salarie, Etab: string; DateD, DateF: TDateTime);
var
  st: string;
//  Q: TQUERY;
  T1: TOB;
  DatD, DatF: TDateTime; // Date Debut et fin exercice
begin
  if TOB_HistoBasesCot <> nil then TOB_HistoBasesCot.Free;
  T1 := TOB_ExerSocial.Detail[0];
  DatD := T1.GetValue('PEX_DATEDEBUT');
  DatF := DateD - 1;
  if DatF < DatD then DatF := DatD; // Cas des reprises de bases en debut d'exercice social
  st := 'select PHB_RUBRIQUE,SUM(PHB_PLAFOND1) P1,SUM(PHB_TRANCHE1) T1,SUM(PHB_TRANCHE2) T2,SUM(PHB_TRANCHE3) T3,SUM(PHB_BASECOT) BC,' +
    'SUM(PHB_PLAFOND2) P2,SUM(PHB_PLAFOND3) P3'
    + ' from HISTOBULLETIN where PHB_NATURERUB="BAS" AND PHB_SALARIE="' + Salarie +
    '" AND PHB_DATEDEBUT >="' + USDateTime(DatD) + '" AND PHB_DATEFIN <="' + USDateTime(DatF) + '"'
    + 'group by PHB_RUBRIQUE';
  TOB_HistoBasesCot := TOB.Create('BASES HISTOBULLETIN', nil, -1);
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not Q.EOF then
  begin
    TOB_HistoBasesCot.LoadDetailDB('BASES HISTOBULLETIN', '', '', Q, False);
  end;
  Ferme(Q);
}
 TOB_HistoBasesCot.LoadDetailDBFROMSQL('HISTOBULLETIN',st);
end;
// Procedur appelées en click droit

procedure Appellesalaries;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'SALARIE_MUL', '', '', '');
{$ENDIF}
end;

procedure AppelleCotisations;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'COTISATION_MUL', '', '', '');
{$ENDIF}
  ChargeCotisations;
end;

procedure AppelleVariables;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'VARIABLE_MUL', '', '', '');
{$ENDIF}
end;

procedure AppelleEtablissements;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'ETABLISSEMENT', '', '', '');
{$ENDIF}
end;

procedure AppelleDossier;
begin
{$IFNDEF EAGLSERVER}
{$IFNDEF EAGLCLIENT}
  ParamSociete(FALSE, '', 'SCO_PGPARAMETRES;SCO_PGCARACTERISTIQUES;SCO_PGCOMPTABILITE;SCO_PGDADS;SCO_PGPREFERENCES;SCO_COORDONNEES', '', nil, nil, nil, nil, 0);
{$ENDIF}
{$ENDIF}
end;

procedure AppelleRemunerations;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'REMUNERATION_MUL', '', '', '');
{$ENDIF}
end;

procedure AppelleCumuls;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'CUMUL', '', '', '');
{$ENDIF}
end;

procedure AppelleProfils;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'PROFIL_MUL', '', '', '');
{$ENDIF}
end;

procedure AppelleTablesdossier;
begin
{$IFNDEF EAGLSERVER}
  AglLanceFiche('PAY', 'COEFFICIENT_MUL', 'PMI_TYPENATURE=INT', '', '');
{$ENDIF}
end;

procedure SuppCotExclus(const ThemeExclus: string; TPE: TOB);
var
  T, T_Cot: TOB;
  Nat, Rub, Them, Themeplus: string;
begin
  Nat := 'COT';
  if ThemeExclus = '' then exit; // si pas de theme alors on sort
  if ThemeExclus = 'RDC' then Themeplus := 'RDS'
  else if ThemeExclus = 'RDS' then Themeplus := 'RDC'
  else Themeplus := 'XYZ';
  T := TPE.FindFirst(['PHB_NATURERUB'], [Nat], TRUE);
  while T <> nil do
  begin
    Rub := T.GetValue('PHB_RUBRIQUE');
    T_Cot := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, Nat], FALSE); // $$$$
    if T_Cot <> nil then // Si Cotisation trouvee
    begin
      Them := T_Cot.GetValue('PCT_THEMECOT');
      if (Them = ThemeExclus) or (Them = Themeplus) then T.Free; // suppression de l'objet dans la TOB si theme identique
    end;
    T := TPE.FindNext(['PHB_NATURERUB'], [Nat], TRUE);
  end;
end;
{ Fonction qui rend les dates de la paie précédante
}

procedure RendDatePaiePrec(var Date1, Date2: TDateTime; const DD, DF: TDateTime);
var
  st: string;
  Q: TQuery;
  ZDate: TDateTime;
  TPR: TOB;
begin
  TOB_PaiePrecedente := TOB.Create('Le Bulletin de la Paie Précédante', nil, -1);
  st := 'SELECT PPU_DATEDEBUT,PPU_DATEFIN,PPU_ETABLISSEMENT,PPU_SALARIE FROM PAIEENCOURS WHERE PPU_ETABLISSEMENT="' + TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT) +
    '" AND PPU_SALARIE="' + TOB_Salarie.GetValeur(iPSA_SALARIE) + '" ORDER BY PPU_DATEDEBUT,PPU_DATEFIN';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_PaiePrecedente.LoadDetailDB('PAIEENCOURS', '', '', Q, FALSE);
}
  TOB_PaiePrecedente.LoadDetailDBFromSQL('PAIEENCOURS', st);
  TPR := TOB_PaiePrecedente.FindFirst(['PPU_ETABLISSEMENT', 'PPU_SALARIE'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur(iPSA_SALARIE)], TRUE); // $$$$
  ferme(Q);
  Date1 := 0;
  Date2 := 0;
  while TPR <> nil do
  begin // recup des dates de la deniere paie
    ZDate := TDateTime(TPR.GetValue('PPU_DATEFIN'));
    if (Zdate < Date2) and (Date2 <> 0) then break; // cas où il y a un bulletin postérieur alors on s'arrete
    if ZDate >= DF then break;
    Date1 := TDateTime(TPR.GetValue('PPU_DATEDEBUT'));
    Date2 := TDateTime(TPR.GetValue('PPU_DATEFIN'));
    TPR := TOB_PaiePrecedente.FindNext(['PPU_ETABLISSEMENT', 'PPU_SALARIE'], [TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT), TOB_Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
  TOB_PaiePrecedente.Free;
  TOB_PaiePrecedente := nil;
end;

{ Fonction qui recherche si on va faire une division par 0
Alors remplace 0 par 1
}

procedure EvalueChaineDiv0(var St: string);
const
  Tableau: array[1..10] of string = (
    '0', '1', '2', '3', '4', '5', '6', '7', '8', '9');

var
  Pos1, Pos2, Pos3: Integer;
  St1: string;
begin
  St1 := '/0';
  Pos1 := Pos(St1, St);
  // PT80  : 11/06/2003 PH V_421 Modification fct de contrôle division par 0 Pour tester si 0.xxx au lieu de O
  if (Pos1 <> 0) and (St[Pos1 + 2] <> '.') and (St[Pos1 + 2] <> ',') then
  begin
    St[Pos1 + 1] := '1';
    for Pos2 := Pos1 - 1 downto 1 do
    begin
      St1 := Copy(St, Pos2, 1);
      if ((St1 >= '0') and (St1 <= '9')) or (St1 = '.') or (St1 = ',') then St[Pos2] := '0' //PT221 remplacer . et ,
      else break;
    end;
  end;
  Pos3 := Length(St);
  St1 := ')0';
  Pos1 := Pos(St1, St);
  if Pos1 = 0 then exit;
  for Pos2 := Pos1 to Pos3 do
  begin
    St1 := Copy(St, 1, Pos1); // recup de la chaine jusqu'à )0 inclus ) exclus 0
    if St[Pos1 + 3] <> '0' then St1 := St1 + Copy(St, Pos1 + 3, Pos3) // on exclus le 0 et la ) qui suit
    else St1 := St1 + Copy(St, Pos1 + 2, Pos3); // On exclus simplement le 0 car il est suivi normalement de 0
    St := St1;
    if (Pos(')0', St) > 0) then EvalueChaineDiv0(St)
    else break;
  end;
end;

{ intégration des mvt congé dans le bulletin
L'objectif est de générer les lignes de libellé
et de calculer toutes les variables prédéfinies utilisées dans le moteur de paie
Recherche du type de profil defini au niveau du salarié et remonte au niveau de etablissement le cas échéant
}

procedure IntegrePaye(Tob_Rub, Tob_prisCP: tob; const Etab, Salarie: string; const DateD, DateF: TDateTime; const Typecp: string; ManqueAcquis: Boolean = False);
var
  TR, TZ, T, TPR, TP, TL, TOB_libelle, THH: TOB;
  i, j: Integer;
  Rub, TypProfil, Profil: string;
  NbPris, Absence, Indemnite, NBHPris: double;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  Rub := '';
  Profil := '';
  TypProfil := '';
  FreeAndNil(TOB_Libelle);
  FreeAndNil(TR);
  NbPris := 0;
  NBHPris := 0;
  Absence := 0;
  Indemnite := 0;
  if Tob_prisCP = nil then
  begin //mvi en cas de suppression mvt cp via bulletin, passer là pour virer lignes de bulletin
  end;
  { DEB PT90 }
  Profil := PGRecupereProfilCge(Etab);
  { Integre dans function PGRecupereProfilCge
   T_ECP:=TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'],[Etab],True);
   TypProfil:=Tob_Salarie.GetValeur (iPSA_TYPPROFILCGE);   // Type Profil : Idem Etab ou Personnalisé
   if (TypProfil='ETB') OR (TypProfil='') then
     begin
     if T_ECP <> NIL then Profil:=T_ECP.GetValue('ETB_PROFILCGE')
     end
   else Profil:=Tob_Salarie.GetValeur (iPSA_PROFILCGE);}
   { FIN PT90 }
  TPR := TOB_ProfilPaies.FindFirst(['PPI_PROFIL'], [Profil], FALSE); // $$$$
  if TPR <> nil then
  begin
    i := 0;
    while i < TPR.Detail.Count do
    begin
      TZ := TPR.Detail[i]; // Recup 1ere rubrique du profil
      if TZ = nil then exit; // Cas improbable où pas de rubrique dans le profil
      Rub := TZ.GetValue('PPM_RUBRIQUE'); // Recuperation de la 1ere Rubrique du Profil CP pour creer les lignes de libelle
      // on relit la rubrique pour voir son paramétrage
      TR := TOB_Rem.findfirst(['PRM_RUBRIQUE'], [Rub], False);
      if TR <> nil then
        if IsOkType(TR, typecp) then break;
      i := i + 1;
    end;
  end; //PT90 Calcul variable même si intégration rubrique sans profil affecté
  // a voir mvi 16-01 JCPSOLDE := 0;  CPMTABSSOL := 0; CPMTINDABSSOL := 0;
  TOB_Libelle := TOB.create('Les libelles des cp', nil, -1);
  if Tob_prisCP <> nil then
  begin
    TP := Tob_prisCP.findfirst([''], [''], True);
    while tp <> nil do
    begin
      if TP.getvalue('PCN_MVTDUPLIQUE') <> 'X' then
      begin
        NbPris := NbPris + TP.getvalue('PCN_JOURS');
        NBHPris := NBHPris + TP.getvalue('PCN_HEURES');
        if TP.GetValue('PCN_MODIFABSENCE') = 'X' then
          Absence := Absence + TP.GetValue('PCN_ABSENCEMANU')
        else
          Absence := Absence + TP.GetValue('PCN_ABSENCE');
        if TP.Getvalue('PCN_MODIFVALO') <> 'X' then
        begin
          Indemnite := Indemnite + TP.GetValue('PCN_VALORETENUE');
          TP.putvalue('PCN_BASE', TP.GetValue('PCN_VALORETENUE'));
        end
        else
        begin
          Indemnite := Indemnite + TP.GetValue('PCN_VALOMANUELLE');
          TP.putvalue('PCN_BASE', TP.GetValue('PCN_VALOMANUELLE'))
        end;
        begin // Constitution de la TOB des Libellés
          TL := TOB.create('Une ligne de libelle', TOB_Libelle, -1);
          TL.AddChampSup('LIBELLE', FALSE);
          TL.PutValue('LIBELLE', TP.Getvalue('PCN_LIBELLE'));
        end;
      end;
      if Typecp = 'SLD' then
      begin
        PositionnePaye(tP, datef);
        TP.putvalue('PCN_DATEFIN', datef);
      end;
      TP := Tob_prisCP.findnext([''], [''], True);
    end;
  end;
{$IFDEF ENCOURS}
  { DEB PT88 }
  if (ManqueAcquis) and (TOB_Libelle <> nil) and (JCPPAYESPOSES - NbPris > 0) then
  begin
    TL := TOB.create('Une ligne de libelle', TOB_Libelle, -1);
    TL.AddChampSup('LIBELLE', FALSE);
    TL.PutValue('LIBELLE', FloatToStr(JCPPAYESPOSES - NbPris) + ' j. CP non intégré manque acquis');
  end;
  { FIN PT88 }
{$ENDIF}

  // mv on vire les libelles de la tob_rub créés précédemment
 {PT69-1 Modification
 recherche des lignes de commentaire que sur la nature non sur les dates, etab et salarié
  T := Tob_Rub.FindFirst(['PHB_NATURERUB','PHB_DATEDEBUT','PHB_DATEFIN','PHB_ETABLISSEMENT','PHB_SALARIE'],
                          ['AAA',DateD,DateF,Etab,CodSal],TRUE);}
  T := Tob_Rub.FindFirst(['PHB_NATURERUB'], ['AAA'], TRUE);
  while (T <> nil) and (Rub <> '') do //PT90-1 Gestion libellé si rubrique existant
  begin
    if (copy(T.GetValue('PHB_RUBRIQUE'), 1, length(Rub) + 1) = Rub + '.') then
      if T.GetValue('PHB_ORIGINELIGNE') = 'CPA' then T.free; { PT83-4 }
    T := Tob_Rub.FindNext(['PHB_NATURERUB'], ['AAA'], TRUE);
  end;

  if (Rub <> '') and (TOB_Libelle <> nil) and (Assigned(Tob_Rub.FindFirst(['PHB_RUBRIQUE'], [Rub], TRUE))) then //PT90-1 Gestion libellé si rubrique existant { PT155 }
  begin
    i := 1;
    for J := 0 to TOB_Libelle.Detail.count - 1 do
    begin
      TL := TOB_Libelle.Detail[J];
      THH := TOB.create('HISTOBULLETIN', Tob_Rub, -1);
      THH.PutValue('PHB_RUBRIQUE', Rub + '.' + IntToStr(i));
      THH.PutValue('PHB_NATURERUB', 'AAA');
      THH.PutValue('PHB_DATEDEBUT', DateD);
      THH.PutValue('PHB_DATEFIN', DateF);
      THH.PutValue('PHB_ETABLISSEMENT', Etab);
      THH.PutValue('PHB_SALARIE', CodSal);
      i := i + 1; // incrementation automatique de l'indice du commentaire @@@ voir pour
      THH.PutValue('PHB_BASEREM', 0);
      THH.PutValue('PHB_TAUXREM', 0);
      THH.PutValue('PHB_COEFFREM', 0);
      THH.PutValue('PHB_MTREM', 0);
      THH.PutValue('PHB_CONSERVATION', 'BUL');
      { PT83-4 Ajout du type CPA pour distinguer les lignes de commentaire calculé
      de celle inséré par l'utilisateur }
      THH.PutValue('PHB_ORIGINELIGNE', 'CPA');
      THH.PutValue('PHB_IMPRIMABLE', 'X');
      THH.PutValue('PHB_TRAVAILN2', Tob_Salarie.GetValeur(iPSA_TRAVAILN2));
      THH.PutValue('PHB_TRAVAILN3', Tob_Salarie.GetValeur(iPSA_TRAVAILN3));
      THH.PutValue('PHB_TRAVAILN4', Tob_Salarie.GetValeur(iPSA_TRAVAILN4));
      THH.PutValue('PHB_TRAVAILN1', Tob_Salarie.GetValeur(iPSA_TRAVAILN1));
      THH.PutValue('PHB_CODESTAT', Tob_Salarie.GetValeur(iPSA_CODESTAT));
      THH.PutValue('PHB_LIBREPCMB1', Tob_Salarie.GetValeur(iPSA_LIBREPCMB1));
      THH.PutValue('PHB_LIBREPCMB2', Tob_Salarie.GetValeur(iPSA_LIBREPCMB2));
      THH.PutValue('PHB_LIBREPCMB3', Tob_Salarie.GetValeur(iPSA_LIBREPCMB3));
      THH.PutValue('PHB_LIBREPCMB4', Tob_Salarie.GetValeur(iPSA_LIBREPCMB4));
      THH.PutValue('PHB_CONFIDENTIEL', Tob_Salarie.GetValeur(iPSA_CONFIDENTIEL));
      if TR <> nil then
      begin
        THH.PutValue('PHB_ORDREETAT', TR.GetValue('PRM_ORDREETAT'));
        THH.PutValue('PHB_OMTSALARIAL', TR.GetValue('PRM_ORDREETAT')); { PT136 }
      end;
      THH.PutValue('PHB_LIBELLE', TL.GetValue('LIBELLE'));
    end;
  end;
  if Typecp = 'PRI' then
  begin
    JCPPRIS := NbPris;
    HCPPRIS := NBHPris;
    CPMTABS := Absence;
    CPMTINDABS := Indemnite;
    JCPPAYES := NbPris;
    HCPPAYES := NBHPris;
    CPMTABS := Absence;
    CPMTINDABS := Indemnite; {PT6}
  end;
  if Typecp = 'SLD' then
  begin
    JCPSOLDE := NbPris;
    HCPASOLDER := NBHPris;
    CPMTABSSOL := Absence;
    CPMTINDABSSOL := Indemnite;
  end;
  //  end;

  if TOB_Libelle <> nil then TOB_Libelle.Free;
  TOB_Libelle := nil;
end;
{ DEB PT90 Function renvoyant le profil congés payés niveau salarié ou établissement }

function PGRecupereProfilCge(const Etab: string): string;
var
  T_ECP: Tob;
  TypProfil: string;
begin
  Result := '';
  T_ECP := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [Etab], True);
  TypProfil := Tob_Salarie.GetValeur(iPSA_TYPPROFILCGE); // Type Profil : Idem Etab ou Personnalisé
  if (TypProfil = 'ETB') or (TypProfil = '') then
  begin
    if T_ECP <> nil then Result := T_ECP.GetValue('ETB_PROFILCGE')
  end
  else Result := Tob_Salarie.GetValeur(iPSA_PROFILCGE);
end;
{ FIN PT90 }

function IsOkType(TR: tob; const typecp: string): boolean;
const
  CP: array[0..6] of PChar = ('0040', '0041', '0043', '0044', '0045', '0047', '0048');
  SLD: array[0..4] of PChar = ('0042', '0046', '0049', '0050', '0052');
var
  i: integer;
begin
  result := false;
  if typecp = 'PRI' then
    for i := 0 to 6 do
    begin
      if (((TR.GetValue('PRM_TYPEBASE') = '03') and (TR.getvalue('PRM_BASEREM') = CP[i])) or
        ((TR.GetValue('PRM_TYPETAUX') = '03') and (TR.getvalue('PRM_TAUXREM') = CP[i])) or
        ((TR.GetValue('PRM_TYPECOEFF') = '03') and (TR.getvalue('PRM_COEFFREM') = CP[i])) or
        ((TR.GetValue('PRM_TYPEMONTANT') = '03') and (TR.getvalue('PRM_MONTANT') = CP[i]))) then
      begin
        result := true;
        exit;
      end;
    end;


  if typecp = 'SLD' then
    for i := 0 to 4 do
    begin
      if (((TR.GetValue('PRM_TYPEBASE') = '03') and (TR.getvalue('PRM_BASEREM') = SLD[i])) or
        ((TR.GetValue('PRM_TYPETAUX') = '03') and (TR.getvalue('PRM_TAUXREM') = SLD[i])) or
        ((TR.GetValue('PRM_TYPECOEFF') = '03') and (TR.getvalue('PRM_COEFFREM') = SLD[i])) or
        ((TR.GetValue('PRM_TYPEMONTANT') = '03') and (TR.getvalue('PRM_MONTANT') = SLD[i]))) then
      begin
        result := true;
        exit;
      end;
    end;
end;
{PT-11 : Ajout ds declarations var Auto : Boolean;Var StMsgErr : String }

function SalIntegreCP(TOB_Sal, TOB_Rub, Tob_AcqEnCours: TOB; const DateD, DateF: TDateTime; const Auto: Boolean; var StMsgErr: string): tob; //PT7
var
  P5T_etab, TP, T1, T2: TOB;
  Etab, Salarie, st: string;
  Datesortie: TDateTime;
//  Q: TQuery;
  PrisBul: double;
  ManqueAcquis: Boolean;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  FreeAndNil(result);
  if BullCompl = 'X' then exit;
  StMsgErr := '';
  PrisBul := 0;
  ManqueAcquis := False; //PT88
  // Modif Ph Quand une tob est free, il faut la remettre à NIL obligatoirement
  if Tob_pris <> nil then if Assigned(Tob_pris) then Tob_pris := nil; //FreeAndNil(TOB_Pris); { PT160 PT160-1	}
  if TOB_Acquis <> nil then if Assigned(TOB_Acquis) then FreeAndNil(TOB_Acquis); { PT160 PT160-1	 }
  if TOB_Solde <> nil then if Assigned(TOB_Solde) then FreeAndNil(TOB_Solde); { PT160 PT160-1 }
  Etab := string(TOB_Sal.GetValeur(iPSA_ETABLISSEMENT));
  Salarie := TOB_Sal.GetValeur(iPSA_SALARIE);
  P5T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
  Tob_Pris := Tob.create('CONGES PRIS', nil, -1);
  Tob_Solde := Tob.create('SOLDE CONGES', nil, -1);

  // alimentation de la tob des acquis
  Tob_Acquis := Tob.create('CONGES ACQUIS', nil, -1);
  st := 'SELECT * from ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + Tob_SAL.GetValeur(iPSA_SALARIE) + '" ' +
    'AND PCN_TYPEMVT="CPA" ' + //PT54 Ajout PCN_TYPEMVT="CPA"
    'AND PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" AND ' +
    // mv pour arrondi  ' AND PCN_JOURS <> PCN_APAYES AND '+
// PT197  '( ( (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS")  ' +
  '( ( (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="ACF")  ' +
    ' AND PCN_DATEFIN <> "' + usdatetime(Datef) + '")' + // ajout mv 6-11-00
    ' OR PCN_TYPECONGE="ARR" OR PCN_TYPECONGE="REP" OR PCN_TYPECONGE="REB" OR ' +
    ' ( ((PCN_TYPECONGE="AJU") OR (PCN_TYPECONGE="AJP") or(PCN_TYPECONGE="REL")) AND PCN_SENSABS = "+"))' +
    ' AND PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '" ' + { PT144 }
//PT197   ' ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';
    ' ORDER BY PCN_PERIODECP DESC, PCN_DATEVALIDITE, PCN_TYPECONGE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then // pas de congés acquis
  begin
    Tob_Acquis.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    AcquisMaj := false;
  end;
  Ferme(Q);
}
    Tob_Acquis.LoadDetailDBFROMSQL('ABSENCESALARIE',st); // Flux optimisé
    if Tob_Acquis.detail.count <= 0 then AcquisMaj := false; // Flux optimisé

  if Assigned(tob_delta) then FreeAndNil(tob_delta); { PT146 }
  ChargeTobDelta(Salarie, datef, tob_delta); { PT146 }

  //DEB PT17  Duplication de la tob des acquis pour traitement de prise de CP
  if Tob_AcqEnCours <> nil then
  begin
    T1 := Tob_AcqEnCours.FindFirst([''], [''], False);
    while T1 <> nil do
    begin
      T1.PutValue('PCN_APAYES', 0);
      T1.PutValue('PCN_DATEPAIEMENT', 0);
      T2 := Tob.create('CP ACQUIS', Tob_Acquis, -1);
      T2.Dupliquer(T1, True, True, True);
      T1 := Tob_AcqEnCours.FindNext([''], [''], False);
    end;
  end;
  //FIN PT17
  if RecupereCongespris(Tob_Pris, Tob_Sal, DateF) then
  begin
    IntegreSolde(Tob_solde, Tob_pris);
    Suivant := suivant + 1; //PT17 Positionne l'ordre du prochain mvt à créer
    if not CalculValoCP(Tob_Pris, Tob_Acquis, P5T_Etab, TOB_SAL, ChpEntete.Ouvres, ChpEntete.Ouvrables, DateD, DateF, suivant, 'PRI', Tob_SAL.GetValeur(iPSA_SALARIE)) then
    begin
      ManqueAcquis := True; //PT88
      if Auto then {DEB PT-11}
        StMsgErr := '1 ou plusieurs mouvements de congés n''ont pu être payés par manque d''acquis pour le salarié ' + Salarie + ' ' + TOB_Sal.GetValeur(iPSA_LIBELLE) + ' ' +
          TOB_Sal.GetValeur(iPSA_PRENOM)
      else {FIN PT-11}
        HShowMessage('5;Calcul bulletin :;1 ou plusieurs mouvements de congés n''ont pu être payés par manque d''acquis;E;O;O;O;;;', '', '');
    end;
  end
  else Tob_Pris := nil;
  //DEB PT17  Réaffectation des modifs apportés à la tob des acquis en cours
  if Tob_AcqEnCours <> nil then
  begin
    T1 := Tob_AcqEnCours.FindFirst([''], [''], False);
    while T1 <> nil do
    begin
      T2 := Tob_Acquis.FindFirst([''], [''], False);
      while T2 <> nil do
      begin
        if T1.GetValue('PCN_ORDRE') = T2.GetValue('PCN_ORDRE') then
        begin
          JCPACQUISACONSOMME := JCPACQUISACONSOMME + T2.GetValue('PCN_APAYES');
          T1.free;
          T1 := Tob.Create('Cp Acq', Tob_AcqEnCours, -1);
          T1.Dupliquer(T2, True, True, True);
          T2.free;
        end;
        T2 := Tob_Acquis.FindNext([''], [''], False);
      end;
      T1 := Tob_AcqEnCours.FindNext([''], [''], False);
    end;
  end;
  //FIN PT17
  PAYECOUR := 0;
  // Est ce que le salarié a une date sortie renseignée et dans le bulletin ?
  DateSortie := Tob_Sal.GetValeur(iPSA_DATESORTIE);
  if ((DateSortie >= DateD) and (Datesortie <= DateF)) then
  begin
    { DEB PT90 }
    if (PGRecupereProfilCge(Etab) = '') then
    begin
      if Auto then
        StMsgErr := 'Génération du mouvement de solde impossible. Vous devez gérer un profil congés payés pour le salarié ' + Salarie + ' ' + TOB_Sal.GetValeur(iPSA_LIBELLE) + ' '
          + TOB_Sal.GetValeur(iPSA_PRENOM)
      else
        PgiBox('Vous devez gérer un profil congés payés pour le salarié ' + Salarie + ' ' + TOB_Sal.GetValeur(iPSA_LIBELLE) + ' ' + TOB_Sal.GetValeur(iPSA_PRENOM),
          'Génération du mouvement de solde impossible');
    end
      { FIN PT90 }
    else
    begin
      if Tob_pris <> nil then
      begin
        TP := Tob_pris.findfirst([''], [''], True);
        while tp <> nil do
        begin
          if TP.getvalue('PCN_MVTDUPLIQUE') <> 'X' then
            PrisBul := PrisBul + TP.getvalue('PCN_JOURS');
          TP := Tob_pris.findnext([''], [''], True);
        end;
      end;
      GenereSolde(Tob_sal, P5T_Etab, Tob_Pris, Tob_Acquis, Tob_Solde, dateF, DateD, suivant, JCPACQUIS, PrisBul, 'C');
      // SoldeAcquis(T_MvtAcquis,True); //PT17 Cas SLD on solde les acquis en cours au cas ou modifiées
      //PT46 Mise en commentaire on solde en validation en fonction du Consommé sur l'acquis en cour
      AjouteAcquiscourant(Tob_Acquis, Tob_sal, DateF, suivant);
    end;
  end;
  //      IntegrePaye(Tob_Rub, Tob_Solde,Etab,Salarie, DateD, DateF,'SLD' )  ;}

  IntegrePaye(Tob_Rub, Tob_Pris, Etab, Salarie, DateD, DateF, 'PRI', ManqueAcquis); //PT88 Ajout paramètre
  result := Tob_Pris;
end;

//modif mv 17/11/00

procedure SalEcritCP(var Tob_cppris: tob);
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  SupprimeMvtSolde(Tob_cppris);

  if Tob_cppris <> nil then
  begin
    TOB_cpPris.SetAllModifie(TRUE);
    TOB_cpPris.InsertOrUpdateDB(false);
  end;
  if T_MvtAcquis <> nil then //PT17 on met d'abord à jour la tob des acqs en cours
  begin //avant celle des acqs anterieurs
    T_MvtAcquis.SetAllModifie(TRUE);
    T_MvtAcquis.InsertorupdateDB(true);
  end;
  if TOB_Acquis <> nil then
  begin
    TOB_Acquis.SetAllModifie(TRUE);
    TOB_Acquis.InsertOrUpdateDB(true);
  end;
  if Tob_solde <> nil then
  begin
    Tob_Solde.SetAllModifie(TRUE);
    Tob_Solde.InsertOrUpdateDB(true);
  end;
  if Tob_AcquisAVirer <> nil then
  begin
    Tob_AcquisAVirer.deletedb(true);
  end;
  if T_MvtAcquisAVirer <> nil then
  begin
    T_MvtAcquisAVirer.deletedb(true);
    FreeAndNil(T_MvtAcquisAVirer); { PT97 13/05/2004 }
  end;
  if TOB_Delta <> nil then { PT146 }
  begin
    TOB_Delta.SetAllModifie(TRUE);
    TOB_Delta.InsertOrUpdateDB(true);
  end; { PT146 }


  {if T_MvtAcquis <> NIL then PT17 mise en commentaire maj au dessus
      begin
      T_MvtAcquis.SetAllModifie (TRUE);
      T_MvtAcquis.InsertorupdateDB(true);
      end;   }
end;

procedure LibereTobCP(var Tob_cppris: tob);
begin
  FreeAndNil(TOB_cpPris);
  FreeAndNil(TOB_Pris); { PT147 }
  FreeAndNil(TOB_Acquis);
  FreeAndNil(T_MvtAcquis);
  FreeAndNil(Tob_Solde);
  FreeAndNil(Tob_AcquisAVirer);
  FreeAndNil(T_MvtAcquisAVirer);

  FreeAndNil(TOB_delta); { PT146 }
end;

procedure SalEcritAbs(var Tob_abs: tob);
begin
  if assigned(Tob_abs) then { PT160-1 }
    if (Tob_abs.detail.count > 0) then // notVide(Tob_abs, true) then  { PT160 }
    begin
      TOB_abs.SetAllModifie(TRUE);
      TOB_abs.InsertOrUpdateDB(false);
    end;
  exit;
end;
//DEB PT71

procedure LibereTobAbs;
begin
  if Tob_ABS <> nil then
  begin
    Tob_ABS.free;
    Tob_ABS := nil;
  end;
  HABSPRIS := 0;
  JABSPRIS := 0;
end;
//FIN PT71
{ DEB PT158 }

procedure LibereTobCalendrier;
begin
  FreeAndNil(GblTob_Semaine); { DEB PT130 }
  FreeAndNil(GblTob_Standard);
  FreeAndNil(GblTob_JourFerie); { FIN PT130 }
end;
{ FIN PT158 }

procedure IntegreAbsenceDansPaye(Tob_Rub, Tob_ABS, Salarie: tob; const DateD, DateF: tdatetime; const Action: string);
var
  T, T1, Tob_LcAbs: tob;
  TypeAbs, Aliment, Arub, Natrub, profil: string;
  Compteur: double;
  i, j, k: integer;
  GereComm, Heure, GenereCom, NetH, HetJ: boolean;
  ProfilH, RubH, AlimH, ProfilJ, RubJ, AlimJ: string; { PT153 }
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;
  if TOB_Abs = nil then exit;
  Compteur := 0; { PT153 }
  { DEB PT131-1 Refonte traitement pour intégration pls motifs idem rubrique }
  Tob_LcAbs := Tob.Create('Les absences paies', nil, -1);

  { Suppression des lignes de commentaire existantes }
  for i := 0 to Tob_Abs.Detail.Count - 1 do
  begin
    T := Tob_Abs.Detail[i];
    TypeAbs := T.getvalue('PCN_TYPECONGE');
    { DEB PT153 }
    RechercheCarMotifAbsence(TypeAbs, ProfilH, RubH, AlimH, ProfilJ, RubJ, AlimJ, GereComm, Heure, NetH, HetJ);
    for k := 1 to 2 do
    begin
      GenereCom := False;
      if k = 1 then //Récupération de la rubrique en heures si renseigné
      begin
        Profil := ProfilH;
        Arub := RubH;
        Aliment := AlimH;
        Compteur := T.GetValue('PCN_HEURES');
        if Heure then GenereCom := True;
      end
      else
        if k = 2 then //Récupération de la rubrique en jours si renseigné
        begin
          Profil := ProfilJ;
          Arub := RubJ;
          Aliment := AlimJ;
          Compteur := T.GetValue('PCN_JOURS');
          if not Heure then GenereCom := True;
        end;
   { FIN PT153 }
      if Profil = '' then Natrub := 'AAA' else Natrub := ChercheNatRub(Salarie, Tob_rub, Profil, Arub);
      ChargeProfil(Salarie, Tob_rub, Profil);
      EnleveCommAbsence(Salarie, Tob_rub, Arub, Natrub, DateD, Datef);
      if Arub = '' then continue;
       { Conversion des absences en tob d'intégration rubrique }
      if Assigned(Tob_LcAbs) then
      begin
        T1 := Tob_LcAbs.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [Natrub, ARub], False);
        if Assigned(T1) then
        begin
          if Aliment = 'BAS' then T1.PutValue('PHB_BASEREM', T1.GetValue('PHB_BASEREM') + Compteur);
          if Aliment = 'TAU' then T1.PutValue('PHB_TAUXREM', T1.GetValue('PHB_TAUXREM') + Compteur); { PT143 }
          if Aliment = 'COE' then T1.PutValue('PHB_COEFFREM', T1.GetValue('PHB_COEFFREM') + Compteur); { PT143 }
          if Aliment = 'MON' then T1.PutValue('PHB_MTREM', T1.GetValue('PHB_MTREM') + Compteur); { PT143 }
        end
        else
        begin
          T1 := Tob.Create('Les absences paies', Tob_LcAbs, -1);
          T1.AddChampSupValeur('TYPETOB', 'AAA');
          T1.AddChampSupValeur('PHB_RUBRIQUE', Arub);
          T1.AddChampSupValeur('PHB_NATURERUB', Natrub);
          if Aliment = 'BAS' then T1.AddChampSupValeur('PHB_BASEREM', Compteur) else T1.AddChampSupValeur('PHB_BASEREM', 0);
          if Aliment = 'TAU' then T1.AddChampSupValeur('PHB_TAUXREM', Compteur) else T1.AddChampSupValeur('PHB_TAUXREM', 0);
          if Aliment = 'COE' then T1.AddChampSupValeur('PHB_COEFFREM', Compteur) else T1.AddChampSupValeur('PHB_COEFFREM', 0);
          if Aliment = 'MON' then T1.AddChampSupValeur('PHB_MTREM', Compteur) else T1.AddChampSupValeur('PHB_MTREM', 0);
        end;
         { Création des libellés à reprendre }
        if (GereComm) and (GenereCom) then { PT153 }
        begin
          T1 := Tob.Create('Les absences paies', Tob_LcAbs, -1);
          T1.AddChampSupValeur('TYPETOB', 'ZZZ');
          T1.AddChampSupValeur('PHB_RUBRIQUE', Arub);
          T1.AddChampSupValeur('PHB_NATURERUB', Natrub);
          T1.AddChampSupValeur('PCN_LIBELLE', T.getvalue('PCN_LIBELLE'));
        end;
        PositionnePaye(T, DateF);
      end;
    end;
  end; //End du for i

  { Parcours & création des lignes d'absences }
  if Assigned(Tob_LcAbs) then
  begin
    Tob_LcAbs.Detail.Sort('PHB_NATURERUB;PHB_RUBRIQUE;TYPETOB');
    for i := 0 to Tob_LcAbs.Detail.Count - 1 do
    begin
      T := Tob_LcAbs.Detail[i];
      if T.GetValue('TYPETOB') = 'AAA' then
      begin
        J := 0;
        T1 := IntegreAbsInRub(tob_rub, Salarie, T.GetValue('PHB_NATURERUB'), T.GetValue('PHB_RUBRIQUE'));
        if Assigned(T1) then
        begin
          T1.PutValue('PHB_BASEREM', T.GetValue('PHB_BASEREM'));
          T1.PutValue('PHB_TAUXREM', T.GetValue('PHB_TAUXREM'));
          T1.PutValue('PHB_COEFFREM', T.GetValue('PHB_COEFFREM'));
          T1.PutValue('PHB_MTREM', T.GetValue('PHB_MTREM'));
        end;
      end
      else
        if T.GetValue('TYPETOB') = 'ZZZ' then
        begin
          Inc(j);
          Ecritlignecomm(Tob_rub, Salarie, T, j, DateD, Datef, T.GetValue('PHB_RUBRIQUE'), T.GetValue('PHB_NATURERUB'));
        end;
    end;
    FreeAndNil(Tob_LcAbs);
  end;
  { FIN PT131-1 }
end;

function ChercheNatRub(Salarie, TPE: TOB; const Profil, rub: string): string;
var
  TPP, TPR: TOB;
  ThemeRub, Rubrique, ThemeExclus: string;
  I: Integer;
begin
  result := '';
  if Profil = '' then exit;
  ThemeRub := '';
  ThemeExclus := '';
  TPP := TOB_ProfilPaies.FindFirst(['PPI_PROFIL'], [Profil], FALSE); // $$$$
  if TPP <> nil then
  begin
    ThemeExclus := TPP.GetValue('PPI_THEMECOT'); // theme des rubriques de cotisations à exclure
    if ThemeExclus <> '' then SuppCotExclus(ThemeExclus, TPE);
    for I := 0 to TPP.Detail.count - 1 do
    begin
      TPR := TPP.Detail[I];
      Rubrique := TPR.GetValue('PPM_RUBRIQUE');
      if Rub = Rubrique then
      begin
        result := TPR.getvalue('PPM_NATURERUB');
        exit;
      end;
    end;
  end;
end;


procedure PositionnePaye(t: tob; const datef: tdatetime);
begin
  if T = nil then exit;
  T.putvalue('PCN_CODETAPE', 'P');
  T.putvalue('PCN_DATEPAIEMENT', datef);
  if VH_Paie.PGEcabMonoBase then T.putvalue('PCN_EXPORTOK', 'X'); { PT134-1 }
end;

procedure AnnuleAbsenceBulletin(const CodeSalarie: string; const DateF: tdatetime);
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  Executesql('UPDATE ABSENCESALARIE SET PCN_DATEPAIEMENT = "' + usdatetime(0) + '", ' +
    'PCN_CODETAPE = "..." WHERE PCN_SALARIE="' + CodeSalarie + //PT23
    '" AND PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" AND  PCN_TYPECONGE <> "PRI" AND ' + //PT54 Ajout AND PCN_SENSABS="-"
    ' PCN_DATEPAIEMENT = "' + usdatetime(Datef) + '"');
end;

procedure EnleveCommAbsence(Salarie, Tob_rub: tob; const Arub, Natrub: string; const dated, datef: tdatetime);
var
  t: tob;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;
  if ARub = '' then exit; { PT153 }

  // mv on vire les libelles de la tob_rub créés précédemment
  T := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    if ((copy(T.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.') and
      (T.GetValue('PHB_ORIGINELIGNE') = 'ABS')) then
      T.free;
    T := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;
{ DEB PT131-2 }

procedure EnleveRubAbsence(Salarie, Tob_rub: tob; const dated, datef: tdatetime);
var
  t: tob;
begin
  if BullCompl = 'X' then exit;
  T := Tob_Rub.FindFirst(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    ['ABS', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    T.free;
    T := Tob_Rub.FindNext(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      ['ABS', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;
{ FIN PT131-2 }

procedure Ecritlignecomm(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const Rub, NatRub: string);
var
  Thh, TR: tob;
begin
  TR := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
  THH := TOB.create('HISTOBULLETIN', Tob_Rub, -1);
  THH.PutValue('PHB_RUBRIQUE', Rub + '.' + IntToStr(i));
  THH.PutValue('PHB_NATURERUB', Natrub);
  THH.PutValue('PHB_DATEDEBUT', DateD);
  THH.PutValue('PHB_DATEFIN', DateF);
  THH.PutValue('PHB_ETABLISSEMENT', Salarie.GetValeur(iPSA_ETABLISSEMENT));
  THH.PutValue('PHB_SALARIE', Salarie.GetValeur(iPSA_SALARIE));
  THH.PutValue('PHB_BASEREM', 0);
  THH.PutValue('PHB_TAUXREM', 0);
  THH.PutValue('PHB_COEFFREM', 0);
  THH.PutValue('PHB_MTREM', 0);
  THH.PutValue('PHB_CONSERVATION', 'BUL');
  THH.PutValue('PHB_LIBELLE', T.GetValue('PCN_LIBELLE'));
  THH.PutValue('PHB_ORIGINELIGNE', 'ABS');
  THH.PutValue('PHB_IMPRIMABLE', 'X');
  THH.PutValue('PHB_TRAVAILN2', Salarie.GetValeur(iPSA_TRAVAILN2));
  THH.PutValue('PHB_TRAVAILN3', Salarie.GetValeur(iPSA_TRAVAILN3));
  THH.PutValue('PHB_TRAVAILN4', Salarie.GetValeur(iPSA_TRAVAILN4));
  THH.PutValue('PHB_TRAVAILN1', Salarie.GetValeur(iPSA_TRAVAILN1));
  THH.PutValue('PHB_CODESTAT', Salarie.GetValeur(iPSA_CODESTAT));
  THH.PutValue('PHB_LIBREPCMB1', Salarie.GetValeur(iPSA_LIBREPCMB1));
  THH.PutValue('PHB_LIBREPCMB2', Salarie.GetValeur(iPSA_LIBREPCMB2));
  THH.PutValue('PHB_LIBREPCMB3', Salarie.GetValeur(iPSA_LIBREPCMB3));
  THH.PutValue('PHB_LIBREPCMB4', Salarie.GetValeur(iPSA_LIBREPCMB4));
  THH.PutValue('PHB_CONFIDENTIEL', Salarie.GetValeur(iPSA_CONFIDENTIEL));
  if TR <> nil then
  begin
    THH.PutValue('PHB_ORDREETAT', TR.GetValue('PRM_ORDREETAT'));
    THH.PutValue('PHB_OMTSALARIAL', TR.GetValue('PRM_ORDREETAT')); { PT136 }
  end;
  //i := i + 1; PT141 Mise en commentaire

end;

procedure RechercheCarMotifAbsence(const Abs: string; var ProfilH, RubriqueH, AlimentH, ProfilJ, RubriqueJ, AlimentJ: string; var GereComm, Heure, NetH, NetJ: boolean);
var
  Q: tQuery;
  st: string;
begin
  Gerecomm := false;
  Heure := false;
  NetH := False; NetJ := False; { PT153 }
  AlimentH := ''; ProfilH := ''; RubriqueH := ''; { PT153 }
  AlimentJ := ''; ProfilJ := ''; RubriqueJ := ''; { PT153 }
  st := 'SELECT * FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFABSENCE = "' + Abs + '"';
  Q := opensql(st, True);
  if Q.eof then
  begin
    Ferme(Q);
    exit;
  end;
  { DEB PT153 }
  // DEB PT149
  Heure := ((Q.findField('PMA_JOURHEURE').asstring = 'HEU') or (Q.findField('PMA_JOURHEURE').asstring = 'HOR')); // PT199
  GereComm := (q.findfield('PMA_GERECOMM').asstring = 'X');
  AlimentH := Q.findfield('PMA_ALIMENT').asstring;
  ProfilH := Q.findfield('PMA_PROFILABS').asstring;
  RubriqueH := Q.findfield('PMA_RUBRIQUE').asstring;
  NetH := (Q.findfield('PMA_ALIMNETH').asstring = 'X');
  AlimentJ := Q.findfield('PMA_ALIMENTJ').asstring;
  ProfilJ := Q.findfield('PMA_PROFILABSJ').asstring;
  RubriqueJ := Q.findfield('PMA_RUBRIQUEJ').asstring;
  NetJ := (Q.findfield('PMA_ALIMNETJ').asstring = 'X');
  // FIN PT149
  { FIN PT153 }
  Ferme(Q);
end;


//Alimentation des valeurs dans les rubriques d'absence
// PT107 IJSS procedure IntegreRub(tob_rub, Salarie: tob; const TypeAbs, aliment: string; const DateD, DateF: tdatetime; const compteur: double; natrub, arub: string);

procedure IntegreRub(tob_rub, Salarie: tob; const TypeAbs, aliment: string; const DateD, DateF: tdatetime; const compteur: double; natrub, arub: string; IJSS, Maintien: Boolean);
var
  TRC, THB: tob;
begin
  if NatRub = 'AAA' then
  begin
    TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARub], FALSE); // $$$$
  end else
  begin
    TRC := Tob_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [ARub, NatRub], FALSE); // $$$$
    if TRC = nil then // Rubrique de base
    begin
      TRC := Tob_Bases.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [ARub, NatRub], FALSE); // $$$$
    end;
  end;

  THB := Tob_rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [NatRub, ARub], FALSE); // $$$$
  if THB = nil then
  begin
    THB := TOB.Create('HISTOBULLETIN', tob_rub, -1);
    if TRC <> nil then
      RemplirHistoBulletin(THB, Salarie, TRC, nil, tob_rub.GetValue('PPU_DATEDEBUT'), tob_rub.GetValue('PPU_DATEFIN')); //mv
  end;
  if THB <> nil then
  begin
    if Aliment = 'BAS' then
      THB.PutValue('PHB_BASEREM', Compteur);
    if Aliment = 'TAU' then
      THB.PutValue('PHB_TAUXREM', Compteur);
    if Aliment = 'COE' then
      THB.PutValue('PHB_COEFFREM', Compteur);
    if Aliment = 'MON' then
      THB.PutValue('PHB_MTREM', Compteur);
    {Positionne indicateur pour indiquer la provenance de la ligne dans le bulletin
    ET indiquer que la saisie n'est pas possible car les champs sont précalculés }
// d PT107 IJSS
    if not (IJSS) and not (maintien) then
      THB.PutValue('PHB_ORIGINELIGNE', 'ABS') // Calculé par les absences
    else
      if (IJSS) then
        THB.PutValue('PHB_ORIGINELIGNE', 'IJS') // Calculé par les IJSS
      else
        if (maintien) then
          THB.PutValue('PHB_ORIGINELIGNE', 'MAI'); // Calculé par le maintien
    // f PT107 IJSS

  end;
end;
{ DEB PT131-1 }

function IntegreAbsInRub(tob_rub, Tob_ParSal: TOB; const natrub, arub: string): Tob;
var
  TRC: tob;
begin
  if NatRub = 'AAA' then
  begin
    TRC := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [ARub], FALSE);
  end else
  begin
    TRC := Tob_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [ARub, NatRub], FALSE);
    if TRC = nil then // Rubrique de base
    begin
      TRC := Tob_Bases.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [ARub, NatRub], FALSE);
    end;
  end;

  Result := Tob_rub.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [NatRub, ARub], FALSE);
  if Result = nil then
  begin
    Result := TOB.Create('HISTOBULLETIN', tob_rub, -1);
    if TRC <> nil then
      RemplirHistoBulletin(Result, Tob_ParSal, TRC, nil, tob_rub.GetValue('PPU_DATEDEBUT'), tob_rub.GetValue('PPU_DATEFIN')); //mv
  end;
  if Assigned(Result) then Result.PutValue('PHB_ORIGINELIGNE', 'ABS');

end;
{ FIN PT131-1 }


procedure GridGriseCell(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas);
var
  R: TRect;
begin
  Canvas.Brush.color := $00C9C9CB; // $00E4E4E4;
  Canvas.Brush.Style := bsBDiagonal;
  Canvas.Pen.Color := Canvas.Brush.color;
  Canvas.Pen.mode := pmMask; //pmCopy;
  Canvas.Pen.Style := psDot; //psClear;
  Canvas.Pen.Width := 1;
  R := GS.CellRect(Acol, ARow);
  Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
end;

procedure GridDeGriseCell(const GS: THGrid; const Acol, Arow: integer; Canvas: Tcanvas);
var
  R: TRect;
begin
  Canvas.Brush.color := $00E4E4E4;
  Canvas.Brush.Style := bsClear;
  Canvas.Pen.Color := Canvas.Brush.color;
  Canvas.Pen.mode := pmMask; //pmCopy;
  Canvas.Pen.Style := psDot; //psClear;
  Canvas.Pen.Width := 1;
  R := GS.CellRect(Acol, ARow);
  Canvas.Rectangle(R.Left, R.Top, R.Right, R.Bottom);
end;
{ function qui calcule la base de précarité .
Identification des dates de début de fin de contrat tq qu'il n'ait pas eu de rupture
ou que le contrat n'indique pas un motif de non calcul de la précarité.
La précarité ne concerne que les contrats de type CDD.
3 cas où la précarité ne se calcule pas :
- refus CDI
- Faute grave
- Autre contrat CDD suivant le contrat
Méthodologie de recherche des contrats CDD :
- si DD lue > DF+5 stockee alors stockage DD et DF
- si DD lue <= DF stockee + 5 et DD lue >= DF stockée et si DF lue <= DD paie alors Stockage DF
- si CCD et Next non CCD alors Next et RAZ DD et DF Stockée
L'objectif est de trouver les dates de contrat CDD ayant pour la paie
Il suffit dès lors de calculer la base de precarite en sommant le cumul dans
la fourchette de dates calculees et de prendre aussi le cumul base de precarite
de la paie en cours
Refonte complète du calcul de la précarité.
Elle est tjrs calculée en fin de contrat même si un autre CDD suit.
Il faut dès lors recalculer les trentiemes pour la précarité déjà calculée afin de
la deduire sur la précarité en cours
}
// Reecriture calcul base de précarité

function RendBasePrecarite(const DateDebut, DateFin: TDateTime; const DuSalarie: string): double;
var
  Q: TQuery;
  St, TypeContrat, Precarite: string;
  DDLue, DFLue, DDOk, DFOk: TDateTime;
  MDDOk, MDFOk, LaDate: TDateTime;
  NbreTrent, RR, NbreC: Double;
  MtCDD: Double;
  NbreCDD, CtratHorsPaie, NbreCDDP, NBREP: Integer;
  Denom: Double;
  LaFin, DebCtr: TDateTime;
begin
  result := 0;
  DDLue := 0;
  DFLue := 0;
  DDOk := 0;
  DFOk := 0;
  MDDOk := 0;
  MDFOk := 0;
  NbreCDD := 0; // Nbre de CDD Finissant dans la paie
  NbreCDDP := 0;
  // Analyse du nombre de contrats CDD du salarié qui se termine dans la paie et qui commence avt la paie
  st := 'SELECT COUNT(*) NBRECDD FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie + '" AND PCI_TYPECONTRAT="CCD" AND ' +
    ' PCI_FINCONTRAT >="' + UsDateTime(DateDebut) + '" AND PCI_FINCONTRAT <="' + UsDateTime(DateFin) + '"' +
    ' AND PCI_DEBUTCONTRAT <"' + UsDateTime(DateDebut) + '"';
  Q := OpenSql(st, TRUE);
  if not Q.EOF then NbreCDD := Q.FindField('NBRECDD').AsInteger;
  Ferme(Q);
  // Analyse du nombre de contrats CDD du salarié dont les périodes sont incluses dans la paie
  st := 'SELECT COUNT(*) NBRECDD FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie + '" AND PCI_TYPECONTRAT="CCD" AND ' +
    ' PCI_FINCONTRAT >="' + UsDateTime(DateDebut) + '" AND PCI_FINCONTRAT <="' + UsDateTime(DateFin) + '"' +
    ' AND PCI_DEBUTCONTRAT >="' + UsDateTime(DateDebut) + '"';
  Q := OpenSql(st, TRUE);
  if not Q.EOF then NbreCDDP := Q.FindField('NBRECDD').AsInteger;
  Ferme(Q);
  // Analyse si contrat CDD qui commence dans la paie qui se termine au delà de la fin de paie
  CtratHorsPaie := 0;
  st := 'SELECT COUNT(*) NBRECDD,MIN (PCI_DEBUTCONTRAT) DEBCTR FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie + '" AND PCI_TYPECONTRAT="CCD" AND ' +
    ' PCI_DEBUTCONTRAT >="' + UsDateTime(DateDebut) + '" AND PCI_FINCONTRAT >"' + UsDateTime(DateFin) + '"';
  Q := OpenSql(st, TRUE);
  if not Q.EOF then
  begin
    CtratHorsPaie := Q.FindField('NBRECDD').AsInteger;
    DebCtr := Q.FindField('DEBCTR').AsDateTime;
  end
  else
  begin
    CtratHorsPaie := 0;
    DebCtr := 0;
  end;
  Ferme(Q);
  // Analyse des différents contrats de travail du salarié soit 1 (dans 99% ces cas) à 4 ou 5 enreg maxi
  // de façon à connaitre les dates debut et fin contrat
  st := 'SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_NONPRECARITE FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie +
    '" ORDER BY PCI_SALARIE,PCI_DEBUTCONTRAT DESC';
  Q := OpenSql(st, TRUE);
  while not Q.EOF do
  begin
    TypeContrat := Q.FindField('PCI_TYPECONTRAT').AsString;
    DDLue := Q.FindField('PCI_DEBUTCONTRAT').AsFloat;
    DFLue := Q.FindField('PCI_FINCONTRAT').AsFloat;
    Precarite := Q.FindField('PCI_NONPRECARITE').AsString;
    if (TypeContrat = 'CCD') and (DFLue >= DateDebut) and (DFLue <= DateFin) and (Precarite = '') then
    begin // Cas où 1er CDD trouvé ou nouveau CDD
      DDOk := DDlue;
      DFOk := DFLue;
      // memorisation du 1er contrat CDD lu dont la date de fin est dans la periode du bulletin     if MDDOk = 0 then MDDOk := DDLue;
      if MDDOk = 0 then MDDOk := DDLue;
      if MDFOk = 0 then MDFOk := DFLue;
    end;
    if (TypeContrat = 'CCD') and (DDLue <= DFOk) and (DDLue > DFOk) and (Precarite <> 'X') then
    begin // Cas CDD prolongé par un autre CDD
    end;
    if (TypeContrat <> 'CCD') and (DDOk <> 0) then
    begin // Cas où il y a un autre contrat qui ne soit pas un CDD
      if (TypeContrat = 'CDI') and (DDLue > DFOk) and (DDLue < DFOk) and (DDLue < DateDebut) then
      begin // cas CDD suivi d'un CDI dans les jours suivants
        DDOk := 0; // pas de calcul de precarite car embauche definitive
        DFOk := 0;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
  if DDOk <> MDDok then NbreC := 2 else NbreC := 1; // on identifie le nombre de contrats (> 1)
  if (DDOk <> 0) and (DFOk <> 0) then
  begin
    // cas 1 seul contrat
    if NbreC = 1 then
    begin
      if DDok >= DateDebut then // 1 Contrat dans la periode de paie
      begin
        if DFok = DateFin then
          result := RendCumulSalSess('18') // cumul base precarite de la paie en cours
        else
        begin // 1 contrat mais date de fin de paie ne correspond pas à la date de fin de contrat
          // ce cas ne devrait pas se produire !!!!
          Denom := DateFin + 1 - DateDebut; // Calcul du denominateur du trentieme
          if Denom > 30 then Denom := 30;
          NbreTrent := CalculTrentieme(DateDebut, DFok) / (Denom);
          if NbreTrent > 1 then NbreTrent := 1;
          result := (RendCumulSalSess('18') * NbreTrent); // cumul base precarite de la paie en cours proratisée
        end;
      end
      else
      begin // contrat débute sur 1 periode précédente et se termine dans la paie
        if DDok <= (DateDebut - 1) then
        begin
          st := 'SELECT COUNT (*) NBREP FROM PAIEENCOURS WHERE PPU_DATEDEBUT="' + UsDatetime(DDOk) + '" AND PPU_SALARIE="' + DuSalarie + '"';
          Q := OpenSql(St, TRUE);
          if not Q.EOF then NBREP := Q.FindField('NBREP').AsInteger
          else NBREP := 0;
          FERME(Q);
          if NbreP = 0 then LaDate := DebutDeMois(DDOk) // base precarite mois precedents
          else LaDate := DDOK;
          RR := ValCumulDate('18', LaDate, DateDebut - 1);
          St := 'SELECT SUM(PHB_BASEREM) AS MTPRECPAYES' +
            ' FROM HISTOBULLETIN' +
            ' LEFT JOIN REMUNERATION ON' +
            ' PHB_NATURERUB=PRM_NATURERUB AND' +
            ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
            ' PRM_INDEMLEGALES="003" AND' +
            ' PHB_SALARIE="' + DuSalarie + '" AND' +
            ' PHB_DATEDEBUT>="' + UsDateTime(LaDate) + '" AND' +
            ' PHB_DATEFIN<="' + UsDateTime(DateDebut - 1) + '"';
          Q := OpenSql(St, TRUE);
          if not Q.EOF then MtCDD := Q.FindField('MTPRECPAYES').AsFloat
          else MtCDD := 0;
          Ferme(Q);
          result := RR - MtCDD;
          // cas où on ne fait pas un bulletin à chaque fin de période de contrat
          if result < 0 then
          begin
            //            NbreTrent := 1;
            st := 'SELECT MAX(PCI_FINCONTRAT) DATECDD FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie +
              '" AND PCI_TYPECONTRAT="CCD" AND ' +
              ' PCI_FINCONTRAT >="' + UsDateTime(DebutDeMois(DDOK)) + '" AND PCI_FINCONTRAT <="' + UsDateTime(DDOK - 1) + '"';
            Q := OpenSql(St, TRUE);
            if not Q.EOF then LaFin := Q.FindField('DATECDD').AsDateTime
            else LaFin := idate1900;
            Ferme(Q);
            if LaFin = Idate1900 then LaFin := DDOK - 1;
            NbreTrent := CalculTrentieme(DebutDeMois(DDOK), LaFin) / 30;
            if NbreTrent > 1 then NbreTrent := 1;
            NbreTrent := 1 - NbreTrent;
            result := ARRONDI(ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOK)) * NbreTrent, 2);
          end;
        end;
        NbreTrent := 1;
        st := 'SELECT COUNT (*) NbreCDD FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie +
          '" AND PCI_TYPECONTRAT="CCD" AND ' +
          ' PCI_DEBUTCONTRAT >="' + UsDateTime(DFOK + 1) + '" AND PCI_DEBUTCONTRAT <="' + UsDateTime(FinDeMois(DFOK)) + '"';
        Q := OpenSql(st, TRUE);
        if not Q.EOF then NbreCDD := Q.FindField('NBRECDD').AsInteger else NbreCDD := 0;
        Ferme(Q);
        if (DateFin <> DFOK) and (NbreCDD > 0) then NbreTrent := CalculTrentieme(DateDebut, DFok) / 30;
        if NbreTrent > 1 then NbreTrent := 1;
        result := result + (RendCumulSalSess('18') * NbreTrent);
      end;
    end
    else // cas où 2 ou + contrats
    begin
      if (NbreCDD = 0) and (CtratHorsPaie = 0) then
      begin
        // il n'y a que des CDD dans la paie pas de chevauchement de contrat avec periode de paie
        // soit x contrats dans la pais
        NbreTrent := 1;
        if NbreCDDP > 0 then result := RendCumulSalSess('18');
      end
      else
      begin
        if (NbreCDD = 0) and (NbreCDDP >= 1) and (CtratHorsPaie > 0) and (Trentieme <> 30) then
          // il y a des CDD dans la paie pas de chevauchement de contrat avec periode de paie precedante
          // mais un CDD qui commence dans la paie et se poursuit en dehors
        begin
          NbreTrent := (FinDeMois(DebCtr) - DebCtr + 1) / 30; // Recupération du nbre de trentieme correspondant au contrat qui continue
          NbreTrent := 1 - (NbreTrent / Trentieme); // Proratisation des trentiémes calculés par rapport au nbre de trentieme de la paie
          result := (RendCumulSalSess('18') * NbreTrent);
        end
        else
        begin
          if ((NbreCDD = 1) and (NbreCDDP > 0)) then
          begin // Contrat(s) dans la paie, mais contrat hors paie(
            result := RendCumulSalSess('18');
          end
          else
          begin
            St := 'SELECT SUM(PHB_BASEREM) AS MTPRECPAYES' +
              ' FROM HISTOBULLETIN' +
              ' LEFT JOIN REMUNERATION ON' +
              ' PHB_NATURERUB=PRM_NATURERUB AND' +
              ' ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE WHERE' +
              ' PRM_INDEMLEGALES="003" AND' +
              ' PHB_SALARIE="' + DuSalarie + '" AND' +
              ' PHB_DATEDEBUT>="' + UsDateTime(DebutDeMois(DDOk)) + '" AND' +
              ' PHB_DATEFIN<="' + UsDateTime(FinDeMois(DDOK)) + '"';
            Q := OpenSql(St, TRUE);
            if not Q.EOF then MtCDD := Q.FindField('MTPRECPAYES').AsFloat
            else MtCDD := 0;
            Ferme(Q);
            result := ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOK)) - MtCDD;
            result := result + ValCumulDate('18', FinDeMois(DDOK) + 1, DateDebut - 1);
            if DateFin <> DFOK then NbreTrent := CalculTrentieme(DateDebut, DFok) / 30;
            if NbreTrent > 1 then NbreTrent := 1;
            result := result + (RendCumulSalSess('18') * NbreTrent);
          end;
        end;
      end;
    end;
  end; // fin si contrats trouves
end;

// PT49 : 18/06/2002 V582 PH Modif Refonte complète du calcul de la précarité

function RendBasePrecarite_Old(DateDebut, DateFin: TDateTime; DuSalarie: string): double;
var
  Q: TQuery;
  St, TypeContrat, Precarite: string;
  DDLue, DFLue, DDOk, DFOk: TDateTime;
  MDDOk, MDFOk, MF: TDateTime;
  NbreTrent, RR, NbreC: Double;
  RendRub: TRendRub;
begin
  result := 0;
  DDLue := 0;
  DFLue := 0;
  DDOk := 0;
  DFOk := 0;
  MDDOk := 0;
  MDFOk := 0;
  // Analyse des différents contrats de travail du salarié soit 1 (dans 99% ces cas) à 4 ou 5 enreg maxi
  st := 'SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT,PCI_NONPRECARITE FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + DuSalarie +
    '" ORDER BY PCI_SALARIE,PCI_DEBUTCONTRAT DESC';
  Q := OpenSql(st, TRUE);
  while not Q.EOF do
  begin
    TypeContrat := Q.FindField('PCI_TYPECONTRAT').AsString;
    DDLue := Q.FindField('PCI_DEBUTCONTRAT').AsFloat;
    DFLue := Q.FindField('PCI_FINCONTRAT').AsFloat;
    Precarite := Q.FindField('PCI_NONPRECARITE').AsString;
    // PT38 : 03/04/2002 V571 PH Modif Calcul base de précarité Seuil ramené à 1 jour
    // PT47 : 03/04/2002 V571 PH Modif Calcul base de précarité Seuil ramené à 0 jour
    if (TypeContrat = 'CCD') and (DFLue >= DateDebut) and (DFLue <= DateFin) and (Precarite = '') then
    begin // Cas où 1er CDD trouvé ou nouveau CDD
      DDOk := DDlue;
      DFOk := DFLue;
      // memorisation du 1er contrat CDD lu dont la date de fin est dans la periode du bulletin     if MDDOk = 0 then MDDOk := DDLue;
      if MDDOk = 0 then MDDOk := DDLue;
      if MDFOk = 0 then MDFOk := DFLue;
    end;
    if (TypeContrat = 'CCD') and (DDLue <= DFOk) and (DDLue > DFOk) and (DDLue > DFOk) and (Precarite <> 'X') then
    begin // Cas CDD prolongé par un autre CDD
      //     DFOk := DFLue; // on recupére la date de fin de contrat et on garde la date de debut de contrat
    end;
    if (TypeContrat <> 'CCD') and (DDOk <> 0) then
    begin // Cas où il y a un autre contrat qui ne soit pas un CDD
      if (TypeContrat = 'CDI') and (DDLue > DFOk) and (DDLue < DFOk) and (DDLue < DateDebut) then
      begin // cas CDD suivi d'un CDI dans les jours suivants
        DDOk := 0; // pas de calcul de precarite car embauche definitive
        DFOk := 0;
      end;
    end;
    Q.Next;
  end;
  Ferme(Q);
  if DDOk <> MDDok then NbreC := 2 else NbreC := 1; // on identifie le nombre de contrats (> 1)
  if (DDOk <> 0) and (DFOk <> 0) then
  begin
    // cas 1 seul contrat
    if NbreC = 1 then
    begin
      if DDok >= DateDebut then // Contrat dans la periode de paie
      begin
        if DFok = DateFin then
          result := RendCumulSalSess('18') // cumul base precarite de la paie en cours
        else
        begin
          {PT52-2
                    NbreTrent := CalculTrentieme(DateDebut,DFok)/30;
          }
          NbreTrent := CalculTrentieme(DateDebut, DFok) / (DateFin + 1 - DateDebut);
          //FIN PT52-2
          if NbreTrent > 1 then NbreTrent := 1;
          result := (RendCumulSalSess('18') * NbreTrent); // cumul base precarite de la paie en cours proratisée
        end;
      end
      else
      begin // contrat débute sur 1 periode précédente et se termine dans la paie
        //       result := ValCumulDate ('18',DDOk, DateDebut-1); // base precarite mois precedents
        MF := FindeMois(DDOk);
        //PT52-1 Ajout de la condition
        if (DDLue <> DDOK) then
        begin
          if MF = (DateDebut - 1) then // on a qu'un prorata sur le mois precedent
          begin
            {PT52-2
                         NbreTrent := CalculTrentieme(DDok, MF)/30;
                         RR := CalculTrentieme (DebutDeMois (DDok), DDok-1)/30;
                         if (RR + NbreTrent) <> 1 then NbreTrent := 1 - RR; // Correctif pour 30/30 dans le mois
            }
            NbreTrent := CalculTrentieme(DDok, MF) / (MF + 1 - DDLue);
            if DebutDeMois(DDok) > DDLue then
            begin
              RR := CalculTrentieme(DebutDeMois(DDok), DDok - 1) / 30;
              if (RR + NbreTrent) <> 1 then NbreTrent := 1 - RR; // Correctif pour 30/30 dans le mois
            end;
            //FIN PT52-2
            if NbreTrent > 1 then NbreTrent := 1;
            if (ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOk))) <> (ValCumulDate('18', DDOk, FinDeMois(DDOk))) then
            begin
              result := result + ValCumulDate('18', DDOk, DateDebut - 1);
            end
            else
            begin
              result := result + ValCumulDate('18', DebutDeMois(DDOk), DateDebut - 1);
              RendRub := ValRubDate('4282', 'REM', DebutDeMois(DDOk), DDOk - 1);
              result := result - RendRub.BasRem;
            end;
            //            result := ValCumulDate ('18',DebutDeMois(DDOk), DateDebut-1) * NbreTrent;
            result := Result + RendCumulSalSess('18'); // PT52-2
          end
          else
          begin // base precarite sur les mois complets + Base precarite proratisée
            result := ValCumulDate('18', MF + 1, DateDebut - 1); // base precarite mois precedents complets
            {PT52-2
                         NbreTrent := CalculTrentieme(DDok,MF)/30;
            }
            if ((MF + 1 - DDLue) > 30) then
              NbreTrent := CalculTrentieme(DDok, MF) / 30
            else
              NbreTrent := CalculTrentieme(DDok, MF) / (MF + 1 - DDLue);
            //FIN PT52-2
            if NbreTrent > 1 then NbreTrent := 1;
            {PT52-2
                         result := result + (ValCumulDate ('18',DDok, FinDeMois (DDOk)) * NbreTrent);
                         end;
                      end;
                   result := Result + RendCumulSalSess ('18'); // base precarite de la paie en cours
            }
            if (ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOk))) <> (ValCumulDate('18', DDOk, FinDeMois(DDOk))) then
            begin
              result := result + ValCumulDate('18', DDOk, FinDeMois(DDOk));
            end
            else
            begin
              result := result + ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOk));
              RendRub := ValRubDate('4282', 'REM', DebutDeMois(DDOk), DDOk - 1);
              result := result - RendRub.BasRem;
            end;
            result := Result + RendCumulSalSess('18'); // base precarite de la paie en cours
          end;
        end
        else
        begin
          if MF = (DateDebut - 1) then // on a qu'un prorata sur le mois precedent
          begin
            NbreTrent := CalculTrentieme(DDok, MF) / (MF + 1 - DDLue);
            if DebutDeMois(DDok) >= DDLue then
            begin
              RR := CalculTrentieme(DebutDeMois(DDok), DDok - 1) / 30;
              if (RR + NbreTrent) <> 1 then NbreTrent := 1 - RR; // Correctif pour 30/30 dans le mois
            end;
            if NbreTrent > 1 then NbreTrent := 1;
            result := ValCumulDate('18', DebutDeMois(DDOk), DateDebut - 1) * NbreTrent;

            NbreTrent := CalculTrentieme(DateDebut, DFok) / (DateFin + 1 - DateDebut);
            if NbreTrent > 1 then
              NbreTrent := 1;
            //             result := result + (ValCumulDate ('18',DateDebut, DateFin) * NbreTrent); // base precarite mois en cours
            result := Result + RendCumulSalSess('18'); // @@@@@@ Pour prendre en compte la base pécaritéen cours mais à payer
          end
          else
          begin
            result := ValCumulDate('18', MF + 1, DateDebut - 1); // base precarite mois precedents complets
            if ((MF + 1 - DDLue) > 30) then
              NbreTrent := CalculTrentieme(DDok, MF) / 30
            else
              NbreTrent := CalculTrentieme(DDok, MF) / (MF + 1 - DDLue);
            if NbreTrent > 1 then NbreTrent := 1;
            result := result + (ValCumulDate('18', DebutDeMois(DDOk), FinDeMois(DDOk)) * NbreTrent);
            result := Result + RendCumulSalSess('18'); // base precarite de la paie en cours
          end;
        end;
        //FIN PT52-2
        //FIN PT52-1
      end;
    end
    else // cas où 2 ou + contrats
    begin
      RR := 0;
      if MDFok = DateFin then result := RendCumulSalSess('18') // base précarite de la paie en cours
      else
      begin
        NbreTrent := ((MDFok + 1 - DateDebut) / 30);
        if NbreTrent > 1 then NbreTrent := 1;
        result := (RendCumulSalSess('18') * NbreTrent); // cumul base precarite de la paie en cours proratisée
      end;
      if MDDok < DateDebut then // // base precarite mois precedents
      begin // contrat débute sur 1 periode précédente et se termine dans la paie
        RR := ValCumulDate('18', MDDOk, DateDebut - 1); // base precarite mois precedents
        if RR = 0 then RR := ValCumulDate('18', DebutDeMois(MDDOk), DateDebut - 1);
      end;
      Result := result + RR;
    end;
  end; // fin si contrats trouves
end;
//------------------------------------------------------------------------------
// Procedure qui permet d'alimenter à chaque validation de bulletin la ligne de
// congés acquis correspondant
//------------------------------------------------------------------------------



function IsNumerique(const S: string): boolean;
const
  Numerique = ['0'..'9', ',', '.'];
var
  i: Integer;
begin
  if S = '' then
  begin
    IsNumerique := true;
    exit;
  end;
  IsNumerique := FALSE;
  if Length(S) < 1 then exit;
  for i := 1 to Length(S) do
    if ((not (S[i] in Numerique)) and (Ord(s[i]) <> 160)) then exit;
  IsNumerique := TRUE;
end;

function IsOkFloat(const val: string): boolean;
begin
  result := false;
  if not IsNumerique(Val) then
  begin
    PGIBox('Zone non numérique', 'Controle de numéricité');
    exit;
  end;
  result := true;
end;

function notVide(T: tob; const MulNiveau: boolean): boolean;
var
  tp: tob;
begin
  result := false;
  if t = nil then
    exit;
  tp := t.findfirst([''], [''], MulNiveau);
  result := (tp <> nil);
end;

procedure GestionCalculAbsence(tob_rub: tob; const DateD, DateF: tdatetime; const typem: string);
var
  mul, ValoDxmn: double;
  NBjSem: integer;
  Valo, MValoMs, PaieSalMaint: string;
  LcT_Etab: Tob; { PT150 }
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;
  //DEB PT57 Intégration de la méthode de valorisation des CP personnalisé salarié
  RendMethodeValoDixieme(NBjSem, MValoMs, ValoDxmn);
  if MValoMs = 'T' then //Calendrier théorique
  begin
    if NbJSem = 5 then Mul := 21.67 else Mul := 26;
  end
  else
  begin
    if MValoMs = 'M' then //Gestion manuelle
    begin
      Mul := ValoDxmn;
    end
    else
    begin //Calendrier réel
      if NbJSem = 5 then Mul := ChpEntete.Ouvres else
        Mul := ChpEntete.Ouvrables;
    end;
    //FIN PT57
  end; { PT127-2 Replacement du end; pour ne pas englober tout le code et permettre le calcul de l'indemnité et absence valo CP }
  { DEB PT150 }
  LcT_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Tob_Salarie.GetValeur(iPSA_ETABLISSEMENT)], False);
  Valo := TrouveModeValorisation(LcT_Etab, Tob_Salarie); //PT59
  PaieSalMaint := TrouvePaieSalMaintien(LcT_Etab, Tob_Salarie);
  { FIN PT150 }
  if TypeM = 'PRI' then
    if Assigned(Tob_Pris) then { PT160-1 }
      if (Tob_Pris.detail.count > 0) then //notVide(Tob_Pris, true) then { PT160 }
      begin
        if ReCalculeValoAbsence(Tob_pris, Valo, PaieSalMaint, Mul, 'PRI') then { PT150 }
        begin
          AgregeCumulPris(tob_pris);
          IntegrePaye(Tob_Rub, Tob_Pris, Tob_Salarie.GetValeur(iPSA_ETABLISSEMENT), CodSal, DateD, DateF, 'PRI');
        end;
      end;
  if TypeM = 'SLD' then
    if Assigned(Tob_Solde) then { PT160-1 }
      if (Tob_Solde.detail.count > 0) then // notVide(Tob_Solde, true) then  { PT160 }
      begin
        if ReCalculeValoAbsence(Tob_Solde, Valo, PaieSalMaint, Mul, 'SLD') then { PT150 }
        begin
          AgregeCumulPris(tob_solde);
          IntegrePaye(Tob_Rub, Tob_solde, Tob_Salarie.GetValeur(iPSA_ETABLISSEMENT), CodSal, DateD, DateF, 'SLD');
        end;
      end;
end;

procedure InitialiseVariableStatCP(AvecAbs : Boolean = TRUE);
begin
  HCPPRIS := 0; // Heures CP pris dans la session
  HCPPAYES := 0; // Heures CP payees dans la session
  HCPASOLDER := 0; // Heures CP à solder dans la session
  JCPACQUIS := 0; // Jours CP acquis globale dans la session
  JCPACQ := 0; // Jours CP acquis dans la session            { PT140 }
  JCPACS := 0; // Jours CP acquis suppl. dans la session     { PT140 }
  JCPACA := 0; // Jours CP acquis Ancienneté dans la session { PT140 }
  JCPACQUISACONSOMME := 0; // Valeur a consommé sur Jours CP acquis dans la session PT17
  JCPACQUISTHEORIQUE := 0; // Jours CP acquis théorique PT75-2
  MCPACQUIS := 0; // Mois acquis ds la session (trentieme)
  BCPACQUIS := 0; // Base acquis ds la session
  ASOLDERAVBULL := 0; // Nbre de jours à solder en cloture de contrat
  JRSAARRONDIR := 0; // CP de la période à solder hors acquis bulletin
  JCPPRIS := 0; // Jours CP pris dans la session
  JCPPAYES := 0; // Jours CP payés dans la session
  JCPPAYESPOSES := 0; // Jours CP pris à intégrer dans la session  PT88
  JCPASOLDER := 0; // Jours CP A Solder dans la session
  CPMTABS := 0; // CP Montant Absence
  CPMTINDABS := 0; // CP Montant Indemnite Absence
  CPMTINDCOMP := 0; // CP Montant Indemnite Compensatrice
  JCPSOLDE := 0;
  CPMTABSSOL := 0;
  CPMTINDABSSOL := 0;
  JPAVANTBULL := 0;
  TopJPAVANTBULL := false;
  Chargement := true;
  PP := True;
  ValoXP0N := 0;
  ValoXP0D := 0;
  // DEB PT191
  if AvecAbs then
  begin
    HABSPRIS := 0; //PT33 Init nouvel variable : Heures Absences pris
    JABSPRIS := 0; //PT33 Init nouvel variable : Jours Absences pris
  end;
  // FIN PT191
end;

{ DEB PT83-3 Extrait de InitialiseVariableStatCP }

procedure InitVarTHTRAVAILLES;
begin
  JTHTRAVAILLES := 0; // Jours théorique travaillés par le salarié sur la période
  HTHTRAVAILLES := 0; // Heures théoriques travaillés par le salarié sur la période //PT43
end;
{ FIN PT83-3 }

procedure ReaffecteDateMvtcp(const DateFin: tdatetime);
var
  i: integer;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if (BullCompl = 'X') or (actionBulcp <> tacreation) or (not assigned(tob_pris)) then
  else
  begin
    for i := 0 to tob_pris.detail.count - 1 do
      with tob_pris.detail[i] do
      begin
        if (getvalue('PCN_DATEFIN') > (10)) and (getvalue('PCN_CODETAPE') <> '...') then //PT23
          putvalue('PCN_DATEFIN', Datefin);
        putvalue('PCN_DATEPAIEMENT', Datefin);
      end;
  end;
end;
// PT32 : 07/02/2002 V571 PH suppression historisation salarié
{
procedure CopyHistosal(var A,N:Histosal);
begin
     A.salarie       := N.SALARIE;
     A.CodeEmploi    := N.CODEEMPLOI;
     A.LibelleEmploi := N.LIBELLEEMPLOI;
     A.Qualification := N.QUALIFICATION;
     A.Coefficient   := N.COEFFICIENT;
     A.Indice        := N.INDICE;
     A.Niveau        := N.NIVEAU;
     A.CodeStat      := N.CODESTAT;
     A.TravailN1     := N.TRAVAILN1;
     A.TravailN2     := N.TRAVAILN2;
     A.TravailN3     := N.TRAVAILN3;
     A.TravailN4     := N.TRAVAILN4;
     A.Groupepaie    := N.Groupepaie;
     A.SalaireMois1  := N.SalaireMois1;
     A.SalaireAnn1   := N.SalaireAnn1;
     A.SalaireMois2  := N.SalaireMois2;
     A.SalaireAnn2   := N.SalaireAnn2;
     A.SalaireMois3  := N.SalaireMois3;
     A.SalaireAnn3   := N.SalaireAnn3;
     A.SalaireMois4  := N.SalaireMois4;
     A.SalaireAnn4   := N.SalaireAnn4;
     A.DtLibre1      := N.Dtlibre1;
     A.Dtlibre2      := N.Dtlibre2;
     A.Dtlibre3      := N.Dtlibre3;
     A.Dtlibre4      := N.Dtlibre4;
     A.Boollibre1    := N.Boollibre1;
     A.Boollibre2    := N.Boollibre2;
     A.Boollibre3    := N.Boollibre3;
     A.Boollibre4    := N.Boollibre4;
     A.CSlibre1      := N.CSlibre1;
     A.CSlibre2      := N.CSlibre2;
     A.CSlibre3      := N.CSlibre3;
     A.CSlibre4      := N.CSlibre4;
     A.Profil        := N.Profil;
     A.PeriodBul     := N.PeriodBul;
//PT-12 02/10/01 V562 PH Rajout historisation salarie profil rémunération
     A.ProfilRem     := N.ProfilRem;
     A.Horairemois   := N.Horairemois;
     A.Horairehebdo  := N.Horairehebdo;
     A.HoraireAnnuel := N.HoraireAnnuel;
     A.TauxHoraire   := N.TauxHoraire;
//PT2
     A.DADSProfessio := N.DADSProfessio;
     A.DADSCategorie := N.DADSCategorie;
     A.TauxPartiel   := N.TauxPartiel;
     A.Condemploi    := N.Condemploi;//PT13
end;

procedure InitialiseBoolean(var M : Histosal;ValB: boolean);
Var
sValB : char;
begin
if Valb = true then
   sValB := 'X' else
   sValB := '-';
with M do begin
     Bsalarie      := sValb;
     BCodeEmploi   := sValb;
     BLibelleEmploi:= sValb;
     BQualification:= sValb;
     BCoefficient  := sValb;
     BIndice       := sValb;
     BNiveau       := sValb;
     BCodeStat     := sValb;
     BTravailN1    := sValb;
     BTravailN2    := sValb;
     BTravailN3    := sValb;
     BTravailN4    := sValb;
     BGroupepaie   := sValb;
     BSalaireMois1 := sValb;
     BSalaireAnn1  := sValb;
     BSalaireMois2 := sValb;
     BSalaireAnn2  := sValb;
     BSalaireMois3 := sValb;
     BSalaireAnn3  := sValb;
     BSalaireMois4 := sValb;
     BSalaireAnn4  := sValb;
     BDtlibre1     := sValb;
     BDtlibre2     := sValb;
     BDtlibre3     := sValb;
     BDtlibre4     := sValb;
     BBoollibre1   := sValb;
     BBoollibre2   := sValb;
     BBoollibre3   := sValb;
     BBoollibre4   := sValb;
     BCSlibre1     := sValb;
     BCSlibre2     := sValb;
     BCSlibre3     := sValb;
     BCSlibre4     := sValb;
//PT-12 02/10/01 V562 PH Rajout historisation salarie profil rémunération
     BProfilRem    := sValb;
     BPeriodBul    := sValb;
     BProfil       := sValb;
     BHorairemois  := sValb;
     BHorairehebdo := sValb;
     BHoraireAnnuel:= sValb;
     BTauxHoraire  := sValb;
//PT2
     BDADSProfessio:= sValb;
     BDADSCategorie:= sValb;
     BTauxPartiel  := sValb;
     BCondemploi   := sValb;//PT13

 end;
end;
}
// FIN PT32
// PT20 : 20/11/01 V562 PH Modif du calcul des dates d'édition automatique en fonction de l'établissement

procedure CalculeDateEdtPaie(var LDatD, LDatF: TDateTime; T_Etab: TOB);
var
  Annee, Mois, Jour, J1, J2: WORD;
  st: string;
begin
  if (T_ETAB.GetValue('ETB_JEDTDU') = 0) and (T_ETAB.GetValue('ETB_JEDTAU') = 0) then exit;
  if (T_ETAB.GetValue('ETB_JEDTDU') <> 0) and (T_ETAB.GetValue('ETB_JEDTAU') <> 0) then
  begin
    DecodeDate(LDatD, Annee, Mois, Jour);
    Jour := T_ETAB.GetValue('ETB_JEDTDU');
    // DEB PT108
    St := IntToStr(Jour) + DateSeparator + IntToStr(Mois) + DateSeparator + IntToStr(Annee);
    if not HEnt1.IsValidDate(st) then
    begin
      Jour := 1;
      LDatD := EncodeDate(Annee, Mois + 1, Jour);
    end
    else LDatD := EncodeDate(Annee, Mois, Jour);
    // FIN PT108
    J1 := T_ETAB.GetValue('ETB_JEDTDU');
    J2 := T_ETAB.GetValue('ETB_JEDTAU');
    if J1 = J2 then
    begin
      LDatF := LDatD; // On édite sur le même mois pour le même jour
      exit;
    end;
    LDatD := PLUSMOIS(LDatD, -1); // Jour sur le mois précédant
    DecodeDate(LDatF, Annee, Mois, Jour);
    Jour := T_ETAB.GetValue('ETB_JEDTAU');
    St := IntToStr(Jour) + DateSeparator + IntToStr(Mois) + DateSeparator + IntToStr(Annee);
    if not HEnt1.IsValidDate(st) then
    begin
      Jour := 1;
      LDatF := EncodeDate(Annee, Mois, Jour);
      LDatF := FINDEMOIS(LDatF);
      exit;
    end;
    LDatF := EncodeDate(Annee, Mois, Jour);
  end;
end;

// PT28 : 27/12/2001 V571 PH Calcul des variables de tupe cumul en tenant compte du mois de RAZ
{ Fonction qui va borner la date de début en fonction du mois de RAZ du cumul.
Elle essaie de renseigner le mois de début et l'année en fonction des dates calculées par rapport
à l'exercice social
Si erreur alors pas de calcul
PT70    23/01/2003 PH V591 Correction du calcul des bornes de dates pour la valorisation d'un cumul antérieur avec une période de raz
Rajout dates début et fin
}

procedure BorneDateCumul(const DateDeb, DateFin: TDateTime; var ZDatD, ZDatF: TDateTime; const PerRaz: string);
var
  AA, MM, JJ: WORD;
  T_Etab: TOB;
  Etab, LeMoisRaz: string;
  LaDateCP: TDateTime;
begin
  if PerRaz = '' then
  begin
    ZDatD := -1;
    ZDatF := -1;
  end
  else if PerRaz = '99' then ZdatD := IDate1900 // Jamais de RAZ
  else if (PerRaz >= '01') and (PerRaz <= '12') then // Raz par rapport à un mois
    RendBorneDateCumul(DateDeb, DateFin, ZDatD, ZDatF, PerRaz)
  else if PerRaz = '20' then // Debut de période CP
  begin
    Etab := string(TOB_Salarie.GetValeur(iPSA_ETABLISSEMENT));
    T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Etab], True);
    if T_Etab <> nil then
    begin
      LaDateCp := T_Etab.getvalue('ETB_DATECLOTURECPN');
      DecodeDate(LaDateCp, AA, MM, JJ);
      LeMoisRaz := IntToStr(MM);
      RendBorneDateCumul(DateDeb, DateFin, ZDatD, ZDatF, LeMoisRaz);
    end
    else
    begin
      ZDatD := -1;
      ZDatF := -1;
    end;
  end;
end;
{Fonction qui renseigne le mois de début et l'année
PT70    23/01/2003 PH V591 Correction du calcul des bornes de dates pour la valorisation d'un cumul antérieur avec une période de raz
Rajout Datedebut et fin et réécriture de la fonction
}

procedure RendBorneDateCumul(const DateDeb, DateFin: TDateTime; var ZDatD, ZdatF: TDateTime; const PerRaz: string);
var
  AA, MM, JJ: WORD;
  A, M, J: WORD;
begin
  DecodeDate(DateFin, AA, MM, JJ);
  if StrToInt(PerRaz) + 1 = MM then // Periode de RAZ = 31/05 donc 06 on ne recupère rien PT201
  begin
    ZDatD := -1;
    ZdatF := -1;
    exit;
  end; // Cas ou le mois en cours correspond au mois de raz donc resultat = 0
  ZdatF := DateDeb - 1; // On se positionne juste avant la date de debut de la paie
  ZdatD := DebutDeMois(PLUSMois(ZDatF, -12)); // debut de mois - 1 an complet
  DeCodeDate(ZDatD, AA, MM, JJ);
  DeCodeDate(DateDeb, A, M, J);
  MM := StrToInt(PerRaz);
  // DEB PT201
  if MM < M then AA := A
  else
  begin
    MM := MM + 1;
    AA := AA + 1;
  end;
  // FIN PT201
  ZDatD := EncodeDate(AA, MM, 01);
end;

// PT35 : 26/03/2002 V571 PH Fonction de raz de ChptEntete

procedure RazChptEntete(const DateF: TDateTime);
begin
  //PT68    19/12/2002 PH V591 Affectation de la date de paiement à la date de fin de paie si non renseignée
  //  Champ DatePai et DateVal renseigné
  with ChpEntete do
  begin
    Reglt := '';
    DatePai := DateF;
    DateVal := Date;
    HorMod := FALSE;
    BasesMod := FALSE;
    TranchesMod := FALSE;
    TrentMod := FALSE;
    DTrent := 0;
    NTrent := 0;
    Ouvres := 0;
    Houvres := 0; //PT61 Init Houvres à 0
    Ouvrables := 0;
    HOuvrables := 0; //PT61 Init Houvrables à 0
    HeuresTrav := 0;
    RegltMod := FALSE;
    Edtdu := 0;
    Edtau := 0;
    CpAcquisMod := False; { PT97 }
  end;
end;
// FIN PT35
// PT60 07/10/2002 V585 PH Recherche d'un element national par préférence DOS,STD,CEG

function ValEltNatPref(const Elt, Pref: string; const DatTrait: TDatetime): double;
var
  ret: double;
  NoDossier: string;
  EltNat: TOB;
  Convention:String;
begin
  result := -123456; // initialisation
  ret := -123456;
  if Pref = 'CEG' then EltNat := TOB_EltNationauxCEG
  else if Pref = 'STD' then EltNat := TOB_EltNationauxSTD
  else EltNat := TOB_EltNationauxDOS;
  if assigned(EltNat) and (EltNat.detail.Count > 0) then
  begin
    Nodossier := '000000';
      // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if Pref = 'DOS' then Nodossier := PgRendNoDossier();
    if (iCHampMontant = 0) then MemorisePel(EltNat.Detail[0]);
    //DEB PT172
    // Si on est STD, tester d'abord en fonction de la convention collective  //PT172
    if (Pref = 'STD') or (Pref = 'CEG') then       //PT188
      Convention := TOB_Salarie.GetValeur(iPSA_CONVENTION)
    else
      Convention := '000';

    Ret := RecupValStd(EltNat,Convention,Pref,Elt,DatTrait);

    // Si on n'a pas trouvé la convention du salarié, regarder toutes conventions
    if (ret = -123456) and (Pref = 'STD') and (Convention <> '000') then
    begin
      Convention := '000';
      Ret := RecupValStd(EltNat,Convention,Pref,Elt,DatTrait);
    end;
    //FIN PT172

    //DEB PT188
    if (ret = -123456) and (Pref = 'CEG') and (Convention <> '000') then
    begin
      Convention := '000';
      Ret := RecupValStd(EltNat,Convention,Pref,Elt,DatTrait);
    end;
    //FIN PT188

    result := ret;
  end;
end;

//DEB PT172
function RecupValStd(EltNat:Tob;Convention:String;Pref:String;Elt:String;DatTrait: TDatetime) : Double;
var
  T: TOB;
  FindNum: integer;
  i: integer;
  datelue, LaDate: TDateTime;
  LeDecalage: string;
begin
  result := -123456; // initialisation
  i := 0;
  FindNum := -1;
  while i < EltNat.Detail.Count do
  begin
    with EltNat.Detail[i] do
    begin
      if (GetValeur(iChampPredefini) = Pref)
      // and (GetValeur(iCHampDossier) = NoDossier)
      and (GetValeur(iCHampCodeElt) = Elt)
      and (GetValeur(iChampConvention) = Convention) then //PT172
      begin
        FindNum := i;
        break;
      end;
    end;
    Inc(i);
  end;

  if FindNum > -1 then
  begin
    while true do
    begin
      t := EltNat.Detail[FindNum];
      DateLue := T.GetValeur(iCHampDateValidite);
      // PT62 07/11/2002 PH V591 Prise en compte du decalage de paie donc le traitement depend de l'élément national
      LeDecalage := T.GetValeur(iCHampDecaleMois);
      LaDate := DatTrait;
      if (VH_Paie.PGDecalage = True) and (LeDecalage = 'X') then LaDate := PLUSMOIS(DatTrait, 1);
      if (VH_Paie.PGDecalagePetit = True) and (LeDecalage = 'X') then LaDate := PLUSMOIS(DatTrait, 1);
      if DateLue <= LaDate then // fin PT62
      begin
        // PT100  : 05/04/2004 PH V_50 FQ 11231 Prise en compte du cas spécifique Alsace Lorraine
        //          if VH_Paie.PGTenueEuro = FALSE then ret := T.GetValeur(iCHampMontant)
        if (RegimeAlsace = 'X') and (T.GetValeur(iChampRegimeAlsace) = 'X') then result := T.GetValeur(iCHampMontant)
        else result := T.GetValeur(iCHampMontantEuro);
      end
      else break;
      Inc(FindNum);
      if FindNum = EltNat.Detail.Count then break;
      with EltNat.Detail[FindNum] do
        if (GetValeur(iChampPredefini) <> Pref)
        // or (GetValeur(iCHampDossier) <> NoDossier)
        or (GetValeur(iCHampCodeElt) <> Elt)
        or (GetValeur(iChampConvention) <> Convention) then break; //PT172
    end;
  end;
end;

//Fonction spécifique pour les éléments dynamiques dossier
function ValEltNatDynDos(const Elt, Pref: string; const DatTrait: TDatetime): double;
var
  ret: double;
  NoDossier: string;
  EltNat: TOB;
  T: TOB;
  FindNum: integer;
  i: integer;
  datelue, LaDate: TDateTime;
  LeDecalage: string;
begin
  result := -123456; // initialisation
  ret := -123456;
  if Pref = 'SAL' then
    EltNat := TOB_EltDynSal
  else if Pref = 'POP' then
    EltNat := TOB_EltDynPop
  else
    EltNat := TOB_EltDynEtab;
  if assigned(EltNat) and (EltNat.detail.Count > 0) then
  begin
    if (iCHampPEDMontant = 0) then MemorisePed(EltNat.Detail[0]);
    i := 0;
    FindNum := -1;
    while i < EltNat.Detail.Count do
    begin
      with EltNat.Detail[i] do
      begin
        if (GetValeur(iCHampPEDCodeElt) = Elt) then
        begin
          FindNum := i;
          break;
        end;
      end;
      Inc(i);
    end;

    if FindNum > -1 then
    begin
      while true do
      begin
        t := EltNat.Detail[FindNum];
        DateLue := T.GetValeur(iCHampPEDDateValidite);
        // PT62 07/11/2002 PH V591 Prise en compte du decalage de paie donc le traitement depend de l'élément national
        LeDecalage := T.GetValeur(iCHampPEDDecaleMois);
        LaDate := DatTrait;
        if (VH_Paie.PGDecalage = True) and (LeDecalage = 'X') then LaDate := PLUSMOIS(DatTrait, 1);
        if (VH_Paie.PGDecalagePetit = True) and (LeDecalage = 'X') then LaDate := PLUSMOIS(DatTrait, 1);
        if DateLue <= LaDate then // fin PT62
        begin
        // Zone Régime Alsace présente mais non gérée dans les éléments dynamiques dossier
        //  if (RegimeAlsace = 'X') and (T.GetValeur(iChampPEDRegimeAlsace) = 'X') then result := T.GetValeur(iCHampPEDMontant)
          ret := T.GetValeur(iCHampPEDMontantEuro);
        end
        else break;
        Inc(FindNum);
        if FindNum = EltNat.Detail.Count then break;
        with EltNat.Detail[FindNum] do
          if (GetValeur(iCHampPEDCodeElt) <> Elt) then break;
      end;
    end;
    result := ret;
  end;
end;
//FIN PT172

Procedure InitMemorise();
Begin
  iPHC_ETABLISSEMENT    := 0;         iPHC_MONTANT          := 0;
  iPHC_SALARIE          := 0;         iPHC_TRAVAILN1        := 0;
  iPHC_DATEDEBUT        := 0;         iPHC_TRAVAILN2        := 0;
  iPHC_DATEFIN          := 0;         iPHC_TRAVAILN3        := 0;
  iPHC_REPRISE          := 0;         iPHC_TRAVAILN4        := 0;
  iPHC_CUMULPAIE        := 0;         iPHC_CODESTAT         := 0;
  iPHB_ETABLISSEMENT    := 0;         iPHC_CONFIDENTIEL     := 0;
  iPHB_SALARIE          := 0;         iPHC_LIBREPCMB1       := 0;
  iPHB_DATEDEBUT        := 0;         iPHC_LIBREPCMB2       := 0;
  iPHB_DATEFIN          := 0;         iPHC_LIBREPCMB3       := 0;
  iPHB_NATURERUB        := 0;         iPHC_LIBREPCMB4       := 0;
  iPHB_TAUXAT           := 0;         iPHB_RUBRIQUE         := 0;
  iPHB_PLAFOND          := 0;         iPHB_LIBELLE          := 0;
  iPHB_PLAFOND1         := 0;         iPHB_IMPRIMABLE       := 0;
  iPHB_PLAFOND2         := 0;         iPHB_TAUXREM          := 0;
  iPHB_PLAFOND3         := 0;         iPHB_COEFFREM         := 0;
  iPHB_CONSERVATION     := 0;         iPHB_TAUXSALARIAL     := 0;
  iPHB_ORDREETAT        := 0;         iPHB_TAUXPATRONAL     := 0;
  iPHB_SENSBUL          := 0;         iPHB_BASEREMIMPRIM    := 0;
  iPHB_ORIGINELIGNE     := 0;         iPHB_TAUXREMIMPRIM    := 0;
  iPHB_ORIGINEINFO      := 0;         iPHB_COEFFREMIMPRIM   := 0;
  iPHB_TRAVAILN1        := 0;         iPHB_BASECOTIMPRIM    := 0;
  iPHB_TRAVAILN2        := 0;         iPHB_TAUXSALIMPRIM    := 0;
  iPHB_TRAVAILN3        := 0;         iPHB_TAUXPATIMPRIM    := 0;
  iPHB_TRAVAILN4        := 0;         iPHB_ORGANISME        := 0;
  iPHB_CODESTAT         := 0;         iPHB_BASECOT          := 0;
  iPHB_CONFIDENTIEL     := 0;         iPHB_MTSALARIAL       := 0;
  iPHB_LIBREPCMB1       := 0;         iPHB_MTPATRONAL       := 0;
  iPHB_LIBREPCMB2       := 0;         iPHB_TRANCHE1         := 0;
  iPHB_LIBREPCMB3       := 0;         iPHB_TRANCHE2         := 0;
  iPHB_LIBREPCMB4       := 0;         iPHB_TRANCHE3         := 0;
  iPHB_BASEREM          := 0;         iPHB_ORDREETAT        := 0;
  iPHB_MTREM            := 0;         iPHB_COTREGUL         := 0;
  iPHB_OMTSALARIAL      := 0;
  iPCL_ALIMCUMUL        := 0;         iPCL_ALIMCUMULCOT     := 0;
  iPCL_COEFFAFFECT      := 0;
  iPCR_CUMULPAIE        := 0;         iPCR_SENS             := 0;
  iPCT_IMPRIMABLE       := 0;         iPCT_DU               := 0;
  iPCT_BASEIMP          := 0;         iPCT_AU               := 0;
  iPCT_TXSALIMP         := 0;         iPCT_MOIS1            := 0;
  iPCT_ORGANISME        := 0;         iPCT_MOIS2            := 0;
  iPCT_ORDREETAT        := 0;         iPCT_MOIS3            := 0;
  iPCT_LIBELLE          := 0;         iPCT_MOIS4            := 0;
  iPCT_FFPAT            := 0;         iPCT_DECBASE          := 0;
  iPCT_DECMTPAT         := 0;         iPCT_TYPEBASE         := 0;
  iPCT_TYPEMINISAL      := 0;         iPCT_TYPETAUXSAL      := 0;
  iPCT_VALEURMINISAL    := 0;         iPCT_TAUXSAL          := 0;
  iPCT_DECMTSAL         := 0;         iPCT_DECTXSAL         := 0;
  iPCT_TYPEMAXISAL      := 0;         iPCT_TYPETAUXPAT      := 0;
  iPCT_VALEURMAXISAL    := 0;         iPCT_TAUXPAT          := 0;
  iPCT_TYPEMINIPAT      := 0;         iPCT_DECTXPAT         := 0;
  iPCT_VALEURMINIPAT    := 0;         iPCT_TYPEFFSAL        := 0;
  iPCT_DECMTPAT         := 0;         iPCT_FFSAL            := 0;
  iPCT_TYPEMAXIPAT      := 0;         iPCT_DECMTSAL         := 0;
  iPCT_VALEURMAXIPAT    := 0;         iPCT_TYPEFFPAT        := 0;
  iPCT_DECMTPAT         := 0;         iPCT_TYPEREGUL        := 0;
  iPCT_DECBASECOT       := 0;         iPCT_TXPATIMP         := 0;
  iPCT_SOUMISREGUL      := 0;         iPCT_TRANCHE1         := 0;
  iPCT_BASECOTISATION   := 0;         iPCT_TRANCHE2         := 0;
  iPCT_TYPEPLAFOND      := 0;         iPCT_TRANCHE3         := 0;
  iPCT_PLAFOND          := 0;         iPCT_ORDREAT          := 0;
  iPCT_TYPETRANCHE1     := 0;         iPCT_CODETRANCHE      := 0;
  iPCT_TYPETRANCHE2     := 0;         iPCT_PRESFINMOIS      := 0;
  iPCT_TYPETRANCHE3     := 0;
  iPRM_IMPRIMABLE       := 0;         iPRM_SENSBUL          := 0;
  iPRM_BASEIMPRIMABLE   := 0;         iPRM_LIBELLE          := 0;
  iPRM_BASEIMPRIMABLE   := 0;         iPRM_THEMEREM         := 0;
  iPRM_TAUXIMPRIMABLE   := 0;         iPRM_DECBASE          := 0;
  iPRM_COEFFIMPRIM      := 0;         iPRM_DECTAUX          := 0;
  iPRM_ORDREETAT        := 0;         iPRM_DECCOEFF         := 0;
  iPRM_LIBCONTRAT       := 0;         iPRM_DECMONTANT       := 0;
  iChampPredefini       := 0;         iCHampPEDDateValidite := 0;
  iCHampDossier         := 0;         iCHampPEDDecaleMois   := 0;
  iCHampCodeElt         := 0;         iCHampPEDMontant      := 0;
  iCHampDateValidite    := 0;         iCHampPEDMontantEuro  := 0;
  iCHampDecaleMois      := 0;         iChampPEDRegimeAlsace := 0;
  iCHampMontant         := 0;         iChampPEDValeurNiveau := 0;
  iCHampMontantEuro     := 0;         iCHampPNRCodeElt      := 0;
  iChampRegimeAlsace    := 0;         iCHampPNRNiveauRequis := 0;
  iChampConvention      := 0;         iCHampPNRNiveauMaxi   := 0;
  iCHampPEDCodeElt      := 0;
  iPSA_SALARIE          := 0;         iPSA_CONFIDENTIEL     := 0;
  iPSA_ETABLISSEMENT    := 0;         iPSA_TAUXHORAIRE      := 0;
  iPSA_TRAVAILN1        := 0;         iPSA_SALAIREMOIS1     := 0;
  iPSA_TRAVAILN2        := 0;         iPSA_SALAIREMOIS2     := 0;
  iPSA_TRAVAILN3        := 0;         iPSA_SALAIREMOIS3     := 0;
  iPSA_TRAVAILN4        := 0;         iPSA_SALAIREMOIS4     := 0;
  iPSA_CODESTAT         := 0;         iPSA_SALAIREMOIS5     := 0;
  iPSA_LIBREPCMB1       := 0;         iPSA_SALAIRANN1       := 0;
  iPSA_LIBREPCMB2       := 0;         iPSA_SALAIRANN2       := 0;
  iPSA_LIBREPCMB3       := 0;         iPSA_SALAIRANN3       := 0;
  iPSA_LIBREPCMB4       := 0;         iPSA_SALAIRANN4       := 0;
  iPSA_TYPPROFILRBS     := 0;         iPSA_SALAIRANN5       := 0;
  iPSA_PROFILRBS        := 0;         iPSA_DATENAISSANCE    := 0;
  iPSA_TYPREDREPAS      := 0;         iPSA_DATEENTREE       := 0;
  iPSA_REDREPAS         := 0;         iPSA_DATESORTIE       := 0;
  iPSA_TYPREDRTT1       := 0;         iPSA_TYPPROFILREM     := 0;
  iPSA_REDRTT1          := 0;         iPSA_PROFILREM        := 0;
  iPSA_TYPREDRTT2       := 0;         iPSA_TYPPROFIL        := 0;
  iPSA_REDRTT2          := 0;         iPSA_PROFIL           := 0;
  iPSA_PROFILTPS        := 0;         iPSA_TYPPERIODEBUL    := 0;
  iPSA_TYPPROFILAFP     := 0;         iPSA_PERIODBUL        := 0;
  iPSA_PROFILAFP        := 0;         iPSA_PROFILTSS        := 0;
  iPSA_TYPPROFILAPP     := 0;         iPSA_TYPPROFILCGE     := 0;
  iPSA_PROFILAPP        := 0;         iPSA_PROFILCGE        := 0;
  iPSA_TYPPROFILRET     := 0;         iPSA_PROFILCDD        := 0;
  iPSA_PROFILRET        := 0;         iPSA_PROFILMUL        := 0;
  iPSA_TYPPROFILMUT     := 0;         iPSA_TYPPROFILFNAL    := 0;
  iPSA_PROFILMUT        := 0;         iPSA_PROFILFNAL       := 0;
  iPSA_TYPPROFILPRE     := 0;         iPSA_TYPPROFILTRANS   := 0;
  iPSA_PROFILPRE        := 0;         iPSA_PROFILTRANS      := 0;
  iPSA_TYPPROFILTSS     := 0;         iPSA_TYPPROFILANC     := 0;
  iPSA_DATEANCIENNETE   := 0;         iPSA_PROFILANCIEN     := 0;
  iPSA_QUALIFICATION    := 0;         iPSA_LIBELLE          := 0;
  iPSA_COEFFICIENT      := 0;         iPSA_PRENOM           := 0;
  iPSA_LIBELLEEMPLOI    := 0;         iPSA_NUMEROSS         := 0;
  iPSA_CIVILITE         := 0;         iPSA_ADRESSE1         := 0;
  iPSA_CPACQUISMOIS     := 0;         iPSA_ADRESSE2         := 0;
  iPSA_NBREACQUISCP     := 0;         iPSA_ADRESSE3         := 0;
  iPSA_TYPDATPAIEMENT   := 0;         iPSA_CODEPOSTAL       := 0;
  iPSA_MOISPAIEMENT     := 0;         iPSA_VILLE            := 0;
  iPSA_JOURPAIEMENT     := 0;         iPSA_INDICE           := 0;
  iPSA_TYPREGLT         := 0;         iPSA_NIVEAU           := 0;
  iPSA_PGMODEREGLE      := 0;         iPSA_CONVENTION       := 0;
  iPSA_REGULANCIEN      := 0;         iPSA_CODEEMPLOI       := 0;
  iPSA_HORHEBDO         := 0;         iPSA_AUXILIAIRE       := 0;
  iPSA_HORAIREMOIS      := 0;         iPSA_BOOLLIBRE1       := 0;
  iPSA_HORANNUEL        := 0;         iPSA_BOOLLIBRE2       := 0;
  iPSA_PERSACHARGE      := 0;         iPSA_BOOLLIBRE3       := 0;
  iPSA_PCTFRAISPROF     := 0;         iPSA_BOOLLIBRE4       := 0;
  iPSA_MULTIEMPLOY      := 0;         iPSA_DADSPROF         := 0;
  iPSA_SALAIREMULTI     := 0;         iPSA_DADSCAT          := 0;
  iPSA_ORDREAT          := 0;         iPSA_TAUXPARTIEL      := 0;
  iPSA_SALAIRETHEO      := 0;         iPSA_CPTYPEMETHOD     := 0;
  iPSA_DATELIBRE1       := 0;         iPSA_VALORINDEMCP     := 0;
  iPSA_DATELIBRE2       := 0;         iPSA_CPTYPEVALO       := 0;
  iPSA_DATELIBRE3       := 0;         iPSA_MVALOMS          := 0;
  iPSA_DATELIBRE4       := 0;         iPSA_VALODXMN         := 0;
  iPSA_VALANCCP         := 0;         iPSA_CPACQUISANC      := 0;
  iPSA_ANCIENNETE       := 0;         iPSA_BASANCCP         := 0;
  iPSA_CALENDRIER       := 0;         iPSA_VALANCCP         := 0;
  iPSA_STANDCALEND      := 0;         iPSA_DATANC           := 0;
  iPSA_SORTIEDEFINIT    := 0;         iPSA_TYPDATANC        := 0;
  iPSA_SEXE             := 0;         iPSA_DATEACQCPANC     := 0;
  IPSA_CONDEMPLOI       := 0;         iPSA_NBRECPSUPP       := 0;
  iPSA_CONGESPAYES      := 0;         iPSA_CPTYPERELIQ      := 0;
  iPSA_CPACQUISSUPP     := 0;         iPSA_RELIQUAT         := 0;
  iPSA_ANCIENPOSTE      := 0;         iPSA_UNITEPRISEFF     := 0;
  iPSA_TYPPAIEVALOMS    := 0;         iPSA_ACTIVITE         := 0;
  iPSA_PAIEVALOMS       := 0;         iPSA_CATDADS          := 0; // PT203
end;


procedure MemorisePhc(UneTob: TOB);
begin
  if Assigned(UneTob) then
    with UneTob do
    begin
      iPHC_ETABLISSEMENT := GetNumChamp('PHC_ETABLISSEMENT');
      iPHC_SALARIE := GetNumChamp('PHC_SALARIE');
      iPHC_DATEDEBUT := GetNumChamp('PHC_DATEDEBUT');
      iPHC_DATEFIN := GetNumChamp('PHC_DATEFIN');
      iPHC_REPRISE := GetNumChamp('PHC_REPRISE');
      iPHC_CUMULPAIE := GetNumChamp('PHC_CUMULPAIE');
      iPHC_MONTANT := GetNumChamp('PHC_MONTANT');
      iPHC_TRAVAILN1 := GetNumChamp('PHC_TRAVAILN1');
      iPHC_TRAVAILN2 := GetNumChamp('PHC_TRAVAILN2');
      iPHC_TRAVAILN3 := GetNumChamp('PHC_TRAVAILN3');
      iPHC_TRAVAILN4 := GetNumChamp('PHC_TRAVAILN4');
      iPHC_CODESTAT := GetNumChamp('PHC_CODESTAT');
      iPHC_CONFIDENTIEL := GetNumChamp('PHC_CONFIDENTIEL');
      iPHC_LIBREPCMB1 := GetNumChamp('PHC_LIBREPCMB1');
      iPHC_LIBREPCMB2 := GetNumChamp('PHC_LIBREPCMB2');
      iPHC_LIBREPCMB3 := GetNumChamp('PHC_LIBREPCMB3');
      iPHC_LIBREPCMB4 := GetNumChamp('PHC_LIBREPCMB4');
    end;
end;

procedure MemorisePhb(THB: TOB);
begin
  if Assigned(THB) then
    with THB do
    begin
      iPHB_ETABLISSEMENT := GetNumChamp('PHB_ETABLISSEMENT');
      iPHB_SALARIE := GetNumChamp('PHB_SALARIE');
      iPHB_DATEDEBUT := GetNumChamp('PHB_DATEDEBUT');
      iPHB_DATEFIN := GetNumChamp('PHB_DATEFIN');
      iPHB_NATURERUB := GetNumChamp('PHB_NATURERUB');
      iPHB_RUBRIQUE := GetNumChamp('PHB_RUBRIQUE');
      iPHB_LIBELLE := GetNumChamp('PHB_LIBELLE');
      iPHB_IMPRIMABLE := GetNumChamp('PHB_IMPRIMABLE');
      iPHB_TAUXREM := GetNumChamp('PHB_TAUXREM');
      iPHB_COEFFREM := GetNumChamp('PHB_COEFFREM');
      iPHB_TAUXSALARIAL := GetNumChamp('PHB_TAUXSALARIAL');
      iPHB_TAUXPATRONAL := GetNumChamp('PHB_TAUXPATRONAL');
      iPHB_BASEREMIMPRIM := GetNumChamp('PHB_BASEREMIMPRIM');
      iPHB_TAUXREMIMPRIM := GetNumChamp('PHB_TAUXREMIMPRIM');
      iPHB_COEFFREMIMPRIM := GetNumChamp('PHB_COEFFREMIMPRIM');
      iPHB_BASECOTIMPRIM := GetNumChamp('PHB_BASECOTIMPRIM');
      iPHB_TAUXSALIMPRIM := GetNumChamp('PHB_TAUXSALIMPRIM');
      iPHB_TAUXPATIMPRIM := GetNumChamp('PHB_TAUXPATIMPRIM');
      iPHB_ORGANISME := GetNumChamp('PHB_ORGANISME');
      iPHB_TAUXAT := GetNumChamp('PHB_TAUXAT');
      iPHB_PLAFOND := GetNumChamp('PHB_PLAFOND');
      iPHB_PLAFOND1 := GetNumChamp('PHB_PLAFOND1');
      iPHB_PLAFOND2 := GetNumChamp('PHB_PLAFOND2');
      iPHB_PLAFOND3 := GetNumChamp('PHB_PLAFOND3');
      iPHB_CONSERVATION := GetNumChamp('PHB_CONSERVATION');
      iPHB_ORDREETAT := GetNumChamp('PHB_ORDREETAT');
      iPHB_SENSBUL := GetNumChamp('PHB_SENSBUL');
      iPHB_ORIGINELIGNE := GetNumChamp('PHB_ORIGINELIGNE');
      iPHB_ORIGINEINFO := GetNumChamp('PHB_ORIGINEINFO');
      iPHB_TRAVAILN1 := GetNumChamp('PHB_TRAVAILN1');
      iPHB_TRAVAILN2 := GetNumChamp('PHB_TRAVAILN2');
      iPHB_TRAVAILN3 := GetNumChamp('PHB_TRAVAILN3');
      iPHB_TRAVAILN4 := GetNumChamp('PHB_TRAVAILN4');
      iPHB_CODESTAT := GetNumChamp('PHB_CODESTAT');
      iPHB_CONFIDENTIEL := GetNumChamp('PHB_CONFIDENTIEL');
      iPHB_LIBREPCMB1 := GetNumChamp('PHB_LIBREPCMB1');
      iPHB_LIBREPCMB2 := GetNumChamp('PHB_LIBREPCMB2');
      iPHB_LIBREPCMB3 := GetNumChamp('PHB_LIBREPCMB3');
      iPHB_LIBREPCMB4 := GetNumChamp('PHB_LIBREPCMB4');
      iPHB_BASEREM := GetNumChamp('PHB_BASEREM');
      iPHB_MTREM := GetNumChamp('PHB_MTREM');
      iPHB_BASECOT := GetNumChamp('PHB_BASECOT');
      iPHB_MTSALARIAL := GetNumChamp('PHB_MTSALARIAL');
      iPHB_MTPATRONAL := GetNumChamp('PHB_MTPATRONAL');
      iPHB_TRANCHE1 := GetNumChamp('PHB_TRANCHE1');
      iPHB_TRANCHE2 := GetNumChamp('PHB_TRANCHE2');
      iPHB_TRANCHE3 := GetNumChamp('PHB_TRANCHE3');
      iPHB_ORDREETAT := GetNumChamp('PHB_ORDREETAT');
      iPHB_COTREGUL := GetNumChamp('PHB_COTREGUL');
      iPHB_OMTSALARIAL := GetNumChamp('PHB_OMTSALARIAL'); { PT136 }
    end;
end;

procedure MemorisePcl(TRechCum: TOB);
begin
  with TRechCum do
  begin
    iPCL_ALIMCUMUL := GetNumChamp('PCL_ALIMCUMUL');
    iPCL_ALIMCUMULCOT := GetNumChamp('PCL_ALIMCUMULCOT');
    iPCL_COEFFAFFECT := GetNumChamp('PCL_COEFFAFFECT');
  end;
end;

procedure MemorisePcr(TRechRub: TOB);
begin
  with TRechRub do
  begin
    iPCR_CUMULPAIE := GetNumChamp('PCR_CUMULPAIE');
    iPCR_SENS := GetNumChamp('PCR_SENS');
  end;
end;

procedure MemorisePct(UneCot: TOB);
begin
  with UneCot do
  begin
    iPCT_IMPRIMABLE := GetNumChamp('PCT_IMPRIMABLE');
    iPCT_BASEIMP := GetNumChamp('PCT_BASEIMP');
    iPCT_TXSALIMP := GetNumChamp('PCT_TXSALIMP');
    iPCT_ORGANISME := GetNumChamp('PCT_ORGANISME');
    iPCT_ORDREETAT := GetNumChamp('PCT_ORDREETAT');
    iPCT_LIBELLE := GetNumChamp('PCT_LIBELLE');
    iPCT_DU := GetNumChamp('PCT_DU');
    iPCT_AU := GetNumChamp('PCT_AU');
    iPCT_MOIS1 := GetNumChamp('PCT_MOIS1');
    iPCT_MOIS2 := GetNumChamp('PCT_MOIS2');
    iPCT_MOIS3 := GetNumChamp('PCT_MOIS3');
    iPCT_MOIS4 := GetNumChamp('PCT_MOIS4');
    iPCT_DECBASE := GetNumChamp('PCT_DECBASE');
    iPCT_TYPEBASE := GetNumChamp('PCT_TYPEBASE');
    iPCT_TYPETAUXSAL := GetNumChamp('PCT_TYPETAUXSAL');
    iPCT_TAUXSAL := GetNumChamp('PCT_TAUXSAL');
    iPCT_DECTXSAL := GetNumChamp('PCT_DECTXSAL');
    iPCT_TYPETAUXPAT := GetNumChamp('PCT_TYPETAUXPAT');
    iPCT_TAUXPAT := GetNumChamp('PCT_TAUXPAT');
    iPCT_DECTXPAT := GetNumChamp('PCT_DECTXPAT');
    iPCT_TYPEFFSAL := GetNumChamp('PCT_TYPEFFSAL');
    iPCT_FFSAL := GetNumChamp('PCT_FFSAL');
    iPCT_DECMTSAL := GetNumChamp('PCT_DECMTSAL');
    iPCT_TYPEFFPAT := GetNumChamp('PCT_TYPEFFPAT');
    iPCT_FFPAT := GetNumChamp('PCT_FFPAT');
    iPCT_DECMTPAT := GetNumChamp('PCT_DECMTPAT');
    iPCT_TYPEMINISAL := GetNumChamp('PCT_TYPEMINISAL');
    iPCT_VALEURMINISAL := GetNumChamp('PCT_VALEURMINISAL');
    iPCT_DECMTSAL := GetNumChamp('PCT_DECMTSAL');
    iPCT_TYPEMAXISAL := GetNumChamp('PCT_TYPEMAXISAL');
    iPCT_VALEURMAXISAL := GetNumChamp('PCT_VALEURMAXISAL');
    iPCT_TYPEMINIPAT := GetNumChamp('PCT_TYPEMINIPAT');
    iPCT_VALEURMINIPAT := GetNumChamp('PCT_VALEURMINIPAT');
    iPCT_DECMTPAT := GetNumChamp('PCT_DECMTPAT');
    iPCT_TYPEMAXIPAT := GetNumChamp('PCT_TYPEMAXIPAT');
    iPCT_VALEURMAXIPAT := GetNumChamp('PCT_VALEURMAXIPAT');
    iPCT_DECMTPAT := GetNumChamp('PCT_DECMTPAT');
    iPCT_DECBASECOT := GetNumChamp('PCT_DECBASECOT');
    iPCT_SOUMISREGUL := GetNumChamp('PCT_SOUMISREGUL');
    iPCT_BASECOTISATION := GetNumChamp('PCT_BASECOTISATION');
    iPCT_TYPEPLAFOND := GetNumChamp('PCT_TYPEPLAFOND');
    iPCT_PLAFOND := GetNumChamp('PCT_PLAFOND');
    iPCT_TYPETRANCHE1 := GetNumChamp('PCT_TYPETRANCHE1');
    iPCT_TYPETRANCHE2 := GetNumChamp('PCT_TYPETRANCHE2');
    iPCT_TYPETRANCHE3 := GetNumChamp('PCT_TYPETRANCHE3');
    iPCT_TYPEREGUL := GetNumChamp('PCT_TYPEREGUL');
    iPCT_TXPATIMP := GetNumChamp('PCT_TXPATIMP');
    iPCT_TRANCHE1 := GetNumChamp('PCT_TRANCHE1');
    iPCT_TRANCHE2 := GetNumChamp('PCT_TRANCHE2');
    iPCT_TRANCHE3 := GetNumChamp('PCT_TRANCHE3');
    iPCT_ORDREAT := GetNumChamp('PCT_ORDREAT');
    iPCT_CODETRANCHE := GetNumChamp('PCT_CODETRANCHE');
    iPCT_PRESFINMOIS := GetNumChamp('PCT_PRESFINMOIS');
  end;
end;

procedure MemorisePrm(UneRem: TOB);
begin
  with UneRem do
  begin
    iPRM_IMPRIMABLE := GetNumChamp('PRM_IMPRIMABLE');
    iPRM_BASEIMPRIMABLE := GetNumChamp('iPRM_BASEIMPRIMABLE');
    iPRM_BASEIMPRIMABLE := GetNumChamp('PRM_BASEIMPRIMABLE');
    iPRM_TAUXIMPRIMABLE := GetNumChamp('PRM_TAUXIMPRIMABLE');
    iPRM_COEFFIMPRIM := GetNumChamp('PRM_COEFFIMPRIM');
    iPRM_ORDREETAT := GetNumChamp('PRM_ORDREETAT');
    iPRM_SENSBUL := GetNumChamp('PRM_SENSBUL');
    iPRM_LIBELLE := GetNumChamp('PRM_LIBELLE');
    iPRM_THEMEREM := GetNumChamp('PRM_THEMEREM');
    iPRM_DECBASE := GetNumChamp('PRM_DECBASE');
    iPRM_DECTAUX := GetNumChamp('PRM_DECTAUX');
    iPRM_DECCOEFF := GetNumChamp('PRM_DECCOEFF');
    iPRM_DECMONTANT := GetNumChamp('PRM_DECMONTANT');
    iPRM_LIBCONTRAT := GetNumChamp('PRM_LIBCONTRAT');
  end;
end;

procedure MemorisePel(UnElt: TOB);
begin
  with UnElt do
  begin
    iChampPredefini := GetNumChamp('PEL_PREDEFINI');
    iCHampDossier := GetNumChamp('PEL_NODOSSIER');
    iCHampCodeElt := GetNumChamp('PEL_CODEELT');
    iCHampDateValidite := GetNumChamp('PEL_DATEVALIDITE');
    iCHampDecaleMois := GetNumChamp('PEL_DECALMOIS');
    iCHampMontant := GetNumChamp('PEL_MONTANT');
    iCHampMontantEuro := GetNumChamp('PEL_MONTANTEURO');
    // PT100  : 05/04/2004 PH V_50 FQ 11231 Prise en compte du cas spécifique Alsace Lorraine
    iChampRegimeAlsace := GetNumChamp('PEL_REGIMEALSACE');
    iChampConvention :=  GetNumChamp('PEL_CONVENTION'); //PT172
  end;
end;

//DEB PT172
procedure MemorisePed(UnElt: TOB);
begin
  with UnElt do
  begin
    iCHampPEDCodeElt := GetNumChamp('PED_CODEELT');
    iCHampPEDDateValidite := GetNumChamp('PED_DATEVALIDITE');
    iCHampPEDDecaleMois := GetNumChamp('PED_DECALMOIS');
    iCHampPEDMontant := GetNumChamp('PED_MONTANT');
    iCHampPEDMontantEuro := GetNumChamp('PED_MONTANTEURO');
    iChampPEDRegimeAlsace := GetNumChamp('PED_REGIMEALSACE');
    iChampPEDValeurNiveau := GetNumChamp('PED_VALEURNIVEAU');
  end;
end;

procedure MemorisePnr(UnElt: TOB);
begin
  with UnElt do
  begin
    iCHampPNRCodeElt := GetNumChamp('PNR_CODEELT');
    iCHampPNRNiveauRequis := GetNumChamp('PNR_TYPENIVEAU');
    iCHampPNRNiveauMaxi := GetNumChamp('PNR_NIVMAXPERSO');
  end;
end;
//FIN PT172

procedure MemorisePsa(Unsal: TOB);
begin
  with UnSal do
  begin
    iPSA_SALARIE := GetNumChamp('PSA_SALARIE');
    iPSA_ETABLISSEMENT := GetNumChamp('PSA_ETABLISSEMENT');
    iPSA_TRAVAILN1 := GetNumChamp('PSA_TRAVAILN1');
    iPSA_TRAVAILN2 := GetNumChamp('PSA_TRAVAILN2');
    iPSA_TRAVAILN3 := GetNumChamp('PSA_TRAVAILN3');
    iPSA_TRAVAILN4 := GetNumChamp('PSA_TRAVAILN4');
    iPSA_CODESTAT := GetNumChamp('PSA_CODESTAT');
    iPSA_LIBREPCMB1 := GetNumChamp('PSA_LIBREPCMB1');
    iPSA_LIBREPCMB2 := GetNumChamp('PSA_LIBREPCMB2');
    iPSA_LIBREPCMB3 := GetNumChamp('PSA_LIBREPCMB3');
    iPSA_LIBREPCMB4 := GetNumChamp('PSA_LIBREPCMB4');
    iPSA_CONFIDENTIEL := GetNumChamp('PSA_CONFIDENTIEL');
    iPSA_TAUXHORAIRE := GetNumChamp('PSA_TAUXHORAIRE');
    iPSA_SALAIREMOIS1 := GetNumChamp('PSA_SALAIREMOIS1');
    iPSA_SALAIREMOIS2 := GetNumChamp('PSA_SALAIREMOIS2');
    iPSA_SALAIREMOIS3 := GetNumChamp('PSA_SALAIREMOIS3');
    iPSA_SALAIREMOIS4 := GetNumChamp('PSA_SALAIREMOIS4');
    iPSA_SALAIREMOIS5 := GetNumChamp('PSA_SALAIREMOIS5');
    iPSA_SALAIRANN1 := GetNumChamp('PSA_SALAIRANN1');
    iPSA_SALAIRANN2 := GetNumChamp('PSA_SALAIRANN2');
    iPSA_SALAIRANN3 := GetNumChamp('PSA_SALAIRANN3');
    iPSA_SALAIRANN4 := GetNumChamp('PSA_SALAIRANN4');
    iPSA_SALAIRANN5 := GetNumChamp('PSA_SALAIRANN5');
    iPSA_DATENAISSANCE := GetNumChamp('PSA_DATENAISSANCE');
    iPSA_DATEENTREE := GetNumChamp('PSA_DATEENTREE');
    iPSA_DATESORTIE := GetNumChamp('PSA_DATESORTIE');
    iPSA_TYPPROFILREM := GetNumChamp('PSA_TYPPROFILREM');
    iPSA_PROFILREM := GetNumChamp('PSA_PROFILREM');
    iPSA_TYPPROFIL := GetNumChamp('PSA_TYPPROFIL');
    iPSA_PROFIL := GetNumChamp('PSA_PROFIL');
    iPSA_TYPPERIODEBUL := GetNumChamp('PSA_TYPPERIODEBUL');
    iPSA_PERIODBUL := GetNumChamp('PSA_PERIODBUL');
    iPSA_TYPPROFILRBS := GetNumChamp('PSA_TYPPROFILRBS');
    iPSA_PROFILRBS := GetNumChamp('PSA_PROFILRBS');
    iPSA_TYPREDREPAS := GetNumChamp('PSA_TYPREDREPAS');
    iPSA_REDREPAS := GetNumChamp('PSA_REDREPAS');
    iPSA_TYPREDRTT1 := GetNumChamp('PSA_TYPREDRTT1');
    iPSA_REDRTT1 := GetNumChamp('PSA_REDRTT1');
    iPSA_TYPREDRTT2 := GetNumChamp('PSA_TYPREDRTT2');
    iPSA_REDRTT2 := GetNumChamp('PSA_REDRTT2');
    iPSA_PROFILTPS := GetNumChamp('PSA_PROFILTPS');
    iPSA_TYPPROFILAFP := GetNumChamp('PSA_TYPPROFILAFP');
    iPSA_PROFILAFP := GetNumChamp('PSA_PROFILAFP');
    iPSA_TYPPROFILAPP := GetNumChamp('PSA_TYPPROFILAPP');
    iPSA_PROFILAPP := GetNumChamp('PSA_PROFILAPP');
    iPSA_TYPPROFILRET := GetNumChamp('PSA_TYPPROFILRET');
    iPSA_PROFILRET := GetNumChamp('PSA_PROFILRET');
    iPSA_TYPPROFILMUT := GetNumChamp('PSA_TYPPROFILMUT');
    iPSA_PROFILMUT := GetNumChamp('PSA_PROFILMUT');
    iPSA_TYPPROFILPRE := GetNumChamp('PSA_TYPPROFILPRE');
    iPSA_PROFILPRE := GetNumChamp('PSA_PROFILPRE');
    iPSA_TYPPROFILTSS := GetNumChamp('PSA_TYPPROFILTSS');
    iPSA_PROFILTSS := GetNumChamp('PSA_PROFILTSS');
    iPSA_TYPPROFILCGE := GetNumChamp('PSA_TYPPROFILCGE');
    iPSA_PROFILCGE := GetNumChamp('PSA_PROFILCGE');
    iPSA_PROFILCDD := GetNumChamp('PSA_PROFILCDD');
    iPSA_PROFILMUL := GetNumChamp('PSA_PROFILCDD');
    iPSA_TYPPROFILFNAL := GetNumChamp('PSA_TYPPROFILFNAL');
    iPSA_PROFILFNAL := GetNumChamp('PSA_PROFILFNAL');
    iPSA_TYPPROFILTRANS := GetNumChamp('PSA_TYPPROFILTRANS');
    iPSA_PROFILTRANS := GetNumChamp('PSA_PROFILTRANS');
    iPSA_TYPPROFILANC := GetNumChamp('PSA_TYPPROFILANC');
    iPSA_PROFILANCIEN := GetNumChamp('PSA_PROFILANCIEN');
    iPSA_LIBELLE := GetNumChamp('PSA_LIBELLE');
    iPSA_PRENOM := GetNumChamp('PSA_PRENOM');
    iPSA_NUMEROSS := GetNumChamp('PSA_NUMEROSS');
    iPSA_ADRESSE1 := GetNumChamp('PSA_ADRESSE1');
    iPSA_ADRESSE2 := GetNumChamp('PSA_ADRESSE2');
    iPSA_ADRESSE3 := GetNumChamp('PSA_ADRESSE3');
    iPSA_CODEPOSTAL := GetNumChamp('PSA_CODEPOSTAL');
    iPSA_VILLE := GetNumChamp('PSA_VILLE');
    iPSA_INDICE := GetNumChamp('PSA_INDICE');
    iPSA_NIVEAU := GetNumChamp('PSA_NIVEAU');
    iPSA_CONVENTION := GetNumChamp('PSA_CONVENTION');
    iPSA_CODEEMPLOI := GetNumChamp('PSA_CODEEMPLOI');
    iPSA_AUXILIAIRE := GetNumChamp('PSA_AUXILIAIRE');
    iPSA_DATEANCIENNETE := GetNumChamp('PSA_DATEANCIENNETE');
    iPSA_QUALIFICATION := GetNumChamp('PSA_QUALIFICATION');
    iPSA_COEFFICIENT := GetNumChamp('PSA_COEFFICIENT');
    iPSA_LIBELLEEMPLOI := GetNumChamp('PSA_LIBELLEEMPLOI');
    iPSA_CIVILITE := GetNumChamp('PSA_CIVILITE');
    iPSA_CPACQUISMOIS := GetNumChamp('PSA_CPACQUISMOIS');
    iPSA_NBREACQUISCP := GetNumChamp('PSA_NBREACQUISCP');
    iPSA_TYPDATPAIEMENT := GetNumChamp('PSA_TYPDATPAIEMENT');
    iPSA_MOISPAIEMENT := GetNumChamp('PSA_MOISPAIEMENT');
    iPSA_JOURPAIEMENT := GetNumChamp('PSA_JOURPAIEMENT');
    iPSA_TYPREGLT := GetNumChamp('PSA_TYPREGLT');
    iPSA_PGMODEREGLE := GetNumChamp('PSA_PGMODEREGLE');
    iPSA_REGULANCIEN := GetNumChamp('PSA_REGULANCIEN');
    iPSA_HORHEBDO := GetNumChamp('PSA_HORHEBDO');
    iPSA_HORAIREMOIS := GetNumChamp('PSA_HORAIREMOIS');
    iPSA_HORANNUEL := GetNumChamp('PSA_HORANNUEL');
    iPSA_PERSACHARGE := GetNumChamp('PSA_PERSACHARGE');
    iPSA_PCTFRAISPROF := GetNumChamp('PSA_PCTFRAISPROF');
    iPSA_MULTIEMPLOY := GetNumChamp('PSA_MULTIEMPLOY');
    iPSA_SALAIREMULTI := GetNumChamp('PSA_SALAIREMULTI');
    iPSA_ORDREAT := GetNumChamp('PSA_ORDREAT');
    iPSA_SALAIRETHEO := GetNumChamp('PSA_SALAIRETHEO');
    iPSA_DATELIBRE1 := GetNumChamp('PSA_DATELIBRE1');
    iPSA_DATELIBRE2 := GetNumChamp('PSA_DATELIBRE2');
    iPSA_DATELIBRE3 := GetNumChamp('PSA_DATELIBRE3');
    iPSA_DATELIBRE4 := GetNumChamp('PSA_DATELIBRE4');
    iPSA_VALANCCP := GetNumChamp('PSA_VALANCCP');
    iPSA_ANCIENNETE := GetNumChamp('PSA_ANCIENNETE');
    iPSA_CALENDRIER := GetNumChamp('PSA_CALENDRIER');
    iPSA_STANDCALEND := GetNumChamp('PSA_STANDCALEND');
    iPSA_BOOLLIBRE1 := GetNumChamp('PSA_BOOLLIBRE1');
    iPSA_BOOLLIBRE2 := GetNumChamp('PSA_BOOLLIBRE2');
    iPSA_BOOLLIBRE3 := GetNumChamp('PSA_BOOLLIBRE3');
    iPSA_BOOLLIBRE4 := GetNumChamp('PSA_BOOLLIBRE4');
    iPSA_DADSPROF := GetNumChamp('PSA_DADSPROF');
    iPSA_DADSCAT := GetNumChamp('PSA_DADSCAT');
    iPSA_TAUXPARTIEL := GetNumChamp('PSA_TAUXPARTIEL');
    iPSA_CPTYPEMETHOD := GetNumChamp('PSA_CPTYPEMETHOD');
    iPSA_VALORINDEMCP := GetNumChamp('PSA_VALORINDEMCP');
    iPSA_CPTYPEVALO := GetNumChamp('PSA_CPTYPEVALO');
    iPSA_MVALOMS := GetNumChamp('PSA_MVALOMS');
    iPSA_VALODXMN := GetNumChamp('PSA_VALODXMN');
    iPSA_CPACQUISANC := GetNumChamp('PSA_CPACQUISANC');
    iPSA_BASANCCP := GetNumChamp('PSA_BASANCCP');
    iPSA_VALANCCP := GetNumChamp('PSA_VALANCCP');
    iPSA_DATANC := GetNumChamp('PSA_DATANC');
    iPSA_TYPDATANC := GetNumChamp('PSA_TYPDATANC');
    iPSA_DATEACQCPANC := GetNumChamp('PSA_DATEACQCPANC');
    iPSA_NBRECPSUPP := GetNumChamp('PSA_NBRECPSUPP');
    iPSA_CPTYPERELIQ := GetNumChamp('PSA_CPTYPERELIQ');
    iPSA_RELIQUAT := GetNumChamp('PSA_RELIQUAT');
    iPSA_SORTIEDEFINIT := GetNumChamp('PSA_SORTIEDEFINIT');
    iPSA_SEXE := GetNumChamp('PSA_SEXE');
    IPSA_CONDEMPLOI := GetNumChamp('PSA_CONDEMPLOI');
    iPSA_CONGESPAYES := GetNumChamp('PSA_CONGESPAYES'); { PT100-1 }
    iPSA_CPACQUISSUPP := GetNumChamp('PSA_CPACQUISSUPP'); { PT100-2 }
    iPSA_ANCIENPOSTE := GetNumChamp('PSA_ANCIENPOSTE'); // PT110
    iPSA_TYPPAIEVALOMS := GetNumChamp('PSA_TYPPAIEVALOMS'); { PT150 }
    iPSA_PAIEVALOMS := GetNumChamp('PSA_PAIEVALOMS'); { PT150 }
    iPSA_UNITEPRISEFF := GetNumChamp('PSA_UNITEPRISEFF');
    iPSA_ACTIVITE := GetNumChamp('PSA_ACTIVITE');
    iPSA_CATDADS := GetNumChamp ('PSA_CATDADS'); // PT203
  end;
end;

function NumChampTobS(const zz: string; const Ind: Integer): Integer;
var
  ii: Integer;
begin
  ii := -1;
  if zz = 'OR' then
  begin
    case ind of
      1: ii := iPSA_TRAVAILN1;
      2: ii := iPSA_TRAVAILN2;
      3: ii := iPSA_TRAVAILN3;
      4: ii := iPSA_TRAVAILN4;
    end;
  end
  else
  begin
    if zz = 'TC' then
    begin
      case ind of
        1: ii := iPSA_LIBREPCMB1;
        2: ii := iPSA_LIBREPCMB2;
        3: ii := iPSA_LIBREPCMB3;
        4: ii := iPSA_LIBREPCMB4;
      end;
    end
    else
    begin
      if zz = 'DT' then
      begin
        case ind of
          1: ii := iPSA_DATELIBRE1;
          2: ii := iPSA_DATELIBRE2;
          3: ii := iPSA_DATELIBRE3;
          4: ii := iPSA_DATELIBRE4;
        end;
      end;
    end;
  end;
  result := ii; // PT119 Affectation de la valeur trouvée
end;

function NumChampProfS(const Ch: string): Integer;
begin
  if Ch = 'PSA_PROFILREM' then result := iPSA_PROFILREM
  else if Ch = 'PSA_PROFIL' then result := iPSA_PROFIL
  else if Ch = 'PSA_PERIODBUL' then result := iPSA_PERIODBUL
  else if Ch = 'PSA_PROFILRBS' then result := iPSA_PROFILRBS
  else if Ch = 'PSA_REDREPAS' then result := iPSA_REDREPAS
  else if Ch = 'PSA_REDRTT1' then result := iPSA_REDRTT1
  else if Ch = 'PSA_REDRTT2' then result := iPSA_REDRTT2
  else if Ch = 'PSA_PROFILTPS' then result := iPSA_PROFILTPS
  else if Ch = 'PSA_PROFILAFP' then result := iPSA_PROFILAFP
  else if Ch = 'PSA_PROFILAPP' then result := iPSA_PROFILAPP
  else if Ch = 'PSA_PROFILRET' then result := iPSA_PROFILRET
  else if Ch = 'PSA_PROFILMUT' then result := iPSA_PROFILMUT
  else if Ch = 'PSA_PROFILPRE' then result := iPSA_PROFILPRE
  else if Ch = 'PSA_PROFILTSS' then result := iPSA_PROFILTSS
  else if Ch = 'PSA_PROFILCGE' then result := iPSA_PROFILCGE
  else if Ch = 'PSA_PROFILCDD' then result := iPSA_PROFILCDD
  else if Ch = 'PSA_PROFILMUL' then result := iPSA_PROFILMUL
  else if Ch = 'PSA_PROFILFNAL' then result := iPSA_PROFILFNAL
  else if Ch = 'PSA_PROFILTRANS' then result := iPSA_PROFILTRANS
  else if Ch = 'PSA_PROFILANCIEN' then result := iPSA_PROFILANCIEN
  else result := -1;
end;
// d PT103 IJSS
{$IFNDEF CCS3}

procedure AnnuleIJSSBulletin(const CodeSalarie: string; const DateF: tdatetime);
begin
  Executesql('UPDATE REGLTIJSS SET PRI_DATEINTEGR = "' + usdatetime(IDate1900) + '" ' + // PT192 DateToStr
    'WHERE PRI_SALARIE="' + CodeSalarie + '" AND ' +
    'PRI_DATEINTEGR = "' + usdatetime(Datef) + '"');
end;
{$ENDIF}
// f PT103 IJSS

// d PT121 Maintien
{$IFNDEF CCS3}

procedure AnnuleMaintienBulletin(const CodeSalarie: string; const DateD, DateF: tdatetime);
begin
  Executesql('DELETE FROM MAINTIEN ' +
    'WHERE PMT_SALARIE="' + CodeSalarie + '" AND ' +
    '((PMT_DATEDEBUT = "' + usdatetime(DateD) + '" AND ' +
    'PMT_DATEFIN = "' + usdatetime(DateF) + '") OR ' +
    '(PMT_DATEDEBUT = "' + usdatetime(2) + '" AND ' +
    'PMT_DATEFIN = "' + usdatetime(2) + '"))');
end;
{$ENDIF}
// f PT121 Maintien
{$IFNDEF CCS3}
// d PT124

procedure EcritCommMaint(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const RubMaintien, NatRub: string);
var
  Thh, TR: tob;
  libelle: string;
begin
  TR := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [RubMaintien], FALSE);
  THH := TOB.create('HISTOBULLETIN', Tob_Rub, -1);
  THH.PutValue('PHB_RUBRIQUE', RubMaintien + '.' + IntToStr(i));
  THH.PutValue('PHB_NATURERUB', Natrub);
  THH.PutValue('PHB_DATEDEBUT', DateD);
  THH.PutValue('PHB_DATEFIN', DateF);
  THH.PutValue('PHB_ETABLISSEMENT', Salarie.GetValeur(iPSA_ETABLISSEMENT));
  THH.PutValue('PHB_SALARIE', Salarie.GetValeur(iPSA_SALARIE));
  THH.PutValue('PHB_BASEREM', 0);
  THH.PutValue('PHB_TAUXREM', 0);
  THH.PutValue('PHB_COEFFREM', 0);
  THH.PutValue('PHB_MTREM', 0);
  THH.PutValue('PHB_CONSERVATION', 'BUL');
// d PT206
  if (T.GetValue('PMT_NBJMAINTIEN') <= 1) then
  begin

    libelle := copy(DateToStr(T.GetValue('PMT_DATEDEBUTABS')),1,5) +
              ' au ' +
              copy(DateToStr(T.GetValue('PMT_DATEFINABS')),1,5) +
              ' '+
              IntToStr(T.GetValue('PMT_NBJMAINTIEN'))+
              ' jour à ' +
              FloatToStrF(T.GetValue('PMT_PCTMAINTIEN'), ffNumber, 20, 2) +
              '%';
  end
  else
  begin
    libelle := copy(DateToStr(T.GetValue('PMT_DATEDEBUTABS')),1,5) +
              ' au ' +
              copy(DateToStr(T.GetValue('PMT_DATEFINABS')),1,5) +
              ' '+
              IntToStr(T.GetValue('PMT_NBJMAINTIEN'))+
              ' jours à ' +
              FloatToStrF(T.GetValue('PMT_PCTMAINTIEN'), ffNumber, 20, 2) +
              '%';
// f PT206
  end;

  if (T.GetValue('PMT_TYPECONGE') = 'NET') or
    (T.GetValue('PMT_TYPECONGE') = 'BRU') then
  begin
    libelle := Trim(T.GetValue('PMT_LIBELLE')) +
      ' ' +
      FloatToStrF(T.GetValue('PMT_MTMAINTIEN'), ffNumber, 20, 2);
  end;
  THH.PutValue('PHB_LIBELLE', libelle);
  THH.PutValue('PHB_ORIGINELIGNE', 'MAI');
  THH.PutValue('PHB_IMPRIMABLE', 'X');
  THH.PutValue('PHB_TRAVAILN2', Salarie.GetValeur(iPSA_TRAVAILN2));
  THH.PutValue('PHB_TRAVAILN3', Salarie.GetValeur(iPSA_TRAVAILN3));
  THH.PutValue('PHB_TRAVAILN4', Salarie.GetValeur(iPSA_TRAVAILN4));
  THH.PutValue('PHB_TRAVAILN1', Salarie.GetValeur(iPSA_TRAVAILN1));
  THH.PutValue('PHB_CODESTAT', Salarie.GetValeur(iPSA_CODESTAT));
  THH.PutValue('PHB_LIBREPCMB1', Salarie.GetValeur(iPSA_LIBREPCMB1));
  THH.PutValue('PHB_LIBREPCMB2', Salarie.GetValeur(iPSA_LIBREPCMB2));
  THH.PutValue('PHB_LIBREPCMB3', Salarie.GetValeur(iPSA_LIBREPCMB3));
  THH.PutValue('PHB_LIBREPCMB4', Salarie.GetValeur(iPSA_LIBREPCMB4));
  THH.PutValue('PHB_CONFIDENTIEL', Salarie.GetValeur(iPSA_CONFIDENTIEL));
  if TR <> nil then
  begin
    THH.PutValue('PHB_ORDREETAT', TR.GetValue('PRM_ORDREETAT'));
    THH.PutValue('PHB_OMTSALARIAL', TR.GetValue('PRM_ORDREETAT')); { PT136 }
  end;
  i := i + 1;
end;
// f PT124
// d PT129

procedure EcritCommIJ(Tob_rub, Salarie, t: tob; var i: integer; const DateD, Datef: tdatetime; const RubIJ, NatRub: string);
var
  Thh, TR: tob;
  libelle: string;
begin
  TR := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [RubIJ], FALSE);
  THH := TOB.create('HISTOBULLETIN', Tob_Rub, -1);
  THH.PutValue('PHB_RUBRIQUE', RubIJ + '.' + IntToStr(i));
  THH.PutValue('PHB_NATURERUB', Natrub);
  THH.PutValue('PHB_DATEDEBUT', DateD);
  THH.PutValue('PHB_DATEFIN', DateF);
  THH.PutValue('PHB_ETABLISSEMENT', Salarie.GetValeur(iPSA_ETABLISSEMENT));
  THH.PutValue('PHB_SALARIE', Salarie.GetValeur(iPSA_SALARIE));
  THH.PutValue('PHB_BASEREM', 0);
  THH.PutValue('PHB_TAUXREM', 0);
  THH.PutValue('PHB_COEFFREM', 0);
  THH.PutValue('PHB_MTREM', 0);
  THH.PutValue('PHB_CONSERVATION', 'BUL');
  libelle := 'IJSS du ' +
    DateToStr(T.GetValue('PRI_DATEDEBUTABS')) +
    ' au ' +
    DateToStr(T.GetValue('PRI_DATEFINABS'));

  THH.PutValue('PHB_LIBELLE', libelle);
  THH.PutValue('PHB_ORIGINELIGNE', 'IJS');
  THH.PutValue('PHB_IMPRIMABLE', 'X');
  THH.PutValue('PHB_TRAVAILN2', Salarie.GetValeur(iPSA_TRAVAILN2));
  THH.PutValue('PHB_TRAVAILN3', Salarie.GetValeur(iPSA_TRAVAILN3));
  THH.PutValue('PHB_TRAVAILN4', Salarie.GetValeur(iPSA_TRAVAILN4));
  THH.PutValue('PHB_TRAVAILN1', Salarie.GetValeur(iPSA_TRAVAILN1));
  THH.PutValue('PHB_CODESTAT', Salarie.GetValeur(iPSA_CODESTAT));
  THH.PutValue('PHB_LIBREPCMB1', Salarie.GetValeur(iPSA_LIBREPCMB1));
  THH.PutValue('PHB_LIBREPCMB2', Salarie.GetValeur(iPSA_LIBREPCMB2));
  THH.PutValue('PHB_LIBREPCMB3', Salarie.GetValeur(iPSA_LIBREPCMB3));
  THH.PutValue('PHB_LIBREPCMB4', Salarie.GetValeur(iPSA_LIBREPCMB4));
  THH.PutValue('PHB_CONFIDENTIEL', Salarie.GetValeur(iPSA_CONFIDENTIEL));

  if TR <> nil then
  begin
    THH.PutValue('PHB_ORDREETAT', TR.GetValue('PRM_ORDREETAT'));
    THH.PutValue('PHB_OMTSALARIAL', TR.GetValue('PRM_ORDREETAT')); { PT136 }
  end;
  i := i + 1;
end;

procedure EnleveCommIJ(Salarie, Tob_rub: tob; const Arub, Natrub: string; const dated, datef: tdatetime);
var
  t: tob;
begin
  // PT14 : 23/10/01 : V562 : PH Rajout traitement bulletion complémentaire et dates édition
  if BullCompl = 'X' then exit;

  T := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    if ((copy(T.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.') and
      (T.GetValue('PHB_ORIGINELIGNE') = 'IJS')) then
      T.free;
    T := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;

end;
// f PT129
{$ENDIF}

procedure Align_ExceptRub(NatureRub: string; T: TOB);
var TR: TOB;
begin
  TR := Paie_RechercheRubrique(NatureRub, T.GetValue('PEN_RUBRIQUE'));
  if TR <> nil then
  begin
    if NatureRub = 'AAA' then
    begin
      if Assigned(TR) and (iPRM_IMPRIMABLE = 0) then MemorisePrm(TR);
      TR.PutValeur(iPRM_LIBELLE, T.GetValue('PEN_LIBELLE'));
      TR.PutValeur(iPRM_IMPRIMABLE, T.GetValue('PEN_IMPRIMABLE'));
      TR.PutValeur(iPRM_BASEIMPRIMABLE, T.GetValue('PEN_BASEIMPRIMABLE'));
      TR.PutValeur(iPRM_TAUXIMPRIMABLE, T.GetValue('PEN_TAUXIMPRIMABLE'));
      TR.PutValeur(iPRM_COEFFIMPRIM, T.GetValue('PEN_COEFFIMPRIM'));
      TR.PutValeur(iPRM_DECTAUX, T.GetValue('PEN_DECTAUX'));
      TR.PutValeur(iPRM_DECCOEFF, T.GetValue('PEN_DECCOEFF'));
      TR.PutValeur(iPRM_DECBASE, T.GetValue('PEN_DECBASE'));
    end
    else
    begin
      if Assigned(TR) and (iPCT_IMPRIMABLE = 0) then MemorisePct(TR);
      TR.PutValeur(iPCT_LIBELLE, T.GetValue('PEN_LIBELLE'));
      TR.PutValeur(iPCT_IMPRIMABLE, T.GetValue('PEN_IMPRIMABLE'));
      TR.PutValeur(iPCT_BASEIMP, T.GetValue('PEN_BASEIMPRIMABLE'));
      TR.PutValeur(iPCT_DECBASE, T.GetValue('PEN_DECBASE'));
      TR.PutValeur(iPCT_TXSALIMP, T.GetValue('PEN_TXSALIMP'));
//    TR.PutValeur (iPCT_FFSALIMP, T.GetValue ('PEN_FFSALIMP'));
      TR.PutValeur(iPCT_TXPATIMP, T.GetValue('PEN_TXPATIMP'));
//    TR.PutValeur (iPCT_FFPATIMP , T.GetValue ('PEN_FFPATIMP'));
      TR.PutValeur(iPCT_ORGANISME, T.GetValue('PEN_ORGANISME'));
      TR.PutValeur(iPCT_DECTXSAL, T.GetValue('PEN_DECTXSAL'));
      TR.PutValeur(iPCT_DECMTSAL, T.GetValue('PEN_DECMTSAL'));
      TR.PutValeur(iPCT_DECTXPAT, T.GetValue('PEN_DECTXPAT'));
      TR.PutValeur(iPCT_DECMTPAT, T.GetValue('PEN_DECMTPAT'));
    end;
  end;
end;

procedure Align_Execpt(NatureRub, Rub: string); // Gestion de la personnalisation
var T: TOB;
  Nat: string;
begin
  if (NatureRub = '') and (Rub <> '') then exit;
  if (NatureRub <> '') and (Rub <> '') then
  begin // Cas 1 seule rubrique execption
    T := TOB_Execpt.FindFirst(['PEN_NATURERUB', 'PEN_RUBRIQUE'], [NatureRub, Rub], FALSE);
    if T <> nil then Align_ExceptRub(NatureRub, T);
  end
  else
  begin // toutes les exceptions d'une nature
    if NatureRub <> '' then
    begin
      T := TOB_Execpt.FindFirst(['PEN_NATURERUB'], [NatureRub], FALSE);
      while T <> nil do
      begin
        Align_ExceptRub(NatureRub, T);
        T := TOB_Execpt.FindNext(['PEN_NATURERUB'], [NatureRub], FALSE);
      end;
    end
    else
    begin // toutes les execptions quelque soit la nature
      T := TOB_Execpt.FindFirst([''], [''], FALSE);
      while T <> nil do
      begin
        Align_ExceptRub(T.GetValue('PEN_NATURERUB'), T);
        T := TOB_Execpt.FindNext([''], [''], FALSE);
      end;
    end;
  end;
end;
// DEB PT137

procedure Charg_TauxAt;
var //Q: TQuery;
  St: string;
begin
  FreeAndNIL(TOB_TauxAT);
  TOB_TauxAT := TOB.Create('Les Taux AT', nil, -1);
  st := 'SELECT * FROM TAUXAT ORDER BY PAT_DATEVALIDITE DESC';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_TauxAT.LoadDetailDB('TAUXAT', '', '', Q, FALSE, False);
  Ferme(Q);
}
  TOB_TauxAT.LoadDetailDBFROMSQL('TAUXAT', st);
end;
// FIN PT137
// d PT159
// suppression des rubrique d'IJSS du bulletin

procedure EnleveRubIJS(Salarie, Tob_rub: tob; const dated, datef: tdatetime);
var
  t: tob;
begin
  if BullCompl = 'X' then exit;
  T := Tob_Rub.FindFirst(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    ['IJS', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    T.free;
    T := Tob_Rub.FindNext(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      ['IJS', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;
// suppression des rubrique de maintien (et garantie) du bulletin

procedure EnleveRubMAI(Salarie, Tob_rub: tob; const dated, datef: tdatetime);
var
  t: tob;
begin
  if BullCompl = 'X' then exit;
  T := Tob_Rub.FindFirst(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    ['MAI', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    T.free;
    T := Tob_Rub.FindNext(['PHB_ORIGINELIGNE', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      ['MAI', DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;

// f PT159

// PT198 Mise à jour de la table PRESENCESALARIE lors de la validation de la paie
Procedure MiseAjourPresence(DateD, DateF : TDateTime; Salarie, Typetraitement : string);
var
T ,  TOB_Compteurrubr, T_CPTRUB, TC, TOB_HistoSaisRub : Tob;
DD, DF, Datedebutbul, Datefinbul : TDatetime;
St, rubrique, compteur, typecalpres : string;
Q: TQUERY;
i : integer;

Begin
// Récupération des données saisies en historique
  st := 'SELECT * FROM HISTOSAISRUB LEFT JOIN REMUNERATION ON PRM_RUBRIQUE = PSD_RUBRIQUE AND ##PRM_PREDEFINI## WHERE PRM_THEMEREM <> "ZZZ" AND' +
        ' PSD_SALARIE = "' + Salarie + '"' +
        ' AND PSD_DATEDEBUT >="' + usdatetime(DateD) + // DebutdeMois
        '" AND PSD_DATEFIN <= "' + usdatetime(DateF) + '"' + //  FindeMois
        ' AND (PSD_BASE <> 0 OR PSD_TAUX <> 0 OR PSD_COEFF <> 0 OR PSD_MONTANT <> 0)'+
        ' AND PSD_ORIGINEMVT = "PRE"'+
        ' ORDER BY PSD_ORIGINEMVT,PSD_RUBRIQUE';

  Tob_Histosaisrub  := Tob.create('Historique saisie Rub', nil, -1);
  Tob_HistoSaisRub.LoadDetailDBFROMSQL('HISTOSAISRUB', St);

  if Tob_HistoSaisRub.detail.count <= 0 then
  begin
    If Assigned(Tob_HistoSaisRub) then
      FreeAndNil(Tob_HistoSaisRub);
    exit;
  end;


  if typetraitement = 'MAJ' then
  // Si Mise à jour du bulletin : recherche des compteurs de présence "à intégrer en paie" qui seront mis à jour
  // PGINDICATPRES <== "INP" et DATEDEBUTBUL et DATEFINBUL compris dans les dates de la paie
    St := 'SELECT DISTINCT PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES,'+
          ' PYR_RUBRIQUE, PYP_TYPECALPRES, max(PYR_DATEVALIDITE) DATEVALIDITE FROM PRESENCESALARIE'+
          ' LEFT JOIN COMPTEURPRESENCE ON PYP_COMPTEURPRES = PYR_COMPTEURPRES AND PYR_DATEVALIDITE'+
          ' <= PYP_DATEDEBUTPRES '+
          ' WHERE PYP_PGINDICATPRES = "AIN" AND PYP_DATEDEBUTBUL >= "'+USDATETIME(DateD)+'" AND'+
          ' PYP_DATEFINBUL <= "'+Usdatetime(DateF)+'" AND PYP_SALARIE = "'+Salarie +'" GROUP BY '+
          ' PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES,PYR_RUBRIQUE, PYP_TYPECALPRES'
  else
   // Si Suppression du bulletin : recherche des compteurs de présence " intégré en paie" qui seront mis à jour
  // PGINDICATPRES <== "AIN" et DATEDEBUTBUL et DATEFINBUL = dates de HISTOSAISRUB
    St := 'SELECT DISTINCT PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES,'+
          ' PYR_RUBRIQUE, PYP_TYPECALPRES, max(PYR_DATEVALIDITE) DATEVALIDITE FROM PRESENCESALARIE'+
          ' LEFT JOIN COMPTEURPRESENCE ON PYP_COMPTEURPRES = PYR_COMPTEURPRES AND PYR_DATEVALIDITE'+
          ' <= PYP_DATEDEBUTPRES '+
          ' WHERE PYP_PGINDICATPRES = "INP" AND PYP_DATEDEBUTBUL >= "'+USDATETIME(DateD)+'" AND'+
          ' PYP_DATEFINBUL <= "'+Usdatetime(DATEF)+'" AND PYP_SALARIE = "'+Salarie +'" GROUP BY '+
          ' PYP_SALARIE, PYP_DATEDEBUTPRES, PYP_DATEFINPRES, PYP_COMPTEURPRES,PYR_RUBRIQUE, PYP_TYPECALPRES';
    Q := OPENSQL(st, true);
    TOB_Compteurrubr := TOB.CREATE('Compteurs/ rubrique',nil,-1);

    WHILE Not Q.EOF do
    begin
     TC := Tob.Create('Compteurs/rubriques',TOB_Compteurrubr,-1);
     TC.AddChampSupValeur('COMPTEUR',Q.findfield('PYP_COMPTEURPRES').asstring);
     TC.AddChampSupValeur('RUBRIQUE',Q.findfield('PYR_RUBRIQUE').asstring);
     TC.AddChampSupValeur('DATEDEBUTPRES',Q.findfield('PYP_DATEDEBUTPRES').asDatetime);
     TC.AddChampSupValeur('DATEFINPRES',Q.findfield('PYP_DATEFINPRES').asDatetime);
     TC.AddChampSupValeur('TYPECALPRES',Q.findfield('PYP_TYPECALPRES').asstring);
     Q.next;
    end;

    ferme(Q);


  For i := 0 to  TOB_HistoSaisRub.Detail.Count -1 do
  begin
    T:= TOB_HistoSaisRub.Detail[i];
    DateDebutBul := T.Getvalue('PSD_DATEDEBUT');
    Datefinbul := T.GetValue('PSD_DATEFIN');
    rubrique := T.Getvalue('PSD_RUBRIQUE');
    T_CPTRUB := TOB_Compteurrubr.findfirst(['RUBRIQUE'], [Rubrique], false);
    if T_CPTRUB <> Nil then
    begin
    Compteur := T_CPTRUB.getvalue('COMPTEUR');
    DD := T_CPTRUB.getvalue('DATEDEBUTPRES');
    DF := T_CPTRUB.Getvalue('DATEFINPRES');
    typecalpres := T_CPTRUB.Getvalue('TYPECALPRES');

    if typetraitement = 'MAJ' then    // Mise à jour du bulletin
    // Met à jour le champ PGINDICATPRES = "INP" , datedebutbul et datefinbul = dates du bulletin
    EXECUTESQL('UPDATE PRESENCESALARIE SET PYP_PGINDICATPRES = "INP", PYP_DATEDEBUTBUL = "'+USdatetime(DateD)+'"'+
               ', PYP_DATEFINBUL = "'+UsDatetime(DateF)+'"  WHERE PYP_SALARIE = "'+Salarie+'" AND PYP_COMPTEURPRES'+
               ' = "'+Compteur+'" AND PYP_DATEDEBUTPRES = "'+USdatetime(DD)+'" AND PYP_DATEFINPRES = "'+Usdatetime(DF)+'"'+
               ' AND PYP_DATEDEBUTBUL = "'+Usdatetime(Datedebutbul)+'" AND PYP_DATEFINBUL = "'+Usdatetime(Datefinbul)+'"'+
               ' AND PYP_TYPECALPRES = "'+typecalpres+'" AND PYP_PGINDICATPRES = "AIN"')

    else   // Suppression du bulletin
    // met à jour le champ PGINDICTAPRES = "AIN" , Datedebutbul et datefinbul = dates de HISTOSAISRUB
    EXECUTESQL('UPDATE PRESENCESALARIE SET PYP_PGINDICATPRES = "AIN", PYP_DATEDEBUTBUL = "'+USdatetime(DateDebutbul)+'"'+
               ', PYP_DATEFINBUL = "'+UsDatetime(DateFinbul)+'"  WHERE PYP_SALARIE = "'+Salarie+'" AND PYP_COMPTEURPRES'+
               ' = "'+Compteur+'" AND PYP_DATEDEBUTPRES = "'+USdatetime(DD)+'" AND PYP_DATEFINPRES = "'+Usdatetime(DF)+'"'+
               ' AND PYP_DATEDEBUTBUL = "'+Usdatetime(Dated)+'" AND PYP_DATEFINBUL = "'+Usdatetime(Datef)+'"'+
               ' AND PYP_TYPECALPRES = "'+typecalpres+'" AND PYP_PGINDICATPRES = "INP"');

    end;

  end;

  If Assigned(Tob_CompteurRubr) then freeandnil(Tob_CompteurRubr);
  If Assigned(Tob_HistoSaisRub) then freeandnil(Tob_HistoSAISrub);
end;

end.


{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2001
Modifié le ... : 25/03/2002
Description .. : Unité qui contient l'ensemble des fonctions et procédures
Suite ........ : utilisées pour la DADS
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
{
PT1     : 21/12/2001 VG V571 Ajout de traces dans le .log (si debuglog)
PT2-1   : 03/01/2002 VG V571 Prise en compte de la date d'entrée et de sortie
                             précédente
PT2-2   : 03/01/2002 VG V571 Suppression de la dernière date doublon
PT3     : 10/01/2002 VG V571 Correction du calcul avec date d'entrée et de
                             sortie précédente
PT4-1   : 14/01/2002 VG V571 Correction pour que les périodes d'inactivité ne
                             soient pas prises en compte dans le calcul des
                             périodes d'activité
PT4-2   : 14/01/2002 VG V571 Modification, dans le cahier des charges 6.0 de
                             l'alimentation de la rubrique S41.G01.00.013
                             (Condemploi)
PT5-1   : 15/01/2002 VG V571 On n'allait pas chercher les valeurs qui étaient
                             dans l'historique du salarié mais dans la fiche
                             salarié pour les champs statut professionnel,
                             statut catégoriel, conditions d'emploi, taux
                             d'activité à temps partiel
PT5-2   : 15/01/2002 VG V571 Le dernier historique salarié créait toujours une
                             rupture
PT6-1   : 16/01/2002 VG V571 Prise en compte des bulletins effectués en dehors
                             des périodes d'activité
PT6-2   : 16/01/2002 VG V571 Si le statut professionnel ou le statut catégoriel
                             ne sont pas renseignés, information dans le fichier
                             de contrôle
PT7-1   : 22/01/2002 VG V571 Test sur la date de sortie complété
PT7-2   : 22/01/2002 VG V571 Si la requête des cotisations ne renvoie aucun
                             enregistrement, on ne fait plus de parcours des
                             enregistrements
PT8-1   : 24/01/2002 VG V571 Classement des motif de rupture en 4 niveaux
PT8-2   : 24/01/2002 VG V571 Le programme pouvait faire une periode d'inactivité
                             (somme isolée) et une période d'activité superposée
PT8-3   : 24/01/2002 VG V571 Gestion des entrées et sorties le même jour
PT8-4   : 24/01/2002 VG V571 Chargement du salarié même si il a une date de
                             sortie
PT9-1   : 25/01/2002 VG V571 Pour les salariés n'ayant pas de paye sur la
                             période, récupération des données etablissement et
                             tauxAT
PT9-2   : 25/01/2002 VG V571 Tri des dates de paie
PT9-3   : 25/01/2002 VG V571 Alimentation des organismes sur toutes les périodes
                             d'activité
PT10-1  : 28/01/2002 VG V571 Correction de l'alimentation pour les salariés
                             sortis le premier jour d'une période de paye
PT10-2  : 28/01/2002 VG V571 Si les organismes n'étaient pas saisis, on avait un
                             plantage
PT11-1  : 30/01/2002 VG V571 Correction pour version sous ORACLE : on fait
                             désormais des comparaison avec '' et null
PT11-2  : 30/01/2002 VG V571 Base imposable 1° taux => Tranche 2 et base
                             imposable 2° taux => Tranche 3
PT12    : 04/02/2002 VG V571 Ajout de contrôles non bloquant sur les adresses et
                             les natures d'emploi trop longs
PT13    : 11/02/2002 VG V571 Nouvelle procédure qui écrit l'entête salarié
                             uniquement s'il y a quelque chose à écrire
PT14    : 26/02/2002 VG V571 Test complémentaire sur les taux de travail à temps
                             partiel et taux cas particulier SS : cas ou ils ne
                             sont pas renseignés entrainait un plantage
PT15    : 26/02/2002 VG V571 Modifications pour la génération du fichier BTP
PT16    : 19/03/2002 VG V571 Le taux à temps partiel n'était pas multiplié par
                             100
PT17    : 25/03/2002 VG V571 Modif champs historique salarie
PT18    : 19/06/2002 VG V585 Les indemnités de congés payés n'étaient jamais
                             alimentés par défaut (S44.G01.00.004)
PT19    : 26/06/2002 VG V585 Si le nombre d'heures travaillées et salariées
                             était négatif, on le force à 0
PT20    : 10/07/2002 VG V585 Si le taux de travail à temps partiel est à '', on
                             n'écrit plus l'enregistrement
PT21-1  : 15/07/2002 VG V585 Adaptation cahier des charges V6R02
PT21-2  : 15/07/2002 VG V585 Duplication des honoraires
PT21-3  : 15/07/2002 VG V585 Récupération automatique des montants des sommes
                             isolées
PT21-4  : 15/07/2002 VG V585 Si le salarié ne fait pas partie de la période, on
                             écrit une alerte
PT21-5  : 15/07/2002 VG V585 Si il n'y a pas de coefficient, écriture de 'sans
                             classement conventionnel' au lieu de ''
PT22    : 02/09/2002 VG V585 Prise en compte des salariés relevant du régime SS
                             "Autre" . FQ N°10194
PT23    : 03/09/2002 VG V585 Utilisation de l'historique pour le calcul de la
                             DADS-U conditionné par un paramètre société
PT24    : 18/09/2002 VG V585 Correction sur les dates
PT25    : 07/10/2002 VG V585 On ne crèe l'enregistrement Taux temps partiel que
                             si la condition d'emploi est différente de
                             'Complet' - FQ N° 10263
PT26    : 09/10/2002 VG V585 Remplacement de PHS_DATEAPPLIC par
                             PHS_DATEEVENEMENT
PT27-1  : 11/10/2002 VG V585 Contrôle du N° de sécu : Si pas OK, alors valeurs
                             de repli + message
PT27-2  : 11/10/2002 VG V585 Lorsqu'il y avait une rupture au cours d'une
                             période de paie, les montant de la période de paie
                             étaient inclus dans les 2 périodes DADS-U
PT28    : 14/10/2002 VG V585 Gestion des montants négatifs : Tous les montants
                             négatifs sont ramenés à zéro dans le calcul
PT29    : 21/10/2002 VG V585 Gestion du journal des evenements
PT30    : 30/10/2002 VG V585 Ecriture des détails des frais professionnels que
                             s'il existe un montant pour frais professionnels
PT31    : 13/11/2002 VG V585 Modification du message inscrit dans le .log
                             lorsque les bases sont à 0
PT32    : 14/11/2002 VG V585 Contrôles supplémentaires sur le taux de travail à
                             temps partiel
PT33-1  : 27/11/2002 VG V591 Si le pays de l'adresse du salarié était à null,
                             plantage lors du calcul
PT33-2  : 27/11/2002 VG V591 idem pour le département de naissance
PT33-3  : 29/11/2002 VG V591 idem pour le motif de sortie
PT33-4  : 29/11/2002 VG V591 idem pour le numéro de sécurité sociale
PT34-1  : 11/12/2002 VG V591 Correction de la duplication des honoraires qui ne
                             fonctionnait plus
PT34-2  : 11/12/2002 VG V591 Taux de travail à temps partiel : On ne doit plus
                             enlever la virgule. Des contrôles sont déjà
                             effectués lors de la génération du fichier.
PT35    : 18/12/2002 VG V591 On récupère POG_NUMAFFILIATION pour alimenter le
                             numéro de rattachement
PT36    : 09/01/2003 VG V591 Correction du PT27-2 : si fin de contrat en fin de
                             paie et nouveau contrat en début de paie suivante,
                             la première paie du deuxième contrat n'était pas
                             reprise
PT37-1  : 10/01/2003 VG V591 Cas ou la requête des cotisations ne renvoie que
                             des enregistrements antérieurs à la période en
                             cours
PT37-2  : 10/01/2003 VG V591 On tronque toutes les données à 50 caractères
                             (longueur du champ) pour éviter un plantage
PT38    : 13/01/2003 VG V591 Si on a décoché dans les paramètres société
                             "Utilisation de l'historisation pour le calcul de
                             la DADS-U", alors il ne faut plus du tout lire
                             l'historique salarié
PT39-1  : 14/01/2003 VG V591 Si les revenus d'activités sont négatifs, on force
                             le montant à 0
PT39-2  : 14/01/2003 VG V591 Sous Oracle, Si condition d'emploi, statut
                             professionnel ou statut catégoriel n'étaient pas
                             renseignés dans l'historique du salarié, on avait
                             un plantage lors du calcul - FQ N°10451
PT40    : 15/01/2003 VG V591 Modification pour reprendre les données issues de
                             l'historique salarié lorsqu'il existe des
                             historiques antérieurs au début d'exercice
PT41    : 20/01/2003 VG V591 Plantage du calcul pour un salarié dont l'adresse
                             est à l'étranger
PT42    : 22/01/2003 VG V591 Traitement des cas particuliers ayant plusieurs
                             ruptures dans la même paye
PT43-1  : 28/01/2003 VG V591 Si la donnée saisie ne se trouve pas dans la combo,
                             on force à la valeur par défaut - FQ N°10470
PT43-2  : 28/01/2003 VG V591 Résolution du problème des cumuls et des bases
                             provenant d'un import pour les salariés rentrés
                             dans l'année . La DADS-U ne prenait pas en compte
                             le premier mois pour les cumuls (cumuls de reprise
                             du 01 au 31 même si entrée dans le mois) et ne
                             prenait aucune valeur pour les bases (bases du
                             début d'exercice ISIS II à la fin d'exercice
                             ISIS II)
PT44    : 10/02/2003 VG V591 L'alimentation de la zone Taux de travail à temps
                             partiel provenant de l'historique se faisait avec
                             le champ PHS_TAUXPARTIEL au lieu de
                             PHS_TTAUXPARTIEL
PT45    : 12/02/2003 VG V591 Le libellé d'emploi n'était jamais repris à cause
                             d'un test à l'envers du PT43-1
PT46    : 20/02/2003 VG V_42 Libération de TOB non faite mise en évidence avec
                             memcheck
PT47    : 03/03/2003 VG V_42 Dans le cas ou aucune paie n'existe sur la période,
                             on prend comme fin de période (date de rupture -1),
                             uniquement en calcul global
PT48-1  : 17/03/2003 VG V_42 Pour les apprentis, S41.G01.06.002 = 0 et
                             S41.G01.06.003 absent lors du calcul (en fonction
                             du contrat de travail) - FQ N°10572
PT48-2  : 17/03/2003 VG V_42 Si le régime SS est différent du régime général, la
                             base plafonnée doit être renseignée à 0
                             FQ N°10573
PT48-3  : 18/03/2003 VG V_42 Pour une somme isolée, le nombre d'heures est forcé
                             à 0 - FQ N°10575
PT49    : 28/03/2003 VG V_42 Concordance AT avec régime SS - FQ N°10574
PT50    : 23/04/2003 VG V_42 DADS-U BTP : le salaire moyen est conditionné
PT51    : 12/08/2003 VG V_42 Quand le salarié est en "Contrat d'Apprentissage
                             Sect. Publ.", on force le contrat à "Contrat
                             Apprentissage entr. +10 sal" - FQ N°10698
PT52    : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT52-1  : 16/09/2003 VG V_42 Traitement du cas des ouvriers non qualifiés
PT53    : 01/10/2003 VG V_42 Résolution du problème d'un salarié sortant le
                             premier jour d'une paie. La dernière paie était
                             comptée 2 fois - FQ N°10721
PT54-1  : 06/10/2003 VG V_42 DADS-U BTP et cahier des charges V7R01 - FQ N°10878
PT54-2  : 06/10/2003 VG V_42 Sous oracle, "Erreur de transaction" : calcul
                             annulé lors du calcul - FQ N°10880
PT55    : 08/10/2003 VG V_42 Duplication des honoraires : Adaptation cahier des
                             charges V7R01 - FQ N°10891
PT56    : 10/10/2003 VG V_42 Gestion du regroupement des organismes et
                             utilisation de la coche "Code DADS-U" - FQ N°10892
PT57    : 15/10/2003 VG V_42 Gestion de l'exonération Loi Fillon - FQ N°10908
PT58    : 23/10/2003 VG V_42 Il n'est pas utile de dupliquer l'exercice
PT59    : 14/11/2003 VG V_42 Correction du PT11-1
PT60    : 07/01/2004 VG V_50 Lors du calcul de la DADS-U, deux périodes étaient
                             calculées pour les salariés entrés et sortis le
                             même jour - FQ N°11039
PT61-1  : 25/02/2004 VG V_50 Dans le fichier de log, on mentionne lorsque le
                             motif de sortie n'est pas renseigné - FQ N°11019
PT61-2  : 25/02/2004 VG V_50 Complément du PT43-2 - FQ N°11048
PT62    : 27/02/2004 VG V_50 Le code risque AT doit être en majuscule
                             FQ N°11081
PT63-1  : 04/03/2004 VG V_50 Correction pour prendre en compte en priorité les
                             contrats dans la création des périodes d'activité
                             FQ N°11091
PT63-2  : 04/03/2004 VG V_50 Ajout d'un ajustement pour que les périodes soient
                             au plus juste
PT63-3  : 04/03/2004 VG V_50 Correction pour prendre en compte les paies qui
                             semblent en dehors de périodes d'activité, avant la
                             date de sortie
PT63-4  : 04/03/2004 VG V_50 L'initialisation était déjà faite
PT64    : 05/03/2004 VG V_50 Alimentation de l'organisme de prévoyance
                             FQ N°11094
PT65    : 08/03/2004 VG V_50 Alimentation du code unité d'expression des temps
                             (BTP) par "heure" puisqu'on récupère le cumul 20
                             (heures travaillées) - FQ N°11095
PT66    : 14/04/2004 VG V_50 Complément d'alimentation de la prévoyance
PT67-1  : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT67-2  : 05/07/2004 VG V_50 Optimisation du traitement
PT67-3  : 05/07/2004 VG V_50 Calcul des périodes d'inactivité
PT67-4  : 05/07/2004 VG V_50 Gestion des prud'hommes
PT67-5  : 05/07/2004 VG V_50 Spécificités BTP
PT67-6  : 05/07/2004 VG V_50 On ne pouvait pas saisir 99999 comme taux AT
                             FQ N°11118
PT67-7  : 05/07/2004 VG V_50 Message '98,88' n'est pas une valeur en virgule
                             flottante correcte - FQ N°11270
PT67-8  : 05/07/2004 VG V_50 DADS-U BTP : recupération heures salaries/employeur
                             erronnée - FQ N°11261
PT68    : 06/09/2004 VG V_50 Le premier mois des salariés entrés en cours de
                             mois sous ISIS II n'étaient pas repris. Dorénavant,
                             on repprend les cumuls - FQ N°11261
PT69    : 29/09/2004 VG V_50 préciser les dates d'absence calculés - FQ N°11598
PT70    : 04/10/2004 VG V_50 Modification de l'alimentation pour les DADS-U BTP
                             1 enregistrement par période d'activité
PT71    : 04/10/2004 VG V_50 Complément du PT43-2 - FQ N°11639
PT72    : 11/10/2004 VG V_50 Signaler si le nombre d'heures est à zéro
                             FQ N°11663
PT73    : 12/10/2004 VG V_50 Nombre d'enfants couverts : "Inconnu" à la place de
                             "99" - FQ N°11668
PT74    : 13/10/2004 VG V_50 Alimentation conditionnée des heures rémunérées ou
                             salariées - FQ N°11694
PT75-1  : 21/10/2004 VG V_50 Ajout du code convention collective à 9999
PT75-2  : 21/10/2004 VG V_50 Numéro de période incrémenté de 1 entre chaque
                             période
PT76    : 16/11/2004 VG V_60 Complément du PT43-2 - FQ N°11639
PT77    : 30/11/2004 VG V_60 Correction récupération données prud'homales
                             FQ N°11809
PT78    : 03/12/2004 VG V_60 Manipulation des périodes DADS-U en saisie
                             complémentaire : Erreur de transaction - FQ N°11824
PT79    : 06/01/2005 VG V_60 Le segment S46.G01.02.002 Temps d'arrêt des
                             périodes d'inactivité doit être * 100 - FQ N°11894
PT80    : 10/01/2005 VG V_60 La zone épargne retraire reprend la base des
                             rubriques de cotisation au lieu de la part patronal
                             et salariale - FQ N°11912
PT81    : 12/01/2005 VG V_60 On force les données prud'homales à une valeur par
                             défaut (1 4 1) si la zone n'est pas renseignée au
                             niveau du salarié - FQ N°11808
PT82    : 21/01/2005 VG V_60 Le calcul bouclait dans le cas où il existait un
                             bulletin complémentaire dont la date de début et de
                             fin correspondait à la date de sortie du salarié
                             FQ N°11930
PT83    : 25/01/2005 VG V_60 Cas des paies d'un jour sans contrat de travail
                             FQ N°11927
PT84    : 02/02/2005 VG V_60 "Travail à l'étranger ou frontalier" mal renseigné
                             FQ N°11957
PT85    : 08/03/2005 VG V_60 Signaler si le nombre d'heures est à zéro en raison
                             d'une somme isolée - FQ N°11774
PT86    : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT87    : 11/10/2005 VG V_60 Pour un dossier en PETIT décalage paie, le segment
                             S41.G01.00.009 doit prendre la valeur 01
                             FQ N°12197
PT88    : 17/10/2005 VG V_60 Lorsqu'il existe un contrat, on reprend comme motif
                             de fin de période le motif indiqué dans le
                             "Motif fin de contrat" - FQ N°11933
PT89    : 23/11/2005 VG V_65 Correction du PT88
PT90-1  : 08/12/2005 VG V_65 Cas des apprentis loi 1979 - alimentation de
                             S41.G01.06.002 (base exo brute) avec montant de la
                             rémunération - FQ N°12743
PT90-2  : 08/12/2005 VG V_65 Cas des apprentis loi 1987 - alimentation de
                             S41.G01.00.035.001 avec le montant ayant servi de
                             base au calcul des cotisations - FQ N°12747
PT91    : 13/12/2005 VG V_65 Il faut tenir compte de PHB_SENSBUL pour le calcul
                             des montants issus des rémunérations - FQ N°12752
PT92    : 14/12/2005 VG V_65 Correction PT90-1 - FQ N°12743
PT93    : 05/01/2006 VG V_65 Alimentation durée effective du travail pour IRC si
                             égale à zéro - FQ N°12759
PT94    : 05/01/2006 VG V_65 Pour l'alimentation des conditions contractuelles,
                             on ne prenait en compte que le premier contrat
                             FQ N°12795
PT95    : 10/01/2006 VG V_65 Prise en compte des bulletins complémentaires
                             situés sur le premier jour d'une période. Prend en
                             compte du bulletin complémentaire situé sur le
                             dernier jour d'une période - FQ N°11930 & 12811
PT96    : 10/01/2006 VG V_65 Prise en compte des dates inférieures au 01/01/1900
                             FQ N°12805
PT97-1  : 17/01/2006 VG V_65 Modifications des règles de calcul pour un apprenti
                             loi 1987 - FQ N°12747
PT97-2  : 17/01/2006 VG V_65 Ajout d'un message si le contrat de travail n'est
                             pas '05' mais utilisation d'une base de cotisation
                             paramétrée en exonération '02' - FQ N°12783
PT97-3  : 17/01/2006 VG V_65 Pour l'alimentation des segments S41.G01.00.024 et
                             S41.G01.00.063.001, on considère que les rubriques
                             de rémunération qui sont cochées "Heures chômage
                             partiel" sont des rubriques qui ont des montants
                             positifs mais arrivent dans le bulletin en négatif
                             grâce au champ "PHB_SENSBUL" - FQ N°12831
PT98    : 19/01/2006 SB V_65 FQ 10866 Ajout du predefini
PT99    : 19/01/2006 VG V_65 Correction PT94 : On ne prenait plus en compte les
                             contrats dont le début était antérieur à la période
                             en cours
PT100   : 20/03/2006 VG V_65 Pas de calcul BTP sans calcul de période d'activité
PT101-1 : 13/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                             des erreurs
PT101-2 : 13/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT101-3 : 13/10/2006 VG V_70 On enlève la création des enregistrements à blanc
                             même si les rubriques sont obligatoires
PT101-4 : 13/10/2006 VG V_70 Traitement des VRP multicartes - FQ N°12830
PT101-5 : 13/10/2006 VG V_70 On ne gère le code exo 01 que si le code contrat
                             est à '04' ; on ne gère le code exo 02 que si le
                             code contrat est '05' - FQ N°12881
PT101-6 : 13/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT101-7 : 13/10/2006 VG V_70 Traitement DADS-U complémentaire
PT101-8 : 13/10/2006 VG V_70 Gestion des périodes d'1 jour sans contrat
PT101-9 : 13/10/2006 VG V_70 Affectation des débuts de période par défaut à
                             "008" et des fins de périodes à "098" si fin
                             d'exercice
PT101-10: 13/10/2006 VG V_70 Modification alimentation du nombre d'heures pour
                             la prévoyance
PT101-11: 13/10/2006 VG V_70 Traitement des DADS-U IRCANTEC
PT101-12: 13/10/2006 VG V_70 Alimentation de la base brute CP à l'aide d'un
                             paramétrage - FQ N°12951
PT102-1 : 23/10/2006 VG V_70 Alimentation du code situation familiale BTP avec
                             code 99 si non connue - FQ N°13608
PT102-2 : 23/10/2006 VG V_70 Alimentation de l'exonération VRP uniquement si le
                             montant est différent de 0 - FQ N°13608
PT103-1 : 24/10/2006 VG V_70 Contrôle heures travaillées <= heures payées
                             FQ N°13260
PT104   : 03/11/2006 VG V_70 Alimentation des organismes B pour DADS-U BTP
                             uniquement et inversement - FQ N°13642
PT105   : 06/11/2006 VG V_70 Correction gestion du décalage - FQ N°13655
PT106   : 07/11/2006 VG V_70 Calcul automatique des données DADS-U IRCANTEC
                             FQ N°13644
PT107   : 14/11/2006 VG V_70 Les données de l'établissement n'étaient pas
                             réinitialisées entre chaque calcul - FQ N°13684
PT108   : 16/11/2006 VG V_70 Adaptation cahier des charges V8R04 rectificatif du
                             15/11/2006
PT109   : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                             complémentaire - FQ N°13613
PT110   : 29/11/2006 VG V_70 Suppression du contrôle heures travaillées <=
                             heures payées - FQ N°13260
PT111   : 07/12/2006 VG V_70 Prise en compte de l'historique bulletin pour les
                             extras à qui on ne modifie que les dates entré,
                             sortie, entrée prec, sortie prec - FQ N°13739
PT112-1 : 14/12/2006 VG V_70 Mauvais séquencage des enregistrements
                             établissement dans le cas de préparation des
                             établissements de façon indépendante
PT112-2 : 15/12/2006 VG V_70 Si sortie du salarié le 01/01, seuls les segments
                             S30 sont créés en calcul - FQ N°13776
PT112-3 : 15/12/2006 VG V_70 Ne pas alimenter le code de rattachement si
                             l'organisme n'est pas généré - FQ N°13642
PT113   : 03/01/2007 VG V_72 Test du paramètre tenue euro de la compta au lieu
                             de celui de la paye - FQ N°13799
PT114-1 : 04/01/2007 VG V_72 Suppression de tous les S20 de la table
                             DADSCONTROLE lors du contrôle dossier - FQ N°13811
PT114-2 : 04/01/2007 VG V_72 Mauvaise alimentation du code pays et libellé pays
                             lorsque l'adresse est à l'étranger pour les
                             établissements - FQ N°13812
PT115   : 09/01/2007 VG V_72 Alimentation de la base brute fiscale à 0 pour
                             certains contrats - FQ N°13822
PT116   : 11/01/2007 VG V_72 Suppression des enregistrements ayant le champ
                             exercicedads mal alimenté - FQ N°13827
PT117   : 22/01/2007 VG V_72 Les dossiers en paie décalée annoncaient une fin de
                             déclaration au 30/11 - FQ N°13859
PT118   : 26/01/2007 VG V_72 DADS-U IRCANTEC : Retraite complémentaire non
                             incluse - FQ N°13861
PT119   : 22/02/2007 VG V_72 Duplication des honoraires : mauvaise alimentation
                             du champ EXERCICEDADS
PT120   : 02/03/2007 VG V_72 Alimentation du montant versé par l'employeur pour
                             les inactivités - FQ N°13618
PT121   : 12/03/2007 VG V_72 Mauvaise alimentation du code "Début période
                             d'inactivité" en cas de décalage de paie
                             FQ N°13871
PT122   : 31/05/2007 VG V_72 Ajout de l'alimentation du champ PDE_DADSCDC
PT123   : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                             activité" - FQ N°14568
PT124   : 20/07/2007 FC V_72 FQ 14423 Contrôle valeur régime
PT125   : 30/10/2007 GGU V_80 Gestion des absences "en horaires" (absence d'une
                             journée avec heure de début et de fin)
PT125-1 : 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT125-2 : 20/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14430
PT125-3 : 20/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14432
PT125-4 : 22/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14439
PT125-5 : 30/08/2007 VG V_80 Adaptation cahier des charges V8R05
PT125-6 : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT126   : 12/11/2007 VG V_80 Modification de l'alimentation des revenus
                             d'activité - FQ N°13854
PT127   : 13/11/2007 GGU V_80 Gestion des prévoyances
PT128-1 : 21/11/2007 GGU V_80 FQ 14967 : le format des segments S45.G01.01.002
                             Début période cotisation et S45.G01.01.003 Fin
                             période cotisation n'est pas correct : il doit être
                             sous la forme JJMMAAAA soit 8 caractères numériques
                             --> actuellement on trouve des segments alimentés
                             avec la valeur JJ/MM/AAAA ou JJMM
PT128-2 : 21/11/2007 GGU V_80 FQ 14967 : dans le cas où le montant des bases de
                             alors il ne faudrait pas créer les segments
                             S45.G01.01.006 (actuellement créé en code 02
                             assiette forfaitaire) et S45.G01.01.007.001
                             (actuellement alimenté avec un montant à 0)
PT128-3 : 21/11/2007 GGU V_80 Problème de calcul des bases de cotisations
                             prévoyance : on ne prend qu'une base en compte
PT129-1 : 22/11/2007 VG V_80 Prise en compte des nouveaux champs sur PUBLICOTIS
                             dans le calcul des segments alimentés par une
                             affectations DADS-U - FQ N°14968
PT129-2 : 22/11/2007 VG V_80 Si un enregistrement Détail n'est composé que
                             d'espaces, on ne l'écrit pas - FQ N°14963
PT130-1 : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT130-2 : 28/11/2007 VG V_80 Ecriture correcte des données établissement
                             saisissable sans écrasement des données saisies
PT131   : 28/11/2007 GGU V_80 FQ14967 On récupère le montant de la cotisation de
                             prévoyance au lieu de récupérer le montant de la
                             BASE de cotisation de prévoyance
PT132   : 29/11/2007 FC V_80 Mise en place habilitation
PT133   : 03/12/2007 VG V_80 Alimentation automatique du nouveau champ
                             CODE EXTENSION ALSACE-MOSELLE - FQ N°14945
PT134-1 : 04/12/2007 VG V_80 Prise en compte des salariés à temps partiel avec
                             heures complémentaires pour le calcul des heures
                             exonérées - FQ N°14939
PT134-2 : 04/12/2007 VG V_80 Pas d'écriture de la rémunération heures supp.
                             exonérées si égal à zéro - FQ N°15018
PT135   : 10/12/2007 VG V_80 Certains montants sont avec des décimales -
                             FQ N°13925
PT136   : 10/12/2007 VG V_80 Le paramètre société des heures complémentaires
                             passe de 42 à 44
PT137   : 14/12/2007 VG V_80 Adresse complémentaire passe de 32 à 38
PT138   : 17/12/2007 VG V_80 Adaptation cahier des charges V8R06 - AT dépend du
                             statut professionnel
PT139   : 09/01/2008 VG V_80 Contrôle de cohérence section prud'homale avec
                             section établissement uniquement si différente de
                             '05' - FQ N°15097
PT140   : 11/01/2008 VG V_80 Correction de l'alimentation de la base brute SS
                             dans le cas des apprentis - FQ N°15113
PT141   : 12/01/2008 VG V_80 Correction réduction "Fillon" - FQ N°15104
PT142   : 18/01/2008 VG V_80 Correction sommes isolées - FQ N°15130
PT143   : 22/01/2008 GGU V_80 Correction du cas ou un salarié a plusieurs contrats
                             de prévoyance simultanés - FQ 15139
PT144   : 01/02/2008 VG V_80 Le code décalage, lorsque forcé à '05' pour un
                             salarié, restait forcé pour les salariés suivants
                             FQ N°15194
PT145   : 21/02/2008 VG V_80 En cas de décalage de paie, rupture au 30/11 pour
                             une gestion du mois de décembre (pour les
                             prud'hommes) à part.
PT148   : 04/04/2008 NA V_850 Suppression de la période ; ne pas faire la requête selon les dates de la période si
                              Inactivité
PT151   : 15/07/2008 NA V_80 FQ 15373  Duplication honoraires dans Duplic_un : prendre typed idem procédure duplic
}
unit PgDADSOutils;

interface

uses
     {$IFDEF VER150}
     Variants,
     {$ENDIF}
     UTob,HDebug,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
     hctrls,
     Pgoutils2,
     HEnt1,
     SysUtils,
{$IFNDEF DADSUSEULE}
     P5Util,
{$ENDIF}
     Classes,
     PgDADSCommun,
     UTobDebug,
     paramsoc,
     Controls,
     hmsgbox,
     DateUtils,
     P5Def //MonHabilitation
     ;

Type TCleDADS = record
     Salarie : string;    //Code salarié
     TypeD : string;      //Type de DADS
     Num : integer;       //Numéro de période
     DateDeb : TDateTime; //Date de début de période
     DateFin : TDateTime; //Date de fin de période
     Exercice : string;   //Exercice
end;


Type TParamCalc = record
     Siren : string;             //Siren de l'entreprise, issu de SO_SIRET
     TypeDecla : string;         //Nature de la déclaration, tablette PGDADSNATURE
     TypeD : string;             //Type interne, 001 pour déclaration normale
     Libelle : string;           //Nom de la société, issu de SO_LIBELLE
     Du : string;                //Date de début de la déclaration
     Au : string;                //Date de fin de la déclaration
     Neant : boolean;            //Déclaration néant
     Fraction : string;          //Fraction en cours
     FractionTot : string;       //Nombre total de fraction
     TobEtab : Tob;              //Tob des établissement
     Siege : string;             //Etablissement siège, issu de SO_ETABLISDEFAUT
     Depose : string;            //Etablissement déposant, issu de SO_PGCDEPOSE
     Client : string;            //Numéro de client chez l'emetteur, issu de SO_PGNUMCLIENT
     Siret : string;             //Siret du destinataire, issu de SO_PGSIRETDESTIN
{PT125-5
     Communication : string;     //Communication du CRE, issu de SO_PGCOMMCRE
}     
     Coordonnee : string;        //Coordonnées du destinataire du CRE
{PT125-5
     Civilite : string;          //Civilité de la personne destinataire du CRE, issu de SO_PGCIVILCRE
     Personne : string;          //Personne destinataire du CRE, issu de SO_PGNOMCRE
}     
     Routage : string;           //Routage décompte point retraite, issu de SO_PGROUTAGE
     Routage1 : string;          //Tri n° 1, issu de SO_PGROUTAGE1, tablette PGDADSROUTAGE2
     Routage2 : string;          //Tri n° 2, issu de SO_PGROUTAGE2, tablette PGDADSROUTAGE2
     Caracteristique : string;   //Caractéristiques déclaration, tablette PGCARDADS
end;

Type TypeEtab = record
     Code : string;              //Code établissement, issu de ET_ETABLISSEMENT
     Siret : string;             //Siret, issu de ET_SIRET
     Libelle : string;           //Enseigne, issu de ET_LIBELLE
     Adr_Adresse1 : string;      //ET_ADRESSE1
     Adr_Adresse2 : string;      //ET_ADRESSE2
     Adr_Adresse3 : string;      //ET_ADRESSE3
     Adr_Complement : string;    //Complément d'adresse, issu de ET_ADRESSE2+ET_ADRESSE3
     Adr_Numero : string;        //Numero dans la voie, issu de ET_ADRESSE1
     Adr_NomRue : string;        //Nature et nom de la voie, issu de ET_ADRESSE1
     Adr_CP : string;            //Code postal, issu de ET_CODEPOSTAL
     Adr_Ville : string;         //Bureau distributeur, issu de ET_VILLE
     Adr_PaysCode : string;      //Code pays, issu de ET_PAYS
     Adr_PaysNom : string;       //Nom pays, issu de ET_PAYS
     Effectif : string;          //Effectif au 31/12
     SansSal : string;           //Code établissement sans salarié
     TaxeSal : string;           //Code assujetissement à la taxe sur les salaires, issu de ETB_PRORATATVA
     NAF : string;               //Code NAF, issu de ET_APE
     PrudhSect : string;         //Section prud'hommale, issue de ETB_PRUDHSECT
     PrudhDero : string;         //Section prud'hommale dérogatoire, issue de ETB_PRUDH
     Taxe : string;              //Assujettissement taxe d'apprentissage
     Formation : string;         //Assujettissement participation formation professionnelle
end;

Type DIrcantec = record
     ChampsHisto : string;       //Champs Histo
     TableHisto : string;        //Table Histo
     ChampJoin : string;         //Champ Join
     Prefixe : string;           //Préfixe de la table
end;

procedure ChargeTOBCommun();
procedure LibereTOBCommun();
procedure ChargeTOBDADS();
procedure LibereTOBDADS();
procedure FreeTobDADS();
procedure ChargeTOBSal(Salarie : string);
procedure LibereTOB();
procedure ChargeTOBInact(Salarie : string);
procedure LibereTOBInact();
procedure ChargeTOBHon(OrdreHonor : integer);
procedure LibereTOBHon();
procedure CreeEntete (CleDADS : TCleDADS; MotifDeb, MotifFin, Decal : string);
procedure DeletePeriode(Salarie : string; Ordre : integer);
procedure CreeDetail (CleDADS : TCleDADS; Numseg : integer; Segment, Valeur,
                      Affich : string);
procedure DeleteDetail(Salarie : string; Ordre : integer);
procedure DuplicPrudh(Salarie, Origine : string; Ordre : integer);
procedure DeleteTout(TypeD : string);
procedure CalculElemSal(Salarie : string);
procedure CalculSal (Salarie : string);
function  RemplitPeriode(Salarie : string; TDateRupDeb, TDateRupFin : TOB) : integer;
procedure GereSomIso(T : TOB;Salarie, TypeE : string; Periode : integer;
          var NumOrdre : integer);
procedure GereExo (T : TOB;Champ, Salarie, TypeE : string; Periode : integer;
                   var NumOrdre : integer; BufContrat : string='');
procedure RemplitPrudh (Salarie : string; Periode : integer);
Function RemplitFromPublicotis (Salarie, Num : string; PerDeb, PerFin : TDateTime) : double;
procedure RemplitIRCANTEC (Salarie : string; Periode : integer);
procedure RemplitInact(Salarie : string);
procedure CalculBTP (Salarie : string);
procedure AdresseNormalisee(origine : string; var numero, nom : string);
procedure Duplic (OrdreHonor : integer);
procedure Duplic_un (OrdreHonor : integer; Segment : string);
procedure CalculS20_Entreprise(ParamCalc : TParamCalc);
procedure CalculS80_Etablissement(Etab : TypeEtab);
procedure InitEtab (var Etab : TypeEtab);
procedure MajExercice ();
procedure ChargeEtabNonFictif (TEtabDADS : TOB);


var
TAbs, TAbs1, TConvention, TCumuls, TDADSUE, TEtab, THistoOrg, THonor2 : TOB;
TOrganisme, TRemBTP, TSalCompl, TTauxAT : TOB;

DateDecalage, DebExer2, FinDAnnee, FinExer2 : TDateTime;

BufProf, Decalage, StDateDebut, StPerDeb, StPerFin : string;

NumPer, Periode : integer;

Trace : TStringList;

Const
CDC ='V08R06';

implementation

uses DB
  , PGPopulOutils //PT127
  , Math
;

{$IFDEF DADSUSEULE}
function AncienneteMois(const DateEntree, DateFin: TDateTime): WORD;
var
PremMois, PremAnnee: WORD;
begin
PremMois:= 0;
PremAnnee:= 0;
result:= 0;
if DateEntree = 0 then
   result:= 0
else
   begin
   AglNombreDeMoisComplet(DateEntree, DateFin, PremMois, PremAnnee, result);
   if DateEntree = DateFin then
      result:= 0;
   end;
end;

function AncienneteAnnee(const DateEntree, DateFin: TDateTime): WORD;
begin
result:= (AncienneteMois(DateEntree, DateFin)) div 12;
end;
{$ENDIF}

//PT67-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Chargement des TOBs communes à tous les salariés
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure ChargeTOBCommun();
var
StConvention, StEtab, StOrganisme, StTauxAT : string;
QRechConvention, QRechEtab, QRechOrganisme, QRechTauxAT : TQuery;
begin
//Chargement de la TOB ETABLISSEMENT
StEtab:= 'SELECT ET_ETABLISSEMENT, ET_LIBELLE, ET_SIRET, ET_APE,'+
         ' ETB_HORAIREETABL, ETB_PRUDH, ETB_REGIMEALSACE, ET_FICTIF'+
         ' FROM ETABLISS'+
         ' LEFT JOIN ETABCOMPL ON'+
         ' ET_ETABLISSEMENT=ETB_ETABLISSEMENT';
QRechEtab:=OpenSql(StEtab,TRUE);
TEtab := TOB.Create('Les etablissements', NIL, -1);
TEtab.LoadDetailDB('ETAB','','',QRechEtab,False);
Ferme(QRechEtab);

//Chargement de la TOB TAUXAT
StTauxAT:= 'SELECT PAT_ETABLISSEMENT, PAT_ORDREAT, PAT_DATEVALIDITE,'+
           ' PAT_CODERISQUE, PAT_TAUXAT, PAT_SECTIONAT, PAT_CODEBUREAU'+
           ' FROM TAUXAT ORDER BY PAT_DATEVALIDITE DESC';
QRechTauxAT:=OpenSql(StTauxAT,TRUE);
TTauxAT := TOB.Create('Les Taux AT', NIL, -1);
TTauxAT.LoadDetailDB('AT','','',QRechTauxAT,False);
Ferme(QRechTauxAT);

//Chargement de la TOB Organisme
StOrganisme:= 'SELECT POG_ETABLISSEMENT, POG_INSTITUTION, POG_NUMAFFILIATION,'+
             ' POG_REGROUPEMENT, POG_REGROUPDADSU'+
             ' FROM ORGANISMEPAIE WHERE'+
             ' POG_INSTITUTION<>"" AND'+
             ' POG_INSTITUTION IS NOT null AND'+
             ' POG_INSTITUTION=POG_REGROUPEMENT'+
             ' ORDER BY POG_ETABLISSEMENT, POG_INSTITUTION';
QRechOrganisme:=OpenSql(StOrganisme,TRUE);
TOrganisme := TOB.Create('Les Organismes', NIL, -1);
TOrganisme.LoadDetailDB('ORGANISME','','',QRechOrganisme,False);
Ferme(QRechOrganisme);

//PT86
//Chargement de la TOB Convention
StConvention:= 'SELECT PCV_CONVENTION, PCV_IDCC'+
             ' FROM CONVENTIONCOLL WHERE'+
             ' ##PCV_PREDEFINI##'+
             ' PCV_IDCC<>"" AND'+
             ' PCV_IDCC IS NOT null';
QRechConvention:=OpenSql(StConvention,TRUE);
TConvention := TOB.Create('Les Conventions', NIL, -1);
TConvention.LoadDetailDB('CONVENTION','','',QRechConvention,False);
Ferme(QRechConvention);
//FIN PT86
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Libération des TOBs communes
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibereTOBCommun();
begin
FreeAndNil(TEtab);
FreeAndNil(TTauxAT);
FreeAndNil(TOrganisme);
FreeAndNil(TConvention);    //PT86
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/08/2001
Modifié le ... :   /  /
Description .. : Chargement des TOB destination
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ChargeTOBDADS();
begin
TDADSUE := TOB.Create('Mère DADSU Entête', nil, -1);
TDADSUD := TOB.Create('Mère DADSU Détail', nil, -1);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/08/2001
Modifié le ... :   /  /
Description .. : Libération des TOBs destination utilisées pour la DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibereTOBDADS();
begin
TDADSUE.SetAllModifie (TRUE);
TDADSUE.InsertDB(nil,FALSE);
{PT101-2
FreeAndNil(TDADSUE);
}
TDADSUD.SetAllModifie (TRUE);
TDADSUD.InsertDB(nil,FALSE);
{PT101-2
FreeAndNil(TDADSUD);
}
FreeTobDADS;
//FIN PT101-2
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Free des Tob DADS-U
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure FreeTobDADS();
begin
FreeAndNil (TDADSUE);
FreeAndNil (TDADSUD);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/07/2001
Modifié le ... : 24/08/2001
Description .. : Chargement des TOB necessaires au calcul de la DADSU
Suite ........ : pour un salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ChargeTOBSal(Salarie : string);
var
StBase, StContrat, StCot, StHistoCumSal, StHistoOrg, StHistoSal : String;
StPaieCours, StRem, StRemBTP, StSal : String;
QRechBase, QRechContrat, QRechCot, QRechHistoCumSal, QRechHistoOrg : TQuery;
QRechHistoSal, QRechPaieCours, QRechRem, QRechRemBTP, QRechSal : TQuery;
begin
ChTOBSal(Salarie);

//PT15
//Chargement de la TOB DEPORTSAL
if (GetParamSocSecur('SO_PGBTP', '-') = True) then
   begin
   StSal:='SELECT PSE_SALARIE, PSE_BTP, PSE_BTPADHESION, PSE_BTPANCIENNETE,'+
          ' PSE_BTPASSEDIC, PSE_BTPSALMOYEN, PSE_BTPHORAIRE, PSE_STATUTCOTIS,'+
          ' PSE_BTPPOSITION, PSE_METIERBTP, PSE_BTPAFFILIRCIP'+
          ' FROM DEPORTSAL'+
          ' LEFT JOIN SALARIES ON'+
          ' PSE_SALARIE=PSA_SALARIE WHERE'+
          ' PSE_BTP="X" AND'+
          ' PSA_SALARIE="'+Salarie+'" AND'+
          ' ((PSA_DATEENTREE<="'+UsDateTime(FinExer)+'") OR'+
          ' (PSA_DATEENTREEPREC<="'+UsDateTime(FinExer)+'"))';

   QRechSal:=OpenSql(StSal,TRUE);
   TSalCompl := TOB.Create('Les salaries compl', NIL, -1);
   TSalCompl.LoadDetailDB('SALCOMPL','','',QRechSal,False);
   Ferme(QRechSal);
   end;
//FIN PT15

//Chargement de la TOB PAIEENCOURS
StPaieCours:= 'SELECT PPU_ETABLISSEMENT, PPU_SALARIE, PPU_DATEDEBUT,'+
              ' PPU_DATEFIN, PPU_CBRUTFISCAL, PPU_CNETIMPOSAB, ET_SIRET'+
              ' FROM PAIEENCOURS'+
              ' LEFT JOIN ETABLISS ON'+
              ' ET_ETABLISSEMENT=PPU_ETABLISSEMENT WHERE'+
              ' PPU_SALARIE="'+Salarie+'" AND'+
              ' PPU_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
              ' PPU_DATEDEBUT>="'+UsDateTime(DebExer)+'"'+
              ' ORDER BY PPU_DATEDEBUT';
QRechPaieCours:=OpenSql(StPaieCours,TRUE);
TPaie := TOB.Create('Les Paies', NIL, -1);
TPaie.LoadDetailDB('PAIE','','',QRechPaieCours,False);
Ferme(QRechPaieCours);

//Chargement de la TOB CONTRATTRAVAIL
StContrat:= 'SELECT PCI_SALARIE, PCI_ORDRE, PCI_TYPECONTRAT, PCI_DEBUTCONTRAT,'+
            ' PCI_FINCONTRAT, PCI_MOTIFSORTIE, PMS_DADSU'+
            ' FROM CONTRATTRAVAIL'+
            ' LEFT JOIN MOTIFSORTIEPAY ON'+
            ' PCI_MOTIFSORTIE=PMS_CODE WHERE'+
            ' PCI_SALARIE="'+Salarie+'"';
QRechContrat:=OpenSql(StContrat,TRUE);
TContrat := TOB.Create('Les Contrats', NIL, -1);
TContrat.LoadDetailDB('CONTRAT','','',QRechContrat,False);
Ferme(QRechContrat);

//Chargement de la TOB HISTOCUMSAL
StHistoCumSal:= 'SELECT PHC_SALARIE, PHC_CUMULPAIE, PHC_MONTANT,'+
                ' PHC_DATEDEBUT, PHC_DATEFIN, PHC_REPRISE'+
                ' FROM HISTOCUMSAL WHERE'+
                ' PHC_SALARIE = "'+Salarie+'" AND'+
                ' (PHC_CUMULPAIE="02" OR PHC_CUMULPAIE="09" OR'+
                ' (PHC_CUMULPAIE="13" AND PHC_REPRISE="X") OR'+
                ' (PHC_CUMULPAIE="15" AND PHC_REPRISE="X") OR'+
                ' PHC_CUMULPAIE="20" OR PHC_CUMULPAIE="21" OR'+
                ' PHC_CUMULPAIE="30") AND'+
                ' PHC_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
                ' PHC_DATEDEBUT>="'+UsDateTime(DebExer)+'"'+
                ' ORDER BY PHC_DATEDEBUT';
QRechHistoCumSal:=OpenSql(StHistoCumSal,TRUE);
THistoCumSal := TOB.Create('Les HistoCum', NIL, -1);
THistoCumSal.LoadDetailDB('HISTOCUM','','',QRechHistoCumSal,False);
Ferme(QRechHistoCumSal);

//Chargement de la TOB TAUXAT
{PT67-2
StTauxAT:= 'SELECT PAT_ETABLISSEMENT, PAT_ORDREAT, PAT_DATEVALIDITE,'+
           ' PAT_CODERISQUE, PAT_TAUXAT, PAT_SECTIONAT, PAT_CODEBUREAU'+
           ' FROM TAUXAT ORDER BY PAT_DATEVALIDITE DESC';
QRechTauxAT:=OpenSql(StTauxAT,TRUE);
TTauxAT := TOB.Create('Les Taux AT', NIL, -1);
TTauxAT.LoadDetailDB('AT','','',QRechTauxAT,False);
Ferme(QRechTauxAT);
}

//Chargement de la TOB HISTOBULLETIN/BASES DE COTISATION
StBase:= 'SELECT PHB_ETABLISSEMENT, PHB_SALARIE, PHB_TAUXAT, PHB_BASECOT,'+
         ' PHB_PLAFOND, PHB_TRANCHE1, PHB_TRANCHE2, PHB_TRANCHE3,'+
         ' PHB_MTPATRONAL, PCT_BRUTSS, PCT_PLAFONDSS, PCT_BASECSGCRDS,'+
         ' PCT_BASECRDS, PCT_DADSTOTIMPTSS, PCT_DADSMONTTSS, PCT_DADSEXOBASE,'+
         ' PCT_DADSEPARGNE, PHB_DATEDEBUT, PHB_DATEFIN, PHB_MTSALARIAL'+
         ' FROM HISTOBULLETIN'+
         ' LEFT JOIN COTISATION ON'+
         ' PHB_RUBRIQUE=PCT_RUBRIQUE AND'+
         ' PHB_NATURERUB=PCT_NATURERUB WHERE'+
         ' ##PCT_PREDEFINI## (PCT_BRUTSS="X" OR PCT_PLAFONDSS="X" OR'+
         ' PCT_BASECSGCRDS="X" OR PCT_BASECRDS="X" OR PCT_DADSTOTIMPTSS="X" OR'+
         ' PCT_DADSMONTTSS="X" OR PCT_DADSEPARGNE="X" OR'+
         ' PCT_DADSEXOBASE<>"" OR PCT_DADSEXOCOT <>"") AND'+
         ' PHB_SALARIE="'+Salarie+'" AND'+
         ' PHB_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
         ' PHB_DATEDEBUT>="'+UsDateTime(DebExer)+'" ORDER BY PHB_DATEDEBUT';
QRechBase:=OpenSql(StBase,TRUE);
TBase := TOB.Create('Les Bases', NIL, -1);
TBase.LoadDetailDB('HISTO_BULLETIN','','',QRechBase,False);
Ferme(QRechBase);
//PGVisuUnObjet(TBase, '', '');

//Chargement de la TOB HISTOBULLETIN/REMUNERATION
StRem:= 'SELECT PHB_SALARIE, PHB_DATEDEBUT, PHB_DATEFIN, PHB_BASEREM,'+
        ' PHB_MTREM, PHB_SENSBUL, PRM_AVANTAGENAT, PRM_SOMISOL,'+
        ' PRM_RETSALAIRE, PRM_FRAISPROF, PRM_CHQVACANCE, PRM_IMPOTRETSOURC,'+
        ' PRM_INDEXPATRIAT, PRM_HCHOMPAR, PRM_INDEMINTEMP, PRM_INDEMCP,'+
        ' PRM_BTPSALAIRE'+
        ' FROM HISTOBULLETIN'+
        ' LEFT JOIN REMUNERATION ON'+
        ' PHB_RUBRIQUE=PRM_RUBRIQUE AND'+
        ' PHB_NATURERUB=PRM_NATURERUB WHERE'+
        ' ##PRM_PREDEFINI## (PRM_AVANTAGENAT="N" OR PRM_AVANTAGENAT="L" OR'+
        ' PRM_AVANTAGENAT="V" OR PRM_AVANTAGENAT="A" OR PRM_AVANTAGENAT="T" OR'+
        ' PRM_SOMISOL="01" OR PRM_SOMISOL="03" OR PRM_SOMISOL="04" OR'+
        ' PRM_RETSALAIRE="X" OR PRM_FRAISPROF="X" OR PRM_CHQVACANCE="X" OR'+
        ' PRM_IMPOTRETSOURC="X" OR PRM_INDEXPATRIAT="X" OR PRM_HCHOMPAR="X" OR'+
        ' PRM_INDEMINTEMP="X" OR PRM_INDEMCP="X" OR PRM_BTPSALAIRE = "X") AND'+
        ' PHB_SALARIE="'+Salarie+'" AND'+
        ' PHB_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
        ' PHB_DATEDEBUT>="'+UsDateTime(DebExer)+'" ORDER BY PHB_DATEDEBUT';
QRechRem:=OpenSql(StRem,TRUE);
TRem := TOB.Create('Les Rémunérations', NIL, -1);
TRem.LoadDetailDB('HISTOBULL','','',QRechRem,False);
Ferme(QRechRem);

//PT15
//Chargement de la TOB HISTOBULLETIN/REMUNERATION pour BTP
if (GetParamSocSecur('SO_PGBTP', '-') = True) then
   begin
   StRemBTP:= 'SELECT PHB_SALARIE, PHB_DATEDEBUT, PHB_DATEFIN, PHB_BASEREM,'+
              ' PHB_MTREM, PHB_SENSBUL, PRM_BTPARRET'+
              ' FROM HISTOBULLETIN'+
              ' LEFT JOIN REMUNERATION ON'+
              ' PHB_RUBRIQUE=PRM_RUBRIQUE AND'+
              ' PHB_NATURERUB=PRM_NATURERUB WHERE'+
              ' ##PRM_PREDEFINI## (PRM_BTPARRET>="01" AND'+
              ' PRM_BTPARRET<="08") AND'+
              ' PHB_SALARIE="'+Salarie+'" AND'+
              ' PHB_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
              ' PHB_DATEDEBUT>="'+UsDateTime(DebExer)+'"'+
              ' ORDER BY PHB_SALARIE, PRM_BTPARRET, PHB_DATEDEBUT';
   QRechRemBTP:=OpenSql(StRemBTP,TRUE);
   TRemBTP := TOB.Create('Les Rémunérations BTP', NIL, -1);
   TRemBTP.LoadDetailDB('HISTO_BULLBTP','','',QRechRemBTP,False);
   Ferme(QRechRemBTP);
   end;
//FIN PT15

//Chargement de la TOB HISTOBULLETIN/ORGANISME
StHistoOrg:= 'SELECT PHB_SALARIE, POG_ETABLISSEMENT, POG_INSTITUTION,'+
             ' POG_NUMAFFILIATION, POG_REGROUPEMENT, POG_REGROUPDADSU,'+
             ' PHB_DATEDEBUT, PHB_DATEFIN'+
             ' FROM HISTOBULLETIN'+
             ' LEFT JOIN ORGANISMEPAIE ON'+
             ' PHB_ORGANISME=POG_ORGANISME WHERE'+
             ' PHB_SALARIE="'+Salarie+'" AND'+
             ' POG_INSTITUTION<>" " AND'+
             ' POG_INSTITUTION<>"" AND'+
             ' (PHB_BASECOT<>0 OR PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0) AND'+
             ' PHB_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
             ' PHB_DATEDEBUT>="'+UsDateTime(DebExer)+'"'+
             ' ORDER BY POG_ETABLISSEMENT, POG_INSTITUTION';
QRechHistoOrg:=OpenSql(StHistoOrg,TRUE);
THistoOrg := TOB.Create('Les organismes', NIL, -1);
THistoOrg.LoadDetailDB('HISTOBULLORG','','',QRechHistoOrg,False);
Ferme(QRechHistoOrg);

//Chargement de la TOB HISTOBULLETIN/COTISATION
StCot:= 'SELECT PHB_SALARIE, PHB_DATEDEBUT, PHB_DATEFIN, PHB_BASECOT,'+
        ' PHB_MTPATRONAL, PHB_MTSALARIAL, PCT_DADSEXOCOT, PCT_MAJORATA2,'+
        ' PCT_ALLEGEMENTA2'+
        ' FROM HISTOBULLETIN'+
        ' LEFT JOIN COTISATION ON'+
        ' PHB_RUBRIQUE=PCT_RUBRIQUE AND'+
        ' PHB_NATURERUB=PCT_NATURERUB WHERE'+
        ' ##PCT_PREDEFINI## (PCT_DADSEXOCOT<>"" OR PCT_MAJORATA2="X" OR'+
        ' PCT_ALLEGEMENTA2="X") AND'+
        ' PHB_SALARIE="'+Salarie+'" AND'+
        ' (PHB_BASECOT<>0 OR PHB_MTSALARIAL<>0 OR PHB_MTPATRONAL<>0) AND'+
        ' PHB_DATEFIN<="'+UsDateTime (FinExer)+'" AND'+
        ' PHB_DATEDEBUT>="'+UsDateTime (DebExer)+'" ORDER BY PHB_DATEDEBUT';
QRechCot:= OpenSql(StCot,TRUE);
TCot:= TOB.Create ('Les cotisations', NIL, -1);
TCot.LoadDetailDB ('HISTOBULLCOT','','',QRechCot,False);
Ferme (QRechCot);

//PT17  : 25/03/2002 VG V571 Modif champs historique salarie
//Chargement de la TOB HISTOSALARIE
StHistoSal:= 'SELECT PHS_SALARIE, PHS_DATEEVENEMENT, PHS_DADSPROF,'+
             ' PHS_BDADSPROF, PHS_DADSCAT, PHS_BDADSCAT, PHS_TTAUXPARTIEL,'+
             ' PHS_BTTAUXPARTIEL, PHS_CONDEMPLOI, PHS_BCONDEMPLOI, PHS_PROFIL,'+
             ' PHS_BPROFIL'+
             ' FROM HISTOSALARIE WHERE'+
             ' PHS_SALARIE="'+Salarie+'" ORDER BY PHS_DATEEVENEMENT';
QRechHistoSal:=OpenSql(StHistoSal,TRUE);
THistoSal := TOB.Create('Les historiques', NIL, -1);
THistoSal.LoadDetailDB('HISTOSAL','','',QRechHistoSal,False);
Ferme(QRechHistoSal);

//PT56
//Chargement de la TOB Organisme
{PT67-2
StOrganisme:= 'SELECT POG_ETABLISSEMENT, POG_INSTITUTION, POG_NUMAFFILIATION,'+
             ' POG_REGROUPEMENT, POG_REGROUPDADSU'+
             ' FROM ORGANISMEPAIE WHERE'+
             ' POG_INSTITUTION <> "" AND'+
             ' POG_INSTITUTION IS NOT null AND'+
             ' POG_INSTITUTION=POG_REGROUPEMENT ORDER BY'+
             ' POG_ETABLISSEMENT, POG_INSTITUTION';
QRechOrganisme:=OpenSql(StOrganisme,TRUE);
TOrganisme := TOB.Create('Les Organismes', NIL, -1);
TOrganisme.LoadDetailDB('ORGANISME','','',QRechOrganisme,False);
Ferme(QRechOrganisme);
}
//FIN PT56

//TOB destination
ChargeTOBDADS;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/07/2001
Modifié le ... :   /  /
Description .. : Libération des TOBs utilisées pour la DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibereTOB();
begin
LibTOB;

if (GetParamSocSecur('SO_PGBTP', '-') = True) then
   FreeAndNil(TSalCompl);
FreeAndNil(TPaie);
FreeAndNil(TContrat);
FreeAndNil(THistoCumSal);
FreeAndNil(TBase);
FreeAndNil(TRem);
if (GetParamSocSecur('SO_PGBTP', '-') = True) then
   FreeAndNil(TRemBTP);
FreeAndNil(THistoOrg);
FreeAndNil(TCot);
FreeAndNil(THistoSal);
LibereTOBDADS;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Chargement des TOB necessaires au calcul des périodes
Suite ........ : d'inactivité de la DADSU pour un salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ChargeTOBInact(Salarie : string);
var
StAbs, StAbs1 : String;
QRechAbs, QRechAbs1 : TQuery;
begin
//Chargement de la TOB SALARIES
StAbs1:= 'SELECT PSA_SALARIE, PSA_LIBELLE, PSA_PRENOM'+
         ' FROM SALARIES WHERE'+
         ' PSA_SALARIE =  "'+Salarie+'"';
QRechAbs1:=OpenSql(StAbs1,TRUE);
TAbs1 := TOB.Create('Le salarié', NIL, -1);
TAbs1.LoadDetailDB('SAL','','',QRechAbs1,False);
Ferme(QRechAbs1);

//Chargement de la TOB ABSENCESALARIE/MOTIFABSENCE
StAbs:= 'SELECT PMA_TYPEABS, PMA_JOURHEURE, PMA_RUBRIQUE, PMA_RUBRIQUEJ,'+
        ' PCN_SALARIE, PCN_DATEDEBUTABS, PCN_DATEFINABS, PCN_JOURS, PCN_HEURES'+
        ' FROM ABSENCESALARIE'+
        ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI##'+{ PT98 }
        ' PCN_TYPECONGE=PMA_MOTIFABSENCE WHERE'+
        ' PCN_SALARIE="'+Salarie+'" AND'+
        ' PCN_CODETAPE="P" AND'+
        ' PMA_TYPEABS<>"RTT" AND'+
        ' PMA_TYPEABS<>"" AND'+
        ' PCN_ETATPOSTPAIE<>"NAN"'+
        ' ORDER BY PCN_DATEDEBUTABS, PCN_DATEFINABS';
QRechAbs:=OpenSql(StAbs,TRUE);
TAbs := TOB.Create('Les absences', NIL, -1);
TAbs.LoadDetailDB('ABSENCES','','',QRechAbs,False);
Ferme(QRechAbs);

//TOB destination
ChargeTOBDADS;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Libération des TOBs utilisées pour le calcul des périodes
Suite ........ : d'inactivité de la DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibereTOBInact();
begin
FreeAndNil(TAbs1);
FreeAndNil(TAbs);
LibereTOBDADS;
end;

//PT21-2
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/07/2002
Modifié le ... :   /  /
Description .. : Chargement des TOBs pour la duplication des honoraires
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure ChargeTOBHon(OrdreHonor : integer);
var
StHonor : string;
QRechHonor : TQuery;
begin

//Chargement de la TOB HONORAIRES
StHonor:='SELECT *'+
         ' FROM DADSDETAIL WHERE'+
         ' PDS_SALARIE LIKE "--H%" AND'+
         ' PDS_ORDRE = '+IntToStr(OrdreHonor)+' AND'+
         ' PDS_DATEDEBUT = "'+UsDateTime(DebExer)+'" AND'+
         ' PDS_DATEFIN = "'+UsDateTime(FinExer)+'"'+
         ' ORDER BY PDS_TYPE,PDS_ORDRESEG,PDS_SEGMENT';

QRechHonor:=OpenSql(StHonor,TRUE);
THonor2 := TOB.Create('Les honoraires', NIL, -1);
THonor2.LoadDetailDB('HONOR','','',QRechHonor,False);
Ferme(QRechHonor);

//TOB destination
ChargeTOBDADS;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/07/2002
Modifié le ... :
Description .. : Libération des TOBs utilisées pour la duplication des
Suite ........ : honoraires DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure LibereTOBHon();
begin
FreeAndNil(THonor2);
LibereTOBDADS;
end;
//FIN PT21-2


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/07/2001
Modifié le ... :   /  /
Description .. : Création d'une entête DADSPERIODES dans la TOB
Suite ........ : associée
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CreeEntete (CleDADS : TCleDADS; MotifDeb, MotifFin, Decal : string);
var
TDADSUEntete : Tob;
begin
TDADSUEntete:=TOB.Create ('DADSPERIODES', TDADSUE, -1);
{PT101-6
TDADSUEntete.PutValue('PDE_SALARIE', Salarie);
TDADSUEntete.PutValue('PDE_TYPE', TypeD);
TDADSUEntete.PutValue('PDE_ORDRE', Num);
TDADSUEntete.PutValue('PDE_DATEDEBUT', DateDeb);
TDADSUEntete.PutValue('PDE_DATEFIN', DateFin);
TDADSUEntete.PutValue('PDE_MOTIFDEB', MotifDeb);
TDADSUEntete.PutValue('PDE_MOTIFFIN', MotifFin);
if Trace <> nil then
   Trace.Add (Salarie+' : création de la période '+IntToStr(Num)+' du '+
             DateToStr(DateDeb)+' au '+DateToStr(DateFin));
}
TDADSUEntete.PutValue('PDE_SALARIE', CleDADS.Salarie);
TDADSUEntete.PutValue('PDE_TYPE', CleDADS.TypeD);
TDADSUEntete.PutValue('PDE_ORDRE', CleDADS.Num);
TDADSUEntete.PutValue('PDE_DATEDEBUT', CleDADS.DateDeb);
TDADSUEntete.PutValue('PDE_DATEFIN', CleDADS.DateFin);
TDADSUEntete.PutValue('PDE_EXERCICEDADS', CleDADS.Exercice);
TDADSUEntete.PutValue('PDE_MOTIFDEB', MotifDeb);
TDADSUEntete.PutValue('PDE_MOTIFFIN', MotifFin);
TDADSUEntete.PutValue('PDE_DECALEE', Decal);   //PT101-2
TDADSUEntete.PutValue('PDE_DADSCDC', CDC);     //PT122
if Trace <> nil then
   Trace.Add (CleDADS.Salarie+' : création de la période '+
              IntToStr (CleDADS.Num)+' du '+ DateToStr (CleDADS.DateDeb)+' au '+
              DateToStr (CleDADS.DateFin));
//FIN PT101-6
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 29/08/2001
Modifié le ... :   /  /
Description .. : Suppression des enregistrements période
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure DeletePeriode(Salarie : string; Ordre : integer);
var
StDate, StDelete : string;
begin
{Si ordre = -100, on supprime l'ensemble des périodes d'activité d'un salarié
pour un type donné. Si ordre = -200, on supprime l'ensemble des périodes
d'inactivité d'un salarié pour un type donné.
}
//PT125-6
FinDAnnee:= StrToDate ('31/12/'+Copy (DateToStr (FinExer), 7, 4));
// pt148 if ((TypeD<>'001') and (TypeD<>'201')) then
if ((TypeD<>'001') and (TypeD<>'201') and ((ordre >0) or (ordre = -100)) ) then  // pt148
   StDate:= ' PDE_DATEDEBUT>="'+UsDateTime (DebExer)+'" AND'+
            ' PDE_DATEFIN<="'+UsDateTime (FinExer)+'" AND'
else
   StDate:= '';
//FIN PT125-6

if (Ordre = -100) then
   StDelete := 'DELETE FROM DADSPERIODES WHERE'+
               ' PDE_SALARIE = "'+Salarie+'" AND'+
               ' PDE_TYPE = "'+TypeD+'" AND'+
               ' PDE_ORDRE >= 0 AND'+StDate+
               ' PDE_EXERCICEDADS = "'+PGExercice+'"'
else
if (Ordre = -200) then
   StDelete := 'DELETE FROM DADSPERIODES WHERE'+
               ' PDE_SALARIE= "'+Salarie+'" AND '+
               ' PDE_TYPE = "'+TypeD+'" AND'+
               ' PDE_ORDRE < 0 AND'+StDate+
               ' PDE_EXERCICEDADS = "'+PGExercice+'"'
else
   StDelete := 'DELETE FROM DADSPERIODES WHERE'+
               ' PDE_SALARIE= "'+Salarie+'" AND '+
               ' PDE_TYPE = "'+TypeD+'" AND'+
               ' PDE_ORDRE = '+IntToStr(Ordre)+' AND'+StDate+
               ' PDE_EXERCICEDADS = "'+PGExercice+'"';
ExecuteSQL(StDelete) ;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2001
Modifié le ... :   /  /
Description .. : Création d'un élément DADSDETAIL dans la TOB associée
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CreeDetail (CleDADS : TCleDADS; Numseg : integer; Segment, Valeur,
                      Affich : string);
var
TDADSUDetail : Tob;
begin
//PT129-2
if (Trim (Valeur)='') then
   exit;
//FIN PT129-2   
if (Length (Valeur)>50) then
   Valeur:= Copy (Valeur, 1, 50);
if (Length (Affich)>50) then
   Affich:= Copy (Affich, 1, 50);
TDADSUDetail:= TOB.Create ('DADSDETAIL', TDADSUD, -1);
{PT101-6
TDADSUDetail.PutValue ('PDS_SALARIE', Salarie);
TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
TDADSUDetail.PutValue ('PDS_ORDRE', Num);
TDADSUDetail.PutValue ('PDS_DATEDEBUT', DateDeb);
TDADSUDetail.PutValue ('PDS_DATEFIN', DateFin);
TDADSUDetail.PutValue ('PDS_ORDRESEG', Numseg);
TDADSUDetail.PutValue ('PDS_SEGMENT', Segment);
TDADSUDetail.PutValue ('PDS_DONNEE', Valeur);
TDADSUDetail.PutValue ('PDS_DONNEEAFFICH', Affich);
}
TDADSUDetail.PutValue ('PDS_SALARIE', CleDADS.Salarie);
TDADSUDetail.PutValue ('PDS_TYPE', CleDADS.TypeD);
TDADSUDetail.PutValue ('PDS_ORDRE', CleDADS.Num);
TDADSUDetail.PutValue ('PDS_DATEDEBUT', CleDADS.DateDeb);
TDADSUDetail.PutValue ('PDS_DATEFIN', CleDADS.DateFin);
TDADSUDetail.PutValue ('PDS_EXERCICEDADS', CleDADS.Exercice);
TDADSUDetail.PutValue ('PDS_ORDRESEG', Numseg);
TDADSUDetail.PutValue ('PDS_SEGMENT', Segment);
{PT125-6
TDADSUDetail.PutValue ('PDS_DONNEE', Valeur);
}
TDADSUDetail.PutValue ('PDS_DONNEE', Trim (Valeur));
//FIN PT125-6
TDADSUDetail.PutValue ('PDS_DONNEEAFFICH', Affich);
Debug ('Paie PGI/Sauve DADS-U : '+Segment+' "'+Valeur+'"');
//FIN PT101-6
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 29/08/2001
Modifié le ... :   /  /
Description .. : Suppression des enregistrements détail
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure DeleteDetail(Salarie : string; Ordre : integer);
var
StDate, StDelete : string;
begin
{
Si ordre = -100, on supprime l'ensemble des enregistrements concernant les
périodes d'activité d'un salarié pour un type donné.
Si ordre = -200, on supprime l'ensemble des enregistrements concernant les
périodes d'inactivité d'un salarié pour un type donné.
Si ordre = -300, on supprime l'ensemble des enregistrements d'une même nature
(établissement, société, honoraire) pour un type donné.
Si ordre = -400, on supprime l'ensemble des enregistrements d'une même nature
(établissement, société, honoraire) pour un type donné et l'ordre 1.
}
//PT125-6
FinDAnnee:= StrToDate ('31/12/'+Copy (DateToStr (FinExer), 7, 4));
// pt148 if ((TypeD<>'001') and (TypeD<>'201')) then
if ((TypeD<>'001') and (TypeD<>'201') and ((ordre >0) or (ordre = -100) or (ordre = -300)) ) then   //pt148
   StDate:= ' PDS_DATEDEBUT>="'+UsDateTime (DebExer)+'" AND'+
            ' PDS_DATEFIN<="'+UsDateTime (FinExer)+'" AND'
else
   StDate:= '';
//FIN PT125-6

if (Ordre = -100) then
   StDelete:= 'DELETE FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE= "'+Salarie+'" AND'+
              ' PDS_TYPE = "'+TypeD+'" AND'+
              ' PDS_ORDRE >= 0 AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"'
else
if (Ordre = -200) then
   StDelete:= 'DELETE FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE= "'+Salarie+'" AND'+
              ' PDS_TYPE = "'+TypeD+'" AND'+
              ' PDS_ORDRE < 0 AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"'
//PT112-1
else
if (Ordre = -300) then
   StDelete:= 'DELETE FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE LIKE "'+Salarie+'%" AND'+
              ' PDS_TYPE = "'+TypeD+'" AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"'
//FIN PT112-1
//PT130-2
else
if (Ordre = -400) then
   StDelete:= 'DELETE FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE LIKE "'+Salarie+'%" AND'+
              ' PDS_TYPE = "'+TypeD+'" AND'+
              ' PDS_ORDRE = 1 AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"'
//FIN PT130-2
else
   StDelete:= 'DELETE FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE= "'+Salarie+'" AND'+
              ' PDS_TYPE = "'+TypeD+'" AND'+
              ' PDS_ORDRE = '+IntToStr(Ordre)+' AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"';
ExecuteSQL(StDelete) ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/06/2004
Modifié le ... :   /  /
Description .. : Cette procédure est utilisée pour :
Suite ........ : - dupliquer les éléments prud'hommaux dans les cas
Suite ........ : suivants :
Suite ........ : * en cas de suppression de la dernière période, duplication
Suite ........ : de la période en cours vers la période précédente
Suite ........ : * en cas de création de la dernière période, duplication de la
Suite ........ : période précédente vers la période en cours
Suite ........ : - suppression des éléments prud'hommaux sur la
Suite ........ : période précédente dans le cas de création de la dernière
Suite ........ : période
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure DuplicPrudh(Salarie, Origine : string; Ordre : integer);
var
StCherche, StMAJ : string;
QRech : TQuery;
OrdreRech : integer;
begin
{PT101-7
if (TypeD='001') then
}
if ((TypeD='001') or (TypeD='201')) then
//PT101-7
   begin
//Recherche de toutes les périodes
   StCherche:= 'SELECT * FROM DADSPERIODES WHERE'+
               ' PDE_SALARIE= "'+Salarie+'" AND '+
               ' PDE_TYPE = "'+TypeD+'" AND'+
               ' PDE_ORDRE > 0 AND'+
               ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
               ' ORDER BY PDE_DATEDEBUT DESC';
   QRech:=OpenSql(StCherche,TRUE);
   if (not QRech.EOF) then
      begin
      OrdreRech := QRech.FindField ('PDE_ORDRE').AsInteger;
//Vérification que ma période est la dernière (PDE_DATEDEBUT)
      if (Ordre=OrdreRech) then
//Si oui, je récupère le numéro de la période suivante
         begin
         QRech.Next;
         if (not QRech.EOF) then
            begin
            OrdreRech := QRech.FindField ('PDE_ORDRE').AsInteger;
            if (origine='D') then
               begin         //Suppression
//je mets à jour les enregistrements prud'hommaux
               StCherche:= 'SELECT PDS_SALARIE'+
                           ' FROM DADSDETAIL WHERE'+
                           ' PDS_SALARIE= "'+Salarie+'" AND '+
                           ' PDS_TYPE = "'+TypeD+'" AND'+
                           ' PDS_ORDRE ='+IntToStr(OrdreRech)+' AND'+
                           ' PDS_ORDRESEG>700 AND'+
                           ' PDS_ORDRESEG<=710 AND'+
                           ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               if (ExisteSQL(StCherche)=True) then
                  begin
                  StMAJ:= 'DELETE FROM DADSDETAIL WHERE'+
                          ' PDS_SALARIE= "'+Salarie+'" AND '+
                          ' PDS_TYPE = "'+TypeD+'" AND'+
                          ' PDS_ORDRE ='+IntToStr(OrdreRech)+' AND'+
                          ' PDS_ORDRESEG>700 AND'+
                          ' PDS_ORDRESEG<=710 AND'+
                          ' PDS_EXERCICEDADS = "'+PGExercice+'"';
                  ExecuteSQL(StMAJ);
                  end;
               StMAJ:= 'UPDATE DADSDETAIL SET'+
                       ' PDS_ORDRE='+IntToStr(OrdreRech)+' WHERE'+
                       ' PDS_SALARIE= "'+Salarie+'" AND '+
                       ' PDS_TYPE = "'+TypeD+'" AND'+
                       ' PDS_ORDRE ='+IntToStr(Ordre)+' AND'+
                       ' PDS_ORDRESEG>700 AND'+
                       ' PDS_ORDRESEG<=710 AND'+
                       ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               ExecuteSQL(StMAJ);
               StMAJ:= 'DELETE FROM DADSDETAIL WHERE'+
                       ' PDS_SALARIE= "'+Salarie+'" AND '+
                       ' PDS_TYPE = "'+TypeD+'" AND'+
                       ' PDS_ORDRE ='+IntToStr(Ordre)+' AND'+
                       ' PDS_ORDRESEG>700 AND'+
                       ' PDS_ORDRESEG<=710 AND'+
                       ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               ExecuteSQL(StMAJ);
               end
            else
               begin         //Création
//je mets à jour les enregistrements prud'hommaux
               StCherche:= 'SELECT PDS_SALARIE'+
                           ' FROM DADSDETAIL WHERE'+
                           ' PDS_SALARIE= "'+Salarie+'" AND '+
                           ' PDS_TYPE = "'+TypeD+'" AND'+
                           ' PDS_ORDRE ='+IntToStr(Ordre)+' AND'+
                           ' PDS_ORDRESEG>700 AND'+
                           ' PDS_ORDRESEG<=710 AND'+
                           ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               if (ExisteSQL(StCherche)=True) then
                  begin
                  StMAJ:= 'DELETE FROM DADSDETAIL WHERE'+
                          ' PDS_SALARIE= "'+Salarie+'" AND '+
                          ' PDS_TYPE = "'+TypeD+'" AND'+
                          ' PDS_ORDRE ='+IntToStr(Ordre)+' AND'+
                          ' PDS_ORDRESEG>700 AND'+
                          ' PDS_ORDRESEG<=710 AND'+
                          ' PDS_EXERCICEDADS = "'+PGExercice+'"';
                  ExecuteSQL(StMAJ);
                  end;
               StMAJ:= 'UPDATE DADSDETAIL SET'+
                       ' PDS_ORDRE='+IntToStr(Ordre)+' WHERE'+
                       ' PDS_SALARIE= "'+Salarie+'" AND '+
                       ' PDS_TYPE = "'+TypeD+'" AND'+
                       ' PDS_ORDRE ='+IntToStr(OrdreRech)+' AND'+
                       ' PDS_ORDRESEG>700 AND'+
                       ' PDS_ORDRESEG<=710 AND'+
                       ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               ExecuteSQL(StMAJ);
               StMAJ:= 'DELETE FROM DADSDETAIL WHERE'+
                       ' PDS_SALARIE= "'+Salarie+'" AND '+
                       ' PDS_TYPE = "'+TypeD+'" AND'+
                       ' PDS_ORDRE ='+IntToStr(OrdreRech)+' AND'+
                       ' PDS_ORDRESEG>700 AND'+
                       ' PDS_ORDRESEG<=710 AND'+
                       ' PDS_EXERCICEDADS = "'+PGExercice+'"';
               ExecuteSQL(StMAJ);
               end;
            end;
         end;
      end;
   Ferme(QRech);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Suppression de tous les enregistrements DADS-U
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure DeleteTout(TypeD : string);
var
StDate, StDelete : string;
begin
{On supprime l'ensemble des enregistrements des salariés pour un type donné dans
les tables DADSPERIODES et DADSDETAIL : cas du calcul de tous les salariés avec
multi-sélection}
//PT125-6
FinDAnnee:= StrToDate ('31/12/'+Copy (DateToStr (FinExer), 7, 4));
if ((TypeD<>'001') and (TypeD<>'201')) then
   StDate:= ' PDE_DATEDEBUT>="'+UsDateTime (DebExer)+'" AND'+
            ' PDE_DATEFIN<="'+UsDateTime (FinExer)+'" AND'
else
   StDate:= '';
//FIN PT125-6

StDelete := 'DELETE FROM DADSPERIODES WHERE'+
            ' PDE_SALARIE NOT LIKE "--%" AND'+
            ' PDE_TYPE = "'+TypeD+'" AND'+StDate+
            ' PDE_EXERCICEDADS = "'+PGExercice+'"';
ExecuteSQL(StDelete) ;
StDelete := 'DELETE FROM DADSDETAIL WHERE'+
            ' PDS_SALARIE NOT LIKE "--%" AND'+
            ' PDS_TYPE = "'+TypeD+'" AND'+StDate+
            ' PDS_EXERCICEDADS = "'+PGExercice+'"';
ExecuteSQL(StDelete) ;

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/08/2002
Modifié le ... :   /  /
Description .. : Calcul des éléments de la fiche salarié uniquement pour la
Suite ........ : DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CalculElemSal(Salarie : string);
var
Buffer, CodeIso, Dept, Libelle, nom, numero : string;
TSalD : TOB;
Longueur, ValidSS : integer;
CleDADS : TCleDADS;
ErreurDADSU : TControle;
begin
Erreur := False;

//Recherche du salarié concerné dans la TOB Salariés
TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);

if TSalD <> nil then
   begin
{Ecriture dans le fichier de contrôle du matricule, nom et prénom du salarié
concerné}
{PT101-1
   Writeln(FRapport, '');
   Writeln(FRapport, '');
   Writeln(FRapport, 'Salarié : '+TSalD.GetValue('PSA_SALARIE')+' '+
                     TSalD.GetValue('PSA_LIBELLE')+' '+
                     TSalD.GetValue('PSA_PRENOM'));
}
//Remplissage d'un enregistrement salarié
//Remplissage du détail commun à tout le salarié
//PT101-6
   CleDADS.Salarie:= Salarie;
   CleDADS.TypeD:= TypeD;
   CleDADS.Num:= 0;
   CleDADS.DateDeb:= DebExer;
   CleDADS.DateFin:= FinExer;
   CleDADS.Exercice:= PGExercice;

   ErreurDADSU.Salarie:= Salarie;
   ErreurDADSU.TypeD:= TypeD;
   ErreurDADSU.Num:= 0;
   ErreurDADSU.DateDeb:= DebExer;
   ErreurDADSU.DateFin:= FinExer;
   ErreurDADSU.Exercice:= PGExercice;
//FIN PT101-6

{Si le Numéro de sécu est renseigné : remplissage, sinon, en fonction de la
civilité, renseignement par les valeurs "de repli"}
   if (TSalD.GetValue('PSA_NUMEROSS') = null) then
      ValidSS := 1
   else
      ValidSS:= TestNumeroSS (TSalD.GetValue('PSA_NUMEROSS'),
                              TSalD.GetValue('PSA_SEXE'));
   if ((ValidSS = 0) or (ValidSS = 2)) then
      CreeDetail (CleDADS, 1, 'S30.G01.00.001',
                  Copy(TSalD.GetValue('PSA_NUMEROSS'), 1, 13),
                  Copy(TSalD.GetValue('PSA_NUMEROSS'), 1, 13))
   else
      begin
      if (TSalD.GetValue('PSA_CIVILITE') = 'MR') then
         begin
         CreeDetail (CleDADS, 1, 'S30.G01.00.001', '1999999999999',
                     '1999999999999');
{PT101-1
         EcrireErreur('          Le N° de sécurité Sociale est invalide, il a'+
                      ' été forcé à 1999999999999', Erreur);
}
         ErreurDADSU.Segment:= 'S30.G01.00.001';
         ErreurDADSU.Explication:= 'Le N° de sécurité Sociale est invalide, il'+
                                   ' a été forcé à 1999999999999';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT101-1
         end
      else
         begin
         CreeDetail (CleDADS, 1, 'S30.G01.00.001', '2999999999999',
                     '2999999999999');
{PT101-1
         EcrireErreur('          Le N° de sécurité Sociale est invalide, il a'+
                      ' été forcé à 2999999999999', Erreur);
}
         ErreurDADSU.Segment:= 'S30.G01.00.001';
         ErreurDADSU.Explication:= 'Le N° de sécurité Sociale est invalide, il'+
                                   ' a été forcé à 2999999999999';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT101-1
         end
      end;
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.001');

{Si le nom de jeune fille est renseigné, je le prends comme nom patronymique
sinon, je prends le nom}
   if ((TSalD.GetValue('PSA_NOMJF') = '') or
      (TSalD.GetValue('PSA_NOMJF') = null)) then
      CreeDetail (CleDADS, 2, 'S30.G01.00.002', TSalD.GetValue('PSA_LIBELLE'),
                  TSalD.GetValue('PSA_LIBELLE'))
   else
      begin
      CreeDetail (CleDADS, 2, 'S30.G01.00.002', TSalD.GetValue('PSA_NOMJF'),
                  TSalD.GetValue('PSA_NOMJF'));
      CreeDetail (CleDADS, 4, 'S30.G01.00.004', TSalD.GetValue('PSA_LIBELLE'),
                  TSalD.GetValue('PSA_LIBELLE'));
      end;
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.002');

{Si le prénom n'est pas renseigné, écriture d'une erreur dans le fichier de
contrôle}
   if ((TSalD.GetValue('PSA_PRENOM') = '') or
      (TSalD.GetValue('PSA_PRENOM') = null)) then
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié : Prénom', Erreur)
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.003';
      ErreurDADSU.Explication:= 'Le prénom n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
//FIN PT101-1
   else
      CreeDetail (CleDADS, 3, 'S30.G01.00.003', TSalD.GetValue('PSA_PRENOM'),
                  TSalD.GetValue('PSA_PRENOM'));

   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.003');

   if ((TSalD.GetValue('PSA_SURNOM') <> '') and
      (TSalD.GetValue('PSA_SURNOM') <> null)) then
      CreeDetail (CleDADS, 5, 'S30.G01.00.006', TSalD.GetValue('PSA_SURNOM'),
                  TSalD.GetValue('PSA_SURNOM'));
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.006');

   if (TSalD.GetValue('PSA_CIVILITE') = 'MR') then
      CreeDetail (CleDADS, 6, 'S30.G01.00.007', '01',
                  TSalD.GetValue('PSA_CIVILITE'))
   else
   if (TSalD.GetValue('PSA_CIVILITE') = 'MME') then
      CreeDetail (CleDADS, 6, 'S30.G01.00.007', '02',
                  TSalD.GetValue('PSA_CIVILITE'))
   else
   if (TSalD.GetValue('PSA_CIVILITE') = 'MLE') then
      CreeDetail (CleDADS, 6, 'S30.G01.00.007', '03',
                  TSalD.GetValue('PSA_CIVILITE'));
   if ((TSalD.GetValue('PSA_CIVILITE') = '') or
      (TSalD.GetValue('PSA_CIVILITE') = null)) then
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié : Civilité', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.007';
      ErreurDADSU.Explication:= 'La civilité n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.007');

{Concaténation de adresse2 et adresse3 afin de remplir le complément d'adresse}
   Buffer := '';
   if ((TSalD.GetValue('PSA_ADRESSE2') <> '') and
      (TSalD.GetValue('PSA_ADRESSE2') <> null)) then
      begin
      if ((TSalD.GetValue('PSA_ADRESSE3') <> '') and
         (TSalD.GetValue('PSA_ADRESSE3') <> null)) then
         Buffer:= TSalD.GetValue('PSA_ADRESSE2')+' '+
                  TSalD.GetValue('PSA_ADRESSE3')
      else
         Buffer:= TSalD.GetValue('PSA_ADRESSE2');
      end
   else
      if ((TSalD.GetValue('PSA_ADRESSE3') <> '') and
         (TSalD.GetValue('PSA_ADRESSE3') <> null)) then
         Buffer := TSalD.GetValue('PSA_ADRESSE3');
   if (Buffer <> '') then
      begin
{vérification que la longueur ne dépasse pas 38, sinon, tronquage}
      Longueur:= Length (Buffer);
{PT137
      if (Longueur > 32) then
         begin
         Buffer := Copy(Buffer,1,32);
}
      if (Longueur>38) then
         begin
         Buffer:= Copy (Buffer, 1, 38);
//FIN PT137
{PT101-1
         EcrireErreur('Erreur:   L''adresse est trop longue, vérifiez qu''elle est conforme à :', Erreur);
         EcrireErreur('          Adresse 1 : N° et Rue', Erreur);
         EcrireErreur('          Adresse 2 : Complément (escalier ...)', Erreur);
         EcrireErreur('          Adresse 3 : Complément (Résidence ...)', Erreur);
}
         ErreurDADSU.Segment:= 'S30.G01.00.008.001';
         ErreurDADSU.Explication:= 'L''adresse a été tronquée';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT101-1
         end;
      CreeDetail (CleDADS, 7, 'S30.G01.00.008.001', Buffer, Buffer);
      end;
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.001');

{Adresse1 est coupée en 2 : d'un côté, le numéro, de l'autre côté, le nom}
   numero := '';
   nom := '';
   if ((TSalD.GetValue('PSA_ADRESSE1') <> '') and
      (TSalD.GetValue('PSA_ADRESSE1') <> null)) then
      AdresseNormalisee (TSalD.GetValue('PSA_ADRESSE1'), numero, nom);

   if (numero <> '') then
      CreeDetail (CleDADS, 8, 'S30.G01.00.008.003', numero, numero);
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.003');

   if (nom <> '') then
      begin
{vérification que la longueur ne dépasse pas 26, sinon, tronquage}
      Longueur := Length(nom);
      if (Longueur > 26) then
         begin
         nom := Copy(nom,1,26);
{PT101-1
         EcrireErreur('Erreur:   L''adresse est trop longue, vérifiez qu''elle est conforme à :', Erreur);
         EcrireErreur('          Adresse 1 : N° et Rue', Erreur);
         EcrireErreur('          Adresse 2 : Complément (escalier ...)', Erreur);
         EcrireErreur('          Adresse 3 : Complément (Résidence ...)', Erreur);
}
         ErreurDADSU.Segment:= 'S30.G01.00.008.006';
         ErreurDADSU.Explication:= 'L''adresse a été tronquée';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT101-1
         end;
      CreeDetail (CleDADS, 9, 'S30.G01.00.008.006', nom, nom);
      end;
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.006');

{On ne récupère le code postal que s'il est renseigné et qu'il est numérique sur
5 chiffres}
   if (TSalD.GetValue('PSA_CODEPOSTAL') <> '') then
      begin
      if ((TSalD.GetValue('PSA_PAYS') <> 'FRA') or
         ((TSalD.GetValue('PSA_CODEPOSTAL') > '00000') and
         (TSalD.GetValue('PSA_CODEPOSTAL') < '99999'))) then
         CreeDetail (CleDADS, 11, 'S30.G01.00.008.010',
                     TSalD.GetValue('PSA_CODEPOSTAL'),
                     TSalD.GetValue('PSA_CODEPOSTAL'))
      else
{PT101-1
         EcrireErreur ('Erreur:   Fiche Salarié : Code Postal', Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S30.G01.00.008.010';
         ErreurDADSU.Explication:= 'Le code postal est incorrect';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
      end
   else
{PT101-1
      EcrireErreur ('Erreur:   Fiche Salarié : Code Postal', Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S30.G01.00.008.010';
         ErreurDADSU.Explication:= 'Le code postal n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.010');

{On ne récupère la ville que si elle est renseignée, sinon, message d'erreur}
   if ((TSalD.GetValue('PSA_VILLE') = '') or
      (TSalD.GetValue('PSA_VILLE') = null)) then
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié : Adresse', Erreur)
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.008.012';
      ErreurDADSU.Explication:= 'La ville n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
//FIN PT101-1
   else
      CreeDetail (CleDADS, 12, 'S30.G01.00.008.012',
                  PGUpperCase(TSalD.GetValue('PSA_VILLE')),
                  TSalD.GetValue('PSA_VILLE'));
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.012');

   if ((TSalD.GetValue ('PSA_PAYS')<>'FRA') and
//PT125-1
      (TSalD.GetValue ('PSA_PAYS')<>'GUF') and
      (TSalD.GetValue ('PSA_PAYS')<>'GLP') and
      (TSalD.GetValue ('PSA_PAYS')<>'MCO') and
      (TSalD.GetValue ('PSA_PAYS')<>'MTQ') and
      (TSalD.GetValue ('PSA_PAYS')<>'NCL') and
      (TSalD.GetValue ('PSA_PAYS')<>'PYF') and
      (TSalD.GetValue ('PSA_PAYS')<>'SPM') and
      (TSalD.GetValue ('PSA_PAYS')<>'REU') and
      (TSalD.GetValue ('PSA_PAYS')<>'ATF') and
      (TSalD.GetValue ('PSA_PAYS')<>'WLF') and
      (TSalD.GetValue ('PSA_PAYS')<>'MYT') and
//FIN PT125-1
      (TSalD.GetValue ('PSA_PAYS')<>'') and
      (TSalD.GetValue ('PSA_PAYS')<>null)) then
      begin
      PaysISOLib(TSalD.GetValue('PSA_PAYS'), CodeIso, Libelle);
      CreeDetail (CleDADS, 13, 'S30.G01.00.008.013', CodeIso, CodeIso);
      CreeDetail (CleDADS, 14, 'S30.G01.00.008.014', Libelle, Libelle);
      end;
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.013');
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.008.014');

{On ne récupère que les numériques de la date de naissance}
   Buffer := '';
   if TSalD.GetValue('PSA_DATENAISSANCE') <> null then
      begin
      ForceNumerique(DateToStr(TSalD.GetValue('PSA_DATENAISSANCE')), Buffer);
      CreeDetail (CleDADS, 15, 'S30.G01.00.009', Buffer,
                  DateToStr(TSalD.GetValue('PSA_DATENAISSANCE')));
      end;
   if (Length(Buffer) <> 8) then
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié : Date de naissance', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.009';
      ErreurDADSU.Explication:= 'La date de naissance est erronnée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.009');

{Récupération de la commune de naissance}
   if ((TSalD.GetValue('PSA_COMMUNENAISS') <> '') and
      (TSalD.GetValue('PSA_COMMUNENAISS') <> null)) then
      CreeDetail (CleDADS, 16, 'S30.G01.00.010',
                  TSalD.GetValue('PSA_COMMUNENAISS'),
                  TSalD.GetValue('PSA_COMMUNENAISS'));
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.010');

{Si le pays de naissance est la France, alors, récupération du département et
message d'erreur si la commune de naissance et le département de naissance ne
sont pas renseignés . Si le pays de naissance n'est pas la France, alors
département = '99'}
   if (TSalD.GetValue('PSA_PAYSNAISSANCE') = 'FRA') then
      begin
      Dept := '';
      if ((TSalD.GetValue('PSA_DEPTNAISSANCE') <> '') and
         (TSalD.GetValue('PSA_DEPTNAISSANCE') <> null)) then
         Dept := TSalD.GetValue('PSA_DEPTNAISSANCE');
      if (Dept = '20A') then
         Dept := '2A';
      if (Dept = '20B') then
         Dept := '2B';
      CreeDetail (CleDADS, 17, 'S30.G01.00.011', Dept, Dept);

      if ((TSalD.GetValue('PSA_COMMUNENAISS') = '') or
         (TSalD.GetValue('PSA_COMMUNENAISS') = null)) then
{PT101-1
         EcrireErreur('Erreur:   Fiche Salarié : Commune de naissance', Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S30.G01.00.010';
         ErreurDADSU.Explication:= 'La commune de naissance n''est pas renseignée';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1

      if ((TSalD.GetValue('PSA_DEPTNAISSANCE') = '') or
         (TSalD.GetValue('PSA_DEPTNAISSANCE') = null)) then
{PT101-1
         EcrireErreur ('Erreur:   Fiche Salarié : Département de naissance',
                       Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S30.G01.00.011';
         ErreurDADSU.Explication:= 'Le département de naissance n''est pas'+
                                   ' renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
      end
   else
      CreeDetail (CleDADS, 17, 'S30.G01.00.011', '99', '99');
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.011');

{Si le pays de naissance n'est pas renseigné, création de l'enregistrement avec
'' (champ obligatoire) et erreur}
   if ((TSalD.GetValue('PSA_PAYSNAISSANCE') <> '') and
      (TSalD.GetValue('PSA_PAYSNAISSANCE') <> null)) then
      begin
      if ((TypeD='001') or (TypeD='201')) then
         CreeDetail (CleDADS, 18, 'S30.G01.00.012',
                     RechDom('TTPAYS', TSalD.GetValue('PSA_PAYSNAISSANCE'), FALSE),
                     TSalD.GetValue('PSA_PAYSNAISSANCE'))
      else
         begin
         PaysISOLib(TSalD.GetValue('PSA_PAYSNAISSANCE'), CodeIso, Libelle);
         CreeDetail (CleDADS, 18, 'S30.G01.00.012', CodeIso,
                     TSalD.GetValue('PSA_PAYSNAISSANCE'));
         end;
      end
   else
{PT101-1 - PT101-3
      begin
      CreeDetail (Salarie, TypeD, 0, DebExer, FinExer, 18, 'S30.G01.00.012',
                 '', '');
      EcrireErreur('Erreur:   Fiche Salarié : Pays de naissance', Erreur);
      end;
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.012';
      ErreurDADSU.Explication:= 'Le pays de naissance n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.012');

{Si la nationalité n'est pas renseignée, création de l'enregistrement avec ''
(champ obligatoire) et erreur}
   if ((TSalD.GetValue('PSA_NATIONALITE') <> '') and
      (TSalD.GetValue('PSA_NATIONALITE') <> null)) then
      begin
{PT101-7
      if (TypeD='001') then
}
      if ((TypeD='001') or (TypeD='201')) then
//FIN PT101-7
         CreeDetail (CleDADS, 19, 'S30.G01.00.013',
                     RechDom('TTPAYS', TSalD.GetValue('PSA_NATIONALITE'), FALSE),
                     TSalD.GetValue('PSA_NATIONALITE'))
      else
         begin
         PaysISOLib(TSalD.GetValue('PSA_NATIONALITE'), CodeIso, Libelle);
         CreeDetail (CleDADS, 19, 'S30.G01.00.013', CodeIso,
                     TSalD.GetValue('PSA_NATIONALITE'));
         end;
      end
   else
{PT101-1 - PT101-3
      begin
      CreeDetail (Salarie, TypeD, 0, DebExer, FinExer, 19, 'S30.G01.00.013',
                 '', '');
      EcrireErreur('Erreur:   Fiche Salarié : Nationalité', Erreur);
      end;
}
      begin
      ErreurDADSU.Segment:= 'S30.G01.00.013';
      ErreurDADSU.Explication:= 'La nationalité n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   Debug('Paie PGI/Calcul DADS-U : S30.G01.00.013');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/07/2001
Modifié le ... :   /  /
Description .. : Calcul des éléments d'un salarié pour la DADSU
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CalculSal(Salarie : string);
var
Etab : string;
TContratD, TDADSUDetail, TDate, TDateDeb, TDateFille, TDateFin : TOB;
TDateInact, TDateInactD, TDateInactF, THistoSalD, TPaieD, TSalD : TOB;
Boucle, NbreDate, PeriodeFaite : integer;
PaieDansContrat : boolean;
begin
PeriodeFaite:=0;
//Recherche du salarié concerné dans la TOB Salariés
TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);

if TSalD <> nil then
   begin
//Calcul des dates de ruptures salariés
   TDate := TOB.Create('Les Dates', NIL, -1);

{
Si salarié sorti avant le début d'exercice, alors
   DEBUT1
   rupture car bulletin    :
           début d'exercice,  NIV=D, ME=095, MS=096, PO=A, DD=0, DF=0
   rupture car bulletin    :
           fin d'exercice,    NIV=D, ME=095, MS=096, PO=Z, DD=0, DF=0
   FIN1
sinon
   DEBUT2
   si entrée précédente dans l'exercice, alors
      DEBUT3
      rupture                 :
           début d'exercice,  NIV=D, ME=095,         PO=Z, DD=0, DF=0
      rupture car entrée prec :
           entrée précédente, NIV=B, ME=001,         PO=A, DD=0, DF=-1
      si sortie précédente avant la fin d'exercice, alors
         DEBUT4
         rupture car sortie prec :
           sortie précédente, NIV=C,                 PO=Z, DD=1, DF=0
         rupture                 :
           fin d'exercice,    NIV=D,         MS=096, PO=Z, DD=0, DF=0
         FIN4
      FIN3
   sinon
      DEBUT5
      si sortie précédente dans l'exercice, alors
         DEBUT6
         rupture                 :
           début d'exercice,  NIV=C, ME=097,         PO=A, DD=0, DF=0
         rupture car sortie prec :
           sortie précédente, NIV=C,                 PO=Z, DD=1, DF=0
         rupture                 :
           fin d'exercice,    NIV=D,         MS=096, PO=Z, DD=0, DF=0
         FIN6
      FIN5
   si salarié entré dans l'exercice, alors
      DEBUT7
      rupture                 :
           début d'exercice,  NIV=D, ME=095,         PO=Z, DD=0, DF=0
      rupture car entrée      :
           entrée,            NIV=B, ME=001,         PO=A, DD=0, DF=-1
      FIN7
   sinon
      si salarié entré entre la date de fin d'exercice et la fin d'année
         DEBUT8B
         rupture car entrée      :
              entrée,            NIV=B, ME=001, MS=008, PO=A, DD=0, DF=-1
         FIN8B
      sinon
         DEBUT8
         rupture                 :
              début d'exercice,  NIV=D, ME=097,         PO=A, DD=0, DF=0
         FIN8
   si salarié sorti dans l'exercice, alors
      DEBUT9
      rupture car sortie      :
           sortie           , NIV=C,         MS=008 ou 010 ou 012 ou 016, PO=Z, DD=1, DF=0
      rupture                 :
           fin d'exercice,    NIV=D,         MS=096, PO=Z, DD=0, DF=0
      FIN9
   sinon
      si salarié entré entre la date de fin d'exercice et la fin d'année
         DEBUT9B
         rupture car fin d'année :
              fin d'année,       NIV=C,         MS=098, PO=Z, DD=0, DF=0
         FIN9B
      DEBUT10
      rupture                 :
           fin d'exercice,    NIV=C,         MS=098, PO=Z, DD=0, DF=0
      FIN10
   si changement d'établissement, alors
      DEBUT11
      rupture établissement   :
           chgt etab     ,    NIV=E, ME=019, MS=020, PO=A, DD=0, DF=-1
      FIN11
   si début contrat dans l'exercice, alors
      DEBUT12
      rupture contrat         :
           début contrat ,    NIV=A, ME=001, MS=008, PO=A, DD=0, DF=-1
      FIN12
   sinon
      si début contrat antérieur à début d'exercice et
         fin contrat postèrieur à début d'exercice, alors
         DEBUT12B
         rupture                 :
              début d'exercice,  NIV=A, ME=097,         PO=A, DD=0, DF=0
         FIN12B
   si fin contrat dans l'exercice, alors
      DEBUT13
      rupture contrat         :
           fin contrat   ,    NIV=A, ME=001, MS=008, PO=Z, DD=1, DF=0
      FIN13
   si historique salarié=condition emploi, alors
      DEBUT14
      rupture histo salarié   :
           historique sal,    NIV=E, ME=021, MS=022, PO=A, DD=0, DF=-1
      FIN14
   si historique salarié=statut professionnel, alors
      DEBUT15
      rupture histo salarié   :
           historique sal,    NIV=E, ME=023, MS=024, PO=A, DD=0, DF=-1
      FIN15
   si historique salarié=statut catégoriel, alors
      DEBUT16
      rupture histo salarié   :
           historique sal,    NIV=E, ME=025, MS=026, PO=A, DD=0, DF=-1
      FIN16
   si historique salarié=taux de travail à temps partiel, alors
      DEBUT17
      rupture histo salarié   :
           historique sal,    NIV=E, ME=031, MS=032, PO=A, DD=0, DF=-1
      FIN17
   si historique salarié=profil, alors
      DEBUT18
      rupture histo salarié   :
           historique sal,    NIV=E, ME=021, MS=022, PO=A, DD=0, DF=-1
      FIN18
   FIN2
}
{PT96
   if ((TSalD.GetValue('PSA_DATESORTIE') < DebExer) AND
      (TSalD.GetValue('PSA_DATESORTIE') <> IDate1900) AND
      (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
}
   if ((TSalD.GetValue('PSA_DATESORTIE') < DebExer) AND
      (TSalD.GetValue('PSA_DATESORTIE') > IDate1900) AND
      (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
      begin
//DEBUT1
      TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
      While ((TPaieD <> nil) and
            ((TPaieD.GetValue('PPU_DATEDEBUT')<DebExer) or
            (TPaieD.GetValue('PPU_DATEDEBUT')>FinExer))) do
            TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
      if ((TPaieD <> nil) and
         (TPaieD.GetValue('PPU_DATEDEBUT')>=DebExer) and
         (TPaieD.GetValue('PPU_DATEDEBUT')<=FinExer)) then
         begin
         TDateFille := TOB.Create('', TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', DebExer);
         TDateFille.PutValue('NIVEAU', 'D');
         TDateFille.PutValue('MOTIFENTREE', '095');
         TDateFille.PutValue('MOTIFSORTIE', '096');
         TDateFille.PutValue('PERIODEOK', 'A');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture début exercice  : '+
               DateToStr(DebExer)+',  NIV=C, ME=095, MS=096, PO=A');

         TDateFille := TOB.Create('', TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', FinExer);
         TDateFille.PutValue('NIVEAU', 'D');
         TDateFille.PutValue('MOTIFENTREE', '095');
         TDateFille.PutValue('MOTIFSORTIE', '096');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
               DateToStr(FinExer)+',    NIV=C, ME=095, MS=096, PO=Z');
         end;
//FIN1
      end
   else
      begin
//DEBUT2
      if ((TSalD.GetValue('PSA_DATEENTREEPREC') >= DebExer) and
         (TSalD.GetValue('PSA_DATEENTREEPREC') <= FinExer)) then
         begin
//DEBUT3
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', DebExer);
         TDateFille.PutValue('NIVEAU', 'D');
         TDateFille.PutValue('MOTIFENTREE', '095');
         TDateFille.PutValue('MOTIFSORTIE', '000');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture début exercice  : '+
               DateToStr(DebExer)+',  NIV=C, ME=095,         PO=Z');

//PT111
         TPaieD:= TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
         While ((TPaieD <> nil) and
               ((TPaieD.GetValue('PPU_DATEDEBUT')<DebExer) or
               (TPaieD.GetValue('PPU_DATEDEBUT')>FinExer))) do
               TPaieD:= TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
         While ((TPaieD <> nil) and
               (TPaieD.GetValue('PPU_DATEDEBUT')<TSalD.GetValue('PSA_DATEENTREEPREC'))) do
               begin
               PaieDansContrat:= False;
               TContratD:= TContrat.FindFirst(['PCI_SALARIE'], [Salarie], TRUE);
               While ((TContratD <> nil) and
                     (TContratD.GetValue('PCI_SALARIE') = Salarie)) do
                     begin
                     if (((TPaieD.GetValue('PPU_DATEDEBUT')>=TContratD.GetValue('PCI_DEBUTCONTRAT')) and
                        (TPaieD.GetValue('PPU_DATEDEBUT')<=TContratD.GetValue ('PCI_FINCONTRAT'))) or
                        ((TContratD.GetValue('PCI_DEBUTCONTRAT')>=TPaieD.GetValue('PPU_DATEDEBUT')) and
                        (TContratD.GetValue ('PCI_DEBUTCONTRAT')<=TPaieD.GetValue('PPU_DATEFIN')))) then
                        PaieDansContrat:= True;
                     TContratD:= TContrat.FindNext(['PCI_SALARIE'], [Salarie], TRUE);
                     end;

               if (PaieDansContrat=False) then
                  begin
                  TDateFille:= TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', TPaieD.GetValue('PPU_DATEDEBUT'));
                  if (TPaieD.GetValue('PPU_DATEDEBUT')=DebExer) then
                     TDateFille.PutValue('NIVEAU', 'B')
                  else
                     TDateFille.PutValue('NIVEAU', 'D');
                  TDateFille.PutValue('MOTIFENTREE', '095');
                  TDateFille.PutValue('MOTIFSORTIE', '096');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture début bulletin  : '+
                        DateToStr(DebExer)+',  NIV=D, ME=095, MS=096, PO=A');

                  TDateFille:= TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', TPaieD.GetValue('PPU_DATEFIN'));
                  TDateFille.PutValue('NIVEAU', 'D');
                  TDateFille.PutValue('MOTIFENTREE', '095');
                  TDateFille.PutValue('MOTIFSORTIE', '096');
                  TDateFille.PutValue('PERIODEOK', 'Z');
                  TDateFille.PutValue('AJUDEB', 1);
                  TDateFille.PutValue('AJUFIN', 0);
                  Debug ('Paie PGI/Calcul DADS-U : rupture fin bulletin    : '+
                        DateToStr(DebExer)+',  NIV=D, ME=095, MS=096, PO=Z');
                  end;
               TPaieD:= TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
               end;
//FIN PT111

         TDateFille:= TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATEENTREEPREC'));
         TDateFille.PutValue('NIVEAU', 'B');
         TDateFille.PutValue('MOTIFENTREE', '001');
         TDateFille.PutValue('MOTIFSORTIE', '000');
         TDateFille.PutValue('PERIODEOK', 'A');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', -1);
         Debug ('Paie PGI/Calcul DADS-U : rupture entrée prec.    : '+
               DateToStr(TSalD.GetValue('PSA_DATEENTREEPREC'))+
               ',  NIV=A, ME=001,         PO=A');
{PT96
         if ((TSalD.GetValue('PSA_DATESORTIEPREC') <= FinExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> IDate1900) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> null)) then
}
         if ((TSalD.GetValue('PSA_DATESORTIEPREC') <= FinExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') > IDate1900) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> null)) then
            begin
//DEBUT4
            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
            TDateFille.PutValue('DATE', FinExer);
            TDateFille.PutValue('NIVEAU', 'D');
            TDateFille.PutValue('MOTIFSORTIE', '096');
            TDateFille.PutValue('PERIODEOK', 'Z');
            TDateFille.PutValue('AJUDEB', 0);
            TDateFille.PutValue('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
                  DateToStr(FinExer)+',  NIV=B,         MS=096, PO=Z');
            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
{PT101-8
            TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIEPREC'));
}
{PT112-2
            if (TSalD.GetValue('PSA_DATEENTREEPREC')=TSalD.GetValue('PSA_DATESORTIEPREC')) then
}
            if ((TSalD.GetValue('PSA_DATEENTREEPREC')=TSalD.GetValue('PSA_DATESORTIEPREC')) or
               (DebExer=TSalD.GetValue('PSA_DATESORTIEPREC'))) then
//FIN PT112-2
               TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIEPREC')+1)
            else
               TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIEPREC'));
            TDateFille.PutValue('NIVEAU', 'C');
{PT101-9
            TDateFille.PutValue('MOTIFSORTIE', '000');
}
            TDateFille.PutValue('MOTIFSORTIE', '008');
            TDateFille.PutValue('PERIODEOK', 'Z');
            TDateFille.PutValue('AJUDEB', 1);
{PT101-8
            TDateFille.PutValue('AJUFIN', 0);
}
{PT112-2
            if (TSalD.GetValue('PSA_DATEENTREEPREC')=TSalD.GetValue('PSA_DATESORTIEPREC')) then
}
            if ((TSalD.GetValue('PSA_DATEENTREEPREC')=TSalD.GetValue('PSA_DATESORTIEPREC')) or
               (DebExer=TSalD.GetValue('PSA_DATESORTIEPREC'))) then
//FIN PT112-2
               TDateFille.PutValue('AJUFIN', -1)
            else
               TDateFille.PutValue('AJUFIN', 0);
//FIN PT101-8
            Debug ('Paie PGI/Calcul DADS-U : rupture sortie prec.    : '+
                  DateToStr(TSalD.GetValue('PSA_DATESORTIEPREC'))+',  NIV=C,'+
                  '                 PO=Z');
{PT101-1
            EcrireErreur ('Alerte:   Saisie complémentaire DADS-U :'+
                          ' Motif de sortie à saisir si pas de contrat', Erreur);
}
//FIN4
            end;
//FIN3
         end
      else
         begin
//DEBUT5
{PT96
         if ((TSalD.GetValue('PSA_DATESORTIEPREC') <= FinExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') >= DebExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> IDate1900) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> null)) then
}
         if ((TSalD.GetValue('PSA_DATESORTIEPREC') <= FinExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') >= DebExer) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') > IDate1900) AND
            (TSalD.GetValue('PSA_DATESORTIEPREC') <> null)) then
            begin
//DEBUT6
            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
            TDateFille.PutValue('DATE', DebExer);
            TDateFille.PutValue('NIVEAU', 'C');
            TDateFille.PutValue('MOTIFENTREE', '097');
            TDateFille.PutValue('MOTIFSORTIE', '000');
            TDateFille.PutValue('PERIODEOK', 'A');
            TDateFille.PutValue('AJUDEB', 0);
            TDateFille.PutValue('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture début exercice  : '+
                  DateToStr(DebExer)+',  NIV=B, ME=097,         PO=A');

            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
            TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIEPREC'));
            TDateFille.PutValue('NIVEAU', 'C');
{PT101-9
            TDateFille.PutValue('MOTIFSORTIE', '000');
}
            TDateFille.PutValue('MOTIFSORTIE', '008');
            TDateFille.PutValue('PERIODEOK', 'Z');
            TDateFille.PutValue('AJUDEB', 1);
            TDateFille.PutValue('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture sortie prec.    : '+
                  DateToStr(TSalD.GetValue('PSA_DATESORTIEPREC'))+',  NIV=B,'+
                  '                 PO=Z');
{PT101-1
            EcrireErreur ('Erreur:   Saisie complémentaire DADS-U :'+
                          ' Motif de sortie à saisir si pas de contrat', Erreur);
}
            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
            TDateFille.PutValue('DATE', FinExer);
            TDateFille.PutValue('NIVEAU', 'D');
            TDateFille.PutValue('MOTIFSORTIE', '096');
            TDateFille.PutValue('PERIODEOK', 'Z');
            TDateFille.PutValue('AJUDEB', 0);
            TDateFille.PutValue('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
                  DateToStr(FinExer)+',  NIV=C,         MS=096, PO=Z');
//FIN6
            end;
//FIN5
         end;

      if ((TSalD.GetValue('PSA_DATEENTREE') >= DebExer) and
         (TSalD.GetValue('PSA_DATEENTREE') <= FinExer)) then
         begin
//DEBUT7
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', DebExer);
         TDateFille.PutValue('NIVEAU', 'D');
         TDateFille.PutValue('MOTIFENTREE', '095');
         TDateFille.PutValue('MOTIFSORTIE', '000');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture début exercice  : '+
               DateToStr(DebExer)+',  NIV=C, ME=095,         PO=Z');
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATEENTREE'));
         TDateFille.PutValue('NIVEAU', 'B');
         TDateFille.PutValue('MOTIFENTREE', '001');
{PT101-9
         TDateFille.PutValue('MOTIFSORTIE', '000');
}
         TDateFille.PutValue('MOTIFSORTIE', '008');
         TDateFille.PutValue('PERIODEOK', 'A');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', -1);
         Debug ('Paie PGI/Calcul DADS-U : rupture entrée          : '+
               DateToStr(TSalD.GetValue('PSA_DATEENTREE'))+',  NIV=A, ME=001,'+
               '         PO=A');
//FIN7
         end
      else
         begin
//PT125-6
         if ((TSalD.GetValue ('PSA_DATEENTREE')>=FinExer) and
            (TSalD.GetValue ('PSA_DATEENTREE')<=FinDAnnee) and
            ((TypeD='001') or (TypeD='201'))) then
            begin
//DEBUT7B
            TDateFille:= TOB.Create ('', TDate, -1);
            TDateFille.AddChampSup ('DATE', FALSE);
            TDateFille.AddChampSup ('NIVEAU', FALSE);
            TDateFille.AddChampSup ('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup ('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup ('PERIODEOK', FALSE);
            TDateFille.AddChampSup ('AJUDEB', FALSE);
            TDateFille.AddChampSup ('AJUFIN', FALSE);
            TDateFille.PutValue ('DATE', TSalD.GetValue ('PSA_DATEENTREE'));
            TDateFille.PutValue ('NIVEAU', 'B');
            TDateFille.PutValue ('MOTIFENTREE', '001');
            TDateFille.PutValue ('MOTIFSORTIE', '008');
            TDateFille.PutValue ('PERIODEOK', 'A');
            TDateFille.PutValue ('AJUDEB', 0);
            TDateFille.PutValue ('AJUFIN', -1);
            Debug ('Paie PGI/Calcul DADS-U : rupture entrée          : '+
                   DateToStr (TSalD.GetValue ('PSA_DATEENTREE'))+',  NIV=A,'+
                   ' ME=001,         PO=A');
//FIN7B
            end
         else
            begin
//DEBUT8
            TDateFille := TOB.Create('',TDate,-1);
            TDateFille.AddChampSup('DATE', FALSE);
            TDateFille.AddChampSup('NIVEAU', FALSE);
            TDateFille.AddChampSup('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup('PERIODEOK', FALSE);
            TDateFille.AddChampSup('AJUDEB', FALSE);
            TDateFille.AddChampSup('AJUFIN', FALSE);
            TDateFille.PutValue('DATE', DebExer);
            TDateFille.PutValue('NIVEAU', 'D');
            TDateFille.PutValue('MOTIFENTREE', '097');
            TDateFille.PutValue('MOTIFSORTIE', '000');
            TDateFille.PutValue('PERIODEOK', 'A');
            TDateFille.PutValue('AJUDEB', 0);
            TDateFille.PutValue('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture début exercice  : '+
                  DateToStr(DebExer)+',  NIV=C, ME=097,         PO=A');
            end;
//FIN PT125-6
//FIN8
         end;

{PT96
      if ((TSalD.GetValue('PSA_DATESORTIE') <= FinExer) AND
         (TSalD.GetValue('PSA_DATESORTIE') >= DebExer) AND
         (TSalD.GetValue('PSA_DATESORTIE') <> IDate1900) AND
         (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
}
      if ((TSalD.GetValue('PSA_DATESORTIE') <= FinExer) AND
         (TSalD.GetValue('PSA_DATESORTIE') >= DebExer) AND
         (TSalD.GetValue('PSA_DATESORTIE') > IDate1900) AND
         (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
         begin
//DEBUT9
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
{PT101-8
         TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIE'));
}
{PT112-2
         if (TSalD.GetValue('PSA_DATEENTREE')=TSalD.GetValue('PSA_DATESORTIE')) then
}
         if ((TSalD.GetValue('PSA_DATEENTREE')=TSalD.GetValue('PSA_DATESORTIE')) or
            (DebExer=TSalD.GetValue('PSA_DATESORTIE'))) then
//FIN PT112-2
            TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIE')+1)
         else
            TDateFille.PutValue('DATE', TSalD.GetValue('PSA_DATESORTIE'));
//FIN PT101-8
         TDateFille.PutValue('NIVEAU', 'C');
         if ((TSalD.GetValue('PSA_MOTIFSORTIE') <> '') and
            (TSalD.GetValue('PSA_MOTIFSORTIE') <> null)) then
            begin
            case StrToInt(TSalD.GetValue('PSA_MOTIFSORTIE')) of
                 31..37, 81 : TDateFille.PutValue('MOTIFSORTIE', '008');
                 59 : TDateFille.PutValue('MOTIFSORTIE', '010');
                 11..20 : TDateFille.PutValue('MOTIFSORTIE', '012');
                 38, 39 : TDateFille.PutValue('MOTIFSORTIE', '016');
                 403 : TDateFille.PutValue('MOTIFSORTIE', '018');
            else
{PT101-9
                 TDateFille.PutValue('MOTIFSORTIE', '999');
}
                 TDateFille.PutValue('MOTIFSORTIE', '008');
                 end;
            end
         else
{PT101-9
            begin
            TDateFille.PutValue('MOTIFSORTIE', '000');
            EcrireErreur ('Erreur:   Fiche Salarié :'+
                          ' Motif de sortie à saisir si pas de contrat', Erreur);
            end;
}
            TDateFille.PutValue('MOTIFSORTIE', '008');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 1);
{PT101-8
         TDateFille.PutValue('AJUFIN', 0);
}
{PT112-2
         if (TSalD.GetValue('PSA_DATEENTREE')=TSalD.GetValue('PSA_DATESORTIE')) then
}
         if ((TSalD.GetValue('PSA_DATEENTREE')=TSalD.GetValue('PSA_DATESORTIE')) or
            (DebExer=TSalD.GetValue('PSA_DATESORTIE'))) then
//FIN PT112-2
            TDateFille.PutValue('AJUFIN', -1)
         else
            TDateFille.PutValue('AJUFIN', 0);
//FIN PT101-8
         Debug ('Paie PGI/Calcul DADS-U : rupture sortie          : '+
               DateToStr(TSalD.GetValue('PSA_DATESORTIE'))+',  NIV=B,'+
               '         MS=008 ou 010 ou 012 ou 016, PO=Z');
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', FinExer);
         TDateFille.PutValue('NIVEAU', 'D');
         TDateFille.PutValue('MOTIFSORTIE', '096');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
               DateToStr(FinExer)+',  NIV=C,         MS=096, PO=Z');
//FIN9
         end
      else
         begin
         if ((TSalD.GetValue ('PSA_DATEENTREE')>FinExer) and
            (TSalD.GetValue ('PSA_DATEENTREE')<=FinDAnnee) and
            ((TypeD='001') or (TypeD='201'))) then
            begin
//DEBUT9B
            TDateFille:= TOB.Create ('',TDate,-1);
            TDateFille.AddChampSup ('DATE', FALSE);
            TDateFille.AddChampSup ('NIVEAU', FALSE);
            TDateFille.AddChampSup ('MOTIFENTREE', FALSE);
            TDateFille.AddChampSup ('MOTIFSORTIE', FALSE);
            TDateFille.AddChampSup ('PERIODEOK', FALSE);
            TDateFille.AddChampSup ('AJUDEB', FALSE);
            TDateFille.AddChampSup ('AJUFIN', FALSE);
            TDateFille.PutValue ('DATE', FinDAnnee);
            TDateFille.PutValue ('NIVEAU', 'C');
            TDateFille.PutValue ('MOTIFSORTIE', '098');
            TDateFille.PutValue ('PERIODEOK', 'Z');
            TDateFille.PutValue ('AJUDEB', 0);
            TDateFille.PutValue ('AJUFIN', 0);
            Debug ('Paie PGI/Calcul DADS-U : rupture fin année       : '+
                   DateToStr (FinDAnnee)+',  NIV=B,         MS=098, PO=Z');
//FIN9B
            end;
{PT145
         else
}
//FIN PT125-6
{PT145
            begin
}
//DEBUT10
         TDateFille := TOB.Create('',TDate,-1);
         TDateFille.AddChampSup('DATE', FALSE);
         TDateFille.AddChampSup('NIVEAU', FALSE);
         TDateFille.AddChampSup('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup('PERIODEOK', FALSE);
         TDateFille.AddChampSup('AJUDEB', FALSE);
         TDateFille.AddChampSup('AJUFIN', FALSE);
         TDateFille.PutValue('DATE', FinExer);
         TDateFille.PutValue('NIVEAU', 'C');
         TDateFille.PutValue('MOTIFSORTIE', '098');
         TDateFille.PutValue('PERIODEOK', 'Z');
         TDateFille.PutValue('AJUDEB', 0);
         TDateFille.PutValue('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
                DateToStr(FinExer)+',  NIV=B,         MS=098, PO=Z');
//FIN10
{PT145
            end;
}
         end;

      TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
      if (TPaieD <> nil) then
         begin
         Etab := TPaieD.GetValue('PPU_ETABLISSEMENT');
         While ((TPaieD <> nil) and
               (TPaieD.GetValue('PPU_SALARIE') = Salarie)) do
               begin
               if (Etab <> TPaieD.GetValue('PPU_ETABLISSEMENT')) then
                  begin
//DEBUT11
                  TDateFille := TOB.Create('',TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', TPaieD.GetValue('PPU_DATEDEBUT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '019');
                  TDateFille.PutValue('MOTIFSORTIE', '020');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture début bulletin  : '+
                        DateToStr(TPaieD.GetValue('PPU_DATEDEBUT'))+',  NIV=D,'+
                        ' ME=019, MS=020, PO=A');
                  Etab := TPaieD.GetValue('PPU_ETABLISSEMENT');
//FIN11
                  end;
               TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
               end;
         end;

      TContratD :=TContrat.FindFirst(['PCI_SALARIE'], [Salarie], TRUE);
      While ((TContratD <> nil) and
            (TContratD.GetValue('PCI_SALARIE') = Salarie)) do
            begin
            if ((TContratD.GetValue('PCI_DEBUTCONTRAT') >= DebExer) and
               (TContratD.GetValue('PCI_DEBUTCONTRAT') <= FinExer)) then
               begin
//DEBUT12
               TDateFille := TOB.Create('',TDate,-1);
               TDateFille.AddChampSup('DATE', FALSE);
               TDateFille.AddChampSup('NIVEAU', FALSE);
               TDateFille.AddChampSup('MOTIFENTREE', FALSE);
               TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
               TDateFille.AddChampSup('PERIODEOK', FALSE);
               TDateFille.AddChampSup('AJUDEB', FALSE);
               TDateFille.AddChampSup('AJUFIN', FALSE);
               TDateFille.PutValue('DATE', TContratD.GetValue('PCI_DEBUTCONTRAT'));
               TDateFille.PutValue('NIVEAU', 'A');
               TDateFille.PutValue('MOTIFENTREE', '001');
               TDateFille.PutValue('MOTIFSORTIE', '008');
               TDateFille.PutValue('PERIODEOK', 'A');
               TDateFille.PutValue('AJUDEB', 0);
               TDateFille.PutValue('AJUFIN', -1);
               Debug ('Paie PGI/Calcul DADS-U : rupture début contrat   : '+
                     DateToStr(TContratD.GetValue('PCI_DEBUTCONTRAT'))+',  NIV=D,'+
                     ' ME=001, MS=008, PO=A');
//FIN12
               end
            else
               begin
               if ((TContratD.GetValue('PCI_DEBUTCONTRAT') < DebExer) and
                  (TContratD.GetValue('PCI_FINCONTRAT') >= DebExer)) then
                  begin
//DEBUT12B
                  TDateFille := TOB.Create('',TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', DebExer);
                  TDateFille.PutValue('NIVEAU', 'A');
                  TDateFille.PutValue('MOTIFENTREE', '097');
                  TDateFille.PutValue('MOTIFSORTIE', '000');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', 0);
                  Debug ('Paie PGI/Calcul DADS-U : rupture début exercice   : '+
                        DateToStr(DebExer)+',  NIV=A, ME=097, PO=A');
//FIN12B
                  end;
               end;

            if ((TContratD.GetValue('PCI_FINCONTRAT') >= DebExer) and
               (TContratD.GetValue('PCI_FINCONTRAT') <= FinExer)) then
               begin
//DEBUT13
               TDateFille := TOB.Create('',TDate,-1);
               TDateFille.AddChampSup('DATE', FALSE);
               TDateFille.AddChampSup('NIVEAU', FALSE);
               TDateFille.AddChampSup('MOTIFENTREE', FALSE);
               TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
               TDateFille.AddChampSup('PERIODEOK', FALSE);
               TDateFille.AddChampSup('AJUDEB', FALSE);
               TDateFille.AddChampSup('AJUFIN', FALSE);
{PT101-8
               TDateFille.PutValue('DATE', TContratD.GetValue('PCI_FINCONTRAT'));
}
{PT112-2
               if (TContratD.GetValue('PCI_DEBUTCONTRAT')=TContratD.GetValue('PCI_FINCONTRAT')) then
}
               if ((TContratD.GetValue('PCI_DEBUTCONTRAT')=TContratD.GetValue('PCI_FINCONTRAT')) or
                  (DebExer=TContratD.GetValue('PCI_FINCONTRAT'))) then
//FIN PT112-2
                  TDateFille.PutValue ('DATE',
                                       TContratD.GetValue ('PCI_FINCONTRAT')+1)
               else
                  TDateFille.PutValue ('DATE',
                                       TContratD.GetValue ('PCI_FINCONTRAT'));
//FIN PT101-8
               TDateFille.PutValue('NIVEAU', 'A');
               TDateFille.PutValue('MOTIFENTREE', '001');
{PT88
               TDateFille.PutValue('MOTIFSORTIE', '008');
}
               if ((TContratD.GetValue('PMS_DADSU')<> null) and
{PT89
                  (TContratD.GetValue('PMS_DADSU')<>'000')) then
}
                  (TContratD.GetValue('PMS_DADSU')<>'')) then
                  TDateFille.PutValue('MOTIFSORTIE', TContratD.GetValue('PMS_DADSU'))
               else
                  TDateFille.PutValue('MOTIFSORTIE', '008');
//FIN PT88
               TDateFille.PutValue('PERIODEOK', 'Z');
               TDateFille.PutValue('AJUDEB', 1);
{PT101-8
               TDateFille.PutValue('AJUFIN', 0);
}
{PT112-2
               if (TContratD.GetValue('PCI_DEBUTCONTRAT')=TContratD.GetValue('PCI_FINCONTRAT')) then
}
               if ((TContratD.GetValue('PCI_DEBUTCONTRAT')=TContratD.GetValue('PCI_FINCONTRAT')) or
                  (DebExer=TContratD.GetValue('PCI_FINCONTRAT'))) then
//FIN PT112-2
                  TDateFille.PutValue('AJUFIN', -1)
               else
                  TDateFille.PutValue('AJUFIN', 0);
//FIN PT101-8
               Debug ('Paie PGI/Calcul DADS-U : rupture fin contrat     : '+
                     DateToStr(TContratD.GetValue('PCI_FINCONTRAT'))+',  NIV=D,'+
                     ' ME=001, MS=008, PO=Z');
//FIN13
               end;
            TContratD :=TContrat.FindNext(['PCI_SALARIE'], [Salarie], TRUE);
            end;

      if (GetParamSocSecur('SO_PGHDADSHISTO', '-') = True) then
         begin
         THistoSalD := THistoSal.FindFirst(['PHS_SALARIE', 'PHS_BCONDEMPLOI'],
                                           [Salarie, 'X'], TRUE);
         While (THistoSalD <> nil) do
               begin
               if ((THistoSalD.GetValue('PHS_DATEEVENEMENT') >= DebExer) and
                  (THistoSalD.GetValue('PHS_DATEEVENEMENT') <= FinExer)) then
                  begin
//DEBUT14
                  TDateFille := TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', THistoSalD.GetValue('PHS_DATEEVENEMENT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '021');
                  TDateFille.PutValue('MOTIFSORTIE', '022');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture car. activité   : '+
                        DateToStr(THistoSalD.GetValue('PHS_DATEEVENEMENT'))+','+
                        '  NIV=D, ME=021, MS=022, PO=A');
//FIN14
                  end;
               THistoSalD := THistoSal.FindNext(['PHS_SALARIE', 'PHS_BCONDEMPLOI'],
                                                [Salarie, 'X'], TRUE);
               end;

         THistoSalD := THistoSal.FindFirst(['PHS_SALARIE', 'PHS_BDADSPROF'],
                                           [Salarie, 'X'], TRUE);
         While (THistoSalD <> nil) do
               begin
               if ((THistoSalD.GetValue('PHS_DATEEVENEMENT') >= DebExer) and
                  (THistoSalD.GetValue('PHS_DATEEVENEMENT') <= FinExer)) then
                  begin
//DEBUT15
                  TDateFille := TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', THistoSalD.GetValue('PHS_DATEEVENEMENT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '023');
                  TDateFille.PutValue('MOTIFSORTIE', '024');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture statut prof.    : '+
                        DateToStr(THistoSalD.GetValue('PHS_DATEEVENEMENT'))+','+
                        '  NIV=D, ME=023, MS=024, PO=A');
//FIN15
                  end;
               THistoSalD := THistoSal.FindNext(['PHS_SALARIE', 'PHS_BDADSPROF'],
                                                [Salarie, 'X'], TRUE);
               end;

         THistoSalD := THistoSal.FindFirst(['PHS_SALARIE', 'PHS_BDADSCAT'],
                                           [Salarie, 'X'], TRUE);
         While (THistoSalD <> nil) do
               begin
               if ((THistoSalD.GetValue('PHS_DATEEVENEMENT') >= DebExer) and
                  (THistoSalD.GetValue('PHS_DATEEVENEMENT') <= FinExer)) then
                  begin
//DEBUT16
                  TDateFille := TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', THistoSalD.GetValue('PHS_DATEEVENEMENT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '025');
                  TDateFille.PutValue('MOTIFSORTIE', '026');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture statut categ.   : '+
                        DateToStr(THistoSalD.GetValue('PHS_DATEEVENEMENT'))+','+
                        '  NIV=D, ME=025, MS=026, PO=A');
//FIN16
                  end;
               THistoSalD := THistoSal.FindNext(['PHS_SALARIE', 'PHS_BDADSCAT'],
                                                [Salarie, 'X'], TRUE);
               end;

         THistoSalD := THistoSal.FindFirst(['PHS_SALARIE', 'PHS_BTTAUXPARTIEL'],
                                           [Salarie, 'X'], TRUE);
         While (THistoSalD <> nil) do
               begin
               if ((THistoSalD.GetValue('PHS_DATEEVENEMENT') >= DebExer) and
                  (THistoSalD.GetValue('PHS_DATEEVENEMENT') <= FinExer)) then
                  begin
//DEBUT17
                  TDateFille := TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', THistoSalD.GetValue('PHS_DATEEVENEMENT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '031');
                  TDateFille.PutValue('MOTIFSORTIE', '032');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture taux partiel    : '+
                        DateToStr(THistoSalD.GetValue('PHS_DATEEVENEMENT'))+','+
                        '  NIV=D, ME=031, MS=032, PO=A');
//FIN17
                  end;
               THistoSalD:= THistoSal.FindNext(['PHS_SALARIE','PHS_BTTAUXPARTIEL'],
                                                [Salarie, 'X'], TRUE);
               end;

         THistoSalD := THistoSal.FindFirst(['PHS_SALARIE', 'PHS_BPROFIL'],
                                           [Salarie, 'X'], TRUE);
         While (THistoSalD <> nil) do
               begin
               if ((THistoSalD.GetValue('PHS_DATEEVENEMENT') >= DebExer) and
                  (THistoSalD.GetValue('PHS_DATEEVENEMENT') <= FinExer)) then
                  begin
//DEBUT18
                  TDateFille := TOB.Create('', TDate,-1);
                  TDateFille.AddChampSup('DATE', FALSE);
                  TDateFille.AddChampSup('NIVEAU', FALSE);
                  TDateFille.AddChampSup('MOTIFENTREE', FALSE);
                  TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateFille.AddChampSup('PERIODEOK', FALSE);
                  TDateFille.AddChampSup('AJUDEB', FALSE);
                  TDateFille.AddChampSup('AJUFIN', FALSE);
                  TDateFille.PutValue('DATE', THistoSalD.GetValue('PHS_DATEEVENEMENT'));
                  TDateFille.PutValue('NIVEAU', 'E');
                  TDateFille.PutValue('MOTIFENTREE', '021');
                  TDateFille.PutValue('MOTIFSORTIE', '022');
                  TDateFille.PutValue('PERIODEOK', 'A');
                  TDateFille.PutValue('AJUDEB', 0);
                  TDateFille.PutValue('AJUFIN', -1);
                  Debug ('Paie PGI/Calcul DADS-U : rupture profil          : '+
                        DateToStr(THistoSalD.GetValue('PHS_DATEEVENEMENT'))+','+
                        '  NIV=D, ME=021, MS=022, PO=A');
//FIN18
                  end;
               THistoSalD := THistoSal.FindNext(['PHS_SALARIE', 'PHS_BPROFIL'],
                                                [Salarie, 'X'], TRUE);
               end;
         end;

// Classement des dates dans un ordre chronologique
      TDate.Detail.Sort('DATE;NIVEAU;MOTIFENTREE;MOTIFSORTIE;PERIODEOK');

//Comptage du nombre de période
      NbreDate := TDate.FillesCount(1);

{
if V_PGI.Debug=True then
   PGVisuUnObjet(TDate, '', '');
}

// Suppression des dates doublon de niveau supérieur à A ou B
      TDateDeb:= TDate.FindFirst ([''],[''],FALSE);
      TDateFin:= TDate.FindNext ([''],[''],FALSE);
      NumPer:= 1;
      While (NumPer <> NbreDate) do
            begin
            if TDateDeb.GetValue ('DATE')=TDateFin.GetValue ('DATE') then
               begin
               if (TDateFin.GetValue ('NIVEAU')>=TDateDeb.GetValue ('NIVEAU')) then
                  begin
                  TDate.Detail[NumPer].Free;
                  TDateFin:= TDate.FindNext ([''],[''],FALSE);
                  NbreDate:= NbreDate-1;
                  end
               else
                  begin
                  NumPer:= NumPer+1;
                  TDateDeb:= TDateFin;
                  TDateFin:= TDate.FindNext ([''],[''],FALSE);
                  end;
               end
            else
               begin
               NumPer:= NumPer+1;
               TDateDeb:= TDateFin;
               TDateFin:= TDate.FindNext ([''],[''],FALSE);
               end;
            end;
      if (TDateDeb.GetValue ('PERIODEOK')='A') then
         begin
         TDateFille:= TOB.Create ('',TDate,-1);
         TDateFille.AddChampSup ('DATE', FALSE);
         TDateFille.AddChampSup ('NIVEAU', FALSE);
         TDateFille.AddChampSup ('MOTIFENTREE', FALSE);
         TDateFille.AddChampSup ('MOTIFSORTIE', FALSE);
         TDateFille.AddChampSup ('PERIODEOK', FALSE);
         TDateFille.AddChampSup ('AJUDEB', FALSE);
         TDateFille.AddChampSup ('AJUFIN', FALSE);
         TDateFille.PutValue ('DATE', FinExer);
         TDateFille.PutValue ('NIVEAU', 'D');
{PT101-9
         TDateFille.PutValue('MOTIFSORTIE', '999');
}
         TDateFille.PutValue ('MOTIFSORTIE', '098');
         TDateFille.PutValue ('PERIODEOK', 'Z');
         TDateFille.PutValue ('AJUDEB', 0);
         TDateFille.PutValue ('AJUFIN', 0);
         Debug ('Paie PGI/Calcul DADS-U : rupture fin exercice    : '+
                DateToStr(FinExer)+',  NIV=B,         MS=999, PO=Z');
         end;
//FIN2         
      end;

//TobDebug (TDate);
//Comptage du nombre de période
   NbreDate := TDate.FillesCount(1);

{Boucle sur les périodes du salarié et appel de la fonction RemplitPeriode si
PO=A ou s'il s'agit d'une somme isolée
Cas d'une somme isolée : création des périodes basées sur les périodes de paie
(date début de paie => date fin de paie)}
   TDateDeb := TDate.FindFirst([''],[''],FALSE);
   TDateFin := TDate.FindNext([''],[''],FALSE);
   NumPer := 1;
   Periode:= 0;
   Boucle := 0;
   While (NumPer <> NbreDate) and (TDateFin<>nil) do
         begin
         if (TDateDeb.GetValue('PERIODEOK') = 'A') then
            begin
{PT109
            PeriodeFaite:= RemplitPeriode (Salarie, TypeD, TDateDeb, TDateFin);
}
            PeriodeFaite:= RemplitPeriode (Salarie, TDateDeb, TDateFin);
{PT100
            if (TypeD <> '001') then
}
{PT101-7
            if ((TypeD <> '001') and (PeriodeFaite <> 0)) then
}
            if ((TypeD <> '001') and (TypeD <> '201') and
               (PeriodeFaite <> 0)) then
//FIN PT101-7
{PT109
               CalculBTP (Salarie, TypeD);
}
               CalculBTP (Salarie);
            end
         else
            begin
            TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
            While ((TPaieD <> nil) and
                  ((TPaieD.GetValue('PPU_DATEDEBUT')<TDateDeb.GetValue('DATE')) or
                  (TPaieD.GetValue('PPU_DATEDEBUT')>TDateFin.GetValue('DATE')))) do
                  TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
            While ((TPaieD <> nil) and
                  (TPaieD.GetValue('PPU_DATEDEBUT')>TDateDeb.GetValue('DATE')) and
                  (TPaieD.GetValue('PPU_DATEDEBUT')<TDateFin.GetValue('DATE'))) do
                  begin
                  TDateInact := TOB.Create('Les Dates Inactivités', NIL, -1);
                  TDateInactD := TOB.Create('', TDateInact,-1);
                  TDateInactD.AddChampSup('DATE', FALSE);
                  TDateInactD.AddChampSup('NIVEAU', FALSE);
                  TDateInactD.AddChampSup('MOTIFENTREE', FALSE);
                  TDateInactD.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateInactD.AddChampSup('AJUDEB', FALSE);
                  TDateInactD.AddChampSup('AJUFIN', FALSE);
                  TDateInactD.PutValue('DATE', TPaieD.GetValue('PPU_DATEDEBUT'));
                  TDateInactD.PutValue('NIVEAU', 'D');
                  TDateInactD.PutValue('MOTIFENTREE', '095');
                  TDateInactD.PutValue('MOTIFSORTIE', '096');
                  TDateInactD.PutValue('AJUDEB', 0);
                  TDateInactD.PutValue('AJUFIN', -1);

                  TDateInactF := TOB.Create('', TDateInact,-1);
                  TDateInactF.AddChampSup('DATE', FALSE);
                  TDateInactF.AddChampSup('NIVEAU', FALSE);
                  TDateInactF.AddChampSup('MOTIFENTREE', FALSE);
                  TDateInactF.AddChampSup('MOTIFSORTIE', FALSE);
                  TDateInactF.AddChampSup('AJUDEB', FALSE);
                  TDateInactF.AddChampSup('AJUFIN', FALSE);
                  TDateInactF.PutValue('DATE', TPaieD.GetValue('PPU_DATEFIN'));
                  TDateInactF.PutValue('NIVEAU', 'D');
                  TDateInactF.PutValue('MOTIFENTREE', '095');
                  TDateInactF.PutValue('MOTIFSORTIE', '096');
                  TDateInactF.PutValue('AJUDEB', 1);
                  TDateInactF.PutValue('AJUFIN', 0);

                  TDADSUDetail:=TDADSUE.FindFirst(['PDE_SALARIE', 'PDE_DATEDEBUT'],
                                                  [Salarie, TDateInactD.GetValue('DATE')], TRUE);

                  if (TDADSUDetail = nil) then
                     begin
{PT109
                     PeriodeFaite:= RemplitPeriode (Salarie, TypeD,
                                                    TDateInactD, TDateInactF);
}
                     PeriodeFaite:= RemplitPeriode (Salarie, TDateInactD,
                                                    TDateInactF);
{PT100
                     if (TypeD <> '001') then
}
{PT101-7
                     if ((TypeD <> '001') and (PeriodeFaite <> 0)) then
}
                     if ((TypeD <> '001') and (TypeD <> '201') and
                        (PeriodeFaite <> 0)) then
//FIN PT101-7
{PT109
                        CalculBTP (Salarie, TypeD);
}
                        CalculBTP (Salarie);
                     NumPer:= NumPer+1;
                     NbreDate:= NbreDate+1;
                     Boucle:= 1;
                     FreeAndNil(TDADSUDetail);
                     end;
                  TPaieD:= TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
                  FreeAndNil (TDateInact);
                  end;
            end;
         if Boucle = 0 then
            NumPer:= NumPer+1
         else
            Boucle := 0;
         TDateDeb := TDateFin;
         TDateFin := TDate.FindNext([''],[''],FALSE);
         end;

   if (PeriodeFaite<>0) then
{PT109
      RemplitPrudh (Salarie, TypeD, PeriodeFaite);
}
      RemplitPrudh (Salarie, PeriodeFaite);

   FreeAndNil(TDate);
   end
else
   begin
{PT101-1
   Writeln(FRapport, '');
   Writeln(FRapport, 'Salarié : '+Salarie);
   Writeln(FRapport, 'Erreur    Le salarié a une date d''entrée postérieure à la fin de période');
   Writeln(FRapport, '          Le traitement est annulé pour ce salarié');
}
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 16/07/2001
Modifié le ... :   /  /
Description .. : Pour une période donnée, calcul des éléments qui
Suite ........ : rempliront les tables DADSPERIODES et DADSDETAIL
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
function RemplitPeriode(Salarie : string; TDateRupDeb, TDateRupFin : TOB):integer;
var
TBaseD, TContratD, TConventionD, TCotD, TEtabD, THistoCumSalD, THistoOrgD : TOB;
THistoSalD, TInstitD, TInstitution, TInstitutionFille, TOrganismeD : TOB;
TPaieD, TPrev, TPrevD, TRemD, TSalD, TTauxATD : TOB;
Buf, Buf2, BufContrat, BufDest, BufOrig, Cumul, HistoOrg : string;
MotifPerDeb, MotifPerFin, NIC, Prudh, SIREN, StCumul, StDernierMois : string;
StEtab : string;
BaseBruteFisc, HDixieme, Total, Tot1, Tot2, Tot3, Tot4, Tot5 : Double;
QRechCumul, QRechDernierMois, QRechEtab : TQuery;
DateResult : TDateTime;
AnneeA, JourJ, MoisM : Word;
i, Longueur, NbreHistoOrg, NbreInst, NumOrdre : integer;
Agirc, BoolFraisProf, Fait, Ircantec : boolean;
CleDADS : TCleDADS;
ErreurDADSU : TControle;
PrevoyanceLevelFind : Boolean;
TempTobPeriodesContrats : Tob; //PT143
stCode : String; //PT143

//Deb PT127
  Q : TQuery;
  isAffiliation : Boolean;
  EtabSalarie, PopSalarie : String;
  BaseSS : Double;
  NbEnfants : Integer;
  TobPeriodesContrats, TobPeriode, TobContrats, TobRechContrat, TempTob : Tob;
  IndexPeriode : Integer;
  Procedure AddPeriode(InTob, TobContrat : Tob);  //PT143
  var
    DateDebPeriode, DateFinPeriode : TDateTime;
    Deb, Fin : TDateTime;
    Institution, TypeContrat, NomPop, RefContrat, PopCouverte : String;
    NbEnfCouverts : Integer;
    Montant : Double;
  begin
    Deb           := TobContrat.GetDateTime('POP_DATEDEBUT');
    Fin           := TobContrat.GetDateTime('POP_DATEFIN');
    Institution   := TobContrat.GetString('POP_INSTITUTION');
    TypeContrat   := TobContrat.GetString('POP_TYPECONTRAT');
    NomPop        := TobContrat.GetString('POP_NOMPOP');
    RefContrat    := TobContrat.GetString('POP_REFERCONTRAT');
    PopCouverte   := TobContrat.GetString('POP_CODEPOPPREV');
    NbEnfCouverts := TobContrat.GetInteger('POP_NBENFANT');
    Montant       := 0;
    DateDebPeriode := StrToDate(StPerDeb);
    DateFinPeriode := StrToDate(StPerFin);
//    if (Deb > DateFinPeriode) then exit;
//    if (Fin < DateDebPeriode) then exit;
    With Tob.Create('Périodes contrat de prévoyance', InTob, -1) do
    begin
      AddChampSupValeur('DATEDEB_CONTRAT', Deb);
      AddChampSupValeur('DATEFIN_CONTRAT', Fin);
      AddChampSupValeur('DATEDEB', Max(Deb,DateDebPeriode));
      AddChampSupValeur('DATEFIN', Min(Fin,DateFinPeriode));
      AddChampSupValeur('INSTITUTION', Institution);
      AddChampSupValeur('TYPECONTRAT', TypeContrat);
      AddChampSupValeur('NOMPOP', NomPop);
      AddChampSupValeur('REFERCONTRAT', RefContrat);
      AddChampSupValeur('MONTANT', Montant, CSTDouble);
//PT128-3      AddChampSupValeur('PREDEFINI', Predefini);
//PT128-3      AddChampSupValeur('NODOSSIER', NoDossier);
//PT128-3      AddChampSupValeur('NATURERUB', NatureRub);
//PT128-3      AddChampSupValeur('RUBRIQUE', Rub);
      AddChampSupValeur('CODEPOPPREV', PopCouverte);
      AddChampSupValeur('NBENFANT', NbEnfCouverts, CSTInteger);
      AddChampSupValeur('CODEEVT', '');
    end;
  end;
(*PT143
  Procedure DecoupePeriode(var TobPeriodes : Tob; TobContrat : Tob);
  var
    DateDebContrat, DateFinContrat, DateDebPeriode, DateFinPeriode : TDateTime;
    IndexPeriode, IndexPeriodeDetail : Integer;
    tempTobPeriode, NewTobPeriodes : Tob;
  begin
    for IndexPeriode := 0 to TobPeriodes.detail.count -1 do
    begin
      DateDebContrat := TobContrat.GetDateTime('POP_DATEDEBUT');
      DateFinContrat := TobContrat.GetDateTime('POP_DATEFIN');
      tempTobPeriode := TobPeriodes.detail[IndexPeriode];
      DateDebPeriode := tempTobPeriode.GetDateTime('DATEDEB');
      DateFinPeriode := tempTobPeriode.GetDateTime('DATEFIN');
      if (DateFinContrat < DateDebPeriode) or (DateDebContrat > DateFinPeriode) then
        continue;
      if (DateDebContrat < DateDebPeriode) then DateDebContrat := DateDebPeriode;
      if (DateFinContrat > DateFinPeriode) then DateFinContrat := DateFinPeriode;
      if (DateDebContrat > DateDebPeriode) then
        AddPeriode(tempTobPeriode, DateDebPeriode, DateDebContrat-1, tempTobPeriode.GetString('INSTITUTION'),
                   tempTobPeriode.GetString('TYPECONTRAT'), tempTobPeriode.GetString('NOMPOP'),
                   tempTobPeriode.GetString('REFERCONTRAT')
//PT128-3                   , tempTobPeriode.GetString('PREDEFINI'),
//PT128-3                   tempTobPeriode.GetString('NODOSSIER'), tempTobPeriode.GetString('NATURERUB'),
//PT128-3                   tempTobPeriode.GetString('RUBRIQUE')
                   , tempTobPeriode.GetString('CODEPOPPREV'),
                   tempTobPeriode.GetInteger('NBENFANT'), tempTobPeriode.GetDouble('MONTANT')
        );
      AddPeriode(tempTobPeriode, DateDebContrat, DateFinContrat, TobContrat.GetString('POP_INSTITUTION'),
                 TobContrat.GetString('POP_TYPECONTRAT'), TobContrat.GetString('POP_NOMPOP'),
                 TobContrat.GetString('POP_REFERCONTRAT')
//PT128-3                 , TobContrat.GetString('POC_PREDEFINI'),
//PT128-3                 TobContrat.GetString('POC_NODOSSIER'), TobContrat.GetString('POC_NATURERUB'),
//PT128-3                 TobContrat.GetString('POC_RUBRIQUE')
                 , TobContrat.GetString('POP_CODEPOPPREV'),
                 TobContrat.GetInteger('POP_NBENFANT'),0
      );
      if (DateFinContrat < DateFinPeriode) then
        AddPeriode(tempTobPeriode, DateFinContrat+1, DateFinPeriode, tempTobPeriode.GetString('INSTITUTION'),
                   tempTobPeriode.GetString('TYPECONTRAT'), tempTobPeriode.GetString('NOMPOP'),
                   tempTobPeriode.GetString('REFERCONTRAT')
//PT128-3                   , tempTobPeriode.GetString('PREDEFINI'),
//PT128-3                   tempTobPeriode.GetString('NODOSSIER'), tempTobPeriode.GetString('NATURERUB'),
//PT128-3                   tempTobPeriode.GetString('RUBRIQUE')
                   , tempTobPeriode.GetString('CODEPOPPREV'),
                   tempTobPeriode.GetInteger('NBENFANT'), tempTobPeriode.GetDouble('MONTANT')
        );
    end;
    NewTobPeriodes := Tob.Create(TobPeriodes.NomTable, TobPeriodes.Parent, -1);
    for IndexPeriode := 0 to TobPeriodes.detail.count -1 do
    begin
      tempTobPeriode := TobPeriodes.detail[IndexPeriode];
      if tempTobPeriode.Detail.count > 0 then
      begin
        for IndexPeriodeDetail := 0 to tempTobPeriode.Detail.count -1 do
        begin
          with tob.create('Périodes contrat de prévoyance', NewTobPeriodes, -1) do
          begin
            Dupliquer(tempTobPeriode.Detail[IndexPeriodeDetail], true, true);
          end;
        end;
      end else begin
        with tob.create('Périodes contrat de prévoyance', NewTobPeriodes, -1) do
        begin
          Dupliquer(tempTobPeriode, true, true);
        end;
      end;
    end;
    FreeAndNil(TobPeriodes);
    TobPeriodes := NewTobPeriodes;
  end;
*)
(*PT143
  Function Rassemble(TobPeriodes : Tob ) : Boolean;
  var
    indexPeriode, IndexComparaison : Integer;
    T1, T2 : Tob;
    diff : Boolean;
  begin
    result := False;
    for indexPeriode := TobPeriodes.detail.count -2 downto 0 do
    begin
      T1 := TobPeriodes.detail[indexPeriode];
      T2 := TobPeriodes.detail[indexPeriode + 1 ];
      diff := False;
      //Assertion :
{$IFDEF TESTGGU}
      if not (T1.GetDateTime('DATEDEB') < T2.GetDateTime('DATEDEB')) then
        PGIError('PgDADSOutils - Erreur d''assertion : Les périodes ne sont pas ordonnées');
{$ENDIF}
      //Vérification de la continuité des champs date
      if   (T1.GetDateTime('DATEFIN') +1) <> (T2.GetDateTime('DATEDEB')) then
        diff := True;
      if diff = False then
      begin
        for IndexComparaison := 2 to T1.NombreChampSup-1 do //On ne compare pas les 2 premier champs qui sont les champs dates (on les a déjà vérifier)
        begin
          if (T1.GetValeur(1000+IndexComparaison) <> T2.GetValeur(1000+IndexComparaison)) then
          begin
            diff := True;
            break;
          end;
        end;
      end;
      if diff = False then  //Le contrat est le même, on regroupe les 2 périodes
      begin
        T1.PutValue('DATEFIN', T2.GetValue('DATEFIN'));
        T2.ChangeParent(nil, -1);
        FreeAndNil(T2);
      end;
    end;
  end;
*)
  Function SommeRubSalarie(NatureRub, CodeRub, Salarie : String; DebutPeriode, FinPeriode, DebExer, FinExer : TDateTime; MontantBaseCot : Boolean = False) : Double;
  var
    THisto, THistoD : Tob;
    StSQL : string;
    QRechSQL : TQuery;
    Function RecupMontant(THistoDetail : Tob) : Double;
    begin
      if MontantBaseCot then begin   //PT131
        if (THistoDetail.GetValue ('PHB_SENSBUL')<>'M') then
          result := THistoDetail.GetValue ('PHB_BASECOT')
        else
          result := -THistoDetail.GetValue ('PHB_BASECOT');
      end else begin
        if (THistoDetail.GetValue ('PHB_SENSBUL')<>'M') then
          result := THistoDetail.GetValue ('PHB_MTSALARIAL')
        else
          result := -THistoDetail.GetValue ('PHB_MTSALARIAL');
      end;
    end;
  begin
    result := 0;
    StSQL:= 'SELECT PHB_SALARIE, PHB_DATEDEBUT, PHB_DATEFIN,'+
            ' PHB_BASEREM, PHB_TAUXREM, PHB_COEFFREM, PHB_MTREM,'+
            ' PHB_BASECOT, PHB_TAUXSALARIAL, PHB_MTSALARIAL,'+
            ' PHB_SENSBUL'+
            ' FROM HISTOBULLETIN '+
            ' WHERE PHB_RUBRIQUE = "'+CodeRub+'"'+
            ' AND PHB_NATURERUB="'+NatureRub+'" '+
            ' AND PHB_SALARIE="'+Salarie+'" '+
            ' AND PHB_DATEFIN<="'+UsDateTime(FinExer)+'" '+
            ' AND PHB_DATEDEBUT>="'+UsDateTime(DebExer)+'" '+
            ' ORDER BY PHB_DATEDEBUT';
    QRechSQL:= OpenSql (StSQL, TRUE);
    THisto:= TOB.Create ('Les historiques', NIL, -1);
    THisto.LoadDetailDB ('HISTOPUBLIC','','',QRechSQL,False);
    Ferme (QRechSQL);
    THistoD:= THisto.FindFirst ([''], [''], TRUE);
    While (THistoD <> nil) do
    begin
      { On récupère le montant }
      if ((THistoD.GetValue ('PHB_DATEDEBUT')<DebutPeriode) or (THistoD.GetValue ('PHB_DATEDEBUT')>FinPeriode)) then
      begin
        if (((THistoD.GetValue ('PHB_DATEFIN')>DebutPeriode) and (THistoD.GetValue ('PHB_DATEFIN')<=FinPeriode))
         or ((THistoD.GetValue ('PHB_DATEDEBUT')<DebutPeriode) and (THistoD.GetValue ('PHB_DATEFIN')>=FinPeriode))) then
         result:= result+RecupMontant(THistoD);
      end else begin
        if (THistoD.GetValue ('PHB_DATEDEBUT')>=DebutPeriode) and (THistoD.GetValue ('PHB_DATEDEBUT')<=FinPeriode) then
          result:= result+RecupMontant(THistoD);
      end;
      THistoD:= THisto.FindNext ([''], [''], TRUE);
    end;
    FreeAndNil (THisto);
  end;
//Fin PT127
begin
result:=0;
TTauxATD:=nil;
//Recherche du salarié concerné dans la TOB Salariés
TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);

//Récupération des données des Tob Rupture
StPerDeb:= DateToStr(StrToDate(TDateRupDeb.GetValue('DATE'))+
           TDateRupDeb.GetValue('AJUDEB'));

MotifPerDeb := TDateRupDeb.GetValue('MOTIFENTREE');
StPerFin:= DateToStr(StrToDate(TDateRupFin.GetValue('DATE'))+
           TDateRupFin.GetValue('AJUFIN'));

MotifPerFin := TDateRupFin.GetValue('MOTIFSORTIE');

TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
While ((TPaieD <> nil) and
      ((TPaieD.GetValue('PPU_DATEDEBUT')>StrToDate(StPerDeb)) or
      (TPaieD.GetValue('PPU_DATEFIN')<StrToDate(StPerDeb)))) do
      TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
if ((TPaieD <> nil) and
   ((TPaieD.GetValue('PPU_DATEDEBUT')<=StrToDate(StPerDeb)) or
   (TPaieD.GetValue('PPU_DATEFIN')>=StrToDate(StPerDeb)))) then
   begin
   Buf := TPaieD.GetValue('PPU_DATEFIN');
   if ((TDateRupDeb.GetValue('NIVEAU') = 'A') and
      ((StrToDate(StPerDeb)<>TPaieD.GetValue('PPU_DATEDEBUT')))) then
      TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
   if (TPaieD <> nil) then
      begin
      StPerDeb:= TPaieD.GetValue('PPU_DATEDEBUT');
      if ((TPaieD.GetValue('PPU_DATEDEBUT')=TPaieD.GetValue('PPU_DATEFIN')) and
         (TDateRupDeb.GetValue('NIVEAU')<>'A')) then
         begin
         Buf2:= TPaieD.GetValue('PPU_DATEFIN');
         TPaieD:= TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
         if ((TPaieD <> nil) and
            (TPaieD.GetValue('PPU_DATEFIN')>StrToDate(StPerFin))) then
            StPerFin:= Buf2;
         end;
      end
   else
      StPerDeb := DateToStr(StrToDate(Buf)+1);
   end;

TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
While ((TPaieD <> nil) and
      ((TPaieD.GetValue('PPU_DATEDEBUT')>StrToDate(StPerFin)) or
      (TPaieD.GetValue('PPU_DATEFIN')<StrToDate(StPerFin)))) do
      TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
if ((TPaieD <> nil) and
   ((TPaieD.GetValue('PPU_DATEDEBUT')<=StrToDate(StPerFin)) or
   (TPaieD.GetValue('PPU_DATEFIN')>=StrToDate(StPerFin)))) then
   StPerFin := TPaieD.GetValue('PPU_DATEFIN');

if (StrToDate(StPerDeb)>StrToDate(StPerFin)) then
   exit;

//PT145
if ((Periode<>0) and (StrToDate(StPerDeb)>FinExer)) then
   begin
   result:= Periode;
   exit;
   end;
//FIN PT145

//PT125-5
if (MotifPerDeb='095') then
   begin
{PT142
   StPerDeb:= StPerFin;
}
   StDateDebut:= StPerFin;
//FIN PT142
   MotifPerFin:= '096';
   end
//PT142
else
   StDateDebut:= StPerDeb;
//FIN PT142   
//FIN PT125-5

Periode:= Periode+1;

{PT101-1
Writeln(FRapport, '********** PERIODE N° '+IntToStr(Periode));
Writeln(FRapport, '********** du '+StPerDeb+' au '+StPerFin);
}

//Remplissage de l'entête d'une période
{PT101-6
CreeEntete (Salarie, TypeD, Periode, StrToDate(StPerDeb), StrToDate(StPerFin),
           MotifPerDeb, MotifPerFin);
}
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= StrToDate (StPerDeb);
CleDADS.DateFin:= StrToDate (StPerFin);
CleDADS.Exercice:= PGExercice;

ErreurDADSU.Salarie:= Salarie;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= Periode;
ErreurDADSU.DateDeb:= StrToDate (StPerDeb);
ErreurDADSU.DateFin:= StrToDate (StPerFin);
ErreurDADSU.Exercice:= PGExercice;

CreeEntete (CleDADS, MotifPerDeb, MotifPerFin, Decalage);
//FIN PT101-6

result:= Periode;

//Remplissage du détail d'une période
{PT125-5
if ((TSalD.GetValue('PSA_DATESORTIE') < DebExer) AND
   (TSalD.GetValue('PSA_DATESORTIE') > IDate1900) AND
   (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
   CreeDetail (CleDADS, 1, 'S41.G01.00.001', '0101', '01/01/'+PGExercice)
else
   begin
}
{PT142
ForceNumerique(StPerDeb, Buf);
}
ForceNumerique (StDateDebut, Buf);
//FIN PT142
{PT125-5
CreeDetail (CleDADS, 1, 'S41.G01.00.001', Copy (Buf, 1, 4), StPerDeb);
}
CreeDetail (CleDADS, 1, 'S41.G01.00.001', Buf, StPerDeb);
//FIN PT125-5
{PT125-5
   end;
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.001');

CreeDetail (CleDADS, 2, 'S41.G01.00.002.001', MotifPerDeb, MotifPerDeb);
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.002.001');

{PT125-5
if ((TSalD.GetValue('PSA_DATESORTIE') < DebExer) AND
   (TSalD.GetValue('PSA_DATESORTIE') > IDate1900) AND
   (TSalD.GetValue('PSA_DATESORTIE') <> null)) then
   CreeDetail (CleDADS, 7, 'S41.G01.00.003', '0101', '01/01/'+PGExercice)
else
   begin
}
ForceNumerique(StPerFin, Buf);
{PT125-5
CreeDetail (CleDADS, 7, 'S41.G01.00.003', Copy(Buf, 1, 4), StPerFin);
}
CreeDetail (CleDADS, 7, 'S41.G01.00.003', Buf, StPerFin);
//FIN PT125-5
{PT125-5
   end;
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.003');

{PT101-9
if (MotifPerFin = '000') then
   EcrireErreur ('Erreur:   Saisie complémentaire Salarié :'+
                 ' Le motif de fin de période n''est pas renseigné',
                 Erreur)
else
if (MotifPerFin = '999') then
   EcrireErreur ('Erreur:   Saisie complémentaire Salarié :'+
                 ' Le motif de sortie n''est pas connu dans la DADS-U',
                 Erreur);
}
//FIN PT101-9
CreeDetail (CleDADS, 8, 'S41.G01.00.004.001', MotifPerFin, MotifPerFin);
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.004.001');

TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
While ((TPaieD <> nil) and
      ((TPaieD.GetValue('PPU_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TPaieD.GetValue('PPU_DATEDEBUT')>StrToDate(StPerFin)))) do
      TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
if ((TPaieD <> nil) and
   (TPaieD.GetValue('PPU_DATEDEBUT')>=StrToDate(StPerDeb)) and
   (TPaieD.GetValue('PPU_DATEDEBUT')<=StrToDate(StPerFin))) then
   begin
   EtabSalarie := TPaieD.GetValue('PPU_ETABLISSEMENT');//PT127
   BufOrig := TPaieD.GetValue('ET_SIRET');
   ForceNumerique(BufOrig, BufDest);
   if ControlSiret(BufDest)=False then
{PT101-1
      EcrireErreur ('Erreur:   Fiche Etablissement : Le SIRET n''est pas valide',
                   Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.005';
      ErreurDADSU.Explication:= 'Le SIRET de l''établissement n''est pas valide';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
//PT101-2
   ForceNumerique (GetParamSoc('SO_SIRET'), SIREN);
   if (Copy (BufDest, 1, 9)<> Copy (SIREN, 1, 9)) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.005';
      ErreurDADSU.Explication:= 'Le SIREN de l''établissement ne correspond'+
                                ' pas au SIREN de l''entreprise';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-2
   NIC := Copy(BufDest, 10, 5);
   CreeDetail (CleDADS, 13, 'S41.G01.00.005', NIC,
               TPaieD.GetValue('PPU_ETABLISSEMENT'));
   end
else
   begin
   TEtabD := TEtab.FindFirst(['ET_ETABLISSEMENT'],
                             [TSalD.GetValue('PSA_ETABLISSEMENT')], TRUE);
   if ((TEtabD <> nil) and (TEtabD.GetValue('ET_SIRET') <> '') and
      (TEtabD.GetValue('ET_SIRET') <> null)) then
      begin
      EtabSalarie := TSalD.GetValue('PSA_ETABLISSEMENT');//PT127
      BufOrig := TEtabD.GetValue('ET_SIRET');
      ForceNumerique(BufOrig, BufDest);
      if ControlSiret(BufDest)=False then
{PT101-1
         EcrireErreur ('Erreur:   Fiche Etablissement : Le SIRET n''est pas valide',
                      Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.005';
         ErreurDADSU.Explication:= 'Le SIRET de l''établissement n''est pas'+
                                   ' valide';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
//PT101-2
      ForceNumerique (GetParamSoc('SO_SIRET'), SIREN);
      if (Copy (BufDest, 1, 9)<> Copy (SIREN, 1, 9)) then
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.005';
         ErreurDADSU.Explication:= 'Le SIREN de l''établissement ne correspond'+
                                   ' pas au SIREN de l''entreprise';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-2
      NIC := Copy(BufDest, 10, 5);
      BufOrig := TEtabD.GetValue('ET_ETABLISSEMENT');
      CreeDetail (CleDADS, 13, 'S41.G01.00.005', NIC, BufOrig);
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.005');

if (TSalD.GetValue('PSA_MULTIEMPLOY') = '-') then
   CreeDetail (CleDADS, 14, 'S41.G01.00.008.001', '01', '01')
else
   CreeDetail (CleDADS, 14, 'S41.G01.00.008.001', '02', '02');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.008.001');

{A traiter avec la gestion des carrières et des postes de travail
Pour l'instant, initialisé à inconnu}
{PT101-2
CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
           StrToDate(StPerFin), 15, 'S41.G01.00.008.002', '01', '01');
}
CreeDetail (CleDADS, 15, 'S41.G01.00.008.002', '03', '03');
//FIN PT101-2           
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.008.002');

{PT101-2
if (VH_Paie.PGDecalage = TRUE) then
   CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
              StrToDate(StPerFin), 16, 'S41.G01.00.009', '03', '03')
else
   CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
              StrToDate(StPerFin), 16, 'S41.G01.00.009', '01', '01');
}
//PT125-6
if (StrToDate(StPerDeb)>FinExer) then
{PT144
   Decalage:= '05';
}
   Buf:= '05'
else
   Buf:= Decalage;
//FIN PT144
//FIN PT125-6

{PT144
if (Decalage='02') then
}
if (Buf='02') then
//FIN PT144
   begin
   if (DateDecalage<StrToDate(StPerDeb)) then
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', '03', '03')
   else
   if ((DateDecalage>=StrToDate(StPerDeb)) and
      (DateDecalage<=StrToDate(StPerFin))) then
{PT144
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', Decalage, Decalage)
}
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', Buf, Buf)
//FIN PT144
   else
   if (DateDecalage>StrToDate(StPerFin)) then
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', '01', '01');
   end
else
{PT144
if (Decalage='04') then
}
if (Buf='04') then
//FIN PT144
   begin
   if (DateDecalage<StrToDate(StPerDeb)) then
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', '01', '01')
   else
   if ((DateDecalage>=StrToDate(StPerDeb)) and
      (DateDecalage<=StrToDate(StPerFin))) then
{PT144
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', Decalage, Decalage)
}
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', Buf, Buf)
//FIN PT144
   else
   if (DateDecalage>StrToDate(StPerFin)) then
      CreeDetail (CleDADS, 16, 'S41.G01.00.009', '03', '03');
   end
else
{PT144
   CreeDetail (CleDADS, 16, 'S41.G01.00.009', Decalage, Decalage);
}
   CreeDetail (CleDADS, 16, 'S41.G01.00.009', Buf, Buf);
//FIN PT144
//FIN PT101-2
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.009');

if ((TSalD.GetValue('PSA_LIBELLEEMPLOI') <> '') and
   (TSalD.GetValue('PSA_LIBELLEEMPLOI') <> null) and
   (RechDom('PGLIBEMPLOI', TSalD.GetValue('PSA_LIBELLEEMPLOI'), FALSE) <> '')) then
   begin
   Buf := RechDom('PGLIBEMPLOI', TSalD.GetValue('PSA_LIBELLEEMPLOI'), FALSE);
   Longueur := Length(Buf);
   CreeDetail (CleDADS, 17, 'S41.G01.00.010', Buf,
               TSalD.GetValue('PSA_LIBELLEEMPLOI'));
   end
else
{PT101-1 - PT101-3
   CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
              StrToDate(StPerFin), 17, 'S41.G01.00.010', '', '');
   EcrireErreur('Erreur:   Fiche Salarié : Libellé emploi', Erreur);
}
   begin
   ErreurDADSU.Segment:= 'S41.G01.00.010';
   ErreurDADSU.Explication:= 'Le libellé emploi n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT101-1
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.010');

if ((TSalD.GetValue('PSA_CODEEMPLOI') <> '')  and
   (TSalD.GetValue('PSA_CODEEMPLOI') <> null)) then
   CreeDetail (CleDADS, 18, 'S41.G01.00.011', TSalD.GetValue('PSA_CODEEMPLOI'),
               TSalD.GetValue('PSA_CODEEMPLOI'));
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.011');

BufContrat := '';
TContratD := TContrat.FindFirst(['PCI_SALARIE'], [Salarie], TRUE);
While ((TContratD <> nil) and
      (((TContratD.GetValue('PCI_DEBUTCONTRAT')<StrToDate(StPerDeb)) and
      (TContratD.GetValue('PCI_FINCONTRAT')<StrToDate(StPerDeb)) and
      (TContratD.GetValue('PCI_FINCONTRAT')>Idate1900)) or
      (TContratD.GetValue('PCI_DEBUTCONTRAT')>StrToDate(StPerFin)))) do
      TContratD := TContrat.FindNext(['PCI_SALARIE'], [Salarie], TRUE);
if ((TContratD <> nil) and (TContratD.GetValue('PCI_TYPECONTRAT') <> '') and
   (TContratD.GetValue('PCI_TYPECONTRAT') <> null)) then
   begin
   BufContrat := TContratD.GetValue('PCI_TYPECONTRAT');
   if (BufContrat = 'CAC') then
      BufContrat := 'CAP';
   CreeDetail (CleDADS, 19, 'S41.G01.00.012.001',
               RechDom('PGTYPECONTRAT', BufContrat, TRUE), BufContrat);
   end
else
   begin
   CreeDetail (CleDADS, 19, 'S41.G01.00.012.001', '01', 'CDI');
{PT101-1
   EcrireErreur('          Le contrat a été forcé à ''CDI''', Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.00.012.001';
   ErreurDADSU.Explication:= 'Le contrat a été forcé à ''CDI''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.012.001');

CreeDetail (CleDADS, 20, 'S41.G01.00.012.002', '01', '01');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.012.002');

Buf2 := '';
if ((TSalD.GetValue('PSA_CONDEMPLOI') <> '') and
   (TSalD.GetValue('PSA_CONDEMPLOI') <> null)) then
   Buf2 := TSalD.GetValue('PSA_CONDEMPLOI');

if (GetParamSocSecur('SO_PGHDADSHISTO', '-') = True) then
   begin
   THistoSalD := THistoSal.FindFirst(['PHS_SALARIE'], [Salarie], TRUE);
   While ((THistoSalD <> nil) and
         (THistoSalD.GetValue('PHS_SALARIE') = Salarie)) do
         begin
         if (THistoSalD.GetValue('PHS_DATEEVENEMENT')<= StrToDate(StPerDeb)) then
            begin
            if (THistoSalD.GetValue('PHS_CONDEMPLOI') <> null) then
               Buf2 := THistoSalD.GetValue('PHS_CONDEMPLOI')
            else
               Buf2 := 'W';
            end;
         THistoSalD := THistoSal.FindNext(['PHS_SALARIE'], [Salarie], TRUE);
         end;
   end;

if Buf2 = '' then
   begin
   Buf := '90';
   Buf2 := 'W';
{PT101-1
   EcrireErreur ('          Fiche Salarié :'+
                 ' Condition d''emploi a été forcée à ''Sans complément de contrat''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.00.013';
{PT123
   ErreurDADSU.Explication:= 'La condition d''emploi a été forcée à ''Sans'+
                             ' complément de contrat''';
}
   ErreurDADSU.Explication:= 'La caractéristique activité a été forcée à'+
                             ' ''Sans complément de contrat''';
//FIN PT123
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;

if Buf2 = 'C' then
   Buf := '01'
else
if Buf2 = 'P' then
   Buf := '02'
else
if Buf2 = 'I' then
   Buf := '04'
else
if Buf2 = 'D' then
   Buf := '05'
else
if Buf2 = 'S' then
   Buf := '06'
else
if Buf2 = 'V' then
   Buf := '07'
else
if Buf2 = 'O' then
   Buf := '08'
else
if Buf2 = 'N' then
   Buf := '09'
else
if Buf2 = 'F' then
   Buf := '10'
else
if Buf2 = 'W' then
   Buf := '90';
CreeDetail (CleDADS, 21, 'S41.G01.00.013', Buf, Buf2);
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.013');

BufProf := '';
if ((TSalD.GetValue('PSA_DADSPROF') <> '') and
   (TSalD.GetValue('PSA_DADSPROF') <> null)) then
   BufProf := TSalD.GetValue('PSA_DADSPROF');

if (GetParamSocSecur('SO_PGHDADSHISTO', '-') = True) then
   begin
   THistoSalD := THistoSal.FindFirst(['PHS_SALARIE'], [Salarie], TRUE);
   While ((THistoSalD <> nil) and
         (THistoSalD.GetValue('PHS_SALARIE') = Salarie)) do
         begin
         if (THistoSalD.GetValue('PHS_DATEEVENEMENT')<= StrToDate(StPerDeb)) then
            begin
            if (THistoSalD.GetValue('PHS_DADSPROF') <> null) then
               BufProf := THistoSalD.GetValue('PHS_DADSPROF')
            else
               BufProf := '';
            end;
         THistoSalD := THistoSal.FindNext(['PHS_SALARIE'], [Salarie], TRUE);
         end;
   end;
if (BufProf <> '') then
   begin
   if (BufProf = 'ZZZ') then
      BufProf := '01';
   CreeDetail (CleDADS, 22, 'S41.G01.00.014', BufProf, BufProf)
   end
else
   begin
   CreeDetail (CleDADS, 22, 'S41.G01.00.014', '90', '90');
{PT101-1
   EcrireErreur ('          Fiche Salarié :'+
                 ' Statut professionnel a été forcé à ''Pas de statut''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.00.014';
   ErreurDADSU.Explication:= 'Le statut professionnel a été forcé à ''Pas de'+
                             ' statut''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;

Debug('Paie PGI/Calcul DADS-U : S41.G01.00.014');

Buf := '';
if ((TSalD.GetValue('PSA_DADSCAT') <> '') and
   (TSalD.GetValue('PSA_DADSCAT') <> null)) then
   Buf := TSalD.GetValue('PSA_DADSCAT');

if (GetParamSocSecur('SO_PGHDADSHISTO', '-') = True) then
   begin
   THistoSalD := THistoSal.FindFirst(['PHS_SALARIE'], [Salarie], TRUE);
   While ((THistoSalD <> nil) and
         (THistoSalD.GetValue('PHS_SALARIE') = Salarie)) do
         begin
         if (THistoSalD.GetValue('PHS_DATEEVENEMENT')<= StrToDate(StPerDeb)) then
            begin
            if (THistoSalD.GetValue('PHS_DADSCAT') <> null) then
               Buf := THistoSalD.GetValue('PHS_DADSCAT')
            else
               Buf := '';
            end;
         THistoSalD := THistoSal.FindNext(['PHS_SALARIE'], [Salarie], TRUE);
         end;
   end;

//PT125-2
if (((BufProf='01') or (BufProf='02') or (BufProf='05') or (BufProf='14')) and
   (Buf<>'04')) then
   begin
   Buf:= '04';
   ErreurDADSU.Segment:= 'S41.G01.00.015.002';
   ErreurDADSU.Explication:= 'En raison de la valeur du statut professionnel,'+
                             ' le statut catégoriel a été forcé à  ''Non cadre''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
   end;

if (((BufProf='12') or (BufProf='13') or (BufProf='22') or (BufProf='23') or
   (BufProf='24') or (BufProf='25')) and (Buf<>'01')) then
   begin
   Buf:= '01';
   ErreurDADSU.Segment:= 'S41.G01.00.015.002';
   ErreurDADSU.Explication:= 'En raison de la valeur du statut professionnel,'+
                             ' le statut catégoriel a été forcé à  ''Cadre''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT125-2

{PT101-2
if (Buf <> '') then
   CreeDetail (CleDADS, 23, 'S41.G01.00.015', Buf, Buf)
else
   begin
   CreeDetail (CleDADS, 23, 'S41.G01.00.015', '90', '90');
   EcrireErreur ('          Fiche Salarié :'+
                 ' Statut catégoriel a été forcé à ''Pas de statut''', Erreur);
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.015');
}

TConventionD := TConvention.FindFirst(['PCV_CONVENTION'],
                                      [TSalD.GetValue('PSA_CONVENTION')], TRUE);
if (TConventionD<>nil) then
   begin
{PT101-2
   if ((TConventionD.GetValue('PCV_IDCC') <> '') and
      (TConventionD.GetValue('PCV_IDCC') <> null)) then
      CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
                 StrToDate(StPerFin), 24, 'S41.G01.00.016',
                 TConventionD.GetValue('PCV_IDCC'),
                 TConventionD.GetValue('PCV_IDCC'))
}
   if ((TConventionD.GetValue ('PCV_IDCC')<>'') and
      (TConventionD.GetValue ('PCV_IDCC')<>null)) then
      begin
      if ((Buf<>'') and (Buf<>'90')) then
         begin
         if (Buf='01') then
            begin
            CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', Buf, Buf);
            CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', Buf, Buf);
            end
         else
            begin
            CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '02', '02');
            CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', Buf, Buf);
            end;
         end
      else
         begin
         CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '90', '90');
         CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', '04', '04');
         ErreurDADSU.Segment:= 'S41.G01.00.015.001';
         ErreurDADSU.Explication:= 'Le statut catégoriel a été forcé à'+
                                   ' ''Pas de statut''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;

      CreeDetail (CleDADS, 25, 'S41.G01.00.016',
                  TConventionD.GetValue('PCV_IDCC'),
                  TConventionD.GetValue('PCV_IDCC'));
      end
//FIN PT101-2
   else
      begin
{PT101-2
      CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
                 StrToDate(StPerFin), 24, 'S41.G01.00.016', '9999', '9999');
}
      CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '90', '90');
      CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', '04', '04');
      ErreurDADSU.Segment:= 'S41.G01.00.015.001';
      ErreurDADSU.Explication:= 'Le statut catégoriel a été forcé à'+
                                ' ''Pas de statut''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);

      CreeDetail (CleDADS, 25, 'S41.G01.00.016', '9999', '9999');
//FIN PT101-2
{PT101-1
      EcrireErreur ('          Saisie complémentaire Salarié :'+
                    ' Code Convention collective a été forcé à ''9999''',
                    Erreur);
}
      ErreurDADSU.Segment:= 'S41.G01.00.016';
      ErreurDADSU.Explication:= 'Le code Convention collective a été forcé à'+
                                ' ''9999''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end
else
   begin
{PT101-2
   CreeDetail (Salarie, TypeD, Periode, StrToDate(StPerDeb),
              StrToDate(StPerFin), 24, 'S41.G01.00.016', '9999', '9999');
}
   CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '90', '90');
   CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', '04', '04');
   ErreurDADSU.Segment:= 'S41.G01.00.015.001';
   ErreurDADSU.Explication:= 'Le statut catégoriel a été forcé à'+
                             ' ''Pas de statut''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);

   CreeDetail (CleDADS, 25, 'S41.G01.00.016', '9999', '9999');
//FIN PT101-2
{PT101-1
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Code Convention collective a été forcé à ''9999''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.00.016';
   ErreurDADSU.Explication:= 'Le code Convention collective a été forcé à'+
                             ' ''9999''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.015');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.016');


if ((TSalD.GetValue('PSA_COEFFICIENT') <> '') and
   (TSalD.GetValue('PSA_COEFFICIENT') <> null) and
   (RechDom('PGCOEFFICIENT', TSalD.GetValue('PSA_COEFFICIENT'), FALSE) <> '')) then
   CreeDetail (CleDADS, 26, 'S41.G01.00.017',
              RechDom('PGCOEFFICIENT', TSalD.GetValue('PSA_COEFFICIENT'), FALSE),
              TSalD.GetValue('PSA_COEFFICIENT'))
else
   begin
   CreeDetail (CleDADS, 26, 'S41.G01.00.017', 'sans classement conventionnel',
               '');
{PT101-1
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Coefficient a été forcé à ''sans classement conventionnel''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.00.017';
   ErreurDADSU.Explication:= 'Le coefficient a été forcé à ''sans classement'+
                             ' conventionnel''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.017');

//DEB PT124
if (TSalD.GetValue('PSA_TYPEREGIME') = '-') then
   begin
   if ((TSalD.GetValue('PSA_REGIMESS') = '') or
      (TSalD.GetValue('PSA_REGIMESS') = null)) then
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié : Régime SS', Erreur)
}
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.018.001';
      ErreurDADSU.Explication:= 'Le régime SS n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
//FIN PT101-1
   else
{PT125-2
      CreeDetail (CleDADS, 27, 'S41.G01.00.018.001', TSalD.GetValue('PSA_REGIMESS'),
                  TSalD.GetValue('PSA_REGIMESS'));
}
      begin
      if ((TSalD.GetValue ('PSA_REGIMESS')='200') and (MotifPerDeb='085') and
         (MotifPerFin='086')) then
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.018.001';
         ErreurDADSU.Explication:= 'Le régime SS n''est pas valide pour cette'+
                                   ' période';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
      CreeDetail (CleDADS, 27, 'S41.G01.00.018.001', TSalD.GetValue ('PSA_REGIMESS'),
                  TSalD.GetValue ('PSA_REGIMESS'));
      end;
//FIN PT125-2
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.018.001');
   end
else
   begin
   if ((TSalD.GetValue('PSA_REGIMEMAL') = '') or
      (TSalD.GetValue('PSA_REGIMEMAL') = null)) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.018.002';
      ErreurDADSU.Explication:= 'Le régime obligatoire risque maladie n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
   else
      CreeDetail (CleDADS, 28, 'S41.G01.00.018.002', TSalD.GetValue('PSA_REGIMEMAL'),
                  TSalD.GetValue('PSA_REGIMEMAL'));
   Debug('Paie PGI/Calcul DADS-U : S41.G01.00.018.002');

   if ((TSalD.GetValue('PSA_REGIMEAT') = '') or
      (TSalD.GetValue('PSA_REGIMEAT') = null)) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.018.003';
      ErreurDADSU.Explication:= 'Le régime obligatoire risque AT n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
   else
{PT125-2
      CreeDetail (CleDADS, 29, 'S41.G01.00.018.003', TSalD.GetValue('PSA_REGIMEAT'),
                  TSalD.GetValue('PSA_REGIMEAT'));
}
      begin
      if ((TSalD.GetValue ('PSA_REGIMESS')<>'901') and (MotifPerDeb='085') and
         (MotifPerFin='086')) then
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.018.003';
         ErreurDADSU.Explication:= 'Le régime obligatoire risque AT n''est pas'+
                                   ' valide pour cette période';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
      CreeDetail (CleDADS, 29, 'S41.G01.00.018.003', TSalD.GetValue ('PSA_REGIMEAT'),
                  TSalD.GetValue ('PSA_REGIMEAT'));
      end;
   Debug('Paie PGI/Calcul DADS-U : S41.G01.00.018.003');

   if ((TSalD.GetValue('PSA_REGIMEVIP') = '') or
      (TSalD.GetValue('PSA_REGIMEVIP') = null)) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.018.004';
      ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PP) n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
   else
      CreeDetail (CleDADS, 30, 'S41.G01.00.018.004', TSalD.GetValue('PSA_REGIMEVIP'),
                  TSalD.GetValue('PSA_REGIMEVIP'));
   Debug('Paie PGI/Calcul DADS-U : S41.G01.00.018.004');

   if ((TSalD.GetValue('PSA_REGIMEVIS') = '') or
      (TSalD.GetValue('PSA_REGIMEVIS') = null)) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.018.005';
      ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PS) n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end
   else
      CreeDetail (CleDADS, 31, 'S41.G01.00.018.005', TSalD.GetValue('PSA_REGIMEVIS'),
                  TSalD.GetValue('PSA_REGIMEVIS'));
   Debug('Paie PGI/Calcul DADS-U : S41.G01.00.018.005');
   end;
//FIN PT124

//PT133
TEtabD := TEtab.FindFirst(['ET_ETABLISSEMENT'],
                          [TSalD.GetValue('PSA_ETABLISSEMENT')], TRUE);
if (TEtabD.GetValue ('ETB_REGIMEALSACE')='X') then
   CreeDetail (CleDADS, 32, 'S41.G01.00.018.006', '01', '01');
//FIN PT133

CreeDetail (CleDADS, 33, 'S41.G01.00.019', Salarie, Salarie);
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.019');

if (TSalD.GetValue('PSA_TAUXPARTIEL') <> null) then
   begin
   BufOrig := FloatToStr(TSalD.GetValue('PSA_TAUXPARTIEL'));
   if (GetParamSocSecur('SO_PGHDADSHISTO', '-') = True) then
      begin
      THistoSalD := THistoSal.FindFirst(['PHS_SALARIE'], [Salarie], TRUE);
      While ((THistoSalD <> nil) and
            (THistoSalD.GetValue('PHS_SALARIE') = Salarie)) do
            begin
            if (THistoSalD.GetValue('PHS_DATEEVENEMENT')<= StrToDate(StPerDeb)) then
               BufOrig := THistoSalD.GetValue('PHS_TTAUXPARTIEL');
            THistoSalD := THistoSal.FindNext(['PHS_SALARIE'], [Salarie], TRUE);
            end;
      end;
   if ((BufOrig <> '0') and (BufOrig <> '') and (Buf2 = 'P')) then
      begin
      BufOrig := FloatToStr(Arrondi(StrToFloat(BufOrig), 2));
      BufDest := FloatToStr(StrToFloat(BufOrig)*100);
      CreeDetail (CleDADS, 34, 'S41.G01.00.020', BufDest, BufOrig);
      end
   else
      if (Buf2 = 'P') then
{PT101-1
         EcrireErreur ('Erreur:   Fiche Salarié :'+
                       ' Le taux de travail à temps partiel n''est pas renseigné',
                       Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.020';
         ErreurDADSU.Explication:= 'Le taux de travail à temps partiel n''est'+
                                   ' pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.020');

Tot1:= 0;
{PT101-2
if (MotifPerDeb <> '095') then
}
if ((Buf2<>'D') and (BufProf <> '14') and (BufProf <> '15') and
   (BufProf <> '16') and (MotifPerDeb <> '095')) then
//FIN PT101-2
   begin
   THistoCumSalD:= THistoCumSal.FindFirst (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                           [Salarie, '20'], TRUE);
   While ((THistoCumSalD <> nil) and
         ((THistoCumSalD.GetValue ('PHC_DATEDEBUT')<StrToDate (StPerDeb)) or
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')>StrToDate (StPerFin)))) do
         begin
         if ((THistoCumSalD.GetValue ('PHC_DATEFIN')>StrToDate (StPerDeb)) and
            (THistoCumSalD.GetValue ('PHC_DATEFIN')<=StrToDate (StPerFin))) then
            Tot1:= Tot1+THistoCumSalD.GetValue ('PHC_MONTANT');
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '20'], TRUE);
         end;
   While ((THistoCumSalD <> nil) and
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')>=StrToDate (StPerDeb)) and
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')<=StrToDate (StPerFin))) do
         begin
         Tot1:= Tot1+THistoCumSalD.GetValue ('PHC_MONTANT');
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '20'], TRUE);
         end;
   if (Tot1<0) then
      begin
      Tot1:= 0;
{PT101-1
      EcrireErreur ('          Saisie complémentaire Salarié :'+
                    ' Le nombre d''heures travaillées était négatif, il a été'+
                    ' forcé à 0', Erreur);
}
      ErreurDADSU.Segment:= 'S41.G01.00.021';
      ErreurDADSU.Explication:= 'Le nombre d''heures travaillées était négatif,'+
                                ' il a été forcé à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end
   else
      if (Tot1=0) then
{PT101-1
         EcrireErreur ('          Saisie complémentaire Salarié : Le nombre'+
                       ' d''heures travaillées est égal à 0.', Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.021';
         ErreurDADSU.Explication:= 'Le nombre d''heures travaillées est égal à 0';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1

   if (Arrondi(Tot1, 0)<>0) then
      CreeDetail (CleDADS, 35, 'S41.G01.00.021', FloatToStr (Arrondi(Tot1, 0)),
                  FloatToStr (Arrondi (Tot1, 0)));
   end;
{PT101-2
else
   EcrireErreur ('          Le nombre d''heures travaillées est égal à 0 car'+
                 ' la période générée correspond à une somme isolée.', Erreur);
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.021');

Tot2:= 0;
if ((Buf2<>'D') and (BufProf <> '14') and (BufProf <> '15') and
   (BufProf <> '16') and (MotifPerDeb <> '095')) then
   begin
   THistoCumSalD:= THistoCumSal.FindFirst (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                           [Salarie, '21'], TRUE);
   While ((THistoCumSalD <> nil) and
         ((THistoCumSalD.GetValue ('PHC_DATEDEBUT')<StrToDate (StPerDeb)) or
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')>StrToDate (StPerFin)))) do
         begin
         if ((THistoCumSalD.GetValue ('PHC_DATEFIN')>StrToDate (StPerDeb)) and
            (THistoCumSalD.GetValue ('PHC_DATEFIN')<=StrToDate (StPerFin))) then
            Tot2:= Tot2+THistoCumSalD.GetValue ('PHC_MONTANT');
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '21'], TRUE);
         end;
   While ((THistoCumSalD <> nil) and
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')>=StrToDate (StPerDeb)) and
         (THistoCumSalD.GetValue ('PHC_DATEDEBUT')<=StrToDate (StPerFin))) do
         begin
         Tot2:= Tot2+THistoCumSalD.GetValue ('PHC_MONTANT');
         THistoCumSalD:= THistoCumSal.FindNext (['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '21'], TRUE);
         end;
   if (Tot2<0) then
      begin
      Tot2:= 0;
{PT101-1
      EcrireErreur ('          Saisie complémentaire Salarié :'+
                    ' Le nombre d''heures salariées était négatif, il a été'+
                    ' forcé à 0', Erreur);
}
      ErreurDADSU.Segment:= 'S41.G01.00.022';
      ErreurDADSU.Explication:= 'Le nombre d''heures payées était négatif,'+
                                ' il a été forcé à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end
   else
      if (Tot2=0) then
{PT101-1
         EcrireErreur ('          Saisie complémentaire Salarié : Le nombre'+
                       ' d''heures salariées est égal à 0.', Erreur);
}
         begin
         ErreurDADSU.Segment:= 'S41.G01.00.022';
         ErreurDADSU.Explication:= 'Le nombre d''heures payées est égal à 0';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT101-1
//PT103-1
   if (Tot2<Tot1) then
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.022';
      ErreurDADSU.Explication:= 'Le nombre d''heures travaillées est'+
                                ' supérieur au total d''heures payées';
{PT110
      ErreurDADSU.CtrlBloquant:= True;
}
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT103-1
   if (Arrondi (Tot2, 0)<>0) then
      CreeDetail (CleDADS, 36, 'S41.G01.00.022', FloatToStr (Arrondi (Tot2, 0)),
                  FloatToStr (Arrondi (Tot2, 0)));
   end;
{PT101-2
else
   if (MotifPerDeb='095') then
      EcrireErreur ('          Le nombre d''heures salariées est égal à 0 car'+
                    ' la période générée correspond à une somme isolée.',
                    Erreur);
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.022');

if ((Total < 1200) or (MotifPerFin = '008') or (MotifPerFin = '010') or
   (MotifPerFin = '012') or (MotifPerFin = '014') or (MotifPerFin = '016') or
   (MotifPerFin = '020')) then
   begin
   StDernierMois:= 'SELECT PHC_DATEDEBUT, PHC_MONTANT'+
                   ' FROM HISTOCUMSAL WHERE'+
                   ' PHC_CUMULPAIE="21" AND'+
                   ' PHC_SALARIE = "'+Salarie+'" AND'+
                   ' PHC_MONTANT >= 60 AND'+
                   ' PHC_DATEDEBUT <= "'+UsDateTime(StrToDate(StPerFin))+'" AND'+
                   ' PHC_DATEDEBUT >= "'+UsDateTime(StrToDate(StPerDeb))+'"'+
                   ' ORDER BY PHC_DATEDEBUT DESC';
   QRechDernierMois:=OpenSql(StDernierMois,TRUE);
   if (not QRechDernierMois.EOF) then
      DateResult := QRechDernierMois.FindField ('PHC_DATEDEBUT').AsDateTime;
   Ferme(QRechDernierMois);
   if (DateResult <> 0) then
      begin
      DecodeDate(DateResult, AnneeA, MoisM, JourJ);
      CreeDetail (CleDADS, 37, 'S41.G01.00.023', ColleZeroDevant(MoisM, 2),
                  ColleZeroDevant(MoisM, 2));
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.023');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_HCHOMPAR'], [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_HCHOMPAR'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')='M') then
         Total := Total + TRemD.GetValue('PHB_BASEREM')
      else
         Total := Total - TRemD.GetValue('PHB_BASEREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_HCHOMPAR'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total > 0) then
   CreeDetail (CleDADS, 38, 'S41.G01.00.024', FloatToStr(Arrondi(Total, 0)),
               FloatToStr(Arrondi(Total, 0)));
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.024');

Fait := FALSE;
{PT101-2
if ((TSalD.GetValue('PSA_REGIMESS') <> '122') and
   (TSalD.GetValue('PSA_REGIMESS') <> '200')) then
   begin
   CreeDetail (CleDADS, 33, 'S41.G01.00.025', '98', '98');
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Régime de base obligatoire différent de régime général :'+
                 ' la section AT a été forcé à ''98''',
                Erreur);
   Fait := TRUE;
   end
else
}
if ((((TSalD.GetValue('PSA_REGIMESS') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = '-'))      //PT124
  or ((TSalD.GetValue('PSA_REGIMEAT') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = 'X')))     //PT124
{PT138
  and (BufProf<>'07')) then
}
  and (BufProf<>'07') and (BufProf<>'36') and (BufProf<>'37') and
      (BufProf<>'38') and (BufProf<>'39')) then
//FIN PT138
//FIN PT101-2
   begin
   TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
   While ((TPaieD <> nil) and
         ((TPaieD.GetValue('PPU_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TPaieD.GetValue('PPU_DATEDEBUT')>StrToDate(StPerFin)))) do
         TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
   if ((TPaieD <> nil) and
      (TPaieD.GetValue('PPU_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TPaieD.GetValue('PPU_DATEDEBUT')<=StrToDate(StPerFin))) then
      begin
      if (TPaieD <> nil) then
         begin
         TTauxATD := TTauxAT.FindFirst(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'],
                                       [TPaieD.GetValue('PPU_ETABLISSEMENT'),
                                       TSalD.GetValue('PSA_ORDREAT')], TRUE);
         While ((TTauxATD <> nil) and
               (TTauxATD.GetValue('PAT_DATEVALIDITE')>StrToDate(StPerFin))) do
               TTauxATD := TTauxAT.FindNext(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'],
                                            [TPaieD.GetValue('PPU_ETABLISSEMENT'),
                                            TSalD.GetValue('PSA_ORDREAT')], TRUE);

         if (TTauxATD <> nil) then
            begin
            if ((TTauxATD.GetValue('PAT_SECTIONAT') <> '') and
               (TTauxATD.GetValue('PAT_SECTIONAT') <> '98') and	//PT101-2
               (TTauxATD.GetValue('PAT_SECTIONAT') <> null)) then
               begin
               CreeDetail (CleDADS, 39, 'S41.G01.00.025',
                           TTauxATD.GetValue('PAT_SECTIONAT'),
                           TTauxATD.GetValue('PAT_SECTIONAT'));
               Fait := TRUE;
               end;
            end;
         end;
      end
   else
      begin
      TTauxATD := TTauxAT.FindFirst(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'],
                                    [TSalD.GetValue('PSA_ETABLISSEMENT'),
                                    TSalD.GetValue('PSA_ORDREAT')], TRUE);
      While ((TTauxATD <> nil) and
            (TTauxATD.GetValue('PAT_DATEVALIDITE')>StrToDate(StPerFin))) do
            TTauxATD := TTauxAT.FindNext(['PAT_ETABLISSEMENT', 'PAT_ORDREAT'],
                                         [TSalD.GetValue('PSA_ETABLISSEMENT'),
                                         TSalD.GetValue('PSA_ORDREAT')], TRUE);

      if (TTauxATD <> nil) then
         begin
         if ((TTauxATD.GetValue('PAT_SECTIONAT') <> '') and
            (TTauxATD.GetValue('PAT_SECTIONAT') <> '98') and	//PT101-2
            (TTauxATD.GetValue('PAT_SECTIONAT') <> null)) then
            begin
            CreeDetail (CleDADS, 39, 'S41.G01.00.025',
                        TTauxATD.GetValue('PAT_SECTIONAT'),
                        TTauxATD.GetValue('PAT_SECTIONAT'));
            Fait := TRUE;
            end;
         end;
      end;
   end;
{PT101-2
if Fait = FALSE then
   begin
   CreeDetail (CleDADS, 33, 'S41.G01.00.025', '98', '98');
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Section AT a été forcé à ''98''',
                 Erreur);
   end
else
}
if ((TTauxATD <> nil) and (TTauxATD.GetValue('PAT_SECTIONAT') = '99')) then
{PT101-1
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Section AT a été forcé à ''99''',
                 Erreur);
}
   begin
   ErreurDADSU.Segment:= 'S41.G01.00.025';
   ErreurDADSU.Explication:= 'La section AT a été forcé à ''99''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT101-1
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.025');

Fait := FALSE;
{PT101-2
if ((TSalD.GetValue('PSA_REGIMESS') <> '122') and
   (TSalD.GetValue('PSA_REGIMESS') <> '200')) then
   begin
   CreeDetail (CleDADS, 34, 'S41.G01.00.026', '98888', '98888');
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Régime de base obligatoire différent de régime général :'+
                 ' le code risque AT a été forcé à ''98888''',
                Erreur);
   Fait := TRUE;
   end
else
}
Buf:= '';
if ((((TSalD.GetValue('PSA_REGIMESS') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = '-'))    //PT124
  or ((TSalD.GetValue('PSA_REGIMEAT') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = 'X')))   //PT124
{PT138
  and (BufProf<>'07')) then
}
  and (BufProf<>'07') and (BufProf<>'36') and (BufProf<>'37') and
      (BufProf<>'38') and (BufProf<>'39')) then
//FIN PT138
//FIN PT101-2
   begin
   if (TTauxATD <> nil) then
      begin
      if ((TTauxATD.GetValue('PAT_CODERISQUE') <> '') and
         (TTauxATD.GetValue('PAT_CODERISQUE') <> null)) then
         begin
         if (TTauxATD.GetValue('PAT_SECTIONAT') = '99') then
            begin
            CreeDetail (CleDADS, 40, 'S41.G01.00.026', '99999', '99999');
{PT101-1
            EcrireErreur ('          Saisie complémentaire Salarié :'+
                          ' Le code risque AT a été forcé à ''99999''',
                          Erreur);
}
            ErreurDADSU.Segment:= 'S41.G01.00.026';
            ErreurDADSU.Explication:= 'Le code risque AT a été forcé à ''99999''';
            ErreurDADSU.CtrlBloquant:= False;
            EcrireErreur(ErreurDADSU);
//FIN PT101-1
            end
         else
            begin
            Buf:= PGUpperCase(TTauxATD.GetValue('PAT_CODERISQUE'));
            if (((BufContrat <> 'CTT') and ((Buf='745BA') or (Buf='745BB') or
               (Buf='745BD'))) or ((BufProf<>'05') and (Buf='923AC'))) then
               begin
{PT101-2
               EcrireErreur ('Erreur:   Saisie complémentaire Salarié : Le'+
                            ' code risque AT a été forcé à ''''', Erreur);
               CreeDetail (CleDADS, 34, 'S41.G01.00.026', '', '');
}
               Buf:= '';
               end
            else
               CreeDetail (CleDADS, 40, 'S41.G01.00.026', Buf, Buf);
            end;
         Fait := TRUE;
         end;
      end;
   end;
{PT101-2
if Fait = FALSE then
   begin
   CreeDetail (CleDADS, 34, 'S41.G01.00.026', '98888', '98888');
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Code risque AT a été forcé à ''98888''',
                 Erreur);
   end;
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.026');

{PT101-2
if ((TTauxATD <> nil) and (TTauxATD.GetValue('PAT_CODEBUREAU') = 'BUR')) then
}
if ((TTauxATD <> nil) and (TTauxATD.GetValue('PAT_CODEBUREAU') = 'BUR') and
   (Buf <> '') and (Buf <> '753CB') and (Buf <> '911AA') and
   (Buf <> '753CA') and (Buf <> '401ZA') and (Buf <> '753CC') and
   (Buf <> '631AZ') and (Buf <> '926CC') and (Buf <> '926CF') and
   (Buf <> '631AB') and (Buf <> '923AC') and (Buf <> '926CE') and
   (Buf <> '926CD') and (Buf <> '401ZB') and (Buf <> '802CA') and
   (Buf <> '751CA') and (Buf <> '802AA') and (Buf <> '853KA') and
   (Buf <> '752EC') and (Buf <> '853CA') and (Buf <> '752EB') and
   (Buf <> '804CA') and (Buf <> '752EA') and (Buf <> '950ZD') and
   (Buf <> '950ZC') and (Buf <> '950ZB') and (Buf <> '511TG') and
   (Buf <> '950ZA') and (Buf <> '913EG') and (Buf <> '913EF') and
   (Buf <> '526GA') and (Buf <> '913EE') and (Buf <> '913ED') and
   (Buf <> '913EC') and (Buf <> '853KC') and (Buf <> '853KB') and
{PT125-6
   (Buf <> '752ED') and (Buf <> '745AB') and (Buf <> '999ZA') and
   (Buf <> '524RB') and (Buf <> '853HB') and (Buf <> '853HA') and
   (Buf <> '745BB') and (Buf <> '745BD') and (Buf <> '911AE') and
   (Buf <> '745BA') and (Buf <> '853KD') and (Buf <> '511TH') and
   (Buf <> '853KE')) then
}
   (Buf <> '752ED') and (Buf <> '745AB') and (Buf <> '524RB') and
   (Buf <> '853HB') and (Buf <> '853HA') and (Buf <> '745BB') and
   (Buf <> '745BD') and (Buf <> '911AE') and (Buf <> '745BA') and
   (Buf <> '853KD') and (Buf <> '511TH') and (Buf <> '853KE')) then
//FIN PT125-6
//FIN PT101-2
   CreeDetail (CleDADS, 41, 'S41.G01.00.027', 'B', 'B');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.027');

{PT101-2
Fait := FALSE;
if ((TSalD.GetValue('PSA_REGIMESS') <> '122') and
   (TSalD.GetValue('PSA_REGIMESS') <> '200')) then
   begin
   CreeDetail (CleDADS, 36, 'S41.G01.00.028', '98888', FloatToStr(9888/100));
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Régime de base obligatoire différent de régime général :'+
                 ' le taux AT a été forcé à ''98888''',
                Erreur);
   Fait := TRUE;
   end
else
}
if ((((TSalD.GetValue('PSA_REGIMESS') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = '-'))    //PT124
  or ((TSalD.GetValue('PSA_REGIMEAT') = '200') and (TSalD.GetValue('PSA_TYPEREGIME') = 'X')))   //PT124
{PT138
  and (BufProf<>'07')) then
}
  and (BufProf<>'07') and (BufProf<>'36') and (BufProf<>'37') and
      (BufProf<>'38') and (BufProf<>'39')) then
//FIN PT138
//FIN PT101-2
   begin
   if (TTauxATD <> nil) then
      begin
      if (TTauxATD.GetValue('PAT_SECTIONAT') = '99') then
         begin
         CreeDetail (CleDADS, 42, 'S41.G01.00.028', '99999',
                     FloatToStr(9999/100));
{PT101-1
         EcrireErreur ('          Saisie complémentaire Salarié :'+
                       ' Le Taux AT a été forcé à ''99999''',
                       Erreur);
}
         ErreurDADSU.Segment:= 'S41.G01.00.028';
         ErreurDADSU.Explication:= 'Le Taux AT a été forcé à ''99999''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT101-1
         Fait := TRUE;
         end
      else
         begin
         Total := TTauxATD.GetValue('PAT_TAUXAT');
         CreeDetail (CleDADS, 42, 'S41.G01.00.028',
                     FloatToStr(Arrondi(Total*100, 0)),
                     FloatToStr(Arrondi(Total, 2)));
         if Total <> 0 then
            Fait := TRUE;
         end;
      end;
   end;
{PT101-2
if Fait = FALSE then
   begin
   CreeDetail (CleDADS, 36, 'S41.G01.00.028', '98888', FloatToStr(9888/100));
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Le Taux AT a été forcé à ''98888''', Erreur);
   end;
}
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.028');

Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_BRUTSS'], [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if (((TBaseD.GetValue('PHB_DATEFIN')>StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')<=StrToDate(StPerFin))) or
         ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')>=StrToDate(StPerFin)))) then
         Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BRUTSS'],
                               [Salarie, 'X'], TRUE);
      end;
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BRUTSS'],
                               [Salarie, 'X'], TRUE);
      end;

{PT101-4
if (BufContrat='CAP') then
}
{PT125-6
if ((BufContrat='CAP') or (BufProf='07')) then
}
{PT140
if ((BufContrat='CAP') or (BufProf='07') or (BufProf='36')) then
}
if ((BufContrat='CAA') or (BufProf='07') or (BufProf='36')) then
//FIN PT140
//FIN PT125-6
   Total:= 0;
BaseSS := Arrondi(Total, 0);
CreeDetail (CleDADS, 43, 'S41.G01.00.029.001',
            FloatToStr(Abs(BaseSS)), FloatToStr(BaseSS));
if (Total < 0) then
   CreeDetail (CleDADS, 44, 'S41.G01.00.029.002', 'N', 'N');

//PT140
if ((BufContrat='CAA') or (BufContrat='CAP')) then
   CreeDetail (CleDADS, 45, 'S41.G01.00.029.003', '02', '02')
else
//FIN PT140
   CreeDetail (CleDADS, 45, 'S41.G01.00.029.003', '01', '01');   //PT125-6
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.029');

if (BufContrat='CAP') then
   CreeDetail (CleDADS, 46, 'S41.G01.00.030.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)))
else
   begin
   Total := 0;
   TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_PLAFONDSS'],
                             [Salarie, 'X'], TRUE);
   While ((TBaseD <> nil) and
         ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
         begin
         if (((TBaseD.GetValue('PHB_DATEFIN')>StrToDate(StPerDeb)) and
            (TBaseD.GetValue('PHB_DATEFIN')<=StrToDate(StPerFin))) or
            ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) and
            (TBaseD.GetValue('PHB_DATEFIN')>=StrToDate(StPerFin)))) then
            Total:= Total+TBaseD.GetValue('PHB_TRANCHE1');
         TBaseD:= TBase.FindNext(['PHB_SALARIE', 'PCT_PLAFONDSS'],
                                 [Salarie, 'X'], TRUE);
         end;
   While ((TBaseD <> nil) and
         (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         Total := Total + TBaseD.GetValue('PHB_TRANCHE1');
         TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_PLAFONDSS'],
                                  [Salarie, 'X'], TRUE);
         end;
{PT101-4
   if (TSalD.GetValue('PSA_REGIMESS') <> '200') then
}
{PT124
   if ((TSalD.GetValue('PSA_REGIMESS') <> '200') or (BufProf='07')) then
}
   if ((((TSalD.GetValue ('PSA_REGIMESS')<>'200') and
      (TSalD.GetValue ('PSA_TYPEREGIME')='-')) or
      ((TSalD.GetValue ('PSA_REGIMEAT')<>'200') and
      (TSalD.GetValue ('PSA_TYPEREGIME')='X'))) or 
{PT125-6
    (BufProf='07')) then
}
    (BufProf='07') or (BufProf='36')) then
//FIN PT125-6
//FIN PT124
      Total := 0;
   CreeDetail (CleDADS, 46, 'S41.G01.00.030.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total < 0) then
      CreeDetail (CleDADS, 47, 'S41.G01.00.030.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.030');

//PGVisuUnObjet(TBase, '', '');
Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE','PCT_BASECSGCRDS'],
                          [Salarie,'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TBaseD := TBase.FindNext(['PHB_SALARIE','PCT_BASECSGCRDS'],
                               [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BASECSGCRDS'],
                               [Salarie, 'X'], TRUE);
      end;

THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                        [Salarie, '15'], TRUE);
While ((THistoCumSalD <> nil) and
      ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
         (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
         Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '15'], TRUE);
      end;
While ((THistoCumSalD <> nil) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '15'], TRUE);
      end;

{PT140
if (BufContrat='CAP') then
}
if ((BufContrat='CAA') or (BufContrat='CAP')) then
//FIN PT140
   Total:= 0;
CreeDetail (CleDADS, 48, 'S41.G01.00.032.001',
            FloatToStr(Abs(Arrondi(Total, 0))), FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 49, 'S41.G01.00.032.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.032');

Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_BASECRDS'],
                          [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BASECRDS'],
                               [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BASECRDS'],
                               [Salarie, 'X'], TRUE);
      end;

THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                        [Salarie, '13'], TRUE);
While ((THistoCumSalD <> nil) and
      ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
         (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
         Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '13'], TRUE);
      end;
While ((THistoCumSalD <> nil) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '13'], TRUE);
      end;

{PT140
if (BufContrat='CAP') then
}
if ((BufContrat='CAA') or (BufContrat='CAP')) then
//FIN PT140
   Total:= 0;
CreeDetail (CleDADS, 50, 'S41.G01.00.033.001',
            FloatToStr(Abs(Arrondi(Total, 0))), FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 51, 'S41.G01.00.033.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.033');

Buf := '';
if ((TSalD.GetValue('PSA_TRAVETRANG') <> '') and
   (TSalD.GetValue('PSA_TRAVETRANG') <> null)) then
   Buf := TSalD.GetValue('PSA_TRAVETRANG');
if (Buf = 'F') then
   CreeDetail (CleDADS, 52, 'S41.G01.00.034', '01', Buf)
else
if (Buf = 'E') then
   CreeDetail (CleDADS, 52, 'S41.G01.00.034', '02', Buf);
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.034');

BaseBruteFisc:= 0;
{PT115
if (BufContrat = 'CAA') then
}
if ((BufContrat = 'CAA') or (BufContrat = 'CAU') or (BufContrat = 'CAV') or
   (BufContrat = 'CEJ') or (BufContrat = 'CES')) then
//FIN PT115
   BaseBruteFisc:= 0
else
if (BufContrat = 'CAP') then
   begin
   TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_DADSEXOBASE'], [Salarie, '02'], TRUE);
   While ((TBaseD <> nil) and
         ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
         TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSEXOBASE'], [Salarie, '02'], TRUE);
   While ((TBaseD <> nil) and
         (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         BaseBruteFisc:= BaseBruteFisc + TBaseD.GetValue('PHB_BASECOT');
         TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSEXOBASE'], [Salarie, '02'], TRUE);
         end;
   end
else
   begin
   THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                           [Salarie, '02'], TRUE);
   While ((THistoCumSalD <> nil) and
         ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
         (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
         begin
         if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
            (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
            BaseBruteFisc:= BaseBruteFisc+THistoCumSalD.GetValue('PHC_MONTANT');
         THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '02'], TRUE);
         end;
   While ((THistoCumSalD <> nil) and
         (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         BaseBruteFisc:= BaseBruteFisc+THistoCumSalD.GetValue('PHC_MONTANT');
         THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                [Salarie, '02'], TRUE);
         end;
   end;

CreeDetail (CleDADS, 53, 'S41.G01.00.035.001',
            FloatToStr (Abs (Int (BaseBruteFisc))),
            FloatToStr (Int (BaseBruteFisc)));
if (BaseBruteFisc < 0) then
   CreeDetail (CleDADS, 54, 'S41.G01.00.035.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.035');

Tot1 := 0;
Tot2 := 0;
Tot3 := 0;
Tot4 := 0;
Tot5 := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                        [Salarie, 'N'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'N'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Tot1 := Tot1 + TRemD.GetValue('PHB_MTREM')
      else
         Tot1 := Tot1 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'N'], TRUE);
      end;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                        [Salarie, 'L'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'L'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Tot2 := Tot2 + TRemD.GetValue('PHB_MTREM')
      else
         Tot2 := Tot2 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'L'], TRUE);
      end;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                        [Salarie, 'V'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'V'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Tot3 := Tot3 + TRemD.GetValue('PHB_MTREM')
      else
         Tot3 := Tot3 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'V'], TRUE);
      end;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                        [Salarie, 'A'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'A'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Tot4 := Tot4 + TRemD.GetValue('PHB_MTREM')
      else
         Tot4 := Tot4 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'A'], TRUE);
      end;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                        [Salarie, 'T'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'T'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Tot5 := Tot5 + TRemD.GetValue('PHB_MTREM')
      else
         Tot5 := Tot5 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_AVANTAGENAT'],
                             [Salarie, 'T'], TRUE);
      end;
Total:= Tot1 + Tot2 + Tot3 + Tot4 + Tot5;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 55, 'S41.G01.00.037.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 56, 'S41.G01.00.037.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.037');

if (Tot1 <> 0) then
   CreeDetail (CleDADS, 57, 'S41.G01.00.038', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.038');

if (Tot2 <> 0) then
   CreeDetail (CleDADS, 58, 'S41.G01.00.039', 'L', 'L');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.039');

if (Tot3 <> 0) then
   CreeDetail (CleDADS, 59, 'S41.G01.00.040', 'V', 'V');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.040');

if (Tot4 <> 0) then
   CreeDetail (CleDADS, 60, 'S41.G01.00.041', 'A', 'A');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.041');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_RETSALAIRE'],
                        [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_RETSALAIRE'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_RETSALAIRE'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 61, 'S41.G01.00.042.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 62, 'S41.G01.00.042.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.042');

if (TSalD.GetValue('PSA_REMPOURBOIRE') = 'X') then
   CreeDetail (CleDADS, 63, 'S41.G01.00.043', 'P', 'P');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.043');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_FRAISPROF'], [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_FRAISPROF'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_FRAISPROF'],
                             [Salarie, 'X'], TRUE);
      end;
BoolFraisProf := FALSE;

if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 64, 'S41.G01.00.044.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 65, 'S41.G01.00.044.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.044');

   if (TSalD.GetValue('PSA_ALLOCFORFAIT') = 'X') then
      begin
      BoolFraisProf := TRUE;
      CreeDetail (CleDADS, 66, 'S41.G01.00.045', 'F', 'F');
      end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.045');

   if (TSalD.GetValue('PSA_REMBJUSTIF') = 'X') then
      begin
      BoolFraisProf := TRUE;
      CreeDetail (CleDADS, 67, 'S41.G01.00.046', 'R', 'R');
      end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.046');

   if (TSalD.GetValue('PSA_PRISECHARGE') = 'X') then
      begin
      BoolFraisProf := TRUE;
      CreeDetail (CleDADS, 68, 'S41.G01.00.047', 'P', 'P');
      end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.047');

   if (TSalD.GetValue('PSA_AUTREDEPENSE') = 'X') then
      begin
      BoolFraisProf := TRUE;
      CreeDetail (CleDADS, 69, 'S41.G01.00.048', 'D', 'D');
      end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.048');
   if (BoolFraisProf = FALSE) then
{PT101-1
      EcrireErreur ('Erreur:   Saisie complémentaire Salarié : Le montant des'+
                    ' frais professionnels est égal à '+FloatToStr(Arrondi(Total, 0))+#13#10+
                    '                 mais aucune zone correspondante n''est cochée',
                   Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S41.G01.00.044.001';
      ErreurDADSU.Explication:= 'Le montant des frais professionnels est'+
                                ' renseigné mais aucune zone correspondante'+
                                ' n''est cochée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   end;

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_CHQVACANCE'],
                        [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_CHQVACANCE'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_CHQVACANCE'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 70, 'S41.G01.00.049.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 71, 'S41.G01.00.049.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.049');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_IMPOTRETSOURC'],
                        [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_IMPOTRETSOURC'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_IMPOTRETSOURC'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 72, 'S41.G01.00.052.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 73, 'S41.G01.00.052.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.052');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_INDEXPATRIAT'],
                        [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_INDEXPATRIAT'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_INDEXPATRIAT'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 74, 'S41.G01.00.053.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 75, 'S41.G01.00.053.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.053');

Total := 0;
Tot1 := 0;
Tot2 := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_DADSTOTIMPTSS'],
                          [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSTOTIMPTSS'],
                               [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_BASECOT');
      Tot1 := Tot1 + TBaseD.GetValue('PHB_TRANCHE2');
      Tot2 := Tot2 + TBaseD.GetValue('PHB_TRANCHE3');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSTOTIMPTSS'],
                               [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 76, 'S41.G01.00.055.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 77, 'S41.G01.00.055.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.055');

if (Tot1 <> 0) then
   begin
   CreeDetail (CleDADS, 78, 'S41.G01.00.056.001', FloatToStr (Abs (Int (Tot1))),
               FloatToStr (Int (Tot1)));
   if (Tot1 < 0) then
      CreeDetail (CleDADS, 79, 'S41.G01.00.056.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.056');

if (Tot2 <> 0) then
   begin
   CreeDetail (CleDADS, 80, 'S41.G01.00.057.001', FloatToStr (Abs (Int (Tot2))),
               FloatToStr (Int (Tot2)));
   if (Tot2 < 0) then
      CreeDetail (CleDADS, 81, 'S41.G01.00.057.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.057');

Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_DADSMONTTSS'],
                          [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSMONTTSS'],
                               [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_MTPATRONAL');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSMONTTSS'],
                               [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 82, 'S41.G01.00.058.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 83, 'S41.G01.00.058.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.058');

Tot1 := 0;
Tot2 := 0;
THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                        [Salarie, '09'], TRUE);
While ((THistoCumSalD <> nil) and
      ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
         (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
         Tot1 := Tot1 + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '09'], TRUE);
      end;
While ((THistoCumSalD <> nil) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Tot1 := Tot1 + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '09'], TRUE);
      end;
{PT126
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_HCHOMPAR'], [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_HCHOMPAR'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')='M') then
         Tot2 := Tot2 + TRemD.GetValue('PHB_MTREM')
      else
         Tot2 := Tot2 - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_HCHOMPAR'],
                             [Salarie, 'X'], TRUE);
      end;
Total := Tot1-Tot2;

CreeDetail (CleDADS, 84, 'S41.G01.00.063.001', FloatToStr (Abs (Int (Total))),
            FloatToStr (Int (Total)));
if Total < 0 then
}
CreeDetail (CleDADS, 84, 'S41.G01.00.063.001', FloatToStr (Abs (Int (Tot1))),
            FloatToStr (Int (Tot1)));
if (Tot1<0) then
//FIN PT126
   CreeDetail (CleDADS, 85, 'S41.G01.00.063.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.063');

Total := 0;
TRemD := TRem.FindFirst(['PHB_SALARIE', 'PRM_INDEMCP'], [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      ((TRemD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TRemD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_INDEMCP'],
                             [Salarie, 'X'], TRUE);
While ((TRemD <> nil) and
      (TRemD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TRemD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TRemD.GetValue('PHB_MTREM')
      else
         Total := Total - TRemD.GetValue('PHB_MTREM');
      TRemD := TRem.FindNext(['PHB_SALARIE', 'PRM_INDEMCP'],
                             [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 86, 'S41.G01.00.065.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if Total < 0 then
      CreeDetail (CleDADS, 87, 'S41.G01.00.065.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.065');

Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_DADSEPARGNE'],
                          [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if (((TBaseD.GetValue('PHB_DATEFIN')>StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')<StrToDate(StPerFin))) or
         ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')>StrToDate(StPerFin)))) then
         Total:= Total+TBaseD.GetValue('PHB_MTPATRONAL')+
                 TBaseD.GetValue('PHB_MTSALARIAL');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSEPARGNE'],
                               [Salarie, 'X'], TRUE);
      end;
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total:= Total+TBaseD.GetValue('PHB_MTPATRONAL')+
              TBaseD.GetValue('PHB_MTSALARIAL');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_DADSEPARGNE'],
                               [Salarie, 'X'], TRUE);
      end;
if (Total <> 0) then
   begin
   CreeDetail (CleDADS, 90, 'S41.G01.00.067.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total < 0) then
      CreeDetail (CleDADS, 91, 'S41.G01.00.067.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.067');

if (Tot5 > 0) then
   CreeDetail (CleDADS, 92, 'S41.G01.00.068', 'T', 'T');
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.068');

//PT125-6
Total:= 0;
TCotD:= TCot.FindFirst (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'], [Salarie, 'X'], TRUE);
While ((TCotD<>nil) and
      ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) or
      (TCotD.GetValue ('PHB_DATEDEBUT')>StrToDate (StPerFin)))) do
      begin
      if (((TCotD.GetValue ('PHB_DATEFIN')>StrToDate (StPerDeb)) and
         (TCotD.GetValue ('PHB_DATEFIN')<=StrToDate (StPerFin))) or
         ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) and
         (TCotD.GetValue ('PHB_DATEFIN')>=StrToDate (StPerFin)))) then
         Total:= Total+TCotD.GetValue ('PHB_BASECOT');
      TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'],
                             [Salarie, 'X'], TRUE);
      end;
While ((TCotD<>nil) and
      (TCotD.GetValue ('PHB_DATEDEBUT')>=StrToDate (StPerDeb)) and
      (TCotD.GetValue ('PHB_DATEDEBUT')<=StrToDate (StPerFin))) do
      begin
      Total:= Total+TCotD.GetValue ('PHB_BASECOT');
      TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'],
                             [Salarie, 'X'], TRUE);
      end;

if (Total<>0) then           //PT134-2
   begin
   CreeDetail (CleDADS, 100, 'S41.G01.00.073.001',
               FloatToStr (Abs (Int (Total))), FloatToStr (Int (Total)));
   if (Total<0) then
      CreeDetail (CleDADS, 101, 'S41.G01.00.073.002', 'N', 'N');
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.00.073');
//FIN PT125-6

HistoOrg := '';
NumOrdre := 1;
Ircantec:= False;      //PT118
Agirc:= False;
{
if V_PGI.Debug=True then
   PGVisuUnObjet(THistoOrg, '', '');
}
NbreHistoOrg := THistoOrg.FillesCount(1);
if NbreHistoOrg > 0 then
   begin
   TInstitution := TOB.Create('Les Institutions', NIL, -1);
   TPaieD := TPaie.FindFirst(['PPU_SALARIE'], [Salarie], TRUE);
   While ((TPaieD <> nil) and
         ((TPaieD.GetValue('PPU_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TPaieD.GetValue('PPU_DATEDEBUT')>StrToDate(StPerFin)))) do
         TPaieD := TPaie.FindNext(['PPU_SALARIE'], [Salarie], TRUE);
   if ((TPaieD <> nil) and
      (TPaieD.GetValue('PPU_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TPaieD.GetValue('PPU_DATEDEBUT')<=StrToDate(StPerFin))) then
      begin
      THistoOrgD := THistoOrg.FindFirst(['PHB_SALARIE', 'POG_ETABLISSEMENT'],
                                        [Salarie, TPaieD.GetValue('PPU_ETABLISSEMENT')],
                                        TRUE);
      if (THistoOrgD <> nil) then
         begin
         While (THistoOrgD <> nil) do
               begin
               if ((THistoOrgD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
                  (THistoOrgD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin)) and
                  (THistoOrgD.GetValue('POG_INSTITUTION') <> null)) then
                  begin
                  if (THistoOrgD.GetValue('POG_REGROUPDADSU')='X') then
                     TInstitD := TInstitution.FindFirst(['INSTITUTION'],
                                                        [THistoOrgD.GetValue('POG_REGROUPEMENT')],
                                                        TRUE)
                  else
                     TInstitD := TInstitution.FindFirst(['INSTITUTION'],
                                                        [THistoOrgD.GetValue('POG_INSTITUTION')],
                                                        TRUE);

                  if (TInstitD = nil) then
                     begin
                     TInstitutionFille := TOB.Create('',TInstitution,-1);
                     TInstitutionFille.AddChampSup('INSTITUTION', FALSE);
                     TInstitutionFille.AddChampSup('NUMAFFILIATION', FALSE);
                     if (THistoOrgD.GetValue('POG_REGROUPDADSU')='X') then
                        begin
                        TInstitutionFille.PutValue('INSTITUTION',
                                                   THistoOrgD.GetValue('POG_REGROUPEMENT'));
                        TOrganismeD := TOrganisme.FindFirst(['POG_INSTITUTION'],
                                                            [THistoOrgD.GetValue('POG_REGROUPEMENT')],
                                                            TRUE);
                        if (TOrganismeD <> nil) then
                           TInstitutionFille.PutValue('NUMAFFILIATION',
                                                      TOrganismeD.GetValue('POG_NUMAFFILIATION'))
                        else
                           TInstitutionFille.PutValue('NUMAFFILIATION',
                                                      THistoOrgD.GetValue('POG_NUMAFFILIATION'));
                        end
                     else
                        begin
                        TInstitutionFille.PutValue('INSTITUTION',
                                                   THistoOrgD.GetValue('POG_INSTITUTION'));
                        TInstitutionFille.PutValue('NUMAFFILIATION',
                                                   THistoOrgD.GetValue('POG_NUMAFFILIATION'));
                        end;
                     end;
                  end;
               THistoOrgD := THistoOrg.FindNext(['PHB_SALARIE', 'POG_ETABLISSEMENT'],
                                                [Salarie, TPaieD.GetValue('PPU_ETABLISSEMENT')],
                                                TRUE);
               end;

         NbreInst:= TInstitution.FillesCount (1);
         For i := 0 to NbreInst-1 do
             begin
             TInstitD:= TInstitution.Detail[i];
             HistoOrg:= TInstitD.GetValue ('INSTITUTION');
             if ((HistoOrg <> '') and (HistoOrg <> null) and
                (not (isnumeric (HistoOrg)))) then
                begin
{PT104
                if ((TypeD>='002') and (TypeD<='005') and
                   (Pos('B', HistoOrg)=1)) then
                   begin
                   Buf:=Copy(HistoOrg, 3, 2);
                   CreeDetail (CleDADS, 100+NumOrdre, 'S41.G01.01.001', Buf,
                               HistoOrg);
                   end
                else
                   CreeDetail (CleDADS, 100+NumOrdre, 'S41.G01.01.001',
                               HistoOrg, HistoOrg);
}
                if (((TypeD>='002') and (TypeD<='005')) or
                   ((TypeD>='202') and (TypeD<='205'))) then
                   begin
                   if (Pos ('B', HistoOrg)=1) then
                      begin
                      Buf:=Copy (HistoOrg, 3, 2);
                      CreeDetail (CleDADS, 110+NumOrdre, 'S41.G01.01.001', Buf,
                                  HistoOrg);
//PT112-3
                      if ((TInstitD.GetValue ('NUMAFFILIATION') <> '') and
                         (TInstitD.GetValue ('NUMAFFILIATION') <> null)) then
                         CreeDetail (CleDADS, 110+NumOrdre+1, 'S41.G01.01.002',
                                     TInstitD.GetValue ('NUMAFFILIATION'),
                                     TInstitD.GetValue ('NUMAFFILIATION'));
                      NumOrdre:= NumOrdre+2;
//FIN PT112-3
                      end;
                   end
                else
                   begin
                   if (Pos ('B', HistoOrg)<>1) then
                      begin
                      if (Pos('I', HistoOrg)=1) then
{PT106
                         Buf:= Copy (HistoOrg, 1, 1)+'0'+Copy(HistoOrg, 2, 3)
}
                         begin
                         Buf:= Copy (HistoOrg, 1, 1)+'0'+Copy(HistoOrg, 2, 3);
                         RemplitIRCANTEC (Salarie, Periode);
                         Ircantec:= True;          //PT118
                         end
//FIN PT106
                      else
                         Buf:= HistoOrg;
                      CreeDetail (CleDADS, 110+NumOrdre, 'S41.G01.01.001',
                                  Buf, HistoOrg);
//PT112-3
                      if ((TInstitD.GetValue ('NUMAFFILIATION') <> '') and
                         (TInstitD.GetValue ('NUMAFFILIATION') <> null)) then
                         CreeDetail (CleDADS, 110+NumOrdre+1, 'S41.G01.01.002',
                                     TInstitD.GetValue ('NUMAFFILIATION'),
                                     TInstitD.GetValue ('NUMAFFILIATION'));
                      NumOrdre:= NumOrdre+2;
//FIN PT112-3
                      end;
                   end;
//FIN PT104
{PT112-3
                if ((TInstitD.GetValue ('NUMAFFILIATION') <> '') and
                   (TInstitD.GetValue ('NUMAFFILIATION') <> null)) then
                   CreeDetail (CleDADS, 100+NumOrdre+1, 'S41.G01.01.002',
                               TInstitD.GetValue ('NUMAFFILIATION'),
                               TInstitD.GetValue ('NUMAFFILIATION'));
                NumOrdre:= NumOrdre+2;
}
                end
             else
                HistoOrg:='';

//PT125-2
             if ((Pos ('A', HistoOrg)=1) or (Pos ('C', HistoOrg)=1) or
                (Pos ('G', HistoOrg)=1)) then
                begin
                Agirc:= True;
                if (Buf2='N') then
                   begin
                   ErreurDADSU.Segment:= 'S41.G01.00.013';
                   ErreurDADSU.Explication:= 'La caractéristique activité ne'+
                                             ' convient pas à ce régime'+
                                             ' complémentaire';
                   ErreurDADSU.CtrlBloquant:= True;
                   EcrireErreur (ErreurDADSU);
                   end;

                if ((BufProf<>'01') and (BufProf<>'02') and (BufProf<>'03') and
                   (BufProf<>'04') and (BufProf<>'05') and (BufProf<>'06') and
                   (BufProf<>'07') and (BufProf<>'08') and (BufProf<>'09') and
                   (BufProf<>'11') and (BufProf<>'12') and (BufProf<>'13') and
                   (BufProf<>'14') and (BufProf<>'15') and (BufProf<>'16') and
                   (BufProf<>'18') and (BufProf<>'19') and (BufProf<>'20') and
                   (BufProf<>'21') and (BufProf<>'22') and (BufProf<>'23') and
                   (BufProf<>'24') and (BufProf<>'25') and (BufProf<>'26') and
                   (BufProf<>'27') and (BufProf<>'28') and (BufProf<>'75') and
                   (BufProf<>'90')) then
                   begin
                   ErreurDADSU.Segment:= 'S41.G01.00.014';
                   ErreurDADSU.Explication:= 'Le statut professionnel ne'+
                                             ' convient pas à ce régime'+
                                             ' complémentaire';
                   ErreurDADSU.CtrlBloquant:= True;
                   EcrireErreur(ErreurDADSU);
                   end;
                end;
//FIN PT125-2
             end;
         end;
      end;
   end;
if ((HistoOrg = '') or
   (HistoOrg = null)) then
   begin
   CreeDetail (CleDADS, 110+NumOrdre, 'S41.G01.01.001', '90000', '90000');
{PT101-1
   EcrireErreur ('          Saisie complémentaire Salarié :'+
                 ' Code régime complémentaire a été forcé à ''90000''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S41.G01.01.001';
   ErreurDADSU.Explication:= 'Le code régime complémentaire a été forcé à'+
                             ' ''90000''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S41.G01.01.001');

//PT101-4
if (BufProf='07') then
   begin
   if (BaseBruteFisc<>0) then //PT102-2
      begin
      CreeDetail (CleDADS, 201, 'S41.G01.02.001', '21', '21');
      CreeDetail (CleDADS, 202, 'S41.G01.02.002.001',
                  FloatToStr (Abs (Int (BaseBruteFisc))),
                  FloatToStr (Int (BaseBruteFisc)));
      if (BaseBruteFisc < 0) then
         CreeDetail (CleDADS, 203, 'S41.G01.02.002.002', 'N', 'N');
{PT125-4
      end;
}

      Total:= 0;
      THistoCumSalD:= THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '30'], TRUE);
      While ((THistoCumSalD <> nil) and
            ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
            (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
            begin
            if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
               (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
               Total:= Total+THistoCumSalD.GetValue('PHC_MONTANT');
            THistoCumSalD:= THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                  [Salarie, '30'], TRUE);
            end;
      While ((THistoCumSalD <> nil) and
            (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
            (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
            begin
            Total:= Total+THistoCumSalD.GetValue('PHC_MONTANT');
            THistoCumSalD:= THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                                  [Salarie, '30'], TRUE);
            end;

      if (BaseBruteFisc<Total) then
         begin
         if (BaseBruteFisc<>0) then //PT102-2
            begin
            CreeDetail (CleDADS, 301, 'S41.G01.03.001', '21', '21');
            CreeDetail (CleDADS, 302, 'S41.G01.03.002.001',
                        FloatToStr (Abs (Int (BaseBruteFisc))),
                        FloatToStr (Int (BaseBruteFisc)));
            if (BaseBruteFisc < 0) then
               CreeDetail (CleDADS, 303, 'S41.G01.03.002.002', 'N', 'N');
            end;
         end
      else
         begin
         if (BaseBruteFisc<>0) then //PT102-2
            begin
            CreeDetail (CleDADS, 301, 'S41.G01.03.001', '21', '21');
            CreeDetail (CleDADS, 302, 'S41.G01.03.002.001',
                        FloatToStr(Abs(Arrondi(Total, 0))),
                        FloatToStr(Arrondi(Total, 0)));
            if (Total < 0) then
               CreeDetail (CleDADS, 303, 'S41.G01.03.002.002', 'N', 'N');
            end
//PT125-4
         else
            begin
            ErreurDADSU.Segment:= 'S41.G01.03.001';
            ErreurDADSU.Explication:= 'Une base brute exceptionnelle est'+
                                      ' renseignée sans base plafonnée'+
                                      ' exceptionnelle';
            ErreurDADSU.CtrlBloquant:= True;
            EcrireErreur (ErreurDADSU);
            end;
//FIN PT125-4
         end;
      end;   //PT125-4
   end;
//FIN PT101-4

NumOrdre := 400;
{PT109
GereSomIso (TRem, Salarie, '01', TypeD, Periode, NumOrdre);
GereSomIso (TRem, Salarie, '02', TypeD, Periode, NumOrdre);
GereSomIso (TRem, Salarie, '03', TypeD, Periode, NumOrdre);
GereSomIso (TRem, Salarie, '04', TypeD, Periode, NumOrdre);
GereSomIso (TRem, Salarie, '05', TypeD, Periode, NumOrdre);

NumOrdre := 600;
if (BufContrat='CAA') then         //PT101-5
   GereExo (THistoCumSal, 'PHC_CUMULPAIE', Salarie, '01', TypeD, Periode,
            NumOrdre, BufContrat);
if (BufContrat='CAP') then        //PT101-5
   GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '02', TypeD, Periode, NumOrdre,
            BufContrat);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '03', TypeD, Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '06', TypeD, Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '07', TypeD, Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '09', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '10', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '11', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '12', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '13', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '14', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '15', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '16', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '17', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '18', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '19', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '20', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '21', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '22', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '23', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '24', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '26', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '27', TypeD, Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '31', TypeD, Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '33', TypeD, Periode, NumOrdre);
//PT101-2
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '34', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '35', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '36', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '37', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '38', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '39', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '40', TypeD, Periode, NumOrdre);
//FIN PT101-2
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '51', TypeD, Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '52', TypeD, Periode, NumOrdre);
}
GereSomIso (TRem, Salarie, '01', Periode, NumOrdre);
{PT125-4
GereSomIso (TRem, Salarie, '02', Periode, NumOrdre);
}
GereSomIso (TRem, Salarie, '03', Periode, NumOrdre);
//PT125-4
if (Agirc=False) then
   GereSomIso (TRem, Salarie, '04', Periode, NumOrdre);
{
GereSomIso (TRem, Salarie, '05', Periode, NumOrdre);
}
//FIN PT125-4

NumOrdre := 600;
if (BufContrat='CAA') then         //PT101-5
   GereExo (THistoCumSal, 'PHC_CUMULPAIE', Salarie, '01', Periode, NumOrdre,
            BufContrat);
if (BufContrat='CAP') then        //PT101-5
   GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '02', Periode, NumOrdre,
            BufContrat);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '03', Periode, NumOrdre);
{PT125-6
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '06', Periode, NumOrdre);
}
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '07', Periode, NumOrdre);
{PT125-6
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '09', Periode, NumOrdre);
}
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '10', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '11', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '12', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '13', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '14', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '15', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '16', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '17', Periode, NumOrdre);
{PT125-6
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '18', Periode, NumOrdre);
}
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '19', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '20', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '21', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '22', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '23', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '24', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '26', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '27', Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '31', Periode, NumOrdre);
GereExo (TCot, 'PCT_DADSEXOCOT', Salarie, '33', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '34', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '35', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '36', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '37', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '38', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '39', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '40', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '51', Periode, NumOrdre);
GereExo (TBase, 'PCT_DADSEXOBASE', Salarie, '52', Periode, NumOrdre);
//FIN PT109

//PT125-6
Total:= RemplitFromPublicotis (Salarie, '171', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 801, 'S41.G30.10.001', '01', '01');
   CreeDetail (CleDADS, 802, 'S41.G30.10.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 803, 'S41.G30.10.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '172', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 804, 'S41.G30.10.001', '02', '02');
   CreeDetail (CleDADS, 805, 'S41.G30.10.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 806, 'S41.G30.10.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '173', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 807, 'S41.G30.10.001', '03', '03');
   CreeDetail (CleDADS, 808, 'S41.G30.10.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 809, 'S41.G30.10.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '174', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 810, 'S41.G30.10.001', '04', '04');
   CreeDetail (CleDADS, 811, 'S41.G30.10.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 812, 'S41.G30.10.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '191', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 831, 'S41.G30.15.001', '01', '01');
   CreeDetail (CleDADS, 832, 'S41.G30.15.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 833, 'S41.G30.15.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '192', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 834, 'S41.G30.15.001', '02', '02');
   CreeDetail (CleDADS, 835, 'S41.G30.15.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 836, 'S41.G30.15.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '201', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 841, 'S41.G30.20.001', '01', '01');
   CreeDetail (CleDADS, 842, 'S41.G30.20.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 843, 'S41.G30.20.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '202', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 844, 'S41.G30.20.001', '02', '02');
   CreeDetail (CleDADS, 845, 'S41.G30.20.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 846, 'S41.G30.20.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '203', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 847, 'S41.G30.20.001', '03', '03');
   CreeDetail (CleDADS, 848, 'S41.G30.20.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 849, 'S41.G30.20.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '204', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 850, 'S41.G30.20.001', '04', '04');
   CreeDetail (CleDADS, 851, 'S41.G30.20.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 852, 'S41.G30.20.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '211', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 861, 'S41.G30.25.001', '01', '01');
   CreeDetail (CleDADS, 862, 'S41.G30.25.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 863, 'S41.G30.25.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '212', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 864, 'S41.G30.25.001', '02', '02');
   CreeDetail (CleDADS, 865, 'S41.G30.25.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 866, 'S41.G30.25.002.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '221', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 871, 'S41.G30.30.001.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 872, 'S41.G30.30.001.002', 'N', 'N');
   end;

Total:= RemplitFromPublicotis (Salarie, '222', CleDADS.DateDeb, CleDADS.DateFin);
if (Total<>0) then
   begin
   CreeDetail (CleDADS, 873, 'S41.G30.30.002.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total<0) then
      CreeDetail (CleDADS, 874, 'S41.G30.30.002.002', 'N', 'N');
   end;

Tot1:= 0;  //Pour le montant de la rémunération
Tot2:= 0;  //Pour le montant de la réduction salariale
TCotD:= TCot.FindFirst (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'], [Salarie, 'X'], TRUE);
While ((TCotD<>nil) and
      ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) or
      (TCotD.GetValue ('PHB_DATEDEBUT')>StrToDate (StPerFin)))) do
      begin
      if (((TCotD.GetValue ('PHB_DATEFIN')>StrToDate (StPerDeb)) and
         (TCotD.GetValue ('PHB_DATEFIN')<=StrToDate (StPerFin))) or
         ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) and
         (TCotD.GetValue ('PHB_DATEFIN')>=StrToDate (StPerFin)))) then
         begin
         Tot1:= Tot1+TCotD.GetValue ('PHB_BASECOT');
         Tot2:= Tot2+TCotD.GetValue ('PHB_MTSALARIAL');
         end;
      TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'],
                             [Salarie, 'X'], TRUE);
      end;
While ((TCotD<>nil) and
      (TCotD.GetValue ('PHB_DATEDEBUT')>=StrToDate (StPerDeb)) and
      (TCotD.GetValue ('PHB_DATEDEBUT')<=StrToDate (StPerFin))) do
      begin
      Tot1:= Tot1+TCotD.GetValue ('PHB_BASECOT');
      Tot2:= Tot2+TCotD.GetValue ('PHB_MTSALARIAL');
      TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_ALLEGEMENTA2'],
                             [Salarie, 'X'], TRUE);
      end;
if (Tot1<>0) then
   begin
   if (Buf2<>'P') then
      CreeDetail (CleDADS, 881, 'S41.G30.35.001.001', '01','01')
   else
      CreeDetail (CleDADS, 881, 'S41.G30.35.001.001', '02','02');

   CreeDetail (CleDADS, 882, 'S41.G30.35.001.002',
               FloatToStr (Abs (Int (Tot1))), FloatToStr (Int (Tot1)));
   if (Tot1<0) then
      CreeDetail (CleDADS, 883, 'S41.G30.35.001.003', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G30.35.001');

   Tot3:= 0;           //Pour le nombre d'heures
   Tot4:= 0;           //Pour le montant patronal
   TCotD:= TCot.FindFirst (['PHB_SALARIE', 'PCT_MAJORATA2'], [Salarie, 'X'], TRUE);
   While ((TCotD<>nil) and
         ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) or
         (TCotD.GetValue ('PHB_DATEDEBUT')>StrToDate (StPerFin)))) do
         begin
         if (((TCotD.GetValue ('PHB_DATEFIN')>StrToDate (StPerDeb)) and
            (TCotD.GetValue ('PHB_DATEFIN')<=StrToDate (StPerFin))) or
            ((TCotD.GetValue ('PHB_DATEDEBUT')<StrToDate (StPerDeb)) and
            (TCotD.GetValue ('PHB_DATEFIN')>=StrToDate (StPerFin)))) then
            begin
            Tot3:= Tot1+TCotD.GetValue ('PHB_BASECOT');
            Tot4:= Tot2+TCotD.GetValue ('PHB_MTPATRONAL');
            end;
         TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_MAJORATA2'], [Salarie, 'X'], TRUE);
         end;
   While ((TCotD<>nil) and
         (TCotD.GetValue ('PHB_DATEDEBUT')>=StrToDate (StPerDeb)) and
         (TCotD.GetValue ('PHB_DATEDEBUT')<=StrToDate (StPerFin))) do
         begin
         Tot3:= Tot3+TCotD.GetValue ('PHB_BASECOT');
         Tot4:= Tot4+TCotD.GetValue ('PHB_MTPATRONAL');
         TCotD:= TCot.FindNext (['PHB_SALARIE', 'PCT_MAJORATA2'], [Salarie, 'X'], TRUE);
         end;
//PT134-1
   if (Tot3=0) then
      begin
{PT136
      Cumul:= GetParamSocSecur ('SO_PGCUMULHEURESSUP', '42');
}
      Cumul:= GetParamSocSecur ('SO_PGCUMULHEURESSUP', '44');
      StCumul:= 'SELECT SUM (PHC_MONTANT) AS TOTAL'+
                ' FROM HISTOCUMSAL WHERE'+
                ' PHC_SALARIE="'+Salarie+'" AND'+
                ' PHC_CUMULPAIE="'+Cumul+'" AND'+
                ' PHC_DATEFIN<="'+UsDateTime (StrToDate (StPerFin))+'" AND'+
                ' PHC_DATEDEBUT>="'+UsDateTime (StrToDate (StPerDeb))+'"';
      if (ExisteSQL (StCumul)) then
         begin
         QRechCumul:= OpenSQL (StCumul, True);
         Tot3:= QRechCumul.FindField ('TOTAL').AsFloat;
         Ferme (QRechCumul);
         end;
      end;
//FIN PT134-1
{PT135
   CreeDetail (CleDADS, 884, 'S41.G30.35.002', FloatToStr (Tot3*100),
               FloatToStr (Tot3));
}
   CreeDetail (CleDADS, 884, 'S41.G30.35.002',
               FloatToStr (Arrondi (Tot3*100, 0)),
               FloatToStr (Arrondi (Tot3, 2)));
//FIN PT135
Debug('Paie PGI/Calcul DADS-U : S41.G30.35.002');

   Tot4:= -Tot4;
   CreeDetail (CleDADS, 885, 'S41.G30.35.003.001',
               FloatToStr (Abs (Int (Tot4))), FloatToStr (Int (Tot4)));
   if (Tot4<0) then
      CreeDetail (CleDADS, 886, 'S41.G30.35.003.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G30.35.003');

   Tot2:= -Tot2;
   CreeDetail (CleDADS, 887, 'S41.G30.35.004.001',
               FloatToStr (Abs (Int (Tot2))), FloatToStr (Int (Tot2)));
   if (Tot2<0) then
      CreeDetail (CleDADS, 888, 'S41.G30.35.004.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G30.35.004');
   end;
//FIN PT125-6

Total := 0;
THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                        [Salarie, '20'], TRUE);
While ((THistoCumSalD <> nil) and
      ((THistoCumSalD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if ((THistoCumSalD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
         (THistoCumSalD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
         Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '20'], TRUE);
      end;
While ((THistoCumSalD <> nil) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (THistoCumSalD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '20'], TRUE);
      end;
{PT118
if Total >= 0 then
}
if ((Total>=0) and (Ircantec=False)) then
//FIN PT118
   begin
   CreeDetail (CleDADS, 911, 'S44.G01.00.001', '01', '01');
Debug('Paie PGI/Calcul DADS-U : S44.G01.00.001');

{PT101-10
   HDixieme := Arrondi(Total, 0)+Arrondi(Frac(Total)*100/60, 2);
   CreeDetail (CleDADS, 752, 'S44.G01.00.002', FloatToStr(HDixieme*100),
               FloatToStr(HDixieme));
}
{PT135
   CreeDetail (CleDADS, 912, 'S44.G01.00.002', FloatToStr (Total*100),
               FloatToStr (Total));
}
   CreeDetail (CleDADS, 912, 'S44.G01.00.002',
               FloatToStr (Arrondi (Total*100, 0)),
               FloatToStr (Arrondi (Total, 2)));
//FIN PT135               
//FIN PT101-10
Debug('Paie PGI/Calcul DADS-U : S44.G01.00.002');
   end;

NumOrdre := 1;
//Debut PT127
{ Recherche des contrats liés au salarié sous forme de périodes : }
TobContrats := Tob.Create('Liste des contrats de prévoyance', nil, -1);
//début PT128-3
//TobContrats.LoadDetailFromSQL('select POP_INSTITUTION, POP_REFERCONTRAT, POP_TYPECONTRAT, POP_NOMPOP, POP_DATEDEBUT, '
//                             +'POP_DATEFIN, POP_CODEPOPPREV, POP_NBENFANT, POC_PREDEFINI, POC_NODOSSIER, POC_NATURERUB, POC_RUBRIQUE '
//                             +'from contratprev left join COTISCONTRAT '
//                             +'on ((POC_INSTITUTION = POP_INSTITUTION) AND (POC_REFERCONTRAT = POP_REFERCONTRAT))');
//
TobContrats.LoadDetailFromSQL('select POP_INSTITUTION, POP_REFERCONTRAT, POP_TYPECONTRAT, POP_NOMPOP, POP_DATEDEBUT, '
                             +'POP_DATEFIN, POP_CODEPOPPREV, POP_NBENFANT '
                             +'from contratprev');
//fin PT128-3

TobPeriodesContrats := Tob.Create('Périodes des contrats de prévoyance', nil, -1);
//AddPeriode(TobPeriodesContrats, StrToDate(StPerDeb), StrToDate(StPerFin), '', '', '', '', '', 0, 0); //PT128-3 '', '', '', '',
PrevoyanceLevelFind := False;
{ Au niveau Salarie }
if not PrevoyanceLevelFind then
begin
  TobRechContrat := TobContrats.FindFirst(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['SAL', Salarie], False);
  While Assigned(TobRechContrat) do
  begin
    PrevoyanceLevelFind := True;
    { On découpe la période avec les périodes des contrats au niveau dossier }
    AddPeriode(TobPeriodesContrats, TobRechContrat); //PT143
//PT143    DecoupePeriode(TobPeriodesContrats, TobRechContrat);
    TobRechContrat := TobContrats.FindNext(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['SAL', Salarie], False);
  end;
end;
{ Au niveau Population }
if not PrevoyanceLevelFind then
begin
  PopSalarie := '';
  { On regarde si on a utiliser les populations pour paramétrer les associations
    cotisation - contrats de prévoyance }
  TobRechContrat := TobContrats.FindFirst(['POP_TYPECONTRAT'], ['POP'], False);
  if Assigned(TobRechContrat) then
  begin
    if CheckValidPopulParametres('PAI') then
    begin
      { Si oui, on vérifie la validité des populations, puis on recherche la population
        du salarie }
      Q:=OpenSql('select PNA_POPULATION from salariepopul where pna_salarie = "'+Salarie+'" and PNA_TYPEPOP = "PAI"',TRUE);
      if (not Q.EOF) then
        PopSalarie := Q.Fields[0].AsString;
      Ferme(Q);
    end else begin
      ErreurDADSU.Segment:= 'S45.G01.01.001';
      ErreurDADSU.Explication:= 'La prévoyance a été paramétrée en fonction de populations et les populations de type "PAI" ne sont pas valides.';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
    end;
  end;
  if PopSalarie <> '' then
  begin
    TobRechContrat := TobContrats.FindFirst(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['POP', PopSalarie], False);
    While Assigned(TobRechContrat) do
    begin
      PrevoyanceLevelFind := True;
      { On découpe la période avec les périodes des contrats au niveau population }
      AddPeriode(TobPeriodesContrats, TobRechContrat); //PT143
//PT143      DecoupePeriode(TobPeriodesContrats, TobRechContrat);
      TobRechContrat := TobContrats.FindNext(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['POP', PopSalarie], False);
    end;
  end;
end;
{ Au niveau Etablissement }
if not PrevoyanceLevelFind then
begin
  TobRechContrat := TobContrats.FindFirst(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['ETB', EtabSalarie], False);
  While Assigned(TobRechContrat) do
  begin
    PrevoyanceLevelFind := True;
    { On découpe la période avec les périodes des contrats au niveau dossier }
    AddPeriode(TobPeriodesContrats, TobRechContrat); //PT143
//PT143    DecoupePeriode(TobPeriodesContrats, TobRechContrat);
    TobRechContrat := TobContrats.FindNext(['POP_TYPECONTRAT', 'POP_NOMPOP'], ['ETB', EtabSalarie], False);
  end;
end;
{ Au niveau Dossier }
if not PrevoyanceLevelFind then
begin
  TobRechContrat := TobContrats.FindFirst(['POP_TYPECONTRAT'], ['DOS'], False);
  While Assigned(TobRechContrat) do
  begin
    PrevoyanceLevelFind := True;
    { On découpe la période avec les périodes des contrats au niveau dossier }
    AddPeriode(TobPeriodesContrats, TobRechContrat); //PT143
//PT143    DecoupePeriode(TobPeriodesContrats, TobRechContrat);
    TobRechContrat := TobContrats.FindNext(['POP_TYPECONTRAT'], ['DOS'], False);
  end;
end;
//Début PT143
{ Pour chaque période, on cherche le code évènement associé }
For IndexPeriode := 0 to TobPeriodesContrats.Detail.Count -1 do
begin
  TempTobPeriodesContrats := TobPeriodesContrats.Detail[IndexPeriode];
  { Si la date de début du contrat est antérieur à la date de début renseignée, c'est
    qu'on a changer de période DADS-U depuis le début du contrat. Dans ce cas, on est
    en continuité, code 90 }
  if TempTobPeriodesContrats.GetDateTime('DATEDEB_CONTRAT') < TempTobPeriodesContrats.GetDateTime('DATEDEB') then
    stCode := '90'
  else begin
    { On regarde si on a un contrat pour la même institution qui se termine la
      veille ou le jour du début du contrat }
    if Assigned(TobPeriodesContrats.FindFirst(['INSTITUTION',
                                               'DATEFIN_CONTRAT'],
                                              [TempTobPeriodesContrats.GetString('INSTITUTION'),
                                               TempTobPeriodesContrats.GetDateTime('DATEDEB_CONTRAT')],
                                              False))
      or
       Assigned(TobPeriodesContrats.FindFirst(['INSTITUTION',
                                               'DATEFIN_CONTRAT'],
                                              [TempTobPeriodesContrats.GetString('INSTITUTION'),
                                               TempTobPeriodesContrats.GetDateTime('DATEDEB_CONTRAT')-1],
                                              False)) then
      stCode := '50' { C'est un "changement des conditions contractuelles" code 50 }

    else
      stCode := '01';{ C'est une "affiliation" code 01 }
  end;
  TempTobPeriodesContrats.PutValue('CODEEVT', stCode);
end;
{ On supprime les lignes qui ne sont pas dans la période } 
For IndexPeriode := 0 to TobPeriodesContrats.Detail.Count -1 do
begin
  TempTobPeriodesContrats := TobPeriodesContrats.Detail[IndexPeriode];
  if not (    (StrToDate(StPerDeb) <= TempTobPeriodesContrats.GetDateTime('DATEDEB'))
          and (TempTobPeriodesContrats.GetDateTime('DATEDEB') <= StrToDate(StPerFin))
          and (StrToDate(StPerDeb) <= TempTobPeriodesContrats.GetDateTime('DATEFIN'))
          and (TempTobPeriodesContrats.GetDateTime('DATEFIN') <= StrToDate(StPerFin))
         ) then
  begin
    TempTobPeriodesContrats.ChangeParent(nil, -1);
    FreeAndNil(TempTobPeriodesContrats);
  end;
end;
(*  if IndexPeriode = 0 then
  begin
    try
      isAffiliation := (TobPeriodesContrats.Detail[IndexPeriode].GetDateTime('DATEDEB') = TSalD.GetDateTime('PSA_DATEENTREE'));
    except
      isAffiliation := False;
    end;
    { si le début de la période est égale à la date d'entrée dans l'entreprise,
      c'est qu'on est en affiliation code 01 }
    if isAffiliation then
    begin
      TobPeriodesContrats.Detail[IndexPeriode].PutValue('CODEEVT', '01')
    end else { Sinon, la première période est en code 90 }
      TobPeriodesContrats.Detail[IndexPeriode].PutValue('CODEEVT', '90')
  end else begin { et pour les suivantes : }
    { Si la période précédente est vide (pas de référence contrat) c'est qu'on
      est en affiliation code 01}
    if TobPeriodesContrats.Detail[IndexPeriode -1].GetString('REFERCONTRAT') = '' then
      TobPeriodesContrats.Detail[IndexPeriode].PutValue('CODEEVT', '01')
    { Si la période précédente est renseignée, on est en "changement des
      conditions contractuelles" code 50 }
    else
      TobPeriodesContrats.Detail[IndexPeriode].PutValue('CODEEVT', '50');
  end;
end;
*)
//Fin PT143
//{ On rassemble les contrats identiques coupés en plusieurs périodes }
//Rassemble(TobPeriodesContrats);
{ Pour chaque contrat trouvé, on cherche la cotisation associée et son montant }
//début PT128-3
FreeAndNil(TobContrats);
TobContrats := Tob.Create('Liste des contrats de prévoyance', nil, -1);
TobContrats.LoadDetailFromSQL('select POP_INSTITUTION, POP_REFERCONTRAT, POP_TYPECONTRAT, POP_NOMPOP, POP_DATEDEBUT, '
                             +'POP_DATEFIN, POP_CODEPOPPREV, POP_NBENFANT, POC_PREDEFINI, POC_NODOSSIER, POC_NATURERUB, POC_RUBRIQUE '
                             +'from contratprev left join COTISCONTRAT '
                             +'on ((POC_INSTITUTION = POP_INSTITUTION) AND (POC_REFERCONTRAT = POP_REFERCONTRAT))');
For IndexPeriode := TobPeriodesContrats.Detail.Count -1 downto 0 do
begin

  TobRechContrat := TobContrats.FindFirst(['POP_INSTITUTION', 'POP_REFERCONTRAT']
                                         , [  TobPeriodesContrats.Detail[IndexPeriode].getString('INSTITUTION')
                                            , TobPeriodesContrats.Detail[IndexPeriode].getString('REFERCONTRAT')]
                                         , False);
  While Assigned(TobRechContrat) do
  begin
//    TobPeriodesContrats.Detail[IndexPeriode].PutValue('MONTANT',
//      TobPeriodesContrats.Detail[IndexPeriode].GetDouble('MONTANT') +
//      SommeRubSalarie(TobPeriodesContrats.Detail[IndexPeriode].GetString('NATURERUB'),
//                      TobPeriodesContrats.Detail[IndexPeriode].GetString('RUBRIQUE'),
//                      Salarie,
//                      TobPeriodesContrats.Detail[IndexPeriode].GetDateTime('DATEDEB'),
//                      TobPeriodesContrats.Detail[IndexPeriode].GetDateTime('DATEFIN'),
//                      DebExer, FinExer));
    TobPeriodesContrats.Detail[IndexPeriode].PutValue('MONTANT',
      TobPeriodesContrats.Detail[IndexPeriode].GetDouble('MONTANT') +
      SommeRubSalarie(TobRechContrat.GetString('POC_NATURERUB'),
                      TobRechContrat.GetString('POC_RUBRIQUE'),
                      Salarie,
                      TobPeriodesContrats.Detail[IndexPeriode].GetDateTime('DATEDEB'),
                      TobPeriodesContrats.Detail[IndexPeriode].GetDateTime('DATEFIN'),
                      DebExer, FinExer,
                      True)); //PT131

    TobRechContrat := TobContrats.FindNext(['POP_INSTITUTION', 'POP_REFERCONTRAT']
                                          , [  TobPeriodesContrats.Detail[IndexPeriode].getString('INSTITUTION')
                                             , TobPeriodesContrats.Detail[IndexPeriode].getString('REFERCONTRAT')]
                                          , False);
  end;
end;
//Fin PT128-3
{ On supprime les périodes vides }
For IndexPeriode := TobPeriodesContrats.Detail.Count -1 downto 0 do
begin
  if TobPeriodesContrats.Detail[IndexPeriode].GetString('REFERCONTRAT') = '' then
  begin
    TempTob := TobPeriodesContrats.Detail[IndexPeriode];
    TempTob.ChangeParent(nil, -1);
    FreeAndNil(TempTob);
  end;
end;
{ Ecriture en base de la S45.G01 }
if TobPeriodesContrats.Detail.Count > 0 then
begin
  ForceNumerique(DateToStr(TSalD.GetValue('PSA_DATEENTREE')), Buf);
  CreeDetail (CleDADS, 921, 'S45.G01.00.001', Buf, DateToStr(TSalD.GetValue('PSA_DATEENTREE')));
Debug('Paie PGI/Calcul DADS-U : S45.G01.00.001');
end;
For IndexPeriode := 0 to TobPeriodesContrats.Detail.Count -1 do
begin
  CreeDetail (CleDADS, 921+NumOrdre,    'S45.G01.01.001',
              TobPeriodesContrats.Detail[IndexPeriode].GetString('CODEEVT'),
              TobPeriodesContrats.Detail[IndexPeriode].GetString('CODEEVT'));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.001');
  //deb PT128-1
  Buf := '';
  if TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEDEB') <> null then
  begin
    ForceNumerique(DateToStr(TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEDEB')), Buf);
    CreeDetail (CleDADS, 921+NumOrdre+1, 'S45.G01.01.002', Buf,
                DateToStr(TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEDEB')));
  end;
  if (Length(Buf) <> 8) then
  begin
    ErreurDADSU.Segment:= 'S45.G01.01.002';
    ErreurDADSU.Explication:= 'La date de début de période de cotisation est erronnée';
    ErreurDADSU.CtrlBloquant:= True;
    EcrireErreur(ErreurDADSU);
  end;
  //fin PT128-1
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.002');
  //deb PT128-1
  Buf := '';
  if TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEFIN') <> null then
  begin
    ForceNumerique(DateToStr(TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEFIN')), Buf);
    CreeDetail (CleDADS, 921+NumOrdre+2, 'S45.G01.01.003', Buf,
                DateToStr(TobPeriodesContrats.Detail[IndexPeriode].GetValue('DATEFIN')));
  end;
  if (Length(Buf) <> 8) then
  begin
    ErreurDADSU.Segment:= 'S45.G01.01.003';
    ErreurDADSU.Explication:= 'La date de fin de période de cotisation est erronnée';
    ErreurDADSU.CtrlBloquant:= True;
    EcrireErreur(ErreurDADSU);
  end;
  //fin PT128-1
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.003');
  CreeDetail (CleDADS, 921+NumOrdre+3,  'S45.G01.01.004.001',
              'P'+TobPeriodesContrats.Detail[IndexPeriode].GetString('INSTITUTION'),
              TobPeriodesContrats.Detail[IndexPeriode].GetString('INSTITUTION'));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.004.001');
{ Code organisme gestionnaire
  CreeDetail (CleDADS, 921+NumOrdre+4,  'S45.G01.01.004.002',
              '',
              '');   }
  CreeDetail (CleDADS, 921+NumOrdre+5,  'S45.G01.01.005',
              TobPeriodesContrats.Detail[IndexPeriode].GetString('REFERCONTRAT'),
              TobPeriodesContrats.Detail[IndexPeriode].GetString('REFERCONTRAT'));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.005');

  Total := TobPeriodesContrats.Detail[IndexPeriode].Getdouble('MONTANT');
  //Début PT128-2
  Total := Arrondi(Total, 0);
  if Total = 0 then
  begin
    ErreurDADSU.Segment:= 'S45.G01.01.007.001';
    ErreurDADSU.Explication:= 'La base de prévoyance contractuelle a été calculé à 0';
    ErreurDADSU.CtrlBloquant:= False;
    EcrireErreur(ErreurDADSU);
  end;
  //fin PT128-2  
  if (BaseSS <> Total) then  //PT128-2    (Total <> 0) and (BaseSS <> 0) and
  begin
    CreeDetail (CleDADS, 921+NumOrdre+6,  'S45.G01.01.006', '02', '02');
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.006');
    CreeDetail (CleDADS, 921+NumOrdre+7,  'S45.G01.01.007.001', FloatToStr(Abs(Total)), FloatToStr(Total));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.007.001');
    if Total < 0 then
    begin
      CreeDetail (CleDADS, 921+NumOrdre+8,  'S45.G01.01.007.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.007.002');
    end;
  end;
  { Population Couverte }
  CreeDetail (CleDADS, 921+NumOrdre+9,  'S45.G01.01.008',
              TobPeriodesContrats.Detail[IndexPeriode].GetString('CODEPOPPREV'),
              TobPeriodesContrats.Detail[IndexPeriode].GetString('CODEPOPPREV'));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.008');
  { Situation familiale }
  if ((TSalD.GetValue('PSA_SITUATIONFAMIL') <> '') and
    (TSalD.GetValue('PSA_SITUATIONFAMIL') <> null)) then
    CreeDetail (CleDADS, 921+NumOrdre+10, 'S45.G01.01.009',
                RechDom('PGSITUATIONFAMIL', TSalD.GetValue('PSA_SITUATIONFAMIL'), TRUE),
                TSalD.GetValue('PSA_SITUATIONFAMIL'))
  else
    CreeDetail (CleDADS, 921+NumOrdre+10, 'S45.G01.01.009', '90', '90');
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.009');
  { Nombre d'enfants couverts }
  NbEnfants := TobPeriodesContrats.Detail[IndexPeriode].GetInteger('NBENFANT');
  if NbEnfants > 91 then NbEnfants := 90;
  CreeDetail (CleDADS, 921+NumOrdre+11, 'S45.G01.01.010', IntToStr(NbEnfants), IntToStr(NbEnfants));
Debug('Paie PGI/Calcul DADS-U : S45.G01.01.010');
  NumOrdre := NumOrdre + 12;
end;     
FreeAndNil(TobContrats);
FreeAndNil(TobPeriodesContrats);
//Fin PT127
if NbreHistoOrg > 0 then
   FreeAndNil(TInstitution);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 20/06/2002
Modifié le ... :   /  /
Description .. : Création des enregistrements concernant les sommes isolées
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure GereSomIso(T : TOB;Salarie, TypeE : string; Periode : integer;
          var NumOrdre : integer);
var
Total : integer;
TD : TOB;
AnneeA, JourJ, MoisM : Word;
CleDADS : TCleDADS;
begin
//PT101-6
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= StrToDate (StPerDeb);
CleDADS.DateFin:= StrToDate (StPerFin);
CleDADS.Exercice:= PGExercice;
//FIN PT101-6
Total := 0;
TD := T.FindFirst(['PHB_SALARIE', 'PRM_SOMISOL'], [Salarie, TypeE], TRUE);
While ((TD <> nil) and
      ((TD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      TD := T.FindNext(['PHB_SALARIE', 'PRM_SOMISOL'], [Salarie, TypeE], TRUE);
While ((TD <> nil) and
      (TD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
{PT91
      Total := Total + TD.GetValue('PHB_MTREM');
}
      if (TD.GetValue('PHB_SENSBUL')<>'M') then
         Total := Total + TD.GetValue('PHB_MTREM')
      else
         Total := Total - TD.GetValue('PHB_MTREM');
//FIN PT91
      TD := T.FindNext(['PHB_SALARIE', 'PRM_SOMISOL'], [Salarie, TypeE], TRUE);
      end;
if (Total <> 0) then
   begin
   NumOrdre := NumOrdre+1;
   CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.001', TypeE, TypeE);
Debug('Paie PGI/Calcul DADS-U : S41.G01.04.001 '+IntToStr(NumOrdre));
   NumOrdre := NumOrdre+1;
   DecodeDate(FinExer, AnneeA, MoisM, JourJ);
   CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.002', IntToStr(AnneeA),
              IntToStr(AnneeA));
Debug('Paie PGI/Calcul DADS-U : S41.G01.04.002 '+IntToStr(NumOrdre));
   NumOrdre := NumOrdre+1;
   CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.003.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   NumOrdre := NumOrdre+1;
   if (Total < 0) then
      CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.003.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.04.003 '+IntToStr(NumOrdre));
   NumOrdre := NumOrdre+1;
   CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.004.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   NumOrdre := NumOrdre+1;
   if (Total < 0) then
      CreeDetail (CleDADS, NumOrdre, 'S41.G01.04.004.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.04.004 '+IntToStr(NumOrdre));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/11/2001
Modifié le ... :   /  /
Description .. : Création des enregistrements concernant les exonérations
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure GereExo(T : TOB;Champ, Salarie, TypeE : string;
          Periode : integer; var NumOrdre : integer; BufContrat : string='');
var
Total1, Total2 : Double;
TBaseD, TD : TOB;
CleDADS : TCleDADS;
ErreurDADSU : TControle;
begin
//PT101-6
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= StrToDate(StPerDeb);
CleDADS.DateFin:= StrToDate(StPerFin);
CleDADS.Exercice:= PGExercice;

ErreurDADSU.Salarie:= Salarie;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= Periode;
ErreurDADSU.DateDeb:= StrToDate(StPerDeb);
ErreurDADSU.DateFin:= StrToDate(StPerFin);
ErreurDADSU.Exercice:= PGExercice;
//FIN PT101-6

Total1 := 0;
Total2 := 0;
if Champ = 'PCT_DADSEXOBASE' then
   begin
   TD := T.FindFirst(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
   While ((TD <> nil) and
         ((TD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
         TD := T.FindNext(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
   While ((TD <> nil) and
         (TD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (TD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         Total1 := Total1 + TD.GetValue('PHB_BASECOT');
         Total2 := Total2 + TD.GetValue('PHB_TRANCHE1');
         TD := T.FindNext(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
         end;
   end
else
//PT92
if ((Champ = 'PHC_CUMULPAIE') and (BufContrat='CAA')) then
   begin
   TD := T.FindFirst(['PHC_SALARIE', Champ], [Salarie, '09'], TRUE);
   While ((TD <> nil) and
         ((TD.GetValue('PHC_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TD.GetValue('PHC_DATEDEBUT')>StrToDate(StPerFin)))) do
         begin
         if ((TD.GetValue('PHC_DATEFIN')>StrToDate(StPerDeb)) and
            (TD.GetValue('PHC_DATEFIN')<=StrToDate(StPerFin))) then
            Total1 := Total1 + TD.GetValue('PHC_MONTANT');
         TD := T.FindNext(['PHC_SALARIE', Champ], [Salarie, '09'], TRUE);
         end;
   While ((TD <> nil) and
         (TD.GetValue('PHC_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (TD.GetValue('PHC_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         Total1 := Total1 + TD.GetValue('PHC_MONTANT');
         TD := T.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'], [Salarie, '09'], TRUE);
         end;
   Total2:= Total1;
   end
else
//FIN PT92
   begin
   TD := T.FindFirst(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
   While ((TD <> nil) and
         ((TD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
         (TD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
         TD := T.FindNext(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
   While ((TD <> nil) and
         (TD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
         (TD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
         begin
         TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_BRUTSS'],
                                   [Salarie, 'X'], TRUE);
         While ((TBaseD <> nil) and
               ((TBaseD.GetValue('PHB_DATEDEBUT')<TD.GetValue('PHB_DATEDEBUT')) or
               (TBaseD.GetValue('PHB_DATEDEBUT')>TD.GetValue('PHB_DATEFIN')))) do
               TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BRUTSS'],
                                        [Salarie, 'X'], TRUE);
         While ((TBaseD <> nil) and
               (TBaseD.GetValue('PHB_DATEDEBUT')>=TD.GetValue('PHB_DATEDEBUT')) and
               (TBaseD.GetValue('PHB_DATEDEBUT')<=TD.GetValue('PHB_DATEFIN'))) do
               begin
               if (TD.GetValue('PHB_MTPATRONAL')<>0) then                      //PT141
                  Total1:= Total1+TBaseD.GetValue ('PHB_BASECOT');

               TBaseD:= TBase.FindNext (['PHB_SALARIE', 'PCT_BRUTSS'],
                                        [Salarie, 'X'], TRUE);
               end;

         TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_PLAFONDSS'],
                                   [Salarie, 'X'], TRUE);
         While ((TBaseD <> nil) and
               ((TBaseD.GetValue('PHB_DATEDEBUT')<TD.GetValue('PHB_DATEDEBUT')) or
               (TBaseD.GetValue('PHB_DATEDEBUT')>TD.GetValue('PHB_DATEFIN')))) do
               TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_PLAFONDSS'],
                                        [Salarie, 'X'], TRUE);
         While ((TBaseD <> nil) and
               (TBaseD.GetValue('PHB_DATEDEBUT')>=TD.GetValue('PHB_DATEDEBUT')) and
               (TBaseD.GetValue('PHB_DATEDEBUT')<=TD.GetValue('PHB_DATEFIN'))) do
               begin
               if (TD.GetValue('PHB_MTPATRONAL')<>0) then                      //PT141
                  Total2:= Total2+TBaseD.GetValue ('PHB_TRANCHE1');

               TBaseD:= TBase.FindNext (['PHB_SALARIE', 'PCT_PLAFONDSS'],
                                        [Salarie, 'X'], TRUE);
               end;
         TD := T.FindNext(['PHB_SALARIE', Champ], [Salarie, TypeE], TRUE);
         end;
   end;

{PT90-1
if ((Total1 > 0) or (Total2 > 0) or ((BufContrat='CAA') and
   (TypeE = '01'))) then
}
if ((Total1 > 0) or (Total2 > 0)) then
   begin
//PT97-2
{PT101-2
   if ((BufContrat<>'CAP') and (TypeE='02')) then
      EcrireErreur ('Erreur:   Saisie complémentaire Salarié : Le'+
                    ' salarié comporte une exonération ''apprenti loi de 1987'''+
                    ' mais son contrat n''est pas ''apprenti +10''', Erreur);
}
//FIN PT97-2
   NumOrdre := NumOrdre+1;
   CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.001', TypeE, TypeE);
Debug('Paie PGI/Calcul DADS-U : S41.G01.06.001 '+IntToStr(NumOrdre));

   NumOrdre := NumOrdre+1;
{PT90-1
   if (BufContrat='CAA') then
      Total1 := 0;
}

   CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.002.001',
               FloatToStr(Abs(Arrondi(Total1, 0))),
               FloatToStr(Arrondi(Total1, 0)));
   NumOrdre := NumOrdre+1;
   if (Total1 < 0) then
      CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.002.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S41.G01.06.002 '+IntToStr(NumOrdre));

{PT90-1
   if (BufContrat<>'CAA') then
      begin
}
   NumOrdre := NumOrdre+1;
   if ((BufContrat = 'CAP') and (TypeE = '02')) then
      begin
      CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.003.001',
                  FloatToStr(Abs(Arrondi(Total1, 0))),
                  FloatToStr(Arrondi(Total1, 0)));
      NumOrdre := NumOrdre+1;
      if (Total1 < 0) then
         CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.003.002', 'N', 'N');
      end
   else
      begin
      CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.003.001',
                  FloatToStr(Abs(Arrondi(Total2, 0))),
                  FloatToStr(Arrondi(Total2, 0)));
      NumOrdre := NumOrdre+1;
      if (Total2 < 0) then
         CreeDetail (CleDADS, NumOrdre, 'S41.G01.06.003.002', 'N', 'N');
      end;
//FIN PT90-1      
Debug('Paie PGI/Calcul DADS-U : S41.G01.06.003 '+IntToStr(NumOrdre));
   end;

//PT125-6
//FIN PT125-6   
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/06/2004
Modifié le ... :   /  /
Description .. : Pour une période donnée, remplissage des éléments
Suite ........ : prud'hommaux
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure RemplitPrudh(Salarie : string; Periode : integer);
var
NumOrdre : integer;
TEtabD, TSalD : TOB;
CleDADS : TCleDADS;
ErreurDADSU : TControle;
DateBuf : TDateTime;
BufCode, BufLib : string;
begin
TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);

//PT101-6
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= FinExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;

ErreurDADSU.Salarie:= Salarie;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= Periode;
ErreurDADSU.DateDeb:= FinExer;
ErreurDADSU.DateFin:= FinExer;
ErreurDADSU.Exercice:= PGExercice;
//FIN PT101-6

{PT101-7
if (TypeD='001') then
}
if ((TypeD='001') or (TypeD='201')) then
//FIN PT101-7
   begin
{PT125-6
   NumOrdre:= 688;
}
   NumOrdre:= 701;
//FIN PT125-6
{PT96
   if ((TSalD.GetValue('PSA_DATESORTIE')>FinAnnee(FinExer)) or
      (TSalD.GetValue('PSA_DATESORTIE')=IDate1900) or
      (TSalD.GetValue('PSA_DATESORTIE')=null)) then
}
{PT101-2
   if ((TSalD.GetValue('PSA_DATESORTIE')>FinAnnee(FinExer)) or
      (TSalD.GetValue('PSA_DATESORTIE')<=IDate1900) or
      (TSalD.GetValue('PSA_DATESORTIE')=null)) then
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.008', '01', '01')
   else
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.008', '02', '02');

   NumOrdre := 689;
   if ((TSalD.GetValue('PSA_PRUDHCOLL') <> '') and
      (TSalD.GetValue('PSA_PRUDHCOLL') <> null)) then
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.009', '0'+TSalD.GetValue('PSA_PRUDHCOLL'),
                 TSalD.GetValue('PSA_PRUDHCOLL'))
   else
      begin
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.009', '01', '1');
      EcrireErreur ('          Saisie complémentaire Salarié :'+
                    ' Le collège prud''homal a été forcé à ''Salarié''',
                    Erreur);
      end;

   NumOrdre := 690;
   if ((TSalD.GetValue('PSA_PRUDHSECT') <> '') and
      (TSalD.GetValue('PSA_PRUDHSECT') <> null)) then
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.010', '0'+TSalD.GetValue('PSA_PRUDHSECT'),
                 TSalD.GetValue('PSA_PRUDHSECT'))
   else
      begin
      CreeDetail (Salarie, TypeD, Periode, FinExer, FinExer, NumOrdre,
                 'S41.G02.00.010', '04', '4');
      EcrireErreur ('          Saisie complémentaire Salarié :'+
                    ' La section prud''homale a été forcé à'+
                    ' ''Activités diverses''',
                    Erreur);
      end;
Debug('Paie PGI/Calcul DADS-U : S41.G02.00.010');
}
{PT125-6
   DateBuf:= PremierJourSemaine (NumSemaine (FinAnnee (FinExer)),
                                 StrToInt (PGExercice))+4;
   if ((StrToDate (StPerDeb)<DateBuf) and (StrToDate (StPerFin)>DateBuf)) then
}
   DateBuf:= DayOfTheWeek (FinAnnee (FinExer));
   if (DateBuf<DayFriday) then
      DateBuf:= FinAnnee (FinExer)-(DateBuf+2)
   else
      DateBuf:= FinAnnee (FinExer)-(DateBuf-5);
{PT142
   if (((StrToDate (StPerDeb)<=DateBuf) and (StrToDate (StPerFin)>DateBuf)) or
      ((StrToDate (StPerFin)=FinExer) and (FinExer<>FinAnnee (FinExer)) and
      ((TSalD.GetValue('PSA_DATESORTIE')>DateBuf) or
      (TSalD.GetValue('PSA_DATESORTIE')<=IDate1900) or
      (TSalD.GetValue('PSA_DATESORTIE')=null)))) then
}
   if (((StrToDate (StDateDebut)<=DateBuf) and (StrToDate (StPerFin)>DateBuf)) or
      ((StrToDate (StPerFin)=FinExer) and (FinExer<>FinAnnee (FinExer)) and
      ((TSalD.GetValue ('PSA_DATESORTIE')>DateBuf) or
      (TSalD.GetValue ('PSA_DATESORTIE')<=IDate1900) or
      (TSalD.GetValue ('PSA_DATESORTIE')=null)))) then
//FIN PT142
//FIN PT125-6
      begin
      CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.008', '01', '01');
{PT125-6
      NumOrdre:= 689;
}
      NumOrdre:= 702;
//FIN PT125-6
      if ((TSalD.GetValue ('PSA_PRUDHCOLL') <> '') and
         (TSalD.GetValue ('PSA_PRUDHCOLL') <> null)) then
         CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.009',
                     '0'+TSalD.GetValue ('PSA_PRUDHCOLL'),
                     TSalD.GetValue ('PSA_PRUDHCOLL'))
      else
         begin
         CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.009', '01', '1');
         ErreurDADSU.Segment:= 'S41.G02.00.009';
         ErreurDADSU.Explication:= 'Le collège prud''homal a été forcé à'+
                                   ' ''Salarié''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur (ErreurDADSU);
         end;
Debug('Paie PGI/Calcul DADS-U : S41.G02.00.009');

{PT125-6
      NumOrdre:= 690;
}
      NumOrdre:= 703;

      TEtabD:= TEtab.FindFirst (['ET_ETABLISSEMENT'],
                                [TSalD.GetValue ('PSA_ETABLISSEMENT')], TRUE);
//FIN PT125-6
      if ((TSalD.GetValue ('PSA_PRUDHSECT')<>'') and
         (TSalD.GetValue ('PSA_PRUDHSECT') <> null)) then
         begin
         BufCode:= TSalD.GetValue ('PSA_PRUDHSECT'); //PT125-6
         CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.010', '0'+BufCode, BufCode);
//PT125-6
{PT139
         if ('0'+BufCode<>TEtabD.GetValue ('ETB_PRUDH')) then
}
         if (('0'+BufCode<>TEtabD.GetValue ('ETB_PRUDH')) and
            (BufCode<>'5')) then
//FIN PT139
            begin
            ErreurDADSU.Segment:= 'S41.G02.00.010';
            ErreurDADSU.Explication:= 'La section prud''homale du salarié ne'+
                                      ' correspond pas à la section'+
                                      ' prud''homale de l''établissement';
            ErreurDADSU.CtrlBloquant:= True;
            EcrireErreur (ErreurDADSU);
            end;
//FIN PT125-6
         end
      else
         begin
{PT125-6
         CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.010', '04', '4');
         ErreurDADSU.Segment:= 'S41.G02.00.010';
         ErreurDADSU.Explication:= 'La section prud''homale a été forcée à'+
                                   ' ''Activités diverses''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur (ErreurDADSU);
}
         if (TEtabD.GetValue ('ETB_PRUDH')<>'') then
            BufCode:= Copy (TEtabD.GetValue ('ETB_PRUDH'), 2, 1)
         else
            BufCode:= '4';
         BufLib:= RechDom ('PGSECTIONPRUD', BufCode, False);
         CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.010', '0'+BufCode, BufCode);
         ErreurDADSU.Segment:= 'S41.G02.00.010';
         ErreurDADSU.Explication:= 'La section prud''homale a été forcée à'+
                                   ' '''+BufLib+'''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur (ErreurDADSU);
//FIN PT125-6
         end;
Debug('Paie PGI/Calcul DADS-U : S41.G02.00.010');
      end
   else
      CreeDetail (CleDADS, NumOrdre, 'S41.G02.00.008', '02', '02');
Debug('Paie PGI/Calcul DADS-U : S41.G02.00.008');
//FIN PT101-2
   end;
end;


//PT101-11
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
Function RemplitFromPublicotis (Salarie, Num : string; PerDeb, PerFin : TDateTime) : double;
var
Total : double;
DI : DIrcantec;
Predefini, StNature, StSQL : string;
QRechNature, QRechSQL : TQuery;
THisto, THistoD, TNature, TNatureD : TOB;
begin
result:= 0;
Total:= 0;
StNature:= 'SELECT PUO_NATURERUB, PUO_UTILSEGMENT, PUO_PREDEFINI'+
           ' FROM PUBLICOTIS WHERE'+
           ' ##PUO_PREDEFINI## PUO_UTILSEGMENT="'+Num+'"'+
           ' GROUP BY PUO_NATURERUB, PUO_UTILSEGMENT, PUO_PREDEFINI';
QRechNature:= OpenSql (StNature, TRUE);
TNature:= TOB.Create ('Les NatureRub', NIL, -1);
TNature.LoadDetailDB ('MA_NATURE', '', '', QRechNature, False);
Ferme(QRechNature);

{PT129-1
TNatureD:= TNature.FindFirst ([''], [''], True);
}
TNatureD:= TNature.FindFirst (['PUO_PREDEFINI'], ['DOS'], True);
if (TNatureD<>nil) then
   Predefini:= 'DOS'
else
   begin
   TNatureD:= TNature.FindFirst (['PUO_PREDEFINI'], ['STD'], True);
   if (TNatureD<>nil) then
      Predefini:= 'STD'
   else
      begin
      TNatureD:= TNature.FindFirst (['PUO_PREDEFINI'], ['CEG'], True);
      if (TNatureD<>nil) then
         Predefini:= 'CEG'
      else
         begin
         FreeAndNil (TNature);
         exit;
         end;
      end;
   end;

TNatureD:= TNature.FindFirst (['PUO_PREDEFINI'], [Predefini], True);
//FIN PT129-1
While (TNatureD<>nil) do
      begin
      if (TNatureD.GetValue ('PUO_NATURERUB')<>'CUM') then
         begin
         DI.ChampsHisto:= ' PHB_SALARIE, PHB_DATEDEBUT, PHB_DATEFIN,'+
                          ' PHB_BASEREM, PHB_TAUXREM, PHB_COEFFREM, PHB_MTREM,'+
                          ' PHB_BASECOT, PHB_TAUXSALARIAL, PHB_TAUXPATRONAL,'+
                          ' PHB_MTSALARIAL, PHB_MTPATRONAL, PHB_SENSBUL';
         DI.TableHisto:= 'HISTOBULLETIN';
         DI.ChampJoin:= 'PHB_RUBRIQUE';
         DI.Prefixe:= 'PHB';
         end
      else
         begin
         DI.ChampsHisto:= ' PHC_SALARIE, PHC_DATEDEBUT, PHC_DATEFIN,'+
                          ' PHC_MONTANT';
         DI.TableHisto:= 'HISTOCUMSAL';
         DI.ChampJoin:= 'PHC_CUMULPAIE';
         DI.Prefixe:= 'PHC';
         end;
      StSQL:= 'SELECT PUO_NATURERUB, PUO_RUBRIQUE, PUO_UTILSEGMENT,'+
              ' PUO_CHAMPRETENU,'+DI.ChampsHisto+
              ' FROM PUBLICOTIS'+
              ' LEFT JOIN '+DI.TableHisto+' ON'+
              ' PUO_RUBRIQUE='+DI.ChampJoin+' WHERE'+
              ' ##PUO_PREDEFINI## PUO_UTILSEGMENT="'+Num+'" AND'+
              ' PUO_PREDEFINI="'+Predefini+'" AND'+
              ' '+DI.Prefixe+'_SALARIE="'+Salarie+'" AND'+
              ' '+DI.Prefixe+'_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
              ' '+DI.Prefixe+'_DATEDEBUT>="'+UsDateTime(DebExer)+'"'+
              ' ORDER BY '+DI.Prefixe+'_DATEDEBUT';
      QRechSQL:= OpenSql (StSQL, TRUE);
      THisto:= TOB.Create ('Les historiques', NIL, -1);
      THisto.LoadDetailDB ('HISTOPUBLIC','','',QRechSQL,False);
      Ferme (QRechSQL);

      THistoD:= THisto.FindFirst ([''], [''], TRUE);
      While ((THistoD <> nil) and
            ((THistoD.GetValue (DI.Prefixe+'_DATEDEBUT')<PerDeb) or
            (THistoD.GetValue (DI.Prefixe+'_DATEDEBUT')>PerFin))) do
            begin
            if (((THistoD.GetValue (DI.Prefixe+'_DATEFIN')>PerDeb) and
               (THistoD.GetValue (DI.Prefixe+'_DATEFIN')<=PerFin)) or
               ((THistoD.GetValue (DI.Prefixe+'_DATEDEBUT')<PerDeb) and
               (THistoD.GetValue (DI.Prefixe+'_DATEFIN')>=PerFin))) then
               begin
               if (THistoD.GetValue ('PUO_NATURERUB')='REM') then
                  begin
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='05') then
                     begin
                     if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                        Total:= Total+THistoD.GetValue ('PHB_BASEREM')
                     else
                        Total:= Total-THistoD.GetValue ('PHB_BASEREM');
                     end
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='07') then
                     Total:= Total+THistoD.GetValue ('PHB_COEFFREM')
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='08') then
                     begin
                     if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                        Total:= Total+THistoD.GetValue ('PHB_MTREM')
                     else
                        Total:= Total-THistoD.GetValue ('PHB_MTREM');
                     end
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='06') then
                     Total:= Total+THistoD.GetValue ('PHB_TAUXREM');
                  end
               else
               if (THistoD.GetValue ('PUO_NATURERUB')='COT') then
                  begin
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='12') then
                     begin
                     if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                        Total:= Total+THistoD.GetValue ('PHB_BASECOT')
                     else
                        Total:= Total-THistoD.GetValue ('PHB_BASECOT');
                     end
//PT125-6
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='13') then
                     begin
                     if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                        Total:= Total+THistoD.GetValue ('PHB_MTPATRONAL')
                     else
                        Total:= Total-THistoD.GetValue ('PHB_MTPATRONAL');
                     end
//FIN PT125-6
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='14') then
                     begin
                     if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                        Total:= Total+THistoD.GetValue ('PHB_MTSALARIAL')
                     else
                        Total:= Total-THistoD.GetValue ('PHB_MTSALARIAL');
                     end
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='25') then
                     Total:= Total+THistoD.GetValue ('PHB_TAUXSALARIAL')
//PT125-6
                  else
                  if (THistoD.GetValue ('PUO_CHAMPRETENU')='26') then
                     Total:= Total+THistoD.GetValue ('PHB_TAUXPATRONAL');
//FIN PT125-6
                  end
               else
               if (THistoD.GetValue ('PUO_NATURERUB')='CUM') then
                  Total:= Total+THistoD.GetValue('PHC_MONTANT');
               end;
            THistoD:= THisto.FindNext ([''], [''], TRUE);
            end;

      While ((THistoD <> nil) and
            (THistoD.GetValue (DI.Prefixe+'_DATEDEBUT')>=PerDeb) and
            (THistoD.GetValue (DI.Prefixe+'_DATEDEBUT')<=PerFin)) do
            begin
            if (THistoD.GetValue ('PUO_NATURERUB')='REM') then
               begin
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='05') then
                  begin
                  if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                     Total:= Total+THistoD.GetValue ('PHB_BASEREM')
                  else
                     Total:= Total-THistoD.GetValue ('PHB_BASEREM');
                  end
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='07') then
                  Total:= Total+THistoD.GetValue ('PHB_COEFFREM')
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='08') then
                  begin
                  if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                     Total:= Total+THistoD.GetValue ('PHB_MTREM')
                  else
                     Total:= Total-THistoD.GetValue ('PHB_MTREM');
                  end
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='06') then
                  Total:= Total+THistoD.GetValue ('PHB_TAUXREM');
               end
            else
            if (THistoD.GetValue ('PUO_NATURERUB')='COT') then
               begin
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='12') then
                  begin
                  if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                     Total:= Total+THistoD.GetValue ('PHB_BASECOT')
                  else
                     Total:= Total-THistoD.GetValue ('PHB_BASECOT');
                  end
//PT125-6
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='13') then
                  begin
                  if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                     Total:= Total+THistoD.GetValue ('PHB_MTPATRONAL')
                  else
                     Total:= Total-THistoD.GetValue ('PHB_MTPATRONAL');
                  end
//FIN PT125-6
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='14') then
                  begin
                  if (THistoD.GetValue ('PHB_SENSBUL')<>'M') then
                     Total:= Total+THistoD.GetValue ('PHB_MTSALARIAL')
                  else
                     Total:= Total-THistoD.GetValue ('PHB_MTSALARIAL');
                  end
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='25') then
                  Total:= Total+THistoD.GetValue ('PHB_TAUXSALARIAL')
//PT125-6
               else
               if (THistoD.GetValue ('PUO_CHAMPRETENU')='26') then
                  Total:= Total+THistoD.GetValue ('PHB_TAUXPATRONAL');
//FIN PT125-6
               end
            else
            if (THistoD.GetValue ('PUO_NATURERUB')='CUM') then
               Total:= Total+THistoD.GetValue('PHC_MONTANT');
            THistoD:= THisto.FindNext ([''], [''], TRUE);
            end;
{PT129-1
      TNatureD:= TNature.FindNext ([''], [''], True);
}
      TNatureD:= TNature.FindNext (['PUO_PREDEFINI'], [Predefini], True);
//FIN PT129-1
      end;

FreeAndNil (THisto);
FreeAndNil (TNature);
result:= Total;
end;
//FIN PT101-11

//PT106
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/11/2006
Modifié le ... :   /  /    
Description .. : Calcul des éléments IRCANTEC
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure RemplitIRCANTEC (Salarie : string; Periode : integer);
var
CleDADS : TCleDADS;
Total : Double;
TSalD : TOB;
begin
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= StrToDate (StPerDeb);
CleDADS.DateFin:= StrToDate (StPerFin);
CleDADS.Exercice:= PGExercice;

CreeDetail (CleDADS, 891, 'S42.G01.00.001', '07', '07');

{PT125-6
NombreMois (CleDADS.DateDeb, CleDADS.DateFin, PremMois, PremAnnee, NbreMois);
CreeDetail (CleDADS, 892, 'S42.G01.00.002', IntToStr (NbreMois),
            IntToStr (NbreMois));
}            

//PT108
{PT125-6
if ((BufProf='63') or (BufProf='65') or (BufProf='66') or (BufProf='68')) then
}
if ((BufProf='63') or (BufProf='65')) then
//FIN PT125-6
   begin
   TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
{PT135
   CreeDetail (CleDADS, 893, 'S42.G01.00.004',
               FloatToStr (TSalD.GetValue ('PSA_HORHEBDO')*100),
               FloatToStr (TSalD.GetValue ('PSA_HORHEBDO')));
}
   CreeDetail (CleDADS, 893, 'S42.G01.00.004',
               FloatToStr (Arrondi (TSalD.GetValue ('PSA_HORHEBDO')*100, 0)),
               FloatToStr (Arrondi (TSalD.GetValue ('PSA_HORHEBDO'), 2)));
//FIN PT135               
   end;
//FIN PT108

Total:= RemplitFromPublicotis (Salarie, '511', CleDADS.DateDeb, CleDADS.DateFin);
CreeDetail (CleDADS, 894, 'S42.G01.00.007.001',
            FloatToStr(Abs(Arrondi(Total, 0))),
            FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 895, 'S42.G01.00.007.002', 'N', 'N');

Total:= RemplitFromPublicotis (Salarie, '512', CleDADS.DateDeb, CleDADS.DateFin);
CreeDetail (CleDADS, 896, 'S42.G01.00.008.001',
            FloatToStr(Abs(Arrondi(Total, 0))),
            FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 897, 'S42.G01.00.008.002', 'N', 'N');

//PT108
if ((BufProf='52') or (BufProf='53') or (BufProf='54') or (BufProf='56') or
   (BufProf='57') or (BufProf='58') or (BufProf='59')) then
   begin
   Total:= RemplitFromPublicotis (Salarie, '513', CleDADS.DateDeb,
                                  CleDADS.DateFin);
   CreeDetail (CleDADS, 898, 'S42.G01.00.009.001',
               FloatToStr(Abs(Arrondi(Total, 0))),
               FloatToStr(Arrondi(Total, 0)));
   if (Total < 0) then
      CreeDetail (CleDADS, 899, 'S42.G01.00.009.002', 'N', 'N');
   end;
//FIN PT108

{A implémenter plus tard
CreeDetail (CleDADS, 901, 'S42.G02.00.001', '', '');
CreeDetail (CleDADS, 902, 'S42.G02.00.002.001', '', '');
CreeDetail (CleDADS, 903, 'S42.G02.00.002.002', '', '');
}
end;
//FIN PT106


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/06/2004
Modifié le ... :   /  /
Description .. : Calcul des éléments des périodes d'inactivité
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure RemplitInact(Salarie : string);
var
TAbsD, TAbs1D, THisto, THistoD : TOB;
Ordre : integer;
PerDeb, PerFin : TDateTime;
Buffer, Motif, StSQL, Unite : string;
CleDADS : TCleDADS;
QRechSQL : TQuery;
Ircantec : boolean;
Total : double;
begin
Ordre:= -1;
TAbs1D := TAbs1.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
if (TAbs1D <> nil) then
   begin
{PT101-1
   Writeln(FRapport, '');
   Writeln(FRapport, '');
   Writeln(FRapport, 'Salarié : '+TAbs1D.GetValue('PSA_SALARIE')+' '+
           TAbs1D.GetValue('PSA_LIBELLE')+' '+TAbs1D.GetValue('PSA_PRENOM'));
}
   end;
TAbsD := TAbs.FindFirst(['PCN_SALARIE'], [Salarie], TRUE);
While (TAbsD <> nil) do
      begin
      if (((TAbsD.GetValue('PCN_DATEDEBUTABS')>=DebExer) and
         (TAbsD.GetValue('PCN_DATEDEBUTABS')<=FinExer)) or
         ((TAbsD.GetValue('PCN_DATEFINABS')>=DebExer) and
         (TAbsD.GetValue('PCN_DATEFINABS')<=FinExer)) or
         ((TAbsD.GetValue('PCN_DATEDEBUTABS')<DebExer) and
         (TAbsD.GetValue('PCN_DATEFINABS')>FinExer))) then
         begin
{PT125-6
         if (TAbsD.GetValue('PCN_DATEDEBUTABS')<DebExer) then
            PerDeb:= DebExer
         else
}
         PerDeb:= TAbsD.GetValue('PCN_DATEDEBUTABS');
         if (TAbsD.GetValue('PCN_DATEFINABS')>FinExer) then
            PerFin:= FinExer
         else
            PerFin:= TAbsD.GetValue('PCN_DATEFINABS');

{PT101-1
         Writeln (FRapport, '********** PERIODE N° '+IntToStr(-Ordre));
         Writeln (FRapport, '********** '+
                  RechDom('PGTYPEABS', TAbsD.GetValue('PMA_TYPEABS'), FALSE)+
                  ' du '+DateToStr(TAbsD.GetValue('PCN_DATEDEBUTABS'))+
                  ' au '+DateToStr(TAbsD.GetValue('PCN_DATEFINABS')));
}
         Motif:= RechDom('PGTYPEABS', TAbsD.GetValue('PMA_TYPEABS'), True);
         if (Motif<>'') then
            begin
{PT101-6
            CreeEntete (Salarie, TypeD, Ordre, PerDeb, PerFin, Motif, '');
}
            CleDADS.Salarie:= Salarie;
            CleDADS.TypeD:= TypeD;
            CleDADS.Num:= Ordre;
            CleDADS.DateDeb:= PerDeb;
            CleDADS.DateFin:= PerFin;
            CleDADS.Exercice:= PGExercice;
            CreeEntete (CleDADS, Motif, '', Decalage);
//FIN PT101-6
            CreeDetail (CleDADS, 1, 'S46.G01.00.001', Motif,
                        TAbsD.GetValue('PMA_TYPEABS'));
Debug('Paie PGI/Calcul DADS-U : S46.G01.00.001');

            ForceNumerique(DateToStr(TAbsD.GetValue('PCN_DATEDEBUTABS')), Buffer);
{PT125-6
            CreeDetail (CleDADS, 2, 'S46.G01.00.002', Copy(Buffer, 1, 4),
                        DateToStr(TAbsD.GetValue('PCN_DATEDEBUTABS')));
}
            CreeDetail (CleDADS, 2, 'S46.G01.00.002', Buffer,
                        DateToStr(TAbsD.GetValue('PCN_DATEDEBUTABS')));
//FIN PT125-6
Debug('Paie PGI/Calcul DADS-U : S46.G01.00.002');

            ForceNumerique(DateToStr(TAbsD.GetValue('PCN_DATEFINABS')), Buffer);
{PT125-6
            CreeDetail (CleDADS, 3, 'S46.G01.00.003', Copy(Buffer, 1, 4),
                        DateToStr(TAbsD.GetValue('PCN_DATEFINABS')));
}
            CreeDetail (CleDADS, 3, 'S46.G01.00.003', Buffer,
                        DateToStr(TAbsD.GetValue('PCN_DATEFINABS')));
//FIN PT125-6                        
Debug('Paie PGI/Calcul DADS-U : S46.G01.00.003');

//PT121
{PT125-6
            Buffer:= DateToStr (DebExer);
            if ((Copy (Buffer, 1, 5)='01/12') and
               ((Copy (TypeD, 2, 2)='01') or (Copy (TypeD, 2, 2)='02'))) then
               Buffer:= '01/01/'+ IntToStr (StrToInt (Copy (Buffer, 7, 4))+1);
            DateBuf:= DebutDeMois (StrToDate (Buffer));

            if ((PerDeb<DateBuf) and (PerFin<DateBuf)) then
               CreeDetail (CleDADS, 4, 'S46.G01.00.004', '02', '02')
            else
            if ((PerDeb<DateBuf) and (PerFin>DateBuf)) then
               CreeDetail (CleDADS, 4, 'S46.G01.00.004', '03', '03');
}
            if (PerDeb<DebExer) then
               CreeDetail (CleDADS, 4, 'S46.G01.00.004', '01', '01');
//FIN PT125-6
Debug('Paie PGI/Calcul DADS-U : S46.G01.00.004');
//FIN PT121

//PT101-11
            StSQL:= 'SELECT PDS_SALARIE'+
                    ' FROM DADSDETAIL WHERE'+
                    ' PDS_SALARIE="'+Salarie+'" AND'+
                    ' PDS_SEGMENT="S41.G01.01.001" AND'+
                    ' (PDS_DONNEEAFFICH="I001" OR PDS_DONNEEAFFICH="I002")';
            Ircantec:= ExisteSql (StSQL);
            if (Ircantec=True) then
               begin
               CreeDetail (CleDADS, 5, 'S46.G01.01.001', '01','01');

               Total:= RemplitFromPublicotis (Salarie, '691', PerDeb, PerFin);
               CreeDetail (CleDADS, 6, 'S46.G01.01.002.001',
                           FloatToStr(Abs(Arrondi(Total, 0))),
                           FloatToStr(Arrondi(Total, 0)));
               if (Total < 0) then
                  CreeDetail (CleDADS, 7, 'S46.G01.01.002.002', 'N', 'N');
               end
            else
               begin
//FIN PT101-11

{PT101-2
            if (TAbsD.GetValue('PMA_JOURHEURE')='HEU') then
}
               if ((TAbsD.GetValue('PMA_JOURHEURE')='HEU') or
                   (TAbsD.GetValue('PMA_JOURHEURE')='HOR') or //PT125
                  (Motif ='07')) then
//FIN PT101-2
                  Unite:= '01'
               else
                  Unite:= '03';
{PT125-6
               CreeDetail (CleDADS, 9, 'S46.G01.02.001', Unite, Unite);
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.001');
}
               CreeDetail (CleDADS, 9, 'S46.G01.02.001.001', Unite, Unite);
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.001.001');
               CreeDetail (CleDADS, 10, 'S46.G01.02.001.002', '02', '02');
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.001.002');
//FIN PT125-6

               if (Unite='01') then
                  CreeDetail (CleDADS, 11, 'S46.G01.02.002',
                              TAbsD.GetValue('PCN_HEURES')*100,
                              TAbsD.GetValue('PCN_HEURES'))
               else
                  CreeDetail (CleDADS, 11, 'S46.G01.02.002',
                              TAbsD.GetValue('PCN_JOURS')*100,
                              TAbsD.GetValue('PCN_JOURS'));
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.002');

//PT101-2
               if (Motif ='07') then
                  begin
//Chargement de la TOB MOTIFABSENCE/HISTOBULLETIN
                  StSQL:= 'SELECT PHB_MTREM'+
                          ' FROM HISTOBULLETIN'+
                          ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI##'+
                          ' PHB_RUBRIQUE=PMA_RUBRIQUE OR'+
                          ' PHB_RUBRIQUE=PMA_RUBRIQUEJ WHERE'+
                          ' PHB_SALARIE="'+Salarie+'" AND'+
                          ' PHB_DATEDEBUT<="'+UsDateTime(PerFin)+'" AND'+
                          ' PHB_DATEFIN>="'+UsDateTime(PerFin)+'" AND'+
                          ' PHB_NATURERUB="AAA" AND'+
                          ' PMA_TYPEABS="CHO"';
                  QRechSQL:= OpenSql (StSQL, TRUE);
                  THisto:= TOB.Create ('Les historiques', NIL, -1);
                  THisto.LoadDetailDB ('HISTOBULLETIN','','',QRechSQL,False);
                  Ferme(QRechSQL);

{PT120
                  CreeDetail (CleDADS, 11, 'S46.G01.02.003.001',
                              FloatToStr (Arrondi (StrToFloat (THisto.GetValue('PHB_MTREM')), 0)),
                              FloatToStr (StrToFloat (THisto.GetValue('PHB_MTREM'))));
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.003.001');

                  if (THisto.GetValue('PHB_MTREM') < 0) then
}
                  THistoD:= THisto.FindFirst([''], [''], False);
                  CreeDetail (CleDADS, 12, 'S46.G01.02.003.001',
                              FloatToStr (Arrondi (StrToFloat (THistoD.GetValue('PHB_MTREM')), 0)),
                              FloatToStr (StrToFloat (THistoD.GetValue('PHB_MTREM'))));
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.003.001');

                  if (THistoD.GetValue('PHB_MTREM') < 0) then
//FIN PT120
                     CreeDetail (CleDADS, 13, 'S46.G01.02.003.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S46.G01.02.003.002');
                  FreeAndNil (THisto);
                  end;
//FIN PT101-2
               end;
            Ordre:= Ordre-1;
            end;
         end;
      TAbsD := TAbs.FindNext(['PCN_SALARIE'], [Salarie], TRUE);
      end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/02/2002
Modifié le ... :
Description .. : Pour une période donnée, calcul des éléments qui
Suite ........ : rempliront les tables DADSPERIODES et DADSDETAIL en
Suite ........ : ce qui concerne les Congés payés pour les caisses BTP
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CalculBTP(Salarie : string);
var
TEtabD, TRemD, TSalD, TSalComplD : TOB;
Anciennete : WORD;
Buf, Horaire : string;
MontantH, NMois, Total : Double;
THistoCumSalD : TOB;
DerDate : TDateTime;
AnneeA, DerMois, JourJ, MoisM : Word;
CleDADS : TCleDADS;
ErreurDADSU : TControle;
begin
//Recherche du salarié concerné dans la TOB Salariés
TSalD := TSal.FindFirst(['PSA_SALARIE'], [Salarie], TRUE);
TSalComplD := TSalCompl.FindFirst(['PSE_SALARIE'], [Salarie], TRUE);
//PT101-6
CleDADS.Salarie:= Salarie;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Periode;
CleDADS.DateDeb:= StrToDate (StPerDeb);
CleDADS.DateFin:= StrToDate (StPerFin);
CleDADS.Exercice:= PGExercice;

ErreurDADSU.Salarie:= Salarie;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Num:= Periode;
ErreurDADSU.DateDeb:= StrToDate (StPerDeb);
ErreurDADSU.DateFin:= StrToDate (StPerFin);
ErreurDADSU.Exercice:= PGExercice;
//FIN PT101-6

{Ecriture dans le fichier de contrôle du matricule, nom et prénom du salarié
concerné}
{PT101-1
Writeln(FRapport, '********** Congés payés BTP pour le salarié : '+
                  TSalD.GetValue('PSA_SALARIE')+' '+
                  TSalD.GetValue('PSA_LIBELLE')+' '+
                  TSalD.GetValue('PSA_PRENOM'));
}

{Traitement effectué uniquement si les compléments BTP du salarié ont été
trouvés sinon écriture d'une erreur dans le fichier de contrôle}
if (TSalComplD <> nil) then
   begin
{Si le numéro d'adhésion n'est pas renseigné, écriture d'une erreur dans le
fichier de contrôle}
   if ((TSalComplD.GetValue('PSE_BTPADHESION') <> '') and
      (TSalComplD.GetValue('PSE_BTPADHESION') <> null)) then
      CreeDetail (CleDADS, 971, 'S66.G01.00.001',
                  Copy(TSalComplD.GetValue('PSE_BTPADHESION'), 1, 15),
                  Copy(TSalComplD.GetValue('PSE_BTPADHESION'), 1, 15))
   else
{PT101-1
      EcrireErreur('Erreur:   Fiche Salarié BTP : Numéro d''adhésion', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S66.G01.00.001';
      ErreurDADSU.Explication:= 'Le numéro d''adhésion n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   end
else
{PT101-1
   EcrireErreur('Erreur:   Fiche Salarié BTP absente', Erreur);
}
   begin
   ErreurDADSU.Segment:= '';
   ErreurDADSU.Explication:= 'La fiche Salarié BTP absente';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT101-1
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.001');

{PT125-6
CreeDetail (CleDADS, 972, 'S66.G01.00.002', '01', '01');
}
CreeDetail (CleDADS, 972, 'S66.G01.00.002.001', '01', '01');
CreeDetail (CleDADS, 973, 'S66.G01.00.002.002', '02', '02');
//FIN PT125-6

Total := 0;
THistoCumSalD := THistoCumSal.FindFirst(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                        [Salarie, '20'], TRUE);
While (THistoCumSalD <> nil) do
      begin
      Total := Total + THistoCumSalD.GetValue('PHC_MONTANT');
      THistoCumSalD := THistoCumSal.FindNext(['PHC_SALARIE', 'PHC_CUMULPAIE'],
                                             [Salarie, '20'], TRUE);
      end;
{PT135
CreeDetail (CleDADS, 974, 'S66.G01.00.003', FloatToStr(Total*100),
            FloatToStr(Total));
}
CreeDetail (CleDADS, 974, 'S66.G01.00.003', FloatToStr (Arrondi (Total*100, 0)),
            FloatToStr (Arrondi (Total, 2)));
//FIN PT135               
if Total = 0 then
{PT101-1
   EcrireErreur ('          Saisie complémentaire DADS-U BTP :'+
                 ' Le temps de travail est égal à 0', Erreur);
}
   begin
   ErreurDADSU.Segment:= 'S66.G01.00.003';
   ErreurDADSU.Explication:= 'Le temps de travail est égal à 0';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT101-1
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.003');

{PT96
if ((TSalD.GetValue('PSA_DATEANCIENNETE') <> IDate1900) and
   (TSalD.GetValue('PSA_DATEANCIENNETE') <> null)) then
}
if ((TSalD.GetValue('PSA_DATEANCIENNETE') > IDate1900) and
   (TSalD.GetValue('PSA_DATEANCIENNETE') <> null)) then
   begin
   Anciennete:= AncienneteAnnee (TSalD.GetValue('PSA_DATEANCIENNETE'), FinExer );
   CreeDetail (CleDADS, 975, 'S66.G01.00.004', IntToStr(Anciennete),
               IntToStr(Anciennete));
   end
else
   begin
{PT101-1
   CreeDetail (CleDADS, 904, 'S66.G01.00.004', '0', '0');
   EcrireErreur('Erreur:   Fiche Salarié : Date d''ancienneté', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.004';
   ErreurDADSU.Explication:= 'La date d''ancienneté n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.004');

if (TSalComplD <> nil) then
   begin
{PT96
   if ((TSalComplD.GetValue('PSE_BTPANCIENNETE') <> IDate1900) and
      (TSalComplD.GetValue('PSE_BTPANCIENNETE') <> null)) then
}
   if ((TSalComplD.GetValue('PSE_BTPANCIENNETE') > IDate1900) and
      (TSalComplD.GetValue('PSE_BTPANCIENNETE') <> null)) then
      begin
      Anciennete:= AncienneteAnnee (TSalComplD.GetValue('PSE_BTPANCIENNETE'), FinExer );
      CreeDetail (CleDADS, 976, 'S66.G01.00.005', IntToStr(Anciennete),
                  IntToStr(Anciennete));
      end
   else
      begin
      CreeDetail (CleDADS, 976, 'S66.G01.00.005', '99', '99');
{PT101-1
      EcrireErreur('          Fiche Salarié BTP :'+
                   ' L''ancienneté dans le BTP a été forcé à ''inconnue''',
                   Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.005';
      ErreurDADSU.Explication:= 'L''ancienneté dans le BTP a été forcé à'+
                                ' ''inconnue''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.005');

if (TSalComplD <> nil) then
   begin
   if ((TSalComplD.GetValue('PSE_BTPASSEDIC') <> '') and
      (TSalComplD.GetValue('PSE_BTPASSEDIC') <> null)) then
      begin
      Buf := TSalComplD.GetValue('PSE_BTPASSEDIC');
      if Buf='OUI' then
         CreeDetail (CleDADS, 977, 'S66.G01.00.006', '01', Buf)
      else
         CreeDetail (CleDADS, 977, 'S66.G01.00.006', '02', Buf);
      end
   else
      begin
      CreeDetail (CleDADS, 977, 'S66.G01.00.006', '01', '01');
{PT101-1
      EcrireErreur ('          Fiche Salarié BTP :'+
                    ' Code bénéficiaire ASSEDIC', Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.006';
      ErreurDADSU.Explication:= 'Le code bénéficiaire ASSEDIC a été forcé à'+
                                ' ''oui''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.006');

if (TSalComplD <> nil) then
   begin
   if ((TSalComplD.GetValue('PSE_BTPSALMOYEN') <> '') and
      (TSalComplD.GetValue('PSE_BTPSALMOYEN') <> null)) then
      CreeDetail (CleDADS, 978, 'S66.G01.00.007',
                  TSalComplD.GetValue('PSE_BTPSALMOYEN'),
                  TSalComplD.GetValue('PSE_BTPSALMOYEN'))
   else
      begin
      CreeDetail (CleDADS, 978, 'S66.G01.00.007', '01', '01');
{PT101-1
      EcrireErreur ('          Fiche Salarié BTP : Unité du salaire moyen',
                    Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.007';
      ErreurDADSU.Explication:= 'L''unité du salaire moyen a été forcée à'+
                                ' ''salaire horaire''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.007');

if ((TSalComplD <> nil) and
   (TSalComplD.GetValue('PSE_BTPSALMOYEN') <> '03')) then
   begin
   Total:= 0;
   NMois:= 1;
   TRemD:= TRem.FindFirst(['PHB_SALARIE', 'PRM_BTPSALAIRE'],
                          [Salarie, 'X'], TRUE);
   if TRemD <> nil then
      begin
      DerDate := TRemD.GetValue('PHB_DATEDEBUT');
      DecodeDate(DerDate, AnneeA, MoisM, JourJ);
      DerMois := MoisM;
      While (TRemD <> nil) do
            begin
            DerDate := TRemD.GetValue('PHB_DATEDEBUT');
            DecodeDate(DerDate, AnneeA, MoisM, JourJ);
            if (DerMois <> MoisM) then
               NMois := NMois + 1;
            DerMois := MoisM;
{PT91
            Total:= Total+TRemD.GetValue('PHB_BASEREM');
}
            if (TRemD.GetValue('PHB_SENSBUL')<>'M') then
               Total:= Total+TRemD.GetValue('PHB_BASEREM')
            else
               Total:= Total-TRemD.GetValue('PHB_BASEREM');
//FIN PT91
            TRemD:= TRem.FindNext(['PHB_SALARIE', 'PRM_BTPSALAIRE'],
                                  [Salarie, 'X'], TRUE);
            end;
      end;
   Total := Total / NMois;
   CreeDetail (CleDADS, 979, 'S66.G01.00.008',
               FloatToStr(Arrondi(Total*100, 0)),
               FloatToStr(Arrondi(Total, 2)));
   if (Total = 0) then
{PT101-1
      EcrireErreur ('          Saisie complémentaire DADS-U BTP :'+
                    ' Le salaire moyen est égal à 0', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S66.G01.00.008';
      ErreurDADSU.Explication:= 'Le salaire moyen est égal à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.008');

Buf := '';
if (TSalComplD <> nil) then
   begin
   if ((TSalComplD.GetValue('PSE_BTPHORAIRE') <> '') and
      (TSalComplD.GetValue('PSE_BTPHORAIRE') <> null)) then
      begin
      Buf := TSalComplD.GetValue('PSE_BTPHORAIRE');
      if (Buf = 'HEB') then
         Horaire := '01'
      else
         Horaire := '02';
      end
   else
      begin
      Buf := 'MEN';
      Horaire := '02';
{PT101-1
      EcrireErreur ('          Fiche Salarié BTP :'+
                    ' Le code type horaire a été forcé à ''mensuel''', Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.009';
      ErreurDADSU.Explication:= 'Le code type horaire a été forcé à ''mensuel''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end
else
   begin
   Buf := 'MEN';
   Horaire := '02';
{PT101-1
   EcrireErreur ('          Fiche Salarié BTP :'+
                 ' Le code type horaire a été forcé à ''mensuel''', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.009';
   ErreurDADSU.Explication:= 'Le code type horaire a été forcé à ''mensuel''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;

CreeDetail (CleDADS, 980, 'S66.G01.00.009', Horaire, Buf);
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.009');

MontantH := 0;
if Horaire = '01' then
   begin
   if (TSalD.GetValue('PSA_HORHEBDO') <>  null) then
      MontantH := TSalD.GetValue('PSA_HORHEBDO')
   else
{PT101-1
      EcrireErreur('          Fiche Salarié : horaire hebdomadaire', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S66.G01.00.010';
      ErreurDADSU.Explication:= 'L''horaire hebdomadaire est à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   end
else
   begin
   if (TSalD.GetValue('PSA_HORAIREMOIS') <> null) then
      MontantH := TSalD.GetValue('PSA_HORAIREMOIS')
   else
{PT101-1
      EcrireErreur('          Fiche Salarié : horaire mensuel', Erreur);
}
      begin
      ErreurDADSU.Segment:= 'S66.G01.00.010';
      ErreurDADSU.Explication:= 'L''horaire mensuel est à 0';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT101-1
   end;
{PT135
CreeDetail (CleDADS, 981, 'S66.G01.00.010', FloatToStr(MontantH*100),
            FloatToStr(MontantH));
}
CreeDetail (CleDADS, 981, 'S66.G01.00.010', FloatToStr(Arrondi (MontantH*100, 0)),
            FloatToStr (Arrondi (MontantH, 2)));
//FIN PT135               
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.010');

TEtabD := TEtab.FindFirst(['ET_ETABLISSEMENT'],
                          [TSalD.GetValue('PSA_ETABLISSEMENT')], TRUE);
if ((TEtabD <> nil) and (TEtabD.GetValue('ETB_HORAIREETABL') <> null)) then
   begin
   if Horaire = '01' then
      CreeDetail (CleDADS, 982, 'S66.G01.00.011',
                  FloatToStr(Arrondi(TEtabD.GetValue('ETB_HORAIREETABL')*1200/52, 0)),
                  FloatToStr(Arrondi(TEtabD.GetValue('ETB_HORAIREETABL')*12/52, 0)))
   else
      CreeDetail (CleDADS, 982, 'S66.G01.00.011',
                  FloatToStr(Arrondi(TEtabD.GetValue('ETB_HORAIREETABL')*100, 0)),
                  FloatToStr(Arrondi(TEtabD.GetValue('ETB_HORAIREETABL'), 2)));
   end
else
   begin
   CreeDetail (CleDADS, 982, 'S66.G01.00.011', '0', '0');
{PT101-1
   EcrireErreur('          Fiche Etablissement '+TSalD.GetValue('PSA_ETABLISSEMENT')+
               ': horaire de l''établissement', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.011';
   ErreurDADSU.Explication:= 'L''horaire de l''établissement est à 0';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.011');

{PT102-1
if (TSalD.GetValue('PSA_SITUATIONFAMIL') <> '') then
}
if ((TSalD.GetValue('PSA_SITUATIONFAMIL') <> '') and
   (TSalD.GetValue('PSA_SITUATIONFAMIL') <> '90')) then
   CreeDetail (CleDADS, 983, 'S66.G01.00.012',
               RechDom('PGSITUATIONFAMIL', TSalD.GetValue('PSA_SITUATIONFAMIL'), TRUE),
               TSalD.GetValue('PSA_SITUATIONFAMIL'))
else
   begin
   CreeDetail (CleDADS, 983, 'S66.G01.00.012', '99', '90');
{PT101-1
   EcrireErreur ('          Fiche Salarié :'+
                 ' Le code situation familiale a été forcé à ''non connue''',
                 Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.012';
   ErreurDADSU.Explication:= 'Le code situation familiale a été forcé à ''non'+
                             ' connue''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.012');

//PT103-2
if (TSalD.GetValue('PSA_PERSACHARGE')='99') then
   CreeDetail (CleDADS, 984, 'S66.G01.00.013', '90', '90')
else
//FIN PT103-2
   CreeDetail (CleDADS, 984, 'S66.G01.00.013', TSalD.GetValue('PSA_PERSACHARGE'),
               TSalD.GetValue('PSA_PERSACHARGE'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.013');

if (FloatToStr(TSalD.GetValue('PSA_PCTFRAISPROF'))<>'0') then
   CreeDetail (CleDADS, 985, 'S66.G01.00.014',
               FormatFloat('00', TSalD.GetValue('PSA_PCTFRAISPROF')),
               TSalD.GetValue('PSA_PCTFRAISPROF'))
else
   begin
   CreeDetail (CleDADS, 985, 'S66.G01.00.014', '99',
               TSalD.GetValue('PSA_PCTFRAISPROF'));
{PT101-1
   EcrireErreur('          Fiche Salarié : Le taux d''abattement pour frais'+
                ' professionnels a été forcé à ''pas d''abattement''', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.014';
   ErreurDADSU.Explication:= 'Le taux d''abattement pour frais professionnels'+
                             ' a été forcé à ''pas d''abattement''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.014');

if (TSalComplD <> nil) then
   begin
   if ((TSalComplD.GetValue('PSE_STATUTCOTIS') <> '') and
      (TSalComplD.GetValue('PSE_STATUTCOTIS') <> null)) then
      CreeDetail (CleDADS, 986, 'S66.G01.00.015',
                  TSalComplD.GetValue('PSE_STATUTCOTIS'),
                  TSalComplD.GetValue('PSE_STATUTCOTIS'))
   else
      begin
      CreeDetail (CleDADS, 986, 'S66.G01.00.015', '90', '90');
{PT101-1
      EcrireErreur('          Fiche Salarié BTP : Le statut cotisant a été forcé'+
                   ' à ''Pas congés payés ni autres BTP''', Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.015';
      ErreurDADSU.Explication:= 'Le statut cotisant a été forcé à ''Pas congés'+
                                ' payés ni autres BTP''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end
else
   begin
   CreeDetail (CleDADS, 986, 'S66.G01.00.015', '90', '90');
{PT101-1
   EcrireErreur('          Fiche Salarié BTP : Le statut cotisant a été forcé à'+
                ' ''Pas congés payés ni autres BTP''', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.015';
   ErreurDADSU.Explication:= 'Le statut cotisant a été forcé à ''Pas congés'+
                             ' payés ni autres BTP''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.015');

{PT101-2
if (TSalD.GetValue('PSA_NIVEAU') <> '') then
   CreeDetail (CleDADS, 916, 'S66.G01.00.016', TSalD.GetValue('PSA_NIVEAU'),
               TSalD.GetValue('PSA_NIVEAU'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.016');

if (TSalD.GetValue('PSA_COEFFICIENT') <> '') then
   CreeDetail (CleDADS, 917, 'S66.G01.00.017',
               TSalD.GetValue('PSA_COEFFICIENT'),
               TSalD.GetValue('PSA_COEFFICIENT'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.017');

if ((TSalComplD <> nil) and ((TSalComplD.GetValue('PSE_BTPPOSITION') <> '') and
   (TSalComplD.GetValue('PSE_BTPPOSITION') <> null))) then
   CreeDetail (CleDADS, 918, 'S66.G01.00.018',
               TSalComplD.GetValue('PSE_BTPPOSITION'),
               TSalComplD.GetValue('PSE_BTPPOSITION'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.018');

if ((TSalComplD <> nil) and ((TSalComplD.GetValue('PSE_BTPECHELON') <> '') and
   (TSalComplD.GetValue('PSE_BTPECHELON') <> null))) then
   CreeDetail (CleDADS, 919, 'S66.G01.00.019',
               TSalComplD.GetValue('PSE_BTPECHELON'),
               TSalComplD.GetValue('PSE_BTPECHELON'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.019');

if ((TSalComplD <> nil) and ((TSalComplD.GetValue('PSE_BTPCATEGORIE') <> '') and
   (TSalComplD.GetValue('PSE_BTPCATEGORIE') <> null))) then
   CreeDetail (CleDADS, 920, 'S66.G01.00.020',
               TSalComplD.GetValue('PSE_BTPCATEGORIE'),
               TSalComplD.GetValue('PSE_BTPCATEGORIE'));
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.020');
}

if (TSalComplD <> nil) then
   begin
   if ((TSalComplD.GetValue('PSE_METIERBTP') <> '') and
      (TSalComplD.GetValue('PSE_METIERBTP') <> null)) then
      CreeDetail (CleDADS, 991, 'S66.G01.00.021',
                  TSalComplD.GetValue('PSE_METIERBTP'),
                  TSalComplD.GetValue('PSE_METIERBTP'))
   else
      begin
      CreeDetail (CleDADS, 991, 'S66.G01.00.021', 'A0000', 'A0000');
{PT101-1
      EcrireErreur('          Fiche Salarié BTP : Le code métier BTP a été'+
                   ' forcé à ''Agenceur''', Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.021';
      ErreurDADSU.Explication:= 'Le code métier BTP a été forcé à ''Agenceur''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;

   if ((TSalComplD.GetValue('PSE_BTPAFFILIRCIP') <> '') and
      (TSalComplD.GetValue('PSE_BTPAFFILIRCIP') <> null)) then
      CreeDetail (CleDADS, 992, 'S66.G01.00.022',
                  TSalComplD.GetValue('PSE_BTPAFFILIRCIP'),
                  TSalComplD.GetValue('PSE_BTPAFFILIRCIP'))
   else
      begin
      CreeDetail (CleDADS, 992, 'S66.G01.00.022', '01', '01');
{PT101-1
      EcrireErreur('          Fiche Salarié BTP : Le type affiliation IRC/IP a'+
                   ' été forcé à ''Ouvrier''', Erreur);
}
      ErreurDADSU.Segment:= 'S66.G01.00.022';
      ErreurDADSU.Explication:= 'Le type affiliation IRC/IP a été forcé à'+
                                ' ''Ouvrier''';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
//FIN PT101-1
      end;
   end
else
   begin
   CreeDetail (CleDADS, 991, 'S66.G01.00.021', 'A0000', 'A0000');
{PT101-1
   EcrireErreur('          Fiche Salarié BTP : Le code métier BTP a été forcé'+
                ' à ''Agenceur''', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.021';
   ErreurDADSU.Explication:= 'Le code métier BTP a été forcé à ''Agenceur''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1

   CreeDetail (CleDADS, 992, 'S66.G01.00.022', '01', '01');
{PT101-1
   EcrireErreur('          Fiche Salarié BTP : Le Le type affiliation IRC/IP a'+
                ' été forcé à ''Ouvrier''', Erreur);
}
   ErreurDADSU.Segment:= 'S66.G01.00.022';
   ErreurDADSU.Explication:= 'Le type affiliation IRC/IP a été forcé à'+
                             ' ''Ouvrier''';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
//FIN PT101-1
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.021');
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.022');

{PT101-12
Total := 0;
TBaseD := TBase.FindFirst(['PHB_SALARIE', 'PCT_BRUTSS'], [Salarie, 'X'], TRUE);
While ((TBaseD <> nil) and
      ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) or
      (TBaseD.GetValue('PHB_DATEDEBUT')>StrToDate(StPerFin)))) do
      begin
      if (((TBaseD.GetValue('PHB_DATEFIN')>StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')<=StrToDate(StPerFin))) or
         ((TBaseD.GetValue('PHB_DATEDEBUT')<StrToDate(StPerDeb)) and
         (TBaseD.GetValue('PHB_DATEFIN')>=StrToDate(StPerFin)))) then
         Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BRUTSS'],
                               [Salarie, 'X'], TRUE);
      end;
While ((TBaseD <> nil) and
      (TBaseD.GetValue('PHB_DATEDEBUT')>=StrToDate(StPerDeb)) and
      (TBaseD.GetValue('PHB_DATEDEBUT')<=StrToDate(StPerFin))) do
      begin
      Total := Total + TBaseD.GetValue('PHB_BASECOT');
      TBaseD := TBase.FindNext(['PHB_SALARIE', 'PCT_BRUTSS'],
                               [Salarie, 'X'], TRUE);
      end;
}
Total:= RemplitFromPublicotis (Salarie, '783', StrToDate(StPerDeb),
                               StrToDate(StPerFin));
//FIN PT101-12
CreeDetail (CleDADS, 993, 'S66.G01.00.023.001',
            FloatToStr(Abs(Arrondi(Total, 0))),
            FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 994, 'S66.G01.00.023.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.023');

//PT101-12
Total:= RemplitFromPublicotis (Salarie, '784', StrToDate(StPerDeb),
                               StrToDate(StPerFin));
//FIN PT101-12
CreeDetail (CleDADS, 995, 'S66.G01.00.024.001',
            FloatToStr(Abs(Arrondi(Total, 0))),
            FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 996, 'S66.G01.00.024.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.024');

//PT101-2
Total:= RemplitFromPublicotis (Salarie, '785', StrToDate(StPerDeb),
                               StrToDate(StPerFin));
CreeDetail (CleDADS, 997, 'S66.G01.00.025.001',
            FloatToStr(Abs(Arrondi(Total, 0))),
            FloatToStr(Arrondi(Total, 0)));
if (Total < 0) then
   CreeDetail (CleDADS, 998, 'S66.G01.00.025.002', 'N', 'N');
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.025');

if (TSalComplD <> nil) then
   begin
   if (TSalComplD.GetValue('PSE_BTPPOSITION') <> null) then
      CreeDetail (CleDADS, 999, 'S66.G01.00.026',
                  TSalComplD.GetValue('PSE_BTPPOSITION'),
                  TSalComplD.GetValue('PSE_BTPPOSITION'));
   if ((TSalComplD.GetValue('PSE_BTPPOSITION') = '') or
      (TSalComplD.GetValue('PSE_BTPPOSITION') = null)) then
      begin
      ErreurDADSU.Segment:= 'S66.G01.00.026';
      ErreurDADSU.Explication:= 'Le code classification BTP n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
   end;
Debug('Paie PGI/Calcul DADS-U : S66.G01.00.026');
//FIN PT101-2
end;


{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2001
Modifié le ... : 27/08/2001
Description .. : Récupération à partir d'un champ ou d'une chaine du
Suite ........ : numéro de voie et du nom de la voie
Mots clefs ... : PAIE,PGDADSU,ADRESSE
*****************************************************************}
procedure AdresseNormalisee(origine : string; var numero, nom : string);
var
Adresse : string;
i, Long : integer;
begin
numero := '     ';
nom := '                                   ';
Adresse := origine;

i := 1;
Long := Length(Adresse);

While ((Adresse[i]>='0') and (Adresse[i]<='9')) do
      begin
      numero[i] := Adresse[i];
      i := i+1;
      end;

While ((Adresse[i] = ',') or (Adresse[i] = ' ')) do
      i:= i+1;

nom := Copy(Adresse, i, Long-i+1);
numero := Trim(numero);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/07/2002
Modifié le ... :   /  /
Description .. : Duplication d'un honoraire
Mots clefs ... : PAIE;PGDADSU;PGHONORAIRE
*****************************************************************}
procedure Duplic (OrdreHonor : integer);
var
THonorD : TOB;
StHonor : string;
OrdreOrig : integer;
CleDADS : TCleDADS;
begin
StHonor:= '--H';
{PT109
StType:= '001';
}

{PT101-6
CreeEntete(StHonor, StType, OrdreHonor, DebExer2, FinExer2, '', '');
}
CleDADS.Salarie:= StHonor+IntToStr(OrdreHonor);
{PT109
CleDADS.TypeD:= StType;
}
CleDADS.TypeD:= TypeD;
//FIN PT109
CleDADS.Num:= OrdreHonor;
CleDADS.DateDeb:= DebExer2;
CleDADS.DateFin:= FinExer2;
{PT119
CleDADS.Exercice:= PGExercice;
}
CleDADS.Exercice:= PGExercice2;
CreeEntete (CleDADS, '', '', Decalage);
//FIN PT101-6

THonorD := THonor2.FindFirst([''], [''], TRUE);
if THonorD <> nil then
   OrdreOrig := THonorD.GetValue('PDS_ORDRE')
else
   OrdreOrig := 0;

Duplic_un (OrdreHonor, 'S70.G01.00.001');
Duplic_un (OrdreHonor, 'S70.G01.00.002');
Duplic_un (OrdreHonor, 'S70.G01.00.002.001');
Duplic_un (OrdreHonor, 'S70.G01.00.002.002');
Duplic_un (OrdreHonor, 'S70.G01.00.003.001');
Duplic_un (OrdreHonor, 'S70.G01.00.003.002');
Duplic_un (OrdreHonor, 'S70.G01.00.004.001');
Duplic_un (OrdreHonor, 'S70.G01.00.004.003');
Duplic_un (OrdreHonor, 'S70.G01.00.004.004');
Duplic_un (OrdreHonor, 'S70.G01.00.004.006');
Duplic_un (OrdreHonor, 'S70.G01.00.004.007');
Duplic_un (OrdreHonor, 'S70.G01.00.004.009');
Duplic_un (OrdreHonor, 'S70.G01.00.004.010');
Duplic_un (OrdreHonor, 'S70.G01.00.004.012');
Duplic_un (OrdreHonor, 'S70.G01.00.004.013');
Duplic_un (OrdreHonor, 'S70.G01.00.004.014');
Duplic_un (OrdreHonor, 'S70.G01.00.004.015');
Duplic_un (OrdreHonor, 'S70.G01.00.005');
Duplic_un (OrdreHonor, 'S70.G01.00.006');
Duplic_un (OrdreHonor, 'S70.G01.00.007');
Duplic_un (OrdreHonor, 'S70.G01.00.008');
Duplic_un (OrdreHonor, 'S70.G01.00.009');
Duplic_un (OrdreHonor, 'S70.G01.00.010');
Duplic_un (OrdreHonor, 'S70.G01.00.011');
Duplic_un (OrdreHonor, 'S70.G01.00.013');
Duplic_un (OrdreHonor, 'S70.G01.00.014');
Duplic_un (OrdreHonor, 'S70.G01.00.016');

if Trace <> nil then
   Trace.Add ('                               Duplication à partir de'+
              ' l''honoraire N°'+IntToStr (OrdreOrig));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/07/2004
Modifié le ... :   /  /
Description .. : Duplication d'un segment honoraire
Mots clefs ... : PAIE;PGDADSU;PGHONORAIRE
*****************************************************************}
procedure Duplic_un (OrdreHonor : integer; Segment : string);
var
THonorD : TOB;
StHonor: string;
CleDADS : TCleDADS;
begin
StHonor := '--H';
// pt151 StType := '001';
//PT101-6
CleDADS.Salarie:= StHonor+IntToStr(OrdreHonor);
// pt151 CleDADS.TypeD:= StType;
CleDADS.TypeD:= TypeD;  // pt151
CleDADS.Num:= OrdreHonor;
CleDADS.DateDeb:= DebExer2;
CleDADS.DateFin:= FinExer2;
{PT119
CleDADS.Exercice:= PGExercice;
}
CleDADS.Exercice:= PGExercice2;
//FIN PT101-6

//Recherche de l'enregistrement concerné dans la TOB Honoraire
THonorD := THonor2.FindFirst(['PDS_SEGMENT'], [Segment], TRUE);
if THonorD <> nil then
  CreeDetail (CleDADS, THonorD.GetValue('PDS_ORDRESEG'), Segment,
              THonorD.GetValue('PDS_DONNEE'),
              THonorD.GetValue('PDS_DONNEEAFFICH'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/04/2006
Modifié le ... :   /  /
Description .. : Calcul des données entreprise et mise en tob
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CalculS20_Entreprise(ParamCalc : TParamCalc);
var
BufDest, Buffer, BufOrig, CodeIso, Libelle, NIC, NIC2, NomRue, numero : string;
SIR, SIR2 : string;
TEtabD : TOB;
CleDADS : TCleDADS;
NbreMois, PremAnnee, PremMois : word;
begin
CleDADS.Salarie:= '++'+Copy (ParamCalc.Libelle, 0, 15);
CleDADS.TypeD:= ParamCalc.TypeD;
CleDADS.Num:= 0;
CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;
BufOrig:= ParamCalc.Siren;
ForceNumerique(BufOrig, BufDest);
SIR := Copy(BufDest, 1, 9);
{PT114-1
DeleteErreur (CleDADS.Salarie, 'S20');
}
DeleteErreur ('', 'S20');
//FIN PT114-1
CreeDetail (CleDADS, 1, 'S20.G01.00.001', SIR, SIR);

CreeDetail (CleDADS, 2, 'S20.G01.00.002', ParamCalc.Libelle, ParamCalc.Libelle);

Buffer:= Copy(ParamCalc.TypeDecla,1,2);
BufOrig:= ParamCalc.Du;
//PT105
if ((Copy (BufOrig, 1, 5)='01/12') and
   (Copy (ParamCalc.Caracteristique, 1, 1)='A')) then
   BufOrig:= '01/01/'+ IntToStr (StrToInt (Copy (BufOrig, 7, 4))+1);
//FIN PT105
BufOrig:= DateToStr (DebutDeMois (StrToDate (BufOrig)));
ForceNumerique(BufOrig, BufDest);
CreeDetail (CleDADS, 3, 'S20.G01.00.003.001', BufDest, BufDest);

//Calcul du nombre de mois
{PT105
NombreMois (StrToDate (ParamCalc.Du), StrToDate (ParamCalc.Au), PremMois,
            PremAnnee, NbreMois);
if (NbreMois>12) then
   NbreMois:= 12;
BufOrig:= ParamCalc.Au;
BufOrig:= DateToStr (FinDeMois (StrToDate (BufOrig)));
}
{PT117
NombreMois (StrToDate (BufOrig), StrToDate (ParamCalc.Au), PremMois,
            PremAnnee, NbreMois);
if (NbreMois>12) then
   begin
   NbreMois:= 12;
   BufOrig:= DateToStr (FinDeMois (PlusMois (StrToDate (BufOrig), 12)));
   end
else
   begin
   BufOrig:= ParamCalc.Au;
   BufOrig:= DateToStr (FinDeMois (StrToDate (BufOrig)));
   end;
}
NombreMois (StrToDate (ParamCalc.Du), StrToDate (ParamCalc.Au), PremMois,
            PremAnnee, NbreMois);
if (NbreMois>12) then
   NbreMois:= 12;
BufOrig:= DateToStr (FinDeMois (PlusMois (StrToDate (BufOrig)-1, NbreMois)));
//FIN PT117
//FIN PT105
ForceNumerique (BufOrig, BufDest);
CreeDetail (CleDADS, 4, 'S20.G01.00.003.002', BufDest, BufDest);

CreeDetail (CleDADS, 5, 'S20.G01.00.004.001', Copy(ParamCalc.TypeDecla,1,2),
            ParamCalc.TypeDecla);

if (ParamCalc.Neant=True) then
   CreeDetail (CleDADS, 6, 'S20.G01.00.004.002', '55', '55')
else
   CreeDetail (CleDADS, 6, 'S20.G01.00.004.002', Copy(ParamCalc.TypeDecla,3,2),
               ParamCalc.TypeDecla);

CreeDetail (CleDADS, 7, 'S20.G01.00.005',
            ParamCalc.Fraction+ParamCalc.FractionTot,
            ParamCalc.Fraction+ParamCalc.FractionTot);

{PT113
if (GetParamSocSecur('SO_TENUEEURO', 'X') = False) then
}
if (GetParamSocSecur('SO_PGTENUEEURO', 'X') = False) then
   CreeDetail (CleDADS, 9, 'S20.G01.00.007', 'FRF', 'FRF')
else
   CreeDetail (CleDADS, 9, 'S20.G01.00.007', 'EUR', 'EUR');

TEtabD:= ParamCalc.TobEtab.FindFirst (['ET_ETABLISSEMENT'],
                                      [ParamCalc.Siege], TRUE);
if TEtabD <> nil then
   begin
   BufOrig := TEtabD.GetValue('ET_SIRET');
   ForceNumerique(BufOrig, BufDest);
   NIC := Copy(BufDest, 10, 5);
   CreeDetail (CleDADS, 10, 'S20.G01.00.008', NIC, NIC);

   if (TEtabD.GetValue('ET_ADRESSE2') <> '') then
      begin
      if (TEtabD.GetValue('ET_ADRESSE3') <> '') then
         Buffer:= TEtabD.GetValue('ET_ADRESSE2')+' '+
                  TEtabD.GetValue('ET_ADRESSE3')
      else
         Buffer := TEtabD.GetValue('ET_ADRESSE2');
      end
   else
      if (TEtabD.GetValue('ET_ADRESSE3') <> '') then
         Buffer := TEtabD.GetValue('ET_ADRESSE3');
   if (Buffer <> '') then
      CreeDetail (CleDADS, 11, 'S20.G01.00.009.001', Buffer, Buffer);

   if (TEtabD.GetValue('ET_ADRESSE1') <> '') then
      AdresseNormalisee (TEtabD.GetValue('ET_ADRESSE1'), numero, NomRue);

   if (numero <> '') then
      CreeDetail (CleDADS, 12, 'S20.G01.00.009.003', numero,
                  TEtabD.GetValue('ET_ADRESSE1'));

   if (NomRue <> '') then
      CreeDetail (CleDADS, 13, 'S20.G01.00.009.006', NomRue,
                  TEtabD.GetValue('ET_ADRESSE2'));

   if ((TEtabD.GetValue('ET_PAYS') <> 'FRA') or
      ((TEtabD.GetValue('ET_CODEPOSTAL') > '00000') and
      (TEtabD.GetValue('ET_CODEPOSTAL') < '99999'))) then
      CreeDetail (CleDADS, 14, 'S20.G01.00.009.010',
                  TEtabD.GetValue('ET_CODEPOSTAL'),
                  TEtabD.GetValue('ET_CODEPOSTAL'));

   CreeDetail (CleDADS, 15, 'S20.G01.00.009.012',
               PGUpperCase(TEtabD.GetValue('ET_VILLE')),
               TEtabD.GetValue('ET_VILLE'));

   if ((TEtabD.GetValue ('ET_PAYS')<>'FRA') and
//PT125-1
      (TEtabD.GetValue ('ET_PAYS')<>'GUF') and
      (TEtabD.GetValue ('ET_PAYS')<>'GLP') and
      (TEtabD.GetValue ('ET_PAYS')<>'MCO') and
      (TEtabD.GetValue ('ET_PAYS')<>'MTQ') and
      (TEtabD.GetValue ('ET_PAYS')<>'NCL') and
      (TEtabD.GetValue ('ET_PAYS')<>'PYF') and
      (TEtabD.GetValue ('ET_PAYS')<>'SPM') and
      (TEtabD.GetValue ('ET_PAYS')<>'REU') and
      (TEtabD.GetValue ('ET_PAYS')<>'ATF') and
      (TEtabD.GetValue ('ET_PAYS')<>'WLF') and
      (TEtabD.GetValue ('ET_PAYS')<>'MYT') and
//FIN PT125-1
      (TEtabD.GetValue ('ET_PAYS')<>'')) then
      begin
      PaysISOLib(TEtabD.GetValue('ET_PAYS'), CodeIso, Libelle);
      CreeDetail (CleDADS, 16, 'S20.G01.00.009.013', CodeIso, CodeIso);
      CreeDetail (CleDADS, 17, 'S20.G01.00.009.014', Libelle, Libelle);
      end;
   end;

if ParamCalc.Depose <> ParamCalc.Siege then
   begin
   TEtabD:= ParamCalc.TobEtab.FindFirst (['ET_ETABLISSEMENT'],
                                         [ParamCalc.Depose], TRUE);
   if TEtabD <> nil then
      begin
      BufOrig := TEtabD.GetValue('ET_SIRET');
      ForceNumerique(BufOrig, BufDest);

      NIC2 := Copy(BufDest, 10, 5);
      CreeDetail (CleDADS, 18, 'S20.G01.00.010', NIC2, NIC2);

      CreeDetail (CleDADS, 19, 'S20.G01.00.011', TEtabD.GetValue('ET_LIBELLE'),
                  TEtabD.GetValue('ET_LIBELLE'));

      if (TEtabD.GetValue('ET_ADRESSE2') <> '') then
         begin
         if (TEtabD.GetValue('ET_ADRESSE3') <> '') then
            Buffer:= TEtabD.GetValue('ET_ADRESSE2')+' '+
                     TEtabD.GetValue('ET_ADRESSE3')
         else
            Buffer := TEtabD.GetValue('ET_ADRESSE2');
         end
      else
         if (TEtabD.GetValue('ET_ADRESSE3') <> '') then
            Buffer := TEtabD.GetValue('ET_ADRESSE3');
      if (Buffer <> '') then
         CreeDetail (CleDADS, 20, 'S20.G01.00.012.001', Buffer, Buffer);

      if (TEtabD.GetValue('ET_ADRESSE1') <> '') then
         AdresseNormalisee (TEtabD.GetValue('ET_ADRESSE1'), numero, NomRue);

      if (numero <> '') then
         CreeDetail (CleDADS, 21, 'S20.G01.00.012.003', numero,
                     TEtabD.GetValue('ET_ADRESSE1'));

      if (NomRue <> '') then
         CreeDetail (CleDADS, 22, 'S20.G01.00.012.006', NomRue,
                     TEtabD.GetValue('ET_ADRESSE2'));

      CreeDetail (CleDADS, 23, 'S20.G01.00.012.010',
                  TEtabD.GetValue('ET_CODEPOSTAL'),
                  TEtabD.GetValue('ET_CODEPOSTAL'));

      CreeDetail (CleDADS, 24, 'S20.G01.00.012.012',
                  PGUpperCase(TEtabD.GetValue('ET_VILLE')),
                  TEtabD.GetValue('ET_VILLE'));

      if ((TEtabD.GetValue ('ET_PAYS')<>'FRA') and
//PT125-1
         (TEtabD.GetValue ('ET_PAYS')<>'GUF') and
         (TEtabD.GetValue ('ET_PAYS')<>'GLP') and
         (TEtabD.GetValue ('ET_PAYS')<>'MCO') and
         (TEtabD.GetValue ('ET_PAYS')<>'MTQ') and
         (TEtabD.GetValue ('ET_PAYS')<>'NCL') and
         (TEtabD.GetValue ('ET_PAYS')<>'PYF') and
         (TEtabD.GetValue ('ET_PAYS')<>'SPM') and
         (TEtabD.GetValue ('ET_PAYS')<>'REU') and
         (TEtabD.GetValue ('ET_PAYS')<>'ATF') and
         (TEtabD.GetValue ('ET_PAYS')<>'WLF') and
         (TEtabD.GetValue ('ET_PAYS')<>'MYT') and
//FIN PT125-1
         (TEtabD.GetValue ('ET_PAYS') <> '')) then
         begin
         PaysISOLib(TEtabD.GetValue('ET_PAYS'), CodeIso, Libelle);
         CreeDetail (CleDADS, 25, 'S20.G01.00.012.013', CodeIso, CodeIso);
         CreeDetail (CleDADS, 26, 'S20.G01.00.012.014', Libelle, Libelle);
         end;
      end;
   end;

if (ParamCalc.Client <> '') then
{PT125-3
   CreeDetail (CleDADS, 27, 'S20.G01.00.013', ParamCalc.Client,
               ParamCalc.Client);
}
   CreeDetail (CleDADS, 27, 'S20.G01.00.013.001', ParamCalc.Client,
               ParamCalc.Client);
//FIN PT125-3

{PT125-3
if ((ParamCalc.Siret <> '') and (ParamCalc.Communication <> '') and
   ((ParamCalc.Coordonnee <> '') or ((ParamCalc.Civilite <> '') and
   (ParamCalc.Personne <> '')))) then
}
if ((ParamCalc.Siret<>'') and (ParamCalc.Coordonnee<>'')) then
   begin
   ForceNumerique(ParamCalc.Siret, BufDest);
   SIR2 := Copy (BufDest, 1, 9);
   NIC2 := Copy(BufDest, 10, 5);
{PT125-3
   if ((SIR2<>'') and (NIC2<>'') and (ParamCalc.Communication<>'')) then
}
   if ((SIR2<>'') and (NIC2<>'')) then
//FIN PT125-3
      begin
      CreeDetail (CleDADS, 28, 'S20.G01.00.014.001', SIR2, SIR2);
      CreeDetail (CleDADS, 29, 'S20.G01.00.014.002', NIC2, NIC2);

{PT125-3
      CreeDetail (CleDADS, 30, 'S20.G01.00.015', ParamCalc.Communication,
                  ParamCalc.Communication);

      if (ParamCalc.Communication='03') then
}
      CreeDetail (CleDADS, 31, 'S20.G01.00.016.001', ParamCalc.Coordonnee,
                  ParamCalc.Coordonnee)
{PT125-3
      else
      if (ParamCalc.Communication='05') then
         begin
         if (ParamCalc.Civilite='MR') then
            Buffer:='01'
         else
         if (ParamCalc.Civilite='MME') then
            Buffer:='02'
         else
         if (ParamCalc.Civilite='MLE') then
            Buffer:='03';

         CreeDetail (CleDADS, 32, 'S20.G01.00.016.002', Buffer,
                     ParamCalc.Civilite);

         CreeDetail (CleDADS, 33, 'S20.G01.00.016.003', ParamCalc.Personne,
                     ParamCalc.Personne);
         end;
}         
      end;
   end;

if ((ParamCalc.Routage <> '') and (ParamCalc.Routage1 <> '')) then
   begin
   CreeDetail (CleDADS, 34, 'S20.G01.00.017.001', ParamCalc.Routage,
               ParamCalc.Routage);
   CreeDetail (CleDADS, 35, 'S20.G01.00.017.002', ParamCalc.Routage1,
               ParamCalc.Routage1);
   if (ParamCalc.Routage2 <> '') then
      CreeDetail (CleDADS, 36, 'S20.G01.00.017.003', ParamCalc.Routage2,
                  ParamCalc.Routage2);
   end;

CreeDetail (CleDADS, 37, 'S20.G01.00.018', ParamCalc.Caracteristique,
            ParamCalc.Caracteristique);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/07/2006
Modifié le ... :   /  /
Description .. : Calcul des données établissement et mise en tob
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure CalculS80_Etablissement(Etab : TypeEtab);
var
CleDADS : TCleDADS;
Buffer, NICEtab, SirEtab, StExiste : string;
begin
CleDADS.Salarie:= Etab.Code;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= 1;

CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
CleDADS.Exercice:= PGExercice;

{PT109
DeleteDetail (CleDADS.Salarie, CleDADS.TypeD, -100);
}
{PT125-6
DeleteDetail (CleDADS.Salarie, -100);
}
DeleteDetail (CleDADS.Salarie, CleDADS.Num);
//FIN PT125-6
//FIN PT109

SirEtab:= Copy(Etab.Siret, 1, 9);
NICEtab:= Copy(Etab.Siret, 10, 5);

CreeDetail (CleDADS, 0, 'S80.G01.00.001.001', SirEtab, Etab.Siret);
CreeDetail (CleDADS, 1, 'S80.G01.00.001.002', NICEtab, Etab.Siret);
CreeDetail (CleDADS, 2, 'S80.G01.00.002', Etab.Libelle, Etab.Libelle);

Buffer:= '';
if (Etab.Adr_Adresse2 <> '') then
   begin
   if (Etab.Adr_Adresse3 <> '') then
      Buffer:= Etab.Adr_Adresse2+' '+Etab.Adr_Adresse3
   else
      Buffer:= Etab.Adr_Adresse2;
   end
else
if (Etab.Adr_Adresse3 <> '') then
   Buffer:= Etab.Adr_Adresse3;

if (Buffer <> '') then
   CreeDetail (CleDADS, 3, 'S80.G01.00.003.001', Buffer, Buffer);

if (Etab.Adr_Adresse1 <> '') then
    AdresseNormalisee (Etab.Adr_Adresse1, Etab.Adr_Numero, Etab.Adr_NomRue);

if (Etab.Adr_Numero <> '') then
   CreeDetail (CleDADS, 4, 'S80.G01.00.003.003', Etab.Adr_Numero,
               Etab.Adr_Adresse1);

if (Etab.Adr_NomRue <> '') then
   CreeDetail (CleDADS, 5, 'S80.G01.00.003.006', Etab.Adr_NomRue,
               Etab.Adr_Adresse2);

if (Etab.Adr_CP <> '') then
   begin
   if ((Etab.Adr_PaysCode <> 'FRA') or ((Etab.Adr_CP > '00000') and
      (Etab.Adr_CP < '99999'))) then
      CreeDetail (CleDADS, 6, 'S80.G01.00.003.010', Etab.Adr_CP, Etab.Adr_CP);
   end;

if (Etab.Adr_Ville <> '') then
   CreeDetail (CleDADS, 7, 'S80.G01.00.003.012', PGUpperCase(Etab.Adr_Ville),
               Etab.Adr_Ville);

{PT125-1
if ((Etab.Adr_PaysCode <> 'FRA') and (Etab.Adr_PaysCode <> '')) then
}
if ((Etab.Adr_PaysCode<>'FRA') and (Etab.Adr_PaysCode<>'GUF') and
   (Etab.Adr_PaysCode<>'GLP') and (Etab.Adr_PaysCode<>'MCO') and
   (Etab.Adr_PaysCode<>'MTQ') and (Etab.Adr_PaysCode<>'NCL') and
   (Etab.Adr_PaysCode<>'PYF') and (Etab.Adr_PaysCode<>'SPM') and
   (Etab.Adr_PaysCode<>'REU') and (Etab.Adr_PaysCode<>'ATF') and
   (Etab.Adr_PaysCode<>'WLF') and (Etab.Adr_PaysCode<>'MYT') and
   (Etab.Adr_PaysCode<>'')) then
//FIN PT125-1
   begin
{PT114-2
   PaysISOLib (Etab.Adr_PaysCode, CodeIso, Libelle);
   CreeDetail (CleDADS, 8, 'S80.G01.00.003.013', CodeIso, CodeIso);
   CreeDetail (CleDADS, 9, 'S80.G01.00.003.014', Libelle, Libelle);
}
   CreeDetail (CleDADS, 8, 'S80.G01.00.003.013', Etab.Adr_PaysCode,
               Etab.Adr_PaysCode);
   CreeDetail (CleDADS, 9, 'S80.G01.00.003.014', Etab.Adr_PaysNom,
               Etab.Adr_PaysNom);
//FIN PT114-2
   end;

CreeDetail (CleDADS, 10, 'S80.G01.00.004.001', Etab.Effectif, Etab.Effectif);

if (Etab.SansSal<>'') then
   CreeDetail (CleDADS, 11, 'S80.G01.00.004.002', Etab.SansSal, Etab.SansSal);

CreeDetail (CleDADS, 12, 'S80.G01.00.005', Etab.TaxeSal, Etab.TaxeSal);
CreeDetail (CleDADS, 13, 'S80.G01.00.006', Etab.NAF, Etab.NAF);
{PT125-6
CreeDetail (CleDADS, 14, 'S80.G01.00.007', Etab.Prudh, Etab.Prudh);
}
CreeDetail (CleDADS, 14, 'S80.G01.00.007.001', Etab.PrudhSect, Etab.PrudhSect);
if ((Etab.PrudhDero<>'') and (Etab.PrudhDero<>null)) then
   CreeDetail (CleDADS, 15, 'S80.G01.00.007.002', Etab.PrudhDero, Etab.PrudhDero);
//FIN PT125-6

//Taxe d'apprentissage et Formation professionnelle continue
//Vérification d'une saisie existante
StExiste:= 'SELECT * FROM DADSDETAIL WHERE'+
           ' PDS_SALARIE="'+Etab.Code+'" AND'+
           ' PDS_TYPE="'+CleDADS.TypeD+'" AND'+
           ' PDS_ORDRE=2 AND'+
           ' PDS_DATEDEBUT="'+UsDateTime (CleDADS.DateDeb)+'" AND'+
           ' PDS_DATEFIN="'+UsDateTime (CleDADS.DateFin)+'" AND'+
           ' PDS_EXERCICEDADS="'+CleDADS.Exercice+'"';
//Si elle n'existe pas alors, valeur par défaut : établissement non assujetti
if (ExisteSQL (StExiste)=False) then
   begin
   CleDADS.Num:= 2;             //PT130-2
   CreeDetail (CleDADS, 71, 'S80.G62.05.001', '02', '02');
   CreeDetail (CleDADS, 81, 'S80.G62.10.001', '02', '02');
   end;
end;

//PT107
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Initialisation des données établissement et mise en tob
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure InitEtab (var Etab : TypeEtab);
begin
Etab.Code:= '';
Etab.Siret:= '';
Etab.Libelle:= '';
Etab.Adr_Adresse1:= '';
Etab.Adr_Adresse2:= '';
Etab.Adr_Adresse3:= '';
Etab.Adr_Complement:= '';
Etab.Adr_Numero:= '';
Etab.Adr_NomRue:= '';
Etab.Adr_CP:= '';
Etab.Adr_Ville:= '';
Etab.Adr_PaysCode:= '';
Etab.Adr_PaysNom:= '';
Etab.Effectif:= '';
Etab.SansSal:= '';
Etab.TaxeSal:= '';
Etab.NAF:= '';
{PT125-6
Etab.Prudh:= '';
}
Etab.PrudhSect:= '';
Etab.PrudhDero:= '';
//FIN PT125-6
end;
//FIN PT107

//PT116
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. : Suppression des enregistrements ayant le champs
Suite ........ : exercicedads mal alimenté
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure MajExercice ();
var
St : string;
begin
if (GetParamSoc('SO_PGDECALAGE')) then
   begin
   if (ExisteSQL ('SELECT * FROM DADSPERIODES WHERE'+
                  ' PDE_TYPE="001" AND'+
                  ' PDE_DATEDEBUT>="'+UsDateTime (EncodeDate (2005, 12, 01))+'" AND'+
                  ' PDE_DATEFIN>="'+UsDateTime (EncodeDate (2005, 12, 01))+'" AND'+
                  ' (PDE_EXERCICEDADS="2005" OR PDE_EXERCICEDADS="")')) then
      begin
      st:= 'update dadsperiodes set'+
           ' pde_exercicedads="2006" where'+
           ' pde_datedebut>="'+UsDateTime (EncodeDate (2005, 12, 01))+'" and'+
           ' pde_type="001" and'+
           ' pde_exercicedads="2005"';
      ExecuteSQL (St);
      st:= 'update dadsdetail set'+
           ' pds_exercicedads="2006" where'+
           ' pds_datedebut>="'+UsDateTime (EncodeDate (2005, 12, 01))+'" and'+
           ' pds_type="001" and'+
           ' pds_exercicedads="2005"';
      ExecuteSQL(St);
    end;
   end
else
   begin
   if (ExisteSQL ('SELECT * FROM DADSPERIODES WHERE'+
                  ' PDE_TYPE="001" AND'+
                  ' PDE_DATEDEBUT>="'+UsDateTime (EncodeDate (2006, 01, 01))+'" AND'+
                  ' PDE_DATEFIN>="'+UsDateTime (EncodeDate (2006, 01, 01))+'" AND'+
                  ' (PDE_EXERCICEDADS="2005" OR PDE_EXERCICEDADS="")')) then
      begin
      st:= 'update dadsperiodes set'+
           ' pde_exercicedads="2006" where'+
           ' pde_datedebut>="'+UsDateTime (EncodeDate (2006, 01, 01))+'" and'+
           ' pde_type="001" and'+
           ' pde_exercicedads="2005")';
      ExecuteSQL(St);
      st:= 'update dadsdetail set'+
           ' pds_exercicedads="2006" where'+
           ' pds_datedebut>="'+UsDateTime (EncodeDate (2006, 01, 01))+'" and'+
           ' pds_type="001" and'+
           ' pds_exercicedads="2005")';
      ExecuteSQL(St);
    end;
   end;
end;
//FIN PT116

//PT130-1
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/11/2007
Modifié le ... :   /  /
Description .. : Chargement dans une tob passée en paramètres des
Suite ........ : établissements dont ET_FICTIF<>"X"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure ChargeEtabNonFictif (TEtabDADS : TOB);
var
EtabSelect : string;
QRechEtab : TQuery;
LesEtablis,Etabl,Habilitation,StEtab : string;
k:integer;
begin
EtabSelect:= 'SELECT ET_ETABLISSEMENT'+
             ' FROM ETABLISS WHERE'+
             ' ET_FICTIF<>"-"';
if (ExisteSQL (EtabSelect)) or ((Assigned(MonHabilitation)) and
   (MonHabilitation.LeSQL<>'')) then  //PT132
   begin
   //DEB PT132
   StEtab := '';
   if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
   begin
      LesEtablis := MonHabilitation.LesEtab;
      Etabl := ReadTokenSt(LesEtablis);
      Habilitation := '';
      k := 0;
      while Etabl <> '' do
      begin
        k := k + 1;
        if Etabl <> '' then
        begin
          if k > 1 then Habilitation := Habilitation + ',';
          Habilitation := Habilitation + '"' + Etabl + '"';
        end;
        Etabl := ReadTokenSt(LesEtablis);
      end;
      if k > 0 then
        StEtab:= ' AND ET_ETABLISSEMENT IN ('+Habilitation+MonHabilitation.SqlPop+')';
   end;
   //FIN PT132

   EtabSelect:= 'SELECT ET_ETABLISSEMENT'+
                ' FROM ETABLISS WHERE'+
                ' ET_FICTIF<>"X"' + StEtab;
   QRechEtab:= OpenSql (EtabSelect, TRUE);
   TEtabDADS.LoadDetailDB ('Mes etab', '', '', QRechEtab, False);
   Ferme (QRechEtab);
   end;
end;
//FIN PT130-1

end.

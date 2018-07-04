{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/06/2001
Modifié le ... : 18/02/2002
Description .. : Unité de gestion de la saisie complémentaire de la DADS-U
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1-1  : 03/01/2002 VG V571 Suppression du tri sur les dates de début et de fin
                            qui n'était pas correct lors du calcul
PT1-2  : 03/01/2002 VG V571 Ajout d'un test sur la cohérence des dates de début
                            et de fin de période
PT1-3  : 03/01/2002 VG V571 Fermeture du fichier de contrôle pour pouvoir le
                            visualiser
PT2    : 07/01/2002 VG V571 Si les indemnités intempéries ET les indemnités CP
                            étaient remplies, plantage lors de la validation car
                            clé dupliquée
PT3    : 10/01/2002 VG V571 Correction validation des enregistrements retraite
PT4    : 14/01/2002 VG V571 Modification, dans le cahier des charges 6.0 de
                            l'alimentation de la rubrique S41.G01.00.013
                            (Condemploi)
PT5    : 21/01/2002 VG V571 on ne pouvait pas saisir de lettres pour le code
                            risque AT
PT6    : 25/01/2002 VG V571 Lors de la création d'une nouvelle période, il ne
                            faut pas effacer les éléments de la fiche salarié
PT7    : 28/01/2002 VG V571 Correction de l'alimentation pour les salariés
                            sortis le premier jour d'une période de paye
PT8    : 01/02/2002 VG V571 Les revenus d'activités sont désormais obligatoires;
                            si la valeur est à zéro, on ne supprime plus
                            l'enregistrement
PT9    : 04/02/2002 VG V571 En cas de suppression des enregistrements de la
                            dernière période existante, on supprime aussi les
                            enregistrements entête salarié
PT10   : 18/02/2002 VG V571 Après avoir fait une édition d'une période
                            d'activité, on est en version de démonstration
PT11-1 : 19/03/2002 VG V571 Le taux à temps partiel n'était pas multiplié par
                            100
PT11-2 : 19/03/2002 VG V571 Gestion de la DADSU BTP
PT12   : 22/04/2002 VG V571 Si le taux à temps partiel n'était pas renseigné,
                            plantage au moment de la validation de la saisie
                            complémentaire
PT13   : 16/05/2002 VG V582 Pour les femmes mariées, lors de la validation sur
                            la saisie complémentaire, le nom marital reprenait
                            le nom de jeune fille
PT14   : 04/06/2002 VG V582 En saisie complémentaire DADS-U, la création d'une
                            période pour un salarié pour qui le calcul n'a pas
                            été fait récupérait le N°SS sur 15 caractères au
                            lieu de 13
PT15-1 : 20/06/2002 VG V585 Requête non fermée
PT15-2 : 20/06/2002 VG V585 Affichage du libellé des sommes isolées
PT15-3 : 20/06/2002 VG V585 Le libellé des exonérations était toujours celui de
                            la première ligne
PT16   : 15/07/2002 VG V585 Adaptation cahier des charges V6R02
PT17   : 20/08/2002 VG V585 Les informations salarié sont désormais visible en
                            saisie de période d'activité
PT18   : 02/09/2002 VG V585 Ajout d'un traitement de mise à jour des éléments
                            salarié
PT19   : 03/09/2002 VG V585 Prise en compte des salariés relevant du régime SS
                            "Autre" . FQ N°10194
PT20   : 07/10/2002 VG V585 On ne crèe l'enregistrement Taux temps partiel que
                            si la condition d'emploi est différente de 'Complet'
                            FQ N° 10263
PT21-1 : 21/10/2002 VG V585 Gestion du journal des evenements
PT21-2 : 21/10/2002 VG V585 Amélioration du traitement de création de période
PT21-3 : 21/10/2002 VG V585 Pour valider la période, l'établissement est
                            obligatoire
PT21-4 : 21/10/2002 VG V585 Puisque le coefficient est forcé, on ne rend plus la
                            saisie obligatoire
PT22   : 30/10/2002 VG V585 Ecriture des détails des frais professionnels que
                            s'il existe un montant pour frais professionnels
PT23-1 : 05/11/2002 VG V585 Affichage du nom marital dans le titre de la fenêtre
                            FQ N°10293
PT23-2 : 05/11/2002 VG V585 La suppression du code dans les grilles "Régime
                            complémentaire", "Exonération", "Somme isolée",
                            "Base brute spécifique", "Base plafonnée spécifique"
                            provoque désormais l'effacement de toute la ligne
                            FQ N°10288
PT24-1 : 14/11/2002 VG V585 On force le motif de fin de période si le motif de
                            début de période est à '095'
PT24-2 : 14/11/2002 VG V585 Contrôles supplémentaires sur le taux de travail à
                            temps partiel
PT25   : 11/12/2002 VG V591 Taux de travail à temps partiel : On ne doit plus
                            enlever la virgule. Des contrôles sont déjà
                            effectués lors de la génération du fichier.
PT26   : 14/01/2003 VG V591 Reprise de la FQ N°10288 : S'il existe un code
                            organisme, on ne demande plus la saisie d'un
                            organisme lors de la validation
PT27   : 12/02/2003 VG V_42 Si la donnée saisie ne se trouve pas dans la combo,
                            on force à la valeur par défaut - FQ N°10470
PT28   : 14/02/2003 VG V_42 Passage des zones de saisie de type date au format
                            date
PT29-1 : 20/02/2003 VG V_42 Libération de TOB non faite mise en évidence avec
                            memcheck
PT29-2 : 20/02/2003 VG V_42 En saisie complémentaire, la commune de naissance
                            était écrite même si non connue
PT30   : 03/03/2003 VG V_42 Ajout d'un paramètre à la fonction RemplitPeriode et
                            mise en tob des dates au format "DATE"
PT31-1 : 13/03/2003 VG V_42 Gestion du journal des evenements
PT31-2 : 18/03/2003 VG V_42 Pour une somme isolée, le nombre d'heures est forcé
                            à 0 . On interdit la saisie du motif de fin
                            "Somme isolée" si le motif de début n'est pas
                            "Somme isolée" - FQ N°10576
PT31-3 : 18/03/2003 VG V_42 Pour les apprentis, S41.G01.06.002 = 0 et
                            S41.G01.06.003 absent lors du calcul (en fonction du
                            contrat de travail) - FQ N°10577
PT31-4 : 19/03/2003 VG V_42 Les dates de debut et de fin passent de type otDate
PT31-5 : 19/03/2003 VG V_42 Contrôle que la date de fin ne soit pas inférieure à
                            la date de début et que les dates de début et de fin
                            ne soient pas à l'exterieur de l'exercice
PT31-6 : 19/03/2003 VG V_42 A la validation, Si aucun régime complémentaire
                            n'est renseigné, on force celui-ci à '90000' au lieu
                            de mettre un message
PT31-7 : 25/03/2003 VG V_42 Ecriture des détails de l'avantage en nature que
                            s'il existe un montant d'avantage en nature
PT32   : 28/03/2003 VG V_42 Concordance AT avec régime SS - FQ N°10574
PT33   : 31/03/2003 VG V_42 Correction sur le code type horaire BTP
PT34   : 02/04/2003 VG V_42 Correction pour la suppression des éléments DADS-U
                            BTP
PT35-1 : 23/04/2003 VG V_42 DADS-U BTP : le salaire moyen est conditionné
PT35-2 : 23/04/2003 VG V_42 DADS-U BTP : Modification pour pouvoir supprimer une
                            période d'inactivité
PT35-3 : 23/04/2003 VG V_42 DADS-U BTP : Dans les périodes d'inactivités, les
                            dates de début et de fin sont forcées si non
                            renseignées
PT35-4 : 23/04/2003 VG V_42 DADS-U BTP : Dans les périodes d'inactivités, le
                            nombre d'heure doit être renseigné
PT36   : 12/08/2003 VG V_42 Quand le salarié est en "Contrat d'Apprentissage
                            Sect. Publ.", on force le contrat à "Contrat
                            Apprentissage entr. +10 sal" - FQ N°10698
PT37   : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT37-1 : 16/09/2003 VG V_42 Traitement du cas des ouvriers non qualifiés
PT38   : 03/10/2003 VG V_42 Adaptation nouveau code PCS
PT39   : 06/10/2003 VG V_42 DADS-U BTP et cahier des charges V7R01 - FQ N°10878
PT40   : 17/10/2003 VG V_42 Reprise automatique du N° d'affiliation de
                            l'organisme si un organisme est modifié- FQ N°10907
PT41   : 24/10/2003 VG V_42 Le nom du fichier .log n'était pas initialisé
PT42   : 12/12/2003 VG V_50 En saisie complémentaire de la DADS-U, la zone
                            département de naissance est systématiquement mise à
                            "99"
PT43   : 13/01/2004 VG V_50 Adaptation cahier des charges V7R01
PT44   : 10/02/2004 VG V_50 Le nom du fichier de log n'était pas initialisé
                            FQ N°11100
PT45   : 25/02/2004 VG V_50 Modification du message lors du calcul - FQ N°11038
PT46   : 27/02/2004 VG V_50 Le code risque AT doit être en majuscule
                            FQ N°11081
PT47-1 : 05/03/2004 VG V_50 Alimentation de l'organisme de prévoyance
                            FQ N°11094
PT47-2 : 04/03/2004 VG V_50 Ajout d'un ajustement pour que les périodes soient
                            au plus juste
PT48-1 : 08/03/2004 VG V_50 Dans les grilles de saisie, on ne permet de saisir
                            que des codes valides - FQ N°11109
PT48-2 : 08/03/2004 VG V_50 Si le code section = 99,  on complète le taux AT et
                            le code risque avec les valeurs 99999 - FQ N°11118
PT49-1 : 16/03/2004 VG V_50 Correction : En création, si le code pays de
                            l'adresse n'était pas renseigné au niveau de la
                            fiche salarié, on enregistrait le code ''
PT49-2 : 16/03/2004 VG V_50 Les dates de début et de fin de périodes étaient mal
                            initialisées
PT49-3 : 16/03/2004 VG V_50 Le test du contrôle sur le code postal de l'adresse
                            était incomplet
PT50   : 31/03/2004 VG V_50 On ne pose plus la question pour la réinitialisation
                            du fichier de log. Le fichier est réinitialisé si le
                            précédent calcul datait de plus de 6 mois.
                            Le fichier est systématiquement ouvert en fin de
                            traitement
PT51-1 : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT51-2 : 05/07/2004 VG V_50 Ergonomie
PT51-3 : 05/07/2004 VG V_50 Spécificités BTP
PT51-4 : 05/07/2004 VG V_50 Contrôles
PT51-5 : 05/07/2004 VG V_50 DADS-U BTP : recupération heures salaries/employeur
                            erronnée - FQ N°11261
PT51-6 : 05/07/2004 VG V_50 Gestion des prud'hommes
PT51-7 : 05/07/2004 VG V_50 Optimisation du traitement
PT52   : 24/09/2004 VG V_50 Adaptation cahier des charges V8R01
PT53   : 04/10/2004 VG V_50 Modification de l'alimentation pour les DADS-U BTP
                            1 enregistrement par période d'activité
PT54-1 : 12/10/2004 VG V_50 Alerte lorsque "Situation familiale BTP" non
                            renseignée - FQ N°11667
PT54-2 : 12/10/2004 VG V_50 Nombre d'enfants couverts : "Inconnu" à la place de
                            "99" - FQ N°11668
PT55   : 13/10/2004 VG V_50 Alimentation conditionnée des heures rémunérées ou
                            salariées - FQ N°11694
PT56   : 21/10/2004 VG V_50 Numéro de période incrémenté de 1 entre chaque
                            période
PT57   : 02/11/2004 VG V_50 Ajout du code convention collective à 9999
PT58-1 : 30/11/2004 VG V_60 Correction récupération données prud'homales
                            FQ N°11809
PT58-2 : 30/11/2004 VG V_60 Correction de l'impression qui plantait si la date
                            d'entrée (Retraite/Prévoyance) n'était pas
                            renseignée - FQ N°11809
PT59   : 03/12/2004 VG V_60 Manipulation des périodes DADS-U en saisie
                            complémentaire : Erreur de transaction - FQ N°11824
PT60   : 06/01/2005 VG V_60 La touche F10 ne fonctionnait pas - FQ N°11779
PT61   : 12/01/2005 VG V_60 Le segment S45.G01.01.006 n'est pas obligatoire
                            FQ N°11915
PT62   : 20/01/2005 VG V_60 Impossible de supprimer une période d'activité en
                            saisie des périodes d'activités en CWAS - FQ N°11940
PT63   : 01/02/2005 VG V_60 Edition fiche DADS-U en CWAS
PT64   : 02/02/2005 VG V_60 "Travail à l'étranger ou frontalier" mal renseigné
                            FQ N°11957
PT65   : 03/02/2005 VG V_60 Contrôle des montants saisis dans les tableaux
                            FQ N°11109
PT66   : 07/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT67-1 : 11/10/2005 VG V_60 Pour un dossier en PETIT décalage paie, le segment
                            S41.G01.00.009 doit prendre la valeur 01
                            FQ N°12197
PT67-2 : 11/10/2005 VG V_60 DADS-U Congés payés BTP : la liste des institutions
                            affiche actuellement toutes les institutions sauf
                            celles du BTP - FQ N°12235
PT68   : 29/12/2005 PH V_650 Correction PT66 Boucle infinie
PT69   : 03/01/2006 PH V_651 FQ 12281 Multi-périodes DADS-U
PT70   : 04/01/2006 VG V_65 Suppression du contrôle sur les apprentis qui
                            mettait à 0 le montant des exonérations - FQ N°12783
PT71   : 13/03/2006 VG V_65 Mauvais codage du code organisme destinataire dans
                            le cas d'un organisme BTP - FQ N°12972
PT72   : 16/03/2006 VG V_65 Erreur lors de l'enregistrement des données si "Base
                            brute pour calcul C.P." ou "Base brute de
                            cotisations" sont à ''
PT73-1 : 16/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                            des erreurs
PT73-2 : 16/10/2006 VG V_70 Gestion du décalage de paie - FQ N°12860
PT73-3 : 16/10/2006 VG V_70 Modification de l'affichage des motifs de début et
                            de fin de période - FQ N°12809
PT73-4 : 16/10/2006 VG V_70 Modification du message de suppression - FQ N°12871
PT73-5 : 16/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT73-6 : 16/10/2006 VG V_70 Pour les déclarations BTP, suppression des onglets
                            inutiles - FQ N°13143
PT73-7 : 16/10/2006 VG V_70 Traitement DADS-U complémentaire
PT73-8 : 16/10/2006 VG V_70 Traitement des DADS-U IRCANTEC
PT73-9 : 16/10/2006 VG V_70 Utilisation d'un type pour la cle DADS-U
PT74   : 23/10/2006 VG V_70 Correction contrôle des apprentis - FQ N°13614
PT75   : 24/10/2006 VG V_70 Contrôle anciennetés BTP - FQ N°13304
PT76   : 31/10/2006 VG V_70 Si le montant de la base brute exonérée est
                            supérieure à la base brute SS, il n'y a pas
                            d'anomalie - FQ N°13624
PT77-1 : 03/11/2006 VG V_70 Alimentation des organismes B pour DADS-U BTP
                            uniquement et inversement, affichage onglet IRCANTEC
                            uniquement si pas DADS-U BTP - FQ N°13643
PT77-2 : 03/11/2006 VG V_70 Format d'enregistrement de "durée hebdomadaire du
                            poste de travail" - FQ N°13650
PT78   : 07/11/2006 VG V_70 Suppression des contrôles AT pour les VRP
                            multicartes - FQ N°13626
PT79-1 : 08/11/2006 VG V_70 Effacement des libellés des motifs de début et fin
                            de périodes - FQ N°13662
PT79-2 : 08/11/2006 VG V_70 Correction affichages régimes complémentaires et
                            IRCANTEC - FQ N°13660
PT80   : 09/11/2006 VG V_70 Pas d'enregistrement prud'hommaux si le contrat est
                            de droit public
PT81   : 16/11/2006 VG V_70 Adaptation cahier des charges V8R04 rectificatif du
                            15/11/2006
PT82   : 21/11/2006 VG V_70 Permettre de déclarer des honoraires en DADS-U
                            complémentaire - FQ N°13613
PT83   : 23/11/2006 VG V_70 Modification du contrôle IRCANTEC - FQ N°13650
PT84   : 01/12/2006 VG V_70 Correction PT79-2 - FQ N°13660
PT85   : 15/12/2006 VG V_70 Contrôle du numéro de rattachement IRCANTEC
                            FQ N°13768
PT86   : 18/12/2006 VG V_70 Ne pas alimenter le code de rattachement si
                            l'organisme n'est pas généré - FQ N°13643
PT87   : 09/01/2007 VG V_72 Alimentation de la base brute fiscale à 0 pour
                            certains contrats - FQ N°13822
PT88   : 11/01/2007 VG V_72 Suppression des enregistrements DADSCONTROLE lors de
                            la suppression d'une période - FQ N°13810
PT89   : 17/01/2007 VG V_72 Plantage en CWAS lorsqu'on veut modifier le motif de
                            début de période
PT90   : 26/01/2007 VG V_72 DADS-U IRCANTEC : Retraite complémentaire non
                            incluse - FQ N°13861
PT91   : 13/07/2007 VG V_72 "Condition d''emploi" remplacé par "Caractéristique
                            activité" - FQ N°14568
PT92-1 : 20/07/2007 FC V_72 FQ 14423 Gestion des différents régimes
PT92-2 : 16/08/2007 VG V_72 FQ 14423 Ajout des contrôles
PT93-1 : 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT93-2 : 20/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14430
PT93-3 : 22/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14439
PT93-4 : 03/09/2007 VG V_80 Adaptation cahier des charges V8R05
PT93-5 : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT93-6 : 20/11/2007 NA V_80 FQ 13636 Ajout des contrôles
PT94   : 21/11/2007 GGU V_80 FQ 14967 ne pas créer le segment S45.G01.01.004.002
                            code organisme gestionnaire si la valeur n'est pas
                            saisie (segment à blanc) ; (saisie sur instruction
                            institution)
PT95-1 : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT95-2 : 28/11/2007 VG V_80 Correction contrôle CNBF - FQ N°14974
PT95-3 : 28/11/2007 VG V_80 Correction contrôle IRCANTEC - FQ N°14979
PT96   : 28/11/2007 GGU V_80 FQ14967 Date de début / fin de période incorrecte
                             si le salarié rentre en cours de période.
PT97   : 29/11/2007 NA V_80 FQ14976 Base brute/ base plafonnée exonération: ne
                            pas créer le segment si = 0
PT98   : 30/11/2007 VG V_80 Désactivation du bouton suivant si pas de période
                            suivante - FQ N°14134
PT99   : 04/12/2007 NA V-80 FQ N° 13769 et 14033: Gestion des bases
                            exceptionnelles et exo sans les décimales
PT100  : 05/12/2007 VG V_80 Affichage du caractère interdit - FQ N°14961
PT101  : 06/12/2007 NA V_80 FQ N°14421 Gestion des codes autorisés pour les
                            régimes maladie/AT ...
PT102  : 06/12/2007 VG V_80 Message en entrée dans la fiche en CWAS - FQ N°14991
PT103  : 10/12/2007 VG V_80 Décimales acceptées pour le nombre d'heures ou de
                            jours supplémentaires - FQ N°14980
PT104  : 10/12/2007 GGU V_80 FQ 15028 - pb affichage des champs lors de la
                            validation
PT104-1: 21/12/2007 GGU V_80 FQ 15028 : la date début période est obligatoire
                            pour les codes événement 01, 02, 03, 04, 50
                            la date fin de période est obligatoire pour les
                            codes événement 02, 03, 04, 50
PT105  : 10/12/2007 VG V_80 Ajout contrôle volontariat associatif - FQ N°14759
PT106  : 09/01/2008 VG V_80 Contrôle de cohérence section prud'homale avec
                            section établissement uniquement si différente de
                            '05' - FQ N°15097
PT107-1: 11/01/2008 VG V_80 Correction de l'alimentation de la nature base SS
                            dans le cas des apprentis - FQ N°15113
PT107-2: 11/01/2008 VG V_80 Pour exonération des heures sup, si le code est à
                            "Aucun", on n'enregistre pas les montants
                            FQ N°15113
PT108  : 18/01/2008 VG V_80 Correction sommes isolées - FQ N°15130
PT111  : 15/07/2008 NA V_80 FQ 14970 Saisie prévoyance : mettre date début et date fin au format date
}
unit UTOFPG_DADSPERIODES;

interface
uses
         {$IFDEF VER150}
         Variants,
         {$ENDIF}
         UTOF,
{$IFNDEF EAGLCLIENT}
         DBCtrls,
         {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
         FE_Main,
         EdtREtat,
{$ELSE}
         UtileAGL,
         MaineAgl,
{$ENDIF}
         Hctrls,
         ComCtrls,
         HEnt1,
         HMsgBox,
         StdCtrls,
         Classes,
         sysutils,
         UTob,
         HTB97,
         Vierge,
         PgDADSOutils,
         PgDADSCommun,
         Pgoutils2,
         EntPaie,
         ParamDat,
         Controls,
         LookUp,
         ParamSoc,
{$IFNDEF DADSUSEULE}
         P5Def,
{$ENDIF}
         windows,
         PgDADSControles,
         UTofPG_DADSPrevEvt;

Type TMotif = record
     ControlMotif : TControl;    //TControl
     TitreMotif : string;        //Titre de la fenêtre
     TableMotif : string;        //Table
     ColonneMotif : string;      //Code
     SelectMotif : string;       //Libellé
     WhereMotif : string;        //Condition
     EditMotif : string;         //THEdit
     LabelMotif : string;        //TLabel
     end;

Type
      TOF_PG_DADSPERIODES = Class (TOF)
      public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad; override ;
        procedure OnClose ; override ;

      private
        Basebrutespec, Baseplafspec, Exoner, Prevoyance, Regimecomp : THGrid;
        Somiso : THGrid;

        CodeRegime, Fraction, LibSal, Nbre, NbreTot, Sal, State : string;

        Calculer, Daccord, PerPrec, PerSuiv, Salarie, SalMAJ : TToolBarButton97;
        SalPrem, SalPrec, SalSuiv, SalDern : TToolBarButton97;

        QMul, QPer : TQUERY;     // Query recuperee du mul

        T_Etab, T_Periode, T_Sal : TOB; // Tob du salarie et de son etablissement

        ControleOK, FocusRegime, SalChange : Boolean;

        DebMotif1, DebMotif2, DebMotif3, FinMotif1, FinMotif2 : THEdit;
        FinMotif3, MetierBTP : THEdit;

        CondEmploi : THValComboBox;

        AGFF, AllocForfait, Alsace, AutreAvant, AutreDepense : TCheckBox;
        BaseReelle, CodeBureau, Logement, MultiEmp, Nourriture : TCheckBox;
        NTIC, Ouvre, PriseCharge, PrudDP, PrudPresent, RembJustif : TCheckBox;
        RemPourboire, TypeRegime, Voiture : TCheckBox;

        LeMotif : TMotif;

        procedure New(Sender: TObject);
        procedure Del(Sender: TObject);
        procedure AfficheCaption();
        procedure ChargeZones ();
        procedure SauveZones();
        procedure MetABlanc();
        procedure MetABlancSal();
        function ControleConform() : boolean;
        function UpdateTable() : boolean;
        procedure PerPrecClick (Sender: TObject);
        procedure PerSuivClick (Sender: TObject);
{$IFNDEF EAGLCLIENT}
        procedure SalPremClick (Sender: TObject);
        procedure SalPrecClick (Sender: TObject);
        procedure SalSuivClick (Sender: TObject);
        procedure SalDernClick (Sender: TObject);
{$ENDIF}
        Function BougeSal(Button: TNavigateBtn) : boolean ;
        procedure GereQuerySal();
        Function BougePer(Button: TNavigateBtn) : boolean ;
        procedure GereQueryPer();
        procedure MAJQuery();
        procedure CalculerClick (Sender: TObject);
        procedure SalarieClick (Sender: TObject);
        procedure SalarieMAJClick (Sender: TObject);
        procedure Impression(Sender: TObject);
        procedure Validation(Sender: TObject);
        procedure DateDoubleClick (Sender: TObject);
        procedure GrilleElipsisClick(Sender: TObject);
        procedure MetierBTPElipsisClick(Sender: TObject);
        procedure GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
        procedure GrilleCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
        procedure GrilleEnter (Sender: TObject);
        procedure MotifChange(Sender: TObject);
        procedure MotifChangeClick(Sender: TObject);
        procedure CondEmploiChange(Sender: TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure ChangeTypeRegime(Sender: TObject); //PT92-1
        procedure PrevoyanceDblClick (Sender: TObject);
     END;

var PGDADSEtab : string;

implementation

uses Grids;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.OnArgument(stArgument: String);
var
Pages : TPageControl;
Arg, StPlus : string;
TabSheet : TTabsheet;
begin
Inherited ;
//Récupération des arguments
Arg:= stArgument;
State:= Trim(ReadTokenPipe(Arg,';'));
Sal:= Trim(ReadTokenPipe(Arg,';')) ;
TypeD:= Trim(ReadTokenPipe(Arg,';')) ;
Fraction:= Trim(ReadTokenPipe(Arg,';')) ;

TFVierge(Ecran).OnKeyDown:= FormKeyDown;

//THEdit
MetierBTP:= THEdit (GetControl ('EMETIERBTP'));
if MetierBTP <> NIL then
   MetierBTP.OnElipsisClick:= MetierBTPElipsisClick;

//Combos
{PT73-3
DebMotif1:= THValComboBox (GetControl ('CDEBMOTIF1'));
if DebMotif1 <> NIL then
   DebMotif1.OnChange:= MotifChange;
DebMotif2:= THValComboBox (GetControl ('CDEBMOTIF2'));
if DebMotif2 <> NIL then
   DebMotif2.OnChange:= MotifChange;
DebMotif3:= THValComboBox (GetControl ('CDEBMOTIF3'));
if DebMotif3 <> NIL then
   DebMotif3.OnChange:= MotifChange;
FinMotif1:= THValComboBox (GetControl ('CFINMOTIF1'));
if FinMotif1 <> NIL then
   FinMotif1.OnChange:= MotifChange;
FinMotif2:= THValComboBox (GetControl ('CFINMOTIF2'));
if FinMotif2 <> NIL then
   FinMotif2.OnChange:= MotifChange;
FinMotif3:= THValComboBox (GetControl ('CFINMOTIF3'));
if FinMotif3 <> NIL then
   FinMotif3.OnChange:= MotifChange;
}
DebMotif1:= THEdit (GetControl ('EDEBMOTIF1'));
if DebMotif1 <> NIL then
   DebMotif1.OnElipsisClick:= MotifChangeClick;
DebMotif2:= THEdit (GetControl ('EDEBMOTIF2'));
if DebMotif2 <> NIL then
   DebMotif2.OnElipsisClick:= MotifChangeClick;
DebMotif3:= THEdit (GetControl ('EDEBMOTIF3'));
if DebMotif3 <> NIL then
   DebMotif3.OnElipsisClick:= MotifChangeClick;
FinMotif1:= THEdit (GetControl ('EFINMOTIF1'));
if FinMotif1 <> NIL then
   FinMotif1.OnElipsisClick:= MotifChangeClick;
FinMotif2:= THEdit (GetControl ('EFINMOTIF2'));
if FinMotif2 <> NIL then
   FinMotif2.OnElipsisClick:= MotifChangeClick;
FinMotif3:= THEdit (GetControl ('EFINMOTIF3'));
if FinMotif3 <> NIL then
   FinMotif3.OnElipsisClick:= MotifChangeClick;
//FIN PT73-3

CondEmploi:= THValComboBox (GetControl ('CCONDEMPLOI'));
if CondEmploi <> NIL then
   CondEmploi.OnChange:= CondEmploiChange;

//CheckBox
//DEB PT92-1
TypeRegime:= TCheckBox (getcontrol ('CHTYPEREGIME'));
if (TypeRegime<>nil) then
   TypeRegime.OnClick:= ChangeTypeRegime;
//FIN PT92-1
MultiEmp:= TCheckBox (GetControl ('CHMULTIEMP'));
//PT93-5
Alsace:= TCheckBox (GetControl ('CHALSACE'));
BaseReelle:= TCheckBox (GetControl ('CHBASE'));
//FIN PT93-5
CodeBureau:= TCheckBox (GetControl ('CHCODEBUREAU'));
Nourriture:= TCheckBox (GetControl ('CHNOURRITURE'));
Logement:= TCheckBox (GetControl ('CHLOGEMENT'));
Voiture:= TCheckBox (GetControl ('CHVOITURE'));
AutreAvant:= TCheckBox (GetControl ('CHAUTREAVANT'));
RemPourboire:= TCheckBox (GetControl ('CHREMPOURBOIRE'));
AllocForfait:= TCheckBox (GetControl ('CHALLOCFORFAIT'));
RembJustif:= TCheckBox (GetControl ('CHREMBJUSTIF'));
PriseCharge:= TCheckBox (GetControl ('CHPRISECHARGE'));
AutreDepense:= TCheckBox (GetControl ('CHAUTREDEPENSE'));
NTIC:= TCheckBox (GetControl ('CHNTIC'));
PrudDP:= TCheckBox (GetControl ('CHPRUDHPUBLIC'));
PrudPresent:= TCheckBox (GetControl ('CHPRESENT3112'));
//PT93-5
AGFF:= TCheckBox (GetControl ('CHAGFF'));
Ouvre:= TCheckBox (GetControl ('CHOUVRE'));
//FIN PT93-5


//Grids
Regimecomp:= THGRID (GetControl ('GREGIMECOMP'));
Basebrutespec:= THGRID (GetControl ('GBASEBRUTESPEC'));
Baseplafspec:= THGRID (GetControl ('GBASEPLAFSPEC'));
Exoner:= THGRID (GetControl ('GEXONER'));
Somiso:= THGRID (GetControl ('GSOMISO'));
Prevoyance:= THGrid (GetControl ('GPREVOYANCE'));

SetControlText('LSALARIE_', Sal);
SetControlText('LFRACTION_', Fraction);

T_Sal:= NIL;
T_Etab:= NIL;
// Positionnement sur le premier onglet
Pages:= TPageControl(GetControl('PAGES'));
if Pages<>nil then
   Pages.ActivePageIndex:=0;

{PT73-7
if TypeD <> '001' then
}
if ((TypeD <> '001') and (TypeD <> '201')) then
   begin
   TabSheet:= TTabsheet(GetControl('TBTP'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= True;
//PT98
   TabSheet:= TTabsheet(GetControl('TEXO'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;
//FIN PT98
//PT73-6
   TabSheet:= TTabsheet(GetControl('TAUTRE'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;

   TabSheet:= TTabsheet(GetControl('TBASES'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;

   TabSheet:= TTabsheet(GetControl('TPRUD'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;

   TabSheet:= TTabsheet(GetControl('TIRCANTEC'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;

   TabSheet:= TTabsheet(GetControl('TRETPREV'));
   if TabSheet <> nil then
      TabSheet.TabVisible:= False;
//FIN PT73-6
   end;

// recuperation de la query du multicritere
QMul:= TFVierge(Ecran).FMULQ ;

// Gestion du navigateur
SalPrem:= TToolbarButton97(GetControl('BSALPREM'));
SalPrec:= TToolbarButton97(GetControl('BSALPREC'));
SalSuiv:= TToolbarButton97(GetControl('BSALSUIV'));
SalDern:= TToolbarButton97(GetControl('BSALDERN'));
{$IFNDEF EAGLCLIENT}
if SalPrem <> NIL then
   begin
   SalPrem.Enabled:= True;
   SalPrem.Visible:= True;
   SalPrem.OnClick:= SalPremClick;
   end;

if SalPrec <> NIL then
   begin
   SalPrec.Enabled:= True;
   SalPrec.Visible:= True;
   SalPrec.OnClick:= SalPrecClick;
   end;

if SalSuiv <> NIL then
   begin
   SalSuiv.Enabled:= True;
   SalSuiv.Visible:= True;
   SalSuiv.OnClick:= SalSuivClick;
   end;

if SalDern <> NIL then
   begin
   SalDern.Enabled:= True;
   SalDern.Visible:= True;
   SalDern.OnClick:= SalDernClick;
   end;
{$ELSE}
if SalPrem <> NIL then
   SalPrem.Visible:= False;

if SalPrec <> NIL then
   SalPrec.Visible:= False;

if SalSuiv <> NIL then
   SalSuiv.Visible:= False;

if SalDern <> NIL then
   SalDern.Visible:= False;
{$ENDIF}

SalChange:= True;

PerPrec:= TToolbarButton97(GetControl('BPERPREC'));
if PerPrec <> NIL then
   begin
   PerPrec.Enabled:= True;
   PerPrec.Visible:= True;
   PerPrec.OnClick:= PerPrecClick;
   end;

PerSuiv:= TToolbarButton97(GetControl('BPERSUIV'));
if PerSuiv <> NIL then
   begin
   PerSuiv.Enabled:= True;
   PerSuiv.Visible:= True;
   PerSuiv.OnClick:= PerSuivClick;
   end;

TFVierge(Ecran).Binsert.OnClick:= New;
TFVierge(Ecran).BDelete.OnClick:= Del;

Calculer:= TToolbarButton97(GetControl('BCALCULER'));
if Calculer <> NIL then
   begin
   Calculer.Enabled:= True;
   Calculer.Visible:= True;
   Calculer.OnClick:= CalculerClick;
   end;

Salarie:= TToolbarButton97(GetControl('BSALARIE'));
if Salarie <> NIL then
   Salarie.OnClick:= SalarieClick;

SalMAJ:= TToolbarButton97(GetControl('BSALMAJ'));
if SalMAJ <> NIL then
   SalMAJ.OnClick:= SalarieMAJClick;

TFVierge(Ecran).BImprimer.OnClick:= Impression;
//TFVierge(Ecran).BImprimer.Visible:= False;

Daccord:= TToolbarButton97(GetControl('BDACCORD'));
if Daccord <> NIL then
   begin
   Daccord.Enabled:= True;
   Daccord.Visible:= True;
   Daccord.OnClick:= Validation;
   end;

if RegimeComp <> NIL then
   begin
   RegimeComp.OnCellEnter:= GrilleCellEnter;
   RegimeComp.OnCellExit:= GrilleCellExit;
   RegimeComp.OnElipsisClick:= GrilleElipsisClick;
   RegimeComp.OnEnter:= GrilleEnter;              //PT84
   RegimeComp.HideSelectedWhenInactive:= true;
   RegimeComp.ColWidths [0]:=40;
   RegimeComp.ColWidths [1]:=170;
   RegimeComp.ColWidths [2]:=193;
   end;

if Basebrutespec <> NIL then
   begin
   Basebrutespec.OnCellEnter:= GrilleCellEnter;
   Basebrutespec.OnCellExit:= GrilleCellExit;
   Basebrutespec.OnElipsisClick:= GrilleElipsisClick;
   Basebrutespec.HideSelectedWhenInactive:= true;
   Basebrutespec.ColAligns[2]:= taRightJustify;
   Basebrutespec.ColWidths [0]:= 40;
   Basebrutespec.ColWidths [1]:= 470;
   Basebrutespec.ColWidths [2]:= 80;
   Basebrutespec.Coltypes[1] := 'R'; // pt99
   Basebrutespec.Coltypes[2] := 'R'; // pt99
   Basebrutespec.ColFormats[1]:= '#######0' ; // pt99
   Basebrutespec.ColFormats[2]:= '#######0' ; // pt99
   end;

if Baseplafspec <> NIL then
   begin
   Baseplafspec.OnCellEnter:= GrilleCellEnter;
   Baseplafspec.OnCellExit:= GrilleCellExit;
   Baseplafspec.OnElipsisClick:= GrilleElipsisClick;
   Baseplafspec.HideSelectedWhenInactive:= true;
   Baseplafspec.ColAligns[2]:= taRightJustify;
   Baseplafspec.ColWidths [0]:= 40;
   Baseplafspec.ColWidths [1]:= 470;
   Baseplafspec.ColWidths [2]:= 80;
   Baseplafspec.Coltypes[1] := 'R'; // pt99
   Baseplafspec.Coltypes[2] := 'R'; // pt99
   Baseplafspec.ColFormats[1]:= '#######0' ; // pt99
   Baseplafspec.ColFormats[2]:= '#######0' ; // pt99
   end;

if Exoner <> NIL then
   begin
   Exoner.OnCellEnter:= GrilleCellEnter;
   Exoner.OnCellExit:= GrilleCellExit;
   Exoner.OnElipsisClick:= GrilleElipsisClick;
   Exoner.HideSelectedWhenInactive:= true;
   Exoner.ColAligns[2]:= taRightJustify;
   Exoner.ColAligns[3]:= taRightJustify;
   Exoner.ColWidths [0]:= 40;
   Exoner.ColWidths [1]:= 390;
   Exoner.ColWidths [2]:= 80;
   Exoner.ColWidths [3]:= 80;
   Exoner.Coltypes[2] := 'R'; // pt99
   Exoner.Coltypes[3] := 'R'; // pt99
   Exoner.ColFormats[2]:= '#######0' ; // pt99
   Exoner.ColFormats[3]:= '#######0' ; // pt99
   end;

if Somiso <> NIL then
   begin
   Somiso.OnCellEnter:= GrilleCellEnter;
   Somiso.OnCellExit:= GrilleCellExit;
   Somiso.OnElipsisClick:= GrilleElipsisClick;
   Somiso.HideSelectedWhenInactive:= true;
   Somiso.ColAligns[3]:= taRightJustify;
   Somiso.ColAligns[4]:= taRightJustify;
   Somiso.ColWidths [0]:= 40;
   Somiso.ColWidths [1]:= 339;
   Somiso.ColWidths [2]:= 50;
   Somiso.ColWidths [3]:= 80;
   Somiso.ColWidths [4]:= 80;
   end;

if Prevoyance <> NIL then
   begin
   Prevoyance.OnCellEnter:= GrilleCellEnter;
   Prevoyance.OnCellExit:= GrilleCellExit;
   Prevoyance.OnElipsisClick:= GrilleElipsisClick;
   Prevoyance.OnDblClick:= PrevoyanceDblClick;    //PT102
   Prevoyance.HideSelectedWhenInactive:= true;
   Prevoyance.ColEditables[0]:= False;   //PT102
   Prevoyance.ColAligns[1]:= taCenter;
   Prevoyance.ColAligns[2]:= taCenter;
   Prevoyance.ColAligns[6]:= taRightJustify;
   Prevoyance.ColWidths[0]:= 184;
   Prevoyance.ColWidths[1]:= 60;
   Prevoyance.ColWidths[2]:= 60;
   Prevoyance.ColWidths[3]:= 135;
{PT93-5
   Prevoyance.ColWidths[4]:= 131;
   Prevoyance.ColWidths[5]:= 172;
   Prevoyance.ColWidths[6]:= 69;
   Prevoyance.ColWidths[7]:= 146;
   Prevoyance.ColWidths[8]:= 120;
   Prevoyance.ColWidths[9]:= 63;
}
   Prevoyance.ColWidths[4]:= 135;
   Prevoyance.ColWidths[5]:= 131;
   Prevoyance.ColWidths[6]:= 172;
   Prevoyance.ColWidths[7]:= 69;
   Prevoyance.ColWidths[8]:= 146;
   Prevoyance.ColWidths[9]:= 120;
   Prevoyance.ColWidths[10]:= 63;

   // deb pt111
   Prevoyance.ColTypes[1]    := 'D';
   Prevoyance.ColFormats [1] := ShortDateFormat;
   Prevoyance.ColTypes[2]    := 'D';
   Prevoyance.ColFormats [2] := ShortDateFormat;
   // fin pt111
   
//FIN PT93-5
{PT102
   Prevoyance.ColFormats[0]:= 'CB=PGDADSPREV';
}
{PT93-5
   Prevoyance.ColFormats[5]:= 'CB=PGDADSBASEPREV';
   Prevoyance.ColFormats[7]:= 'CB=PGDADSPREVPOP';
   Prevoyance.ColFormats[8]:= 'CB=PGSITUATIONFAMIL';
}
   Prevoyance.ColFormats[6]:= 'CB=PGDADSBASEPREV';
   Prevoyance.ColFormats[8]:= 'CB=PGDADSPREVPOP';
   Prevoyance.ColFormats[9]:= 'CB=PGSITUATIONFAMIL';
//FIN PT93-5
   end;

if (VH_Paie.PGPCS2003) then
   SetControlProperty ('EPCS', 'datatype', 'PGCODEPCSESE') ;

//PT95-1
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('CETABLISSEMENT', 'Plus', StPlus);
//FIN PT95-1

// Deb PT101  Initialisation des codes à afficher en fonction du type de régime
StPlus:= ' AND CO_CODE <> "150"';
SetControlProperty ('CREGIMEVIS', 'Plus', StPlus);
SetControlProperty ('CREGIMEVIP', 'Plus', StPlus);
Stplus := ' AND CO_CODE <> "120" AND CO_CODE <> "121" AND CO_CODE <> "147" AND CO_CODE <> "148" AND '+
' CO_CODE <> "149" AND CO_CODE <> "150" AND CO_CODE <> "157"';
Setcontrolproperty('CREGIMEAT', 'Plus', Stplus);
Stplus := ' AND CO_CODE <> "120" AND CO_CODE <> "121" AND CO_CODE <> "147" AND CO_CODE <> "148" AND '+
' CO_CODE <> "150" AND CO_CODE <> "157"';
Setcontrolproperty('CREGIMEMAL', 'Plus', Stplus);
Stplus := ' AND CO_CODE <> "120" AND CO_CODE <> "121" AND CO_CODE <> "147" AND CO_CODE <> "148" AND '+
' CO_CODE <> "149" AND CO_CODE <> "157"';
Setcontrolproperty('CREGIMESS', 'Plus', Stplus);
// fin pt101

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : procédure exécutée sur le Binsert.onclick
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.New(Sender: TObject);
var
IMax : Integer;
QRechNum :TQuery;
begin
UpdateTable;
Ecran.Caption:= 'Saisie complémentaire DADS-U : Salarié '+Sal+' '+LibSal+' '+
                GetControlText('EPRENOM')+' Période nouvelle';

MetABlanc;

State := 'CREATION';
IMax:=1;
QRechNum:=OpenSQL('SELECT MAX(PDE_ORDRE)'+
                  ' FROM DADSPERIODES WHERE'+
                  ' PDE_SALARIE ="'+Sal+'" AND'+
                  ' PDE_TYPE="'+TypeD+'" AND'+
                  ' PDE_ORDRE > 0 AND'+
                  ' PDE_EXERCICEDADS = "'+PGExercice+'"',TRUE);

if Not QRechNum.EOF then
   try
   IMax:= QRechNum.Fields[0].AsInteger+1;
   except
         on E: EConvertError do
            IMax:= 1;
   end;
Ferme(QRechNum);
SetControlText('ORDREDADS', IntToStr(IMax));
SetControlText('TYPEDADS', TypeD);
SetControlText('DATEDEBDADS', DateToStr(IDate1900));
SetControlText('DATEFINDADS', DateToStr(IDate1900));

// Gestion du navigateur
if SalPrem <> NIL then
   SalPrem.Enabled := False;

if SalPrec <> NIL then
   SalPrec.Enabled := False;

if SalSuiv <> NIL then
   SalSuiv.Enabled := False;

if SalDern <> NIL then
   SalDern.Enabled := False;

if PerPrec <> NIL then
   PerPrec.Enabled := False;

if PerSuiv <> NIL then
   PerSuiv.Enabled := False;

TFVierge(Ecran).Binsert.Enabled := False;
TFVierge(Ecran).BDelete.Enabled := False;

if Calculer <> NIL then
   Calculer.Enabled := False;

//SalChange := True;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : procédure exécutée sur le bdelete.onclick
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.Del(Sender: TObject);
var
Rep : integer;
StExiste : string;
begin
{PT73-4
Rep:=PGIAsk ('Voulez vous supprimer cette fiche ?', 'Suppression DADS-U');
}
Rep:=PGIAsk ('Confirmez-vous la suppression de la période en cours ?',
             'Suppression DADS-U');
if Rep=mrNo then
   exit
else
   begin
   try
      begintrans;
{PT82
      DuplicPrudh(Sal, TypeD, 'D', StrToInt(GetControlText('ORDREDADS')));
      DeletePeriode(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
}
      DuplicPrudh(Sal, 'D', StrToInt(GetControlText('ORDREDADS')));
      DeletePeriode(Sal, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, StrToInt(GetControlText('ORDREDADS')));
//FIN PT82
      DeleteErreur (Sal, 'S4', StrToInt(GetControlText('ORDREDADS'))); //PT88
      Trace := TStringList.Create;
      if Trace <> nil then
         Trace.Add (Sal+' : Suppression de la période '+
                    GetControlText ('ORDREDADS'));

{$IFNDEF DADSUSEULE}
      CreeJnalEvt('001', '043', 'OK', NIL, NIL, Trace);
{$ENDIF}
      if Trace <> nil then
         FreeAndNil (Trace);

      StExiste:= 'SELECT PDS_SALARIE'+
                 ' FROM DADSDETAIL WHERE'+
                 ' PDS_SALARIE="'+Sal+'" AND'+
                 ' PDS_TYPE="'+TypeD+'" AND'+
                 ' PDS_ORDRE > 0 AND'+
                 ' PDS_EXERCICEDADS = "'+PGExercice+'"';

      if ExisteSQL(StExiste)=False then
{PT82
         DeleteDetail(Sal, TypeD, 0);
}
         begin
         DeleteDetail(Sal, 0);
         DeleteErreur (Sal, 'S3', 0); //PT88
         end;
      MAJQuery;
      CommitTrans;
   Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
              'Suppression DADS-U');
      END;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.AfficheCaption();
var
PRUD : TTabsheet;
GBFNC : TGroupBox;
begin
if ((Nbre<>'0') and (NbreTot<>'0')) then
   begin
   if (Nbre <> '1') then
      PerPrec.Enabled := TRUE
   else
      PerPrec.Enabled := FALSE;

//PT73-6
   if ((TypeD = '001') or (TypeD = '201')) then
      begin
      PRUD:= TTabsheet (GetControl ('TPRUD'));
      GBFNC:= TGroupBox (GetControl ('GBIRCANTECFNC'));
      if (Nbre <> NbreTot) then
         begin
{PT98
         PerSuiv.Enabled:= TRUE;
}
         if (PRUD <> nil) then
            PRUD.TabVisible:= False;
         if (GBFNC <> nil) then
            GBFNC.Visible:= False;
         end
      else
         begin
{PT98
         PerSuiv.Enabled:= FALSE;
}
         if (PRUD <> nil) then
            PRUD.TabVisible:= True;
         if (GBFNC <> nil) then
            GBFNC.Visible:= True;
         end;
      end;
//FIN PT73-6

//PT98
      if (Nbre <> NbreTot) then
         PerSuiv.Enabled:= TRUE
      else
         PerSuiv.Enabled:= FALSE;
//FIN PT98

   SetControlText('LPERIODE', 'Salarié '+Sal+' '+LibSal+' '+
                 GetControlText('EPRENOM')+'   Période '+Nbre+'/'+NbreTot);
   Ecran.Caption:= 'Saisie complémentaire DADS-U : '+GetControlText('LPERIODE');
   end
else
   begin
   Ecran.Caption:= 'Saisie complémentaire DADS-U : Salarié '+Sal+' '+LibSal+' '+
                   GetControlText('EPRENOM')+'   Période nouvelle';
   TFVierge(Ecran).Binsert.Enabled := False;
   TFVierge(Ecran).BDelete.Enabled := False;
   if Calculer <> NIL then
      Calculer.Enabled := False;

   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /    
Description .. : Procédure exécutée sur le chargement de la fiche
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.OnLoad;
begin
Inherited ;
ChargeZones ;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2001
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.OnClose;
begin
Ferme(QPer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/09/2001
Modifié le ... :
Description .. : Gestion du double-click sur les dates
Mots clefs ... : PAIE;PGDADSU;DATE
*****************************************************************}
procedure TOF_PG_DADSPERIODES.DateDoubleClick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/06/2001
Modifié le ... : 26/07/2001
Description .. : Chargement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.ChargeZones();
var
CodeIso, Libelle, nom, numero, StBuf, StConvention, StDADSD : String;
StMotifSortie, StPer, StPrudh, StSal, sWhere : String;
QCount, QRechBuf, QRechConvention, QRechDADSD, QRechInstit : TQuery;
QRechMetier, QRechPrudh, QRechSal : TQuery;
TDetail, TDetailBTP, TDetailD, TDetailPrudh, TDetailSal : Tob;
DateEntree, DateSortie : TDateTime;
i, j : integer;
ExisteIRCANTEC : boolean;
TabSheet : TTabsheet;
begin
MetABlanc;
MetABlancSal;
SetControlText('ENOM', '');
SetControlText('EPRENOM', '');

ExisteIRCANTEC:= False;	//PT73-8

if (SalChange = True) then
   begin
   StPer:= 'SELECT COUNT(PDE_SALARIE) AS NBRE'+
           ' FROM DADSPERIODES WHERE'+
           ' PDE_SALARIE = "'+Sal+'" AND'+
           ' PDE_TYPE="'+TypeD+'" AND'+
           ' PDE_ORDRE > 0 AND'+
           ' PDE_EXERCICEDADS = "'+PGExercice+'"';

   QCount:=OpenSql(StPer, True);
   if (not QCount.EOF) then
      NbreTot := IntToStr(QCount.FindField('NBRE').AsInteger);
   Ferme(QCount);

   if (QPer <> nil) then
      Ferme(QPer);
   StPer:= 'SELECT *'+
           ' FROM DADSPERIODES WHERE'+
           ' PDE_SALARIE = "'+Sal+'" AND'+
           ' PDE_TYPE="'+TypeD+'" AND'+
           ' PDE_ORDRE > 0 AND'+
           ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
           ' ORDER BY PDE_DATEDEBUT';

   QPer:=OpenSql(StPer, True);
   QPer.First;
   PerPrec.Enabled := FALSE;
   if QPer.EOF then
      begin
      PerSuiv.Enabled := FALSE;
      Nbre := '0';
      end
   else
      begin
      PerSuiv.Enabled := TRUE;
      Nbre := '1';
      end;
   SalChange := False;
   end;

if (not QPer.eof) then
   begin
   if (QPer.FindField ('PDE_ORDRE').Asinteger = 0) then
      begin
      SetControlText('ORDREDADS', '1');
      State := 'CREATION';
      end
   else
      begin
      SetControlText('ORDREDADS', IntToStr(QPer.FindField ('PDE_ORDRE').Asinteger));
      State:= 'MODIFICATION';
      end;

   SetControlText('TYPEDADS', QPer.FindField ('PDE_TYPE').Asstring);
   SetControlText('DATEDEBDADS', DateToStr(QPer.FindField ('PDE_DATEDEBUT').AsDateTime));
   SetControlText('DATEFINDADS', DateToStr(QPer.FindField ('PDE_DATEFIN').AsDateTime));
{PT73-3
   SetControlText('CDEBMOTIF1', QPer.FindField ('PDE_MOTIFDEB').Asstring);
   SetControlText('CFINMOTIF1', QPer.FindField ('PDE_MOTIFFIN').Asstring);
}
   SetControlText('EDEBMOTIF1', QPer.FindField ('PDE_MOTIFDEB').Asstring);
   SetControlText('EFINMOTIF1', QPer.FindField ('PDE_MOTIFFIN').Asstring);
//FIN PT73-3
   end
else
   begin
   SetControlText('ORDREDADS', '1');
   State := 'CREATION';
   SetControlText('TYPEDADS', '');
   SetControlText('DATEDEBDADS', DateToStr(IDate1900));
   SetControlText('DATEFINDADS', DateToStr(IDate1900));
{PT73-3
   SetControlText('CDEBMOTIF1', '');
   SetControlText('CFINMOTIF1', '');
}
   SetControlText('EDEBMOTIF1', '');
   SetControlText('EFINMOTIF1', '');
//FIN PT73-3
   end;

if (State = 'CREATION') then // chargement uniquement en mode creation
   begin
   StSal:= 'SELECT PSA_ETABLISSEMENT, PSA_DATEENTREE, PSA_MOTIFENTREE,'+
           ' PSA_DATESORTIE, PSA_MOTIFSORTIE, PSA_NUMEROSS, PSA_LIBELLE,'+
           ' PSA_PRENOM, PSA_NOMJF, PSA_SURNOM, PSA_CIVILITE, PSA_ADRESSE1,'+
           ' PSA_ADRESSE2, PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE, PSA_PAYS,'+
           ' PSA_DATENAISSANCE, PSA_COMMUNENAISS, PSA_DEPTNAISSANCE,'+
           ' PSA_PAYSNAISSANCE, PSA_NATIONALITE, PSA_MULTIEMPLOY,'+
           ' PSA_LIBELLEEMPLOI, PSA_CODEEMPLOI, PSA_CONDEMPLOI, PSA_DADSPROF,'+
           ' PSA_DADSCAT, PSA_CONVENTION, PSA_COEFFICIENT, PSA_INDICE,'+
           ' PSA_NIVEAU, PSA_REGIMESS, PSA_TAUXPARTIEL, PSA_ORDREAT,'+
           ' PSA_TYPEREGIME, PSA_REGIMEMAL, PSA_REGIMEAT, PSA_REGIMEVIP, PSA_REGIMEVIS,'+ //PT92-1
           ' PSA_TAUXPARTSS, PSA_TRAVETRANG, PSA_REMPOURBOIRE,'+
           ' PSA_ALLOCFORFAIT, PSA_REMBJUSTIF, PSA_PRISECHARGE,'+
           ' PSA_AUTREDEPENSE, PSA_PRUDHCOLL, PSA_PRUDHSECT, PSA_PRUDHVOTE,'+
           ' PSA_DADSDATE, PSA_DADSFRACTION,  PSA_UNITETRAVAIL,'+
           ' PSA_SITUATIONFAMIL, PSA_DATEANCIENNETE'+
           ' FROM SALARIES WHERE'+
           ' PSA_SALARIE="'+Sal+'"';

   QRechSal:=OpenSql(StSal,TRUE);
   SetControlText('LFRACTION_', Fraction);

   if (not QRechSal.EOF) then
      begin
{Si le nom de jeune fille est renseigné, je le prends comme nom patronymique
sinon, je prends le nom}
      LibSal := QRechSal.FindField('PSA_LIBELLE').Asstring;
      if ((QRechSal.FindField('PSA_NOMJF').Asstring = '') or
         (QRechSal.FindField('PSA_NOMJF').Asstring = null)) then
         SetControlText('ENOM', QRechSal.FindField('PSA_LIBELLE').Asstring)
      else
         begin
         SetControlText('ENOM', QRechSal.FindField('PSA_NOMJF').Asstring);
         SetControlText('LNOMJF', QRechSal.FindField('PSA_LIBELLE').Asstring);
         end;

      SetControlText('EPRENOM', QRechSal.FindField('PSA_PRENOM').Asstring);
      SetControlText('CETABLISSEMENT', QRechSal.FindField('PSA_ETABLISSEMENT').Asstring);

      DateEntree := QRechSal.FindField('PSA_DATEENTREE').AsDateTime;
      if (DateEntree >= DebExer) then
         begin
         SetControlText('DATEDEBDADS', DateToStr(DateEntree));
{PT73-3
         SetControlText('CDEBMOTIF1', '001');
}
         SetControlText('EDEBMOTIF1', '001');
         end
      else
         begin
         SetControlText('DATEDEBDADS', DateToStr(DebExer));
{PT73-3
         SetControlText('CDEBMOTIF1', '097');
}
         SetControlText('EDEBMOTIF1', '097');
         end;

      DateSortie := QRechSal.FindField('PSA_DATESORTIE').AsDateTime;
      StMotifSortie := QRechSal.FindField('PSA_MOTIFSORTIE').Asstring;
      if ((DateSortie <= FinExer) and
         (DateSortie > IDate1900)) then
         begin
         SetControlText('DATEFINDADS', DateToStr(DateSortie));
         if (StMotifSortie <> '') then
            begin
            case StrToInt(StMotifSortie) of
{PT73-3
                 31..35, 81 : SetControlText('CFINMOTIF1', '008');
                 59 : SetControlText('CFINMOTIF1', '010');
                 11..20 : SetControlText('CFINMOTIF1', '012');
                 38, 39 : SetControlText('CFINMOTIF1', '016');
}
                 31..37, 81 : SetControlText('EFINMOTIF1', '008');
                 59 : SetControlText('EFINMOTIF1', '010');
                 11..20 : SetControlText('EFINMOTIF1', '012');
                 38, 39 : SetControlText('EFINMOTIF1', '016');
                 403 : SetControlText('EFINMOTIF1', '018');
//FIN PT73-3
                 end;
            end;
         end
      else
         begin
         SetControlText('DATEFINDADS', DateToStr(FinExer));
{PT73-3
         SetControlText('CFINMOTIF1', '098');
}
         SetControlText('EFINMOTIF1', '098');
         end;

      SetControlText('CCONDEMPLOI', QRechSal.FindField('PSA_CONDEMPLOI').Asstring);
      SetControlText('CSTATUTPROF', QRechSal.FindField('PSA_DADSPROF').Asstring);
      SetControlText('CSTATUTCAT', QRechSal.FindField('PSA_DADSCAT').Asstring);
      SetControlText('CREGIMESS', QRechSal.FindField('PSA_REGIMESS').Asstring);
      //DEB PT92-1
      SetControlText('CHTYPEREGIME', QRechSal.FindField('PSA_TYPEREGIME').Asstring);
      SetControlText('CREGIMEMAL', QRechSal.FindField('PSA_REGIMEMAL').Asstring);
      SetControlText('CREGIMEAT', QRechSal.FindField('PSA_REGIMEAT').Asstring);
      SetControlText('CREGIMEVIP', QRechSal.FindField('PSA_REGIMEVIP').Asstring);
      SetControlText('CREGIMEVIS', QRechSal.FindField('PSA_REGIMEVIS').Asstring);
      if QRechSal.FindField('PSA_TYPEREGIME').Asstring = '-' then
      begin
        SetControlEnabled('CREGIMESS',True);
        SetControlEnabled('CREGIMEMAL',False);
        SetControlEnabled('CREGIMEAT',False);
        SetControlEnabled('CREGIMEVIP',False);
        SetControlEnabled('CREGIMEVIS',False);
      end
      else
      begin
        SetControlEnabled('CREGIMESS',False);
        SetControlEnabled('CREGIMEMAL',True);
        SetControlEnabled('CREGIMEAT',True);
        SetControlEnabled('CREGIMEVIP',True);
        SetControlEnabled('CREGIMEVIS',True);
      end;
      //FIN PT92-1

      if MultiEmp <> NIL then
         MultiEmp.Checked := QRechSal.FindField('PSA_MULTIEMPLOY').Asstring='X';

      SetControlText('ELIBELLEEMPLOI', QRechSal.FindField('PSA_LIBELLEEMPLOI').Asstring);
      SetControlText('EPCS', QRechSal.FindField('PSA_CODEEMPLOI').Asstring);
      SetControlText('ECOEFF', QRechSal.FindField('PSA_COEFFICIENT').Asstring);
      SetControlText('DECALAGE', Decalage);
      if (QRechSal.FindField('PSA_CONVENTION').Asstring <> '') then
         begin
         StConvention:= 'SELECT PCV_CONVENTION, PCV_IDCC'+
                        ' FROM CONVENTIONCOLL WHERE'+
                        ' ##PCV_PREDEFINI##'+
                        ' PCV_CONVENTION = "'+QRechSal.FindField('PSA_CONVENTION').Asstring+'"';
         QRechConvention:= OpenSql (StConvention,TRUE);
         SetControlText('ECONVCOLL', QRechConvention.FindField('PCV_IDCC').Asstring);
         Ferme(QRechConvention);
         end
      else
         SetControlText('ECONVCOLL', '9999');
      SetControlText('ESECTIONAT', QRechSal.FindField('PSA_ORDREAT').Asstring);
      SetControlText('CTRAVETRANG', QRechSal.FindField('PSA_TRAVETRANG').Asstring);

      if RemPourboire <> NIL then
         RemPourboire.Checked := QRechSal.FindField('PSA_REMPOURBOIRE').Asstring='X';
      if AllocForfait <> NIL then
         AllocForfait.Checked := QRechSal.FindField('PSA_ALLOCFORFAIT').Asstring='X';
      if RembJustif <> NIL then
         RembJustif.Checked := QRechSal.FindField('PSA_REMBJUSTIF').Asstring='X';
      if PriseCharge <> NIL then
         PriseCharge.Checked := QRechSal.FindField('PSA_PRISECHARGE').Asstring='X';
      if AutreDepense <> NIL then
         AutreDepense.Checked := QRechSal.FindField('PSA_AUTREDEPENSE').Asstring='X';
      SetControlText('CDUREETRAV', QRechSal.FindField('PSA_UNITETRAVAIL').Asstring);
      SetControlText('LNUMSS', Copy(QRechSal.FindField('PSA_NUMEROSS').Asstring, 1, 13));
      SetControlText('LCIVILITE', QRechSal.FindField('PSA_CIVILITE').Asstring);
      SetControlText('LSURNOM', QRechSal.FindField('PSA_SURNOM').Asstring);

      numero := '';
      nom := '';
      AdresseNormalisee (QRechSal.FindField('PSA_ADRESSE1').Asstring, numero,
                        nom);
      SetControlText('LADRNUM', numero);
      SetControlText('LADRNOM', nom);
      if (QRechSal.FindField('PSA_ADRESSE2').Asstring <> '') then
         begin
         if (QRechSal.FindField('PSA_ADRESSE3').Asstring <> '') then
            SetControlText('LADRCOMPL', QRechSal.FindField('PSA_ADRESSE2').Asstring+
                           ' '+QRechSal.FindField('PSA_ADRESSE3').Asstring)
         else
            SetControlText('LADRCOMPL', QRechSal.FindField('PSA_ADRESSE2').Asstring);
         end
      else
         if (QRechSal.FindField('PSA_ADRESSE3').Asstring <> '') then
            SetControlText('LADRCOMPL', QRechSal.FindField('PSA_ADRESSE3').Asstring);
      SetControlText('LCODEPOSTAL', QRechSal.FindField('PSA_CODEPOSTAL').Asstring);
      SetControlText('LVILLE', PGUpperCase(QRechSal.FindField('PSA_VILLE').Asstring));
      PaysISOLib(QRechSal.FindField('PSA_PAYS').Asstring, CodeIso, Libelle);
      if ((CodeIso <> '') and (Libelle <> '')) then
         begin
         SetControlText('LCODEPAYS', CodeIso);
         SetControlText('LNOMPAYS', Libelle);
         end
      else
         begin
         SetControlText('LCODEPAYS', '...');
         SetControlText('LNOMPAYS', '...');
         end;
      SetControlText('LDATENAISSANCE', QRechSal.FindField('PSA_DATENAISSANCE').Asstring);
      SetControlText('LCOMMUNENAISS', QRechSal.FindField('PSA_COMMUNENAISS').Asstring);
      SetControlText('LPAYSNAISSANCE', QRechSal.FindField('PSA_PAYSNAISSANCE').Asstring);
      SetControlText('LDEPTNAISSANCE', QRechSal.FindField('PSA_DEPTNAISSANCE').Asstring);
      if (GetControlText('LDEPTNAISSANCE') = '20A') then
         SetControlText('LDEPTNAISSANCE', '2A');
      if (GetControlText('LDEPTNAISSANCE') = '20B') then
         SetControlText('LDEPTNAISSANCE', '2B');
      SetControlText('LNATIONALITE', QRechSal.FindField('PSA_NATIONALITE').Asstring);
      end;
   Ferme(QRechSal);

// Remplissage des THLabel associés aux Elipsis
   if (GetControlText('ELIBELLEEMPLOI') <> '') then
      SetControlText('LLIBELLEEMPLOI_', RechDom('PGLIBEMPLOI',
                     GetControlText('ELIBELLEEMPLOI'), FALSE));

   if (GetControlText('EPCS') <> '') then
      begin
      if VH_Paie.PGPCS2003 then
         SetControlText('LPCS_', RechDom('PGCODEPCSESE',GetControlText('EPCS'),
                        False))
      else
         SetControlText('LPCS_', RechDom('PGCODEEMPLOI', GetControlText('EPCS'),
                        FALSE));
      end;

   if (GetControlText('ECOEFF') <> '') then
      SetControlText('LCOEFF_', RechDom('PGCOEFFICIENT', GetControlText('ECOEFF'), FALSE));

   if (GetControlText('ECONVCOLL') <> '') then
      SetControlText('LCONVCOLL_', RechDom('PGIDCC', GetControlText('ECONVCOLL'), FALSE));
   end
else
   if (State = 'MODIFICATION') then // chargement uniquement en mode modification
      begin
      StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
                ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
                ' PDS_DONNEEAFFICH'+
                ' FROM DADSDETAIL WHERE'+
                ' PDS_SALARIE="'+Sal+'" AND'+
                ' PDS_TYPE="'+TypeD+'" AND'+
                ' PDS_ORDRE=0 AND'+
                ' PDS_EXERCICEDADS = "'+PGExercice+'"'+
                ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,PDS_DATEFIN';
      QRechDADSD:=OpenSql(StDADSD,TRUE);
      TDetailSal := TOB.Create('Les détails salariés', NIL, -1);
      TDetailSal.LoadDetailDB('DADSDETAIL','','',QRechDADSD,False);
      Ferme(QRechDADSD);

      StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
                ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
                ' PDS_DONNEEAFFICH'+
                ' FROM DADSDETAIL WHERE'+
                ' PDS_SALARIE="'+Sal+'" AND'+
                ' PDS_TYPE="'+TypeD+'" AND'+
                ' PDS_ORDRE='+GetControlText('ORDREDADS')+' AND'+
                ' PDS_EXERCICEDADS = "'+PGExercice+'"'+
                ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,PDS_DATEFIN';
      QRechDADSD:=OpenSql(StDADSD,TRUE);
      TDetail := TOB.Create('Les détails', NIL, -1);
      TDetail.LoadDetailDB('DADSDETAIL','','',QRechDADSD,False);
      Ferme(QRechDADSD);

      StDADSD:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
                ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
                ' PDS_DONNEEAFFICH'+
                ' FROM DADSDETAIL WHERE'+
                ' PDS_SALARIE="'+Sal+'" AND'+
                ' PDS_TYPE="'+TypeD+'" AND'+
                ' PDS_ORDRE='+GetControlText('ORDREDADS')+' AND'+
                ' PDS_EXERCICEDADS = "'+PGExercice+'"'+
                ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,PDS_DATEFIN';
      QRechDADSD:=OpenSql(StDADSD,TRUE);
      TDetailBTP := TOB.Create('Les détails', NIL, -1);
      TDetailBTP.LoadDetailDB('DADSDETAIL','','',QRechDADSD,False);
      Ferme(QRechDADSD);

      StPrudh:= 'SELECT PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_DATEDEBUT,'+
                ' PDS_DATEFIN, PDS_ORDRESEG, PDS_SEGMENT, PDS_DONNEE,'+
                ' PDS_DONNEEAFFICH'+
                ' FROM DADSDETAIL WHERE'+
                ' PDS_SALARIE="'+Sal+'" AND'+
                ' PDS_TYPE="'+TypeD+'" AND'+
                ' PDS_ORDRE='+GetControlText('ORDREDADS')+' AND'+
                ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                ' PDS_SEGMENT LIKE "S41.G02%"'+
                ' ORDER BY PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEDEBUT,PDS_DATEFIN';
      QRechPrudh:=OpenSql(StPrudh,TRUE);
      TDetailPrudh := TOB.Create('Les détails', NIL, -1);
      TDetailPrudh.LoadDetailDB('DADSDETAIL','','',QRechPrudh,False);
      Ferme(QRechPrudh);

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LNUMSS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.002'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('ENOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         LibSal := TDetailD.GetValue('PDS_DONNEEAFFICH');
         end;

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.003'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EPRENOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.004'], TRUE);
      If (TDetailD <> NIL) then
         begin
         SetControlText('LNOMJF', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         LibSal := TDetailD.GetValue('PDS_DONNEEAFFICH');
         end;

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.006'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LSURNOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.007'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LCIVILITE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LADRCOMPL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.003'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LADRNUM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.006'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LADRNOM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.010'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LCODEPOSTAL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.012'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LVILLE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.013'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LCODEPAYS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.008.014'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LNOMPAYS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.009'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LDATENAISSANCE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.010'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LCOMMUNENAISS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.011'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LDEPTNAISSANCE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.012'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LPAYSNAISSANCE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailSal.FindFirst(['PDS_SEGMENT'], ['S30.G01.00.013'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('LNATIONALITE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('DATEDEBDADS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

{PT73-3
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CDEBMOTIF1', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.002'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CDEBMOTIF2', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.003'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CDEBMOTIF3', TDetailD.GetValue('PDS_DONNEEAFFICH'));
}
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.001'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EDEBMOTIF1', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (DebMotif1);
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.002'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EDEBMOTIF2', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (DebMotif2);
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.002.003'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EDEBMOTIF3', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (DebMotif3);
         end;
//FIN PT73-3

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.003'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('DATEFINDADS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

{PT73-3
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CFINMOTIF1', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.002'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CFINMOTIF2', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.003'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CFINMOTIF3', TDetailD.GetValue('PDS_DONNEEAFFICH'));
}
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.001'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EFINMOTIF1', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (FinMotif1);
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.002'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EFINMOTIF2', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (FinMotif2);
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.004.003'], TRUE);
      if (TDetailD <> NIL) then
         begin
         SetControlText('EFINMOTIF3', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         MotifChange (FinMotif3);
         end;
//FIN PT73-3

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.005'], TRUE);
      if (TDetailD <> NIL) then
         begin
         if (Length (TDetailD.GetValue('PDS_DONNEEAFFICH'))>3) then
            begin
            StBuf:= 'SELECT ET_ETABLISSEMENT'+
                    ' FROM ETABLISS WHERE'+
                    ' ET_SIRET LIKE "%'+TDetailD.GetValue('PDS_DONNEEAFFICH')+'"';
            QRechBuf:= OpenSql (StBuf, TRUE);
            if (not QRechBuf.EOF) then
               SetControlText ('CETABLISSEMENT',
                               QRechBuf.FindField('ET_ETABLISSEMENT').AsString);
            Ferme (QRechBuf);
            end
         else
            SetControlText ('CETABLISSEMENT',
                            TDetailD.GetValue('PDS_DONNEEAFFICH'));
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.008.001'], TRUE);
      if ((MultiEmp <> NIL) and (TDetailD <> NIL)) then
         MultiEmp.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='02';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.009'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('DECALAGE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.010'], TRUE);
      if (TDetailD <> NIL) then
         begin
         if (Length (TDetailD.GetValue('PDS_DONNEEAFFICH'))>3) then
            begin
            StBuf:= 'SELECT CC_CODE'+
                   ' FROM CHOIXCOD WHERE'+
                   ' CC_LIBELLE="'+TDetailD.GetValue('PDS_DONNEEAFFICH')+'" AND'+
                   ' CC_TYPE = "PLE"';
            QRechBuf:= OpenSql (StBuf, TRUE);
            if (not QRechBuf.EOF) then
               SetControlText ('ELIBELLEEMPLOI',
                               QRechBuf.FindField('CC_CODE').AsString);
            Ferme (QRechBuf);
            end
         else
            SetControlText ('ELIBELLEEMPLOI',
                            TDetailD.GetValue('PDS_DONNEEAFFICH'));
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.011'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EPCS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.012.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CCONTRAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.012.002'], TRUE);
         if ((PrudDP <> NIL) and (TDetailD <> NIL)) then
            PrudDP.Checked:= TDetailD.GetValue('PDS_DONNEEAFFICH')='02';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.013'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CCONDEMPLOI', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.014'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CSTATUTPROF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.015.002'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CSTATUTCAT', TDetailD.GetValue('PDS_DONNEEAFFICH'))
//PT73-5
      else
         SetControlText('CSTATUTCAT', '90');
//FIN PT73-5

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.016'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECONVCOLL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.017'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECOEFF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.018.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText('CREGIMESS', TDetailD.GetValue('PDS_DONNEEAFFICH'))
//DEB PT92-1
      else
         SetControlText ('CHTYPEREGIME', 'X');

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.018.002'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CREGIMEMAL', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.018.003'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CREGIMEAT', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.018.004'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CREGIMEVIP', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.018.005'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CREGIMEVIS', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      if (GetControlText ('CHTYPEREGIME')='-') then
         begin
         SetControlEnabled ('CREGIMESS', True);
         SetControlEnabled ('CREGIMEMAL', False);
         SetControlEnabled ('CREGIMEAT', False);
         SetControlEnabled ('CREGIMEVIP', False);
         SetControlEnabled ('CREGIMEVIS', False);
         end
      else
         begin
         SetControlEnabled ('CREGIMESS', False);
         SetControlEnabled ('CREGIMEMAL', True);
         SetControlEnabled ('CREGIMEAT', True);
         SetControlEnabled ('CREGIMEVIP', True);
         SetControlEnabled ('CREGIMEVIS', True);
         end;
//FIN PT92-1

//PT93-4
      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.018.006'], TRUE);
      if ((Alsace <> NIL) and (TDetailD <> NIL)) then
         Alsace.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='01';
//FIN PT93-4

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.020'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ETAUXPARTIEL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.021'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EHEURESTRAV', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.022'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EHEURESSAL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.023'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EDERNIERMOIS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.024'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EHEURESCHOMPAR', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.025'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ESECTIONAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.026'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ERISQUEAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.027'], TRUE);
      if ((CodeBureau <> NIL) and (TDetailD <> NIL)) then
         CodeBureau.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='B';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.028'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ETAUXAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.029.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBRUTSS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT93-4
      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.029.003'], TRUE);
      if ((BaseReelle <> NIL) and (TDetailD <> NIL)) then
         BaseReelle.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='01';
//FIN PT93-4

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.030.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EPLAFONDSS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.032.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBASECSG', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.033.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBASECRDS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.034'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CTRAVETRANG', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.035.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBRUTFISC', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.037.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EAVANTAGENAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.038'], TRUE);
      if ((Nourriture <> NIL) and (TDetailD <> NIL)) then
         Nourriture.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='N';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.039'], TRUE);
      if ((Logement <> NIL) and (TDetailD <> NIL)) then
         Logement.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='L';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.040'], TRUE);
      if ((Voiture <> NIL) and (TDetailD <> NIL)) then
         Voiture.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='V';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.041'], TRUE);
      if ((AutreAvant <> NIL) and (TDetailD <> NIL)) then
         AutreAvant.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='A';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.042.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ERETSALAIRE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.043'], TRUE);
      if ((RemPourboire <> NIL) and (TDetailD <> NIL)) then
         RemPourboire.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='P';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.044.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EFRAISPROF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.045'], TRUE);
      if ((AllocForfait <> NIL) and (TDetailD <> NIL)) then
         AllocForfait.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='F';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.046'], TRUE);
      if ((RembJustif <> NIL) and (TDetailD <> NIL)) then
         RembJustif.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='R';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.047'], TRUE);
      if ((PriseCharge <> NIL) and (TDetailD <> NIL)) then
         PriseCharge.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='P';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.048'], TRUE);
      if ((AutreDepense <> NIL) and (TDetailD <> NIL)) then
         AutreDepense.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='D';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.049.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECHQVACANCE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.052.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EIMPOTRETSOURC', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.053.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EINDEXPATRIAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.055.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ETOTIMPOTAXE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.056.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBASEIMPO1', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.057.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EBASEIMPO2', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.058.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ETAXESAL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.063.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EREVENUSACTIV', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.065.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECP', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.066.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EAUTREREV', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.067.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EEPARGNE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.068'], TRUE);
      if ((NTIC <> NIL) and (TDetailD <> NIL)) then
         NTIC.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='T';

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.069.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EINDIMPATRIAT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT73-5
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.070.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ESERVICEPERS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT93-5
      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G01.00.071'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CCNBF', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
//FIN PT93-5

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.072.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECPPLAF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.00.073.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EHEURESSUPP', TDetailD.GetValue('PDS_DONNEEAFFICH'));
//FIN PT73-5

      if RegimeComp <> NIL then
         begin
         j := 110;
         for i := 1 to 5 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*2)-1)], TRUE);
             if TDetailD <> NIL then
                RegimeComp.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             if (RegimeComp.CellValues[0,i] <> '') then
                begin
                sWhere:= 'SELECT PIP_ABREGE'+
                         ' FROM INSTITUTIONPAYE WHERE'+
                         ' ##PIP_PREDEFINI## PIP_INSTITUTION = "'+RegimeComp.CellValues[0,i]+'"';
                QRechInstit := OpenSql(sWhere,TRUE);
                if (not QRechInstit.EOF) then
                   RegimeComp.CellValues[1,i] := QRechInstit.Fields[0].asstring;
                Ferme(QRechInstit);
                if ((RegimeComp.CellValues[0,i] = 'I001') or
                   (RegimeComp.CellValues[0,i] = 'I002')) then
                   ExisteIRCANTEC:= True;
                end;

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*2))], TRUE);
             if TDetailD <> NIL then
                RegimeComp.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

      if Basebrutespec <> NIL then
         begin
         j := 200;
         for i := 1 to 4 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
             if TDetailD <> NIL then
                Basebrutespec.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             if Basebrutespec.CellValues[0,i] <> '' then
                Basebrutespec.CellValues[1,i] := RechDom('PGBASEBRUTESPEC', Basebrutespec.CellValues[0,i], FALSE);
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if TDetailD <> NIL then
                Basebrutespec.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

      if Baseplafspec <> NIL then
         begin
         j := 300;
         for i := 1 to 4 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
             if TDetailD <> NIL then
                Baseplafspec.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             if Baseplafspec.CellValues[0,i] <> '' then
                Baseplafspec.CellValues[1,i] := RechDom('PGBASEPLAFSPEC', Baseplafspec.CellValues[0,i], FALSE);

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if TDetailD <> NIL then
                Baseplafspec.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

      if Somiso <> NIL then
         begin
         j := 400;
         for i := 1 to 4 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*6)-5)], TRUE);
             if TDetailD <> NIL then
                Somiso.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             if Somiso.CellValues[0,i] <> '' then
                Somiso.CellValues[1,i] := RechDom('PGSOMISO', Somiso.CellValues[0,i], FALSE);

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*6)-4)], TRUE);
             if TDetailD <> NIL then
                Somiso.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*6)-3)], TRUE);
             if TDetailD <> NIL then
                Somiso.CellValues[3,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*6)-1)], TRUE);
             if TDetailD <> NIL then
                Somiso.CellValues[4,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S41.G01.05.002.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('ECSGSPECM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      if Exoner <> NIL then
         begin
         j := 600;
         for i := 1 to 10 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*5)-4)], TRUE);
             if TDetailD <> NIL then
                Exoner.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             if Exoner.CellValues[0,i] <> '' then
                Exoner.CellValues[1,i] := RechDom('PGEXONERATION', Exoner.CellValues[0,i], FALSE);

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*5)-3)], TRUE);
             if TDetailD <> NIL then
                Exoner.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*5)-1)], TRUE);
             if TDetailD <> NIL then
                Exoner.CellValues[3,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

      TDetailD := TDetailPrudh.FindFirst(['PDS_SEGMENT'],
                                         ['S41.G02.00.008'], TRUE);
      if ((PrudPresent <> NIL) and (TDetailD <> NIL)) then
         PrudPresent.Checked:= TDetailD.GetValue('PDS_DONNEEAFFICH')='01';

      TDetailD := TDetailPrudh.FindFirst(['PDS_SEGMENT'],
                                         ['S41.G02.00.009'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CPRUDHCOLL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailPrudh.FindFirst(['PDS_SEGMENT'],
                                         ['S41.G02.00.010'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CPRUDHSECT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT93-5
      j:= 800;
      for i:= 1 to 4 do
          begin
          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
          if (TDetailD<>NIL) then
             begin
             TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if (TDetailD<>NIL) then
                begin
                if (i=1) then
                   SetControlText ('EEPARGNESALP', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=2) then
                   SetControlText ('EEPARGNESALI', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=3) then
                   SetControlText ('EEPARGNESALA', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=4) then
                   SetControlText ('EEPARGNESALD', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
                end;
             end;
          end;

      j:= 820;
      for i:= 1 to 3 do
          begin
          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*2)-1)], TRUE);
          if (TDetailD<>NIL) then
             SetControlText ('EACTIONSN'+IntToStr (i),
                             TDetailD.GetValue('PDS_DONNEEAFFICH'));

          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*2))], TRUE);
          if (TDetailD<>NIL) then
             SetControlText ('EACTIONSV'+IntToStr (i),
                             TDetailD.GetValue('PDS_DONNEEAFFICH'));
          end;

      j:= 830;
      for i:= 1 to 2 do
          begin
          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
          if (TDetailD<>NIL) then
             begin
             TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if (TDetailD<>NIL) then
                begin
                if (i=1) then
                   SetControlText ('EPPRESTO', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=2) then
                   SetControlText ('EPPTRANSPORT', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
                end;
             end;
          end;

      j:= 840;
      for i:= 1 to 4 do
          begin
          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
          if (TDetailD<>NIL) then
             begin
             TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if (TDetailD<>NIL) then
                begin
                if (i=1) then
                   SetControlText ('ECASPARTS', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=2) then
                   SetControlText ('ECASPARTA', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=3) then
                   SetControlText ('ECASPARTV', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=4) then
                   SetControlText ('ECASPARTG', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
                end;
             end;
          end;

      j:= 860;
      for i:= 1 to 2 do
          begin
          TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-2)], TRUE);
          if (TDetailD<>NIL) then
             begin
             TDetailD:= TDetail.FindFirst (['PDS_ORDRESEG'], [IntToStr(j+(i*3)-1)], TRUE);
             if (TDetailD<>NIL) then
                begin
                if (i=1) then
                   SetControlText ('EILICENC', TDetailD.GetValue ('PDS_DONNEEAFFICH'))
                else
                if (i=2) then
                   SetControlText ('EIRETRAITE', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
                end;
             end;
          end;

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.30.001.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EGSS', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.30.002.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EGCSG', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.35.001.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('CETYPEREM', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.35.001.002'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EEMONTANTREM', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.35.002'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EENBJOURS', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.35.003.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EEMONTANTPAT', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S41.G30.35.004.001'], TRUE);
      if (TDetailD<>NIL) then
         SetControlText ('EEMONTANTSAL', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
//FIN PT93-5

//PT73-8
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('CIRCANTECPERIODICITE',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

{PT93-5
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.002'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECNBPAIE',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));
}

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.004'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECHEBDO',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.007.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECBRUT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.008.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECPLAFOND',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G01.00.009.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECMEDECINS',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT95-3
      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G02.00.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECFNCNUM',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S42.G02.00.002.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText ('EIRCANTECFNCMONTANT',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));
//FIN PT95-3

//FIN PT73-8

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S44.G01.00.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('CDUREETRAV', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S44.G01.00.002'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EDUREETRAV', TDetailD.GetValue('PDS_DONNEEAFFICH'));

//PT93-5
      TDetailD:= TDetail.FindFirst (['PDS_SEGMENT'], ['S44.G01.00.003'], TRUE);
      if ((AGFF <> NIL) and (TDetailD<>NIL)) then
         AGFF.Checked := TDetailD.GetValue('PDS_DONNEEAFFICH')='01';
//FIN PT93-5

      TDetailD := TDetail.FindFirst(['PDS_SEGMENT'], ['S45.G01.00.001'], TRUE);
      if (TDetailD <> NIL) then
         SetControlText('EENTREE', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      if Prevoyance <> NIL then
         begin
{PT93-5
         j:= 801;
}
         j:= 921;
//FIN PT93-5         
         for i := 1 to 4 do
             begin
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-11)], TRUE);
             if TDetailD <> NIL then
{PT102
                Prevoyance.CellValues[0,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
}
                Prevoyance.CellValues[0,i]:= RechDom ('PGDADSPREV',
                                                      TDetailD.GetValue ('PDS_DONNEEAFFICH'),
                                                      False);

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-10)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[1,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-9)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[2,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-8)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[3,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

//PT93-5
             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-7)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[4,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
//FIN PT93-5

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-6)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[5,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-5)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[6,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-4)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[7,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-2)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[8,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12)-1)], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[9,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');

             TDetailD := TDetail.FindFirst(['PDS_ORDRESEG'], [IntToStr(j+(i*12))], TRUE);
             if TDetailD <> NIL then
                Prevoyance.CellValues[10,i] := TDetailD.GetValue('PDS_DONNEEAFFICH');
             end;
         end;

//Récupération des éléments BTP
      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPADHES', TDetailD.GetValue('PDS_DONNEEAFFICH'));

{PT93-5
      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.002'], TRUE);
}
      TDetailD:= TDetailBTP.FindFirst (['PDS_SEGMENT'], ['S66.G01.00.002.001'], TRUE);
//FIN PT93-5
      If (TDetailD<>NIL) then
         SetControlText ('CBTPUNITTEMPS', TDetailD.GetValue ('PDS_DONNEEAFFICH'));

//PT93-5
      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.002.002'], TRUE);
      if ((Ouvre<>NIL) and (TDetailD<>NIL)) then
         Ouvre.Checked:= TDetailD.GetValue('PDS_DONNEEAFFICH')='02';
//FIN PT93-5

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.003'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPTEMPS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.004'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPANCIENENT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.005'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPANCIENPROF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.006'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPASSEDIC', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.007'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPUNITSALMOY', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.008'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPSALMOY', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.009'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPUNITHOR', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.010'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPHORSAL', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.011'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPHOREMP', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.012'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPSITFAM', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.013'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPENFANT', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.014'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPFRAISPROF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.015'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPSTATUTCOTIS', TDetailD.GetValue('PDS_DONNEEAFFICH'));

{PT73-5
      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.016'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPNIVEAU', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.017'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPCOEFF', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.018'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPPOSITION', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.019'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPECHELON', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.020'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('EBTPCATEGORIE', TDetailD.GetValue('PDS_DONNEEAFFICH'));
}
      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.021'], TRUE);
      If (TDetailD <> NIL) then
         begin
         SetControlText ('EMETIERBTP', TDetailD.GetValue('PDS_DONNEEAFFICH'));
         sWhere:= 'SELECT PMB_METIERLIB'+
                  ' FROM METIERBTP WHERE'+
                  ' ##PMB_PREDEFINI##'+
                  ' PMB_METIERBTP = "'+GetControlText('EMETIERBTP')+'"';
         QRechMetier:= OpenSql(sWhere,TRUE);
         if (not QRechMetier.EOF) then
            SetControlText ('LMETIERBTP_', QRechMetier.Fields[0].asstring);
         Ferme (QRechMetier);
         end;

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.022'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText('CBTPAFFILIRC', TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.023.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText ('EBTPBASECPBTP',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));

      TDetailD := TDetailBTP.FindFirst(['PDS_SEGMENT'], ['S66.G01.00.024.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText ('EBTPBASECOTIS',
                         TDetailD.GetValue('PDS_DONNEEAFFICH'));
//PT73-5
      TDetailD:= TDetailBTP.FindFirst (['PDS_SEGMENT'], ['S66.G01.00.025.001'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText ('EBTPBASECOTISOPPBTP',
                         TDetailD.GetValue ('PDS_DONNEEAFFICH'));

      TDetailD:= TDetailBTP.FindFirst (['PDS_SEGMENT'], ['S66.G01.00.026'], TRUE);
      If (TDetailD <> NIL) then
         SetControlText ('ECLASSIFBTP', TDetailD.GetValue ('PDS_DONNEEAFFICH'));
//FIN PT73-5

// Remplissage des THLabel associés aux Elipsis
      If (GetControlText('ELIBELLEEMPLOI') <> '') then
         SetControlText('LLIBELLEEMPLOI_', RechDom('PGLIBEMPLOI',
                        GetControlText('ELIBELLEEMPLOI'), FALSE));

      If (GetControlText('EPCS') <> '') then
         begin
         if VH_Paie.PGPCS2003 then
            SetControlText('LPCS_', RechDom('PGCODEPCSESE',
                           GetControlText('EPCS'), False))
         else
            SetControlText('LPCS_', RechDom('PGCODEEMPLOI',
                           GetControlText('EPCS'), FALSE));
         end;

      If (GetControlText('ECOEFF') <> '') then
         SetControlText('LCOEFF_', RechDom('PGCOEFFICIENT',
                        GetControlText('ECOEFF'), FALSE));

      If (GetControlText('ECONVCOLL') <> '') then
         SetControlText('LCONVCOLL_', RechDom('PGIDCC',
                        GetControlText('ECONVCOLL'), FALSE));
      FreeAndNil(TDetailSal);
      FreeAndNil(TDetail);
      FreeAndNil(TDetailBTP);
      FreeAndNil(TDetailPrudh);
      end;

//PT73-8
TabSheet:= TTabsheet(GetControl('TIRCANTEC'));
if (TabSheet <> nil) then
   begin
   if (ExisteIRCANTEC=True) then
{PT77-1
      TabSheet.TabVisible:= True
}
      begin
      if ((TypeD='001') or (TypeD='201')) then
         TabSheet.TabVisible:= True;
      end
   else
      TabSheet.TabVisible:= False;
   end;
//FIN PT73-8
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/08/2001
Modifié le ... :
Description .. : Enregistrement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SauveZones();
var
Buffer, BufOrig, BufDest, CodeIso, Libelle, NIC, SIREN, StEtab, StEvt : string;
StSal : String;
QRechEtab, QRechEvt : TQuery;
PerDeb, PerFin : TDateTime;
i, j, Ordre : integer;
AnneeA, JourJ, MoisM : Word;
CleDADS : TCleDADS;
Ircantec : Boolean;
begin
Trace := TStringList.Create;

Ordre := StrToInt(GetControlText('ORDREDADS'));
PerDeb := StrToDate(GetControlText('DATEDEBDADS'));
PerFin := StrToDate(GetControlText('DATEFINDADS'));

{PT73-9
CreeEntete(Sal, TypeD, Ordre, PerDeb, PerFin, GetControlText('CDEBMOTIF1'),
          GetControlText('CFINMOTIF1'));
}
CleDADS.Salarie:= Sal;
CleDADS.TypeD:= TypeD;
CleDADS.Num:= Ordre;
CleDADS.DateDeb:= PerDeb;
CleDADS.DateFin:= PerFin;
CleDADS.Exercice:= PGExercice;
CreeEntete (CleDADS, GetControlText('EDEBMOTIF1'), GetControlText('EFINMOTIF1'),
            Decalage);
//FIN PT73-9

//PT73-9
CleDADS.Num:= 0;
CleDADS.DateDeb:= DebExer;
CleDADS.DateFin:= FinExer;
//FIN PT73-9

If (GetControlText('LNUMSS') <> '') then
   CreeDetail (CleDADS, 1, 'S30.G01.00.001', GetControlText('LNUMSS'),
               GetControlText('LNUMSS'))
else
   begin
   if (GetControlText('LCIVILITE') = 'MR') then
      CreeDetail (CleDADS, 1, 'S30.G01.00.001', '1999999999999',
                  '1999999999999')
   else
      CreeDetail (CleDADS, 1, 'S30.G01.00.001', '2999999999999',
                  '2999999999999');
   end;

CreeDetail (CleDADS, 2, 'S30.G01.00.002', GetControlText('ENOM'),
            GetControlText('ENOM'));
if ((GetControlText('LNOMJF') <> '') and
   (GetControlText('LNOMJF') <> '...')) then
   CreeDetail (CleDADS, 4, 'S30.G01.00.004', GetControlText('LNOMJF'),
               GetControlText('LNOMJF'));

if (GetControlText('EPRENOM') <> '') then
   CreeDetail (CleDADS, 3, 'S30.G01.00.003', GetControlText('EPRENOM'),
               GetControlText('EPRENOM'));

if ((GetControlText('LSURNOM') <> '') and
   (GetControlText('LSURNOM') <> '...')) then
   CreeDetail (CleDADS, 5, 'S30.G01.00.006', GetControlText('LSURNOM'),
               GetControlText('LSURNOM'));

if (GetControlText('LCIVILITE') = 'MR') then
   CreeDetail (CleDADS, 6, 'S30.G01.00.007', '01', GetControlText('LCIVILITE'))
else
if (GetControlText('LCIVILITE') = 'MME') then
   CreeDetail (CleDADS, 6, 'S30.G01.00.007', '02', GetControlText('LCIVILITE'))
else
if (GetControlText('LCIVILITE') = 'MLE') then
   CreeDetail (CleDADS, 6, 'S30.G01.00.007', '03', GetControlText('LCIVILITE'));

if ((GetControlText('LADRCOMPL') <> '...') and
   (GetControlText('LADRCOMPL') <> '')) then
   CreeDetail (CleDADS, 7, 'S30.G01.00.008.001', GetControlText('LADRCOMPL'),
               GetControlText('LADRCOMPL'));

if ((GetControlText('LADRNUM') <> '...') and
   (GetControlText('LADRNUM') <> '')) then
   CreeDetail (CleDADS, 8, 'S30.G01.00.008.003', GetControlText('LADRNUM'),
               GetControlText('LADRNUM'));

if ((GetControlText('LADRNOM') <> '...') and
   (GetControlText('LADRNOM') <> '')) then
   CreeDetail (CleDADS, 9, 'S30.G01.00.008.006', GetControlText('LADRNOM'),
               GetControlText('LADRNOM'));

if (GetControlText('LCODEPOSTAL') <> '') then
   CreeDetail (CleDADS, 11, 'S30.G01.00.008.010', GetControlText('LCODEPOSTAL'),
               GetControlText('LCODEPOSTAL'))
else
   CreeDetail (CleDADS, 11, 'S30.G01.00.008.010', '', '');

CreeDetail (CleDADS, 12, 'S30.G01.00.008.012',
            PGUpperCase(GetControlText('LVILLE')), GetControlText('LVILLE'));

if ((GetControlText('LCODEPAYS')<>'...') and
{PT93-1
   (GetControlText('LCODEPAYS')<>'FR')) then
}
   (GetControlText('LCODEPAYS')<>'FR') and
   (GetControlText('LCODEPAYS')<>'GF') and
   (GetControlText('LCODEPAYS')<>'GP') and
   (GetControlText('LCODEPAYS')<>'MC') and
   (GetControlText('LCODEPAYS')<>'MQ') and
   (GetControlText('LCODEPAYS')<>'NC') and
   (GetControlText('LCODEPAYS')<>'PF') and
   (GetControlText('LCODEPAYS')<>'PM') and
   (GetControlText('LCODEPAYS')<>'RE') and
   (GetControlText('LCODEPAYS')<>'TF') and
   (GetControlText('LCODEPAYS')<>'WF') and
   (GetControlText('LCODEPAYS')<>'YT')) then
//FIN PT93-1   
   begin
   CreeDetail (CleDADS, 13, 'S30.G01.00.008.013', GetControlText('LCODEPAYS'),
               GetControlText('LCODEPAYS'));
   CreeDetail (CleDADS, 14, 'S30.G01.00.008.014', GetControlText('LNOMPAYS'),
               GetControlText('LNOMPAYS'));
   end;

ForceNumerique(GetControlText('LDATENAISSANCE'), Buffer);
if Length(Buffer) = 8 then
   CreeDetail (CleDADS, 15, 'S30.G01.00.009', Buffer,
               GetControlText('LDATENAISSANCE'))
else
   CreeDetail (CleDADS, 15, 'S30.G01.00.009', '', '');

if (GetControlText('LCOMMUNENAISS') <> '...') then
   CreeDetail (CleDADS, 16, 'S30.G01.00.010', GetControlText('LCOMMUNENAISS'),
               GetControlText('LCOMMUNENAISS'));

if (GetControlText('LPAYSNAISSANCE') = 'FRA') then
   CreeDetail (CleDADS, 17, 'S30.G01.00.011', GetControlText('LDEPTNAISSANCE'),
               GetControlText('LDEPTNAISSANCE'))
else
   CreeDetail (CleDADS, 17, 'S30.G01.00.011', '99', '99');

if (GetControlText('LPAYSNAISSANCE') <> '') then
   begin
   if ((TypeD='001') or (TypeD='201')) then
      CreeDetail (CleDADS, 18, 'S30.G01.00.012',
                  RechDom('TTPAYS', GetControlText('LPAYSNAISSANCE'), FALSE),
                  GetControlText('LPAYSNAISSANCE'))
   else
      begin
      PaysISOLib(GetControlText('LPAYSNAISSANCE'), CodeIso, Libelle);
      CreeDetail (CleDADS, 18, 'S30.G01.00.012', CodeIso,
                  GetControlText('LPAYSNAISSANCE'));
      end;
   end
else
   CreeDetail (CleDADS, 18, 'S30.G01.00.012', '', '');

if (GetControlText('LNATIONALITE') <> '') then
   begin
   if ((TypeD='001') or (TypeD='201')) then
      CreeDetail (CleDADS, 19, 'S30.G01.00.013',
                  RechDom('TTPAYS', GetControlText('LNATIONALITE'), FALSE),
                  GetControlText('LNATIONALITE'))
   else
      begin
      PaysISOLib(GetControlText('LNATIONALITE'), CodeIso, Libelle);
      CreeDetail (CleDADS, 19, 'S30.G01.00.013', CodeIso,
                  GetControlText('LNATIONALITE'));
      end;
   end
else
   CreeDetail (CleDADS, 19, 'S30.G01.00.013', '', '');

//PT73-9
CleDADS.Num:= Ordre;
CleDADS.DateDeb:= PerDeb;
CleDADS.DateFin:= PerFin;
//FIN PT73-9

{PT108
ForceNumerique(DateToStr(PerDeb), Buffer);
}
if (GetControlText ('EDEBMOTIF1')='095') then
   ForceNumerique (DateToStr (PerFin), Buffer)
else
   ForceNumerique (DateToStr (PerDeb), Buffer);
//FIN PT108
{PT93-4
CreeDetail (CleDADS, 1, 'S41.G01.00.001', Copy(Buffer, 1, 4),
            DateToStr(PerDeb));
}
CreeDetail (CleDADS, 1, 'S41.G01.00.001', Buffer, DateToStr (PerDeb));
//FIN PT93-4

{PT73-3
CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 2, 'S41.G01.00.002.001',
           GetControlText('CDEBMOTIF1'), GetControlText('CDEBMOTIF1'));

if (GetControlText('CDEBMOTIF2')<>'') then
   begin
   CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 3, 'S41.G01.00.002.002',
               GetControlText('CDEBMOTIF2'), GetControlText('CDEBMOTIF2'));
   if (GetControlText('CDEBMOTIF3')<>'') then
      CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 4, 'S41.G01.00.002.003',
                  GetControlText('CDEBMOTIF3'), GetControlText('CDEBMOTIF3'));
   end;
}
CreeDetail (CleDADS, 2, 'S41.G01.00.002.001', GetControlText('EDEBMOTIF1'),
            GetControlText('EDEBMOTIF1'));

{PT73-5
if (GetControlText('EDEBMOTIF2')<>'') then
}
if ((GetControlText('EDEBMOTIF2')<>'') and
   (GetControlText('EDEBMOTIF2')<>GetControlText('EDEBMOTIF1'))) then
   begin
   CreeDetail (CleDADS, 3, 'S41.G01.00.002.002', GetControlText('EDEBMOTIF2'),
               GetControlText('EDEBMOTIF2'));
{PT73-5
   if (GetControlText('EDEBMOTIF3')<>'') then
}
   if ((GetControlText('EDEBMOTIF3')<>'') and
      (GetControlText('EDEBMOTIF3')<>GetControlText('EDEBMOTIF1')) and
      (GetControlText('EDEBMOTIF3')<>GetControlText('EDEBMOTIF2'))) then
      CreeDetail (CleDADS, 4, 'S41.G01.00.002.003',
                  GetControlText('EDEBMOTIF3'), GetControlText('EDEBMOTIF3'));
   end;
//FIN PT73-3

ForceNumerique(DateToStr(PerFin), Buffer);
{PT93-4
CreeDetail (CleDADS, 7, 'S41.G01.00.003', Copy(Buffer, 1, 4),
            DateToStr(PerFin));
}
CreeDetail (CleDADS, 7, 'S41.G01.00.003', Buffer, DateToStr (PerFin));
//FIN PT93-4

{PT73-3
CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 8, 'S41.G01.00.004.001',
            GetControlText('CFINMOTIF1'), GetControlText('CFINMOTIF1'));

if (GetControlText('CFINMOTIF2')<>'') then
   begin
   CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 9, 'S41.G01.00.004.002',
               GetControlText('CFINMOTIF2'), GetControlText('CFINMOTIF2'));
   if (GetControlText('CFINMOTIF3')<>'') then
      CreeDetail (Sal, TypeD, Ordre, PerDeb, PerFin, 10, 'S41.G01.00.004.003',
                  GetControlText('CFINMOTIF3'), GetControlText('CFINMOTIF3'));
   end;
}
CreeDetail (CleDADS, 8, 'S41.G01.00.004.001', GetControlText('EFINMOTIF1'),
            GetControlText('EFINMOTIF1'));

{PT73-5
if (GetControlText('EFINMOTIF2')<>'') then
}
if ((GetControlText('EFINMOTIF2')<>'') and
   (GetControlText('EFINMOTIF2')<>GetControlText('EFINMOTIF1'))) then
   begin
   CreeDetail (CleDADS, 9, 'S41.G01.00.004.002', GetControlText('EFINMOTIF2'),
               GetControlText('EFINMOTIF2'));
{PT73-5
   if (GetControlText('EFINMOTIF3')<>'') then
}
   if ((GetControlText('EFINMOTIF3')<>'') and
      (GetControlText('EFINMOTIF3')<>GetControlText('EFINMOTIF1')) and
      (GetControlText('EFINMOTIF3')<>GetControlText('EFINMOTIF2'))) then
      CreeDetail (CleDADS, 10, 'S41.G01.00.004.003',
                  GetControlText('EFINMOTIF3'), GetControlText('EFINMOTIF3'));
   end;
//FIN PT73-3

StEtab := 'SELECT ET_SIRET'+
          ' FROM ETABLISS WHERE'+
          ' ET_ETABLISSEMENT = "'+GetControlText('CETABLISSEMENT')+'"';
QRechEtab:=OpenSQL(StEtab,TRUE);
if (not QRechEtab.EOF) then
   begin
   BufOrig:= QRechEtab.FindField ('ET_SIRET').Asstring;
   ForceNumerique (BufOrig, BufDest);
   if ControlSiret (BufDest)=False then
      PGIBox ('Le SIRET de l''établissement '+GetControlText ('CETABLISSEMENT')+
              ' n''est pas valide', 'Saisie complémentaire DADS-U');
//PT73-5
   ForceNumerique (GetParamSoc ('SO_SIRET'), SIREN);
   if (Copy (BufDest, 1, 9)<>Copy (SIREN, 1, 9)) then
      PGIBox ('Le SIREN de l''établissement '+GetControlText ('CETABLISSEMENT')+
              ' ne correspond pas au SIREN de l''entreprise',
              'Saisie complémentaire DADS-U');
//FIN PT73-5
   NIC:= Copy (BufDest, 10, 5);
   end
else
   begin
   PGIBox ('L''établissement n''existe pas', 'Saisie complémentaire DADS-U');
   SetFocusControl ('CETABLISSEMENT');
   end;
Ferme(QRechEtab);
CreeDetail (CleDADS, 13, 'S41.G01.00.005', NIC,
            GetControlText ('CETABLISSEMENT'));

if (MultiEmp <> NIL) then
   begin
   if (MultiEmp.Checked = FALSE) then
      CreeDetail (CleDADS, 14, 'S41.G01.00.008.001', '01', '01')
   else
      CreeDetail (CleDADS, 14, 'S41.G01.00.008.001', '02', '02');

{A traiter avec la gestion des carrières et des postes de travail
Pour l'instant, initialisé à inconnu}
{PT73-5
CreeDetail (CleDADS, 15, 'S41.G01.00.008.002', '01', '01');
}
CreeDetail (CleDADS, 15, 'S41.G01.00.008.002', '03', '03');
//FIN PT73-5           
   end;

//PT93-5
if (PerDeb>FinExer) then
   SetControlText ('DECALAGE', '05');
//FIN PT93-5

CreeDetail (CleDADS, 16, 'S41.G01.00.009', GetControlText('DECALAGE'),
            GetControlText('DECALAGE'));
if ((GetControlText('ELIBELLEEMPLOI') <> '') and
   (RechDom('PGLIBEMPLOI', GetControlText('ELIBELLEEMPLOI'), FALSE) <> '')) then
   CreeDetail (CleDADS, 17, 'S41.G01.00.010',
               RechDom('PGLIBEMPLOI', GetControlText('ELIBELLEEMPLOI'), FALSE),
               GetControlText('ELIBELLEEMPLOI'))
else
   CreeDetail (CleDADS, 17, 'S41.G01.00.010', '', '');

if (GetControlText('EPCS') <> '') then
   CreeDetail (CleDADS, 18, 'S41.G01.00.011', GetControlText('EPCS'),
               GetControlText('EPCS'));
//CCONTRAT
if (GetControlText('CCONTRAT') <> '') then
   begin
   if (GetControlText('CCONTRAT') = 'CAC') then
      SetControlText('CCONTRAT', 'CAP');
   CreeDetail (CleDADS, 19, 'S41.G01.00.012.001',
               RechDom('PGTYPECONTRAT', GetControlText('CCONTRAT'), TRUE),
               GetControlText('CCONTRAT'));
   end
else
   CreeDetail (CleDADS, 19, 'S41.G01.00.012.001', '01', '01');

if ((PrudDP <> NIL) and (PrudDP.Checked = TRUE)) then
   CreeDetail (CleDADS, 20, 'S41.G01.00.012.002', '02', '02')
else
   CreeDetail (CleDADS, 20, 'S41.G01.00.012.002', '01', '01');

{PT73-1
if (GetControlText('CCONDEMPLOI') = '') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '90',
               GetControlText('CCONDEMPLOI'))
else
}
if (GetControlText('CCONDEMPLOI') = 'C') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '01',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'P') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '02',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'I') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '04',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'D') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '05',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'S') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '06',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'V') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '07',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'O') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '08',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'N') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '09',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'F') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '10',
               GetControlText('CCONDEMPLOI'))
else
if (GetControlText('CCONDEMPLOI') = 'W') then
   CreeDetail (CleDADS, 21, 'S41.G01.00.013', '90',
               GetControlText('CCONDEMPLOI'));


if (GetControlText('CSTATUTPROF') <> '') then
{PT73-5
   begin
   if (GetControlText('CSTATUTPROF') = 'ZZZ') then
      SetControlText('CSTATUTPROF', '01');
}
   CreeDetail (CleDADS, 22, 'S41.G01.00.014', GetControlText('CSTATUTPROF'),
               GetControlText('CSTATUTPROF'));
{PT73-5
   end;
else
   CreeDetail (CleDADS, 22, 'S41.G01.00.014', '90', '90');
}

{PT73-5
if (GetControlText('CSTATUTCAT') <> '') then
   CreeDetail (CleDADS, 23, 'S41.G01.00.015', GetControlText('CSTATUTCAT'),
               GetControlText('CSTATUTCAT'));
else
   CreeDetail (CleDADS, 23, 'S41.G01.00.015', '90', '90');
}
if ((GetControlText('CSTATUTCAT') <> '') and
   (GetControlText('CSTATUTCAT') <> '90') and
   (GetControlText('ECONVCOLL') <> '') and
   (RechDom('PGIDCC', GetControlText('ECONVCOLL'), FALSE) <> '')) then
   begin
   if (GetControlText('CSTATUTCAT') = '01') then
      begin
      CreeDetail (CleDADS, 23, 'S41.G01.00.015.001',
                  GetControlText('CSTATUTCAT'), GetControlText('CSTATUTCAT'));
      CreeDetail (CleDADS, 24, 'S41.G01.00.015.002',
                  GetControlText('CSTATUTCAT'), GetControlText('CSTATUTCAT'));
      end
   else
      begin
      CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '02', '02');
      CreeDetail (CleDADS, 24, 'S41.G01.00.015.002',
                  GetControlText('CSTATUTCAT'), GetControlText('CSTATUTCAT'));
      end;
   end
else
   begin
   CreeDetail (CleDADS, 23, 'S41.G01.00.015.001', '90', '90');
   CreeDetail (CleDADS, 24, 'S41.G01.00.015.002', '04', '04');
   end;
//FIN PT73-5

if ((GetControlText('ECONVCOLL') <> '') and
   (RechDom('PGIDCC', GetControlText('ECONVCOLL'), FALSE) <> '')) then
   CreeDetail (CleDADS, 25, 'S41.G01.00.016', GetControlText('ECONVCOLL'),
               GetControlText('ECONVCOLL'));
{PT73-5
else
   CreeDetail (CleDADS, 24, 'S41.G01.00.016', '9999', '9999');
}

if ((GetControlText('ECOEFF') <> '') and
   (RechDom('PGCOEFFICIENT', GetControlText('ECOEFF'), FALSE) <> '')) then
   CreeDetail (CleDADS, 26, 'S41.G01.00.017',
               RechDom('PGCOEFFICIENT', GetControlText('ECOEFF'), FALSE),
               GetControlText('ECOEFF'))
else
   CreeDetail (CleDADS, 26, 'S41.G01.00.017', 'sans classement conventionnel',
               '');

//DEB PT92-1
if (GetControlText('CHTYPEREGIME') = '-') then
   CreeDetail (CleDADS, 27, 'S41.G01.00.018.001', GetControlText ('CREGIMESS'),
               GetControlText ('CREGIMESS'))
else
   begin
   CreeDetail (CleDADS, 28, 'S41.G01.00.018.002', GetControlText ('CREGIMEMAL'),
               GetControlText ('CREGIMEMAL'));
   CreeDetail (CleDADS, 29, 'S41.G01.00.018.003', GetControlText ('CREGIMEAT'),
               GetControlText ('CREGIMEAT'));
   CreeDetail (CleDADS, 30, 'S41.G01.00.018.004', GetControlText ('CREGIMEVIP'),
               GetControlText ('CREGIMEVIP'));
   CreeDetail (CleDADS, 31, 'S41.G01.00.018.005', GetControlText ('CREGIMEVIS'),
               GetControlText ('CREGIMEVIS'));
end;
//FIN PT92-1
//PT93-5
if ((Alsace<>NIL) and (Alsace.Checked=TRUE)) then
   CreeDetail (CleDADS, 32, 'S41.G01.00.018.006', '01', '01');
//FIN PT93-5

CreeDetail (CleDADS, 33, 'S41.G01.00.019', Sal, Sal);

BufOrig := GetControlText('ETAUXPARTIEL');
if ((BufOrig <> '0') and (BufOrig <> '') and
   (GetControlText('CCONDEMPLOI') = 'P')) then
   begin
   BufOrig := FloatToStr(Arrondi(StrToFloat(BufOrig), 2));
   BufDest := FloatToStr(StrToFloat(BufOrig)*100);
   CreeDetail (CleDADS, 34, 'S41.G01.00.020', BufDest, BufOrig);
   end;

{PT73-3
if (GetControlText('CDEBMOTIF1') = '095') then
}
if (GetControlText('EDEBMOTIF1') = '095') then
   SetControlText('EHEURESTRAV', '');
if ((GetControlText('EHEURESTRAV') <> '') and
   (Arrondi(StrToFloat(GetControlText('EHEURESTRAV')), 0)>0)) then
   CreeDetail (CleDADS, 35, 'S41.G01.00.021',
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESTRAV')), 0)),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESTRAV')), 0)));

if ((GetControlText('CSTATUTPROF') = '14') or
   (GetControlText('CSTATUTPROF') = '15') or
   (GetControlText('CSTATUTPROF') = '16') or
{PT73-3
   (GetControlText('CDEBMOTIF1') = '095')) then
}
   (GetControlText('EDEBMOTIF1') = '095')) then
   SetControlText('EHEURESSAL', '');

if ((GetControlText('EHEURESSAL') <> '') and
   (Arrondi(StrToFloat(GetControlText('EHEURESSAL')), 0)>0) and
   (GetControlText('CCONDEMPLOI')<>'D') and
   (GetControlText('CSTATUTPROF')<>'14') and
   (GetControlText('CSTATUTPROF')<>'15') and
   (GetControlText('CSTATUTPROF')<>'16')) then
   begin
   CreeDetail (CleDADS, 36, 'S41.G01.00.022',
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESSAL')), 0)),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESSAL')), 0)));
   if ((Arrondi(StrToFloat(GetControlText('EHEURESSAL')), 0) > 1200) and
{PT73-3
      (GetControlText('CFINMOTIF1') <> '008') and
      (GetControlText('CFINMOTIF1') <> '010') and
      (GetControlText('CFINMOTIF1') <> '012') and
      (GetControlText('CFINMOTIF1') <> '014') and
      (GetControlText('CFINMOTIF1') <> '016') and
      (GetControlText('CFINMOTIF1') <> '020')) then
}
      (GetControlText('EFINMOTIF1') <> '008') and
      (GetControlText('EFINMOTIF1') <> '010') and
      (GetControlText('EFINMOTIF1') <> '012') and
      (GetControlText('EFINMOTIF1') <> '014') and
      (GetControlText('EFINMOTIF1') <> '016') and
      (GetControlText('EFINMOTIF1') <> '020')) then
//FIN PT73-3
      SetControlText('EDERNIERMOIS', '');
   end;

if ((GetControlText('EDERNIERMOIS') <> '0') and
   (GetControlText('EDERNIERMOIS') <> '') and
   (StrToInt(GetControlText('EDERNIERMOIS')) <= 12)) then
   CreeDetail (CleDADS, 37, 'S41.G01.00.023', GetControlText('EDERNIERMOIS'),
               GetControlText('EDERNIERMOIS'));

if ((GetControlText('EHEURESCHOMPAR') <> '0') and
   (GetControlText('EHEURESCHOMPAR') <> '')) then
   CreeDetail (CleDADS, 38, 'S41.G01.00.024',
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESCHOMPAR')), 0)),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EHEURESCHOMPAR')), 0)));

{PT73-5
if ((GetControlText('CREGIMESS') <> '122') and
   (GetControlText('CREGIMESS') <> '200')) then
   CreeDetail (CleDADS, 33, 'S41.G01.00.025', '98', '98')
else
}
{PT92-1
if ((GetControlText('CREGIMESS') = '200') and
   (GetControlText('CSTATUTPROF') <> '07')) then
}
if ((((GetControlText ('CREGIMESS')='200') and
   (GetControlText ('CHTYPEREGIME')='-')) or
   ((GetControlText ('CREGIMEAT')='200') and
   (GetControlText ('CHTYPEREGIME')='X'))) and
   (GetControlText ('CSTATUTPROF')<>'07') and
   (GetControlText ('CSTATUTPROF')<>'36') and
   (GetControlText ('CSTATUTPROF')<>'37') and
   (GetControlText ('CSTATUTPROF')<>'38') and
   (GetControlText ('CSTATUTPROF')<>'39')) then
//FIN PT92-1   
   begin
   if ((GetControlText('ESECTIONAT') <> '') and
      (GetControlText('ESECTIONAT') <> '98')) then
      begin
      CreeDetail (CleDADS, 39, 'S41.G01.00.025', GetControlText('ESECTIONAT'),
                  GetControlText('ESECTIONAT'));
      if (GetControlText('ESECTIONAT')='99') then
         begin
         SetControlText('ERISQUEAT', '99999');
         SetControlText('ETAUXAT', FloatToStr(9999/100));
         end;
      end;
{PT73-5
   else
      CreeDetail (CleDADS, 33, 'S41.G01.00.025', '98', '98');
}
   end;

{PT73-5
if ((GetControlText('CREGIMESS') <> '122') and
   (GetControlText('CREGIMESS') <> '200')) then
   CreeDetail (CleDADS, 34, 'S41.G01.00.026', '98888', '98888')
else
}
Buffer:= '';
{PT92-1
if ((GetControlText('CREGIMESS') = '200') and
   (GetControlText('CSTATUTPROF') <> '07')) then
}
if ((((GetControlText ('CREGIMESS')='200') and
   (GetControlText ('CHTYPEREGIME')='-')) or
   ((GetControlText ('CREGIMEAT')='200') and
   (GetControlText ('CHTYPEREGIME')='X'))) and
   (GetControlText ('CSTATUTPROF')<>'07') and
   (GetControlText ('CSTATUTPROF')<>'36') and
   (GetControlText ('CSTATUTPROF')<>'37') and
   (GetControlText ('CSTATUTPROF')<>'38') and
   (GetControlText ('CSTATUTPROF')<>'39')) then
   begin
   Buffer:= PGUpperCase(GetControlText('ERISQUEAT'));
   if (GetControlText('ERISQUEAT') <> '') then
      CreeDetail (CleDADS, 40, 'S41.G01.00.026', Buffer, Buffer);
{PT73-5
   else
      CreeDetail (CleDADS, 34, 'S41.G01.00.026', '98888', '98888');
}
   end;

{PT73-5
if ((CodeBureau <> NIL) and (CodeBureau.Checked = TRUE)) then
}
if ((CodeBureau <> NIL) and (CodeBureau.Checked = TRUE) and (Buffer <> '') and
   (Buffer <> '753CB') and (Buffer <> '911AA') and (Buffer <> '753CA') and
   (Buffer <> '401ZA') and (Buffer <> '753CC') and (Buffer <> '631AZ') and
   (Buffer <> '926CC') and (Buffer <> '926CF') and (Buffer <> '631AB') and
   (Buffer <> '923AC') and (Buffer <> '926CE') and (Buffer <> '926CD') and
   (Buffer <> '401ZB') and (Buffer <> '802CA') and (Buffer <> '751CA') and
   (Buffer <> '802AA') and (Buffer <> '853KA') and (Buffer <> '752EC') and
   (Buffer <> '853CA') and (Buffer <> '752EB') and (Buffer <> '804CA') and
   (Buffer <> '752EA') and (Buffer <> '950ZD') and (Buffer <> '950ZC') and
   (Buffer <> '950ZB') and (Buffer <> '511TG') and (Buffer <> '950ZA') and
   (Buffer <> '913EG') and (Buffer <> '913EF') and (Buffer <> '526GA') and
   (Buffer <> '913EE') and (Buffer <> '913ED') and (Buffer <> '913EC') and
   (Buffer <> '853KC') and (Buffer <> '853KB') and (Buffer <> '752ED') and
{PT93-5
   (Buffer <> '745AB') and (Buffer <> '999ZA') and (Buffer <> '524RB') and
   (Buffer <> '853HB') and (Buffer <> '853HA') and (Buffer <> '745BB') and
   (Buffer <> '745BD') and (Buffer <> '911AE') and (Buffer <> '745BA') and
   (Buffer <> '853KD') and (Buffer <> '511TH') and (Buffer <> '853KE')) then
}
   (Buffer <> '745AB') and (Buffer <> '524RB') and (Buffer <> '853HB') and
   (Buffer <> '853HA') and (Buffer <> '745BB') and (Buffer <> '745BD') and
   (Buffer <> '911AE') and (Buffer <> '745BA') and (Buffer <> '853KD') and
   (Buffer <> '511TH') and (Buffer <> '853KE')) then
//FIN PT93-5
   CreeDetail (CleDADS, 41, 'S41.G01.00.027', 'B', 'B');

{PT73-5
if ((GetControlText('CREGIMESS') <> '122') and
   (GetControlText('CREGIMESS') <> '200')) then
   CreeDetail (CleDADS, 36, 'S41.G01.00.028', '98888', FloatToStr(9888/100))
else
}
{PT92-1
if ((GetControlText('CREGIMESS') = '200') and
   (GetControlText('CSTATUTPROF') <> '07')) then
}
if ((((GetControlText ('CREGIMESS')='200') and
   (GetControlText ('CHTYPEREGIME')='-')) or
   ((GetControlText ('CREGIMEAT')='200') and
   (GetControlText ('CHTYPEREGIME')='X'))) and
   (GetControlText ('CSTATUTPROF')<>'07') and
   (GetControlText ('CSTATUTPROF')<>'36') and
   (GetControlText ('CSTATUTPROF')<>'37') and
   (GetControlText ('CSTATUTPROF')<>'38') and
   (GetControlText ('CSTATUTPROF')<>'39')) then
   begin
   if ((GetControlText('ETAUXAT') <> '') and
      (GetControlText('ETAUXAT') <> FloatToStr(9888/100))) then
      begin
      if (GetControlText('ETAUXAT') <> FloatToStr(9999/100)) then
         CreeDetail (CleDADS, 42, 'S41.G01.00.028',
                     FloatToStr(Arrondi(StrToFloat(GetControlText('ETAUXAT'))*100, 0)),
                     FloatToStr(Arrondi(StrToFloat(GetControlText('ETAUXAT')), 2)))
      else
         CreeDetail (CleDADS, 42, 'S41.G01.00.028', '99999',
                     FloatToStr(9999/100));
      end;
{PT73-5
   else
      CreeDetail (CleDADS, 36, 'S41.G01.00.028', '98888', FloatToStr(9888/100));
}
   end;

{PT93-5
if ((GetControlText('EBRUTSS') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
}
if ((GetControlText ('EBRUTSS') <> '') and
   (GetControlText ('CCONTRAT') <> 'CAA') and
   (GetControlText ('CSTATUTPROF')<>'07') and
   (GetControlText ('CSTATUTPROF')<>'36')) then
//FIN PT93-5
   begin
   CreeDetail (CleDADS, 43, 'S41.G01.00.029.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBRUTSS')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EBRUTSS')), 0)));
   if (StrToFloat(GetControlText('EBRUTSS')) < 0) then
      CreeDetail (CleDADS, 44, 'S41.G01.00.029.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 43, 'S41.G01.00.029.001', '0', '0');

//PT93-5
{PT107-1
if ((BaseReelle<>NIL) and (BaseReelle.Checked=TRUE)) then
}
if ((BaseReelle<>NIL) and (BaseReelle.Checked=TRUE) and
   (GetControlText ('CCONTRAT')<>'CAA') and
   (GetControlText ('CCONTRAT')<>'CAP')) then
//FIN PT107-1
   CreeDetail (CleDADS, 45, 'S41.G01.00.029.003', '01', '01')
else
   CreeDetail (CleDADS, 45, 'S41.G01.00.029.003', '02', '02');
//FIN PT93-5

{PT93-5
if ((GetControlText('EPLAFONDSS') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
}
if ((GetControlText ('EPLAFONDSS') <> '') and
   (GetControlText ('CCONTRAT') <> 'CAA') and
   (GetControlText ('CSTATUTPROF')<>'07') and
   (GetControlText ('CSTATUTPROF')<>'36')) then
//FIN PT93-5
   begin
   CreeDetail (CleDADS, 46, 'S41.G01.00.030.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EPLAFONDSS')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EPLAFONDSS')), 0)));
   if (StrToFloat(GetControlText('EPLAFONDSS')) < 0) then
      CreeDetail (CleDADS, 47, 'S41.G01.00.030.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 46, 'S41.G01.00.030.001', '0', '0');


if (GetControlText('EBASECSG') <> '') then
   begin
   CreeDetail (CleDADS, 48, 'S41.G01.00.032.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBASECSG')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EBASECSG')), 0)));
   if (StrToFloat(GetControlText('EBASECSG')) < 0) then
      CreeDetail (CleDADS, 49, 'S41.G01.00.032.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 48, 'S41.G01.00.032.001', '0', '0');

if (GetControlText('EBASECRDS') <> '') then
   begin
   CreeDetail (CleDADS, 50, 'S41.G01.00.033.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBASECRDS')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('EBASECRDS')), 0)));
   if (StrToFloat(GetControlText('EBASECRDS')) < 0) then
      CreeDetail (CleDADS, 51, 'S41.G01.00.033.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 50, 'S41.G01.00.033.001', '0', '0');

if (GetControlText('CTRAVETRANG') = 'F') then
   CreeDetail (CleDADS, 52, 'S41.G01.00.034', '01',
               GetControlText('CTRAVETRANG'))
else
if (GetControlText('CTRAVETRANG') = 'E') then
   CreeDetail (CleDADS, 52, 'S41.G01.00.034', '02',
               GetControlText('CTRAVETRANG'));

{PT87
if ((GetControlText('EBRUTFISC') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
}
if ((GetControlText('EBRUTFISC') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA') and
   (GetControlText('CCONTRAT') <> 'CAU') and
   (GetControlText('CCONTRAT') <> 'CAV') and
   (GetControlText('CCONTRAT') <> 'CEJ') and
   (GetControlText('CCONTRAT') <> 'CES')) then
//FIN PT87
   begin
   CreeDetail (CleDADS, 53, 'S41.G01.00.035.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EBRUTFISC'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EBRUTFISC')))));
   if (StrToFloat(GetControlText('EBRUTFISC')) < 0) then
      CreeDetail (CleDADS, 54, 'S41.G01.00.035.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 53, 'S41.G01.00.035.001', '0', '0');

if ((GetControlText('EAVANTAGENAT') <> '0') and
   (GetControlText('EAVANTAGENAT') <> '')) then
   begin
   CreeDetail (CleDADS, 55, 'S41.G01.00.037.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EAVANTAGENAT'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EAVANTAGENAT')))));
   if (StrToFloat(GetControlText('EAVANTAGENAT')) < 0) then
      CreeDetail (CleDADS, 56, 'S41.G01.00.037.002', 'N', 'N');

   if ((Nourriture <> NIL) and (Nourriture.Checked = TRUE)) then
      CreeDetail (CleDADS, 57, 'S41.G01.00.038', 'N', 'N');

   if ((Logement <> NIL) and (Logement.Checked = TRUE)) then
      CreeDetail (CleDADS, 58, 'S41.G01.00.039', 'L', 'L');

   if ((Voiture <> NIL) and (Voiture.Checked = TRUE)) then
      CreeDetail (CleDADS, 59, 'S41.G01.00.040', 'V', 'V');

   if ((AutreAvant <> NIL) and (AutreAvant.Checked = TRUE)) then
      CreeDetail (CleDADS, 60, 'S41.G01.00.041', 'A', 'A');

   if ((NTIC <> NIL) and (NTIC.Checked = TRUE)) then
      CreeDetail (CleDADS, 92, 'S41.G01.00.068', 'T', 'T');
   end;

if ((GetControlText('ERETSALAIRE') <> '0') and
   (GetControlText('ERETSALAIRE') <> '')) then
   begin
   CreeDetail (CleDADS, 61, 'S41.G01.00.042.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('ERETSALAIRE'))))),
               FloatToStr(Int(StrToFloat(GetControlText('ERETSALAIRE')))));
   if (StrToFloat(GetControlText('ERETSALAIRE')) < 0) then
      CreeDetail (CleDADS, 62, 'S41.G01.00.042.002', 'N', 'N');
   end;

if ((RemPourboire <> NIL) and (RemPourboire.Checked = TRUE)) then
   CreeDetail (CleDADS, 63, 'S41.G01.00.043', 'P', 'P');

if ((GetControlText('EFRAISPROF') <> '0') and
   (GetControlText('EFRAISPROF') <> '')) then
   begin
   CreeDetail (CleDADS, 64, 'S41.G01.00.044.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EFRAISPROF'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EFRAISPROF')))));
   if (StrToFloat(GetControlText('EFRAISPROF')) < 0) then
      CreeDetail (CleDADS, 65, 'S41.G01.00.044.002', 'N', 'N');

   if ((AllocForfait <> NIL) and (AllocForfait.Checked = TRUE)) then
      CreeDetail (CleDADS, 66, 'S41.G01.00.045', 'F', 'F');

   if ((RembJustif <> NIL) and (RembJustif.Checked = TRUE)) then
      CreeDetail (CleDADS, 67, 'S41.G01.00.046', 'R', 'R');

   if ((PriseCharge <> NIL) and (PriseCharge.Checked = TRUE)) then
      CreeDetail (CleDADS, 68, 'S41.G01.00.047', 'P', 'P');

   if ((AutreDepense <> NIL) and (AutreDepense.Checked = TRUE)) then
      CreeDetail (CleDADS, 69, 'S41.G01.00.048', 'D', 'D');
   end;

if ((GetControlText('ECHQVACANCE') <> '0') and
   (GetControlText('ECHQVACANCE') <> '')) then
   begin
   CreeDetail (CleDADS, 70, 'S41.G01.00.049.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('ECHQVACANCE'))))),
               FloatToStr(Int(StrToFloat(GetControlText('ECHQVACANCE')))));
   if (StrToFloat(GetControlText('ECHQVACANCE')) < 0) then
      CreeDetail (CleDADS, 71, 'S41.G01.00.049.002', 'N', 'N');
   end;

if ((GetControlText('EIMPOTRETSOURC') <> '0') and
   (GetControlText('EIMPOTRETSOURC') <> '')) then
   begin
   CreeDetail (CleDADS, 72, 'S41.G01.00.052.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EIMPOTRETSOURC'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EIMPOTRETSOURC')))));
   if (StrToFloat(GetControlText('EIMPOTRETSOURC')) < 0) then
      CreeDetail (CleDADS, 73, 'S41.G01.00.052.002', 'N', 'N');
   end;

if ((GetControlText('EINDEXPATRIAT') <> '0') and
   (GetControlText('EINDEXPATRIAT') <> '')) then
   begin
   CreeDetail (CleDADS, 74, 'S41.G01.00.053.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EINDEXPATRIAT'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EINDEXPATRIAT')))));
   if (StrToFloat(GetControlText('EINDEXPATRIAT')) < 0) then
      CreeDetail (CleDADS, 75, 'S41.G01.00.053.002', 'N', 'N');
   end;

if ((GetControlText('ETOTIMPOTAXE') <> '0') and
   (GetControlText('ETOTIMPOTAXE') <> '')) then
   begin
   CreeDetail (CleDADS, 76, 'S41.G01.00.055.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('ETOTIMPOTAXE'))))),
               FloatToStr(Int(StrToFloat(GetControlText('ETOTIMPOTAXE')))));
   if (StrToFloat(GetControlText('ETOTIMPOTAXE')) < 0) then
      CreeDetail (CleDADS, 77, 'S41.G01.00.055.002', 'N', 'N');
   end;

if ((GetControlText('EBASEIMPO1') <> '0') and
   (GetControlText('EBASEIMPO1') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
   begin
   CreeDetail (CleDADS, 78, 'S41.G01.00.056.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EBASEIMPO1'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EBASEIMPO1')))));
   if (StrToFloat(GetControlText('EBASEIMPO1')) < 0) then
      CreeDetail (CleDADS, 79, 'S41.G01.00.056.002', 'N', 'N');
   end;

if ((GetControlText('EBASEIMPO2') <> '0') and
   (GetControlText('EBASEIMPO2') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
   begin
   CreeDetail (CleDADS, 80, 'S41.G01.00.057.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EBASEIMPO2'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EBASEIMPO2')))));
   if (StrToFloat(GetControlText('EBASEIMPO2')) < 0) then
      CreeDetail (CleDADS, 81, 'S41.G01.00.057.002', 'N', 'N');
   end;

if ((GetControlText('ETAXESAL') <> '0') and
   (GetControlText('ETAXESAL') <> '') and
   (GetControlText('CCONTRAT') <> 'CAA')) then
   begin
   CreeDetail (CleDADS, 82, 'S41.G01.00.058.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('ETAXESAL'))))),
               FloatToStr(Int(StrToFloat(GetControlText('ETAXESAL')))));
   if (StrToFloat(GetControlText('ETAXESAL')) < 0) then
      CreeDetail (CleDADS, 83, 'S41.G01.00.058.002', 'N', 'N');
   end;

if (GetControlText('EREVENUSACTIV') <> '') then
   begin
   CreeDetail (CleDADS, 84, 'S41.G01.00.063.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EREVENUSACTIV'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EREVENUSACTIV')))));
   if (StrToFloat(GetControlText('EREVENUSACTIV')) < 0) then
      CreeDetail (CleDADS, 85, 'S41.G01.00.063.002', 'N', 'N');
   end
else
   CreeDetail (CleDADS, 84, 'S41.G01.00.063.001', '0', '0');

if ((GetControlText('ECP') <> '0') and (GetControlText('ECP') <> '')) then
   begin
   CreeDetail (CleDADS, 86, 'S41.G01.00.065.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('ECP')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('ECP')), 0)));
   if (StrToFloat(GetControlText('ECP')) < 0) then
      CreeDetail (CleDADS, 87, 'S41.G01.00.065.002', 'N', 'N');
   end;

if ((GetControlText('EAUTREREV') <> '0') and
   (GetControlText('EAUTREREV') <> '')) then
   begin
   CreeDetail (CleDADS, 88, 'S41.G01.00.066.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EAUTREREV'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EAUTREREV')))));
   if (StrToFloat(GetControlText('EAUTREREV')) < 0) then
      CreeDetail (CleDADS, 89, 'S41.G01.00.066.002', 'N', 'N');
   end;

if ((GetControlText('EEPARGNE') <> '0') and
   (GetControlText('EEPARGNE') <> '')) then
   begin
   CreeDetail (CleDADS, 90, 'S41.G01.00.067.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EEPARGNE'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EEPARGNE')))));
   if (StrToFloat(GetControlText('EEPARGNE')) < 0) then
      CreeDetail (CleDADS, 91, 'S41.G01.00.067.002', 'N', 'N');
   end;

if ((GetControlText('EINDIMPATRIAT') <> '0') and
   (GetControlText('EINDIMPATRIAT') <> '')) then
   begin
   CreeDetail (CleDADS, 93, 'S41.G01.00.069.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('EINDIMPATRIAT'))))),
               FloatToStr(Int(StrToFloat(GetControlText('EINDIMPATRIAT')))));
   if (StrToFloat(GetControlText('EINDIMPATRIAT')) < 0) then
      CreeDetail (CleDADS, 94, 'S41.G01.00.069.002', 'N', 'N');
   end;

//PT73-5
if ((GetControlText('ESERVICEPERS') <> '0') and
   (GetControlText('ESERVICEPERS') <> '')) then
   begin
   CreeDetail (CleDADS, 95, 'S41.G01.00.070.001',
               FloatToStr(Abs(Int(StrToFloat(GetControlText('ESERVICEPERS'))))),
               FloatToStr(Int(StrToFloat(GetControlText('ESERVICEPERS')))));
   if (StrToFloat(GetControlText('ESERVICEPERS')) < 0) then
      CreeDetail (CleDADS, 96, 'S41.G01.00.070.002', 'N', 'N');
   end;

//PT93-5
if (GetControlText('CCNBF') <> '') then
   CreeDetail (CleDADS, 97, 'S41.G01.00.071', GetControlText('CCNBF'),
               GetControlText('CCNBF'));
//FIN PT93-5

if ((GetControlText('ECPPLAF') <> '0') and
   (GetControlText('ECPPLAF') <> '')) then
   begin
   CreeDetail (CleDADS, 98, 'S41.G01.00.072.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('ECPPLAF')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('ECPPLAF')), 0)));
   if (StrToFloat(GetControlText('ECPPLAF')) < 0) then
      CreeDetail (CleDADS, 99, 'S41.G01.00.072.002', 'N', 'N');
   end;
//FIN PT73-5

//PT93-5
if ((GetControlText ('EHEURESSUPP')<>'0') and
   (GetControlText ('EHEURESSUPP')<>'')) then
   begin
   CreeDetail (CleDADS, 100, 'S41.G01.00.073.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EHEURESSUPP')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EHEURESSUPP')), 0)));
   if (StrToFloat (GetControlText ('EHEURESSUPP'))<0) then
      CreeDetail (CleDADS, 101, 'S41.G01.00.073.002', 'N', 'N');
   end;
//FIN PT93-5

Ircantec:= False;          //PT90
if (RegimeComp <> NIL) then
   begin
   j := 110;
   for i := 1 to 5 do
       begin
       if ((RegimeComp.CellValues[0,i] <> '') and
          ((not(isnumeric(RegimeComp.CellValues[0,i]))) or
          (RegimeComp.CellValues[0,i] = '90000'))) then
          begin
{PT77-1
          if ((TypeD>='002') and (TypeD<='005') and
             (Pos('B', RegimeComp.CellValues[0,i])=1)) then
             Buffer:= Copy(RegimeComp.CellValues[0,i], 3, 2)
          else
//PT73-8
          if (Pos('I', RegimeComp.CellValues[0,i])=1) then
             Buffer:= Copy(RegimeComp.CellValues[0,i], 1, 1)+'0'+
                      Copy(RegimeComp.CellValues[0,i], 2, 3)
          else
//FIN PT73-8
             Buffer:= RegimeComp.CellValues[0,i];
          CreeDetail (CleDADS, j+(i*2)-1, 'S41.G01.01.001', Buffer,
                      RegimeComp.CellValues[0,i]);
}
          if (((TypeD>='002') and (TypeD<='005')) or
             ((TypeD>='202') and (TypeD<='205'))) then
             begin
             if (Pos ('B', RegimeComp.CellValues[0,i])=1) then
                begin
                Buffer:= Copy(RegimeComp.CellValues[0,i], 3, 2);
                CreeDetail (CleDADS, j+(i*2)-1, 'S41.G01.01.001', Buffer,
                            RegimeComp.CellValues[0,i]);
//PT86
                if (RegimeComp.CellValues[2,i] <> '') then
                   CreeDetail (CleDADS, j+(i*2), 'S41.G01.01.002',
                               RegimeComp.CellValues[2,i],
                               RegimeComp.CellValues[2,i]);
//FIN PT86
                end;
             end
          else
             begin
             if (Pos ('B', RegimeComp.CellValues[0,i])<>1) then
                begin
//PT73-8
                if (Pos('I', RegimeComp.CellValues[0,i])=1) then
                   begin
                   Buffer:= Copy(RegimeComp.CellValues[0,i], 1, 1)+'0'+
                            Copy(RegimeComp.CellValues[0,i], 2, 3);
                   Ircantec:= True;
                   end
                else
//FIN PT73-8
                   Buffer:= RegimeComp.CellValues[0,i];
                CreeDetail (CleDADS, j+(i*2)-1, 'S41.G01.01.001', Buffer,
                            RegimeComp.CellValues[0,i]);
//PT86
                if (RegimeComp.CellValues[2,i] <> '') then
                   CreeDetail (CleDADS, j+(i*2), 'S41.G01.01.002',
                               RegimeComp.CellValues[2,i],
                               RegimeComp.CellValues[2,i]);
//FIN PT86
                end;
             end;
//FIN PT77-1

{PT86
          if (RegimeComp.CellValues[2,i] <> '') then
             CreeDetail (CleDADS, j+(i*2), 'S41.G01.01.002',
                         RegimeComp.CellValues[2,i],
                         RegimeComp.CellValues[2,i]);
}
          end;
       end;
   end;

if (Basebrutespec <> NIL) then
   begin
   j := 200;
   for i := 1 to 4 do
       begin
       if ((Basebrutespec.CellValues[0,i] <> '') and
          (Basebrutespec.CellValues[2,i] <> '') and
          (Basebrutespec.CellValues[2,i] <> '0') and
          (isnumeric(Basebrutespec.CellValues[2,i]))) then
          begin
          CreeDetail (CleDADS, j+(i*3)-2, 'S41.G01.02.001',
                      Basebrutespec.CellValues[0,i],
                      Basebrutespec.CellValues[0,i]);
          CreeDetail (CleDADS, j+(i*3)-1, 'S41.G01.02.002.001',
                      FloatToStr(Abs(Arrondi(StrToFloat(Basebrutespec.CellValues[2,i]), 0))),
                      FloatToStr(Arrondi(StrToFloat(Basebrutespec.CellValues[2,i]), 0)));
          if (StrToFloat(Basebrutespec.CellValues[2,i]) < 0) then
             CreeDetail (CleDADS, j+(i*3), 'S41.G01.02.002.002', 'N', 'N');
          end
       else
          begin
          Basebrutespec.CellValues[0,i]:= '';
          Basebrutespec.CellValues[1,i]:= '';
          Basebrutespec.CellValues[2,i]:= '';
          end;
       end;
   end;

if (Baseplafspec <> NIL) then
   begin
   j := 300;
   for i := 1 to 4 do
       begin
       if ((Baseplafspec.CellValues[0,i] <> '') and
          (Baseplafspec.CellValues[2,i] <> '') and
          (Baseplafspec.CellValues[2,i] <> '0') and
          (isnumeric(Baseplafspec.CellValues[2,i]))) then
          begin
          CreeDetail (CleDADS, j+(i*3)-2, 'S41.G01.03.001',
                      Baseplafspec.CellValues[0,i],
                      Baseplafspec.CellValues[0,i]);
          CreeDetail (CleDADS, j+(i*3)-1, 'S41.G01.03.002.001',
                      FloatToStr(Abs(Arrondi(StrToFloat(Baseplafspec.CellValues[2,i]), 0))),
                      FloatToStr(Arrondi(StrToFloat(Baseplafspec.CellValues[2,i]), 0)));
          if (StrToFloat(Baseplafspec.CellValues[2,i]) < 0) then
             CreeDetail (CleDADS, j+(i*3), 'S41.G01.03.002.002', 'N', 'N');
          end
       else
          begin
          Baseplafspec.CellValues[0,i]:= '';
          Baseplafspec.CellValues[1,i]:= '';
          Baseplafspec.CellValues[2,i]:= '';
          end;
       end;
   end;

if (Somiso <> NIL) then
   begin
   DecodeDate(FinExer, AnneeA, MoisM, JourJ);
   j := 400;
   for i := 1 to 4 do
       begin
       if ((Somiso.CellValues[0,i] <> '') and
          (Somiso.CellValues[2,i]<=IntToStr(AnneeA))) then
          begin
          CreeDetail (CleDADS, j+(i*6)-5, 'S41.G01.04.001',
                      Somiso.CellValues[0,i], Somiso.CellValues[0,i]);
          CreeDetail (CleDADS, j+(i*6)-4, 'S41.G01.04.002',
                      Somiso.CellValues[2,i], Somiso.CellValues[2,i]);

          if ((Somiso.CellValues[3,i] <> '') and
             (isnumeric(Somiso.CellValues[3,i]))) then
             begin
             CreeDetail (CleDADS, j+(i*6)-3, 'S41.G01.04.003.001',
                         FloatToStr(Abs(Arrondi(StrToFloat(Somiso.CellValues[3,i]), 0))),
                         FloatToStr(Arrondi(StrToFloat(Somiso.CellValues[3,i]), 0)));
             if (StrToFloat(Somiso.CellValues[3,i]) < 0) then
                CreeDetail (CleDADS, j+(i*6)-2, 'S41.G01.04.003.002', 'N', 'N');
             end
          else
             begin
             CreeDetail (CleDADS, j+(i*6)-3, 'S41.G01.04.003.001', '0','0');
             Somiso.CellValues[3,i]:= '0';
             end;

          if ((Somiso.CellValues[4,i] <> '') and
             (isnumeric(Somiso.CellValues[4,i]))) then
             begin
             CreeDetail (CleDADS, j+(i*6)-1, 'S41.G01.04.004.001',
                         FloatToStr(Abs(Arrondi(StrToFloat(Somiso.CellValues[4,i]), 0))),
                         FloatToStr(Arrondi(StrToFloat(Somiso.CellValues[4,i]), 0)));
             if (StrToFloat(Somiso.CellValues[4,i]) < 0) then
                CreeDetail (CleDADS, j+(i*6), 'S41.G01.04.004.002', 'N', 'N');
             end
          else
             begin
             CreeDetail (CleDADS, j+(i*6)-1, 'S41.G01.04.004.001', '0','0');
             Somiso.CellValues[4,i]:= '0';
             end;
          end;
       end;
   end;

if (GetControlText('ECSGSPECM') <> '') then
   begin
   CreeDetail (CleDADS, 501, 'S41.G01.05.001', '01', '01');

   CreeDetail (CleDADS, 502, 'S41.G01.05.002.001',
               FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('ECSGSPECM')), 0))),
               FloatToStr(Arrondi(StrToFloat(GetControlText('ECSGSPECM')), 0)));
   if (StrToFloat(GetControlText('ECSGSPECM')) < 0) then
      CreeDetail (CleDADS, 503, 'S41.G01.05.002.002', 'N', 'N');
   end;

if (Exoner <> NIL) then
   begin
   j := 600;
   for i := 1 to 10 do
       begin
       if ((Exoner.CellValues[0,i] <> '') and
          (Exoner.CellValues[2,i] <> '') and
          (Exoner.CellValues[2,i] <> '0') and   // pt97
          (isnumeric(Exoner.CellValues[2,i]))) then
          begin
          CreeDetail (CleDADS, j+(i*5)-4, 'S41.G01.06.001',
                      Exoner.CellValues[0,i], Exoner.CellValues[0,i]);

          CreeDetail (CleDADS, j+(i*5)-3, 'S41.G01.06.002.001',
                      FloatToStr(Abs(Arrondi(StrToFloat(Exoner.CellValues[2,i]), 0))),
                      FloatToStr(Arrondi(StrToFloat(Exoner.CellValues[2,i]), 0)));
          if (StrToFloat(Exoner.CellValues[2,i]) < 0) then
             CreeDetail (CleDADS, j+(i*5)-2, 'S41.G01.06.002.002', 'N', 'N');

          if ((Exoner.CellValues[3,i] <> '') and
              (Exoner.CellValues[3,i] <> '0') and // pt97
             (isnumeric(Exoner.CellValues[3,i]))) then
             begin
             CreeDetail (CleDADS, j+(i*5)-1, 'S41.G01.06.003.001',
                         FloatToStr(Abs(Arrondi(StrToFloat(Exoner.CellValues[3,i]), 0))),
                         FloatToStr(Arrondi(StrToFloat(Exoner.CellValues[3,i]), 0)));
             if (StrToFloat(Exoner.CellValues[3,i]) < 0) then
                CreeDetail (CleDADS, j+(i*5), 'S41.G01.06.003.002', 'N', 'N');
             end
          else
             Exoner.CellValues[3,i]:= '';
          end
       else
          begin
             Exoner.CellValues[0,i]:= '';
             Exoner.CellValues[1,i]:= '';
             Exoner.CellValues[2,i]:= '';
             Exoner.CellValues[3,i]:= '';
          end;
       end;
   end;

{PT80
if (Nbre = NbreTot) then
}
if ((Nbre = NbreTot) and (PrudDP <> NIL) and (PrudDP.Checked <> TRUE)) then
   begin
   CleDADS.DateDeb:= PerFin;	//PT73-9
   if ((PrudPresent <> NIL) and (PrudPresent.Checked = TRUE)) then
      begin
      CreeDetail (CleDADS, 701, 'S41.G02.00.008', '01', '01');
//PT73-5
      if (GetControlText('CPRUDHCOLL') <> '') then
         begin
         CreeDetail (CleDADS, 702, 'S41.G02.00.009',
                     '0'+GetControlText('CPRUDHCOLL'),
                     GetControlText('CPRUDHCOLL'));

         CreeDetail (CleDADS, 703, 'S41.G02.00.010',
                     '0'+GetControlText('CPRUDHSECT'),
                     GetControlText('CPRUDHSECT'));
         end;
//FIN PT73-5
      end
   else
      CreeDetail (CleDADS, 701, 'S41.G02.00.008', '02', '02');
   end;

CleDADS.DateDeb:= PerDeb;	//PT73-9

//PT93-5
if ((GetControlText ('EEPARGNESALP')<>'') and
   (GetControlText ('EEPARGNESALP')<>'0')) then
   begin
   CreeDetail (CleDADS, 801, 'S41.G30.10.001', '01', '01');

   CreeDetail (CleDADS, 802, 'S41.G30.10.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEPARGNESALP')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEPARGNESALP')), 0)));
   if (StrToFloat(GetControlText ('EEPARGNESALP'))<0) then
                CreeDetail (CleDADS, 803, 'S41.G30.10.002.002', 'N', 'N');
   end;

if ((GetControlText ('EEPARGNESALI')<>'') and
   (GetControlText ('EEPARGNESALI')<>'0')) then
   begin
   CreeDetail (CleDADS, 804, 'S41.G30.10.001', '02', '02');

   CreeDetail (CleDADS, 805, 'S41.G30.10.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEPARGNESALI')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEPARGNESALI')), 0)));
   if (StrToFloat(GetControlText ('EEPARGNESALI'))<0) then
                CreeDetail (CleDADS, 806, 'S41.G30.10.002.002', 'N', 'N');
   end;

if ((GetControlText ('EEPARGNESALA')<>'') and
   (GetControlText ('EEPARGNESALA')<>'0')) then
   begin
   CreeDetail (CleDADS, 807, 'S41.G30.10.001', '03', '03');

   CreeDetail (CleDADS, 808, 'S41.G30.10.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEPARGNESALA')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEPARGNESALA')), 0)));
   if (StrToFloat(GetControlText ('EEPARGNESALA'))<0) then
                CreeDetail (CleDADS, 809, 'S41.G30.10.002.002', 'N', 'N');
   end;

if ((GetControlText ('EEPARGNESALD')<>'') and
   (GetControlText ('EEPARGNESALD')<>'0')) then
   begin
   CreeDetail (CleDADS, 810, 'S41.G30.10.001', '04', '04');

   CreeDetail (CleDADS, 811, 'S41.G30.10.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEPARGNESALD')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEPARGNESALD')), 0)));
   if (StrToFloat(GetControlText ('EEPARGNESALD'))<0) then
                CreeDetail (CleDADS, 812, 'S41.G30.10.002.002', 'N', 'N');
   end;

if ((GetControlText ('EACTIONSN1')<>'') and
   (GetControlText ('EACTIONSN1')>'0') and
   (GetControlText ('EACTIONSV1')<>'') and
   (GetControlText ('EACTIONSV1')>'0')) then
   begin
   CreeDetail (CleDADS, 821, 'S41.G30.11.001.001',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN1')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN1')), 0)));

   CreeDetail (CleDADS, 822, 'S41.G30.11.001.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV1'))*100, 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV1')), 2)));
   end;

if ((GetControlText ('EACTIONSN2')<>'') and
   (GetControlText ('EACTIONSN2')>'0') and
   (GetControlText ('EACTIONSV2')<>'') and
   (GetControlText ('EACTIONSV2')>'0')) then
   begin
   CreeDetail (CleDADS, 823, 'S41.G30.11.001.001',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN2')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN2')), 0)));

   CreeDetail (CleDADS, 824, 'S41.G30.11.001.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV2'))*100, 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV2')), 2)));
   end;

if ((GetControlText ('EACTIONSN3')<>'') and
   (GetControlText ('EACTIONSN3')>'0') and
   (GetControlText ('EACTIONSV3')<>'') and
   (GetControlText ('EACTIONSV3')>'0')) then
   begin
   CreeDetail (CleDADS, 825, 'S41.G30.11.001.001',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN3')), 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSN3')), 0)));

   CreeDetail (CleDADS, 826, 'S41.G30.11.001.002',
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV3'))*100, 0)),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EACTIONSV3')), 2)));
   end;

if ((GetControlText ('EPPRESTO')<>'') and
   (GetControlText ('EPPRESTO')<>'0')) then
   begin
   CreeDetail (CleDADS, 831, 'S41.G30.15.001', '01', '01');

   CreeDetail (CleDADS, 832, 'S41.G30.15.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EPPRESTO')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EPPRESTO')), 0)));
   if (StrToFloat(GetControlText ('EPPRESTO'))<0) then
                CreeDetail (CleDADS, 833, 'S41.G30.15.002.002', 'N', 'N');
   end;

if ((GetControlText ('EPPTRANSPORT')<>'') and
   (GetControlText ('EPPTRANSPORT')<>'0')) then
   begin
   CreeDetail (CleDADS, 834, 'S41.G30.15.001', '02', '02');

   CreeDetail (CleDADS, 835, 'S41.G30.15.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EPPTRANSPORT')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EPPTRANSPORT')), 0)));
   if (StrToFloat(GetControlText ('EPPTRANSPORT'))<0) then
                CreeDetail (CleDADS, 836, 'S41.G30.15.002.002', 'N', 'N');
   end;

if ((GetControlText ('ECASPARTS')<>'') and
   (GetControlText ('ECASPARTS')<>'0')) then
   begin
   CreeDetail (CleDADS, 841, 'S41.G30.20.001', '01', '01');

   CreeDetail (CleDADS, 842, 'S41.G30.20.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('ECASPARTS')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('ECASPARTS')), 0)));
   if (StrToFloat(GetControlText ('ECASPARTS'))<0) then
                CreeDetail (CleDADS, 843, 'S41.G30.20.002.002', 'N', 'N');
   end;

if ((GetControlText ('ECASPARTA')<>'') and
   (GetControlText ('ECASPARTA')<>'0')) then
   begin
   CreeDetail (CleDADS, 844, 'S41.G30.20.001', '02', '02');

   CreeDetail (CleDADS, 845, 'S41.G30.20.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('ECASPARTA')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('ECASPARTA')), 0)));
   if (StrToFloat(GetControlText ('ECASPARTA'))<0) then
                CreeDetail (CleDADS, 846, 'S41.G30.20.002.002', 'N', 'N');
   end;

if ((GetControlText ('ECASPARTV')<>'') and
   (GetControlText ('ECASPARTV')<>'0')) then
   begin
   CreeDetail (CleDADS, 847, 'S41.G30.20.001', '03', '03');

   CreeDetail (CleDADS, 848, 'S41.G30.20.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('ECASPARTV')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('ECASPARTV')), 0)));
   if (StrToFloat(GetControlText ('ECASPARTV'))<0) then
                CreeDetail (CleDADS, 849, 'S41.G30.20.002.002', 'N', 'N');
   end;

if ((GetControlText ('ECASPARTG')<>'') and
   (GetControlText ('ECASPARTG')<>'0')) then
   begin
   CreeDetail (CleDADS, 850, 'S41.G30.20.001', '04', '04');

   CreeDetail (CleDADS, 851, 'S41.G30.20.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('ECASPARTG')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('ECASPARTG')), 0)));
   if (StrToFloat(GetControlText ('ECASPARTG'))<0) then
                CreeDetail (CleDADS, 852, 'S41.G30.20.002.002', 'N', 'N');
   end;

if ((GetControlText ('EILICENC')<>'') and
   (GetControlText ('EILICENC')<>'0')) then
   begin
   CreeDetail (CleDADS, 861, 'S41.G30.25.001', '01', '01');

   CreeDetail (CleDADS, 862, 'S41.G30.25.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EILICENC')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EILICENC')), 0)));
   if (StrToFloat(GetControlText ('EILICENC'))<0) then
                CreeDetail (CleDADS, 863, 'S41.G30.25.002.002', 'N', 'N');
   end;

if ((GetControlText ('EIRETRAITE')<>'') and
   (GetControlText ('EIRETRAITE')<>'0')) then
   begin
   CreeDetail (CleDADS, 864, 'S41.G30.25.001', '02', '02');

   CreeDetail (CleDADS, 865, 'S41.G30.25.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EIRETRAITE')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EIRETRAITE')), 0)));
   if (StrToFloat(GetControlText ('EIRETRAITE'))<0) then
                CreeDetail (CleDADS, 866, 'S41.G30.25.002.002', 'N', 'N');
   end;

if ((GetControlText ('EGSS')<>'') and (GetControlText ('EGSS')<>'0') and
   (GetControlText ('EGCSG')<>'') and (GetControlText ('EGCSG')<>'0')) then
   begin
   CreeDetail (CleDADS, 871, 'S41.G30.30.001.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EGSS')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EGSS')), 0)));
   if (StrToFloat(GetControlText ('EGSS'))<0) then
                CreeDetail (CleDADS, 872, 'S41.G30.30.001.002', 'N', 'N');

   CreeDetail (CleDADS, 873, 'S41.G30.30.002.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EGCSG')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EGCSG')), 0)));
   if (StrToFloat(GetControlText ('EGCSG'))<0) then
                CreeDetail (CleDADS, 874, 'S41.G30.30.002.002', 'N', 'N');
   end;

{PT107-2
if ((GetControlText ('EEMONTANTREM')<>'') and
   (GetControlText ('EEMONTANTREM')<>'0')) then
}
if ((GetControlText ('EEMONTANTREM')<>'') and
   (GetControlText ('EEMONTANTREM')<>'0') and
   (GetControlText ('CETYPEREM')<>'')) then
//FIN PT107-2
   begin
   CreeDetail (CleDADS, 881, 'S41.G30.35.001.001', GetControlText ('CETYPEREM'),
               GetControlText ('CETYPEREM'));

   CreeDetail (CleDADS, 882, 'S41.G30.35.001.002',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEMONTANTREM')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEMONTANTREM')), 0)));
   if (StrToFloat(GetControlText ('EEMONTANTREM'))<0) then
                CreeDetail (CleDADS, 883, 'S41.G30.35.001.003', 'N', 'N');

{PT103
   CreeDetail (CleDADS, 884, 'S41.G30.35.002',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EENBJOURS')), 0))),
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EENBJOURS')), 0))));
}
   CreeDetail (CleDADS, 884, 'S41.G30.35.002',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EENBJOURS'))*100, 0))),
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EENBJOURS')), 2))));
//FIN PT103

   CreeDetail (CleDADS, 885, 'S41.G30.35.003.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEMONTANTPAT')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEMONTANTPAT')), 0)));
   if (StrToFloat(GetControlText ('EEMONTANTPAT'))<0) then
                CreeDetail (CleDADS, 886, 'S41.G30.35.003.002', 'N', 'N');

   CreeDetail (CleDADS, 887, 'S41.G30.35.004.001',
               FloatToStr (Abs (Arrondi (StrToFloat (GetControlText ('EEMONTANTSAL')), 0))),
               FloatToStr (Arrondi (StrToFloat (GetControlText ('EEMONTANTSAL')), 0)));
   if (StrToFloat(GetControlText ('EEMONTANTSAL'))<0) then
                CreeDetail (CleDADS, 888, 'S41.G30.35.004.002', 'N', 'N');
   end;

//FIN PT93-5

//PT73-8
if ((RegimeComp.CellValues[0, 1] = 'I001') or
   (RegimeComp.CellValues[0, 1] = 'I002') or
   (RegimeComp.CellValues[0, 2] = 'I001') or
   (RegimeComp.CellValues[0, 2] = 'I002') or
   (RegimeComp.CellValues[0, 3] = 'I001') or
   (RegimeComp.CellValues[0, 3] = 'I002') or
   (RegimeComp.CellValues[0, 4] = 'I001') or
   (RegimeComp.CellValues[0, 4] = 'I002') or
   (RegimeComp.CellValues[0, 5] = 'I001') or
   (RegimeComp.CellValues[0, 5] = 'I002')) then
   begin
   CreeDetail (CleDADS, 891, 'S42.G01.00.001',
               GetControlText ('CIRCANTECPERIODICITE'),
               GetControlText ('CIRCANTECPERIODICITE'));

{PT93-5
   CreeDetail (CleDADS, 892, 'S42.G01.00.002',
               GetControlText ('EIRCANTECNBPAIE'),
               GetControlText ('EIRCANTECNBPAIE'));
}

{PT77-2
   CreeDetail (CleDADS, 893, 'S42.G01.00.004',
               GetControlText ('EIRCANTECHEBDO'),
               GetControlText ('EIRCANTECHEBDO'));
}
{PT81
   if (GetControlText('EIRCANTECHEBDO') <> '') then
}
{PT93-5
   if ((GetControlText('CSTATUTPROF') = '63') or
      (GetControlText('CSTATUTPROF') = '65') or
      (GetControlText('CSTATUTPROF') = '66') or
      (GetControlText('CSTATUTPROF') = '68')) then
}
   if ((GetControlText('CSTATUTPROF') = '63') or
      (GetControlText('CSTATUTPROF') = '65')) then
//FIN PT93-5
//FIN PT81
      CreeDetail (CleDADS, 893, 'S42.G01.00.004',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EIRCANTECHEBDO'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EIRCANTECHEBDO')), 2)));
//FIN PT77-2

   if ((GetControlText('EIRCANTECBRUT') <> '')) then
      begin
      CreeDetail (CleDADS, 894, 'S42.G01.00.007.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText ('EIRCANTECBRUT')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText ('EIRCANTECBRUT')), 0)));

      if (StrToFloat(GetControlText('EIRCANTECBRUT')) < 0) then
         CreeDetail (CleDADS, 895, 'S42.G01.00.007.002', 'N', 'N');
      end
   else
      CreeDetail (CleDADS, 894, 'S42.G01.00.007.001', '0', '0');

   if ((GetControlText ('EIRCANTECPLAFOND')<>'')) then
      begin
      CreeDetail (CleDADS, 896, 'S42.G01.00.008.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText ('EIRCANTECPLAFOND')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText ('EIRCANTECPLAFOND')), 0)));

      if (StrToFloat(GetControlText('EIRCANTECPLAFOND')) < 0) then
         CreeDetail (CleDADS, 897, 'S42.G01.00.008.002', 'N', 'N');
      end
   else
      CreeDetail (CleDADS, 896, 'S42.G01.00.008.001', '0', '0');

{PT81
   if ((GetControlText('EIRCANTECMEDECINS') <> '0') and
      (GetControlText('EIRCANTECMEDECINS') <> '')) then
}
   if ((GetControlText('CSTATUTPROF') = '52') or
      (GetControlText('CSTATUTPROF') = '53') or
      (GetControlText('CSTATUTPROF') = '54') or
      (GetControlText('CSTATUTPROF') = '56') or
      (GetControlText('CSTATUTPROF') = '57') or
      (GetControlText('CSTATUTPROF') = '58') or
      (GetControlText('CSTATUTPROF') = '59')) then
//FIN PT81
      begin
      CreeDetail (CleDADS, 898, 'S42.G01.00.009.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText ('EIRCANTECMEDECINS')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText ('EIRCANTECMEDECINS')), 0)));

      if (StrToFloat(GetControlText('EIRCANTECMEDECINS')) < 0) then
         CreeDetail (CleDADS, 899, 'S42.G01.00.009.002', 'N', 'N');
      end;

   if (Nbre = NbreTot) then
      begin
      if ((GetControlText('EIRCANTECFNCNUM') <> '')) then
         begin
         CreeDetail (CleDADS, 901, 'S42.G02.00.001',
                     GetControlText('EIRCANTECFNCNUM'),
                     GetControlText('EIRCANTECFNCNUM'));

         if ((GetControlText('EIRCANTECFNCMONTANT') <> '')) then
            begin
            CreeDetail (CleDADS, 902, 'S42.G02.00.002.001',
                        FloatToStr(Abs(Arrondi(StrToFloat(GetControlText ('EIRCANTECFNCMONTANT')), 0))),
                        FloatToStr(Arrondi(StrToFloat(GetControlText ('EIRCANTECFNCMONTANT')), 0)));

            if (StrToFloat(GetControlText('EIRCANTECFNCMONTANT')) < 0) then
               CreeDetail (CleDADS, 903, 'S42.G02.00.002.002', 'N', 'N');
            end
         else
            CreeDetail (CleDADS, 902, 'S42.G02.00.002.001', '0', '0');

         end;
      end;
   end;
//FIN PT73-8

//PT90
If (Ircantec = False) then
   begin
   if (GetControlText('CDUREETRAV') <> '') then
      CreeDetail (CleDADS, 911, 'S44.G01.00.001', GetControlText('CDUREETRAV'),
                  GetControlText('CDUREETRAV'));

   if (GetControlText('EDUREETRAV') <> '') then
      CreeDetail (CleDADS, 912, 'S44.G01.00.002',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EDUREETRAV'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EDUREETRAV')), 2)));

//PT93-5
   if ((AGFF<>NIL) and (AGFF.Checked=TRUE)) then
      CreeDetail (CleDADS, 913, 'S44.G01.00.003', '01', '01');
//FIN PT93-5
   end;
//FIN PT90

if ((GetControlText('EENTREE') <> '') and
   (GetControlText('EENTREE') <> '  /  /    ')) then
   begin
   ForceNumerique(GetControlText('EENTREE'), Buffer);
   CreeDetail (CleDADS, 921, 'S45.G01.00.001', Buffer,
               GetControlText('EENTREE'));

   if (Prevoyance <> NIL) then
      begin
      j := 921;
      for i := 1 to 4 do
          begin
          if (Prevoyance.CellValues[0,i] <> '') then
             begin
{PT102
             CreeDetail (CleDADS, j+(i*12)-11, 'S45.G01.01.001',
                         Prevoyance.CellValues[0,i],
                         Prevoyance.CellValues[0,i]);
}
             StEvt:= 'SELECT CO_CODE'+
                     ' FROM COMMUN WHERE'+
                     ' CO_TYPE="PDR" AND'+
                     ' CO_LIBELLE="'+Prevoyance.CellValues[0,i]+'"';
             QRechEvt:= OpenSQL(StEvt,TRUE) ;
             if Not QRechEvt.EOF then
                Buffer:= QRechEvt.FindField ('CO_CODE').AsString;
             Ferme (QRechEvt);

             CreeDetail (CleDADS, j+(i*12)-11, 'S45.G01.01.001', Buffer, Buffer);
//FIN PT102             
//PT96             if Prevoyance.CellValues[0,i] <> '90' then
//PT96                begin
             ForceNumerique(Prevoyance.CellValues[1,i], Buffer);
             CreeDetail (CleDADS, j+(i*12)-10, 'S45.G01.01.002',
//PT96                           Copy(Buffer,1,4), Prevoyance.CellValues[1,i]);
                         Buffer, Prevoyance.CellValues[1,i]);


             ForceNumerique(Prevoyance.CellValues[2,i], Buffer);
             CreeDetail (CleDADS, j+(i*12)-9, 'S45.G01.01.003',
//PT96                            Copy(Buffer,1,4), Prevoyance.CellValues[2,i]);
                         Buffer, Prevoyance.CellValues[2,i]);
//PT96                end;

{PT93-5
             CreeDetail (CleDADS, j+(i*11)-7, 'S45.G01.01.004',
                         'P'+Prevoyance.CellValues[3,i],
                         Prevoyance.CellValues[3,i]);
}
             CreeDetail (CleDADS, j+(i*12)-8, 'S45.G01.01.004.001',
                         'P'+Prevoyance.CellValues[3,i],
                         Prevoyance.CellValues[3,i]);
             if Prevoyance.CellValues[4,i] <> '' then  //PT94
               CreeDetail (CleDADS, j+(i*12)-7, 'S45.G01.01.004.002',
                           Prevoyance.CellValues[4,i],
                           Prevoyance.CellValues[4,i]);
//FIN PT93-5
             CreeDetail (CleDADS, j+(i*12)-6, 'S45.G01.01.005',
                         Prevoyance.CellValues[5,i],
                         Prevoyance.CellValues[5,i]);

             if ((Prevoyance.CellValues[6,i] <> '') and
                (Prevoyance.CellValues[7,i] <> '') and
{PT73-5
                (Prevoyance.CellValues[6,i] <> '0') and
}                
                (isnumeric(Prevoyance.CellValues[7,i]))) then
                begin
                CreeDetail (CleDADS, j+(i*12)-5, 'S45.G01.01.006',
                            Prevoyance.CellValues[6,i],
                            Prevoyance.CellValues[6,i]);
                CreeDetail (CleDADS, j+(i*12)-4, 'S45.G01.01.007.001',
                            FloatToStr(Abs(Arrondi(StrToFloat(Prevoyance.CellValues[7,i]), 0))),
                            FloatToStr(Arrondi(StrToFloat(Prevoyance.CellValues[7,i]), 0)));
                if (StrToFloat(Prevoyance.CellValues[7,i]) < 0) then
                   CreeDetail (CleDADS, j+(i*12)-3, 'S45.G01.01.007.002', 'N',
                               'N');
                end
             else
                begin
                Prevoyance.CellValues[6,i]:= '';
                Prevoyance.CellValues[7,i]:= '';
                end;
             CreeDetail (CleDADS, j+(i*12)-2, 'S45.G01.01.008',
                         Prevoyance.CellValues[8,i],
                         Prevoyance.CellValues[8,i]);
             CreeDetail (CleDADS, j+(i*12)-1, 'S45.G01.01.009',
                         RechDom('PGSITUATIONFAMIL', Prevoyance.CellValues[9,i], TRUE),
                         Prevoyance.CellValues[9,i]);
             if (isnumeric(Prevoyance.CellValues[10,i])) then
                CreeDetail (CleDADS, j+(i*12), 'S45.G01.01.010',
                            Prevoyance.CellValues[10,i],
                            Prevoyance.CellValues[10,i])
             else
                CreeDetail (CleDADS, j+(i*12), 'S45.G01.01.010', '90',
                            'Inconnu');
             end;
          end;
      end;
   end;

{PT73-7
if (TypeD <> '001') then
}
if ((TypeD <> '001') and (TypeD<>'201')) then
//FIN PT73-7
   begin
   if (GetControlText('EBTPADHES') <> '') then
      CreeDetail (CleDADS, 971, 'S66.G01.00.001', GetControlText('EBTPADHES'),
                  GetControlText('EBTPADHES'));

   if (GetControlText ('CBTPUNITTEMPS')<>'') then
{PT93-5
      CreeDetail (CleDADS, 972, 'S66.G01.00.002',
                  GetControlText('CBTPUNITTEMPS'),
                  GetControlText('CBTPUNITTEMPS'));
}
      CreeDetail (CleDADS, 972, 'S66.G01.00.002.001',
                  GetControlText('CBTPUNITTEMPS'),
                  GetControlText('CBTPUNITTEMPS'));

   if ((Ouvre<>NIL) and (Ouvre.Checked=TRUE)) then
      CreeDetail (CleDADS, 973, 'S66.G01.00.002.002', '02', '02')
   else
      CreeDetail (CleDADS, 973, 'S66.G01.00.002.002', '01', '01');
//FIN PT93-5

   if (GetControlText('EBTPTEMPS') <> '') then
      CreeDetail (CleDADS, 974, 'S66.G01.00.003',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPTEMPS'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPTEMPS')), 2)));

   if (GetControlText('EBTPANCIENENT') <> '') then
      CreeDetail (CleDADS, 975, 'S66.G01.00.004',
                  GetControlText('EBTPANCIENENT'),
                  GetControlText('EBTPANCIENENT'));

   if (GetControlText('EBTPANCIENPROF') <> '') then
      CreeDetail (CleDADS, 976, 'S66.G01.00.005',
                  GetControlText('EBTPANCIENPROF'),
                  GetControlText('EBTPANCIENPROF'))
   else
      CreeDetail (CleDADS, 976, 'S66.G01.00.005', '99', '99');

   if (GetControlText('CBTPASSEDIC') <> '') then
      begin
      if GetControlText('CBTPASSEDIC') = 'OUI' then
         CreeDetail (CleDADS, 977, 'S66.G01.00.006', '01',
                     GetControlText('CBTPASSEDIC'))
      else
         CreeDetail (CleDADS, 977, 'S66.G01.00.006', '02',
                     GetControlText('CBTPASSEDIC'));
      end
//PT73-5
   else
      CreeDetail (CleDADS, 977, 'S66.G01.00.006', '01', '01');
//FIN PT73-5

   if (GetControlText('CBTPUNITSALMOY') <> '') then
      CreeDetail (CleDADS, 978, 'S66.G01.00.007',
                  GetControlText('CBTPUNITSALMOY'),
                  GetControlText('CBTPUNITSALMOY'))
//PT73-5
   else
      CreeDetail (CleDADS, 978, 'S66.G01.00.007', '01', '01');
//FIN PT73-5

   if ((GetControlText('EBTPSALMOY') <> '') and
      (GetControlText('CBTPUNITSALMOY') <> '03')) then
      CreeDetail (CleDADS, 979, 'S66.G01.00.008',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPSALMOY'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPSALMOY')), 2)));

   Buffer := '';
   if (GetControlText('CBTPUNITHOR') <> '') then
      begin
      if (GetControlText('CBTPUNITHOR') = 'HEB') then
         Buffer := '01'
      else
         Buffer := '02';
      end
   else
      begin
      SetControlText('CBTPUNITHOR', 'MEN');
      Buffer := '02';
      end;
   CreeDetail (CleDADS, 980, 'S66.G01.00.009', Buffer,
               GetControlText('CBTPUNITHOR'));

   if (GetControlText('EBTPHORSAL') <> '') then
      CreeDetail (CleDADS, 981, 'S66.G01.00.010',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPHORSAL'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPHORSAL')), 2)));

   if (GetControlText('EBTPHOREMP') <> '') then
      CreeDetail (CleDADS, 982, 'S66.G01.00.011',
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPHOREMP'))*100, 0)),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPHOREMP')), 2)));

   if (GetControlText('CBTPSITFAM') <> '') then
      begin
      if (GetControlText('CBTPSITFAM') = '90') then
         CreeDetail (CleDADS, 983, 'S66.G01.00.012', '99', '90')
      else
         CreeDetail (CleDADS, 983, 'S66.G01.00.012',
                     RechDom('PGSITUATIONFAMIL', GetControlText('CBTPSITFAM'), TRUE),
                     GetControlText('CBTPSITFAM'));
      end;

   if (GetControlText('EBTPENFANT') <> '') then
      CreeDetail (CleDADS, 984, 'S66.G01.00.013', GetControlText('EBTPENFANT'),
                  GetControlText('EBTPENFANT'));

   if ((GetControlText('EBTPFRAISPROF') <> '') and
      (GetControlText('EBTPFRAISPROF') <> '0')) then
      CreeDetail (CleDADS, 985, 'S66.G01.00.014',
                  FORMAT_STRING(GetControlText('EBTPFRAISPROF'), 2),
                  GetControlText('EBTPFRAISPROF'))
   else
      CreeDetail (CleDADS, 985, 'S66.G01.00.014', '99',
                  GetControlText('EBTPFRAISPROF'));


   if (GetControlText('CBTPSTATUTCOTIS') <> '') then
      CreeDetail (CleDADS, 986, 'S66.G01.00.015',
                  GetControlText('CBTPSTATUTCOTIS'),
                  GetControlText('CBTPSTATUTCOTIS'))
   else
      CreeDetail (CleDADS, 986, 'S66.G01.00.015', '90',
                  GetControlText('CBTPSTATUTCOTIS'));

{PT73-5
   if (GetControlText('EBTPNIVEAU') <> '') then
      CreeDetail (CleDADS, 916, 'S66.G01.00.016', GetControlText('EBTPNIVEAU'),
                  GetControlText('EBTPNIVEAU'));

   if (GetControlText('EBTPCOEFF') <> '') then
      CreeDetail (CleDADS, 917, 'S66.G01.00.017', GetControlText('EBTPCOEFF'),
                  GetControlText('EBTPCOEFF'));

   if (GetControlText('EBTPPOSITION') <> '') then
      CreeDetail (CleDADS, 918, 'S66.G01.00.018',
                  GetControlText('EBTPPOSITION'),
                  GetControlText('EBTPPOSITION'));

   if (GetControlText('EBTPECHELON') <> '') then
      CreeDetail (CleDADS, 919, 'S66.G01.00.019', GetControlText('EBTPECHELON'),
                  GetControlText('EBTPECHELON'));

   if (GetControlText('EBTPCATEGORIE') <> '') then
      CreeDetail (CleDADS, 920, 'S66.G01.00.020',
                  GetControlText('EBTPCATEGORIE'),
                  GetControlText('EBTPCATEGORIE'));

   CreeDetail (CleDADS, 921, 'S66.G01.00.021', GetControlText('EMETIERBTP'),
               GetControlText('EMETIERBTP'));
}
   if (GetControlText('EMETIERBTP') <> '') then
      CreeDetail (CleDADS, 991, 'S66.G01.00.021', GetControlText('EMETIERBTP'),
                  GetControlText('EMETIERBTP'))
   else
      CreeDetail (CleDADS, 991, 'S66.G01.00.021', 'A0000', 'A0000');
//FIN PT73-5

{PT73-5
   CreeDetail (CleDADS, 922, 'S66.G01.00.022', GetControlText('CBTPAFFILIRC'),
               GetControlText('CBTPAFFILIRC'));
}
   if (GetControlText('CBTPAFFILIRC') <> '') then
      CreeDetail (CleDADS, 992, 'S66.G01.00.022', GetControlText('CBTPAFFILIRC'),
                  GetControlText('CBTPAFFILIRC'))
   else
      CreeDetail (CleDADS, 992, 'S66.G01.00.022', '01', '01');
//FIN PT73-5
   if (GetControlText('EBTPBASECPBTP') <> '') then
      begin
      CreeDetail (CleDADS, 993, 'S66.G01.00.023.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBTPBASECPBTP')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPBASECPBTP')), 0)));
      if (StrToFloat(GetControlText('EBTPBASECPBTP')) < 0) then
         CreeDetail (CleDADS, 994, 'S66.G01.00.023.002', 'N', 'N');
      end;

   if (GetControlText('EBTPBASECOTIS') <> '') then
      begin
      CreeDetail (CleDADS, 995, 'S66.G01.00.024.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBTPBASECOTIS')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPBASECOTIS')), 0)));
      if (StrToFloat(GetControlText('EBTPBASECOTIS')) < 0) then
         CreeDetail (CleDADS, 996, 'S66.G01.00.024.002', 'N', 'N');
      end;

//PT73-5
   if (GetControlText('EBTPBASECOTISOPPBTP') <> '') then
      begin
      CreeDetail (CleDADS, 997, 'S66.G01.00.025.001',
                  FloatToStr(Abs(Arrondi(StrToFloat(GetControlText('EBTPBASECOTISOPPBTP')), 0))),
                  FloatToStr(Arrondi(StrToFloat(GetControlText('EBTPBASECOTISOPPBTP')), 0)));
      if (StrToFloat(GetControlText('EBTPBASECOTISOPPBTP')) < 0) then
         CreeDetail (CleDADS, 998, 'S66.G01.00.025.002', 'N', 'N');
      end;

   CreeDetail (CleDADS, 999, 'S66.G01.00.026', GetControlText('ECLASSIFBTP'),
               GetControlText('ECLASSIFBTP'));
//FIN PT73-5
   end;

{PT82
DuplicPrudh(Sal, TypeD, 'C', Ordre);
}
DuplicPrudh(Sal, 'C', Ordre);
StSal := 'UPDATE SALARIES SET PSA_DADSDATE = "'+UsDateTime(Date)+'" WHERE'+
         ' PSA_SALARIE="'+Sal+'"';
ExecuteSQL(StSal) ;

{$IFNDEF DADSUSEULE}
CreeJnalEvt('001', '042', 'OK', NIL, NIL, Trace);
{$ENDIF}
if Trace <> nil then
   FreeAndNil (Trace);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/09/2001
Modifié le ... :   /  /
Description .. : Effacement des zones de la fiche de saisie complémentaire
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MetABlanc();
var
i, j : integer;
begin
//THEdit
SetControlText ('DATEDEBDADS', DateToStr(IDate1900));
SetControlText ('DATEFINDADS', DateToStr(IDate1900));
//PT73-3
SetControlText ('EDEBMOTIF1', '');
SetControlText ('EDEBMOTIF2', '');
SetControlText ('EDEBMOTIF3', '');
SetControlText ('EFINMOTIF1', '');
SetControlText ('EFINMOTIF2', '');
SetControlText ('EFINMOTIF3', '');
//FIN PT73-3
SetControlText ('ESECTIONAT', '');
SetControlText ('ETAUXPARTIEL', '');
SetControlText ('EHEURESTRAV', '');
SetControlText ('EHEURESSAL', '');
SetControlText ('EDERNIERMOIS', '');
SetControlText ('EHEURESCHOMPAR', '');
SetControlText ('EBRUTSS', '');
SetControlText ('EPLAFONDSS', '');
SetControlText ('EBASECSG', '');
SetControlText ('EBASECRDS', '');
SetControlText ('EBRUTFISC', '');
SetControlText ('EAVANTAGENAT', '');
SetControlText ('EFRAISPROF', '');
SetControlText ('ERISQUEAT', '');
SetControlText ('ETAUXAT', '');
SetControlText ('ERETSALAIRE', '');
SetControlText ('ECHQVACANCE', '');
SetControlText ('EIMPOTRETSOURC', '');
SetControlText ('EINDEXPATRIAT', '');
SetControlText ('ETOTIMPOTAXE', '');
SetControlText ('EBASEIMPO1', '');
SetControlText ('EBASEIMPO2', '');
SetControlText ('ETAXESAL', '');
SetControlText ('EREVENUSACTIV', '');
SetControlText ('ECP', '');
SetControlText ('EAUTREREV', '');
SetControlText ('EEPARGNE', '');
SetControlText ('EINDIMPATRIAT', '');
//PT73-5
SetControlText ('ESERVICEPERS', '');
SetControlText ('ECPPLAF', '');
//FIN PT73-5
//PT93-5
SetControlText ('EHEURESSUPP', '');
SetControlText ('EEPARGNESALP', '');
SetControlText ('EEPARGNESALI', '');
SetControlText ('EEPARGNESALA', '');
SetControlText ('EEPARGNESALD', '');
SetControlText ('EACTIONSN1', '');
SetControlText ('EACTIONSV1', '');
SetControlText ('EACTIONSN2', '');
SetControlText ('EACTIONSV2', '');
SetControlText ('EACTIONSN3', '');
SetControlText ('EACTIONSV3', '');
SetControlText ('EPPRESTO', '');
SetControlText ('EPPTRANSPORT', '');
SetControlText ('ECASPARTS', '');
SetControlText ('ECASPARTA', '');
SetControlText ('ECASPARTV', '');
SetControlText ('ECASPARTG', '');
SetControlText ('EILICENC', '');
SetControlText ('EIRETRAITE', '');
SetControlText ('EGSS', '');
SetControlText ('EGCSG', '');
SetControlText ('EEMONTANTREM', '');
SetControlText ('EENBJOURS', '');
SetControlText ('EEMONTANTPAT', '');
SetControlText ('EEMONTANTSAL', '');
//FIN PT93-5
SetControlText ('ELIBELLEEMPLOI', '');
SetControlText ('EPCS', '');
SetControlText ('ECOEFF', '');
SetControlText ('ECONVCOLL', '');
SetControlText ('ECSGSPECM', '');
//PT73-8
{PT93-5
SetControlText ('EIRCANTECNBPAIE', '');
}
SetControlText ('EIRCANTECHEBDO', '');
SetControlText ('EIRCANTECBRUT', '');
SetControlText ('EIRCANTECPLAFOND', '');
SetControlText ('EIRCANTECMEDECINS', '');
SetControlText ('EIRCANTECFNCNUM', '');
SetControlText ('EIRCANTECFNCMONTANT', '');
//FIN PT73-8
SetControlText ('EDUREETRAV', '');
SetControlText ('EENTREE', '');
SetControlText ('EBTPADHES', '');
SetControlText ('EBTPTEMPS', '');
SetControlText ('EBTPSALMOY', '');
SetControlText ('EBTPHORSAL', '');
SetControlText ('EBTPHOREMP', '');
SetControlText ('EBTPANCIENENT', '');
SetControlText ('EBTPANCIENPROF', '');
SetControlText ('EBTPENFANT', '');
SetControlText ('EBTPFRAISPROF', '');
{PT73-5
SetControlText ('EBTPNIVEAU', '');
SetControlText ('EBTPCOEFF', '');
SetControlText ('EBTPPOSITION', '');
SetControlText ('EBTPECHELON', '');
SetControlText ('EBTPCATEGORIE', '');
}
SetControlText ('EMETIERBTP', '');


//Combos
SetControlText ('CETABLISSEMENT', '');
SetControlText ('DECALAGE', '');  //PT73-2
{PT73-3
SetControlText ('CDEBMOTIF1', '');
SetControlText ('CDEBMOTIF2', '');
SetControlText ('CDEBMOTIF3', '');
SetControlText ('CFINMOTIF1', '');
SetControlText ('CFINMOTIF2', '');
SetControlText ('CFINMOTIF3', '');
}
SetControlText ('CCONDEMPLOI', '');
SetControlText ('CSTATUTPROF', '');
SetControlText ('CSTATUTCAT', '');
SetControlText ('CREGIMESS', '');
//DEB PT92-1
{PT93-5
SetControlText ('CHTYPEREGIME', '-');
}
SetControlText ('CREGIMEAT', '');
SetControlText ('CREGIMEMAL', '');
SetControlText ('CREGIMEVIP', '');
SetControlText ('CREGIMEVIS', '');
//FIN PT92-1
SetControlText ('CCONTRAT', '');
SetControlText ('CTRAVETRANG', '');
//PT93-5
SetControlText ('CCNBF', '');
SetControlText ('CETYPEREM', '');
//FIN PT93-5
SetControlText ('CPRUDHCOLL', '');
SetControlText ('CPRUDHSECT', '');
SetControlText ('CIRCANTECPERIODICITE', '');  //PT73-8
SetControlText ('CDUREETRAV', '');
SetControlText ('CBTPUNITTEMPS', '');
SetControlText ('CBTPUNITSALMOY', '');
SetControlText ('CBTPUNITHOR', '');
SetControlText ('CBTPSITFAM', '');
SetControlText ('CBTPASSEDIC', '');
SetControlText ('CBTPSTATUTCOTIS', '');
SetControlText ('CBTPAFFILIRC', '');

//CheckBox
//PT93-5
if (TypeRegime<>nil) then
   TypeRegime.Checked:= False;
if (Alsace<>NIL) then
   Alsace.Checked:= False;
if (BaseReelle<>NIL) then
   BaseReelle.Checked:= True;
//FIN PT93-5
if Nourriture <> NIL then
   Nourriture.Checked:= False;
if Logement <> NIL then
   Logement.Checked:= False;
if Voiture <> NIL then
   Voiture.Checked:= False;
if AutreAvant <> NIL then
   AutreAvant.Checked:= False;
if NTIC <> NIL then
   NTIC.Checked:= False;
if AllocForfait <> NIL then
   AllocForfait.Checked:= False;
if RembJustif <> NIL then
   RembJustif.Checked:= False;
if PriseCharge <> NIL then
   PriseCharge.Checked:= False;
if AutreDepense <> NIL then
   AutreDepense.Checked:= False;
if CodeBureau <> NIL then
   CodeBureau.Checked:= False;
if MultiEmp <> NIL then
   MultiEmp.Checked:= False;
if PrudDP <> NIL then
   PrudDP.Checked:= False;
if PrudPresent <> NIL then
   PrudPresent.Checked:= False;
if RemPourboire <> NIL then
   RemPourboire.Checked:= False;
//PT93-5
if (AGFF<>NIL) then
   AGFF.Checked:= False;
//FIN PT93-5

//Labels
SetControlText ('LMETIERBTP_', '...');
SetControlText ('LLIBELLEEMPLOI_', '...');
SetControlText ('LPCS_', '...');
SetControlText ('LCOEFF_', '...');
SetControlText ('LSALARIE_', Sal);
SetControlText ('LCONVCOLL_', '...');
//PT79-1
SetControlText ('LDEBMOTIF1', '...');
SetControlText ('LDEBMOTIF2', '...');
SetControlText ('LDEBMOTIF3', '...');
SetControlText ('LFINMOTIF1', '...');
SetControlText ('LFINMOTIF2', '...');
SetControlText ('LFINMOTIF3', '...');
//FIN PT79-1

// Mise à blanc des THGrid
for i := 1 to 5 do
    for j := 0 to 2 do
        Regimecomp.CellValues[j,i] := '';
for i := 1 to 10 do
    for j := 0 to 3 do
        Exoner.CellValues[j,i] := '';
for i := 1 to 4 do
    for j := 0 to 4 do
        Somiso.CellValues[j,i] := '';
for i := 1 to 4 do
    for j := 0 to 2 do
        Basebrutespec.CellValues[j,i] := '';
for i := 1 to 4 do
    for j := 0 to 2 do
        Baseplafspec.CellValues[j,i] := '';
for i := 1 to 4 do
{PT93-5
    for j := 0 to 9 do
}
    for j := 0 to 10 do
//FIN PT93-5    
        Prevoyance.CellValues[j,i] := '';
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/01/2002
Modifié le ... :   /  /
Description .. : Effacement des zones de la fiche de saisie complémentaire
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MetABlancSal();
begin
SetControlText('LNUMSS', '...');
SetControlText('LCIVILITE', '...');
SetControlText('LNOMJF', '...');
SetControlText('LSURNOM', '...');
SetControlText('LADRCOMPL', '...');
SetControlText('LADRNUM', '...');
SetControlText('LADRNOM', '...');
SetControlText('LCODEPOSTAL', '...');
SetControlText('LVILLE', '...');
SetControlText('LCODEPAYS', '...');
SetControlText('LNOMPAYS', '...');
SetControlText('LDATENAISSANCE', '...');
SetControlText('LCOMMUNENAISS', '...');
SetControlText('LPAYSNAISSANCE', '...');
SetControlText('LDEPTNAISSANCE', '...');
SetControlText('LNATIONALITE', '...');

{NE PAS METTRE A BLANC (CHARGE DANS LE ONARGUMENT)
if Lab_Fraction <> NIL then
   Lab_Fraction.Caption := '...';
}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Contrôle des données
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_PG_DADSPERIODES.ControleConform() : boolean;
var
DateDebBuf, DateFinBuf : TDateTime;
Buffer, CInterdit, Mess, StEtab : string;
BoolFraisProf , ctlbasebrute, ctlbaseplaf : boolean; // pt93-6
i, j, Ordre, rep : integer;
ErreurDADSU : TControle;
DoubleBuf : Double;
QEtab : TQuery;
begin
result := TRUE;
Mess:='';

Ordre:= StrToInt(GetControlText('ORDREDADS'));

//PT73-9
ErreurDADSU.Salarie:= Sal;
ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.DateDeb:= StrToDate(GetControlText('DATEDEBDADS'));
ErreurDADSU.DateFin:= StrToDate(GetControlText('DATEFINDADS'));
ErreurDADSU.Exercice:= PGExercice;
//FIN PT73-9

//PT73-1
If ((GetControlText('LNUMSS') = '1999999999999') or
   (GetControlText('LNUMSS') = '2999999999999')) then
   begin
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.001';
   ErreurDADSU.Explication:= 'Le N° de sécurité Sociale est invalide, il a'+
                             ' été forcé à "'+GetControlText('LNUMSS')+'"';
   ErreurDADSU.CtrlBloquant:= False;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT73-1

if (GetControlText('EPRENOM')= '') then
   begin
   result := FALSE;
   Mess := Mess+ '#13#10 - Fiche Salarié : Prénom absent';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.003';
   ErreurDADSU.Explication:= 'Le prénom n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if (GetControlText('LCIVILITE') = '') then
   begin
   result := FALSE;
   Mess := Mess+ '#13#10 - Fiche salarié : Civilité absente';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.007';
   ErreurDADSU.Explication:= 'La civilité n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if (((GetControlText('LCODEPAYS') = 'FR') or
   (GetControlText('LCODEPAYS') = '...')) and
   ((GetControlText('LCODEPOSTAL') = '') or
   (GetControlText('LCODEPOSTAL') < '00000') or
   (GetControlText('LCODEPOSTAL') > '99999'))) then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Adresse incorrecte';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.008.010';
   ErreurDADSU.Explication:= 'Le code postal est incorrect';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

//PT73-1
if (GetControlText('LVILLE') = '') then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Adresse incorrecte';
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.008.012';
   ErreurDADSU.Explication:= 'La ville n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
   end;
//FIN PT73-1

ForceNumerique(GetControlText('LDATENAISSANCE'), Buffer);
if (Length(Buffer) <> 8) then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Date de naissance erronée';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.009';
   ErreurDADSU.Explication:= 'La date de naissance est erronée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if ((GetControlText('LPAYSNAISSANCE') = 'FRA') and
   ((GetControlText('LCOMMUNENAISS') = '') or
   (GetControlText('LCOMMUNENAISS') = '...'))) then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Commune de naissance absente';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.010';
   ErreurDADSU.Explication:= 'La commune de naissance n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if ((GetControlText('LPAYSNAISSANCE') = 'FRA') and
   (GetControlText('LDEPTNAISSANCE') = '')) then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Département de naissance absent';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.011';
   ErreurDADSU.Explication:= 'Le département de naissance n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if (GetControlText('LPAYSNAISSANCE') = '') then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Pays de naissance absent';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.012';
   ErreurDADSU.Explication:= 'Le pays de naissance n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if (GetControlText('LNATIONALITE') = '') then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Fiche salarié : Nationalité absente';
//PT73-1
   ErreurDADSU.Num:= 0;
   ErreurDADSU.Segment:= 'S30.G01.00.013';
   ErreurDADSU.Explication:= 'La nationalité n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if ((StrToDate(GetControlText('DATEDEBDADS'))<DebExer) or
   (StrToDate(GetControlText('DATEDEBDADS'))>FinExer)) then
   SetControlText('DATEDEBDADS', DateToStr(IDate1900));

if (GetControlText ('DATEDEBDADS')=DateToStr (IDate1900)) then
   begin
   result:= FALSE;
   Mess:= Mess+ '#13#10 - Date de début de période absente';
   SetFocusControl ('DATEDEBDADS');
//PT73-1
   ErreurDADSU.Num:= Ordre;
   ErreurDADSU.Segment:= 'S41.G01.00.001';
   ErreurDADSU.Explication:= 'La date de début de période n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur (ErreurDADSU);
//FIN PT73-1
   end;

{PT73-3
if (GetControlText('CDEBMOTIF1') = '') then
}
if (GetControlText('EDEBMOTIF1') = '') then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Motif de début de période absent';
{PT73-3
   SetFocusControl('CDEBMOTIF1');
}
   SetFocusControl('EDEBMOTIF1');
//PT73-1
   ErreurDADSU.Num:= Ordre;
   ErreurDADSU.Segment:= 'S41.G01.00.002.001';
   ErreurDADSU.Explication:= 'Le motif de début de période n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

if ((StrToDate(GetControlText('DATEDEBDADS'))>StrToDate(GetControlText('DATEFINDADS'))) or
   (StrToDate(GetControlText('DATEFINDADS'))<DebExer) or
   (StrToDate(GetControlText('DATEFINDADS'))>FinExer)) then
   SetControlText('DATEFINDADS', DateToStr(IDate1900));

if (GetControlText('DATEFINDADS') = DateToStr(IDate1900)) then
   begin
   result := FALSE;
   Mess := Mess+ '#13#10 - Date de fin de période absente';
   SetFocusControl('DATEFINDADS');
//PT73-1
   ErreurDADSU.Num:= Ordre;
   ErreurDADSU.Segment:= 'S41.G01.00.003';
   ErreurDADSU.Explication:= 'La date de fin de période n''est pas renseignée';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;

{PT73-3
if ((GetControlText('CDEBMOTIF1') <> '095') and
   (GetControlText('CFINMOTIF1') = '096')) then
   SetControlText('CFINMOTIF1', '');

if (GetControlText('CFINMOTIF1') = '') then
   begin
   result := FALSE;
   Mess := Mess+ '#13#10 - Motif de fin de période';
   SetFocusControl('CFINMOTIF1');
   end;
}
if ((GetControlText('EDEBMOTIF1') <> '095') and
   (GetControlText('EFINMOTIF1') = '096')) then
   SetControlText('EFINMOTIF1', '');

if (GetControlText('EFINMOTIF1') = '') then
   begin
   result:= FALSE;
   Mess:= Mess+'#13#10 - Motif de fin de période absent';
   SetFocusControl('EFINMOTIF1');
//PT73-1
   ErreurDADSU.Num:= Ordre;
   ErreurDADSU.Segment:= 'S41.G01.00.004.001';
   ErreurDADSU.Explication:= 'Le motif de fin de période n''est pas renseigné';
   ErreurDADSU.CtrlBloquant:= True;
   EcrireErreur(ErreurDADSU);
//FIN PT73-1
   end;
//FIN PT73-3

if (State<>'CREATION') then
   begin
   if (GetControlText ('CETABLISSEMENT')='') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Etablissement erroné';
      SetFocusControl('CETABLISSEMENT');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.005';
      ErreurDADSU.Explication:= 'Le SIRET de l''établissement n''est pas valide';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;

//PT73-1
   for i := 1 to 5 do
       begin
       if ((RegimeComp.CellValues[0,i] = 'F001') and
          (MultiEmp.Checked = TRUE)) then
          begin
          if (GetControlText('CCONDEMPLOI')<>'N') then
             begin
             SetControlText ('CCONDEMPLOI', 'N');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.008.001';
{PT91
             ErreurDADSU.Explication:= 'La condition d''emploi a été forcée à'+
                                       ' "Temps non complet (Secteur Public)"';
}
             ErreurDADSU.Explication:= 'La caractéristique activité a été'+
                                       ' forcée à "Temps non complet (Secteur'+
                                       ' Public)"';
             ErreurDADSU.CtrlBloquant:= False;
             EcrireErreur(ErreurDADSU);
             break;
             end;
          end;
       end;
//FIN PT73-1

//PT73-2
   if (GetControlText('DECALAGE') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Décalage de paie absent';
      SetFocusControl('DECALAGE');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.009';
      ErreurDADSU.Explication:= 'Le code de décalage de paie n''est pas'+
                                ' renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT73-2

//PT73-1
   for i := 1 to 5 do
       begin
       if (((RegimeComp.CellValues[0,i] = 'F001') or
          (RegimeComp.CellValues[0,i] = 'F002')) and
          ((GetControlText ('DECALAGE') = '02') or
          (GetControlText ('DECALAGE') = '03') or
          (GetControlText ('DECALAGE') = '04'))) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence code décalage de paie/régime'+
                 ' complémentaire';
          SetFocusControl ('DECALAGE');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.00.009';
          ErreurDADSU.Explication:= 'Incohérence contrôle C2 du cahier des'+
                                    ' charges';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
          break;
          end;
       end;

//PT93-5
   if ((GetControlText ('DECALAGE')='05') and
      (((StrToDate(GetControlText('DATEDEBDADS'))<StrToDate ('0112'+Copy (DateToStr (FinExer), 7, 4))) and
      (StrToDate(GetControlText('DATEDEBDADS'))>StrToDate ('3112'+Copy (DateToStr (FinExer), 7, 4)))) or
      ((StrToDate(GetControlText('DATEFINDADS'))<StrToDate ('0112'+Copy (DateToStr (FinExer), 7, 4))) and
      (StrToDate(GetControlText('DATEFINDADS'))>StrToDate ('3112'+Copy (DateToStr (FinExer), 7, 4)))))) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code décalage de paie/dates de période';
      SetFocusControl('DECALAGE');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.009';
      ErreurDADSU.Explication:= 'Le code de décalage de paie n''est pas'+
                                ' compatible avec les dates de période';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;

   if ((GetControlText ('DECALAGE')='05') and (PrudDP.Checked=True)) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code décalage de paie/droit public';
      SetFocusControl('DECALAGE');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.009';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-03 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT93-5

   for i := 1 to 5 do
       begin
       if (((RegimeComp.CellValues[0,i] = 'F001') or
          (RegimeComp.CellValues[0,i] = 'F002')) and
          (GetControlText ('CCONTRAT') <> 'SC')) then
          begin
          SetControlText ('CCONTRAT', 'SC');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.00.012.001';
          ErreurDADSU.Explication:= 'Le contrat a été forcé à "Sans contrat"';
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur (ErreurDADSU);
          break;
          end;
       end;

   if (GetControlText('CCONTRAT')='') then
      begin
      SetControlText ('CCONTRAT', 'CDI');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.012.001';
      ErreurDADSU.Explication:= 'Le contrat a été forcé à "CDI"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

   for i := 1 to 10 do
       begin
{PT74
       if (((GetControlText ('CCONTRAT') = 'CAA') or
          (GetControlText ('CCONTRAT') = 'CAP')) and
          (Exoner.CellValues[0,i] <> '01') and
          (Exoner.CellValues[0,i] <> '02') and
          (Exoner.CellValues[0,i] <> '03')) then
}
       if ((GetControlText ('CCONTRAT') = 'CAA') or
          (GetControlText ('CCONTRAT') = 'CAP')) then
          begin
{PT93-5
          if ((Exoner.CellValues[0,i] <> '01') and
             (Exoner.CellValues[0,i] <> '02') and
             (Exoner.CellValues[0,i] <> '03')) then
}
          if ((Exoner.CellValues[0,i] <> '01') and
             (Exoner.CellValues[0,i] <> '02') and
             (Exoner.CellValues[0,i] <> '03') and
             (GetControlText ('DECALAGE')<>'05')) then
//FIN PT93-5
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Incohérence code contrat de'+
                    ' travail/exonération';
             SetFocusControl ('CCONTRAT');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.012.001';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2-03 du cahier'+
                                       ' des charges';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end
          else
             break;
          end;
       end;

//PT93-5
   for i := 1 to 5 do
       begin
       if ((RegimeComp<>NIL) and ((Pos ('A', RegimeComp.CellValues[0,i])=1) or
          ((Pos ('C', RegimeComp.CellValues[0,i])=1) and
          (IsNumeric (Copy (RegimeComp.CellValues[0,i], 1, 3)))) or
          (Pos ('G', RegimeComp.CellValues[0,i])=1)) and
          (GetControlText ('CCONTRAT')<>'CDI') and
          (GetControlText ('CCONTRAT')<>'CCD') and
          (GetControlText ('CCONTRAT')<>'CTT') and
          (GetControlText ('CCONTRAT')<>'CAA') and
          (GetControlText ('CCONTRAT')<>'CAP') and
          (GetControlText ('CCONTRAT')<>'CEJ') and
          (GetControlText ('CCONTRAT')<>'CES') and
          (GetControlText ('CCONTRAT')<>'CNE') and
          (GetControlText ('CCONTRAT')<>'CAU') and
          (GetControlText ('CCONTRAT')<>'CAV') and
          (GetControlText ('CCONTRAT')<>'SC')) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence code contrat de'+
                 ' travail/régime complémentaire';
          SetFocusControl ('CCONTRAT');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.00.012.001';
          ErreurDADSU.Explication:= 'Incohérence contrôle C2-04 du cahier des'+
                                    ' charges';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
          break;
          end;
      end;

   if ((GetControlText ('CCONTRAT')='20') and
      (GetControlText ('CSTATUTPROF')<>'40') and
      (GetControlText ('CSTATUTPROF')<>'41') and
      (GetControlText ('CSTATUTPROF')<>'42') and
      (GetControlText ('CSTATUTPROF')<>'43') and
      (GetControlText ('CSTATUTPROF')<>'44')) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code contrat de'+
             ' travail/statut professionnel';
      SetFocusControl ('CCONTRAT');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.012.001';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-05 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;
//FIN PT93-5

   for i := 1 to 5 do
       begin
       if (((RegimeComp.CellValues[0,i] = 'F001') or
          (RegimeComp.CellValues[0,i] = 'F002')) and
          (GetControlText ('CCONDEMPLOI') <> 'C') and
          (GetControlText ('CCONDEMPLOI') <> 'P') and
          (GetControlText ('CCONDEMPLOI') <> 'N')) then
          begin
          SetControlText ('CCONDEMPLOI', 'C');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.00.013';
{PT91
          ErreurDADSU.Explication:= 'La condition d''emploi a été forcée à '+
                                    '"Temps plein"';
}
          ErreurDADSU.Explication:= 'La caractéristique activité a été forcée'+
                                    ' à "Temps plein"';
//FIN PT91
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur (ErreurDADSU);
          break;
          end;
      end;

   if (GetControlText('CCONDEMPLOI')='') then
      begin
      SetControlText ('CCONDEMPLOI', 'W');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.013';
{PT91
      ErreurDADSU.Explication:= 'La condition d''emploi a été forcée à '+
                                '"Sans spécificité de l''activité"';
}
      ErreurDADSU.Explication:= 'La caractéristique activité a été forcée à'+
                                ' "Sans spécificité de l''activité"';
//FIN PT91
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

//PT93-2
   for i:= 1 to 5 do
       begin
       if ((RegimeComp<>NIL) and ((Pos ('A', RegimeComp.CellValues[0,i])=1) or
          ((Pos ('C', RegimeComp.CellValues[0,i])=1) and
          (IsNumeric (Copy (RegimeComp.CellValues[0,i], 1, 3)))) or
          (Pos ('G', RegimeComp.CellValues[0,i])=1)) and
          (GetControlText('CCONDEMPLOI')='N')) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence caractéristique activité/régime'+
                 ' complémentaire';
          SetFocusControl ('CCONDEMPLOI');
          ErreurDADSU.Segment:= 'S41.G01.00.013';
          ErreurDADSU.Explication:= 'Incohérence caractéristique activité/'+
                                    'régime complémentaire';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur(ErreurDADSU);
          end;
       end;
//FIN PT93-2

//PT93-2
   for i:= 1 to 5 do
       begin
       if ((RegimeComp<>NIL) and ((Pos ('A', RegimeComp.CellValues[0,i])=1) or
          ((Pos ('C', RegimeComp.CellValues[0,i])=1) and
          (IsNumeric (Copy (RegimeComp.CellValues[0,i], 1, 3)))) or
          (Pos ('G', RegimeComp.CellValues[0,i])=1)) and
          ((GetControlText ('CSTATUTPROF')<>'01') and
          (GetControlText ('CSTATUTPROF')<>'02') and
          (GetControlText ('CSTATUTPROF')<>'03') and
          (GetControlText ('CSTATUTPROF')<>'04') and
          (GetControlText ('CSTATUTPROF')<>'05') and
          (GetControlText ('CSTATUTPROF')<>'06') and
          (GetControlText ('CSTATUTPROF')<>'07') and
          (GetControlText ('CSTATUTPROF')<>'08') and
          (GetControlText ('CSTATUTPROF')<>'09') and
          (GetControlText ('CSTATUTPROF')<>'11') and
          (GetControlText ('CSTATUTPROF')<>'12') and
          (GetControlText ('CSTATUTPROF')<>'13') and
          (GetControlText ('CSTATUTPROF')<>'14') and
          (GetControlText ('CSTATUTPROF')<>'15') and
          (GetControlText ('CSTATUTPROF')<>'16') and
          (GetControlText ('CSTATUTPROF')<>'18') and
          (GetControlText ('CSTATUTPROF')<>'19') and
          (GetControlText ('CSTATUTPROF')<>'20') and
          (GetControlText ('CSTATUTPROF')<>'21') and
          (GetControlText ('CSTATUTPROF')<>'22') and
          (GetControlText ('CSTATUTPROF')<>'23') and
          (GetControlText ('CSTATUTPROF')<>'24') and
          (GetControlText ('CSTATUTPROF')<>'25') and
          (GetControlText ('CSTATUTPROF')<>'26') and
          (GetControlText ('CSTATUTPROF')<>'27') and
          (GetControlText ('CSTATUTPROF')<>'28') and
          (GetControlText ('CSTATUTPROF')<>'75') and
          (GetControlText ('CSTATUTPROF')<>'90'))) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence statut professionnel/régime'+
                 ' complémentaire';
          SetFocusControl ('CSTATUTPROF');
          ErreurDADSU.Segment:= 'S41.G01.00.014';
          ErreurDADSU.Explication:= 'Incohérence statut professionnel/régime'+
                                    ' complémentaire';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur(ErreurDADSU);
          break;
          end;
       end;

   for i:= 1 to 5 do
       begin
       if ((RegimeComp<>NIL) and ((Pos ('A', RegimeComp.CellValues[0,i])=1) or
          ((Pos ('C', RegimeComp.CellValues[0,i])=1) and
          (IsNumeric (Copy (RegimeComp.CellValues[0,i], 1, 3)))) or
          (Pos ('G', RegimeComp.CellValues[0,i])=1)) and
          ((GetControlText ('CSTATUTPROF')<>'01') and
          (GetControlText ('CSTATUTPROF')<>'02') and
          (GetControlText ('CSTATUTPROF')<>'03') and
          (GetControlText ('CSTATUTPROF')<>'04') and
          (GetControlText ('CSTATUTPROF')<>'05') and
          (GetControlText ('CSTATUTPROF')<>'06') and
          (GetControlText ('CSTATUTPROF')<>'07') and
          (GetControlText ('CSTATUTPROF')<>'08') and
          (GetControlText ('CSTATUTPROF')<>'09') and
          (GetControlText ('CSTATUTPROF')<>'11') and
          (GetControlText ('CSTATUTPROF')<>'12') and
          (GetControlText ('CSTATUTPROF')<>'13') and
          (GetControlText ('CSTATUTPROF')<>'14') and
          (GetControlText ('CSTATUTPROF')<>'15') and
          (GetControlText ('CSTATUTPROF')<>'16') and
          (GetControlText ('CSTATUTPROF')<>'18') and
          (GetControlText ('CSTATUTPROF')<>'19') and
          (GetControlText ('CSTATUTPROF')<>'20') and
          (GetControlText ('CSTATUTPROF')<>'21') and
          (GetControlText ('CSTATUTPROF')<>'22') and
          (GetControlText ('CSTATUTPROF')<>'23') and
          (GetControlText ('CSTATUTPROF')<>'24') and
          (GetControlText ('CSTATUTPROF')<>'25') and
          (GetControlText ('CSTATUTPROF')<>'26') and
          (GetControlText ('CSTATUTPROF')<>'27') and
          (GetControlText ('CSTATUTPROF')<>'28') and
          (GetControlText ('CSTATUTPROF')<>'75') and
          (GetControlText ('CSTATUTPROF')<>'90'))) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence statut professionnel/régime'+
                 ' complémentaire';
          SetFocusControl ('CSTATUTPROF');
          ErreurDADSU.Segment:= 'S41.G01.00.014';
          ErreurDADSU.Explication:= 'Incohérence statut professionnel/régime'+
                                    ' complémentaire';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur(ErreurDADSU);
          break;
          end;
       end;
       
//FIN PT93-2

   for i := 1 to 5 do
       begin
       if (((RegimeComp.CellValues[0,i] = 'F001') or
          (RegimeComp.CellValues[0,i] = 'F002')) and
          (GetControlText ('CSTATUTPROF') <> '60') and
          (GetControlText ('CSTATUTPROF') <> '62') and
          (GetControlText ('CSTATUTPROF') <> '63') and
          (GetControlText ('CSTATUTPROF') <> '65') and
          (GetControlText ('CSTATUTPROF') <> '66') and
          (GetControlText ('CSTATUTPROF') <> '68')) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Incohérence code statut professionnel/régime'+
                 ' complémentaire';
          SetFocusControl ('CSTATUTPROF');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.00.014';
          ErreurDADSU.Explication:= 'Incohérence contrôle C2-01 du cahier des'+
                                    ' charges';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
          break;
          end;
       end;

   if (GetControlText('CSTATUTPROF')='') then
      begin
      SetControlText ('CSTATUTPROF', '90');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Le statut professionnel a été forcé à "Pas'+
                                ' de statut"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end
   else
   if (((GetControlText('CSTATUTPROF')='61') or
      (GetControlText('CSTATUTPROF')='64') or
{PT93-5
      (GetControlText('CSTATUTPROF')='67') or
      (GetControlText('CSTATUTPROF')='70')) and
}
      (GetControlText('CSTATUTPROF')='67')) and
//FIN PT93-5
      (((GetControlText('CREGIMESS') <> '200') and
      (GetControlText('CHTYPEREGIME') = '-')) or   //PT92-1
      ((GetControlText('CREGIMEAT') <> '200') and
      (GetControlText('CHTYPEREGIME') = 'X')))) then   //PT92-1
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code statut professionnel/régime SS';
      SetFocusControl ('CSTATUTPROF');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-02 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;

   for i := 1 to 5 do
       begin
{PT74
       if ((RegimeComp.CellValues[0,i] <> 'I001') and
          ((GetControlText('CSTATUTPROF')='52') or
          (GetControlText('CSTATUTPROF')='53') or
          (GetControlText('CSTATUTPROF')='54') or
          (GetControlText('CSTATUTPROF')='56') or
          (GetControlText('CSTATUTPROF')='57') or
          (GetControlText('CSTATUTPROF')='58') or
          (GetControlText('CSTATUTPROF')='59'))) then
}
{PT93-5
       if (RegimeComp.CellValues[0,i] <> 'I001') then
}
{PT95-3
       if ((RegimeComp.CellValues[0,i]<>'I001') and
          (RegimeComp.CellValues[0,i]<>'90000')) then
}
       if ((RegimeComp.CellValues[0,i]<>'') and
          (RegimeComp.CellValues[0,i]<>'I001') and
          (RegimeComp.CellValues[0,i]<>'90000')) then
//FIN PT95-3
//FIN PT93-5
          begin
          if ((GetControlText('CSTATUTPROF')='52') or
             (GetControlText('CSTATUTPROF')='53') or
             (GetControlText('CSTATUTPROF')='54') or
             (GetControlText('CSTATUTPROF')='56') or
             (GetControlText('CSTATUTPROF')='57') or
             (GetControlText('CSTATUTPROF')='58') or
             (GetControlText('CSTATUTPROF')='59')) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Incohérence code statut professionnel/régime'+
                    ' complémentaire';
             SetFocusControl ('CSTATUTPROF');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.014';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2-03 du cahier des'+
                                       ' charges';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end
          else
             break;
          end;
       end;

   if ((RegimeComp <> NIL) and ((RegimeComp.CellValues[0,1] = 'I002') or
      (RegimeComp.CellValues[0,2] = 'I002') or
      (RegimeComp.CellValues[0,3] = 'I002') or
      (RegimeComp.CellValues[0,4] = 'I002') or
      (RegimeComp.CellValues[0,5] = 'I002')) and
      (GetControlText('CSTATUTPROF') <> '40') and
      (GetControlText('CSTATUTPROF') <> '41') and
      (GetControlText('CSTATUTPROF') <> '42') and
{PT93-5
      (GetControlText('CSTATUTPROF') <> '43')) then
}
      (GetControlText('CSTATUTPROF') <> '43') and
      (GetControlText('CSTATUTPROF') <> '44')) then
//FIN PT93-5
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code statut professionnel/régime'+
             ' complémentaire';
      SetFocusControl ('CSTATUTPROF');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Un assuré relevant de l''IRCANTEC élus doit'+
                                ' avoir un "statut professionnel" de type "élu"';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;

   if ((RegimeComp <> NIL) and ((RegimeComp.CellValues[0,1] = 'I001') or
      (RegimeComp.CellValues[0,2] = 'I001') or
      (RegimeComp.CellValues[0,3] = 'I001') or
      (RegimeComp.CellValues[0,4] = 'I001') or
      (RegimeComp.CellValues[0,5] = 'I001')) and
      ((GetControlText('CSTATUTPROF') = '40') or
      (GetControlText('CSTATUTPROF') = '41') or
      (GetControlText('CSTATUTPROF') = '42') or
{PT93-5
      (GetControlText('CSTATUTPROF') = '43'))) then
}
      (GetControlText('CSTATUTPROF') = '43') or
      (GetControlText('CSTATUTPROF') = '44'))) then
//FIN PT93-5
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code statut professionnel/régime'+
             ' complémentaire';
      SetFocusControl ('CSTATUTPROF');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Un assuré relevant de l''IRCANTEC ne doit pas'+
                                ' avoir un "statut professionnel" de type "élu"';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;

//PT93-5
   if ((GetControlText ('CSTATUTPROF')='36') and
      ((GetControlText ('ECASPARTV')='') or
      (GetControlText ('ECASPARTV')='0'))) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Montant relatif au volontariat associatif absent';
      SetFocusControl ('CSTATUTPROF');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Le montant relatif au volontariat associatif'+
                                ' doit être renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;
//FIN PT93-5

//PT105
   if ((GetControlText ('CSTATUTPROF')<>'36') and
      (GetControlText ('ECASPARTV')<>'') and
      (GetControlText ('ECASPARTV')<>'0')) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Statut professionnel incohérent';
      SetFocusControl ('CSTATUTPROF');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.014';
      ErreurDADSU.Explication:= 'Le montant relatif au volontariat associatif'+
                                ' est incompatible avec le statut professionnel';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur (ErreurDADSU);
      end;
//FIN PT105

   if (GetControlText('CSTATUTCAT')='') then
      begin
      SetControlText ('CSTATUTCAT', '90');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.015.001';
      ErreurDADSU.Explication:= 'Le statut catégoriel a été forcé à "Pas de'+
                                ' statut"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

//PT93-2
   if (((GetControlText('CSTATUTPROF')='01') or
      (GetControlText('CSTATUTPROF')='02') or
      (GetControlText('CSTATUTPROF')='14')) and
      (GetControlText('CSTATUTCAT')<>'04')) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code statut professionnel/code statut'+
             ' catégoriel';
      SetFocusControl ('CSTATUTCAT');
      ErreurDADSU.Segment:= 'S41.G01.00.015.002';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-02 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;

   if (((GetControlText('CSTATUTPROF')='12') or
      (GetControlText('CSTATUTPROF')='13') or
      (GetControlText('CSTATUTPROF')='22') or
      (GetControlText('CSTATUTPROF')='23') or
      (GetControlText('CSTATUTPROF')='24') or
      (GetControlText('CSTATUTPROF')='25')) and
      (GetControlText('CSTATUTCAT')<>'01')) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Incohérence code statut professionnel/code statut'+
             ' catégoriel';
      SetFocusControl ('CSTATUTCAT');
      ErreurDADSU.Segment:= 'S41.G01.00.015.002';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-03 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT93-2

   if (GetControlText('ECONVCOLL')='') then
      begin
      SetControlText ('ECONVCOLL', '9999');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.016';
      ErreurDADSU.Explication:= 'Le code Convention collective a été forcé à'+
                                ' "9999"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

//DEB PT92-1
   if (GetControlText('CHTYPEREGIME')='-') then
      begin
      if (GetControlText ('CREGIMESS')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime de base S.S. absent';
         SetFocusControl ('CREGIMESS');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.001';
         ErreurDADSU.Explication:= 'Le régime de base S.S. n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;

//PT92-2
// pt101 Suppression des controles, les codes autorisés sont gérés dans la tablette
{      if ((GetControlText ('CREGIMESS')='147') or
         (GetControlText ('CREGIMESS')='148') or
         (GetControlText ('CREGIMESS')='149')) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime de base S.S. erroné';
         SetFocusControl ('CREGIMESS');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.001';
         ErreurDADSU.Explication:= 'Le régime de base S.S. est renseigné avec'+
                                   ' une valeur interdite';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end; }
//FIN PT92-2
//PT93-2
      if ((GetControlText ('CREGIMESS')='200') and
         ((GetControlText ('EDEBMOTIF1')='085') or
         (GetControlText ('EDEBMOTIF2')='085') or
         (GetControlText ('EDEBMOTIF3')='085')) and
         ((GetControlText ('EFINMOTIF1')='086') or
         (GetControlText ('EFINMOTIF2')='086') or
         (GetControlText ('EFINMOTIF3')='086'))) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Incohérence statut professionnel/régime SS';
         SetFocusControl ('CREGIMESS');
         ErreurDADSU.Segment:= 'S41.G01.00.018.001';
         ErreurDADSU.Explication:= 'Incohérence statut professionnel/régime SS';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//FIN PT93-2
      end
   else
      begin
      if (GetControlText ('CREGIMEMAL')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire risque maladie absent';
         SetFocusControl ('CREGIMEMAL');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.002';
         ErreurDADSU.Explication:= 'Le régime obligatoire risque maladie'+
                                   ' n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//PT92-2
// Pt101 suppression des controles , les codes autorisés sont gérés dans la tablette
 {     if ((GetControlText ('CREGIMEMAL')='147') or
         (GetControlText ('CREGIMEMAL')='148') or
         (GetControlText ('CREGIMEMAL')='150') or
         (GetControlText ('CREGIMEMAL')='160')) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire risque maladie erroné';
         SetFocusControl ('CREGIMEMAL');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.002';
         ErreurDADSU.Explication:= 'Le régime obligatoire risque maladie est'+
                                   ' renseigné avec une valeur interdite';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;   }
//FIN PT92-2

      if (GetControlText ('CREGIMEAT')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire risque AT absent';
         SetFocusControl ('CREGIMEAT');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.003';
         ErreurDADSU.Explication:= 'Le régime obligatoire risque AT n''est pas'+
                                   ' renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//PT93-2
      if ((GetControlText ('CREGIMEAT')<>'901') and
         ((GetControlText ('EDEBMOTIF1')='085') or
         (GetControlText ('EDEBMOTIF2')='085') or
         (GetControlText ('EDEBMOTIF3')='085')) and
         ((GetControlText ('EFINMOTIF1')='086') or
         (GetControlText ('EFINMOTIF2')='086') or
         (GetControlText ('EFINMOTIF3')='086'))) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Incohérence statut professionnel/régime '+
                ' obligatoire risque AT';
         SetFocusControl ('CREGIMEAT');
         ErreurDADSU.Segment:= 'S41.G01.00.018.003';
         ErreurDADSU.Explication:= 'Incohérence statut professionnel/régime'+
                                   ' obligatoire risque AT';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//FIN PT93-2

      if (GetControlText ('CREGIMEVIP')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire vieillesse (PP) absent';
         SetFocusControl ('CREGIMEVIP');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.004';
         ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PP)'+
                                   ' n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//PT92-2
// pt101 suppression des controles , les codes autorisés sont gérés dans la tablette
  {    if ((GetControlText ('CREGIMEVIP')='150') or
         (GetControlText ('CREGIMEVIP')='160')) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire vieillesse (PP) erroné';
         SetFocusControl ('CREGIMEVIP');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.004';
         ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PP) est'+
                                   ' renseigné avec une valeur interdite';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;  }
//FIN PT92-2

      if (GetControlText ('CREGIMEVIS')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire vieillesse (PS) absent';
         SetFocusControl ('CREGIMEVIS');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.005';
         ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PP)'+
                                   ' n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//PT92-2
//Pt101 : suppression des controles, les codes autorisés sont gérés dans la tablette
  {    if ((GetControlText ('CREGIMEVIS')='150') or
         (GetControlText ('CREGIMEVIS')='160')) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Régime obligatoire vieillesse (PS) erroné';
         SetFocusControl ('CREGIMEVIS');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.018.005';
         ErreurDADSU.Explication:= 'Le régime obligatoire vieillesse (PS) est'+
                                   ' renseigné avec une valeur interdite';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;  }
//FIN PT92-2
      end;
//FIN PT92-1
//FIN PT73-1

   if (GetControlText('CCONDEMPLOI') = 'P') then
      begin
      if (GetControlText('ETAUXPARTIEL') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Taux d''activité à temps partiel';
         SetFocusControl('ETAUXPARTIEL');
         end
      else
      if (StrToFloat(GetControlText('ETAUXPARTIEL')) >= 100) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Le taux d''activité à temps partiel doit être'+
                ' inférieur à 100';
         SetFocusControl('ETAUXPARTIEL');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.020';
         ErreurDADSU.Explication:= 'Le taux d''activité à temps partiel doit'+
                                   ' être inférieur à 100';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;

//PT73-1
      for i := 1 to 5 do
          begin
          if (((RegimeComp.CellValues[0,i] = 'F001') or
             (RegimeComp.CellValues[0,i] = 'F002')) and
             (StrToFloat(GetControlText('ETAUXPARTIEL')) < 50)) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Incohérence taux de travail à temps partiel'+
                    '/régime complémentaire';
             SetFocusControl ('ETAUXPARTIEL');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.020';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2-02 du cahier'+
                                       ' des charges';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
      end;
//FIN PT73-1

   if ((GetControlText('ELIBELLEEMPLOI') = '') or
      (GetControlText('LLIBELLEEMPLOI_') = '') or
      (GetControlText('LLIBELLEEMPLOI_') = '...')) then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Libellé d''emploi absent';
      SetFocusControl('ELIBELLEEMPLOI');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.020';
      ErreurDADSU.Explication:= 'Le libellé d''emploi n''est pas renseigné';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

//PT73-1
   if ((GetControlText ('EHEURESTRAV') <> '') and
      (GetControlText ('EHEURESSAL') <> '') and
      (StrToInt (GetControlText ('EHEURESTRAV')) > StrToInt (GetControlText ('EHEURESSAL')))) then
      begin
{PT81
      SetControlText ('EHEURESTRAV', GetControlText ('EHEURESSAL'));
}
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.021';
{PT81
      ErreurDADSU.Explication:= 'Le nombre d''heures travaillées était'+
}
      ErreurDADSU.Explication:= 'Le nombre d''heures travaillées est'+
//FIN PT81
                                ' supérieur au total d''heures payées';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT73-1

//PT78
{PT92-1
   if ((GetControlText('CREGIMESS') = '200') and
      (GetControlText('CSTATUTPROF') <> '07')) then
}
{PT93-5
   if ((((GetControlText('CREGIMESS') = '200') and
      (GetControlText('CHTYPEREGIME') = '-')) or
      ((GetControlText('CREGIMEAT') = '200') and
      (GetControlText('CHTYPEREGIME') = 'X'))) and
      (GetControlText('CSTATUTPROF') <> '07')) then
}
   if ((((GetControlText('CREGIMESS') = '200') and
      (GetControlText('CHTYPEREGIME') = '-')) or
      ((GetControlText('CREGIMEAT') = '200') and
      (GetControlText('CHTYPEREGIME') = 'X'))) and
      (GetControlText('CSTATUTPROF') <> '07') and
      (GetControlText('CSTATUTPROF') <> '36') and
      (GetControlText('CSTATUTPROF') <> '37') and
      (GetControlText('CSTATUTPROF') <> '38') and
      (GetControlText('CSTATUTPROF') <> '39')) then
//FIN PT93-5
      begin
{PT73-5
      if (GetControlText('ESECTIONAT') = '') then
}
      if ((GetControlText('ESECTIONAT') = '') or
         (GetControlText('ESECTIONAT') = '98')) then
//FIN PT73-5
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Section AT absente';
         SetFocusControl('ESECTIONAT');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.025';
         ErreurDADSU.Explication:= 'La section AT n''est pas renseignée';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;

      Buffer:= PGUpperCase(GetControlText('ERISQUEAT'));
{PT93-5
      if (((GetControlText('CCONTRAT')<>'CTT') and ((Buffer='745BA') or
         (Buffer='745BB') or (Buffer='745BD'))) or
         ((GetControlText('CSTATUTPROF')<>'05') and (Buffer='923AC'))) then
}
      if ((GetControlText('CSTATUTPROF')<>'05') and (Buffer='923AC')) then
//FIN PT93-5
         SetControlText('ERISQUEAT', '');

{PT73-5
      if (GetControlText('ERISQUEAT') = '') then
}
      if ((GetControlText('ERISQUEAT') = '') or
         (GetControlText('ERISQUEAT') = '98888')) then
//FIN PT73-5
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Risque AT absent';
         SetFocusControl('ERISQUEAT');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.026';
         ErreurDADSU.Explication:= 'Le code risque AT n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;

{PT73-5
      if (GetControlText('ETAUXAT') = '') then
}
      if ((GetControlText ('ETAUXAT') = '') or
         (GetControlText ('ETAUXAT') = FloatToStr (9888/100))) then
//FIN PT73-5
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Taux AT absent';
         SetFocusControl('ETAUXAT');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.028';
         ErreurDADSU.Explication:= 'Le Taux AT n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;
   end;
//FIN PT78

   if (GetControlText('EBRUTSS') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Base brute S.S. absente';
      SetFocusControl('EBRUTSS');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.029.001';
      ErreurDADSU.Explication:= 'La base brute S.S. n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

   if (GetControlText('EPLAFONDSS') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Base limitée au plafond S.S. absente';
      SetFocusControl('EPLAFONDSS');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.030.001';
      ErreurDADSU.Explication:= 'La base limitée au plafond S.S. n''est pas'+
                                ' renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

   if (GetControlText('EBASECSG') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Base C.S.G. absente';
      SetFocusControl('EBASECSG');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.032.001';
      ErreurDADSU.Explication:= 'La Base C.S.G. n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

   if (GetControlText('EBASECRDS') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Base C.R.D.S. absente';
      SetFocusControl('EBASECRDS');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.033.001';
      ErreurDADSU.Explication:= 'La Base C.R.D.S. n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

   if (GetControlText('EBRUTFISC') = '') then
      begin
      result:= FALSE;
      Mess:= Mess+'#13#10 - Base brute fiscale absente';
      SetFocusControl('EBRUTFISC');
//PT73-1
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.035.001';
      ErreurDADSU.Explication:= 'La Base brute fiscale n''est pas renseignée';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
//FIN PT73-1
      end;

   BoolFraisProf:= FALSE;
   if ((GetControlText('EAVANTAGENAT') > '0') and
      (GetControlText('EAVANTAGENAT') <> '')) then
      begin
      if ((Nourriture <> NIL) and (Nourriture.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((Logement <> NIL) and (Logement.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((Voiture <> NIL) and (Voiture.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((AutreAvant <> NIL) and (AutreAvant.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((NTIC <> NIL) and (NTIC.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if (BoolFraisProf = FALSE) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Coche(s) Avantage en nature absente';
         SetFocusControl('EAVANTAGENAT');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.037.001';
         ErreurDADSU.Explication:= 'Aucune coche "Avantage en nature" n''est'+
                                   ' renseignée';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;
      end;

   BoolFraisProf:= FALSE;
   if ((GetControlText('EFRAISPROF') > '0') and
      (GetControlText('EFRAISPROF') <> '')) then
      begin
      if ((AllocForfait <> NIL) and (AllocForfait.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((RembJustif <> NIL) and (RembJustif.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((PriseCharge <> NIL) and (PriseCharge.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if ((AutreDepense <> NIL) and (AutreDepense.Checked = TRUE)) then
         BoolFraisProf:= TRUE;

      if (BoolFraisProf = FALSE) then
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Coche(s) Frais Professionnels absente';
         SetFocusControl('EFRAISPROF');
//PT73-1
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G01.00.044.001';
         ErreurDADSU.Explication:= 'Aucune coche "Frais Professionnels" n''est'+
                                   ' renseignée';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
//FIN PT73-1
         end;
      end;

//PT73-1
   if (GetControlText('EREVENUSACTIV')=GetControlText('EAUTREREV')) then
      begin
      result:= FALSE;
      Mess:= Mess+ '#13#10 - Montant des autres revenus nets imposables erroné';
      SetFocusControl('EAUTREREV');
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.00.066.001';
      ErreurDADSU.Explication:= 'Incohérence contrôle C2-01 du cahier des'+
                                ' charges';
      ErreurDADSU.CtrlBloquant:= True;
      EcrireErreur(ErreurDADSU);
      end;
//FIN PT73-1

//PT93-5
   for i := 1 to 5 do
       begin
       if (RegimeComp.CellValues[0,i]='CNBF') then
          begin
          if (GetControlText ('CCNBF')='') then
             begin
             result:= False;
             Mess:= Mess+'#13#10 - Classe d''extension CNBF non renseignée';
             SetFocusControl('CCNBF');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.071';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2 du cahier des'+
                                       ' charges';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end
       else
       if ((RegimeComp.CellValues[0,i]<>'CNBF') and
          (RegimeComp.CellValues[0,i]<>'')) then
          begin
          if (GetControlText ('CCNBF')<>'') then
             begin
             result:= False;
             Mess:= Mess+'#13#10 - Organisme CNBF non renseigné';
             SetFocusControl('CCNBF');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.00.071';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2 du cahier des'+
                                       ' charges';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
       end;
//FIN PT93-5

   if ((RegimeComp <> NIL) and (RegimeComp.CellValues[0,1] = '') and
      (RegimeComp.CellValues[0,2] = '') and
      (RegimeComp.CellValues[0,3] = '') and
      (RegimeComp.CellValues[0,4] = '') and
      (RegimeComp.CellValues[0,5] = '')) then
{PT73-1
      RegimeComp.CellValues[0,1]:= '90000';
}
      begin
      RegimeComp.CellValues[0,1]:= '90000';
      ErreurDADSU.Num:= Ordre;
      ErreurDADSU.Segment:= 'S41.G01.01.001';
      ErreurDADSU.Explication:= 'Le code régime complémentaire a été forcé à'+
                                ' "90000"';
      ErreurDADSU.CtrlBloquant:= False;
      EcrireErreur(ErreurDADSU);
      end;

   for i := 1 to 5 do
       begin
       if (((RegimeComp.CellValues[0,i] = 'I001') or
          (RegimeComp.CellValues[0,i] = 'I002') or
          (RegimeComp.CellValues[0,i] = 'F001') or
          (RegimeComp.CellValues[0,i] = 'F002') or
          (RegimeComp.CellValues[0,i] = 'R001')) and
          (RegimeComp.CellValues[2,i] = '')) then
          begin
          result:= FALSE;
          Mess:= Mess+ '#13#10 - Numéro de rattachement régime complémentaire absent';
          SetFocusControl('GREGIMECOMP');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.01.001';
          ErreurDADSU.Explication:= 'Le numéro de rattachement régime'+
                                    ' complémentaire n''est pas renseigné';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
          break;
          end
       else
       if ((RegimeComp.CellValues[0,i] = 'I001') or
          (RegimeComp.CellValues[0,i] = 'I002')) then
          begin
          if (Length (RegimeComp.CellValues[2,i])=8) then
             rep:= 0
          else
             rep:= 1;

          if (rep = 0) then
             begin
             Buffer:= Copy (RegimeComp.CellValues[2,i], 2, 3)+
                      Copy (RegimeComp.CellValues[2,i], 7, 2);
{PT85
             rep:= ControleCar (Buffer, '0', False);
}
{PT100
             rep:= ControleCar (Buffer, 'A', False);
}
             rep:= ControleCar (Buffer, 'A', False, CInterdit);
//FIN PT100
             end;

          if (rep = 0) then
             begin
             Buffer:= '1'+Copy (RegimeComp.CellValues[2,i], 1, 1)+
                      Copy (RegimeComp.CellValues[2,i], 5, 2);
{PT100
             rep:= ControleCar (Buffer, 'N', False);
}
             rep:= ControleCar (Buffer, 'N', False, CInterdit);
//FIN PT100
             end;

          if (rep <> 0) then
             begin
             result:= FALSE;
             Mess:= Mess+ '#13#10 - Numéro de rattachement régime complémentaire erroné';
             RegimeComp.CellValues[2,i]:= '';
             SetFocusControl('GREGIMECOMP');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.01.001';
             ErreurDADSU.Explication:= 'Le numéro de rattachement régime'+
                                       ' complémentaire est incorrect';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
       end;

   for i := 1 to 5 do
       begin
       if ((RegimeComp.CellValues[0,i] = 'I001') or
          (RegimeComp.CellValues[0,i] = 'I002')) then
          begin
          for j := 1 to 5 do
              begin
              if ((RegimeComp.CellValues[0,j] <> 'I001') and
                 (RegimeComp.CellValues[0,j] <> 'I002') and
                 (RegimeComp.CellValues[0,j] <> '')) then
                 begin
                 Mess:= Mess+'#13#10 - Incohérence entre régimes'+
                        ' complémentaires';
                 ErreurDADSU.Num:= Ordre;
                 ErreurDADSU.Segment:= 'S41.G01.01.001';
                 ErreurDADSU.Explication:= 'Il existe un code régime'+
                                           ' complémentaire IRCANTEC associé à'+
                                           ' un autre code régime';
                 ErreurDADSU.CtrlBloquant:= False;
                 EcrireErreur (ErreurDADSU);
                 break;
                 end;
              end;
          end;
       end;

//PT93-5
   for i := 1 to 5 do
       begin
       if (RegimeComp.CellValues[0,i]='CNBF') then
          begin
          if ((GetControlText ('CREGIMEVIP')<>'157') or
             (GetControlText ('CREGIMEVIS')<>'157')) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Incohérence régime complémentaire/régime'+
                    ' vieillesse';
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.01.001';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2-09 du cahier'+
                                       ' des charges';
             ErreurDADSU.CtrlBloquant:= False;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end
       else
          begin
{PT95-2
          if ((GetControlText ('CREGIMEVIP')='157') or
             (GetControlText ('CREGIMEVIS')='157')) then
}
          if ((RegimeComp.CellValues[0,i]<>'') and
             ((GetControlText ('CREGIMEVIP')='157') or
             (GetControlText ('CREGIMEVIS')='157'))) then
//FIN PT95-2
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Incohérence régime complémentaire/régime'+
                    ' vieillesse';
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.01.001';
             ErreurDADSU.Explication:= 'Incohérence contrôle C2-10 du cahier'+
                                       ' des charges';
             ErreurDADSU.CtrlBloquant:= False;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
       end;
//FIN PT93-5

//PT93-3
   ctlbasebrute := true;
   for i:= 1 to 4 do
       begin
       if ((Basebrutespec.CellValues[0,i]='01') or
          (Basebrutespec.CellValues[0,i]='03') or
          (Basebrutespec.CellValues[0,i]='04') or
          (Basebrutespec.CellValues[0,i]='05') or
          (Basebrutespec.CellValues[0,i]='06') or
          (Basebrutespec.CellValues[0,i]='08') or
          (Basebrutespec.CellValues[0,i]='09') or
          (Basebrutespec.CellValues[0,i]='10') or
          (Basebrutespec.CellValues[0,i]='11') or
          (Basebrutespec.CellValues[0,i]='12') or
          (Basebrutespec.CellValues[0,i]='14') or
          (Basebrutespec.CellValues[0,i]='15') or
          (Basebrutespec.CellValues[0,i]='16') or
          (Basebrutespec.CellValues[0,i]='19') or
          (Basebrutespec.CellValues[0,i]='21') or
          (Basebrutespec.CellValues[0,i]='24') or
          (Basebrutespec.CellValues[0,i]='25') or
          (Basebrutespec.CellValues[0,i]='26') or
          (Basebrutespec.CellValues[0,i]='27') or
          (Basebrutespec.CellValues[0,i]='28') or
          (Basebrutespec.CellValues[0,i]='57')) then
          begin
          ctlbasebrute:= False;
          for j:= 1 to 4 do
              begin
              if (Baseplafspec.CellValues[0,j]=Basebrutespec.CellValues[0,i]) then
                 ctlbasebrute:= True;
              end;
          if (ctlbasebrute=False) then
             begin
             result := false;
             Mess:= Mess+ '#13#10 - Incohérence entre bases brutes et bases'+
                    ' plafonnées';
             SetFocusControl ('GBASEPLAFSPEC');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.03.001';
             ErreurDADSU.Explication:= 'Il existe une base brute'+
                                       ' exceptionnelle sans base plafonnée'+
                                       ' exceptionnelle';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
       end;

   ctlbaseplaf := true;
   for i:= 1 to 4 do
       begin
       if ((Baseplafspec.CellValues[0,i]='01') or
          (Baseplafspec.CellValues[0,i]='03') or
          (Baseplafspec.CellValues[0,i]='04') or
          (Baseplafspec.CellValues[0,i]='05') or
          (Baseplafspec.CellValues[0,i]='06') or
          (Baseplafspec.CellValues[0,i]='08') or
          (Baseplafspec.CellValues[0,i]='09') or
          (Baseplafspec.CellValues[0,i]='10') or
          (Baseplafspec.CellValues[0,i]='11') or
          (Baseplafspec.CellValues[0,i]='12') or
          (Baseplafspec.CellValues[0,i]='14') or
          (Baseplafspec.CellValues[0,i]='15') or
          (Baseplafspec.CellValues[0,i]='16') or
          (Baseplafspec.CellValues[0,i]='19') or
          (Baseplafspec.CellValues[0,i]='21') or
          (Baseplafspec.CellValues[0,i]='24') or
          (Baseplafspec.CellValues[0,i]='25') or
          (Baseplafspec.CellValues[0,i]='26') or
          (Baseplafspec.CellValues[0,i]='27') or
          (Baseplafspec.CellValues[0,i]='28') or
          (Baseplafspec.CellValues[0,i]='57')) then
          begin
          ctlbaseplaf:= False;
          for j := 1 to 4 do
              begin
              if (Basebrutespec.CellValues[0,j]=Baseplafspec.CellValues[0,i]) then
                 ctlbaseplaf:= True;
              end;
          if (ctlbaseplaf=False) then
             begin
             result := false;
             Mess:= Mess+ '#13#10 - Incohérence entre bases brutes et bases'+
                    ' plafonnées';
             SetFocusControl ('GBASEBRUTESPEC');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.03.001';
             ErreurDADSU.Explication:= 'Il existe une base plafonnée'+
                                       ' exceptionnelle sans base brute'+
                                       ' exceptionnelle';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             break;
             end;
          end;
       end;
//FIN PT93-3

// Deb PT93-6
   ctlbasebrute := true;
   for i := 1 to 3 do
       begin
       if Basebrutespec.CellValues[0,i] <> '' then
          begin
          for j := i+1 to 4 do

              begin
              if  Basebrutespec.CellValues[0,i] = Basebrutespec.CellValues[0,j] then
                  begin
                  result := false;
                  ctlbasebrute := false;
                  end;
              end;
          end;
       end;
       if (ctlbasebrute=False) then
             begin
             Mess:= Mess+ '#13#10 - Saisie du même code dans les bases'+
                    ' brutes exceptionnelles';
             SetFocusControl ('GBASEBRUTESPEC');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.03.001';
             ErreurDADSU.Explication:= 'Plusieurs codes identiques'+
                                       ' sont saisis dans les bases brutes exceptionnelles';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             end;

    ctlbaseplaf := true;
   for i := 1 to 3 do
       begin
       if Baseplafspec.CellValues[0,i] <> '' then
          begin
          for j := i+1 to 4 do
              begin
              if  Baseplafspec.CellValues[0,i] = Baseplafspec.CellValues[0,j] then
                 begin
                 result := false;
                 ctlbaseplaf := false;
                 end;
              end;
          end;
       end;
       if (ctlbaseplaf=False) then
             begin
             Mess:= Mess+ '#13#10 - Saisie du même code dans les bases'+
                    ' plafonnées exceptionnelles';
             SetFocusControl ('GBASEPLAFSPEC');
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.03.001';
             ErreurDADSU.Explication:= 'Plusieurs codes identiques'+
                                       ' sont saisis dans les bases plafonnées exceptionnelles';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur (ErreurDADSU);
             end;
   // fin pt93-6

   for i := 1 to 5 do
       begin
       if ((RegimeComp.CellValues[0,i] = 'I001') or
          (RegimeComp.CellValues[0,i] = 'I002')) then
          begin
          for j := 1 to 4 do
              begin
              if ((Somiso.CellValues[0,j] <> '') and
                 (Somiso.CellValues[0,j] <> '04')) then
                 begin
                 result:= FALSE;
                 Mess:= Mess+ '#13#10 - Code type de sommes isolées erroné';
                 SetFocusControl('GSOMISO');
                 ErreurDADSU.Num:= Ordre;
                 ErreurDADSU.Segment:= 'S41.G01.04.001';
                 ErreurDADSU.Explication:= 'Le type de somme isolée est'+
                                           ' incorrect';
                 ErreurDADSU.CtrlBloquant:= True;
                 EcrireErreur (ErreurDADSU);
                 break;
                 end;
              end;
          end;

//PT93-3
      if ((RegimeComp <> NIL) and ((Pos ('A', RegimeComp.CellValues[0,i])=1) or
         ((Pos ('C', RegimeComp.CellValues[0,i])=1) and
         (IsNumeric (Copy (RegimeComp.CellValues[0,i], 1, 3)))) or
         (Pos ('G', RegimeComp.CellValues[0,i])=1))) then
          begin
          for j := 1 to 4 do
              begin
              if ((Somiso.CellValues[0,j] <> '') and
                 (Somiso.CellValues[0,j] <> '01') and
                 (Somiso.CellValues[0,j] <> '03')) then
                 begin
                 result:= FALSE;
                 Mess:= Mess+ '#13#10 - Code type de sommes isolées erroné';
                 SetFocusControl('GSOMISO');
                 ErreurDADSU.Num:= Ordre;
                 ErreurDADSU.Segment:= 'S41.G01.04.001';
                 ErreurDADSU.Explication:= 'Le type de somme isolée est'+
                                           ' incorrect';
                 ErreurDADSU.CtrlBloquant:= True;
                 EcrireErreur (ErreurDADSU);
                 break;
                 end;
              end;
          end;
//FIN PT93-3
       end;

   for j := 1 to 4 do
       begin
       if (Somiso.CellValues[2,j] > Copy (DateToStr (FinExer), 7, 4)) then
          begin
          result:= FALSE;
          Mess:= Mess+ '#13#10 - Année de rattachement sommes isolées erronée';
          SetFocusControl('GSOMISO');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.04.002';
          ErreurDADSU.Explication:= 'L''année de rattachement est incorrecte';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur (ErreurDADSU);
          break;
          end;
       end;

   for i := 1 to 4 do
       begin
       if ((IsNumeric (Somiso.CellValues[3,i])) and
          (StrToFloat(Somiso.CellValues[3,i])>StrToFloat(GetControlText('EBRUTSS')))) then
          begin
          Somiso.CellValues[3,i]:= GetControlText('EBRUTSS');
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.04.003.001';
          ErreurDADSU.Explication:= 'Une somme isolée brute était supérieure à'+
                                    ' la base brute SS';
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur (ErreurDADSU);
          end;
       end;
//FIN PT73-1

{PT73-5
   for i := 1 to 10 do
       begin
       if ((Exoner.CellValues[0,i] <> '') and
          (Exoner.CellValues[2,i] <> '') and
          (isnumeric(Exoner.CellValues[2,i]))) then
          begin
          if ((GetControlText('CCONTRAT')<>'CAP') and
             (Exoner.CellValues[0,i]='02')) then
             begin
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S41.G01.06.001';
             ErreurDADSU.Explication:= 'Le salarié comporte une exonération'+
                                       ' "apprenti loi de 1987" mais son'+
                                       ' contrat n''est pas "apprenti +10"';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur(ErreurDADSU);
             end;
          end;
       end;
}

//PT73-1
   for i := 1 to 10 do
       begin
       if ((IsNumeric (Exoner.CellValues[2,i])) and
          (StrToFloat(Exoner.CellValues[2,i])>StrToFloat(GetControlText('EBRUTSS')))) then
          begin
{PT76
          Exoner.CellValues[2,i]:= GetControlText('EBRUTSS');
}
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.06.002.001';
          ErreurDADSU.Explication:= 'Une base brute soumise à exonération est'+
                                    ' supérieure à la base brute SS';
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur (ErreurDADSU);
          end;
       end;

   for i := 1 to 10 do
       begin
       if ((IsNumeric (Exoner.CellValues[3,i])) and
          (StrToFloat(Exoner.CellValues[3,i])>StrToFloat(GetControlText('EPLAFONDSS')))) then
          begin
{PT76
          Exoner.CellValues[3,i]:= GetControlText('EPLAFONDSS');
}
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S41.G01.06.003.001';
          ErreurDADSU.Explication:= 'Une base plafonnée soumise à exonération'+
                                    ' est supérieure à la base plafonnée SS';
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur (ErreurDADSU);
          end;
       end;

   if (Nbre = NbreTot) then
      begin
      DateFinBuf:= PremierJourSemaine (NumSemaine (FinAnnee (FinExer)),
                                       StrToInt (PGExercice))+4;
      if (PrudPresent.Checked = TRUE) then
         begin
         if (((StrToDate (GetControlText ('DATEDEBDADS'))<DateFinBuf) and
            (StrToDate (GetControlText ('DATEFINDADS'))<DateFinBuf)) or
            ((StrToDate (GetControlText ('DATEDEBDADS'))>DateFinBuf) and
            (StrToDate (GetControlText ('DATEFINDADS'))>DateFinBuf))) then
            begin
            ErreurDADSU.Num:= Ordre;
            ErreurDADSU.Segment:= 'S41.G02.00.008';
            ErreurDADSU.Explication:= 'Veuillez vérifier l''information'+
                                      ' "Salarié présent au dernier vendredi'+
                                      ' de l''année"';
            ErreurDADSU.CtrlBloquant:= False;
            EcrireErreur(ErreurDADSU);
            end;
         end
      else
         begin
         if ((StrToDate (GetControlText ('DATEDEBDADS'))<DateFinBuf) and
            (StrToDate (GetControlText ('DATEFINDADS'))>DateFinBuf)) then
            begin
            ErreurDADSU.Num:= Ordre;
            ErreurDADSU.Segment:= 'S41.G02.00.008';
            ErreurDADSU.Explication:= 'Veuillez vérifier l''information'+
                                      ' "Salarié présent au dernier vendredi'+
                                      ' de l''année"';
            ErreurDADSU.CtrlBloquant:= False;
            EcrireErreur(ErreurDADSU);
            end;
         end;

      if ((GetControlText('CPRUDHCOLL')='') and (PrudDP.Checked = False) and
         (PrudPresent.Checked = TRUE)) then
         begin
         SetControlText ('CPRUDHCOLL', '1');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G02.00.009';
         ErreurDADSU.Explication:= 'Le collège prud''homal a été forcé à'+
                                   ' "Salarié"';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;

      if ((GetControlText('CPRUDHCOLL')<>'') and ((PrudDP.Checked = True) or
         (PrudPresent.Checked = False))) then
         begin
         SetControlText ('CPRUDHCOLL', '');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G02.00.009';
         ErreurDADSU.Explication:= 'Le collège prud''homal a été forcé à'+
                                   ' "Aucun"';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;

//PT93-5
      StEtab:= 'SELECT ETB_PRUDH'+
               ' FROM ETABCOMPL WHERE'+
               ' ETB_ETABLISSEMENT="'+GetControlText ('CETABLISSEMENT')+'"';
      QEtab:= OpenSql(StEtab, True);
      if (not QEtab.EOF) then
         Buffer:= QEtab.FindField ('ETB_PRUDH').AsString;
      Ferme (QEtab);
//FIN PT93-5

      if (GetControlText ('CPRUDHSECT')='') then
         begin
{PT93-5
         SetControlText ('CPRUDHSECT', '4');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G02.00.010';
         ErreurDADSU.Explication:= 'La section prud''homale a été forcée à'+
                                   ' "Activités diverses"';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur (ErreurDADSU);
}
         if (Buffer='') then
            Buffer:= '4';
         SetControlText ('CPRUDHSECT', Buffer);
         Buffer:= RechDom ('PGSECTIONPRUD', Buffer, False);
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S41.G02.00.010';
         ErreurDADSU.Explication:= 'La section prud''homale a été forcée à'+
                                   ' '''+Buffer+'''';
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur (ErreurDADSU);
//FIN PT93-5
         end
//PT93-5
      else
         begin
{PT106
         if ('0'+GetControlText ('CPRUDHSECT')<>Buffer) then
}         
         if (('0'+GetControlText ('CPRUDHSECT')<>Buffer) and
            (GetControlText ('CPRUDHSECT')<>'5')) then
//FIN PT106
            begin
            result:= FALSE;
            Mess:= Mess+ '#13#10 - Section prud''homale du salarié incohérente';
            SetFocusControl ('CPRUDHSECT');
            ErreurDADSU.Num:= Ordre;
            ErreurDADSU.Segment:= 'S41.G02.00.010';
            ErreurDADSU.Explication:= 'La section prud''homale du salarié ne'+
                                      ' correspond pas à la section'+
                                      ' prud''homale de l''établissement';
            ErreurDADSU.CtrlBloquant:= True;
            EcrireErreur (ErreurDADSU);
            end;
         end;
//FIN PT93-5
      end;
//FIN PT73-1

//PT73-8
   if ((RegimeComp <> NIL) and ((RegimeComp.CellValues[0,1] = 'I001') or
      (RegimeComp.CellValues[0,1] = 'I002') or
      (RegimeComp.CellValues[0,2] = 'I001') or
      (RegimeComp.CellValues[0,2] = 'I002') or
      (RegimeComp.CellValues[0,3] = 'I001') or
      (RegimeComp.CellValues[0,3] = 'I002') or
      (RegimeComp.CellValues[0,4] = 'I001') or
      (RegimeComp.CellValues[0,4] = 'I002') or
      (RegimeComp.CellValues[0,5] = 'I001') or
      (RegimeComp.CellValues[0,5] = 'I002'))) then
      begin
      if (GetControlText ('CIRCANTECPERIODICITE') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Périodicité de paiement des salaires IRCANTEC absente';
         SetFocusControl('CIRCANTECPERIODICITE');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.001';
         ErreurDADSU.Explication:= 'La périodicité de paiement des salaires'+
                                   ' IRCANTEC est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;

{PT93-5
      if (GetControlText ('EIRCANTECNBPAIE') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Nombre de périodicités de paie IRCANTEC';
         SetFocusControl('EIRCANTECNBPAIE');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.002';
         ErreurDADSU.Explication:= 'Le nombre de périodicités de paie IRCANTEC'+
                                   ' est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
}

//PT81
{PT93-5
      if (((GetControlText ('EIRCANTECHEBDO') = '') or
         (GetControlText ('EIRCANTECHEBDO') = '0')) and
         ((GetControlText('CSTATUTPROF') = '63') or
         (GetControlText('CSTATUTPROF') = '65') or
         (GetControlText('CSTATUTPROF') = '66') or
         (GetControlText('CSTATUTPROF') = '68'))) then
}
      if (((GetControlText ('EIRCANTECHEBDO')='') or
         (GetControlText ('EIRCANTECHEBDO')='0')) and
         ((GetControlText ('CSTATUTPROF')='63') or
         (GetControlText ('CSTATUTPROF')='65'))) then
//FIN PT93-5
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Durée hebdomadaire de travail du poste absente';
         SetControlText ('EIRCANTECHEBDO', '0');
         SetFocusControl('EIRCANTECHEBDO');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.004';
         ErreurDADSU.Explication:= 'La durée hebdomadaire de travail du poste'+
                                   ' IRCANTEC est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;

      if ((GetControlText ('EIRCANTECHEBDO') <> '') and
         (GetControlText ('EIRCANTECHEBDO') <> '0') and
{PT83
         (StrToInt (GetControlText ('EIRCANTECHEBDO')) >= 28) and
}
{PT93-5
         (Arrondi (StrToFloat (GetControlText ('EIRCANTECHEBDO'))*100, 2) >= 2800) and
         ((GetControlText('CSTATUTPROF') = '63') or
         (GetControlText('CSTATUTPROF') = '65') or
         (GetControlText('CSTATUTPROF') = '66') or
         (GetControlText('CSTATUTPROF') = '68'))) then
}
         (Arrondi (StrToFloat (GetControlText ('EIRCANTECHEBDO'))*100, 2) >= 2800) and
         ((GetControlText ('CSTATUTPROF')='63') or
         (GetControlText ('CSTATUTPROF')='65'))) then
//FIN PT93-5
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Durée hebdomadaire de travail du poste incorrecte';
         SetFocusControl('EIRCANTECHEBDO');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.004';
         ErreurDADSU.Explication:= 'La durée hebdomadaire de travail du poste'+
                                   ' IRCANTEC doit être inférieure à 28 heures';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//FIN PT81

      if (GetControlText ('EIRCANTECBRUT') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Base brute IRCANTEC absente';
         SetFocusControl('EIRCANTECBRUT');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.007.001';
         ErreurDADSU.Explication:= 'La base brute IRCANTEC est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;

      if (GetControlText ('EIRCANTECPLAFOND') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Base plafonnée IRCANTEC absente';
         SetFocusControl('EIRCANTECPLAFOND');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.008.001';
         ErreurDADSU.Explication:= 'La base plafonnée IRCANTEC est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;

//PT81
{PT93-5
      if (((GetControlText ('EIRCANTECMEDECINS') = '') or
         (GetControlText ('EIRCANTECMEDECINS') = '0')) and
         ((GetControlText('CSTATUTPROF') = '52') or
         (GetControlText('CSTATUTPROF') = '53') or
         (GetControlText('CSTATUTPROF') = '54') or
         (GetControlText('CSTATUTPROF') = '56') or
         (GetControlText('CSTATUTPROF') = '57') or
         (GetControlText('CSTATUTPROF') = '58') or
         (GetControlText('CSTATUTPROF') = '59'))) then
}
      if ((GetControlText ('EIRCANTECMEDECINS')='') and
         ((GetControlText ('CSTATUTPROF')='52') or
         (GetControlText ('CSTATUTPROF')='53') or
         (GetControlText ('CSTATUTPROF')='54') or
         (GetControlText ('CSTATUTPROF')='56') or
         (GetControlText ('CSTATUTPROF')='57') or
         (GetControlText ('CSTATUTPROF')='58') or
         (GetControlText ('CSTATUTPROF')='59')) and
         (GetControlText ('DECALAGE')<>'05')) then
//FIN PT93-5
         begin
         result:= FALSE;
         Mess:= Mess+ '#13#10 - Rémunération totale médecins absente';
         SetControlText ('EIRCANTECMEDECINS', '0');
         SetFocusControl('EIRCANTECMEDECINS');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S42.G01.00.009.001';
         ErreurDADSU.Explication:= 'La rémunération totale des médecins'+
                                   ' IRCANTEC est obligatoire';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur (ErreurDADSU);
         end;
//FIN PT81
      end;
//FIN PT73-8

   for i := 1 to 4 do
       begin
       if ((Prevoyance <> nil) and (Prevoyance.CellValues[0,i] <> '') and
          (Prevoyance.CellValues[0,i] <> RechDom('PGDADSPREV', '90', False))) then
          begin
//PT104-1
          if (not(IsValidDate (Prevoyance.CellValues[1,i]))) and
             (   (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '01', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '02', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '03', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '04', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '50', False))
             ) then
//PT104-1          if (not(IsValidDate (Prevoyance.CellValues[1,i]))) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Début de période de couverture absent';
//PT73-1
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S45.G01.01.002';
             ErreurDADSU.Explication:= 'La date de début de période de'+
                                       ' couverture n''est pas renseignée';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur(ErreurDADSU);
//FIN PT73-1
             end;

//PT104-1
          if (not(IsValidDate (Prevoyance.CellValues[2,i]))) and
             (   (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '02', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '03', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '04', False))
              or (Prevoyance.CellValues[0,i] = RechDom('PGDADSPREV', '50', False))
             ) then
//PT104-1          if (not(IsValidDate(Prevoyance.CellValues[2,i]))) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Fin de période de couverture absente';
//PT73-1
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S45.G01.01.003';
             ErreurDADSU.Explication:= 'La date de fin de période de'+
                                       ' couverture n''est pas renseignée';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur(ErreurDADSU);
//FIN PT73-1
             end;

//PT73-1
          if ((result = True) and (Prevoyance.CellValues[1,i] <> '') and (Prevoyance.CellValues[2,i] <> '') and  //PT104-1
             (StrToDate (Prevoyance.CellValues[1,i])> StrToDate (Prevoyance.CellValues[2,i]))) then
             begin
             result:= FALSE;
             Mess:= Mess+'#13#10 - Début de période de couverture';
             ErreurDADSU.Num:= Ordre;
             ErreurDADSU.Segment:= 'S45.G01.01.002';
             ErreurDADSU.Explication:= 'La date de début de période de'+
                                       ' couverture est posterieure à la date'+
                                       ' de fin de période de couverture';
             ErreurDADSU.CtrlBloquant:= True;
             EcrireErreur(ErreurDADSU);
             end;
//FIN PT73-1
          end;

//PT93-5
       if ((Prevoyance.CellValues[1,i]<>'') and
          (Prevoyance.CellValues[2,i]<>'') and
          ((StrToDate (Prevoyance.CellValues[1,i])<>StrToDate (GetControlText ('DATEDEBDADS'))) or
          (StrToDate (Prevoyance.CellValues[2,i])<>StrToDate (GetControlText ('DATEFINDADS')))) and
          ((Prevoyance.CellValues[6,i]='') or
          (Prevoyance.CellValues[7,i]=''))) then
          begin
          result:= FALSE;
          Mess:= Mess+'#13#10 - Base de prévoyance non renseignée';
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S45.G01.01.007.001';
          ErreurDADSU.Explication:= 'La base de prévoyance n''est pas renseignée';
          ErreurDADSU.CtrlBloquant:= True;
          EcrireErreur(ErreurDADSU);
          end;
//FIN PT93-5

//PT73-5
{PT93-5
       if (Prevoyance.CellValues[9,i]>'91') then
}
       if (Prevoyance.CellValues[10,i]>'91') then
//FIN PT93-5
          begin
{PT93-5
          Prevoyance.CellValues[9,i]:= '90';
}
          Prevoyance.CellValues[10,i]:= '90';
//FIN PT93-5
          ErreurDADSU.Num:= Ordre;
          ErreurDADSU.Segment:= 'S45.G01.01.010';
          ErreurDADSU.Explication:= 'Le nombre d''enfants couverts a été forcé'+
                                    ' à 90';
          ErreurDADSU.CtrlBloquant:= False;
          EcrireErreur(ErreurDADSU);
          end;
//FIN PT73-5
       end;

{PT73-1 + PT73-7
   if ((TypeD <> '001') and (GetControlText('CBTPSITFAM') = '')) then
      begin
      result:= FALSE;
      Mess:= Mess+ '#13#10 - Situation familiale BTP';
      SetFocusControl('CBTPSITFAM');
      end;
}
   if ((TypeD<>'001') and (TypeD<>'201')) then
      begin
      if (GetControlText('EBTPADHES') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Numéro d''adhésion BTP absent';
         SetFocusControl('EBTPADHES');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.001';
         ErreurDADSU.Explication:= 'Le numéro d''adhésion BTP n''est pas'+
                                   ' renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;

      if (GetControlText('EBTPANCIENENT') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Date d''ancienneté BTP absente';
         SetFocusControl('EBTPANCIENENT');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.004';
         ErreurDADSU.Explication:= 'La date d''ancienneté BTP n''est pas'+
                                   ' renseignée';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;

//PT75
      DoubleBuf:= Arrondi (StrToFloat (GetControlText ('EBTPANCIENENT')), 0);
      if (GetControlText ('EBTPANCIENENT') <> FloatToStr (DoubleBuf)) then
         begin
         SetControlText ('EBTPANCIENENT', FloatToStr (DoubleBuf));
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.004';
         ErreurDADSU.Explication:= 'La date d''ancienneté BTP a été arrondie à'+
                                   ' '+FloatToStr (DoubleBuf);
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;

      DoubleBuf:= Arrondi (StrToFloat (GetControlText ('EBTPANCIENPROF')), 0);
      if (GetControlText('EBTPANCIENPROF') <> FloatToStr (DoubleBuf)) then
         begin
         SetControlText ('EBTPANCIENPROF', FloatToStr (DoubleBuf));
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.005';
         ErreurDADSU.Explication:= 'La date d''ancienneté dans la profession'+
                                   ' BTP a été arrondie à '+
                                   FloatToStr (DoubleBuf);
         ErreurDADSU.CtrlBloquant:= False;
         EcrireErreur(ErreurDADSU);
         end;
//FIN PT75

      if (((GetControlText ('CBTPUNITSALMOY') = '01') or
         (GetControlText ('CBTPUNITSALMOY') = '02')) and
         (GetControlText ('EBTPSALMOY')= '')) then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Salaire moyen absent';
         SetFocusControl('EBTPSALMOY');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.008';
         ErreurDADSU.Explication:= 'Le salaire moyen BTP n''est pas renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;

      if (GetControlText('CBTPSITFAM') = '') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Situation familiale BTP absente';
         SetFocusControl('CBTPSITFAM');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.012';
         ErreurDADSU.Explication:= 'La situation familiale BTP n''est pas'+
                                   ' renseignée';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;

      if (GetControlText('EBTPENFANT') > '90') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Nombre d''enfants BTP';
         SetFocusControl('EBTPENFANT');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.013';
         ErreurDADSU.Explication:= 'Le nombre d''enfants est supérieur à "90"';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;

      if (GetControlText('ECLASSIFBTP')='') then
         begin
         result:= FALSE;
         Mess:= Mess+'#13#10 - Code classification BTP absent';
         SetFocusControl('ECLASSIFBTP');
         ErreurDADSU.Num:= Ordre;
         ErreurDADSU.Segment:= 'S66.G01.00.026';
         ErreurDADSU.Explication:= 'Le code classification BTP n''est pas'+
                                   ' renseigné';
         ErreurDADSU.CtrlBloquant:= True;
         EcrireErreur(ErreurDADSU);
         end;
      end;
//FIN PT73-1 + PT73-7
   end;

if result = FALSE then
   begin
   PGIBox ('Anomalies :'+Mess,'Informations erronées DADS-U');
   end;

QPer.First;
While not QPer.Eof do
      begin
      if (GetControlText('ORDREDADS')=IntToStr(QPer.FindField ('PDE_ORDRE').Asinteger)) then
         QPer.Next
      else
         begin
         DateDebBuf:= QPer.FindField('PDE_DATEDEBUT').AsDateTime;
         DateFinBuf:= QPer.FindField('PDE_DATEFIN').AsDateTime;
         if ((StrToDate(GetControlText('DATEDEBDADS'))> DateDebBuf) and
            (StrToDate(GetControlText('DATEDEBDADS'))< DateFinBuf)) then
            begin
            Mess:= 'La date de début '+GetControlText('DATEDEBDADS')+
                   ' est comprise entre le#13#10'+DateToStr(DateDebBuf)+' et'+
                   ' le '+DateToStr(DateFinBuf)+' qui sont les dates de début'+
                   ' et de fin d''une autre période';
            PGIBox (Mess,'Contrôle DADS-U');
            result:= FALSE;
            end;

         if ((StrToDate(GetControlText('DATEFINDADS'))>DateDebBuf) and
            (StrToDate(GetControlText('DATEFINDADS'))<DateFinBuf)) then
            begin
            Mess:= 'La date de fin '+GetControlText('DATEFINDADS')+' est'+
                   ' comprise entre le#13#10'+DateToStr(DateDebBuf)+' et le '+
                   DateToStr(DateFinBuf)+' qui sont les dates de début et de'+
                   ' fin d''une autre période';
            PGIBox (Mess,'Contrôle DADS-U');
            result := FALSE;
            end;
         QPer.Next;
         end;
      end;
QPer.First;
While ((not QPer.Eof) and
      (IntToStr(QPer.FindField('PDE_ORDRE').AsInteger)<>GetControlText('ORDREDADS'))) do
      QPer.Next;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/09/2001
Modifié le ... :   /  /
Description .. : Mise à jour des tables
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
function TOF_PG_DADSPERIODES.UpdateTable(): boolean;
var
Ordre, Rep : integer;
begin
result := FALSE;
Ordre := StrToInt(GetControlText('ORDREDADS'));

Rep:= PGIAsk ('Voulez vous sauvegarder votre saisie ?',
              'Saisie complémentaire DADS-U') ;
if Rep=mrNo then
   exit
else
   result := TRUE;

try
   begintrans;
//PT73-1
   DeleteErreur (Sal, 'S', 0);
   DeleteErreur (Sal, 'S', Ordre);
//FIN PT73-1
   ControleOK := ControleConform;
   if ControleOK = TRUE then
      begin
{PT82
      DeletePeriode(Sal, TypeD, Ordre);
      DeleteDetail(Sal, TypeD, 0);
      DeleteDetail(Sal, TypeD, Ordre);
}
      DeletePeriode(Sal, Ordre);
      DeleteDetail(Sal, 0);
      DeleteDetail(Sal, Ordre);
//FIN PT82
      
      ChargeTOBDADS;
      DeleteErreur ('', 'SKO');	//PT73-1
      SauveZones;
      EcrireErreurKO;	//PT73-1
      LibereTOBDADS;
      MAJQuery;
      TFVierge(Ecran).Binsert.Enabled := True;
      TFVierge(Ecran).BDelete.Enabled := True;
      if Calculer <> NIL then
         Calculer.Enabled := True;
      end;
   CommitTrans;
Except
   result := FALSE;
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base', 'Mise à jour DADS-U');
   END;
end;

{$IFNDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrem"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalPremClick(Sender: TObject);
begin
BougeSal(nbFirst) ;
SalPrem.Enabled := FALSE;
SalPrec.Enabled := FALSE;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalPrec"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalPrecClick(Sender: TObject);
begin
BougeSal(nbPrior) ;
if QMul.BOF then
   begin
   SalPrem.Enabled := FALSE;
   SalPrec.Enabled := FALSE;
   end;
SalSuiv.Enabled := TRUE;
SalDern.Enabled := TRUE;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalSuiv"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalSuivClick(Sender: TObject);
begin
BougeSal(nbNext) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
if QMul.EOF then
   begin
   SalSuiv.Enabled := FALSE;
   SalDern.Enabled := FALSE;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "SalDern"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalDernClick(Sender: TObject);
begin
BougeSal(nbLast) ;
SalPrem.Enabled := TRUE;
SalPrec.Enabled := TRUE;
SalSuiv.Enabled := FALSE;
SalDern.Enabled := FALSE;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Déplacement au niveau du salarié
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
Function TOF_PG_DADSPERIODES.BougeSal(Button: TNavigateBtn) : boolean ;
BEGIN
UpdateTable;
result:=TRUE ;

if Button=nbDelete then
   begin
   if QMul.EOF = FALSE then
      begin
      QMul.Next;
      if QMul.EOF = TRUE then
         begin
         QMul.prior ;
         if QMul.BOF then
            Close;
         end
      end
   else
      begin
      if QMul.BOF = FALSE then
         begin
         QMul.prior;
         if QMul.BOF = TRUE then
            Close;
         end;
      end;
   end;

if QMul <> NIL then
   begin
   Case Button of
        nblast : QMul.Last;
        nbfirst : QMul.First;
        nbnext : QMul.Next;
        nbprior : QMul.prior;
        end;
   end;
GereQuerySal;
END ;


procedure TOF_PG_DADSPERIODES.GereQuerySal();
begin
if QMul = NIL then
   exit;
Sal := QMul.FindField('PSA_SALARIE').AsString;
if T_Periode <> NIL then
   FreeAndNil (T_Periode);

SalChange := True;
ChargeZones;
AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "PerPrec"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.PerPrecClick(Sender: TObject);
begin
BougePer(nbPrior) ;
Nbre := IntToStr(StrToInt(Nbre)-1);
if QPer.BOF then
   begin
   Nbre := IntToStr(StrToInt(Nbre)+1);
   PerPrec.Enabled := FALSE;
   end;

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "PerSuiv"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.PerSuivClick(Sender: TObject);
begin
BougePer(nbNext) ;
Nbre := IntToStr(StrToInt(Nbre)+1);
if QPer.EOF then
   begin
   Nbre := IntToStr(StrToInt(Nbre)-1);
   PerSuiv.Enabled := FALSE;
   end;

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : MAJ de la query des périodes
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MAJQuery();
var
StPer : string;
QCount : TQuery;
begin
Ferme(QPer);
StPer:= 'SELECT COUNT(PDE_SALARIE) AS NBRE'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE = "'+Sal+'" AND'+
        ' PDE_TYPE="'+TypeD+'" AND'+
        ' PDE_ORDRE > 0 AND'+
        ' PDE_EXERCICEDADS = "'+PGExercice+'"';

QCount:=OpenSql(StPer, True);
if (not QCount.EOF) then
   NbreTot := IntToStr(QCount.FindField('NBRE').AsInteger)
else
   NbreTot := '0';
Ferme(QCount);

StPer:= 'SELECT *'+
        ' FROM DADSPERIODES WHERE'+
        ' PDE_SALARIE = "'+Sal+'" AND'+
        ' PDE_TYPE="'+TypeD+'" AND'+
        ' PDE_ORDRE > 0 AND'+
        ' PDE_EXERCICEDADS = "'+PGExercice+'"'+
        ' ORDER BY PDE_DATEDEBUT';

QPer:=OpenSql(StPer, True);
QPer.First;
Nbre := IntToStr(1);
while ((not QPer.Eof) and
      (IntToStr(QPer.FindField('PDE_ORDRE').AsInteger) <> GetControlText('ORDREDADS'))) do    //DB2
      begin
      QPer.Next;
      Nbre := IntToStr(StrToInt(Nbre)+1);
      end;

if (QPer.EOF) then
   begin
   Nbre := IntToStr(1);
   QPer.First;
   if (not QPer.Eof) then
      SetControlText ('ORDREDADS',
                      IntToStr(QPer.FindField('PDE_ORDRE').AsInteger))
   else
      SetControlText ('ORDREDADS', '0');
   ChargeZones;
   end;

AfficheCaption;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Déplacement au niveau de la période
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
Function TOF_PG_DADSPERIODES.BougePer(Button: TNavigateBtn) : boolean ;
BEGIN
UpdateTable;
result:=TRUE ;

if QPer <> NIL then
   begin
   Case Button of
        nbfirst : QPer.First;
        nbprior : QPer.prior;
        nbnext : QPer.Next;
        nblast : QPer.Last;
        end;
   end;
GereQueryPer;
END ;


procedure TOF_PG_DADSPERIODES.GereQueryPer();
begin
if QPer = NIL then
   exit;
if T_Periode <> NIL then
   FreeAndNil (T_Periode);

ChargeZones;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2001
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Calculer"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.CalculerClick(Sender: TObject);
var
TDate, TDateFille, TDateDeb, TDateFin : TOB;
BufDest : String;
PeriodeFaite, Rep : integer;
begin
Rep := PGIAsk ('Les éventuelles modifications saisies dans la période#13#10'+
               'd''activité seront annulées si vous confirmez le calcul#13#10'+
               'de cette période.#13#10'+
               'Voulez-vous continuer ?', 'Saisie complémentaire DADS-U') ;
if Rep=mrNo then
   exit;

Trace := TStringList.Create;
ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);

{$IFDEF EAGLCLIENT}
{PT73-1
NomFic := VH_Paie.PgCheminEagl+'\'+BufDest+'_DADSU_PGI.log';
}
{$ELSE}
{PT73-1
NomFic := V_PGI.DatPath+'\'+BufDest+'_DADSU_PGI.log';
}
{$ENDIF}

{PT73-1
if FileExists(NomFic) then
   begin
   Maintenant := Now;
   DateCalcul := Now;
   FileAttrs := 0;
   FileAttrs := FileAttrs + faAnyFile;
   if FindFirst(NomFic, FileAttrs, sr) = 0 then
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         DateCalcul := FileDateToDateTime(sr.Time);
      sysutils.FindClose(sr);
      end;

   if (PlusMois(Maintenant, -6)>DateCalcul) then
      DeleteFile(PChar(NomFic));
   end;

AssignFile(FRapport, NomFic);
if FileExists(NomFic) then
   begin
   Append(FRapport);
   Writeln(FRapport, '');
   end
else
   begin
   ReWrite(FRapport);
   Writeln(FRapport, 'Attention, Le dernier calcul se trouve en fin du fichier');
   end;

Writeln(FRapport, '_____________________________________');
Writeln(FRapport, 'Début de calcul : '+DateTimeToStr(Now));
Writeln(FRapport, '');
Writeln(FRapport, 'Paramètres : Fraction '+GetControlText('LFRACTION_'));
Writeln(FRapport, '             Période du '+DateToStr(DebExer)+' au '+
        DateToStr(FinExer));
Writeln(FRapport, '             Salarié '+Sal);
}

if StrToDate(GetControlText('DATEDEBDADS'))<=StrToDate(GetControlText('DATEFINDADS')) then
   begin
   try
      begintrans;

      ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);
      if ControlSiret(BufDest)=False then
         PGIBox ('Le SIRET de la société n''est pas valide.#13#10'+
                 'Vous devez le vérifier en y accédant par le module#13#10'+
                 'Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées.#13#10'+
                 'Si vous travaillez en environnement multi-dossiers,#13#10'+
                 'vous pouvez y accéder par le Bureau PGI/Annuaire',
                 'Saisie complémentaire DADS-U');

{PT82
      DuplicPrudh(Sal, TypeD, 'C', StrToInt(GetControlText('ORDREDADS')));
      DeletePeriode(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, TypeD, StrToInt(GetControlText('ORDREDADS')));
}
      DuplicPrudh(Sal, 'C', StrToInt(GetControlText('ORDREDADS')));
      DeletePeriode(Sal, StrToInt(GetControlText('ORDREDADS')));
      DeleteDetail(Sal, StrToInt(GetControlText('ORDREDADS')));
//FIN PT82

      ChargeTOBCommun;
      ChargeTOBSal(Sal);
      DeleteErreur ('', 'SKO');	//PT73-1

//Récup des dates de ruptures salariés
      TDate := TOB.Create('Les Dates', NIL, -1);
      TDateFille := TOB.Create('',TDate,-1);
      TDateFille.AddChampSup('DATE', FALSE);
      TDateFille.AddChampSup('NIVEAU', FALSE);
      TDateFille.AddChampSup('MOTIFENTREE', FALSE);
      TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
      TDateFille.AddChampSup('AJUDEB', FALSE);
      TDateFille.AddChampSup('AJUFIN', FALSE);
      TDateFille.PutValue('DATE', StrToDate(GetControlText('DATEDEBDADS')));
      TDateFille.PutValue('NIVEAU', 'C');
{PT73-3
      TDateFille.PutValue('MOTIFENTREE', GetControlText('CDEBMOTIF1'));
}
      TDateFille.PutValue('MOTIFENTREE', GetControlText('EDEBMOTIF1'));
      TDateFille.PutValue('AJUDEB', 0);
      TDateFille.PutValue('AJUFIN', 0);

      TDateFille := TOB.Create('',TDate,-1);
      TDateFille.AddChampSup('DATE', FALSE);
      TDateFille.AddChampSup('NIVEAU', FALSE);
      TDateFille.AddChampSup('MOTIFENTREE', FALSE);
      TDateFille.AddChampSup('MOTIFSORTIE', FALSE);
      TDateFille.AddChampSup('AJUDEB', FALSE);
      TDateFille.AddChampSup('AJUFIN', FALSE);
      TDateFille.PutValue('DATE', StrToDate(GetControlText('DATEFINDADS')));
      TDateFille.PutValue('NIVEAU', 'C');
{PT73-3
      TDateFille.PutValue('MOTIFSORTIE', GetControlText('CFINMOTIF1'));
}
      TDateFille.PutValue('MOTIFSORTIE', GetControlText('EFINMOTIF1'));
      TDateFille.PutValue('AJUDEB', 0);
      TDateFille.PutValue('AJUFIN', 0);

//Boucle sur les périodes du salarié
      TDateDeb := TDate.FindFirst([''],[''],FALSE);
      TDateFin := TDate.FindNext([''],[''],FALSE);

      Periode:= StrToInt(GetControlText('ORDREDADS'))-1;
      DeleteErreur (Sal, 'S4', Periode+1);	//PT73-1
{PT82
      PeriodeFaite:= RemplitPeriode (Sal, TypeD, TDateDeb, TDateFin);
}
      PeriodeFaite:= RemplitPeriode (Sal, TDateDeb, TDateFin);

{PT73-7
      if (TypeD <> '001') then
         CalculBTP (Sal, TypeD);
      if ((PeriodeFaite<>0) and (Nbre=NbreTot)) then
         RemplitPrudh (Sal, TypeD, PeriodeFaite);
}
      if ((TypeD<>'001') and (TypeD<>'201')) then
{PT82
         CalculBTP (Sal, TypeD)
}
         CalculBTP (Sal)
      else
         begin
         if ((PeriodeFaite<>0) and (Nbre=NbreTot)) then
{PT82
            RemplitPrudh (Sal, TypeD, PeriodeFaite);
}
{PT108
}
            begin
            if (GetControlText ('EDEBMOTIF1')='095') then
               StDateDebut:= GetControlText('DATEFINDADS')
            else
               StDateDebut:= GetControlText('DATEDEBDADS');
            RemplitPrudh (Sal, PeriodeFaite);
            end;
//FIN PT108
         end;
//FIN PT73-7
      EcrireErreurKO;	//PT73-1
      LibereTOB;
      LibereTOBCommun;
      ExecuteSQL('UPDATE SALARIES SET'+
                 ' PSA_DADSDATE = "'+UsDateTime(Date)+
                 '" WHERE PSA_SALARIE = "'+Sal+'"');
      FreeAndNil (TDate);
      PGIBox('Calcul terminé','Calcul de la DADS-U');
{PT73-1
      Writeln(FRapport, 'Calcul terminé : '+DateTimeToStr(Now));
      CloseFile(FRapport);
      ShellExecute ( 0, PCHAR('open'),PChar('WordPad'),
                   PChar(NomFic),Nil,SW_RESTORE);
}
      CommitTrans;
   Except
      Rollback;
      PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
             'Calcul de la DADS-U');
{PT73-1
      Writeln(FRapport, 'Calcul annulé : '+DateTimeToStr(Now));
      CloseFile(FRapport);
}
      END;
   ChargeZones;
   end
else
   begin
   PGIBox ('La date de début est supérieure à la date de fin',
          'Calcul de la DADS-U');
{PT73-1
   Writeln(FRapport, 'Calcul annulé : '+DateTimeToStr(Now));
   CloseFile(FRapport);
}
   end;
{$IFNDEF DADSUSEULE}
CreeJnalEvt('001', '040', 'OK', NIL, NIL, Trace);
{$ENDIF}
if Trace <> nil then
   FreeAndNil (Trace);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/08/2002
Modifié le ... :   /  /
Description .. : Click sur le bouton SALMAJ
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalarieMAJClick (Sender: TObject);
var
BufDest, NomFic, StExiste : String;
FileAttrs, Rep : integer;
DateCalcul, Maintenant : TDateTime;
sr : TSearchRec;
begin
Rep := PGIAsk ('Attention, ce traitement va remettre à jour les données#13#10'+
              'issues de la fiche salarié. #13#10'+
              'Voulez-vous continuer ?', 'Saisie complémentaire DADS-U') ;
if Rep=mrNo then
   exit;

ForceNumerique(GetParamSoc('SO_SIRET'), BufDest);
{$IFDEF EAGLCLIENT}
NomFic := VH_Paie.PgCheminEagl+'\'+BufDest+'_DADSU_PGI.log';
{$ELSE}
NomFic := V_PGI.DatPath+'\'+BufDest+'_DADSU_PGI.log';
{$ENDIF}

if FileExists(NomFic) then
   begin
   Maintenant := Now;
   DateCalcul := Now;
   FileAttrs := 0;
   FileAttrs := FileAttrs + faAnyFile;
   if FindFirst(NomFic, FileAttrs, sr) = 0 then
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         DateCalcul := FileDateToDateTime(sr.Time);
      sysutils.FindClose(sr);
      end;

   if (PlusMois(Maintenant, -6)>DateCalcul) then
      DeleteFile(PChar(NomFic));
   end;

{PT73-1
AssignFile(FRapport, NomFic);
if FileExists(NomFic) then
   begin
   Append(FRapport);
   Writeln(FRapport, '');
   end
else
   begin
   ReWrite(FRapport);
   Writeln(FRapport, 'Attention, Le dernier calcul se trouve en fin du fichier');
   end;

Writeln(FRapport, '_____________________________________');
Writeln(FRapport, 'Début de calcul des éléments de la fiche salarié : '+DateTimeToStr(Now));
Writeln(FRapport, '');
Writeln(FRapport, 'Paramètres : Fraction '+GetControlText('LFRACTION_'));
Writeln(FRapport, '             Période du '+DateToStr(DebExer)+' au '+
        DateToStr(FinExer));
}

try
   begintrans;
   StExiste:= 'SELECT PDS_SALARIE'+
              ' FROM DADSDETAIL WHERE'+
              ' PDS_SALARIE="'+Sal+'" AND'+
              ' PDS_TYPE="'+TypeD+'" AND'+
              ' PDS_ORDRE=0 AND'+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"';
   if (ExisteSQL(StExiste) = TRUE) then
      begin
{PT82
      DeleteDetail (Sal, TypeD, 0);
}
      DeleteDetail (Sal, 0);
      ChTOBSal (Sal);
      ChargeTOBDADS;
//PT73-1
      DeleteErreur (Sal, 'S3', 0);
      DeleteErreur ('', 'SKO');
//FIN PT73-1
{PT82
      CalculElemSal (Sal, TypeD);
}
      CalculElemSal (Sal);
      EcrireErreurKO;	//PT73-1
      LibTOB;
      LibereTOBDADS;
      end
   else
      begin
      PGIBox('Traitement impossible : Le salarié '+Sal+' n''a jamais été calculé',
             'Mise à jour données salarié');
{PT73-1
      Writeln(FRapport, 'Mise à jour des données du salarié impossible : Le salarié '+Sal+' n''a jamais été calculé');
}
      end;
   PGIBox('Traitement terminé','Mise à jour données salarié');
{PT73-1
   Writeln(FRapport, 'Mise à jour des données du salarié terminée : '+DateTimeToStr(Now));
   CloseFile(FRapport);
   ShellExecute(0, PCHAR('open'),PChar('WordPad'), PChar(NomFic),Nil, SW_RESTORE);
}
   CommitTrans;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
           'Mise à jour données salarié');
{PT73-1
   Writeln(FRapport, 'Mise à jour des données du salarié annulée : '+DateTimeToStr(Now));
   CloseFile(FRapport);
}
   END;

ChargeZones;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... : 19/12/2001  JL modification LanceEtat et requête
Description .. : Click sur le bouton Imprimer
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.Impression(Sender: TObject);
var
Pages : TPageControl;
Rech, StPages : String;
begin
Rech:='SELECT PDE_SALARIE, PDE_DATEDEBUT, PDE_DATEFIN, PDE_TYPE, PDE_ORDRE,'+
      ' PDS_ORDRE, PDS_SALARIE'+
      ' FROM DADSPERIODES'+
      ' LEFT JOIN DADSDETAIL ON'+
      ' PDE_SALARIE=PDS_SALARIE AND'+
      ' PDE_TYPE=PDS_TYPE AND'+
      ' PDE_ORDRE=PDS_ORDRE AND'+
      ' PDE_DATEDEBUT=PDS_DATEDEBUT AND'+
      ' PDE_DATEFIN=PDS_DATEFIN AND'+
      ' PDE_EXERCICEDADS=PDS_EXERCICEDADS'+
      ' WHERE PDE_SALARIE="'+Sal+'" AND'+
      ' PDE_TYPE = "'+TypeD+'" AND'+
      ' PDE_ORDRE = '+GetControlText('ORDREDADS')+' AND'+
      ' PDE_EXERCICEDADS = "'+PGExercice+'"';

Pages := TPageControl(GetControl('PAGES'));
if (GetControlText('EENTREE')='  /  /    ') then
   SetControlText('EENTREE', DateToStr(IDate1900));
StPages := AglGetCriteres (Pages, FALSE);
LanceEtat('E','PDA','PDF',True,False,False,Pages,Rech,'',False,0,StPages);
if (GetControlText('EENTREE')=DateToStr(IDate1900)) then
   SetControlText('EENTREE', '');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Click sur le bouton valider
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.Validation(Sender: TObject);
var
Pages : TPageControl;
SavRow : Integer;//PT104
begin
//DEBUT PT104 Pour régler des problèmes d'affichage des champs avec tablette dans la grille
// il faut positionner le focus sur une case sans tablette. On choisit ici
// la dernière colonne (choix arbitraire). Il faut aussi forcer la validation de
// la cellule au cas ou on ne soit pas sorti de la cellule après modification.
Pages:= TPageControl(GetControl('PAGES'));
if (Prevoyance <> NIL) and (Pages.ActivePage.Name = 'TRETPREV') then
  begin
  Prevoyance.SetFocus;
  Prevoyance.UpdateCell;
  Prevoyance.Col := Prevoyance.ColCount-1;
  SavRow := Prevoyance.Row;
  Prevoyance.Row := 1;
  Prevoyance.Row := Prevoyance.RowCount-1;
  Prevoyance.Row := SavRow;
  end;
//FIN PT104

UpdateTable;
if (ControleOK = TRUE) then
   begin
   if (State = 'CREATION') then
      begin
      SalPrem.Enabled := TRUE;
      SalPrec.Enabled := TRUE;
      SalSuiv.Enabled := TRUE;
      SalDern.Enabled := TRUE;
      end;
   State := 'MODIFICATION';
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2001
Modifié le ... :   /  /
Description .. : Click sur le bouton Salarié
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.SalarieClick(Sender: TObject);
begin
AGLLanceFiche ('PAY', 'SALARIE', '',Sal, '');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/10/2001
Modifié le ... :   /  /
Description .. : Click sur l'ellipsis d'une grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.GrilleElipsisClick(Sender: TObject);
var
sWhere : string;
QRechInstit : TQuery;
begin
if ((Sender = RegimeComp) and (RegimeComp.Col = 0)) then
   begin
{PT73-7
   if (TypeD<>'001') then
}
   if ((TypeD<>'001') and (TypeD<>'201')) then
//FIN PT73-7
      sWhere:= '##PIP_PREDEFINI## PIP_INSTITUTION LIKE "B%"'
   else
      sWhere:= '##PIP_PREDEFINI## PIP_INSTITUTION LIKE "A%" OR'+
               ' PIP_INSTITUTION LIKE "C%" OR PIP_INSTITUTION LIKE "F%" OR'+
               ' PIP_INSTITUTION LIKE "G%" OR PIP_INSTITUTION LIKE "I%"';
   LookUpList (Regimecomp,'Organismes','INSTITUTIONPAYE','PIP_INSTITUTION','PIP_ABREGE',swhere,'PIP_INSTITUTION',TRUE,-1);
   sWhere:= 'SELECT PIP_ABREGE, POG_NUMAFFILIATION'+
            ' FROM INSTITUTIONPAYE'+
            ' LEFT JOIN ORGANISMEPAIE ON'+
            ' PIP_INSTITUTION=POG_INSTITUTION WHERE'+
            ' ##PIP_PREDEFINI##'+
            ' PIP_INSTITUTION = "'+RegimeComp.CellValues[0,RegimeComp.Row]+'" AND'+
            ' POG_ETABLISSEMENT = "'+GetControlText('CETABLISSEMENT')+'"';
   QRechInstit := OpenSql(sWhere,TRUE);
   if (not QRechInstit.EOF) then
      begin
      RegimeComp.CellValues[1,RegimeComp.Row] := QRechInstit.Fields[0].asstring;
      RegimeComp.CellValues[2,RegimeComp.Row] := QRechInstit.Fields[1].asstring;
      end
   else
      begin
      RegimeComp.CellValues[1,RegimeComp.Row] := '';
      RegimeComp.CellValues[2,RegimeComp.Row] := '';
      end;
   Ferme(QRechInstit);
   end;

if ((Sender = Basebrutespec) and (Basebrutespec.Col = 0)) then
   begin
   sWhere:='##PBB_PREDEFINI## PBB_LIBELLE <> ""';
   LookUpList (Basebrutespec,'Types','BASEBRUTESPEC','PBB_CODE','PBB_LIBELLE',sWhere,'PBB_CODE',TRUE,-1);
   sWhere:= 'SELECT PBB_LIBELLE'+
            ' FROM BASEBRUTESPEC WHERE'+
            ' ##PBB_PREDEFINI## PBB_CODE = "'+Basebrutespec.CellValues[0,Basebrutespec.Row]+'"';
   QRechInstit := OpenSql(sWhere,TRUE);
   if (not QRechInstit.EOF) then
      Basebrutespec.CellValues[1,Basebrutespec.Row] := QRechInstit.Fields[0].asstring;
   Ferme(QRechInstit);
   end;

if ((Sender = Baseplafspec) and (Baseplafspec.Col = 0)) then
   begin
   sWhere:='##PBB_PREDEFINI## PBB_BASEPLAFOND <> ""';
   LookUpList (Baseplafspec,'Types','BASEBRUTESPEC','PBB_CODE','PBB_BASEPLAFOND',sWhere,'PBB_CODE',TRUE,-1);
   sWhere:= 'SELECT PBB_BASEPLAFOND'+
            ' FROM BASEBRUTESPEC WHERE'+
            ' ##PBB_PREDEFINI## PBB_CODE = "'+Baseplafspec.CellValues[0,Baseplafspec.Row]+'"';
   QRechInstit := OpenSql(sWhere,TRUE);
   if (not QRechInstit.EOF) then
      Baseplafspec.CellValues[1,Baseplafspec.Row] := QRechInstit.Fields[0].asstring;
   Ferme(QRechInstit);
   end;

if ((Sender = Somiso) and (Somiso.Col = 0)) then
   begin
   sWhere:='##PBB_PREDEFINI## PBB_SOMISO <> ""';
   LookUpList (Somiso,'Sommes versées pour','BASEBRUTESPEC','PBB_CODE','PBB_SOMISO',sWhere,'PBB_CODE',TRUE,-1);
   sWhere:= 'SELECT PBB_SOMISO'+
            ' FROM BASEBRUTESPEC WHERE'+
            ' ##PBB_PREDEFINI## PBB_CODE = "'+Somiso.CellValues[0,Somiso.Row]+'"';
   QRechInstit := OpenSql(sWhere,TRUE);
   if (not QRechInstit.EOF) then
      Somiso.CellValues[1,Somiso.Row] := QRechInstit.Fields[0].asstring;
   Ferme(QRechInstit);
   end;

if ((Sender = Exoner) and (Exoner.Col = 0)) then
   begin
   sWhere:='##PBB_PREDEFINI## PBB_LEXONERATION <> ""';
   LookUpList (Exoner,'Catégories d''exonération','BASEBRUTESPEC','PBB_CODE','PBB_LEXONERATION',sWhere,'PBB_CODE',TRUE,-1);
   sWhere:= 'SELECT PBB_LEXONERATION'+
            ' FROM BASEBRUTESPEC WHERE'+
            ' ##PBB_PREDEFINI## PBB_CODE = "'+Exoner.CellValues[0,Exoner.Row]+'"';
   QRechInstit := OpenSql(sWhere,TRUE);
   if (not QRechInstit.EOF) then
      Exoner.CellValues[1,Exoner.Row] := QRechInstit.Fields[0].asstring;
   Ferme(QRechInstit);
   end;

if Sender = Prevoyance then
   begin
   if ((Prevoyance.Col = 1) OR (Prevoyance.Col = 2) AND (Prevoyance.row <> 0)) then
      DateDoubleClick (Sender);

   if ((Prevoyance.Col = 3) AND (Prevoyance.Row <> 0)) then
      begin
      sWhere:= '##PIP_PREDEFINI## PIP_INSTITUTION < "A000"';
      LookUpList (Prevoyance,'Organismes','INSTITUTIONPAYE','PIP_INSTITUTION','PIP_LIBELLE',swhere,'PIP_INSTITUTION',TRUE,-1);
      end;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/06/2004
Modifié le ... :   /  /
Description .. : Click sur l'ellipsis métier BTP
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MetierBTPElipsisClick(Sender: TObject);
var
sWhere : string;
QRechMetier : TQuery;
begin
LookUpList (MetierBTP,'Metiers BTP','METIERBTP','PMB_METIERBTP',
           'PMB_CONVCOLL, PMB_SPECIALITE, PMB_STATUTBTP, PMB_METIERLIB',
           '##PMB_PREDEFINI##',
           'PMB_CONVCOLL,PMB_SPECIALITE,PMB_METIERBTP',TRUE,-1);
sWhere:= 'SELECT PMB_METIERLIB'+
         ' FROM METIERBTP WHERE'+
         ' ##PMB_PREDEFINI## PMB_METIERBTP = "'+GetControlText('EMETIERBTP')+'"';
QRechMetier := OpenSql(sWhere,TRUE);
if (not QRechMetier.EOF) then
   SetControlText('LMETIERBTP_', QRechMetier.Fields[0].asstring);
Ferme(QRechMetier);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/10/2001
Modifié le ... :   /  /
Description .. : Entée dans une cellule d'une grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.GrilleCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
if (Sender = RegimeComp) then
   begin
   if (Regimecomp.Col=0) then
{PT79-2
      Regimecomp.ElipsisButton:= TRUE
}
      begin
      Regimecomp.ElipsisButton:= TRUE;
      CodeRegime:= Regimecomp.CellValues[(Regimecomp.Col),(Regimecomp.Row)];
      end
//FIN PT79-2
   else
      Regimecomp.ElipsisButton:= FALSE;

   FocusRegime:= True;           //PT84
   end;

if Sender = Basebrutespec then
   begin
   if Basebrutespec.Col = 0 then
      Basebrutespec.ElipsisButton := TRUE
   else
      Basebrutespec.ElipsisButton := FALSE;
   end;

if Sender = Baseplafspec then
   begin
   if Baseplafspec.Col = 0 then
      Baseplafspec.ElipsisButton := TRUE
   else
      Baseplafspec.ElipsisButton := FALSE;
   end;

if Sender = Somiso then
   begin
   if Somiso.Col = 0 then
      Somiso.ElipsisButton := TRUE
   else
      Somiso.ElipsisButton := FALSE;
   end;

if Sender = Exoner then
   begin
   if Exoner.Col = 0 then
      Exoner.ElipsisButton := TRUE
   else
      Exoner.ElipsisButton := FALSE;
   end;

if Sender = Prevoyance then
   begin
   if (((Prevoyance.Col = 1) OR (Prevoyance.Col = 2)  OR
      (Prevoyance.Col = 3)) AND
      (Prevoyance.row <> 0)) then
      Prevoyance.ElipsisButton := TRUE
   else
      Prevoyance.ElipsisButton := FALSE;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/10/2001
Modifié le ... :   /  /
Description .. : Sortie d'une cellule d'une grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.GrilleCellExit (Sender: TObject;
                                              var ACol,ARow: Integer;
                                              var Cancel: Boolean);
var
ResultRech, SWhere : string;
QRechInstit : TQuery;
ExisteIRCANTEC : boolean;
TabSheet : TTabsheet;
begin
if Sender = RegimeComp then
   begin
   ExisteIRCANTEC:= False;	//PT73-8
{PT84
   if (ACol = 0) then
}
   if ((ACol = 0) and (FocusRegime=True)) then
      begin
      ResultRech:=RechDom('PGINSTITUTION', Regimecomp.CellValues[ACol,ARow], FALSE);
      if ((ResultRech <> 'Error') and (ResultRech <> '')) then
         begin
         sWhere:= 'SELECT PIP_ABREGE, POG_NUMAFFILIATION'+
                  ' FROM INSTITUTIONPAYE'+
                  ' LEFT JOIN ORGANISMEPAIE ON'+
                  ' PIP_INSTITUTION=POG_INSTITUTION WHERE'+
                  ' ##PIP_PREDEFINI##'+
                  ' PIP_INSTITUTION = "'+RegimeComp.CellValues[ACol,ARow]+'" AND'+
                  ' POG_ETABLISSEMENT = "'+GetControlText('CETABLISSEMENT')+'"';
         QRechInstit := OpenSql(sWhere,TRUE);
         if (not QRechInstit.EOF) then
            begin
            RegimeComp.CellValues[ACol+1,ARow]:= QRechInstit.Fields[0].asstring;
            RegimeComp.CellValues[ACol+2,ARow]:= QRechInstit.Fields[1].asstring;
            end
         else
            if (RegimeComp.CellValues[ACol,ARow]<>CodeRegime) then              //PT79-2
               begin
               RegimeComp.CellValues[ACol+1,ARow]:= '';
               RegimeComp.CellValues[ACol+2,ARow]:= '';
               end;
         Ferme(QRechInstit);
         end
      else
         if (RegimeComp.CellValues[ACol,ARow]<>CodeRegime) then                 //PT79-2
            begin
            Regimecomp.CellValues[ACol+1,ARow] := '';
            Regimecomp.CellValues[ACol+2,ARow] := '';
            end;
//PT73-8
      if ((RegimeComp.CellValues[ACol, 1] = 'I001') or
         (RegimeComp.CellValues[ACol, 1] = 'I002') or
         (RegimeComp.CellValues[ACol, 2] = 'I001') or
         (RegimeComp.CellValues[ACol, 2] = 'I002') or
         (RegimeComp.CellValues[ACol, 3] = 'I001') or
         (RegimeComp.CellValues[ACol, 3] = 'I002') or
         (RegimeComp.CellValues[ACol, 4] = 'I001') or
         (RegimeComp.CellValues[ACol, 4] = 'I002') or
         (RegimeComp.CellValues[ACol, 5] = 'I001') or
         (RegimeComp.CellValues[ACol, 5] = 'I002')) then
         ExisteIRCANTEC:= True;
      TabSheet:= TTabsheet(GetControl('TIRCANTEC'));
      if (TabSheet <> nil) then
         begin
         if (ExisteIRCANTEC=True) then
            TabSheet.TabVisible:= True
         else
            TabSheet.TabVisible:= False;
         end;
//FIN PT73-8
      end;
   end;
// DEB PT99
  if Sender = EXONER then
   begin
    if (Acol = 2) or (Acol = 3) then
    Exoner.CellValues[Acol, Arow] := FormatFloat(Exoner.ColFormats[Acol], Valeur(Exoner.Cells[Acol, Arow]));
   end;
// Fin PT99

if Sender = Basebrutespec then
   begin
   if (ACol = 0) then
      begin
      ResultRech:=RechDom('PGBASEBRUTESPEC', Basebrutespec.CellValues[ACol,ARow], FALSE);
      if ((ResultRech <> 'Error') and (ResultRech <> '')) then
         Basebrutespec.CellValues[ACol+1,ARow] := ResultRech
      else
         begin
         Basebrutespec.CellValues[ACol,ARow] := '';
         Basebrutespec.CellValues[ACol+1,ARow] := '';
         Basebrutespec.CellValues[ACol+2,ARow] := '';
         end;
      end
      else   // deb pt99
      begin
       if (Acol = 2) then
       BaseBruteSpec.CellValues[Acol, Arow] := FormatFloat(BaseBruteSpec.ColFormats[Acol], Valeur(BaseBruteSpec.Cells[Acol, Arow]));
      end;    // fin pt99

   end;

if Sender = Baseplafspec then
   begin
   if (ACol = 0) then
      begin
      ResultRech:=RechDom('PGBASEPLAFSPEC', Baseplafspec.CellValues[ACol,ARow], FALSE);
      if ((ResultRech <> 'Error') and (ResultRech <> '')) then
         Baseplafspec.CellValues[ACol+1,ARow] := ResultRech
      else
         begin
         Baseplafspec.CellValues[ACol,ARow] := '';
         Baseplafspec.CellValues[ACol+1,ARow] := '';
         Baseplafspec.CellValues[ACol+2,ARow] := '';
         end;
      end
      else // deb pt99
      begin
       if (Acol = 2) then
       BaseplafSpec.CellValues[Acol, Arow] := FormatFloat(BaseplafSpec.ColFormats[Acol], Valeur(BaseplafSpec.Cells[Acol, Arow]));
      end;    // fin pt99

   end;

if Sender = Somiso then
   begin
   if (ACol = 0) then
      begin
      ResultRech:=RechDom('PGDADSSOMISO', Somiso.CellValues[ACol,ARow], FALSE);
      if ((ResultRech <> 'Error') and (ResultRech <> '')) then
         Somiso.CellValues[ACol+1,ARow] := ResultRech
      else
         begin
         Somiso.CellValues[ACol,ARow] := '';
         Somiso.CellValues[ACol+1,ARow] := '';
         Somiso.CellValues[ACol+2,ARow] := '';
         Somiso.CellValues[ACol+3,ARow] := '';
         Somiso.CellValues[ACol+4,ARow] := '';
         end;
      end;
   end;

if Sender = Exoner then
   begin
   if (ACol = 0) then
      begin
      ResultRech:=RechDom('PGEXONERATION', Exoner.CellValues[ACol,ARow], FALSE);
      if ((ResultRech <> 'Error') and (ResultRech <> '')) then
         Exoner.CellValues[ACol+1,ARow] := ResultRech
      else
         begin
         Exoner.CellValues[ACol,ARow] := '';
         Exoner.CellValues[ACol+1,ARow] := '';
         Exoner.CellValues[ACol+2,ARow] := '';
         Exoner.CellValues[ACol+3,ARow] := '';
         end;
      end;
   end;

if Sender = Prevoyance then
   begin
   if ((ACol = 3) and (Prevoyance.CellValues[ACol,ARow] <> '')) then
      begin
      ResultRech:=RechDom('PGINSTITUTION', Prevoyance.CellValues[ACol,ARow], FALSE);
      if ((ResultRech = 'Error') or (ResultRech = '')) then
         Prevoyance.CellValues[ACol,ARow] := '';
      end;
   end;
end;

//PT84
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/12/2006
Modifié le ... :   /  /
Description .. : Entrée sur une grille
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.GrilleEnter (Sender: TObject);
begin
FocusRegime:= False;
end;
//FIN PT84

{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 14/11/2002
Modifié le ... :   /  /
Description .. : Modification de la valeur du motif d'entrée
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MotifChange(Sender: TObject);
var
sWhere : string;
QRechMotif : TQuery;
begin
{PT73-3
if (GetControlText('CDEBMOTIF1') = '095') then
   begin
   if GetControlText ('CFINMOTIF1') <> '096' then
      SetControlText('CFINMOTIF1', '096');
   SetControlEnabled('CFINMOTIF1', FALSE);
   end
else
   begin
   if (GetControlText('CFINMOTIF1') = '096') then
      begin
      SetControlText('CFINMOTIF1', '');
      SetControlEnabled('CFINMOTIF1', TRUE);
      end;
   end;

if (GetControlText('CDEBMOTIF1') <> '') then
   SetControlEnabled('CDEBMOTIF2', True)
else
   begin
   if (GetControlText('CDEBMOTIF2') <> '') then
      SetControlText('CDEBMOTIF2', '');
   SetControlEnabled('CDEBMOTIF2', False);
   end;

if (GetControlText('CDEBMOTIF2') <> '') then
   SetControlEnabled('CDEBMOTIF3', True)
else
   begin
   if (GetControlText('CDEBMOTIF3') <> '') then
      SetControlText('CDEBMOTIF3', '');
   SetControlEnabled('CDEBMOTIF3', False);
   end;

if (GetControlText('CFINMOTIF1') <> '') then
   SetControlEnabled('CFINMOTIF2', True)
else
   begin
   if (GetControlText('CFINMOTIF2') <> '') then
      SetControlText('CFINMOTIF2', '');
   SetControlEnabled('CFINMOTIF2', False);
   end;

if (GetControlText('CFINMOTIF2') <> '') then
   SetControlEnabled('CFINMOTIF3', True)
else
   begin
   if (GetControlText('CFINMOTIF3') <> '') then
      SetControlText('CFINMOTIF3', '');
   SetControlEnabled('CFINMOTIF3', False);
   end;
}
if (GetControlText('EDEBMOTIF1') = '095') then
   begin
   if GetControlText ('EFINMOTIF1') <> '096' then
      SetControlText('EFINMOTIF1', '096');
   SetControlEnabled('EFINMOTIF1', FALSE);
   end
else
   begin
   if (GetControlText('EFINMOTIF1') = '096') then
      begin
      SetControlText ('EFINMOTIF1', '');
      SetControlText ('LFINMOTIF1', '...');
      SetControlEnabled('EFINMOTIF1', TRUE);
      end;
   end;

if (GetControlText('EDEBMOTIF1') <> '') then
   SetControlEnabled('EDEBMOTIF2', True)
else
   begin
   if (GetControlText('EDEBMOTIF2') <> '') then
      SetControlText('EDEBMOTIF2', '');
   SetControlEnabled('EDEBMOTIF2', False);
   end;

if (GetControlText('EDEBMOTIF2') <> '') then
   SetControlEnabled('EDEBMOTIF3', True)
else
   begin
   if (GetControlText('EDEBMOTIF3') <> '') then
      SetControlText('EDEBMOTIF3', '');
   SetControlEnabled('EDEBMOTIF3', False);
   end;

if (GetControlText('EFINMOTIF1') <> '') then
   SetControlEnabled('EFINMOTIF2', True)
else
   begin
   if (GetControlText('EFINMOTIF2') <> '') then
      SetControlText('EFINMOTIF2', '');
   SetControlEnabled('EFINMOTIF2', False);
   end;

if (GetControlText('EFINMOTIF2') <> '') then
   SetControlEnabled('EFINMOTIF3', True)
else
   begin
   if (GetControlText('EFINMOTIF3') <> '') then
      SetControlText('EFINMOTIF3', '');
   SetControlEnabled('EFINMOTIF3', False);
   end;

if (Sender=DebMotif1) then
   begin
   LeMotif.ControlMotif:= DebMotif1;
   LeMotif.TitreMotif:= '1er motif de début';
   LeMotif.TableMotif:= 'MOTIFENTREESAL';
   LeMotif.ColonneMotif:= 'PME_CODE';
   LeMotif.SelectMotif:= 'PME_LIBELLE';
   LeMotif.WhereMotif:= '';
   LeMotif.EditMotif:= 'EDEBMOTIF1';
   LeMotif.LabelMotif:= 'LDEBMOTIF1';
   end
else
if (Sender=FinMotif1) then
   begin
   LeMotif.ControlMotif:= FinMotif1;
   LeMotif.TitreMotif:= '1er motif de fin';
   LeMotif.TableMotif:= 'MOTIFSORTIEPAY';
   LeMotif.ColonneMotif:= 'PMS_DADSU';
   LeMotif.SelectMotif:= 'PMS_LIBELLEDADS';
   LeMotif.WhereMotif:= 'PMS_DADSU <> "000"';
   LeMotif.EditMotif:= 'EFINMOTIF1';
   LeMotif.LabelMotif:= 'LFINMOTIF1';
   end
else
if (Sender=DebMotif2) then
   begin
   LeMotif.ControlMotif:= DebMotif2;
   LeMotif.TitreMotif:= '2ème motif de début';
   LeMotif.TableMotif:= 'MOTIFENTREESAL';
   LeMotif.ColonneMotif:= 'PME_CODE';
   LeMotif.SelectMotif:= 'PME_LIBELLE';
   LeMotif.WhereMotif:= '';
   LeMotif.EditMotif:= 'EDEBMOTIF2';
   LeMotif.LabelMotif:= 'LDEBMOTIF2';
   end
else
if (Sender=FinMotif2) then
   begin
   LeMotif.ControlMotif:= FinMotif2;
   LeMotif.TitreMotif:= '2ème motif de fin';
   LeMotif.TableMotif:= 'MOTIFSORTIEPAY';
   LeMotif.ColonneMotif:= 'PMS_DADSU';
   LeMotif.SelectMotif:= 'PMS_LIBELLEDADS';
   LeMotif.WhereMotif:= 'PMS_DADSU <> "000"';
   LeMotif.EditMotif:= 'EFINMOTIF2';
   LeMotif.LabelMotif:= 'LFINMOTIF2';
   end
else
if (Sender=DebMotif3) then
   begin
   LeMotif.ControlMotif:= DebMotif3;
   LeMotif.TitreMotif:= '3ème motif de début';
   LeMotif.TableMotif:= 'MOTIFENTREESAL';
   LeMotif.ColonneMotif:= 'PME_CODE';
   LeMotif.SelectMotif:= 'PME_LIBELLE';
   LeMotif.WhereMotif:= '';
   LeMotif.EditMotif:= 'EDEBMOTIF3';
   LeMotif.LabelMotif:= 'LDEBMOTIF3';
   end
else
if (Sender=FinMotif3) then
   begin
   LeMotif.ControlMotif:= FinMotif3;
   LeMotif.TitreMotif:= '3ème motif de fin';
   LeMotif.TableMotif:= 'MOTIFSORTIEPAY';
   LeMotif.ColonneMotif:= 'PMS_DADSU';
   LeMotif.SelectMotif:= 'PMS_LIBELLEDADS';
   LeMotif.WhereMotif:= 'PMS_DADSU <> "000"';
   LeMotif.EditMotif:= 'EFINMOTIF3';
   LeMotif.LabelMotif:= 'LFINMOTIF3';
   end;

sWhere:= 'SELECT '+LeMotif.SelectMotif+
         ' FROM '+LeMotif.TableMotif+' WHERE '+
         LeMotif.ColonneMotif+' = "'+GetControlText (LeMotif.EditMotif)+'"';
QRechMotif := OpenSql(sWhere,TRUE);
if (not QRechMotif.EOF) then
   SetControlText(LeMotif.LabelMotif, QRechMotif.Fields[0].asstring);
Ferme(QRechMotif);
//FIN PT73-3
end;

//PT73-3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/05/2006
Modifié le ... :   /  /
Description .. : Modification de la valeur du motif d'entrée par click
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.MotifChangeClick(Sender: TObject);
var
StSelect, sWhere : string;
begin
MotifChange (Sender);

{PT89
sWhere:= 'SELECT '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif+
         ' FROM '+LeMotif.TableMotif+' WHERE '+
         LeMotif.WhereMotif+' GROUP BY '+LeMotif.ColonneMotif+', '+
         LeMotif.SelectMotif;
}
StSelect:= 'SELECT '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif+
           ' FROM '+LeMotif.TableMotif;
if (LeMotif.WhereMotif<> '') then
   begin
   sWhere:= StSelect+' WHERE '+LeMotif.WhereMotif;
   sWhere:= sWhere+' GROUP BY '+LeMotif.ColonneMotif+', '+LeMotif.SelectMotif;
   end;
{$IFNDEF EAGLCLIENT}
LookUpList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, sWhere);
{$ELSE}
LookupList (LeMotif.ControlMotif, LeMotif.TitreMotif, LeMotif.TableMotif,
            LeMotif.ColonneMotif, LeMotif.SelectMotif, LeMotif.WhereMotif,
            LeMotif.ColonneMotif, TRUE, -1, StSelect);
{$ENDIF}
//FIN PT89

MotifChange (Sender);
end;
//FIN PT73-3


{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 14/11/2002
Modifié le ... :   /  /
Description .. : Modification de la valeur de la condition d'emploi
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.CondEmploiChange(Sender: TObject);
begin
if (GetControlText('CCONDEMPLOI') = 'P') then
   SetControlEnabled('ETAUXPARTIEL', TRUE)
else
   begin
   SetControlText('ETAUXPARTIEL', '');
   SetControlEnabled('ETAUXPARTIEL', FALSE);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
case Key of
     VK_F10: if ((GetControlVisible('BDACCORD')) and
               (GetControlEnabled('BDACCORD'))) then
               Daccord.Click; //Validation
     end;

TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F6: if ((GetControlVisible('BCALCULER')) and
               (GetControlEnabled('BCALCULER'))) then
               Calculer.Click; //Calcul des éléments
     ord('S'): if ((GetControlVisible('BSALARIE')) and
                  (GetControlEnabled('BSALARIE')) and (ssAlt in Shift)) then
                  Salarie.Click; //Visualisation de la fiche salarié
     ord('M'): if ((GetControlVisible('BSALMAJ')) and
                  (GetControlEnabled('BSALMAJ')) and (ssAlt in Shift)) then
                  SalMAJ.Click; //MAJ Fiche salarié
     VK_F3: begin
            if ((GetControlVisible('BSALPREM')) and
               (GetControlEnabled('BSALPREM')) and (ssCtrl in Shift)) then
               SalPrem.Click      //Premier salarié
            else
               if ((GetControlVisible('BSALPREC')) and
                  (GetControlEnabled('BSALPREC')) and (Shift = [])) then
                  SalPrec.Click;     //Salarié précédent
            end;
     VK_F4: begin
            if ((GetControlVisible('BSALDERN')) and
               (GetControlEnabled('BSALDERN')) and (ssCtrl in Shift)) then
                SalDern.Click     //Dernier salarié
            else
                if ((GetControlVisible('BSALSUIV')) and
                   (GetControlEnabled('BSALSUIV')) and (Shift = [])) then
                   SalSuiv.Click;     //Salarié suivant
            end;
     ord('P'): if ((GetControlVisible('BPERPREC')) and
                  (GetControlEnabled('BPERPREC')) and (ssAlt in Shift)) then
                  PerPrec.Click; //Période précédente
     ord('N'): if ((GetControlVisible('BPERSUIV')) and
                  (GetControlEnabled('BPERSUIV')) and (ssAlt in Shift)) then
                  PerSuiv.Click; //Période suivante
     end;
end;

//DEB PT92-1
procedure TOF_PG_DADSPERIODES.ChangeTypeRegime(Sender: TObject);
var
  Chk : TCheckBox;
begin
  inherited;
  Chk := TCheckBox(Sender);
  if (Chk.Checked = True) then
  begin
    SetControlEnabled('CREGIMESS', False);
    SetControlText('CREGIMESS', '');
    SetControlEnabled('CREGIMEMAL', True);
    SetControlEnabled('CREGIMEAT', True);
    SetControlEnabled('CREGIMEVIP', True);
    SetControlEnabled('CREGIMEVIS', True);
    if (GetControlText('CREGIMEMAL') = '') then SetControlText('CREGIMEMAL', '200');
    if (GetControlText('CREGIMEAT') = '') then SetControlText('CREGIMEAT', '200');
    if (GetControlText('CREGIMEVIP') = '') then SetControlText('CREGIMEVIP', '200');
    if (GetControlText('CREGIMEVIS') = '') then SetControlText('CREGIMEVIS', '200');
  end
  else
  begin
    SetControlEnabled('CREGIMESS', True);
    if (GetControlText ('CREGIMESS') = '') then SetControlText('CREGIMESS', '200');
    SetControlEnabled('CREGIMEMAL', False);
    SetControlEnabled('CREGIMEAT', False);
    SetControlEnabled('CREGIMEVIP', False);
    SetControlEnabled('CREGIMEVIS', False);
    SetControlText('CREGIMEMAL', '');
    SetControlText('CREGIMEAT', '');
    SetControlText('CREGIMEVIP', '');
    SetControlText('CREGIMEVIS', '');
  end;
end;
//FIN PT92-1

//PT102
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/12/2007
Modifié le ... :   /  /    
Description .. : Double-click sur la grille prévoyance
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_DADSPERIODES.PrevoyanceDblClick (Sender: TObject);
var
Code, ResultRech, Ret, StEvt: string;
QRechEvt : TQuery;
i, Rep : integer;
begin
if (Prevoyance.Col=0) then
   begin
   StEvt:= 'SELECT CO_CODE'+
           ' FROM COMMUN WHERE'+
           ' CO_TYPE="PDR" AND'+
           ' CO_LIBELLE="'+Prevoyance.CellValues[Prevoyance.Col,Prevoyance.Row]+'"';
   QRechEvt:= OpenSQL(StEvt,TRUE) ;
   if Not QRechEvt.EOF then
      ResultRech:= QRechEvt.FindField ('CO_CODE').AsString;
   Ferme (QRechEvt);

   if ((ResultRech<>'Error') and (ResultRech<>'')) then
      Code:= ResultRech;
   Ret:= AGLLanceFiche ('PAY', 'DADS_PREVEVT', '', '', Code);
   if ((Ret='') or (Ret='Error')) then
      begin
      Rep:= PGIAsk ('Voulez-vous supprimer la ligne de contrat en cours ?',
                  'Contrat de prévoyance');
      if (Rep=mrYes) then
         begin
         for i:=0 to 10 do
             Prevoyance.CellValues[i,Prevoyance.Row]:= '';
         end;
      end
   else
      Prevoyance.CellValues[Prevoyance.Col,Prevoyance.Row]:= RechDom ('PGDADSPREV', Ret, FALSE);
   end;
end;
//FIN PT102


Initialization
registerclasses([TOF_PG_DADSPERIODES]) ;

end.

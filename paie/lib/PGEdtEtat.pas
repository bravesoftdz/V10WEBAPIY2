{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Edition du calendrier, temps de travail, Absences..
Mots clefs ... : PAIE,ABSENCE,PGDUCS
*****************************************************************
PT1   : 10/09/2001 SB V547 Violation d'acces : RendQueryCalendrier : Au lieu de
                           passer la requête en paramètre je passe une string
PT2   : 14/09/2001 SB V547 Bug modèle d'etat des cumuls
                           Fn appellée GETMONTRUB au lieu de GETMONTCUM
                           création var PGEDITION
PT3   : 14/09/2001 SB V547 Cumul de reprise non intégré dans l'edition des
                           cumuls
                           Deplacement de code
PT4   : 28/09/2001 SB V547 Edition des cumuls variable WhereCumul à déclarer en
                           globale
PT5   : 01/10/2001 SB V560 Multiplicateur retourné erroné..

PT6   : 05/10/2001 MF V652 Modification édition DUCS   - PGDUCS
                           1- Sous titre particulier pour DUCS IRC Annuelle
                           2- Les ligne de Sous Totaux ne font pas partie du
                              total déclaré
                           3- Mise en place de la bande overlay (partie gauche
                              du corps)
PT7   : 15/10/2001 MF V562 Modification édition DUCS - PGDUCS
                           1- La table INSTITUTIONPAYE remplace la table
                              INSTITUTIONPAIE
PT8   : 08/11/2001 MF V562 Modification édition DUCS - PGDUCS
                           1- Mise en commentaire code inutile
                           2- Correction calcul effectif total
                           3- Correction de l'édition des montants en pied de
                              page, en particulier les montants négatifs
                           4- Correction de l'édition de blancs pour les
                              réductions bas salaire Urssaf.
PT9   : 09/11/2001 MF V652 Modification édition DUCS - PGDUCS
                           suite a modification des valeurs de la tablette
                           PGMONNAIE, correction de la méthode d'édition de la
                           monnaie .
PT10  : 12/11/2001 SB V562 Date de reprise des cumuls en edition de bulletin
                           On ne tient pas compte de la date entrée..
PT11  : 16/11/2001 SB V562 DateDebutExerSoc déclaré en globale
PT12  : 20/11/2001 MF V562 PGDUCS
                           correction conversion Nbre d'hommes et Nbre de femmes
                           quand non renseigné.
PT13  : 07/12/2001 SB V563 Fiche de bug n°355 Les montants additionnés =0 ne
                           s'imprime pas
PT14  : 07/01/2002 MF V571 PGDUCS
                           1- Cas des lignes de sous-totaux pour IRC, le nom de
                              la caisse est limité à 7 caractères
                           2- Cas des forfaits : Dans la colonne Taux on édite
                               l'effectif.
                           3- Traitement du champ PDD_LIBELLESUITE Le nbre de
                              ligne détail est incrémenté de 1.
PT15  : 08/01/2002 SB V571 Edition de cumul : cumul affecté au 2nd bulletin du
                           mois ne s'edite pas si premier bulletin non affecté
PT16  : 16/01/2002 SB V571 Fiche de bug n°414 : Suppression de la rupture
                           périodique
PT17  : 18/01/2002 SB V571 Fiche de bug n°420 : Edition du journal de paie ajout
                           Tranche 1 et 2 et de la base csg et crds
PT18  : 18/01/2002 JL V571 Modification des TOB pour éditions DADS-U et ajout
                           fonction FunctPGDADSUNbPer >> mis en commentaire le
                           30/01/2002
PT19  : 23/01/2002 MF V571 PGDUCS
                           On ne fait plus de ColleZeroDevant sur la variable
                           Etab. Correction : pour les établissement numérotés
                           "1  " le montant déclaré était à blanc.
PT20  : 23/01/2002 SB V571 Edition des cumuls allègement du traitement
PT21-1: 25/01/2002 SB V571 Edition DADS : Les requêtes FN n'appliquent pas les
                           critères de selection de la fiche : Edition pour un
                           étab => Chargement de tous les Etabs => Temps de
                           traitement > ++
PT21-2: 25/01/2002 SB V571 Déallocation mémoire  : Contrôle tob<> nil
PT22  : 28/01/2002 SB V571 Selection alphanumérique : modification de critères
                           de selection
PT23  : 29/01/2002 JL V571 Modification de la requête pour édition
                           récapitulative pour gestion des organismes.
PT24  : 30/01/2002 SB V571 Calcul Homme Femme etat des Charges Sociales
PT25  : 14/02/2002 MF V571 PGDUCS
                           Adaptation des fonctions d'édition de la DUCS pour le
                           traitement des ruptures
PT26  : 13/03/2002 SB V571 Fiche de n°10040 Cumul mois en cours non réinitialisé
                           en édition de bulletin
PT27  : 25/03/2002 SB V571 Fiche de n°415 Edition de l'établissement sur fiche
                           individuelle
PT28  : 29/03/2002 MF V571 PGDUCS
                           Modification du traitement des ducs IRC (cas
                           personnalisé)
PT29  : 29/04/2002 SB V571 Fiche de bug n°306,367,10081 : Calcul du décompte des jours eronnés
                           Modification de la fn IfJourNonTravaille intégrant
                           des contrôles de jours spécifiques
                           Ajout d'arrondi pour l'affichage des heures
PT30  : 30/05/2002 SB V582 Controle de vraisemblance, Jour spécifique
                           réinitialise systématique le jour comme travaillé
PT30-2: 07/06/2002 SB V852 Controle de vraisemblance, Calendrier bimensuel
                           personnalisé non récupérer
PT31  : 11/06/2002 SB V582 FnGetLesCumul : Restriction de la requête au cumul
                           édité..
PT32  : 12/06/2002 SB V585 Fiche de bug n° 10109 : Edition Calendrier, décompte
                           erroné
PT33  : 12/06/2002 SB V582 Fiche de bug n° 10157 : Edition du journal de paie
                           Anomalie dans totalisation, Les montants du dernier
                           salarié ne sont pas totalisés
PT34  : 21/06/2002 MF V585 PGDUCS
                           Traitement des fonctions propres à la VLU (Annexe à
                           la DUCS) - PGDUCS
PT35  : 25/07/2002 SB V585 Dû à l'intégration des mvts d'annulation d'absences
                           Contrôle des requêtes si typemvt et sensabs en
                           critere
PT36  : 30/07/2002 SB V585 FQ n° 10200 Conversion monnaie inversée des tranches
                           et bases sur journal de paie
PT37  : 05/09/2002 MF V585 PGDUCS
                           rectification du formatage de la zone code-libellé
PT38  : 10/09/2002 MF V585 PGDUCS
                           Fiche qualité 10018 : Rectification du libellé :
                           'Nombre de salarié ou d'assurés..."
PT39  : 11/09/2002 SB V582 Mise en commentaire de l'appel de la moulinette de
                           maj des bases Cp erronées
PT40  : 18/09/2002 SB V585 FQ n°10022 Récupération de l'organisme à editer sur
                           bulletin paramétrer au niveau etab ou salarié
PT41  : 30/09/2002 VG V585 Adaptation cahier des charges DADS-U V6R02
PT42-1: 02/10/2002 SB V585 FQ n°10212 Critère non récupéré pour calcul H/F
                           charges sociales
PT42-2: 02/10/2002 SB V585 FQ n°10214 Bulletin bimensuel non edité si selection
                           mensualité de fin
PT43  : 10/10/2002 SB V585 Optimisation traitement d'edition
PT43-1  16/12/2002 SB V585 FQ 10123 intégration confidentialité
PT44  : 11/10/2002 JL V585 DADS mention aucun organisme si code = 90 000
PT45  : 14/11/2002 VG V585 Edition des traitements et salaires DADSU
PT46  : 26/11/2002 JL V591 DADS : édition récapitulative modification si pas de
                           rupture organisme
PT47  : 16/12/2002 JL V591 DADS : Civilité  M. au lieu de MR
PT48-1: 20/12/2002 SB V591 FQ 10391 Calendrier pris en compte de la date entrée
                           et sortie pour calcul tps de travail
PT49  : 06/01/2003 SB V591 FQ 10426 ProvisionCP : Totalisation etab non raz..
PT50  : 08/01/2003 SB V585 Calcul de la date de RAZ des cumuls érroné si
                           mois défini>mois édité
PT51  : 29/01/2003 SB V585 Chargemt des organismes depuis de début et fin de
                           mois période et non date premier bulletin
PT52  : 05/02/2003 SB V591 Reprise des cumuls au debut du mois donné
PT53  : 07/02/2003 SB V591 FQ 10396 Edition Bulletin : zone organisme dont
                           établissement sur - de 3 caractères
PT54  : 14/02/2003 VG V_42 Initialisation du result de la fonction
                           FunctPGDADSUSalaires
PT55  : 17/02/2003 SB V_42 Optimisation traitement d'édition
PT56  : 26/02/2003 VG V_42 Edition des périodes d'inactivité DADSU - FQ N° 10463
PT57  : 19/03/2003 SB V_42 FQ 10597 réinitialisation du rib bancaire
                           systématique et non que pour les virements
PT58  : 17/03/2003 SB V_42 Optimisation code pour editon Fiche individuelle et
                           Journal de paie
PT59  : 25/03/2003 SB V_42 Incremantation d'année de DateDebutExerSoc si edition
                           périodique à cheval sur pls exercices
PT60  : 03/04/2003 MF V_42 Fiche qualité n° 10457 (édition libellé ST pour IRC
                           standardisée. - PGDUCS
PT61  : 10/04/2003 MF V_42 Modification liées aux ducs récapitulatives - calcul
                           du  net à payer
PT62  : 11/06/2003 SB V_42 FQ 10598 & 367 Calendrier : Décompte des jours
                           ouvrables non travaillés sur la grille
PT63-1: 04/08/2003 SB V_42 Journal de paie : init des cumuls à zero déplacés
PT63-2: 04/08/2003 SB V_42 FQ 10746 Journal de paie : Ajout de la tranche 3
PT64  : 12/08/2003 VG V_42 Edition de la DADS Bilatérale
PT65  : 12/08/2003 SB V_42 FQ 10755 Edition Bulletin : division par zéro
                           impossible
PT66  : 21/08/2003 JL V_42 FQ 10778 Correction requête médecine travail pour
                           oracle (DATE=mot réservé)
PT67  : 03/09/2003 PH V_42 Edition des variables : impression des libellés des
                           bases de cotisation
PT68  : 14/05/2003 JL V_421 FQ 10662 nature du contrat non trouver pour CDD
PT69-1: 10/09/2003 VG V_42 Initialisation du résultat lors de la DADS
                           récapitulative
                   JL V_42 Idem pour autres fcts de libération de TOB  (FQ 10814)
PT69-2: 10/09/2003 VG V_42 Les noms des fonctions doivent être en majuscule
PT70  : 11/09/2003 VG V_42 Les segments des états de la DADS-U ne sont plus
                           appellés par le champ ORDRESEG (qui change à chaque
                           version !) mais par le champ SEGMENT
PT71  : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT72  : 16/10/2003 VG V_42 Modification de l'état PDA/PDB
PT73  : 22/10/2003 VG V_42 Retour en arrière de l'AGL (idem PT69-2)
PT74  : 24/10/2003 VG V_42 Ajout de critères pour les états DADS-U
PT75-1: 27/11/2003 SB V_50 Publication dans CalcOleGenericPaie des fonctions
PT75-2: 27/11/2003 SB V_50 Suppression des variables globales
PT76    09/12/2003 SB V_50 FQ 10370 Edition compteurs Cp selon paramétrage
PT77    18/12/2003 MF V_50 Fcts DUCS - StrToFloat remplacée par Valeur
PT78  : 19/01/2003 JL V_50 Ajout contôlre date sur DADSDETAIL pour erreur doublon
PT79    15/01/2004 JL V_50 Correction édition frais formations
PT80  : 06/02/2004 MF V_50 FQ 11087 : calcul effectif total pour URSSAF (les apprentis
                           (n'en font pas partie)
PT82  : 06/04/2004 VG V_50 Correction édition DADS-U BTP
PT83  : 15/04/2004 SB V_50 Renomminaton des noms de champ de la tob provision
PT84  : 29/04/2004 SB V_50 FQ 11263 Edition du solde Cp à identique de la provision CP
PT85  : 09/07/2004 VG V_50 Adaptation cahier des charges DADS-U V8R00
PT86  : 09/10/2004 PH V_50 On force la recherche de l'exercice social si chgt de critère sinon ??????
PT87  : 12/10/2004 PH V_50 Compatibilité ascendante des états avec la V5. Il n'y avait pas le compteur de pages
                           passé en paramètre
PT88  : 19/11/2004 MF V_60 FQ 11714 : DUCS correction du calcul du nbre de pages.
PT89  : 20/12/2004 PH V_60 Tous les tableaux de string contenant des doubles formatés
PT90  : 25/12/2004 PH V_60 FQ 11922 gestion des exercices CP pour provisions CP et RTT rajout fct INITPROVDOTCP
PT91  : 02/02/2005 VG V_60 Edition DADS - FQ N°11957
PT92  : 03/02/2005 PH V_602 FQ 11337 Prise en compte 1ere position du cumul en alpha ==> Gestion dynamique des tableaux
                      de cumuls
PT93  : 25/03/2005 SB V_60 FQ 11990 Ajout option pour rech. etab de l'historique
PT94  : 10/05/2005 MF V_602 FQ 12267 On ne fait plus de ColleZeroDevant sur la
                            variable Organisme. Correction : pour les types
                            d'organisme numérotés "13" le montant déclaré était
                            à 0.
PT95  : 26/05/2005 SB V_60  FQ 12308 Modification de l'ordre de présentation des rubriques
PT96  : 27/06/2005 PH V_60 Compatibilité Etats provisions CP et RTT par module Dotations
PT97  : 03/10/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT98  : 03/10/2005 SB V_65 FQ 12626 Bulletin : Anomalie suite cumul alphanumérique
PT99  : 18/01/2006 SB V_65 modification de la clause de Confidentialité
PT100   23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT101 : 06/02/2006 MF V_65 FQ 12180 correction récupération arrondi montant et base de la codif.
PT102 : 13/03/2006 RM V_65 Ajout nouvelle fonction FUNCTPGCALCULANCIENNETE
---- JL 20/03/2006 modification clé annuaire ----

PT103   07/04/2006 SB V_65 FQ 12941 Traitement nouveau mode de décompte motif d'absence
PT104   14/04/2006 SB V_65 FQ 12696 Correction anomalie calendrier absence bimensuel
PT105   19/05/2006 SB V_65 Appel des fonctions communes pour identifier V_PGI.driver
PT106   01/06/2006 SB V_65 Suppression fonction pour traitement de la reprise CP
PT107   16/06/2006 SB V_65 FQ 12670 Rech. de l'organisme à editer sur bulletin
PT108   19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT109   20/07/2006 SB V_65 FQ 12813 Edit. bulletin Refonte rech. date debut exercice social
PT110   27/07/2006 GGS V_70 FQ 12911 Prise en compte des salariés avec date sortie < idat1900
PT111   03/01/2007 SB V_70 Conversion en CWAS pour les états libres sur provisions
PT112   09/02/2007 MF V_702 FQ 13070 : Traitement des codifications ALSACE MOSELLE
PT113   22/03/2007 FC V_70 FQ 13356 : rajout de l'ancienneté en toutes lettres via la fonction GETTXTANCIENNETE
PT115   26/04/2007 FC V_70 Libération mémoire
PT116 : 30/04/2007 FL V_720 Pour l'édition des frais par stagiaire, on ne passe plus par des fonctions @.
PT117   23/05/2007 JL V_720 Modif clé table fraissalform
PT118   12/06/2007 GGU V_72 FQ 13541 : Un effectif global équivalent temps plein
PT119   13/07/2007 FC V_72 FQ 13326 tenir compte du paramètre société pour calculer la date prévision visite médicale
PT120   28/08/2007 FC V_80 Externalisation des fonctions @ d'édition du bulletin dans une DLL
PT121   03/10/2007 FC V_80 Suppression fonction ReadParam + correction bug quand le dernier bulletin est sur 2 pages
PT122   09/10/2007 NA Modif fonction Fncreatecalendrier : plantage de Tbis.dupliquer
PT123   11/10/2007 Mf V_80 mise en place loi TEPA
PT124   09/11/2007 FC V_80 FQ 14889 Prendre le RIB SALAIRE au lieu de prendre le RIB PRINCIPAL
}
unit PGEdtEtat;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  SysUtils, HEnt1, HCtrls,
  Forms, HDebug,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DB,  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTOB, CalcOLEGenericAFF, PgOutils2,
  ParamSoc,ULibEditionPaie;

//function ReadParam(var St: string): string; //PT121
function CalcOLEEtatPG(sf, sp: string): variant;
//PT120 function FunctPgEditBulletin(Sf, Sp: string): variant;
//PT120 procedure FnCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page: string); { PT75-2 }
{ PT75-1 Publication dans CalcOleGenericPaie des fonctions
              - Function  FunctPgEditFicInd
              - Function  FnCreateFid
              - Function  FunctPgEditCumul
              - Function  FnGetReduction & FnGetLesReduction => FunctPgEditReduction
 Suppression de :
              - FnCreateFidRecap (fonctionnement intégré dans FnCreateFid)      }

function FunctPgEditJournalPaie(Sf, Sp: string): variant;
function FnGetLesCumul(DD, DF, StWhere: string): string;
function FunctPgEditProvisionCP(Sf, Sp: string): variant;
function FunctPGEditDucs(Sf, Sp: string): variant; // PGDUCS
function FunctPGFinEditDucs(): variant; // PGDUCS
function FunctPGDebEditDucs(): variant; // PGDUCS
function FunctPGDADSU(sf, sp: string): variant;
function FunctPGDADSUHonoraires(sf, sp: string): variant;
function FunctPGDADSURecap(sf, sp: string): variant;
function FunctPGDADSUSalaires(sf, sp: string): variant;
function FunctPGDADSUInactivite(sf, sp: string): variant;
function FunctPGDADSUNumPeriode(sf, sp: string): variant;
//Function  FunctPGDADSUNbPer(sf,sp: string) : variant ;    //PT18
// MEDECINE DU TRAVAIL
function FunctVisitesMed(Sf, Sp: string): Variant;
function FunctMedecineAge(Sf, Sp: string): Variant;
function FunctMedecineVisiteSuiv(Sf, Sp: string): Variant;
// MOUVEMENT MAIN OEUVRE
function FunctMMOEtab(Sf, Sp: string): Variant;
function FunctDateMMO(Sf, Sp: string): Variant;
function functContratMMO(Sf, Sp: string): Variant;
function FunctSortieMMO(Sf, Sp: string): Variant;
function FunctEntreeMMO(Sf, Sp: string): Variant;
//MSA
function MSAHeuresMontant(Sf, Sp: string): Variant;
function LibDateMSA(Sf, Sp: string): Variant;
function CreateTobMSA(Sf, Sp: string): Variant;

function CreateCritere(ChampRupt, ValRupt, StWhere: string): string;
function ConvertPrefixe(StWhere, DePref, APref: string): string;
function FnCreateCalendrier(Salarie: string; DateDeb: TDateTime; GDCalendrier: THGrid; Tob_MotifAbs : Tob): string; { PT103 }
procedure ReinitGrille(GDCalendrier: THGrid; DateDeb: TDateTime);
function InitCumulAbsence(Salarie, Etab: string; DateDeb, DateFin, DateEntree, DateSortie: TDateTime): string;
procedure AffectCumulAbsence(GDCalendrier: THGrid; C, L: integer; TypeConges, JourHeure: string);
//Procedure InitTobCalendrier(DateDeb : TDateTime);
//Formation
function InitTobFormation(Sf, Sp: string): Variant;
{PT116
function InitTobFrais(Sf, Sp: string): Variant;
function InitLibelleFrais(Sf, Sp: string): Variant;
function AfficheLibFrais(Sf, Sp: string): Variant;
function AfficheFrais(Sf, Sp: string): Variant;
}
function InitTobComparatif(Sf, Sp: string): Variant;
function FormationComparatif(Sf, Sp: string): Variant;
function InscBudgetComparatif(Sf, Sp: string): Variant;
function CreateTobFormationGlobal(Sf, Sp: string): Variant;
function FormationGlobal(Sf, Sp: string): Variant;
function CreateTobForRealise(Sf, Sp: string): Variant;
function RealiseForMensuel(Sf, Sp: string): Variant;
function CreateTobSuviRespF(Sf, Sp: string): Variant;
function SuiviResponsableFormation(Sf, Sp: string): Variant;
function CreateTobBudgetFor(Sf, Sp: string): Variant;
function CalCoutBudgetFor(Sf, Sp: string): Variant;
function CreateTobGestIndFor(Sf, Sp: string): Variant;
function GestIndFormation(Sf, Sp: string): Variant;
function InitTobSuiviGestInd(Sf, Sp: string): Variant;
function SuiviGestFor(Sf, Sp: string): Variant;
function InitTobGestIndComp(Sf, Sp: string): Variant;
function InitTobForSuiviGestColl(Sf, Sp: string): Variant;
function ForSuiviGestColl(Sf, Sp: string): Variant;
function RendNbStagiaresParAge(Sf, Sp: string): Variant;
// Congés spectacles
function CongesSpectacles(Sf, Sp: string): Variant;
function MajTableCongesSpectacles(Sf, Sp: string): Variant;
//DADS Bilatérale
function FunctPGDADSB(sf, sp: string): variant;
//PT102 Calcul l'age à partir d'une date donnée en entrée (Ancienneté)
function FunctPGCalculAnciennete(Sf, Sp: string): Variant;

var
  PgChampRupt {,PgTauxConvert,PGTypeFiche}: string; { PT75-2 }
  PGGlbEtabDe, PGGlBEtabA, PGGlbSalDe, PGGlbSalA, PGGlbOrg, PGGlbEtabMMO: string; //PT21-1
  PGGlbTypeDADS: string;
  PGSalarie: boolean; { PT75-2 }
  PGCalendrierEtat_Ecran: TForm;
//  PGGrilleRepriseCP: THGrid; PT106
  LigneOptique: string; //PGDUCS

  MemDateDeb, MemDateFin: TDateTime;
  MemEtab: string;


implementation

uses PgOutils, PgEditOutils, EntPaie, PGCongespayes, P5Util, PGCalendrier, pgcommun
{$IFNDEF EABSENCES}
{$IFNDEF EMANAGER}
  , OlePaiePgi
{$ENDIF}
{$ENDIF}
  , CalcOLEGenericPaie { PT75-1 }
  ;



var
  TobPaie, TobEtat, TRech, TCumPaie: TOB;
  Tob_CumulAbsence, Tob_AbsMois1: Tob;
  TCum, TRubCum, TobProvisionCp, TMvtSal: TOB;
  Tob_OrganismeBull, Tob_DateClotureCp, Tob_TauxBull: Tob;
  DADSUCodeEtab: string;
  Niveau, StWhere, LibBanque, organisme, cumul21: string;
  TabCumul, TabSens: array of string;  // PT92
  TabCumulEx, TabCumulMois: array of double; // PT89 // PT92
  TabMontCum: array[1..12] of Double;
  TabJour: array[1..7] of Char;
  Brut, BrutAbbat, Retenue, Patronal, Net: double;
  Assujet, DiversPlus, DiversMoins, CoutGlobalSal, Tranche1, Tranche2, Tranche3, BaseCsg, BaseCrds: double;
  NbRepasPris, Allegement, Majoration, Minoration: double;
  NbMois: integer;
  PN, AN, RN, BN, PN1, AN1, RN1, BN1: Double;
  DDP, DFP, DDPN1, DFPN1, DateDN1, DateFN1, DateDN: TDateTime;
  CumulMonnaie: Boolean;
  CMDD, CMDF, GblDateDeb, DebRepriseMois1, DebRepriseMois2, GblDateSortie, GblDateEntree, DateDebutExerSoc: TDateTime; //PT11 //PT48-1
  CMSal, CMEtab, GblSalarie, GblEtab, GblJourHeure, NumMois, WhereCumul: string; //PT-4
  Indice: Integer;
  Nb_JourCumul, Nb_HeureCumul, EffectifJour1, EffectifHeure1, TotNbjour, TotNbHeure: double;
  EffectifJour2, EffectifHeure2: double;
  BasePP, MoisPP, RestantsPP, BaseN1, MoisN1, AcquisN1, PrisN1, RestantsN1: double;
  BaseN, MoisN, AcquisN, PrisN, RestantsN, BasePro: double;
  TOBDADSHON, TOBDADSPER, TOBDADSRECAP, TobDadsNumPer, TobMedecine, TobContratMMO: Tob;
  TobMMOEtab, TobPaieMMO1, TobPaieMMO2, TOBDADSSAL, TOBDADSETAB: Tob;
  TobMSAP1, TobMSAP2, TobMSAP3: Tob;
  TobEntreeMMO, TobSortieMMO: Tob;
  DateDebDADS, DateFinDADS: TDateTime;
  MontantEuro: string; //Chèques
  Position: Char; //Chèques
  numposition: Integer; //Chèques
  TestCheque: Boolean; //Chèques
  Tob_JournalPaie, Tob_TypeCumul: Tob;
  TobFraisFor, TobInscForm, TobFormation: Tob; //FORMATION
  LesFrais: array[1..15] of string; //FORMATION
  TobCoutStageBud, TobForInscritsNom, TobForInscritsNonNom, TobForRealise, TobForRealiseInsc: Tob; //FORMATION
  TobFGFrais, TobFGInsc, TobFGStage, TobFGSessions, TobFGFormations, TobGestIndRealise, TobGestIndBudget, TobGestIndStage, TobGestIndSession: Tob; //FORMATION
  ChampFormation1, ChampFormation2, ChampFormation3, ChampFormation4, MillesimeFormation: string;
  GblStParamEtat: string;
  //-------------------------------------------------------  { PT75-2 }
  NoPageNum, NbPage: Integer; //PGDUCS
  AncNumPagSys, WMontant, WBaseCotisation: string; //PGDUCS
  WDeclare, WAcompte, WRegul, WAPayer: string; //PGDUCS
  RecaDecl: double; //PGDUCS PT61
  GblEditBulCp: string;
  MaxCum: Integer;
  Tob_ExerciceSocial : Tob;  { PT109 }


  //----------------------------------------------------------------------

{PT121 function ReadParam(var St: string): string;
var
  i: Integer;
begin
  i := Pos(';', St);
  if i <= 0 then i := Length(St) + 1;
  result := Copy(St, 1, i - 1);
  Delete(St, 1, i);
end;}

//----------------------------------------------------------------------

function CalcOLEEtatPG(sf, sp: string): variant;
var
  QRech, Q: TQuery;
  DateDeb, Datefin, hr1: TDateTime;
  NomTable, Salarie, cumul, nature, rubrique: string;
  Val,  TType, Prefixe, Table: string;
  {$IFNDEF EMANAGER}
  Val1, Categorie,DD, DF : String;
  {$ENDIF EMANAGER}
  champ, TypeBase: string;
  i: integer;
  NbJour, NbHeure, Dbl1, Dbl2: double;
  GDCalendrier: THGrid;
  CalOuAbs, st, st1, st2: string;
  Tob_MotifAbs, T_MotifAbs : Tob; { PT103 }
begin
  Debug('Paie PGI : fonction appellée : ' + sf + '  ' + sp);
  result := 'Non affecté';
  {DEB PT43 Initialisation d'une variable de debut de traitement des états
  Afin d'optimiser les tps de traitement des editions}
  if Sf = 'PGINITPARAMETAT' then
  begin
    result := ' ';
    GblStParamEtat := '';
    exit;
  end;

  if (sf = 'PAIECREATEFID') or (Sf = 'PAIECREATECUMSAL') then { PT75-1 }
  begin
    result := '';
    if (GblStParamEtat <> '') and (GblStParamEtat = sp) then exit;
    GblStParamEtat := Sp;
  end;
  {FIN PT43}

  { PT75-1 }
  if Pos('PAIE', Sf) = 1 then
  begin
    result := PGCalcOLEGenericPaie(Sf, Sp);
    Exit;
  end;

  {----------------------------------------------------------------------------------------
                                 FN BULLETIN DE PAIE
  ------------------------------------------------------------------------------------------}
  {Bulletin de paie , pour une session de paie
     @CREATEBULL([XX_DATEDEBUT];[XX_DATEFIN];[XX_SALARIE];[XX_ETABLISSEMENT])
     @RECHNIVEAU()   Pour affichage du net imposable // Max ordre bulletin
     @GetTauxRem([PHB_NATURERUB],[PHB_RUBRIQUE]) //Retourne le mutiplicateur  Taux Coeff
     @GETCUM(CUMULPAIE;[PHB_DATEDEBUT];[PHB_DATEFIN];[PHB_SALARIE];[PHB_ETABLISSEMENT])
        Retourne le montant d'un cumul pour une session,un salarié et un établissement
     @GETRIB(BP;[Aux];[Mode])
        Retourne le rib du salarié en fonction de son code auxiliaire et de son mode de reglement
     @GETLIBBANQUE(BP;[PPU_AUXILIAIRE];[PPU_PGMODEREGLE])
     @GETCUMULEX([CRI_XX_VARIABLECUM])
        Retourne la somme du Cumul du debut de l'exercice social jusqu'à la fin de session de paie éditée
     @GETCUMULMOIS([SO_PGCUMULMOIS1])
     @GETN1(DateDeb,DateFin,CodeSal,Etab)
        Affectation du solde des congés payés periode acquis pris restants
  }

(*PT120
  if (sf = 'CREATEBULL') or (sf = 'RECHNIVEAU') or (sf = 'GETTAUXREM')
    or (sf = 'GETRIB') or (sf = 'GETLIBBANQUE') or (sf = 'GETORGANISME')
    or (sf = 'GETCUM') or (sf = 'GETCUMULEX') or (sf = 'GETCUMULMOIS')
    or (sf = 'GETN1') or (sf = 'GETN') or (sf = 'GETRESTN1') or (sf = 'GETRESTN')
    or (sf = 'GETACQN1') or (sf = 'GETACQN') or (sf = 'GETPRN1') or (sf = 'GETPRN')
    or (sf = 'GETTXTANCIENNETE') then      //PT113
  begin
    result := FunctPGEditBulletin(Sf, Sp);
    Exit;
  end;*)

  { PT75-1 Publication dans CalcOleGenericPaie des fonctions
                - Function  FunctPgEditFicInd
                - Function  FnCreateFid
                - Function  FunctPgEditCumul
                - Function  FnGetReduction & FnGetLesReduction => FunctPgEditReduction
   Suppression de :
               - FnCreateFidRecap (fonctionnement intégré dans FnCreateFid)      }


  {-------------------------------------------------------------------------------
                                 JOURNAL DE PAIE
  --------------------------------------------------------------------------------
  @GETLESCUMUL([PHC_DATEDEBUT];[PHC_DATEFIN];[PSA_SALARIE];[PSA_ETABLISSEMENT])
  }
  if (sf = 'GETLESCUMUL') or (sf = 'GETDIVERS')
    or (sf = 'GETCUMULPJP') or (sf = 'GETCUMUL') or (SF = 'INITCUMULSAL') then
  begin
    result := FunctPgEditJournalPaie(Sf, Sp);
    exit;
  end;

  {----------------------------------------------------------------------------
                      PROVISION POUR CONGES PAYES
  ------------------------------------------------------------------------------}
  { @INITPROVISIONCP([PSA_ETABLISSEMENT];[CRI_DATEARRET];critere;champ rupture;val rupture)
       Init de la tob des congés pour un établissement jusqu'à la date d'arrêt saisie
    @INITPROVISIONSLDCP([PSA_ETABLISSEMENT];[CRI_DATEARRET];critere;champ rupture;val rupture)
       Init de la tob des congés pour un établissement jusqu'à la date d'arrêt saisie
    @GETDATE%  Init des dates des periodes n-1 et n
    @GETMVT Init des mvts pour le salarie concerné
    @%%%PP récupération des mvts pour la periode précédente
    @%%%N1 récupération des mvts pour la periode n-1
    @%%%N récupération des mvts pour la periode en cours}
// PT90 Rajout fct INITPROVDOTCP
  if (sf = 'INITPROVISIONCP') or (sf = 'INITPROVISIONSLDCP') or (sf = 'INITPROVDOTCP')
    or (Sf = 'GETDATEDN1') or (Sf = 'GETDATEFN1') or (Sf = 'GETDATEDN') or (sf = 'GETMVT')
    or (sf = 'GETBASEPP') or (sf = 'GETMOISPP') or (sf = 'GETRESTANTSPP')
    or (sf = 'GETBASEN1') or (sf = 'GETMOISN1') or (sf = 'GETACQUISN1') or (sf = 'GETPRISN1') or (sf = 'GETRESTANTSN1')
    or (sf = 'GETBASEN') or (sf = 'GETMOISN') or (sf = 'GETACQUISN') or (sf = 'GETPRISN') or (sf = 'GETRESTANTSN')
    or (sf = 'GETBASEPRO') or (sf = 'GETMONTANTCP') then
  begin
    Result := FunctPgEditProvisionCP(Sf, Sp);
    Exit;
  end;

  {----------------------------------------------------------------------------
                                   CALENDRIER
  ------------------------------------------------------------------------------}

  //@GETCALENDRIER([PSA_SALARIE]
  //Alimente la grille calendrier pour recup valeur
  if sf = 'GETCALENDRIER' then
  begin
    Salarie := READTOKENST(sp); 
    GblSalarie := Salarie;
    DateDeb := StrToDate(READTOKENST(sp));  
    GblDateDeb := DateDeb;
    DebRepriseMois1 := StrToDate(READTOKENST(sp));
    DebRepriseMois2 := StrToDate(READTOKENST(sp));
    GDCalendrier := THGrid(PGGetControl('GDCALENDRIER', PGCalendrierEtat_Ecran));
    result := '';
    GblDateSortie := idate1900;
    GblDateEntree := idate1900; //PT48-1 ajout date entree
    if (salarie = '') or (DateDeb < 2) then Exit;

    Tob_MotifAbs := tob.create('tob_virtuelle', nil, -1);                         { PT103 }
    Tob_MotifAbs.loaddetaildb('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', nil, False);{ PT103 }

    if (GDCalendrier <> nil) then result := FnCreateCalendrier(Salarie, DateDeb, GDCalendrier,Tob_MotifAbs );
    GDCalendrier := THGrid(PGGetControl('GDCALENDRIER2', PGCalendrierEtat_Ecran));
    if (GDCalendrier <> nil) then result := FnCreateCalendrier(Salarie, PlusMois(DateDeb, 1), GDCalendrier,Tob_MotifAbs);

    {PT32 Modification du fonctionnement de calcul des éléments du tableau récapitulatif
    La requête renvoie les mvts d'abs. anterieur à la periode demandé depuis le date de cumul
    Par la suite on additionnera le decompte effectué à partir des éléments de la grille pour les periodes demandées}
    DateFin := GblDateDeb - 1; //FindeMois(GblDateDeb);   PT32
    //Récapitulatif Tob Absence Mois 1
{Flux optimisé
    Q := OpenSql('SELECT PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_JOURS,PCN_HEURES ' +
      //,SUM(PCN_JOURS) AS JOUR ,SUM(PCN_HEURES) AS HEURE '+   'GROUP BY PCN_TYPECONGE' PT32
      'FROM ABSENCESALARIE ' +
      'WHERE ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-" ) '+
      'OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-")) ' + //PT35 Ajout PCN_SENSABS="-"
      'AND PCN_ETATPOSTPAIE <> "NAN" '+ //PT108
      'AND PCN_SALARIE="' + Salarie + '" ' +
      'AND ((PCN_DATEDEBUTABS>="' + UsDateTime(DebRepriseMois1) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DateFin) + '") ' +
      'OR (PCN_DATEFINABS>="' + UsDateTime(DebRepriseMois1) + '" AND PCN_DATEFINABS<="' + UsDateTime(DateFin) + '" )) ', True);
    Tob_AbsMois1 := Tob.create('Les absencces', nil, -1);
    Tob_AbsMois1.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    Ferme(Q);
}
    st := 'SELECT PCN_TYPECONGE,PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_JOURS,PCN_HEURES ' +
      'FROM ABSENCESALARIE ' +
      'WHERE ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-" ) '+
      'OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-")) ' + //PT35 Ajout PCN_SENSABS="-"
      'AND PCN_ETATPOSTPAIE <> "NAN" '+
      'AND PCN_SALARIE="' + Salarie + '" ' +
      'AND ((PCN_DATEDEBUTABS>="' + UsDateTime(DebRepriseMois1) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DateFin) + '") ' +
      'OR (PCN_DATEFINABS>="' + UsDateTime(DebRepriseMois1) + '" AND PCN_DATEFINABS<="' + UsDateTime(DateFin) + '" )) ';
    Tob_AbsMois1 := Tob.create('Les absencces', nil, -1);
    Tob_AbsMois1.LoadDetailDBFROMSQL('ABSENCESALARIE', St);

    //PT32 Parcours de la tob pour modifier les mouvements à cheval sur la fin du mois anterieur et le debut du premier mois
    //Afin de ne pas les comptabliser 2 fois
    for i := 0 to Tob_AbsMois1.detail.count - 1 do
      if (Tob_AbsMois1.detail[i].GetValue('PCN_DATEDEBUTABS') < GblDateDeb) and (Tob_AbsMois1.detail[i].GetValue('PCN_DATEFINABS') >= GblDateDeb) then
      begin
        { DEB PT103 }
        // CalculDuree(Tob_AbsMois1.detail[i].GetValue('PCN_DATEDEBUTABS'), DateFin, Salarie, GblEtab, '', NbJour, NbHeure);
        if Assigned(Tob_MotifAbs) then
          T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [Tob_AbsMois1.detail[i].GetValue('PCN_TYPECONGE')], False);
        CalculNbJourAbsence(Tob_AbsMois1.detail[i].GetValue('PCN_DATEDEBUTABS'), DateFin, Salarie,
        GblEtab, Tob_AbsMois1.detail[i].GetValue('PCN_TYPECONGE'), T_MotifAbs, NbJour, NbHeure);
        { FIN PT103 }
        Tob_AbsMois1.detail[i].PutValue('PCN_DATEFINABS', Datefin);
        Tob_AbsMois1.detail[i].PutValue('PCN_JOURS', NbJour);
        Tob_AbsMois1.detail[i].PutValue('PCN_HEURES', NbHeure);
      end;
    FreeAndNil(T_MotifAbs);   { PT103 }
    FreeAndNil(Tob_MotifAbs); { PT103 }
    {PT32 Mise en commentaire Le decompte du deuxième mois s'effectue à partir des éléments de la grille
    Utilisation d'une seule date de reprise    puis suppression }
    result := '';
    exit;
  end;

  if sf = 'GETCUMULCALABS' then
  begin
    CalouAbs := READTOKENST(sp);
    NumMois := READTOKENST(sp);
    NbJour := 0;
    NbHeure := 0;
    result := '';
    DateFin := Idate1900;
    if NumMois = '1' then DateFin := FindeMois(GblDateDeb);
    if NumMois = '2' then DateFin := FindeMois(PlusMois(GblDateDeb, 1));
    //Test si Mois > DateSortie
    if (GblDateSortie < DebutdeMois(DateFin)) and (GblDateSortie > Idate1900) then exit; //
    if CalouAbs = 'EFFECTIF' then
    begin
      if NumMois = '1' then
      begin
        {PT48-1 Pris en compte date entrée sortie}
        DateDeb := DebRepriseMois1;
        if (GblDateEntree > idate1900) then
          if (GblDateEntree >= DebRepriseMois1) then
            DateDeb := GblDateEntree;
        if (GblDateSortie > idate1900) then
          if (GblDateSortie <= FindeMois(DateFin)) and (GblDateSortie >= DebutDeMois(DateFin)) then
            DateFin := GblDateSortie;
        { DEB PT103 }
//        CalculDuree(DateDeb, DateFin, GblSalarie, GblEtab, '', NbJour, NbHeure);
        CalculNbJourAbsence(DateDeb, DateFin, GblSalarie, GblEtab,'',nil, NbJour, NbHeure);
        { FIN PT103 }
        {FIN PT48-1}
        EffectifJour1 := NbJour;
        EffectifHeure1 := NbHeure;
      end
      else
        if NumMois = '2' then
      begin
        {PT48-1 Pris en compte date entrée sortie}
        DateDeb := DebRepriseMois2;
        if (GblDateEntree > idate1900) then
          if (GblDateEntree >= DebRepriseMois2) then
            DateDeb := GblDateEntree;
        if (GblDateSortie > idate1900) then
          if (GblDateSortie <= FindeMois(DateFin)) and (GblDateSortie >= DebutDeMois(DateFin)) then
            DateFin := GblDateSortie;
        { DEB PT103 }
        //CalculDuree(DateDeb, DateFin, GblSalarie, GblEtab, '', NbJour, NbHeure);
        CalculNbJourAbsence(DateDeb, DateFin, GblSalarie, GblEtab,'',nil, NbJour, NbHeure);
        { FIN PT103 }
        {FIN PT48-1}
        EffectifJour2 := NbJour;
        EffectifHeure2 := NbHeure;
      end;
    end
    else //PT32 Utilisation de End Else
      if CalouAbs = 'TRAV' then
    begin
      if NumMois = '1' then
      begin
        GDCalendrier := THGrid(PGGetControl('GDCALENDRIER', PGCalendrierEtat_Ecran));
        if IsNumeric(GDCalendrier.CellValues[1, 32]) then Dbl1 := StrToFloat(GDCalendrier.CellValues[1, 32]) else Dbl1 := 0;
        if IsNumeric(GDCalendrier.CellValues[10, 32]) then Dbl2 := StrToFloat(GDCalendrier.CellValues[10, 32]) else Dbl2 := 0;
        NbJour := EffectifJour1 - (Dbl1 - Dbl2);
        NbHeure := EffectifHeure1 - (Dbl1 - Dbl2);
      end
      else
      begin
        GDCalendrier := THGrid(PGGetControl('GDCALENDRIER2', PGCalendrierEtat_Ecran));
        if IsNumeric(GDCalendrier.CellValues[1, 32]) then Dbl1 := StrToFloat(GDCalendrier.CellValues[1, 32]) else Dbl1 := 0;
        if IsNumeric(GDCalendrier.CellValues[10, 32]) then Dbl2 := StrToFloat(GDCalendrier.CellValues[10, 32]) else Dbl2 := 0;
        NbJour := EffectifJour2 - (Dbl1 - Dbl2);
        NbHeure := EffectifHeure2 - (Dbl1 - Dbl2);
      end;
      {       if NumMois='1' then    //PT32 Utilisation de End Else
               Begin
               NbJour:=Tob_AbsMois1.Somme('JOUR',[''],[''],False);
               NbJour:=EffectifJour1-NbJour;
               NbHeure:=Tob_AbsMois1.Somme('HEURE',[''],[''],False);
               NbHeure:=EffectifHeure1-NbHeure;
               End
             else
               if NumMois='1' then
                 Begin
                 NbJour:=Tob_AbsMois2.Somme('JOUR',[''],[''],False);
                 NbJour:=EffectifJour2-NbJour;
                 NbHeure:=Tob_AbsMois2.Somme('HEURE',[''],[''],False);
                 NbHeure:=EffectifHeure2-NbHeure;
                 End;}
    end
    else
      if (CalouAbs <> 'EFFECTIF') and (CalouAbs <> 'TRAV') then
    begin
      NbJour := Tob_AbsMois1.Somme('PCN_JOURS', ['PCN_TYPECONGE'], [CalOuAbs], False);
      NbHeure := Tob_AbsMois1.Somme('PCN_HEURES', ['PCN_TYPECONGE'], [CalOuAbs], False);
      TotNbjour := TotNbjour + NbJour;
      TotNbHeure := TotNbHeure + NbHeure;
      {Tob_temp:=nil;DEB PT32 Utilisation de Tob_AbsMois1.Somme du à la modif de la Requête
      Tob_temp := Tob_AbsMois1.FindFirst(['PCN_TYPECONGE'],[CalouAbs],False);
      if Tob_Temp<>nil then
        Begin
        NbJour:=Tob_temp.GetValue('JOUR'); TotNbjour:=TotNbjour+NbJour;
        NbHeure:=Tob_temp.GetValue('HEURE'); TotNbHeure:=TotNbHeure+NbHeure;
        End;}
      //PT32 integration des absences décomptées dans les grilles;Parcours du premier mois
      GDCalendrier := THGrid(PGGetControl('GDCALENDRIER', PGCalendrierEtat_Ecran));
      if GDCalendrier <> nil then
        for i := 0 to GDCalendrier.ColCount - 1 do
        begin
          if GDCalendrier.Cells[i, 0] = CalouAbs then
          begin
            if IsNumeric(GDCalendrier.CellValues[i, 32]) then Dbl1 := StrToFloat(GDCalendrier.CellValues[i, 32]) else Dbl1 := 0;
            NbJour := NbJour + Dbl1;
            TotNbjour := TotNbjour + NbJour;
            NbHeure := NBHeure + Dbl1;
            TotNbHeure := TotNbHeure + NbHeure;
            Break;
          end;
        end;
      if NumMois = '2' then //PT32 Si 2eme mois Parcours du  deuxième mois
      begin
        GDCalendrier := THGrid(PGGetControl('GDCALENDRIER2', PGCalendrierEtat_Ecran));
        if GDCalendrier <> nil then
          for i := 0 to GDCalendrier.ColCount - 1 do
          begin
            if GDCalendrier.Cells[i, 0] = CalouAbs then
            begin
              if IsNumeric(GDCalendrier.CellValues[i, 32]) then Dbl1 := StrToFloat(GDCalendrier.CellValues[i, 32]) else Dbl1 := 0;
              NbJour := NbJour + Dbl1;
              TotNbjour := TotNbjour + NbJour;
              NbHeure := NbHeure + Dbl1;
              TotNbHeure := TotNbHeure + NbHeure;
              Break;
            end;
          end;
      end; //FIN PT32
    end;
    if GblJourHeure = 'HEU' then result := NbHeure else result := NbJour;
    if Result = 0 then result := '';
    exit;
  end;

  //------------------------------------------------------------------------------

  // GESTION ASSOCIE DES RUBRIQUES   (ALIMENTATION DES CUMULS DES RUBRIQUES)
    //@CREATECUMULRUB(NOMTABLE);
    //Charge tob
  if sf = 'CREATECUMULRUB' then
  begin
    NomTable := Trim(READTOKENST(sp));
    TobEtat := Tob.create('La Tob chargee', nil, -1);
    TOBEtat.LoadDetailDB(NomTable, '', '', nil, FALSE);
    exit;
  end;
  //@GETRUBCUMSENS([PCR_NATURERUB];[PCR_RUBRIQUE])
  if sf = 'GETRUBCUMSENS' then
  begin
// DEV PT92
    if MaxCum = 0 then
    begin
      Q := OpenSql('SELECT COUNT (*) NBRE FROM CUMULPAIE WHERE PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
      if not Q.EOF then
      begin
        MaxCum := Q.FindField('NBRE').asInteger;
        SetLength(TabCumul, MaxCum+1);
        SetLength(TabSens, MaxCum+1);
      end;
      Ferme(Q);
    end;

    Nature := Trim(READTOKENST(sp));
    Rubrique := Trim(READTOKENST(sp));
    for i := 1 to MaxCum do
    begin
      TabCumul[i] := '';
      TabSens[i] := '';
    end;
// FIN PT92
    i := 0;
    if TobEtat <> nil then
      TRubCum := TobEtat.FindFirst(['PCR_NATURERUB', 'PCR_RUBRIQUE'], [Nature, Rubrique], FALSE);
    if TRubCum <> nil then
      while TRubCum <> nil do
      begin
        i := i + 1;
        TabCumul[i] := TRubCum.GetValue('PCR_CUMULPAIE');
        TabSens[i] := TRubCum.GetValue('PCR_SENS');
        TRubCum := TobEtat.FindNext(['PCR_NATURERUB', 'PCR_RUBRIQUE'], [Nature, Rubrique], FALSE);
      end;
    result := '';
    exit;
  end;
  //@GETSENS([XX_NATURERUB];[XX_RUBRIQUE];[XX_CUMULPAIE])
  if sf = 'GETSENS' then
  begin
    Cumul := Trim(READTOKENST(sp));
// DEB PT92
    for i := 1 to MaxCum do
    begin
      if TabCumul[i] = Cumul then
      begin
        result := TabSens[i];
        Break;
      end;
      if i = MaxCum then result := '';
    end;
    exit;
  end;
// FIN PT92

  //------------------------------------------------------------------------------
  //Fonction renvoyant le libelle d'un combo (tablette)
  //@GETLIBELLE([VALEUR];[XX_TYPE];[PREFIXE])
  if sf = 'GETLIBELLE' then
  begin
    Val := Trim(READTOKENST(sp));
    TType := Trim(READTOKENST(sp));
    //Recherche de la tablette concerné
    Q := OpenSql('SELECT DO_COMBO FROM DECOMBOS WHERE DO_TYPE="' + TType + '" ', True);
    Prefixe := Q.FindField('DO_COMBO').asstring;
    Ferme(Q);
    if (Prefixe <> '') and (Val <> '') then Result := RechDom(Prefixe, Val, False)
    else result := '';
    exit;
  end;

  if sf = 'RECHDOM' then
  begin
    Val := Trim(READTOKENST(sp));
    Table := Trim(READTOKENST(sp));
    if (Table <> '') and (Val <> '') then Result := RechDom(Table, val, False) else result := '';
    if Table = 'PGLIBQUALIFICATION' then
    begin
      QRech := OpenSql('SELECT PMI_LIBELLE FROM MINIMUMCONVENT WHERE PMI_NATURE="QUA" ' +
        'AND ##PMI_PREDEFINI## PMI_CODE="' + Val + '"', true);
      if not QRech.eof then //PORTAGECWAS
        result := QRech.findField('PMI_LIBELLE').Asstring;
      Ferme(Qrech);
    end;
    exit;
  end;

  if sf = 'GETLIBVAR' then //VARIABLES DE PAIE
  begin
    TypeBase := Trim(READTOKENST(sp));
    Val := Trim(READTOKENST(sp));
    if TypeBase = '' then
    begin
      Result := '';
      Exit;
    end;
    // Eléments nationaux
    if TYPEBASE = '02' then Table := 'PGELEMENTNAT';
    // Champ de type cotisation
    if ((TYPEBASE = '12') or (TYPEBASE = '13') or (TYPEBASE = '14')) then Table := 'PGCOTISATION';
    // Champ de type base de base PT67  : 03/09/2003 PH V_42 Edition des variables : impression des libellés des bases de cotisation
    if (TYPEBASE = '16') or (TYPEBASE = '17') or (TYPEBASE = '18') or (TYPEBASE = '19') then Table := 'PGBASECOTISATION';
    // Champ de type rémunération
    if ((TYPEBASE = '05') or (TYPEBASE = '06') or (TYPEBASE = '07') or
      (TYPEBASE = '08') or (TYPEBASE = '09') or (TYPEBASE = '10'))
      then Table := 'PGREMUNERATION';
    // Champ de type variable
    if TYPEBASE = '03' then Table := 'PGVARIABLE';
    // Champ de la tanle salarié
    if TYPEBASE = '15' then Table := 'PGCHAMPSALARIE';
    // champ de la table des Ages
    if TYPEBASE = '20' then Table := 'PGTABINTAGE';
    // champ de la table Ancienneté
    if TYPEBASE = '21' then Table := 'PGTABINTANC';
    // champ de la table Minimum conventionnels
    if TYPEBASE = '22' then Table := 'PGTABINTDIV';
    if Table = '' then Result := '';
    if Table <> '' then Result := RechDom(Table, Val, FALSE);
    exit;
  end;

  {----------------------------------------------------------------------------------------
                          Fonstions EDITION DE LA DUCS  -  PGDUCS
  ------------------------------------------------------------------------------------------}
  if (sf = 'SOUSTITRE') or (sf = 'LIBELLEPERIODE') or (sf = 'NBPAGE') or
    //PT8-1   (sf = 'RAZNOPAG') OR (sf = 'PLUSUNEPAG') OR  (sf = 'TYPDECL') OR
  (sf = 'PLUSUNEPAG') or (sf = 'TYPDECL') or
    (sf = 'FORMATCODIF') or (sf = 'FORMATEFFECTIF') or (sf = 'FORMATBASE') or
    (sf = 'FORMATTAUX') or (sf = 'FORMATMONTANT') or (sf = 'FORMATMONNAIE') or
    (sf = 'SIGNEBASE') or (sf = 'SIGNEMT') or (sf = 'SIGNEDECLARE') or
    (sf = 'SIGNEACOMPTE') or (sf = 'SIGNEREGUL') or (sf = 'SIGNEAPAYER') or
    (sf = 'FORMATAPAYER') or (sf = 'FORMATDECLARE') or (sf = 'FORMATDECL') or // PT61
  (sf = 'FORMATACOMPTE') or (sf = 'FORMATREGUL') or
    (sf = 'CESSATION') or (sf = 'CONTINUATION') or
    (sf = 'LIGNEOPTIQUE') or (sf = 'LIGNEOPTIQ2') or
    (sf = 'LIBEFFTOT') or (sf = 'EFFTOT') or // PT6-3
  (sf = 'TIT1') or (sf = 'TIT2') or (sf = 'TIT3') or // PT6-3
  (sf = 'VLUHOMMES') or (sf = 'VLUFEMMES') then // PT34 VLU
  begin
    result := FunctPGEditDucs(Sf, Sp);
    Exit;
  end;
  //-----------------------------------------------------------------------------
  //-----------------------------------------------------------------------------
  //@GETINDEXOFGRILLE([PSA_SALARIE])
  {if sf = 'GETINDEXOFGRILLE' then PT106 Mise en commentaire
  begin
    Salarie := Trim(READTOKENST(sp));
    result := '';
    if PGGrilleRepriseCP <> nil then
      for i := 0 to PGGrilleRepriseCP.RowCount - 1 do
      begin
        if PGGrilleRepriseCP.Cells[0, i] = salarie then Result := IntToStr(i);
        if result <> '' then break;
      end;
    Exit;
  end;   }
  //----------------------DADS-U
  if sf = 'DADSINSTITUTION' then
  begin
    if Sp = '90000' then Result := 'Aucun Organisme' //PT44
    else result := RechDom('PGINSTITUTION', Sp, False);
  end;
  if sf = 'DADSPER' then result := FunctPGDADSU(sf, sp);
  if sf = 'DADSHON1' then result := FunctPGDADSUHonoraires(sf, sp);
  if sf = 'DADSPERRECAP' then result := FunctPGDADSURecap(sf, sp);
  if sf = 'DADSSALAIRES' then result := FunctPGDADSUSalaires(sf, sp); //PT45
  if sf = 'DADSINACT' then result := FunctPGDADSUInactivite(sf, sp); //PT56
  if sf = 'DADSNUMPERIODE' then result := FunctPGDADSUNumPeriode(sf, sp);
  //If sf='DADSNBPER' Then begin result:=FunctPGDADSUNbPer(sf,sp) ;exit;end;
  if sf = 'DADSDATE' then
  begin
    DateDebDADS := StrToDateTime(READTOKENST(sp));
    DateFinDADS := StrToDateTime(READTOKENST(sp));
    result := ''; //PT69-1
  end;
  //----------------------DADS Bilatérale
  //PT64
  if Sf = 'DADSB' then
    result := FunctPGDADSB(sf, sp);
  //FIN PT64
  //-----------------------MEDECINE DU TRAVAIL
  if Sf = 'MEDECINEAGE' then
  begin
    result := FunctMedecineAge(Sf, Sp);
    Exit;
  end;
  if Sf = 'MEDECINESUIV' then
  begin
    Result := FunctMedecineVisiteSuiv(Sf, Sp);
    Exit;
  end;
  if sf = 'VISITEMDT' then
  begin
    Result := FunctVisitesMed(Sf, Sp);
    Exit;
  end;

  //--------------------- MOUVEMENT DE MAIN D'OEUVRE
  if sf = 'SORTIEMMO' then
  begin
    result := FunctSortieMMO(Sf, Sp);
    Exit;
  end;
  if sf = 'ENTREEMMO' then
  begin
    result := FunctEntreeMMO(Sf, Sp);
    Exit;
  end;
  if sf = 'CONTRATMMO' then
  begin
    result := FunctContratMMO(Sf, Sp);
    Exit;
  end;
  if sf = 'MMO' then
  begin
    Result := FunctMMOEtab(Sf, Sp);
    Exit;
  end;
  if Sf = 'MMOMOTIF' then
  begin
    Result := FunctDateMMO(Sf, Sp);
    Exit;
  end;
  //---------------------MSA
  if Sf = 'CREATETOBMSA' then
  begin
    result := CreateTobMSA(Sf, Sp);
    Exit;
  end;
  if Sf = 'MSACALCLUL' then
  begin
    result := MSAHeuresMontant(Sf, Sp);
    Exit;
  end;
  if Sf = 'LIBDATEMSA' then
  begin
    result := LibDateMSA(Sf, Sp);
    Exit;
  end;
  if Sf = 'NUMSSMSA' then Result := FormatCase(READTOKENST(Sp), 'STR', 2);

  //---------------------FORMATION
  if Sf = 'INITTOBFORMATION' then
  begin
    InitTobFormation(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if sF = 'CREATETOBBUDGET' then
  begin
    CreateTobBudgetFor(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if sF = 'CALCOUTBUDBUDGET' then result := CalCoutBudgetFor(Sf, Sp);
  if Sf = 'FREETOBBUDGET' then
  begin
    if TobCoutStageBud <> nil then
    begin
      TobCoutStageBud.free;
      TobCoutStageBud := nil;
    end;
    Result := ''; //PT69-1
  end;
  { PT116
  if Sf = 'TOBFRAISFOR' then
  begin
    InitTobFrais(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'INITLIBFRAISFOR' then
  begin
    InitLibelleFrais(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'FRAISFOR' then result := AfficheFrais(Sf, Sp);
  if Sf = 'LIBFRAISFOR' then result := AfficheLibFrais(Sf, Sp);
  if Sf = 'FREETOBFRAISFOR' then
  begin
    TobFraisFor.Free;
    Result := ''; //PT69-1
  end;
  }
  if Sf = 'INITTOBCOMP' then Result := InitTobComparatif(Sf, Sp);
  if Sf = 'COMPFOR' then Result := FormationComparatif(Sf, Sp);
  if Sf = 'COMPINSC' then Result := InscBudgetComparatif(Sf, Sp);
  if Sf = 'INITTOBFORREALISE' then
  begin
    CreateTobForRealise(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'FORREALISE' then Result := RealiseForMensuel(Sf, Sp);
  if Sf = 'CREATETOBFORGLOBAL' then
  begin
    CreateTobFormationGlobal(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'FORMATIONGLOBAL' then Result := FormationGlobal(Sf, Sp);
  if Sf = 'INITTOBSUIVIRESPF' then
  begin
    CreateTobSuviRespF(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'SUIVIFORRESP' then Result := SuiviResponsableFormation(Sf, Sp);
  //If Sf='NITTOBFORSUIVGESTINDCOMP' Then Result:=CreateTobGestIndFor(Sf,Sp);
  if Sf = 'FORSUIVGESTINDCOMP' then Result := GestIndFormation(Sf, Sp);
  if Sf = 'INITTOBFORSUIVGESTIND' then
  begin
    InitTobSuiviGestInd(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'INITTOBFORSUIVGESTINDCOMP' then
  begin
    InitTobGestIndComp(Sf, Sp);
    Result := ''; //PT69-1
  end;
  if Sf = 'FORSUIVGESTIND' then result := SuiviGestFor(Sf, Sp);
  if Sf = 'INITTOBSUIVIGESTCOLL' then result := InitTobForSuiviGestColl(Sf, Sp);
  if Sf = 'SUIVIGESTCOLL' then result := ForSuiviGestColl(Sf, Sp);
  if Sf = 'PGRENDNBSTAGIAIRE' then result := RendNbStagiaresParAge(Sf, Sp);
  if sf = 'FREETOBSUIVIGESTCOLL' then
  begin
    if TobGestIndSession <> nil then
    begin
      TobGestIndSession.free;
      TobGestIndSession := nil;
    end;
    if TobGestIndRealise <> nil then
    begin
      TobGestIndRealise.free;
      TobGestIndRealise := nil;
    end;
    Result := ''; //PT69-1
  end;
  if (Sf = 'FREETOBFORSUIVGESTINDCOMP') or (Sf = 'FREETOBFORGLOBAL') then
  begin
    if TobGestIndSession <> nil then
    begin
      TobGestIndSession.free;
      TobGestIndSession := nil;
    end;
    if TobGestIndRealise <> nil then
    begin
      TobGestIndRealise.free;
      TobGestIndRealise := nil;
    end;
    if TobGestIndBudget <> nil then
    begin
      TobGestIndBudget.free;
      TobGestIndBudget := nil;
    end;
    if TobGestIndStage <> nil then
    begin
      TobGestIndStage.free;
      TobGestIndStage := nil;
    end;
    Result := ''; //PT69-1
  end;
  if Sf = 'FREETOBFORSUIVGESTIND' then
  begin
    if TobGestIndSession <> nil then
    begin
      TobGestIndSession.free;
      TobGestIndSession := nil;
    end;
    if TobGestIndRealise <> nil then
    begin
      TobGestIndRealise.free;
      TobGestIndRealise := nil;
    end;
    Result := ''; //PT69-1
  end;
  if Sf = 'FREETOBFORSUIVGESTINDCOMP' then
  begin
    if TobGestIndRealise <> nil then
    begin
      TobGestIndRealise.free;
      TobGestIndRealise := nil;
    end;
    if TobGestIndBudget <> nil then
    begin
      TobGestIndBudget.free;
      TobGestIndBudget := nil;
    end;
    if TobGestIndStage <> nil then
    begin
      TobGestIndStage.free;
      TobGestIndStage := nil;
    end;
    if TobGestIndSession <> nil then
    begin
      TobGestIndSession.free;
      TobGestIndSession := nil;
    end;
    Result := ''; //PT69-1
  end;
  if (Sf = 'FREETOBFORGLOBAL') or (Sf = 'FREETOBFORREALISE') or (Sf = 'FREETOBSUIVIRESPF') then
  begin
    if TobFGFrais <> nil then
    begin
      TobFGFrais.free;
      TobFGFrais := nil;
    end;
    if TobFGInsc <> nil then
    begin
      TobFGInsc.free;
      TobFGInsc := nil;
    end;
    if TobFGStage <> nil then
    begin
      TobFGStage.free;
      TobFGStage := nil;
    end;
    if TobFGSessions <> nil then
    begin
      TobFGSessions.free;
      TobFGSessions := nil;
    end;
    if TobFGFormations <> nil then
    begin
      TobFGFormations.free;
      TobFGFormations := nil;
    end;
    Result := ''; //PT69-1
  end;
  if Sf = 'FREETOBCOMPFOR' then
  begin
    if TobForInscritsNom <> nil then
    begin
      TobForInscritsNom.free;
      TobForInscritsNom := nil;
    end;
    if TobForInscritsNonNom <> nil then
    begin
      TobForInscritsNonNom.free;
      TobForInscritsNonNom := nil;
    end;
    if TobForRealise <> nil then
    begin
      TobForRealise.free;
      TobForRealise := nil;
    end;
    if TobForRealiseInsc <> nil then
    begin
      TobForRealiseInsc.free;
      TobForRealiseInsc := nil;
    end;
    Result := ''; //PT69-1
  end;
  // Congés spectacles
  if Sf = 'CONGESSPEC' then result := CongesSpectacles(Sf, Sp);
  if Sf = 'MAJCONGESSPEC' then result := MajTableCongesSpectacles(Sf, Sp);

  {----------------------------------------------------------------------------------------
                          Fonctions EDITION DES CHEQUES
  ------------------------------------------------------------------------------------------}
  if sf = 'LIBCHEQUE' then
  begin
    TestCheque := False;
    NumPosition := 65;
    MontantEuro := (READTOKENST(sp));
    if MontantEuro <> '' then Position := MontantEuro[NumPosition];
    if position = '' then TestCheque := True;
    if Position = ' ' then TestCheque := True;
    while TestCheque = False do
    begin
      NumPosition := NumPosition - 1;
      Position := MontantEuro[NumPosition];
      if position = '' then TestCheque := True;
      if Position = ' ' then TestCheque := True;
    end;
    Result := NumPosition;
    exit;
  end;
  //@GETRUBVENTIL([V_COMPTE])
  if Sf = 'GETRUBVENTIL' then
  begin
    result := READTOKENST(sp);
    result := READTOKENST(sp);
    exit;
  end;

  { PT75-1 Publication dans CalcOleGenericPaie des fonctions
                - Function  FunctPgEditChSocial
   Suppression de :
               - Sf=GETNBSEXE (fonctionnement intégré dans FunctPgEditChSocial }

  //-----------------------------------------------------------------------------
  //-----------------------------------------------------------------------------
  //-----------------------------------------------------------------------------
  if sf = 'FREEFN' then
  begin
    Debug('Paie PGI : Déallocation memoire des données');
    Champ := Trim(READTOKENST(sp));
    if champ = 'PCG' then
    begin
      if TMvtSal <> nil then
      begin
        TMvtSal.free;
        TMvtSal := nil;
      end;
      if TobProvisionCp <> nil then
      begin
        TobProvisionCp.free;
        TobProvisionCp := nil;
      end;
    end;
    if Champ = 'PCU' then
    begin
      if Tob_TypeCumul <> nil then
      begin
        Tob_TypeCumul.Free;
        Tob_TypeCumul := nil;
      end; //PT20
      if TCumPaie <> nil then
      begin
        TCumPaie.Free;
        TCumPaie := nil;
      end;
    end;
    { PT75-1 Publication dans CalcOleGenericPaie des free }
    if Champ = 'PBP' then
    begin
      if TCum <> nil then
      begin
        TCum.free;
        Tcum := nil;
      end;
      if Tob_TauxBull <> nil then
      begin
        Tob_TauxBull.free;
        Tob_TauxBull := nil;
      end;
      if TobPaie <> nil then
      begin
        TobPaie.Free;
        TobPaie := nil;
      end;
//PT122    if sp = 'X' then  { DEB PT109 }
//PT122      Begin
      If Assigned(Tob_DateClotureCp) then FreeAndNil(Tob_DateClotureCp);
      If Assigned(Tob_ExerciceSocial) then FreeAndNil(Tob_ExerciceSocial);
      If Assigned(Tob_OrganismeBull) then FreeAndNil(Tob_OrganismeBull);
//PT122      end;         { FIN PT109 }
    end;
    if Champ = 'PGR' then
    begin
      if TobEtat <> nil then
      begin
        TobEtat.free;
        TObEtat := nil;
      end;
      if TRech <> nil then
      begin
        TRech.free;
        TRech := nil;
      end;
      if TRubCum <> nil then
      begin
        TRubCum.free;
        TRubCum := nil;
      end;
    end;
    if Champ = 'PJP' then
    begin
      if TRech <> nil then
      begin
        TRech.free;
        TRech := nil;
      end;
      if Tob_JournalPaie <> nil then
      begin
        Tob_JournalPaie.free;
        Tob_JournalPaie := nil;
      end;
    end;
    result := '';
    exit;
  end;
  if ((sf = 'FREEDADSPER') or (sf = 'freedadsper')) then //PT69-2 PT73
  begin //DEB PT21-2
    //   if TobNbPer<>nil then Begin TobNbPer.Free;  TobNbPer:=Nil; End;
    if TobDadsNumPer <> nil then
      FreeAndNil(TobDadsNumPer);

    if TOBDADSPER <> nil then
      FreeAndNil(TOBDADSPER);

    //PT45
    if TOBDADSSAL <> nil then
      FreeAndNil(TOBDADSSAL);

    if TOBDADSETAB <> nil then
      FreeAndNil(TOBDADSETAB);
    //FIN PT45

    result := '';
    exit;
  end;
  if ((sf = 'FREEDADSHON') or (sf = 'freedadshon')) then //PT69-2 PT73
  begin
    if TOBDADSHON <> nil then
    begin
      TOBDADSHON.Free;
      TOBDADSHON := nil;
    end;
    result := ''; //FIN PT21-2
    exit;
  end;
  if sf = 'FREETOBMEDECINE' then
  begin
    if TobMedecine <> nil then
    begin
      TobMedecine.Free;
      TobMedecine := nil;
    end;
    Result := ''; //PT69-1
    Exit;
  end;
  if sf = 'FREETOBMMO' then //PT69-2
  begin
    if TobEntreeMMO <> nil then
    begin
      TobEntreeMMO.free;
      TobEntreeMMO := nil;
    end;
    if TobSortieMMO <> nil then
    begin
      TobSortieMMO.Free;
      TobSortieMMO := nil;
    end;
    if TobContratMMO <> nil then
    begin
      TobContratMMO.Free;
      TobContratMMO := nil;
    end;
    if TobMMOEtab <> nil then
    begin
      TobMMOEtab.Free;
      TobMMOEtab := nil;
    end;
    if TobPaieMMO1 <> nil then
    begin
      TobPaieMMO1.Free;
      TobPaieMMO1 := nil;
    end;
    if TobPaieMMO2 <> nil then
    begin
      TobPaieMMO2.Free;
      TobPaieMMO2 := nil;
    end;
    Result := ''; //PT69-1
    Exit;
  end;
  if Sf = 'FREEMSA' then
  begin
    if TobMSAP1 <> nil then
    begin
      TobMSAP1.Free;
      TobMSAP1 := nil;
    end;
    if TobMSAP2 <> nil then
    begin
      TobMSAP2.Free;
      TobMSAP2 := nil;
    end;
    if TobMSAP3 <> nil then
    begin
      TobMSAP3.Free;
      TobMSAP3 := nil;
    end;
    Result := ''; //PT69-1
    exit;
  end;

  //DEB PTXX Bilan Social
  {@GETBSxyz (Val,Val1;Categorie;[CRI_XX_VARIABLEDEB];[CRI_XX_VARIABLEFIN])

  }
{$IFNDEF EABSENCES}
{$IFNDEF EMANAGER}
  if SF = 'GETINITBSANC' then
  begin
    DD := READTOKENST(sp);
    DF := READTOKENST(sp);
    StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), '', '', '');
    InitBSAnciennete(StWhere);
    exit;
  end;
  if SF = 'GETBSFREEANC' then
  begin
    FreeBSAnciennete;
    exit;
  end;

  if Copy(SF, 1, 11) = 'GETBSLIBABS' then
  begin
    val := READTOKENST(sp);
    if val = '1' then result := RechDom('PGMOTIFABSENCE', VH_Paie.PgBSAbs1, False)
    else result := RechDom('PGMOTIFABSENCE', VH_Paie.PgBSAbs2, False);
    exit;
  end;
  if (SF = 'GETBSCPPRIS') or (SF = 'GETBSABS') then
  begin
    DD := READTOKENST(sp);
    DF := READTOKENST(sp);
    Categorie := READTOKENST(sp);
    val := '';
    if SF = 'GETBSCPPRIS' then val := 'PRI'
    else
    begin
      val1 := READTOKENST(sp);
      if val1 = '1' then val := VH_Paie.PgBSAbs1
      else if val1 = '2' then val := VH_Paie.PgBSAbs2
      else if val1 = 'A' then val := 'Autres';
    end;
    result := 0;
    StWhere := '';
    StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), '', '', Categorie);
    result := RendBSAbsences(StrToDate(DD), StrToDate(DF), val, StWhere);
    exit;
  end;
  if (copy(SF, 1, 8) = 'GETBSREM') or (copy(SF, 1, 9) = 'GETBSHEUR') then
  begin
    DD := READTOKENST(sp);
    DF := READTOKENST(sp);
    Categorie := READTOKENST(sp);
    result := 0;
    StWhere := '';
    val := '';
    if copy(SF, 9, 1) = 'B' then Val := '01'
    else if copy(SF, 9, 1) = 'N' then Val := '10'
    else if copy(SF, 1, 9) = 'GETBSHEUR' then Val := '20';
    StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), '', '', Categorie);
    result := RendBSCumul(StrToDate(DD), StrToDate(DF), val, StWhere);
    exit;
  end;
  if copy(SF, 1, 5) = 'GETBS' then
  begin
    Val := READTOKENST(sp);
    Val1 := READTOKENST(sp);
    DD := READTOKENST(sp);
    DF := READTOKENST(sp);
    Categorie := READTOKENST(sp);
    result := 0;
    StWhere := '';

    if Copy(SF, 6, 3) = 'SEX' then StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), 'SEXE', Val, Categorie)
    else if Copy(SF, 6, 3) = 'NAT' then StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), 'NATIONALITE', Val, Categorie)
    else if Copy(SF, 6, 3) = 'EMP' then StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), 'CONDEMPLOI', Val, Categorie);
    if Copy(SF, 6, 3) = 'AGE' then
    begin
      StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), '', '', Categorie);
      StWhere := StWhere + RendBSAge(StrToInt(Val), StrToInt(Val1), StrToDate(DF));
    end;
    if StWhere <> '' then result := ExecBSSQL(StWhere);
    if Copy(SF, 6, 3) = 'ETP' then //PT118 Effectif équivalent temps plein
    begin
      StWhere := RendBSLeWhere(StrToDate(DD), StrToDate(DF), '', '', Categorie);
      result := ExecBSSumTP(StWhere);
    end;
    if Copy(SF, 6, 3) = 'ANC' then
    begin
      result := RendBSAnciennete(StrToInt(Val), StrToInt(Val1), StrToDate(DD), StrToDate(DF), Categorie);
    end;
    exit;
  end;
  //FIN PTXX Bilan Social
{$ENDIF}
{$ENDIF}

  {----------------------------------------------------------------------------
   ----------------------------------------------------------------------------
                               GESTION AFFAIRE
   ----------------------------------------------------------------------------
   ----------------------------------------------------------------------------}
  {FONCTION GESTION AFFAIRE RECUPERER POUR EDITION CALENDRIER HORAIRE }
  if sf = 'RECUPHEURE' then
  begin
    hr1 := StrToFloat(READTOKENST(sp)); // double à convertir
    st1 := FormatFloat('0.00', HeureBase100To60(hr1));
    st2 := ReadTokenPipe(st1, ',');
    if ((st2 = '00') or (st2 = '0')) and (st1 = '00') then
      result := '' else result := st2 + 'h' + st1;
  end;

  //PT102 Calcul l'age à partir d'une date donnée en entrée (Ancienneté)
  if Sf = 'CALCULANCIENNETE' then
  begin
    result := FunctPGCalculAnciennete(Sf, Sp);
    Exit;
  end;

end;

//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------
//----------------------------------------------------------------------------

{----------------------------------------------------------------------------
                            BULLETIN DE PAIE
Initialisation de tobs pour le calcul du bulletin de paie en mode édition (HISTOBULLETIN)
Les Sommes Calculées reprennent les montants des rubriques en fonction
de leur sens (+ ou -) de leur ordre d'affichage.
Seulement les rubriques imprimables sont cumulées
----------------------------------------------------------------------------}
(*PT120
procedure FnCreateBull(DateDeb, DateFin: TDateTime; Salarie, Etab, Page: string);
var
  QOrg, QNiv, QMontEx, QRechCumul, QPrem, QDCloture, QExer, Q: TQuery;
  CpAnnee, CpMois, CpJour, EnAnnee, EnMois, EnJour, RMois: Word;
  DTcloturePi, DTDebutPi, Tdate, CumulDD, DateJamais, DateCP {PT10,DateEntree}: TDateTime;
  i: integer;
  Raz, PclCumul, st, WhereCumMois, StCum, CodeCalcul, EtabOrg: string;
  TOrg, Tob_DateCp: TOB;
  Taux, Coeff, Multiplicateur: double;
begin
  Debug('Paie PGI : Sous fonction appellée : FnCreateBull');
  organisme := '';
  niveau := '';
  cumul21 := '';
  if (Datefin < 1) or (DateDeb < 1) or (Salarie = '') or (Etab = '') then exit;
  //PGLanceBulletin init à true au lancement du bulletin
  //de facon à reproduire le code qu'une seul fois : gain de temps
  //if PGLanceBulletin=True then { PT75-2 }
  // DEB PT86 Rupture si changement de dates ou d'établissement
  // PT87 rajout test si pas de page alors on recharge systématiquement compatibilité état personnalisé en 5.0
//PT122  if (Page = '1') or (Page = '') or (DateDeb <> MemDateDeb) or (DateFin <> MemDateFin) or (Etab <> MemEtab) or (DerPage='X') then //PT121
//PT122  begin
    MemDateDeb := DateDeb; // rupture sur les dates donc rechargement
    MemDateFin := DateFin;
    MemEtab := Etab;
    // FIN PT86
    //init l'organisme associé à l'établissement du salarie
    //DEB PT40
    if Assigned(Tob_OrganismeBull) then FreeAndNil(Tob_OrganismeBull); //PT115
    //DEB PT121
{    QOrg := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_TYPEDITORG,PSA_EDITORG,ETB_EDITORG ' +
      'FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
      'WHERE PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DebutDeMois(DateDeb)) + '" ' + //PT51 Ajout DebutDeMois et FinDeMois
      'AND PPU_DATEFIN<="' + UsDateTime(FinDeMois(DateFin)) + '")', True);
    if not QOrg.eof then
    begin
      Tob_OrganismeBull := Tob.create('ORGANISMEPAIE', nil, -1);
      Tob_OrganismeBull.LoadDetailDB('ORGANISMEPAIE', '', '', QOrg, FALSE);
    end
    else FreeAndNil(Tob_OrganismeBull);
    Ferme(QOrg);  }
    st := 'SELECT PSA_ETABLISSEMENT,PSA_SALARIE,PSA_TYPEDITORG,PSA_EDITORG,ETB_EDITORG ' +
      'FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
      'WHERE PSA_SALARIE IN (SELECT PPU_SALARIE FROM PAIEENCOURS ' +
      'WHERE PPU_DATEFIN>="' + UsDateTime(DebutDeMois(DateDeb)) + '" ' + //PT51 Ajout DebutDeMois et FinDeMois
      'AND PPU_DATEFIN<="' + UsDateTime(FinDeMois(DateFin)) + '")';
    Tob_OrganismeBull := Tob.create('ORGANISMEPAIE', nil, -1);
    Tob_OrganismeBull.LoadDetailDbFromSQL('ORGANISMEPAIE',st) ;
    //FIN PT121
    //FIN PT40

    //init date debut exercice soxial pour l'alimentation des cumuls
    if Assigned(Tob_ExerciceSocial) then FreeAndNil(Tob_ExerciceSocial); //PT115
    //DEB PT121
{    QExer := OpenSql('SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN',True); // PT109
    if not QExer.eof then //PORTAGECWAS
      Begin
      Tob_ExerciceSocial := Tob.Create('Les exercices social',nil,-1); // PT109
      Tob_ExerciceSocial.LoadDetailDB('','','',QExer,False);           // PT109
      End
    else
      FreeAndNil(Tob_ExerciceSocial);                                       // PT109
    Ferme(QExer); }
    st := 'SELECT PEX_DATEDEBUT,PEX_DATEFIN FROM EXERSOCIAL ORDER BY PEX_DATEFIN';
    Tob_ExerciceSocial := Tob.Create('Les exercices social',nil,-1); // PT109
    Tob_ExerciceSocial.LoadDetailDbFromSQL('EXERSOCIAL',st) ;
    //FIN PT121

    //init DateCP pour l'alimentation des cumuls
    if Assigned(Tob_DateClotureCp) then FreeAndNil(Tob_DateClotureCp);  //PT115
    Tob_DateClotureCp := Tob.create('Date de cloture', nil, -1);
    if Tob_DateClotureCp <> nil then
    begin
      QDCloture := OpenSql('SELECT ETB_ETABLISSEMENT,ETB_DATECLOTURECPN FROM ETABCOMPL ', True);
      while not QDCloture.eof do
      begin
        Tob_DateCp := Tob.create('Date de cloture', Tob_DateClotureCp, -1);
        if Tob_DateCp = nil then break;
        Tob_DateCp.AddChampSup('ETABLISSEMENT', False);
        Tob_DateCp.AddChampSup('DATECLOTURE', False);
        TDate := QDCloture.FindField('ETB_DATECLOTURECPN').AsDateTime;
        if (TDate > 0) then
        begin
          RMois := 1;
          RendPeriode(DTcloturePi, DTDebutPi, TDate, DateFin);
          DecodeDate(DateFin, EnAnnee, EnMois, EnJour);
          DecodeDate(DTcloturePi, CpAnnee, Cpmois, Cpjour);
          if EnMois <= CpMois then
          begin
            EnAnnee := EnAnnee - 1;
            RMois := Cpmois + 1;
          end;
          if EnMois > Cpmois then RMois := Cpmois + 1;
          if CpMois = 12 then RMois := 1;
          DateCP := EncodeDate(EnAnnee, RMois, 1);
        end
        else DateCP := idate1900;
        Tob_DateCp.PutValue('ETABLISSEMENT', QDCloture.FindField('ETB_ETABLISSEMENT').AsString);
        Tob_DateCp.PutValue('DATECLOTURE', DateCP);
        QDCloture.next;
      end;
      Ferme(QDCloture);
    end;

    //Cumuls choisis ds paamsoc et à editer
    WhereCumul := '';
    if VH_PAIE.PGCumul01 <> '' then WhereCumul := ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul01 + '" ';
    if VH_PAIE.PGCumul02 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul02 + '" ';
    if VH_PAIE.PGCumul03 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul03 + '" ';
    if VH_PAIE.PGCumul04 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul04 + '" ';
    if VH_PAIE.PGCumul05 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul05 + '" ';
    if VH_PAIE.PGCumul06 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCumul06 + '" ';
    if VH_PAIE.PGCgAcq1 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgAcq1 + '" ';
    if VH_PAIE.PGCgPris1 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgPris1 + '" ';
    if VH_PAIE.PGCgAcq2 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgAcq2 + '" ';
    if VH_PAIE.PGCgPris2 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgPris2 + '" ';
    if VH_PAIE.PGCgAcq3 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgAcq3 + '" ';
    if VH_PAIE.PGCgPris3 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgPris3 + '" ';
    if VH_PAIE.PGCgAcq4 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgAcq4 + '" ';
    if VH_PAIE.PGCgPris4 <> '' then WhereCumul := WhereCumul + ' OR PHC_CUMULPAIE="' + VH_PAIE.PGCgPris4 + '" ';
    //PGLanceBulletin:=False; PT75-2
  end;

  //Emplacement net imposable
  QNiv := OpenSql('SELECT MAX(PHB_ORDREETAT) AS ORDRE FROM HISTOBULLETIN ' +
    'WHERE PHB_DATEDEBUT="' + USDateTime(DateDeb) + '" AND PHB_DATEFIN="' + USDateTime(Datefin) + '" ' +
    'AND PHB_NATURERUB<>"BAS" ' +
    'AND (PHB_MTSALARIAL+PHB_MTPATRONAL+PHB_MTREM<>0 OR PHB_RUBRIQUE like "%.%") ' +
    'AND PHB_IMPRIMABLE="X" AND PHB_ORDREETAT>0 AND PHB_ORDREETAT<7 ' +   { PT95 }
    'AND PHB_SALARIE="' + salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '"', True);
  if not QNiv.eof then //PORTAGECWAS
    niveau := IntToStr(QNiv.FindField('ORDRE').AsInteger)
  else
    niveau := '3';
  Ferme(QNiv);
  //init l'organisme associé à l'établissement du salarie
  //PT40 Recherche de l'organisme par le champ EDITORG
  if Tob_OrganismeBull <> nil then
  begin
    TOrg := Tob_OrganismeBull.FindFirst(['PSA_SALARIE'], [Salarie], False);
    if TOrg <> nil then
    begin
      { DEB PT53 Recomposition du code établissement sur 3 caractères }
      if (PGisOracle) then EtabOrg := Etab { PT105 }
      else
      begin
        EtabOrg := StringOfchar(' ', 3);
        for i := 1 to Length(Etab) do EtabOrg[i] := Etab[i];
      end;
      if (TOrg.GetValue('PSA_TYPEDITORG') = 'ETB') and (TOrg.GetValue('ETB_EDITORG') <> null) then
        organisme := RechDom('PGORGAFFILIATION', EtabOrg + TOrg.GetValue('ETB_EDITORG'), False)
      else
        if (TOrg.GetValue('PSA_TYPEDITORG') = 'PER') and (TOrg.GetValue('PSA_EDITORG') <> null) then
        organisme := RechDom('PGORGAFFILIATION', EtabOrg + TOrg.GetValue('PSA_EDITORG'), False)
      else
        organisme := '';
    end;
    if (organisme = '') or (organisme = 'Error') then
      organisme := RechDom('PGORGAFFILIATION', EtabOrg + '001', False);
    if (organisme = '') or (organisme = 'Error') then                 { PT107 }
      organisme := RechDom('PGORGAFFILIATION', Etab + '001', False);  { PT107 }
    { FIN PT53 }
  end;
  //init heures payées + Cumul du mois en cours
  WhereCumMois := '';
  // DEB PT92
  if (Page = '1') or (Page = '') then
  begin
    MaxCum := 0;
*)
//    Q := OpenSql('SELECT COUNT (*) NBRE FROM CUMULPAIE WHERE PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
(*    if not Q.EOF then
    begin
      MaxCum := Q.FindField('NBRE').asInteger;
      SetLength(TabCumul, MaxCum+1);
      SetLength(TabCumulMois, MaxCum+1);
      SetLength(TabCumulEx, MaxCum+1);
    end;
    Ferme(Q);
  end;
  for i := 1 to MaxCum do
  begin
    if (Page = '1') or (Page = '') then TabCumul[i] := '';
    TabCumulMois[i] := 0;
    TabCumulEx[i] := 0;
  end;
  if (Page = '1') or (Page = '') then
  begin
    Q := OpenSql('SELECT PCL_CUMULPAIE FROM CUMULPAIE WHERE PCL_CUMULPAIE <> "" AND ##PCL_PREDEFINI##', True);
    i := 1;
    while not Q.EOF do
    begin // Boucle de chargement des numéros de cumuls
      TabCumul[i] := Q.FindField('PCL_CUMULPAIE').AsString;
      i := i + 1;
      Q.Next;
    end;
    Ferme(Q);
  end;
  //  for i := 1 to 100 do TabCumulMois[i] := 0; // PT89
// FIN PT92
    // PT26 Réinitialisation des cumuls mois en cours
  if (VH_Paie.PgTypeCumulMois1 = 'CMO') and (VH_Paie.PgCumulMois1 <> '') then WhereCumMois := ' OR PHC_CUMULPAIE="' + VH_Paie.PgCumulMois1 + '" ';
  if (VH_Paie.PgTypeCumulMois2 = 'CMO') and (VH_Paie.PgCumulMois2 <> '') then WhereCumMois := WhereCumMois + ' OR PHC_CUMULPAIE="' + VH_Paie.PgCumulMois2 + '" ';
  if (VH_Paie.PgTypeCumulMois3 = 'CMO') and (VH_Paie.PgCumulMois3 <> '') then WhereCumMois := WhereCumMois + ' OR PHC_CUMULPAIE="' + VH_Paie.PgCumulMois3 + '" ';
  QRechCumul := OpenSql('SELECT PHC_CUMULPAIE,SUM(PHC_MONTANT) AS MONTANT FROM HISTOCUMSAL ' +
    'WHERE PHC_SALARIE="' + Salarie + '" AND PHC_ETABLISSEMENT="' + Etab + '"' +
    'AND PHC_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND (PHC_CUMULPAIE="21" ' + WhereCumMois + ' )' +
    'GROUP BY PHC_CUMULPAIE ', TRUE);
  while not QRechCumul.Eof do
  begin
    StCum := QRechCumul.FindField('PHC_CUMULPAIE').AsString;
    if StCum = '21' then
      Cumul21 := FloatToStr(QRechCumul.FindField('MONTANT').AsFloat);
// DEB PT92
    for i := 1 to MaxCum do
    begin
      if TabCumul[i] = StCum then
      begin
        TabCumulMois[i] := QRechCumul.FindField('MONTANT').AsFloat; // PT89
        Break;
      end;
    end;
// FIN PT92
    QRechCumul.Next;
  end;
  Ferme(QRechCumul);
  //init Tableau les cumuls bas de bulletin
  raz := '';
  CumulDD := DateDeb;
  //  for i := 1 to 100 do TabCumulEx[i] := 0; // PT89 // PT92
    {Periode raz cumul choix :
                          Jamais      : DateJamais
                          Exer Cp     : DateCp
                          Exer social : DateDebutExerSoc}

  QPrem := OpenSQl('SELECT MIN(PHC_DATEDEBUT) AS DATEDEB FROM HISTOCUMSAL ' +
    'WHERE PHC_SALARIE="' + Salarie + '"', True);
  if not QPrem.EOF then //PORTAGECWAS
    DateJamais := QPrem.FindField('DATEDEB').AsDateTime
  else
    DateJamais := idate1900;
  Ferme(QPrem);
  {PT10 mise en commentaire
   Q:=OpenSQl('SELECT MAX(PPU_DATEENTREE) AS DATEENTREE FROM PAIEENCOURS '+
   'WHERE PPU_SALARIE="'+Salarie+'" AND PPU_DATEFIN<="'+UsDateTime(DateFin)+'"',true);
   DateEntree:=Q.FindField('DATEENTREE').AsDateTime;
   Ferme(Q); }

  if Tob_DateClotureCp <> nil then
  begin
    Tob_DateCp := Tob_DateClotureCp.FindFirst(['ETABLISSEMENT'], [etab], False);
    if Tob_DateCp <> nil then
      DateCp := Tob_DateCp.GetValue('DATECLOTURE')
    else
      DateCp := Idate1900;
  end
  else
    DateCp := Idate1900;
  { DEB PT109 }
  DateDebutExerSoc := idate1900;
  if Tob_ExerciceSocial <> nil then
    For i := 0 to Tob_ExerciceSocial.detail.count-1 do
    Begin
      if (Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT') <= DateDeb)
        AND (Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEFIN') >= DateFin)  then
      Begin
        DateDebutExerSoc :=  Tob_ExerciceSocial.detail[i].GetValue('PEX_DATEDEBUT');
        Break;
      End;
    End;
  { FIN PT109 }
  //Alimente les cumuls en fonction de sa periode de remise à zero
  QRechCumul := OpenSql('SELECT DISTINCT PHC_CUMULPAIE,PCL_RAZCUMUL FROM HISTOCUMSAL  ' +
    'LEFT JOIN CUMULPAIE ON PCL_CUMULPAIE=PHC_CUMULPAIE ' +
    'WHERE ##PCL_PREDEFINI## PHC_SALARIE="' + salarie + '" ' +
    'AND (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="07" OR PHC_CUMULPAIE="09" ' +
    'OR PHC_CUMULPAIE="20" OR PHC_CUMULPAIE="30" OR PHC_CUMULPAIE="40" ' + WhereCumul + ' ) ' +
    'ORDER BY PHC_CUMULPAIE', True);

  while not QRechCumul.Eof do
  begin
    PclCumul := QRechCumul.FindField('PHC_CUMULPAIE').AsString;
    //if IsNumeric(PclCumul) then  PT98 Mise en commentaire
    //  if (StrToInt(PclCumul) > 0) and (StrToInt(PclCumul) < 100) then
    Raz := QRechCumul.FindField('PCL_RAZCUMUL').AsString;
    if Raz = '00' then
        begin
        CumulDD := DateDebutExerSoc;
          { PT109 Mise en commentaire } { DEB PT59 Incrémentation d'année pour se positionner sur la date de debut d'exercice coorespondant }
          while PlusDate(CumulDD, 1, 'A') < DateFin do
            CumulDD := PlusDate(CumulDD, 1, 'A');
          while PlusDate(CumulDD, -1, 'A') > DateFin do
            CumulDD := PlusDate(CumulDD, -1, 'A');
          { FIN PT59 }
        end;
     if (raz = '01') or (raz = '02') or (raz = '03') or (raz = '04') or (raz = '05') or (raz = '06') or (raz = '07') or (raz = '08') or (raz = '09') or (raz = '10') or (raz = '11') or (raz = '12') then
        begin
          DecodeDate(Datedeb, CpAnnee, CpMois, CpJour);
          if CpMois < StrtoInt(raz) then Cpannee := Cpannee - 1;
          {PT50 mise en commentaire on doit tenir compte de RAZ dans l'encodedate
          If CpMois>=StrToInt(raz) then CpMois:=StrToInt(raz);
          CumulDD:=EncodeDate(CpAnnee,CpMois,Cpjour);}
          CumulDD := EncodeDate(CpAnnee, StrToInt(raz), 1);
        end;
     if raz = '99' then cumulDD := DateJamais;
     if raz = '20' then cumulDD := DateCP;
     //if DateEntree>cumulDD then CumulDD:=DateEntree; PT10 mise en commentaire
     if (Datefin > 0) and (CumulDD > 0) and (Salarie <> '') and (PclCumul <> '') then
        begin
          st := 'SELECT SUM(PHC_MONTANT) AS MONTANT,PHC_CUMULPAIE ' +
            'FROM HISTOCUMSAL WHERE PHC_SALARIE="' + Salarie + '" AND ' +
            'PHC_CUMULPAIE="' + PclCumul + '" ' +
            'AND PHC_DATEDEBUT>="' + UsDateTime(CumulDD) + '" ' +
            'AND PHC_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
            'GROUP BY PHC_CUMULPAIE';
          QMontEx := OpenSql(st, TRUE);
          if not QMontEx.eof then //PORTAGECWAS
          begin
// DEB PT92
            for i := 1 to MaxCum do
            begin
              if TabCumul[i] = PclCumul then
              begin
                TabCumulEx[i] := QMontEx.FindField('MONTANT').AsFloat; // PT89
                Break;
              end;
            end;
            //         TabCumulEx[StrToInt(PclCumul)] := QMontEx.FindField('MONTANT').AsFloat; // PT89
// FIN PT92
          end;
          Ferme(QMontEx);
        end;
     QRechCumul.Next;
  end;
  ferme(QRechCumul);
  //Init de la zone taux rem    affectation du multiplicateur
{Flux optimisé
  Q := OpenSql('SELECT PHB_NATURERUB,PHB_RUBRIQUE,PHB_TAUXREM,PHB_COEFFREM,PRM_CODECALCUL ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
    'WHERE PHB_SALARIE="' + Salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '"' +
    'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND PHB_NATURERUB="AAA" AND PHB_IMPRIMABLE="X" AND PRM_CODECALCUL<>"01" ' +
    'AND (PHB_TAUXREMIMPRIM="X" OR PHB_COEFFREMIMPRIM="X") ' +
    'ORDER BY PHB_ORDREETAT,PHB_RUBRIQUE ', True);
  Tob_TauxBull := Tob.create('Le taux calculé', nil, -1);
  Tob_TauxBull.LoadDetailDB('HISTO_BULLETIN', '', '', Q, FALSE);
  Ferme(Q);
}
  st := 'SELECT PHB_NATURERUB,PHB_RUBRIQUE,PHB_TAUXREM,PHB_COEFFREM,PRM_CODECALCUL ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN REMUNERATION ON PHB_NATURERUB=PRM_NATURERUB AND ##PRM_PREDEFINI## PHB_RUBRIQUE=PRM_RUBRIQUE ' +
    'WHERE PHB_SALARIE="' + Salarie + '" AND PHB_ETABLISSEMENT="' + Etab + '"' +
    'AND PHB_DATEDEBUT>="' + UsDateTime(DateDeb) + '" AND PHB_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'AND PHB_NATURERUB="AAA" AND PHB_IMPRIMABLE="X" AND PRM_CODECALCUL<>"01" ' +
    'AND (PHB_TAUXREMIMPRIM="X" OR PHB_COEFFREMIMPRIM="X") ' +
    'ORDER BY PHB_ORDREETAT,PHB_RUBRIQUE ';

  Tob_TauxBull := Tob.create('Le taux calculé', nil, -1);
  Tob_TauxBull.LoadDetailDBFROMSQL ('HISTO_BULLETIN',St);

  for i := 0 to Tob_TauxBull.detail.count - 1 do
  begin
    Tob_TauxBull.detail[i].AddChampSup('MULTIPLICATEUR', False);
    CodeCalcul := Tob_TauxBull.detail[i].GetValue('PRM_CODECALCUL');
    Taux := Tob_TauxBull.detail[i].GetValue('PHB_TAUXREM');
    Coeff := Tob_TauxBull.detail[i].GetValue('PHB_COEFFREM');
    Multiplicateur := 0;
    { PT58 Utilisation de Else }
    if CodeCalcul = '02' then Multiplicateur := Taux * Coeff
    else if CodeCalcul = '03' then Multiplicateur := Taux / 100 * Coeff
    else if CodeCalcul = '04' then Multiplicateur := Taux
    else if CodeCalcul = '05' then Multiplicateur := Taux {PT-5 : Taux/100 }
    else if CodeCalcul = '08' then Multiplicateur := Coeff;
    if Coeff <> 0 then { PT65 Test Coeff <> 0 pour pouvoir diviser }
    begin
      if CodeCalcul = '06' then Multiplicateur := Taux / Coeff
      else if CodeCalcul = '07' then Multiplicateur := Taux / 100 / Coeff;
    end;
    Tob_TauxBull.detail[i].PutValue('MULTIPLICATEUR', Multiplicateur);
  end;
end; *)

(*PT120
function FunctPGEditBulletin(Sf, Sp: string): variant;
var
  DateDeb, Datefin : TDateTime;
  DateAncte : TDateTime;     //PT113
  Nb, j, NbA, NbM : integer; //PT113
  NbAnneeMois : Double;      //PT113
  Auxiliaire, Salarie, Etab: string;
  cumul, nature, rubrique, champ, mode, ordre, champs, TypeCumulMois, Val, Page: string;
  Tob_Temp: Tob;
  QRech: TQuery;
  i: integer;
  CMois, CJour, aad, aaf: Word;
begin
  Debug('Paie PGI : Sous fonction appellée : FunctPGEditBulletin ' + Sf + '  ' + Sp);
  result := '';
  //Bulletin de paie , pour une session de paie
  //@CREATEBULL([XX_DATEDEBUT];[XX_DATEFIN];[XX_SALARIE];[XX_ETABLISSEMENT])
  if sf = 'CREATEBULL' then
  begin
    Datedeb := StrToInt(ReadParam(sp));
    CMDD := Datedeb;
    Datefin := StrToInt(ReadParam(sp));
    CMDF := Datefin;
    Salarie := Trim(ReadParam(sp));
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then Salarie := ColleZeroDevant(StrToInt(Salarie), 10);
    CMSal := Salarie;
    Etab := ReadParam(sp);
    CMEtab := Etab;
    Page := ReadParam(sp); { PT75-2 }
    if (Etab <> '') and (Salarie <> '') and (DateDeb > 0) and (DateFin > 0) then
      FnCreateBull(DateDeb, DateFin, Salarie, Etab, Page);
    result := '';
    exit;
  end;

  // @RECHNIVEAU()   Pour affichage du net imposable // Max ordre bulletin
  if sf = 'RECHNIVEAU' then
  begin
{    //if (StrToInt(niveau) < 6) and (niveau <> '') then result := niveau; PT95
    if StrToInt(niveau) = 6 then result := '5';}
    if niveau = '' then result := '4'    { PT95 }
    else result := StrToInt(niveau); 
    exit;
  end;
  //Retourne le mutiplicateur  Taux Coeff
  //@GetTauxRem([PHB_NATURERUB],[PHB_RUBRIQUE])
  if sf = 'GETTAUXREM' then
  begin
    result := '';
    if Tob_TauxBull = nil then exit;
    Nature := ReadParam(sp);
    rubrique := ReadParam(sp);
    if nature <> 'AAA' then exit;
    Tob_Temp := Tob_TauxBull.FindFirst(['PHB_RUBRIQUE'], [rubrique], False);
    if Tob_Temp <> nil then result := Tob_Temp.Getvalue('MULTIPLICATEUR');
    exit;
  end;

  //@GETCUM(CUMULPAIE;[PHB_DATEDEBUT];[PHB_DATEFIN];[PHB_SALARIE];[PHB_ETABLISSEMENT])
  //Bulletin de paie  ,
  //Retourne le montant d'un cumul pour une session,un salarié et un établissement
  if sf = 'GETCUM' then
  begin
    result := Cumul21;
    exit;
  end;

  {@GETRIB(BP;[Aux];[Mode])
  Retourne le rib du salarié en fonction de son code auxiliaire et de son mode de
  reglement
  BP : pour l'utilisation sur l'édition du bulletin de Paie
  }
  if sf = 'GETRIB' then
  begin
    champ := ReadParam(sp);
    Auxiliaire := ReadParam(sp);
    mode := ReadParam(sp);
    LibBanque := ''; //PT57 réinitialisation du rib systématique
    if (mode = '008') and (Auxiliaire <> '') and (champ <> '') then
    begin
      if champ <> 'BP' then
      begin
        QRech := OpenSql('SELECT ' + champ + ' FROM RIB WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_PRINCIPAL="X"', TRUE);
        if not QRech.eof then Result := QRech.Fields[0].Asstring; //PORTAGECWAS
        Ferme(QRech);
      end
      else
        if champ = 'BP' then
      begin
        LibBanque := '';
        val := '';
        QRech := OpenSql('SELECT  R_DOMICILIATION,R_ETABBQ,R_GUICHET,R_NUMEROCOMPTE,R_CLERIB ' +
          'FROM RIB WHERE R_AUXILIAIRE="' + Auxiliaire + '" AND R_SALAIRE="X"', TRUE); //PT124
        if not QRech.Eof then //PORTAGECWAS
        begin
          for i := 1 to QRech.FieldCount - 1 do
            val := val + QRech.Fields[i].Asstring + '  ';
          LibBanque := QRech.FindField('R_DOMICILIATION').asstring;
        end;
        Result := Trim(Val);
        Ferme(QRech);
      end;
    end
    else
      Result := '';
    exit;
  end;
  //@GETLIBBANQUE(BP;[PPU_AUXILIAIRE];[PPU_PGMODEREGLE])

  if sf = 'GETLIBBANQUE' then
  begin
    result := LibBanque;
    exit;
  end;
  if sf = 'GETORGANISME' then
  begin
    result := organisme;
    exit;
  end;
  {@GETCUMULEX([CRI_XX_VARIABLECUM])
  Bulletin de paie , Retourne la somme du Cumul du debut de l'exercice social
  jusqu'à la fin de session de paie éditée
  }
  if sf = 'GETCUMULEX' then
  begin
    cumul := Trim(ReadParam(sp));
    if cumul <> '' then
    begin
// DEB PT92
      for i := 1 to MaxCum do
      begin
        if TabCumul[i] = Cumul then
        begin
          if TabCumulEx[i] = 0 then // PT89
            result := ''
          else result := TabCumulEx[i];
          Break;
        end;
      end;
    end
// FIN PT92    
    else
      result := '';
    exit;
  end;

  //@GETCUMULMOIS([SO_PGCUMULMOIS1])
  if sf = 'GETCUMULMOIS' then
  begin //Soit code cumul soit code remuneration
    ordre := Trim(ReadParam(sp));
    cumul := Trim(ReadParam(sp));
    if cumul = '' then
    begin
      result := '';
      exit;
    end;
    if Length(cumul) = 2 then //Cas d'un cumul
    begin
// DEB PT92
      for i := 1 to MaxCum do
      begin
        if TabCumul[i] = Cumul then
        begin
          if TabCumulMois[i] = 0 then // PT89
            result := ''
          else result := TabCumulMois[i];
          Break;
        end;
      end;
// FIN PT92
    end;
    if Length(Cumul) = 4 then //Cas d'une rémunération
    begin
      Champs := '';
      if ordre = '1' then TypeCumulMois := VH_Paie.PgTypeCumulMois1;
      if ordre = '2' then TypeCumulMois := VH_Paie.PgTypeCumulMois2;
      if ordre = '3' then TypeCumulMois := VH_Paie.PgTypeCumulMois3;
      if TypeCumulMois = '05' then Champs := 'PHB_BASEREM';
      if TypeCumulMois = '06' then Champs := 'PHB_TAUXREM';
      if TypeCumulMois = '07' then Champs := 'PHB_COEFFREM';
      if TypeCumulMois = '08' then Champs := 'PHB_MTREM';
      if (Champs <> '') and (Cumul <> '') then
        if (CMEtab <> '') and (CMSAL <> '') and (CMDD > 0) and (CMDF > 0) then
        begin
          QRech := OpenSql('SELECT ' + Champs + ' FROM HISTOBULLETIN ' +
            ' WHERE PHB_NATURERUB="AAA"  ' +
            ' AND PHB_SALARIE="' + CMSal + '" AND PHB_ETABLISSEMENT="' + CMEtab + '" ' +
            ' AND PHB_RUBRIQUE="' + Cumul + '" ' +
            ' AND PHB_DATEDEBUT>="' + UsDateTime(CMDD) + '" ' +
            ' AND PHB_DATEFIN<="' + UsDateTime(CMDF) + '" ', TRUE);
          if not QRech.eof then //PORTAGECWAS
            result := FloatToStr(QRech.FindField(champs).AsFloat);
          Ferme(QRech);
        end
        else
          result := '';
    end;
    exit;
  end;

  // Affectation du solde des congés payés periode acquis pris restants
 // @GETN1(DateDeb,DateFin,CodeSal,Etab)
  if sf = 'GETN1' then
  begin
    Datedeb := StrToDate(ReadParam(sp));
    Datefin := StrToDate(ReadParam(sp));
    Salarie := Trim(ReadParam(sp));
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      Salarie := ColleZeroDevant(StrToInt(Salarie), 10);
    Etab := ReadParam(sp);
    PN := 0;
    AN := 0;
    RN := 0;
    BN := 0;
    PN1 := 0;
    AN1 := 0;
    RN1 := 0;
    BN1 := 0;
    DDP := 0;
    DFP := 0;
    DDPN1 := 0;
    DFPN1 := 0;
    Result := '';
    if VH_PAie.PGCongesPayes then
    begin
      {PT55 Mise en commentaire : intégration procédure RechercheCumulCP
         Q :=Opensql('SELECT ETB_CONGESPAYES FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Etab+'"',True);
          if not Q.eof then
            if (Q.FindField('ETB_CONGESPAYES').AsString) ='X'  then
              Begin}
      RechercheCumulCP(Datedeb, Datefin, Salarie, Etab, PN, AN, RN, BN, PN1, AN1, RN1, BN1, DDP, DFP, DDPN1, DFPN1, GblEditBulCp);
      if (GblEditBulCp = '') or (GblEditBulCp = 'N') then exit; //PT76
      if (DDPN1 > idate1900) and (DFPN1 > iDate1900) then
      begin
        Decodedate(DDPN1, aad, Cmois, CJour);
        Decodedate(DFPN1, aaf, CMois, CJour);
        Result := copy(inttoStr(aad), 3, 2) + '/' + Copy(intToStr(aaf), 3, 2);
      end;
      {        End
            else
              Result:='';
          Ferme(Q);}
    end;
    exit;
  end;

  if sf = 'GETN' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N1') then exit; //PT76
    if (DDP <> 0) and (DFP <> 0) then
    begin
      decodedate(DDP, aad, Cmois, CJour);
      Decodedate(DFP, aaf, CMois, CJour);
      Result := copy(inttoStr(aad), 3, 2) + '/' + Copy(intToStr(aaf), 3, 2);
    end;
    exit;
  end;

  if sf = 'GETACQN1' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N') then exit; //PT76
    if AN1 <> 0 then Result := AN1;
    exit;
  end;
  if sf = 'GETACQN' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N1') then exit; //PT76
    if AN <> 0 then Result := AN;
    exit;
  end;
  if sf = 'GETPRN1' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N') then exit; //PT76
    if PN1 <> 0 then Result := PN1;
    exit;
  end;
  if sf = 'GETPRN' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N1') then exit; //PT76
    if PN <> 0 then Result := PN;
    exit;
  end;
  if sf = 'GETRESTN1' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N') then exit; //PT76
    if RN1 <> 0 then Result := RN1;
    exit;
  end;
  if sf = 'GETRESTN' then
  begin
    Result := '';
    if (GblEditBulCp = '') or (GblEditBulCp = 'N1') then exit; //PT76
    if RN <> 0 then Result := RN;
    exit;
  end;
  //DEB PT113
  if sf = 'GETTXTANCIENNETE' then
  begin
    Result := '';
    DateAncte := StrToDate(ReadParam(sp));
    DateFin   := StrToDate(ReadParam(sp));
    DiffMoisJour (DateAncte, DateFin, Nb, j);
    NbAnneeMois := Nb/12;
    NbA := StrToInt(FloatToStr(Int(NbAnneeMois)));
    NbM := Round(Frac(NbAnneeMois)*12);
    if (NbA = 0) then
    begin
      if (Nb <> 0) then
        Result := 'Soit ' + IntToStr(Nb) + ' mois';
    end
    else
    begin
      Result := 'Soit ' + IntToStr(NbA);
      if (NbA = 1) then
        Result := Result + ' an'
      else
        Result := Result + ' ans';
      if (NbM <> 0) then
         Result := Result + ' et ' + IntToStr(NbM) + ' mois';
    end;
  end;
  //FIN PT113
end;  *)


{ PT75-1 Publication dans CalcOleGenericPaie des fonctions
              - Function  FunctPgEditFicInd
              - Function  FnCreateFid
              - Function  FunctPgEditCumul
              - Function  FnGetReduction & FnGetLesReduction => FunctPgEditReduction
 Suppression de :
             - FnCreateFidRecap (fonctionnement intégré dans FnCreateFid)      }

{----------------------------------------------------------------------------
                            JOURNAL DE PAIE
----------------------------------------------------------------------------}

function FunctPgEditJournalPaie(Sf, Sp: string): variant;
var
  Salarie, Rupt1, RuptSess, DD, DF, Divers, StI: string;
  TTemp: Tob;
begin
  result := '';
  //@GETLESCUMUL([PHC_DATEDEBUT];[PHC_DATEFIN];[PSA_SALARIE];[PSA_ETABLISSEMENT])
  if sf = 'GETLESCUMUL' then
  begin
    StI := READTOKENST(sp);
    Rupt1 := READTOKENST(sp);
    RuptSess := READTOKENST(sp);
    DD := READTOKENST(sp);
    DF := READTOKENST(sp);       
    StWhere := READTOKENST(sp);
    if (StI = 'TITRE') and ((Rupt1 <> '') or (RuptSess <> '')) then exit
    else
      if (StI = 'RUPT1') and (RuptSess <> '') then exit;
    //  Rupt:=READTOKENST(sp); PT17
    //  StWhere:=PGCritere;
    //  if (PgChampRupt<>'') and  (Rupt<>'') then StWhere:=StWhere+' AND '+PgChampRupt+'="'+Rupt+'"';
    if (DD <> '') and (DF <> '') then
      result := FnGetLesCumul(DD, DF, StWhere) //PT17 Retrait du code salarié
    else
      result := '';
    //For i:=1 to 100 do TabCumul[i]:='0'; //PT33 RAZ du tableau des cumuls.. PT58 TabCumul n'est plus utilisé
    GblSalarie := '';
    exit;
  end;
  if sf = 'INITCUMULSAL' then //PT17 Initialisation des valeurs salariés
  begin
    Salarie := Trim(READTOKENST(sp));
    Debug('Paie Pgi : Traitement salarié=' + Salarie);
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      Salarie := ColleZeroDevant(StrToInt(Salarie), 10);
    { DEB PT58 TabCumul n'est plus utilisé
    For i:=1 to 100 do TabCumul[i]:='0'; }
  //  Assujet:=0;DiversPlus:=0;DiversMoins:=0;Tranche1:=0;Tranche2:=0;BaseCsg:=0;BaseCrds:=0; PT63-1 Code déplacé
    if Tob_JournalPaie = nil then
    begin
      result := '';
      exit;
    end;
    TTemp := Tob_JournalPaie.FindFirst(['SALARIE'], [Salarie], False);
    if Assigned(TTemp) then
      with TTemp do
      begin
        Assujet := GetValue('ASSUJET');
        DiversPlus := GetValue('PLUS');
        DiversMoins := GetValue('MOINS');
        Tranche1 := GetValue('TRANCHE1');
        Tranche2 := GetValue('TRANCHE2');
        Tranche3 := GetValue('TRANCHE3'); //PT63-2
        BaseCsg := GetValue('BASECSG');
        BaseCrds := GetValue('BASECRDS');
        {For i:=1 to 100 do   PT58 les cumuls ne sont plus répurérés par fn @
          Begin
          StI:=IntToStr(i);
          if Length(Sti)=1 then Sti:='0'+Sti;
          if TTemp.FieldExists(Sti) then
            if IsNumeric(TTemp.GetValue(Sti)) then
               TabCumul[i]:=FloatToStr(TTemp.GetValue(Sti));
          End;
        FIN PT58 }
        TTemp.free;
      end;
    result := 0; //PT33 Result 0 pour édition bande
  end;

  if sf = 'GETDIVERS' then
  begin
    result := '';
    Divers := Trim(READTOKENST(sp));
    { DEB PT58 Refonte code }
    if (divers = '9') and (Assujet <> 0) then
    begin
      result := Assujet;
      Assujet := 0; //PT63-1
    end
    else
      if (divers = '101') and (DiversPlus <> 0) then
    begin
      result := DiversPlus;
      DiversPlus := 0; //PT63-1
    end
    else
      if (divers = '102') and (DiversMoins <> 0) then
    begin
      result := DiversMoins;
      DiversMoins := 0; //PT63-1
    end
    else
      if (Divers = 'Tranche1') and (Tranche1 <> 0) then //DEB PT36 Ajout clause conversion
    begin
      result := Tranche1;
      Tranche1 := 0; //PT63-1
    end
    else
      if (Divers = 'Tranche2') and (Tranche2 <> 0) then
    begin
      result := Tranche2;
      Tranche2 := 0; //PT63-1
    end
    else { DEB PT63-2 }
      if (Divers = 'Tranche3') and (Tranche3 <> 0) then
    begin
      result := Tranche3;
      Tranche3 := 0;
    end { FIN PT63-2 }
    else
      if (Divers = 'BaseCSG') and (BaseCsg <> 0) then
    begin
      result := BaseCsg;
      BaseCsg := 0; //PT63-1
    end
    else
      if (Divers = 'BaseCRDS') and (BaseCrds <> 0) then
    begin
      result := BaseCrds;
      BaseCrds := 0 //PT63-1
    end;
    { FIN PT58 }
    exit; //FIN PT36
  end;

  (*DEB PT58 GETCUMULPJP n'est plus utilisé
  if sf='GETCUMULPJP' then
    begin
    cumul := Trim(READTOKENST(sp));
    if (cumul<>'')  then
      begin
      ICumul:=StrToInt(Cumul);
      if TabCumul[ICumul]='' then begin result:=''; exit; end;
      if (PgMonnaieInv=True) and (Cumul<>'20') then   //PT36 Le cumul 20 est un cumul heures, on ne le convertit pas
       Result:=StrToFloat(TabCumul[ICumul])*StrToFloat(PgTauxConvert)
      else
       Result:=StrToFloat(TabCumul[ICumul]);
  { PT17 Mise en commentaire Modification des fn d'appel puis supprimer }
      end
    else
      Result:='';
    exit;
    end;
    FIN PT58 *)
  { PT17  Mise en commentaire Modification des fn d'appel }
end;
{Initialisation de tobs pour l'affichage des cumuls (HISTOCUMSAL)
Récupération du cumul pour un salarié, un établissement à une session de paie donnée
PT17 Modification de la fonction de chargement de la tob : On ne charge plus par salarié mais la totalité}

function FnGetLesCumul(DD, DF, StWhere: string): string;
var
  QRech: TQuery;
  Tob_Divers {,Tob_CumulJournal}, T1, TTemp: Tob;
  Salarie, st: string;
begin
  Debug('Paie Pgi : Chargement des données');
  if V_PGI.Confidentiel = '0' then StWhere := StWhere + ' AND PPU_CONFIDENTIEL<="'+IntToStr(V_PGI.NiveauAccesConf)+'"'; //PT43-1 PT99
  StWhere := ConvertPrefixe(StWhere, 'PPU', 'PHC');
  {DEB PT58 Mise en commentaire : chargement cumul via requête principal sur table paieencours
  QRech:=OpenSql('SELECT SUM(PHC_MONTANT) as MONTANT,PHC_SALARIE,PHC_CUMULPAIE FROM HISTOCUMSAL '+
         'WHERE PHC_DATEFIN>="'+UsDateTime(StrToDate(DD))+'" AND PHC_DATEFIN<="'+UsDateTime(StrToDate(DF))+'" '+
         'AND PHC_REPRISE="-" AND (PHC_CUMULPAIE="01" OR PHC_CUMULPAIE="02" '+   //PT31 : ajout Cond.
         'OR PHC_CUMULPAIE="07" OR PHC_CUMULPAIE="08" OR PHC_CUMULPAIE="09" '+
         'OR PHC_CUMULPAIE="10" OR PHC_CUMULPAIE="20" OR PHC_CUMULPAIE="40" )'+
         ' '+StWhere+' '+
         'GROUP BY PHC_SALARIE,PHC_CUMULPAIE',TRUE);
  Tob_CumulJournal:=Tob.Create('Les cumuls',nil,-1);
  Tob_CumulJournal.LoadDetailDB('HISTO_CUMSAL','','',QRech,False);
  Ferme(QRech);
  FIN PT58 }

  StWhere := ConvertPrefixe(StWhere, 'PHC', 'PHB'); //PT17 Ajout de la Tranche 1 et 2 et de la base csg et crds Modif. requête
{Flux optimisé
  QRech := OpenSql('SELECT PHB_SALARIE,PHB_RUBRIQUE,PHB_ORDREETAT,PHB_SENSBUL,PHB_MTREM,PHB_TRANCHE1,PHB_TRANCHE2,PHB_TRANCHE3,' + //PT63-2
    'PHB_BASECOT,PCT_BRUTSS,PCT_BASECSGCRDS,PCT_BASECRDS ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PCT_RUBRIQUE=PHB_RUBRIQUE ' +
    'WHERE ( (PHB_MTREM<>0 AND (PHB_ORDREETAT=6  OR PHB_ORDREETAT=7)) ' + //PT95
    'OR ( PCT_BRUTSS="X" AND (PHB_TRANCHE1<>0 OR PHB_TRANCHE2<>0 OR PHB_TRANCHE3<>0)) ' + //PT63-2
    'OR ( PHB_BASECOT<>0 AND (PCT_BASECSGCRDS="X" OR PCT_BASECRDS="X"))) ' +
    '' + StWhere + ' ' +
    'AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DD)) + '" ' +
    'AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DF)) + '" ' +
    'ORDER BY PHB_SALARIE', TRUE);
  Tob_Divers := Tob.Create('Divers ', nil, -1);
  Tob_Divers.LoadDetailDB('HISTO_CUMSAL', '', '', QRech, False);
  Ferme(QRech);
}
  St := 'SELECT PHB_SALARIE,PHB_RUBRIQUE,PHB_ORDREETAT,PHB_SENSBUL,PHB_MTREM,PHB_TRANCHE1,PHB_TRANCHE2,PHB_TRANCHE3,' + //PT63-2
    'PHB_BASECOT,PCT_BRUTSS,PCT_BASECSGCRDS,PCT_BASECRDS ' +
    'FROM HISTOBULLETIN ' +
    'LEFT JOIN COTISATION ON PHB_NATURERUB=PCT_NATURERUB AND ##PCT_PREDEFINI## PCT_RUBRIQUE=PHB_RUBRIQUE ' +
    'WHERE ( (PHB_MTREM<>0 AND (PHB_ORDREETAT=6  OR PHB_ORDREETAT=7)) ' + //PT95
    'OR ( PCT_BRUTSS="X" AND (PHB_TRANCHE1<>0 OR PHB_TRANCHE2<>0 OR PHB_TRANCHE3<>0)) ' + //PT63-2
    'OR ( PHB_BASECOT<>0 AND (PCT_BASECSGCRDS="X" OR PCT_BASECRDS="X"))) ' +
    '' + StWhere + ' ' +
    'AND PHB_DATEFIN>="' + UsDateTime(StrToDate(DD)) + '" ' +
    'AND PHB_DATEFIN<="' + UsDateTime(StrToDate(DF)) + '" ' +
    'ORDER BY PHB_SALARIE';
  Tob_Divers := Tob.Create('Divers ', nil, -1);
  Tob_Divers.LoadDetailDBFROMSQL('HISTO_CUMSAL', st);

  Tob_JournalPaie := Tob.create('Le journal de paie', nil, -1);
  //Les cumuls
  Debug('Paie Pgi : Traitement des données');
  (* PT58 Mise en commentaire  Tob_CumulJournal n'est plus utilisé
   T1:=Tob_CumulJournal.FindFirst([''],[''],False);
  While T1<>nil Do
    Begin
    TTemp:=Tob_JournalPaie.FindFirst(['SALARIE'],[T1.GetValue('PHC_SALARIE')],False);
    if TTemp=nil then
      Begin
      TTemp:=Tob.create('La fille journal de paie',Tob_JournalPaie,-1);
      if TTemp<>nil then
        Begin
        TTemp.AddChampSup('SALARIE',False); TTemp.PutValue('SALARIE',T1.GetValue('PHC_SALARIE'));
        TTemp.AddChampSup(T1.GetValue('PHC_CUMULPAIE'),True);
        TTemp.AddChampSup('ASSUJET',True);  TTemp.PutValue('ASSUJET',0);
        TTemp.AddChampSup('MOINS',True);    TTemp.PutValue('MOINS',0);
        TTemp.AddChampSup('PLUS',True);     TTemp.PutValue('PLUS',0);
        TTemp.AddChampSup('TRANCHE1',True); TTemp.PutValue('TRANCHE1',0);
        TTemp.AddChampSup('TRANCHE2',True); TTemp.PutValue('TRANCHE2',0);
        TTemp.AddChampSup('BASECSG',True);  TTemp.PutValue('BASECSG',0);
        TTemp.AddChampSup('BASECRDS',True);  TTemp.PutValue('BASECRDS',0);
        End;
      End;
    if TTemp<>nil then
      Begin
      if TTemp.FieldExists(T1.GetValue('PHC_CUMULPAIE')) then
        TTemp.PutValue(T1.GetValue('PHC_CUMULPAIE'),T1.getValue('MONTANT'))
      Else
        Begin
        TTemp.AddChampSup(T1.GetValue('PHC_CUMULPAIE'),True);
        TTemp.PutValue(T1.GetValue('PHC_CUMULPAIE'),T1.getValue('MONTANT'));
        End;
      End;
    T1:=Tob_CumulJournal.FindNext([''],[''],False);
    End;
   PT58 Restructuration de la Tob_Divers *)
  //Les divers
  while Tob_Divers.detail.count > 0 do
  begin
    T1 := Tob_Divers.detail[0];
    Salarie := T1.GetValue('PHB_SALARIE');
    Debug('Traitement salarié : ' + Salarie);
    TTemp := Tob_JournalPaie.FindFirst(['SALARIE'], [Salarie], False);
    if TTemp = nil then
    begin
      TTemp := Tob.create('La fille journal de paie', Tob_JournalPaie, -1);
      if Assigned(TTemp) then
        with TTemp do
        begin
          AddChampSupValeur('SALARIE', Salarie);
          AddChampSupValeur('ASSUJET', 0);
          AddChampSupValeur('MOINS', 0);
          AddChampSupValeur('PLUS', 0);
          AddChampSupValeur('TRANCHE1', 0);
          AddChampSupValeur('TRANCHE2', 0);
          AddChampSupValeur('TRANCHE3', 0); //PT63-2
          AddChampSupValeur('BASECSG', 0);
          AddChampSupValeur('BASECRDS', 0);
        end;
    end;
    if Assigned(TTemp) then
    begin
      if T1.GetValue('PCT_BASECSGCRDS') = 'X' then
        TTemp.PutValue('BASECSG', TTemp.getValue('BASECSG') + T1.getValue('PHB_BASECOT'));
      if T1.GetValue('PCT_BASECRDS') = 'X' then
        TTemp.PutValue('BASECRDS', TTemp.getValue('BASECRDS') + T1.getValue('PHB_BASECOT'));
      if T1.GetValue('PCT_BRUTSS') = 'X' then
      begin
        TTemp.PutValue('TRANCHE1', TTemp.getValue('TRANCHE1') + T1.getValue('PHB_TRANCHE1'));
        TTemp.PutValue('TRANCHE2', TTemp.getValue('TRANCHE2') + T1.getValue('PHB_TRANCHE2'));
        TTemp.PutValue('TRANCHE3', TTemp.getValue('TRANCHE3') + T1.getValue('PHB_TRANCHE3')); //PT63-2
      end;
      if T1.GetValue('PHB_ORDREETAT') = '6' then     { PT95 }
      begin
        if T1.GetValue('PHB_SENSBUL') <> 'M' then
          TTemp.PutValue('ASSUJET', TTemp.getValue('ASSUJET') + T1.getValue('PHB_MTREM'))
        else
          TTemp.PutValue('ASSUJET', TTemp.getValue('ASSUJET') - T1.getValue('PHB_MTREM'));
      end
      else
        if T1.GetValue('PHB_ORDREETAT') = '7' then  { PT95 }
      begin
        if T1.GetValue('PHB_SENSBUL') <> 'M' then
          TTemp.PutValue('PLUS', TTemp.getValue('PLUS') + T1.getValue('PHB_MTREM'))
        else
          TTemp.PutValue('MOINS', TTemp.getValue('MOINS') + T1.getValue('PHB_MTREM'));
      end;
    end;
    T1.free;
  end;
  { FIN PT58 }
  if Tob_Divers <> nil then Tob_Divers.Free;
  // PT58 if Tob_CumulJournal<>nil then Tob_CumulJournal.Free;
  Result := '';
end;



{------------------------------------------------------------------------
                            PROVISION ET SOLDE CP
------------------------------------------------------------------------}

function FunctPgEditProvisionCP(Sf, Sp: string): variant;
var
  Salarie, DateArret: string;
  Critere, StVal: string;
  Dbltemp: double;
begin
  //@INITPROVISIONCP([PSA_ETABLISSEMENT];[CRI_DATEARRET];critere;champ rupture;val rupture)
  //Init de la tob des congés pour un établissement jusqu'à la date d'arrêt saisie
  //@INITPROVISIONSLDCP([PSA_ETABLISSEMENT];[CRI_DATEARRET];critere;champ rupture;val rupture)
  // PT90 rajout fct INITPROVDOTCP
  if (sf = 'INITPROVISIONCP') or (sf = 'INITPROVISIONSLDCP') or (sf = 'INITPROVDOTCP') then
  begin
    //  Etab:=Trim(READTOKENST(sp));   PT43 ne charge plus la tob par établissement
    DateArret := Trim(READTOKENST(sp));
    if (sf = 'INITPROVDOTCP') AND (IsNumeric (DateArret))then
         DateArret := DateToStr (Valeur(DateArret)); // PT96

    {$IFDEF EAGLCLIENT}
    if IsNumeric (DateArret) then DateArret := DateToStr (Valeur(DateArret)); {PT111}
    {$ENDIF}

    Critere := READTOKENST(sp);
    { RuptChamp:=READTOKENST(sp);
     RuptVal:=READTOKENST(sp);
     If (RuptChamp<>'') and (RuptVal<>'') then Critere:=Critere+' AND '+RuptChamp+'="'+RuptVal+'" ';}
     //
    // PT90 rajout fct INITPROVDOTCP
    if (sf = 'INITPROVISIONCP') or (sf = 'INITPROVISIONSLDCP') or (sf = 'INITPROVDOTCP') then { PT84 }
    begin
      //MAJBaseCPAcquis(True,Etab,''); PT39 Mise en commentaire
    // PT90 rajout fct INITPROVDOTCP
      if sf <> 'INITPROVDOTCP' then TobProvisionCp := CalculeProvisionCp({PT43 Etab,}Critere, StrToDate(DateArret), DateDN1, DateFN1, DateDN)
      else FgInitDotProvCp(Critere, StrToDate(DateArret), DateDN1, DateFN1, DateDN);
    end;
    (*  Else PT84 Mise en commentaire
        if (sf='INITPROVISIONSLDCP') then
          TobProvisionCp:=CalculeSoldeCp({Etab,}Critere,StrToDate(DateArret),DateDN1,DateFN1,DateDN);*)
    BasePP := 0;
    MoisPP := 0;
    RestantsPP := 0;
    BaseN1 := 0;
    MoisN1 := 0;
    AcquisN1 := 0;
    PrisN1 := 0;
    RestantsN1 := 0;
    BaseN := 0;
    MoisN := 0;
    AcquisN := 0;
    PrisN := 0;
    RestantsN := 0;
    BasePro := 0;
    result := '';
    exit;
  end;
  //Init des dates des periodes n-1 et n
  if Sf = 'GETDATEDN1' then
  begin
    if DateDN1 > 0 then Result := DateToStr(DateDN1) else result := '';
    exit;
  end;
  if Sf = 'GETDATEFN1' then
  begin
    if DateFN1 > 0 then Result := DateToStr(DateFN1) else result := '';
    exit;
  end;
  if Sf = 'GETDATEDN' then
  begin
    if DateDN > 0 then Result := DateToStr(DateDN) else result := '';
    exit;
  end;

  if sf = 'GETMVT' then //Init des mvts pour le salarie concerné
  begin
    Salarie := Trim(READTOKENST(sp));
    if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
      Salarie := ColleZeroDevant(StrToInt(Salarie), 10);
    if TobProvisionCp <> nil then
      TMvtSal := TobProvisionCp.FindFirst(['SALARIE'], [Salarie], FALSE);
    result := '';
    exit;
  end;

  if sf = 'GETBASEPP' then //récupération des mvts pour la periode précédente
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPBASERELN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPBASERELN1');
      BasePP := BasePP + StrToFloat(TMvtSal.getvalue('CPBASERELN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETMOISPP' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPMOISRELN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPMOISRELN1');
      MoisPP := MoisPP + StrToFloat(TMvtSal.getvalue('CPMOISRELN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;

  if sf = 'GETRESTANTSPP' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPRESTRELN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPRESTRELN1');
      RestantsPP := RestantsPP + StrToFloat(TMvtSal.getvalue('CPRESTRELN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;

  if sf = 'GETBASEN1' then //récupération des mvts pour la periode n-1
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPBASEN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPBASEN1');
      BaseN1 := BaseN1 + StrToFloat(TMvtSal.getvalue('CPBASEN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETMOISN1' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPMOISN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPMOISN1');
      MoisN1 := MoisN1 + StrToFloat(TMvtSal.getvalue('CPMOISN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;

  if sf = 'GETACQUISN1' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPACQUISN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPACQUISN1');
      AcquisN1 := AcquisN1 + StrToFloat(TMvtSal.getvalue('CPACQUISN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETPRISN1' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPPRISN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPPRISN1');
      PrisN1 := PrisN1 + StrToFloat(TMvtSal.getvalue('CPPRISN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETRESTANTSN1' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPRESTN1') = '0' then result := '' else Result := TMvtSal.getvalue('CPRESTN1');
      RestantsN1 := RestantsN1 + StrToFloat(TMvtSal.getvalue('CPRESTN1')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETBASEN' then //récupération des mvts pour la periode en cours
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPBASEN') = '0' then result := '' else Result := TMvtSal.getvalue('CPBASEN');
      Dbltemp := StrToFloat(TMvtSal.getvalue('CPBASEN')); { PT83 }
      BaseN := BaseN + Dbltemp;
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETMOISN' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPMOISN') = '0' then result := '' else Result := TMvtSal.getvalue('CPMOISN');
      MoisN := MoisN + StrToFloat(TMvtSal.getvalue('CPMOISN')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETACQUISN' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPACQUISN') = '0' then result := '' else Result := TMvtSal.getvalue('CPACQUISN');
      AcquisN := AcquisN + StrToFloat(TMvtSal.getvalue('CPACQUISN')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETPRISN' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPPRISN') = '0' then result := '' else Result := TMvtSal.getvalue('CPPRISN');
      PrisN := PrisN + StrToFloat(TMvtSal.getvalue('CPPRISN')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETRESTANTSN' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPRESTN') = '0' then result := '' else Result := TMvtSal.getvalue('CPRESTN');
      RestantsN := RestantsN + StrToFloat(TMvtSal.getvalue('CPRESTN')); { PT83 }
    end
    else
      result := '';
    exit;
  end;
  if sf = 'GETBASEPRO' then
  begin
    if TMvtSal <> nil then
    begin
      if TMvtSal.getvalue('CPPROVISION') = '0' then result := '' else Result := TMvtSal.getvalue('CPPROVISION');
      BasePro := BasePro + StrToFloat(TMvtSal.getvalue('CPPROVISION')); { PT83 }
    end
    else
      result := '';
    exit;
  end;

  if sf = 'GETMONTANTCP' then
  begin
    StVal := READTOKENST(sp);
    result := '';
    //DEB PT49 Aprés le result on raz les totaux établissement
    if (stVal = 'BPP') and (BasePP <> 0) then
    begin
      result := BasePP;
      BasePP := 0;
    end;
    if (stVal = 'MsPP') and (MoisPP <> 0) then
    begin
      result := MoisPP;
      MoisPP := 0;
    end;
    if (stVal = 'RestPP') and (RestantsPP <> 0) then
    begin
      result := RestantsPP;
      RestantsPP := 0;
    end;
    if (stVal = 'BN1') and (BaseN1 <> 0) then
    begin
      result := BaseN1;
      BaseN1 := 0;
    end;
    if (stVal = 'MsN1') and (MoisN1 <> 0) then
    begin
      result := MoisN1;
      MoisN1 := 0;
    end;
    if (stVal = 'AcqN1') and (AcquisN1 <> 0) then
    begin
      result := AcquisN1;
      AcquisN1 := 0;
    end;
    if (stVal = 'PsN1') and (PrisN1 <> 0) then
    begin
      result := PrisN1;
      PrisN1 := 0;
    end;
    if (stVal = 'RestN1') and (RestantsN1 <> 0) then
    begin
      result := RestantsN1;
      RestantsN1 := 0;
    end;
    if (stVal = 'BN') and (BaseN <> 0) then
    begin
      result := BaseN;
      BaseN := 0;
    end;
    if (stVal = 'MsN') and (MoisN <> 0) then
    begin
      result := MoisN;
      MoisN := 0;
    end;
    if (stVal = 'AcqN') and (AcquisN <> 0) then
    begin
      result := AcquisN;
      AcquisN := 0;
    end;
    if (stVal = 'PsN') and (PrisN <> 0) then
    begin
      result := PrisN;
      PrisN := 0;
    end;
    if (stVal = 'RestN') and (RestantsN <> 0) then
    begin
      result := RestantsN;
      RestantsN := 0;
    end;
    if (stVal = 'BPro') and (BasePro <> 0) then
    begin
      result := BasePro;
      BasePro := 0;
    end;
    //FIN PT49
  end;

end;



function CreateCritere(ChampRupt, ValRupt, StWhere: string): string;
var
  j, poschamp, Deuxieme, Fin, lg: integer;
  ExpRupt: string;
begin
  ExpRupt := '';
  Deuxieme := 3;
  Fin := 0;
  if ChampRupt <> '' then
  begin
    PosChamp := Pos('(', StWhere);
    if PosChamp > 0 then Delete(StWhere, PosChamp, 1);
    PosChamp := Pos(')', StWhere);
    if PosChamp > 0 then Delete(StWhere, PosChamp, 1);

    PosChamp := Pos('AND ' + ChampRupt, StWhere);
    if PosChamp > 0 then
    begin
      for j := (PosChamp + 3) to Length(StWhere) do
      begin
        lg := length(StWhere);
        Deuxieme := Deuxieme + 1;
        if (j + 4) = lg then
        begin
          Fin := 3;
          Break;
        end else Fin := 0;
        if (StWhere[j] = ' ') and (StWhere[j + 1] = 'A') and (StWhere[j + 2] = 'N') and (StWhere[j + 3] = 'D') and (StWhere[j + 4] = ' ') then Break;
      end;
      Delete(StWhere, PosChamp, Deuxieme + Fin);
    end;
    Deuxieme := 3;
    PosChamp := Pos('AND ' + ChampRupt, StWhere);
    if PosChamp > 0 then
    begin
      for j := (PosChamp + 3) to Length(StWhere) do
      begin
        Deuxieme := Deuxieme + 1;
        if (j + 4) = length(StWhere) then
        begin
          Fin := 3;
          Break;
        end else Fin := 0;
        if (StWhere[j] = ' ') and (StWhere[j + 1] = 'A') and (StWhere[j + 2] = 'N') and (StWhere[j + 3] = 'D') and (StWhere[j + 4] = ' ') then Break;
      end;
      Delete(StWhere, PosChamp, Deuxieme + Fin);
    end;

    if ValRupt <> '' then ExpRupt := ' AND ' + ChampRupt + '="' + ValRupt + '" ';
    StWhere := StWhere + ExpRupt;
  end;

  Result := Trim(StWhere);
end;

function ConvertPrefixe(StWhere, DePref, APref: string): string;
var
  pospref: integer;
begin
  if StWhere <> '' then
    while Pos(DePref, StWhere) > 0 do
    begin
      pospref := Pos(DePref, StWhere);
      StWhere[(pospref + 1)] := APref[2];
      StWhere[(pospref + 2)] := APref[3];
    end;
  result := Trim(StWhere);
end;


function FnCreateCalendrier(Salarie: string; DateDeb: TDateTime; GDCalendrier: THGrid; Tob_MotifAbs : Tob): string; { PT103 }
var
  Q, QSemaine: TQuery;
  i, j, Ijf, PremierJour, Jour, NbJourMois, IndJr, DOWJour, L, C, R, Col, Lig: integer;
  NbTempsAbs, NbTempsAbsRestant, DecompteTemps, Nb_j, Nb_h: Double; //,NbJourAbs,NbJourAbsRestant
  DureeTotal, Somme, DblAbs, Duree, DureePerso: Double;
  calend, StandCalend, EtabStandCalend, HeureDeb, HeureFin, TypeAbs: string;
  Temp, NbJrTravEtab, Repos1, Repos2, JourHeure: string;
  DebutPerEnCours, DateDebAbsence, DateFinAbsence, DD, DF, DateFin, DateEntree, DateSortie: TDateTime;
  DatePerso, DateDebExer, DateFinExer: TDateTime;
  JourPerso: Integer;
  StTypConges, Etab, TCong, StQuery, Critere: string;
  Tob_Semaine, Tob_Standard, Tob_JourFerie, T_Jour, T_Bis, Tob_Absence, Tob_JourPerso: Tob;
  YY, MM, JJ: Word;
  NotOkJour, AbsSur2Mois, PrisTotal, OkJourTravail, AM, PM, JourRepos: Boolean;
  T_MotifAbs : Tob;                                          { PT103 }
  CalCivil, AvecFerie, DcptOuvre, DcptOuvrable : Boolean;    { PT103 }
begin
  DateFin := FindeMois(DateDeb);
  NbJourMois := StrToInt(Copy(DateToStr(DateFin), 1, 2));
  TabJour[1] := 'D';
  TabJour[2] := 'L';
  TabJour[3] := 'M';
  TabJour[4] := 'M';
  TabJour[5] := 'J';
  TabJour[6] := 'V';
  TabJour[7] := 'S';
  //raz grille a faire
  ReinitGrille(GDCalendrier, DateDeb);
  //Init Des Tobs
  ChargeTobCalendrier(DateDeb, DateFin, Salarie, True, True,True, Tob_Semaine, Tob_Standard, Tob_JourFerie, Tob_absence, Etab, Calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure, DateEntree, DateSortie); { PT93 }
  GBlEtab := Etab;
  GblJourHeure := JourHeure;
  GblDateSortie := DateSortie;
  GblDateEntree := DateEntree;
  if Tob_Semaine = nil then exit; //PT30-2 Ajout ligne
  i := 1; //1 : positionne premier jour du mois
  //Test si dateentrée salarié dans mois
  DD := DateDeb;
  DF := DateFin;
  if (DD < DateEntree) and (DF > DateEntree) then
    repeat
      DD := DD + 1;
      i := i + 1; //incrémente les jours pour débuter le calendrier à la date entree
    until (DD = DateEntree) or (DD > DF);
  //  init du premier jour et du nb de jour du mois demandé
  Premierjour := DayOfWeek(DD);
  if Premierjour = 1 then Jour := 7 else Jour := PremierJour - 1;
  T_Jour := Tob_Semaine.FindFirst(['ACA_JOUR'], [jour], False);
  DureeTotal := 0;
  while (T_Jour <> nil) and (i <= NbJourMois) do
  begin
    if (DateSortie > 2) and (DateSortie <> null) then
      if DD > DateSortie then Break; //test i date > date de sortie
    //PT29 Modif IfJourNonTravaille : intégration des tobs, recup des jours heures calendrier si jour travaillé
    OkJourTravail := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DD, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM);
    if OkJourTravail = False then
    begin
      if JourHeure <> 'HEU' then Duree := nb_j else Duree := Nb_H; //PT29 Recup jour heure calendrier
      DureeTotal := DureeTotal + Duree;
      if Duree <> 0 then GDCalendrier.Cells[2, i] := FloatToStr(Duree);
    end;
    if Jour = 7 then
    begin
      Jour := 1;
      T_Jour := Tob_Semaine.FindFirst(['ACA_JOUR'], [jour], False);
    end
    else
    begin
      Jour := Jour + 1;
      T_Jour := Tob_Semaine.FindNext(['ACA_JOUR'], [jour], False);
    end;
    i := i + 1;
    DD := DD + 1;
  end;
  //Init des jours personnalisés du calendrier spéc. et Standard sur le calendrier civile
  //DEB PT29 Modification fonctionnalité
  while DD <= DateFin do
  begin
  // pt122 T_Bis := nil;
    T_Bis:=Tob.create('calendrier spécifique au salarié',nil,-1);   // pt122
    T_Jour := nil;
    if Tob_Standard <> nil then T_Jour := Tob_Standard.FindFirst(['ACA_DATE'], [DD], False);
    if T_Jour <> nil then T_Bis.Dupliquer(T_Jour, True, True);
    T_Jour := Tob_Semaine.FindFirst(['ACA_DATE'], [DD], False);
    if T_Jour <> nil then T_Bis.Dupliquer(T_Jour, True, True);
    if T_Bis <> nil then
    begin
      nb_j := 0;
      nb_h := 0;
      DatePerso := T_Bis.GetDatetime('ACA_DATE');
    // deb pt122
     // if (T_Bis.GetValue('ACA_HEUREDEB1') <> 0) and (T_Bis.GetValue('ACA_HEUREFIN1') <> 0) then nb_j := 0.5;
     // if (T_Bis.GetValue('ACA_HEUREDEB2') <> 0) and (T_Bis.GetValue('ACA_HEUREFIN2') <> 0) then nb_j := nb_j + 0.5;
     // DureePerso := Arrondi(T_Bis.GetValue('ACA_DUREE'), 2);
     //if JourHeure <> 'HEU' then DureePerso := nb_j; //Edition en heure ou en jour
     // DureeTotal := DureeTotal + DureePerso;
     // JourPerso := StrToInt(Copy(DateToStr(DatePerso), 1, 2));
     // GDCalendrier.Cells[2, JourPerso] := FloatToStr(DureePerso);
     OkJourTravail := IfJourNonTravaille(T_Bis, Tob_Standard, Tob_JourFerie, DD, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM);
     if OkJourTravail = False then
     begin
      if JourHeure <> 'HEU' then DureePerso := nb_j else Dureeperso := Nb_H; //Edition en heure ou en jour
      DureeTotal := DureeTotal + Dureeperso;
      JourPerso := StrToInt(Copy(DateToStr(DatePerso), 1, 2));
      if dureeperso <> 0 then GDCalendrier.Cells[2, JourPerso] := FloatToStr(DureePerso);
     end; // fin pt122
    end;
    DD := DD + 1;
  end;
  //FIN PT29 Modification fonctionnalité
  GDCalendrier.Cells[2, 32] := FloatToStr(DureeTotal);
  //Init tob cumul absence
  Nb_JourCumul := 0;
  Nb_HeureCumul := 0;
  result := InitCumulAbsence(salarie, Etab, DateDeb, DateFin, DateEntree, DateSortie);
  if JourHeure <> 'HEU' then
    GDCalendrier.Cells[2, 33] := FloatToStr(Nb_JourCumul)
  else
    GDCalendrier.Cells[2, 33] := FloatToStr(Nb_HeureCumul);
  //Integre Absence & CP grille colonne
  Col := 3;
  TypeAbs := '';
  DureeTotal := 0;
  for i := 0 to Tob_absence.Detail.Count - 1 do
  begin
    if (TypeAbs <> Tob_Absence.detail[i].Getvalue('PCN_TYPECONGE')) and (TypeAbs <> '') then
    begin
      Col := Col + 1;
      DureeTotal := 0;
      if Col = 11 then Break;
      if GDCalendrier.ColCount <= Col then GDCalendrier.ColCount := Col + 1;
    end;
    TCong := Tob_Absence.detail[i].Getvalue('PCN_TYPECONGE');
    GDCalendrier.Cells[Col, 0] := TCong;
    TypeAbs := Tob_Absence.detail[i].Getvalue('PCN_TYPECONGE');
    if JourHeure <> 'HEU' then NbTempsAbs := Tob_Absence.detail[i].Getvalue('PCN_JOURS')
    else NbTempsAbs := Tob_Absence.detail[i].Getvalue('PCN_HEURES');
    AffectCumulAbsence(GDCalendrier, Col, 33, Tob_Absence.detail[i].Getvalue('PCN_TYPECONGE'), JourHeure);
    DateDebAbsence := Tob_Absence.detail[i].Getvalue('PCN_DATEDEBUTABS');
    DateFinAbsence := Tob_Absence.detail[i].Getvalue('PCN_DATEFINABS');
    GDCalendrier.Cells[Col, 34] := Tob_Absence.detail[i].Getvalue('PMA_PRISTOTAL');
    // jour:=StrToInt(Copy(DateToStr(DateDebAbsence),1,2));
    NbTempsAbsRestant := NbTempsAbs;
    { DEB PT103 }
    if Assigned(Tob_MotifAbs) then
       T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [TypeAbs], False);
    if Assigned(T_MotifAbs) then
     Begin
     CalCivil     := (T_MotifAbs.GetValue('PMA_CALENDCIVIL') = 'X');
     AvecFerie    := (T_MotifAbs.GetValue('PMA_SSJOURFERIE') = '-');
     DcptOuvre    := (T_MotifAbs.GetValue('PMA_OUVRES') = 'X');
     DcptOuvrable := (T_MotifAbs.GetValue('PMA_OUVRABLE') = 'X');
     end;

    { Recherche du premier jour d'absence dans le calendrier }
    if DateDebAbsence < DateDeb then
      repeat
        if  AvecFerie then begin if (IfJourFerie(Tob_JourFerie,DateDebAbsence)) then NotOkJour := True; end
        else if CalCivil then NotOkJour := False
        else if DcptOuvre then begin if (DayOfWeek(DateDebAbsence) in [1,7]) then NotOkJour := True; end
        else if DcptOuvrable then begin if (DayOfWeek(DateDebAbsence) in [1]) then NotOkJour := True; end
        else
        Begin
        //PT29 Modif IfJourNonTravaille : intégration des tobs, recup des jours heures calendrier si jour travaillé
        NotOkJour := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbsence, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, True, True, Nb_j, Nb_h, AM, PM);
        End;
        if NotOkJour = False then
          begin
          if JourHeure = 'HEU' then DecompteTemps := Nb_h else DecompteTemps := Nb_j;
          NbTempsAbsRestant := NbTempsAbsRestant - DecompteTemps;
          end;
        DateDebAbsence := DateDebAbsence + 1;
      until DateDebAbsence = DateDeb; 
    { FIN PT103 }
    if (DateDebAbsence >= DateDeb) and (NbTempsAbsRestant > 0) then
    begin
      jour := StrToInt(Copy(DateToStr(DateDebAbsence), 1, 2));
      repeat           { DEB PT103 }
        NotOkJour := False; nb_h := 0; nb_j := 1;
        if (AvecFerie) AND (IfJourFerie(Tob_JourFerie,DateDebAbsence)) then NotOkJour := True { PT104 }
        else if CalCivil then NotOkJour := False
        else if (DcptOuvre)  AND (DayOfWeek(DateDebAbsence) in [1,7]) then NotOkJour := True { PT104 }
        else if (DcptOuvrable) AND (DayOfWeek(DateDebAbsence) in [1]) then NotOkJour := True  { PT104 }
        else
        Begin          { FIN PT103 }
        //Début de traitement sur calendrier salarié
        //PT29 Modif IfJourNonTravaille : intégration des tobs, recup des jours heures calendrier si jour travaillé
        NotOkJour := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbsence, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, True, True, Nb_j, Nb_h, AM, PM);
        { DEB PT62 Cas d'un jour ouvrable non travaillé dans l'entreprise }
        if (TCong = 'PRI') and (NotOkJour) then
        begin
          OkJourTravail := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateDebAbsence, True, nb_j, nb_h, AM, PM);
          if (OkJourTravail) and (nb_j <> -1) then
          begin
            JourRepos := False;
            IfJourReposHebdo(DayOfWeek(DateDebAbsence), NbJrTravEtab, Repos1, Repos2, JourRepos);
            if not JourRepos then
            begin
              NotOkJour := False;
              nb_h := 0;
              nb_j := 1;
            end;
          end;
        end;
        { FIN PT62 }
        //Fin de traitement sur calendrier salarié
        End;

        if (NotOkJour = False) then
        begin
          {Temp:=GDCalendrier.Cells[2,Jour];  PT29 Recup jour heure calendrier , Arrondi
          if temp<>'' then DecompteTemps:=StrToFloat(Temp) else DecompteTemps:=0;}
          if JourHeure = 'HEU' then DecompteTemps := Nb_h else DecompteTemps := Nb_j;
          if (GDCalendrier.Cells[Col, Jour]) <> '' then DblAbs := StrToFloat(GDCalendrier.Cells[Col, Jour]) else DblAbs := 0;
          if NbTempsAbsRestant <= DecompteTemps then GDCalendrier.Cells[Col, Jour] := FloatToStr(Arrondi(DblAbs + NbTempsAbsRestant, 2));
          if NbTempsAbsRestant > DecompteTemps then GDCalendrier.Cells[Col, Jour] := FloatToStr(DblAbs + DecompteTemps);
          DureeTotal := DureeTotal + StrToFloat(GDCalendrier.Cells[Col, Jour]) - DblAbs;
          NbTempsAbsRestant := Arrondi(NbTempsAbsRestant - DecompteTemps, 2);
          GDCalendrier.Cells[Col, 32] := FloatToStr(DureeTotal);
        end;
        jour := jour + 1;
        DateDebAbsence := DateDebAbsence + 1;
        //PT29 On décompte même aprés la fin d'absence si le restant >0 Retrait cond. :(DateDebAbsence>DateFinAbsence) or
      until (DateDebAbsence > FinDeMois(DateDeb)) or (NbTempsAbsRestant <= 0);
    end;

  end; //End For
  if T_Jour <> nil then T_Jour.free;
  FreeAndnil(T_Bis); // Pt122
  FreeAndNil(Tob_Standard);
  FreeAndNil(Tob_Semaine);
  FreeAndNil(Tob_Absence);
  FreeAndNil(Tob_CumulAbsence);
  FreeAndNil(Tob_JourFerie);

  //Calcul du total général
  Col := 10;
  Lig := 33;
  GDCalendrier.Cells[11, 0] := 'EFF';
  nb_j := 0;
  for j := 1 to Lig do
  begin
    Somme := 0;
    for i := 2 to Col do
    begin
      if (GDCalendrier.Cells[i, j] <> '') and (i = 2) then
        Somme := Somme + Arrondi(StrToFloat(GDCalendrier.Cells[i, j]), 2); //PT29 Arrondi
      if (GDCalendrier.Cells[i, j] <> '') and (i > 2) and (GDCalendrier.Cells[i, 34] = 'X') then
        Somme := Somme - Arrondi(StrToFloat(GDCalendrier.Cells[i, j]), 2); //PT29 Arrondi
    end;
    if Somme > 0 then
      GDCalendrier.Cells[11, j] := FloatToStr(Somme)
    else
    begin
      nb_j := nb_j + somme;
      GDCalendrier.Cells[11, j] := '';
    end;
    //Réajustement des jours ouvrables non travaillés et décompté en CP et non en temps de travail
    if j >= 32 then
    begin
      if nb_j < 0 then nb_j := nb_j * -1;
      GDCalendrier.Cells[11, j] := FloatToStr(Somme + Nb_j);
    end;
  end;
end;



procedure ReinitGrille(GDCalendrier: THGrid; DateDeb: TDateTime);
var
  i, j, C, R: integer;
  DD, DF: TDateTime;
  NbJourMois: integer;
begin
  DD := DateDeb;
  DF := FindeMois(Datedeb);
  NbJourMois := StrToInt(Copy(DateToStr(DF), 1, 2));
  GDCalendrier.ColCount := 12;
  GDCalendrier.FixedCols := 1;
  GDCalendrier.FixedRows := 1;
  GDCalendrier.RowCount := 35;
  for R := 0 to GDCalendrier.RowCount - 1 do
    for C := 0 to GDCalendrier.ColCount - 1 do
    begin
      GDCalendrier.Cells[C, R] := '';
    end;

  j := NbJourMois + 1;
  for i := 1 to J do
  begin
    GDCalendrier.Cells[0, i] := TabJour[DayOfWeek(DD)];
    DD := DD + 1;
    if DD > DF then Break;
  end;
  j := NbJourMois;
  for i := 1 to j do
  begin
    GDCalendrier.Cells[1, i] := IntToStr(i);
    if i > NbJourMois then break;
  end;
  GDCalendrier.Cells[1, 32] := 'TOT';
  GDCalendrier.Cells[0, 33] := 'CUM';
  GDCalendrier.Cells[0, 34] := 'Pris';

  GDCalendrier.Cells[2, 0] := 'HOR';
end;




function InitCumulAbsence(Salarie, Etab: string; DateDeb, DateFin, DateEntree, DateSortie: TDateTime): string;
var
  Q: TQuery;
  DateDebExer: TDateTime;
  St : String;
begin
  DateDebExer := 0;
  result := '';
  if (DateDeb < DateEntree) and (DateFin < DateEntree) then exit;
  if (DateDeb > DateSortie) and (DateFin > DateSortie) and (DateSortie > 2) then exit;

  //Calcul du cumul de la date de debut d'exercice à la date de fin en cours
  Q := OpenSql('SELECT PEX_EXERCICE,PEX_LIBELLE,PEX_DATEDEBUT,PEX_DATEFIN ' +
    'FROM EXERSOCIAL ' +
    'WHERE PEX_DATEDEBUT<="' + USDateTime(DateDeb) + '" ' +
    'AND PEX_DATEFIN>="' + USDateTime(DateFin) + '" ', True);
  if not Q.eof then
  begin
    DateDebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
    Result := Q.FindField('PEX_LIBELLE').AsString;
  end;
  Ferme(Q);

  if (DateDeb <= DateEntree) and (DateFin >= DateEntree) then DateDebExer := DateEntree;

  if (DateDeb <= DateSortie) and (DateFin >= DateSortie) and (DateSortie > 2) then DateFin := DateSortie;

  if (DateDeb > DateEntree) and (DateFin > DateEntree) and (DateDebExer < DateEntree) then DateDebExer := DateEntree;
  { DEB PT103 }
//  CalculDuree(DateDebExer, DateFin, Salarie, Etab, '', Nb_JourCumul, Nb_HeureCumul);
  CalculNbJourAbsence(DateDebExer, DateFin, Salarie, Etab,'',nil, Nb_JourCumul, Nb_HeureCumul);
  { FIN PT103 }
  if (DateDebExer > 2) and (DateFin > 2) then
  begin
{Flux optimisé
    Q := OpenSql('SELECT PCN_SALARIE,PCN_TYPECONGE,SUM(PCN_JOURS) AS JOUR,' +
      'SUM(PCN_HEURES) AS HEURE FROM ABSENCESALARIE ' +
      'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' +   // PT100
      'WHERE PCN_SALARIE="' + Salarie + '" ' +
      'AND (PCN_DATEDEBUTABS>="' + UsDateTime(DateDebExer) + '" ' +
      'AND PCN_DATEDEBUTABS<="' + UsDateTime(DateFin) + '" ' +
      'OR PCN_DATEFINABS>="' + UsDateTime(DateDebExer) + '" ' +
      'AND PCN_DATEFINABS<="' + UsDateTime(DateFin) + '" ) ' +
      'AND ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-" ) '+
      'OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-" )) ' + //PT35 Ajout AND PCN_SENSABS="-"
      'AND PCN_ETATPOSTPAIE <> "NAN" AND PMA_EDITION="X" ' + // PT108
      'GROUP BY PCN_SALARIE,PCN_TYPECONGE ', True);
    Tob_CumulAbsence := Tob.Create('Les cumuls absences', nil, -1);
    Tob_CumulAbsence.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    Ferme(Q);
}
   st := 'SELECT PCN_SALARIE,PCN_TYPECONGE,SUM(PCN_JOURS) AS JOUR,' +
      'SUM(PCN_HEURES) AS HEURE FROM ABSENCESALARIE ' +
      'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PMA_MOTIFABSENCE=PCN_TYPECONGE ' +   // PT100
      'WHERE PCN_SALARIE="' + Salarie + '" ' +
      'AND (PCN_DATEDEBUTABS>="' + UsDateTime(DateDebExer) + '" ' +
      'AND PCN_DATEDEBUTABS<="' + UsDateTime(DateFin) + '" ' +
      'OR PCN_DATEFINABS>="' + UsDateTime(DateDebExer) + '" ' +
      'AND PCN_DATEFINABS<="' + UsDateTime(DateFin) + '" ) ' +
      'AND ((PCN_TYPEMVT="ABS" AND PCN_SENSABS="-" ) '+
      'OR (PCN_TYPECONGE="PRI" AND PCN_TYPEMVT="CPA" AND PCN_MVTDUPLIQUE="-" )) ' + //PT35 Ajout AND PCN_SENSABS="-"
      'AND PCN_ETATPOSTPAIE <> "NAN" AND PMA_EDITION="X" ' + // PT108
      'GROUP BY PCN_SALARIE,PCN_TYPECONGE ';
   Tob_CumulAbsence := Tob.Create('Les cumuls absences', nil, -1);
   Tob_CumulAbsence.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
  end;
end;


procedure AffectCumulAbsence(GDCalendrier: THGrid; C, L: integer; TypeConges, JourHeure: string);
var
  T: Tob;
begin
  if TypeConges = '' then exit;
  if Tob_CumulAbsence = nil then exit;
  T := Tob_CumulAbsence.FindFirst(['PCN_TYPECONGE'], [TypeConges], False);
  if T <> nil then
  begin
    if JourHeure <> 'HEU' then
      GDCalendrier.Cells[C, L] := T.GetValue('JOUR')
    else
      GDCalendrier.Cells[C, L] := T.GetValue('HEURE');
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 18/07/2001
Modifié le ... :   /  /
Description .. : Fonctions liées à l'état DUCS.
Suite ........ : formatage conditionnel des différentes zones de l'état
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

function FunctPGEditDucs(Sf, Sp: string): variant;
var
  QRech: TQuery;
  Codif, Codification, Libelle, StLongEditable, blanc : string;
  Effectif, TypeCotisation, NatureOrg, BaseCotisation : string;
  BoolCalcul, TauxCotisation, Montant, Institution    : string;
  CodifBaseArr, CodifMtArr, BaseArr, MtArr            : string;
  Monnaie, Acompte, Regul, Apayer, Declare            : string;
  Libperiod, Periode, Mois, Trimestre, An, NumPagSys  : string;
  Etab, Organisme, DebPer, FinPer, LigneOptiq2        : string;
  NbSalFpe                                            : string; // PT6-3
  NbHommes, NbFemmes                                  : string; // PT8-2
  // d PT80
  // NbApprentis : integer;
  NbApprentis                                         : string;
  // f PT 80
  NumMois                                             : integer;
  LongEditable, i, NbDetail, NbreMois, NbreJour       : integer;
  NAcompte, NRegul                                    : double;
  QQ                                                  : Tquery;
  StQQ                                                : string;
  Period1, Period2, TypeDecl1, TypeDecl2, PeriodEtat  : string;
  NumDucs                                             : string; // PT25
  DebutST, LongEditST                                 : integer; // PT28
  Atampon                                             : string; // PT28
  Predef                                              : string; // PT101
  AlsaceMoselle                                       : string; // PT112
begin
  Result := '';
  //@SOUSTITRE([POG_NATUREORG];[SYS_PAGE],[POG_POSTOTAL],[POG_LONGTOTAL])
  if sf = 'SOUSTITRE' then
    // mise en forme du sous-titre qui varie selon la nature de la DUCS
  begin
    NatureOrg := READTOKENST(sp);
    Periode := READTOKENST(sp);
    // début PT28
    Atampon := READTOKENST(sp);
    if isnumeric(Atampon) then
      DebutST := StrToInt(Atampon);
    Atampon := READTOKENST(sp);
    if isnumeric(Atampon) then
      LongEditST := StrToInt(Atampon);
    // fin PT28
      // Affectation du sous-titre
    if (NatureOrg = '100') then
      Result := '                    (ART. R243-13 DU CODE DE LA SECURITE SOCIALE)';
    if (NatureOrg = '200') then
      Result := '                           (ART. R351-4 DU CODE DU TRAVAIL)';
    if (NatureOrg = '300') then
      // PT6-1
    begin
      // début PT28
      if (DebutST <> 0) and (LongEditST <> 0) then
        if (Copy(Periode, 3, 2) = '00') then
          Result := '                               RECAPITULATIF ANNUEL '
        else
          // PT6-1 fin
          Result := ''
      else
        // fin PT28
        if (Copy(Periode, 3, 2) = '00') then
        Result := '                      RECAPITULATIF ANNUEL STRANDIRSEE I.R.C.'
      else
        // PT6-1 fin
        Result := '                                STANDARDISEE I.R.C.';
    end;
    if (NatureOrg = '600') then
      Result := '(DECRET N° 1282 DU 29/12/76 MODIFIE, ' +
        'CONVENTION MSA/UNEDIC, REGLEMENT CAMARA DU 10/12/1991)';
    if (NatureOrg = '700') then
      Result := '';
  end;

  // @NBPAGE([PDU_ETABLISSEMENT];[PDU_ORGANISME];[PDU_DATEDEBUT];[PDU_DATEFIN];
  //         [PDU_NUM])
  if sf = 'NBPAGE' then
    // détermination du nombre de pages.
  begin
    RecaDecl := 0.00; //PT61
    Etab := READTOKENST(sp);
    Organisme := READTOKENST(sp);
    DebPer := READTOKENST(sp);
    FinPer := READTOKENST(sp);
    NumDucs := READTOKENST(sp); // PT25
    if (NumDucs = '') then NumDucs := '0'; // DB2
    StQQ := 'PDD_ETABLISSEMENT = "' + Etab + '" AND ' +
      'PDD_ORGANISME = "' + Organisme + '" AND ' +
      'PDD_DATEDEBUT = "' + UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) + '" AND ' +
      'PDD_DATEFIN = "' + UsDateTime(StrToDate(DateToStr(Valeur(FinPer)))) + '" AND ' +
      'PDD_NUM = ' + NumDucs; // PT25      // DB2
    QQ := OpenSQL('SELECT Count(*) FROM DUCSDETAIL WHERE ' + StQQ, True);
    NbDetail := 0; // PORTAGECWAS
    if not QQ.EOF then // PORTAGECWAS
      NbDetail := QQ.Fields[0].AsInteger;
    Ferme(QQ);
    // debut PT14-3
    StQQ := '';
    StQQ := 'PDD_ETABLISSEMENT = "' + Etab + '" AND ' +
      'PDD_ORGANISME = "' + Organisme + '" AND ' +
      'PDD_DATEDEBUT = "' + UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) + '" AND ' +
      'PDD_DATEFIN = "' + UsDateTime(StrToDate(DateToStr(Valeur(FinPer)))) + '" AND ' +
      'PDD_NUM = ' + NumDucs + ' AND ' + // PT25  // DB2
    // d PT88    'PDD_LIBELLESUITE <> ""';
    '(PDD_LIBELLESUITE <> "" OR PDD_TYPECOTISATION="S")';
    // f PT88
    QQ := OpenSQL('SELECT Count(*) FROM DUCSDETAIL WHERE ' + StQQ, True);
    if not QQ.EOF then // PORTAGECWAS
      NbDetail := NbDetail + QQ.Fields[0].AsInteger;
    Ferme(QQ);
    // fin PT14-3
    if (NbDetail = 0) then
    begin
      NbPage := 1;
    end
    else
    begin
      NbPage := NbDetail div 22;
      if (NbPage * 22) < NbDetail then NbPage := NbPage + 1;
    end;
    Result := IntToStr(NbPage);
  end;

  //@RAZNOPAG()       PT8-1
  //IF sf = 'RAZNOPAG' then
  // RAZ du n° de page sur rupture établissement et sur rupture organisme
  // begin
  //  if (AncNumPagSys <>'1') AND (NoPageNum > NbPage) then
  //     NoPageNum := 0 ;
  //  Result := '';
  // end;

  //@PLUSUNEPAG([SYS_PAGE])
  if sf = 'PLUSUNEPAG' then
    // +1 sur n+ de page à chaque en tête
  begin
    NumPagSys := READTOKENST(sp);
    if (NumPagSys <> AncNumPagSys) and (NoPageNum >= NbPage) then
      NoPageNum := 0;
    if (NumPagSys <> AncNumPagSys) then
      NoPageNum := NoPageNum + 1;
    Result := IntToStr(NoPageNum);
    AncNumPagSys := NumPagSys;
  end;
  // @PAGPLUS([PAG]
  if sf = 'PAGPLUS' then
    // +1 Sur N° de page, utilisé pour les tests de dernière page
  begin
    Result := IntToStr((StrToInt(READTOKENST(sp))) + 1);
  end;
  // @TYPDECL([PDU_DATEDEBUT];[PDU_DATEFIN];[POG_PERIODICITDUCS];
  // [POG_AUTREPERIODUCS];[POG_PERIODCALCUL];[POG_AUTPERCALCUL])
  if sf = 'TYPDECL' then
  begin
    DebPer := READTOKENST(sp);
    FinPer := READTOKENST(sp);
    Period1 := READTOKENST(sp);
    Period2 := READTOKENST(sp);
    TypeDecl1 := READTOKENST(sp);
    TypeDecl2 := READTOKENST(sp);
    DiffMoisJour(StrToDate(DateToStr(Valeur(DebPer))), StrToDate(DateToStr(Valeur(FinPer))), NbreMois, NbreJour);
    NbreMois := NbreMois + 1;
    PeriodEtat := 'M';
    if (NbreMois = 3) then PeriodEtat := 'T';
    if (NbreMois = 6) then PeriodEtat := 'S';
    if (NbreMois = 12) then PeriodEtat := 'A';
    if (PeriodEtat = Period1) then
      Result := TypeDecl1
    else
      if (PeriodEtat = Period2) then
      Result := TypeDecl2;
  end;
  //@LIBELLEPERIODE([PDU_ABREGEPERIODE])
  if sf = 'LIBELLEPERIODE' then
    // mise en forme du sous-titre qui varie selon la nature de la DUCS
  begin
    Periode := READTOKENST(sp);
    libperiod := '';
    An := copy(Periode, 1, 2);
    Mois := copy(Periode, 4, 1);
    Trimestre := copy(Periode, 3, 1);
    if (Mois <> '0') and (trimestre <> '0') then
      // Il s'agit d'une déclaration mensuelle
    begin
      NumMois := ((StrToInt(Trimestre) - 1) * 3) + StrToInt(Mois);
      Mois := IntToStr(NumMois);
      if Length(Mois) = 1 then Mois := '0' + Mois;
      libperiod := RechDom('PGMOIS', Mois, FALSE) + ' 20' + An;
    end
    else
      if (Mois = '0') and (Trimestre <> '0') then
      // Il s'agit d'une déclaration trimestrielle
      LibPeriod := Trimestre + ' TRIMESTRE 20' + An
    else
      if ((Mois = '0') and (Trimestre = '0')) then
      // Il s'agit d'une déclaration annuelle
      libperiod := 'ANNEE 20' + An;

    Result := Uppercase(libperiod);
  end;
  //@FORMATCODIF([PDD_CODIFEDITEE];[POG_LONGEDITABLE];[PDD_LIBELLE];
  //         [PDD_TYPECOTISATION];[POG_NATUREORG];[PDD_INSTITUTION],
  //         [POG_POSTOTAL],[POG_LONGTOTAL])
  if sf = 'FORMATCODIF' then
    // mise en forme de la zone "Code Libellé de cotisation"
    // présentation particulière pour les lignes de type "Sous Total",
    // en fonction du type de DUCS.
  begin
    LongEditable := 0;
    blanc := '     ';
    Codification := READTOKENST(sp);
    StLongEditable := READTOKENST(sp);
    if (StLongEditable <> '') then LongEditable := StrToInt(StLongEditable);
    Libelle := READTOKENST(sp);
    TypeCotisation := READTOKENST(sp);
    NatureOrg := READTOKENST(sp);
    Institution := READTOKENST(sp);
    // début PT28
    Atampon := READTOKENST(sp);
    if isnumeric(Atampon) then
      DebutST := StrToInt(Atampon);
    Atampon := READTOKENST(sp);
    if isnumeric(Atampon) then
      LongEditST := StrToInt(Atampon);
    if (TypeCotisation = 'S') and (NatureOrg = '200') then
      Codif := blanc + ' ' + Codification + ' ' + Libelle
    else
      if (TypeCotisation = 'S') and (NatureOrg = '300') then
    begin
      // début PT28
      if (DebutST <> 0) and (LongEditST <> 0) then
      begin
        Codif := Codification + ' ';
        for i := 1 to 19 - (Length(TrimRight(libelle))) do
        begin
          Codif := Codif + ' ';
        end;
        Codif := Codif + TrimRight(libelle);
      end
      else
      begin
        // fin PT28

        {PT7  début
                      QRech:=OpenSql('SELECT PIN_ABREGE '+
                                    'FROM INSTITUTIONPAIE '+
                                    'WHERE PIN_INSTITUTION="'+Institution+'" ',TRUE);}
        QRech := OpenSql('SELECT PIP_ABREGE ' +
          'FROM INSTITUTIONPAYE ' +
          'WHERE PIP_INSTITUTION="' + Institution + '" ', TRUE);
        // PT7 fin
        // PT14-1         Libelle := QRech.Fields[0].Asstring;
        Libelle := ''; // PORTAGECWAS
        if not QRech.EOF then // PORTAGECWAS
          Libelle := Copy(QRech.Fields[0].Asstring, 1, 7); // PT14-1

        Ferme(QRech);
        // ci-après les lignes seront à activer si le terme "AGFF" doit apparaître sur
        // la ligne de Sous-total
        //              if (Codification = '8') then
        //                Codif := blanc + '    S/T ' + Institution + ' AGFF ' +
        //                         Copy(Libelle,1,2)
        //              else
        Codif := blanc + '    S/T ' + Institution + ' ' + Libelle; // PT60
      end; // PT28
    end
    else
      if (TypeCotisation = 'I') then
      Codif := blanc + ' ' + Libelle
    else
      // PT37          Codif := copy(blanc,1,5-LongEditable) + copy(Codification, 1, LongEditable) + ' ' + Libelle;
      Codif := copy(Codification, 1, LongEditable) + copy(blanc, 1, 5 - length(Codification)) + ' ' + Libelle;
    Result := Codif;
  end;
  // @FORMATEFFECTIF([PDD_EFFECTIF];[PDD_TYPECOTISATION])
  if sf = 'FORMATEFFECTIF' then
    // mise en forme de la zone "Nombre de salariés ou d'assurés
    // édition de ****** si cotisation de type "taux AT" ou "Sous Total"
  begin
    Effectif := READTOKENST(sp);
    if Effectif = '' then
      Result := ''
    else
    begin
      TypeCotisation := READTOKENST(sp);
      if (TypeCotisation = 'A') or (TypeCotisation = 'S') then
        Result := '******'
      else
        if (TypeCotisation = 'I') then
        Result := ' '
      else
        Result := FormatFloat('#0', Valeur(Effectif)); // PT77
    end;
  end;
  // @FORMATBASE([PDD_BASECOTISATION];[PDD_TYPECOTISATION];[PDD_CODIFEDITEE];
  //             [POG_NATUREORG];[CRI_TYPDECL];[POG_BASETYPARR];
  //             [PDD_CODIFICATION];[ALSACEMOSELLE])
  if sf = 'FORMATBASE' then
    // mise en forme de la zone "Base"
    // édition de *********** si cotisation de type "taux AT" ou
    // "Montant (GMP)" ou "Sous Total"
    // pas d'éditon si cotisation "déduction bas salaires" (URSSAF)
    // édition avec les centimes pour une DUCS BTP de type "déclaratif"
  begin
    BaseCotisation := READTOKENST(sp);
    if (BaseCotisation = '') then BaseCotisation := '0';
    WBaseCotisation := BaseCotisation; // signe
    if (Valeur(BaseCotisation) < 0) then // PT77
      BaseCotisation := FloatToStr(Valeur(BaseCotisation) * -1); // signe   // PT77
    if BaseCotisation = '' then
      Result := '0'
    else
    begin
      TypeCotisation := READTOKENST(sp);
      Codification := READTOKENST(sp);
      NatureOrg := READTOKENST(sp);
      BoolCalcul := READTOKENST(sp);
      BaseArr := READTOKENST(sp);
      Codif := READTOKENST(sp);
      AlsaceMoselle := READTOKENST(sp); //PT112

      // récup arrondi de la base pour la codif
// d PT112
      if  (AlsaceMoselle <> 'X') then
      begin
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '100D') then
          Codif := '1A0100D';
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '100P') then
          Codif := '1A0100P';
        QRech := OpenSql('SELECT PDP_BASETYPARR, PDP_PREDEFINI ' +
                         'FROM DUCSPARAM ' +
                         'WHERE ##PDP_PREDEFINI## PDP_CODIFICATION="' +
                         Codif + '" ', TRUE);
      end
      else
      begin
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '101D') then
          Codif := '1A0101D';
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '101P') then
          Codif := '1A0101P';
        QRech := OpenSql('SELECT PDP_BASETYPARR, PDP_PREDEFINI ' +
                         'FROM DUCSPARAM ' +
                         'WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE="' +
                         Codif + '" ', TRUE);
      end;
// f PT112

      CodifBaseArr := ''; // PORTAGECWAS
//      if not QRech.EOF then // PORTAGECWAS
      while not QRech.EOF do
      begin
        CodifBaseArr := QRech.Fields[0].Asstring;
        Predef := QRech.Fields[1].Asstring;
        if (Predef = 'DOS') then break;
        QRech.Next;
      end;
// f PT101      
      Ferme(QRech);
      if (CodifBaseArr <> '') then
        BaseArr := CodifBaseArr;
      // début PT8-4
      if ((Codification = '498') and (NatureOrg = '100')) or
        (TypeCotisation = 'I') then
        Result := ''
      else
        // fin PT8-4
        if (TypeCotisation = 'A') or (TypeCotisation = 'M')
        or (TypeCotisation = 'S') then
        Result := '***********'
      else
        //début PT8-4    if ((Codification= '498') AND (NatureOrg = '100')) OR
        //              (TypeCotisation = 'I') then
        //              Result := ''
        //fin PT8-4         else
        if ((NatureOrg = '700') and (BoolCalcul = '-')) or (BaseArr = '') then
        Result := FormatFloat('#0.00', Valeur(BaseCotisation)) // PT77
      else
        Result := FormatFloat('#0', Valeur(BaseCotisation)); // PT77
    end;
  end;
  // @SIGNEBASE()
  if sf = 'SIGNEBASE' then
    // édition du sihne si valeur négative
  begin
    if (WBaseCotisation = '') then WBaseCotisation := '0';
    if (Valeur(WBaseCotisation) < 0) then // PT77
      result := '-'
    else
      result := ' ';
  end;
  //@FORMATTAUX([PDD_TAUXCOTISATION];[PDD_CODIFEDITEE];[POG_NATUREORG];
  //            [CRI_TYPDECL];[PDD_TYPECOTISATION];[PDD_EFFECTIF])
  if sf = 'FORMATTAUX' then
    // mise en forme de la zone "Taux ou Quantité"
    // pas d'éditon si cotisation "déduction bas salaires" (URSSAF)
    // édition de ****** si déclaration BTP de type "déclaratif"  ou de
    // type Montant (GMP)
    // édition d'un nombre entier si cotisation de type "quantité" (forfait)
  begin
    TauxCotisation := READTOKENST(sp);
    if TauxCotisation = '' then
      Result := ''
    else
    begin
      Codification := READTOKENST(sp);
      NatureOrg := READTOKENST(sp);
      BoolCalcul := READTOKENST(sp);
      TypeCotisation := READTOKENST(sp);
      Effectif := READTOKENST(sp); // PT14-2
      if ((Codification = '498') and (NatureOrg = '100')) or
        (TypeCotisation = 'I') then
        Result := ''
      else
        if ((NatureOrg = '700') and (BoolCalcul = '-')) or
        (TypeCotisation = 'M') or (TypeCotisation = 'S') then
        Result := '********'
      else
        if (TypeCotisation = 'Q') then
        // PT14-2             Result := FormatFloat('#0',StrToFloat(TauxCotisation))
        Result := FormatFloat('#0', Valeur(Effectif)) // PT14-2 // PT77
      else
// d PT123
        if (TypeCotisation = 'C') or (TypeCotisation = 'D') then    // base & montant ou quantité montant
        Result := ''
      else
// f PT123
        Result := FormatFloat('#0.0000', Valeur(TauxCotisation)); // PT77
    end;
  end;

  // @FORMATMONTANT([PDD_MTCOTISAT];[PDD_TYPECOTISATION];[POG_NATUREORG];
  //                [CRI_TYPDECL];[POG_MTTYPARR];[PDD_CODIFICATION];[ALSACEMOSELLE])
  if sf = 'FORMATMONTANT' then
    // mise en forme de la zone "Montant"
    // pas d'édition si déclaration BTP de type "déclaratif"
    // pas d'édition si cotisation de type Taux AT
  begin
    Montant := READTOKENST(sp);
    if (Montant = '') then Montant := '0';
    WMontant := Montant; // signe
    if (Valeur(Montant) < 0) then // PT77
      Montant := FloatToStr(Valeur(Montant) * -1); // signe // PT77
    if Montant = '0' then
      Result := ''
    else
    begin
      TypeCotisation := READTOKENST(sp);
      NatureOrg := READTOKENST(sp);
      BoolCalcul := READTOKENST(sp);
      MtArr := READTOKENST(sp);
      Codif := READTOKENST(sp);
      AlsaceMoselle := READTOKENST(sp); //PT112
      // récup arrondi du montant pour la codif
// d PT112
      if (AlsaceMoselle <> 'X' ) then
      begin
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '100D') then
          Codif := '1A0100D';
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '100P') then
          Codif := '1A0100P';
        QRech := OpenSql('SELECT PDP_MTTYPARR, PDP_PREDEFINI ' +
                         'FROM DUCSPARAM ' +
                         'WHERE ##PDP_PREDEFINI## PDP_CODIFICATION="' +
                         Codif + '" ', TRUE);
      end
      else
      begin
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '101D') then
          Codif := '1A0101D';
        if (Copy(Codif, 1, 2) = '1A') and (Copy(Codif, 4, 7) = '101P') then
          Codif := '1A0101P';
        QRech := OpenSql('SELECT PDP_MTTYPARR, PDP_PREDEFINI ' +
                         'FROM DUCSPARAM ' +
                         'WHERE ##PDP_PREDEFINI## PDP_CODIFALSACE="' +
                         Codif + '" ', TRUE);
      end;
// f PT112

      CodifMtArr := ''; // PORTAGECWAS
//      if not QRech.EOF then // PORTAGECWAS
        while not QRech.EOF do
        begin
          CodifMtArr := QRech.Fields[0].Asstring;
          Predef := QRech.fields[1].Asstring;
          if (Predef = 'DOS') then break;
          QRech.Next;
        end;
// f PT101
      Ferme(QRech);
      if (CodifMtArr <> '') then
        MtArr := CodifMtArr;

      if ((NatureOrg = '700') and (BoolCalcul = '-')) or
        (TypeCotisation = 'A') or (TypeCotisation = 'I') then
        Result := ''
      else
        if (MtArr = '') then
        Result := FormatFloat('#0.00', Valeur(Montant)) // PT77
      else
        Result := FormatFloat('#0', Valeur(Montant)); // PT77
      // d PT61
      if (TypeCotisation <> 'A') and (TypeCotisation <> 'S') then
      begin
        if (Valeur(WMontant) < 0) then // PT77
          RecaDecl := RecaDecl - Valeur(Montant) // PT77
        else
          RecaDecl := RecaDecl + Valeur(Montant); // PT77
      end;
      // f PT61
    end;
  end;
  // @SIGNEMT()
  if sf = 'SIGNEMT' then
    // édition du signe si valeur négative
  begin
    if (WMontant = '') then WMontant := '0';
    if (Valeur(WMontant) < 0) then // PT77
      result := '-'
    else
      result := ' ';
  end;

  // @FORMATMONNAIE([MONNAIETENUE])
  if sf = 'FORMATMONNAIE' then
    // mise en forme de la zone "Unité Monétaire"
    // édition sur la dernière page uniquement
  begin
    Monnaie := READTOKENST(sp);

    if (NbPage <> NoPageNum) then
      Result := ' '
    else
      if (Monnaie = '001') or (Monnaie = 'FRF') then // PT9
      Result := '(francs)'
    else
      Result := '(euros)';
  end;
  // @FORMATDECLARE([TYPDECL];[NATUREORG];[ETABLISSEMENT];
  //                [ORGANISME];[DATEDEBUT];[DATEFIN];[PDU_NUM])
  if sf = 'FORMATDECLARE' then
    // mise en forme de la zone "TOTAL DECLARE"
    // édition sur la dernière page uniquement
    // sinon édition de **************
    // édition si différent de 0
    // Non renseignée pour BTP si édition de type déclaratif
  begin
    BoolCalcul := READTOKENST(sp);
    NatureOrg := READTOKENST(sp);
    if isnumeric(NatureOrg) then
      NatureOrg := ColleZeroDevant(StrToInt(NatureOrg), 3);
    Etab := READTOKENST(sp);
    // PT19     if  isnumeric(Etab) then
    // PT19        Etab := ColleZeroDevant(StrToInt(Etab),3);
    Organisme := READTOKENST(sp);
    // PT94    if isnumeric(Organisme) then
    // PT94     Organisme := ColleZeroDevant(StrToInt(Organisme), 3);
    DebPer := READTOKENST(sp);
    FinPer := READTOKENST(sp);
    NumDucs := READTOKENST(sp); // PT25
    if (NumDucs = '') then NumDucs := '0'; // DB2
    if (NbPage <> NoPageNum) then
      Result := '***************'
    else
      if (BoolCalcul = '-') and (NatureOrg = '700') then
      Result := ' '
    else
    begin
      StQQ := 'PDD_ETABLISSEMENT = "' + Etab + '" AND ' +
        'PDD_ORGANISME = "' + Organisme + '" AND ' +
        'PDD_DATEDEBUT = "' + UsDateTime(StrToDate(DebPer)) + '" AND ' +
        'PDD_DATEFIN = "' + UsDateTime(StrToDate(FinPer)) + '" AND ' +
        'PDD_NUM = ' + NumDucs + ' AND ' + // PT25  // DB2
      'PDD_TYPECOTISATION <> "S"'; // PT6-2
      QQ := OpenSQL('SELECT SUM(PDD_MTCOTISAT) AS DECLAR FROM DUCSDETAIL WHERE ' + StQQ, True);
      Declare := ''; // PORTAGECWAS
      if not QQ.EOF then // PORTAGECWAS
        Declare := FormatFloat('#0.00', QQ.FindField('DECLAR').asFloat);
      WDeclare := Declare; // signe
      if (Valeur(Declare) < 0) then // PT77
        // PT8-3     Declare := FloatToStr(StrToFloat(Declare) *-1);      // signe
        Result := FormatFloat('#0.00', Valeur(Declare) * -1) // PT77
      else // signe
        if (Declare <> '0,00') then
        Result := Declare
      else
        Result := ' ';
      Ferme(QQ);
    end;
  end;

  // d PT61
  if sf = 'FORMATDECL' then
  begin
    Result := FormatFloat('#0.00', RecaDecl);
    Wdeclare := FormatFloat('#0.00', RecaDecl);
  end;
  //f PT61

  // @SIGNEDECLARE()
  if sf = 'SIGNEDECLARE' then
    // édition du signe si valeur négative
  begin
    if (WDeclare = '') then WDeclare := '0';
    if (Valeur(WDeclare) < 0) then // PT77
      result := '-'
    else
      result := ' ';
  end;
  // @FORMATAPAYER([DECLAR];[TYPDECL];[NATUREORG];[ACOMPTES];[REGULARISATION])
  if sf = 'FORMATAPAYER' then
    // mise en forme de la zone "Montant à payer"
    // édition sur la dernière page uniquement
    // sinon édition de **************
    // édition si différent de 0
    // Non renseignée pour BTP si édition de type déclaratif

  begin
    Declare := READTOKENST(sp);
    if (Declare = ' ') then Declare := '0,00';
    BoolCalcul := READTOKENST(sp);
    NatureOrg := READTOKENST(sp);
    if isnumeric(NatureOrg) then
      NatureOrg := ColleZeroDevant(StrToInt(NatureOrg), 3);
    Acompte := READTOKENST(sp);
    Regul := READTOKENST(sp);
    if (NbPage <> NoPageNum) then
      Result := '***************'
    else
      if (BoolCalcul = '-') and (NatureOrg = '700') then
      Result := ' '
    else
    begin
      //PT8-3              NAcompte := Valeur(Acompte);
      //PT8-3            NRegul := Valeur(Regul);
      NAcompte := Valeur(WAcompte);
      NRegul := Valeur(WRegul);
      Apayer := FormatFloat('#0.00',
        (Valeur(WDeclare) - NAcompte + NRegul)); // PT77
      //PT3-3                    (StrToFloat(Declare)-NAcompte+NRegul));
      WApayer := Apayer; // signe
      if (Valeur(Apayer) < 0) then // PT77
        //PT3-3                Apayer := FloatToStr(StrToFloat(Apayer) *-1);      // signe
        Result := FormatFloat('#0.00', Valeur(Apayer) * -1) // signe // PT77
      else
        if (Apayer <> '0,00') then
        Result := Apayer
      else
        Result := ' ';
    end;

  end;
  // @SIGNEAPAYER()
  if sf = 'SIGNEAPAYER' then
    // édition du signe si valeur négative
  begin
    if (WAPayer = '') then WAPayer := '0';
    if (Valeur(WAPayer) < 0) then // PT77
      result := '-'
    else
      result := ' ';
  end;
  // @FORMATACOMPTE([ACOMPTES];[TYPDECL];[NATUREORG])
  if sf = 'FORMATACOMPTE' then
    // mise en forme de la zone "Acompte"
    // édition sur la dernière page uniquement
    // sinon édition de **************
    // édition si différent de 0
    // Non renseignée pour BTP si édition de type déclaratif
  begin
    Acompte := READTOKENST(sp);
    WAcompte := Acompte; // signe
    NAcompte := Valeur(Acompte);
    BoolCalcul := READTOKENST(sp);
    NatureOrg := READTOKENST(sp);
    if isnumeric(NatureOrg) then
      NatureOrg := ColleZeroDevant(StrToInt(NatureOrg), 3);
    if (NbPage <> NoPageNum) then
      Result := '***************'
    else
      if ((BoolCalcul = '-') and (NatureOrg = '700')) then
      result := ' '
    else
      if (NAcompte < 0) then
      // PT8-3            Acompte := FloatToStr(NAcompte *-1);      // signe
      Result := FormatFloat('#0.00', NAcompte * -1) // signe
    else
      //PT8-3         if (Acompte = '') OR (Acompte ='0,00') OR
      //PT3-3            ((BoolCalcul = '-') AND (NatureOrg = '700')) then
      if (Acompte = '') or (Acompte = '0,00') then
      Result := ' '
    else
      Result := FormatFloat('#0.00', NAcompte);
  end;

  // @SIGNEACOMPTE()
  if sf = 'SIGNEACOMPTE' then
    // édition du signe si valeur négative
  begin
    if (WAcompte = '') then WAcompte := '0';
    NAcompte := Valeur(WAcompte);
    if (NAcompte < 0) then
      result := '-'
    else
      result := ' ';
  end;

  // @FORMATREGUL([REGULARISATION];[NATUREORG];[TYPDECL])
  if sf = 'FORMATREGUL' then
    // mise en forme de la zone "Régularisations diverses"
    // édition sur la dernière page uniquement
    // sinon édition de **************
    // édition si différent de 0
    // Non renseignée pour l'ASSEDIC
    // Non renseignée pour BTP si édition de type déclaratif
  begin
    Regul := READTOKENST(sp);
    NRegul := Valeur(Regul);
    WRegul := Regul; // signe
    NatureOrg := READTOKENST(sp);
    if isnumeric(NatureOrg) then
      NatureOrg := ColleZeroDevant(StrToInt(NatureOrg), 3);
    BoolCalcul := READTOKENST(sp);
    if (NbPage <> NoPageNum) then
      Result := '***************'
    else
      if (NatureOrg = '200') or ((BoolCalcul = '-') and (NatureOrg = '700')) then
      result := ' '
    else
      if (NRegul < 0) then
      // PT8-3     Regul := FloatToStr(NRegul *-1);      // signe
      Result := FormatFloat('#0.00', NRegul * -1) // signe
    else
      // PT8-3         if (Regul = '') OR (Regul = '0,00') OR (NatureOrg = '200')OR
      // PT8-3              ((BoolCalcul = '-') AND (NatureOrg = '700')) then
      if (Regul = '') or (Regul = '0,00') then
      Result := ' '
    else
      Result := FormatFloat('#0.00', NRegul);
  end;

  // @SIGNEREGUL()
  if sf = 'SIGNEREGUL' then
    // édition du signe si valeur négative
  begin
    if (WRegul = '') then WRegul := '0';
    NRegul := Valeur(WRegul);
    if (NRegul < 0) then
      result := '-'
    else
      result := ' ';
  end;

  // @LIGNEOPTIQUE([PDU_LIGNEOPTIQUE])
  if sf = 'LIGNEOPTIQUE' then
  begin
    LigneOptique := READTOKENST(sp);
    result := LigneOptique;
  end;

  // @LIGNEOPTIQ2([TYPDECL])
  if sf = 'LIGNEOPTIQ2' then
  begin
    BoolCalcul := READTOKENST(sp);
    if (NbPage = NoPageNum) and (BoolCalcul = 'X') then
    begin
      LigneOptiq2 := LigneOptique;
      result := LigneOptiq2;
    end
    else
      result := ' ';
  end;

  // @CESSATION([PDU_CESSATION])
  if (sf = 'CESSATION') then
  begin
    Result := ' ';
    DebPer := READTOKENST(sp);
    if (DebPer <> '') and (UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) <> UsDateTime(IDate1900)) then
      // PT6-3
      if (NoPageNum = 1) then
        Result := 'X'
      else
        Result := ' ';
    // fin PT6-3
  end;

  // @CONTINUATION([PDU_CONTINUATION])
  if (sf = 'CONTINUATION') then
  begin
    Result := ' ';
    DebPer := READTOKENST(sp);
    if (DebPer <> '') and (UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) <> UsDateTime(IDate1900)) then
      // PT6-3
      if (NoPageNum = 1) then
        Result := 'X'
      else
        Result := ' ';

  end;

  // @LIBEFFTOT([POG_NATUREORG])
  if (sf = 'LIBEFFTOT') then
  begin
    NatureOrg := READTOKENST(sp);
    if (NatureOrg = '300') or (NatureOrg = '600') then
      // IRC ou MSA
      Result := ' '
    else
      // URSSAF, ASSEDIC ou BTP
      if (NoPageNum = 1) then
      Result := 'TOTAL :'
    else
      Result := '';
  end;

  // @EFFTOT([POG_NATUREORG];[PDU_NBSALFPE];[PDU_TOTAPPRENTI];[PDU_TOTHOMMES];[PDU_TOTFEMMES])
  if (sf = 'EFFTOT') then
  begin
    NatureOrg := READTOKENST(sp);
    NbSalFpe := READTOKENST(sp);
    // d PT80
    //   NbSalHap :=  READTOKENST(sp);
    NbApprentis := READTOKENST(sp);
    // f PT80
    NbHommes := READTOKENST(sp);
    NbFemmes := READTOKENST(sp);
    if (NbHommes = '') then NbHommes := '0'; // PT12
    if (NbFemmes = '') then NbFemmes := '0'; // PT12
    // d PT80
    if (NbApprentis = '') then NbApprentis := '0';
    //   NbApprentis := StrToInt(NbSalFpe)-StrToInt(NbSalHap);
    // f PT80
    if (NatureOrg = '100') then
    begin
      // URSSAF
      if (NoPageNum = 1) then
        //      Result := NbSalHap      // PT8-2
        Result := IntToStr(StrToInt(NbHommes) + StrToInt(NbFemmes) - StrToInt(NbApprentis)) // PT80
      else
        Result := '';
    end
    else
      if (NatureOrg = '300') or (NatureOrg = '600') then
      // IRC ou MSA
      Result := ' '
    else
    begin
      // ASSEDIC ou BTP
      if (NoPageNum = 1) then
        //       Result := NbSalFpe                                     // PT8-2
        Result := IntToStr(StrToInt(NbHommes) + StrToInt(NbFemmes)) // PT8-2
      else
        Result := '';
    end;
  end;

  // @TIT1([POG_NATUREORG])
  if (sf = 'TIT1') then
  begin
    NatureOrg := READTOKENST(sp);
    if (NatureOrg = '300') or (NatureOrg = '600') then
      // IRC ou MSA
      Result := ' '
    else
    begin
      // URSSAF, ASSEDIC, BTP
      if (NoPageNum = 1) then
        Result := 'Nombre de salariés ou'
      else
        Result := '';
    end;
  end;

  // @TIT2([POG_NATUREORG])
  if (sf = 'TIT2') then
  begin
    NatureOrg := READTOKENST(sp);
    if (NatureOrg = '300') or (NatureOrg = '600') then
      // IRC ou MSA
      Result := ' '
    else
    begin
      // URSSAF, ASSEDIC, BTP
      if (NoPageNum = 1) then
        Result := 'd''assurés au dernier' // PT38
      else
        Result := '';
    end;
  end;

  // @TIT3([POG_NATUREORG])
  if (sf = 'TIT3') then
  begin
    NatureOrg := READTOKENST(sp);
    if (NatureOrg = '300') or (NatureOrg = '600') then
      // IRC ou MSA
      Result := ' '
    else
    begin
      // URSSAF, ASSEDIC, BTP
      if (NoPageNum = 1) then
        Result := 'jour de la période :'
      else
        Result := '';
    end;
  end;
  // fin PT6-3
  // début PT34 VLU
  // @VLUHOMMES([POG_ETABLISSEMENT];[PDU_DATEDEBUT];[PDU_DATEFIN])
  if (sf = 'VLUHOMMES') then
  begin
    Etab := READTOKENST(sp);
    DebPer := READTOKENST(sp);
    FinPer := READTOKENST(sp);

    StQQ := '';
//PT110 remplacé PSA_DATESORTIE=...idate1900 par <=    
    StQQ := 'PSA_ETABLISSEMENT="' + Etab + '" AND ' +
      '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR ' +
      'PSA_DATESORTIE IS NULL ' +
      'OR PSA_DATESORTIE>="' + UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) + '")' +
      ' AND ' +
      'PSA_DATEENTREE<="' + UsDateTime(StrToDate(DateToStr(Valeur(FinPer)))) + '"';

    QQ := OpenSQL('SELECT COUNT(*) FROM SALARIES WHERE ' + StQQ + ' AND PSA_SEXE="M"', True);
    Result := ''; // PORTAGECWAS
    if not QQ.EOF then // PORTAGECWAS
      Result := FormatFloat('#0', Valeur(IntToStr(QQ.Fields[0].AsInteger))); // PT77
    Ferme(QQ);
  end;
  // @VLUFEMMES([POG_ETABLISSEMENT];[PDU_DATEDEBUT];[PDU_DATEFIN])
  if (sf = 'VLUFEMMES') then
  begin
    Etab := READTOKENST(sp);
    DebPer := READTOKENST(sp);
    FinPer := READTOKENST(sp);

    StQQ := '';
//PT110 remplacé PSA_DATESORTIE=...idate1900 par <=
    StQQ := 'PSA_ETABLISSEMENT="' + Etab + '" AND ' +
      '(PSA_DATESORTIE<="' + UsDateTime(IDate1900) + '" OR ' +
      'PSA_DATESORTIE IS NULL ' +
      'OR PSA_DATESORTIE>="' + UsDateTime(StrToDate(DateToStr(Valeur(DebPer)))) + '")' +
      ' AND ' +
      'PSA_DATEENTREE<="' + UsDateTime(StrToDate(DateToStr(Valeur(FinPer)))) + '"';

    QQ := OpenSQL('SELECT COUNT(*) FROM SALARIES WHERE ' + StQQ + ' AND PSA_SEXE="F"', True);
    Result := ''; // PORTAGECWAS
    if not QQ.EOF then // PORTAGECWAS
      Result := FormatFloat('#0', Valeur(IntToStr(QQ.Fields[0].AsInteger))); // PT77
    Ferme(QQ);
  end;
  // fin PT34 VLU

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Remise à zéro du n° de page en début d'édition des DUCS
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

function FunctPGDebEditDucs(): variant;
begin
  NoPageNum := 0;
  result := NoPageNum;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MF
Créé le ...... : 05/09/2001
Modifié le ... :   /  /
Description .. : Remise à zéro de la variable AncNumPagSys en fin
Suite ........ : d'édition de DUCS
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

function FunctPGFinEditDucs(): variant;
begin
  AncNumPagSys := '';
  result := AncNumPagSys;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : JL
Créé le ...... : 08/11/2001
Modifié le ... :   /  /
Description .. :Edition de la DADS-U
Suite ........ :
Mots clefs ... : PAIE, PGDADS-U
*****************************************************************}
function FunctPGDADSURecap(sf, sp: string): variant;
var
TobDadsP: Tob;
QDADSPER: TQuery;
EtabDads, Organisme, Sal, Segment, StOrg, StSal, StType, ValeurDads: string;
begin
ValeurDADS:= '0';
sal:= (READTOKENST(sp));
Segment:= (READTOKENST(sp));
EtabDADS:= (READTOKENST(sp));
Organisme:= (READTOKENST(sp));
StSal:= '';
ValeurDads:= '0';
if organisme <> '' then
   begin
   if EtabDADS <> '' then
      begin
      if PGGlbSalDe <> '' then
         StSal:= 'AND D1.PDS_SALARIE>="'+PGGlbSalDe+'" ';
      if PGGlbSalA <> '' then
         StSal:= StSal+'AND D1.PDS_SALARIE<="'+PGGlbSalA+'" ';
      if PGGlbOrg <> '' then
         StOrg:= ' AND D3.PDS_DONNEEAFFICH="'+PGGlbOrg+'"';
      if TOBDADSRECAP = nil then
         begin
         ValeurDADS:= '0';
         if (PGGlbTypeDADS = '001') then
            StType:= ' D1.PDS_TYPE="001" AND'
         else
            StType:= ' D1.PDS_TYPE<>"001" AND';
         QDADSPER:= OpenSQL ('SELECT D1.PDS_SALARIE, D1.PDS_ORDRE,'+
                            ' D1.PDS_ORDRESEG, D1.PDS_SEGMENT,'+
                            ' (D1.PDS_DONNEEAFFICH) AS VALEUR,'+
                            ' (D2.PDS_DONNEEAFFICH) ETABLISS,'+
                            ' (D3.PDS_DONNEEAFFICH) AS Org'+
                            ' FROM DADSDETAIL D1'+
                            ' LEFT JOIN DADSDETAIL D2 ON'+
                            ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                            ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                            ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                            ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                            ' D1.PDS_DATEFIN=D2.PDS_DATEFIN'+
                            ' LEFT JOIN DADSDETAIL D3 ON'+
                            ' D1.PDS_SALARIE=D3.PDS_SALARIE AND'+
                            ' D1.PDS_ORDRE=D3.PDS_ORDRE AND'+
                            ' D1.PDS_TYPE=D3.PDS_TYPE AND'+
                            ' D1.PDS_DATEDEBUT=D3.PDS_DATEDEBUT AND'+
                            ' D1.PDS_DATEFIN=D3.PDS_DATEFIN WHERE'+
                            ' D1.PDS_ORDRE<>0 AND'+StType+
                            ' D1.PDS_DATEDEBUT>="'+UsDateTime(DateDebDADS)+'" AND'+
                            ' D1.PDS_SEGMENT>="S41.G01.00.022" AND'+
                            ' D1.PDS_SEGMENT<="S41.G01.00.069" AND'+
                            ' D1.PDS_DATEDEBUT<="'+UsDateTime(DateFinDADS)+'" AND'+
                            ' D1.PDS_SALARIE NOT LIKE "--H%" '+StSal+' AND'+
                            ' D2.PDS_SEGMENT="S41.G01.00.005" AND'+
                            ' D2.PDS_DONNEEAFFICH>="'+PGGlbEtabDe+'" AND'+
                            ' D2.PDS_DONNEEAFFICH<="'+PGGlbEtabA+'" AND'+
                            ' D3.PDS_SEGMENT="S41.G01.01.001"'+StOrg+
                            ' ORDER BY D1.PDS_SALARIE,D1.PDS_ORDRE', True);
         TOBDADSRECAP:= TOB.Create('DADS périodes', nil, -1);
         TOBDADSRECAP.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
         Ferme(QDADSPer);
         end;
      TobDadsP:= TOBDADSRECAP.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT', 'ETABLISS', 'Org'],
                                        [Sal, Segment, EtabDADS, Organisme], False);
      while TobdadsP <> nil do
            begin
            ValeurDADS:= IntToStr(StrToInt(ValeurDads)+
                         StrToInt(TobDadsP.GetValue('VALEUR')));
            TobDadsP:= TOBDADSRECAP.FindNext(['PDS_SALARIE', 'PDS_SEGMENT', 'ETABLISS', 'Org'],
                                             [Sal, Segment, EtabDADS, Organisme], False);
            end;
      end
   else
      begin
      if PGGlbSalDe <> '' then
         StSal:= ' AND D1.PDS_SALARIE>="'+PGGlbSalDe+'"';
      if PGGlbSalA <> '' then
         StSal:= StSal+' AND D1.PDS_SALARIE<="'+PGGlbSalA+'"';
      if PGGlbOrg <> '' then
         StOrg:= ' AND D3.PDS_DONNEEAFFICH="'+PGGlbOrg+'"';
      if TOBDADSRECAP = nil then
         begin
         ValeurDADS:= '0';
         if (PGGlbTypeDADS = '001') then
            StType:= ' D1.PDS_TYPE="001" AND'
         else
            StType:= ' D1.PDS_TYPE<>"001" AND';
         QDADSPER:= OpenSQL('SELECT D1.PDS_SALARIE, D1.PDS_ORDRE,'+
                            ' D1.PDS_ORDRESEG, D1.PDS_SEGMENT,'+
                            ' (D1.PDS_DONNEEAFFICH) AS VALEUR,'+
                            ' (D3.PDS_DONNEEAFFICH) AS Org'+
                            ' FROM DADSDETAIL D1'+
                            ' LEFT JOIN DADSDETAIL D3 ON'+
                            ' D1.PDS_SALARIE=D3.PDS_SALARIE AND'+
                            ' D1.PDS_ORDRE=D3.PDS_ORDRE AND'+
                            ' D1.PDS_TYPE=D3.PDS_TYPE AND'+
                            ' D1.PDS_DATEDEBUT=D3.PDS_DATEDEBUT AND'+
                            ' D1.PDS_DATEFIN=D3.PDS_DATEFIN WHERE'+
                            ' D1.PDS_ORDRE<>0 AND'+StType+
                            ' D1.PDS_DATEDEBUT>="'+UsDateTime(DateDebDADS)+'" AND'+
                            ' D1.PDS_SEGMENT>="S41.G01.00.022" AND'+
                            ' D1.PDS_SEGMENT<="S41.G01.00.069" AND'+
                            ' D1.PDS_DATEDEBUT<="'+UsDateTime(DateFinDADS)+'" AND'+
                            ' D1.PDS_SALARIE NOT LIKE "--H%" '+StSal+' AND'+
                            ' D3.PDS_SEGMENT="S41.G01.01.001" '+StOrg+
                            ' ORDER BY D1.PDS_SALARIE,D1.PDS_ORDRE', True);
         TOBDADSRECAP:= TOB.Create('DADS périodes', nil, -1);
         TOBDADSRECAP.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
         Ferme(QDADSPer);
         end;
      TobDadsP:= TOBDADSRECAP.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT', 'Org'],
                                        [Sal, Segment, Organisme], False);
      while TobdadsP <> nil do
            begin
            ValeurDADS:= IntToStr(StrToInt(ValeurDads)+
                         StrToInt(TobDadsP.GetValue('VALEUR')));
            TobDadsP:= TOBDADSRECAP.FindNext(['PDS_SALARIE', 'PDS_SEGMENT', 'Org'],
                                             [Sal, Segment, Organisme], False);
            end;
      end;
   end
else
   begin
   if EtabDADS <> '' then
      begin
      if PGGlbSalDe <> '' then
         StSal:= ' AND D1.PDS_SALARIE>="'+PGGlbSalDe+'"';
      if PGGlbSalA <> '' then
         StSal:= StSal+' AND D1.PDS_SALARIE<="'+PGGlbSalA+'"';
      if TOBDADSRECAP = nil then
         begin
         ValeurDADS:= '0';
         if (PGGlbTypeDADS = '001') then
            StType:= ' D1.PDS_TYPE="001" AND'
         else
            StType:= ' D1.PDS_TYPE<>"001" AND';
         QDADSPER:= OpenSQL('SELECT D1.PDS_SALARIE, D1.PDS_ORDRE,'+
                            ' D1.PDS_ORDRESEG, D1.PDS_SEGMENT,'+
                            ' (D1.PDS_DONNEEAFFICH) AS VALEUR,'+
                            ' (D2.PDS_DONNEEAFFICH) AS ETABLISS'+
                            ' FROM DADSDETAIL D1'+
                            ' LEFT JOIN DADSDETAIL D2 ON'+
                            ' D1.PDS_SALARIE=D2.PDS_SALARIE AND'+
                            ' D1.PDS_ORDRE=D2.PDS_ORDRE AND'+
                            ' D1.PDS_TYPE=D2.PDS_TYPE AND'+
                            ' D1.PDS_DATEDEBUT=D2.PDS_DATEDEBUT AND'+
                            ' D1.PDS_DATEFIN=D2.PDS_DATEFIN WHERE'+
                            ' D1.PDS_ORDRE<>0 AND'+StType+
                            ' D1.PDS_DATEDEBUT>="'+UsDateTime(DateDebDADS)+'" AND'+
                            ' D1.PDS_SEGMENT>="S41.G01.00.022" AND'+
                            ' D1.PDS_SEGMENT<="S41.G01.00.069" AND'+
                            ' D1.PDS_DATEDEBUT<="'+UsDateTime(DateFinDADS)+'" AND'+
                            ' D1.PDS_SALARIE NOT LIKE "--H%" '+StSal+' AND'+
                            ' D2.PDS_SEGMENT="S41.G01.00.005" AND'+
                            ' D2.PDS_DONNEEAFFICH>="'+PGGlbEtabDe+'" AND'+
                            ' D2.PDS_DONNEEAFFICH<="'+PGGlbEtabA+'"'+
                            ' ORDER BY D1.PDS_SALARIE,D1.PDS_ORDRE', True);
         TOBDADSRECAP:= TOB.Create('DADS périodes', nil, -1);
         TOBDADSRECAP.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
         Ferme(QDADSPer);
         end;
      TobDadsP:= TOBDADSRECAP.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT', 'ETABLISS'],
                                        [Sal, Segment, EtabDADS], False);
      while TobdadsP <> nil do
            begin
            ValeurDADS:= IntToStr(StrToInt(ValeurDads)+
                         StrToInt(TobDadsP.GetValue('VALEUR')));
            TobDadsP:= TOBDADSRECAP.FindNext(['PDS_SALARIE', 'PDS_SEGMENT', 'ETABLISS'],
                                             [Sal, Segment, EtabDADS], False);
            end;
      end
   else
      begin
      if PGGlbSalDe <> '' then
         StSal:= ' AND D1.PDS_SALARIE>="'+PGGlbSalDe+'"';
      if PGGlbSalA <> '' then
         StSal:= StSal+' AND D1.PDS_SALARIE<="'+PGGlbSalA+'"';
      if TOBDADSRECAP = nil then
         begin
         ValeurDADS:= '0';
         if (PGGlbTypeDADS = '001') then
            StType:= ' D1.PDS_TYPE="001" AND'
         else
            StType:= ' D1.PDS_TYPE<>"001" AND';
         QDADSPER:= OpenSQL('SELECT D1.PDS_SALARIE, D1.PDS_ORDRE,'+
                            ' D1.PDS_ORDRESEG, D1.PDS_SEGMENT,'+
                            ' (D1.PDS_DONNEEAFFICH) AS VALEUR'+
                            ' FROM DADSDETAIL D1 WHERE'+
                            ' D1.PDS_ORDRE<>0 AND'+StType+
                            ' D1.PDS_DATEDEBUT>="'+UsDateTime(DateDebDADS)+'" AND'+
                            ' D1.PDS_SEGMENT>="S41.G01.00.022" AND'+
                            ' D1.PDS_SEGMENT<="S41.G01.00.069" AND'+
                            ' D1.PDS_DATEDEBUT<="'+UsDateTime(DateFinDADS)+'" AND'+
                            ' D1.PDS_SALARIE NOT LIKE "--H%" '+StSal+
                            ' ORDER BY D1.PDS_SALARIE,D1.PDS_ORDRE', True);
         TOBDADSRECAP:= TOB.Create('DADS périodes', nil, -1);
         TOBDADSRECAP.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
         Ferme(QDADSPer);
         end;
      TobDadsP:= TOBDADSRECAP.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT'],
                                        [Sal, Segment], False);
      while TobdadsP <> nil do
            begin
            ValeurDADS:= IntToStr(StrToInt(ValeurDads)+
                         StrToInt(TobDadsP.GetValue('VALEUR')));
            TobDadsP:= TOBDADSRECAP.FindNext(['PDS_SALARIE', 'PDS_SEGMENT'],
                                             [Sal, Segment], False);
            end;
      end;
   end;
result:= ValeurDADS;
exit;
end;


//PT45
{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 12/11/2002
Modifié le ... :
Description .. : Edition des traitements et salaires d'un salarié : édition des
Suite ........ : cumuls
Suite ........ :
Mots clefs ... : PAIE, PGDADSU
*****************************************************************}
function FunctPGDADSUSalaires(sf, sp: string): variant;
var
  TobDadsP: Tob;
  QDADSPER: TQuery;
  Ordre, StSal, Sal, Segment, ValeurDads: string;
begin
  result := ' '; //PT54
  ValeurDADS := '0';
  sal := (READTOKENST(sp));
  Segment := (READTOKENST(sp));
  Ordre := READTOKENST(sp);
  StSal := '';
  if PGGlbSalDe <> '' then
    StSal := 'AND PDS_SALARIE>="' + PGGlbSalDe + '" ';

  if Ordre = '0' then
  begin
    if TOBDADSPER = nil then
    begin
      QDADSPER := OpenSQL('SELECT PDS_SALARIE, PDS_ORDRE, PDS_ORDRESEG,' +
        ' PDS_SEGMENT, PDS_DONNEEAFFICH' +
        ' FROM DADSDETAIL WHERE' +
        ' PDS_ORDRE=0 AND' +
        ' (PDS_TYPE="001" OR PDS_TYPE="201") AND' +
        ' (PDS_SEGMENT="S30.G01.00.008.003" OR' +
        ' PDS_SEGMENT="S30.G01.00.008.006" OR' +
        ' PDS_SEGMENT="S30.G01.00.008.010" OR' +
        ' PDS_SEGMENT="S30.G01.00.008.012") AND' +
        ' PDS_DATEDEBUT>="' + UsDateTime(DateDebDADS) + '" AND' +
        ' PDS_DATEDEBUT<="' + UsDateTime(DateFinDADS) + '" AND' +
        ' PDS_SALARIE NOT LIKE "--H%" ' + StSal +
        ' Order by PDS_SALARIE, PDS_ORDRE', True);
      TOBDADSPER := TOB.Create('DADS adresses', nil, -1);
      TOBDADSPER.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
      Ferme(QDADSPer);
    end;
    TobDadsP := TOBDADSPER.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT'], [Sal, Segment], False);
    if TobdadsP <> nil then
      result := TobDadsP.GetValue('PDS_DONNEEAFFICH');
    exit;
  end;

  if Ordre = '1' then
  begin
    if TOBDADSSAL = nil then
    begin
      QDADSPER := OpenSQL('SELECT PDS_SALARIE, PDS_ORDRE, PDS_ORDRESEG,' +
        ' PDS_SEGMENT, PDS_DONNEEAFFICH' +
        ' FROM DADSDETAIL WHERE' +
        ' PDS_ORDRE<>0 AND' +
        ' (PDS_TYPE="001" OR PDS_TYPE="201") AND' +
        ' (PDS_SEGMENT="S41.G01.00.005" OR' +
        ' PDS_SEGMENT="S41.G01.00.021" OR' +
        ' PDS_SEGMENT="S41.G01.00.022" OR' +
        ' PDS_SEGMENT="S41.G01.00.035.001" OR' +
        ' PDS_SEGMENT="S41.G01.00.044.001" OR' +
        ' PDS_SEGMENT="S41.G01.00.063.001" OR' +
        ' PDS_SEGMENT="S41.G01.00.066.001") AND' +
        ' PDS_DATEDEBUT>="' + UsDateTime(DateDebDADS) + '" AND' +
        ' PDS_DATEDEBUT<="' + UsDateTime(DateFinDADS) + '" AND' +
        ' PDS_SALARIE NOT LIKE "--H%" ' + StSal +
        ' Order by PDS_SALARIE, PDS_ORDRE', True);
      TOBDADSSAL := TOB.Create('DADS Cumuls', nil, -1);
      TOBDADSSAL.LoadDetailDB('DADSDETAIL', '', '', QDADSPER, False);
      Ferme(QDADSPER);
    end;
    TobDadsP := TOBDADSSAL.FindFirst(['PDS_SALARIE', 'PDS_SEGMENT'], [Sal, Segment], False);
    while TobdadsP <> nil do
    begin
      if (Segment = 'S41.G01.00.005') then
        DADSUCodeEtab := TobDadsP.GetValue('PDS_DONNEEAFFICH')
      else
        ValeurDADS := IntToStr(StrToInt(ValeurDads) + StrToInt(TobDadsP.GetValue('PDS_DONNEEAFFICH')));
      TobDadsP := TOBDADSSAL.FindNext(['PDS_SALARIE', 'PDS_SEGMENT'], [Sal, Segment], False);
    end;
    result := ValeurDADS;
    exit;
  end;

  if Ordre = '2' then
  begin
    if TOBDADSETAB = nil then
    begin
      QDADSPER := OpenSQL('SELECT ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1,' +
        ' ET_ADRESSE2, ET_CODEPOSTAL, ET_VILLE, ET_SIRET,' +
        ' ET_APE' +
        ' FROM ETABLISS', True);
      TOBDADSETAB := TOB.Create('Etablissement', nil, -1);
      TOBDADSETAB.LoadDetailDB('ETABLISS', '', '', QDADSPER, False);
      Ferme(QDADSPER);
    end;
    TobDadsP := TOBDADSETAB.FindFirst(['ET_ETABLISSEMENT'], [DADSUCodeEtab], False);
    if TobdadsP <> nil then
    begin
      if (Segment = '1') then
      begin
        result := TobDadsP.GetValue('ET_LIBELLE');
        exit;
      end;
      if (Segment = '2') then
      begin
        result := TobDadsP.GetValue('ET_ADRESSE1');
        exit;
      end;
      if (Segment = '3') then
      begin
        result := TobDadsP.GetValue('ET_ADRESSE2');
        exit;
      end;
      if (Segment = '4') then
      begin
        result := TobDadsP.GetValue('ET_CODEPOSTAL');
        exit;
      end;
      if (Segment = '5') then
      begin
        result := TobDadsP.GetValue('ET_VILLE');
        exit;
      end;
      if (Segment = '6') then
      begin
        result := TobDadsP.GetValue('ET_SIRET');
        exit;
      end;
      if (Segment = '7') then
      begin
        result := TobDadsP.GetValue('ET_APE');
        exit;
      end;
    end;
  end;
end;
//FIN PT45


//PT56
{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 24/02/2003
Modifié le ... :
Description .. : Edition des périodes d'inactivité
Mots clefs ... : PAIE, PGDADSU
*****************************************************************}

function FunctPGDADSUInactivite(sf, sp: string): variant;
var
  TobDadsP: Tob;
  QDADSPER: TQuery;
  Ordre, StSal, Sal, Segment, ValeurDads: string;
begin
  result := ' ';
  ValeurDADS := '0';
  sal := READTOKENST(sp);
  Ordre := READTOKENST(sp);
  Segment := READTOKENST(sp);
  StSal := '';
  if PGGlbSalDe <> '' then
    StSal := 'AND PDS_SALARIE>="' + PGGlbSalDe + '" ';
  if PGGlbSalA <> '' then
    StSal := StSal + 'AND PDS_SALARIE<="' + PGGlbSalA + '" ';

  if TOBDADSPER = nil then
  begin
    QDADSPER := OpenSQL('SELECT PDS_SALARIE, PDS_ORDRE, PDS_ORDRESEG,' +
      ' PDS_SEGMENT, (PDS_DONNEEAFFICH) AS VALEUR' +
      ' FROM DADSDETAIL WHERE' +
      ' PDS_ORDRE<0 AND' +
      ' PDS_TYPE="001" AND' +
      ' PDS_DATEDEBUT>="' + UsDateTime(DateDebDADS) + '" AND' +
      ' PDS_DATEDEBUT<="' + UsDateTime(DateFinDADS) + '" AND' +
      ' PDS_SALARIE NOT LIKE "--H%" ' + StSal +
      ' Order by PDS_SALARIE, PDS_ORDRE DESC', True);
    TOBDADSPER := TOB.Create('DADS adresses', nil, -1);
    TOBDADSPER.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
    Ferme(QDADSPer);
  end;
  if Segment = '0' then
    TobDadsP := TOBDADSPER.FindFirst(['PDS_SALARIE', 'PDS_ORDRE'], [Sal, Ordre], False)
  else
    TobDadsP := TOBDADSPER.FindFirst(['PDS_SALARIE', 'PDS_ORDRE', 'PDS_SEGMENT'], [Sal, Ordre, Segment], False);
  if TobdadsP <> nil then
  begin
    if (Segment = '0') then
      result := -(TobDadsP.GetValue('PDS_ORDRE'))
    else
    begin
      if (Segment = 'S46.G01.00.001') then
        result := RechDom('PGTYPEABS', TobDadsP.GetValue('VALEUR'), False)
          //PT85
      else
        if (Segment = 'S46.G01.02.001') then
        result := RechDom('PGDUREETRAVAIL', TobDadsP.GetValue('VALEUR'), False)
          //FIN PT85
      else
        result := TobDadsP.GetValue('VALEUR');
    end;
  end;
  exit;
end;
//FIN PT56


function FunctPGDADSUNumPeriode(sf, sp: string): variant;
var
  Q: TQuery;
  TobDadsNP: Tob;
  NBPeriodeDADS, ComptePer: Integer;
  Ordre, Sal, StEtab, StSal, StType: string;
begin
  NBPeriodeDADS := 1;
  sal := (READTOKENST(sp));
  Ordre := (READTOKENST(sp));
  StSal := '';
  if PGGlbSalDe <> '' then StSal := ' AND PDS_SALARIE>="' + PGGlbSalDe + '" '; //PT21-1 Ajout critère Sal de
  if PGGlbSalA <> '' then StSal := StSal + ' AND PDS_SALARIE<="' + PGGlbSalA + '" '; //PT21-1 Ajout Sal etab à
  if PGGlbEtabDe <> '' then StEtab := ' AND PDS_DONNEEAFFICH>="' + PGGlbEtabDe + '"';
  if PGGlBEtabA <> '' then StEtab := StEtab + ' AND PDS_DONNEEAFFICH<="' + PGGlBEtabA + '"';
  if TobDadsNumPer = nil then
  begin
    //PT82
    if (PGGlbTypeDADS = '001') then
      StType := ' PDS_TYPE="001" AND'
    else
      StType := ' PDS_TYPE<>"001" AND';
    //FIN PT82
    Q := OpenSQL('SELECT DISTINCT(PDS_ORDRE),PDS_SALARIE' +
      ' FROM DADSDETAIL WHERE' +
      ' PDS_DATEDEBUT>="' + UsDateTime(DateDebDADS) + '"' + StSal + StEtab + //PT21-1 Ajout Cond. StSal StEtab
      ' AND PDS_DATEDEBUT<="' + UsDateTime(DateFinDADS) + '" AND' +
      ' PDS_ORDRE<>0 AND' + StType +
      ' PDS_SEGMENT="S41.G01.00.005" Order by PDS_ORDRE', True);
    TobDadsNumPer := Tob.Create('les numeros de periodes', nil, -1);
    TobDadsNumPer.LoadDetailDB('DADSDETAIL', '', '', Q, False);
    Ferme(Q);
  end;
  if Ordre = 'NbPeriode' then
  begin
    ComptePer := 0;
    TobDadsNP := TobDadsNumPer.FindFirst(['PDS_SALARIE'], [Sal], False);
    while TobDadsNP <> nil do
    begin
      ComptePer := ComptePer + 1;
      TobDadsNP := TobDadsNumPer.FindNext(['PDS_SALARIE'], [Sal], False);
    end;
    Result := ComptePer;
    Exit;
  end;
  TobDadsNP := TobDadsNumPer.FindFirst(['PDS_SALARIE'], [Sal], False);
  while (TobDadsNP.GetValue('PDS_ORDRE') <> Ordre) do
  begin
    NBPeriodeDADS := NBPeriodeDADS + 1;
    TobDadsNP := TobDadsNumPer.FindNext(['PDS_SALARIE'], [Sal], False);
  end;
  Result := NBPeriodeDADS;
end;

function FunctPGDADSU(sf, sp: string): variant;
var
TOBDads: Tob;
QDADSPER: TQuery;
St, StSal, Sal, Segment, Ordre, ValeurDads, libDADS, SegmentNum: string;
begin
sal := (READTOKENST(sp));
Ordre := (READTOKENST(sp));
Segment := (READTOKENST(sp));
LibDADS := (READTOKENST(sp));
StSal := '';
St := '';
if PGGlbSalDe <> '' then
   StSal := 'AND PDS_SALARIE>="' + PGGlbSalDe + '" ';
if PGGlbSalA <> '' then
   StSal := StSal + 'AND PDS_SALARIE<="' + PGGlbSalA + '" ';
if (PGGlbTypeDADS = '001') then
   StSal := StSal + 'AND PDS_TYPE="001" '
else
   StSal := StSal + 'AND PDS_TYPE<>"001" ';
StSal:= StSal+'AND PDS_DATEDEBUT>="'+UsDateTime(DateDebDADS)+'" AND'+
        ' PDS_DATEFIN<="'+UsDateTime(DateFinDADS)+'" ';
if StSal <> '' then
   St := 'WHERE' + Copy(StSal, 4, Length(StSal));
if TOBDADSPER = nil then
   begin
   QDADSPER:= OpenSQL ('SELECT PDS_SALARIE,PDS_ORDRE,PDS_ORDRESEG,'+
                       ' PDS_SEGMENT, PDS_DONNEEAFFICH'+
                       ' FROM DADSDETAIL '+St+
                       ' ORDER BY PDS_ORDRE', True);
   TOBDADSPER:= TOB.Create('DADS périodes', nil, -1);
   TOBDADSPER.LoadDetailDB('DADSDETAIL', '', '', QDADSPer, False);
   Ferme(QDADSPer);
   end;

if (isnumeric(Segment)) then
   TOBDads:= TOBDADSPER.FindFirst (['PDS_SALARIE', 'PDS_ORDRE', 'PDS_ORDRESEG'],
                                   [Sal, Ordre, Segment], False)
else
   TOBDads:= TOBDADSPER.FindFirst (['PDS_SALARIE', 'PDS_ORDRE', 'PDS_SEGMENT'],
                                   [Sal, Ordre, Segment], False);
if TOBDads <> nil then
   begin
   if (isnumeric(Segment)) then
      SegmentNum := TOBDads.GetValue('PDS_SEGMENT');
   ValeurDADS := TOBDads.GetValue('PDS_DONNEEAFFICH');
   if (Segment = 'S30.G01.00.007') and (ValeurDADS = 'MR') then
      ValeurDADS := 'M.';
   end
else
   begin
   ValeurDADS := '';
   if (Segment = 'S30.G01.00.008.003') and (ordre = '0') then
      ValeurDADS := '';
   if (Segment = 'S30.G01.00.008.006') and (ordre = '0') then
      ValeurDADS := '';
   Result := ValeurDADS;
   exit;
   end;
if StrToInt(ordre) > 0 then
   begin
   if (Segment = 'S41.G01.00.005') then
      ValeurDADS := RechDom('TTETABLISSEMENT', ValeurDADS, False);
   if (Segment = 'S41.G01.00.002') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
//PT97
   if (Segment = 'S41.G01.00.002.001') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
   if (Segment = 'S41.G01.00.002.002') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
   if (Segment = 'S41.G01.00.002.003') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
   if (Segment = 'S41.G01.00.002.004') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
   if (Segment = 'S41.G01.00.002.005') then
      ValeurDADS := RechDom('PGMOTIFENTREE', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004.001') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004.002') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004.003') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004.004') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.004.005') then
      ValeurDADS := RechDom('PGMOTIFSORTIEDADS', ValeurDADS, False);
//FIN PT97
   if (Segment = 'S41.G01.00.013') then
      ValeurDADS := RechDom('PGCONDEMPLOI', ValeurDADS, False);
   if (Segment = 'S41.G01.00.014') then
      ValeurDADS := RechDom('PGSPROFESSIONNEL', ValeurDADS, False);
   if (Segment = 'S41.G01.00.015') then
      ValeurDADS := RechDom('PGSCATEGORIEL', ValeurDADS, False);
   if (Segment = 'S41.G01.00.012') then
      ValeurDADS := RechDom('PGTYPECONTRAT', ValeurDADS, False);
//PT97
   if (Segment = 'S41.G01.00.012.001') then
      ValeurDADS := RechDom('PGTYPECONTRAT', ValeurDADS, False);
//FIN PT97
   if (Segment = 'S41.G01.00.010') then
      ValeurDADS := RechDom('PGLIBEMPLOI', ValeurDADS, False);
   if (Segment = 'S41.G01.00.018') then
      ValeurDADS := RechDom('PGREGIMESS', ValeurDADS, False);
   if (Segment = 'S41.G01.00.017') then
      ValeurDADS := RechDom('PGCOEFFICIENT', ValeurDADS, False);
   if (Segment = 'S41.G01.00.034') then
{PT91
      begin
      if ValeurDADS = '01' then ValeurDADS := 'F';
      if ValeurDADS = '02' then ValeurDADS := 'E';
      ValeurDADS := RechDom('PGTRAVAILETRANGER', ValeurDADS, False);
      end;
}
      ValeurDADS := RechDom('PGTRAVAILETRANGER', ValeurDADS, False);
//FIN PT91
   if (Segment = 'S41.G01.00.011') then
      begin
      if VH_Paie.PGPCS2003 then
         ValeurDADS := RechDom('PGCODEPCSESE', ValeurDADS, False)
      else
         ValeurDADS := RechDom('PGCODEEMPLOI', ValeurDADS, False);
      end;
   if libDADS = 'Libelle' then
      begin
      if (SegmentNum = 'S41.G01.01.001') then
         ValeurDads := RechDom('PGINSTITUTION', ValeurDads, False);

      if (SegmentNum = 'S41.G01.02.001') then
         ValeurDADS := RechDom('PGBASEBRUTESPEC', ValeurDADS, False);

      if (SegmentNum = 'S41.G01.03.001') then
         ValeurDADS := RechDom('PGBASEPLAFSPEC', ValeurDADS, False);

      if (SegmentNum = 'S41.G01.04.001') then
         ValeurDADS := RechDom('PGSOMISO', ValeurDADS, False);

      if (SegmentNum = 'S41.G01.06.001') then
         ValeurDADS := RechDom('PGEXONERATION', ValeurDADS, False);

      if (Segment = 'S44.G01.00.001') then
         ValeurDADS := RechDom('PGDUREETRAVAIL', ValeurDADS, False);
      end;
   if (SegmentNum = 'S45.G01.01.001') then
      ValeurDADS := RechDom('PGDADSPREV', ValeurDADS, False);
   if (SegmentNum = 'S45.G01.01.006') then
      ValeurDADS := RechDom('PGDADSBASEPREV', ValeurDADS, False);
   if (SegmentNum = 'S45.G01.01.008') then
      ValeurDADS := RechDom('PGDADSPREVPOP', ValeurDADS, False);
   if (SegmentNum = 'S45.G01.01.009') then
      ValeurDADS := RechDom('PGSITUATIONFAMIL', ValeurDADS, False);
   if (Segment = 'S66.G01.00.002') then
      ValeurDADS := RechDom('PGDUREETRAVAIL', ValeurDADS, False);
   if (Segment = 'S66.G01.00.007') then
      ValeurDADS := RechDom('PGDADSBTPSALMOY', ValeurDADS, False);
   if (Segment = 'S66.G01.00.009') then
      ValeurDADS := RechDom('PGDADSBTPHORAIRE', ValeurDADS, False);
   if (Segment = 'S66.G01.00.012') then
      ValeurDADS := RechDom('PGSITUATIONFAMIL', ValeurDADS, False);
   if (Segment = 'S66.G01.00.015') then
      ValeurDADS := RechDom('PGDADSBTPSTATUT', ValeurDADS, False);
   if (Segment = 'S66.G01.00.016') then
      ValeurDADS := RechDom('PGLIBNIVEAU', ValeurDADS, False);
   if (Segment = 'S66.G01.00.017') then
      ValeurDADS := RechDom('PGCOEFFICIENT', ValeurDADS, False);
   if (Segment = 'S66.G01.00.021') then
      ValeurDADS := RechDom('PGMETIERBTP', ValeurDADS, False);
   if (Segment = 'S66.G01.00.022') then
      ValeurDADS := RechDom('PGDADSBTPAFFILIRC', ValeurDADS, False);
   if ((Segment = '1001') or (Segment = '1008') or (Segment = '1015') or
      (Segment = '1022') or (Segment = '1029') or (Segment = '1036') or
      (Segment = '1043') or (Segment = '1050') or (Segment = '1057') or
      (Segment = '1064') or (Segment = '1071') or (Segment = '1078') or
      (Segment = '1085') or (Segment = '1092') or (Segment = '1099') or
      (Segment = '1106') or (Segment = '1113') or (Segment = '1120') or
      (Segment = '1127') or (Segment = '1134') or (Segment = '1141')) then
      ValeurDADS := RechDom('PGDADSBTPTYPEARRET', ValeurDADS, False);
   if ((Segment = '1004') or (Segment = '1011') or (Segment = '1018') or
      (Segment = '1025') or (Segment = '1032') or (Segment = '1039') or
      (Segment = '1046') or (Segment = '1053') or (Segment = '1060') or
      (Segment = '1067') or (Segment = '1074') or (Segment = '1081') or
      (Segment = '1088') or (Segment = '1095') or (Segment = '1102') or
      (Segment = '1109') or (Segment = '1116') or (Segment = '1123') or
      (Segment = '1130') or (Segment = '1137') or (Segment = '1144')) then
      ValeurDADS := RechDom('PGDUREETRAVAIL', ValeurDADS, False);
   end;

if StrToInt(ordre) = 0 then
   begin
   if (Segment = 'S30.G01.00.012') then
      ValeurDADS := RechDom('TTPAYS', ValeurDADS, False);
   if (Segment = 'S30.G01.00.013') then
      ValeurDADS := RechDom('TTPAYS', ValeurDADS, False);
   end;
   result := ValeurDADS;
   exit;
end;

function FunctPGDADSUHonoraires(sf, sp: string): variant;
var
QDADSHON: TQuery;
TOBHON: Tob;
Segment, Ordre, ValeurDads, SegmentNum: string;
begin
Ordre:= (READTOKENST(sp));
Segment:= (READTOKENST(sp));
if TOBDADShON = nil then
   begin
   QDADSHon:= OpenSQL ('SELECT PDS_ORDRE,PDS_ORDRESEG,PDS_SEGMENT,'+
                       ' PDS_DONNEEAFFICH'+
                       ' FROM DADSDETAIL WHERE'+
                       ' PDS_SALARIE LIKE "--H%"'+
                       ' ORDER BY PDS_ORDRE', True);
   TOBDADSHON:= TOB.Create('DADS Honoraires', nil, -1);
   TOBDADSHON.LoadDetailDB('DADSDETAIL', '', '', QDADSHon, False);
   Ferme(QDADSHon);
   end;
//PT71
if (isnumeric(Segment)) then
   TOBHON:= TOBDADSHON.FindFirst(['PDS_ORDRE', 'PDS_ORDRESEG'], [Ordre, Segment], False)
else
//FIN PT71
   TOBHON:= TOBDADSHON.FindFirst(['PDS_ORDRE', 'PDS_SEGMENT'], [Ordre, Segment], False);
if TOBHON <> nil then
   begin
//PT71
   if (isnumeric(Segment)) then
      SegmentNum:= TOBHON.GetValue('PDS_SEGMENT');
   ValeurDADS:= TOBHON.GetValue('PDS_DONNEEAFFICH');
   end
//FIN PT71
else
   begin
   ValeurDADS:= '';
   Result:= ValeurDADS;
   Exit;
   end;
{PT41
If (Segment='18') Then ValeurDADS:=RechDom('PGDADSHONMODALITES',ValeurDADS,False);
If (Segment='19') Then ValeurDADS:=RechDom('PGDADSHONTAUXREDUIT',ValeurDADS,False);
If (Segment='21') Then ValeurDADS:=RechDom('TTETABLISSEMENT',ValeurDADS,False);
If (Segment='22') Then ValeurDADS:=RechDom('PGEXERCICESOCIAUX',ValeurDADS,False);
If (Segment='7') Then ValeurDADS:=RechDom('PGADRESSECOMPL',ValeurDADS,False);
If (Segment='23') Then ValeurDADS:=RechDom('PGDADSHONTYPEREM',ValeurDADS,False);
If (Segment='25') Then ValeurDADS:=RechDom('PGDADSHONTYPEREM',ValeurDADS,False);
If (Segment='27') Then ValeurDADS:=RechDom('PGDADSHONTYPEREM',ValeurDADS,False);
If (Segment='29') Then ValeurDADS:=RechDom('PGDADSHONTYPEREM',ValeurDADS,False);
If (Segment='31') Then ValeurDADS:=RechDom('PGDADSHONTYPEREM',ValeurDADS,False);
}
if (Segment = 'S70.G01.00.010') then
   ValeurDADS := RechDom('PGDADSHONMODALITES', ValeurDADS, False);
if (Segment = 'S70.G01.00.011') then
   ValeurDADS := RechDom('PGDADSHONTAUXREDUIT', ValeurDADS, False);
if (Segment = 'S70.G01.00.014') then
   ValeurDADS := RechDom('TTETABLISSEMENT', ValeurDADS, False);
if (Segment = 'S70.G01.00.015') then
   ValeurDADS := RechDom('PGEXERCICESOCIAUX', ValeurDADS, False);
if (Segment = 'S70.G01.00.004.004') then
   ValeurDADS := RechDom('PGADRESSECOMPL', ValeurDADS, False);
if ((Segment = '101') or (Segment = '104') or (Segment = '107') or
   (Segment = '110') or (Segment = '113')) then
   ValeurDADS := RechDom('PGDADSHONTYPEREM', ValeurDADS, False);
//FIN PT41
result := ValeurDADS;
exit;
end;

{***************************************************************
            EDITION MOUVEMENT DE MAIN D'OEUVRE
*****************************************************************}

function FunctSortieMMO(Sf, Sp: string): Variant; // Fonction qui renvoi le motif de sortie MMO
var
  Q: TQuery;
  T: Tob;
  CodeMMO: string;
begin
  CodeMMO := READTOKENST(Sp);
  if TobSortieMMO = nil then
  begin
    Q := OpenSQL('SELECT PMS_MAINOEUVRE,PMS_CODE FROM MOTIFSORTIEPAY', True);
    TobSortieMMO := Tob.Create('Les entrées MMO', nil, -1);
    TobSortieMMO.LoadDetailDB('MOTIFSORTIEPAY', '', '', Q, False);
    Ferme(Q);
  end;
  T := TobSortieMMO.FindFirst(['PMS_CODE'], [CodeMMO], False);
  if T = nil then
  begin
    Result := '';
    Exit;
  end;
  Result := T.GetValue('PMS_MAINOEUVRE');
end;

function FunctEntreeMMO(Sf, Sp: string): Variant; // Fonction qui renvoi le motif entrée MMO
var
  Q: TQuery;
  T: Tob;
  CodeMMO: string;
begin
  CodeMMO := READTOKENST(Sp);
  if TobEntreeMMO = nil then
  begin
    Q := OpenSQL('SELECT PME_MAINOEUVRE,PME_CODE FROM MOTIFENTREESAL', True);
    TobEntreeMMO := Tob.Create('Les entrées MMO', nil, -1);
    TobEntreeMMO.LoadDetailDB('MOTIFENTREESAL', '', '', Q, False);
    Ferme(Q);
  end;
  T := TobEntreeMMO.FindFirst(['PME_CODE'], [CodeMMO], False);
  if T = nil then
  begin
    Result := '';
    Exit;
  end;
  Result := T.GetValue('PME_MAINOEUVRE');
end;

function FunctMMOEtab(Sf, Sp: string): Variant; // Fonction pour afficher l'adresse de l'établissement
var
  champ, Etab: string;
  T: Tob;
  Q: TQuery;
begin
  Champ := (READTOKENST(sp));
  Etab := READTOKENST(Sp);
  if TobMMOEtab = nil then
  begin
    Q := OpenSQL('SELECT ET_ETABLISSEMENT,ET_SIRET,ET_APE,ET_LIBELLE,ET_ADRESSE1,ET_ADRESSE2,ET_ADRESSE3,ET_CODEPOSTAL,' +
      ' ET_VILLE,ANN_APNOM,ANN_APRUE1,ANN_APRUE2,ANN_APRUE3,ANN_APCPVILLE' +
      ' FROM ETABLISS ' +
      ' left join ETABCOMPL on ET_ETABLISSEMENT=ETB_ETABLISSEMENT' +
      ' left join ANNUAIRE on ETB_CODEDDTEFPGU=ANN_GUIDPER' +
      ' WHERE ET_ETABLISSEMENT="' + PGGlbEtabMMO + '"', True);
    TobMMOEtab := TOB.Create('établissement', nil, -1);
    TobMMOEtab.LoadDetailDB('ETABLISS', '', '', Q, False);
    Ferme(Q);
  end;
  T := TobMMOEtab.FindFirst([''], [''], False);
  if T <> nil then
  begin
    if Champ = 'NOMDDTEFP' then Result := T.GetValue('ANN_APNOM');
    if Champ = 'RUE1DDTEFP' then Result := T.GetValue('ANN_APRUE1');
    if Champ = 'RUE2DDTEFP' then Result := T.GetValue('ANN_APRUE2');
    if Champ = 'RUE3DDTEFP' then Result := T.GetValue('ANN_APRUE3');
    if Champ = 'CPVILLEDDTEFP' then Result := T.GetValue('ANN_APCPVILLE');
    if Champ = 'SIRET' then Result := T.GetValue('ET_SIRET');
    if Champ = 'APE' then Result := T.GetValue('ET_APE');
    if Champ = 'ADR1' then Result := T.GetValue('ET_ADRESSE1');
    if Champ = 'ADR2' then Result := T.GetValue('ET_ADRESSE2');
    if Champ = 'ADR3' then Result := T.GetValue('ET_ADRESSE3');
    if Champ = 'CP' then Result := T.GetValue('ET_CODEPOSTAL');
    if Champ = 'VILLE' then Result := T.GetValue('ET_VILLE');
    if Champ = 'NOMETAB' then Result := T.GetValue('ET_LIBELLE');
  end
  else Result := '';
end;

function FunctDateMMO(Sf, Sp: string): Variant; // Fonction qui vérifie si sortie/entrée pour cause de changement établissement
var
  T1, T2: TOB;
  Salarie, EtabDecl, Etab1, Etab2: string;
  DebutMMO, DebutMMO2: TDateTime;
  Q: TQuery;
begin
  Salarie := (READTOKENST(SP));
  EtabDecl := (READTOKENST(Sp));
  DebutMMO := StrToDateTime(READTOKENST(sp));
  DebutMMO2 := PlusMois(DebutMMO, -1);
  if TobPaieMMO1 = nil then
  begin
    Q := OpenSQL('SELECT PPU_SALARIE,PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_DATEDEBUT="' + UsDateTime(DebutMMO) + '"', True);
    TobPaieMMO1 := TOB.Create('les paies 1', nil, -1);
    TobPaieMMO1.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
    Ferme(Q);
  end;
  if TobPaieMMO2 = nil then
  begin
    Q := OpenSQL('SELECT PPU_SALARIE,PPU_ETABLISSEMENT FROM PAIEENCOURS WHERE PPU_DATEDEBUT="' + UsDateTime(DebutMMO2) + '"', True);
    TobPaieMMO2 := TOB.Create('les paies 2', nil, -1);
    TobPaieMMO2.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
    Ferme(Q);
  end;
  T1 := TobPaieMMO1.FindFirst(['PPU_SALARIE'], [Salarie], False);
  T2 := TobPaieMMO2.FindFirst(['PPU_SALARIE'], [Salarie], False);
  if T1 = nil then
  begin
    ;
    Result := 'AUCUN';
    Exit;
  end;
  if T2 = nil then
  begin
    ;
    Result := 'AUCUN';
    Exit;
  end;
  Etab1 := T1.GetValue('PPU_ETABLISSEMENT');
  Etab2 := T2.GetValue('PPU_ETABLISSEMENT');
  if Etab1 <> Etab2 then
  begin
    if Etab1 = EtabDecl then Result := 'ENTREE';
    if Etab2 = EtabDecl then Result := 'SORTIE';
  end
  else Result := 'AUCUN';
end;

function FunctContratMMO(Sf, Sp: string): Variant; // Récupération du type de contrat de travail
var
  T: TOB;
  Q: TQuery;
  DateDebut, DateFin: TDateTime;
  Salarie: string;
begin
  Salarie := READTOKENST(Sp);
  DateDebut := StrToDate(READTOKENST(Sp));
  DateFin := StrToDate(READTOKENST(Sp));
  if TobContratMMO = nil then
  begin
    //DEBUT PT68
    Q := OpenSQL('SELECT PCI_SALARIE,PCI_TYPECONTRAT FROM CONTRATTRAVAIL WHERE PCI_DEBUTCONTRAT<="' + UsDateTime(DateFin) + '" ' +
      'AND (PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '" OR PCI_FINCONTRAT>="' + UsDateTime(DateDebut) + '") ' +
      'ORDER BY PCI_DEBUTCONTRAT DESC', True);
    //FIN PT68
    TobContratMMO := Tob.Create('Les contrats', nil, -1);
    TobContratMMO.LoadDetailDB('CONTRATTRAVAIL', '', '', Q, False);
    Ferme(Q);
  end;
  T := TobContratMMO.FindFirst(['PCI_SALARIE'], [Salarie], False);
  if T = nil then
  begin
    Result := '';
    Exit;
  end;
  result := T.GetValue('PCI_TYPECONTRAT');
end;

// MEDECINE DU TRAVAIL

function FunctMedecineAge(Sf, Sp: string): Variant; // Calcul de l'âge
var
  DateNaiss: TDateTime;
  SDateNaiss: string;
  aaaa, mm, jj, dd, oo, yyyy: Word;
begin
  SDateNaiss := READTOKENST(Sp);
  if SDateNaiss = '' then
  begin
    Result := '';
    Exit;
  end;
  DateNaiss := StrToDate(SDateNaiss);
  DecodeDate(DateNaiss, aaaa, mm, jj);
  DecodeDate(Date, yyyy, oo, dd);
  result := yyyy - aaaa;
end;

function FunctMedecineVisiteSuiv(Sf, Sp: string): Variant; // Calcul de la date de la prochaine visite (Date Visite + 1 an)
var
  DateVisite: TDateTime;
  SDateVisite: string;
  PgPeriodVisitMed : String; //PT119
  Tob_Ferie: Tob;            // PT119
  k: integer;                // PT119
begin
  SDateVisite := READTOKENST(Sp);
  if SDateVisite = '' then
  begin
    Result := '';
    Exit;
  end;
  DateVisite := StrToDate(SDateVisite);
  //DEB PT119
//  DateVisite := PlusDate(DateVisite, 2, 'A'); // Visite tous les 2 ans au lieu de 1
  PgPeriodVisitMed := VarToStr(GetParamSocSecur('SO_PGPERIODVISITMED', ''));
  if (PgPeriodVisitMed = '3') then
    DateVisite := PlusDate(DateVisite, 1, 'A');
  if (PgPeriodVisitMed = '4') then
    DateVisite := PlusDate(DateVisite, 2, 'A');
  if (PgPeriodVisitMed = '5') then
    DateVisite := PlusDate(DateVisite, 6, 'M');
  if (PgPeriodVisitMed = '6') then
    DateVisite := PlusDate(DateVisite, 3, 'M');

  if not Assigned(Tob_Ferie) then Tob_Ferie := ChargeTobFerie(DateVisite,PlusDate(DateVisite,10,'J'));
  k := 10;
  While k > 0 do
    Begin
    if (IfJourFerie(Tob_Ferie,DateVisite)) OR (DayOfWeek(DateVisite) in [1,7]) then
        Begin
        DateVisite := PlusDate(DateVisite,1,'J');
        Continue;
        End
    else
      Break;
    End;
  Tob_Ferie.Free;
  //FIN PT199
  Result := DateVisite;
end;

function FunctVisitesMed(Sf, Sp: string): Variant; // Renvoi la date de la dernière visite du salarié si il en a déja effectué
var
  T: Tob;
  Q: TQuery;
  DateVisite: TDateTime;
  Salarie: string;
begin
  Salarie := READTOKENST(sp);
  Q := OpenSQL('SELECT MAX(PVM_DATEVISITE) DATEV,PVM_SALARIE FROM VISITEMEDTRAV GROUP BY PVM_SALARIE', True); //PT66
  TobMedecine := Tob.Create('Les dernières visites', nil, -1);
  TobMedecine.LoadDetailDB('VISITEMEDTRAV', '', '', Q, False);
  Ferme(Q);
  T := TobMedecine.FindFirst(['PVM_SALARIE'], [Salarie], False);
  if T = nil then
  begin
    result := 'Aucune visite';
    Exit;
  end;
  dateVisite := T.GetValue('DATEV'); //PT66
  Result := DateToStr(DateVisite);
end;


function MSAHeuresMontant(Sf, Sp: string): Variant;
var
  Salarie, Numperiode, TypeMontant: string;
  T: Tob;
begin
  Salarie := READTOKENST(sp);
  NumPeriode := READTOKENST(Sp);
  TypeMontant := READTOKENST(Sp);
  if NumPeriode = '1' then T := TobMSAP1.FindFirst(['PPU_SALARIE'], [Salarie], false);
  if NumPeriode = '2' then T := TobMSAP2.FindFirst(['PPU_SALARIE'], [Salarie], false);
  if NumPeriode = '3' then T := TobMSAP3.FindFirst(['PPU_SALARIE'], [Salarie], false);
  if T <> nil then
  begin
    if typeMontant = 'H' then Result := T.GetValue('HEURESTRAV');
    if TypeMontant = 'R' then Result := T.GetValue('BRUT');
  end
  else Result := '';
end;

function CreatetobMSA(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Etablissement: string;
  DateDebut1, DateFin1, DateDebut2, DateFin2, DateDebut3, DateFin3: TdateTime;
begin
  Etablissement := READTOKENST(Sp);
  DateDebut1 := StrToDate(READTOKENST(Sp));
  DateDebut2 := PlusDate(DateDebut1, 1, 'M');
  DateDebut3 := PlusDate(DateDebut1, 2, 'M');
  DateFin1 := FinDeMois(DateDebut1);
  Datefin2 := FinDeMois(DateDebut2);
  DateFin3 := FinDeMois(DateDebut3);
  Q := OpenSQL('SELECT PPU_SALARIE,SUM(PPU_CBRUT) AS BRUT,SUM(PHC_MONTANT) AS HEURESTRAV' +
    ' FROM PAIEENCOURS' +
    ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"' +
    ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DateDebut1) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin1) + '"' +
    ' AND PPU_ETABLISSEMENT="' + Etablissement + '" GROUP BY PPU_SALARIE', True);
  TobMSAP1 := Tob.Create('La tob MSA', nil, -1);
  TobMSAP1.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PPU_SALARIE,SUM(PPU_CBRUT) AS BRUT,SUM(PHC_MONTANT) AS HEURESTRAV' +
    ' FROM PAIEENCOURS' +
    ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"' +
    ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DateDebut2) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin2) + '"' +
    ' AND PPU_ETABLISSEMENT="' + Etablissement + '" GROUP BY PPU_SALARIE', True);
  TobMSAP2 := Tob.Create('La tob MSA', nil, -1);
  TobMSAP2.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PPU_SALARIE,SUM(PPU_CBRUT) AS BRUT,SUM(PHC_MONTANT) AS HEURESTRAV' +
    ' FROM PAIEENCOURS' +
    ' LEFT JOIN HISTOCUMSAL ON PPU_DATEDEBUT=PHC_DATEDEBUT AND PPU_DATEFIN=PHC_DATEFIN AND PPU_SALARIE=PHC_SALARIE AND PHC_CUMULPAIE="21"' +
    ' WHERE PPU_DATEDEBUT>="' + UsDateTime(DateDebut3) + '" AND PPU_DATEFIN<="' + UsDateTime(DateFin3) + '"' +
    ' AND PPU_ETABLISSEMENT="' + Etablissement + '" GROUP BY PPU_SALARIE', True);
  TobMSAP3 := Tob.Create('La tob MSA', nil, -1);
  TobMSAP3.LoadDetailDB('PAIEENCOURS', '', '', Q, False);
  Ferme(Q);
end;

function LibDateMSA(Sf, Sp: string): Variant;
var
  periode: string;
begin
  Periode := READTOKENST(Sp);
  if Periode = 'D2' then Result := DateToStr(FinDeMois(StrToDate(READTOKENST(Sp))));
  if Periode = 'D3' then Result := DatetoStr(PlusMois(StrTodate(READTOKENST(Sp)), 1));
  if Periode = 'D4' then Result := DateToStr(FinDeMois(PlusMois(StrTodate(READTOKENST(Sp)), 1)));
  if Periode = 'D5' then Result := DateToStr(PlusMois(StrTodate(READTOKENST(Sp)), 2));
  if Periode = 'D6' then Result := DateToStr(FindeMois(PlusMois(StrTodate(READTOKENST(Sp)), 2)));
end;

function InitTobFormation(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Millesime: string;
begin
  Millesime := READTOKENST(Sp);
  Q := OpenSQL('SELECT PFI_SALARIE,PFI_NBINSC,PFI_ETABLISSEMENT,PFI_COUTREELSAL,PFI_FRAISFORFAIT FROM INSCFORMATION' +
    ' WHERE PFI_MILLESIME="' + Millesime + '"', True);
  TobInscForm := Tob.Create('Les inscriptions', nil, -1);
  TobInscform.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFO_SALARIE,PFO_FRAISLIBRE1,PFO_FRAISLIBRE2,PFO_FRAISLIBRE3,PFO_FRAISLIBRE4,' +
    'PFO_FRAISLIBRE5,PFO_FRAISLIBRE6,PFO_MILLESIME,PFO_CODESTAGE FROM FORMATIONS' +
    ' WHERE PFO_MILLESIME="' + Millesime + '"', True);
  TobFormation := Tob.Create('Les Formations', nil, -1);
  TobFormAtion.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
end;

{PT116
function InitTobFrais(Sf, Sp: string): Variant;
var
  Q: TQuery;
  //LeStage,LaSession,LeMillesime:String;
  DateDebut, DateFin: TDateTime;
  Resp, Where, TypeEdit: string;
begin
  {LeStage:=READTOKENST(Sp);
  LaSession:=READTOKENST(Sp);
  LeMillesime:=READTOKENST(Sp);
  Q:=OpenSQL('SELECT PFS_SALARIE,PFS_QUANTITE,PFS_MONTANT,PFS_MONTANTPLAF,PFS_FRAISSALFOR FROM FRAISSALFORM '+
  ' WHERE PFS_CODESTAGE="'+LeStage+'" AND PFS_ORDRE="'+LaSession+'" AND PFS_MILLESIME="'+LeMillesime+'"',True);
  TobFraisFor:=Tob.Create('FRAISSALFORM',Nil,-1);
  TobFraisFor.LoadDetailDB('FRAISSALFORM','','',Q,False);
  Ferme(Q);}
  {
  DateDebut := StrToDate(READTOKENST(Sp));
  DateFin := StrToDate(READTOKENST(Sp));
  Resp := READTOKENST(Sp);
  TypeEdit := READTOKENST(Sp);
  if TypeEdit = 'STAGIAIRE' then
  begin
    if Resp <> '' then Where := 'WHERE PFO_RESPONSFOR="' + Resp + '" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '"' //PT79
    else Where := 'WHERE PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '"';
    Q := OpenSQL('SELECT PFS_SALARIE,PFS_QUANTITE,PFS_MONTANT,PFS_MONTANTPLAF,PFS_FRAISSALFOR,PFS_ORDRE,PFS_CODESTAGE FROM FRAISSALFORM ' +
      'LEFT JOIN FORMATIONS ON PFS_SALARIE=PFO_SALARIE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_ORDRE=PFO_ORDRE AND PFS_MILLESIME=PFO_MILLESIME ' + Where, True);
    TobFraisFor := Tob.Create('Lesfrais', nil, -1);
    TobFraisFor.LoadDetailDB('LesFrais', '', '', Q, False);
    Ferme(Q);
  end
  else
  begin
    Where := 'WHERE PAN_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PAN_DATEFIN<="' + UsDateTime(DateFin) + '"';
    Q := OpenSQL('SELECT PFS_SALARIE,PFS_QUANTITE,PFS_MONTANT,PFS_MONTANTPLAF,PFS_FRAISSALFOR,PFS_ORDRE,PFS_CODESTAGE FROM FRAISSALFORM ' +
      'LEFT JOIN SESSIONANIMAT ON PFS_SALARIE=PAN_SALARIE AND PFS_CODESTAGE=PAN_CODESTAGE AND PFS_ORDRE=PAN_ORDRE AND PFS_MILLESIME=PAN_MILLESIME ' + Where, True);
    TobFraisFor := Tob.Create('Lesfrais', nil, -1);
    TobFraisFor.LoadDetailDB('LesFrais', '', '', Q, False);
    Ferme(Q);
  end;
end;

function InitLibelleFrais(Sf, Sp: string): Variant;
var
  i: Integer;
  Q: TQuery;
begin
  Q := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PFA"', True);
  i := 1;
  while not Q.Eof do
  begin
    LesFrais[i] := Q.FindField('CC_CODE').AsString;
    i := i + 1;
    Q.Next;
  end;
  Ferme(Q);
end;

function AfficheLibFrais(Sf, Sp: string): Variant;
var
  NumFrais: string;
begin
  NumFrais := READTOKENST(Sp);
  if LesFrais[StrToInt(NumFrais)] <> '' then Result := RechDom('PGFRAISSALFORM', LesFrais[StrToInt(NumFrais)], False)
  else Result := '';
end;

function AfficheFrais(Sf, Sp: string): Variant;
var
  Salarie, Frais, TypeMontant, Champ, Stage, Session: string;
  T: Tob;
begin
  {Salarie:=READTOKENST(Sp);
  Frais:=READTOKENST(Sp);
  TypeMontant:=READTOKENST(Sp);
  Champ:='';
  If TypeMontant='MR' Then Champ:='PFS_MONTANT';
  If TypeMontant='MP' Then Champ:='PFS_MONTANTPLAF';
  Frais:=LesFrais[StrToInt(Frais)];
  T:=TobFraisFor.FindFirst(['PFS_SALARIE','PFS_FRAISSALFOR'],[Salarie,Frais],False);
  If t<>Nil Then
     begin
     Result:=FloatToStr(T.GetValue(Champ));
     end
  Else Result:='0';}
  {
  Salarie := READTOKENST(Sp);
  Stage := READTOKENST(SP);
  Session := READTOKENST(Sp);
  Frais := READTOKENST(Sp);
  TypeMontant := READTOKENST(Sp);
  Champ := '';
  if TypeMontant = 'MR' then Champ := 'PFS_MONTANT';
  if TypeMontant = 'MP' then Champ := 'PFS_MONTANTPLAF';
  Frais := LesFrais[StrToInt(Frais)];
  T := TobFraisFor.FindFirst(['PFS_SALARIE', 'PFS_CODESTAGE', 'PFS_ORDRE', 'PFS_FRAISSALFOR'], [Salarie, Stage, StrToInt(Session), Frais], False);
  if t <> nil then
  begin
    Result := FloatToStr(T.GetValue(Champ));
  end
  else Result := '0';
end; }

function InitTobComparatif(Sf, Sp: string): Variant;
var
  QS, QN, QF: TQuery;
  DateFin, DateDebut: TDateTime;
  Millesime, Stage: string;
  QExer: TQuery;
begin
  Millesime := READTOKENST(Sp);
  Stage := READTOKENST(Sp);
  DateDebut := IDate1900;
  DateFin := IDate1900;
  Qexer := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
  if not Qexer.eof then
  begin
    DateDebut := QExer.FindField('PFE_DATEDEBUT').AsDateTime; // PortageCWAS
    DateFin := QExer.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(QExer);
  QF := OpenSQL('SELECT  COUNT(PFO_SALARIE) AS NBSAL, PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR FROM FORMATIONS' +
    ' WHERE PFO_CODESTAGE="' + Stage + '" AND' +
    ' PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' AND PFO_SALARIE NOT IN (SELECT PFI_SALARIE FROM INSCFORMATION WHERE PFI_CODESTAGE="' + Stage + '"' +
    ' AND PFI_MILLESIME="' + Millesime + '")' +
    ' GROUP BY PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR', true);
  TobForRealise := Tob.Create('LesFormations', nil, -1);
  TobForRealise.LoadDetailDB('FORMATIONS', '', '', QF, False);
  Ferme(QF);
  QF := OpenSQL('SELECT  PFO_SALARIE,PFO_ETABLISSEMENT,PFO_LIBEMPLOIFOR FROM FORMATIONS' +
    ' WHERE PFO_CODESTAGE="' + Stage + '" AND' +
    ' PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' AND PFO_SALARIE IN (SELECT PFI_SALARIE FROM INSCFORMATION WHERE PFI_CODESTAGE="' + Stage + '"' +
    ' AND PFI_MILLESIME="' + Millesime + '")', true);
  TobForRealiseInsc := Tob.Create('LesFormations avec insc', nil, -1);
  TobForRealiseInsc.LoadDetailDB('FORMATIONS', '', '', QF, False);
  Ferme(QF);
  QS := OpenSQL('SELECT PFI_SALARIE,PFI_NBINSC,PFI_ETABLISSEMENT,PFI_LIBEMPLOIFOR FROM INSCFORMATION' +
    ' WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" AND PFI_SALARIE<>""', True);
  TobForInscritsNom := Tob.Create('Les Salariés', nil, -1);
  TobForInscritsNom.LoadDetailDB('INSCFORMATION', '', '', QS, False);
  Ferme(QS);
  QN := OpenSQL('SELECT PFI_SALARIE,PFI_NBINSC,PFI_ETABLISSEMENT,PFI_LIBEMPLOIFOR FROM INSCFORMATION' +
    ' WHERE PFI_CODESTAGE="' + Stage + '" AND PFI_MILLESIME="' + Millesime + '" AND PFI_SALARIE=""', True);
  TobForInscritsNonNom := Tob.Create('Les inscriptions', nil, -1);
  TobForInscritsNonNom.LoadDetailDB('INSCFORMATION', '', '', QN, False);
  Ferme(QN);
end;

function FormationComparatif(Sf, Sp: string): Variant;
var
  Etablissement, LibEmploi, Salarie, Valeur: string;
  TS, TN: Tob;
  Nb: Integer;
begin
  Salarie := READTOKENST(Sp);
  Etablissement := READTOKENST(Sp);
  LibEmploi := READTOKENST(Sp);
  if Salarie <> '' then
  begin
    TS := TobForInscritsNom.FindFirst(['PFI_SALARIE'], [Salarie], False);
    if TS <> nil then
    begin
      Result := 'OUI';
      Exit;
    end;
  end;
  TN := TobForInscritsNonNom.FindFirst(['PFI_ETABLISSEMENT', 'PFI_LIBEMPLOIFOR'], [Etablissement, LibEmploi], False);
  Valeur := 'NON';
  while TN <> nil do
  begin
    Nb := TN.GetValue('PFI_NBINSC');
    if Nb > 0 then
    begin
      Nb := Nb - 1;
      TN.ChangeParent(TobForInscritsNonNom, -1);
      TN.PutValue('PFI_NBINSC', Nb);
      Valeur := 'OUI';
      Break;
    end
    else TN := TobForInscritsNonNom.FindNext(['PFI_ETABLISSEMENT', 'PFI_LIBEMPLOIFOR'], [Etablissement, LibEmploi], False);
  end;
  Result := Valeur;
end;

function InscBudgetComparatif(Sf, Sp: string): Variant;
var
  Etablissement, LibEmploi, Salarie, NbInsc: string;
  T: Tob;
  Nb, NbFor: Integer;
begin
  Salarie := READTOKENST(Sp);
  Etablissement := READTOKENST(Sp);
  LibEmploi := READTOKENST(Sp);
  NbInsc := READTOKENST(Sp);
  if Salarie <> '' then
  begin
    T := TobForRealiseInsc.FindFirst(['PFO_SALARIE'], [Salarie], False);
    if T <> nil then result := 'OUI'
    else Result := 'NON';
  end
  else
  begin
    T := TobForRealise.FindFirst(['PFO_ETABLISSEMENT', 'PFO_LIBEMPLOIFOR'], [Etablissement, LibEmploi], False);
    if T <> nil then
    begin
      NbFor := T.GetValue('NBSAL');
      if NbFor > 0 then
      begin
        T.ChangeParent(TobForRealise, -1);
        Nb := NbFor - StrToInt(NbInsc);
        T.PutValue('NBSAL', Nb);
        if Nb < 0 then
        begin
          Nb := StrToInt(NbInsc) - NbFor;
          Result := 'Reste ' + IntToStr(Nb) + ' salariés';
        end
        else Result := 'OUI';
      end
      else Result := 'NON';
    end
    else Result := 'NON';
  end;
end;

function CreateTobFormationGlobal(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Millesime: string;
  DateDebut, DateFin: TDateTime;
begin
  Millesime := READTOKENST(Sp);
  Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
  DateDebut := IDate1900;
  dateFin := IDate1900;
  if not Q.eof then
  begin
    DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime; // PortageCWAS
    DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR,SUM(PFO_NBREHEURE) NBHEURES,COUNT (PFO_SALARIE) NBSAL,' +
    'PST_CODESTAGE,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8' +
    ' FROM SESSIONSTAGE' +
    ' LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME ' +
    ' LEFT JOIN STAGE ON PST_CODESTAGE=PSS_CODESTAGE AND PST_MILLESIME=PSS_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' GROUP BY PSS_MILLESIME,PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR,' +
    'PST_CODESTAGE,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8', True);
  TobGestIndSession := Tob.Create('Les sessions', nil, -1);
  TobGestIndSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PST_MILLESIME,PST_NBSTAMAX,PST_DUREESTAGE,PST_CODESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,' +
    'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8' +
    ' FROM STAGE ' +
    'LEFT JOIN INSCFORMATION ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME WHERE PST_MILLESIME="' + Millesime + '"' +
    ' GROUP BY PST_MILLESIME,PST_NBSTAMAX,PST_DUREESTAGE,PST_CODESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,' +
    'PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8', True);
  TobGestIndStage := Tob.Create('Les Stages', nil, -1);
  TobGestIndStage.LoadDetailDB('STAGE', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFI_CODESTAGE,PFI_MILLESIME,PFI_CODESTAGE,PFI_SALARIE,PFI_NBINSC,PFI_DUREESTAGE,PFI_FRAISFORFAIT,PFI_COUTREELSAL,' +
    'PST_CODESTAGE,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8' +
    ' FROM INSCFORMATION ' +
    'LEFT JOIN STAGE ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME WHERE PFI_MILLESIME="' + Millesime + '"', True);

  TobGestIndBudget := Tob.Create('Budgeté', nil, -1);
  TobGestIndBudget.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFO_MILLESIME,PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE,SUM(PFS_MONTANT)AS FRAIS,' +
    'PST_CODESTAGE,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8' +
    ' FROM FORMATIONS ' +
    ' LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME ' +
    'LEFT JOIN FRAISSALFORM ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_MILLESIME=PFO_MILLESIME WHERE ' + //PT117
    'PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" GROUP BY PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE,' +
    'PFO_MILLESIME,PST_CODESTAGE,PST_FORMATION1,PST_FORMATION2,PST_FORMATION3,PST_FORMATION4,PST_FORMATION5,PST_FORMATION6,PST_FORMATION7,PST_FORMATION8', True);
  TobGestIndRealise := Tob.Create('Réalisé', nil, -1);
  TobGestIndRealise.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
end;

function FormationGlobal(Sf, Sp: string): Variant;
var
  Colonne, TypeBudget, ValeurChamp, Champ: string;
  TB, TR, TF, TI: tob;
  NbInscrit, NbMax, NbSta: Integer;
  NbH, NbDivE, NbDivD, NbCout, NbHTot, NbHSta: Double;
  CoutSal, CoutForfait, CoutStage, CoutUnit, CoutAnim, CoutTotal, CoutFrais: Double;
begin
  Colonne := READTOKENST(Sp);
  TypeBudget := READTOKENST(Sp);
  ValeurChamp := READTOKENST(Sp);
  Champ := READTOKENST(Sp);
  if TypeBudget = 'P' then
  begin
    if Colonne = 'NBSTA' then
    begin
      NbSta := 0;
      TI := TobGestIndBudget.FindFirst([Champ], [ValeurChamp], False);
      while TI <> nil do
      begin
        NbSta := NbSta + TI.GetValue('PFI_NBINSC');
        TI := TobGestIndBudget.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := NbSta;
    end;
    if Colonne = 'NBH' then
    begin
      NbH := 0;
      TB := TobGestIndStage.FindFirst([Champ], [ValeurChamp], False);
      while TB <> nil do
      begin
        NbH := NbH + TB.GetValue('PST_DUREESTAGE');
        TB := TobGestIndStage.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := NbH;
    end;
    if Colonne = 'COUT' then
    begin
      TB := TobGestIndStage.FindFirst([Champ], [ValeurChamp], False);
      while TB <> nil do
      begin
        Couttotal := 0;
        NbMax := TB.GetValue('PST_NBSTAMAX');
        NbInscrit := TB.GetValue('NBINSC');
        if NbMax > 0 then
        begin
          NbDivD := arrondi(NbInscrit / NbMax, 2);
          NbDivE := arrondi(NbDivD, 0);
          if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
          else NbCout := NbDivE
        end
        else NbCout := 1;
        TI := TobGestIndBudget.FindFirst(['PFI_CODESTAGE', 'PFI_MILLESIME'], [TB.getValue('PST_CODESTAGE'), TB.getValue('PST_MILLESIME')], False);
        while TI <> nil do
        begin
          CoutSal := TI.GetValue('PFI_COUTREELSAL');
          CoutForfait := TI.GetValue('PFI_FRAISFORFAIT');
          CoutTotal := CoutTotal + CoutSal + CoutForfait;
          TI := TobGestIndBudget.FindNext(['PFI_CODESTAGE', 'PFI_MILLESIME'], [TB.getValue('PST_CODESTAGE'), TB.getValue('PST_MILLESIME')], False);
        end;
        CoutAnim := TB.GetValue('PST_COUTSALAIR');
        CoutAnim := CoutAnim * NbCout;
        CoutStage := TB.GetValue('PST_COUTBUDGETE');
        CoutStage := CoutStage * NbCout;
        CoutUnit := TB.GetValue('PST_COUTUNITAIRE');
        CoutUnit := CoutUnit * NbInscrit;
        CoutTotal := CoutTotal + CoutUnit + CoutAnim + CoutStage;
        TB := TobGestIndStage.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := CoutTotal;
    end;
  end;
  if TypeBudget = 'R' then
  begin
    if Colonne = 'NBSTA' then
    begin
      NbSta := 0;
      TF := TobGestIndRealise.FindFirst([Champ], [ValeurChamp], False);
      while TF <> nil do
      begin
        NbSta := NbSta + 1;
        TF := TobGestIndRealise.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := NbSta;
    end;
    if Colonne = 'NBH' then
    begin
      NbH := 0;
      TF := TobGestIndRealise.FindFirst([Champ], [ValeurChamp], False);
      while TF <> nil do
      begin
        NbH := NbH + TF.GetValue('PFO_NBREHEURE');
        TF := TobGestIndRealise.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := NbH;
    end;
    if Colonne = 'COUT' then
    begin
      CoutTotal := 0;
      TR := TobGestIndSession.FindFirst([Champ], [ValeurChamp], False);
      while TR <> nil do
      begin
        NbInscrit := TR.GetValue('NBSAL');
        NbHTot := TR.GetValue('NBHEURES');
        TF := TobGestIndRealise.FindFirst(['PFO_CODESTAGE', 'PFO_ORDRE', 'PFO_MILLESIME'], [TR.GetValue('PSS_CODESTAGE'), TR.GetValue('PSS_ORDRE'), TR.GetValue('PSS_MILLESIME')], False);
        while TF <> nil do
        begin
{$IFDEF EAGLCLIENT}
          if TF.GetValue('FRAIS') <> Null then
          begin
            if IsNumeric(TF.GetValue('FRAIS')) then CoutFrais := TF.GetValue('FRAIS')
            else CoutFrais := 0;
          end
          else CoutFrais := 0;
{$ELSE}
          if TF.GetValue('FRAIS') <> Null then CoutFrais := TF.GetValue('FRAIS')
          else CoutFrais := 0;
{$ENDIF}
          CoutSal := TF.GetValue('PFO_COUTREELSAL');
          NbHSta := NbHSta + TF.GetValue('PFO_NBREHEURE');
          CoutTotal := CoutTotal + CoutFrais + CoutSal;
          TF := TobGestIndRealise.FindNext(['PFO_CODESTAGE', 'PFO_ORDRE', 'PFO_MILLESIME'], [TR.GetValue('PSS_CODESTAGE'), TR.GetValue('PSS_ORDRE'), TR.GetValue('PSS_MILLESIME')], False);
        end;
        CoutStage := TR.GetValue('PSS_COUTPEDAG');
        CoutStage := Arrondi((CoutStage * NbHSta) / NbHTot, 2);
        CoutAnim := TR.GetValue('PSS_COUTSALAIR');
        CoutAnim := Arrondi((CoutAnim * NbHSta) / NbHTot, 2);
        CoutUnit := TR.GetValue('PSS_COUTUNITAIRE');
        CoutUnit := CoutUnit * NbInscrit;
        CoutTotal := CoutTotal + CoutUnit + CoutAnim + CoutStage;
        TR := TobGestIndSession.FindNext([Champ], [ValeurChamp], False);
      end;
      Result := CoutTotal;
    end;
  end;
end;

function CreateTobForRealise(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Axe, Millesime: string;
  DateDebut, DateFin: TDateTime;
begin
  Millesime := READTOKENST(Sp);
  Axe := READTOKENST(Sp);
  DateDebut := IDate1900;
  DateFin := IDate1900;
  MillesimeFormation := Millesime;
  Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
  if not Q.eof then
  begin
    DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime; // PortageCWAS
    DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PFO_COUTREELSAL) as COUTSAL,MONTH(PFO_DATEFIN) AS MOIS,' + Axe + ' FROM FORMATIONS ' +
    'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    ' GROUP BY MONTH(PFO_DATEFIN),' + Axe, True);
  TobFGFormations := TOB.Create('Les formations', nil, -1);
  TobFGFormations.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PFS_MONTANT)AS FRAIS,MONTH(PFO_DATEFIN) AS MOIS,' + Axe +
    ' FROM  FRAISSALFORM ' +
    'LEFT JOIN FORMATIONS ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_MILLESIME=PFO_MILLESIME AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE ' +//PT117
    'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'GROUP BY MONTH(PFO_DATEFIN),' + Axe, True);
  TobFGFrais := TOB.Create('Les Frais', nil, -1);
  TobFGFrais.LoadDetailDB('FRAISSALFORM', '', '', Q, False);
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PSS_COUTPEDAG) AS COUTPEDAG,SUM(PSS_COUTSALAIR) AS COUTANIM,MONTH(PSS_DATEFIN) AS MOIS,' + Axe +
    ' FROM SESSIONSTAGE ' +
    'LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME ' +
    'WHERE PSS_EFFECTUE="X" AND PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'GROUP BY MONTH(PSS_DATEFIN),' + Axe, True);
  TobFGSessions := TOB.Create('Le réalisé', nil, -1);
  TobFGSessions.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
end;

function RealiseForMensuel(Sf, Sp: string): Variant;
var
  TF, TR, TS, TCUnit: Tob;
  Axe, Valeur, Mois, TypeC: string;
  MOntant, CoutUnitaire: Double;
  DDForCoutUnit, DFForCoutUnit: TDateTime;
  Q: TQuery;
  i: Integer;
begin
  Valeur := READTOKENST(Sp);
  Mois := READTOKENST(Sp);
  TypeC := READTOKENST(Sp);
  Axe := READTOKENST(Sp);
  Montant := 0;
  if TypeC = 'SAL' then
  begin
    TF := TobFGFormations.FindFirst(['MOIS', Axe], [Mois, Valeur], False);
    TR := TobFGFrais.FindFirst(['MOIS', Axe], [Mois, Valeur], False);
    if TF <> nil then Montant := TF.GetValue('COUTSAL');
    if TR <> nil then
    begin
{$IFDEF EAGLCLIENT}
      if TR.GetValue('FRAIS') <> Null then
      begin
        if IsNumeric(TR.GetValue('FRAIS')) then Montant := Montant + TR.GetValue('FRAIS');
      end;
{$ELSE}
      if TR.GetValue('FRAIS') <> Null then Montant := Montant + TR.GetValue('FRAIS');
{$ENDIF}
    end;
  end;
  if TypeC = 'PEDAG' then
  begin
    TS := TobFGSessions.FindFirst(['MOIS', Axe], [Mois, Valeur], False);
    if TS <> nil then Montant := TS.GetValue('COUTPEDAG');
    if TS <> nil then Montant := Montant + TS.GetValue('COUTANIM');
    DDForCoutUnit := EncodeDate(StrToInt(MillesimeFormation), StrToInt(Mois), 01);
    DFForCoutUnit := FinDeMois(DDForCoutUnit);
    Q := OpenSQL('SELECT PSS_COUTUNITAIRE * Count(PFO_SALARIE) TOTALUNIT FROM SESSIONSTAGE LEFT JOIN FORMATIONS ' +
      'ON PFO_MILLESIME=PSS_MILLESIME AND PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE ' +
      'LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME ' +
      'WHERE PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DDForCoutUnit) + '" AND PFO_DATEFIN<="' + UsDateTime(DFForCoutUnit) + '" ' +
      'AND ' + Axe + '="' + Valeur + '" ' +
      'GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_MILLESIME,PSS_COUTUNITAIRE', true);
    TCUnit := Tob.Create('les couts unitaires', nil, -1);
    TCUnit.LoadDetailDB('les couts', '', '', Q, False);
    Ferme(Q);
    CoutUnitaire := 0;
    for i := 0 to TCUnit.Detail.Count - 1 do
    begin
      CoutUnitaire := CoutUnitaire + TCUnit.Detail[i].GetValue('TOTALUNIT');
    end;
    TCUnit.Free;
    Montant := Montant + CoutUnitaire;
  end;
  if Montant <> 0 then Result := Montant
  else Result := '';
end;

function CreateTobSuviRespF(Sf, Sp: string): Variant;
var
  Millesime, RuptureStage, GroupBy, Champ: string;
  DateDebut, DateFin: TDateTime;
  Q: TQuery;
begin
  Millesime := READTOKENST(SP);
  RuptureStage := READTOKENST(SP);
  if RuptureStage = 'X' then GroupBy := ',PST_CODESTAGE'
  else GroupBy := '';
  Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
  Datedebut := IDate1900;
  DateFin := IDate1900;
  if not Q.eof then
  begin
    DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime; // PortageCWAS
    DateFin := Q.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PFS_MONTANT)AS FRAIS,PFO_RESPONSFOR' + GroupBy +
    ' FROM  FRAISSALFORM ' +
    'LEFT JOIN STAGE ON PST_CODESTAGE=PFS_CODESTAGE AND PST_MILLESIME=PFS_MILLESIME ' +
    'LEFT JOIN FORMATIONS ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_MILLESIME=PFO_MILLESIME AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE ' + //PT117
    'WHERE PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" ' +
    'GROUP BY PFO_RESPONSFOR' + GroupBy, True);
  TobFGFrais := TOB.Create('Les Frais', nil, -1);
  TobFGFrais.LoadDetailDB('FRAISSALFORM', '', '', Q, False);
  Ferme(Q);

  Q := OpenSQl('SELECT SUM(PFI_COUTREELSAL) AS COUTSAL,SUM(PFI_FRAISFORFAIT) AS FORFAIT,SUM(PFI_NBINSC) AS NBINSC,SUM(PFI_DUREESTAGE) AS DUREE,PFI_RESPONSFOR' + groupBy +
    ' FROM INSCFORMATION ' +
    'LEFT JOIN  STAGE ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PFI_MILLESIME ' +
    'WHERE PFI_MILLESIME="' + Millesime + '" AND PFI_DATEACCEPT<>"' + UsDateTime(IDate1900) + '" ' +
    'GROUP BY PFI_RESPONSFOR' + GroupBy, True);
  TobFGInsc := TOB.Create('Les inscriptions', nil, -1);
  TobFGInsc.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PFO_COUTREELSAL) AS COUTSAL,COUNT (PFO_SALARIE) AS NBSAL,SUM(PFO_NBREHEURE) AS NBHEURE,PFO_RESPONSFOR' + GroupBy + ' ' +
    'FROM  FORMATIONS ' +
    'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME ' +
    'LEFT JOIN SESSIONSTAGE ON PSS_CODESTAGE=PFO_CODESTAGE AND PSS_MILLESIME=PFO_MILLESIME AND PSS_ORDRE=PFO_ORDRE ' +
    'WHERE PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" AND PFO_EFFECTUE="X" ' +
    ' AND PSS_EFFECTUE="X" ' +
    'GROUP BY PFO_RESPONSFOR' + GroupBy, True);
  TobFGFormations := TOB.Create('Les formations', nil, -1);
  TobFGFormations.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);

  if GroupBy <> '' then
  begin
    Champ := GroupBy;
    GroupBy := 'GROUP BY PST_CODESTAGE';
  end;
  Q := OpenSQl('SELECT SUM(PST_COUTBUDGETE) AS COUTBUDGETE,SUM(PST_COUTSALAIR) AS COUTANIM' + Champ + ' ' +
    'FROM STAGE ' +
    'LEFT JOIN INSCFORMATION ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PFI_MILLESIME ' +
    'WHERE PST_MILLESIME="' + Millesime + '" AND PST_ACCEPTBUD="X" ' +
    GroupBy, True);
  TobFGStage := TOB.Create('Le budget', nil, -1);
  TobFGStage.LoadDetailDB('STAGE', '', '', Q, False);
  Ferme(Q);

  Q := OpenSQL('SELECT SUM(PSS_COUTPEDAG) AS COUTPEDAG,SUM(PSS_COUTSALAIR) AS COUTANIM' + Champ + ' ' +
    'FROM SESSIONSTAGE ' +
    'LEFT JOIN STAGE ON PSS_CODESTAGE=PST_CODESTAGE AND PSS_MILLESIME=PST_MILLESIME ' +
    'WHERE PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '" AND PSS_EFFECTUE="X" ' +
    GroupBy, True);
  TobFGSessions := TOB.Create('Le réalisé', nil, -1);
  TobFGSessions.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);

end;

function SuiviResponsableFormation(Sf, Sp: string): Variant;
var
  TF, TI, TR: Tob;
  Resp, Stage, TypeM, TypeF: string;
  Montant: Double;
begin
  TypeF := READTOKENST(Sp);
  TypeM := READTOKENST(Sp);
  Resp := READTOKENST(Sp);
  Stage := READTOKENST(Sp);
  Montant := 0;
  if TypeF = 'P' then
  begin
    if TypeM = 'H' then
    begin
      if Stage <> '' then TI := TobFGInsc.Findfirst(['PFI_RESPONSFORPFI_RESPONSFOR', 'PST_CODESTAGE'], [Resp, Stage], False)
      else TI := TobFGInsc.Findfirst(['PFI_RESPONSFOR'], [Resp], False);
      if TI <> nil then
      begin
        Montant := TI.GetValue('DUREE');
      end;
    end;
    if TypeM = 'C' then
    begin
      if Stage <> '' then TI := TobFGInsc.Findfirst(['PFI_RESPONSFOR', 'PST_CODESTAGE'], [Resp, Stage], False)
      else TI := TobFGInsc.Findfirst(['PFI_RESPONSFOR'], [Resp], False);
      if TI <> nil then
      begin
        Montant := TI.GetValue('COUTSAL');
        Montant := Montant + TI.GetValue('FORFAIT');
      end;
    end;
  end;
  if TypeF = 'R' then
  begin
    if TypeM = 'H' then
    begin
      if Stage <> '' then TF := TobFGFormations.FindFirst(['PFO_RESPONSFOR', 'PST_CODESTAGE'], [Resp, Stage], False)
      else TF := TobFGFormations.FindFirst(['PFO_RESPONSFOR'], [Resp], False);
      if TF <> nil then Montant := TF.GetValue('NBHEURE');
    end;
    if TypeM = 'C' then
    begin
      if Stage <> '' then
      begin
        TF := TobFGFormations.FindFirst(['PFO_RESPONSFOR', 'PST_CODESTAGE'], [Resp, Stage], False);
        TR := TobFGFrais.FindFirst(['PFO_RESPONSFOR', 'PST_CODESTAGE'], [Resp, Stage], False);
      end
      else
      begin
        TF := TobFGFormations.FindFirst(['PFO_RESPONSFOR'], [Resp], False);
        TR := TobFGFrais.FindFirst(['PFO_RESPONSFOR'], [Resp], False);
      end;
      if TR <> nil then
      begin
{$IFDEF EAGLCLIENT}
        if TR.GetValue('FRAIS') <> Null then
        begin
          if IsNumeric(TR.GetValue('FRAIS')) then Montant := TR.GetValue('FRAIS');
        end;
{$ELSE}
        if TR.GetValue('FRAIS') <> Null then Montant := TR.GetValue('FRAIS');
{$ENDIF}
      end;
      if TF <> nil then Montant := Montant + TF.GetValue('COUTSAL');
    end;
  end;
  Result := Montant;
end;

function CreateTobBudgetFor(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Millesime: string;
begin
  Millesime := READTOKENST(Sp);
  Q := OpenSQL('SELECT PST_CODESTAGE,PST_MILLESIME,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,SUM(PFI_NBINSC) AS NBINSC' +
    ' FROM STAGE LEFT JOIN INSCFORMATION ON PST_CODESTAGE=PFI_CODESTAGE AND PST_MILLESIME=PST_MILLESIME ' +
    'WHERE PST_MILLESIME="' + Millesime + '"' +
    ' GROUP BY PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR,PST_NBSTAMAX,PST_CODESTAGE,PST_MILLESIME', True);
  TobCoutStageBud := Tob.Create('Le couts du budget', nil, -1);
  TobCoutStageBud.LoadDetailDB('STAGE', '', '', Q, False);
  Ferme(Q);
end;

function CalCoutBudgetFor(Sf, Sp: string): Variant;
var
  T: Tob;
  Stage, Millesime, NbInsc: string;
  NbInscStage, NbMax: Integer;
  CoutFonc, CoutUnit, CoutAnim, NbDivD, NbDivE, NbCout: Double;
begin
  Stage := READTOKENST(Sp);
  Millesime := READTOKENST(Sp);
  NbInsc := READTOKENST(Sp);
  CoutUnit := 0;
  CoutAnim := 0;
  Coutfonc := 0;
  T := TobCoutStageBud.FindFirst(['PST_CODESTAGE', 'PST_MILLESIME'], [Stage, Millesime], False);
  if T <> nil then
  begin
    CoutFonc := T.GetValue('PST_COUTBUDGETE');
    CoutUnit := T.GetValue('PST_COUTUNITAIRE');
    CoutAnim := T.GetValue('PST_COUTSALAIR');
    NbMax := T.GetValue('PST_NBSTAMAX');
    NbInscStage := T.GetValue('NBINSC');
    if NbMax > 0 then
    begin
      NbDivD := NbInscStage / NbMax;
      NbDivE := arrondi(NbDivD, 0);
      if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
      else NbCout := NbDivE
    end
    else NbCout := 1;
    CoutFonc := (NbCout * CoutFonc) / NbInscStage;
    CoutAnim := (NbCout * CoutAnim) / NbInscStage;
  end;

  Result := (StrToInt(NbInsc)) * (coutFonc + CoutAnim + Coutunit);
end;


function CreateTobGestIndFor(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Millesime: string;
  DateDebut, DateFin: TDateTime;
begin
  Millesime := READTOKENST(Sp);
  Q := OpenSQL('SELECT PFE_DATEDEBUT,PFE_DATEFIN FROM EXERFORMATION WHERE PFE_MILLESIME="' + Millesime + '"', True);
  DateDebut := IDate1900;
  DateFin := IDate1900;
  if not Q.eof then
  begin
    DateDebut := Q.FindField('PFE_DATEDEBUT').AsDateTime; // PortageCWAS
    Datefin := Q.FindField('PFE_DATEFIN').AsDateTime;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT PFI_SALARIE,PFI_NBINSC,PFI_DUREESTAGE,PFI_FRAISFORFAIT,PFI_COUTREELSAL FROM INSCFORMATION WHERE PFI_MILLESIME="' + Millesime + '"', True);
  TobGestIndBudget := Tob.Create('Budgeté', nil, -1);
  TobGestIndBudget.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE,SUM(PFS_MONTANT)AS FRAIS FROM FORMATIONS ' +
    'LEFT JOIN FRAISSALFORM ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_MILLESIME=PFO_MILLESIME WHERE ' +//PT117
    'PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" AND PFO_EFFECTUE="X" GROUP BY PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE', True);
  TobGestIndRealise := Tob.Create('Réalisé', nil, -1);
  TobGestIndRealise.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PST_DUREESTAGE,PST_CODESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR FROM INSCFORMATION WHERE PST_MILLESIME="' + Millesime + '"', True);
  TobGestIndStage := Tob.Create('Les Stages', nil, -1);
  TobGestIndStage.LoadDetailDB('STAGE', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR FROM FORMATIONS WHERE PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '"', True);
  TobGestIndSession := Tob.Create('Les sessions', nil, -1);
  TobGestIndSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
end;

function GestIndFormation(Sf, Sp: string): Variant;
var
  TSession, TStage, TInsc, TForm: Tob;
  Salarie, Stage, Millesime: string;
  TypeMontant, Colonne: string;
  Session: Integer;
  NbStagiaire, NbInsc: Integer;
  NbheuresTot, NbHeuresSta: Double;
  TotalPrev, TotalRea: Double;
  CoutSal, CoutUnit, CoutStage, CoutAnim, CoutFrais, NbDivD, NbDivE, NbCout: Double;
begin
  Salarie := READTOKENST(Sp);
  Stage := READTOKENST(Sp);
  Millesime := READTOKENST(Sp);
  Colonne := READTOKENST(Sp);
  TypeMontant := READTOKENST(Sp);
  TotalRea := 0;
  TotalPrev := 0;
  if TypeMontant = 'H' then
  begin
    TInsc := TobGestIndBudget.FindFirst(['PFI_SALARIE', 'PFI_CODESTAGE'], [Salarie, Stage], False);
    if TInsc <> nil then
    begin
      TStage := TobGestIndStage.FindFirst(['PST_CODESTAGE'], [Stage], False);
      if TStage <> nil then TotalPrev := TStage.GetValue('PST_DUREESTAGE');
    end;
    TForm := TobGestIndRealise.FindFirst(['PFO_SALARIE', 'PFO_CODESTAGE'], [Salarie, Stage], False);
    if TForm <> nil then
    begin
      TotalRea := TForm.GetValue('PFO_NBREHEURE');
    end;
  end;
  if TypeMontant = 'C' then
  begin
    TInsc := TobGestIndBudget.FindFirst(['PFI_SALARIE', 'PFI_CODESTAGE'], [Salarie, Stage], False);
    TStage := TobGestIndStage.FindFirst(['PST_CODESTAGE'], [Stage], False);
    if (Tstage <> nil) and (TInsc <> nil) then
    begin
      TStage := TobGestIndStage.FindFirst(['PST_CODESTAGE'], [Stage], False);
      NbInsc := TStage.GetValue('NBINSC');
      NbStagiaire := TStage.GetValue('PST_NBSTAMAX');
      if NbStagiaire > 0 then
      begin
        NbDivD := NbInsc / NbStagiaire;
        NbDivE := arrondi(NbDivD, 0);
        if NbDivE - NbDivD < 0 then NbCout := NbDivE + 1
        else NbCout := NbDivE
      end
      else NbCout := 1;
      CoutAnim := TStage.GetValue('PST_COUTSALAIR');
      CoutAnim := CoutAnim * NbCout;
      CoutAnim := CoutAnim / NbInsc;
      CoutStage := TStage.GetValue('PST_COUTBUDGETE');
      CoutStage := CoutStage * NbCout;
      CoutStage := CoutStage / NbInsc;
      CoutUnit := TStage.GetValue('PST_COUTUNITAIRE');
      CoutFrais := TInsc.GetValue('PFI_FRAISFORFAIT');
      CoutSal := TInsc.GetValue('PFI_COUTREELSAL');
      TotalPrev := CoutAnim + CoutStage + CoutUnit + CoutFrais + CoutSal;
    end;
    TForm := TobGestIndRealise.FindFirst(['PFO_SALARIE', 'PFO_CODESTAGE'], [Salarie, Stage], False);
    if TForm <> nil then
    begin
      Session := TForm.GetValue('PFO_ORDRE');
      TSession := TobGestIndSession.FindFirst(['PSS_CODESTAGE', 'PSS_ORDRE'], [Stage, Session], False);
      NbHeuresTot := TSession.GetValue('NBHEURES');
      CoutAnim := TSession.GetValue('PSS_COUTSALAIR');
      CoutAnim := CoutAnim / NbHeuresTot;
      CoutStage := TSession.GetValue('PSS_COUTPEDAG');
      CoutStage := CoutStage / NbHeuresTot;
      CoutUnit := TSession.GetValue('PSS_COUTUNITAIRE');
{$IFDEF EAGLCLIENT}
      if TForm.GetValue('FRAIS') <> Null then
      begin
        if IsNumeric(TForm.GetValue('FRAIS')) then CoutFrais := TForm.GetValue('FRAIS')
        else CoutFrais := 0;
      end
      else CoutFrais := 0;
{$ELSE}
      if TForm.GetValue('FRAIS') <> Null then CoutFrais := TForm.GetValue('FRAIS')
      else CoutFrais := 0;
{$ENDIF}
      CoutSal := TForm.GetValue('PFO_COUTREELSAL');
      NbHeuresSta := TForm.GetValue('PFO_NBREHEURE');
      CoutStage := Arrondi(CoutStage * NbHeuresSta, 2);
      CoutAnim := Arrondi(CoutAnim * NbHeuresSta, 2);
      TotalRea := CoutAnim + CoutStage + CoutUnit + CoutFrais + CoutSal;
    end;
  end;
  if Colonne = 'REA' then
  begin
    if (TotalPrev <> 0) and (TotalRea <> 0) then
    begin
      if TotalPrev >= TotalRea then Result := TotalRea
      else Result := TotalPrev;
    end
    else Result := 0;
  end;
  if Colonne = 'NREA' then
  begin
    if TotalPrev > TotalRea then Result := TotalPrev - TotalRea
    else Result := 0;
  end;
  if Colonne = 'NPREV' then
  begin
    if TotalRea > TotalPrev then Result := TotalRea - TotalPrev
    else Result := 0;
  end;
end;

function InitTobSuiviGestInd(Sf, Sp: string): Variant;
var
  Q: TQuery;
  DateDebut, DateFin: TDateTime;
begin
  DateDebut := StrToDate(READTOKENST(Sp));
  DateFin := StrToDate(READTOKENST(Sp));
  Q := OpenSQL('SELECT PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE,SUM(PFS_MONTANT) FRAIS FROM FORMATIONS ' +
    'LEFT JOIN FRAISSALFORM ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_MILLESIME=PFO_MILLESIME WHERE ' +//PT117
    'PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" GROUP BY PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE', True);
  TobGestIndRealise := Tob.Create('Réalisé', nil, -1);
  TobGestIndRealise.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR,SUM (PFO_NBREHEURE) NBHEURES FROM SESSIONSTAGE' +
    ' LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR', True);
  TobGestIndSession := Tob.Create('Les sessions', nil, -1);
  TobGestIndSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
end;

function SuiviGestFor(Sf, Sp: string): Variant;
var
  Stage, Session, Salarie: string;
  TS, TF: Tob;
  NBheuresSal, NBHeuresTot: Double;
  CoutSession, CoutAnim, CoutUnit, CoutFrais, CoutSal: Double;
begin
  Stage := READTOKENST(Sp);
  Session := READTOKENST(Sp);
  Salarie := READTOKENST(Sp);
  CoutUnit := 0;
  CoutSession := 0;
  CoutAnim := 0;
  TS := TobGestIndSession.FindFirst(['PSS_CODESTAGE', 'PSS_ORDRE'], [Stage, Session], False);
  if TS <> nil then
  begin
    NBHeuresTot := TS.GetValue('NBHEURES');
    CoutUnit := TS.GetValue('PSS_COUTUNITAIRE');
    CoutAnim := (TS.GetValue('PSS_COUTSALAIR')) / NBHeuresTot;
    CoutSession := (TS.GetValue('PSS_COUTPEDAG')) / NBHeuresTot;
  end;
  TF := TobGestIndRealise.FindFirst(['PFO_SALARIE', 'PFO_CODESTAGE', 'PFO_ORDRE'], [Salarie, Stage, Session], False);
  if TF <> nil then
  begin
{$IFDEF EAGLCLIENT}
    if TF.GetValue('FRAIS') <> Null then
    begin
      if IsNumeric(TF.GetValue('FRAIS')) then CoutFrais := TF.GetValue('FRAIS')
      else CoutFrais := 0;
    end
    else CoutFrais := 0;
{$ELSE}
    if TF.GetValue('FRAIS') <> Null then CoutFrais := TF.GetValue('FRAIS')
    else CoutFrais := 0;
{$ENDIF}
    NBheuresSal := TF.GetValue('PFO_NBREHEURE');
    CoutSal := TF.GetValue('PFO_COUTREELSAL');
    CoutAnim := Arrondi(CoutAnim * NBheuresSal, 2);
    CoutSession := Arrondi(CoutSession * NBheuresSal, 2);
  end;
  Result := CoutFrais + CoutSal + CoutSession + CoutUnit + CoutAnim;
end;

function InitTobGestIndComp(Sf, Sp: string): Variant;
var
  Q: TQuery;
  Millesime: string;
  DateDebut, DateFin: TDateTime;
begin
  Millesime := READTOKENST(Sp);
  DateDebut := StrToDate(READTOKENST(Sp));
  DateFin := StrToDate(READTOKENST(Sp));
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR,SUM (PFO_NBREHEURE) NBHEURES FROM SESSIONSTAGE' +
    ' LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR', True);
  TobGestIndSession := Tob.Create('Les sessions', nil, -1);
  TobGestIndSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT SUM(PFI_NBINSC) AS NBINSC,PST_NBSTAMAX,PST_DUREESTAGE,PST_CODESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR FROM STAGE ' +
    'LEFT JOIN INSCFORMATION ON PFI_CODESTAGE=PST_CODESTAGE AND PFI_MILLESIME=PST_MILLESIME WHERE PST_MILLESIME="' + Millesime + '"' +
    ' GROUP BY PST_NBSTAMAX,PST_DUREESTAGE,PST_CODESTAGE,PST_COUTBUDGETE,PST_COUTUNITAIRE,PST_COUTSALAIR', True);
  TobGestIndStage := Tob.Create('Les Stages', nil, -1);
  TobGestIndStage.LoadDetailDB('STAGE', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFI_CODESTAGE,PFI_SALARIE,PFI_NBINSC,PFI_DUREESTAGE,PFI_FRAISFORFAIT,PFI_COUTREELSAL FROM INSCFORMATION WHERE PFI_MILLESIME="' + Millesime + '"', True);
  TobGestIndBudget := Tob.Create('Budgeté', nil, -1);
  TobGestIndBudget.LoadDetailDB('INSCFORMATION', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE,SUM(PFS_MONTANT)AS FRAIS FROM FORMATIONS ' +
    'LEFT JOIN FRAISSALFORM ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_MILLESIME=PFO_MILLESIME WHERE ' +//PT117
    'PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" GROUP BY PFO_NBREHEURE,PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,PFO_CODESTAGE', True);
  TobGestIndRealise := Tob.Create('Réalisé', nil, -1);
  TobGestIndRealise.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
end;

function InitTobForSuiviGestColl(Sf, Sp: string): Variant;
var
  DateDebut, DateFin: TDateTime;
  Q: TQuery;
  SQL: string;
begin
  DateDebut := StrToDate(READTOKENST(Sp));
  DateFin := StrToDate(READTOKENST(Sp));
  ChampFormation1 := READTOKENST(Sp);
  ChampFormation2 := READTOKENST(Sp);
  ChampFormation3 := READTOKENST(Sp);
  ChampFormation4 := READTOKENST(Sp);
  SQl := ChampFormation1;
  if ChampFormation2 <> '' then SQL := SQL + ',' + ChampFormation2;
  if ChampFormation3 <> '' then SQL := SQL + ',' + ChampFormation3;
  if ChampFormation4 <> '' then SQL := SQL + ',' + ChampFormation4;
  if (ChampFormation1 <> 'PFO_CODESTAGE') and (ChampFormation2 <> 'PFO_CODESTAGE') and (ChampFormation3 <> 'PFO_CODESTAGE') and (ChampFormation4 <> 'PFO_CODESTAGE') then SQL := SQL + ',PFO_CODESTAGE';
  Q := OpenSQL('SELECT PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,SUM(PFS_MONTANT)AS FRAIS,' + SQL + ' FROM FORMATIONS ' +
    'LEFT JOIN FRAISSALFORM ON PFS_TYPEPLANPREV=PFO_TYPEPLANPREV AND PFS_SALARIE=PFO_SALARIE AND PFS_ORDRE=PFO_ORDRE AND PFS_CODESTAGE=PFO_CODESTAGE AND PFS_MILLESIME=PFO_MILLESIME ' +//PT117
    'LEFT JOIN STAGE ON PFO_CODESTAGE=PST_CODESTAGE AND PFO_MILLESIME=PST_MILLESIME WHERE ' +
    'PFO_EFFECTUE="X" AND PFO_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PFO_DATEFIN<="' + UsDateTime(DateFin) + '" GROUP BY PFO_SALARIE,PFO_COUTREELSAL,PFO_ORDRE,' + SQL, True);
  TobGestIndRealise := Tob.Create('Réalisé', nil, -1);
  TobGestIndRealise.LoadDetailDB('FORMATIONS', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR,COUNT (PFO_SALARIE) AS NBSAL FROM SESSIONSTAGE' +
    ' LEFT JOIN FORMATIONS ON PFO_CODESTAGE=PSS_CODESTAGE AND PFO_ORDRE=PSS_ORDRE AND PFO_MILLESIME=PSS_MILLESIME ' +
    'WHERE PFO_EFFECTUE="X" AND PSS_DATEDEBUT>="' + UsDateTime(DateDebut) + '" AND PSS_DATEFIN<="' + UsDateTime(DateFin) + '"' +
    ' GROUP BY PSS_CODESTAGE,PSS_ORDRE,PSS_COUTPEDAG,PSS_COUTUNITAIRE,PSS_COUTSALAIR', True);
  TobGestIndSession := Tob.Create('Les sessions', nil, -1);
  TobGestIndSession.LoadDetailDB('SESSIONSTAGE', '', '', Q, False);
  Ferme(Q);
end;

function ForSuiviGestColl(Sf, Sp: string): Variant;
var
  Val1, Val2, Val3, Val4: string;
  TF, TS: Tob;
  CoutUnit, CoutSession, CoutAnim: Double;
  NbSta: Integer;
  Session, Stage: string;
  CoutTotal, CoutSal, CoutFrais: double;
begin
  Val1 := READTOKENST(Sp);
  Val2 := READTOKENST(Sp);
  Val3 := READTOKENST(Sp);
  Val4 := READTOKENST(Sp);
  CoutTotal := 0;
  if (Val1 = '') and (Val2 = '') and (Val3 = '') and (Val4 = '') then
  begin
    Result := 0;
    Exit;
  end;
  if (Val1 <> '') and (Val2 = '') and (Val3 = '') and (Val4 = '') then TF := TobGestIndRealise.FindFirst([ChampFormation1], [Val1], False);
  if (Val2 <> '') and (Val3 = '') and (Val4 = '') then TF := TobGestIndRealise.FindFirst([ChampFormation1, ChampFormation2], [Val1, Val2], False);
  if (Val3 <> '') and (Val4 = '') then TF := TobGestIndRealise.FindFirst([ChampFormation1, ChampFormation2, ChampFormation3], [Val1, Val2, Val3], False);
  if Val4 <> '' then TF := TobGestIndRealise.FindFirst([ChampFormation1, ChampFormation2, ChampFormation3, ChampFormation4], [Val1, Val2, Val3, Val4], False);
  while TF <> nil do
  begin
    Stage := TF.GetValue('PFO_CODESTAGE');
    Session := IntToStr(TF.GetValue('PFO_ORDRE'));
    TS := TobGestIndSession.findFirst(['PSS_CODESTAGE,PSS_ORDRE'], [Stage, Session], False);
    CoutUnit := 0;
    CoutSession := 0;
    CoutAnim := 0;
    if TS <> nil then
    begin
      CoutUnit := TS.Getvalue('PSS_COUTUNITAIRE');
      CoutSession := TS.Getvalue('PSS_COUTPEDAG');
      CoutAnim := TS.Getvalue('PSS_COUTSALAIR');
      NbSta := TS.Getvalue('NBSAL');
      coutAnim := CoutAnim / NbSta;
      CoutSession := CoutSession / NbSta;
    end;
    CoutSal := TF.GetValue('PFO_COUTREELSAL');
{$IFDEF EAGLCLIENT}
    if TF.GetValue('FRAIS') <> Null then
    begin
      if IsNumeric(TF.GetValue('FRAIS')) then CoutFrais := TF.GetValue('FRAIS')
      else CoutFrais := 0;
    end
    else CoutFrais := 0;
{$ELSE}
    if TF.GetValue('FRAIS') <> Null then CoutFrais := TF.GetValue('FRAIS')
    else CoutFrais := 0;
{$ENDIF}
    CoutTotal := CoutTotal + CoutSal + CoutFrais + CoutUnit + CoutSession + CoutAnim;
    if (Val1 <> '') and (Val2 = '') and (Val3 = '') and (Val4 = '') then TF := TobGestIndRealise.FindNext([ChampFormation1], [Val1], False);
    if (Val2 <> '') and (Val3 = '') and (Val4 = '') then TF := TobGestIndRealise.FindNext([ChampFormation1, ChampFormation2], [Val1, Val2], False);
    if (Val3 <> '') and (Val4 = '') then TF := TobGestIndRealise.FindNext([ChampFormation1, ChampFormation2, ChampFormation3], [Val1, Val2, Val3], False);
    if Val4 <> '' then TF := TobGestIndRealise.FindNext([ChampFormation1, ChampFormation2, ChampFormation3, ChampFormation4], [Val1, Val2, Val3, Val4], False);
  end;
  Result := CoutTotal;
end;

function RendNbStagiaresParAge(Sf, Sp: string): Variant;
var
  AgeMin, AgeMax, NbSal: Integer;
  Sexe: string;
  DateNaissanceMin, DateNaissanceMax: TDateTime;
  DateDebExer, DateFinExer: TDateTime;
  aa, mm, jj: word;
  Q: TQuery;
begin
  DateDebExer := StrToDate(READTOKENST(Sp));
  DateFinExer := StrToDate(READTOKENST(Sp));
  Sexe := READTOKENST(Sp);
  AgeMin := StrToInt(READTOKENST(Sp));
  AgeMax := StrToInt(READTOKENST(Sp));
  DecodeDate(DateDebExer, aa, mm, jj);
  DateNaissanceMin := EncodeDate(aa - AgeMin, mm, jj);
  if AgeMax > 0 then DateNaissanceMax := EncodeDate(aa - AgeMax, mm, jj)
  else DateNaissanceMax := IDate1900;
  Q := OpenSQL('SELECT COUNT(DISTINCT PFO_SALARIE) NBSAL FROM FORMATIONS LEFT JOIN SALARIES ON PFO_SALARIE=PSA_SALARIE ' +
    ' WHERE PFO_EFFECTUE="X" AND ' +
    'PFO_DATEDEBUT>="' + UsDateTime(DateDebExer) + '" AND PFO_DATEDEBUT<="' + UsDateTime(DateFinExer) + '" ' +
    'AND PSA_SEXE="' + Sexe + '" AND PSA_DATENAISSANCE<"' + UsDateTime(DateNaissanceMin) + '" AND PSA_DATENAISSANCE>"' + UsDateTime(DateNaissanceMax) + '"', True);
  if not Q.Eof then NbSal := Q.FindField('NBSAL').AsInteger
  else NbSal := 0;
  Ferme(Q);
  Result := NbSal;
end;

// Congés spectacles

function CongesSpectacles(Sf, Sp: string): Variant;
var
  Valeur: string;
  TypeValeur: string;
begin
  Valeur := READTOKENST(Sp);
  TypeValeur := READTOKENST(Sp);
  Valeur := PGUpperCase(Valeur);
  if TypeValeur = 'DATE' then Valeur := FormatCase(Valeur, 'STR', 5);
  Result := Valeur;
end;

function MajTableCongesSpectacles(Sf, Sp: string): Variant;
var
  Salarie, dateDeb, DateFin, Signataire, Lieu, dateDeliv, NumCertif: string;
begin
  Salarie := READTOKENST(Sp);
  dateDeb := READTOKENST(Sp);
  DateFin := READTOKENST(Sp);
  Signataire := READTOKENST(Sp);
  Lieu := READTOKENST(Sp);
  dateDeliv := READTOKENST(Sp);
  NumCertif := READTOKENST(Sp);
  if NumCertif = '' then Numcertif := '0';
  NumCertif := ColleZeroDevant(StrToInt(NumCertif), 8);
  ExecuteSQL('UPDATE CONGESSPEC SET ' +
    'PCS_NOMSIGN="' + Signataire + '",' +
    'PCS_LIEU="' + Lieu + '",' +
    'PCS_NUMCERTIFICAT="' + NumCertif + '",' +
    'PCS_DATEDELIV="' + dateDeliv + '" ' +
    'WHERE PCS_SALARIE="' + Salarie + '" ' +
    'AND PCS_DATEDEBUT="' + UsDateTime(StrToDate(dateDeb)) + '" ' +
    'AND PCS_DATEFIN="' + UsDateTime(StrToDate(DateFin)) + '"');
  result := '';
end;

//PT64
{***********A.G.L.Privé.*****************************************
Auteur  ...... : VG
Créé le ...... : 12/08/2003
Modifié le ... :
Description .. : Edition des éléments de la DADS bilatérale
Suite ........ :
Mots clefs ... : PAIE, PGDADSB
*****************************************************************}

function FunctPGDADSB(sf, sp: string): variant;
var
  BisTer, NomRue, Num, periode, TypeFonc: string;
begin
  result := '';
  TypeFonc := READTOKENST(sp);

  if (TypeFonc = 'PER') then
  begin
    periode := (READTOKENST(sp));
    if (periode <> '') then
      result := copy(periode, 1, 2) + '/' + copy(periode, 3, 2) + '    ' +
        copy(periode, 5, 2) + '/' + copy(periode, 7, 2);
  end
  else
    if (TypeFonc = 'ADR') then
  begin
    Num := READTOKENST(sp);
    BisTer := READTOKENST(sp);
    NomRue := READTOKENST(sp);
    BisTer := RechDom('PGADRESSECOMPL', BisTer, False);
    result := Num + ' ' + BisTer + ' ' + NomRue;
  end;
end;
//FIN PT64

//PT102 Calcul l'age à partir d'une date donnée en entrée (Ancienneté)
Function FunctPGCalculAnciennete(Sf, Sp: string): Variant;
var
  DateLu: TDateTime;
  SDate : string;

Begin
  SDate := READTOKENST(Sp);
  If Trim(SDate) = '' then
  Begin
    Result := '';
    Exit;
  End;
  DateLu := StrToDate(SDate);
  result := AncienneteAnnee(DateLu, Date);
End;

end.


{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/06/2001
Modifié le ... :   /  /
Description .. : Gestion des CP
Mots clefs ... : PAIE;CP
*****************************************************************}
{
PT1   : 15/06/2001 SB      Cloture de l'exercice CP
                           Cas : Incrémentation de la periode CP des mvts SLD de
                           +1;
                           Le recalcul tient compte de la date de paiement.
                           Pour ces mvts la date concerne la periode en cours au
                           lieu de la periode anterieur donc lors du recalcul
                           décalage d'une periode..
PT2   : 27/06/2001 SB      Cas d'un acquis en cours de 0
                           Message d'erreur : Division par zéro en virgule
                           flottante..
PT3   : 28/06/2001 SB      Prise de congés sur mvt de reliquat (Fiche bug n°259)
                            -Le mvt de pris n'est pas considéré comme Payé
                            -Le mvt dupliqué et consommant l'acquis n'a pas de
                           date de paiement affecté
                           En suppression de bulletin (ou option Voir Congés
                           Payés sur bulletin) :
                            -Le mvt de pris n'est pas correctement réinitialisé
                            -Le mvt dupliqué et consommant l'acquis n'est pas
                           supprimé
                           En réinitialisation de bulletin
                            -Le mvt dupliqué et consommant l'acquis est regénéré
PT4   : 23/08/2001 PH      Correction du bug du bug de la cloture des cp d'un
                           etablissement
PT5   : 28/08/2001 SB V547 Division par zéro interdit ajout de test au préalable
PT6   : 04/09/2001 SB V547 Fiche de bug n°287
                           Suppression de bulletin
                           Compteur consommé erroné
PT7   : 05/09/2001 SB      Correction Provision
                           le Solde CP ne se calque plus sur la provision donc
                           nouvel fonction GetFromTobsSoldeCP
PT8   : 05/09/2001 SB V547 L'acquisition annuelle des CP ne correspond plus à 25
                           ou 30 mais à l'acquis mensuelle*12
                           On passe en parametre ETB_NBREACQUISCP et non
                           ETB_NBJOUTRAV
PT9   : 12/09/2001 SB V547 la date de paiement du mvt dupliqué (PRI) sur la
                           periode n-2 na correspond à une date comprise dans la
                           periode 2 donc la clôture ne repositionne pas le mvt
                           sur la bonne période
                           Ajout DTCloture ds fn DupliquerMvtPris
PT10  : 12/11/2001 SB V563 Provision CP Selon la periode on prend ou non en
                           compte le reliquat
                           Modif. funct GetFromTobsCalculProvision ajout
                           variable AvecRelNeg
PT11-1: 13/11/2001 SB V563 Prise de CP sur Acquis en cours
PT11-2:                    Arrondi du SLD

PT12  : 05/11/2001 SB V562 Calcul de la duree d'absence ne doit pas dépendre de
                           la gestion des CP
PT13  : 22/11/2001 SB V563 Pour ORACLE, Le champ PCN_CODETAPE doit être
                           renseigné
PT14  : 23/11/2001 SB V563 Fiche de bug n°367 Décompte des jours spécifiques
PT15  : 28/12/2001 SB V571 Calcul de la duree si une seule plage horaire
                           renseignée
PT16  : 07/01/2002 SB V571 Arrondi des CP Acquis
PT17  : 14/01/2002 SB V571 Prise en compte dans l'utilitaire CompteurCpApayes
                           des AJP
PT18  : 14/01/2002 SB V571 L'ancienneté est alimenté si Année système = Année
                           Bulletin
PT19  : 25/02/2002 SB V571 Fiche de bug n°10039 L'affectation de la base ACQ est
                           recalculé pour ts les mvts
PT20  : 08/03/2002 SB V571 Calcul Anciennete ne fonction pas si Idem Etab et
                           Date entree salarie non soldés non cloturés
PT21  : 18/03/2002 SB V571 Fiche de bug n°10053 Acq non généré pour SLD sans Arr
PT22  : 26/03/2002 V571 SB Arrondi des mvts de SOLDES
PT23  : 27/03/2002 SB V571 Le calcul de la valorisation au dixième ne prend pas
                           en compte l'arrondi du mois pour un SLD
PT24  : 09/04/2002 SB V571 L'arrondi du SLD se calcul sur le restant des 2
                           périodes.il doit correspondre à l'arrondi de la
                           période en cours..
PT25  : 23/04/2002 SB V571 Fiche de bug 10021 : Modification
                           CalculJoursTheorique pour calcul heures théoriques
                           travaillés
PT26  : 02/05/2002 SB V571 Erreur : Pour un SLD, On consomme systématiquement
                           l'acquis en cours dans sa totalité même si on ne
                           solde pas la totalité sur ce mvt (cas AJU Ou AJP) du
                           fait de l'utilisation de pls tob pour l'integration
                           des acquis en cours
PT27-1: 22/05/2002 SB V582 Fiche de bug n°10129 point 1
                           Validation d'un bulletin pré-solde ou cloturé
PT27-2: 24/05/2002 SB V582 Fiche de bug n°10129 point 2 Gestion de pls SLD dans
                           le même mois, affectation des cons. et codeEtape
                           erronné
                           LA suppression du bulletin ne doit déconsommé les
                           mvts que jusqu'au dernier SLD effectué si existe..
PT27-3: 28/05/2002 SB V582 Fiche de bug n°10129 point 3 : On récupère les mvts
                           d'acquis posterieur au dernier SLD en edition de
                           bulletin
PT28  : 22/05/2002 SB V582 Calul de l'indemnité journalière au 10eme erroné
                           On ne doit pas intégré les mvts ACA dans le décompte
                           des Acquis Total
PT29  : 24/05/2002 SB V582 On genere un mvt de SLD même si = à 0
PT30  : 19/06/2002 SB V582 arrondi sur nb de jours restants à prendre
                           Bug : génération d'un mvt Pri duplique à 0
PT31  : 19/06/2002 SB V582 Calcul de l'indemnité au maintien sur Reliquat n'est
                           pas repris si plus avantageux que le dixième et
                           réinitialiser systématiquement à 0
PT32  : 25/07/2002 SB V585 Dû à l'intégration des mvts d'annulation d'absences
                           Contrôle des requêtes si typemvt et sensabs en
                           critere
PT33  : 01/08/2002 SB V582 Ajout d'un arrondi dans le décompte des acquis
                           apayer lors de la prise de CP
PT34-1: 05/09/2002 SB V585 Intégration de la gestion salarié de la méthode de
                           valorisation au maintien des CP
PT34-2: 05/09/2002 SB V585 FQ 10231 Encodage incorrect sur date ancienneté si
                           date d'acquisition 29/02
PT34-3: 05/09/2002 SB V585 Calcul de l'anciennete imbriqué dans test date
                           d'acquisition
PT34-4: 16/09/2002 SB V585 FQ n°10233 récupération de la date acquisition
                           ancienneté si personnalisé salarié à activer
PT34-5: 17/09/2002 SB V585 FQ n°10181 Consommation d'acquis : on se positionne
                           par défaut sur les mvts REL..
PT35  : 23/09/2002 SB V585 FQ n°10242 Provision CP Edition du reliquat sur
                           Periode 1
PT36  : 03/10/2002 SB V585 Optimisation requête saisie de la paie et editions
PT37  : 11/10/2002 SB V585 Pour utilisation monobase des Appli. Paie & Econges
                           On charge les CP Pris validé
PT48  : 17/10/2002 SB V585 FQ n°10235 Intégration du calcul des variables 28,29
                           heures ouvrées ouvrables
PT49  : 23/10/2002 SB V585 Cloture CP : On ne recupère pas dans la génération du
                           reliquat de la nouvelle période
                           la base et le nbre de mois des reliquats + de la
                           période n-1 si consommés
PT50  : 20/11/2002 PH V591 Portage eAGL
PT51  : 20/12/2002 SB V591 FQ 10391 Recupération du calendrier affecté au
                           salarié
PT52  : 30/12/2002 SB V591 FQ 10408 Pour le calcul du dixième (PRI ou SLD) on ne
                           doit pas tenir compte des AJP
PT52   remodifié le 09/01/2003
PT53  : 07/01/2003 SB V591 fn CalculPeriode test si date cloture <= 01/01/1900
PT54  : 15/01/2003 SB V585 Provision CP Suppression de la Tob mère des filles
                           dupliquées
PT55  : 30/01/2003 SB V591 Vidage Tob_Abs en sortie et ré entrée de bulletin
PT56  : 07/02/2003 SB V591 FQ 10490 Violation d'accès
PT57  : 17/02/2003 SB V_42 Edition Cp bulletin : contrôle gestion CP
PT58  : 11/03/2003 SB V_42 Optimisation des traitements
PT59  : 07/05/2003 SB V591 FQ 10682 Suite PT54 : Mouvement REL + non intégré sur
                           N1
PT60  : 19/05/2003 SB V_42 FQ 10693 Refonte décompte des jours CP,absence, temps
                           de travail
PT61-1: 02/06/2003 SB V_42 FQ 10705 Mvt SLD de Cloture non valorisé..
PT61-2: 04/06/2003 SB V_42 FQ 10523 Intégration de l'acquis bulletin calculé sur
                           une variable de paie
PT61-3: 11/06/2003 SB V_42 FQ 10718 Recherche Ancienneté niveau etab ou salarié
                           érroné
PT61-4: 24/06/2003 SB V_42 FQ 10454 Calcul au maintien sur premier bulletin
                           implique un recalcul des CP
PT61-5: 16/07/2003 SB V_42 FQ 10705 Affectation erroné mvt de cloture si
                           paiement du solde
PT62-1: 25/07/2003 SB V_42 Dysfonctionnement decompte SLD sur N-1 avec
                           ajustement
PT62-2: 25/07/2003 SB V_42 Ajout d'un arrondi sur le calcul de décompte des
                           ajustements - sur les acquis
PT62-3: 07/08/2003 SB V_42 FQ 10738 Refonte calcul periode cloture CP
PT63  : 10/10/2003 SB V_42 Initialisation jours cp pris à intégrer
PT64  : 12/11/2003 SB V_42 Non prise en compte des SLD de cloture en raz
                           compteur payes CP
PT65  : 13/11/2003 SB V_42 Prise en compte des SLD de cloture en suppr et
                           edition de bulletin
PT66  : 18/11/2003 SB V_50 FQ 10794 Controle profil Congés payés existant, suppr
                           des SLD de sortie non payé
PT67  : 09/12/2003 SB V_50 FQ 10370 Edition compteurs Cp selon paramétrage  -> Voir PT80
PT68-1: 16/01/2004 SB V_50 FQ 11063 Décompte demi journée ouvrable
PT68-2: 16/01/2004 SB V_50 Edition bulletin période Cp antérieur..prise en
                           compte des mvts de cloture ou non
PT69  : 26/01/2004 SB V_50 FQ 11120 Decompte érronné si prise de congés payés
                           sur date de sortie salarié
PT70-1: 16/02/2004 SB V_50 FQ 11111 Provision CP : Ne pas calculé de provision
                           sur N-1 si restant=0
PT70-2: 17/02/2004 SB V_50 FQ 11111 Provision CP : Sur N-1 calcul provision sur
                           restant reliquat et sur restant acquis
PT70-3: 17/02/2004 SB V_50 FQ 11111 Provision CP : Intégration des arrondi de
                           SLD
PT71  : 12/03/2004 SB V_50 FQ 11162 Encodage de la date de cloture erroné si fin
                           fevrier
PT72  : 18/03/2004 SB V_50 FQ 11111 Refonte calcul provision pour séparation des
                           reliquats + des acquis
PT73  : 18/03/2004 SB V_50 FQ 11020 Modification utilisateur des acquis sur la
                           saisie de la paie
PT74-1: 09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                           salarié
PT74-2: 09/04/2004 SB V_50 FQ 11237 Ajout Gestion idem etab jours Cp Suppl
PT74-3: 16/04/2004 SB V_50 Distinction des AJU sur reliquat ou sur Acquis
                           (ACQ,ACA,ACS)
                           Calcul de l'indem. Cp Xeme sur reliquat avec AJU
                           Suppression Bulletin : Deconsommation des pris sur
                           reliquat
                           Provision Cp : Calcul de la provision
PT74-4: 09/04/2004 SB V_50 Ajout contrôle chargement ancienneté CP
PT75-1: 29/04/2004 SB V_50 FQ 11248 Calcul duree : Blocage si saisie du matin au
                           matin sur jour non travaillé
PT75-2: 29/04/2004 SB V_50 FQ 11249 Prise CP, après avoir pris sur reliquat le
                           FindNext ne poursuit pas sur l'acquis suivant
PT75-3: 30/04/2004 SB V_50 FQ 11111 Provision CP : Intégration des mvts SLD de
                           cloture CP
PT76  : 07/05/2004 SB V_50 FQ 11199 Dysfonctionnement Calcul Indemnité CP sur
                           1er bulletin
PT77  : 18/05/2004 SB V_50 FQ 11266 CalculDuree : Prise en compte des jours et
                           heures spécifiques du calendrier pour les absences
PT78  : 25/06/2004 VG V_50 La TOB mère étant réelle et à vide, un enregistrement
                           à blanc était créé lors du TOB.InsertOrUpdateDb
PT79  : 10/09/2004 MF V_50 Traitement des IJSS et du maintien
PT80  : 28/09/2004 JL V_50 Modifs requête pour FQ 10370 : si changement
                           d'établissement la requête ne renvoyait aucun
                           résultat
PT81  : 04/10/2004 MF V_50 Traitement des IJSS et du maintien (Tob_Abs en
                           Append)
PT82  : 15/11/2004 PH V_60 FQ 11703 Salarié sorti à la date d'arrêté
PT83  : 01/12/2004 PH V_60 Calcul au dixième faux Modif apportée sur la V5
PT84  : 06/12/2004 PH V_60 Modification Annulation Bulletin : Prise en compte
                           du 1er bulletin
PT85  : 04/01/2005 PH V_60 Rajout parenthèse pour SQL compatible avec V500
PT86  : 10/01/2005 PH V_60 FQ 11910 Erreur ORACLE récup de champ à NULL
PT87  : 25/01/2005 PH V_60 FQ 11922 GEstion des dates exerices CP pour édition
                           CP et RTT
PT88  : 31/01/2005 SB V_60 Maj champ proc. InitialiseTobAbsenceSalarie
PT89  : 31/01/2005 SB V_60 Export des proc. communes dans pgcommun
PT90  : 01/02/2005 SB V_60 Cloture Cp : On ne tient pas compte du Rel +
                           anterieur pour calcul Arrondi.
PT91  : 03/02/2005 SB V_60 FQ 11692 Initialisation des dates à Idate1900
PT92  : 07/02/2005 SB V_60 ProvisionCP : Optimisation traitement
PT93  : 14/02/2005 SB V_60 FQ 11967 : CalculDuree, pas de calcul calendaire si
                           pas de durée calendrier
PT94  : 02/03/2005 MF V_60 FQ 12058 : si maintien non géré (décoche en prepa
                           auto) il faut intégrer tous les types d'absence
PT95  : 15/03/2005 SB V_60 FQ 12021, FQ 12088 Solde & Provision CP : Conversion
                           TypeVariant incorrect sous oracle
PT96  : 22/03/2005 SB V_60 FQ 11931 Calcul duree, si sans decompte des jours non
                           travaillé, pas de recherche calendrier etablissement
PT97  : 29/03/2005 SB V_60 FQ 11990 Optimisation traitement Calendrier
PT98  : 31/03/2005 SB V_60 FQ 12127 Affichage Compteurs Cp erronés sur bulletin
                           pour mvt rel N+1
PT99  : 02/06/2005 SB V_60 FQ 12327 Econges : En monobase mise à jour du
                           exportok via la paie
PT100 : 01/07/2005 SB V_60 FQ 12383 Prise en compte des sorties inférieur à
                           IDate1900
PT101 : 05/07/2005 SB V_65 FQ 12384 Bulletin ICCP Xeme : ne pas tenir compte des
                           ACS
PT102 : 07/07/2005 SB V_65 FQ 12065 Génération Acquis = 0 si cumul 11 > 0
PT103 : 06/09/2005 SB V_65 FQ 12384 Bulletin ICCP : Rédéfinition des acquis pris
                           en compte pour le diviseur
PT104 : 16/09/2005 SB V_65 DQ 12420 Bulletin SLD : défaut d'intégration des SLD
                           de clôture
PT105 : 17/11/2005 SB V_65 FQ 12700 Conges : Ajout clause date validité <=
                           datefin bulltin
PT106 : 08/12/2005 SB V_65 FQ 12700 Traitement de la tob des ajustement - en tob
                           globale
PT107 : 27/03/2006 PH V_65 Prise en compte des champs motifs absences dédoublés
                           (Heures,Jours)
PT108 : 02/05/2006 SB V_65 FQ 11426 Mis en place paramètre pour calcul au
                           maintien CP
PT109 : 24/04/2006 SB V_65 FQ 12786 Anomalie calcul variable 0027 en prep. auto
PT110 : 09/05/2006 MF V_65 FQ 13030 correction : en modification chrgt TOB
                           Absences
PT111 : 19/05/2006 SB V_65 FQ 13154 Calcul du Xième différent si Paramsoc CHR
                           6ème semaine CP suppl.
PT112 : 19/05/2006 SB V_65 FQ 13155 Distinction des mvt REP d'acs, et d'aca dans
                           calcul Xième
PT113 : 02/06/2006 SB V_65 Optimisation mémoire
PT114 : 16/06/2006 SB V_65 FQ 13298 Provision cp : anonalie sur code
                           alphanumérique en minuscule
PT115-1:07/07/2006 SB V_65 FQ 11426 Anomalie chargement du cumul12 sur paie en
                           cours
PT115-2:07/07/2006 SB V_65 Anomalie déallocation mémoire de la Tob_pris =>
                           distinction du nom entre la tob globale et la tob
                           procédural
PT115-3:31/08/2006 SB V_70 Déallocation mémoire
PT116 : 12/07/2006 SB V_65 FQ 13320 Intégration des absences en création,
                           exclusion des absences annulées
PT117 : 20/07/2006 SB V_65 Déallocation mémoire de la tob des editions Solde &
                           Provisions CP
PT118 : 28/07/2006 MF V_65 correction : Qd reglt IJ sans absence de type IJSS le
                           Tob_Abs était vidée à tort, les autres absences (non
                           IJSS) étaient intégrées dans le bulletin mais pas maj
                           dans ABSENCESALARIE.
PT119 : 12/09/2006 SB V_65 FQ 13486 Correction calcul arrondi SLDC
PT120 : 23/11/2006 SB V_70 FQ 13703 Correction violation d'acces
PT121 : 25/01/2007 SB V_70 FQ 11992 Ajout valeur methode date d'acquisition CP
PT122 : 25/01/2007 SB V_70 Contrôle cohérence saisie date si modif. paramètres
                           régionaux PC en Anglais
PT123 : 08/03/2007 PH V_70 Repris la valeur du cumul 12 en cours de calcul
PT124 : 29/05/2007 PH V_70 FQ 14031 Division par 0
PT125 : 27/06/2007 VG V_72 MemCheck
PT126 : 17/08/2007 PH V_72 FQ 14042 et 14393 Vérifcp et mode silence= true
PT127 : 28/08/2007 FC V_72 Extraction de fonctions dans ULibEditionPaie
PT128 : 05/10/2007 MF V_80 Traitement des congés de fractionnement
PT129 : 07/12/2007 FC V_80 FQ 15019 Violation d'accès en sortant du zoom CP en multi
PT131 : 05/06/2008 MF V_82  FQ 459 et FQ 15466
PT134 : 08/09/2008 PH  V810 Optimisation gestion du fractionnement
}
unit PgCongesPayes;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  SysUtils, HCtrls, HEnt1, ed_tools, Hdebug,
{$IFNDEF EAGLCLIENT}
  Db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
(*  Mise en commentaire des uses inutiles
  Fiche,fichlist,
{$ELSE}
  eFiche, eFichList,
{$ENDIF}
*)
  Controls, HMsgBox, UTOB
  ,ULibEditionPaie;//PT127

const
  DCP: integer = 2; // Nbre de décimales pour le calcul des congés payés

  { Traitement bulletin CP & Absences }
//PT89 procedure InitialiseTobAbsenceSalarie(TCP: tob);
function CalculValoCP(Tb_Pris, TobAcquis, TobEtab, TobSAL: tob; NbJOuvres, NbJOuvrables: double; datedeb, datefin: tdatetime; var Next: integer; Typecp, salarie: string): boolean;
function TrouveModeValorisation(T_etab, T_Sal: tob): string; //PT36 Passage en paramètre de la tob salarie
procedure RendMethodeValoDixieme(var NBjSem: integer; var MValoMs: string; var ValoDxmn: double);
function JoursReference(NbJoursAcquis: double): Integer;
function AssezAcquis(T_Acquis: tob; Nbjourspri: double; var prispayes, acquisres: double; calculAcquis: boolean): boolean;
procedure AgregeCumulPris(T_Pris: Tob);
procedure DupliquerMvtPris(var TP, TA, Tb_Pris: tob; MS, DX, JTraites, JATraiter: double; var Next: integer; DTcloture, DD, DF: TdateTime; ModeValo: string);
procedure ChargeTobDelta(Salarie: string; DateF: TDateTime; var Lctob_delta: tob); { PT105 } { PT106 }
function CalculDelta(Lctob_delta: tob; DTclotureP, dateFB: tdatetime; Periode: integer): double; { PT106 }
procedure DeSoldeTobDelta(Salarie: string; DtSld: TDateTime; StWhere: string);
function CalculValoX(DTclotureP, DateFB, DateDB: tdatetime; periode: integer; NbJoursRef: double; Typecp: string; T_Acquis, lctob_delta: tob): double; { PT106 }
function CalculValoReliquat(T_Acquis, Lctob_delta: TOB; NbJoursRef: double): Double; { PT106 }
function CalculValoAbsence(NbJ: double; DebMoisprec, FinMoisprec: Tdatetime; TypeCp: string): double;
function ReCalculeValoAbsence(Tb_pris: tob; Valo, PaieSalMaint: string; Mul: double; typecp: string): boolean; { PT108 }
procedure EcritMvtPris(T: tob; CumMs, CumX, Absence: double; ModeValo: string);
// d PT79
// function  RecupereAbsences (Var Tob_ABS:tob;Datef:tdatetime;Salarie,Action:string):boolean;
// f PT79
{PT131
function RecupereAbsences(var Tob_ABS: tob; Datef: tdatetime; Salarie, Action: string; GenereMaintien, IJSS: boolean): boolean;}
function RecupereAbsences(var Tob_ABS: tob; Datef: tdatetime; Salarie, Action: string; GenereMaintien, IJSS: boolean; TypeAbs : string): boolean;
function RecupereCongespris(Tob_Pris, TobSal: Tob; DtFinP: TDateTime): boolean;
function ModificationCongesPris(Tob_pris, TobSal: Tob; DtFinP, Dateentree: tdatetime): boolean;
procedure CalculVarOuvresOuvrablesMois(TobEtab, TobSal, Tob_Cal: Tob; datedebut, datefin: tdatetime; Ouvres: boolean; var Nb_j, Nb_H: double); { PT48 Modification de la function en procedure }
procedure RecupereCongesPrisPayes(LcTob_pris, LcTob_solde: tob; Datef: tdatetime; Salarie, typeCp: string); { PT115-2 }
//function AlimCongesAcquis(Tobsalarie, Tobetablissement: Tob; datedebut, datefin: tdatetime; numerateur, denominateur: integer; TOB_RUB: Tob; Ecrit: boolean; var base, Nbmois: double; var No_ordre: integer; var CpMois, CpSupp, CpAnc: double): double;
// PT128 function ChargeAcquisParametre(Tobsalarie, Tobetablissement, TOB_RUB: Tob; datedebut, datefin: tdatetime; OnLoad, OnRefresh: Boolean; var base, Nbmois, CpMois, CpSupp, CpAnc: double): double; { PT73  }
function ChargeAcquisParametre(Tobsalarie, Tobetablissement, TOB_RUB: Tob; datedebut, datefin: tdatetime; OnLoad, OnRefresh: Boolean; var base, Nbmois, CpMois, CpSupp, CpAnc, CpFract: double): double; { PT73  }
procedure InitTobAcquisCp(Tobsalarie, Tobetablissement: Tob; datedebut, datefin: tdatetime; base, Nbmois, CpMois, CpSupp, CpAnc: double; var No_ordre: integer); { PT73  }
// PT128 function ChargeT_MvtAcquis(Salarie: string; DD, DF: TDateTime; Mois, Supp, Anc: Double; var Modif, New: Boolean): boolean;
function ChargeT_MvtAcquis(Salarie: string; DD, DF: TDateTime; Mois, Supp, Anc, Fract: Double; var Modif, New: Boolean): boolean;
procedure AffectBaseAcquisEncours;
procedure AjouteAcquiscourant(Tob_Acquis, Tob_sal: tob; DateF: tdatetime; var No_Ordre: integer);
procedure RetireAcquiscourant(Tob_Acquis: tob);
procedure AnnuleCongesAcquis(CodeSalarie: string; datedebut, datefin: tdatetime);
procedure AnnuleCongesAcquisDeTob(CodeSalarie: string; datedebut, datefin: tdatetime; T_A: tob);
procedure ChangeCongesAcquisDeTob(CodeSalarie: string; datedebut, datefin: tdatetime; T_A: tob);
procedure AnnuleCongesAcquisMasse(datedebut, datefin: tdatetime);
procedure SupprimeCPMasse(DD, DF: tdatetime);
procedure SupprimeCP(T_pris, T_acquis: tob; SalEnCours, EtabEnCours: string; DTClot, DD, FF: tdatetime; StWhere: string); { PT106 }
procedure DeconsommeMvtAcq(T_Acquis: Tob; Sal: string; Periode: integer; CptPer, NbConsRelMvt: Double; FF: TDateTime);
procedure AnnuleCongesPris(CodeSalarie, Etab: string; datedebut, datefin: tdatetime);
procedure CreeArrondi(Arr: double; Salarie, etab, clot: string; var suivant: integer; periode: integer; dateE: tdatetime; T: tob);
//PT127 function IsAcquis(TypeConge, Sens: string): boolean;
//PT127 function IsPris(TypeConge: string): boolean;
{ Cloture / Decloture des Congés Payés}
function CloturePeriodeCP(datefinperiode: TDatetime; Etablissement, MethodeRel: string): boolean;
function DeCloturePeriodeCP(FinPeriodeN: TDatetime; Etab, EtabLibelle: string): Boolean;
procedure ClotureMvtAcquis(T_ListeMvt: Tob; salarie: string; reliquat: integer; DatedebutperiodeP, DatefinperiodeP: TDateTime);
procedure GenereMvtCloture(T_MVTINS, TS: Tob; NoOrdre, periode: integer; TypeConges, sens, libelle, codeetape: string; datevalidite, AutreDate: Tdatetime; nbjour, nbmois, base: double);
{ Edition du Bulletin }
//PT127 procedure RechercheCumulCP(DD, DF: tdatetime; Salarie, Etab: string; var PN, AN, RN, BN, PN1, AN1, RN1, BN1: double; var DDP, DFP, DDPN1, DFPN1: tdatetime; var EditbulCP: string);
{ Utilitaire CP}
//PT89 function CalculPeriode(DtClot, DTValidite: TDatetime): integer;
//PT127 function RendPeriode(var DTcloturePi, DTDebutPi: tdatetime; DTCloture, datevalidite: tdatetime): integer;
//PT127 procedure RechercheDate(var DDebut, Dfin, Datecloture: tdatetime; Periode: string);
procedure CalculDuree(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
procedure CalculJoursTheorique(Tob_sal: tob; datedebut, datefin: tdatetime; var Varj, Varh: double); //PT-25
//PT127 procedure AffichelibelleAcqPri(Periode, Salarie: string; DateDeb, Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double; edit, SansCloture: boolean);
//PT127 procedure CompteCp(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase, valo: double); { PT70-2 Ajout ConsRel }
procedure CompteCpAPAYES(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double);
function IsValeurNumeric(const S: string): boolean;
//PT89 function PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime; { PT71 }
{ Solde de tout compte }
procedure GenereSolde(Tob_Sal, Tob_etab, Tob_Pris, Tob_Acquis, Tob_Solde: tob; DateF, DateD: tdatetime; var next: integer; AcquisBulletin, PrisBul: double; Action: string);
procedure CalculAcquisPeriodeSolde(T_etab, T_Acquis: tob; DTDEBP, DTFINP, DateD, DateF: tdatetime; var NbJAcqP, NbJacq, NBJPayes: double; salarie: string); { PT105 }
procedure SupprimeMvtSolde(Tob_pris: tob);
procedure SoldeAcquis(Tob_Acq: tob; Apayes: Boolean; LcDateFin: TDateTime); //PT-11-1 Ajout variable Apayes { PT106 }
procedure IntegreSolde(Tob_solde, Tob_pris: tob);
procedure RecalculePeriode(Etab: string; DateClotCp: TDateTime);
{ Calcul Provision CP & Solde CP}
function CalculeProvisionCp({Etab,}StWhere: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime): Tob;
function CalculeSoldeCp({Etab,}StWhere: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime): Tob;
procedure ChargeTobsCalculProvision(TAbs, TSld: tob; StWhere: string; DateFin: tdatetime); { PT72 }
function GetFromTobsCalculProvision(Tabs: Tob; Salarie, Mvt, Periode: string; AvecArr, AvecRelPos, AvecRelNeg: Boolean): Tob;
function GetFromTobsSoldeCP(Tabs: Tob; Salarie, Mvt, Periode: string; Edit, SansCloture: Boolean): Tob;
procedure FormateTobprovision(TW: tob);
procedure CalculProvision(Salarie: string; TAbs, TSld, TW: tob; DateEtat, DTcloture: TDatetime; NbJoursReference: integer); { PT72 }
function GetMvtFromTobProvision(Tabs, TSld: Tob; Salarie, Typ, Periode: string; DateEtat, DateSolde: TDateTime): Tob; { PT72 }

{ Calendrier }
function ChargeCalendrierSalarie(Etablissement, Salarie, Calend, StandCalend, EtabStandCalend: string): Tob;
// @@@@
procedure FgInitDotProvCp(Etab: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime);
function TrouvePaieSalMaintien(T_etab, T_Sal: tob): string; { PT108 }
function RecupSalMaintien(PaieSalMaint: string; DateAbs, DateBull: TDateTime): Double; { PT108 }

implementation

uses p5util,
  pgCommun,
  PgOutils,
  //      PgVisuObjet,
  PgCalendrier,
  uPaieEtabCompl,
  UTOFPGUtilitaireCP,
  EntPaie;

var
  T_Etab: tob;
  JOuvres, JOuvrables: double;
  DTFin, DTDeb: TDatetime;

  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : Paie Pgi
  Créé le ...... : 28/01/2003
  Modifié le ... :   /  /
  Description .. :  Valorisation de l'absence au dixième et au maintient, retour
  Suite ........ : avec la TOB des congés enrichie des calculs et de la TOB
  Suite ........ : des Acquis à écrire en cas de validation
  Mots clefs ... : PAIE;CP
  *****************************************************************}

function CalculValoCP(Tb_Pris, TobAcquis, TobEtab, TobSAL: tob; NbJOuvres, NbJOuvrables: double; Datedeb, datefin: tdatetime; var Next: integer; TypeCp, salarie: string): boolean;
var
  Mvt, Sens, st, ModeValo, MValoMs: string;
  TP, TA, TD, Test, tob_deltaP, {tob_delta, PT106 } T_Acquis: TOB;
  Q: TQuery;
  ValoX, ValoMS, CumSMAP, CumSBAP, CumSBAP1, CumSMAP1, NMois, BASE, acquisres: double;
  CumMS, CumX, NBj, Nbjourspri, Congeacquis, MS, DX, Absence, APayes, JTraites: double;
  JATraiter, prispayes, cumul12, ValoDxmn, HVal, delta, deltaInit, TJours, TJAPayes: double;
  ConsoAjustement, IndemJourRel: double; //PT62-1
  ValoJX, DeltaJ: array[0..10] of double;
  DTFinP, DTDebutP, DTFinP1, DTValidite, DTcloture, FinMoisprec, DebMoisprec: TDateTime;
  Datedebut: TDateTime;
  NbJSem, NbJoursReference, NoP, IPAcquis, CodeRgpt, code, per, ordre: integer;
  aa, mm, jj: word;
  TYPEIMPUTE, PaieSalMaint: string; // PT86 , PT108
//  WTypeConge  : string; 
begin
  { Typecp = 'PRI' ou 'SLD' }
  result := true;
  if Tb_Pris = nil then exit;
  T_Etab := TobEtab;
  T_Acquis := TobAcquis;
  DTcloture := T_ETAB.getvalue('ETB_DATECLOTURECPN');
  { ne doit jamais arriver }
  if DTcloture = 0 then exit;
  prispayes := 0;
  IndemJourRel := 0; { PT74-3 }
  AcquisRes := 0;
  if TypeCp = 'SLD' then AcquisMaj := true;
  // Tob_delta := nil; { PT106 }
  { ChargeTobDelta : alimentation de la tob des ajustements pour chercher le delta }
//  ChargeTobDelta(Salarie,datefin, tob_delta); { PT105 } { PT106 }
  delta := 0;
  for NoP := 0 to 10 do
  begin
    DeltaJ[NoP] := 0;
    DeltaJ[NoP] := CalculDelta(tob_delta, DTcloture, datefin, NoP);
    delta := delta + deltaJ[NoP];
  end;
  deltaInit := delta;

  ModeValo := TrouveModeValorisation(T_etab, TobSAL); //PT36

  PaieSalMaint := TrouvePaieSalMaintien(T_etab, TobSAL); { PT108 }

  if Assigned(T_Acquis) and (T_Acquis.detail.count > 0) then // notvide(T_Acquis, true) then PT113
    AssezAcquis(T_Acquis, 0, prispayes, acquisres, True);
  { basé sur calendrier réel ou théorique }
  { DEB PT34-1 Intégration de la méthode de valorisation au maintien cas personnalisé salarié }
  RendMethodeValoDixieme(NBjSem, MValoMs, ValoDxmn);
  if MValoMs = 'T' then { Calendrier théorique }
  begin
    JOuvres := 21.67;
    JOuvrables := 26;
  end
  else
  begin
    if MValoMs = 'M' then { Gestion manuelle }
    begin
      JOuvres := ValoDxmn;
      JOuvrables := ValoDxmn;
    end
    else
    begin { Calendrier réel }
      if typecp = 'PRI' then
      begin
        JOuvres := NBJOuvres;
        JOuvrables := NBJOuvrables;
      end
      else
      begin
        {PT48 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables }
        { Dans ce cas seuls les jours nous intéresse.. }
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebutdeMois(Datedeb), FindeMois(Datedeb), TRUE, JOuvres, HVal);
        CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebutdeMois(Datedeb), FindeMois(Datedeb), FALSE, JOuvrables, HVal);
      end;
    end;
  end;
  { FIN PT34-1 }
  DTFin := Datefin;
  Datedebut := DebutDeMois(DateFin);
  DTDeb := Datedebut;
  FinMoisprec := datedebut - 1;
  Debmoisprec := DebutdeMois(FinMoisPrec);
  (* PT108 Refonte traitement pour recherche salaire perçu en fonction de la date de début d'absence
  cumul12 := Valcumuldate('12', DebMoisprec, FinMoisprec);
  if cumul12 = 0 then Cumul12 := Valcumuldate('12', Datedeb, Datefin);
  if Cumul12 = 0 then Cumul12 := RendCumulSalSess('12');
  if FirstBull then Cumul12 := RendCumulSalSess('12'); { PT76 }  *)

  { balayage de la TOBACQUIS sur période P (en cours) }
  { et P-1 (périodes antérieures non cloturées) }
  {pour réaliser le calcul de l'indemnité journalière des congés au Xème. cf CalculValoX }
  { Calcul cumul des nbremois d'acquisitions,cumul des bases acquises }
  decodedate(DTcloture, aa, mm, jj);
  DTFinP1 := PGEncodeDateBissextile(aa - 1, mm, jj); { PT71 }
  DTDebutP := DTFinP1 + 1;
  NbjSem := T_Etab.getvalue('ETB_NBJOUTRAV');
  if TobSal.GetValue('PSA_CPACQUISMOIS') = 'ETB' then //DEB PT-8
    NbJoursReference := JoursReference(T_Etab.getvalue('ETB_NBREACQUISCP'))
  else
    NbJoursReference := JoursReference(TobSal.getvalue('PSA_NBREACQUISCP')); //FIN PT-8
  for NoP := 0 to 10 do
    ValoJX[NoP] := CalculValoX(DTcloture, datefin, datedeb, NoP, NbJoursReference, Typecp, T_Acquis, tob_delta);
  { Balayage de chacun des éléments de la TOB Congé (Tob_pris). }
  { Calcul de la valorisation au maintien et au Xème }
  CumMS := 0;
  CumX := 0;
  if Assigned(Tb_Pris) and (Tb_Pris.detail.count > 0) then { PT113 if notvide(Tb_Pris, true) then} TP := Tb_Pris.findfirst([''], [''], True);
  if Assigned(T_ACQUIS) and (T_ACQUIS.detail.count > 0) then //if notvide(T_ACQUIS, true) then PT113
  begin { DEB PT34-5 On se positionne toujours sur le mvt de reliquat pour consommer }
    TA := T_ACQUIS.findfirst(['PCN_TYPECONGE'], ['REL'], True);
    if TA <> nil then
    begin
      { PT74-3 Calcul de la valorisation du reliquat, intégration des ajustements sur reliquats }
      IndemJourRel := CalculValoReliquat(T_ACQUIS, tob_delta, NbJoursReference);
      { si reliquat consommé alors on se positionne sur premier acquis de la tob }
      if Arrondi(TA.GetValue('PCN_JOURS'), 2) - Arrondi(TA.GetValue('PCN_APAYES'), 2) <= 0 then
        TA := T_ACQUIS.findfirst([''], [''], True);
    end
    else { si reliquat inexistant on se positionne sur premier acquis de la tob }
      TA := T_ACQUIS.findfirst([''], [''], True);
  end; { FIN PT34-5 }
  NOP := -1; Congeacquis := 0;
  while TP <> nil do
  begin
    if (TP.getValue('PCN_CODERGRPT') <> -1) then
    begin
      Mvt := TP.getvalue('PCN_TYPECONGE');
      Nbjourspri := TP.getvalue('PCN_JOURS');
      { est-ce que je pourrai payer totalement ma ligne de congés pris? }
      if not AssezAcquis(T_Acquis, Nbjourspri + delta, prispayes, acquisres, false) then
      begin
        TP.free;
        result := false;
        TP := Tb_Pris.findnext([''], [''], True);
        continue;
      end;

      { PT108 Recherche du salaire perçu précedant la date de début d'absence et selon le paramétrage défini }
      Cumul12 := RecupSalMaintien(PaieSalMaint, TP.GetValue('PCN_DATEDEBUTABS'), DateDeb);

      delta := 0;
      ValoX := 0;
      ValoMS := 0;
      ConsoAjustement := 0; //PT62-1
      while ((TA <> nil) and (Nbjourspri > 0)) do
      begin
//        WTypeConge  := TA.getvalue('PCN_TYPECONGE');    //??
        Per := TA.getvalue('PCN_PERIODECP');
        { DEB PT33 Erreur d'arrondi lors de la soustraction des 2 champs de la Tob }
        { Remplacement des TA.GetValue par les variables }
        TJours := Arrondi(TA.GetValue('PCN_JOURS'), DCP);
        TJAPayes := Arrondi(TA.getvalue('PCN_APAYES'), DCP);
        { FIN PT33 }
        if (TJours - Arrondi(TJAPayes + ConsoAjustement, DCP) <= 0) then
        begin
          TA := T_ACQUIS.findnext([''], [''], True);
          continue;
        end;
        ConsoAjustement := 0; //PT62-1
        { je commence par défalquer le deltaj de la période avant de payer les congés restants }
        if (Arrondi(TJours - TJAPayes - DeltaJ[per], DCP) <= 0) then //PT62-2 Ajout de l'arrondi
        begin
          DeltaJ[per] := DeltaJ[per] - ((TJours - TJAPayes));
          TA := T_ACQUIS.findnext([''], [''], True);
          continue;
        end
        else
        begin
          Congeacquis := TJours - DeltaJ[per];
          { PT62-1 A ce niveau l'ajustement - restant est pris en partie sur l'ACQ }
          //*** SB ConsoAjustement := deltaj[per] ;
          DeltaJ[per] := 0;
        end;
        { si j'arrive là, le delta période est soldé }
        APayes := TA.getvalue('PCN_APAYES');
        Mvt := TA.getvalue('PCN_TYPECONGE');
        if (Congeacquis - APayes) >= NbJourspri then
        begin
          TA.putvalue('PCN_APAYES', Arrondi(Apayes + Nbjourspri, DCP));
          NBj := Nbjourspri;
        end
        else
        begin
          TA.putvalue('PCN_APAYES', Arrondi(Congeacquis, DCP));
          Nbj := Congeacquis - APayes;
        end;

        if NoP = -1 then
          NoP := RendPeriode(DTFinP, DTDebutP, DTcloture, TA.getvalue('PCN_DATEVALIDITE'));
// PT128        if NoP <> RendPeriode(DTFinP, DTDebutP, DTcloture, TA.getvalue('PCN_DATEVALIDITE')) then
        if ((TA.getvalue('PCN_TYPECONGE')='ACF') and (NoP <> TA.getvalue('PCN_PERIODECP'))) or
           ((TA.getvalue('PCN_TYPECONGE')<>'ACF') and (NoP <> RendPeriode(DTFinP, DTDebutP, DTcloture, TA.getvalue('PCN_DATEVALIDITE')))) then
        begin
          { on change de période de paiement donc éclatement du mvt de pris }
          { le mvt dupliqué est considéré traité, dc on l'insère avant le tp courant }
          JTraites := TP.getvalue('PCN_JOURS') - Nbjourspri;
          JATraiter := Nbjourspri;
          DupliquerMvtPris(TP, TA, Tb_pris, CumMS, CumX, JTraites, JATraiter, Next, DtCloture, DebMoisPrec, FinMoisPrec, ModeValo);
          { Je me positionne sur la deuxième partie du mvt éclaté à traiter }
          TP := Tb_Pris.findnext([''], [''], True);
          if TP <> nil then TP := Tb_Pris.findnext([''], [''], True); { pour se retrouver sur TU }
          if TP = nil then break; { ne doit jamais arriver }
          CumMs := 0;
          CumX := 0;
          NoP := RendPeriode(DTFinP, DTDebutP, DTCloture, TA.getvalue('PCN_DATEVALIDITE'));
        end;

        { si le mvt est de type reliquat, il s'adresse à une autre période, }
        { dc on génère un mvt pour cette fraction des congés }
        { le mvt dupliqué est considéré traité, dc on l'insère avant le tp courant }
        // DEB PT86
        if TA.getvalue('PCN_TYPEIMPUTE') <> NULL then TYPEIMPUTE := TA.getvalue('PCN_TYPEIMPUTE')
        else TYPEIMPUTE := '';
        if (Mvt = 'REL') or ((Mvt = 'AJU') and (TYPEIMPUTE = 'REL')) then { PT74-3 } // FIN PT86
        begin
          TP.PutValue('PCN_PERIODEPY', TA.getvalue('PCN_PERIODECP'));
          TP.PutValue('PCN_PERIODECP', TA.getvalue('PCN_PERIODECP'));
          if (Mvt = 'AJU') then St := 'Ajustement sur ' else St := ''; { PT74-3 }
          TP.putvalue('PCN_PERIODEPAIE', St + 'Reliquat au ' + datetostr(TA.getvalue('PCN_DATEVALIDITE')));

          MS := 0; //PT-31 Initialisation à 0
          if (NbJSem = 5) and (Jouvres <> 0) then MS := cumul12 * Nbj / Jouvres; { PT-5 }
          if (NbJsem = 6) and (Jouvrables <> 0) then MS := cumul12 * Nbj / Jouvrables; { PT-5 }

          (*if (TA.getvalue('PCN_NBREMOIS')*10*NbJoursReference) <>0 then  { PT-5 }
            DX := (12*TA.getvalue('PCN_BASE') * Nbj ) / (TA.getvalue('PCN_NBREMOIS')*10*NbJoursReference)
          else
            DX:=0;*)
          Dx := IndemJourRel * Nbj;
          JTraites := NBJ;
          JATraiter := NbJoursPri - NBJ;

          if (((NBJourspri < TP.getvalue('PCN_JOURS')) or { on a déjà payé sur d'autres acquis }
            (TP.getvalue('PCN_JOURS') > (CongeAcquis - APayes)))) then { on ne pourra pas tt payer sur la mvt acquis }
          begin
            DupliquerMvtPris(TP, TA, Tb_pris, MS, DX, JTraites, JATraiter, Next, DtCloture, DebMoisPrec, FinMoisPrec, ModeValo);
            TP := Tb_Pris.findnext([''], [''], True);
            if TP <> nil then TP := Tb_Pris.findnext([''], [''], True); { pour se retrouver sur TU }
            if TP = nil then break; { ne doit jamais arriver }
            CumMs := 0;
            CumX := 0;
            NoP := -1;
          end
          else
          begin
            TP.putvalue('PCN_PERIODEPAIE', TA.getvalue('PCN_LIBELLE'));
            TP.putValue('PCN_PERIODEPY', TA.getvalue('PCN_PERIODECP'));
            TP.PutValue('PCN_PERIODECP', TP.getValue('PCN_PERIODEPY'));
            TP.PutValue('PCN_CODETAPE', 'P');
            TP.PutValue('PCN_DATEPAIEMENT', DTfin);
            if VH_Paie.PGEcabMonoBase then TP.PutValue('PCN_EXPORTOK', 'X'); { PT99 }
            CumMs := MS;
            CumX := DX;
            NoP := -1;
          end;
        end { End du (Mvt = 'REL') }
        else
        begin
          { calcul du maintien}
          if (NbJSem = 5) and (Jouvres <> 0) then CumMS := CumMS + (cumul12 * Nbj / Jouvres); { PT-5 }
          if (NbJsem = 6) and (Jouvrables <> 0) then CumMS := CumMs + (cumul12 * Nbj / Jouvrables); { PT-5 }
          {PT61-4 le cumul 12 pouvant ne pas être encore alimenté par le programme,}
          {   un recalcul des CP est nécessaire}
          if Cumul12 = 0 then TopRecalculCPBull := True;
          { Calcul au dixième }
          CumX := CumX + (Nbj * ValoJX[NoP]);
          TP.putvalue('PCN_PERIODEPAIE', 'Période du ' + datetostr(DTDebutP) + ' au ' + datetostr(DTFinP));
          TP.putvalue('PCN_DATEPAIEMENT', DTFinP);
          TP.putValue('PCN_PERIODEPY', TA.getvalue('PCN_PERIODECP'));
          TP.PutValue('PCN_PERIODECP', TP.getValue('PCN_PERIODEPY'));
          TP.PutValue('PCN_CODETAPE', 'P');
          if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK', 'X'); { PT99 }
        end;
        { DEB PT75-2 }
        if (Arrondi(Congeacquis, DCP) = Arrondi(TA.getvalue('PCN_APAYES'), DCP)) then
        begin
          Ordre := TA.getvalue('PCN_ORDRE');
          if Mvt = 'REL' then
          begin
            TA := T_ACQUIS.findFirst([''], [''], True);
            if (TA.GetValue('PCN_TYPECONGE') = 'REL') and (Ordre = TA.getvalue('PCN_ORDRE')) then
              TA := T_ACQUIS.findNext([''], [''], True);
          end
          else
            TA := T_ACQUIS.findnext([''], [''], True);
        end;
        { FIN PT75-2 }
        Nbjourspri := Arrondi(NbJourspri - Nbj, DCP); //PT-30 Ajout Arrondi
      end; //while ta <> nil and nbjourspris > 0
      if TP = nil then exit;
      // PT108 Absence := CalculValoAbsence(PaieSalMaint,TP.getValue('PCN_JOURS'),TP.getValue('PCN_DATEDEBUTABS'),Debmoisprec, Finmoisprec, TypeCP);
      EcritMvtPris(TP, CumMs, CumX, CumMs, ModeValo); { PT108 }
      CumX := 0;
      CumMS := 0;
    end;
    NoP := -1;
    TP := Tb_Pris.findnext([''], [''], True);
  end;
  AgregeCumulPris(Tb_Pris);

  {if tob_delta <> nil then PT106 Mise en commentaire
  begin
    tob_delta.free;
    tob_delta := nil;
  end;                 }
  TobAcquis := T_Acquis;
  JOuvres := 0;
  JOuvrables := 0;
  DTFin := 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : Rend le mode de valorisation de l'imdemnité CP
Suite ........ : D : Dixième de la base
Suite ........ : M : Maintien du salaire
Suite ........ : X : Au plus avantageux
Mots clefs ... : PAIE;CP
*****************************************************************}

function TrouveModeValorisation(T_etab, T_Sal: tob): string;
begin
  result := 'X';
  { PT36 Passage en paramètre de la tob salarie suppr. de la requête }
  if T_Sal.GetValue('PSA_CPTYPEMETHOD') = 'PER' then
    Result := T_Sal.GetValue('PSA_VALORINDEMCP')
  else
    Result := T_Etab.getvalue('ETB_VALORINDEMCP');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Recupère la méthode de valorisation au maintien
Mots clefs ... : PAIE;CP
*****************************************************************}
{ DEB PT34-1 Intégration de la méthode de valorisation au maintien cas personnalisé salarié }

procedure RendMethodeValoDixieme(var NBjSem: integer; var MValoMs: string; var ValoDxmn: double);
var
  T_etab: tob;
begin
  T_Etab := TOB_Etablissement.FindFirst(['ETB_ETABLISSEMENT'], [Tob_Salarie.getvalue('PSA_ETABLISSEMENT')], True);
  if T_etab <> nil then
  begin
    NbJSem := T_Etab.getvalue('ETB_NBJOUTRAV');
    MValoMs := T_ETAB.getvalue('ETB_MVALOMS');
    ValoDxmn := T_ETAB.getvalue('ETB_VALODXMN');
  end;
  if Tob_Salarie.GetValue('PSA_CPTYPEVALO') = 'PER' then
  begin
    MValoMs := Tob_Salarie.getvalue('PSA_MVALOMS');
    ValoDxmn := Tob_Salarie.getvalue('PSA_VALODXMN');
  end;
end;
{ FIN PT34-1 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 28/01/2003
Description .. : En fonction du nbre de jours travaillés de l'établissement on
Suite ........ : calcul le nbre de jours de référence pour le calcul cp
Mots clefs ... : PAIE;CP
*****************************************************************}

function JoursReference(NbJoursAcquis: double): Integer;
begin
  if NbJoursAcquis > 0 then //DEB PT-8
    result := StrToInt(FloatToStr(Arrondi(NbJoursAcquis * 12, 0)))
  else
    result := 0; //FIN PT-8
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : balayage de la tob des acquis pour savoir si je pourrai payer
Suite ........ : en entier le mvt pris
Mots clefs ... : PAIE;CP
*****************************************************************}

function AssezAcquis(T_Acquis: tob; Nbjourspri: double; var prispayes, acquisres: double; calculAcquis: boolean): boolean;
var
  TA: tob;
begin
  result := false;
  if CalculAcquis then
  begin
    TA := T_ACQUIS.findfirst([''], [''], True);
    while TA <> nil do
    begin
      Acquisres := Acquisres + TA.GetValue('PCN_JOURS') - TA.GetValue('PCN_APAYES');
      TA := T_ACQUIS.findnext([''], [''], True);
    end;
    exit;
  end;
  if (arrondi(Acquisres, 2) >= Arrondi((Nbjourspri + prispayes), 2)) then
  begin
    result := true;
    prispayes := prispayes + Nbjourspri;
  end
  else result := false;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 03/02/2003
Description .. :  Pour chaque mouvement ayant fait l'objet d'un éclatement,
Suite ........ : on agrège les calculs détails pour les stocker au niveau du
Suite ........ : mvt original
Suite ........ : On alimente aussi les zones periode CP, période PY et date
Suite ........ : début, date fin du Mvt Général
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AgregeCumulPris(T_Pris: Tob);
var
  S_Absence, S_ValoX, S_ValoMS, S_Valoretenue, S_TotValo, S_TotAbsence: double;
  code, ordre, PY, PerCP: integer;
  TP, TA: Tob;
  Top_manuel, Top_ManuelA: boolean;
  Salarie: string;
begin
  if not ((Assigned(T_Pris)) and (T_Pris.detail.count > 0)) {PT113 notvide(T_Pris, true) } then exit;
  TP := T_PRIS.findfirst(['PCN_CODERGRPT'], [-1], True);
  while TP <> nil do
  begin
    PY := 99;
    PerCP := 99;
    Salarie := TP.getvalue('PCN_SALARIE');
    Ordre := TP.getvalue('PCN_ORDRE');
    { il s'agit d'un mvt éclaté, je cumule les mt au niveau du pal }
    code := TP.getvalue('PCN_ORDRE');
    S_Absence := 0;
    S_TotAbsence := 0;
    S_ValoX := 0;
    S_ValoMS := 0;
    S_Valoretenue := 0;
    S_TotValo := 0;
    Top_manuel := false;
    Top_ManuelA := false;
    TA := T_PRIS.findfirst(['PCN_CODERGRPT'], [code], True);
    while TA <> nil do
    begin
      S_Absence := S_Absence + TA.GetValue('PCN_ABSENCE');
      S_ValoX := S_ValoX + TA.GetValue('PCN_VALOX');
      S_ValoMS := S_ValoMS + TA.GetValue('PCN_VALOMS');
      S_Valoretenue := S_Valoretenue + TA.GetValue('PCN_VALORETENUE');
      if TA.GetValue('PCN_MODIFVALO') = 'X' then
      begin
        S_TotValo := S_TotValo + TA.GetValue('PCN_VALOMANUELLE');
        Top_manuel := true;
      end
      else
        S_TotValo := S_TotValo + TA.GetValue('PCN_VALORETENUE');
      if TA.GetValue('PCN_MODIFABSENCE') = 'X' then
      begin
        S_TotAbsence := S_TotAbsence + TA.GetValue('PCN_ABSENCEMANU');
        Top_manuelA := true;
      end
      else
        S_TotAbsence := S_TotAbsence + TA.GetValue('PCN_ABSENCE');
      TP.PutValue('PCN_DATEPAIEMENT', TA.Getvalue('PCN_DATEPAIEMENT'));
      TP.putValue('PCN_DATEDEBUT', TA.Getvalue('PCN_DATEDEBUT'));
      TP.putValue('PCN_DATEFIN', TA.Getvalue('PCN_DATEFIN'));
      if PY > TA.getvalue('PCN_PERIODEPY') then
        PY := TA.getvalue('PCN_PERIODEPY');
      if PerCP > TA.getvalue('PCN_PERIODECP') then
        PerCP := TA.getvalue('PCN_PERIODECP');
      TA := T_PRIS.findNext(['PCN_CODERGRPT'], [code], True);
    end;
    TP.PutValue('PCN_ABSENCE', S_Absence);
    TP.PutValue('PCN_VALOX', S_ValoX);
    TP.PutValue('PCN_VALOMS', S_ValoMS);
    TP.PutValue('PCN_VALORETENUE', S_Valoretenue);
    TP.PutValue('PCN_PERIODECP', PerCP);
    TP.PutValue('PCN_PERIODEPY', PY);
    if Top_Manuel then
    begin
      TP.PutValue('PCN_VALOMANUELLE', S_totValo);
      TP.PutValue('PCN_MODIFVALO', 'X');
      TP.PutValue('PCN_BASE', S_TotValo);
    end
    else
    begin
      TP.PutValue('PCN_VALOMANUELLE', 0);
      TP.PutValue('PCN_MODIFVALO', '-');
      TP.PutValue('PCN_BASE', S_Valoretenue);
    end;
    if Top_ManuelA then
    begin
      TP.PutValue('PCN_ABSENCEMANU', S_totAbsence);
      TP.PutValue('PCN_MODIFABSENCE', 'X');
    end
    else
      TP.PutValue('PCN_ABSENCEMANU', 0);
    TP := T_PRIS.findfirst(['PCN_CODERGRPT'], [-1], True);
    while tp.getvalue('PCN_ORDRE') <> ordre do
      TP := T_PRIS.findnext(['PCN_CODERGRPT'], [-1], True);
    { repositionnement sur le Tp avant lecture du suivant }
    TP := T_PRIS.findnext(['PCN_CODERGRPT'], [-1], True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : procedure d'éclatement du CP pri en pls mvts pri pour
Suite ........ : chaque période consommé
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure DupliquerMvtPris(var TP, TA, Tb_Pris: tob; MS, DX, JTraites, JATraiter: double; var Next: integer; DTcloture, DD, DF: TdateTime; ModeValo: string);
var
  TD, TU: Tob;
  TypeCp: string;
//  Absence: double; PT108
begin
  TD := TOB.Create('ABSENCESALARIE', Tb_Pris, (TP.GetIndex) + 1);
  TD.Dupliquer(TP, FALSE, TRUE, TRUE);
  TD.putvalue('PCN_ORDRE', Next);
  Next := Next + 1;
  TD.Putvalue('PCN_CODETAPE', 'P');
  if VH_Paie.PGEcabMonoBase then TD.putvalue('PCN_EXPORTOK', 'X'); { PT99 }
  TD.PutValue('PCN_MVTDUPLIQUE', 'X');
  if TP.getvalue('PCN_TYPECONGE') = 'PRI' then {DEB PT-9}
    TD.PutValue('PCN_DATEPAIEMENT', PlusDate(DTcloture, -1, 'A'))
  else {FIN PT-9}
    TD.PutValue('PCN_DATEPAIEMENT', DF); {PT-3  ligne ajouté le 29/06/2001 }
  TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
  TD.putvalue('PCN_JOURS', Arrondi(JTraites, DCP)); // PT-22
  if Tp.getvalue('PCN_TYPECONGE') = 'SLD' then TypeCp := 'SLD'
  else TypeCp := 'PRI';
 // Absence := CalculValoAbsence(JTraites, DD, DD, DF, typeCp);  PT108
  EcritMvtPris(TD, Ms, DX, Ms, ModeValo); { PT108 }

  TU := TOB.Create('ABSENCESALARIE', Tb_Pris, (TP.GetIndex) + 2);
  TU.Dupliquer(TP, FALSE, TRUE, TRUE);
  TU.putvalue('PCN_ORDRE', Next);
  Next := Next + 1;
  TU.Putvalue('PCN_CODETAPE', 'P');
  if VH_Paie.PGEcabMonoBase then TU.putvalue('PCN_EXPORTOK', 'X'); { PT99 }
  TU.PutValue('PCN_MVTDUPLIQUE', 'X');
  TU.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_ORDRE'));
  TU.PutValue('PCN_JOURS', Arrondi(JATraiter, DCP)); // PT-22
  if (TP.GetValue('PCN_CODERGRPT') <> TP.getvalue('PCN_ORDRE'))
    and (TP.GetValue('PCN_CODERGRPT') <> 0) then
  begin
    TU.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_CODERGRPT'));
    TD.PutValue('PCN_CODERGRPT', TP.getvalue('PCN_CODERGRPT'));
    Tp.free;
  end
  else
  begin
    TP.PutValue('PCN_CODERGRPT', -1);
    TP.PutValue('PCN_PERIODEPAIE', ' ');
    TP.PutValue('PCN_PERIODEPY', -1);
    TP.PutValue('PCN_PERIODECP', -1);
    TP.PutValue('PCN_CODETAPE', 'P'); {PT-3  ligne ajouté le 28/06/2001 }
    if VH_Paie.PGEcabMonoBase then TP.putvalue('PCN_EXPORTOK', 'X'); { PT99 }
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : charge les ajustements de congés payés négatif
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure ChargeTobDelta(Salarie: string; DateF: TDateTime; var Lctob_delta: tob); { PT105 } { PT106 }
var
  st: string;
begin
  { alimentation de la tob des ajustements pour chercher le delta }
  st := 'SELECT * FROM ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + SALARIE + '" AND PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'AND PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" ' +
    'AND (( PCN_TYPECONGE="AJP" OR PCN_TYPECONGE="AJU") AND PCN_SENSABS = "-") ' +
    'AND PCN_DATEVALIDITE <= "' + usdatetime(DateF) + '" ' + { PT105 } { PT106 }
    'ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then // pas de congés acquis
  begin
    LcTob_delta := Tob.create('Tob Delta', nil, -1);               // PT106
    LcTob_delta.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);  // PT106
  end;
  Ferme(Q);
}
  LcTob_delta := Tob.create('Tob Delta', nil, -1);
  LcTob_delta.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
if LcTob_delta.detail.count <= 0 then FreeAndNil (LcTob_Delta);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : Calcul du nombre de jours AJP, AJU sens négatif - posés
Suite ........ : sur le période
Mots clefs ... : PAIE;CP
*****************************************************************}

function CalculDelta(Lctob_delta: tob; DTclotureP, dateFB: tdatetime; Periode: integer): double; { PT106 }
var
  Datevalidite, datedebut, datefin: tdatetime;
  Nbjours, Cumjours: double;
  aa, mm, jj: word;
  i: integer;
  TA: Tob;
begin
  CumJours := 0;
  CalculPeriode(DTClotureP, DateFB);
  datefin := DTclotureP;
  decodedate(datefin, aa, mm, jj);
  datedebut := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT71 }
  i := 0;
  while i < periode do
  begin
    datefin := datedebut - 1;
    decodedate(datefin, aa, mm, jj);
    datedebut := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT71 }
    i := i + 1;
  end;
  if Lctob_delta <> nil then { PT106 }
  begin
    TA := Lctob_delta.Findfirst([''], [''], True); { PT106 }
    while TA <> nil do
    begin
      NbJours := TA.getvalue('PCN_JOURS');
      DateValidite := TA.Getvalue('PCN_DATEVALIDITE');
      if ((DateValidite >= Datedebut) and (Datevalidite <= DateFin)) then
        CumJours := CumJours + Nbjours;
      TA := Lctob_delta.Findnext([''], [''], True); { PT106 }
    end;
  end;
  result := CumJours;
end;

{ DEB PT106 }

procedure DeSoldeTobDelta(Salarie: string; DtSld: TDateTime; StWhere: string);
var
  st: string;
  TA: Tob;
begin
  if DtSld = idate1900 then exit;
  FreeAndNil(Tob_delta);

  st := 'SELECT * FROM ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + SALARIE + '" AND PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'AND PCN_CODETAPE = "S" AND PCN_PERIODECP < 2 ' + StWhere +
    'AND (( PCN_TYPECONGE="AJP" OR PCN_TYPECONGE="AJU") AND PCN_SENSABS = "-") ' +
    'AND PCN_DATEVALIDITE <= "' + usdatetime(DtSld) + '" ' + { PT105 }
    'ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then // pas de congés acquis
  begin
    Tob_delta := Tob.create('Tob delta', nil, -1);
    Tob_delta.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
  end;
  Ferme(Q);
}
  Tob_delta := Tob.create('Tob delta', nil, -1);
  Tob_delta.LoadDetailDBFROMSQL('ABSENCESALARIE', St);
if Tob_Delta.detail.count <= 0 then FreeAndNil (Tob_Delta);

  if Assigned(Tob_delta) then
  begin
    TA := tob_delta.Findfirst([''], [''], True);
    while TA <> nil do
    begin
      if TA.Getvalue('PCN_DATEVALIDITE') <= DtSld then
        TA.putValue('PCN_CODETAPE', '...');
      TA := Tob_delta.Findnext([''], [''], True);
    end;
    TOB_Delta.SetAllModifie(TRUE);
    TOB_Delta.InsertOrUpdateDB(true);
  end;
end;
{ FIN PT106 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : calcul de l'indemnité journalière d'un jour au dixième pour la
Suite ........ : période passée
Suite ........ :  période = 0 => période courante
Suite ........ :  période = 1 => période précédente
Mots clefs ... : PAIE;CP
*****************************************************************}

function CalculValoX(DTclotureP, dateFB, DateDB: tdatetime; periode: integer; NbJoursRef: double; Typecp: string; T_Acquis, Lctob_delta: tob): double; { PT106 }
var
  Mvt, sens, TypeImpute: string;
  Datevalidite, datedebut, datefin: TDateTime;
  Arr, NMois, Base, CumSBA, CumSMA, Nbjours, Cumjours: double;
  aa, mm, jj: word;
  i, PB: integer;
  TA: Tob;
begin
  CumJours := 0;
  PB := CalculPeriode(DTClotureP, DateFB);
  if Typecp = 'PRO' then Datefin := DateFB
  else Datefin := DTclotureP;
  decodedate(datefin, aa, mm, jj);
  datedebut := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT71 }
  i := 0;
  while i < periode do
  begin
    datefin := datedebut - 1;
    decodedate(datefin, aa, mm, jj);
    datedebut := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT71 }
    i := i + 1;
  end;
  CumSBA := 0;
  CumSMA := 0;
  if T_Acquis <> nil then
  begin
    TA := T_ACQUIS.Findfirst([''], [''], True);
    while TA <> nil do
    begin
      Mvt := TA.getvalue('PCN_TYPECONGE');
      Sens := TA.getvalue('PCN_SENSABS');
      DateValidite := TA.Getvalue('PCN_DATEVALIDITE');
      NMois := TA.Getvalue('PCN_NBREMOIS');
      BASE := TA.Getvalue('PCN_BASE');
      NbJours := TA.getvalue('PCN_JOURS');
      // DEB PT86
      if TA.getvalue('PCN_TYPEIMPUTE') <> NULL then TypeImpute := TA.getvalue('PCN_TYPEIMPUTE')
      else TypeImpute := '';
      //      TypeImpute := TA.getvalue('PCN_TYPEIMPUTE'); { PT74-3 }
      // FIN PT86
            { on ne prend pas le mvt d'acquis du bulletin courant. }
            { Necessaire s'ils'agit d'un mvt de type SLD }
      if Typecp <> 'PRO' then //****SB <==> PH PT83 Remis idem V5
      begin { Sauf l'arrondi qui doit être cumulé à CumJours }
        if ((DateValidite >= DateDB) and (DateValidite <= DateFB) and (Mvt <> 'ARR')) then //PT-23 Ajout cond.(Mvt<>'ARR')
        begin
          TA := T_ACQUIS.Findnext([''], [''], True);
          continue;
        end;
      end;
      { PT-28 on ne tient pas compte des ACA dans le decompte des jours acquis }
      { PT52  on ne tient pas compte des AJP dans le calcul du dixieme }
      { La tob acquis charge aussi les AJP et AJU + }
      { Donc les AJU + sont cumulés par la tob acquis }
      {   et les AJU - sont cumulés par la tob_Delta }
      { PT101 On ne tient pas compte des ACS dans le décompte des jours acquis clause (Mvt <> 'ACS') }
      if (Mvt <> 'REL') and (Mvt <> 'ACA') and (Mvt <> 'ACS') and (Mvt <> 'AJP')
        and (Mvt <> 'ACF') // PT128
        and (TypeImpute <> 'REL') and (TypeImpute <> 'ACA') and (TypeImpute <> 'ACS') then { PT74-3 } { PT101 } { PT112 }
      begin
        if ((DateValidite >= Datedebut) and (Datevalidite <= DateFin)) then
        begin
          if Sens = '+' then
          begin
            CumSBA := CumSBA + BASE;
            CumSMA := CumSMA + NMois;
            CumJours := CumJours + Nbjours;
          end
          else
            if Sens = '-' then
            begin
              CumSBA := CumSBA - BASE;
              CumSMA := CumSMA - NMois;
              CumJours := CumJours - Nbjours;
            end;

      { DEB PT103 On ne tient pas compte de l'arrondi du solde }
          if (Mvt = 'ARR') and (PB = Periode) then
          begin
            CumSBA := CumSBA - BASE;
            CumSMA := CumSMA - NMois;
            CumJours := CumJours - Nbjours;
          end;
      { FIN PT103 }
      { DEB PT119 Dplct de la fin de condition pour intégrer l'exlusion de l'arrondi dans la condition de période }
        end;
      end;
      { FIN PT119}
    { DEB PT111 }
      if (VH_Paie.PGChr6Semaine) and (Mvt = 'ACS') and ((DateValidite >= Datedebut) and (Datevalidite <= DateFin)) then
      begin
        if Sens = '+' then
        begin
          CumSBA := CumSBA + BASE;
          CumSMA := CumSMA + NMois;
          CumJours := CumJours + Nbjours;
        end
        else
          if Sens = '-' then
          begin
            CumSBA := CumSBA - BASE;
            CumSMA := CumSMA - NMois;
            CumJours := CumJours - Nbjours;
          end;
      end;
    { FIN PT111 }
      TA := T_ACQUIS.Findnext([''], [''], True);
    end; //FIN du While
  end;
  {PT52  on ne tient pas compte de tous les  AJP dans le calcul du dixieme }
  { NB : La Tob_delta charge que les AJP et AJU -}
  if Lctob_delta <> nil then { PT106 }
  begin
    TA := Lctob_delta.Findfirst(['PCN_TYPECONGE'], ['AJU'], True); //PT52 ajout PCN_TYPECONGE { PT106 }
    while TA <> nil do
    begin
      Sens := TA.getvalue('PCN_SENSABS');
      Mvt := TA.getvalue('PCN_TYPECONGE');
      DateValidite := TA.Getvalue('PCN_DATEVALIDITE');
      NMois := TA.Getvalue('PCN_NBREMOIS');
      BASE := TA.Getvalue('PCN_BASE');
      NbJours := TA.getvalue('PCN_JOURS');
      // DEB PT86
      if TA.getvalue('PCN_TYPEIMPUTE') <> NULL then TypeImpute := TA.getvalue('PCN_TYPEIMPUTE')
      else TypeImpute := '';
      //      TypeImpute := TA.getvalue('PCN_TYPEIMPUTE'); { PT74-3 }
      // FIN PT86
      if ((DateValidite >= Datedebut) and (Datevalidite <= DateFin)) and (Sens = '-') and (TypeImpute = 'ACQ') then { PT74-3 }
      begin
        CumSBA := CumSBA - BASE;
        CumSMA := CumSMA - NMois;
        CumJours := CumJours - Nbjours;
      end;
      TA := Lctob_delta.Findnext(['PCN_TYPECONGE'], ['AJU'], True); { PT106 }
    end;
  end;
  if CumSBA < 0 then CumSBA := 0;
  if CumSMA < 0 then CumSMA := 0;
  if CumJours < 0 then CumJours := 0;
  if (((CumSBA = 0) or (CumSMA = 0)) and (typecp <> 'SLD')) then
  begin
    result := 0;
    exit;
  end;
  { valorisation d'un jour de congé acquis sur période au dixième }
  if ((Typecp = 'SLD') and ((DateSSalarie <= DateFB) and (DateSSalarie >= DateFB))) then
  begin
    if (PB = Periode) then { on est sur la période en cours. }
    begin { Ne pas oublier ce qu' on est en train d'acquérir }
      ValoXP0N := CumSBA; { intégration de ce qu'on va acquérir sur ce bulletin }
      ValoXP0D := CumSMA;
      { si insertion rubrique alimentant cumul 11 ds bulletin .. }
      BCPACQUIS := Arrondi(RendCumulSalSess('11'), DCP);
      if (10 * (arrondi(CumJours + JCPACQ, DCP))) = 0 then {Deb PT-2} { PT103 }
        result := 0
      else {Fin PT-2}
      begin
        if VH_Paie.PGChr6Semaine = False then { PT111 }
        begin
          { DEB PT103 On calcul un arrondi seulement sur les acquis }
          Arr := 1 + Int(CumJours + JCPACQ) - (CumJours + JCPACQ);
          if Arr = 1 then Arr := 0;
          if Arrondi(CumJours + JCPACQ + Arr, DCP) <> 0 then // PT124 Au cas où
            result := ((CumSBA + BCPACQUIS)) / (10 * Arrondi(CumJours + JCPACQ + Arr, DCP))
          else result := 0;
          { FIN PT103 }
        end
        else { DEB PT111 }
        begin
          if ((CumSMA + MCPACQUIS) * NbJoursRef) <> 0 then // PT124
          // base CP / nbre de mois * 12 / nbre jours référence * nbre de jours à solder non arrondi * 0.10
            result := ((CumSBA + BCPACQUIS) * 12) / (10 * (CumSMA + MCPACQUIS) * NbJoursRef) //* (CumJours + JCPACQ + JCPACS)
          else result := 0;
        end; { FIN PT111 }
      end;
    end
    else
    begin { on n'est pas sur la période en cours }
      if Cumjours = 0 then
        result := 0
      else
        result := CumSBA / (Cumjours * 10);
    end;
  end
  else { Calcul d'un PRI sur toutes les périodes }
    if (CumSMA * 10 * NbJoursRef) = 0 then
      result := 0
    else
      result := (12 * CumSBA) / (CumSMA * 10 * NbJoursRef);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 16/04/2004
Modifié le ... :   /  /
Description .. : Calcul de l'indemnité journalière imputée sur reliquat
Mots clefs ... : PAIE;CP
*****************************************************************}
{ DEB PT74-3 }

function CalculValoReliquat(T_Acquis, Lctob_delta: TOB; NbJoursRef: double): Double; { PT106 }
var
  TA: Tob;
  Base, Mois: Double;

begin
  Result := 0;
  if not Assigned(T_Acquis) then exit;
  Base := 0;
  Mois := 0;
  TA := T_ACQUIS.findfirst(['PCN_TYPECONGE', 'PCN_SENSABS'], ['REL', '+'], True);
  if Assigned(TA) then
  begin
    Base := TA.GetValue('PCN_BASE');
    Mois := TA.GetValue('PCN_NBREMOIS');
  end;
  TA := T_ACQUIS.findfirst(['PCN_TYPECONGE', 'PCN_TYPEIMPUTE', 'PCN_SENSABS'], ['AJU', 'REL', '+'], True);
  while Assigned(TA) do
  begin
    Base := Base + TA.GetValue('PCN_BASE');
    Mois := Mois + TA.GetValue('PCN_NBREMOIS');
    TA := T_ACQUIS.FindNext(['PCN_TYPECONGE', 'PCN_TYPEIMPUTE', 'PCN_SENSABS'], ['AJU', 'REL', '+'], True);
  end;
  if Assigned(Lctob_delta) then { PT106 }
  begin
    TA := Lctob_delta.findfirst(['PCN_TYPECONGE', 'PCN_TYPEIMPUTE', 'PCN_SENSABS'], ['AJU', 'REL', '-'], True); { PT106 }
    while Assigned(TA) do
    begin
      Base := Base - TA.GetValue('PCN_BASE');
      Mois := Mois - TA.GetValue('PCN_NBREMOIS');
      TA := Lctob_delta.FindNext(['PCN_TYPECONGE', 'PCN_TYPEIMPUTE', 'PCN_SENSABS'], ['AJU', 'REL', '-'], True); { PT106 }
    end;
  end;

  if (Base = 0) or (Mois = 0) or (NbJoursRef = 0) then
    Result := 0
  else
    Result := (arrondi((12 * Base) / (Mois * 10 * NbJoursRef), DCP));
end;
{ FIN PT74-3 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction qui valorise l'absence au maintient de salaire
Suite ........ : 1- Si on est en ouvrable, cum11/NbJouvrables * Nbj
Suite ........ : 2- Si on est en ouvré, cum11/NbJouvres * Nbj
Mots clefs ... : PAIE;CP
*****************************************************************}

function CalculValoAbsence(NbJ: double; DebMoisprec, FinMoisprec: Tdatetime; TypeCp: string): double;
var
  NbJSem: Integer; //ph
  cumul12, divr: double;
  JSouvres, JSouvrables, HVal: double;
  DebMoisCours, FinMoisCours: tdatetime;
begin
  divr := 0;
  cumul12 := Valcumuldate('12', DebMoisprec, FinMoisprec);
  if Cumul12 = 0 then Cumul12 := RendCumulSalSess('12');
  if FirstBull then Cumul12 := RendCumulSalSess('12'); { PT76 }
  NbJSem := T_Etab.getvalue('ETB_NBJOUTRAV');
  if Typecp = 'SLD' then
  begin
    DebMoisCours := FinMoisprec + 1;
    FinMoisCours := FindeMois(DebMoisCours);
    { PT48 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables }
    { Dans ce cas seuls les jours nous intéresse.. }
    { JSouvrables := NombreJoursOuvresOuvrablesMois(T_Etab,DebMoisCours,FinMoisCours,false); }
    { JSouvres    := NombreJoursOuvresOuvrablesMois(T_Etab,DebMoisCours,FinMoisCours,true); }
    CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebMoisCours, FinMoisCours, TRUE, JSouvres, HVal);
    CalculVarOuvresOuvrablesMois(T_Etab, TOB_Salarie, nil, DebMoisCours, FinMoisCours, FALSE, JSouvrables, HVal);
  end;
  result := Cumul12 * Nbj;
  if (NbJSem = 5) and (TypecP = 'SLD') then divr := JSouvres;
  if (NbJSem = 6) and (TypecP = 'SLD') then divr := JSouvrables;
  if (NbJSem = 5) and (TypecP = 'PRI') then divr := Jouvres;
  if (NbJSem = 6) and (TypecP = 'PRI') then divr := Jouvrables;
  if divr <> 0 then result := result / divr
  else result := 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;CP
*****************************************************************}

function ReCalculeValoAbsence(Tb_pris: tob; Valo, PaieSalMaint: string; Mul: double; typecp: string): boolean;
var
  T: Tob;
  MS, cumul12, ValeurJ: double;
 // Datedebut, FinMoisprec, Debmoisprec: TDateTime;   PT108
  modif: boolean;
begin
  result := false;
  modif := false;
  {PT108 Mise en commentaire
  Datedebut := DebutDeMois(DateFin);
  FinMoisprec := datedebut - 1;
  Debmoisprec := DebutdeMois(FinMoisPrec);
  cumul12 := Valcumuldate('12', DebMoisprec, FinMoisprec);
  if Cumul12 <= 0 then Cumul12 := RendCumulSalSess('12');
  if Mul <> 0 then ValeurJ := Cumul12 / Mul
  else ValeurJ := 0;}
  T := Tb_pris.findfirst([''], [''], true);
  while T <> nil do
  begin
    if T.getvalue('PCN_CODERGRPT') = -1 then
    begin
      T := Tb_pris.findnext([''], [''], true);
      continue;
    end;
  { DEB PT108 }
    Cumul12 := RecupSalMaintien(PaieSalMaint, T.getvalue('PCN_DATEDEBUTABS'), T.getvalue('PCN_DATEDEBUT'));
    if Cumul12 <= 0 then Cumul12 := RendCumulSalSess('12');
    if Mul <> 0 then ValeurJ := Cumul12 / Mul
    else ValeurJ := 0;
  { FIN PT108 }
    if ((T.getvalue('PCN_ABSENCE') = arrondi(T.getvalue('PCN_JOURS') * ValeurJ, DCP)) and (typecp = 'PRI')) then
      exit;
    modif := true;
    MS := arrondi(T.getvalue('PCN_JOURS') * ValeurJ, DCP);
    T.putvalue('PCN_ABSENCE', MS);
    T.putvalue('PCN_VALOMS', MS);
    if Valo = 'M' then T.putvalue('PCN_VALORETENUE', MS)
    else
      if Valo <> 'D' then
      begin
        if T.getvalue('PCN_VALOX') > MS then T.putvalue('PCN_VALORETENUE', T.getvalue('PCN_VALOX'))
        else T.putvalue('PCN_VALORETENUE', MS);
      end;
    T.PutValue('PCN_BASE', T.getvalue('PCN_VALORETENUE'));
    T := Tb_pris.findnext([''], [''], true);
  end;
  result := modif;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : Alimentation des champs tables sur mvt CP PRI
Suite ........ : maj du maintien, dixième, absence
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure EcritMvtPris(T: tob; CumMs, CumX, Absence: double; Modevalo: string);
begin
  T.putvalue('PCN_VALOX', Arrondi(CumX, DCP));
  T.putvalue('PCN_VALOMS', Arrondi(CumMS, DCP));
  T.PutValue('PCN_DATEFIN', DTfin);
  T.PutValue('PCN_DATEDEBUT', DTDeb);

  if modeValo = 'M' then T.putvalue('PCN_VALORETENUE', Arrondi(CumMS, DCP))
  else
    if ModeValo = 'D' then T.putvalue('PCN_VALORETENUE', Arrondi(CumX, DCP))
    else
      if CumX > CumMS then T.putvalue('PCN_VALORETENUE', Arrondi(CumX, DCP))
      else
        T.putvalue('PCN_VALORETENUE', Arrondi(CumMS, DCP));
  T.PutValue('PCN_ABSENCE', Arrondi(Absence, DCP));
  T.PutValue('PCN_BASE', T.getvalue('PCN_VALORETENUE'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 04/10/2002
Modifié le ... :   /  /
Description .. : Alimente la tob des absences hors CP
Suite ........ : Renvoie True si absence non payé existe
Mots clefs ... :
*****************************************************************}
// d PT79
//function RecupereAbsences(Var Tob_ABS:tob;Datef:tdatetime;Salarie,Action:string):boolean;

{PT131
function RecupereAbsences(var Tob_ABS: tob; Datef: tdatetime; Salarie, Action: string; GenereMaintien, IJSS: boolean): boolean;}
function RecupereAbsences(var Tob_ABS: tob; Datef: tdatetime; Salarie, Action: string; GenereMaintien, IJSS: boolean; typeAbs : string): boolean;
// f PT79
var
  st: string;
  T: Tob;
  Q: TQuery;
begin
  result := False;
  { PT36 Mise en commentaire puis supprimer : intégré dans la requête de chargement}
  if Action = 'M' then
  begin // PT79
    st := 'SELECT * FROM ABSENCESALARIE WHERE' +
      ' PCN_SALARIE = "' + Salarie + '"' + { PT37 Ajout clause AND PCN_VALIDRESP = "VAL" }
      ' AND (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" AND PCN_TYPECONGE <>  "PRI" AND PCN_VALIDRESP="VAL")' + //PT32 Ajout AND PCN_SENSABS="-"
      ' AND ((PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '"' +
      ' AND PCN_DATEPAIEMENT <= "' + usdatetime(10) + '")' +
      // d PT79
    ' OR (PCN_DATEPAIEMENT="' + usdatetime(Datef) + '" AND PCN_CODETAPE="P"))';
//d PT94
//    if (IJSS) then st := st + ' AND PCN_GESTIONIJSS = "X"';
{ d PT110    if (IJSS) and (GenereMaintien) then st := st + ' AND PCN_GESTIONIJSS = "X"';
// f PT94
    if (GenereMaintien) and not (IJSS) then st := st + ' AND PCN_GESTIONIJSS <> "X"'; f PT110 }
    st := st + ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE' + { PT36 }
    // f PT79
    // DEB PT107
    ' WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE <> "" AND PMA_RUBRIQUE is not null) OR (PMA_RUBRIQUEJ <> "" AND PMA_RUBRIQUEJ is not null)) ' +
      ' ORDER BY PCN_TYPECONGE,PCN_DATEVALIDITE'
    // FIN PT107
  end // PT79
  else
  begin // PT79
    st := 'SELECT * FROM ABSENCESALARIE WHERE' +
      ' PCN_SALARIE = "' + Salarie + '"' + { PT37 Ajout clause AND PCN_VALIDRESP = "VAL" }
      ' AND (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-" AND PCN_TYPECONGE <>  "PRI" AND PCN_VALIDRESP="VAL")' + //PT32 Ajout AND PCN_SENSABS="-"
      ' AND PCN_ETATPOSTPAIE <> "NAN" AND PCN_DATEVALIDITE <= "' + usdatetime(Datef) + '"' + { PT116 }
      // d PT79
    ' AND PCN_DATEPAIEMENT <= "' + usdatetime(10) + '"';
// d PT94
//    if (IJSS) then st := st + ' AND PCN_GESTIONIJSS = "X"';
    if (IJSS) and (GenereMaintien) then st := st + ' AND PCN_GESTIONIJSS = "X"';
// f PT94
    if (GenereMaintien) and not (IJSS) then
      st := st + ' AND PCN_GESTIONIJSS <> "X"';

    st := st + ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE' + { PT36 }
    // f PT79
    // DEB PT107
{d PT131
      ' WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE <> "" AND PMA_RUBRIQUE is not null) OR (PMA_RUBRIQUEJ <> "" AND PMA_RUBRIQUEJ is not null)) ' +
      ' ORDER BY PCN_TYPECONGE,PCN_DATEVALIDITE';}
      ' WHERE ##PMA_PREDEFINI## (PMA_RUBRIQUE <> "" AND PMA_RUBRIQUE is not null) OR (PMA_RUBRIQUEJ <> "" AND PMA_RUBRIQUEJ is not null) ' ;
    if (GenereMaintien) and (typeAbs <> '') then
      st := st + 'AND PMA_TYPEABS = "'+ typeAbs +'"';

    st := st+') ORDER BY PCN_TYPECONGE,PCN_DATEVALIDITE';
// f PT131
    // FIN PT107
  end; // PT79
  Q := OpenSql(st, TRUE);
  if not Q.eof then
  begin
    {PT78
         Tob_ABS := Tob.create('ABSENCESALARIE',nil,-1);
    }
    if (Tob_ABS = nil) then // PT81
      Tob_ABS := Tob.create('Mes ABSENCESALARIE', nil, -1);
    // d PT81     Tob_ABS.LoadDetailDB('ABSENCESALARIE','','',Q,False);
    //   La Tob_Abs est en maj car dans le cas du traitement des IJ on récupère
    //   d'abord les absences non IJSS et ensuite les absences IJSS - C'est à partir
    //   Tob_abs que les ABSENCESALARIE sont maj en SauvegardeBul
    Tob_ABS.LoadDetailDB('ABSENCESALARIE', '', '', Q, True);
    //f PT81
    HABSPRIS := Tob_ABS.Somme('PCN_HEURES', [''], [''], False); { PT25 Affectation Var Heures Absence pris }
    JABSPRIS := Tob_ABS.Somme('PCN_JOURS', [''], [''], False); { PT25 Affectation Var Jours Absence pris }
    { PT36 Mise en commentaire puis supprimer : intégré dans la requête de chargement }
    { Si absence non payé existe renvoie true pour réintégration paye }
    if TOB_ABS.Findfirst(['PCN_CODETAPE'], ['...'], false) <> nil then result := true;
  end
  else
//PT118   if TOB_ABS <> nil then LibereTobAbs; { PT55 }
    if (not IJSS) and (TOB_ABS <> nil) then LibereTobAbs; { PT55 }

  Ferme(Q);
  if (Action = 'M') and (Tob_ABS <> nil) and (result) then
  begin { On réintègre toutes les absences même celles déjà payées }
    t := TOB_ABS.Findfirst(['PCN_CODETAPE'], ['P'], false);
    while t <> nil do
    begin
      T.putvalue('PCN_CODETAPE', '...'); { PT-13 }
      T.putvalue('PCN_DATEFIN', 0);
      T.putvalue('PCN_DATEPAIEMENT', 0);
      if (VH_Paie.PGEcabMonoBase) and (T.getValue('PCN_SAISIEDEPORTEE') = 'X') then T.putvalue('PCN_EXPORTOK', '-'); { PT99 }
      T := TOB_ABS.Findnext([''], [''], false)
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 28/01/2003
Description .. : Recherche des Mvt de congés payés intégrables
Mots clefs ... : PAIE;CP
*****************************************************************}

function RecupereCongespris(Tob_pris, TobSal: Tob; DtFinP: TDateTime): boolean;
var
  st: string;
begin
  st := 'SELECT * FROM ABSENCESALARIE WHERE' +
    ' PCN_SALARIE = "' + TobSal.getvalue('PSA_SALARIE') + '"' +
    ' AND PCN_TYPEMVT="CPA" ' + { PT32 Ajout } { PT37 Ajout clause AND PCN_VALIDRESP = "VAL" }
    ' AND ((PCN_TYPECONGE = "PRI" AND PCN_VALIDRESP = "VAL") OR PCN_TYPECONGE = "SLD")' +
    ' AND PCN_DATEVALIDITE <= "' + usdatetime(DTFinP) + '"' +
    ' AND PCN_DATEFIN <= "' + usdatetime(10) + '"' +
    ' ORDER BY PCN_DATEVALIDITE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if (Q.eof) then
    result := false
  else
  begin
    Tob_pris.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    JCPPAYESPOSES := Tob_pris.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['PRI'], False); //PT63
    result := true;
  end;
  Ferme(Q);
  }
  Tob_pris.LoadDetailDBFROMSQL('ABSENCESALARIE', St);
  if Tob_pris.Detail.count >= 1 then
  begin
    JCPPAYESPOSES := Tob_pris.Somme('PCN_JOURS', ['PCN_TYPECONGE'], ['PRI'], False); //PT63
    result := true;
  end
  else
    result := FALSE;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 28/01/2003
Description .. : En retour de l'option Voir Congés Payés (saisie du bulletin)
Suite ........ : on verifie si l'état de saisie des congés à changer :
Suite ........ : CREATION, MODIFICATION, SUPPRESSION
Mots clefs ... : PAIE;CP
*****************************************************************}

function ModificationCongesPris(Tob_pris, TobSal: Tob; DtFinP, Dateentree: tdatetime): boolean;
var
  st: string;
  TobInt, T, TP, TTest: Tob;
  Ladate: TDateTime;
begin
  TobInt := Tob.Create('Tob intermediaire', nil, -1);
  st := 'SELECT * FROM ABSENCESALARIE WHERE' +
    ' PCN_SALARIE = "' + TobSal.getvalue('PSA_SALARIE') + '"' +
    ' AND PCN_TYPEMVT="CPA" ' + { PT32 Ajout } { PT37 Ajout clause AND PCN_VALIDRESP = "VAL" }
    ' AND ((PCN_TYPECONGE = "PRI" AND PCN_VALIDRESP = "VAL") OR PCN_TYPECONGE = "SLD")' +
    ' AND PCN_DATEVALIDITE <= "' + usdatetime(DTFinP) + '"' +
    ' AND PCN_DATEFIN <= "' + usdatetime(10) + '"' +
    ' AND PCN_DATEMODIF >="' + ustime(DateEntree) + '"' +
    ' ORDER BY PCN_DATEVALIDITE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  result := false;
  if (not Q.eof) then
  begin
}
  TobInt.LoadDetailDBFROMSQL('tobfille intermédiaire', st);
  T := TobInt.FindFirst([''], [''], false);
//  Ferme(Q);
  while T <> nil do
  begin
    Ladate := T.getvalue('PCN_DATEMODIF');
    if Ladate > DateEntree then
    begin
      result := true;
//PT129      Tobint.free;
      FreeAndNil (TobInt);
      exit;
    end;
    T := TobInt.Findnext([''], [''], false);
  end;
// end;
  { Controle si il n'y a pas eu de suppression }
  TP := Tob.create('CONGES PRIS', nil, -1);
  if not RecupereCongespris(TP, TobSal, DtFinP) then
  begin
    TP.free;
    if not ((Assigned(Tob_pris)) and (Tob_pris.detail.count > 0)) then
       begin
       FreeAndNil (TobInt);
       exit;
       end;
    T := Tob_Pris.FindFirst([''], [''], false);
    while t <> nil do
    begin
      t.free;
      result := true;
      T := Tob_Pris.FindNext([''], [''], false);
    end;
    FreeAndNil (TobInt);
    exit;
  end;
  if Tob_Pris <> nil then
  begin
    T := Tob_Pris.FindFirst([''], [''], false);
    while T <> nil do
    begin
      TTest := TP.findfirst(['PCN_TYPEMVT', 'PCN_RESSOURCE', 'PCN_SALARIE', 'PCN_ORDRE'],
        [T.getvalue('PCN_TYPEMVT'), T.getvalue('PCN_RESSOURCE'),
        T.getvalue('PCN_SALARIE'), T.getvalue('PCN_ORDRE')], true);
      if TTest = nil then
      begin
        T.free;
        result := true;
        FreeAndNil (TobInt);
        exit;
      end;
      T := Tob_Pris.FindNext([''], [''], false);
    end;
  end;
  if TP <> nil then TP.free;
FreeAndNil (TobInt);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... :   /  /
Description .. : Calcule le nombre de jours ouvrés / ouvrables d'un mois
Mots clefs ... : PAIE;CP
*****************************************************************}
{ PT48 Modification de la function en procedure, recup en parametre heures et jours ouvrés ouvrables }

procedure CalculVarOuvresOuvrablesMois(TobEtab, TobSal, Tob_Cal: Tob; datedebut, datefin: tdatetime; Ouvres: boolean; var Nb_j, Nb_H: double);
var
  Jour_j: TDatetime;
  Numero, RH1, RH2: integer;
  Gerecp, c: string;
  T: Tob;
begin
  if TOBEtab = nil then exit;
  RH2 := 0;
  GereCp := TobEtab.getvalue('ETB_CONGESPAYES');
  if GereCp = 'X' then
  begin
    RH1 := strtoint(copy(TobEtab.getvalue('ETB_1ERREPOSH'), 1, 1));
    if ouvres then
    begin
      c := TobEtab.getvalue('ETB_2EMEREPOSH');
      if (length(c) > 0) then
        RH2 := strtoint(copy(TobEtab.getvalue('ETB_2EMEREPOSH'), 1, 1))
      else
        RH2 := 0;
      if RH2 = 0 then
        if RH1 > 1 then
          RH2 := RH1 - 1
        else
          RH2 := 7;
    end;
  end
  else
  begin
    if ouvres then
    begin
      RH1 := 7;
      RH2 := 1;
    end
    else
      RH1 := 1;
  end;
  Nb_j := 0;
  Nb_h := 0;
  Jour_j := datedebut;
  while Jour_j <= datefin do
  begin
    Numero := DayOfWeek(Jour_j);
    if ((Numero <> RH1) and (Numero <> RH2)) then
    begin
      Nb_j := Nb_j + 1;
      {DEB PT48 Calcul des heures ouvrés ou ouvrables}
      if Tob_Cal <> nil then
      begin
        if Numero = 1 then Numero := 7 else Numero := Numero - 1;
        T := Tob_Cal.FindFirst(['ACA_JOUR'], [Numero], False);
        if T <> nil then nb_h := Nb_H + T.GetValue('ACA_DUREE');
      end;
      {FIN PT48}
    end;
    Jour_j := Jour_j + 1;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 03/02/2003
Description .. : Recupere les mvts de CP PRI et SLD payés
Mots clefs ... : PAIE;CP
*****************************************************************}
{ PT56 Reprise syntaxe procedure V591 }

procedure RecupereCongesPrisPayes(LcTob_pris, LcTob_solde: tob; Datef: tdatetime; Salarie, TypeCp: string); { PT115-2 }
var
  st, StValidResp: string;
  T: tob;
  nb: double;
begin
  { PT37 Ajout variable StValidResp pour CP Pris }
  if TypeCp = 'PRI' then StValidResp := 'AND PCN_VALIDRESP="VAL" ' else StValidResp := '';
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" AND ' + { PT32 Ajout Clause }
    ' PCN_SALARIE = "' + Salarie + '" AND PCN_TYPECONGE = "' + Typecp + '"' + StValidResp +
    ' AND PCN_DATEFIN = "' + usdatetime(Datef) + '"' +
    ' AND PCN_DATEPAIEMENT >= "' + usdatetime(10) + '" ORDER BY PCN_DATEVALIDITE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then
  begin
}
  if TypeCp = 'PRI' then LcTob_pris.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
  if TypeCp = 'SLD' then
  begin
    LcTob_Solde.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
    if (Assigned(LcTob_Solde)) and (LcTob_Solde.detail.count > 0) then
      T := LcTob_solde.findfirst([''], [''], true)
    else T := nil;
    nb := 0;
    while T <> nil do
    begin
      if T.getvalue('PCN_CODERGRPT') <> -1 then
        nb := nb + T.getvalue('PCN_JOURS');
      T := LcTob_solde.findnext([''], [''], true);
    end;
    JCPSOLDE := nb;
  end;
{
  end;
  Ferme(Q);
}
  if Assigned(Tob_Pris) then { PT120 }
    Tob_Pris := LcTob_pris; //mv 13-11-00   { PT115-2 SB Je sais pas pourquoi, on conserve qd même!! }
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Procedure qui permet d'alimenter à chaque validation de
Suite ........ : bulletin la ligne de congés acquis correspondant
Mots clefs ... : PAIE;CP
*****************************************************************}

(*function AlimCongesAcquis(Tobsalarie, Tobetablissement: Tob; datedebut, datefin: tdatetime; numerateur, denominateur: integer; TOB_RUB: Tob; Ecrit: boolean; var base, Nbmois: double; var No_ordre: integer; var CpMois, CpSupp, CpAnc: double): double;
var
  Cum11, nbj: double;
  CPACQUISMOIS, CpAcquisAnc, Baseanc, Valanc, TDossier, DateCPAnc, TypVar, Profil: string;
  DateCompleteCPAnc, DTClot, DateSortie, DateDebutAca0106, DateFinAca0106: TdateTime;
  aaj, mmj, jjj, yy, mm, jj: word;
  T_ACQ, T_ACA, T_ACS: TOB;
  ChangePerAca0106: integer;
  MvtAcaCree: Boolean;
begin

  NbJ := 0; //PAYECOUR;  //PT-26 Inutile car on ne recharge plus les acquis en validation
  result := 0;
  base := 0;
  NbMois := 0;
  ChangePerAca0106 := 0;
  DateCompleteCPAnc := idate1900;
  DateSortie := TobSalarie.GetValue('PSA_DATESORTIE');
  Profil := PGRecupereProfilCge(TobSalarie.GetValue('PSA_ETABLISSEMENT')); //PT66

  if ((tobsalarie = nil) or (tobetablissement = nil)) then exit;
  if (T_MvtAcquis <> nil) and (Ecrit = True) then { PT-11-1 ajout cond. (Ecrit=True) }
  begin
    T_MvtAcquis.free;
    T_MvtACquis := nil;
  end;
  //PT-11-1 mise en commentaire  T_MvtAcquis:=TOB.Create('CONGES PAYE SALARIE mere',Nil,-1) ;


    //   AnnuleCongesAcquis(TobSalarie.getvalue('PSA_SALARIE'),datedebut,datefin);
  { si idem établissement, on va chercher la valeur dans la table etabcompl, }
  { parce qu'elle n'est pas forcément à jour }
  { DEB PT61-2 Prise en compte de l'alimentation variable paie si renseigné }
  CPACQUISMOIS := Tobsalarie.getvalue('PSA_CPACQUISMOIS');
  if CPACQUISMOIS = 'ETB' then
    JCPACQUISTHEORIQUE := TobEtablissement.getvalue('ETB_NBREACQUISCP')
  else
    JCPACQUISTHEORIQUE := Tobsalarie.getvalue('PSA_NBREACQUISCP');

  if (Tobsalarie.getvalue('PSA_TYPNBACQUISCP') = 'ETB') and (TobEtablissement.getvalue('ETB_NBACQUISCP') <> '') then
    CpMois := ValVariable(TobEtablissement.getvalue('ETB_NBACQUISCP'), Datedebut, Datefin, TOB_RUB)
  else
    if (Tobsalarie.getvalue('PSA_TYPNBACQUISCP') = 'PER') and (Tobsalarie.getvalue('PSA_NBACQUISCP') <> '') then
    CpMois := ValVariable(Tobsalarie.getvalue('PSA_NBACQUISCP'), Datedebut, Datefin, TOB_RUB)
  else
    CpMois := JCPACQUISTHEORIQUE;
  { FIN PT61-2 }

  DTClot := TobEtablissement.getvalue('ETB_DATECLOTURECPN');
  CpAcquisAnc := Tobsalarie.getvalue('PSA_CPACQUISANC');
  if CPACQUISAnc = 'ETB' then
  begin
    BaseAnc := TobEtablissement.getvalue('ETB_BASANCCP');
    ValAnc := TobEtablissement.getvalue('ETB_VALANCCP');
    { PT34-4 Mise en commentaire puis supprimé }
  end
  else
  begin
    BaseAnc := Tobsalarie.getvalue('PSA_BASANCCP');
    ValAnc := Tobsalarie.getvalue('PSA_VALANCCP');
    //   DateCPAnc       := Tobsalarie.getvalue('PSA_DATEACQCPANC');  PT34-4 Mise en commentaire
  end;
  //PT34-4 Gestion de la date d'acquisition de l'ancienneté idem etab ou perso salarié
  if Tobsalarie.getvalue('PSA_DATANC') = 'ETB' then
  begin
    if TobEtablissement.getvalue('ETB_TYPDATANC') <> '1' then
      DateCPAnc := TobEtablissement.getvalue('ETB_DATEACQCPANC')
    else
      DateCPAnc := Tobsalarie.getvalue('PSA_DATEENTREE');
  end
  else
  begin
    if Tobsalarie.getvalue('PSA_TYPDATANC') <> '1' then
      DateCPAnc := Tobsalarie.getvalue('PSA_DATEACQCPANC')
    else
      DateCPAnc := Tobsalarie.getvalue('PSA_DATEENTREE');
  end;

  { PT34-1 Lignes de code déplacés, imbriqués dans test date d'acquisition anciénneté }
  CpSupp := Tobsalarie.getvalue('PSA_NBRECPSUPP');

  Cum11 := RendCumulSalSess('11');
  if denominateur > 0 then
    NbMois := Numerateur / denominateur else NbMois := 0;
  CPMois := Arrondi(CpMois * NbMois, DCP); //PT-16
  CpSupp := CpSupp * NbMois;
  decodedate(datefin, aaj, mmj, jjj); //PT-18 on ne récupère pas la Date Système mais DateFin Bulletin
  if Length(DateCpAnc) = 4 then DateCpAnc := Copy(DateCpAnc, 1, 2) + '/' + Copy(DateCpAnc, 3, 2) + '/2000';
  if isValidDate(DateCpAnc) then
  begin
    DecodeDate(StrToDate(DateCpAnc), YY, MM, JJ); { PT-20 Utilisation de la fn de decodage }
    { PT-20 Mise en commentaire puis supprimer }
    if not isValidDate(IntToStr(jj) + '/' + IntToStr(mm) + '/' + IntToStr(aaj)) then
      jj := jj - 1; //PT34-2
    DateCompleteCPAnc := PGEncodeDateBissextile(aaj, mm, jj); { PT71 }
    if ((DateCompleteCPAnc >= datedebut) and (DateCompleteCPAnc <= datefin)) then
    begin { Calcul de l'ancienneté si date d'acquisition compris dans période de paie }
      { DEB PT34-3 Lignes de code replacés en fn de la validité de la date d'acquisition de l'ancienneté }
      if ((BaseAnc = 'VAR') and (Valanc <> '')) then
        CpAnc := ValVariable(ValAnc, Datedebut, Datefin, TOB_RUB)
      else
        if (BaseAnc = 'VAL') and (Isnumerique(Valanc)) then
        CpAnc := Valeur(ValAnc)
      else
        if BaseAnc = 'TAB' then
      begin
        TypVar := 'ANC';
        TDossier := Valanc; { PT61-3 Recupération paramétrage }
        CpAnc := ValVarTable(TypVar, TDossier, DateDebut, DateFin, TOB_Rub, 'CP'); // Ancienneté
      end;
    end
      //FIN PT34-3
    else
      CpAnc := 0;
  end
  else
    CpAnc := 0;


  { 3 alimentation des lignes de congés }
  { ------ Mvt ACQ ------- }
  if Ecrit then
  begin
    T_MvtAcquis := TOB.Create('CONGES PAYE SALARIE mere', nil, -1);
    if CpMois > 0 then
    begin
      T_ACQ := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
      InitialiseTobAbsenceSalarie(T_ACQ);
      T_ACQ.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
      T_ACQ.PutValue('PCN_ORDRE', No_Ordre);
      T_ACQ.PutValue('PCN_CODERGRPT', No_Ordre);
      T_ACQ.PutValue('PCN_DATEDEBUT', datedebut);
      T_ACQ.PutValue('PCN_DATEFIN', datefin);
      T_ACQ.PutValue('PCN_TYPECONGE', 'ACQ');
      T_ACQ.PutValue('PCN_SENSABS', '+');
      T_ACQ.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(datefin));
      T_ACQ.PutValue('PCN_DATEMODIF', Date);
      T_ACQ.PutValue('PCN_DATEVALIDITE', Datefin);
      T_ACQ.PutValue('PCN_PERIODECP', CalculPeriode(DTClot, Datefin));
      T_ACQ.PutValue('PCN_JOURS', Arrondi(CpMois, DCP));
      T_ACQ.PutValue('PCN_BASE', Arrondi(Cum11, DCP));
      T_ACQ.PutValue('PCN_NBREMOIS', Arrondi(NbMois, DCP));
      T_ACQ.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
      T_ACQ.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
      T_ACQ.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
      T_ACQ.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
      T_ACQ.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
      T_ACQ.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
      T_ACQ.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));
      if Nbj > 0 then
        if Nbj < Arrondi(CpMois, DCP) then
        begin
          T_ACQ.putValue('PCN_APAYES', Arrondi(CpMois - Nbj, DCP));
          Nbj := 0;
        end
        else
        begin
          T_ACQ.putValue('PCN_APAYES', Arrondi(CpMois, DCP));
          Nbj := Nbj - Arrondi(CpMois, DCP);
        end;
      if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then //PT66
        T_ACQ.PutValue('PCN_CODETAPE', 'S');
      Base := Arrondi(Cum11, DCP);
    end;

    { ------ Mvt ACS ------- }
    if CpSupp > 0 then
    begin
      T_ACS := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
      No_ordre := No_ordre + 1;
      InitialiseTobAbsenceSalarie(T_ACS);
      T_ACS.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
      T_ACS.PutValue('PCN_ORDRE', No_Ordre);
      T_ACS.PutValue('PCN_CODERGRPT', No_Ordre);
      T_ACS.PutValue('PCN_DATEDEBUT', datedebut);
      T_ACS.PutValue('PCN_DATEFIN', datefin);
      T_ACS.PutValue('PCN_TYPECONGE', 'ACS');
      T_ACS.PutValue('PCN_SENSABS', '+');
      T_ACS.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(datefin));
      T_ACS.PutValue('PCN_DATEMODIF', Date);
      T_ACS.PutValue('PCN_DATEVALIDITE', Datefin);
      T_ACS.PutValue('PCN_PERIODECP', CalculPeriode(DTClot, Datefin));
      T_ACS.PutValue('PCN_JOURS', Arrondi(CpSupp, DCP));
      T_ACS.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
      T_ACS.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
      T_ACS.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
      T_ACS.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
      T_ACS.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
      T_ACS.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
      T_ACS.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));

      if Nbj > 0 then
        if Nbj < Arrondi(CpSupp, DCP) then
        begin
          T_ACS.putValue('PCN_APAYES', Arrondi(CpSupp - Nbj, DCP));
          Nbj := 0;
        end
        else
        begin
          T_ACS.putValue('PCN_APAYES', Arrondi(CpSupp, DCP));
          Nbj := Nbj - Arrondi(CpSupp, DCP);
        end;
      if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then //PT66
        T_ACS.PutValue('PCN_CODETAPE', 'S');
    end;
    { ------ Mvt ACA ------- }
    MvtAcaCree := False;
    if DateCpAnc = '0106' then ChangePerAca0106 := 1; //modif sb 07/06/01
    DateDebutAca0106 := DateDebut;
    DateFinAca0106 := DateFin;
    if ChangePerAca0106 <> 0 then
    begin
      DateDebutAca0106 := PlusMois(datedebut, ChangePerAca0106 * (-1));
      DateFinAca0106 := datedebut - 1;
      MvtAcaCree := ExisteSql('SELECT * FROM ABSENCESALARIE ' +
        'WHERE PCN_SALARIE="' + Tobsalarie.getvalue('PSA_SALARIE') + '" ' +
        'AND PCN_TYPEMVT="CPA" ' + //PT32 AJout AND PCN_TYPEMVT="CPA"
        'AND PCN_TYPECONGE="ACA" AND PCN_DATEVALIDITE="' + UsDateTime(DateFinAca0106) + '"');
    end; //modif sb 07/06/01
    decodedate(Date, aaj, mmj, jjj);
    if DateCpAnc <> '' then
    begin
      if ((DateCompleteCPAnc >= datedebut) and (DateCompleteCPAnc <= datefin)
        and (CpAnc > 0)) and (MvtAcaCree = False) then
      begin
        T_ACA := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
        No_ordre := No_ordre + 1;
        InitialiseTobAbsenceSalarie(T_ACA);
        T_ACA.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
        T_ACA.PutValue('PCN_ORDRE', No_Ordre);
        T_ACA.PutValue('PCN_CODERGRPT', No_Ordre);
        T_ACA.PutValue('PCN_DATEDEBUT', DateDebutAca0106);
        T_ACA.PutValue('PCN_DATEFIN', DateFinAca0106);
        T_ACA.PutValue('PCN_TYPECONGE', 'ACA');
        T_ACA.PutValue('PCN_SENSABS', '+');
        T_ACA.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(DateFinAca0106));
        T_ACA.PutValue('PCN_DATEMODIF', Date);
        T_ACA.PutValue('PCN_DATEVALIDITE', DateFinAca0106);
        T_ACA.PutValue('PCN_PERIODECP', (CalculPeriode(DTClot, DateFinAca0106))); //modif SB    07/06/01
        T_ACA.PutValue('PCN_JOURS', Arrondi(CpAnc, DCP));
        T_ACA.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
        T_ACA.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
        T_ACA.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
        T_ACA.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
        T_ACA.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
        T_ACA.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
        T_ACA.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));

        if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then //PT66
          T_ACA.PutValue('PCN_CODETAPE', 'S');
        if Nbj > 0 then
          if Nbj < Arrondi(CpAnc, DCP) then
          begin
            T_ACA.putValue('PCN_APAYES', Arrondi(CpAnc - Nbj, DCP));
            Nbj := 0;
          end
          else
          begin
            T_ACA.putValue('PCN_APAYES', Arrondi(CpAnc, DCP));
            Nbj := Nbj - Arrondi(CpMois, DCP);
          end;
      end;
    end;
  end;
  if Ecrit then
  begin
    if ((DateSortie >= DateDebut) and (DateSortie <= DateFin)) and (Profil <> '') then //PT66
      SoldeAcquis(T_MvtAcquis, False); //PT-11-1 Ajout False
    if Tob_AcquisAVirer = nil then
      Tob_AcquisAVirer := TOB.Create('CP A VIRER', nil, -1);
    if Tob_AcquisAVirer <> nil then
      ChangeCongesAcquisdetob(TobSalarie.getvalue('PSA_SALARIE'), datedebut, datefin, tob_acquis);
    suivant := No_ordre + 1;
  end;
  result := CpMois + CPSupp + CpAnc;
end;   *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 19/03/2004
Modifié le ... :   /  /
Description .. : Chargement des acquis en fonction :
Suite ........ : => Paramétrage CP
Suite ........ : => Saisie d'un trentième
Suite ........ : => Saisie manuelle sur bulletin
Mots clefs ... : PAIE;CONGESPAYES
*****************************************************************}
{ DEB PT73 }

// PT128 function ChargeAcquisParametre(Tobsalarie, Tobetablissement, TOB_RUB: Tob; datedebut, datefin: tdatetime; OnLoad, OnRefresh: Boolean; var base, Nbmois, CpMois, CpSupp, CpAnc: double): double;
function ChargeAcquisParametre(Tobsalarie, Tobetablissement, TOB_RUB: Tob; datedebut, datefin: tdatetime; OnLoad, OnRefresh: Boolean; var base, Nbmois, CpMois, CpSupp, CpAnc, CpFract: double): double;
var
  BaseAnc, ValAnc, DateCPAnc: string;
  DateCompleteCPAnc: TDateTime;
  Q: TQuery;
begin
  result := 0;
  base := 0;
  NbMois := 0;

  if ((tobsalarie = nil) or (tobetablissement = nil)) then exit;

  { Alimentation des acquis non forcé, recherche au niveau paramétrage }
  if not ChpEntete.CpAcquisMod then
  begin
    //   AnnuleCongesAcquis(TobSalarie.getvalue('PSA_SALARIE'),datedebut,datefin);
    { si idem établissement, on va chercher la valeur dans la table etabcompl, }
    { parce qu'elle n'est pas forcément à jour }
    { DEB PT61-2 Prise en compte de l'alimentation variable paie si renseigné }
    if Tobsalarie.getvalue('PSA_CPACQUISMOIS') = 'ETB' then
      JCPACQUISTHEORIQUE := TobEtablissement.getvalue('ETB_NBREACQUISCP')
    else
      JCPACQUISTHEORIQUE := Tobsalarie.getvalue('PSA_NBREACQUISCP');

    if (Tobsalarie.getvalue('PSA_TYPNBACQUISCP') = 'ETB') and (TobEtablissement.getvalue('ETB_NBACQUISCP') <> '') then
      CpMois := ValVariable(TobEtablissement.getvalue('ETB_NBACQUISCP'), Datedebut, Datefin, TOB_RUB)
    else
      if (Tobsalarie.getvalue('PSA_TYPNBACQUISCP') = 'PER') and (Tobsalarie.getvalue('PSA_NBACQUISCP') <> '') then
        CpMois := ValVariable(Tobsalarie.getvalue('PSA_NBACQUISCP'), Datedebut, Datefin, TOB_RUB)
      else
        CpMois := JCPACQUISTHEORIQUE;
    { FIN PT61-2 }

    if Tobsalarie.getvalue('PSA_CPACQUISANC') = 'ETB' then
    begin
      BaseAnc := TobEtablissement.getvalue('ETB_BASANCCP');
      ValAnc := TobEtablissement.getvalue('ETB_VALANCCP');
      { PT34-4 Mise en commentaire puis supprimé }
    end
    else
    begin
      BaseAnc := Tobsalarie.getvalue('PSA_BASANCCP');
      ValAnc := Tobsalarie.getvalue('PSA_VALANCCP');
    end;
    // Gestion de la date d'acquisition de l'ancienneté idem etab ou perso salarié
    if (Tobsalarie.getvalue('PSA_DATANC') = 'ETB') and (TobEtablissement.getvalue('ETB_TYPDATANC') <> '') then { PT74-4 }
    begin
      if (TobEtablissement.getvalue('ETB_TYPDATANC') = '0') or (TobEtablissement.getvalue('ETB_TYPDATANC') = '2') or (TobEtablissement.getvalue('ETB_TYPDATANC') = '4') then { PT121 }
        DateCPAnc := TobEtablissement.getvalue('ETB_DATEACQCPANC')
      else
        if (TobEtablissement.getvalue('ETB_TYPDATANC') = '1') then { PT121 }
          DateCPAnc := Tobsalarie.getvalue('PSA_DATEENTREE')
        else
          if (TobEtablissement.getvalue('ETB_TYPDATANC') = '3') then { PT121 }
            DateCPAnc := Tobsalarie.getvalue('PSA_DATEANCIENNETE');
    end
    else
      if (Tobsalarie.getvalue('PSA_DATANC') = 'PER') and (Tobsalarie.getvalue('PSA_TYPDATANC') <> '') then { PT74-4 }
      begin
        if (Tobsalarie.getvalue('PSA_TYPDATANC') = '0') or (Tobsalarie.getvalue('PSA_TYPDATANC') = '2') or (Tobsalarie.getvalue('PSA_TYPDATANC') = '4') then { PT121 }
          DateCPAnc := Tobsalarie.getvalue('PSA_DATEACQCPANC')
        else
          if (Tobsalarie.getvalue('PSA_TYPDATANC') = '1') then { PT121 }
            DateCPAnc := Tobsalarie.getvalue('PSA_DATEENTREE')
          else
            if (Tobsalarie.getvalue('PSA_TYPDATANC') = '3') then { PT121 }
              DateCPAnc := Tobsalarie.getvalue('PSA_DATEANCIENNETE'); { PT121 }
      end
      else DateCPAnc := '';
    { DEB PT74-2 }
    if Tobsalarie.getvalue('PSA_CPACQUISSUPP') = 'ETB' then
      CpSupp := TobEtablissement.getvalue('ETB_NBRECPSUPP')
    else
      CpSupp := Tobsalarie.getvalue('PSA_NBRECPSUPP');
    { FIN PT74-2 }
    if ChpEntete.DTrent > 0 then NbMois := ChpEntete.NTrent / ChpEntete.DTrent else NbMois := 0;
    CPMois := Arrondi(CpMois * NbMois, DCP); //PT-16
    CpSupp := CpSupp * NbMois;

    DateCompleteCPAnc := PgCalculDateAnciennete(datefin, DateCpAnc); { PT122 }


    if ((DateCompleteCPAnc >= datedebut) and (DateCompleteCPAnc <= datefin)) then
    begin { Calcul de l'ancienneté si date d'acquisition compris dans période de paie }
      if ((BaseAnc = 'VAR') and (Valanc <> '')) then
        CpAnc := ValVariable(ValAnc, Datedebut, Datefin, TOB_RUB)
      else
        if (BaseAnc = 'VAL') and (Isnumerique(Valanc)) then
          CpAnc := Valeur(ValAnc)
        else
          if BaseAnc = 'TAB' then
            CpAnc := ValVarTable('ANC', Valanc, DateDebut, DateFin, TOB_Rub, 'CP'); // Ancienneté
    end
    else
      CpAnc := 0;
// d PT128 PT134
      if (NOT CalcFractOk) then
      begin
      CpFract := 0;
      Q := OpenSql('SELECT PCN_JOURS FROM ABSENCESALARIE ' +
        'WHERE PCN_TYPEMVT="CPA" AND PCN_DATEVALIDITE<="' + UsDateTime(DateFin) + '" ' +
        'AND PCN_DATEDEBUT="' + UsDateTime(DebutdeMois(PlusDate(DateDebut,-1,'M'))) + '" '+
        'AND PCN_DATEFIN="' + UsDateTime(FindeMois(PlusDate(DateDebut,-1,'M'))) + '" ' +
        'AND PCN_SALARIE="' + Tobsalarie.GetValue('PSA_SALARIE') + '" ' +
        'AND PCN_TYPECONGE="ACF" ', True);
      while not Q.eof do
      begin
        CpFract:= CpFract + Q.FindField('PCN_JOURS').AsFloat;
        Q.next;
      end;
      Ferme(Q);
      CalcFractOk := TRUE;
      end;
// f PT128 PT134
    if OnRefresh then { PT73 13/05/2004 }
    begin
      ChpEntete.CpAcq := CpMois;
      ChpEntete.CpSupl := CpSupp;
      ChpEntete.CpAnc := CpAnc;
    end;
  end
  else
  begin
    { Alimentation des acquis forcé, récupération des zones de la fiche }
    if ChpEntete.DTrent > 0 then NbMois := ChpEntete.NTrent / ChpEntete.DTrent else NbMois := 0;
    CpMois := ChpEntete.CpAcq;
    CpSupp := ChpEntete.CpSupl;
    CpAnc := ChpEntete.CpAnc;
    if OnLoad then
    begin
{PT128      Q := OpenSql('SELECT PCN_TYPECONGE,PCN_JOURS FROM ABSENCESALARIE ' +
        'WHERE PCN_TYPEMVT="CPA" AND PCN_DATEVALIDITE="' + UsDateTime(DateFin) + '" ' +
        'AND PCN_DATEDEBUT="' + UsDateTime(DateDebut) + '" AND PCN_DATEFIN="' + UsDateTime(DateFin) + '" ' +
        'AND PCN_SALARIE="' + Tobsalarie.GetValue('PSA_SALARIE') + '" ' +
        'AND ( PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="ACA" )', True);
*}
      Q := OpenSql('SELECT PCN_TYPECONGE,PCN_JOURS FROM ABSENCESALARIE ' +
        'WHERE PCN_TYPEMVT="CPA" '+
        'AND PCN_SALARIE="' + Tobsalarie.GetValue('PSA_SALARIE') + '" ' +
        'AND ((PCN_DATEVALIDITE="' + UsDateTime(DateFin) + '" ' +
        'AND PCN_DATEDEBUT="' + UsDateTime(DateDebut) + '" AND PCN_DATEFIN="' + UsDateTime(DateFin) + '" ' +
        'AND ( PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="ACA" )) '+
        'OR (PCN_DATEVALIDITE<="' +UsDateTime(DateFin) + '" ' +
        'AND PCN_DATEDEBUT="' + UsDateTime(DebutdeMois(PlusDate(DateDebut,-1,'M'))) + '" ' +
        'AND PCN_DATEFIN="' + UsDateTime(FindeMois(PlusDate(DateDebut,-1,'M'))) + '" ' +
        'AND PCN_SALARIE="' + Tobsalarie.GetValue('PSA_SALARIE') + '" ' +
        'AND PCN_TYPECONGE="ACF" ))', True);
      while not Q.eof do
      begin
        if Q.FindField('PCN_TYPECONGE').AsString = 'ACQ' then CpMois := Q.FindField('PCN_JOURS').AsFloat;
        if Q.FindField('PCN_TYPECONGE').AsString = 'ACS' then CpSupp := Q.FindField('PCN_JOURS').AsFloat;
        if Q.FindField('PCN_TYPECONGE').AsString = 'ACA' then CpAnc := Q.FindField('PCN_JOURS').AsFloat;
        if Q.FindField('PCN_TYPECONGE').AsString = 'ACF' then CpFract := Q.FindField('PCN_JOURS').AsFloat;
        Q.next;
      end;
      Ferme(Q);
      ChpEntete.CpAcq := CpMois; { PT73 03/05/2004 }
      ChpEntete.CpSupl := CpSupp; { PT73 03/05/2004 }
      ChpEntete.CpAnc := CpAnc; { PT73 03/05/2004 }
    end;

  end;
  Base := Arrondi(RendCumulSalSess('11'), DCP);
  JCPACQ := CpMois; { PT103 Distinction pour calcul ICCP }
  JCPACS := CPSupp; { PT103 Distinction pour calcul ICCP }
  JCPACA := CpAnc; { PT103 Distinction pour calcul ICCP }
//PT128  result := CpMois + CPSupp + CpAnc;
  result := CpMois + CPSupp + CpAnc + CpFract;
end;
{ FIN PT73 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 19/03/2004
Modifié le ... :   /  /
Description .. : Génération de la tob des acquis CP
Mots clefs ... : PAIE;CONGESPAYES
*****************************************************************}
{ DEB PT73 }

procedure InitTobAcquisCp(Tobsalarie, Tobetablissement: Tob; datedebut, datefin: tdatetime; base, Nbmois, CpMois, CpSupp, CpAnc: double; var No_ordre: integer);
var
  T_ACQ, T_ACA, T_ACS: TOB;
  Profil: string;
  DTClot, DateSortie: TdateTime;
  Cum11, NbJ: Double;
begin
  NbJ := 0;
  if ((tobsalarie = nil) or (tobetablissement = nil)) then exit;

  if Assigned(T_MvtAcquis) then FreeAndNil(T_MvtAcquis);

  DateSortie := TobSalarie.GetValue('PSA_DATESORTIE');
  Profil := PGRecupereProfilCge(TobSalarie.GetValue('PSA_ETABLISSEMENT'));
  DTClot := TobEtablissement.getvalue('ETB_DATECLOTURECPN');
  Cum11 := RendCumulSalSess('11');
  { 3 alimentation des lignes de congés }
  { ------ Mvt ACQ ------- }
  T_MvtAcquis := TOB.Create('CONGES PAYE SALARIE mere', nil, -1);
//  if CpMois > 0 then  PT102 Mise en commentaire, généré systématiquement
  begin
    T_ACQ := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
    InitialiseTobAbsenceSalarie(T_ACQ);
    T_ACQ.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
    T_ACQ.PutValue('PCN_ORDRE', No_Ordre);
    T_ACQ.PutValue('PCN_CODERGRPT', No_Ordre);
    T_ACQ.PutValue('PCN_DATEDEBUT', datedebut);
    T_ACQ.PutValue('PCN_DATEFIN', datefin);
    T_ACQ.PutValue('PCN_TYPECONGE', 'ACQ');
    T_ACQ.PutValue('PCN_SENSABS', '+');
    T_ACQ.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(datefin));
    T_ACQ.PutValue('PCN_DATEMODIF', Date);
    T_ACQ.PutValue('PCN_DATEVALIDITE', Datefin);
    T_ACQ.PutValue('PCN_PERIODECP', CalculPeriode(DTClot, Datefin));
    T_ACQ.PutValue('PCN_JOURS', Arrondi(CpMois, DCP));
    T_ACQ.PutValue('PCN_BASE', Arrondi(Cum11, DCP));
    T_ACQ.PutValue('PCN_NBREMOIS', Arrondi(NbMois, DCP));
    T_ACQ.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
    T_ACQ.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
    T_ACQ.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
    T_ACQ.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
    T_ACQ.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
    T_ACQ.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
    T_ACQ.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));
    if Nbj > 0 then
      if Nbj < Arrondi(CpMois, DCP) then
      begin
        T_ACQ.putValue('PCN_APAYES', Arrondi(CpMois - Nbj, DCP));
        Nbj := 0;
      end
      else
      begin
        T_ACQ.putValue('PCN_APAYES', Arrondi(CpMois, DCP));
        Nbj := Nbj - Arrondi(CpMois, DCP);
      end;
    if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then
      T_ACQ.PutValue('PCN_CODETAPE', 'S');
    Base := Arrondi(Cum11, DCP);
  end;

  { ------ Mvt ACS ------- }
  if CpSupp > 0 then
  begin
    T_ACS := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
    No_ordre := No_ordre + 1;
    InitialiseTobAbsenceSalarie(T_ACS);
    T_ACS.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
    T_ACS.PutValue('PCN_ORDRE', No_Ordre);
    T_ACS.PutValue('PCN_CODERGRPT', No_Ordre);
    T_ACS.PutValue('PCN_DATEDEBUT', datedebut);
    T_ACS.PutValue('PCN_DATEFIN', datefin);
    T_ACS.PutValue('PCN_TYPECONGE', 'ACS');
    T_ACS.PutValue('PCN_SENSABS', '+');
    T_ACS.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(datefin));
    T_ACS.PutValue('PCN_DATEMODIF', Date);
    T_ACS.PutValue('PCN_DATEVALIDITE', Datefin);
    T_ACS.PutValue('PCN_PERIODECP', CalculPeriode(DTClot, Datefin));
    T_ACS.PutValue('PCN_JOURS', Arrondi(CpSupp, DCP));
    T_ACS.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
    T_ACS.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
    T_ACS.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
    T_ACS.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
    T_ACS.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
    T_ACS.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
    T_ACS.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));

    if Nbj > 0 then
      if Nbj < Arrondi(CpSupp, DCP) then
      begin
        T_ACS.putValue('PCN_APAYES', Arrondi(CpSupp - Nbj, DCP));
        Nbj := 0;
      end
      else
      begin
        T_ACS.putValue('PCN_APAYES', Arrondi(CpSupp, DCP));
        Nbj := Nbj - Arrondi(CpSupp, DCP);
      end;
    if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then
      T_ACS.PutValue('PCN_CODETAPE', 'S');
  end;
  { ------ Mvt ACA ------- }
  if CpAnc > 0 then
  begin
    T_ACA := TOB.Create('ABSENCESALARIE', T_MvtAcquis, -1);
    No_ordre := No_ordre + 1;
    InitialiseTobAbsenceSalarie(T_ACA);
    T_ACA.PutValue('PCN_SALARIE', Tobsalarie.getvalue('PSA_SALARIE'));
    T_ACA.PutValue('PCN_ORDRE', No_Ordre);
    T_ACA.PutValue('PCN_CODERGRPT', No_Ordre);
    T_ACA.PutValue('PCN_DATEDEBUT', DateDebut);
    T_ACA.PutValue('PCN_DATEFIN', DateFin);
    T_ACA.PutValue('PCN_TYPECONGE', 'ACA');
    T_ACA.PutValue('PCN_SENSABS', '+');
    T_ACA.PutValue('PCN_LIBELLE', 'Acquis au ' + datetostr(DateFin));
    T_ACA.PutValue('PCN_DATEMODIF', Date);
    T_ACA.PutValue('PCN_DATEVALIDITE', DateFin);
    T_ACA.PutValue('PCN_PERIODECP', (CalculPeriode(DTClot, DateFin)));
    T_ACA.PutValue('PCN_JOURS', Arrondi(CpAnc, DCP));
    T_ACA.PutValue('PCN_ETABLISSEMENT', TobEtablissement.getvalue('ETB_ETABLISSEMENT'));
    T_ACA.Putvalue('PCN_TRAVAILN1', Tobsalarie.getvalue('PSA_TRAVAILN1'));
    T_ACA.Putvalue('PCN_TRAVAILN2', Tobsalarie.getvalue('PSA_TRAVAILN2'));
    T_ACA.Putvalue('PCN_TRAVAILN3', Tobsalarie.getvalue('PSA_TRAVAILN3'));
    T_ACA.Putvalue('PCN_TRAVAILN4', Tobsalarie.getvalue('PSA_TRAVAILN4'));
    T_ACA.Putvalue('PCN_CODESTAT', Tobsalarie.getvalue('PSA_CODESTAT'));
    T_ACA.Putvalue('PCN_CONFIDENTIEL', Tobsalarie.getvalue('PSA_CONFIDENTIEL'));

    if ((DateSortie >= DateDebut) and (Datesortie <= Datefin)) and (Profil <> '') then
      T_ACA.PutValue('PCN_CODETAPE', 'S');
    if Nbj > 0 then
      if Nbj < Arrondi(CpAnc, DCP) then
      begin
        T_ACA.putValue('PCN_APAYES', Arrondi(CpAnc - Nbj, DCP));
      end
      else
      begin
        T_ACA.putValue('PCN_APAYES', Arrondi(CpAnc, DCP));
      end;
  end;

  if ((DateSortie >= DateDebut) and (DateSortie <= DateFin)) and (Profil <> '') then
    SoldeAcquis(T_MvtAcquis, False, DateFin); { PT106 }
  if Tob_AcquisAVirer = nil then
    Tob_AcquisAVirer := TOB.Create('CP A VIRER', nil, -1);
  if Tob_AcquisAVirer <> nil then
    ChangeCongesAcquisdetob(TobSalarie.getvalue('PSA_SALARIE'), datedebut, datefin, tob_acquis);
  suivant := No_ordre + 1;
end;
{ FIN PT73 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : - charge la tob des congés acquis sur session en cours
Suite ........ : - Renvoie True si mvt d'acq bulletin soldé ou cloturé afin de
Suite ........ : bloquer la validation CP..
Mots clefs ... : PAIE;CP
*****************************************************************}
{PT-27-1 Modif de la procédure en function
Renvoie True si mvt d'acq bulletin soldé ou cloturé afin de bloquer la validation CP..}

// PT128 function ChargeT_MvtAcquis(Salarie: string; DD, DF: TDateTime; Mois, Supp, Anc: Double; var Modif, New: Boolean): boolean;
function ChargeT_MvtAcquis(Salarie: string; DD, DF: TDateTime; Mois, Supp, Anc, Fract: Double; var Modif, New: Boolean): boolean; 
var
  T: Tob;
  Nb, NbPlus: integer;
  st: string;
//PT128  DejaCreeACQ, DejaCreeACS, DejaCreeACA: Boolean;
  DejaCreeACQ, DejaCreeACS, DejaCreeACA, DejaCreeACF: Boolean;
  Cum11: Double;
begin
  { PT73 Refonte fonctionnement chargement acquis en cours..en fonction de la saisie sur bulletin }
  result := False; { PT-27-1 }
  if T_MvtAcquis <> nil then
  begin
    T_MvtAcquis.free;
    T_MvtAcquis := nil
  end;
  New := False;
  Nb := 0;
  NbPlus := 0;
  DejaCreeACQ := False;
  DejaCreeACA := False;
  DejaCreeACS := False;
  DejaCreeACF := False; // PT128
  { alimentation de la tob des acquis }
{*
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + Salarie + '" ' +
    'AND PCN_TYPEMVT="CPA" ' + //PT32 Ajout PCN_TYPEMVT="CPA"
  //'AND PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S" '+     PT-27-1 mise en commentaire
  'AND (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS") ' +
    'AND PCN_DATEVALIDITE = "' + usdatetime(DF) + '" ' +
    'ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';
*}
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + Salarie + '" ' +
    'AND PCN_TYPEMVT="CPA" ' +
    'AND (((PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS") ' +
    'AND PCN_DATEVALIDITE = "' + usdatetime(DF) + '" ) ' +
    'OR (PCN_TYPECONGE="ACF" AND PCN_DATEVALIDITE <= "' + usdatetime(DF) + '" )) ' +
    'ORDER BY PCN_DATEVALIDITE, PCN_TYPECONGE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then // pas de congés acquis
  begin
}
  T_MvtAcquis := TOB.Create('CONGES PAYE SALARIE mere', nil, -1);
  T_MvtAcquis.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
  Modif := false;
  Nb := T_MvtAcquis.detail.count;
  if T_MvtAcquis.detail.count <= 0 then FreeAndNIL(T_MvtAcquis);
//  end;
//  Ferme(Q);

   { DEB PT102  Recherche cumul11 > 0 pour générer mvt ACQ = 0 jours }
  Cum11 := RendCumulSalSess('11');
  if Cum11 = 0 then Cum11 := Valcumuldate('11', DD, DF);
   { FIN PT102 }
  { Top indicateur mvt acquis crée en cours de saisie de bulletin }
  if (Mois > 0) or ((Cum11 > 0) and (Mois = 0)) then NbPlus := 1; { PT102 }
  if Supp > 0 then NbPlus := NbPlus + 1;
  if Anc > 0 then NbPlus := NbPlus + 1;
  if Fract > 0 then NbPlus := NbPlus + 1;   // PT128
  if (NbPlus > 0) and (NbPlus > Nb) then New := True;

  { Réimputation du Jour acquis modifié sur jour acquis TOB }
  { Si saisie 0 alors mise en tob de suppression du mouvement }
  { Si mouvement déjà mise en tob suppression alors free de la tob de chargement }
  if (T_MvtAcquis <> nil) then
  begin
    if T_MvtAcquisAVirer = nil then
      T_MvtAcquisAVirer := TOB.Create('CP A VIRER', nil, -1);
    T := T_MvtAcquis.FindFirst([''], [''], False);
    while T <> nil do
    begin
      { Si mvt C ou S, on ne recharge pas les CP pour ne pas fausser le décompte }
      if T.GetValue('PCN_CODETAPE') <> '...' then result := True; { PT-27-1 }
      if (T.GetValue('PCN_TYPECONGE') = 'ACQ') then
      begin
        DejaCreeACQ := True;
        if T.GetValue('PCN_JOURS') <> Mois then
        begin
          Modif := True;
          if (Mois = 0) and (Cum11 = 0) then { PT102 }
          begin
            if Assigned(T_MvtAcquisAVirer.FindFirst(['PCN_ORDRE'], [T.GetValue('PCN_ORDRE')], False)) then
              T.free
            else
              T.ChangeParent(T_MvtAcquisAVirer, -1);
            { T.free; PT73 03/05/2004 Mise en commentaire }
          end
          else
            T.PutValue('PCN_JOURS', Mois);
        end;
      end
      else
        if (T.GetValue('PCN_TYPECONGE') = 'ACS') then
        begin
          DejaCreeACS := True;
          if T.GetValue('PCN_JOURS') <> Supp then
          begin
            Modif := True;
            if (Supp = 0) and (T_MvtAcquisAVirer <> nil) then
            begin
              if Assigned(T_MvtAcquisAVirer.FindFirst(['PCN_ORDRE'], [T.GetValue('PCN_ORDRE')], False)) then
                T.free
              else
                T.ChangeParent(T_MvtAcquisAVirer, -1)
                { T.free; PT73 03/05/2004 Mise en commentaire }
            end
            else
              T.PutValue('PCN_JOURS', Supp);
          end;
        end
        else
          if (T.GetValue('PCN_TYPECONGE') = 'ACA') then
          begin
            DejaCreeACA := True;
            if T.GetValue('PCN_JOURS') <> Anc then
            begin
              Modif := True;
              if (Anc = 0) and (T_MvtAcquisAVirer <> nil) then
              begin
                if Assigned(T_MvtAcquisAVirer.FindFirst(['PCN_ORDRE'], [T.GetValue('PCN_ORDRE')], False)) then
                  T.free
                else
                  T.ChangeParent(T_MvtAcquisAVirer, -1);
            { T.free; PT73 03/05/2004 Mise en commentaire }
              end
              else
                T.PutValue('PCN_JOURS', Anc);
            end;
          end
          else
// d PT128
          if (T.GetValue('PCN_TYPECONGE') = 'ACF') then
          begin
            DejaCreeACS := True;
            if T.GetValue('PCN_JOURS') <> Fract then
            begin
              Modif := True;
              if (Fract = 0) and (T_MvtAcquisAVirer <> nil) then
              begin
                if Assigned(T_MvtAcquisAVirer.FindFirst(['PCN_ORDRE'], [T.GetValue('PCN_ORDRE')], False)) then
                  T.free
                else
                  T.ChangeParent(T_MvtAcquisAVirer, -1)
                  { T.free; PT73 03/05/2004 Mise en commentaire }
                end
                else
                  T.PutValue('PCN_JOURS', Fract);
               end;
// f PT128
            end;
      T := T_MvtAcquis.FindNext([''], [''], False);
    end;
    { Top indicateur mvt acquis crée en cours de saisie de bulletin }
    if ((Mois > 0) or ((Cum11 > 0) and (Mois = 0))) and (DejaCreeACQ = False) then New := True; { PT102 }
    if (Supp > 0) and (DejaCreeACS = False) then New := True;
    if (Anc > 0) and (DejaCreeACA = False) then New := True;
    if (Fract > 0) and (DejaCreeACF = False) then New := True;   // PT128
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Affecte la base congés payés du mouvement acquis
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AffectBaseAcquisEncours;
var
  T: Tob;
  Cum11: Double;
begin
  { PT-19 Mise en commentaire puis supprimer }
  if T_MvtAcquis <> nil then
  begin
    T := T_MvtAcquis.FindFirst(['PCN_TYPECONGE'], ['ACQ'], False);
    while T <> nil do
    begin
      { DEB PT102 Suppression du mvt ACQ si jour=0 & cumul11=0 }
      Cum11 := Arrondi(RendCumulSalSess('11'), DCP);
      if (T.GetValue('PCN_JOURS') = 0) and (Cum11 = 0) then T.free
      else
        T.PutValue('PCN_BASE', Cum11);
      { FIN PT102 }
      T := T_MvtAcquis.FindNext(['PCN_TYPECONGE'], ['ACQ'], False);
    end;
  end;
end;

procedure AjouteAcquiscourant(Tob_Acquis, Tob_sal: tob; DateF: tdatetime; var No_Ordre: integer);
var
  T_ACC: tob;
begin
  T_ACC := TOB.Create('ABSENCESALARIE', Tob_acquis, -1);
  No_ordre := No_ordre + 1;
  InitialiseTobAbsenceSalarie(T_ACC);
  T_ACC.PutValue('PCN_SALARIE', Tob_sal.getvalue('PSA_SALARIE'));
  T_ACC.PutValue('PCN_ORDRE', No_Ordre);
  T_ACC.PutValue('PCN_TYPECONGE', 'ACC');
  T_ACC.PutValue('PCN_SENSABS', '+');
  T_ACC.PutValue('PCN_LIBELLE', 'Acquis courant');
  T_ACC.PutValue('PCN_DATEVALIDITE', Datef);
  T_ACC.PutValue('PCN_JOURS', Arrondi(JCPACQUIS, DCP));
  T_ACC.PutValue('PCN_BASE', Arrondi(BCPACQUIS, DCP));
  T_ACC.PutValue('PCN_NBREMOIS', Arrondi(MCPACQUIS, DCP));
  T_ACC.PutValue('PCN_PERIODECP', 0);
  T_ACC.PutValue('PCN_ETABLISSEMENT', Tob_sal.getvalue('PSA_ETABLISSEMENT'));
end;

procedure RetireAcquiscourant(Tob_Acquis: tob);
var
  T: tob;
begin
  if Assigned(Tob_acquis) and (Tob_acquis.detail.count > 0) then // notvide(Tob_acquis, true) then PT113
  begin
    T := Tob_Acquis.findfirst(['PCN_TYPECONGE'], ['ACC'], false);
    while T <> nil do
    begin
      PAYECOUR := PAYECOUR + T.getvalue('PCN_APAYES');
      T.free;
      T := Tob_Acquis.findnext(['PCN_TYPECONGE'], ['ACC'], false)
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Procedure qui annule les congés acquis générés dans une
Suite ........ : session de paye entre datedebut et datefin par requête SQL
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AnnuleCongesAcquis(CodeSalarie: string; datedebut, datefin: tdatetime);
begin
  EXECUTESQL('DELETE FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + CodeSalarie +
    '" AND PCN_DATEVALIDITE <= "' + UsDateTime(datefin) +
    '" AND PCN_DATEVALIDITE >= "' + UsDateTime(datedebut) +
    '" AND PCN_TYPEMVT = "CPA"' + { PT32 AJout AND PCN_TYPEMVT="CPA" }
    '  AND (PCN_TYPECONGE = "ACS" OR PCN_TYPECONGE = "ACQ" OR PCN_TYPECONGE = "ACA")');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Procedure qui annule les congés acquis générés de la TOB
Suite ........ : dans
Suite ........ : une session de paye entre datedebut et datefin
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AnnuleCongesAcquisDeTob(CodeSalarie: string; datedebut, datefin: tdatetime; T_A: tob);
var
  T: Tob;
  TC: string;
begin
  if not ((Assigned(T_a) and (T_a.detail.count > 0))) {PT113 notVide(T_a, true)} then exit;
  T := T_A.FindFirst(['PCN_DATEDEBUT', 'PCN_DATEFIN'], [datedebut, datefin], true);
  while T <> nil do
  begin
    TC := T.getvalue('PCN_TYPECONGE');
    if ((TC = 'ACA') or (TC = 'ACS') or (TC = 'ACQ')) then T.free;
    T := T_A.Findnext(['PCN_DATEDEBUT', 'PCN_DATEFIN'], [datedebut, datefin], true);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Procedure qui charge les congés acquis à supprimer selon
Suite ........ : une
Suite ........ : date de validité comprise entre debut et fin validité
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure ChangeCongesAcquisDeTob(CodeSalarie: string; datedebut, datefin: tdatetime; T_A: tob);
var
  st: string;
begin
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + CodeSalarie +
    '" AND PCN_DATEVALIDITE <= "' + UsDateTime(datefin) +
    '" AND PCN_DATEVALIDITE >= "' + UsDateTime(datedebut) +
    '" AND PCN_TYPEMVT = "CPA"' + { PT32 AJout AND PCN_TYPEMVT="CPA" }
    ' AND (PCN_TYPECONGE = "ACS" OR PCN_TYPECONGE = "ACQ" OR PCN_TYPECONGE = "ACA")';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if (Q.eof) then
  begin
    Ferme(Q);
    exit;
  end;
  Tob_AcquisAVirer.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
  Ferme(Q);
}
  Tob_AcquisAVirer.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Procedure qui annule en masse les congés acquis selon
Suite ........ : une date de validité comprise entre debut et fin validité
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AnnuleCongesAcquisMasse(datedebut, datefin: tdatetime);
begin
  EXECUTESQL('DELETE FROM ABSENCESALARIE WHERE ' +
    '  PCN_DATEVALIDITE <= "' + UsDateTime(datefin) +
    '" AND PCN_DATEVALIDITE >= "' + UsDateTime(datedebut) +
    '" AND PCN_TYPEMVT = "CPA"' + //PT32 AJout AND PCN_TYPEMVT="CPA"
    ' AND (PCN_TYPECONGE = "ACS" OR PCN_TYPECONGE = "ACQ" OR PCN_TYPECONGE = "ACA")');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Procedure qui supprime en masse tous les congés générés,
Suite ........ : traités dans une session de paie
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure SupprimeCPMasse(DD, DF: tdatetime);
var
  st, EtabEncours, SalEnCours: string;
  Q: tQuery;
  T, Ta, T_Pris, T_Acquis, T_PrisASup, T_AcqASup: Tob;
  DTClot: TDatetime;
  ModSal: boolean;
begin
  T_Pris := Tob.create('PRIS A ANNULER', nil, -1);
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'AND (PCN_TYPECONGE = "PRI"  OR PCN_TYPECONGE = "SLD" ) ' +
    'AND PCN_DATEPAIEMENT > "' + Usdatetime(10) + '" ' +
    'AND (PCN_DATEFIN>= "' + Usdatetime(DD) + '" AND PCN_DATEFIN <= "' + Usdatetime(DF) + '") ' +
    'AND PCN_GENERECLOTURE <> "X" ORDER BY PCN_SALARIE,PCN_ETABLISSEMENT';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if (Q.eof) then
  begin
    Ferme(Q);
    exit;
  end;
}
  T_Pris.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
//Ferme(Q);
  if T_Pris.detail.count <= 0 then exit;

  T_Acquis := Tob.create('ACQUIS A CREDITER', nil, -1);
  st := 'SELECT * FROM ABSENCESALARIE WHERE ' +
    'PCN_CODETAPE <> "C" AND PCN_TYPEMVT="CPA" AND ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'PCN_APAYES >= 0 AND ' +
    '(PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS" OR ' +
    ' PCN_TYPECONGE="ARR" OR PCN_TYPECONGE="REP" OR ' +
    ' ((PCN_TYPECONGE="REL" OR PCN_TYPECONGE ="AJU" OR PCN_TYPECONGE ="AJP") AND PCN_SENSABS = "+"))' +
    ' ORDER BY PCN_SALARIE ASC, PCN_DATEVALIDITE DESC,PCN_ETABLISSEMENT ASC';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if Q.eof then
  begin
    Ferme(Q);
    exit;
  end; //pas de congés pris
}
  T_Acquis.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
//  Ferme(Q);
  if T_Acquis.detail.count <= 0 then exit;

  AcquisMaj := false;
  T_prisASup := Tob.create('PRIS A SUPPRIMER', nil, -1);
  T_AcqASup := Tob.create('ACQUIS A SUPPRIMER', nil, -1);
  T := T_Pris.FindFirst([''], [''], True);
  EtabEncours := '';
  SalEnCours := '';
  DTClot := 0;
  while T <> nil do
  begin
    ModSal := false;
    if EtabEncours <> T.getvalue('PCN_ETABLISSEMENT') then
    begin
      Q := Opensql('SELECT ETB_DATECLOTURECPN FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' +
        T.getvalue('PCN_ETABLISSEMENT') + '"', True);
      if not Q.eof then
        DTClot := Q.findfield('ETB_DATECLOTURECPN').AsVariant;
      Ferme(Q);
      EtabEncours := T.getvalue('PCN_ETABLISSEMENT');
    end;
    if SalEnCours <> T.getvalue('PCN_SALARIE') then
    begin
      SalEnCours := T.getvalue('PCN_SALARIE');
      ModSal := True;
    end;
    if ModSal then
    begin
      while t <> nil do
      begin
        T.changeParent(T_prisAsup, -1);
        T := T_Pris.Findnext(['PCN_SALARIE', 'PCN_ETABLISSEMENT'], [SalEnCours, EtabEncours], True);
      end;
      Ta := T_Acquis.Findfirst(['PCN_SALARIE', 'PCN_ETABLISSEMENT'], [SalEnCours, EtabEncours], True);
      while ta <> nil do
      begin
        Ta.changeParent(T_AcqAsup, -1);
        Ta := T_Acquis.Findnext(['PCN_SALARIE', 'PCN_ETABLISSEMENT'], [SalEnCours, EtabEncours], True);
      end;
      SupprimeCP(T_prisASup, T_acqASup, SalEnCours, EtabEnCours, DTClot, DD, DF, '');
    end;
    T := T_Pris.FindNext([''], [''], True);
  end;

  T_acqASup.SetAllModifie(TRUE);
  T_acqASup.InsertOrUpdateDB(false);
  T_prisASup.SetAllModifie(TRUE);
  T_prisASup.InsertOrUpdateDB(false);


  if T_Pris <> nil then T_Pris.free;
  if T_Acquis <> nil then T_Acquis.free;
  if T_prisASup <> nil then T_prisASup.free;
  if T_acqASup <> nil then T_acqASup.free;

  st := 'DELETE FROM ABSENCESALARIE WHERE ' +
    ' (PCN_TYPECONGE="SLD" OR PCN_TYPECONGE="ARR") AND PCN_TYPEMVT="CPA" AND PCN_DATEVALIDITE<="' +
    Usdatetime(DF) + '" AND PCN_DATEVALIDITE>="' + Usdatetime(DD) + '" AND PCN_GENERECLOTURE="-"';
  ExecuteSql(st);


end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Procedure qui supprime tous les congés générés,  traités
Suite ........ : dans une session de paie
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure SupprimeCP(T_pris, T_acquis: tob; SalEnCours, EtabEnCours: string; DTClot, DD, FF: tdatetime; StWhere: string); { PT106 }
var
  Codergrpt: string;
  TP, T_SUP: Tob;
  CptPer, CptPer0, CptPer1, NbConsRelMvt: double;
  DtSld: TDateTime; { PT106 }
begin
  T_SUP := Tob.create('tob à supprimer', nil, -1);
  NbConsRelMvt := 0;
  DtSld := Idate1900; { PT106 }
  { PT-6 partie mis en commentaire puis supprimer  }

  { 1) on remet à blanc les mvts payés sur le bulletin en question }
  { on vire la notion d'établissement }
  TP := T_Pris.findfirst(['PCN_SALARIE'], [SalEnCours], True);
  while TP <> nil do
  begin
    Codergrpt := TP.getvalue('PCN_CODERGRPT');
    if Pos('Reliquat', TP.GetValue('PCN_PERIODEPAIE')) > 0 then
      NbConsRelMvt := NbConsRelMvt + TP.getvalue('PCN_JOURS');
    { PT-3  ligne modifié le 29/06/2001 Nouveau code : TP.Getvalue('PCN_CODERGRPT') = '-1' }
    { PT-6  le PT-3 a été remodifié,l'ancien code remis en place }
    if ((TP.Getvalue('PCN_MVTDUPLIQUE') <> 'X') and
      ((TP.getvalue('PCN_GENERECLOTURE') <> 'X') or
      (TP.getvalue('PCN_GENERECLOTURE') = 'X') and (TP.getvalue('PCN_TYPECONGE') = 'SLD'))) then
    begin
      { PT-6 partie mis en commentaire puis supprimer }
      if (TP.getvalue('PCN_GENERECLOTURE') <> 'X') and (TP.getvalue('PCN_TYPECONGE') = 'SLD') then
        DtSld := TP.GetValue('PCN_DATEFIN'); { PT106 }
      TP.PutValue('PCN_DATEPAIEMENT', 0);
      TP.PutValue('PCN_DATEDEBUT', idate1900); //PT91
      TP.PutValue('PCN_DATEFIN', idate1900); //PT91
      TP.PutValue('PCN_ABSENCE', 0);
      TP.PutValue('PCN_ABSENCEMANU', 0);
      TP.PutValue('PCN_MODIFABSENCE', '-');
      TP.PutValue('PCN_BASE', 0);
      TP.PutValue('PCN_VALOX', 0);
      TP.PutValue('PCN_VALOMS', 0);
      TP.PutValue('PCN_VALORETENUE', 0);
      TP.PutValue('PCN_VALOMANUELLE', 0);
      TP.PutValue('PCN_MODIFVALO', '-');
      TP.PutValue('PCN_PERIODEPAIE', '');
      TP.PutValue('PCN_CODERGRPT', TP.getValue('PCN_ORDRE'));
      TP.Putvalue('PCN_PERIODECP', CalculPeriode(DTClot, TP.GetValue('PCN_DATEVALIDITE')));
      TP.PutValue('PCN_PERIODEPY', -1);
      TP.PutValue('PCN_CODETAPE', '...'); //PT-13
      if (VH_Paie.PGEcabMonoBase) and (TP.getValue('PCN_SAISIEDEPORTEE') = 'X') then TP.putvalue('PCN_EXPORTOK', '-'); { PT99 }
    end;
    TP := T_Pris.findnext(['PCN_SALARIE'], [SalEnCours], True);
  end;
  { On vire les mvts dupliqués et solde de tout compte }
  CptPer := 0;
  CptPer0 := 0;
  CptPer1 := 0; { PT-6 Ajout ligne }
  TP := T_Pris.findfirst(['PCN_SALARIE'], [SalEnCours], True);
  while TP <> nil do
  begin
    if ((TP.Getvalue('PCN_MVTDUPLIQUE') = 'X') or
      ((TP.getvalue('PCN_TYPECONGE') = 'SLD')
      and (TP.getvalue('PCN_GENERECLOTURE') <> 'X'))) then
      TP.changeparent(T_Sup, -1);
    if TP.Getvalue('PCN_MVTDUPLIQUE') = '-' then
      CptPer := CptPer + TP.Getvalue('PCN_JOURS'); { DEB1 PT-6 Ajout ligne }
    if (TP.Getvalue('PCN_MVTDUPLIQUE') = 'X') and (TP.Getvalue('PCN_PERIODECP') = 0) then
      CptPer0 := CptPer0 + TP.Getvalue('PCN_JOURS');
    if (TP.Getvalue('PCN_MVTDUPLIQUE') = 'X') and (TP.Getvalue('PCN_PERIODECP') = 1) then
      CptPer1 := CptPer1 + TP.Getvalue('PCN_JOURS'); //FIN1 PT-6 Ajout ligne
    TP := T_Pris.findnext(['PCN_SALARIE'], [SalEnCours], True);
  end;
  T_Sup.DeleteDB(false);
  if T_sup <> nil then T_sup.free;

  { 2) on crédite les mvts acquis débités à hauteur du nb de jours payés }
  { Déconsomme les mvt en fonction des 2 périodes }
  if ((CptPer0 <> 0) or (CptPer1 <> 0)) and (CptPer = (CptPer0 + CptPer1)) then { DEB2 PT-6 Ajout ligne }
  begin
    DeconsommeMvtAcq(T_Acquis, SalEnCours, 0, CptPer0, NbConsRelMvt, FF);
    DeconsommeMvtAcq(T_Acquis, SalEnCours, 1, CptPer1, NbConsRelMvt, FF);
  end
  else

    DeconsommeMvtAcq(T_Acquis, SalEnCours, -10, CptPer, NbConsRelMvt, FF); { FIN2 PT-6 Ajout ligne }

  DeSoldeTobDelta(SalEnCours, DtSld, StWhere); { PT106 }
  { PT-6 partie mis en commentaire puis supprimer   }
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Procedure qui déconsomme les mvts soldés et/ou
Suite ........ : consommés : cas d'un solde de tout compte ou CP PRI
Mots clefs ... : PAIE;CP
*****************************************************************}
{ PT-6 Nouvelle procédure : DeconsommeMvtAcq }

procedure DeconsommeMvtAcq(T_Acquis: Tob; Sal: string; Periode: integer; CptPer, NbConsRelMvt: Double; FF: TDateTime);
var
  TA: TOB;
  Nb_J: double;
  TypeImpute: string;
begin
 if CptPer > 0 then
  begin
    { DEB PT74-3 On déconsomme les acquis sur reliquat ou ajustement de reliquat }
    while (NbConsRelMvt > 0) and (CptPer > 0) and (Periode <> 0) do
    begin
      { ajustement sur reliquat }
      TA := T_Acquis.findfirst(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE', 'PCN_TYPEIMPUTE'], [Sal, '1', 'AJU', 'REL'], True);
      while Assigned(TA) do
      begin
        Nb_J := Ta.getvalue('PCN_APAYES');
        if (NbConsRelMvt > 0) and (Nb_j <> 0) then
        begin
          if (Nb_J > NbConsRelMvt) then
            Ta.putValue('PCN_APAYES', Ta.getvalue('PCN_APAYES') - NbConsRelMvt)
          else
            Ta.putValue('PCN_APAYES', 0);
        end;
        NbConsRelMvt := Arrondi(NbConsRelMvt - Nb_J, DCP);
        CptPer := Arrondi(CptPer - Nb_J, DCP);
        TA := T_Acquis.findNext(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE', 'PCN_TYPEIMPUTE'], [Sal, '1', 'AJU', 'REL'], True);
      end;
      { reliquat }
      if NbConsRelMvt > 0 then
      begin
        TA := T_Acquis.findfirst(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Sal, '1', 'REL'], True);
        while Assigned(TA) do
        begin
          Nb_J := Ta.getvalue('PCN_APAYES');
          if (NbConsRelMvt > 0) and (Nb_j <> 0) then
          begin
            if (Nb_J > NbConsRelMvt) then
            begin
              Ta.putValue('PCN_APAYES', Ta.getvalue('PCN_APAYES') - NbConsRelMvt);
              CptPer := Arrondi(CptPer - NbConsRelMvt, DCP); { PT105 }
              NbConsRelMvt := 0; { PT105 }
            end
            else
            begin
              Ta.putValue('PCN_APAYES', 0);
              NbConsRelMvt := Arrondi(NbConsRelMvt - Nb_J, DCP);
              CptPer := Arrondi(CptPer - Nb_J, DCP);
            end;
          end;
          TA := T_Acquis.findNext(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Sal, '1', 'REL'], True);
        end;
      end;
    end;
    { FIN PT74-3 }
    { On déconsomme les acquis }
    if Periode = -10 then
      TA := T_Acquis.findfirst(['PCN_SALARIE'], [Sal], True)
    else
      TA := T_Acquis.findfirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Sal, Periode], True);
    while (TA <> nil) do
    begin
      if Ta.GetValue('PCN_PERIODECP') < 2 then { PT104 Ne traite pas les périodes clôturés }
      begin
        Nb_J := Ta.getvalue('PCN_APAYES');
        if (CptPer > 0) and (Nb_j <> 0) then
        begin
           // PT86
          if TA.getvalue('PCN_TYPEIMPUTE') <> NULL then TypeImpute := TA.getvalue('PCN_TYPEIMPUTE')
          else TypeImpute := '';
          if not (((Ta.getvalue('PCN_TYPECONGE') = 'REL') or ((Ta.getvalue('PCN_TYPECONGE') = 'AJU') and (TypeImpute = 'REL'))) and (Ta.getvalue('PCN_SENSABS') = '+')) then { PT74-3  FIN PT86}
          begin
            if (Nb_J > CptPer) then
              Ta.putValue('PCN_APAYES', Ta.getvalue('PCN_APAYES') - CptPer)
            else
              Ta.putValue('PCN_APAYES', 0);
          end;
           (*  PT74-3 Mise en commentaire puis supprimé  *)
          CptPer := Arrondi(CptPer - Nb_J, DCP);
        end;
        Ta.PutValue('PCN_CODETAPE', '...'); //PT-13
      end;
      if Periode = -10 then
        TA := T_Acquis.findnext(['PCN_SALARIE'], [Sal], True)
      else
        TA := T_Acquis.findnext(['PCN_SALARIE', 'PCN_PERIODECP'], [Sal, Periode], True);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Procedure qui annule les mvts PRI ET SLD sur congés du
Suite ........ : bulletin de paye en cours d'annulation
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure AnnuleCongesPris(CodeSalarie, Etab: string; datedebut, datefin: tdatetime);
var
  st, StWhere: string;
  Q: tQuery;
  T_Pris, T_Acquis: Tob;
  DTClot: TDatetime;
begin
  Q := Opensql('SELECT ETB_DATECLOTURECPN FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' +
    Etab + '"', True);
  if not Q.eof then
    DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime
  else
    DTClot := Idate1900;
  Ferme(Q);
  { 1) on remet à blanc les mvts payés sur le bulletin en question }

  T_Pris := Tob.create('PRIS A ANNULER', nil, -1);
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'AND PCN_SALARIE = "' + CodeSalarie + '" '; { PT37 Ajout clause AND PCN_VALIDRESP = "VAL" }
  // DEB PT84
  if not FirstBull then st := st + 'AND PCN_DATEFIN = "' + Usdatetime(datefin) + '" '
  else // DEB PT85
    st := st + 'AND (PCN_DATEDEBUT = "' + Usdatetime(dateDebut) +
      '" AND ((PCN_DATEFIN = "' + Usdatetime(Idate1900) + '" AND PCN_MODIFVALO <>"X")) OR PCN_DATEFIN = "' + Usdatetime(datefin) + '") ';
  // FIN PT85
  st := st + ' AND ((((PCN_TYPECONGE = "PRI" AND PCN_VALIDRESP="VAL") OR PCN_TYPECONGE = "SLD" ) ' +
    // FIN PT84
//      'AND PCN_GENERECLOTURE <>"X" '+ PT65 On charge aussi les SLD de cloture
  'AND PCN_DATEPAIEMENT > "' + Usdatetime(10) + '") ' +
    ' OR (PCN_TYPECONGE = "SLD" AND PCN_GENERECLOTURE="-" AND PCN_CODETAPE="..."))'; //PT66 Suppr. des SLD de sortie non payé
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if (Q.eof) then
  begin
    Ferme(Q);
    if T_Pris <> nil then T_Pris.free; // PT115-3
    exit;
  end;
}
  T_Pris.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
//  Ferme(Q);
  if T_Pris.detail.count <= 0 then
  begin
    T_Pris.free;
    exit;
  end;
  { DEB PT-27-2 On récupère les mvts d'acquis posterieur au dernier SLD }
  { Pour ne pas détoper les mvts d'acquis soldés des SLD antérieur }
  st := 'SELECT PCN_DATEVALIDITE FROM ABSENCESALARIE WHERE PCN_SALARIE="' + CodeSalarie + '" ' +
    'AND PCN_DATEVALIDITE<"' + Usdatetime(datefin) + '" AND PCN_TYPECONGE="SLD" AND PCN_TYPEMVT="CPA" ' + //PT32 AJout PCN_TYPEMVT="CPA"
    'AND PCN_GENERECLOTURE="-" ' + //PT64
    'ORDER BY PCN_DATEVALIDITE DESC';
  Q := OpenSql(st, TRUE);
  if (not Q.eof) then
    StWhere := ' AND PCN_DATEVALIDITE>"' + USDateTime(Q.FindField('PCN_DATEVALIDITE').AsDateTime) + '" '
  else
    StWhere := '';
  Ferme(Q);
  {FIN PT-27-2}
  T_Acquis := Tob.create('ACQUIS A CREDITER', nil, -1);
  st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
    'AND PCN_SALARIE = "' + CodeSalarie + '" ' +
    'AND PCN_CODETAPE <> "C" AND PCN_APAYES >= 0 ' +
// PT128    'AND (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS" ' +
    'AND (PCN_TYPECONGE="ACQ" OR PCN_TYPECONGE="ACA" OR PCN_TYPECONGE="ACS" OR PCN_TYPECONGE="ACF" ' +
    'OR PCN_TYPECONGE="ARR" OR PCN_TYPECONGE="REP" ' +
    'OR ((PCN_TYPECONGE="REL" OR PCN_TYPECONGE ="AJU" OR PCN_TYPECONGE ="AJP") AND PCN_SENSABS = "+")) ' +
    'AND PCN_DATEVALIDITE<="' + Usdatetime(datefin) + '" ' +
    StWhere + { PT-27-2 ajout clause DATEVALIDITE ET StWhere }
    ' ORDER BY PCN_DATEVALIDITE DESC,PCN_ORDRE DESC';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  if Q.eof then
  begin
    Ferme(Q);
    if T_Pris <> nil then T_Pris.free; // PT115-3
    if T_Acquis <> nil then T_Acquis.free;  // PT115-3
    exit;
  end; // pas de congés pris
}
  T_Acquis.LoadDetailDBFROMSQL('ABSENCESALARIE', st);
  if T_Acquis.detail.count <= 0 then
  begin
    if T_Pris <> nil then T_Pris.free; // PT115-3
    if T_Acquis <> nil then T_Acquis.free; // PT115-3
    exit;
  end;
  AcquisMaj := false;

  SupprimeCP(T_pris, T_acquis, CodeSalarie, Etab, DTClot, DateDebut, DateFin, StWhere); { PT106 }

  T_Acquis.SetAllModifie(TRUE);
  T_Acquis.InsertOrUpdateDB(false);
  T_Pris.SetAllModifie(TRUE);
  T_Pris.InsertOrUpdateDB(false);

  if T_Pris <> nil then T_Pris.free;
  if T_Acquis <> nil then T_Acquis.free;


  st := 'DELETE FROM ABSENCESALARIE WHERE PCN_SALARIE ="' + CodeSalarie +
    '" AND (PCN_TYPECONGE="SLD" OR PCN_TYPECONGE="ARR") AND PCN_TYPEMVT="CPA" AND PCN_DATEVALIDITE="' +
    Usdatetime(datefin) + '" AND PCN_GENERECLOTURE="-"';
  ExecuteSql(st);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : procédure qui génère le mouvement d' arrondi
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure CreeArrondi(Arr: double; Salarie, etab, clot: string; var suivant: integer; periode: integer; dateE: tdatetime; T: tob);
var
  Q: tquery;
  st: string;
begin
  st := 'SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL' +
    ' FROM SALARIES WHERE PSA_SALARIE = "' + Salarie + '"';
  Q := opensql(st, true);
  if clot = '-' then OrdreArrBull := Suivant;
  InitialiseTobAbsenceSalarie(T);
  {  PT50   20/11/2002 V591 PH Portage eAGL }
  if (not Q.EOF) and (Q <> nil) then
  begin
    T.Putvalue('PCN_TRAVAILN1', Q.findfield('PSA_TRAVAILN1').AsString);
    T.Putvalue('PCN_TRAVAILN2', Q.findfield('PSA_TRAVAILN2').AsString);
    T.Putvalue('PCN_TRAVAILN3', Q.findfield('PSA_TRAVAILN3').AsString);
    T.Putvalue('PCN_TRAVAILN4', Q.findfield('PSA_TRAVAILN4').AsString);
    T.Putvalue('PCN_CODESTAT', Q.findfield('PSA_CODESTAT').AsString);
    T.Putvalue('PCN_CONFIDENTIEL', Q.findfield('PSA_CONFIDENTIEL').AsString);
  end;
  T.PutValue('PCN_SALARIE', Salarie);
  T.PutValue('PCN_RESSOURCE', '');
  T.PutValue('PCN_ORDRE', Suivant);
  T.PutValue('PCN_TYPECONGE', 'ARR');
  T.PutValue('PCN_SENSABS', '+');
  T.PutValue('PCN_LIBELLE', 'Arrondi au ' + datetostr(dateE));
  T.PutValue('PCN_DATEMODIF', Date);
  T.PutValue('PCN_DATEDEBUT', dateE);
  T.PutValue('PCN_DATEFIN', dateE);
  T.PutValue('PCN_DATEVALIDITE', dateE);
  T.PutValue('PCN_PERIODECP', periode);
  T.PutValue('PCN_JOURS', Arrondi(Arr, DCP));
  T.PutValue('PCN_ETABLISSEMENT', etab);
  T.PutValue('PCN_GENERECLOTURE', clot);
  suivant := suivant + 1;
  Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction d'initialisation d'un mouvement congés payés
Suite ........ : A appeler avant tout renseignement de nouveau mvt
Mots clefs ... : PAIE;CP
*****************************************************************}

(* PT89 procedure InitialiseTobAbsenceSalarie(TCP: tob);
begin
  TCP.Putvalue('PCN_TYPEMVT'             , 'CPA' );
  TCP.Putvalue('PCN_SALARIE'             , ''    );
  TCP.Putvalue('PCN_DATEDEBUT'           , 0     );
  TCP.Putvalue('PCN_DATEFIN'             , 0     );
  TCP.Putvalue('PCN_ORDRE'               , 0     );
  TCP.Putvalue('PCN_MVTPRIS'             , ''    );          { PT88 }
  TCP.PutValue('PCN_PERIODECP'           , -1    );
  TCP.PutValue('PCN_PERIODEPY'           , -1    );
  TCP.Putvalue('PCN_TYPECONGE'           , ''    );
  TCP.Putvalue('PCN_SENSABS'             , ''    );
  TCP.Putvalue('PCN_LIBELLE'             , ''    );
  TCP.Putvalue('PCN_GENERECLOTURE'       ,'-'    );
  //  TCP.Putvalue('PCN_DATEMODIF', 0);                      { PT88 }
  TCP.Putvalue('PCN_DATESOLDE'           , 0     );
  TCP.Putvalue('PCN_DATEVALIDITE'        , 0     );
  TCP.Putvalue('PCN_DATEDEBUTABS'        , 0     );
  TCP.Putvalue('PCN_DEBUTDJ'             , ''    );          { PT88 }
  TCP.Putvalue('PCN_DATEFINABS'          , 0     );
  TCP.Putvalue('PCN_FINDJ'               , ''    );          { PT88 }
  TCP.Putvalue('PCN_DATEPAIEMENT'        , 0     );
//  TCP.Putvalue('PCN_ETATPOSTPAIE'        , 0);
  TCP.Putvalue('PCN_CODETAPE'            , '...' ); //PT-13
  TCP.Putvalue('PCN_JOURS'               , 0     );
  TCP.Putvalue('PCN_HEURES'              , 0     );          { PT88 }
  TCP.Putvalue('PCN_BASE'                , 0     );
  TCP.Putvalue('PCN_NBREMOIS'            , 0     );
  TCP.Putvalue('PCN_CODERGRPT'           , 0     );
  TCP.PutValue('PCN_MVTDUPLIQUE'         , '-'   );
  TCP.Putvalue('PCN_ABSENCE'             , 0     );
  TCP.Putvalue('PCN_MODIFABSENCE'        , '-'   );
  TCP.Putvalue('PCN_APAYES'              , 0     );
  TCP.Putvalue('PCN_VALOX'               , 0     );
  TCP.Putvalue('PCN_VALOMS'              , 0     );
  TCP.Putvalue('PCN_VALORETENUE'         , 0     );
  TCP.Putvalue('PCN_VALOMANUELLE'        , 0     );          { PT88 }
  TCP.Putvalue('PCN_MODIFVALO'           , '-'   );
  TCP.Putvalue('PCN_PERIODEPAIE'         ,''     );
  TCP.Putvalue('PCN_ETABLISSEMENT'       , ''    );
  TCP.Putvalue('PCN_TRAVAILN1'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN2'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN3'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN4'           , ''    );
  TCP.Putvalue('PCN_CODESTAT'            , ''    );          { PT88 }
// TCP.PutValue('PCN_CONFIDENTIEL','-');
//  TCP.Putvalue('PCN_TOPCONVERT'            , '');
  TCP.Putvalue('PCN_SAISIEDEPORTE'       , '-'   );          { DEB PT88 }
  TCP.Putvalue('PCN_VALIDSALARIE'        , ''    );
  TCP.Putvalue('PCN_VALIDRESP'           , ''    );
  TCP.Putvalue('PCN_EXPORTOK'            , ''    );
  TCP.Putvalue('PCN_LIBCOMPL1'           , ''    );
  TCP.Putvalue('PCN_LIBCOMPL2'           , ''    );
  TCP.Putvalue('PCN_VALIDABSENCE'        , ''    );
  TCP.Putvalue('PCN_OKFRACTION'          , ''    );
  TCP.Putvalue('PCN_NBJCARENCE'          , 0     );
  TCP.Putvalue('PCN_NBJCALEND'           , 0     );
  TCP.Putvalue('PCN_NBJIJSS'             , 0     );
  TCP.Putvalue('PCN_IJSSSOLDEE'          , '-'   );
  TCP.Putvalue('PCN_GESTIONIJSS'         , '-'   );          { FIN PT88 }
end;    *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : fonction qui rend true si mouvement d'acquisition congés
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
function IsAcquis(TypeConge, Sens: string): boolean;
begin
  result := false;
  if ((typeconge = 'ACQ') or
    (typeconge = 'ACS') or
    (typeconge = 'ACA') or
    (typeconge = 'AJU') or
    (typeconge = 'ARR') or
    ((typeconge = 'REL') and (Sens = '+')) or
    (typeconge = 'REP')) then result := true;
end; *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : fonction qui rend true si mouvement de prise congés
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
function IsPris(TypeConge: string): boolean;
begin
  result := false;
  if ((typeconge = 'PRI') or (typeconge = 'SLD') or
    (typeconge = 'AJP') or (typeconge = 'CPA')) then result := true;
end;  *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : --- CLOTURE D'UNE SESSION DE CONGES PAYES ---
Suite ........ : 1 - Suppression de mvts pouvant correspondre à une
Suite ........ : clôture de cette période CP
Suite ........ : 2 - Calcul du reliquat de congés à reporter sur les acquis de
Suite ........ : la nouvelle période d'acquisition.1 reliquat = 2 mvts (- fin de
Suite ........ : période cloturée, + début nvelle période)
Suite ........ : 3 - Calcul des congés acquis sur la période d'acquisition et
Suite ........ : arrondi à l'entier  supérieur
Mots clefs ... : PAIE;CP
*****************************************************************}

function CloturePeriodeCP(datefinperiode: TDatetime; Etablissement, MethodeRel: string): boolean;
var
  Datedebutperiode, DatedebutperiodeP, DatefinPeriodeP: TDateTime;
  DateDebutPeriodeS, DateSortie: TDateTime;
  salarie, st, MethodeReliquat, GestionRel, LibMvt, Per, PerN1: string;
  aa, mm, jj: word;
  Noordre, Reliquat {,PT62-3 DecalagePer}: integer;
  T_listesal, T_MVTINS, T_listeMvt, TR, TS, T1: tob;
  Arr, SommePArr, PN, AN, RN, BN, MB, PN1, AN1, RN1, BN1, MB1, Valo: double;
begin
  { Calcul des débuts et fin de période sur lesquelles on va travailler }
  try
    begintrans;
    InitMoveProgressForm(nil, 'Clôture des congés payés', 'Veuillez patienter SVP ...', 100, FALSE, TRUE);
    Decodedate(datefinperiode, aa, mm, jj);
    DateDebutPeriodeS := datefinperiode + 1;
    //DateFinPeriodeS   := (EncodeDate(aa+1,mm,jj));
    Datedebutperiode := (PGEncodeDateBissextile(aa - 1, mm, jj) + 1); { PT71 }
    DatedebutperiodeP := (PGEncodeDateBissextile(aa - 2, mm, jj) + 1); { PT71 }
    DatefinPeriodeP := DatedebutPeriode - 1;

    { Suppression des éventuels mouvements générés par une clôture de cette }
    { même session de CP }
    ExecuteSql('DELETE FROM ABSENCESALARIE WHERE PCN_ETABLISSEMENT= "' + etablissement + '"' +
      'AND PCN_GENERECLOTURE = "X" AND PCN_TYPEMVT="CPA" ' + { PT32 AJout PCN_TYPEMVT="CPA" }
      'AND ((PCN_TYPECONGE="REL" AND PCN_DATEVALIDITE="' + UsDateTime(DatefinPeriodeP) + '") ' +
      'OR  ((PCN_TYPECONGE="REL" OR PCN_TYPECONGE="SLD") AND PCN_DATEVALIDITE="' + UsDateTime(DatedebutPeriode) + '") ' +
      'OR   (PCN_TYPECONGE="ARR" AND PCN_DATEVALIDITE="' + UsDateTime(DateFinPeriode) + '"))');

    T_MVTINS := Tob.create('Mere des lignes a inserer', nil, -1);
    T_ListeSal := TOB.Create('Liste des salarie etablissement', nil, -1);
    st := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,' +
      'PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CPTYPERELIQ,PSA_RELIQUAT,PSA_DATESORTIE ' +
      'FROM SALARIES WHERE PSA_ETABLISSEMENT ="' + Etablissement + '"';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    T_ListeSal.LoadDetailDB('SALARIES', '', '', Q, False);
    ferme(Q);
}
    T_ListeSal.LoadDetailDBFROMSQL('SALARIES', st);

    T_listeMvt := Tob.create('ABSENCESALARIE', nil, -1);
    st := 'SELECT * FROM ABSENCESALARIE LEFT JOIN SALARIES ON PCN_SALARIE=PSA_SALARIE ' +
      'WHERE PSA_ETABLISSEMENT ="' + Etablissement + '" ' +
      'AND PCN_TYPEMVT="CPA" ' +
      'ORDER BY PCN_SALARIE , PCN_DATEVALIDITE ';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    T_listeMvt.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    ferme(Q);
}
    T_listeMvt.LoadDetailDBFROMSQL('ABSENCESALARIE', St);

    { Au préalable on change le parent des reliquats positive et consommés pour ne pas }
    { comptabiliser la base et le nbre de mois de ces mvts dans la nouvelle periode.. }
    TS := T_ListeSal.FindFirst([''], [''], false);
    while TS <> nil do
    begin //B TS
      DateSortie := TS.GetValue('PSA_DATESORTIE');
      Gestionrel := TS.getvalue('PSA_CPTYPERELIQ');
      if ((Gestionrel <> 'ETB') and (Gestionrel <> 'PER')) then GestionRel := 'ETB';
      if gestionrel = 'ETB' then MethodeReliquat := MethodeRel
      else
        if gestionrel = 'PER' then MethodeReliquat := TS.getvalue('PSA_RELIQUAT');
      Salarie := TS.getvalue('PSA_SALARIE');
      VerificationCP(Salarie, TS.getvalue('PSA_ETABLISSEMENT'), TRUE); //PT126
      NoOrdre := IncrementeSeqNoOrdre('CPA', SALARIE);
      MoveCurProgressForm('Clôture des mouvements du salarié : ' + Salarie + ' ' + TS.getvalue('PSA_LIBELLE') + ' ' + TS.getvalue('PSA_PRENOM'));
      if MethodeReliquat <> '' then Reliquat := StrToInt(MethodeReliquat) else Reliquat := 5;

      { Recupération du solde des CP pour la periode cloturé et la periode antérieur }
      Per := inttostr(CalculPeriode(DateFinPeriode, DateFinperiode));
      PerN1 := Inttostr(StrToInt(Per) + 1);

      CompteCp(T_listeMvt, Per, Salarie, 0, PN, AN, RN, BN, MB, Valo);
      CompteCp(T_listeMvt, PerN1, Salarie, 0, PN1, AN1, RN1, BN1, MB1, Valo);

      { DEB PT49 : On décremente au compteur la base et nb mois des reliquats positives et consommés du salarié }
      TR := T_listeMvt.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE', 'PCN_SENSABS'], [Salarie, PerN1, 'REL', '+'], False);
      while TR <> nil do
      begin
        if Arrondi(TR.GetValue('PCN_JOURS'), DCP) = Arrondi(TR.GetValue('PCN_APAYES'), DCP) then
        begin
          MB1 := MB1 - TR.GetValue('PCN_NBREMOIS');
          BN1 := BN1 - TR.GetValue('PCN_BASE');
        end;
        TR := T_listeMvt.FindNext(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE', 'PCN_SENSABS'], [Salarie, PerN1, 'REL', '+'], False);
      end;
      { FIN PT49 }
      { Cloture les mvts de la periode DateDebperiodeP à DateFinperiodeP }
      ClotureMvtAcquis(T_listeMvt, Salarie, Reliquat, DatedebutperiodeP, DatefinperiodeP);
      if (DateSortie >= DateDebutPeriodeS) or (Datesortie <= idate1900) then { PT100 }
      begin //Begin IF 1
        case Reliquat of
          0:
            begin { Remise à zero }
              { Genere reliquat negatif à DateFinperiodeP Cloturé }
              LibMvt := 'Reliquat - au ' + datetostr(datefinPeriodeP) + ' RAZ';
              if RN1 > 0 then
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 2, 'REL', '-', LibMvt, 'C', DatefinperiodeP, Datefinperiode, RN1, MB1, BN1);
            end;
          1:
            begin { Conservation des droits }
              { Genere reliquat negatif à DateFinperiodeP }
              { Genere reliquat positif à DateDebutperiode }
              if RN1 > 0 then
              begin
                LibMvt := 'Reliquat - au ' + datetostr(datefinPeriodeP);
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 2, 'REL', '-', LibMvt, 'C', DatefinperiodeP, Datefinperiode, RN1, MB1, BN1);
                NoOrdre := NoOrdre + 1;
                LibMvt := 'Reliquat + au ' + datetostr(datedebutPeriode);
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 1, 'REL', '+', LibMvt, '...', DateDebutPeriode, Datefinperiode, RN1, MB1, BN1); //PT-13
              end;
              { Incrémente la periode des mvts de + 1 }
            end;
          2:
            begin { Paiement du solde }
              { Genere reliquat negatif à DateFinperiodeP }
              { Genere reliquat positif à DateDebutperiode }
              { Genere un solde negatif à DateDebutperiode }
              if RN1 > 0 then
              begin
                LibMvt := 'Reliquat - au ' + datetostr(datefinPeriodeP) + '. Mvt soldé au ' + datetostr(DateDebutPeriode);
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 2, 'REL', '-', LibMvt, 'S', DatefinperiodeP, Datefinperiode, RN1, MB1, BN1); //PT61-5
                NoOrdre := NoOrdre + 1;
                LibMvt := 'Reliquat + au ' + datetostr(datedebutPeriode);
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 1, 'REL', '+', LibMvt, '...', DateDebutPeriode, Datefinperiode, RN1, MB1, BN1); //PT-13 //PT61-5
                NoOrdre := NoOrdre + 1;
                LibMvt := 'Solde ' + floattostr(Arrondi(RN1, DCP)) + 'j au ' + datetostr(datedebutPeriode);
                GenereMvtCloture(T_MVTINS, TS, NoOrdre, 1, 'SLD', '-', LibMvt, '...', DateDebutPeriode, Datefinperiode, RN1, MB1, BN1); //PT-13 //PT61-5
              end;
            end;
        end; //End du Case
        NoOrdre := NoOrdre + 1;
        { Genere un arrondi positif à DateFinPeriode }
        SommePArr := AN;
        // PT90 Mise en commentaire if MethodeReliquat = '1' then SommePArr := SommePArr + RN1; // il faut en tenir compte dans l'arrondi
        Arr := (1 + Int(SommePArr) - SommePArr);
        { génération d'un arrondi si nécessaire }
        if ((Arr > 0) and (Arr < 1)) then
        begin //B6
          T1 := TOB.Create('ABSENCESALARIE', T_MVTINS, -1);
          CreeArrondi(Arr, Salarie, TS.getvalue('PSA_ETABLISSEMENT'), 'X', NoOrdre, 1, DatefinPeriode, T1);
        end; //E6
      end; //End IF 1
      TS := T_ListeSal.FindNext([''], [''], false);
    end; //E TS
    { Incrémente la periode des mvts de + 1 }
    TS := T_ListeMvt.FindFirst([''], [''], false);
    while TS <> nil do
    begin //B2
      (* PT62-3 Mise en commentaire, incrémentation de la période
      if ((Ispris(TS.getvalue('PCN_TYPECONGE'))) and (TS.getvalue('PCN_DATEFIN')>10)) then
         begin //B3
         TempDate:=TS.getvalue('PCN_DATEPAIEMENT');        //SB 08/06/2001
         If TempDate=idate1900 then TempDate:=TS.getvalue('PCN_DATEVALIDITE');  //SB 08/06/2001
            DecalagePer:=0;
         {DEB PT-1}
         if (TS.GetValue('PCN_TYPECONGE')='SLD') AND (TS.getvalue('PCN_MVTDUPLIQUE')='X') AND (TS.GetValue('PCN_PERIODEPY')=1) then
            DecalagePer:=1;
         {FIN PT-1}
         {DEB PT-9}
         if (TS.GetValue('PCN_TYPECONGE')='PRI') AND (TS.GetValue('PCN_PERIODECP')=CalculPeriode(DateFinPeriodeS,TempDate)) then
            DecalagePer:=1;
         {FIN PT-9}
         TS.PutValue('PCN_PERIODECP',CalculPeriode(DateFinPeriodeS,TempDate)+DecalagePer);//PT1
         if TS.getvalue('PCN_PERIODECP') >1 then TS.PutValue('PCN_CODETAPE', 'C');
         end //E3
      else
         TS.PutValue('PCN_PERIODECP',CalculPeriode(DateFinPeriodeS,TS.getvalue('PCN_DATEVALIDITE'))); *)
      TS.PutValue('PCN_PERIODECP', TS.GetValue('PCN_PERIODECP') + 1);
      TS.UpdateDB(false);
      TS := T_ListeMvt.FindNext([''], [''], false);
    end; //E2
    if ((Assigned(T_MVTINS) and (T_MVTINS.detail.count > 0))) { PT113 notVide(T_MVTINS, True)} then
    begin //B1
      T_MVTINS.SetAllModifie(TRUE);
      T_MVTINS.InsertDB(nil, true);
      FreeAndNil(T_MVTINS); //T_MVTINS.Free; PT113
    end; //E1
    FiniMoveProgressForm;

    FreeAndNil(T_listesal); //if notVide(T_listesal, True) then T_listesal.free; PT113
    FreeAndNil(T_listeMvt); //if notVide(T_listeMvt, True) then T_listeMvt.free; PT113

    CommitTrans;
    result := True;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la clôture Congés payés', 'Clôture période CP');
    Result := False;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : DECLOTURE DE LA SESSION DE CONGES PAYES
Mots clefs ... : PAIE;CP
*****************************************************************}

function DeCloturePeriodeCP(FinPeriodeN: TDatetime; Etab, EtabLibelle: string): Boolean;
var
  DebPeriodeN1, DebPeriodeN2, FinPeriodeN1, FinPeriodeN2: TDateTime;
  Init, YY, MM, JJ: Word;
begin
  Result := False;
  if FinPeriodeN <= idate1900 then exit;

  FinPeriodeN1 := PlusDate(FinPeriodeN, -1, 'A');
  DecodeDate(FinPeriodeN1, YY, MM, JJ); { PT71 }
  FinPeriodeN1 := PGEncodeDateBissextile(YY, MM, JJ); { PT71 }

  FinPeriodeN2 := PlusDate(FinPeriodeN1, -1, 'A');
  DecodeDate(FinPeriodeN2, YY, MM, JJ); { PT71 }
  FinPeriodeN2 := PGEncodeDateBissextile(YY, MM, JJ); { PT71 }

  DebPeriodeN1 := FinPeriodeN2 + 1;

  DebPeriodeN2 := PlusDate(DebPeriodeN1, -1, 'A');
  DecodeDate(DebPeriodeN2, YY, MM, JJ); { PT71 }
  DebPeriodeN2 := PGEncodeDateBissextile(YY, MM, JJ); { PT71 }


  if ExisteSql('SELECT DISTINCT PPU_SALARIE FROM PAIEENCOURS ' +
    'WHERE PPU_ETABLISSEMENT="' + etab + '" ' +
    'AND PPU_DATEFIN>"' + UsDateTime(FinPeriodeN) + '" ') = TRUE then
  begin
    PGIBox('Vous ne pouvez "déclôturer" la période CP de l''établissement "' + EtabLibelle + '" : ' +
      'Des bulletins de paie sont alimentés pour la période du ' + DateToStr(FinPeriodeN + 1) + ' ' +
      'au ' + DateToStr(PlusMois(FinPeriodeN, 1)) + '.', '"Déclôture" de congés payés');
    exit;
  end;

  Init := PgiAsk('Attention, vous allez "declôturer" la période de congés payés ' +
    '#13#10( du ' + datetostr(DebPeriodeN2) + ' au ' + datetostr(finPeriodeN2) + '). ' +
    '#13#10Et réouvrir la période ( du ' + datetostr(DebPeriodeN1) + ' au ' + datetostr(finPeriodeN1) + ') ' +
    'de l''établissement "' + EtabLibelle + '" ', '"Déclôture" de congés payés');
  if Init = mrYes then
  begin
    try
      begintrans;
      InitMoveProgressForm(nil, '"Declôture" des congés payés de l''établissement ' + EtabLibelle, 'Veuillez patienter SVP ...', 7, FALSE, TRUE);
      MoveCurProgressForm('Suppression des mouvements générés par la clôture');
      ExecuteSql('DELETE FROM ABSENCESALARIE WHERE PCN_ETABLISSEMENT= "' + etab + '"' +
        'AND PCN_GENERECLOTURE = "X" ' +
        'AND ((PCN_TYPECONGE="REL" AND PCN_DATEVALIDITE="' + UsDateTime(FinPeriodeN2) + '") ' +
        'OR  ((PCN_TYPECONGE="REL" OR PCN_TYPECONGE="SLD") AND PCN_DATEVALIDITE="' + UsDateTime(DebPeriodeN1) + '") ' +
        'OR   (PCN_TYPECONGE="ARR" AND PCN_DATEVALIDITE="' + UsDateTime(FinPeriodeN1) + '"))');
      MoveCurProgressForm('Mise à jour des mouvements payés puis cloturés ou soldés sur la période(n-1).');
      ExecuteSql('UPDATE ABSENCESALARIE SET PCN_CODETAPE="P" ' +
        'WHERE PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + etab + '") ' +
        'AND (PCN_CODETAPE="C" OR PCN_CODETAPE="S") ' +
        'AND ((( PCN_DATEVALIDITE>="' + UsDateTime(DebPeriodeN1) + '" AND PCN_DATEVALIDITE<="' + UsDateTime(FinPeriodeN1) + '" ' +
        '  AND (PCN_MVTDUPLIQUE="X" OR PCN_MVTPRIS="PRI" OR PCN_PERIODEPY=1)) ' +
        'OR (   PCN_DATEVALIDITE>="' + UsDateTime(DebPeriodeN2) + '" AND PCN_DATEVALIDITE<="' + UsDateTime(FinPeriodeN2) + '" ' +
        '   AND PCN_TYPECONGE="CPA")))');
      MoveCurProgressForm('Mise à jour des mouvements payés puis soldés sur la période(n-2).');
      ExecuteSql('UPDATE ABSENCESALARIE SET PCN_CODETAPE="P" ' +
        'WHERE PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + etab + '") ' +
        'AND (PCN_CODETAPE="C" OR PCN_CODETAPE="S") ' +
        'AND PCN_DATEVALIDITE>="' + UsDateTime(DebPeriodeN2) + '" ' +
        'AND PCN_DATEVALIDITE<="' + UsDateTime(FinPeriodeN2) + '"  ' +
        'AND PCN_PERIODEPY<>-1 ' +
        'AND EXISTS (SELECT PCN_TYPEMVT FROM ABSENCESALARIE ABS WHERE ABSENCESALARIE.PCN_SALARIE=ABS.PCN_SALARIE ' +
        'AND PCN_TYPECONGE="SLD" AND PCN_GENERECLOTURE="-" AND PCN_TYPEMVT="CPA")');
      MoveCurProgressForm('Mise à jour des mouvements acquis puis soldés sur la période(n-2).');
      ExecuteSql('UPDATE ABSENCESALARIE SET PCN_CODETAPE="S" ' +
        'WHERE PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + etab + '") ' +
        'AND (PCN_CODETAPE="C" OR PCN_CODETAPE="S") ' +
        'AND PCN_DATEVALIDITE>="' + UsDateTime(DebPeriodeN2) + '" ' +
        'AND PCN_DATEVALIDITE<="' + UsDateTime(FinPeriodeN2) + '"  ' +
        'AND PCN_PERIODEPY=-1 ' +
        'AND EXISTS (SELECT PCN_TYPEMVT FROM ABSENCESALARIE ABS WHERE ABSENCESALARIE.PCN_SALARIE=ABS.PCN_SALARIE ' +
        'AND PCN_TYPECONGE="SLD" AND PCN_GENERECLOTURE="-" AND PCN_TYPEMVT="CPA")');
      MoveCurProgressForm('Mise à jour des mouvements acquis puis cloturés sur la période(n-2).');
      ExecuteSql('UPDATE ABSENCESALARIE SET PCN_CODETAPE="..." ' + //PT13
        'WHERE PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + etab + '") ' +
        'AND (PCN_CODETAPE="C" OR PCN_CODETAPE="S") ' +
        'AND PCN_PERIODEPY=-1 ' +
        'AND PCN_DATEVALIDITE>="' + UsDateTime(DebPeriodeN2) + '" ' +
        'AND PCN_DATEVALIDITE<="' + UsDateTime(FinPeriodeN2) + '" ' +
        'AND NOT EXISTS (SELECT PCN_TYPEMVT FROM ABSENCESALARIE ABS WHERE ABSENCESALARIE.PCN_SALARIE=ABS.PCN_SALARIE ' +
        'AND PCN_TYPECONGE="SLD" AND PCN_GENERECLOTURE="-" AND PCN_TYPEMVT="CPA")');
      MoveCurProgressForm('Mise à jour en cours...');
      RecalculePeriode(etab, FinPeriodeN1);
      FiniMoveProgressForm;
      CommitTrans;
      Result := True;
    except
      Rollback;
      Result := False;
      PGIBox('Une erreur est survenue lors de la "declôture" congés payés de l''établissement "' + EtabLibelle + '"', '"Declôture" période CP');
    end;
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui clôture les mvts acquis
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure ClotureMvtAcquis(T_ListeMvt: Tob; salarie: string; reliquat: integer; DatedebutperiodeP, DatefinperiodeP: TDateTime);
var
  TC: Tob;
  Typeconge: string;
  DateValidite: TDateTime;
begin
  { traitement du Salarie }
  TC := T_listeMvt.findfirst(['PCN_SALARIE'], [Salarie], True);
  while TC <> nil do
  begin //B10
    TypeConge := TC.getvalue('PCN_TYPECONGE');
    Datevalidite := TC.getvalue('PCN_DATEVALIDITE');
//PT128    if ((((typeConge = 'ACQ') or (typeConge = 'ACA') or (typeConge = 'ACS') or
    if ((((typeConge = 'ACQ') or (typeConge = 'ACA') or (typeConge = 'ACS') or (typeConge = 'ACF') or
      (typeConge = 'REL') or (typeConge = 'ARR') or (typeConge = 'REP') or
      (typeConge = 'AJU') or (typeConge = 'AJP')))) then
    begin //B19
// d PT128
      if  (typeConge = 'ACF') then
      begin
        if (TC.getvalue('PCN_PERIODECP') = 1) then
        begin
          if reliquat = 2 then
          begin
            TC.PutValue('PCN_CODETAPE', 'S');
            TC.putValue('PCN_DATESOLDE', DatefinperiodeP);
          end
          else
            TC.PutValue('PCN_CODETAPE', 'C');
        end;
      end
      else
// f PT128
      if ((Datevalidite <= DatefinperiodeP) and (Datevalidite >= DatedebutperiodeP)) then
      begin //B20
        if reliquat = 2 then
        begin //B22
          TC.PutValue('PCN_CODETAPE', 'S');
          TC.putValue('PCN_DATESOLDE', DatefinperiodeP);
        end //E22
        else
          TC.PutValue('PCN_CODETAPE', 'C');
      end; //E20
    end; //E19
    TC.UpdateDB(false);
    TC := T_listeMvt.findnext(['PCN_SALARIE'], [Salarie], True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : procédure qui génère les mouvements de clotûre
Mots clefs ... : PAIE,CP
*****************************************************************}

procedure GenereMvtCloture(T_MVTINS, TS: Tob; NoOrdre, periode: integer; TypeConges, sens, libelle, codeetape: string; datevalidite, AutreDate: Tdatetime; nbjour, nbmois, base: double);
var
  T1: TOB;
begin
  T1 := TOB.Create('ABSENCESALARIE', T_MVTINS, -1);
  InitialiseTobAbsenceSalarie(T1);
  T1.PutValue('PCN_ETABLISSEMENT', TS.getvalue('PSA_ETABLISSEMENT'));
  T1.Putvalue('PCN_TRAVAILN1', TS.getvalue('PSA_TRAVAILN1'));
  T1.Putvalue('PCN_TRAVAILN2', TS.getvalue('PSA_TRAVAILN2'));
  T1.Putvalue('PCN_TRAVAILN3', TS.getvalue('PSA_TRAVAILN3'));
  T1.Putvalue('PCN_TRAVAILN4', TS.getvalue('PSA_TRAVAILN4'));
  T1.Putvalue('PCN_CODESTAT', TS.getvalue('PSA_CODESTAT'));
  T1.Putvalue('PCN_CONFIDENTIEL', TS.getvalue('PSA_CONFIDENTIEL'));
  T1.PutValue('PCN_SALARIE', TS.getvalue('PSA_SALARIE'));
  T1.PutValue('PCN_ORDRE', Noordre);
  T1.PutValue('PCN_TYPECONGE', TypeConges);
  T1.PutValue('PCN_SENSABS', Sens);
  T1.PutValue('PCN_GENERECLOTURE', 'X');
  T1.PutValue('PCN_LIBELLE', libelle);
  T1.PutValue('PCN_DATEMODIF', Date);
  T1.PutValue('PCN_DATEVALIDITE', DateValidite);
  T1.PutValue('PCN_PERIODECP', periode);
  T1.PutValue('PCN_DATEDEBUT', AutreDate);
  if TypeConges = 'SLD' then AutreDate := idate1900;
  T1.PutValue('PCN_DATEFIN', AutreDate);
  T1.PutValue('PCN_JOURS', Arrondi(nbjour, DCP));
  T1.Putvalue('PCN_NBREMOIS', Arrondi(nbmois, DCP));
  T1.PutValue('PCN_BASE', Arrondi(base, DCP));
  T1.putValue('PCN_CODETAPE', codeetape);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui permet le calculer les soldes n-1 et n pour
Suite ........ : edition sur bulletin
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
procedure RechercheCumulCP(DD, DF: tdatetime; Salarie, Etab: string; var PN, AN, RN, BN, PN1, AN1, RN1, BN1: double; var DDP, DFP, DDPN1, DFPN1: tdatetime; var EditbulCP: string);
var
  DTClot, DDebut, Dfin, DSolde: TDateTime;
  Q: TQuery;
  Periodei: Integer;
  Periode, PeriodeN1, st: string;
  MB: Double;
begin
  EditbulCP := '';
  DTClot := Idate1900;
  { PT67 Recharge paramétrage édition compteurs }
  Q := Opensql('SELECT ETB_CONGESPAYES,ETB_DATECLOTURECPN,ETB_EDITBULCP,' + //PT57 Ajout champ ETB_CONGESPAYES
    'PSA_TYPEDITBULCP,PSA_EDITBULCP,PSA_CONGESPAYES ' + { PT74-1 }
    'FROM SALARIES ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    'WHERE PSA_SALARIE="' + Salarie + '"', True); //PT80 Suppression de WHERE ETB_ETABLISSEMENT="'+Etab+'"
  if not Q.eof then
  begin
    if Q.findfield('ETB_CONGESPAYES').AsString = '-' then
    begin
      Ferme(Q);
      exit;
    end; //PT57 Ajout contrôle
    if Q.findfield('PSA_CONGESPAYES').AsString = '-' then
    begin
      Ferme(Q);
      exit;
    end; { PT74-1 }
    DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
    if Q.FindField('PSA_TYPEDITBULCP').AsString = 'ETB' then EditbulCP := Q.FindField('ETB_EDITBULCP').AsString
    else EditbulCP := Q.FindField('PSA_EDITBULCP').AsString;

  end;
  Ferme(Q);
  { DEB PT-27-3 On récupère les mvts d'acquis posterieur au dernier SLD }
  { Pour éditer sur le bulletin que les mvts du dernier contrat}
  st := 'SELECT PCN_DATEVALIDITE FROM ABSENCESALARIE WHERE PCN_SALARIE="' + Salarie + '" ' +
    'AND PCN_DATEVALIDITE<"' + Usdatetime(DF) + '" AND PCN_TYPECONGE="SLD" AND PCN_TYPEMVT="CPA" ' + //PT32 AJout PCN_TYPEMVT="CPA"
    'AND PCN_GENERECLOTURE="-" ' + //PT65 On tient compte que des SLD de cloture
    'ORDER BY PCN_DATEVALIDITE DESC';
  Q := OpenSql(st, TRUE);
  if (not Q.eof) then
    DSolde := Q.FindField('PCN_DATEVALIDITE').AsDateTime
  else
    DSolde := iDate1900;
  Ferme(Q);
  {FIN PT-27-3}
  DDebut := 0;
  Dfin := 0;
  Periodei := CalculPeriode(DtClot, DF);
  Periode := inttostr(Periodei);
  PeriodeN1 := Inttostr(Periodei + 1);
  RechercheDate(DDebut, Dfin, DtClot, Periode);
  AffichelibelleAcqPri(Periode, Salarie, DSolde, DF, PN, AN, RN, BN, MB, False, True); //PT-27-3 Ajout DSolde
  DDP := DDebut;
  DFP := Dfin;
  DDebut := 0;
  Dfin := 0;
  RechercheDate(DDebut, Dfin, DtClot, PeriodeN1);
  AffichelibelleAcqPri(PeriodeN1, Salarie, DSolde, DF, PN1, AN1, RN1, BN1, MB, True, False); //PT-27-3 Ajout DSolde
  DDPN1 := DDebut;
  DFPN1 := Dfin;
end; *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : fonction qui rend la période d'un mouvement en fonction de
Suite ........ : la date de fin période CP
Mots clefs ... : PAIE;CP
*****************************************************************}

(* PT89 function CalculPeriode(DTClot, DTValidite: TDatetime): integer;
var
  Dtdeb, DtFin, DtFinS: TDATETIME;
  aa, mm, jj: word;
  i: integer;
begin
  result := -1;
  if DTClot <= idate1900 then exit; //PT53
  Decodedate(DTclot, aa, mm, jj);
  DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1; { PT71 }
  DtFin := DtClot;
  DtFinS := PGEncodeDateBissextile(AA + 1, MM, JJ); { PT71 }
  if Dtvalidite > Dtfins then
  begin
    result := -9;
    exit;
  end;
  if DtValidite > DtClot then exit;
  result := 0;
  i := 0;
  while not ((DTValidite >= DtDeb) and (DTValidite <= DtFin)) do
  begin
    i := i + 1;
    if i > 50 then exit; // pour ne pas boucler au cas où....
    result := result + 1;
    DtFin := DtDeb - 1;
    Decodedate(DTFin, aa, mm, jj);
    DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1; { PT71 }
  end;
end;                  *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 28/01/2003
Modifié le ... : 03/02/2003
Description .. : renvoie le n° de période dans lequel est inclue la
Suite ........ : datevalidite
Suite ........ :  DTcloturePi : datecloture courante établissement
Suite ........ :  si résultat : 0 => période courante
Suite ........ :                   1 => période - 1
Suite ........ :                   2 => période - 2
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
function RendPeriode(var DTcloturePi, DTDebutPi: tdatetime; DTCloture, datevalidite: tdatetime): integer;
var
  i: integer;
  aa, mm, jj: word;
begin
  result := -1;
  DTcloturePi := DTCloture;
  decodedate(DTCloture, aa, mm, jj);
  DTDebutPi := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1; { PT71 }
  i := 0;
  while i < 10 do { au delà de p-10 , il y a sûrement une erreur ! }
  begin
    if ((datevalidite >= DTDebutPi) and (datevalidite <= DTcloturePi)) then
    begin
      result := i;
      exit;
    end;
    i := i + 1;
    DTcloturePi := DTDebutPi - 1;
    decodedate(DTcloturePi, aa, mm, jj);
    DTDebutPi := PGEncodeDateBissextile(aa - 1, MM, JJ) + 1; { PT71 }
  end;
  if result = -1 then
    { on essaie dans l'autre sens pour essayer de récupérer les bons intervalles }
  begin
    DTcloturePi := DTCloture;
    decodedate(DTCloture, aa, mm, jj);
    DTDebutPi := PGEncodeDateBissextile(aa, mm, jj) + 1; { PT71 }
    DTcloturePi := PGEncodeDateBissextile(aa + 1, mm, jj);
    i := 0;
    while i < 10 do { au delà de p+10 , ca va , on arrête ! }
    begin
      if (datevalidite >= DTDebutPi) and (datevalidite <= DTcloturePi) then exit;
      i := i + 1;
      DTdebutPi := DTcloturePi + 1;
      decodedate(DTcloturePi, aa, mm, jj);
      DTcloturePi := PGEncodeDateBissextile(aa + 1, mm, jj); { PT71 }
    end;
  end;
end; *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui en fonction de la période renvoie les dates
Suite ........ : de session de paie
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
procedure RechercheDate(var DDebut, Dfin, Datecloture: tdatetime; Periode: string);
var
  Noperiode, i: integer;
  aa, mm, jj: word;
begin
  DDebut := 0;
  DFin := 0;
  if Isnumeric(Periode) then NoPeriode := strtoint(Periode)
  else exit;
  DFin := Datecloture;
  Decodedate(DFin, aa, mm, jj);
  DDebut := PGEncodeDateBissextile(aa - 1, mm, jj); { PT71 }
  DDebut := DDebut + 1;
  i := 0;
  while i < NoPeriode do
  begin
    DFin := DDebut - 1;
    Decodedate(DFin, aa, mm, jj);
    DDebut := PGEncodeDateBissextile(aa - 1, mm, jj); { PT71 }
    DDebut := DDebut + 1;
    i := i + 1;
  end;
end; *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Calcul la durée d'absence, du temps de travail  en fonction
Suite ........ : du calendrier salarié
Mots clefs ... : PAIE;CP;ABSENCE
*****************************************************************}

procedure CalculDuree(DateDebut, DateFin: TDatetime; Salarie, Etablissement, TypeMvt: string; var Nb_J, Nb_H: double; Nodjm: integer = 0; Nodjp: integer = 1);
var
  { PT60 Mise en commentaire
  Tob_Cal,TCal                        : Tob;
   NbJourTrav,i,NDay,No                : integer;
   Duree                               : double;
   RH1,RH2,PRH2,st                     : string;
   DDay                                : TDateTime;
   TabDay                              : array [1..7] of boolean ;
   TabHour,TabHourPlage1,TabHourPlage2 : array [1..7] of double; //PT-15
   FJourFerie                          : TJourFerie;
   MvtCP,MvtAbS                        : boolean; }
  calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure: string;
  Tob_Semaine, Tob_Standard, Tob_JourFerie, Tob_absence: Tob;
  DateEntree, DateSortie, DateDebAbs, DateFinAbs: TDateTime;
  CptJour, CptHeure, Nb_JSal: Double;
  SSDecompte, JourNonTrav, JourNonTravStd, AM, PM, AMDeb, AMFin, PMDeb, PMFin, AMStd, PMStd: boolean;
begin
  { DEB PT60 NOUVELLE FONCTION DE CALCUL DE TEMPS
     Principe de calcul du decompte CP :
  1- On positionne la date de fin à la dernière demi-journée travaillé par le salarié
     On positionne la date de debut à la première demi journée travaillé par l'entreprise
  2- On recherche la demi journée de reprise de travail pour l'entreprise et le salarié
  3- On comptabilise le temps de travail sur le calendrier entreprise
   }
  Nb_J := 0; { PT93 }
  Nb_H := 0;
  SSDecompte := False;
  //Init Des Tobs
  ChargeTobCalendrier(DateDebut, DateFin, Salarie, FALSE, True, False, Tob_Semaine, Tob_Standard, Tob_JourFerie, Tob_absence, Etablissement, Calend, StandCalend, EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure, DateEntree, DateSortie); { PT97 }
  if Tob_Semaine = nil then exit;
  if Tob_Semaine.Somme('ACA_DUREE', [''], [''], False) = 0 then Exit; { PT93 }
  DateDebAbs := DateDebut;
  DateFinAbs := DateFin;
  CptJour := 0;
  CptHeure := 0;
  JourNonTrav := True;
  {Pour calcul congés payés pris on positionne la date de fin sur le dernier jour
  travaillé afin de tester le lendemain du dernier jour travaillé par la suite}
  if TypeMvt = 'PRI' then
  begin
    SSDecompte := (RechDom('PGSALDECOMPTE', salarie, false) = 'X');
    //1- Repositionnement des dates de debut et fin
    while (JourNonTrav = True) and (DateDebAbs <= DateFinAbs) do
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_jSal, Nb_h, AMDeb, PMDeb);
      if (JourNonTrav = True) then DateDebAbs := DateDebAbs + 1;
    end;
    if DateDebAbs > DateFinAbs then DateDebAbs := DateDebut;
    JourNonTrav := True;
    while (JourNonTrav = True) and (DateFinAbs >= DateDebAbs) do
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateFinAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_jSal, Nb_h, AMFin, PMFin);
      if (JourNonTrav = True) { AND (Nodjp=1) PT75-1 } then DateFinAbs := DateFinAbs - 1;
    end;
    if DateFinAbs < DateDebAbs then DateFinAbs := DateFin;
    //2- Recherche de la demi journée de reprise du standard
    JourNonTrav := True;
    while (JourNonTrav = True) and (Nodjp = 1) and (Assigned(Tob_Standard)) do
    begin
      if (DateSortie < DateFinAbs) and (DateSortie > idate1900) then Break; //PT69 PT100
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateFinAbs + 1, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AMFin, PMFin);
      JourNonTravStd := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateFinAbs + 1, True, Nb_j, Nb_h, AMStd, PMStd);
      {Un jour non travaillé par l'entreprise est soit un jour de repos hebdo, soit un jour férié, soit un jour ouvrable
      Dans le cas d'un jour de repos hebdo , on saute la journée
      Dans le cas d'un jour férié, on vérifie si c'est un jour de travail du calendrier,( Nb_j = -1 pour un jour ferié)
      Tant Que Jour pas travaillé pour salarié et jour travaillé pour l'entreprise}
      if ((JourNonTrav = True) and (JourNonTravStd = False))
        or ((JourNonTrav = True) and (JourNonTravStd = True) and (Nb_j <> -1)) then
      begin
        DateFinAbs := DateFinAbs + 1;
        Nodjp := 1;
      end
      else
        //Cas d'un jour férié correspondant à un jour travaillé dans l'entreprise
        if (JourNonTrav = True) and (JourNonTravStd = True) and (Nb_j = -1) then
        begin
          JourNonTravStd := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateFinAbs + 1, False, Nb_j, Nb_h, AMStd, PMStd);
          if (JourNonTravStd = False) then
          begin
            DateFinAbs := DateFinAbs + 1;
            Nodjp := 1;
          end
          else
            JourNonTrav := False;
        end
        else
          if (JourNonTravStd = True) then JourNonTrav := False;
    end;
  end;
  //PT69 Repositionnement de la date de début et de fin en fonction de la date d'entrée et de sortie
  if DateEntree > DateDebAbs then DateDebAbs := DateEntree;
  if DateEntree > DateFinAbs then
  begin
    Nb_j := 0;
    Nb_h := 0;
    exit;
  end; { PT69 16/04/04 }
  if (DateSortie < DateFinAbs) and (DateSortie > idate1900) then DateFinAbs := DateSortie; { PT100 }
  //3- On comptabilise le temps de travail sur le calendrier entreprise si type Cp acquis = IdemEtab
  repeat
    if (TypeMvt = 'PRI') and (Nb_JSal > 0) then
    begin
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM);
      //Jour non travaillé pas le salarié
      if ((JourNonTrav = True) or (NB_J < 1)) and (SSDecompte = False) and (Assigned(Tob_Standard)) then //si travail demi journée salarié { PT96 }
      begin
        JourNonTrav := IfJourNonTravailStd(Tob_Standard, Tob_JourFerie, DateDebAbs, True, Nb_j, Nb_h, AM, PM);
        //Dans le cas d'un jour non travaillé mais ouvrable on comptabilise
        if (JourNonTrav = True) and (Nb_j = 0) then
          if (IntToStr(DayOfWeek(DateDebAbs)) <> Repos1) and (IntToStr(DayOfWeek(DateDebAbs)) <> Repos2) then
            CptJour := CptJour + 1;
        //Dans le cas d'un jour travaillé par l'entreprise mais pas par le salarié, on décompte si case non cochée
        if (JourNonTrav = False) and (SSDecompte = True) then JourNonTrav := True;
      end;
    end
    else //Ou on comptabilise sur le calendrier salarié pour les autres motifs d'absence, le calcul du temps de travail..
      JourNonTrav := IfJourNonTravaille(Tob_Semaine, Tob_Standard, Tob_JourFerie, DateDebAbs, DateEntree, DateSortie, NbJrTravEtab, Repos1, Repos2, False, True, Nb_j, Nb_h, AM, PM); { PT77 }
    //une fois la période de calcul des jours calculé on peut décompte le temps consommé
    if (JourNonTrav = False) then
    begin
      //Dans le cas d'une demi-journée sur une journée complète travaillé
      if (DateDebAbs = DateDebut) and (Nodjm = 1) and (AM = True) and (PM = True) then
      begin
        Nb_j := Nb_j / 2;
        Nb_H := Nb_H / 2
      end;
      if (DateDebAbs = DateFin) and (Nodjp = 0) and (AM = True) and (PM = True) then
      begin
        Nb_j := Nb_j / 2;
        Nb_H := Nb_H / 2
      end;
      //Dans le cas d'une demi-journée sur une journée ouvré ou ouvrable travaillé à demi temps
      if (SSDecompte = False) and (Nb_j = 0.5) and (((AM = True) and (PM = FALSE)) or ((AM = False) and (PM = True))) then Nb_j := 1; //PT68-1 PT96
      CptJour := CptJour + Nb_j;
      CptHeure := CptHeure + Nb_H;
    end;
    DateDebAbs := DateDebAbs + 1;
  until (DateDebAbs > DateFinAbs);


  FreeAndNil(Tob_Standard);
  FreeAndNil(Tob_Semaine);
  FreeAndNil(Tob_JourFerie);
  {if Tob_Standard<>nil then  Begin Tob_Standard.free; Tob_Standard:=nil; end;
  if Tob_Semaine<>nil then   Begin Tob_Semaine.free; Tob_semaine:=nil; end;
  if Tob_JourFerie<>nil then Begin Tob_JourFerie.free; Tob_JourFerie:=nil; end; }
  Nb_J := CptJour;
  Nb_H := CptHeure;

  { FIN PT60 }

  (** PT60 Mise en commentaire de l'ancien mode de calcul
    { Rq : Nodjm, Nodjp : no de 1/2 journée 0: matin, 1: après midi }
    MvtCP := false ; MvtAbS := false;
    if CalcCalend = true then
      MvtAbs := true { ds ce cas, on fait un calcul d'intervalles de type absence }
    else
       begin
       if RechDom('PGMVTCONGES',TypeMvt,FALSE) <> '' then
         MvtCP := true
       else
         MvtAbS := True;
       end;
    Nb_H := 0;
    FJourFerie:= TJourFerie.Create;
    for i:= 1 to 7 do begin TabDay[i] := False; TabHour[i] := 0; end;

    st := 'SELECT ETB_1ERREPOSH,ETB_2EMEREPOSH,ETB_NBJOUTRAV FROM ETABCOMPL '+
          'WHERE ETB_ETABLISSEMENT="'+Etablissement+'"';
    Q:=OpenSQL(st,TRUE) ;
    if Not Q.EOF then
       Begin
       NbJourTrav := Q.FindField('ETB_NBJOUTRAV').AsInteger;
       { repos  hebdomadaire niveau établissement  }
       RH1 := Copy(Q.FindField('ETB_1ERREPOSH').AsString,1,1);
       RH2 := Copy(Q.FindField('ETB_2EMEREPOSH').AsString,1,1);
       End
    else
       NbJourTrav := 0;
    Ferme(Q);
    // if (NbJourTrav = 0) and (CalcCalend=False) then exit;  //PT-12 Mise en commentaire
    { Recherche des jours habituellement travaillés pour ce salarié }
    Tob_Cal:=ChargeCalendrierSalarie(Etablissement,Salarie,'','','');//PT51 ajout paramètre

    { En delphi, le 1er jour de la semaine est le Dimanche, et le dernier le Samedi }
    { Ds la gestion du calendrier, le 1er est le lundi, le dernier, le Dimanche }
    if Tob_Cal<>nil then
       begin
       TCal:=Tob_Cal.FindFirst([''],[''],False);
       while Tcal<>nil do   //Not Q.EOF   PT-14
          Begin
          No    := TCal.GetValue('ACA_JOUR');//Q.FindField('ACA_JOUR').AsInteger;
          if No < 7 then No := No+1 else No := 7;
          Duree := TCal.GetValue('ACA_DUREE');//Q.FindField('ACA_DUREE').Asfloat;
          if (Duree <> 0) AND (TCal.GetValue('ACA_JOUR')<>0)then  //Attention pour jour non travaillé No=0
             begin
             TabDay[No]:=True;
             TabHour[No] := Duree;
             TabHourPlage1[No]:=Arrondi(TCal.getValue('ACA_HEUREFIN1')-TCal.getValue('ACA_HEUREDEB1'),2);//DEB PT-15
             TabHourPlage2[No]:=Arrondi(TCal.getValue('ACA_HEUREFIN2')-TCal.getValue('ACA_HEUREDEB2'),2);//FIN PT-15
             end;
          TCal:=Tob_Cal.FindNext([''],[''],False);
          end;    // while
       end;
    if MvtCp then
       begin
       if (RH1 <> '') AND (RH1 <> '0') then  //PT-14 Attention Rh1 peut être égale à 0
          begin
          TabDay[strtoint(RH1)]:= false;
          TabHour[strtoint(RH1)]:= 0;
          end;
       if (RH2 <> '') AND (RH2 <> '0') then  //PT-14 Attention Rh2 peut être égale à 0
          begin
          TabDay[strtoint(RH2)]:= false;
          TabHour[strtoint(RH2)]:= 0;
          end;
       end;
    DDay := Datedebut; Nb_j := 0; Nb_H := 0;
    { En ouvrable, on calcule un pseudo RH2 = jour qui précède RH1 }
    PRH2:='';
    if ((NbJourtrav =6) and (RH2=''))then
       begin
       if strtoint(RH1) = 1 then PRH2 := '7'
       else                      PRH2 := inttostr(strtoint(RH1) - 1);
       end;
    while DDay <= datefin do
       begin
       { DEB PT-14 Test si jour ne correspond pas à un jour spécifique du calendrier }
       TCal:=nil;
       if Tob_Cal <>nil then
          Begin
          TCal:=Tob_Cal.FindFirst(['ACA_DATE','ACA_FERIETRAVAIL'],[DDay,'-'],False);
          if Tcal<>nil then
             If TCal.GetValue('ACA_DUREE')>0 then
                begin
                Nb_j := Nb_J + 1;
                Nb_H := NB_h + TCal.GetValue('ACA_DUREE');
                End;
          End;
       {FIN PT-14 }
       if (not FJourFerie.TestJourFerie(DDay)) and (TCal=nil) then   ////PT-14 Ajout Cond. TCal=Nil
          begin
          NDay := Dayofweek(DDay);
          if (TabDay[NDay]) then
             begin
             if ((DDay = datedebut) and (datedebut=datefin) and (nodjm=nodjp)) then
                begin
                nb_j :=1/2;
                if nodjm=0 then Nb_H:=TabHourPlage1[NDay] else Nb_H:=TabHourPlage2[NDay];//DEB PT-15
                //Nb_H := TabHour[NDay]/2; FIN PT-15 misen en commentaire
                end
             else
                if (((DDay = datedebut) and (Nodjm = 1)) or ((DDay = dateFin) and (Nodjp = 0))) then
                   begin
                   nb_j := nb_j + 1/2;
                   if (DDay = datedebut) and (Nodjm = 1) then Nb_H:=NB_h+TabHourPlage2[NDay];//DEB PT-15
                   if (DDay = dateFin) and   (Nodjp = 0) then Nb_H:=NB_h+TabHourPlage1[NDay];
                   //Nb_H := NB_h + TabHour[NDay]/2;  //FIN PT-15 misen en commentaire
                   end
                else
                   begin
                   Nb_j := Nb_J + 1;
                   Nb_H := NB_h + TabHour[NDay];
                   end;
             end
          else
             if MvtCP then
                if  ((inttostr(NDay)= PRH2) and (NbJourTrav = 6)and(Nb_J >0)) then
                   begin
                   if ((DDay = datedebut) and (datedebut=datefin) and (nodjm=nodjp)) then
                      begin
                      nb_j :=1/2;
                      Nb_H := TabHour[NDay]/2;
                      end
                   else
                      if (((DDay=datedebut) and (Nodjm=1)) or ((DDay=dateFin) and (Nodjp=0))) then
                         begin
                         nb_j := nb_j + 1/2;
                         Nb_H := NB_h + TabHour[NDay]/2;
                         end
                      else
                         begin
                         Nb_j := Nb_J + 1;
                         Nb_H := NB_h + TabHour[NDay];
                         end;
                   end
                else
                   if ((inttostr(NDay)<>RH1)and(inttostr(NDay)<>RH2)
                       and (Nb_j > 0) and (not TabDay[Nday] )) then
                      begin
                      if ((DDay = datedebut) and (datedebut=datefin) and (nodjm=nodjp)) then
                         begin
                         nb_j :=1/2;
                         Nb_H := TabHour[NDay]/2;
                         end
                      else
                         if (((DDay = datedebut) and (Nodjm = 1)) or
                             ((DDay = dateFin)   and (Nodjp = 0)) ) then
                            begin
                            nb_j := nb_j + 1/2;
                            Nb_H := NB_h + TabHour[NDay]/2;
                            end
                         else
                            begin
                            Nb_j := Nb_J + 1;
                            Nb_H := NB_h + TabHour[NDay];
                            end;
                         end;
          end; //end of if
    { Pour les salariés à calendrier spécifique qui ne travaille pas un jour de la semaine }
    { mais travaillé par l'établissement, on test si le jour suivant }
    { le dernier jour de prise de congé correspond à son jour de repos spécifique }
    DDay := DDay + 1;
    end;   // end of while
    { si on est en ouvrable et que le lendemain de la prise de congé est la veille du }
    { repos hebdomadaire, alors on ajoute 1 jour }
    if ((NbJourtrav =6)and  (PRH2 = inttostr(Dayofweek(Datefin+1)))and (MvtCp)) then
       Nb_j := NB_j + 1;
    Nb_j :=Arrondi(Nb_j,DCP);
    Nb_H :=Arrondi(Nb_h,DCP);
    FJourFerie.free;
    if Tob_Cal<>nil then Tob_Cal.free;   **)
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : renvoie le nombre de jours et d'heures travaillés dans une
Suite ........ : session de paie
Mots clefs ... : PAIE;CP
*****************************************************************}
{PT-25 Modification de la function en procedure pour recup à la fois Jours et Heures
théoriques travaillés
function  CalculJoursTheorique(Tob_sal:tob;datedebut,datefin:tdatetime):double;}

procedure CalculJoursTheorique(Tob_sal: tob; datedebut, datefin: tdatetime; var Varj, Varh: double); //PT-25
var NbJrTravEtab, Repos1, Repos2, E1, Calend, StandCalend, EtabStandCalend, JourHeure: string;
  TETABSAL, T1: Tob;
  Date1, Date2, D1, D2: TdateTime;
  Nb_j, Nb_h: Double;
  AM, PM: Boolean;
begin
  Varh := 0;
  VarJ := 0;
 { PT97 procédure appellé dans le moteur de la paie,
        calcul du jours et heures à partir des tobs globales de la paie,
  CalculDuree(DateDebut, datefin, Tob_sal.getvalue('PSA_SALARIE'), Tob_sal.getvalue('PSA_ETABLISSEMENT'), 'ABS', VarJ, VarH); }
  Date1 := DateDebut;
  Date2 := Datefin;
  E1 := TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT); { PT109 }
  D1 := TOB_SALARIE.GetValeur(iPSA_DATEENTREE); { PT109 }
  D2 := TOB_SALARIE.GetValeur(iPSA_DATESORTIE); { PT109 }
  if (Date1 > Date2) then exit;
  if (TOB_Salarie.GetValeur(iPSA_CALENDRIER) <> '') then
  begin
    TETABSAL := TOB_Etablissement.findfirst(['ETB_ETABLISSEMENT'], [TOB_SALARIE.GetValeur(iPSA_ETABLISSEMENT)], True);
    NbJrTravEtab := IntToStr(TETABSAL.GetValue('ETB_NBJOUTRAV'));
    Repos1 := TETABSAL.GetValue('ETB_1ERREPOSH'); if Repos1 <> '' then repos1 := Copy(repos1, 1, 1);
    Repos2 := TETABSAL.GetValue('ETB_2EMEREPOSH'); if Repos2 <> '' then repos2 := Copy(repos2, 1, 1);
     { DEB PT109 }
    if not Assigned(GblTob_Semaine) then
    begin
      ChargeTobCalendrier(DateDebut, Datefin, TOB_Salarie.GetValeur(iPSA_SALARIE),
        FALSE, True, False, GblTob_Semaine, GblTob_Standard, GblTob_JourFerie, T1,
        E1, Calend, StandCalend,
        EtabStandCalend, NbJrTravEtab, Repos1, Repos2, JourHeure, D1, D2);
    end;
     { FIN PT109 }
    while (Date1 <= Date2) do
    begin
      if IfJourNonTravaille(GblTob_Semaine, GblTob_Standard, GblTob_JourFerie,
        Date1, D1, D2, NbJrTravEtab, Repos1, Repos2, { PT109 }
        False, True, Nb_j, Nb_h, AM, PM) = False then
      begin
        VarJ := VarJ + Nb_j;
        Varh := Varh + Nb_h;
      end;
      Date1 := Date1 + 1;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui calcul le solde des CP pour une période
Suite ........ : donnée
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
procedure AffichelibelleAcqPri(Periode, Salarie: string; DateDeb, Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double; edit, SansCloture: boolean);
var
st: string;
valo: Double;
Q: TQuery;
TAbs: Tob;
begin
Pris:= 0;
Acquis:= 0;
Restants:= 0;
Base:= 0;
Moisbase:= 0;
if (Salarie = '') then
   exit;
{affichage du libellé correspondant à la période sélectionnée}
Tabs:= Tob.create ('ABSENCESALARIE', nil, -1);
st:= 'SELECT *'+
     ' FROM ABSENCESALARIE WHERE'+
     ' PCN_SALARIE="'+Salarie+'" AND'+
     ' PCN_TYPEMVT="CPA" AND'+
     ' PCN_PERIODECP='+Periode+' AND'+
     ' PCN_CODERGRPT<>-1 AND'+
     ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'"';
if (edit) and (StrToInt(periode) > 0) then
   begin
   St:= 'SELECT *'+
        ' FROM ABSENCESALARIE WHERE'+
        ' PCN_SALARIE="'+Salarie+'" AND'+
        ' PCN_DATEVALIDITE<="'+USDateTime (DateFin)+'" AND'+
        ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'" AND'+
        ' PCN_TYPEMVT="CPA" AND'+
        ' PCN_PERIODECP='+Periode+' AND'+
        ' PCN_CODERGRPT<>-1 AND'+
        ' ((PCN_GENERECLOTURE="-") OR'+
        ' (PCN_GENERECLOTURE="X" AND'+
        ' ((PCN_SENSABS="+" AND'+
        ' PCN_TYPECONGE<>"SLD") OR'+
        ' (PCN_TYPECONGE="SLD"))))';
{J'exclue les mvts générés par la cloture de période précédente}
   end;

if SansCloture = True then
   begin
   st:= 'SELECT *'+
        ' FROM ABSENCESALARIE WHERE'+
        ' PCN_SALARIE="'+Salarie+'" AND'+
        ' PCN_DATEVALIDITE<="'+USDateTime (DateFin)+'" AND'+
        ' PCN_DATEVALIDITE>"'+USDateTime (DateDeb)+'" AND'+
        ' PCN_TYPEMVT="CPA" AND'+
        ' PCN_PERIODECP='+Periode+' AND'+
        ' PCN_CODERGRPT<>-1 AND'+
        ' (PCN_GENERECLOTURE="-")';
   end;
{Flux optimisé
  Q := Opensql(st, True);
  if not Q.eof then
    Tabs.loaddetaildb('ABSENCESALARIE', '', '', Q, false, false);
  Ferme(Q);
  if Tabs = nil then exit;
}
Tabs.loaddetaildbFromSQL ('ABSENCESALARIE', st);
if (Tabs.detail.count<=0) then
   begin
   FreeAndNil (Tabs);     //PT125
   exit;
   end;
Valo:= 0;
CompteCP (TAbs, Periode, Salarie, Datefin, Pris, Acquis, Restants, Base,
          Moisbase, Valo);
{PT125
if (Tabs <> nil) then
   Tabs.free;
}
FreeAndNil (Tabs);
//FIN PT125
Pris:= Arrondi (Pris, DCP);
Acquis:= Arrondi (Acquis, DCP);
Base:= Arrondi (Base, DCP);
Restants:= Arrondi (Restants, DCP);
end;  *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui balaye chaque mvt CP et alimente les
Suite ........ : compteurs CP
Mots clefs ... : PAIE;CP
*****************************************************************}
(*PT127
procedure CompteCp(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase, valo: double);
var
  Typeconge, Sens: string;
  Jours, MBase: double;
  T: Tob;
begin
  { datefin = 0 signifie qu'on n'est pas dans un bulletin }
  if datefin = 0 then datefin := 99999;
  Acquis := 0;
  Pris := 0;
  Base := 0;
  Restants := 0;
  MoisBase := 0;
  T := TCP.findfirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  while T <> nil do
  begin { PCN_CODERGRPT <> -1  }
    if T.Getvalue('PCN_CODERGRPT') <> '-1' then
    begin
      Typeconge := T.getvalue('PCN_TYPECONGE');
      Sens := T.getvalue('PCN_SENSABS');
      Jours := T.getvalue('PCN_JOURS');
      MBase := T.getvalue('PCN_BASE');
      if IsAcquis(TypeConge, Sens) then
        if Sens = '+' then
        begin
          Acquis := ACquis + Jours;
          MoisBase := MoisBase + T.getvalue('PCN_NBREMOIS');
          Restants := Restants + Jours;
          Base := Base + MBase;
          if (T.getvalue('PCN_NBREMOIS')) = 0 then
            Valo := Valo + 0
          else
            Valo := Valo + (arrondi((12 * MBase) / (T.getvalue('PCN_NBREMOIS') * 10), DCP) * Jours);
        end
        else
        begin
          MoisBase := MoisBase - T.getvalue('PCN_NBREMOIS');
          Base := Base - MBase;
          Acquis := ACquis - Jours;
          Restants := Restants - Jours;
        end; { on prend les CPA de la période car ils st forcément payés ! }
      if ((IsPris(TypeConge)) and ((T.getvalue('PCN_DATEFIN') > 10) and
        (T.getvalue('PCN_DATEFIN') <= datefin) and
        (T.getvalue('PCN_DATEPAIEMENT') > 10)) or (Typeconge = 'CPA')) then
        if Sens = '-' then
        begin
          Pris := Pris + Jours;
          Restants := Restants - Jours;
        end
        else
        begin
          Pris := Pris - Jours;
          Restants := Restants + Jours;
        end;
      if ((TypeConge = 'REL') and (sens = '-')) then
        Restants := Restants - Jours;
    end;
    T := TCP.findnext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  end; //while
end;  *)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : procedure qui balaye chaque mvt CP et alimente les
Suite ........ : compteurs CP payés
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure CompteCpAPAYES(TCP: tob; Periode, Salarie: string; Datefin: tdatetime; var Pris, Acquis, Restants, Base, Moisbase: double);
var
  Typeconge, Sens: string;
  Jours, MBase: double;
  T: Tob;
begin
  Acquis := 0;
  Pris := 0;
  Base := 0;
  Restants := 0;
  MoisBase := 0;
  T := TCP.findfirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  while T <> nil do
  begin
    Typeconge := T.getvalue('PCN_TYPECONGE');
    Sens := T.getvalue('PCN_SENSABS');
    Jours := T.getvalue('PCN_JOURS');
    MBase := T.getvalue('PCN_BASE');
    { DEB PT-17 }
    if (TypeConge = 'AJP') and (Sens = '-') then Restants := Restants - Jours;
    if (TypeConge = 'AJP') and (Sens = '+') then Restants := Restants + Jours;
    { FIN PT-17 }
    if not IsAcquis(TypeConge, Sens) then
    begin
      T := TCP.findnext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
      continue;
    end;
    if Sens = '+' then
    begin
      Acquis := ACquis + Jours;
      MoisBase := MoisBase + T.getvalue('PCN_NBREMOIS');
      Restants := Restants + Jours;
      Restants := Restants - T.getvalue('PCN_APAYES');
      Base := Base + MBase;
    end
    else
    begin
      MoisBase := MoisBase - T.getvalue('PCN_NBREMOIS');
      Base := Base - MBase;
      Acquis := ACquis - Jours;
      Restants := Restants - Jours;
    end;
    if ((TypeConge = 'REL') and (sens = '-')) then
      Restants := Restants - Jours;
    T := TCP.findnext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], false);
  end; //while
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Fonction presque idem isnumeric sans blanc ni '-'
Mots clefs ... : PAIE;
*****************************************************************}

function IsValeurNumeric(const S: string): boolean;
const
  Num = ['0'..'9', '.', ','];
var
  i: Integer;
begin
  IsValeurNumeric := FALSE;
  for i := 1 to Length(S) do
    if ((not (S[i] in Num)) and (Ord(s[i]) <> 160)) then exit;
  IsValeurNumeric := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui génère le solde de tout compte dans un
Suite ........ : bulletin
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure GenereSolde(Tob_Sal, Tob_etab, Tob_Pris, Tob_Acquis, Tob_Solde: tob; DateF, DateD: tdatetime; var next: integer; AcquisBulletin, PrisBul: double; Action: string);
var
  DateSortie, DTFinP, DTDebP: TDateTime;
  NbJAcq, Arr, NBJPayes, NBJAcqP, cpt: double;
  T: Tob;
  Periode: integer;
  salarie: string;
begin
  Salarie := Tob_Sal.GetValue('PSA_SALARIE');
  DateSortie := Tob_Sal.GetValue('PSA_DATESORTIE');
  if not ((DateSortie >= DateD) and (DateSortie <= Datef)) then exit;
  NbJAcq := 0;
  NBJPayes := 0;
  cpt := 0;
  Periode := RendPeriode(DTFinP, DTDebP, Tob_etab.getvalue('ETB_DATECLOTURECPN'), DateF);

  CalculAcquisPeriodeSolde(Tob_etab, Tob_Acquis, DTDebP, DTFinP, DateD, DateF, NBJAcqP, NbJAcq, NBJPayes, salarie); { PT105 }
  if TopJPAvantBull = false then
  begin
    JPAVANTBULL := NBJPayes - PrisBul;
    TopJPAVANTBULL := true;
  end;
  //mvi ??  if ASOLDERAVBULL=0 then
  ASOLDERAVBULL := NbJAcq - JPAVANTBULL - PrisBul;
  if JRSAARRONDIR = 0 then JRSAARRONDIR := NBJAcqP;
  NBJAcqP := NBJAcqP + AcquisBulletin;
  NBJAcq := NBJAcq + AcquisBulletin;
  Arr := 0; { PT111 SI CHR pas d'arrondi de solde }
  if VH_Paie.PGChr6Semaine = False then { PT111 }
  begin
     {DEB PT-11-2 //PT-24 On arrondi l'acquis de la periode N et non le total des 2 périodes (ASOLDERAVBULL) -JCPACQUISACONSOMME }
    Arr := Arrondi((1 + Int(JRSAARRONDIR + AcquisBulletin) - (JRSAARRONDIR + AcquisBulletin)), dcp); // mv 05-02
    if ((Arr > 0) and (Arr < 1)) then
    begin
      T := TOB.Create('ABSENCESALARIE', Tob_Acquis, -1);
      Next := Next + 1;
      CreeArrondi(Arr, Tob_Sal.getvalue('PSA_SALARIE'), tob_sal.getvalue('PSA_ETABLISSEMENT'), '-', Next, Periode, DateF, T); //mvi
    end
    else Arr := 0;
      { FIN PT-11-2 }
  end;
  if ACtion = 'C' then
  begin
    { PT-29 Mise en commentaire }
    {if ((ASOLDERAVBULL+AcquisBulletin+arr-JCPACQUISACONSOMME) > 0) then //PT-11-1   ajout JCPACQUISACONSOMME}
    if Arr = 0 then Next := Next + 1; { PT-21 Le num ordre doit être incrementé.. }
    T := TOB.Create('ABSENCESALARIE', Tob_Solde, -1);
    InitialiseTobAbsenceSalarie(T);
    T.PutValue('PCN_SALARIE', Tob_Sal.getvalue('PSA_SALARIE'));
    T.PutValue('PCN_ORDRE', Next);
    T.PutValue('PCN_TYPECONGE', 'SLD');
    T.PutValue('PCN_SENSABS', '-');
    T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + Arr + AcquisBulletin - JCPACQUISACONSOMME, DCP)) + 'j au ' + datetostr(dateF)); //PT-11-1
    T.PutValue('PCN_DATEMODIF', Date);
    T.PutValue('PCN_DATEVALIDITE', Datef);
    T.PutValue('PCN_PERIODECP', 0);
    T.PutValue('PCN_DATEDEBUT', Datef);
    T.PutValue('PCN_DATEFIN', Datef);
    T.PutValue('PCN_JOURS', Arrondi(ASOLDERAVBULL + Arr + AcquisBulletin - JCPACQUISACONSOMME, DCP)); //PT-11-1
    T.putValue('PCN_CODETAPE', '...'); { PT-13 }
    T.PutValue('PCN_ETABLISSEMENT', Tob_Sal.getvalue('PSA_ETABLISSEMENT'));
    T.Putvalue('PCN_TRAVAILN1', Tob_Sal.getvalue('PSA_TRAVAILN1'));
    T.Putvalue('PCN_TRAVAILN2', Tob_Sal.getvalue('PSA_TRAVAILN2'));
    T.Putvalue('PCN_TRAVAILN3', Tob_Sal.getvalue('PSA_TRAVAILN3'));
    T.Putvalue('PCN_TRAVAILN4', Tob_Sal.getvalue('PSA_TRAVAILN4'));
    T.Putvalue('PCN_CODESTAT', Tob_Sal.getvalue('PSA_CODESTAT'));
    T.Putvalue('PCN_CONFIDENTIEL', Tob_Sal.getvalue('PSA_CONFIDENTIEL'));
    Next := Next + 1;
    { PT-29 }
  end;
  if Action = 'M' then
  begin
    T := Tob_solde.findfirst([''], [''], true);
    while t <> nil do
    begin
      if ((t.getvalue('PCN_CODERGRPT') = -1) or (Tob_solde.detail.count = 1)) then
      begin
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin + arr, DCP)) + 'j au ' + datetostr(dateF));
        T.PutValue('PCN_JOURS', Arrondi(ASOLDERAVBULL + AcquisBulletin + arr, DCP)); //PT-11-2
        T := Tob_solde.findnext([''], [''], true);
        continue;
      end;
      if ((t.getvalue('PCN_PERIODECP') <> 0)) then
      begin
        cpt := cpt + T.getValue('PCN_JOURS');
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin {mv  05-02+arr}, DCP)) + 'j au ' + datetostr(dateF));
      end;
      T := Tob_solde.findnext([''], [''], true);
    end;
    T := Tob_solde.findfirst(['PCN_PERIODEPY'], ['0'], true);
    while t <> nil do
    begin
      if t.getvalue('PCN_CODERGRPT') <> -1 then
      begin
        T.PutValue('PCN_LIBELLE', 'Solde ' + floattostr(Arrondi(ASOLDERAVBULL + AcquisBulletin + Arr, DCP)) + 'j au ' + datetostr(dateF));
        T.PutValue('PCN_JOURS', Arrondi(ASOLDERAVBULL + AcquisBulletin - cpt + arr, DCP));
      end;
      T := Tob_solde.findnext(['PCN_PERIODEPY'], ['0'], true);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : procedure qui solde les acquis de la periode bulletin
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure CalculAcquisPeriodeSolde(T_etab, T_Acquis: tob; DTDEBP, DTFINP, DateD, DateF: tdatetime; var NbJAcqP, NbJacq, NBJPayes: double; salarie: string); { PT105 }
var
  Ta {, tob_delta}: tob; { PT106 }
  i: integer;
  deltaJAutres, deltaJ0: double;
begin
 // tob_delta := nil; { PT106 }
  deltaJAutres := 0;
//  ChargeTobDelta(Salarie,datef, tob_delta); { PT105 } { PT106 }
  DeltaJ0 := CalculDelta(tob_delta, DTFINP, DTFINP, 0);
  for i := 1 to 10 do
    deltaJAutres := deltaJAutres + CalculDelta(tob_delta, DTFINP, DTFINP, i);
  NbjAcqp := 0;
  NbjAcq := 0;
  NBJPayes := 0;
  TA := T_Acquis.findfirst([''], [''], True);
  while TA <> nil do
  begin
    if (TA.getvalue('PCN_DATEVALIDITE') <= DateF) and (TA.getvalue('PCN_DATEVALIDITE') >= 0) then { PT105 } { PT106 }
    begin
      NbJAcq := NbJAcq + TA.GetValue('PCN_JOURS');
      NBJPayes := NBJPayes + TA.GetValue('PCN_APAYES');
      TA.PutValue('PCN_CODETAPE', 'S');
    end;
    if (TA.getvalue('PCN_DATEVALIDITE') <= DateF) and (TA.getvalue('PCN_DATEVALIDITE') >= DTDebP) then { PT105 } { PT106 }
    begin
      NbJAcqP := NbJAcqP + TA.GetValue('PCN_JOURS');
      TA.PutValue('PCN_CODETAPE', 'S');
    end;
    TA := T_Acquis.findnext([''], [''], True);
  end;
  { DEB PT104 Si SLD de cloture alors décremente à la hauteur du REL additionnée }
  if Assigned(Tob_Solde) then
  begin
    TA := Tob_Solde.findfirst([''], [''], True);
    while TA <> nil do
    begin
      if (TA.GetValue('PCN_TYPECONGE') = 'SLD') and (TA.GetValue('PCN_GENERECLOTURE') = 'X') then
        NbJAcq := NbJAcq - TA.GetValue('PCN_JOURS');
      TA := Tob_Solde.findnext([''], [''], True);
    end;
  end;
  { FIN PT104 }
  NBJAcq := NBJAcq - DeltaJ0 - deltaJAutres;
  if NBJAcq < 0 then NBJAcq := 0;
  NbJAcqP := NbJAcqP - DeltaJ0;
  if NBJAcqP < 0 then NBJAcqP := 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui vide la Tob_pris des mvts de SLD
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure SupprimeMvtSolde(Tob_pris: tob);
var
  T: tob;
begin
  if Tob_pris = nil then exit;
  T := Tob_Pris.findfirst(['PCN_TYPECONGE'], ['SLD'], false);
  while T <> nil do
  begin
    T.Free;
    T := Tob_Pris.findfirst(['PCN_TYPECONGE'], ['SLD'], false);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : positionne le code étape de tous les mvts d'acquis à S pour
Suite ........ : solde
Mots clefs ... : PAIE;CP
*****************************************************************}
{ PT-11-1 }

procedure SoldeAcquis(Tob_Acq: tob; Apayes: Boolean; LcDateFin: TDateTime); { PT106 }
var
  T: Tob;
  Cons: Double;
begin
  if Tob_Acq = nil then exit;
  //OkPayeCour:=(PayeCour>0);
  T := Tob_Acq.findfirst([''], [''], True);
  while T <> nil do
  begin
    if (T.GetValue('PCN_DATEVALIDITE') <= LcDateFin) then { PT106 }
      T.putvalue('PCN_CODETAPE', 'S');
      { DEB PT-26 }
    if (Apayes) and (PayeCour > 0) then
    begin
      Cons := T.Getvalue('PCN_APAYES');
      if Cons + PayeCour > T.GetValue('PCN_JOURS') then
      begin
        T.putvalue('PCN_APAYES', T.GetValue('PCN_JOURS')); { PT-11-1 Afin de consomme les acquis }
        PayeCour := PayeCour - (T.GetValue('PCN_JOURS') - Cons);
      end
      else
      begin
        T.putvalue('PCN_APAYES', T.Getvalue('PCN_APAYES') + PAYECOUR);
        PAYECOUR := 0;
      end;
    end;
      { FIN PT-26 }
(*     else
       if (Apayes=True) and (OkPayeCour=False) then { PT-27-2 On consomme de toute facon si PayeCour>0 }
         T.putvalue('PCN_APAYES',T.GetValue('PCN_JOURS'));   *)
    T := Tob_Acq.findnext([''], [''], True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : balaye la tob_pris, si on trouve un mvt de solde, on le
Suite ........ : bascule ds la tob_solde
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure IntegreSolde(Tob_solde, Tob_pris: tob);
var
  T: tob;
begin
  SLDCLOTURE := False; //PT61-1
  if Tob_pris = nil then exit;
  T := Tob_pris.findfirst(['PCN_TYPECONGE'], ['SLD'], True); //'PCN_GENERECLOTURE'
  while T <> nil do
  begin
    if T.GetValue('PCN_GENERECLOTURE') = 'X' then SLDCLOTURE := True; //PT61-1
    T.ChangeParent(Tob_solde, -1);
    T := Tob_pris.findnext(['PCN_TYPECONGE'], ['SLD'], True);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : affecte la période des mvts CP en traitement de clotûre des
Suite ........ :  congés
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure RecalculePeriode(Etab: string; DateClotCp: TDateTime);
var
  st: string;
  T_listeMvt, TS: Tob;
  PeriodePaye, DecalagePer: Integer;
begin
  try
    BeginTrans;
    T_listeMvt := Tob.create('ABSENCESALARIE', nil, -1);
    st := 'SELECT * FROM ABSENCESALARIE WHERE PCN_ETABLISSEMENT ="' + Etab + '" ' +
      'AND PCN_TYPEMVT="CPA" ORDER BY PCN_SALARIE , PCN_DATEVALIDITE'; //PT32 Ajout PCN_TYPEMVT="CPA"
{Flux optimisé
    Q := OpenSql(st, TRUE);
    if not Q.eof then
      T_listeMvt.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    ferme(Q);
}
    T_listeMvt.LoadDetailDBFROMSQL('ABSENCESALARIE', st);

    if (T_listeMVt <> nil) and (T_listeMVt.detail.count > 0) then
    begin
      TS := T_ListeMvt.FindFirst([''], [''], false);
      while TS <> nil do
      begin
        if ((IsPris(TS.getvalue('PCN_TYPECONGE'))) and (TS.getvalue('PCN_DATEPAIEMENT') > 10)) then
        begin
          {DEB PT-2}
          PeriodePaye := TS.GetValue('PCN_PERIODEPY');
          if (TS.GetValue('PCN_TYPECONGE') = 'SLD') and (TS.getvalue('PCN_MVTDUPLIQUE') = 'X') and (PeriodePaye = 1) then DecalagePer := 1 else DecalagePer := 0;
          TS.PutValue('PCN_PERIODECP', CalculPeriode(DateClotCp, TS.getvalue('PCN_DATEPAIEMENT')) + DecalagePer);
          {FIN PT-2}
          CalculPeriode(DateClotCp, TS.getvalue('PCN_DATEPAIEMENT'));
        end
        else
        begin
          TS.PutValue('PCN_PERIODECP', CalculPeriode(DateClotCp, TS.getvalue('PCN_DATEVALIDITE')));
          CalculPeriode(DateClotCp, TS.getvalue('PCN_DATEVALIDITE'));
        end;
        TS.UpdateDB(false);
        TS := T_ListeMvt.FindNext([''], [''], false);
      end;
    end;
    T_listeMvt.free;
    CommitTrans;
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la mise à jour des périodes CP', 'Ecriture établissement');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : - Charge les tobs nécessaire aux calculs de la provision
Suite ........ : - Calcul par salarié la provison
Suite ........ :
Mots clefs ... : PAIE;CP
*****************************************************************}

function CalculeProvisionCp({Etab,}StWhere: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime): Tob;
var
  TSAL, TW, Tabs, TSld: tob;
  Salarie, st: string;
  NbJoursReference {,NbJoursReferenceEtab}: integer;
  DTcloture: tdatetime;
  aa, mm, jj: word;
  iChamp0, iChamp1, iChamp2, iChamp3, iChamp4: integer;
begin
  DTcloture := Idate1900;
  // TPro : tob virtuelle des provisions
  Result := TOB.Create('Tob des provisions', nil, -1); { PT117 }
  Tsal := TOB.Create('Tob des Salaries', nil, -1);
  TAbs := TOB.Create('ABSENCE_SALARIE', nil, -1);
  TSld := TOB.Create('ABSENCE_SALARIE', nil, -1); { PT72 }
  { PT36 Charge tob tous établissement }
  StWhere := StWhere + ' AND (PSA_DATESORTIE <= "' + usdatetime(10) + '" OR PSA_DATESORTIE >= "' + usdatetime(DateEtat) + '"'
    + ' or PSA_DATESORTIE is null) '; { PT92 }
  st := 'SELECT PSA_SALARIE,PSA_CPACQUISMOIS,PSA_NBREACQUISCP,ETB_NBREACQUISCP,ETB_DATECLOTURECPN ' +
    'FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    ' WHERE ' + StWhere + ' ORDER BY PSA_ETABLISSEMENT,PSA_SALARIE'; { PT114 }
  Debug('Paie Pgi : Chargement des données');
{Flux optimisé
  Q := Opensql(st, true);
  if not Q.eof then
    Tsal.loaddetaildb('TobVirtuelle_des_SALARIEs', '', '', Q, false);
  ferme(Q);
}
  Tsal.loaddetaildbFROMSQL('TobVirtuelle_des_SALARIEs', st);

  { PT36 Optimisation traitement d'edition Mise en commentaire puis supprimer }
  Debug('Paie Pgi : Chargement des données.');
  ChargeTobsCalculProvision(TAbs, TSld, StWhere, DateEtat); { PT72 }
  // PT82   rajout du test si au moins 1 enregistrement salarié dans la TOB
  if (TSal <> nil) and (Tsal.Detail.count > 0) and (Result <> nil) and (Tabs <> nil) then { PT117 }
  begin
    { PT58 récupération du N° de champ , optimisation du traitement }
    iChamp0 := tsal.detail[0].GetNumChamp('PSA_SALARIE');
    iChamp1 := tsal.detail[0].GetNumChamp('PSA_CPACQUISMOIS');
    iChamp2 := tsal.detail[0].GetNumChamp('ETB_NBREACQUISCP');
    iChamp3 := tsal.detail[0].GetNumChamp('PSA_NBREACQUISCP');
    iChamp4 := tsal.detail[0].GetNumChamp('ETB_DATECLOTURECPN');
    while tsal.detail.count > 0 do
    begin
      Salarie := Tsal.detail[0].GetValeur(iChamp0);
      if Tsal.detail[0].GetValeur(iChamp1) = 'ETB' then //DEB PT8
        NbJoursReference := JoursReference(Tsal.detail[0].GetValeur(iChamp2))
      else
        NbJoursReference := JoursReference(Tsal.detail[0].GetValeur(iChamp3)); //FIN PT8
      DTcloture := Tsal.detail[0].GetValeur(iChamp4);
      TW := Tob.create('Tob des provisions', Result, -1); { PT117 }
      if assigned(TW) then
      begin
        FormateTobprovision(TW);
        Debug('Paie Pgi : Calcul provision salarié=' + Salarie);
        CalculProvision(Salarie, Tabs, TSld, TW, DateEtat, DTcloture, NbJoursReference); { PT72 }
      end;
      tsal.detail[0].free;
    end;
    { FIN PT58 }
  end
  else PGIINFO('Traitement impossible, le salarié n''est pas présent à la date d''arrêté.', '');
  //result := TPro; { PT117 Mise en commentaire }
  RendPeriode(DateFN1, DateDN, DTcloture, DateEtat);
  DateFN1 := DateDN - 1;
  Decodedate(DateDN, aa, mm, jj);
  DateDN1 := PGEncodeDateBissextile(aa - 1, mm, jj); { PT71 }
  FreeAndNil(TSal); //if Notvide(TSal, true) then Tsal.free; PT113
  FreeAndNil(TAbs); //if Notvide(TAbs, true) then TAbs.free; PT113
  FreeAndNil(TSld); { PT72 }
end;
// DEB PT87
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/01/2004
Modifié le ... :   /  /
Description .. : - Fonction d'initailisation des dates des
Suite ........ : - exercices CP pour provisions CP module Dotation
Suite ........ :
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure FgInitDotProvCp(Etab: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime);
var
  Q: tQuery;
  st: string;
  DTcloture: tdatetime;
  aa, mm, jj: word;
begin
  DTcloture := Idate1900;
  st := 'SELECT ETB_DATECLOTURECPN FROM ETABCOMPL WHERE ETB_ETABLISSEMENT ="' + Etab +
    '" AND ETB_CONGESPAYES="X"';
  Debug('Paie Pgi : Calcul des périodes CP');
  Q := Opensql(st, true);
  if not Q.eof then DtCloture := Q.findField('ETB_DATECLOTURECPN').AsDateTime;
  ferme(Q);
  RendPeriode(DateFN1, DateDN, DTcloture, DateEtat);
  DateFN1 := DateDN - 1;
  Decodedate(DateDN, aa, mm, jj);
  DateDN1 := PGEncodeDateBissextile(aa - 1, mm, jj);
end;
// FIN PT87
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Calcul le solde de chaque salarié
Mots clefs ... : PAIE;CP
*****************************************************************}

function CalculeSoldeCp({Etab,}StWhere: string; DateEtat: Tdatetime; var DateDN1, DateFN1, DateDN: tdatetime): Tob;
var
  TPro, T, TSAL, TW, Tabs, Tob_abs: tob;
  Q: tQuery;
  Salarie, st, Periode, PeriodeN1, PeriodeN2: string;
  Periodei: integer;
  DTcloture, DDebut, DFin: tdatetime;
  aa, mm, jj: word;
  PN, AN, RN, BN, MB, PN1, AN1, RN1, BN1, PN2, AN2, RN2, BN2, valo: double;
begin
  result := nil;
  // TPro : tob virtuelle des provisions
  TPro := TOB.Create('Tob des provisions', nil, -1);
  Tsal := TOB.Create('Tob des Salaries', nil, -1);
  TAbs := TOB.Create('ABSENCE_SALARIE', nil, -1);

  Debug('Paie Pgi : Chargement des données.');
  st := 'SELECT PSA_SALARIE,ETB_NBREACQUISCP,ETB_DATECLOTURECPN ' +
    'FROM SALARIES ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    // 'WHERE PSA_ETABLISSEMENT= "'+ Etab + '" '+
// ' '+StWhere+' '+
  'WHERE (PSA_DATESORTIE <= "' + usdatetime(10) + '" OR PSA_DATESORTIE >= "' + usdatetime(DateEtat) + '"' +
    ' or PSA_DATESORTIE is null) ' + StWhere;

// BEURK
  Q := Opensql(st, true);
  if not Q.eof then
  begin
    DTcloture := Q.findfield('ETB_DATECLOTURECPN').asDateTime; // BEURK
    Tsal.loaddetaildb('SALARIES', '', '', Q, false);
  end
  else
  begin
    TPro.free;
    Tsal.free;
    ferme(Q);
    exit;
  end;
  ferme(Q);

  ChargeTobsCalculProvision(TAbs, nil, StWhere, DateEtat); { PT72 }

  if (TSal <> nil) and (TPro <> nil) and (Tabs <> nil) then
  begin
    T := Tsal.findfirst([''], [''], false);
    while T <> nil do
    begin
      Salarie := T.getvalue('PSA_SALARIE');

      TW := Tob.create('Tob des provisions', TPro, -1);
      if TW <> nil then
      begin
        Debug('Paie Pgi : Calcul solde salarié=' + Salarie);
        FormateTobprovision(TW);
        Periodei := CalculPeriode(DTcloture, DateEtat);
        Periode := inttostr(Periodei);
        PeriodeN1 := Inttostr(Periodei + 1);
        PeriodeN2 := Inttostr(Periodei + 2);
        RechercheDate(DDebut, Dfin, DtCloture, Periode);

        Tob_abs := GetFromTobsSoldeCP(Tabs, Salarie, '', Periode, False, True); { PT-7 Modification de la fn appellée }
        CompteCP(Tob_abs, Periode, Salarie, DateEtat, PN, AN, RN, BN, MB, Valo);
        DateDN := DDebut;
        DDebut := 0;
        Dfin := 0;
        RechercheDate(DDebut, Dfin, DtCloture, PeriodeN1);
        Tob_abs := GetFromTobsSoldeCP(Tabs, Salarie, '', PeriodeN1, True, False); { PT-7 }
        CompteCP(Tob_abs, PeriodeN1, Salarie, DateEtat, PN1, AN1, RN1, BN1, MB, Valo);
        DateDN1 := DDebut;
        DateFN1 := Dfin;
        DateDN := DDebut;
        DDebut := 0;
        Dfin := 0;
        RechercheDate(DDebut, Dfin, DtCloture, PeriodeN2);
        Tob_abs := GetFromTobsSoldeCP(Tabs, Salarie, '', PeriodeN2, False, False); { PT-7 }
        CompteCP(Tob_abs, PeriodeN2, Salarie, DateEtat, PN2, AN2, RN2, BN2, MB, Valo);
        DateDN1 := DDebut;
        DateFN1 := Dfin;

        TW.PutValue('SALARIE', Salarie);
        TW.PutValue('RESTANTSPP', RN2);
        TW.PutValue('ACQUISN1', AN1);
        TW.PutValue('PRISN1', PN1);
        TW.PutValue('RESTANTSN1', RN1);
        TW.PutValue('ACQUISN', AN);
        TW.PutValue('PRISN', PN);
        TW.PutValue('RESTANTSN', RN);
      end;
      T := Tsal.findnext([''], [''], false);
    end;
  end;
  result := TPro;
  RendPeriode(DateFN1, DateDN, DTcloture, DateEtat);
  DateFN1 := DateDN - 1;
  Decodedate(DateDN, aa, mm, jj);
  DateDN1 := PGEncodeDateBissextile(aa - 1, mm, jj); { PT71 }
  FreeAndNil(TSal); //if Notvide(TSal, true) then Tsal.free; PT113
  FreeAndNil(TAbs); //if Notvide(TAbs, true) then TAbs.free; PT113
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : procedure qui charge la tob des mvts pour traitement de
Suite ........ : l'édition de la provision CP
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure ChargeTobsCalculProvision(TAbs, TSld: tob; StWhere: string; DateFin: tdatetime);
var
  St: string;
begin
  if TAbs = nil then exit;
  { PT36 Recupération des critères de sélection }
  if Trim(StWhere) <> '' then
    StWhere := 'AND PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE ' + Copy(StWhere, 5, Length(StWhere)) + ') ';
  { DEB PT72 Recherche SLD de sortie antérieur à la date d'arrêté }
  if Assigned(TSld) then
  begin
    St := 'SELECT DISTINCT PCN_SALARIE,PCN_DATEVALIDITE ' +
      'FROM ABSENCESALARIE WHERE PCN_GENERECLOTURE="-" ' +
      'AND PCN_TYPECONGE="SLD" AND PCN_TYPEMVT="CPA" ' +
      'AND PCN_DATEVALIDITE < "' + usdatetime(DateFin) + '" ' + StWhere +
      'ORDER BY PCN_SALARIE,PCN_DATEVALIDITE DESC';
{Flux optimisé
    Q := Opensql(st, True);
    if not Q.Eof then
      TSld.loaddetaildb('ABSENCE_SALARIE', '', '', Q, false, false);
    Ferme(Q);
  end;
}
    TSld.loaddetaildbFROMSQL('ABSENCE_SALARIE', st);
  end;

  { FIN PT72 }
  { ne pas tenir cpte de l'etablissement }
  st := 'SELECT PSA_ETABLISSEMENT,PCN_SALARIE,PCN_DATEFIN,PCN_PERIODECP,PCN_TYPECONGE,PCN_CODETAPE,' + { PT114 }
    'PCN_SENSABS,PCN_DATEPAIEMENT,PCN_JOURS,PCN_BASE,PCN_NBREMOIS,PCN_TYPEIMPUTE,' +
    'PCN_CODERGRPT,PCN_GENERECLOTURE,PCN_APAYES,PCN_PERIODEPAIE,PCN_DATEVALIDITE ' + { PT70-2 Ajout du apayes PT72 }
    'FROM ABSENCESALARIE,SALARIES ' + { PT92 } { PT114 }
    //'WHERE PSA_ETABLISSEMENT="'+Etab+'" PT36 Charge tob tous établissement
  'WHERE PSA_SALARIE=PCN_SALARIE AND PCN_CODERGRPT <> -1 AND PCN_TYPEMVT="CPA" ' + { PT114 }
    'AND PCN_DATEVALIDITE <= "' + usdatetime(DateFin) + '" ' + StWhere +
    ' ORDER BY PSA_ETABLISSEMENT,PCN_SALARIE,PCN_PERIODECP'; { PT114 }
  // 'AND (PCN_CODETAPE <> "C" AND PCN_CODETAPE <> "S")';
{Flux optimisé
  Q := Opensql(st, True);
  if not Q.eof then
    Tabs.loaddetaildb('ABSENCE_SALARIE', '', '', Q, false, false);
  Ferme(Q);
}
  Tabs.loaddetaildbFROMSQL('ABSENCE_SALARIE', st);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : fonction qui constitue à partir de la tob des mvts CP, une
Suite ........ : tob de mvts CP permettant le calcul de la provision
Suite ........ : Distinction de la période, des mvts REL
Mots clefs ... : PAIE;CP
*****************************************************************}

function GetFromTobsCalculProvision(Tabs: Tob; Salarie, Mvt, Periode: string; AvecArr, AvecRelPos, AvecRelNeg: Boolean): Tob;
var
  TMereAbs, TAbsSal, TAbsRech: Tob;
  SensAbs, Typeconge, MvtCloture: string;
begin
  Result := nil;
  if Tabs = nil then exit;
  TMereAbs := TOB.Create('Absence dupliquée', nil, -1);
  if mvt <> '' then
    TAbsSal := Tabs.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Salarie, Periode, Mvt], False)
  else
    TAbsSal := Tabs.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
  while TAbsSal <> nil do
  begin
    Typeconge := TAbsSal.GetValue('PCN_TYPECONGE');
    SensAbs := TAbsSal.GetValue('PCN_SENSABS');
    MvtCloture := TAbsSal.GetValue('PCN_GENERECLOTURE'); { PT70-3 }
    TAbsRech := TOB.Create('Absence dupliqué', TMereAbs, -1);
    { DEB PT59 : pour ne pas supprimer des mouvements non dupliqués }
    if (Typeconge <> 'REL') and (Typeconge <> 'ARR') then {DEB PT-7}
    begin
      TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      TAbsSal.free; { PT54 une fois dupliquée on suppr la fille PT59}
    end
    else
      if (Typeconge = 'ARR') and (AvecArr = True) then
      begin
        TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
        TAbsSal.free;
      end
      else
        if ((Typeconge = 'REL') and (SensAbs = '+')) and (AvecRelPos = True) then
        begin
          TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
          TAbsSal.free;
        end
        else
          if ((Typeconge = 'REL') and (SensAbs = '-')) and (AvecRelNeg = True) then //  PT-10 Ajout cond.
          begin
            TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
            TAbsSal.free;
          end {FIN PT-7}
      { FIN PT59 }
          else
      { DEB PT70-3 }
            if (Typeconge = 'ARR') and (MvtCloture = '-') then
            begin
              TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
              TAbsSal.free;
            end;
    { FIN PT70-3 }
    if mvt <> '' then
      TAbsSal := Tabs.FindNext(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Salarie, Periode, Mvt], False)
    else
      TAbsSal := Tabs.FindNext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
  end;
  Result := TMereAbs;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : fonction qui constitue à partir de la tob des mvts CP, une
Suite ........ : tob de mvts CP permettant le calcul du solde CP
Mots clefs ... : PAIE;CP
*****************************************************************}
{Nouvelle fonction PT-7}

function GetFromTobsSoldeCP(Tabs: Tob; Salarie, Mvt, Periode: string; Edit, SansCloture: Boolean): Tob;
var
  TMereAbs, TAbsSal, TAbsRech: Tob;
  SensAbs, Typeconge: string;
  Cloturer: Boolean;
begin
  Result := nil;
  if Tabs = nil then exit;
  TMereAbs := TOB.Create('Absence dupliquée', nil, -1);
  if mvt <> '' then
    TAbsSal := Tabs.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Salarie, Periode, Mvt], False)
  else
    TAbsSal := Tabs.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
  while TAbsSal <> nil do
  begin
    Typeconge := TAbsSal.GetValue('PCN_TYPECONGE');
    Cloturer := (TAbsSal.GetValue('PCN_GENERECLOTURE') = 'X');
    SensAbs := TAbsSal.GetValue('PCN_SENSABS');
    TAbsRech := TOB.Create('Absence dupliqué', TMereAbs, -1);
    if (SansCloture = False) and (edit = False) then
      TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
    if (SansCloture = True) and (Cloturer = False) then
      TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
    if (edit = True) and (StrToInt(periode) > 0) then
      if ((Cloturer = False) or ((Cloturer = True) and (((SensAbs = '+') and (Typeconge <> 'SLD')) or (Typeconge = 'SLD')))) then
        TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);

    if mvt <> '' then
      TAbsSal := Tabs.FindNext(['PCN_SALARIE', 'PCN_PERIODECP', 'PCN_TYPECONGE'], [Salarie, Periode, Mvt], False)
    else
      TAbsSal := Tabs.FindNext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
  end;
  Result := TMereAbs;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... : 03/02/2003
Description .. : Formate la tab provision
Mots clefs ... : PAIE;CP
*****************************************************************}

procedure FormateTobprovision(TW: tob);
begin
  with TW do
  begin
    AddChampSup('SALARIE', False);
    AddChampSup('CPBASERELN1', False); //BASEPP
    AddChampSup('CPMOISRELN1', False); //MOISPP
    AddChampSup('CPACQUISRELN1', False); //ACQUISPP
    AddChampSup('CPPRISRELN1', False); //PRISPP
    AddChampSup('CPRESTRELN1', False); //RESTANTSPP
    AddChampSup('CPPROVRELN1', False);

    AddChampSup('CPBASEN1', False); //BASEN1
    AddChampSup('CPMOISN1', False); //MOISN1
    AddChampSup('CPACQUISN1', False); //ACQUISN1
    AddChampSup('CPPRISN1', False); //PRISN1
    AddChampSup('CPRESTN1', False); //RESTANTSN1
    AddChampSup('CPPROVN1', False);

    AddChampSup('CPBASEN', False); //BASEN
    AddChampSup('CPMOISN', False); //MOISN
    AddChampSup('CPACQUISN', False); //ACQUISN
    AddChampSup('CPPRISN', False); //PRISN
    AddChampSup('CPRESTN', False); //RESTANTSN
    AddChampSup('CPPROVN', False);

    AddChampSup('CPPROVISION', False); //BASEPRO
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : Calcul de la provision par salarié
Mots clefs ... : PAIE;CP
*****************************************************************}
{ DEB PT72 : Refonte de la procedure :
             L'édition s'édite sur les deux dernières périodes antérieur à la date d'arrêté
             Au délà les périodes ne sont plus édités
             Les reliquats + de la seconde période ne sont plus intégrés aux acquis de cette période
             Ils sont édités en période précédente
             Les mvts de SLD de sortie ne sont plus édités si date d'arrêté superieur à la date du solde }

procedure CalculProvision(Salarie: string; TAbs, TSld, TW: tob; DateEtat, DTcloture: TDatetime; NbJoursReference: integer);
var
  AN, PN, RN, BN, MB, ANRel, PNRel, RNRel, BNRel, MBRel, Valo: double;
  ValoRel: double;
  i: Integer;
  PeriodeN, StI: string;
  DateSolde: TDateTime;
  T, LcTob_Abs: Tob;
begin
  { PT74-3 Refonte calcul }
  DateSolde := IDate1900;
  TW.PutValue('SALARIE', Salarie);
  { Recherche de la date du dernier solde de sortie }
  if Assigned(TSld) then
  begin
    T := TSld.FindFirst(['PCN_SALARIE'], [Salarie], False);
    if Assigned(T) then
    begin
      DateSolde := T.GetValue('PCN_DATEVALIDITE');
      T.Free;
    end;
  end;
  PeriodeN := inttostr(CalculPeriode(DTcloture, DateEtat));
  i := 1;
  AnRel := 0; ValoRel := 0;
  while i < 3 do
  begin
    AN := 0;
    PN := 0;
    BN := 0;
    MB := 0;
    PNRel := 0;
    RNRel := 0;
    BNRel := 0;
    MBRel := 0;
    Valo := 0;
    { Chargement des acquis de la période }{ PT92 Modif Tob_abs en LcTob_abs}
    Debug('Salarie=' + Salarie);
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'ACQ', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
    begin
      AN := LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['+'], False);
      AN := AN - LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['-'], False);
      BN := LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['+'], False);
      BN := BN - LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['-'], False);
      MB := LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['+'], False);
      MB := MB - LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['-'], False);
    end;
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'AJUACQ', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
    begin
      AN := AN + LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['+'], False);
      AN := AN - LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['-'], False);
      BN := BN + LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['+'], False);
      BN := BN - LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['-'], False);
      MB := MB + LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['+'], False);
      MB := MB - LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['-'], False);
    end;
    { Chargement des Pris payés de la période }
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'PRI', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
    begin
      PN := LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['-'], False);
      PN := PN - LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['+'], False);
    end;
    { Chargement des Pris payés sur reliquat de la période }
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'PRIREL', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
      PNRel := LcTob_abs.somme('PCN_JOURS', [''], [''], False);
    if PN < 0 then
    begin
      PNRel := PNRel + PN;
      PN := 0;
    end;

    { Chargement des Ajustement sur reliquat de la période }
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'AJUREL', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
    begin
      ANRel := LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['+'], False);
      ANRel := ANRel - LcTob_abs.somme('PCN_JOURS', ['PCN_SENSABS'], ['-'], False);
      BNRel := LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['+'], False);
      BNRel := BNRel - LcTob_abs.somme('PCN_BASE', ['PCN_SENSABS'], ['-'], False);
      MBRel := LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['+'], False);
      MBRel := MBRel - LcTob_abs.somme('PCN_NBREMOIS', ['PCN_SENSABS'], ['-'], False);
      RNRel := Arrondi(ANRel - PNRel, DCP);
    end;
    { Chargement des reliquats + de la période }
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    LcTob_abs := GetMvtFromTobProvision(Tabs, TSld, Salarie, 'REL', PeriodeN, DateEtat, DateSolde);
    if Assigned(LcTob_abs) then
    begin
      ANRel := ANRel + LcTob_abs.somme('PCN_JOURS', [''], [''], False);
      BNRel := BNRel + LcTob_abs.somme('PCN_BASE', [''], [''], False);
      MBRel := MBRel + LcTob_abs.somme('PCN_NBREMOIS', [''], [''], False);
      RNRel := Arrondi(ANRel - PNRel, DCP);
    end;
    FreeAndNil(LcTob_abs); { PT92 Suppression tob }
    if (ANRel <= 0) and (PNRel > 0) then
    begin
      PN := Arrondi(PN + PNRel, DCP);
      PNRel := 0;
      RNRel := Arrondi(ANRel - PNRel, DCP);
    end;
    RN := Arrondi(AN - PN, DCP);
    if (PN > 0) or (AN > 0) then
    begin
      if (RN < 0) and (AN > 0) and (RNRel > 0) then
      begin
        PNRel := Arrondi(PNRel - RN, DCP);
        RN := 0;
      end
    end;
    if (ANRel > 0) then RNRel := Arrondi(ANRel - PNRel, DCP);

    { Calcul des provisions }
    if Arrondi(MB * 10 * NbJoursReference, DCP) <> 0 then
      Valo := (arrondi((12 * BN) / (MB * 10 * NbJoursReference), DCP)) * RN;
    if (RNRel > 0) and (i = 2) then
      if (Arrondi(MBRel * 10 * NbJoursReference, DCP) <> 0) then
        ValoRel := (arrondi((12 * BNRel) / (MBRel * 10 * NbJoursReference), DCP)) * RNRel;

    if i = 1 then StI := '' else StI := '1';
    TW.PutValue('CPBASEN' + StI, BN); //BASEN & BASEN1
    TW.PutValue('CPMOISN' + StI, MB); //MOISN & MOISN1
    TW.PutValue('CPACQUISN' + StI, AN); //ACQUISN & ACQUISN1
    TW.PutValue('CPPRISN' + StI, PN); //PRISN & PRISN1
    TW.PutValue('CPRESTN' + StI, RN); //RESTANTSN & RESTANTSN1
    TW.PutValue('CPPROVN' + StI, Valo); //PROVN & PROVN1
    if i > 1 then
    begin
      TW.PutValue('CPBASERELN1', BNRel); //BASEPP
      TW.PutValue('CPMOISRELN1', MBRel); //MOISPP
      TW.PutValue('CPACQUISRELN1', ANRel); //ACQUISPP
      TW.PutValue('CPPRISRELN1', PNRel); //PRISPP
      TW.PutValue('CPRESTRELN1', RNRel); //RESTANTSPP
      TW.PutValue('CPPROVRELN1', ValoRel); //PROVN & PROVN1
    end;
    TW.PutValue('CPPROVISION', Valeur(TW.GetValue('CPPROVISION')) + Valo + ValoRel); //BASEPRO
    PeriodeN := IntToStr(ValeurI(PeriodeN) + 1);
    Inc(i);
  end;
  { PT92 Suppression tob }
  LcTob_abs := TABs.findFirst(['PCN_SALARIE'], [Salarie], FALSE);
  while Assigned(LcTob_abs) do
  begin
    FreeAndNil(LcTob_abs);
    LcTob_abs := TABs.findNext(['PCN_SALARIE'], [Salarie], FALSE);
  end;
  if not (isnumeric(Salarie)) then { DEB PT114 }
  begin
    LcTob_abs := TABs.findFirst(['PCN_SALARIE'], [LowerCase(Salarie)], FALSE);
    while Assigned(LcTob_abs) do
    begin
      FreeAndNil(LcTob_abs);
      LcTob_abs := TABs.findNext(['PCN_SALARIE'], [LowerCase(Salarie)], FALSE);
    end;
  end; { FIN PT114 }

end;
{ FIN PT72 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie PGI
Créé le ...... : 18/03/2004
Modifié le ... :   /  /
Description .. : Extrait de la tob des CP les différents types de mouvement
Mots clefs ... : PAIE;CONGESPAYES
*****************************************************************}
{ DEB PT72 }

function GetMvtFromTobProvision(Tabs, TSld: Tob; Salarie, Typ, Periode: string; DateEtat, DateSolde: TDateTime): Tob;
var
  TMereAbs, TAbsSal, TAbsRech: Tob;
  SensAbs, Typeconge, TypeImpute: string;
  i: Integer;
begin
  Result := nil;
  if Tabs = nil then exit;
  TMereAbs := TOB.Create('Absence dupliquée', nil, -1);
  //TAbsSal := Tabs.FindFirst(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
  i := 0; { PT92 Suppression du FindFirst, FindNext }
  while (i <= Tabs.detail.count - 1) do
  begin
    TAbsSal := Tabs.Detail[i];
    if UPPERCase(TAbsSal.getvalue('PCN_SALARIE')) <> Salarie then { PT114 }
      Break;
    if TAbsSal.getvalue('PCN_PERIODECP') <> Periode then
    begin
      i := i + 1;
      continue;
    end;
    TAbsRech := TOB.Create('Absence dupliqué', TMereAbs, -1);
    Typeconge := TAbsSal.GetValue('PCN_TYPECONGE');
    SensAbs := TAbsSal.GetValue('PCN_SENSABS');
    // DEB PT86
    if TAbsSal.getvalue('PCN_TYPEIMPUTE') <> NULL then TypeImpute := TAbsSal.getvalue('PCN_TYPEIMPUTE')
    else TypeImpute := '';
    //    TypeImpute := TAbsSal.GetValue('PCN_TYPEIMPUTE'); { PT74-3 }
    // FIN PT86
    if (DateSolde <> Idate1900) and (DateSolde <= DateEtat) and (TAbsSal.GetValue('PCN_CODETAPE') <> '...') and (TAbsSal.GetValue('PCN_DATEVALIDITE') <= DateSolde) then
    begin
      // PT92 TAbsSal.Free;
    end
    else
      if (Typ = 'ACQ') and (IsAcquis(TypeConge, SensAbs) = True) and (TypeConge <> 'REL') and (TypeConge <> 'ARR') and (TypeConge <> 'AJU') then
      begin
        TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      // PT92 TAbsSal.free;
      end
      else
        if (Typ = 'ACQ') and (TypeConge = 'ARR')
          and (((TAbsSal.getvalue('PCN_DATEVALIDITE') < DateEtat) and (TAbsSal.getvalue('PCN_GENERECLOTURE') = 'X'))
          or ((TAbsSal.getvalue('PCN_DATEVALIDITE') <= DateEtat) and (TAbsSal.getvalue('PCN_GENERECLOTURE') = '-'))) then
        begin
          TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      // PT92 TAbsSal.free;
        end
        else
          if (Typ = 'AJUACQ') and (TypeConge = 'AJU') and (TypeImpute = 'ACQ') then
          begin
            TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      // PT92 TAbsSal.free;
          end
          else
            if (Typ = 'AJUREL') and (TypeConge = 'AJU') and (TypeImpute = 'REL') then
            begin
              TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      // PT92 TAbsSal.free;
            end
            else
              if ((Typ = 'PRI') or (Typ = 'PRIREL')) and ((IsPris(TypeConge)) and ((TAbsSal.getvalue('PCN_DATEFIN') > 10)
                and (TAbsSal.getvalue('PCN_DATEFIN') <= dateEtat) and (TAbsSal.getvalue('PCN_DATEPAIEMENT') > 10))
                or (Typeconge = 'CPA')) then
              begin
                if (Typ = 'PRIREL') and ((TypeConge = 'PRI') or (TypeConge = 'SLD')) and (Pos('Reliquat', TAbsSal.GetString('PCN_PERIODEPAIE')) > 0) then { PT75-3 Ajout clause PT95 }
                begin
                  TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
        // PT92 TAbsSal.free;
                end
                else
                  if (Typ = 'PRI') and (Pos('Reliquat', TAbsSal.GetString('PCN_PERIODEPAIE')) < 1) then { PT95 }
                  begin
                    TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
        // PT92 TAbsSal.free;
                  end;
              end
              else
                if (Typ = 'REL') and (Typeconge = 'REL') and (SensAbs = '+') then
                begin
                  TAbsRech.Dupliquer(TAbsSal, FALSE, TRUE, TRUE);
      // PT92 TAbsSal.free;
                end;
    // PT92 TAbsSal := Tabs.FindNext(['PCN_SALARIE', 'PCN_PERIODECP'], [Salarie, Periode], False);
    i := i + 1;
  end;
  Result := TMereAbs;
end;
{ FIN PT72 }

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 20/12/2002
Modifié le ... : 03/02/2003
Description .. : Charge une tob du calendrier salarié
Mots clefs ... : PAIE;CALENDRIER
*****************************************************************}
{PT51 Modification de la function on récupère le calendrier en fonction du paramétrage salarié}

function ChargeCalendrierSalarie(Etablissement, Salarie, Calend, StandCalend, EtabStandCalend: string): Tob;
var
  StQ: string;
  Q: TQuery;
begin
  result := nil;
  StQ := '';
  if StandCalend = '' then
  begin
    Q := OpenSql('SELECT PSA_CALENDRIER,PSA_STANDCALEND,ETB_STANDCALEND FROM SALARIES ' +
      'LEFT JOIN ETABCOMPL ON ETB_ETABLISSEMENT="' + Etablissement + '" ' +
      'WHERE PSA_SALARIE="' + Salarie + '"', True);
    if not Q.eof then
    begin
      Calend := Q.FindField('PSA_CALENDRIER').asstring;
      StandCalend := Q.FindField('PSA_STANDCALEND').asstring;
      EtabStandCalend := Q.FindField('ETB_STANDCALEND').asstring;
    end;
    Ferme(Q);
  end;
  { 1 - on cherche au niveau calendrier spécifique salarie }
  if StandCalend = 'PER' then
  begin
    Q := Opensql('SELECT * FROM CALENDRIER WHERE ACA_SALARIE="' + Salarie + '" ORDER BY ACA_JOUR', True);
    if (Q.eof) and (calend <> '') then
      { 2 - on cherche au niveau standard de calendrier salarie }
      StQ := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + Calend + '" AND ACA_SALARIE="***" ORDER BY ACA_JOUR'
    else
      if (not Q.eof) then
        StQ := 'SELECT * FROM CALENDRIER WHERE ACA_SALARIE="' + Salarie + '" ORDER BY ACA_JOUR';
    Ferme(Q);
  end
  else
    { 2 - on cherche au niveau standard de calendrier salarie }
    if StandCalend = 'ETS' then
    begin
      if Calend = '' then exit;
      StQ := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + Calend + '" AND ACA_SALARIE="***" ORDER BY ACA_JOUR';
    end
    else
    { 3 - on cherche au niveau calendrier standard établissement }
      if StandCalend = 'ETB' then
      begin
        if EtabStandCalend = '' then exit;
        StQ := 'SELECT * FROM CALENDRIER WHERE ACA_STANDCALEN="' + EtabStandCalend + '" AND ACA_SALARIE="***" ORDER BY ACA_JOUR';
      end;
  {PT51 Mise en commentaire puis supprimer  }
  { DEB PT-14 }
  if StQ <> '' then { PT51 }
  begin
{Flux optimisé
    Q := OpenSql(StQ, True);
    Result := Tob.create('Le calendrier', nil, -1);
    Result.LoadDetailDB('Cal_endrier', '', '', Q, False);
    Ferme(Q);
}
    Result := Tob.create('Le calendrier', nil, -1);
    Result.LoadDetailDBFROMSQL('Cal_endrier', stQ);
  end
  else
  begin { Aucun calendrier affecté au salarié }
    result := nil;
    Exit;
  end;
  { FIN PT-14 }
end;
{ DEB PT71 }

(* PT89 function PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
  Result := Idate1900;
  if IsValidDate(IntToStr(JJ) + '/' + IntToStr(MM) + '/' + IntToStr(AA)) then
    Result := encodedate(AA, MM, JJ);
  if (MM = 2) and ((JJ = 28) or (JJ = 29)) then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end;
end;
{ FIN PT71 }*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : PT108 : Renvoie le paramètre : Salaire paie retenu pour le calcul du
Suite ........ : maintien
Mots clefs ... : PAIE;CP
*****************************************************************}

function TrouvePaieSalMaintien(T_etab, T_Sal: tob): string;
begin
  result := 'ANT';
  if T_Sal.GetValeur(iPSA_TYPPAIEVALOMS) = 'PER' then
    Result := T_Sal.GetValeur(iPSA_PAIEVALOMS)
  else
    Result := T_Etab.getvalue('ETB_PAIEVALOMS');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : PT108 : Renvoie le cumul 12 correspondant à la période defini en
Suite ........ : fonction du paramétrage
Mots clefs ... : PAIE;CP
*****************************************************************}
{ PT108 }

function RecupSalMaintien(PaieSalMaint: string; DateAbs, DateBull: TDateTime): Double;
var DBPrecPaiAbs, DFPrecPaiAbs, DBEncpaieAbs, DFEncpaieAbs: TDateTime;
  DBPrecPaiBull, DFPrecPaiBull, DBEncpaieBull, DFEncpaieBull: TDateTime;
  Mt1, Mt2: Double;
begin
  Result := 0;
  if (PaieSalMaint = 'ANT') or (PaieSalMaint = 'ENC') or (PaieSalMaint = 'VAN') then
  begin { Dans le cas où l'on retient le salaire sur la date du bulletin }
    DBEncpaieBull := DebutDeMois(DateBull);
    DFEncpaieBull := FinDeMois(DateBull);
    DFPrecPaiBull := DBEncpaieBull - 1;
    DBPrecPaiBull := DebutdeMois(DFPrecPaiBull);
    if PaieSalMaint = 'ANT' then
    begin
      Result := Valcumuldate('12', DBPrecPaiBull, DFPrecPaiBull);
      if Result = 0 then Result := Valcumuldate('12', DBEncpaieBull, DFEncpaieBull);
    end
    else
      if PaieSalMaint = 'ENC' then
      begin
//         Result := Valcumuldate('12', DBEncpaieBull, DFEncpaieBull);
        Result := RendCumulSalSess('12'); // PT123
      end
      else
      begin
        Mt1 := Valcumuldate('12', DBPrecPaiBull, DFPrecPaiBull);
        Mt2 := Valcumuldate('12', DBEncpaieBull, DFEncpaieBull);
        if Mt2 = 0 then Mt2 := RendCumulSalSess('12'); { PT115-1 }
        if Mt1 > Mt2 then Result := Mt1 else result := Mt2;
      end;
  end
  else { Dans le cas où l'on retient le salaire sur la date d'absence }
    if (PaieSalMaint = 'ZAN') or (PaieSalMaint = 'ZNC') or (PaieSalMaint = 'ZVA') then
    begin
      DBEncpaieAbs := DebutDeMois(DateAbs);
      DFEncpaieAbs := FinDeMois(DateAbs);
      DFPrecPaiAbs := DBEncpaieAbs - 1;
      DBPrecPaiAbs := DebutdeMois(DFPrecPaiAbs);
      if PaieSalMaint = 'ZAN' then
      begin
        Result := Valcumuldate('12', DBPrecPaiAbs, DFPrecPaiAbs);
        if Result = 0 then Result := Valcumuldate('12', DBEncpaieAbs, DFEncpaieAbs);
      end
      else
        if PaieSalMaint = 'ZNC' then
        begin
          Result := Valcumuldate('12', DBEncpaieAbs, DFEncpaieAbs);
        end
        else
        begin
          Mt1 := Valcumuldate('12', DBPrecPaiAbs, DFPrecPaiAbs);
          Mt2 := Valcumuldate('12', DBEncpaieAbs, DFEncpaieAbs);
          if Mt1 > Mt2 then Result := Mt1 else result := Mt2;
        end;
    end;

  { Dans tous les cas }
  if FirstBull then Result := RendCumulSalSess('12')
  else if Result = 0 then Result := RendCumulSalSess('12');

end;

end.


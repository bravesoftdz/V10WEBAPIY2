{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... :   /  /
Description .. : Unit de définition de fonctions génériques de la paie
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 25/06/2001 PH V536 Suppression  des .Next
PT2   : 27/06/2001 PH V536 Utilisation de la fonction PresenceComplexe
PT3   : 28/06/2001 PH V536 On ne ventile pas dans analytique de la paie sur un
                           compte d'attente
PT4   : 03/07/2001 VG V540 Rajout d'une fonction qui rend l'exercice précédent
PT5   : 02/08/2001 VG V547 Génération de fichier sur lecteur réseau
PT6   : 16/08/2001 PH V547 PGRempliDefaut devient une fonction générique de la
                           paie
PT7   : 27/08/2001 SB V547 MajConvention procedure ExisteSql répétitive
                           Integré dans une clause conditionnelle
PT8   : 03/09/2001 VG V547 Ajout procédure de purge des caractères non
                           numériques d'une chaîne.
PT9   : 25/10/2001 SB V562 Mise à jour du ORDREETAT pour les rubriques de
                           cotisation..
PT10  : 03/10/2001 PH V562 Traitement ventilations analytiques provenant d'un
                           import
PT11  : 12/12/2001 PH V562 Controle du mois renseigné sinon 1 par defaut car
                           cela provoque une erreur de conversion
PT12  : 21/01/2002 PH V571 Réaffectation des dates début et fin dans
                           YVA_IDENTIFIANT en création de bulletin
PT13  : 01/03/2002 SB V571 Moulinette des maj des bases des mvts d'acquis congés
                           payés
PT14  : 04/04/2002 Ph V571 Prise en compte des THDBRichEditOLE en lecture seule
PT15  : 10/04/2002 Ph V571 Traitement des cumuls si pair et > 50 et STD alors Ok
PT16  : 17/04/2002 SB V571 Plantage SQL sur InsertDb de Moulinette des maj des
                           bases des mvts d'acquis congés payés
PT17  : 29/04/2002 PH V582 Optimisation SQL en supprimant like
PT18  : 30/04/2002 PH V582 Modifs des libellés de champs libres dans dechamp
PT19  : 11/06/2002 PH V582 Fonction qui les périodes de la saisie des elts
                           variables
PT20  : 15/07/2002 VG V585 Procédure qui retourne le code ISO d'un pays sur 2
                           caractères et son libellé à partir du code ISO sur 3
                           caractères + Fonction qui remplace les minuscules par
                           des majuscules non accentuées
PT21  : 25/07/2002 SB V585 Dû à l'intégration des mvts d'annulation d'absences
                           Contrôle des requêtes si typemvt et sensabs en
                           critere
PT22  : 09/08/2002 MF V585 modif fct ForceNumérique pour initialiser le champ
                           résultat
PT23  : 06/09/2002 PH V585 Traitement libelle des champs elt de salaire 5
PT24  : 17/09/2002 PH V585 Optimisation requete sql en remplaçant like par = car
                           valeur exacte
PT25  : 17/09/2002 PH V585 Optimisation Ventilation analytique salarié non
                           rechargées systématiquement à chaque rubrique
PT26  : 18/09/2002 PH V585 Indication si le elt de salaire sont mensuels ou
                           annuels dans les intitulés des zones
PT27  : 10/10/2002 PH V585 recherche de tous les champs composant la clé des
                           ventilations comptables
PT28  : 10/10/2002 VG V585 Implémentation de la fonction TestNumeroSS qui se
                           trouvait dans UTOMSalaries ET dans UtomInterimaires
PT29  : 16/12/2002 PH V591 Fonction dupliquerpaie rendue compatible CWAS
PT30  : 30/12/2002 SB V591 FQ 10401 Création nouvelle fonction pour test
                           exercice actif
PT31  : 08/01/2003 PH V591 Suppression PT17 car ne prend plus en compte specif
                           sal/rub . Optimisation faite pour CEGID.
PT32  : 09/01/2003 PH V591 Gestion dynamique des modules de la paie en fonction
                           des serias
PT33  : 31/01/2003 PH V591 Modif tablette PGSALPREPAUTO pour gérer PT8 de
                           UTofPG_MulPrepAuto
PT34  : 02/04/2003 PH V_42 Fonction controle de cohérence et recup modif sur
                           NODOSSIER
PT35  : 06/05/2003 SB V_42 Orthographe
// ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
// ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
PT36  : 23/06/2003 PH V_42 correction erreur de focalisation fenetre d'une
                           variable CEGID en modif autorisée
PT37  : 01/07/2003 JL V_42 FQ : 10652 Contrôle de la clé du numSS par rapport au
                           département 20 si 2a ou 2b
PT38  : 11/08/2003 PH V_42 Boucle sur la tob de la table PAIEVENTIL
PT38-1: 28/08/2003 PH V_42 Correction si pas de ventilations analytiques alors
                           pas de type de ventilation par défaut
PT39  : 04/09/2003 PH V_42 Suppression menu 49 en CWAS
PT40  : 04/09/2003 PH V_42 Chgt suffixe tablettes CODEEMPLOI dans le cas passage
                           2003
PT41  : 04/11/2003 PH V_50 Fonction de dédoublonnage des ventilations analytique
PT42  : 18/12/2003 VG V_50 FQ 10542 Tolérance N° SS provisoire
PT43  : 06/01/2004 PH V_50 Fonction pour complèter les préventilations
                           analytiques alors qu'elles ne sont pas renseignées
                           dans le fichier d'import
PT44  : 30/01/2004 PH V_50 Mise en place du MENU 48
PT45  : 10/05/2004 PH V_50 restriction droit accès au paramétrage CEGID
PT46  : 03/06/2004 PH V_50 FQ 11271 CAs des pré-ventilations analytiques
                           provenant fichier d'import pour une rem tq il existe
                           une cotisation avec le même numéro
PT47  : 05/06/2004 PH V_50 Mise en place gestion des menus complémentaires
                           Gestion de compétences et bilan social
PT48  : 03/08/2004 VG V_50 N° sécurité sociale salarié né en Corse après
                           01/01/1976 - FQ N°11342
PT49  : 09/08/2004 JL V_50 Gestion affichage libellé champs libres compétences
PT50  : 11/08/2004 PH V_50 FQ 11030 Test du pourcentage <> 0 avant affectation
PT51  : 24/09/2004 VG V_50 Les fonctions suivantes ont été déplacées dans
                           PgOutils2 : AffectDefautCode, ColleZeroDevant,
                           RendExerSocialPrec, ForceNumerique, ControlSiret
PT52  : 28/09/2004 VG V_50 Optimisation
PT53  : 02/11/2004 PH V_60 Prise en compte nouveau type de ventilation VSA = VEN
                           mais pour le mois en cours. les enreg sont stockés
                           dans PAIEVENTIl et proviennent du fichier d'import
PT54  : 08/11/2004 PH V_60 procédure de contrôle de date courte devient une
                           fonction pour validité de la zône saisie
PT55  : 17/11/2004 JL V_60 Remplissage tabletet PGZONELIBRE par défaut
PT56  : 17/11/2004 JL V_60 Gestion du DIF, utilisation du champ PFO_CIF
PT57  : 20/12/2004 PH V_60 Gestion d'1 tableau dynamique pour utilisation de la
                           fonction PresenceComplexe
PT58  : 31/01/2005 SB V_65 Export des proc. communes dans pgoutils2
PT59  : 18/04/2005 SB V_65 FQ 11993 Ajout champ pha_ordreetat
PT60  : 15/02/2005 JL V_60 Ajout module 303 augmentation
PT61  : 22/04/2005 PH V_60 Filtrage des tablettes pour la confidentialité
                           établissement
PT62  : 13/09/2005 PH V_60 FQ 11966 On autorise la modification de la gestion
                           associée si Rub CEGID et Si Réviseur
PT63  : 27/09/2005 SB V65  FQ 12602 Ajout mode silence supprime fichier
PT64  : 05/10/2005 PH V65  Cas où brut=0 mais ventilations analytiques des
                           cotisations à partir de la ventilation salarié
PT65  : 13/10/2005 PH V65  Initialisation de la tablette PGCATBILAN avec une
                           valeurs par defaut 000 sans catégorie
PT66  : 19/10/2005 VG V_60 Fonction SupprimeFichier déplacée dans PgOutils2
PT67  : 18/04/2006 NA V_65 Procédure PaieLectureSeule : Autoriser la barre de
                           déplacement si lecture seule pour les THGRID
PT68  : 29/05/2006 PH V_65 Accès module bilan social idem les autres modules
PT69  : 12/06/2006 SB V_65 Refonte des zoom sur etats
PT70  : 13/06/2006 PH V_65 FQ 12964 Ventilations analytiques sur matricule alpha
                           de longueur variable
PT71  : 15/06/2006 PH V_65 Correction de la tablette JUTYPECIVIL Erreur ORACLE
PT72  : 23/06/2006 VG V_70 Suppression du menu "specifiques"
PT73  : 17/10/2006 SB V_70 Refonte Fn pour utilisation portail => exporter vers
                           pgcalendrier
PT74  : 26/10/2006 PH V_70 Mise en place gestion des habilitations
PT75  : 16/11/2006 PH V_70 On interdit le paramètrage STD en PAIE50
PT76  : 24/11/2006 PH V_70 Utilitaire reaffectation de numéro de dossier en
                           multisoc
PT77  : 11/12/2006 PH V_70 Gestion de l'analytique sur les rubriques de régul =
                           idem rubrique non regul
PT78  : 19/01/2006 SB V_70 Ajout module des présence
PT79  : 26/01/2007 FC V_70 Ajout d'une fonction qui renvoie la clause where
                           existante à laquelle est rajoutée la clause pour
                           l'habilitation
PT80  : 12/02/2007 FC V_70 Ajout d'une nouvelle procédure pour la duplication
                           d'un élément national dossier
PT81  : 06/03/2007 FC V_70 Ajout de la convention collective pour la duplication
                           d'un élément national
PT82  : 07/03/2007 FC V_70 Récupération du mode d'édition
PT83  : 03/05/2007 PH V_70 FQ 13960 Création elt nationaux STD
PT84  : 11/05/2007 PH V_70 Moulinette mutil dossier affectation du numéro de
                           dossier
PT85  : 27/06/2007 VG V_72 Procedure pour mettre à jour les tiers sur
                           modification du salarié - FQ N°13476
PT86  : 29/06/2007 VG V_72 Affichage modules présence et masse salariale
PT87  : 03/07/2007 VG V_72 Pour tests nouveaux modules - à supprimer pour
                           diffusion
PT88  : 01/08/2007 GGU V_8 FQ 14544 : indice de liste hors limite
PT89  : 04/09/2007 FC V_80 FQ 14605 : Duplication d'un établissement
PT90  : 22/11/2007 FLO V_81 Ajout d'une fonction rendant le dernier mois clos
                            d'un exercice
PT91  : 04/12/2007 FLO V_81 Optimisation du chargement des tablettes libres
PT92  : 14/12/2007 FLO V_81 Ajout d'une fonction pour le chargement de la liste
                           des dossiers d'un regroupement et d'une autre pour le
                           lancement d'une autre instance de l'application
PT94  : 03/04/2008 FLO V_81 Ajout de la fonction GereAccesPredefini pour limiter
                           les choix en fonction des droits utilisateur
}

unit Pgoutils;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  SysUtils,
  HCtrls,
  HEnt1,
  Classes,
{$IFNDEF EAGLSERVER}
  LookUp,
{$IFDEF EAGLCLIENT}
  eFiche,
  eFichList,
  emul,
  MaineAgl,
{$ELSE}
  Fiche,
  fichlist,
  mul,
  MenuOLG,
  Fe_main,
{$ENDIF}
{$ENDIF}
{$IFDEF EAGLCLIENT}
  UtileAGL,
  MenuOLX,
  Spin,
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  Db,
  HDB,
  DBCtrls,
{$ENDIF}
  Hqry,
  StdCtrls,
  Graphics,
  HStatus,
  HRichOle,
  Hdebug,
  Controls,
  forms,
  ComCtrls,
  extCtrls,
  HMsgBox,
  UTOM,
  UTOB,
  HTB97,
  Buttons,
  Grids,
  ParamSoc,
  Math,
  Dialogs,
  Ed_tools,
{$IFNDEF PLUGIN}
  PgOutils2,
{$ENDIF}  
  Windows //PT91
  ;

{$IFNDEF PLUGIN}
procedure DupliquerPaie(PrefixeTable: string; FicheOrigine: TForm);
procedure AccesFicheDupliquer(Ecran: TForm; Predefini: string; var NoDossier: string; var LectureSeule: Boolean);
procedure RecupControl(ControlParent: TWinControl);
function RechEnrAssocier(NomTable: Hstring; NomChamp: array of Hstring; ValChamp: array of variant): BOOLEAN;
{$IFDEF EAGLCLIENT}
function TestValLookupChamp(ControlChamp: THEdit; ValChamp: string): integer;
{$ELSE}
function TestValLookupChamp(ControlChamp: THDBEdit; ValChamp: string): integer;
{$ENDIF}
{$IFDEF EAGLCLIENT}
procedure ReinitMoisValidite(Control1, Control2, Control3: THValComboBox);
{$ELSE}
procedure ReinitMoisValidite(Control1, Control2, Control3: THDBValComboBox);
{$ENDIF}
{$IFDEF EAGLCLIENT}
function ControleDateCourte(MZone4: thedit; Bdd_zone: string; BlancAutorise: boolean): Boolean;
{$ELSE}
function ControleDateCourte(MZone4: thdbedit; Bdd_zone: string; BlancAutorise: boolean): Boolean;
{$ENDIF}
//PT58 function RendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
//PT58 function RendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
function CodePair(Code: string): Boolean;
procedure DonneCodeDupliquer(Ecode, EPred, EDate, EConv: string); //PT81
procedure DonneCodeDupliquerDos(Ecode, ENiveau, EValNiveau, EDate: string); //PT80
procedure DonneCodeDupliquerEtab(Ecode, ELib: string); //PT89
function IncrementeSeqNoOrdre(cle_type, cle_sal: string): Integer;
function PGZeroAGauche(St: string; L: integer): string;
function RendBornesExerSocial(var DateDebut, DateFin: TDateTime; ComboExer: string): Boolean;
procedure RendTrimestreEnCours(MoisEnCours, DebExer, FinExer: TDateTime; var DebTrim, FinTrim, DebSem, FinSem: TDateTime);
function IsPeriodeActif(DD, DF: TdateTime): boolean; //PT30
function ControlMoisAnneeExer(Mois, Annee: string; var AnneeOk: string): boolean;
function TestRubrique(Predefini, rubrique: string; MaxRub: integer; TrtCum: Boolean = FALSE): boolean;
function TestRubriqueCegStd(Predefini, rubrique: string; MaxRub: integer): boolean;
function MesTestRubrique(Menu, Predefini: string; MaxRub: Integer): string;
procedure PaieLectureSeule(Fiche: TForm; Interdit: Boolean);
procedure AccesPredefini(Predefini: string; var CEG, STD, DOS: Boolean);
function RendTabEltNationaux(Edit: TControl): Boolean;
function DecodeRefPaieAnal(Rubrique, Nature, RefA: string): Boolean;
function PreVentileLignePaie(TOB_VenRem, TOB_VenCot, TOB_Rub: TOB; Salarie, TypeVentil: string; DateDebut, DateFin: TDateTime; AvecRaz: Boolean = TRUE): TOB;
function PreChargeVentileLignePaie(Salarie: string; DateDebut, DateFin: TDateTime): TOB;
function EncodeRefPaie(Salarie: string; DateDebut, DateFin: TDateTime; LaTOB_Rub: TOB): string;
procedure ControlAffecteAnal(Salarie: string; DateDebut, DateFin: TDateTime; TOB_VenRem, TOB_VenCot, TOB_Rubrique, TOB_Anal: TOB);
procedure PGRempliDefaut;
function ConvertiFichierVirement(NomFichier, Support: string): string;
function AffectVentilSalarie(TSal, TAxes, TSection: TOB; CreationSection: Boolean): Boolean;
procedure RendSectionSalarie(TSal: TOB; var Sect: array of string);
function PgCreationSection(TSection: TOB; Sect: array of string; CreationSection: Boolean): Boolean;
procedure PGCompleteSection(var Sect: string; Axe: string);
procedure PGReaffSectionAnal(TOB_Rub, TOBAna: tob; LDDebut: TDateTime = 0; LDFin: TDateTime = 0);
//Procedure PGReaffSectionAnal (TOB_Rub, TOBAna : tob);
function SauvegardeAnal(TOBAna, TOB_Rubrique, TOB_VentilRem, TOB_VentilCot: TOB; Comment: Boolean; LeSal: string; DateD, DateF: TDateTime): Boolean;
function RendComptePreVentil(var NatureRub, Rubrique, Sens: string; TOB_VentilRem, TOB_VentilCot, UneTOB: TOB): string;
procedure VisibiliteOrganisation(Ecran: TForm);
function RendExerciceCorrespondant(DateFin: TDateTime): string;
function RendStringListSal(Prefixe: string; ListSal: TStringList): string;

//Mise à jour
procedure MajConvention;
procedure MajRubrique;
procedure MAJAttestations;
procedure MAJCotisation;
//procedure MAJBaseCPAcquis(Verrouillage : Boolean; etab, sal : string);   //PT13
procedure PgModifNomChamp;
//   PT19 11/06/2002 V582 PH Fonction qui les périodes de la saisie des elts variables
function RendPeriodeSaisVar(var ExerPerEncours, DebPer, FinPer: string): Boolean;
//PT25  : 17/09/2002 V585 PH Optimisation Ventilation analytique salarié non rechargées systématiquement à chaque rubrique
procedure VideVentilAnaSalarie;
// PT27  : 10/10/2002 V585 PH recherche de tous les champs composant la clé des ventilations comptables
procedure RendRefPaieAnal(var Sal, Rub, Nat: string; var DatD, DatF: TDateTime; RefA: string);
function EncodeRefPaieAlpha(Salarie, Nat, Rub: string; DateDebut, DateFin: TDateTime): string;
// PT32  : 09/01/2003 PH V591 Gestion dynamique des modules de la paie en fonction des serias
procedure ChargeModules_PAIE();
function PGIsPeriodeClot(Tob_Exercice: TOB; DatePer: TDateTime): Boolean;
//PT34  : 02/04/2003 PH V_421 Fonction controle de cohérence et recup modif sur NODOSSIER
procedure ReparationNoDossier();
procedure ReaffectationNoDossier(); // PT76
//PT40  : 04/09/2003 PH V_421 Chgt suffixe tablettes CODEEMPLOI dans le cas passage 2003
procedure ChangePCS82A03();
// PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
procedure CompleteVentileLignePaie(LaTOB_Rub, TOBAnalytique, AnaTOB_VenSal: TOB; Salarie: string; DateDebut, DateFin: TDateTime; NumAxe: Integer);
//Copie de la fonction GetControl de UTOF
function PGGetControl(ControlName: string; Ecran: TForm): TControl;
procedure PgGetZoomPaieEdt(Qui: integer; Quoi: string);
function RecupClauseHabilitationLookupList(WhereRenseigne: string): string; //PT79
procedure ReaffNoDossierMulti (); // PT84
procedure SynchroniseTiers (CodeTiers : string);
  Function RendDernierMoisClos (EtatCloture : String; Decalage : Boolean) : String; //PT90
  Procedure GereAccesPredefini (Combo : THValComboBox); //PT94
{$IFNDEF EAGLSERVER}
  Function ChargeDossiersDuRegroupement (CodeReg : String; Grille : THGrid; Mode, AnneeRef, Selection : String) : Boolean; //PT92
  Procedure LanceApplication (var ProcessInfo : TProcessInformation; Dossier, Element : String; GereMessages : Boolean); //PT92
{$ENDIF}
{$ENDIF} //PLUGIN
  function  GererCritereGroupeConfTous (FiltreDonnees:string='') : String;  //PT19  //TEMPO

{$IFNDEF PLUGIN}
var
  PGDateDupliquer, PGCodeDupliquer, PGCodePredefini, PGConvDupliquer: string; //PT81
  PGNiveauDupliquer, PGValNiveauDupliquer: string; //PT80
  PGLibDupliquer : string;      //PT89
  PGModeEdition: string; //PT82
  AnaTOB_VenSal, AnaTOB_VenSalMois, AnaTOB_VenSalRub, AnaTOB_VSalRubImp: TOB; // PT53
  cle: integer;
  MessageInterApp: UINT; //PT91
{$ENDIF}

implementation
{$IFNDEF PLUGIN}
uses
  EntPaie,
  Ent1,
// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}
  P5Util,
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
{$IFDEF COMPTAPAIE}
  UTofPG_GeneCompta,
{$ENDIF}
  P5Def,
  Utilpgi,
  StrUtils;
{$ENDIF}

// PT12 21/01/02 V571 PH Réaffectation des dates début et fin dans YVA_IDENTIFIANT en création de bulletinfunction RendIdentAna (RF : String; LDDebut, LDFin : TDateTime) : string;

{$IFNDEF PLUGIN}

function RendIdentAna(RF: string; LDDebut, LDFin: TDateTime): string;
var
  LSal, DD, DF, Nat, Rub: string;
begin
  LSal := ReadTokenSt(RF);
  DD := ReadTokenSt(RF);
  DF := ReadTokenSt(RF);
  Nat := ReadTokenSt(RF);
  Rub := ReadTokenSt(RF);
  Result := LSal + ';' // Code Salarie
    + FormatDateTime('ddmmyyyy', LDDebut) + ';' // Date Debut bulletin
    + FormatDateTime('ddmmyyyy', LDFin) + ';' // Date Fin bulletin
    + Nat + ';' // Nature Rubrique REM ou COT
    + Rub; // Code rubrique
end;
// FIN PT12

procedure DonneCodeDupliquer(Ecode, EPred, EDate, EConv: string);
begin
  PGCodeDupliquer := Ecode;
  PGCodePredefini := EPred;
  PGDateDupliquer := EDate;
  PGConvDupliquer := EConv; //PT81
end;

//DEB PT80

procedure DonneCodeDupliquerDos(Ecode, ENiveau, EValNiveau, EDate: string);
begin
  PGCodeDupliquer := Ecode;
  PGNiveauDupliquer := ENiveau;
  PGValNiveauDupliquer := EValNiveau;
  PGDateDupliquer := EDate;
end;
//FIN PT80

//DEB PT89
procedure DonneCodeDupliquerEtab(Ecode, ELib: string);
begin
  PGCodeDupliquer := Ecode;
  PGLibDupliquer := ELib;
end;
//FIN PT89

function CodePair(Code: string): Boolean;
var
  LgCode: integer;
begin
  result := FALSE;
  //if V_PGI.Debug = TRUE then exit;
  if Code <> '' then
  begin
    LgCode := length(Code);
    if (Code[LgCode] = '0') or (Code[LgCode] = '2') or (Code[LgCode] = '4') or (Code[LgCode] = '6') or (Code[LgCode] = '8') then
      result := True
    else
      result := False;
  end;
end;

{ fonction renvoyant true si un code existe dans une autre table
Utilisé Utom Cotisation et Utom Rémunération...
PT2 27/06/2001 V536 Utilisation de la fonction générique AGL PresenceComplexe por tester la présence
physique de VALEURS dans différentes tables afin d'empecher la suppression si le champ est utilisé dans
d'autre(s) table(s).
Exemple : Pouvoir supprimer une cotisation si elle n'est pas utilisée dans les historiques, les rubriqueProfil ...
}

function RechEnrAssocier(NomTable: Hstring; NomChamp: array of Hstring; ValChamp: array of variant): BOOLEAN;
var
  Nbre, j: Integer;
  LesTypes, LesComp, LesVal: array[0..9] of Hstring;
  Montableau: array of Hstring;
begin
  result := True;
  // DEB PT57
  j := 0;
  for nbre := Low(NomChamp) to High(NomChamp) do
  begin
    if NomChamp[Nbre] <> '' then j := j + 1;
  end;
  SetLength(Montableau, j);
  for nbre := Low(NomChamp) to j - 1 do
  begin
    Montableau[Nbre] := NomChamp[Nbre];
  end;
  // FIN PT57
  for Nbre := 0 to 9 do
  begin
    LesComp[Nbre] := '';
    LesTypes[Nbre] := '';
  end;
  if (Low(NomChamp) < 0) or (High(NomChamp) > 9) then
  begin
    PgiBox('Incohérence de test de suppression sur la table ' + NomTable, 'Abandon de la suppression');
    exit;
  end;
  for nbre := Low(NomChamp) to High(NomChamp) do
  begin // Affectation par défaut car on traite que des strings et l'égalité comme opérateur
    if NomChamp[Nbre] = '' then break;
    LesComp[Nbre] := '=';
    LesTypes[Nbre] := 'S';
  end; // Conversion du tableau de variant en string car auparavant utilisation de tob avec tableau de variant
  for nbre := Low(ValChamp) to High(ValChamp) do
    Lesval[Nbre] := VarToStr(ValChamp[Nbre]);

  result := (PresenceComplexe(NomTable, MonTableau, LesComp, LesVal, LesTypes)); // PT57
end;


// Fonction de duplication d'un enregistrement dans une table

procedure DupliquerPaie(PrefixeTable: string; FicheOrigine: TForm);
var
  TOB_TablePaie: TOB;
  NomTable: string;
begin
  // PT29  : 16/12/2002 PH V591 Fonction dupliquerpaie rendue compatible CWAS
  if Strlen(PChar(PrefixeTable)) <= 3 then NomTable := PrefixeToTable(PrefixeTable)
  else NomTable := PrefixeTable;
  if (NomTable <> '') then
  begin
    TOB_TablePaie := TOB.Create(NomTable, nil, -1);
    TOB_TablePaie.GetEcran(FicheOrigine, nil);
    //**********Générer un nouvel enregistrement
{$IFNDEF EAGLSERVER}
    if (NomTable = 'JEUECRPAIE') or (NomTable = 'ETABLISS') then    //PT89
      TFFicheListe(FicheOrigine).Bouge(nbInsert)
    else
      TFFiche(FicheOrigine).Bouge(nbInsert);
{$ENDIF}
    //**********Affecter les valeurs Récup
    TOB_TablePaie.PutEcran(FicheOrigine, nil);
    TOB_TablePaie.free;
  end;
end;

procedure AccesFicheDupliquer(Ecran: TForm; Predefini: string; var NoDossier: string; var LectureSeule: Boolean);
var
  CEG, STD, DOS: Boolean;
begin
  AccesPredefini('TOUS', CEG, STD, DOS);
  if Predefini = 'CEG' then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(Ecran, (CEG = False));
    NoDossier := '000000';
  end;
  if Predefini = 'STD' then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(Ecran, (STD = False));
    NoDossier := '000000';
  end;
  if Predefini = 'DOS' then
  begin
    LectureSeule := False;
    PaieLectureSeule(Ecran, False);
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    NoDossier := PgRendNoDossier();
  end;
end;


//procédure affectant Null aux champs de l'onglet non visible (CALCUL OU TRANCHE/REGUL)

procedure RecupControl(ControlParent: TWinControl);
var
  i: integer;
  ChampTrouve: string;
  ChampControl: TControl;
{$IFDEF EAGLCLIENT}
  Combo: THValComboBox;
  Edit: THEdit;
  Check: TCheckBox;
  Spin: TSpinEdit;
{$ELSE}
  Combo: THDBValComboBox;
  Edit: THDBEdit;
  Check: TDBCheckBox;
  Spin: THDBSpinEdit;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  for i := 0 to ControlParent.ControlCount - 1 do
  begin
    ChampControl := ControlParent.Controls[i];
    ChampTrouve := AnsiUpperCase(ChampControl.Name);
    if ChampControl is THValComboBox then
    begin
      Combo := THValComboBox(ChampControl);
      Combo.Value := '';
    end;
    if ChampControl is THEdit then
    begin
      Edit := THEdit(ChampControl);
      Edit.text := '';
    end;
    if ChampControl is TCheckBox then
    begin
      Check := TCheckBox(ChampControl);
      Check.Checked := False;
    end;
    if ChampControl is TSpinEdit then
    begin
      Spin := TSpinEdit(ChampControl);
      Spin.Value := 2;
    end;
  end;
{$ELSE}
  for i := 0 to ControlParent.ControlCount - 1 do
  begin
    ChampControl := ControlParent.Controls[i];
    ChampTrouve := AnsiUpperCase(ChampControl.Name);
    if ChampControl is THDBValComboBox then
    begin
      Combo := THDBValComboBox(ChampControl);
      Combo.Value := '';
    end;
    if ChampControl is THDBEdit then
    begin
      Edit := THDBEdit(ChampControl);
      Edit.text := '';
    end;
    if ChampControl is TDBCheckBox then
    begin
      Check := TDBCheckBox(ChampControl);
      Check.Checked := False;
    end;
    if ChampControl is THDBSpinEdit then
    begin
      Spin := THDBSpinEdit(ChampControl);
      Spin.Value := 2;
    end;
  end;
{$ENDIF}
end;

//******Test sur Valeur Saisie Des THDBEdit (UTOM COTISATION ET REMUNERATION)
//      par rapport au Contenu des tablettes associées
{$IFDEF EAGLCLIENT}

function TestValLookupChamp(ControlChamp: THEdit; ValChamp: string): integer;
{$ELSE}

function TestValLookupChamp(ControlChamp: THDBEdit; ValChamp: string): integer;
{$ENDIF$}
var
  DataTypeName, ValLibelle: string;
  Error: integer;
begin
  Error := 0;
  DataTypeName := '';
  DataTypeName := ControlChamp.DataType;
  if (DataTypeName <> '') and (ValChamp <> '') then
  begin
    ValLibelle := RechDom(DataTypeName, ValChamp, FALSE);
    if (ValLibelle = '') or (ValLibelle = 'Error') then Error := 1; //**si champs n'existe pas
  end;
  result := Error;
end;


//procedure réinitialisant les 4 champs Mois (Onglet Validité) si un prend la valeur <<Aucun>>
//Utilisé dans Utom Cotisation et Rémunération
{$IFDEF EAGLCLIENT}

procedure ReinitMoisValidite(Control1, Control2, Control3: THValComboBox);
{$ELSE}

procedure ReinitMoisValidite(Control1, Control2, Control3: THDBValComboBox);
{$ENDIF}
var
  ValExist: Boolean;
  reponse: Word;
  Vide: string;
begin
  vide := '';
  ValExist := False;
  if (Control1 <> nil) and (Control1.Value <> '') then ValExist := True;
  if (Control2 <> nil) and (Control2.Value <> '') then ValExist := True;
  if (Control3 <> nil) and (Control3.Value <> '') then ValExist := True;

  if ValExist = True then
  begin
    reponse := HShowMessage('5;Validité :;Voulez-vous réinitialiser tous les champs intitulés Mois!;' +
      'Q;YN;Y;N;;;', '', '');
    if reponse = 6 then //Oui
    begin
      Control1.value := Vide;
      Control2.value := Vide;
      Control3.value := Vide;
    end;
  end;

end;


//------------------------------------------------------------------------------
// Controle que la date saisie en jj/mm est valide
// on ajoute une année 00 (pour 2000, parce que bissextile)
// et on passe la date dans le isvaliddatetime.
//------------------------------------------------------------------------------
{$IFDEF EAGLCLIENT}
// PT54 procédure devient fonction pour avoir un code retour de la fonction sur la validité de la zône

function ControleDateCourte(MZone4: thedit; Bdd_zone: string; BlancAutorise: boolean): Boolean;
{$ELSE}

function ControleDateCourte(MZone4: thdbedit; Bdd_zone: string; BlancAutorise: boolean): Boolean;
{$ENDIF}
var
  jj, mm: string;
  MResult: boolean;
begin //PORTAGE CWAS
  result := TRUE;
  if MZone4 = nil then exit;
  if not MZone4.enabled then exit;
  mm := Copy(MZone4.text, 3, 2);
  jj := Copy(MZone4.text, 1, 2);
  // PT36  : 23/06/2003 PH V_421 correction erreur de focalisation fenetre d'une variable CEGID en modif autorisée
  if (((jj = '') or (jj = '  ')) and ((mm = '') or (mm = '  ')) and BlancAutorise) then exit;
  Mresult := IsValidDate(jj + DateSeparator + mm + DateSeparator + '00');
  if not Mresult then
  begin
    result := FALSE;
    MZone4.SetFocus;
    ShowMessage('La date doit avoir le format jj/mm et être valide');
  end
end;
// FIN PT54
{ Fonction qui rend les periodes en cours du dernier exercice actif
}

(* PT58 function RendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
var
  Q: TQuery;
begin
  result := FALSE;
  DebPer := DateToStr(idate1900);
  FinPer := DateToStr(idate1900);
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DebPer := Q.FindField('PEX_DEBUTPERIODE').AsString;
    FinPer := Q.FindField('PEX_FINPERIODE').AsString;
    ExerPerEncours := Q.FindField('PEX_EXERCICE').AsString;
    result := TRUE;
    //  PT1
  end;
  Ferme(Q);
end;   *)

{  PT19 11/06/2002 V582 PH Fonction qui les périodes de la saisie des elts variables
Fonction qui rend les periodes saisie des elts variables du dernier exercice actif
}

function RendPeriodeSaisVar(var ExerPerEncours, DebPer, FinPer: string): Boolean;
var
  Q: TQuery;
begin
  result := FALSE;
  DebPer := DateToStr(idate1900);
  FinPer := DateToStr(idate1900);
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DebPer := Q.FindField('PEX_DEBUTSAISVAR').AsString;
    FinPer := Q.FindField('PEX_FINSAISVAR').AsString;
    ExerPerEncours := Q.FindField('PEX_EXERCICE').AsString;
    result := TRUE;
    //  PT1
  end;
  Ferme(Q);
end;
{ Fonction qui rend l'année et le mois de l'exercice en cours
Elle donne en fait le dernier exercice Actif
}

(* PT58 function RendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
var
  Q: TQuery;
  DatF: TDateTime;
  Jour, Mois, Annee: WORD;
begin
  result := FALSE;
  DebExer := idate1900;
  FinExer := idate1900;
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DatF := Q.FindField('PEX_FINPERIODE').AsFloat; //Q.Fields[7].AsFloat; // Recup date de fin periode en cours
    DecodeDate(DatF, Annee, Mois, Jour);
    MoisE := ColleZeroDevant(Mois, 2);
    ComboExer := Q.FindField('PEX_EXERCICE').AsString; //Q.Fields[0].AsString; // recup Combo identifiant exercice
    AnneeE := Q.FindField('PEX_ANNEEREFER').AsString; //Q.Fields[8].AsString; // recup Annee de exercice
    DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
    FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
    result := TRUE;
    // PT1
  end;
  Ferme(Q);
end;    *)


{ Fonction qui rend la date de debut et la date de fin de l'exercice en cours en passant l'exercice
Elle donne en fait le dernier exercice Actif
}

function RendBornesExerSocial(var DateDebut, DateFin: TDateTime; ComboExer: string): Boolean;
var
  Q: TQuery;
begin
  result := FALSE;
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" AND PEX_EXERCICE="' + ComboExer + '" GROUP BY PEX_DATEDEBUT DESC', TRUE);
  if not Q.EOF then
  begin
    DateDebut := Q.FindField('PEX_DATEDEBUT').AsFloat; //Q.Fields[2].AsFloat; // Recup date de DEBUT exercice social
    DateFin := Q.FindField('PEX_DATEFIN').AsFloat; //Q.Fields[3].AsFloat; // Recup date de fin exercice social
    result := TRUE;
    // PT1
  end;
  Ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : SB & MF
Créé le ...... : 25/07/2001
Modifié le ... : 25/07/2001
Description .. : procedure qui rend la date de debut et de fin du trimestre en
Suite ........ : cours, de début et de fin du semestre en cours de l'exercice
Suite ........ : en cours.
Mots clefs ... : PAIE, PGDUCS
*****************************************************************}

procedure RendTrimestreEnCours(MoisEnCours, DebExer, FinExer: TDateTime; var DebTrim, FinTrim, DebSem, FinSem: TDateTime);
var
  IntMois: Integer;
  YY, JJ, MM: WORD;
  NumTrim: Extended;
begin //PORTAGE CWAS
  DebTrim := MoisEncours;
  FinTrim := FindeMois(MoisEnCours);
  DecodeDate(MoisEnCours, YY, MM, JJ);
  IntMois := MM;
  // Récupère numero de trimestre
  if (IntMois = 12) and (VH_PAIE.PGDecalage = True) then NumTrim := 1
  else
  begin
    if VH_PAIE.PGDecalage = False then IntMois := IntMois - 1;
    NumTrim := Int(IntMois / 3);
    NumTrim := NumTrim + 1;
  end;
  //Récupère date debut et fin de trimestre
  if NumTrim = 1 then
  begin
    DebTrim := DebExer;
    FinTrim := FindeMois(PlusMois(DebTrim, 2));
  end;
  if NumTrim = 2 then
  begin
    DebTrim := PlusMois(DebExer, 3);
    FinTrim := FindeMois(PlusMois(DebTrim, 2));
  end;
  if NumTrim = 3 then
  begin
    DebTrim := PlusMois(DebExer, 6);
    FinTrim := FindeMois(PlusMois(DebTrim, 2));
  end;
  if NumTrim = 4 then
  begin
    DebTrim := PlusMois(DebExer, 9);
    FinTrim := FindeMois(PlusMois(DebTrim, 2));
  end;

  //Récupère date debut et fin de Semestre
  if (NumTrim = 1) or (NumTrim = 2) then
  begin
    DebSem := DebExer;
    FinSem := FindeMois(PlusMois(DebSem, 6));
  end;
  if (NumTrim = 3) or (NumTrim = 4) then
  begin
    DebSem := PlusMois(DebExer, 5);
    FinSem := FindeMois(PlusMois(DebSem, 6));
  end;
end;
//PT30 nouvelle function
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 30/12/2002
Modifié le ... : 30/12/2002
Description .. : Function qui test si la période est comprise dans un
Suite ........ : exercice social actif
Mots clefs ... : PAIE,EXERCICE
*****************************************************************}

function IsPeriodeActif(DD, DF: TdateTime): boolean;
begin
  result := ExisteSql('SELECT PEX_EXERCICE FROM EXERSOCIAL  ' +
    'WHERE PEX_DATEDEBUT<="' + UsDateTime(DD) + '" ' +
    'AND PEX_DATEFIN>="' + UsDateTime(DF) + '" ' +
    'AND PEX_ACTIF="X" ');
  if Result = False then
    PGIBox('Le mois et l''année ne correspondent pas à un exercice social actif.', 'Exercice Social');
end;

//------------------------------------------------------------------------------
// fonction d'incrément du numéro d'ordre pour la table
// ABSENCESALARIE
//------------------------------------------------------------------------------

function IncrementeSeqNoOrdre(cle_type, cle_sal: string): Integer;
var
  q: TQuery;
  StWhere: string;
begin
  result := 1;
  if cle_type <> '' then StWhere := 'AND PCN_TYPEMVT = "' + cle_type + '" ' else StWhere := '';
  q := OpenSQL('select max (PCN_ORDRE) as maxno from ABSENCESALARIE ' +
    ' where PCN_SALARIE = "' + cle_sal + '" ' + StWhere, TRUE);
  //cle :  PCN_TYPEMVT,PCN_SALARIE,PCN_DATEDEBUT,PCN_DATEFIN,PCN_ORDRE
  if not q.eof then
{$IFDEF EAGLCLIENT}
    if q.Fields[0].AsInteger <> 0 then
{$ELSE}
    if not q.Fields[0].IsNull then
{$ENDIF}
      result := ((q.Fields[0]).AsInteger + 1);

  Ferme(q);
end;
// Complete une chaine de caractere par des 0 à gauche

function PGZeroAGauche(St: string; L: integer): string;
begin
  St := Copy(St, 1, L);
  while Length(St) < L do St := '0' + St;
  Result := St;
end;

function ControlMoisAnneeExer(Mois, Annee: string; var AnneeOk: string): boolean;
var
  st: string;
  DebMois, FinMois: TDateTime;
  Q: TQuery;
  Okok: Boolean;
begin
  Okok := False;
  AnneeOk := Annee;
  //  PT11 12/12/01 V562 PH Controle du mois renseigné car cela provoque une erreur de conversion
  if Mois = '' then Mois := '01';
  if AnneeOk = '' then
  begin
    OkOk := False;
    AnneeOk := '1900';
  end;
  if (VH_Paie.PGDecalage = TRUE) and (AnneeOk <> '') then
  begin
    if Mois = '12' then AnneeOk := IntToStr(StrToInt(Annee) - 1)
    else AnneeOk := Annee;
  end;
  if (AnneeOk <> '') then
  begin
    Debmois := EncodeDate(StrToInt(AnneeOk), StrToInt(Mois), 01);
    FinMois := FINDEMOIS(DebMois);
    st := 'Select PEX_DATEDEBUT,PEX_DATEFIN from EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC';
    Q := OpenSql(st, TRUE);
    Okok := FALSE;
    // Boucle de recherche dans les exercices pour identifier dans quel exercice se trouve une date maxi 10 enreg
    // si 10 de paie en ligne. Comme ordre décroissant alors 1 seule boucle dans la majorité des cas
    while not Q.EOF do
    begin
      if (DebMois >= Q.FindField('PEX_DATEDEBUT').AsDateTime) and (FINMois <= Q.FindField('PEX_DATEFIN').AsDateTime) then
      begin
        Okok := TRUE;
        break;
      end;
      Q.Next;
    end;
    Ferme(Q);
  end;
  if Okok = FALSE then
  begin
    PGIBox('Le mois et l''année ne correspondent pas à un exercice social actif.', 'Exercice Social');
  end;
  result := Okok;
end;
{Dans tous les cas les nombres pairs sont prédéfini CEGID
Dans le cas des nombres impairs :
  - Rubriques de cotisations ou remunération
       - Termine 1 ou 3 : prédéfini Standard
       - Termine 5,7 ou 9 : prédéfini Dossier
  - Cumul
       - 50 premier nombre impair sont prédéfini CEGID ( plus ts les nb pairs)
       - Au dela :
           - Termine 1 ou 3 : prédéfini Standard
           - Termine 5,7 ou 9 : prédéfini Dossier
   - Elt Nationaux et Variable
       - 1000 premier nombre impair sont prédéfini CEGID ( plus ts les nb pairs)
       - Au dela :
           - Termine 1 ou 3 : prédéfini Standard
           - Termine 5,7 ou 9 : prédéfini Dossier sauf Elt national
           }
//  PT15 10/04/2002 V571 Ph Traitement des cumuls si pair et > 50 et STD alors Ok

function TestRubrique(Predefini, rubrique: string; MaxRub: integer; TrtCum: Boolean = FALSE): boolean;
var
  Pair: Boolean;
  Lg: integer;
begin
  result := False;
  Pair := CodePair(rubrique);
  if Pair = true then
  begin
    if Predefini = 'CEG' then result := True;
    if Predefini = 'STD' then result := False;
    if Predefini = 'DOS' then result := False;
  end;
  if (Pair = False) and (MaxRub = 0) then // 0 : rubriques de cotisations et de rémunérations
  begin
    Lg := length(rubrique);
    if Predefini = 'CEG' then result := False;
    if (Predefini = 'STD') and ((Rubrique[Lg] = '1') or (Rubrique[Lg] = '3')) then result := True;
    if (Predefini = 'DOS') and ((Rubrique[Lg] = '5') or (Rubrique[Lg] = '7') or (Rubrique[Lg] = '9')) then result := True;
  end;
  //  PT15 10/04/2002 V571 Ph Traitement des cumuls si pair et > 50 et STD alors Ok
  if TrtCum and (StrToInt(Rubrique) > MaxRub) and (Predefini = 'STD') and (Pair = true) then
  begin
    result := True;
    exit;
  end;
  if (Pair = False) and (MaxRub <> 0) then
  begin
    if (StrToInt(Rubrique) <= MaxRub) and (Predefini = 'CEG') then result := True;
    if (StrToInt(Rubrique) <= MaxRub) and (Predefini <> 'CEG') then result := False;
    if (StrToInt(Rubrique) > MaxRub) and (Predefini = 'CEG') then result := False;
    if (StrToInt(Rubrique) > MaxRub) and (Predefini <> 'CEG') then
    begin
      Lg := length(rubrique);
      if (Predefini = 'STD') and ((Rubrique[Lg] = '1') or (Rubrique[Lg] = '3')) then result := True;
      if (Predefini = 'DOS') and ((Rubrique[Lg] = '5') or (Rubrique[Lg] = '7') or (Rubrique[Lg] = '9')) then result := True;
      if (Predefini = 'DOS') and ((Rubrique[Lg] = '1') or (Rubrique[Lg] = '3')) then result := False;
      if (Predefini = 'STD') and ((Rubrique[Lg] = '5') or (Rubrique[Lg] = '7') or (Rubrique[Lg] = '9')) then result := False;
    end;
  end;
end;

function TestRubriqueCegStd(Predefini, rubrique: string; MaxRub: integer): boolean;
var
  Pair: Boolean;
  Lg: integer;
begin
  result := False;
  Pair := CodePair(rubrique);
  Lg := length(rubrique);
  if Pair = true then
  begin
    if Predefini = 'CEG' then result := True;
    if Predefini = 'STD' then result := False;
  end;
  if (Pair = False) and (MaxRub = 0) then // 0 : conventions collectives
  begin
    if Predefini = 'CEG' then result := False;
    if (Predefini = 'STD') and ((Rubrique[Lg] = '1') or (Rubrique[Lg] = '3') or (Rubrique[Lg] = '5') or (Rubrique[Lg] = '7') or (Rubrique[Lg] = '9')) then result := True;
  end;

  if (Pair = False) and (MaxRub <> 0) then
  begin
    if (StrToInt(Rubrique) <= MaxRub) and (Predefini = 'CEG') then result := True;
    if (StrToInt(Rubrique) <= MaxRub) and (Predefini <> 'CEG') then result := False;
    if (StrToInt(Rubrique) > MaxRub) and (Predefini = 'CEG') then result := False;
    if (StrToInt(Rubrique) > MaxRub) and (Predefini <> 'CEG') then
    begin
      if (Predefini = 'STD') and ((Rubrique[Lg] = '1') or (Rubrique[Lg] = '3') or (Rubrique[Lg] = '5') or (Rubrique[Lg] = '7') or (Rubrique[Lg] = '9')) then result := True;
      if (Predefini = 'STD') and ((Rubrique[Lg] = '0') or (Rubrique[Lg] = '2') or (Rubrique[Lg] = '4') or (Rubrique[Lg] = '6') or (Rubrique[Lg] = '8')) then result := False;
{$IFDEF CPS1} // PT83
      if (StrToInt(Rubrique) > MaxRub) and (Predefini = 'DOS') then result := true;
{$ENDIF}
    end;
  end;
end;

function MesTestRubrique(Menu, Predefini: string; MaxRub: Integer): string;
var
  Mes, MesSup: string;
begin
  mes := 'Vous devez renseigner un code';
  if maxrub = 0 then MesSup := '';
  if maxrub > 0 then MesSup := ' et supérieur à ' + IntToStr(MaxRub) + ' '; //PT35 Orthographe
  if (Predefini = 'CEG') then //and (Menu<>'ELT') then
  begin
    if maxrub > 0 then mes := mes + ' inférieur à ' + IntToStr(MaxRub) + ' ou pair'; //PT35 Orthographe
    if maxrub = 0 then mes := mes + ' pair';
  end;
  if (Menu = 'ELT') and (Predefini = 'STD') then
  begin
    mes := mes + ' impair' + MesSup;
    Result := mes;
    exit;
  end;
  if (Menu = 'CONV') and (Predefini = 'STD') then
  begin
    mes := mes + ' impair';
    Result := mes;
    exit;
  end;
  if Predefini = 'STD' then mes := mes + ' impair terminant par 1 ou 3 ' + MesSup;
  if Predefini = 'DOS' then mes := mes + ' impair terminant par 5,7 ou 9' + MesSup;
  if mes = '' then mes := 'Votre code est incorrect';
  Result := mes + '.'; //PT35 Ajout .
end;



procedure PaieLectureSeule(Fiche: TForm; Interdit: Boolean);
var
  i: integer;
  ActionNormale, ActionInverse: Boolean;
begin
  ActionNormale := not Interdit;
  ActionInverse := Interdit;
  with Fiche do
  begin
    for i := 0 to ComponentCount - 1 do
    begin
      if (Components[i] is TBitBtn) or (Components[i] is TToolbarButton97) then
      begin
        if ((UpperCase(Components[i].Name) = 'BINSERT') or
          (UpperCase(Components[i].Name) = 'BDELETE') or
          (UpperCase(Components[i].Name) = 'BANNULER') or
          (UpperCase(Components[i].Name) = 'BVALIDER')) then
        begin
          if Components[i] is TBitBtn then TBitBtn(Components[i]).Enabled := ActionNormale else
            if Components[i] is TToolbarButton97 then TToolbarButton97(Components[i]).Enabled := ActionNormale;
{$IFDEF EAGLCLIENT}
          if UpperCase(Components[i].Name) = 'BDELETE' then TToolbarButton97(Components[i]).Visible := ActionNormale;
{$ENDIF}
        end;
      end else
{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
        if Components[i] is THDBRichEditOLE then THDBRichEditOLE(Components[i]).Enabled := ActionNormale;
      //   else   if Components[i] is THDBRichEdit then THDBRichEdit(Components[i]).ReadOnly:=TRUE else
      if Components[i] is THDBGrid then THDbGrid(Components[i]).ReadOnly := ActionInverse else
{$ENDIF}
{$ENDIF}
        //      if Components[i] is THRichEditOLE then THRichEditOLE(Components[i]).ReadOnly:=TRUE else

        if Components[i] is TRichEdit then TRichEdit(Components[i]).ReadOnly := ActionInverse else
          if Components[i] is TCustomEdit then TCustomEdit(Components[i]).Enabled := ActionNormale else
            if Components[i] is TCustomComboBox then TCustomComboBox(Components[i]).Enabled := ActionNormale else
              if Components[i] is TCustomCheckBox then TCustomCheckBox(Components[i]).Enabled := ActionNormale else
                if Components[i] is TRadioButton then TRadioButton(Components[i]).Enabled := ActionNormale else
                  if Components[i] is TCustomListBox then TCustomListBox(Components[i]).Enabled := ActionNormale else
                    if Components[i] is TCustomRadioGroup then TCustomRadioGroup(Components[i]).Enabled := ActionNormale else
        // if Components[i] is TCustomGroupBox then TCustomGroupBox(Components[i]).Enabled:=ActionNormale else
        // PT67
{        if Components[i] is TCustomGrid then TCustomGrid(Components[i]).Enabled := ActionNormale else
        if Components[i] is THGrid then
           begin
           if Interdit = TRUE then THGrid(Components[i]).Options := THGrid(Components[i]).Options - [goEditing]
           else THGrid(Components[i]).Options := THGrid(Components[i]).Options + [goEditing];
           end; }
        // PT67
                      if Components[i] is TCustomGrid then
                      begin
                        if Interdit = TRUE then THGrid(Components[i]).Options := THGrid(Components[i]).Options - [goEditing]
                        else THGrid(Components[i]).Options := THGrid(Components[i]).Options + [goEditing];
                      end;
    end;
  end;
end;

procedure AccesPredefini(Predefini: string; var CEG, STD, DOS: Boolean);
var
  Q: TQuery;
begin
  CEG := False;
  STD := False;
  DOS := False;
  //V_PGI.debug si non coché alors Lecture Seule
// PT45  : 10/05/2004 PH V_50 restriction droit accès au paramétrage CEGID (CEG := V_PGI.Debug)
// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}
  if (Predefini = 'CEG') or (Predefini = 'TOUS') then CEG := DroitCegPaie;
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
// DEB PT62
  if (Predefini = 'CEG') then
  begin
{$IFNDEF EAGLSERVER}
    Q := OpenSql('SELECT US_CONTROLEUR FROM UTILISAT WHERE US_UTILISATEUR="' + V_PGI.FUser + '"', True);
    if not Q.EOF then //PORTAGECWAS
      STD := (Q.FindField('US_CONTROLEUR').AsString = 'X');
    Ferme(Q);
{$ENDIF}
  end;
// FIN PT62
  if (Predefini = 'STD') or (Predefini = 'TOUS') then
  begin
{$IFNDEF EAGLSERVER}
    Q := OpenSql('SELECT US_CONTROLEUR FROM UTILISAT WHERE US_UTILISATEUR="' + V_PGI.FUser + '"', True);
    if not Q.EOF then //PORTAGECWAS
      STD := (Q.FindField('US_CONTROLEUR').AsString = 'X');
    if STD then
      STD := JaiLeDroitTag(200101);
    Ferme(Q);
{$ENDIF}
  end;
  if (Predefini = 'DOS') or (Predefini = 'TOUS') then
  begin
    DOS := JaiLeDroitTag(200102);
  end;
// PT75
{$IFDEF CPS1}
  STD := FALSE;
{$ENDIF}
end;

function RendTabEltNationaux(Edit: TControl): Boolean;
var
  Sql: string;
begin
{$IFNDEF EAGLSERVER}
  Sql := 'SELECT DISTINCT PEL_CODEELT,PEL_LIBELLE FROM ELTNATIONAUX ORDER BY PEL_CODEELT';
  result := LookUpList(Edit, 'Element national', 'ELTNATIONAUX', 'PEL_CODEELT,PEL_DATEVALIDITE', 'PEL_LIBELLE', Sql, 'PEL_CODEELT', True, -1);
{$ENDIF}
end;

{ Fonction qui indique par rapport à la ventilation analytique si on est bien
 sur la bonne rubrique correspond à la ligne que l'on traite
}

function DecodeRefPaieAnal(Rubrique, Nature, RefA: string): Boolean;
var
  Sal, DatD, DatF, Nat, Rub: string;
begin
  Result := FALSE;
  Sal := ReadTokenSt(RefA);
  DatD := ReadTokenSt(RefA);
  DatF := ReadTokenSt(RefA);
  Nat := ReadTokenSt(RefA);
  Rub := ReadTokenSt(RefA);
  if (Copy(Nat, 1, 1) = Copy(Nature, 1, 1)) and (Copy(Rub, 1, 4) = Copy(Rubrique, 1, 4)) then result := TRUE; // PT77
end;
// PT27  : 10/10/2002 V585 PH recherche de tous les champs composant la clé des ventilations comptables
{Fonction qui decode la composition de la clé analytique pour retrouver
 tous les champs
}

procedure RendRefPaieAnal(var Sal, Rub, Nat: string; var DatD, DatF: TDateTime; RefA: string);
var
  DD, DF: string;
begin
  Sal := ReadTokenSt(RefA);
  DD := ReadTokenSt(RefA);
  DF := ReadTokenSt(RefA);
  DD := Copy(DD, 1, 2) + '/' + Copy(DD, 3, 2) + '/' + Copy(DD, 5, 4);
  DF := Copy(DF, 1, 2) + '/' + Copy(DF, 3, 2) + '/' + Copy(DF, 5, 4);
  DatD := StrToDate(DD);
  DatF := StrToDate(DF);
  Nat := ReadTokenSt(RefA);
  Rub := ReadTokenSt(RefA);
end;

{ Fonction de constitution de la clé d'acces au ventilation du bulletin
Salarie/Nature/rubrique/DateDebut/Datefin mais à partir de paramètres et non de la TOB_Rub
}

function EncodeRefPaieAlpha(Salarie, Nat, Rub: string; DateDebut, DateFin: TDateTime): string;
begin
  Result := '';
  Result := Salarie + ';' // Code Salarie
    + FormatDateTime('ddmmyyyy', DateDebut) + ';' // Date Debut bulletin
    + FormatDateTime('ddmmyyyy', DateFin) + ';'; // Date Fin bulletin
  if Rub <> '' then result := result + Nat + ';' + Rub // Nature Rubrique REM=AAA ou COT + Rubrique
  else result := result + Nat;
end;
// FIN PT27
{ Fonction de constitution de la clé d'acces au ventilation du bulletin
Salarie/Nature/rubrique/DateDebut/Datefin
Attention les bases de cotisations ne sont pas traitées
}

function EncodeRefPaie(Salarie: string; DateDebut, DateFin: TDateTime; LaTOB_Rub: TOB): string;
begin
  Result := '';
  if LaTOB_Rub = nil then Exit;
  Result := Salarie + ';' // Code Salarie
    + FormatDateTime('ddmmyyyy', DateDebut) + ';' // Date Debut bulletin
    + FormatDateTime('ddmmyyyy', DateFin) + ';' // Date Fin bulletin
    + Copy(LaTOB_Rub.GetValue('PHB_NATURERUB'), 1, 1) + ';' // Nature Rubrique REM=AAA ou COT
    + LaTOB_Rub.GetValue('PHB_RUBRIQUE'); // Code rubrique
end;
// Fonction de memorisation de la TOB des ventilation par defaut du salarié et du bulletin

procedure RendVentilAnaSalarie(TOB_VenSal, TOB_VenSalRub: TOB; Salarie: string);
var
  St: string;
begin
  TOB_VenSal := TOB.Create('Analytique Salarié', nil, -1);
  St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + Salarie + '" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_VenSal.LoadDetailDB('VENTIL', '', '', Q, FALSE);
  ferme(Q);
}
  TOB_VenSal.LoadDetailDBFROMSQL('VENTIL', st);

  // Tob des ventilations analytiques prédéfinies pour le salarié pour des rubriques
  TOB_VenSalRub := TOB.Create('Analytique Salarié Rubrique', nil, -1);
  //   PT17 29/04/2002 V582 PH Optimisation SQL en supprimant like
  St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "PG%" AND V_COMPTE = "' + Salarie + '%" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOB_VenSalRub.LoadDetailDB('VENTIL', '', '', Q, FALSE);
  ferme(Q);
}
  TOB_VenSalRub.LoadDetailDBFromSQL('VENTIL', st);
end;
// fonction qui detruit les TOB des ventilations analytiques par Salarié et par Salaries/rubriques
//PT25  : 17/09/2002 V585 PH Optimisation Ventilation analytique salarié non rechargées systématiquement à chaque rubrique

procedure VideVentilAnaSalarie;
begin
  if AnaTOB_VenSal <> nil then
  begin
    AnaTOB_VenSal.Free;
    AnaTOB_VenSal := nil;
  end;
  if AnaTOB_VenSalRub <> nil then
  begin
    AnaTOB_VenSalRub.Free;
    AnaTOB_VenSalRub := nil;
  end;
  if AnaTOB_VSalRubImp <> nil then
  begin
    AnaTOB_VSalRubImp.Free;
    AnaTOB_VSalRubImp := nil;
  end;
  // DEB PT53
  if AnaTOB_VenSalMois <> nil then
  begin
    AnaTOB_VenSalMois.Free;
    AnaTOB_VenSalMois := nil;
  end;
  // FIN PT53
end;
{ Fonction de preventilation des lignes analytiques
Recup des ventilations par defaut Rubrique,Salarie, Salarié/rubrique
dans la table VENTIL et confection des ventilations du bulletin
dans la table VENTANA.
Les enregs de la table VENTANA serviront à retrouver les sections analytiques
pour la génération des  écritures comptables et pour les lignes analytiques
de la paie.
On regarde si la création des sections est automatique et composée en fonction
des paramètres société
TOB_VenSal : Tob des ventilations du salarié
TOB_VenRem : Tob des ventilations des rémunérations
TOB_VenCot : Tob des ventilations des cotisations  ==> n'a pas de sens à priori
TOB_VenSalRub : Tob des ventilations Salarié/rubrique
TOB_Rub : Tob du bulletin en cours de traitement
TypeVentil : PRE preventilation alors TOB qui contient des preventil de la table VENTIL
             ANA ventilations définitives du bulletin
Cette Fonction parcourt la TOB_Rub qui est la tob des lignes du bulletin pour concevoir
toutes les ventilations analytiques par defaut au niveau du bulletin
Attention : la fonction rend la TOB des ventilations analytiques , il faut donc que la fonction qui l'appelle
detruise cette TOB et sauvegarde les donnees le cas échéant.
Les Ventilations des cotisations dependent de celles des remunerations sauf celles qui peuvent
avoir une ventilation specifique
On creera lors du calcul du bulletin des ventilations de type Sal;DateDebut;DateFin; COT
pour toutes les cotisations. Ces ventilations sont calculees, proratisées en fonction des
montants des rémunérations présentent sur le bulletin. Mais elles dependent surtout de la personnalisation
de la ventilation des rémunérations faites sur le bulletin. Voir fonction PGReaffSectionAnal

PT10 03/10/01 V562 PH Traitement ventilations analytiques provenant d'un import
La prise en compte de ces ventilations se fait uniquement en creation d'un bulletin
et sur une période donnée.
PT25  : 17/09/2002 V585 PH Optimisation Ventilation analytique salarié non rechargées systématiquement à chaque rubrique
        changement des noms des TOB donc refonte des acces et de la liberation des TOB faite par fonction generique
}

function PreVentileLignePaie(TOB_VenRem, TOB_VenCot, TOB_Rub: TOB; Salarie, TypeVentil: string; DateDebut, DateFin: TDateTime; AvecRaz: Boolean = TRUE): TOB;
var
  i: integer;
  RefA, St, VentilTrouv: string;
  Nature, Rubrique: string;
  LaTOB_Rub, TOBAnalytique, TOBA, Tob_PreVentImp, Tob_PreVentil: TOB;
  AxesImp, AxesSal: string;
  T: TOB;
  ii: Integer;
begin
  TOBA := nil;
  TOBAnalytique := TOB.Create('Les ventilations du bulletin', nil, -1);
  if TypeVentil = 'PRE' then LaTOB_Rub := TOB_Rub.FindFirst(['PHB_SALARIE'], [Salarie], TRUE)
  else LaTOB_Rub := TOB_Rub;
  // Tob des ventilations analytiques prédéfinies pour le salarié
  if AnaTOB_VenSal = nil then
  begin
    AnaTOB_VenSal := TOB.Create('Analytique Salarié', nil, -1);
    St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + Salarie + '" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    AnaTOB_VenSal.LoadDetailDB('VENTIL', '', '', Q, FALSE);
    ferme(Q);
}
    AnaTOB_VenSal.LoadDetailDBFromSQL('VENTIL', st);
  end;

  // DEB PT53
  { Nouvelles ventilations type en provenance du fichier d'import : ventilation type uniquement
    valables pour la période de traitement
  }
  if AnaTOB_VenSalMois = nil then
  begin
    AnaTOB_VenSalMois := TOB.Create('Analytique_Salarié_Mois', nil, -1);
    St := 'SELECT * FROM PAIEVENTIL WHERE PAV_NATURE LIKE "VS%" AND PAV_COMPTE = "'
      + Salarie + '" AND PAV_DATEDEBUT >= "' + USDateTime(DateDebut)
      + '" AND  PAV_DATEFIN <= "' + USDateTime(DateFin) + '" ORDER BY PAV_NATURE,PAV_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    AnaTOB_VenSalMois.LoadDetailDB('PAIEVENTIL', '', '', Q, FALSE);
    ferme(Q);
}
    AnaTOB_VenSalMois.LoadDetailDBFROMSQL('PAIEVENTIL', st);
  end;
  // FIN PT53

    // Tob des ventilations analytiques prédéfinies pour le salarié pour des rubriques
  if AnaTOB_VenSalRub = nil then
  begin
    AnaTOB_VenSalRub := TOB.Create('Analytique Salarié Rubrique', nil, -1);
    //   PT17 29/04/2002 V582 PH Optimisation SQL en supprimant like
    // PT31  : 08/01/2003 PH V591 Suppression PT17 car ne prend plus en compte specif sal/rub
    St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "PG%" AND V_COMPTE like "' + Salarie + '%" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    AnaTOB_VenSalRub.LoadDetailDB('VENTIL', '', '', Q, FALSE);
    ferme(Q);
}
    AnaTOB_VenSalRub.LoadDetailDBFROMSQL('VENTIL',st);
  end;
  // PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
  AxesSal := '-----';
  T := AnaTOB_VenSal.FindFirst([''], [''], FALSE);
  while T <> nil do
  begin
    ii := ValeurI(Copy(T.GetValue('V_NATURE'), 3, 1));
    AxesSal[ii] := 'X'; // Recuperation des axes analytiques gérés au niveau salarié
    T := AnaTOB_VenSal.FindNext([''], [''], FALSE);
  end;
  // FIN PT43
  // Tob des ventilations analytiques prédéfinies pour le salarié pour des rubriques provenant d'un import
  if AnaTOB_VSalRubImp = nil then
  begin
    AnaTOB_VSalRubImp := TOB.Create('Anal Salarié Rubrique Import', nil, -1);
    // PT24  : 17/09/2002 V585 PH Optimisation requete sql en remplaçant like par = car valeur exacte
    // PT31  : 08/01/2003 PH V591 Suppression PT17 car ne prend plus en compte specif sal/rub
    St := 'SELECT * FROM PAIEVENTIL WHERE PAV_NATURE LIKE "PG%" AND PAV_COMPTE like "'
      + Salarie + '%" AND PAV_DATEDEBUT >= "' + USDateTime(DateDebut)
      + '" AND  PAV_DATEFIN <= "' + USDateTime(DateFin) + '" ORDER BY PAV_NATURE,PAV_COMPTE';
{Flux optimisé
    Q := OpenSql(st, TRUE);
    AnaTOB_VSalRubImp.LoadDetailDB('PAIEVENTIL', '', '', Q, FALSE);
    ferme(Q);
}
    AnaTOB_VSalRubImp.LoadDetailDBFROMSQL('PAIEVENTIL',st);
  end;
  i := 0;
  while LaTOB_Rub <> nil do
  begin
    Nature := LaTOB_Rub.GetValue('PHB_NATURERUB');
    Rubrique := LaTOB_Rub.GetValue('PHB_RUBRIQUE');
    AxesImp := '-----'; // PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
    //      if VH_Paie.PGSectionAnal = FALSE then // Section analytique non creee automatiquement
    if (Nature <> 'BAS') and (Copy(Rubrique, 5, 1) <> '.') then // car les bases de cotisation ne sont pas ventilées analytiquement
    begin // On elimine aussi les lignes de commentaires
      // preventilation des salaries/remunerations
      VentilTrouv := '';
      Tob_PreVentil := AnaTOB_VenSalRub.findFirst(['V_NATURE', 'V_COMPTE'], ['PG1', Salarie + ';' + Rubrique], TRUE);
      if Tob_PreVentil <> nil then VentilTrouv := 'P';
      // Recherche s'il existe des preventilations provenant d'un import pour le salarie,la rubrique rémunération, la periode
      Tob_PreVentImp := AnaTOB_VSalRubImp.findFirst(['PAV_NATURE', 'PAV_COMPTE'], ['PG1', Salarie + ';' + Rubrique], TRUE);
      if Tob_PreVentImp <> nil then
      begin
        Tob_PreVentil := Tob_PreVentImp;
        VentilTrouv := 'I';
      end;
      if Tob_PreVentil = nil then
      begin
        if Nature = 'AAA' then
        begin // recherche ventilation analytique par defaut pour les remunerations
          Tob_PreVentil := TOB_VenRem.findFirst(['V_NATURE', 'V_COMPTE'], ['RR1', Rubrique], TRUE);
          if Tob_PreVentil <> nil then VentilTrouv := 'R';
        end
        else
        begin // recherche ventilation analytique par defaut pour les cotisations
          Tob_PreVentil := TOB_VenCot.findFirst(['V_NATURE', 'V_COMPTE'], ['RC1', Rubrique], TRUE);
          if Tob_PreVentil <> nil then VentilTrouv := 'C';
        end;
        // DEB PT53
        if Tob_PreVentil = nil then
        begin
          Tob_PreVentil := AnaTOB_VenSalMois.findFirst([''], [''], TRUE);
          if Tob_PreVentil <> nil then VentilTrouv := 'VS'; // Preventilation specifique au salarie pour la période
        end;
        // FIN PT53
        if Tob_PreVentil = nil then
        begin
          Tob_PreVentil := AnaTOB_VenSal.findFirst([''], [''], TRUE);
          if Tob_PreVentil <> nil then VentilTrouv := 'S'; // Preventilation specifique au salarie
        end;
      end;
      //PT38-1    else VentilTrouv := 'P'; // Ventilations specifiques Salaries / Remunerations
      if (Tob_PreVentil <> nil) and (VentilTrouv <> '') then // pas de ventilation trouvee alors on traite la ligne histobulletin suivante
      begin
        // boucle sur la TOB des preventilations pour trouver toutes les ventilations sur un axe et pour parcourir tous les Axes
        while (Tob_PreVentil <> nil) do
        begin
          if (Nature <> 'COT') or (VentilTrouv = 'C') then
          begin // si remuneration ou ventil specifique cotisation
            i := i + 1; // incrementation automatique du numéro de ligne
            TOBA := TOB.Create('VENTANA', TOBAnalytique, -1);
            TOBA.PutValue('YVA_TABLEANA', 'PPU');
            RefA := EncodeRefPaie(Salarie, DateDebut, DateFin, LaTOB_Rub);
            TOBA.PutValue('YVA_IDENTIFIANT', RefA);
            TOBA.PutValue('YVA_NATUREID', 'PG');
            TOBA.PutValue('YVA_IDENTLIGNE', FormatFloat('000', i));
            if (VentilTrouv <> 'I') and (VentilTrouv <> 'VS') then // PT53
            begin
              TOBA.PutValue('YVA_SECTION', Tob_PreVentil.GetValue('V_SECTION'));
              TOBA.PutValue('YVA_POURCENTAGE', Tob_PreVentil.GetValue('V_TAUXMONTANT'));
              TOBA.PutValue('YVA_NUMVENTIL', Tob_PreVentil.GetValue('V_NUMEROVENTIL'));
              TOBA.PutValue('YVA_AXE', 'A' + Copy(Tob_PreVentil.GetValue('V_NATURE'), 3, 1));
            end
            else
            begin
              TOBA.PutValue('YVA_SECTION', Tob_PreVentil.GetValue('PAV_SECTION'));
              TOBA.PutValue('YVA_POURCENTAGE', Tob_PreVentil.GetValue('PAV_TAUXMONTANT'));
              TOBA.PutValue('YVA_NUMVENTIL', Tob_PreVentil.GetValue('PAV_NUMEROVENTIL'));
              TOBA.PutValue('YVA_AXE', 'A' + Copy(Tob_PreVentil.GetValue('PAV_NATURE'), 3, 1));
            end;
          end; // fin si rem ou cot non deja traitee
          // PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
          if (VentilTrouv = 'I') and (Assigned(TOBA)) then // PT46 Rajout test si TOB allouée
          begin
            ii := ValeurI(Copy(TOBA.GetValue('YVA_AXE'), 2, 1));
            AxesImp[ii] := 'X'; // Recuperation des axes analytiques gérés au niveau salarié mais dans le fichier import
          end;
          // FIN PT43
          if VentilTrouv = 'S' then Tob_PreVentil := AnaTOB_VenSal.findNext([''], [''], TRUE)
          else if VentilTrouv = 'VS' then Tob_PreVentil := AnaTOB_VenSalMois.findNext([''], [''], TRUE) // PT53
          else if VentilTrouv = 'P' then Tob_PreVentil := AnaTOB_VenSalRub.findNext(['V_COMPTE'], [Salarie + ';' + Rubrique], TRUE)
          else if VentilTrouv = 'C' then Tob_PreVentil := TOB_VenCot.findNext(['V_COMPTE'], [Rubrique], TRUE)
          else if VentilTrouv = 'R' then Tob_PreVentil := TOB_VenRem.findNext(['V_COMPTE'], [Rubrique], TRUE)
            // PT38  : 11/08/2003 PH V_421 Boucle sur la tob de la table PAIEVENTIL
          else if VentilTrouv = 'I' then Tob_PreVentil := AnaTOB_VSalRubImp.findNext(['PAV_COMPTE'], [Salarie + ';' + Rubrique], TRUE);

        end; // Fin de la boucle
      end; // Fin ventilation trouvee
      // PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
      for ii := 1 to 5 do
      begin // Cas où les ventilations ne sont pas trouvées dans le fichier d'import mais elles existent au niveau salarié
        if (AxesImp[ii] = '-') and (AxesSal[ii] = 'X') and (VentilTrouv = 'I') then
          CompleteVentileLignePaie(LaTOB_Rub, TOBAnalytique, AnaTOB_VenSal, Salarie, DateDebut, DateFin, ii);
      end;
      // FIN PT43
    end; // Fin si la rubrique est une cotisation ou une remuneration
    if TypeVentil = 'PRE' then LaTOB_Rub := TOB_Rub.FindNext(['PHB_SALARIE'], [Salarie], TRUE) // ligne bulletin suivante
    else LaTOB_Rub := nil;
  end;
  //PT25  : 17/09/2002 V585 PH Optimisation Ventilation analytique salarié non rechargées systématiquement à chaque rubrique
  //if AnaTOB_VenSal <> NIL then begin AnaTOB_VenSal.Free; AnaTOB_VenSal := NIL; end;
  //if AnaTOB_VenSalRub <> NIL then begin AnaTOB_VenSalRub.Free; AnaTOB_VenSalRub := NIL; end;
  //if AnaTOB_VSalRubImp <> NIL then begin AnaTOB_VSalRubImp.Free; AnaTOB_VSalRubImp := NIL; end;
  if AvecRaz then VideVentilAnaSalarie;
  // FIN PT10
  result := TOBAnalytique;
end;
// PT43  Fonction pour complèter les préventilations analytiques alors qu'elles ne sont pas renseignées dans le fichier import
{ Fonction de complément des ventilations analytiques dans le cas du traitement du fichier d'import
Uniquement dans le cas où on ne définit que les ventilations sur un seul axe alors que pour le salarié
plusieurs axes ont été renseignés. On ne traite que les ventilations analytiques salariés.
}

procedure CompleteVentileLignePaie(LaTOB_Rub, TOBAnalytique, AnaTOB_VenSal: TOB; Salarie: string; DateDebut, DateFin: TDateTime; NumAxe: Integer);
var
  RefA, Quoi: string;
  TOBA, Tob_PreVentil: TOB;
  i: integer;
begin
  Quoi := 'SA' + IntToStr(NumAxe);
  i := 0;
  Tob_PreVentil := AnaTOB_VenSal.findFirst(['V_NATURE'], [Quoi], TRUE);
  // boucle sur la TOB des preventilations pour trouver toutes les ventilations sur l'axe
  while (Tob_PreVentil <> nil) do
  begin
    i := i + 1; // incrementation automatique du numéro de ligne
    TOBA := TOB.Create('VENTANA', TOBAnalytique, -1);
    TOBA.PutValue('YVA_TABLEANA', 'PPU');
    RefA := EncodeRefPaie(Salarie, DateDebut, DateFin, LaTOB_Rub);
    TOBA.PutValue('YVA_IDENTIFIANT', RefA);
    TOBA.PutValue('YVA_NATUREID', 'PG');
    TOBA.PutValue('YVA_IDENTLIGNE', FormatFloat('000', i));
    TOBA.PutValue('YVA_SECTION', Tob_PreVentil.GetValue('V_SECTION'));
    TOBA.PutValue('YVA_POURCENTAGE', Tob_PreVentil.GetValue('V_TAUXMONTANT'));
    TOBA.PutValue('YVA_NUMVENTIL', Tob_PreVentil.GetValue('V_NUMEROVENTIL'));
    TOBA.PutValue('YVA_AXE', 'A' + Copy(Tob_PreVentil.GetValue('V_NATURE'), 3, 1));
    Tob_PreVentil := AnaTOB_VenSal.findNext(['V_NATURE'], [Quoi], TRUE)
  end;
end; // FIN PT43

{ Fonction de controle des doublons de la la tob des lignes analytiques
  Si doublon alors on elimine automatiquement la 2eme ligne.
}
// PT41  : 04/11/2003 PH V_500 Fonction de dédoublonnage des ventilations analytiques

procedure PgControlAna(TobAnalytiq: Tob);
var
  IdentAna, LigneAna, AxeAna, NumVAna: string;
  T1: TOB;
begin
  if not Assigned(TobAnalytiq) then exit;
  IdentAna := '';
  LigneAna := '';
  AxeAna := '';
  NumVAna := '';
  TobAnalytiq.Detail.Sort('YVA_IDENTIFIANT;YVA_IDENTLIGNE;YVA_AXE;YVA_NUMVENTIL');
  T1 := TobAnalytiq.FindFirst([''], [''], FALSE);
  while T1 <> nil do
  begin
    if (T1.getValue('YVA_IDENTIFIANT') = IdentAna) and (T1.getValue('YVA_IDENTLIGNE') = LigneAna) and
      (T1.getValue('YVA_AXE') = AxeAna) and (T1.getValue('YVA_NUMVENTIL') = NumVAna) then
      T1.Free
    else
    begin
      IdentAna := T1.getValue('YVA_IDENTIFIANT');
      LigneAna := T1.getValue('YVA_IDENTLIGNE');
      AxeAna := T1.getValue('YVA_AXE');
      NumVAna := T1.getValue('YVA_NUMVENTIL');
    end;
    T1 := TobAnalytiq.FindNext([''], [''], FALSE);
  end;
end;
// FIN PT41
{ Fonction de reaffectation des ventana en fonction des decompositions analytiques des remunerations ==> pour les cotisations
Principe : les cotisations ont toutes les mêmes ventilations et dependent donc des ventilations
des remunerations.
On ne prend que les rémunérations cde type Salaires cad on elimine celles qui ont un ordre de presentation >=5
Donc dans VENTANA : Identifiant : Salarie;Datedebut;DateFin;COT pour toutes les cotisations  Cas 1
sinon                           : Salarie;Datedebut;DateFin;C;xxxx pour une cotisation specifique

Pour faire cette affectation par defaut, il faut par axe et pour chaque section definir les pourcentages
et remplacer le cas 1
Le cas 2 ne sera jamais effacée car il correspond à une saisie
}

procedure PGReaffSectionAnal(TOB_Rub, TOBAna: tob; LDDebut: TDateTime = 0; LDFin: TDateTime = 0);
var
  i, NumAxe: Integer;
  TA, TVAxes, TV, TR, TC, TOBA: TOB;
  Ident, Axe, Sect: string;
  Rub, Rubrique: string;
  Pourc, Mt: Double;
  Montant: array[1..5] of Double;
  LSal, Nat, RF: string;
  DatD, DatF: TDateTime;
  Q: TQuery;
begin
  for i := 1 to 5 do Montant[i] := 0;
  if TOBAna = nil then exit;
  //PGVisuUnObjet (ToBAna,'','') ;

  TA := TOBAna.FindFirst([''], [''], FALSE);
  while TA <> nil do
  begin
    Ident := TA.GetVAlue('YVA_IDENTIFIANT');
    // PT27  : 10/10/2002 V585 PH recherche de tous les champs composant la clé des ventilations comptables
    RendRefPaieAnal(LSal, Rub, Nat, DatD, DatF, Ident);
    if (Nat[1] <> 'C') and (Nat[1] <> 'A') then TA.free;
    if Nat[1] = 'C' then
    begin
      if (Nat = 'COT') then TA.Free; // On elimine les TOB ventana pour ttes les cotisations
    end
    else
    begin
      // PT12 21/01/02 V571 PH Réaffectation des dates début et fin dans YVA_IDENTIFIANT en création de bulletin
      if (LDDebut <> 0) and (LDFin <> 0) then
      begin // Cas création de bulletin alors on renseigne Dates début et fin avec les dates de paies
        RF := TA.GetValue('YVA_IDENTIFIANT');
        RF := RendIdentAna(RF, LDDebut, LDFin); // on met les bonnes dates de la paie en creation de bulletin
        TA.PutValue('YVA_IDENTIFIANT', RF);
      end;
    end;
    TA := TOBAna.FindNext([''], [''], FALSE);
  end;
  //PGVisuUnObjet (ToBAna,'','') ;

  TVAxes := TOB.Create('Les nouvelles Ventana', nil, -1);
  TR := TOB_Rub.FindFirst(['PHB_NATURERUB'], ['AAA'], TRUE); // ligne bulletin remunerations
  while TR <> nil do
  begin
    Rubrique := TR.GetValue('PHB_RUBRIQUE');
    //    Mt := TR.GetValue('PHB_MTREM');  mt := TR.GetValue('PHB_ORDREETAT');
    if (TR.GetValue('PHB_ORDREETAT') < 7) and (TR.GetValue('PHB_MTREM') <> 0) then
    begin
      Rubrique := TR.GetValue('PHB_RUBRIQUE');
      TA := TOBAna.FindFirst([''], [''], FALSE);
      while TA <> nil do
      begin
        Ident := TA.GetVAlue('YVA_IDENTIFIANT');
        // PT27  : 10/10/2002 V585 PH recherche de tous les champs composant la clé des ventilations comptables
        RendRefPaieAnal(LSal, Rub, Nat, DatD, DatF, Ident);
        if (Nat[1] <> 'C') and (Rubrique = Rub) then // on ne traite que les remunerations
        begin // On a trouve les ventilations pour la rubrique concernée
          Axe := TA.GetValue('YVA_AXE');
          NumAxe := StrToInt(Copy(Axe, 2, 1));
          Sect := TA.GetValue('YVA_SECTION');
          TV := TVAxes.FindFirst(['YVA_AXE', 'YVA_SECTION'], [Axe, Sect], FALSE);
          if TV = nil then
          begin
            TV := TOB.Create('VENTANA', TVAxes, -1);
            TV.PutValue('YVA_TABLEANA', TA.GetValue('YVA_TABLEANA'));
            ident := EncodeRefPaieAlpha(LSal, 'COT', '', DatD, DatF);
            TV.PutValue('YVA_IDENTIFIANT', Ident);
            TV.PutValue('YVA_AXE', Axe);
            TV.PutValue('YVA_NATUREID', TA.GetValue('YVA_NATUREID'));
            TV.PutValue('YVA_IDENTLIGNE', TA.GetValue('YVA_IDENTLIGNE'));
            TV.PutValue('YVA_SECTION', Sect);
            TV.PutValue('YVA_NUMVENTIL', TA.GetValue('YVA_NUMVENTIL'));
            TV.PutValue('YVA_POURCENTAGE', 0);
            TV.PutValue('YVA_MONTANT', 0);
            TV.PutValue('YVA_MONTANTDEV', 0);
            //                  TV.PutValue ('YVA_MONTANTCON', 0);
          end;
          Mt := Arrondi(TR.GetValue('PHB_MTREM') * TA.GetValue('YVA_POURCENTAGE') / 100, 6);
          TV.PutValue('YVA_MONTANT', TV.GetValue('YVA_MONTANT') + Mt);
          Montant[NumAxe] := Montant[NumAxe] + Mt;
        end;
        TA := TOBAna.FindNext([''], [''], FALSE);
      end;
    end;
    TR := TOB_Rub.FindNext(['PHB_NATURERUB'], ['AAA'], TRUE); // ligne bulletin remunerations suivantes
  end;
  if TVAxes <> nil then
  begin
    TV := TVAxes.FindFirst([''], [''], FALSE);
    while TV <> nil do
    begin // Boucle de reaffectation  des pourcentages par axe/section pour les cotisations
      Axe := TV.GetValue('YVA_AXE');
      NumAxe := StrToInt(Copy(Axe, 2, 1));
      if Montant[NumAxe] <> 0 then Pourc := Arrondi((TV.GetValue('YVA_MONTANT') / Montant[NumAxe]) * 100, 6)
      else Pourc := 0; // PT50
      TV.PutValue('YVA_MONTANT', 0);
      TV.PutValue('YVA_POURCENTAGE', Pourc);
      TV := TVAxes.FindNext([''], [''], FALSE);
    end;
    // DEB PT64
    if TVAxes.detail.Count = 0 then
    begin // Recup des ventilations type salariés car les rémunérations ne sont pas valorisées
      i := 0;
      Q := OpenSQL('SELECT * FROM VENTIL WHERE V_NATURE lIKE "SA%" AND V_COMPTE="' + LSal + '"', FALSE);
      while not Q.EOF do
      begin
        TOBA := TOB.Create('VENTANA', TVAxes, -1);
        TOBA.PutValue('YVA_TABLEANA', 'PPU');
        ident := EncodeRefPaieAlpha(LSal, 'COT', '', DatD, DatF);
        TOBA.PutValue('YVA_IDENTIFIANT', Ident);
        TOBA.PutValue('YVA_NATUREID', 'PG');
        i := i + 1;
        TOBA.PutValue('YVA_IDENTLIGNE', FormatFloat('000', i));
        TOBA.PutValue('YVA_SECTION', Q.FindField('V_SECTION').AsString);
        TOBA.PutValue('YVA_POURCENTAGE', Q.FindField('V_TAUXMONTANT').AsFloat);
        TOBA.PutValue('YVA_NUMVENTIL', Q.FindField('V_NUMEROVENTIL').AsInteger);
        TOBA.PutValue('YVA_AXE', 'A' + Copy(Q.FindField('V_NATURE').AsString, 3, 1));
        Q.Next;
      end;
    end;
    // FIN PT64
    TV := TVAxes.FindFirst([''], [''], FALSE);
    while TV <> nil do
    begin // Boucle de reaffectation  des pourcentages par axe/section pour les cotisations
      TC := TOBAna.FindFirst(['YVA_TABLEANA', 'YVA_IDENTIFIANT', 'YVA_NATUREID', 'YVA_IDENTLIGNE', 'YVA_AXE', 'YVA_NUMVENTIL'],
        [TV.GetValue('YVA_TABLEANA'), TV.GetValue('YVA_IDENTIFIANT'), TV.GetValue('YVA_NATUREID'), TV.GetValue('YVA_IDENTLIGNE'), TV.GetValue('YVA_AXE'),
        TV.GetValue('YVA_NUMVENTIL')], FALSE);
      while TC <> nil do
      begin // Boucle de dedoublonnage des numeros de ventilation on regarde si la clé n'existe pas dans TOBana avant de l'intercaler
        TV.PutValue('YVA_NUMVENTIL', TV.GetValue('YVA_NUMVENTIL') + 1);
        TC := TOBAna.FindFirst(['YVA_TABLEANA', 'YVA_IDENTIFIANT', 'YVA_NATUREID', 'YVA_IDENTLIGNE', 'YVA_AXE', 'YVA_NUMVENTIL'],
          [TV.GetValue('YVA_TABLEANA'), TV.GetValue('YVA_IDENTIFIANT'), TV.GetValue('YVA_NATUREID'), TV.GetValue('YVA_IDENTLIGNE'), TV.GetValue('YVA_AXE'),
          TV.GetValue('YVA_NUMVENTIL')], FALSE);
      end;
      TV.ChangeParent(TOBAna, -1);
      TV := TVAxes.FindNext([''], [''], FALSE);
    end;
  end;
  //   PGVisuUnObjet (ToBAna,'','') ;
  if TVAxes <> nil then TVAxes.Free;
  PgControlAna(TobAna);
end;

// Fonction de creation et de controle existance

function PgCreationSection(TSection: TOB; Sect: array of string; CreationSection: Boolean): Boolean;
var
  i: Integer;
  TSect, TS, T1, THH: TOB;
  St: string;
begin
  Result := False;
  if not CreationSection then exit;
  TS := TOB.Create('Les SECTIONS Paie', nil, -1); // Creation de la tob des sections
  for i := 0 to 4 do
  begin // boucle sur les 5 Axes pour voir si les sections existent
    St := Sect[i]; // Attention la clé est composée de la section et de l'axe
    if St = '' then continue;
    TSect := TSection.FindFirst(['S_SECTION', 'S_AXE'], [St, 'A' + IntToStr(i + 1)], FALSE);
    if (TSect = nil) then
    begin
      T1 := TOB.Create('SECTION', TS, -1);
      T1.PutValue('S_SECTION', St);
      T1.PutValue('S_LIBELLE', 'Section paie ' + St);
      T1.PutValue('S_ABREGE', 'Section ' + St);
      T1.PutValue('S_SENS', 'M');
      T1.PutValue('S_SOLDEPROGRESSIF', 'X');
      T1.PutValue('S_AXE', 'A' + IntToStr(i + 1));
      THH := TOB.create('SECTION', TSection, -1); // on insere aussi dans la tob des sections la section créée sinon on va l'ecrire x fois
      THH.Dupliquer(T1, FALSE, TRUE, TRUE); // donc on duplique
    end;
  end;
  if (TS <> nil) and (TS.Detail.Count > 0) then
  begin // insertion des sections créees
    TS.InsertDB(nil, FALSE);
    TS.Free;
    TS := nil;
  end;
  if TS <> nil then TS.free;
end;
{ Fonction d'affectation des ventilations analytiques salariés par defaut en fonction
de la gestion automatique analytique des paramsoc de la paie.
On considère que la table des axes est définie et les axes sont connus
Il se pose le problème de la création automatique des sections.
On remplit la table VENTIL des ventilations analytiques salariés.
Attention, si on modifie les affectations salariés (Etablissement,4 Codes Stats,
4 tablettes libres de type combo, il faut alors réaffecter les ventilations par défaut.
TAxes Tob des axes analytiques
TSection Tob des sections analytiques
}

function AffectVentilSalarie(TSal, TAxes, TSection: TOB; CreationSection: Boolean): Boolean;
var
  TVentil, TV, Ta: TOB;
  St: string;
  i: Integer;
  Sect: array[1..5] of string; // Tableau des 5 sections analytiques possibles
begin
  result := FALSE;
  if TSal = nil then exit;
  if not VH_Paie.PGSectionAnal then exit;
  RendSectionSalarie(TSal, Sect); // remplit le tableau des 5 sections analytiques
  // Boucle pour la creation et la vérification de l'existance des sections
  PgCreationSection(TSection, Sect, CreationSection);
  TVentil := TOB.Create('Les Ventilsations du Salarie', nil, -1);
  for i := 1 to 5 do
  begin
    Ta := TAxes.FindFirst(['X_AXE'], ['A' + IntToStr(i)], FALSE);
    if (Ta <> nil) and (Ta.GetValue('X_FERME') <> 'X') and (Sect[i] <> '') then // Axe non fermé
    begin
      TV := TOB.Create('VENTIL', TVentil, -1);
      TV.PutValue('V_NATURE', 'SA' + IntToStr(i));
      TV.PutValue('V_COMPTE', TSal.GetValue('PSA_SALARIE'));
      TV.PutValue('V_SECTION', Sect[i]);
      TV.PutValue('V_TAUXMONTANT', 100);
      TV.PutValue('V_TAUXQTE1', 0);
      TV.PutValue('V_TAUXQTE2', 0);
      TV.PutValue('V_NUMEROVENTIL', 1);
      TV.PutValue('V_SOCIETE', '');
      TV.PutValue('V_MONTANT', 0);
      TV.PutValue('V_SOUSPLAN1', '');
      TV.PutValue('V_SOUSPLAN2', '');
      TV.PutValue('V_SOUSPLAN3', '');
      TV.PutValue('V_SOUSPLAN4', '');
      TV.PutValue('V_SOUSPLAN5', '');
      TV.PutValue('V_SOUSPLAN6', '');
    end;
  end;
  try
    BeginTrans;
    St := 'DELETE FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + TSal.GetVAlue('PSA_SALARIE') + '"';
    ExecuteSQL(st);
    TVentil.SetAllModifie(TRUE);
    TVentil.InsertDB(nil, FALSE);
    CommitTrans;
    result := TRUE;
  except
    Rollback;
    PGIError('Erreur création des ventilations analytiques du Salarié ' + TSal.GetValue('PSA_SALARIE') + ' ' + TSal.GetVAlue('PSA_LIBELLE') + ' ' + TSal.GetVAlue('PSA_PRENOM'),
      'Ventilations analytiques salariés');
  end;
  Ta.Free;
  TVentil.Free;
end;
// Fonction qui donne les sections analytiques par defaut en fonction des paramètres societe

procedure RendSectionSalarie(TSal: TOB; var Sect: array of string);
var
  i, j, k: Integer;
  LaSection, Letyp: string;
begin
  for i := 0 to 4 do Sect[i] := '######';

  for i := 0 to 4 do
  begin
    LaSection := '';
    for j := 1 to 3 do
    begin
      LeTyp := GetParamSoc('SO_PGAX' + IntToStr(i + 1) + 'ANAL' + IntToStr(j));
      if LeTyp <> '' then
      begin // cas etablissement
        if LeTyp = 'ETB' then LaSection := LaSection + TSal.GetValue('PSA_ETABLISSEMENT')
        else
        begin
          if pos('O', LeTyp) > 0 then // Cas code organisations OR1
          begin
            k := StrToInt(Copy(LeTyp, 3, 1));
            if (k > 0) and (k < 5) then LaSection := LaSection + TSal.GetValue('PSA_TRAVAILN' + IntToStr(k));
          end;
          if pos('C', LeTyp) > 0 then // Cas code libres TC1
          begin
            k := StrToInt(Copy(LeTyp, 3, 1));
            if (k > 0) and (k < 5) then LaSection := LaSection + TSal.GetValue('PSA_LIBREPCMB' + IntToStr(k));
          end;
          if LeTyp = 'STA' then LaSection := LaSection + TSal.GetValue('PSA_CODESTAT'); // cas code stat
        end;
      end;
    end;
    if LaSection <> '' then PGCompleteSection(LaSection, 'A' + IntToStr(i + 1));
    Sect[i] := LaSection;
  end;
end;
// Bourrage à droite du code section par rapport à la longueur des codes sections par axe

procedure PGCompleteSection(var Sect: string; Axe: string);
{$IFDEF COMPTAPAIE}
var
  Lg, ll, i: Integer;
  Bourre: string;
{$ENDIF}
begin
{$IFDEF COMPTAPAIE}
  Bourre := VH^.Cpta[AxeToFb(Axe)].Cb; // Recup caractère de bourrage des comptes généraux
  ll := Length(Sect);
  Lg := VH^.Cpta[AxeToFb(Axe)].Lg;
  if ll < Lg then
  begin
    for i := 1 to (Lg - ll) do Sect := Sect + Bourre;
  end
  else
    if ll > Lg then Sect := Copy(Sect, 1, Lg);
  // Si on a une longueur de section > à la longueur maxi dans la compta alors on limite BOF mais mieux que rien car evite ACCESS VIO
{$ENDIF}
end;
{ Fonction de prechargement = chargement de la Tob Analytique en consultation,
On recupère toutes les lignes du bulletin
Il n'y a pas de reaffectation
}

function PreChargeVentileLignePaie(Salarie: string; DateDebut, DateFin: TDateTime): TOB;
var
  RefA, St: string;
  TOBAnal: TOB;
begin
  TOBAnal := TOB.Create('Analytique bulletin Salarié', nil, -1);
  RefA := Salarie + ';' // Code Salarie
    + FormatDateTime('ddmmyyyy', DateDebut) + ';' // Date Debut bulletin
    + FormatDateTime('ddmmyyyy', DateFin) + ';'; // Date Fin bulletin
  St := 'SELECT * FROM VENTANA WHERE YVA_TABLEANA = "PPU" AND YVA_IDENTIFIANT like "' + RefA + '%" ORDER BY YVA_AXE,YVA_NATUREID,YVA_IDENTLIGNE';
{Flux optimisé
  Q := OpenSql(st, TRUE);
  TOBAnal.LoadDetailDB('VENTANA', '', '', Q, FALSE);
  ferme(Q);
}
  TOBAnal.LoadDetailDBFROMSQL('VENTANA', st);
  result := TOBAnal;
end;

{ Fonction de controle que chaque ligne du bulletin a une ventilation analytique
ceci pour donner les ventilations analytiques de rubriques rajouter dans un profil
en modification du bulletin ou en re preparation auto des bulletins
}

procedure ControlAffecteAnal(Salarie: string; DateDebut, DateFin: TDateTime; TOB_VenRem, TOB_VenCot, TOB_Rubrique, TOB_Anal: TOB);
var
  LaTOB_Rub, LaTobAnal, TPR: TOB;
  RefA, NatureRub: string;
begin
  LaTOB_Rub := TOB_Rubrique.FindFirst(['PHB_SALARIE'], [Salarie], TRUE);
  while LaTOB_Rub <> nil do
  begin
    NatureRub := LaTOB_Rub.GetValue('PHB_NATURERUB');
    if (NatureRub = 'BAS') or (Copy(LaTOB_Rub.GetValue('PHB_RUBRIQUE'), 5, 1) = '.') then
    begin // On ne traite pas les bases de cotisations et les lignes de commentaire
      LaTOB_Rub := TOB_Rubrique.FindNext(['PHB_SALARIE'], [Salarie], TRUE);
      continue;
    end;
    RefA := EncodeRefPaie(Salarie, DateDebut, DateFin, LaTOB_Rub);
    LaTobAnal := TOB_Anal.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
    if LaTobAnal = nil then
    begin
      LaTobAnal := PreVentileLignePaie(TOB_VenRem, TOB_VenCot, LaTOB_Rub, Salarie, 'INS', DateDebut, DateFin, FALSE);
      if LaTobAnal <> nil then
      begin
        TPR := LaTobAnal.findfirst([''], [''], TRUE);
        while TPR <> nil do
        begin
          TPR.ChangeParent(TOB_Anal, -1);
          TPR := LaTobAnal.findnext([''], [''], TRUE);
        end;
        if LaTobAnal <> nil then LaTobAnal.Free;
        TOB_Anal.Detail.Sort('YVA_TABLEANA;YVA_IDENTIFIANT');
      end;
    end;
    LaTOB_Rub := TOB_Rubrique.FindNext(['PHB_SALARIE'], [Salarie], TRUE);
  end;
  //PT25  : 17/09/2002 V585 PH Optimisation Ventilation analytique salarié non rechargées systématiquement à chaque rubrique
  VideVentilAnaSalarie;
end;

function ConvertiFichierVirement(NomFichier, Support: string): string;
var
  IntPos: integer;
begin
  NomFichier := Trim(NomFichier);
  result := NomFichier;
  if Support = 'DIS' then
  begin
    IntPos := Pos(':\', NomFichier);
    IntPos := IntPos + Pos('\\', NomFichier); {PT5}
    if IntPos > 0 then
      result := 'A:\' + ExtractFileName(NomFichier) {PT5}
    else
      Result := 'A:\' + NomFichier;
    PGIBox('Penser à insérer une disquette dans le lecteur!', 'Virement');
  end;
  if Support = 'TEL' then
  begin
    IntPos := Pos(':\', NomFichier);
    IntPos := IntPos + Pos('\\', NomFichier);
    if IntPos > 0 then
    begin
      Result := NomFichier;
    end // NomFichier[1]:='C'; {PT5}
    else Result := 'C:\' + NomFichier;
  end;
end;
{  PT6 16/08/01 V547 PH PGRempliDefaut devient une fonction générique de la paie
Cette fonction permet l'initialisation par défaut des tablettes choixcod remplies avec
des valeurs par défaut de la paie. C'est le cas des tablettes :
PGQUALIFBASE, PGTYPEORGANISME
}

procedure PGRempliDefaut;
var
  Q: TQuery;
  St: string;
begin
  St := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PTG" AND CC_CODE ="001"'; // PGTYPEORGANISME
  Q := OpenSql(St, TRUE);
  if Q.EOF then
  begin
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","001","URSSAF","URSSAF","URSSAF")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","002","ASSEDIC","ASSEDIC","ASSEDIC")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","003","AGIRC","AGIRC","AGIRC")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","004","ARRCO","ARRCO","ARRCO")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","005","PREVOYANCE","PREVOYANCE","PREVOYANCE")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","006","MSA","MSA","MSA")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","007","Taxe sur les salaires","Taxes salaires","Taxe sur les salaires")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","008","Autres taxes","Autres taxes","Autres taxes")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","009","CCVRP","CCVRP","CCVRP")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","010","IRPVRP","IRPVRP","IRPVRP")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PTG","011","IRREP","IRREP","IRREP")');
  end;
  Ferme(Q);

  // FIN PT6 16/08/01 V547 PH
  {St := 'SELECT PRC_SALARIE FROM REPOSSALARIE WHERE PRC_SALARIE="ZZZZZZZZZZZZZZZZZ"';  // PTXX @@@@
  Q := OpenSql (St, TRUE);
  if Q.EOF then
   begin
   try
   BeginTrans;
   Wt := '"'+UsDateTime(iDate1900)+'"';
   st := 'update absencesalarie set pcn_codetape="..." where pcn_typemvt="CPA" and pcn_typeconge="PRI" and (pcn_codetape is null or pcn_codetape="")';
   ExecuteSQL(st);
   st := 'update absencesalarie set pcn_validresp="VAL" where pcn_typemvt="CPA" and pcn_typeconge="PRI" and (pcn_validresp is null or pcn_validresp="")';
   ExecuteSQL(st);
   st := 'INSERT INTO REPOSSALARIE (PRC_SALARIE,PRC_DATEDEBUT,PRC_DATEFIN,PRC_ORDRE,PRC_TYPEMVT,PRC_NBREHEURE,PRC_DATEMVT,PRC_DATESAISIE) '+
             ' VALUES '+
             ' ("ZZZZZZZZZZZZZZZZZ",'+Wt+','+Wt+',"0","ZZZ","0",'+Wt+','+Wt+')';
   ExecuteSQL(st);
   CommitTrans;
   Except
   Rollback;
   end;
   end;
  Ferme (Q);}
  //DEBUT PT55
  St := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PGZ" AND CC_CODE>="SC1" AND CC_CODE<="SC5"'; // PGZONELIBRE
  Q := OpenSql(St, TRUE);
  if Q.EOF then
  begin
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PGZ","SC1","Tablette libre scoring 1","Tablette libre 1","")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PGZ","SC2","Tablette libre scoring 2","Tablette libre 2","")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PGZ","SC3","Tablette libre scoring 3","Tablette libre 3","")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PGZ","SC4","Tablette libre scoring 4","Tablette libre 4","")');
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PGZ","SC5","Tablette libre scoring 5","Tablette libre 5","")');
  end;
  Ferme(Q);
  //FIN PT55
  // DEB PT65
  St := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="PCL" AND CC_CODE ="000"'; // PGTYPEORGANISME
  Q := OpenSql(St, TRUE);
  if Q.EOF then
  begin
    ExecuteSQL('INSERT INTO CHOIXCOD (CC_TYPE,CC_CODE,CC_LIBELLE,CC_ABREGE,CC_LIBRE) ' +
      ' VALUES ' +
      ' ("PCL","000","Sans catégorie","Sans catégorie","Sans catégorie")');
  end;
  ferme(Q);
  // FIN PT65
end;

procedure MajConvention;
var
  Tob_Convention: Tob;
  AncValeur, NewValeur: string;
  i: integer;
begin
  if ExisteSQL('SELECT PCV_CONVENTION FROM CONVENTIONCOLL') = False then exit;
  Tob_Convention := Tob.Create('Les conventions', nil, -1);
  Tob_Convention.LoadDetailDB('CONVENTIONCOLL', '', '', nil, False);
  for i := 0 to Tob_Convention.Detail.Count - 1 do
  begin
    NewValeur := '';
    AncValeur := Tob_Convention.Detail[i].Getvalue('PCV_CONVENTION');
    if (isnumeric(AncValeur)) and (Length(AncValeur) < 3) then
      NewValeur := ColleZeroDevant(StrToInt(AncValeur), 3);
    if (NewValeur <> '') then
    begin
      if (ExisteSQL('SELECT PCV_CONVENTION FROM CONVENTIONCOLL ' + {PT7}
        'WHERE PCV_CONVENTION="' + NewValeur + '"') = True) then
      begin
        PGIBox('La convention collective : ' + AncValeur + ' correspond à la convention 001!#13#10' +
          'Vous devez en supprimer une!.', 'Duplication de convention collective');
        exit;
      end;
      try
        BeginTrans;
        ExecuteSql('UPDATE CONVENTIONCOLL SET PCV_CONVENTION="' + NewValeur + '" WHERE PCV_CONVENTION="' + AncValeur + '" ');
        ExecuteSql('UPDATE SALARIES  SET PSA_CONVENTION="' + NewValeur + '" WHERE PSA_CONVENTION="' + AncValeur + '" ');
        ExecuteSql('UPDATE ETABCOMPL SET ETB_CONVENTION="' + NewValeur + '" WHERE ETB_CONVENTION="' + AncValeur + '" ');
        ExecuteSql('UPDATE ETABCOMPL SET ETB_CONVENTION1="' + NewValeur + '" WHERE ETB_CONVENTION1="' + AncValeur + '" ');
        ExecuteSql('UPDATE ETABCOMPL SET ETB_CONVENTION2="' + NewValeur + '" WHERE ETB_CONVENTION2="' + AncValeur + '" ');
        ExecuteSql('UPDATE MINCONVPAIE SET PCP_CONVENTION="' + NewValeur + '" WHERE PCP_CONVENTION="' + AncValeur + '" ');
        ExecuteSql('UPDATE MINIMUMCONVENT SET PMI_CONVENTION="' + NewValeur + '" WHERE PMI_CONVENTION="' + AncValeur + '" ');
        ExecuteSql('UPDATE PAIEENCOURS SET PPU_CONVENTION="' + NewValeur + '" WHERE PPU_CONVENTION="' + AncValeur + '" ');
        CommitTrans;
      except
        Rollback;
      end;
    end;
  end; //End For
  Tob_Convention.free;
end;

procedure MajRubrique;
var
  St: string;
  Nbre: Integer;
  Q: TQuery;
begin
  exit;
  //if (V_PGI_Env.NoDossier='000000') then  exit;
  // **** Refonte accès V_PGI_env ***** V_PGI_env.ModeFonc remplacé par PgRendModeFonc () *****
  if (PgRendModeFonc() <> 'MULTI') then exit;

  try
    if v_pgi.NumBuild = '113' then
    begin
      BeginTrans;
      ExecuteSql('DELETE FROM COTISATION WHERE PCT_PREDEFINI="DOS" AND PCT_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM REMUNERATION WHERE PRM_PREDEFINI="DOS" AND PRM_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM CUMULRUBRIQUE WHERE PCR_PREDEFINI="DOS" AND PCR_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM CUMULRUBRIQUE WHERE PCR_PREDEFINI="CEG" AND PCR_NODOSSIER <> "000000" ');
      ExecuteSql('DELETE FROM ELTNATIONAUX WHERE PEL_PREDEFINI="DOS" AND PEL_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM PROFILPAIE WHERE PPI_PREDEFINI="DOS" AND PPI_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM MINIMUMCONVENT WHERE PMI_PREDEFINI="DOS" AND PMI_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM MINCONVPAIE WHERE PCP_PREDEFINI="DOS" AND PCP_NODOSSIER="000000" ');
      ExecuteSql('DELETE FROM MOTIFABSENCE WHERE PMA_PREDEFINI="DOS" AND PMA_NODOSSIER="000000" ');
      CommitTrans;
    end;
  except
    Rollback;
  end;

  try
    if v_pgi.NumBuild >= '525' then
    begin // Test de doublons
      BeginTrans;
      Nbre := 0;
      St := 'select count (*) from cumulrubrique where ((PCR_RUBRIQUE LIKE "%%%1") OR (PCR_RUBRIQUE LIKE "%%%3") ' +
        ') AND (PCR_PREDEFINI="DOS") AND (PCR_NODOSSIER <> "000000") AND ' +
        '(PCR_NATURERUB = "AAA")';

      Q := OpenSql(St, TRUE);
      if not Q.EOF then Nbre := Q.Fields[0].AsInteger;
      Ferme(Q);
      if Nbre > 0 then
      begin
        St := 'delete from cumulrubrique where ((PCR_RUBRIQUE LIKE "%%%1") OR (PCR_RUBRIQUE LIKE "%%%3") ' +
          ') AND (PCR_PREDEFINI="DOS") AND (PCR_NODOSSIER <> "000000") AND ' +
          ' (PCR_NATURERUB = "AAA") AND ' +
          ' EXISTS (SELECT PCR_RUBRIQUE FROM cumulrubrique AS CR where (CR.PCR_PREDEFINI="DOS") AND (CR.PCR_NODOSSIER <> "000000") AND (CR.PCR_NATURERUB = "AAA") AND (CR.PCR_NODOSSIER <> CUMULRUBRIQUE.PCR_NODOSSIER) ' +
          'AND (CR.PCR_RUBRIQUE = CUMULRUBRIQUE.PCR_RUBRIQUE) AND (CR.PCR_CUMULPAIE = CUMULRUBRIQUE.PCR_CUMULPAIE))';
        ExecuteSql(st);
        St := 'update cumulrubrique set PCR_NODOSSIER="000000",PCR_PREDEFINI="STD" where ((PCR_RUBRIQUE LIKE "%%%1") OR (PCR_RUBRIQUE LIKE "%%%3") ' +
          ' ) AND (PCR_PREDEFINI="DOS") AND (PCR_NODOSSIER <> "000000") AND ' +
          ' (PCR_NATURERUB = "AAA") ';
        ExecuteSql(st);
      end;
      CommitTrans;
    end
  except
    Rollback;
  end;

end;
// Fonction de controle des enregistrements des attestations ne s'execute qu'une fois

procedure MAJAttestations;
var
  StDelete, StExiste, StNum, StSalarie: string;
  QRechAtt: TQuery;
begin
  //Cas où enregistrement dans ATTSALAIRES et non dans ATTESTATIONS
  StExiste := 'SELECT PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST' +
    ' FROM ATTSALAIRES' +
    ' WHERE PAL_ORDRE not in (SELECT PAS_ORDRE FROM ATTESTATIONS)';
  if ExisteSQL(StExiste) = True then
  begin
    StDelete := 'DELETE FROM ATTSALAIRES WHERE PAL_ORDRE not in (SELECT PAS_ORDRE FROM ATTESTATIONS)';
    ExecuteSQL(StDelete);
  end;
  //Cas où ORDRE IDEM mais SALARIES different
  StExiste := 'SELECT  PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST, PAS_SALARIE,' +
    ' PAS_TYPEATTEST' +
    ' FROM ATTSALAIRES' +
    ' LEFT JOIN ATTESTATIONS ON PAS_ORDRE=PAL_ORDRE' +
    ' WHERE PAS_SALARIE <> PAL_SALARIE';

  if ExisteSQL(StExiste) = True then
  begin
    StDelete := 'DELETE FROM ATTSALAIRES WHERE PAL_ORDRE in (SELECT PAS_ORDRE FROM ATTESTATIONS WHERE PAS_SALARIE <> PAL_SALARIE)';
    ExecuteSQL(StDelete);
    StDelete := 'DELETE FROM ATTESTATIONS WHERE PAS_ORDRE not in (SELECT PAL_ORDRE FROM ATTSALAIRES)';
    ExecuteSQL(StDelete);
  end;

  //CAS Où ORDRE IDEM mais TYPEATTEST different
  StExiste := 'SELECT PAL_ORDRE, PAL_SALARIE, PAL_TYPEATTEST, PAS_SALARIE,' +
    ' PAS_TYPEATTEST' +
    ' FROM ATTSALAIRES' +
    ' LEFT JOIN ATTESTATIONS ON PAS_ORDRE=PAL_ORDRE' +
    ' WHERE PAS_TYPEATTEST <> PAL_TYPEATTEST ';

  if ExisteSQL(StExiste) = True then
  begin
    StDelete := 'DELETE FROM ATTSALAIRES WHERE PAL_ORDRE in (SELECT PAS_ORDRE FROM ATTESTATIONS WHERE PAS_TYPEATTEST <> PAL_TYPEATTEST)';
    ExecuteSQL(StDelete);
    StDelete := 'DELETE FROM ATTESTATIONS WHERE PAS_ORDRE not in (SELECT PAL_ORDRE FROM ATTSALAIRES)';
    ExecuteSQL(StDelete);
  end;

  //CAS Où Attestations ASSEDIC absente, MAJ fiche salarié
  //PT52
  StSalarie := 'SELECT PSA_SALARIE' +
    ' FROM SALARIES WHERE' +
    ' PSA_ASSEDIC="X" AND' +
    ' PSA_SALARIE NOT IN' +
    '(SELECT PAS_SALARIE' +
    ' FROM ATTESTATIONS WHERE' +
    ' PAS_TYPEATTEST="ASS")';

  QRechAtt := OpenSql(StSalarie, True);
  while not QRechAtt.Eof do
  begin
    StNum := QRechAtt.FindField('PSA_SALARIE').Asstring;
    StDelete := 'UPDATE SALARIES SET' +
      ' PSA_ASSEDIC="-" WHERE' +
      ' PSA_SALARIE="' + StNum + '"';
    ExecuteSQL(StDelete);
    QRechAtt.Next;
  end;
  Ferme(QRechAtt);
end;
// fonction de sauvegarde des ventilations analytiques et de l'affectation des historiques analytiques

function SauvegardeAnal(TOBAna, TOB_Rubrique, TOB_VentilRem, TOB_VentilCot: TOB; Comment: Boolean; LeSal: string; DateD, DateF: TDateTime): Boolean;
var
  NatureRub, Rubrique, RefA, MRefA, st, General: string; // PT70
  T1, T_HistoAna, T2, TAna: TOB;
  Mt: Double;
  Sens: string;
  ll: Integer; // PT70
begin
  result := TRUE;
  if TOBAna = nil then exit;
  // Suppression des lignes analytiques avant tout traitement
  try
    BeginTrans;
    st := 'DELETE FROM HISTOANALPAIE WHERE PHA_SALARIE="' + LeSal + '" AND PHA_DATEDEBUT="' + USDateTime(DateD) + '" AND PHA_DATEFIN="' + USDateTime(DateF) + '"';
    ExecuteSQL(St);
    RefA := LeSal + ';' // Code Salarie
      + FormatDateTime('ddmmyyyy', DateD) + ';' // Date Debut bulletin
      + FormatDateTime('ddmmyyyy', DateF) + ';'; // Date Fin bulletin
    st := 'DELETE FROM VENTANA WHERE YVA_TABLEANA="PPU" AND YVA_IDENTIFIANT like "' + RefA + '%"'; // suppression des ventilations analytiques du bulletin
    ExecuteSQL(St);
    TOBAna.SetAllModifie(TRUE);
    // PGVisuUnObjet (TOBANA, 'Les ecritures analytiques','');
    result := TOBAna.InsertDB(nil, FALSE);
    if Comment then MoveCur(FALSE);
    // Boucle de constitution des lignes historiques analytiques
    T1 := TOB_Rubrique.findfirst(['PHB_SALARIE'], [LeSal], FALSE);
    T_HistoAna := TOB.create('HISTO ANAL PAIE', nil, -1);
    while T1 <> nil do
    begin
      NatureRub := T1.GetValue('PHB_NATURERUB');
      Rubrique := T1.GetValue('PHB_RUBRIQUE');
      Mt := T1.GetValue('PHB_BASEREM') + T1.GetValue('PHB_MTREM') + T1.GetValue('PHB_BASECOT') +
        T1.GetValue('PHB_MTSALARIAL') + T1.GetValue('PHB_MTPATRONAL');
      // Récupération du numéro de compte, on se base sur les ventilations par défaut des rubriques idem génération comptables
      // Mais on ne va pas générer de lignes analytiques dans la paie si on n'a pas renseigner de compte
      // Cela correspond à une rubrique de calcul intermédiaire. Compte attente = '99999999'
      // PT3 28/06/01 V536 PH
      General := RendComptePreVentil(NatureRub, Rubrique, Sens, TOB_VentilRem, TOB_VentilCot, T1);
      if (NatureRub <> 'BAS') and (Mt <> 0) and (General <> '99999999') then // on ne traite que les lignes valorisées
      begin // on ne traite que les remunerations et les cotisations
        RefA := EncodeRefPaie(LeSal, DateD, DateF, T1);
        T2 := TOBAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
        // en premier s'il y a eu personnalisation des ventilations de la cotisation Sinon Cas commun des cotisations
        if (T2 = nil) and (NatureRub = 'COT') then
        begin
        // DEV PT70
          MRefA := RefA;
          RefA := Copy(RefA, 1, 29) + 'COT';
          T2 := TOBAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
          if (T2 = nil) and (VH_Paie.PgTypeNumSal = 'ALP') then
          begin
            ll := POS('C;', MRefA);
            if ll > 0 then RefA := Copy(MRefA, 1, ll - 1) + 'COT';
            T2 := TOBAna.FindFirst(['YVA_IDENTIFIANT'], [RefA], TRUE);
          end;
        // FIN PT70
        end;
        while T2 <> nil do
        begin
          TAna := TOB.create('HISTOANALPAIE', T_HistoAna, -1);
          TAna.PutValue('PHA_ETABLISSEMENT', T1.GetValue('PHB_ETABLISSEMENT'));
          TAna.PutValue('PHA_SALARIE', T1.GetValue('PHB_SALARIE'));
          TAna.PutValue('PHA_DATEDEBUT', T1.GetValue('PHB_DATEDEBUT'));
          TAna.PutValue('PHA_DATEFIN', T1.GetValue('PHB_DATEFIN'));
          TAna.PutValue('PHA_NATURERUB', T1.GetValue('PHB_NATURERUB'));
          TAna.PutValue('PHA_RUBRIQUE', T1.GetValue('PHB_RUBRIQUE'));
          TAna.PutValue('PHA_COTREGUL', T1.GetValue('PHB_COTREGUL'));
          TAna.PutValue('PHA_LIBELLE', T1.GetValue('PHB_LIBELLE'));
          TAna.PutValue('PHA_TRAVAILN1', T1.GetValue('PHB_TRAVAILN1'));
          TAna.PutValue('PHA_TRAVAILN2', T1.GetValue('PHB_TRAVAILN2'));
          TAna.PutValue('PHA_TRAVAILN3', T1.GetValue('PHB_TRAVAILN3'));
          TAna.PutValue('PHA_TRAVAILN4', T1.GetValue('PHB_TRAVAILN4'));
          TAna.PutValue('PHA_CODESTAT', T1.GetValue('PHB_CODESTAT'));
          TAna.PutValue('PHA_LIBREPCMB1', T1.GetValue('PHB_LIBREPCMB1'));
          TAna.PutValue('PHA_LIBREPCMB2', T1.GetValue('PHB_LIBREPCMB2'));
          TAna.PutValue('PHA_LIBREPCMB3', T1.GetValue('PHB_LIBREPCMB3'));
          TAna.PutValue('PHA_LIBREPCMB4', T1.GetValue('PHB_LIBREPCMB4'));
          TAna.PutValue('PHA_CONFIDENTIEL', T1.GetValue('PHB_CONFIDENTIEL'));
          TAna.PutValue('PHA_ORGANISME', T1.GetValue('PHB_ORGANISME'));
          TAna.PutValue('PHA_AXE', T2.GetValue('YVA_AXE'));
          TAna.PutValue('PHA_SECTION', T2.GetValue('YVA_SECTION'));
          TAna.PutValue('PHA_GENERAL', General);
          TAna.PutValue('PHA_ORDREETAT', T1.GetValue('PHB_ORDREETAT')); { PT59 }
          if NatureRub = 'AAA' then // Remunerations
          begin
            Mt := ARRONDI(T1.GetValue('PHB_BASEREM') * (T2.GetValue('YVA_POURCENTAGE') / 100), 5);
            TAna.PutValue('PHA_BASEREM', Mt);
            Mt := ARRONDI(T1.GetValue('PHB_MTREM') * (T2.GetValue('YVA_POURCENTAGE') / 100), 5);
            if Sens = 'M' then Mt := Mt * -1;
            TAna.PutValue('PHA_MTREM', Mt);
          end
          else
          begin // cotisations
            Mt := ARRONDI(T1.GetValue('PHB_BASECOT') * (T2.GetValue('YVA_POURCENTAGE') / 100), 5);
            TAna.PutValue('PHA_BASECOT', Mt);
            Mt := ARRONDI(T1.GetValue('PHB_MTSALARIAL') * (T2.GetValue('YVA_POURCENTAGE') / 100), 5);
            if Sens = 'M' then Mt := Mt * -1;
            TAna.PutValue('PHA_MTSALARIAL', Mt);
            Mt := ARRONDI(T1.GetValue('PHB_MTPATRONAL') * (T2.GetValue('YVA_POURCENTAGE') / 100), 5);
            if Sens = 'M' then Mt := Mt * -1;
            TAna.PutValue('PHA_MTPATRONAL', Mt);
          end;
          T2 := TOBAna.FindNext(['YVA_IDENTIFIANT'], [RefA], TRUE);
        end; // Fin boucle sur les sections analytiques multi Axes
      end;
      T1 := TOB_Rubrique.FindNext(['PHB_SALARIE'], [LeSal], FALSE);
    end;
    if T_HistoAna <> nil then
    begin
      T_HistoAna.SetAllModifie(TRUE);
      result := T_HistoAna.InsertDB(nil, FALSE);
      T_HistoAna.free;
    end;
    CommitTrans;
  except
    Rollback;
    PGIError('Une erreur est survenue lors de la validation des ventilations analytiques', 'Ventilations analytiques');
    result := FALSE;
  end;

end;
{ fonction qui rend le numero de compte gene pour histoanalpaie
Parcours des pré ventilations par defaut des cotisations et des rémunérations dans les
tables VENTICOTPAIE et VENTIREMPAIE
}

function RendComptePreVentil(var NatureRub, Rubrique, Sens: string; TOB_VentilRem, TOB_VentilCot, UneTOB: TOB): string;
var
  TOrg: TOB;
  Nodossier, TheRacine: string;
begin
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  Nodossier := PgRendNoDossier();

  result := '99999999'; // compte attente correspond à une rubrique pour laquelle on n'a pas trouvé de compte associé
  Sens := 'P';
  // Recherche du compte de charge affecté à la cotisation
  if Naturerub = 'COT' then
  begin
    Torg := Tob_VentilCot.FindFirst(['PVT_PREDEFINI', 'PVT_NODOSSIER', 'PVT_RUBRIQUE'], ['DOS', NoDossier, Rubrique], FALSE);
    if Torg = nil then
      Torg := Tob_VentilCot.FindFirst(['PVT_PREDEFINI', 'PVT_NODOSSIER', 'PVT_RUBRIQUE'], ['STD', '000000', Rubrique], FALSE);
    if Torg = nil then
      Torg := Tob_VentilCot.FindFirst(['PVT_PREDEFINI', 'PVT_NODOSSIER', 'PVT_RUBRIQUE'], ['CEG', '000000', Rubrique], FALSE);
    if Torg <> nil then
    begin
      Sens := Torg.GetValue('PVT_SENS1');
      TheRacine := Copy(Torg.GetValue('PVT_RACINE1'), 1, VH_Paie.PGLongRacine);
      if TheRacine = '' then
      begin
        TheRacine := Copy(Torg.GetValue('PVT_RACINE2'), 1, VH_Paie.PGLongRacine);
        Sens := Torg.GetValue('PVT_SENS2');
      end;
    end
  end
  else // cas d'une remuneration , on ne prend pas les bases de cotisations
  begin
    Torg := Tob_VentilRem.FindFirst(['PVS_PREDEFINI', 'PVS_NODOSSIER', 'PVS_RUBRIQUE'], ['DOS', NoDossier, Rubrique], FALSE);
    if Torg = nil then
      Torg := Tob_VentilRem.FindFirst(['PVS_PREDEFINI', 'PVS_NODOSSIER', 'PVS_RUBRIQUE'], ['STD', '000000', Rubrique], FALSE);
    if Torg = nil then
      Torg := Tob_VentilRem.FindFirst(['PVS_PREDEFINI', 'PVS_NODOSSIER', 'PVS_RUBRIQUE'], ['CEG', '000000', Rubrique], FALSE);
    if Torg <> nil then
    begin
      Sens := Torg.GetValue('PVS_SENS1');
      TheRacine := Copy(Torg.GetValue('PVS_RACINE1'), 1, VH_Paie.PGLongRacine);
      if TheRacine = '' then
      begin
        TheRacine := Copy(Torg.GetValue('PVS_RACINE2'), 1, VH_Paie.PGLongRacine);
        Sens := Torg.GetValue('PVS_SENS2');
      end;
    end;
  end;
{$IFDEF COMPTAPAIE}
  if TheRacine <> '' then result := ConfectionCpte(TheRacine, 'H', nil, FALSE, UneTOB);
{$ENDIF}
end;

procedure VisibiliteOrganisation(Ecran: TForm);
var
  TLieu, TA, TDE, TCodeStat: THLabel;
  ChampLieu, ChampLieu_: THEdit;
  num, nr: integer;
  numero, libelle, PrefTable: string;
begin
{$IFNDEF EAGLSERVER}
  if Ecran is TFFiche then PrefTable := PrefixeToTable(TFFiche(Ecran).TableName)
  else
    if Ecran is TFMul then
    begin
      if TFMul(Ecran).name = 'CONGESPAY_MUL' then PrefTable := 'PCN';
    end;
{$ENDIF}
  TCodeStat := THLabel(PGGetControl(('T' + PrefTable + '_CODESTAT'), Ecran));
  for num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PGNbreStatOrg then exit;
    Libelle := '';
    if Num = 1 then libelle := VH_Paie.PGLibelleOrgStat1;
    if Num = 2 then libelle := VH_Paie.PGLibelleOrgStat2;
    if Num = 3 then libelle := VH_Paie.PGLibelleOrgStat3;
    if Num = 4 then libelle := VH_Paie.PGLibelleOrgStat4;
    if (Libelle <> '') then
    begin
      ChampLieu := THEdit(PGGetControl((PrefTable + '_TRAVAILN' + Numero), Ecran));
      if ChampLieu <> nil then
      begin
        ChampLieu.Visible := TRUE;
      end;
      ChampLieu_ := THEdit(PGGetControl((PrefTable + '_TRAVAILN' + Numero + '_'), Ecran));
      if ChampLieu_ <> nil then
      begin
        ChampLieu_.Visible := TRUE;
      end;
      if (numero = '1') or (numero = '2') then
      begin
        TDE := THLabel(PGGetControl(('TDE' + Numero), Ecran));
        if TDE <> nil then TDE.visible := True;
        TA := THLabel(PGGetControl(('TA' + Numero), Ecran));
        if TA <> nil then TA.visible := True;
      end;
      TLieu := THLabel(PGGetControl(('T' + PrefTable + '_TRAVAILN' + Numero), Ecran));
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
    end;
  end;
  if VH_Paie.PGNbreStatOrg < 4 then
    for Nr := VH_Paie.PGNbreStatOrg to 4 do
    begin
      TLieu := THLabel(PGGetControl(('T' + PrefTable + '_TRAVAILN' + InttoStr(Nr + 1)), Ecran));
      if (TLieu <> nil) then TLieu.Caption := '';
    end;

  //Visibilité du code statistique
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    ChampLieu := THEdit(PGGetControl(PrefTable + '_CODESTAT', Ecran));
    ChampLieu_ := THEdit(PGGetControl(PrefTable + '_CODESTAT_', Ecran));
    if (TCodeStat <> nil) and (ChampLieu <> nil) and (ChampLieu_ <> nil) then
    begin
      TCodeStat.Visible := True;
      TCodeStat.caption := VH_Paie.PGLibCodeStat;
      ChampLieu.Visible := True;
      ChampLieu_.Visible := True;
    end;
  end;
end;

function RendExerciceCorrespondant(DateFin: TDateTime): string;
var
  Q: TQuery;
begin
  result := '';
  if (DateFin = idate1900) then exit;
  Q := OpenSQL('SELECT PEX_EXERCICE FROM EXERSOCIAL ' +
    'WHERE PEX_DATEDEBUT<="' + USDateTime(DateFin) + '" ' +
    'AND PEX_DATEFIN>="' + USDateTime(DateFin) + '" ', TRUE);
  if not Q.EOF then //PORTAGECWAS
    Result := Q.FindField('PEX_EXERCICE').AsString;
  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 26/07/2001
Modifié le ... : 26/07/2001
Description .. : Convertit une StringList de code salarie en string
Mots clefs ... : PAIE,SALARIE,CHAINE,FUSIONWORD
*****************************************************************}
// Pour modifier AND en WHERE
//StSal:='WHERE '+Copy(StSal,Pos('AND',StSal)+3,Length(StSal));
function RendStringListSal(Prefixe: string; ListSal: TStringList): string;
var
  i: integer;
  StSal: string;
begin
  StSal := '';
  for i := 0 to ListSal.Count - 1 do
  begin
    if i = 0 then StSal := 'AND (';
    StSal := StSal + Prefixe + '_SALARIE="' + ListSal[i] + '" ';
    if i <> ListSal.Count - 1 then StSal := StSal + 'OR ';
  end;
  if StSal <> '' then StSal := StSal + ')';
  Result := StSal;
end;


//DEB PT9

procedure MAJCotisation;
var
  Q: TQuery;
  St: string;
begin
  St := '';
  Q := OpenSql('SELECT PCT_RUBRIQUE FROM COTISATION WHERE PCT_ORDREETAT=0', True);
  while not Q.eof do
  begin
    St := St + ' PHB_RUBRIQUE="' + Q.FindField('PCT_RUBRIQUE').AsString + '" OR';
    Q.next;
  end;
  Ferme(Q);
  if St <> '' then
  begin
    St := Copy(St, 1, Length(St) - 3);
    St := 'WHERE (' + St + ') AND PHB_ORDREETAT=0';
    ExecuteSql('UPDATE COTISATION SET PCT_ORDREETAT=3 WHERE PCT_ORDREETAT=0');
    ExecuteSql('UPDATE HISTOBULLETIN SET PHB_ORDREETAT=3 ' + St);
  end;
end;
//FIN PT9


//DEB PT13  Moulinette des maj des bases des mvts d'acquis congés payés
//Déplacer dans UtofPGUtilitaireCP
//FIN PT13

//   PT18 30/04/2002 V582 PH Modifs des libellés de champs libres dans dechamp
{ La fonction analyse tous les champs libres et les champs dont les libellés sont
paramètrables et substitue les libellés définis dans dechamp pour les remplacer
par les libellés saisis par les utilisateurs. Ces libellés seront alors visibles
dans toutes les listes et les cubes et tobviewer.
Les champs non utilisés seront indiqués pour information
}

function PgRemplitCLauseEtab(ch1, ch2: string): string;
begin
  if ch1 = '' then ch1 := ch2
  else ch1 := ch1 + ' AND ' + ch2;
  result := ch1;
end;

procedure PgModifNomChamp;
var
  iTable, iChamp: integer;
  St, Prefixe, TheEtab, LesEtab: string;
  TOB_DesChamps, TobRH, TobCP: TOB; //PT91
  T1: TOB;
  zz, i: Integer;
  LaTablette, LaChaine, LaValeur, LePlus, LeChamp: string;
begin
    //PT91 - Début
    // Chargement des libellés des champs libres RH
    TobRH := TOB.Create('LesRH', Nil, -1);
    TobRH.LoadDetailFromSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "RH%"');

    // Chargement des libellés des champs libres CP
    TobCP := TOB.Create('LesCP', Nil, -1);
    TobCP.LoadDetailFromSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "CP%"');
    //PT91 - Fin

    TOB_DesChamps := TOB.Create('La liste des champs libres', nil, -1);
    st := 'SELECT DISTINCT(DH_PREFIXE) FROM DECHAMPS WHERE DH_NOMCHAMP LIKE "P%_TRAVAILN1" OR DH_NOMCHAMP LIKE "P%_LIBREPCMB1"'
      + ' OR DH_NOMCHAMP LIKE "P%_DATELIBRE1" OR DH_NOMCHAMP LIKE "P%_BOOLLIBRE1" OR DH_NOMCHAMP LIKE "P%_SALAIRANN1"'
      + ' OR DH_NOMCHAMP LIKE "P%_SALAIREMOIS1" OR DH_NOMCHAMP LIKE "P%_FORMATION1" OR DH_NOMCHAMP LIKE "P%_TABLELIBRERH1"'
      + ' OR DH_NOMCHAMP LIKE "P%_TABLELIBRECR1" OR DH_NOMCHAMP="PFO_CIF"'
      + ' GROUP BY DH_PREFIXE';     //PT91
    {Flux optimisé
      Q := OpenSql(st, TRUE);
      TOB_DesChamps.LoadDetailDB('DECHAMPS', '', '', Q, FALSE);
      ferme(Q);
    }
    TOB_DesChamps.LoadDetailDBFROMSQL('DECHAMPS', st);

    T1 := TOB_DesChamps.FindFirst([''], [''], TRUE);
    while T1 <> nil do
    begin
      Prefixe := T1.GetValue('DH_PREFIXE');
      iTable := PrefixeToNum(Prefixe);
      ChargeDeChamps(iTable, Prefixe);
      {$IFNDEF EAGLSERVER}
      // XP 10.10.2007 Optimisation !!! Et on pourrait continuer !!!
      for iChamp := 1 to high(V_PGI.DeChamps[iTable]) do
      begin
        //PT91 - Début
        // Traitement des champs organisation
        if Pos(Prefixe + '_TRAVAILN', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PGNbreStatOrg >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGLIBORGSTAT'+IntToStr(zz))
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        // Traitement des champs combo Libre
        if Pos(Prefixe + '_LIBREPCMB', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PgNbCombo >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGLIBCOMBO'+IntToStr(zz))
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        // Traitement des champs date Libre
        if Pos(Prefixe + '_DATELIBRE', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PgNbDate >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGLIBDATE'+IntToStr(zz))
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        // Traitement des champs boolean Libre
        if Pos(Prefixe + '_BOOLLIBRE', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PgNbCoche >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGLIBCOCHE'+IntToStr(zz))
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        // Traitement des champs salaires mensuels
        if Pos(Prefixe + '_SALAIREMOIS', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PgNbSalLib >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGSALLIB'+IntToStr(zz)) + '(mensuel)'
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        // Traitement des champs salaires annuels
        if Pos(Prefixe + '_SALAIRANN', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.PgNbSalLib >= zz then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGSALLIB'+IntToStr(zz)) + '(annuel)'
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        if Pos(Prefixe + '_FORMATION', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            // Traitement des champs formation
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            if VH_Paie.NbFormationLibre >= zz Then
                V_PGI.DEChamps[iTable, iChamp].Libelle := GetParamSoc('SO_PGFForlibre' + IntToStr(zz))
            else
                V_PGI.DEChamps[iTable, iChamp].Libelle := V_PGI.DEChamps[iTable, iChamp].Libelle + ' Non utilisé';
        end else
        if Pos(Prefixe + '_TABLELIBRERH', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            // Traitement des champs libres RH
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            //V_PGI.DEChamps[iTable, iChamp].Libelle := RECHDOM('GCZONELIBRE', 'RH' + IntToStr(zz), False);
            If TobRH.FindFirst(['CC_CODE'],['RH' + IntToStr(zz)], False) <> Nil Then
                  V_PGI.DEChamps[iTable, iChamp].Libelle := TobRH.FindFirst(['CC_CODE'],['RH' + IntToStr(zz)], False).GetValue('CC_LIBELLE');
        end else
        if Pos(Prefixe + '_TABLELIBRECR', V_PGI.DEChamps[iTable, iChamp].Nom) > 0 then
        begin
            // Traitement des champs libres CR
            If Not IsNumeric(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1)) Then Continue;
            zz := StrToInt(RightStr(V_PGI.DEChamps[iTable, iChamp].Nom, 1));
            //V_PGI.DEChamps[iTable, iChamp].Libelle := RECHDOM('GCZONELIBRE', 'CP' + IntToStr(zz), False);
            If TobCP.FindFirst(['CC_CODE'],['CP' + IntToStr(zz)], False) <> Nil Then
                V_PGI.DEChamps[iTable, iChamp].Libelle := TobCP.FindFirst(['CC_CODE'],['CP' + IntToStr(zz)], False).GetValue('CC_LIBELLE');
        end else
        if V_PGI.DEChamps[iTable, iChamp].Nom = 'PFO_CIF' then
        begin
          V_PGI.DEChamps[iTable, iChamp].Libelle := 'DIF';
        end;
        //PT91 - Fin
        //FIN PT26
      end;
      {$ENDIF}
      T1 := TOB_DesChamps.FindNext([''], [''], TRUE);
    end;

    TOB_DesChamps.Free;
    FreeAndNil(TobCP);//PT91
    FreeAndNil(TobRH);//PT91

    //FIN PT56
    // FIN PT49

  // Reste les champs contrats de la table salarié   A faire et à verifier
  st := 'PGORGAFFILIATION';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'POG_NUMAFFILIATION||" "||POG_LIBELLE';
  end;
  // PT33  : 31/01/2003 PH V591 Modif tablette PGSALPREPAUTO pour gérer PT8 de UTofPG_MulPrepAuto
  st := 'PGSALPREPAUTO';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].Where := '&#@';
  end;
  //PT40  : 04/09/2003 PH V_421 Chgt suffixe tablettes CODEEMPLOI dans le cas passage 2003
  ChangePCS82A03();
  // DEB PT61 Filtrage des tablettes pour la confidentialité établissement
  st := 'PGSALARIE';
  iChamp := TTToNum(St);
  // DEB PT74
  PgRendHabilitation(); // Fonction d'initialisation de l'objet habilitation
  if Assigned(MonHabilitation) and not (MonHabilitation.active) then
  begin
    TheEtab := PGRendEtabUser();
    if TheEtab <> '' then TheEtab := ' AND PSA_ETABLISSEMENT="' + TheEtab + '"';
  end
  else
  begin // On est dans le cas de la gestion des habilitations on va fitrer les tablettes salaries avec les critères
    TheEtab := MonHabilitation.LeSQL;
  end;
  // FIN PT74
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM';
    if TheEtab <> '' then
      V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab);
  end;
  st := 'PGSALARIECONGES';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM';
    if TheEtab <> '' then
      V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab);
  end;
  st := 'PGSALARIEDEPORTE';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM';
    if TheEtab <> '' then
      // Pour cette tablette, rajouter le tri qui a été supprimé dans DECLA
      V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab + 'ORDER BY PSA_LIBELLE');
  end;
  st := 'PGSALPREPAUTO';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM';
    if TheEtab <> '' then
      V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab);
  end;

  if Assigned(MonHabilitation) then
  begin
    st := 'PGSALARIELOOKPIPE';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
  //    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM';
      if TheEtab <> '' then
        V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab);
    end;

    st := 'PGSALARIELOOK';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
  //    V_PGI.DECombos[iChamp].ChampLib := 'PSA_LIBELLE||" "||PSA_PRENOM Nom';
      if TheEtab <> '' then
        V_PGI.DECombos[iChamp].Where := PgRemplitCLauseEtab(V_PGI.DECombos[iChamp].Where, TheEtab);
    end;
  end;

  st := 'JUTYPECIVIL';
  iChamp := TTToNum(St);
  if iChamp > 0 then
  begin
    V_PGI.DECombos[iChamp].ChampLib := 'JTC_TYPECIV CIVILITE';
  end;
// DEB PT74 Gestion des habilitations
  if MonHabilitation.Active then pgidentpop(MonHabilitation.codepop, MonHabilitation.population); // Remplissage de la liste des valeurs autorisées
  for zz := 0 to MonHabilitation.MesDroits.Count - 1 do
  begin
    st := MonHabilitation.MesDroits[zz]; //PT88 [i]
    LaTablette := ReadTokenst(St);
    iChamp := TTToNum(LaTablette);
    LeChamp := ReadTokenst(St); // Pour passer au champ suivant car le champ indique le champ à controler et non la valeur retournée de la tablette
    LeChamp := V_PGI.DECombos[iChamp].code;
    if iChamp > 0 then
    begin
      if V_PGI.DECombos[iChamp].Where <> '' then LePlus := ' AND (' + LeChamp
      else LePlus := '';
      LaChaine := ' IN (';
      i := 0;
      while st <> '' do
      begin
        i := i + 1;
        LaValeur := ReadTokenst(St);
        if i > 1 then LaChaine := LaChaine + ',';
        LaChaine := LaChaine + '"' + LaValeur + '"';
      end;
      if i > 0 then
      begin
        LaChaine := LePlus + LaChaine + '))';
        V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where + LaChaine;
      end;
    end;
  end;
  if MonHabilitation.Active then
  begin
    st := 'PGETABCONGES';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
      LesEtab := MonHabilitation.LesEtab;
      St := ReadTokenst(LesEtab);
      if st <> '' then
      begin
        LaChaine := ' ETB_ETABLISSEMENT IN (';
        i := 0;
        while st <> '' do
        begin
          i := i + 1;
          LaValeur := ReadTokenst(St);
          if i > 1 then LaChaine := LaChaine + ',';
          LaChaine := LaChaine + '"' + LaValeur + '"';
        end;
        if i > 0 then
        begin
          LaChaine := LaChaine + ')';
          if V_PGI.DECombos[iChamp].Where = '' then V_PGI.DECombos[iChamp].Where := LaChaine
          else V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where + ' AND ' + LaChaine;
        end;
      end;
    end;
    st := 'TTETABABR';
    iChamp := TTToNum(St);
    if iChamp > 0 then
    begin
      LesEtab := MonHabilitation.LesEtab;
      St := ReadTokenst(LesEtab);
      if st <> '' then
      begin
        LaChaine := ' ET_ETABLISSEMENT IN (';
        i := 0;
        while st <> '' do
        begin
          i := i + 1;
          LaValeur := ReadTokenst(St);
          if i > 1 then LaChaine := LaChaine + ',';
          LaChaine := LaChaine + '"' + LaValeur + '"';
        end;
        if i > 0 then
        begin
          LaChaine := LaChaine + ')';
          if V_PGI.DECombos[iChamp].Where = '' then V_PGI.DECombos[iChamp].Where := LaChaine
          else V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where + ' AND ' + LaChaine;
        end;
      end;
    end;
  end
  else
  begin
    if VH^.ProfilUserC[prEtablissement].ForceEtab then
    begin
      st := 'TTETABLISSEMENT';
      iChamp := TTToNum(St);
      if iChamp > 0 then
      begin
        if TheEtab <> '' then V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where
          + ' AND ET_ETABLISSEMENT="' + PGRendEtabUser() + '"';
      end;

      st := 'PGETABCONGES';
      iChamp := TTToNum(St);
      if iChamp > 0 then
      begin
        if TheEtab <> '' then V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where
          + ' AND ETB_ETABLISSEMENT="' + PGRendEtabUser() + '"';
      end;

      st := 'TTETABABR';
      iChamp := TTToNum(St);
      if iChamp > 0 then
      begin
        if TheEtab <> '' then V_PGI.DECombos[iChamp].Where := V_PGI.DECombos[iChamp].Where
          + ' AND ET_ETABLISSEMENT="' + PGRendEtabUser() + '"';
      end;
    end;
  // FIN PT61
  end;
// FIN PT74
end;

//PT20
//PT40  : 04/09/2003 PH V_421 Chgt suffixe tablettes CODEEMPLOI dans le cas passage 2003
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 04/09/2003
Modifié le ... :   /  /
Description .. : Changement des suffixes des tablettes CODE EMPLOI
Suite ........ : Passage norme 1982-2003
Mots clefs ... : PAIE
*****************************************************************}

procedure ChangePCS82A03();
var
  St: string;
  iChamp: integer;
begin
  if VH_Paie.PGPCS2003 then
  begin
    st := 'PGCODEEMPLOI';
    iChamp := TTToNum(St);
    if iChamp > 0 then V_PGI.DECombos[iChamp].champ := '';
    st := 'PGCODEPCSESE';
    iChamp := TTToNum(St);
    if iChamp > 0 then V_PGI.DECombos[iChamp].champ := V_PGI.DECombos[iChamp].champ + 'CODEEMPLOI;';
  end;
end;
// FIN PT40

// PT32  : 09/01/2003 PH V591 Gestion dynamique des modules de la paie en fonction des serias
// PT44  30/01/2004 PH V_50 Mise en place du MENU 48

procedure ChargeModules_PAIE();
var
  LesModules, LesImages: array of integer;
  i, j, k, max: integer;
  OkOk: Boolean;
const
{$IFNDEF RHSEUL}
{$IFDEF PRESENCE} //PT86
  PaieModules: array[0..6] of integer = (42, 41, 46, 49, 43, 310, 347); { PT78 Ajout 347 }
{$ELSE}
  PaieModules: array[0..5] of integer = (42, 41, 46, 49, 43, 310);
{$ENDIF PRESENCE}
{PT86
  PaieModulesPCL: array[0..5] of integer = (42, 41, 46, 49, 43, 347);
}
  PaieModulesPCL: array[0..4] of integer = (42, 41, 46, 49, 43);
//FIN PT86
  PaieImages: array[0..6] of integer = (42, 41, 46, 49, 43, 0, 80); { PT78 Ajout 9 }
{PT86
  ModulesCommuns: array[0..5] of integer = (42, 41, 46, 49, 310, 347);
}
  ModulesCommuns: array[0..4] of integer = (42, 41, 46, 49, 310);
//FIN PT86
{$ELSE}
{PT86
  PaieModules: array[0..6] of integer = (47, 48, 44, 303, 49, 310, 347);
  PaieModulesPCL: array[0..5] of integer = (47, 48, 44, 303, 49, 347);
  PaieImages: array[0..6] of integer = (47, 48, 44, 3, 49, 0, 80);
  ModulesCommuns: array[0..2] of integer = (49, 310, 347);
}
  PaieModules: array[0..5] of integer = (47, 48, 44, 303, 49, 310);
  PaieModulesPCL: array[0..4] of integer = (47, 48, 44, 303, 49);
  PaieImages: array[0..5] of integer = (47, 48, 44, 3, 49, 0);
  ModulesCommuns: array[0..1] of integer = (49, 310);
//FIN PT86
{$ENDIF}
  ModuleECab = [43];
  ModuleForm = [47];
  ModuleComp = [48]; // PT47 Module gestion de carrière et des compétences
  ModuleBilan = [44]; // PT47 Module Bilan social détaillé et tableaux de bord
begin
{$IFNDEF EAGLSERVER}
  if V_PGI.ModePcl <> '1' then
  begin
    max := High(PaieModules);
    SetLength(LesModules, High(PaieModules) + 1);
    SetLength(LesImages, High(PaieModules) + 1);
  end
  else
  begin
    max := High(PaieModulesPCL);
    SetLength(LesModules, High(PaieModulesPCL) + 1);
    SetLength(LesImages, High(PaieModulesPCL) + 1);
  end;
{$ELSE}
  max := High(PaieModules);
  SetLength(LesModules, High(PaieModules) + 1);
  SetLength(LesImages, High(PaieModules) + 1);
{$ENDIF}
  J := 0;
  for i := 0 to max do
  begin
    OkOK := False;
    for k := 0 to High(ModulesCommuns) do
    begin
      if (PaieModules[i] = ModulesCommuns[k]) then OkOk := True;
    end;
    if (PaieModules[i] in ModuleECab) and (VH_Paie.PgeAbsences) then OkOk := True;
    if (PaieModules[i] in ModuleForm) and (VH_Paie.PgSeriaFormation) then OkOk := True;
    if (PaieModules[i] in ModuleComp) and (VH_Paie.PgSeriaCompetence) then OkOk := True; // PT47
    if (PaieModules[i] in ModuleBilan) and (VH_Paie.PgSeriaBilan) then OkOk := True; // PT68

    if (not OkOk) and (V_PGI.VersionDemo) and (not VH_Paie.PgSeriaPaie) then Okok := TRUE; // Version de démo
{PT86
    if (PaieModules[i] = 303) and (VH_Paie.PgSeriaCompetence) then OkOk := True; //PT60
}
    if (PaieModules[i] = 303) and (VH_Paie.PgSeriaMasseSalariale) then OkOk := True; //PT60

    if (PaieModules[i] = 303) and (FileExists('UDC.CEG')) then OkOk := True; //PT87

//PT72
    if (PaieModules[i] = 310) then
      OkOk := False;
//FIN PT72

    if (PaieModules[i] = 347) and (VH_Paie.PgSeriaPresence) then OkOk := True; { PT78 }

    if (PaieModules[i] = 347) and (FileExists('UDC.CEG')) then OkOk := True; //PT87

{ PT68
    if (PaieModules[i] in ModuleBilan) then
    begin
      if (VH_Paie.PgSeriaBilan) or V_PGI.Sav then OkOk := True // PT47
      else OkOk := FALSE
    end;
   }
    if OkOk then
    begin
      LesModules[j] := PaieModules[i];
      LesImages[j] := PaieImages[i];
      inc(J);
    end;
  end;
  SetLength(LesModules, J);
  SetLength(LesImages, J);
{$IFNDEF EAGLSERVER}
  FMenuG.SetModules(LesModules, LesImages);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 22/01/2003
Modifié le ... :   /  /
Description .. : Renvoie True si période de paie cloturer
Mots clefs ... : PAIE;CLOTURE
*****************************************************************}

function PGIsPeriodeClot(Tob_Exercice: TOB; DatePer: TDateTime): Boolean;
var
  T: TOB;
  i: Integer;
  Cloture: string;
  JJ, MM, YY: Word;
begin
  result := False;
  if Tob_Exercice = nil then exit;
  Cloture := '';
  for i := 0 to Tob_Exercice.Detail.count - 1 do
  begin
    T := Tob_Exercice.Detail[i];
    if (T.GetValue('PEX_DATEDEBUT') <= DatePer) and (T.GetValue('PEX_DATEFIN') >= DatePer)
      and (T.GetValue('PEX_CLOTURE') <> '') then
    begin
      Cloture := T.GetValue('PEX_CLOTURE');
      DeCodeDate(DatePer, JJ, MM, YY);
      if Length(Cloture) >= MM then Result := (Cloture[MM] = 'X');
      Break;
    end
  end;
end;

//PT34  : 02/04/2003 PH V_421 Fonction controle de cohérence et recup modif sur NODOSSIER

procedure ReparationNoDossier();
var
  St, St1, LATABLE, Pref, LaCle, LaCle1, Champ: string;
  Q, Q1: TQuery;
  rep, i, j, k: Integer;
begin
{
  AccesPredefini('CEG', CEG, STD, DOS);
  if not CEG then
  begin
    PgiBox('Vous n''avez pas les droits pour utiliser cette fonctionnalité', 'Utilitaire de contrôle');
    exit;
  end;
  }
    rep := PGIAsk('Voulez vous exécuter l''utilitaire de contrôle de cohérence', 'Utilitaire de contrôle');
    if Rep <> mrYes then exit;
    i := 0;
    rep := PGIAsk('Voulez vous corriger les erreurs éventuelles en même temps', 'Utilitaire de contrôle');
    if Rep = mrCancel then exit;
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  Q := OpenSQL('select * from detables where dt_commun="S" and DT_DOMAINE="P"', TRUE);
  InitMoveProgressForm(nil, 'Contrôle en cours ...', 'Veuillez patienter SVP ...', 30, FALSE, TRUE);
  while not Q.EOF do
  begin
    LATABLE := Q.FindField('DT_NOMTABLE').AsSTring;
    if (LATABLE <> 'PAIEPARIM') and (LATABLE <> 'DADSLEXIQUE') and (LATABLE <> 'BASEBRUTESPEC') then
    begin
      Pref := Q.FindField('DT_PREFIXE').AsSTring;
      LaCle := Q.FindField('DT_CLE1').AsSTring;
      LaCle1 := ReadTokenPipe(LaCle, ','); // Elimination des 2 1er champs de la cle qui est tjrs NODOSSIER ET PREDEFINI
      LaCle1 := ReadTokenPipe(LaCle, ','); // car traites dans la requete
      if rep = mrYes then
        st := 'UPDATE ' + LATABLE + '  set ' + LATABLE + '.' + PREF + '_NODOSSIER="000000"'
      else st := 'SELECT COUNT(*) NBRE FROM ' + LATABLE;
      St1 := 'SELECT COUNT(*) NBRE FROM ' + LATABLE;
      St1 := St1 + ' WHERE (' + LATABLE + '.' + Pref + '_NODOSSIER = "" OR ' + LATABLE + '.' + Pref + '_NODOSSIER IS NULL OR ' + LATABLE + '.' + Pref + '_NODOSSIER <> "000000")';
      st := st + ' WHERE (' + LATABLE + '.' + Pref + '_NODOSSIER = "" OR ' + LATABLE + '.' + Pref + '_NODOSSIER IS NULL OR ' + LATABLE + '.' + Pref + '_NODOSSIER <> "000000")';
      st := st + ' AND NOT EXISTS (SELECT TA1.' + PREF + '_PREDEFINI FROM ' + LATABLE + ' TA1 WHERE TA1.' + Pref + '_NODOSSIER="000000" ' +
        'AND TA1.' + PREF + '_PREDEFINI=' + LATABLE + '.' + PREF + '_PREDEFINI ';
      while LaCle <> '' do
      begin
        j := Pos(',', LaCle);
        if j > 0 then Champ := Copy(LaCle, 1, j - 1)
        else Champ := LaCle;
        if (POS('NOSSIER', Champ) < 1) or (POS('PREDEFINI', Champ) < 1) then
          St := St + ' AND ' + LATABLE + '.' + Champ + ' = TA1.' + Champ;
        LaCle1 := ReadTokenPipe(LaCle, ','); // champ suivant
      end; // boucle sur les champs de la clé
      V_PGI.DebugSQL := FALSE;
      V_PGI.Debug := FALSE;
      st := st + ')';
      Q1 := OpenSQL(St1, TRUE);
      if not Q1.EOF then k := Q1.FindField('NBRE').AsInteger
      else k := 0;
      Ferme(Q1);
      j := 0;
      if k > 0 then
      begin
        if rep = mrNo then
        begin
          Q1 := OpenSQL(St, TRUE);
          if not Q1.EOF then j := Q1.FindField('NBRE').AsInteger
          else j := 0;
          Ferme(Q1);
        end
        else j := ExecuteSql(st);
      end;
      if (rep = mrNo) and (k > 0) then
      begin
        V_PGI.DebugSQL := TRUE;
        V_PGI.Debug := TRUE;

        Debug(Q.FindField('DT_LIBELLE').AsSTring + '==> ' + IntToStr(k) + ' Ligne(s) en anomalie');
        if j > 0 then Debug(IntToStr(j) + ' Ligne(s) pouvant être corrigée(s)');
        Debug('-------------------------------------------------------------------------------------');
        i := i + k;
      end
      else i := i + j;
      MoveCurProgressForm(Q.FindField('DT_LIBELLE').AsSTring + ' analysée ');
    end; // Fin du if
    Q.NEXT;
  end;
  ferme(Q);
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  FiniMoveProgressForm;
  if i > 0 then
    PgiBox(IntToStr(i) + ' enregistrement(s) incohérent(s) ont été trouvé(s)', 'Utilitaire de contrôle')
  else PgiBox('Fin du traitement, aucune anomalie détectée', 'Utilitaire de contrôle')
end;

function PGGetControl(ControlName: string; Ecran: TForm): TControl;
begin
  Result := nil;
  if Ecran = nil then exit;
  Result := TControl(Ecran.FindComponent(ControlName));
end;

{ DEB PT69 }

procedure PgGetZoomPaieEdt(Qui: integer; Quoi: string);
var
  predefini, Dossier, param, nat, LaFiche: string;
  j: integer;
  Ok: Boolean;
begin
{$IFNDEF EAGLSERVER}
  Ok := False;
  Param := Quoi;
  if (Param = '') then exit;
  j := Pos(';', Param);
  if j > 0 then Ok := True;
  if Ok = True then
  begin
    if (Copy(Param, 1, 3) = 'CEG') or (Copy(Param, 1, 3) = 'STD') or (Copy(Param, 1, 3) = 'DOS') then
      Predefini := ReadTokenSt(Param);
    // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
    if Predefini = 'DOS' then Dossier := PgRendNoDossier()
    else dossier := '000000';
    Quoi := ReadTokenSt(Param);
    Nat := ReadTokenSt(Param);
  end;

  case Qui of
    440: begin
        LaFiche := 'VARIABLE_';
        if (Nat = 'CUM') or (Nat = 'COT') or (Nat = 'REM') then LaFiche := LaFiche + 'CUM'
        else if (Nat = 'VAL') or (Nat = 'ALE') then LaFiche := LaFiche + 'VAL'
        else LaFiche := LaFiche + Nat;
        AglLanceFiche('PAY', LaFiche, '', Predefini + ';' + Dossier + ';' + Quoi, Nat + ';CONSULTATION');
      end;
    450: AglLanceFiche('PAY', 'SALARIE', '', Quoi, 'ACTION=CONSULTATION');
    460: AglLanceFiche('PAY', 'CUMULS', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    470..471:
      begin
        if (Nat = 'COT') or (Nat = 'BAS') then
          AglLanceFiche('PAY', 'COTISATION', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION')
        else
          if (Nat = 'AAA') then
            AglLanceFiche('PAY', 'REMUNERATION', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
      end;
    480: AglLanceFiche('PAY', 'PROFIL', '', Predefini + ';' + Dossier + ';' + Quoi, 'ACTION=CONSULTATION');
    490: AglLanceFiche('PAY', 'ORGANISME', '', Quoi, 'ACTION=CONSULTATION');

  end;
{$ENDIF}
end;
{ FIN PT69 }
// DEB PT76

procedure ReaffectationNoDossier();
var
  St, LeDossier, LATABLE, Pref : string;
  Q, Q1: TQuery;
  rep, i, j, k: Integer;
  T_DOS, T_TAB, T_DUPL, T1: TOB;
  RedirC: Boolean;
begin
  st := 'SELECT DS_NOMBASE FROM DESHARE WHERE DS_NOMTABLE = "COTISATION"';
  Q := OpenSql(st, TRUE);
  if not Q.EOF then RedirC := TRUE
  else RedirC := FALSE;
  Ferme(Q);
  T_DOS := TOB.create('Mes lignes dossier', nil, -1);
  if not RedirC then st := 'select DOS_NODOSSIER from dossier WHERE DOS_NODOSSIER = "' + V_PGI.NoDossier + '"'
  else st := 'select DOS_NODOSSIER from dossier WHERE DOS_NODOSSIER <> ""';
  Q := OpenSql(st, TRUE);
  T_DOS.LoadDetailDB('DOSSIER', '', '', Q, FALSE);
  Ferme(Q);
  if T_DOS.detail.count < 1 then
  begin
    PgiInfo('Vous n''êtes pas en configuration multi-dossier, Abandon du traitement', 'Utilitaire de contrôle');
    exit;
  end;
  rep := PGIAsk('Voulez vous exécuter l''utilitaire de contrôle de cohérence multi-dossier', 'Utilitaire de contrôle');
  if Rep <> mrYes then exit;
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  Q := OpenSQL('select * from detables where dt_commun="S" and DT_DOMAINE="P"', TRUE);
  InitMoveProgressForm(nil, 'Contrôle en cours ...', 'Veuillez patienter SVP ...', 30, FALSE, TRUE);
  while not Q.EOF do
  begin
    LATABLE := Q.FindField('DT_NOMTABLE').AsSTring;
    if (LATABLE <> 'PAIEPARIM') then
    begin
      Pref := Q.FindField('DT_PREFIXE').AsSTring;
      st := 'SELECT COUNT(*) NBRE FROM ' + LATABLE + ' WHERE ' + Pref + '_PREDEFINI="DOS" AND ' + Pref + '_NODOSSIER="000000"';
      Q1 := OpenSQL(St, TRUE);
      k := Q1.FindField('NBRE').AsInteger;
      Ferme(Q1);
      if k > 0 then
      begin // Il y a des enregistrements à traiter
        T_TAB := TOB.Create('MES ENREG', nil, 1);
        st := 'SELECT * FROM ' + LATABLE + ' WHERE ' + Pref + '_PREDEFINI="DOS" AND ' + Pref + '_NODOSSIER="000000"';
        Q1 := OpenSQL(St, TRUE);
        T_TAB.LoadDetailDB(LATABLE, '', '', Q1, FALSE);
        Ferme(Q1);
        T_DUPL := TOB.Create('LES ENREG DUPLIQUER', nil, 1);
        for j := 0 to T_DOS.detail.count - 1 do
        begin // Boucle du la liste des dossiers
          LeDossier := T_DOS.detail[j].GetValue('DOS_NODOSSIER');
          for i := 0 to T_TAB.detail.count - 1 do
          begin // Pour chaque enregistrement, on le duplique avec le bon numéro dossier
            T1 := TOB.create(LATABLE, T_DUPL, -1); // on insere aussi dans la tob des sections la section créée sinon on va l'ecrire x fois
            T1.Dupliquer(T_TAB.detail[i], FALSE, TRUE, TRUE); // donc on duplique
            T1.PutValue(Pref + '_NODOSSIER', LeDossier);
          end;
        end;
        T_DUPL.InsertDB(nil, FALSE);
        T_TAB.DeleteDB();
        FreeAndNil(T_TAB);
        FreeAndNil(T_DUPL);
      end;
      MoveCurProgressForm(Q.FindField('DT_LIBELLE').AsSTring + ' analysée ');
    end; // Fin du if
    Q.NEXT;
  end;
  ferme(Q);
  FreeAndNil(T_DOS);
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  FiniMoveProgressForm;
  PgiBox('Fin du traitement', 'Utilitaire de contrôle');
end;
// FIN PT76

//DEB PT79

function RecupClauseHabilitationLookupList(WhereRenseigne: string): string;
var
  Longueur: integer;
  stHabilitation: string;
begin
  if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
    if copy(MonHabilitation.LeSQL, 1, 3) <> 'PSA' then
    begin
      Longueur := Length(MonHabilitation.LeSQL);
      Longueur := Longueur - 2;
      stHabilitation := MidStr(MonHabilitation.LeSQL, 4, Longueur);
      stHabilitation := 'PSA' + stHabilitation;
      if WhereRenseigne <> '' then
        WhereRenseigne := WhereRenseigne + ' AND ' + stHabilitation
      else
        WhereRenseigne := stHabilitation;
    end
    else
      if WhereRenseigne <> '' then
        WhereRenseigne := WhereRenseigne + ' AND ' + MonHabilitation.LeSQL
      else
        WhereRenseigne := MonHabilitation.LeSQL;
  result := WhereRenseigne;
end;
//FIN PT79
// DEB PT84
function PGGetMonParamsoc(LaBase, LeNom: string): string;
var
  Q: TQuery;
begin
  Q := OpenSql('SELECT SOC_DATA FROM ' + GetBase(LaBase, 'PARAMSOC')
    + ' WHERE SOC_NOM="' + LeNom + '"', true);
  if not Q.Eof then
    result := Q.Fields[0].AsString
  else Result := '';
  Ferme(Q);
end;

procedure PGAffecteNodossierMulti(LaBase, NumDoss : String);
var LATABLE, st, Pref: string;
  k : Integer;
  Q, Q1 : TQuery;
begin
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  Q := OpenSQL('select * from detables where dt_commun="S" and DT_DOMAINE="P"', TRUE);
  InitMoveProgressForm(nil, 'Contrôle en cours de la base '+LaBase, 'Veuillez patienter SVP ...', 30, FALSE, TRUE);
  while not Q.EOF do
  begin
    LATABLE := Q.FindField('DT_NOMTABLE').AsSTring;
    if (LATABLE <> 'PAIEPARIM') then
    begin
      Pref := Q.FindField('DT_PREFIXE').AsSTring;
      LATABLE := GetBase (LaBase, LATABLE);
      st := 'SELECT COUNT(*) NBRE FROM ' + LATABLE + ' WHERE ' + Pref + '_PREDEFINI="DOS" AND ' + Pref + '_NODOSSIER="000000"';
      Q1 := OpenSQL(St, TRUE);
      k := Q1.FindField('NBRE').AsInteger;
      Ferme(Q1);
      if k > 0 then
      begin // Il y a des enregistrements à traiter
        ExecuteSQL ('UPDATE '+LATABLE+ ' SET '+Pref+'_NODOSSIER="'+ NumDoss+ '" WHERE ' + Pref + '_PREDEFINI="DOS" AND ' + Pref + '_NODOSSIER="000000"');
      end;
    end;
  end;
  Ferme(Q);
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  FiniMoveProgressForm;
end;

procedure ReaffNoDossierMulti ();
var UneBase, stLesBases, NumDoss: string;
    Rep : Integer;
    SavEnableDeShare : Boolean;
begin
{$IFNDEF EAGLSERVER}
  rep := PGIAsk('Voulez vous exécuter l''utilitaire de contrôle de cohérence des numéros de dossier multi-dossier', 'Utilitaire de contrôle');
  if Rep <> mrYes then exit;
  SavEnableDeShare := V_PGI.enableDEShare;
  V_PGI.enableDEShare := True;
  try
    stLesBases := GetBasesMS('##MULTISOC', true); //'##MULTISOC' est le nom du regroupement
    // Remplissage table temporaire pour chaque base :
    UneBase := ReadTokenSt(stLesBases);
    if UneBase = '' then // cas pas de regroupement traitement base par base
    begin
      UneBase :=V_PGI.FCurrentAlias;
      if UneBase = '' then
      begin
        PgiError ('Nom de base vide', 'Utilitaire de contrôle');
        exit;
      end;
      NumDoss := PGGetMonParamsoc(UneBase, 'SO_NODOSSIER');
      if (NumDoss ='') OR (NumDoss ='000000') then
      begin
        PgiInfo ('Numéro de dossier = '+NumDoss+'. Traitement inutile.','Utilitaire de contrôle');
        exit;
      end;
    end;

    while UneBase <> '' do
    begin
      NumDoss := PGGetMonParamsoc(UneBase, 'SO_NODOSSIER');
      PGAffecteNodossierMulti(UneBase, NumDoss);
      UneBase := ReadTokenSt(stLesBases);
    end;
  finally
    V_PGI.enableDEShare := SavEnableDeShare;
  end;
  PgiBox('Fin du traitement', 'Utilitaire de contrôle');
// FIN PT84
{$ENDIF}
  end;

//PT85
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 27/06/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;TIERS;SALARIES
*****************************************************************}
procedure SynchroniseTiers (CodeTiers : string);
var
FTob : TOB;
Q1 : TQuery;
Auxi, Libell, Preno, St, T_AUXILIAIRE   :string;
LongLib, LongPre : integer;
begin
//Salarié vers Tiers
St:= 'SELECT *'+
     ' FROM TIERS WHERE'+
     ' T_TIERS="'+CodeTiers+'" AND'+
     ' T_NATUREAUXI="SAL"';
Q1:= OpenSQL (St, TRUE);
if Q1.eof then
   begin
   Ferme (Q1);
   exit;
   end;
Auxi:= Q1.FindField ('T_AUXILIAIRE').AsString ;
T_AUXILIAIRE:= '"'+Auxi+'"';

FTob:= TOB.Create ('TIERS', Nil, -1) ;
FTob.SelectDB (T_AUXILIAIRE, Q1, TRUE) ;
Ferme (Q1);

St:= 'SELECT PSA_LIBELLE, PSA_PRENOM, PSA_ADRESSE1, PSA_ADRESSE2,'+
     ' PSA_ADRESSE3, PSA_CODEPOSTAL, PSA_VILLE, PSA_TELEPHONE'+
     ' FROM SALARIES WHERE'+
     ' PSA_AUXILIAIRE="'+CodeTiers+'"';
Q1:= OpenSQL (St, TRUE);
if not Q1.eof then
   begin
   Libell:= Q1.FindField ('PSA_LIBELLE').AsString;
   LongLib:= Length(Libell);
   Preno:= Q1.FindField('PSA_PRENOM').AsString;
   LongPre:= Length(Preno);
   Libell:= Copy (Libell, 1, LongLib)+' '+Copy (Preno, 1, LongPre);
   FTob.PutValue ('T_LIBELLE', Copy (Libell, 1, 35));
   FTob.PutValue ('T_ABREGE', Copy (Libell, 1, 17));
   FTob.PutValue ('T_ADRESSE1', Q1.FindField ('PSA_ADRESSE1').Asstring);
   FTob.PutValue ('T_ADRESSE2', Q1.FindField ('PSA_ADRESSE2').Asstring);
   FTob.PutValue ('T_ADRESSE3', Q1.FindField ('PSA_ADRESSE3').Asstring);
   FTob.PutValue ('T_CODEPOSTAL', Q1.FindField ('PSA_CODEPOSTAL').Asstring);
   FTob.PutValue ('T_VILLE', Q1.FindField ('PSA_VILLE').Asstring);
   FTob.PutValue ('T_TELEPHONE', Q1.FindField ('PSA_TELEPHONE').Asstring);
   FTob.InsertOrUpdateDB (True);
   end;
FreeAndNil (FTob);
Ferme (Q1);
end;
//FIN PT85

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 22/11/2007 / PT90
Modifié le ... :   /  /
Description .. : Rend le dernier mois clos par rapport à l'état de l'exerice
Mots clefs ... :
*****************************************************************}
Function RendDernierMoisClos (EtatCloture : String; Decalage : Boolean) : String;
var i : Integer;
    lib : String;
Begin
    Lib := '';

    i := Pos('-', EtatCloture);
    If i = 0 Then i := 13;
    i := i - 1;

    If Not Decalage Then
    Begin
        Lib := RechDom('PGMOIS', Format('%2.2d',[i]), False);
    End
    Else
    Begin
        If i=1 Then i:=12
        Else i := i-1;
        Lib := RechDom('PGMOIS', Format('%2.2d',[i]), False);
    End;

    Result := Lib;
End;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 03/04/2008 / PT94
Modifié le ... :   /  /    
Description .. : Adapte la liste de choix d'une combo "Prédéfini"
Mots clefs ... : 
*****************************************************************}
Procedure GereAccesPredefini (Combo : THValComboBox);
var CEG,STD,DOS : Boolean;
    Num         : Integer;
Begin
	AccesPredefini('TOUS', CEG, STD, DOS);
	If Not CEG then
	Begin
		Num := Combo.Values.IndexOf('CEG');
		If Num >= 0 Then Combo.Values.Delete(Num); Combo.Items.Delete(Num);
	End;
	if Not STD then
	Begin
		Num := Combo.Values.IndexOf('STD');
		If Num >= 0 Then Combo.Values.Delete(Num); Combo.Items.Delete(Num);
	End;
	if Not DOS then
	Begin
		Num := Combo.Values.IndexOf('DOS');
		If Num >= 0 Then Combo.Values.Delete(Num); Combo.Items.Delete(Num);
	End;
End;

{$IFNDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007 / PT92
Modifié le ... :   /  /
Description .. : Chargement de la liste des dossiers d'un regroupement
Suite ........ : dans une grille
Mots clefs ... :
Paramètres ... : Mode : Modes gérés : <vide>, EXERSOCIAL, VIREMENT
.............. : AnneeRef : Sert pour le mode EXERSOCIAL
.............. : Selection : Liste des éléments à sélectionner par défaut (forme : val1;val2)
*****************************************************************}
Function ChargeDossiersDuRegroupement (CodeReg : String; Grille : THGrid; Mode, AnneeRef, Selection : String) : Boolean;
Var TobDossiers,T       : TOB;
    Q                   : TQuery;
    Detail,Temp         : String;
    StSQL,Bases,Groupes : String;
    AccessGranted,Trouve: Boolean;
    Mois,Annee,DernierMoisRegle			: String;
    J,M,A				: WORD;
    i                   : Integer;
	DatFin				: TDateTime;
    DBPrefixe			: String;
    Etat,Actif,Decalage	: String;
Begin
    Result := True;

	If Grille <> Nil Then
    Begin
        // Initialisations
        StSQL        := '';
        AccessGranted := False;

        // Récupération des bases faisant partie du regroupement
        Q := OpenSQL('SELECT YMD_DETAILS FROM YMULTIDOSSIER WHERE YMD_CODE="'+CodeReg+'"', True);
        If Not Q.EOF Then Detail := Q.FindField('YMD_DETAILS').AsString;
        Ferme(Q);

        // Découpage des informations pour récupérer chaque base
        If Detail <> '' Then
        Begin
            // La zone est composée de 2 parties : liste des bases , retour chariot, liste des groupes
            Bases   := Copy(Detail, 1, Pos(#13,Detail)-1);
            Groupes := Copy(Detail, Pos(#13,Detail)+2, Length(Detail));

            // Dossiers
            Temp := ReadTokenSt(Bases);
            While Temp <> '' Do
            Begin
                StSQL := StSQL + '"' + ReadTokenPipe(Temp, '|') + '",';
                Temp := ReadTokenSt(Bases);
            End;

            // Groupes
            Temp := ReadTokenSt(Groupes);
            While Temp <> '' Do
            Begin
                If V_PGI.Groupe = Temp Then AccessGranted := True;
                Temp := ReadTokenSt(Groupes);
            End;
        End;

        // Pas d'accès à ce regroupement pour l'utilisateur
        If Not AccessGranted Then
        Begin
            PGIError(TraduireMemoire('Vous n''avez pas accès au regroupement spécifié.'), TraduireMemoire('Sélection des dossiers du regroupement'));
            Result := False;
            Exit;
        End;

        // Chargement de la liste des dossiers dans la grille
        If StSQL <> '' Then
        Begin
            TobDossiers := TOB.Create('LesDossiers', Nil, -1);
            TobDossiers.LoadDetailFromSQL('SELECT DOS_NODOSSIER,DOS_LIBELLE FROM DOSSIER WHERE DOS_NOMBASE IN ('+Copy(StSQL, 1, Length(StSQL)-1)+')');

			// Mise à jour de la grille : Mode à vide : Seuls les dossiers sont affichés
			If (Mode = '') Or (Mode = 'REGROUPEMENT') Then
			Begin
	            TobDossiers.PutGridDetail(Grille, False, False, 'DOS_NODOSSIER;DOS_LIBELLE');
	        End
	        // Mode = EXERSOCIAL : On affiche en plus l'exercice social en cours sur chaque dossier
	        Else If Mode = 'EXERSOCIAL' Then
	        Begin
	        	For i := 0 To TobDossiers.Detail.Count-1 Do
	        	Begin
	        		StSQL := 'SELECT PEX_ANNEEREFER, PEX_ACTIF, PEX_CLOTURE FROM '+TobDossiers.Detail[i].GetValue('DOS_LIBELLE')+'.DBO.EXERSOCIAL';
					If AnneeRef<>'' Then StSQL := StSQL + ' WHERE PEX_ANNEEREFER="'+AnneeRef+'"';
	        		StSQL := StSQL + ' ORDER BY PEX_DATEDEBUT DESC';

	        		Q := OpenSQL(StSQL, True);
	        		If Not Q.EOF Then
	        		Begin
	        			TobDossiers.Detail[i].AddChampSupValeur('ANNEE', Q.FindField('PEX_ANNEEREFER').AsString);
	        			Etat  := Q.FindField('PEX_CLOTURE').AsString;
	        			Actif := Q.FindField('PEX_ACTIF').AsString;
	        			Ferme(Q);
	        			
	        			Decalage := '-';
	        	        Q := OpenSQL('SELECT SOC_DATA FROM '+TobDossiers.Detail[i].GetValue('DOS_LIBELLE')+'.DBO.PARAMSOC WHERE SOC_NOM="SO_PGDECALAGE"', True);
                		If Not Q.EOF Then Decalage := Q.FindField('SOC_DATA').AsString;
                		Ferme(Q);

	        			If Actif = 'X' Then Actif:='OUI' Else Actif:='NON';
	        			TobDossiers.Detail[i].AddChampSupValeur('ACTIF',    RechDom('PGOUINON', Actif, False));

	        			If Decalage = 'X' Then Decalage:='OUI' Else Decalage:='NON';
	        			TobDossiers.Detail[i].AddChampSupValeur('DECAL', 	RechDom('PGOUINON', Decalage, False));
	        			TobDossiers.Detail[i].AddChampSupValeur('MOISCLOS', RendDernierMoisClos(Etat, (Decalage='OUI')));
	        		End
	        		Else
	        		Begin
	        			TobDossiers.Detail[i].AddChampSupValeur('ANNEE', 	'');
	        			TobDossiers.Detail[i].AddChampSupValeur('ACTIF',    '');
	        			TobDossiers.Detail[i].AddChampSupValeur('DECAL',	'');
	        			TobDossiers.Detail[i].AddChampSupValeur('MOISCLOS', '');
	        		End;
	        	End;

	        	Grille.ColCount   := 6;
	        	Grille.Cells[2,0] := TraduireMemoire('Exercice');
	        	Grille.Cells[3,0] := TraduireMemoire('Actif');
	        	Grille.Cells[4,0] := TraduireMemoire('Décalage paie');
	        	Grille.Cells[5,0] := TraduireMemoire('Dernier mois clos');
	        	
	        	Grille.ColAligns[2] := taRightJustify;
	        	Grille.ColAligns[3] := taCenter;
	        	Grille.ColAligns[4] := taCenter;
	        	
	        	Grille.ColWidths[0] := 70;
	        	Grille.ColWidths[1] := 100;
	        	Grille.ColWidths[2] := 60;
	        	Grille.ColWidths[3] := 70;
	        	Grille.ColWidths[4] := 70;
	        	Grille.ColWidths[5] := 90;
	        	TobDossiers.PutGridDetail(Grille, False, False, 'DOS_NODOSSIER;DOS_LIBELLE;ANNEE;ACTIF;DECAL;MOISCLOS');
	        End
	        // Mode = VIREMENT : On affiche en plus le dernier mois réglé et le mois à virer
	        Else If Mode = 'VIREMENT' Then
	        Begin
	            For i := 0 To TobDossiers.Detail.Count-1 Do
	            Begin
	                T := TobDossiers.Detail[i];
	            	If V_PGI.ModePcl = '1' Then DBPrefixe := 'DB' Else DBPrefixe := '';
	            	DBPrefixe := DBPrefixe + TobDossiers.Detail[i].GetValue('DOS_LIBELLE') + '.DBO.';
	
		            // Récupération du mois à virer
		            Mois := '';
					Q:= OpenSQL ('SELECT ##TOP 1## PEX_FINPERIODE,PEX_ANNEEREFER FROM '+DBPrefixe+'EXERSOCIAL WHERE'+
		                         ' PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', True);
					If Not Q.EOF Then
					Begin
					   	DatFin := Q.FindField('PEX_FINPERIODE').AsDateTime;
					   	Annee  := Q.FindField('PEX_ANNEEREFER').AsString;
					   	DecodeDate(DatFin, A, M, J);
					   	Mois := ColleZeroDevant(M, 2);
					End;
					Ferme(Q);
					
					// Dernier mois réglé
					DernierMoisRegle := 'Aucun';
					Q:= OpenSQL ('SELECT MAX(PPU_DATEFIN) AS DATEMAX FROM '+DBPrefixe+'PAIEENCOURS WHERE'+
		                         ' PPU_TOPREGLE="X"', True);
					If Not Q.EOF Then
					Begin
					   	DatFin := Q.FindField('DATEMAX').AsDateTime;
	                    If DatFin <> iDate1900 Then
	                    Begin
	    				   	DecodeDate(DatFin, A, M, J);
	                        DernierMoisRegle := RechDom('PGMOIS',ColleZeroDevant(M, 2),False)
	                    End;
					End;
					Ferme(Q);
	
	    			T.AddChampSupValeur('MOIS', Mois);
	    			T.AddChampSupValeur('MOISLIB', RechDom('PGMOIS',Mois,False));
	    			T.AddChampSupValeur('ANNEE', Annee);
	    			T.AddChampSupValeur('DERNIER', DernierMoisRegle);
				End;

				// Mise à jour de la grille : Mode à vide : Seuls les dossiers sont affichés
	            TobDossiers.PutGridDetail(Grille, False, False, 'DOS_NODOSSIER;DOS_LIBELLE;DERNIER;MOISLIB;ANNEE');
	
	            // Adaptation de la taille des colonnes
	            Grille.ColAligns[2] := taLeftJustify;
	            Grille.ColAligns[3] := taLeftJustify;
	            Grille.ColAligns[4] := taRightJustify;
	        End;

            // Libération des objets en mémoire
            FreeAndNil(TobDossiers);

            // Sélection des éléments déjà paramétrés par l'utilisateur (Nième ouverture de l'écran)
            Temp := ReadTokenSt(Selection);
            While Temp <> '' Do
            Begin
                Trouve := False; i := 1;
                While (Not Trouve) And (i < Grille.RowCount) Do
                Begin
                    If Grille.CellValues[0,i] = Temp Then
                    Begin
                    	Grille.FlipSelection(i);
                        Trouve := True;
                    End;
                    i := i + 1;
                End;
                Temp := ReadTokenSt(Selection);
            End;
        End;
    End;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 14/12/2007 / PT92
Modifié le ... :   /  /
Description .. : Lancement d'une nouvelle instance de la même application.
Suite ........ : Penser (au besoin avant d'appeler la fonction) à faire une fenêtre d'attente dépendante de la mainform,
Suite ........ : bloquer les boutons valider, sélectionner tout, etc, et à interdire la fermeture de la fenêtre
Suite ........ : dans le OnClose tant que hProcess <> 0. Ajouter également dans le OnArgument de la fiche à ouvrir
Suite ........ : la ligne VH_PAIE.FicheOuverte := True;
Suite ........ : Exemples : UTOFSalarie_mul, UTOFPGGenerationVirement
Mots clefs ... :
Paramètres ... : ProcessInfo : Element permettant de gérer l'état de l'application lancée
.............. :               (Une fois l'application lancée, la propriété hProcess est <> 0)
.............. : Dossier     : Dossier sur lequel la nouvelle application doit se connecter
.............. : Element     : Paramètre(s) supplémentaire(s) à passer à l'application (ex : /TAG=48220)
.............. : GereMessages: True > Gère les messages interapp
.............. :              (sert notamment à indiquer à l'appli maître l'état du chargement de la nouvelle application)
*****************************************************************}
procedure LanceApplication (var ProcessInfo : TProcessInformation; Dossier, Element : String; GereMessages : Boolean);
Var  StartInfo   : TStartupInfo;
     Fin         : Boolean;
     Chaine      : PAnsiChar;
     lpMsg       : TagMSG;
     Param       : String;
begin
    // Mise à zéro de la structure StartInfo
    FillChar(StartInfo,SizeOf(StartInfo),#0);

    // Seule la taille est renseignée, toutes les autres options
    // laissées à zéro prendront les valeurs par défaut
    StartInfo.cb := SizeOf(StartInfo);

    // Construction de la ligne de commande à exécuter
    {$IFDEF EAGLCLIENT}
    Param := Param + ' /SERVER='+ V_PGI.FTaskServer;
    {$ENDIF}
    Param := Param + ' /USER=' + V_PGI.UserLogin;
    Param := Param + ' /PASSWORD=' + V_PGI.Password;
    Param := Param + ' /DATE=' + DateToStr(V_PGI.DateEntree);

    Param := Param + ' /DOSSIER=' + Dossier;
    If V_PGI.ModePCL = '1' Then
    Begin
        Param := Param + ' /MODEPCL=1';
    End;
    Param := Param + ' '+ Element;

    Chaine := PAnsiChar('"'+ParamStr(0)+'" '+Param);

    // Lancement de la ligne de commande
    // Exemple :
    // /SERVER=SRV-DEV-PAYE /USER=CEGID /PASSWORD=46171FC5C3 /DATE=23/11/2007 /DOSSIER=PAIE750_V800 /FICHE=SALARIE;0000000001
    If CreateProcess(Nil, Chaine, Nil, Nil, False, 0, Nil, Nil, StartInfo,ProcessInfo) Then
    Begin
        // L'application est bien lancée, on va en attendre la fin
        // ProcessInfo.hProcess contient le handle du process principal de l'application
        Fin:=False;

        Repeat
            // On attend la fin de l'application
            Case WaitForSingleObject(ProcessInfo.hProcess, 100) Of
                WAIT_OBJECT_0 :
                    Begin
                    	If GereMessages Then
                        Begin
                            MoveCurProgressForm(TraduireMemoire('Déconnexion'));
                        End;
                        Fin   := True; // L'application est terminée, on sort
                    End;

                WAIT_TIMEOUT  :
                    Begin
                    	If GereMessages Then
                    	Begin
	                        // Au cours du chargement, l'application envoie des messages pour indiquer
	                        // son état d'avancement
	                        If PeekMessage(lpMsg, 0, MessageInterApp, MessageInterApp, PM_REMOVE or PM_NOYIELD) <> False Then
	                        Begin
	                            Case lpMsg.wParam Of
	                            1:
	                                MoveCurProgressForm(TraduireMemoire('Initialisation'));
	                            2:
	                                MoveCurProgressForm(TraduireMemoire('Connexion à la base'));
	                            3:
	                                MoveCurProgressForm(TraduireMemoire('Chargement des paramètres'));

                                // Cas particuliers des fiches lancées
                                10: // Fiche quelconque
	                                MoveCurProgressForm(TraduireMemoire('Chargement des données'));
	                            11: // Fiche Salarié
	                                MoveCurProgressForm(TraduireMemoire('Chargement du salarié'));
                                21: // Fiche Virement_Mul
	                                MoveCurProgressForm(TraduireMemoire('Initialisation des virements'));
	                            end;
	                        end;
	                    end;
                    end;
            End;
            // Mise à jour de la fenêtre pour que l'application ne paraisse pas bloquée.
            Application.ProcessMessages;
        Until Fin;
    End
    Else RaiseLastOSError;

    // Fermeture du handle ouvert
    CloseHandle (ProcessInfo.hProcess);
    ProcessInfo.hProcess := 0;
End;
{$ENDIF}

{$ENDIF} //PLUGIN

function GererCritereGroupeConfTous (FiltreDonnees:string='') : String;
begin
  Result := '';

  if FiltreDonnees='' then  //LM20071008
    begin
    // Dossier sans groupes
    if GetParamsocDpSecur('SO_MDDOSSANSGRP', True) then
      Result := 'NOT EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER) OR ';
    // Dossiers dans des groupes auxquels l'utilisateur en cours a droit
    Result := Result
                  + 'EXISTS (SELECT 1 FROM GRPDONNEES '
                           +'LEFT JOIN LIENDOSGRP ON GRP_NOM=LDO_NOM AND GRP_ID=LDO_GRPID '
                           +'LEFT JOIN LIENDONNEES ON GRP_NOM=LND_NOM AND GRP_ID=LND_GRPID '
                           +'WHERE GRP_NOM="GROUPECONF" AND LDO_NODOSSIER=DOS_NODOSSIER AND LND_USERID="'+V_PGI.User+'")'
    end
  else
    Result := 'not exists (select 1 FROM DOSSIERGRP_FILTREDONNEES ' +
                          'where DOG_NODOSSIER=DOS_NODOSSIER  and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' +
              'or exists (select 1 from DOSSIERGRP_FILTREDONNEES, USERCONF_FILTREDONNEES '+
                         'where UCO_GROUPECONF = DOG_GROUPECONF '+
                         'and UCO_USER="'+ V_PGI.User+'" ' +
                         'and DOG_NODOSSIER=DOS_NODOSSIER '+
                         'and DOG_FILTREDONNEES="'+ FiltreDonnees+'") ' ; //LM20071008
end;

end.

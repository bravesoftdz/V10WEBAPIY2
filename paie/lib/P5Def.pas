{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... : 25/06/2001
Description .. : Unit de définition de fonctions génériques de la paie
Mots clefs ... : PAIE;PGBULLETIN
*****************************************************************}
{
PT1   : 25/06/2001 PH V536 suppression  de Q.Next le cas échéant
PT2   : 21/08/2001 PH V547 rajout date de fin dans les fonctions
                           RechercheRubriqueProfil,RubriqDuProfil,RecupRubProfil
                           pour saisie de reprise des bases de cotisations
PT3   : 07/09/2001 PH V547 Controle des dates de la paie par rapport à une
                           période cloturée
PT4   : 07/09/2001 PH V547 Fonction RendMois mise dans P5DEF existait dans
                           UTofPG_CLOTURE
PT5   : 07/09/2001 PH V547 Test du profil FNAL on recherche le profil FNAL que
                           dans la table Salarie ou dans les paramsoc. test plus
                           sévère
PT6   : 02/10/2001 PH V562 Prise en compte du profil rémunérations
PT7   : 19/03/2002 PH V571 Rajout fonction affichage des champs libres COMBO
                           salarié
PT8   : 02/04/2002 PH V571 Traitement du profil retraite
PT9   : 03/06/2002 PH V582 Gestion historiques des évènements
PT10  : 27/06/2002 PH V582 tests si composant non nul avant affectation
PT11  : 27/06/2002 PH V582 Fonction de visibilité des boites à cocher idem combo
PT12  : 02/07/2002 PH V582 Fonction initialisation de tous les champs TOB
                           Salaries
PT13  : 23/09/2002 PH V585 Suppression des lignes en doublons
PT14  : 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 au
                           lieu de null
PT15  : 03/01/2003 PH V591 FQ10417 initialisation nationalités
PT16  : 02/04/2003 PH V421 Fonction qui rend le numéro de dossier
PT17  : 02/04/2003 PH V421 Fonction de suppression en masse des bulletins de
                            paie
PT18  : 20/05/2003 MF V_42 fct affichage des champs libres salarié  PSA_LIBREx
PT19  : 17/06/2003 PH V421 Fonction qui rend le mode de fonctionnement MONO ou
                            MULTI
PT20  : 04/07/2003    V421 procedure InitSalDef : suppression de
                            PSA_DATEENTREECRC,PSA_DATESORTIECRC
PT21  : 08/09/2003 PH V421 Procedure de chargement de la table de
                            correspondance PCS
PT22  : 23/09/2003 PH V421 Rajout suppression saisie arret dans la suppression
                            en masse des bulletins
PT23  : 09/12/2003 JL V_50 Correction saisie arrêt : calcul dernière mensualité
                           de saisie arrêt si soldé
PT24  : 14/04/2004 MF V_50 Ajout procédure VisibiliteChampRupt
PT25  : 25/05/2004 PH V_50 Rajout fonction création automatique des salariés
                           dans la table INTERIMAIRES
PT26  : 17/06/2004 MF V_50 Ajout traitement de suppression des IJSS qd
                           suppression bulletin (menu Outils)
PT27  : 12/08/2004 JL V_50 FQ 11148 Prise en compte du montant de l'acompte dans
                           le calcul de la saisie arret
PT26  : 17/06/2004 MF V_50 Ajout traitement de suppression des IJSS qd
                           suppression bulletin (menu Outils)
PT28  : 17/06/2004 MF V_50 Ajout traitement de suppression des lignes de
                           maintien qd suppression bulletin (menu Outils)
PT29  : 17/09/2004 JL V_50 Gestion des remboursements partiels pour saisie arrêt
PT30  : 24/09/2004 VG V_50 Déplacement de PGRendNoDossier dans PgOutils2
PT31  : 07/10/2004 JL V_50 Saisie arrêt : pre^t limité a 10% du salaire net
PT32  : 16/02/2005 JL V_60 FQ 12004 Maj champ effectue = - en suppression saisie arret
PT33  : 19/04/2005 PH V_60 Fonction filtrage par droit utilisateur/etablissement
PT34  : 12/07/2005 SB V_60 FQ 12308 Modification de l'ordre de présentation des bulletins
PT35  : 01/08/2005 JL V_60 FQ 12469 Remboursement anticipé saisie arrêt
PT36  : 21/09/2005 PH V_60 Fonctions de suppression du répertoire temporaire de la paie pour la
                           mise en base des fichiers PDF
PT37  : 07/10/2005 PH V_60 FQ 12629 Cas exclusion profil CSG/CRDS deductible et non deductible ensembles
PT38  : 21/10/2005 SB V_65 FQ 12657 Violation accès : Restriction utilisateur
PT39  : 29/11/2005 PH V_65 Procedure gestion message process_serveur
PT40  : 13/01/2006 PH V_65 FQ 12821 Rajout pré-alimentation des etabs dans les éditions de la DADS-U
PT41  : 03/05/2006 PH V_65 prise en compte driver dbMSSQL2005
PT42  : 19/05/2006 SB V_65 Appel des fonctions communes pour identifier V_PGI.driver
PT43  : 27/09/2006 JL V_70 FQ  13537 Modif calcul arriéré saisie arrêt
PT44  : 26/10/2006 PH V_70 Mise en place gestion des habilitations
PT45  : 22/01/2007 PH V_70 Gestion de la rétroactivité
PT46  : 25/01/2007 GGS V_80 Journal Evenement
PT47  : 21/02/2007 NA  V-80 Recherche de la clause SQL pour les populaztions :Le critère CATDADS n'est pas en historique
PT48  : 13/04/2007 FC V_72 FQ 14096 montant saisie-arrêt disparait après validation du bulletin de paie
PT49  : 25/04/2007 JL V_72 Correction fonction duplication interimaire
PT50  : 04/07/2007 MF V_72 FQ 14511 correction access vio sur états chaînés qd habilitation
PT51  : 20/07/2007 FC V_72 FQ 14423 DADS_U rajout de régimes
PT52  : 31/07/2007 JL V_80 FQ 13982 Gestion valeur 90 et 99 pour champ personne a charge
PT53  : 04/09/2007 FC V_80 FQ 14734 Gestion des habilitations, message incohérent sur saisie historique
PT54  : 16/10/2007 GGU V_80 FQ 14498 Solder automatiquement le reste du prêt au solde de tous comptes
PT55  : 17/10/2007 GGU V_80 FQ 14827 Décloturer la saisie arrêt si on supprime une mensualité après cloture
PT56  : 17/10/2007 GGU V_80 Enlever la saisie arrêt des exe GC -> car nécéssite P5Util qui n'est pas en GC
PT57  : 29/10/2007 NA  V-810 Si suppression du bulletin et Gestion de la présence : Mise à jour de la table PRESENCESALARIE
PT58  : 19/11/2007 GGU V_80 Correction d'une fuite mémoire en préparation automatique sur le calcul des saisies arrêt
PT60  : 29/02/2008 GGU V81 FQ 15273 Erreur du calcul lors du solde de tous comptes si on a coché "solde des prêts lors du solde de tout compte"
}
unit P5Def;

interface

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  StdCtrls,
{$IFDEF EAGLCLIENT}
  UtileAGL,
{$ELSE}
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}
  uPaieCotisations,
  uPaieRemunerations,
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA
  Hent1,
  HCtrls,
  HMsgBox,
  ComCtrls,
  HRichEdt,
  HRichOLE,
  ParamSoc,
  ExtCtrls,
  Grids,
  ImgList,
  Mask,
  richedit,
  UTOB,
  EntPaie,
  //     UTOF,
  PgOutils2;

// Définition des différents types de traitements possibles pour la saisie et le calcul du bulletin
type
  TActionBulletin = (PremCreation, taCreation, taModification, taConsultation, taPrepAuto, taCalcul, taEnvers, taSuppression);
// DEB PT44
type
  T_Habilit = class
    LeSQL, LesEtab, SqlPop, LibellePop: string; // chaine SQL
    Active, CritEtab, Habilitation: Boolean;
    CodePop, Population, AvecCreat, LibelleHab: string;
    MesDroits: TStringList;
    LesPrefixes: TStringList;
    function InclureSal(LePrefixe: string): Boolean;
    constructor Create;
    destructor Destroy; override;
  end;

var MonHabilitation: T_Habilit;
  LeModeT: string;
// FIN PT44
  Paie1550: Integer;
function SurchargeCase(Base: string; T: array of string; V: array of integer): integer;
procedure PaieRazPile(Grille: THGrid);
procedure RendMoisAnnee(Zdate: TDateTime; var Mois, Annee: string);
function RendCalendrierMois(Etab, Annee, Mois: string; TOB_JoursTrav: TOB): TOB;
function RendDateExerSocial(DateDebut, DateFin: TDateTime; var DatD, DatF: TDateTime; MoisPris: Boolean = FALSE): Boolean;
function MaxDebutExerSocial(): variant;
// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}
procedure RechercheRubriqueProfil(Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases: TOB; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
// Rend dans TPE la liste des rubiques en provenance des profils salaries
procedure RubriqDuProfil(Champ1, Champ2: string; Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases: TOB; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
// Recherche le profil salarie : Personnalisé ou Idem Etab ...
procedure RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases: TOB; Profil: string; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
// Rend dans TPE la liste des rubiques en provenance d'un profil
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA

procedure VisibiliteChampSalarie(Numero: string; Champ1, Champ2: TControl; Champ3: TControl = nil);
// sur les codes Organisations pour la gestion de l'onglet PCOMPLEMENT des fiches multicritere  salarie
procedure VisibiliteStat(Champ1, Champ2: TControl; Champ3: TControl = nil);
procedure VisibiliteChampLibreX(Numero: string; Champ1, Champ2: TControl); // PT18
function MemCmpPaie(Chaine1, Chaine2: string; Position, Longueur: Integer): Boolean;
//  PT3 07/09/01 V547 PH Controle des dates de la paie par rapport aux clotures
function ControlPaieCloture(DateDebut, DateFin: TDateTime): boolean;
// PT4 07/09/01 V547 PH Fonction RendMois mise dans P5DEF existait dans UTofPG_CLOTURE
function RendMois(Indice: WORD): WORD;
// PT7 19/03/02 V571 PH Rajout fonction affichage des champs libres COMBO salarié
procedure VisibiliteChampLibreSal(Numero: string; Champ1, Champ2: TControl; Champ3: TControl = nil);
//  PT9 03/06/02 V582 PH Gestion historiques des évènements
function CreeJnalEvt(TypeEvt, Evt, Statut: string; Lbox, LboxS: TlistBox; Trace: TStringList = nil; TraceE: TStringList = nil): Boolean;
// PT11 27/06/02 V582 PH Fonction de visibilité des boites à cocher idem combo
procedure VisibiliteBoolLibreSal(Numero: string; Champ1: TControl);
// PT12 02/07/02 V582 PH Fonction initialisation de tous les champs TOB Salaries
procedure InitSalDef(TS, TEtab: TOB; Etab: string);
// PT17  : 02/04/03 V_421 PH Fonction de suppression en masse des bulletins de paie

// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}
procedure SuppressionBull(Zdate1, Zdate2: TDateTime; CodeSalarie, Etab, BullCompl: string; TheCP: BooLean);
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
// PT19  : 17/06/03 V_421 PH Fonction qui rend le mode de fonctionnement MONO ou MULTI
function PGRendModeFonc(): string;
//PT21  : 08/09/03 V_421 PH Procedure de chargement de la table de correspondance PCS
procedure ChargePCS();
{$IFNDEF GCGC}   //PT56
//PT22  : 23/09/03 V_421 PH Rajout suppression saisie arret dans la suppression en masse des bulletins
function PGSuppHistoRetenue(Salarie: string; DateDebPaie, DateFinPaie: TDateTime): Boolean;
function GetTobSaisieArret(Salarie : String; DateFinPaie : TDateTime) : Tob;
function PGCalcSaisieArret(Salarie: string; DateDebPaie, DateFinPaie: TDateTime; NbACharge: Integer; SalaireNet, Acompte: Double): Tob; //PT27
function PGCalcPartieSaisissable(DateDebPaie: TDateTime; NbACharge: Integer; SalaireNet, ValeurRMI: Double; LeType: string): Double;
function PGMajHistoRetenue(TobMajHisto: Tob; DatePaiement: TDateTime): Boolean;
{$ENDIF}
procedure VisibiliteChampBoollX(Numero: string; Champ1: TControl);
procedure VisibiliteChampRupt(Numero: string; Champ1, Champ2: TControl); // PT24

procedure DupliqSalInterimaire;
// DEB PT33 Cette fonction sera appelée par la TOF via un pointeur pour faire un filtre/confidentialité par etab/utilisateur
function PGAffecteEtabByUser(FF: TForm): Boolean; // AffecteFiltre: boolean Affectation etablissement par utilisateur
procedure PGPositionneEtabUser(LeEtab: TControl; IgnoreForce: Boolean = False);
procedure PGEnabledEtabUser(LeEtab: TControl);
function PGRendEtabUser(): string;
// FIN PT33
function PGGetTempPath: string;
procedure PGDeleteRepertTemporaire(sDir: string);
procedure PGDeleteRepertSTDPDF(sDir: string);
function PGRendNbreSal(): Integer;
procedure MessageCgi(S: string); // PT39
// DEB PT44
procedure PGAjoutComposant(THBST: TTabsheet; FF: TForm);
procedure PGrechwhere(Codepop, population: string; histo: boolean; var st1, st2: string; Habilitation : Boolean = FALSE);
function PGConstitutionwhere(T1, T2: Tob): string;
function PGControlHab(FF: TForm): Boolean;
procedure PgRendHabilitation(codePop: string = ''; Population: string = ''; Libellepop: string = '');
procedure PgVideHabilitation();
procedure CreateObjHabilitation();

procedure pgidentpop(Codepop, population: string);
// PT44
{$IFNDEF GCGC}
// PT45 Gestion de la retractivité des cotisations
procedure ChargeRegulCot(LeSalarie, TPE: TOB; DD, DF: TDateTime);
procedure AnnulRegulCot(LeSalarie: string; DD, DF: TDateTime);
procedure PaieConceptPlanPaie(Ecran: TForm);
procedure PaieConceptTabMinPaie(Ecran: TForm);
{$ENDIF}

implementation

uses
// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}
  P5Util,
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
  Ent1,
  HTB97,
  HDebug,
  majtable,
// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}
  PgCongesPayes,
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
  pgcommun;


// DEB PT39

procedure MessageCgi(S: string);
begin
{$IFDEF EAGLSERVER}
  Writeln(s);
  Readln(s);
{$ENDIF}
end;
// FIN PT39
// PT36

function PGGetTempPath: string;
var
  Path: array[0..255] of Char;
begin
  GetTempPath(255, Path); // recup du repertoire Temp -> voir si la methode est la bonne
  Result := StrPas(Path);
  if V_PGI.Debug then debug('Répertoire temporaire=' + Result);
end;

procedure PGDeleteRepertTemporaire(sDir: string);
var
  iIndex: Integer;
  SearchRec: TSearchRec;
  sFileName: string;
begin
  sDir := sDir + '\*.*';
  iIndex := FindFirst(sDir, faAnyFile, SearchRec);
  while iIndex = 0 do
  begin
    sFileName := ExtractFileDir(sDir) + '\' + SearchRec.Name;
    if SearchRec.Attr = faDirectory then
    begin
      if (SearchRec.Name <> '') and
        (SearchRec.Name <> '.') and
        (SearchRec.Name <> '..') then
        PGDeleteRepertTemporaire(sFileName);
    end else
    begin
      if SearchRec.Attr <> 0 then
        FileSetAttr(sFileName, 0);
      Sysutils.DeleteFile(sFileName);
    end;
    iIndex := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);
  RemoveDir(ExtractFileDir(sDir));
end;

procedure PGDeleteRepertSTDPDF(sDir: string);
var
  iIndex, i: Integer;
  SearchRec: TSearchRec;
  sFileName, LeFic: string;
const TableauPdf: array[0..6] of string = ('ACCTRAVAIL.PDF', 'ATTEST_SPECTACLES.PDF', 'ATTEST_SPECTACLES_SALARIE.PDF', 'DECLA_AT.PDF', 'DUCS.PDF', 'DUE.PDF', 'MALADIE.PDF');
begin
  sDir := sDir + '\*.pdf';
  iIndex := FindFirst(sDir, faAnyFile, SearchRec);
  while iIndex = 0 do
  begin
    sFileName := ExtractFileDir(sDir) + '\' + SearchRec.Name;
    if SearchRec.Attr <> 0 then
      FileSetAttr(sFileName, 0);
    LeFic := UPPERCASE(SearchRec.Name);
    for i := 0 to 6 do
    begin
      if LeFic = TableauPDF[i] then Sysutils.DeleteFile(sFileName);
    end;
    iIndex := FindNext(SearchRec);
  end;
  Sysutils.FindClose(SearchRec);
end;
// FIN PT36

function SurchargeCase(Base: string; T: array of string; V: array of integer): integer;
var
  i: integer;
  FindIt: Boolean;
  S: string;
begin
  FindIt := False;
  for i := 0 to SizeOf(T) do
  begin
    S := ReadTokenSt(T[i]);
    while S <> '' do
    begin
      if S = Base then
      begin
        FindIt := True;
        break;
      end;
      S := ReadTokenSt(T[i]);
    end;
    if FindIt then break;
  end;
  if FindIt then Result := V[i] else Result := -1;
end;

{ Reinitialisation d'un THGRID sans ecraser la valeur courante de la ligne sur
laquelle on est positionné
}

procedure PaieRazPile(Grille: THGrid);
var
  C, R: integer;
begin
  for R := Grille.FixedRows to Grille.RowCount - 1 do
    for C := 0 to Grille.ColCount - 1 do
    begin
      Grille.Cells[C, R] := '';
      Grille.Objects[C, R] := nil;
    end;
end;
{ Rend le Mois et l'annee d'une date sous forme de string
Attention, le format de la date est de type ShortDate cad que l'année
est sur 2 caractères
}

procedure RendMoisAnnee(Zdate: TDateTime; var Mois, Annee: string);
var
  aa, mm, jj: word;
begin
  decodedate(ZDate, aa, mm, jj);
  Annee := inttostr(aa);
  Mois := inttostr(mm);
  {                     modif mv
  Annee:= Copy (DateToStr (Zdate), 7,2);
  if StrToInt (Annee) < 80 then Annee :='20'+Annee
   else Annee :='19'+Annee;
  Mois := Copy (DateToStr (Zdate), 4,2); // @@@@}
end;
{ fonction qui rend la Tob calendrier des jours travaillés
}

function RendCalendrierMois(Etab, Annee, Mois: string; TOB_JoursTrav: TOB): TOB;
var
  T1: TOB;
begin
  T1 := TOB_JoursTrav.FindFirst(['PJT_ANNEE', 'PJT_MOIS'], [Annee, mois], TRUE);
  if etab <> '' then
  begin
    result := TOB_JoursTrav.FindFirst(['PJT_ETABLISSEMENT', 'PJT_ANNEE', 'PJT_MOIS'], [Etab, Annee, mois], TRUE);
    if result = nil then result := T1;
  end
  else result := T1;
end;
{ Fonction qui rend les dates de début et de fin pour sélectionner les lignes à prendre en compte dans les
différents historiques.
Attention : cette fonction exclus la période de paie en cours !!!!!!!
Attention : cette fonction doit prendre en compte les exercices décalés !
2 solutions : Régarder la date de paiement des salaires
              ou saisir dans la table exersocial les dates début et date de fin à prendre en compte
                 en fonction du décalage de paie ==> voir paramsoc
}

function RendDateExerSocial(DateDebut, DateFin: TDateTime; var DatD, DatF: TDateTime; MoisPris: Boolean = FALSE): Boolean;
var
  Q: TQuery;
  st: string;
begin
  result := FALSE;
  st := 'SELECT * FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + UsDateTime(DateDebut) + '" AND PEX_DATEFIN>="' + UsDateTime(DateFin) + '"' +
    ' ORDER BY PEX_DATEDEBUT DESC';
  Q := OpenSQL(st, TRUE);
  if not Q.EOF then
  begin
    DatD := Q.Fields[2].AsFloat; // Recup date de debut exercice social
    result := TRUE;
    // PT1
  end;
  Ferme(Q);
  if MoisPris = FALSE then DatF := DateDebut - 1 // Regul des montants jusqu'à la date (non incluse) de debut de la session
  else DatF := DateFin;
end;
{ Fonction qui rend la + grande date de debut de l'exercice social
Cela devrait correspondre à la date de debut de l'exercice social traité
Cette fonction est publiee dans aglinitpaie pour etre connue du script.
}
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2001
Modifié le ... :   /  /
Description .. : Unit de définition de fonctions génériques de la paie
Mots clefs ... : PAIE
*****************************************************************}

function MaxDebutExerSocial(): variant;
var
  DateD: Double;
  Q: TQuery;
begin
  DateD := 0;
  Q := OPENSQL('SELECT MAX (PEX_DATEDEBUT) FROM EXERSOCIAL', TRUE);
  if not Q.EOF then
  begin
    DateD := Q.Fields[0].AsFloat;
    // PT1 suppression
  end;
  Ferme(Q);
  result := DateToStr(DateD);
end;
{ fonction de recherche de la liste des rubriques de Bases ou autres faisant partie du
bulletin
}
// PT2 21/08/01 V547 PH rajout date de fin
// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}

procedure RechercheRubriqueProfil(Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases: TOB; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
var
  Q: TQuery;
  ThemeProfil: string;
begin
  ThemeProfil := '';
  // Chargement des rubriques de chaque profil reférencé au niveau salarié
  // PT6 02/10/01 V562 PH Prise en compte du profil r2munérations
  // Profil rémunération salarié
  RubriqDuProfil('PSA_TYPPROFILREM', 'PSA_PROFILREM', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Salarié modèle de bulletin
  RubriqDuProfil('PSA_TYPPROFIL', 'PSA_PROFIL', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Periodicite bulletin
  RubriqDuProfil('PSA_TYPPERIODEBUL', 'PSA_PERIODBUL', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Réduction Bas Salaire
  RubriqDuProfil('PSA_TYPPROFILRBS', 'PSA_PROFILRBS', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Réduction REPAS
  RubriqDuProfil('PSA_TYPREDREPAS', 'PSA_REDREPAS', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Réduction RTT 1 Loi Aubry 2
  RubriqDuProfil('PSA_TYPREDRTT1', 'PSA_REDRTT1', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Réduction RTT 2 Loi Robien
  RubriqDuProfil('PSA_TYPREDRTT2', 'PSA_REDRTT2', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Temps Partiel
  RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases, Salarie.GetValue('PSA_PROFILTPS'), BasesSeules, ZDateDebut, ZDateFin);
  // Profil Abattement pour frais Professionnels
  RubriqDuProfil('PSA_TYPPROFILAFP', 'PSA_PROFILAFP', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Gestion des Appoints
  RubriqDuProfil('PSA_TYPPROFILAPP', 'PSA_PROFILAPP', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  //  PT8 02/04/02 V571 PH Traitement du profil retraite
  // PRofil Retraite
  RubriqDuProfil('PSA_TYPPROFILRET', 'PSA_PROFILRET', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // PRofil Mutuelle
  RubriqDuProfil('PSA_TYPPROFILMUT', 'PSA_PROFILMUT', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Prévoyance
  RubriqDuProfil('PSA_TYPPROFILPRE', 'PSA_PROFILPRE', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Taxe sur les Salaires
  RubriqDuProfil('PSA_TYPPROFILTSS', 'PSA_PROFILTSS', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil Congés Payés
  RubriqDuProfil('PSA_TYPPROFILCGE', 'PSA_PROFILCGE', Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases, BasesSeules, ZDateDebut, ZDateFin);
  // Profil fin de CDD
  RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases, Salarie.GetValue('PSA_PROFILCDD'), BasesSeules, ZDateDebut, ZDateFin);
  // Profil Multi Employeur
  RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases, Salarie.GetValue('PSA_PROFILMUL'), BasesSeules, ZDateDebut, ZDateFin);
  // Chargement des rubriques pour les profils CAS Particuliers dans ce cas grand maximum 2 à 3 cas pour un salarié
  Q := OPENSQL('SELECT * FROM PROFILSPECIAUX WHERE PPS_ETABSALARIE="-" AND PPS_CODE="' + Salarie.GetValue('PSA_SALARIE') + '"', TRUE);
  while not Q.EOF do
  begin
    ThemeProfil := Q.FindField('PPS_THEMEPROFIL').AsString;
    RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases, Q.FindField('PPS_PROFIL').AsString, BasesSeules, ZDateDebut, ZDateFin);
    Q.Next;
  end;
  Ferme(Q);
end;
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA


{ Recherche en cascade d'un profil à partir du salarié à l'établissement
La fonction a pour but de rechercher le profil mis par défaut au niveau de l'établissement
dans le cas où le salarié n'est pas personnalisé
Profil FNAL : Idem Dossier donc lecture des paramètres de la paie
}
// PT2 21/08/01 V547 PH rajout date de fin
// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}

procedure RubriqDuProfil(Champ1, Champ2: string; Salarie, T_Etab, TPE, TOB_ProfilP, Tob_LesBases: TOB; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
var
  suffixe, Profil: string;
begin
  suffixe := ExtractSuffixe(Champ2); // recup du suffixe du champ profil à traiter
  Profil := string(Salarie.GetValue(Champ1)); // Type Profil : Idem Etab ou Personnalisé
  {  PT5 07/09/01 V547 PH Test du profil FNAL si Profil '' et non FNAL alors on recherche idem etab
   ceci pour compenser le fait que la zone n'est pas saisie ou mal initialisée ne provoque pas d'erreur mais
   une exception en vision SAV
  }
  if ((Profil = 'ETB') or (Profil = '')) and (suffixe <> 'PROFILFNAL') then Profil := T_Etab.GetValue('ETB_' + suffixe)
  else
  begin // Il n'y a que le profil FNAL qui est paramètrable au niveau du dossier
    if (suffixe = 'PROFILFNAL') and (Profil = 'DOS') then Profil := string(GetParamSoc(suffixe))
    else Profil := string(Salarie.GetValue('PSA_' + suffixe));
  end;
  if Profil <> '' then RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases, Profil, BasesSeules, ZDateDebut, ZDateFin);
end;
// PT2 21/08/01 V547 PH rajout date de fin
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA

// DEBUT CCMX-CEGID ORGANIGRAMME DA
{$IFNDEF GCGC}

procedure RecupRubProfil(Salarie, TPE, TOB_ProfilP, Tob_LesBases: TOB; Profil: string; BasesSeules: Boolean; ZDateDebut, ZDateFin: TDateTime);
var
  TPP, TPR, THB, T, T_Cot, Rubrique: TOB;
  Them, ThemeRub, Rub, ThemeExclus, Themeplus: string;
  I: Integer;
begin
  if Profil = '' then exit;
  ThemeRub := '';
  ThemeExclus := '';
  if ThemeExclus = 'RDC' then Themeplus := 'RDS'
  else if ThemeExclus = 'RDS' then Themeplus := 'RDC'
  else Themeplus := 'XYZ';
  TPP := TOB_ProfilP.FindFirst(['PPI_PROFIL'], [Profil], FALSE);
  if TPP <> nil then
  begin
    ThemeExclus := TPP.GetValue('PPI_THEMECOT'); // theme des rubriques de cotisations à exclure
    if (ThemeExclus <> '') and (BasesSeules = FALSE) then
    begin
      T := TPE.FindFirst(['PHB_NATURERUB'], ['COT'], TRUE);
      while T <> nil do
      begin
        Rub := T.GetValue('PHB_RUBRIQUE');
        T_Cot := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE', 'PCT_NATURERUB'], [Rub, 'COT'], FALSE); // $$$$
        if T_Cot <> nil then // Si Cotisation trouvee
        begin
          Them := T_Cot.GetValue('PCT_THEMECOT');
          if (Them = ThemeExclus) or (Them = Themeplus) then T.Free; // suppression de l'objet dans la TOB si theme identique
        end;
        T := TPE.FindNext(['PHB_NATURERUB'], ['COT'], TRUE);
      end;
    end; // Fin Theme A exclure

    for I := 0 to TPP.Detail.count - 1 do
    begin
      TPR := TPP.Detail[I];
      Rub := TPR.GetValue('PPM_RUBRIQUE');
      Rubrique := Tob_LesBases.FindFirst(['PCT_NATURERUB', 'PCT_RUBRIQUE'], ['BAS', Rub], FALSE);
      if Rubrique <> nil then
      begin
        if TPE.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE'], [TPR.GetValue('PPM_NATURERUB'), TPR.GetValue('PPM_RUBRIQUE')], FALSE) = nil then // $$$$
        begin
          // on va ecrire une ligne bulletin THB comme fille de la TOB TPE
          THB := TOB.Create('HISTOBULLETIN', TPE, -1);
          THB.PutValue('PHB_ETABLISSEMENT', Salarie.GetValue('PSA_ETABLISSEMENT'));
          THB.PutValue('PHB_SALARIE', Salarie.GetValue('PSA_SALARIE'));
          THB.PutValue('PHB_DATEDEBUT', ZDateDebut);
          //  PT2 21/08/01 V547 PH rajout date de fin au lieu de la date de début
          THB.PutValue('PHB_DATEFIN', ZDateFin);
          THB.PutValue('PHB_NATURERUB', Rubrique.GetValue('PCT_NATURERUB'));
          THB.PutValue('PHB_RUBRIQUE', Rubrique.GetValue('PCT_RUBRIQUE'));
          THB.PutValue('PHB_LIBELLE', Rubrique.GetValue('PCT_LIBELLE'));
          THB.PutValue('PHB_IMPRIMABLE', Rubrique.GetValue('PCT_IMPRIMABLE'));
          THB.PutValue('PHB_BASEREM', 0);
          THB.PutValue('PHB_TAUXREM', 0);
          THB.PutValue('PHB_COEFFREM', 0);
          THB.PutValue('PHB_MTREM', 0);
          //  PT13 23/09/02 V585 PH Suppression des lignes en doublons provoque des erreurs en SAV
          {         THB.PutValue('PHB_BASEREMIMPRIM',Rubrique.GetValue('PCT_BASEIMPRIMABLE')) ;
                   THB.PutValue('PHB_TAUXREMIMPRIM',Rubrique.GetValue('PCT_TAUXIMPRIMABLE')) ;
                   THB.PutValue('PHB_COEFFREMIMPRIM',Rubrique.GetValue('PCT_COEFFIMPRIM')) ;
                   THB.PutValue('PHB_ORDREETAT',Rubrique.GetValue('PCT_ORDREETAT')) ;
                   THB.PutValue('PHB_SENSBUL',Rubrique.GetValue('PCT_SENSBUL')) ;  }
          // FIN PT13
          THB.PutValue('PHB_BASECOT', 0);
          THB.PutValue('PHB_TAUXSALARIAL', 0);
          THB.PutValue('PHB_TAUXPATRONAL', 0);
          THB.PutValue('PHB_MTSALARIAL', 0);
          THB.PutValue('PHB_MTPATRONAL', 0);
          THB.PutValue('PHB_BASECOTIMPRIM', Rubrique.GetValue('PCT_BASEIMP'));
          THB.PutValue('PHB_TAUXSALIMPRIM', Rubrique.GetValue('PCT_TXSALIMP'));
          THB.PutValue('PHB_TAUXPATIMPRIM', Rubrique.GetValue('PCT_TXPATIMP'));
          THB.PutValue('PHB_ORDREETAT', 3); { PT34 }
          THB.PutValue('PHB_OMTSALARIAL', Rubrique.GetValue('PCT_ORDREETAT')); { PT34 }
          THB.PutValue('PHB_SENSBUL', 'P');
          THB.PutValue('PHB_TRAVAILN2', Salarie.GetValue('PSA_TRAVAILN2'));
          THB.PutValue('PHB_TRAVAILN3', Salarie.GetValue('PSA_TRAVAILN3'));
          THB.PutValue('PHB_TRAVAILN4', Salarie.GetValue('PSA_TRAVAILN4'));
          THB.PutValue('PHB_TRAVAILN1', Salarie.GetValue('PSA_TRAVAILN1'));
          THB.PutValue('PHB_CODESTAT', Salarie.GetValue('PSA_CODESTAT'));
          THB.PutValue('PHB_CONSERVATION', 'BUL'); // Pour indiquer que la rubrique vient d'un bulletin
        end;
      end; // Si rubrique OK
    end; // Fin du FOR
  end;
end;
// PTXX
{$ENDIF}
// FIN CCMX-CEGID ORGANIGRAMME DA

procedure VisibiliteChampSalarie(Numero: string; Champ1, Champ2: TControl; Champ3: TControl = nil);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Rupt: TRadioButton;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  Rupt := TRadioButton(Champ3);
  if Num > VH_Paie.PGNbreStatOrg then
  begin
    //  PT10 27/06/02 V582 PH tests si composant non nul avant affectation
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    if Rupt <> nil then Rupt.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
    if Rupt <> nil then Rupt.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    if Num = 1 then TLieu.Caption := VH_Paie.PGLibelleOrgStat1
    else if Num = 2 then TLieu.Caption := VH_Paie.PGLibelleOrgStat2
    else if Num = 3 then TLieu.Caption := VH_Paie.PGLibelleOrgStat3
    else TLieu.Caption := VH_Paie.PGLibelleOrgStat4;
  end;
end;

// PT7 19/03/02 V571 PH Rajout fonction affichage des champs libres COMBO salarié
// PTXX

procedure VisibiliteChampLibreSal(Numero: string; Champ1, Champ2: TControl; Champ3: TControl = nil);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Rupt: TRadioButton;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  Rupt := TRadioButton(Champ3);
  if Num > VH_Paie.PgNbCombo then
  begin
    // PT10 27/06/02 V582 PH tests si composant non nul avant affectation
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    if Rupt <> nil then Rupt.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    if Rupt <> nil then Rupt.Visible := TRUE;
    if Num = 1 then TLieu.Caption := VH_Paie.PgLibCombo1
    else if Num = 2 then TLieu.Caption := VH_Paie.PgLibCombo2
    else if Num = 3 then TLieu.Caption := VH_Paie.PgLibCombo3
    else TLieu.Caption := VH_Paie.PgLibCombo4;
  end;
end;
// FIN PT7
// PT11 27/06/02 V582 PH Fonction de visibilité des boites à cocher idem combo

procedure VisibiliteBoolLibreSal(Numero: string; Champ1: TControl);
var
  TLieu: TCheckBox;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  TLieu := TCheckBox(Champ1);
  if Num > VH_Paie.PgNbCoche then
  begin
    if TLieu <> nil then TLieu.Visible := FALSE;
    exit;
  end;
  if TLieu <> nil then TLieu.Visible := TRUE;
  if TLieu <> nil then
  begin
    if Num = 1 then TLieu.Caption := VH_Paie.PgLibCoche1
    else if Num = 2 then TLieu.Caption := VH_Paie.PgLibCoche2
    else if Num = 3 then TLieu.Caption := VH_Paie.PgLibCoche3
    else TLieu.Caption := VH_Paie.PgLibCoche4;
  end;
end;
// PTXX

procedure VisibiliteStat(Champ1, Champ2: TControl; Champ3: TControl = nil);
var
  TStat: THLabel;
  Stat: THValComboBox;
  Rupt: TRadioButton;
begin
  Stat := THValComboBox(Champ1);
  TStat := THLabel(Champ2);
  Rupt := TRadioButton(Champ3);
  if (Stat <> nil) and (TStat <> nil) then
  begin
    Stat.Enabled := FALSE;
    Stat.Visible := FALSE;
    TStat.Caption := '';
    TStat.Visible := FALSE;
    if Rupt <> nil then Rupt.Visible := FALSE;
  end else exit;
  if VH_Paie.PGLibCodeStat <> '' then
  begin
    if Stat <> nil then
    begin
      Stat.Enabled := TRUE;
      Stat.Visible := TRUE;
    end;
    if TStat <> nil then
    begin
      TStat.Caption := VH_Paie.PGLibCodeStat;
      TStat.Visible := TRUE;
    end;
    if Rupt <> nil then Rupt.Visible := TRUE;
  end;
end;
// d PT18 fonction affichage des champs libres salarié  PSA_LIBREx

procedure VisibiliteChampLibreX(Numero: string; Champ1, Champ2: TControl);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);
  if Num > VH_Paie.PgNbSalLib then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
    if Num = 1 then TLieu.Caption := VH_Paie.PgSalLib1
    else if Num = 2 then TLieu.Caption := VH_Paie.PgSalLib2
    else if Num = 3 then TLieu.Caption := VH_Paie.PgSalLib3
    else TLieu.Caption := VH_Paie.PgSalLib4;
  end;
end;
// FIN PT18
// PTxx

procedure VisibiliteChampBoollX(Numero: string; Champ1: TControl);
var
  ChampLieu: TCheckBox;
  Num: Integer;
begin
  Num := StrToInt(Numero);
  ChampLieu := TCheckBox(Champ1);
  if Num > VH_Paie.PgNbCoche then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
    if Num = 1 then ChampLieu.Caption := VH_Paie.PgLibCoche1
    else if Num = 2 then ChampLieu.Caption := VH_Paie.PgLibCoche2
    else if Num = 3 then ChampLieu.Caption := VH_Paie.PgLibCoche3
    else ChampLieu.Caption := VH_Paie.PgLibCoche4;
  end;
end;
// FIN PTxx
// d PT24

procedure VisibiliteChampRupt(Numero: string; Champ1, Champ2: TControl);
var
  TLieu: THLabel;
  ChampLieu: THValComboBox;
  Num: Integer;

begin

  Num := StrToInt(Numero);

  ChampLieu := THValComboBox(Champ1);
  TLieu := THLabel(Champ2);

  if (Num > (VH_Paie.PgNbreStatOrg + VH_Paie.PgNbCombo + 1)) then
  begin
    if ChampLieu <> nil then ChampLieu.Visible := FALSE;
    if TLieu <> nil then TLieu.Visible := FALSE;
    exit;
  end;
  if ChampLieu <> nil then
  begin
    ChampLieu.Enabled := TRUE;
    ChampLieu.Visible := TRUE;
  end;
  if TLieu <> nil then
  begin
    TLieu.Visible := TRUE;
  end;
end;
// f PT24
// Fonction de comparaison de 2 chaines à partir d'une position sur une longueur

function MemCmpPaie(Chaine1, Chaine2: string; Position, Longueur: Integer): Boolean;
begin
  result := FALSE;
  if (Length(Chaine1) < Longueur) or (Length(Chaine2) < Longueur) then exit; // si longueur chaine < longueur des chaines à tester
  if (Length(Chaine1) < Position + Longueur - 1) then exit;
  if (Length(Chaine2) < Position + Longueur - 1) then exit;
  if (Copy(Chaine1, Position, Longueur) = Copy(Chaine2, Position, Longueur)) then result := TRUE;
end;

{  PT3 07/09/01 V547 PH Controle des dates de la paie par rapport aux clotures
On recherche en premier l'exercice sur lequel on travaille,
si non trouvé alerte et on continue quand même pour faire des paies sur 2 exercices
Rend TRUE si le controle est bon.
}

function ControlPaieCloture(DateDebut, DateFin: TDateTime): boolean;
var
  st, Cloture: string;
  Q: TQuery;
  AA, MD, MF, JJ: WORD;
  okok, Indice, i: Integer;
begin
  result := FALSE;
  st := 'SELECT PEX_CLOTURE FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + UsDateTime(DateDebut) + '" AND PEX_DATEFIN>="' + UsDateTime(DateFin) + '"' +
    ' ORDER BY PEX_DATEDEBUT DESC';
  Q := OpenSQL(st, TRUE);
  if not Q.EOF then Cloture := Q.FindField('PEX_CLOTURE').AsString
  else
  begin
    Ferme(Q);
    OkOk := PGIAsk('Attention, les dates du bulletins ne sont pas incluses dans un exercice social#13#10 Voulez vous continuer', 'Contrôle des dates du bulletin');
    if OkOk = 6 then result := TRUE;
    exit;
  end;
  Ferme(Q);
  // recuperation des mois
  DecodeDate(DateDebut, AA, MD, JJ);
  DecodeDate(DateFin, AA, MF, JJ);
  Indice := MD;
  if VH_Paie.PGDecalage = TRUE then
  begin
    if Indice = 1 then Indice := 2 else
    begin
      if Indice = 12 then Indice := 1 else
        Indice := Indice + 1;
    end;
  end;
  if MD = MF then // On a une paie sur 1 mois
  begin
    if Cloture[Indice] = '-' then result := TRUE; // OK mois non cloturé
  end
  else
  begin // on a une paie sur plusieurs mois
    result := TRUE;
    for i := MD to MF do
    begin
      if VH_Paie.PGDecalage = TRUE then
      begin
        if i = 1 then Indice := 12 else
        begin
          if I = 12 then Indice := 1 else
            Indice := i + 1;
        end;
      end else Indice := i;
      if Cloture[Indice] = 'X' then
      begin
        Result := FALSE;
        break;
      end; // Si on trouve un seul mois cloturé alors erreur
    end;
  end;
end;

{ PT4 07/09/01 V547 PH Fonction RendMois mise dans P5DEF existait dans UTofPG_CLOTURE
Elle retourne l'indice correspond au test du champ PEX_CLOTURE pour savoir si un mois
est cloturé
}

function RendMois(Indice: WORD): WORD;
begin
  result := Indice;
  if VH_Paie.PGDecalage = TRUE then
  begin
    if Indice = 1 then result := 12 else
      result := Indice - 1;
  end;
end;
{ Fonction de creation enregistrement journal des évènements de lan paie
Prise en charge complète du traitement. attention, tous les arguments doivent
être renseignés sauf le statut.
}
// PT9 03/06/02 V582 PH Gestion historiques des évènements

function CreeJnalEvt(TypeEvt, Evt, Statut: string; Lbox, LboxS: TlistBox; Trace: TStringList = nil; TraceE: TStringList = nil): Boolean;
var
  TT: TOB;
  QQ: TQuery;
  NumOrdre: Integer;
  DS, D: Boolean;
begin
  result := FALSE;
  if ((LBox = nil) and (Trace = nil)) or (TypeEvt = '') or (Evt = '') then exit;
  NumOrdre := 0;
  TT := TOB.Create('PAIEJEVT', nil, -1);
  TT.PutValue('PJE_FAMEVENT', TypeEvt);
  TT.PutValue('PJE_EVENT', Evt);
  TT.PutValue('PJE_LIBELLE', RechDom('PGEVT', Evt, FALSE));
  TT.PutValue('PJE_DATEEVENT', Now); //PT46 Now remplace Date
  TT.PutValue('PJE_UTILISATEUR', V_PGI.User);
  DS := V_PGI.DebugSQL;
  D := V_PGI.Debug;
  V_PGI.DebugSQL := FALSE;
  V_PGI.Debug := FALSE;
  QQ := OpenSQL('SELECT MAX(PJE_NUMORDRE) FROM PAIEJEVT', True);
  if not QQ.EOF then NumOrdre := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumOrdre);
  TT.PutValue('PJE_NUMORDRE', Numordre);
  if Trace = nil then
  begin
    if (LBoxS <> nil) then TT.PutValue('PJE_BLOCNOTE', LBox.items.Text +
        'Anomalies identifiées :' + #13 + #10 + LBoxS.items.Text)
    else TT.PutValue('PJE_BLOCNOTE', LBox.items.Text);
  end
  else
  begin
    if (LBox <> nil) then TT.PutValue('PJE_BLOCNOTE', LBox.items.Text + ' ' + Trace.Text)
    else
    begin
      if TraceE <> nil then TT.PutValue('PJE_BLOCNOTE', Trace.Text + ' ' + TraceE.Text)
      else TT.PutValue('PJE_BLOCNOTE', Trace.Text);
    end;
  end;
  if Statut = '' then TT.PutValue('PJE_ETATEVENT', 'OK')
  else TT.PutValue('PJE_ETATEVENT', Statut);
  try
    TT.InsertDB(nil);
    result := TRUE;
  except
    result := FALSE;
  end;
  V_PGI.DebugSQL := DS;
  V_PGI.Debug := D;
  TT.Free;
end;
// FIN PT9
// PT12 02/07/02 V582 PH Fonction initialisation de tous les champs TOB Salaries

procedure InitSalDef(TS, TEtab: TOB; Etab: string);
var
  Te: TOB;
begin
  TS.PutValue('PSA_TYPPROFIL', 'ETB');
  TS.PutValue('PSA_TYPPROFILREM', 'ETB');
  TS.PutValue('PSA_TYPPERIODEBUL', 'ETB');
  TS.PutValue('PSA_TYPPROFILRBS', 'ETB');
  TS.PutValue('PSA_TYPPROFILAFP', 'ETB');
  TS.PutValue('PSA_TYPPROFILAPP', 'ETB');
  TS.PutValue('PSA_TYPPROFILRET', 'ETB');
  TS.PutValue('PSA_TYPPROFILMUT', 'ETB');
  TS.PutValue('PSA_TYPPROFILPRE', 'ETB');
  TS.PutValue('PSA_TYPPROFILTSS', 'ETB');
  TS.PutValue('PSA_TYPPROFILCGE', 'ETB');
  TS.PutValue('PSA_TYPPROFILANC', 'ETB');
  TS.PutValue('PSA_CPACQUISMOIS', 'ETB');
  TS.PutValue('PSA_CPTYPEMETHOD', 'ETB');
  TS.PutValue('PSA_CPTYPERELIQ', 'ETB');
  TS.PutValue('PSA_CPACQUISANC', 'ETB');
  TS.PutValue('PSA_DATANC', 'ETB');
  TS.PutValue('PSA_TYPPROFILTRANS', 'ETB');
  TS.PutValue('PSA_TYPPROFILFNAL', 'DOS');
  TS.PutValue('PSA_STANDCALEND', 'ETB');
  TS.PutValue('PSA_TYPREDREPAS', 'ETB');
  TS.PutValue('PSA_TYPREDRTT1', 'ETB');
  TS.PutValue('PSA_TYPREDRTT2', 'ETB');
  TS.PutValue('PSA_TYPDADSFRAC', 'ETB');
  TS.PutValue('PSA_TYPPRUDH', 'ETB');
  TS.PutValue('PSA_TYPREGLT', 'ETB');
  TS.PutValue('PSA_TYPVIRSOC', 'ETB');
  TS.PutValue('PSA_TYPDATPAIEMENT', 'ETB');
  TS.PutValue('PSA_TYPPAIACOMPT', 'ETB');
  TS.PutValue('PSA_TYPACPSOC', 'ETB');
  TS.PutValue('PSA_TYPPAIFRAIS', 'ETB');
  TS.PutValue('PSA_TYPFRAISSOC', 'ETB');
  TS.PutValue('PSA_TYPJOURHEURE', 'ETB');
  TS.PutValue('PSA_DATEENTREE', Date);
  TS.PutValue('PSA_DATEANCIENNETE', Idate1900);
  // PT14  : 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 au lieu de null
  TS.PutValue('PSA_DATENAISSANCE', Idate1900);
  TS.PutValue('PSA_DATESORTIE', Idate1900);
  TS.PutValue('PSA_DATELIBRE1', Idate1900);
  TS.PutValue('PSA_DATELIBRE2', Idate1900);
  TS.PutValue('PSA_DATELIBRE3', Idate1900);
  TS.PutValue('PSA_DATELIBRE4', Idate1900);
  TS.PutValue('PSA_DATENTREECRC', Idate1900);
  TS.PutValue('PSA_DATSORTIECRC', Idate1900);
  TS.PutValue('PSA_ORDREAT', 1);
  TS.PutValue('PSA_DADSFRACTION', 1);
  TS.PutValue('PSA_REGIMESS', '200');
  //DEB PT51
  TS.PutValue('PSA_TYPEREGIME', '-');
  TS.PutValue('PSA_REGIMEMAL', '');
  TS.PutValue('PSA_REGIMEAT', '');
  TS.PutValue('PSA_REGIMEVIP', '');
  TS.PutValue('PSA_REGIMEVIS', '');
  //FIN PT51
  TS.PutValue('PSA_TAUXPARTSS', 0);
  TS.PutValue('PSA_TAUXPARTIEL', 0);
  TS.PutValue('PSA_DADSDATE', IDate1900);
  TS.PutValue('PSA_PRISEFFECTIF', 'X');
  TS.PutValue('PSA_UNITEPRISEFF', 1);
  TS.PutValue('PSA_TYPJOURHEURE', 'DOS');
  TS.PutValue('PSA_TYPPROFILFNAL', 'DOS');
  //PT15  : 03/01/2003 PH V591 FQ10417 initialisation nationalités
  TS.PutValue('PSA_NATIONALITE', 'FRA');
  TS.PutValue('PSA_PAYSNAISSANCE', 'FRA');

  if Etab = '' then exit;
  Te := TEtab.FindFirst(['ETB_ETABLISSEMENT'], [Etab], FALSE);
  if Te <> nil then
  begin
    TS.PutValue('PSA_CONVENTION', Te.getValue('ETB_CONVENTION'));
    TS.PutValue('PSA_PROFILANCIEN', Te.GetValue('ETB_PROFILANCIEN'));
    TS.PutValue('PSA_PROFIL', Te.GetValue('ETB_PROFIL'));
    TS.PutValue('PSA_PROFILREM', Te.GetValue('ETB_PROFILREM'));
    TS.PutValue('PSA_PERIODBUL', Te.GetValue('ETB_PERIODBUL'));
    TS.PutValue('PSA_PROFILRBS', Te.GetValue('ETB_PROFILRBS'));
    TS.PutValue('PSA_REDREPAS', Te.GetValue('ETB_REDREPAS'));
    TS.PutValue('PSA_REDRTT1', Te.GetValue('ETB_REDRTT1'));
    TS.PutValue('PSA_REDRTT2', Te.GetValue('ETB_REDRTT2'));
    TS.PutValue('PSA_PROFILAFP', Te.GetValue('ETB_PROFILAFP'));
    TS.PutValue('PSA_PCTFRAISPROF', Te.GetValue('ETB_PCTFRAISPROF'));
    TS.PutValue('PSA_PROFILAPP', Te.GetValue('ETB_PROFILAPP'));
    TS.PutValue('PSA_PROFILRET', Te.GetValue('ETB_PROFILRET'));
    TS.PutValue('PSA_PROFILMUT', Te.GetValue('ETB_PROFILMUT'));
    TS.PutValue('PSA_PROFILPRE', Te.GetValue('ETB_PROFILPRE'));
    TS.PutValue('PSA_PROFILTSS', Te.GetValue('ETB_PROFILTSS'));
    TS.PutValue('PSA_PROFILTRANS', Te.GetValue('ETB_PROFILTRANS'));
    TS.PutValue('PSA_CALENDRIER', Te.GetValue('ETB_STANDCALEND'));
    TS.PutValue('PSA_PROFILFNAL', VH_Paie.PGProfilFnal);
    TS.PutValue('PSA_PROFILCGE', Te.GetValue('ETB_PROFILCGE'));
    TS.PutValue('PSA_NBREACQUISCP', Te.GetValue('ETB_NBREACQUISCP'));
    TS.PutValue('PSA_VALORINDEMCP', Te.GetValue('ETB_VALORINDEMCP'));
    TS.PutValue('PSA_RELIQUAT', Te.GetValue('ETB_RELIQUAT'));
    TS.PutValue('PSA_BASANCCP', Te.GetValue('ETB_BASANCCP'));
    TS.PutValue('PSA_TYPDATANC', Te.GetValue('ETB_TYPDATANC'));
    TS.PutValue('PSA_JOURHEURE', Te.GetValue('ETB_JOURHEURE'));
    TS.PutValue('PSA_JOURHEURE', VH_PAIE.PGJourHeure);
    TS.PutValue('PSA_DADSFRACTION', Te.GetValue('ETB_CODESECTION'));
    TS.PutValue('PSA_PRUDHCOLL', Te.GetValue('ETB_PRUDHCOLL'));
    TS.PutValue('PSA_PRUDHSECT', Te.GetValue('ETB_PRUDHSECT'));
    TS.PutValue('PSA_PRUDHVOTE', Te.GetValue('ETB_PRUDHVOTE'));
    TS.PutValue('PSA_PGMODEREGLE', Te.GetValue('ETB_PGMODEREGLE'));
    TS.PutValue('PSA_RIBVIRSOC', Te.GetValue('ETB_RIBSALAIRE'));
    TS.PutValue('PSA_PAIACOMPTE', Te.GetValue('ETB_PAIACOMPTE'));
    TS.PutValue('PSA_RIBACPSOC', Te.GetValue('ETB_RIBACOMPTE'));
    TS.PutValue('PSA_PAIFRAIS', Te.GetValue('ETB_PAIFRAIS'));
    TS.PutValue('PSA_RIBFRAISSOC', Te.GetValue('ETB_RIBFRAIS'));
    TS.PutValue('PSA_MOISPAIEMENT', Te.GetValue('ETB_MOISPAIEMENT'));
    TS.PutValue('PSA_JOURPAIEMENT', Te.GetValue('ETB_JOURPAIEMENT'));
  end;
end;


// PT19  : 17/06/03 V_421 PH Fonction qui rend le mode de fonctionnement MONO ou MULTI

function PGRendModeFonc(): string;
var
  st: string;
  Q: TQuery;
begin
  //  St := GetParamSoc('SO_MODEFONC');
{$IFDEF EAGLSERVER}
  // PROCESS-SERVEUR
  result := 'MONO';
{$ELSE}
  if LeModeT = '' then
  begin
    Q := OpenSQL('SELECT SO_MODEFONC FROM SOCIETE', True);
    if not Q.EOF then st := Q.FindField('SO_MODEFONC').AsString;
    Ferme(Q);
    if (St = 'DB0') or (V_PGI.ModePcl = '1') then LeModeT := 'MULTI'
    else LeModeT := 'MONO';
  end;
  result := LeModeT;
{$ENDIF}
end;

// PT17  : 02/04/03 V_421 PH Fonction de suppression en masse des bulletins de paie
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 17/06/2004
Modifié le ... :   /  /
Description .. : ATTENTION !!
Suite ........ : A faire
Suite ........ : Vérif des diff entre OnDelete(SaisBul) et SuppressionBull
Suite ........ : (P5DEF)
Mots clefs ... :
*****************************************************************}
// DEBUT CCMX-CEGID DA VERSION6
{$IFNDEF GCGC}

procedure SuppressionBull(Zdate1, Zdate2: TDateTime; CodeSalarie, Etab, BullCompl: string; TheCP: BooLean);
var
  St, RefA: string;
  Nbre: longint;
  Trace: TStringList;
begin
  BeginTrans;
  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  Trace := TStringList.Create;
  try
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
    ExecuteSQL('UPDATE HISTOSAISRUB SET PSD_DATEINTEGRAT ="' + usdatetime(0) +
      ' " WHERE PSD_SALARIE = "' + CodeSalarie + '"' +
      ' AND (PSD_DATEINTEGRAT >="' + usdatetime(ZDate1) +
      '" AND PSD_DATEINTEGRAT <="' + usdatetime(ZDate2) + '")');
    // PT22  : 23/09/03 V_421 PH Rajout suppression saisie arret dans la suppression en masse des bulletins
    PGSuppHistoRetenue(CodeSalarie, ZDate1, ZDate2);

    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    if BullCompl <> 'X' then AnnuleAbsenceBulletin(CodeSalarie, ZDate2);

    // d PT26 IJSS - Raz date intégration dans REGLTIJSS
{$IFNDEF CCS3}
    if (BullCompl <> 'X') and (VH_Paie.PGGestIJSS) then
      AnnuleIJSSBulletin(CodeSalarie, ZDate2);
{$ENDIF}
    // f PT26 IJSS

    // d PT28 Maintien - Raz date de paie dans MAINTIEN
{$IFNDEF CCS3}
    if (BullCompl <> 'X') and (VH_Paie.PGMaintien) then
      AnnuleMaintienBulletin(CodeSalarie, ZDate1, ZDate2);
{$ENDIF}
    // f PT28 Maintien

    if Nbre < 0 then Showmessage('Une erreur s''est produite pendant l''annulation');
    //PT9 : 22/10/2001 V562 PH Gestion cas particulier du bulletin complémentaire et Dates edition
    if (BullCompl <> 'X') and ((VH_paie.PGCongesPayes) and (TheCP)) then
    begin
      AnnuleCongesPris(CodeSalarie, Etab, ZDate1, ZDate2); // reaffectation des cp pris et acquis
      AnnuleCongesAcquis(CodeSalarie, ZDate1, ZDate2);
    end;
    // Suppression historique analytique
    if VH_Paie.PGAnalytique = TRUE then
    begin
      st := 'DELETE FROM HISTOANALPAIE WHERE PHA_SALARIE="' + CodeSalarie + '" AND PHA_DATEDEBUT="' + USDateTime(ZDate1) + '" AND PHA_DATEFIN="' + USDateTime(ZDate2) + '"';
      ExecuteSQL(St);
      RefA := CodeSalarie + ';' // Code Salarie
        + FormatDateTime('ddmmyyyy', ZDate1) + ';' // Date Debut bulletin
        + FormatDateTime('ddmmyyyy', ZDate2) + ';'; // Date Fin bulletin
      st := 'DELETE FROM VENTANA WHERE YVA_TABLEANA="PPU" AND YVA_IDENTIFIANT like "' + RefA + '%"'; // suppression des ventilations analytiques du bulletin
      ExecuteSQL(St);
    end;
    // Destruction enregistrement historique salarié
    // PT24 : 07/02/2002 V571 PH mise en place au niveau Historisation salarié
    if VH_Paie.PgHistorisation then
      ExecuteSQL('DELETE FROM HISTOSALARIE WHERE PHS_SALARIE="' + CodeSalarie + '" AND PHS_DATEEVENEMENT="' + USDateTime(ZDate1) + '" AND PHS_DATEAPPLIC="' + USDateTime(ZDate2) +
        '"');
    // FIN PT24
    // Annulation de l'intégration des rubriques de réguls des cotisations
    AnnulRegulCot(CodeSalarie, ZDate1, Zdate2); // PT45

    // PT57  Mise à jour table PRESENCESALARIE ; Le champ PGINDICATPRES = "AIP" au lieu de "INP"
    if VH_Paie.PGMODULEPRESENCE then
    MiseAjourPresence(Zdate1, Zdate2, Codesalarie, 'SUP');

    CommitTrans;
    // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
    Trace.add('Suppression bulletin du ' + DateToStr(ZDate1) + ' au ' + DateToStr(Zdate2) + ' pour le salarié ' + CodeSalarie);
    CreeJnalEvt('001', '002', 'OK', nil, nil, Trace);
  except
    Rollback;
    PGIBox('Une erreur est survenue lors de la suppression du bulletin', 'Saisie du bulletin');
    // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
    Trace.add('Erreur de suppression bulletin du ' + DateToStr(ZDate1) + ' au ' + DateToStr(Zdate2) + ' pour le salarié ' + CodeSalarie);
    CreeJnalEvt('001', '002', 'ERR', nil, nil, Trace);
  end;
  // PT39 : 03/06/2002 V582 PH Gestion historique des évènements
  Trace.Free;
end;
//PT21  : 08/09/03 V_421 PH Procedure de chargement de la table de correspondance PCS
{$ENDIF}
// FIN CCMX-CEGID DA VERSION6
{***********A.G.L.***********************************************
Auteur  ...... : Dev Paie
Créé le ...... : 08/09/2003
Modifié le ... :   /  /
Description .. : Fonction de chargement de la table de correspondance
Suite ........ : des Codes PCS .
Suite ........ : Utilise la fonction d'import de bob de GalOutil
Suite ........ : Attention, GalOutil est non compatible CWAS
Mots clefs ... : PAIE;PCS
*****************************************************************}

procedure ChargePCS();
{$IFNDEF EAGLCLIENT}
var
  Q: TQuery;
  ii: Integer;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  Q := OpenSql('SELECT COUNT (*) NBRE FROM PCSESE', true);
  if not Q.eof then ii := Q.findfield('NBRE').AsInteger;
  Ferme(Q);
{$ENDIF}
end;

{$IFNDEF GCGC}   //PT56
{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Fonction de suppression de l'historique d'une retenue sur
Suite ........ : salaire pour un salarié et une période de paie
Mots clefs ... :
*****************************************************************}

function PGSuppHistoRetenue(Salarie: string; DateDebPaie, DateFinPaie: TDateTime): Boolean;
var
  THisto: Tob;
  Q: TQuery;
  i, NumOrdre: Integer;
  Salarieretenue, Where, st: string;
  Echeancier, Decloture: Boolean;  //PT55
begin
  if Salarie = '' then Where := ''
  else Where := 'PHR_SALARIE="' + Salarie + '" AND';
  st := 'SELECT * FROM HISTORETENUE WHERE ' + Where + ' PHR_DATEDEBUT="' + UsDateTime(DateDebPaie) + '" ' +
    'AND PHR_DATEFIN="' + UsDateTime(DateFinPaie) + '"';
  THisto := Tob.Create('HISTORETENUE', nil, -1);
{Flux optimisé
  Q := OpenSQL('SELECT * FROM HISTORETENUE WHERE ' + Where + ' PHR_DATEDEBUT="' + UsDateTime(DateDebPaie) + '" ' +
    'AND PHR_DATEFIN="' + UsDateTime(DateFinPaie) + '"', True);
  THisto.LoadDetailDB('HISTORETENUE', '', '', Q, False);
  Ferme(Q);
}
  //PT55
  Decloture := False;
  THisto.LoadDetailDBFROMSQL('HISTORETENUE', st);
  for i := 0 to THisto.Detail.Count - 1 do
  begin
    Salarieretenue := THisto.Detail[i].Getvalue('PHR_SALARIE');
    NumOrdre := THisto.Detail[i].Getvalue('PHR_ORDRE');
    Q := OpenSQL('SELECT PRE_ECHEANCIER, PRE_CLOTURE, PRE_REMBEFFECTUE FROM RETENUESALAIRE WHERE PRE_SALARIE="' + Salarieretenue + '" AND PRE_ORDRE=' + IntToStr(NumOrdre) + '', true);
    if not Q.Eof then
    begin
      if Q.FindField('PRE_ECHEANCIER').AsString = 'X' then
        Echeancier := True
      else
        Echeancier := False;
      //PT55
      if (Q.FindField('PRE_CLOTURE').AsString = 'X') or (Q.FindField('PRE_REMBEFFECTUE').AsString = 'X') then
        Decloture := True;
    end
    else Echeancier := False;
    Ferme(Q);
    if Echeancier = True then
    begin
      THisto.Detail[i].PutValue('PHR_MONTANT', 0);
      THisto.Detail[i].PutValue('PHR_EFFECTUE', '-'); //PT32
      THisto.Detail[i].PutValue('PHR_DATEPAIEMENT', IDate1900);
      THisto.Detail[i].PutValue('PHR_CUMULARRIERE', 0);
      THisto.Detail[i].PutValue('PHR_CUMULVERSE', 0);
      THisto.Detail[i].PutValue('PHR_REPRISEARR', 0);
      THisto.Detail[i].PutValue('PHR_ARRIERE', 0);
      THisto.Detail[i].UpdateDB(False);
    end
    else THisto.Detail[i].DeleteDB;
    //PT55
    if Decloture then
    begin
      ExecuteSQL('UPDATE RETENUESALAIRE SET PRE_CLOTURE="-",PRE_REMBEFFECTUE="-" ' +
                 'WHERE PRE_SALARIE="' + Salarie + '" AND PRE_ORDRE=' + IntToStr(NumOrdre) + '');
    end;
  end;
  THisto.Free;
  Result := True;
end;
{$ENDIF}


{$IFNDEF GCGC}   //PT56
{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 01/10/2007
Modifié le ... :   /  /
Description .. : Fonction qui renvoie une tob contenant la liste des saisies
Suite ........ : arrêt du salarié
Mots clefs ... :
*****************************************************************}
function GetTobSaisieArret(Salarie : String; DateFinPaie : TDateTime) : Tob;
begin
  // Récupération des retenues en cours pour le salarié à la période concernée
  result := Tob.Create('Les retenues', nil, -1);
//  result.LoadDetailFromSQL('SELECT PRE_ECHEANCIER,PRE_LIBELLE,PRE_SALARIE,PRE_ORDRE,PRE_RETENUESAL,PRE_MONTANTMENS,PRE_MONTANTTOT,PRE_DATEDEBUT,PRE_DATEFIN,PRE_REMBMAX,' +
//    'PTR_RUBRIQUE,PTR_NIVEAURET,PTR_TYPEFRACTION,PTR_TYPALIMPAIE,PTR_LIBELLE ' +
//    'FROM RETENUESALAIRE ' +
//    'LEFT JOIN TYPERETENUE ON PRE_RETENUESAL=PTR_RETENUESAL ' +
//    'WHERE PTR_NIVEAURET>0 AND PRE_ACTIF="X" AND ' +  //PT35 //PT48
//    '((PRE_REMBMAX="X") ' +  //PT48
//    ' OR (PRE_REMBMAX<>"X" AND PRE_DATEFIN>="' + UsdateTime(DateFinPaie) + '" AND PRE_REMBEFFECTUE<>"X") ' +  //PT23  //PT48
//    ' OR (PRE_REMBMAX<>"X" AND PRE_REMBEFFECTUE<>"X")) ' + //PT23
//    'AND PRE_SALARIE="' + Salarie + '" AND ' +
//    'PRE_DATEDEBUT<="' + UsDateTime(DateFinPaie) + '" ' +
//    'ORDER BY PTR_NIVEAURET,PRE_DATEDEBUT');
  { On ramène les saisies commencées et non cloturées, et les rubriques de la
  période en cours de calcul (pour qu'elles soient recalculées à l'identique
  quand on ouvre un bulletin non clôturé) }
  result.LoadDetailFromSQL('SELECT PRE_ECHEANCIER, PRE_LIBELLE, PRE_SALARIE, '
                          +' PRE_ORDRE, PRE_RETENUESAL, PRE_MONTANTMENS, PRE_MONTANTTOT, '
                          +' PRE_DATEDEBUT, PRE_DATEFIN, PRE_REMBMAX, PTR_RUBRIQUE, '
                          +' PTR_NIVEAURET, PTR_TYPEFRACTION, PTR_TYPALIMPAIE, PTR_LIBELLE '
                          +' FROM RETENUESALAIRE '
                          +' LEFT JOIN TYPERETENUE ON PRE_RETENUESAL=PTR_RETENUESAL '
                          +' WHERE PTR_NIVEAURET > 0 '
                          +'   AND PRE_ACTIF="X" '
                          +'   AND PRE_SALARIE="' + Salarie + '" '
                          +'  AND PRE_DATEDEBUT<="' + UsDateTime(DateFinPaie) + '" '
                          +'  AND (     (PRE_REMBEFFECTUE <> "X") '
                          +'         OR EXISTS(SELECT 1 FROM HISTORETENUE '
                          +'                    WHERE (PHR_SALARIE=PRE_SALARIE) '
                          +'                      AND (PHR_ORDRE=PRE_ORDRE) '
                          +'                      AND (PHR_DATEFIN="' + UsDateTime(DateFinPaie) + '" ) '
                          +'                  ) '
                          +'      ) '
                          +'ORDER BY PTR_NIVEAURET,PRE_DATEDEBUT ');
end;
{$ENDIF}


{$IFNDEF GCGC}   //PT56
{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Cette fonction renvoi une tob contenant toute les retenues
Suite ........ : pouvant être imputé au salaire d'un salarié pour une période
Suite ........ : de paie en fonction des limtations des différents type de
Suite ........ : retenue
Mots clefs ... :
*****************************************************************}

function PGCalcSaisieArret(Salarie: string; DateDebPaie, DateFinPaie: TDateTime; NbACharge: Integer; SalaireNet, Acompte: Double): Tob;
var
  TobRetenueSal: Tob;
  Q: TQuery;
  TotalRetenue, PartiePensionEtat, PartieAutreCreancier, RMI, MtRetenue, MtIntermediaire, PartiePret: Double;
  MtMens, MtArriere, MtReprise, CumulVerse, CumulArriere, SalaireNetAC, CumulTheo, MtTot, CumulRet: Double;
  i, NumOrdre: Integer;
  TypeSaisie, TypeFraction, CodeElement: string;
  RembTotalite, AvecEcheancier: Boolean;
  CorrectionNbCharge: Integer;
begin
  if (NbACharge = 90) or (NbACharge = 99) then CorrectionNbCharge := 0 //PT52
  else CorrectionNbCharge := NbACharge;
  Result := nil;
  SalaireNetAC := SalaireNet + Acompte; //PT27
  // Récupération des retenues en cours pour le salarié à la période concernée
//  Q := OpenSQL('SELECT PRE_ECHEANCIER,PRE_LIBELLE,PRE_SALARIE,PRE_ORDRE,PRE_RETENUESAL,PRE_MONTANTMENS,PRE_MONTANTTOT,PRE_DATEDEBUT,PRE_DATEFIN,PRE_REMBMAX,' +
//    'PTR_RUBRIQUE,PTR_NIVEAURET,PTR_TYPEFRACTION,PTR_TYPALIMPAIE,PTR_LIBELLE ' +
//    'FROM RETENUESALAIRE ' +
//    'LEFT JOIN TYPERETENUE ON PRE_RETENUESAL=PTR_RETENUESAL ' +
//    'WHERE PTR_NIVEAURET>0 AND PRE_ACTIF="X" AND ' + //PT35 //PT48
//    '((PRE_REMBMAX="X") ' + //PT48
//    ' OR (PRE_REMBMAX<>"X" AND PRE_DATEFIN>="' + UsdateTime(DateFinPaie) + '" AND PRE_REMBEFFECTUE<>"X") ' + //PT23  //PT48
//    ' OR (PRE_REMBMAX<>"X" AND PRE_REMBEFFECTUE<>"X")) ' + //PT23
//    'AND PRE_SALARIE="' + Salarie + '" AND ' +
//    'PRE_DATEDEBUT<="' + UsDateTime(DateFinPaie) + '" ' +
//    'ORDER BY PTR_NIVEAURET,PRE_DATEDEBUT', True);
// { Clause WHERE suite PT35 et avant PT48
//    'WHERE PTR_NIVEAURET>0 AND PRE_ACTIF="X" AND PRE_REMBEFFECTUE<>"X" AND ' + //PT35
//    '((PRE_REMBMAX="X") OR (PRE_REMBMAX<>"X" AND PRE_DATEFIN>="' + UsdateTime(DateFinPaie) + '")) ' + //PT23
//    'AND PRE_SALARIE="' + Salarie + '" AND ' +
//    'PRE_DATEDEBUT<="' + UsDateTime(DateFinPaie) + '" ' +
// }
//  TobRetenueSal := Tob.Create('Les retenues', nil, -1);
//  TobRetenueSal.LoadDetailDB('les retenues', '', '', Q, False);
//  Ferme(Q);
  TobRetenueSal := GetTobSaisieArret(Salarie, DateFinPaie);
  if (not Assigned(TobRetenueSal)) or (TobRetenueSal.Detail.Count = 0) then
  begin
//    TobRetenueSal.Free;
    FreeAndNil(TobRetenueSal);
    exit;
  end;
  // Récupération valeur RMI dans éléments nationaux
  //************** Fct ValEltNat dans P5 Util
  CodeElement := '';
  Q := OpenSQL('SELECT PTR_CALCFRACTION,PCF_CODEELT FROM TYPERETENUE ' +
    'LEFT JOIN CALCFRACTION ON PTR_CALCFRACTION=PCF_CALCFRACTION ' +
    'WHERE PTR_RETENUESAL="003"', True);
  if not Q.eof then
  begin
    CodeElement := Q.FindField('PCF_CODEELT').AsString;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT PEL_MONTANTEURO FROM ELTNATIONAUX WHERE PEL_CODEELT="' + CodeElement + '" AND ' +
    'PEL_PREDEFINI="CEG" AND PEL_DATEVALIDITE<="' + UsDateTime(Date) + '" ORDER BY PEL_DATEVALIDITE DESC', true);
  if not Q.Eof then RMI := Q.Fields[0].AsFloat
  else RMI := 0;
  Ferme(Q);
  if SalaireNetAC <= RMI then
  begin    //Debut PT58
    FreeAndNil(TobRetenueSal);
    exit;
  end;
//   Exit; //Fin PT58
  {Découpage du salaire en 2 parties}
  PartieAutreCreancier := PGCalcPartieSaisissable(DateDebPaie, CorrectionNbCharge, SalaireNetAC, RMI, TypeFraction);
  PartiePensionEtat := SalaireNetAC - RMI - PartieAutreCreancier;
  PartiePret := (SalaireNetAC * 10) / 100; //PT31
  CumulRet := 0;
  for i := 0 to TobRetenueSal.Detail.Count - 1 do
  begin
    TotalRetenue := 0;
    MtTot := 0;
    if TobRetenueSal.Detail[i].Getvalue('PRE_REMBMAX') = 'X' then RembTotalite := True
    else RembTotalite := False;
    if TobRetenueSal.Detail[i].Getvalue('PRE_ECHEANCIER') = 'X' then AvecEcheancier := True
    else AvecEcheancier := False;
    TypeSaisie := TobRetenueSal.Detail[i].Getvalue('PTR_TYPEFRACTION');
    NumOrdre := TobRetenueSal.Detail[i].Getvalue('PRE_ORDRE');
    if RembTotalite = False then
    begin
      if AvecEcheancier = True then
      begin
        MtMens := 0;
        Q := OpenSQL('SELECT PHR_MONTANTMENS FROM HISTORETENUE WHERE ' +
          'PHR_SALARIE="' + Salarie + '" AND PHR_ORDRE=' + IntToStr(NumOrdre) + ' AND ' + // DB2
          'PHR_DATEDEBUT="' + UsDateTime(DateDebPaie) + '" AND PHR_DATEFIN="' + UsDateTime(DateFinPaie) + '"', True);
        if not Q.Eof then MtMens := Q.FindField('PHR_MONTANTMENS').AsFloat;
        TobRetenueSal.Detail[i].PutValue('PRE_MONTANTMENS', MtMens);
        Ferme(Q);
        MtTot := TobRetenueSal.Detail[i].Getvalue('PRE_MONTANTTOT');   //PT60 Le montant total n'était pas calculé alors qu'on s'en sert en cas de solde de tous comptes 
      end
      else
      begin
        MtMens := TobRetenueSal.Detail[i].Getvalue('PRE_MONTANTMENS');
        MtTot := TobRetenueSal.Detail[i].Getvalue('PRE_MONTANTTOT');
      end;
    end
    else begin
      MtTot := TobRetenueSal.Detail[i].Getvalue('PRE_MONTANTTOT');
      MtMens := TobRetenueSal.Detail[i].Getvalue('PRE_MONTANTTOT');
    end;
    MtRetenue := MtMens;
    //DEBUT PT29
    CumulVerse := 0;
    CumulArriere := 0;
    Q := OpenSQL('SELECT SUM(PHR_MONTANT) CUMULVERSE,SUM(PHR_MONTANTMENS) MONTANTTHEO FROM HISTORETENUE ' +
      'WHERE PHR_SALARIE="' + Salarie +
      '" AND PHR_ORDRE=' + IntToStr(NumOrdre) + ' ' + // DB2
      'AND PHR_DATEPAIEMENT <"' + UsDateTime(DateDebPaie) + '" AND PHR_EFFECTUE="X"', True);  //PT60 On ne prends que les saisies précédentes
    if not Q.Eof then
    begin
      CumulVerse := Q.FindField('CUMULVERSE').AsFloat;
      CumulTheo := Q.FindField('MONTANTTHEO').AsFloat;
      CumulArriere := CumulTheo - CumulVerse;
      if CumulArriere < 0 then CumulArriere := 0;
    end;
    Ferme(Q);
    if RembTotalite = True then MtRetenue := MtMens - CumulVerse
    else if AvecEcheancier = True then MtRetenue := MtMens + CumulArriere
    else
    begin
      if CumulVerse + MtMens > MtTot then MtRetenue := (MtTot - CumulVerse)
      else MtRetenue := MtMens + CumulArriere;
    end;
    //FIN PT29

    {        If TypeSaisie='FAI' then     // MIS EN COMMENTAIRE LE 27/10/03 CAR ON FORCE LA LIMITATION AU RMI DANS TOUS LES CAS
            begin
                    If MtRetenue <= PartiePensionEtat then
                    begin
                            TotalRetenue := MtRetenue;
                            PartiePensionEtat := PartiePensionEtat - MtRetenue;
                    end
                    else
                    begin
                            MtIntermediaire := MtRetenue - PartiePensionEtat;
                            If MtIntermediaire <= PartiePensionEtat then
                            begin
                                    TotalRetenue := PartiePensionEtat + MtIntermediaire;
                                    PartiePensionEtat := PartiePensionEtat - MtIntermediaire;
                                    PartieEtat := 0;
                            end
                            else
                            begin
                                    MtIntermediaire := MtRetenue - PartieEtat - PartiePension;
                                    If MtIntermediaire <= PartieAutreCreancier then
                                    begin
                                            TotalRetenue := PartieEtat + PartiePension + MtIntermediaire;
                                            PartiePension := 0;
                                            PartieEtat := 0;
                                            PartieAutreCreancier := PartieAutreCreancier - MtIntermediaire;
                                    end
                                    else
                                    begin
                                            TotalRetenue := PartieEtat + PartiePension + PartieAutreCreancier;
                                            PartiePension := 0;
                                            PartieEtat := 0;
                                            PartieAutreCreancier := 0;
                                    end;
                            end;
                    end;
            end;                                     }
    if (TypeSaisie = 'FRI') or (typesaisie = 'FAI') then
    begin
      {Cas d'une retenue sur la partie relativement insaisissable,
      si le montant est supérieur on prend ensuite sur la partie saisissable}
      if MtRetenue <= PartiePensionEtat then
      begin
        TotalRetenue := MtRetenue;
        PartiePensionEtat := PartiePensionEtat - MtRetenue;
      end
      else
      begin
        MtIntermediaire := MtRetenue - PartiePensionEtat;
        if MtIntermediaire <= PartieAutreCreancier then
        begin
          TotalRetenue := PartiePensionEtat + MtIntermediaire;
          PartiePensionEtat := 0;
          PartieAutreCreancier := PartieAutreCreancier - MtIntermediaire;
        end
        else
        begin
          TotalRetenue := PartiePensionEtat + PartieAutreCreancier;
          PartiePensionEtat := 0;
          PartieAutreCreancier := 0;
        end;
      end;
    end;
    if TypeSaisie = 'FSA' then
    begin
      if (CumulRet + PartieAutreCreancier) > (SalaireNetAC - RMI) then PartieAutreCreancier := SalaireNet - (RMI + CumulRet);
      if PartieAutreCreancier < 0 then PartieAutreCreancier := 0;
      {Cas d'une retenue sur la partie saisissable du salaire}
      if MtRetenue <= PartieAutreCreancier then
      begin
        TotalRetenue := MtRetenue;
        PartieAutreCreancier := PartieAutreCreancier - MtRetenue;
      end
      else
      begin
        TotalRetenue := PartieAutreCreancier;
        PartieAutreCreancier := 0;
      end;
    end;
    // DEBUT PT31
    if TypeSaisie = 'F10' then
    begin
      if (CumulRet + PartiePret) > (SalaireNetAC - RMI) then PartiePret := SalaireNet - (RMI + CumulRet);
      if PartiePret < 0 then PartiePret := 0;
      {Cas d'une retenue limité a 10 % du salaire net}
      if MtRetenue <= PartiePret then
      begin
        TotalRetenue := MtRetenue;
        PartiePret := PartiePret - MtRetenue;
      end
      else
      begin
        TotalRetenue := PartiePret;
        PartiePret := 0;
      end;
      //PT54 FQ 14498 Solde automatique des prêts lors du solde de tout compte
      if SoldeDeToutCompte And GetParamSocSecur('SO_PGSAISIEARRETSOLDEPRET', False) then
      begin
        TotalRetenue := (MtTot - CumulVerse);
        PartiePret := 0;
      end;
    end;
    // FIN PT31
    MtArriere := 0;
    MtReprise := 0;
    //DEBUT PT43
    if MtMens > TotalRetenue then MtArriere := MtMens - TotalRetenue;
    if MtMens < TotalRetenue then MtReprise := TotalRetenue - MtMens;
    //FIN PT43
    CumulRet := CumulRet + TotalRetenue;
    TobRetenueSal.Detail[i].AddChampSupValeur('MONTANTBUL', Totalretenue, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('DATEDEBPAIE', DateDebPaie, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('DATEFINPAIE', DateFinPaie, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('ARRIERE', MtArriere, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('REPRISE', MtReprise, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('CUMULARRIERE', CumulArriere, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('CUMULVERSE', CumulVerse, False);
    TobRetenueSal.Detail[i].AddChampSupValeur('INTBULL', '-', False);
  end;
  Result := TobRetenueSal;
end;
{$ENDIF}

{$IFNDEF GCGC}   //PT56
{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Fonction de calcul de la partie saisissable pour les retenues
Suite ........ : sur salaires selon les quotités par tranches de rémunération
Mots clefs ... :
*****************************************************************}

function PGCalcPartieSaisissable(DateDebPaie: TDateTime; NbACharge: Integer; SalaireNet, ValeurRMI: Double; LeType: string): Double;
var
  TobTrancheRem: Tob;
  i: Integer;
  Total, SalaireInf, SalaireSup, AugACharge, TotalTranche, SalaireCalc: Double;
  Denominateur, Numerateur: Integer;
  Q: TQuery;
  DateValidite: TDateTime;
  CorrectionNbCharge: Integer;
begin
  if (NbACharge = 90) or (NbACharge = 99) then CorrectionNbCharge := 0 //PT52
  else CorrectionNbCharge := NbACharge;
  if SalaireNet <= ValeurRMI then
  begin
    result := 0;
    Exit;
  end;
  // récupération des tranches de rémunération de la table TRANCHERETENUESAL
  DateValidite := IDate1900;
  Q := OpenSQL('SELECT PTR_CALCFRACTION,PCD_DATEDEBUT FROM TYPERETENUE ' +
    'LEFT JOIN CALCFRACTIONDETAIL ON PTR_CALCFRACTION=PCD_CALCFRACTION ' +
    'WHERE PTR_RETENUESAL="003" AND PCD_DATEDEBUT<="' + UsDateTime(DateDebPaie) + '" ' +
    'ORDER BY PCD_DATEDEBUT DESC', True);
  if not Q.eof then
  begin
    Q.First;
    DateValidite := Q.FindField('PCD_DATEDEBUT').AsdateTime;
  end;
  Ferme(Q);
  Q := OpenSQL('SELECT PCD_MONTANTINF,PCD_MONTANTSUP,PCD_AUGACHARGE,PCD_NUMERATEUR,PCD_DENOMINATEUR FROM CALCFRACTIONDETAIL' +
    ' WHERE PCD_DATEDEBUT="' + UsDateTime(DateValidite) + '"' +
    ' ORDER BY PCD_ORDRE', True);
  TobTrancheRem := Tob.Create('Les rémunérations', nil, -1);
  TobTrancheRem.LoadDetailDB('Les rémunérations', '', '', Q, False);
  Ferme(Q);
  Total := 0;
  for i := 0 to TobTrancheRem.Detail.Count - 1 do
  begin
    SalaireInf := TobTrancheRem.Detail[i].GetValue('PCD_MONTANTINF');
    SalaireSup := TobTrancheRem.Detail[i].GetValue('PCD_MONTANTSUP');
    AugACharge := TobTrancheRem.Detail[i].GetValue('PCD_AUGACHARGE');
    Numerateur := TobTrancheRem.Detail[i].GetValue('PCD_NUMERATEUR');
    Denominateur := TobTrancheRem.Detail[i].GetValue('PCD_DENOMINATEUR');
    // Calcul augmentation des seuils selon le nbre de personnes à charge
    if SalaireInf > 0 then SalaireInf := (AugACharge * CorrectionNbCharge) + SalaireInf;
    if SalaireSup > 0 then SalaireSup := (AugACharge * CorrectionNbCharge) + SalaireSup;
    if SalaireNet > SalaireInf then
    begin
      if (SalaireNet > SalaireSup) and (SalaireSup > 0) then SalaireCalc := SalaireSup - SalaireInf
      else SalaireCalc := SalaireNet - SalaireInf;
      TotalTranche := arrondi((SalaireCalc * Numerateur) / Denominateur, 2);
      if SalaireNet - (Total + TotalTranche) < ValeurRMI then TotalTranche := SalaireNet - Total - ValeurRMI;
      Total := Total + TotalTranche;
    end;
  end;
  TobTrancheRem.Free;
  result := Total;
end;
{$ENDIF}

{$IFNDEF GCGC}   //PT56
{***********A.G.L.***********************************************
Auteur  ...... : Jérôme LEFEVRE
Créé le ...... : 15/04/2003
Modifié le ... : 15/04/2003
Description .. : Fonction de MAJ de l'historique des retenues sur salaires,
Suite ........ : attention la tob passé en paramètre est celle issue de la
Suite ........ : fonction PGCalcSaisieArret.
Suite ........ : Les arriérés ne sont pas gérés
Mots clefs ... :
*****************************************************************}

function PGMajHistoRetenue(TobMajHisto: Tob; DatePaiement: TDateTime): Boolean;
var
  TobHistoRetenue, TH: Tob;
  Q: TQuery;
  Salarie: string;
  DateDebut, DateFin: TDateTime;
  Ordre, i: Integer;
  MtMensuel, MtPaye, Arriere, Reprise, CumulArriere, CumulVerse, MtTot: Double;
begin
  Result := False;
  if TobMajHisto = nil then exit;
  // Maj de l'historique à partirde la tob de la fonction PGCalcSaisieArret
  for i := 0 to TobMajHisto.Detail.Count - 1 do
  begin
    if (TobMajHisto.Detail[i].GetValue('INTBULL') = 'X') then // and (TobMajHisto.Detail[i].GetValue('CREERHISTO') = 'X') then //PT23
    begin
      Salarie := TobMajHisto.Detail[i].GetValue('PRE_SALARIE');
      DateDebut := TobMajHisto.Detail[i].GetValue('DATEDEBPAIE');
      DateFin := TobMajHisto.Detail[i].GetValue('DATEFINPAIE');
      Ordre := TobMajHisto.Detail[i].GetValue('PRE_ORDRE');
      MtMensuel := TobMajHisto.Detail[i].GetValue('PRE_MONTANTMENS');
      MtTot := TobMajHisto.Detail[i].GetValue('PRE_MONTANTTOT');
      MtPaye := TobMajHisto.Detail[i].GetValue('MONTANTBUL');
      Arriere := TobMajHisto.Detail[i].GetValue('ARRIERE');
      Reprise := TobMajHisto.Detail[i].GetValue('REPRISE');
      CumulArriere := TobMajHisto.Detail[i].GetValue('CUMULARRIERE');
      CumulVerse := TobMajHisto.Detail[i].GetValue('CUMULVERSE');
      CumulArriere := CumulArriere + Arriere - Reprise;
      CumulVerse := CumulVerse + MtPaye;
      Q := OpenSQL('SELECT * FROM HISTORETENUE WHERE ' +
        'PHR_SALARIE="' + Salarie +
        '" AND PHR_ORDRE=' + IntToStr(Ordre) + ' ' + // DB2
        'AND PHR_DATEDEBUT="' + UsDateTime(DateDebut) + '"', True);
      TobHistoRetenue := Tob.Create('Historique', nil, -1);
      TobHistoRetenue.LoadDetailDB('Historique', '', '', Q, false);
      Ferme(Q);
      TH := Tob.Create('HISTORETENUE', TobHistoRetenue, -1);
      TH.PutValue('PHR_SALARIE', Salarie);
      TH.PutValue('PHR_ORDRE', Ordre);
      TH.PutValue('PHR_DATEDEBUT', DateDebut);
      TH.PutValue('PHR_DATEFIN', DateFin);
      TH.PutValue('PHR_MONTANT', MtPaye);
      TH.PutValue('PHR_MONTANTMENS', MtMensuel);
      TH.PutValue('PHR_REPRISEARR', Reprise);
      TH.PutValue('PHR_ARRIERE', Arriere);
      TH.PutValue('PHR_DATEPAIEMENT', DatePaiement);
      TH.PutValue('PHR_CUMULARRIERE', CumulArriere);
      TH.PutValue('PHR_CUMULVERSE', CumulVerse);
      TH.PutValue('PHR_EFFECTUE', 'X');
      TH.InsertOrUpdateDB(FALSE);
      TobHistoRetenue.Free;
      if CumulVerse = MtTot then ExecuteSQL('UPDATE RETENUESALAIRE SET PRE_CLOTURE="X",PRE_REMBEFFECTUE="X" ' +
          'WHERE PRE_SALARIE="' + Salarie + '" AND PRE_ORDRE=' + IntToStr(Ordre) + '');
    end;
  end;
  Result := True;
end;
// PT25  : 25/05/2004 V_5.0 PH Rajout fonction création automatique des salariés dans la table INTRIMAIRES
{$ENDIF}

procedure DupliqSalInterimaire;
var
  TobInterim, T, TS, TobSalaries, TobChamp: Tob;
  Q: TQuery;
  Salarie, TypeChamp, Champ, Prefixe, Suffixe: string;
  UneDate: TDateTime;
  i: Integer;
begin
  Q := OpenSQL('SELECT * FROM DECHAMPS WHERE DH_PREFIXE="PSI"', True);
  TobChamp := Tob.Create('MesChamps', nil, -1);
  TobChamp.LoadDetailDB('MesChamps', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT * FROM INTERIMAIRES WHERE PSI_TYPEINTERIM="SAL"', True);
  TobInterim := Tob.Create('INTERIMAIRES', nil, -1);
  TobInterim.LoadDetailDB('INTERIMAIRES', '', '', Q, False);
  Ferme(Q);
  Q := OpenSQL('SELECT * FROM SALARIES', True);
  TobSalaries := Tob.Create('les SALARIES', nil, -1);
  TobSalaries.LoadDetailDB('les SALARIES', '', '', Q, False);
  Ferme(Q);
  Ts := TobSalaries.FindFirst([''], [''], False);
  while ts <> nil do
  begin
    Salarie := Ts.GetValue('PSA_SALARIE');
    T := TobInterim.FindFirst(['PSI_INTERIMAIRE'], [Salarie], False);
    if T = nil then T := Tob.Create('INTERIMAIRES', TobInterim, -1);
    for i := 0 to TobChamp.Detail.Count - 1 do
    begin
      Champ := TobChamp.Detail[i].GetValue('DH_NOMCHAMP');
      TypeChamp := TobChamp.Detail[i].GetValue('DH_TYPECHAMP');
      Prefixe := ReadTokenPipe(Champ, '_');
      Suffixe := Champ;
      if (Suffixe = 'DISPOMOIS') or (Suffixe = 'SALAIREDEM') or (suffixe = 'EMAIL') or (Suffixe = 'PORTABLE') then continue
        // Attention aux champs existants dans la table interimaires et qui n'existent pas dans la table salaries
      else if Suffixe = 'TYPEINTERIM' then T.PutValue('PSI_TYPEINTERIM', 'SAL')
      else if Suffixe = 'INTERIMAIRE' then T.PutValue('PSI_INTERIMAIRE', Salarie)
      // DEBUT PT49
      else if Suffixe = 'NODOSSIER' then T.PutValue('PSI_NODOSSIER', V_PGI.NoDossier)
      else if Suffixe = 'PREDEFINI' then T.PutValue('PSI_PREDEFINI', 'DOS')
      // FIN PT49
        //      else if ((suffixe = 'EMAIL') or (Suffixe = 'PORTABLE')) then T.PutValue('PSI_' + Suffixe, '')
      else
      begin
        if TypeChamp = 'DATE' then
        begin
          if TS.GetValue('PSA_' + Suffixe) <> Null then UneDate := TS.GetValue('PSA_' + Suffixe)
          else UneDate := IDate1900;
          T.PutValue('PSI_' + Suffixe, UneDate);
        end
        else T.PutValue('PSI_' + Suffixe, Ts.GetValue('PSA_' + Suffixe));
      end;
    end;
    T.SetAllModifie(TRUE); //PT49
    Ts := TobSalaries.Findnext([''], [''], False);
  end;
  TobChamp.Free;
//  TobInterim.SetAllModifie(TRUE);   PT49
  TobInterim.InsertOrUpdateDB(False);
  TobInterim.Free;
  TobSalaries.Free;
end;
// FIN PT25
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 19/04/2005
Modifié le ... : 19/04/2005
Description .. : Fonction qui renseigne les valeurs autorisées par défaut
Suite ........ : au niveau de la sélection des établissements.
Suite ........ : Elle prend  en compte les combos et les lookup.
Suite ........ :
Suite ........ : Le filtrage est identique à celui de la compta et de la
Suite ........ : GC/GRC.
Suite ........ :
Suite ........ : Elle est basée sur la restriction des utilisateurs dans le
Suite ........ : pgimajver
Mots clefs ... : CONFIDENTIALITE
*****************************************************************}

procedure PGPositionneEtabUser(LeEtab: TControl; IgnoreForce: Boolean = False);
var
  Etab, LEtab: string;
  Forcer: Boolean;
begin
  if LeEtab = nil then Exit;
  if not VH^.EtablisCpta then Exit;
  if not LeEtab.Visible then Exit;
  Etab := VH^.ProfilUserC[prEtablissement].Etablissement;
  LEtab := RechDOM('TTETABLISSEMENT', Etab, FALSE);
  if (LEtab = '') or (UPPERCASE(LETAB) = 'ERROR') then Etab := ''; // Etablissement inconnu
  if Etab = '' then Exit;
  Forcer := VH^.ProfilUserC[prEtablissement].ForceEtab;
  if LeEtab is ThValComboBox then
  begin
    if ThValCOmboBox(LeEtab).Values.IndexOf(Etab) < 0 then Exit;
    if ((not LeEtab.Enabled) and (ThValCOmboBox(LeEtab).Value <> Etab)) then Exit;
    ThValCOmboBox(LeEtab).Value := Etab;
  end
  else if LeEtab is ThMultiValComboBox then
  begin
    if ThMultiValCOmboBox(LeEtab).Values.IndexOf(Etab) < 0 then Exit;
    if ((not LeEtab.Enabled) and (ThMultiValCOmboBox(LeEtab).Text <> Etab)) then Exit;
    ThMultiValCOmboBox(LeEtab).Text := Etab;
  end
  else if LeEtab is ThEdit then
  begin
    if ((not LeEtab.Enabled) and (ThEdit(LeEtab).Text <> Etab)) then Exit;
    ThEdit(LeEtab).Text := Etab;
  end;
  if ((Forcer) and (not IgnoreForce)) then LeEtab.Enabled := False;
end;

procedure PGEnabledEtabUser(LeEtab: TControl);
var
  Etab, LEtab: string;
begin
  if LeEtab = nil then Exit;
  if not VH^.EtablisCpta then Exit;
  if not LeEtab.Visible then Exit;
  Etab := VH^.ProfilUserC[prEtablissement].Etablissement;
  LEtab := RechDOM('TTETABLISSEMENT', Etab, FALSE);
  if (LEtab = '') or (UPPERCASE(LETAB) = 'ERROR') then Etab := ''; // Etablissement inconnu
  if Etab = '' then Exit;
  if VH^.ProfilUserC[prEtablissement].ForceEtab then LeEtab.enabled := FALSE;
end;

function PGRendEtabUser(): string;
var LEtab: string;
begin
  if not VH^.EtablisCpta then result := ''
  else result := VH^.ProfilUserC[prEtablissement].Etablissement;
  if result = '' then exit; // Pas d'etablissement donc pas de restriction
  LEtab := RechDOM('TTETABLISSEMENT', result, FALSE);
  if (LEtab = '') or (UPPERCASE(LETAB) = 'ERROR') then result := ''; // Etablissement inconnu
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 19/04/2005
Modifié le ... :   /  /
Description .. : Fonction de filtrage des TOF par etablissement et par
Suite ........ : utilisateur.
Suite ........ : Cette fonction est appelée par la TOF via un pointeur
Suite ........ : sur fonction.
Mots clefs ... : CONFIDENTIALITE
*****************************************************************}

function PGAffecteEtabByUser(FF: TForm): Boolean; // ; AffecteFiltre: boolean Affectation etablissement par utilisateur
var
  i: integer;
  LeControl: TControl;
  BTN: TToolBarButton97;
  EtabTrouve: Boolean; { PT38 }
  St1, St2: string;
begin
// PT44 fonction revue entièrement
  result := FALSE;
  EtabTrouve := FALSE;
  if MonHabilitation.Active then
  begin
    PGrechwhere(MonHabilitation.Codepop, MonHabilitation.population, False, st1, st2);
    if St1 <> '' then MonHabilitation.LeSQL := st1;
  end
  else
  begin
    for i := 0 to FF.ComponentCount - 1 do
    begin
      LeControl := TControl(FF.Components[i]);
      if LeControl is THLabel then continue;
      if (LeControl.Name = 'VCBETAB') or (LeControl.Name = 'ETAB') then
      begin
        PGPositionneEtabUser(LeControl);
        EtabTrouve := True; { PT38 }
        continue;
      end;
      if (pos('ETABLISSEMENT', LeControl.Name) > 0) or (pos('VARETAB', LeControl.Name) > 0)
        or (pos('ETAB2', LeControl.Name) > 0) or (pos('ETAB1', LeControl.Name) > 0) then // PT40
      begin
        PGPositionneEtabUser(LeControl);
        EtabTrouve := True; { PT38 }
      end
      else
        if pos('ETABLISSEMENT_', LeControl.Name) > 0 then
        begin
          PGPositionneEtabUser(LeControl);
          EtabTrouve := True; { PT38 }
        end
        else
        begin // Prise en compte du lookup salarié pour avoir une liste avec Nom,prénom
          if (LeControl is THEdit) and ((pos('SALARIE', LeControl.Name) > 0) or (pos('SALARIE_', LeControl.Name) > 0)) then
          begin
            THedit(LeControl).DataType := 'PGSALARIELOOK';
          end
          else Continue;
        end;
    end;
  end;
  if (FF.Name = 'ETATSCHAINES') then exit; // PT50
  if (not MonHabilitation.Active) and EtabTrouve then
  begin
    for i := 0 to FF.ComponentCount - 1 do
    begin
      LeControl := TControl(FF.Components[i]);
      if (LeControl.Name = 'BCherche') and (FF.Components[i] is TToolBarButton97) then
      begin
        BTN := TToolBarButton97(FF.Components[i]);
        if Assigned(Btn) then BTN.OnClick(nil);
        exit;
      end;
    end;
  end;
  if MonHabilitation.Active then
  begin
    for i := 0 to FF.ComponentCount - 1 do
    begin
      LeControl := TControl(FF.Components[i]);
      if LeControl is TTabsheet then
        if (LeControl.Name = 'PCritere') or (LeControl.Name = 'Standards') then
        begin
          if (MonHabilitation.Active) and (MonHabilitation.SqlPop <> '') then
            if MonHabilitation.Habilitation then
            begin
              if POS (MonHabilitation.SqlPop,MonHabilitation.LeSQL) = 0  then
                MonHabilitation.LeSQL := MonHabilitation.LeSQL + MonHabilitation.SqlPop;
            end
            else MonHabilitation.LeSQL := MonHabilitation.SqlPop;
          PGAjoutComposant(TTabsheet(LeControl), FF);
          Break;
        end
    end;
    if (FF.Name = 'ETATSCHAINES') then exit; // PT50
    for i := 0 to FF.ComponentCount - 1 do
    begin
      LeControl := TControl(FF.Components[i]);
      if (LeControl.Name = 'BCherche') and (FF.Components[i] is TToolBarButton97) then
      begin
        BTN := TToolBarButton97(FF.Components[i]);
        if Assigned(Btn) then
          if BTN.Visible then //PT53
            BTN.OnClick(nil);
        break;
      end;
    end;
  end;

end;

function PGRendSQLMaster(): string;
begin
  { DEB PT42 }
  if PGisOracle then result := ' '
  else
    if PGisMssql or PGisSYBASE then result := '@@select cast(name as varchar) labase from master..sysdatabases'
    else result := '';
  { FIN PT42 }
{  case V_PGI.Driver of PT42 Mise en commentaire
    dbORACLE7, dbORACLE8:
    dbMSSQL, dbMSSQL2005,dbSYBASE: result := '@@select cast(name as varchar) labase from master..sysdatabases';// PT41
//    dbMSSQL, dbSYBASE: result := '@@select name from master.dbo.sysdatabases ';
  else result := '';
  end;}
end;

function PGRendListeBasesPGI(): TOB;
var LaTob, T1: TOB;
  Q, Q1: TQuery;
  st, labase: string;
  i: Integer;
  Suivante: Boolean;
const TableSysSQL: array[0..5] of string = ('master', 'model', 'msdb', 'Northwind', 'pubs', 'tempdb');
begin
  St := PGRendSQLMaster();
  if St = '' then
  begin
    Result := nil;
    exit;
  end;
  try
    Q := OpenSql(st, TRUE); //-1,'',FALSE,'',TRUE);
    if Q.EOF then
    begin
      Result := nil;
    end
    else
    begin
      LaTob := TOB.Create('La liste des tables', nil, -1);
      while not Q.EOF do
      begin
        try
{
      if TableExiste('DECHAMPS', Q.findField('labase').AsString, FALSE) then
      begin
        T1 := TOB.create('_Labase', LaTob, -1);
        T1.AddChampSup('NOMBASE', FALSE);
        T1.PutValue('NOMBASE', Q.findField('labase').AsString+'.DBO.');
      end;
}
          Labase := Q.Fields[0].AsString;
          if (LaBase <> '') then
          begin
            Suivante := FALSE;
            for i := 0 to 5 do
            begin
              if LaBase = TableSysSQL[i] then Suivante := TRUE;
            end;
            if not Suivante then
            begin
              st := '@@select count(*) from ' + LaBase + '.dbo.sysobjects where name=' + '''PAIEENCOURS'' ';
              Q1 := OpenSql(st, TRUE);
              if not Q1.EOF then
              begin
                if Q1.Fields[0].AsInteger > 0 then // C'est une base PGI
                begin
                  T1 := TOB.create('_Labase', LaTob, -1);
                  T1.AddChampSup('NOMBASE', FALSE);
                  T1.PutValue('NOMBASE', LaBase + '.DBO.');
                end;
              end;
              FERME(Q1);
            end;
          end;
        except
        end;
        Q.Next;
      end;
      result := LaTob;
    end;
    Ferme(Q);
  except
    Result := nil;
  end;
end;

function PGRendNbreSal(): Integer;
var
  Q: Tquery;
  UneTob: TOB;
  i, LeNbre: Integer;
  LaBase, st: string;
begin
  UneTob := PGRendListeBasesPGI();
  LeNbre := 0;
  if UneTob <> nil then
  begin
    if UneTob.detail.count - 1 > 0 then
    begin
      for i := 0 to UneTob.detail.count - 1 do
      begin
        LaBase := UneTob.detail[i].GetValue('NOMBASE');
        St := 'SELECT COUNT(*) NBRE FROM ' + LaBase + 'SALARIES WHERE PSA_DATESORTIE <= "' + UsDateTime(IDate1900) + '" OR PSA_DATESORTIE > "' + UsDateTime(Now) + '"';
        Q := OpenSQL(st, TRUE); // FALSE, -1, '', FALSE, LaBase);
        if not Q.EOF then LeNbre := LeNbre + Q.FindField('NBRE').AsInteger;
        ferme(Q);
      end;
    end;
  end;
  result := LeNbre;
end;
// DEB PT44

procedure CreateObjHabilitation();
begin
  if not Assigned(MonHabilitation) then
  begin
    MonHabilitation := T_Habilit.Create;
    MonHabilitation.LeSQL := '';
    MonHabilitation.SqlPop := '';
    MonHabilitation.LibellePop := '';
    MonHabilitation.LibelleHab := '';
    MonHabilitation.CodePop := '';
    MonHabilitation.Population := '';
    MonHabilitation.Active := FALSE;
    MonHabilitation.Habilitation := FALSE;
    MonHabilitation.MesDroits := TStringList.Create;
    MonHabilitation.LesPrefixes := TStringList.Create;
  end;
end;

procedure PgRendHabilitation(codePop: string = ''; Population: string = ''; Libellepop: string = '');
var Q: TQuery;
  St, St1, St2, LeGroupe: string;
begin
  CreateObjHabilitation();
  if CodePop = '' then
  begin
    st := 'SELECT PHL_CODEPOP,PHL_POPULATION,PHL_LIBELLE FROM PGHABILITATIONS WHERE PHL_USER="' + V_PGI.User + '" AND ##PHL_PREDEFINI##';
    Q := OPENSQL(st, True);
    if not Q.Eof then
    begin
      MonHabilitation.CodePop := Q.FindField('PHL_CODEPOP').AsString;
      MonHabilitation.Population := Q.FindField('PHL_POPULATION').AsString;
      if (MonHabilitation.CodePop <> '') and (MonHabilitation.Population <> '') then
      begin
        MonHabilitation.Active := TRUE;
        MonHabilitation.Habilitation := TRUE;
        MonHabilitation.LibelleHab := Q.FindField('PHL_LIBELLE').AsString;
      end;
      Ferme(Q);
    end
    else
    begin
      Ferme(Q);
      st := 'SELECT US_GROUPE FROM UTILISAT WHERE US_UTILISATEUR="' + V_PGI.User + '"';
      Q := OPENSQL(st, True);
      if not Q.Eof then // Recherche du groupe
        LeGroupe := Q.findField('US_GROUPE').AsString;
      Ferme(Q);
      st := 'SELECT PHL_CODEPOP,PHL_POPULATION,PHL_LIBELLE FROM PGHABILITATIONS WHERE PHL_USERGRP="' + LeGroupe + '" AND ##PHL_PREDEFINI##';
      Q := OPENSQL(st, True);
      if not Q.Eof then
      begin
        MonHabilitation.CodePop := Q.FindField('PHL_CODEPOP').AsString;
        MonHabilitation.Population := Q.FindField('PHL_POPULATION').AsString;
        if (MonHabilitation.CodePop <> '') and (MonHabilitation.Population <> '') then
        begin
          MonHabilitation.Active := TRUE;
          MonHabilitation.Habilitation := TRUE;          
          MonHabilitation.LibelleHab := Q.FindField('PHL_LIBELLE').AsString;
        end;
      end;
      Ferme(Q);
    end;
  end
  else
  begin
    MonHabilitation.CodePop := CodePop;
    MonHabilitation.Population := Population;
    MonHabilitation.Active := TRUE;
    MonHabilitation.LibelleHab := Libellepop;
    MonHabilitation.Habilitation := FALSE;
   end;
  if MonHabilitation.Active then
  begin
    PGrechwhere(MonHabilitation.Codepop, MonHabilitation.population, False, st1, st2);
    if St1 <> '' then MonHabilitation.LeSQL := st1;
    st := 'SELECT dd.dh_prefixe LEPREFIXE from dechamps dd left join detables ' +
      'on dt_prefixe=dd.dh_prefixe where dd.dh_nomchamp like "P%_ETABLISSEMENT" ' +
      'and exists (SELECT DA.Dh_NOMCHAMP FROM DEchamps DA where da.dh_nomchamp=dd.dh_prefixe||"_SALARIE")';
    Q := OPENSQL(st, True);
    while not Q.Eof do
    begin
      MonHabilitation.LesPrefixes.Add(Q.FindField('LEPREFIXE').AsString);
      Q.Next;
    end;
    Ferme(Q);
  end;
end;

function PGControlHab(FF: TForm): Boolean;
var LePrefixe, st, Etab: string;
  i, j: Integer;
  LeControl, LePref: TControl;
begin
  CreateObjHabilitation();
  result := FALSE;
  if Assigned(MonHabilitation) then
  begin
//    Monhabilitation.LeSql := '';
    for i := 0 to FF.ComponentCount - 1 do
    begin
      LeControl := TControl(FF.Components[i]);
      if LeControl is THLabel then continue;
      if pos('ETABLISSEMENT', LeControl.Name) > 0 then // Le composant etablissement existe
      begin
        LePrefixe := Copy(LeControl.Name, 1, 3);
        if LePrefixe[1] <> 'P' then
        begin
          for j := 0 to FF.ComponentCount - 1 do
          begin
            LePref := TControl(FF.Components[j]);
            if LePref.name = 'LEPREFIXE' then
            begin
              LePrefixe := THEdit(LePref).Text;
              break;
            end;
          end;
        end;
        if LePrefixe[1] = 'P' then // Table du domaine de la paie
        begin
          if MonHabilitation.InclureSal(LePrefixe) then // Champ Salarie et/ou etablissement
          begin
            st := LePrefixe + '_SALARIE IN ' + ' (SELECT PSA_SALARIE FROM SALARIES WHERE ' + MonHabilitation.LeSQL + MonHabilitation.SqlPop + ')';
            Monhabilitation.LeSql := st;
          end
          else
          begin // Champs etablissement seul comme critere de la population
            st := MonHabilitation.LesEtab;
            Etab := ReadTokenSt(St);
            MonHabilitation.LeSQL := '';
            j := 0;
            while Etab <> '' do
            begin
              j := j + 1;
              if Etab <> '' then
              begin
                if j > 1 then MonHabilitation.LeSQL := MonHabilitation.LeSQL + ',';
                MonHabilitation.LeSQL := MonHabilitation.LeSQL + '"' + Etab + '"';
              end;
              Etab := ReadTokenSt(St);
            end;
            if j > 0 then
              MonHabilitation.LeSQL := LeControl.Name + ' IN (' + MonHabilitation.LeSQL + MonHabilitation.SqlPop + ')';
          end;
          Result := TRUE;
          break;
        end;
      end;
    end;
  end;
  if (MonHabilitation.Active) and (MonHabilitation.SqlPop <> '') then
    if MonHabilitation.Habilitation then
    begin
      if POS (MonHabilitation.SqlPop,MonHabilitation.LeSQL) = 0  then
        MonHabilitation.LeSQL := MonHabilitation.LeSQL + MonHabilitation.SqlPop;
    end
    else MonHabilitation.LeSQL := MonHabilitation.SqlPop;
end;

procedure PGAjoutComposant(THBST: TTabsheet; FF: TForm);
var
  C: THEdit;
  i: Integer;
  LeControl: TControl;
  Creat: Boolean;
begin
  if not (MonHabilitation.Active) then exit;
  creat := true;
  for i := 0 to FF.ComponentCount - 1 do
  begin
    LeControl := TControl(FF.Components[i]);
    if LeControl.name = 'XX_WHERE2' then
    begin
      Creat := FALSE;
      C := THEdit(LeControl);
      if V_PGI.SAV then
        if C <> nil then C.Visible := TRUE;
      break;
    end;
  end;
  if Creat then
  begin
    C := THEdit.Create(FF);
    C.Parent := THBST;
    C.Left := 450;
    C.Width := 150;
    C.Name := 'XX_WHERE2';
    C.Text := '';
    C.Color := clRed;
    C.Visible := FALSE;
    if V_PGI.SAV then C.Visible := TRUE;
  end;
  if (MonHabilitation.Active) and (PGControlHab(FF)) then
    if Assigned(C) then
      C.Text := MonHabilitation.LeSQL;
end;

procedure PgVideHabilitation();
begin
  if Assigned(MonHabilitation) then
  begin
    MonHabilitation.MesDroits.Free;
    MonHabilitation.LesPrefixes.free;
    MonHabilitation.free;
    MonHabilitation := nil;
  end;
end;

procedure pgidentpop(codepop, population: string);
var
  Q1, Q2, Q3: Tquery;
  st, valident, sti, numident, nomchamp, nomtablette, car: string;
  nbident, i: integer;
begin
// Recherche de la valeur des identifiants
  Q1 := Opensql('SELECT PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4  FROM ORDREPOPULATION' +
    ' WHERE PPC_POPULATION = "' + population + '"', true);

// Recherche du nom des identifiants dans la table PAIEPARIM

  Q2 := OPENSQL('SELECT PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,' +
    ' PPO_TYPEPOP FROM CODEPOPULATION WHERE PPO_CODEPOP = "' + codepop + '" ', True);
  if not Q2.Eof then
  begin
    nbident := Q2.findfield('PPO_NBIDENT').asinteger;
    for i := 1 to nbident do
    begin
      Sti := IntToStr(i);
      numident := Q2.findfield('PPO_IDENT' + sti).asstring;
      valident := '';
      if not Q1.EOF then valident := Q1.findfield('PPC_VALIDENT' + sti).asstring;
      if valident <> '' then
      begin
        Q3 := Opensql('SELECT PAI_LIENASSOC,PAI_PREFIX,PAI_SUFFIX from PAIEPARIM WHERE PAI_IDENT = "' + numident + '"', true);
        nomtablette := Q3.findfield('PAI_LIENASSOC').asstring;
        nomchamp := Q3.findfield('PAI_PREFIX').asstring + '_' + Q3.findfield('PAI_SUFFIX').asstring;
        st := '';
        st := nomtablette + ';' + nomchamp + ';' + valident;
        car := copy(st, length(st), length(st));
        if car = ';' then st := copy(st, 1, length(st) - 1);
        if Assigned(MonHabilitation) then
        begin
          MonHabilitation.MesDroits.Add(St);
          if (POS('PSA_ETABLISSEMENT', st)) > 0 then
          begin
            MonHabilitation.CritEtab := True;
            MonHabilitation.LesEtab := valident;
          end;
          ferme(Q3);
        end;
      end;
    end;
  end;
  ferme(Q2);
  ferme(Q1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 16/06/2006
Modifié le ... :   /  /
Description .. : Recherche de la clause where pour une population
Mots clefs ... :
*****************************************************************}

procedure PGrechwhere(Codepop, population: string; histo: boolean; var st1, st2: string; Habilitation : Boolean = FALSE);
var

  st, nomchamp, sti, stj, numident: string;
  Q1, Q2, Q3: Tquery;
  nbident, i, j: integer;
  Tids, Tidh, Tob_Identsal, Tob_identhis, Tob_valident: TOB;

begin

// Recherche du nom des identifiants dans la table PAIEPARIM
  st := 'SELECT PPO_NBIDENT, PPO_IDENT1, PPO_IDENT2, PPO_IDENT3, PPO_IDENT4,PPO_LIBELLE,PPO_TYPEPOP FROM CODEPOPULATION' +
    ' Where PPO_CODEPOP = "' + codepop + '" ';
  Q1 := OPENSQL(st, True);
  if not Q1.Eof then
  begin
    nbident := Q1.findfield('PPO_NBIDENT').asinteger;
    Tob_Identsal := Tob.Create('Nom identifiants salarie', nil, -1);
    Tob_Identhis := Tob.Create('Nom identifiants histo', nil, -1);

    for i := 1 to nbident do
    begin
      Sti := IntToStr(i);
      numident := Q1.findfield('PPO_IDENT' + sti).asstring;

      Q2 := Opensql('select PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_IDENT = "' + numident + '"', true);
  // si historisation : les cbamps SEXE (27), CONVENTION (60), CATBILAN (500),CATDADS (402) doivent être pris dans
  // la fiche SALARIE car ils ne sont pas historisés
  // Deb PT47
      if ((histo = false) or (numident = '27') or (numident = '60') or (numident = '500') or (numident = '402')) then // fin PT47
      begin
        nomchamp := Q2.findfield('PAI_PREFIX').asstring + '_' + Q2.findfield('PAI_SUFFIX').asstring;
        Tids := Tob.Create('Identifiant', Tob_Identsal, -1);
        Tids.AddChampSup('NOMCHAMPS', False);
        Tids.PutValue('NOMCHAMPS', Nomchamp);
        Tids.AddChampSup('ORDREIDENT', False);
        Tids.PutValue('ORDREIDENT', sti);
      end
      else
      begin
        nomchamp := 'a.PHS_' + Q2.findfield('PAI_SUFFIX').asstring;
        for j := 1 to 4 do
        begin
          stj := inttostr(j);
          if Q2.findfield('PAI_SUFFIX').asstring = 'LIBREPCMB' + stj then nomchamp := 'a.PHS_CBLIBRE' + stj;
        end;
        Tidh := Tob.Create('Identifiant', Tob_Identhis, -1);
        Tidh.AddChampSup('NOMCHAMPS', False);
        Tidh.PutValue('NOMCHAMPS', Nomchamp);
        Tidh.AddChampSup('ORDREIDENT', False);
        Tidh.PutValue('ORDREIDENT', sti);
      end;
    end;
    ferme(Q2);
  // Recherche de la valeur des identifiants
    Q3 := Opensql('select PPC_VALIDENT1,PPC_VALIDENT2, PPC_VALIDENT3, PPC_VALIDENT4  from ORDREPOPULATION' +
      ' where PPC_POPULATION = "' + population + '"', true);
    if not Q3.EOF then
    begin
      Tob_ValIdent := Tob.Create('Valeur identifiants', nil, -1);
      Tob_ValIdent.LoadDetailDB('Valeur identifiants', '', '', Q3, False);
  // constitution de la clause WHERE  (pour table SALARIES et pour table HISTOSALARIE
      st1 := '';
      st2 := '';

      if Assigned(tob_identsal) then st1 := PGConstitutionwhere(Tob_identsal, Tob_Valident);
      if Assigned(tob_identhis) then st2 := PGConstitutionwhere(Tob_identhis, Tob_Valident);

      freeandnil(Tob_Identsal);
      freeandnil(Tob_Identhis);
      freeandnil(Tob_ValIdent);
    end;
    ferme(Q3);
  end;
  ferme(Q1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 16/06/2006
Modifié le ... :   /  /
Description .. : Constitution de la clause Where à partir des valeurs
Suite ........ : sélectionnées pour un identifiant
Mots clefs ... :
*****************************************************************}

function PGconstitutionwhere(T1, T2: Tob): string;
// T1 contient le nom des champs, T2 contient la valeur de l'identifiant
var
  nomchamp, listeid, sti, valident, listewhere, st: string;
  T_ident, T_Valeurid: Tob;

begin
  result := '';
  st := '';
  if T1 <> nil then
  begin
    T_Valeurid := T2.findfirst([''], [''], TRUE);
    T_ident := T1.FindFirst([''], [''], TRUE);
    while T_ident <> nil do
    begin
      sti := T_ident.getvalue('ORDREIDENT');
      nomchamp := T_ident.GetValue('NOMCHAMPS'); // nom du champ correspondant au n° identifiant
      listeid := T_Valeurid.Getvalue('PPC_VALIDENT' + sti); // liste de valeur de l'identifiant

      listewhere := '';
      valident := READTOKENST(Listeid);
      while valident <> '' do // lecture de chaque valeur
      begin
        listewhere := listewhere + '"' + valident + '",';
        valident := READTOKENST(Listeid);
      end;

      if listewhere <> '' then begin // si au moins une valeur à comparer
        st := st + nomchamp + ' ' + 'in' + ' ';
        st := st + '(' + copy(listewhere, 1, length(listewhere) - 1) + ')' + ' ' + 'And' + ' ';
      end;

      T_ident := T1.FindNext([''], [''], TRUE); // lecture de l'identifiant suivant
    end;

    if st <> '' then result := copy(st, 1, length(st) - 4);
  end;
end;

{ T_Habilit }

constructor T_Habilit.Create;
begin
  inherited Create;
end;

destructor T_Habilit.Destroy;
begin
  inherited Destroy;
end;

function T_Habilit.InclureSal(LePrefixe: string): Boolean;
var i: Integer;
begin
  result := FALSE;
  for i := 0 to MonHabilitation.LesPrefixes.count - 1 do
  begin
    if lePrefixe = MonHabilitation.LesPrefixes[i] then
    begin
      result := TRUE;
      break;
    end;
  end;
end;
// FIN PT44
// DEB PT45
{$IFNDEF GCGC}

procedure ChargeRegulCot(LeSalarie, TPE: TOB; DD, DF: TDateTime);
var
  TRC, THB, T1: TOB;
  Q: TQuery;
  i: Integer;
  St, Rub, Nat, cc: string;
  Mt: Double;
  NbreDec: Integer;
  cas: WORD;
begin
  if LeSalarie = nil then exit;
  St := 'SELECT * FROM PGHISTRETRO WHERE PGT_SALARIE="' + LeSalarie.GetValue('PSA_SALARIE') + '" AND ' +
    'PGT_DATEVALIDITE <= "' + UsDateTime(DF) + '" AND PGT_DATEVALIDITE >="' + UsDateTime(DD) + '"' +
    ' AND PGT_DATEINTEG="' + UsdateTime(IDate1900) + '"';
{Flux optimisé
  Q := OpenSql(St, true);
  if not Q.EOF then
  begin
    TRC := TOB.Create('PGHISTRETRO', nil, -1);
    TRC.LoadDetailDB('PGHISTRETRO', '', '', Q, False);
}
  TRC := TOB.Create('PGHISTRETRO', nil, -1);
  TRC.LoadDetailDBFROMSQL('PGHISTRETRO', St);
  for i := 0 to TRC.detail.count - 1 do
  begin
    Rub := TRC.detail[i].GetValue('PGT_RUBRIQUE');
    Nat := TRC.detail[i].GetValue('PGT_NATURERUB');
    if TPE.FindFirst(['PHB_NATURERUB', 'PHB_RUBRIQUE', 'PHB_COTREGUL'], [Nat, Rub, 'REG'], FALSE) = nil then // $$$$
    begin
      if Nat = 'COT' then
        T1 := TOB_Cotisations.FindFirst(['PCT_RUBRIQUE'], [Rub], FALSE)
      else T1 := TOB_Rem.FindFirst(['PRM_RUBRIQUE'], [Rub], FALSE);
      THB := TOB.Create('HISTOBULLETIN', TPE, -1);
      RemplirHistoBulletin(THB, LeSalarie, T1, nil, TPE.GetValue('PPU_DATEDEBUT'), TPE.GetValue('PPU_DATEFIN'));
      THB.PutValue('PHB_RUBRIQUE', Rub + '.R');
      THB.PutValue('PHB_COTREGUL', 'REG');
      THB.PutValue('PHB_LIBELLE', TRC.detail[i].GetValue('PGT_LIBELLE1'));
      if Nat = 'COT' then
      begin
        THB.PutValue('PHB_BASECOT', TRC.detail[i].GetValue('PGT_DIFFBASECOT'));
        THB.PutValue('PHB_TAUXSALARIAL', TRC.detail[i].GetValue('PGT_DIFFTXSAL'));
        THB.PutValue('PHB_TAUXPATRONAL', TRC.detail[i].GetValue('PGT_DIFFTXPAT'));
        Mt := (TRC.detail[i].GetValue('PGT_DIFFBASECOT') * TRC.detail[i].GetValue('PGT_DIFFTXSAL')) / 100;
        Mt := Arrondi(Mt, T1.GetValue('PCT_DECMTSAL'));
        THB.PutValue('PHB_MTSALARIAL', Mt);
        Mt := (TRC.detail[i].GetValue('PGT_DIFFBASECOT') * TRC.detail[i].GetValue('PGT_DIFFTXPAT')) / 100;
        Mt := Arrondi(Mt, T1.GetValue('PCT_DECMTPAT'));
        THB.PutValue('PHB_MTPATRONAL', Mt);
      end
      else
      begin
        THB.PutValue('PHB_RUBRIQUE', Rub + '.R');
        THB.PutValue('PHB_BASEREM', TRC.detail[i].GetValue('PGT_DIFFBASEREM'));
        THB.PutValue('PHB_TAUXREM', TRC.detail[i].GetValue('PGT_DIFTAUXREM'));
        THB.PutValue('PHB_COEFFREM', TRC.detail[i].GetValue('PGT_DIFCOEFFREM'));
        Mt := TRC.detail[i].GetValue('PGT_DIFFMTREM');
        Cc := T1.GetValue('PRM_CODECALCUL');
        cas := RechCasCodeCalcul(Cc);
        NbreDec := T1.GetValue('PRM_DECMONTANT');
        if Mt = 0 then
          Mt := ValoriseMt(cas, Cc, NbreDec, TRC.detail[i].GetValue('PGT_DIFFBASEREM'),
            TRC.detail[i].GetValue('PGT_DIFTAUXREM'), TRC.detail[i].GetValue('PGT_DIFCOEFFREM'));
        THB.PutValue('PHB_MTREM', Mt);
      end;
    end;
  end;
  if Assigned(TRC) then FreeANdNil(TRC);
end;
// Fonction de détopage des enreg de la rétroactivité en cas de suppression de bulletin

procedure AnnulRegulCot(LeSalarie: string; DD, DF: TDateTime);
begin
  ExecuteSql('UPDATE PGHISTRETRO SET PGT_DATEINTEG="' + UsdateTime(IDate1900) + '" WHERE PGT_SALARIE="' + LeSalarie + '" AND ' +
    'PGT_DATEVALIDITE <= "' + UsDateTime(DF) + '" AND PGT_DATEVALIDITE >="' + UsDateTime(DD) + '"');
end;
// FIN PT45

procedure PaieConceptPlanPaie(Ecran: TForm);
var STD, DOS: Boolean;
  LeControl: TControl;
begin
  if Ecran = nil then exit;
  STD := JaiLeDroitTag(200101);
  DOS := JaiLeDroitTag(200102);
  if ((not STD) and (not DOS)) then
  begin
    LeControl := TControl(Ecran.FindComponent('BDUPLIQUER'));
    if LeControl <> nil then LeControl.Visible := FALSE;
  end;
  if (not DOS) then
  begin
    LeControl := TControl(Ecran.FindComponent('BINSERT'));
    if LeControl <> nil then LeControl.Visible := FALSE;
  end;
end;

procedure PaieConceptTabMinPaie(Ecran: TForm);
var STD, DOS: Boolean;
  LeControl: TControl;
begin
  if Ecran = nil then exit;
  STD := JaiLeDroitTag(200111);
  DOS := JaiLeDroitTag(200112);
  if ((not STD) and (not DOS)) then
  begin
    LeControl := TControl(Ecran.FindComponent('BDUPLIQUER'));
    if LeControl <> nil then LeControl.Visible := FALSE;
  end;
  if (not DOS) then
  begin
    LeControl := TControl(Ecran.FindComponent('BINSERT'));
    if LeControl <> nil then LeControl.Visible := FALSE;
  end;
end;
{$ENDIF}

end.


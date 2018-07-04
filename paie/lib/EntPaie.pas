{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 11/07/2001
Modifié le ... :   /  /
Description .. : Gestion des paramètres sociétés
Mots clefs ... : PAIE;PARAMETRESOC
*****************************************************************}
{
PT1   : 11/07/2001 SB V540 Suppression de l'option Edition des codes org général
                           Remplacé par une option pour chaque code org
PT2   : 25/07/2001 VG V540 Ajout champs dans Historique salarié
PT3   : 08/08/2001 PH V540 Ajout PGCHEMINEAGL chemin de stockage des données à
                           transiter entre différentes bases
PT4   : 02/10/2001 PH V562 Ajout historisation profil remunérations
PT5   : 15/10/2001 VG V562 Ajout champs dans Historique salarié
PT6   : 10/12/2001 PH V562 Ajout indicateur si saisie déportée sérialisée donc
                           activation du menu 43 PgeAbsences
PT7   : 21/02/2002 JL V571 Ajout champ MSA,Itermittents,Responsables
PT8   : 22/02/2002 VG V571 Ajout champ BTP
PT9   : 10/04/2002 JL V571 Ajout champs gestion intérimaires
PT10  : 11/06/2002 PH V582 Ajout champs gestion des primes eAGL
PT11  : 02/07/2002 PH V582 Ajout champs autorisation de controle de creation
                           enrg dans recupxls
PT12  : 25/07/2002 SB V585 FQ n°10192 Ajout Champ Rtt
PT13  : 26/07/2002 JL V585 Ajout champ formation
PT14  : 30/08/2002 Ph V585 Ajout champ bilan social simplifié
PT15  : 03/09/2002 VG V585 Ajout champ pour utilisation de l'historique pour le
                           calcul de la DADS-U
PT16  : 06/09/2002 PH V585 Ajout champ pour 5eme zone element de salaire
PT17  : 23/09/2002 PH V585 Ajout champ méthode préférence date de la préparation
                           automatique
PT18  : 26/09/2002 PH V585 Ajout champ liste des champs gérés dans le
                           référentiel métier
PT19  : 02/10/2002 SB V585 Ajout Booléen pour recupération histo absences dans
                           calendrier excel
PT20  : 02/10/2002 JL V591 Ajout champs pour gérer format et incrément des codes
                           des stages pour formation
PT21  : 05/12/2002 SB V591 Ajout des paramètres Econgés
PT22  : 09/01/2003 PH V591 Nouveau boolean PgSeriaFormation indiquant état
                           sérialisation du module formation
PT23  : 24/04/2003 SB V_42 Integration des params pour la gestion des ressources
PT24  : 07/05/2003 MF V_42 Ajout booléen pour Ticket restaurant + 2 string code
                           client
PT25  : 18/06/2003 JL V_42 Ajout booléen pour gestion saisie arrêt + frais fixe
                           formation -> 27/04/2007 suppression frais fixe
PT26  : 18/08/2003 JL V_42 Ajout champ gestion des menus formation
PT27  : 28/08/2003 PH V_42 Ajout paramsoc nomenclature PCS 2003
PT28  : 06/05/2004 MF V_50 Ajout booléen pour Gestion IJSS + 1 string Critère de
                           maintien
PT29  : 28/06/2004 PH V_50 Ajout booléen pour Bilan social, gestion de carrière,
                           participation
PT30  : 29/06/2004 MF V_50 Ajout booléen pour Gestion du maintien, Gestion
                           Acompte, Gestion FRP, Scoring formation
PT31  : 05/07/2004 PH V_50 Ajout booléen pour Sérialisation des modules
                           complémentaires getion de carière et bilan social
PT32  : 06/10/2004 JL V_50 Ajout n° attestation ASSEDID intermittents pour gérer
                           incrément auto. -> 10/05/2007 suppression de
                           SO_PGNUMAEMINTER
PT33  : 29/03/2005 SB V_60 FQ 11990 Ajout gestion des abences
PT34  : 18/05/2005 PH V_60 FQ 11801 Ajout arrondi calcul paie à l'envers
PT35  : 18/05/2005 PH V_60 FQ 10977 Ajout gestion du libelle du tiers au lieu du
                           compte général
PT36  : 17/06/2005 MF V_60 FQ 12388 Ajout du paramètre SO_PGHISTOMAINTIEN
                           (méthode de récup du maintien)
PT37  : 11/10/2005 PH V_60 ParamSoc gestion des responsables non salarié
PT38  : 14/04/2006 SB V_65 FQ 11953 Paramsoc absence à cheval sur plusieurs mois
PT39  : 09/05/2006 PH V_65 FQ 12835 Paramsoc taux de charges
PT40  : 19/05/2006 SB V_65 FQ 13154 Paramsoc CHR 6ème semaine CP suppl.
PT41  : 30/05/2006 PH V_65 FQ 13156 Séria Diode
PT42  : 12/09/2006 NA V702 Ajout Gestion des IDR et élement national pour le
                           taux d'actualisation
PT43  : 07/11/2006 SB V_70 Ajout paramsoc pour paramètre export etemptation
PT44  : 29/12/2006 PH V_70 Ajout paramsoc pour insertion automatique des
                           rubriques manquantes dans le bulletin
PT45  : 19/01/2006 SB V_70 Ajout paramsoc seria présence
PT46  : 13/04/2007 FL V702 FQ 13568 Ajout des méthodes de calcul au prévisionnel
PT47  : 27/04/2007 VG V_72 Suppression paramSoc PgEAI
PT48  : 27/04/2007 GGU V72 Ajout paramsoc pour la validation des populations
PT49  : 04/05/2007 MF V_72 Séria des IDR
PT50  : 09/05/2007 FL V_72 FQ 13567 Ajout des paramètres de population pour la
                           formation
PT51  : 11/06/2007 VG V_72 Adaptation nouvelles méthodes CBPPath
PT52  : 02/10/2007 NA V_800 Ajout gestion de la présence ou non
PT53  : 04/12/2007 FL V_81 Ajout d'un membre de VH_PAIE pour le contrôle du lancement de l'application en mode FICHE
PT54  : 21/12/2007 FL V_81 Ajout d'un membre de VH_PAIE pour la vérification du lancement de la fiche en mode FICHE
PT55  : 20/03/2008 FL V_81 Ajout des base de partage
PT56  : 04/04/2008 FL V_81 Ajout du champ PGForMailAdr pour stocker l'adresse du responsable formation en copie des mails
PT57  : 07/04/2008 NA V_850 Ajout gestion des contrats d'assurance
PT58  : 16/04/2008 MF V_81 Ajout paramsocpour la factorisation des DUCS PGFACTODUCS
PT59  : 23/06/2008 NA V_850 Chèque déjeuner : ajout du champ "PGFACTETABL" pour gérer plusieurs codes clients par établissement
}

unit EntPaie;

interface

uses Controls, HEnt1, HMsgBox, paramsoc,
{$IFNDEF EAGLSERVER}
     PGRepertoire,
{$ENDIF}
     EntPgi;

type
  LaVariablePaie = class{$IFDEF EAGLSERVER}(TLaVariable){$ENDIF EAGLSERVER}
  private
  public
    PGNbreStatOrg, PGIterationEnvers: Integer;
    PGLibelleOrgStat1, PGLibelleOrgStat2, PGLibelleOrgStat3, PGLibelleOrgStat4, PGLibCodeStat: string;
    PGSaisieBulletin, PGProfilFnal, PGMethodHeures: string; // COMBO dans ParamSoc
    PGIncSalarie, PGCalculBulletin, PGDecalage, PGDecalagePetit, PGAnalytique, PGResponsables, PGMsa, PGIntermittents: Boolean; //PT7
    PGCongesPayes, PGInterimaires, PGCodeInterim, PGAbsence : boolean; //PT9 PT33
    PGBTP: boolean; //PT8
    PGMonnaieTenue, PGRubAcompte: string; //Monnaie de tenue
    PGTenueEuro: Boolean;
    PGDateBasculEuro: TDateTime;
    PGCheminRech, PGCheminSav, PGCheminEagl, PGPhotoSal, PGEnversNet, PGContextePaie: string;
    PGCumul01, PGCumul02, PGCumul03, PGCumul04, PGCumul05, PGCumul06: string;
    PGCritSaisRub: string;
    PGLibCg1, PGLibCg2, PGLibCg3, PGLibCg4: string;
    PGCgAcq1, PGCgPris1, PGCgAcq2, PGCgPris2, PGCgAcq3, PGCgPris3, PGCgAcq4, PGCgPris4: string;
    PgLibDate1, PgLibDate2, PgLibDate3, PgLibDate4: string; //champs libre Date
    PgLibCombo1, PgLibCombo2, PgLibCombo3, PgLibCombo4: string; //champs libre Tablette
    PgLibCoche1, PgLibCoche2, PgLibCoche3, PgLibCoche4: string; //champs libre CheckBox
    PgNbDate, PgNbCoche, PgNbCombo, PgNbSalLib: Integer;
    PgSalLib1, PgSalLib2, PgSalLib3, PgSalLib4, PgSalLib5: string; //Champs Salaire
    PgRedHeure, PGRedRem, PgTypRedRepas, PgRedRepas: string; //Réductions
    PgSeriaPaie, PgSeriaAnalyses, PgTiersAuxiAuto, PgeAbsences: Boolean; // Modules sérialisés
    PgSeriaFormation, PgSeriaDADSB, PgSeriaCompetence, PgSeriaBilan, SeriaDiodePcl: Boolean; // Modules sérialisés // PT31 PT41
    PgTypeNumSal: string; // Matricule salarié Alpha ou numérique
    PGSectionAnal, PgCreationSection: Boolean; // Creation automatique des sections analytiques
    PGAx1Anal1, PGAx1Anal2, PGAx1Anal3: string;
      // Décomposition des sections analytiques Axe 1     PGAx2Anal1,PGAx2Anal2,PGAx2Anal3 : String; // Décomposition des sections analytiques Axe 1
    PGAx2Anal1, PGAx2Anal2, PGAx2Anal3: string; // Décomposition des sections analytiques Axe 2
    PGAx3Anal1, PGAx3Anal2, PGAx3Anal3: string; // Décomposition des sections analytiques Axe 3
    PGAx4Anal1, PGAx4Anal2, PGAx4Anal3: string; // Décomposition des sections analytiques Axe 4
    PGAx5Anal1, PGAx5Anal2, PGAx5Anal3: string; // Décomposition des sections analytiques Axe 5
    PGCptNetAPayer, PGPreselect0, PGPreselect1, PGPreselect2, PGPreselect3, PGPre4select1, PGPre4select2, PGPre4select3: string; // Compte Net A Payer + Choix confection du compte
    PGModeleEcr, PGJournalPaie, PGVENTILMULETAB: string; // Modele ecriture et Journal OD par defaut
    PGLongRacine, PGLongRacine421, PGLongRacin4, PGNbreRac1, PGNbreRac2, PGNbreRac3, PGNbre4Rac1, PGNbre4Rac2, PGNbre4Rac3: Integer;
      // Nbre de caracteres composant le compte pour chaque préselection
    PGDecompoCl4, PGDecompoCl6: Boolean; // Decomposition des comptes de la classe 4,6
    PgBulDefaut: string; //Modele de bulletin par défaut
    PgTypeCumulMois1, PgTypeCumulMois2, PgTypeCumulMois3: string; //Type Cumul du mois bas de bulletin
    PgCumulMois1, PgCumulMois2, PgCumulMois3, PGIntegODPaie: string; //Cumul mois
    PgLibCumulMois1, PgLibCumulMois2, PgLibCumulMois3: string; //Libelle cumul mois
//    PgCodeEmploi, PgLibelleEmploi, PgQualification, PgCoefficient, PgIndice, PgNiveau: boolean;
//    PgCodeStat, PgHTravailN1, PgHTravailN2, PgHTravailN3, PgHTravailN4, PgHGroupePaie: boolean;
//    PGHMSLIB1, PGHANLIB1, PGHMSLIB2, PGHANLIB2, PGHMSLIB3, PGHANLIB3, PGHMSLIB4, PGHANLIB4: boolean;
//    PGHDATE1, PGHDATE2, PGHDATE3, PGHDATE4, PGHBOOL1, PGHBOOL2, PGHBOOL3, PGHBOOL4: boolean;
//    PGHPCMB1, PGHPCMB2, PGHPCMB3, PGHPCMB4, PgProfil, PgPeriodBul, PgProfilRem: boolean;
//    PGHCONFIRMATION, PgHHORAIREMOIS, PgHHORHEBDO, PgHHORANNUEL, PgHTAUXHORAIRE: boolean;
    PgHistorisation, PgTypeEcriture : boolean;
//    PGHDADSPROF, PGHDADSCAT, PGHTAUXPARTIEL, PGHCONDEMPLOI : boolean;
    PGDADSHISTO, PGMODIFCOEFF: boolean;
    PGJourHeure: string; //edition calendrier en jour ou en heure
    PgTypSalLib1, PgTypSalLib2, PgTypSalLib3, PgTypSalLib4, PgTypSalLib5: boolean;
    PgEditOrg1, PgEditOrg2, PgEditOrg3, PgEditOrg4, PgEditCodeStat: Boolean; //edition des codes org
    // PT10 : 11/06/2002 : V582 PH Ajout champs gestion des primes eAGL
    PgGestionDesPrimes: BooLean;
    PgLibSaisPrim: string;
    PgPrimIntegCompl, PgPrimMoisPrec: Boolean;
    PgPrimAffichSal1, PgPrimAffichSal2, PgPrimAffichSal3, PgPrimAffichSal4, PgPrimAffichSal5: Boolean;
    // PT11 : 02/07/2002 : V582 PH Ajout champs autorisation de controle de creation enreg
    //PgEAI,  ggg
    PGRecupAbsHisto: Boolean;
    PgRttAcquis, PgRttPris: string; //PT12 Ajout Champ Rtt
    PGFNbFraisLibre, NBFormationLibre: Integer; //PT13 FORMATION
    FormationLibre1, FormationLibre2, FormationLibre3, FormationLibre4, FormationLibre5, FormationLibre6, FormationLibre7, FormationLibre8: string;
    PGTypeCodeStage, PGForMethodeCalc, PGForValoSalaire: string; //PT20
    PGForMethodeCalcPrev, PGForValoSalairePrev : string; //PT46
    PGForGestPlafByPop, PGForGestFraisByPop : Boolean; //PT50
    PGStageFormAuto: Boolean; //PT20
    PgBSAbs1, PgBSAbs2, PgDatePrepAuto: string; // Bilan social simplifié = code absence
    PGChampBudget: string; // liste des champs gérés dans le référentiel métier
    PGEcabBaseDeporte, PGEcabMonoBase, PGEcabHierarchie, PgEcabGestionMail, PgEcabFractionnement: Boolean;
    PGEcabValiditeMail, PGEcabSaisieRemp: Integer; //PT21 paramètres Econgés
    PGECabDateIntegration, PGETempDateVal : TDateTime;  { PT43 }
    PgLienRessource: Boolean;
    PgTypeAffectRes: string;
    PgNbCarNomRes, PgNbCarPrenRes: integer;
    PGTicketRestau, PGFactEtabl: boolean; // PT24  PT59	
    PGCodeClient: string; // PT24
    PGCodeRattach: string; // PT24
    PGSaisieArret: Boolean; //PT25
    PGForOPCAInterne: string; // PT26
    PGForPrevisionnel, PGForValidPrev, PGForValidSession, PGForGestionOPCA: Boolean; // PT26
    PGPCS2003: Boolean; // PT27
    PGGestIJSS: boolean; // PT28
    PGCritMaintien: string; // PT28
    PGBILANSOCIALDET, PGGESTIONCARRIERE, PGPARTICIPINT, PGMODIFLIGNEIMP: Boolean; //PT29
    PGMaintien, PGGESTFRP, PGSCORINGFORM, PGGESTACPT, PGCPTALIBAUXI : Boolean; // PT30 PT35
    PGRUBFRP, PGDECOMP4, PGDECOMP6: string;
    PGArrondiEnvers : Double; // PT34
    PGHistoMaintien : String; // PT36
    PGRESPONSNONSAL : Boolean; // PT37 Gestion des responsables non salariés
    PGAbsenceCheval : Boolean;  { PT38 }
    PGtauxCharges   : Double; // PT39
    PGChr6Semaine   : Boolean;  //PT40
    PGIDRELT        : String;   // PT42
    PGRepPublic     : String; { PT43 }
    PGIntegAutoRub  : Boolean; // PT44
    PgSeriaMasseSalariale : Boolean;
    PgSeriaPresence : Boolean;  { PT45 }
    PGPopulValides  : String; //PT48
    PGSeriaIDR      : Boolean;  // PT49
    PGModulepresence : Boolean; // PT52
    PGContratassu   : Boolean; // PT57
    ModeFiche		: Boolean; //PT53
    FicheOuverte    : Boolean; //PT54
    BasesPartage    : Array [0..9] of String; //PT55
    PGForMailAdr    : String; //PT56
    PGFactoDucs     : Boolean;  // PT58
  end;

{$IFNDEF EAGLSERVER}
var
  VH_Paie: LaVariablePaie;
{$ENDIF !EAGLSERVER}

{$IFDEF EAGLSERVER}
function VH_Paie: LaVariablePaie;
{$ENDIF EAGLSERVER}

procedure InitLaVariablePaie;
procedure ChargeParamsPaie;
function ReChargeParamsPaie(C: TControl): boolean;
function ExitParamsPaieControl(C: TControl): boolean;
implementation
{$IFNDEF TOXPAIE}
uses Ent1;
{$ENDIF}

procedure InitLaVariablePaie;
begin
  {$IFNDEF EAGLSERVER}
    VH_Paie := LaVariablePaie.Create;
  {$ENDIF !EAGLSERVER}
end;

procedure ChargeParamsPaie;
begin
  VH_Paie.PGRESPONSABLES := (GetParamSoc('SO_PGRESPONSABLES')); //PT7
  VH_Paie.PGMsa := (GetParamSoc('SO_PGMSA'));
  VH_Paie.PGIntermittents := (GetParamSoc('SO_PGINTERMITTENTS'));
  VH_Paie.PGBTP := (GetParamSoc('SO_PGBTP')); //PT8
  VH_Paie.PGCodeInterim := (GetParamSoc('SO_PGCODEINTERIM')); //PT9
  VH_Paie.PGInterimaires := (GetParamSoc('SO_PGINTERIMAIRES')); //PT9
  VH_Paie.PGNbreStatOrg := getparamsoc('SO_PGNBRESTATORG'); // Gestion des Stats de la PAIE ET GRH
  if VH_Paie.PGNbreStatOrg > 4 then VH_Paie.PGNbreStatOrg := 4;
  VH_Paie.PGLibCodeStat := (GetParamSoc('SO_PGLIBCODESTAT'));
  VH_Paie.PGLibelleOrgStat1 := (GetParamSoc('SO_PGLIBORGSTAT1'));
  VH_Paie.PGLibelleOrgStat2 := (GetParamSoc('SO_PGLIBORGSTAT2'));
  VH_Paie.PGLibelleOrgStat3 := (GetParamSoc('SO_PGLIBORGSTAT3'));
  VH_Paie.PGLibelleOrgStat4 := (GetParamSoc('SO_PGLIBORGSTAT4'));
  VH_Paie.PGSaisieBulletin := (GetParamSoc('SO_PGSAISIEBULLETIN')); // Methode de presentation du bulletin : Simplifiee ou detaillee
  VH_Paie.PGProfilFnal := (GetParamSoc('SO_PGPROFILFNAL')); // Profil FNAL Dossier
  //mv VH_Paie.PGIncSalarie:=(GetParamSoc('SO_PGINCSALARIE')='X') ;
  VH_Paie.PGIncSalarie := (GetParamSoc('SO_PGINCSALARIE'));
  //mv VH_Paie.PGCalculBulletin:=(GetParamSoc('SO_PGCALCULBULLETIN')='X') ; // Calcul automatique de la paie ou sur action utilisateur
  VH_Paie.PGCalculBulletin := (GetParamSoc('SO_PGCALCULBULLETIN')); // Calcul automatique de la paie ou sur action utilisateur
  VH_Paie.PGDecalage := (GetParamSoc('SO_PGDECALAGE')); // Entreprise en décalage O/N
  VH_Paie.PGDecalagePetit := (GetParamSoc('SO_PGDECALAGEPETIT')); // Entreprise en petit décalage O/N
  //mv VH_Paie.PGDecalage:=(GetParamSoc('SO_PGDECALAGE')='X') ;  // Entreprise en décalage O/N
  //mv VH_Paie.PGAnalytique:=(GetParamSoc('SO_PGANALYTIQUE')='X') ; // Gestion de l'analytique
  VH_Paie.PGAnalytique := (GetParamSoc('SO_PGANALYTIQUE')); // Gestion de l'analytique
  //mv VH_Paie.PGCongesPayes:=(GetParamSoc('SO_PGCONGES')='X') ;    // Gestion des CP
  VH_Paie.PGCongesPayes := (GetParamSoc('SO_PGCONGES')); // Gestion des CP
  VH_Paie.PGMonnaieTenue := (GetParamSoc('SO_PGMONNAIETENUE')); // Monnaie de Tenue de la PAIE
  //mv VH_Paie.PGTenueEuro:=(GetParamSoc('SO_PGTENUEEURO')='X') ;   // Dossier tenue en EURO
  VH_Paie.PGTenueEuro := (GetParamSoc('SO_PGTENUEEURO')); // Dossier tenue en EURO
  VH_Paie.PGDateBasculEuro := (GetParamSoc('SO_PGDATEBASCULEURO')); // Date de basculement en EURO
{PT51
  VH_Paie.PGCheminRech := (GetParamSoc('SO_PGCHEMINRECH')); // gestion des chemins pour editions parser Certificats, ...
  VH_Paie.PGCheminSav := (GetParamSoc('SO_PGCHEMINSAV')); // chemins de sauvegarde des documents
  VH_Paie.PGCheminEagl := (GetParamSoc('SO_PGCHEMINEAGL'));
}
{$IFNDEF EAGLSERVER}
  VH_Paie.PGCheminRech := VerifieCheminPG (GetParamSoc('SO_PGCHEMINRECH')); // gestion des chemins pour editions parser Certificats, ...
  VH_Paie.PGCheminSav := VerifieCheminPG (GetParamSoc('SO_PGCHEMINSAV')); // chemins de sauvegarde des documents
  VH_Paie.PGCheminEagl := VerifieCheminPG (GetParamSoc('SO_PGCHEMINEAGL'));
{$ENDIF}
//FIN PT51
  VH_Paie.PGMethodHeures := (GetParamSoc('SO_PGMETHODHEURES'));
  VH_Paie.PGPhotoSal := (GetParamSoc('SO_PGPHOTOSAL')); // Gestion des Photos Salariés
  VH_Paie.PGCumul01 := (GetParamSoc('SO_PGCUMUL01')); // Cumuls imprimes sur edition de bulletin
  VH_Paie.PGCumul02 := (GetParamSoc('SO_PGCUMUL02'));
  VH_Paie.PGCumul03 := (GetParamSoc('SO_PGCUMUL03'));
  VH_Paie.PGCumul04 := (GetParamSoc('SO_PGCUMUL04'));
  VH_Paie.PGCumul05 := (GetParamSoc('SO_PGCUMUL05'));
  VH_Paie.PGCumul06 := (GetParamSoc('SO_PGCUMUL06'));
  VH_Paie.PGEnversNet := (GetParamSoc('SO_PGENVERSNET')); // Calcul de la paie envers sur le net à payer ou cout ou ...
  VH_Paie.PGContextePaie := (GetParamSoc('SO_PGCONTEXTPAIE')); // Indique quand le contexte de la paie cad ttes les TOB doivent être chargées pour le calcul du bulletin
  VH_Paie.PGIterationEnvers := (GetParamSoc('SO_PGMAXITERATENVERS')); // Nombre Maximal iyeration pour le calcul de la paie à l'envers
  if VH_Paie.PGIterationEnvers = 0 then VH_Paie.PGIterationEnvers := 50;
  if VH_Paie.PGIterationEnvers > 100 then VH_Paie.PGIterationEnvers := 100; // Correction des Valeurs pour éviter des recherches infinies
  VH_Paie.PGCritSaisRub := (GetParamSoc('SO_PGCRITSAISRUB'));
  VH_Paie.PGLibCg1 := (GetParamSoc('SO_PGLIBCG1')); //Libelle des Absences édités en bas de bulletin
  VH_Paie.PGLibCg2 := (GetParamSoc('SO_PGLIBCG2'));
  VH_Paie.PGLibCg3 := (GetParamSoc('SO_PGLIBCG3'));
  VH_Paie.PGLibCg4 := (GetParamSoc('SO_PGLIBCG4'));
  VH_Paie.PGCgAcq1 := (GetParamSoc('SO_PGCGACQ1')); //Type d'Absence à éditer en bas de bulletin
  VH_Paie.PGCgPris1 := (GetParamSoc('SO_PGCGPRIS1'));
  VH_Paie.PGCgAcq2 := (GetParamSoc('SO_PGCGACQ2'));
  VH_Paie.PGCgPris2 := (GetParamSoc('SO_PGCGPRIS2'));
  VH_Paie.PGCgAcq3 := (GetParamSoc('SO_PGCGACQ3'));
  VH_Paie.PGCgPris3 := (GetParamSoc('SO_PGCGPRIS3'));
  VH_Paie.PGCgAcq4 := (GetParamSoc('SO_PGCGACQ4'));
  VH_Paie.PGCgPris4 := (GetParamSoc('SO_PGCGPRIS4'));
  VH_Paie.PgLibDate1 := (GetParamSoc('SO_PGLIBDATE1')); //Champs Libre Date
  VH_Paie.PgLibDate2 := (GetParamSoc('SO_PGLIBDATE2'));
  VH_Paie.PgLibDate3 := (GetParamSoc('SO_PGLIBDATE3'));
  VH_Paie.PgLibDate4 := (GetParamSoc('SO_PGLIBDATE4'));
  VH_Paie.PgLibCombo1 := (GetParamSoc('SO_PGLIBCOMBO1')); //Champs Libre Tablette
  VH_Paie.PgLibCombo2 := (GetParamSoc('SO_PGLIBCOMBO2'));
  VH_Paie.PgLibCombo3 := (GetParamSoc('SO_PGLIBCOMBO3'));
  VH_Paie.PgLibCombo4 := (GetParamSoc('SO_PGLIBCOMBO4'));
  VH_Paie.PgLibCoche1 := (GetParamSoc('SO_PGLIBCOCHE1')); //Champs Libre CheckBox
  VH_Paie.PgLibCoche2 := (GetParamSoc('SO_PGLIBCOCHE2'));
  VH_Paie.PgLibCoche3 := (GetParamSoc('SO_PGLIBCOCHE3'));
  VH_Paie.PgLibCoche4 := (GetParamSoc('SO_PGLIBCOCHE4'));
//  VH_Paie.PgProfil := (GetParamSoc('SO_PGPROFIL'));
//  VH_Paie.PgPeriodBul := (GetParamSoc('SO_PGPERIODBUL'));
  //PT4 : 02/10/01 : V562 : PH Ajout historisation profil remunérations
//  VH_Paie.PgProfilRem := (GetParamSoc('SO_PGPROFILREM'));
  VH_Paie.PgNbDate := (GetParamSoc('SO_PGNBDATE')); //nbre de champs libre gérés
  VH_Paie.PgNbCoche := (GetParamSoc('SO_PGNBCOCHE'));
  VH_Paie.PgNbCombo := (GetParamSoc('SO_PGNBCOMBO'));
  VH_Paie.PgNbSalLib := (GetParamSoc('SO_PGNBSALLIB')); //Champs Salaire
  VH_Paie.PgSalLib1 := (GetParamSoc('SO_PGSALLIB1'));
  VH_Paie.PgSalLib2 := (GetParamSoc('SO_PGSALLIB2'));
  VH_Paie.PgSalLib3 := (GetParamSoc('SO_PGSAlLIB3'));
  VH_Paie.PgSalLib4 := (GetParamSoc('SO_PGSALLIB4'));
  VH_Paie.PgRedHeure := (GetParamSoc('SO_PGREDHEURE')); //Réductions
  VH_Paie.PGRedRem := (GetParamSoc('SO_PGREDREM'));
  VH_Paie.PgTypRedRepas := (GetParamSoc('SO_PGTYPREDREPAS'));
  VH_Paie.PgRedRepas := (GetParamSoc('SO_PGREDREPAS'));
  VH_Paie.PgTypeNumSal := (GetParamSoc('SO_PGTYPENUMSAL'));
  //mv VH_Paie.PgTiersAuxiAuto :=(GetParamSoc('SO_PGTIERSAUXIAUTO')='X');
  VH_Paie.PgTiersAuxiAuto := (GetParamSoc('SO_PGTIERSAUXIAUTO'));
  VH_Paie.PGAx1Anal1 := (GetParamSoc('SO_PGAX1ANAL1'));
  VH_Paie.PGAx1Anal2 := (GetParamSoc('SO_PGAX1ANAL2'));
  VH_Paie.PGAx1Anal3 := (GetParamSoc('SO_PGAX1ANAL3'));
  VH_Paie.PGAx2Anal1 := (GetParamSoc('SO_PGAX2ANAL1'));
  VH_Paie.PGAx2Anal2 := (GetParamSoc('SO_PGAX2ANAL2'));
  VH_Paie.PGAx2Anal3 := (GetParamSoc('SO_PGAX2ANAL3'));
  VH_Paie.PGAx3Anal1 := (GetParamSoc('SO_PGAX3ANAL1'));
  VH_Paie.PGAx3Anal2 := (GetParamSoc('SO_PGAX3ANAL2'));
  VH_Paie.PGAx3Anal3 := (GetParamSoc('SO_PGAX3ANAL3'));
  VH_Paie.PGAx4Anal1 := (GetParamSoc('SO_PGAX4ANAL1'));
  VH_Paie.PGAx4Anal2 := (GetParamSoc('SO_PGAX4ANAL2'));
  VH_Paie.PGAx4Anal3 := (GetParamSoc('SO_PGAX4ANAL3'));
  VH_Paie.PGAx5Anal1 := (GetParamSoc('SO_PGAX5ANAL1'));
  VH_Paie.PGAx5Anal2 := (GetParamSoc('SO_PGAX5ANAL2'));
  VH_Paie.PGAx5Anal3 := (GetParamSoc('SO_PGAX5ANAL3'));
  VH_Paie.PGCptNetAPayer := (GetParamSoc('SO_PGCPTNETAPAYER'));
  VH_Paie.PGPreselect0 := (GetParamSoc('SO_PGPRESELECT0'));
  VH_Paie.PGPreselect1 := (GetParamSoc('SO_PGPRESELECT1'));
  VH_Paie.PGPreselect2 := (GetParamSoc('SO_PGPRESELECT2'));
  VH_Paie.PGPreselect3 := (GetParamSoc('SO_PGPRESELECT3'));
  VH_Paie.PGPre4select1 := (GetParamSoc('SO_PGPRE4SELECT1'));
  VH_Paie.PGPre4select2 := (GetParamSoc('SO_PGPRE4SELECT2'));
  VH_Paie.PGPre4select3 := (GetParamSoc('SO_PGPRE4SELECT3'));
  VH_Paie.PGModeleEcr := (GetParamSoc('SO_PGMODELEECR'));
  VH_Paie.PGJournalPaie := (GetParamSoc('SO_PGJOURNALPAIE'));
  //mv VH_Paie.PGSectionAnal :=(GetParamSoc('SO_PGSECTIONANAL')='X');
  VH_Paie.PGSectionAnal := (GetParamSoc('SO_PGSECTIONANAL'));
  VH_Paie.PGLongRacine := (GetParamSoc('SO_PGLONGRACINE')); //  Nbre de caractéres de la racine classe 6
  VH_Paie.PGDecompoCl4 := (GetParamSoc('SO_PGDECOMPOCL4')); // Decomposition des comptes de la classe 4
  VH_Paie.PGDecompoCl6 := (GetParamSoc('SO_PGDECOMPOCL6')); // Decomposition des comptes de la classe 6
  VH_Paie.PGNbreRac1 := (GetParamSoc('SO_PGNBRERAC1')); // Nbre de caracteres composition compte classe 6
  if VH_Paie.PGNbreRac1 < 0 then VH_Paie.PGNbreRac1 := 0;
  if VH_Paie.PGNbreRac1 > 3 then VH_Paie.PGNbreRac1 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PGNbreRac2 := (GetParamSoc('SO_PGNBRERAC2'));
  if VH_Paie.PGNbreRac2 < 0 then VH_Paie.PGNbreRac2 := 0;
  if VH_Paie.PGNbreRac2 > 3 then VH_Paie.PGNbreRac2 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PGNbreRac3 := (GetParamSoc('SO_PGNBRERAC3'));
  if VH_Paie.PGNbreRac3 < 0 then VH_Paie.PGNbreRac3 := 0;
  if VH_Paie.PGNbreRac3 > 3 then VH_Paie.PGNbreRac3 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PGNbre4Rac1 := (GetParamSoc('SO_PGNBRE4RAC1')); // Nbre de caracteres composition compte classe 6
  if VH_Paie.PGNbre4Rac1 < 0 then VH_Paie.PGNbre4Rac1 := 0;
  if VH_Paie.PGNbre4Rac1 > 3 then VH_Paie.PGNbre4Rac1 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PGNbre4Rac2 := (GetParamSoc('SO_PGNBRE4RAC2'));
  if VH_Paie.PGNbre4Rac2 < 0 then VH_Paie.PGNbre4Rac2 := 0;
  if VH_Paie.PGNbre4Rac2 > 3 then VH_Paie.PGNbre4Rac2 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PGNbre4Rac3 := (GetParamSoc('SO_PGNBRE4RAC3'));
  if VH_Paie.PGNbre4Rac3 < 0 then VH_Paie.PGNbre4Rac3 := 0;
  if VH_Paie.PGNbre4Rac3 > 3 then VH_Paie.PGNbre4Rac3 := 3; // Maxi 3 caractères car une combo
  VH_Paie.PgRubAcompte := (GetParamSoc('SO_PGRUBACOMPTE'));
  //VH_Paie.PgEditCodeOrg:=(GetParamSoc('SO_PGEDITCODEORG')); //DEB PT-1
  VH_Paie.PgEditOrg1 := (GetParamSoc('SO_PGEDITORG1')); //Edition sur bulletion code org
  VH_Paie.PgEditOrg2 := (GetParamSoc('SO_PGEDITORG2'));
  VH_Paie.PgEditOrg3 := (GetParamSoc('SO_PGEDITORG3'));
  VH_Paie.PgEditOrg4 := (GetParamSoc('SO_PGEDITORG4')); //FIN PT-1
  VH_Paie.PgEditCodeStat := (GetParamSoc('SO_PGEDITCODESTAT'));
  VH_Paie.PGLongRacine421 := (GetParamSoc('SO_PGLONGRACINE421')); //  Nbre de caractéres de la racine du compte  421
  VH_Paie.PGLongRacin4 := (GetParamSoc('SO_PGLONGRACIN4')); //  Nbre de caractéres de la racine classe 4
  VH_Paie.PgBulDefaut := (GetParamSoc('SO_PGBULDEFAUT')); //Modèle d'edition du bulletin par défaut
  VH_Paie.PgTypeCumulMois1 := GetParamSoc('SO_PGTYPECUMULMOIS1'); //Type Cumul du mois
  VH_Paie.PgTypeCumulMois2 := GetParamSoc('SO_PGTYPECUMULMOIS2');
  VH_Paie.PgTypeCumulMois3 := GetParamSoc('SO_PGTYPECUMULMOIS3');
  VH_Paie.PgCumulMois1 := GetParamSoc('SO_PGCUMULMOIS1'); //Cumul mois bas de bulletin
  VH_Paie.PgCumulMois2 := GetParamSoc('SO_PGCUMULMOIS2');
  VH_Paie.PgCumulMois3 := GetParamSoc('SO_PGCUMULMOIS3');
  VH_Paie.PgLibCumulMois1 := GetParamSoc('SO_PGLIBCUMULMOIS1'); //Libelle Cumul mois bas de bulletin
  VH_Paie.PgLibCumulMois2 := GetParamSoc('SO_PGLIBCUMULMOIS2');
  VH_Paie.PgLibCumulMois3 := GetParamSoc('SO_PGLIBCUMULMOIS3');
  VH_Paie.PGDADSHISTO := (GetParamSoc('SO_PGHDADSHISTO')); //PT15
  VH_Paie.PGMODIFCOEFF := (GetParamSoc('SO_PGCOEFFEVOSAL'));
  VH_Paie.PgHistorisation := (GetParamSoc('SO_PGHISTORISATION'));
//  VH_Paie.PgHConfirmation := (GetParamSoc('SO_PGHCONFIRMATION'));
{  VH_Paie.PgCodeEmploi := (GetParamSoc('SO_PGCODEEMPLOI'));
  VH_Paie.PgLibelleEmploi := (GetParamSoc('SO_PGLIBELLEEMPLOI'));
  VH_Paie.PgQualification := (GetParamSoc('SO_PGQUALIFICATION'));
  VH_Paie.PgCoefficient := (GetParamSoc('SO_PGCOEFFICIENT'));
  VH_Paie.PgIndice := (GetParamSoc('SO_PGINDICE'));
  VH_Paie.PgNiveau := (GetParamSoc('SO_PGNIVEAU'));
  VH_Paie.PgCodeStat := (GetParamSoc('SO_PGCODESTAT'));
  VH_Paie.PgHTravailN1 := (GetParamSoc('SO_PGHTRAVAILN1'));
  VH_Paie.PgHTravailN2 := (GetParamSoc('SO_PGHTRAVAILN2'));
  VH_Paie.PgHTravailN3 := (GetParamSoc('SO_PGHTRAVAILN3'));
  VH_Paie.PgHTravailN4 := (GetParamSoc('SO_PGHTRAVAILN4'));
  VH_Paie.PgHHORAIREMOIS := (GetParamSoc('SO_PGHHORAIREMOIS'));
  VH_Paie.PgHHORHEBDO := (GetParamSoc('SO_PGHHORHEBDO'));
  VH_Paie.PgHHORANNUEL := (GetParamSoc('SO_PGHHORANNUEL'));
  VH_Paie.PgHTAUXHORAIRE := (GetParamSoc('SO_PGHTAUXHORAIRE'));
  VH_Paie.PgHGroupePaie := false; //(GetParamSoc('SO_PGHGROUPEPAIE'));
  VH_Paie.PGHMSLIB1 := (GetParamSoc('SO_PGHMSLIB1'));
  VH_Paie.PGHANLIB2 := (GetParamSoc('SO_PGHANLIB1'));
  VH_Paie.PGHMSLIB2 := (GetParamSoc('SO_PGHMSLIB2'));
  VH_Paie.PGHANLIB2 := (GetParamSoc('SO_PGHANLIB2'));
  VH_Paie.PGHMSLIB3 := (GetParamSoc('SO_PGHMSLIB3'));
  VH_Paie.PGHANLIB3 := (GetParamSoc('SO_PGHANLIB3'));
  VH_Paie.PGHMSLIB4 := (GetParamSoc('SO_PGHMSLIB4'));
  VH_Paie.PGHANLIB4 := (GetParamSoc('SO_PGHANLIB4'));
  VH_Paie.PGHDATE1 := (GetParamSoc('SO_PGHDATE1'));
  VH_Paie.PGHDATE2 := (GetParamSoc('SO_PGHDATE2'));
  VH_Paie.PGHDATE3 := (GetParamSoc('SO_PGHDATE3'));
  VH_Paie.PGHDATE4 := (GetParamSoc('SO_PGHDATE4'));
  VH_Paie.PGHBOOL1 := (GetParamSoc('SO_PGHBOOL1'));
  VH_Paie.PGHBOOL2 := (GetParamSoc('SO_PGHBOOL2'));
  VH_Paie.PGHBOOL3 := (GetParamSoc('SO_PGHBOOL3'));
  VH_Paie.PGHBOOL4 := (GetParamSoc('SO_PGHBOOL4'));
  VH_Paie.PGHPCMB1 := (GetParamSoc('SO_PGHPCMB1'));
  VH_Paie.PGHPCMB2 := (GetParamSoc('SO_PGHPCMB2'));
  VH_Paie.PGHPCMB3 := (GetParamSoc('SO_PGHPCMB3'));
  VH_Paie.PGHPCMB4 := (GetParamSoc('SO_PGHPCMB4'));
  //PT2
  VH_Paie.PGHDADSPROF := (GetParamSoc('SO_PGHDADSPROF'));
  VH_Paie.PGHDADSCAT := (GetParamSoc('SO_PGHDADSCAT'));
  VH_Paie.PGHTAUXPARTIEL := (GetParamSoc('SO_PGHDADSPARTIEL'));
  VH_Paie.PGHCONDEMPLOI := (GetParamSoc('SO_PGHCONDEMPLOI')); //PT5
  }
  VH_Paie.PGIntegODPaie := (GetParamSoc('SO_PGINTEGODPAIE')); // methode integration des OD de paie
  VH_Paie.PGJourHeure := (GetParamSoc('SO_PGJOURHEURE')); //edition calendrier en jour ou en heure
  VH_Paie.PgVentilMulEtab := (GetParamSoc('SO_PGVENTILMULETAB')); // traitement muti etablissement génération comptable
  VH_Paie.PgCreationSection := (GetParamSoc('SO_PGCREATIONSECTION')); // Autorisation de création des sections par axe dans la compta
  VH_Paie.PgTypeEcriture := (GetParamSoc('SO_PGTYPEECRITURE')); // Si simulation alors = X
  VH_Paie.PgTypSalLib1 := (GetParamSoc('SO_PGTYPSALLIB1')); // Si elemnts variables fiches salaries sont des montants = X
  VH_Paie.PgTypSalLib2 := (GetParamSoc('SO_PGTYPSALLIB2')); // Si elemnts variables fiches salaries sont des montants = X
  VH_Paie.PgTypSalLib3 := (GetParamSoc('SO_PGTYPSALLIB3')); // Si elemnts variables fiches salaries sont des montants = X
  VH_Paie.PgTypSalLib4 := (GetParamSoc('SO_PGTYPSALLIB4')); // Si elemnts variables fiches salaries sont des montants = X
  VH_Paie.PgTypSalLib5 := (GetParamSoc('SO_PGTYPSALLIB5')); // Si elemnts variables fiches salaries sont des montants = X
  // PT10 : 11/06/2002 : V582 PH Ajout champs gestion des primes eAGL
  VH_Paie.PGGestionDesPrimes := (GetParamSoc('SO_PGGESTIONDESPRIMES'));
  VH_Paie.PGLibSaisPrim := (GetParamSoc('SO_PGLIBSAISPRIM'));
  VH_Paie.PgPrimIntegCompl := (GetParamSoc('SO_PGPRIMINTEGCOMPL'));
  VH_Paie.PgPrimAffichSal1 := (GetParamSoc('SO_PGPRIMAFFICHSAL1'));
  VH_Paie.PgPrimAffichSal2 := (GetParamSoc('SO_PGPRIMAFFICHSAL2'));
  VH_Paie.PgPrimAffichSal3 := (GetParamSoc('SO_PGPRIMAFFICHSAL3'));
  VH_Paie.PgPrimAffichSal4 := (GetParamSoc('SO_PGPRIMAFFICHSAL4'));
  // PT16  : 06/09/2002 PH V585 Ajout champ pour 5eme zone element de salaire
  VH_Paie.PgPrimAffichSal5 := (GetParamSoc('SO_PGPRIMAFFICHSAL5'));
  VH_Paie.PgSalLib5 := (GetParamSoc('SO_PGSALLIB5'));
  VH_Paie.PgTypSalLib5 := (GetParamSoc('SO_PGTYPSALLIB5')); // Si elemnts variables fiches salaries sont des montants = X

  VH_Paie.PgPrimMoisPrec := (GetParamSoc('SO_PGPRIMMOISPREC'));
  // PT11 : 02/07/2002 : V582 PH Ajout champs autorisation de controle de creation enreg
{ggg
  VH_Paie.PgEAI := (GetParamSoc('SO_PGEAI'));
}
  //PT12 Ajout Champ Rtt
  VH_Paie.PgRttAcquis := (GetParamSoc('SO_PGRTTACQUIS'));
  VH_Paie.PgRttPris := (GetParamSoc('SO_PGRTTPRIS'));
  //PT13 FORMATION
  VH_Paie.NBFormationLibre := (GetParamSoc('SO_PGFNBFORLIBRE'));
  VH_Paie.FormationLibre1 := (GetParamSoc('SO_PGFFORLIBRE1'));
  VH_Paie.FormationLibre2 := (GetParamSoc('SO_PGFFORLIBRE2'));
  VH_Paie.FormationLibre3 := (GetParamSoc('SO_PGFFORLIBRE3'));
  VH_Paie.FormationLibre4 := (GetParamSoc('SO_PGFFORLIBRE4'));
  VH_Paie.FormationLibre5 := (GetParamSoc('SO_PGFFORLIBRE5'));
  VH_Paie.FormationLibre6 := (GetParamSoc('SO_PGFFORLIBRE6'));
  VH_Paie.FormationLibre7 := (GetParamSoc('SO_PGFFORLIBRE7'));
  VH_Paie.FormationLibre8 := (GetParamSoc('SO_PGFFORLIBRE8'));
  VH_Paie.PGFNbFraisLibre := (GetParamSoc('SO_PGFNBFRAISLIBRE'));
  //PT20
  VH_Paie.PGTypeCodeStage := (GetParamSoc('SO_PGTYPECODESTAGE'));
  VH_Paie.PGStageFormAuto := (GetParamSoc('SO_PGSTAGEAUTO'));
  VH_Paie.PGForMethodeCalc := (GetParamSoc('SO_PGFORMETHODECALC'));
  VH_Paie.PGForValoSalaire := (GetParamSoc('SO_PGFORVALOSALAIRE'));
  //PT46
  VH_Paie.PGForMethodeCalcPrev := (GetParamSoc('SO_PGFORMETHODECALCPREV'));
  VH_Paie.PGForValoSalairePrev := (GetParamSoc('SO_PGFORVALOSALAIREPREV'));
  //PT50
  VH_Paie.PGForGestPlafByPop := (GetParamSoc('SO_PGFORPOPPLAF'));
  VH_Paie.PGForGestFraisByPop := (GetParamSoc('SO_PGFORPOPFRAIS'));
  // PT14 : 30/08/2002   V585 Ph Ajout champ bilan social simplifié
  VH_Paie.PgBSAbs1 := (GetParamSoc('SO_PGBSABS1'));
  VH_Paie.PgBSAbs2 := (GetParamSoc('SO_PGBSABS2'));
  // PT17  : 23/09/2002 PH V585 Ajout champ méthode préférence date de la préparation automatique
  VH_Paie.PgDatePrepAuto := (GetParamSoc('SO_PGDATEPREPAUTO'));
  // PT18  : 26/09/2002 PH V585 Ajout champ liste des champs gérés dans le référentiel métier
  VH_Paie.PGChampBudget := (GetParamSoc('SO_PGCHAMPBUDGET'));
  VH_Paie.PGRecupAbsHisto := GetParamSoc('SO_PGRECUPABSHISTO');
  VH_Paie.PGEcabMonoBase := GetParamSoc('SO_PGECABMONOBASE'); //DEB PT21 Econges
  if (NomHalley = 'eCABS5') and (VH_Paie.PGEcabMonoBase = False) then
    VH_Paie.PgEcabBaseDeporte := True
  else
    VH_Paie.PGEcabBaseDeporte := GetParamSoc('SO_PGECABBASEDEPORTE');
  VH_Paie.PGEcabHierarchie := GetParamSoc('SO_PGECABHIERARCHIE');
  VH_Paie.PgEcabGestionMail := GetParamSoc('SO_PGECABGESTIONMAIL');
  VH_Paie.PGEcabValiditeMail := GetParamSoc('SO_PGECABVALIDITEMAIL');
  VH_Paie.PgEcabFractionnement := GetParamSoc('SO_PGECABFRACTIONNEMENT');
  VH_Paie.PGECabDateIntegration := GetParamSoc('SO_PGECABDATEINTEGRATION');
  VH_Paie.PGEcabSaisieRemp := GetParamSoc('SO_PGECABSAISIEREMP'); //FIN PT21 Econges

  VH_Paie.PgLienRessource := GetParamSoc('SO_PGLIENRESSOURCE'); //DEB PT23
  VH_Paie.PgTypeAffectRes := GetParamSoc('SO_PGTYPEAFFECTRES');
  VH_Paie.PgNbCarNomRes := GetParamSoc('SO_PGNBCARNOMRES');
  VH_Paie.PgNbCarPrenRes := GetParamSoc('SO_PGNBCARPRENRES'); //FIN PT23

  VH_Paie.PGTicketRestau := GetParamSocSecur('SO_PGTICKETRESTAU', FALSE); // PT24
  VH_Paie.PGCodeClient := GetParamSocSecur('SO_PGCODECLIENT',''); // PT24
  VH_Paie.PGCodeRattach := GetParamSocSecur('SO_PGCODERATTACH',''); // PT24
  VH_paie.PGFactEtabl := GetparamsocSecur('SO_PGFACTETABL', FALSE); // pt59
  VH_Paie.PGSaisieArret := GetParamSoc('SO_PGSAISIEARRET'); // PT25
  VH_Paie.PGForOPCAInterne := GetParamSoc('SO_PGFORMOPCAINTERNE'); //DEBUT PT26
  VH_Paie.PGForPrevisionnel := GetParamSoc('SO_PGFORMPREVISIONNEL');
  VH_Paie.PGForValidPrev := GetParamSoc('SO_PGFORMVALIDPREV');
  VH_Paie.PGForValidSession := GetParamSoc('SO_PGFORMVALIDREA');
  VH_Paie.PGForGestionOPCA := GetParamSoc('SO_PGFORMOPCA'); // FIN PT26
  If GetParamSocSecur('SO_PGFORMAILCOPIE', False) = True Then //PT56
  	VH_PAIE.PGForMailAdr := GetParamSocSecur('SO_PGFORMAILADR', '') 
  Else
  	VH_PAIE.PGForMailAdr := '';
  //PT27  : 28/08/2003 PH V_42 Ajout paramsoc nomenclature PCS 2003
  VH_Paie.PGPCS2003 := GetParamSoc('SO_PGPCS2003'); // FIN PT27
  VH_Paie.PGGestIJSS := GetParamSocSecur('SO_PGGESTIJSS', FALSE); // PT28
  VH_Paie.PGCritMaintien := GetParamSocSecur('SO_PGCRITMAINTIEN',''); // PT28
  // PT29 28/06/2004 PH V_50 Ajout booléen pour Bilan social, gestion de carrière, participation
  VH_Paie.PGPARTICIPINT := GetParamSoc('SO_PGPARTICIPINT');
  VH_Paie.PGBILANSOCIALDET := GetParamSoc('SO_PGBILANSOCIALDET');
  VH_Paie.PGGESTIONCARRIERE := GetParamSoc('SO_PGGESTIONCARRIERE');
  VH_Paie.PGMODIFLIGNEIMP := GetParamSoc('SO_PGMODIFLIGNEIMP'); // authorisation modification ligne d'import
  VH_Paie.PGMaintien := GetParamSocSecur('SO_PGMAINTIEN', FALSE); // PT30
  VH_Paie.PGRUBFRP := GetParamSoc('SO_PGRUBFRP');
  VH_Paie.PGGESTFRP := GetParamSoc('SO_PGGESTFRP');
  VH_Paie.PGSCORINGFORM := GetParamSoc('SO_PGSCORINGFORM');
  VH_Paie.PGGESTACPT := GetParamSoc('SO_PGGESTACPT');
  VH_Paie.PGDECOMP4 := GetParamSoc('SO_PGDECOMP4');
  VH_Paie.PGDECOMP6 := GetParamSoc('SO_PGDECOMP6');
{$IFNDEF EABSENCES}
  VH_Paie.PGAbsence := (GetParamSocSecur ('SO_PGABSENCE', TRUE)); // Gestion des absences PT33
  VH_Paie.PGArrondiEnvers := (GetParamSocSecur('SO_PGARRONDIENVERS',2)); // Arrondi paie envers PT34
  VH_PAIE.PGCPTALIBAUXI := (GetParamSoc('SO_PGCPTALIBAUXI')); // PT35 récup du libellé compte de tiers
// d PT36
  VH_Paie.PGHistoMaintien := GetParamSocSecur('SO_PGHISTOMAINTIEN','001');     // PT36
  VH_PAIE.PGRESPONSNONSAL := (GetParamSocSecur('SO_PGRESPONSNONSAL',FALSE)); // PT37
// f PT36
{$ENDIF}
  VH_Paie.PGAbsenceCheval := (GetParamSocSecur ('SO_PGABSENCECHEVAL', False)); // Autoriser les absences à cheval { PT38 }
  VH_Paie.PGTauxCharges := (GetParamSocSecur ('SO_PGTAUXCHARGES', 0)); // taux de charges RTT { PT39 }

  VH_Paie.PGChr6Semaine := (GetParamSocSecur ('SO_PGCHR6SEMAINE', False)); //PT40
  VH_Paie.PGIDRELT := (GetParamSocSecur ('SO_PGIDRELT', '')); //PT42
  VH_Paie.PGPopulValides := (GetParamSocSecur ('SO_PGPOPULVALIDES', ''));  //PT48
  VH_Paie.PGModulePresence := (GetParamSocSecur ('SO_PGMODULEPRESENCE', False)); //PT52
  VH_Paie.PGContratassu := (GetParamSocSecur ('SO_PGCONTRATASSU', False)); //PT57


  if GetParamSocSecur('SO_IFDEFCEGID', FALSE) then
    Begin
    VH_Paie.PGETempDateVal :=  GetParamSocSecur('SO_PGETEMPDATEVAL', Idate1900);  { PT43 }
    VH_Paie.PGRepPublic    :=  GetParamSocSecur('SO_PGREPPUBLIC', 'C:');  { PT43 }
    End
  else
    Begin
    VH_Paie.PGETempDateVal := Idate1900;
    VH_Paie.PGRepPublic    :=  'C:';
    End;
  VH_Paie.PGIntegAutoRub := (GetParamSocSecur ('SO_PGINTEGAUTORUB', False)); //PT44
  VH_Paie.PGFactoDucs := GetParamSocSecur('SO_PGFACTODUCS',False);     // PT58
end;

function ReChargeParamsPaie(C: TControl): boolean;
begin
  VH_Paie.PGCongesPayes := getparamsoc('SO_PGCONGES');
  result := true;
end;

function ExitParamsPaieControl(C: TControl): boolean;
var
  Nam: string;
  init: word;
begin
  Result := True;
  Nam := C.Name;
  if Nam = 'SO_PGCONGES' then
    if (GetParamSoc('SO_PGCONGES') = '-') then
      if VH_Paie.PGCongesPayes then
      begin
        Init :=
          HShowMessage('1;Congés payés;Attention, la modification de la date de clôture va décaler tous les calculs congés payés.#13#10 Si vous souhaitez changer de période, clôturez celle ci pour ouvrir la suivante.#13#10 Etes vous sûr de vouloir continuer ?;Q;YN;N;N;', '', '');
        if Init = mrYes then exit
        else
        begin
          SetParamSoc('SO_PGCONGES', 'X');
          exit;
        end;
      end;
end;

{$IFDEF EAGLSERVER}
function VH_Paie: LaVariablePaie;
begin
//  BBI Web services
//  Result := LaVariablePaie(RegisterVHSession('VH_GC', LaVariablePaie))
  Result := LaVariablePaie(RegisterVHSession('VH_PAIE', LaVariablePaie))
//  BBI Fin Web services
end;
{$ENDIF EAGLSERVER}

end.



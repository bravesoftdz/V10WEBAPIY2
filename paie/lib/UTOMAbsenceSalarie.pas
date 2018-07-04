{***********UNITE*************************************************
Auteur  ...... : PAIE PGI : M. Vignal
Créé le ...... : 12/06/2001
Modifié le ... :   /  /
Description .. : Gestion des absences
Mots clefs ... : PAIE;ABSENCES
*****************************************************************

PT1-1 : 12/06/2001 Ph      Modif Boucle sur query remplacer par SUM
PT1-2 : 12/06/2001 Ph      Modif contenu OnChangeField modification Affichage
                           titre et des boutons à chaque champs ???
                           fait sur OnLoadRecord au lieu de OnChangeField
PT1-3 : 12/06/2001 Ph      Reecriture de l'affichage du recap CP et RTT avec les
                           fonctions AGL
PT3   : 13/06/2001 SB      Modif contenu OnchangeField sur PCN_TYPECONGE
                           Ajout contenu OnUpdateRecord
                           Le mvt de type reprise Pris CPA affecte PCN_CODEETAPE
                           à payé
                           le fait de redefinir sur un mvt PRI rend enabled
                           l'onglet..
                           On affect le codeetape et la date de paiement sur la
                           validation
PT4   : 15/06/2001 SB V547 Mode Debug:En debug permet quelques zones accesibles
PT5   :    07/2001    V547 Visualisation de l'indemnité CP des mvts SLD
PT6-1 : 28/08/2001 SB V547 PCN_NBREMOIS positionner à 1, le 0 est interdit
PT6-2 : 28/08/2001 SB V547 Contrôle de modification de fiche sur le code
                           Test valeur avant SetField
PT7-1 : 24/10/2001 SB V562 Affectation de la tablette des abs. deportés
                           Maj du PCN_LIBELLE new proc. AffectLibelle
PT7-2 : 24/10/2001 SB V562 Autorise la saisie par anticipation
PT7-3 : 25/10/2001 SB V562 Raffraichissement et MAJ du titre de l'ecran
PT7-4 : 25/10/2001 SB V562 On force la validation resp et exportOK pour saisie
                           non eagl
PT7-5 : 26/10/2001 SB V562 Si modif date deb abs alors réaffect date fin abs si
                           inferieur
PT7-6 : 30/10/2001 SB V562 Controle de cohérence MAT/PAM sur même journée
PT7-7 : 30/10/2001 SB V562 Si jour=0 alors message erreur
PT7-8 : 31/10/2001 SB V562 Gestion de la modification de la periode d'absence
                           Jusqu'à présent on positionnait BDateEnter sur
                           l'evenement onenter des composants
                           Inconvenient : Les messages d'erreur apparaissait pls
                           fois sur validation
                           Now, on recupère la valeur en entrée on test si elle
                           a changé sur le fieldname.
                           si oui on recontrole la cohérence et la durée de
                           l'absence
                           On ne gère plus d'evenement sur les composants
                           DATEDEB,DATEFIN,MATIN,APRESMIDI pour savoir si la
                           periode a été modifié
PT7-9 : 31/10/2001 SB V562 Control si typeconge renseigné
PT7-10: 31/10/2001 SB V562 Champvisible champEnabled
PT7-11: 31/10/2001 SB V562 Reconstruction des messages d''erreur..
                           Utilisation de PGIBox au lieu de LastErrorMsg,d'un
                           tableau de message et d'une variable globale Error
                           En validation LastError recupère la valeur de Error
PT7-12: 31/10/2001 SB V562 Suppression de LAVERTISSEMENT remplacé par PGENERAL
PT7-13: 31/10/2001 SB V562 Ajout du restant RTT
PT7-14: 31/10/2001 SB V562 Ajout codergrpt sur requête test si mvt existe déjà
PT8   : 26/09/2001 SB V562 Reconception du module, Mise en place du navigateur
PT9   : 12/11/2001 SB V562 Imputation du periode de paie pour les mvts AJU
PT10  : 14/11/2001 SB V562 Ajout du nom du responsable en saisie déportée
PT11  : 14/11/2001 SB V562 Ajout Cond. pour affectation de la periode CP
PT12  : 19/11/2001 SB V562 Contrôle saisi abs sur mvt déjà existant
PT13  : 22/11/2001 SB V563 Pour ORACLE, Le champ PCN_CODETAPE doit être
                           renseigné
PT14  : 30/11/2001 SB V563 Le REB est remplacé par le mvt AJU : controle de
                           saisie
PT15  : 29/01/2002 SB V571 Fiche de bug n°444 : Date de validité erronnée
PT16  : 22/02/2002 SB V571 Ne doit pas modifié le DataSet sur Onchange
PT17  : 26/02/2002 SB V571 L'adm. à acces à la validation des mvts en attente
PT17-2: 27/02/2002 SB V571 Le Libelle Recap sur fiche econges
PT18  : 27/03/2002 SB V571 L'adm. à acces à la validation des mvts en attente
PT19  : 23/04/2002 SB V571 Fiche de bug n°10000 : Tronquage du libllé absence
                           pour libellé paie limité à 35
PT20-1: 15/07/2002 SB V582 Intégration de la gestion des mails envoyés au
                           responsable
PT20-2: 15/07/2002 SB V582 Affectation Tablette motif saisi responsable et
                           utilisation de PCN_VALIDSALARIE pour toper saisie
                           resp ou sal pour création mvt resp.
PT20-3: 15/07/2002 SB V582 Intégration d'une notion legislatif
PT20-4: 15/07/2002 SB V582 Décompte calendaire pour Congé paternité
PT20-5: 15/07/2002 SB V582 Contrôle saisi absence pas + antérieur d'un mois
PT20-6: 15/07/2002 SB V582 Contrôle Assez restant erronné, modif des messages
PT21  : 19/07/2002 SB V582 Visu de l'indemnité pour les PRI ET SLD Cloturé
PT22  : 25/07/2002 SB V585 FQ n° 10177 Association des attestations et motifs
                           absences
PT23-1: 01/08/2002 SB V585 FQ n° 10199 Controle de vraisemblance sur suppression
PT24  : 16/09/2002 SB V585 FQ n° 10192 Chargement des données MOTIFABSENCE dans
                           une tob au lieu de générer des requêtes
PT24-1: 16/09/2002 SB V585 FQ n° 10192 Intégration du calcul calendrier civil
PT24-2: 16/09/2002 SB V585 FQ n° 10192 Distinction des mvts d'absence RTT par le
                           champ PMA_TYPERTT et non par le typeconge
PT24-3: 16/09/2002 SB V585 Zone PCN_SENSABS enabled en saisie des Absences pour
                           les Cp PRI
PT25  : 04/11/2002 SB V585 Intégration du calcul des compteurs Abs en cours
PT26-1: 05/12/2002 SB V591 Intégration du remplaçant obligatoire
PT26-2: 05/12/2002 SB V591 Intégration des nouveaux paramètres sociétés
PT26-3: 05/12/2002 SB V591 Ajout champ PRS_CUMRTTREST
PT27  : 11/12/2002 SB V591 FQ 10311 Reprise des dates d'absence pour la
                           génération de l'attestation
PT28  : 19/12/2002 SB V591 FQ 10382 Contrôle date entrée et suppression requête
                           inutile
PT29  : 07/01/2003 SB V591 réaffectation date validité si date de fin absence
                           modifié
PT30  : 08/01/2003 SB V591 Modif PT20-5 utilisation du paramètres sociétés
PT31  : 22/01/2003 SB V591 FQ 10437 Modif PT- 9 Modification du test en saisie
                           date AJP
PT32  : 28/01/2003 SB V591 Econges : Révision chargement récap et test solde
                           suffissant
PT33  : 19/02/2003 SB V595 FQ 10498 Faute orthographe
PT34  : 18/03/2003 SB V591 Refonte chargement motifabsence
PT35  : 26/03/2003 SB V42C FQ 10654 Controle dateentree, datesortie inopérant
PT36  : 08/08/2003 SB V42  Mvt AJP : affectation periode en cours pour
                           imputation bulletin
PT37-1: 15/09/2003 SB V42  FQ 10781 Saisie mvt PRI si gestion CP active
PT37-2: 15/09/2003 SB V42  FQ 10789 Contrôle code salarié saisie et existant
PT38  : 01/10/2003 SB V_42 FQ 10868 Affectation date absence selon mvt saisie
PT39  : 08/10/2003 SB V_42 Econges Spéc. CEGID Gestion des abences
PT40-1: 28/10/2003 SB V_42 Portage CWAS : Impression fiche
PT40-2: 28/10/2003 SB V_42 Portage CWAS : Bouton delete invisible
PT40-3: 30/10/2003 SB V_42 Calcul duree si affectation salarié
PT40-4: 31/10/2003 SB V_42 FQ 10943 Gestion des mails pour salarié
PT41  : 12/03/2004 SB V_50 FQ 11162 Encodage de la date de cloture erroné si fin
                           fevrier
PT42-1: 19/03/2004 SB V_50 FQ 11160 Restriction au mot de passe du jour à
                           l'accès en debug
PT42-2: 22/03/2004 SB V_50 FQ 11160 & FQ 11187 Refonte valeur par défaut sur
                           saisie mouvement d'AJU, AJP et reprise
PT42-3: 22/03/2004 SB V_50 FQ 11160 Contrôle saisie, suppression mvt AJU, AJP,
                           REP, CPA sur période fermée
                           Ajout Gestion Type imputation sur ajustement acquis
PT43-1: 31/03/2004 SB V_50 FQ 11209 Nombre de mois à zéro par défaut excepté
                           pour les mvts de REP
PT43-2: 09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                           salarié
PT44-1: 29/04/2004 SB V_50 FQ 11160 Les congés pris sont systématiquement sur la
                           période 0
PT44-2: 29/04/2004 SB V_50 FQ 11160 Mvt CP saisie sur période close topé 'C'
PT45  : 03/05/2004 SB V_50 FQ 11280 Modif champ TYPERTT => TYPEABS
PT46  : 12/05/2004 SB V_50 FQ 11264 Identification de saisie salarié
PT47-1: 12/05/2004 SB V_50 FQ 11259 Message clôture limité à la saisie CP
PT47-2: 14/05/2004 SB V_50 FQ 11278 Réaffectation date validité limité à la saisie CP
PT48  : 17/05/2004 SB V_50 FQ 11279 Econges : Motif saisissable par le salarié
PT49  : 09/06/2004 MF V_50 Traitement des IJSS : modification du traitement de
                            saisie, modification des absences. Ajout du
                            traitement de suivi des IJSS
PT50  : 06/07/2004 PH V_50 FQ 11404 Ecps5 : Mauvaise initialisation des dates en
                           creation
PT51  : 05/08/2004 VG V_50 La TOM est utilisée pour le traitement de 2 fiches :
                           MVTABSENCE et MVTCONGEPAYE - FQ N°11450
PT52  : 20/09/2004 PH V_50 FQ 11279 Econges : Motif saisissable par le salarié
                           uniquement en création sinon on perd tous les motifs
                           absences Responsable en modification
PT53  : 27/09/2004 PH V_50 FQ 11634 Erreur CWAS mvt Ajustement CP pris
PT54  : 17/11/2004 MF V_60 FQ 11725 correction contrôle dates entrée et sortie
                           du salarié.
PT55  : 14/01/2005 SB V_65 FQ 11820 En saisie absence, inutile de réaffecter les
                           dates si modif Typeconge
PT56-1: 10/02/2005 SB V_65 FQ 11902 Mise en place du navigateur en saisie
                           absence Paie
PT56-2: 10/02/2005 SB V_65 Ajout controle suivi IJSS si mvt congés payés
PT57  : 17/02/2005 MF V_60 le nb jours de carence IJSS dépend du type d'absence.
PT58  : 17/03/2005 MF V_60 FQ 12102 correction calcul nbre jours calendaires
                           (Pb CWAS)
PT59  : 17/03/2005 SB V_60 FQ 11208 Rectification contrôle saisie ajustement
PT60  : 22/04/2005 PH V_60 Filtrage des salariés confidentialité établissement
PT61  : 01/06/2005 MF V_60 FQ 12342 : le nb jours IJSS = 0 qd
                           nb jours calendaires-carence IJ < 0
PT62  : 02/06/2005 SB V_60 FQ 12338 Econges : Btn Delete non visible
PT63  : 05/07/2005 SB V_60 FQ 12424 Restriction message au mvt Congés Payés
PT64  : 25/07/2005 SB V_65 Ajout btn pour affichage du récapitulatif
PT65  : 09/09/2005 SB V_65 FQ 11313 Gestion lien planning/fiche pour planning
                           des absences de la paie
PT66  : 15/09/2005 SB V_65 Econges : recherche responsable table intérimaire
PT67  : 22/09/2005 PH V_65 Econges : recherche responsable table intérimaire
                           Suite en attendant mieux
PT68  : 26/09/2005 SB V_65 FQ 12349 Suppression contrôle existance des
                           fichiers .pdf
PT69  : 10/10/2005 SB V_65 FQ 12621 Controle mouvement CP limité au
                           typemvt ='CPA'
PT70  : 24/10/2005 SB V_65 ECONGES : Recherche du nom du l'utilisateur validant
PT71  : 17/11/2005 SB V_65 FQ 12689 ECONGES : Suppression du terme Outlook
PT72  : 17/11/2005 SB V_65 FQ 12700 Ajout contrôle date validité = date sortie
PT73  : 05/12/2005 SB V_65 EConges Ajout champ etat
PT74  : 06/12/2005 SB V_65 FQ 12692 Modification de la tablette des absences
PT75  : 23/01/2006 SB V_65 FQ 10866 Ajout clause predefini motif d'absence
PT76-1: 28/02/2006 SB V_65 FQ 10397 Gestion d'annulation
PT76-2: 28/02/2006 SB V_65 Pour recharger l'affichage des onglets en navigation
PT77  : 21/03/2006 MF V_65 Suivi des IJSS, mise au point CWAS affichage du
                           matricule du nom salarié et du libellé de l'absence
                           sur l'onglet IJSS
PT78  : 21/03/2006 MF V_65 IJSS :  alimentation nbre jours carence par défaut
PT79  : 07/04/2006 SB V_65 FQ 12941 Traitement nouveau mode de décompte motif
                           d'absence
PT80  : 13/04/2006 SB V_65 Traitement gestion maximum
PT81  : 13/04/2006 SB V_65 FQ 12122 Suppression des rglt IJSS
PT82  : 14/04/2006 SB V_65 FQ 12621 Controle mouvement CP limité au
                           typemvt ='CPA'
PT83  : 14/04/2006 SB V_65 FQ 11953 Paramsoc absence à cheval sur plusieurs mois
PT84  : 21/04/2006 SB V_65 FQ 13126 Ajout méthode pour contrôle affaire
                           existante
PT85  : 04/05/2006 SB V_65 FQ 12767 Ajout nom société mail econges
PT86  : 19/05/2006 SB V_65 FQ 13155 intégration des REP d'acs, et d'aca
PT87  : 23/05/2006 SB V_65 FQ 13193 Controle absence existente
PT88  : 01/06/2006 SB V_65 FQ 13198 Controle absence existente
PT89  : 09/06/2006 SB V_65 FQ 13202 Controle absence existente
PT90  : 12/07/2006 SB V_65 Raffraichissement de la date d'entree et de sortie
PT91  : 18/07/2006 MF V_70 FQ 13388 Correction Bdelete qui ne s'affiche pas
PT91-2: 20/07/2006 SB V_70 FQ 13388 Correction Bdelete qui ne s'affiche pas
PT92  : 17/10/2006 SB V_70 Refonte Fn pour utilisation portail => exporter vers
                           pgcalendrier
PT93  : 30/11/2006 SB V_70 FQ 13719 Anomalie affectation de l'établissement
PT94  : 25/01/2007 SB V_70 FQ 13173 Ajout événement annulation au journal
                           d'évènement
PT95  : 25/01/2007 GGU V_80 Gestion de la fiche de saisie des mouvements de
                           présence
PT96  : 21/05/2007 FC V_72 En vision SAV dans econges, champs inexistants
PT97  : 20/06/2007 FC V_72 FQ 14196 Accès à la saisie des absences pour le
                           salarié confidentiel si on saisit son matricule dans
                           les critères
PT98  : 05/07/2007 NA V_80 Ajout Gestion du GUID
PT99  : 25/07/2007 FLO V_80 Présence - Gestion du champs "Nb Heures de nuit"
PT100 : 02/08/2007 FLO V_80 Présence - Externalisation du contrôle des
                           mouvements de présence
PT101 : 06/08/2007 FLO V_80 Présence - Limitation des mouvements de présence à
                           la journée, Duplication
PT102 : 10/08/2007 FLO V_80 Présence - Recalcul automatique des compteurs lors
                           de la saisie d'un évènement
PT103 : 12/08/2007 PH V_80 FQ 14744 Mauvaise initialisation PCN_CODERGRPT en
                           création CP Pris
PT104   08/10/2007 MF  V_80 mise en place du traitement des jours de fractionnement
PT105 : 17/10/2007 VG V_80 Mise à jour du planning unifié
PT106 : 11/12/2007 FC V_810 Appel depuis fiche salarié des absences, on création,
                            on pouvait changer de n° salarié
PT107 29/12/2007 FC V_81 Concept accessibilité fiche salarié
PT108 26/12/2007 PH V_81 FQ 13173 Rajout du code motif absence dans libellé mouvement annulé du
                         journal des événements
PT109 : 06/10/2008 SJ FQ n°15289 Permettre la saisie de nouvelle absence sur des périodes antécedante à la date d'entrée (cas plusieurs contrats)
}
unit UTOMAbsenceSalarie;

interface
uses
   {$IFDEF VER150}
   Variants,
   {$ENDIF}
   StdCtrls,
   Controls,
   Classes,
   forms,
   sysutils,
   ComCtrls,
   HCtrls,
   HEnt1,
   HMsgBox,
   UTOM,
   UTOB,
   HTB97,
{$IFDEF EAGLCLIENT}
   UtileAGL,
   MaineAgl,
   eFiche,
{$ELSE}
   Fiche,
   HDB,
   Fe_Main,
   db,
{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
   EdtREtat,
   LicUtil,
{$ENDIF}
   AGLInit,
   MailOL,
   paramsoc,
   ULibEditionPaie;

type
  TOM_AbsenceSalarie = class(TOM)
    procedure OnArgument(stArgument: string); override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnClose; override;
    procedure OnAfterUpdateRecord; override; //PT25
    procedure OnAfterDeleteRecord; override; //PT102
  private
    GblTypeConge: string;
    etablissement, salarie, Origine, GblTypeMvt, Action, libordre, FicPrec, oldValidResp: string;
    Nom, Prenom, calendrier, standcalend, SalarieEnter: string;
    GblDateValidite: TDateTime;
    GblPeriodeCP: Integer;
    DTclot, {DDebPaieAju,DFinPaieAju,PT31} {DateModifRecap,PT30} MailDate: TdateTime; //PT26-2 ajout date
    // QMul : TQUERY;     // Query recuperee du mul
    Premier_passage, CPEtab, CpSal, Bok: boolean; // variable de travail { PT43-2 }
    DDAbs, DFAbs, EJoursPris, GblJours: double; { PT59 }
    DebDJ, FinDJ, Error, GblOrdre: integer;
    BDateEnter, FirstValid, OkPassage, OkMailEnvoie, AValider, IsMailRefOrAnu, PGAcces: boolean; { PT42-1 }
    ComplMessage, Resp, NomResp, MailResp, MailSal: string;
    Tob_MotifAbs, T_MotifAbs, Tob_Exercice,T_Recap{,Tob_Recapitulatif}: Tob; //PT24 //PT??
    DateEntree, DateSortie: TDateTime; //PT28
    Tob_AbsIJSS: TOB; // PT49
    IJSSaSolder: Boolean; // PT49
    GblModifDate, OnFerme, GblPasControl: Boolean; { PT55 } { PT56-1 } { PT59 }
    GblEnterTypeMvt, DerniereCreate, ParamDD, ParamDF: string; { PT56-1 }
    GblEnterTypeConge, GblEnterTypeImpute, GblEnterSensAbs: string; { PT59 }
    GblEnterEtatPostPaie : string; { PT94 }
    procedure enableChampOngletMvtCp;
    procedure enableChampOngletModeAgl;
    procedure InitialiseZonesCongesPayes;
    procedure ControleZonesOngletCongesPayes;
    function RechercheReprise(Validite: Tdatetime; TypeConge, TypeImpute : string): Tquery; { PT86 }
    procedure RechercheExerciceCp(Validite: tdatetime; var DTdeb, DtFin: tdatetime);
    procedure BindemniteClick(Sender: TObject);
    procedure ChargeLatob(Tobpris: tob);
    procedure TCalendrierClick(Sender: TObject);
    //     procedure NextOnClick(Sender: TObject);    Ancienne version
    //     procedure PrevOnClick(Sender: TObject);    Ancienne version
    //     procedure FirstOnClick(Sender: TObject);   Ancienne version
    //     procedure LastOnClick(Sender: TObject);    Ancienne version
    //     procedure BValiderOnClick(Sender: TObject);  Ancienne version
    //     procedure DateEnter(Sender: TObject);    PT- 7-8
    //     Function  mBouge(Button: TNavigateBtn) : boolean ; Ancienne version
    procedure AfficheTitre;
    procedure ValidRespEnter(Sender: TObject);
    procedure AffectLibelle; //PT- 7-1
    procedure OnExitSalarie(Sender: TObject); //PT- 8
    procedure OnEnterSalarie(Sender: TObject); //PT- 8
    procedure OnExitHeure(Sender: TObject); //PT95
    procedure OnExitDate(Sender:TObject);
    procedure BAttestClick(Sender: TObject); //PT22
    procedure ChargeTob_MotifAbsence;
    procedure AffectInfoRecap(Sal: string); //PT32  PT92
    procedure ImprimerClick(Sender: TObject); //PT40-1
    procedure MajIJSS; // PT49
    procedure TRecapClick(Sender: TObject);
    procedure GestionOngletAnnulation; { PT76-1 }
    procedure InitEntreeSortie; { PT90 }
    procedure ControleSaisiePresence; //PT100
    procedure Dupliquer (Sender : Tobject); //PT101
  end;

implementation

uses PgCongesPayes,
  PgOutils,
  P5util,
  EntPaie,
  PGOutilsEagl,
  PgOutils2,
  P5Def,
  PGCommun,
  PgCalendrier,
  PgPresence,
  HeureUtil,
  Ed_Tools,
  Menus,
  DateUtils,
  UTobDebug,
  PgPlanningUnifie;

procedure TOM_AbsenceSalarie.OnArgument(stArgument: string);
var
{$IFDEF EAGLCLIENT}
  Edit: THEdit;
  MType: THValComboBox;
{$ELSE}
  MType: THDBValComboBox;
  Edit: THDBEdit;
{$ENDIF}
  Arg, taction, St: string;
  BIndemnite, Tcalendrier, Btn, BAttest: TToolbarbutton97;
  StWhere : String;
  MenuPop : TPopUpMenu;
  i : Integer;
begin
  inherited;
  //PT34 chargement
  tob_motifabs := tob.create('tob_virtuelle', nil, -1);
  tob_motifabs.loaddetaildb('MOTIFABSENCE', '', 'PMA_MOTIFABSENCE', nil, False);
  ChargementTablette('PMA', '');  { PT74 On recharge les tablettes motif abssence }

  Arg := stArgument;
  salarie := Trim(ReadTokenPipe(Arg, ';'));
  if pos('!', Salarie) > 0 then
  begin
    st := salarie; // st := n°salarié ! nom salarié
    salarie := Trim(ReadTokenPipe(sT, '!'));
    Nom := ST;
  end;
  //DEB PT97
  StWhere := SQLConf('SALARIES');
  if StWhere <> '' then
    StWhere := ' AND ' + StWhere;
  if not existesql('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE= "' + salarie + '"' + StWhere) then
    Salarie := '';
  //FIN PT97
  Origine       := Trim(ReadTokenPipe(Arg, ';'));
  Etablissement := Trim(ReadTokenPipe(Arg, ';'));
  Action        := Trim(ReadTokenPipe(Arg, ';'));
  LibOrdre      := Trim(ReadTokenPipe(Arg, ';'));
  FicPrec       := Trim(ReadTokenPipe(Arg, ';'));
  tAction       := ReadtokenPipe(Action, '=');
  ParamDD       := Trim(ReadTokenPipe(Arg, ';'));
  ParamDF       := Trim(ReadTokenPipe(Arg, ';'));

  { DEB PT42-1 }
{$IFNDEF EAGLCLIENT}
  PGAcces := ((V_PGI.PassWord = CryptageSt(DayPass(Date))) and (V_PGI.debug));
{$ELSE}
  PGAcces := FAlse;
{$ENDIF}
  { FIN PT42-1 }
 //DEB PT- 7-8 Mise en commentaire
 //Suppression de la mise en commentaire voir ancienne version
 //FIN PT- 7-8 Mise en commentaire

  BIndemnite := TToolbarbutton97(getcontrol('BINDEMNITE'));
  if Bindemnite <> nil then Bindemnite.OnClick := BindemniteClick;

  Tcalendrier := TToolBarButton97(GetControl('BCALENDRIER'));
  if Tcalendrier <> nil then Tcalendrier.OnClick := TcalendrierClick;

  BAttest := TToolBarButton97(GetControl('BATTEST')); //PT22
  if BAttest <> nil then BAttest.OnClick := BAttestClick;

  SetControlVisible('BIMPRIMER', True); //DEB PT40-1
  Btn := TToolBarButton97(GetControl('BIMPRIMER'));
  if btn <> nil then Btn.OnClick := ImprimerClick; //FIN PT40-1

  Btn := TToolBarButton97(GetControl('BRECAP')); //PT64
  if btn <> nil then Btn.OnClick := TRecapClick; //PT64

  {  PT- 8 mise en commentaire
  //Suppression de la mise en commentaire voir ancienne version
     }

{$IFDEF EAGLCLIENT}
  MType := THValComboBox(getcontrol('PCN_VALIDRESP'));
{$ELSE}
  MType := THDBValComboBox(getcontrol('PCN_VALIDRESP'));
{$ENDIF}
  if Mtype <> nil then Mtype.onenter := ValidRespEnter;

  {DEB PT- 8}
{$IFDEF EAGLCLIENT}
  Edit := THedit(getcontrol('PCN_SALARIE'));
{$ELSE}
  Edit := THDBedit(getcontrol('PCN_SALARIE'));
{$ENDIF}
  if Edit <> nil then
  begin
    Edit.OnEnter := OnEnterSalarie;
    Edit.OnExit := OnExitSalarie;
  end;
  {FIN PT- 8}
 //PT51
  if (Origine = 'A') then
  begin
    // d PT49 IJSS
{$IFDEF CCS3}
    TTabSheet(GetControl('PIJSS')).TabVisible := False;
{$ENDIF}
    if VH_Paie.PGGestIJSS = False then
      TTabSheet(GetControl('PIJSS')).TabVisible := False;
    // f PT49 IJSS
  end;
  //FIN PT51


  if Origine = 'P' then  //Deb PT95
  begin
    if GetControl('HDEB') is THedit then
      (GetControl('HDEB') as THedit).OnExit := OnExitHeure;
    if GetControl('HFIN') is THedit then
      (GetControl('HFIN') as THedit).OnExit := OnExitHeure;
{$IFDEF EAGLCLIENT}
    if GetControl('PCN_DATEDEBUTABS') is THedit then
      (GetControl('PCN_DATEDEBUTABS') as THedit).OnExit := OnExitDate;
    if GetControl('PCN_DATEFINABS') is THedit then
      (GetControl('PCN_DATEFINABS') as THedit).OnExit := OnExitDate;
    (GetControl('PCN_NBHEURESNUIT') As THEdit).OnExit := OnExitHeure;//PT99
{$ELSE}
    if GetControl('PCN_DATEDEBUTABS') is THDBedit then
      (GetControl('PCN_DATEDEBUTABS') as THDBedit).OnExit := OnExitDate;
    if GetControl('PCN_DATEFINABS') is THDBedit then
      (GetControl('PCN_DATEFINABS') as THDBedit).OnExit := OnExitDate;
    (GetControl('PCN_NBHEURESNUIT') As THDBEdit).OnExit := OnExitHeure;//PT99
{$ENDIF}

    //PT101 - Début
    If Action <> 'CREATION' Then
    Begin
          SetControlVisible('BDUPLIQUER', True);

          // Association de l'action
          MenuPop := TPopUpMenu(GetControl('POPDUPLIQUER'));
          If MenuPop <> Nil then
          Begin
               For i := 0 to MenuPop.Items.Count - 1 do
                    MenuPop.Items[i].OnClick := Dupliquer;
          End;
    End;
    //PT101 - Fin
  end;  //Fin PT95

  //DEB PT107
  if Action='CONSULTATION' then
  begin
    SetControlProperty('BATTEST','enabled',false);
    SetControlProperty('BINDEMNITE','visible',false);
    SetControlProperty('BCALENDRIER','enabled',false);
    SetControlProperty('BValider','visible',false);
  end;
  //FIN PT107
end;

procedure TOM_AbsenceSalarie.ValidRespEnter(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
  MType: THValComboBox;
{$ELSE}
  MType: THDBValComboBox;
{$ENDIF}
begin
  if ORIGINE = 'P' then Exit; //PT95
{$IFDEF EAGLCLIENT}
  MType := THValComboBox(getcontrol('PCN_VALIDRESP'));
{$ELSE}
  MType := THDBValComboBox(getcontrol('PCN_VALIDRESP'));
{$ENDIF}
  if Mtype <> nil then
    OldValidResp := Mtype.value;
end;

procedure TOM_AbsenceSalarie.AfficheTitre;
begin
  if (ACTION = 'MODIFICATION') then
  begin
    if Origine = 'C' then
      TFFiche(Ecran).caption := TraduireMemoire('Congés payés du salarié :') + ' '+ salarie + ' ' + Nom + ' ' + Prenom { PT56-1 }
    else if Origine = 'A' then
      TFFiche(Ecran).caption := TraduireMemoire('Absence du salarié :'     ) + ' '+ salarie + ' ' + Nom + ' ' + Prenom; { PT56-1 }
  end;

  //else  PT- 7-3
  if ORIGINE = 'E' then //PT- 7-3 MAJ du titre de l'ecran
  begin
    if ACTION = 'CREATION' then TFFiche(Ecran).caption := TraduireMemoire('Saisie des absences du salarié')+ ' ' + Nom + ' ' + Prenom
    else TFFiche(Ecran).caption := TraduireMemoire('Gestion des absences du salarié') +' '+ Nom + ' ' + Prenom;
  end;

  if ORIGINE = 'P' then //PT95
    TFFiche(Ecran).caption := TraduireMemoire('Mouvements du salarié :'    ) + ' '+ salarie + ' ' + Nom + ' ' + Prenom;

  UpdateCaption(TFFiche(Ecran));
end;



procedure TOM_AbsenceSalarie.TCalendrierClick(Sender: TObject);
begin
  { PT- 8 Mise en commentaire remonter ds la procedure OnLoadRecord
  Q := Opensql('Select PSA_CALENDRIER, PSA_STANDCALEND,PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="'+ Salarie+'"', True);
  if not Q.eof then
     begin
     calendrier := Q.findfield('PSA_CALENDRIER').AsString;
     lib        := Q.findfield('PSA_LIBELLE').AsString;
     standcalend := Q.findfield('PSA_STANDCALEND').AsString;
     end
     else exit; ferme(Q);}
  if ((standcalend = 'ETB') or (standcalend = 'ETS')) then
    AgllanceFiche('AFF', 'HORAIRESTD', '', '',
      'TYPE:STD;CODE:' + Salarie + ';LIBELLE:' + Nom + ';STANDARD:' + Calendrier);
  if (standcalend = 'PER') then
    AgllanceFiche('AFF', 'HORAIRESTD', '', '',
      'TYPE:SAL;CODE:' + SALARIE + ';LIBELLE:' + Nom + ';STANDARD:' + CALENDRIER);
end;

procedure TOM_AbsenceSalarie.OndeleteRecord;
var
  Pris, Acquis, Restants, Base, MBase: Double; { PT42-3 }
  TypeAbs, Retour : String;
  DD, DF : TDateTime;
begin
  inherited;
  DDabs := 0;
  DFABS := 0;
  if Action = 'MODIFICATION' then
  begin
    DDAbs := getfield('PCN_DATEDEBUTABS');
    DFAbs := getfield('PCN_DATEFINABS');
  end;
  if (Origine <> 'P') then //Deb PT95
  begin
    DebDJ := 0;
    FinDJ := 1;
    if Action = 'MODIFICATION' then
    begin
      AffecteNodemj(getfield('PCN_DEBUTDJ'), DebDj);
      AffecteNodemj(getfield('PCN_FINDJ'), FinDj);
    end;
    //DEB PT23-1
    if (Origine <> 'E') and (PGAcces = False) then //non econges { PT42-1 }
      if (GetField('PCN_CODETAPE') <> '...') and ((Getfield('PCN_TYPECONGE') = 'PRI') or (Getfield('PCN_TYPEMVT') = 'CPA') or (Getfield('PCN_TYPECONGE') = 'AJU') or (Getfield('PCN_TYPECONGE') = 'AJP')) then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez pas supprimer ce mouvement.', Ecran.caption);
      end;
    { DEB PT42-3 }
    if (GetField('PCN_APAYES') > 0) and (GetField('PCN_SENSABS') = '+') and ((Getfield('PCN_TYPECONGE') = 'AJU') or (Getfield('PCN_TYPECONGE') = 'AJP')) then
    begin
      LastError := 1;
      PGIBox('Vous ne pouvez pas supprimer ce mouvement.#13#10' +
        'Une prise de congés est imputé sur ce dernier.', Ecran.caption);
    end;
    if (LastError = 0) and (Origine = 'C') and (GetField('PCN_TYPECONGE') <> 'PRI') and (GetField('PCN_PERIODECP') > 0) then
    begin
      AffichelibelleAcqPri((GetField('PCN_PERIODECP') - 1), Salarie, Idate1900, 0, Pris, Acquis, Restants, Base, MBase, false, False);
      if (Pris > 0) then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez pas supprimer ce mouvement pour cette période.#13#10' +
          'Vous avez entamé vos congés payés sur la période suivante.', Ecran.Caption);
      end;
    end;
    { FIN PT42-3 }
    //FIN PT23-1
    //PT25 Recalcul des compteurs en cours lors de la suppression
    if (LastError = 0) and (Origine = 'E') then
    begin //PT34
  { DEB PT92 }
      PgSupprRecapitulatif(Tob_MotifAbs,GetField('PCN_SALARIE'),GetField('PCN_TYPECONGE'),EJoursPris,Error,Retour);
      if Error <> 0 then
         Begin
         LastError := Error;
         PgiBox(RecupMessageErrorAbsence(Error), Ecran.caption);
         End;
       if Retour <> '' then TFFiche(Ecran).Retour := Retour else TFFiche(Ecran).Retour := 'SUPPR'; // pt95

     (* if (GetField('PCN_TYPECONGE') = 'PRI') then
      begin
        try
          BeginTrans;
          ExecuteSql('UPDATE RECAPSALARIES SET PRS_PRIATTENTE=PRS_PRIATTENTE-' + StrfPoint(EJoursPris) +
            ' WHERE PRS_SALARIE="' + GetField('PCN_SALARIE') + '"');
          CommitTrans;
        except
          Rollback;
          LastError := 1;
          PGIBox('Echec lors du calcul du récapitulatif.Suppression annulée', Ecran.caption);
        end;
      end
      else
      begin
        if assigned(T_MotifAbs) then //PT34
          if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'RTT') then { PT45 }
          begin
            try
              BeginTrans;
              ExecuteSql('UPDATE RECAPSALARIES SET PRS_RTTATTENTE=PRS_RTTATTENTE-' + StrfPoint(EJoursPris) +
                ' WHERE PRS_SALARIE="' + GetField('PCN_SALARIE') + '"');
              CommitTrans;
              TFFiche(Ecran).Retour := 'SUPPR';
            except
              Rollback;
              LastError := 1;
              PGIBox('Echec lors du calcul du récapitulatif.Suppression annulée', Ecran.caption);
            end;
          end;
      end; *)
  { FIN PT92 }
      //FIN PT34
    end;

    { DEB PT81 }
    If LastError = 0 then
      Begin
      if assigned(T_MotifAbs) then
        Begin
         TypeAbs := T_MotifAbs.GetValue('PMA_TYPEABS');
         DD := getfield('PCN_DATEDEBUTABS');
         DF := GetField('PCN_DATEFINABS');
         ExecuteSql('DELETE FROM REGLTIJSS '+
                    'WHERE PRI_SALARIE = "'+GetField('PCN_SALARIE')+'" '+
                    'AND PRI_TYPEABS = "'+TypeAbs+'" '+
                    'AND PRI_DATEINTEGR ="'+USDateTime(idate1900)+'" '+
                    'AND PRI_DATEDEBUTABS >= "'+USDateTime(DD)+'" '+
                    'AND PRI_DATEFINABS <="'+USDateTime(DF)+'" ');
         End;
      End;
    { FIN PT81 }
  end;
   If (Origine = 'P') and (LastError = 0)  then  TFFiche(Ecran).Retour := 'SUPPR'; //PT95

//PT105
DeletePGYPL (GetField ('PCN_GUID'), '');
//FIN PT105
end;

procedure TOM_AbsenceSalarie.OnNewRecord;
var
GUID : string;  // pt98
begin
  inherited;
  // Deb pt98
  GUid := AglGetGuid();
  setfield('PCN_GUID', GUID);
  // Fin pt98
  
  Action := 'CREATION';
  TFFiche(Ecran).FTypeAction := taCreat;
  OkPassage := True; //PT- 8
  GblModifDate := False; { PT56-1 }
  SetField('PCN_SALARIE', Salarie);
  SetField('PCN_DATEDEBUT', idate1900); //DEB PT- 9
  SetField('PCN_DATEFIN', idate1900);
  SetField('PCN_DATEVALIDITE', idate1900); //FIN PT- 9
  SetField('PCN_CONFIDENTIEL', 0);
  //DEB PT- 12 on affecte l'ordre du mvt en création et réaffect en validation
  if (pos('ORDREG', libordre) = 0) and (Salarie <> '') then
    if (Origine = 'P') then
      SetField('PCN_ORDRE', IncrementeSeqNoOrdre('PRE', Salarie))
    else
      SetField('PCN_ORDRE', IncrementeSeqNoOrdre('CPA', Salarie))
  else
  begin
    SetField('PCN_ORDRE', suivant);
    suivant := suivant + 1;
  end;
  //FIN PT- 12
  GblOrdre := GetField('PCN_ORDRE'); { PT56-1 }

  if (Origine = 'P') then //Deb PT95
  begin
    SetField('PCN_TYPEMVT', 'PRE');
    SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFPRESENCE');
//    SetField('PCN_TYPECONGE', '');
  end else begin          //Fin PT95
    SetField('PCN_TYPEMVT', 'CPA');
    SetField('PCN_DEBUTDJ', 'MAT');
    SetField('PCN_FINDJ', 'PAM');
    setfield('PCN_MVTDUPLIQUE', '-');
    SetField('PCN_MVTORIGINE', 'SAL'); { PT46 }
    SetField('PCN_TYPECONGE', 'PRI');
    SetField('PCN_ETATPOSTPAIE','VAL'); { PT76-1 }
    SetField('PCN_NBJIJSS',0); // PT91
    if (Origine = 'C') then
    begin
      SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMVTCONGESS');
      if Salarie <> '' then
        SetControlEnabled('PCN_SALARIE', False); {PT- 8}
    end;
    //DEB PT106
    if (Origine = 'A') and (Salarie <> '') then
        SetControlEnabled('PCN_SALARIE', False);
    //FIN PT106
  //     else     Ne sert à rien car c'est la tablette de la fiche
  //     if Origine = 'E' then SetControlProperty('PCN_TYPECONGE','Datatype','PGMABSENCED');
    if Origine = 'E' then
    begin
      setfield('PCN_SENSABS', '-');
      SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCED'); //PT- 7-1 Affectation de la tablette des abs. deportés
      SetControlText('PGENERAL', 'Saisie d''un mouvement d''absence'); //PT- 7-12
      if FicPrec = 'RESP' then SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCERESP') //PT20-2 : Tablette motif saisi responsable
      else if (FicPrec = 'SAL') then SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCESAL'); { PT48 }
    end;
  end;
  SetFocusControl('PCN_TYPECONGE');
  InitialiseZonesCongesPayes;
  GblEnterTypeMvt := GetField('PCN_TYPEMVT'); { PT56-1 }
end;

procedure TOM_AbsenceSalarie.BindemniteClick(Sender: TObject);
var
  Tobpris: tob;
begin
  if GetField('PCN_TYPEMVT') = 'CPA' then {PT- 8 ajout condition}
  begin
    TobPris := Tob.create('CONGES PRIS', nil, -1);
    ChargeLaTob(TobPris);
    TheTob := TobPris;
    AglLanceFiche('PAY', 'INDEMNITECP', '', '', 'Consultation');
    TheTob := nil;
    Tobpris.free;
  end;
end;

procedure TOM_AbsenceSalarie.ChargeLatob(Tobpris: tob);
var
  st: string;
  Q: tQuery;
begin
  st := 'SELECT * FROM ABSENCESALARIE ' +
    'WHERE PCN_SALARIE = "' + getfield('PCN_SALARIE') + '" ' +
    'AND PCN_TYPEMVT="' + getfield('PCN_TYPEMVT') + '"';
    // Deb PT95
  if (Origine = 'P') then
  begin
    st := st + ' AND PCN_TYPEMVT="PRE" ';
    st := st + ' AND PCN_ORDRE = ' + inttostr(getfield('PCN_ORDRE')) + ' ';
  end else begin
    st := st + ' AND PCN_TYPEMVT<>"PRE" ';
    // Fin PT95
    if getfield('PCN_CODERGRPT') <> -1 then
      // 1 seul enregistrement ds la tob
      st := st + 'AND PCN_ORDRE = ' + inttostr(getfield('PCN_ORDRE')) + ''
    else
      // enregistrement éclaté lors du paiement0
      st := st + ' AND PCN_CODERGRPT = ' + inttostr(getfield('PCN_ORDRE')) + ' ORDER BY PCN_DATEVALIDITE';
  end;
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then
    TobPris.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
  ferme(Q);
end;

procedure TOM_AbsenceSalarie.OnloadRecord;
var
  Validresp, Exportok,St: string;
  Q,Q1: TQuery;
  GB: TGroupBox;
  TT: TTabsheet;
begin
  inherited;
  if not (DS.State in [dsInsert]) then DerniereCreate := ''; { PT56-1 }
  GblModifDate := False; { PT55 }

  if (Origine <> 'P') then  // PT95
  begin
    // d PT49 IJSS
    IJSSaSolder := False;
    if (Origine <> 'I') then
      // champs inaccessibles s'il ne s'agit pas du suivi des IJSS
    begin
      //  SetControlProperty('PIJSS','Enabled', False);
      setcontrolenabled('PCN_NBJCALEND', False);
      setcontrolenabled('PCN_NBJCARENCE', True);
      setcontrolenabled('PCN_NBJIJSS', False);
      setcontrolenabled('PCN_IJSSSOLDEE', False);
    end
    else
      // suivi IJSS - onglet PGeneral inactif
    begin
      SetControlProperty('PGeneral', 'Enabled', False);
    end;

    // IJSS - positionnement sur l'onglet PIJSS
    if (Origine = 'I') then
    begin
      setcontrolenabled('BINSERT', false);
      setcontrolvisible('BINSERT', false);
      TPageControl(GetControl('PAGES')).ActivePage := TTabSheet(GetControl('PIJSS'));
    end;
    // f PT49 IJSS
  end else begin  // PT95
    SetControltext('HDEB',GetControlText('PCN_HDEB'));
    SetControltext('HFIN',GetControlText('PCN_HFIN'));
    SetControlEnabled('PCN_DATEFINABS', False); //PT101
  end;
  //PT30 DateModifRecap:=Idate1900;
  //DDebPaieAju:=idate1900;DFinPaieAju:=idate1900; //PT- 9 //PT31
  Resp := '';
  NomResp := '';
  MailResp := '';
  //BDateEnter := false;
  FirstValid := True;
  Premier_passage := True;
  if salarie = '' then Salarie := GeTField('PCN_SALARIE');
  if (GeTField('PCN_SALARIE') <> '') and (Salarie <> GeTField('PCN_SALARIE')) then { PT56-1 }
    Salarie := GeTField('PCN_SALARIE');
  setcontrolenabled('BINSERT', false);
  if (Origine <> 'P') then  // PT95
  begin
    // d PT498 IJSS
    if (Origine <> 'I') then
      Setcontrolenabled('BDELETE', (not ((Action = 'CREATION') and (origine = 'A'))))
    else
      // suivi des IJSS - Pas de suppression, pas d'attestation
    begin
      Setcontrolvisible('BDELETE', FALSE);
      Setcontrolenabled('BDELETE', FALSE);
      Setcontrolvisible('BATTEST', FALSE);
      Setcontrolenabled('BATTEST', FALSE);
    end;
    // f PT49 IJSS
  {$IFDEF EAGLCLIENT} //PT40-2
    if Origine <> 'E' then SetControlVisible('BDELETE', (not ((Action = 'CREATION') and (origine = 'A'))));
  {$ENDIF}
  end;
  Setcontrolenabled('BDEFAIRE', (Action <> 'MODIFICATION'));
  SetControlenabled('BVALIDER', (LibOrdre <> 'ORDREG'));
  OkPassage := False; //PT- 8
  SetControlVisible('SORTIE',False);  { PT72 }
  SetControlVisible('LSORTIE',False); { PT72 }
  {DEB PT- 8 }
  DTclot := idate1900;  { PT90 }
  Etablissement := '';  Nom := '';  Prenom := '';
  calendrier := '';  standcalend := '';
  CPEtab := False;   CPSal := False; { PT43-2 }
  DateEntree := Idate1900; //PT28 ajout date entree sortie
  DateSortie := Idate1900;
  if Salarie <> '' then
    Begin
    Q := opensql('SELECT PSA_LIBELLE,PSA_PRENOM,PSA_CALENDRIER,PSA_DATEENTREE,PSA_DATESORTIE,' + //PT28 ajout date entree sortie
    'PSA_STANDCALEND,ETB_ETABLISSEMENT,ETB_DATECLOTURECPN,ETB_CONGESPAYES,PSA_CONGESPAYES ' +
    'FROM SALARIES ' +
    'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
    'WHERE PSA_SALARIE= "' + salarie + '" ', TRUE);
    if not Q.eof then
      begin
      Etablissement := Q.findfield('ETB_ETABLISSEMENT').Asstring;
      Nom := Q.findfield('PSA_LIBELLE').Asstring;
      Prenom := Q.findfield('PSA_PRENOM').Asstring;
      CPEtab := (Q.findfield('ETB_CONGESPAYES').Asstring = 'X'); { PT43-2 }
      CPSal := (Q.findfield('PSA_CONGESPAYES').Asstring = 'X'); { PT43-2 }
      DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
      calendrier := Q.findfield('PSA_CALENDRIER').AsString;
      standcalend := Q.findfield('PSA_STANDCALEND').AsString;
      DateEntree := Q.findfield('PSA_DATEENTREE').AsDateTime; //PT28 ajout date entree sortie
      DateSortie := Q.findfield('PSA_DATESORTIE').AsDateTime;
      InitEntreeSortie;  { PT90 }
      end;
    End;
  if GetField('PCN_ETABLISSEMENT') <> Etablissement then { PT93 }
    SetField('PCN_ETABLISSEMENT', Etablissement);
  Ferme(Q);
  GblTypeMvt := GetField('PCN_TYPEMVT');
  GblOrdre := GetField('PCN_ORDRE');
  GblEnterTypeMvt := GetField('PCN_TYPEMVT'); { PT56-1 }
  GblDateValidite := GetField('PCN_DATEVALIDITE');
  GblPeriodeCP := GetField('PCN_PERIODECP');
  GblJours := GetField('PCN_JOURS'); { DEB PT59 }
  GblEnterTypeConge := GetField('PCN_TYPECONGE');
  GblEnterTypeImpute := GetField('PCN_TYPEIMPUTE');
  GblEnterSensAbs := GetField('PCN_SENSABS'); { FIN PT59 }
  if (Origine = 'P') then  // PT95
  begin
    SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFPRESENCELIB'); { PT74 }
  end;
  if Origine = 'A' then
  begin
    Bok := true;
    SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCELIB'); { PT74 }
  end
  else
    if Origine = 'C' then
  begin
    if RechDom('PGMVTCONGES', GetField('PCN_TYPECONGE'), True) = 'S' then
      SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMVTCONGESS')
    else
      SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMVTCONGES');
  end
  else
    if ((Origine = 'E') and (FicPrec = 'RESP')) then
  begin
    TT := TTabsheet(getcontrol('PGENERAL'));
    if TT <> nil then
      TT.caption := TraduireMemoire('Gestion des mouvements du salarié ') + Nom + ' (' + Salarie + ')';
    if GetField('PCN_VALIDSALARIE') = 'RES' then //PT20-2 : Tablette motif saisi responsable
      SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCERESP');
  end;
  OldValidResp := GetField('PCN_VALIDRESP');
  {FIN PT- 8}
  Validresp := '';
  Exportok := '';
  ValidResp := GetField('PCN_VALIDRESP');
  Exportok := GetField('PCN_EXPORTOK');
  {DEB PT- 7-8 On recupère les valeurs en chargement de fiche}
  if getfield('PCN_DATEDEBUTABS') <> null then
    DDAbs := getfield('PCN_DATEDEBUTABS')
  else
    DDAbs := idate1900;
  if getfield('PCN_DATEFINABS') <> null then
    DFAbs := getfield('PCN_DATEFINABS')
  else
    DFAbs := idate1900;
  if (Origine <> 'P') then  // PT95
  begin
    AffecteNodemj(getfield('PCN_DEBUTDJ'), DebDJ);
    AffecteNodemj(getfield('PCN_FINDJ'), FinDJ);
  end;
  {FIN PT- 7-8 }
  // on est en eagl
  if (Origine = 'E') then
  begin
    { DEB PT40-4 Initialisation des infos nécessaires pour envoie mail }
    EJoursPris := GetField('PCN_JOURS');
    Resp := '';
    NomResp := '';
    MailResp := '';
    OkMailEnvoie := False;
    MailDate := Idate1900; //PT26-2
    MailSal := '';
    SetControltext('NOMRESP', '');
    if (FicPrec = 'SAL') then
    begin
      //DEB PT- 10
      Q := OpenSql('SELECT PSE_SALARIE,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSA_LIBELLE,PSA_PRENOM FROM DEPORTSAL ' + //PT26-2 Ajout EMAILDATE
        'LEFT JOIN SALARIES ON PSE_SALARIE=PSA_SALARIE ' + //PT20-1 modif. requête
        'WHERE PSE_SALARIE IN (SELECT RESP.PSE_RESPONSABS FROM DEPORTSAL RESP ' +
        'WHERE RESP.PSE_SALARIE="' + GetField('PCN_SALARIE') + '")', True);
      if not Q.eof then
      begin
        //DEB PT20-1 Gestion des mails responsable
        Resp := Q.FindField('PSE_SALARIE').AsString;
        NomResp := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString;
        MailResp := Q.FindField('PSE_EMAILPROF').AsString;
        OkMailEnvoie := (Q.FindField('PSE_EMAILENVOYE').AsString = 'X');
        MailDate := Q.FindField('PSE_EMAILDATE').AsDateTime; //PT26-2 Pour gestion validité
        // DEB PT67
        if NomResp = ' ' then
        begin
          St := 'SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE ="' + Resp + '"';
          Q1 := OpenSql(St, True);
          if not Q1.eof then
            NomResp := Q1.FindField('PSI_LIBELLE').AsString + ' ' + Q1.FindField('PSI_PRENOM').AsString;
          Ferme(Q1);
        end;
        // FIN PT67
        SetControltext('NOMRESP', TraduireMemoire('Par')+' '+ NomResp);
        //FIN PT20-1
      end;
      Ferme(Q);
      //FIN PT- 10
    end
    else
      if (FicPrec = 'RESP') or (FicPrec = 'ADM') then
    begin
      { DEB PT66 }
      if (GblTypeSal = 'PSA') OR (GblTypeSal = 'SAL') OR (GblTypeSal = '')then
        St := 'SELECT PSE_SALARIE,PSE_RESPONSABS,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSA_LIBELLE,PSA_PRENOM FROM DEPORTSAL ' +
        'LEFT JOIN SALARIES ON PSE_RESPONSABS=PSA_SALARIE ' +
        'WHERE PSE_SALARIE="' + GetField('PCN_SALARIE') + '"'
      else
      if GblTypeSal = 'INT' then
        St := 'SELECT PSE_SALARIE,PSE_RESPONSABS,PSE_EMAILENVOYE,PSE_EMAILPROF,PSE_EMAILDATE,' +
        'PSI_LIBELLE,PSI_PRENOM FROM DEPORTSAL ' +
        'LEFT JOIN INTERIMAIRES ON PSE_RESPONSABS=PSI_INTERIMAIRE ' +
        'WHERE PSE_SALARIE="' + GetField('PCN_SALARIE') + '"';
      { FIN PT66 }
      Q := OpenSql(St, True);
      if not Q.eof then
      begin
        Resp := Q.FindField('PSE_RESPONSABS').AsString;
        if (GblTypeSal = 'PSA') OR (GblTypeSal = 'SAL')then
        NomResp := Q.FindField('PSA_LIBELLE').AsString + ' ' + Q.FindField('PSA_PRENOM').AsString
        else
        if GblTypeSal = 'INT' then
          NomResp := Q.FindField('PSI_LIBELLE').AsString + ' ' + Q.FindField('PSI_PRENOM').AsString;
        // DEB PT67
        if Trim(NomResp) = '' then
        begin
          St := 'SELECT PSI_LIBELLE,PSI_PRENOM FROM INTERIMAIRES WHERE PSI_INTERIMAIRE ="' + Resp + '"';
          Q1 := OpenSql(St, True);
          if not Q1.eof then
            NomResp := Q1.FindField('PSI_LIBELLE').AsString + ' ' + Q1.FindField('PSI_PRENOM').AsString;
          Ferme(Q1);
        end;
        // FIN PT67
        MailSal := Q.FindField('PSE_EMAILPROF').AsString;
        SetControltext('NOMRESP', TraduireMemoire('Par')+' '+ NomResp);
      end;
      Ferme(Q);
    end;
    { FIN PT40-4 }

    if (GetField('PCN_VALIDABSENCE') <> '') and (GetField('PCN_VALIDABSENCE') <> Resp) then
      Begin  { DEB PT70 }
      St := RechDom('PGSALARIE', GetField('PCN_VALIDABSENCE'), False);
      if (St <> '') then
        SetControltext('NOMRESP', TraduireMemoire('Par')+ ' ' +St )
      else
        Begin
        St := RechDom('PGSALARIEINT', GetField('PCN_VALIDABSENCE'), False);
        if (St <> '') then
           SetControltext('NOMRESP', TraduireMemoire('Par')+' ' +St )
        else
           SetControltext('NOMRESP', TraduireMemoire('Par l''administrateur') );
        End;
      End;   { FIN PT70 }
    Setcontrolenabled('BDEFAIRE', false);

    if ACTION = 'CREATION' then // PT52
      if (FicPrec = 'SAL')  then
      begin
        if (Validresp <> 'ATT') or (Exportok <> '-') then
        begin
          SetControlText('PGENERAL', 'Absence en consultation seule'); //PT- 7-12
          SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCED'); { PT48 }
        end
        else
        begin
          SetControlText('PGENERAL', 'Modification de l''absence');
          SetControlProperty('PCN_TYPECONGE', 'Datatype', 'PGMOTIFABSENCESAL'); { PT48 }
        end;
      end
      else
      begin
        if (Exportok = 'X') or (ValidREsp = 'REF') or (ValidResp = 'NAN') then
          SetControlText('PGENERAL', 'Absence en consultation seule') //PT- 7-12
        else
          SetControlText('PGENERAL', 'Absence en consultation seule'); //PT- 7-12
      end;

    //alimentation des zones récapitulatifs CP et RTT pour le salarié
    AffectInfoRecap(Salarie); //PT32 Integration affichage recap dans procedure  { PT92 }


    GB := TGroupBox(getcontrol('TITRERECAPCP'));
    if GB <> nil then
    begin
      GB.Caption := TraduireMemoire('Récapitulatif congés payés');
      if (VH_Paie.PGECabDateIntegration <> idate1900) then
      begin
        //DEB PT20-5 pour recup datemotifrecap
        //PT30 DateModifRecap:=q.findfield('PRS_DATEMODIF').asdatetime;
        gb.Caption := TraduireMemoire (gb.caption) + ' '+TraduireMemoire('au')+' '+ datetostr(VH_Paie.PGECabDateIntegration); {PT30 DateModifRecap}
        //FIN PT20-5
      end;
    end;
    GB := TGroupBox(getcontrol('TITRERECAPRTT'));
    if GB <> nil then
    begin
      GB.Caption := TraduireMemoire('Récapitulatif RTT');
      { if not q.eof then //PT- 17-2 Mise en commentaire
         gb.Caption := gb.caption + ' au ' + datetostr(q.findfield('PRS_DATEMODIF').asdatetime);}
    end;
  end;

  ChargeTob_MotifAbsence;  { PT76-2 Pour recharger l'affichage des onglets en navigation }

  // PT1-2 PH 12-06-2001
  if Origine = 'E' then
    enableChampOngletModeAgl
  else
    enableChampOngletMvtCp;
  AfficheTitre;
  GblTypeConge := GetField('PCN_TYPECONGE');
// d PT77
{$IFDEF EAGLCLIENT}
  if getcontrol('PCN_SALARIE1') <> nil then //PT96
    SetControlText('PCN_SALARIE1',GetField('PCN_SALARIE'));
  if getcontrol('PCN_LIBELLE1') <> nil then //PT96
    SetControlText('PCN_LIBELLE1',GetField('PCN_LIBELLE'));
  if getcontrol('NOMSAL1') <> nil then      //PT96
    SetControlText('NOMSAL1',Nom);
{$ENDIF}
  // f PT77
  GblTypeConge := GetField('PCN_TYPECONGE');

  GblEnterEtatPostPaie := GetField('PCN_ETATPOSTPAIE'); { PT94 }

end;

procedure TOM_AbsenceSalarie.OnUpdateRecord;
var
  st, StTerme, pStAffaire: string;
  Q: TQuery;
  Body: TStringList;
{$IFDEF EAGLCLIENT}
  D: THedit;
{$ELSE}
  D: THDBedit;
{$ENDIF}
  Pris, Acquis, Restants, Base, MBase, Rel: Double; { PT42-3 }
  NbHeures : Real; //PT100
  PGYPL : TypePGYPL;
begin
  inherited;
  if (Action = 'CONSULTATION') then exit;
  if (Origine = 'P') then error := 0;
  if GetField('PCN_TYPEMVT') <> GblTypeMvt then { DEB PT56-1 }
  begin
    SetField('PCN_ORDRE', IncrementeSeqNoOrdre(GetField('PCN_TYPEMVT'), Salarie));
    if (Origine <> 'P') then  // PT95
      SetField('PCN_CODERGRPT', getfield('PCN_ORDRE'));
    GblTypeMvt := GetField('PCN_TYPEMVT');
  end; { FIN PT56-1 }

  OnFerme := False; { DEB PT56-1 }
  if (DS.State in [dsInsert]) then
    DerniereCreate := GetField('PCN_SALARIE') + ';' + GetField('PCN_TYPECONGE') + ';' + IntToStr(GetField('PCN_ORDRE'))
  else
    if (DerniereCreate = GetField('PCN_SALARIE') + ';' + GetField('PCN_TYPECONGE') + ';' + IntToStr(GetField('PCN_ORDRE'))) then OnFerme := True; // le bug arrive on se casse !!!
  { FIN PT56-1 }

  if (Error = 0) or ((Error <> 1) and (Error <> 10)) then //Deb PT- 7-11 Error 1 et 10 initialisé ds le changefield
  begin
    if Origine = 'P' then //PT95
    begin
      ControleZonesOngletCongesPayes;
      //PT100 - Début
      ControleSaisiePresence;

      If Error = 0 Then
      Begin
           NbHeures := 0;
           {$IFDEF EAGLCLIENT}
               D := THEdit(getcontrol('PCN_HEURES'));
           {$ELSE}
               D := THDBEdit(getcontrol('PCN_HEURES'));
           {$ENDIF}
           If Assigned(D) Then NbHeures := Valeur(D.Text);
           {$IFDEF EAGLCLIENT}
               D := THEdit(getcontrol('PCN_NBHEURESNUIT'));
           {$ELSE}
               D := THDBEdit(getcontrol('PCN_NBHEURESNUIT'));
           {$ENDIF}
        //   If Assigned(D) Then NbHeures := NbHeures + Valeur(D.Text);
           Error := ControleMouvementsPresence (Action,T_MotifAbs,Getfield('PCN_SALARIE'),Getfield('PCN_TYPECONGE'),
                                          GetField('PCN_DATEDEBUTABS'),GetField('PCN_DATEFINABS'),GetField('PCN_HDEB'),GetField('PCN_HFIN'),
                                          NbHeures, GetField('PCN_ORDRE'), GetField('PCN_RESSOURCE'));
           If Error <> 0 Then
           Begin
               Case Error Of
                 15:SetFocusControl('PCN_DATEDEBUTABS');
                 16:SetFocusControl('PCN_DATEFINABS');
                 32:SetFocusControl('HDEB');
                 17:Begin
                       ComplMessage := FloatToStr(T_MotifAbs.GetValue('PMA_JRSMAXI'));
                       SetFocusControl('PCM_HEURES');
                    End;
                 18:Begin
                       ComplMessage := FloatToStr(T_MotifAbs.GetValue('PMA_JRSMAXI'));
                       SetFocusControl('PCM_HEURES');
                    End;
                 23:SetFocusControl('PCN_DATEDEBUTABS');
               End;
           End;
      End;
      //PT100 - Fin
    end else ControleZonesOngletCongesPayes;
  end;

  if (Error <> 0) and (Error <> 14) then
    If Origine = 'P' then  //PT95
      PgiBox(RecupMessageErrorPresence(Error)+ ComplMessage, Ecran.caption)
    else
      PgiBox(RecupMessageErrorAbsence(Error) + ComplMessage, Ecran.caption);  { PT92 }
  LastError := Error; //Fin PT- 7-11
  if LastError <> 0 then exit;
  {DEB PT28 Mise en commentaire, chargé dans le OnLoadRecord
  st := 'SELECT PSA_DATESORTIE FROM SALARIES WHERE PSA_SALARIE="'+salarie+'"';
    Q := Opensql(st,true);
    if not Q.eof then
       begin
       DateZ := Q.findfield('PSA_DATESORTIE').asdatetime;
       if DateZ > 10 then}
  if DateSortie > idate1900 then
  begin
    {ferme(Q); PT28 Utilisation DateSortie}
    st := 'SELECT PPU_SALARIE,PPU_DATEDEBUT,PPU_DATEFIN FROM PAIEENCOURS WHERE PPU_SALARIE = "' + Salarie + '"' +
      ' AND (PPU_DATEDEBUT <= "' + usdatetime(DateSortie) + '" AND PPU_DATEFIN >= "' + usdatetime(DateSortie) + '")';
    Q := Opensql(st, true);
    if not Q.eof then
      HShowMessage('1;Attention;Le mouvement de solde a déjà été généré dans le bulletin du ' +
        datetostr(Q.findfield('PPU_DATEDEBUT').asdatetime) + ' au ' + datetostr(Q.findfield('PPU_DATEFIN').asdatetime) + ' .#13#10' +
        ' Si vous réalisez des modifications de mouvements CP, ils ne seront pas pris en compte dans le bulletin.#13#10' +
        ' Si vous souhaitez recalculer le mouvement de solde, supprimez et recréez votre bulletin;W;Y;Y;N;', '', '');
    Ferme(Q);
  end;
  {  end;
FIN PT28}
  // DEB PT60
  St := PGRendEtabUser();
  if St <> '' then
  begin
    St := RechDom('PGSALARIE', Salarie, FALSE);
    if St = '' then
    begin
      LastError := 1;
      PGIBox('Vous ne pouvez pas saisir de mouvements d''absences pour ce matricule salarié.', Ecran.Caption); { PT59 }
      Exit;
    end;
  end;
  // FIN PT60

  if Action = 'MODIFICATION' then
    Premier_passage := True;
  if (getfield('PCN_TYPECONGE') = 'AJP') AND (GetField('PCN_TYPEMVT') = 'CPA') then  { PT82 }
  begin
    //      setfield('PCN_DATEFIN'        ,getfield('PCN_DATEVALIDITE'));  PT- 9
    setfield('PCN_DATEPAIEMENT', getfield('PCN_DATEVALIDITE'));
    SetField('PCN_DATEDEBUTABS', getfield('PCN_DATEVALIDITE'));
    SetField('PCN_DATEFINABS', getfield('PCN_DATEVALIDITE'));
  end
  else
    if (getField('PCN_TYPECONGE') = 'CPA') AND (GetField('PCN_TYPEMVT') = 'CPA') then //DEB PT- 3 { PT82 }
  begin
    if getfield('PCN_DATEPAIEMENT') < 10 then SetField('PCN_DATEPAIEMENT', Date);
    SetField('PCN_CODETAPE', 'P');
  end; //FIN PT- 3
{$IFDEF EAGLCLIENT}
  D := THedit(getcontrol('PCN_DATEDEBUTABS'));
{$ELSE}
  D := THDBedit(getcontrol('PCN_DATEDEBUTABS'));
{$ENDIF}
  if D <> nil then
    if Valeur(D.Text) = 0 then Setfield('PCN_DATEDEBUTABS', 0);
{$IFDEF EAGLCLIENT}
  D := THedit(getcontrol('PCN_DATEFINABS'));
{$ELSE}
  D := THDBedit(getcontrol('PCN_DATEFINABS'));
{$ENDIF}
  if D <> nil then
    if Valeur(D.Text) = 0 then Setfield('PCN_DATEFINABS', 0);
  {DEB PT- 15 Mise en commentaire affectation datevalité
  D := THedit(getcontrol('PCN_DATEVALIDITE'));
  if D <> nil then
     IF Valeur(D.Text) <= idate1900 then Setfield('PCN_DATEVALIDITE',idate1900); PT- 9}
  if GetField('PCN_DATEVALIDITE') <> null then
    if HEnt1.IsValidDate(GetField('PCN_DATEVALIDITE')) then
      if GetField('PCN_DATEVALIDITE') < idate1900 then Setfield('PCN_DATEVALIDITE', idate1900);
  //FIN PT- 15

  if (Origine = 'A') or (Origine = 'E') then
  begin
    if (GetField('PCN_TYPECONGE') = 'PRI') and (GetField('PCN_TYPEMVT') <> 'CPA') then
      SetField('PCN_TYPEMVT', 'CPA')
    else
      if (GetField('PCN_TYPECONGE') <> 'PRI') and (GetField('PCN_TYPEMVT') <> 'ABS') then
      SetField('PCN_TYPEMVT', 'ABS');
  end
  else
    if Origine = 'C' then
    if GetField('PCN_TYPEMVT') <> 'CPA' then SetField('PCN_TYPEMVT', 'CPA');

  { Dans le cas d'une saisie CP autre que PRI }
  { DEB PT42-3 }
  if (LastError = 0) and (Origine = 'C') and (GetField('PCN_TYPECONGE') <> 'PRI') and (GetField('PCN_PERIODECP') > 0) and (PGAcces = False) then
  begin
    if (ACTION = 'CREATION') then
    begin
      AffichelibelleAcqPri((GetField('PCN_PERIODECP') - 1), Salarie, Idate1900, 0, Pris, Acquis, Restants, Base, MBase, false, False);
      if Pris > 0 then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez saisir de mouvement pour cette période.#13#10' +
          'Vous avez entamé vos congés payés sur la période suivante.', Ecran.Caption);
        Exit;
      end;
      if (GetField('PCN_TYPECONGE') <> 'PRI') and (GetField('PCN_TYPECONGE') <> 'SLD') and (GetField('PCN_PERIODECP') >= 2) and (GetField('PCN_CODETAPE') <> 'C') then
        SetField('PCN_CODETAPE', 'C'); { PT44-2 }
    end;
    if (GetField('PCN_TYPECONGE') = 'CPA') or (GetField('PCN_TYPECONGE') = 'AJP') or ((GetField('PCN_TYPECONGE') = 'AJU') and (GetField('PCN_SENSABS') = '-')) AND (GetField('PCN_TYPEMVT') = 'CPA') then { PT82 }
    begin
      Rel := 0;
      AffichelibelleAcqPri(GetField('PCN_PERIODECP'), Salarie, Idate1900, 0, Pris, Acquis, Restants, Base, MBase, false, False);
      { DEB PT59 Ne tient pas compte du mouvement en cours }
      if (GblEnterTypeConge = 'AJU') and (GblEnterSensAbs = '-') then
      begin
        if (GblEnterTypeImpute = 'ACQ') then Acquis := Acquis + GblJours
        else
          if (GblEnterTypeImpute = 'REL') then Rel := Rel + GblJours;
      end
      else
        if (GblEnterTypeConge = 'AJP') and (GblEnterSensAbs = '+') then
      begin
        Pris := Pris + GblJours;
        Restants := Restants - GblJours;
      end
      else
        if (GblEnterTypeConge = 'AJP') and (GblEnterSensAbs = '-') then
      begin
        Pris := Pris - GblJours;
        Restants := Restants + GblJours
      end
      else
        if (GblEnterTypeConge = 'CPA') then
      begin
        Pris := Pris - GblJours;
        Restants := Restants + GblJours;
      end;
      { FIN PT59 }
      { Contrôle de cohérence }
      if (GetField('PCN_TYPECONGE') = 'AJU') and (GetField('PCN_SENSABS') = '-') then
      begin
        Q := OpenSql('SELECT PCN_TYPECONGE,PCN_JOURS,PCN_SENSABS FROM ABSENCESALARIE ' +
          'WHERE PCN_TYPEMVT="CPA" AND PCN_SALARIE="' + Salarie + '" ' +
          'AND (PCN_TYPECONGE="REL" OR (PCN_TYPECONGE="AJU" AND PCN_TYPEIMPUTE="REL")) ' +
          'AND PCN_PERIODECP=' + IntToStr(GetField('PCN_PERIODECP')), True);
        while not Q.eof do
        begin
          if (Q.FindField('PCN_TYPECONGE').AsString = 'REL') then
            Rel := Rel + Q.FindField('PCN_JOURS').AsFloat
          else
            if (Q.FindField('PCN_TYPECONGE').AsString = 'AJU') and (Q.FindField('PCN_SENSABS').AsString = '+') then
            Rel := Rel + Q.FindField('PCN_JOURS').AsFloat
          else
            if (Q.FindField('PCN_TYPECONGE').AsString = 'AJU') and (Q.FindField('PCN_SENSABS').AsString = '-') then
            Rel := Rel - Q.FindField('PCN_JOURS').AsFloat;
          Q.Next;
        end;
        Ferme(Q);
        if (GetField('PCN_TYPEIMPUTE') = 'ACQ') and (GetField('PCN_JOURS') > Acquis - Rel) then
        begin
          LastError := 1;
          PGIBox('Vous ne pouvez saisir d''ajustement d''acquis négatif supérieur aux acquis de la période : ' + FloatToStr(Acquis - Rel) + ' .', Ecran.Caption);
          Exit;
        end
        else
          if (GetField('PCN_TYPEIMPUTE') = 'REL') and (GetField('PCN_JOURS') > Rel) then
        begin
          LastError := 1;
          PGIBox('Vous ne pouvez saisir d''ajustement d''acquis négatif supérieur au reliquat de la période : ' + FloatToStr(Rel) + ' .', Ecran.Caption); { PT59 }
          Exit;
        end;
      end
      else
        if (GetField('PCN_TYPECONGE') = 'AJP') and (GetField('PCN_SENSABS') = '+') and (GetField('PCN_JOURS') > Pris) then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez saisir d''ajustement de pris positif supérieur aux pris payés de la période : ' + FloatToStr(Pris) + ' .', Ecran.Caption); { PT59 }
        Exit;
      end
      else
        if (GetField('PCN_TYPECONGE') = 'AJP') and (GetField('PCN_SENSABS') = '-') and (GetField('PCN_JOURS') > Restants) then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez saisir d''ajustement de pris négatif supérieur aux restants de la période : ' + FloatToStr(Restants) + ' .', Ecran.Caption); { PT59 }
        Exit;
      end
      else
        if (GetField('PCN_TYPECONGE') = 'CPA') and (GetField('PCN_JOURS') > Restants) then
      begin
        LastError := 1;
        PGIBox('Vous ne pouvez saisir une reprise de pris supérieur aux restants de la période : ' + FloatToStr(Restants) + ' .', Ecran.Caption); { PT59 }
        Exit;
      end;
    end;
  end;
  { FIN PT42-3 }

  { Dans le cas d'une saisie déportée }
  if (LastError = 0) and (ACTION = 'CREATION') and (Origine = 'E') then
  begin
{$IFDEF ETEMPS}
    { DEB PT39 }
    if not IfMotifabsenceSaisissable(GetField('PCN_TYPECONGE'), Ficprec) then
    begin
      PgiError(RecupMessageErrorAbsence(27), Ecran.caption);  { PT92 }
      LastError := 27;
      exit;
    end
    else //PT39
      if (FicPrec = 'ADM') and ((GetField('PCN_TYPECONGE') = 'MAL') or (GetField('PCN_TYPECONGE') = 'MAT')
      or (GetField('PCN_TYPECONGE') = 'PAT') or (GetField('PCN_TYPECONGE') = 'ACT')
      or (GetField('PCN_TYPECONGE') = 'AM1') or (GetField('PCN_TYPECONGE') = 'AMS')) then
    begin
      if GetField('PCN_VALIDRESP') <> 'VAL' then SetField('PCN_VALIDRESP', 'VAL');
      if GetField('PCN_EXPORTOK') <> 'X' then SetField('PCN_EXPORTOK', 'X');
    end;
    { FIN PT39 }
{$ENDIF}
    { PT84  Vérifie la présence du salarié dans la gestion d'affaire }
      If not EstLibre(GetField('PCN_DATEDEBUTABS'),GetField('PCN_DATEFINABS'),
      GetField('PCN_SALARIE'),GetField('PCN_DEBUTDJ'),GetField('PCN_FINDJ'),pStAffaire) then
         Begin
         PgiBox('Vous avez une intervention planifiée sur cette période pour l''affaire : '+
         pStAffaire+'#13#10 Merci de régulariser le planning avant de saisir cette absence.',Ecran.caption);  { }
         LastError := 1;
         exit;
         End;

    //PT26-1 Saisie du remplaçant sur site obligatoire
    if (VH_Paie.PGEcabSaisieRemp > 0) and (isnumeric(GetControlText('PCN_JOURS'))) then
      if (Trim(GetControlText('PCN_LIBCOMPL1')) = '') and (StrToFloat(GetControlText('PCN_JOURS')) >= VH_Paie.PGEcabSaisieRemp) then
      begin
        PgiError(RecupMessageErrorAbsence(22), Ecran.caption);  { PT92 }
        LastError := 22;
        exit;
      end;
    //PT20-3 Intégration d'une notion legislatif
    //PT26-2 Paramétrage dynamique de l'envoi de mail
    if VH_Paie.PgEcabFractionnement then
      if (getfield('PCN_TYPECONGE') = 'PRI') and (ficPrec = 'SAL') then
        if PgiAsk(RecupMessageErrorAbsence(21), Ecran.caption) = Mrno then  { PT92 }
        begin
          LastError := 1;
          exit;
        end;
    //SetControlEnabled('PCN_TYPECONGE',False); déplacé after update
    if (VH_Paie.PgEcabGestionMail) and (LastError = 0) then
    begin
      Body := TStringList.Create;
      if (ficPrec = 'SAL')  and (MailResp <> '') then
      begin
        //DEB PT20-1 Gestion de l'envoi de mail lors de la validation salarié
        //PT26-2 Ajout des paramètres sociétés : gestion de la durée de validité du mail
        //Test si raz ou validité du  mail pour envoie
        if (OkMailEnvoie = False) or (VH_Paie.PGEcabValiditeMail = 0)
          or ((MailDate <> idate1900) and ((Date - MailDate) > VH_Paie.PGEcabValiditeMail)) then
        begin
          St := GetParamSocSecur('SO_LIBELLE','');     { PT85 }
          if St <> '' then St := ' de la société '+St; { PT85 }
          St := 'Le salarié ' + RechDom('PGSALARIE', GetField('PCN_SALARIE'), False) + St;  { PT85 }
          St := St + ' vient de saisir une absence ' + RechDom('PGMOTIFABSENCED', GetField('PCN_TYPECONGE'), False);
          St := St + ' du ' + DateToStr(GetField('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_DEBUTDJ'), False);
          St := St + ' au ' + DateToStr(GetField('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_FINDJ'), False);
          St := St + ' soit une durée de ' + FloatToStr(GetField('PCN_JOURS')) + ' jours.';
          Body.Add(st);
          Q := OpenSql('SELECT COUNT (*) AS NB FROM ABSENCESALARIE ' +
            'LEFT JOIN DEPORTSAL ON  PSE_SALARIE=PCN_SALARIE ' +
            'LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE ' +  { PT75 }
            'WHERE PSE_RESPONSABS="' + Resp + '" AND PMA_MOTIFEAGL="X" ' +
            'AND PCN_VALIDRESP="ATT"', True);
          if not Q.Eof then
            if Q.FindField('NB').AsInteger > 0 then
            begin
              St := 'Par ailleurs, vous avez ' + IntToStr(Q.FindField('NB').AsInteger) + ' autre(s) absence(s) à valider.';
              Body.Add(St);
            end;
          Ferme(Q);
          if PGEnvoieMailCollaborateur('ECongés : Absence à valider', MailResp, Body) then
          begin
            ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="X",PSE_EMAILDATE="' + UsDateTime(date) + '" WHERE PSE_SALARIE="' + Resp + '"');
            PGIInfo('Un message a été envoyé sur la messagerie de votre responsable ' + NomResp, Ecran.caption); { PT71 }
          end;
        end;
      end
      else
        { DEB PT40-4 si création absence pour salarié alors envoie mail }
        if ((FicPrec = 'RESP') or (FicPrec = 'ADM')) and (MailSal <> '') then
      begin
        St := GetParamSocSecur('SO_LIBELLE','');     { PT85 }
        if St <> '' then St := ' de la société '+St; { PT85 }
        if FicPrec = 'ADM' then St := 'L''administrateur ' + RechDom('PGSALARIE', LeSalarie, False)+ St else St := 'Le responsable ' + NomResp + St; { PT85 }
        St := St + ' vient de saisir une absence ' + RechDom('PGMOTIFABSENCED', GetField('PCN_TYPECONGE'), False);
        St := St + ' du ' + DateToStr(GetField('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_DEBUTDJ'), False);
        St := St + ' au ' + DateToStr(GetField('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_FINDJ'), False);
        St := St + ' soit une durée de ' + FloatToStr(GetField('PCN_JOURS')) + ' jours.';
        Body.Add(st);
        if PGEnvoieMailCollaborateur('ECongés : Absence saisie', MailSal, Body) then
          PGIInfo('Un message a été envoyé sur la messagerie du salarié ' + RechDom('PGSALARIE', Salarie, False) + '.', Ecran.caption); { PT71 }
      end;
      { FIN PT40-4 }
      Body.Free;
    end;
  end;
  //RAZ du boolean de l'envoi de mail du responsable lors de la validation de l'abs salarié
  //PT26-2 Mise à jour date d'envoi
  if (LastError = 0) and (Origine = 'E') then
  begin
    if (ACTION = 'MODIFICATION') then
    begin
      if (FicPrec <> 'SAL') and (MailResp <> '') and (OkMailEnvoie = True) and (GetField('PCN_VALIDRESP') <> 'ATT') then
        ExecuteSql('UPDATE DEPORTSAL SET PSE_EMAILENVOYE="-" WHERE PSE_SALARIE="' + Resp + '"');
      { DEB PT40-4 en modification si absence refusée ou annulée }
      if ((FicPrec = 'RESP') or (FicPrec = 'ADM')) and (IsMailRefOrAnu) and (MailSal <> '') then
      begin
        if getfield('PCN_VALIDRESP') = 'REF' then StTerme := 'de refuser' else StTerme := 'd''annuler';
        Body := TStringList.Create;
        St := GetParamSocSecur('SO_LIBELLE','');     { PT85 }
        if St <> '' then St := ' de la société '+St; { PT85 }
        if FicPrec = 'ADM' then St := 'L''administrateur ' + RechDom('PGSALARIE', LeSalarie, False) + St else St := 'Le responsable ' + NomResp + St; { PT85 }
        St := St + ' vient ' + StTerme + ' l''absence ' + RechDom('PGMOTIFABSENCED', GetField('PCN_TYPECONGE'), False);
        St := St + ' du ' + DateToStr(GetField('PCN_DATEDEBUTABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_DEBUTDJ'), False);
        St := St + ' au ' + DateToStr(GetField('PCN_DATEFINABS')) + ' ' + RechDom('PGDEMIJOURNEE', GetField('PCN_FINDJ'), False);
        St := St + ' soit une durée de ' + FloatToStr(GetField('PCN_JOURS')) + ' jours.';
        Body.Add(st);
        if getfield('PCN_VALIDRESP') = 'REF' then StTerme := 'refusée' else StTerme := 'annulée';
        if PGEnvoieMailCollaborateur('ECongés : Absence ' + StTerme, MailSal, Body) then
          PGIInfo('Un message a été envoyé sur la messagerie du salarié ' + RechDom('PGSALARIE', Salarie, False) + '.', Ecran.caption); { PT71 }
        Body.Free;
      end;
      { FIN PT40-4 }

    end;
    //FIN PT20-1
    if (AValider) and (FicPrec <> 'SAL') and (LeSalarie <> GetField('PCN_VALIDABSENCE')) then
    begin
      SetField('PCN_VALIDABSENCE', LeSalarie);
      AValider := False;
    end;
  end;
  if Assigned(T_MotifAbs) then { DEB PT56-2 }
    if (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X') then
    begin
      // d PT49 IJSS
      if (IJSSaSolder = TRUE) and (VH_Paie.PGGestIJSS = True) then
      begin
        if (PGIAsk('Le nombre de jours d''IJSS est à zéro.#10#13' +
          ' Voulez-vous solder le mouvement ?', 'IJSS') = mrYes) then
          SetField('PCN_IJSSSOLDEE', 'X')
        else
          SetField('PCN_IJSSSOLDEE', '-');
      end;
      // f PT49 IJSS
    end; { FIN PT56-2 }
  if GetField('PCN_CODETAPE') = '' then SetField('PCN_CODETAPE', '...');
// DEB PT103
  if Action = 'CREATION' then
     begin
     if (GetField('PCN_TYPECONGE') = 'PRI') and (GetField('PCN_TYPEMVT') = 'CPA') then
        begin
        if (GetField ('PCN_ORDRE') <> GetField ('PCN_CODERGRPT')) then
           setField ('PCN_CODERGRPT',GetField ('PCN_ORDRE'));
        end;

     end;
// FIN PT103
//PT105
if ((GetField('PCN_TYPEMVT')='ABS') or
   ((GetField  ('PCN_TYPEMVT')='CPA') and
   (GetField ('PCN_TYPECONGE')='PRI'))) then
   begin
   PGYPL.Guid:= GetField ('PCN_GUID');
   PGYPL.Salarie:= GetField ('PCN_SALARIE');
   PGYPL.DateDebutAbs:= GetField ('PCN_DATEDEBUTABS');
   PGYPL.DateFinAbs:= GetField ('PCN_DATEFINABS');
   PGYPL.Libelle:= GetField ('PCN_LIBELLE');
   if (GetField ('PCN_ETATPOSTPAIE')<>'NAN') then
      CreatePGYPL (nil, PGYPL)
   else
      DeletePGYPL (PGYPL.Guid, '');
   end;
//FIN PT105

end;

//
//                       PARTIE ONCHANGE
//
//
//

procedure TOM_AbsenceSalarie.OnChangeField(F: TField);
var
  {PORTAGECWAS}
  Periode, Nodjm, NodjP: integer;
  temp, ExerPerEncours, DebPer, FinPer, St: string;
  Nb_J, Nb_H: double;
  init, aa, mm, jj, aaC, mmC, jjC: word;
  Prochain31Mai: tdatetime;
  CalDuree  : Boolean;
  Q: TQuery;
  {SalVide,} Change: boolean; // PT54  { PT90 } //PT100
  StWhere : String; //PT97
begin
  inherited;
  CalDuree := False;
  if (Action = 'CONSULTATION') then exit;
  //Debug(F.FieldName) ;
  // PT1-2 PH 12-06-2001  AfficheTitre;
  //PT- 16 if ((Action <> 'CONSULTATION') and not (DS.state in [dsinsert, dsedit] ))then
  //PT- 16   DS.edit;   Mise en commentaire
  {PORTAGECWAS Suppression des prise de contrôles }
    { DEB PT40-3 }
  if (F.FieldName = 'PCN_SALARIE') then
  begin
    Change := False;  { PT90 }
    if (Salarie <> Getfield('PCN_SALARIE')) then
    begin
      // d PT54
      Change := True;  { PT90 }
//      SalVide := False; //PT100
//      if (Salarie = '') then //PT100
        // la variable Salarie = '' : cas création sans sélection du salarié
//        SalVide := True; //PT100
      // f PT54
      CalDuree := True;
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Getfield('PCN_SALARIE')) < 11) and (isnumeric(Getfield('PCN_SALARIE'))) then
        Salarie := ColleZeroDevant(Getfield('PCN_SALARIE'), 10)
      else
        Salarie := Getfield('PCN_SALARIE');
    end;
    St := PGRendEtabUser(); // @@@@
    if (St <> '') and (Getfield('PCN_SALARIE') <>'' ) then
    begin
      St := RechDom('PGSALARIE', Salarie, FALSE);
      if St = '' then
      begin
        if ORIGINE = 'P' then //PT95
          PGIError ('Vous n''avez pas les droits pour saisir des mouvements pour ce salarié.', Ecran.Caption)
        else
          PGIError ('Vous n''avez pas les droits pour saisir des absences pour ce salarié.', Ecran.Caption);
        SetFocusControl ('PCN_SALARIE');
        Salarie := '';
        Setfield('PCN_SALARIE', '');
        DTclot := idate1900;
        Etablissement := '';
        Nom := '';
        Prenom := '';
        calendrier := '';
        standcalend := '';
        CPEtab := False;
        CPSal := False;
        DateEntree := Idate1900;
        DateSortie := Idate1900;
      end;
    end;

    // d PT54
    if (Change) then  { PT90 }
      // cas création sans sélection du salarié il faut recharger les données du salarié
    begin   { DEB PT90 }
        DTclot := idate1900;
        Etablissement := '';     Nom := '';             Prenom := '';
        calendrier := '';        standcalend := '';
        CPEtab := False;         CPSal := False;
        DateEntree := Idate1900; DateSortie := Idate1900;
      if Salarie <> '' then  { FIN PT90 }
        Begin
        //DEB PT97 Vérifier la confidentialité du salarié
        StWhere := SQLConf('SALARIES');
        if StWhere <> '' then
          StWhere := ' AND ' + StWhere;
        //FIN PT97
        Q := opensql('SELECT PSA_LIBELLE,PSA_PRENOM,PSA_CALENDRIER,PSA_DATEENTREE,PSA_DATESORTIE,' + //PT28 ajout date entree sortie
        'PSA_STANDCALEND,ETB_ETABLISSEMENT,ETB_DATECLOTURECPN,ETB_CONGESPAYES,PSA_CONGESPAYES ' +
        'FROM SALARIES ' +
        'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
        'WHERE PSA_SALARIE= "' + salarie + '" ' + StWhere, TRUE); //PT97
        if not Q.eof then
        begin
          Etablissement := Q.findfield('ETB_ETABLISSEMENT').Asstring;
          Nom := Q.findfield('PSA_LIBELLE').Asstring;
          Prenom := Q.findfield('PSA_PRENOM').Asstring;
          CPEtab := (Q.findfield('ETB_CONGESPAYES').Asstring = 'X');
          CPSal := (Q.findfield('PSA_CONGESPAYES').Asstring = 'X');
          DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
          calendrier := Q.findfield('PSA_CALENDRIER').AsString;
          standcalend := Q.findfield('PSA_STANDCALEND').AsString;
          DateEntree := Q.findfield('PSA_DATEENTREE').AsDateTime;
          DateSortie := Q.findfield('PSA_DATESORTIE').AsDateTime;
          InitEntreeSortie;  { PT90 }
          end
        else
          begin
          // salarié inexistant
          Salarie := '';
          Setfield('PCN_SALARIE', '');
          InitEntreeSortie;  { PT90 }
          end;
        Ferme(Q);
        End
      else InitEntreeSortie;
    end;
    // f PT54
  end;
  { FIN PT40-3 }

  if (F.FieldName = 'PCN_TYPECONGE') and (Getfield('PCN_TYPECONGE') <> '') then
  begin
    if Getfield('PCN_TYPECONGE') = 'PRI' then
    begin
      if (GetField('PCN_MVTPRIS') <> 'PRI') then Setfield('PCN_MVTPRIS', 'PRI');
      if (GetField('PCN_TYPEMVT') <> 'CPA') then SetField('PCN_TYPEMVT', 'CPA');
      if (GetField('PCN_PERIODEPY') <> -1) and (GetField('PCN_CODETAPE') = '...') then //PT- 13
        SetField('PCN_PERIODEPY', -1);
    end
    else
    begin
      if GetField('PCN_MVTPRIS') <> '' then setfield('PCN_MVTPRIS', '');
      if (GetField('PCN_TYPEMVT') <> 'ABS') and ((Origine = 'A') or (Origine = 'E')) then
        SetField('PCN_TYPEMVT', 'ABS');
    end;
    if (GblTypeConge <> Getfield('PCN_TYPECONGE')) then
    begin
      if (Origine <> 'E') then InitialiseZonesCongesPayes;
      CalDuree := True;
      GblTypeConge := Getfield('PCN_TYPECONGE');
    end;
    AffectLibelle; //PT- 7-1 MAJ Titre Ecran if Action = 'CREATION' then
    if ((GetField('PCN_TYPECONGE') = 'PRI') or (GetField('PCN_TYPECONGE') = 'CPA')) and (GetField('PCN_SENSABS') <> '-') AND (GetField('PCN_TYPEMVT') = 'CPA') then { PT82 }
      SetField('PCN_SENSABS', '-')
    else
      if (GetField('PCN_TYPECONGE') = 'REP') and (GetField('PCN_SENSABS') <> '+') AND (GetField('PCN_TYPEMVT') = 'CPA') then { PT82 }
      SetField('PCN_SENSABS', '+');
    {PT- 6-2}
    ChargeTob_MotifAbsence; //PT24

    //PT1-2 PH 12-06-2001
    if Origine = 'E' then enableChampOngletModeAgl else enableChampOngletMvtCp;

    if (GetField('PCN_TYPEMVT') = 'CPA') then { PT47-2 }
    begin
      GblPasControl := True; { PT59 Si modif typeconge pas d'affichage de controle }
      if (GetField('PCN_PERIODECP') = -1) and (getfield('PCN_TYPECONGE') <> 'PRI') then
        Begin
        if (DateSortie > Idate1900) AND (DateSortie<DtClot) then   { PT72 }
          Setfield('PCN_DATEVALIDITE', DateSortie)                 { PT72 }
        else
          Setfield('PCN_DATEVALIDITE', DtClot); { PT42-2 }
        End;
      {DEB PT- 9}
      if getfield('PCN_TYPECONGE') = 'AJP' then
      begin
        { DEB PT36 }
        if RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) then
        begin
          // DEB PT53 Rajout StrToDate pour avoir le bon type de donnée
          Setfield('PCN_DATEVALIDITE', StrToDate(FinPer)); { PT42-2 }
          if GetField('PCN_DATEDEBUT') <= idate1900 then setfield('PCN_DATEDEBUT', StrToDate(DebPer));
          if GetField('PCN_DATEFIN') <= idate1900 then setfield('PCN_DATEFIN', StrToDate(FinPer));
          // FIN PT53
        end
        else
          if (DateSortie > Idate1900) AND (DateSortie<DtClot) then     { PT72 }
            Setfield('PCN_DATEVALIDITE', DateSortie)                   { PT72 }
          else
            Setfield('PCN_DATEVALIDITE', DtClot); { PT59 }
        { FIN PT36 }
      end
      else
      begin
        if (GetField('PCN_DATEDEBUT') <> idate1900) and (GetField('PCN_CODETAPE') = '...') and (GetField('PCN_TYPECONGE') = 'PRI') then
          setfield('PCN_DATEDEBUT', idate1900);
        if (GetField('PCN_DATEFIN') <> idate1900) and (GetField('PCN_CODETAPE') = '...') and (GetField('PCN_TYPECONGE') = 'PRI') then
          setfield('PCN_DATEFIN', idate1900);
        GblPasControl := False; { PT59 }
      end;
      {FIN PT- 9}
    end;
    // d PT49 IJSS
    if Assigned(T_MotifAbs) then
      if (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X') then MajIJSS;
    // f PT49 IJSS
  end;
  {DEB PT- 9}
  if (F.Fieldname = 'PCN_DATEDEBUT') and (getfield('PCN_TYPECONGE') = 'AJP') then
    //PT31 if ((GetField('PCN_DATEDEBUT')<DDebPaieAju) or (GetField('PCN_DATEDEBUT')>DFinPaieAju)) AND (DDebPaieAju>idate1900) and  (DFinPaieAju>idate1900) then
  begin {PT31 Test si période de paie close}
    if Tob_Exercice = nil then
    begin
      Tob_Exercice := Tob.Create('Les exercices', nil, -1);
      Tob_Exercice.LoadDetailDb('EXERSOCIAL', '', '', nil, False);
    end;
    if (PGIsPeriodeClot(Tob_Exercice, GetField('PCN_DATEDEBUT'))) and (not (DS.State in [DsBrowse])) then
      PgiInfo('La date de début de session correspond à une période de paie close.', Ecran.caption);
    {PT31  Mise en commentaire : modification du test
    PgiBox('Vous devez renseigner une période d''imputation de paie comprise entre le '+DateToStr(DDebPaieAju)+' et le '+DateToStr(DFinPaieAju)+'.',Ecran.caption);
    SetField('PCN_DATEDEBUT',DDebPaieAju);}
  end;
  if (F.Fieldname = 'PCN_DATEFIN') and (getfield('PCN_TYPECONGE') = 'AJP') then
    //PT31 if ((GetField('PCN_DATEFIN')<DDebPaieAju) or (GetField('PCN_DATEFIN')>DFinPaieAju)) AND (DDebPaieAju>idate1900) and  (DFinPaieAju>idate1900) then
  begin {PT31 Test si période de paie close}
    if Tob_Exercice = nil then
    begin
      Tob_Exercice := Tob.Create('Les exercices', nil, -1);
      Tob_Exercice.LoadDetailDb('EXERSOCIAL', '', '', nil, False);
    end;
    if (PGIsPeriodeClot(Tob_Exercice, GetField('PCN_DATEFIN'))) and (not (DS.State in [DsBrowse])) then
      PgiInfo('La date de fin de session correspond à une période de paie close.', Ecran.caption);
    {PT31 PgiBox('Vous devez renseigner une période d''imputation de paie comprise entre le '+DateToStr(DDebPaieAju)+' et le '+DateToStr(DFinPaieAju)+'.',Ecran.caption);
    SetField('PCN_DATEFIN',DFinPaieAju);}
  end;
  {FIN PT- 9}

  if (F.Fieldname = 'PCN_DATEVALIDITE') then
  begin
    if getfield('PCN_DATEVALIDITE') = null then exit;
    if (DtClot <> idate1900) and (CPEtab) and ((Getfield('PCN_TYPECONGE') = 'PRI') or (Origine <> 'A')) then {PT- 8} { PT43-2 }
    begin
      Periode := CalculPeriode(DTClot, getfield('PCN_DATEVALIDITE'));
      {PT- 6-2}if (GetField('PCN_PERIODECP') <> Periode) and (GetField('PCN_PERIODEPY') = -1) and (Periode <> 50) then //PT- 11 Ajout cond. and (Periode<>50)
      begin
        if (Getfield('PCN_TYPECONGE') = 'PRI') and (GetField('PCN_CODETAPE') = '...') then Periode := 0; { PT44-1 }
        if GetField('PCN_PERIODECP') <> Periode then SetField('PCN_PERIODECP', Periode);  { PT91-2 }
      end;
    end;
    if (CPEtab) and (not GblPasControl) and (ACTION <> 'CREATION') and ((Getfield('PCN_TYPECONGE') = 'AJU') or (Getfield('PCN_TYPECONGE') = 'AJP')) and (GblPeriodeCP <> Getfield('PCN_PERIODECP')) then { PT43-2 } { PT59 }
    begin
      PGIBox('Vous ne pouvez saisir une date de validité correspondant à une période différente.', Ecran.caption);
      if GetField('PCN_DATEVALIDITE') <> GblDateValidite then
      begin
        GblPasControl := True; { PT59 }
        SetField('PCN_DATEVALIDITE', GblDateValidite);
        GblPasControl := False; { PT59 }
      end;
    end;
    {DEB PT- 9}
    {PT31 Mise en commmentaire
    if getfield('PCN_TYPECONGE') = 'AJP' then
      Begin
      RechercheDate(DDeb,Dfin,DTClot,IntToStr(CalculPeriode( DTClot,getfield('PCN_DATEVALIDITE'))));
      DDebPaieAju:=PlusDate(DDeb,1,'A');
      DFinPaieAju:=PlusDate(DFin,1,'A');
      End
    else
      Begin
      DDebPaieAju:=idate1900;
      DFinPaieAju:=idate1900;
      End        }
    {FIN PT- 9}
  end;
  if ((F.Fieldname = 'PCN_DATEPAIEMENT') and (getfield('PCN_TYPECONGE') = 'CPA')) AND (GetField('PCN_TYPEMVT') = 'CPA') then  { PT82 }
  begin
    if GetField('PCN_DATEVALIDITE') <> getField('PCN_DATEPAIEMENT') then {PT- 6-2}
      SetField('PCN_DATEVALIDITE', getField('PCN_DATEPAIEMENT'));
    if GetField('PCN_DATEDEBUT') <> getField('PCN_DATEPAIEMENT') then {PT- 6-2}
      SetField('PCN_DATEDEBUT', getField('PCN_DATEPAIEMENT'));
    if GetField('PCN_DATEFIN') <> getField('PCN_DATEPAIEMENT') then {PT- 6-2}
      SetField('PCN_DATEFIN', getField('PCN_DATEPAIEMENT'));
  end;
  { DEB PT55 }
  if (Origine = 'A') and ((F.Fieldname = 'PCN_DATEDEBUTABS') or (F.Fieldname = 'PCN_DATEFINABS') or (F.Fieldname = 'PCN_DATEFINABS') or (F.Fieldname = 'PCN_DATEVALIDITE')) then
  begin
    GblModifDate := True;
  end;
  { FIN PT55 }
  if (((F.Fieldname = 'PCN_DATEDEBUTABS') or (F.Fieldname = 'PCN_DATEFINABS'))
    or (F.fieldname = 'PCN_DEBUTDJ') or (F.fieldname = 'PCN_FINDJ') or (CalDuree = True)) then
  begin
    Temp := F.FieldName; //PT- 7-5 ajout ligne La methode GetField réaffecte la propriété fieldname
    Nodjm := 0;
    NodjP := 1;
    AffecteNodemj(getfield('PCN_DEBUTDJ'), Nodjm);
    AffecteNodemj(getfield('PCN_FINDJ'), Nodjp);
    //if BDateEnter = true then  PT- 7-8 Remplacement : On test si une zone a été modifié
    if (DDAbs <> getfield('PCN_DATEDEBUTABS')) or (DFAbs <> getfield('PCN_DATEFINABS'))
      or (DebDJ <> Nodjm) or (FinDJ <> NodjP) or (BDateEnter = True) or (CalDuree = True) then
    begin
      BDateEnter := FALSE;
      {DEB PT- 7-8 On reaffect les valeurs }
      if getfield('PCN_DATEDEBUTABS') <> null then DDAbs := getfield('PCN_DATEDEBUTABS')
      else DDAbs := idate1900;
      if getfield('PCN_DATEFINABS') <> null then DFAbs := getfield('PCN_DATEFINABS')
      else DFAbs := idate1900;
      DebDJ := Nodjm;
      FinDJ := Nodjp;
      {FIN PT- 7-8}
      if (Origine <> 'P') and (Temp = 'PCN_DATEFINABS') and (getField('PCN_DATEFINABS') > idate1900) then
        if GetField('PCN_DATEVALIDITE') <> getField('PCN_DATEFINABS') then {PT- 6-2}
          setfield('PCN_DATEVALIDITE', getfield('PCN_DATEFINABS'));
      BDateEnter := false;
      if ((GetField('PCN_TYPECONGE') = 'PRI') or (Origine = 'A') or (Origine = 'E')) then
      begin
        { PORTAGECWAS }
        error := 0; //Deb PT- 7-11
        //DEB PT- 7-6 Controle de cohérence MAT/PAM sur même journée
        if (getfield('PCN_DATEFINABS') = getfield('PCN_DATEDEBUTABS')) and (Nodjm = 1) and (Nodjp = 0) then
        begin
          PgiBox('Vous ne pouvez être absent de l''aprés midi au matin sur une même journée.', Ecran.caption);  { PT92 }
          SetField('PCN_DEBUTDJ', 'MAT');
          AffecteNodemj('MAT', Nodjm);
          DebDJ := Nodjm;
          SetField('PCN_FINDJ', 'PAM');
          AffecteNodemj('PAM', Nodjp);
          FinDJ := Nodjp;
        end; //FIN PT- 7-6
        if (getfield('PCN_DATEFINABS') < getfield('PCN_DATEDEBUTABS')) then
        begin
          if (Temp = 'PCN_DATEDEBUTABS') then //PT- 7-5 Si modif date deb abs alors réaffect date fin abs si inferieur
          begin
            SetField('PCN_DATEFINABS', getfield('PCN_DATEDEBUTABS'));
            SetField('PCN_DATEVALIDITE', getfield('PCN_DATEDEBUTABS')); //PT29 réaffect aussi date validité
          end
          else
            if (Error = 0) then
          begin
            Error := 1; //PT- 7-11 LastErrorMsg:='La date de fin ne peut être inférieure à la date début' ;
            PgiBox('La date de fin ne peut être inférieure à la date de début.', Ecran.caption);
            //setFocuscontrol(F.fieldName); PT- 7-11
            exit;
          end;
        end;
        if (GetField('PCN_TYPEMVT') = 'CPA') and (CPSal) then { PT47-1 }
        begin
          decodedate(getfield('PCN_DATEDEBUTABS'), aa, mm, jj);
          decodedate(DtClot, aaC, mmC, jjC); { PT41 }
          if mm > mmC then Prochain31Mai := PGEncodeDateBissextile(aa + 1, mmC, jjC) { PT41 }
          else Prochain31Mai := PGEncodeDateBissextile(aa, mmC, jjC); { PT41 }
          if (Prochain31Mai < getfield('PCN_DATEFINABS')) then
          begin
            if (Error = 0) then
            begin
              Error := 10; //PT- 7-11 LastErrorMsg:='Un mouvement d''absence ne peut être à cheval sur le 31 Mai' ;
              PgiBox('Un mouvement d''absence ne peut être à cheval sur la date de clôture.', Ecran.caption); { PT41 }
              exit;
            end;
          end;
        end;
        { DEB PT79 }
        CalculNbJourAbsence(GetField('PCN_DATEDEBUTABS'), GetField('PCN_DATEFINABS'), Salarie,
        Etablissement, Getfield('PCN_TYPECONGE'), T_MotifAbs, Nb_J, Nb_H, Nodjm, Nodjp);
       (* Nb_J := 0;
        Nb_H := 0;
        if Assigned(T_MotifAbs) then CalCivil := (T_MotifAbs.GetValue('PMA_CALENDCIVIL') = 'X') //PT24-1
        else CalCivil := False;
        //DEB PT20-4 Décompte calendaire pour motif indexé sur calendrier civil
        if (CalCivil) then //PT24-1 Modif. clause
        begin
          Nb_J := Getfield('PCN_DATEFINABS') - getfield('PCN_DATEDEBUTABS') + 1;
          if getfield('PCN_DEBUTDJ') = 'PAM' then Nb_J := Nb_J - 0.5;
          if getfield('PCN_FINDJ') = 'MAT' then Nb_J := Nb_J - 0.5;
          Nb_H := Nb_J * 7;
        end
        else
          //FIN PT20-4
          CalculDuree(GetField('PCN_DATEDEBUTABS'), GetField('PCN_DATEFINABS'), Salarie,
            Etablissement, Getfield('PCN_TYPECONGE'), Nb_J, Nb_H, Nodjm, Nodjp);
        { PT- 7-8 Mise en commentaire : test reporté au dessus
        if ((getfield('PCN_DATEDEBUTABS') <> DDAbs) or(getfield('PCN_DATEFINABS') <> DFAbs)
        or (Nodjm<>DebDJ) or(Nodjp<>FinDJ)) then
          begin} *)
        { FIN PT79 }
        AffecteNodemj(getfield('PCN_DEBUTDJ'), DebDj);
        AffecteNodemj(getfield('PCN_FINDJ'), FinDj);
        DdAbs := getfield('PCN_DATEDEBUTABS');
        DfAbs := getfield('PCN_DATEFINABS');
        SetField('PCN_JOURS', Nb_J);
        SetField('PCN_HEURES', Nb_H);
        // end;
        AffectLibelle; //PT- 7-1 MAJ Titre Ecran if Action = 'CREATION' then
        // d PT49 IJSS calcul nbj calend, nbj carence, nbj IJSS
{ PT58       if ((F.Fieldname = 'PCN_DATEDEBUTABS') or
          (F.Fieldname = 'PCN_DATEFINABS')) and }
        if (assigned(T_MotifAbs)) and
          (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X') then MajIJSS;
        // f PT49 IJSS
      end;
    end;
  end;
  if f.fieldName = 'PCN_VALIDRESP' then
  begin
    if (getfield('PCN_VALIDRESP') <> OldValidResp) then
      AValider := True;
    if not FirstValid then
      if (((getfield('PCN_VALIDRESP') = 'REF') or (getfield('PCN_VALIDRESP') = 'NAN'))
        and (oldValidResp <> getfield('PCN_VALIDRESP'))) then
      begin
        IsMailRefOrAnu := False;
        Init := HShowMessage('1;Validation mouvement absence;Cette modification rendra le mouvement définivement inaccessible.#13#10 Voulez-vous continuer ?;Q;YN;N;N;', '', '');
        if Init <> mryes then
          SetField('PCN_VALIDRESP', oldvalidresp)
        else
          if (getfield('PCN_VALIDRESP') = 'REF') or (getfield('PCN_VALIDRESP') = 'NAN') then IsMailRefOrAnu := True;
      end;
    oldvalidresp := getfield('PCN_VALIDRESP');
    Firstvalid := false;
  end;
  {PT1-2 PH 12-06-2001
    if Origine = 'E' then
       enableChampOngletModeAgl else
       enableChampOngletMvtCp; }

  if (f.fieldName = 'PCN_NBREMOIS') then {PT- 6-1}
  begin
    if (GetField('PCN_NBREMOIS') <= 0) and (GetField('PCN_TYPEMVT')='CPA') and ((GetField('PCN_TYPECONGE') = 'REP') AND (GetField('PCN_TYPEIMPUTE') = 'AC2')) then { PT82 } { PT86 }
    begin
      PGIBox('Vous devez saisir un nombre de mois supérieur à zéro.', 'Nombre de mois :');
      SetField('PCN_NBREMOIS', 1);
    end;
  end;

  { DEB PT42-3 }
  if (f.fieldName = 'PCN_TYPEIMPUTE') and (GetField('PCN_TYPECONGE') = 'AJU') and (Origine = 'C') then
  begin
    SetControlEnabled('PCN_DATEVALIDITE', GetField('PCN_TYPEIMPUTE') <> 'REL');
    if (GetField('PCN_TYPEIMPUTE') = 'REL') and (Salarie <> '') and IsNumeric(GetField('PCN_PERIODECP')) then
    begin
      if GetField('PCN_PERIODECP') <= 0 then St := ' AND PCN_PERIODECP=1 '
      else St := ' AND PCN_PERIODECP=' + IntToStr(GetField('PCN_PERIODECP'));

      Q := OpenSql('SELECT PCN_DATEVALIDITE FROM ABSENCESALARIE ' +
        'WHERE PCN_TYPEMVT="CPA" AND PCN_SALARIE="' + Salarie + '" ' +
        'AND PCN_TYPECONGE="REL" ' + St, True);
      if not Q.eof then
        if (GetField('PCN_DATEVALIDITE')) <> (Q.FindField('PCN_DATEVALIDITE').AsDateTime) then
          SetField('PCN_DATEVALIDITE', Q.FindField('PCN_DATEVALIDITE').AsDateTime);
      Ferme(Q);
    end;
  end;
  { FIN PT42-3 }

// d PT49 IJSS
  if ((f.fieldName = 'PCN_NBJCARENCE') or (f.fieldName = 'PCN_NBJCALEND'))
    and Assigned(T_MotifAbs) then
    if (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X') then
    begin
      if GetField('PCN_NBJIJSS') <>  (Getfield('PCN_NBJCALEND') - getfield('PCN_NBJCARENCE')) then
       SetField('PCN_NBJIJSS',Getfield('PCN_NBJCALEND') - getfield('PCN_NBJCARENCE'));
// d PT61 FQ 12342
      if (Getfield('PCN_NBJIJSS') < 0) then   SetField('PCN_NBJIJSS',0);
// f PT61 FQ 12342
    end;
  // f PT49 IJSS

  if (F.FieldName = 'PCN_ETATPOSTPAIE') then  GestionOngletAnnulation; { PT76-1 }


end;



//------------------------------------------------------------------------------
//- détermine, pour l'onglet Congés Payés, la visibilité et les options de saisie pour
//- champs de cet onglet
//------------------------------------------------------------------------------

procedure TOM_AbsenceSalarie.ControleZonesOngletCongesPayes;
var
{$IFDEF EAGLCLIENT}
  EJours, EHeures, DateDebutcp, DateFincp, DateValidite: THEdit;
  DBSalarie: THEdit;
  TypeConge: THValComboBox;
{$ELSE}
  EJours, EHeures, DateDebutcp, DateFincp, DateValidite: THDBEdit;
  DBSalarie: THDBEdit;
  TypeConge: THDBValComboBox;
{$ENDIF}
  Q: TQuery;
  nbj: double;
  Init: word;
  st, Salarie, FinDj, DebutDj, StDate: string;
  Periode: integer;
  DateD, DateF {, Datesortie PT28 var globale}, DateDebAbs: tDateTime;
  TypeRtt: Boolean;
  YYD, MMD, JJ, YYF, MMF: WORD;
  //debut PT109
  Tob_Contrat : tob;
  i : integer;
  Danscontrat : boolean;
  //fin PT109
begin
{$IFDEF EAGLCLIENT}
  EJours := THEdit(getcontrol('PCN_JOURS'));
  EHeures := THEdit(getcontrol('PCN_HEURES'));
  DateDebutcp := THEdit(getcontrol('PCN_DATEDEBUTABS'));
  DateFincp := THEdit(getcontrol('PCN_DATEFINABS'));
  DateValidite := THEdit(getcontrol('PCN_DATEVALIDITE'));
  TypeConge := THValComboBox(getcontrol('PCN_TYPECONGE'));
  DBSalarie := THEdit(getcontrol('PCN_SALARIE'));
{$ELSE}
  EJours := THDBEdit(getcontrol('PCN_JOURS'));
  EHeures := THDBEdit(getcontrol('PCN_HEURES'));
  DateDebutcp := THDBEdit(getcontrol('PCN_DATEDEBUTABS'));
  DateFincp := THDBEdit(getcontrol('PCN_DATEFINABS'));
  DateValidite := THDBEdit(getcontrol('PCN_DATEVALIDITE'));
  TypeConge := THDBValComboBox(getcontrol('PCN_TYPECONGE'));
  DBSalarie := THDBEdit(getcontrol('PCN_SALARIE'));
{$ENDIF}
  if DBSalarie <> nil then
    Salarie := DBSalarie.text;
  st := 'select PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL' +
    ' FROM SALARIES WHERE PSA_SALARIE = "' + Salarie + '"';
  Q := opensql(st, true);
  if not Q.eof then
  begin
    if getfield('PCN_TRAVAILN1') = '' then SetField('PCN_TRAVAILN1', Q.findfield('PSA_TRAVAILN1').AsString);
    if getfield('PCN_TRAVAILN2') = '' then SetField('PCN_TRAVAILN2', Q.findfield('PSA_TRAVAILN2').AsString);
    if getfield('PCN_TRAVAILN3') = '' then SetField('PCN_TRAVAILN3', Q.findfield('PSA_TRAVAILN3').AsString);
    if getfield('PCN_TRAVAILN4') = '' then SetField('PCN_TRAVAILN4', Q.findfield('PSA_TRAVAILN4').AsString);
    if getfield('PCN_CODESTAT') = '' then SetField('PCN_CODESTAT', Q.findfield('PSA_CODESTAT').AsString);
    if getfield('PCN_CONFIDENTIEL') = '' then SetField('PCN_CONFIDENTIEL', Q.findfield('PSA_CONFIDENTIEL').AsString);
  end;
  Ferme(Q);
  Error := 0;
  ComplMessage := ''; //PT- 7-11
  { DEB PT37-1 }
  if Getfield('PCN_TYPEMVT') = 'CPA' then
  begin
    if VH_PAIE.PGCongesPayes = False then
    begin
      Error := 24;
      Setfocuscontrol('PCN_TYPECONGE');
      exit;
    end
    else
      if (CPEtab = False) and (Etablissement <> '') then { PT43-2 }
    begin
      Error := 25;
      Setfocuscontrol('PCN_TYPECONGE');
      exit;
    end
    else { DEB PT43-2 }
      if (CPSal = False) then
    begin
      Error := 29;
      Setfocuscontrol('PCN_TYPECONGE');
      exit;
    end; { FIN PT43-2 }
  end;
  { FIN PT37-1 }
  if (Getfield('PCN_SALARIE') = '') or (Nom = '') then
  begin
    Error := 26;
    Setfocuscontrol('PCN_SALARIE');
    exit;
  end;
  if Ejours <> nil then
  begin
    if isnumeric(EJours.text) then
    begin
      if (Valeur(EJours.text) <= 0) and (GetField('PCN_TYPECONGE') <> 'AJU') then //PT- 7-7 Si jour=0 alors message error //PT- 14 Ajout Cond. AJU
      begin
        Error := 3; //PT- 7-11 LastErrorMsg:='Le nombre de jours doit être positif' ;
        Setfocuscontrol('PCN_JOURS');
        exit;
      end;
    end
    else
    begin
      Error := 2; //PT- 7-11 LastErrorMsg:='Le nombre de jours est invalide' ;
      setfocuscontrol('PCN_JOURS');
      exit;
    end;
    if GetField('PCN_TYPEMVT') = 'ABS' then
    begin
      if not assigned(T_MotifAbs) then exit; //PT34
      { DEB PT80 }
      if not assigned (EJours) then exit;
      if not assigned (EHeures) then exit;
      If ControleGestionMaximum(Getfield('PCN_SALARIE'),Getfield('PCN_TYPECONGE'),T_MotifAbs,
                GetField('PCN_DATEDEBUTABS'),Valeur(EJours.text),Valeur(EHeures.text)) then
        Begin
        if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'JOU' then
           begin
           Error := 17;
           ComplMessage := floattostr(T_MotifAbs.GetValue('PMA_JRSMAXI'));
           exit;
          end
        else
        if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU' then
           begin
           Error := 18;
           ComplMessage := floattostr(T_MotifAbs.GetValue('PMA_JRSMAXI'));
           exit;
           end;
        End;
       { FIN PT80 }
     end;
  end;

  if ((Typeconge <> nil) and (EJours <> nil)) then
  begin
    if GetField('PCN_TYPECONGE') = '' then ////PT- 7-9 Control si typecongé renseigné PT39
    begin
      Error := 4; //PT- 7-11
      setfocuscontrol('PCN_TYPECONGE');
      exit;
    end;
    if (typeconge.value = 'CPA') AND (Getfield('PCN_TYPEMVT') = 'CPA') then   { PT69 }
    begin
       Setfield('PCN_DATEVALIDITE', getfield('PCN_DATEPAIEMENT'));
       Q := RechercheReprise(getfield('PCN_DATEVALIDITE'), 'CPA','AC2'); { PT86 }
       if not Q.eof then
         begin
         Q.First;
         while not Q.eof do
           begin
           if Q.FindField('PCN_ORDRE').asinteger <> getfield('PCN_ORDRE') then
             begin
             Error := 8; //PT- 7-11 LastErrorMsg:='Il ne peut y avoir 2 mvt de type reprise pris sur une même période CP' ;
             setfocuscontrol('PCN_TYPECONGE');
             Ferme(Q);
             exit;
             end;
           Q.next;
           end;
         end;
       Ferme(Q);
       Q := RechercheReprise(getfield('PCN_DATEVALIDITE'), 'REP','AC2'); { PT86 }
       if not Q.eof then
         nbj := Q.findfield('PCN_JOURS').asfloat
       else
         nbj := 0;
       if nbj < Valeur(Ejours.text) then
          begin
          Error := 5; //PT- 7-11 LastErrorMsg:='Le nombre de jours pris repris ne peut être supérieur au nombre de jours acquis repris ('+floattostr(nbj)+')' ;
          ComplMessage := floattostr(nbj);
          setfocuscontrol('PCN_JOURS');
          Ferme(Q);
          exit;
          end;
       Q.edit;
       Q.FindField('PCN_APAYES').Asfloat := Arrondi(Valeur(Ejours.text), DCP);
       Q.Post;
       Ferme(Q);
    end;
    if (typeconge.value = 'REP') AND (Getfield('PCN_TYPEMVT') = 'CPA') then     { PT69 }
    begin
      Q := RechercheReprise(getfield('PCN_DATEVALIDITE'), 'REP', GetField('PCN_TYPEIMPUTE')); { PT86 }
      if not Q.eof then
      begin
        Q.First;
        if getfield('PCN_ORDRE') <> Q.FindField('PCN_ORDRE').asinteger then
        begin
          Error := 6; //PT- 7-11 LastErrorMsg:='Il ne peut y avoir 2 mvt de type reprise acquis sur une même période CP' ;
          setfocuscontrol('PCN_TYPECONGE');
          Ferme(Q);
          exit;
        end;
        Ferme(Q);
      end;
    end;
    //DEB PT- 14
    if (typeconge.value = 'AJU') AND (Getfield('PCN_TYPEMVT') = 'CPA') then     { PT69 }
    begin
      if (Isnumeric(Ejours.text)) and (GetField('PCN_BASE') <> null) and (GetField('PCN_MOIS') <> null) then
      begin
        if (Valeur(Ejours.text) = 0) and (GetField('PCN_BASE') = 0) and (GetField('PCN_NBREMOIS') = 0) then
        begin
          Error := 19;
          setfocuscontrol('PCN_BASE');
          exit;
        end;
      end
      else
      begin
        Error := 19;
        setfocuscontrol('PCN_BASE');
        exit;
      end;
    end;
    //FIN PT- 14
  {PT- 7-2 Mise en commentaire Autorise la saisie par anticipation
           Voir ancienne version}
    if Assigned(T_MotifAbs) then TypeRtt := (T_MotifAbs.GetValue('PMA_TYPEABS') = 'RTT') //PT24-2 Gestion des motifs de type RTT { PT45 }
    else TypeRtt := False;
    if (Origine = 'E') and ((TYPECONGE.value = 'PRI') or (TypeRtt)) then
    begin
      //DEB PT20-5 Contrôle saisi absence pas + antérieur d'un mois
      DateDebAbs := GetField('PCN_DATEDEBUTABS');
      if (PlusMois(VH_Paie.PGECabDateIntegration, -1) > DateDebAbs) and (FicPrec <> 'ADM') then {PT30 DateModifRecap}
      begin
        Error := 20;
        Exit;
      end;
      //FIN PT20-5
      ControleAssezRestant(T_Recap,{Salarie,} TYPECONGE.value,Action, Valeur(Ejours.text),EJoursPris);  { PT92 }
    end;
  end;
  if ((TypeConge <> nil) and (Ejours <> nil) and (DateDebutcp <> nil)
    and (DateFincp <> nil) and (DateValidite <> nil)) then
  begin
    if Typeconge.Value = '' then
    begin
      Error := 9; //PT- 7-11 LastErrorMsg:='Type de mouvement obligatoire' ;
      exit;
    end;
    if getfield('PCN_DATEDEBUTABS') = null then
    begin
      if Typeconge.value = 'PRI' then
      begin
        Error := 11; //PT- 7-11 LastErrorMsg:='Date début obligatoire' ;
        exit;
      end;
    end;
    if getfield('PCN_DATEFINABS') = null then
    begin
      if Typeconge.value = 'PRI' then
      begin
        Error := 12; //PT- 7-11 LastErrorMsg:='Date fin obligatoire' ;
        exit;
      end;
    end;
    if getfield('PCN_DATEVALIDITE') = null then
    begin
      Error := 13; //PT- 7-11 LastErrorMsg:='Date validité obligatoire' ;
      setfocuscontrol('PCN_DATEVALIDITE');
      exit;
    end;
    {DEB PT28 Mise en commentaire, DateEntree, Date Sortie chargé dans le OnLoadrecord
    On test par rapport à la date de debut et fin abs
    Datesortie := 0;
    st := 'SELECT PSA_DATESORTIE FROM SALARIES WHERE PSA_SALARIE = "'+getfield('PCN_SALARIE')+'"';
    Q := opensql(st, True);
    if not Q.eof then Datesortie := Q.findfield('PSA_DATESORTIE').asdatetime;
    Ferme(Q);}{ DEB PT35 }
    if ((GetField('PCN_TYPECONGE') = 'PRI') and (GetField('PCN_TYPEMVT') = 'CPA')) or (GetField('PCN_TYPEMVT') = 'ABS') then
    begin
      if ((getfield('PCN_DATEFINABS') <> idate1900) and (datesortie > idate1900) and (datesortie < getfield('PCN_DATEFINABS'))) then
      begin
        Error := 16; //PT- 7-11 LastErrorMsg:='La date de validité doit être antérieure à la date de sortie du salarié '+ datetostr(datesortie); ;
        exit;
      end;
      if ((getfield('PCN_DATEDEBUTABS') <> idate1900) and (dateEntree <> idate1900) and (dateEntree > getfield('PCN_DATEDEBUTABS'))) then
      begin
      //debut //PT109
      st := 'select PCI_debutcontrat, PCI_fincontrat from contrattravail ' +
      'where PCI_salarie = ' + getfield('PCN_salarie');
      Tob_Contrat := Tob.Create('Contrat',nil,-1);
      Tob_Contrat.LoadDetailDbFromSQL('Contrat',st) ;
      Danscontrat := false;
        If Tob_contrat <> Nil then
         begin
            i := 0;
            While (i <= Tob_Contrat.Detail.Count - 1) and (not Danscontrat) do
              begin
                If ((Tob_Contrat.Detail[i].Getvalue('PCI_debutcontrat') <= getfield('PCN_DATEDEBUTABS')) and (Tob_Contrat.Detail[i].Getvalue('PCI_fincontrat') >= getfield('PCN_DATEFINABS'))) then
                   Danscontrat := true
                else
                    i := i+1;
              end;
         end;
       If not Danscontrat  then
        begin //fin PT109
         Error := 23;
         exit;
        end;
      If Tob_contrat <> Nil then
        FREEANDNIL(Tob_contrat); //PT109
      end; { FIN PT35 }
    end;

    if  (GetField('PCN_DATEVALIDITE') <> null) and (GetField('PCN_TYPEMVT') = 'CPA') then   { PT63 }
      if GetField('PCN_DATEVALIDITE') > idate1900 then
      begin
        Periode := CalculPeriode(DTClot, getfield('PCN_DATEVALIDITE'));
        if Periode > 1 then
        begin
          Init := HShowMessage('1;Mouvements Congés payés;Attention, La date de validité de ce mouvement le rendra désormais inaccessible.#13#10Etes vous sûr de vouloir continuer ?;Q;YNC;N;N;', '', '');
          if Init = mrYes then
          else Error := 14; //PT- 7-11
        end;
        if periode = -9 then
        begin
          Error := 7;
          //PT- 7-11 LastErrorMsg:='Il est interdit de saisir des mouvements avec une date postérieure à la date de clôture + 1 an' ;
          exit;
        end;
      end;
  end;
  // contrôle du non chevauchement des périodes
  if ((TypeConge <> nil) and (DateDebutcp <> nil) and (DateFincp <> nil)) then
    //PT- 12 le field de PCN_TYPEMVT n'étant pas encore affecté on récupère la valeur de mvt
    if ((TypeConge.value = 'PRI') or (GetField('PCN_TYPEMVT') = 'ABS')) then
      Begin
      DateD := getfield('PCN_DATEDEBUTABS');
      DateF := getfield('PCN_DATEFINABS');
      FinDj := getfield('PCN_FINDJ');
      DebutDj := getfield('PCN_DEBUTDJ');
      if DateD <> DateF then    { DEB PT88 }
          StDate := 'OR (PCN_DATEDEBUTABS ="' + usdatetime(Datef) + '" AND PCN_DEBUTDJ = "MAT") '+
                    'OR (PCN_DATEFINABS ="' + usdatetime(Dated) + '" AND PCN_FINDJ = "PAM") ';
                                { FIN PT88 }


      St := PgGetClauseAbsAControler(T_MotifAbs,Tob_MotifAbs,GetField('PCN_TYPECONGE'));  { PT89 }

      St := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT,PCN_ORDRE ' +
         'FROM ABSENCESALARIE ' +
         'WHERE PCN_SALARIE = "' + Salarie + '" ' + St +
         'AND PCN_ETATPOSTPAIE <> "NAN" '+ { PT76-1 Ne pas tenir compte des mouvements annulés }
         'AND (((PCN_DATEDEBUTABS >"' + usdatetime(DateD) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(DateF) + '") ' +
         'OR (PCN_DATEFINABS >"' + usdatetime(Datef) + '" AND PCN_DATEDEBUTABS < "' + usdatetime(Dated) + '") ' +
         'OR (PCN_DATEFINABS <"' + usdatetime(Datef) + '" AND PCN_DATEFINABS > "' + usdatetime(DateD) + '"))' +
         'OR (PCN_DATEFINABS ="' + usdatetime(Datef) + '" AND PCN_FINDJ = "' + (FinDj) + '") ' +
         'OR (PCN_DATEDEBUTABS ="' + usdatetime(Dated) + '" AND PCN_DEBUTDJ = "' + (Debutdj) + '") ' +
         'OR (PCN_DATEFINABS ="' + usdatetime(Dated) + '" AND PCN_FINDJ = "' + Debutdj + '") ' +
         'OR (PCN_DATEDEBUTABS ="' + usdatetime(Datef) + '" AND PCN_DEBUTDJ = "' + Findj + '") '+
         StDate+')'; { PT88 }

       if Origine = 'E' then
           st := st + ' AND (PCN_VALIDRESP <> "NAN" AND PCN_VALIDRESP <> "REF")';
         Q := opensql(st, True);
         if not Q.eof then
           begin { DEB PT56-1 }
           if ((Action = 'CREATION') and (Q.RecordCount = 1))
           or ((Action = 'MODIFICATION') and (Q.RecordCount > 1))
           or not ((Action = 'MODIFICATION') and (Q.RecordCount = 1)
           and (GblOrdre = Q.FindField('PCN_ORDRE').AsInteger)
           and (GblEnterTypeMvt = Q.FindField('PCN_TYPEMVT').AsString)) then
             begin
             Error := 15;
             Ferme(Q);
             exit;
             end; { FIN PT56-1 }
           end;
         ferme(Q);
      end;

  { DEB PT42-3 }
  if (GetField('PCN_TYPEMVT') = 'CPA') and ((Getfield('PCN_TYPECONGE') = 'AJP')
    or (Getfield('PCN_TYPECONGE') = 'AJU')
    or (Getfield('PCN_TYPECONGE') = 'REP')
    or (Getfield('PCN_TYPECONGE') = 'CPA')) then
  begin
    St := 'SELECT MAX(PCN_DATEVALIDITE) DATEVAL FROM ABSENCESALARIE ' +
      'WHERE PCN_SALARIE = "' + Salarie + '" AND PCN_TYPEMVT="CPA" ' +
      'AND PCN_TYPECONGE="SLD" AND PCN_GENERECLOTURE="-" ';
    Q := opensql(st, True);
    if not Q.eof then
      if GetField('PCN_DATEVALIDITE') <= Q.FindField('DATEVAL').AsDateTime then
      begin
        Error := 28;
        ComplMessage := DateToStr(Q.FindField('DATEVAL').AsDateTime);
        Ferme(Q);
        exit;
      end;
    Ferme(Q);
  end;
  { FIN PT42-3 }

  { DEB PT83 Contôle absence à cheval }
  If not VH_Paie.PGAbsenceCheval then
    if (GetField('PCN_TYPEMVT') = 'ABS') OR (Getfield('PCN_TYPECONGE') = 'PRI') Then
      Begin
      DecodeDate(GetField('PCN_DATEDEBUTABS'),YYD,MMD,JJ);
      DecodeDate(GetField('PCN_DATEFINABS'),YYF,MMF,JJ);
      IF (MMD <> MMF) OR (YYD <> YYF) then
        Begin
        Error := 30;
        exit;
        End;
      End;  
{ FIN PT83 }
end;


function TOM_AbsenceSalarie.RechercheReprise(Validite: Tdatetime; TypeConge,TypeImpute : string): Tquery; { PT86 }
var
  Q: TQuery;
  DtFin, DTDeb: TDateTime;
  st: string;
begin
  RechercheExerciceCp(Validite, DTdeb, DtFin);
  st := 'SELECT PCN_ORDRE,PCN_JOURS,PCN_APAYES FROM ABSENCESALARIE '+
        'WHERE PCN_SALARIE ="' + SALARIE+ '" AND PCN_TYPEMVT="CPA" '+     { PT82 }
        'AND PCN_TYPECONGE = "' + TypeConge + '" AND PCN_DATEVALIDITE >= "' + usdatetime(DtDeb) +
    '" AND PCN_DATEVALIDITE <= "' + usdatetime(DTfin) + '" '+
    'AND PCN_TYPEIMPUTE="'+TypeImpute+'" '; { PT86 }
  Q := OpenSQL(st, false);
  result := Q
end;

procedure TOM_AbsenceSalarie.RechercheExerciceCp(Validite: tdatetime; var DTdeb, DtFin: tdatetime);
var
  aa, mm, jj: word;
  i: integer;
begin

  DtDeb := 0;
  DTfin := 0;
  if Dtclot = 0 then exit;
  decodedate(Dtclot, aa, mm, jj);
  Dtdeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT41 }
  DTfin := DTclot;
  i := 0;
  while i < 10 do
  begin
    if ((Validite >= DTDeb) and (Validite <= DtFin)) then exit;
    DtFin := DtDeb - 1;
    decodedate(DtFin, aa, mm, jj);
    Dtdeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT41 }
    i := i + 1;
  end;
end;

procedure TOM_AbsenceSalarie.InitialiseZonesCongesPayes;
begin
  SetField('PCN_LIBELLE', '');
  SetField('PCN_DATEMODIF', Date);
  BDateEnter := True; //PT- 7-8 pour calcul duree en chargement de fiche
  SetField('PCN_HEURES', 0);
  SetField('PCN_TRAVAILN1', '');
  SetField('PCN_TRAVAILN2', '');
  SetField('PCN_TRAVAILN3', '');
  SetField('PCN_TRAVAILN4', '');
  if (Origine = 'P') then //Deb PT95
  begin
    SetField('PCN_DATEDEBUTABS', Idate1900);
    SetField('PCN_DATEFINABS', Idate1900);
    SetField('PCN_HDEB', StrToDateTime('00:00')); //PT100
    SetField('PCN_HFIN', StrToDateTime('00:00')); //PT100
    SetControlText('HDEB', '00:00');
    SetControlText('HFIN', '00:00');
    //On met une valeur par défaut aux champs non utilisés pour éviter les null dans la table
    // Debut de l'initialisation des champs non utilisés
    SetField('PCN_DATEDEBUT', Idate1900);
    SetField('PCN_DATEFIN', Idate1900);
    SetField('PCN_MVTPRIS', '');
    SetField('PCN_PERIODECP', 0);
    SetField('PCN_PERIODEPY', 0);
    SetField('PCN_SENSABS', '');
    //SetField('PCN_MVTORIGINE', '');
    SetField('PCN_MVTORIGINE', 'SAL');
    SetField('PCN_TYPEIMPUTE', '');
    SetField('PCN_GENERECLOTURE', '-');
    SetField('PCN_DATESOLDE', Idate1900);
    SetField('PCN_DATEVALIDITE', Idate1900);
    setField('PCN_DEBUTDJ', '');
    setField('PCN_FINDJ', '');
    SetField('PCN_DATEPAIEMENT', Idate1900);
    SetField('PCN_DATEANNULATION', Idate1900);
    setField('PCN_MOTIFANNUL', '');
    setField('PCN_ANNULEPAR', '');
    setField('PCN_ETATPOSTPAIE', 'VAL');
    SetField('PCN_CODETAPE', '...');
    //setField('PCN_CODETAPE', '');
    SetField('PCN_JOURS', 0);
    SetField('PCN_BASE', 0);
    SetField('PCN_NBREMOIS', 0);
    SetField('PCN_CODERGRPT', 0);
    SetField('PCN_MVTDUPLIQUE', '-');
    SetField('PCN_ABSENCE', 0);
    SetField('PCN_ABSENCEMANU', 0);
    SetField('PCN_MODIFABSENCE', '-');
    SetField('PCN_APAYES', 0);
    SetField('PCN_VALOX', 0);
    SetField('PCN_VALOMS', 0);
    SetField('PCN_VALORETENUE', 0);
    SetField('PCN_VALOMANUELLE', 0);
    SetField('PCN_MODIFVALO', '-');
    SetField('PCN_PERIODEPAIE', '');
    SetField('PCN_TOPCONVERT', '-');
    SetField('PCN_SAISIEDEPORTEE', '-');
   // SetField('PCN_VALIDSALARIE', '');
    SetField('PCN_VALIDSALARIE', 'SAL');
   // SetField('PCN_VALIDRESP', '');
    SetField('PCN_VALIDRESP', 'VAL');
   // SetField('PCN_EXPORTOK', '');
    SetField('PCN_EXPORTOK', 'X');
    SetField('PCN_LIBCOMPL1', '');
    SetField('PCN_LIBCOMPL2', '');
    SetField('PCN_VALIDABSENCE', '');
    SetField('PCN_OKFRACTION', '');
    SetField('PCN_NBJCARENCE', 0);
    SetField('PCN_NBJCALEND', 0);
    SetField('PCN_NBJIJSS', 0);
    SetField('PCN_IJSSSOLDEE', '-');
    SetField('PCN_GESTIONIJSS', '-');
    SetField('PCN_REPRISEARRET', Idate1900);
    SetField('PCN_DATEATTEST', Idate1900);
    if HEnt1.IsValidDate(ParamDF) then SetField('PCN_DATEVALIDITE', StrToDate(ParamDF));
    if HEnt1.IsValidDate(ParamDD) then SetField('PCN_DATEDEBUTABS', StrToDate(ParamDD));
    if HEnt1.IsValidDate(ParamDF) then SetField('PCN_DATEFINABS', StrToDate(ParamDF));
    // Fin de l'initialisation des champs non utilisés
  end else begin  //Fin PT95
    SetField('PCN_ETABLISSEMENT', Etablissement); //PT100
    SetField('PCN_PERIODEPY', -1);
    SetField('PCN_PERIODECP', -1);
    if (GetField('PCN_TYPECONGE') = 'PRI') then  SetField('PCN_PERIODECP', 0); { PT91-2 }
    if Origine = 'A' then SetField('PCN_SENSABS', '-');
    SetField('PCN_DATESOLDE', IDate1900); // PT50
    { DEB PT38 Si non absence et non conges payés pris alors idate1900 }
    if (GetField('PCN_TYPEMVT') <> 'CPA') or (GetField('PCN_TYPECONGE') = 'PRI') or (Origine = 'E') then
    begin { DEB PT55 }
      if not GblModifDate then SetField('PCN_DATEVALIDITE', Date);
      if not GblModifDate then SetField('PCN_DATEDEBUTABS', Date);
      if not GblModifDate then SetField('PCN_DATEFINABS', Date);
      if HEnt1.IsValidDate(ParamDF) then SetField('PCN_DATEVALIDITE', StrToDate(ParamDF));
      if HEnt1.IsValidDate(ParamDD) then SetField('PCN_DATEDEBUTABS', StrToDate(ParamDD));
      if HEnt1.IsValidDate(ParamDF) then SetField('PCN_DATEFINABS', StrToDate(ParamDF));
    end { FIN PT55 }
    else
    begin
      SetField('PCN_DATEDEBUTABS', Idate1900);
      SetField('PCN_DATEFINABS', Idate1900);
    end;
    { FIN PT38 }
  {PORTAGECWAS
    SetField('PCN_DATEVALIDITE'  ,null);
    SetField('PCN_DATEDEBUTABS'  ,null);
    SetField('PCN_DATEFINABS'    ,null);}
    setField('PCN_DEBUTDJ', 'MAT');
    setField('PCN_FINDJ', 'PAM');
    if (GetField('PCN_TYPECONGE') = 'CPA') and (GetField('PCN_TYPEMVT') = 'CPA') then
      if (DateSortie <> idate1900) and (DateSortie < DtClot ) then      { PT72 }
      SetField('PCN_DATEPAIEMENT', DateSortie)                          { PT72 }
    else
      SetField('PCN_DATEPAIEMENT', DtClot) { PT42-2 }
    else
      SetField('PCN_DATEPAIEMENT', IDate1900); // PT50
    SetField('PCN_CODETAPE', '...'); //PT- 13
    SetField('PCN_JOURS', 0);
    SetField('PCN_BASE', 0);
    { DEB PT43-1 }
    if (GetField('PCN_TYPECONGE') = 'REP') and (GetField('PCN_TYPEMVT') = 'CPA') AND (GetField('PCN_TYPEIMPUTE') = 'AC2') then { PT86 }
      SetField('PCN_NBREMOIS', 1) {PT- 6-1}
    else
      SetField('PCN_NBREMOIS', 0);
    { FIN PT43-1 }
    SetField('PCN_MVTDUPLIQUE', '-');
    SetField('PCN_ABSENCE', 0);
    SetField('PCN_ABSENCEMANU', 0);
    SetField('PCN_MODIFABSENCE', '-');
    SetField('PCN_APAYES', 0);
    SetField('PCN_VALOX', 0);
    SetField('PCN_VALOMS', 0);
    SetField('PCN_VALORETENUE', 0);
    SetField('PCN_VALOMANUELLE', 0);
    SetField('PCN_MODIFMANU', '-');
    SetField('PCN_PERIODEPAIE', '');
    SetField('PCN_CODERGRPT', getfield('PCN_ORDRE'));
    SetField('PCN_TYPEIMPUTE', ''); { PT86 }
    if (GetField('PCN_TYPECONGE') = 'AJU') and (GetField('PCN_TYPEMVT') = 'CPA') then
      SetField('PCN_TYPEIMPUTE', 'ACQ'); { PT42-3 }

    if ((GetField('PCN_TYPECONGE') = 'REP') OR (GetField('PCN_TYPECONGE') = 'CPA')) and (GetField('PCN_TYPEMVT') = 'CPA') then { PT86 }
      SetField('PCN_TYPEIMPUTE', 'AC2'); { PT86 }

    if Origine = 'E' then
    begin
      //         SetField('PCN_TYPEMVT'          ,'ABS');
      SetField('PCN_SENSABS', '-');
      SetField('PCN_SAISIEDEPORTEE', 'X');
      SetField('PCN_VALIDSALARIE', Copy(FicPrec, 1, 3)); //PT20-2 Utilisation de PCN_VALIDSALARIE pour toper saisie resp ou sal
      SetField('PCN_VALIDRESP', 'ATT');
      SetField('PCN_EXPORTOK', '-');
    end
    else
    begin //PT- 7-4 On force la validation resp et exportOK pour saisie non eagl
      SetField('PCN_VALIDSALARIE', 'SAL'); //PT20-2
      SetField('PCN_VALIDRESP', 'VAL');
      SetField('PCN_EXPORTOK', 'X');
      SetField('PCN_SAISIEDEPORTEE', '-');
    end;
  end;
end;

procedure TOM_AbsenceSalarie.enableChampOngletModeAgl;
var
  enabledsaisok : boolean;
  Q : TQuery;
  St, LeMotif : String;
begin
  enabledsaisok := False;
  if (ficprec = 'SAL') then
    enabledsaisok := (getfield('PCN_VALIDRESP') = 'ATT') and (getfield('PCN_EXPORTOK') <> 'X');
  if ficprec = 'RESP' then
    enabledsaisok := (getfield('PCN_EXPORTOK') <> 'X') and (getfield('PCN_VALIDRESP') <> 'REF') and (getfield('PCN_VALIDRESP') <> 'NAN') and (GetField('PCN_VALIDSALARIE') <> 'SAL');
  if ficprec = 'ADM' then //PT- 18
    enabledsaisok := (getfield('PCN_EXPORTOK') <> 'X');
  //DEB PT20-2 Mise en commentaire de "and (Ficprec='SAL')"  pour création resp et adm
  SetcontrolEnabled('BDELETE', (enabledsaisok)); //and (Ficprec='SAL')
  SetControlEnabled('PCN_TYPECONGE', (enabledsaisok)); //and (Ficprec='SAL')
  SetControlEnabled('PCN_DATEDEBUTABS', (enabledsaisok)); //and (Ficprec='SAL')
  SetControlEnabled('PCN_DATEFINABS', (enabledsaisok)); //and (Ficprec='SAL'));
  SetControlEnabled('PCN_DEBUTDJ', (enabledsaisok)); //and (Ficprec='SAL'));
  SetControlEnabled('PCN_FINDJ', (enabledsaisok)); //and (Ficprec='SAL'));
  SetControlEnabled('PCN_JOURS', (enabledsaisok)); //and (Ficprec='SAL'));
  SetControlEnabled('PCN_LIBCOMPL1', (enabledsaisok));
  SetControlEnabled('PCN_LIBCOMPL2', (enabledsaisok));
  //FIN PT20-2
  SetControlEnabled('PCN_VALIDRESP', (((FicPrec = 'RESP') or (FicPrec = 'ADM')) and (getfield('PCN_EXPORTOK') = '-'))); // and (enabledsaisok) //PT- 17 ajout OR (FicPrec='ADM')
  //SetControlEnabled('LVALIDRESP'          , (FicPrec='RESP')and (enabledsaisok));  PT- 7-10
  //SetControlEnabled('LLIBELLE'            , (FicPrec='RESP')and (enabledsaisok));  PT- 7-10
  SetControlEnabled('PCN_LIBELLE', False); //(FicPrec='RESP')and (enabledsaisok));
  //SetControlEnabled('LJOURS'              , (FicPrec='RESP')and (enabledsaisok));  PT- 7-10
  SetControlEnabled('PCN_JOURS', False); //(FicPrec='RESP')and (enabledsaisok));
  LeMotif := GetField ('PCN_TYPECONGE');
  st := 'SELECT * FROM MOTIFABSENCE WHERE PMA_MOTIFABSENCE="'+LeMotif+'" AND ##PMA_PREDEFINI##';
  Q := OpenSql (St, true);
  if NOT Q.EOF then
  begin
    if (Q.FindField ('PMA_TYPEABS').AsString = 'DEG') AND  (Q.FindField ('PMA_JOURHEURE').AsString = 'HEU') then
    begin
    SetControlEnabled('PCN_JOURS', true);
    SetControlEnabled('PCN_HEURES', true);
    end;
  end;
  FERME (Q);
  SetControlEnabled('BDELETE', (enabledsaisok)); //not(Ficprec='RESP')and

  if (Origine = 'E') or (Origine = 'P') then  // pt95
  begin
    SetcontrolEnabled('BDELETE', (IfMotifabsenceSaisissable(GetField('PCN_TYPECONGE'), Ficprec))); //PT39
    { DEB PT62 }
    if ficprec = 'SAL' then
      SetControlVisible('BDELETE', (getfield('PCN_VALIDRESP') = 'ATT') and (getfield('PCN_EXPORTOK') <> 'X') AND (GetField('PCN_VALIDSALARIE') = 'SAL'));
    if ficprec = 'RESP' then
      SetControlVisible('BDELETE', (getfield('PCN_EXPORTOK') <> 'X') and (getfield('PCN_VALIDRESP') <> 'REF') and (getfield('PCN_VALIDRESP') <> 'NAN') and (GetField('PCN_VALIDSALARIE') <> 'SAL'));
    if ficprec = 'ADM' then
      SetControlVisible('BDELETE', (getfield('PCN_EXPORTOK') <> 'X'));
    { FIN PT62 }
    if HEnt1.IsValidDate(ParamDD) then SetControlEnabled('PCN_DATEDEBUTABS', False);
    if HEnt1.IsValidDate(ParamDF) then SetControlEnabled('PCN_DATEFINABS', False); //and (Ficprec='SAL'));
    if HEnt1.IsValidDate(ParamDD) then SetControlEnabled('PCN_SALARIE', False);  { PT65 }
  end;

{$IFDEF EAGLCLIENT} //PT40-2
  if Origine <> 'E' then SetControlVisible('BDELETE', enabledsaisok);
{$ENDIF}

  {DEB PT- 7-10 Champvisible champEnabled }
  if (Ficprec = 'SAL') and (getfield('PCN_VALIDRESP') = 'ATT') and (ACTION = 'MODIFICATION') then
  begin
    SetControlEnabled('PCN_TYPECONGE', False);
    SetControlEnabled('BDELETE', True);
    SetControlEnabled('BDEFAIRE', True);
  end;
  {FIN PT- 7-10 Champvisible champEnabled }
end;

procedure TOM_AbsenceSalarie.enableChampOngletMvtCp;
var
{$IFDEF EAGLCLIENT}
  MtypeConge: THValComboBox;
{$ELSE}
  MtypeConge: THDBValComboBox;
{$ENDIF}
// PT104  BPRI, BACA, BACQ, BACS, BAJU, BAJP, BARR, BREL, BCPA, BREP, BSLD, PAYE, BPER, BABS, CONS: boolean;
  BPRI, BACA, BACQ, BACS, BACF, BAJU, BAJP, BARR, BREL, BCPA, BREP, BSLD, PAYE, BPER, BABS, CONS: boolean;
  Lvalidite, Libprispaye: thlabel;
  Periode: integer;
begin
  if ACTION = 'MODIFICATION' then {DEB PT- 8}
  begin
    PaieLectureSeule(TFFiche(Ecran), False);
    // d  PT49 IJSS
    if getcontrol('ENTREE') <> nil then     //PT96
      setcontrolenabled('ENTREE', False);      { PT72 }
    setcontrolenabled('SORTIE', False);      { PT72 }
    if getcontrol('PCN_SALARIE1') <> nil then //PT96
      setcontrolenabled('PCN_SALARIE1', False);
    if getcontrol('PCN_LIBELLE1') <> nil then  //PT96
      setcontrolenabled('PCN_LIBELLE1', False);
    if (Origine <> 'I') then
    begin
      setcontrolenabled('PCN_NBJCALEND', False);
      setcontrolenabled('PCN_NBJCARENCE', True);
      setcontrolenabled('PCN_NBJIJSS', False);
      setcontrolenabled('PCN_IJSSSOLDEE', False);
    end;
    // f PT49 IJSS
    SetControlEnabled('PCN_SALARIE', False);
    SetControlEnabled('PCN_CODETAPE', False);
  end
  else
    if (ACTION = 'CREATION') and (Origine = 'A') then
  begin
    //if (salarie <> '') then SetControlEnabled('PCN_SALARIE', False);   { PT56-1 }
    if getcontrol('PCN_SALARIE1') <> nil then //PT96
      setcontrolenabled('PCN_SALARIE1', False); // PT49 IJSS

    SetControlEnabled('PCN_TYPECONGE', True);
    SetControlEnabled('PCN_SENSABS', True);
    SetControlEnabled('PCN_DATEDEBUTABS', True);
    SetControlEnabled('PCN_DEBUTDJ', True);
    SetControlEnabled('PCN_DATEFINABS', True);
    SetControlEnabled('PCN_FINDJ', True);
    SetControlEnabled('PCN_JOURS', True);
    SetControlEnabled('PCN_HEURES', True);
    SetControlEnabled('PCN_DATEPAIEMENT', True);
    SetControlEnabled('PCN_DATEVALIDITE', True);
    SetControlEnabled('PCN_LIBELLE', True);
   end;
  {FIN PT- 8}

  if (Origine <> 'P') then  // PT95
  begin

  // 03-04 je remplace partout BREB par { or BREB} or BAJU fusion mvts REB avec AJP et AJU
    BPRI := False;
    BACA := False;
    BACQ := False;
    BACS := False;
    BACF := False;     // PT104
    BAJU := False;
    BAJP := False;
    BARR := False;
    BREL := False; {BREG := False;BREB := False;}
    BCPA := False;
    BREP := False;
    BSLD := False;
    BABS := False;
  {$IFDEF EAGLCLIENT}
    MTypeConge := THValComboBox(getcontrol('PCN_TYPECONGE'));
  {$ELSE}
    MTypeConge := THDBValComboBox(getcontrol('PCN_TYPECONGE'));
  {$ENDIF}
    if MtypeConge = nil then exit;
    if MTypeConge.value = 'PRI' then
      BPRI := True;
    PAYE := ((getfield('PCN_CODETAPE') = 'P') or (getfield('PCN_CODETAPE') = 'C')) and (MTypeConge.value <> 'CPA'); //PT21 Code erroné : CODEETAPE='S'
    if GetField('PCN_TYPEMVT') = 'ABS' then
      BABS := True;
    CONS := (GetField('PCN_APAYES') > 0); { PT42-3 }
    if MTypeConge <> nil then
    begin  { DEB PT82 Ajout clause TYPEMVT }
      if (MTypeConge.value = 'ACA') AND (GetField('PCN_TYPEMVT')='CPA') then BACA := True;
      if (MTypeConge.value = 'ACQ') AND (GetField('PCN_TYPEMVT')='CPA') then BACQ := True;
      if (MTypeConge.value = 'ACS') AND (GetField('PCN_TYPEMVT')='CPA') then BACS := True;
      if (MTypeConge.value = 'ACF') AND (GetField('PCN_TYPEMVT')='CPA') then BACF := True;// PT104
      if (MTypeConge.value = 'AJU') AND (GetField('PCN_TYPEMVT')='CPA') then BAJU := True;
      if (MTypeConge.value = 'AJP') AND (GetField('PCN_TYPEMVT')='CPA') then BAJP := True;
      if (MTypeConge.value = 'ARR') AND (GetField('PCN_TYPEMVT')='CPA') then BARR := True;
      if (MTypeConge.value = 'REL') AND (GetField('PCN_TYPEMVT')='CPA') then BREL := True;
      if (MTypeConge.value = 'CPA') AND (GetField('PCN_TYPEMVT')='CPA') then BCPA := True;
      if (MTypeConge.value = 'REP') AND (GetField('PCN_TYPEMVT')='CPA') then BREP := True;
      if (MTypeConge.value = 'SLD') AND (GetField('PCN_TYPEMVT')='CPA') then BSLD := True;
     { FIN PT82 }
      if Origine = 'E' then
      begin
        SetControlEnabled('PCN_JOURS', false);
        SetControlEnabled('PCN_HEURES', false);
        SetControlEnabled('PCN_DATEVALIDITE', false);
        //    SetControlEnabled('LJOURS'          , false); PT- 7-10
        //    SetControlEnabled('LLIBELLE'        , false); PT- 7-10
        SetControlEnabled('LVALIDITE', false);
        SetcontrolEnabled('BDELETE', (Action <> 'CREATION'));
        SetcontrolEnabled('BDEFAIRE', false);
      end;
      SetcontrolEnabled('BCUM', Origine = 'E');
      SetcontrolVisible('BCUM', Origine = 'E');
      if (Origine = 'C') then
      begin
        Lvalidite := thlabel(getcontrol('LVALIDITE'));
        if Lvalidite <> nil then
        begin
          if BPRI then LValidite.caption := TraduireMemoire('A payer à partir du ')
          else LValidite.caption := TraduireMemoire('Date de validité');
        end;

        if getfield('PCN_DATEVALIDITE') <> null then
        begin
          Periode := CalculPeriode(DtClot, getfield('PCN_DATEVALIDITE'));
          if Periode > 1 then
            BPER := False
          else BPER := True;
        end
        else
          BPER := True;
// d PT104
        SetControlVisible('PCN_BASE', BACQ or BACA or BACS or BACF or BREL or BREP or BCPA { or BREB} or BAJU);
        SetControlVisible('LBASE', BACQ or BACA or BACS or BACF or BREL or BREP or BCPA { or BREB} or BAJU);
        SetControlVisible('PCN_NBREMOIS', BACQ or BACA or BACS or BACF  or BREL or BREP { or BREB} or BAJU);
        SetControlVisible('LNBMOIS', BACQ or BACA or BACS or BACF or BREL or BREP { or BREB} or BAJU);

        SetControlVisible('PCN_APAYES', (BACQ or BACA or BACS or BACF or BREL or BREP { or BREB} or BAJU) and (PGAcces)); { PT42-1 }
        SetControlVisible('LAPAYES', (BACQ or BACA or BACS or BACF or BREL or BREP { or BREB} or BAJU) and (PGAcces)); { PT42-1 }
        {Deb PT- 8}
        SetControlVisible('PCN_PERIODECP', (BACQ or BACA or BACS or BACF or BREL or BREP { or BREB} or BAJU) and (PGAcces)); { PT42-1 }
        SetControlVisible('TPCN_PERIODECP', (BACQ or BACA or BACS or BACF or BREL or BREP { or BREB} or BAJU) and (PGAcces)); { PT42-1 }
        SetControlEnabled('PCN_CODETAPE', (PGAcces)); { PT42-1 }
        {Fin PT- 8}
// f PT104

        SetControlVisible('PCN_DATEVALIDITE', (not BCPA));
        SetControlVisible('LVALIDITE', (not BCPA));
        SetControlVisible('PCN_DATEPAIEMENT', BCPA);
        SetControlEnabled('PCN_DATEPAIEMENT', (BCPA) and (Action = 'CREATION'));
        SetControlVisible('LPAIEMENT', BCPA);
        SetControlVisible('PCN_JOURS', {not BREB} true);
        SetControlVisible('LJOURS', {not BREB} true);

        SetControlVisible('PCN_DATEDEBUTABS', BPRI);
        SetControlVisible('PCN_DEBUTDJ', BPRI);
        SetControlVisible('LDEBUTCP', BPRI);
        SetControlVisible('PCN_DATEFINABS', BPRI);
        SetControlVisible('PCN_FINDJ', BPRI);
        SetControlVisible('LFINCP', BPRI);

// d PT104
        SetControlEnabled('PCN_BASE', (((BCPA { or BREB} or BAJU or not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD))) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_NBREMOIS', ({ or BREB}(BAJU or not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD)) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_TYPECONGE', ((not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD)) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_SENSABS', ({ or BREB}BAJP or BAJU or not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD or BPRI or BCPA or BREP)) and BPER);
        SetControlEnabled('PCN_JOURS', ((not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD { or BREB})) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_HEURES', ((not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD { or BREB})) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_DATEVALIDITE', ((not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD)) and BPER) and not ((BCPA or BREP) and (Action <> 'CREATION')));
        SetControlEnabled('PCN_LIBELLE', (not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD or (PAYE))) and BPER);
// f PT104
        SetControlEnabled('PCN_DATEDEBUTABS', BPER);
        SetControlEnabled('PCN_DEBUTDJ', BPER);
        SetControlEnabled('PCN_DATEFINABS', BPER);
        SetControlEnabled('PCN_FINDJ', BPER);

        SetControlVisible('PCN_TYPEIMPUTE', (BAJU OR BREP OR BCPA)); { PT42-3 } { PT86 }
        SetControlVisible('TPCN_TYPEIMPUTE', (BAJU OR BREP OR BCPA)); { PT42-3 } { PT86 }
        SetControlEnabled('PCN_TYPEIMPUTE', (BAJU OR BREP)); { PT42-3 } { PT86 }
        if (not BABS) AND ((BAJU) OR (BREP)) then SetControlproperty('PCN_TYPEIMPUTE','Plus',' AND CO_LIBRE Like "%'+GetField('PCN_TYPECONGE')+'%" '); { PT86 }
        if (not BABS) AND  (BCPA) then SetControlproperty('PCN_TYPEIMPUTE','Plus',' AND CO_LIBRE Like "%REP%" '); { PT86 }
      end;

      if ((BPRI or BABS) and PAYE) then
      begin
        SetControlEnabled('PCN_TYPECONGE', not (PAYE));
        SetControlEnabled('PCN_SENSABS', not (PAYE));
        SetControlEnabled('PCN_DATEDEBUTABS', not (PAYE));
        SetControlEnabled('PCN_DEBUTDJ', not (PAYE));
        SetControlEnabled('PCN_DATEFINABS', not (PAYE));
        SetControlEnabled('PCN_FINDJ', not (PAYE));
        SetControlEnabled('PCN_JOURS', not (PAYE));
        SetControlEnabled('PCN_HEURES', not (PAYE));
        SetControlEnabled('PCN_DATEPAIEMENT', not (PAYE));
        SetControlEnabled('PCN_DATEVALIDITE', not (PAYE));
        SetControlEnabled('PCN_LIBELLE', not (PAYE));
      end;
      { DEB PT42-3 }
      if (BAJU or BAJP) and (CONS) then
      begin
        SetControlEnabled('PCN_BASE', False);
        SetControlEnabled('PCN_NBREMOIS', False);
        SetControlEnabled('PCN_TYPECONGE', False);
        SetControlEnabled('PCN_SENSABS', False);
        SetControlEnabled('PCN_JOURS', False);
        SetControlEnabled('PCN_DATEVALIDITE', False);
        SetControlEnabled('PCN_DATEDEBUT', False);
        SetControlEnabled('PCN_DATEFIN', False);
        SetControlEnabled('PCN_TYPEIMPUTE', False);
      end;
      { FIN PT42-3 }

// d PT104
      SetControlEnabled('BDELETE', (not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD or PAYE)) and not ((BCPA or BREP) and (Action <> 'CREATION')));
  {$IFDEF EAGLCLIENT} //PT40-2
      if Origine <> 'E' then SetControlVisible('BDELETE', (not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD or PAYE)) and not ((BCPA or BREP) and (Action <> 'CREATION')));
  {$ENDIF}
      SetControlEnabled('BDEFAIRE', (not (BACQ or BACA or BACS or BACF or BREL or BARR or BSLD or PAYE)) and not ((BCPA or BREP) and (Action <> 'CREATION')));
//f PT104
    end;
    SetControlVisible('LPRISPAYE', ((BPRI or BABS) and PAYE));
    SetControlEnabled('BINDEMNITE', ((BPRI or BSLD) and PAYE)); {PT- 5 : ajout BSLD} {PT- 8 retrait or BABS}
    if ((BABS) and PAYE) then
    begin
      Libprispaye := THlabel(getcontrol('LPRISPAYE'));
      if Libprispaye <> nil then
      begin
        if GetField('PCN_TYPEMVT') = 'ABS' then // PT-7 ajout TypeMvt='ABS' et else SOUAD
          Libprispaye.caption := TraduireMemoire('Mouvement intégré à la paye du ') + datetostr(getfield('PCN_DATEPAIEMENT'))
        else
          Libprispaye.caption := TraduireMemoire('Mouvement Pris et Payé : non modifiable');
      end;
    end;
    {DEB PT- 4}
    if PGAcces = True then { PT42-1 }
    begin
      SetControlEnabled('BDELETE', True);
      SetControlEnabled('BDEFAIRE', True);
      SetControlEnabled('PCN_JOURS', True);
      SetControlEnabled('PCN_APAYES', True);
      SetControlVisible('PCN_APAYES', True);
      SetControlEnabled('LAPAYES', True);
      SetControlVisible('LAPAYES', True);
      SetControlEnabled('PCN_DATEVALIDITE', True);
      SetControlEnabled('PCN_PERIODECP', True); //DEB PT- 8 Ajout ligne
      SetControlVisible('PCN_PERIODECP', True);
      SetControlEnabled('TPCN_PERIODECP', True);
      SetControlVisible('TPCN_PERIODECP', True); //FIN PT- 8 Ajout ligne
      SetControlEnabled('PCN_BASE', True);
      SetControlEnabled('PCN_ABSENCE', True);
      SetControlEnabled('PCN_NBREMOIS', True);
    end;
    {FIN PT- 4}
    SetControlVisible('TPCN_DATEDEBUT', BAJP);
    SetControlVisible('PCN_DATEDEBUT', BAJP);
    SetControlVisible('TPCN_DATEFIN', BAJP);
    SetControlVisible('PCN_DATEFIN', BAJP);

    if (Origine = 'A') or (Origine = 'ABSANNUL') then //PT24-3 Ajout Cond.  { PT76-1 }
    begin
      SetControlEnabled('PCN_SENSABS', (BAJP or not (BPRI or PAYE or BABS))); // and BPER
      SetControlProperty('PIJSS', 'Tabvisible', (not BPRI)); { PT56-2 }
      if Assigned(T_MotifAbs) then    { PT76-2 }
              SetControlProperty('PIJSS', 'Tabvisible',(T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X'))
    end;
    {DEB PT65}
    if HEnt1.IsValidDate(ParamDD) then SetControlEnabled('PCN_DATEDEBUTABS', False);
    if HEnt1.IsValidDate(ParamDF) then SetControlEnabled('PCN_DATEFINABS', False);
    if HEnt1.IsValidDate(ParamDD) then SetControlEnabled('PCN_SALARIE', False);
    {FIN PT65}

    GestionOngletAnnulation; { PT76-1 }
  end else begin //Fin de "if (Origine <> 'P') then  // PT95"
    if Assigned(T_MotifAbs) then
    begin
      if (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') then
      begin          //Gestion en heures
        SetControlCaption('TPCN_HEURES',TraduireMemoire('Nb heures'));
        SetControlVisible('TPCN_HDEB', True);
        SetControlVisible('HDEB', True);
        SetControlVisible('TPCN_HFIN', True);
        SetControlVisible('HFIN', True);
        //PT99 - Début
        If Origine = 'P' Then
        Begin
             SetControlCaption('TPCN_HEURES',TraduireMemoire('Nb heures totales'));
             SetControlVisible('PCN_NBHEURESNUIT', True);
             SetControlVisible('TPCN_NBHEURESNUIT', True);
        End;
        //PT99 - Fin
      end else begin //Gestion en quantités
        SetControlCaption('TPCN_HEURES',TraduireMemoire('Quantité'));
        SetControlVisible('TPCN_HDEB', False);
        SetControlVisible('HDEB', False);
        SetControlVisible('TPCN_HFIN', False);
        SetControlVisible('HFIN', False);
        //PT99 - Début
        If Origine = 'P' Then
        Begin
             SetControlVisible('PCN_NBHEURESNUIT',  False);
             SetControlVisible('TPCN_NBHEURESNUIT', False);
        End;
        //PT99 - Fin
      end;
      SetControlEnabled('PCN_DATEFINABS', False); //PT101
    end;
  end; //Fin du else de "if (Origine <> 'P') then  // PT95"
end;


procedure TOM_AbsenceSalarie.AffectLibelle;
var
  Titre: string;
begin
  if (GetField('PCN_TYPECONGE') = 'PRI') or (GetField('PCN_TYPEMVT') = 'ABS') then
  begin
    if GetField('PCN_TYPECONGE') = 'PRI' then Titre := 'Congés payés'
    else
      if GetField('PCN_TYPECONGE') <> '' then
      Titre := RechDom('PGMOTIFABSENCE', GetField('PCN_TYPECONGE'), False)
    else Titre := '';
    if (getfield('PCN_DATEDEBUTABS') <> null) and (getfield('PCN_DATEFINABS') <> null) and (titre <> '') then
    begin
      SetField('PCN_LIBELLE',RendLibAbsence(GetField('PCN_TYPECONGE'),GetField('PCN_DEBUTDJ'),GetField('PCN_FINDJ'),getfield('PCN_DATEDEBUTABS'),getfield('PCN_DATEFINABS'))); { PT92 }
      {if Length(Titre) > 14 then Titre := Copy(Titre, 1, 14); //PT- 19  Tronquage du libellé : recup des 14 premiers caractères
      titre := titre + ' ' + FormatDateTime('dd/mm/yy', getfield('PCN_DATEDEBUTABS')) + ' au ' + FormatDateTime('dd/mm/yy', getfield('PCN_DATEFINABS'));
      if GetField('PCN_LIBELLE') <> titre then
        SetField('PCN_LIBELLE', titre);}
    end
    else SetField('PCN_LIBELLE', '');
  end;
  if ORIGINE = 'P' then //Deb PT95
  begin
    if GetField('PCN_TYPECONGE') <> '' then
    begin
      //Faire une fonction RendLibPresence dans PGCalendrier
      Titre := RechDom('PGMOTIFPRESENCE', GetField('PCN_TYPECONGE'), False);
      SetField('PCN_LIBELLE',RendLibPresence(GetField('PCN_TYPECONGE'),getfield('PCN_DATEDEBUTABS'),getfield('PCN_DATEFINABS'))); { PT92 }
    end;
  end;  //Fin PT95
end;

procedure TOM_AbsenceSalarie.OnEnterSalarie(Sender: TObject);
begin
  SalarieEnter := GetControlText('PCN_SALARIE');
end;

procedure TOM_AbsenceSalarie.OnExitSalarie(Sender: TObject);
var
  Sal: string;
  Q: TQuery;
  StWhere :String;
begin
  Sal := GetControlText('PCN_SALARIE'); {DEB PT- 8}
  //AffectDefautCode que si gestion du code salarié en Numérique
  if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Sal) < 11) and (isnumeric(Sal)) then
  begin
    Sal := ColleZeroDevant(StrToInt(Sal), 10);
    SetControlText('PCN_SALARIE', sal);
  end;
  if (SalarieEnter <> Sal) and (OkPassage = False) then
  begin
    Salarie := GetControlText('PCN_SALARIE');
    if (Salarie <> '') then
    begin
      Q := opensql('SELECT PSA_LIBELLE,PSA_PRENOM,PSA_CALENDRIER,PSA_STANDCALEND,PSA_CONGESPAYES,' +
        'ETB_ETABLISSEMENT,ETB_DATECLOTURECPN,ETB_CONGESPAYES ' +
        'FROM SALARIES ' +
        'LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT=ETB_ETABLISSEMENT ' +
        'WHERE PSA_SALARIE= "' + salarie + '" ' + StWhere, TRUE);
      if not Q.eof then
      begin
        Etablissement := Q.findfield('ETB_ETABLISSEMENT').Asstring;
        Nom := Q.findfield('PSA_LIBELLE').Asstring;
        Prenom := Q.findfield('PSA_PRENOM').Asstring;
        CPEtab := (Q.findfield('ETB_CONGESPAYES').Asstring = 'X'); { PT43-2 }
        CPSal := (Q.findfield('PSA_CONGESPAYES').Asstring = 'X'); { PT43-2 }
        DTClot := Q.findfield('ETB_DATECLOTURECPN').AsDateTime;
        calendrier := Q.findfield('PSA_CALENDRIER').AsString;
        standcalend := Q.findfield('PSA_STANDCALEND').AsString;
      end
      else
      begin
        DTclot := idate1900;
        Etablissement := '';
        Nom := '';
        Prenom := '';
        calendrier := '';
        standcalend := '';
        CPEtab := False;
        CPSal := False; { PT43-2 }
      end;
      if GetField('PCN_ETABLISSEMENT') = '' then SetField('PCN_ETABLISSEMENT', Etablissement);
      SetField('PCN_ORDRE', IncrementeSeqNoOrdre(GetField('PCN_TYPEMVT'), Salarie));
    end
    else
    begin
      DTclot := idate1900;
      Etablissement := '';
      Nom := '';
      Prenom := '';
      calendrier := '';
      standcalend := '';
      CPEtab := False;
      CPSal := False; { PT43-2 }
    end;
    Ferme(Q);
  end;
  OkPassage := False;
  {FIN PT- 8}
end;

//DEB PT22 Sur Onclick du bouton on genère l'attestion associé en création

procedure TOM_AbsenceSalarie.BAttestClick(Sender: TObject);
var
  TypeAttest, Fichier, Fiche: string;
begin
  TypeAttest := '';
  if not assigned(T_MotifAbs) then exit; //PT34
  TypeAttest := T_MotifAbs.GetValue('PMA_TYPEATTEST');
  if TypeAttest <> '' then
  begin
    Fichier := RechDom('PGTYPEATTESTATION', TypeAttest, True);
    if Fichier = 'ACCTRAVAIL' then Fiche := 'ACCIDENTTRAVAIL' else Fiche := Fichier;
    AglLanceFiche('PAY', Fiche, '', '',GetField('PCN_SALARIE') +';CREATION;' +
      T_MotifAbs.GetValue('PMA_TYPEABS') +';' +DateToStr(GetField('PCN_DATEDEBUTABS')))
    { PT68 Mise en commentaire 
    if FileExists(VH_Paie.PGCheminRech + '\' + Fichier + '.pdf') then //PT27 ajout date absence
      //     AglLanceFiche ('PAY',Fiche,'','',GetField('PCN_SALARIE')+';CREATION;'+DateToStr(GetField('PCN_DATEDEBUTABS')))
      // PT49 IJSS - Le type d'absence passé en paramètre à l'attestation
      //             Il sera utiliser pour alimenter les cases à cocher
      AglLanceFiche('PAY', Fiche, '', '',GetField('PCN_SALARIE') +';CREATION;' +
      T_MotifAbs.GetValue('PMA_TYPEABS') +';' +DateToStr(GetField('PCN_DATEDEBUTABS')))
    else
      PGIBox('le fichier ' + Fichier + '.pdf n''existe pas sous le répertoire spécifié!', Ecran.caption);}
  end
  else
    PgiBox('Aucun type d''attestation n''a été paramétré pour ce motif d''absence.', Ecran.caption);
end;
//Fin PT22

//DEB PT24 Charge caracteristiques motifabsence du type de l'absence

procedure TOM_AbsenceSalarie.ChargeTob_MotifAbsence;
var
  St: string;
begin
  //DEB PT34
  if (GetField('PCN_TYPEMVT') = 'CPA') and (GetField('PCN_TYPECONGE') <> 'PRI') then
  begin
    t_motifabs := nil;
    Exit;
  end;
  t_motifabs := nil;
  St := Getfield('PCN_TYPECONGE');
  if st <> '' then
    T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [St], False);
  //FIN PT34
end;
//FIN PT24

procedure TOM_AbsenceSalarie.OnClose;
begin
  inherited;
  FreeAndNil(tob_motifabs); //PT34
  {PT31 Free de la tob}
  if Tob_Exercice <> nil then
  begin
    Tob_Exercice.free;
    Tob_Exercice := nil;
  end;
  if Origine = 'E' then
    if T_Recap <> nil then  { PT92 }
    begin
      T_Recap.free;
      T_Recap := nil;
    end;
end;

procedure TOM_AbsenceSalarie.OnAfterUpdateRecord;
Var
  TEvent : TStringList; { PT94 }
begin
  inherited;
  //PT25 Recalcul des compteurs en cours lors de la validation
  if (LastError = 0) then
  begin
    if (Origine = 'E') then
    begin
      SetControlEnabled('PCN_TYPECONGE', False);
      CalculRecapAbsEnCours(GetField('PCN_SALARIE'));
      //TFFiche(Ecran).Retour := GetField('PCN_VALIDRESP');
      TFFiche(Ecran).Retour := IntToStr(GetField('PCN_ORDRE')) + ';' +
        GetField('PCN_TYPECONGE') + ';' +
        DateToStr(GetField('PCN_DATEDEBUTABS')) + ';' +
        DateToStr(GetField('PCN_DATEFINABS')) + ';' +
        GetField('PCN_DEBUTDJ') + ';' +
        GetField('PCN_FINDJ') + ';' +
        FloatToStr(GetField('PCN_JOURS')) + ';' +
        FloatToStr(GetField('PCN_HEURES')) + ';' +
        GetField('PCN_LIBELLE') + ';' +
        GetField('PCN_LIBCOMPL1') + ';' +
        GetField('PCN_LIBCOMPL2') + ';' +
        GetField('PCN_VALIDRESP') + ';' +
        GetField('PCN_EXPORTOK') ;
      
    end
    else { DEB PT56-1 }
      if (Origine = 'A') or (Origine = 'P')  then   // pt95
      Begin
         SetControlEnabled('PCN_SALARIE', False);
         {DEB PT65}
         TFFiche(Ecran).Retour := IntToStr(GetField('PCN_ORDRE')) + ';' +
         GetField('PCN_TYPECONGE') + ';' +
         DateToStr(GetField('PCN_DATEDEBUTABS')) + ';' +
         DateToStr(GetField('PCN_DATEFINABS')) + ';' +
         GetField('PCN_DEBUTDJ') + ';' +
         GetField('PCN_FINDJ') + ';' +
         FloatToStr(GetField('PCN_JOURS')) + ';' +
         FloatToStr(GetField('PCN_HEURES')) + ';' +
         GetField('PCN_LIBELLE') + ';' +
         GetField('PCN_LIBCOMPL1') + ';' +
         GetField('PCN_LIBCOMPL2') + ';' +
         GetField('PCN_VALIDRESP') + ';' +
         GetField('PCN_EXPORTOK')+ ';' +
         GetField('PCN_CODETAPE');

         { FIN PT65 }
  //    If (Origine = 'P') Then CompteursARecalculer(DDAbs, Salarie); //PT102
  // modif NA le 07/02/207 Si module presence coché dans les paramètres société alors met à jour les compteurs à recalculer
     if VH_Paie.PGMODULEPRESENCE then CompteursARecalculer(DDAbs, Salarie); // PT102

         End;

    ACTION := 'MODIFICATION';
    GblOrdre := GetField('PCN_ORDRE');
    GblEnterTypeMvt := GetField('PCN_TYPEMVT');
    GblJours := GetField('PCN_JOURS'); { DEB PT59 }
    GblEnterTypeConge := GetField('PCN_TYPECONGE');
    GblEnterTypeImpute := GetField('PCN_TYPEIMPUTE');
    GblEnterSensAbs := GetField('PCN_SENSABS'); { FIN PT59 }
  end;

  { DEB PT94 }
   if (GblEnterEtatPostPaie <>GetField('PCN_ETATPOSTPAIE')) AND  (GetField('PCN_ETATPOSTPAIE')= 'NAN') then
     Begin
     TEvent := TStringList.create;
     // PT108 Rajout du code du motif absence pour un mvt annulé
     TEvent.Add('L''absence '+ GetField('PCN_TYPECONGE') + ' du '+DateToStr(GetField('PCN_DATEDEBUTABS'))+' au '+ DateToStr(GetField('PCN_DATEFINABS')) +'du salarié '+GetField('PCN_SALARIE')+' a été annulée.');
     CreeJnalEvt('002','131','OK',nil, nil, TEvent);
     if TEvent <> nil then TEvent.free;
     End;
  { FIN PT94 }

  if OnFerme then Ecran.Close; { FIN PT56-1 }
end;

procedure TOM_AbsenceSalarie.OnAfterDeleteRecord;
begin
  inherited;
     If (LastError = 0) And (Origine = 'P') Then CompteursARecalculer(DDAbs, Salarie); //PT102
end;

//PT32 Récupération du récap, alimentation caption de la fiche

procedure TOM_AbsenceSalarie.AffectInfoRecap(Sal: string);  { PT92 }
begin
{ DEB  PT92 }
  T_Recap := ChargeTob_Recapitulatif(Salarie);
 { Q := Opensql('SELECT * FROM RECAPSALARIES WHERE PRS_SALARIE = "' + Salarie + '"', true);
  if not Q.Eof then
  begin
    Tob_Recapitulatif := Tob.Create('Récapitulatif Salarié', nil, -1);
    Tob_Recapitulatif.LoadDetailDB('Récapitulatif Salarié', '', '', Q, False);
    T := Tob_Recapitulatif.FindFirst(['PRS_SALARIE'], [Salarie], False);
  end
  else
  begin
    Tob_Recapitulatif := nil;
    T := nil;
  end;
  Ferme(Q);  }
{ FIN PT92 }
  // PT1-3 Ph Réecriture
  if T_Recap <> nil then
  begin
    { //PT- 7-13 //PT26-3 mise en commentaire puis supprimer}
    {PT64 SetControlText('LE_ACQUISN1', T.GetValue('PRS_ACQUISN1'));
    SetControlText('LE_PRISN1', T.GetValue('PRS_PRISN1'));
    SetControlText('LE_ACQUISN', T.GetValue('PRS_ACQUISN'));
    SetControlText('LE_PRISN', T.GetValue('PRS_PRISN'));
    SetControlText('LE_CUMRTTACQUIS', T.GetValue('PRS_CUMRTTACQUIS'));
    SetControlText('LE_CUMRTTPRIS', T.GetValue('PRS_CUMRTTPRIS'));}
    SetControlText('LE_RESTN1', T_Recap.GetValue('PRS_RESTN1'));     { PT92 }
    SetControlText('LE_RESTN', T_Recap.GetValue('PRS_RESTN'));       { PT92 }
    SetControlText('LE_RESTRTT', T_Recap.GetValue('PRS_CUMRTTREST'));    { PT92 }
  end
  else
  begin
    {PT64 SetControlText('LE_ACQUISN1', '0');
    SetControlText('LE_PRISN1', '0');
    SetControlText('LE_ACQUISN', '0');
    SetControlText('LE_PRISN', '0');
    SetControlText('LE_CUMRTTACQUIS', '0');
    SetControlText('LE_CUMRTTPRIS', '0');
    SetControlText('LE_CUMRTTPRIS', '0'); //PT- 7-13}
    SetControlText('LE_RESTN1', '0');
    SetControlText('LE_RESTN', '0');
    SetControlText('LE_RESTRTT', '0'); //PT26-3
  end;
end;

{ DEB PT40-1 }

procedure TOM_AbsenceSalarie.ImprimerClick(Sender: TObject);
var
  StSql: string;
begin
  if origine = 'E' then
  begin
    StSql := 'SELECT PCN_SALARIE,PCN_TYPEMVT,PCN_TYPECONGE,PCN_LIBELLE,' +
      'PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_DEBUTDJ,PCN_FINDJ,' +
      'PCN_JOURS,PCN_HEURES,PCN_LIBCOMPL1,PCN_LIBCOMPL2,PCN_VALIDRESP,' +
      'PCN_EXPORTOK,PCN_VALIDABSENCE,PCN_VALIDSALARIE,'+   { PT73 }
      'PCN_DATEMODIF,PCN_DATECREATION FROM ABSENCESALARIE ' +
      'WHERE PCN_SALARIE="' + GetField('PCN_SALARIE') + '" ' +
      'AND PCN_TYPEMVT="' + GEtField('PCN_TYPEMVT') + '" ' +
      'AND PCN_ORDRE=' + IntToStr(GetField('PCN_ORDRE')) + ' ';
    LanceEtat('E', 'PCG', 'PAB', True, False, False, nil, StSql, '', False);
  end
  else
  begin
    StSql := 'AND PCN_SALARIE="' + GetField('PCN_SALARIE') + '" ' +
      'AND PCN_TYPEMVT="' + GetField('PCN_TYPEMVT') + '" ' +
      'AND PCN_ORDRE=' + IntToStr(GetField('PCN_ORDRE')) + '';

    LanceEtat('E', 'PAY', 'PCN', True, False, False, nil, StSql, '', False);
  end;
end;
{ FIN PT40-1 }
{ DEB PT40-4 Fonction générique pour envoie de mail }

{ FIN PT40-4 }
// d PT49 IJSS
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 09/06/2004
Modifié le ... : 09/06/2004
Description .. : procédure MajIJSS
Suite ........ : Calcul des champs liés aux IJSS :   Nbre jours calendaires
Suite ........ : d'absence,  nbre jours carence, nbre jours d'IJSS.
Suite ........ : Le nbre de jours de carence vaut de 0 à 3. On tient compte
Suite ........ : de la carence précédente si les absences sont
Suite ........ : consécutives
Mots clefs ... : IJSS
*****************************************************************}

procedure TOM_AbsenceSalarie.MajIJSS;
var
  st: string;
  Q: TQuery;
  WDateFin, WDatedeb: TDateTime;
  WCarence: double;
  AbsIJSS: TOB;
  CarenceIJSS: integer; // PT57
begin
  if Origine = 'P' then exit; // PT95

  // d PT57
  if assigned(T_MotifAbs) then
  begin
// d PT78
{    if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'ATJ') then
      // Accident de trajet
      CarenceIJSS := 1
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'ATR') then
      // Accident du travail
      CarenceIJSS := 1
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'MAN') then
      // Maladie non professionnelle
      CarenceIJSS := 3
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'MAP') then
      // Maladie professionnelle
      CarenceIJSS := 1
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'MAT') then
      // Maternité
      CarenceIJSS := 0
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'PAT') then
      // Paternité
      CarenceIJSS := 0
    else
      if (T_MotifAbs.GetValue('PMA_TYPEABS') = 'THE') then
      // Mi-temps thérapeutique
      CarenceIJSS := 0;}
      CarenceIJSS := T_MotifAbs.GetValue('PMA_CARENCEIJSS') ;
// f PT78
  end;
  // f PT57
  if GetField('PCN_GESTIONIJSS')<>'X' then
    SetField('PCN_GESTIONIJSS', 'X');
  // création de la TOB des absences pour les 12 mois précédents
  Tob_AbsIJSS := Tob.Create('Absences IJSS', nil, -1);
  st := 'SELECT PCN_SALARIE, PCN_DATEDEBUTABS, PCN_DATEFINABS, ' +
    'PCN_LIBELLE, PCN_NBJCARENCE,PCN_GESTIONIJSS ' +
    'FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + Salarie +
    '" AND PCN_GESTIONIJSS="X" AND ' +
    'PCN_DATEDEBUTABS >="' + UsDateTime(PlusMois(DDabs, -12)) + '" AND ' +
    'PCN_DATEDEBUTABS <="' + UsDateTime(DDabs - 1) + '" ' +
    'ORDER BY PCN_SALARIE, PCN_DATEFINABS';
  Q := OpenSql(st, TRUE);
  if not (Q.eof) then
    Tob_AbsIJSS.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
  ferme(Q);

  // détermination de la carence cumulée depuis 12 mois
  WDateFin := IDate1900;
  WCarence := 0;

  AbsIJSS := Tob_AbsIJSS.FindFirst([''], [''], false);
  while (AbsIJSS <> nil) do
  begin
    if (AbsIJSS.GetValue('PCN_DATEDEBUTABS') <> WDateFin + 1) then
      WCarence := AbsIJSS.GetValue('PCN_NBJCARENCE')
    else
      WCarence := WCarence + AbsIJSS.GetValue('PCN_NBJCARENCE');
    //d PT57
    if (WCarence > CarenceIJSS) then WCarence := CarenceIJSS;
    //    if (WCarence > 3) then WCarence := 3;
    // f PT57
    WDateFin := AbsIJSS.GetValue('PCN_DATEFINABS');
    AbsIJSS := Tob_AbsIJSS.FindNext([''], [''], false);
  end;

  // calcul du nombre de jours calendaires d'absence
  if GetField('PCN_NBJCALEND') <> (Getfield('PCN_DATEFINABS') - getfield('PCN_DATEDEBUTABS') + 1) then
// d PT91
  begin
    SetField('PCN_NBJCALEND',
      Getfield('PCN_DATEFINABS') - getfield('PCN_DATEDEBUTABS') + 1);
    SetField('PCN_NBJCARENCE', CarenceIJSS);
    SetField('PCN_IJSSSOLDEE', '-');
  end;
// f // PT91
  // calcul du nbre de jours de carence
  WDateDeb := GetField('PCN_DATEDEBUTABS');
  if (WDateDeb = WDateFin + 1) then
// d // PT91
  begin
    // d PT57
    if (GetField('PCN_NBJCARENCE') <> (CarenceIJSS - WCarence)) then
       SetField('PCN_NBJCARENCE', CarenceIJSS - WCarence) ;
      //    SetField('PCN_NBJCARENCE', 3 - WCarence)
// f PT57
  end
// f // PT91
  else
    // d PT57
    //    SetField('PCN_NBJCARENCE', 3);
    if (GetField('PCN_NBJCARENCE') <> CarenceIJSS) and (Action = 'CREATION') then // PT91
       SetField('PCN_NBJCARENCE', CarenceIJSS);
  // f PT57

  if (GetField('PCN_NBJCALEND') < GetField('PCN_NBJCARENCE')) then   // PT91
    SetField('PCN_NBJCARENCE', GetField('PCN_NBJCALEND'));

  // calcul du nombre de jours d'IJSS
  if (GetField('PCN_NBJIJSS') <> (Getfield('PCN_NBJCALEND') - getfield('PCN_NBJCARENCE'))) then  // PT91
    SetField('PCN_NBJIJSS',
      Getfield('PCN_NBJCALEND') - getfield('PCN_NBJCARENCE'));

  if (Valeur(GetField('PCN_NBJIJSS')) = 0) then
    IJSSaSolder := True
  else
    IJSSaSolder := False;

  FreeAndNil(Tob_AbsIJSS);

end; {fin MajIJSS}
// f PT49 IJSS

procedure TOM_AbsenceSalarie.TRecapClick(Sender: TObject);
begin
AglLanceFiche ('PAY','RECAPSALCP', '', GetField ('PCN_SALARIE'), ''); { PT64 }
end;

{ DEB PT76-1 }
procedure TOM_AbsenceSalarie.GestionOngletAnnulation;
begin
if Origine = 'ABSANNUL' then
  Begin
  SetControlEnabled('PCN_DATEANNULATION'  ,(GetField('PCN_ETATPOSTPAIE') <> 'VAL'));
  SetControlEnabled('PCN_MOTIFANNUL'      ,(GetField('PCN_ETATPOSTPAIE') <> 'VAL'));
  SetControlEnabled('PCN_ANNULEPAR'       ,False);
  SetControlEnabled('BINSERT'             ,False);
  SetControlEnabled('BDELETE'             ,False);
  SetControlEnabled('BATTEST'             ,False);
  SetControlEnabled('BCALENDRIER'         ,False);
  if GetField('PCN_ETATPOSTPAIE')='VAL' then
    Begin
    If GetField('PCN_DATEANNULATION') <> Idate1900 then SetField('PCN_DATEANNULATION',Idate1900);
    If GetField('PCN_MOTIFANNUL')     <> ''        then SetField('PCN_MOTIFANNUL','');
    If GetField('PCN_ANNULEPAR')      <> ''        then SetField('PCN_ANNULEPAR','');
    End
  else
    if GetField('PCN_ETATPOSTPAIE')='NAN' then
       Begin
       If GetField('PCN_DATEANNULATION') = Idate1900  then SetField('PCN_DATEANNULATION',Date);
       If GetField('PCN_MOTIFANNUL')     = ''         then SetField('PCN_MOTIFANNUL',Copy('Annulation '+GetField('PCN_LIBELLE'),1,35));
       If GetField('PCN_ANNULEPAR')      = ''         then SetField('PCN_ANNULEPAR',V_PGI.User);
    End

  End
else
   SetControlProperty('TBPOSTPAIE','Tabvisible',False);
end;
{ FIN PT76-1 }
{ DEB PT90 }
procedure TOM_AbsenceSalarie.InitEntreeSortie;
begin
if getcontrol('ENTREE') <> nil then     //PT96
  SetControlText('ENTREE',DateToStr(DateEntree));    { DEB PT72 }
if DateSortie > Idate1900 then
    Begin
    SetControlVisible('SORTIE',True);
    SetControlVisible('LSORTIE',True);
    SetControlText('SORTIE',DateToStr(DateSortie));
    End                                             { FIN PT72 }
else
  Begin
  SetControlVisible('SORTIE',False);
  SetControlVisible('LSORTIE',False);
  End;
end;
{ FIN PT90 }

//Deb PT95
procedure TOM_AbsenceSalarie.OnExitHeure(Sender: TObject);
Var
  St : String;
  H1,H2,DH,HN : Double;
  J1,J2 : TDateTime;
begin
  St := UpperCase(TControl(Sender).Name);
  if  (Origine = 'P') and (St = 'HDEB')
  and (getfield('PCN_DATEFINABS') = getfield('PCN_DATEDEBUTABS'))
  and (GetControlText('HFIN') < GetControlText('HDEB')) then
  begin
    SetControlText('HFIN', getControlText('HDEB'));
    SetField('PCN_HFIN', getControlText('HDEB'));
  end;

  If (St = 'HDEB') or (St = 'HFIN') or (St = 'PCN_DATEDEBUTABS') or (St = 'PCN_DATEFINABS') Or (St = 'PCN_NBHEURESNUIT') then //PT99
  Begin
    H1 := StrTimeToFloat(GetControlText('HDEB'),True);
    H2 := StrTimeToFloat(GetControlText('HFIN'),True);
    J1 := GetField('PCN_DATEDEBUTABS');
    J2 := GetField('PCN_DATEFINABS');
    HN := GetField('PCN_NBHEURESNUIT'); //PT99
  // DH := (J2-J1)*24 + (H2 - H1) - HN;  //PT99
    DH := (J2-J1)*24 + (H2 - H1);
      //DH := CalculEcartHeure(D1,D2);
    if DH <> GetField('PCN_HEURES') then
    Begin
      If DS.State in [DsBrowse] then DS.Edit;
      SetField('PCN_HEURES',DH);
    end;
    If (St <> 'PCN_NBHEURESNUIT') And (GetControlText(st) <> GetField('PCN_'+St)) then //PT99
    Begin
      If DS.State in [DsBrowse] then DS.Edit;
      SetField('PCN_'+St,GetControlText(st));
    End;
  End;
end;
//Fin PT95

procedure TOM_AbsenceSalarie.OnExitDate(Sender: TObject);
begin
  if  (Origine = 'P') and (UpperCase(TControl(Sender).Name) = 'PCN_DATEDEBUTABS')
  {and (getfield('PCN_DATEFINABS') < getfield('PCN_DATEDEBUTABS'))} then  //PT101
    SetField('PCN_DATEFINABS', getfield('PCN_DATEDEBUTABS'));
  if assigned(T_MotifAbs) then
  if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU' then
    OnExitHeure(Sender);
  SetField('PCN_LIBELLE',RendLibPresence(GetField('PCN_TYPECONGE'),getfield('PCN_DATEDEBUTABS'),getfield('PCN_DATEFINABS'))); { PT92 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 02/08/2007 / PT100
Modifié le ... :   /  /
Description .. : Contrôle de remplissage des zones obligatoires pour la saisie
Suite ........ : des mouvements de présence
Mots clefs ... : 
*****************************************************************}
procedure TOM_AbsenceSalarie.ControleSaisiePresence;
begin
  if getcontrolText('PCN_TYPECONGE') = '' then
  begin
    Error := 9;
    setfocuscontrol('PCN_TYPECONGE');
    exit;
  end;
  if (GetField('PCN_DATEDEBUTABS') = (Idate1900)) then
  begin
    Error := 11;
    setfocuscontrol('PCN_DATEDEBUTABS');
    exit;
  end;
  if (GetField('PCN_DATEFINABS') = (Idate1900)) then
  begin
    Error := 12;
    setfocuscontrol('PCN_DATEFINABS');
    exit;
  end;
  if (GetField('PCN_DATEFINABS') < GetField('PCN_DATEDEBUTABS')) then
  begin
    Error := 1;
    setfocuscontrol('PCN_DATEFINABS');
    exit;
  end;
  if    (GetField('PCN_DATEFINABS') = GetField('PCN_DATEDEBUTABS'))
    and (getcontrolText('HFIN') < getcontrolText('HDEB')) then
  begin
    Error := 31;
    setfocuscontrol('HFIN');
    exit;
  end;
  if    (GetField('PCN_HEURES') < 0) Then
  begin
    Error := 31;
    setfocuscontrol('HDEB');
    exit;
  end;
End;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 06/08/2007 / PT101
Modifié le ... :   /  /
Description .. : Duplication d'un mouvement de présence
Mots clefs ... :
*****************************************************************}
procedure TOM_AbsenceSalarie.Dupliquer (Sender : Tobject);
var T,TobSemaine : TOB;
    DD : TDateTime;
    HD : TDateTime;
    Ordre, i, NbJoursOuvres  : Integer;
    ReposHebdo1, ReposHebdo2 : String;
    Q : TQuery;
begin
     // Duplication pour la journée : On ne modifie que la date
     If (TMenuItem(Sender).Name = 'J') Then
     Begin
          // Récupération des données actuelles
          T := TOB.Create('ABSENCESALARIE', Nil, -1);
          T.GetEcran(Ecran);

          // Création du nouvel enregistrement
         (GetControl('BINSERT') as TToolbarButton97).OnClick(Sender);

          // Récupération des anciennes données : à cause des variables de classe de partout, on ne peut pas
          // utiliser le putEcran. Il faut donc tout faire à la main
          SetField('PCN_SALARIE', T.GetValue('PCN_SALARIE'));
          SetField('PCN_TYPECONGE', T.GetValue('PCN_TYPECONGE'));

          // Par défaut, on met le lendemain
          DD := T.GetValue('PCN_DATEDEBUTABS');
          DD := DD + 1;
          SetField('PCN_DATEDEBUTABS', DD);
          SetField('PCN_DATEFINABS',   DD);

          // Récupération des heures
          HD := T.GetDateTime('PCN_HDEB');
          SetField('PCN_HDEB', HD); SetControlText('HDEB', Copy(TimeToStr(HD),1,5));
          HD := T.GetDateTime('PCN_HFIN');
          SetField('PCN_HFIN', HD); SetControlText('HFIN', Copy(TimeToStr(HD),1,5));
          SetField('PCN_HEURES', T.GetValue('PCN_HEURES'));
          SetField('PCN_NBHEURESNUIT', T.GetValue('PCN_NBHEURESNUIT'));

          SetFocusControl('PCN_DATEDEBUTABS');

          // Rafraîchit le libellé
          AffectLibelle;

          T.Free;
     End
     // Duplication à la semaine : Il faut créer plusieurs enregistrements
     Else If (TMenuItem(Sender).Name = 'S') Then
     Begin
          ReposHebdo1 := '0'; ReposHebdo2 := '0';

          // Récupération du paramétrage de l'établissement auquel appartient le salarié
          Q := OpenSQL('SELECT ETB_NBJOUTRAV, ETB_1ERREPOSH, ETB_2EMEREPOSH FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="'+Etablissement+'"', True);
          If Not Q.EOF Then
          Begin
               NbJoursOuvres := Q.FindField('ETB_NBJOUTRAV').AsInteger;
               If Q.FindField('ETB_1ERREPOSH').AsString  <> '' Then ReposHebdo1 := Copy(Q.FindField('ETB_1ERREPOSH').AsString,1,1);
               If Q.FindField('ETB_2EMEREPOSH').AsString <> '' Then ReposHebdo2 := Copy(Q.FindField('ETB_2EMEREPOSH').AsString,1,1);

               // Dimanche chômé par défaut si aucun paramétrage
               If (ReposHebdo1 = '0') And (ReposHebdo2 = '0') And (NbJoursOuvres = 0) Then ReposHebdo1 := '1';
          End;
          Ferme(Q);

          // Création de la TOB qui contiendra les données générées
          TobSemaine := TOB.Create('LaSemaine', Nil, -1);

          DD := GetField('PCN_DATEDEBUTABS');

          // Si on est le dernier jour de la semaine, pas de génération
          If (DayOfWeek(DD) <> 7) Then
          Begin
               // Récupération du nouveau numéro d'ordre
               Ordre := IncrementeSeqNoOrdre('PRE', Salarie);
               DD := DD + 1;

               For i := DayOfWeek(DD) To 7 Do
               Begin
                    If (i <> StrToInt(ReposHebdo1)) And (i <> StrToInt(ReposHebdo2)) Then
                    Begin
                         // Création des données dans la TOB principale
                         T := TOB.Create ('ABSENCESALARIE', TobSemaine, -1);

                         T.PutValue('PCN_TYPEMVT',          'PRE');
                         T.PutValue('PCN_SALARIE',          Salarie);
                         T.PutValue('PCN_GUID',             AglGetGUID());
                         T.PutValue('PCN_ORDRE',            Ordre);
                         T.PutValue('PCN_TYPECONGE',        GetField('PCN_TYPECONGE'));
                         T.PutValue('PCN_DATEDEBUTABS',     DD);
                         T.PutValue('PCN_DATEFINABS',       DD);
                         T.PutValue('PCN_DATEVALIDITE',     DD);
                         T.PutValue('PCN_HDEB',             GetField('PCN_HDEB'));
                         T.PutValue('PCN_HFIN',             GetField('PCN_HFIN'));
                         T.PutValue('PCN_HEURES',           GetField('PCN_HEURES'));
                         T.PutValue('PCN_NBHEURESNUIT',     GetField('PCN_NBHEURESNUIT'));
                         T.PutValue('PCN_JOURS',            GetField('PCN_JOURS'));
                         T.PutValue('PCN_LIBELLE',          RendLibPresence(GetField('PCN_TYPECONGE'),DD,DD));

                         T.PutValue('PCN_ETABLISSEMENT',    Etablissement);

                         Q := OpenSQL('SELECT PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL' +
                                      ' FROM SALARIES WHERE PSA_SALARIE = "' + Salarie + '"', True);
                         If Not Q.EOF Then
                         Begin
                              T.PutValue('PCN_TRAVAILN1',   Q.FindField('PSA_TRAVAILN1').AsString);
                              T.PutValue('PCN_TRAVAILN2',   Q.FindField('PSA_TRAVAILN2').AsString);
                              T.PutValue('PCN_TRAVAILN3',   Q.FindField('PSA_TRAVAILN3').AsString);
                              T.PutValue('PCN_TRAVAILN4',   Q.FindField('PSA_TRAVAILN4').AsString);
                              T.PutValue('PCN_CODESTAT',    Q.FindField('PSA_CODESTAT').AsString);
                              T.PutValue('PCN_CONFIDENTIEL',Q.FindField('PSA_CONFIDENTIEL').AsString);
                         End;
                         Ferme(Q);

                         // Valeurs par défaut
                         T.PutValue('PCN_MVTORIGINE',       'SAL');
                         T.PutValue('PCN_ETATPOSTPAIE',     'VAL');
                         T.PutValue('PCN_CODETAPE',         '...');
                         T.PutValue('PCN_VALIDSALARIE',     'SAL');
                         T.PutValue('PCN_VALIDRESP',        'VAL');
                         T.PutValue('PCN_EXPORTOK',         'X');

                         Ordre := Ordre + 1;
                    End;
                    DD    := DD + 1;
               End;
          End;

          If TobSemaine.Detail.Count > 0 Then
          Begin
               TobSemaine.InsertOrUpdateDB();
               PGIInfo (Format(TraduireMemoire('%s évènements ont été créés.'),[IntToStr(TobSemaine.Detail.Count)]), TraduireMemoire('Opération effectuée'));
          End
          Else
               PGIInfo (TraduireMemoire('Aucun évènement n''a été créé.'));

          FreeAndNil(TobSemaine);

          TToolbarButton97(GetControl('BFerme')).Click;
     End;
end;

initialization
  registerclasses([TOM_AbsenceSalarie]);

end.


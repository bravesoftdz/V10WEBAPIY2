{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/06/2001                                                                              
Modifié le ... : 12/06/2001
Description .. : gestion de l'établissement complémentaire
Mots clefs ... : PAIE,PGETAB
*****************************************************************}
{
PT1   : 12/06/2001 V547 SB Suppression des congés payés lors de la désactivation
                           des CP.
                           Bug : Suppression brutale de ts les mvts ABSENCES.
                           On ne doit supprimer que les mvts de type CPA
PT2   : 15/06/2001 V547 SB Cas : Recalcul des periodes sur mvt de SLD eclatés
                           sur pls periodes. Le recalcul tient compte de la date
                           de paiement. Pour ces mvts la date concerne la
                           periode en cours au lieu de la periode anterieur..
PT3   : 31/07/2001 V547 SB Modification champ suffixe MODEREGLE
PT4   : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
PT5   : 31/08/2001 V547 SB Le calendrier ne doit pas être réinitialiser
                           quand on décoche la gestion des CP..
                           (Fiche bug n°277)
PT6   : 10/09/2001 V547 PH Initialisation CP coché par defaut
PT7   : 10/09/2001 V547 SB Champ ETB_PGMODEREGLE DataType non renseigné
PT8   : 01/10/2001 V562 PH Nouveau champ ETB_PROFILREM Profil rémunération
PT9-1 : 19/10/2001 V562 SB Fiche de bug n°340
                           On effectue des update SQL pour mettre à jour des ENR
                           donc plantage. Suppression de ces UPDATE
PT9-2 : 19/10/2001 V562 SB Test si numeric au cas ou je modifie le onargument
PT10  : 26/10/2001 V562 JL Modification gestion medecine et DDTEFP
PT11  : 12/11/2001 V563 VG Contrôle de la fraction DADS-U
PT12  : 21/11/2001 V563 PH Initialisation des bnouveaux champs
PT13  : 22/11/2001 V563 SB Pour ORACLE, Le champ PCN_CODETAPE doit être
                           renseigné
PT14  : 29/11/2001 V563 VG Mise à jour de tous les salariés de l'établissement
                           qui ont idem étab au niveau de la fraction
PT15  : 15/01/2002 V571 SB Erreur Convertion type variant sur décoche CP
PT16  : 29/03/2002 V571 JL Ajout champs MSA et Intermittents du spectacle
PT17  : 02/04/2002 V571 Ph Ajout champ profil retraite
PT18  : 23/04/2002 V582 Ph tests et autorisation suppression etablissement
                           complémentaire
PT19  : 06/05/2002 V582 VG Affichage de l'onglet MSA
PT20  : 10/05/2002 V582 JL Vérification horaire de référence
                           Fiche de bug n° 10110
PT21  : 16/05/2002 V582 VG Suppression de l'affichage de certains paramètres
                           pour la version S3
PT22  : 17/05/2002 V582 JL Champs n° intermittent invisible si non cochés au
                           lieu d'enabled
PT23  : 13/06/2002 V582 JL Propriété champ activité MSA modifiées -> SUPPRIME LE
                           17/12/2002 V591 car présent ds la fiche
PT24-1: 18/09/2002 V585 SB FQ n°10022 Ajout champ Organisme à editer sur
                           bulletin
PT24-2: 23/09/2002 V585 SB FQ n°10227 Affichage des libellés des ribs
PT25  : 16/09/2002 V585 SB Export de la partie decloture dans une procedure à
                           part
PT26  : 21/10/2002 V585 SB Génération d'évènements pour traçage
PT27  : 17/12/2002 V591 PH Portage CWAS = THGrid sur grille des taux AT
PT28  : 16/01/2003 V591 PH FQ 10432 gestion des taux at avec rafraichissement de
                           la liste des taux
PT29  : 12/03/2003 VG V_42 Ajout de la gestion des acomptes en version S3
PT30  : 26/03/2003 VG V_42 Contrôle de l'horaire établissement et du prorata TVA
                           FQ N° 10614
PT31  : 07/04/2003 SB V_42 FQ 10582 ajout controle vraisemblance Test au minimum
PT32  : 24/04/2003 SB V_42 Intégration de la gestion des activités
PT33  : 07/05/2003 MF V_42 Ajout champ Type tickets restaurant
PT34  : 15/05/2003 PH V_42 Correction mauvaise initialisation champ CP avec NULL
PT35  : 27/06/2003 MF V_42 Suppression Etablissement : vérif Cde titres restau.
PT36  : 31/07/2003 PH V_42 FQ 10201 controle contenu champ test au minimum
PT37  : 17/09/2003 PH V_42 FQ 10110 creation etabcompl par une TOB au lieu de
                           query et initilaisation des champs
PT38-1: 23/09/2003 SB V_42 Portage Cwas : Suppression de la double zone
                           Ancienneté
PT39  : 30/09/2003 VG V_42 FQ 10848 Création de l'établissement : Initialisation
                           de ETB_CONGESPAYES à '-' et ETB_CODESECTION à '1'
PT40  : 01/10/2003 SB V_42 CP : Suppression des événements Onenter sur champ CP
                           Utilsation d'une simple variable globale initialiser
                           sur OnloadRecord et OnChangeField
PT41  : 02/10/2003 PH V_42 Deplacement de la fonction car elle change
                           f.FieldName en CWAS et les controles ne fonctionnent
                           plus
PT42  : 12/10/2003 PH V_42 Rechargement des caches en CWAS
PT43  : 13/10/2003 PH V_42 FQ 10877 test de la valeur saisie dans zone test au
                           minimum
PT44  : 14/10/2003 VG V_42 Edition de l'établissement
PT45  : 15/10/2003 SB V_42 Maj de date modification de la table des salaries
                           pour introduction PAIE
PT46  : 09/12/2003 SB V_50 FQ 10370 Edition compteurs Cp selon paramétrage
PT47  : 12/02/2004 VG V_50 Modification de l'édition - FQ N°10988
PT48  : 12/03/2004 SB V_50 FQ 11162 Encodage de de la date de cloture erroné si
                           fin fevrier
PT49  : 22/03/2004 SB V_50 Retrait de l'option de décloture des CP
PT50  : 02/04/2004 JL V_50 MSA EDI : Mise à jour changement d'activité
PT51  : 05/04/2004 PH V_50 FQ 11231 Prise en compte Alsace LOrraine
PT52-1: 09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                           salarié
PT52-2: 09/04/2004 SB V_50 FQ 11237 Ajout Gestion idem etab jours Cp Suppl
PT53  : 20/04/2004 PH V_50 Portage DB2
PT54  : 23/04/2004 SB V_50 FQ 11265 Valeur en création à -1 pour code medecine
                           du travail & DDTEFP
PT55-1: 18/05/2004 SB V_50 FQ 11136 Refonte contrôle sur coche CP
PT55-2: 18/05/2004 SB V_50 FQ 11136 Refonte message d'erreur CP
PT56  : 27/05/2004 SB V_50 FQ 11308 Double message d'alerte en cloture CP si
                           paies non éffectuées
PT57  : 28/06/2004 PH V_50 Appel à la fonction de création de l'établissement
                           social
PT58  : 19/08/2004 PH V_50 FQ 11432 Controle contenu de la zone jour acquis
                           bulletin obligatoire
PT59  : 23/08/2004 PH V_50 FQ 11511 Message info en mise en place des CPs
                           informant de la date de clôture
PT60  : 25/08/2004 PH V_50 FQ 11052 Contrôle des champs dates d'édition
PT61  : 29/09/2004 JL V_50 FQ 11637 Zone jour paiement griser si
                           mois paiement = fin de mois
PT62  : 05/10/2004 VG V_50 Utilisation de la TOM pour le TD Bilatéral -
                           FQ N°11654
PT63  : 12/10/2004 VG V_50 Masquage des tickets restaurants et des IJSS -
                           FQ N°11679
PT64  : 15/10/2004 VG V_50 MSA en S3
PT65  : 01/02/2005 VG V_60 Edition fiche établissement en CWAS
PT66  : 02/08/2005 PH V_60 FQ 12486 Ergonomie
PT67  : 31/08/2005 PH V_60 FQ 12491 Clôture CP dans le menu Paie Uniquement
PT68  : 26/09/2005 SB V_65 FQ 12452 Refonte appel méthode pour compatibilité
                           CWAS
PT69  : 20/10/2005 SB V_65 FQ 11693 Lancement du mul banque en CWAS
---- JL 20/03/2006 modification clé annuaire ----
PT70  : 03/04/2006 SB V_65 FQ 11426 Mis en place paramètre pour calcul au
                           maintien CP
PT71  : 03/04/2006 JL V_65 FQ 12462 mettre à jour CCN salarié en fonction de
                           celle de l'établissement
PT72  : 04/04/2006 JL V_65 Gestion UG MSA
PT73  : 06/04/2006 JL V_65 FQ 13048 accès multicritères annuaire pour ddtefp et
                           médecine du travail + 17/10/2006 Ajout
                           rafraichissement combo
PT74  : 11/04/2006 JL V_65 FQ 13063 Modif affichage si pas de gestion MSA
PT75  : 10/05/2006 VG V_65 Contrôles cohérence entre le code NAF et le code
                           section prud'hommal - FQ N°12823
PT76  : 02/06/2006 SB V_65 Optimisation mémoire
PT77-1: 16/06/2006 SB V_65 FQ 13270 Acces zone date de cloture
PT77-2: 16/06/2006 SB V_65 Modification de la requête de suppression des mvts CP
PT78  : 26/06/2006 VG V_70 Correction PT75 : Les exceptions étaient toujours en
                           anomalie - FQ N°12823
PT79  : 19/09/2006 MF V_70 FQ 13500 - contrôle du n° de compte banque en
                           validation de la fiche
PT80  : 22/11/2006 MF V_701 FQ 13500 - modification du contrôle du n° de compte
PT81  : 22/01/2007 MF V_72 Titres restaurant (saisie du code point de livraison
                           ACCOR)
PT82-1: 25/01/2007 SB V_70 FQ 11992 Ajout valeur methode date d'acquisition CP
PT82-2: 25/01/2007 SB V_70 Contrôle cohérence saisie date si modif. paramètres
                           régionaux PC en Anglais
PT83  : 25/01/2007 SB V_70 Reprise PT41
PT84  : 12/02/2007 MF V_70 plus de contrôle du code point de livraison su
                           OnChangeField
PT85-1: 14/03/2007 VG V_72 BQ_GENERAL n'est pas forcément unique
PT85-2: 14/03/2007 VG V_72 Erreur lors de l'accès à la fiche établissement
                           social
PT86  : 07/08/2007 PH V_72 FQ 14135 14565 acces aux multi_critères des comptes
                           banquaires
PT87  : 24/10/2007 VG V_80 Modification des champs prud'hommaux salarié si
                           modification des champs prud'hommaux établissement
PT88  : 29/05/2008 NA V_850 Ticket restaurant : libellé point de livraison = code client si fournisseur = Chéque déjeuner
                            et si coche Facturation / établissement (un code client <> par établ)
PT91  : 04/03/2008 FC V_90 FQ 15270 historiser la convention collective comme les autres champs historisés par avance
PT94  : 30/07/2008 JPP V_80 FQ 15538 Initalisation du masque du champ 'ETB_DATEACQCPANC' lorsque l'on clique sur 'Autre Date'
                            Dans le Combo 'ETB_TYPDATANC'
PT95  : 19/09/2008 JS       FQ 15541  Conservation du champs Valorisation au Maintien
                            Dans le Combo 'ETB_MVALOMS'
}

unit UTOMEtabCompl;

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
{$IFNDEF EAGLCLIENT}
  Dialogs,
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,
  DBCtrls,
  Fiche,
  Fe_main,
  //  BanqueCp,
  EdtREtat,
{$ELSE}
  eFiche,
  MaineAgl,
  UtileAGL,
  {$IFNDEF DADSB}
  UTOFMULPARAMGEN,
  {$ENDIF}
{$ENDIF}
  {$IFNDEF DADSB}
  TRMULCOMPTEBQ_TOF,
  BANQUECP_TOM,
  {$ENDIF}
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  UTOB,
  HTB97,
  Ent1,
  ParamSoc,
  Windows,
  AglInit,
  PgOutilsTreso, Commun, Constantes,
  ULibEditionPaie;

type

  TOM_EtabCompl = class(TOM)
  public
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
{$IFNDEF DADSB}
    procedure OnUpdateRecord; override;
{$ENDIF}
    procedure OnDeleteRecord; override;
    procedure OnClose; override;
    {       procedure OnNewRecord  ;                      override ; PT39 }
  private
    etablissement, libelle, {NbJouTrav,typdatanc,} ValtypDatAnc: string;
    GblBasAnccp, GblNbJoutrav, GblTypDatAnc, GblEditBulCP: string; //PT40
    Activite, Profil, ProfilRbs, ProfilAfp, ProfilApp, ProfilRet, ProfilMut, ProfilPre, ProfilTss: string;
    ProfilCge, ProfilAncien, ProfilTrans, PeriodBul, ProfilRem, StandCalend: string;
    NbreAcquisCp, ValorIndempCp, ValancCp, DateAcqCpAnc, Reliquat, basancp, datanc, PaieValoMs : string; { PT70 }
    RedRepas, RedRtt1, RedRtt2, ModeReglement, RibSalaire, PourcentAfp: string;
    PaiAcompte, RibAcpSoc, PaiFrais, RibFraisSoc, MoisPaiement, JourPaiement, JourHeure: string;
    TypeFraction, EditOrg, NbAcquisCp, TicketRestau: string; //PT33
    Cloture, LoadFiche, DecocheCP, ForceValidCp : boolean; { PT55-1 }
    changedatecloture, ExistCloture: boolean;
    ANow, MNow, DNow: Word;
    DateEnter: tdatetime;
    NbCpSupp: Double; { PT52-2 }
    NumOnglet: string; {VG 16/03/01 Gestion de l'affichage des onglets}
    Imprim: TToolBarButton97;
    GblCongesPayes : String; { PT68 }
    ConventionColl,UGMsa : String; //PT71
    PointLivraison : string; // PT81
    fournisseur:         string; // PT81
    College, LieuVote, Section : string;
    procedure ComboKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{$IFNDEF DADSB}
    procedure ControleZonesOngletCongesPayes;
{$ENDIF}
    procedure EnableChampsOngletCongesPayes;
    procedure InitialiseZonesCongesPayes;
    procedure MDateFExit(Sender: TObject);
    procedure MDateFEnter(Sender: TObject);
{$IFNDEF DADSB}
    procedure ClotcpClick(Sender: TObject);
    procedure DeClotcpClick(Sender: TObject);
{$ENDIF}
    procedure CongesPayesRechercheTabletteAnciennete;
    procedure PCasRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: boolean);
{$IFNDEF DADSB}
    procedure BTnCompteBqClick(Sender: TObject);
{$ENDIF}
    procedure PagesChange(Sender: TObject);
    procedure CongesPayesClick(Sender: TObject);
    //       procedure CongesPayesCheckenter(Sender : TObject); PT40 suppression
    //       procedure NbJoutravEnter(Sender: TObject);
    //       procedure MBasanccpEnter(Sender : TObject); PT40
    //       procedure TypDatAncEnter(Sender : TObject);
    procedure PositionneFocusOngletCP;
    procedure BFermeclick(Sender: TObject);
    procedure InitialiseMvtCp(etab: string);
{$IFNDEF DADSB}
    procedure Calendclick(Sender: TObject);
{$ENDIF}
    procedure OnEnterProfil(Sender: Tobject); //PT32
    // PT10 PT4 : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
    procedure AccesNumIS(Sender: TObject); //PT16
    //   PT18   23/04/2002 V582 Ph suppression etablissement complémentaire
    function IsMouvementer: Boolean;
    // PT28   16/01/2003 V591 PH FQ 10432 gestion des taux at avec rafraichissement de la liste des taux
    procedure FTAUXDblClick(Sender: TObject);
    procedure AfficheTauxAT;
    procedure ChargeProfilActivite(Lechamp, Nat: string); //PT32
    procedure Imprimer(Sender: TObject);
    procedure ChargeVarGblIdemEtab; //PT46
    procedure AccesMedecine(Sender : TObject);
    procedure AccesDDtefp(Sender : TObject);
    procedure AjoutModifsHisto(Champ,ChampType,Valeur : String);
// d PT81
{$IFDEF EAGLCLIENT}
   procedure SaisieCodeLivr(Sender : TObject);
{$ENDIF}
// f PT81
  end;

const
  // libellés des messages
  TexteMessage: array[1..15] of string = (
    {1}'Vous ne pouvez pas supprimer l''établissement par défaut.',
    {2}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des salariés',
    {3}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des paies',
    {4}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des historiques',
    {5}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des intérimaires',
    {6}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des taux AT',
    {7}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des absences',
    {8}'Vous ne pouvez pas supprimer cet établissement, il est référencé par la saisie par rubrique',
    {9}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des organismes',
    {10}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des DUCS',
    {11}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des masques saisie par rubrique',
    {12}'Vous ne pouvez pas supprimer cet établissement, il est référencé par ces horaires',
    {13}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des donneurs d''ordre',
    {14}'Vous ne pouvez pas supprimer cet établissement, il est référencé par des commandes de titres restaurant', // PT35
    {15}'15'
    );

implementation

uses EntPaie,
  PgOutils2,
{$IFNDEF DADSB}
  PgOutils,
  PgCongesPayes,
  P5Util,
  P5Def,
{$ENDIF}
  PgCommun,
  PGCalendrier;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/06/2001
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE,PGETAB
*****************************************************************}

procedure TOM_EtabCompl.OnArgument(stArgument: string);
var
{$IFNDEF EAGLCLIENT}
  MDateF: THDBEdit;
  ISCheckBox1, ISCheckBox2: TDBCheckBox;
  MCombo: THDBVALComboBox;
{$ELSE}
  MDateF: THEdit;
  ISCheckBox1, ISCheckBox2: TCheckBox;
  MCombo: THVALComboBox;
{$ENDIF}
{$IFDEF CCS3}
  GBX: TGroupBox;
{$ENDIF}
  st, arg: string;
  ClotCp, DeClotCp, Btn, MFerme, TCalend: TToolBarButton97;
  Themes: THValCombobox;
  PCas: THGrid;
  i: Integer;
  Q: TQuery;
  MaPageControl: TPageControl;
  BtnFerme: TToolbarButton97;
  // PT10 PT4 : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
  FlisteTaux: THGrid;
{$IFDEF DADSB}
  TabShDADSB: TTabSheet;
{$ENDIF}

// d PT81
{$IFDEF EAGLCLIENT}
  Ticklivr : THEdit;
{$ENDIF}
// f PT81

begin
  inherited;
  Arg := stArgument;
  St := Trim(ReadTokenPipe(Arg, ';'));
  etablissement := St;
  DecocheCP := false;
  ForceValidCp := False; { PT55-1 }
  LoadFiche := false;
  // if BCJ <> NIL then BCJ.DataSource := 'SFiche';
  BtnFerme := TToolbarButton97(GetControl('BDELETE'));
  if BtnFerme <> nil then BtnFerme.ModalResult := 1;
  SetControlProperty('ETB_PGMODEREGLE', 'Datatype', 'PGMODEREGLE'); //PT7

  Decodedate(dATE, ANow, MNow, DNow);

  ExistCloture := ExisteSQL('SELECT PCN_SALARIE FROM ABSENCESALARIE ' +
    'LEFT JOIN SALARIES ON PSA_SALARIE=PCN_SALARIE ' +
    'WHERE PSA_ETABLISSEMENT="' + etablissement + '" ' +
    'AND PCN_GENERECLOTURE="X" AND PCN_TYPEMVT="CPA"');
{$IFNDEF EAGLCLIENT}
  MDateF := THDBEdit(getcontrol('ETB_DATECLOTURECPN'));
{$ELSE}
  MDateF := THEdit(getcontrol('ETB_DATECLOTURECPN'));
{$ENDIF}
  if MDateF <> nil then
  begin
    MDateF.OnExit := MDateFExit;
    MDateF.OnEnter := MDateFEnter;
  end;
  //PT62
{$IFNDEF DADSB}
  ClotCp := TToolBarButton97(getcontrol('TBCLOTCP'));
  if Clotcp <> nil then
    ClotCp.Onclick := ClotCpClick;

  DeClotCp := TToolBarButton97(getcontrol('TBDECLOTCP'));
  if DeClotcp <> nil then
    DeClotCp.Onclick := DeClotCpClick;

  Btn := TToolBarButton97(GetControl('BTNCOMPTEBQ'));
  if Btn <> nil then
    Btn.OnClick := BTnCompteBqClick;
{$ELSE}
  SetControlVisible('TBCLOTCP', False);
  SetControlVisible('TBDECLOTCP', False);
{$IFNDEF EAGLCLIENT}
  SetControlVisible('BTNCOMPTEBQ', False);
{$ENDIF}
{$ENDIF}
  //FIN PT62

    // ETABLISSEMENT := Copy (StArgument, 1, 3);
    // LIBELLE := Copy (StArgument,Pos(';',StArgument)+1, length(StArgument) - Pos(';',StArgument));
  LIBELLE := Trim(ReadTokenPipe(Arg, ';'));
  Q := OpenSQL('SELECT * FROM ETABCOMPL WHERE ETB_ETABLISSEMENT="' + ETABLISSEMENT + '"', FALSE);
  if Q.EOF then
  begin
    // PT37  : 17/09/2003 PH V_421 FQ 10110 creation etabcompl par une TOB au lieu de query et initilaisation des champs
    PGCreerEtabComp(ETABLISSEMENT, LIBELLE, ''); // PT57
    // FIN PT37
  end;
  Ferme(Q);

{PT85-2
  if Getfield('ETB_ETABLISSEMENT') <> ETABLISSEMENT then
    Setfield('ETB_ETABLISSEMENT', ETABLISSEMENT);
  if Getfield('ETB_LIBELLE') <> LIBELLE then
  begin
    Setfield('ETB_LIBELLE', Libelle);
    Ecran.caption := 'Etablissement : ' + Etablissement + ' ' + Libelle;
    UpdateCaption(TFFiche(Ecran));
  end;
}
    SetControlText('ETB_ETABLISSEMENT', ETABLISSEMENT);
    SetControlText('ETB_LIBELLE', Libelle);
    Ecran.caption := 'Etablissement : ' + Etablissement + ' ' + Libelle;
    UpdateCaption(TFFiche(Ecran));
//FIN PT85-2

  {VG 16/03/01 Gestion de l'affichage des onglets}
  NumOnglet := Trim(ReadTokenPipe(arg, ';'));
  {FIN VG 16/03/01 Gestion de l'affichage des onglets}

  Themes := THValCombobox(GetControl('THEME'));
  PCas := THGrid(GetControl('PCAS'));
  if (Themes <> nil) and (PCas <> nil) then
  begin
    PCas.RowCount := Themes.Items.Count + 1;
    for i := 0 to Themes.Items.Count - 1 do PCas.Cells[0, i + 1] := Themes.Items[i];
    PCas.OnRowEnter := PCasRowEnter;
  end;
  MFerme := TToolBarButton97(getcontrol('BFERME'));
  if MFerme <> nil then
    MFerme.OnClick := BFermeClick;  

  MaPageControl := TPageControl(getControl('Pages'));
  if MaPageControl <> nil then
    MaPageControl.OnChange := PagesChange;

  changeDateCloture := false;
  Cloture := false;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILCGE'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILCGE'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_RELIQUAT'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_RELIQUAT'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_1ERREPOSH'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_1ERREPOSH'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_2EMEREPOSH'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_2EMEREPOSH'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_BASANCCP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_BASANCCP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_TYPDATANC'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_TYPDATANC'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_VALORINDEMCP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_VALORINDEMCP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_MVALOMS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_MVALOMS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnKeyDown := ComboKeyDown;
  //PT62
{$IFNDEF DADSB}
  TCalend := TToolBarButton97(getcontrol('TCALENDRIER'));
  if TCalend <> nil then
    TCalend.Onclick := CalendClick;
{$ELSE}
  SetControlVisible('TCALENDRIER', False);
  SetControlVisible('TJOURFERIE', False);
  TabShDADSB := TTabSheet(GetControl('TBSHTPROFIL'));
  if TabShDADSB <> nil then
    TabShDADSB.TabVisible := False;
  TabShDADSB := TTabSheet(GetControl('PCONGES'));
  if TabShDADSB <> nil then
    TabShDADSB.TabVisible := False;
  TabShDADSB := TTabSheet(GetControl('PRIB'));
  if TabShDADSB <> nil then
    TabShDADSB.TabVisible := False;
  TabShDADSB := TTabSheet(GetControl('TBSHTTAUXAT'));
  if TabShDADSB <> nil then
    TabShDADSB.TabVisible := True;
  TabShDADSB := TTabSheet(GetControl('PINTERMITTENTS'));
  if TabShDADSB <> nil then
    TabShDADSB.TabVisible := False;
  SetControlEnabled('ETB_CODESECTION', False);
  SetControlEnabled('FLTAUXAT', False);
  SetControlEnabled('GBSALARIE', False);
  SetControlEnabled('GBETAB', False);
  SetControlEnabled('ETB_PRUDHCOLL', False);
  SetControlEnabled('ETB_PRUDHSECT', False);
  SetControlEnabled('ETB_PRUDHVOTE', False);
  SetControlEnabled('ETB_PRUDH', False);
{$ENDIF}
  //FIN PT62
    // PT10 PT4 : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
  // FIN PT4  et PT10
//PT62
{$IFNDEF DADSB}
  if VH_Paie.PGIntermittents = True then //PT16  Affichage de l'onglet que si gestion (booléen ds paramsoc)
  begin
    SetControlVisible('PINTERMITTENTS', true);
    SetControlVisible('GBINTERMITTENTS', true);
{$IFNDEF EAGLCLIENT}
    ISCheckBox1 := TDBCheckBox(getcontrol('ETB_ISLICSPEC'));
    ISCheckBox2 := TDBCheckBox(getcontrol('ETB_ISLABELP'));
{$ELSE}
    ISCheckBox1 := TCheckBox(getcontrol('ETB_ISLICSPEC'));
    ISCheckBox2 := TCheckBox(getcontrol('ETB_ISLABELP'));
{$ENDIF}
    if ISCheckBox1 <> nil then ISCheckBox1.OnClick := AccesNumIS;
    if ISCheckBox2 <> nil then ISCheckBox2.OnClick := AccesNumIS;
  end;
  if VH_Paie.PGMsa = True then
  begin
    SetControlVisible('PINTERMITTENTS', true);
    SetControlVisible('GBMSA', true);
  end
  else
  //DEBUT PT74
  begin
    SetControlProperty('GBINTERMITTENTS','Top',TGroupBox(GetControl('GBMSA')).Top);
    SetControlProperty('GBINTERMITTENTS','Top',TGroupBox(GetControl('GBMSA')).Top);
    SetControlCaption('PINTERMITTENTS','Intermittents');
  end;
  //FIN PT74
{$ENDIF}
//FIN PT62
  //PT21
{$IFDEF CCS3}
  SetControlVisible('TETB_AUTRENUMERO', FALSE);
  SetControlVisible('ETB_AUTRENUMERO', FALSE);
  SetControlVisible('TETB_CODEDDTEFPGU', FALSE);
  SetControlVisible('ETB_CODEDDTEFPGU', FALSE);
//PT63
  SetControlVisible('TETB_TYPTICKET', FALSE);
  SetControlVisible('ETB_TYPTICKET', FALSE);
  SetControlVisible('ETB_SUBROGATION', FALSE);
//FIN PT63
  {PT29
    SetControlVisible('TETB_RIBACOMPTE',FALSE);
    SetControlVisible('ETB_RIBACOMPTE',FALSE);
    SetControlVisible('GRBXREGLE',FALSE);
  }
  SetControlVisible('TETB_PAIFRAIS', FALSE);
  SetControlVisible('ETB_PAIFRAIS', FALSE);
  GBX := TGroupBox(getcontrol('GRBXREGLE'));
  if GBX <> nil then
    GBX.Left := 170;
  //FIN PT29
  SetControlVisible('TETB_RIBFRAIS', FALSE);
  SetControlVisible('ETB_RIBFRAIS', FALSE);
  SetControlVisible('GRBXSALAIRES1', FALSE);
  SetControlVisible('GRBXSALAIRESEDT', FALSE);
{PT64
  SetControlVisible('PINTERMITTENTS', FALSE);
}  
  GBX := TGroupBox(getcontrol('GRBXSALAIRES'));
  if GBX <> nil then
    GBX.Left := 170;
{$ENDIF}
  //FIN PT21
  // PT28   16/01/2003 V591 PH FQ 10432 gestion des taux at avec rafraichissement de la liste des taux
  FlisteTaux := THGrid(GetControl('FLTAUXAT'));
  if FlisteTaux <> nil then FlisteTaux.OnDblClick := FTAUXDblClick;

  { DEB PT32 gestion d'evenement pour raffraichir tablette }
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILAFP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILAFP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILANCIEN'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILANCIEN'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILAPP'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILAPP'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILCGE'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILCGE'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILMUT'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILMUT'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PERIODBUL'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PERIODBUL'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILPRE'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILPRE'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFIL'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFIL'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILRBS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILRBS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILREM'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILREM'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILRET'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILRET'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_REDREPAS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_REDREPAS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_REDRTT1'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_REDRTT1'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_REDRTT2'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_REDRTT2'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILTRANS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILTRANS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_PROFILTSS'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_PROFILTSS'));
{$ENDIF}
  if MCombo <> nil then MCombo.OnEnter := OnEnterProfil;
  { FIN PT32 }

//PT44
// Gestion de l'impression
  Imprim := TToolBarButton97(GetControl('BImprimer'));
  if Imprim <> nil then
  begin
    Imprim.Visible := True;
    Imprim.OnClick := Imprimer;
  end;
  //FIN PT44

  Btn := TToolBarButton97(GetControl('BMEDECINE'));
  If Btn <> Nil then Btn.OnClick := AccesMedecine;
  Btn := TToolBarButton97(GetControl('BDDTEFP'));
  If Btn <> Nil then Btn.OnClick := AccesDDtefp;

// d PT81
  // Tickets restaurant (ACCOR) - Code point de livraison
  fournisseur := '';
  if (VH_Paie.PGTicketRestau) then
  begin
    fournisseur := GetParamSocSecur ('SO_PGTYPECDETICKET','');
    if ((fournisseur = '003') or (fournisseur = '004')) then    // pt88
    // ACCOR  ou chèque déjeuner
    begin
      if (V_PGI.NumVersionBase < 800) then
      // (spécif V7)
      begin
        if (not VH_Paie.PGMsa) then
        // on ne traite le code point de livraion que si dossier non MSA
        begin
          SetControlVisible('TETB_TICKLIVR', TRUE);
          SetControlVisible('ETB_AUTRENUMERO_', TRUE);
          SetControlEnabled('ETB_AUTRENUMERO_', TRUE);
          SetControlVisible('TETB_AUTRENUMERO', FALSE);
          SetControlVisible('ETB_AUTRENUMERO', FALSE);
          SetControlEnabled('ETB_AUTRENUMERO', FALSE);
{$IFDEF EAGLCLIENT}
          Ticklivr := THEdit(GetControl('ETB_AUTRENUMERO_'));
          TickLivr.OnExit := SaisieCodeLivr;
{$ENDIF}
        end
        else
        // dossier MSA - code point de livraison non traité
        begin
          SetControlVisible('TETB_TICKLIVR', FALSE);
          SetControlVisible('ETB_AUTRENUMERO_', FALSE);
          SetControlEnabled('ETB_AUTRENUMERO_', FALSE);
        end;
      end
      else
      // ( A partir V8)
      begin
        SetControlVisible('TETB_TICKLIVR', TRUE);
        SetControlVisible('ETB_TICKLIVR', TRUE);
        SetControlEnabled('ETB_TICKLIVR', TRUE);
        // deb pt88
        if (fournisseur = '004') then
        begin
          if (VH_Paie.PGFactEtabl) then
          SetcontrolCaption('TETB_TICKLIVR', 'Code client')
          else
          begin
           SetControlVisible('TETB_TICKLIVR', false);
           SetControlVisible('ETB_TICKLIVR', false);
           SetControlEnabled('ETB_TICKLIVR', false);
          end;
         end;
         // fin pt88
      end;
    end
    else
    // SODEXHO, NATEXIS
    begin
      if (V_PGI.NumVersionBase < 800) then
      begin
        SetControlVisible('TETB_TICKLIVR', FALSE);
        SetControlVisible('ETB_AUTRENUMERO_',FALSE);
        SetControlEnabled('ETB_AUTRENUMERO_',FALSE);
      end
      else
      begin
        SetControlVisible('TETB_TICKLIVR', FALSE);
        SetControlVisible('ETB_TICKLIVR', FALSE);
        SetControlEnabled('ETB_TICKLIVR', FALSE);
      end;
    end;
  end;
// f PT81

//PT85-1
SetPlusBanqueCP (GetControl ('ETB_RIBSALAIRE'));
SetPlusBanqueCP (GetControl ('ETB_RIBFRAIS'));
SetPlusBanqueCP (GetControl ('ETB_RIBACOMPTE'));
SetPlusBanqueCP (GetControl ('ETB_RIBDUCSEDI'));
//FIN PT85-1
  SetControlVisible('DATEHISTO',False); //PT91
end;

//PT62
{$IFNDEF DADSB}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/06/2001
Modifié le ... :   /  /
Description .. : Click sur le calendrier
Mots clefs ... : PAIE,PGETAB
*****************************************************************}

procedure TOM_EtabCompl.Calendclick(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  MCombo: THDBValComboBox;
{$ELSE}
  MCombo: THValComboBox;
{$ENDIF}
  TobNom: tob;
begin
  TobNom := Tob.create('NomCalend', nil, -1);
  if TobNom <> nil then
    TobNom.AddChampSup('CALENDRIER_MAJ', false);
  TheTob := TobNom;
{$IFNDEF EAGLCLIENT}
  MCombo := THDBValComboBox(getcontrol('ETB_STANDCALEND'));
{$ELSE}
  MCombo := THValComboBox(getcontrol('ETB_STANDCALEND'));
{$ENDIF}
  if MCombo <> nil then
  begin
    AglLanceFiche('AFF', 'HORAIRESTD', '', '', 'CODE:' + MCombo.value + ','';'';'';STANDARD:' + MCombo.value);
    MCombo.Value := TobNom.getvalue('CALENDRIER_MAJ');
  end;
  TheTob := nil;
  //if notVide(TobNom, false) then TobNom.free; PT76
  FreeAndNil(TobNom);
end;
{$ENDIF}
//FIN PT62


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 12/06/2001
Modifié le ... :   /  /
Description .. : Click sur BFermer
Mots clefs ... : PAIE,PGETAB
*****************************************************************}

procedure TOM_EtabCompl.BFermeclick(Sender: TObject);
begin
  Ecran.Close;
end;

procedure TOM_EtabCompl.ComboKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F5:
      begin
        Key := 0;
{$IFNDEF EAGLCLIENT}
        if Sender is THDBValcombobox then THdbValcombobox(Sender).DroppedDown := True;
{$ELSE}
        if Sender is THValcombobox then THValcombobox(Sender).DroppedDown := True;
{$ENDIF}
      end;
  end;
end;

{PT39
procedure TOM_EtabCompl.OnNewRecord;
var
DateFerm : tdatetime;
begin
inherited;

TFFiche(Ecran).FTypeAction := taCreat ;

InitialiseZonesCongesPayes;
SetField('ETB_CONGESPAYES','X');
DateFerm := EncodeDate(ANow, 05, 31);
if DateFerm <= Now then
   DateFerm := EncodeDate(ANow+1, 05, 31);
SetField('ETB_DATECLOTURECPN', DateFerm);
SetField('ETB_JOURHEURE',VH_Paie.PGJourHeure);
//  PT6 : 10/09/2001  PH V547 Initialisation CP coché par defaut
SetField('ETB_CONGESPAYES','X');
SetField ('ETB_BCJOURPAIEMENT',0);
//   PT12 : 21/11/2001 V563 PH Initialisation des bnouveaux champs
SetField ('ETB_JEDTDU',0);
SetField ('ETB_JEDTAU',0);
SetField ('ETB_PRUDHCOLL','1');
SetField ('ETB_PRUDHSECT','4');
SetField ('ETB_PRUDHVOTE','1');
// PT 16 Initialisation des champs MSA et intermittents
SetField('ETB_ISLICSPEC','-');
SetField('ETB_ISOCCAS','-');
SetField('ETB_ISLABELP','-');
end;
}

procedure TOM_EtabCompl.MDateFEnter(Sender: TObject);
var
{$IFNDEF EAGLCLIENT}
  Medit: thdbedit;
{$ELSE}
  Medit: thedit;
{$ENDIF}
begin
{$IFNDEF EAGLCLIENT}
  Medit := Thdbedit(getcontrol('ETB_DATECLOTURECPN'));
{$ELSE}
  Medit := Thedit(getcontrol('ETB_DATECLOTURECPN'));
{$ENDIF}
  if Medit = nil then exit;
  if Isvaliddate(Medit.text) then
    DateEnter := getfield('ETB_DATECLOTURECPN')
  else DateEnter := 0;
end;

procedure TOM_EtabCompl.MDateFExit(Sender: TObject);
var
  init: word;
begin

  if ((GetField('ETB_DATECLOTURECPN') <> null) and (GetField('ETB_DATECLOTURECPN') <> Idate1900) // beurk
    and (GetField('ETB_DATECLOTURECPN') <> DateEnter) and (DateEnter <> 0)) then
    Init := HShowMessage('1;Congés payés;Attention, ' +
      'la modification de la date de clôture va décaler tous les calculs congés payés.' +
      '#13#10 Si vous souhaitez changer de période, clôturez celle ci pour ouvrir la suivante.' +
      '#13#10 Etes vous sûr de vouloir continuer ?;Q;YN;N;N;', '', '')
  else
  begin
    ChangeDatecloture := True;
    if IsValidDate(GetControlText('ETB_DATECLOTURECPN')) then //DEB PT15
      DateEnter := getfield('ETB_DATECLOTURECPN')
    else
      DateEnter := 0; //FIN PT15
    exit;
  end;
  if Init = mrYes then // J'affecte FinPeriodeN avec FinPeriodeS
  begin
    DS.edit;
    //    key := '*';
    //    ParamDate (Ecran, Sender, Key);
    ChangeDatecloture := True;
    DateEnter := getfield('ETB_DATECLOTURECPN');
  end;
  if Init = mrNo then
    Setfield('ETB_DATECLOTURECPN', DateEnter);
end;

{ PT55-1 Refonte procedure CongesPayesClick }

procedure TOM_EtabCompl.CongesPayesClick(Sender: TObject);
var
  Init, InitCp: word;
  DateFerm: Tdatetime;
  CongesPayes : String;
begin
  { DEB PT68 }
  //  if Ds.state in [DsBrowse] then exit;
  CongesPayes := GetField('ETB_CONGESPAYES');
  if  (CongesPayes = GblCongesPayes) or (GblCongesPayes = '') then  exit
  else  GblCongesPayes := CongesPayes;
  { FIN PT68 }

  if (not VH_paie.PGCongesPayes) and (GetField('ETB_CONGESPAYES') = 'X') then
  begin
    PgiBox('Veuillez cocher l''option gestion des congés payés au niveau dossier avant le niveau établissement.', 'Gestion des congés payés impossible!'); { PT52-1 }
    SetField('ETB_CONGESPAYES', '-');
    EnableChampsOngletCongesPayes;
    exit;
  end;

  if DecocheCP then exit;

  if (VH_paie.PGCongesPayes) then
  begin
    if (GetField('ETB_CONGESPAYES') = 'X') then
    begin { DEB PT52-1 }
      if ExisteSql('SELECT PSA_SALARIE FROM SALARIES WHERE PSA_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '" AND PSA_CONGESPAYES="-"') then
      begin
        InitCp := HShowMessage('1;Congés payés;Voulez-vous activer la gestion des congés payés pour les salariés non affectés?;Q;YNC;N;N;', '', '');
        if InitCp = mrYes then
        begin
          ExecuteSql('UPDATE SALARIES SET PSA_CONGESPAYES="X" WHERE PSA_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"');
          ForceValidCp := True;
        end;
        EnableChampsOngletCongesPayes;
      end; { FIN PT52-1 }
      DateFerm := getfield('ETB_DATECLOTURECPN');
      if DateFerm <= Idate1900 then
      begin
        DateFerm := EncodeDate(ANow, 05, 31);
        if DateFerm <= Now then DateFerm := EncodeDate(ANow + 1, 05, 31);
        Setfield('ETB_DATECLOTURECPN', DateFerm);
      end;
      // PT59 Message info
      PgiBox('Attention, la date de clôture des CP sera le ' + DateToStr(Dateferm) + ',#13#10il faudra la modifier avant de faire vos paies si vous reprenez un dossier avant le 1er juin', Ecran.caption);
    end
    else
      if (GetField('ETB_CONGESPAYES') = '-') then
    begin
      Init := HShowMessage('1;Congés payés;Attention, vous désactivez la gestion des congés payés pour cet établissement. Voulez-vous continuer ?;Q;YNC;N;N;', '', '');
      if Init = mrYes then
      begin
        InitCp := HShowMessage('1;Congés payés;Souhaitez-vous une initialisation des mouvements de congés payés ?;Q;YNC;N;N;', '', '');
        if InitCp = mrYes then begin InitialiseMvtCp(etablissement); ExistCloture := False; end; { PT77-1 }
        ExecuteSql('UPDATE SALARIES SET PSA_CONGESPAYES="-" WHERE PSA_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"'); { PT52-1 }
        ForceValidCp := True;
        InitialiseZonesCongesPayes;
        DecocheCP := true;
        SetFocusControl('PCONGES');
      end
      else
        SetField('ETB_CONGESPAYES', 'X');
    end;
    EnableChampsOngletCongesPayes;
  end;
end;


//PT62
{$IFNDEF DADSB}

procedure TOM_EtabCompl.ClotCpClick(Sender: TObject);
var
  FinPeriodeN, DebPeriodeS, FinPeriodeS: Tdatetime;
  Init, aa, mm, jj: word;
  DatecloturecpN: thedit;
  MethodeRel, St : string;
  BValide: ttoolbarbutton97;
  TEvent: TStringList; //PT26
begin
// DEB PT67
  St := 'La clôture des congés payés doit être exécutée dans le module Paie#13#10'+
        'Menu Salariés - commande Congés payés-Clôture des congés payés';
  PgiBox (St, 'Clôture des congés payés');
  Exit;
// FIN PT67 Je laisse le code dessous au cas où ....
  DatecloturecpN := thedit(Getcontrol('ETB_DATECLOTURECPN'));
  if DatecloturecpN <> nil then
  begin
    if IsValidDate(getfield('ETB_DATECLOTURECPN')) then
      FinPeriodeN := getfield('ETB_DATECLOTURECPN')
    else
      exit;
  end
  else
    exit;
  DebPeriodeS := FinPeriodeN + 1;
  decodedate(FinPeriodeN, aa, mm, jj);
  FinPeriodeS := PGEncodeDateBissextile(aa + 1, mm, jj); { PT48 }
  Init := HShowMessage('2;Congés payés;Attention, vous allez clôturer la période de congés payés #13#10et ouvrir la suivante ( du ' + datetostr(DebPeriodeS) + ' au' +
    datetostr(finPeriodeS) + ').#13#10 '
    + 'Etes-vous sûr d''avoir terminé vos paies jusqu''au ' + datetostr(FinPeriodeN) +
    ' ?#13#10 Attention, la clôture entraîne l''écriture de la fiche établissement;Q;YNC;N;N;', '', '');
  if Init = mrYes then // J'affecte FinPeriodeN avec FinPeriodeS
  begin { DEB PT56 09/06/2004 }
    if ExisteSql('SELECT PPU_SALARIE FROM PAIEENCOURS WHERE PPU_DATEFIN>="' + USDateTime(DebutDeMois(FinPeriodeN)) + '" ' +
      'AND PPU_DATEFIN <= "' + USDateTime(FinPeriodeN) + '"') = False then
      if PgiAsk('Attention! Vous n''avez pas effectué les paies sur tous les mois #13#10 ' +
        'de la période congés payés allant du ' + DateToStr(DebutDeMois(PlusMois(FinPeriodeN, -11))) +
        ' au ' + DateToStr(FinPeriodeN) + '!#13#10 Voulez-vous clôturer?', 'Clôture congés payés') = MrNo then
        exit; { FIN PT56 }
    DS.edit;
    //DEB PT26 Initialisation de TEVENT
    TEvent := TStringList.Create;
    TEvent.Add('Clôture des congés payés de l''établissement : ' + etablissement + ' ' + GetField('ETB_LIBELLE'));
    TEvent.Add('Date de clôture au : ' + DateToStr(FinPeriodeN));
    //FIN PT26
    Setfield('ETB_DATECLOTURECPN', FinPeriodeS);
    MethodeRel := getfield('ETB_RELIQUAT');
    if CloturePeriodeCP(FinPeriodeN, etablissement, MethodeRel) then
    begin
      Cloture := true;
      SetControlEnabled('TBDECLOTCP', False);
      SetControlEnabled('ETB_DATECLOTURECPN', False);
      BValide := ttoolbarbutton97(getcontrol('BVALIDER'));
      if BValide <> nil then BValide.OnClick(Sender);
      //DEB PT26  Génération évènement
      TEvent.Add('Clôture des congés payés OK.');
      CreeJnalEvt('002', '012', 'OK', nil, nil, TEvent);
    end
    else
    begin
      TEvent.Add('Une erreur est survenue lors de la clôture des congés payés.');
      CreeJnalEvt('002', '012', 'ERR', nil, nil, TEvent);
    end;
    //FIN PT26
    if TEvent <> nil then TEvent.free; //PT26
  end;
end;

procedure TOM_EtabCompl.DeClotcpClick(Sender: TObject);
var
  FinPeriodeN, FinPeriodeN1: Tdatetime;
  DatecloturecpN: thedit;
  BValide: ttoolbarbutton97;
  TEvent: TStringList; //PT26
begin
  DatecloturecpN := thedit(Getcontrol('ETB_DATECLOTURECPN'));
  if DatecloturecpN <> nil then
  begin
    if IsValidDate(getfield('ETB_DATECLOTURECPN')) then
      FinPeriodeN := getfield('ETB_DATECLOTURECPN')
    else
      exit;
  end
  else exit;
  //DEB PT26  Initialisation de TEVENT
  TEvent := TStringList.Create;
  TEvent.Add('Declôture des congés payés de l''établissement : ' + etablissement + ' ' + GetField('ETB_LIBELLE'));
  TEvent.Add('Date de declôture au : ' + DateToStr(FinPeriodeN));
  //FIN PT26
  {DEB PT25 Utilisation d'une procedure  pour la decloture des CP
  suite création mul de cloture et decloture multietablissement}
  FinPeriodeN1 := PlusDate(FinPeriodeN, -1, 'A');
  if DeCloturePeriodeCP(FinPeriodeN, etablissement, GetField('ETB_LIBELLE')) then
  begin
    Setfield('ETB_DATECLOTURECPN', FinPeriodeN1);
    BValide := ttoolbarbutton97(getcontrol('BVALIDER'));
    if BValide <> nil then BValide.OnClick(Sender);
    PGiBox('Traitement Terminé.', '"Declôture" période CP');
    //DEB PT26   Génération évènement
    TEvent.Add('Declôture des congés payés OK.');
    CreeJnalEvt('002', '013', 'OK', nil, nil, TEvent);
  end
  else
  begin
    TEvent.Add('Une erreur est survenue lors de la declôture des congés payés.');
    CreeJnalEvt('002', '013', 'ERR', nil, nil, TEvent);
  end;
  //FIN PT26
//FIN PT25
  if TEvent <> nil then TEvent.free; //PT26
end;
{$ENDIF}
//FIN PT62


procedure TOM_Etabcompl.OnLoadRecord;
var
  TCalend, TCompte, TFerie: TToolBarButton97; {VG 16/03/01 Gestion de l'affichage des onglets}
  TCar, TPro, TCon, TReg, TDad, TInter: TTabSheet; {VG 16/03/01 Gestion de l'affichage des onglets}
  EtabPage: TPageControl; {VG 16/03/01 Gestion de l'affichage des onglets}
  // PT10 PT4 : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
  Etab : string;
begin
  inherited;
  GblCongesPayes := GetField('ETB_CONGESPAYES') ; { PT68 }
  ChargementTablette('BQ', ''); //rechargement des tablettes banques
  ChargementTablette('PQ', '');
  ChargementTablette('ET', '');
  ChargementTablette('ETB', '');
  {-----------------------------------------------------------------------
  INIT DES CHAMPS PROFIL IDEMETABL de la fiche SALARIES }
  // PT10 PT4 : 31/08/2001 V547 PH Rajout champs gestion medecine du travail et DDTEFP
  Etab := GetField('ETB_ETABLISSEMENT');


  //**SB
  ChargeVarGblIdemEtab; //PT46
  SetControlProperty('ETB_EDITORG', 'Plus', GetField('ETB_ETABLISSEMENT')); //PT24-1

  //------------------------------------------------------------------
  LoadFiche := True;
  //  EDT_ETABLISSEMENT := THDBEdit(GetControl ('ETB_ETABLISSEMENT'));
   // if EDT_ETABLISSEMENT <> NIL then EDT_ETABLISSEMENT.Text := ETABLISSEMENT;
   // EDT_LIBELLE := THDBEdit(GetControl ('ETB_LIBELLE'));
   // if EDT_LIBELLE <> NIL then EDT_LIBELLE.Text := LIBELLE;

    {VG 16/03/01 Gestion de l'affichage des onglets}
  if (IsNumeric(NumOnglet)) and (NumOnglet <> '0') and (NumOnglet <> '') then //PT9-2 SB 19/10/2001 Test si numeric au cas ou je modifie le onargument
  begin
    EtabPage := TPageControl(GetControl('Pages'));
    if EtabPage <> nil then
      EtabPage.ActivePageIndex := (StrToInt(NumOnglet) - 1);
    TCar := TTabSheet(GetControl('PGeneral'));
    if TCar <> nil then
      TCar.TabVisible := TCar.PageIndex = (StrToInt(NumOnglet) - 1);
    TPro := TTabSheet(GetControl('TBSHTPROFIL'));
    if TPro <> nil then
      TPro.TabVisible := TPro.PageIndex = (StrToInt(NumOnglet) - 1);
    TCon := TTabSheet(GetControl('PCONGES'));
    if TCon <> nil then
    begin
      TCon.TabVisible := TCon.PageIndex = (StrToInt(NumOnglet) - 1);
      TCalend := TToolBarButton97(getcontrol('TCALENDRIER'));
      if TCalend <> nil then
        TCalend.Visible := TCon.PageIndex = (StrToInt(NumOnglet) - 1);
      TFerie := TToolBarButton97(getcontrol('TJOURFERIE'));
      if TFerie <> nil then
        TFerie.Visible := TCon.PageIndex = (StrToInt(NumOnglet) - 1);
    end;
    TReg := TTabSheet(GetControl('PRIB'));
    if TReg <> nil then
    begin
      TReg.TabVisible := TReg.PageIndex = (StrToInt(NumOnglet) - 1);
      TCompte := TToolBarButton97(getcontrol('BTNCOMPTEBQ'));
      if TCompte <> nil then
        TCompte.Visible := TReg.PageIndex = (StrToInt(NumOnglet) - 1);
    end;
    TDad := TTabSheet(GetControl('TBSHTTAUXAT'));
    if TDad <> nil then
      TDad.TabVisible := TDad.PageIndex = (StrToInt(NumOnglet) - 1);
    {PT19}
    TInter := TTabSheet(GetControl('PINTERMITTENTS'));
    if TInter <> nil then
      TInter.TabVisible := TInter.PageIndex = (StrToInt(NumOnglet) - 1);
    {FIN PT19 }
  end;
  {FIN VG 16/03/01 Gestion de l'affichage des onglets}

  EnableChampsOngletCongesPayes;
  { // Beurk
    Themes:=THValCombobox(GetControl('THEME')) ;
    PPart:=THValCombobox(GetControl('PPART')) ;
    PCas:=THGrid(GetControl('PCAS')) ;
    if (PPart=Nil) or (PCas=Nil) or (Themes=Nil) then exit ;
    PPart.Plus:='' ;
    TPS:=TOB.Create('profils speciaux',Nil,-1) ;
    TPS.LoadDetailDB('PROFILSPECIAUX','"X";"'+GetField('ETB_ETABLISSEMENT')+'"','',Nil,FALSE,FALSE) ;
    For i:=0 to TPS.Detail.Count-1 do
      BEGIN
      Theme:=TPS.Detail[i].GetValue('PPS_THEMEPROFIL') ;
      PPart.Value:=TPS.Detail[i].GetValue('PPS_PROFIL') ;
      Profil:=PPart.Text ;
      for j:=0 to Themes.Items.Count-1 do
         if Theme=Themes.Values[j] then BEGIN PCas.Cells[1,j+1]:=Profil ; break ; END ;
      END ;
    TPS.Free ;
  }
  //PT16 Contrôle des booléen pour accès différents N°
  if GetField('ETB_ISLICSPEC') = 'X' then SetControlVisible('ETB_ISNUMLIC', True) //PT22
  else SetControlVisible('ETB_ISNUMLIC', False);
  if GetField('ETB_ISLABELP') = 'X' then SetControlVisible('ETB_ISNUMLAB', True)
  else SetControlVisible('ETB_ISNUMLAB', False);
  // PT28   16/01/2003 V591 PH FQ 10432 gestion des taux at avec rafraichissement de la liste des taux
  AfficheTauxAT;
  //PT40 Raffraîchissement des zones CP si modificationdu type d'alimentation
  GblBasAnccp := GetField('ETB_BASANCCP');
  GblNbJoutrav := GetField('ETB_NBJOUTRAV');
  GblTypDatAnc := GetField('ETB_TYPDATANC');
  ConventionColl := GetField('ETB_CONVENTION'); //PT71
  UGMsa := GetField('ETB_MSAUNITEGES'); //PT72
// d PT81
  if ((VH_Paie.PGTicketRestau) and ((fournisseur = '003') or (fournisseur = '004'))) then   // pt88
  // ACCOR  ou chèque déjeuner // pt88
  begin
    if (V_PGI.NumVersionBase < 800) then
    // (spécif V7)
    begin
      PointLivraison := GetField('ETB_AUTRENUMERO');
// d PT81
{$IFDEF EAGLCLIENT}
      SetControlText ('ETB_AUTRENUMERO_',PointLivraison);
      SetField ('ETB_AUTRENUMERO',PointLivraison);
{$ENDIF}
// f PT81
    end
    else
    // (à partir V8)
    begin
      PointLivraison := GetField('ETB_TICKLIVR');
    end;
  end;
// f PT81
//PT87
College:= GetField('ETB_PRUDHCOLL');
Section:= GetField('ETB_PRUDHSECT');
LieuVote:= GetField('ETB_PRUDHVOTE');
//FIN PT87
end;

//PT62
{$IFNDEF DADSB}

procedure TOM_Etabcompl.OnUpdateRecord;
var
  st, st1: string;
  Etab, nb: string;
  inb: integer;
  LeLib, LaVariable: string;
// d PT80
{$IFNDEF EAGLCLIENT}
  MCombo: THDBValComboBox;
{$ELSE}
  MCombo: THValComboBox;
{$ENDIF}
  ErrBq : boolean;
// f PT80
  PointLivr       :     string;        // PT81
  GereHisto : Boolean;
begin
  inherited;
//d PT80
  if (GetActiveTabSheet('Pages').name = 'PRIB') then
  begin
    ErrBq := false;

{$IFNDEF EAGLCLIENT}
    MCombo := THDBValComboBox(GetControl('ETB_RIBSALAIRE'));
{$ELSE}
    MCombo := THValComboBox(GetControl('ETB_RIBSALAIRE'));
{$ENDIF}
    if (((MCombo.Text <> '') and (MCombo.Text <> ' ')) and
        ((GetField('ETB_RIBSALAIRE') = '') or (GetField('ETB_RIBSALAIRE') = ' ')))  then
      ErrBq := true;

{$IFNDEF EAGLCLIENT}
    MCombo := THDBValComboBox(GetControl('ETB_RIBFRAIS'));
{$ELSE}
    MCombo := THValComboBox(GetControl('ETB_RIBFRAIS'));
{$ENDIF}
    if (((MCombo.Text <> '') and (MCombo.Text <> ' ')) and
        ((GetField('ETB_RIBFRAIS') = '') or (GetField('ETB_RIBFRAIS') = ' '))) then
      ErrBq := true;

{$IFNDEF EAGLCLIENT}
    MCombo := THDBValComboBox(GetControl('ETB_RIBACOMPTE'));
{$ELSE}
    MCombo := THValComboBox(GetControl('ETB_RIBACOMPTE'));
{$ENDIF}
    if (((MCombo.Text <> '') and (MCombo.Text <> ' ')) and
        ((GetField('ETB_RIBACOMPTE') = '') or (GetField('ETB_RIBACOMPTE') = ' '))) then
      ErrBq := true;

{$IFNDEF EAGLCLIENT}
    MCombo := THDBValComboBox(GetControl('ETB_RIBDUCSEDI'));
{$ELSE}
    MCombo := THValComboBox(GetControl('ETB_RIBDUCSEDI'));
{$ENDIF}
    if (((MCombo.Text <> '') and (MCombo.Text <> ' ')) and
        ((GetField('ETB_RIBDUCSEDI') = '') or (GetField('ETB_RIBDUCSEDI') = ' '))) then
      ErrBq := true;

    if ErrBq then
      PGIBox('Vérifier le paramètrage des banques sélectionnées (N° compte général non renseigné)', 'Identifiants bancaires');
  end;

// d PT79
{   if (GetField('ETB_RIBSALAIRE') = '') or (GetField('ETB_RIBSALAIRE') = ' ') then
     begin
       PGIBox('Vérifier le paramètrage de la banque sélectionnée (N° compte général non renseigné)', 'Banque salaires');
     end;

   if (GetField('ETB_RIBFRAIS') = '') or (GetField('ETB_RIBFRAIS') = ' ') then
     begin
       PGIBox('Vérifier le paramètrage de la banque sélectionnée ((N° compte général non renseigné)', 'Banque frais');
     end;

   if (GetField('ETB_RIBACOMPTE') = '') or (GetField('ETB_RIBACOMPTE') = ' ') then
     begin
       PGIBox('Vérifier le paramètrage de la banque sélectionnée (N° compte général non renseigné)', 'Banque acomptes');
     end;
   if (GetField('ETB_RIBDUCSEDI') = '') or (GetField('ETB_RIBDUCSEDI') = ' ') then
     begin
       PGIBox('Vérifier le paramètrage de la banque sélectionnée (N° compte général non renseigné)', 'Banque charges sociales');
     end;*}
// f PT79
// f PT80
  // PT36  : 31/07/2003 PH V_42 FQ 10201 controle contenu champ test au minimum
  if GetField('ETB_TYPSMIC') = '' then SetField('ETB_SMIC', '');
  if GetField('ETB_SMIC') = '' then
  begin
    if GetField('ETB_TYPSMIC') <> '' then
    begin
      LastError := 1;
      PgiBox('Vous devez renseigner la valeur minimum', TFFiche(Ecran).Caption);
      SetFocusControl('ETB_SMIC');
      exit;
    end;
  end;
  // FIN PT36
  // PT43  : 13/10/2003 PH V_421 FQ 10877 test de la valeur saisie dans zone test au minimum
  if GetField('ETB_TYPSMIC') <> '' then
  begin
    if GetField('ETB_TYPSMIC') = 'ELN' then St1 := 'PGELEMENTNAT'
    else if GetField('ETB_TYPSMIC') = 'VAR' then St1 := 'PGVARIABLE'
    else st1 := '';
    if st1 <> '' then
    begin
      st := RechDom(St1, GetField('ETB_SMIC'), FALSE);
      if (St = '') or (St = 'Error') then
      begin
        LastError := 1;
        PgiBox('Vous devez renseigner la valeur minimum', TFFiche(Ecran).Caption);
        SetFocusControl('ETB_SMIC');
        exit;
      end;
    end;
  end;
  // Fin PT43
  // DEB PT58
  LaVariable := GetField('ETB_NBACQUISCP');
  if LaVariable <> '' then
  begin
    LeLib := RechDom('PGVARIABLE', LaVariable, FALSE);
    if (LeLib = '') or (LeLib = 'Error') then
    begin
      LastError := 1;
      PgiBox('Vous devez renseigner le champ jour acquis bulletin avec une variable existante', TFFiche(Ecran).Caption);
      SetFocusControl('ETB_NBACQUISCP');
    end;
  end;
  // FIN PT58
  if GetField('ETB_HORAIREETABL') = 0 then //DEBUT PT20
  begin
    LastError := 1;
    PgiBox('Vous devez renseigner l''horaire de référence de l''établissement', TFFiche(Ecran).Caption);
    SetFocusControl('ETB_HORAIREETABL');
  end; //FIN PT20
  ControleZonesOngletCongesPayes;
  if ((changedatecloture) and (not cloture)) and
    (Getfield('ETB_DATECLOTURECPN') <> null) and
    (Getfield('ETB_DATECLOTURECPN') <> Idate1900) then // beurk
  begin
    RecalculePeriode(Getfield('ETB_ETABLISSEMENT'), Getfield('ETB_DATECLOTURECPN'));
    changedatecloture := false;
  end;

// d PT81
  if ((VH_Paie.PGTicketRestau) and ((fournisseur = '003') or (fournisseur = '004'))) then   // pt88
  // ACCOR ou chèque déjeuner pt88
  begin
    if (V_PGI.NumVersionBase < 800) then
    // (spécif V7)
    begin
//      PointLivr := GetField('ETB_AUTRENUMERO');
      PointLivr := GetControlText ('ETB_AUTRENUMERO_');
    end
    else
    // ( à partir V8)
    begin
      PointLivr := GetField('ETB_TICKLIVR');
    end;
    if (fournisseur = '003') then
    begin  // pt88
    // si ACCOR
      if ((not IsNumeric(PointLivr)) or (length(PointLivr) <> 6)) then
      begin
        LastError := 1;
        PgiBox ('Le code point de livraison doit être numérique et#13#10'+
                'd''une longueur de 6 caractères ', TFFiche(Ecran).Caption);
        if (V_PGI.NumVersionBase < 800) then
        // (spécif V7)
        begin
          SetFocusControl('ETB_AUTRENUMERO_')
        end
        else
        // ( à partir V8)
        begin
          SetFocusControl('ETB_TICKLIVR');
        end;
        exit;
     end;
    // deb pt88
    end
    else
    begin
    // si chèque déjeuner
      if (VH_Paie.PGFactEtabl) then
      begin
        if ((not IsNumeric(PointLivr)) or (length(PointLivr) <> 5)) then
        begin
         LastError := 1;
         PgiBox ('Le code client doit être numérique et#13#10'+
                'd''une longueur de 5 caractères ', TFFiche(Ecran).Caption);
         SetFocusControl('ETB_TICKLIVR');
         exit;
        end;
       end;
      end;
    // fin pt88
  end;
// f PT81
  {-----------------------------------------------------------------------
  MAJ DES CHAMPS PROFIL IDEMETABL de la fiche SALARIES }
  //DEBUT PT186 Ajout maj historique
   // modif mv : je cast tout en string
  Etab := GetField('ETB_ETABLISSEMENT');
  If GetParamSocSecur('SO_PGHISTOAVANCE',False) = True then GereHisto := True
  else GereHisto := False;

  //DEB PT91
  if (GereHisto) and (GetControlText('DATEHISTO')='') then
  begin
    LastError := 1;
    PgiBox('Vous devez renseigner la date de prise en compte de la modification de la convention collective', TFFiche(Ecran).Caption);
    SetFocusControl('DATEHISTO');
    exit;
  end;
  //FIN PT91

  { DEB PT32 }
  { DEB PT45 }
  if Activite <> string(GetField('ETB_ACTIVITE')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_ACTIVITE="' + string(GetField('ETB_ACTIVITE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPACTIVITE="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
    If GereHisto then AjoutModifsHisto('PSA_ACTIVITE','PSA_TYPACTIVITE',string(GetField('ETB_ACTIVITE')));
  end;
  { FIN PT32 }
  if Profil <> string(GetField('ETB_PROFIL')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFIL="' + string(GetField('ETB_PROFIL')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFIL="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
      If GereHisto then AjoutModifsHisto('PSA_PROFIL','PSA_TYPPROFIL',string(GetField('ETB_PROFIL')));
    end;
  //  PT8 : 01/10/2001 V562 PH Nouveau champ ETB_PROFILREM
  if ProfilRem <> string(GetField('ETB_PROFILREM')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILREM="' + string(GetField('ETB_PROFILREM')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILREM="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
      If GereHisto then AjoutModifsHisto('PSA_PROFILREM','PSA_TYPPROFILREM',string(GetField('ETB_PROFILREM')));
    end;
  // FIN PT8
  if (PeriodBul <> string(GetField('ETB_PERIODBUL'))) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PERIODBUL="' + string(GetField('ETB_PERIODBUL')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPERIODEBUL="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
      If GereHisto then AjoutModifsHisto('PSA_PERIODBUL','PSA_TYPPERIODEBUL',string(GetField('ETB_PERIODBUL')));
    end;
  if ProfilRbs <> string(GetField('ETB_PROFILRBS')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILRBS="' + string(GetField('ETB_PROFILRBS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILRBS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILRBS','PSA_TYPPROFILRBS',string(GetField('ETB_PROFILRBS')));
  end;
  if RedRepas <> string(GetField('ETB_REDREPAS')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_REDREPAS="' + string(GetField('ETB_REDREPAS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPREDREPAS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
    If GereHisto then AjoutModifsHisto('PSA_REDREPAS','PSA_TYPREDREPAS',string(GetField('ETB_REDREPAS')));
  end;
  if RedRtt1 <> string(GetField('ETB_REDRTT1')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_REDRTT1="' + string(GetField('ETB_REDRTT1')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPREDRTT1="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_REDRTT1','PSA_TYPREDRTT1',string(GetField('ETB_REDRTT1')));
  end;
  if RedRtt2 <> string(GetField('ETB_REDRTT2')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_REDRTT2="' + string(GetField('ETB_REDRTT2')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPREDRTT2="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_REDRTT2','PSA_TYPREDRTT2',string(GetField('ETB_REDRTT2')));
  end;
  if ProfilAfp <> string(GetField('ETB_PROFILAFP')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILAFP="' + string(GetField('ETB_PROFILAFP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILAFP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILAFP','PSA_TYPPROFILAFP',string(GetField('ETB_PROFILAFP')));
  end;
  if GetField('ETB_PCTFRAISPROF') <> null then // beurk
    if (PourcentAfp <> string(GetField('ETB_PCTFRAISPROF'))) then
    begin
      nb := GetField('ETB_PCTFRAISPROF');
      nb := StrFPoint(valeur(nb));
      ExecuteSQL('UPDATE SALARIES SET PSA_PCTFRAISPROF=' + (Nb) + ',PSA_DATEMODIF="' + UsTime(Now) + '" ' + //DB2
        'WHERE PSA_TYPPROFILAFP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
      nb := '';
    end;
  if ProfilApp <> string(GetField('ETB_PROFILAPP')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILAPP="' + string(GetField('ETB_PROFILAPP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILAPP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILAPP','PSA_TYPPROFILAPP',string(GetField('ETB_PROFILAPP')));
   end;
  // PT17   02/04/2002 V571 Ph Ajout champ profil retraite
  if ProfilRet <> string(GetField('ETB_PROFILRET')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILRET="' + string(GetField('ETB_PROFILRET')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILRET="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
      If GereHisto then AjoutModifsHisto('PSA_PROFILRET','PSA_TYPPROFILRET',string(GetField('ETB_PROFILRET')));
  end;
  if ProfilMut <> string(GetField('ETB_PROFILMUT')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILMUT="' + string(GetField('ETB_PROFILMUT')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILMUT="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILMUT','PSA_TYPPROFILMUT',string(GetField('ETB_PROFILMUT')));
  end;
  if ProfilPre <> string(GetField('ETB_PROFILPRE')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILPRE="' + string(GetField('ETB_PROFILPRE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILPRE="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILPRE','PSA_TYPPROFILPRE',string(GetField('ETB_PROFILPRE')));
  end;
  if ProfilTss <> string(GetField('ETB_PROFILTSS')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILTSS="' + string(GetField('ETB_PROFILTSS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILTSS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILTSS','PSA_TYPPROFILTSS',string(GetField('ETB_PROFILTSS')));
  end;
  if ProfilCge <> string(GetField('ETB_PROFILCGE')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILCGE="' + string(GetField('ETB_PROFILCGE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILCGE="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILCGE','PSA_TYPPROFILCGE',string(GetField('ETB_PROFILCGE')));
  end;
  if ProfilAncien <> string(GetField('ETB_PROFILANC')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILANCIEN="' + string(GetField('ETB_PROFILANCIEN')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILANC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILANCIEN','PSA_TYPPROFILANC',string(GetField('ETB_PROFILANC')));
  end;
  if ProfilTrans <> string(GetField('ETB_PROFILTRANS')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILTRANS="' + string(GetField('ETB_PROFILTRANS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPROFILTRANS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_PROFILTRANS','PSA_TYPPROFILTRANS',string(GetField('ETB_PROFILTRANS')));
  end;
  //CONGES PAYES
  if StandCalend <> string(GetField('ETB_STANDCALEND')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_CALENDRIER="' + string(GetField('ETB_STANDCALEND')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_STANDCALEND="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_CALENDRIER','PSA_STANDCALEND',string(GetField('ETB_STANDCALEND')));
  end;
  if JourHeure <> string(GetField('ETB_JOURHEURE')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_JOURHEURE="' + string(GetField('ETB_JOURHEURE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPJOURHEURE="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if GetField('ETB_NBREACQUISCP') <> null then // beurk
    if NbreAcquisCp <> string(GetField('ETB_NBREACQUISCP')) then
    begin
      nb := GetField('ETB_NBREACQUISCP');
      nb := StrFPoint(valeur(nb));
      ExecuteSQL('UPDATE SALARIES SET PSA_NBREACQUISCP=' + nb + ',PSA_DATEMODIF="' + UsTime(Now) + '" ' + //DB2
        'WHERE PSA_CPACQUISMOIS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    end;
  if NbAcquisCp <> string(GetField('ETB_NBACQUISCP')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_NBACQUISCP="' + string(GetField('ETB_NBACQUISCP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPNBACQUISCP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  { DEB PT52-2 }
  if NbCpSupp <> Valeur(GetField('ETB_NBRECPSUPP')) then
  begin
    nb := GetField('ETB_NBRECPSUPP');
    nb := StrFPoint(valeur(nb));
    ExecuteSQL('UPDATE SALARIES SET PSA_NBRECPSUPP=' + nb + ',PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_CPACQUISSUPP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  end; { FIN PT52-2 }
  if ValorIndempCp <> string(GetField('ETB_VALORINDEMCP')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_VALORINDEMCP="' + string(GetField('ETB_VALORINDEMCP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_CPTYPEMETHOD="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if basancp <> string(GetField('ETB_BASANCCP')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_BASANCCP="' + string(GetField('ETB_BASANCCP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_CPACQUISANC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if ValancCp <> string(GetField('ETB_VALANCCP')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_VALANCCP="' + string(GetField('ETB_VALANCCP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_CPACQUISANC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if DateAcqCpAnc <> string(GetField('ETB_DATEACQCPANC')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_DATEACQCPANC="' + string(GetField('ETB_DATEACQCPANC')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_DATANC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if ValtypDatAnc <> string(GetField('ETB_TYPDATANC')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_TYPDATANC="' + string(GetField('ETB_TYPDATANC')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_DATANC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  //if DatAnc<>string(GetField('ETB_DATANC')) then
  //   ExecuteSQL('UPDATE SALARIES SET PSA_DATANC="'+string(GetField('ETB_DATANC'))+'" AND PSA_ETABLISSEMENT="'+Etab+'"') ;
  if Reliquat <> string(GetField('ETB_RELIQUAT')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_RELIQUAT="' + string(GetField('ETB_RELIQUAT')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_CPTYPERELIQ="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');

  { DEB PT70  }
  if PaieValoMs <> string(GetField('ETB_PAIEVALOMS')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_PAIEVALOMS="' + string(GetField('ETB_PAIEVALOMS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPAIEVALOMS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  { FIN PT70  }


  { DEB PT46 }
  if GblEditBulCP <> string(GetField('ETB_EDITBULCP')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_EDITBULCP="' + string(GetField('ETB_EDITBULCP')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPEDITBULCP="ETB" AND PSA_ETABLISSEMENT="' + Etab + '" ');
  { FIN PT46 }
  //REGLEMENT
  if ModeReglement <> string(GetField('ETB_PGMODEREGLE')) then {PT3}
    ExecuteSQL('UPDATE SALARIES SET PSA_PGMODEREGLE="' + string(GetField('ETB_PGMODEREGLE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPREGLT="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if RibSalaire <> string(GetField('ETB_RIBSALAIRE')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_RIBVIRSOC="' + string(GetField('ETB_RIBSALAIRE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPVIRSOC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if PaiAcompte <> string(GetField('ETB_PAIACOMPTE')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_PAIACOMPTE="' + string(GetField('ETB_PAIACOMPTE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPAIACOMPT="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if RibAcpSoc <> string(GetField('ETB_RIBACOMPTE')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_RIBACPSOC="' + string(GetField('ETB_RIBACOMPTE')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPACPSOC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if PaiFrais <> string(GetField('ETB_PAIFRAIS')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_PAIFRAIS="' + string(GetField('ETB_PAIFRAIS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPPAIFRAIS="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if RibFraisSoc <> string(GetField('ETB_RIBFRAIS')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_RIBFRAISSOC="' + string(GetField('ETB_RIBFRAIS')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPFRAISSOC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if MoisPaiement <> string(GetField('ETB_MOISPAIEMENT')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_MOISPAIEMENT="' + string(GetField('ETB_MOISPAIEMENT')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPDATPAIEMENT="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  if GetField('ETB_JOURPAIEMENT') <> null then // beurk
    if JourPaiement <> string(GetField('ETB_JOURPAIEMENT')) then
    begin
      inb := GetField('ETB_JOURPAIEMENT');
      ExecuteSQL('UPDATE SALARIES SET PSA_JOURPAIEMENT=' + InttoStr(inb) + ',PSA_DATEMODIF="' + UsTime(Now) + '" ' + //DB2
        'WHERE PSA_TYPDATPAIEMENT="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    end;

  //DADS-U
  //PT14
  if TypeFraction <> string(GetField('ETB_CODESECTION')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_DADSFRACTION="' + string(GetField('ETB_CODESECTION')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPDADSFRAC="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  //PT24-1
  if EditOrg <> string(GetField('ETB_EDITORG')) then
    ExecuteSQL('UPDATE SALARIES SET PSA_EDITORG="' + string(GetField('ETB_EDITORG')) + '",PSA_DATEMODIF="' + UsTime(Now) + '" ' +
      'WHERE PSA_TYPEDITORG="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
  { FIN PT45 }
  // d PT33
  if TicketRestau <> string(GetField('ETB_TYPTICKET')) then
    ExecuteSQL('UPDATE DEPORTSAL SET PSE_TYPTICKET="' +
      string(GetField('ETB_TYPTICKET')) +
      '" WHERE PSE_TYPETYPTK="ETB" AND ' +
      'PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES ' +
      ' WHERE PSA_ETABLISSEMENT="' + Etab + '")');


  // f PT33

  //DEB PT50
  if IsFieldModified('ETB_MSAACTIVITE') then
    ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSAACTIVITE="' +
      (GetField('ETB_MSAACTIVITE')) +
      '" WHERE PSE_MSATYPEACT="ETB" AND ' +
      'PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES ' +
      ' WHERE PSA_ETABLISSEMENT="' + Etab + '")');

  //FIN PT50
  //DEBUT PT71
  if ConventionColl <> string(GetField('ETB_CONVENTION')) then
  begin
    ExecuteSQL('UPDATE SALARIES SET PSA_CONVENTION="' + string(GetField('ETB_CONVENTION')) + '" ' +
      'WHERE PSA_TYPCONVENTION="ETB" AND PSA_ETABLISSEMENT="' + Etab + '"');
    If GereHisto then AjoutModifsHisto('PSA_CONVENTION','PSA_TYPCONVENTION',string(GetField('ETB_CONVENTION')));
  end;
  //FIN PT71
  //DEBUT PT72
  if UGMsa <> string(GetField('ETB_MSAUNITEGES')) then
    ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSAUNITEGES="' + string(GetField('ETB_MSAUNITEGES')) + '" ' +
      'WHERE PSE_MSATYPUNITEG="ETB" AND  ' +
      'PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES ' +
      ' WHERE PSA_ETABLISSEMENT="' + Etab + '")');
  //FIN PT72

// d PT81
 //  if ((VH_Paie.PGTicketRestau) and (fournisseur = '003')) then    pt88
 if ((VH_Paie.PGTicketRestau) and ((fournisseur = '003') or (fournisseur = '004'))) then     // pt88
  // ACCOR ou chéque déjeuner pt88
  begin
    if (V_PGI.NumversionBase < 800) then
    // (spécif V7)
    begin
      if PointLivraison <> string(GetField('ETB_AUTRENUMERO')) then
        ExecuteSQL('UPDATE DEPORTSAL SET PSE_MSAUNITEGES="' +
                    string(GetField('ETB_AUTRENUMERO')) +
                    '" WHERE PSE_MSATYPUNITEG = "ETB" AND ' +
                    'PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES ' +
                   ' WHERE PSA_ETABLISSEMENT="' + Etab + '")');
    end
    else
    // (à partir V8)
    begin
      if PointLivraison <> string(GetField('ETB_TICKLIVR')) then
        ExecuteSQL('UPDATE DEPORTSAL SET PSE_TICKLIVR = "' +
                   string(GetField('ETB_TICKLIVR')) +
                   '" WHERE PSE_TYPTICKLIVR = "ETB" AND ' +
                   'PSE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES ' +
                   ' WHERE PSA_ETABLISSEMENT="' + Etab + '")');
    end;
  end;
// f PT81
//PT87
if (College<>string (GetField ('ETB_PRUDHCOLL'))) then
   ExecuteSQL ('UPDATE SALARIES SET'+
               ' PSA_PRUDHCOLL="'+string (GetField ('ETB_PRUDHCOLL'))+'" WHERE'+
               ' PSA_TYPPRUDH="ETB" AND'+
               ' PSA_ETABLISSEMENT="'+Etab+'"');

if (Section<>string (GetField ('ETB_PRUDHSECT'))) then
   ExecuteSQL ('UPDATE SALARIES SET'+
               ' PSA_PRUDHSECT="'+string (GetField ('ETB_PRUDHSECT'))+'" WHERE'+
               ' PSA_TYPPRUDH="ETB" AND'+
               ' PSA_ETABLISSEMENT="'+Etab+'"');

if (LieuVote<>string (GetField ('ETB_PRUDHVOTE'))) then
   ExecuteSQL ('UPDATE SALARIES SET'+
               ' PSA_PRUDHVOTE="'+string (GetField ('ETB_PRUDHVOTE'))+'" WHERE'+
               ' PSA_TYPPRUDH="ETB" AND'+
               ' PSA_ETABLISSEMENT="'+Etab+'"');
//FIN PT87

  ChargeVarGblIdemEtab; //PT46
  if LastError = 0 then ForceValidCp := False; { PT55-1 }
end;
{$ENDIF}
//FIN PT62

//
//                       PARTIE ONCHANGE
//
//
//

procedure TOM_EtabCompl.OnChangeField(F: TField);
var
{$IFNDEF EAGLCLIENT}
MEdit: THDBEdit;
i: Integer;
{$ELSE}
MEdit: THEdit;
{$ENDIF}
aa, mm, jj: Word;
Ladate: Tdatetime;
jour, mois, champ, st, StPlus: string;
Zone: TCOntrol;
Q: TQuery;
// PT84 PointLivr: string; // PT81

begin
inherited;

if (F.Fieldname = 'ETB_CONGESPAYES') then
   begin { DEB PT70 }
   if (GetField('ETB_CONGESPAYES') = 'X') then
      Begin
      DecocheCp:= False;
      If GetField('ETB_VALORINDEMCP')='' then
         SetField('ETB_VALORINDEMCP','X');
      If GetField('ETB_RELIQUAT')='' then
         SetField('ETB_RELIQUAT','0');
      End;             { FIN PT70 }
   CongesPayesClick(nil);
   end;

if (F.Fieldname = 'ETB_PROFILCGE') and (Getfield('ETB_CONGESPAYES') = 'X') then
   begin
   Zone:= getcontrol('ETB_PROFILCGE');
   InitialiseCombo(Zone);
   end;

if (F.FieldName = 'ETB_JEDTDU') or (F.FieldName = 'ETB_JEDTAU') then
   begin
   if GetField('ETB_JEDTDU') < GetField('ETB_JEDTAU') then
      PgiBox ('Le jour édité du doit être supérieur au jour édité au afin'+
              ' d''avoir une cohérence des dates d''édition du bulletin',
              Ecran.caption);
   if ((GetField('ETB_JEDTDU'))-(GetField('ETB_JEDTAU')) <> 1) and
      (GetField('ETB_JEDTAU') > 0) then
      PgiBox ('Attention, vos dates d''édition ne seront pas cohérentes car'+
              ' elles ne correspondront pas à un mois complet.#13#10'+
              'Reportez vous à l''aide en ligne', Ecran.caption);
   end;

// Nombre de jours de travail par semaine
if (F.FieldName = 'ETB_NBJOUTRAV') and (Getfield('ETB_CONGESPAYES') = 'X') then
// Controle format correct
   begin
   SetControlEnabled('ETB_2EMREPOSH', GetField('ETB_NBJOUTRAV') = '5');
   SetControlEnabled('L2EME', GetField('ETB_NBJOUTRAV') = '5');
   if GetField('ETB_NBJOUTRAV') <> GblNbJoutrav then
      begin
      if GetField('ETB_NBJOUTRAV') = '6' then
         begin
         SetField('ETB_2EMEREPOSH', '');
         SetField('ETB_1ERREPOSH', '1DI');
         Setfield('ETB_NBREACQUISCP', 2.5);
         if (getfield('ETB_MVALOMS') = 'T') then
            Setfield('ETB_VALODXMN', 26);
         end
      else
         if GetField('ETB_NBJOUTRAV') = '5' then
            begin
            SetField('ETB_2EMEREPOSH', '1DI');
            SetField('ETB_1ERREPOSH', '7SA');
            Setfield('ETB_NBREACQUISCP', 2.08);
            if (getfield('ETB_MVALOMS') = 'T') then
               Setfield('ETB_VALODXMN', 21.67);
            end;
      end;
   GblNbJoutrav := GetField('ETB_NBJOUTRAV');
   end;

//Nombre de jours CP acquis par mois
(* 25/07/2006 Mise en commentaire, incompatibilité Delphi7 et contrôle inutile 
if (F.Fieldname='ETB_NBREACQUISCP') and (Getfield('ETB_CONGESPAYES')='X') then
   begin
// controle format correct
{$IFNDEF EAGLCLIENT}
   MEdit:= THDBEdit(Getcontrol('ETB_NBREACQUISCP'));
{$ELSE}
   MEdit:= THEdit(Getcontrol('ETB_NBREACQUISCP'));
{$ENDIF}
   if MEdit <> nil then
      begin
{$IFNDEF EAGLCLIENT}
      for i := 1 to length(MEdit.Text) do
          begin
          if not F.IsValidChar(MEdit.Text[i]) then   
             begin
             ShowMessage('Format incorrect');
             SetFocusControl('ETB_NBREACQUISCP');
             break;
             end;
          end; // for
{$ENDIF}
      end;
   end;   *)

{Nombre de jours CP acquis supplémentaire }
if (F.FieldName = 'ETB_NBRECPSUPP') then
   begin
   if (HEnt1.Valeur(GetField('ETB_NBRECPSUPP')) < 0) then
      begin
      PgiBox ('Le nombre de jours acquis supplémentaire doit être positif.',
              Ecran.caption);
      if Hent1.Valeur(GetField('ETB_NBRECPSUPP')) <> 0 then
         SetField('ETB_NBRECPSUPP', '0');
      end;
   end;

// Base congés supplémentaire ancienneté
if (F.FieldName = 'ETB_BASANCCP') and (Getfield('ETB_CONGESPAYES') = 'X') then
   begin
   if (GetField('ETB_BASANCCP') = '') then
      begin
      if GetField('ETB_VALANCCP') <> '' then
         setfield('ETB_VALANCCP', '');
      if GetField('ETB_TYPDATANC') <> '' then
         setfield('ETB_TYPDATANC', '');
      if GetField('ETB_DATEACQCPANC') <> '' then
         setfield('ETB_DATEACQCPANC', '');
      end;
   if (GblBasAnccp <> GetField('ETB_BASANCCP')) then
      if GetField('ETB_VALANCCP') <> '' then
         setfield('ETB_VALANCCP', '');
   GblBasAnccp:= GetField('ETB_BASANCCP');
   CongesPayesRechercheTabletteAnciennete;
   Zone:= getcontrol('ETB_BASANCCP');
   InitialiseCombo(Zone);
   end;

if (F.FieldName = 'ETB_VALODXMN') and (Getfield('ETB_CONGESPAYES') = 'X') then
   begin
   if ((GetField('ETB_VALODXMN')>=100) and (GetField('ETB_MVALOMS')='M')) then
      begin
      PGIBox('La valeur saisie est incohérente', 'Jour maintien');
      SetField('ETB_VALODXMN', 0);
      SetFocuscontrol('ETB_VALODXMN');
      end;
    end;

// type date ancienneté : saisie date ou date entrée salarié
if (F.FieldName = 'ETB_TYPDATANC') and (Getfield('ETB_CONGESPAYES') = 'X') then
   begin
   if Getfield('ETB_TYPDATANC') <> GblTypDatAnc then
      begin
      if (Getfield('ETB_TYPDATANC') = '1') OR (Getfield('ETB_TYPDATANC') = '3') then  { PT82-1 }
         setfield('ETB_DATEACQCPANC', '')
      else
         if (Getfield('ETB_TYPDATANC') = '0') OR (Getfield('ETB_TYPDATANC') = '4') then  { PT82-1 }
            begin
            Ladate:= getfield('ETB_DATECLOTURECPN');
            if Ladate > 0 then
               begin
               if   (Getfield('ETB_TYPDATANC') = '0') then  { PT82-1 }
               decodedate(Ladate + 1, aa, mm, jj)
               else
               decodedate(Ladate, aa, mm, jj);   { PT82-1 }
               if jj > 9 then
                  jour:= inttostr(jj)
               else
                  jour:= '0'+inttostr(jj);
               if mm > 9 then
                  mois:= inttostr(mm)
               else
                  mois:= '0'+inttostr(mm);
               setfield('ETB_DATEACQCPANC', jour + mois);
               end;
            end
         else if (Getfield('ETB_TYPDATANC') = '2') then //PT94
              setfield('ETB_DATEACQCPANC', '') //PT94
      end;
   SetcontrolEnabled ('ETB_DATEACQCPANC', ((getfield('ETB_TYPDATANC')='2') and
                      (getfield('ETB_CONGESPAYES') = 'X')));
   GblTypDatAnc:= GetField('ETB_TYPDATANC');
   end;

if (F.fieldName='ETB_VALORINDEMCP') and (Getfield('ETB_CONGESPAYES')='X') then
   begin
   SetcontrolEnabled('ETB_VALODXMN', (getfield('ETB_VALORINDEMCP') <> 'D'));
   SetcontrolEnabled('LETB_VALODXMN', (getfield('ETB_VALORINDEMCP') <> 'D'));
// debutPT95
{   if (getfield('ETB_VALORINDEMCP') = 'D') then
      begin
      if (Ds.state in [DsEdit]) then
        begin
        Setfield('ETB_MVALOMS', '');
        Setfield('ETB_VALODXMN', 0);
        end;
      end
   else }
//fin PT95
      if (getfield('ETB_VALORINDEMCP')='M') or
         (getfield('ETB_VALORINDEMCP')='X') then
         if getfield('ETB_MVALOMS') = '' then
            Setfield('ETB_MVALOMS', 'T');
   if GetField('ETB_PAIEVALOMS')='' then
      SetField('ETB_PAIEVALOMS','ANT');  { PT70 }
   end;

if (F.fieldName = 'ETB_MVALOMS') and (Getfield('ETB_CONGESPAYES') = 'X') then
   begin
   SetcontrolEnabled('ETB_VALODXMN', (getfield('ETB_MVALOMS') = 'M'));
   SetcontrolEnabled('LETB_VALODXMN', (getfield('ETB_MVALOMS') = 'M'));
   if (getfield('ETB_MVALOMS') = 'R') then
      Setfield('ETB_VALODXMN', 0);
   if (getfield('ETB_MVALOMS') = 'T') then
      begin
      if getfield('ETB_NBJOUTRAV') = '5' then
         Setfield('ETB_VALODXMN', 21.67)
      else
         Setfield('ETB_VALODXMN', 26);
      end;
   end;

if F.fieldName = 'ETB_TYPSMIC' then
   begin
   Zone:= GetControl('ETB_TYPSMIC');
   InitialiseCombo(zone);
   champ:= GetField('ETB_TYPSMIC');
   SetControlEnabled('ETB_SMIC', TRUE);
   if (DS.State in [DSEdit]) and (GetField('ETB_SMIC') <> '') then
      SetField('ETB_SMIC', '');
   if champ = 'ELN' then
      SetControlProperty('ETB_SMIC', 'DataType', 'PGELEMENTNAT')
   else
   if GetField('ETB_TYPSMIC') = 'VAR' then
      SetControlProperty('ETB_SMIC', 'DataType', 'PGVARIABLE')
   else
      SetControlEnabled('ETB_SMIC', FALSE);
   end;

if F.fieldName = 'ETB_SMIC' then
   begin
   if GetField('ETB_TYPSMIC') = 'ELN' then
      SetControlText ('TLIB_SMIC',
                      RechDom('PGELEMENTNAT', GetField('ETB_SMIC'), FALSE))
   else
   if GetField('ETB_TYPSMIC') = 'VAR' then
      SetControlText ('TLIB_SMIC',
                      RechDom('PGVARIABLE', GetField('ETB_SMIC'), FALSE))
   else
      SetControlText('TLIB_SMIC', '');
   end;

// Onglet DADS-U
if (F.FieldName = ('ETB_CODESECTION')) then
   begin
   if ((GetField('ETB_CODESECTION') <> '') and
      (GetField('ETB_CODESECTION') > Getparamsoc('SO_PGFRACTION'))) then
      begin
      PGIBox ('La valeur saisie est supérieure à '+Getparamsoc('SO_PGFRACTION')+
              '#13#10qui est la valeur maximale des paramètres société',
              'Fraction DADS');
      SetField('ETB_CODESECTION', '1');
      end;
   end;

//Onglet Reglement
if (F.FieldName = 'ETB_RIBSALAIRE') or (F.FieldName = 'ETB_RIBFRAIS') or
   (F.FieldName = 'ETB_RIBACOMPTE') or (F.FieldName = 'ETB_RIBDUCSEDI') then
   begin
// PT79   if GetField(F.FieldName) <> '' then
   if (GetField(F.FieldName) <> '') and (GetField(F.FieldName) <> ' ') then
      begin
      StPlus:= PGBanqueCP (True);   //PT85-1
      Q:= OpenSql ('SELECT BQ_ETABBQ,BQ_GUICHET,BQ_NUMEROCOMPTE,BQ_CLERIB'+
                   ' FROM BANQUECP WHERE'+
                   ' BQ_GENERAL="'+GetField(F.FieldName)+'"'+StPlus, True);
      if not Q.eof then
         begin
         St:= 'Compte : '+Q.Findfield('BQ_ETABBQ').AsString;
         St:= St+'  '+Q.Findfield('BQ_GUICHET').AsString;
         St:= St+'  '+Q.Findfield('BQ_NUMEROCOMPTE').AsString;
         St:= St+'  '+Q.Findfield('BQ_CLERIB').AsString;
         SetControlText('T' + Copy(F.FieldName, 5, Length(F.FieldName)), St);
         end;
      Ferme(Q);
      end
   else
      SetControlText('T'+Copy(F.FieldName, 5, Length(F.FieldName)), '');
   end;

// Onglet Caractéristiques
if (F.FieldName = ('ETB_HORAIREETABL')) then
   begin
   if ((GetField('ETB_HORAIREETABL') <> 0) and
      (GetField('ETB_HORAIREETABL') > 300)) then
      begin
      PGIBox ('La valeur saisie est incohérente',
              'Horaire de référence établissement');
      SetField('ETB_HORAIREETABL', 0);
      SetFocuscontrol('ETB_HORAIREETABL');
      end;
   end;

if (F.FieldName = ('ETB_PRORATATVA')) then
   begin
   if ((GetField('ETB_PRORATATVA') <> 0) and
      (GetField('ETB_PRORATATVA') > 100)) then
      begin
      PGIBox('La valeur saisie est incohérente', '% Prorata TVA');
      SetField('ETB_PRORATATVA', 0);
      SetFocuscontrol('ETB_PRORATATVA');
      end;
   end;

//Onglet Profils
if (F.FieldName = ('ETB_TAUXVERSTRANS')) then
   begin
   if ((GetField('ETB_TAUXVERSTRANS') <> 0) and
      (GetField('ETB_TAUXVERSTRANS') >= 100)) then
      begin
      PGIBox('La valeur saisie est incohérente', 'Taux de versement');
      SetField('ETB_TAUXVERSTRANS', 0);
      SetFocuscontrol('ETB_TAUXVERSTRANS');
      end;
   end;

if (F.FieldName = ('ETB_MOISPAIEMENT')) then
   begin
   if getField('ETB_MOISPAIEMENT') = 'FDM' then
      SetControlEnabled('ETB_JOURPAIEMENT', False)
   else
      SetControlEnabled('ETB_JOURPAIEMENT', True);
   end;

if (F.FieldName = ('ETB_BCMOISPAIEMENT')) then
   begin
   if getField('ETB_BCMOISPAIEMENT') = 'FDM' then
      SetControlEnabled('ETB_BCJOURPAIEMENT', False)
   else
      SetControlEnabled('ETB_BCJOURPAIEMENT', True);
   end;



//PT75
if (F.FieldName = ('ETB_PRUDH')) then
   begin
   st:= 'SELECT ET_APE'+
        ' FROM ETABLISS'+
        ' LEFT JOIN ETABCOMPL ON'+
        ' ET_ETABLISSEMENT=ETB_ETABLISSEMENT WHERE'+
        ' ET_ETABLISSEMENT="'+Getfield ('ETB_ETABLISSEMENT')+'"';
   Q:= OpenSql(st, TRUE);
   if not Q.EOF then
      begin
      if ((((Q.Fields[0].Asstring='725Z') or (Q.Fields[0].Asstring='748B') or
         (Q.Fields[0].Asstring='921G') or (Q.Fields[0].Asstring='924Z') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)='05') and
         (Q.Fields[0].Asstring<>'050C')) or        //PT78
         ((Copy (Q.Fields[0].Asstring, 1, 2)>='10') and
         (Copy (Q.Fields[0].Asstring, 1, 2)<='36') and
         (Q.Fields[0].Asstring<>'151F')) or        //PT78
         (Copy (Q.Fields[0].Asstring, 1, 2)='40') or
         (Copy (Q.Fields[0].Asstring, 1, 2)='41') or
         (Copy (Q.Fields[0].Asstring, 1, 2)='45')) and
         (GetField ('ETB_PRUDH')<>'01')) or
         (((Q.Fields[0].Asstring='151F') or (Q.Fields[0].Asstring='741J') or
         (Q.Fields[0].Asstring='747Z') or (Q.Fields[0].Asstring='748A') or
         (Q.Fields[0].Asstring='748G') or (Q.Fields[0].Asstring='748H') or
         (Q.Fields[0].Asstring='851H') or (Q.Fields[0].Asstring='922F') or
         (Copy (Q.Fields[0].Asstring, 1, 2)='37') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)>='50') and
         (Copy (Q.Fields[0].Asstring, 1, 2)<='52')) or
         (Copy (Q.Fields[0].Asstring, 1, 2)='55') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)>='60') and
         (Copy (Q.Fields[0].Asstring, 1, 2)<='71') and
         (Q.Fields[0].Asstring<>'602C') and (Q.Fields[0].Asstring<>'660G') and
         (Q.Fields[0].Asstring<>'701C')) or //PT78
         (Copy (Q.Fields[0].Asstring, 1, 2)='90') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)='93') and
         (Q.Fields[0].Asstring<>'930K'))) and      //PT78
         (GetField ('ETB_PRUDH')<>'02')) or
         (((Q.Fields[0].Asstring='050C') or
         (Copy (Q.Fields[0].Asstring, 1, 2)='01') or
         (Copy (Q.Fields[0].Asstring, 1, 2)='02')) and
         (GetField ('ETB_PRUDH')<>'03')) or
         (((Q.Fields[0].Asstring='602C') or (Q.Fields[0].Asstring='660G') or
         (Q.Fields[0].Asstring='701C') or (Q.Fields[0].Asstring='930K') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)>='72') and
         (Copy (Q.Fields[0].Asstring, 1, 2)<='75') and
         (Q.Fields[0].Asstring<>'725Z') and (Q.Fields[0].Asstring<>'741J') and
         (Q.Fields[0].Asstring<>'747Z') and (Q.Fields[0].Asstring<>'748A') and
         (Q.Fields[0].Asstring<>'748B') and (Q.Fields[0].Asstring<>'748G') and
         (Q.Fields[0].Asstring<>'748H')) or //PT78
         (Copy (Q.Fields[0].Asstring, 1, 2)='80') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)='85') and
         (Q.Fields[0].Asstring<>'851H')) or        //PT78
         (Copy (Q.Fields[0].Asstring, 1, 2)='91') or
         ((Copy (Q.Fields[0].Asstring, 1, 2)='92') and
         (Q.Fields[0].Asstring<>'921G') and (Q.Fields[0].Asstring<>'922F') and
         (Q.Fields[0].Asstring<>'924Z')) or //PT78
         ((Copy (Q.Fields[0].Asstring, 1, 2)>='95') and
         (Copy (Q.Fields[0].Asstring, 1, 2)<='97')) or
         (Copy (Q.Fields[0].Asstring, 1, 2)='99')) and
         (GetField ('ETB_PRUDH')<>'04'))) then
         begin
         PgiBox ('La section prud''hommale établissement n''est pas#13#10'+
                 'compatible avec le code APE', TFFiche(Ecran).Caption);
         SetFocusControl('ETB_PRUDH');
         end;
      end;
   Ferme(Q);
   end;
//FIN PT75

  //DEB PT91
  if (F.FieldName = ('ETB_CONVENTION')) then
  begin
    if (ConventionColl <> string(GetField('ETB_CONVENTION'))) and (GetParamSocSecur('SO_PGHISTOAVANCE',False) = True) then
      SetControlVisible('DATEHISTO',True)
    else
      SetControlVisible('DATEHISTO',False);
  end;
  //FIN PT91

// d PT81
{* PT84
// d PT ds.State n'est jamais en dsinsert ==> on ne contrôle le champ que sur OnUpdateRecord
  if (((F.FieldName = ('ETB_AUTRENUMERO')) or (F.FieldName = ('ETB_TICKLIVR'))) and      // PT81
    (VH_Paie.PGTicketRestau) and (fournisseur = '003')) then
  // Titres restaurant : ACCOR
  begin
    if (V_PGI.NumVersionBase < 800) then
    // ( spécif V7)
    begin
      PointLivr := GetField('ETB_AUTRENUMERO');
    end
    else
    // (à partir V8)
    begin
      PointLivr := GetField('ETB_TICKLIVR');
    end;
    if ((not IsNumeric(PointLivr)) or (length(PointLivr) <> 6)) then
    begin
      PgiBox ('Le code point de livraison doit être numérique et#13#10'+
              'd''une longueur de 6 caractères ', TFFiche(Ecran).Caption);
      if (V_PGI.NumVersionBase < 800) then
      // (spécif V7)
      begin
        SetFocusControl('ETB_AUTRENUMERO_');
      end
      else
      // (à partir V8)
      begin
        SetFocusControl('ETB_TICKLIVR');
      end;
    end;
  end;
*}
// f PT81


//PT83 !!! Toujours en dernier car en WebAccess efface la valeur du F.FieldName
if (Getfield('ETB_CONGESPAYES') = 'X') then EnableChampsOngletCongesPayes;
end;

procedure TOM_EtabCompl.PagesChange(Sender: TObject);
var
  P: TPageControl;
  T: TTabSheet;
begin
  P := TPageControl(GetControl('PAGES'));
  if P <> nil then T := P.ActivePage
  else T := nil;
  // on dirige le controle  vesr la procedure liée à l'onglet courant.
  if T <> nil then
  begin
    SetControlVisible('TBCLOTCP', T.Name = 'PCONGES');
    SetControlVisible('TBDECLOTCP', False); { PT49 }
  end;
end;

//PT62
{$IFNDEF DADSB}
//------------------------------------------------------------------------------
//- détermine, pour l'onglet Congés Payés, la visibilité et les options de saisie pour
//- champs de cet onglet
//------------------------------------------------------------------------------

procedure TOM_EtabCompl.ControleZonesOngletCongesPayes;
var
{$IFNDEF EAGLCLIENT}
  MEdit: THDBEdit;
  MCombo, MCombo1: ThdbValcombobox;
  MCheckBox: TDBCheckBox;
{$ELSE}
  MEdit: THEdit;
  MCombo, MCombo1: ThValcombobox;
  MCheckBox: TCheckBox;
{$ENDIF}
  Ouvrable: boolean;
begin
{$IFNDEF EAGLCLIENT}
  MCheckBox := TDBCheckBox(getcontrol('ETB_CONGESPAYES'));
{$ELSE}
  MCheckBox := TCheckBox(getcontrol('ETB_CONGESPAYES'));
{$ENDIF}
  if MCheckbox = nil then exit;
  if MCheckBox.Checked then
  begin
    // Date de cloture cp saisie
    if not IsValidDate(string(getfield('ETB_DATECLOTURECPN'))) then
    begin
      LastError := 125;
      PgiBox('Veuillez saisir une date de clôture valide.', Ecran.caption); { PT55-2 }
      PositionneFocusOngletCP;
      SetFocusControl('ETB_DATECLOTURECPN');
      exit;
    end;

    // ETB_NBJOUTRAV : Controle zone égale à 5 ou 6
    { DEB PT40 }
    if ((Valeur(GetField('ETB_NBJOUTRAV')) > 6) or (Valeur(GetField('ETB_NBJOUTRAV')) < 5)) then
    begin
      LastError := 102;
      PgiBox('Veuillez saisir la valeur 5 (ouvrés) ou (6) ouvrables.', Ecran.caption); { PT55-2 }
      PositionneFocusOngletCP;
      SetFocusControl('ETB_NBJOUTRAV');
      exit;
    end;
    Ouvrable := (GetField('ETB_NBJOUTRAV') = '5');
    { FIN PT40 }
    // ETB_1ERREPOSH et ETB_2EMEREPOSH
    // Controle de la saisie des jours de repos hebdomadaires
{$IFNDEF EAGLCLIENT}
    Mcombo := THDBValComboBox(Getcontrol('ETB_1ERREPOSH'));
{$ELSE}
    Mcombo := THValComboBox(Getcontrol('ETB_1ERREPOSH'));
{$ENDIF}
    if Mcombo <> nil then
    begin
      if MCombo.value = '' then
      begin
        LastError := 105;
        PgiBox('Veuillez renseigner le 1er jour de repos hebdomadaire.', Ecran.caption); { PT55-2 }
        PositionneFocusOngletCP;
        SetFocusControl('ETB_1ERREPOSH');
        exit;
      end;
    end;
    if ouvrable then
    begin
{$IFNDEF EAGLCLIENT}
      Mcombo1 := THDBValComboBox(Getcontrol('ETB_2EMEREPOSH'));
{$ELSE}
      Mcombo1 := THValComboBox(Getcontrol('ETB_2EMEREPOSH'));
{$ENDIF}
      if MCombo1 <> nil then
      begin
        if MCombo1.value = '' then
        begin
          LastError := 106;
          PgiBox('Veuillez renseigner le 2ème jour de repos hebdomadaire.', Ecran.caption); { PT55-2 }
          PositionneFocusOngletCP;
          SetFocusControl('ETB_2EMEREPOSH');
          exit;
        end;
      end;
      if ((Mcombo <> nil) and (Mcombo1 <> nil)) then
      begin
        if MCombo.value = Mcombo1.value then
        begin
          LastError := 107;
          PgiBox('Veuillez saisir le 2ème jour de repos hebdomadaire différent du premier.', Ecran.Caption); { PT55-2 }
          PositionneFocusOngletCP;
          SetFocusControl('ETB_2EMEREPOSH');
          exit;
        end;
      end;
    end;
{$IFNDEF EAGLCLIENT}
    Mcombo := THDBValComboBox(Getcontrol('ETB_RELIQUAT'));
{$ELSE}
    Mcombo := THValComboBox(Getcontrol('ETB_RELIQUAT'));
{$ENDIF}
    if Mcombo <> nil then
    begin
      if MCombo.value = '' then
      begin
        LastError := 120;
        PgiBox('Congés payés : Veuillez renseigner la méthode de gestion de reliquat.', Ecran.caption); { PT55-2  et PT66 }
        PositionneFocusOngletCP;
        SetFocusControl('ETB_RELIQUAT');
        exit;
      end;
    end;
{$IFNDEF EAGLCLIENT}
    Mcombo := THDBValComboBox(Getcontrol('ETB_VALORINDEMCP'));
{$ELSE}
    Mcombo := THValComboBox(Getcontrol('ETB_VALORINDEMCP'));
{$ENDIF}
    if Mcombo <> nil then
    begin
      if MCombo.value = '' then
      begin
        LastError := 125;
        PgiBox('Veuillez saisir la méthode de valorisation des indemnité CP.', Ecran.caption); { PT55-2 }
        PositionneFocusOngletCP;
        SetFocusControl('ETB_VALORINDEMCP');
        exit;
      end;
    end;
{$IFNDEF EAGLCLIENT}
    Mcombo := THDBValComboBox(Getcontrol('ETB_MVALOMS'));
{$ELSE}
    Mcombo := THValComboBox(Getcontrol('ETB_MVALOMS'));
{$ENDIF}
    if Mcombo <> nil then
    begin
      if MCombo.value = '' then
      begin
        LastError := 130;
        PgiBox('Veuillez saisir la méthode de calcul du maintien de salaire.', Ecran.caption); { PT55-2 }
        PositionneFocusOngletCP;
        SetFocusControl('ETB_MVALOMS');
        exit;
      end;
    end;

    // ETB_NBREACQUISCP : Controle zone non supérieure à 30
{$IFNDEF EAGLCLIENT}
    MEdit := THDBEdit(Getcontrol('ETB_NBREACQUISCP'));
{$ELSE}
    MEdit := THEdit(Getcontrol('ETB_NBREACQUISCP'));
{$ENDIF}
    if MEdit <> nil then
    begin
      if Valeur(MEdit.text) > 10 then
      begin
        LastError := 104;
        PositionneFocusOngletCP;
        PgiBox('Veuillez saisir un nombre de jours acquis inférieur à 10.', Ecran.caption); { PT55-2 }
        SetFocusControl('ETB_NBREACQUISCP');
        exit;
      end;

    { DEB PT82-2 }
   if (GetField('ETB_DATEACQCPANC') <> '') then
   if not PgOkFormatDateJJMM(Copy(Getfield ('ETB_DATEACQCPANC'),1,2)+Copy(Getfield ('ETB_DATEACQCPANC'),3,2)) then
     Begin
     PgiBox('Format date incorrect. Vous devez saisir une date au format jj/mm.',Ecran.caption);
     if GetField('ETB_DATEACQCPANC') <> '' then  setfield('ETB_DATEACQCPANC', '');
     End;
   { FIN PT82-2 }
    end;
  end; // end du test  "if MCheckBox.Checked"
end;
{$ENDIF}
//FIN PT62


procedure TOM_EtabCompl.PositionneFocusOngletCP;
var
  P: TPageControl;
  T: TTabSheet;
begin
  P := TPageControl(GetControl('PAGES'));
  if P <> nil then
  begin
    T := TTabsheet(getcontrol('PCONGES'));
    P.ActivePage := T;
  end;
end;

procedure TOM_EtabCompl.InitialiseZonesCongesPayes;
begin
  SetField('ETB_DATECLOTURECPN', Idate1900); // beurk
  SetField('ETB_ETABLISSEMENT', etablissement);
  SetField('ETB_LIBELLE', LIBELLE);
  SetField('ETB_NBJOUTRAV', 0);
  SetField('ETB_NBREACQUISCP', 0);
  SetField('ETB_NBACQUISCP', '');
  SetField('ETB_NBRECPSUPP', 0); { PT52-2 }
  SetField('ETB_TYPDATANC', '1'); //PORTAGE CWAS Champ combo non integer
  SetField('ETB_DATEACQCPANC', '');
  SetField('ETB_VALORINDEMCP', '');
  SetField('ETB_RELIQUAT', '');
  SetField('ETB_PAIEVALOMS', ''); { PT70 }
  SetField('ETB_PROFILCGE', '');
  SetField('ETB_BASANCCP', '');
  SetField('ETB_VALANCCP', '');
  SetField('ETB_PERIODECP', 0);
  SetField('ETB_1ERREPOSH', '');
  SetField('ETB_2EMEREPOSH', '');
  //SetField('ETB_STANDCALEND','');   {PT- 5}
  SetField('ETB_MVALOMS', '');
  SetField('ETB_VALODXMN', 0);
  SetField('ETB_TYPDATANC', '');
  SetField('ETB_MVALOMS', '');
  SetField('ETB_EDITBULCP', ''); //PT46  { PT55-1 }
{$IFNDEF EAGLCLIENT}
  // PT38-1     MCombo := THDBVALComboBox(Getcontrol('ETB_BASANCCP'));
{$ELSE}
  //      MCombo := THVALComboBox(Getcontrol('ETB_BASANCCP'));
{$ENDIF}
  //      SetControlVisible ('EETB_VALANCCP', (MCombo.value = 'VAL')); ***SB
  //      SetControlVisible ('CETB_VALANCCP', (MCombo.value <> 'VAL'));

end;

procedure TOM_EtabCompl.EnableChampsOngletCongesPayes;
begin
  SetControlEnabled('TBCLOTCP', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('TBDECLOTCP', ((Getfield('ETB_CONGESPAYES') = 'X') and ExistCloture = True));
  SetControlEnabled('CONGES', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_NBJOUTRAV', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_DATECLOTURECPN', ((Getfield('ETB_CONGESPAYES') = 'X') and ExistCloture = False));
  if Getfield('ETB_CONGESPAYES') = '-' then SetControlEnabled('ETB_DATECLOTURECPN', False); { PT55-1 }
  SetControlEnabled('ETB_NBREACQUISCP', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_NBACQUISCP', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_NBRECPSUPP', (Getfield('ETB_CONGESPAYES') = 'X')); { PT52-2 }
  SetControlEnabled('ETB_VALORINDEMCP', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_RELIQUAT', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_PAIEVALOMS', (Getfield('ETB_CONGESPAYES') = 'X')); { PT70 }
  SetControlEnabled('ETB_PROFILCGE', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_BASANCCP', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_1ERREPOSH', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetControlEnabled('ETB_2EMEREPOSH', ((getfield('ETB_NBJOUTRAV') = '5') and (Getfield('ETB_CONGESPAYES') = 'X')));
  SetControlEnabled('L2EME', ((getfield('ETB_NBJOUTRAV') = '5') and (Getfield('ETB_CONGESPAYES') = 'X')));
  SetControlEnabled('ETB_MVALOMS', (Getfield('ETB_CONGESPAYES') = 'X'));
  SetcontrolEnabled('ETB_VALODXMN', ((getfield('ETB_MVALOMS') = 'M') and (Getfield('ETB_CONGESPAYES') = 'X')));
  SetcontrolEnabled('LETB_VALODXMN', ((getfield('ETB_MVALOMS') = 'M') and (Getfield('ETB_CONGESPAYES') = 'X')));
  { DEB PT38-1 }
  SetControlProperty('ETB_VALANCCP', 'ElipsisButton', True);
  if Getfield('ETB_BASANCCP') = 'VAR' then
    SetControlProperty('ETB_VALANCCP', 'Datatype', 'PGVARIABLE')
  else
    if Getfield('ETB_BASANCCP') = 'TAB' then
    SetControlProperty('ETB_VALANCCP', 'Datatype', 'PGTABINTANC')
  else
  begin
    SetControlProperty('ETB_VALANCCP', 'Datatype', '');
    SetControlProperty('ETB_VALANCCP', 'ElipsisButton', False);
  end;
  SetControlEnabled('ETB_VALANCCP', ((Getfield('ETB_BASANCCP') <> '') and (Getfield('ETB_CONGESPAYES') = 'X')));
  SetControlEnabled('ETB_TYPDATANC', ((Getfield('ETB_BASANCCP') <> '') and (Getfield('ETB_CONGESPAYES') = 'X')));
  SetcontrolEnabled('ETB_DATEACQCPANC', ((getfield('ETB_TYPDATANC') = '2') and (Getfield('ETB_CONGESPAYES') = 'X')));
  { FIN PT38-1 }
  SetcontrolEnabled('ETB_EDITBULCP', (Getfield('ETB_CONGESPAYES') = 'X')); //PT46
end;

procedure TOM_EtabCompl.CongesPayesRechercheTabletteAnciennete;
begin
  { DEB PT38-1 }
  SetControlProperty('ETB_VALANCCP', 'ElipsisButton', True);
  if GetField('ETB_BASANCCP') = 'VAR' then
    SetControlProperty('ETB_VALANCCP', 'Datatype', 'PGVARIABLE')
  else
    if GetField('ETB_BASANCCP') = 'TAB' then
    SetControlProperty('ETB_VALANCCP', 'Datatype', 'PGTABINTANC')
  else
  begin
    SetControlProperty('ETB_VALANCCP', 'Datatype', '');
    SetControlProperty('ETB_VALANCCP', 'ElipsisButton', False);
  end;
  { FIN PT38-1 }
end;

procedure TOM_Etabcompl.PCasRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: boolean);
var
  PPart: THValCombobox;
  Theme: THValCombobox;
begin
  PPart := THValCombobox(GetControl('PPART'));
  Theme := THValCombobox(GetControl('THEME'));
  if (PPArt <> nil) and (Theme <> nil) then PPart.Plus := Theme.Values[THGrid(Sender).Row - 1];
end;

procedure TOM_Etabcompl.InitialiseMvtCp(etab: string);
begin
  //PT1 : 12/06/01 On ne doit supprimer que les mvts de type CPA
  ExecuteSQL('DELETE FROM ABSENCESALARIE WHERE PCN_TYPEMVT="CPA" '+ 
             'AND PCN_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES '+  { PT77-2 }
             'WHERE PSA_ETABLISSEMENT="' + etab + '")');
end;

//PT62
{$IFNDEF DADSB}

procedure TOM_Etabcompl.BTnCompteBqClick(Sender: TObject);
begin
 TRLanceFiche_BanqueCP('TR','TRMULCOMPTEBQ','','','');
end;
{$ENDIF}
//FIN PT62

{
////                       Mes erreurs                                      ////
LastError;LastErrorMsg;Onglet
------------------------- Onglet Congés Payés
101 ; LastErrorMsg:='Format incorrect' ; Onglet Congés Payés
102 ; LastErrorMsg:='Cette valeur ne peut être supérieure à 6'
103 ; LastErrorMsg:='Format incorrect' ; Onglet Congés Payés
104 ; LastErrorMsg:='Cette valeur ne peut être supérieure à 10'
105 ; LastErrorMsg:='Le 1er jour de repos hebdomadaire doit être renseigné' ;
106 ; LastErrorMsg:='Le 2ème jour de repos hebdomadaire doit être renseigné' ;
107 ; LastErrorMsg:='Les 2 jours de repos hebdomadaire ne peuvent pas être identiques' ;

}
// PT-4

procedure TOM_ETABCOMPL.AccesNumIS(Sender: TObject); //PT16
begin
  if Sender = nil then Exit;
{$IFNDEF EAGLCLIENT}
  if TDBCheckBox(Sender).Name = 'ETB_ISLICSPEC' then
  begin
    if TDBCheckBox(Sender).Checked = True then SetControlVisible('ETB_ISNUMLIC', true) //PT22
    else SetControlVisible('ETB_ISNUMLIC', False);
  end;
  if TDBCheckBox(Sender).Name = 'ETB_ISLABELP' then
  begin
    if TDBCheckBox(Sender).Checked = True then SetControlVisible('ETB_ISNUMLAB', true) //PT22
    else SetControlVisible('ETB_ISNUMLAB', False);
  end;
{$ELSE}
  if TCheckBox(Sender).Name = 'ETB_ISLICSPEC' then
  begin
    if TCheckBox(Sender).Checked = True then SetControlVisible('ETB_ISNUMLIC', true) //PT22
    else SetControlVisible('ETB_ISNUMLIC', False);
  end;
  if TCheckBox(Sender).Name = 'ETB_ISLABELP' then
  begin
    if TCheckBox(Sender).Checked = True then SetControlVisible('ETB_ISNUMLAB', true) //PT22
    else SetControlVisible('ETB_ISNUMLAB', False);
  end;
{$ENDIF}
end;

//   PT18   23/04/2002 V582 Ph suppression etablissement complémentaire

procedure TOM_EtabCompl.OnDeleteRecord;
var
  Etab: string;
begin
  inherited;
  Etab := GetField('ETB_ETABLISSEMENT');
  if Etab = '' then Exit;
  if Etab = VH^.EtablisDefaut then
  begin
    LastError := 1;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;
  if IsMouvementer then Exit;
end;

function TOM_EtabCompl.IsMouvementer: Boolean;
var
  Sql, St: string;
begin
  Result := True;
  Sql := 'Select ETB_ETABLISSEMENT From ETABCOMPL Where ETB_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '" And ';
  St := Sql + '(Exists (Select PSA_ETABLISSEMENT From SALARIES Where PSA_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 2;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PPU_ETABLISSEMENT From PAIEENCOURS Where PPU_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 3;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PHB_ETABLISSEMENT From HISTOBULLETIN Where PHB_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 4;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PEI_ETABLISSEMENT From EMPLOIINTERIM Where PEI_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 5;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PAT_ETABLISSEMENT From TAUXAT Where PAT_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 6;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PCN_ETABLISSEMENT From ABSENCESALARIE Where PCN_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 7;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PSD_ETABLISSEMENT From HISTOSAISRUB Where PSD_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 8;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select POG_ETABLISSEMENT From ORGANISMEPAIE Where POG_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 9;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PDU_ETABLISSEMENT From DUCSENTETE Where PDU_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 10;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PMR_ETABLISSEMENT From MASQUESAISRUB Where PMR_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 11;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PJT_ETABLISSEMENT From JOURSTRAV Where PJT_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 12;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  St := Sql + '(Exists (Select PDO_ETABLISSEMENT From DONNEURORDRE Where PDO_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 13;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;

  // d PT35
  St := Sql + '(Exists (Select PRT_ETABLISSEMENT From CDETICKETS Where PRT_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '"))';
  if ExisteSQL(St) then
  begin
    LastError := 14;
    LastErrorMsg := TexteMessage[LastError];
    Exit;
  end;
  // f PT35
  Result := False;
end;
// FIN PT18
// PT28   16/01/2003 V591 PH FQ 10432 gestion des taux at avec rafraichissement de la liste des taux

procedure TOM_EtabCompl.FTAUXDblClick(Sender: TObject);
var
  st: string;
begin
  st := GetField('ETB_ETABLISSEMENT');
  AglLanceFiche('PAY', 'TAUXAT', GetField('ETB_ETABLISSEMENT'), '', GetField('ETB_ETABLISSEMENT'));
  AfficheTauxAT;
end;

procedure TOM_EtabCompl.AfficheTauxAT;
var
  St: string;
  T1: TOB;
  Q: TQuery;
begin
  // PORTAGECWAS
  // PT27   17/12/2002 V591 PH Portage CWAS = THGrid sur grille des taux AT
  if (THGrid(GetControl('FLTAUXAT')) <> nil) then
  begin
    st := 'SELECT PAT_ETABLISSEMENT, PAT_ORDREAT, PAT_DATEVALIDITE, PAT_TAUXAT, PAT_LIBELLE,' +
      ' CO_LIBELLE as BUREAU, PAT_CODERISQUE, PAT_SECTIONAT  FROM TAUXAT ' +
      ' LEFT JOIN COMMUN ON PAT_CODEBUREAU=CO_CODE AND CO_TYPE="PCB" ' +
      ' WHERE PAT_ETABLISSEMENT="' + GetField('ETB_ETABLISSEMENT') + '" ORDER BY PAT_ETABLISSEMENT ,PAT_ORDREAT ,PAT_DATEVALIDITE';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then
    begin
      T1 := TOB.Create('Les taux', nil, -1);
      T1.LoadDetailDb('TAUXAT', '', '', Q, FALSE, False);
      if T1.detail.count > 0 then
        T1.PutGridDetail(THGrid(GetControl('FLTAUXAT')), FALSE, TRUE, 'PAT_ETABLISSEMENT;PAT_ORDREAT;PAT_DATEVALIDITE;PAT_TAUXAT;PAT_LIBELLE;BUREAU;PAT_CODERISQUE;PAT_SECTIONAT',
          True);
      T1.Free;
      //   HMTrad.ResizeGridColumns (THGRID(GetControl('FRIB')));
    end;
    Ferme(Q);
  end;
  // FIN PT27
  // FIN PT28
end;
{ DEB PT32 Affectation .plus profil pour raffraichir tablette }

procedure TOM_EtabCompl.ChargeProfilActivite(Lechamp, Nat: string);
begin
  if RechDom('PGTYPEPROFIL', Nat, True) = 'ACT' then
  begin
    if GetField('ETB_ACTIVITE') <> '' then
      SetControlProperty(LeChamp, 'Plus', 'AND (PPI_ACTIVITE="' + GetField('ETB_ACTIVITE') + '" OR PPI_ACTIVITE="" OR PPI_ACTIVITE IS NULL)')
    else
      SetControlProperty(LeChamp, 'Plus', '');
  end;
end;
{ FIN PT32 }
{ DEB PT32 evenement sur entrée profil pour raffraichir tablette }

procedure TOM_EtabCompl.OnEnterProfil(Sender: Tobject);
var
  Name: string;
begin
  if sender = nil then exit;
  { Affectation de la propriété plus pour chaque profil gérant l'activité }
{$IFNDEF EAGLCLIENT}
  Name := THDBValComboBox(sender).Name;
{$ELSE}
  Name := THValComboBox(sender).Name;
{$ENDIF}
  if Name = 'ETB_PROFILAFP' then ChargeProfilActivite('ETB_PROFILAFP', 'AFP')
  else if Name = 'ETB_PROFILANCIEN' then ChargeProfilActivite('ETB_PROFILANCIEN', 'ANC')
  else if Name = 'ETB_PROFILAPP' then ChargeProfilActivite('ETB_PROFILAPP', 'APP')
  else if Name = 'ETB_PROFILCGE' then ChargeProfilActivite('ETB_PROFILCGE', 'CGE')
  else if Name = 'ETB_PROFILMUT' then ChargeProfilActivite('ETB_PROFILMUT', 'MUT')
  else if Name = 'ETB_PERIODBUL' then ChargeProfilActivite('ETB_PERIODBUL', 'PER')
  else if Name = 'ETB_PROFILPRE' then ChargeProfilActivite('ETB_PROFILPRE', 'PRE')
  else if Name = 'ETB_PROFIL' then ChargeProfilActivite('ETB_PROFIL', 'PRO')
  else if Name = 'ETB_PROFILRBS' then ChargeProfilActivite('ETB_PROFILRBS', 'RBS')
  else if Name = 'ETB_PROFILREM' then ChargeProfilActivite('ETB_PROFILREM', 'REM')
  else if Name = 'ETB_PROFILRET' then ChargeProfilActivite('ETB_PROFILRET', 'RET')
  else if Name = 'ETB_REDREPAS' then ChargeProfilActivite('ETB_REDREPAS', 'RRE')
  else if Name = 'ETB_REDRTT1' then ChargeProfilActivite('ETB_REDRTT1', 'RT1')
  else if Name = 'ETB_REDRTT2' then ChargeProfilActivite('ETB_REDRTT2', 'RT2')
  else if Name = 'ETB_PROFILTRANS' then ChargeProfilActivite('ETB_PROFILTRANS', 'TRA')
  else if Name = 'ETB_PROFILTSS' then ChargeProfilActivite('ETB_PROFILTSS', 'TSS');
end;
{ FIN PT32 }

procedure TOM_EtabCompl.OnClose;
begin
  inherited;
{$IFDEF EAGLCLIENT}
  AvertirCacheServer('ETABCOMPL');
{$ENDIF}
  { DEB PT55-1 }
  if ForceValidCp then
  begin
    LastError := 1;
    PgiBox('Vous avez modifié la gestion des congés payés, vous devez valider votre saisie.', Ecran.caption);
  end;
  { FIN PT55-1 }
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/10/2003
Modifié le ... :   /  /
Description .. : Imprimer
Mots clefs ... : PAIE;PGETAB
*****************************************************************}
procedure TOM_EtabCompl.Imprimer(Sender: TObject);
var
Pages: TPageControl;
StPages : string;
begin
//PT85-1
SetControlText ('TRIBSALAIRE_BQCP', RechDom ('TTBANQUECP',
                GetControlText ('ETB_RIBSALAIRE'), False,
                'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE));
SetControlText ('TRIBFRAIS_BQCP', RechDom ('TTBANQUECP',
                GetControlText ('ETB_RIBFRAIS'), False,
                'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE));
SetControlText ('TRIBACOMPTE_BQCP', RechDom ('TTBANQUECP',
                GetControlText ('ETB_RIBACOMPTE'), False,
                'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE));
SetControlText ('TRIBDUCSEDI_BQCP', RechDom ('TTBANQUECP',
                GetControlText ('ETB_RIBDUCSEDI'), False,
                'BQ_NODOSSIER="'+V_PGI.NoDossier+'" AND'+BQCLAUSEWHERE));
//FIN PT85-1
Pages := TPageControl(getControl('Pages'));
StPages := AglGetCriteres (Pages, FALSE);
LanceEtat('E', 'PAY', 'ETB', True, False, False, Pages, '', '', False, 0, StPages);
end;


procedure TOM_EtabCompl.ChargeVarGblIdemEtab;
begin
  Activite := GetField('ETB_ACTIVITE'); //PT32
  Profil := GetField('ETB_PROFIL');
  PeriodBul := GetField('ETB_PERIODBUL');
  //  PT8 : 01/10/2001 V562 PH Nouveau champ ETB_PROFILREM
  ProfilRem := GetField('ETB_PROFILREM');
  ProfilRbs := GetField('ETB_PROFILRBS');
  ProfilAfp := GetField('ETB_PROFILAFP');
  {temp:=GetField('ETB_PCTFRAISPROF');
  if temp=null then    }
  if (GetField('ETB_PCTFRAISPROF') = null) then PourcentAfp := '0' // beurk
  else PourcentAfp := GetField('ETB_PCTFRAISPROF');
  ProfilApp := GetField('ETB_PROFILAPP');
  // PT17   02/04/2002 V571 Ph Ajout champ profil retraite
  ProfilRet := GetField('ETB_PROFILRET');
  ProfilMut := GetField('ETB_PROFILMUT');
  ProfilPre := GetField('ETB_PROFILPRE');
  ProfilTss := GetField('ETB_PROFILTSS');
  ProfilCge := GetField('ETB_PROFILCGE');
  ProfilAncien := GetField('ETB_PROFILANCIEN');
  ProfilTrans := GetField('ETB_PROFILTRANS');
  StandCalend := GetField('ETB_STANDCALEND');
  if GetField('ETB_NBREACQUISCP') = null then NbreAcquisCp := '' // beurk
  else NbreAcquisCp := GetField('ETB_NBREACQUISCP');
  NbAcquisCp := GetField('ETB_NBACQUISCP');
  NbCpSupp := Valeur(GetField('ETB_NBRECPSUPP')); { PT52-2 }
  ValorIndempCp := GetField('ETB_VALORINDEMCP');
  basancp := GetField('ETB_BASANCCP');
  ValancCp := GetField('ETB_VALANCCP');
  DateAcqCpAnc := GetField('ETB_DATEACQCPANC');
  ValtypDatAnc := GetField('ETB_TYPDATANC');
  DatAnc := GetField('ETB_DATANC');
  Reliquat := GetField('ETB_RELIQUAT');
  PaieValoMs := GetField('ETB_PAIEVALOMS');   { PT70 }
  RedRepas := GetField('ETB_REDREPAS');
  RedRtt1 := GetField('ETB_REDRTT1');
  RedRtt2 := GetField('ETB_REDRTT2');
  ModeReglement := GetField('ETB_PGMODEREGLE'); {PT3}
  RibSalaire := GetField('ETB_RIBSALAIRE');
  PaiAcompte := GetField('ETB_PAIACOMPTE');
  RibAcpSoc := GetField('ETB_RIBACOMPTE');
  PaiFrais := GetField('ETB_PAIFRAIS');
  RibFraisSoc := GetField('ETB_RIBFRAIS');
  MoisPaiement := GetField('ETB_MOISPAIEMENT');
  if GetField('ETB_JOURPAIEMENT') = null then JourPaiement := '0' // beurk
  else JourPaiement := GetField('ETB_JOURPAIEMENT');
  JourHeure := GetField('ETB_JOURHEURE');
  if JourHeure = '' then
  begin
    SetField('ETB_JOURHEURE', VH_Paie.PGJourHeure);
    JourHeure := VH_Paie.PGJourHeure;
  end;
  TypeFraction := GetField('ETB_CODESECTION'); //PT14
  EditOrg := GetField('ETB_EDITORG'); //PT24-1
  TicketRestau := GetField('ETB_TYPTICKET'); // PT33
  GblEditBulCP := GetField('ETB_EDITBULCP'); //PT46
// d PT81
  if (VH_Paie.PGTicketRestau) then
  begin
    if (V_PGI.NumVersionBase < 800) then
    // (spécif V7)
    begin
      PointLivraison := GetField('ETB_AUTRENUMERO');
    end
    else
    // ( à partir V8)
    begin
      PointLivraison := GetField('ETB_TICKLIVR');
    end;
  end;
// f PT81
end;
{ FIN PT46 }

//DEBUT PT73
procedure TOM_ETABCOMPL.AccesMedecine(Sender : TObject);
Var Ret : String;
    {$IFNDEF EAGLCLIENT}
    Combo : THDBValComboBox;
    {$ELSE}
    Combo : THValComboBox;
    {$ENDIF}
begin
     Ret := AGLLanceFiche('PAY','PGANNUAIRE','','','COMBO;MED');
     If Ret <> 'VIDE' then
     begin
          {$IFNDEF EAGLCLIENT}
          Combo := THDBValComboBox(GetControl('ETB_MEDTRAVGU'));
          {$ELSE}
          Combo := THValComboBox(GetControl('ETB_MEDTRAVGU'));
          {$ENDIF}
          Combo.ReLoad;
          ForceUpdate;
          SetField('ETB_MEDTRAVGU',Ret);
          SetControlText('ETB_MEDTRAVGU',Ret);
     end;
end;

procedure TOM_ETABCOMPL.AccesDDtefp(Sender : TObject);
Var Ret : String;
    {$IFNDEF EAGLCLIENT}
    Combo : THDBValComboBox;
    {$ELSE}
    Combo : THValComboBox;
    {$ENDIF}
begin
     Ret := AGLLanceFiche('PAY','PGANNUAIRE','','','COMBO;DDT');
     If Ret <> 'VIDE' then
     begin
          {$IFNDEF EAGLCLIENT}
          Combo := THDBValComboBox(GetControl('ETB_CODEDDTEFPGU'));
          {$ELSE}
          Combo := THValComboBox(GetControl('ETB_CODEDDTEFPGU'));
          {$ENDIF}
          Combo.ReLoad;
          ForceUpdate;
          SetField('ETB_CODEDDTEFPGU',Ret);
          SetControlText('ETB_CODEDDTEFPGU',Ret);
     end;

end;
//FIN PT73

// d PT81
{$IFDEF EAGLCLIENT}
procedure TOM_ETABCOMPL.SaisieCodeLivr(Sender : TObject);
begin
          ForceUpdate;
          SetField ('ETB_AUTRENUMERO',GetControlText('ETB_AUTRENUMERO_')) ;
end;
{$ENDIF}
// f PT81
//D PT186

procedure TOM_ETABCOMPL.AjoutModifsHisto(Champ,ChampType,Valeur : String);
var i : Integer;
    TobSal,TobHisto,TH : Tob;
    Q : TQuery;
    Tablette,LeType : String;
    Etab : String;
    DateHisto:TDateTime; //PT91
begin
  //DEB PT91
  if Champ = 'PSA_CONVENTION' then
    DateHisto := StrToDateTime(GetControlText('DATEHISTO'))
  else
    DateHisto := V_PGI.DateEntree;
  //FIN PT91
  If Not ExisteSQL('SELECT PPP_PGINFOSMODIF FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'" AND PPP_PGTYPEINFOLS="SAL" AND PPP_HISTORIQUE="X"') then Exit;
  Q := OpenSQL('SELECT PPP_LIENASSOC,PPP_PGTYPEDONNE FROM PARAMSALARIE WHERE PPP_PGINFOSMODIF="'+Champ+'" AND PPP_PREDEFINI="CEG"',True);
  If Not Q.Eof then
  begin
    Tablette := Q.FindField('PPP_LIENASSOC').AsString;
    LeType := Q.FindField('PPP_PGTYPEDONNE').AsString;
  end;
  Ferme(Q);
  Etab := GetField('ETB_ETABLISSEMENT');
  ExecuteSQL('DELETE FROM PGHISTODETAIL WHERE PHD_PGINFOSMODIF="'+Champ+'" AND PHD_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE '+ChampType+'="ETB" AND PSA_ETABLISSEMENT="'+Etab+'") '+
  'AND PHD_DATEAPPLIC="'+UsDateTime(DateHisto)+'"');        //PT91
  TobHisto := Tob.Create('Ligne historique',Nil,-1);
  Q := OpenSQL('SELECT PSA_SALARIE FROM SALARIES WHERE '+ChampType+'="ETB" AND PSA_ETABLISSEMENT="'+Etab+'"',True);
  TobSal := Tob.Create('Les salaries',Nil,-1);
  TobSal.LoadDetailDB('Les salaries','','',Q,False);
  Ferme(Q);
  For i := 0 to TobSal.Detail.Count - 1 do
  begin
     TH := Tob.Create('PGHISTODETAIL',TobHisto,-1);
     TH.PutValue('PHD_SALARIE',TobSal.Detail[i].Getvalue('PSA_SALARIE'));
     TH.PutValue('PHD_ETABLISSEMENT','');
     TH.PutValue('PHD_ORDRE',-1);
     TH.PutValue('PHD_GUIDHISTO',AglGetGuid());
     TH.PutValue('PHD_PGINFOSMODIF',Champ);
     TH.PutValue('PHD_PGTYPEHISTO','002');
     TH.PutValue('PHD_ANCVALEUR','');
     TH.PutValue('PHD_NEWVALEUR',Valeur);
     TH.PutValue('PHD_TYPEVALEUR',LeType);
     TH.PutValue('PHD_TABLETTE',Tablette);
     TH.PutValue('PHD_DATEAPPLIC',DateHisto);    //PT91
     TH.PutValue('PHD_TRAITEMENTOK','X');
     TH.PutValue('PHD_DATEFINVALID',IDate1900);
     TH.PutValue('PHD_TYPEBUDG','');
     TH.PutValue('PHD_NUMAUG',0);
     TH.PutValue('PHD_PGTYPEINFOLS','SAL');
     TH.PutValue('PHD_ANNEE','');
//     TH.InsertDB(Nil);
//     TH.Free;
  end;
  TobSal.Free;
  TobHisto.InsertDB(Nil);
  TobHisto.Free;

end;
//F PT186

initialization
  registerclasses([TOM_ETABCOMPL]);

end.




{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Unit de génération des ODs de paie dans la compta.
Suite ........ : 3 méthodes sont  disponlibles :
Suite ........ : - dans compta pgi en immédiat
Suite ........ : - idem mais en différé
Suite ........ : - format interne à la paie pour être traiter par une autre
Suite ........ : compta
Suite ........ : Possibilité de travailler soit en journal de simulation, soit
Suite ........ : avec un journal OD dont les écritures seront validées et non
Suite ........ : modifiables
Mots clefs ... : PAIE;GENERCOMPTA
*****************************************************************}
{ Traitement des ecritures comptables
Si ecritures de simulations alors pas d'affectation  de numéro de piece et pas mise à jour des soldes des comptes
Tient compte de la decomposition des comptes en fonction des paramètrages société
Traite les ecritures analytiques comme des filles de la tob des ecritures générales
Si déséquilibre alors on ne peut pas passer les ecritures comptables
}

{ PT-1 : 02/08/2001 : V547  PH  Les écritures de simulations ne sont pas validées
         ce qui permettra de les modifier et de valider alors le journal
  PT-2 : 07/09/2001 : V547  PH Correction génération multi etab
  PT-3 : 21/09/2001 : V547  PH Ecart de conversion
  PT-4 : 22/11/01   : V562  PH ORACLE, Le champ PHB_ORGANISME est obligatoirement renseigné
  PT-5 : 03/12/01   : V563  PH Correction pb compte 421 Collectif test inversé
  PT-6 : 03/12/01   : V563  PH test equilibre débit/crédir sur chaque FOLIO
  PT-7 : 11/12/01   : V563  PH test si les compteurs des journaux existent
  PT-8 : 19/02/02   : V571  PH tests J_MODESAISIE (BOR,LIB,-) et G_NATUREGENE
       pour affectation Mode de saisie et du lettrage corrects dans la ligne ecriture
  PT-9 : 19/02/02   : V571  PH test compte de la classe 2 (remboursement de prets)
  PT10 : 22/03/02   : V575  PH activation passege des ODs en différe dans la compta PGI
  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer et non '421'
  PT12 : 26/03/02   : V575  PH Rajout ligne totalisation dans la TOB pour imprimer les totaux
  PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
  PT14 : 02/05/02   : V582  PH Référence interne dans toutes les lignes ecritures
  PT15 : 02/05/02   : V582  PH compte 425 et 427 avec auxiliaire salarié traitement idem 421  Point supprimé ==> PT21
  PT16 : 02/05/02   : V582  PH Si compte lettrable alors date échéance = date écriture car si plusieurs date de
                      paiement il faudrait générer les Ods pas date de paiement pour chaque bulletin
  PT17 : 03/05/02   : V582  PH Pb AGL 540 et disconcordance des champs entre les tables ECRITURE et ANALYTIQ Y_VISION
  PT18 : 04/06/02   : V582  PH Récupération et personnalisation de la fonction insertiondifféree
  PT19 : 04/06/02   : V582  PH Gestion de la trace en cas d'erreur : Ajout paramètre ZTraceErr
  PT20 : 05/06/02   : V582  PH Gestion historique des évènements
  PT21 : 12/08/02   : V585  PH Rajout du compte auxilaire pour les comptes de racine 42 (ex  425)
  PT22 : 13/08/02   : V585  PH Test si ecriture de simulation et journal pas de type piece alors erreur
  PT23 : 23/08/02   : V585  PH Traitement des comptes classe 5 pour jnal de paiement
  PT24 : 23/08/02   : V585  PH Confection du numéro de compte en fonction du mois de paie
  PT25 : 17/12/02   : V585  PH Test de la query non nulle avant de lancer la génération
  PT26 : 07/01/03   : V591  PH Affectation champs table ecriture si Auxiliaire lettrable FQ 10428
  PT27 : 04/03/03   : V_42  PH Initialisation du champ E_QUALIFORIGINE à GEN pour la validation des écritures de simuls
  PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine en fonction des paramètres
                               société génération comptable
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  PT29 : 19/08/03  : V_421 PH FQ 10759,10532,10507,10294
  PT30 : 10/09/03  : V_421 PH Suppression des traitements concernant les champs EURO de la compta
  PT31 : 25/09/03  : V_421 PH FQ 10726 Comptabilisation sur des organismes non CEGID non reproduit
                      ==> Si erreur alors message indique que le compte n'est pas présent dans le modèle
  PT32 : 12/10/03  : V_421 PH Test des ecritures négatives uniquement en fin de calcul pour traiter les paies
                               et les écritures centrailisées
  PT33 : 14/11/03  : V_500 PH Affinage message anomalie compte non present dans le modele uniquement si déséquilibre
  PT34 : 24/02/04  : V_500 PH Prise en compte des comptes des classes 1 et 9
  PT35 : 25/02/04  : V_500 PH On ne tient plus compte de la monnaie de tenue de la compta si elle n'est pas tenue en EURO
                              car le franc n'existe plus
  PT36 : 09/03/2004 : V_500 FQ 11154 PH traitement cas particulier compte dont les longueurs ne sont pas prédéfinies.
  PT37 : 11/03/2004 : V_500 FQ 11146 PH Prise en compte des comptes collectifs salariés autre que 42 ==> racine classe 4
  PT38 : 11/03/2004 : V_500 FQ 11154 PH Recherche du compte bancaire de l'établissement dans le cas d'un journal de réglement
  PT39 : 19/03/2004 : V_500 FQ 11173 PH Affection de la valeur par defaut correspondant à la racine du compte
  PT40 : 24/03/2004 : V_500 FQ 11207 PH Affectation du mode de réglement par defaut sur les comptes de auxiliaires lettrables
  PT41 : 31/03/2004 : V_500 PH compte du type du modèle qui corrrespond à un journal ODs et non de provisions
  PT42 : 06/05/2004 : V_500 PH FQ 11173 Mise en place liste d'exportation
  PT43 : 08/06/2004 : V_500 PH FQ 11334 Titre avant validation
  PT44 : 08/06/2004 : V_500 PH FQ 10990 Onglet message erreurs dans jnal evt
  PT45 : 11/08/2004 : V_500 PH FQ 11314 gestion des comptes inexistants si écriture dans compta PGI
  PT46 : 11/08/2004 : V_500 PH FQ 11492 Ecritures analytiques au format paie PGI
  PT47 : 11/08/2004 : V_500 PH Decomposition des racines de comptes sélectionnés.
  PT48 : 06/09/2004 : V_500 PH Prise en compte des différents type de modèles
  PT49 : 13/09/2004 : V_500 PH FQ 11591 Abandon complet si folio non equilibré
  PT50 : 13/10/2004 : V_500 PH FQ 11318 Cohérence des champs de la compta
  PT51 : 20/12/2004 : V_600 PH Message alerte si Journal type bordereau et 2 folios dans le modèle
  PT52 : 25/01/2005 : V_601 PH FQ 11948 Propagation du champ confidentiel dans la compta
  PT53 : 09/02/2005 : V_602 PH Refonte complète du source pour optimisation des traitements
  PT53-1 28/02/2005 ; V_602 PH Modifs fcts et Appels à ConfectionCpte et LeChampSal Qmul remplacée par TPPU
  PT54 : 14/02/2005 : V_602 PH On ne décompose plus sur la longueur de la racine si racine non définie Suite PT47
  PT55 : 01/03/2005 : V_602 PH FQ 12052 Prise en compte des décompositions pour les dotations CP et RTT
  PT56-1 : 18/05/2005 : V_602 PH FQ 11977 Libellé du compte auxiliaire au lieu du compte général
  PT56-2 : 18/05/2005 : V_602 PH FQ 11948 prise en compte de la confidentialité du compte auxiliaire et non du cpte géné
  PT57 : 18/05/2005 : V_600 PH FQ 12237 synchronisation avec S1.
  PT58 : 29/06/2005 : V_600 PH FQ 12045 Erreur de transaction avec méthode format paie pgi.
  PT59 : 10/08/2005 : V_600 PH FQ 11796 Message indiquant journal de type bordereau donc 1 folio
  PT60 : 29/09/2005 : V_600 PH FQ 12382 Prise en compte critère salarié datesortie en update paieencours
  PT61 : 29/09/2005 : V_600 PH FQ 12599 Ergonomie Messaged'anomalie
  PT62 : 03/04/2006 : V_650 PH FQ 12983 Rechargement des paramsoc de la compta
  PT63 : 03/04/2006 : V_650 PH FQ 13000 Suppression erreur mais pas maj top ecriture générée
  PT64 : 03/05/2006 : V_650 PH FQ 12931,12932 ??? même pb libellé compte auxiliaire
  PT65 : 04/05/2006 : V_650 PH FQ 12970 Controle de la longueur du compte Net A payer
  PT65-1 : 12/06/2006 : V_650 PH FQ 12970 Suite Controle de la longueur du compte Net A payer < et >
  PT66 : 27/11/2006 : V_700 PH FQ 13698 Organisme inconnu n'est plus une anomalie bloquante.
  PT67 : 16/04/2007 : V_700 PH Plus de contrôle cohérence compta si non intégration dans compta PGI
  PT68 : 22/05/2007 : V_700 PH Prise en compte Vista sur les nouveaux répertoires d'installation
  PT69 : 11/06/2007 : V_720 VG Adaptation nouvelles méthodes cbppath
  PT70 : 25/06/2007 : V_720 PH FQ 11796 Ergonomie
  PT71 : 26/06/2007 : V_720 PH FQ 13757 Compte de nature DIV et lettrable
  PT72 : 13/09/2007 : V_80  FC FQ 13716 Edition du journal des erreurs
  }
 
unit UTofPG_GeneCompta;
interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  Windows,
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  ComCtrls,
  HTB97,
{$IFNDEF EAGLCLIENT}
  db,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  HDB,
  EdtREtat,
{$ELSE}
  HQry, UtileAgl,
{$ENDIF}
  Grids,
  HCtrls,
  HEnt1,
  Ent1,
  EntPaie,
  HMsgBox,
{$IFNDEF EAGLSERVER}
  UTOF,
{$ENDIF}
  UTOB,
  SaisUtil,
  UtilSais,
  //      UTOM,
  Vierge,
  AGLInit,
  UtilPgi,
  cbpIniFiles,
  Paramsoc,
  UDossierSelect,
  Galoutil,
  Galsystem,
  shellapi,
  EntDP,
  //  PGVisuObjet,
  //  PT18 : 04/06/02   : V582  PH Récupération et personnalisation de la focntion insertiondifféree
  //      FactCpta, donc suppression du uses de FactCpta
  ed_tools,
  PGRepertoire;

type
  T_ErrCpta = (rcOk, rcRef, rcPar);

type
  T_AnalPaie = class
    Cpte, Etab, Auxi, NumP, NumL: string; // compte,etablissement,Auxiliaire,numero de piece, numero de ligne
    Anal: TList; // liste des decompositions analytiques valorisees
    constructor Create;
    destructor Destroy; override; // V_42 Qualité
  end;

type
  T_DetAnalPaie = class // Ligne detail section/axe
    Section, Ax: string;
    DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double; // Montants Debit/credit dans chaque monnaie
  end;

type
  T_ContreP = class
    Etab, Prem421, PremCpt: string;
  end;

  {
  procedure CompleteCompte (var LeCompte : String); // Complete le numero de compte avec les caracteres de bourrage
  function  LeChampSal(LeTypChamp, TypeQ : String;  Q: TQUERY; TheTob : TOB): STring; // recupère la valeur du champ associé à chaque ligne de bulletin en fction des paramSoc
  }

type
  TOF_PG_GeneCompta = class(TOF)
  private
    LblDD, LblDF, DateEcr: THLabel; // Date debut, fin des paies traitées et par défaut de la date des ecritures paie
    BtnLance: TToolbarButton97;
    JeuPaie, JournalPaie: THValComboBox;
    QMul: TQUERY; // Query recuperee du mul
    Trace, TraceErr: TListBox;
    ErrorGene: TStringList;
    NbDecimale: Integer; // Nbre de décimale de la compta
    // PT53 definition des valeurs des champs dans la TOB
    iE_DEBIT, iE_CREDIT, iE_GENERAL, iE_ETABLISSEMENT, iE_CONTREPARTIGEN, iE_AUXILIAIRE, iE_RIB, iE_ETATLETTRAGE: Integer;
    iE_MODEPAIE, iE_ENCAISSEMENT, iE_CONFIDENTIEL, iE_ECHE, iE_NUMECHE, iE_DATEECHEANCE, iE_DATECOMPTABLE: Integer;
    iE_ANA, iE_DATEPAQUETMIN, iE_DATEPAQUETMAX, IE_NUMEROPIECE, iE_NUMLIGNE, iE_NUMORDRE, iE_JOURNAL, iE_LIBELLE, iE_REFINTERNE: Integer;
    iE_DEBITDEV, iE_CREDITDEV, iE_TYPEMVT, iE_CONTREPARTIEGEN: Integer;
    LeWhere, CritSal: string; //PT60 contenu du where fait par le multicritère pour faire update en fin de traitement
    ModeSaisieJal: string; // Mode de saisie du journal
    Etabl, Method: string; // Etablissement sur lequel on fait les OD de paie
    RefPaie: string;
    TTA: TList; // liste des ventilations (cumuls analytiques)
    Tob_Ventana, Tob_Vensal: TOB; // Tob des ventilations analytiques du salarie
    AnalCompte: TOB; // Tob qui contient les ventilations analytiques pour un compte
    ComptaEuro, PaieEuro: Boolean; // tenue euro des differents modules
    procedure LanceGeneCompta(Sender: TObject);
    procedure InitEcritDefaut(TOBE: TOB); // Initialisation TOB ecriture avec ts les champs remplis par defaut
    procedure LignePaieECR(MM: RMVT; TOBE: TOB); // Rempli une ligne ecriture en fonction de la ligne du bulletin
    procedure RenseigneClefComptaPaie(Etab: string; var MM: RMVT);
    procedure RemplirT_ecr(TypVentil: string; Montant: Double; TypMvt, Libelle: string; T_Ecr, TOB_Generaux: TOB; Compte, CpteAuxi: string; LeFolio,
      LeNumEcrit: Integer); // Rempli une tob paie preligne ecriture comptable avant affectation définitive numéropiece ...
    function RendMontantHistoBul(NatureRub, Sens, TypeSalPat: string; Q: TQuery): Double; // rend le montant de la ligne en fonction du type du guide ecritures
    function AlimEcrCompta(LeLib, LeLib1, Etablissement, Nature, Compte, LeSens, TypeSalPat: string; Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux: TOB; Q: TQuery;
      QuelRacineFolio: Integer; QSal: TQuery; TraceErr: TListBox = nil): Boolean;
    procedure EquilibreEcrPaie(TOBEcr: TOB; MM: RMVT; NbDec: integer);
    function CreerLigneEcartPaie(TOBEcr, TOBL: TOB; MM: RMVT; EcartD, EcartP, EcartE: Double): T_ErrCpta;
    procedure CumulAnalPaie(TypVentil, Cpte, Etab, Auxi, NumP, NumL: string; DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double; AVentil: array of Boolean);
    //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques format paie PGI alors on conserve la TOB TOBAna
    procedure EcrPaieVersAna(TOB_Generaux, TOBEcr, TOBAna: TOB; OldP, OldL, NewP, NewL: Integer; Conservation: Boolean = FALSE);
    procedure PaieGenereAttente(MontantAxe, MontantAxeEuro: array of double; TOB_Generaux, TOBEcr, TOBAna: TOB; TotalEcriture, TotalDevise, TotalEuro: Double;
      Conservation: Boolean = FALSE);
    procedure ReecritPaiesCompta;
    //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
    function RendLongueurRacine(LeCompte: string): Integer;
    procedure AnalyseAnalytique(MaMethode: string; TOBAna: TOB; PaieEuro, ComptaEuro: Boolean; NbDec: Integer; Conservation: Boolean; Pan: TPageControl);
    // DEB PT53
    function RendGene(LaTobG: TOB; LeCompte: string): TOB;
    function RendCotisation(LaTobC: TOB; LaCot: string): TOB;
    function RendVentiOrg(LaTobOrg: TOB; org: string): TOB;
    function RendVentirem(LaTobRem: TOB; Rem: string): TOB;
    function RendVenticot(LaTobCot: TOB; Cot: string): TOB;
    function RendTContreP(TContreP: TOB; Etab: string): TOB;
    function RendTiers(LaTobT: TOB; LeTiers: string): TOB;
    procedure SauvLesEcrit(TL, TS: TOB);
    procedure MemoriseNumChamp(TT: TOB);
    // FIN PT53
    //Fonctions et procédure TRA
    procedure ExportPaieTRA;
    procedure BImprimerClik (Sender: TObject);  //PT72
  public
    procedure OnArgument(Arguments: string); override;
  end;

function ConfectionCpte(LaRacine, TypeQ: string; Q: TQUERY; C421: Boolean; TheTob: TOB; ZTraceErr: TListBox = nil): string;
function LeChampSal(LeTypChamp, TypeQ: string; Q: TQUERY; TheTob: TOB): string;
procedure CompleteCompte(var LeCompte: string);
function PGInsertionDifferee(TOBEcr: TOB): boolean;

implementation

uses SaisBul,
  PgCongesPayes,
  P5Util,
  P5Def,
  cbpPath,
  PgOutils,
  PgOutils2;

//  PT18 : 04/06/02   : V582  PH Récupération et personnalisation de la focntion insertiondifféree

function PGInsertionDifferee(TOBEcr: TOB): boolean;
var
  EcrDiff, TOBE: TOB;
  StEcr, LaRef, LaNat, LeUser: string;
  DD: TDateTime;
  LeRang: integer;
  Q: TQuery;
begin
  Result := True;
  if TOBEcr = nil then Exit;
  if TOBEcr.Detail.Count <= 0 then Exit;
  TOBE := TOBEcr.Detail[0];
  EcrDiff := TOB.Create('COMPTADIFFEREE', nil, -1);
  DD := TOBE.GetValue('E_DATECOMPTABLE');
  EcrDiff.PutValue('GCD_DATEPIECE', DD);
  LaRef := 'DIF';
  EcrDiff.PutValue('GCD_REFPIECE', LaRef);
  LaNat := TOBE.GetValue('E_NATUREPIECE');
  EcrDiff.PutValue('GCD_NATURECOMPTA', LaNat);
  LeUser := V_PGI.User;
  EcrDiff.PutValue('GCD_USER', LeUser);
  StEcr := TOBEcr.SaveToBuffer(True, False, '');
  EcrDiff.PutValue('GCD_BLOBECRITURE', StEcr);
  Q := OpenSQL('SELECT MAX(GCD_RANG) FROM COMPTADIFFEREE WHERE GCD_DATEPIECE="' + UsDateTime(DD) + '" AND GCD_REFPIECE="' + LaRef + '" AND GCD_USER="' + LeUser
    + '"', True);
  if not Q.EOF then LeRang := Q.Fields[0].AsInteger + 1 else LeRang := 0;
  Ferme(Q);
  EcrDiff.PutValue('GCD_RANG', LeRang);
  Result := EcrDiff.InsertDB(nil);
  EcrDiff.Free;

end;
// FIN PT18

// Confection du numéro de compte
//   PT19 : 04/06/02   : V582  PH Gestion de la trace en cas d'erreur : Ajout paramètre ZTraceErr et traitement

function ConfectionCpte(LaRacine, TypeQ: string; Q: TQUERY; C421: Boolean; TheTob: TOB; ZTraceErr: TListBox = nil): string;
var
  LeCompte, LePlus: string;
  St, LaZone: string; // PT47
  Adecompo: Boolean; // PT47
begin
  LeCompte := LaRacine;

  if C421 then // cas du compte net à payer
  begin
    LePlus := LeChampSal(VH_Paie.PGPreselect0, TypeQ, Q, TheTob);
    if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGLongRacine421, VH_Paie.PGLongRacine421);
    CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
    result := LeCompte;
    exit;
  end;
  //   PT34 : 24/02/03 Classe 1 ou 9
  if ((Copy(LeCompte, 1, 1) = '6') or (Copy(LeCompte, 1, 1) = '7') or (Copy(LeCompte, 1, 1) = '9')) and not VH_Paie.PgDeCompoCl6 then
  begin // cas non decomposition des comptes de la classe 6
    CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
    result := LeCompte;
    exit;
  end;
  //   PT-9 : 19/02/02   : V571  PH test compte de la classe 2 (remboursement de prets)
  //   PT34 : 24/02/03 Classe 1 ou 9
  if ((Copy(LeCompte, 1, 1) = '4') or (Copy(LeCompte, 1, 1) = '2') or (Copy(LeCompte, 1, 1) = '1')) and not VH_Paie.PgDeCompoCl4 then
  begin // cas non decomposition des comptes de la classe 4
    CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
    result := LeCompte;
    exit;
  end;
  //   PT34 : 24/02/03 Classe 1 ou 9
  if (Copy(LeCompte, 1, 1) = '6') or (Copy(LeCompte, 1, 1) = '7') or (Copy(LeCompte, 1, 1) = '9') then
  begin // DEB PT47
    ADecompo := TRUE;
    if VH_Paie.PGDECOMP6 <> '' then
    begin // Cas de la décomposition de quelques racines saisies au préalable
      ADecompo := FALSE;
      LaZone := VH_Paie.PGDECOMP6;
      St := ReadTokenst(LaZone);
      while st <> '' do
      begin
        if Copy(LeCompte, 1, Strlen(PChar(st))) = st then
        begin
          ADecompo := TRUE;
          break;
        end;
        St := ReadTokenst(LaZone);
      end;
    end;
    if ADecompo then
    begin // FIN PT47
      if VH_Paie.PGNbreRac1 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPreselect1, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbreRac1, VH_Paie.PGNbreRac1)
        else if Q <> nil then
          if ZTraceErr <> nil then // DEB PT53-1
          begin
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.getvalue('PHB_SALARIE'));
          end; // FIN PT53-1
      end;
      if VH_Paie.PGNbreRac2 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPreselect2, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbreRac2, VH_Paie.PGNbreRac2)
        else if Q <> nil then
          if ZTraceErr <> nil then
          begin // DEB PT53-1
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.GetValue('PHB_SALARIE'));
          end; // FIN PT53-1
      end;
      if VH_Paie.PGNbreRac3 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPreselect3, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbreRac3, VH_Paie.PGNbreRac3)
        else if Q <> nil then
          if ZTraceErr <> nil then
          begin // DEB PT53-1
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.GetValue('PHB_SALARIE'));
          end; // FIN PT53-1
      end;
      CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
      result := LeCompte;
      exit;
    end; // PT47
  end;
  //   PT-9 : 19/02/02   : V571  PH test compte de la classe 2 (remboursement de prets)
  //   PT34 : 24/02/03 Classe 1 ou 9
  if (Copy(LeCompte, 1, 1) = '4') or (Copy(LeCompte, 1, 1) = '2') or (Copy(LeCompte, 1, 1) = '1') then
  begin // traitement decomposition des comptes de la classe 4
    ADecompo := TRUE; // DEB PT47
    if VH_Paie.PGDECOMP4 <> '' then
    begin // Cas de la décomposition de quelques racines saisies au préalable
      ADecompo := FALSE;
      LaZone := VH_Paie.PGDECOMP4;
      St := ReadTokenst(LaZone);
      while st <> '' do
      begin
        if Copy(LeCompte, 1, Strlen(PChar(st))) = st then
        begin
          ADecompo := TRUE;
          break;
        end;
        St := ReadTokenst(LaZone);
      end;
    end;
    if ADecompo then
    begin // FIN PT47
      if VH_Paie.PGNbre4Rac1 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPre4select1, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbre4Rac1, VH_Paie.PGNbre4Rac1)
        else if Q <> nil then
          if ZTraceErr <> nil then // DEB PT53-1
          begin
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.GetValue('PHB_SALARIE'));
          end; // FIN PT53-1
      end;
      if VH_Paie.PGNbre4Rac2 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPre4select2, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbre4Rac2, VH_Paie.PGNbre4Rac2)
        else if Q <> nil then
          if ZTraceErr <> nil then
          begin // DEB PT53-1
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.GetValue('PHB_SALARIE'));
          end; // FIN PT53-1
      end;
      if VH_Paie.PGNbre4Rac3 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPre4select3, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbre4Rac3, VH_Paie.PGNbre4Rac3)
        else if Q <> nil then
          if ZTraceErr <> nil then
          begin // DEB PT53-1
            if TheTob = nil then ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString)
            else ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + TheTob.GetValue('PHB_SALARIE').AsString);
            ;
          end; // FIN PT53-1
      end;
      CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
      result := LeCompte;
      exit;
    end; // PT47
  end;
  //   PT39 : 19/03/2004 : V_500 PH Affection de la valeur par defaut correspondant à la racine du compte
  CompleteCompte(LeCompte); // Bourrage à droite du numéro de compte
  result := LeCompte;
end;
// FIN PT19
// Bourrage à droite du numéro de compte par rapport à la longueur des comptes de la compta

procedure CompleteCompte(var LeCompte: string);
var
  Lg, ll, i: Integer;
  Bourre: string;
begin
  Bourre := VH^.Cpta[fbGene].Cb; // Recup caractère de bourrage des comptes généraux
  ll := Length(LeCompte);
  Lg := VH^.Cpta[fbGene].Lg;
  if ll < Lg then
  begin
    for i := 1 to (Lg - ll) do LeCompte := LeCompte + Bourre;
  end
  else
    if ll > Lg then LeCompte := Copy(LeCompte, 1, Lg);
  // Si on a une longueur de compte > à la longueur maxi dans la compta alors on limite BOF mais mieux que rien car evite ACCESS VIO

end;

// Composition du compte général en fonction du paramètrage

function LeChampSal(LeTypChamp, TypeQ: string; Q: TQUERY; TheTob: TOB): string;
var
  LeChamp, Prefixe: string;
  LaDate: TdateTime;
  JJ, MM, AA: WORD;
  LeMois: string;
begin
  Result := '';
  LeChamp := '';
  // DEB PT55 Prise en compte des provisions CP et RTT - Rajout accès table PROVI
  if TypeQ = 'P' then Prefixe := 'PPU'
  else if TypeQ = 'H' then Prefixe := 'PHB'
  else if TypeQ = 'D' then Prefixe := 'PDC'
  else exit;
  // FIN PT55
 //   PT24 : 23/08/02   : V585  PH Confection du numéro de compte en fonction du mois de paie
  if LeTypChamp = 'DAT' then
  begin
    // DEB PT53-1
    if TheTob <> nil then LaDate := TheTob.GetValue(Prefixe + '_DATEFIN')
    else if Q <> nil then LaDate := Q.FindField(Prefixe + '_DATEFIN').AsDateTime
    else LaDate := IDate1900;
    // FIN PT53-1
    if LaDate > 0 then
    begin
      DecodeDate(LaDate, AA, MM, JJ);
      LeMois := '0' + IntToStr(MM);
      if Length(LeMois) < 3 then LeMois := '0' + LeMois;
      result := LeMois;
    end;
    exit;
  end;
  // FIN PT24
  if LeTypChamp = 'ETB' then Lechamp := Prefixe + '_ETABLISSEMENT';
  if LeTypChamp = 'OR1' then Lechamp := Prefixe + '_TRAVAILN1';
  if LeTypChamp = 'OR2' then Lechamp := Prefixe + '_TRAVAILN2';
  if LeTypChamp = 'OR3' then Lechamp := Prefixe + '_TRAVAILN3';
  if LeTypChamp = 'OR4' then Lechamp := Prefixe + '_TRAVAILN4';
  if LeTypChamp = 'STA' then Lechamp := Prefixe + '_CODESTAT';
  if LeTypChamp = 'TC1' then Lechamp := Prefixe + '_LIBREPCMB1';
  if LeTypChamp = 'TC2' then Lechamp := Prefixe + '_LIBREPCMB2';
  if LeTypChamp = 'TC3' then Lechamp := Prefixe + '_LIBREPCMB3';
  if LeTypChamp = 'TC4' then Lechamp := Prefixe + '_LIBREPCMB4';

  if LeChamp <> '' then
  begin
    // DEB PT53-1
    if TheTob <> nil then result := TheTob.GetValue(LeChamp)
    else if Q <> nil then Result := Q.FindField(Lechamp).AsString;
    // FIN PT53-1
  end;
end;
// Fonction principale de traitement de la génération des ecritures comptables

procedure TOF_PG_GeneCompta.LanceGeneCompta(Sender: TObject);
var
  Salarie, Nature, Organisme, Nodossier, Rubrique, TypeSalPat, Metab, MDR_Def, TheAuxi: string;
  Pan: TPageControl;
  Tbsht: TTabSheet;
  QuelRacine, LeFolio, LeFolio1, LeFolio2, LeNumEcrit, LeNumEcrit1, LeNumEcrit2, i, OkOk: Integer;
  //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
  ConservAna, EcrPositif: Boolean;
  DateS, ZZDate1, ZZDate2: TDateTime;
  SystemTime0: TSystemTime;
  valid, Compte421, Compte641, Col421, Col42x, OkTrouve, OkTrouve6, Lettrable, LettrableAux, ASauter: boolean;
  St, St1, TheRacine, Racine, Compte, Compte6, LeCompte, LeSens, CpteAuxi, sRib: string;
  Tob_GuideEcr, Tob_VentiOrg, Tob_VentiRem, Tob_VentiCot, Tob_HistoBulletin: TOB; // TOB correspondant à des tablers
  // PT53 Nouvelles TOB pour optimisation des traitements
  Tob_VentiOrgSTD, Tob_VentiRemSTD, Tob_VentiCotSTD, Tob_VentiOrgCEG, Tob_VentiRemCEG, Tob_VentiCotCEG: TOB;
  Torg, Tob_LesEcrit, T_Ecr, TPR, Tob_Cotisations, TOB_Generaux, Tob_Tiers, TTiers: TOB; // Tob de  travail;
  TEcritPaie, TOD, TGene, TContreP, TCP, TOBANA, TZZ, T42, MaTOB, Ta, TModeRegl, TRegl: TOB;
  T_Etab, Tetb, T_PAIES, TPPU, TSauvLesEcrit: TOB; // PT38 Tob des etablissements
  Q, QJal, QSal: TQuery;
  SalBrut, SalNet, SalCharges, Ecart, Ecart1, TotDebit, TotCredit, TotDebit2, TotCredit2, Montant: double;
  // Montants debit/credit folio 1 et 2 MtD1, MtD2, MtC1, MtC2 : Double
  MM: RMVT;
  NumEcritOD1, NumEcritOD2, NumFolioOD1, NumFolioOD2, NBrePiece, NbDec: Integer; // Renumerotation des numeros de lignes dans le Folio
  PremCpt, Prem421, IntegODPaie: string; // 1er Compte et 1er compte 421 trouvés pour la gestion des contre parties
  ContreP: T_ContreP; // Gestion des contreparties ==> pour memoriser lres 1er compte et 1er compte 421
  OldP, OldL, NewP, NewL, rep: Integer;
  TheTrace: TStringList;
  LgRac, ij, rs, NbrePaie, jj: Integer;
  Aleat: Double;
  TheModRegl, NatGene, GLet, GPoint, LeLib, LeLib1, TypDot: string;
  Sts: string;
//*TRA
  bOk : Boolean;
begin
//*TRA
  bOk := False;
  Salarie := '';
  Etab := '';
  NBrePiece := 1;
  PremCpt := '';
  Prem421 := '';
  RefPaie := 'Paie du ' + LblDD.Caption + ' au ' + LblDF.Caption;
  // On regarde quelles sont les monnaies de tenue de la paie et de la compta
  PaieEuro := VH_Paie.PGTenueEuro; // Monnaie de tenue de la paie
  ComptaEuro := VH^.TenueEuro; // Monnaie de tenue de la compta
  //   PT35 : 25/02/03  : V_500 PH On ne tient plus compte de la monnaie de tenue de la compta
  if not ComptaEuro and PaieEuro then ComptaEuro := PaieEuro;
  IntegODPaie := VH_Paie.PGIntegODPaie; // Méthode d'intégration de la paie
  NbDecimale := V_PGI.OkdecV;
  //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
  if (IntegODPaie = 'ECP') then //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  begin
    ConservAna := TRUE;
  end
  else ConservAna := FALSE;

  // Recupere le nombre de décimale de la compta si non renseigné prend 2 par defaut
  if NbDecimale = 0 then NbDecimale := 2;
  if (VH_Paie.PGLongRacine = 0) or (VH_Paie.PGLongRacine421 = 0) or (VH_Paie.PGLongRacin4 = 0) then
  begin
    PGIBox('Renseignez toutes les longueurs des racines de vos comptes !', 'Modifiez les paramètres société');
    exit;
  end;
  if IntegODPaie = '' then
  begin
    PGIBox('La génération ne peut pas être lancée', 'Vous n''avez pas indiqué de méthode de comptabilisation dans les paramètres société !');
    exit;
  end;
  if IntegODPaie <> 'ECP' then
  begin
    //  PT10 : 22/03/02   : V575  PH activation passege des ODs en différe dans la compta PGI
    if (IntegODPaie <> 'IMM') and (IntegODPaie <> 'DIF') and (IntegODPaie <> 'TRA') then
    begin
      PGIBox('La génération ne peut pas être lancée', 'La méthode de comptabilisation dans les paramètres société n''est pas validée !');
      exit;
    end;
  end;
  //   PT-7 : 11/12/01   : V563  PH test si les compteurs des journaux existent
  if not VH_Paie.PGTypeEcriture then Q := OpenSQL('Select J_COMPTEURNORMAL from JOURNAL Where J_JOURNAL="' + JournalPaie.Value + '"', True)
  else Q := OpenSQL('Select J_COMPTEURSIMUL from JOURNAL Where J_JOURNAL="' + JournalPaie.Value + '"', True);
  if (not Q.EOF) or (VH_Paie.PGIntegODPaie = 'ECP') then ferme(Q)
  else
  begin
    Ferme(Q);
    rep := PgiAsk('Attention, vos compteurs de journaux n''existent pas ?#13#10 Voulez vous continuer ?', 'Génération des ODs de Paie');
    if rep <> mrYes then exit;
  end;
  Q := OpenSQL('SELECT MR_MODEREGLE,MR_MP1,MR_TAUX1 FROM MODEREGL ORDER BY MR_MODEREGLE', True);
  if not Q.EOF then
  begin
    Q.First;
    MDR_Def := Q.FindField('MR_MP1').AsString;
  end;
  TModeRegl := TOB.Create('Les ModedePaiement', nil, -1);
  TModeRegl.LoadDetailDB('MODEREGL', '', '', Q, False);
  ferme(Q);
  if MDR_Def = '' then
  begin
    rep := PgiAsk('Attention, vous n''avez pas de mode de réglement dans la comptabilité?#13#10 Voulez vous continuer ?', 'Génération des ODs de Paie');
    if rep <> mrYes then exit;
  end;

  // PT38
  st := 'SELECT ETB_ETABLISSEMENT,ETB_RIBSALAIRE FROM ETABCOMPL';
{Flux optimisé
  Q := OpenSql(St, TRUE);
  T_Etab := TOB.Create('Mes etablissements', nil, -1);
  T_Etab.LoadDetailDB('ETABSOCIAL', '', '', Q, False);
  ferme(Q);
}
  T_Etab := TOB.Create('Mes etablissements', nil, -1);
  T_Etab.LoadDetailDBFROMSQL('ETABCOMPL', st);

  // FIN PT38

    // On regarde si le compte 421000 est un compte collectif pour lui mettre un auxiliaire
  Col421 := FALSE;
  St :=
    'SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_CENTRALISABLE,G_COLLECTIF,G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_VENTILABLE,G_LETTRABLE,G_POINTABLE,G_MODEREGLE,G_CONFIDENTIEL FROM GENERAUX ORDER BY G_GENERAL'; // PT52
{Flux optimisé
  Q := OpenSql(st, TRUE);
  // Chargement de la TOB des comptes généraux pour avoir les caractéristiques de chaque compte général
  TOB_Generaux := TOB.Create('Les comptes généraux', nil, -1);
  TOB_Generaux.LoadDetailDB('GENERAUX', '', '', Q, False);
  ferme(Q);
}
  TOB_Generaux := TOB.Create('Les comptes généraux', nil, -1);
  TOB_Generaux.LoadDetailDBFROMSQL('GENERAUX', st);

  //  TPR := TOB_Generaux.FindFirst(['G_GENERAL'], [VH_Paie.PGCptNetAPayer], FALSE);
  // PT65 Si la longueur du compte net à payer saisi > longueur des comptes ds la compta alors on prend la longueur de la compta
  if (Length(VH_Paie.PGCptNetAPayer) <> VH^.Cpta[fbGene].Lg) then
   CompleteCompte (VH_Paie.PGCptNetAPayer); // PT65-1

  TPR := RendGene(TOB_Generaux, VH_Paie.PGCptNetAPayer);
  //   PT-5 : 03/12/01   : V563  PH Correction pb compte 421 Collectif test inversé
  if TPR <> nil then Col421 := not (TPR.GetValue('G_COLLECTIF') = 'X');
  st := VH_Paie.PGCptNetAPayer;
  if not Col421 then st := st + ' qui est collectif et avec auxiliaire ' else st := st + ' qui n''est pas collectif et sans auxiliaire';
  rep := PgiAsk('Attention, vous avez un compte Net à Payer ' + st + '#13#10 Voulez vous continuer ?', 'Génération des ODs de Paie');
  if rep <> mrYes then
  begin
    TOB_Generaux.free;
    TOB_Generaux := nil;
    exit;
  end;
  TTA := TList.Create; // Création de la liste des ventilations analytiques des écritures
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  Nodossier := PgRendNoDossier();
  Pan := TPageControl(GetControl('PANELPREP'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La génération ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
  if ((JournalPaie = nil) or (JournalPaie.Value = '')) and (VH_Paie.PGIntegODPaie <> 'ECP') then
  begin
    PGIBox('La génération ne peut pas être lancée', 'car vous n''avez pas sélectionné de journal');
    exit;
  end;
  Trace.Items.Add('Chargement des informations nécessaires ');

  QJal := OpenSQL('SELECT J_MODESAISIE FROM JOURNAL WHERE J_JOURNAL="' + JournalPaie.Value + '"', True);
  if not QJal.EOF then ModeSaisieJal := QJal.Fields[0].AsString
  else ModeSaisieJal := '-';
  Ferme(QJal);
  //   PT22 : 13/08/02   : V585  PH Test si ecriture de simulation et journal pas de type piece alors erreur
  if VH_PAie.PgTypeEcriture then // Ecriture de simulation
  begin
    if ModeSaisieJal <> '-' then // si journal <> piece alors erreur
    begin
      PGIBox('La génération ne peut pas être lancée car vous ne pouvez pas avoir #13#10 des écritures de simulation avec un journal de type bordereau ou libre', ECran.Caption);
      exit;
    end;
  end;
  // FIN PT22
  if (JeuPaie = nil) or (JeuPaie.Value = '') then
  begin
    PGIBox('La génération ne peut pas être lancée #13#10car vous n''avez pas sélectionné de modèle d''écritures', Ecran.caption); // PT61
    exit;
  end;
  St := 'SELECT * FROM JEUECRPAIE ' +
    'WHERE ##PJP_PREDEFINI## PJP_NOJEU =' + JeuPaie.Value + ''; // DB2
  Q := OpenSql(st, TRUE);
  if not Q.EOF then TypDot := Q.FindField('PJP_TYPEPROVCP').AsString;
  FERME(Q);
  //   PT41 : 31/03/2004 : V_500 PH compte du type du modèle qui corrrespond à un journal ODs et non de provisions
  if TypDot <> 'COD' then // PT48
  begin
    PGIBox('Le modèle d''écritures ne correspond à une écriture de type ODs de paie', Ecran.Caption);
    exit;
  end;
  // FIN PT41
  // DEB PT62
  if (not ChargeMagExo(True)) AND (VH_Paie.PGIntegODPaie = 'IMM') then // PT67
  begin
    PGIError('Erreur sur le chargement des paramètres de la comptabilité', Ecran.Caption);
    exit;
  end;
  // FIN PT62
  BtnLance.Enabled := FALSE;
  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  DateD := StrToDate(LblDD.Caption);
  DateF := StrToDate(LblDF.Caption);
  // Chargement des TOB pour la génération des ecritures
  Tob_GuideEcr := TOB.Create('Le Guide des ecritures', nil, -1);
  St := 'SELECT * FROM GUIDEECRPAIE WHERE ##PGC_PREDEFINI## PGC_JEUECR=' + // DB2
    JeuPaie.Value +
    ' ORDER BY PGC_PREDEFINI,PGC_NODOSSIER,PGC_JEUECR,PGC_NUMFOLIO,PGC_NUMECRIT'; // DB2
  Q := OpenSql(st, TRUE);
  Tob_GuideEcr.LoadDetailDB('GUIDEECRPAIE', '', '', Q, False);
  Ferme(Q);
  // Recherche si les comptes 421 et/ou 641 sont utilisés dans le guide
  Compte421 := FALSE;
  Compte641 := FALSE;
  for I := 0 to Tob_GuideEcr.Detail.count - 1 do
  begin
    TPR := Tob_GuideEcr.Detail[I];
    LeCompte := TPR.GetValue('PGC_GENERAL');
    //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
    if (Copy(LeCompte, 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TPR.GetValue('PGC_ALIM421') <> '') then Compte421 := TRUE;
    if (Copy(LeCompte, 1, 3) = '641') and (TPR.GetValue('PGC_ALIM421') <> '') then Compte641 := TRUE;
  end;
  // Chargement TOB Cotisations
  Tob_Cotisations := TOB.Create('Les cotisations', nil, -1);
  St := 'SELECT * FROM COTISATION WHERE PCT_NATURERUB="COT" AND ##PCT_PREDEFINI##  ORDER BY PCT_RUBRIQUE'; //**//
{Flux optimisé
  Q := OpenSql(st, TRUE);
  Tob_Cotisations.LoadDetailDB('COTISATION', '', '', Q, False);
  Ferme(Q);
}
  Tob_Cotisations.LoadDetailDBFROMSQL('COTISATION', st);

  Tob_LesEcrit := TOB.Create('Les ecritures vues de la paie', nil, -1);
  // Chargement des TOB des differents types de ventilations PT53
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="DOS" AND PVO_NODOSSIER="' + PGRendNoDossier + '" ORDER BY PVO_TYPORGANISME', TRUE);
  Tob_VentiOrg := TOB.Create('Les ventilations des organismes', nil, -1);
  Tob_VentiOrg.LoadDetailDB('VENTIORGPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st :='SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="DOS" AND PVO_NODOSSIER="' + PGRendNoDossier + '" ORDER BY PVO_TYPORGANISME';
  Tob_VentiOrg := TOB.Create('Les ventilations des organismes', nil, -1);
  Tob_VentiOrg.LoadDetailDBFROMSQL('VENTIORGPAIE', st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="STD" ORDER BY PVO_TYPORGANISME', TRUE);
  Tob_VentiOrgSTD := TOB.Create('Les ventilations des organismes STD', nil, -1);
  Tob_VentiOrgSTD.LoadDetailDB('VENTIORGPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st :='SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="STD" ORDER BY PVO_TYPORGANISME';
  Tob_VentiOrgSTD := TOB.Create('Les ventilations des organismes STD', nil, -1);
  Tob_VentiOrgSTD.LoadDetailDBFROMSQL('VENTIORGPAIE', st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="CEG" ORDER BY PVO_TYPORGANISME', TRUE);
  Tob_VentiOrgCEG := TOB.Create('Les ventilations des organismes CEG', nil, -1);
  Tob_VentiOrgCEG.LoadDetailDB('VENTIORGPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT * FROM VENTIORGPAIE WHERE PVO_PREDEFINI="CEG" ORDER BY PVO_TYPORGANISME';
  Tob_VentiOrgCEG := TOB.Create('Les ventilations des organismes CEG', nil, -1);
  Tob_VentiOrgCEG.LoadDetailDBFROMSQL('VENTIORGPAIE', st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="DOS" AND PVS_NODOSSIER="' + PGRendNoDossier + '"  ORDER BY PVS_RUBRIQUE', TRUE);
  Tob_VentiRem := TOB.Create('Les ventilations des rémunérations', nil, -1);
  Tob_VentiRem.LoadDetailDB('VENTIREMPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="DOS" AND PVS_NODOSSIER="' + PGRendNoDossier + '"  ORDER BY PVS_RUBRIQUE';
  Tob_VentiRem := TOB.Create('Les ventilations des rémunérations', nil, -1);
  Tob_VentiRem.LoadDetailDBFROMSQL('VENTIREMPAIE',St);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="STD" ORDER BY PVS_RUBRIQUE', TRUE);
  Tob_VentiRemSTD := TOB.Create('Les ventilations des rémunérations STD', nil, -1);
  Tob_VentiRemSTD.LoadDetailDB('VENTIREMPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="STD" ORDER BY PVS_RUBRIQUE';
  Tob_VentiRemSTD := TOB.Create('Les ventilations des rémunérations STD', nil, -1);
  Tob_VentiRemSTD.LoadDetailDBFROMSQL('VENTIREMPAIE',st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="CEG" ORDER BY PVS_RUBRIQUE', TRUE);
  Tob_VentiRemCEG := TOB.Create('Les ventilations des rémunérations CEG', nil, -1);
  Tob_VentiRemCEG.LoadDetailDB('VENTIREMPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT * FROM VENTIREMPAIE WHERE PVS_PREDEFINI="CEG" ORDER BY PVS_RUBRIQUE';
  Tob_VentiRemCEG := TOB.Create('Les ventilations des rémunérations CEG', nil, -1);
  Tob_VentiRemCEG.LoadDetailDBFROMSQL('VENTIREMPAIE',st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="DOS" AND PVT_NODOSSIER="' + PGRendNoDossier + '"  ORDER BY PVT_RUBRIQUE', TRUE);
  Tob_VentiCot := TOB.Create('Les ventilations des cotisations', nil, -1);
  Tob_VentiCot.LoadDetailDB('VENTICOTPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st :='SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="DOS" AND PVT_NODOSSIER="' + PGRendNoDossier + '"  ORDER BY PVT_RUBRIQUE';
  Tob_VentiCot := TOB.Create('Les ventilations des cotisations', nil, -1);
  Tob_VentiCot.LoadDetailDBFROMSQL('VENTICOTPAIE', st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="STD" ORDER BY PVT_RUBRIQUE', TRUE);
  Tob_VentiCotSTD := TOB.Create('Les ventilations des cotisations STD', nil, -1);
  Tob_VentiCotSTD.LoadDetailDB('VENTICOTPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="STD" ORDER BY PVT_RUBRIQUE';
  Tob_VentiCotSTD := TOB.Create('Les ventilations des cotisations STD', nil, -1);
  Tob_VentiCotSTD.LoadDetailDBFROMSQL('VENTICOTPAIE',st);
{Flux optimisé
  Q := OpenSql('SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="CEG" ORDER BY PVT_RUBRIQUE', TRUE);
  Tob_VentiCotCEG := TOB.Create('Les ventilations des cotisations CEG', nil, -1);
  Tob_VentiCotCEG.LoadDetailDB('VENTICOTPAIE', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st :='SELECT * FROM VENTICOTPAIE WHERE PVT_PREDEFINI="CEG" ORDER BY PVT_RUBRIQUE';
  Tob_VentiCotCEG := TOB.Create('Les ventilations des cotisations CEG', nil, -1);
  Tob_VentiCotCEG.LoadDetailDBFROMSQL('VENTICOTPAIE', st);

  // PT 56-1 Modif SQL pour récup des nouveaux champs utiles de la table tiers
{Flux optimisé
  Q := OpenSQL('SELECT T_AUXILIAIRE,T_LETTRABLE,T_MODEREGLE,T_CONFIDENTIEL,T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="SAL" ORDER BY T_AUXILIAIRE', TRUE);
  Tob_Tiers := TOB.Create('Les tiers salaries', nil, -1);
  Tob_Tiers.LoadDetailDB('TIERS', '', '', Q, FALSE, False);
  Ferme(Q);
}
  st := 'SELECT T_AUXILIAIRE,T_LETTRABLE,T_MODEREGLE,T_CONFIDENTIEL,T_LIBELLE FROM TIERS WHERE T_NATUREAUXI="SAL" ORDER BY T_AUXILIAIRE';
  Tob_Tiers := TOB.Create('Les tiers salaries', nil, -1);
  Tob_Tiers.LoadDetailDBFROMSQL('TIERS', St);

  // FIN PT53
  Trace.Items.Add('Fin de Chargement des informations nécessaires');
  if QMul = nil then
  begin
    PGIError('Erreur sur la liste des salariés', 'Abandon du traitement');
    exit;
  end;
  QMul.First;
  // if Method = 'MUL' then MEtab := QMul.FindField ('PPU_ETABLISSEMENT').AsString;
  GetLocalTime(SystemTime0);
  Trace.Items.Add('Début du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  Trace.Refresh;
  // DEB PT53 Chargement en TOB au lieu de naviguer sur la query
  NbrePaie := 0;
  StS := QMUL.SQL[0];
  ij := POS('WHERE', Sts);
  if ij <> 0 then
  begin
    st := Copy(Sts, 1, ij - 1);
    st := FindEtReplace(st, '''', '"', TRUE);
    Sts := st + ' ' + LeWhere + ' ORDER BY PPU_ETABLISSEMENT,PSA_SALARIE'
  end;
{Flux optimisé
  Q := OpenSql(StS, TRUE);
  T_PAIES := TOB.Create('Les paies A traiter', nil, -1);
  T_PAIES.LoadDetailDB('_PAIEENCOURS_', '', '', Q, FALSE, False);
  Ferme(Q);
}
  T_PAIES := TOB.Create('Les paies A traiter', nil, -1);
  T_PAIES.LoadDetailDBFROMSQL('_PAIEENCOURS_', Sts);

  //  while not QMul.EOF do ==> remplacement Findfirst par une boucle for
  for jj := 0 to T_PAIES.detail.count - 1 do
  begin
    TPPU := T_PAIES.Detail[JJ];
    Salarie := TPPU.GetValue('PSA_SALARIE');
    // FIN PT53
    St := 'Select PSA_SALARIE,PSA_AUXILIAIRE,PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
    QSal := OpenSql(St, TRUE);
    if QSal.EOF then
    begin
      TraceErr.Items.Add('Erreur le salarié ' + Salarie + ' est inconnu');
      //      QMul.Next;
      Ferme(QSal);
      continue;
    end;

    Etab := TPPU.GetValue('PPU_ETABLISSEMENT');
    if NbrePaie = 0 then MEtab := Etab;
    if (Etab <> Etabl) and (Etabl <> '') and (Method = 'MON') then // cas Mono établissement
    begin // cas ou le salarié n'appartient pas au moment de la paie à l'établissement selectionné
      //      QMul.Next;
      continue;
    end;
    //  if (Etab <> Metab) AND (Method = 'MUL') then ;// Multi Etablissement et rupture sur code Etablissement
// DEB PT53
    NbrePaie := Nbrepaie + 1; // deb PT53
    if Method = 'PRI' then Etab := VH^.EtablisDefaut;
    if (Etab <> MEtab) and (Method <> 'PRI') then
    begin // rupture code établissement ==>sauvegarde de la TOB.
      if TSauvLesEcrit = nil then TSauvLesEcrit := TOB.Create('Les ecritures vues de la paie', nil, -1);
      SauvLesEcrit(Tob_LesEcrit, TSauvLesEcrit);
      Metab := Etab;
    end;
    // FIN PT53
    ZZDate1 := TPPU.GetValue('PPU_DATEDEBUT');
    ZZDate2 := TPPU.GetValue('PPU_DATEFIN');
    if (VH_Paie.PGAnalytique = TRUE) then
    begin // chargement des ventilations analytiques au niveau de chaque salarié/bulletin
      if Tob_Ventana <> nil then
      begin
        Tob_Ventana.Free;
        Tob_Ventana := nil;
      end;
      Tob_Ventana := PreChargeVentileLignePaie(Salarie, ZZDate1, ZZDate2);
      if (Tob_Ventana = nil) or (TOB_Ventana.Detail.Count <= 0) then // pas de ventilation pour le salarie alors que analytique géré
      begin
        TraceErr.Items.Add('Erreur le salarié ' + Salarie + ' n''a pas de préventilation analytique');
      end;
      if TOB_VenSal <> nil then
      begin
        TOB_VenSal.Free;
        TOB_VenSal := nil;
      end;
      TOB_VenSal := TOB.Create('Analytique Salarié', nil, -1);
      St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + Salarie + '" ORDER BY V_NATURE,V_COMPTE';
{Flux optimisé
      Q := OpenSql(st, TRUE);
      TOB_VenSal.LoadDetailDB('VENTIL', '', '', Q, FALSE);
      ferme(Q);
}
      TOB_VenSal.LoadDetailDBFROMSQL('VENTIL', st);

    end;
    // recuperation des cumuls pour alimentation du compte 421 en fonction du paramètrage des ecritures
    SalBrut := TPPU.GetValue('PPU_CBRUT');
    SalNet := TPPU.GetValue('PPU_CNETAPAYER');
    SalCharges := TPPU.GetValue('PPU_CCOUTSALARIE');
    // Affectation des lignes Ecritures pour les comptes 421 et 641x alimentés par Brut,Net,Charges Salariales
    if (Compte421 = TRUE) or (Compte641 = TRUE) then // cas du net à payer, il faut prévoir le compte auxiliaire du salarié
    begin
      for I := 0 to Tob_GuideEcr.Detail.count - 1 do
      begin
        TPR := Tob_GuideEcr.Detail[I];
        LeCompte := TPR.GetValue('PGC_GENERAL');
        //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
        if (MemCmpPaie(LeCompte, Copy(VH_Paie.PGCptNetAPayer, 1, 3), 1, 3) and Compte421 = TRUE) or
          (MemCmpPaie(LeCompte, '641', 1, 3) and Compte641 = TRUE) or (MemCmpPaie(LeCompte, '42', 1, 2)) or
          //   PT23 : 23/08/02   : V585  PH Traitement des comptes classe 5 pour jnal de paiement : rajout test
        (not MemCmpPaie(LeCompte, '43', 1, 2) and (TPR.GetVAlue('PGC_ALIM421') <> '')) then // test sur la racine compte associé à rem ou organisme
        begin
          Col42x := TRUE;
          LeFolio := StrToInt(TPR.GetValue('PGC_NUMFOLIO')); // ==> E_NUMEROPIECE à incrementer en fonction de la compta
          LeNumEcrit := TPR.GetValue('PGC_NUMECRIT'); // ==> E_NUMLIGNE = Numéro de la ligne dans la piece
          // Dans le cas Non Collectif alors une ecriture centralisée sinon une écriture 421 par salarié avec remplissage du compte auxliaire
          //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
          //     if (MemCmpPaie (LeCompte,Copy (VH_Paie.PGCptNetAPayer,1,3),1, 3)) then // beurk = longueur race 421
          if (MemCmpPaie(LeCompte, Copy(VH_Paie.PGCptNetAPayer, 1, VH_Paie.PGLongRacine421), 1, VH_Paie.PGLongRacine421)) then // $$$$
          begin
            TheRacine := Copy(LeCompte, 1, VH_Paie.PGLongRacine421);
            //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
            Compte := ConfectionCpte(TheRacine, 'P', QMul, TRUE, TPPU, TraceErr); // PT53-1
            if not Col421 then
            begin
              TheAuxi := QSal.FindField('PSA_AUXILIAIRE').AsString;
              if TheAuxi = '' then TraceErr.Items.Add('Le salarié ' + Salarie + ' n''a pas de compte auxiliaire');
              T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_ETABLISSEMENT', 'E_NUMEROPIECE', 'E_NUMLIGNE', 'E_AUXILIAIRE'], [Compte, Etab, LeFolio,
                LeNumEcrit, QSal.FindField('PSA_AUXILIAIRE').AsString], FALSE)
            end
            else T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_ETABLISSEMENT', 'E_NUMEROPIECE', 'E_NUMLIGNE'], [Compte, Etab, LeFolio, LeNumEcrit], FALSE);
          end
          else
          begin
            //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
            LgRac := RendLongueurRacine(LeCompte);
            TheRacine := Copy(LeCompte, 1, LgRac); // $$$$
            // PT38
            if (Copy(TheRacine, 1, 3) = '512') then
            begin
              Tetb := T_Etab.FindFirst(['ETB_ETABLISSEMENT'], [TPPU.GetValue('PPU_ETABLISSEMENT')], FALSE);
              if Tetb <> nil then
                if (Tetb.GetValue('ETB_RIBSALAIRE') <> '') then TheRacine := Tetb.GetValue('ETB_RIBSALAIRE')
            end;
            // FIN PT38
                        //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
            Compte := ConfectionCpte(TheRacine, 'P', QMul, FALSE, TPPU, TraceErr); // PT53-1
            //  PT23 : 23/08/02   : V585  PH Traitement des comptes classe 5 pour jnal de paiement : rajout findfirst
            T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_ETABLISSEMENT', 'E_NUMEROPIECE', 'E_NUMLIGNE'], [Compte, Etab, LeFolio, LeNumEcrit], FALSE);
          end;
          if T_Ecr = nil then
          begin // si ligne ecriture inexistante alors creation
            T_Ecr := TOB.Create('ECRITURE', Tob_LesEcrit, -1);
            InitEcritDefaut(T_Ecr);
            RenseigneClefComptaPaie(Etab, MM);
            LignePaieEcr(MM, T_Ecr);
          end;
          Nature := TPR.GetVAlue('PGC_ALIM421');
          if Nature = 'BRU' then Montant := SalBrut
          else if Nature = 'NET' then Montant := SalNet
          else if Nature = 'SAL' then Montant := SalCharges
          else Montant := 0; // Et Oui Ph
          CpteAuxi := '';
          //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
          if (MemCmpPaie(LeCompte, Copy(VH_Paie.PGCptNetAPayer, 1, 3), 1, 3)) and not COL421 then CpteAuxi := QSal.FindField('PSA_AUXILIAIRE').AsString;
          if Nature <> '' then
            RemplirT_Ecr('S', Montant, TPR.GetValue('PGC_TYPMVT'), TPR.GetValue('PGC_LIBELLE'), T_Ecr, TOB_Generaux, Compte, CpteAuxi, LeFolio, LeNumEcrit);
        end; // Fin si compte 421
      end; // Fin de la boucle sur le guide des ecritures
    end; // Fin traitement 421 cas particulier

    DateD := StrToDate(LblDD.Caption);
    DateF := StrToDate(LblDF.Caption);
    // DateE:= QMul.FindField ('PSA_DATEENTREE').AsDateTime ;

    if Tob_HistoBulletin <> nil then
    begin
      Tob_HistoBulletin.Free;
      Tob_HistoBulletin := nil;
    end;
    Tob_HistoBulletin := Tob.create('Les lignes du bulletin', nil, -1);
    st := 'Select * from histobulletin where PHB_SALARIE="' + Salarie + '" AND PHB_ETABLISSEMENT="' + TPPU.GetValue('PPU_ETABLISSEMENT') + '"' +
      ' AND PHB_DATEDEBUT="' + USDateTime(ZZDate1) + '" AND PHB_DATEFIN="' + USDateTime(ZZDate2) +
      '" ORDER BY PHB_ETABLISSEMENT,PHB_SALARIE,PHB_DATEDEBUT,PHB_DATEFIN,PHB_NATURERUB,PHB_RUBRIQUE';
    Q := OpenSql(st, TRUE);
    while not Q.EOF do // traitement de la liste des rubriques composant le bulletin
    begin
   // Initialisation des montants debit/credit MtD1 := 0;MtD2 := 0;MtC1 := 0;MtC2 := 0;
      Nature := Q.FindField('PHB_NATURERUB').AsString;
      Rubrique := Q.FindField('PHB_RUBRIQUE').AsString;
      if RechCommentaire(Rubrique) = true then
      begin
        Q.Next;
        continue;
      end; // Cas ligne de commentaire non traitée
      if Nature = 'COT' then // 2 cas : Organismes compte de la classe 4 ou Rubrique compte de charge
      begin
        // recherche du compte par rapport aux types d'organismes
//        TOrg := Tob_Cotisations.FindFirst(['PCT_RUBRIQUE'], [Rubrique], FALSE);
        TOrg := RendCotisation(Tob_Cotisations, Rubrique); // PT53
        if TOrg = nil then
        begin
          TraceErr.Items.Add('Cotisation inconnue ' + Rubrique + ' pour le salarie ' + Salarie);
          Q.Next;
          continue;
        end;

        Organisme := Q.FindField('PHB_ORGANISME').AsString;
        //  PT-4 : 22/11/01   : V562  PH ORACLE, Le champ PHB_ORGANISME est obligatoirement renseigné
        if (Organisme = '') or (Organisme = '----') or (Organisme = '....') then
        begin
          Trace.Items.Add('==> Attention, organisme inconnu pour le salarie ' + Salarie + ' pour la cotisation ' + Q.FindField('PHB_RUBRIQUE').AsString); // PT66 Changement de trace
          Q.Next;
          continue;
        end;
        compte := '';
        // Recherche en cascade de la ventilation de l'organisme PT53 dans chaque TOB
        Torg := RendVentiOrg(Tob_VentiOrg, Organisme);
        if Torg = nil then
          Torg := RendVentiOrg(Tob_VentiOrgSTD, Organisme);
        if Torg = nil then
          Torg := RendVentiOrg(Tob_VentiOrgCEG, Organisme);
        if Torg <> nil then
        begin // Cas traitement des organismes comptes concernés de la classe 4
          //   Cas des charges salariales sur les organismes cas FOLIO 1
          //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
          LgRac := RendLongueurRacine(Torg.GetValue('PVO_RACINE1'));
          TheRacine := Copy(Torg.GetValue('PVO_RACINE1'), 1, LgRac); // $$$$
          QuelRacine := 1; // Va correspondre au numéro de folio dans le modèle d'écriture
          LeSens := Torg.GetValue('PVO_SENS1');
          TypeSalPat := Torg.GetValue('PVO_PARTSALPAT1');
          if TheRacine <> '' then
          begin
            //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
            Compte := ConfectionCpte(TheRacine, 'H', Q, FALSE, nil, TraceErr);
            //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
            if Compte <> '' then
            begin
              LeLib := 'La ventilation du folio 1 de l''organisme ' + Torg.GetValue('PVO_TYPORGANISME');
              LeLib1 := ' -->' + RechDom('PGTYPEORGANISME', Torg.GetValue('PVO_TYPORGANISME'), FALSE) + ' de la cotisation ' + Rubrique;
            end
            else
            begin
              LeLib := '';
              LeLib1 := '';
            end;
            AlimEcrCompta(LeLib, LeLib1, Etab, Nature, Compte, LeSens, TypeSalPat, Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux, Q, QuelRacine, QSal);
          end;
          //   Cas des charges patronales sur les organismes cas FOLIO 2
          //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
          LgRac := RendLongueurRacine(Torg.GetValue('PVO_RACINE2'));
          TheRacine := Copy(Torg.GetValue('PVO_RACINE2'), 1, LgRac); // $$$$
          QuelRacine := 2;
          LeSens := Torg.GetValue('PVO_SENS2');
          TypeSalPat := Torg.GetValue('PVO_PARTSALPAT2');
          if TheRacine <> '' then
          begin
            //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
            Compte := ConfectionCpte(TheRacine, 'H', Q, FALSE, nil, TraceErr);
            //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
            if Compte <> '' then
            begin
              LeLib := 'La ventilation du folio 2 de l''organisme ' + Torg.GetValue('PVO_TYPORGANISME');
              LeLib1 := ' --> ' + RechDom('PGTYPEORGANISME', Torg.GetValue('PVO_TYPORGANISME') + ' de la cotisation ' + Rubrique, FALSE);
            end
            else
            begin
              LeLib := '';
              LeLib1 := '';
            end;
            AlimEcrCompta(LeLib, LeLib1, Etab, Nature, Compte, LeSens, TypeSalPat, Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux, Q, QuelRacine, QSal);
          end;
        end;
        // Recherche du compte de charge affecté à la cotisation PT53 dans chaque TOB
        Torg := RendVenticot(Tob_VentiCot, Rubrique);
        if Torg = nil then
          Torg := RendVenticot(Tob_VentiCotSTD, Rubrique);
        if Torg = nil then
          Torg := RendVenticot(Tob_VentiCotCEG, Rubrique);
        if Torg <> nil then
        begin
          //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
          LgRac := RendLongueurRacine(Torg.GetValue('PVT_RACINE1'));
          TheRacine := Copy(Torg.GetValue('PVT_RACINE1'), 1, LgRac); // $$$
          QuelRacine := 1;
          LeSens := Torg.GetValue('PVT_SENS1');
          TypeSalPat := Torg.GetValue('PVT_PARTSALPAT1');
          if TheRacine = '' then
          begin
            //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
            LgRac := RendLongueurRacine(Torg.GetValue('PVT_RACINE2'));
            TheRacine := Copy(Torg.GetValue('PVT_RACINE2'), 1, LgRac); // $$$
            QuelRacine := 2;
            LeSens := Torg.GetValue('PVT_SENS2');
            TypeSalPat := Torg.GetValue('PVT_PARTSALPAT2');
          end;
          //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
          Compte := ConfectionCpte(TheRacine, 'H', Q, FALSE, nil, TraceErr);
          //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
          if Compte <> '' then LeLib := 'La ventilation de la cotisation ' + Rubrique
          else LeLib := '';
          AlimEcrCompta(LeLib, '', Etab, Nature, Compte, LeSens, TypeSalPat, Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux, Q, QuelRacine, QSal);
        end
        else Trace.Items.Add('Ventilation de la cotisation ' + Rubrique + ' inconnue ?');
      end
      else // cas d'une remuneration , on ne prend pas les bases de cotisations
      begin
        if Nature = 'AAA' then
        begin // Rémunération alors on cherche dans la table VENTIREMPAIE PT53 dans chaque TOB
          Torg := RendVentirem(Tob_VentiRem, Rubrique);
          if Torg = nil then
            Torg := RendVentirem(Tob_VentiRemSTD, Rubrique);
          if Torg = nil then
            Torg := RendVentirem(Tob_VentiRemCEG, Rubrique);
          if Torg <> nil then
          begin
            //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
            LgRac := RendLongueurRacine(Torg.GetValue('PVS_RACINE1'));
            TheRacine := Copy(Torg.GetValue('PVS_RACINE1'), 1, LgRac); // $$$$
            QuelRacine := 1;
            LeSens := Torg.GetValue('PVS_SENS1');
            TypeSalPat := 'S';
            if TheRacine = '' then
            begin
              //   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
              LgRac := RendLongueurRacine(Torg.GetValue('PVS_RACINE2'));
              TheRacine := Copy(Torg.GetValue('PVS_RACINE2'), 1, LgRac); // $$$$
              QuelRacine := 2;
              LeSens := Torg.GetValue('PVS_SENS2');
              TypeSalPat := 'S';
            end;
            //   PT19 : 04/06/02   : V582  PH Ajout paramètre TraceErr
            Compte := ConfectionCpte(TheRacine, 'H', Q, FALSE, nil, TraceErr);
            //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
            LeLib := 'La ventilation de la rémunération ' + Rubrique;
            AlimEcrCompta(LeLib, '', Etab, Nature, Compte, LeSens, TypeSalPat, Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux, Q, QuelRacine, QSal, TraceErr);
          end;
        end
        else
        begin
          Q.Next;
          continue;
        end; // Cas d'une base de cotisation ==> rubrique non traitée
      end;
      if compte = '' then
      begin
        Q.Next;
        continue;
      end;
      // pas de compte trouvé donc ligne non passée en compta car ligne de calcul intermédiaire par exemple

      // Rajouter ici Fonction pour les ventilations analytiques sur les comptes de la classe 6

      // Rajouter ici les ventilations analytiques de la paie pour chaque rubrique/Salarie/DateDebut/dateFin

      Q.NEXT;
    end; // Fin du while sur la query liste des lignes du bulletins
    Ferme(Q);

    Trace.Items.Add('Salarié ' + Salarie + ' ' + TPPU.GetValue('PSA_LIBELLE') + ' ' +
      ' Etablissement : ' + RechDom('TTETABLISSEMENT', TPPU.GetValue('PPU_ETABLISSEMENT'), FALSE) + ' ');
    Trace.Refresh;
    if QSal <> nil then Ferme(QSal);
    //    QMul.NEXT;
  end; // Fin boucle sur la Query du mul

  //  Tob_LesEcrit.Detail.Sort('E_ETABLISSEMENT;E_NUMEROPIECE;E_NUMLIGNE'); PT53
    // confection de la TOB des lignes des ecritures de paie
  TEcritPaie := TOB.create('Les od de la paie', nil, -1);
  // PT-6 : 03/12/01   : V563  PH test equilibre débit/crédir sur chaque FOLIO
  TotDebit := 0;
  TotCredit := 0;
  TotDebit2 := 0;
  TotCredit2 := 0;
  // DEB PT53
  // Boucle pour éliminer les lignes vides le cas échéant et memorisation des 2 comptes pour la gestion des contre parties
//  ASauter := FALSE;
  // Fonction pour reformer la TOB
  if (Method <> 'PRI') and (TSauvLesEcrit <> nil) then
  begin
    if TSauvLesEcrit.detail.count > 0 then
    begin
      SauvLesEcrit(Tob_LesEcrit, TSauvLesEcrit); // recopie de la dernière TOB par etablissement calculée
      FreeAndNIL(Tob_LesEcrit);
      Tob_LesEcrit := TSauvLesEcrit;
    end;
  end;
  Tob_LesEcrit.Detail.Sort('E_ETABLISSEMENT;E_NUMEROPIECE;E_NUMLIGNE');
  MEtab := '';
  // FIN PT53
  if Method = 'MUL' then TContreP := TOB.Create('Les contre parties par établissement', nil, -1);

  TPR := Tob_LesEcrit.FindFirst([''], [''], FALSE);
  while TPR <> nil do
  begin
    //  PT-2 : 07/09/2001 : V547  PH Correction génération multi etab
    //     if MEtab = '' then MEtab := TPR.GetValue ('E_ETABLISSEMENT');
    ASauter := FALSE;

    if (TPR.GetValeur(iE_DEBIT) = 0) and (TPR.GetValeur(iE_CREDIT) = 0) then
    begin
      TPR.free;
      ASauter := TRUE;
    end;
    if not ASauter then
    begin
      if Method <> 'MUL' then
      begin
        //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
        if (Prem421 = '') and (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then Prem421 := TPR.GetValeur(iE_GENERAL);
        if (PremCpt = '') and (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) <> Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then PremCpt := TPR.GetValeur(iE_GENERAL);
      end
      else
      begin
        st := TPR.GetValeur(iE_ETABLISSEMENT);
        if (st <> '') then
        begin
          //          TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
          TCP := RendTContreP(TcontreP, TPR.GetValeur(iE_ETABLISSEMENT));
          if TCP = nil then
          begin
            TCP := TOB.create('Une ligne', TContreP, -1);
            TCP.AddChampSup('ETABLISSEMENT', FALSE);
            TCP.PutValue('ETABLISSEMENT', TPR.GetValeur(iE_ETABLISSEMENT));
            TCP.AddChampSup('PREM421', FALSE);
            TCP.AddChampSup('FOLIO1', FALSE);
            TCP.AddChampSup('FOLIO2', FALSE);
            //   PT-2 : 07/09/2001 : V547  PH rajout AddChampSup NUMECR1 et 2
            TCP.AddChampSup('NUMECR1', FALSE);
            TCP.AddChampSup('NUMECR2', FALSE);
            TCP.AddChampSup('PREMCPT', FALSE);
            TCP.PutValue('FOLIO1', ''); // Valeurs par defaut
            TCP.PutValue('FOLIO2', '');
            TCP.PutValue('PREM421', '');
            TCP.PutValue('PREMCPT', '');
            TCP.PutValue('NUMECR1', 1);
            TCP.PutValue('NUMECR2', 1);
          end;
          //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
          if (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TCP.GetValue('PREM421') = '') then
            TCP.PutValue('PREM421', TPR.GetValeur(iE_GENERAL));
          if (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) <> Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TCP.GetValue('PREMCPT') = '') then
            TCP.PutValue('PREMCPT', TPR.GetValeur(iE_GENERAL));
          //   PT-2 : 07/09/2001 : V547  Rajout stockage Etablissement en cours de traitement
        end; // Rupture Etablissement donc piece differente donc contrepartie aussi
      end;
    end;
    TPR := Tob_LesEcrit.FindNext([''], [''], FALSE);
  end;
  // Boucle pour affectation définitive de la tob des ecritures comptables après la fin des centralisations
  NbDec := V_PGI.OkdecV;
  for I := 0 to Tob_LesEcrit.Detail.count - 1 do
  begin
    TPR := Tob_LesEcrit.Detail[I];
    // Gestion des contre parties si 4210000 alors 1ere ligne au dessous ou au dessus soit  PremCpt
    //                            sinon  Prem421 1er compte 42100000 trouvé
    Compte := TPR.GetValeur(iE_GENERAL);
    if Method <> 'MUL' then
    begin
      //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
      if (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then TPR.PutValeur(iE_CONTREPARTIEGEN, PremCpt)
      else TPR.PutValeur(iE_CONTREPARTIEGEN, Prem421);
    end
    else
    begin
      //      TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
      TCP := RendTContreP(TcontreP, TPR.GetValeur(iE_ETABLISSEMENT));
      //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
      if (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then
      begin
        if TCP <> nil then TPR.PutValeur(iE_CONTREPARTIEGEN, TCP.GetValue('PREMCPT'));
      end
      else
      begin
        if TCP <> nil then TPR.PutValeur(iE_CONTREPARTIEGEN, TCP.GetValue('PREM421'));
      end;

    end;
    //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
    if (Copy(TPR.GetValeur(iE_GENERAL), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TPR.GetValeur(iE_AUXILIAIRE) <> '') then
    begin // Compte Auxiliaire renseigné alors on va rechercher le lettrage du tiers et son RIB
      CpteAuxi := TPR.GetValeur(iE_AUXILIAIRE);
      sRib := '';
      Q := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + CpteAuxi + '" AND R_SALAIRE="X"', True);
      if not Q.EOF then sRib := EncodeRIB(Q.FindField('R_ETABBQ').AsString, Q.FindField('R_GUICHET').AsString,
          Q.FindField('R_NUMEROCOMPTE').AsString, Q.FindField('R_CLERIB').AsString,
          Q.FindField('R_DOMICILIATION').AsString);
      Ferme(Q);
      TPR.PutValeur(iE_RIB, sRib);
      TTiers := RendTiers(Tob_Tiers, CpteAuxi);
      Lettrable := FALSE;
      if TTiers <> nil then Lettrable := (TTiers.GetValue('T_LETTRABLE') = 'X');
      //  PT26 : 07/01/03   : V591  PH Affectation champs table ecriture si Auxiliaire lettrable
      if Lettrable then
      begin
        TPR.PutValeur(iE_ETATLETTRAGE, 'AL');
        //        TPR.PutValue('E_NUMECHE', 1);
        //        TPR.PutValue('E_ECHE', 'X');
        //        TPR.PutValue('E_MODEPAIE', Q.FindField('T_MODEREGLE').AsString);
      end
      else
      begin
        TPR.PutValeur(iE_ETATLETTRAGE, '');
        //        TPR.PutValue('E_NUMECHE', 0);
        //        TPR.PutValue('E_ECHE', '-');
        //        TPR.PutValue('E_MODEPAIE', '');
      end;
      // FIN PT26
      if TTiers <> nil then
      begin // FQ 10715
        TRegl := TModeRegl.FindFirst(['MR_MODEREGLE'], [TTiers.GetValue('T_MODEREGLE')], FALSE);
        if TRegl <> nil then TPR.PutValeur(iE_MODEPAIE, TRegl.GetValue('MR_MP1'))
        else TPR.PutValeur(iE_MODEPAIE, 'CHQ'); // PT50 FQ 11318
      end
      else TPR.PutValeur(iE_MODEPAIE, 'CHQ');
    end;
    if (TPR.GetValeur(iE_DEBIT) > 0) or (TPR.GetValeur(iE_CREDIT) < 0) then TPR.PutValeur(iE_ENCAISSEMENT, 'ENC')
    else TPR.PutValeur(iE_ENCAISSEMENT, 'DEC'); // Affectation code encaissement
    //    TGene := TOB_Generaux.FindFirst(['G_GENERAL'], [TPR.GetValue('E_GENERAL')], FALSE);
    TGene := RendGene(TOB_Generaux, TPR.GetValeur(iE_GENERAL)); // PT53
    if TGene <> nil then
    begin
      // Gestion des confodentiels sur les comptes généraux. Pas de gestion de la confidentialité des auxilaires
      TPR.PutValeur(iE_CONFIDENTIEL, TGene.GetValue('G_CONFIDENTIEL')); // PT52
      // si lettrable alors gestion des echéance et du mode de paiement
      if (TGene.GetValue('G_LETTRABLE') = 'X') then
      begin
        TPR.PutValeur(iE_NUMECHE, 1);
        TPR.PutValeur(iE_ECHE, 'X');
        if MDR_Def <> '' then TPR.PutValeur(iE_MODEPAIE, MDR_Def)
        else TPR.PutValeur(iE_MODEPAIE, 'CHQ'); // PT50 FQ 11318
        //   PT16 : 02/05/02   : V582  PH Si compte lettrable alors date échéance = date écriture
        TPR.PutValeur(iE_DATEECHEANCE, TPR.GetValeur(iE_DATECOMPTABLE));
      end
      else
      begin
        //   PT26 : 07/01/03   : V591  PH Affectation champs table ecriture si Auxiliaire lettrable
        if (TPR.GetValeur(iE_NUMECHE) <> 1) then
        begin
          TPR.PutValeur(iE_NUMECHE, 0);
          TPR.PutValeur(iE_ECHE, '-');
        end;
      end;
      TPR.PutValeur(iE_DATEPAQUETMIN, TPR.GetValeur(iE_DATECOMPTABLE));
      TPR.PutValeur(iE_DATEPAQUETMAX, TPR.GetValeur(iE_DATECOMPTABLE));
      // FIN PT26
      if (TGene.GetValue('G_VENTILABLE') = 'X') then TPR.PutValeur(iE_ANA, 'X')
      else TPR.PutValeur(iE_ANA, '-'); // FQ 10661
      //  PT-8 : 19/02/02   : V571  PH test G_NATUREGENE pour affectation si ecirture lettrable
      if (TGene.GetValue('G_LETTRABLE') = 'X') and
        ((TGene.GetValue('G_NATUREGENE') = 'TIC') or (TGene.GetValue('G_NATUREGENE') = 'TID')or (TGene.GetValue('G_NATUREGENE') = 'DIV')) then // PT71
        TPR.PutValeur(iE_ETATLETTRAGE, 'AL');
      // FIN PT-8
      if ((TGene.GetValue('G_NATUREGENE') = 'COS') or (TGene.GetValue('G_NATUREGENE') = 'COC') or (TGene.GetValue('G_NATUREGENE') = 'COF') or (TGene.GetValue('G_NATUREGENE') =
        'COD')) and (TPR.Getvaleur(iE_AUXILIAIRE) = '')
        then // FQ 10641
        TraceErr.Items.Add('Ecriture compte ' + TPR.GetValeur(iE_GENERAL) + 'Folio ' + IntToStr(TPR.GetValeur(iE_NUMEROPIECE)) + ' ligne ' + IntToStr(TPR.GetValeur(iE_NUMLIGNE)) +
          ' n''a pas de compte auxiliaire associé');
      LettrableAux := FALSE;
      TheModRegl := TGene.GetValue('G_MODEREGLE');
      if (TPR.GetValeur(iE_AUXILIAIRE) <> '') then
      begin
        // DEB PT53
        TTiers := RendTiers(Tob_Tiers, TPR.GetValeur(iE_AUXILIAIRE)); // PT64 Remplacé CpteAuxi
        if TTiers <> nil then
        begin
          LettrableAux := (TTiers.GetValue('T_LETTRABLE') = 'X');
          TheModRegl := TTiers.GetValue('T_MODEREGLE');
          TPR.putValeur(iE_LIBELLE, TTiers.getvalue('T_LIBELLE')); // PT56-1 Libellé ecriture
          TPR.PutValeur(iE_CONFIDENTIEL, TTiers.getvalue('T_CONFIDENTIEL')); // PT56-2 Confidentialité compte auxiliaire
        end;
        // FIN PT53
      end;
      NatGene := TGene.GetValue('G_NATUREGENE');
      GLet := TGene.GetValue('G_LETTRABLE');
      GPoint := TGene.GetValue('G_POINTABLE');
      if (((NatGene = 'TIC') or (NatGene = 'TID') or (NatGene = 'DIV')) and (GLet = 'X')) or // PT71
        (((NatGene = 'BQE') or (NatGene = 'CAI')) and (GPoint = 'X')) or
        ((TPR.Getvaleur(iE_AUXILIAIRE) <> '') and LettrableAux)
        then // FQ 10661
      begin
        TPR.PutValeur(iE_ETATLETTRAGE, 'AL'); // PT50 FQ 11318
        TPR.PutValeur(iE_ECHE, 'X');
        TPR.PutValeur(iE_NUMECHE, 1);
      end
      else
      begin
        TPR.PutValeur(iE_ETATLETTRAGE, 'RI'); // PT50
        TPR.PutValeur(iE_ECHE, '-');
        TPR.PutValeur(iE_NUMECHE, 0);
      end;
    end
    else
    begin // Cas où le compte général n'existe pas
      if (IntegODPaie = 'IMM') or (IntegODPaie = 'DIF') then
      begin
        TraceErr.Items.Add('Le compte général ' + TPR.GetValeur(iE_GENERAL) + ' n''existe pas');
      end;
    end;
    //   PT40 : 24/03/2004 : V_500 FQ 11207 PH Affectation du mode de réglement par defaut sur les comptes de auxiliaires lettrables
    if LettrableAux then
    begin
      TPR.PutValeur(iE_MODEPAIE, TheModRegl);
      TPR.PutValeur(iE_ETATLETTRAGE, 'AL');
    end;
    // FIN PT40

    TPR.PutValeur(iE_NUMORDRE, TPR.GetValeur(iE_NUMLIGNE)); // Pour saisie par bordereau
    TOd := TOB.create('ECRITODPAIE', TEcritPaie, -1);
    if TOd <> nil then
    begin
      Tod.PutValue('PEC_ETABLISSEMENT', TPR.GetValeur(iE_ETABLISSEMENT));
      Tod.PutValue('PEC_JOURNAL', TPR.GetValeur(iE_JOURNAL));
      Tod.PutValue('PEC_GENERAL', TPR.GetValeur(iE_GENERAL));
      Tod.PutValue('PEC_DATECOMPTABLE', TPR.GetValeur(iE_DATECOMPTABLE));
      Tod.PutValue('PEC_NUMEROPIECE', TPR.GetValeur(iE_NUMEROPIECE));
      Tod.PutValue('PEC_NUMLIGNE', TPR.GetValeur(iE_NUMLIGNE));
      Tod.PutValue('PEC_AUXILIAIRE', TPR.GetValeur(iE_AUXILIAIRE));
      Tod.PutValue('PEC_LIBELLE', TPR.GetValeur(iE_LIBELLE));
      Tod.PutValue('PEC_REFINTERNE', TPR.GetValeur(iE_REFINTERNE));
      if PaieEuro = ComptaEuro then
      begin // on remplit toujours les montants dans la monnaie de tenue de la paie
        Tod.PutValue('PEC_DEBIT', ARRONDI(TPR.GetValeur(iE_DEBIT), NbDec));
        Tod.PutValue('PEC_CREDIT', ARRONDI(TPR.GetValeur(iE_CREDIT), NbDec));
        // PT-6 : 03/12/01   : V563  PH test equilibre débit/crédir sur chaque FOLIO
        if Tod.GetValue('PEC_NUMEROPIECE') = 1 then
        begin
          TotDebit := TotDebit + ARRONDI(TPR.GetValeur(iE_DEBIT), NbDec);
          TotCredit := TotCredit + ARRONDI(TPR.GetValeur(iE_CREDIT), NbDec);
        end
        else
        begin
          TotDebit2 := TotDebit2 + ARRONDI(TPR.GetValeur(iE_DEBIT), NbDec);
          TotCredit2 := TotCredit2 + ARRONDI(TPR.GetValeur(iE_CREDIT), NbDec);
        end;
      end
      else
      begin
        //  PT30      Tod.PutValue('PEC_DEBIT', ARRONDI(TPR.GetValue('E_DEBITEURO'), NbDec));
        //  PT30      Tod.PutValue('PEC_CREDIT', ARRONDI(TPR.GetValue('E_CREDITEURO'), NbDec));
                // PT-6 : 03/12/01   : V563  PH test equilibre débit/crédir sur chaque FOLIO
        if Tod.GetValue('PEC_NUMEROPIECE') = 1 then
        begin
          //  PT30      TotDebit := TotDebit + ARRONDI(TPR.GetValue('E_DEBITEURO'), NbDec);
          //  PT30      TotCredit := TotCredit + ARRONDI(TPR.GetValue('E_CREDITEURO'), NbDec);
        end
        else
        begin
          // PT30          TotDebit2 := TotDebit2 + ARRONDI(TPR.GetValue('E_DEBITEURO'), NbDec);
          //  PT30         TotCredit2 := TotCredit2 + ARRONDI(TPR.GetValue('E_CREDITEURO'), NbDec);
        end;
      end;
      if VH_PAIE.PGTenueEuro then Tod.PutValue('PEC_MONNAIE', 'EUR')
      else Tod.PutValue('PEC_MONNAIE', 'FRF');
    end;
  end;
  // PT-6 : 03/12/01   : V563  PH test equilibre débit/crédir sur chaque FOLIO
  TotDebit := Arrondi(TotDebit, NbDec);
  TotCredit := Arrondi(TotCredit, NbDec);
  TotDebit2 := Arrondi(TotDebit2, NbDec);
  TotCredit2 := Arrondi(TotCredit2, NbDec);

  St := 'Débit = ' + StrfMontant(TotDebit + TotDebit2, 15, NbDec, '', TRUE) + '   Crédit = ' + StrfMontant(TotCredit + TotCredit2, 15, NbDec, '', TRUE);
  GetLocalTime(SystemTime0);
  Trace.Items.Add('Fin du traitement des paies à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  //   PT12 : 26/03/02   : V575  PH Rajout ligne totalisation dans la TOB pour imprimer les totaux
  TZZ := TOB.create('ECRITODPAIE', TEcritPaie, -1);
  TZZ.PutValue('PEC_ETABLISSEMENT', '');
  TZZ.PutValue('PEC_JOURNAL', '');
  TZZ.PutValue('PEC_GENERAL', '');
  TZZ.PutValue('PEC_DATECOMPTABLE', NULL);
  TZZ.PutValue('PEC_NUMEROPIECE', 99999999);
  TZZ.PutValue('PEC_NUMLIGNE', 99999999);
  TZZ.PutValue('PEC_AUXILIAIRE', '');
  TZZ.PutValue('PEC_LIBELLE', 'Totaux : ');
  //   PT14 : 02/05/02   : V582  PH Référence interne dans toutes les lignes ecritures
  TZZ.PutValue('PEC_REFINTERNE', RefPaie);
  TZZ.PutValue('PEC_DEBIT', Arrondi(TotDebit, NbDec) + Arrondi(TotDebit2, NbDec));
  TZZ.PutValue('PEC_CREDIT', Arrondi(TotCredit, NbDec) + Arrondi(TotCredit2, NbDec));
  TZZ.PutValue('PEC_MONNAIE', '');
  st := st + ';' +
    'PEC_ETABLISSEMENT,PEC_JOURNAL,PEC_DATECOMPTABLE,PEC_NUMEROPIECE,PEC_NUMLIGNE,PEC_GENERAL,PEC_AUXILIAIRE,PEC_LIBELLE,PEC_DEBIT,PEC_CREDIT,PEC_REFINTERNE';
  try
    Aleat := Random(97531);
    if (JournalPaie.Value = '') then
      st1 := 'DELETE FROM ECRITODPAIE WHERE PEC_DATECOMPTABLE="' + USDateTime(DateF) + '"'
    else st1 := 'DELETE FROM ECRITODPAIE WHERE PEC_JOURNAL="' + JournalPaie.Value + '" AND PEC_DATECOMPTABLE="' + USDateTime(DateF) + '"';
    if Etabl <> '' then
    begin
      st1 := st1 + ' AND PEC_ETABLISSEMENT = "' + Etabl + '"';
    end;
    ExecuteSQL(st1);
    MaTob := TOB.Create('Ma TOB', nil, -1);
    for i := 0 to TEcritPaie.detail.count - 1 do
    begin
      if TEcritPaie.detail[i].GetValue('PEC_GENERAL') <> '' then
      begin
        Ta := TOB.create('ECRITODPAIE', MaTob, -1);
        Ta.Dupliquer(TEcritPaie.detail[i], FALSE, TRUE, TRUE);
        Ta.PutValue('PEC_ALEAT', Aleat);
        Ta.PutValue('PEC_NUMLIGNE', i + 1);
      end;
    end;
    begintrans;
    MaTob.InsertDB(nil, FALSE);
    CommitTrans;
    st := ' AND PEC_ALEAT = ' + FloatToStr(Aleat);
    LanceEtat('E', 'PAN', 'POD', True, FALSE, False, nil, St, '', False);
    //   PT42 : 06/05/2004 : V_500 PH FQ 11173 Mise ne place liste d'exportation
    if GetControlText('LISTEEXPORT') = 'X' then
      LanceEtatTOB('E', 'PAN', 'POD', MaTob, True, true, False, Pan, '', '', False);
    FreeAndNIL(MaTob);
    ExecuteSql('DELETE FROM ECRITODPAIE WHERE PEC_ALEAT = ' + FloatToStr(Aleat));
  except
    Rollback;
    PGIError('Une erreur est survenue lors de l''édition du journal', 'Génération comptable');
  end;

  //PGVisuUnObjet (TEcritPaie, 'Les ODs de paie', St) ; // visualisation des lignes au format Paie PGI
  OkOk := mrYes;
  TZZ.Free;
  Ecart := Arrondi(TotDebit - TotCredit, 2);
  Ecart1 := Arrondi(TotDebit2 - TotCredit2, 2);
  if (Ecart <> 0) or (Ecart1 <> 0) then
  begin
    st := '';
    if Ecart <> 0 then st := 'Attention, le folio 1 n''est pas équilibré.';
    if Ecart1 <> 0 then st := st + 'Attention, le folio 2 n''est pas équilibré';
    st := st + '#13#10Votre écriture n''est pas équilibrée, vérifiez vos erreurs et les ventilations ';
    PGIError(St, 'Abandon du traitement');
    if IntegODPaie = 'ECP' then // Format paie Pgi
    begin
      if VH_Paie.PGCHEMINEAGL = '' then PgiError('Vous n''avez pas renseigné le chemin de stockage dans les paramètres société', 'Abandon du traitement')
      else
      begin
        if Etabl = '' then st := VH_Paie.PGCHEMINEAGL + '\ODPAIEPGI.TXT'
        else st := VH_Paie.PGCHEMINEAGL + '\ODPAIEPGI' + Etabl + '.TXT';
        TEcritPaie.SaveToFile(St, FALSE, TRUE, FALSE, '');
      end;
    end;
    OkOk := mrNo;
  end;

  //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
  //   PT33 : 14/11/03  : V_500 PH Affinage message anomalie compte non present dans le modele uniquement si déséquilibre
  if Assigned(ErrorGene) and (IntegODPaie <> 'ECP') then // PT45 // and ((Ecart <> 0) or (Ecart1 <> 0))
  begin
    for ij := 0 to ErrorGene.Count - 1 do
    begin
      if ((Ecart <> 0) or (Ecart1 <> 0)) then TraceErr.Items.Add(ErrorGene.Strings[ij]);
    end;
  end;
  if Assigned(ErrorGene) then FreeAndNil(ErrorGene);

  if (TraceErr.Items.Count > 1) and (IntegODPaie <> 'ECP') then
  begin
    st := 'Vous avez des comptes inexistants, vérifiez vos erreurs ';
    PGIError(St, 'Abandon du traitement');
    Tbsht := TTabSheet(GetControl('TBSHTERROR'));
    if Tbsht <> nil then Pan.ActivePage := Tbsht;
    OkOk := mrNo;
  end;
  //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
  //    AND (IntegODPaie <> 'ECP')
  if (OkOk = mrYes) then
  begin // B1
    NumEcritOD1 := 1;
    NumEcritOD2 := 1;
    // Affectation definitive du numéro de piece
    if VH_Paie.PGAnalytique then TOBANA := TOB.Create('LES ECRITURES ANA', nil, -1);

    // recherche du nombre de folios décrits dans le Guide
    for I := 0 to Tob_GuideEcr.Detail.count - 1 do
    begin
      TPR := Tob_GuideEcr.Detail[I];
      if (StrToInt(TPR.GetValue('PGC_NUMFOLIO')) = 1) then NBrePiece := 1;
      if (StrToInt(TPR.GetValue('PGC_NUMFOLIO')) = 2) then
      begin
        NBrePiece := 2;
        break;
      end;
    end;
    // Affectation par defaut des numéros de pièce par rapport au journal
    //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
    if (IntegODPaie <> 'ECP-') then // dans le cas où on integre dans la compta PGI
    begin // B2
      if Method <> 'MUL' then // cas mono etab ou sur etab principal
      begin // B3
        if not VH_Paie.PGTypeEcriture then NumFolioOD1 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'N'), MM.DateC)
        else NumFolioOD1 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'O'), MM.DateC);
        if NBrePiece = 2 then
        begin
          if not VH_Paie.PGTypeEcriture then NumFolioOD2 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'N'), MM.DateC)
          else NumFolioOD2 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'O'), MM.DateC);
        end;
        // DEB PT51
        if (NbrePiece = 2) and (NumFolioOD1 = NumFolioOD2) then
        begin
          St := 'Attention, votre journal est de type bordereau alors #13#10que vous voulez gérer 2 pièces !';
          st := St + '#13#10Si vous continuez alors vos 2 folios seront mis dans la même pièce.';
          st := St + '#13#10 Etes vous sûr(e) de vouloir continuer ?';
          OkOk := PgiAsk(St, Ecran.Caption);
          if OkOk <> MrYes then PgiInfo('Le traitement est abandonné');
        end;
        // FIN PT51
        // Affectation definitive du numéro de piece - boucle sur les ecritures pour affecter les numeros de pieces et de lignes
        for I := 0 to Tob_LesEcrit.Detail.count - 1 do
        begin // B4
          TPR := Tob_LesEcrit.Detail[I];
          OldP := TPR.GetValeur(iE_NUMEROPIECE);
          OldL := TPR.GetValeur(iE_NUMLIGNE);
          if (TPR.GetValeur(iE_NUMEROPIECE) = 1) then
          begin // B5
            NewP := NumFolioOD1;
            NewL := NumEcritOD1;
            TPR.PutValeur(iE_NUMEROPIECE, NumFolioOD1);
            TPR.PutValeur(iE_NUMLIGNE, NumEcritOD1);
            NumEcritOD1 := NumEcritOD1 + 1;
            //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
            if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
          end // F5
          else
          begin // B6
            if (TPR.GetValeur(iE_NUMEROPIECE) = 2) then
            begin // B7
              NewP := NumFolioOD1;
              NewL := NumEcritOD1;
              TPR.PutValeur(iE_NUMEROPIECE, NumFolioOD2);
              TPR.PutValeur(iE_NUMLIGNE, NumEcritOD2);
              NumEcritOD2 := NumEcritOD2 + 1;
              //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
              if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
            end; // F7
          end; // F6
        end; // Fin du for  // F4
      end // F3 Fin du cas mono ou etab principal
      else // Multi etab donc recherche et affectation des numéros de pieces pour chaque etablissement
      begin // B8
        //   PT-2 : 07/09/2001 : V547  PH Recherche depuis le debut de la TOB et non sur un etablissement
        TCP := TContreP.FindFirst([''], [''], FALSE);
        while (TCP <> nil) and (NBrePiece > 0) do
        begin // B9  Boucle affectation par journal des numéros de pieces
          if TCP.GetValue('FOLIO1') = '' then
          begin // B10
            if not VH_Paie.PGTypeEcriture then NumFolioOD1 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'N'), MM.DateC)
            else NumFolioOD1 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'O'), MM.DateC);
            TCP.PutValue('FOLIO1', NumFolioOD1);
            if NBrePiece = 2 then
            begin // B11
              if not VH_Paie.PGTypeEcriture then NumFolioOD2 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'N'), MM.DateC)
              else NumFolioOD2 := GetNewNumJal(JournalPaie.Value, (MM.Simul = 'O'), MM.DateC);
              TCP.PutValue('FOLIO2', NumFolioOD2);
            end; // F11
          end; // F10
          //   PT-2 : 07/09/2001 : V547  PH Boucle sur tous les enregistrements etablissements
          TCP := TContreP.FindNext([''], [''], FALSE);
        end; // F9
        // DEB PT59 Message alerte mais non bloquant au cas ....
        if (NbrePiece = 2) and (NumFolioOD1 = NumFolioOD2) then
        begin
          St := 'Attention, votre journal est de type bordereau alors #13#10que vous voulez gérer 2 pièces !';
          st := St + '#13#10Si vous continuez alors vous n''aurez qu''un seul folio.';  /// PT70
          st := St + '#13#10 Etes vous sûr(e) de vouloir continuer ?';
          OkOk := PgiAsk(St, Ecran.Caption);
          if OkOk <> MrYes then PgiInfo('Le traitement est abandonné');
        end;
        // FIN PT59
        //PGVisuUnObjet (TContreP, 'Les contreparties', '') ; // visualisation des lignes au format Paie PGI
        EcrPositif := TRUE;
        for I := 0 to Tob_LesEcrit.Detail.count - 1 do
        begin // B12
          TPR := Tob_LesEcrit.Detail[I];
          OldP := 0;
          OldL := 0;
          NewP := 0;
          NewL := 0;
          if EcrPositif then
          begin
            if TPR.GetValeur(iE_DEBIT) < 0 then EcrPositif := FALSE;
            if TPR.GetValeur(iE_CREDIT) < 0 then EcrPositif := FALSE;
          end;
          //          TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
          TCP := RendTContreP(TcontreP, TPR.GetValeur(iE_ETABLISSEMENT));
          if TCP <> nil then
          begin
            OldP := TPR.GetValeur(iE_NUMEROPIECE);
            OldL := TPR.GetValeur(iE_NUMLIGNE);
            NewP := TCP.GetValue('FOLIO1');
            NewL := TCP.GetValue('NUMECR1');
          end;
          if (TPR.GetValeur(iE_NUMEROPIECE) = 1) then
          begin // B13
            TPR.PutValeur(iE_NUMEROPIECE, TCP.GetValue('FOLIO1'));
            TPR.PutValeur(iE_NUMLIGNE, TCP.GetValue('NUMECR1'));
            TCP.PutValue('NUMECR1', TCP.GetValue('NUMECR1') + 1);
            //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
            if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
          end // F13
          else
          begin // B14
            if (TPR.GetValeur(iE_NUMEROPIECE) = 2) then
            begin // B15
              TPR.PutValeur(iE_NUMEROPIECE, TCP.GetValue('FOLIO2'));
              TPR.PutValeur(iE_NUMLIGNE, TCP.GetValue('NUMECR2'));
              TCP.PutValue('NUMECR2', TCP.GetValue('NUMECR2') + 1);
              //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
              if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
            end; // F15
          end; // F14 Fin du Else
        end; // F12 fin du for boucle sur les ecritures sur affectation numeros de pieces et de lignes
      end; // F8 fin du else cas mono ou principal = cas multi etab
    end; // F2 fin si on retraite les ecritures = cas integration dans la paie pgi
    // Equilibre écriture en fonction des écarts de conversion et du paramètrage de la compta
    //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
    //  if IntegODPaie <> 'ECP' then
    EquilibreEcrPaie(Tob_LesEcrit, MM, NbDec);
  end;
  //   PT32 : 12/10/03  : V_421 PH Test des ecritures négatives uniquement en fin de calcul
  EcrPositif := TRUE;
  for I := 0 to Tob_LesEcrit.Detail.count - 1 do
  begin
    if EcrPositif then
      if (TPR.GetValeur(iE_DEBIT) < 0) or (TPR.GetValeur(iE_CREDIT) < 0) then EcrPositif := FALSE;
  end;

  if (not EcrPositif) and (not VH^.MontantNegatif) then
  begin
    PgiError('Vous avez une écriture négative alors que vous n''autorisez pas les montants négatifs #13#10 dans la comptabilité PGI ', Ecran.caption);
    OkOk := MrNo;
  end;
  if (OkOk = mrYes) then
  begin
    if IntegODPaie = 'ECP' then st := 'Voulez-vous générer un fichier des ODs au format Paie PGI'
    else if IntegODPaie = 'IMM' then
    begin // DEB PT43
      if VH_PAIE.PgTypeEcriture then st := 'Voulez-vous enregistrer immédiatement vos écritures de simulation dans la comptabilité'
      else st := 'Voulez-vous enregistrer immédiatement vos écritures dans la comptabilité';
    end // FIN PT43
    else if IntegODPaie = 'DIF' then st := 'Voulez-vous enregistrer en différé vos écritures'
    else if IntegODPaie = 'EFC' then st := 'Voulez-vous générer un fichier des ODs au format Export Comptabilté PGI';
  end; // F1 Fin si integration autorisee
  //   PT42 : 06/05/2004 : V_500 PH FQ 11173 Mise ne place liste d'exportation
  // PT49 Rajout test si OkOk pas d'erreur
  if VH_Paie.PGAnalytique and (OkOk = mrYes) then AnalyseAnalytique(IntegODPaie, TOBAna, PaieEuro, ComptaEuro, NbDec, ConservAna, Pan);

  bOk := True;
  if OkOk = mrYes then OkOk := PGIAsk(St, Ecran.Caption);
  if OkOk = mrNo then bok := False;
  if OkOk = mrYes then
  begin
    if IntegODPaie = 'ECP' then // Format paie Pgi
    begin
      if VH_Paie.PGCHEMINEAGL = '' then PgiError('Vous n''avez pas renseigné le chemin de stockage', 'Paramètres société')
      else
      begin
        if Etabl = '' then st := VH_Paie.PGCHEMINEAGL + '\ODPAIEPGI.TXT'
        else st := VH_Paie.PGCHEMINEAGL + '\ODPAIEPGI' + Etabl + '.TXT';
        //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques dans le cas export format interne à la paie
        if not VH_Paie.PGAnalytique then TEcritPaie.SaveToFile(St, FALSE, TRUE, FALSE, '')
        else
        begin
          TEcritPaie.SaveToFile(St, FALSE, TRUE, FALSE, '');
          TOBANA.SaveToFile(St + 'Ana', FALSE, TRUE, FALSE, '');
        end;
      end;
    end
    else // autres cas
    begin
      if IntegODPaie = 'IMM' then // Integration immediate dans la paie PGI
      begin
        //      PGVisuUnObjet (Tob_LesEcrit, 'Les ODs de paie', '') ; // visualisation des lignes au format Paie PGI
//        Tob_LesEcrit.SaveToFile('C:\Ecriture.txt', FALSE, TRUE, FALSE, ''); // @@@@@@@@@@@
        if not Tob_LesEcrit.InsertDBByNivel(False) then V_PGI.IoError := oeUnknown
        else
        begin
          if not VH_Paie.PGTypeEcriture then MajSoldesEcritureTOB(Tob_LesEcrit, True); // MAJ des soldes comptables
        end;
      end;
      if IntegODPaie = 'DIF' then // Integration différé dans la paie PGI
      begin
        //  PT18 : 04/06/02   : V582  PH Récupération et personnalisation de la focntion insertiondifféree
        if not PGInsertionDifferee(Tob_LesEcrit) then V_PGI.IoError := oeUnknown;
      end;
    end; // fin du else
    ReecritPaiesCompta;
  end; // fin integration autorisee
  if IntegODPaie <> 'ECP' then OkOk := mrNO; // Force pour ne pas ecrire dans la cas d'une integration dans COMPTA PGIS5
  if (OkOk = mrYes) then
  begin // pas ecriture dans les tables de la paie des ecritures sera reservé pour les export vers autres compta que PGIS5
    try
      begintrans;
      if (JournalPaie.Value = '') then
        st := 'DELETE FROM ECRITODPAIE WHERE PEC_DATECOMPTABLE="' + USDateTime(DateF) + '"'
      else st := 'DELETE FROM ECRITODPAIE WHERE PEC_JOURNAL="' + JournalPaie.Value + '" AND PEC_DATECOMPTABLE="' + USDateTime(DateF) + '"';
      if Etabl <> '' then
      begin
        st := st + ' AND PEC_ETABLISSEMENT = "' + Etabl + '"';
      end;
      ExecuteSQL(st);
      // DEB PT58
      for i := 0 to TEcritPaie.detail.count - 1 do
      begin
        if TEcritPaie.detail[i].GetValue('PEC_GENERAL') <> '' then
        begin
          TEcritPaie.detail[i].PutValue('PEC_ALEAT', Aleat);
          TEcritPaie.detail[i].PutValue('PEC_NUMLIGNE', i + 1);
        end;
      end;
      // FIN PT58
      TEcritPaie.InsertDB(nil, FALSE);
      CommitTrans;
    except
      Rollback;
      PGIError('Une erreur est survenue lors de la Génération comptable', Ecran.caption);
    end;

  end;
  FreeAndNil(TEcritPaie);
  FreeAndNil(TContreP);
  FreeAndNIL(Tob_GuideEcr);
  FreeAndNIL(Tob_VentiOrg);
  FreeAndNIL(Tob_VentiRem);
  FreeAndNIL(Tob_VentiCot);
  // DEB PT73
  FreeAndNIL(Tob_VentiOrgSTD);
  FreeAndNIL(Tob_VentiRemSTD);
  FreeAndNIL(Tob_VentiCotSTD);
  FreeAndNIL(Tob_VentiOrgCEG);
  FreeAndNIL(Tob_VentiRemCEG);
  FreeAndNIL(Tob_VentiCotCEG);
  FreeAndNIL(Tob_Tiers);
  // FIN PT53
      // PT38
  FreeAndNIL(T_Etab);

  GetLocalTime(SystemTime0);
  //   PT20 : 05/06/02   : V582  PH Gestion historique des évènements
  if TraceErr.Items.Count > 1 then Trace.Items.Add('Traitement interrompu à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)))
  else Trace.Items.Add('Fin du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  PGIInfo('Le traitement est terminé : ' + IntToStr(NbrePaie) + ' Bulletins traités.', Ecran.caption);

  TheTrace := TStringList.Create;
  if (TraceErr.Items.Count > 1) or (OkOk <> mrYes) then
  begin
    TheTrace.Add('Génération interrompue car :');
    if Ecart <> 0 then TheTrace.Add('Attention, le folio 1 n''est pas équilibré.');
    if Ecart1 <> 0 then TheTrace.Add('Attention, le folio 2 n''est pas équilibré');
    St := 'Débit = ' + StrfMontant(TotDebit + TotDebit2, 15, NbDec, '', TRUE) + '   Crédit = ' + StrfMontant(TotCredit + TotCredit2, 15, NbDec, '', TRUE);
    TheTrace.Add(St);
    TheTrace.Add('Vous avez des erreurs ou vous avez interrompu le traitement');
    TheTrace.Add('Abandon du traitement');
    // DEB PT44 rajout de la trace concernant les erreurs
    TheTrace.add('===> Liste des erreurs rencontrées');
    for ij := 0 to TraceErr.Items.Count - 1 do
      TheTrace.add(TraceErr.Items.Strings[ij]);
    // FIN PT44
    CreeJnalEvt('001', '016', 'ERR', Trace, nil, TheTrace);
  end
  else
  begin
    TheTrace.Add('Génération terminée');
    St := 'Débit = ' + StrfMontant(TotDebit + TotDebit2, 15, NbDec, '', TRUE) + '   Crédit = ' + StrfMontant(TotCredit + TotCredit2, 15, NbDec, '', TRUE);
    TheTrace.Add(St);
    if IntegODPaie = 'DIF' then TheTrace.Add('Intégration en différée dans la comptabilité PGI')
    else if IntegODPaie = 'IMM' then TheTrace.Add('Intégration dans la comptabilité PGI')
    else TheTrace.Add('Génération fichier OD au format paie pgi');
    TheTrace.Add('Fin du traitement');
    CreeJnalEvt('001', '016', 'OK', Trace, nil, TheTrace);
  end;
  TheTrace.Free;
  // FIN PT20

  if TraceErr.Items.Count > 1 then
  begin
    Tbsht := TTabSheet(GetControl('TBSHTERROR'));
    if Tbsht <> nil then Pan.ActivePage := Tbsht;
  end;
  // RAZ liste des ventilations analytiques par ecritures comptables
  for i := 0 to TTA.Count - 1 do
  begin
    VideListe(T_AnalPaie(TTA[i]).Anal);
    T_AnalPaie(TTA[i]).Anal.Free;
    T_AnalPaie(TTA[i]).Anal := nil;
  end;
  VideListe(TTA);
  TTA.Free;
  TTA := nil;
  // Raz tob des ventilations analytiques salarie
  if Tob_Ventana <> nil then
  begin
    Tob_Ventana.Free;
    Tob_Ventana := nil;
  end;
  if TOB_VenSal <> nil then
  begin
    TOB_VenSal.Free;
    TOB_VenSal := nil;
  end;
  if TOBANA <> nil then
  begin
    TOBANA.Free;
    TOBANA := nil;
  end;
  FreeANdNil(AnalCompte);
  //  RAZ des TOB Utilisées
  FreeAndNil(Tob_HistoBulletin);
  FreeAndNil(Tob_Cotisations);
  FreeAndNil(TOB_Generaux);
  FreeAndNil(TOB_LesEcrit);
  FreeAndNil(TModeRegl);
//*TRA : oui pour enregister les écritures, aucunes erreurs, TRA coché
  if (bOk) and (IntegODPaie = 'TRA') then
  begin
    if PGIAsk('Voulez vous générer le fichier d''export TRA ?') = mrYes then
      ExportPaieTRA;
  end;
end;

procedure TOF_PG_GeneCompta.OnArgument(Arguments: string);
var
  F: TFVierge;
  st: string;
  BtnImp : TToolbarButton97;
begin
  inherited;
  st := Trim(Arguments);
  LblDD := THLabel(GetControl('LBLDU'));
  LblDF := THLabel(GetControl('LBLAU'));
  DateEcr := THLabel(GetControl('DATEFIN'));
  if LblDD <> nil then LblDD.Caption := ReadTokenSt(st);
  if LblDF <> nil then
  begin
    LblDF.Caption := ReadTokenSt(st);
    DateEcr.Caption := LblDF.Caption;
  end;
  Etabl := ReadTokenSt(st);
  Method := ReadTokenSt(st);
  LeWhere := ReadTokenSt(st);
  CritSal := ReadTokenSt(st); // PT60
  if Method = 'PRI' then Etabl := VH^.EtablisDefaut; // recuperation etablissement par defaut
  BtnLance := TToolbarButton97(GetControl('BTNLANCE'));
  if Method = '' then // ¨Par defaut cas mono etablissement
    if BtnLance <> nil then SetControlEnabled('BTNLANCE', FALSE);
  if BtnLance <> nil then BtnLance.OnClick := LanceGeneCompta;
  JournalPaie := THValComboBox(GetControl('CBXJNAL'));
  if JournalPaie <> nil then
  begin
    JournalPaie.Value := VH_Paie.PGJournalPaie;
    //    if VH_Paie.PGIntegODPaie = 'ECP' then JournalPaie.DataType := '' ;
  end;
  if VH_Paie.PGIntegODPaie = 'ECP' then //@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
  begin
    if JournalPaie <> nil then
    begin
      JournalPaie.Style := csDropDown;
      JournalPaie.MaxLength := 3; // Code journal de la compta à 3 caractères
    end;
  end;
  JeuPaie := THValComboBox(GetControl('CBXJEU'));
  if JeuPaie <> nil then JeuPaie.Value := VH_Paie.PGModeleEcr;
  if not (Ecran is TFVierge) then exit;
  //   PT25 : 17/12/02   : V585  PH Test de la query non nulle avant de lancer la génération
  F := TFVierge(Ecran);
  if F <> nil then
  begin
{$IFDEF EAGLCLIENT}
    QMUL := THQuery(F.FMULQ).TQ;
{$ELSE}
    QMUL := F.FMULQ;
{$ENDIF}
  end;
  if QMUL = nil then
  begin
    PgiError('La liste des paies est vide,traitement impossible', Ecran.caption);
    SetControlEnabled('BTNLANCE', FALSE);
  end;
  //pour export TRA on rend visible la boite à cocher
  SetControlVisible('EXPORTTRA', True);

  //DEB PT72
  BtnImp := TToolbarButton97 (GetControl ('Bimprimer'));
  if BtnImp <> NIL then BtnImp.OnClick := BImprimerClik;
  //FIN PT72
end;
// Initialisation par défaut des champs d'une ligne écriture

procedure TOF_PG_GeneCompta.LignePaieECR(MM: RMVT; TOBE: TOB);
begin
  {RMVT}
  TOBE.PutValue('E_JOURNAL', MM.Jal);
  TOBE.PutValue('E_EXERCICE', MM.Exo);
  TOBE.PutValue('E_DATECOMPTABLE', MM.DateC);
  TOBE.PutValue('E_ETABLISSEMENT', MM.Etabl);
  TOBE.PutValue('E_DEVISE', MM.CodeD);
  TOBE.PutValue('E_TAUXDEV', MM.TauxD);
  TOBE.PutValue('E_DATETAUXDEV', MM.DateTaux);
  TOBE.PutValue('E_QUALIFPIECE', MM.Simul);
  TOBE.PutValue('E_NATUREPIECE', MM.Nature);
  TOBE.PutValue('E_NUMEROPIECE', MM.Num);
  TOBE.PutValue('E_VALIDE', CheckToString(MM.Valide));
  TOBE.PutValue('E_DATEECHEANCE', MM.DateC);
  TOBE.PutValue('E_PERIODE', GetPeriode(MM.DateC));
  TOBE.PutValue('E_SEMAINE', NumSemaine(MM.DateC));
  TOBE.PutValue('E_REGIMETVA', VH^.RegimeDefaut);
  TOBE.PutValue('E_ETATLETTRAGE', 'RI');
  TOBE.PutValue('E_MODEPAIE', '');
  TOBE.PutValue('E_COTATION', 1);
  //  PT27 : 04/03/03   : V_42  PH Initialisation du champ E_QUALIFORIGINE à GEN pour la validation des écritures de simuls
  TOBE.PutValue('E_QUALIFORIGINE', 'GEN');
  TOBE.PutValue('E_VISION', 'DEM');
  TOBE.PutValue('E_ECRANOUVEAU', 'N');
  //  PT-8 : 19/02/02   : V571  E_MODESAISIE (BOR,LIB,-)=MM.ModeSaisieJal au lieu de '-'
  TOBE.PutValue('E_MODESAISIE', MM.ModeSaisieJal);
  // FIN PT8
  TOBE.PutValue('E_CONTROLETVA', 'RIE');
  TOBE.PutValue('E_ETAT', '0000000000');

end;
// Remplissage par défaut des champs d'une ligne écriture ==> au moins une valeur cohérente

procedure TOF_PG_GeneCompta.InitEcritDefaut(TOBE: TOB);
var
  i: integer;
begin
  if (iE_DEBIT = 0) then MemoriseNumChamp(TOBE);
  TOBE.PutValue('E_EXERCICE', '');
  TOBE.PutValue('E_JOURNAL', '');
  TOBE.PutValue('E_NUMEROPIECE', 0);
  TOBE.PutValue('E_DATECOMPTABLE', V_PGI.DateEntree);
  TOBE.PutValue('E_NUMLIGNE', 0);
  TOBE.PutValue('E_DATEREFEXTERNE', IDate1900);
  TOBE.PutValue('E_GENERAL', '');
  TOBE.PutValue('E_AUXILIAIRE', '');
  TOBE.PutValue('E_DEBIT', 0);
  TOBE.PutValue('E_CREDIT', 0);
  //   PT14 : 02/05/02   : V582  PH Référence interne dans toutes les lignes ecritures
  TOBE.PutValue('E_REFINTERNE', RefPaie);
  TOBE.PutValue('E_LIBELLE', '');
  TOBE.PutValue('E_NATUREPIECE', 'OD');
  TOBE.PutValue('E_QUALIFPIECE', 'N');
  TOBE.PutValue('E_TYPEMVT', 'DIV');
  TOBE.PutValue('E_VALIDE', '-');
  TOBE.PutValue('E_ETAT', '0000000000');
  TOBE.PutValue('E_REFEXTERNE', '');
  TOBE.PutValue('E_UTILISATEUR', V_PGI.User);
  TOBE.PutValue('E_CONTROLEUR', '');
  TOBE.PutValue('E_DATECREATION', Date);
  TOBE.PutValue('E_DATEMODIF', NowH);
  TOBE.PutValue('E_SOCIETE', V_PGI.CodeSociete);
  TOBE.PutValue('E_ETABLISSEMENT', '');
  TOBE.PutValue('E_BLOCNOTE', '');
  TOBE.PutValue('E_VISION', 'DEM');
  TOBE.PutValue('E_REFLIBRE', '');
  TOBE.PutValue('E_AFFAIRE', '');
  TOBE.PutValue('E_TVAENCAISSEMENT', '-');
  TOBE.PutValue('E_REGIMETVA', '');
  TOBE.PutValue('E_TVA', '');
  TOBE.PutValue('E_TPF', '');
  TOBE.PutValue('E_NUMEROIMMO', 0);
  TOBE.PutValue('E_BUDGET', '');
  TOBE.PutValue('E_CONTREPARTIEGEN', '');
  TOBE.PutValue('E_CONTREPARTIEAUX', '');
  TOBE.PutValue('E_COUVERTURE', 0);
  TOBE.PutValue('E_LETTRAGE', '');
  TOBE.PutValue('E_LETTRAGEDEV', '-');
  TOBE.PutValue('E_REFPOINTAGE', '');
  TOBE.PutValue('E_SUIVDEC', '');
  TOBE.PutValue('E_DATEPOINTAGE', IDate1900);
  TOBE.PutValue('E_MODEPAIE', '');
  TOBE.PutValue('E_NOMLOT', '');
  TOBE.PutValue('E_NIVEAURELANCE', 0);
  // PT30  TOBE.PutValue('E_DEBITEURO', 0);
  // PT30  TOBE.PutValue('E_CREDITEURO', 0);
  // PT30  TOBE.PutValue('E_SAISIEEURO', '-');
  TOBE.PutValue('E_DEVISE', V_PGI.DevisePivot);
  TOBE.PutValue('E_DEBITDEV', 0);
  TOBE.PutValue('E_CREDITDEV', 0);
  TOBE.PutValue('E_TAUXDEV', 0);
  TOBE.PutValue('E_CONTROLE', '-');
  TOBE.PutValue('E_TIERSPAYEUR', '');
  TOBE.PutValue('E_QTE1', 0);
  TOBE.PutValue('E_QTE2', 0);
  TOBE.PutValue('E_QUALIFQTE1', '...');
  TOBE.PutValue('E_QUALIFQTE2', '...');
  TOBE.PutValue('E_ECRANOUVEAU', 'N');
  TOBE.PutValue('E_DATEVALEUR', IDate1900);
  TOBE.PutValue('E_RIB', '');
  TOBE.PutValue('E_DATEPAQUETMIN', V_PGI.DateEntree);
  TOBE.PutValue('E_REFRELEVE', '');
  TOBE.PutValue('E_DATEPAQUETMAX', V_PGI.DateEntree);
  TOBE.PutValue('E_COUVERTUREDEV', 0);
  TOBE.PutValue('E_ETATLETTRAGE', 'RI');
  TOBE.PutValue('E_ENCAISSEMENT', 'RIE');
  TOBE.PutValue('E_COTATION', 0);
  TOBE.PutValue('E_TYPEANOUVEAU', '');
  TOBE.PutValue('E_EMETTEURTVA', '-');
  TOBE.PutValue('E_NUMECHE', 0);
  TOBE.PutValue('E_NUMPIECEINTERNE', '');
  TOBE.PutValue('E_ANA', '-');
  TOBE.PutValue('E_DATEECHEANCE', IDate1900);
  TOBE.PutValue('E_ECHE', '-');
  TOBE.PutValue('E_DATERELANCE', IDate1900);
  TOBE.PutValue('E_FLAGECR', '');
  TOBE.PutValue('E_DATETAUXDEV', V_PGI.DateEntree);
  TOBE.PutValue('E_CONTROLETVA', 'RIE');
  TOBE.PutValue('E_CONFIDENTIEL', '0');
  TOBE.PutValue('E_MULTIPAIEMENT', '');
  TOBE.PutValue('E_TRACE', '');
  TOBE.PutValue('E_CONSO', '');
  TOBE.PutValue('E_ORIGINEPAIEMENT', IDate1900);
  TOBE.PutValue('E_CREERPAR', 'PG');
  TOBE.PutValue('E_EXPORTE', '---');
  TOBE.PutValue('E_TRESOLETTRE', '-');
  // PT30  TOBE.PutValue('E_COUVERTUREEURO', 0);
  // PT30  TOBE.PutValue('E_LETTRAGEEURO', '-');
  TOBE.PutValue('E_REFLETTRAGE', '');
  TOBE.PutValue('E_NATURETRESO', '');
  TOBE.PutValue('E_BANQUEPREVI', '');
  TOBE.PutValue('E_QUALIFORIGINE', '');
  TOBE.PutValue('E_CFONBOK', '-');
  TOBE.PutValue('E_NUMORDRE', 0);
  TOBE.PutValue('E_MODESAISIE', '-');
  TOBE.PutValue('E_EQUILIBRE', '-');
  TOBE.PutValue('E_AVOIRRBT', '-');
  TOBE.PutValue('E_PIECETP', '');
  TOBE.PutValue('E_IMMO', '');
  TOBE.PutValue('E_NUMTRAITECHQ', '');
  TOBE.PutValue('E_NUMCFONB', '');
  TOBE.PutValue('E_NUMGROUPEECR', 0);
  TOBE.PutValue('E_CODEACCEPT', 'NON');
  TOBE.PutValue('E_SAISIMP', 0);
  // PT30  TOBE.PutValue('E_SAISIMPEURO', 0);
  TOBE.PutValue('E_PERIODE', GetPeriode(V_PGI.DateEntree));
  TOBE.PutValue('E_SEMAINE', NumSemaine(V_PGI.DateEntree));
  TOBE.PutValue('E_ETATREVISION', '');
  TOBE.PutValue('E_IO', 'X'); // PT57
  TOBE.PutValue('E_PAQUETREVISION', 0);
  TOBE.PutValue('E_REFGESCOM', '');
  TOBE.PutValue('E_REFPAIE', '');
  {Zones libres}
  for i := 0 to 9 do TOBE.PutValue('E_LIBRETEXTE' + IntToStr(i), '');
  for i := 0 to 1 do TOBE.PutValue('E_LIBREBOOL' + IntToStr(i), '-');
  for i := 0 to 3 do
  begin
    TOBE.PutValue('E_TABLE' + IntToStr(i), '');
    TOBE.PutValue('E_LIBREMONTANT' + IntToStr(i), 0);
  end;
  TOBE.PutValue('E_LIBREDATE', iDate1900);
  {Tva Enc}
  for i := 1 to 4 do TOBE.PutValue('E_ECHEENC' + IntToStr(i), 0);
  TOBE.PutValue('E_ECHEDEBIT', 0);
  TOBE.PutValue('E_EDITEETATTVA', '-');
  {Champs particuliers}
  //  PT-2 : 07/09/2001 : V547  PH Suppression maj inexistant dans la TOB
  // TOBE.PutValue('OLDDEBIT',(0))  ; TOBE.PutValue('OLDCREDIT',0) ;

end;

procedure TOF_PG_GeneCompta.RenseigneClefComptaPaie(Etab: string; var MM: RMVT);
begin
  MM.Etabl := Etab;
  if VH_Paie.PGIntegODPaie = 'ECP' then
  begin
    if (JournalPaie.Value = '') and (JournalPaie.text <> '') then MM.Jal := JournalPaie.Text
    else MM.Jal := JournalPaie.Value;
  end
  else MM.Jal := JournalPaie.Value;
  MM.DateC := StrToDate(LBLDF.Caption);
  MM.Exo := QuelExoDT(MM.DateC);
  MM.CodeD := V_PGI.DevisePivot;
  MM.DateTaux := MM.DateC;
  MM.ModeSaisieJal := ModeSaisieJal;
  //@@@@@@@@@@@    MM.ModeOppose := FALSE; // monnaie de tenue
  MM.nature := 'OD';
  MM.TauxD := 1; // devise pivot;
  // PT-1 : 02/08/2001 : V547  PH Ecritures de simulations non validées
  if VH_Paie.PGTypeEcriture then
  begin
    MM.Simul := 'S';
    MM.Valide := FALSE;
  end // Ecriture de simulations
  else
  begin
    MM.Simul := 'N';
    MM.Valide := True;
  end;
  // FIN PT-1


end;

// Fonction qui rempli la TOB paie des ecritures en fonction  de la ligne bulletin

procedure TOF_PG_GeneCompta.RemplirT_Ecr(TypVentil: string; Montant: Double; TypMvt, Libelle: string; T_Ecr, TOB_Generaux: TOB; Compte, CpteAuxi: string;
  LeFolio, LeNumEcrit: Integer);
var
  i: Integer;
  LeMontant, LeMontantF, LeMontantE: Double;
  //  MIdentique: Boolean;
  T_Lecompte: TOB;
  St: string;
  DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double;
  AVentil: array[1..5] of Boolean;
begin
  { Attention aux numero des lignes car s'il y a une décomposition du numéro de compte alors doublon de numéro
  Attention, au numéro de compte auxiliaire sur les comptes 421000....
  PaieEuro
  ComptaEuro
  }
  // il faut verifier les monnaies de tenue pour incrementer les bons montants
//  MIdentique := FALSE;
  DebMD := 0;
  DebMP := 0;
  DebME := 0;
  CreMD := 0;
  CreMP := 0;
  CreME := 0;
  if not PaieEuro then
  begin
    LeMontantF := Montant;
    LeMontantE := FRANCTOEURO(Montant);
  end
  else
  begin
    LeMontantF := EUROTOFRANC(Montant);
    LeMontantE := Montant;
  end;
  LeMontantF := ARRONDI(LeMontantF, V_PGI.OkDecV);
  LeMontantE := ARRONDI(LeMontantE, V_PGI.OkDecV);
  //  if PaieEuro = ComptaEuro then MIdentique := TRUE;
  if TypMvt = 'DEB' then
  begin // Alimentation Debit
    if ComptaEuro then // Compta en EURO
    begin
      DebMD := LeMontantE;
      DebMP := LeMontantE;
      DebME := LeMontantF;
      LeMontant := T_Ecr.GetValeur(iE_DEBIT);
      T_Ecr.PutValeur(iE_DEBIT, ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValeur(iE_DEBITDEV);
      T_Ecr.PutValeur(iE_DEBITDEV, ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      // PT30      LeMontant := T_Ecr.GetValue('E_DEBITEURO');
      // PT30     T_Ecr.PutValue('E_DEBITEURO', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
    end
    else
    begin
      DebMD := LeMontantF;
      DebMP := LeMontantF;
      DebME := LeMontantE;
      LeMontant := T_Ecr.GetValeur(iE_DEBIT);
      T_Ecr.PutValeur(iE_DEBIT, ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValeur(iE_DEBITDEV);
      T_Ecr.PutValeur(iE_DEBITDEV, ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      // PT30      LeMontant := T_Ecr.GetValue('E_DEBITEURO');
      // PT30      T_Ecr.PutValue('E_DEBITEURO', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
    end;
  end
  else // Alimentation Credit
  begin
    if ComptaEuro then // Compta en EURO
    begin
      CreMD := LeMontantE;
      CreMP := LeMontantE;
      CreME := LeMontantF;
      LeMontant := T_Ecr.GetValeur(iE_CREDIT);
      T_Ecr.PutValeur(iE_CREDIT, ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValeur(iE_CREDITDEV);
      T_Ecr.PutValeur(iE_CREDITDEV, ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      // PT30      LeMontant := T_Ecr.GetValue('E_CREDITEURO');
      // PT30      T_Ecr.PutValue('E_CREDITEURO', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
    end
    else
    begin
      CreMD := LeMontantF;
      CreMP := LeMontantF;
      CreME := LeMontantE;
      LeMontant := T_Ecr.GetValeur(iE_CREDIT);
      T_Ecr.PutValeur(iE_CREDIT, ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValeur(iE_CREDITDEV);
      T_Ecr.PutValeur(iE_CREDITDEV, ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      // PT30      LeMontant := T_Ecr.GetValue('E_CREDITEURO');
      // PT30      T_Ecr.PutValue('E_CREDITEURO', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
    end;
  end;
  T_Ecr.PutValeur(iE_GENERAL, Compte);
  T_Ecr.PutValeur(iE_AUXILIAIRE, CpteAuxi);
  T_Ecr.PutValeur(iE_NUMEROPIECE, LeFolio);
  T_Ecr.PutValeur(iE_NUMLIGNE, LeNumEcrit);
  T_Ecr.PutValeur(iE_LIBELLE, Libelle);
  T_Ecr.PutValeur(iE_REFINTERNE, RefPaie);
  //  T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [Compte], FALSE);
  T_Lecompte := RendGene(TOB_Generaux, Compte);
  if not VH_Paie.PGAnalytique then exit; // pas de gestion analytique donc fin procedure
  if (Tob_Ventana = nil) or (TOB_Ventana.Detail.Count <= 0) then exit; // pas de ventilation pour le salarie alors que analytique géré
  if (T_Lecompte <> nil) and (T_LeCompte.GetValue('G_VENTILABLE') = 'X') then
  begin
    for i := 1 to 5 do
    begin
      St := 'G_VENTILABLE' + IntToStr(i);
      AVentil[i] := (T_LeCompte.GetValue(St) = 'X'); // Tableau des Axes Ventilables
    end;
    CumulAnalPaie(TypVentil, Compte, T_Ecr.GetValeur(iE_ETABLISSEMENT), CpteAuxi, IntToStr(LeFolio), IntToStr(LeNumEcrit),
      DebMD, DebMP, DebME, CreMD, CreMP, CreME, AVentil);
  end;

end;
// Fonction qui decline toutes les lignes analytiques en fonction des ventilations trouvees
//   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques format paie PGI alors on conserve la TOB TOBAna

procedure TOF_PG_GeneCompta.EcrPaieVersAna(TOB_Generaux, TOBEcr, TOBAna: TOB; OldP, OldL, NewP, NewL: Integer; Conservation: Boolean = FALSE);
var
  NomEcr, NomAna: string;
  i, j, k, Itrouv, FNumAna, FNumEcr: integer;
  TotalEcriture, TotalDevise, TotalEuro, Pourc: Double;
  LaLigne: T_AnalPaie;
  XDT: T_DetAnalPaie;
  MontantAxe: array[1..5] of double;
  MontantAxeEuro: array[1..5] of double;
  AnaAttente: Boolean;
  NumVentil: Integer;
  TAnal, T_LeCompte: TOB;
  AnalC: TOB;
begin
  //   PT17 : 03/05/02   : V582  PH Pb AGL 540 et disconcordance des champs entre les tables ECRITURE et ANALYTIQ Y_VISION
  FNumAna := PrefixeToNum('Y');
  FNumEcr := PrefixeToNum('E');
  //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques format paie PGI
  if not Conservation then
  begin
    //    T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [TOBEcr.GetValue('E_GENERAL')], FALSE);
    T_LeCompte := RendGene(TOB_Generaux, TOBEcr.GetValeur(iE_GENERAL));
    if (T_Lecompte = nil) or (T_LeCompte.GetValue('G_VENTILABLE') = '-') then exit;
  end;
  for k := 1 to 5 do
  begin
    MontantAxe[k] := 0;
    MontantAxeEuro[k] := 0;
  end;
  TotalEcriture := TOBEcr.GetValeur(iE_DEBIT) + TOBEcr.GetValeur(iE_CREDIT);
  TotalDevise := TOBEcr.GetValeur(iE_DEBITDEV) + TOBEcr.GetValeur(iE_CREDITDEV);
  // PT30  TotalEuro := TOBEcr.GetValue('E_DEBITEURO') + TOBEcr.GetValue('E_CREDITEURO');
  iTrouv := -1;
  for k := 0 to TTA.Count - 1 do // boucle identification de l'écriture
  begin
    LaLigne := T_AnalPaie(TTA[k]);
    if (LaLigne.Cpte = TOBEcr.GetValeur(iE_GENERAL)) and (LaLigne.Etab = TOBEcr.GetValeur(iE_ETABLISSEMENT)) and (LaLigne.Auxi =
      TOBEcr.GetValeur(iE_AUXILIAIRE)) and
      (LaLigne.NumP = IntToStr(OldP)) and (LaLigne.NumL = IntToStr(OldL)) then
    begin
      iTrouv := k;
      Break;
    end;
  end;
  if iTrouv >= 0 then LaLigne := T_AnalPaie(TTA[iTrouv])
  else exit; // PT46 Pas trouvé de ventilation analytique donc on sort
  if LaLigne = nil then exit;
  for k := 0 to LaLigne.Anal.Count - 1 do
  begin
    XDT := T_DetAnalPaie(LaLigne.Anal[k]);
    if (XDT.DebMP = 0) and (XDT.CreMP = 0) then continue; // Pas de ventilation sur axe/section
    TAnal := TOB.Create('ANALYTIQ', TOBAna, -1);
    for i := 1 to TOBEcr.NbChamps do
    begin
      NomEcr := TOBEcr.GetNomChamp(i);
      NomAna := NomEcr;
      NomAna[1] := 'Y';
      j := TAnal.GetNumChamp(NomAna);
      if j > 0 then
      begin
        //   PT17 : 03/05/02   : V582  PH Pb AGL 540 et disconcordance des champs entre les tables ECRITURE et ANALYTIQ Y_VISION
        if V_PGI.DECHAMPS[FNumEcr, i].Tipe = V_PGI.DECHAMPS[FNumAna, j].Tipe
          then TAnal.PutValeur(j, TOBEcr.GetValeur(i));
      end;
    end;
    TAnal.PutValue(' Y_TAUXDEV', 1); // PT50 FQ 11318 On suppose que l'écriture est dans la monnaie de tenue
    TAnal.PutValue('Y_AXE', XDT.Ax);
    TAnal.PutValue('Y_SECTION', XDT.Section);
    TAnal.PutValue('Y_DEBIT', XDT.DebMP);
    TAnal.PutValue('Y_CREDIT', XDT.CreMP);
    TAnal.PutValue('Y_DEBITDEV', XDT.DebMD);
    TAnal.PutValue('Y_CREDITDEV', XDT.CreMD);
    // PT30    TAnal.PutValue('Y_DEBITEURO', XDT.DebME);
    // PT30    TAnal.PutValue('Y_CREDITEURO', XDT.CreME);
    Pourc := Arrondi((XDT.DebMP + XDT.CreMP) / (TotalEcriture) * 100, 6);
    MontantAxe[StrToInt(Copy(XDT.Ax, 2, 1))] := MontantAxe[StrToInt(Copy(XDT.Ax, 2, 1))] + XDT.DebMP + XDT.CreMP;
    MontantAxeEuro[StrToInt(Copy(XDT.Ax, 2, 1))] := MontantAxeEuro[StrToInt(Copy(XDT.Ax, 2, 1))] + XDT.DebME + XDT.CreME;
    TAnal.PutValue('Y_POURCENTAGE', Pourc);
    TAnal.PutValue('Y_TOTALECRITURE', TotalEcriture);
    TAnal.PutValue('Y_TOTALDEVISE', TotalDevise);
    // PT30    TAnal.PutValue('Y_TOTALEURO', TotalEuro);
  end;
  AnaAttente := FALSE;
  for k := 1 to 5 do
  begin
    if (Arrondi(MontantAxe[k] - TotalEcriture, 2) <> 0) then AnaAttente := TRUE;
    // PT30   or (Arrondi(MontantAxeEuro[k] - TotalEuro, 2) <> 0)
  end;
  // Il faut verifier si l'axe est ventilable sinon on passe le tout sur la section d'attente de l'axe
  TotalEuro := 0;
  if AnaAttente then PaieGenereAttente(MontantAxe, MontantAxeEuro, TOB_Generaux, TOBEcr, TOBAna, TotalEcriture, TotalDevise, TotalEuro, Conservation);
  // Intégration des lignes analytiques dans la TOB des lignes afin de faire l'insertion dans la base
  // PGVisuUnObjet (TOBANA, 'Les ecritures analytiques','');
  NumVentil := 1;
  TAnal := TOBAna.FindFirst([''], [''], FALSE);
  while TAnal <> nil do
  begin
    TAnal.PutValue('Y_NUMVENTIL', NumVentil);
    //   PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques format paie PGI alors on conserve la TOB TOBAna
    if not Conservation then
    begin
      if not Assigned(AnalCompte) then AnalCompte := TOB.Create('Mon analytic compte', nil, -1);
      AnalC := TOB.Create('ANALYTIQ', AnalCompte, -1);
      AnalC.Dupliquer(TAnal, TRUE, TRUE, TRUE);
      TAnal.ChangeParent(TOBEcr, -1);
    end;
    NumVentil := NumVEntil + 1;
    TAnal := TOBAna.FindNext([''], [''], FALSE);
  end;
end;
// Fonction qui genere une ligne analytique sur la section attente de chaque axe
// si on a un ecart entre ecriture générale et le total des lignes analytiques associées

procedure TOF_PG_GeneCompta.PaieGenereAttente(MontantAxe, MontantAxeEuro: array of double; TOB_Generaux, TOBEcr, TOBAna: TOB; TotalEcriture, TotalDevise,
  TotalEuro: Double; Conservation: Boolean = FALSE);
var
  k, j, i: Integer;
  NomEcr, NomAna: string;
  TAnal, T_LeCompte: TOB;
  Axe, Sens, st: string;
  Montant: Double; // Pourc
  AVentil: array[1..5] of boolean;
  FNumAna, FNumEcr: integer;
begin
  //   PT17 : 03/05/02   : V582  PH Pb AGL 540 et disconcordance des champs entre les tables ECRITURE et ANALYTIQ Y_VISION
  FNumAna := PrefixeToNum('Y');
  FNumEcr := PrefixeToNum('E');

  //  T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [TOBEcr.GetValue('E_GENERAL')], FALSE);
  T_LeCompte := RendGene(TOB_Generaux, TOBEcr.GetValeur(iE_GENERAL));
  // PT13 : 28/03/02   : V575  PH Traitement des ventilations analytiques format paie PGI
  for i := 1 to 5 do AVentil[i] := FALSE;
  {if Not Conservation then
     begin
     For i := 1 to 5 do AVentil [i] := TRUE;
     end
     else
     begin}
  if ((T_Lecompte <> nil) and (T_LeCompte.GetValue('G_VENTILABLE') = 'X')) then
  begin
    for i := 1 to 5 do
    begin
      St := 'G_VENTILABLE' + IntToStr(i);
      AVentil[i] := (T_LeCompte.GetValue(St) = 'X'); // Tableau des Axes Ventilables
    end;
  end else exit; // compte non ventilable
  //   end;
  for k := 0 to 4 do
  begin
    if not AVentil[K + 1] then continue;
    if (Arrondi(MontantAxe[k] - TotalEcriture, 2) <> 0) then
      // PT30  or (Arrondi(MontantAxeEuro[k] - TotalEuro, 2) <> 0)
    begin
      TAnal := TOB.Create('ANALYTIQ', TOBAna, -1);
      for i := 1 to TOBEcr.NbChamps do
      begin
        NomEcr := TOBEcr.GetNomChamp(i);
        NomAna := NomEcr;
        NomAna[1] := 'Y';
        j := TAnal.GetNumChamp(NomAna);
        if j > 0 then
        begin
          //   PT17 : 03/05/02   : V582  PH Pb AGL 540 et disconcordance des champs entre les tables ECRITURE et ANALYTIQ Y_VISION
          if V_PGI.DECHAMPS[FNumEcr, i].Tipe = V_PGI.DECHAMPS[FNumAna, j].Tipe
            then TAnal.PutValeur(j, TOBEcr.GetValeur(i));
        end;
      end;
      if TOBEcr.GetValeur(iE_DEBIT) <> 0 then Sens := 'D'
      else Sens := 'C';
      Axe := 'A' + IntToStr(k + 1);
      TAnal.PutValue('Y_AXE', Axe);
      TAnal.PutValue('Y_SECTION', VH^.Cpta[AxeToFb(Axe)].Attente);
      Montant := Arrondi(TotalEcriture - MontantAxe[k], 2);
      //      MontantE := Arrondi(TotalEuro - MontantAxeEuro[k], 2);
//      Pourc := Arrondi((Montant) / (TotalEcriture) * 100, 6);
      if Sens = 'D' then
      begin
        TAnal.PutValue('Y_DEBIT', Montant);
        TAnal.PutValue('Y_CREDIT', 0);
        TAnal.PutValue('Y_DEBITDEV', Montant);
        TAnal.PutValue('Y_CREDITDEV', 0);
        // PT30        TAnal.PutValue('Y_DEBITEURO', MontantE);
        // PT30        TAnal.PutValue('Y_CREDITEURO', 0);
      end
      else
      begin
        TAnal.PutValue('Y_DEBIT', 0);
        TAnal.PutValue('Y_CREDIT', Montant);
        TAnal.PutValue('Y_DEBITDEV', 0);
        TAnal.PutValue('Y_CREDITDEV', Montant);
        // PT30        TAnal.PutValue('Y_DEBITEURO', 0);
        // PT30        TAnal.PutValue('Y_CREDITEURO', MontantE);
      end;
      TAnal.PutValue('Y_TOTALECRITURE', TotalEcriture);
      TAnal.PutValue('Y_TOTALDEVISE', TotalDevise);
      // PT30      TAnal.PutValue('Y_TOTALEURO', TotalEuro);
    end;
  end;

end;
// Fonction qui rend le montant de la ligne en fonction de la rubrique et du paramétrage

function TOF_PG_GeneCompta.RendMontantHistoBul(NatureRub, Sens, TypeSalPat: string; Q: TQuery): Double;
var
  Montant: Double;
  TypeSal: string;
begin
  TypeSal := TypeSalPat;
  Montant := 0;
  if NatureRub = 'BAS' then
  begin
    result := 0;
    exit;
  end;
  if NatureRub = 'AAA' then Montant := Q.FindField('PHB_MTREM').AsFloat // Remunération
  else // Cas d'une cotisation
  begin
    if TypeSal = 'MSA' then Montant := Q.FindField('PHB_MTSALARIAL').AsFloat
    else if TypeSal = 'MPT' then Montant := Q.FindField('PHB_MTPATRONAL').AsFloat;
  end;
  // Gestion des arrondis sur chaque ligne traitée, Attention se refère à la compta
  Montant := Arrondi(Montant, NbDecimale);
  // Affection du sens
  if Sens <> 'P' then Result := Montant * -1
  else result := Montant;
  //  Montant := ARRONDI(Montant, V_PGI.OkDecV); ?????????????????????
end;

function TOF_PG_GeneCompta.AlimEcrCompta(LeLib, LeLib1, Etablissement, Nature, Compte, LeSens, TypeSalPat: string; Tob_GuideEcr, Tob_LesEcrit, TOB_Generaux: TOB; Q: TQuery;
  QuelRacineFolio: Integer; QSal: TQuery; TraceErr: TListBox = nil): Boolean;
var
  I, LeFolio, LeNumEcrit, Longueur: Integer;
  LeCompte, CpteAuxi, Etab, TypAlim: string;
  TPR, T_Ecr: TOB;
  MM: RMVT;
  Montant: Double;
  OkOk: Boolean;
  T42: TOB;
  Col42x: Boolean;
begin
  Etab := Etablissement;
  result := FALSE;
  for I := 0 to Tob_GuideEcr.Detail.count - 1 do
  begin
    TPR := Tob_GuideEcr.Detail[I];
    LeCompte := TPR.GetValue('PGC_GENERAL');
    TypAlim := TPR.GetValue('PGC_ALIM421'); // Recup du type Alimentation
    OkOk := FALSE;
    if TypAlim = '' then OkOk := TRUE;
    if not OkOk then
    begin // Recherche identification alimentation identique Soit charges patronales ou salariale
      if (TypAlim = 'SAL') and ((TypeSalPat = 'MSA') or (TypeSalPat = 'SAL')) then OkOk := TRUE;
      if (TypAlim = 'PAT') and ((TypeSalPat = 'MPT') or (TypeSalPat = 'PAT')) then OkOk := TRUE;
    end;
    if OkOk then
    begin // test sur le folio pour bien reprendre les écritures sur le bon folio cas de la ventilations des organismes
      LeFolio := StrToInt(TPR.GetValue('PGC_NUMFOLIO'));
      if QuelRacineFolio <> LeFolio then OkOk := FALSE;
    end;
    //   longueur := 0;
        //  PT11 : 26/03/02   : V575  PH test par rapport à VH_Paie.PGCptNetAPayer
    if MemCmpPaie(LeCompte, Copy(VH_Paie.PGCptNetAPayer, 1, 3), 1, 3) then Longueur := VH_Paie.PGLongRacine421
    else if MemCmpPaie(LeCompte, '4', 1, 1) then Longueur := VH_Paie.PGLongRacin4
    else Longueur := VH_Paie.PGLongRacine;
    if (MemCmpPaie(LeCompte, Compte, 1, longueur)) and (OkoK) then // test sur la racine compte associé à rem ou organisme
    begin
      result := TRUE;
      LeFolio := StrToInt(TPR.GetValue('PGC_NUMFOLIO')); // ==> E_NUMEROPIECE à incrementer en fonction de la compta
      LeNumEcrit := TPR.GetValue('PGC_NUMECRIT'); // ==> E_NUMLIGNE = Numéro de la ligne dans la piece
      Montant := RendMontantHistoBul(Nature, LeSens, TypeSalPat, Q);
      Montant := ARRONDI(Montant, V_PGI.OkdecV);
      if Copy(Compte, 1, 4) = '6412' then
        RESULT := TRUE;
      CpteAuxi := '';

      //     PT21 : 12/08/02   : V585  PH Rajout du compte auxilaire pour les comptes de racine 42 (ex  425)
      Col42x := TRUE;
      //   PT37 : 11/03/2004 : V_500 FQ 11146 PH Prise en compte des comptes collectifs salariés autre que 42 ==> racine classe 4
      if Copy(Compte, 1, 1) = '4' then
      begin
        //        T42 := TOB_Generaux.FindFirst(['G_GENERAL'], [Compte], FALSE);
        T42 := RendGene(TOB_Generaux, Compte);
        if T42 <> nil then Col42x := not (T42.GetValue('G_COLLECTIF') = 'X');
        if not Col42x then
        begin
          if QSal <> nil then CpteAuxi := QSal.FindField('PSA_AUXILIAIRE').AsString;
          if (CpteAuxi = '') and (TraceErr <> nil) then
            TraceErr.Items.Add('Le salarié ' + QSal.FindField('PSA_SALARIE').AsString + ' n''a pas de compte auxiliaire');
          T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_ETABLISSEMENT', 'E_NUMEROPIECE', 'E_NUMLIGNE', 'E_AUXILIAIRE'], [Compte, Etab, LeFolio, LeNumEcrit,
            QSal.FindField('PSA_AUXILIAIRE').AsString], FALSE);
        end
        else T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_ETABLISSEMENT', 'E_NUMEROPIECE', 'E_NUMLIGNE'], [Compte, Etab, LeFolio, LeNumEcrit], FALSE);
      end
      else
        T_Ecr := Tob_LesEcrit.FindFirst(['E_GENERAL', 'E_NUMEROPIECE', 'E_ETABLISSEMENT'], [Compte, LeFolio, Etab], FALSE);
      if T_Ecr = nil then
      begin // Creation de la ligne ecriture
        T_Ecr := TOB.Create('ECRITURE', Tob_LesEcrit, -1);
        InitEcritDefaut(T_Ecr);
        RenseigneClefComptaPaie(Etab, MM);
        LignePaieEcr(MM, T_Ecr);
      end;

      if Nature <> '' then
        RemplirT_Ecr(Nature + ';' + Q.FindField('PHB_RUBRIQUE').AsString, Montant, TPR.GetValue('PGC_TYPMVT'), TPR.GetValue('PGC_LIBELLE'), T_Ecr, TOB_Generaux,
          Compte, CpteAuxi, LeFolio, LeNumEcrit);
      break;
    end; // fin si compte ok
  end; // Fin du for
  //   PT31 : 25/09/03  : V_421 PH Gestion non presence compte et mauvaise affectation
  if (result = FALSE) and (LeLib <> '') then
  begin
    if not Assigned(ErrorGene) then ErrorGene := TStringList.Create;
    if LeLib1 = '' then ErrorGene.Add(LeLib + ' référence le compte ' + Compte + ' inexistant dans le modèle')
    else
    begin
      ErrorGene.Add(LeLib);
      ErrorGene.Add(LeLib1 + ' référence le compte ' + Compte + ' inexistant dans le modèle');
    end;
  end;

end;
// Fonction de tests et lance équilibrage des ecritures

procedure TOF_PG_GeneCompta.EquilibreEcrPaie(TOBEcr: TOB; MM: RMVT; NbDec: integer);
var
  DD, DP, DE, CD, CP, CE: Double;
  i: integer;
  TOBL: TOB;
  EcartE, EcartP, EcartD: Double;
begin
  DD := 0;
  CD := 0;
  DP := 0;
  CP := 0;
  DE := 0;
  CE := 0;
  TobL := nil;
  for i := 0 to TOBEcr.Detail.Count - 1 do
  begin
    TOBL := TOBEcr.Detail[i];
    DD := DD + TOBL.GetValeur(iE_DEBITDEV);
    CD := CD + TOBL.GetValeur(iE_CREDITDEV);
    DP := DP + TOBL.GetValeur(iE_DEBIT);
    CP := CP + TOBL.GetValeur(iE_CREDIT);
    // PT30    DE := DE + TOBL.GetValue('E_DEBITEURO');
    // PT30    CE := CE + TOBL.GetValue('E_CREDITEURO');
  end;
  EcartD := Arrondi(DD - CD, NbDec);
  EcartP := Arrondi(DP - CP, V_PGI.OkDecV);
  EcartE := Arrondi(DE - CE, V_PGI.OkDecE);
  if ((EcartD = 0) and (EcartP = 0) and (EcartE = 0)) then Exit;
  if ((EcartD <> 0) or (EcartE <> 0) or (EcartP <> 0)) then // Ecart Debit/credit donc gestion ligne ecart
  begin
    if TOBL <> nil then
    begin
      TOBL.PutValeur(iE_ENCAISSEMENT, SensEnc(DP, CP)); // determination du code encaissement
      CreerLigneEcartPaie(TOBEcr, TOBL, MM, EcartD, EcartP, EcartE);
    end;
  end;
end;
// Fonction qui créer la ligne écart car on aura tjrs une erreur d'équilibre à cause des arrondis dans la monnaie inversée

function TOF_PG_GeneCompta.CreerLigneEcartPaie(TOBEcr, TOBL: TOB; MM: RMVT; EcartD, EcartP, EcartE: Double): T_ErrCpta;
var
  CptECC: string;
  TOBE: TOB;
  NumL: integer;
  Debit: boolean;
begin
  Result := rcOk;
  //@@@@@@@@@@@
  {
    if EcartP <> 0 then Debit := (EcartP < 0) else Debit := (EcartE < 0);
    if Debit then
    begin
      CptECC := VH^.EccEuroDebit;
      if CptECC = '' then
      begin
        CptECC := VH^.EccEuroCredit;
        EcartP := -EcartP;
        EcartE := -EcartE;
        EcartD := -EcartD;
      end;
    end else
    begin
      CptECC := VH^.EccEuroCredit;
      if CptECC = '' then
      begin
        CptECC := VH^.EccEuroDebit;
        EcartP := -EcartP;
        EcartE := -EcartE;
        EcartD := -EcartD;
      end;
    end;
    // Erreur sur Compte de conversion
    if CptECC = '' then
    begin
      Result := rcPar;
      Exit;
    end;}
    {Ligne d'écriture TOBL contient la dernière TOB=Ligne de l'écriture
    On la duplique pour affecter les nouvelles valeurs concernant l'écart de conversion}
  TOBE := TOB.Create('ECRITURE', TOBEcr, -1);
  TOBE.Dupliquer(TOBL, FALSE, TRUE, TRUE);
  {Général}
  TOBE.PutValeur(iE_GENERAL, CptECC);
  {Divers}
  TOBE.PutValeur(iE_TYPEMVT, 'ECC');
  TOBE.PutValeur(iE_LIBELLE, 'Ecart de conversion');
  NumL := TOBE.GetValeur(iE_NUMLIGNE) + 1;
  TOBE.PutValeur(iE_NUMLIGNE, NumL);
  {Montants}
  if Debit then
  begin
    //   PT-3 : 21/09/2001 : V547  PH Ecart de conversion
    TOBE.PutValeur(iE_DEBIT, -EcartP);
    TOBE.PutValeur(iE_DEBITDEV, -EcartD);
    // PT30    TOBE.PutValue('E_DEBITEURO', -EcartE);
    TOBE.PutValeur(iE_CREDIT, 0);
    TOBE.PutValeur(iE_CREDITDEV, 0);
    // PT30    TOBE.PutValue('E_CREDITEURO', 0);
  end else
  begin
    TOBE.PutValeur(iE_CREDIT, EcartP);
    TOBE.PutValeur(iE_CREDITDEV, EcartD);
    // PT30    TOBE.PutValue('E_CREDITEURO', EcartE);
    TOBE.PutValeur(iE_DEBIT, 0);
    TOBE.PutValeur(iE_DEBITDEV, 0);
    // PT30    TOBE.PutValue('E_DEBITEURO', 0);
        // FIN PT3
  end;


end;
// Fonction qui cumule les debits/credits pour chaque ecriture et par section/axe

procedure TOF_PG_GeneCompta.CumulAnalPaie(TypVentil, Cpte, Etab, Auxi, NumP, NumL: string; DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double; AVentil: array of
  Boolean);
var
  XDT: T_DetAnalPaie;
  TVa, LaTob: TOB;
  Pourc: Double;
  Section, Ax, LSal: string; // Section, Axe
  Nature, Rub, St: string; // Nature et Rubrique
  NatLue, RubLue: string;
  DatD, DatF: TDateTime;
  i, k, iTrouv, NumAxe: integer;
  LaLigne: T_AnalPaie;
  deb, fin: integer;
begin
  if TOB_Ventana.Detail.Count <= 0 then Exit;
  St := TypVentil; // Recup de la nature et de la rubrique traitée
  if TypVentil <> 'S' then
  begin
    Nature := ReadTokenSt(St);
    Rub := ReadTokenSt(St);
  end;
  iTrouv := -1;
  for k := 0 to TTA.Count - 1 do // boucle identification de l'écriture
  begin
    LaLigne := T_AnalPaie(TTA[k]);
    if (LaLigne.Cpte = Cpte) and (LaLigne.Etab = Etab) and (LaLigne.Auxi = Auxi) and
      (LaLigne.NumP = NumP) and (LaLigne.NumL = NumL) then
    begin
      iTrouv := k;
      Break;
    end;
  end;
  if iTrouv < 0 then
  begin
    LaLigne := T_AnalPaie.Create;
    LaLigne.Cpte := Cpte;
    LaLigne.Etab := Etab;
    LaLigne.Auxi := Auxi;
    LaLigne.NumP := NumP;
    LaLigne.NumL := NumL;
    TTA.add(LaLigne);
  end else
  begin
    LaLigne := T_AnalPaie(TTA[iTrouv]);
  end;

  if TypVentil = 'S' then LaTob := Tob_Vensal // Tob des ventilations analytiques salarie cas compte alimenté par BRUT,NETAPAYER ...
  else LaTob := TOB_Ventana; // Tob des ventilations du bulletins rub par rub

  deb := -1;
  fin := -1;
  if Nature = 'COT' then
  begin // boucle de recherche pour trouver si une cotisation a une ventilation spécifique
    for i := 0 to LaTob.Detail.Count - 1 do
    begin
      TVa := LaTob.Detail[i];
      St := TVa.GetValue('YVA_IDENTIFIANT');
      RendRefPaieAnal(LSal, RubLue, NatLue, DatD, DatF, St);
      if (NatLue <> 'C') or ((NatLue = 'C') and (RubLue <> Rub)) then continue;
      if deb = -1 then deb := i;
      fin := i;
    end;
  end;
  if Deb = -1 then Deb := 0;
  if Fin = -1 then fin := LaTob.Detail.Count - 1;

  for i := deb to fin do
  begin
    TVa := LaTob.Detail[i];
    if TypVentil <> 'S' then
    begin
      St := TVa.GetValue('YVA_IDENTIFIANT');
      RendRefPaieAnal(LSal, RubLue, NatLue, DatD, DatF, St);
      if (NatLue[1] <> Copy(Nature, 1, 1)) or
        ((RubLue <> Rub) and (NatLue[1] = 'A')) then continue;
      if (Deb = 0) and (Fin = LaTob.Detail.Count - 1) and (NatLue = 'C') then continue;
    end;
    if TypVentil <> 'S' then
    begin
      Ax := TVa.GetValue('YVA_AXE');
      NumAxe := StrToInt(Copy(Ax, 2, 1));
    end
    else
    begin
      Ax := Copy(TVa.GetValue('V_NATURE'), 2, 2); // recup de A1 dans SA1 ou A2 dans SA2 ...
      NumAxe := StrToInt(Copy(Ax, 2, 1));
    end;
    if not (AVentil[NumAxe - 1]) then continue; // Axe non ventilable
    if TypVentil <> 'S' then
    begin
      Pourc := TVa.GetValue('YVA_POURCENTAGE');
      Section := TVa.GetValue('YVA_SECTION');
    end
    else
    begin
      Pourc := TVa.GetValue('V_TAUXMONTANT');
      Section := TVa.GetValue('V_SECTION');
    end;
    iTrouv := -1;
    for k := 0 to LaLigne.Anal.Count - 1 do
    begin
      XDT := T_DetAnalPaie(LaLigne.Anal[k]);
      if ((XDT.Section = Section) and (XDT.Ax = Ax)) then
      begin
        iTrouv := k;
        Break;
      end;
    end;
    if (iTrouv < 0) or (XDT = nil) then XDT := T_DetAnalPaie.Create;
    XDT.Section := Section;
    XDT.aX := aX;
    XDT.DebMD := ARRONDI(XDT.DebMD + Pourc * DebMD / 100, V_PGI.OkdecV);
    XDT.DebMP := ARRONDI(XDT.DebMP + Pourc * DebMP / 100, V_PGI.OkdecV);
    XDT.DebME := ARRONDI(XDT.DebME + Pourc * DebME / 100, V_PGI.OkdecV);
    XDT.CreMD := ARRONDI(XDT.CreMD + Pourc * CreMD / 100, V_PGI.OkdecV);
    XDT.CreMP := ARRONDI(XDT.CreMP + Pourc * CreMP / 100, V_PGI.OkdecV);
    XDT.CreME := ARRONDI(XDT.CreME + Pourc * CreME / 100, V_PGI.OkdecV);
    if iTrouv < 0 then LaLigne.Anal.Add(XDT);
  end;
end;

{ T_AnalPaie }

constructor T_AnalPaie.Create;
begin
  inherited Create;
  Anal := TList.Create;

end;

destructor T_AnalPaie.Destroy;
begin
  Anal.Free;
  inherited Destroy;

end;
// Fonction de réécriture des paies traitées pour indiquer que la génération comptable a déjà eu lieu

procedure TOF_PG_GeneCompta.ReecritPaiesCompta;
var
  st, st1: string;
begin
// DEB PT60
  if (POS('PSA_', LeWhere) = 0) then // PT63 Test si pas de champ de la table SALARIE car réécriture dans la table PAIEENCOURS impossible sur critères salarie
  begin
    if CritSal <> '' then
    begin
      LeWhere := FindEtReplace(LeWhere, CritSal, ' ', TRUE);
      st1 := ' AND PPU_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE=PPU_SALARIE ' + CritSal + ')';
    end
    else st1 := '';
    St := 'UPDATE PAIEENCOURS SET PPU_TOPGENECR="X" ' + LeWhere + St1;
    ExecuteSQL(St);
  end;
// FIN PT60
  Trace.Items.Add('Mise à jour des paies générées en comptabilité terminée');
  Trace.Refresh;

end;
//   PT28 : 01/04/03   : V_42  PH Fonction qui rend la longueur de la racine
// PT54 Amélioration PT47. si racine du compte non définie alors on prend le
//      compte tel quel.

function TOF_PG_GeneCompta.RendLongueurRacine(LeCompte: string): Integer;
var
  St, LaZone: string;
  ADecompo: Boolean;
begin
  if (Copy(LeCompte, 1, VH_Paie.PGLongRacine421) = VH_Paie.PGCptNetAPayer) then
  begin
    if VH_Paie.PGLongRacine421 > 0 then result := VH_Paie.PGLongRacine421
    else result := Strlen(PChar(LeCompte));
    exit;
  end;
  //   PT34 : 24/02/03 Classe 1 ou 9
  if ((Copy(LeCompte, 1, 1) = '6') or (Copy(LeCompte, 1, 1) = '7') or (Copy(LeCompte, 1, 1) = '9')) then
  begin
    ADecompo := TRUE;
    if VH_Paie.PGDECOMP6 <> '' then
    begin // Cas de la décomposition de quelques racines saisies au préalable
      ADecompo := FALSE;
      LaZone := VH_Paie.PGDECOMP6;
      St := ReadTokenst(LaZone);
      while st <> '' do
      begin
        if Copy(LeCompte, 1, Strlen(PChar(st))) = st then
        begin
          ADecompo := TRUE;
          break;
        end;
        St := ReadTokenst(LaZone);
      end;
    end;
    if ADecompo then result := VH_Paie.PGLongRacine
    else result := Strlen(PChar(LeCompte));
  end
  else
  begin
    //   PT34 : 24/02/03 Classe 1 ou 9
    if ((Copy(LeCompte, 1, 1) = '4') or (Copy(LeCompte, 1, 1) = '2') or (Copy(LeCompte, 1, 1) = '1')) then
    begin
      ADecompo := TRUE;
      if VH_Paie.PGDECOMP4 <> '' then
      begin // Cas de la décomposition de quelques racines saisies au préalable
        ADecompo := FALSE;
        LaZone := VH_Paie.PGDECOMP4;
        St := ReadTokenst(LaZone);
        while st <> '' do
        begin
          if Copy(LeCompte, 1, Strlen(PChar(st))) = st then
          begin
            ADecompo := TRUE;
            break;
          end;
          St := ReadTokenst(LaZone);
        end;
      end;
      if ADecompo then result := VH_Paie.PGLongRacin4
      else result := Strlen(PChar(LeCompte));
    end
      // PT36 FQ 11154 PH V_500 09/03/2004 traitement cas particulier compte dont les logueurs ne sont pas prédéfinies.
    else result := Strlen(PChar(LeCompte));
  end;
end;
// FIN PT54
//   PT42 : 06/05/2004 : V_500 PH FQ 11173 Mise ne place liste d'exportation

procedure TOF_PG_GeneCompta.AnalyseAnalytique(MaMethode: string; TOBAna: TOB; PaieEuro, ComptaEuro: Boolean; NbDec: Integer; Conservation: Boolean; Pan: TPageControl);
var
  MaTob, Ta, TAnalc: TOB;
  i: Integer;
  St, Monnaie: string;
  Aleat: Double;
  Rep: Integer;
begin
  rep := PgiAsk('Voulez vous imprimer les détails de vos écritures analytiques ?', Ecran.Caption);
  if rep <> MrYes then exit;
  try
    if VH_PAIE.PGTenueEuro then Monnaie := 'EUR'
    else Monnaie := 'FRF';
    Aleat := Random(97931);
    if not Conservation then
    begin
      if (TOBANA = nil) or (TOBAna.detail.count - 1 > 0) then
      begin
        FreeAndNil(TOBAna);
        TOBANA := TOB.Create('LES ECRITURES ANA', nil, -1);
      end;
      if Assigned(AnalCompte) then
      begin
        TAnalc := AnalCompte.FindFirst([''], [''], FALSE);
        while TAnalc <> nil do
        begin
          TAnalc.ChangeParent(TOBANA, -1);
          TAnalc := AnalCompte.FindNext([''], [''], FALSE);
        end;
      end;
    end;

    MaTob := TOB.Create('Ma TOB', nil, -1);
    for i := 0 to TOBAna.detail.count - 1 do
    begin
      if TOBAna.detail[i].GetValue('Y_GENERAL') <> '' then
      begin
        Ta := TOB.create('ECODANAPAIE', MaTob, -1);
        Ta.PutValue('PEA_ETABLISSEMENT', TOBAna.detail[i].GetValue('Y_ETABLISSEMENT'));
        Ta.PutValue('PEA_JOURNAL', TOBAna.detail[i].GetValue('Y_JOURNAL'));
        Ta.PutValue('PEA_GENERAL', TOBAna.detail[i].GetValue('Y_GENERAL'));
        Ta.PutValue('PEA_DATECOMPTABLE', TOBAna.detail[i].GetValue('Y_DATECOMPTABLE'));
        Ta.PutValue('PEA_NUMEROPIECE', TOBAna.detail[i].GetValue('Y_NUMEROPIECE'));
        Ta.PutValue('PEA_NUMLIGNE', i + 1);
        Ta.PutValue('PEA_AXE', TOBAna.detail[i].GetValue('Y_AXE'));
        Ta.PutValue('PEA_SECTION', TOBAna.detail[i].GetValue('Y_SECTION'));
        Ta.PutValue('PEA_AUXILIAIRE', TOBAna.detail[i].GetValue('Y_AUXILIAIRE'));
        Ta.PutValue('PEA_LIBELLE', TOBAna.detail[i].GetValue('Y_LIBELLE'));
        Ta.PutValue('PEA_REFINTERNE', TOBAna.detail[i].GetValue('Y_REFINTERNE'));
        if PaieEuro = ComptaEuro then
        begin // on remplit toujours les montants dans la monnaie de tenue de la paie
          Ta.PutValue('PEA_DEBIT', ARRONDI(TOBAna.detail[i].GetValue('Y_DEBIT'), NbDec));
          Ta.PutValue('PEA_CREDIT', ARRONDI(TOBAna.detail[i].GetValue('Y_CREDIT'), NbDec));
        end
        else
        begin
          // PT30          Ta.PutValue('PEA_DEBIT', ARRONDI(TOBAna.detail[i].GetValue('Y_DEBITEURO'), NbDec));
          // PT30          Ta.PutValue('PEA_CREDIT', ARRONDI(TOBAna.detail[i].GetValue('Y_CREDITEURO'), NbDec));
        end;
        Ta.PutValue('PEA_MONNAIE', Monnaie);
        Ta.PutValue('PEA_ALEAT', Aleat);
      end;
    end;
    MaTob.InsertDB(nil, FALSE);
    if MaTob.detail.count - 1 > 0 then
    begin
      st := ' AND PEA_ALEAT = ' + FloatToStr(Aleat);
      LanceEtat('E', 'PAN', 'POA', True, FALSE, False, nil, St, '', False);
      //   PT42 : 06/05/2004 : V_500 PH FQ 11173 Mise ne place liste d'exportation
      if GetControlText('LISTEEXPORT') = 'X' then LanceEtat('E', 'PAN', 'POA', True, True, False, Pan, St, '', False);
    end;
    FreeAndNIL(MaTob);
    ExecuteSql('DELETE FROM ECODANAPAIE WHERE PEA_ALEAT = ' + FloatToStr(Aleat))
  except
    Rollback;
    PGIError('Une erreur est survenue lors de l''édition du journal analytique', 'Génération comptable');
  end;


end;
// DEB PT53

function TOF_PG_GeneCompta.RendGene(LaTobG: TOB; LeCompte: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobG, 'G_GENERAL', LeCompte);

end;

function TOF_PG_GeneCompta.RendCotisation(LaTobC: TOB; LaCot: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobc, 'PCT_RUBRIQUE', LaCot);

end;

function TOF_PG_GeneCompta.RendVentiOrg(LaTobOrg: TOB; org: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobOrg, 'PVO_TYPORGANISME', Org);

end;

function TOF_PG_GeneCompta.RendVentirem(LaTobRem: TOB; Rem: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobRem, 'PVS_RUBRIQUE', Rem);

end;

function TOF_PG_GeneCompta.RendVenticot(LaTobCot: TOB; Cot: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobCot, 'PVT_RUBRIQUE', Cot);

end;

function TOF_PG_GeneCompta.RendTContreP(TContreP: TOB; Etab: string): TOB;
begin
  result := Paie_RechercheOptimise(TContreP, 'ETABLISSEMENT', Etab);

end;

function TOF_PG_GeneCompta.RendTiers(LaTobT: TOB; LeTiers: string): TOB;
begin
  result := Paie_RechercheOptimise(LaTobT, 'T_AUXILIAIRE', LeTiers);

end;


procedure TOF_PG_GeneCompta.SauvLesEcrit(TL, TS: TOB);
var
  T1: TOB;
begin
  T1 := TL.FindFirst([''], [''], FALSE);
  while T1 <> nil do
  begin
    T1.ChangeParent(TS, -1);
    T1 := TL.FindNext([''], [''], FALSE);
  end;
end;

procedure TOF_PG_GeneCompta.MemoriseNumChamp(TT: TOB);
begin
  with TT do
  begin
    iE_DEBIT := GetNumChamp('E_DEBIT');
    iE_CREDIT := GetNumChamp('E_CREDIT');
    iE_DEBITDEV := GetNumChamp('E_DEBITDEV');
    iE_CREDITDEV := GetNumChamp('E_CREDITDEV');
    iE_GENERAL := GetNumChamp('E_GENERAL');
    iE_ETABLISSEMENT := GetNumChamp('E_ETABLISSEMENT');
    iE_CONTREPARTIGEN := GetNumChamp('E_CONTREPARTIGEN');
    iE_AUXILIAIRE := GetNumChamp('E_AUXILIAIRE');
    iE_RIB := GetNumChamp('E_RIB');
    iE_ETATLETTRAGE := GetNumChamp('E_ETATLETTRAGE');
    iE_MODEPAIE := GetNumChamp('E_MODEPAIE');
    iE_ENCAISSEMENT := GetNumChamp('E_ENCAISSEMENT');
    iE_CONFIDENTIEL := GetNumChamp('E_CONFIDENTIEL');
    iE_ECHE := GetNumChamp('E_ECHE');
    iE_NUMECHE := GetNumChamp('E_NUMECHE');
    iE_DATEECHEANCE := GetNumChamp('E_DATEECHEANCE');
    iE_DATECOMPTABLE := GetNumChamp('E_DATECOMPTABLE');
    iE_ANA := GetNumChamp('E_ANA');
    iE_DATEPAQUETMIN := GetNumChamp('E_DATEPAQUETMIN');
    iE_DATEPAQUETMAX := GetNumChamp('E_DATEPAQUETMAX');
    iE_NUMEROPIECE := GetNumChamp('E_NUMEROPIECE');
    iE_NUMLIGNE := GetNumChamp('E_NUMLIGNE');
    iE_NUMORDRE := GetNumChamp('E_NUMORDRE');
    iE_JOURNAL := GetNumChamp('E_JOURNAL');
    iE_LIBELLE := GetNumChamp('E_LIBELLE');
    iE_REFINTERNE := GetNumChamp('E_REFINTERNE');
    iE_TYPEMVT := GetNumChamp('E_TYPEMVT');
    iE_CONTREPARTIEGEN := GetNumChamp('E_CONTREPARTIEGEN');
  end;
end;
// FIN PT53
//*TRA
procedure TOF_PG_GeneCompta.ExportPaieTRA;
var
  LigneCommande : String;
  strIniFileName : String;
  strFichierTRA : String;
  strRapport : String;
  BaseFichier : String;
  Q : Tquery;
  i : integer;
  Etab : String;
  IniFile        :THIniFile;
  Section : String;
  Journal : String;
  Comsxexe : String;
  Stop : Boolean;
  F : TextFile;
  TobEtab : Tob;
  Buffer : String;
  LePath : String;
begin
BaseFichier := DateTimeToStr(now)+'-'+V_PGI.NoDossier+'-'+LblDD.Caption+'au'+LblDF.Caption;
BaseFichier := StringReplace_(BaseFichier,'/','',[rfReplaceAll]);
BaseFichier := StringReplace_(BaseFichier,':','',[rfReplaceAll]);
BaseFichier := StringReplace_(BaseFichier,' ','',[rfReplaceAll]);

//strIniFileName := GetparamSoc('SO_PGCHEMINSAV') + '\COMSXEXPORT'+V_PGI.NoDossier+'.INI';
{PT69
strIniFileName := GetparamSoc('SO_PGCHEMINSAV')+'\'+BaseFichier+'.INI';
}
strIniFileName := VerifieCheminPG (GetparamSoc('SO_PGCHEMINSAV'))+'\'+BaseFichier+'.INI';
//FIN PT69
if strIniFileName = '' then
begin
  PGIInfo('Impossible de déterminer le nom du fichier d''initialisation'+#13#10+'Transfert non effectué', 'Export TRA');
  exit;
end;

//Génération du fichier .ini
IniFile := THIniFile.Create (strIniFileName);
//strRapport := 'RAPPORT-'+V_PGI.NoDossier+'.TXT';
strRapport := BaseFichier+'.TXT';
//strFichierTRA := GetparamSoc('SO_PGCHEMINSAV') + '\PAIE-'+V_PGI.NoDossier+'.TRA';
{PT69
strFichierTRA := GetparamSoc('SO_PGCHEMINSAV') + '\'+BaseFichier+'.TRA';
}
strFichierTRA := VerifieCheminPG (GetparamSoc('SO_PGCHEMINSAV')) + '\'+BaseFichier+'.TRA';
//FIN PT69
Q := OpenSQL('SELECT DISTINCT PPU_ETABLISSEMENT FROM PGGENERECOMPTA '+LeWhere, True);
Etab := '[';
Q.First;
TobEtab := Tob.Create('les etablissements', nil, -1);
TobEtab.LoadDetailFromSQL('SELECT DISTINCT PPU_ETABLISSEMENT FROM PGGENERECOMPTA '+LeWhere);
While not Q.Eof do
begin
  Etab := Etab + Q.findField('PPU_ETABLISSEMENT').AsString + ';';
  Q.Next;
end;
Etab := Etab + ']';
Ferme(Q);
Journal := '[' + GetparamSoc('SO_PGJOURNALPAIE', False) + ';]';

IniFile.WriteString ('COMMANDE', 'NOMFICHIER',        strFichierTRA);
IniFile.WriteString ('COMMANDE', 'NATURETRANSFERT',   'JRL');
IniFile.WriteString ('COMMANDE', 'TRANSFERTVERS',     'S5');
IniFile.WriteString ('COMMANDE', 'FORMAT',            'ETE');
IniFile.WriteString ('COMMANDE', 'EXERCICE',          '');
IniFile.WriteString ('COMMANDE', 'DATEECR1',          LblDD.Caption);
IniFile.WriteString ('COMMANDE', 'DATEECR2',          LblDF.Caption);
IniFile.WriteString ('COMMANDE', 'JOURNAL',           Journal);
IniFile.WriteString ('COMMANDE', 'DATEARRET',         '');
IniFile.WriteString ('COMMANDE', 'ETABLISSEMENT',     Etab);
//bug de comsx si nature=od alors il sort TOUS les journaux OD
//IniFile.WriteString ('COMMANDE', 'NATUREJRL',         'OD');
IniFile.WriteString ('COMMANDE', 'EXCLURE',           'DEV;REG;SOU;CORR;TL;REL;EXO;PARAM;');
IniFile.WriteString ('COMMANDE', 'MAIL',              '');
IniFile.WriteString ('COMMANDE', 'TYPE',              '[N;S]');
IniFile.WriteString ('COMMANDE', 'RAPPORT',           strRapport);
IniFile.Free;

//Génération ligne de commande pour ComSx
LePath := TCbpPath.GetCegidDistriApp; // PT68
Comsxexe := LePath+'\comSx.exe';
if ((V_PGI.ModePCL='1')) then
  Section := 'DB'+V_PGI.NoDossier+'@'+V_PGI.CurrentAlias
else
  Section := V_PGI.CurrentAlias;

LigneCommande := ' "/USER='+V_PGI.UserLogin+'" /PASSWORD='+V_PGI.PassWord+' /DOSSIER=' + Section + ' /INI=' + strIniFileName + ';EXPORT;Minimized"';

if not FileExists(Comsxexe) then
begin
  PGIInfo('Le programme '+comsxexe+' est introuvable');
  exit;
end;

Stop := False;
//ShellExecute (0, pchar ('open'), pchar (comsxexe), pchar (lignecommande), nil, SW_RESTORE);
if FileExecAndWaitUntil(Comsxexe + LigneCommande, 15, Stop) then
  begin
  for i:=0 to TobEtab.Detail.Count-1 do
  begin
    Q := OpenSQL('SELECT ET_ETABLISSEMENT, ET_LIBELLE FROM ETABLISS WHERE ET_ETABLISSEMENT="'+TobEtab.Detail[i].GetValue('PPU_ETABLISSEMENT')+'"',True);
    TobEtab.Detail[i].AddChampSup('RAISONSOCIALE',False);
    TobEtab.Detail[i].PutValue('RAISONSOCIALE', Q.FindField('ET_LIBELLE').AsString);
    Ferme(Q);
  end;

{PT69
  assignfile(F, GetparamSoc('SO_PGCHEMINSAV') + '\' + strRapport);
}
  assignfile(F, VerifieCheminPG (GetparamSoc('SO_PGCHEMINSAV')) + '\' + strRapport);
//FIN PT69
  Append(F);
  writeln(F, '___________________________________');
  writeln(F, 'Liste des établissements :');
  For i:=0 to TobEtab.Detail.Count-1 do
  begin
    Buffer := '    * '+TobEtab.Detail[i].GetValue('PPU_ETABLISSEMENT')+'  '+TobEtab.Detail[i].GetValue('RAISONSOCIALE');
    writeln(F, Buffer);
  end;

  closefile(F);


  PGIBox('L''exportation TRA s''est bien déroulée.#10#13'+
         'Le fichier se localise dans : '+strFichierTRA,'Export TRA');
  end
else
  PGIInfo('Erreur lors de l''export TRA');


DeleteFile(strIniFileName);



end;

//DEB PT72
procedure TOF_PG_GeneCompta.BImprimerClik(Sender: TObject);
var
  TobEdit, TOBE : Tob;
  ListeErr : TListBox;
  i : integer;
  Titre : String;
  Cbx : THValComboBox;
  Ficlog : String;
  FRapport: TextFile;
begin
  ListeErr := TListbox(Getcontrol('LSTBXERROR'));
  if (ListeErr.Count <> 0) then
  begin
    {$IFNDEF EAGLCLIENT}
    Ficlog := V_PGI.DatPath + '\RapportGeneCompta.log';
    {$ELSE}
    Ficlog := VH_Paie.PGCheminEagl + '\RapportGeneCompta.log';
    {$ENDIF}
    AssignFile(FRapport, Ficlog);
    if FileExists(Ficlog) then
      Append(FRapport)
    else
      ReWrite(FRapport);

    TobEdit := TOB.Create('Leserreurs',nil,-1);

    Titre := 'Liste d''erreurs de la génération comptable et analytique';
    Writeln(FRapport,Titre);

    TOBE := TOB.Create('Leserreursdetail',TobEdit,-1);
    TOBE.AddChampSupValeur('TYPE','E',false);
    TOBE.AddChampSupValeur('LIGNE','Traitement des paies du ' + GetControlText('LBLDU') + ' au ' + GetControlText('LBLAU'));
    Writeln(FRapport, 'Traitement des paies du ' + GetControlText('LBLDU') + ' au ' + GetControlText('LBLAU'));
    TOBE := TOB.Create('Leserreursdetail',TobEdit,-1);
    TOBE.AddChampSupValeur('TYPE','E',false);
    TOBE.AddChampSupValeur('LIGNE','Date des écritures comptables : ' + GetControlText('DATEFIN'));
    Writeln(FRapport,'Date des écritures comptables : ' + GetControlText('DATEFIN'));
    Cbx := THValCombobox(GetControl('CBXJNAL'));
    if Cbx <> nil then
    begin
      TOBE := TOB.Create('Leserreursdetail',TobEdit,-1);
      TOBE.AddChampSupValeur('TYPE','E',false);
      TOBE.AddChampSupValeur('LIGNE','Journal OD de paie : ' + Cbx.Text);
      Writeln(FRapport,'Journal OD de paie : ' + Cbx.Text);
    end;
    Cbx := THValCombobox(GetControl('CBXJEU'));
    if Cbx <> nil then
    begin
      TOBE := TOB.Create('Leserreursdetail',TobEdit,-1);
      TOBE.AddChampSupValeur('TYPE','E',false);
      TOBE.AddChampSupValeur('LIGNE','Modèle d''écritures : ' + Cbx.Text);
      Writeln(FRapport,'Modèle d''écritures : ' + Cbx.Text);
    end;

    For i := 0 to ListeErr.Count -1 do
    begin
      TOBE := TOB.Create('Leserreursdetail',TobEdit,-1);
      TOBE.AddChampSupValeur('TYPE','D',false);
      TOBE.AddChampSupValeur('LIGNE',ListeErr.Items[i],false);
      Writeln(FRapport,ListeErr.Items[i],false);
    end;
    LanceEtatTob('E','PAY','GLS',TobEdit,True,False,False,TPageControl(GetControl('Pages')),'',Titre,False);
    FreeAndNil(TobEdit);

    PGIBox('Le fichier log ' + Ficlog + ' a été créé', 'Génération comptable et analytique');
    CloseFile(FRapport);
  end
  else
    PGIBox('Il n''y a aucune erreur à imprimer');
end;
//FIN PT72

initialization
  registerclasses([TOF_PG_GeneCompta]);
end.


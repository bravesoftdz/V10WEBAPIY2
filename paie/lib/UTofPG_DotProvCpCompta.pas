{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. : Unit de génération des ODs de paie de constatation
Suite ........ : provisions congés payes
Mots clefs ... : PAIE;
*****************************************************************}
{ Traitement des ecritures comptables constatant la provision pour dotation congés payés
}

{
PT1  : 06/05/2004 : V_500 PH Mise ne place liste d'exportation
PT2  : 31/08/2004 : V_500 PH FQ 11557 Correction erreur SQL

==> 15/10/2004 Réalignement du source avec UTofPG_GeneCompta concernant les améliorations
    et la correction des FQ
PT3  : 22/12/2004 : V_600 PH Borne inférieure établissement prise en compte
PT4  : 01/03/2005 : V_602 PH FQ 12052 Prise en compte des décompositions idem ODs de paie
PT5  : 29/09/2005 : V_602 PH FQ 12599 ergonomie message d'anomalie
PT6  : 13/04/2006 : V_650 PH FQ 12843 Suppression du controle des ventilations analytiques si on ne gère pas l'analytique

}

unit UTofPG_DotProvCpCompta;

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
  UTOF,
  UTOB,
  SaisUtil,
  UtilSais,
  //      UTOM,
  Vierge,
  AGLInit,
  UtilPgi,
  ed_tools;

type
  T_ErrCpta = (rcOk, rcRef, rcPar);

type
  T_AnalPaie = class
    Cpte, Etab, Auxi, NumP, NumL: string; // compte,etablissement,Auxiliaire,numero de piece, numero de ligne
    Anal: TList; // liste des decompositions analytiques valorisees
    constructor Create;
    destructor Destroy; override;
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


type
  TOF_PG_DotProvCpCompta = class(TOF)
  private
    LblDD: THLabel; // Date debut, fin des paies traitées et par défaut de la date des ecritures paie
    BtnLance: TToolbarButton97;
    JeuPaie, JournalPaie: THValComboBox;
    QMul: TQUERY; // Query recuperee du mul
    Trace, TraceErr: TListBox;
    ErrorGene: TStringList;
    NbDecimale: Integer; // Nbre de décimale de la compta
    ModeSaisieJal: string; // Mode de saisie du journal
    Etabl, Method: string; // Etablissement sur lequel on fait les OD de paie
    RefPaie: string;
    TTA: TList; // liste des ventilations (cumuls analytiques)
    Tob_Ventana, Tob_Vensal: TOB; // Tob des ventilations analytiques du salarie
    AnalCompte: TOB; // Tob qui contient les ventilations analytiques pour un compte
    ComptaEuro, PaieEuro: Boolean; // tenue euro des differents modules
    procedure LanceDotProvCpCompta(Sender: TObject);
    procedure InitEcritDefaut(TOBE: TOB); // Initialisation TOB ecriture avec ts les champs remplis par defaut
    procedure LignePaieECR(MM: RMVT; TOBE: TOB); // Rempli une ligne ecriture en fonction de la ligne du bulletin
    procedure RenseigneClefComptaPaie(Etab: string; var MM: RMVT);
    procedure RemplirT_ecr(TypVentil: string; Montant: Double; TypMvt, Libelle: string; T_Ecr, TOB_Generaux: TOB; Compte, CpteAuxi: string; LeFolio,
      LeNumEcrit: Integer); // Rempli une tob paie preligne ecriture comptable avant affectation définitive numéropiece ...
    procedure EquilibreEcrPaie(TOBEcr: TOB; MM: RMVT; NbDec: integer);
    function CreerLigneEcartPaie(TOBEcr, TOBL: TOB; MM: RMVT; EcartD, EcartP, EcartE: Double): T_ErrCpta;
    procedure CumulAnalPaie(TypVentil, Cpte, Etab, Auxi, NumP, NumL: string; DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double; AVentil: array of Boolean);
    procedure EcrPaieVersAna(TOB_Generaux, TOBEcr, TOBAna: TOB; OldP, OldL, NewP, NewL: Integer; Conservation: Boolean = FALSE);
    procedure PaieGenereAttente(MontantAxe, MontantAxeEuro: array of double; TOB_Generaux, TOBEcr, TOBAna: TOB; TotalEcriture, TotalDevise, TotalEuro: Double;
      Conservation: Boolean = FALSE);
    function RendLongueurRacine(LeCompte: string): Integer;
    procedure AnalyseAnalytique(MaMethode: string; TOBAna: TOB; PaieEuro, ComptaEuro: Boolean; NbDec: Integer; Conservation: Boolean; Pan: TPageControl);
    function ChargeVentileLigAna(Salarie: string; DateDebut, DateFin: TDateTime): TOB;

  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation

uses SaisBul,
  PgCongesPayes,
  P5Util,
  P5Def,
  PgOutils,
  PgOutils2,
  UTofPG_GeneCompta,
  PgVisuObjet;

{ DEB PT4 Mise en commentaire des fonctions car on va utiliser celles des UtofPg_GeneCompta
de façon à décomposer de la même façon les comptes pour la génération des ODs de paie et
les dotations pour CP et RTT

function LeChampSalDot(LeTypChamp: string; Q: TQUERY; TheTob: TOB): string;
var
  LeChamp, Prefixe: string;
  LaDate: TdateTime;
  JJ, MM, AA: WORD;
  LeMois: string;
begin
  Result := '';
  LeChamp := '';
  if LeTypChamp = 'DAT' then
  begin
    if Q <> nil then LaDate := Q.FindField('PDC_DATEARRET').AsDateTime
    else if TheTob <> nil then LaDate := TheTob.GetValue('PDC_DATEARRET');
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
  if LeTypChamp = 'ETB' then Lechamp := 'PDC_ETABLISSEMENT';
  if LeTypChamp = 'OR1' then Lechamp := 'PDC_TRAVAILN1';
  if LeTypChamp = 'OR2' then Lechamp := 'PDC_TRAVAILN2';
  if LeTypChamp = 'OR3' then Lechamp := 'PDC_TRAVAILN3';
  if LeTypChamp = 'OR4' then Lechamp := 'PDC_TRAVAILN4';
  if LeTypChamp = 'STA' then Lechamp := 'PDC_CODESTAT';
  if LeTypChamp = 'TC1' then Lechamp := 'PDC_LIBREPCMB1';
  if LeTypChamp = 'TC2' then Lechamp := 'PDC_LIBREPCMB2';
  if LeTypChamp = 'TC3' then Lechamp := 'PDC_LIBREPCMB3';
  if LeTypChamp = 'TC4' then Lechamp := 'PDC_LIBREPCMB4';

  if LeChamp <> '' then
  begin
    if Q <> nil then Result := Q.FindField(Lechamp).AsString
    else if TheTob <> nil then result := TheTob.GetValue(LeChamp);
  end;
end;
// Reprise du PT47 pour alignement avec la gestion des Ods paie: Decomposition des racines en fonction de la sélection
function ConfectionCpteDot(LaRacine, TypeQ: string; Q: TQUERY; C421: Boolean; TheTob: TOB; ZTraceErr: TListBox = nil): string;
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
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
      end;
      if VH_Paie.PGNbreRac2 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPreselect2, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbreRac2, VH_Paie.PGNbreRac2)
        else if Q <> nil then
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
      end;
      if VH_Paie.PGNbreRac3 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPreselect3, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbreRac3, VH_Paie.PGNbreRac3)
        else if Q <> nil then
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
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
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 1 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
      end;
      if VH_Paie.PGNbre4Rac2 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPre4select2, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbre4Rac2, VH_Paie.PGNbre4Rac2)
        else if Q <> nil then
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 2 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
      end;
      if VH_Paie.PGNbre4Rac3 <> 0 then
      begin
        LePlus := LeChampSal(VH_Paie.PGPre4select3, TypeQ, Q, TheTob);
        if LePlus <> '' then LeCompte := LeCompte + Copy(LePlus, 4 - VH_Paie.PGNbre4Rac3, VH_Paie.PGNbre4Rac3)
        else if Q <> nil then
          if ZTraceErr <> nil then ZTraceErr.Items.Add('Erreur 3 Confection compte ' + Lecompte + ' pour le salarié ' + Q.FindField('PHB_SALARIE').AsString);
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

procedure ZZCompleteCompte(var LeCompte: string);
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
// FIN PT4
}
// Fonction principale de traitement de la génération des ecritures comptables

procedure TOF_PG_DotProvCpCompta.LanceDotProvCpCompta(Sender: TObject);
var
  Salarie, Nature, Organisme, Nodossier, Rubrique, TypeSalPat, Metab, MDR_Def, TheAuxi: string;
  Pan: TPageControl;
  Tbsht: TTabSheet;
  QuelRacine, LeFolio, LeFolio1, LeFolio2, LeNumEcrit, LeNumEcrit1, LeNumEcrit2, i, OkOk: Integer;
  ConservAna, EcrPositif: Boolean;
  DateS, ZZDate1, ZZDate2: TDateTime;
  SystemTime0: TSystemTime;
  valid, Compte641, Col421, Col42x, OkTrouve, OkTrouve6, Lettrable, LettrableAux, ASauter: boolean;
  St, St1, TheRacine, Racine, Compte, Compte6, LeCompte, LeSens, CpteAuxi, sRib: string;
  Tob_GuideEcr, Tob_HistoBulletin: TOB; // TOB correspondant à des tables
  Torg, Tob_LesEcrit, T_Ecr, TPR, TOB_Generaux: TOB; // Tob de  travail;
  TEcritPaie, TOD, TGene, TContreP, TCP, TOBANA, TZZ, T42, MaTOB, Ta, TModeRegl, TRegl: TOB;
  T_Etab, Tetb: TOB; // PT38 Tob des etablissements
  Q, QJal, QSal: TQuery;
  MtD1, MtD2, MtC1, MtC2, SalBrut, SalNet, SalCharges, Ecart, Ecart1, TotDebit, TotCredit, TotDebit2, TotCredit2, Montant: double;
  // Montants debit/credit folio 1 et 2
  MM: RMVT;
  NumEcritOD1, NumEcritOD2, NumFolioOD1, NumFolioOD2, NBrePiece, NbDec: Integer; // Renumerotation des numeros de lignes dans le Folio
  Prem421, PremCpt, IntegODPaie: string; // 1er Compte et 1er compte 421 trouvés pour la gestion des contre parties
  ContreP: T_ContreP; // Gestion des contreparties ==> pour memoriser lres 1er compte et 1er compte 421
  OldP, OldL, NewP, NewL, rep: Integer;
  TheTrace: TStringList;
  LgRac, ij: Integer;
  Aleat: Double;
  TheModRegl, NatGene, GLet, GPoint, LeLib, LeLib1, TypDot, TypVentil: string;
begin
  Salarie := '';
  Etab := '';
  NBrePiece := 1;
  PremCpt := '';
  RefPaie := 'Provisions congés payés au ' + LblDD.Caption;
  // On regarde quelles sont les monnaies de tenue de la paie et de la compta
  PaieEuro := VH_Paie.PGTenueEuro; // Monnaie de tenue de la paie
  ComptaEuro := VH^.TenueEuro; // Monnaie de tenue de la compta
  if not ComptaEuro and PaieEuro then ComptaEuro := PaieEuro;
  IntegODPaie := VH_Paie.PGIntegODPaie; // Méthode d'intégration de la paie
  NbDecimale := V_PGI.OkdecV;
  if (IntegODPaie = 'ECP') then
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
    if (IntegODPaie <> 'IMM') and (IntegODPaie <> 'DIF') then
    begin
      PGIBox('La génération ne peut pas être lancée', 'La méthode de comptabilisation dans les paramètres société n''est pas validée !');
      exit;
    end;
  end;
  //  test si les compteurs des journaux existent
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

  st := 'SELECT ETB_ETABLISSEMENT,ETB_RIBSALAIRE FROM ETABCOMPL';
  Q := OpenSql(St, TRUE);
  T_Etab := TOB.Create('Mes etablissements', nil, -1);
  T_Etab.LoadDetailDB('ETABSOCIAL', '', '', Q, False);
  ferme(Q);

  // On regarde si le compte 421000 est un compte collectif pour lui mettre un auxiliaire
  St :=
    'SELECT G_GENERAL,G_LIBELLE,G_NATUREGENE,G_CENTRALISABLE,G_COLLECTIF,G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_VENTILABLE,G_LETTRABLE,G_POINTABLE,G_MODEREGLE FROM GENERAUX ORDER BY G_GENERAL';
  Q := OpenSql(st, TRUE);
  // Chargement de la TOB des comptes généraux pour avoir les caractéristiques de chaque compte général
  TOB_Generaux := TOB.Create('Les comptes généraux', nil, -1);
  TOB_Generaux.LoadDetailDB('GENERAUX', '', '', Q, False);
  ferme(Q);

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
  //  Test si ecriture de simulation et journal pas de type piece alors erreur
  if VH_PAie.PgTypeEcriture then // Ecriture de simulation
  begin
    if ModeSaisieJal <> '-' then // si journal <> piece alors erreur
    begin
      PGIBox('La génération ne peut pas être lancée car vous ne pouvez pas avoir #13#10 des écritures de simulation avec un journal de type bordereau ou libre', ECran.Caption);
      exit;
    end;
  end;

  if (JeuPaie = nil) or (JeuPaie.Value = '') then
  begin
    PGIBox('La génération ne peut pas être lancée, #13#10car vous n''avez pas sélectionné de modèle d''écritures', Ecran.caption); // PT5
    exit;
  end;
  St := 'SELECT * FROM JEUECRPAIE WHERE ##PJP_PREDEFINI## PJP_NOJEU =' + JeuPaie.Value + ''; // DB2
  Q := OpenSql(st, TRUE);
  if not Q.EOF then TypDot := Q.FindField('PJP_TYPEPROVCP').AsString;
  FERME(Q);
  if TypDot = 'COD' then
  begin
    PGIBox('Le modèle d''écritures ne correspond à une écriture de provisions CP ou RTT', Ecran.Caption);
    exit;
  end;

  BtnLance.Enabled := FALSE;
  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  DateD := StrToDate(LblDD.Caption);
  DateF := DateD;
  // Chargement des TOB pour la génération des ecritures
  Tob_GuideEcr := TOB.Create('Le Guide des ecritures', nil, -1);
  St := 'SELECT * FROM GUIDEECRPAIE WHERE ##PGC_PREDEFINI## PGC_JEUECR=' + JeuPaie.Value + // DB2
    ' ORDER BY PGC_PREDEFINI,PGC_NODOSSIER,PGC_JEUECR,PGC_NUMFOLIO,PGC_NUMECRIT'; // DB2
  Q := OpenSql(st, TRUE);
  Tob_GuideEcr.LoadDetailDB('GUIDEECRPAIE', '', '', Q, False);
  Ferme(Q);
  // Recherche si les comptes 421 et/ou 641 sont utilisés dans le guide
  Compte641 := FALSE;
  for I := 0 to Tob_GuideEcr.Detail.count - 1 do
  begin
    TPR := Tob_GuideEcr.Detail[I];
    LeCompte := TPR.GetValue('PGC_GENERAL');
    if (Copy(LeCompte, 1, 3) = '641') and (TPR.GetValue('PGC_ALIM421') <> '') then Compte641 := TRUE;
  end;
  Tob_LesEcrit := TOB.Create('Les ecritures vues de la paie', nil, -1);
  // Chargement des TOB des differents types de ventilations
  Trace.Items.Add('Fin de Chargement des informations nécessaires');
  if QMul = nil then
  begin
    PGIError('Erreur sur la liste des salariés', 'Abandon du traitement');
    exit;
  end;
  QMul.First;

  GetLocalTime(SystemTime0);
  Trace.Items.Add('Début du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  Trace.Refresh;
  while not QMul.EOF do
  begin
    Salarie := QMul.FindField('PDC_SALARIE').AsString;
    St := 'Select PSA_SALARIE,PSA_AUXILIAIRE,PSA_LIBELLE FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
    QSal := OpenSql(St, TRUE);
    if QSal.EOF then
    begin
      TraceErr.Items.Add('Erreur le salarié ' + Salarie + ' est inconnu');
      QMul.Next;
      Ferme(QSal);
      continue;
    end;

    Etab := QMul.FindField('PDC_ETABLISSEMENT').AsString;
    if (Etab <> Etabl) and (Etabl <> '') and (Method = 'MON') then // cas Mono établissement
    begin // cas ou le salarié n'appartient pas au moment de la paie à l'établissement selectionné
      QMul.Next;
      continue;
    end;

    if Method = 'PRI' then Etab := VH^.EtablisDefaut;
    ZZDate1 := QMul.FindField('PDC_DATEDEBUT').AsDateTime;
    ZZDate2 := QMul.FindField('PDC_DATEFIN').AsDateTime;
    if (VH_Paie.PGAnalytique = TRUE) then
    begin // chargement des ventilations analytiques au niveau de chaque salarié/bulletin
      if Tob_Ventana <> nil then
      begin
        Tob_Ventana.Free;
        Tob_Ventana := nil;
      end;
      Tob_Ventana := ChargeVentileLigAna(Salarie, ZZDate1, ZZDate2);

      if TOB_VenSal <> nil then
      begin
        TOB_VenSal.Free;
        TOB_VenSal := nil;
      end;
      TOB_VenSal := TOB.Create('Analytique Salarié', nil, -1);
      St := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "SA%" AND V_COMPTE="' + Salarie + '" ORDER BY V_NATURE,V_COMPTE';
      Q := OpenSql(st, TRUE);
      TOB_VenSal.LoadDetailDB('VENTIL', '', '', Q, FALSE);
      ferme(Q);
      // DEB PT6
      if ((Tob_Ventana = nil) or (TOB_Ventana.Detail.Count <= 0)) and
        ((TOB_VenSal = nil) or (TOB_VenSal.Detail.Count <= 0)) then
        TraceErr.Items.Add('Erreur le salarié ' + Salarie + ' n''a pas de ventilation analytique');
     // FIN PT6
    end;
    //

    // Récupération des bons champs en foncion du type de dotations pour provisions CP ou RTT
    if TypDot = 'CPP' then
    begin
      SalBrut := QMul.FindField('PDC_CPPROVISION').AsFloat;
      SalCharges := QMul.FindField('PDC_CPCHARGESPAT').AsFloat;
    end
    else
    begin
      SalBrut := QMul.FindField('PDC_PROVRTT').AsFloat;
      SalCharges := QMul.FindField('PDC_RTTCHARGESPAT').AsFloat;
    end;

    if SalBrut = 0 then
    begin // Si pas de provision alors on passe au salarié suivant
      QMul.Next;
      Ferme(QSal);
      continue;
    end;

    for I := 0 to Tob_GuideEcr.Detail.count - 1 do
    begin
      TPR := Tob_GuideEcr.Detail[I];
      LeCompte := TPR.GetValue('PGC_GENERAL');
      Col42x := TRUE;
      LeFolio := StrToInt(TPR.GetValue('PGC_NUMFOLIO')); // ==> E_NUMEROPIECE à incrementer en fonction de la compta
      LeNumEcrit := TPR.GetValue('PGC_NUMECRIT'); // ==> E_NUMLIGNE = Numéro de la ligne dans la piece
      // A rajouter le traitement du compte Collectif à savoir ==> recup compte auxiliaire
      if (MemCmpPaie(LeCompte, Copy(VH_Paie.PGCptNetAPayer, 1, VH_Paie.PGLongRacine421), 1, VH_Paie.PGLongRacine421)) then
      begin
        TheRacine := Copy(LeCompte, 1, VH_Paie.PGLongRacine421);
        Compte := ConfectionCpte(TheRacine, 'D', QMul, TRUE, nil, TraceErr);
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
        LgRac := RendLongueurRacine(LeCompte);
        TheRacine := Copy(LeCompte, 1, LgRac);
        if (Copy(TheRacine, 1, 3) = '512') then
        begin
          Tetb := T_Etab.FindFirst(['ETB_ETABLISSEMENT'], [QMul.FindField('PPU_ETABLISSEMENT').AsString], FALSE);
          if Tetb <> nil then
            if (Tetb.GetValue('ETB_RIBSALAIRE') <> '') then TheRacine := Tetb.GetValue('ETB_RIBSALAIRE')
        end;
        Compte := ConfectionCpte(TheRacine, 'D', QMul, FALSE, nil, TraceErr);
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
      else if Nature = 'PAT' then Montant := SalCharges
      else Montant := 0;
      CpteAuxi := '';
      if Nature <> '' then
      begin
        if (Tob_Ventana <> nil) and (TOB_Ventana.Detail.Count > 0) then TypVentil := 'COT'
        else TypVentil := 'S';
        RemplirT_Ecr(TypVentil, Montant, TPR.GetValue('PGC_TYPMVT'), TPR.GetValue('PGC_LIBELLE'), T_Ecr, TOB_Generaux, Compte, CpteAuxi, LeFolio, LeNumEcrit);
      end;
    end; // Fin de la boucle sur le guide des ecritures

    Trace.Items.Add('Salarié ' + Salarie + ' ' +
      ' Etablissement : ' + RechDom('TTETABLISSEMENT', QMul.FindField('PDC_ETABLISSEMENT').AsString, FALSE) + ' ');
    Trace.Refresh;
    if QSal <> nil then Ferme(QSal);
    QMul.NEXT;
  end; // Fin boucle sur la Query du mul
  //  PGVisuUnObjet (Tob_LesEcrit, 'Les ODs de paie', 'E_ETABLISSEMENT;E_NUMEROPIECE;E_NUMLIGNE;E_GENERAL;E_AUXILIAIRE;E_DEBIT;E_CREDIT') ;
  Tob_LesEcrit.Detail.Sort('E_ETABLISSEMENT;E_NUMEROPIECE;E_NUMLIGNE');
  // confection de la TOB des lignes des ecritures de paie
  TEcritPaie := TOB.create('Les od de la paie', nil, -1);

  TotDebit := 0;
  TotCredit := 0;
  TotDebit2 := 0;
  TotCredit2 := 0;

  // Boucle pour éliminer les lignes vides le cas échéant et memorisation des 2 comptes pour la gestion des contre parties
  ASauter := FALSE;
  Tob_LesEcrit.Detail.Sort('E_ETABLISSEMENT;E_NUMEROPIECE;E_NUMLIGNE');
  MEtab := '';

  if Method = 'MUL' then TContreP := TOB.Create('Les contre parties par établissement', nil, -1);

  TPR := Tob_LesEcrit.FindFirst([''], [''], FALSE);
  while TPR <> nil do
  begin
    ASauter := FALSE;

    if (TPR.GetValue('E_DEBIT') = 0) and (TPR.GetValue('E_CREDIT') = 0) then
    begin
      TPR.free;
      ASauter := TRUE;
    end;
    if not ASauter then
    begin
      if Method <> 'MUL' then
      begin
        if (Prem421 = '') and (Copy(TPR.GetValue('E_GENERAL'), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then Prem421 := TPR.GetValue('E_GENERAL');
        if (PremCpt = '') and (Copy(TPR.GetValue('E_GENERAL'), 1, 3) <> Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then PremCpt := TPR.GetValue('E_GENERAL');
      end
      else
      begin
        st := TPR.GetValue('E_ETABLISSEMENT');
        if (st <> '') then
        begin
          TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
          if TCP = nil then
          begin
            TCP := TOB.create('Une ligne', TContreP, -1);
            TCP.AddChampSup('ETABLISSEMENT', FALSE);
            TCP.PutValue('ETABLISSEMENT', TPR.GetValue('E_ETABLISSEMENT'));
            TCP.AddChampSup('PREM421', FALSE);
            TCP.AddChampSup('FOLIO1', FALSE);
            TCP.AddChampSup('FOLIO2', FALSE);
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

          if (Copy(TPR.GetValue('E_GENERAL'), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TCP.GetValue('PREM421') = '') then
            TCP.PutValue('PREM421', TPR.GetValue('E_GENERAL'));
          if (Copy(TPR.GetValue('E_GENERAL'), 1, 3) <> Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TCP.GetValue('PREMCPT') = '') then
            TCP.PutValue('PREMCPT', TPR.GetValue('E_GENERAL'));
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
    Compte := TPR.GetValue('E_GENERAL');
    if Method <> 'MUL' then
    begin
      if (Copy(TPR.GetValue('E_GENERAL'), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then TPR.PutValue('E_CONTREPARTIEGEN', PremCpt)
      else TPR.PutValue('E_CONTREPARTIEGEN', Prem421);
    end
    else
    begin
      TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
      if (Copy(TPR.GetValue('E_GENERAL'), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) then
      begin
        if TCP <> nil then TPR.PutValue('E_CONTREPARTIEGEN', TCP.GetValue('PREMCPT'));
      end
      else
      begin
        if TCP <> nil then TPR.PutValue('E_CONTREPARTIEGEN', TCP.GetValue('PREM421'));
      end;

    end;

    if (Copy(TPR.GetValue('E_GENERAL'), 1, 3) = Copy(VH_Paie.PGCptNetAPayer, 1, 3)) and (TPR.GetValue('E_AUXILIAIRE') <> '') then
    begin // Compte Auxiliaire renseigné alors on va rechercher le lettrage du tiers et son RIB
      CpteAuxi := TPR.GetValue('E_AUXILIAIRE');
      sRib := '';
      Q := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + CpteAuxi + '" AND R_SALAIRE="X"', True);
      if not Q.EOF then sRib := EncodeRIB(Q.FindField('R_ETABBQ').AsString, Q.FindField('R_GUICHET').AsString,
          Q.FindField('R_NUMEROCOMPTE').AsString, Q.FindField('R_CLERIB').AsString,
          Q.FindField('R_DOMICILIATION').AsString);
      Ferme(Q);
      TPR.PutValue('E_RIB', sRib);
      Q := OpenSQL('SELECT T_LETTRABLE,T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="' + CpteAuxi + '"', TRUE);
      Lettrable := FALSE;
      if not Q.EOF then Lettrable := (Q.FindField('T_LETTRABLE').AsString = 'X');
      if Lettrable then
      begin
        TPR.PutValue('E_ETATLETTRAGE', 'AL');
      end
      else
      begin
        TPR.PutValue('E_ETATLETTRAGE', '');
      end;
      if not Q.EOF then
      begin
        TRegl := TModeRegl.FindFirst(['MR_MODEREGLE'], [Q.FindField('T_MODEREGLE').AsString], FALSE);
        if TRegl <> nil then TPR.PutValue('E_MODEPAIE', TRegl.GetValue('MR_MP1'))
        else TPR.PutValue('E_MODEPAIE', 'CHQ');
      end
      else TPR.PutValue('E_MODEPAIE', 'CHQ');
      ferme(Q);
    end;

    if (TPR.GetValue('E_DEBIT') > 0) or (TPR.GetValue('E_CREDIT') < 0) then TPR.PutValue('E_ENCAISSEMENT', 'ENC')
    else TPR.PutValue('E_ENCAISSEMENT', 'DEC'); // Affectation code encaissement
    TGene := TOB_Generaux.FindFirst(['G_GENERAL'], [TPR.GetValue('E_GENERAL')], FALSE);
    if TGene <> nil then
    begin // si lettrable alors gestion des echéance et du mode de paiement
      if (TGene.GetValue('G_LETTRABLE') = 'X') then
      begin
        TPR.PutValue('E_NUMECHE', 1);
        TPR.PutValue('E_ECHE', 'X');
        if MDR_Def <> '' then TPR.PutValue('E_MODEPAIE', MDR_Def)
        else TPR.PutValue('E_MODEPAIE', 'CHQ'); // PT50 FQ 11318
        TPR.PutValue('E_DATEECHEANCE', TPR.GetValue('E_DATECOMPTABLE'));
      end
      else
      begin
        if (TPR.GetValue('E_NUMECHE') <> 1) then
        begin
          TPR.PutValue('E_NUMECHE', 0);
          TPR.PutValue('E_ECHE', '-');
        end;
      end;
      TPR.PutValue('E_DATEPAQUETMIN', TPR.GetValue('E_DATECOMPTABLE'));
      TPR.PutValue('E_DATEPAQUETMAX', TPR.GetValue('E_DATECOMPTABLE'));
      if (TGene.GetValue('G_VENTILABLE') = 'X') then TPR.PutValue('E_ANA', 'X')
      else TPR.PutValue('E_ANA', '-');

      if (TGene.GetValue('G_LETTRABLE') = 'X') and
        ((TGene.GetValue('G_NATUREGENE') = 'TIC') or (TGene.GetValue('G_NATUREGENE') = 'TID')) then
        TPR.PutValue('E_ETATLETTRAGE', 'AL');
      if ((TGene.GetValue('G_NATUREGENE') = 'COS') or (TGene.GetValue('G_NATUREGENE') = 'COC') or (TGene.GetValue('G_NATUREGENE') = 'COF') or (TGene.GetValue('G_NATUREGENE') =
        'COD'))
        and (TPR.Getvalue('E_AUXILIAIRE') = '')
        then
        TraceErr.Items.Add('Ecriture compte ' + TPR.GetValue('E_GENERAL') + 'Folio ' + IntToStr(TPR.GetValue('E_NUMEROPIECE')) + ' ligne ' + IntToStr(TPR.GetValue('E_NUMLIGNE')) +
          ' n''a pas de compte auxiliaire associé');
      LettrableAux := FALSE;
      TheModRegl := TGene.GetValue('G_MODEREGLE');
      if (TPR.GetValue('E_AUXILIAIRE') <> '') then
      begin
        Q := OpenSQL('SELECT T_LETTRABLE,T_MODEREGLE FROM TIERS WHERE T_AUXILIAIRE="' + CpteAuxi + '"', TRUE);
        if not Q.EOF then
        begin
          LettrableAux := (Q.FindField('T_LETTRABLE').AsString = 'X');
          TheModRegl := Q.FindField('T_MODEREGLE').AsString;
        end;
        Ferme(Q);
      end;
      NatGene := TGene.GetValue('G_NATUREGENE');
      GLet := TGene.GetValue('G_LETTRABLE');
      GPoint := TGene.GetValue('G_POINTABLE');
      if (((NatGene = 'TIC') or (NatGene = 'TID')) and (GLet = 'X')) or
        (((NatGene = 'BQE') or (NatGene = 'CAI')) and (GPoint = 'X')) or
        ((TPR.Getvalue('E_AUXILIAIRE') <> '') and LettrableAux)
        then
      begin
        TPR.PutValue('E_ETATLETTRAGE', 'AL'); // PT50 FQ 11318
        TPR.PutValue('E_ECHE', 'X');
        TPR.PutValue('E_NUMECHE', 1);
      end
      else
      begin
        TPR.PutValue('E_ETATLETTRAGE', 'RI'); // PT50
        TPR.PutValue('E_ECHE', '-');
        TPR.PutValue('E_NUMECHE', 0);
      end;
    end
    else
    begin // Cas où le compte général n'existe pas
      if (IntegODPaie = 'IMM') or (IntegODPaie = 'DIF') then
      begin
        TraceErr.Items.Add('Le compte ' + TPR.GetValue('E_GENERAL') + ' n''existe pas');
      end;
    end;
    if LettrableAux then
    begin
      TPR.PutValue('E_MODEPAIE', TheModRegl);
      TPR.PutValue('E_ETATLETTRAGE', 'AL');
    end;

    TPR.PutValue('E_NUMORDRE', TPR.GetValue('E_NUMLIGNE')); // Pour saisie par bordereau
    TOd := TOB.create('ECRITODPAIE', TEcritPaie, -1);
    if TOd <> nil then
    begin
      Tod.PutValue('PEC_ETABLISSEMENT', TPR.GetValue('E_ETABLISSEMENT'));
      Tod.PutValue('PEC_JOURNAL', TPR.GetValue('E_JOURNAL'));
      Tod.PutValue('PEC_GENERAL', TPR.GetValue('E_GENERAL'));
      Tod.PutValue('PEC_DATECOMPTABLE', TPR.GetValue('E_DATECOMPTABLE'));
      Tod.PutValue('PEC_NUMEROPIECE', TPR.GetValue('E_NUMEROPIECE'));
      Tod.PutValue('PEC_NUMLIGNE', TPR.GetValue('E_NUMLIGNE'));
      Tod.PutValue('PEC_AUXILIAIRE', TPR.GetValue('E_AUXILIAIRE'));
      Tod.PutValue('PEC_LIBELLE', TPR.GetValue('E_LIBELLE'));
      Tod.PutValue('PEC_REFINTERNE', TPR.GetValue('E_REFINTERNE'));
      if PaieEuro = ComptaEuro then
      begin // on remplit toujours les montants dans la monnaie de tenue de la paie
        Tod.PutValue('PEC_DEBIT', ARRONDI(TPR.GetValue('E_DEBIT'), NbDec));
        Tod.PutValue('PEC_CREDIT', ARRONDI(TPR.GetValue('E_CREDIT'), NbDec));
        // test equilibre débit/crédit sur chaque FOLIO
        if Tod.GetValue('PEC_NUMEROPIECE') = 1 then
        begin
          TotDebit := TotDebit + ARRONDI(TPR.GetValue('E_DEBIT'), NbDec);
          TotCredit := TotCredit + ARRONDI(TPR.GetValue('E_CREDIT'), NbDec);
        end
        else
        begin
          TotDebit2 := TotDebit2 + ARRONDI(TPR.GetValue('E_DEBIT'), NbDec);
          TotCredit2 := TotCredit2 + ARRONDI(TPR.GetValue('E_CREDIT'), NbDec);
        end;
      end;

      if VH_PAIE.PGTenueEuro then Tod.PutValue('PEC_MONNAIE', 'EUR')
      else Tod.PutValue('PEC_MONNAIE', 'FRF');
    end;
  end;

  TotDebit := Arrondi(TotDebit, NbDec);
  TotCredit := Arrondi(TotCredit, NbDec);
  TotDebit2 := Arrondi(TotDebit2, NbDec);
  TotCredit2 := Arrondi(TotCredit2, NbDec);

  St := 'Débit = ' + StrfMontant(TotDebit + TotDebit2, 15, NbDec, '', TRUE) + '   Crédit = ' + StrfMontant(TotCredit + TotCredit2, 15, NbDec, '', TRUE);
  TZZ := TOB.create('ECRITODPAIE', TEcritPaie, -1);
  TZZ.PutValue('PEC_ETABLISSEMENT', '');
  TZZ.PutValue('PEC_JOURNAL', '');
  TZZ.PutValue('PEC_GENERAL', '');
  TZZ.PutValue('PEC_DATECOMPTABLE', NULL);
  TZZ.PutValue('PEC_NUMEROPIECE', 99999999);
  TZZ.PutValue('PEC_NUMLIGNE', 99999999);
  TZZ.PutValue('PEC_AUXILIAIRE', '');
  TZZ.PutValue('PEC_LIBELLE', 'Totaux : ');
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
    i := ExecuteSQL(st1);
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
    FreeAndNIL(MaTob);
    CommitTrans;
    st := ' AND PEC_ALEAT = ' + FloatToStr(Aleat);
    LanceEtat('E', 'PAN', 'POD', True, FALSE, False, nil, St, '', False);
// PT1  : 06/05/2004 : V_500 PH Mise ne place liste d'exportation
    if GetControlText('LISTEEXPORT') = 'X' then LanceEtatTOB('E', 'PAN', 'POD', nil, True, True, False, Pan, St, '', False);
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
        if Etabl = '' then st := VH_Paie.PGCHEMINEAGL + '\ODCPPAIEPGI.TXT'
        else st := VH_Paie.PGCHEMINEAGL + '\ODCPPAIEPGI' + Etabl + '.TXT';
        TEcritPaie.SaveToFile(St, FALSE, TRUE, FALSE, '');
      end;
    end;
    OkOk := mrNo;
  end;

  if Assigned(ErrorGene) and (IntegODPaie <> 'ECP') then //((Ecart <> 0) or (Ecart1 <> 0)) then
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
        // Affectation definitive du numéro de piece - boucle sur les ecritures pour affecter les numeros de pieces et de lignes
        for I := 0 to Tob_LesEcrit.Detail.count - 1 do
        begin // B4
          TPR := Tob_LesEcrit.Detail[I];
          OldP := TPR.GetValue('E_NUMEROPIECE');
          OldL := TPR.GetValue('E_NUMLIGNE');
          if (TPR.GetValue('E_NUMEROPIECE') = 1) then
          begin // B5
            NewP := NumFolioOD1;
            NewL := NumEcritOD1;
            TPR.PutValue('E_NUMEROPIECE', NumFolioOD1);
            TPR.PutValue('E_NUMLIGNE', NumEcritOD1);
            NumEcritOD1 := NumEcritOD1 + 1;
            if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
          end // F5
          else
          begin // B6
            if (TPR.GetValue('E_NUMEROPIECE') = 2) then
            begin // B7
              NewP := NumFolioOD1;
              NewL := NumEcritOD1;
              TPR.PutValue('E_NUMEROPIECE', NumFolioOD2);
              TPR.PutValue('E_NUMLIGNE', NumEcritOD2);
              NumEcritOD2 := NumEcritOD2 + 1;
              if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
            end; // F7
          end; // F6
        end; // Fin du for  // F4
      end // F3 Fin du cas mono ou etab principal
      else // Multi etab donc recherche et affectation des numéros de pieces pour chaque etablissement
      begin // B8
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
          TCP := TContreP.FindNext([''], [''], FALSE);
        end; // F9

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
            if TPR.GetValue('E_DEBIT') < 0 then EcrPositif := FALSE;
            if TPR.GetValue('E_CREDIT') < 0 then EcrPositif := FALSE;
          end;
          TCP := TContreP.FindFirst(['ETABLISSEMENT'], [TPR.GetValue('E_ETABLISSEMENT')], FALSE);
          if TCP <> nil then
          begin
            OldP := TPR.GetValue('E_NUMEROPIECE');
            OldL := TPR.GetValue('E_NUMLIGNE');
            NewP := TCP.GetValue('FOLIO1');
            NewL := TCP.GetValue('NUMECR1');
          end;
          if (TPR.GetValue('E_NUMEROPIECE') = 1) then
          begin // B13
            TPR.PutValue('E_NUMEROPIECE', TCP.GetValue('FOLIO1'));
            TPR.PutValue('E_NUMLIGNE', TCP.GetValue('NUMECR1'));
            TCP.PutValue('NUMECR1', TCP.GetValue('NUMECR1') + 1);
            if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
          end // F13
          else
          begin // B14
            if (TPR.GetValue('E_NUMEROPIECE') = 2) then
            begin // B15
              TPR.PutValue('E_NUMEROPIECE', TCP.GetValue('FOLIO2'));
              TPR.PutValue('E_NUMLIGNE', TCP.GetValue('NUMECR2'));
              TCP.PutValue('NUMECR2', TCP.GetValue('NUMECR2') + 1);
              if VH_Paie.PGAnalytique then EcrPaieVersAna(TOB_Generaux, TPR, TOBANA, OldP, OldL, NewP, NewL, ConservAna);
            end; // F15
          end; // F14 Fin du Else
        end; // F12 fin du for boucle sur les ecritures sur affectation numeros de pieces et de lignes
      end; // F8 fin du else cas mono ou principal = cas multi etab
    end; // F2 fin si on retraite les ecritures = cas integration dans la paie pgi
    // Equilibre écriture en fonction des écarts de conversion et du paramètrage de la compta
    EquilibreEcrPaie(Tob_LesEcrit, MM, NbDec);
  end;

  EcrPositif := TRUE;
  for I := 0 to Tob_LesEcrit.Detail.count - 1 do
  begin
    if EcrPositif then
      if (TPR.GetValue('E_DEBIT') < 0) or (TPR.GetValue('E_CREDIT') < 0) then EcrPositif := FALSE;
  end;

  if (not EcrPositif) and (not VH^.MontantNegatif) then
  begin
    PgiError('Vous avez une écriture négative alors que vous n''autorisez pas les montants négatifs #13#10 dans la comptabilité PGI ', Ecran.caption);
    OkOk := MrNo;
  end;
  if (OkOk = mrYes) then
  begin
    if IntegODPaie = 'ECP' then st := 'Voulez-vous générer un fichier des ODs au format Paie PGI'
    else if IntegODPaie = 'IMM' then st := 'Voulez-vous enregistrer immédiatement vos écritures dans la comptabilité'
    else if IntegODPaie = 'DIF' then st := 'Voulez-vous enregistrer en différé vos écritures'
    else if IntegODPaie = 'EFC' then st := 'Voulez-vous générer un fichier des ODs au format Export Comptabilté PGI';
  end; // F1 Fin si integration autorisee
// PT1  : 06/05/2004 : V_500 PH Mise ne place liste d'exportation
  if VH_Paie.PGAnalytique and (OkOk = mrYes) then AnalyseAnalytique(IntegODPaie, TOBAna, PaieEuro, ComptaEuro, NbDec, ConservAna, Pan);

  if OkOk = mrYes then OkOk := PGIAsk(St, Ecran.Caption);
  if OkOk = mrYes then
  begin
    if IntegODPaie = 'ECP' then // Format paie Pgi
    begin
      if VH_Paie.PGCHEMINEAGL = '' then PgiError('Vous n''avez pas renseigné le chemin de stockage', 'Paramètres société')
      else
      begin
        if Etabl = '' then st := VH_Paie.PGCHEMINEAGL + '\ODCPPAIEPGI.TXT'
        else st := VH_Paie.PGCHEMINEAGL + '\ODCPPAIEPGI' + Etabl + '.TXT';
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
        if not Tob_LesEcrit.InsertDBByNivel(False) then V_PGI.IoError := oeUnknown
        else
        begin
          if not VH_Paie.PGTypeEcriture then MajSoldesEcritureTOB(Tob_LesEcrit, True); // MAJ des soldes comptables
        end;
      end;
      if IntegODPaie = 'DIF' then // Integration différé dans la paie PGI
      begin
        if not PGInsertionDifferee(Tob_LesEcrit) then V_PGI.IoError := oeUnknown;
      end;
    end; // fin du else
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
      i := ExecuteSQL(st);
      TEcritPaie.InsertDB(nil, FALSE);
      CommitTrans;
    except
      Rollback;
      PGIError('Une erreur est survenue lors de la Génération comptable', Ecran.caption);
    end;

  end;
  TEcritPaie.Free;
  TEcritPaie := nil;
  if TContreP <> nil then
  begin
    TContreP.Free;
    TContreP := nil;
  end;
  Tob_GuideEcr.Free;
  Tob_GuideEcr := nil;
  FreeAndNIL(T_Etab);

  GetLocalTime(SystemTime0);
  if TraceErr.Items.Count > 1 then Trace.Items.Add('Traitement interrompu à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)))
  else Trace.Items.Add('Fin du traitement à :' + TimeToStr(SystemTimeToDateTime(SystemTime0)));
  PGIInfo('Le traitement est terminé', Ecran.caption);

  TheTrace := TStringList.Create;
  if (TraceErr.Items.Count > 1) or (OkOk <> mrYes) then
  begin
    TheTrace.Add('Génération interrompue car :');
    if Ecart <> 0 then TheTrace.Add('Attention, le folio 1 n''est pas équilibré.');
    if Ecart1 <> 0 then TheTrace.Add('Attention, le folio 2 n''est pas équilibré');
    St := 'Débit = ' + StrfMontant(TotDebit + TotDebit2, 15, NbDec, '', TRUE) + '   Crédit = ' + StrfMontant(TotCredit + TotCredit2, 15, NbDec, '', TRUE);
    TheTrace.Add(St);
    TheTrace.Add('Vous avez des erreurs ');
    TheTrace.Add('Abandon du traitement');
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
  if TOB_Generaux <> nil then
  begin
    TOB_Generaux.Free;
    TOB_Generaux := nil;
  end;
  if TOB_LesEcrit <> nil then
  begin
    TOB_LesEcrit.Free;
    TOB_LesEcrit := nil;
  end;
  FreeAndNil(TModeRegl);
end;

procedure TOF_PG_DotProvCpCompta.OnArgument(Arguments: string);
var
  F: TFVierge;
  st, Etab1: string;
  LaDate: TDateTime;
begin
  inherited;
  st := Trim(Arguments);
  LblDD := THLabel(GetControl('LBLDU'));
  if LblDD <> nil then LblDD.Caption := ReadTokenSt(st);
  LaDate := StrToDate(LblDD.Caption);
  SetControlText('DATEFIN', LblDD.Caption);
  Etab := ReadTokenSt(st);
  Etab1 := ReadTokenSt(st);
  Method := ReadTokenSt(st);
  if Method = 'PRI' then Etabl := VH^.EtablisDefaut; // recuperation etablissement par defaut
  BtnLance := TToolbarButton97(GetControl('BTNLANCE'));
  if Method = '' then // ¨Par defaut cas mono etablissement
    if BtnLance <> nil then SetControlEnabled('BTNLANCE', FALSE);
  SetControlText('LBLMETHODE', RechDom('PGVENTILMULTIETAB', Method, FALSE));
  if BtnLance <> nil then BtnLance.OnClick := LanceDotProvCpCompta;
  JournalPaie := THValComboBox(GetControl('CBXJNAL'));
  if JournalPaie <> nil then JournalPaie.Value := VH_Paie.PGJournalPaie;
  if VH_Paie.PGIntegODPaie = 'ECP' then
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
  F := TFVierge(Ecran);
  if F <> nil then
  begin
    st := ' SELECT * FROM PROVCP WHERE PDC_DATEARRET="' + UsDateTime(LaDate) + '"';
    if Etab <> '' then
    begin
      if Method = 'PRI' then St := st + ' AND PDC_ETABLISSEMENT="' + Etabl + '"'
      else // PT2
      // PT3 Borne inférieure prise en compte
        if Method = 'MON' then St := st + ' AND PDC_ETABLISSEMENT>="' + Etab + '" AND PDC_ETABLISSEMENT <="' +
          Etab1 + '"';
    end;
    QMul := OpenSql(st, TRUE);
  end;
  if QMUL = nil then
  begin
    PgiError('Vous n''avez pas de provisions calculées à la date d''arrêté, traitement impossible', Ecran.caption);
    SetControlEnabled('BTNLANCE', FALSE);
  end;
end;
// Initialisation par défaut des champs d'une ligne écriture

procedure TOF_PG_DotProvCpCompta.LignePaieECR(MM: RMVT; TOBE: TOB);
var
  i: integer;
  RefI: string;
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

procedure TOF_PG_DotProvCpCompta.InitEcritDefaut(TOBE: TOB);
var
  i: integer;
begin
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
  TOBE.PutValue('E_PERIODE', GetPeriode(V_PGI.DateEntree));
  TOBE.PutValue('E_SEMAINE', NumSemaine(V_PGI.DateEntree));
  TOBE.PutValue('E_ETATREVISION', '');
  TOBE.PutValue('E_IO', '');
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
end;

procedure TOF_PG_DotProvCpCompta.RenseigneClefComptaPaie(Etab: string; var MM: RMVT);
begin
  MM.Etabl := Etab;
  if VH_Paie.PGIntegODPaie = 'ECP' then
  begin
    if (JournalPaie.Value = '') and (JournalPaie.text <> '') then MM.Jal := JournalPaie.Text
    else MM.Jal := JournalPaie.Value;
  end
  else MM.Jal := JournalPaie.Value;
  MM.DateC := StrToDate(LBLDD.Caption);
  MM.Exo := QuelExoDT(MM.DateC);
  MM.CodeD := V_PGI.DevisePivot;
  MM.DateTaux := MM.DateC;
  MM.ModeSaisieJal := ModeSaisieJal;
  MM.nature := 'OD';
  MM.TauxD := 1; // devise pivot;
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

end;

// Fonction qui rempli la TOB paie des ecritures en fonction  de la ligne bulletin

procedure TOF_PG_DotProvCpCompta.RemplirT_Ecr(TypVentil: string; Montant: Double; TypMvt, Libelle: string; T_Ecr, TOB_Generaux: TOB; Compte, CpteAuxi: string;
  LeFolio, LeNumEcrit: Integer);
var
  NumLig, i: Integer;
  LeMontant, LeMontantF, LeMontantE: Double;
  MIdentique: Boolean;
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
  MIdentique := FALSE;
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
  if PaieEuro = ComptaEuro then MIdentique := TRUE;
  if TypMvt = 'DEB' then
  begin // Alimentation Debit
    if ComptaEuro then // Compta en EURO
    begin
      DebMD := LeMontantE;
      DebMP := LeMontantE;
      DebME := LeMontantF;
      LeMontant := T_Ecr.GetValue('E_DEBIT');
      T_Ecr.PutValue('E_DEBIT', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValue('E_DEBITDEV');
      T_Ecr.PutValue('E_DEBITDEV', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
    end
    else
    begin
      DebMD := LeMontantF;
      DebMP := LeMontantF;
      DebME := LeMontantE;
      LeMontant := T_Ecr.GetValue('E_DEBIT');
      T_Ecr.PutValue('E_DEBIT', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValue('E_DEBITDEV');
      T_Ecr.PutValue('E_DEBITDEV', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
    end;
  end
  else // Alimentation Credit
  begin
    if ComptaEuro then // Compta en EURO
    begin
      CreMD := LeMontantE;
      CreMP := LeMontantE;
      CreME := LeMontantF;
      LeMontant := T_Ecr.GetValue('E_CREDIT');
      T_Ecr.PutValue('E_CREDIT', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValue('E_CREDITDEV');
      T_Ecr.PutValue('E_CREDITDEV', ARRONDI(LeMontant + LeMontantE, V_PGI.OkDecV));
    end
    else
    begin
      CreMD := LeMontantF;
      CreMP := LeMontantF;
      CreME := LeMontantE;
      LeMontant := T_Ecr.GetValue('E_CREDIT');
      T_Ecr.PutValue('E_CREDIT', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
      LeMontant := T_Ecr.GetValue('E_CREDITDEV');
      T_Ecr.PutValue('E_CREDITDEV', ARRONDI(LeMontant + LeMontantF, V_PGI.OkDecV));
    end;
  end;
  T_Ecr.PutValue('E_GENERAL', Compte);
  T_Ecr.PutValue('E_AUXILIAIRE', CpteAuxi);
  T_Ecr.PutValue('E_NUMEROPIECE', LeFolio);
  T_Ecr.PutValue('E_NUMLIGNE', LeNumEcrit);
  T_Ecr.PutValue('E_LIBELLE', Libelle);
  T_Ecr.PutValue('E_REFINTERNE', RefPaie);
  T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [Compte], FALSE);
  if not VH_Paie.PGAnalytique then exit; // pas de gestion analytique donc fin procedure
  if ((Tob_Ventana = nil) or (TOB_Ventana.Detail.Count <= 0)) and
    ((Tob_VenSal = nil) or (TOB_VenSal.Detail.Count <= 0))
    then exit; // pas de ventilation pour le salarie alors que analytique géré
  if (T_Lecompte <> nil) and (T_LeCompte.GetValue('G_VENTILABLE') = 'X') then
  begin
    for i := 1 to 5 do
    begin
      St := 'G_VENTILABLE' + IntToStr(i);
      AVentil[i] := (T_LeCompte.GetValue(St) = 'X'); // Tableau des Axes Ventilables
    end;
    CumulAnalPaie(TypVentil, Compte, T_Ecr.GetValue('E_ETABLISSEMENT'), CpteAuxi, IntToStr(LeFolio), IntToStr(LeNumEcrit),
      DebMD, DebMP, DebME, CreMD, CreMP, CreME, AVentil);
  end;
end;
// Fonction qui decline toutes les lignes analytiques en fonction des ventilations trouvees

procedure TOF_PG_DotProvCpCompta.EcrPaieVersAna(TOB_Generaux, TOBEcr, TOBAna: TOB; OldP, OldL, NewP, NewL: Integer; Conservation: Boolean = FALSE);
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
  FNumAna := PrefixeToNum('Y');
  FNumEcr := PrefixeToNum('E');
  if not Conservation then
  begin
    T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [TOBEcr.GetValue('E_GENERAL')], FALSE);
    if (T_Lecompte = nil) or (T_LeCompte.GetValue('G_VENTILABLE') = '-') then exit;
  end;
  for k := 1 to 5 do
  begin
    MontantAxe[k] := 0;
    MontantAxeEuro[k] := 0;
  end;
  TotalEcriture := TOBEcr.GetValue('E_DEBIT') + TOBEcr.GetValue('E_CREDIT');
  TotalDevise := TOBEcr.GetValue('E_DEBITDEV') + TOBEcr.GetValue('E_CREDITDEV');
  iTrouv := -1;
  for k := 0 to TTA.Count - 1 do // boucle identification de l'écriture
  begin
    LaLigne := T_AnalPaie(TTA[k]);
    if (LaLigne.Cpte = TOBEcr.GetValue('E_GENERAL')) and (LaLigne.Etab = TOBEcr.GetValue('E_ETABLISSEMENT')) and (LaLigne.Auxi =
      TOBEcr.GetValue('E_AUXILIAIRE')) and
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
    Pourc := Arrondi((XDT.DebMP + XDT.CreMP) / (TotalEcriture) * 100, 6);
    MontantAxe[StrToInt(Copy(XDT.Ax, 2, 1))] := MontantAxe[StrToInt(Copy(XDT.Ax, 2, 1))] + XDT.DebMP + XDT.CreMP;
    MontantAxeEuro[StrToInt(Copy(XDT.Ax, 2, 1))] := MontantAxeEuro[StrToInt(Copy(XDT.Ax, 2, 1))] + XDT.DebME + XDT.CreME;
    TAnal.PutValue('Y_POURCENTAGE', Pourc);
    TAnal.PutValue('Y_TOTALECRITURE', TotalEcriture);
    TAnal.PutValue('Y_TOTALDEVISE', TotalDevise);
  end;
  AnaAttente := FALSE;
  for k := 1 to 5 do
  begin
    if (Arrondi(MontantAxe[k] - TotalEcriture, 2) <> 0) then AnaAttente := TRUE;
  end;
  // Il faut verifier si l'axe est ventilable sinon on passe le tout sur la section d'attente de l'axe

  if AnaAttente then PaieGenereAttente(MontantAxe, MontantAxeEuro, TOB_Generaux, TOBEcr, TOBAna, TotalEcriture, TotalDevise, TotalEuro, Conservation);
  // Intégration des lignes analytiques dans la TOB des lignes afin de faire l'insertion dans la base
  // PGVisuUnObjet (TOBANA, 'Les ecritures analytiques','');
  NumVentil := 1;
  TAnal := TOBAna.FindFirst([''], [''], FALSE);
  while TAnal <> nil do
  begin
    TAnal.PutValue('Y_NUMVENTIL', NumVentil);
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

procedure TOF_PG_DotProvCpCompta.PaieGenereAttente(MontantAxe, MontantAxeEuro: array of double; TOB_Generaux, TOBEcr, TOBAna: TOB; TotalEcriture, TotalDevise,
  TotalEuro: Double; Conservation: Boolean = FALSE);
var
  k, j, i: Integer;
  NomEcr, NomAna: string;
  TAnal, T_LeCompte: TOB;
  Axe, Sens, st: string;
  Pourc, Montant, MontantE: Double;
  AVentil: array[1..5] of boolean;
  FNumAna, FNumEcr: integer;
begin
  FNumAna := PrefixeToNum('Y');
  FNumEcr := PrefixeToNum('E');

  T_LeCompte := TOB_Generaux.FindFirst(['G_GENERAL'], [TOBEcr.GetValue('E_GENERAL')], FALSE);
  for i := 1 to 5 do AVentil[i] := FALSE;
  if ((T_Lecompte <> nil) and (T_LeCompte.GetValue('G_VENTILABLE') = 'X')) then
  begin
    for i := 1 to 5 do
    begin
      St := 'G_VENTILABLE' + IntToStr(i);
      AVentil[i] := (T_LeCompte.GetValue(St) = 'X'); // Tableau des Axes Ventilables
    end;
  end else exit; // compte non ventilable
  for k := 0 to 4 do
  begin
    if not AVentil[K + 1] then continue;
    if (Arrondi(MontantAxe[k] - TotalEcriture, 2) <> 0) then
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
          if V_PGI.DECHAMPS[FNumEcr, i].Tipe = V_PGI.DECHAMPS[FNumAna, j].Tipe
            then TAnal.PutValeur(j, TOBEcr.GetValeur(i));
        end;
      end;
      if TOBEcr.GetValue('E_DEBIT') <> 0 then Sens := 'D'
      else Sens := 'C';
      Axe := 'A' + IntToStr(k + 1);
      TAnal.PutValue('Y_AXE', Axe);
      TAnal.PutValue('Y_SECTION', VH^.Cpta[AxeToFb(Axe)].Attente);
      Montant := Arrondi(TotalEcriture - MontantAxe[k], 2);
      MontantE := Arrondi(TotalEuro - MontantAxeEuro[k], 2);
      Pourc := Arrondi((Montant) / (TotalEcriture) * 100, 6);
      if Sens = 'D' then
      begin
        TAnal.PutValue('Y_DEBIT', Montant);
        TAnal.PutValue('Y_CREDIT', 0);
        TAnal.PutValue('Y_DEBITDEV', Montant);
        TAnal.PutValue('Y_CREDITDEV', 0);
      end
      else
      begin
        TAnal.PutValue('Y_DEBIT', 0);
        TAnal.PutValue('Y_CREDIT', Montant);
        TAnal.PutValue('Y_DEBITDEV', 0);
        TAnal.PutValue('Y_CREDITDEV', Montant);
      end;
      TAnal.PutValue('Y_TOTALECRITURE', TotalEcriture);
      TAnal.PutValue('Y_TOTALDEVISE', TotalDevise);
    end;
  end;
end;

procedure TOF_PG_DotProvCpCompta.EquilibreEcrPaie(TOBEcr: TOB; MM: RMVT; NbDec: integer);
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
  for i := 0 to TOBEcr.Detail.Count - 1 do
  begin
    TOBL := TOBEcr.Detail[i];
    DD := DD + TOBL.GetValue('E_DEBITDEV');
    CD := CD + TOBL.GetValue('E_CREDITDEV');
    DP := DP + TOBL.GetValue('E_DEBIT');
    CP := CP + TOBL.GetValue('E_CREDIT');
  end;
  EcartD := Arrondi(DD - CD, NbDec);
  EcartP := Arrondi(DP - CP, V_PGI.OkDecV);
  EcartE := Arrondi(DE - CE, V_PGI.OkDecE);
  if ((EcartD = 0) and (EcartP = 0) and (EcartE = 0)) then Exit;
  if ((EcartD <> 0) or (EcartE <> 0) or (EcartP <> 0)) then // Ecart Debit/credit donc gestion ligne ecart
  begin
    TOBL.PutValue('E_ENCAISSEMENT', SensEnc(DP, CP)); // determination du code encaissement
    CreerLigneEcartPaie(TOBEcr, TOBL, MM, EcartD, EcartP, EcartE);
  end;
end;
// Fonction qui créer la ligne écart car on aura tjrs une erreur d'équilibre à cause des arrondis dans la monnaie inversée

function TOF_PG_DotProvCpCompta.CreerLigneEcartPaie(TOBEcr, TOBL: TOB; MM: RMVT; EcartD, EcartP, EcartE: Double): T_ErrCpta;
var
  CptECC: string;
  TOBE: TOB;
  NumL: integer;
  Debit: boolean;
begin
  Result := rcOk;
  TOBE.PutValue('E_GENERAL', CptECC);
  {Divers}
  TOBE.PutValue('E_TYPEMVT', 'ECC');
  TOBE.PutValue('E_LIBELLE', 'Ecart de conversion');
  NumL := TOBE.GetValue('E_NUMLIGNE') + 1;
  TOBE.PutValue('E_NUMLIGNE', NumL);
  {Montants}
  if Debit then
  begin
    TOBE.PutValue('E_DEBIT', -EcartP);
    TOBE.PutValue('E_DEBITDEV', -EcartD);
    TOBE.PutValue('E_CREDIT', 0);
    TOBE.PutValue('E_CREDITDEV', 0);
  end else
  begin
    TOBE.PutValue('E_CREDIT', EcartP);
    TOBE.PutValue('E_CREDITDEV', EcartD);
    TOBE.PutValue('E_DEBIT', 0);
    TOBE.PutValue('E_DEBITDEV', 0);
  end;

end;
// Fonction qui cumule les debits/credits pour chaque ecriture et par section/axe

procedure TOF_PG_DotProvCpCompta.CumulAnalPaie(TypVentil, Cpte, Etab, Auxi, NumP, NumL: string; DebMD, DebMP, DebME, CreMD, CreMP, CreME: Double; AVentil: array of
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
  St := TypVentil; // Recup de la nature et de la rubrique traitée
  if TypVentil <> 'S' then
    Nature := ReadTokenSt(St);
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

  for i := 0 to LaTob.Detail.Count - 1 do
  begin
    TVa := LaTob.Detail[i];
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
    if iTrouv < 0 then XDT := T_DetAnalPaie.Create;
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

function TOF_PG_DotProvCpCompta.RendLongueurRacine(LeCompte: string): Integer;
begin
  if (Copy(LeCompte, 1, VH_Paie.PGLongRacine421) = VH_Paie.PGCptNetAPayer) then
  begin
    result := VH_Paie.PGLongRacine421;
    exit;
  end;
  // Classe 1 ou 7,9
  if ((Copy(LeCompte, 1, 1) = '6') or (Copy(LeCompte, 1, 1) = '7') or (Copy(LeCompte, 1, 1) = '9')) then
    result := VH_Paie.PGLongRacine
  else
  begin
    //   4 ou 2,1
    if ((Copy(LeCompte, 1, 1) = '4') or (Copy(LeCompte, 1, 1) = '2') or (Copy(LeCompte, 1, 1) = '1')) then
      result := VH_Paie.PGLongRacin4
    else result := Strlen(PChar(LeCompte));
  end;
end;
// PT1  : 06/05/2004 : V_500 PH Mise ne place liste d'exportation

procedure TOF_PG_DotProvCpCompta.AnalyseAnalytique(MaMethode: string; TOBAna: TOB; PaieEuro, ComptaEuro: Boolean; NbDec: Integer; Conservation: Boolean; Pan: TPageControl);
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
// PT1  : 06/05/2004 : V_500 PH Mise ne place liste d'exportation
      if GetControlText('LISTEEXPORT') = 'X' then LanceEtatTOB('E', 'PAN', 'POA', nil, True, TRUE, False, Pan, St, '', False);
    end;
    FreeAndNIL(MaTob);
    ExecuteSql('DELETE FROM ECODANAPAIE WHERE PEA_ALEAT = ' + FloatToStr(Aleat))
  except
    Rollback;
    PGIError('Une erreur est survenue lors de l''édition du journal analytique', 'Génération comptable');
  end;

end;

procedure TOF_PG_DotProvCpCompta.OnClose;
begin
  inherited;
  if QMUL <> nil then FERME(QMUL);
end;


function TOF_PG_DotProvCpCompta.ChargeVentileLigAna(Salarie: string; DateDebut, DateFin: TDateTime): TOB;
var
  RefA, St: string;
  Q: TQuery;
  TOBAnal: TOB;
begin
  TOBAnal := TOB.Create('Analytique bulletin Salarié', nil, -1);
  RefA := Salarie + ';' // Code Salarie
    + FormatDateTime('ddmmyyyy', DateDebut) + ';' // Date Debut bulletin
    + FormatDateTime('ddmmyyyy', DateFin) + ';'; // Date Fin bulletin
  St := 'SELECT * FROM VENTANA WHERE YVA_TABLEANA = "PPU" AND YVA_IDENTIFIANT like "' + RefA + 'COT" ORDER BY YVA_AXE,YVA_NATUREID,YVA_IDENTLIGNE';
  Q := OpenSql(st, TRUE);
  TOBAnal.LoadDetailDB('VENTANA', '', '', Q, FALSE);
  ferme(Q);
  result := TOBAnal;
end;

initialization
  registerclasses([TOF_PG_DotProvCpCompta]);
end.


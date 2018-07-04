{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 16/12/2004
Modifié le ... : 16/12/2004
Description .. :  CA - 05/07/1999 - On enchaîne l'annulation dans le cas
Suite ........ : d'un élément exceptionnel suivi d'un changement de plan.
Suite ........ :  CA - 01/06/2004 - Suppression du choix de la liste des
Suite ........ : immobilisations à annuler : on fait systématiquement pour
Suite ........ : toutes les immobilisations concernées   ( sauf en S1 )
Suite ........ : FQ 15081 - 16/12/2004 - CA - Ajout de VarIsNull pour
Suite ........ : détecter le cas des champs à NULL sous Oracle. Cela
Suite ........ : entraînait un plantage et l'annulation d el'opération était
Suite ........ : impossible.
Suite......... : FQ 16780 - 27/09/2005 - mbo - qd on annule une cession :  ne pas réinitialiser le code
                 i_opechangeplan si un exceptionnel a été auparavant
Suite......... : FQ 16782 - 28/09/2005 - mbo - si exceptionnel = reprise : le montant apparait
                 non signé dans l'affichage opération en cours
                 // TGA 10/2005 Gestion de la dépréciation pour CRC200210
Suite......... : PGR 11/2005 Changement des conditions d'amortissement
Mots clefs ... : BTY 11/05 Révision plan d'amortissement avec Calcul plan futur avec VNC Et Gestion plan fiscal CRC 2002-10
                 TGA 16/11/2005 affichage de la reprise de dépréciation en valeur absolue
Suite......... : PGR 12/2005 Changement des conditions d'amortissement(récup I_REPRISEECO et I_REPRISEFISC)
Suite......... : FQ 16897 - mbo - 03/01/2006 Tva doublée suite annulation de cession + Montant HT faux
Suite..........: FQ 17259 BTY 01/06 Nelle colonne dans IMMO pour la dépréciation
Suite..........: BTY - 01/06 En annulation dépréciation, OPECHANGEPLAN peut aussi concerner DPR
Suite..........: BTY - 03/06 FQ 17446 En annulation modif bases, restituer I_JOURNALA
Suite..........: BTY - 04/06 FQ 17516 Gérer REG opération de changement de regroupement
                 TGA 04/04/2006 maj i_reprisedep
Suite..........: BTY - 04/06 FQ 17629 RAZ Règle de calcul de la PMValue en annulation de cession
Suite..........: MBO - 25/04/2006 - FQ 17923 : nouvelle opération : modif date de mise en service
Suite..........: MBO - 17/05/2006 - fq 17569 : impact des dates deb amort sur le changt de méthode
Suite..........: BTY - 05/06 FQ 18119 Positionner les indicateurs de modif de compta après annulation de l'opération
Suite..........: MVG - 12/07/2006 pour correction conseil de compil
Suite..........: MBO - 11/09/2006 - gestion de la prime d'équipement PRI
Suite..........: MBO - 11/09/2006 - correction sur MajOpeEnCoursImmo ou parfois mms n'était pas géré et dpr en double
Suite..........: MBO - 09/10/2006 - gestion de la subvention d'équipement
Suite..........: BTY - 10/06 - Nelle opération RPR : Réduction de la prime d'équipement
Suite..........: BTY - 10/06 - FQ 18963 En annulation d'opération à partir de la fiche immo, on reste dans la fiche
Suite..........: BTY - 10/06 - FQ 18976 En retour d'annulation de cession à partir de la fiche immo,
Suite..........: pour une immo SANS plan fiscal, on mettait une base fiscale => message de modif bloquant en CWAS
Suite..........: et pb potentiel de toute façon
Suite..........: MBO - 24/10/2006 - en annulation de cession : on restitue les antérieurs saisis subvention stockés ds immolog
Suite..........: BTY - 11/06 - Nelle opération RSB : Réduction de la subvention d'investissement
Suite..........: BTY - 11/06 - FQ 19091 Si plusieurs fois la même opération sur une immo, on affichait toujours le 1er
Suite..........: BTY - 11/06 - Annulation de révision de plan CDM, CD2, CD3 : restaurer la durée d'inaliénabilité
Suite..........: BTY - 12/06 - FQ 19280 Annulation de levée option - Restituer l'ancienne date de fin de contrat
Suite..........: MBO - 03/07 - FQ 17512 Annulation de cession : récup antérieurs dérog et réintég fiscale
Suite..........: MBO - 20/03/2007 correction d'un messagebox qui ne passe plus en delphi 7
Suite..........: BTY - 04/07 Nelle gestion fiscale en révision de plan d'amortissement
Suite..........: MBO - 17/04/2007 - Annulation d'un remplacement de composant
Suite..........: BTY - 04/07 Nelle gestion fiscale en modification des bases
Suite..........: MBO - 08/06/2007 - FQ 20557 : Annul cession : init des zones i_repcedeco, i_repcedfisc et i_qtcede
Suite..........: BTY - 18/09/07 - fq 21754 Nouveaux paramètres immos : Sur cession, afficher si le jour de sortie a été amorti
*****************************************************************}
unit OpEnCour;

interface


uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  ComCtrls,
  ExtCtrls,
  HPanel, hmsgbox, HSysMenu, HTB97, HRichEdt, HRichOLE, Grids,
  Hctrls, StdCtrls,
  {$IFNDEF EAGLCLIENT}
    {$IFDEF ODBCDAC}
    odbcconnection, odbctable, odbcquery, odbcdac,
    {$ELSE}
      {$IFNDEF DBXPRESS}
      dbtables,
      {$ELSE}
      uDbxDataSet,
      {$ENDIF}
    {$ENDIF}
  {$ENDIF EAGLCLIENT}
  {$IFDEf VER150}
  Variants,
  {$ENDIF}
  majtable,
  HEnt1,
  Hqry,
  Math,
  UiUtil,
  HStatus,
  ImOuPlan,
  ImEnt,
  utob
  {$IFDEF EAGLCLIENT}
  ,MaineAGL
  {$ELSE}
  ,HDB
  ,Db
  ,FE_Main
  ,DBGrids
  {$ENDIF}
  ,AmSortie
  ;

type TypeChamp = (tcChaine,tcTT,tcMontant,tcDuree, tcMontantPlus);
type
  TLogLieuGeo = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    LieuGeo : String;
  end;

Type
  TLogEtablissement = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    Etabl : String;
  end;

type
  TLogCession = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateOpe : TDateTime;
    QteCedee : double;
    VoCedee : double;
    PlanActifAv : integer;
    TvaRecuperable : double;
    TvaRecuperee : double;
    Vnc : double;
    RepriseEco : double;
    RepriseFisc : double;
    MontantExc : double;
    TypeExc : string;
    BaseEco : double;
    BaseFisc : double;
    BaseTP: double ;
    RepriseSBV : double;  // ajout mbo pour stockage antérieurs saisis subvention
// 18/09/07 Nouveaux paramètres immos
    JourSortie : string;
// 18/09/07 Fin Nouveaux paramètres immos
  end;

type
  PLogChPlan = ^TLogChPlan;

  TLogChPlan = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateOpe : TDateTime;
    PlanActifAv : integer;
    PlanActifAp : integer;
    MontantExc : double;
    TypeExc: string;
    RevisionEco : double;
    RevisionFisc : double;
    DureeEco : integer;
    DureeFisc : integer;
    MethodeEco : string;
    MethodeFisc : string;
    DatePieceA : TDateTime;
    DateAmort : TDateTime;
    // BTY 11/05 Cas du calcul plan futur avec VNC et Plan fiscal CRC 2002-10
    CodeEclat : string;
    BaseFiscAvMB : double;
    RepriseEco : double;
    RepriseFisc : double;
    DureePriSbv : double; // durée d'inaliénabilité
    // 04/07 Nelle gestion fiscale
    CalculVNF : string;
    GestionFiscale : string;
    NumVersion : string;
    RepriseDR : double;
    RepriseFEC : double;
  end;

type
  TLogMutation = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateOpe : TDateTime;
    CodeMutation : string;
    CpteMutation : string;
  end;

Type
  TLogEclatement = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateOpe : TDateTime;
    CodeEclate : string;
    MontantEclate : double;
    QteEclate : double;
    PlanActifAv : integer;
  end;

// TGA 10/2005 CRC200210
Type
  TLogDepreciation = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    depreciation : double;
    PlanActifAv : integer;
  end;

// mbo fq 17923 04/2006
Type
  TLogModifService = record
    //Entete
    Lequel : string;
    Ordre : integer;
    DebEco : string;
    DebFis : string;
    OrdreSerie : integer;
    TypeOpe : string;

    DateOpReelle : TDateTime;  // date avt modif
    PlanActifAv : integer;
  end;


//PGR 11/2005 Changement des conditions d'amortissement
Type
  TLogChgtCondAmort = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    PlanActifAv : integer;
    MethodeEco : string;
    DureeEco : integer;
    DateOpReelle : TDateTime;
    Duree : string;
    RepriseEco : double;
    RepriseFisc : double;
    // ajout mbo fq 17569
    datedebeco : string;
    datedebfis : string;
 end;


Type
  TLogLeveeOpt = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateFinCt : TDateTime; // FQ 19280
  end;

type
  TLogModifBases = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    DateOpe : TDateTime;
    MontantHT : double;
    BaseEco : double;
    BaseFisc : double;
    PlanActifAv : integer;
    RepriseEco : double;
    RepriseFisc : double;
// BTY 10/05 Annulation de modif base avec plan fiscal CRC 2002-10
    MethodeFisc : string;
    DureeFisc : integer;
// BTY 03/06 Champ 'Calcul plan futur avec VNC' à restituer
    CodeEclat : string;
    // 04/07 Nelle gestion fiscale
    CalculVNF : string;
    GestionFiscale : string;
    NumVersion : string;
    RepriseDR : double;
    RepriseFEC : double;
  end;

// BTY 04/06 FQ 17516
type
  TLogGroupe = record
//Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
//Corps
    Regroupement : String;
  end;

Type
  TLogPrime = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    Prime : double;
    PlanActifAv : integer;
    MethodeEco : string;
    DureeEco : integer;
  end;

Type
  TLogSBV = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    Subvention : double;
    PlanActifAv : integer;
    MethodeEco : string;
    DureeEco : integer;
  end;

// Réduction de prime d'équipement
Type
  TLogReducPrime = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    Prime : double;
    Base : double;
    Suramort : double;
    PlanActifAv : integer;
  end;

// Réduction de la subvention d'investissement
Type
  TLogReducSBV = record
    //Entete
    Lequel : string;
    Ordre : integer;
    OrdreSerie : integer;
    TypeOpe : string;
    //Corps
    SBV : double;
    Reprise : double;
    PlanActifAv : integer;
  end;


Type
  TOpeEnCours = class(TForm)
    HPanel1: THPanel;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    HPageControl1: TPageControl;
    General: TTabSheet;
    HDetail: THGrid;
    HPanel2: THPanel;
    InfoImmo: TLabel;
    HPanel3: THPanel;
    BImprimer: TToolbarButton97;
    bSupprime: TToolbarButton97;
    BValider: TToolbarButton97;
    bFerme: TToolbarButton97;
    ToolbarButton973: TToolbarButton97;
    bAccedeImmo: TToolbarButton97;
    FListe: THGrid;
    TabSheet1: TTabSheet;
    BlocNote: THRichEditOle;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure bSupprimeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure bAccedeImmoClick(Sender: TObject);
    procedure ToolbarButton973Click(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FListeRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
  private
    { Déclarations privées }
    fLequel0 : string;
    fSuppression : boolean;
    fOrdre : integer;
    fOrdreSerie : integer;
    CodeImmoDest : string;
    fTypeOpe : string;
    fbTrouveDialogOpeEnCours : boolean;
    fTypeAction : TActionFiche;
    LogLieuGeo : TLogLieuGeo;
    // TGA 10/2005 CRC200210
    LogDepreciation : TLogDepreciation;
    //MBO FQ 17923
    LogModifService : TLogModifService;
    //PGR 11/2005 Changement des conditions d'amortissement
    LogChgtCondAmort : TLogChgtCondAmort;
    // BTY 04/06 FQ 17516
    LogGroupe : TLogGroupe;
    LogEtabl : TLogEtablissement;
    LogEclatement : TLogEclatement;
    LogMutation : TLogMutation;
    LogChPlan : TLogChPlan;
    LogCession : TLogCession;
    LogLeveeOpt : TLogLeveeOpt;
    LogModifBases : TLogModifBases;
    bChangement: boolean ;
    fListeOperation : TOB;
    fLigneOperation : integer;
    AncienLogCDA : boolean;
    LogPrime : TLogPrime;      // ajout mbo pour la prime 11.09.2006
    LogSBV : TLogSBV;          // ajout mbo pour la subvention 11.09.2006
    LogReducPrime : TLogReducPrime;      // 10/06 BTY Réduction de la prime
    LogReducSBV : TLogReducSBV;          // 11/06 BTY Réduction de la subvention
    procedure AfficheDetail;
    procedure AnnuleDerniereOperation;
    procedure AnnuleCession;
    procedure EnregAnnuleOp ;
    procedure AnnuleEtablissement(LogEtabl : TLogEtablissement);
    procedure AnnuleLieu(LogLieuGeo : TLogLieuGeo);
    // BTY 04/06 FQ 17516
    procedure AnnuleRegroupement (LogGroupe : TLogGroupe);
    procedure AffecteColonneOperation (Col : integer;Champ : string;tc :TypeChamp;TT:String);
    function TraiteAnnulOpeSerie : TIOErr;
    procedure TraiteEnregLog(T : TOB);
    procedure TraiteEnregLogCession(T : TOB);
    procedure TraiteEnregLogMutation(T : TOB);
    procedure TraiteEnregLogEclatement(T : TOB);
    procedure TraiteEnregLogEtablissement(T : TOB);
    procedure TraiteEnregLogLieuGeo(T : TOB);
    procedure TraiteEnregLogLeveeOption(T : TOB);
    procedure TraiteEnregLogModifBases(T : TOB);
    // TGA 10/2005 CRC200210
    procedure TraiteEnregLogdepreciation(T : TOB);
    procedure AnnuleDepreciation(LogDepreciation : TLogDepreciation);
    //PGR 11/2005 Changement des conditions d'amortissement
    procedure TraiteEnregLogChgtCondAmort(T : TOB);
    procedure AnnuleChgtCondAmort(LogChgtCondAmort : TLogChgtCondAmort);
    // BTY 04/06 FQ 17516
    procedure TraiteEnregLogRegroupement(T : TOB);
    // MBO 04/06 FQ 17923
    procedure TraiteEnregLogService(T : TOB);
    procedure AnnuleModifService(LogModifService : TLogModifService);
    // mbo 11/09/2006
    procedure TraiteEnregLogPrime(T : TOB);
    procedure AnnulePrime(LogPrime : TLogPrime);
    // mbo 09/10/2006
    procedure TraiteEnregLogSBV(T : TOB);
    procedure AnnuleSBV(LogSBV : TLogSBV);
    // BTY 10/06 Réduction de prime
    procedure TraiteEnregLogReducPrime(T : TOB);
    procedure TraiteEnregLogReducSBV(T : TOB);
    // BTY 10/06 Réduction de prime
    procedure AnnuleReductionPrime(LogReducPrime : TLogReducPrime);
    procedure AnnuleReductionSBV(LogReducSBV : TLogReducSBV);
  public
    { Déclarations publiques }
  end;
procedure TraiteEnregLogChPlan(T : TOB; var St : TLogChPlan); overload;
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TraiteEnregLogChPlan(Q : TQuery ; var St : TLogChPlan); overload;
{$ENDIF}
function  OperationsEnCours (Lequel0, Libelle0 : string;bSuppression : boolean;LeMode : TActionFiche): boolean ;

implementation
uses ChanPlan, ImEclate, ImMutati, Outils, IMMO_TOM, (*ImSelSup,*) ImLevOpt, ImModBas, ImPlan;
{$R *.DFM}
function GetWindowOperation(Window: HWnd;Param : Longint): WordBool; stdcall;
Var lpClassName : Array[0..80] Of Char;
begin
  result := true;
  GetClassName(Window,lpClassName,80);
  if lpClassName = 'TOpeEnCours' then result := false;
end;

function OperationsEnCours(Lequel0,Libelle0: string; bSuppression: boolean; LeMode: TActionFiche): boolean;
var
  OpeEnCours: TOpeEnCours;
  PP : THPanel ;
  bFound : boolean;
begin
  bFound := not EnumThreadWindows(GetCurrentThreadID, @GetWindowOperation, 0);
  OpeEnCours := TOpeEnCours.Create(Application);
  OpeEnCours.InfoImmo.Caption := Lequel0+' - '+Libelle0;
  OpeEnCours.fLequel0 := Lequel0;
  OpeEnCours.fSuppression := bSuppression;
  OpeEnCours.fbTrouveDialogOpeEnCours := bFound;
  OpeEnCours.fTypeAction := LeMode;
  if bSuppression then
  begin
    OpeEnCours.bValider.Hint:=OpeEnCours.HM.Mess[0];
    OpeEnCours.Caption:=OpeEnCours.HM.Mess[3];
    OpeEnCours.bSupprime.Visible:= (LeMode <> taConsult);
  end
  else
  begin
    OpeEnCours.bValider.Visible:=false;
    OpeEnCours.BImprimer.Left:=OpeEnCours.bValider.Left;
    OpeEnCours.bSupprime.Visible:= (LeMode <> taConsult);
  end;

  PP:=FindInsidePanel ;
  if PP=Nil then
  BEGIN
    try
      OpeEnCours.ShowModal;
    finally
      result:=OpeEnCours.bChangement ;
      OpeEnCours.Free;
    end;
  END else
  BEGIN
    InitInside(OpeEnCours,PP) ;
    OpeEnCours.Show ;
    result:=OpeEnCours.bChangement ;
  END ;
end;

procedure TOpeEnCours.FormShow(Sender: TObject);
var QTmp : TQuery;
    Requete,Where,WhereDate,Selection,Jointure,ListeChampImmoLog,NumDerniereOpe : string;
begin
  bChangement:=false ;
  {$IFDEF SERIE1}
  HelpContext:=511020 ;
  {$ELSE}
  if (fSuppression) then HelpContext:=2101600 else HelpContext := 2101500;
  {$ENDIF}
  bSupprime.Visible := (not fbTrouveDialogOpeEnCours) and (not fSuppression) and (fTypeAction <> taConsult);
  QTmp := OpenSql('SELECT MAX(IL_ORDRE) MAXNUM FROM IMMOLOG WHERE IL_IMMO="'+fLequel0+'"',TRUE);
  NumDerniereOpe := QTmp.FindField('MAXNUM').AsString;
  Ferme(QTmp);
  //PGR 11/2005 Changement des conditions d'amortissement : Ajout IL_DATEOPREELLE IL_DUREE IL_REVISIONREPECO
  //PGR 12/2005 Changement des conditions d'amortissement : Ajout IL_CUMANTCESECO IL_CUMANTCESFIS
  ListeChampImmoLog:= 'IL_IMMO,IL_REPRISEECO,IL_REPRISEFISC,IL_PLANACTIFAV,IL_DATEOPREELLE,IL_DUREE,IL_REVISIONREPECO,'
      +'IL_PLANACTIFAP,IL_TVARECUPEREE,IL_TVARECUPERABLE,IL_ORDRESERIE,IL_VNC,IL_CUMANTCESECO,IL_CUMANTCESFIS,'
      +'IL_ORDRE,IL_TYPEOP,IL_DATEOP,IL_TYPEMODIF,IL_MOTIFCES,IL_BLOCNOTE,'
      +'IL_DATEOP,IL_CPTEMUTATION,IL_CODEMUTATION,IL_VOCEDEE,IL_MONTANTCES,'
      +'IL_PVALUE,IL_QTECEDEE,IL_MONTANTECL,IL_QTEECLAT,IL_CODEECLAT,'
      +'IL_METHODEECO,IL_DUREEECO,IL_REVISIONECO,IL_REVISIONFISC,IL_ETABLISSEMENT,'
      +'IL_METHODEFISC,IL_DUREEFISC,IL_TYPEDOT,IL_MONTANTDOT,IL_CODECB,IL_LIEUGEO,IL_CUMANTCESECO,IL_CUMANTCESFIS,'
      +'IL_MONTANTAVMB,IL_BASEECOAVMB,IL_BASEFISCAVMB,IL_MONTANTEXC,IL_TYPEEXC,IL_BASETAXEPRO,IL_LIBELLE,IL_VERSION,IL_REVISIONDOTECO';
  Selection:='SELECT '+ListeChampImmoLog+',CO_LIBELLE ';
  Jointure:='FROM IMMOLOG LEFT JOIN COMMUN ON IL_TYPEOP=CO_CODE ';
  if fSuppression then
     Where:='(IL_IMMO="'+fLequel0+'") AND (IL_ORDRE='+NumDerniereOpe+'))'
  else Where:='(IL_IMMO="'+fLequel0+'")) ORDER BY IL_ORDRE';
  WhereDate := '(IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" AND IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+'")';
  Where:='WHERE '+WhereDate+' AND ((CO_TYPE="TOA") AND (CO_CODE<>"ACQ") AND (CO_CODE<>"CLO") AND '+Where;

  Requete:=Selection+Jointure+Where;
  QTmp := OpenSQL ( Requete , True );
  try
    fListeOperation := TOB.Create ('', nil, - 1);
    fListeOperation.LoadDetailDB ('', '', '', QTmp, False);
  finally
    Ferme (QTmp);
  end;
  if (fListeOperation.Detail.Count > 0) then
  begin
    fLigneOperation := 1;
    fListeOperation.PutGridDetail(FListe,False,False,'CO_LIBELLE;IL_DATEOP;IL_LIBELLE');
    AfficheDetail;
  end
  else ModalResult:=mrCancel;
end;

procedure TOpeEnCours.AffecteColonneOperation (Col : integer; Champ : string; tc :TypeChamp;TT:String);
var Info:string;
    ColSize : integer;
    TLog : TOB;
    TitreColonne: string;
    // 18/09/07 Nouveaux paramètres immos
    JourSortie:Boolean;
    // 18/09/07 Fin Nouveaux paramètres immos
begin
  TLog := fListeOperation.Detail[fLigneOperation-1];
  HDetail.ColAligns[Col]:= taCenter;
  // HDetail.Cells[Col,0] :=ChampToLibelle(Champ );
  // BTY 04/06 FQ 17516 Regroupement stocké dans IL_LIEUGEO => modifier titre de cette colonne
  if Tlog.GetValue('IL_TYPEOP') = 'REG' Then
     TitreColonne := 'Ancien Regroupement'
  else
     TitreColonne := ChampToLibelle(Champ);

  HDetail.Cells[Col,0] := TitreColonne;

  // 18/09/07 Nouveaux paramètres immos
  JourSortie := ((Tlog.GetValue('IL_TYPEOP') ='CES') and (Tlog.GetValue('IL_MOTIFCES') <> '999')
                  and (Col = 5) and (Champ = 'IL_CODECB') );
  // 18/09/07 Fin Nouveaux paramètres immos

  if Tlog.GetValue('IL_TYPEOP') = 'MMS' Then
  begin
     if col = 1 then
        HDetail.Cells[Col,0] := 'Début Amort éco.'
     else
     begin
        if col = 2 then
           HDetail.Cells[Col,0] := 'Début amort. fiscal'
        else
           HDetail.Cells[Col,0] := 'Mise en service';
     end;
  end;

  if Tlog.GetValue('IL_TYPEOP') = 'PRI' Then
  begin
     if col = 1 then
        HDetail.Cells[Col,0] := 'Méthode amort.'
     else
     begin
        if col = 2 then
           HDetail.Cells[Col,0] := 'Durée d''amortissement.'
        else
           HDetail.Cells[Col,0] := 'Prime';
     end;
  end;

  if Tlog.GetValue('IL_TYPEOP') = 'SBV' Then
  begin
     if col = 1 then
        HDetail.Cells[Col,0] := 'Méthode amort.'
     else
     begin
        if col = 2 then
           HDetail.Cells[Col,0] := 'Durée d''amortissement.'
        else
           HDetail.Cells[Col,0] := 'Subvention';
     end;
  end;

  // BTY 10/06 Réduction de prime
  if Tlog.GetValue('IL_TYPEOP') = 'RPR' Then
     begin
     case col of
       0 :  HDetail.Cells[Col,0] := 'Prime' ;
       1 :  HDetail.Cells[Col,0] := 'Base' ;
       end;
     end;

  // BTY 11/06 Réduction de subvention
  if Tlog.GetValue('IL_TYPEOP') = 'RSB' Then
     begin
     case col of
       0 :  HDetail.Cells[Col,0] := 'Subvention' ;
       1 :  HDetail.Cells[Col,0] := 'Reprise déjà pratiquée' ;
       end;
     end;


  // fq 17569 mbo
  if (Tlog.GetValue('IL_TYPEOP') = 'CDA') and not (AncienLogCDA) then
  begin
     if col = 0 then
        HDetail.Cells[Col,0] := 'Méthode éco.'
     else
     begin
        if col = 1 then
           HDetail.Cells[Col,0] := 'Durée éco.'
        else if col = 2 then
           HDetail.Cells[Col,0] := 'Mise en service'
        else if col = 3 then
           HDetail.Cells[Col,0] := 'Début amort. éco.'
        else
           HDetail.Cells[Col,0] := 'Début amort. fiscal';
     end;
  end;

  // 18/09/07 Nouveaux paramètres immos
  // Colonne supplémentaire en cession
  if JourSortie then HDetail.Cells[5,0] := 'Amort. J sortie';
  // 18/09/07 Fin Nouveaux paramètres immos

  if (Champ='IL_QTECEDEE') and (TLog.GetValue(Champ) = 0) then Info:=StrFMontant(1,0,V_PGI.OkDecV,'',True)
  else
  if tc = tcMontant then
    Begin
     // TGA 16/11/2005 affichage de la reprise de dépréciation en valeur absolue
     if Tlog.GetValue('IL_TYPEOP') ='DPR' Then
       Info := StrFMontant(ABS(TLog.GetValue(Champ)),0,V_PGI.OkDecV,'',True)
     Else
       Info := StrFMontant(TLog.GetValue(Champ),0,V_PGI.OkDecV,'',True)
    end
  else
    // 18/09/07 Nouveaux paramètres immos
    // Cessions faites avant le paramètre SO_AMORTJSORTIE
    if (JourSortie) and (Tlog.GetValue('IL_CODECB')='') then Info := 'OUI'
    // 18/09/07 Fin Nouveaux paramètres immos
  else if tc=tcMontantPlus then info:=StrFMontant(TLog.GetValue(Champ),0,V_PGI.OkDecV,'',True) + '     '
  else if tc=tcTT then Info := RechDom(TT,TLog.GetValue(Champ),False)
  else if tc = tcDuree then Info :=IntToStr(TLog.GetValue(Champ))+' '+HM.Mess[18]
  else if not VarIsNull(TLog.GetValue(Champ)) then Info := TLog.GetValue(Champ)
  else Info := '';

  // BTY 04/06 FQ 17516 Regroupement stocké dans IL_LIEUGEO => modifier titre de cette colonne
  //ColSize := MaxIntValue ([HDetail.Canvas.TextExtent(ChampToLibelle(Champ )).cx,HDetail.Canvas.TextExtent(Info).cx]);
  ColSize := MaxIntValue ([HDetail.Canvas.TextExtent(TitreColonne).cx,HDetail.Canvas.TextExtent(Info).cx]);
  HDetail.ColWidths[Col]:= ColSize+10;
  // 18/09/07 Nouveaux paramètres immos
  if JourSortie then HDetail.ColWidths[Col]:= ColSize+28;
  // 18/09/07 Fin Nouveaux paramètres immos

  HDetail.Cells[Col,1] := Info;
end;

procedure TOpeEnCours.AfficheDetail;
var stTypeOpe : string;
  affmont : double;
  info : string;
  QTmp : TQuery;
  CodeImmo : String;
 begin
  if (fListeOperation=nil) or ((fLigneOperation-1) >= fListeOperation.Detail.Count) then exit;
  stTypeOpe := fListeOperation.Detail[fLigneOperation-1].GetValue('IL_TYPEOP');
  bSupprime.Enabled:=(fTypeAction<>taConsult) and (stTypeOpe<>'EUR');

  bAccedeImmo.Visible:=false; // seulement pour ECL et MUT

  if (stTypeOpe='CES') or (stTypeOpe='CEP') then
  begin
    HDetail.ColCount := 5;
    AffecteColonneOperation (0,'IL_MOTIFCES',tcTT,'TIMOTIFCESSION');
    AffecteColonneOperation (1,'IL_VOCEDEE',tcMontant,'');
    AffecteColonneOperation (2,'IL_MONTANTCES',tcMontant,'');
    AffecteColonneOperation (3,'IL_PVALUE',tcMontant,'');
    AffecteColonneOperation (4,'IL_QTECEDEE',tcMontant,'');

// 18/09/07 Nouveaux paramètres immos
    if (stTypeOpe='CES') and (fListeOperation.Detail[fLigneOperation-1].GetValue('IL_MOTIFCES') <> '999') then
    begin
      HDetail.ColCount := 6;
      AffecteColonneOperation (5,'IL_CODECB',tcChaine,'');
    end;
// 18/09/07 Fin Nouveaux paramètres immos

  end
  else if (stTypeOpe='ECL') then
  begin
    HDetail.ColCount := 3;
    AffecteColonneOperation (0,'IL_MONTANTECL',tcMontant,'');
    AffecteColonneOperation (1,'IL_QTEECLAT',tcMontant,'');
    AffecteColonneOperation (2,'IL_CODEECLAT',tcChaine,'');
    CodeImmoDest := fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODEECLAT');
    bAccedeImmo.Visible:=true; // seulement pour ECL et LEV
  end
  else if (stTypeOpe='MUT') then
  begin
    HDetail.ColCount := 2;
    AffecteColonneOperation (0,'IL_CPTEMUTATION',tcChaine,'');
    AffecteColonneOperation (1,'IL_CODEMUTATION',tcChaine,'');
    if not VarIsNull(fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODEMUTATION')) then
      CodeImmoDest := fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODEMUTATION')
    else CodeImmoDest := '';
  end
  else if (stTypeOpe='LEV') then
  begin
    HDetail.ColCount := 1;
    AffecteColonneOperation (0,'IL_CODECB',tcChaine,'');
    if not VarIsNull(fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODECB')) then
      CodeImmoDest := fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODECB')
    else CodeImmoDest := '';
    bAccedeImmo.Visible:=true; // seulement pour ECL et LEV
  end
  else if (stTypeOpe='ELE') OR (stTypeOpe='ELC') OR (stTypeOpe='EEC') then //EPZ 26/10/98
  begin
    HDetail.ColCount := 1;
    // mbo FQ 16782 AffecteColonneOperation (0,'IL_MONTANTDOT',tcMontant,'');
    affmont :=fListeOperation.Detail[fLigneOperation-1].GetValue('IL_MONTANTDOT') ;
    if fListeOperation.Detail[fLigneOperation-1].GetValue('IL_TYPEDOT') = 'RDO' then affmont := affmont*(-1);
    HDetail.ColAligns[0]:= taCenter;
    HDetail.Cells[0,0] :=ChampToLibelle('IL_MONTANTDOT');
    Info := StrFMontant(affmont,0,V_PGI.OkDecV,'',True);
    HDetail.ColWidths[0]:= MaxIntValue ([HDetail.Canvas.TextExtent(ChampToLibelle('IL_MONTANTDOT')).cx,HDetail.Canvas.TextExtent(Info).cx]);
    HDetail.ColWidths[0]:=HDetail.ColWidths[0]+10;
    HDetail.Cells[0,1] := Info;

  end
  //else if (stTypeOpe='CDM') then
  else if ((stTypeOpe='CDM') or (stTypeOpe='CD2') or (stTypeOpe='CD3')) then
  begin
    HDetail.ColCount := 3;
    AffecteColonneOperation(0,'IL_METHODEECO',tcTT,'TIMETHODEIMMO');
    AffecteColonneOperation(1,'IL_DUREEECO',tcDuree,'');
    AffecteColonneOperation(2,'IL_REVISIONECO',tcMontant,'');
    if not VarIsNull(fListeOperation.Detail[fLigneOperation-1].GetValue('IL_METHODEFISC')) then
    begin
      if (fListeOperation.Detail[fLigneOperation-1].GetValue('IL_METHODEFISC')<>'') then
      begin
           HDetail.ColCount := 6;
           AffecteColonneOperation(3,'IL_METHODEFISC',tcTT,'TIMETHODEIMMO');
           AffecteColonneOperation(4,'IL_DUREEFISC',tcDuree,'');
           AffecteColonneOperation(5,'IL_REVISIONFISC',tcMontant,'');
      end;
    end;
  end
  else if (stTypeOpe='ETA') then
  begin
    HDetail.ColCount := 1;
    AffecteColonneOperation(0,'IL_ETABLISSEMENT',tcTT,'TTETABLISSEMENT');
  end
  else if (stTypeOpe='LIE') then
  begin
    HDetail.ColCount := 1;
    AffecteColonneOperation(0,'IL_LIEUGEO',tcTT,'TILIEUGEO');
  end
  // BTY 10/05 Opération modif bases avec plan fiscal CRC2002-10
  //else if (stTypeOpe='MBA') then
  else if (stTypeOpe='MBA') or (stTypeOpe='MB2') then
  begin
     HDetail.ColCount := 3;
     AffecteColonneOperation (0,'IL_MONTANTAVMB',tcMontant,'');
     AffecteColonneOperation (1,'IL_BASEECOAVMB',tcMontant,'');
     AffecteColonneOperation (2,'IL_BASEFISCAVMB',tcMontant,'');
  end
  //PGR 11/2005 Changement des conditions d'amortissement
  else if (stTypeOpe='CDA') then
  begin
     // modif mbo fq 17569
     if not VarIsNull(fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODEMUTATION'))
     and (fListeOperation.Detail[fLigneOperation-1].GetValue('IL_CODEMUTATION')='') then
        AncienLogCDA := true
     else
        AncienLogCDA := false;

     if (AncienLogCDA) then
     begin
        HDetail.ColCount := 3;
        AffecteColonneOperation(0,'IL_METHODEECO',tcTT,'TIMETHODEIMMO');
        AffecteColonneOperation(1,'IL_DUREEECO',tcDuree,'');
        AffecteColonneOperation(2,'IL_DATEOPREELLE',tcChaine,'');
     end else
     begin
        HDetail.ColCount := 5;
        AffecteColonneOperation(0,'IL_METHODEECO',tcTT,'TIMETHODEIMMO');
        AffecteColonneOperation(1,'IL_DUREEECO',tcDuree,'');
        AffecteColonneOperation(2,'IL_DATEOPREELLE',tcChaine,'');
        AffecteColonneOperation(3,'IL_CODEMUTATION',tcChaine,'');  // date deb éco
        AffecteColonneOperation(4,'IL_CODEECLAT',tcChaine,''); // date deb fiscale
     end;
  end
  else if (stTypeOpe='DPR') then
  begin
     HDetail.ColCount := 1;
     AffecteColonneOperation (0,'IL_MONTANTDOT',tcMontant,'');
  end
  // BTY 04/06 FQ 17516 Changement de regroupement devient une opération
  else if (stTypeOpe='REG') then
  begin
    HDetail.ColCount := 1;
    AffecteColonneOperation(0,'IL_LIEUGEO',tcTT,'AMREGROUPEMENT');
  end
  // FQ 17923 - mbo - ajout mise à jour date de mise en service
  else if (stTypeOpe='MMS') then
  begin
    HDetail.ColCount := 3;
    AffecteColonneOperation(0,'IL_DATEOPREELLE',tcChaine,'');
    AffecteColonneOperation(1,'IL_CODEMUTATION',tcChaine,'');
    AffecteColonneOperation(2,'IL_CODEECLAT',tcChaine,'');
  end

  else if (stTypeOpe ='PRI') then   // ajout mbo pour la prime d'équipement
  begin
    HDetail.ColCount := 3;
    AffecteColonneOperation(0,'IL_MONTANTEXC',tcMontantPlus,'');
    AffecteColonneOperation(1,'IL_METHODEECO',tcTT,'TIMETHODEIMMO');
    AffecteColonneOperation(2,'IL_DUREEECO',tcDuree,'');
  end

  else if (stTypeOpe ='SBV') then   // ajout mbo pour la subvention d'équipement
  begin
    HDetail.ColCount := 3;
    AffecteColonneOperation(0,'IL_MONTANTEXC',tcMontantPlus,'');
    AffecteColonneOperation(1,'IL_METHODEECO',tcTT,'TIMETHODEIMMO');
    AffecteColonneOperation(2,'IL_DUREEECO',tcDuree,'');
  end
  else if (stTypeOpe ='RPR') then       // BTY Réduction de prime
  begin
    HDetail.ColCount := 2;
    AffecteColonneOperation(0,'IL_MONTANTEXC',tcMontantPlus,'');
    AffecteColonneOperation(1,'IL_BASETAXEPRO',tcMontantPlus,'');
  end
  else if (stTypeOpe ='RSB') then      // BTY Réduction de subvention
  begin
    HDetail.ColCount := 2;
    AffecteColonneOperation(0,'IL_MONTANTEXC',tcMontantPlus,'');
    AffecteColonneOperation(1,'IL_MONTANTAVMB',tcMontantPlus,'');
  end;

  // TGA affichage du Bloc-note
  // FQ 19091 Préciser la date sinon on prend le 1er bloc-note pour la même opération
  codeimmo:=fListeOperation.Detail[fLigneOperation-1].GetValue('IL_IMMO');
  Blocnote.clear;
  QTmp := nil;
  Begin
    try
      QTmp := OpenSql('SELECT IL_BLOCNOTE FROM IMMOLOG WHERE ' +
            '(IL_IMMO="'+CodeImmO+'")AND'+'(IL_TypeOp="'+stTypeOpe+'") AND ' +
            '(IL_DATEOP>="'+USDateTime(VHImmo^.Encours.Deb)+'" AND ' +
            ' IL_DATEOP<="'+USDateTime(VHImmo^.Encours.Fin)+'")', True);

      // BTY 12/10/05 Correction pour CWAS
      StringToRich (BlocNote,QTmp.FindField('IL_BLOCNOTE').AsString);
      //BlocNote.Lines.Assign(QTmp.FindField('IL_BLOCNOTE'));
    finally
      Ferme(QTmp);
    end;
  end;

end;

//-----------------------------------------------------------------------------------

procedure TOpeEnCours.BValiderClick(Sender: TObject);
begin
  if (fSuppression and (not fbTrouveDialogOpeEnCours)) then
      begin
      AnnuleDerniereOperation;
      // BTY FQ 18963  Indicateur de sortie correcte d'une annulation d'opération
      bChangement:=true ;
    end;
end;

//-----------------------------------------------------------------------------------

procedure TOpeEnCours.AnnuleDerniereOperation;
var OldVErr : TIOErr;
  LaSortie : TAmSortie;
begin
  OldVErr:=V_PGI.IoError;
  if fTypeOpe='EEC' then HM.execute(24,'','')
  else
    begin
    BEGINTRANS ;
    try
      if (fTypeOpe = 'CES') then
      begin
        LaSortie := TAmSortie.Create (fListeOperation.Detail[fLigneOperation-1].GetValue('IL_IMMO'));
        try
          if LaSortie.EstSortie then
          begin
            LaSortie.Execute;
          end;
        finally
          LaSortie.Free;
        end;
      end else
      begin
        TraiteEnregLog(fListeOperation.Detail[fLigneOperation-1]);
        if (fTypeOpe='CES') or (fTypeOpe='CEP') then AnnuleCession else EnregAnnuleOp ;
        ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+fLequel0+'" AND IL_ORDRE='+IntToStr(fOrdre)) ;
        TraiteAnnulOpeSerie ;
      end;
      COMMITTRANS ;
      // FQ 18119 Positionner les indicateurs de modif
      VHImmo^.ChargeOBImmo := True;
      ImMarquerPublifi (True);
    except
      HM.execute(2,Caption,'') ;
      ROLLBACK ;
      end ;
    V_PGI.IoError := OldVErr;
    end ;
end;

procedure TOpeEnCours.EnregAnnuleOp ;

{BTY 04/06 Pour diminuer messages compilateur
var Q: TQuery;
    QueryImmo : TQuery;
    codemut: string;
    ordre : string;}

begin
  if fTypeOpe ='MUT' then fLequel0 := AnnuleMutation(LogMutation,HM)
  else	if fTypeOpe ='ECL' then AnnuleEclatement(LogEclatement,HM)
  else	if fTypeOpe ='ETA' then AnnuleEtablissement(LogEtabl)
  else	if fTypeOpe ='LIE' then AnnuleLieu(LogLieuGeo)
  //PGR 11/2005 Changement des conditions d'amortissement
  else	if fTypeOpe ='CDA' then  AnnuleChgtCondAmort(LogChgtCondAmort)
  // TGA 10/2005 CRC200210
  else	if fTypeOpe ='DPR' then  AnnuleDepreciation(LogDepreciation)
  else  if fTypeOpe ='LEV' then AnnuleLeveeOption(LogLeveeOpt,HM)
  else  if OpeChangementPlan(fTypeOpe) then AnnuleChangementPlan(LogChPlan)
  // BTY 04/06 FQ 17516 Annuler l'opération de changement de regroupement
  else  if fTypeOpe = 'REG' then AnnuleRegroupement(LogGroupe)
  else  if fTypeOpe = 'MMS' then AnnuleModifService(LogModifService)
  // BTY 10/05 Opération modif bases avec plan fiscal CRC2002-10
  // else if fTypeOpe ='MBA' then AnnuleModifBases (LogModifBases,HM);
  else  if (fTypeOpe ='MBA') OR (fTypeOpe ='MB2') then AnnuleModifBases (LogModifBases,HM)
  else if fTypeOpe = 'PRI' then AnnulePrime(LogPrime)
  else if fTypeOpe = 'SBV' then AnnuleSBV(LogSBV)
  else if fTypeOpe = 'RPR' then AnnuleReductionPrime(LogReducPrime) // BTY 10/06
  else if fTypeOpe = 'RSB' then AnnuleReductionSBV(LogReducSBV);    // BTY 11/06
end;

procedure TOpeEnCours.AnnuleCession;
var Q: TQuery;
    QteCedee, TValeur, VoCedee,TTvaR, TTvaE,wMttHt: double ; // BTY 04/06 ,MntExc : double ;
    Plan : TPlanAmort;
    PlanCourant,NumPlanPrec : string;
    Composant : boolean;
begin

  Plan:=TPlanAmort.Create(true) ;
  Q:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogCession.Lequel+'"', FALSE);
  if not Q.EOF then
  begin
    Q.Edit;

    // ajout mbo pour annulation remplacement de composant
    //ajout mbo pour suppression opération remplacement de composant
    if trim(Q.FindField('I_REMPLACEE').AsString) <>'' then
    begin
       Composant := true;

       //1) suppression du composant ds immoamor
       ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Q.FindField('I_REMPLACEE').AsString+'"');
       //2) suppression du composant immo
       ExecuteSQL('DELETE FROM IMMO WHERE I_IMMO="'+Q.FindField('I_REMPLACEE').AsString+'"');
       //3) suppression du composant immolog
       ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+Q.FindField('I_REMPLACEE').AsString+'"');
       // pour l'immo d'origine on la traite comme une cession normale + réinitialisation des zones composant
    end else
       Composant := false;

    QteCedee:= LogCession.QteCedee;
    if (QteCedee=0) and (LogCession.TypeOpe='CES') then QteCedee:=1;
    VoCedee := LogCession.VoCedee;
    NumPlanPrec:=IntToStr(LogCession.PlanActifAv);
    PlanCourant:=Q.FindField('I_PLANACTIF').AsString;
    // ajout mbo - FQ 16780
    //MntExc:=Q.FindField('I_MONTANTEXC').AsFloat; BTY 04/06 Pour diminuer messages compilateur
    Q.FindField('I_VNC').AsFloat:=Q.FindField('I_VNC').AsFloat+LogCession.Vnc;
    Q.FindField('I_QUANTITE').AsFloat:=Q.FindField('I_QUANTITE').AsFloat+QteCedee;
    // FQ 20557 Q.FindField('I_QTCEDE').AsFloat:= Q.FindField('I_QTCEDE').AsFloat-QteCedee;
    Q.FindField('I_QTCEDE').AsFloat:= 0;
    wMttHt:=Q.FindField('I_MONTANTHT').AsFloat+VoCedee; //YCP 10/09/01 +LogCession.TvaRecuperee-LogCession.TvaRecuperable ;

    // fq 16897 - tva doublée suite annulation de cession - mbo 3.01.2005
    //TTvaR:=Q.FindField('I_TVARECUPERABLE').AsFloat + LogCession.TvaRecuperable;
    //TTvaE:=Q.FindField('I_TVARECUPEREE').AsFloat + LogCession.TvaRecuperee;
    TTvaR:=LogCession.TvaRecuperable;
    TTvaE:=LogCession.TvaRecuperee;
    // mbo 03.01.06 calcul montant ht faux
    //Q.FindField('I_MONTANTHT').AsFloat:=wMttHt-TTvaR+TTvaE ;
    Q.FindField('I_MONTANTHT').AsFloat:=wMttHt;

    TValeur := Q.FindField('I_MONTANTHT').AsFloat;
    Q.FindField('I_TVARECUPERABLE').AsFloat:=TTvaR ;
    Q.FindField('I_TVARECUPEREE').AsFloat:=TTvaE ;
    if (LogCession.BaseTP <> 0) then Q.FindField('I_BASETAXEPRO').AsFloat:= LogCession.BaseTP
    else Q.FindField('I_BASETAXEPRO').AsFloat:=TValeur+TTvaR-TTvaE ;
    if (LogCession.BaseEco <> 0) then Q.FindField('I_BASEECO').AsFloat:= LogCession.BaseEco
    else Q.FindField('I_BASEECO').AsFloat:=TValeur+TTvaR-TTvaE ;
    if (LogCession.BaseFisc <> 0) then Q.FindField('I_BASEFISC').AsFloat:= LogCession.BaseFisc
    // FQ 18976 Base fiscale même sans plan fiscal => pb potentiel et message bloquant en cwas
    //else Q.FindField('I_BASEFISC').AsFloat:=TValeur+TTvaR-TTvaE ;
    else if Q.FindField('I_METHODEFISC').AsString = '' then
            Q.FindField('I_BASEFISC').AsFloat:= Arrondi(0, V_PGI.OkDecV)
         else Q.FindField('I_BASEFISC').AsFloat:=TValeur+TTvaR-TTvaE ;

    Q.FindField('I_BASEAMORFINEXO').AsFloat:=Q.FindField('I_BASEAMORFINEXO').AsFloat+VoCedee;
    Q.FindField('I_REPRISEECO').AsFloat:=Q.FindField('I_REPRISEECO').AsFloat + LogCession.RepriseEco;
    Q.FindField('I_REPRISEFISCAL').AsFloat:=Q.FindField('I_REPRISEFISCAL').AsFloat + LogCession.RepriseFisc;

    // FQ 20557
    //Q.FindField('I_REPCEDECO').AsFloat := Q.FindField('I_REPCEDECO').AsFloat - LogCession.RepriseEco;
    //Q.FindField('I_REPCEDFISC').AsFloat := Q.FindField('I_REPCEDFISC').AsFloat - LogCession.RepriseFisc;
    Q.FindField('I_REPCEDECO').AsFloat := 0;
    Q.FindField('I_REPCEDFISC').AsFloat := 0;

    if LogCession.TypeExc = '' then  // Uniquement dans le cas ou l'exceptionnel a eu lieu avant la sortie
    begin
      Q.FindField('I_MONTANTEXC').AsFloat := Q.FindField('I_MONTANTEXC').AsFloat + LogCession.MontantExc;
      Q.FindField('I_MONTANTEXCCED').AsFloat := Q.FindField('I_MONTANTEXCCED').AsFloat - LogCession.MontantExc;
    end else
      // ajout test mbo - fq 16780 on ne met pas à jour opechangeplan si présence d'une dotation exceptionnelle
      //IF MntExc = 0 then  // = dotation exceptionnelle
      if Q.FindField('I_MONTANTEXC').AsFloat = 0 then
         MajOpeEnCoursImmo ( Q, 'I_OPECHANGEPLAN','ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="ELC','-');

    Q.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
    ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+Q.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;
    Plan.Recupere(Q.FindField('I_IMMO').AsString,NumPlanPrec);
    Q.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
    Q.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);

    // TGA 04/04/2006
    Q.FindField('I_REPRISEDEP').AsFloat := Q.FindField('I_REPRISEDEPCEDEE').AsFloat;
    Q.FindField('I_REPRISEDEPCEDEE').AsFloat := 0;
    Q.FindField('I_DATECESSION').AsDateTime := idate1900;
    // BTY 04/06 FQ 17629 Remettre à NOR la règle de calcul de la PMValue
    Q.FindField('I_REGLECESSION').AsString := 'NOR';
    //ajout mbo pour prime d'équipement - 11.09.2006
    Q.FindField('I_SBVPRI').AsFloat := Q.FindField('I_SBVPRIC').AsFloat;
    Q.FindField('I_SBVPRIC').AsFloat := 0;
    Q.FindField('I_REPRISEUO').AsFloat := Q.FindField('I_REPRISEUOCEDEE').AsFloat;
    Q.FindField('I_REPRISEUOCEDEE').AsFloat := 0;
    //ajout mbo pour subvention d'équipement - 09/10/2006
    Q.FindField('I_SBVMT').AsFloat := Q.FindField('I_SBVMTC').AsFloat;
    Q.FindField('I_SBVMTC').AsFloat := 0;
    //ajout mbo pour remettre les antérieurs saisis - 24/10/2006
    Q.FindField('I_CORRECTIONVR').AsFloat := LogCession.RepriseSBV;

    //fq 17512 - mbo - 20/03/2007 - antérieurs dérogatoire et réintégration fiscale
    Q.FindField('I_REPRISEDR').AsFloat := Q.FindField('I_REPRISEFDRCEDEE').AsFloat;
    Q.FindField('I_REPRISEFEC').AsFloat := Q.FindField('I_REPRISEFECCEDEE').AsFloat;
    Q.FindField('I_REPRISEFDRCEDEE').AsFloat := 0;
    Q.FindField('I_REPRISEFECCEDEE').AsFloat := 0;

    // ajout mbo pour annulation de remplacement de composant
    if Composant then
    begin
       // mbo - 02.07.07 attention si immo remplaçante il faut mettre i_typer à irm
       if Q.FindField('I_REMPLACE').AsString <> '' then
          Q.FindField('I_TYPER').AsString:= 'IRM'
       else
          Q.FindField('I_TYPER').AsString:= '';

       Q.FindField('I_REMPLACEE').AsString:= '';
       // on ne remet pas à blanc car l'immo d'origine est peut être elle même un remplacement
       // Q.FindField('I_REMPLACE').AsString :='';
       Q.FindField('I_OPEREMPL').AsString:='-';
    end;

    Plan.free ;
    MajOpeEnCoursImmo ( Q,'I_OPECESSION', 'CEP" OR IL_TYPEOP="CES', '-');
    Q.Post ;
    Ferme(Q);
    end;
end;

procedure TOpeEnCours.bSupprimeClick(Sender: TObject);
begin
  if fLigneOperation<>fListeOperation.Detail.Count then
   // Modif mbo 20.03.2007 PB DELPHI 7 MessageBox(0,PChar(HM.Mess[19]),Pchar(HM.Mess[3]),MB_OK)
   // BTY 04/07 attention HM.Mess[19] était formaté avec son n° de message dans la dfm
    PGIInfo(HM.Mess[19],HM.Mess[3]) //XVI Conseil Compile...
  else if (HM.Execute(1,Caption,'')=mrYes) then
  begin
    AnnuleDerniereOperation;
    bChangement:=true ;
    ModalResult:=mrOk ;
    close ;
  end;
end;

procedure TOpeEnCours.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if fListeOperation <> nil then FreeAndNil(fListeOperation);
  HDetail.VidePile(true);
  if isInside(Self) then Action:=caFree ;
end;

function TOpeEnCours.TraiteAnnulOpeSerie : TIOErr;
var TSerie : TOB;
    QTmp : TQuery;
    OldOrdreSerie,OldLequel : string; rT : TIOErr; OpeIni, OpeSuiv : string;
    iOrdre : integer;
    CodeImmo : string;
    LeIn: string ;
    i : integer;
begin
  rt := oeOK;
  result:= rt;
  OldOrdreSerie:=fListeOperation.Detail[fLigneOperation-1].GetValue('IL_ORDRESERIE');

  // TGA 05/10/2005
  IF OldOrdreSerie='-1' Then
    begin
      result:= rt;
      exit;
    end;
  // Fin TGA 05/10/2005

  OpeIni       :=fListeOperation.Detail[fLigneOperation-1].GetValue('IL_TYPEOP');
  QTmp:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_ORDRESERIE='+OldOrdreSerie,True) ;
  try
    TSerie := TOB.Create ('', nil, -1);
    TSerie.LoadDetailDB('','','',QTmp,False);
  finally
    Ferme (QTmp);
  end;
  OldLequel:=fLequel0;
  if TSerie.Detail.Count=0 then
  begin
    TSerie.Free;
    result:= rt;
    exit;
  end;
  OpeSuiv:=TSerie.Detail[0].GetValue('IL_TYPEOP');

  // BTY 10/05 Traiter le couple ELC + MB2 (modif base avec plan fiscal CRC2002-10 générant un exceptionnel)
  //if (OpeSuiv<>OpeIni) and (((OpeSuiv='CDM') and (OpeIni='ELE')) // on enchaîne l'annulation si un élément exceptionnel a été effectué en même temps qu'un changement de plan
  //  or (OpeSuiv = 'ELC') or ((OpeIni ='ELC') and (OpeSuiv='MBA'))) then
  if (OpeSuiv<>OpeIni) and (((OpeSuiv='CDM') and (OpeIni='ELE')) // on enchaîne l'annulation si un élément exceptionnel a été effectué en même temps qu'un changement de plan
  or (OpeSuiv = 'ELC') or ((OpeIni ='ELC') and (OpeSuiv='MBA'))
  or  ((OpeIni ='ELC') and (OpeSuiv='MB2'))) then
  begin
    TraiteEnregLog(TSerie.Detail[0]);
    if (fTypeOpe='CES') or (fTypeOpe='CEP') then AnnuleCession else EnregAnnuleOp ;
    CodeImmo := TSerie.Detail[0].GetValue('IL_IMMO');
    iOrdre := TSerie.Detail[0].GetValue('IL_ORDRE');
    ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'" AND IL_ORDRE='+IntToStr(iOrdre)) ;
    result := rt;
    TSerie.Free;
    exit;
  end else TSerie.Free;
  if HM.execute(20,'','') <> mrYes then begin result:= rt; exit; end;
  {$IFDEF SERIE1}
  LeIn:=AGLLanceFiche('I','MULIMMOLOG','','','IL_ORDRESERIE='+OldOrdreSerie) ;
  {$ELSE}
  LeIn:='X';
  {$ENDIF}
  if LeIn<>'' then
  begin
    {$IFDEF SERIE1}
    QTmp:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO IN ('+LeIn+') AND IL_ORDRESERIE='+OldOrdreSerie,FALSE) ;
    {$ELSE}
    QTmp:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_ORDRESERIE='+OldOrdreSerie,FALSE) ;
    {$ENDIF}
    try
      TSerie := TOB.Create ('', nil, -1);
      TSerie.LoadDetailDB('','','',QTmp,False);
    finally
      Ferme (QTmp);
    end;
    for i:=0 to TSerie.Detail.Count - 1 do
    begin
      TraiteEnregLog(TSerie.Detail[i]);
      if (fTypeOpe='CES') or (fTypeOpe='CEP') then AnnuleCession else EnregAnnuleOp ;
      ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+TSerie.Detail[i].GetValue('IL_IMMO')+'" AND IL_ORDRE='+IntToStr(TSerie.Detail[i].GetValue('IL_ORDRE'))) ;
    end ;
    TSerie.Free;
  end ;
end ;

procedure TOpeEnCours.AnnuleEtablissement(LogEtabl : TLogEtablissement);
var
  CodeCourant,AncienEtabl : string;
  QueryImmo : TQuery;
begin
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogEtabl.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;
  AncienEtabl:=LogEtabl.Etabl;
  CodeCourant:=LogEtabl.Lequel ;
  QueryImmo.FindField('I_ETABLISSEMENT').AsString := AncienEtabl;
  MajOpeEnCoursImmo ( QueryImmo,'I_OPEETABLISSEMENT', 'ETA', '-');
  QueryImmo.post;
  Ferme(QueryImmo);
end;

procedure TOpeEnCours.AnnuleLieu(LogLieuGeo : TLogLieuGeo);
var
  CodeCourant,AncienLieu : string;
  QueryImmo : TQuery;
begin
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogLieuGeo.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;
  AncienLieu:=LogLieuGeo.LieuGeo;
  CodeCourant:=LogLieuGeo.Lequel;
  QueryImmo.FindField('I_LIEUGEO').AsString := AncienLieu;
  MajOpeEnCoursImmo ( QueryImmo,'I_OPELIEUGEO', 'LIE', '-');
  QueryImmo.post;
  Ferme(QueryImmo);
end;

// TGA 10/2005 CRC200210
procedure TOpeEnCours.AnnuleDepreciation(LogDepreciation : TLogDepreciation);
var QueryImmo : TQuery;
    Plan : TPlanAmort;
    PlanCourant,NumPlanPrec : string;
    //Q: TQuery; BTY 04/06 Pour diminuer les messages compilateur

    //codemut: string;
    //ordre : string;
    //MtDepre: double;
begin

  // Lecture de immolog pour récupérer l'ancien code mutation
  // et le montant de la dépréciation avant destruction
  //ordre := intTostr(LogDepreciation.Ordre);
  //Q:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+
  //LogDepreciation.Lequel+'" AND IL_ORDRE="'+Ordre+'"', FALSE);
  //if not Q.EOF then
  //  begin
  //     Q.Edit;
  //     CodeMut :=Q.FindField('IL_CODEMUTATION').AsString;
  //     mtdepre :=Q.FindField('IL_MONTANTDOT').AsFloat;
  //  end;
  //Ferme(Q);

  Plan:=TPlanAmort.Create(true) ;
  NumPlanPrec:=IntToStr(LogDepreciation.PlanActifAv);

  // Lecture de IMMMO
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogDepreciation.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;

  // Maj de immoamor
  PlanCourant:=QueryImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+QueryImmo.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;

  // maj du code mutation
  //QueryImmo.FindField('I_JOURNALA').AsString := CodeMut;

  // maj du mt dépréciaition
  QueryImmo.FindField('I_REVISIONECO').AsFloat := 0 ;
  Queryimmo.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
  // BTY 03/06 Attention FQ 17476 en 02/06, impact sur I_REVISIONFISCALE
  QueryImmo.FindField('I_REVISIONFISCALE').AsFloat := 0 ;

  Plan.Recupere(Queryimmo.FindField('I_IMMO').AsString,NumPlanPrec);
  QueryImmo.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
  QueryImmo.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);
  Plan.free ;

  // BTY 11/05 NON, I_OPERATION traité dans l'appel MajOpeEnCoursImmo suivant
  // sup coche opération en cours
  // MajOpeEnCoursImmo ( QueryImmo,'I_OPERATION', 'DPR', '-');

  // BTY 01/06 NON, I_OPECHANGEPLAN traité correctement dans MajOpeEnCoursImmo
  // Si pas d'exceptionnel maj 'I_OPECHANGEPLAN'
  // if QueryImmo.FindField('I_MONTANTEXC').AsFloat = 0 then
  //   begin
     // MajOpeEnCoursImmo ( QueryImmo, 'I_OPECHANGEPLAN','ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="ELC','-');
     // BTY 11/05  Prendre en compte les cas CD2 et CD3
     // PGR 12/05  Ajout CDA
     // BTY 01/06 Décocher I_OPECHANGEPLAN si une seule opération parmi ELE/CDM/CD2/CD3/CDA/ELC/DPR
     // mbo fq 17923 + MMS opération de changement date de mise en service
     MajOpeEnCoursImmo ( QueryImmo,
               'I_OPECHANGEPLAN',
               'ELE" OR IL_TYPEOP="MMS" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC" OR IL_TYPEOP="DPR'
                         ,'-');
//                         'ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC'
//                         ,'-');
     // BTY 01/06 Nelle zone de IMMO
     MajOpeEnCoursImmo ( QueryImmo, 'I_OPEDEPREC', 'DPR', '-');
   // end;

  QueryImmo.post;
  Ferme(QueryImmo);
end;

//MBO - FQ 17923 - 04/2006
procedure TOpeEnCours.AnnuleModifService(LogModifService : TLogModifService);
var QueryImmo : TQuery;
    //YCP suppr. Conseil , Plan : TPlanAmort; MVG 12/07/2006
    PlanCourant,NumPlanPrec : string;
begin
//  Plan:=TPlanAmort.Create(true) ;
  NumPlanPrec:=IntToStr(LogModifService.PlanActifAv);

  // Lecture de IMMMO
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogModifService.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;

  // Maj de immoamor
  PlanCourant:=QueryImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+QueryImmo.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;

  // maj de la date de mise en service
  Queryimmo.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
  QueryImmo.FindField('I_DATEAMORT').AsDateTime := LogModifService.DateOpReelle ;
  // maj date debut amortissement eco
  QueryImmo.FindField('I_DATEDEBECO').AsDateTime := StrToDate(LogModifService.DebEco) ;
  // maj date debut amortissement fiscal
  QueryImmo.FindField('I_DATEDEBFIS').AsDateTime := StrToDate(LogModifService.DebFis) ;

//  Plan.Recupere(Queryimmo.FindField('I_IMMO').AsString,NumPlanPrec);
//  Plan.free ;

     // BTY 01/06 Décocher I_OPECHANGEPLAN si une seule opération parmi ELE/CDM/CD2/CD3/CDA/ELC/DPR
     // MBO 11.09.2006 correction pour ajouter mms (dpr en double)
                      //               'ELE" OR IL_TYPEOP="DPR" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC" OR IL_TYPEOP="DPR'
     MajOpeEnCoursImmo ( QueryImmo,
               'I_OPECHANGEPLAN',
               'ELE" OR IL_TYPEOP="DPR" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC" OR IL_TYPEOP="MMS'
                         ,'-');
//                         'ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC'
//                         ,'-');

  QueryImmo.post;
  Ferme(QueryImmo);

end;


{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 11/09/2006
Modifié le ... : 05/10/2006
Description .. : Annulation  de l'opération prime
Mots clefs ... :
*****************************************************************}
procedure TOpeEnCours.AnnulePrime(LogPrime : TLogPrime);
var QueryImmo : TQuery;
    PlanCourant,NumPlanPrec : string;
begin
  NumPlanPrec:=IntToStr(LogPrime.PlanActifAv);

  // Lecture de IMMMO
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogPrime.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;

  if QueryImmo.FindField('I_IMMOORIGINEECL').AsString <> '' then
     HM.execute(25,'','');

  // Maj de immoamor
  PlanCourant:=QueryImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+QueryImmo.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;

  // réinitialisation des zones liées à la prime
  Queryimmo.FindField('I_SBVPRI').AsFloat:=0.00;    // montant de la prime
  if QueryImmo.FindField('I_SBVMT').AsFloat = 0.00 then
     QueryImmo.FindField('I_CORVRCEDDE').AsFloat := 0.00 ; //durée d'inaliénabilité
  QueryImmo.FindField('I_REPRISEUO').AsFloat := 0.00 ;  // antérieurs prime
  Queryimmo.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
     MajOpeEnCoursImmo ( QueryImmo,
               'I_OPECHANGEPLAN',
               'ELE" OR IL_TYPEOP="DPR" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC" OR IL_TYPEOP="MMS'
                         ,'-');

  QueryImmo.post;
  Ferme(QueryImmo);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Maryse BOUDIN
Créé le ...... : 09/10/2006
Modifié le ... :
Description .. : Annulation  de l'opération SUBVENTION
Mots clefs ... :
*****************************************************************}
procedure TOpeEnCours.AnnuleSBV(LogSBV : TLogSBV);
var QueryImmo : TQuery;
    PlanCourant,NumPlanPrec : string;
begin
  NumPlanPrec:=IntToStr(LogSBV.PlanActifAv);

  // Lecture de IMMMO
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogSBV.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;

  if QueryImmo.FindField('I_IMMOORIGINEECL').AsString <> '' then
     HM.execute(26,'','');

  // Maj de immoamor
  PlanCourant:=QueryImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+QueryImmo.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;

  // réinitialisation des zones liées à la subvention
  Queryimmo.FindField('I_SBVMT').AsFloat:=0.00;    // montant de la subention
  if QueryImmo.FindField('I_SBVPRI').AsFloat = 0.00 then
     QueryImmo.FindField('I_CORVRCEDDE').AsFloat := 0.00 ; //durée d'inaliénabilité
  QueryImmo.FindField('I_CORRECTIONVR').AsFloat := 0.00 ;  // antérieurs subvention
  QueryImmo.FindField('I_CPTSBVB').AsString := '';
  // le compte de reprise de sbv n'est plus stocké - QueryImmo.FindField('I_CPTSBVR').AsString := '';
  QueryImmo.FindField('I_DPIEC').AsString := '-';

  Queryimmo.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
     MajOpeEnCoursImmo ( QueryImmo,
               'I_OPECHANGEPLAN',
               'ELE" OR IL_TYPEOP="DPR" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="ELC" OR IL_TYPEOP="MMS'
                         ,'-');

  QueryImmo.post;
  Ferme(QueryImmo);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 10/10/2006
Modifié le ... :   /  /
Description .. : Annulation de l'opération de réduction d'une prime
Suite ........ : d'équipement
Mots clefs ... :
*****************************************************************}
procedure TOpeEnCours.AnnuleReductionPrime(LogReducPrime : TLogReducPrime);
var QImmo : TQuery;
    PlanCourant, NumPlanPrec : string;
begin
  NumPlanPrec:=IntToStr(LogReducPrime.PlanActifAv);

  // Lecture de IMMMO
  QImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogReducPrime.Lequel+'"', FALSE) ;
  QImmo.Edit ;

  // Maj de IMMOAMOR
  PlanCourant:= QImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="' + QImmo.FindField('I_IMMO').AsString +
              '" AND IA_NUMEROSEQ=' + PlanCourant) ;

  // Récup anciennes zones prime
  Qimmo.FindField('I_SBVPRI').AsFloat:= LogReducPrime.Prime;
  QImmo.FindField('I_REPRISEUO').AsFloat := LogReducPrime.Suramort;
  QImmo.FindField('I_PLANACTIF').AsString:= NumPlanPrec;

  // Maj code opération
  // Champ bidon I_OPECHANGEPLAN pour éviter plantage ds la procédure
  MajOpeEnCoursImmo ( QImmo, 'I_OPECHANGEPLAN', '', '-');

  QImmo.Post;
  Ferme(QImmo);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 07/11/2006
Modifié le ... :   /  /
Description .. : Annulation de l'opération de réduction d'une subvention
Suite ........ : d'investissement
Mots clefs ... :
*****************************************************************}
procedure TOpeEnCours.AnnuleReductionSBV(LogReducSBV : TLogReducSBV);
var QImmo : TQuery;
    PlanCourant, NumPlanPrec : string;
begin
  NumPlanPrec:=IntToStr(LogReducSBV.PlanActifAv);

  // Lecture de IMMMO
  QImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogReducSBV.Lequel+'"', FALSE) ;
  QImmo.Edit ;

  // Maj de IMMOAMOR
  PlanCourant:= QImmo.FindField('I_PLANACTIF').AsString;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="' + QImmo.FindField('I_IMMO').AsString +
              '" AND IA_NUMEROSEQ=' + PlanCourant) ;

  // Récup anciennes zones subvention
  Qimmo.FindField('I_SBVMT').AsFloat:= LogReducSBV.SBV;
  QImmo.FindField('I_CORRECTIONVR').AsFloat := LogReducSBV.Reprise;
  QImmo.FindField('I_PLANACTIF').AsString:= NumPlanPrec;

  // Maj code opération
  // Champ bidon I_OPECHANGEPLAN pour éviter plantage ds la procédure
  MajOpeEnCoursImmo ( QImmo, 'I_OPECHANGEPLAN', '', '-');

  QImmo.Post;
  Ferme(QImmo);
end;



//PGR 11/2005 Changement des conditions d'amortissement
procedure TOpeEnCours.AnnuleChgtCondAmort(LogChgtCondAmort : TLogChgtCondAmort);
var QueryImmo : TQuery;
    Plan : TPlanAmort;
    PlanCourant,NumPlanPrec : string;
    //Q: TQuery; BTY 04/06 Pour diminuer les messages compilateur
    taux: double;

begin
  NumPlanPrec:=IntToStr(LogChgtCondAmort.PlanActifAv);

  // Lecture de IMMMO
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogChgtCondAmort.Lequel+'"', FALSE) ;
  PlanCourant:=QueryImmo.FindField('I_PLANACTIF').AsString;
  QueryImmo.Edit ;

  Queryimmo.FindField('I_PLANACTIF').AsString:=NumPlanPrec;
  Queryimmo.FindField('I_METHODEECO').AsString:=LogChgtCondAmort.MethodeEco;
  Queryimmo.FindField('I_DUREEECO').AsInteger:=LogChgtCondAmort.DureeEco;
  Queryimmo.FindField('I_DATEAMORT').AsDateTime:=LogChgtCondAmort.DateOpReelle;
  Queryimmo.FindField('I_DUREEREPRISE').AsString:=LogChgtCondAmort.Duree;
  Queryimmo.FindField('I_REPRISEECO').AsFloat:=LogChgtCondAmort.RepriseEco;
  Queryimmo.FindField('I_REPRISEFISCAL').AsFloat:=LogChgtCondAmort.RepriseFisc;
  // modif mbo fq 17569 - nécessité de gérer les enreg avant modif de dates
  if Trim(LogChgtCondAmort.datedebeco) = '' then
     AncienLogCDA := true
  else
     AncienLogCDA := false;

  if AncienLogCDA = false then
  begin
     Queryimmo.FindField('I_DATEDEBECO').AsDateTime:= StrToDate(LogChgtCondAmort.datedebeco);
     Queryimmo.FindField('I_DATEDEBFIS').AsDateTime:= StrToDate(LogChgtCondAmort.datedebfis);
  end else
  begin
       if LogChgtCondAmort.MethodeEco <> 'LIN' then
          Queryimmo.FindField('I_DATEDEBECO').AsDateTime:= LogChgtCondAmort.dateOpReelle
       else
          Queryimmo.FindField('I_DATEDEBECO').AsDateTime:= Queryimmo.FindField('I_DATEPIECEA').AsDateTIme;

       if Queryimmo.FindField('I_METHODEFISC').AsString = 'DEG' then
          Queryimmo.FindField('I_DATEDEBFIS').AsDateTime:= Queryimmo.FindField('I_DATEPIECEA').AsDateTime
       else
          Queryimmo.FindField('I_DATEDEBFIS').AsDateTime:= LogChgtCondAmort.dateOpReelle;
  end;

  taux := GetTaux( Queryimmo.FindField('I_METHODEECO').AsString,
          Queryimmo.FindField('I_DATEDEBECO').AsDateTime,
          Queryimmo.FindField('I_DATEDEBECO').AsDateTime,
          Queryimmo.FindField('I_DUREEECO').AsInteger);
  Queryimmo.FindField('I_TAUXECO').AsFloat:=taux;
  //PGR - 12/2005 Ajout MajOpeEnCoursImmo
  MajOpeEnCoursImmo ( QueryImmo,
     'I_OPECHANGEPLAN',
     'ELE" OR IL_TYPEOP="MMS" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="CDA" OR IL_TYPEOP="DPR" OR IL_TYPEOP="ELC'
     ,'-');

  Plan:=TPlanAmort.Create(true) ;
  Plan.Recupere(Queryimmo.FindField('I_IMMO').AsString,NumPlanPrec);

  QueryImmo.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
  QueryImmo.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);
  Plan.free ;

  // Maj de immoamor
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+QueryImmo.FindField('I_IMMO').AsString+
                 '" AND IA_NUMEROSEQ='+PlanCourant) ;
  QueryImmo.post;
  Ferme(QueryImmo);
end;

// BTY 04/06 FQ 17516
procedure TOpeEnCours.AnnuleRegroupement (LogGroupe : TLogGroupe);
var
  CodeCourant,AncienGroupe : string;
  QueryImmo : TQuery;
begin
  QueryImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogGroupe.Lequel+'"', FALSE) ;
  QueryImmo.Edit ;
  AncienGroupe:=LogGroupe.Regroupement;
  CodeCourant:=LogGroupe.Lequel;
  QueryImmo.FindField('I_GROUPEIMMO').AsString := AncienGroupe;
  MajOpeEnCoursImmo ( QueryImmo,'I_OPEREG', 'REG', '-');
  QueryImmo.post;
  Ferme(QueryImmo);
end;
//

procedure TOpeEnCours.bAccedeImmoClick(Sender: TObject);
var QTmp : TQuery;
begin
  QTmp := OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmoDest+'"',TRUE);
  AMLanceFiche_FicheImmobilisation(QTmp.FindField('I_IMMO').AsString,taConsult,'') ;
  Ferme(QTmp);
end;

procedure TOpeEnCours.TraiteEnregLog(T : TOB);
begin
  fLequel0 := T.GetValue('IL_IMMO') ;
  fOrdre := T.GetValue('IL_ORDRE');
  fOrdreSerie := T.GetValue('IL_ORDRESERIE');
  fTypeOpe := T.GetValue('IL_TYPEOP');
  if (fTypeOpe='CEP') or (fTypeOpe='CES') then TraiteEnregLogCession(T)
  else if fTypeOpe ='MUT' then TraiteEnregLogMutation(T)
  else if fTypeOpe ='ECL' then TraiteEnregLogEclatement(T)
  else if fTypeOpe ='ETA' then TraiteEnregLogEtablissement(T)
  else if fTypeOpe ='LIE' then TraiteEnregLogLieuGeo(T)
  // TGA 10/2005 CRC200210
  else if fTypeOpe ='DPR' then TraiteEnregLogDepreciation(T)
  //PGR 11/2005 Changement des conditions d'amortissement
  else if fTypeOpe ='CDA' then TraiteEnregLogChgtCondAmort(T)
  else if fTypeOpe ='LEV' then TraiteEnregLogLeveeOption(T)
  //else if fTypeOpe ='MBA' then TraiteEnregLogModifBases (T)
  else if (fTypeOpe ='MBA') OR (fTypeOpe = 'MB2') then TraiteEnregLogModifBases (T)
  // BTY 04/06 FQ 17516
  else if fTypeOpe ='REG' then TraiteEnregLogRegroupement(T)
  else if fTypeOpe ='MMS' then TraiteEnregLogService(T)  // mbo fq 17923
  else if OpeChangementPlan(fTypeOpe) then TraiteEnregLogChPlan(T,LogChPlan)
  else if fTypeOpe ='PRI' then TraiteEnregLogPrime(T)  // mbo 11.09.2006
  else if fTypeOpe = 'SBV' then TraiteEnregLogSBV(T)  //mbo 09.10.2006
  else if fTypeOpe = 'RPR' then TraiteEnregLogReducPrime(T)  // BTY 10/06 Réduction prime
  else if fTypeOpe = 'RSB' then TraiteEnregLogReducSBV(T);   // BTY 11/06 Réduction subvention
end;

procedure TOpeEnCours.TraiteEnregLogCession(T : TOB);
begin
  LogCession.Lequel := T.GetValue('IL_IMMO');
  LogCession.Ordre := T.GetValue('IL_ORDRE');
  LogCession.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogCession.TypeOpe := T.GetValue('IL_TYPEOP');
  LogCession.DateOpe := T.GetValue('IL_DATEOP');
  LogCession.QteCedee := T.GetValue('IL_QTECEDEE');
  LogCession.VoCedee := T.GetValue('IL_VOCEDEE');
  LogCession.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
  LogCession.TvaRecuperable := T.GetValue('IL_TVARECUPERABLE');
  LogCession.TvaRecuperee := T.GetValue('IL_TVARECUPEREE');
  LogCession.Vnc := T.GetValue('IL_VNC');
  LogCession.RepriseEco := T.GetValue('IL_REPRISEECO');
  LogCession.RepriseFisc := T.GetValue('IL_REPRISEFISC');
  LogCession.MontantExc := T.GetValue('IL_MONTANTEXC');
  if not VarIsNull (T.GetValue('IL_TYPEEXC')) then
    LogCession.TypeExc := T.GetValue('IL_TYPEEXC')
  else LogCession.TypeExc := '';
  LogCession.BaseEco := T.GetValue('IL_BASEECOAVMB');
  LogCession.BaseFisc := T.GetValue('IL_BASEFISCAVMB');
  LogCession.BaseTP := T.GetValue('IL_BASETAXEPRO');
  // Ajout mbo pour antérieurs saisis subvention 24/10/2006
  LogCession.RepriseSBV := T.GetValue('IL_MONTANTAVMB');
  // 18/09/07 Nouveaux paramètres immos
  LogCession.JourSortie := T.GetValue('IL_CODECB');
  if (T.GetValue('IL_CODECB')= '') then
     LogCession.JourSortie := 'OUI';
  // 18/09/07 Fin Nouveaux paramètres immos
end;

procedure TOpeEnCours.TraiteEnregLogMutation(T : TOB);
begin
  LogMutation.Lequel := T.GetValue('IL_IMMO');
  LogMutation.Ordre := T.GetValue('IL_ORDRE');
  LogMutation.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogMutation.TypeOpe := T.GetValue('IL_TYPEOP');
  LogMutation.DateOpe := T.GetValue('IL_DATEOP');
  if not VarIsNull (T.GetValue('IL_CODEMUTATION')) then
    LogMutation.CodeMutation := T.GetValue('IL_CODEMUTATION')
  else LogMutation.CodeMutation := '';
  if not VarIsNull (T.GetValue('IL_CPTEMUTATION')) then
    LogMutation.CpteMutation := T.GetValue('IL_CPTEMUTATION')
  else LogMutation.CpteMutation := '';
end;

procedure TOpeEnCours.TraiteEnregLogEclatement(T : TOB);
begin
  LogEclatement.Lequel := T.GetValue('IL_IMMO');
  LogEclatement.Ordre := T.GetValue('IL_ORDRE');
  LogEclatement.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogEclatement.TypeOpe := T.GetValue('IL_TYPEOP');
  LogEclatement.DateOpe := T.GetValue('IL_DATEOP');
  LogEclatement.CodeEclate := T.GetValue('IL_CODEECLAT');
  LogEclatement.MontantEclate := T.GetValue('IL_MONTANTECL');
  LogEclatement.QteEclate := T.GetValue('IL_QTEECLAT');
  LogEclatement.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

// mbo - 12.09.2006 pour gestion prime d'équipement
procedure TOpeEnCours.TraiteEnregLogPrime(T : TOB);
begin
  LogPrime.Lequel := T.GetValue('IL_IMMO');
  LogPrime.Ordre := T.GetValue('IL_ORDRE');
  LogPrime.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogPrime.TypeOpe := T.GetValue('IL_TYPEOP');
  LogPrime.Prime   := T.GetValue('IL_MONTANTEXC');
  LogPrime.DureeEco := T.GetValue('IL_DUREEECO');
  LogPrime.MethodeEco := T.GetValue('IL_METHODEECO');
  LogPrime.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

// mbo - 09.10.2006 pour gestion subvention
procedure TOpeEnCours.TraiteEnregLogSBV(T : TOB);
begin
  LogSBV.Lequel := T.GetValue('IL_IMMO');
  LogSBV.Ordre := T.GetValue('IL_ORDRE');
  LogSBV.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogSBV.TypeOpe := T.GetValue('IL_TYPEOP');
  LogSBV.Subvention   := T.GetValue('IL_MONTANTEXC');
  LogSBV.DureeEco := T.GetValue('IL_DUREEECO');
  LogSBV.MethodeEco := T.GetValue('IL_METHODEECO');
  LogSBV.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

// BTY 10/06 Réduction de la prime d'équipement
procedure TOpeEnCours.TraiteEnregLogReducPrime(T : TOB);
begin
  LogReducPrime.Lequel := T.GetValue('IL_IMMO');
  LogReducPrime.Ordre := T.GetValue('IL_ORDRE');
  LogReducPrime.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogReducPrime.TypeOpe := T.GetValue('IL_TYPEOP');
  LogReducPrime.Prime   := T.GetValue('IL_MONTANTEXC');
  LogReducPrime.Base   := T.GetValue('IL_BASETAXEPRO');
  LogReducPrime.Suramort   := T.GetValue('IL_MONTANTAVMB');
  LogReducPrime.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

// BTY 11/06 Réduction de la subvention d'investissement
procedure TOpeEnCours.TraiteEnregLogReducSBV(T : TOB);
begin
  LogReducSBV.Lequel := T.GetValue('IL_IMMO');
  LogReducSBV.Ordre := T.GetValue('IL_ORDRE');
  LogReducSBV.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogReducSBV.TypeOpe := T.GetValue('IL_TYPEOP');
  LogReducSBV.SBV   := T.GetValue('IL_MONTANTEXC');
  LogReducSBV.Reprise   := T.GetValue('IL_MONTANTAVMB');
  LogReducSBV.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

procedure TOpeEnCours.TraiteEnregLogEtablissement(T : TOB);
begin
  LogEtabl.Lequel := T.GetValue('IL_IMMO');
  LogEtabl.Ordre := T.GetValue('IL_ORDRE');
  LogEtabl.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogEtabl.TypeOpe := T.GetValue('IL_TYPEOP');
  LogEtabl.Etabl := T.GetValue('IL_ETABLISSEMENT');
end;

procedure TOpeEnCours.TraiteEnregLogLieuGeo(T : TOB);
begin
  LogLieuGeo.Lequel := T.GetValue('IL_IMMO');
  LogLieuGeo.Ordre := T.GetValue('IL_ORDRE');
  LogLieuGeo.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogLieuGeo.TypeOpe := T.GetValue('IL_TYPEOP');
  LogLieuGeo.LieuGeo := T.GetValue('IL_LIEUGEO');
end;

// TGA 10/2005 CRC200210
procedure TOpeEnCours.TraiteEnregLogDepreciation(T : TOB);
begin
  LogDepreciation.Lequel := T.GetValue('IL_IMMO');
  LogDepreciation.Ordre := T.GetValue('IL_ORDRE');
  LogDepreciation.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogDepreciation.TypeOpe := T.GetValue('IL_TYPEOP');
  LogDepreciation.Depreciation := ABS(T.GetValue('IL_MONTANTDOT'));
  LogDepreciation.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

// mbo 04/2006 - fq 17923
procedure TOpeEnCours.TraiteEnregLogService(T : TOB);
begin
  LogModifService.Lequel := T.GetValue('IL_IMMO');
  LogModifService.Ordre := T.GetValue('IL_ORDRE');
  LogModifService.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogModifService.TypeOpe := T.GetValue('IL_TYPEOP');
  LogModifService.DebEco := T.GetValue('IL_CODEMUTATION');
  LogModifService.DebFis := T.GetValue('IL_CODEECLAT');
  LogModifService.DateOpReelle := T.GetValue('IL_DATEOPREELLE');
  LogModifService.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
end;

//PGR 11/2005 Changement des conditions d'amortissement
procedure TOpeEnCours.TraiteEnregLogChgtCondAmort(T : TOB);
begin
  LogChgtCondAmort.Lequel := T.GetValue('IL_IMMO');
  LogChgtCondAmort.Ordre := T.GetValue('IL_ORDRE');
  LogChgtCondAmort.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogChgtCondAmort.TypeOpe := T.GetValue('IL_TYPEOP');
  LogChgtCondAmort.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
  LogChgtCondAmort.MethodeEco := T.GetValue('IL_METHODEECO');
  LogChgtCondAmort.DureeEco := T.GetValue('IL_DUREEECO');
  LogChgtCondAmort.DateOpReelle := T.GetValue('IL_DATEOPREELLE');
  LogChgtCondAmort.Duree := T.GetValue('IL_DUREE');
  //PGR - 12/2005 - Changement des conditions d'amortissement(récup I_REPRISEECO et I_REPRISEFISC)
  LogChgtCondAmort.RepriseEco := T.GetValue('IL_CUMANTCESECO');
  LogChgtCondAmort.RepriseFisc := T.GetValue('IL_CUMANTCESFIS');
  // mbo fq 17569
  LogChgtCondAmort.datedebeco := T.GetValue('IL_CODEMUTATION');
  LogChgtCondAmort.datedebfis := T.GetValue('IL_CODEECLAT');

end;

procedure TOpeEnCours.TraiteEnregLogLeveeOption(T : TOB);
begin
  LogLeveeOpt.Lequel := T.GetValue('IL_IMMO');
  LogLeveeOpt.Ordre := T.GetValue('IL_ORDRE');
  LogLeveeOpt.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogLeveeOpt.TypeOpe := T.GetValue('IL_TYPEOP');
  // FQ 19280
  LogLeveeOpt.DateFinCt := T.GetValue('IL_DATEOPREELLE');
end;

procedure TOpeEnCours.TraiteEnregLogModifBases(T : TOB);
begin
  LogModifBases.Lequel := T.GetValue('IL_IMMO');
  LogModifBases.Ordre := T.GetValue('IL_ORDRE');
  LogModifBases.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogModifBases.TypeOpe := T.GetValue('IL_TYPEOP');
  LogModifBases.MontantHT := T.GetValue('IL_MONTANTAVMB');
  LogModifBases.BaseEco := T.GetValue('IL_BASEECOAVMB');
  LogModifBases.BaseFisc := T.GetValue('IL_BASEFISCAVMB');
  LogModifBases.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
  LogModifBases.RepriseEco := T.GetValue('IL_CUMANTCESECO');
  LogModifBases.RepriseFisc :=  T.GetValue('IL_CUMANTCESFIS');
  // BTY 10/05 Annulation modif base avec plan fiscal CRC2002-10
  if not VarIsNull (T.GetValue('IL_METHODEFISC')) then
         LogModifBases.MethodeFisc := T.GetValue('IL_METHODEFISC')
  else   LogModifBases.MethodeFisc := '';
  LogModifBases.DureeFisc :=  T.GetValue('IL_DUREEFISC');
  // BTY 03/06 FQ 17446
  LogModifBases.CodeEclat := T.GetValue('IL_CODEECLAT');
  // 04/07 Nelle gestion fiscale
  LogModifBases.CalculVNF := T.GetValue('IL_LIEUGEO');
  LogModifBases.GestionFiscale := T.GetValue('IL_ETABLISSEMENT');
  LogModifBases.NumVersion := T.GetValue('IL_VERSION');
  LogModifBases.RepriseDR := T.GetValue('IL_REVISIONDOTECO');
  LogModifBases.RepriseFEC := T.GetValue('IL_REVISIONREPECO');
end;

// BTY 04/06 FQ 17516 Changement de regroupement devient une opération
procedure TOpeEnCours.TraiteEnregLogRegroupement(T : TOB);
begin
  LogGroupe.Lequel := T.GetValue('IL_IMMO');
  LogGroupe.Ordre := T.GetValue('IL_ORDRE');
  LogGroupe.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  LogGroupe.TypeOpe := T.GetValue('IL_TYPEOP');
  LogGroupe.Regroupement := T.GetValue('IL_LIEUGEO');
end;


procedure TraiteEnregLogChPlan(T : TOB; var St : TLogChPlan);
begin
  St.Lequel := T.GetValue('IL_IMMO');
  St.Ordre := T.GetValue('IL_ORDRE');
  St.OrdreSerie := T.GetValue('IL_ORDRESERIE');
  St.TypeOpe := T.GetValue('IL_TYPEOP');
  St.DateOpe := T.GetValue('IL_DATEOP');
  St.PlanActifAv := T.GetValue('IL_PLANACTIFAV');
  St.PlanActifAp := T.GetValue('IL_PLANACTIFAP');
  St.MontantExc := T.GetValue('IL_MONTANTEXC');
  St.TypeExc := T.GetValue('IL_TYPEEXC');
  St.RevisionEco := T.GetValue('IL_REVISIONECO');
  St.RevisionFisc := T.GetValue('IL_REVISIONFISC');
  St.DureeEco := T.GetValue('IL_DUREEECO');
  St.DureeFisc := T.GetValue('IL_DUREEFISC');
  St.MethodeEco := T.GetValue('IL_METHODEECO');
  if not VarIsNull (T.GetValue('IL_METHODEFISC')) then
    St.MethodeFisc := T.GetValue('IL_METHODEFISC')
  else St.MethodeFisc := '';
  if St.MethodeEco='NAM' then  // Astuce pour récupérer les dates initiales avant révision du plan
  begin
    if not VarIsNull (T.GetValue('IL_CODEMUTATION')) then
    begin
      if (T.GetValue('IL_CODEMUTATION')<>'') then
      begin
        St.DatePieceA := Str8ToDate(Copy(T.GetValue('IL_CODEMUTATION'),1,8),True);
        St.DateAmort := Str8ToDate(Copy(T.GetValue('IL_CODEMUTATION'),9,8), True);
      end;
    end;
  end;
  // BTY 11/05 Cas du calcul plan futur avec VNC et plan fiscal CRC 2002-10
  St.CodeEclat := T.GetValue('IL_CODEECLAT');
  St.BaseFiscAvMB := T.GetValue('IL_BASEFISCAVMB');
  St.RepriseEco := T.GetValue('IL_REPRISEECO');
  St.RepriseFisc := T.GetValue('IL_REPRISEFISC');
  // BTY Durée d'inaliénabilité d'immo non amortissable ayant eu une prime ou subvention
  if VarIsNull(T.GetValue('IL_CODECB')) = False then  St.DureePriSbv := 0
  else St.DureePriSbv := StrToFloat (T.GetValue('IL_CODECB'));
  // 04/07 Plan futur avec la VNF
  if not VarIsNull (T.GetValue('IL_LIEUGEO')) then
     St.CalculVNF := T.GetValue('IL_LIEUGEO')
  else St.CalculVNF := '';
  // 04/07 Gestion fiscale
  if not VarIsNull (T.GetValue('IL_ETABLISSEMENT')) then
     St.GestionFiscale := T.GetValue('IL_ETABLISSEMENT')
  else St.GestionFiscale := '';
  St.NumVersion := T.GetValue('IL_VERSION');
  St.RepriseDR := T.GetValue ('IL_REVISIONDOTECO');
  St.RepriseFEC := T.GetValue ('IL_REVISIONREPECO');
end;
{$IFDEF EAGLCLIENT}
{$ELSE}
procedure TraiteEnregLogChPlan(Q : TQuery ; var St : TLogChPlan);
begin
  St.Lequel := Q.FindField('IL_IMMO').AsString ;
  St.Ordre := Q.FindField('IL_ORDRE').AsInteger ;
  St.OrdreSerie := Q.FindField('IL_ORDRESERIE').AsInteger ;
  St.TypeOpe := Q.FindField('IL_TYPEOP').AsString;
  St.DateOpe := Q.FindField('IL_DATEOP').AsDateTime ;
  St.PlanActifAv := Q.FindField('IL_PLANACTIFAV').AsInteger ;
  St.PlanActifAp := Q.FindField('IL_PLANACTIFAP').AsInteger ;
  St.MontantExc := Q.FindField('IL_MONTANTEXC').AsFloat;
  St.TypeExc := Q.FindField('IL_TYPEEXC').AsString ;
  St.RevisionEco := Q.FindField('IL_REVISIONECO').AsFloat;
  St.RevisionFisc := Q.FindField('IL_REVISIONFISC').AsFloat;
  St.DureeEco := Q.FindField('IL_DUREEECO').AsInteger ;
  St.DureeFisc := Q.FindField('IL_DUREEFISC').AsInteger ;
  St.MethodeEco := Q.FindField('IL_METHODEECO').AsString ;
  St.MethodeFisc := Q.FindField('IL_METHODEFISC').AsString;

  // BTY 11/05 Cas du calcul plan futur avec VNC et plan fiscal CRC 2002-10
  St.CodeEclat := Q.FindField('IL_CODEECLAT').AsString;
  St.BaseFiscAvMB := Q.FindField('IL_BASEFISCAVMB').AsFloat;
  St.RepriseEco := Q.FindField('IL_REPRISEECO').AsFloat;
  St.RepriseFisc := Q.FindField('IL_REPRISEFISC').AsFloat;
  // BTY Durée d'inaliénabilité d'immo non amortissable ayant eu une prime ou subvention
  St.DureePriSbv := Q.FindField('IL_CODECB').AsFloat;
  // 04/07 Plan futur avec la VNF
  St.CalculVNF := Q.FindField('IL_LIEUGEO').AsString;
  // 04/07 Gestion fiscale
  St.Gestionfiscale := Q.FindField('IL_ETABLISSEMENT').AsString;
  St.NumVersion := Q.FindField('IL_VERSION').AsString;
  St.RepriseDR :=  Q.FindField('IL_REVISIONDOTECO').AsFloat;
  St.RepriseFEC :=  Q.FindField('IL_REVISIONREPECO').AsFloat;
end;
{$ENDIF}

procedure TOpeEnCours.ToolbarButton973Click(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TOpeEnCours.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
    if Key=VK_F10 then
    begin
      ModalResult:=mrOk ;
      bValiderClick(BValider);
    end;
end;

procedure TOpeEnCours.FListeRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  fLigneOperation := Ou;
  AfficheDetail;
  bSupprime.Enabled :=  (fListeOperation.Detail.Count > 0) and (fTypeAction <> taConsult);
end;



end.

